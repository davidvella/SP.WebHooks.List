﻿using Microsoft.Azure;
using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.Taxonomy;
using Microsoft.WindowsAzure.Storage;
using Microsoft.WindowsAzure.Storage.Queue;
using Newtonsoft.Json;
using Noggle.TikaOnDotNet.Parser;
using OfficeDevPnP.Core;
using SharePoint.WebHooks.Common.Models;
using SharePoint.WebHooks.Common.SQL;
using SharePoint.WebHooks.Common.TextAnalytics;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Threading.Tasks;
using MoreLinq;

namespace SharePoint.WebHooks.Common
{
    /// <summary>
    /// Helper class that deals with asynchronous and synchronous SharePoint list web hook events processing
    /// </summary>
    public class ChangeManager
    {
        #region Constants and variables
        public const string StorageQueueName = "sharepointlistwebhookevent";
        private string _accessToken;
        #endregion

        #region Async processing...add item to queue
        /// <summary>
        /// Add the notification message to an Azure storage queue
        /// </summary>
        /// <param name="storageConnectionString">Storage account connection string</param>
        /// <param name="notification">Notification message to add</param>
        public void AddNotificationToQueue(string storageConnectionString, NotificationModel notification)
        {
            var storageAccount =CloudStorageAccount.Parse(storageConnectionString);

            // Get queue... create if does not exist.
            var queueClient = storageAccount.CreateCloudQueueClient();
            var queue = queueClient.GetQueueReference(StorageQueueName);
            queue.CreateIfNotExists();

            // add message to the queue
            queue.AddMessage(new CloudQueueMessage(JsonConvert.SerializeObject(notification)));
        }
        #endregion

        #region Synchronous processing of a web hook notification
        /// <summary>
        /// Processes a received notification. This typically is triggered via an Azure Web Job that reads the Azure storage queue
        /// </summary>
        /// <param name="notification">Notification to process</param>
        public void ProcessNotification(NotificationModel notification)
        {
            ClientContext cc = null;
            try
            {
                #region Setup an app-only client context
                var am = new AuthenticationManager();

                var url = $"https://{CloudConfigurationManager.GetSetting("TenantName")}{notification.SiteUrl}";
                var realm = TokenHelper.GetRealmFromTargetUrl(new Uri(url));
                var clientId = CloudConfigurationManager.GetSetting("ClientId");
                var clientSecret = CloudConfigurationManager.GetSetting("ClientSecret");

                cc = new Uri(url).DnsSafeHost.Contains("spoppe.com")
                    ? am.GetAppOnlyAuthenticatedContext(url, realm, clientId, clientSecret,
                        acsHostUrl: "windows-ppe.net", globalEndPointPrefix: "login")
                    : am.GetAppOnlyAuthenticatedContext(url, clientId, clientSecret);

                cc.ExecutingWebRequest += Cc_ExecutingWebRequest;
                #endregion

                #region Grab the list for which the web hook was triggered
                var lists = cc.Web.Lists;
                var listId = new Guid(notification.Resource);
                var results = cc.LoadQuery(lists.Where(lst => lst.Id == listId));
                cc.ExecuteQueryRetry();
                var changeList = results.FirstOrDefault();
                if (changeList == null)
                {
                    // list has been deleted in between the event being fired and the event being processed
                    return;
                }
                #endregion

                #region Grab the list changes and do something with them
                // grab the changes to the provided list using the GetChanges method 
                // on the list. Only request Item changes as that's what's supported via
                // the list web hooks
                var changeQuery = new ChangeQuery(false, true)
                {
                    Item = true,
                    FetchLimit = 1000, // Max value is 2000, default = 1000
                    DeleteObject = false,
                    Add = true,
                    Update = true,
                    SystemUpdate = false
                };

                // grab last change token from database if possible
                using (var dbContext = new SharePointWebHooks())
                {
                    ChangeToken lastChangeToken = null;
                    var id = new Guid(notification.SubscriptionId);

                    var listWebHookRow = dbContext.ListWebHooks.Find(id);
                    if (listWebHookRow != null)
                    {
                        lastChangeToken = new ChangeToken
                        {
                            StringValue = listWebHookRow.LastChangeToken
                        };
                    }

                    // Start pulling down the changes
                    var allChangesRead = false;
                    do
                    {
                        // should not occur anymore now that we record the starting change token at 
                        // subscription creation time, but it's a safety net
                        if (lastChangeToken == null)
                        {
                            lastChangeToken = new ChangeToken
                            {
                                StringValue =
                                    $"1;3;{notification.Resource};{DateTime.Now.AddMinutes(-5).ToUniversalTime().Ticks.ToString()};-1"
                            };
                            // See https://blogs.technet.microsoft.com/stefan_gossner/2009/12/04/content-deployment-the-complete-guide-part-7-change-token-basics/
                        }

                        // Assign the change token to the query...this determines from what point in
                        // time we'll receive changes
                        changeQuery.ChangeTokenStart = lastChangeToken;

                        // Execute the change query
                        var changes = changeList.GetChanges(changeQuery);
                        cc.Load(changes);
                        cc.ExecuteQueryRetry();

                        // If item is changed more than once
                        var uniqueChanges = changes.Cast<ChangeItem>().AsEnumerable().DistinctBy(change => change.ItemId).ToList();

                        if (uniqueChanges.Any())
                        {
                            foreach (var change in uniqueChanges)
                            {
                                lastChangeToken = change.ChangeToken;

                                try
                                {
                                    // do "work" with the found change
                                    DoWork(cc, changeList, change);
                                }
                                catch (Exception)
                                {
                                    // ignored
                                }
                            }

                            // We potentially can have a lot of changes so be prepared to repeat the 
                            // change query in batches of 'FetchLimit' until we've received all changes
                            if (changes.Count < changeQuery.FetchLimit)
                            {
                                allChangesRead = true;
                            }
                        }
                        else
                        {
                            allChangesRead = true;
                        }
                        // Are we done?
                    } while (allChangesRead == false);

                    // Persist the last used change token as we'll start from that one
                    // when the next event hits our service
                    if (listWebHookRow != null)
                    {
                        // Only persist when there's a change in the change token
                        if (!listWebHookRow.LastChangeToken.Equals(lastChangeToken.StringValue, StringComparison.InvariantCultureIgnoreCase))
                        {
                            listWebHookRow.LastChangeToken = lastChangeToken.StringValue;
                            dbContext.SaveChanges();
                        }
                    }
                    else
                    {
                        // should not occur anymore now that we record the starting change token at 
                        // subscription creation time, but it's a safety net
                        dbContext.ListWebHooks.Add(new ListWebHooks()
                        {
                            Id = id,
                            ListId = listId,
                            LastChangeToken = lastChangeToken.StringValue,
                        });
                        dbContext.SaveChanges();
                    }
                }
                #endregion

                #region "Update" the web hook expiration date when needed
                // Optionally add logic to "update" the expiration date time of the web hook
                // If the web hook is about to expire within the coming 5 days then prolong it
                if (notification.ExpirationDateTime.AddDays(-5) >= DateTime.Now) return;
                var webHookManager = new WebHookManager();
                var updateResult = Task.WhenAny(
                    webHookManager.UpdateListWebHookAsync(
                        url,
                        listId.ToString(),
                        notification.SubscriptionId,
                        CloudConfigurationManager.GetSetting("WebHookEndPoint"),
                        DateTime.Now.AddMonths(3),
                        _accessToken)
                ).Result;

                if (updateResult.Result == false)
                {
                    throw new Exception(
                        $"The expiration date of web hook {notification.SubscriptionId} with endpoint {CloudConfigurationManager.GetSetting("WebHookEndPoint")} could not be updated");
                }
                #endregion
            }
            catch (Exception ex)
            {
                // Log error
                Console.WriteLine(ex.ToString());
            }
            finally
            {
                // ReSharper disable once ConstantConditionalAccessQualifier
                cc?.Dispose();
            }
        }

        /// <summary>
        /// Method doing actually something with the changes obtained via the web hook notification. 
        /// </summary>
        private static void DoWork(ClientContext cc, List changeList, Change change)
        {
            //Get Fields, if no fields then don't do anything
            var taxonomyTerms = CloudConfigurationManager.GetSetting("TaxonomyTermNames")
                ?.Split(",".ToCharArray(), StringSplitOptions.RemoveEmptyEntries);

            if (!taxonomyTerms.Any()) return;

            // Get the list item from the Change List
            // Note that this is the ID of the item in the list, not a reference to its position.
            var targetListItem = changeList.GetItemById(((ChangeItem) change).ItemId);
            cc.Load(targetListItem.File);

            // Get the File Binary Stream
            var streamResult = targetListItem.File.OpenBinaryStream();
            cc.ExecuteQueryRetry();

            string textFromStream;

            try
            {
                // Get Text Rendition of document binary
                var tika = new Tika();
                textFromStream = tika.ParseToString(streamResult.Value);
            }
            catch (Exception ex)
            {
                Trace.TraceWarning($"Tika Error: {ex}");
                return;
            }

            // Get Key phrases from text rendition
            var client = new TextAnalyticsClient();
            var result = client.GetStringPhrasesEntities(textFromStream).Result;

            // list of distinct key phrases
            var keyPhrases = result as string[] ?? result.ToArray();
            Trace.TraceInformation($"Key Phrases: {string.Join(",", keyPhrases)}");

            try
            {
                var results = taxonomyTerms.Select(term =>
                    SetTaxFieldValueAgainstKeyPhrase(cc, changeList, targetListItem, term, keyPhrases)).ToList();
            }
            catch (Exception ex)
            {
                Trace.TraceError($"Error: {ex}");
            }

        }

        private enum Result { Pass, Fail };

        /// <summary>
        /// 
        /// </summary>
        /// <param name="context"></param>
        /// <param name="changeList"></param>
        /// <param name="item"></param>
        /// <param name="taxonomyTermName"></param>
        /// <param name="keyPhrases"></param>
        /// <returns></returns>
        private static Result SetTaxFieldValueAgainstKeyPhrase(ClientContext context,List changeList, ListItem item, string taxonomyTermName, IEnumerable<string> keyPhrases)
        {
            try
            {
                Trace.TraceInformation($"TermName: {taxonomyTermName}");

                // Get Field and cast to taxonomy
                var txField = context.CastTo<TaxonomyField>(
                    changeList.Fields.GetByInternalNameOrTitle(
                        taxonomyTermName));
                context.Load(txField);
                context.ExecuteQuery();

                if(!txField.AllowMultipleValues) throw new ArgumentException("Can only set values on Taxonomy Fields that allow Multiple Values.");

                Trace.TraceInformation($"TermSetId: {txField.TermSetId}");

                var matchedTerms = GetMatchedKeywordsFromMms(keyPhrases,
                    GetTermsFromMms(context, txField.TermSetId));

                Trace.TraceInformation($"Matched Terms: {string.Join(",", matchedTerms.Select(term => term.Name))}");

                //Create Taxonomy Field Value and set on item.
                var taxonomyFieldValue = string.Join(";#", matchedTerms.Select(term => $"-1;#{term.Name}|{term.Id}"));
                var tx = new TaxonomyFieldValueCollection(context, taxonomyFieldValue, txField);
                txField.SetFieldValueByValueCollection(item, tx);

                item.SystemUpdate();
                context.ExecuteQuery();

                return Result.Pass;
            }
            catch (Exception ex)
            {
                // This occurs either because: 
                // 1. The column does not exist.
                // 2. The column is not a Taxonomy Field
                // 3. The column doesn't allow Multiple Values. The will error as it is not supported.
                Trace.TraceWarning($"Error: {ex}");
                return Result.Fail;
            }
        }

        /// <summary>
        /// Return all term from the term store of the specified term set id.
        /// </summary>
        /// <param name="cc">The Client Context</param>
        /// <param name="guidOfTermSet"></param>
        /// <returns>All term from the term store of the specified term set id.</returns>
        private static TermCollection GetTermsFromMms(ClientContext cc, Guid guidOfTermSet)
        {
            if (cc == null) throw new ArgumentNullException(nameof(cc));
            if (guidOfTermSet == null) throw new ArgumentNullException(nameof(guidOfTermSet));
            
            //
            // Get access to taxonomy CSOM.
            //
            var taxonomySession = TaxonomySession.GetTaxonomySession(cc);
            var termStore = taxonomySession.GetDefaultSiteCollectionTermStore();
            cc.Load(termStore,
                store => store.Id,
                store => store.Groups.Include(
                    groupArg => groupArg.Id,
                    groupArg => groupArg.Name
                )
            );
            cc.ExecuteQuery();

            //Requires you know the GUID of your Term Set, and the Name.
            var termSet = termStore.GetTermSet(guidOfTermSet);
            var terms = termSet.GetAllTerms();
            cc.Load(terms);
            cc.Load(termSet);
            cc.ExecuteQuery();

            return terms;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="keywords"></param>
        /// <param name="allTermsFromTermSet"></param>
        /// <returns></returns>
        private static List<Term> GetMatchedKeywordsFromMms(IEnumerable<string> keywords, TermCollection allTermsFromTermSet)
        {
            var keywordsToLower = keywords.Select(word => word.ToLower()).ToList();
            var matchedKeyWordsFromMms = new List<Term>();
            foreach (var term in allTermsFromTermSet)
            {
                if (!keywordsToLower.Contains(term.Name.ToLower())) continue;
                matchedKeyWordsFromMms.Add(term);
            }
            return matchedKeyWordsFromMms;
        }

        private void Cc_ExecutingWebRequest(object sender, WebRequestEventArgs e)
        {
            // Capture the OAuth access token since we want to reuse that one in our REST requests
            _accessToken = e.WebRequestExecutor.RequestHeaders.Get("Authorization").Replace("Bearer ", "");
        }
        #endregion
    }
}

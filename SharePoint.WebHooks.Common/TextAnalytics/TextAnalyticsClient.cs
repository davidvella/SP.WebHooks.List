using Microsoft.Azure;
using Microsoft.Azure.CognitiveServices.Language.TextAnalytics;
using Microsoft.Azure.CognitiveServices.Language.TextAnalytics.Models;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using SharePoint.WebHooks.Common.Utils;

namespace SharePoint.WebHooks.Common.TextAnalytics
{
    /// <summary>
    ///  The Text Analytics API is a suite of text analytics web services built with best-in-class
    ///  Microsoft machine learning algorithms. The API can be used to analyze unstructured
    ///  text for tasks such as sentiment analysis, key phrase extraction and language
    ///  detection. No training data is needed to use this API; just bring your text data.
    ///  This API uses advanced natural language processing techniques to deliver best
    ///  in class predictions. Further documentation can be found in https://docs.microsoft.com/en-us/azure/cognitive-services/text-analytics/overview
    /// </summary>
    class TextAnalyticsClient : Microsoft.Azure.CognitiveServices.Language.TextAnalytics.TextAnalyticsClient
    {
        private const int MaxDocumentInRequest = 1000;

        public TextAnalyticsClient() : base(new ApiKeyServiceClientCredentials(CloudConfigurationManager.GetSetting("TextAnalyticsApiKey")))
        {
            Endpoint = CloudConfigurationManager.GetSetting("TextAnalyticsEndpoint");
        }

        /// <summary>
        /// The API returns a list of strings denoting the key talking points in the
        /// input text.
        /// </summary>
        /// <remarks>
        /// See the &lt;a
        /// href="https://docs.microsoft.com/en-us/azure/cognitive-services/text-analytics/overview#supported-languages"&gt;Text
        /// Analytics Documentation&lt;/a&gt; for details about the languages that are
        /// supported by key phrase extraction.
        /// </remarks>
        /// <param name="text">
        /// The text to be passed to the Text Analytics Api
        /// </param>
        /// <param name="showStats">
        /// (optional) if set to true, response will contain input and document level
        /// statistics.
        /// </param>
        /// <param name="cancellationToken">The cancellation token.</param>
        public async Task<KeyPhraseBatchResult> KeyPhrasesStringAsync(string text, bool? showStats = null, CancellationToken cancellationToken = default)
        {
            return await this.KeyPhrasesAsync(showStats, new MultiLanguageBatchInput(await GetMultiLanguageInput(text, showStats, cancellationToken)), cancellationToken);
        }

        /// <summary>
        /// The API returns a list of recognized entities in a given document.
        /// </summary>
        /// <remarks>
        /// To get even more information on each recognized entity we recommend using
        /// the Bing Entity Search API by querying for the recognized entities names.
        /// See the &lt;a
        /// href="https://docs.microsoft.com/en-us/azure/cognitive-services/text-analytics/text-analytics-supported-languages"&gt;Supported
        /// languages in Text Analytics API&lt;/a&gt; for the list of enabled
        /// languages.
        /// </remarks>
        /// <param name="text">
        /// The text to be passed to the Text Analytics Api
        /// </param>
        /// <param name="showStats">
        /// (optional) if set to true, response will contain input and document level
        /// statistics.
        /// </param>
        /// <param name="cancellationToken">The cancellation token.</param>
        public async Task<EntitiesBatchResult> EntitiesStringAsync(string text, bool? showStats = null, CancellationToken cancellationToken = default)
        {
            return await this.EntitiesAsync(showStats, new MultiLanguageBatchInput(await GetMultiLanguageInput(text, showStats, cancellationToken)), cancellationToken);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="text"></param>
        /// <param name="showStats"></param>
        /// <param name="cancellationToken"></param>
        /// <returns></returns>
        public async Task<IEnumerable<string>> GetStringPhrasesEntities(string text, bool? showStats = null, CancellationToken cancellationToken = default)
        {
            var input = await GetMultiLanguageInput(text, showStats, cancellationToken);

            // Get Entities and KeyPhrases
            var entities = this.EntitiesAsync(showStats, new MultiLanguageBatchInput(input), cancellationToken);
            var phrases = this.KeyPhrasesAsync(showStats, new MultiLanguageBatchInput(input), cancellationToken);

            await Task.WhenAll(phrases, entities);

            // Select Strings from both results
            var phrasesStrings = phrases.Result.Documents.SelectMany(row => row.KeyPhrases).ToList();
            var entitiesStrings = entities.Result.Documents
                .SelectMany(row => row.Entities.Select(entity => entity.Name)).ToList();

            return phrasesStrings.Union(entitiesStrings).Distinct();
        }

        private async Task<IList<MultiLanguageInput>> GetMultiLanguageInput(string text, bool? showStats = null, CancellationToken cancellationToken = default)
        {
            // Get Sentences from string
            var sentences = StringUtils.FormatTextToSentences(text);

            // Get language of inputs
            var languageInput = sentences.Select((sentence, index) => new LanguageInput(id: index.ToString(), text: sentence)).Take(MaxDocumentInRequest).ToArray();

            var langResults = await this.DetectLanguageAsync(showStats, new LanguageBatchInput(languageInput), cancellationToken);

            // Make MultiLanguageInput
            var inputDocuments = langResults.Documents.Select(doc =>
                new MultiLanguageInput(doc.DetectedLanguages.OrderByDescending(lang => lang.Score).First().Iso6391Name,
                    doc.Id, languageInput[int.Parse(doc.Id)].Text)).ToArray();

            return inputDocuments;
        }
        
    }
        
}

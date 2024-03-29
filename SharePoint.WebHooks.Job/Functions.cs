using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Azure.WebJobs;
using SharePoint.WebHooks.Common;
using SharePoint.WebHooks.Common.Models;

namespace SharePoint.WebHooks.Job
{
    public class Functions
    {
        // This function will get triggered/executed when a new message is written 
        // on an Azure Queue called queue.
        public static void ProcessQueueMessage([QueueTrigger(ChangeManager.StorageQueueName)] NotificationModel notification, TextWriter log)
        {
            log.WriteLine($"Processing subscription {notification.SubscriptionId} for site {notification.SiteUrl}");
            var changeManager = new ChangeManager();
            changeManager.ProcessNotification(notification);
        }
    }
}

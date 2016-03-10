using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Mail;
using System.Web;
using System.Web.Services;
using System.Web.Services.Protocols;
using Newtonsoft.Json.Linq;
using NLog;
using NLog.Internal;
using ConfigurationManager = System.Configuration.ConfigurationManager;

namespace TFS2013Slack
{
    /// <summary>
    /// Summary description for WorkItemChanged
    /// </summary>
    [WebService(Namespace = "http://tempuri.org/")]
    [WebServiceBinding(ConformsTo = WsiProfiles.BasicProfile1_1)]
    [System.ComponentModel.ToolboxItem(false)]
    // To allow this Web Service to be called from script, using ASP.NET AJAX, uncomment the following line. 
    // [System.Web.Script.Services.ScriptService]
    public class WorkItemChanged : System.Web.Services.WebService
    {
        private readonly Logger _logger = LogManager.GetCurrentClassLogger();
        private string TFS_USERNAME;
        private string TFS_PASSWORD;

        public WorkItemChanged()
        {
            this.TFS_USERNAME = ConfigurationManager.AppSettings.Get("TfsUsername");
            this.TFS_PASSWORD = ConfigurationManager.AppSettings.Get("TfsPassword");
        }

        [SoapDocumentMethod(Action = "http://schemas.microsoft.com/TeamFoundation/2005/06/Services/Notification/03/Notify", RequestNamespace = "http://schemas.microsoft.com/TeamFoundation/2005/06/Services/Notification/03")]
        [WebMethod]
        //public void Notify (string eventXml, string tfsIdentityXml)
        public void Notify(string eventXml)
        {
            try
            {
                HandleEvent(eventXml);
            }
            catch (Exception ex)
            {
                _logger.Error(ex);
                throw;
            }
        }

        private string GetTaskUrl(string areaPath, string workItemId)
        {
            return string.Format("https://tfs.datalinksoftware.com/tfs/DefaultCollection{0}/_workitems#_a=edit&id={1}", areaPath, workItemId);
        }

        private void PostCompletedTaskToSlack(string areaPath, string id, string title, string changedBy, string parentId, string parentTitle, string parentWorkItemType)
        {
            var slackChannel = ConfigurationManager.AppSettings.Get("channel_" + areaPath);
            title = SanitizeTitle(title);
            parentTitle = SanitizeTitle(parentTitle);
            var taskUrl = GetTaskUrl(areaPath, id);
            var message = "";
            if (parentId != null)
            {
                var parentUrl = GetTaskUrl(areaPath, parentId);
                message += string.Format("<{0}|{1} {2}: {3}> >", parentUrl, parentWorkItemType, parentId, parentTitle);
            }
            message += string.Format("<{0}|Task {1}: {2}> completed by {3}", taskUrl, id, title, changedBy);

            PostMessageToSlack(message, slackChannel);
        }
        private void PostNewBugToSlack(string areaPath, string workItemId, string workItemTitle, string changedBy)
        {
            var slackChannel = ConfigurationManager.AppSettings.Get("channel_" + areaPath);
            var taskUrl = GetTaskUrl(areaPath, workItemId);
            workItemTitle = SanitizeTitle(workItemTitle);
            string message = "";
            message += string.Format("<{0}|Bug {1}: {2}> added by {3}", taskUrl, workItemId, workItemTitle, changedBy);
            PostMessageToSlack(message, slackChannel);
        }

        private void PostMessageToSlack(string message, string slackChannel)
        {
            var slackUrl = ConfigurationManager.AppSettings.Get("SlackUrl");
            var slack = new SlackClient(slackUrl);
            slack.PostMessage(message, slackChannel);
        }

        private Tuple<string, string, string> GetParent(string id)
        {
            string url =
                string.Format("https://tfs.datalinksoftware.com/tfs/DefaultCollection/_api/_wit/workitems?__v=5&ids={0}", id);
            using (WebClient client = new WebClient())
            {
                client.Credentials = new NetworkCredential(TFS_USERNAME, TFS_PASSWORD);
                var response = client.DownloadString(url);
                dynamic responseObject = JObject.Parse(response);
                var relations = responseObject.__wrappedArray[0].relations;
                foreach (var relation in relations)
                {
                    if (relation.LinkType == -2)
                    {
                        // found parent relation!
                        var parentId = relation.ID.ToString();
                        var details = GetWorkItemDetails(parentId);
                        return details;
                    }
                }
                return null;
            }
        }

        private Tuple<string, string, string> GetWorkItemDetails(string id)
        {
            string url =
                string.Format("https://tfs.datalinksoftware.com/tfs/DefaultCollection/_api/_wit/workitems?__v=5&ids={0}", id);
            using (WebClient client = new WebClient())
            {
                client.Credentials = new NetworkCredential(TFS_USERNAME, TFS_PASSWORD);
                var response = client.DownloadString(url);
                dynamic responseObject = JObject.Parse(response);
                string name = responseObject.__wrappedArray[0].fields["1"].ToString();
                string workItemType = responseObject.__wrappedArray[0].fields["25"].ToString();
                workItemType = workItemType.Replace("Product Backlog Item", "PBI");
                return new Tuple<string, string, string>(id, name, workItemType);
            }
        }

        private string SanitizeTitle(string title)
        {
            return title.Replace(">", "_");
        }

        private void HandleEvent(string eventXml)
        {
            var eventData = new TfsEventData(eventXml);
            var oldState = eventData.OldState;
            var state = eventData.State;

            if (eventData.WorkItemType == "Task" && oldState != state && state == "Done")
            {
                // state has changed
                var parent = GetParent(eventData.WorkItemId);
                PostCompletedTaskToSlack(eventData.AreaPath, eventData.WorkItemId, eventData.WorkItemTitle, eventData.ChangedBy, parent.Item1, parent.Item2, parent.Item3);
            }

            else if (eventData.WorkItemType == "Bug" && String.IsNullOrWhiteSpace(oldState) && state == "New")
            {
                // new bug was inserted
                PostNewBugToSlack(eventData.AreaPath, eventData.WorkItemId, eventData.WorkItemTitle, eventData.ChangedBy);
            }
        }


        [WebMethod]
        public string HelloWorld()
        {
            //var parent = GetParent("18657");
            //PostCompletedTaskToSlack("\\CareBookScrum", "9283", "my cool task", "travis collins", "9260", "parentTitle", "PBI");
            //PostNewBugToSlack("\\CareBookScrum", "9283", "my cool task", "travis collins");
            return "Hello World";
        }
    }
}

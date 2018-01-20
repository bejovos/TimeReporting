using System;
using System.Collections.Generic;
using Microsoft.TeamFoundation.Client;
using Microsoft.TeamFoundation.WorkItemTracking.Client;
using Microsoft.TeamFoundation.VersionControl.Client;
using System.Windows.Forms;

namespace OutlookAddIn1
{
    public class ConnectionWithTFS
    {
        private TfsTeamProjectCollection tfs = null;
        private TfsTeamProjectCollection GetTFS()
        {
            if (tfs == null)
            {
                Helpers.DebugInfo("Trying to connect to tfs.materialise.net:443");
                tfs = new TfsTeamProjectCollection(new System.Uri("https://tfs.materialise.net:443/tfs/Materialise%20Software"));                
            }                
            return tfs;
        }

        public IEnumerable<Changeset> GetMyChangesets()
        {   
            try
            {
                VersionControlServer vcs = GetTFS().GetService<VersionControlServer>();
                QueryHistoryParameters parameters = new QueryHistoryParameters("$/", RecursionType.Full);
                parameters.Author = vcs.AuthorizedUser;
                parameters.MaxResults = 5;
                parameters.IncludeDownloadInfo = false;
                parameters.IncludeChanges = false;
                Helpers.DebugInfo("Trying to get changesets for user: " + vcs.AuthorizedUser);
                return vcs.QueryHistory(parameters);
            }
            catch (Exception ex)
            {
                Helpers.DebugInfo("GetMyChangesets exception: " + ex.Message);
                tfs = null;
                throw;
            }
        }

        public WorkItemCollection GetWorkItem(int id)
        {
            try
            {
                WorkItemStore store = GetTFS().GetService<WorkItemStore>();
                var queryString = string.Format(@"SELECT [Id], [Title], [Completed Work], 
                [State], [Remaining Work] FROM WorkItems WHERE [Id] = '{0}'", id);
                return store.Query(queryString);
            }
            catch (Exception ex)
            {
                Helpers.DebugInfo("GetWorkItem exception: " + ex.Message);
                tfs = null;
                throw;
            }
        }

        public WorkItemCollection GetChildTasks(int id)
        {
            try
            {
                WorkItemStore store = GetTFS().GetService<WorkItemStore>(); 
                var queryString = string.Format(@"SELECT [Target].[Id] FROM WorkItemLinks WHERE " +
                        "[System.Links.LinkType] = 'System.LinkTypes.Hierarchy-Forward' AND " +
                        "Target.[System.WorkItemType] = 'Task' AND " +
                        "Source.[System.Id] = '{0}'", id);
                var query = new Query(store, queryString);

                queryString = @"SELECT [Id], [Title] FROM WorkItems WHERE [System.AssignedTo]= @me AND ([ID]='0'";
                foreach (WorkItemLinkInfo link in query.RunLinkQuery())
                    if (link.TargetId != 0)
                        queryString += string.Format(@" OR [Id] = '{0}'", link.TargetId);
                return store.Query(queryString + ")");
            }
            catch (Exception ex)
            {
                Helpers.DebugInfo("GetChildTasks exception: " + ex.Message);
                tfs = null;
                throw;
            }
        }

        public string IncreaseReportedTime(int id, int minutes, bool requireConfirmation)
        {
            Helpers.DebugInfo("Trying to add some time to TFS item: Id: " + id + " time: " + minutes + " minutes");

            foreach (WorkItem workItem in GetWorkItem(id))
            {
                workItem.PartialOpen();
                double time = minutes / (double)60;

                if ((string)workItem.Fields["State"].Value == "New")
                    workItem.Fields["State"].Value = "Active";

                double completed_work =
                    (workItem.Fields["Completed Work"].Value != null ?
                    (double)workItem.Fields["Completed Work"].Value : 0);

                string title = workItem.Title;
                string message = "Id: " + workItem.Id + Environment.NewLine +
                    "Title: " + workItem.Title + Environment.NewLine +
                    "Time: " + minutes / (double)60 + Environment.NewLine +
                    "Completed Work: " + completed_work + (time >= 0 ? " + " : " - ") + Math.Abs(time) + Environment.NewLine;

                completed_work += time;
                workItem.Fields["Completed Work"].Value = (completed_work > 0 ? completed_work : 0);

                bool is_closed = !workItem.Fields["Remaining Work"].IsEditable;
                if (!is_closed) 
                {
                    double remaining_work =
                    (workItem.Fields["Remaining Work"].Value != null ?
                    (double)workItem.Fields["Remaining Work"].Value : 0);
                     remaining_work -= time;
                     message += "Remaining Work: " + remaining_work + (time >= 0 ? " - " : " + ") + Math.Abs(time);
                     workItem.Fields["Remaining Work"].Value = (remaining_work > 0 ? remaining_work : 0);
                }

                if (requireConfirmation)
                {
                    if (MessageBox.Show(message, "Reporting item...", MessageBoxButtons.OKCancel) != DialogResult.OK)
                        throw new Exception("Aborted");
                    Helpers.DebugInfo("Confirmed by user: Id: " + id + " time: " + minutes + " minutes");
                }
                else Helpers.DebugInfo("Silently confirmed: Id: " + id + " time: " + minutes + " minutes");

                workItem.Save();
                workItem.Close();
                return title;
            }            
            throw new Exception("Work item not found!");
        }

    }
}

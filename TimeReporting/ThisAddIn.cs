using System;
using System.Collections.Generic;
using System.Windows.Forms;
using Outlook = Microsoft.Office.Interop.Outlook;
using System.Deployment.Application;

namespace OutlookAddIn1
{
    public partial class ThisAddIn
    {
        private static string _TimeReportedProperty = "TimeReportedWithThisAppointment"; //integer
        private static string _TaskIdProperty = "TaskIdUsedForReporting"; //integer
        private static string _PreviousAssignedSubjectProperty = "PreviousAssignedSubject"; //string
        private static string _RestoredFromDeletedProperty = "AppointmentRestoredFromDeleted"; //integer
        private static string _MyLastModificationTimeProperty = "MyLastModificationTime"; //date time
        private static string _CategoryName = "Reported";

        public ConnectionWithTFS _tfs = new ConnectionWithTFS();
        Timer _ReportingQueueTimer = new Timer();
        Timer _UpdateTimer = new Timer();        

        public Outlook.Explorer _Explorer = null;
        Outlook.Items _CalendarItems = null;
        Outlook.Items _DeletedItems = null;        
        public System.Collections.Specialized.StringCollection _RecentWorkItems = null;

        private static void DeleteProperty(Outlook.AppointmentItem aitem, string property)
        {
            if (aitem.UserProperties.Find(property) != null)
                aitem.UserProperties.Find(property).Delete();
        }

        private void AddNewRecentWorkItems(string subject)
        {
            Helpers.DebugInfo("Updating recent work items: " + subject);

            _RecentWorkItems.Remove(subject);
            _RecentWorkItems.Insert(0, subject);
            
            while (_RecentWorkItems.Count > 10)
                _RecentWorkItems.RemoveAt(_RecentWorkItems.Count - 1);

            Properties.Settings.Default["RecentWorkItems"] = _RecentWorkItems;
            Properties.Settings.Default.Save();
        }

        public void UpdateTimerLoop(object sender, EventArgs e)
        {
            _UpdateTimer.Stop();
            try
            {
                ApplicationDeployment ad = ApplicationDeployment.CurrentDeployment;
                if ((DateTime.Now - ad.TimeOfLastUpdateCheck).Hours > 6)
                {
                    Helpers.DebugInfo("Trying to get updates...");

                    var checkinfo = ad.CheckForDetailedUpdate();
                    if (checkinfo != null)
                    {
                        string shown_version = (string)Properties.Settings.Default["UpdateLastShownVersion"];
                        string new_version = checkinfo.AvailableVersion.ToString();
                        Helpers.DebugInfo("Available version: " + new_version);
                        if (new_version != shown_version)
                        {
                            Properties.Settings.Default["UpdateLastShownVersion"] = new_version;
                            Properties.Settings.Default.Save();
                            MessageBox.Show(
                                "New version is available (" + new_version + ") at" + Environment.NewLine +
                                ApplicationDeployment.CurrentDeployment.UpdateLocation,
                                "TimeReporting " + ApplicationDeployment.CurrentDeployment.CurrentVersion.ToString(), 
                                MessageBoxButtons.OK);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Helpers.DebugInfo("Exception during updating: " + ex.Message);
            }
            _UpdateTimer.Start();
        }

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            try
            {
                var key = Microsoft.Win32.Registry.CurrentUser.OpenSubKey("Software", true).CreateSubKey("TimeReporting");
                RegistryHelpers.ReadValue<string>(ref Settings.debugFile, key, "DebugFile");
                RegistryHelpers.ReadValue<bool>(ref Settings.showMessageBoxWhenAppointementCreated, key, "ShowMessageBoxWhenAppointementCreated");
                RegistryHelpers.ReadValue<bool>(ref Settings.showMessageBoxWhenAppointementEdited, key, "ShowMessageBoxWhenAppointementEdited");
                RegistryHelpers.ReadValue<bool>(ref Settings.showMessageBoxWhenAppointementDeleted, key, "ShowMessageBoxWhenAppointementDeleted");
                RegistryHelpers.ReadValue<bool>(ref Settings.includeMyChildTasksInCommitList, key, "IncludeMyChildTasksInCommitList");
            }
            catch (Exception ex)
            { 
                Helpers.DebugInfo("Startup: Registry reading exception: " + ex.Message);                
            }

            Helpers.DebugInfo("Startup...");

            //try
            //{
            //    if (ApplicationDeployment.IsNetworkDeployed)
            //    {
            //        string current_version = "";
            //        _UpdateTimer.Interval = 60 * 60 * 1000; // every hour
            //        _UpdateTimer.Tick += UpdateTimerLoop;
            //        _UpdateTimer.Start();

            //        current_version = ApplicationDeployment.CurrentDeployment.CurrentVersion.ToString();
            //        Helpers.DebugInfo("Current version: " + current_version);

            //        if (Properties.Settings.Default["UpdateLastShownVersion"] == null ||
            //            (string)Properties.Settings.Default["UpdateLastShownVersion"] == "")
            //        {
            //            Properties.Settings.Default["UpdateLastShownVersion"] = current_version;
            //            Properties.Settings.Default.Save();
            //        }
            //    }
            //}
            //catch (Exception ex)
            //{
            //    Helpers.DebugInfo("Startup: Current deployment exception: " + ex.Message);
            //}

            try
            {
                _Explorer = this.Application.ActiveExplorer();
                Outlook.MAPIFolder calendarFolder = _Explorer.Session.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderCalendar);
                _CalendarItems = calendarFolder.Items;
                _CalendarItems.ItemAdd += MyAfterItemAdded;
                _CalendarItems.ItemChange += MyAfterItemChanged;

                Outlook.MAPIFolder deletedFolder = _Explorer.Session.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderDeletedItems);
                _DeletedItems = deletedFolder.Items;
                _DeletedItems.ItemAdd += MyAfterItemRemoved;

                if (_Explorer.Session.Categories[_CategoryName] == null)
                    _Explorer.Session.Categories.Add(_CategoryName);
                _Explorer.Session.Categories[_CategoryName].Color = Outlook.OlCategoryColor.olCategoryColorGreen;
                _RecentWorkItems = Properties.Settings.Default.RecentWorkItems;
                if (_RecentWorkItems == null)
                    _RecentWorkItems = new System.Collections.Specialized.StringCollection();
                _ReportingQueueTimer.Tick += ReportingQueueLoop;
                _ReportingQueueTimer.Interval = 1;
                _ReportingQueueTimer.Stop();
            }
            catch (Exception ex)
            {
                Helpers.DebugInfo("Startup: Callback handlers exception: " + ex.Message);
                throw;
            }

            if (Settings.debugFile != "")
                MessageBox.Show(@"DebugFile is enabled (HKEY_CURRENT_USER\Software\TimeReporting\DebugFile): " + Environment.NewLine + Settings.debugFile);
        }

        class DelayedReporting
        {
            public enum Mode { New, Edit, Remove };
            public Outlook.AppointmentItem item = null;
            public Mode mode = Mode.New;
            public DelayedReporting(Outlook.AppointmentItem i_item, Mode i_mode)
            {
                item = i_item;
                mode = i_mode;
            }
        }
        List<DelayedReporting> _DelayedReportingQueue = new List<DelayedReporting>();

        #region Callback handlers

        public void MyAfterItemChanged(object obj)
        {
            Outlook.AppointmentItem aitem = obj as Outlook.AppointmentItem;
            if (aitem == null || aitem.IsRecurring ||
                aitem.MeetingStatus != Outlook.OlMeetingStatus.olNonMeeting)
                return;

            Helpers.DebugInfo("ItemChanged callback is fired with appointment: Id: " + aitem.GlobalAppointmentID +
                " Subject: " + aitem.Subject + " Duration: " + aitem.Duration);

            if (aitem.UserProperties.Find(_TimeReportedProperty) == null ||
                aitem.UserProperties.Find(_TaskIdProperty) == null)
            {
                _DelayedReportingQueue.Add(new DelayedReporting(aitem, DelayedReporting.Mode.New));
                _ReportingQueueTimer.Start();
                return;
            }
            if (aitem.Categories != _CategoryName)
            {
                Helpers.DebugInfo("Reverted category");
                aitem.Categories = _CategoryName;
                aitem.Save();
            }                
            if (aitem.Subject != aitem.UserProperties.Find(_PreviousAssignedSubjectProperty).Value)
            {
                Helpers.DebugInfo("Reverted subject: old: " + aitem.Subject + " new: " + aitem.UserProperties.Find(_PreviousAssignedSubjectProperty).Value);
                aitem.Subject = aitem.UserProperties.Find(_PreviousAssignedSubjectProperty).Value;
                aitem.Save();
            }
            if (aitem.Duration != aitem.UserProperties.Find(_TimeReportedProperty).Value)
            {                
                _DelayedReportingQueue.Add(new DelayedReporting(aitem, DelayedReporting.Mode.Edit));
                _ReportingQueueTimer.Start();
            }            
        }

        public void MyAfterItemRemoved(object obj)
        {
            Outlook.AppointmentItem aitem = obj as Outlook.AppointmentItem;
            if (aitem == null)            
                return;

            Helpers.DebugInfo("ItemRemoved callback is fired with appointment: Id: " + aitem.GlobalAppointmentID +
                " Subject: " + aitem.Subject + " Duration: " + aitem.Duration);

            if (aitem.UserProperties.Find(_TimeReportedProperty) == null ||
                aitem.UserProperties.Find(_TaskIdProperty) == null)
                return;
            aitem.Categories = "";
            aitem.Save();

            _DelayedReportingQueue.Add(new DelayedReporting(aitem, DelayedReporting.Mode.Remove));
            _ReportingQueueTimer.Start();
        }

        public void MyAfterItemAdded(object obj)
        {
            Outlook.AppointmentItem aitem = obj as Outlook.AppointmentItem;
            if (aitem == null || aitem.IsRecurring ||
                aitem.MeetingStatus != Outlook.OlMeetingStatus.olNonMeeting)
                return;

            Helpers.DebugInfo("ItemAdded callback is fired with appointment: Id: " + aitem.GlobalAppointmentID +
                " Subject: " + aitem.Subject + " Duration: " + aitem.Duration);

            if (aitem.UserProperties.Find(_RestoredFromDeletedProperty) != null)
                aitem.UserProperties.Find(_RestoredFromDeletedProperty).Delete();                
            else
            {
                if (aitem.UserProperties.Find(_TaskIdProperty) != null ||
                    aitem.UserProperties.Find(_TimeReportedProperty) != null ||
                    aitem.UserProperties.Find(_PreviousAssignedSubjectProperty) != null)
                    aitem.Categories = "";
                DeleteProperty(aitem, _TaskIdProperty);
                DeleteProperty(aitem, _TimeReportedProperty);
                DeleteProperty(aitem, _PreviousAssignedSubjectProperty);
            }
            aitem.Save();            

            // outlook sends events in really random order => trying to speed-up it
            MyAfterItemChanged(obj);
        }

        #endregion

        private void UpdateAppointment(Outlook.AppointmentItem aitem, int id, 
            string user_title, string work_item_title, int duration)
        {
            string subject_with_time = Helpers.GenerateSubject(id, user_title, true);
            aitem.Categories = _CategoryName;
            aitem.Subject = subject_with_time;
            aitem.Duration = duration;
            aitem.ReminderSet = false;
            aitem.UserProperties.Add(_TimeReportedProperty, Outlook.OlUserPropertyType.olInteger).Value = duration;
            aitem.UserProperties.Add(_TaskIdProperty, Outlook.OlUserPropertyType.olInteger).Value = id;
            aitem.UserProperties.Add(_PreviousAssignedSubjectProperty, Outlook.OlUserPropertyType.olText).Value = subject_with_time;
            aitem.Save();
            AddNewRecentWorkItems(Helpers.GenerateSubject(id, work_item_title));
        }

        // main processing
        private void ProcessTask(DelayedReporting task)
        {
            Outlook.AppointmentItem aitem = task.item;
            if (aitem.Subject == null)
                return;

            if (task.mode == DelayedReporting.Mode.Remove)
            {
                if (aitem.UserProperties.Find(_TimeReportedProperty) == null ||
                    aitem.UserProperties.Find(_TaskIdProperty) == null)
                    return;
                int id = aitem.UserProperties.Find(_TaskIdProperty).Value;
                int duration = aitem.UserProperties.Find(_TimeReportedProperty).Value;
                string subject = aitem.UserProperties.Find(_PreviousAssignedSubjectProperty).Value;
                try
                {
                    _tfs.IncreaseReportedTime(id, -duration, Settings.showMessageBoxWhenAppointementDeleted);
                }
                catch (Exception ex)
                {
                    Outlook.AppointmentItem newItem = (Outlook.AppointmentItem)this.Application.CreateItem(Outlook.OlItemType.olAppointmentItem);
                    newItem.Start = aitem.Start;
                    newItem.Subject = aitem.Subject;
                    newItem.Duration = duration;
                    newItem.Categories = _CategoryName;
                    newItem.UserProperties.Add(_TaskIdProperty, Outlook.OlUserPropertyType.olInteger).Value = id;
                    newItem.UserProperties.Add(_TimeReportedProperty, Outlook.OlUserPropertyType.olInteger).Value = duration;
                    newItem.UserProperties.Add(_PreviousAssignedSubjectProperty, Outlook.OlUserPropertyType.olText).Value = subject;

                    // used to prevent firing "new item event"
                    newItem.UserProperties.Add(_RestoredFromDeletedProperty, Outlook.OlUserPropertyType.olInteger).Value = 1;
                    newItem.Save();
                    if (ex.Message != "Aborted")
                        MessageBox.Show(ex.Message, "TimeReporting");
                }

                DeleteProperty(aitem, _TaskIdProperty);
                DeleteProperty(aitem, _TimeReportedProperty);
                DeleteProperty(aitem, _PreviousAssignedSubjectProperty);
                aitem.Categories = "";
                aitem.Subject = Helpers.GenerateSubject(id, "");
                aitem.Save();
            }
            else if (task.mode == DelayedReporting.Mode.New)
            {
                if (aitem.UserProperties.Find(_TimeReportedProperty) != null &&
                    aitem.UserProperties.Find(_TaskIdProperty) != null)
                    return;
                if (aitem.UserProperties.Find(_MyLastModificationTimeProperty) != null &&
                    aitem.UserProperties.Find(_MyLastModificationTimeProperty).Value >= aitem.LastModificationTime)
                    return;
                int id = Helpers.ParseID(aitem.Subject);
                if (id == 0)
                    return;

                try
                {
                    int duration = aitem.Duration;
                    string user_title = Helpers.ParseTitle(aitem.Subject);
                    if (user_title != "")
                        user_title += " ";
                    string work_item_title = _tfs.IncreaseReportedTime(id, duration, Settings.showMessageBoxWhenAppointementCreated);
                    UpdateAppointment(aitem, id, user_title + work_item_title, work_item_title, duration);
                }
                catch (Exception ex)
                {
                    aitem.UserProperties.Add(_MyLastModificationTimeProperty, Outlook.OlUserPropertyType.olDateTime).Value
                        = DateTime.Now.AddSeconds(5);
                    aitem.Save();

                    if (ex.Message != "Aborted")
                        MessageBox.Show(ex.Message, "TimeReporting");
                }
            }
            else if (task.mode == DelayedReporting.Mode.Edit)
            {
                if (aitem.UserProperties.Find(_TimeReportedProperty) == null ||
                    aitem.UserProperties.Find(_TaskIdProperty) == null)
                    return;
                int duration = aitem.UserProperties.Find(_TimeReportedProperty).Value;
                int new_duration = aitem.Duration;
                if (new_duration == duration)
                    return;
                try
                {
                    int id = aitem.UserProperties.Find(_TaskIdProperty).Value;                    
                    string user_title = Helpers.ParseTitleWithTime(aitem.Subject);
                    string work_item_title = _tfs.IncreaseReportedTime(id, new_duration - duration, Settings.showMessageBoxWhenAppointementEdited);
                    UpdateAppointment(aitem, id, user_title, work_item_title, new_duration);
                }
                catch (Exception ex)
                {
                    aitem.Duration = aitem.UserProperties.Find(_TimeReportedProperty).Value;
                    aitem.Save();
                    if (ex.Message != "Aborted")
                        MessageBox.Show(ex.Message, "TimeReporting");
                }
            }  
        }

        bool queue_in_processing = false;
        public void ReportingQueueLoop(object sender, EventArgs e)
        {
            _ReportingQueueTimer.Stop();
            if (queue_in_processing == true)
                return ;

            queue_in_processing = true;

            while (_DelayedReportingQueue.Count > 0)
            {
                try
                {
                    DelayedReporting task = _DelayedReportingQueue[0];
                    Helpers.DebugInfo("ReportingQueue is going to process appointment: Id: " + task.item.GlobalAppointmentID +
                        " Subject: " + task.item.Subject + " Duration: " + task.item.Duration);                    
                    _DelayedReportingQueue.RemoveAt(0);
                    ProcessTask(task);
                }
                catch (Exception ex)
                {
                    Helpers.DebugInfo("ReportingQueue exception: " + ex.Message);
                }
            }    
            queue_in_processing = false;
        }


        #region VSTO generated code

        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
        }

        protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            return new MyContextMenu();
        }
        
        #endregion
    }
}


using System;
using System.IO;
using System.Linq;
using System.Text;
using System.Reflection;
using System.Runtime.InteropServices;
using Microsoft.TeamFoundation.WorkItemTracking.Client;
using Microsoft.TeamFoundation.VersionControl.Client;
using Office = Microsoft.Office.Core;
using Outlook = Microsoft.Office.Interop.Outlook;
using System.Collections.Generic;

namespace OutlookAddIn1
{
    [ComVisible(true)]
    public class MyContextMenu : Office.IRibbonExtensibility
    {
        private Office.IRibbonUI ribbon;

        public MyContextMenu()
        {
        }

        public bool GetMultipleItemContextMenuVisible(Office.IRibbonControl control)
        {
            return true;
        }

        public bool GetPasteFromClipboardVisible(Office.IRibbonControl control)
        {
            return (Helpers.GetIDFromClipboard() > 0);
        }
        
        private void UpdateSelectedAppointmentsOrCreateNew(string tag)
        {
            Outlook.View view = Globals.ThisAddIn._Explorer.CurrentView as Outlook.View;
            if (view.ViewType != Outlook.OlViewType.olCalendarView)
                return;
            Outlook.Folder folder = Globals.ThisAddIn._Explorer.CurrentFolder as Outlook.Folder;
            Outlook.CalendarView calView = view as Outlook.CalendarView;

            if (Globals.ThisAddIn._Explorer.Selection.Count != 0)
            {
                foreach (object obj in Globals.ThisAddIn._Explorer.Selection)
                {
                    Outlook.AppointmentItem aitem = obj as Outlook.AppointmentItem;
                    if (aitem == null)
                        continue;

                    if (aitem.IsRecurring || 
                        aitem.MeetingStatus != Outlook.OlMeetingStatus.olNonMeeting)
                    {
                        // clone this appointment                        
                        DateTime dateStart = aitem.Start;
                        DateTime dateEnd = aitem.End;
                        aitem = folder.Items.Add("IPM.Appointment") as Outlook.AppointmentItem;
                        aitem.ReminderSet = false;
                        aitem.Start = dateStart;
                        aitem.End = dateEnd;
                    }

                    if (aitem.Subject == null)
                        aitem.Subject = tag;
                    else aitem.Subject = tag + " " + aitem.Subject.Trim();
                    aitem.Save();
                }
            }
            else
            {                
                DateTime dateStart = calView.SelectedStartTime;
                DateTime dateEnd = calView.SelectedEndTime;
                Outlook.AppointmentItem aitem = folder.Items.Add("IPM.Appointment") as Outlook.AppointmentItem;
                aitem.ReminderSet = false;
                aitem.Start = dateStart;
                aitem.End = dateEnd;
                aitem.Subject = tag;
                aitem.Save();
            }
        }

        public void OnAction(Office.IRibbonControl control)
        {   
            UpdateSelectedAppointmentsOrCreateNew(control.Tag);
        }

        public void PasteFromClipboard(Office.IRibbonControl control)
        {
            int id = Helpers.GetIDFromClipboard();
            if (id <= 0)
                return;
            UpdateSelectedAppointmentsOrCreateNew(Helpers.GenerateSubject(id, ""));
        }

        public string GetContentLatestReporting(Office.IRibbonControl control)
        {
            StringBuilder builder = new StringBuilder(@"<menu xmlns=""http://schemas.microsoft.com/office/2006/01/customui"" >");
            int index = 0;
            foreach (string text in Globals.ThisAddIn._RecentWorkItems)
            {
                int id = Helpers.ParseID(text);
                if (id == 0)
                    continue;
                string tag = Helpers.GenerateSubject(id, "");
                ++index;
                builder.Append(@"<button id=""button" + index + @""" label=""" + 
                    System.Security.SecurityElement.Escape(text) + @""" onAction=""OnAction"" tag=""" + tag + @""" />");
            }
            builder.Append(@"</menu>");
            return builder.ToString();
        }

        #region IRibbonExtensibility Members

        public string GetCustomUI(string ribbonID)
        {
            return GetResourceText("OutlookAddIn1.MyContextMenu.xml");
        }

        #endregion

        #region Ribbon Callbacks
        //Create callback methods here. For more information about adding callback methods, visit http://go.microsoft.com/fwlink/?LinkID=271226

        public void Ribbon_Load(Office.IRibbonUI ribbonUI)
        {
            this.ribbon = ribbonUI;
        }

        #endregion

        #region Helpers

        private static string GetResourceText(string resourceName)
        {
            Assembly asm = Assembly.GetExecutingAssembly();
            string[] resourceNames = asm.GetManifestResourceNames();
            for (int i = 0; i < resourceNames.Length; ++i)
            {
                if (string.Compare(resourceName, resourceNames[i], StringComparison.OrdinalIgnoreCase) == 0)
                {
                    using (StreamReader resourceReader = new StreamReader(asm.GetManifestResourceStream(resourceNames[i])))
                    {
                        if (resourceReader != null)
                        {
                            return resourceReader.ReadToEnd();
                        }
                    }
                }
            }
            return null;
        }

        #endregion
    }
}

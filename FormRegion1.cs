using Microsoft.Office.Interop.Outlook;
using System;

namespace OutlookAddInGlobalAppointmentIdTest
{
    partial class FormRegion1
    {
        #region Form Region Factory 

        [Microsoft.Office.Tools.Outlook.FormRegionMessageClass(Microsoft.Office.Tools.Outlook.FormRegionMessageClassAttribute.Appointment)]
        [Microsoft.Office.Tools.Outlook.FormRegionName("OutlookAddInGlobalAppointmentIdTest.FormRegion1")]
        public partial class FormRegion1Factory
        {
            // Occurs before the form region is initialized.
            // To prevent the form region from appearing, set e.Cancel to true.
            // Use e.OutlookItem to get a reference to the current Outlook item.
            private void FormRegion1Factory_FormRegionInitializing(object sender, Microsoft.Office.Tools.Outlook.FormRegionInitializingEventArgs e)
            {
            }
        }

        #endregion
        private AppointmentItem _lastAppointmentItem = null;

        // Occurs before the form region is displayed.
        // Use this.OutlookItem to get a reference to the current Outlook item.
        // Use this.OutlookFormRegion to get a reference to the form region.
        private void FormRegion1_FormRegionShowing(object sender, System.EventArgs e)
        {
            ShowGlobalAppointmentId();
            ((AppointmentItem)this.OutlookItem).PropertyChange += FormRegion1_PropertyChange;
        }

        private void FormRegion1_PropertyChange(string Name)
        {
            ShowGlobalAppointmentId();
        }


        // Occurs when the form region is closed.
        // Use this.OutlookItem to get a reference to the current Outlook item.
        // Use this.OutlookFormRegion to get a reference to the form region.
        private void FormRegion1_FormRegionClosed(object sender, System.EventArgs e)
        {
            _lastAppointmentItem = null;
        }

        private void ShowGlobalAppointmentId()
        {
            string message = "No GlobalAppointmentId";
            try
            {
                AppointmentItem item = (AppointmentItem)this.OutlookItem;
                message = item.GlobalAppointmentID;
                if (String.IsNullOrEmpty(message)) { }
            }
            catch (System.Exception ex)
            {
                message = ex.Message;
            }
            System.Action action = new System.Action(() =>
            {
                textBoxGlobalAppointmentId.Text = message;
            });
            if (textBoxGlobalAppointmentId.InvokeRequired)
                textBoxGlobalAppointmentId.Invoke(action);
            else
                action();
        }

        private void ShowiCalUID()
        {
            string message = "No iCalUID";
            try
            {
                AppointmentItem item = (AppointmentItem)this.OutlookItem;
                message = "Not implemented";
            }
            catch (System.Exception ex)
            {
                message = ex.Message;
            }
            System.Action action = new System.Action(() =>
            {
                textBoxGlobalAppointmentId.Text = message;
            });
            if (textBoxGlobalAppointmentId.InvokeRequired)
                textBoxGlobalAppointmentId.Invoke(action);
            else
                action();
        }

        private void buttonRefresh_Click(object sender, EventArgs e)
        {
            ShowGlobalAppointmentId();
        }

        private void buttonWatch_Click(object sender, EventArgs e)
        {
            FormIdWatcher watcher = new FormIdWatcher((AppointmentItem)this.OutlookItem);
            ((AppointmentItem)this.OutlookItem).Save();
            watcher.Show();
            Globals.ThisAddIn.Application.ActiveInspector().Close(OlInspectorClose.olSave);
        }
    }
}

using Microsoft.Office.Interop.Outlook;
using System;
using System.Windows.Forms;

namespace OutlookAddInGlobalAppointmentIdTest
{
    public partial class FormIdWatcher : Form
    {
        private AppointmentItem _watchedAppointment;
        private Folder _appointmentFolder;
        private int _itemChangeCount = 0;
        private int _triggerChangeCount = 4;

        public FormIdWatcher(AppointmentItem WatchedAppointment)
        {
            InitializeComponent();
            _watchedAppointment = WatchedAppointment;
            _watchedAppointment.PropertyChange += _watchedAppointment_PropertyChange;
            _watchedAppointment.Unload += _watchedAppointment_Unload;
            _watchedAppointment.AfterWrite += _watchedAppointment_AfterWrite;
            if (_watchedAppointment.Saved)
                _triggerChangeCount--;

            _appointmentFolder = _watchedAppointment.Parent;
            _appointmentFolder.Items.ItemChange += Items_ItemChange;
        }

        private void Items_ItemChange(object Item)
        {
            if (((AppointmentItem)Item).Subject==_watchedAppointment.Subject)
            {
                _watchedAppointment.Close(OlInspectorClose.olDiscard);
                _watchedAppointment = (AppointmentItem)Item;
                _itemChangeCount++;
                if (_itemChangeCount == _triggerChangeCount && _watchedAppointment.GlobalAppointmentID==null)
                    // At this point, we should have GlobalAppointmentId - reopening the item allows us to read it
                    buttonOpen_Click(this, null);
                else
                    ShowGlobalAppointmentId($"ItemChange detected ({_itemChangeCount})");
            }
        }

        private void _watchedAppointment_AfterWrite()
        {
            string message = "Appointment written";
            System.Action action = new System.Action(() =>
            {
                textBoxGlobalAppointmentId.Text = message;
            });
            if (textBoxGlobalAppointmentId.InvokeRequired)
                textBoxGlobalAppointmentId.Invoke(action);
            else
                action();
        }

        private void _watchedAppointment_Unload()
        {
            _watchedAppointment = null;
            string message = "Appointment unloaded";
            System.Action action = new System.Action(() =>
            {
                textBoxGlobalAppointmentId.Text = message;
            });
            if (textBoxGlobalAppointmentId.InvokeRequired)
                textBoxGlobalAppointmentId.Invoke(action);
            else
                action();
        }

        private void _watchedAppointment_PropertyChange(string Name)
        {
            ShowGlobalAppointmentId($"{Name} property changed");
        }

        private void ShowGlobalAppointmentId(string message = null)
        {
            if (String.IsNullOrEmpty(message))
                message = textBoxGlobalAppointmentId.Text;
            if (String.IsNullOrEmpty(message))
                message = "No GlobalAppointmentId";

            try
            {
                if (!String.IsNullOrEmpty(_watchedAppointment.GlobalAppointmentID))
                    message = _watchedAppointment.GlobalAppointmentID;
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

        private void buttonOpen_Click(object sender, EventArgs e)
        {
            _watchedAppointment.Display();
            ShowGlobalAppointmentId();
        }
    }
}

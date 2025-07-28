using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Outlook = Microsoft.Office.Interop.Outlook;
using Microsoft.Office.Interop.Outlook;
using System.Reflection;
using System.Diagnostics;

namespace LishcaAddIn
{
    public partial class MeetingControl : UserControl
    {
        MessageText message;
        protected static DatabaseManager database = DatabaseManager.Instance;
        private Contact[] contacts;

        public MeetingControl()
        {
            InitializeComponent();
        }

        public void LoadAppointItem(Outlook.AppointmentItem appointment)
        {
            listInv.Items.Clear();
            backgroundWorkerGetContacts.RunWorkerAsync(appointment);
        }

        private void PopulateTitleComboBox()
        {
            titleComboBox.Items.Add(Title.New);
            titleComboBox.Items.Add(Title.Change);
            titleComboBox.Items.Add(Title.ChangeDate);
            titleComboBox.Items.Add(Title.ChangeHours);
            titleComboBox.Items.Add(Title.Check);
            titleComboBox.Items.Add(Title.NewInvitees);
            titleComboBox.Items.Add(Title.Postponed);
            titleComboBox.Items.Add(Title.UnPostpone);
            titleComboBox.Items.Add(Title.Cancelled);
        }

        private void ListInv_ItemCheck(object sender, ItemCheckEventArgs e)
        {
            //Contact contact = this.contacts[e.Index];
            Contact contact = this.listInv.Items[e.Index] as Contact;
            if (e.NewValue == CheckState.Checked)
            {
                message.phoneNums.Add(contact.number);
            }
            else
            {
                message.phoneNums.RemoveAll(num => num == contact.number);
                listInv.ItemCheck -= ListInv_ItemCheck;
                DeselectFromCBAll(contact);
                listInv.ItemCheck += ListInv_ItemCheck;
            }
            UpdateQR();
        }

        private void DeselectFromCBAll(Contact contact)
        {
            for (int i = 0; i < this.listInv.Items.Count; i++)
            {
                Contact compare = this.listInv.Items[i] as Contact;
                if (compare.number == contact.number)
                {
                    this.listInv.SetItemChecked(i, false);
                }
            }
        }

        private void PopulateContactsListBox()
        {
            for (int i = 0; i < contacts.Length; i++)
            {
                this.listInv.Items.Add(contacts[i], contacts[i].pkidim == null);
                if (contacts[i].pkidim != null)
                {
                    for (int j = 0; j < contacts[i].pkidim.Length; j++) // should be recursive
                    {
                        this.listInv.Items.Add(contacts[i].pkidim[j], true);
                    }
                }
                else
                {
                    this.listInv.SetItemChecked(i, true);
                }
            }
        }

        private string[] LoadAddressesFromAppointItem(Outlook.AppointmentItem appointment)
        {
            const string PR_SMTP_ADDRESS = "http://schemas.microsoft.com/mapi/proptag/0x39FE001E";
            string[] result = new string[Appointment.Recipients.Count - 1];
            for (int i = 1; i < Appointment.Recipients.Count; i++) // start with 1, for first Recipient is the Inviter (ignored)
            {
                Outlook.PropertyAccessor pa = Appointment.Recipients[i + 1].PropertyAccessor;
                result[i - 1] = pa.GetProperty(PR_SMTP_ADDRESS).ToString();
            }
            return result;
        }

        private void UpdateQR()
        {
            this.pictureBox1.Image = message.GetQrCode();
        }

        private void EditText_Click(object sender, EventArgs e)
        {
            using (RichTextEditor form = new RichTextEditor
                (LocationCB.Checked ? message.FullMessageWithLocation : message.FullMessage))
            {
                form.Enabled = true;
                if (form.ShowDialog() == DialogResult.OK)
                {
                    this.pictureBox1.Image = message.UpdateQrCode(form.ReturnValue);
                }
            }
        }

        private void titleComboBox_TextChanged(object sender, EventArgs e)
        {
            this.message.title.text = titleComboBox.Text;
            UpdateQR();
        }

        private void CheckBoxAll_CheckStateChanged(object sender, EventArgs e)
        {
            for (int i = 0; i < listInv.Items.Count; i++)
            {
                this.listInv.SetItemCheckState(i, CheckBoxAll.CheckState);
            }
        }

        private void LocationCB_CheckStateChanged(object sender, EventArgs e)
        {
            if (LocationCB.CheckState == CheckState.Checked)
            {
                this.pictureBox1.Image = message.UpdateQrCode(message.FullMessageWithLocation);
            }
            else if (LocationCB.CheckState == CheckState.Unchecked)
            {
                this.pictureBox1.Image = message.UpdateQrCode(message.FullMessage);
            }
        }

        private async void BackgroundWorkerGetContacts_DoWorkAsync(object sender, DoWorkEventArgs e)
        {
            Outlook.AppointmentItem appointment = (Outlook.AppointmentItem)e.Argument;
            string[] addresses = LoadAddressesFromAppointItem(appointment);

            var task1 = database.GetContactsFromAddresses(addresses);
            task1.Wait();
            this.contacts = task1.Result; // RETRIEVES DATA FROM GUI, WHICH TAKES TIME, returns contact only if not found
            message = new MessageText(appointment, this.contacts);
        }

        private void BackgroundWorkerGetContacts_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            PopulateContactsListBox();
            PopulateTitleComboBox();
            UpdateQR();
            listInv.ItemCheck += ListInv_ItemCheck;
        }
    }
}

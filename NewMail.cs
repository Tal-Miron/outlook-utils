using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace LishcaAddIn
{
    public partial class NewMail : Form
    {
        private string[] addresses;
        protected static DatabaseManager database = DatabaseManager.Instance;
        public BindingList<Contact> missingContacts { get; private set; }
        private Dictionary<string, string> addressesToDisplayName = new Dictionary<string, string>();

        public NewMail(Outlook.AppointmentItem appointment)
        {
            addresses = LoadAddressesFromAppointItem(appointment);
            if (addresses.Length <= 0)
            {
                this.DialogResult = DialogResult.Abort;
                this.Close();
                return;
            }
            InitializeComponent();
            loadContactsBW.RunWorkerAsync();
        }

        private void InitializeBinding()
        {
            ListMails.DataSource = missingContacts;
            ListMails.DisplayMember = "displayName";
            ListMails.ValueMember = "displayName";

            UnitsList.SelectedIndex = 0;
            Rank[] ranks = InfoManager.AllRanks;
            RanksList.DataSource = ranks;
            RanksList.SelectedIndex = 0;
        }

        private string[] LoadAddressesFromAppointItem(Outlook.AppointmentItem appointment)
        {
            const string PR_SMTP_ADDRESS = "http://schemas.microsoft.com/mapi/proptag/0x39FE001E";
            
            if (appointment.Recipients.Count < 1)
                return new string[0];
            
            string[] result = new string[appointment.Recipients.Count - 1];
            
            for (int i = 1; i < appointment.Recipients.Count; i++) // start with 1, for first Recipient is the Inviter (ignored)
            {
                Outlook.PropertyAccessor pa = appointment.Recipients[i + 1].PropertyAccessor;
                result[i - 1] = pa.GetProperty(PR_SMTP_ADDRESS).ToString();
                if (!addressesToDisplayName.ContainsKey(result[i - 1]))
                    addressesToDisplayName.Add(result[i - 1], appointment.Recipients[i + 1].Name);
            }
            return result;
        }

        private void ListMails_SelectedIndexChanged(object sender, EventArgs e)
        {
            UpdateTexts();
        }

        private void UpdateTexts()
        {
            int i = this.ListMails.SelectedIndex;
            if (i < 0)
                return;
            
            this.FullNameTextBox.Text = missingContacts[i].name ?? string.Empty;
            this.FullNameTextBox.RightToLeft = RightToLeft.No;
            this.PhoneNumberTextBox.Text = missingContacts[i].number ?? string.Empty;
            this.checkBoxPakId.Checked = missingContacts[i].pakid;
            this.UnitsList.Text = missingContacts[i].unit.name;
            
            int rankI = FindRankIndex(missingContacts[i].rank);
            if (rankI != -1)
                this.RanksList.SelectedIndex = rankI;
            
            this.LocationText.Text = (i + 1).ToString() + "/" + missingContacts.Count.ToString();
            AdjustComboBoxFontSize();
            ClosePakIdTab();
        }

        private void AdjustComboBoxFontSize()
        {
            int availableWidth = ListMails.Width - 45;
            using (Graphics g = ListMails.CreateGraphics())
            {
                string text = ListMails.Text;
                SizeF textSize = g.MeasureString(text, ListMails.Font);
                if (textSize.Width > availableWidth)
                {
                    float newFontSize = ListMails.Font.Size * availableWidth / textSize.Width;
                    ListMails.Font = new Font(ListMails.Font.FontFamily, newFontSize);
                }
            }
        }

        private int FindUnitIndex(Unit unit)
        {
            for (int i = 0; i < UnitsList.Items.Count; i++)
            {
                Unit temp = UnitsList.Items[i] as Unit;
                if (temp != null && temp.id == unit.id)
                {
                    return i;
                }
            }
            return -1;
        }

        private int FindRankIndex(Rank rank)
        {
            for (int i = 0; i < RanksList.Items.Count; i++)
            {
                Rank temp = RanksList.Items[i] as Rank;
                if (temp != null && temp.code == rank.code)
                {
                    return i;
                }
            }
            return -1;
        }

        private void NextBtn_Click(object sender, EventArgs e)
        {
            this.ListMails.SelectedIndex = (this.ListMails.SelectedIndex + 1) % missingContacts.Count;
            UpdateTexts();
        }

        private void BackBtn_Click(object sender, EventArgs e)
        {
            this.ListMails.SelectedIndex = (this.ListMails.SelectedIndex - 1 + missingContacts.Count) % missingContacts.Count;
            UpdateTexts();
        }

        private bool ValidSelection()
        {
            Contact selected = (Contact)this.ListMails.SelectedItem;
            bool result = true;
            
            if (!selected.SetName(this.FullNameTextBox.Text))
            {
                this.FullNameTextBox.BackColor = Color.LightPink;
                this.FullNameTextBox.TextChanged += BackgroundBack;
                result = false;
            }
            if (!selected.SetNumber(this.PhoneNumberTextBox.Text))
            {
                this.PhoneNumberTextBox.BackColor = Color.LightPink;
                this.PhoneNumberTextBox.TextChanged += BackgroundBack;
                result = false;
            }
            if (selected.unit.name == this.UnitsList.Text)
            {
                this.UnitsList.BackColor = Color.LightPink;
                this.UnitsList.TextChanged += BackgroundBack;
                result = false;
            }
            if (!selected.SetUnit(UnitsList.SelectedItem as Unit))
            {
                this.UnitsList.BackColor = Color.LightPink;
                this.UnitsList.TextChanged += BackgroundBack;
                result = false;
            }
            if (!selected.SetRank(RanksList.SelectedItem as Rank))
            {
                this.RanksList.BackColor = Color.LightPink;
                this.RanksList.TextChanged += BackgroundBack;
                result = false;
            }
            selected.pakid = checkBoxPakId.Checked;
            return result;
        }

        private async void SaveBtn_Click(object sender, EventArgs e)
        {
            Contact selected = (Contact)this.ListMails.SelectedItem;
            if (ValidSelection())
                await AddContactToDBAsync(selected);
            else
                return;
            
            if (missingContacts.Count < 2)
            {
                this.Close();
                return;
            }
            this.ListMails.SelectedIndex = (this.ListMails.SelectedIndex + 1) % missingContacts.Count;
            missingContacts.Remove(selected);
            UpdateTexts();
        }

        private async Task<bool> AddContactToDBAsync(Contact contact)
        {
            /*
            Task<List<Contact>> task = 
                task.RunSynchronously();
                task.Wait();*/
            List<Contact> similars = await database.GetSimilarContactsAsync(contact);
            if (similars.Count > 0)
            {
                ViewContacts SimForm = new ViewContacts(similars);
                SimForm.Enabled = true;
                SimForm.ShowDialog();
                if (SimForm.DialogResult == DialogResult.Yes)
                {
                    return await database.AddMail1Async(similars[SimForm.index].id, contact.addresses[0]);
                }
            }
            return await database.AddContactAsync(contact);
        }

        private void BackgroundBack(object sender, EventArgs e)
        {
            if (sender == this.PhoneNumberTextBox)
            {
                this.PhoneNumberTextBox.BackColor = Color.GhostWhite;
                this.PhoneNumberTextBox.TextChanged -= BackgroundBack;
            }
            if (sender == this.FullNameTextBox)
            {
                this.FullNameTextBox.BackColor = Color.GhostWhite;
                this.FullNameTextBox.TextChanged -= BackgroundBack;
            }
            if (sender == this.UnitsList)
            {
                this.UnitsList.BackColor = Color.GhostWhite;
                this.UnitsList.TextChanged -= BackgroundBack;
            }
        }

        private async void IgnoreContact_ClickAsync(object sender, EventArgs e)
        {
            Contact selected = (Contact)this.ListMails.SelectedItem;
            if (this.IgnoreForever.CheckState == CheckState.Checked)
                await database.AddIgnoredAddressAsync(selected.addresses[0]);
            
            if (missingContacts.Count <= 1)
            {
                this.Close();
                return;
            }
            this.ListMails.SelectedIndex = (this.ListMails.SelectedIndex + 1) % missingContacts.Count;
            missingContacts.Remove(selected);
            UpdateTexts();
        }

        private async void loadContactsBW_DoWork(object sender, DoWorkEventArgs e)
        {
            FullNameTextBox.Enabled = false;
            PhoneNumberTextBox.Enabled = false;
            UnitsList.Enabled = false;
            IgnoreContact.Enabled = false;
            IgnoreContact.Enabled = false;
            SaveBtn.Enabled = false;
            NextBtn.Enabled = false;
            BackBtn.Enabled = false;
            Debug.WriteLine("is working?");
            var task1 = database.GetGULContactsFromAddressesAsync(addresses);
            task1.Wait();
            this.missingContacts = task1.Result; //RETRIEVES DATA FROM GUL WICH TAKES TIME, returns contact with address only if not found
            
            if (this.missingContacts.Count > 0)
            {
                var task2 = database.GetAllUnitsAsync();
                task2.Wait();
                Unit[] units = task2.Result;
                UnitsList.DataSource = units;
            }
        }

        private void CompletedHelper()
        {
            FullNameTextBox.Visible = true;
            PhoneNumberTextBox.Visible = true;
            UnitsList.Visible = true;
            IgnoreContact.Visible = true;
            IgnoreContact.Visible = true;
            SaveBtn.Visible = true;
            NextBtn.Visible = true;
            BackBtn.Visible = true;
            IgnoreForever.Visible = true;
            label4.Visible = true;
            label2.Visible = true;
            ListMails.Visible = true;
            LocationText.Visible = true;
            this.LoadingAnimationLabel.Visible = false;
            FullNameTextBox.Enabled = true;
            PhoneNumberTextBox.Enabled = true;
            UnitsList.Enabled = true;
            IgnoreContact.Enabled = true;
            IgnoreContact.Enabled = true;
            SaveBtn.Enabled = true;
            NextBtn.Enabled = true;
            BackBtn.Enabled = true;
            RanksList.Enabled = true;
            label6.Enabled = true;
            RanksList.Visible = true;
            label6.Visible = true;
            AddDisplayNames();
            InitializeBinding();
            UpdateTexts();
        }

        private void AddDisplayNames()
        {
            foreach (Contact contact in this.missingContacts)
            {
                string name = "לא זוהה";
                addressesToDisplayName.TryGetValue(contact.addresses[0], out name);
                contact.displayName = name;
            }
        }

        private void loadContactsBW_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            if (this.missingContacts.Count == 0)
            {
                this.Close();
                return;
            }
            //this.progressBar.Value = 100;
            CompletedHelper();
        }

        private async void SavePakIdBtn_Click(object sender, EventArgs e)
        {
            int result;
            Contact pakid = ValidPakIdSelection(out result);
            if (result == 0)
                return;
            if (result == 1)
            {
                List<Contact> similars = await database.GetSimilarContactsAsync(pakid);
                if (similars.Count > 0)
                {
                    ViewContacts SimForm = new ViewContacts(similars);
                    SimForm.Enabled = true;
                    SimForm.ShowDialog();
                    if (SimForm.DialogResult == DialogResult.Yes)
                    {
                        pakid = similars[SimForm.index];
                    }
                }

                Contact selected = (Contact)ListMails.SelectedItem;
                List<Contact> newList = new List<Contact>();
                newList = new List<Contact>();
                if (selected.pkididm != null)
                    newList = selected.pkididm.ToList();
                newList.Add(pakid);
                selected.pkididm = newList.ToArray();
            }

            ClosePakIdTab();
        }

        private Contact ValidPakIdSelection(out int result)
        {
            result = 1;
            Contact newPakid = new Contact();
            if (!newPakid.SetName(this.PakIdNameTB.Text))
            {
                this.FullNameTextBox.BackColor = Color.LightPink;
                this.FullNameTextBox.TextChanged += BackgroundBack;
                result = 0;
            }
            if (!newPakid.SetNumber(this.PakIdPhoneNumTB.Text))
            {
                this.PhoneNumberTextBox.BackColor = Color.LightPink;
                this.PhoneNumberTextBox.TextChanged += BackgroundBack;
                result = 0;
            }
            if (!newPakid.SetUnit(UnitsList.SelectedItem as Unit))
            {
                this.UnitsList.BackColor = Color.LightPink;
                this.UnitsList.TextChanged += BackgroundBack;
                result = 0;
            }
            newPakid.SetRank(new Rank(0, "חייל"));
            newPakid.pakid = true;
            return newPakid;
        }

        private void AddPakIdBtn_Click(object sender, EventArgs e)
        {
            OpenPakIdTab();
        }

        private void OpenPakIdTab()
        {
            AddPakIdBtn.Hide();
            this.PakIdNameTB.Show();
            this.SavePakIdBtn.Show();
            this.PakIdPhoneNumTB.Show();
            this.Size = new Size(484, 433);
        }

        private void ClosePakIdTab()
        {
            AddPakIdBtn.Show();
            this.PakIdNameTB.Hide();
            this.PakIdNameTB.ResetText();
            this.SavePakIdBtn.Hide();
            this.PakIdPhoneNumTB.Hide();
            this.PakIdPhoneNumTB.ResetText();
            this.Size = new Size(484, 370);
        }
    }
}

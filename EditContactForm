using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using static System.Windows.Forms.ListBox;
using static System.Windows.Forms.ListView;

namespace LishcaAddIn
{
    public partial class EditContactForm : Form
    {
        protected static DatabaseManager database = DatabaseManager.Instance;
        //private Contact SelectedContact;
        public EditContactForm(Unit[] units)
        {
            InitializeComponent();
            PopulateUnitsList(units);
            UnitsListView.SelectedIndexChanged += 
                UnitsListView_SelectedIndexChanged;
            PakidPanel.Hide();
        }

        public static async Task<EditContactForm> CreateAsync()
        {
            Unit[] units = await database.GetAllUnitsAsync();
            return new EditContactForm(units);
        }

        private void PopulateUnitsList(Unit[] units)
        {
            UnitsListView.Sorted = true;
            UnitsListView.DataSource = units;
            UnitsListView.DisplayMember = "name";
        }

        private async void UnitsListView_SelectedIndexChanged(object sender, EventArgs e)
        {
            Contact[] data = null;
            Unit unit = null;
            SelectedObjectCollection selected = UnitsListView.SelectedItems;
            if (selected != null)
                if (selected.Count > 0)
                    unit = selected[0] as Unit;

            if(unit != null)
                data = await database.GetContactsFromUnit(unit);
            if (data != null)
            {
                PopulateContactsList(data);
            }
            //else*/
            //Show error
        }

        private void PopulateContactsList(Contact[] contacts)
        {
            ContactsListView.Sorted = true;
            ContactsListView.DataSource = contacts;
            ContactsListView.DisplayMember = "name";
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            int foundIndex = UnitsListView.FindString(textBox1.Text);
            if (foundIndex > -1)
            {
                var found = UnitsListView.Items[foundIndex];
                if (found != null)
                    UnitsListView.SelectedItem = found;
            }
        }

        private void textBox2_TextChanged(object sender, EventArgs e)
        {
            int foundIndex = ContactsListView.FindString(textBox2.Text);
            if (foundIndex > -1)
            {
                var found = ContactsListView.Items[foundIndex];
                if (found != null)
                    ContactsListView.SelectedItem = found;
            }
        }

        private async void ContactsListView_SelectedIndexChanged(object sender, EventArgs e)
        {
            SelectedObjectCollection selected = 
                ContactsListView.SelectedItems;
            if (selected != null)
                if (selected.Count > 0)
                {
                    ContactView cv = new ContactView(selected[0] as Contact)
                    {
                        Dock = DockStyle.Fill
                    };
                    this.ContactPanel.Controls.Clear();
                    this.ContactPanel.Controls.Add(cv);
                    tableLayoutPanel3.Enabled = true;
                    PakidPanel.Hide();
                    await cv.EnterEditMode();
                }
        }

        private async void DeleteBtn_Click(object sender, EventArgs e)
        {
            foreach (Control c in ContactPanel.Controls)
            {
                if (c is ContactView)
                {
                    if (await database.DeleteContact((c as 
                        ContactView).contact))
                    {
                        Contact[] contacts = ContactsListView.DataSource as 
                            Contact[];
                        contacts.ToList().RemoveAll( con => con.id == (c as 
                            ContactView).contact.id);
                        ContactsListView.DataSource = contacts.ToArray();
                        tableLayoutPanel3.Enabled = false;
                        //ContactPanel.Enabled = false;
                        //DisableButtons();
                    }
                }
            }
        }

        private void DisableButtons()
        {
            DeleteBtn.Enabled = false;
            SaveBtn.Enabled = false;
            AddPakid.Enabled = false;
        }

        private void EnableButtons()
        {
            DeleteBtn.Enabled = true;
            SaveBtn.Enabled = true;
            AddPakid.Enabled = true;
        }

        private async void SearchNameBtn_Click(object sender, EventArgs e)
        {
            string searchText = SearchNameTextBox.Text;
            if (!(string.IsNullOrEmpty(searchText) || 
                string.IsNullOrWhiteSpace(searchText)))
            {
                Contact[] data = await database.GetContactFromName
                    (searchText);
                PopulateContactsList(data);
            }
        }

        private async void SearchNumberBtn_Click(object sender, EventArgs e)
        {
            string searchText = SearchNumberTextBox.Text;
            if (!(string.IsNullOrEmpty(searchText) || 
                string.IsNullOrWhiteSpace(searchText) || searchText.Length != 
                9))
            {
                Contact[] data = await database.GetContactFromNumberAsync
                    (searchText);
                PopulateContactsList(data);
            }
        }

        private void AddPakid_Click(object sender, EventArgs e)
        {
            PakidPanel.Show();
            PakidNameTB.Show();
            PakidNameTB.Text = string.Empty;
            PakidPhoneNumTB.Show();
            PakidPhoneNumTB.Text = string.Empty;
            SavePakidBtn.Show();
            AddPakid.Enabled = false;
        }

        private async void SavePakidBtn_ClickAsync(object sender, EventArgs e)
        {
            AddPakid.Enabled = true;
            int result;
            Contact mefaked = null, pakid;
            foreach (Control c in ContactPanel.Controls)
            {
                if (c is ContactView)
                {
                    mefaked = (c as ContactView).contact;
                }
            }
            if (mefaked is null)
                return;
            else
                pakid = ValidPakidSelection(out result, mefaked);

            if (result == 0)
                return;
            if (result == 1)
            {
                List<Contact> similars = await 
                    database.GetSimilarContactsAsync(pakid);
                if (similars.Count > 0)
                {
                    ViewContacts SimForm = new ViewContacts(similars);
                    SimForm.Enabled = true;
                    SimForm.ShowDialog();
                    if (SimForm.DialogResult == DialogResult.Yes)
                    {
                        pakid = similars[SimForm.index];
                        if(await database.AddPakidToExistingContact(pakid.id, 
                            mefaked.id))
                            PakidPanel.Hide();
                        return;
                    }
                }
            }

            if (await database.AddPakidToExistingContact(pakid, 
                mefaked.id))
                PakidPanel.Hide();
        }

        // EditContactForm.cs - Partial class methods from lines 210-273

private Contact ValidPakidSelection(out int result, Contact meFaked)
{
    result = 1;
    Contact newPakid = new Contact();
    if (!newPakid.SetName(this.PakidNameTB.Text))
    {
        this.PakidNameTB.BackColor = Color.LightPink;
        this.PakidNameTB.TextChanged += BackgroundBack;
        result = 0;
    }
    if (!newPakid.SetNumber(this.PakidPhoneNumTB.Text))
    {
        this.PakidPhoneNumTB.BackColor = Color.LightPink;
        this.PakidPhoneNumTB.TextChanged += BackgroundBack;
        result = 0;
    }
    if (!newPakid.SetUnit(meFaked.unit))
    {
        result = 0;
    }
    newPakid.SetRank(new Rank(0, "חייל"));
    newPakid.pakid = true;
    return newPakid;
}

private void BackgroundBack(object sender, EventArgs e)
{
    if (sender == this.PakidPhoneNumTB)
    {
        this.PakidPhoneNumTB.BackColor = Color.GhostWhite;
        this.PakidPhoneNumTB.TextChanged -= BackgroundBack;
    }
    if (sender == this.PakidNameTB)
    {
        this.PakidNameTB.BackColor = Color.GhostWhite;
        this.PakidNameTB.TextChanged -= BackgroundBack;
    }
}

private async void SaveBtn_Click(object sender, EventArgs e)
{
    SaveBtn.Enabled = false;
    ContactView cv = null;
    foreach (Control c in ContactPanel.Controls)
    {
        if (c is ContactView)
        {
            cv = c as ContactView;
        }
    }
    
    if (cv is null)
        return;
        
    if (cv.ValidSelection())
        await cv.ChangeContactToDBAsync();
    SaveBtn.Enabled = true;
}

using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace LishcaAddIn
{
    public partial class EditUnitForm : Form
    {
        Unit MainUnit; // represents current unit displayed 
        BindingList<Unit> ChildUnits = new BindingList<Unit>();
        Unit MotherUnit;
        protected static DatabaseManager database = DatabaseManager.Instance;

        public EditUnitForm()
        {
            InitializeComponent();
            this.backgroundWorkerLoadUnits.RunWorkerAsync();
        }

        public EditUnitForm(Unit unit)
        {
            InitializeComponent();
            this.MainUnit = unit;
            PopulateListsAsync().RunSynchronously();
            UpdateTexts();
        }

        private void UpdateTexts()
        {
            this.UnitNameBtn.Text = MainUnit.name;
            this.UnitNameTextBox.Text = MainUnit.name;
            int unitI = FindUnitIndex();
            if (unitI != -1)
            {
                this.MotherUnitsList.SelectedIndex = unitI;
                this.NUSapSymbol.SelectedIndex = unitI;
                this.MotherUnit = this.MotherUnitsList.Items[unitI] as Unit;
                this.MotherUnitBtn.Text = MotherUnit.name;
            }

            int codeI = FindCodeIndex();
            if (codeI != -1)
                this.AllUnitCodes.SelectedIndex = codeI;

            this.ChildUnitsLB.DataSource = ChildUnits;
        }

        private int FindCodeIndex()
        {
            for (int i = 0; i < AllUnitCodes.Items.Count; i++)
            {
                if ((string)AllUnitCodes.Items[i] == MainUnit.SAPSymbol)
                {
                    return i;
                }
            }
            return -1;
        }

        private int FindUnitIndex()
        {
            for (int i = 0; i < MotherUnitsList.Items.Count; i++)
            {
                Unit temp = MotherUnitsList.Items[i] as Unit;
                if (temp != null && temp.id == MainUnit.motherUnit)
                {
                    return i;
                }
            }
            return -1;
        }

        private Unit FindUnit(int id)
        {
            for (int i = 0; i < MotherUnitsList.Items.Count; i++)
            {
                Unit temp = MotherUnitsList.Items[i] as Unit;
                if (temp != null && temp.id == id)
                {
                    return temp;
                }
            }
            return null;
        }

        private async Task PopulateListsAsync()
        {
            Unit[] allUnits = await database.GetAllUnitsAsync();
            MotherUnitsList.DataSource = allUnits;
            //MotherUnitsList.SelectedIndex = 1;
            NUSapSymbol.SelectedIndex = 1;

            string[] symbols = await database.GetAllSAPSymbolsAsync();
            AllUnitCodes.DataSource = symbols;
            NUSapSymbol.DataSource = symbols;

            UpdateChildUnits(allUnits);
        }

        private void UpdateChildUnits(Unit[] allUnits)
        {
            ChildUnits.Clear();
            for (int i = 0; i < allUnits.Length; i++)
            {
                if (allUnits[i].motherUnit == MainUnit.id)
                    ChildUnits.Add(allUnits[i]);
            }
        }

        private async void AddUnitBtn_Click(object sender, EventArgs e)
        {
            Unit newChildUnit = new Unit();
            if (newChildUnit.SetName(NUNameTextBox.Text))
            {
                newChildUnit.SAPSymbol = string.Empty; // or NUSapSymbol.Text;
                newChildUnit.motherUnit = MainUnit.id;
                newChildUnit = await database.AddNewUnitAsync(newChildUnit);
                if (newChildUnit != null)
                {
                    if (newChildUnit.name != null)
                        ChildUnits.Add(newChildUnit);
                }
            }
            else
            {
                this.NUNameTextBox.BackColor = Color.LightPink;
                this.NUNameTextBox.TextChanged += BackroundBack;
            }
        }

        private async void UpdateUnitBtn_Click(object sender, EventArgs e)
        {
            if (this.MainUnit.SetName(UnitNameTextBox.Text))
            {
                this.UnitNameTextBox.BackColor = Color.LightGreen;
                this.UnitNameTextBox.TextChanged += BackroundBack;
            }
            else
            {
                this.UnitNameTextBox.BackColor = Color.LightPink;
                this.UnitNameTextBox.TextChanged += BackroundBack;
            }

            this.MainUnit.motherUnit = (MotherUnitsList.SelectedItem as Unit).id;
            this.MainUnit.SAPSymbol = this.AllUnitCodes.Text;
            await database.UpdateUnitAsync(MainUnit);
            UpdateTexts();
        }

        private void BackroundBack(object sender, EventArgs e)
        {
            if (sender == this.NUNameTextBox)
            {
                this.NUNameTextBox.BackColor = Color.GhostWhite;
                this.NUNameTextBox.TextChanged -= BackroundBack;
            }
        }

        private void ModeSwitchBtn_Click(object sender, EventArgs e)
        {
            if (ModeSwitchBtn.Text == "שינוי")
            {
                ModeSwitchBtn.Text = "צפיה";
            }
            else
            {
                ModeSwitchBtn.Text = "שינוי";
            }
        }

        private async void DeleteUnitBtn_Click(object sender, EventArgs e)
        {
            if (this.MainUnit.id == 6 || this.MainUnit.id == 5 || this.MainUnit.id == 4) 
            {
                FreeText errForm = new FreeText("איסור מחיקת יח", "מחיקת היח", "אסור למחוק את היחידה");
                errForm.Enabled = true;
                errForm.ShowDialog();
                return;
            }

            if (ChildUnits.Count > 0)
            {
                FreeText errForm = new FreeText("שגיאת מחיקה", "יחידה עם יחידות משנה", "יחידה עם יחידות משנה לא ניתנת למחיקה");
                errForm.Enabled = true;
                errForm.ShowDialog();
                return;
            }

            await database.RemoveChildlessUnitAsync(this.MainUnit.id);
            this.MainUnit = await database.GetMainUnitAsync();
        }

        private async void RefreshBtn_ClickAsync(object sender, EventArgs e)
        {
            this.MainUnit = await database.GetUnitAsync(this.MainUnit.id);
            await PopulateListsAsync();
            UpdateTexts();
        }

        private async void MotherUnitBtn_Click(object sender, EventArgs e)
        {
            this.MainUnit = await database.GetUnitAsync(this.MainUnit.motherUnit);
            await PopulateListsAsync();
            UpdateTexts();
        }

        private async void backgroundWorkerLoadUnits_DoWork(object sender, DoWorkEventArgs e)
        {
            var task1 = database.GetMainUnitAsync();
            task1.Wait();
            this.MainUnit = task1.Result;
            PopulateListsAsync().Wait();
        }

        private void backgroundWorkerLoadUnits_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            UpdateTexts();
            NUNameTextBox.Show();
            NUSapSymbol.Show();
            AddUnitBtn.Show();
            UnitNameTextBox.Show();
            DeleteUnitBtn.Show();
            UpdateUnitBtn.Show();
            MotherUnitsList.Show();
            AllUnitCodes.Show();
        }

        private async void ChildUnitsLB_DoubleClick(object sender, EventArgs e)
        {
            if (ChildUnitsLB.SelectedItems.Count > 0)
            {
                Unit selected = (Unit)ChildUnitsLB.SelectedItems[0];
                this.MainUnit = selected;
                await PopulateListsAsync();
                UpdateTexts();
            }
        }
    }
}

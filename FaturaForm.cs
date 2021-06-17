using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Kale_Mobilya
{
    public partial class FaturaForm : Form
    {
        public FaturaForm()
        {
            InitializeComponent();

            //Yılları Ekle ve Varsayılanı Seç
            int foryear = 2019;
            for (int i = 0; i < (DateTime.Now.Year - 2019); i++)
            {
                foryear++;
                CmbYil1.Items.Add(foryear);
            }
            CmbYil1.SelectedIndex = DateTime.Now.Year - 2019;

            //Aylardan Varsayılanı Seç
            int ay = DateTime.Now.Month;
            if (ay == 1)
            { CmbAy1.SelectedIndex = 0; }
            else if (ay == 2)
            { CmbAy1.SelectedIndex = 1; }
            else if (ay == 3)
            { CmbAy1.SelectedIndex = 2; }
            else if (ay == 4)
            { CmbAy1.SelectedIndex = 3; }
            else if (ay == 5)
            { CmbAy1.SelectedIndex = 4; }
            else if (ay == 6)
            { CmbAy1.SelectedIndex = 5; }
            else if (ay == 7)
            { CmbAy1.SelectedIndex = 6; }
            else if (ay == 8)
            { CmbAy1.SelectedIndex = 7; }
            else if (ay == 9)
            { CmbAy1.SelectedIndex = 8; }
            else if (ay == 10)
            { CmbAy1.SelectedIndex = 9; }
            else if (ay == 11)
            { CmbAy1.SelectedIndex = 10; }
            else if (ay == 12)
            { CmbAy1.SelectedIndex = 11; }
        }

        private void BtnFatura1_Click(object sender, EventArgs e)
        {
            this.DialogResult = DialogResult.Abort;
            this.Close();
        }

        private void BtnOk1_Click(object sender, EventArgs e)
        {
            int ays = DateTime.Now.Month;
            int yils = DateTime.Now.Year;
            int ay2 = 0;
            if (CmbAy1.SelectedIndex == 0)
            {
                ay2 = 1;
            }
            else if (CmbAy1.SelectedIndex == 1)
            {
                ay2 = 2;
            }
            else if (CmbAy1.SelectedIndex == 2)
            {
                ay2 = 3;
            }
            else if (CmbAy1.SelectedIndex == 3)
            {
                ay2 = 4;
            }
            else if (CmbAy1.SelectedIndex == 4)
            {
                ay2 = 5;
            }
            else if (CmbAy1.SelectedIndex == 5)
            {
                ay2 = 6;
            }
            else if (CmbAy1.SelectedIndex == 6)
            {
                ay2 = 7;
            }
            else if (CmbAy1.SelectedIndex == 7)
            {
                ay2 = 8;
            }
            else if (CmbAy1.SelectedIndex == 8)
            {
                ay2 = 9;
            }
            else if (CmbAy1.SelectedIndex == 9)
            {
                ay2 = 10;
            }
            else if (CmbAy1.SelectedIndex == 10)
            {
                ay2 = 11;
            }
            else if (CmbAy1.SelectedIndex == 11)
            {
                ay2 = 12;
            }

            Form1.FTarihYil = Convert.ToInt32(CmbYil1.SelectedItem);
            Form1.FTarihAy = ay2;

            if ((Convert.ToInt32(CmbYil1.SelectedItem) == yils && ay2 <= ays) || (Convert.ToInt32(CmbYil1.SelectedItem) < yils))
            {
                this.DialogResult = DialogResult.OK;
            }
            else
            {
                MessageBox.Show("Gelecek bir ay seçtiniz!", "Uyarı", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }           
        }

        private void BtnFaturaListe_Click(object sender, EventArgs e)
        {
            this.DialogResult = DialogResult.Abort;
            this.Close();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            this.DialogResult = DialogResult.Cancel;
            this.Close();
        }
    }
}

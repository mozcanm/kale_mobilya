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
    public partial class CizelgeForm : Form
    {
        public CizelgeForm()
        {
            InitializeComponent();

            //Yılları Ekle ve Varsayılanı Seç
            int foryear = 2019;
            for (int i = 0; i < (DateTime.Now.Year - 2019 + 1); i++)
            {
                foryear++;
                CmbCizYil1.Items.Add(foryear);
            }
            CmbCizYil1.SelectedIndex = DateTime.Now.Year - 2019;

            //Aylardan Varsayılanı Seç
            int ay = DateTime.Now.Month;
            if (ay == 1)
            { CmbCizAy1.SelectedIndex = 0; }
            else if (ay == 2)
            { CmbCizAy1.SelectedIndex = 1; }
            else if (ay == 3)
            { CmbCizAy1.SelectedIndex = 2; }
            else if (ay == 4)
            { CmbCizAy1.SelectedIndex = 3; }
            else if (ay == 5)
            { CmbCizAy1.SelectedIndex = 4; }
            else if (ay == 6)
            { CmbCizAy1.SelectedIndex = 5; }
            else if (ay == 7)
            { CmbCizAy1.SelectedIndex = 6; }
            else if (ay == 8)
            { CmbCizAy1.SelectedIndex = 7; }
            else if (ay == 9)
            { CmbCizAy1.SelectedIndex = 8; }
            else if (ay == 10)
            { CmbCizAy1.SelectedIndex = 9; }
            else if (ay == 11)
            { CmbCizAy1.SelectedIndex = 10; }
            else if (ay == 12)
            { CmbCizAy1.SelectedIndex = 11; }
        }

        private void BtnCizOk1_Click(object sender, EventArgs e)
        {
            int ays = DateTime.Now.Month;
            int yils = DateTime.Now.Year;
            int ay2 = 0;

            if (CmbCizAy1.SelectedIndex == 0)
            {
                ay2 = 1;
            }
            else if (CmbCizAy1.SelectedIndex == 1)
            {
                ay2 = 2;
            }
            else if (CmbCizAy1.SelectedIndex == 2)
            {
                ay2 = 3;
            }
            else if (CmbCizAy1.SelectedIndex == 3)
            {
                ay2 = 4;
            }
            else if (CmbCizAy1.SelectedIndex == 4)
            {
                ay2 = 5;
            }
            else if (CmbCizAy1.SelectedIndex == 5)
            {
                ay2 = 6;
            }
            else if (CmbCizAy1.SelectedIndex == 6)
            {
                ay2 = 7;
            }
            else if (CmbCizAy1.SelectedIndex == 7)
            {
                ay2 = 8;
            }
            else if (CmbCizAy1.SelectedIndex == 8)
            {
                ay2 = 9;
            }
            else if (CmbCizAy1.SelectedIndex == 9)
            {
                ay2 = 10;
            }
            else if (CmbCizAy1.SelectedIndex == 10)
            {
                ay2 = 11;
            }
            else if (CmbCizAy1.SelectedIndex == 11)
            {
                ay2 = 12;
            }

            Form1.CTarihYil = Convert.ToInt32(CmbCizYil1.SelectedItem);
            Form1.CTarihAy = ay2;
            this.DialogResult = DialogResult.OK;
        }
    }
}

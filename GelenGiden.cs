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
    public partial class GelenGiden : Form
    {
        public GelenGiden()
        {
            InitializeComponent();
        }

        private void BtnTarih2_Click(object sender, EventArgs e)
        {
            this.DialogResult = DialogResult.Cancel;
            this.Close();
        }

        private void BtnTarih1_Click(object sender, EventArgs e)
        {
            //var tarih1 = DtpTarih1.Value;
            //var tarih2 = DtpTarih2.Value;
            Form1.Ttarih1 = DtpTarih1.Value;
            Form1.Ttarih2 = DtpTarih2.Value;
            if (Form1.Ttarih1 <= Form1.Ttarih2)
            {
                this.DialogResult = DialogResult.OK;
            }
            else
            {
                MessageBox.Show("'Tarih 1', 'Tarih 2' de küçük olmalı!", "Uyarı", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }  
        }
    }
}

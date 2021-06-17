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
    public partial class DetayForm : Form
    {
        KaleMobilyaDataContext ctx = new KaleMobilyaDataContext();

        public DetayForm()
        {
            InitializeComponent();

            lblID.Text = Form1.KisiIdDetay.ToString();
            lblID.Tag = Form1.KisiIdDetay;

            if (Form1.KisiAdDetay != null)
            {
                lblAd.Text = Form1.KisiAdDetay.ToString();
            }
            else
            {
                label1.Visible = lblAd.Visible = false;
            }

            if (Form1.KisiFirmaDetay != null)
            {
                lblFirma.Text = Form1.KisiFirmaDetay.ToString();
            }
            else
            {
                label2.Visible = lblFirma.Visible = false;
            }

            if (Form1.KisiTel2Detay!= null)
            {
                lblTel2.Text = Form1.KisiTel2Detay.ToString();
            }
            else
            {
                lblTel2B.Visible = lblTel2.Visible = false;
            }

            if (Form1.KisiTel1Detay != null)
            {
                lblTel1.Text = Form1.KisiTel1Detay.ToString();
            }
            else
            {
                label4.Visible = lblTel1.Visible = false;
            }

            if (Form1.KisiAdresDetay!= null)
            {
                lblAdres.Text = Form1.KisiAdresDetay.ToString();
            }
            else
            {
                lblAdresB.Visible = lblAdres.Visible = false;
            }

            var liste = from kisiler in ctx.Kisis
                        join cariler in ctx.Caris on kisiler.KisiID equals cariler.KisiID
                        join durumlar in ctx.Durums on cariler.DurumID equals durumlar.DurumID
                        orderby cariler.Tarih descending
                        select new
                        {
                            kisiler.KisiID,
                            kisiler.Ad,
                            Durum = durumlar.Durumlar,
                            cariler.Tutar,
                            cariler.Tarih,
                            Açıklama = cariler.Aciklama,
                            durumlar.DurumID,
                            cariler.CariID
                        };

            dataGridView2.DataSource = liste.Where(x => x.KisiID == Convert.ToInt32(lblID.Tag));
            dataGridView2.Columns["KisiID"].Visible = dataGridView2.Columns["Ad"].Visible = dataGridView2.Columns["DurumID"].Visible = dataGridView2.Columns["CariID"].Visible = false;

            dataGridView2.Columns[3].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            dataGridView2.Columns[3].HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleRight;
            dataGridView2.Columns[4].HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dataGridView2.Columns[4].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;

            //Hesaplama
            var hesap = ctx.Kisis.Join(ctx.Caris,
                kisiler => kisiler.KisiID,
                cariler => cariler.KisiID,
                (ki, ca) => new
                {
                    kKisiID = ki.KisiID,
                    cKisiID = ca.KisiID,
                    ca.Tutar,
                    ca.DurumID,
                    ca.Tarih
                }).Select(x => new
                {
                    x.cKisiID,
                    x.kKisiID,
                    x.DurumID,
                    x.Tarih,
                    x.Tutar
                });

            //Alacak
            decimal? alacak = hesap.Where(x => x.cKisiID == Convert.ToInt32(lblID.Tag)).Where(x => x.DurumID == 1).Sum(x => x.Tutar);
            decimal? alindi = hesap.Where(x => x.cKisiID == Convert.ToInt32(lblID.Tag)).Where(x => x.DurumID == 2).Sum(x => x.Tutar);
            decimal? iskonto = hesap.Where(x => x.cKisiID == Convert.ToInt32(lblID.Tag)).Where(x => x.DurumID == 9).Sum(x => x.Tutar);

            if (alacak != null)
            {
                lblAlacak.Text = String.Format("{0:N}\n", alacak);
            }
            else
            {
                lblAlacak.Text = String.Format("{0:N}\n", 0);
            }
            //
            if (alindi != null)
            {
                LblAlindi.Text = String.Format("{0:N}\n", alindi);
            }
            else
            {
                LblAlindi.Text = String.Format("{0:N}\n", 0);
            }
            //
            if (iskonto != null)
            {
                LblIskonto.Text = String.Format("{0:N}\n", iskonto);
            }
            else
            {
                LblIskonto.Text = String.Format("{0:N}\n", 0);
            }
            //
            if (alacak != null && alindi != null && iskonto != null)
            {
                LblKalanAlacak.Text = String.Format("{0:N}\n", (alacak - (alindi + iskonto)));
            }
            else if (alacak != null && alindi == null && iskonto == null)
            {
                LblKalanAlacak.Text = String.Format("{0:N}\n", alacak);
            }
            else if (alacak != null && alindi != null && iskonto == null)
            {
                LblKalanAlacak.Text = String.Format("{0:N}\n", (alacak - alindi));
            }
            else if (alacak != null && alindi == null && iskonto != null)
            {
                LblKalanAlacak.Text = String.Format("{0:N}\n", (alacak - iskonto));
            }
            else
            {
                LblKalanAlacak.Text = String.Format("{0:N}\n", 0);
            }

            //Kalan Alacak Renk Değişimi
            decimal alacak2 = 0;
            decimal alindi2 = 0;
            decimal iskonto2 = 0;

            if (alacak == null || alacak == 0)
            {
                alacak2 = 0;
            }
            else
            {
                alacak2 = (decimal)alacak;
            }

            if (alindi == null || alindi == 0)
            {
                alindi2 = 0;
            }
            else
            {
                alindi2 = (decimal)alindi;
            }

            if (iskonto == null || iskonto == 0)
            {
                iskonto2 = 0;
            }
            else
            {
                iskonto2 = (decimal)iskonto;
            }

            if ((alacak2 - (alindi2 + iskonto2)) == 0)
            {
                LblKalanAlacak.BackColor = Color.ForestGreen;
                label50.BackColor = Color.ForestGreen;
                label48.BackColor = Color.ForestGreen;
                label3.BackColor = Color.ForestGreen;
            }
            else
            {
                LblKalanAlacak.BackColor = Color.Maroon;
                label50.BackColor = Color.Maroon;
                label48.BackColor = Color.Maroon;
                label3.BackColor = Color.Maroon;
            }

            //Borç
            decimal? borc = hesap.Where(x => x.cKisiID == Convert.ToInt32(lblID.Tag)).Where(x => x.DurumID == 3).Sum(x => x.Tutar);
            decimal? odendi = hesap.Where(x => x.cKisiID == Convert.ToInt32(lblID.Tag)).Where(x => x.DurumID == 4).Sum(x => x.Tutar);

            if (borc != null)
            {
                lblBorc.Text = String.Format("{0:N}\n", borc);
            }
            else
            {
                lblBorc.Text = String.Format("{0:N}\n", 0);
            }
            //
            if (odendi != null)
            {
                LblOdendi.Text = String.Format("{0:N}\n", odendi);
            }
            else
            {
                LblOdendi.Text = String.Format("{0:N}\n", 0);
            }
            //
            if (borc != null && odendi != null)
            {
                LblKalanBorc.Text = String.Format("{0:N}\n", (borc - odendi));
            }
            else if (borc != null && odendi == null)
            {
                LblKalanBorc.Text = String.Format("{0:N}\n", borc);
            }
            else
            {
                LblKalanBorc.Text = String.Format("{0:N}\n", 0);
            }
        }


        private void dataGridView2_DataBindingComplete(object sender, DataGridViewBindingCompleteEventArgs e)
        {
            foreach (DataGridViewRow dGVRow in this.dataGridView2.Rows)
            {
                dGVRow.HeaderCell.Value = String.Format("{0}", dGVRow.Index + 1);
            }
            //Genişlik Ayarı
            this.dataGridView2.AutoResizeRowHeadersWidth(DataGridViewRowHeadersWidthSizeMode.AutoSizeToAllHeaders);
        }

        private void txtAraCari_TextChanged(object sender, EventArgs e)
        {
            var sonuc = from kisiler in ctx.Kisis
                        join cariler in ctx.Caris on kisiler.KisiID equals cariler.KisiID
                        join durumlar in ctx.Durums on cariler.DurumID equals durumlar.DurumID
                        orderby cariler.Tarih descending
                        select new
                        {
                            kisiler.KisiID,
                            kisiler.Ad,
                            Durum = durumlar.Durumlar,
                            cariler.Tutar,
                            cariler.Tarih,
                            Açıklama = cariler.Aciklama,
                            durumlar.DurumID,
                            cariler.CariID
                        };
            dataGridView2.DataSource = sonuc.Where(x => x.KisiID == Convert.ToInt32(lblID.Tag) && (x.Açıklama.Contains(txtAraCari.Text) || x.Durum.Contains(txtAraCari.Text) || Convert.ToString(x.Tutar).Contains(txtAraCari.Text)));
            dataGridView2.Columns["KisiID"].Visible = dataGridView2.Columns["Ad"].Visible = dataGridView2.Columns["DurumID"].Visible = dataGridView2.Columns["CariID"].Visible = false;
        }
    }
}

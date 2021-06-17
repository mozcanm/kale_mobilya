using System;
//using System.Collections.Generic;
//using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using ClosedXML.Excel;
//using System.Text;
//using System.Threading.Tasks;
using System.Windows.Forms;
using System.Globalization;
using System.Text.RegularExpressions;

namespace Kale_Mobilya
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
            dataGridView4.CellPainting += new DataGridViewCellPaintingEventHandler(DataGridView4_CellPainting);
            tabControl1.TabPages.Remove(tabPageHesaplama);
            //btnEEkle.Visible = btnESil.Visible = btnGGuncelle.Visible = btnGuncelle.Visible = btnEkle.Visible = btnSil.Visible = false;
        }

        KaleMobilyaDataContext ctx = new KaleMobilyaDataContext();
        bool anahtar = true;
        bool anahtar2 = false;
        public static DateTime Ttarih1;
        public static DateTime Ttarih2;
        public static int FTarihAy;
        public static int FTarihYil;
        public static int CTarihAy;
        public static int CTarihYil;
        public static int KisiIdDetay;
        public static string KisiAdDetay;
        public static string KisiTel1Detay;
        public static string KisiTel2Detay;
        public static string KisiAdresDetay;
        public static string KisiFirmaDetay;

        //Acilis
        private void Form1_Load(object sender, EventArgs e)
        {           
            //Bilgi -> Gider ComboBox Varsayılan a Getir
            int foryear = 2017;
            for (int i = 0; i < (DateTime.Now.Year - 2017); i++)
            {
                foryear++;
                CmbYil.Items.Add(foryear);
            }
            CmbYil.Items.Add("Tümü");
            CmbYil.SelectedIndex = DateTime.Now.Year - 2017;
            CmbBilgi.SelectedIndex = 5;

            int ay = DateTime.Now.Month;
            if (ay == 1)
            { CmbGider.SelectedIndex = 1; }
            else if (ay == 2)
            { CmbGider.SelectedIndex = 2; }
            else if (ay == 3)
            { CmbGider.SelectedIndex = 3; }
            else if (ay == 4)
            { CmbGider.SelectedIndex = 4; }
            else if (ay == 5)
            { CmbGider.SelectedIndex = 5; }
            else if (ay == 6)
            { CmbGider.SelectedIndex = 6; }
            else if (ay == 7)
            { CmbGider.SelectedIndex = 7; }
            else if (ay == 8)
            { CmbGider.SelectedIndex = 8; }
            else if (ay == 9)
            { CmbGider.SelectedIndex = 9; }
            else if (ay == 10)
            { CmbGider.SelectedIndex = 10; }
            else if (ay == 11)
            { CmbGider.SelectedIndex = 11; }
            else if (ay == 12)
            { CmbGider.SelectedIndex = 12; }
            else
            { CmbGider.SelectedIndex = 0; }

            var sonuc = from kisiler in ctx.Kisis
                        orderby kisiler.Ad
                        select new
                        {
                            kisiler.KisiID,
                            kisiler.Ad,
                            kisiler.Firma,
                            kisiler.Tel1,
                            kisiler.Tel2,
                            kisiler.Adres,
                            kisiler.Karaliste
                        };
            dataGridView1.DataSource = sonuc;
            dataGridView1.Columns["Firma"].Visible = dataGridView1.Columns["Tel1"].Visible = dataGridView1.Columns["Tel2"].Visible = dataGridView1.Columns["Adres"].Visible = dataGridView1.Columns["KisiID"].Visible = dataGridView1.Columns["Karaliste"].Visible = false;
            DataGridViewColumn column = dataGridView1.Columns[1];
            column.Width = dataGridView1.Width - 20;

           
            LblKisiSayi.Text = dataGridView1.Rows.Count.ToString();
            LblVeriSayi.Text = ctx.Caris.Count().ToString();

            Dtp1.Value = DateTime.Today.AddDays(-1);
        }


        private void TxtAra_TextChanged(object sender, EventArgs e)
        {
            var sonuc = from kisiler in ctx.Kisis
                        orderby kisiler.Ad
                        select new
                        {
                            kisiler.KisiID,
                            kisiler.Ad,
                            kisiler.Firma,
                            kisiler.Tel1,
                            kisiler.Tel2,
                            kisiler.Adres,
                            kisiler.Karaliste
                        };
            dataGridView1.DataSource = sonuc.Where(x => x.Ad.Contains(txtAra.Text) || x.Firma.Contains(txtAra.Text));
            dataGridView1.Columns["Firma"].Visible = dataGridView1.Columns["Tel1"].Visible = dataGridView1.Columns["Tel2"].Visible = dataGridView1.Columns["Adres"].Visible = dataGridView1.Columns["KisiID"].Visible = dataGridView1.Columns["Karaliste"].Visible = false;
            DataGridViewColumn column = dataGridView1.Columns[1];
            column.Width = dataGridView1.Width - 20;
        }

        private void DataGridView1_SelectionChanged(object sender, EventArgs e)
        {
            DataGridViewRow row = dataGridView1.CurrentRow;
            if (anahtar == true)
            {
                //Düzenle Kısmını Temizle
                txtAciklama.Text = "";
                cmbDurum.SelectedValue = -1;
                dtTarih.Text = DateTime.Now.ToString("d/M/yyyy");
                txtTutar.Text = "";

                txtAraCari.Text = "";
                TxtAraDuzelt.Text = "";


                //Soldaki isim secildiginde sag-ustte labeller degisir.
                
                lblID.Text = row.Cells["KisiID"].Value.ToString();
                lblAd.Text = lblDAd.Text = row.Cells["Ad"].Value.ToString();
                lblID.Tag = row.Cells["KisiID"].Value;                

                LblKaraListe.Tag = row.Cells["Karaliste"].Value;
                if ((Boolean)LblKaraListe.Tag)
                {
                    LblKaraListe.ForeColor = Color.DarkRed;
                }
                else
                {
                    LblKaraListe.ForeColor = Color.White;
                }

                if ((string)row.Cells["Firma"].Value == "")
                {
                    label2.Visible = false;
                    lblFirma.Visible = false;
                }
                else
                {
                    label2.Visible = true;
                    lblFirma.Visible = true;
                    lblFirma.Text = row.Cells["Firma"].Value.ToString();
                }

                if ((string)row.Cells["Tel1"].Value == "")
                {
                    label4.Visible = false;
                    lblTel1.Visible = false;
                }
                else
                {
                    label4.Visible = true;
                    lblTel1.Visible = true;
                    lblTel1.Text = row.Cells["Tel1"].Value.ToString();
                }

                if ((string)row.Cells["Tel2"].Value == "")
                {
                    lblTel2B.Visible = false;
                    lblTel2.Visible = false;
                }
                else
                {
                    lblTel2B.Visible = true;
                    lblTel2.Visible = true;
                    lblTel2.Text = row.Cells["Tel2"].Value.ToString();
                }

                if ((string)row.Cells["Adres"].Value == "")
                {
                    lblAdresB.Visible = false;
                    lblAdres.Visible = false;
                }
                else
                {
                    lblAdresB.Visible = true;
                    lblAdres.Visible = true;
                    lblAdres.Text = row.Cells["Adres"].Value.ToString();
                }
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
                    label60.BackColor = Color.ForestGreen;
                }
                else
                {
                    LblKalanAlacak.BackColor = Color.Maroon;
                    label50.BackColor = Color.Maroon;
                    label48.BackColor = Color.Maroon;
                    label60.BackColor = Color.Maroon;
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


                //Soldaki isim secildiginde sag-alt taki datagridview2 degisir.
                var sonuc2 = from kisiler in ctx.Kisis
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
                dataGridView2.DataSource = sonuc2.Where(x => x.KisiID == Convert.ToInt32(row.Cells["KisiID"].Value));
                dataGridView2.Columns["KisiID"].Visible = dataGridView2.Columns["Ad"].Visible = dataGridView2.Columns["DurumID"].Visible = dataGridView2.Columns["CariID"].Visible = false;

                dataGridView2.Columns[3].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                dataGridView2.Columns[3].HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleRight;
                dataGridView2.Columns[4].HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
                dataGridView2.Columns[4].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                //dataGridView2.AutoResizeColumns();

                //Düzenle ve Sil Kısmı

                txtGAd.Text = lblSAd.Text = row.Cells["Ad"].Value.ToString();
                txtGFirma.Text = lblSFirma.Text = row.Cells["Firma"].Value.ToString();
                txtGTel1.Text = lblSTel1.Text = (string)row.Cells["Tel1"].Value;
                txtGTel2.Text = lblSTel2.Text = (string)row.Cells["Tel2"].Value;
                txtGAdres.Text = lblSAdres.Text = row.Cells["Adres"].Value.ToString();
                txtGAd.Tag = row.Cells["KisiId"].Value;
                if ((Boolean)LblKaraListe.Tag)
                {
                    ChkKaraListe.Checked = true;
                }
                else
                {
                    ChkKaraListe.Checked = false;
                }

                cmbDurum.DisplayMember = "Durumlar";
                cmbDurum.ValueMember = "DurumID";
                cmbDurum.DataSource = ctx.Durums;
                //DataGridView3 Hizalama
                dataGridView3.DataSource = sonuc2.Where(x => x.KisiID == Convert.ToInt32(row.Cells["KisiID"].Value));
                dataGridView3.Columns["KisiID"].Visible = dataGridView3.Columns["Ad"].Visible = dataGridView3.Columns["DurumID"].Visible = dataGridView3.Columns["CariID"].Visible = false;

                dataGridView3.Columns[3].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                dataGridView3.Columns[3].HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleRight;
                dataGridView3.Columns[4].HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
                dataGridView3.Columns[4].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            }
        }

        private void DataGridView2_DataBindingComplete(object sender, DataGridViewBindingCompleteEventArgs e)
        {
            foreach (DataGridViewRow dGVRow in this.dataGridView2.Rows)
            {
                dGVRow.HeaderCell.Value = String.Format("{0}", dGVRow.Index + 1);
            }
            //Genişlik Ayarı
            this.dataGridView2.AutoResizeRowHeadersWidth(DataGridViewRowHeadersWidthSizeMode.AutoSizeToAllHeaders);
        }

        private void DataGridView3_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            DataGridViewRow row2 = dataGridView3.CurrentRow;
            //row 2 was null

            if (row2.Cells["Tutar"].Value != null)
            {
                txtTutar.Text = row2.Cells["Tutar"].Value.ToString();
            }
            else
            {
                txtTutar.Text = "";
            }

            if (row2.Cells["Tarih"].Value != null)
            {
                dtTarih.Text = row2.Cells["Tarih"].Value.ToString();
            }
            else
            {
                dtTarih.Text = DateTime.Now.ToString("M/d/yyyy");
            }

            if (row2.Cells["DurumID"].Value != null)
            {
                cmbDurum.SelectedValue = row2.Cells["DurumID"].Value;
            }
            else
            {
                cmbDurum.SelectedValue = -1;
            }


            if (row2.Cells["Açıklama"].Value != null)
            {
                txtAciklama.Text = row2.Cells["Açıklama"].Value.ToString();
                txtAciklama.Tag = row2.Cells["CariId"].Value;
            }
            else
            {
                txtAciklama.Text = "";
            }
        }

        private void BtnGuncelle_Click(object sender, EventArgs e)
        {
            if (cmbDurum.SelectedIndex != -1)
            {
                Button btn = sender as Button;

                int GuncelledId2 = dataGridView3.SelectedRows[0].Index;
                string GuncelleId = (GuncelledId2 + 1).ToString();
                string GuncelleDurum = dataGridView3.SelectedCells[2].Value.ToString();
                string GuncelleTutar = dataGridView3.SelectedCells[3].Value.ToString();
                string GuncelleTarih = string.Format("{0:dd/MM/yyyy}", dataGridView3.SelectedCells[4].Value);
                string GuncelleAciklama = "Açıklama";
                if (dataGridView3.SelectedCells[5].Value.ToString() == "")
                {
                    GuncelleAciklama = "Açıklama";
                }
                else
                {
                    GuncelleAciklama = dataGridView3.SelectedCells[5].Value.ToString();
                }

                DialogResult sonuc = MessageBox.Show(GuncelleId + ": " + GuncelleDurum + " - " + GuncelleTutar + " - " + GuncelleTarih + " - " + GuncelleAciklama + "\n" + "\n" + "Aşağıdaki yeni veri ile," + "\n" + "\n" + cmbDurum.Text + " - " + txtTutar.Text + " - " + dtTarih.Text + " - " + txtAciklama.Text + "\n" + "\n" + "Güncellenecek. Emin misiniz?", "Güncelleme", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

                //Güncelle
                if (sonuc == DialogResult.Yes)
                {
                    anahtar = false;

                    int id2 = (int)txtAciklama.Tag;

                    Cari c = ctx.Caris.SingleOrDefault(x => x.CariID == id2);
                    c.Tutar = Convert.ToDecimal(txtTutar.Text);
                    c.Tarih = dtTarih.Value;
                    c.DurumID = (byte)cmbDurum.SelectedValue;
                    c.Aciklama = txtAciklama.Text;

                    ctx.SubmitChanges();
                    txtAraCari.Text = "";
                    TxtAraDuzelt.Text = "";

                    //Değişiklikleri altta düzenle bölümünde listele
                    var liste2 = from kisiler in ctx.Kisis
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

                    dataGridView3.DataSource = liste2.Where(x => x.KisiID == Convert.ToInt32(lblID.Tag));
                    dataGridView3.Columns["KisiID"].Visible = dataGridView3.Columns["Ad"].Visible = dataGridView3.Columns["DurumID"].Visible = dataGridView3.Columns["CariID"].Visible = false;

                    dataGridView3.Columns[3].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                    dataGridView3.Columns[3].HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleRight;
                    dataGridView3.Columns[4].HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
                    dataGridView3.Columns[4].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;

                    //Allta Bilgiler Bölümünde Listele
                    dataGridView2.DataSource = liste2.Where(x => x.KisiID == Convert.ToInt32(lblID.Tag));
                    dataGridView2.Columns["KisiID"].Visible = dataGridView2.Columns["Ad"].Visible = dataGridView2.Columns["DurumID"].Visible = dataGridView2.Columns["CariID"].Visible = false;

                    dataGridView2.Columns[3].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                    dataGridView2.Columns[3].HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleRight;
                    dataGridView2.Columns[4].HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
                    dataGridView2.Columns[4].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;

                    //Üstte Bilgiler Bölümünde Listele
                    DataGridViewRow row = dataGridView1.CurrentRow;
                    lblAd.Text = row.Cells["Ad"].Value.ToString();
                    lblDAd.Text = row.Cells["Ad"].Value.ToString();
                    if ((string)row.Cells["Firma"].Value == "")
                    {
                        label2.Visible = false;
                        lblFirma.Visible = false;
                    }
                    else
                    {
                        label2.Visible = true;
                        lblFirma.Visible = true;
                        lblFirma.Text = row.Cells["Firma"].Value.ToString();
                    }

                    if ((string)row.Cells["Tel1"].Value == "")
                    {
                        label4.Visible = false;
                        lblTel1.Visible = false;
                    }
                    else
                    {
                        label4.Visible = true;
                        lblTel1.Visible = true;
                        lblTel1.Text = row.Cells["Tel1"].Value.ToString();
                    }

                    if ((string)row.Cells["Tel2"].Value == "")
                    {
                        lblTel2B.Visible = false;
                        lblTel2.Visible = false;
                    }
                    else
                    {
                        lblTel2B.Visible = true;
                        lblTel2.Visible = true;
                        lblTel2.Text = row.Cells["Tel2"].Value.ToString();
                    }

                    if ((string)row.Cells["Adres"].Value == "")
                    {
                        lblAdresB.Visible = false;
                        lblAdres.Visible = false;
                    }
                    else
                    {
                        lblAdresB.Visible = true;
                        lblAdres.Visible = true;
                        lblAdres.Text = row.Cells["Adres"].Value.ToString();
                    }

                    //Hesaplama
                    var hesap = ctx.Kisis.Join(ctx.Caris,
                        kisiler => kisiler.KisiID,
                        cariler => cariler.KisiID,
                        (ki, ca) => new
                        {
                            kKisiID = ki.KisiID,
                            cKisiID = ca.KisiID,
                            ca.Tutar,
                            ca.DurumID
                        }).Select(x => new
                        {
                            x.cKisiID,
                            x.kKisiID,
                            x.DurumID,
                            x.Tutar
                        });
                    //Alacak
                    var alacak = hesap.Where(x => x.cKisiID == Convert.ToInt32(lblID.Tag)).Where(x => x.DurumID == 1).Sum(x => x.Tutar);
                    var alindi = hesap.Where(x => x.cKisiID == Convert.ToInt32(lblID.Tag)).Where(x => x.DurumID == 2).Sum(x => x.Tutar);

                    if (alacak != null && alindi != null)
                    {
                        lblAlacak.Text = String.Format("{0:N}\n", (alacak - alindi));
                    }
                    else if (alacak != null && alindi == null)
                    {
                        lblAlacak.Text = String.Format("{0:N}\n", alacak);
                    }
                    else
                    {
                        lblAlacak.Text = String.Format("{0:N}\n", 0);
                    }

                    //Borç
                    var borc = hesap.Where(x => x.cKisiID == Convert.ToInt32(lblID.Tag)).Where(x => x.DurumID == 3).Sum(x => x.Tutar);
                    var odendi = hesap.Where(x => x.cKisiID == Convert.ToInt32(lblID.Tag)).Where(x => x.DurumID == 4).Sum(x => x.Tutar);

                    if (borc != null && odendi != null)
                    {
                        lblBorc.Text = String.Format("{0:N}\n", (borc - odendi));
                    }
                    else if (borc != null && odendi == null)
                    {
                        lblBorc.Text = String.Format("{0:N}\n", borc);
                    }
                    else
                    {
                        lblBorc.Text = String.Format("{0:N}\n", 0);
                    }

                    //Anahtar
                    anahtar = true;
                }
            }
            else
            {
                MessageBox.Show("Seçili satır yok!", "Uyarı", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        private void BtnEkle_Click(object sender, EventArgs e)
        {
            if (cmbDurum.SelectedIndex != -1)
            {
                //Ekle
                Cari car = new Cari
                {
                    Aciklama = txtAciklama.Text
                };

                if (txtTutar.Text == "")
                {
                    car.Tutar = 0;
                }
                else
                {
                    car.Tutar = Convert.ToDecimal(txtTutar.Text);
                }

                car.DurumID = (byte)cmbDurum.SelectedValue;
                car.Tarih = dtTarih.Value;
                car.KisiID = Convert.ToInt32(lblID.Tag);

                ctx.Caris.InsertOnSubmit(car);
                ctx.SubmitChanges();
                TxtAraDuzelt.Text = "";
                txtAraCari.Text = "";
                //Listele Aşağıda - Bilgi ve Düzenle
                var liste2 = from kisiler in ctx.Kisis
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

                dataGridView3.DataSource = liste2.Where(x => x.KisiID == Convert.ToInt32(lblID.Tag));
                dataGridView3.Columns["KisiID"].Visible = dataGridView3.Columns["Ad"].Visible = dataGridView3.Columns["DurumID"].Visible = dataGridView3.Columns["CariID"].Visible = false;

                dataGridView3.Columns[3].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                dataGridView3.Columns[3].HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleRight;
                dataGridView3.Columns[4].HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
                dataGridView3.Columns[4].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;

                //Allta Bilgiler Bölümünde Listele
                dataGridView2.DataSource = liste2.Where(x => x.KisiID == Convert.ToInt32(lblID.Tag));
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
                        ca.DurumID
                    }).Select(x => new
                    {
                        x.cKisiID,
                        x.kKisiID,
                        x.DurumID,
                        x.Tutar
                    });
                //Alacak
                var alacak = hesap.Where(x => x.cKisiID == Convert.ToInt32(lblID.Tag)).Where(x => x.DurumID == 1).Sum(x => x.Tutar);
                var alindi = hesap.Where(x => x.cKisiID == Convert.ToInt32(lblID.Tag)).Where(x => x.DurumID == 2).Sum(x => x.Tutar);

                if (alacak != null && alindi != null)
                {
                    lblAlacak.Text = String.Format("{0:N}\n", (alacak - alindi));
                }
                else if (alacak != null && alindi == null)
                {
                    lblAlacak.Text = String.Format("{0:N}\n", alacak);
                }
                else
                {
                    lblAlacak.Text = String.Format("{0:N}\n", 0);
                }

                //Borç
                var borc = hesap.Where(x => x.cKisiID == Convert.ToInt32(lblID.Tag)).Where(x => x.DurumID == 3).Sum(x => x.Tutar);
                var odendi = hesap.Where(x => x.cKisiID == Convert.ToInt32(lblID.Tag)).Where(x => x.DurumID == 4).Sum(x => x.Tutar);

                if (borc != null && odendi != null)
                {
                    lblBorc.Text = String.Format("{0:N}\n", (borc - odendi));
                }
                else if (borc != null && odendi == null)
                {
                    lblBorc.Text = String.Format("{0:N}\n", borc);
                }
                else
                {
                    lblBorc.Text = String.Format("{0:N}\n", 0);
                }

                //Alanları temizle
                txtTutar.Text = null;
                txtAciklama.Text = null;
                dtTarih.Text = Convert.ToString(DateTime.Now);
                cmbDurum.SelectedIndex = -1;
            }
            else
            {
                MessageBox.Show("Lütfen bir 'Durum' seçiniz!", "Uyarı", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        private void BtnSil_Click(object sender, EventArgs e)
        {
            if (cmbDurum.SelectedIndex != -1)
            {
                Button btn = sender as Button;

                int SildId2 = dataGridView3.SelectedRows[0].Index;
                string SilId = (SildId2 + 1).ToString();
                string SilDurum = dataGridView3.SelectedCells[2].Value.ToString();
                string SilTarih = string.Format("{0:dd/MM/yyyy}", dataGridView3.SelectedCells[4].Value);
                string SilAciklama = "Açıklama";
                if (dataGridView3.SelectedCells[5].Value.ToString() == "")
                {
                    SilAciklama = "Açıklama";
                }
                else
                {
                    SilAciklama = dataGridView3.SelectedCells[5].Value.ToString();
                }
                
                DialogResult sonuc = MessageBox.Show("' " + SilId + ": " + SilDurum + " / " + SilTarih + " / " + SilAciklama + " '" + "\n" + "\n" + " silinecek. Emin misiniz?", "Silme", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
                //Sil
                if (sonuc == DialogResult.Yes)
                {
                    if (dataGridView3.CurrentRow == null) return;

                    int cariId = (int)dataGridView3.CurrentRow.Cells["CariID"].Value;

                    Cari c = ctx.Caris.SingleOrDefault(id => id.CariID == cariId);
                    ctx.Caris.DeleteOnSubmit(c);
                    ctx.SubmitChanges();
                    TxtAraDuzelt.Text = "";
                    txtAraCari.Text = "";
                    //Listele Aşağıda - Bilgi ve Düzenle
                    var liste2 = from kisiler in ctx.Kisis
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

                    dataGridView3.DataSource = liste2.Where(x => x.KisiID == Convert.ToInt32(lblID.Tag));
                    dataGridView3.Columns["KisiID"].Visible = dataGridView3.Columns["Ad"].Visible = dataGridView3.Columns["DurumID"].Visible = dataGridView3.Columns["CariID"].Visible = false;

                    dataGridView3.Columns[3].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                    dataGridView3.Columns[3].HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleRight;
                    dataGridView3.Columns[4].HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
                    dataGridView3.Columns[4].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;

                    //Allta Bilgiler Bölümünde Listele
                    dataGridView2.DataSource = liste2.Where(x => x.KisiID == Convert.ToInt32(lblID.Tag));
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
                            ca.DurumID
                        }).Select(x => new
                        {
                            x.cKisiID,
                            x.kKisiID,
                            x.DurumID,
                            x.Tutar
                        });
                    //Alacak
                    var alacak = hesap.Where(x => x.cKisiID == Convert.ToInt32(lblID.Tag)).Where(x => x.DurumID == 1).Sum(x => x.Tutar);
                    var alindi = hesap.Where(x => x.cKisiID == Convert.ToInt32(lblID.Tag)).Where(x => x.DurumID == 2).Sum(x => x.Tutar);

                    if (alacak != null && alindi != null)
                    {
                        lblAlacak.Text = String.Format("{0:N}\n", (alacak - alindi));
                    }
                    else if (alacak != null && alindi == null)
                    {
                        lblAlacak.Text = String.Format("{0:N}\n", alacak);
                    }
                    else
                    {
                        lblAlacak.Text = String.Format("{0:N}\n", 0);
                    }

                    //Borç
                    var borc = hesap.Where(x => x.cKisiID == Convert.ToInt32(lblID.Tag)).Where(x => x.DurumID == 3).Sum(x => x.Tutar);
                    var odendi = hesap.Where(x => x.cKisiID == Convert.ToInt32(lblID.Tag)).Where(x => x.DurumID == 4).Sum(x => x.Tutar);

                    if (borc != null && odendi != null)
                    {
                        lblBorc.Text = String.Format("{0:N}\n", (borc - odendi));
                    }
                    else if (borc != null && odendi == null)
                    {
                        lblBorc.Text = String.Format("{0:N}\n", borc);
                    }
                    else
                    {
                        lblBorc.Text = String.Format("{0:N}\n", 0);
                    }

                    //Alanları temizle
                    txtTutar.Text = null;
                    txtAciklama.Text = null;
                    dtTarih.Text = Convert.ToString(DateTime.Now);
                    cmbDurum.SelectedIndex = -1;
                }   
            }
            else
            {
                MessageBox.Show("Seçili satır yok!", "Uyarı", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        private void DataGridView3_DataBindingComplete(object sender, DataGridViewBindingCompleteEventArgs e)
        {
            foreach (DataGridViewRow dGVRow in this.dataGridView3.Rows)
            {
                dGVRow.HeaderCell.Value = String.Format("{0}", dGVRow.Index + 1);
            }
            //Genişlik Ayarı
            this.dataGridView3.AutoResizeRowHeadersWidth(DataGridViewRowHeadersWidthSizeMode.AutoSizeToAllHeaders);
        }

        private void DataGridView4_DataBindingComplete(object sender, DataGridViewBindingCompleteEventArgs e)
        {
            foreach (DataGridViewRow dGVRow in this.dataGridView4.Rows)
            {
                dGVRow.HeaderCell.Value = String.Format("{0}", dGVRow.Index + 1);
            }
            //Genişlik Ayarı
            this.dataGridView4.AutoResizeRowHeadersWidth(DataGridViewRowHeadersWidthSizeMode.AutoSizeToAllHeaders);           
        }

        private void BtnToplamAlacak_Click(object sender, EventArgs e)
        {
            lblRapor.Text = "ALACAKLARIM";
            dataGridView4.DataSource = null;

            //Hesaplama - Alacağım
            var alacagim = from kisiler in ctx.Kisis
                           join cariler in ctx.Caris on kisiler.KisiID equals cariler.KisiID
                           join durumlar in ctx.Durums on cariler.DurumID equals durumlar.DurumID

                           group cariler by new { cariler.KisiID, kisiler.Karaliste, kisiler.Ad, kisiler.Firma, kisiler.Tel1 } into grup
                           let KisiID = grup.Key.KisiID
                           let Karaliste = grup.Key.Karaliste
                           let Ad = grup.Key.Ad
                           let Firma = grup.Key.Firma
                           let Telefon = grup.Key.Tel1
                           let SonTarih = grup.Where(x => x.DurumID == 2).Max(x => x.Tarih)
                           let Bakiye = grup.Where(x => x.DurumID == 1).Sum(x => (decimal?)x.Tutar) ?? 0
                           let Alındı = grup.Where(x => x.DurumID == 2).Sum(x => (decimal?)x.Tutar) ?? 0
                           let iskonto = grup.Where(x => x.DurumID == 9).Sum(x => (decimal?)x.Tutar) ?? 0
                           let Alacağım = (Bakiye - (Alındı + iskonto))

                           orderby grup.Key.Ad
                           select new
                           {
                               grup.Key.KisiID,
                               grup.Key.Karaliste,
                               SonTarih,
                               grup.Key.Ad,
                               Firma,
                               Telefon,
                               Bakiye,
                               Alındı,
                               Alacağım
                           };

            dataGridView4.DataSource = alacagim.Where(x => x.Alacağım != 0);

            dataGridView4.Columns["KisiID"].Visible = dataGridView4.Columns["Firma"].Visible = dataGridView4.Columns["Karaliste"].Visible = false;

            dataGridView4.Columns[3].HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleLeft;
            dataGridView4.Columns[3].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft;
            dataGridView4.Columns[4].HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleLeft;
            dataGridView4.Columns[5].HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dataGridView4.Columns[5].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            //dataGridView4.Columns[6].HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleRight;
            dataGridView4.Columns[6].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            //dataGridView4.Columns[7].HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleRight;
            dataGridView4.Columns[7].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            //dataGridView4.Columns[8].HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleRight;
            dataGridView4.Columns[8].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;

            dataGridView4.Columns["SonTarih"].Width = 100;
            dataGridView4.Columns["SonTarih"].HeaderText = "Son İşlem";
            dataGridView4.Columns["Ad"].Width = 220;
            dataGridView4.Columns["Telefon"].Width = 115;
            dataGridView4.Columns["Bakiye"].Width = 160;
            dataGridView4.Columns["Alındı"].Width = 133;
            dataGridView4.Columns["Alacağım"].Width = 160;

            /*
            // <- Eski -> //
            //Hesaplama - Alacağım
            var alacagim = from kisiler in ctx.Kisis
                           join cariler in ctx.Caris on kisiler.KisiID equals cariler.KisiID
                           join durumlar in ctx.Durums on cariler.DurumID equals durumlar.DurumID

                           group cariler by new { cariler.KisiID, kisiler.Karaliste, kisiler.Ad, kisiler.Tel1 } into grup
                           let Ad = grup.Key.Ad
                           let KisiID = grup.Key.KisiID
                           let Karaliste = grup.Key.Karaliste
                           let Telefon = grup.Key.Tel1
                           let Bakiye = grup.Where(x => x.DurumID == 1).Sum(x => (decimal?)x.Tutar) ?? 0
                           let Alındı = grup.Where(x => x.DurumID == 2).Sum(x => (decimal?)x.Tutar) ?? 0
                           let iskonto = grup.Where(x => x.DurumID == 9).Sum(x => (decimal?)x.Tutar) ?? 0
                           let Alacağım = (Bakiye - (Alındı + iskonto))

                           orderby grup.Key.Ad
                           select new
                           {
                               grup.Key.KisiID,
                               grup.Key.Karaliste,
                               grup.Key.Ad,
                               Telefon,
                               Bakiye,
                               Alındı,
                               Alacağım
                           };

            dataGridView4.DataSource = alacagim.Where(x => x.Alacağım != 0);
            dataGridView4.Columns["KisiID"].Visible = dataGridView4.Columns["Karaliste"].Visible = false;

            dataGridView4.Columns[2].HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleLeft;
            dataGridView4.Columns[3].HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dataGridView4.Columns[3].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dataGridView4.Columns[4].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            dataGridView4.Columns[5].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            dataGridView4.Columns[6].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;

            dataGridView4.Columns["Ad"].Width = 220;
            dataGridView4.Columns["Telefon"].Width = 115;
            dataGridView4.Columns["Bakiye"].Width = 160;
            dataGridView4.Columns["Alındı"].Width = 133;
            dataGridView4.Columns["Alacağım"].Width = 160;
            */
        }


        private void BtnToplamBorc_Click(object sender, EventArgs e)
        {
            lblRapor.Text = "BORÇLARIM";
            dataGridView4.DataSource = null;
            //Hesaplama - Borcum
            var borcum = from kisiler in ctx.Kisis
                         join cariler in ctx.Caris on kisiler.KisiID equals cariler.KisiID
                         join durumlar in ctx.Durums on cariler.DurumID equals durumlar.DurumID

                         group cariler by new { cariler.KisiID, kisiler.Karaliste, kisiler.Ad, kisiler.Firma, kisiler.Tel1 } into grup
                         let KisiID = grup.Key.KisiID
                         let Karaliste = grup.Key.Karaliste
                         let Ad = grup.Key.Ad
                         let Telefon = grup.Key.Tel1
                         let Bakiye = grup.Where(x => x.DurumID == 3).Sum(x => (decimal?)x.Tutar) ?? 0
                         let Ödendi = grup.Where(x => x.DurumID == 4).Sum(x => (decimal?)x.Tutar) ?? 0
                         let Borcum = (Bakiye - Ödendi)

                         orderby grup.Key.Ad
                         select new
                         {
                             grup.Key.KisiID,
                             grup.Key.Karaliste,
                             grup.Key.Ad,
                             grup.Key.Firma,
                             Telefon,
                             Bakiye,
                             Ödendi,
                             Borcum
                         };

            dataGridView4.DataSource = borcum.Where(x => x.Borcum != 0);
            dataGridView4.Columns["KisiID"].Visible = dataGridView4.Columns["Firma"].Visible = dataGridView4.Columns["Karaliste"].Visible = false;

            dataGridView4.Columns[2].HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleLeft;
            dataGridView4.Columns[4].HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dataGridView4.Columns[4].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dataGridView4.Columns[5].HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleRight;
            dataGridView4.Columns[5].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            //dataGridView4.Columns[4].HeaderCell.Style.Alignment = DataGridViewContentAlignment.BottomRight;
            dataGridView4.Columns[6].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            dataGridView4.Columns[6].HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleRight;
            dataGridView4.Columns[7].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            dataGridView4.Columns[7].HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleRight;

            dataGridView4.Columns["Ad"].Width = 220;
            dataGridView4.Columns["Telefon"].Width = 115;
            dataGridView4.Columns["Bakiye"].Width = 160;
            dataGridView4.Columns["Ödendi"].Width = 150;
            dataGridView4.Columns["Borcum"].Width = 150;
        }

        private void BtnExcel1_Click(object sender, EventArgs e)
        {
            CultureInfo customCulture = (CultureInfo)System.Threading.Thread.CurrentThread.CurrentCulture.Clone();
            customCulture.NumberFormat.NumberDecimalSeparator = ".";
            System.Threading.Thread.CurrentThread.CurrentCulture = customCulture;
            int GelirGider = 0;
            if (lblRapor.Text != "RAPOR")
            {
                //Dosya ismi
                string fileName;

                SaveFileDialog saveFileDialog1 = new SaveFileDialog
                {
                    Filter = "xls files (*.xlsx)|*.xlsx|All files (*.*)|*.*",
                    Title = "To Excel",
                    FileName = "(" + DateTime.Now.ToString("dd-MM-yyyy") + ")" + " " + lblRapor.Text
                };
               

                if (saveFileDialog1.ShowDialog() == DialogResult.OK)
                {
                    fileName = saveFileDialog1.FileName;
                    var workbook = new XLWorkbook();

                    //-Liste A - //

                    //DataTable2 Oluştur
                    DataTable dt = new DataTable();

                    //No2 Sütunu Ekle
                    DataColumn columno = new DataColumn
                    {
                        DataType = System.Type.GetType("System.Int32"),
                        AutoIncrement = true,
                        AutoIncrementSeed = 1,
                        AutoIncrementStep = 1
                    };
                    dt.Columns.Add(columno);
                    dt.Columns["Column1"].ColumnName = "No";

                    //Sütunları Ekle
                    foreach (DataGridViewColumn column in dataGridView4.Columns)
                    {
                        //dt.Columns.Add(column.HeaderText, column.ValueType);
                        dt.Columns.Add(column.HeaderText);
                    }

                    //if (dt.Columns[5].ColumnName == "Kalan_Alacağım" || dt.Columns[3].ColumnName == "Kalan_Alacağım") { dt.Columns["Kalan_Alacağım"].ColumnName = "Kalan Alacağım"; }
                    //if (dt.Columns[5].ColumnName == "Kalan_Borcum" || dt.Columns[4].ColumnName == "Kalan_Borcum") { dt.Columns["Kalan_Borcum"].ColumnName = "Kalan Borcum"; }
                    if (dt.Columns.Contains("Kalan_Alacağım"))
                    {
                        dt.Columns["Kalan_Alacağım"].ColumnName = "Alacağım";
                    }
                    if (dt.Columns.Contains("Kalan_Borcum"))
                    {
                        dt.Columns["Kalan_Borcum"].ColumnName = "Borcum";
                    }

                    //Satırları Ekle
                    foreach (DataGridViewRow row in dataGridView4.Rows)
                    {
                        dt.Rows.Add();
                        //Hücreleri Ekle
                        foreach (DataGridViewCell cell in row.Cells)
                        {
                            if (cell.Value != null)
                            {
                                //dt.Rows[dt.Rows.Count - 1][cell.ColumnIndex + 1] = String.Format("{0:N}\n", cell.Value);
                                dt.Rows[dt.Rows.Count - 1][cell.ColumnIndex + 1] = cell.Value;
                            }
                        }
                    }

                    //Toplam Satırı Ekle
                    DataRow rowToplam = dt.NewRow();
                    dt.Rows.Add();
                    dt.Rows.Add(rowToplam);
                    if (lblRapor.Text == "ALACAKLARIM")
                    {
                        rowToplam[6] = "Toplam :";
                    }
                    else if (lblRapor.Text == "BORÇLARIM")
                    {
                        rowToplam[5] = "Toplam :";
                    }

                    if (lblRapor.Text == "ALACAKLARIM")
                    {
                        for (int i = 6; i < dataGridView4.Columns.Count; i++)
                        {
                            decimal toplam = 0;
                            for (int j = 0; j < dataGridView4.Rows.Count + 1; ++j)
                            {
                                if (j < dataGridView4.Rows.Count)
                                {
                                    toplam += Convert.ToDecimal(dataGridView4.Rows[j].Cells[i].Value);
                                }
                                else
                                {
                                    //rowToplam[i + 1] = String.Format("{0:N}\n", toplam);
                                    rowToplam[i + 1] = toplam;
                                }
                            }
                        }
                    }
                    else if (lblRapor.Text == "BORÇLARIM")
                    {
                        for (int i = 5; i < dataGridView4.Columns.Count; i++)
                        {
                            decimal toplam = 0;
                            for (int j = 0; j < dataGridView4.Rows.Count + 1; ++j)
                            {
                                if (j < dataGridView4.Rows.Count)
                                {
                                    toplam += Convert.ToDecimal(dataGridView4.Rows[j].Cells[i].Value);
                                }
                                else
                                {
                                    //rowToplam[i + 1] = String.Format("{0:N}\n", toplam);
                                    rowToplam[i + 1] = toplam;
                                }
                            }
                        }
                    }

                    //Gereksiz Sütunu Kaldır
                    if (lblRapor.Text == "ALACAKLARIM")
                    {
                        dt.Columns.RemoveAt(1);
                        dt.Columns.RemoveAt(1);
                        dt.Columns.RemoveAt(1);
                        dt.Columns.RemoveAt(2);
                    }
                    else if (lblRapor.Text == "BORÇLARIM")
                    {
                        dt.Columns.RemoveAt(1);
                        dt.Columns.RemoveAt(1);
                        dt.Columns.RemoveAt(2);
                    }
                    
                    /*
                    //Tarihler Arası Gelir-Gider Fark ı Tablo2 Altına Ekle
                    if (ChkGelirGider.Checked == true)
                    {
                        dt.Rows.Add();

                        DataRow rowTarih = dt.NewRow();
                        dt.Rows.Add(rowTarih);
                        rowTarih[2] = (string.Format("{0:dd/MM/yyyy}", Ttarih1)) + "-" + (string.Format("{0:dd/MM/yyyy}", Ttarih2));

                        rowTarih[2] = (string.Format("{0:dd/MM/yyyy}", Ttarih1)) + "-" + (string.Format("{0:dd/MM/yyyy}", Ttarih2));

                        DataRow rowAlindi = dt.NewRow();
                        dt.Rows.Add(rowAlindi);
                        rowAlindi[2] = "Alındı:";

                        DataRow rowOdendi = dt.NewRow();
                        dt.Rows.Add(rowOdendi);
                        //rowOdendi[1] = (string.Format("{0:dd/MM/yyyy}", Ttarih1)) + " - " + (string.Format("{0:dd/MM/yyyy}", Ttarih2));
                        rowOdendi[2] = "Ödendi:";

                        DataRow rowGider = dt.NewRow();
                        dt.Rows.Add(rowGider);
                        //rowGider[1] = (string.Format("{0:dd/MM/yyyy}", Ttarih1)) + " - " + (string.Format("{0:dd/MM/yyyy}", Ttarih2));
                        rowGider[2] = "Giderler:";

                        DataRow rowFark = dt.NewRow();
                        dt.Rows.Add(rowFark);
                        //rowFark[1] = (string.Format("{0:dd/MM/yyyy}", Ttarih1)) + " - " + (string.Format("{0:dd/MM/yyyy}", Ttarih2));
                        rowFark[2] = "Fark:";
                    }
                    */

                    //Excel Sayfasına 'Liste' yi ekle.
                    using (workbook)
                    {
                        workbook.Worksheets.Add(dt, "Liste");
                    }

                    /*
                    //Düzenleme - Paraları Sağa Yaslama
                    if (lblRapor.Text == "ALACAKLARIM")
                    {
                        workbook.Worksheet("Liste").Column(4).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
                        workbook.Worksheet("Liste").Column(5).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
                        workbook.Worksheet("Liste").Column(6).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
                    }
                    else if (lblRapor.Text == "BORÇLARIM")
                    {
                        workbook.Worksheet("Liste").Column(3).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
                        workbook.Worksheet("Liste").Column(4).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
                        workbook.Worksheet("Liste").Column(5).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;  
                    }
                    */
                    /*
                    for (int i = 1; i < dataGridView4.Rows.Count + 4; i++)
                    {
                        for (int j = 5; j < dataGridView4.Columns.Count + 2; j++)
                        {
                            workbook.Worksheet("Liste").Cell(i, j).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
                        }
                    }
                    */

                    //Düzenleme No Sütunu En Altlar Sil, Toplam Font Bold Yap ve Sağa Hizala, Başlıklar Bold, Siyah
                    workbook.Worksheet("Liste").Cell(dataGridView4.Rows.Count + 2, 1).Value = "";
                    workbook.Worksheet("Liste").Cell(dataGridView4.Rows.Count + 3, 1).Value = "";
                    workbook.Worksheet("Liste").Row(dataGridView4.Rows.Count + 3).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;

                    if (lblRapor.Text == "ALACAKLARIM")
                    {
                        for (int i = 4; i < dataGridView4.Columns.Count + 2; i++)
                        {
                            workbook.Worksheet("Liste").Column(i).Style.NumberFormat.NumberFormatId = 4;
                        }
                    }
                    else if (lblRapor.Text == "BORÇLARIM")
                    {
                        for (int i = 4; i < dataGridView4.Columns.Count + 2; i++)
                        {
                            workbook.Worksheet("Liste").Column(i).Style.NumberFormat.NumberFormatId = 4;
                        }
                    }
                    
                    /*
                    //Tarihler Arası Gelir-Gider Aktifse
                    if (ChkGelirGider.Checked == true)
                    {
                        GelirGider = 5;
                        var girdicikti = from kisiler in ctx.Kisis
                                         join cariler in ctx.Caris on kisiler.KisiID equals cariler.KisiID
                                         join durumlar in ctx.Durums on cariler.DurumID equals durumlar.DurumID

                                         group cariler by new { cariler.KisiID, kisiler.Ad, kisiler.Tel1 } into grup
                                         let Ad = grup.Key.Ad
                                         let Telefon = grup.Key.Tel1
                                         let Gider = grup.Where(x => x.DurumID == 5 && (x.Tarih.Value >= Ttarih1 && x.Tarih.Value <= Ttarih2)).Sum(x => (decimal?)x.Tutar) ?? 0
                                         let Ödeme = grup.Where(x => x.DurumID == 10 && (x.Tarih.Value >= Ttarih1 && x.Tarih.Value <= Ttarih2)).Sum(x => (decimal?)x.Tutar) ?? 0
                                         let Ödendi = grup.Where(x => x.DurumID == 4 && (x.Tarih.Value >= Ttarih1 && x.Tarih.Value <= Ttarih2)).Sum(x => (decimal?)x.Tutar) ?? 0
                                         let Alındı = grup.Where(x => x.DurumID == 2 && (x.Tarih.Value >= Ttarih1 && x.Tarih.Value <= Ttarih2)).Sum(x => (decimal?)x.Tutar) ?? 0

                                         orderby grup.Key.Ad
                                         select new
                                         {
                                             grup.Key.Ad,
                                             Telefon,
                                             Gider,
                                             Ödeme,
                                             Ödendi,
                                             Alındı
                                         };

                        var GiderToplam = girdicikti.Sum(x => x.Gider);
                        var OdemeToplam = girdicikti.Sum(x => x.Ödeme);
                        var OdendiToplam = girdicikti.Sum(x => x.Ödendi);
                        var AlindiToplam = girdicikti.Sum(x => x.Alındı);

                        workbook.Worksheet("Liste").Cell(dataGridView4.Rows.Count + 4, 1).Value = "";
                        workbook.Worksheet("Liste").Cell(dataGridView4.Rows.Count + 5, 1).Value = "";
                        workbook.Worksheet("Liste").Cell(dataGridView4.Rows.Count + 6, 1).Value = "";
                        workbook.Worksheet("Liste").Cell(dataGridView4.Rows.Count + 7, 1).Value = "";
                        workbook.Worksheet("Liste").Cell(dataGridView4.Rows.Count + 8, 1).Value = "";
                        workbook.Worksheet("Liste").Cell(dataGridView4.Rows.Count + 9, 1).Value = "";

                        workbook.Worksheet("Liste").Cell(dataGridView4.Rows.Count + 6, 4).Value = AlindiToplam;
                        workbook.Worksheet("Liste").Cell(dataGridView4.Rows.Count + 7, 4).Value = OdendiToplam + OdemeToplam;
                        workbook.Worksheet("Liste").Cell(dataGridView4.Rows.Count + 8, 4).Value = GiderToplam;
                        workbook.Worksheet("Liste").Cell(dataGridView4.Rows.Count + 9, 4).Value = AlindiToplam - (GiderToplam + OdendiToplam + OdemeToplam);

                        workbook.Worksheet("Liste").Cell(dataGridView4.Rows.Count + 5, 3).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
                        workbook.Worksheet("Liste").Cell(dataGridView4.Rows.Count + 6, 3).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
                        workbook.Worksheet("Liste").Cell(dataGridView4.Rows.Count + 7, 3).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
                        workbook.Worksheet("Liste").Cell(dataGridView4.Rows.Count + 8, 3).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
                        workbook.Worksheet("Liste").Cell(dataGridView4.Rows.Count + 9, 3).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
                    }
                    */

                    //Düzenleme No Sütunu En Altlar Sil, Toplam Font Bold Yap ve Sağa Hizala, Başlıklar Bold, Siyah-Devam
                    workbook.Worksheet("Liste").Cell(dataGridView4.Rows.Count + 3, 4).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
                    workbook.Worksheet("Liste").Row(dataGridView4.Rows.Count + 3).Style.Font.SetBold();
                    workbook.Worksheet("Liste").Row(1).Style.Fill.BackgroundColor = XLColor.White;
                    workbook.Worksheet("Liste").Row(1).Style.Font.SetBold();
                    workbook.Worksheet("Liste").Row(1).Style.Font.SetFontColor(XLColor.FromArgb(0, 0, 1));
                    workbook.Worksheet("Liste").Cell(1, 1).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
                    workbook.Worksheet("Liste").PageSetup.Header.Left.AddText(lblRapor.Text).SetBold();
                    workbook.Worksheet("Liste").PageSetup.Header.Right.AddText(string.Format("{0:dd/MM/yyyy}", DateTime.Now));
                    /*
                    //Satır Arkaplan Renkleri Ata
                    for (int i = 3; i < dataGridView4.Rows.Count + 2; i += 2)
                    {
                        workbook.Worksheet("Kale Mobilya").Row(i).Style.Fill.BackgroundColor = XLColor.LightGray;
                    }
                    
                    for (int i = 2; i < dataGridView4.Rows.Count + 4; i += 2)
                    {
                        workbook.Worksheet("Kale Mobilya").Row(i).Style.Fill.BackgroundColor = XLColor.White;
                    }
                    */

                    //Değerler tekrar atanarak sayı hale getiriliyor..
                    if (lblRapor.Text == "ALACAKLARIM")
                    {
                        for (int j = 4; j < dataGridView4.Columns.Count + 2; j++)
                        {
                            for (int i = 2; i < dataGridView4.Rows.Count + GelirGider + 4; i++)
                            {
                                workbook.Worksheet("Liste").Cell(i, j).Value = workbook.Worksheet("Liste").Cell(i, j).Value;
                            }
                        }
                    }
                    else if (lblRapor.Text == "BORÇLARIM")
                    {
                        for (int j = 4; j < dataGridView4.Columns.Count + 2; j++)
                        {
                            for (int i = 2; i < dataGridView4.Rows.Count + GelirGider + 4; i++)
                            {
                                workbook.Worksheet("Liste").Cell(i, j).Value = workbook.Worksheet("Liste").Cell(i, j).Value;
                            }
                        }
                    }
                    
                    //Liste A SON //

                    //<-ALACAKLARIM (Liste - B, Liste - 3)->//
                    if (lblRapor.Text == "ALACAKLARIM")
                    {
                        // Liste - 3   //

                        //DataTable3 Oluştur
                        DataTable dt3 = new DataTable();

                        //No3 Sütunu Ekle
                        DataColumn columno3 = new DataColumn
                        {
                            DataType = System.Type.GetType("System.Int32"),
                            AutoIncrement = true,
                            AutoIncrementSeed = 1,
                            AutoIncrementStep = 1
                        };
                        dt3.Columns.Add(columno3);
                        dt3.Columns["Column1"].ColumnName = "No";

                        //Sütunları Ekle
                        foreach (DataGridViewColumn column in dataGridView4.Columns)
                        {
                            dt3.Columns.Add(column.HeaderText);
                        }

                        if (dt3.Columns.Contains("Kalan_Alacağım"))
                        {
                            dt3.Columns["Kalan_Alacağım"].ColumnName = "Alacağım";
                        }

                        //Satırları Ekle
                        /*
                        foreach (DataGridViewRow row in dataGridView4.Rows)
                        {
                            dt3.Rows.Add();
                            //Hücreleri Ekle
                            foreach (DataGridViewCell cell in row.Cells)
                            {
                                if (cell.Value != null)
                                {
                                    dt3.Rows[dt3.Rows.Count - 1][cell.ColumnIndex + 1] = cell.Value;
                                }
                            }
                        }
                        */
                        foreach (DataGridViewRow row in dataGridView4.Rows)
                        {
                            dt3.Rows.Add();
                        }
                        var a1 = dataGridView4.Rows[0].Cells[0].Value;
                        var a2 = dataGridView4.Rows[0].Cells[1].Value;
                        var a22 = dataGridView4.Rows[0].Cells[2].Value;
                        var a23 = dataGridView4.Rows[0].Cells[3].Value;
                        var a24 = dataGridView4.Rows[3].Cells[0].Value;
                        var a3 = dataGridView4.Rows[1].Cells[1].Value;
                        var a4 = dataGridView4.Rows[1].Cells[2].Value;
                        var a5 = dataGridView4.Rows[1].Cells[3].Value;
                        var a6 = dataGridView4.Rows[1].Cells[4].Value;
                        var a7 = dataGridView4.Rows[1].Cells[5].Value;
                        var a8 = dataGridView4.Rows[2].Cells[6].Value;
                        for (int i = 0; i < dataGridView4.Rows.Count; i++)
                        {                            
                                for (int j = 0; j < dataGridView4.Columns.Count; j++)
                                {
                                if (j == 2)
                                    {
                                        dt3.Rows[i][j + 1] = string.Format("{0:dd/MM/yyyy}", dataGridView4.Rows[i].Cells[j].Value);
                                    }
                                    else
                                    {
                                        dt3.Rows[i][j + 1] = dataGridView4.Rows[i].Cells[j].Value;
                                    }
                                }
                        }

                        //Toplam Satırı Ekle
                        DataRow rowToplam3 = dt3.NewRow();
                        dt3.Rows.Add();
                        dt3.Rows.Add(rowToplam3);
                        rowToplam3[6] = "Toplam :";

                        for (int i = 6; i < dataGridView4.Columns.Count; i++)
                        {
                            decimal toplam = 0;
                            for (int j = 0; j < dataGridView4.Rows.Count + 1; ++j)
                            {
                                if (j < dataGridView4.Rows.Count)
                                {
                                    toplam += Convert.ToDecimal(dataGridView4.Rows[j].Cells[i].Value);
                                }
                                else
                                {
                                    rowToplam3[i + 1] = toplam;
                                }
                            }
                        }

                        //Gereksiz Sütunu Kaldır
                        dt3.Columns.RemoveAt(1);
                        dt3.Columns.RemoveAt(1);
                        dt3.Columns.RemoveAt(3);

                        //Excel Sayfasına 'Liste-3' yi ekle.
                        using (workbook)
                        {
                            workbook.Worksheets.Add(dt3, "Liste-3");
                        }

                        //Düzenleme No Sütunu En Altlar Sil, Toplam Font Bold Yap ve Sağa Hizala, Başlıklar Bold, Siyah
                        workbook.Worksheet("Liste-3").Cell(dataGridView4.Rows.Count + 2, 1).Value = "";
                        workbook.Worksheet("Liste-3").Cell(dataGridView4.Rows.Count + 3, 1).Value = "";
                        workbook.Worksheet("Liste-3").Row(dataGridView4.Rows.Count + 3).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;


                        for (int i = 5; i < dataGridView4.Columns.Count + 2; i++)
                        {
                            workbook.Worksheet("Liste-3").Column(i).Style.NumberFormat.NumberFormatId = 4;
                        }


                        //Düzenleme No Sütunu En Altlar Sil, Toplam Font Bold Yap ve Sağa Hizala, Başlıklar Bold, Siyah-Devam
                        workbook.Worksheet("Liste-3").Cell(dataGridView4.Rows.Count + 3, 4).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
                        workbook.Worksheet("Liste-3").Row(dataGridView4.Rows.Count + 3).Style.Font.SetBold();
                        workbook.Worksheet("Liste-3").Row(1).Style.Fill.BackgroundColor = XLColor.White;
                        workbook.Worksheet("Liste-3").Row(1).Style.Font.SetBold();
                        workbook.Worksheet("Liste-3").Row(1).Style.Font.SetFontColor(XLColor.FromArgb(0, 0, 1));
                        workbook.Worksheet("Liste-3").Cell(1, 1).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
                        workbook.Worksheet("Liste-3").PageSetup.Header.Left.AddText(lblRapor.Text).SetBold();
                        workbook.Worksheet("Liste-3").PageSetup.Header.Right.AddText(string.Format("{0:dd/MM/yyyy}", DateTime.Now));
                        workbook.Worksheet("Liste-3").Column(2).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;


                        //Değerler tekrar atanarak sayı hale getiriliyor..
                        for (int j = 5; j < dataGridView4.Columns.Count + 2; j++)
                        {
                            for (int i = 2; i < dataGridView4.Rows.Count + GelirGider + 4; i++)
                            {
                                workbook.Worksheet("Liste-3").Cell(i, j).Value = workbook.Worksheet("Liste-3").Cell(i, j).Value;
                            }
                        }

                        //Liste-3 SON //

                        // 3 Ayın Altındakiler //
                        //<-Hesaplamalar Evvelen1!-> Başlangıç//
                        lblRapor.Text = "ALACAKLARIM";
                        //tabControl1.TabPages.Add(tabPageHesaplama);
                        dataGridView4.DataSource = null;

                        //Hesaplama - Alacağım1
                        var alacagim1 = from kisiler in ctx.Kisis
                                        join cariler in ctx.Caris on kisiler.KisiID equals cariler.KisiID
                                        join durumlar in ctx.Durums on cariler.DurumID equals durumlar.DurumID

                                        group cariler by new { cariler.KisiID, kisiler.Karaliste, kisiler.Ad, kisiler.Tel1 } into grup
                                        let KisiID = grup.Key.KisiID
                                        let Karaliste = grup.Key.Karaliste
                                        let Ad = grup.Key.Ad
                                        let Telefon = grup.Key.Tel1
                                        let SonTarih = grup.Where(x => x.DurumID == 2).Max(x => x.Tarih)
                                        //let SonTarih2 = grup.Where(x => x.DurumID == 1).Min(x => x.Tarih)
                                        let Bakiye = grup.Where(x => x.DurumID == 1).Sum(x => (decimal?)x.Tutar) ?? 0
                                        let Alındı = grup.Where(x => x.DurumID == 2).Sum(x => (decimal?)x.Tutar) ?? 0
                                        let iskonto = grup.Where(x => x.DurumID == 9).Sum(x => (decimal?)x.Tutar) ?? 0
                                        let Alacağım = (Bakiye - (Alındı + iskonto))

                                        orderby grup.Key.Ad
                                        select new
                                        {
                                            grup.Key.KisiID,
                                            grup.Key.Karaliste,
                                            SonTarih,
                                            //SonTarih2,
                                            grup.Key.Ad,
                                            Telefon,
                                            Bakiye,
                                            Alındı,
                                            Alacağım
                                        };


                        //dataGridView4.DataSource = alacagim1.Where(x => x.Alacağım != 0 && x.Karaliste == false && ((DateTime.Today - x.SonTarih).Value.TotalDays < 61) || ((x.Alındı == null && (DateTime.Today - x.SonTarih2).Value.TotalDays < 61) || (x.Alındı == 0 && (DateTime.Today - x.SonTarih2).Value.TotalDays < 61)));
                        dataGridView4.DataSource = alacagim1.Where(x => x.Alacağım != 0 && x.Karaliste == false && (DateTime.Today - x.SonTarih).Value.TotalDays < 93);

                        //int tarihh = DateTime.Today.Day - new DateTime(2019, 11, 30).Day;
                        /*
                        dataGridView4.Columns["KisiID"].Visible = dataGridView4.Columns["SonTarih"].Visible = dataGridView4.Columns["Karaliste"].Visible = false;

                        dataGridView4.Columns[3].HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleLeft;
                        dataGridView4.Columns[4].HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
                        dataGridView4.Columns[4].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                        dataGridView4.Columns[5].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                        dataGridView4.Columns[6].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                        dataGridView4.Columns[7].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;

                        dataGridView4.Columns["Ad"].Width = 220;
                        dataGridView4.Columns["Telefon"].Width = 115;
                        dataGridView4.Columns["Bakiye"].Width = 160;
                        dataGridView4.Columns["Alındı"].Width = 133;
                        dataGridView4.Columns["Alacağım"].Width = 160;
                        */
                        //<-Hesaplamalar1 Evvelen!-> SON//

                        //DataTable1 Oluştur
                        DataTable dt1 = new DataTable();

                        //No1 Sütunu Ekle
                        DataColumn columno1 = new DataColumn
                        {
                            DataType = System.Type.GetType("System.Int32"),
                            AutoIncrement = true,
                            AutoIncrementSeed = 1,
                            AutoIncrementStep = 1
                        };
                        dt1.Columns.Add(columno1);
                        dt1.Columns["Column1"].ColumnName = "No";

                        //Sütunları Ekle
                        foreach (DataGridViewColumn column in dataGridView4.Columns)
                        {
                            dt1.Columns.Add(column.HeaderText);
                        }

                        //Sütun İsimleri Düzenle
                        if (dt1.Columns.Contains("Kalan_Alacağım"))
                        {
                            dt1.Columns["Kalan_Alacağım"].ColumnName = "Alacağım";
                        }
                        if (dt1.Columns.Contains("Kalan_Borcum"))
                        {
                            dt1.Columns["Kalan_Borcum"].ColumnName = "Borcum";
                        }

                        int Liste1Satir1 = 0;
                        int Liste1Satir2 = 0;
                        int Liste1Satir3 = 0;
                        int Liste1Satir4 = 0;
                        int Liste1BoslukSatir1 = 0;
                        int Liste1BoslukSatir2 = 0;
                        int Liste1BoslukSatir3 = 0;
                        int Liste1BoslukSatir4 = 0;
                        decimal AToplam11 = 0;
                        decimal AToplam12 = 0;
                        decimal AToplam13 = 0;

                        if (dataGridView4.Rows.Count > 0)
                        {
                            Liste1BoslukSatir1 += 2;
                            //Satırları Ekle
                            foreach (DataGridViewRow row in dataGridView4.Rows)
                            {
                                dt1.Rows.Add();
                                Liste1Satir1 += 1;
                                //Hücreleri Ekle
                                foreach (DataGridViewCell cell in row.Cells)
                                {
                                    if (cell.Value != null)
                                    {
                                        dt1.Rows[dt1.Rows.Count - 1][cell.ColumnIndex + 1] = cell.Value;
                                    }
                                }
                            }
                            //Toplmaları Bul
                            for (int i = 0; i < dataGridView4.Rows.Count; i++)
                            {
                                AToplam11 += Convert.ToDecimal(dataGridView4.Rows[i].Cells[5].Value);
                            }
                            for (int i = 0; i < dataGridView4.Rows.Count; i++)
                            {
                                AToplam12 += Convert.ToDecimal(dataGridView4.Rows[i].Cells[6].Value);
                            }
                            for (int i = 0; i < dataGridView4.Rows.Count; i++)
                            {
                                AToplam13 += Convert.ToDecimal(dataGridView4.Rows[i].Cells[7].Value);
                            }
                        }
                        // 3 Ayın Altındakiler - SON //
                        //-//

                        // 3 Ayın Üstündekiler - BAŞLANGIÇ //
                        //<-Hesaplamalar Evvelen2!-> Başlangıç//
                        dataGridView4.DataSource = null;

                        //Hesaplama - Alacağım2
                        var alacagim2 = from kisiler in ctx.Kisis
                                        join cariler in ctx.Caris on kisiler.KisiID equals cariler.KisiID
                                        join durumlar in ctx.Durums on cariler.DurumID equals durumlar.DurumID

                                        group cariler by new { cariler.KisiID, kisiler.Karaliste, kisiler.Ad, kisiler.Tel1 } into grup
                                        let KisiID = grup.Key.KisiID
                                        let Karaliste = grup.Key.Karaliste
                                        let Ad = grup.Key.Ad
                                        let Telefon = grup.Key.Tel1
                                        let SonTarih = grup.Where(x => x.DurumID == 2).Max(x => x.Tarih)
                                        let Bakiye = grup.Where(x => x.DurumID == 1).Sum(x => (decimal?)x.Tutar) ?? 0
                                        let Alındı = grup.Where(x => x.DurumID == 2).Sum(x => (decimal?)x.Tutar) ?? 0
                                        let iskonto = grup.Where(x => x.DurumID == 9).Sum(x => (decimal?)x.Tutar) ?? 0
                                        let Alacağım = (Bakiye - (Alındı + iskonto))

                                        orderby grup.Key.Ad
                                        select new
                                        {
                                            grup.Key.KisiID,
                                            grup.Key.Karaliste,
                                            SonTarih,
                                            grup.Key.Ad,
                                            Telefon,
                                            Bakiye,
                                            Alındı,
                                            Alacağım
                                        };

                        dataGridView4.DataSource = alacagim2.Where(x => x.Alacağım != 0 && (DateTime.Today - x.SonTarih).Value.TotalDays >= 93 && (DateTime.Today - x.SonTarih).Value.TotalDays < 183 && x.Karaliste == false);
                        /*
                        dataGridView4.Columns["KisiID"].Visible = dataGridView4.Columns["SonTarih"].Visible = dataGridView4.Columns["Karaliste"].Visible = false;

                        dataGridView4.Columns[3].HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleLeft;
                        dataGridView4.Columns[4].HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
                        dataGridView4.Columns[4].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                        dataGridView4.Columns[5].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                        dataGridView4.Columns[6].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                        dataGridView4.Columns[7].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;

                        dataGridView4.Columns["Ad"].Width = 220;
                        dataGridView4.Columns["Telefon"].Width = 115;
                        dataGridView4.Columns["Bakiye"].Width = 160;
                        dataGridView4.Columns["Alındı"].Width = 133;
                        dataGridView4.Columns["Alacağım"].Width = 160;
                        */
                        //<-Hesaplamalar2 Evvelen!-> SON//
                        decimal AToplam21 = 0;
                        decimal AToplam22 = 0;
                        decimal AToplam23 = 0;
                        if (dataGridView4.Rows.Count > 0)
                        {
                            Liste1BoslukSatir2 += 2;
                            //Boş Satır Ekle
                            dt1.Rows.Add();
                            //Liste1Satir1 += 1;
                            //2 Ayın Üstündekiler Başlık Ekle
                            DataRow rowBaslik2 = dt1.NewRow();
                            dt1.Rows.Add(rowBaslik2);
                            rowBaslik2[4] = "3-6 Aydır Ödeme Yok";
                            //Liste1Satir2 += 1;
                            //Satırları Ekle
                            foreach (DataGridViewRow row in dataGridView4.Rows)
                            {
                                dt1.Rows.Add();
                                Liste1Satir2 += 1;
                                //Hücreleri Ekle
                                foreach (DataGridViewCell cell in row.Cells)
                                {
                                    if (cell.Value != null)
                                    {
                                        dt1.Rows[dt1.Rows.Count - 1][cell.ColumnIndex + 1] = cell.Value;
                                    }
                                }
                            }
                            //Toplmaları Bul
                            for (int i = 0; i < dataGridView4.Rows.Count; i++)
                            {
                                AToplam21 += Convert.ToDecimal(dataGridView4.Rows[i].Cells[5].Value);
                            }
                            for (int i = 0; i < dataGridView4.Rows.Count; i++)
                            {
                                AToplam22 += Convert.ToDecimal(dataGridView4.Rows[i].Cells[6].Value);
                            }
                            for (int i = 0; i < dataGridView4.Rows.Count; i++)
                            {
                                AToplam23 += Convert.ToDecimal(dataGridView4.Rows[i].Cells[7].Value);
                            }
                        }
                        // 3 Ayın Üstündekiler - SON //

                        // 6 Ayın Üstündekiler - BAŞLANGIÇ //
                        //<-Hesaplamalar Evvelen3!-> Başlangıç//
                        dataGridView4.DataSource = null;

                        //Hesaplama - Alacağım3
                        var alacagim3 = from kisiler in ctx.Kisis
                                        join cariler in ctx.Caris on kisiler.KisiID equals cariler.KisiID
                                        join durumlar in ctx.Durums on cariler.DurumID equals durumlar.DurumID

                                        group cariler by new { cariler.KisiID, kisiler.Karaliste, kisiler.Ad, kisiler.Tel1 } into grup
                                        let KisiID = grup.Key.KisiID
                                        let Karaliste = grup.Key.Karaliste
                                        let Ad = grup.Key.Ad
                                        let Telefon = grup.Key.Tel1
                                        let SonTarih = grup.Where(x => x.DurumID == 2).Max(x => x.Tarih)
                                        let Bakiye = grup.Where(x => x.DurumID == 1).Sum(x => (decimal?)x.Tutar) ?? 0
                                        let Alındı = grup.Where(x => x.DurumID == 2).Sum(x => (decimal?)x.Tutar) ?? 0
                                        let iskonto = grup.Where(x => x.DurumID == 9).Sum(x => (decimal?)x.Tutar) ?? 0
                                        let Alacağım = (Bakiye - (Alındı + iskonto))

                                        orderby grup.Key.Ad
                                        select new
                                        {
                                            grup.Key.KisiID,
                                            grup.Key.Karaliste,
                                            SonTarih,
                                            grup.Key.Ad,
                                            Telefon,
                                            Bakiye,
                                            Alındı,
                                            Alacağım
                                        };

                        dataGridView4.DataSource = alacagim3.Where(x => x.Alacağım != 0 && (DateTime.Today - x.SonTarih).Value.TotalDays >= 183 && x.Karaliste == false);
                        /*
                        dataGridView4.Columns["KisiID"].Visible = dataGridView4.Columns["SonTarih"].Visible = dataGridView4.Columns["Karaliste"].Visible = false;

                        dataGridView4.Columns[3].HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleLeft;
                        dataGridView4.Columns[4].HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
                        dataGridView4.Columns[4].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                        dataGridView4.Columns[5].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                        dataGridView4.Columns[6].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                        dataGridView4.Columns[7].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;

                        dataGridView4.Columns["Ad"].Width = 220;
                        dataGridView4.Columns["Telefon"].Width = 115;
                        dataGridView4.Columns["Bakiye"].Width = 160;
                        dataGridView4.Columns["Alındı"].Width = 133;
                        dataGridView4.Columns["Alacağım"].Width = 160;
                        */
                        //<-Hesaplamalar3 Evvelen!-> SON//
                        decimal AToplam31 = 0;
                        decimal AToplam32 = 0;
                        decimal AToplam33 = 0;
                        if (dataGridView4.Rows.Count > 0)
                        {
                            Liste1BoslukSatir3 += 2;
                            //Boş Satır Ekle
                            dt1.Rows.Add();
                            //Liste1Satir2 += 1;
                            //6 Ayın Üstündekiler Başlık Ekle
                            DataRow rowBaslik3 = dt1.NewRow();
                            dt1.Rows.Add(rowBaslik3);
                            rowBaslik3[4] = "6 Ay Geçti, Ödeme Yok";
                            //Liste1Satir3 += 1;
                            //Satırları Ekle
                            foreach (DataGridViewRow row in dataGridView4.Rows)
                            {
                                dt1.Rows.Add();
                                Liste1Satir3 += 1;
                                //Hücreleri Ekle
                                foreach (DataGridViewCell cell in row.Cells)
                                {
                                    if (cell.Value != null)
                                    {
                                        dt1.Rows[dt1.Rows.Count - 1][cell.ColumnIndex + 1] = cell.Value;
                                    }
                                }
                            }
                            //Toplmaları Bul
                            for (int i = 0; i < dataGridView4.Rows.Count; i++)
                            {
                                AToplam31 += Convert.ToDecimal(dataGridView4.Rows[i].Cells[5].Value);
                            }
                            for (int i = 0; i < dataGridView4.Rows.Count; i++)
                            {
                                AToplam32 += Convert.ToDecimal(dataGridView4.Rows[i].Cells[6].Value);
                            }
                            for (int i = 0; i < dataGridView4.Rows.Count; i++)
                            {
                                AToplam33 += Convert.ToDecimal(dataGridView4.Rows[i].Cells[7].Value);
                            }
                        }
                        // 6 Ayın Üstündekiler - SON //

                        // KırmızıListedekiler - BAŞLANGIÇ //
                        //<-Hesaplamalar Evvelen4!-> Başlangıç//
                        dataGridView4.DataSource = null;

                        //Hesaplama - Alacağım4
                        var alacagim4 = from kisiler in ctx.Kisis
                                        join cariler in ctx.Caris on kisiler.KisiID equals cariler.KisiID
                                        join durumlar in ctx.Durums on cariler.DurumID equals durumlar.DurumID

                                        group cariler by new { cariler.KisiID, kisiler.Karaliste, kisiler.Ad, kisiler.Tel1 } into grup
                                        let KisiID = grup.Key.KisiID
                                        let Karaliste = grup.Key.Karaliste
                                        let Ad = grup.Key.Ad
                                        let Telefon = grup.Key.Tel1
                                        let SonTarih = grup.Where(x => x.DurumID == 2).Max(x => x.Tarih)
                                        let Bakiye = grup.Where(x => x.DurumID == 1).Sum(x => (decimal?)x.Tutar) ?? 0
                                        let Alındı = grup.Where(x => x.DurumID == 2).Sum(x => (decimal?)x.Tutar) ?? 0
                                        let iskonto = grup.Where(x => x.DurumID == 9).Sum(x => (decimal?)x.Tutar) ?? 0
                                        let Alacağım = (Bakiye - (Alındı + iskonto))

                                        orderby grup.Key.Ad
                                        select new
                                        {
                                            grup.Key.KisiID,
                                            grup.Key.Karaliste,
                                            SonTarih,
                                            grup.Key.Ad,
                                            Telefon,
                                            Bakiye,
                                            Alındı,
                                            Alacağım
                                        };

                        dataGridView4.DataSource = alacagim4.Where(x => x.Alacağım != 0 && x.Karaliste == true);
                        /*
                        dataGridView4.Columns["KisiID"].Visible = dataGridView4.Columns["SonTarih"].Visible = dataGridView4.Columns["Karaliste"].Visible = false;

                        dataGridView4.Columns[3].HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleLeft;
                        dataGridView4.Columns[4].HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
                        dataGridView4.Columns[4].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                        dataGridView4.Columns[5].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                        dataGridView4.Columns[6].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                        dataGridView4.Columns[7].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;

                        dataGridView4.Columns["Ad"].Width = 220;
                        dataGridView4.Columns["Telefon"].Width = 115;
                        dataGridView4.Columns["Bakiye"].Width = 160;
                        dataGridView4.Columns["Alındı"].Width = 133;
                        dataGridView4.Columns["Alacağım"].Width = 160;
                        */
                        //<-Hesaplamalar4 Evvelen!-> SON//
                        decimal AToplam41 = 0;
                        decimal AToplam42 = 0;
                        decimal AToplam43 = 0;
                        if (dataGridView4.Rows.Count > 0)
                        {
                            Liste1BoslukSatir4 += 2;
                            //Boş Satır Ekle
                            dt1.Rows.Add();
                            //Liste1Satir3 += 1;
                            //Kırmızı Liste Başlık Ekle
                            DataRow rowBaslik4 = dt1.NewRow();
                            dt1.Rows.Add(rowBaslik4);
                            rowBaslik4[4] = "Kırmızı Liste";
                            //Liste1Satir4 += 1;
                            //Satırları Ekle
                            foreach (DataGridViewRow row in dataGridView4.Rows)
                            {
                                dt1.Rows.Add();
                                Liste1Satir4 += 1;
                                //Hücreleri Ekle
                                foreach (DataGridViewCell cell in row.Cells)
                                {
                                    if (cell.Value != null)
                                    {
                                        dt1.Rows[dt1.Rows.Count - 1][cell.ColumnIndex + 1] = cell.Value;
                                    }
                                }
                            }
                            //Toplmaları Bul
                            for (int i = 0; i < dataGridView4.Rows.Count; i++)
                            {
                                AToplam41 += Convert.ToDecimal(dataGridView4.Rows[i].Cells[5].Value);
                            }
                            for (int i = 0; i < dataGridView4.Rows.Count; i++)
                            {
                                AToplam42 += Convert.ToDecimal(dataGridView4.Rows[i].Cells[6].Value);
                            }
                            for (int i = 0; i < dataGridView4.Rows.Count; i++)
                            {
                                AToplam43 += Convert.ToDecimal(dataGridView4.Rows[i].Cells[7].Value);
                            }
                        }
                        // KırmızıListedekiler - SON //

                        //tabControl1.TabPages.Remove(tabPageHesaplama);
                        //Toplam Satırı Ekle
                        DataRow rowToplam1 = dt1.NewRow();
                        dt1.Rows.Add();
                        //Liste1Satir4 += 1;
                        //int Liste1SatirToplam = Liste1Satir1 + Liste1Satir2 + Liste1Satir3 + Liste1Satir4;
                        dt1.Rows.Add(rowToplam1);
                        rowToplam1[5] = "Toplam :";
                        rowToplam1[6] = AToplam11 + AToplam21 + AToplam31 + AToplam41;
                        rowToplam1[7] = AToplam12 + AToplam22 + AToplam32 + AToplam42;
                        rowToplam1[8] = AToplam13 + AToplam23 + AToplam33 + AToplam43;
                        //Liste1SatirToplam += 1;
                        //Gereksiz Sütunu Kaldır
                        dt1.Columns.RemoveAt(1);
                        dt1.Columns.RemoveAt(1);
                        dt1.Columns.RemoveAt(1);

                        //Excel Sayfasına 'Kale Mobilya' yı ekle.
                        using (workbook)
                        {
                            workbook.Worksheets.Add(dt1, "Liste-B");
                        }

                        //Düzenleme - Paraları Sağa Yaslama
                        workbook.Worksheet("Liste-B").Column(4).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
                        workbook.Worksheet("Liste-B").Column(5).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
                        workbook.Worksheet("Liste-B").Column(6).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;


                        //No lar sil
                        if (Liste1Satir1 != 0)
                        {
                            workbook.Worksheet("Liste-B").Cell(Liste1Satir1 + Liste1BoslukSatir1, 1).Value = "";
                            workbook.Worksheet("Liste-B").Cell(Liste1Satir1 + Liste1BoslukSatir1 + 1, 1).Value = "";
                        }

                        if (Liste1Satir2 != 0)
                        {
                            workbook.Worksheet("Liste-B").Cell(Liste1Satir1 + Liste1Satir2 + Liste1BoslukSatir1 + Liste1BoslukSatir2, 1).Value = "";
                            workbook.Worksheet("Liste-B").Cell(Liste1Satir1 + Liste1Satir2 + Liste1BoslukSatir1 + Liste1BoslukSatir2 + 1, 1).Value = "";
                        }

                        if (Liste1Satir3 != 0)
                        {
                            workbook.Worksheet("Liste-B").Cell(Liste1Satir1 + Liste1Satir2 + Liste1Satir3 + Liste1BoslukSatir1 + Liste1BoslukSatir2 + Liste1BoslukSatir3, 1).Value = "";
                            workbook.Worksheet("Liste-B").Cell(Liste1Satir1 + Liste1Satir2 + Liste1Satir3 + Liste1BoslukSatir1 + Liste1BoslukSatir2 + Liste1BoslukSatir3 + 1, 1).Value = "";
                        }

                        if (Liste1Satir4 != 0)
                        {
                            workbook.Worksheet("Liste-B").Cell(Liste1Satir1 + Liste1Satir2 + Liste1Satir3 + Liste1Satir4 + Liste1BoslukSatir1 + Liste1BoslukSatir2 + Liste1BoslukSatir3 + Liste1BoslukSatir4, 1).Value = "";
                            workbook.Worksheet("Liste-B").Cell(Liste1Satir1 + Liste1Satir2 + Liste1Satir3 + Liste1Satir4 + Liste1BoslukSatir1 + Liste1BoslukSatir2 + Liste1BoslukSatir3 + Liste1BoslukSatir4 + 1, 1).Value = "";
                        }


                        for (int i = 4; i < dataGridView4.Columns.Count - 1; i++)
                        {
                            workbook.Worksheet("Liste-B").Column(i).Style.NumberFormat.NumberFormatId = 4;
                        }

                        workbook.Worksheet("Liste-B").Row(Liste1Satir1 + Liste1Satir2 + Liste1Satir3 + Liste1Satir4 + Liste1BoslukSatir1 + Liste1BoslukSatir2 + Liste1BoslukSatir3 + Liste1BoslukSatir4+1).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
                        //workbook.Worksheet("Liste-B").Row(Liste1Satir1 + Liste1Satir2 + Liste1Satir3 + Liste1Satir4 + Liste1BoslukSatir1 + Liste1BoslukSatir2 + Liste1BoslukSatir3 + Liste1BoslukSatir4).Style.Font.SetBold();
                        workbook.Worksheet("Liste-B").Row(1).Style.Fill.BackgroundColor = XLColor.White;
                        workbook.Worksheet("Liste-B").Row(1).Style.Font.SetBold();

                        if (Liste1Satir1 != 0)
                        {
                            workbook.Worksheet("Liste-B").Row(Liste1Satir1 + Liste1BoslukSatir1 + 1).Style.Font.SetBold();
                        }
                        if (Liste1Satir2 != 0)
                        {
                            workbook.Worksheet("Liste-B").Row(Liste1Satir1 + Liste1Satir2 + Liste1BoslukSatir1 + Liste1BoslukSatir2 + 1).Style.Font.SetBold();
                        }
                        if (Liste1Satir3 != 0)
                        {
                            workbook.Worksheet("Liste-B").Row(Liste1Satir1 + Liste1Satir2 + Liste1Satir3 + Liste1BoslukSatir1 + Liste1BoslukSatir2 + Liste1BoslukSatir3 + 1).Style.Font.SetBold();
                        }
                        if (Liste1Satir4 != 0)
                        {
                            workbook.Worksheet("Liste-B").Row(Liste1Satir1 + Liste1Satir2 + Liste1Satir3 + Liste1Satir4 + Liste1BoslukSatir1 + Liste1BoslukSatir2 + Liste1BoslukSatir3 + Liste1BoslukSatir4 + 1).Style.Font.SetBold();
                        }

                        workbook.Worksheet("Liste-B").Row(1).Style.Font.SetFontColor(XLColor.FromArgb(0, 0, 1));
                        workbook.Worksheet("Liste-B").Cell(1, 1).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
                        workbook.Worksheet("Liste-B").PageSetup.Header.Left.AddText(lblRapor.Text).SetBold();
                        workbook.Worksheet("Liste-B").PageSetup.Header.Right.AddText(string.Format("{0:dd/MM/yyyy}", DateTime.Now));

                        //Değerler tekrar atanarak sayı hale getiriliyor..
                        for (int j = 4; j < 7; j++)
                        {
                            for (int i = 2; i < Liste1Satir1 + Liste1Satir2 + Liste1Satir3 + Liste1Satir4 + Liste1BoslukSatir1 + Liste1BoslukSatir2 + Liste1BoslukSatir3 + Liste1BoslukSatir4 + 2; i++)
                            {
                                workbook.Worksheet("Liste-B").Cell(i, j).Value = workbook.Worksheet("Liste-B").Cell(i, j).Value;
                            }
                        }

                        // No lar düzenleniyor
                        if (Liste1Satir2 != 0)
                        {
                            for (int i = 1; i < Liste1Satir2+1; i++)
                            {
                                workbook.Worksheet("Liste-B").Cell(Liste1Satir1 + Liste1BoslukSatir1 + 1 + i, 1).Value = Liste1Satir1 + i;
                            }
                        }

                        if (Liste1Satir3 != 0)
                        {
                            for (int i = 1; i < Liste1Satir3+1; i++)
                            {
                                workbook.Worksheet("Liste-B").Cell(Liste1Satir1 + Liste1Satir2 + Liste1BoslukSatir1 + Liste1BoslukSatir2 + 1 + i, 1).Value = Liste1Satir1 + Liste1Satir2 + i;
                            }
                        }

                        if (Liste1Satir4 != 0)
                        {
                            for (int i = 1; i < Liste1Satir4+1; i++)
                            {
                                workbook.Worksheet("Liste-B").Cell(Liste1Satir1 + Liste1Satir2 + Liste1Satir3 + Liste1BoslukSatir1 + Liste1BoslukSatir2 + Liste1BoslukSatir3 + 1 + i, 1).Value = Liste1Satir1 + Liste1Satir2 + Liste1Satir3 + i;
                            }
                        }

                        /*
                        int r1 = Liste1Satir1 + 2;
                        string hucrea = "A" + r1;
                        workbook.Worksheet("Liste-B").Cell(hucrea).Value = "2 - 6 Aydır Ödeme Yapmayanlar";
                        */
                        /*
                        // Hücre Birleştir
                        int r1 = Liste1Satir1 + 2;
                        string hucre1 = "A" + r1 + ":" + "B" + r1;
                        string hucre2 = "A" + r1;                       
                        workbook.Worksheet("Liste-B").Range(hucre1).Row(1).Merge();
                        workbook.Worksheet("Liste-B").Cell(hucre2).Value = "2 - 6 Aydır Ödeme Yapmayanlar";
                        */

                        //Alacaklarımı son hâline getir.
                        var alacagim = from kisiler in ctx.Kisis
                                       join cariler in ctx.Caris on kisiler.KisiID equals cariler.KisiID
                                       join durumlar in ctx.Durums on cariler.DurumID equals durumlar.DurumID

                                       group cariler by new { cariler.KisiID, kisiler.Karaliste, kisiler.Ad, kisiler.Firma, kisiler.Tel1 } into grup
                                       let KisiID = grup.Key.KisiID
                                       let Karaliste = grup.Key.Karaliste
                                       let Ad = grup.Key.Ad
                                       let Firma = grup.Key.Firma
                                       let Telefon = grup.Key.Tel1
                                       let SonTarih = grup.Where(x => x.DurumID == 2).Max(x => x.Tarih)
                                       let Bakiye = grup.Where(x => x.DurumID == 1).Sum(x => (decimal?)x.Tutar) ?? 0
                                       let Alındı = grup.Where(x => x.DurumID == 2).Sum(x => (decimal?)x.Tutar) ?? 0
                                       let iskonto = grup.Where(x => x.DurumID == 9).Sum(x => (decimal?)x.Tutar) ?? 0
                                       let Alacağım = (Bakiye - (Alındı + iskonto))

                                       orderby grup.Key.Ad
                                       select new
                                       {
                                           grup.Key.KisiID,
                                           grup.Key.Karaliste,
                                           SonTarih,
                                           grup.Key.Ad,
                                           Firma,
                                           Telefon,
                                           Bakiye,
                                           Alındı,
                                           Alacağım
                                       };

                        dataGridView4.DataSource = alacagim.Where(x => x.Alacağım != 0);

                        dataGridView4.Columns["KisiID"].Visible = dataGridView4.Columns["Firma"].Visible = dataGridView4.Columns["Karaliste"].Visible = false;

                        dataGridView4.Columns[3].HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleLeft;
                        dataGridView4.Columns[3].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft;
                        dataGridView4.Columns[4].HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleLeft;
                        dataGridView4.Columns[5].HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
                        dataGridView4.Columns[5].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                        //dataGridView4.Columns[6].HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleRight;
                        dataGridView4.Columns[6].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                        //dataGridView4.Columns[7].HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleRight;
                        dataGridView4.Columns[7].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                        //dataGridView4.Columns[8].HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleRight;
                        dataGridView4.Columns[8].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;

                        dataGridView4.Columns["SonTarih"].Width = 100;
                        dataGridView4.Columns["SonTarih"].HeaderText = "Son İşlem";
                        dataGridView4.Columns["Ad"].Width = 220;
                        dataGridView4.Columns["Telefon"].Width = 115;
                        dataGridView4.Columns["Bakiye"].Width = 160;
                        dataGridView4.Columns["Alındı"].Width = 133;
                        dataGridView4.Columns["Alacağım"].Width = 160;
                        //Eski
                        /*
                        var alacagim = from kisiler in ctx.Kisis
                                       join cariler in ctx.Caris on kisiler.KisiID equals cariler.KisiID
                                       join durumlar in ctx.Durums on cariler.DurumID equals durumlar.DurumID

                                       group cariler by new { cariler.KisiID, kisiler.Karaliste, kisiler.Ad, kisiler.Tel1 } into grup
                                       let KisiID = grup.Key.KisiID
                                       let Karaliste = grup.Key.Karaliste
                                       let Ad = grup.Key.Ad
                                       let Telefon = grup.Key.Tel1
                                       let Bakiye = grup.Where(x => x.DurumID == 1).Sum(x => (decimal?)x.Tutar) ?? 0
                                       let Alındı = grup.Where(x => x.DurumID == 2).Sum(x => (decimal?)x.Tutar) ?? 0
                                       let iskonto = grup.Where(x => x.DurumID == 9).Sum(x => (decimal?)x.Tutar) ?? 0
                                       let Alacağım = (Bakiye - (Alındı + iskonto))

                                       orderby grup.Key.Ad
                                       select new
                                       {
                                           grup.Key.KisiID,
                                           grup.Key.Karaliste,
                                           grup.Key.Ad,
                                           Telefon,
                                           Bakiye,
                                           Alındı,
                                           Alacağım
                                       };

                        dataGridView4.DataSource = alacagim.Where(x => x.Alacağım != 0);
                        dataGridView4.Columns["KisiID"].Visible = dataGridView4.Columns["Karaliste"].Visible = false;

                        dataGridView4.Columns[2].HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleLeft;
                        dataGridView4.Columns[3].HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
                        dataGridView4.Columns[3].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                        dataGridView4.Columns[4].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                        dataGridView4.Columns[5].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                        dataGridView4.Columns[6].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;

                        dataGridView4.Columns["Ad"].Width = 220;
                        dataGridView4.Columns["Telefon"].Width = 115;
                        dataGridView4.Columns["Bakiye"].Width = 160;
                        dataGridView4.Columns["Alındı"].Width = 133;
                        dataGridView4.Columns["Alacağım"].Width = 160;
                        */
                    }
                    //<-Liste-B SON->//
                   
                    workbook.Worksheet("Liste").Columns().AdjustToContents();
                    workbook.Worksheet("Liste").PageSetup.FitToPages(1, 2);
                    //workbook.Worksheet("Kale Mobilya").PageSetup.AdjustTo(94);

                    //Liste-B
                    if (lblRapor.Text == "ALACAKLARIM")
                    {
                        workbook.Worksheet("Liste-B").Columns().AdjustToContents();
                        workbook.Worksheet("Liste-B").PageSetup.FitToPages(1, 2);
                        workbook.Worksheet("Liste-3").Columns().AdjustToContents();
                        workbook.Worksheet("Liste-3").PageSetup.FitToPages(1, 2);
                    }

                    //Kaydet
                    do
                    {
                        try
                        {
                            workbook.SaveAs(fileName);
                            MessageBox.Show("Excel dosyası kaydedildi.", "Bilgi", MessageBoxButtons.OK, MessageBoxIcon.Information);
                            customCulture.NumberFormat.NumberDecimalSeparator = ",";
                            System.Threading.Thread.CurrentThread.CurrentCulture = customCulture;
                            break;
                        }
                        catch (System.IO.IOException)
                        {
                            MessageBox.Show("Kayıt yapılamadı! Kaydetmeye çalıştığınız dosya açık olabilir.", "Uyarı", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                            customCulture.NumberFormat.NumberDecimalSeparator = ",";
                            System.Threading.Thread.CurrentThread.CurrentCulture = customCulture;
                            break;
                        }
                    } while (true);
                }
            }
            else
            {
                MessageBox.Show("Herhangi bir rapor seçili değil!", "Uyarı", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        private void BtnHesapla_Click(object sender, EventArgs e)
        {
            //Hesapla
            var genel = from kisiler in ctx.Kisis
                        join cariler in ctx.Caris on kisiler.KisiID equals cariler.KisiID
                        join durumlar in ctx.Durums on cariler.DurumID equals durumlar.DurumID

                        group cariler by new { cariler.KisiID, kisiler.Ad, kisiler.Tel1 } into grup
                        let Ad = grup.Key.Ad
                        let Telefon = grup.Key.Tel1
                        let Toplam_Alacağım = grup.Where(x => x.DurumID == 1).Sum(x => (decimal?)x.Tutar) ?? 0
                        let Toplam_Alındı = grup.Where(x => x.DurumID == 2).Sum(x => (decimal?)x.Tutar) ?? 0
                        let Iskonto = grup.Where(x => x.DurumID == 9).Sum(x => (decimal?)x.Tutar) ?? 0
                        let Kalan_Alacağım = (Toplam_Alacağım - (Toplam_Alındı + Iskonto))
                        let Toplam_Borç = grup.Where(x => x.DurumID == 3).Sum(x => (decimal?)x.Tutar) ?? 0
                        let Toplam_Ödendi = grup.Where(x => x.DurumID == 4).Sum(x => (decimal?)x.Tutar) ?? 0
                        let Kalan_Borcum = Toplam_Borç - Toplam_Ödendi

                        orderby grup.Key.Ad
                        select new
                        {
                            grup.Key.Ad,
                            Telefon,
                            Kalan_Alacağım,
                            Kalan_Borcum
                        };

            lblToplamAlacak.Text = String.Format("{0:N}\n", genel.Sum(toplam => toplam.Kalan_Alacağım));
            lblToplamBorc.Text = String.Format("{0:N}\n", genel.Sum(toplam => toplam.Kalan_Borcum));
        }

        private void BtnEEkle_Click(object sender, EventArgs e)
        {
            if (txtEAd.Text != "")
            {
                //Ekle
                Kisi per = new Kisi
                {
                    Ad = txtEAd.Text,
                    Tel1 = txtETel1.Text,
                    Tel2 = txtETel2.Text,
                    Adres = txtEAdres.Text,
                    Firma = txtEFirma.Text,
                    Karaliste = false
                };

                ctx.Kisis.InsertOnSubmit(per);
                ctx.SubmitChanges();

                //Listele
                var sonuc = from kisiler in ctx.Kisis
                            orderby kisiler.Ad
                            select new
                            {
                                kisiler.KisiID,
                                kisiler.Ad,
                                kisiler.Firma,
                                kisiler.Tel1,
                                kisiler.Tel2,
                                kisiler.Adres,
                                kisiler.Karaliste
                            };
                dataGridView1.DataSource = sonuc;
                dataGridView1.Columns["Firma"].Visible = dataGridView1.Columns["Tel1"].Visible = dataGridView1.Columns["Tel2"].Visible = dataGridView1.Columns["Adres"].Visible = dataGridView1.Columns["KisiID"].Visible = dataGridView1.Columns["Karaliste"].Visible = false;
                DataGridViewColumn column = dataGridView1.Columns[1];
                column.Width = dataGridView1.Width - 20;

                MessageBox.Show(txtEAd.Text + " adlı kişi eklendi!", "Bilgi", MessageBoxButtons.OK, MessageBoxIcon.Information);

                //Alanları temizle
                txtEAd.Text = null;
                txtETel1.Text = null;
                txtETel2.Text = null;
                txtEFirma.Text = null;
                txtEAdres.Text = null;
                txtAra.Text = null;  
            }
            else
            {
                MessageBox.Show("Lütfen 'Ad Soyad' giriniz!", "Uyarı", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        private void BtnGGuncelle_Click(object sender, EventArgs e)
        {
            if (txtGAd.Text != "")
            {
                Button btn = sender as Button;
                DialogResult sonuc = MessageBox.Show(lblSAd.Text + " adlı kişi güncellenecek?", "Güncelleme", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                //Güncelle
                if (sonuc == DialogResult.Yes)
                {
                    anahtar = false;

                    int id = (int)txtGAd.Tag;

                    Kisi k = ctx.Kisis.SingleOrDefault(x => x.KisiID == id);
                    k.Ad = txtGAd.Text;
                    k.Firma = txtGFirma.Text;
                    k.Tel1 = txtGTel1.Text;
                    k.Tel2 = txtGTel2.Text;
                    k.Adres = txtGAdres.Text;

                    if (ChkKaraListe.Checked)
                    {
                        k.Karaliste = true;
                    }
                    else
                    {
                        k.Karaliste = false;
                    }

                    ctx.SubmitChanges();

                    //Değişiklikleri solda listele   
                    var liste = from kisiler in ctx.Kisis
                                orderby kisiler.Ad
                                select new
                                {
                                    kisiler.KisiID,
                                    kisiler.Ad,
                                    kisiler.Firma,
                                    kisiler.Tel1,
                                    kisiler.Tel2,
                                    kisiler.Adres,
                                    kisiler.Karaliste
                                };
                    dataGridView1.DataSource = liste;
                    dataGridView1.Columns["Firma"].Visible = dataGridView1.Columns["Tel1"].Visible = dataGridView1.Columns["Tel2"].Visible = dataGridView1.Columns["Adres"].Visible = dataGridView1.Columns["KisiID"].Visible = dataGridView1.Columns["Karaliste"].Visible = false;
                    DataGridViewColumn column = dataGridView1.Columns[1];
                    column.Width = dataGridView1.Width - 20;

                    //Değişiklikleri altta düzenle bölümünde listele
                    var liste2 = from kisiler in ctx.Kisis
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

                    //Allta Bilgiler Bölümünde Listele
                    dataGridView2.DataSource = liste2.Where(x => x.KisiID == Convert.ToInt32(lblID.Tag));
                    dataGridView2.Columns["KisiID"].Visible = dataGridView2.Columns["Ad"].Visible = dataGridView2.Columns["DurumID"].Visible = dataGridView2.Columns["CariID"].Visible = false;

                    dataGridView2.Columns[3].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                    dataGridView2.Columns[3].HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleRight;
                    dataGridView2.Columns[4].HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
                    dataGridView2.Columns[4].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;

                    //Üstte Bilgiler Bölümünde Listele
                    DataGridViewRow row = dataGridView1.CurrentRow;
                    lblAd.Text = row.Cells["Ad"].Value.ToString();
                    lblDAd.Text = row.Cells["Ad"].Value.ToString();

                    LblKaraListe.Tag = row.Cells["Karaliste"].Value;
                    if ((Boolean)LblKaraListe.Tag)
                    {
                        LblKaraListe.ForeColor = Color.DarkRed;
                    }
                    else
                    {
                        LblKaraListe.ForeColor = Color.White;
                    }

                    if ((string)row.Cells["Firma"].Value == "")
                    {
                        label2.Visible = false;
                        lblFirma.Visible = false;
                    }
                    else
                    {
                        label2.Visible = true;
                        lblFirma.Visible = true;
                        lblFirma.Text = row.Cells["Firma"].Value.ToString();
                    }

                    if ((string)row.Cells["Tel1"].Value == "")
                    {
                        label4.Visible = false;
                        lblTel1.Visible = false;
                    }
                    else
                    {
                        label4.Visible = true;
                        lblTel1.Visible = true;
                        lblTel1.Text = row.Cells["Tel1"].Value.ToString();
                    }

                    if ((string)row.Cells["Tel2"].Value == "")
                    {
                        lblTel2B.Visible = false;
                        lblTel2.Visible = false;
                    }
                    else
                    {
                        lblTel2B.Visible = true;
                        lblTel2.Visible = true;
                        lblTel2.Text = row.Cells["Tel2"].Value.ToString();
                    }

                    if ((string)row.Cells["Adres"].Value == "")
                    {
                        lblAdresB.Visible = false;
                        lblAdres.Visible = false;
                    }
                    else
                    {
                        lblAdresB.Visible = true;
                        lblAdres.Visible = true;
                        lblAdres.Text = row.Cells["Adres"].Value.ToString();
                    }

                    //Hesaplama
                    var hesap = ctx.Kisis.Join(ctx.Caris,
                        kisiler => kisiler.KisiID,
                        cariler => cariler.KisiID,
                        (ki, ca) => new
                        {
                            kKisiID = ki.KisiID,
                            cKisiID = ca.KisiID,
                            ca.Tutar,
                            ca.DurumID
                        }).Select(x => new
                        {
                            x.cKisiID,
                            x.kKisiID,
                            x.DurumID,
                            x.Tutar
                        });

                    //Alacak
                    var alacak = hesap.Where(x => x.cKisiID == Convert.ToInt32(lblID.Tag)).Where(x => x.DurumID == 1).Sum(x => x.Tutar);
                    var alindi = hesap.Where(x => x.cKisiID == Convert.ToInt32(lblID.Tag)).Where(x => x.DurumID == 2).Sum(x => x.Tutar);

                    if (alacak != null && alindi != null)
                    {
                        lblAlacak.Text = String.Format("{0:N}\n", (alacak - alindi));
                    }
                    else if (alacak != null && alindi == null)
                    {
                        lblAlacak.Text = String.Format("{0:N}\n", alacak);
                    }
                    else
                    {
                        lblAlacak.Text = String.Format("{0:N}\n", 0);
                    }

                    //Borç
                    var borc = hesap.Where(x => x.cKisiID == Convert.ToInt32(lblID.Tag)).Where(x => x.DurumID == 3).Sum(x => x.Tutar);
                    var odendi = hesap.Where(x => x.cKisiID == Convert.ToInt32(lblID.Tag)).Where(x => x.DurumID == 4).Sum(x => x.Tutar);

                    if (borc != null && odendi != null)
                    {
                        lblBorc.Text = String.Format("{0:N}\n", (borc - odendi));
                    }
                    else if (borc != null && odendi == null)
                    {
                        lblBorc.Text = String.Format("{0:N}\n", borc);
                    }
                    else
                    {
                        lblBorc.Text = String.Format("{0:N}\n", 0);
                    }

                    txtAra.Text = null;

                    //Anahtar
                    anahtar = true;
                }
            }
            else
            {
                MessageBox.Show("Lütfen 'Ad Soyad' giriniz!", "Uyarı", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        private void BtnESil_Click(object sender, EventArgs e)
        {
            Button btn = sender as Button;
            DialogResult sonuc = MessageBox.Show(lblSAd.Text + " adlı kişi silinecek!?", "Uyarı", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
            //Sil
            if (sonuc == DialogResult.Yes)
            {
                //Anahtar
                anahtar = false;

                //Boşsa dön
                if (dataGridView1.CurrentRow == null)
                {
                    MessageBox.Show(lblSAd.Text + " silinmedi. Zira sol sütunda herhangi bir satır gözükmüyor.", "Bilgi", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                } 

                //İşlem
                //Önce Cari Kısmı Sil
                int kisiId = (int)dataGridView1.CurrentRow.Cells["KisiID"].Value;
                var c = ctx.Caris.Where(id => id.KisiID == kisiId);
                ctx.Caris.DeleteAllOnSubmit(c);
                ctx.SubmitChanges();

                //Sonra Kişi Bilgileri Sil
                Kisi k = ctx.Kisis.SingleOrDefault(id => id.KisiID == kisiId);
                ctx.Kisis.DeleteOnSubmit(k);
                ctx.SubmitChanges();

                //Listele
                var listele = from kisiler in ctx.Kisis
                              orderby kisiler.Ad
                              select new
                              {
                                  kisiler.KisiID,
                                  kisiler.Ad,
                                  kisiler.Firma,
                                  kisiler.Tel1,
                                  kisiler.Tel2,
                                  kisiler.Adres,
                                  kisiler.Karaliste
                              };
                dataGridView1.DataSource = listele;
                dataGridView1.Columns["Firma"].Visible = dataGridView1.Columns["Tel1"].Visible = dataGridView1.Columns["Tel2"].Visible = dataGridView1.Columns["Adres"].Visible = dataGridView1.Columns["KisiID"].Visible = dataGridView1.Columns["Karaliste"].Visible = false;
                DataGridViewColumn column = dataGridView1.Columns[1];
                column.Width = dataGridView1.Width - 20;

                //Alanları temizle
                lblSAd.Text = null;
                lblSTel1.Text = null;
                lblSTel2.Text = null;
                lblSFirma.Text = null;
                lblSAdres.Text = null;
                txtGAd.Text = null;
                txtGTel1.Text = null;
                txtGTel2.Text = null;
                ChkKaraListe.Checked = false;
                txtGFirma.Text = null;
                txtGAdres.Text = null;
                //lblAd.Text = null;
                //lblDAd.Text = null;
                //lblTel1.Text = null;
                //lblTel2.Text = null;
                //lblFirma.Text = null;
                //lblAdres.Text = null;
                txtAciklama.Text = null;
                dataGridView2.DataSource = null;
                dataGridView3.DataSource = null;
                cmbDurum.SelectedIndex = -1;
                txtAra.Text = null;
                //lblAlacak.Text = null;
                //lblBorc.Text = null;

                //Değişiklikleri altta düzenle bölümünde listele
                var liste2 = from kisiler in ctx.Kisis
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

                //Allta Bilgiler Bölümünde Listele
                dataGridView2.DataSource = liste2.Where(x => x.KisiID == Convert.ToInt32(lblID.Tag));
                dataGridView2.Columns["KisiID"].Visible = dataGridView2.Columns["Ad"].Visible = dataGridView2.Columns["DurumID"].Visible = dataGridView2.Columns["CariID"].Visible = false;

                dataGridView2.Columns[3].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                dataGridView2.Columns[3].HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleRight;
                dataGridView2.Columns[4].HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
                dataGridView2.Columns[4].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;

                //Üstte Bilgiler Bölümünde Listele
                DataGridViewRow row = dataGridView1.CurrentRow;
                lblAd.Text = row.Cells["Ad"].Value.ToString();
                lblDAd.Text = row.Cells["Ad"].Value.ToString();

                LblKaraListe.Tag = row.Cells["Karaliste"].Value;
                if ((Boolean)LblKaraListe.Tag)
                {
                    LblKaraListe.ForeColor = Color.DarkRed;
                }
                else
                {
                    LblKaraListe.ForeColor = Color.White;
                }

                if ((string)row.Cells["Firma"].Value == "")
                {
                    label2.Visible = false;
                    lblFirma.Visible = false;
                }
                else
                {
                    label2.Visible = true;
                    lblFirma.Visible = true;
                    lblFirma.Text = row.Cells["Firma"].Value.ToString();
                }

                if ((string)row.Cells["Tel1"].Value == "")
                {
                    label4.Visible = false;
                    lblTel1.Visible = false;
                }
                else
                {
                    label4.Visible = true;
                    lblTel1.Visible = true;
                    lblTel1.Text = row.Cells["Tel1"].Value.ToString();
                }

                if ((string)row.Cells["Tel2"].Value == "")
                {
                    lblTel2B.Visible = false;
                    lblTel2.Visible = false;
                }
                else
                {
                    lblTel2B.Visible = true;
                    lblTel2.Visible = true;
                    lblTel2.Text = row.Cells["Tel2"].Value.ToString();
                }

                if ((string)row.Cells["Adres"].Value == "")
                {
                    lblAdresB.Visible = false;
                    lblAdres.Visible = false;
                }
                else
                {
                    lblAdresB.Visible = true;
                    lblAdres.Visible = true;
                    lblAdres.Text = row.Cells["Adres"].Value.ToString();
                }

                //Hesaplama
                var hesap = ctx.Kisis.Join(ctx.Caris,
                    kisiler => kisiler.KisiID,
                    cariler => cariler.KisiID,
                    (ki, ca) => new
                    {
                        kKisiID = ki.KisiID,
                        cKisiID = ca.KisiID,
                        ca.Tutar,
                        ca.DurumID
                    }).Select(x => new
                    {
                        x.cKisiID,
                        x.kKisiID,
                        x.DurumID,
                        x.Tutar
                    });

                //Alacak
                var alacak = hesap.Where(x => x.cKisiID == Convert.ToInt32(lblID.Tag)).Where(x => x.DurumID == 1).Sum(x => x.Tutar);
                var alindi = hesap.Where(x => x.cKisiID == Convert.ToInt32(lblID.Tag)).Where(x => x.DurumID == 2).Sum(x => x.Tutar);

                if (alacak != null && alindi != null)
                {
                    lblAlacak.Text = String.Format("{0:N}\n", (alacak - alindi));
                }
                else if (alacak != null && alindi == null)
                {
                    lblAlacak.Text = String.Format("{0:N}\n", alacak);
                }
                else
                {
                    lblAlacak.Text = String.Format("{0:N}\n", 0);
                }

                //Borç
                var borc = hesap.Where(x => x.cKisiID == Convert.ToInt32(lblID.Tag)).Where(x => x.DurumID == 3).Sum(x => x.Tutar);
                var odendi = hesap.Where(x => x.cKisiID == Convert.ToInt32(lblID.Tag)).Where(x => x.DurumID == 4).Sum(x => x.Tutar);

                if (borc != null && odendi != null)
                {
                    lblBorc.Text = String.Format("{0:N}\n", (borc - odendi));
                }
                else if (borc != null && odendi == null)
                {
                    lblBorc.Text = String.Format("{0:N}\n", borc);
                }
                else
                {
                    lblBorc.Text = String.Format("{0:N}\n", 0);
                }

                txtAra.Text = null;

                //Anahtar
                anahtar = true;
            }
        }

        private void BtnExcel2_Click(object sender, EventArgs e)
        {
            //Dosya ismi
            string fileName;

            SaveFileDialog saveFileDialog1 = new SaveFileDialog
            {
                Filter = "xls files (*.xlsx)|*.xlsx|All files (*.*)|*.*",
                Title = "To Excel",
                FileName = " (" + DateTime.Now.ToString("dd-MM-yyyy") + ")" + " " + lblAd.Text
            };

            if (saveFileDialog1.ShowDialog() == DialogResult.OK)
            {
                fileName = saveFileDialog1.FileName;
                var workbook = new XLWorkbook();

                //DataTable Oluştur
                DataTable dt = new DataTable();

                //No Sütunu Ekle
                DataColumn columno = new DataColumn
                {
                    DataType = System.Type.GetType("System.Int32"),
                    AutoIncrement = true,
                    AutoIncrementSeed = 1,
                    AutoIncrementStep = 1
                };

                dt.Columns.Add(columno);
                dt.Columns["Column1"].ColumnName = "No";

                
                //Sütunları Ekle
                foreach (DataGridViewColumn column in dataGridView2.Columns)
                {
                    //dt.Columns.Add(column.HeaderText, column.ValueType);
                    dt.Columns.Add(column.HeaderText);
                }
               

                //Satırları Ekle
                foreach (DataGridViewRow row in dataGridView2.Rows)
                {
                    dt.Rows.Add();
                    //Hücreleri Ekle
                    foreach (DataGridViewCell cell in row.Cells)
                    {
                        if (cell.Value != null)
                        {
                            dt.Rows[dt.Rows.Count - 1][cell.ColumnIndex + 1] = cell.Value;
                        }
                    }
                }

                //Gereksiz Sütunları Kaldır
                dt.Columns.RemoveAt(1);
                dt.Columns.RemoveAt(1);
                dt.Columns.RemoveAt(5);
                dt.Columns.RemoveAt(5);

                //Tarih Sütunu Zamanı Kaldır
                dt.Columns.Add("Tarihler", typeof(String));
                for (int i = 4; i < dataGridView2.Columns.Count - 3; i++)
                {
                    for (int j = 0; j < dataGridView2.Rows.Count; j++)
                    {
                        dt.Rows[j][5] = string.Format("{0:dd/MM/yyyy}", dataGridView2.Rows[j].Cells[i].Value);
                    }
                }
                int columnNumber = dt.Columns["Tarih"].Ordinal;
                dt.Columns.Remove("Tarih");
                dt.Columns["Tarihler"].SetOrdinal(columnNumber);
                dt.Columns["Tarihler"].ColumnName = "Tarih";


                //Para Ayracı Ekle
                dt.Columns.Add("Tutarlar", typeof(String));
                for (int i = 3; i < dataGridView2.Columns.Count - 4; i++)
                {
                    for (int j = 0; j < dataGridView2.Rows.Count; j++)
                    {
                        dt.Rows[j][5] = string.Format("{0:N}\n", dataGridView2.Rows[j].Cells[i].Value);
                    }
                }
                int columnNumber2 = dt.Columns["Tutar"].Ordinal;
                dt.Columns.Remove("Tutar");
                dt.Columns["Tutarlar"].SetOrdinal(columnNumber2);
                dt.Columns["Tutarlar"].ColumnName = "Tutar";


                //Bakiye-Alacak Satırı Ekle
                DataRow rowAlacak = dt.NewRow();
                dt.Rows.Add();
                dt.Rows.Add(rowAlacak);
                rowAlacak[1] = "Bakiye :";
                rowAlacak[2] = lblAlacak.Text;


                //Alındı Satırı Ekle
                DataRow rowAlindi = dt.NewRow();
                dt.Rows.Add(rowAlindi);
                rowAlindi[1] = "Alındı :";
                rowAlindi[2] = LblAlindi.Text;

                //İskonto Satırı Ekle
                DataRow rowIskonto = dt.NewRow();
                dt.Rows.Add(rowIskonto);
                rowIskonto[1] = "İskonto :";
                rowIskonto[2] = LblIskonto.Text;


                //Kalan-Alacak Satırı Ekle
                DataRow rowKalanAlacak = dt.NewRow();
                dt.Rows.Add(rowKalanAlacak);
                rowKalanAlacak[1] = "Alacak :";
                rowKalanAlacak[2] = LblKalanAlacak.Text;

                dt.Rows.Add();

                //Bakiye-Borç Satırı Ekle
                DataRow rowBorc = dt.NewRow();
                dt.Rows.Add(rowBorc);
                rowBorc[1] = "Bakiye :";
                rowBorc[2] = lblBorc.Text;


                //Ödendi Satırı Ekle
                DataRow rowOdendi = dt.NewRow();
                dt.Rows.Add(rowOdendi);
                rowOdendi[1] = "Ödendi :";
                rowOdendi[2] = LblOdendi.Text;


                //Kalan-Borç Satırı Ekle
                DataRow rowKalanBorc = dt.NewRow();
                dt.Rows.Add(rowKalanBorc);
                rowKalanBorc[1] = "Borç :";
                rowKalanBorc[2] = LblKalanBorc.Text;

                //Excel Sayfasına Kişinin İsmini ekle.
                using (workbook)
                {
                    workbook.Worksheets.Add(dt, lblAd.Text);
                }
                
                //Düzenleme No Sütunu En Altlar Sil, Toplam Font Bold Yap
                workbook.Worksheet(lblAd.Text).Row(dataGridView2.Rows.Count + 3).Style.Font.SetBold();
                workbook.Worksheet(lblAd.Text).Row(dataGridView2.Rows.Count + 4).Style.Font.SetBold();
                workbook.Worksheet(lblAd.Text).Row(dataGridView2.Rows.Count + 5).Style.Font.SetBold();
                workbook.Worksheet(lblAd.Text).Row(dataGridView2.Rows.Count + 6).Style.Font.SetBold();
                workbook.Worksheet(lblAd.Text).Row(dataGridView2.Rows.Count + 7).Style.Font.SetBold();
                workbook.Worksheet(lblAd.Text).Row(dataGridView2.Rows.Count + 8).Style.Font.SetBold();
                workbook.Worksheet(lblAd.Text).Row(dataGridView2.Rows.Count + 9).Style.Font.SetBold();
                workbook.Worksheet(lblAd.Text).Row(dataGridView2.Rows.Count + 10).Style.Font.SetBold();
                workbook.Worksheet(lblAd.Text).Row(1).Style.Fill.BackgroundColor = XLColor.White;
                workbook.Worksheet(lblAd.Text).Row(1).Style.Font.SetBold();
                workbook.Worksheet(lblAd.Text).Row(1).Style.Font.SetFontColor(XLColor.FromArgb(0, 0, 1));
                workbook.Worksheet(lblAd.Text).Cell(1, 1).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
                workbook.Worksheet(lblAd.Text).PageSetup.Header.Left.AddText(lblAd.Text).SetBold();
                workbook.Worksheet(lblAd.Text).PageSetup.Header.Right.AddText(string.Format("{0:dd/MM/yyyy}", DateTime.Now));
                workbook.Worksheet(lblAd.Text).Column(4).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                workbook.Worksheet(lblAd.Text).Column(3).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
                workbook.Worksheet(lblAd.Text).Cell(1, 3).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                workbook.Worksheet(lblAd.Text).Cell(dataGridView2.Rows.Count + 2, 1).Value = "";
                workbook.Worksheet(lblAd.Text).Cell(dataGridView2.Rows.Count + 3, 1).Value = "";
                workbook.Worksheet(lblAd.Text).Cell(dataGridView2.Rows.Count + 4, 1).Value = "";
                workbook.Worksheet(lblAd.Text).Cell(dataGridView2.Rows.Count + 5, 1).Value = "";
                workbook.Worksheet(lblAd.Text).Cell(dataGridView2.Rows.Count + 6, 1).Value = "";
                workbook.Worksheet(lblAd.Text).Cell(dataGridView2.Rows.Count + 7, 1).Value = "";
                workbook.Worksheet(lblAd.Text).Cell(dataGridView2.Rows.Count + 8, 1).Value = "";
                workbook.Worksheet(lblAd.Text).Cell(dataGridView2.Rows.Count + 9, 1).Value = "";
                workbook.Worksheet(lblAd.Text).Cell(dataGridView2.Rows.Count + 10, 1).Value = "";
                workbook.Worksheet(lblAd.Text).Cell(dataGridView2.Rows.Count + 3, 2).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
                workbook.Worksheet(lblAd.Text).Cell(dataGridView2.Rows.Count + 4, 2).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
                workbook.Worksheet(lblAd.Text).Cell(dataGridView2.Rows.Count + 5, 2).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
                workbook.Worksheet(lblAd.Text).Cell(dataGridView2.Rows.Count + 6, 2).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
                workbook.Worksheet(lblAd.Text).Cell(dataGridView2.Rows.Count + 7, 2).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
                workbook.Worksheet(lblAd.Text).Cell(dataGridView2.Rows.Count + 8, 2).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
                workbook.Worksheet(lblAd.Text).Cell(dataGridView2.Rows.Count + 9, 2).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
                workbook.Worksheet(lblAd.Text).Cell(dataGridView2.Rows.Count + 10, 2).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;


                //Satır Arkaplan Renkleri Ata
                for (int i = 3; i < dataGridView2.Rows.Count + 2; i += 2)
                {
                    workbook.Worksheet(lblAd.Text).Row(i).Style.Fill.BackgroundColor = XLColor.LightGray;
                }

                for (int i = 2; i < dataGridView2.Rows.Count + 4; i += 2)
                {
                    workbook.Worksheet(lblAd.Text).Row(i).Style.Fill.BackgroundColor = XLColor.White;
                }
                workbook.Worksheet(lblAd.Text).Row(dataGridView2.Rows.Count + 4).Style.Fill.BackgroundColor = XLColor.White;
                workbook.Worksheet(lblAd.Text).Row(dataGridView2.Rows.Count + 5).Style.Fill.BackgroundColor = XLColor.White;
                workbook.Worksheet(lblAd.Text).Row(dataGridView2.Rows.Count + 6).Style.Fill.BackgroundColor = XLColor.LightGray;
                workbook.Worksheet(lblAd.Text).Row(dataGridView2.Rows.Count + 7).Style.Fill.BackgroundColor = XLColor.White;
                workbook.Worksheet(lblAd.Text).Row(dataGridView2.Rows.Count + 8).Style.Fill.BackgroundColor = XLColor.White;
                workbook.Worksheet(lblAd.Text).Row(dataGridView2.Rows.Count + 9).Style.Fill.BackgroundColor = XLColor.White;
                workbook.Worksheet(lblAd.Text).Row(dataGridView2.Rows.Count + 10).Style.Fill.BackgroundColor = XLColor.LightGray;

                workbook.Worksheet(lblAd.Text).Columns().AdjustToContents();
                //workbook.Worksheet(lblAd.Text).PageSetup.PageOrientation = XLPageOrientation.Landscape;
                workbook.Worksheet(lblAd.Text).PageSetup.FitToPages(1, 2);
                //Kaydet
                do
                {
                    try
                    {
                        workbook.SaveAs(fileName);
                        MessageBox.Show("Excel dosyası kaydedildi.", "Bilgi", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        break;
                    }
                    catch (System.IO.IOException)
                    {
                        MessageBox.Show("Kayıt yapılamadı! Kaydetmeye çalıştığınız dosya açık olabilir.", "Uyarı", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        break;
                    }
                } while (true);
            }
        }

        private void TxtTutar_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar) && (e.KeyChar != ','))
            {
                e.Handled = true;
            }

            // decimal için virgül
            if ((e.KeyChar == ',') && ((sender as TextBox).Text.IndexOf(',') > -1))
            {
                e.Handled = true;
            }
        }


        private void BtnAlindi_Click(object sender, EventArgs e)
        {
            var date1 = Dtp1.Value;
            var date2 = Dtp2.Value;   
            DataGridView5.DataSource = null;
            label6.Text = null;
            if (date1 <= date2)
            {
                LblRapor2.Text = "ALINDI" + " (" + (string.Format("{0:dd/MM/yyyy}", date1)) + " - " + (string.Format("{0:dd/MM/yyyy}", date2)) + ")";
                //Hesaplama - Alındı
                var alindi = from kisiler in ctx.Kisis
                               join cariler in ctx.Caris on kisiler.KisiID equals cariler.KisiID
                               join durumlar in ctx.Durums on cariler.DurumID equals durumlar.DurumID

                               group cariler by new { cariler.KisiID, kisiler.Ad, kisiler.Tel1 } into grup
                               let Ad = grup.Key.Ad
                               let Telefon = grup.Key.Tel1
                               let Alındı = grup.Where(x => (x.DurumID == 2 || x.DurumID == 8) && (x.Tarih.Value >= date1 && x.Tarih.Value <= date2)).Sum(x => (decimal?)x.Tutar) ?? 0

                               orderby grup.Key.Ad
                               select new
                               {
                                   grup.Key.Ad,
                                   Telefon,
                                   Alındı,
                               };
                DataGridView5.DataSource = alindi.Where(x => x.Alındı != 0);

                for (int i = 2; i < DataGridView5.Columns.Count; i++)
                {
                    decimal toplam = 0;
                    for (int j = 0; j < DataGridView5.Rows.Count + 1; ++j)
                    {
                        if (j < DataGridView5.Rows.Count)
                        {
                            toplam += Convert.ToDecimal(DataGridView5.Rows[j].Cells[i].Value);
                        }
                        else
                        {
                            LblTutar.Text = String.Format("{0:N}\n", toplam);
                        }
                    }
                }
            } 
            else
            {
                LblRapor2.Text = "ALINDI" + " (" + (string.Format("{0:dd/MM/yyyy}", date2)) + " - " + (string.Format("{0:dd/MM/yyyy}", date1)) + ")";
                //Hesaplama - Alındı
                var alindi = from kisiler in ctx.Kisis
                               join cariler in ctx.Caris on kisiler.KisiID equals cariler.KisiID
                               join durumlar in ctx.Durums on cariler.DurumID equals durumlar.DurumID

                               group cariler by new { cariler.KisiID, kisiler.Ad, kisiler.Tel1 } into grup
                               let Ad = grup.Key.Ad
                               let Telefon = grup.Key.Tel1
                               let Alındı = grup.Where(x => x.DurumID == 2 && (x.Tarih.Value >= date2 && x.Tarih.Value <= date1)).Sum(x => (decimal?)x.Tutar) ?? 0

                               orderby grup.Key.Ad
                               select new
                               {
                                   grup.Key.Ad,
                                   Telefon,
                                   Alındı,
                               };
                DataGridView5.DataSource = alindi.Where(x => x.Alındı != 0);

                for (int i = 2; i < DataGridView5.Columns.Count; i++)
                {
                    decimal toplam = 0;
                    for (int j = 0; j < DataGridView5.Rows.Count + 1; ++j)
                    {
                        if (j < DataGridView5.Rows.Count)
                        {
                            toplam += Convert.ToDecimal(DataGridView5.Rows[j].Cells[i].Value);
                        }
                        else
                        {
                            LblTutar.Text = String.Format("{0:N}\n", toplam);
                        }
                    }
                }
            }
            
            DataGridView5.Columns[0].HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
            DataGridView5.Columns[1].HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
            DataGridView5.Columns[1].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            DataGridView5.Columns[2].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            DataGridView5.Columns[2].HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
        }

        private void BtnOdendi_Click(object sender, EventArgs e)
        {
            var date1 = Dtp1.Value;
            var date2 = Dtp2.Value;
            DataGridView5.DataSource = null;
            label6.Text = null;
            if (date1 <= date2)
            {
                LblRapor2.Text = "ÖDENDİ" + " (" + (string.Format("{0:dd/MM/yyyy}", date1)) + " - " + (string.Format("{0:dd/MM/yyyy}", date2)) + ")";
                //Hesaplama - Ödendi
                var odendi = from kisiler in ctx.Kisis
                               join cariler in ctx.Caris on kisiler.KisiID equals cariler.KisiID
                               join durumlar in ctx.Durums on cariler.DurumID equals durumlar.DurumID

                               group cariler by new { cariler.KisiID, kisiler.Ad, kisiler.Tel1 } into grup
                               let Ad = grup.Key.Ad
                               let Telefon = grup.Key.Tel1
                               let Ödendi = grup.Where(x => (x.DurumID == 4 || x.DurumID == 10) && (x.Tarih.Value >= date1 && x.Tarih.Value <= date2)).Sum(x => (decimal?)x.Tutar) ?? 0

                               orderby grup.Key.Ad
                               select new
                               {
                                   grup.Key.Ad,
                                   Telefon,
                                   Ödendi
                               };
                DataGridView5.DataSource = odendi.Where(x => x.Ödendi != 0);

                for (int i = 2; i < DataGridView5.Columns.Count; i++)
                {
                    decimal toplam = 0;
                    for (int j = 0; j < DataGridView5.Rows.Count + 1; ++j)
                    {
                        if (j < DataGridView5.Rows.Count)
                        {
                            toplam += Convert.ToDecimal(DataGridView5.Rows[j].Cells[i].Value);
                        }
                        else
                        {
                            LblTutar.Text = String.Format("{0:N}\n", toplam);
                        }
                    }
                }
            }
            else
            {
                LblRapor2.Text = "ÖDENDİ" + " (" + (string.Format("{0:dd/MM/yyyy}", date2)) + " - " + (string.Format("{0:dd/MM/yyyy}", date1)) + ")";
                //Hesaplama - Alındı
                var odendi = from kisiler in ctx.Kisis
                             join cariler in ctx.Caris on kisiler.KisiID equals cariler.KisiID
                             join durumlar in ctx.Durums on cariler.DurumID equals durumlar.DurumID

                             group cariler by new { cariler.KisiID, kisiler.Ad, kisiler.Tel1 } into grup
                             let Ad = grup.Key.Ad
                             let Telefon = grup.Key.Tel1
                             let Ödendi = grup.Where(x => x.DurumID == 4 && (x.Tarih.Value >= date2 && x.Tarih.Value <= date1)).Sum(x => (decimal?)x.Tutar) ?? 0

                             orderby grup.Key.Ad
                             select new
                             {
                                 grup.Key.Ad,
                                 Telefon,
                                 Ödendi,
                             };
                DataGridView5.DataSource = odendi.Where(x => x.Ödendi != 0);

                for (int i = 2; i < DataGridView5.Columns.Count; i++)
                {
                    decimal toplam = 0;
                    for (int j = 0; j < DataGridView5.Rows.Count + 1; ++j)
                    {
                        if (j < DataGridView5.Rows.Count)
                        {
                            toplam += Convert.ToDecimal(DataGridView5.Rows[j].Cells[i].Value);
                        }
                        else
                        {
                            LblTutar.Text = String.Format("{0:N}\n", toplam);
                        }
                    }
                }
            }

            DataGridView5.Columns[0].HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
            DataGridView5.Columns[1].HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
            DataGridView5.Columns[1].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            DataGridView5.Columns[2].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            DataGridView5.Columns[2].HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
        }

        private void BtnExcel3_Click(object sender, EventArgs e)
        {
            CultureInfo customCulture = (CultureInfo)System.Threading.Thread.CurrentThread.CurrentCulture.Clone();
            customCulture.NumberFormat.NumberDecimalSeparator = ".";
            System.Threading.Thread.CurrentThread.CurrentCulture = customCulture;
            if (LblRapor2.Text != "RAPOR")
            {
                //Dosya ismi
                string fileName;

                SaveFileDialog saveFileDialog1 = new SaveFileDialog
                {
                    Filter = "xls files (*.xlsx)|*.xlsx|All files (*.*)|*.*",
                    Title = "To Excel",
                    //FileName = this.Text + " " + LblRapor2.Text
                    FileName = LblRapor2.Text
                };


                if (saveFileDialog1.ShowDialog() == DialogResult.OK)
                {
                    fileName = saveFileDialog1.FileName;
                    var workbook = new XLWorkbook();

                    //DataTable Oluştur
                    DataTable dt = new DataTable();

                    //No Sütunu Ekle
                    DataColumn columno = new DataColumn
                    {
                        DataType = System.Type.GetType("System.Int32"),
                        AutoIncrement = true,
                        AutoIncrementSeed = 1,
                        AutoIncrementStep = 1
                    };

                    dt.Columns.Add(columno);
                    dt.Columns["Column1"].ColumnName = "No";

                    //Sütunları Ekle
                    foreach (DataGridViewColumn column in DataGridView5.Columns)
                    {
                        //dt.Columns.Add(column.HeaderText, column.ValueType);
                        dt.Columns.Add(column.HeaderText);
                    }

                    //Satırları Ekle
                    foreach (DataGridViewRow row in DataGridView5.Rows)
                    {
                        dt.Rows.Add();
                        //Hücreleri Ekle
                        foreach (DataGridViewCell cell in row.Cells)
                        {
                            if (cell.Value != null)
                            {
                                dt.Rows[dt.Rows.Count - 1][cell.ColumnIndex + 1] = cell.Value;
                            }
                        }
                    }

                    //Toplam Satırı Ekle
                    DataRow rowToplam = dt.NewRow();
                    dt.Rows.Add();
                    dt.Rows.Add(rowToplam);
                    rowToplam[2] = "Toplam :";

                    for (int i = 2; i < DataGridView5.Columns.Count; i++)
                    {
                        decimal toplam = 0;
                        for (int j = 0; j < DataGridView5.Rows.Count + 1; ++j)
                        {
                            if (j < DataGridView5.Rows.Count)
                            {
                                toplam += Convert.ToDecimal(DataGridView5.Rows[j].Cells[i].Value);
                            }
                            else
                            {
                                rowToplam[i + 1] = toplam;
                            }
                        }
                    }

                    //Excel Sayfasına 'Kale Mobilya' yı ekle.
                    using (workbook)
                    {
                        workbook.Worksheets.Add(dt, "Kale Mobilya");
                    }

                    //Düzenleme
                    workbook.Worksheet("Kale Mobilya").Column(4).Style.NumberFormat.NumberFormatId = 4;

                    //Paraları Sağa Yaslama
                    /*
                    for (int i = 1; i < DataGridView5.Rows.Count + 4; i++)
                    {
                        for (int j = 4; j < DataGridView5.Columns.Count + 2; j++)
                        {
                            workbook.Worksheet("Kale Mobilya").Cell(i, j).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
                        }
                    }
                    */
                    workbook.Worksheet("Kale Mobilya").Column(4).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
                    workbook.Worksheet("Kale Mobilya").Cell(1, 4).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                    //Düzenleme No Sütunu En Altlar Sil, Toplam Font Bold Yap ve Sağa Hizala, Başlıklar Bold, Siyah
                    workbook.Worksheet("Kale Mobilya").Cell(DataGridView5.Rows.Count + 2, 1).Value = "";
                    workbook.Worksheet("Kale Mobilya").Cell(DataGridView5.Rows.Count + 3, 1).Value = "";
                    workbook.Worksheet("Kale Mobilya").Cell(DataGridView5.Rows.Count + 3, 3).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
                    workbook.Worksheet("Kale Mobilya").Row(DataGridView5.Rows.Count + 3).Style.Font.SetBold();
                    workbook.Worksheet("Kale Mobilya").Row(1).Style.Fill.BackgroundColor = XLColor.White;
                    workbook.Worksheet("Kale Mobilya").Row(1).Style.Font.SetBold();
                    workbook.Worksheet("Kale Mobilya").Row(1).Style.Font.SetFontColor(XLColor.FromArgb(0, 0, 1));
                    workbook.Worksheet("Kale Mobilya").Cell(1, 1).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
                    workbook.Worksheet("Kale Mobilya").PageSetup.Header.Left.AddText(LblRapor2.Text).SetBold();
                    workbook.Worksheet("Kale Mobilya").PageSetup.Header.Right.AddText(string.Format("{0:dd/MM/yyyy}", DateTime.Now));
                    /*
                    //Satır Arkaplan Renkleri Ata
                    for (int i = 3; i < DataGridView5.Rows.Count + 2; i += 2)
                    {
                        workbook.Worksheet("Kale Mobilya").Row(i).Style.Fill.BackgroundColor = XLColor.LightGray;
                    }

                    for (int i = 2; i < DataGridView5.Rows.Count + 4; i += 2)
                    {
                        workbook.Worksheet("Kale Mobilya").Row(i).Style.Fill.BackgroundColor = XLColor.White;
                    }
                    */

                    //Değerler tekrar atanarak sayı hale getiriliyor..
                    for (int i = 2; i < DataGridView5.Rows.Count + 4; i++)
                    {
                        workbook.Worksheet("Kale Mobilya").Cell(i, 4).Value = workbook.Worksheet("Kale Mobilya").Cell(i, 4).Value;
                    }

                    workbook.Worksheet("Kale Mobilya").Columns().AdjustToContents();

                    //Kaydet
                    do
                    {
                        try
                        {
                            workbook.SaveAs(fileName);
                            MessageBox.Show("Excel dosyası kaydedildi.", "Bilgi", MessageBoxButtons.OK, MessageBoxIcon.Information);
                            customCulture.NumberFormat.NumberDecimalSeparator = ",";
                            System.Threading.Thread.CurrentThread.CurrentCulture = customCulture;
                            break;
                        }
                        catch (System.IO.IOException)
                        {
                            MessageBox.Show("Kayıt yapılamadı! Kaydetmeye çalıştığınız dosya açık olabilir.", "Uyarı", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                            customCulture.NumberFormat.NumberDecimalSeparator = ",";
                            System.Threading.Thread.CurrentThread.CurrentCulture = customCulture;
                            break;
                        }
                    } while (true);
                }
            }
            else
            {
                MessageBox.Show("Herhangi bir rapor seçili değil!", "Uyarı", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        private void DataGridView5_DataBindingComplete(object sender, DataGridViewBindingCompleteEventArgs e)
        {
            foreach (DataGridViewRow dGVRow in this.DataGridView5.Rows)
            {
                dGVRow.HeaderCell.Value = String.Format("{0}", dGVRow.Index + 1);
            }
            //Genişlik Ayarı
            this.DataGridView5.AutoResizeRowHeadersWidth(DataGridViewRowHeadersWidthSizeMode.AutoSizeToAllHeaders);
        }

        private void BtnGider_Click(object sender, EventArgs e)
        {
            var date1 = Dtp1.Value;
            var date2 = Dtp2.Value;
            DataGridView5.DataSource = null;
            label6.Text = null;
            if (date1 <= date2)
            {
                LblRapor2.Text = "GİDER (" + (string.Format("{0:dd/MM/yyyy}", date1)) + " - " + (string.Format("{0:dd/MM/yyyy}", date2)) + ")";
                //Hesaplama - Gider
                var odendi = from kisiler in ctx.Kisis
                             join cariler in ctx.Caris on kisiler.KisiID equals cariler.KisiID
                             join durumlar in ctx.Durums on cariler.DurumID equals durumlar.DurumID

                             group cariler by new { cariler.KisiID, kisiler.Ad, kisiler.Tel1 } into grup
                             let Ad = grup.Key.Ad
                             let Telefon = grup.Key.Tel1
                             let Gider = grup.Where(x => x.DurumID == 5 && (x.Tarih.Value >= date1 && x.Tarih.Value <= date2)).Sum(x => (decimal?)x.Tutar) ?? 0

                             orderby grup.Key.Ad
                             select new
                             {
                                 grup.Key.Ad,
                                 Telefon,
                                 Gider,
                             };
                DataGridView5.DataSource = odendi.Where(x => x.Gider != 0);

                for (int i = 2; i < DataGridView5.Columns.Count; i++)
                {
                    decimal toplam = 0;
                    for (int j = 0; j < DataGridView5.Rows.Count + 1; ++j)
                    {
                        if (j < DataGridView5.Rows.Count)
                        {
                            toplam += Convert.ToDecimal(DataGridView5.Rows[j].Cells[i].Value);
                        }
                        else
                        {
                            LblTutar.Text = String.Format("{0:N}\n", toplam);
                        }
                    }
                }
            }
            else
            {
                LblRapor2.Text = "GİDER (" + (string.Format("{0:dd/MM/yyyy}", date2)) + " - " + (string.Format("{0:dd/MM/yyyy}", date1)) + ")";
                //Hesaplama - Gider
                var odendi = from kisiler in ctx.Kisis
                             join cariler in ctx.Caris on kisiler.KisiID equals cariler.KisiID
                             join durumlar in ctx.Durums on cariler.DurumID equals durumlar.DurumID

                             group cariler by new { cariler.KisiID, kisiler.Ad, kisiler.Tel1 } into grup
                             let Ad = grup.Key.Ad
                             let Telefon = grup.Key.Tel1
                             let Gider = grup.Where(x => x.DurumID == 5 && (x.Tarih.Value >= date2 && x.Tarih.Value <= date1)).Sum(x => (decimal?)x.Tutar) ?? 0

                             orderby grup.Key.Ad
                             select new
                             {
                                 grup.Key.Ad,
                                 Telefon,
                                 Gider,
                             };
                DataGridView5.DataSource = odendi.Where(x => x.Gider != 0);

                for (int i = 2; i < DataGridView5.Columns.Count; i++)
                {
                    decimal toplam = 0;
                    for (int j = 0; j < DataGridView5.Rows.Count + 1; ++j)
                    {
                        if (j < DataGridView5.Rows.Count)
                        {
                            toplam += Convert.ToDecimal(DataGridView5.Rows[j].Cells[i].Value);
                        }
                        else
                        {
                            LblTutar.Text = String.Format("{0:N}\n", toplam);
                        }
                    }
                }
            }

            DataGridView5.Columns[0].HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
            DataGridView5.Columns[1].HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
            DataGridView5.Columns[1].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            DataGridView5.Columns[2].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            DataGridView5.Columns[2].HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
        }

        private void BtnGirdiCikti_Click(object sender, EventArgs e)
        {
            var date1 = Dtp1.Value;
            var date2 = Dtp2.Value;

            if (date1 <= date2)
            {
                //Hesaplama - Girdi/Çıktı
                var girdicikti = from kisiler in ctx.Kisis
                                 join cariler in ctx.Caris on kisiler.KisiID equals cariler.KisiID
                                 join durumlar in ctx.Durums on cariler.DurumID equals durumlar.DurumID

                                 group cariler by new { cariler.KisiID, kisiler.Ad, kisiler.Tel1 } into grup
                                 let Ad = grup.Key.Ad
                                 let Telefon = grup.Key.Tel1
                                 let Gider = grup.Where(x => x.DurumID == 5 && (x.Tarih.Value >= date1 && x.Tarih.Value <= date2)).Sum(x => (decimal?)x.Tutar) ?? 0
                                 let Ödeme = grup.Where(x => x.DurumID == 10 && (x.Tarih.Value >= date1 && x.Tarih.Value <= date2)).Sum(x => (decimal?)x.Tutar) ?? 0
                                 let Ödendi = grup.Where(x => x.DurumID == 4 && (x.Tarih.Value >= date1 && x.Tarih.Value <= date2)).Sum(x => (decimal?)x.Tutar) ?? 0
                                 let Alındı = grup.Where(x => (x.DurumID == 2 || x.DurumID == 8) && (x.Tarih.Value >= date1 && x.Tarih.Value <= date2)).Sum(x => (decimal?)x.Tutar) ?? 0

                                 orderby grup.Key.Ad
                                 select new
                                 {
                                     grup.Key.Ad,
                                     Telefon,
                                     Gider,
                                     Ödeme,
                                     Ödendi,
                                     Alındı
                                 };

                var GiderToplam = girdicikti.Sum(x => x.Gider);
                var OdemeToplam = girdicikti.Sum(x => x.Ödeme);
                var OdendiToplam = girdicikti.Sum(x => x.Ödendi);
                var AlindiToplam = girdicikti.Sum(x => x.Alındı);

                label6.Text = string.Format("{0:N}\n", AlindiToplam - (GiderToplam + OdendiToplam + OdemeToplam));
            }
            else
            {
                MessageBox.Show("'Tarih 1', 'Tarih 2' de küçük olmalı!", "Uyarı", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        private void ChkGelirGider_CheckedChanged(object sender, EventArgs e)
        {
            if (ChkGelirGider.Checked)
            {
                GelenGiden gg = new GelenGiden();
                DialogResult sonuc = gg.ShowDialog();
                if (sonuc == DialogResult.OK)
                {
                    ChkGelirGider.Checked = true;
                }
                else if (sonuc==DialogResult.Cancel)
                {
                    ChkGelirGider.Checked = false;
                }
            }  
        }

        private void BtnGelirGider_Click(object sender, EventArgs e)
        {
            CultureInfo customCulture = (CultureInfo)System.Threading.Thread.CurrentThread.CurrentCulture.Clone();
            customCulture.NumberFormat.NumberDecimalSeparator = ".";
            System.Threading.Thread.CurrentThread.CurrentCulture = customCulture;

            //Değişkenler
            var date1 = Dtp1.Value;
            var date2 = Dtp2.Value;
            decimal toplam1 = 0;
            decimal toplam2 = 0;
            decimal toplam3 = 0;
            //Hesaplama - Alındı
            var alindi = from kisiler in ctx.Kisis
                         join cariler in ctx.Caris on kisiler.KisiID equals cariler.KisiID
                         join durumlar in ctx.Durums on cariler.DurumID equals durumlar.DurumID

                         group cariler by new { cariler.KisiID, kisiler.Ad, kisiler.Tel1 } into grup
                         let Ad = grup.Key.Ad
                         let Telefon = grup.Key.Tel1
                         let Alındı = grup.Where(x => (x.DurumID == 2 || x.DurumID == 8) && (x.Tarih.Value >= date1 && x.Tarih.Value <= date2)).Sum(x => (decimal?)x.Tutar) ?? 0

                         orderby grup.Key.Ad
                         select new
                         {
                             grup.Key.Ad,
                             Telefon,
                             Alındı
                         };
            //Hesaplama - Ödendi
            var odendi = from kisiler in ctx.Kisis
                         join cariler in ctx.Caris on kisiler.KisiID equals cariler.KisiID
                         join durumlar in ctx.Durums on cariler.DurumID equals durumlar.DurumID

                         group cariler by new { cariler.KisiID, kisiler.Ad, kisiler.Tel1 } into grup
                         let Ad = grup.Key.Ad
                         let Telefon = grup.Key.Tel1
                         let Ödendi = grup.Where(x => (x.DurumID == 4 || x.DurumID == 10) && (x.Tarih.Value >= date1 && x.Tarih.Value <= date2)).Sum(x => (decimal?)x.Tutar) ?? 0

                         orderby grup.Key.Ad
                         select new
                         {
                             grup.Key.Ad,
                             Telefon,
                             Ödendi
                         };
            //Hesaplama - Gider
            var gider = from kisiler in ctx.Kisis
                        join cariler in ctx.Caris on kisiler.KisiID equals cariler.KisiID
                        join durumlar in ctx.Durums on cariler.DurumID equals durumlar.DurumID

                        group cariler by new { cariler.KisiID, kisiler.Ad, kisiler.Tel1 } into grup
                        let Ad = grup.Key.Ad
                        let Telefon = grup.Key.Tel1
                        let Gider = grup.Where(x => x.DurumID == 5 && (x.Tarih.Value >= date1 && x.Tarih.Value <= date2)).Sum(x => (decimal?)x.Tutar) ?? 0

                        orderby grup.Key.Ad
                        select new
                        {
                            grup.Key.Ad,
                            Gider
                        };
            DataGridView5.DataSource = null;
            if (date1 <= date2)
            {
                //Dosya ismi
                string fileName;

                SaveFileDialog saveFileDialog1 = new SaveFileDialog
                {
                    Filter = "xls files (*.xlsx)|*.xlsx|All files (*.*)|*.*",
                    Title = "To Excel",
                    FileName = "(" + (string.Format("{0:dd/MM/yyyy}", date1)) + " - " + (string.Format("{0:dd/MM/yyyy}", date2)) + ") " + "GENEL"
                };

                if (saveFileDialog1.ShowDialog() == DialogResult.OK)
                {
                    fileName = saveFileDialog1.FileName;
                    var workbook = new XLWorkbook();

                    //DataTable Oluştur
                    DataTable dt = new DataTable();

                    ///                  TÜM SUTUNLARI EKLEME-BAŞLANGIÇ               ///
                    //No Sütunu Ekle
                    DataColumn columno = new DataColumn
                    {
                        DataType = System.Type.GetType("System.Int32"),
                        AutoIncrement = true,
                        AutoIncrementSeed = 1,
                        AutoIncrementStep = 1
                    };

                    dt.Columns.Add(columno);
                    dt.Columns["Column1"].ColumnName = "No";
                    dt.Columns.Add();
                    dt.Columns[1].ColumnName = "Ad - Soyad";
                    dt.Columns.Add();
                    dt.Columns[2].ColumnName = "Telefon";
                    dt.Columns.Add();
                    dt.Columns[3].ColumnName = "Alındı";
                    dt.Columns.Add();
                    dt.Columns[4].ColumnName = ",";
                    dt.Columns.Add();
                    dt.Columns[5].ColumnName = "İsim - Soyisim";
                    dt.Columns.Add();
                    dt.Columns[6].ColumnName = "İletişim";
                    dt.Columns.Add();
                    dt.Columns[7].ColumnName = "Ödendi";
                    dt.Columns.Add();
                    dt.Columns[8].ColumnName = ".";
                    dt.Columns.Add();
                    dt.Columns[9].ColumnName = "Açıklama";
                    dt.Columns.Add();
                    dt.Columns[10].ColumnName = "Gider";
                    ///                    TÜM SUTUNLARI EKLEME-SON                     ///
                    ///                            - - -                                ///
                    ///                  TÜM SATIRLARI EKLEME-BAŞLANGIÇ                 ///
                    //Karşılaştırma için DataTable ler Oluştur - Başlangıç
                    DataTable dt1 = new DataTable();
                    DataTable dt2 = new DataTable();
                    DataTable dt3 = new DataTable();

                    //Datatable Alındı Satırları Oluştur
                    DataGridView5.DataSource = alindi.Where(x => x.Alındı != 0);
                    dt1.Columns.Add();
                    dt1.Columns.Add();
                    dt1.Columns.Add();
                    //Alındı - Satırları Ekle
                    foreach (DataGridViewRow row in DataGridView5.Rows)
                    {
                        dt1.Rows.Add();
                    }
                    DataGridView5.DataSource = null;
                    //Datatable Ödendi Satırları Oluştur
                    DataGridView5.DataSource = odendi.Where(x => x.Ödendi != 0);
                    dt2.Columns.Add();
                    dt2.Columns.Add();
                    dt2.Columns.Add();
                    //Ödendi - Satırları Ekle
                    foreach (DataGridViewRow row in DataGridView5.Rows)
                    {
                        dt2.Rows.Add();
                    }
                    DataGridView5.DataSource = null;
                    //Datatable Gider Satırları Oluştur
                    DataGridView5.DataSource = gider.Where(x => x.Gider != 0);
                    dt3.Columns.Add();
                    dt3.Columns.Add();
                    //Gider - Satırları Ekle
                    foreach (DataGridViewRow row in DataGridView5.Rows)
                    {
                        dt3.Rows.Add();
                    }
                    DataGridView5.DataSource = null;
                    //Karşılaştırma için DataTable ler Oluştur - SON
                    ///
                    //Datatable ler Satır Sayısı Enyüksek Olan İçin Satırları ve Verileri Ekle - BAŞLANGIÇ
                    if (dt1.Rows.Count >= dt2.Rows.Count && dt1.Rows.Count >= dt3.Rows.Count)
                    {
                        DataGridView5.DataSource = alindi.Where(x => x.Alındı != 0);
                        //Satırları Ekle
                        foreach (DataGridViewRow row in DataGridView5.Rows)
                        {
                            dt.Rows.Add();
                            //Hücreleri Ekle
                            foreach (DataGridViewCell cell in row.Cells)
                            {
                                if (cell.Value != null)
                                {
                                    dt.Rows[dt.Rows.Count - 1][cell.ColumnIndex + 1] = cell.Value;
                                }
                            }
                        }
                        //Toplam Satırı Ekle - Alındı
                        DataRow rowToplam = dt.NewRow();
                        dt.Rows.Add();
                        dt.Rows.Add(rowToplam);
                        rowToplam[1] = "Toplam :";
                        rowToplam[2] = "Gelirler :";
                        rowToplam[6] = "Ödemeler :";
                        rowToplam[9] = "Giderler :";
                        for (int i = 2; i < DataGridView5.Columns.Count; i++)
                        {
                            decimal toplam = 0;
                            for (int j = 0; j < DataGridView5.Rows.Count + 1; ++j)
                            {
                                if (j < DataGridView5.Rows.Count)
                                {
                                    toplam += Convert.ToDecimal(DataGridView5.Rows[j].Cells[i].Value);
                                }
                                else
                                {
                                    rowToplam[3] = toplam;
                                    toplam1 = toplam;
                                }
                            }
                        }
                        DataGridView5.DataSource = null;
                        DataGridView5.DataSource = odendi.Where(x => x.Ödendi != 0);
                        for (int i = 0; i < DataGridView5.Columns.Count; i++)
                        {
                            for (int j = 0; j < DataGridView5.Rows.Count; ++j)
                            {
                                dt.Rows[j][i+5] = DataGridView5.Rows[j].Cells[i].Value;
                            }
                        }
                        //Toplam - Ödendi
                        for (int i = 2; i < DataGridView5.Columns.Count; i++)
                        {
                            decimal toplam = 0;
                            for (int j = 0; j < DataGridView5.Rows.Count + 1; ++j)
                            {
                                if (j < DataGridView5.Rows.Count)
                                {
                                    toplam += Convert.ToDecimal(DataGridView5.Rows[j].Cells[i].Value);
                                }
                                else
                                {
                                    rowToplam[7] = toplam;
                                    toplam2 = toplam;
                                }
                            }
                        }
                        DataGridView5.DataSource = null;
                        DataGridView5.DataSource = gider.Where(x => x.Gider != 0);
                        for (int i = 0; i < DataGridView5.Columns.Count; i++)
                        {
                            for (int j = 0; j < DataGridView5.Rows.Count; ++j)
                            {
                                dt.Rows[j][i + 9] = DataGridView5.Rows[j].Cells[i].Value;
                            }
                        }
                        //Toplam - Gider
                        for (int i = 1; i < DataGridView5.Columns.Count; i++)
                        {
                            decimal toplam = 0;
                            for (int j = 0; j < DataGridView5.Rows.Count + 1; ++j)
                            {
                                if (j < DataGridView5.Rows.Count)
                                {
                                    toplam += Convert.ToDecimal(DataGridView5.Rows[j].Cells[i].Value);
                                }
                                else
                                {
                                    rowToplam[10] = toplam;
                                    toplam3 = toplam;
                                }
                            }
                        }
                        DataGridView5.DataSource = null;
                    }
                    else if (dt2.Rows.Count >= dt1.Rows.Count && dt2.Rows.Count >= dt3.Rows.Count)
                    {
                        DataGridView5.DataSource = odendi.Where(x => x.Ödendi != 0);
                        //Satırları Ekle
                        foreach (DataGridViewRow row in DataGridView5.Rows)
                        {
                            dt.Rows.Add();
                            //Hücreleri Ekle
                            foreach (DataGridViewCell cell in row.Cells)
                            {
                                if (cell.Value != null)
                                {
                                    dt.Rows[dt.Rows.Count - 1][cell.ColumnIndex + 5] = cell.Value;
                                }
                            }
                        }
                        //Toplam Satırı Ekle - Ödendi
                        DataRow rowToplam = dt.NewRow();
                        dt.Rows.Add();
                        dt.Rows.Add(rowToplam);
                        rowToplam[1] = "Toplam :";
                        rowToplam[2] = "Gelirler :";
                        rowToplam[6] = "Ödemeler :";
                        rowToplam[9] = "Giderler :";
                        for (int i = 2; i < DataGridView5.Columns.Count; i++)
                        {
                            decimal toplam = 0;
                            for (int j = 0; j < DataGridView5.Rows.Count + 1; ++j)
                            {
                                if (j < DataGridView5.Rows.Count)
                                {
                                    toplam += Convert.ToDecimal(DataGridView5.Rows[j].Cells[i].Value);
                                }
                                else
                                {
                                    rowToplam[7] = toplam;
                                    toplam2 = toplam;
                                }
                            }
                        }
                        DataGridView5.DataSource = null;
                        DataGridView5.DataSource = alindi.Where(x => x.Alındı != 0);
                        for (int i = 0; i < DataGridView5.Columns.Count; i++)
                        {
                            for (int j = 0; j < DataGridView5.Rows.Count; ++j)
                            {
                                dt.Rows[j][i + 1] = DataGridView5.Rows[j].Cells[i].Value;
                            }
                        }
                        //Toplam - Alındı
                        for (int i = 2; i < DataGridView5.Columns.Count; i++)
                        {
                            decimal toplam = 0;
                            for (int j = 0; j < DataGridView5.Rows.Count + 1; ++j)
                            {
                                if (j < DataGridView5.Rows.Count)
                                {
                                    toplam += Convert.ToDecimal(DataGridView5.Rows[j].Cells[i].Value);
                                }
                                else
                                {
                                    rowToplam[3] = toplam;
                                    toplam1 = toplam;
                                }
                            }
                        }
                        DataGridView5.DataSource = null;
                        DataGridView5.DataSource = gider.Where(x => x.Gider != 0);
                        for (int i = 0; i < DataGridView5.Columns.Count; i++)
                        {
                            for (int j = 0; j < DataGridView5.Rows.Count; ++j)
                            {
                                dt.Rows[j][i + 9] = DataGridView5.Rows[j].Cells[i].Value;
                            }
                        }
                        //Toplam - Gider
                        for (int i = 1; i < DataGridView5.Columns.Count; i++)
                        {
                            decimal toplam = 0;
                            for (int j = 0; j < DataGridView5.Rows.Count + 1; ++j)
                            {
                                if (j < DataGridView5.Rows.Count)
                                {
                                    toplam += Convert.ToDecimal(DataGridView5.Rows[j].Cells[i].Value);
                                }
                                else
                                {
                                    rowToplam[10] = toplam;
                                    toplam3 = toplam;
                                }
                            }
                        }
                        DataGridView5.DataSource = null;
                    }
                    else if (dt3.Rows.Count >= dt1.Rows.Count && dt3.Rows.Count >= dt2.Rows.Count)
                    {
                        DataGridView5.DataSource = gider.Where(x => x.Gider != 0);
                        //Satırları Ekle
                        foreach (DataGridViewRow row in DataGridView5.Rows)
                        {
                            dt.Rows.Add();
                            //Hücreleri Ekle
                            foreach (DataGridViewCell cell in row.Cells)
                            {
                                if (cell.Value != null)
                                {
                                    dt.Rows[dt.Rows.Count - 1][cell.ColumnIndex + 9] = cell.Value;
                                }
                            }
                        }
                        //Toplam Satırı Ekle - Gider
                        DataRow rowToplam = dt.NewRow();
                        dt.Rows.Add();
                        dt.Rows.Add(rowToplam);
                        rowToplam[1] = "Toplam :";
                        rowToplam[2] = "Gelirler :";
                        rowToplam[6] = "Ödemeler :";
                        rowToplam[9] = "Giderler :";
                        for (int i = 1; i < DataGridView5.Columns.Count; i++)
                        {
                            decimal toplam = 0;
                            for (int j = 0; j < DataGridView5.Rows.Count + 1; ++j)
                            {
                                if (j < DataGridView5.Rows.Count)
                                {
                                    toplam += Convert.ToDecimal(DataGridView5.Rows[j].Cells[i].Value);
                                }
                                else
                                {
                                    rowToplam[10] = toplam;
                                    toplam3 = toplam;
                                }
                            }
                        }
                        DataGridView5.DataSource = null;
                        DataGridView5.DataSource = alindi.Where(x => x.Alındı != 0);
                        //Satırları Ekle
                        for (int i = 0; i < DataGridView5.Columns.Count; i++)
                        {
                            for (int j = 0; j < DataGridView5.Rows.Count; ++j)
                            {
                                dt.Rows[j][i + 1] = DataGridView5.Rows[j].Cells[i].Value;
                            }
                        }
                        //Toplam - Alındı
                        for (int i = 2; i < DataGridView5.Columns.Count; i++)
                        {
                            decimal toplam = 0;
                            for (int j = 0; j < DataGridView5.Rows.Count + 1; ++j)
                            {
                                if (j < DataGridView5.Rows.Count)
                                {
                                    toplam += Convert.ToDecimal(DataGridView5.Rows[j].Cells[i].Value);
                                }
                                else
                                {
                                    rowToplam[3] = toplam;
                                    toplam1 = toplam;
                                }
                            }
                        }
                        DataGridView5.DataSource = null;
                        DataGridView5.DataSource = odendi.Where(x => x.Ödendi != 0);
                        //Satırları Ekle
                        for (int i = 0; i < DataGridView5.Columns.Count; i++)
                        {
                            for (int j = 0; j < DataGridView5.Rows.Count; ++j)
                            {
                                dt.Rows[j][i + 5] = DataGridView5.Rows[j].Cells[i].Value;
                            }
                        }
                        //Toplam - Ödendi
                        for (int i = 2; i < DataGridView5.Columns.Count; i++)
                        {
                            decimal toplam = 0;
                            for (int j = 0; j < DataGridView5.Rows.Count + 1; ++j)
                            {
                                if (j < DataGridView5.Rows.Count)
                                {
                                    toplam += Convert.ToDecimal(DataGridView5.Rows[j].Cells[i].Value);
                                }
                                else
                                {
                                    rowToplam[7] = toplam;
                                    toplam2 = toplam;
                                }
                            }
                        }
                        DataGridView5.DataSource = null;
                    }
                    //Datatable ler Satır Sayısı Enyüksek Olan İçin Satırları ve Verileri Ekle - SON

                    //Excel Sayfasına 'Kale Mobilya' yı ekle.
                    using (workbook)
                    {
                        workbook.Worksheets.Add(dt, "Kale Mobilya");
                    }

                    /*
                    //Satır Arkaplan Renkleri Ata
                    for (int i = 3; i < (Math.Max(Math.Max(dt1.Rows.Count, dt2.Rows.Count), dt3.Rows.Count)) + 2; i += 2)
                    {
                        workbook.Worksheet("Kale Mobilya").Row(i).Style.Fill.BackgroundColor = XLColor.LightGray;
                    }

                    for (int i = 2; i < (Math.Max(Math.Max(dt1.Rows.Count, dt2.Rows.Count), dt3.Rows.Count)) + 4; i += 2)
                    {
                        workbook.Worksheet("Kale Mobilya").Row(i).Style.Fill.BackgroundColor = XLColor.White;
                    }
                    */
                    workbook.Worksheet("Kale Mobilya").Column(4).Style.NumberFormat.NumberFormatId = 4;
                    workbook.Worksheet("Kale Mobilya").Column(8).Style.NumberFormat.NumberFormatId = 4;
                    workbook.Worksheet("Kale Mobilya").Column(11).Style.NumberFormat.NumberFormatId = 4;

                    //Değerler tekrar atanarak sayı hale getiriliyor..
                     for (int i = 2; i < Math.Max(Math.Max(dt1.Rows.Count, dt2.Rows.Count), dt3.Rows.Count) + 4; i++)
                        {
                            workbook.Worksheet("Kale Mobilya").Cell(i, 4).Value = workbook.Worksheet("Kale Mobilya").Cell(i, 4).Value;
                        }

                    for (int i = 2; i < Math.Max(Math.Max(dt1.Rows.Count, dt2.Rows.Count), dt3.Rows.Count) + 4; i++)
                    {
                        workbook.Worksheet("Kale Mobilya").Cell(i, 8).Value = workbook.Worksheet("Kale Mobilya").Cell(i, 8).Value;
                    }

                    for (int i = 2; i < Math.Max(Math.Max(dt1.Rows.Count, dt2.Rows.Count), dt3.Rows.Count) + 4; i++)
                    {
                        workbook.Worksheet("Kale Mobilya").Cell(i, 11).Value = workbook.Worksheet("Kale Mobilya").Cell(i, 11).Value;
                    }


                    //Düzenleme No Sütunu En Altlar Sil, Toplam Font Bold Yap ve Sağa Hizala, Başlıklar Bold, Siyah
                    workbook.Worksheet("Kale Mobilya").Cell((Math.Max(Math.Max(dt1.Rows.Count, dt2.Rows.Count), dt3.Rows.Count)) + 2, 1).Value = "";
                    workbook.Worksheet("Kale Mobilya").Cell((Math.Max(Math.Max(dt1.Rows.Count, dt2.Rows.Count), dt3.Rows.Count)) + 3, 1).Value = "";
                    workbook.Worksheet("Kale Mobilya").Row((Math.Max(Math.Max(dt1.Rows.Count, dt2.Rows.Count), dt3.Rows.Count)) + 3).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
                    workbook.Worksheet("Kale Mobilya").Cell((Math.Max(Math.Max(dt1.Rows.Count, dt2.Rows.Count), dt3.Rows.Count)) + 4, 1).Value = "";
                    workbook.Worksheet("Kale Mobilya").Cell((Math.Max(Math.Max(dt1.Rows.Count, dt2.Rows.Count), dt3.Rows.Count)) + 5, 1).Value = "";
                    workbook.Worksheet("Kale Mobilya").Cell((Math.Max(Math.Max(dt1.Rows.Count, dt2.Rows.Count), dt3.Rows.Count)) + 5, 2).Value = "KALAN:";
                    workbook.Worksheet("Kale Mobilya").Cell((Math.Max(Math.Max(dt1.Rows.Count, dt2.Rows.Count), dt3.Rows.Count)) + 5, 2).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
                    workbook.Worksheet("Kale Mobilya").Cell((Math.Max(Math.Max(dt1.Rows.Count, dt2.Rows.Count), dt3.Rows.Count)) + 5, 2).Style.Font.SetUnderline();
                    workbook.Worksheet("Kale Mobilya").Cell((Math.Max(Math.Max(dt1.Rows.Count, dt2.Rows.Count), dt3.Rows.Count)) + 5, 3).Style.NumberFormat.NumberFormatId = 4;
                    workbook.Worksheet("Kale Mobilya").Cell((Math.Max(Math.Max(dt1.Rows.Count, dt2.Rows.Count), dt3.Rows.Count)) + 5, 3).Value = toplam1 - (toplam2 + toplam3);
                    workbook.Worksheet("Kale Mobilya").Row(Math.Max(Math.Max(dt1.Rows.Count, dt2.Rows.Count), dt3.Rows.Count) + 3).Style.Font.SetBold();
                    workbook.Worksheet("Kale Mobilya").Row(Math.Max(Math.Max(dt1.Rows.Count, dt2.Rows.Count), dt3.Rows.Count) + 5).Style.Font.SetBold();
                    workbook.Worksheet("Kale Mobilya").Column(4).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
                    workbook.Worksheet("Kale Mobilya").Column(8).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
                    workbook.Worksheet("Kale Mobilya").Column(11).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
                    workbook.Worksheet("Kale Mobilya").Cell(1, 1).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                    workbook.Worksheet("Kale Mobilya").Cell(1, 4).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                    workbook.Worksheet("Kale Mobilya").Cell(1, 8).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                    workbook.Worksheet("Kale Mobilya").Cell(1, 11).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                    workbook.Worksheet("Kale Mobilya").Row(1).Style.Fill.BackgroundColor = XLColor.White;
                    workbook.Worksheet("Kale Mobilya").Row(1).Style.Font.SetBold();
                    workbook.Worksheet("Kale Mobilya").Row(1).Style.Font.SetFontColor(XLColor.FromArgb(0, 0, 1));
                    //workbook.Worksheet("Kale Mobilya").Column(5).Style.Font.SetFontColor(XLColor.White);
                    workbook.Worksheet("Kale Mobilya").Cell(1,5).Style.Font.SetFontColor(XLColor.White);
                    //workbook.Worksheet("Kale Mobilya").Column(9).Style.Font.SetFontColor(XLColor.White);
                    workbook.Worksheet("Kale Mobilya").Cell(1, 9).Style.Font.SetFontColor(XLColor.White);
                    workbook.Worksheet("Kale Mobilya").PageSetup.Header.Center.AddText((string.Format("{0:dd/MM/yyyy}", date1)) + " - " + (string.Format("{0:dd/MM/yyyy}", date2))).SetBold();

                    workbook.Worksheet("Kale Mobilya").Columns().AdjustToContents();
                    workbook.Worksheet("Kale Mobilya").PageSetup.PageOrientation = XLPageOrientation.Landscape;
                    workbook.Worksheet("Kale Mobilya").PageSetup.FitToPages(1, 2);

                    //Kaydet
                    do
                    {
                        try
                        {
                            workbook.SaveAs(fileName);
                            MessageBox.Show("Excel dosyası kaydedildi.", "Bilgi", MessageBoxButtons.OK, MessageBoxIcon.Information);
                            customCulture.NumberFormat.NumberDecimalSeparator = ",";
                            System.Threading.Thread.CurrentThread.CurrentCulture = customCulture;
                            break;
                        }
                        catch (System.IO.IOException)
                        {
                            MessageBox.Show("Kayıt yapılamadı! Kaydetmeye çalıştığınız dosya açık olabilir.", "Uyarı", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                            customCulture.NumberFormat.NumberDecimalSeparator = ",";
                            System.Threading.Thread.CurrentThread.CurrentCulture = customCulture;
                            break;
                        }
                    } while (true);
                }
            }
            else
            {
                MessageBox.Show("'Tarih 1', 'Tarih 2' de küçük olmalı!", "Uyarı", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        private void BtnTemizle_Click(object sender, EventArgs e)
        {
            //Düzenle Kısmını Temizle
            txtAciklama.Text = "";
            cmbDurum.SelectedValue = -1;
            dtTarih.Text = DateTime.Now.ToString("d/M/yyyy");
            txtTutar.Text = "";
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
            dataGridView2.DataSource = sonuc.Where(x => x.KisiID == Convert.ToInt32(lblID.Tag) &&(x.Açıklama.Contains(txtAraCari.Text) || x.Durum.Contains(txtAraCari.Text) || Convert.ToString(x.Tutar).Contains(txtAraCari.Text)));
            dataGridView2.Columns["KisiID"].Visible = dataGridView2.Columns["Ad"].Visible = dataGridView2.Columns["DurumID"].Visible = dataGridView2.Columns["CariID"].Visible = false;
            //DataGridViewColumn column = dataGridView2.Columns[1];
            //column.Width = dataGridView1.Width - 20;
        }


        private void DataGridView4_CellPainting(object sender, DataGridViewCellPaintingEventArgs e)
        {  
            decimal sum = 0;
            if (e.RowIndex == this.dataGridView4.NewRowIndex && e.ColumnIndex > 5)
            {
                for (int i = 0; i < dataGridView4.Rows.Count; i++)
                {
                    sum += (decimal)this.dataGridView4.Rows[i].Cells[e.ColumnIndex].Value;
                }
                e.PaintBackground(e.CellBounds, false);
                e.Graphics.DrawString(dataGridView4.Columns[e.ColumnIndex].Name +  ": "+ String.Format("{0:N}\n", sum), this.dataGridView4.Font, Brushes.Black, e.CellBounds.Left + 2, e.CellBounds.Top + 3);
                e.Handled = true;
            } 
        }

        private void BtnAylik_Click(object sender, EventArgs e)
        {
            CultureInfo customCulture = (CultureInfo)System.Threading.Thread.CurrentThread.CurrentCulture.Clone();
            customCulture.NumberFormat.NumberDecimalSeparator = ".";
            System.Threading.Thread.CurrentThread.CurrentCulture = customCulture;

            //Değişkenler
            var date1 = Dtp1.Value;
            var date2 = Dtp2.Value;

            if (date1 <= date2)
            {
                DataGridView5.DataSource = null;
                //System.Globalization.DateTimeFormatInfo DFI = new System.Globalization.DateTimeFormatInfo();
                //Hesaplama
                var hesap = from cariler in ctx.Caris

                            group cariler by new { cariler.Tarih.Value.Year, cariler.Tarih.Value.Month } into grup
                            let Gelir = grup.Where(x => (x.DurumID == 2 || x.DurumID == 8) && (x.Tarih.Value >= date1 && x.Tarih.Value <= date2)).Sum(x => (decimal?)x.Tutar) ?? 0
                            //let Ödeme = grup.Where(x => x.DurumID == 4 && (x.Tarih.Value >= date1 && x.Tarih.Value <= date2)).Sum(x => (decimal?)x.Tutar) ?? 0
                            let Gider = grup.Where(x => (x.DurumID == 5 || x.DurumID == 4 || x.DurumID == 10) && (x.Tarih.Value >= date1 && x.Tarih.Value <= date2)).Sum(x => (decimal?)x.Tutar) ?? 0

                            orderby grup.Key.Year, grup.Key.Month
                            select new
                            {
                                Yıl = grup.Key.Year,
                                Ay = grup.Key.Month,
                                //Ay = DFI.GetMonthName(grup.Key.Month).ToString(),
                                //AyYıl = string.Format("{0}/{1}", grup.Key.Month, grup.Key.Year),
                                //YılveAy = grup.Key.Year + "-" + grup.Key.Month,
                                Gelir,
                                Gider,
                                Kalan = Gelir - Gider
                            };

                DataGridView5.DataSource = hesap.Where(x => x.Kalan != 0);

                DataGridView5.Columns[0].HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
                DataGridView5.Columns[0].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                DataGridView5.Columns[1].HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
                DataGridView5.Columns[1].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                DataGridView5.Columns[2].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                DataGridView5.Columns[2].HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
                DataGridView5.Columns[3].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                DataGridView5.Columns[3].HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
                DataGridView5.Columns[4].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                DataGridView5.Columns[4].HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;

                
                //Dosya ismi
                string fileName;

                SaveFileDialog saveFileDialog1 = new SaveFileDialog
                {
                    Filter = "xls files (*.xlsx)|*.xlsx|All files (*.*)|*.*",
                    Title = "To Excel",
                    FileName = "(" + (string.Format("{0:dd/MM/yyyy}", date1)) + " - " + (string.Format("{0:dd/MM/yyyy}", date2)) + ") " + "AYLIK"
                };

                if (saveFileDialog1.ShowDialog() == DialogResult.OK)
                {
                    fileName = saveFileDialog1.FileName;
                    var workbook = new XLWorkbook();

                    //DataTable Oluştur
                    DataTable dt = new DataTable();

                    //No Sütunu Ekle
                    DataColumn columno = new DataColumn
                    {
                        DataType = System.Type.GetType("System.Int32"),
                        AutoIncrement = true,
                        AutoIncrementSeed = 1,
                        AutoIncrementStep = 1
                    };

                    dt.Columns.Add(columno);
                    dt.Columns["Column1"].ColumnName = "No";

                    //Sütunları Ekle
                    foreach (DataGridViewColumn column in DataGridView5.Columns)
                    {
                        dt.Columns.Add(column.HeaderText);
                    }

                    //Satırları Ekle
                    foreach (DataGridViewRow row in DataGridView5.Rows)
                    {
                        dt.Rows.Add();
                        //Hücreleri Ekle
                        foreach (DataGridViewCell cell in row.Cells)
                        {
                            if (cell.Value != null)
                            {
                                dt.Rows[dt.Rows.Count - 1][cell.ColumnIndex + 1] = cell.Value;
                            }
                        }
                    }
                    /*
                    //Parasal Değerlere "." ve "," ekle
                    for (int i = 2; i < DataGridView5.Columns.Count; i++)
                    {
                        for (int j = 0; j < DataGridView5.Rows.Count; ++j)
                        {
                            dt.Rows[j][i+1] = String.Format("{0:N}\n", DataGridView5.Rows[j].Cells[i].Value);
                        }
                    }
                    */
                    //Toplam Satırı Ekle
                    DataRow rowToplam = dt.NewRow();
                    dt.Rows.Add();
                    dt.Rows.Add(rowToplam);
                    rowToplam[2] = "Toplam :";

                    for (int i = 2; i < DataGridView5.Columns.Count; i++)
                    {
                        decimal toplam = 0;
                        for (int j = 0; j < DataGridView5.Rows.Count + 1; ++j)
                        {
                            if (j < DataGridView5.Rows.Count)
                            {
                                toplam += Convert.ToDecimal(DataGridView5.Rows[j].Cells[i].Value);
                            }
                            else
                            {
                                rowToplam[i + 1] = toplam;
                            }
                        }
                    }

                    //Excel Sayfasına 'Kale Mobilya' yı ekle.
                    using (workbook)
                    {
                        workbook.Worksheets.Add(dt, "Kale Mobilya");
                    }

                    for (int i = 4; i < DataGridView5.Columns.Count + 2; i++)
                    {
                        workbook.Worksheet("Kale Mobilya").Column(i).Style.NumberFormat.NumberFormatId = 4;
                    }

                    //Değerler tekrar atanarak sayı hale getiriliyor..
                    for (int j = 4; j < DataGridView5.Columns.Count + 2; j++)
                    {
                        for (int i = 2; i < DataGridView5.Rows.Count + 4; i++)
                        {
                            workbook.Worksheet("Kale Mobilya").Cell(i, j).Value = workbook.Worksheet("Kale Mobilya").Cell(i, j).Value;
                        }
                    }

                    //Düzenleme - Paraları Sağa Yaslama
                    for (int j = 4; j < DataGridView5.Columns.Count + 2; j++)
                        {
                            workbook.Worksheet("Kale Mobilya").Column(j).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
                        }
                    
                    //Düzenleme No Sütunu En Altlar Sil, Toplam Font Bold Yap ve Sağa Hizala, Başlıklar Bold, Siyah
                    workbook.Worksheet("Kale Mobilya").Cell(DataGridView5.Rows.Count + 2, 1).Value = "";
                    workbook.Worksheet("Kale Mobilya").Cell(DataGridView5.Rows.Count + 3, 1).Value = "";
                    workbook.Worksheet("Kale Mobilya").Cell(DataGridView5.Rows.Count + 3, 3).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
                    workbook.Worksheet("Kale Mobilya").Row(DataGridView5.Rows.Count + 3).Style.Font.SetBold();
                    workbook.Worksheet("Kale Mobilya").Row(1).Style.Fill.BackgroundColor = XLColor.White;
                    workbook.Worksheet("Kale Mobilya").Row(1).Style.Font.SetBold();
                    workbook.Worksheet("Kale Mobilya").Row(1).Style.Font.SetFontColor(XLColor.FromArgb(0, 0, 1));
                    workbook.Worksheet("Kale Mobilya").PageSetup.Header.Left.AddText("Aylık Rapor");
                    workbook.Worksheet("Kale Mobilya").PageSetup.Header.Center.AddText((string.Format("{0:dd/MM/yyyy}", date1)) + " - " + (string.Format("{0:dd/MM/yyyy}", date2))).SetBold();
                    workbook.Worksheet("Kale Mobilya").PageSetup.Header.Right.AddText(string.Format("{0:dd/MM/yyyy}", DateTime.Now));
                    workbook.Worksheet("Kale Mobilya").Column(1).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                    workbook.Worksheet("Kale Mobilya").Column(2).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                    workbook.Worksheet("Kale Mobilya").Column(3).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                    workbook.Worksheet("Kale Mobilya").Row(1).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                    workbook.Worksheet("Kale Mobilya").Column(1).Style.Font.FontSize = 15;
                    workbook.Worksheet("Kale Mobilya").Column(2).Style.Font.FontSize = 15;
                    workbook.Worksheet("Kale Mobilya").Column(3).Style.Font.FontSize = 15;
                    workbook.Worksheet("Kale Mobilya").Column(4).Style.Font.FontSize = 15;
                    workbook.Worksheet("Kale Mobilya").Column(5).Style.Font.FontSize = 15;
                    workbook.Worksheet("Kale Mobilya").Column(6).Style.Font.FontSize = 15;
                    /*
                    //Satır Arkaplan Renkleri Ata
                    for (int i = 3; i < DataGridView5.Rows.Count + 2; i += 2)
                    {
                        workbook.Worksheet("Kale Mobilya").Row(i).Style.Fill.BackgroundColor = XLColor.LightGray;
                    }

                    for (int i = 2; i < DataGridView5.Rows.Count + 4; i += 2)
                    {
                        workbook.Worksheet("Kale Mobilya").Row(i).Style.Fill.BackgroundColor = XLColor.White;
                    }
                    */

                    workbook.Worksheet("Kale Mobilya").Columns().AdjustToContents();

                    //Kaydet
                    do
                    {
                        try
                        {
                            workbook.SaveAs(fileName);
                            MessageBox.Show("Excel dosyası kaydedildi.", "Bilgi", MessageBoxButtons.OK, MessageBoxIcon.Information);
                            customCulture.NumberFormat.NumberDecimalSeparator = ",";
                            System.Threading.Thread.CurrentThread.CurrentCulture = customCulture;
                            break;
                        }
                        catch (System.IO.IOException)
                        {
                            MessageBox.Show("Kayıt yapılamadı! Kaydetmeye çalıştığınız dosya açık olabilir.", "Uyarı", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                            customCulture.NumberFormat.NumberDecimalSeparator = ",";
                            System.Threading.Thread.CurrentThread.CurrentCulture = customCulture;
                            break;
                        }
                    } while (true);
                }
            }
            else
            {
                MessageBox.Show("'Tarih 1', 'Tarih 2' de küçük olmalı!", "Uyarı", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        
        }

        private void BtnExcel4_Click(object sender, EventArgs e)
        {
            if (dataGridView2.Rows.Count > 0)
            {
                CultureInfo customCulture = (CultureInfo)System.Threading.Thread.CurrentThread.CurrentCulture.Clone();
                customCulture.NumberFormat.NumberDecimalSeparator = ".";
                System.Threading.Thread.CurrentThread.CurrentCulture = customCulture;
                bool FaturaSonuc = false;
                string fay = "";

                if (lblAd.Text == "FATURALAR")
                {
                    //FaturaFormu Aç
                    FaturaForm ff = new FaturaForm();
                    DialogResult sonucf = ff.ShowDialog();

                    if (sonucf == DialogResult.OK)
                    {
                        FaturaSonuc = true;
                        DataGridViewRow row = dataGridView1.CurrentRow;
                        var sonuc2y2 = from kisiler in ctx.Kisis
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
                        dataGridView2.DataSource = sonuc2y2.Where(x => x.KisiID == Convert.ToInt32(row.Cells["KisiID"].Value) && x.Tarih.Value.Month == FTarihAy && x.Tarih.Value.Year == FTarihYil);
                        dataGridView2.Columns["KisiID"].Visible = dataGridView2.Columns["Ad"].Visible = dataGridView2.Columns["DurumID"].Visible = dataGridView2.Columns["CariID"].Visible = false;

                        dataGridView2.Columns[3].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                        dataGridView2.Columns[3].HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleRight;
                        dataGridView2.Columns[4].HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
                        dataGridView2.Columns[4].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;

                        if (FTarihAy == 1)
                        {
                            fay = "OCAK";
                        }
                        else if (FTarihAy == 2)
                        {
                            fay = "ŞUBAT";
                        }
                        else if (FTarihAy == 3)
                        {
                            fay = "MART";
                        }
                        else if (FTarihAy == 4)
                        {
                            fay = "NİSAN";
                        }
                        else if (FTarihAy == 5)
                        {
                            fay = "MAYIS";
                        }
                        else if (FTarihAy == 6)
                        {
                            fay = "HAZİRAN";
                        }
                        else if (FTarihAy == 7)
                        {
                            fay = "TEMMUZ";
                        }
                        else if (FTarihAy == 8)
                        {
                            fay = "AĞUSTOS";
                        }
                        else if (FTarihAy == 9)
                        {
                            fay = "EYLÜL";
                        }
                        else if (FTarihAy == 10)
                        {
                            fay = "EKİM";
                        }
                        else if (FTarihAy == 11)
                        {
                            fay = "KASIM";
                        }
                        else if (FTarihAy == 12)
                        {
                            fay = "ARALIK";
                        }
                    }
                    else if (sonucf == DialogResult.Abort)
                    {
                        FaturaSonuc = false;
                        DataGridViewRow row = dataGridView1.CurrentRow;
                        var sonuc2y = from kisiler in ctx.Kisis
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
                        dataGridView2.DataSource = sonuc2y.Where(x => x.KisiID == Convert.ToInt32(row.Cells["KisiID"].Value));
                        dataGridView2.Columns["KisiID"].Visible = dataGridView2.Columns["Ad"].Visible = dataGridView2.Columns["DurumID"].Visible = dataGridView2.Columns["CariID"].Visible = false;

                        dataGridView2.Columns[3].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                        dataGridView2.Columns[3].HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleRight;
                        dataGridView2.Columns[4].HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
                        dataGridView2.Columns[4].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                    }
                }

                /*
                tabControl1.TabPages.Add(tabPageHesaplama);
                var sonuc = from kisiler in ctx.Kisis
                            join cariler in ctx.Caris on kisiler.KisiID equals cariler.KisiID
                            join durumlar in ctx.Durums on cariler.DurumID equals durumlar.DurumID
                            orderby durumlar.DurumID, cariler.Tarih
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
                dataGridViewHesap.DataSource = sonuc.Where(x => x.KisiID == Convert.ToInt32(lblID.Text));
                dataGridViewHesap.Columns["KisiID"].Visible = dataGridViewHesap.Columns["Ad"].Visible = dataGridViewHesap.Columns["DurumID"].Visible = dataGridViewHesap.Columns["CariID"].Visible = false;
                */

                //Dosya ismi - Kaydet
                string fileName;

                SaveFileDialog saveFileDialog1 = new SaveFileDialog
                {
                    Filter = "xls files (*.xlsx)|*.xlsx|All files (*.*)|*.*",
                    Title = "To Excel",
                };

                if (FaturaSonuc)
                {
                    saveFileDialog1.FileName = "(" + fay + "-" + FTarihYil + ")" + " " + lblAd.Text;
                }
                else if (FaturaSonuc == false)
                {
                    saveFileDialog1.FileName = "(" + DateTime.Now.ToString("dd-MM-yyyy") + ")" + " " + lblAd.Text;
                }


                if (saveFileDialog1.ShowDialog() == DialogResult.OK)
                {
                    fileName = saveFileDialog1.FileName;
                    var workbook = new XLWorkbook();

                    //FATURALAR için Excel 2. Çalışma Sayfası ve Liste oluştur
                    if (lblAd.Text == "FATURALAR")
                    {
                        //DataTable2 Oluştur
                        DataTable dt2 = new DataTable();

                        //No Sütunu Ekle
                        DataColumn columno2 = new DataColumn
                        {
                            DataType = System.Type.GetType("System.Int32"),
                            AutoIncrement = true,
                            AutoIncrementSeed = 1,
                            AutoIncrementStep = 1
                        };

                        dt2.Columns.Add(columno2);
                        dt2.Columns["Column1"].ColumnName = "No";

                        //Sütunları Ekle
                        foreach (DataGridViewColumn column in dataGridView2.Columns)
                        {
                            dt2.Columns.Add(column.HeaderText);
                        }

                        //Gelen Satırları Ekle ve Gelen Toplam Hesapla
                        decimal ToplamGelen = 0;
                        decimal Toplam1 = 0;
                        int a = 0;
                        int satirekle = 0;
                        int gelensatir = 0;
                        for (int i = 0; i < dataGridView2.Rows.Count + 1; i++)
                        {
                            //if (i < dataGridView2.Rows.Count && (dataGridView2[5, i].Value.ToString().Contains("gelen",StringComparer.OrdinalIgnoreCase)))
                            if (i < dataGridView2.Rows.Count && Regex.IsMatch(dataGridView2[5, i].Value.ToString(), Regex.Escape("gelen"), RegexOptions.IgnoreCase))
                            {
                                dt2.Rows.Add();
                                gelensatir += 1;
                                satirekle += 1;
                                for (int j = 0; j < dataGridView2.Columns.Count; j++)
                                {
                                    if (j == 4)
                                    {
                                        dt2.Rows[a][j + 1] = string.Format("{0:dd/MM/yyyy}", dataGridView2.Rows[i].Cells[j].Value);
                                    }
                                    else
                                    {
                                        dt2.Rows[a][j + 1] = dataGridView2.Rows[i].Cells[j].Value;
                                    }
                                }
                                a += 1;
                                Toplam1 += Convert.ToDecimal(dataGridView2.Rows[i].Cells[3].Value);
                            }
                            else
                            {
                                ToplamGelen = Toplam1;
                            }
                        }

                        //Boş Satır Ekle
                        dt2.Rows.Add();

                        //Gelen Toplam Satır Ekle
                        DataRow rowGelenToplam = dt2.NewRow();
                        dt2.Rows.Add(rowGelenToplam);

                        //Boş Satırlar Ekle
                        dt2.Rows.Add();
                        dt2.Rows.Add();
                        satirekle += 4;

                        //Giden Satırları Ekle ve Giden Toplam Hesapla
                        decimal ToplamGiden = 0;
                        decimal Toplam2 = 0;
                        int b = satirekle;
                        for (int i = 0; i < dataGridView2.Rows.Count + 1; i++)
                        {
                            //if (i < dataGridView2.Rows.Count && (dataGridView2[5, i].Value.ToString().Contains("Giden") || dataGridView2[5, i].Value.ToString().Contains("GİDEN")))
                            if (i < dataGridView2.Rows.Count && Regex.IsMatch(dataGridView2[5, i].Value.ToString(), Regex.Escape("giden"), RegexOptions.IgnoreCase))
                            {
                                dt2.Rows.Add();
                                for (int j = 0; j < dataGridView2.Columns.Count; j++)
                                {
                                    if (j == 4)
                                    {
                                        dt2.Rows[b][j + 1] = string.Format("{0:dd/MM/yyyy}", dataGridView2.Rows[i].Cells[j].Value);
                                    }
                                    else
                                    {
                                        dt2.Rows[b][j + 1] = dataGridView2.Rows[i].Cells[j].Value;
                                    }
                                }
                                b += 1;
                                Toplam2 += Convert.ToDecimal(dataGridView2.Rows[i].Cells[3].Value);
                                satirekle += 1;
                            }
                            else
                            {
                                ToplamGiden = Toplam2;
                            }
                        }

                        //Boş Satır Ekle
                        dt2.Rows.Add();

                        //Giden Toplam Satır Ekle
                        DataRow rowGidenToplam = dt2.NewRow();
                        dt2.Rows.Add(rowGidenToplam);
                        satirekle += 2;

                        //Gereksiz Sütunları Kaldır
                        dt2.Columns.RemoveAt(1);
                        dt2.Columns.RemoveAt(1);
                        dt2.Columns.RemoveAt(5);
                        dt2.Columns.RemoveAt(5);

                        //rowGelenToplam[0] = "Gelen";
                        rowGelenToplam[1] = "Toplam :";
                        rowGelenToplam[2] = ToplamGelen;

                        //rowGidenToplam[0] = "Gelen";
                        rowGidenToplam[1] = "Toplam :";
                        rowGidenToplam[2] = ToplamGiden;


                        //Excel Sayfasına 'Liste-1' i ekle.
                        using (workbook)
                        {
                            workbook.Worksheets.Add(dt2, "Liste-1");
                        }

                        //Düzenleme 
                        //No Sütunu Gereksiz Noları Sil
                        workbook.Worksheet("Liste-1").Cell(gelensatir + 2, 1).Value = "";
                        workbook.Worksheet("Liste-1").Cell(gelensatir + 3, 1).Value = "";
                        workbook.Worksheet("Liste-1").Cell(gelensatir + 4, 1).Value = "";
                        workbook.Worksheet("Liste-1").Cell(gelensatir + 5, 1).Value = "";
                        workbook.Worksheet("Liste-1").Cell(satirekle, 1).Value = "";
                        workbook.Worksheet("Liste-1").Cell(satirekle + 1, 1).Value = "";

                        //Giden Satırlarına No Ekle
                        for (int i = 1; i < satirekle - (gelensatir + 5); i++)
                        {
                            workbook.Worksheet("Liste-1").Cell(gelensatir + 5 + i, 1).Value = i;
                        }

                        //Gelen-Giden Toplam da 'Gelen', 'Giden' Yazıları Ekle
                        workbook.Worksheet("Liste-1").Cell(gelensatir + 3, 1).Value = "Gelen";
                        workbook.Worksheet("Liste-1").Cell(satirekle + 1, 1).Value = "Giden";

                        //Tekrar Paraları Ayraç İçin Ekle
                        workbook.Worksheet("Liste-1").Column(3).Style.NumberFormat.NumberFormatId = 4;
                        for (int i = 2; i < dataGridView2.Rows.Count + satirekle + 2; i++)
                        {
                            workbook.Worksheet("Liste-1").Cell(i, 3).Value = workbook.Worksheet("Liste-1").Cell(i, 3).Value;
                        }

                        workbook.Worksheet("Liste-1").Row(1).Style.Fill.BackgroundColor = XLColor.White;
                        workbook.Worksheet("Liste-1").Row(satirekle + 1).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
                        workbook.Worksheet("Liste-1").Row(satirekle + 1).Style.Font.SetBold();
                        workbook.Worksheet("Liste-1").Row(gelensatir + 3).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
                        workbook.Worksheet("Liste-1").Row(gelensatir + 3).Style.Font.SetBold();
                        workbook.Worksheet("Liste-1").Row(1).Style.Font.SetFontColor(XLColor.FromArgb(0, 0, 1));
                        workbook.Worksheet("Liste-1").Cell(1, 1).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
                        workbook.Worksheet("Liste-1").PageSetup.Header.Left.AddText("Fatura-KDV").SetBold();
                        if (FaturaSonuc)
                        {
                            workbook.Worksheet("Liste-1").PageSetup.Header.Center.AddText(fay + " / " + FTarihYil).SetBold();
                        }
                        workbook.Worksheet("Liste-1").PageSetup.Header.Right.AddText(string.Format("{0:dd/MM/yyyy}", DateTime.Now));
                        workbook.Worksheet("Liste-1").Column(4).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                        workbook.Worksheet("Liste-1").Column(3).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
                        workbook.Worksheet("Liste-1").Cell(1, 3).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                    }

                    //---DataTable3 Oluştur İlk Genel Tablo---//

                    DataTable dt3 = new DataTable();

                    //No Sütunu Ekle
                    DataColumn columno3 = new DataColumn
                    {
                        DataType = System.Type.GetType("System.Int32"),
                        AutoIncrement = true,
                        AutoIncrementSeed = 1,
                        AutoIncrementStep = 1
                    };

                    dt3.Columns.Add(columno3);
                    dt3.Columns["Column1"].ColumnName = "No";

                    //Sütunları Ekle
                    foreach (DataGridViewColumn column in dataGridView2.Columns)
                    {
                        dt3.Columns.Add(column.HeaderText);
                    }

                    //Alacak Satırları Ekle ve Alacak Toplam Hesapla
                    decimal ToplamAlacak3 = 0;
                    decimal Toplam13 = 0;
                    int a3 = 0;
                    int SatirEkle3 = 0;
                    int AlacakSatir3 = 0;
                    for (int i = 0; i < dataGridView2.Rows.Count + 1; i++)
                    {
                        if (i < dataGridView2.Rows.Count && dataGridView2[2, i].Value.ToString() == "Alacak")
                        {
                            dt3.Rows.Add();
                            AlacakSatir3 += 1;
                            SatirEkle3 += 1;
                            for (int j = 0; j < dataGridView2.Columns.Count; j++)
                            {
                                if (j == 4)
                                {
                                    dt3.Rows[a3][j + 1] = string.Format("{0:dd/MM/yyyy}", dataGridView2.Rows[i].Cells[j].Value);
                                }
                                else
                                {
                                    dt3.Rows[a3][j + 1] = dataGridView2.Rows[i].Cells[j].Value;
                                }
                            }
                            a3 += 1;
                            Toplam13 += Convert.ToDecimal(dataGridView2.Rows[i].Cells[3].Value);
                        }
                        else
                        {
                            ToplamAlacak3 = Toplam13;
                        }
                    }

                    if (AlacakSatir3 != 0)
                    {
                        //Boş Satır Ekle
                        dt3.Rows.Add();

                        //Alacak Toplam Satır Ekle
                        DataRow rowAlacakToplam3 = dt3.NewRow();
                        dt3.Rows.Add(rowAlacakToplam3);

                        rowAlacakToplam3[3] = "Toplam :";
                        rowAlacakToplam3[4] = ToplamAlacak3;

                        //Boş Satırlar Ekle
                        dt3.Rows.Add();
                        dt3.Rows.Add();

                        SatirEkle3 += 4;
                    }

                    //İskonto Satırları Ekle ve İskonto Toplam Hesapla
                    decimal ToplamIskonto3 = 0;
                    decimal ToplamIs3 = 0;
                    int is3 = SatirEkle3;
                    int IskontoSatir3 = 0;
                    for (int i = 0; i < dataGridView2.Rows.Count + 1; i++)
                    {
                        if (i < dataGridView2.Rows.Count && dataGridView2[2, i].Value.ToString() == "İskonto")
                        {
                            IskontoSatir3 = 1;
                            ToplamIs3 += Convert.ToDecimal(dataGridView2.Rows[i].Cells[3].Value);
                        }
                        else
                        {
                            ToplamIskonto3 = ToplamIs3;
                        }
                    }

                    if (IskontoSatir3 != 0 && AlacakSatir3 != 0)
                    {
                        is3 += 2;
                        dt3.Rows.RemoveAt(SatirEkle3 - 1);
                        dt3.Rows.RemoveAt(SatirEkle3 - 2);

                        //İskonto Toplam Satır Ekle
                        DataRow rowIskontoToplam3 = dt3.NewRow();
                        dt3.Rows.Add(rowIskontoToplam3);

                        rowIskontoToplam3[3] = "İskonto :";
                        rowIskontoToplam3[4] = ToplamIskonto3;

                        //Alacak2 Toplam Satır Ekle
                        DataRow rowAlacak2Toplam3 = dt3.NewRow();
                        dt3.Rows.Add(rowAlacak2Toplam3);

                        rowAlacak2Toplam3[3] = "Alacak :";
                        rowAlacak2Toplam3[4] = ToplamAlacak3 - ToplamIskonto3;

                        //Boş Satırlar Ekle
                        dt3.Rows.Add();
                        dt3.Rows.Add();

                        SatirEkle3 += 2;
                    }

                    //Alındı Satırları Ekle ve Alındı Toplam Hesapla
                    decimal ToplamAlindi3 = 0;
                    decimal Toplam23 = 0;
                    int b3 = SatirEkle3;
                    int AlindiSatir3 = 0;
                    for (int i = 0; i < dataGridView2.Rows.Count + 1; i++)
                    {
                        if (i < dataGridView2.Rows.Count && dataGridView2[2, i].Value.ToString() == "Alındı")
                        {
                            dt3.Rows.Add();
                            AlindiSatir3 += 1;
                            for (int j = 0; j < dataGridView2.Columns.Count; j++)
                            {
                                if (j == 4)
                                {
                                    dt3.Rows[b3][j + 1] = string.Format("{0:dd/MM/yyyy}", dataGridView2.Rows[i].Cells[j].Value);
                                }
                                else
                                {
                                    dt3.Rows[b3][j + 1] = dataGridView2.Rows[i].Cells[j].Value;
                                }
                            }
                            b3 += 1;
                            Toplam23 += Convert.ToDecimal(dataGridView2.Rows[i].Cells[3].Value);
                            SatirEkle3 += 1;
                        }
                        else
                        {
                            ToplamAlindi3 = Toplam23;
                        }
                    }

                    if (AlindiSatir3 != 0)
                    {
                        //Boş Satır Ekle
                        dt3.Rows.Add();

                        //Alındı Toplam Satır Ekle
                        DataRow rowAlindiToplam3 = dt3.NewRow();
                        dt3.Rows.Add(rowAlindiToplam3);

                        rowAlindiToplam3[3] = "Toplam :";
                        rowAlindiToplam3[4] = ToplamAlindi3;

                        //Alındı2 Toplam Satır Ekle
                        DataRow rowAlindi2Toplam3 = dt3.NewRow();
                        dt3.Rows.Add(rowAlindi2Toplam3);

                        rowAlindi2Toplam3[3] = "Kalan :";
                        rowAlindi2Toplam3[4] = ToplamAlacak3 - (ToplamIskonto3 + ToplamAlindi3);

                        //Boş Satırlar Ekle
                        dt3.Rows.Add();
                        dt3.Rows.Add();

                        SatirEkle3 += 5;
                    }
                   
                    //Borç Satırları Ekle ve Borç Toplam Hesapla
                    decimal ToplamBorc3 = 0;
                    decimal Toplam33 = 0;
                    int c3 = SatirEkle3;
                    int BorcSatir3 = 0;
                    for (int i = 0; i < dataGridView2.Rows.Count + 1; i++)
                    {
                        if (i < dataGridView2.Rows.Count && dataGridView2[2, i].Value.ToString() == "Borç")
                        {
                            dt3.Rows.Add();
                            BorcSatir3 += 1;
                            for (int j = 0; j < dataGridView2.Columns.Count; j++)
                            {
                                if (j == 4)
                                {
                                    dt3.Rows[c3][j + 1] = string.Format("{0:dd/MM/yyyy}", dataGridView2.Rows[i].Cells[j].Value);
                                }
                                else
                                {
                                    dt3.Rows[c3][j + 1] = dataGridView2.Rows[i].Cells[j].Value;
                                }
                            }
                            c3 += 1;
                            Toplam33 += Convert.ToDecimal(dataGridView2.Rows[i].Cells[3].Value);
                            SatirEkle3 += 1;
                        }
                        else
                        {
                            ToplamBorc3 = Toplam33;
                        }
                    }

                    if (BorcSatir3 != 0)
                    {
                        //Boş Satır Ekle
                        dt3.Rows.Add();

                        //Borç Toplam Satır Ekle
                        DataRow rowBorcToplam3 = dt3.NewRow();
                        dt3.Rows.Add(rowBorcToplam3);

                        rowBorcToplam3[3] = "Toplam :";
                        rowBorcToplam3[4] = ToplamBorc3;

                        //Boş Satırlar Ekle
                        dt3.Rows.Add();
                        dt3.Rows.Add();

                        SatirEkle3 += 4;
                    }

                    //Ödendi Satırları Ekle ve Ödendi Toplam Hesapla
                    decimal ToplamOdendi3 = 0;
                    decimal Toplam43 = 0;
                    int d3 = SatirEkle3;
                    int OdendiSatir3 = 0;
                    for (int i = 0; i < dataGridView2.Rows.Count + 1; i++)
                    {
                        if (i < dataGridView2.Rows.Count && dataGridView2[2, i].Value.ToString() == "Ödendi")
                        {
                            dt3.Rows.Add();
                            OdendiSatir3 += 1;
                            for (int j = 0; j < dataGridView2.Columns.Count; j++)
                            {
                                if (j == 4)
                                {
                                    dt3.Rows[d3][j + 1] = string.Format("{0:dd/MM/yyyy}", dataGridView2.Rows[i].Cells[j].Value);
                                }
                                else
                                {
                                    dt3.Rows[d3][j + 1] = dataGridView2.Rows[i].Cells[j].Value;
                                }
                            }
                            d3 += 1;
                            Toplam43 += Convert.ToDecimal(dataGridView2.Rows[i].Cells[3].Value);
                            SatirEkle3 += 1;
                        }
                        else
                        {
                            ToplamOdendi3 = Toplam43;
                        }
                    }

                    if (OdendiSatir3 != 0)
                    {
                        //Boş Satır Ekle
                        dt3.Rows.Add();

                        //Ödendi Toplam Satır Ekle
                        DataRow rowOdendiToplam3 = dt3.NewRow();
                        dt3.Rows.Add(rowOdendiToplam3);

                        rowOdendiToplam3[3] = "Toplam :";
                        rowOdendiToplam3[4] = ToplamOdendi3;

                        //Boş Satırlar Ekle
                        dt3.Rows.Add();
                        dt3.Rows.Add();

                        SatirEkle3 += 4;
                    }


                    //Gider Satırları Ekle ve Gider Toplam Hesapla
                    decimal ToplamGider3 = 0;
                    decimal Toplam53 = 0;
                    int e3 = SatirEkle3;
                    int GiderSatir3 = 0;
                    for (int i = 0; i < dataGridView2.Rows.Count + 1; i++)
                    {
                        if (i < dataGridView2.Rows.Count && dataGridView2[2, i].Value.ToString() == "Gider")
                        {
                            dt3.Rows.Add();
                            GiderSatir3 += 1;
                            for (int j = 0; j < dataGridView2.Columns.Count; j++)
                            {
                                if (j == 4)
                                {
                                    dt3.Rows[e3][j + 1] = string.Format("{0:dd/MM/yyyy}", dataGridView2.Rows[i].Cells[j].Value);
                                }
                                else
                                {
                                    dt3.Rows[e3][j + 1] = dataGridView2.Rows[i].Cells[j].Value;
                                }
                            }
                            e3 += 1;
                            Toplam53 += Convert.ToDecimal(dataGridView2.Rows[i].Cells[3].Value);
                            SatirEkle3 += 1;
                        }
                        else
                        {
                            ToplamGider3 = Toplam53;
                        }
                    }

                    if (GiderSatir3 != 0)
                    {
                        //Boş Satır Ekle
                        dt3.Rows.Add();

                        //Gider Toplam Satır Ekle
                        DataRow rowGiderToplam3 = dt3.NewRow();
                        dt3.Rows.Add(rowGiderToplam3);

                        rowGiderToplam3[3] = "Toplam :";
                        rowGiderToplam3[4] = ToplamGider3;

                        //Boş Satırlar Ekle
                        dt3.Rows.Add();
                        dt3.Rows.Add();

                        SatirEkle3 += 4;
                    }


                    //Bilgi Satırları Ekle ve Bilgi Toplam Hesapla
                    decimal ToplamBilgi3 = 0;
                    decimal Toplam63 = 0;
                    int f3 = SatirEkle3;
                    int BilgiSatir3 = 0;
                    for (int i = 0; i < dataGridView2.Rows.Count + 1; i++)
                    {
                        if (i < dataGridView2.Rows.Count && dataGridView2[2, i].Value.ToString() == "Bilgi")
                        {
                            dt3.Rows.Add();
                            BilgiSatir3 += 1;
                            for (int j = 0; j < dataGridView2.Columns.Count; j++)
                            {
                                if (j == 4)
                                {
                                    dt3.Rows[f3][j + 1] = string.Format("{0:dd/MM/yyyy}", dataGridView2.Rows[i].Cells[j].Value);
                                }
                                else
                                {
                                    dt3.Rows[f3][j + 1] = dataGridView2.Rows[i].Cells[j].Value;
                                }
                            }
                            f3 += 1;
                            Toplam63 += Convert.ToDecimal(dataGridView2.Rows[i].Cells[3].Value);
                            SatirEkle3 += 1;
                        }
                        else
                        {
                            ToplamBilgi3 = Toplam63;
                        }
                    }

                    if (BilgiSatir3 != 0)
                    {
                        //Boş Satır Ekle
                        dt3.Rows.Add();

                        //Bilgi Toplam Satır Ekle
                        DataRow rowBilgiToplam3 = dt3.NewRow();
                        dt3.Rows.Add(rowBilgiToplam3);

                        rowBilgiToplam3[3] = "Toplam :";
                        rowBilgiToplam3[4] = ToplamBilgi3;

                        //Boş Satırlar Ekle
                        dt3.Rows.Add();
                        dt3.Rows.Add();

                        SatirEkle3 += 4;
                    }


                    //Mesai Satırları Ekle ve Mesai Toplam Hesapla
                    decimal ToplamMesai3 = 0;
                    decimal Toplam73 = 0;
                    int g3 = SatirEkle3;
                    int MesaiSatir3 = 0;
                    for (int i = 0; i < dataGridView2.Rows.Count + 1; i++)
                    {
                        if (i < dataGridView2.Rows.Count && dataGridView2[2, i].Value.ToString() == "Mesai")
                        {
                            dt3.Rows.Add();
                            MesaiSatir3 += 1;
                            for (int j = 0; j < dataGridView2.Columns.Count; j++)
                            {
                                if (j == 4)
                                {
                                    dt3.Rows[g3][j + 1] = string.Format("{0:dd/MM/yyyy}", dataGridView2.Rows[i].Cells[j].Value);
                                }
                                else
                                {
                                    dt3.Rows[g3][j + 1] = dataGridView2.Rows[i].Cells[j].Value;
                                }
                            }
                            g3 += 1;
                            Toplam73 += Convert.ToDecimal(dataGridView2.Rows[i].Cells[3].Value);
                            SatirEkle3 += 1;
                        }
                        else
                        {
                            ToplamMesai3 = Toplam73;
                        }
                    }

                    if (MesaiSatir3 != 0)
                    {
                        //Boş Satır Ekle
                        dt3.Rows.Add();

                        //Mesai Toplam Satır Ekle
                        DataRow rowMesaiToplam3 = dt3.NewRow();
                        dt3.Rows.Add(rowMesaiToplam3);

                        rowMesaiToplam3[3] = "Toplam :";
                        rowMesaiToplam3[4] = ToplamMesai3;

                        //Boş Satırlar Ekle
                        dt3.Rows.Add();
                        dt3.Rows.Add();

                        SatirEkle3 += 4;
                    }


                    //Gelir Satırları Ekle ve Gelir Toplam Hesapla
                    decimal ToplamGelir3 = 0;
                    decimal Toplam83 = 0;
                    int h3 = SatirEkle3;
                    int GelirSatir3 = 0;
                    for (int i = 0; i < dataGridView2.Rows.Count + 1; i++)
                    {
                        if (i < dataGridView2.Rows.Count && dataGridView2[2, i].Value.ToString() == "Gelir")
                        {
                            dt3.Rows.Add();
                            GelirSatir3 += 1;
                            for (int j = 0; j < dataGridView2.Columns.Count; j++)
                            {
                                if (j == 4)
                                {
                                    dt3.Rows[h3][j + 1] = string.Format("{0:dd/MM/yyyy}", dataGridView2.Rows[i].Cells[j].Value);
                                }
                                else
                                {
                                    dt3.Rows[h3][j + 1] = dataGridView2.Rows[i].Cells[j].Value;
                                }
                            }
                            h3 += 1;
                            Toplam83 += Convert.ToDecimal(dataGridView2.Rows[i].Cells[3].Value);
                            SatirEkle3 += 1;
                        }
                        else
                        {
                            ToplamGelir3 = Toplam83;
                        }
                    }

                    if (GelirSatir3 != 0)
                    {
                        //Boş Satır Ekle
                        dt3.Rows.Add();

                        //Gelir Toplam Satır Ekle
                        DataRow rowGelirToplam3 = dt3.NewRow();
                        dt3.Rows.Add(rowGelirToplam3);

                        rowGelirToplam3[3] = "Toplam :";
                        rowGelirToplam3[4] = ToplamGelir3;

                        //Boş Satırlar Ekle
                        dt3.Rows.Add();
                        dt3.Rows.Add();

                        SatirEkle3 += 4;
                    }


                    //Ödeme Satırları Ekle ve Ödeme Toplam Hesapla
                    decimal ToplamOdeme3 = 0;
                    decimal Toplam93 = 0;
                    int i3 = SatirEkle3;
                    int OdemeSatir3 = 0;
                    for (int i = 0; i < dataGridView2.Rows.Count + 1; i++)
                    {
                        if (i < dataGridView2.Rows.Count && dataGridView2[2, i].Value.ToString() == "Ödeme")
                        {
                            dt3.Rows.Add();
                            OdemeSatir3 += 1;
                            for (int j = 0; j < dataGridView2.Columns.Count; j++)
                            {
                                if (j == 4)
                                {
                                    dt3.Rows[i3][j + 1] = string.Format("{0:dd/MM/yyyy}", dataGridView2.Rows[i].Cells[j].Value);
                                }
                                else
                                {
                                    dt3.Rows[i3][j + 1] = dataGridView2.Rows[i].Cells[j].Value;
                                }
                            }
                            i3 += 1;
                            Toplam93 += Convert.ToDecimal(dataGridView2.Rows[i].Cells[3].Value);
                            SatirEkle3 += 1;
                        }
                        else
                        {
                            ToplamOdeme3 = Toplam93;
                        }
                    }

                    if (OdemeSatir3 != 0)
                    {
                        //Boş Satır Ekle
                        dt3.Rows.Add();

                        //Odeme Toplam Satır Ekle
                        DataRow rowOdemeToplam3 = dt3.NewRow();
                        dt3.Rows.Add(rowOdemeToplam3);

                        rowOdemeToplam3[3] = "Toplam :";
                        rowOdemeToplam3[4] = ToplamOdeme3;

                        SatirEkle3 += 2;
                    }
                    else if (OdemeSatir3 == 0)
                    {
                        dt3.Rows.RemoveAt(SatirEkle3 - 1);
                        dt3.Rows.RemoveAt(SatirEkle3 - 2);
                    }

                    //Gereksiz Sütunları Kaldır
                    dt3.Columns.RemoveAt(1);
                    dt3.Columns.RemoveAt(1);
                    dt3.Columns.RemoveAt(5);
                    dt3.Columns.RemoveAt(5);

                    //Excel Sayfasına 'Liste-1' i ekle.
                    using (workbook)
                    {
                        workbook.Worksheets.Add(dt3, "Liste");
                    }

                    //No Sütunu Gereksiz Noları Sil
                    if (AlacakSatir3 != 0)
                    {
                        workbook.Worksheet("Liste").Cell(AlacakSatir3 + 2, 1).Value = "";
                        workbook.Worksheet("Liste").Cell(AlacakSatir3 + 3, 1).Value = "";
                        workbook.Worksheet("Liste").Cell(AlacakSatir3 + 4, 1).Value = "";
                        workbook.Worksheet("Liste").Cell(AlacakSatir3 + 5, 1).Value = "";

                        workbook.Worksheet("Liste").Cell(AlacakSatir3 + 3, 1).Value = "Alacak";
                    }
                    if (AlindiSatir3 != 0)
                    {
                        workbook.Worksheet("Liste").Cell(b3 + 2, 1).Value = "";
                        workbook.Worksheet("Liste").Cell(b3 + 3, 1).Value = "";
                        workbook.Worksheet("Liste").Cell(b3 + 4, 1).Value = "";
                        workbook.Worksheet("Liste").Cell(b3 + 5, 1).Value = "";
                        workbook.Worksheet("Liste").Cell(b3 + 6, 1).Value = "";

                        for (int i = 1; i < (AlindiSatir3 + 1); i++)
                        {
                            workbook.Worksheet("Liste").Cell((b3 + 1 + i) - AlindiSatir3, 1).Value = i;
                        }

                        workbook.Worksheet("Liste").Cell(b3 + 3, 1).Value = "Alındı";
                    }
                    if (IskontoSatir3 != 0 && AlacakSatir3 != 0)
                    {
                        workbook.Worksheet("Liste").Cell(is3 + 1, 1).Value = "";
                        workbook.Worksheet("Liste").Cell(is3, 1).Value = "";
                    }
                    if (BorcSatir3 != 0)
                    {
                        workbook.Worksheet("Liste").Cell(c3 + 2, 1).Value = "";
                        workbook.Worksheet("Liste").Cell(c3 + 3, 1).Value = "";
                        workbook.Worksheet("Liste").Cell(c3 + 4, 1).Value = "";
                        workbook.Worksheet("Liste").Cell(c3 + 5, 1).Value = "";

                        for (int i = 1; i < (BorcSatir3 + 1); i++)
                        {
                            workbook.Worksheet("Liste").Cell((c3 + 1 + i) - BorcSatir3, 1).Value = i;
                        }

                        workbook.Worksheet("Liste").Cell(c3 + 3, 1).Value = "Borç";
                    }
                    if (OdendiSatir3 != 0)
                    {
                        workbook.Worksheet("Liste").Cell(d3 + 2, 1).Value = "";
                        workbook.Worksheet("Liste").Cell(d3 + 3, 1).Value = "";
                        workbook.Worksheet("Liste").Cell(d3 + 4, 1).Value = "";
                        workbook.Worksheet("Liste").Cell(d3 + 5, 1).Value = "";

                        for (int i = 1; i < (OdendiSatir3 + 1); i++)
                        {
                            workbook.Worksheet("Liste").Cell((d3 + 1 + i) - OdendiSatir3, 1).Value = i;
                        }

                        workbook.Worksheet("Liste").Cell(d3 + 3, 1).Value = "Ödendi";
                    }
                    if (GiderSatir3 != 0)
                    {
                        workbook.Worksheet("Liste").Cell(e3 + 2, 1).Value = "";
                        workbook.Worksheet("Liste").Cell(e3 + 3, 1).Value = "";
                        workbook.Worksheet("Liste").Cell(e3 + 4, 1).Value = "";
                        workbook.Worksheet("Liste").Cell(e3 + 5, 1).Value = "";

                        for (int i = 1; i < (GiderSatir3 + 1); i++)
                        {
                            workbook.Worksheet("Liste").Cell((e3 + 1 + i) - GiderSatir3, 1).Value = i;
                        }

                        workbook.Worksheet("Liste").Cell(e3 + 3, 1).Value = "Gider";
                    }
                    if (BilgiSatir3 != 0)
                    {
                        workbook.Worksheet("Liste").Cell(f3 + 2, 1).Value = "";
                        workbook.Worksheet("Liste").Cell(f3 + 3, 1).Value = "";
                        workbook.Worksheet("Liste").Cell(f3 + 4, 1).Value = "";
                        workbook.Worksheet("Liste").Cell(f3 + 5, 1).Value = "";

                        for (int i = 1; i < (BilgiSatir3 + 1); i++)
                        {
                            workbook.Worksheet("Liste").Cell((f3 + 1 + i) - BilgiSatir3, 1).Value = i;
                        }

                        workbook.Worksheet("Liste").Cell(f3 + 3, 1).Value = "Bilgi";
                    }
                    if (MesaiSatir3 != 0)
                    {
                        workbook.Worksheet("Liste").Cell(g3 + 2, 1).Value = "";
                        workbook.Worksheet("Liste").Cell(g3 + 3, 1).Value = "";
                        workbook.Worksheet("Liste").Cell(g3 + 4, 1).Value = "";
                        workbook.Worksheet("Liste").Cell(g3 + 5, 1).Value = "";

                        for (int i = 1; i < (MesaiSatir3 + 1); i++)
                        {
                            workbook.Worksheet("Liste").Cell((g3 + 1 + i) - MesaiSatir3, 1).Value = i;
                        }

                        workbook.Worksheet("Liste").Cell(g3 + 3, 1).Value = "Mesai";
                    }
                    if (GelirSatir3 != 0)
                    {
                        workbook.Worksheet("Liste").Cell(h3 + 2, 1).Value = "";
                        workbook.Worksheet("Liste").Cell(h3 + 3, 1).Value = "";
                        workbook.Worksheet("Liste").Cell(h3 + 4, 1).Value = "";
                        workbook.Worksheet("Liste").Cell(h3 + 5, 1).Value = "";

                        for (int i = 1; i < (GelirSatir3 + 1); i++)
                        {
                            workbook.Worksheet("Liste").Cell((h3 + 1 + i) - GelirSatir3, 1).Value = i;
                        }

                        workbook.Worksheet("Liste").Cell(h3 + 3, 1).Value = "Gelir";
                    }
                    if (OdemeSatir3 != 0)
                    {
                        workbook.Worksheet("Liste").Cell(i3 + 2, 1).Value = "";
                        workbook.Worksheet("Liste").Cell(i3 + 3, 1).Value = "";

                        for (int i = 1; i < (OdemeSatir3 + 1); i++)
                        {
                            workbook.Worksheet("Liste").Cell((i3 + 1 + i) - OdemeSatir3, 1).Value = i;
                        }

                        workbook.Worksheet("Liste").Cell(i3 + 3, 1).Value = "Ödeme";
                    }

                    //Tekrar Paraları Ayraç İçin Ekle
                    workbook.Worksheet("Liste").Column(3).Style.NumberFormat.NumberFormatId = 4;
                    for (int i = 2; i < dataGridView2.Rows.Count + SatirEkle3 + 2; i++)
                    {
                        workbook.Worksheet("Liste").Cell(i, 3).Value = workbook.Worksheet("Liste").Cell(i, 3).Value;
                    }

                    //Düzenleme Hizalama, Renk vs.
                    workbook.Worksheet("Liste").Row(1).Style.Fill.BackgroundColor = XLColor.White;
                    if (AlacakSatir3 != 0)
                    {
                        workbook.Worksheet("Liste").Row(AlacakSatir3 + 3).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
                        workbook.Worksheet("Liste").Row(AlacakSatir3 + 3).Style.Font.SetBold();
                    }
                    if (AlindiSatir3 != 0)
                    {
                        workbook.Worksheet("Liste").Row(b3 + 3).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
                        workbook.Worksheet("Liste").Row(b3 + 3).Style.Font.SetBold();
                        workbook.Worksheet("Liste").Row(b3 + 4).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
                        workbook.Worksheet("Liste").Row(b3 + 4).Style.Font.SetBold();
                    }

                    if (IskontoSatir3 != 0 && AlacakSatir3 != 0)
                    {
                        workbook.Worksheet("Liste").Row(is3 - 1).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
                        workbook.Worksheet("Liste").Row(is3 - 1).Style.Font.SetBold();
                        workbook.Worksheet("Liste").Row(is3 - 2).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
                        workbook.Worksheet("Liste").Row(is3 - 2).Style.Font.SetBold();
                    }

                    if (BorcSatir3 != 0)
                    {
                        workbook.Worksheet("Liste").Row(c3 + 3).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
                        workbook.Worksheet("Liste").Row(c3 + 3).Style.Font.SetBold();
                    }
                    if (OdendiSatir3 != 0)
                    {
                        workbook.Worksheet("Liste").Row(d3 + 3).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
                        workbook.Worksheet("Liste").Row(d3 + 3).Style.Font.SetBold();
                    }
                    if (GiderSatir3 != 0)
                    {
                        workbook.Worksheet("Liste").Row(e3 + 3).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
                        workbook.Worksheet("Liste").Row(e3 + 3).Style.Font.SetBold();
                    }
                    if (BilgiSatir3 != 0)
                    {
                        workbook.Worksheet("Liste").Row(f3 + 3).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
                        workbook.Worksheet("Liste").Row(f3 + 3).Style.Font.SetBold();
                    }
                    if (MesaiSatir3 != 0)
                    {
                        workbook.Worksheet("Liste").Row(g3 + 3).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
                        workbook.Worksheet("Liste").Row(g3 + 3).Style.Font.SetBold();
                    }
                    if (GelirSatir3 != 0)
                    {
                        workbook.Worksheet("Liste").Row(h3 + 3).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
                        workbook.Worksheet("Liste").Row(h3 + 3).Style.Font.SetBold();
                    }
                    if (OdemeSatir3 != 0)
                    {
                        workbook.Worksheet("Liste").Row(i3 + 3).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
                        workbook.Worksheet("Liste").Row(i3 + 3).Style.Font.SetBold();
                    }
                    //workbook.Worksheet("Liste").Row(SatirEkle3 + 1).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
                    //workbook.Worksheet("Liste").Row(SatirEkle3 + 1).Style.Font.SetBold();
                    workbook.Worksheet("Liste").Row(1).Style.Font.SetFontColor(XLColor.FromArgb(0, 0, 1));
                    workbook.Worksheet("Liste").Cell(1, 1).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
                    workbook.Worksheet("Liste").PageSetup.Header.Left.AddText(lblAd.Text).SetBold();
                    workbook.Worksheet("Liste").PageSetup.Header.Right.AddText(string.Format("{0:dd/MM/yyyy}", DateTime.Now));
                    workbook.Worksheet("Liste").Column(4).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                    workbook.Worksheet("Liste").Column(3).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
                    workbook.Worksheet("Liste").Cell(1, 3).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;

                    //-- Datatable3 - Genel - SON --//

                    //DataTable1 Oluştur
                    DataTable dt = new DataTable();

                    //No Sütunu Ekle
                    DataColumn columno = new DataColumn
                    {
                        DataType = System.Type.GetType("System.Int32"),
                        AutoIncrement = true,
                        AutoIncrementSeed = 1,
                        AutoIncrementStep = 1
                    };

                    dt.Columns.Add(columno);
                    dt.Columns["Column1"].ColumnName = "No";


                    //Sütunları Ekle
                    foreach (DataGridViewColumn column in dataGridView2.Columns)
                    {
                        dt.Columns.Add(column.HeaderText);
                    }


                    //Satırları Ekle
                    foreach (DataGridViewRow row in dataGridView2.Rows)
                    {
                        dt.Rows.Add();
                        //Hücreleri Ekle
                        foreach (DataGridViewCell cell in row.Cells)
                        {
                            if (cell.Value != null)
                            {
                                dt.Rows[dt.Rows.Count - 1][cell.ColumnIndex + 1] = cell.Value;
                            }
                        }
                    }

                    //Gereksiz Sütunları Kaldır
                    dt.Columns.RemoveAt(1);
                    dt.Columns.RemoveAt(1);
                    dt.Columns.RemoveAt(5);
                    dt.Columns.RemoveAt(5);

                    //Tarih Sütunu Zamanı Kaldır
                    dt.Columns.Add("Tarihler", typeof(String));
                    for (int i = 4; i < dataGridView2.Columns.Count - 3; i++)
                    {
                        for (int j = 0; j < dataGridView2.Rows.Count; j++)
                        {
                            dt.Rows[j][5] = string.Format("{0:dd/MM/yyyy}", dataGridView2.Rows[j].Cells[i].Value);
                        }
                    }
                    int columnNumber = dt.Columns["Tarih"].Ordinal;
                    dt.Columns.Remove("Tarih");
                    dt.Columns["Tarihler"].SetOrdinal(columnNumber);
                    dt.Columns["Tarihler"].ColumnName = "Tarih";

                    /*
                    //Para Ayracı Ekle - Eski
                    dt.Columns.Add("Tutarlar", typeof(String));
                    for (int i = 3; i < dataGridView2.Columns.Count - 4; i++)
                    {
                        for (int j = 0; j < dataGridView2.Rows.Count; j++)
                        {
                            dt.Rows[j][5] = string.Format("{0:N}\n", dataGridView2.Rows[j].Cells[i].Value);
                        }
                    }
                    int columnNumber2 = dt.Columns["Tutar"].Ordinal;
                    dt.Columns.Remove("Tutar");
                    dt.Columns["Tutarlar"].SetOrdinal(columnNumber2);
                    dt.Columns["Tutarlar"].ColumnName = "Tutar";
                    */

                    /*
                    //Alacak Toplam Hesapla-Eski
                    for (int i = 3; i < dataGridView2.Columns.Count - 4; i++)
                    {
                         for (int j = 0; j < dataGridView2.Rows.Count; j++)
                         {
                            if (dataGridView2[3, j].Value.ToString() == "Gider")
                            { 
                                GiderT += Convert.ToDecimal(dataGridView2.Rows[j].Cells[i].Value);
                            }
                         }   
                    }
                    */
                    int EklenecekSatir = 0;
                    dt.Rows.Add();
                    //Alacak Toplam Hesapla
                    decimal AlacakT = 0;
                    for (int i = 3; i < 4; i++)
                    {
                        decimal toplam = 0;
                        for (int j = 0; j < dataGridView2.Rows.Count + 1; ++j)
                        {
                            if (j < dataGridView2.Rows.Count && dataGridView2[2, j].Value.ToString() == "Alacak")
                            {
                                toplam += Convert.ToDecimal(dataGridView2.Rows[j].Cells[i].Value);
                            }
                            else// if (j == dataGridView2.Rows.Count && dataGridView2[2, j-1].Value.ToString() == "Alacak")
                            {
                                AlacakT = toplam;
                            }
                        }
                    }
                    //Alacak Toplam Satırı Ekle
                    if (AlacakT != 0)
                    {
                        DataRow rowAlacak = dt.NewRow();
                        dt.Rows.Add(rowAlacak);
                        rowAlacak[1] = "Alacağım :";
                        rowAlacak[2] = AlacakT;
                        EklenecekSatir += 1;
                    }


                    //Alındı Toplam Hesapla
                    decimal AlindiT = 0;
                    for (int i = 3; i < 4; i++)
                    {
                        decimal toplam = 0;
                        for (int j = 0; j < dataGridView2.Rows.Count + 1; ++j)
                        {
                            if (j < dataGridView2.Rows.Count && dataGridView2[2, j].Value.ToString() == "Alındı")
                            {
                                toplam += Convert.ToDecimal(dataGridView2.Rows[j].Cells[i].Value);
                            }
                            else
                            {
                                AlindiT = toplam;
                            }
                        }
                    }
                    //Alındı Toplam Satırı Ekle
                    if (AlindiT != 0)
                    {
                        DataRow rowAlindi = dt.NewRow();
                        dt.Rows.Add(rowAlindi);
                        rowAlindi[1] = "Alındı :";
                        rowAlindi[2] = AlindiT;
                        EklenecekSatir += 1;
                    }


                    //Borç Toplam Hesapla
                    decimal BorcT = 0;
                    for (int i = 3; i < 4; i++)
                    {
                        decimal toplam = 0;
                        for (int j = 0; j < dataGridView2.Rows.Count + 1; ++j)
                        {
                            if (j < dataGridView2.Rows.Count && dataGridView2[2, j].Value.ToString() == "Borç")
                            {
                                toplam += Convert.ToDecimal(dataGridView2.Rows[j].Cells[i].Value);
                            }
                            else
                            {
                                BorcT = toplam;
                            }
                        }
                    }
                    //Borç Toplam Satırı Ekle
                    if (BorcT != 0)
                    {
                        DataRow rowBorc = dt.NewRow();
                        dt.Rows.Add(rowBorc);
                        rowBorc[1] = "Borç :";
                        rowBorc[2] = BorcT;
                        EklenecekSatir += 1;
                    }


                    //Ödendi Toplam Hesapla
                    decimal OdendiT = 0;
                    for (int i = 3; i < 4; i++)
                    {
                        decimal toplam = 0;
                        for (int j = 0; j < dataGridView2.Rows.Count + 1; ++j)
                        {
                            if (j < dataGridView2.Rows.Count && dataGridView2[2, j].Value.ToString() == "Ödendi")
                            {
                                toplam += Convert.ToDecimal(dataGridView2.Rows[j].Cells[i].Value);
                            }
                            else
                            {
                                OdendiT = toplam;
                            }
                        }
                    }
                    //Ödendi Toplam Satırı Ekle
                    if (OdendiT != 0)
                    {
                        DataRow rowOdendi = dt.NewRow();
                        dt.Rows.Add(rowOdendi);
                        rowOdendi[1] = "Ödendi :";
                        rowOdendi[2] = OdendiT;
                        EklenecekSatir += 1;
                    }


                    //Gider Toplam Hesapla
                    decimal GiderT = 0;
                    for (int i = 3; i < 4; i++)
                    {
                        decimal toplam = 0;
                        for (int j = 0; j < dataGridView2.Rows.Count + 1; ++j)
                        {
                            if (j < dataGridView2.Rows.Count && dataGridView2[2, j].Value.ToString() == "Gider")
                            {
                                toplam += Convert.ToDecimal(dataGridView2.Rows[j].Cells[i].Value);
                            }
                            else
                            {
                                GiderT = toplam;
                            }
                        }
                    }
                    //Gider Toplam Satırı Ekle
                    if (GiderT != 0)
                    {
                        DataRow rowGider = dt.NewRow();
                        dt.Rows.Add(rowGider);
                        rowGider[1] = "Gider :";
                        rowGider[2] = GiderT;
                        EklenecekSatir += 1;
                    }


                    //Bilgi Toplam Hesapla
                    decimal BilgiT = 0;
                    for (int i = 3; i < 4; i++)
                    {
                        decimal toplam = 0;
                        for (int j = 0; j < dataGridView2.Rows.Count + 1; ++j)
                        {
                            if (j < dataGridView2.Rows.Count && dataGridView2[2, j].Value.ToString() == "Bilgi")
                            {
                                toplam += Convert.ToDecimal(dataGridView2.Rows[j].Cells[i].Value);
                            }
                            else
                            {
                                BilgiT = toplam;
                            }
                        }
                    }
                    //Bilgi Toplam Satırı Ekle
                    if (BilgiT != 0)
                    {
                        DataRow rowBilgi = dt.NewRow();
                        dt.Rows.Add(rowBilgi);
                        rowBilgi[1] = "Bilgi :";
                        rowBilgi[2] = BilgiT;
                        EklenecekSatir += 1;
                    }


                    //Mesai Toplam Hesapla
                    decimal MesaiT = 0;
                    for (int i = 3; i < 4; i++)
                    {
                        decimal toplam = 0;
                        for (int j = 0; j < dataGridView2.Rows.Count + 1; ++j)
                        {
                            if (j < dataGridView2.Rows.Count && dataGridView2[2, j].Value.ToString() == "Mesai")
                            {
                                toplam += Convert.ToDecimal(dataGridView2.Rows[j].Cells[i].Value);
                            }
                            else
                            {
                                MesaiT = toplam;
                            }
                        }
                    }
                    //Mesai Toplam Satırı Ekle
                    if (MesaiT != 0)
                    {
                        DataRow rowMesai = dt.NewRow();
                        dt.Rows.Add(rowMesai);
                        rowMesai[1] = "Mesai :";
                        rowMesai[2] = MesaiT;
                        EklenecekSatir += 1;
                    }


                    //Gelir Toplam Hesapla
                    decimal GelirT = 0;
                    for (int i = 3; i < 4; i++)
                    {
                        decimal toplam = 0;
                        for (int j = 0; j < dataGridView2.Rows.Count + 1; ++j)
                        {
                            if (j < dataGridView2.Rows.Count && dataGridView2[2, j].Value.ToString() == "Gelir")
                            {
                                toplam += Convert.ToDecimal(dataGridView2.Rows[j].Cells[i].Value);
                            }
                            else
                            {
                                GelirT = toplam;
                            }
                        }
                    }
                    //Gelir Toplam Satırı Ekle
                    if (GelirT != 0)
                    {
                        DataRow rowGelir = dt.NewRow();
                        dt.Rows.Add(rowGelir);
                        rowGelir[1] = "Gelir :";
                        rowGelir[2] = GelirT;
                        EklenecekSatir += 1;
                    }


                    //İskonto Toplam Hesapla
                    decimal IskontoT = 0;
                    for (int i = 3; i < 4; i++)
                    {
                        decimal toplam = 0;
                        for (int j = 0; j < dataGridView2.Rows.Count + 1; ++j)
                        {
                            if (j < dataGridView2.Rows.Count && dataGridView2[2, j].Value.ToString() == "İskonto")
                            {
                                toplam += Convert.ToDecimal(dataGridView2.Rows[j].Cells[i].Value);
                            }
                            else
                            {
                                IskontoT = toplam;
                            }
                        }
                    }
                    //İskonto Toplam Satırı Ekle
                    if (IskontoT != 0)
                    {
                        DataRow rowIskonto = dt.NewRow();
                        dt.Rows.Add(rowIskonto);
                        rowIskonto[1] = "İskonto :";
                        rowIskonto[2] = IskontoT;
                        EklenecekSatir += 1;
                    }


                    //Ödeme Toplam Hesapla
                    decimal OdemeT = 0;
                    for (int i = 3; i < 4; i++)
                    {
                        decimal toplam = 0;
                        for (int j = 0; j < dataGridView2.Rows.Count + 1; ++j)
                        {
                            if (j < dataGridView2.Rows.Count && dataGridView2[2, j].Value.ToString() == "Ödeme")
                            {
                                toplam += Convert.ToDecimal(dataGridView2.Rows[j].Cells[i].Value);
                            }
                            else
                            {
                                OdemeT = toplam;
                            }
                        }
                    }
                    //Ödeme Toplam Satırı Ekle
                    if (OdemeT != 0)
                    {
                        DataRow rowOdeme = dt.NewRow();
                        dt.Rows.Add(rowOdeme);
                        rowOdeme[1] = "Ödeme :";
                        rowOdeme[2] = OdemeT;
                        EklenecekSatir += 1;
                    }
                    /*
                    DataRow rowBorc = dt.NewRow();
                    dt.Rows.Add(rowBorc);
                    rowBorc[1] = "Borcum :";
                    rowBorc[2] = LblKalanBorc.Text;
                    */
                    //Excel Sayfasına Kişinin İsmini ekle.
                    using (workbook)
                    {
                        workbook.Worksheets.Add(dt, lblAd.Text);
                    }

                    //Düzenleme No Sütunu En Altlar Sil, Toplam Font Bold Yap, vs..
                    workbook.Worksheet(lblAd.Text).Column(3).Style.NumberFormat.NumberFormatId = 4;

                    for (int i = 0; i < EklenecekSatir; i++)
                    {
                        workbook.Worksheet(lblAd.Text).Row(dataGridView2.Rows.Count + i + 3).Style.Font.SetBold();
                    }

                    workbook.Worksheet(lblAd.Text).Row(1).Style.Fill.BackgroundColor = XLColor.White;
                    workbook.Worksheet(lblAd.Text).Row(1).Style.Font.SetBold();
                    workbook.Worksheet(lblAd.Text).Row(1).Style.Font.SetFontColor(XLColor.FromArgb(0, 0, 1));
                    workbook.Worksheet(lblAd.Text).Cell(1, 1).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
                    workbook.Worksheet(lblAd.Text).PageSetup.Header.Left.AddText(lblAd.Text).SetBold();
                    workbook.Worksheet(lblAd.Text).PageSetup.Header.Right.AddText(string.Format("{0:dd/MM/yyyy}", DateTime.Now));
                    workbook.Worksheet(lblAd.Text).Column(4).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                    workbook.Worksheet(lblAd.Text).Column(3).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
                    workbook.Worksheet(lblAd.Text).Cell(1, 3).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;

                    for (int i = 0; i < EklenecekSatir + 1; i++)
                    {
                        workbook.Worksheet(lblAd.Text).Cell(dataGridView2.Rows.Count + i + 2, 1).Value = "";
                    }

                    //Değerler tekrar atanarak sayı hale getiriliyor..
                    //workbook.Worksheet(lblAd.Text).Cell(2, 3).Value = workbook.Worksheet(lblAd.Text).Cell(2, 3).Value;
                    for (int i = 2; i < dataGridView2.Rows.Count + EklenecekSatir + 3; i++)
                    {
                        workbook.Worksheet(lblAd.Text).Cell(i, 3).Value = workbook.Worksheet(lblAd.Text).Cell(i, 3).Value;
                    }

                    //Satır Arkaplan Renkleri Ata
                    /*
                    for (int i = 3; i < dataGridView2.Rows.Count + 2; i += 2)
                    {
                        workbook.Worksheet(lblAd.Text).Row(i).Style.Fill.BackgroundColor = XLColor.LightGray;
                    }
                    */
                    /*
                    for (int i = 2; i < dataGridView2.Rows.Count + 4; i += 2)
                    {
                        workbook.Worksheet(lblAd.Text).Row(i).Style.Fill.BackgroundColor = XLColor.White;
                    }
                    */
                    //workbook.Worksheet(lblAd.Text).Row(dataGridView2.Rows.Count + 2).Style.Fill.BackgroundColor = XLColor.White;
                    /*
                    int ListeSatirSayi = workbook.Worksheet("Liste").RowCount();
                    int bolum1 = ListeSatirSayi / 60;
                    int kalan1 = ListeSatirSayi % 60;
                    if (kalan1 != 0)
                    {
                        bolum1 = bolum1 + 1;
                    }
                    */
                    workbook.Worksheet("Liste").Columns().AdjustToContents();
                    if (workbook.Worksheet("Liste").Column(5).Width < 49)
                    {
                        workbook.Worksheet("Liste").PageSetup.AdjustTo(100);
                    }
                    else if (workbook.Worksheet("Liste").Column(5).Width < 59 && workbook.Worksheet("Liste").Column(5).Width >= 49)
                    {
                        workbook.Worksheet("Liste").PageSetup.AdjustTo(90);
                    }
                    else if (workbook.Worksheet("Liste").Column(5).Width >= 59 && workbook.Worksheet("Liste").Column(5).Width < 70)
                    {
                        workbook.Worksheet("Liste").PageSetup.AdjustTo(80);
                    }
                    else if (workbook.Worksheet("Liste").Column(5).Width >= 70 && workbook.Worksheet("Liste").Column(5).Width < 98)
                    {
                        workbook.Worksheet("Liste").PageSetup.FitToPages(1, 2);
                        workbook.Worksheet("Liste").PageSetup.PageOrientation = XLPageOrientation.Landscape;
                    }
                    else if (workbook.Worksheet("Liste").Column(5).Width >= 98)
                    {
                        workbook.Worksheet("Liste").PageSetup.PageOrientation = XLPageOrientation.Landscape;
                        workbook.Worksheet("Liste").PageSetup.FitToPages(1, 3);
                    }
                    //workbook.Worksheet("Liste").PageSetup.PrintAreas.Add(1, 1, ListeSatirSayi, 5);
                    //workbook.Worksheet("Liste").PageSetup.AddVerticalPageBreak(bolum1);

                    workbook.Worksheet(lblAd.Text).Columns().AdjustToContents();
                    if (workbook.Worksheet(lblAd.Text).Column(5).Width < 49)
                    {
                        workbook.Worksheet(lblAd.Text).PageSetup.AdjustTo(100);
                    }
                    else if (workbook.Worksheet(lblAd.Text).Column(5).Width < 59 && workbook.Worksheet(lblAd.Text).Column(5).Width >= 49)
                    {
                        workbook.Worksheet(lblAd.Text).PageSetup.AdjustTo(90);
                    }
                    else if (workbook.Worksheet(lblAd.Text).Column(5).Width >= 59 && workbook.Worksheet(lblAd.Text).Column(5).Width < 70)
                    {
                        workbook.Worksheet(lblAd.Text).PageSetup.AdjustTo(80);
                    }
                    else if (workbook.Worksheet(lblAd.Text).Column(5).Width >= 70 && workbook.Worksheet(lblAd.Text).Column(5).Width < 98)
                    {
                        workbook.Worksheet(lblAd.Text).PageSetup.FitToPages(1, 2);
                        workbook.Worksheet(lblAd.Text).PageSetup.PageOrientation = XLPageOrientation.Landscape;
                    }
                    else if (workbook.Worksheet(lblAd.Text).Column(5).Width >= 98)
                    {
                        workbook.Worksheet(lblAd.Text).PageSetup.PageOrientation = XLPageOrientation.Landscape;
                        workbook.Worksheet(lblAd.Text).PageSetup.FitToPages(1, 3);
                    }

                    if (lblAd.Text == "FATURALAR")
                    {
                        workbook.Worksheet("Liste-1").Columns().AdjustToContents();
                        if (workbook.Worksheet("Liste-1").Column(5).Width < 49)
                        {
                            workbook.Worksheet("Liste-1").PageSetup.AdjustTo(100);
                        }
                        else if (workbook.Worksheet("Liste-1").Column(5).Width < 59 && workbook.Worksheet("Liste-1").Column(5).Width >= 49)
                        {
                            workbook.Worksheet("Liste-1").PageSetup.AdjustTo(90);
                        }
                        else if (workbook.Worksheet("Liste-1").Column(5).Width >= 59 && workbook.Worksheet("Liste-1").Column(5).Width < 70)
                        {
                            workbook.Worksheet("Liste-1").PageSetup.AdjustTo(80);
                        }
                        else if (workbook.Worksheet("Liste-1").Column(5).Width >= 70 && workbook.Worksheet("Liste-1").Column(5).Width < 98)
                        {
                            workbook.Worksheet("Liste-1").PageSetup.FitToPages(1, 2);
                            workbook.Worksheet("Liste-1").PageSetup.PageOrientation = XLPageOrientation.Landscape;
                        }
                        else if (workbook.Worksheet("Liste-1").Column(5).Width >= 98)
                        {
                            workbook.Worksheet("Liste-1").PageSetup.PageOrientation = XLPageOrientation.Landscape;
                            workbook.Worksheet("Liste-1").PageSetup.FitToPages(1, 3);
                        }
                    }


                    //Kaydet
                    do
                    {
                        try
                        {
                            workbook.SaveAs(fileName);
                            MessageBox.Show("Excel dosyası kaydedildi.", "Bilgi", MessageBoxButtons.OK, MessageBoxIcon.Information);
                            //tabControl1.TabPages.Remove(tabPageHesaplama);
                            customCulture.NumberFormat.NumberDecimalSeparator = ",";
                            System.Threading.Thread.CurrentThread.CurrentCulture = customCulture;
                            break;
                        }
                        catch (System.IO.IOException)
                        {
                            MessageBox.Show("Kayıt yapılamadı! Kaydetmeye çalıştığınız dosya açık olabilir.", "Uyarı", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                            //tabControl1.TabPages.Remove(tabPageHesaplama);
                            customCulture.NumberFormat.NumberDecimalSeparator = ",";
                            System.Threading.Thread.CurrentThread.CurrentCulture = customCulture;
                            break;
                        }
                    } while (true);
                }
            }   
        }

        private void TxtDaAra_TextChanged(object sender, EventArgs e)
        {
            if (TxtDaAra.Text.Length > 1)
            {
                TxtDaAra.ForeColor = Color.Blue;
                var sonuc = from kisiler in ctx.Kisis
                            join cariler in ctx.Caris on kisiler.KisiID equals cariler.KisiID
                            join durumlar in ctx.Durums on cariler.DurumID equals durumlar.DurumID
                            orderby cariler.Tarih descending
                            select new
                            {
                                kisiler.KisiID,
                                kisiler.Ad,
                                kisiler.Firma,
                                kisiler.Tel1,
                                kisiler.Tel2,
                                kisiler.Adres,
                                Durum = durumlar.Durumlar,
                                cariler.Tutar,
                                cariler.Tarih,
                                Açıklama = cariler.Aciklama,
                                durumlar.DurumID,
                                cariler.CariID
                            };

                DataGridViewDa.DataSource = sonuc.Where(x => x.Ad.Contains(TxtDaAra.Text) || x.Firma.Contains(TxtDaAra.Text) || x.Durum.Contains(TxtDaAra.Text) || x.Açıklama.Contains(TxtDaAra.Text) || Convert.ToString(x.Tutar).Contains(TxtDaAra.Text));

                DataGridViewDa.Columns["KisiID"].Visible = DataGridViewDa.Columns["Ad"].Visible = DataGridViewDa.Columns["Firma"].Visible = DataGridViewDa.Columns["Tel1"].Visible = DataGridViewDa.Columns["Tel2"].Visible = DataGridViewDa.Columns["Adres"].Visible = DataGridViewDa.Columns["DurumID"].Visible = DataGridViewDa.Columns["CariID"].Visible = false;

                DataGridViewDa.Columns[3].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                DataGridViewDa.Columns[3].HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleRight;
                DataGridViewDa.Columns[4].HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
                DataGridViewDa.Columns[4].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            }
            else
            {
                TxtDaAra.ForeColor = Color.Red;
            }
        }


        private void DataGridViewDa_DataBindingComplete(object sender, DataGridViewBindingCompleteEventArgs e)
        {
            /*
            for (int i = 0; i < DataGridViewDa.Columns.Count; i++)
            {
                DataGridViewDa.Columns[i].SortMode = DataGridViewColumnSortMode.NotSortable;
            }
            */
            foreach (DataGridViewRow dGVRow in this.DataGridViewDa.Rows)
            {
                dGVRow.HeaderCell.Value = String.Format("{0}", dGVRow.Index + 1);
            }
            //Genişlik Ayarı
            this.DataGridViewDa.AutoResizeRowHeadersWidth(DataGridViewRowHeadersWidthSizeMode.AutoSizeToAllHeaders);
            
        }

        private void DataGridViewDa_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            anahtar2 = true;
            if (anahtar2)
            {
                DataGridViewRow row = DataGridViewDa.CurrentRow;

                label57.Visible = true;
                LblDaId.Visible = true;
                LblDaId.Text = row.Cells["KisiID"].Value.ToString();
                KisiIdDetay = (int)row.Cells["KisiID"].Value;

                if ((string)row.Cells["Ad"].Value == "" || row.Cells["Ad"].Value == null)
                {
                    LblDaAd.Visible = false;
                    label47.Visible = false;
                }
                else
                {
                    label47.Visible = true;
                    LblDaAd.Visible = true;
                    KisiAdDetay = LblDaAd.Text = row.Cells["Ad"].Value.ToString();
                }

                if ((string)row.Cells["Firma"].Value == "" || row.Cells["Firma"].Value == null)
                {
                    LblDaFirma.Visible = false;
                    label40.Visible = false;
                    KisiFirmaDetay = null;
                }
                else
                {
                    label40.Visible = true;
                    LblDaFirma.Visible = true;
                    KisiFirmaDetay = LblDaFirma.Text = row.Cells["Firma"].Value.ToString();
                }

                if ((string)row.Cells["Tel1"].Value == "" || row.Cells["Tel1"].Value == null)
                {
                    label45.Visible = false;
                    LblDaTel1.Visible = false;
                    KisiTel1Detay = null;
                }
                else
                {
                    label45.Visible = true;
                    LblDaTel1.Visible = true;
                    KisiTel1Detay = LblDaTel1.Text = row.Cells["Tel1"].Value.ToString();
                }

                if ((string)row.Cells["Tel2"].Value == "" || row.Cells["Tel2"].Value == null)
                {
                    label41.Visible = false;
                    LblDaTel2.Visible = false;
                    KisiTel2Detay = null;
                }
                else
                {
                    label41.Visible = true;
                    LblDaTel2.Visible = true;
                    KisiTel2Detay = LblDaTel2.Text = row.Cells["Tel2"].Value.ToString();
                }

                if ((string)row.Cells["Adres"].Value == "" || row.Cells["Adres"].Value == null)
                {
                    label43.Visible = false;
                    LblDaAdres.Visible = false;
                    KisiAdresDetay = null;
                }
                else
                {
                    label43.Visible = true;
                    LblDaAdres.Visible = true;
                    KisiAdresDetay = LblDaAdres.Text = row.Cells["Adres"].Value.ToString();
                }
            }
        }

        private void DataGridViewDa_SelectionChanged(object sender, EventArgs e)
        {
            if (anahtar2)
            {
                DataGridViewRow row = DataGridViewDa.CurrentRow;

                if ((string)row.Cells["Ad"].Value == "" || row.Cells["Ad"].Value == null)
                {
                    LblDaAd.Visible = false;
                    label47.Visible = false;
                }
                else
                {
                    label47.Visible = true;
                    LblDaAd.Visible = true;
                    LblDaAd.Text = row.Cells["Ad"].Value.ToString();
                }

                if ((string)row.Cells["Firma"].Value == "" || row.Cells["Firma"].Value == null)
                {
                    LblDaFirma.Visible = false;
                    label40.Visible = false;
                }
                else
                {
                    label40.Visible = true;
                    LblDaFirma.Visible = true;
                    LblDaFirma.Text = row.Cells["Firma"].Value.ToString();
                }

                if ((string)row.Cells["Tel1"].Value == "" || row.Cells["Tel1"].Value == null)
                {
                    label45.Visible = false;
                    LblDaTel1.Visible = false;
                }
                else
                {
                    label45.Visible = true;
                    LblDaTel1.Visible = true;
                    LblDaTel1.Text = row.Cells["Tel1"].Value.ToString();
                }

                if ((string)row.Cells["Tel2"].Value == "" || row.Cells["Tel2"].Value == null)
                {
                    label41.Visible = false;
                    LblDaTel2.Visible = false;
                }
                else
                {
                    label41.Visible = true;
                    LblDaTel2.Visible = true;
                    LblDaTel2.Text = row.Cells["Tel2"].Value.ToString();
                }

                if ((string)row.Cells["Adres"].Value == "" || row.Cells["Adres"].Value == null)
                {
                    label43.Visible = false;
                    LblDaAdres.Visible = false;
                }
                else
                {
                    label43.Visible = true;
                    LblDaAdres.Visible = true;
                    LblDaAdres.Text = row.Cells["Adres"].Value.ToString();
                }
            }      
        }

        private void tabControl1_SelectedIndexChanged(object sender, EventArgs e)
        {
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
            var alacak = hesap.Where(x => x.cKisiID == Convert.ToInt32(lblID.Tag)).Where(x => x.DurumID == 1).Sum(x => x.Tutar);
            var alindi = hesap.Where(x => x.cKisiID == Convert.ToInt32(lblID.Tag)).Where(x => x.DurumID == 2).Sum(x => x.Tutar);
            var iskonto = hesap.Where(x => x.cKisiID == Convert.ToInt32(lblID.Tag)).Where(x => x.DurumID == 9).Sum(x => x.Tutar);

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


            if (CmbYil.SelectedIndex == DateTime.Now.Year - 2017 + 1)//Tümü
            {
                var gider = hesap.Where(x => x.cKisiID == Convert.ToInt32(lblID.Tag)).Where(x => x.DurumID == (CmbBilgi.SelectedIndex + 1)).Sum(x => x.Tutar);


                if (gider != null)
                {
                    lblGider.Text = String.Format("{0:N}\n", gider);
                }
                else
                {
                    lblGider.Text = "0,00";
                }
            }
            else //Yıllara Göre
            {
                if (CmbGider.SelectedIndex == 0)//Genel
                {
                    var gider = hesap.Where(x => x.cKisiID == Convert.ToInt32(lblID.Tag)).Where(x => x.DurumID == (CmbBilgi.SelectedIndex + 1) && (x.Tarih.Value.Year == Convert.ToInt32(CmbYil.SelectedItem))).Sum(x => x.Tutar);


                    if (gider != null)
                    {
                        lblGider.Text = String.Format("{0:N}\n", gider);
                    }
                    else
                    {
                        lblGider.Text = "0,00";
                    }
                }
                else if (CmbGider.SelectedIndex == 1 || CmbGider.SelectedIndex == 2 || CmbGider.SelectedIndex == 3 || CmbGider.SelectedIndex == 4 || CmbGider.SelectedIndex == 5 || CmbGider.SelectedIndex == 6 || CmbGider.SelectedIndex == 7 || CmbGider.SelectedIndex == 8 || CmbGider.SelectedIndex == 9 || CmbGider.SelectedIndex == 10 || CmbGider.SelectedIndex == 11 || CmbGider.SelectedIndex == 12)//Aylar
                {
                    int Yil = Convert.ToInt32(CmbYil.SelectedItem);
                    var AyinIlkGunu = new DateTime(Yil, CmbGider.SelectedIndex, 1);
                    var AyinSonGunu = AyinIlkGunu.AddMonths(1).AddDays(-1);

                    var gider = hesap.Where(x => x.cKisiID == Convert.ToInt32(lblID.Tag)).Where(x => x.DurumID == (CmbBilgi.SelectedIndex + 1) && (x.Tarih.Value >= AyinIlkGunu && x.Tarih.Value <= AyinSonGunu)).Sum(x => x.Tutar);


                    if (gider != null)
                    {
                        lblGider.Text = String.Format("{0:N}\n", gider);
                    }
                    else
                    {
                        lblGider.Text = "0,00";
                    }
                }
            }

            //Borç
            var borc = hesap.Where(x => x.cKisiID == Convert.ToInt32(lblID.Tag)).Where(x => x.DurumID == 3).Sum(x => x.Tutar);
            var odendi = hesap.Where(x => x.cKisiID == Convert.ToInt32(lblID.Tag)).Where(x => x.DurumID == 4).Sum(x => x.Tutar);

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

        private void DataGridViewDa_ColumnAdded(object sender, DataGridViewColumnEventArgs e)
        {
            DataGridViewDa.Columns[e.Column.Index].SortMode = DataGridViewColumnSortMode.NotSortable;
        }

        private void CmbGider_SelectedIndexChanged(object sender, EventArgs e)
        {
            decimal gider = 0;
            /*
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
            */
            if (CmbYil.SelectedIndex == DateTime.Now.Year - 2017 + 1)//Tümü Seçiliyse
            {
                CmbGider.SelectedIndex = 0;
                /*
                var gider = hesap.Where(x => x.cKisiID == Convert.ToInt32(lblID.Tag)).Where(x => x.DurumID == (CmbBilgi.SelectedIndex + 1)).Sum(x => x.Tutar);
                */
                for (int i = 3; i < 4; i++)
                {
                    decimal toplam = 0;
                    for (int j = 0; j < dataGridView2.Rows.Count + 1; ++j)
                    {
                        if (j < dataGridView2.Rows.Count && dataGridView2[2, j].Value.ToString() == CmbBilgi.SelectedItem.ToString())
                        {
                            toplam += Convert.ToDecimal(dataGridView2.Rows[j].Cells[i].Value);
                        }
                        else if (j == dataGridView2.Rows.Count)
                        {
                            gider = toplam;
                        }
                    }
                }
                if (gider != 0)
                {
                    lblGider.Text = String.Format("{0:N}\n", gider);
                }
                else
                {
                    lblGider.Text = "0,00";
                }
            }
            else //Yıllara Göre
            {
                if (CmbGider.SelectedIndex == 0)//Genel
                {
                    //var gider = hesap.Where(x => x.cKisiID == Convert.ToInt32(lblID.Tag)).Where(x => x.DurumID == (CmbBilgi.SelectedIndex + 1) && (x.Tarih.Value.Year == Convert.ToInt32(CmbYil.SelectedItem))).Sum(x => x.Tutar);
                    for (int i = 3; i < 4; i++)
                    {
                        decimal toplam = 0;
                        for (int j = 0; j < dataGridView2.Rows.Count + 1; ++j)
                        {
                            if (j < dataGridView2.Rows.Count && dataGridView2[2, j].Value.ToString() == CmbBilgi.SelectedItem.ToString() && ((DateTime)dataGridView2[4, j].Value).Year == Convert.ToInt32(CmbYil.SelectedItem))
                            {
                                toplam += Convert.ToDecimal(dataGridView2.Rows[j].Cells[i].Value);
                            }
                            else if (j == dataGridView2.Rows.Count)
                            {
                                gider = toplam;
                            }
                        }
                    }

                    if (gider != 0)
                    {
                        lblGider.Text = String.Format("{0:N}\n", gider);
                    }
                    else
                    {
                        lblGider.Text = "0,00";
                    }
                }
                else//Aylar
                {
                    int Yil = Convert.ToInt32(CmbYil.SelectedItem);
                    var AyinIlkGunu = new DateTime(Yil, CmbGider.SelectedIndex, 1);
                    var AyinSonGunu = AyinIlkGunu.AddMonths(1).AddDays(-1);

                    //var gider = hesap.Where(x => x.cKisiID == Convert.ToInt32(lblID.Tag)).Where(x => x.DurumID == (CmbBilgi.SelectedIndex + 1) && (x.Tarih.Value >= AyinIlkGunu && x.Tarih.Value <= AyinSonGunu)).Sum(x => x.Tutar);
                    for (int i = 3; i < 4; i++)
                    {
                        decimal toplam = 0;
                        for (int j = 0; j < dataGridView2.Rows.Count + 1; ++j)
                        {
                            if (j < dataGridView2.Rows.Count && dataGridView2[2, j].Value.ToString() == CmbBilgi.SelectedItem.ToString() && (DateTime)dataGridView2[4, j].Value >= AyinIlkGunu && (DateTime)dataGridView2[4, j].Value <= AyinSonGunu)
                            {
                                toplam += Convert.ToDecimal(dataGridView2.Rows[j].Cells[i].Value);
                            }
                            else if (j == dataGridView2.Rows.Count)
                            {
                                gider = toplam;
                            }
                        }
                    }

                    if (gider != 0)
                    {
                        lblGider.Text = String.Format("{0:N}\n", gider);
                    }
                    else
                    {
                        lblGider.Text = "0,00";
                    }
                }
            }
        }

        private void TxtAraDuzelt_TextChanged(object sender, EventArgs e)
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
            dataGridView3.DataSource = sonuc.Where(x => x.KisiID == Convert.ToInt32(lblID.Tag) && (x.Açıklama.Contains(TxtAraDuzelt.Text) || x.Durum.Contains(TxtAraDuzelt.Text) || Convert.ToString(x.Tutar).Contains(TxtAraDuzelt.Text)));
            dataGridView3.Columns["KisiID"].Visible = dataGridView3.Columns["Ad"].Visible = dataGridView3.Columns["DurumID"].Visible = dataGridView3.Columns["CariID"].Visible = false;
        }

        private void CmbBilgi_SelectedIndexChanged(object sender, EventArgs e)
        {
            decimal gider = 0;
            /*
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
            */
            if (CmbYil.SelectedIndex == DateTime.Now.Year - 2017 + 1)//Tümü
            {
                //var gider = hesap.Where(x => x.cKisiID == Convert.ToInt32(lblID.Tag)).Where(x => x.DurumID == (CmbBilgi.SelectedIndex + 1)).Sum(x => x.Tutar);
                for (int i = 3; i < 4; i++)
                {
                    decimal toplam = 0;
                    for (int j = 0; j < dataGridView2.Rows.Count + 1; ++j)
                    {
                        if (j < dataGridView2.Rows.Count && dataGridView2[2, j].Value.ToString() == CmbBilgi.SelectedItem.ToString())
                        {
                            toplam += Convert.ToDecimal(dataGridView2.Rows[j].Cells[i].Value);
                        }
                        else if (j == dataGridView2.Rows.Count)
                        {
                            gider = toplam;
                        }
                    }
                }

                if (gider != 0)
                {
                    lblGider.Text = String.Format("{0:N}\n", gider);
                }
                else
                {
                    lblGider.Text = "0,00";
                }
            }
            else //Yıllara Göre
            {
                if (CmbGider.SelectedIndex == 0)//Genel
                {
                    //var gider = hesap.Where(x => x.cKisiID == Convert.ToInt32(lblID.Tag)).Where(x => x.DurumID == (CmbBilgi.SelectedIndex + 1) && (x.Tarih.Value.Year == Convert.ToInt32(CmbYil.SelectedItem))).Sum(x => x.Tutar);
                    for (int i = 3; i < 4; i++)
                    {
                        decimal toplam = 0;
                        for (int j = 0; j < dataGridView2.Rows.Count + 1; ++j)
                        {
                            if (j < dataGridView2.Rows.Count && dataGridView2[2, j].Value.ToString() == CmbBilgi.SelectedItem.ToString() && ((DateTime)dataGridView2[4, j].Value).Year == Convert.ToInt32(CmbYil.SelectedItem))
                            {
                                toplam += Convert.ToDecimal(dataGridView2.Rows[j].Cells[i].Value);
                            }
                            else if (j == dataGridView2.Rows.Count)
                            {
                                gider = toplam;
                            }
                        }
                    }
                    if (gider != 0)
                    {
                        lblGider.Text = String.Format("{0:N}\n", gider);
                    }
                    else
                    {
                        lblGider.Text = "0,00";
                    }
                }
                else if (CmbGider.SelectedIndex == 1 || CmbGider.SelectedIndex == 2 || CmbGider.SelectedIndex == 3 || CmbGider.SelectedIndex == 4 || CmbGider.SelectedIndex == 5 || CmbGider.SelectedIndex == 6 || CmbGider.SelectedIndex == 7 || CmbGider.SelectedIndex == 8 || CmbGider.SelectedIndex == 9 || CmbGider.SelectedIndex == 10 || CmbGider.SelectedIndex == 11 || CmbGider.SelectedIndex == 12)//Aylar
                {
                    int Yil = Convert.ToInt32(CmbYil.SelectedItem);
                    var AyinIlkGunu = new DateTime(Yil, CmbGider.SelectedIndex, 1);
                    var AyinSonGunu = AyinIlkGunu.AddMonths(1).AddDays(-1);

                    //var gider = hesap.Where(x => x.cKisiID == Convert.ToInt32(lblID.Tag)).Where(x => x.DurumID == (CmbBilgi.SelectedIndex + 1) && (x.Tarih.Value >= AyinIlkGunu && x.Tarih.Value <= AyinSonGunu)).Sum(x => x.Tutar);
                    for (int i = 3; i < 4; i++)
                    {
                        decimal toplam = 0;
                        for (int j = 0; j < dataGridView2.Rows.Count + 1; ++j)
                        {
                            if (j < dataGridView2.Rows.Count && dataGridView2[2, j].Value.ToString() == CmbBilgi.SelectedItem.ToString() && (DateTime)dataGridView2[4, j].Value >= AyinIlkGunu && (DateTime)dataGridView2[4, j].Value <= AyinSonGunu)
                            {
                                toplam += Convert.ToDecimal(dataGridView2.Rows[j].Cells[i].Value);
                            }
                            else if (j == dataGridView2.Rows.Count)
                            {
                                gider = toplam;
                            }
                        }
                    }

                    if (gider != 0)
                    {
                        lblGider.Text = String.Format("{0:N}\n", gider);
                    }
                    else
                    {
                        lblGider.Text = "0,00";
                    }
                }
            }
        }

        private void CmbYil_SelectedIndexChanged(object sender, EventArgs e)
        {
            decimal gider = 0;
            /*
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
            */
            if (CmbYil.SelectedIndex == DateTime.Now.Year - 2017 + 1)//Tümü
            {
                CmbGider.SelectedIndex = 0;

                //var gider = hesap.Where(x => x.cKisiID == Convert.ToInt32(lblID.Tag)).Where(x => x.DurumID == (CmbBilgi.SelectedIndex + 1)).Sum(x => x.Tutar);

                for (int i = 3; i < 4; i++)
                {
                    decimal toplam = 0;
                    for (int j = 0; j < dataGridView2.Rows.Count + 1; ++j)
                    {
                        if (j < dataGridView2.Rows.Count && dataGridView2[2, j].Value.ToString() == CmbBilgi.SelectedItem.ToString())
                        {
                            toplam += Convert.ToDecimal(dataGridView2.Rows[j].Cells[i].Value);
                        }
                        else if (j == dataGridView2.Rows.Count)
                        {
                            gider = toplam;
                        }
                    }
                }

                if (gider != 0) //if (gider != null)
                {
                    lblGider.Text = String.Format("{0:N}\n", gider);
                }
                else
                {
                    lblGider.Text = "0,00";
                }
            }
            else //Yıllara Göre
            {
                if (CmbGider.SelectedIndex == 0)//Genel
                {
                    //var gider = hesap.Where(x => x.cKisiID == Convert.ToInt32(lblID.Tag)).Where(x => x.DurumID == (CmbBilgi.SelectedIndex + 1) && (x.Tarih.Value.Year == Convert.ToInt32(CmbYil.SelectedItem))).Sum(x => x.Tutar);
                    for (int i = 3; i < 4; i++)
                    {
                        decimal toplam = 0;
                        for (int j = 0; j < dataGridView2.Rows.Count + 1; ++j)
                        {
                            if (j < dataGridView2.Rows.Count && dataGridView2[2, j].Value.ToString() == CmbBilgi.SelectedItem.ToString() && ((DateTime)dataGridView2[4, j].Value).Year == Convert.ToInt32(CmbYil.SelectedItem))
                            {
                                toplam += Convert.ToDecimal(dataGridView2.Rows[j].Cells[i].Value);
                            }
                            else if (j == dataGridView2.Rows.Count)
                            {
                                gider = toplam;
                            }
                        }
                    }


                    if (gider != 0)
                    {
                        lblGider.Text = String.Format("{0:N}\n", gider);
                    }
                    else
                    {
                        lblGider.Text = "0,00";
                    }
                }
                else if (CmbGider.SelectedIndex == 1 || CmbGider.SelectedIndex == 2 || CmbGider.SelectedIndex == 3 || CmbGider.SelectedIndex == 4 || CmbGider.SelectedIndex == 5 || CmbGider.SelectedIndex == 6 || CmbGider.SelectedIndex == 7 || CmbGider.SelectedIndex == 8 || CmbGider.SelectedIndex == 9 || CmbGider.SelectedIndex == 10 || CmbGider.SelectedIndex == 11 || CmbGider.SelectedIndex == 12)//Aylar
                {
                    int Yil = Convert.ToInt32(CmbYil.SelectedItem);
                    var AyinIlkGunu = new DateTime(Yil, CmbGider.SelectedIndex, 1);
                    var AyinSonGunu = AyinIlkGunu.AddMonths(1).AddDays(-1);

                    //var gider = hesap.Where(x => x.cKisiID == Convert.ToInt32(lblID.Tag)).Where(x => x.DurumID == (CmbBilgi.SelectedIndex + 1) && (x.Tarih.Value >= AyinIlkGunu && x.Tarih.Value <= AyinSonGunu)).Sum(x => x.Tutar);
                    for (int i = 3; i < 4; i++)
                    {
                        decimal toplam = 0;
                        for (int j = 0; j < dataGridView2.Rows.Count + 1; ++j)
                        {
                            if (j < dataGridView2.Rows.Count && dataGridView2[2, j].Value.ToString() == CmbBilgi.SelectedItem.ToString() && (DateTime)dataGridView2[4, j].Value >= AyinIlkGunu && (DateTime)dataGridView2[4, j].Value <= AyinSonGunu)
                            {
                                toplam += Convert.ToDecimal(dataGridView2.Rows[j].Cells[i].Value);
                            }
                            else if (j == dataGridView2.Rows.Count)
                            {
                                gider = toplam;
                            }
                        }
                    }

                    if (gider != 0)
                    {
                        lblGider.Text = String.Format("{0:N}\n", gider);
                    }
                    else
                    {
                        lblGider.Text = "0,00";
                    }
                }
            }
        }

        private void dataGridView2_RowsRemoved(object sender, DataGridViewRowsRemovedEventArgs e)
        {
            decimal gider = 0;
                if (CmbYil.SelectedIndex == DateTime.Now.Year - 2017 + 1)//Tümü
                {
                    //var gider = hesap.Where(x => x.cKisiID == Convert.ToInt32(lblID.Tag)).Where(x => x.DurumID == (CmbBilgi.SelectedIndex + 1)).Sum(x => x.Tutar);
                    for (int i = 3; i < 4; i++)
                    {
                        decimal toplam = 0;
                        for (int j = 0; j < dataGridView2.Rows.Count + 1; ++j)
                        {
                            if (j < dataGridView2.Rows.Count && dataGridView2[2, j].Value.ToString() == CmbBilgi.SelectedItem.ToString())
                            {
                                toplam += Convert.ToDecimal(dataGridView2.Rows[j].Cells[i].Value);
                            }
                            else if (j == dataGridView2.Rows.Count)
                            {
                                gider = toplam;
                            }
                        }
                    }

                    if (gider != 0)
                    {
                        lblGider.Text = String.Format("{0:N}\n", gider);
                    }
                    else
                    {
                        lblGider.Text = "0,00";
                    }
                }
                else //Yıllara Göre
                {
                    if (CmbGider.SelectedIndex == 0)//Genel
                    {
                        //var gider = hesap.Where(x => x.cKisiID == Convert.ToInt32(lblID.Tag)).Where(x => x.DurumID == (CmbBilgi.SelectedIndex + 1) && (x.Tarih.Value.Year == Convert.ToInt32(CmbYil.SelectedItem))).Sum(x => x.Tutar);
                        for (int i = 3; i < 4; i++)
                        {
                            decimal toplam = 0;
                            for (int j = 0; j < dataGridView2.Rows.Count + 1; ++j)
                            {
                                if (j < dataGridView2.Rows.Count && dataGridView2[2, j].Value.ToString() == CmbBilgi.SelectedItem.ToString() && ((DateTime)dataGridView2[4, j].Value).Year == Convert.ToInt32(CmbYil.SelectedItem))
                                {
                                    toplam += Convert.ToDecimal(dataGridView2.Rows[j].Cells[i].Value);
                                }
                                else if (j == dataGridView2.Rows.Count)
                                {
                                    gider = toplam;
                                }
                            }
                        }
                        if (gider != 0)
                        {
                            lblGider.Text = String.Format("{0:N}\n", gider);
                        }
                        else
                        {
                            lblGider.Text = "0,00";
                        }
                    }
                    else if (CmbGider.SelectedIndex == 1 || CmbGider.SelectedIndex == 2 || CmbGider.SelectedIndex == 3 || CmbGider.SelectedIndex == 4 || CmbGider.SelectedIndex == 5 || CmbGider.SelectedIndex == 6 || CmbGider.SelectedIndex == 7 || CmbGider.SelectedIndex == 8 || CmbGider.SelectedIndex == 9 || CmbGider.SelectedIndex == 10 || CmbGider.SelectedIndex == 11 || CmbGider.SelectedIndex == 12)//Aylar
                    {
                        int Yil = Convert.ToInt32(CmbYil.SelectedItem);
                        var AyinIlkGunu = new DateTime(Yil, CmbGider.SelectedIndex, 1);
                        var AyinSonGunu = AyinIlkGunu.AddMonths(1).AddDays(-1);

                        //var gider = hesap.Where(x => x.cKisiID == Convert.ToInt32(lblID.Tag)).Where(x => x.DurumID == (CmbBilgi.SelectedIndex + 1) && (x.Tarih.Value >= AyinIlkGunu && x.Tarih.Value <= AyinSonGunu)).Sum(x => x.Tutar);
                        for (int i = 3; i < 4; i++)
                        {
                            decimal toplam = 0;
                            for (int j = 0; j < dataGridView2.Rows.Count + 1; ++j)
                            {
                                if (j < dataGridView2.Rows.Count && dataGridView2[2, j].Value.ToString() == CmbBilgi.SelectedItem.ToString() && (DateTime)dataGridView2[4, j].Value >= AyinIlkGunu && (DateTime)dataGridView2[4, j].Value <= AyinSonGunu)
                                {
                                    toplam += Convert.ToDecimal(dataGridView2.Rows[j].Cells[i].Value);
                                }
                                else if (j == dataGridView2.Rows.Count)
                                {
                                    gider = toplam;
                                }
                            }
                        }

                        if (gider != 0)
                        {
                            lblGider.Text = String.Format("{0:N}\n", gider);
                        }
                        else
                        {
                            lblGider.Text = "0,00";
                        }
                    }
                } 
        }

        private void dataGridView2_RowsAdded(object sender, DataGridViewRowsAddedEventArgs e)
        {
            decimal gider = 0;
            if (dataGridView2[2, 0].Value != null)
            {
                if (CmbYil.SelectedIndex == DateTime.Now.Year - 2017 + 1)//Tümü
                {
                    //var gider = hesap.Where(x => x.cKisiID == Convert.ToInt32(lblID.Tag)).Where(x => x.DurumID == (CmbBilgi.SelectedIndex + 1)).Sum(x => x.Tutar);

                    for (int i = 3; i < 4; i++)
                    {
                        decimal toplam = 0;
                        for (int j = 0; j < dataGridView2.Rows.Count + 1; ++j)
                        {
                            if (j < dataGridView2.Rows.Count && dataGridView2[2, j].Value.ToString() == CmbBilgi.SelectedItem.ToString())
                            {
                                toplam += Convert.ToDecimal(dataGridView2.Rows[j].Cells[i].Value);
                            }
                            else if (j == dataGridView2.Rows.Count)
                            {
                                gider = toplam;
                            }
                        }
                    }

                    if (gider != 0)
                    {
                        lblGider.Text = String.Format("{0:N}\n", gider);
                    }
                    else
                    {
                        lblGider.Text = "0,00";
                    }
                }
                else //Yıllara Göre
                {
                    if (CmbGider.SelectedIndex == 0)//Genel
                    {
                        //var gider = hesap.Where(x => x.cKisiID == Convert.ToInt32(lblID.Tag)).Where(x => x.DurumID == (CmbBilgi.SelectedIndex + 1) && (x.Tarih.Value.Year == Convert.ToInt32(CmbYil.SelectedItem))).Sum(x => x.Tutar);
                        for (int i = 3; i < 4; i++)
                        {
                            decimal toplam = 0;
                            for (int j = 0; j < dataGridView2.Rows.Count + 1; ++j)
                            {
                                if (j < dataGridView2.Rows.Count && dataGridView2[2, j].Value.ToString() == CmbBilgi.SelectedItem.ToString() && ((DateTime)dataGridView2[4, j].Value).Year == Convert.ToInt32(CmbYil.SelectedItem))
                                {
                                    toplam += Convert.ToDecimal(dataGridView2.Rows[j].Cells[i].Value);
                                }
                                else if (j == dataGridView2.Rows.Count)
                                {
                                    gider = toplam;
                                }
                            }
                        }
                        if (gider != 0)
                        {
                            lblGider.Text = String.Format("{0:N}\n", gider);
                        }
                        else
                        {
                            lblGider.Text = "0,00";
                        }
                    }
                    else if (CmbGider.SelectedIndex == 1 || CmbGider.SelectedIndex == 2 || CmbGider.SelectedIndex == 3 || CmbGider.SelectedIndex == 4 || CmbGider.SelectedIndex == 5 || CmbGider.SelectedIndex == 6 || CmbGider.SelectedIndex == 7 || CmbGider.SelectedIndex == 8 || CmbGider.SelectedIndex == 9 || CmbGider.SelectedIndex == 10 || CmbGider.SelectedIndex == 11 || CmbGider.SelectedIndex == 12)//Aylar
                    {
                        int Yil = Convert.ToInt32(CmbYil.SelectedItem);
                        var AyinIlkGunu = new DateTime(Yil, CmbGider.SelectedIndex, 1);
                        var AyinSonGunu = AyinIlkGunu.AddMonths(1).AddDays(-1);

                        //var gider = hesap.Where(x => x.cKisiID == Convert.ToInt32(lblID.Tag)).Where(x => x.DurumID == (CmbBilgi.SelectedIndex + 1) && (x.Tarih.Value >= AyinIlkGunu && x.Tarih.Value <= AyinSonGunu)).Sum(x => x.Tutar);
                        for (int i = 3; i < 4; i++)
                        {
                            decimal toplam = 0;
                            for (int j = 0; j < dataGridView2.Rows.Count + 1; ++j)
                            {
                                if (j < dataGridView2.Rows.Count && dataGridView2[2, j].Value.ToString() == CmbBilgi.SelectedItem.ToString() && (DateTime)dataGridView2[4, j].Value >= AyinIlkGunu && (DateTime)dataGridView2[4, j].Value <= AyinSonGunu)
                                {
                                    toplam += Convert.ToDecimal(dataGridView2.Rows[j].Cells[i].Value);
                                }
                                else if (j == dataGridView2.Rows.Count)
                                {
                                    gider = toplam;
                                }

                            }
                        }


                        if (gider != 0)
                        {
                            lblGider.Text = String.Format("{0:N}\n", gider);
                        }
                        else
                        {
                            lblGider.Text = "0,00";
                        }
                    }
                }
            }
            
        }

        private void LblDaAd_Click(object sender, EventArgs e)
        {
            if (LblDaAd.Text != "Ad Soyad")
            {
                txtAra.Text = LblDaAd.Text;
            }   
        }

        private void LblDaAd_MouseEnter(object sender, EventArgs e)
        {
            if (LblDaAd.Text != "Ad Soyad")
            {
                LblDaAd.ForeColor = Color.Blue;
            }   
        }

        private void LblDaAd_MouseLeave(object sender, EventArgs e)
        {
            LblDaAd.ForeColor = Color.Black;
        }

        private void DataGridViewDa_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            DetayForm df = new DetayForm();
            DialogResult sonuc = df.ShowDialog();
        }

        private void dataGridView4_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            DataGridViewRow row = dataGridView4.CurrentRow;
            KisiIdDetay = (int)row.Cells["KisiID"].Value;
        }

        private void dataGridView4_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            DataGridViewRow row = dataGridView4.CurrentRow;
            KisiIdDetay = (int)row.Cells["KisiID"].Value;
            KisiAdDetay = row.Cells["Ad"].Value.ToString();
            KisiFirmaDetay = row.Cells["Firma"].Value.ToString();
            KisiTel1Detay = row.Cells["Telefon"].Value.ToString();
            DetayForm df = new DetayForm();
            DialogResult sonuc = df.ShowDialog();
        }

        private void TxtAraDetay2_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                var sonuc = from kisiler in ctx.Kisis
                            join cariler in ctx.Caris on kisiler.KisiID equals cariler.KisiID
                            join durumlar in ctx.Durums on cariler.DurumID equals durumlar.DurumID
                            orderby cariler.Tarih descending
                            select new
                            {
                                kisiler.KisiID,
                                kisiler.Ad,
                                kisiler.Firma,
                                kisiler.Tel1,
                                kisiler.Tel2,
                                kisiler.Adres,
                                Durum = durumlar.Durumlar,
                                cariler.Tutar,
                                cariler.Tarih,
                                Açıklama = cariler.Aciklama,
                                durumlar.DurumID,
                                cariler.CariID
                            };

                DataGridViewDa.DataSource = sonuc.Where(x => x.Ad.Contains(TxtAraDetay2.Text) || x.Firma.Contains(TxtAraDetay2.Text) || x.Durum.Contains(TxtAraDetay2.Text) || x.Açıklama.Contains(TxtAraDetay2.Text) || Convert.ToString(x.Tutar).Contains(TxtAraDetay2.Text));

                DataGridViewDa.Columns["KisiID"].Visible = DataGridViewDa.Columns["Ad"].Visible = DataGridViewDa.Columns["Firma"].Visible = DataGridViewDa.Columns["Tel1"].Visible = DataGridViewDa.Columns["Tel2"].Visible = DataGridViewDa.Columns["Adres"].Visible = DataGridViewDa.Columns["DurumID"].Visible = DataGridViewDa.Columns["CariID"].Visible = false;

                DataGridViewDa.Columns[3].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                DataGridViewDa.Columns[3].HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleRight;
                DataGridViewDa.Columns[4].HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
                DataGridViewDa.Columns[4].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            }
        }

        private void txtEAd_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.Handled = Char.IsPunctuation(e.KeyChar) || Char.IsSymbol(e.KeyChar);
        }

        private void txtGAd_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.Handled = Char.IsPunctuation(e.KeyChar) || Char.IsSymbol(e.KeyChar);
        }

        private void dataGridView4_RowsAdded(object sender, DataGridViewRowsAddedEventArgs e)
        {
            //DataGridViewRow row4 = dataGridView4.CurrentRow;
            if (dataGridView4.Columns.Contains("SonTarih"))
            {
                foreach (DataGridViewRow dGVRow in this.dataGridView4.Rows)
                {
                    if (Convert.ToInt32(dGVRow.Cells["Karaliste"].Value) == 1)
                    {
                        dGVRow.DefaultCellStyle.BackColor = Color.Red;
                        dGVRow.DefaultCellStyle.ForeColor = Color.White;
                    }
                    else if (dGVRow.Cells["SonTarih"].Value == null || dGVRow.Cells["SonTarih"].Value == DBNull.Value || String.IsNullOrWhiteSpace(dGVRow.Cells["SonTarih"].Value.ToString()))
                    {
                        dGVRow.DefaultCellStyle.BackColor = Color.White;
                        dGVRow.DefaultCellStyle.ForeColor = Color.Black;
                    }
                    else if ((DateTime.Today - (DateTime)dGVRow.Cells["SonTarih"].Value).TotalDays > 31 && (DateTime.Today - (DateTime)dGVRow.Cells["SonTarih"].Value).TotalDays < 62)
                    {
                        dGVRow.DefaultCellStyle.BackColor = Color.FromArgb(255, 231, 231);
                        dGVRow.DefaultCellStyle.ForeColor = Color.Black;
                    }
                    else if ((DateTime.Today - (DateTime)dGVRow.Cells["SonTarih"].Value).TotalDays > 61 && (DateTime.Today - (DateTime)dGVRow.Cells["SonTarih"].Value).TotalDays < 93)
                    {
                        dGVRow.DefaultCellStyle.BackColor = Color.FromArgb(255, 210, 210);
                        dGVRow.DefaultCellStyle.ForeColor = Color.Black;
                    }
                    else if ((DateTime.Today - (DateTime)dGVRow.Cells["SonTarih"].Value).TotalDays > 92 && (DateTime.Today - (DateTime)dGVRow.Cells["SonTarih"].Value).TotalDays < 123)
                    {
                        dGVRow.DefaultCellStyle.BackColor = Color.FromArgb(255, 189, 189);
                        dGVRow.DefaultCellStyle.ForeColor = Color.Black;
                    }
                    else if ((DateTime.Today - (DateTime)dGVRow.Cells["SonTarih"].Value).TotalDays > 122 && (DateTime.Today - (DateTime)dGVRow.Cells["SonTarih"].Value).TotalDays < 154)
                    {
                        dGVRow.DefaultCellStyle.BackColor = Color.FromArgb(255, 168, 168);
                        dGVRow.DefaultCellStyle.ForeColor = Color.Black;
                    }
                    else if ((DateTime.Today - (DateTime)dGVRow.Cells["SonTarih"].Value).TotalDays > 153 && (DateTime.Today - (DateTime)dGVRow.Cells["SonTarih"].Value).TotalDays < 184)
                    {
                        dGVRow.DefaultCellStyle.BackColor = Color.FromArgb(255, 147, 147);
                        dGVRow.DefaultCellStyle.ForeColor = Color.Black;
                    }
                    else if ((DateTime.Today - (DateTime)dGVRow.Cells["SonTarih"].Value).TotalDays > 183 && (DateTime.Today - (DateTime)dGVRow.Cells["SonTarih"].Value).TotalDays < 215)
                    {
                        dGVRow.DefaultCellStyle.BackColor = Color.FromArgb(255, 126, 126);
                        dGVRow.DefaultCellStyle.ForeColor = Color.Black;
                    }
                    else if ((DateTime.Today - (DateTime)dGVRow.Cells["SonTarih"].Value).TotalDays > 214 && (DateTime.Today - (DateTime)dGVRow.Cells["SonTarih"].Value).TotalDays < 246)
                    {
                        dGVRow.DefaultCellStyle.BackColor = Color.FromArgb(255, 125, 125);
                        dGVRow.DefaultCellStyle.ForeColor = Color.Black;
                    }
                    else if ((DateTime.Today - (DateTime)dGVRow.Cells["SonTarih"].Value).TotalDays > 245 && (DateTime.Today - (DateTime)dGVRow.Cells["SonTarih"].Value).TotalDays < 276)
                    {
                        dGVRow.DefaultCellStyle.BackColor = Color.FromArgb(255, 104, 104);
                        dGVRow.DefaultCellStyle.ForeColor = Color.Black;
                    }
                    else if ((DateTime.Today - (DateTime)dGVRow.Cells["SonTarih"].Value).TotalDays > 275 && (DateTime.Today - (DateTime)dGVRow.Cells["SonTarih"].Value).TotalDays < 307)
                    {
                        dGVRow.DefaultCellStyle.BackColor = Color.FromArgb(255, 91, 83);
                        dGVRow.DefaultCellStyle.ForeColor = Color.Black;
                    }
                    else if ((DateTime.Today - (DateTime)dGVRow.Cells["SonTarih"].Value).TotalDays > 306 && (DateTime.Today - (DateTime)dGVRow.Cells["SonTarih"].Value).TotalDays < 337)
                    {
                        dGVRow.DefaultCellStyle.BackColor = Color.FromArgb(255, 81, 81);
                        dGVRow.DefaultCellStyle.ForeColor = Color.Black;
                    }
                    else if ((DateTime.Today - (DateTime)dGVRow.Cells["SonTarih"].Value).TotalDays > 336 && (DateTime.Today - (DateTime)dGVRow.Cells["SonTarih"].Value).TotalDays < 366)
                    {
                        dGVRow.DefaultCellStyle.BackColor = Color.FromArgb(255, 71, 71);
                        dGVRow.DefaultCellStyle.ForeColor = Color.Black;
                    }
                    else if ((DateTime.Today - (DateTime)dGVRow.Cells["SonTarih"].Value).TotalDays > 365)
                    {
                        dGVRow.DefaultCellStyle.BackColor = Color.FromArgb(255, 51, 51);
                        dGVRow.DefaultCellStyle.ForeColor = Color.Black;
                    }
                    else
                    {
                        dGVRow.DefaultCellStyle.BackColor = Color.White;
                        dGVRow.DefaultCellStyle.ForeColor = Color.Black;
                    }
                }
            }
        }

        private void dataGridView1_RowsAdded(object sender, DataGridViewRowsAddedEventArgs e)
        {
            foreach (DataGridViewRow dGVRow in this.dataGridView1.Rows)
            {
                if (Convert.ToInt32(dGVRow.Cells["Karaliste"].Value) == 1)
                {
                    dGVRow.DefaultCellStyle.BackColor = Color.Red;
                }
                else
                {
                    dGVRow.DefaultCellStyle.BackColor = Color.White;
                }
            }
        }

        private void BtnCizelge_Click(object sender, EventArgs e)
        {
            string cay = "";
            DateTime dateValue;
            //ÇizelgeFormu Aç
            CizelgeForm cf = new CizelgeForm();
            DialogResult sonucf = cf.ShowDialog();

            if (sonucf == DialogResult.OK)
            {
                if (CTarihAy == 1)
                {
                    cay = "OCAK";
                }
                else if (CTarihAy == 2)
                {
                    cay = "ŞUBAT";
                }
                else if (CTarihAy == 3)
                {
                    cay = "MART";
                }
                else if (CTarihAy == 4)
                {
                    cay = "NİSAN";
                }
                else if (CTarihAy == 5)
                {
                    cay = "MAYIS";
                }
                else if (CTarihAy == 6)
                {
                    cay = "HAZİRAN";
                }
                else if (CTarihAy == 7)
                {
                    cay = "TEMMUZ";
                }
                else if (CTarihAy == 8)
                {
                    cay = "AĞUSTOS";
                }
                else if (CTarihAy == 9)
                {
                    cay = "EYLÜL";
                }
                else if (CTarihAy == 10)
                {
                    cay = "EKİM";
                }
                else if (CTarihAy == 11)
                {
                    cay = "KASIM";
                }
                else if (CTarihAy == 12)
                {
                    cay = "ARALIK";
                }

                //Dosya ismi - Kaydet
                string fileName;

                SaveFileDialog saveFileDialog1 = new SaveFileDialog
                {
                    Filter = "xls files (*.xlsx)|*.xlsx|All files (*.*)|*.*",
                    Title = "To Excel",
                };

                saveFileDialog1.FileName = "ÇİZELGE_" + cay + "-" + CTarihYil;

                if (saveFileDialog1.ShowDialog() == DialogResult.OK)
                {
                    fileName = saveFileDialog1.FileName;
                    var workbook = new XLWorkbook();

                    //DataTable Cizelge Oluştur
                    DataTable CizelgeDt = new DataTable();

                    //No Sütunu Ekle
                    DataColumn columno = new DataColumn
                    {
                        DataType = System.Type.GetType("System.Int32"),
                    };

                    CizelgeDt.Columns.Add(columno);
                    CizelgeDt.Columns["Column1"].ColumnName = "NO";

                    CizelgeDt.Columns.Add("AY");

                    DataColumn columyil = new DataColumn
                    {
                        DataType = System.Type.GetType("System.Int32"),
                    };

                    CizelgeDt.Columns.Add(columyil);
                    CizelgeDt.Columns["Column1"].ColumnName = "YIL";
                    //CizelgeDt.Columns.Add("YIL");

                    CizelgeDt.Columns.Add("GÜN");
                    CizelgeDt.Columns.Add("AÇIKLAMA");
                    CizelgeDt.Columns.Add("NOT/TUTAR");

                    var CTarih1 = new DateTime(CTarihYil, CTarihAy, 1);
                    byte HaftaGunBas = Convert.ToByte(CTarih1.DayOfWeek); // 0 = Sunday....
                    byte GunSayisiAy = Convert.ToByte(DateTime.DaysInMonth(CTarihYil, CTarihAy));
                    //string dateString = HaftaGunBas + "/" + CTarihAy + "/" + CTarihYil;
                    string dateString = CTarih1.ToString();
                    dateValue = DateTime.Parse(dateString, CultureInfo.CreateSpecificCulture("tr-TR"));

                    CizelgeDt.Rows.Add();

                    //Satırları Ekle
                    for (int i = 0; i < GunSayisiAy; i++)
                    {
                        CizelgeDt.Rows.Add();
                        CizelgeDt.Rows[i + 1][0] = i + 1;
                        CizelgeDt.Rows[i + 1][1] = cay;
                        CizelgeDt.Rows[i + 1][2] = CTarihYil;
                        CizelgeDt.Rows[i + 1][3] = dateValue.ToString("dddd", new CultureInfo("tr-TR"));
                        CTarih1 = CTarih1.AddDays(1);
                        dateString = CTarih1.ToString();
                        dateValue = DateTime.Parse(dateString, CultureInfo.CreateSpecificCulture("tr-TR"));
                    }


                    //Excel Sayfasına 'Çizelge' i ekle.
                    using (workbook)
                    {
                        workbook.Worksheets.Add(CizelgeDt, "Çizelge");
                    }

                   
					//Düzenleme Hizalama, Renk vs.
                    workbook.Worksheet("Çizelge").Row(1).Style.Fill.BackgroundColor = XLColor.White;
                    workbook.Worksheet("Çizelge").Row(1).Style.Font.SetBold();
                    workbook.Worksheet("Çizelge").Row(1).Style.Font.SetFontColor(XLColor.FromArgb(0, 0, 1));
                    workbook.Worksheet("Çizelge").PageSetup.Header.Left.AddText(cay+"/"+CTarihYil).SetBold();
                    workbook.Worksheet("Çizelge").PageSetup.Footer.Center.AddText("kalemobilyadekorasyon.com").SetBold();
                    workbook.Worksheet("Çizelge").Style.Font.FontSize = 15;
                    workbook.Worksheet("Çizelge").Column(2).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                    workbook.Worksheet("Çizelge").Column(3).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;

                    workbook.Worksheet("Çizelge").Columns().AdjustToContents();
                    workbook.Worksheet("Çizelge").Column(1).Width = 3.3;
                    workbook.Worksheet("Çizelge").Column(3).Width = 6;
                    workbook.Worksheet("Çizelge").Column(5).Width = 37;
                    workbook.Worksheet("Çizelge").Column(6).Width = 17;

                    workbook.Worksheet("Çizelge").RowHeight = 23;
                    workbook.Worksheet("Çizelge").PageSetup.AdjustTo(95);
                    //workbook.Worksheet("Çizelge").PageSetup.FitToPages(1, 2);

                    //Kaydet
                    do
                    {
                        try
                        {
                            workbook.SaveAs(fileName);
                            MessageBox.Show("Excel dosyası kaydedildi.", "Bilgi", MessageBoxButtons.OK, MessageBoxIcon.Information);
                            
                            break;
                        }
                        catch (System.IO.IOException)
                        {
                            MessageBox.Show("Kayıt yapılamadı! Kaydetmeye çalıştığınız dosya açık olabilir.", "Uyarı", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                            
                            break;
                        }
                    } while (true);
                }
            }
        }

        private void BtnAraDetay_Click(object sender, EventArgs e)
        {
            if (DataGridViewDa.Rows.Count > 0)
            {
                CultureInfo customCulture = (CultureInfo)System.Threading.Thread.CurrentThread.CurrentCulture.Clone();
                customCulture.NumberFormat.NumberDecimalSeparator = ".";
                System.Threading.Thread.CurrentThread.CurrentCulture = customCulture;

                //Dosya ismi - Kaydet
                string fileName;

                SaveFileDialog saveFileDialog1 = new SaveFileDialog
                {
                    Filter = "xls files (*.xlsx)|*.xlsx|All files (*.*)|*.*",
                    Title = "To Excel",
                };

                saveFileDialog1.FileName = "(" + DateTime.Now.ToString("dd-MM-yyyy") + ")" + " Arama";

                if (saveFileDialog1.ShowDialog() == DialogResult.OK)
                {
                    fileName = saveFileDialog1.FileName;
                    var workbook = new XLWorkbook();

                    //DataTable Oluştur
                    DataTable AraDt = new DataTable();

                    //No Sütunu Ekle
                    DataColumn columno = new DataColumn
                    {
                        DataType = System.Type.GetType("System.Int32"),
                    };
                    AraDt.Columns.Add(columno);
                    AraDt.Columns["Column1"].ColumnName = "NO";
                   
                    //Sütunları Ekle
                    foreach (DataGridViewColumn column in DataGridViewDa.Columns)
                    {
                        AraDt.Columns.Add(column.HeaderText);
                    }

                    //Satırları Ekle
                    foreach (DataGridViewRow row in DataGridViewDa.Rows)
                    {
                        AraDt.Rows.Add();
                        //Hücreleri Ekle
                        foreach (DataGridViewCell cell in row.Cells)
                        {
                            if (cell.Value != null)
                            {
                                AraDt.Rows[AraDt.Rows.Count - 1][cell.ColumnIndex + 1] = cell.Value;
                            }
                        }
                    }

                    //No ları Ekle
                    for (int i = 0; i < DataGridViewDa.RowCount; i++)
                    {
                        AraDt.Rows[i][0] = i + 1;
                    }

                    //Gereksiz Sütunları Kaldır
                    AraDt.Columns.RemoveAt(1);
                    AraDt.Columns.RemoveAt(1);
                    AraDt.Columns.RemoveAt(1);
                    AraDt.Columns.RemoveAt(1);
                    AraDt.Columns.RemoveAt(1);
                    AraDt.Columns.RemoveAt(1);
                    AraDt.Columns.RemoveAt(5);
                    AraDt.Columns.RemoveAt(5);

                    //Tarih Sütunu Zamanı Kaldır
                    for (int i = 0; i < DataGridViewDa.RowCount; i++)
                    {
                        AraDt.Rows[i][3] = string.Format("{0:dd/MM/yyyy}", DataGridViewDa.Rows[i].Cells[8].Value);
                    }


                    //Excel Sayfasına Datatable ekle.
                    using (workbook)
                    {
                        workbook.Worksheets.Add(AraDt, "Arama");
                    }

                    //Düzenleme Hizalama, Renk vs.
                    workbook.Worksheet("Arama").Row(1).Style.Fill.BackgroundColor = XLColor.White;
                    workbook.Worksheet("Arama").Row(1).Style.Font.SetBold();
                    workbook.Worksheet("Arama").Row(1).Style.Font.SetFontColor(XLColor.FromArgb(0, 0, 1));
                    workbook.Worksheet("Arama").PageSetup.Header.Left.AddText("Arama").SetBold();
                    workbook.Worksheet("Arama").PageSetup.Header.Right.AddText(string.Format("{0:dd/MM/yyyy}", DateTime.Now));
                    workbook.Worksheet("Arama").PageSetup.Footer.Center.AddText("kalemobilyadekorasyon.com");
                    workbook.Worksheet("Arama").Column(1).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                    workbook.Worksheet("Arama").Column(4).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                    workbook.Worksheet("Arama").Column(3).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
                    workbook.Worksheet("Arama").Columns().AdjustToContents();

                    if (workbook.Worksheet("Arama").Column(5).Width < 49)
                    {
                        workbook.Worksheet("Arama").PageSetup.AdjustTo(100);
                    }
                    else if (workbook.Worksheet("Arama").Column(5).Width < 59 && workbook.Worksheet("Arama").Column(5).Width >= 49)
                    {
                        workbook.Worksheet("Arama").PageSetup.AdjustTo(90);
                    }
                    else if (workbook.Worksheet("Arama").Column(5).Width >= 59 && workbook.Worksheet("Arama").Column(5).Width < 70)
                    {
                        workbook.Worksheet("Arama").PageSetup.AdjustTo(80);
                    }
                    else if (workbook.Worksheet("Arama").Column(5).Width >= 70 && workbook.Worksheet("Arama").Column(5).Width < 98)
                    {
                        workbook.Worksheet("Arama").PageSetup.FitToPages(1, 2);
                        workbook.Worksheet("Arama").PageSetup.PageOrientation = XLPageOrientation.Landscape;
                    }
                    else if (workbook.Worksheet("Arama").Column(5).Width >= 98)
                    {
                        workbook.Worksheet("Arama").PageSetup.PageOrientation = XLPageOrientation.Landscape;
                        workbook.Worksheet("Arama").PageSetup.FitToPages(1, 3);
                    }

                    //Kaydet
                    do
                    {
                        try
                        {
                            workbook.SaveAs(fileName);
                            MessageBox.Show("Excel dosyası kaydedildi.", "Bilgi", MessageBoxButtons.OK, MessageBoxIcon.Information);

                            break;
                        }
                        catch (System.IO.IOException)
                        {
                            MessageBox.Show("Kayıt yapılamadı! Kaydetmeye çalıştığınız dosya açık olabilir.", "Uyarı", MessageBoxButtons.OK, MessageBoxIcon.Warning);

                            break;
                        }
                    } while (true);
                }
            }
        }
    }
}

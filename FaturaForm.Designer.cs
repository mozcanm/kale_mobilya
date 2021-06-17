namespace Kale_Mobilya
{
    partial class FaturaForm
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(FaturaForm));
            this.label1 = new System.Windows.Forms.Label();
            this.CmbAy1 = new System.Windows.Forms.ComboBox();
            this.CmbYil1 = new System.Windows.Forms.ComboBox();
            this.BtnFatura1 = new System.Windows.Forms.Button();
            this.BtnOk1 = new System.Windows.Forms.Button();
            this.label2 = new System.Windows.Forms.Label();
            this.button1 = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 11.25F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Underline))), System.Drawing.GraphicsUnit.Point, ((byte)(162)));
            this.label1.Location = new System.Drawing.Point(88, 9);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(143, 18);
            this.label1.TabIndex = 0;
            this.label1.Text = "FATURA DÖNEMİ";
            // 
            // CmbAy1
            // 
            this.CmbAy1.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.CmbAy1.Font = new System.Drawing.Font("Microsoft Sans Serif", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(162)));
            this.CmbAy1.FormattingEnabled = true;
            this.CmbAy1.Items.AddRange(new object[] {
            "Ocak",
            "Şubat",
            "Mart",
            "Nisan",
            "Mayıs",
            "Haziran",
            "Temmuz",
            "Ağustos",
            "Eylül",
            "Ekim",
            "Kasım",
            "Aralık"});
            this.CmbAy1.Location = new System.Drawing.Point(12, 51);
            this.CmbAy1.Name = "CmbAy1";
            this.CmbAy1.Size = new System.Drawing.Size(121, 26);
            this.CmbAy1.TabIndex = 1;
            // 
            // CmbYil1
            // 
            this.CmbYil1.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.CmbYil1.Font = new System.Drawing.Font("Microsoft Sans Serif", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(162)));
            this.CmbYil1.FormattingEnabled = true;
            this.CmbYil1.Items.AddRange(new object[] {
            "2019"});
            this.CmbYil1.Location = new System.Drawing.Point(185, 51);
            this.CmbYil1.Name = "CmbYil1";
            this.CmbYil1.Size = new System.Drawing.Size(121, 26);
            this.CmbYil1.TabIndex = 1;
            // 
            // BtnFatura1
            // 
            this.BtnFatura1.Font = new System.Drawing.Font("Microsoft Sans Serif", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(162)));
            this.BtnFatura1.Location = new System.Drawing.Point(174, 161);
            this.BtnFatura1.Name = "BtnFatura1";
            this.BtnFatura1.Size = new System.Drawing.Size(149, 31);
            this.BtnFatura1.TabIndex = 3;
            this.BtnFatura1.Text = "Tüm Faturaları Getir";
            this.BtnFatura1.UseVisualStyleBackColor = true;
            this.BtnFatura1.Click += new System.EventHandler(this.BtnFatura1_Click);
            // 
            // BtnOk1
            // 
            this.BtnOk1.Image = global::Kale_Mobilya.Properties.Resources.Accept_icon;
            this.BtnOk1.Location = new System.Drawing.Point(133, 92);
            this.BtnOk1.Name = "BtnOk1";
            this.BtnOk1.Size = new System.Drawing.Size(51, 46);
            this.BtnOk1.TabIndex = 4;
            this.BtnOk1.UseVisualStyleBackColor = true;
            this.BtnOk1.Click += new System.EventHandler(this.BtnOk1_Click);
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("Microsoft Sans Serif", 11.25F, System.Drawing.FontStyle.Italic, System.Drawing.GraphicsUnit.Point, ((byte)(162)));
            this.label2.Location = new System.Drawing.Point(135, 167);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(38, 18);
            this.label2.TabIndex = 5;
            this.label2.Text = "veya";
            // 
            // button1
            // 
            this.button1.Font = new System.Drawing.Font("Microsoft Sans Serif", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(162)));
            this.button1.Location = new System.Drawing.Point(8, 161);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(125, 31);
            this.button1.TabIndex = 3;
            this.button1.Text = "Görüneni Getir";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // FaturaForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.SystemColors.HighlightText;
            this.ClientSize = new System.Drawing.Size(326, 198);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.BtnFatura1);
            this.Controls.Add(this.BtnOk1);
            this.Controls.Add(this.CmbAy1);
            this.Controls.Add(this.CmbYil1);
            this.Controls.Add(this.label1);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MaximizeBox = false;
            this.MaximumSize = new System.Drawing.Size(342, 236);
            this.MinimizeBox = false;
            this.MinimumSize = new System.Drawing.Size(342, 236);
            this.Name = "FaturaForm";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Faturalar";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.ComboBox CmbAy1;
        private System.Windows.Forms.ComboBox CmbYil1;
        private System.Windows.Forms.Button BtnFatura1;
        private System.Windows.Forms.Button BtnOk1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Button button1;
    }
}
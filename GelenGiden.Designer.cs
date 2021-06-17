namespace Kale_Mobilya
{
    partial class GelenGiden
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(GelenGiden));
            this.DtpTarih1 = new System.Windows.Forms.DateTimePicker();
            this.DtpTarih2 = new System.Windows.Forms.DateTimePicker();
            this.BtnTarih2 = new System.Windows.Forms.Button();
            this.BtnTarih1 = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // DtpTarih1
            // 
            this.DtpTarih1.Location = new System.Drawing.Point(12, 32);
            this.DtpTarih1.Name = "DtpTarih1";
            this.DtpTarih1.Size = new System.Drawing.Size(210, 24);
            this.DtpTarih1.TabIndex = 0;
            // 
            // DtpTarih2
            // 
            this.DtpTarih2.Location = new System.Drawing.Point(12, 62);
            this.DtpTarih2.Name = "DtpTarih2";
            this.DtpTarih2.Size = new System.Drawing.Size(210, 24);
            this.DtpTarih2.TabIndex = 0;
            // 
            // BtnTarih2
            // 
            this.BtnTarih2.BackColor = System.Drawing.SystemColors.Info;
            this.BtnTarih2.Image = global::Kale_Mobilya.Properties.Resources.cancel_icon;
            this.BtnTarih2.Location = new System.Drawing.Point(57, 92);
            this.BtnTarih2.Name = "BtnTarih2";
            this.BtnTarih2.Size = new System.Drawing.Size(54, 48);
            this.BtnTarih2.TabIndex = 1;
            this.BtnTarih2.UseVisualStyleBackColor = false;
            this.BtnTarih2.Click += new System.EventHandler(this.BtnTarih2_Click);
            // 
            // BtnTarih1
            // 
            this.BtnTarih1.BackColor = System.Drawing.SystemColors.Info;
            this.BtnTarih1.Image = global::Kale_Mobilya.Properties.Resources.Accept_icon;
            this.BtnTarih1.Location = new System.Drawing.Point(117, 92);
            this.BtnTarih1.Name = "BtnTarih1";
            this.BtnTarih1.Size = new System.Drawing.Size(54, 48);
            this.BtnTarih1.TabIndex = 1;
            this.BtnTarih1.UseVisualStyleBackColor = false;
            this.BtnTarih1.Click += new System.EventHandler(this.BtnTarih1_Click);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 11.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(162)));
            this.label1.Location = new System.Drawing.Point(5, 7);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(240, 18);
            this.label1.TabIndex = 2;
            this.label1.Text = "Gelir/Gider Carisi Eklensin mi?";
            // 
            // GelenGiden
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(9F, 18F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(254, 150);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.BtnTarih2);
            this.Controls.Add(this.BtnTarih1);
            this.Controls.Add(this.DtpTarih2);
            this.Controls.Add(this.DtpTarih1);
            this.Font = new System.Drawing.Font("Microsoft Sans Serif", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(162)));
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Margin = new System.Windows.Forms.Padding(4);
            this.MaximizeBox = false;
            this.Name = "GelenGiden";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Tarih Seçimi";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion
        private System.Windows.Forms.Button BtnTarih1;
        private System.Windows.Forms.Button BtnTarih2;
        public System.Windows.Forms.DateTimePicker DtpTarih1;
        public System.Windows.Forms.DateTimePicker DtpTarih2;
        private System.Windows.Forms.Label label1;
    }
}
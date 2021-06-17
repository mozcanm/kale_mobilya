namespace Kale_Mobilya
{
    partial class CizelgeForm
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(CizelgeForm));
            this.BtnCizOk1 = new System.Windows.Forms.Button();
            this.CmbCizYil1 = new System.Windows.Forms.ComboBox();
            this.CmbCizAy1 = new System.Windows.Forms.ComboBox();
            this.label1 = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // BtnCizOk1
            // 
            this.BtnCizOk1.Image = global::Kale_Mobilya.Properties.Resources.Accept_icon;
            this.BtnCizOk1.Location = new System.Drawing.Point(116, 97);
            this.BtnCizOk1.Name = "BtnCizOk1";
            this.BtnCizOk1.Size = new System.Drawing.Size(51, 46);
            this.BtnCizOk1.TabIndex = 8;
            this.BtnCizOk1.UseVisualStyleBackColor = true;
            this.BtnCizOk1.Click += new System.EventHandler(this.BtnCizOk1_Click);
            // 
            // CmbCizYil1
            // 
            this.CmbCizYil1.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.CmbCizYil1.Font = new System.Drawing.Font("Microsoft Sans Serif", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(162)));
            this.CmbCizYil1.FormattingEnabled = true;
            this.CmbCizYil1.Items.AddRange(new object[] {
            "2019"});
            this.CmbCizYil1.Location = new System.Drawing.Point(152, 56);
            this.CmbCizYil1.Name = "CmbCizYil1";
            this.CmbCizYil1.Size = new System.Drawing.Size(121, 26);
            this.CmbCizYil1.TabIndex = 6;
            // 
            // CmbCizAy1
            // 
            this.CmbCizAy1.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.CmbCizAy1.Font = new System.Drawing.Font("Microsoft Sans Serif", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(162)));
            this.CmbCizAy1.FormattingEnabled = true;
            this.CmbCizAy1.Items.AddRange(new object[] {
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
            this.CmbCizAy1.Location = new System.Drawing.Point(12, 56);
            this.CmbCizAy1.Name = "CmbCizAy1";
            this.CmbCizAy1.Size = new System.Drawing.Size(121, 26);
            this.CmbCizAy1.TabIndex = 7;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 11.25F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Underline))), System.Drawing.GraphicsUnit.Point, ((byte)(162)));
            this.label1.Location = new System.Drawing.Point(73, 18);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(149, 18);
            this.label1.TabIndex = 5;
            this.label1.Text = "ÇİZELGE DÖNEMİ";
            // 
            // CizelgeForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(285, 154);
            this.Controls.Add(this.BtnCizOk1);
            this.Controls.Add(this.CmbCizYil1);
            this.Controls.Add(this.CmbCizAy1);
            this.Controls.Add(this.label1);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "CizelgeForm";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Çizelge";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button BtnCizOk1;
        private System.Windows.Forms.ComboBox CmbCizYil1;
        private System.Windows.Forms.ComboBox CmbCizAy1;
        private System.Windows.Forms.Label label1;
    }
}
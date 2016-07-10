namespace SunnyExcel
{
    partial class AboutDialog
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(AboutDialog));
            this.btOK = new System.Windows.Forms.Button();
            this.pictureBox1 = new System.Windows.Forms.PictureBox();
            this.label3 = new System.Windows.Forms.Label();
            this.lnOpenXmlSdk = new System.Windows.Forms.LinkLabel();
            this.label2 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.lnOdinSoft = new System.Windows.Forms.LinkLabel();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).BeginInit();
            this.SuspendLayout();
            // 
            // btOK
            // 
            this.btOK.Location = new System.Drawing.Point(327, 227);
            this.btOK.Name = "btOK";
            this.btOK.Size = new System.Drawing.Size(75, 23);
            this.btOK.TabIndex = 0;
            this.btOK.Text = "OK";
            this.btOK.UseVisualStyleBackColor = true;
            this.btOK.Click += new System.EventHandler(this.btOK_Click);
            // 
            // pictureBox1
            // 
            this.pictureBox1.Image = ((System.Drawing.Image)(resources.GetObject("pictureBox1.Image")));
            this.pictureBox1.Location = new System.Drawing.Point(13, 13);
            this.pictureBox1.Name = "pictureBox1";
            this.pictureBox1.Size = new System.Drawing.Size(135, 145);
            this.pictureBox1.TabIndex = 1;
            this.pictureBox1.TabStop = false;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(165, 39);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(149, 12);
            this.label3.TabIndex = 4;
            this.label3.Text = "엑셀 변환 작업용 프로그램";
            // 
            // lnOpenXmlSdk
            // 
            this.lnOpenXmlSdk.AutoSize = true;
            this.lnOpenXmlSdk.Location = new System.Drawing.Point(11, 200);
            this.lnOpenXmlSdk.Name = "lnOpenXmlSdk";
            this.lnOpenXmlSdk.Size = new System.Drawing.Size(337, 12);
            this.lnOpenXmlSdk.TabIndex = 5;
            this.lnOpenXmlSdk.TabStop = true;
            this.lnOpenXmlSdk.Text = "본 프로그램은 Open-XML-SDK를 사용하여 구현 되었습니다.";
            this.lnOpenXmlSdk.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.WebLinkClicked);
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(167, 65);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(78, 12);
            this.label2.TabIndex = 6;
            this.label2.Text = "Version 1.0.0";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(167, 93);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(145, 12);
            this.label4.TabIndex = 7;
            this.label4.Text = "@2016 OdinSoft Co., Ltd.";
            // 
            // lnOdinSoft
            // 
            this.lnOdinSoft.AutoSize = true;
            this.lnOdinSoft.Location = new System.Drawing.Point(167, 117);
            this.lnOdinSoft.Name = "lnOdinSoft";
            this.lnOdinSoft.Size = new System.Drawing.Size(87, 12);
            this.lnOdinSoft.TabIndex = 9;
            this.lnOdinSoft.TabStop = true;
            this.lnOdinSoft.Text = "(주)오딘소프트";
            this.lnOdinSoft.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.WebLinkClicked);
            // 
            // AboutDialog
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(414, 271);
            this.Controls.Add(this.lnOdinSoft);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.lnOpenXmlSdk);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.pictureBox1);
            this.Controls.Add(this.btOK);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.KeyPreview = true;
            this.MinimumSize = new System.Drawing.Size(430, 310);
            this.Name = "AboutDialog";
            this.Text = "AboutDialog";
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.AboutDialog_FormClosing);
            this.Load += new System.EventHandler(this.AboutDialog_Load);
            this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.AboutDialog_KeyDown);
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button btOK;
        private System.Windows.Forms.PictureBox pictureBox1;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.LinkLabel lnOpenXmlSdk;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.LinkLabel lnOdinSoft;
    }
}
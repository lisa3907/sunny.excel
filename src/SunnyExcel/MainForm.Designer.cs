namespace SunnyExcel
{
    partial class MainForm
    {
        /// <summary>
        /// 필수 디자이너 변수입니다.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// 사용 중인 모든 리소스를 정리합니다.
        /// </summary>
        /// <param name="disposing">관리되는 리소스를 삭제해야 하면 true이고, 그렇지 않으면 false입니다.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form 디자이너에서 생성한 코드

        /// <summary>
        /// 디자이너 지원에 필요한 메서드입니다. 
        /// 이 메서드의 내용을 코드 편집기로 수정하지 마세요.
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(MainForm));
            this.ssStatusBar = new System.Windows.Forms.StatusStrip();
            this.status_message = new System.Windows.Forms.ToolStripStatusLabel();
            this.tsProgressBar = new System.Windows.Forms.ToolStripProgressBar();
            this.msMainMenu = new System.Windows.Forms.MenuStrip();
            this.파일FToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.miExit = new System.Windows.Forms.ToolStripMenuItem();
            this.tsHelp = new System.Windows.Forms.ToolStripMenuItem();
            this.miAbout = new System.Windows.Forms.ToolStripMenuItem();
            this.ofBeforeExcel = new System.Windows.Forms.OpenFileDialog();
            this.panel1 = new System.Windows.Forms.Panel();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.tbCreated = new System.Windows.Forms.TextBox();
            this.label5 = new System.Windows.Forms.Label();
            this.tbModifier = new System.Windows.Forms.TextBox();
            this.label4 = new System.Windows.Forms.Label();
            this.tbAuthor = new System.Windows.Forms.TextBox();
            this.label3 = new System.Windows.Forms.Label();
            this.tbBefore = new System.Windows.Forms.TextBox();
            this.btBefore = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.tbNumberOfRowPerPage = new System.Windows.Forms.TextBox();
            this.tbStartRowNumber = new System.Windows.Forms.TextBox();
            this.tbHeightOfRow = new System.Windows.Forms.TextBox();
            this.label8 = new System.Windows.Forms.Label();
            this.label7 = new System.Windows.Forms.Label();
            this.label6 = new System.Windows.Forms.Label();
            this.cbAfterOpen = new System.Windows.Forms.CheckBox();
            this.btAfter = new System.Windows.Forms.Button();
            this.tbAfter = new System.Windows.Forms.TextBox();
            this.label2 = new System.Windows.Forms.Label();
            this.ssStatusBar.SuspendLayout();
            this.msMainMenu.SuspendLayout();
            this.groupBox1.SuspendLayout();
            this.groupBox2.SuspendLayout();
            this.SuspendLayout();
            // 
            // ssStatusBar
            // 
            this.ssStatusBar.ImageScalingSize = new System.Drawing.Size(20, 20);
            this.ssStatusBar.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.status_message,
            this.tsProgressBar});
            this.ssStatusBar.Location = new System.Drawing.Point(0, 675);
            this.ssStatusBar.Name = "ssStatusBar";
            this.ssStatusBar.Padding = new System.Windows.Forms.Padding(1, 0, 16, 0);
            this.ssStatusBar.Size = new System.Drawing.Size(896, 26);
            this.ssStatusBar.TabIndex = 2;
            this.ssStatusBar.Text = "statusStrip1";
            // 
            // status_message
            // 
            this.status_message.Name = "status_message";
            this.status_message.Size = new System.Drawing.Size(609, 21);
            this.status_message.Spring = true;
            this.status_message.Text = "Ready";
            this.status_message.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // tsProgressBar
            // 
            this.tsProgressBar.Name = "tsProgressBar";
            this.tsProgressBar.Size = new System.Drawing.Size(229, 20);
            // 
            // msMainMenu
            // 
            this.msMainMenu.ImageScalingSize = new System.Drawing.Size(20, 20);
            this.msMainMenu.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.파일FToolStripMenuItem,
            this.tsHelp});
            this.msMainMenu.Location = new System.Drawing.Point(0, 0);
            this.msMainMenu.Name = "msMainMenu";
            this.msMainMenu.Padding = new System.Windows.Forms.Padding(7, 2, 0, 2);
            this.msMainMenu.Size = new System.Drawing.Size(896, 28);
            this.msMainMenu.TabIndex = 3;
            this.msMainMenu.Text = "menuStrip1";
            // 
            // 파일FToolStripMenuItem
            // 
            this.파일FToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.miExit});
            this.파일FToolStripMenuItem.Name = "파일FToolStripMenuItem";
            this.파일FToolStripMenuItem.Size = new System.Drawing.Size(68, 24);
            this.파일FToolStripMenuItem.Text = "파일(&F)";
            // 
            // miExit
            // 
            this.miExit.Name = "miExit";
            this.miExit.ShortcutKeys = ((System.Windows.Forms.Keys)((System.Windows.Forms.Keys.Alt | System.Windows.Forms.Keys.F4)));
            this.miExit.Size = new System.Drawing.Size(202, 26);
            this.miExit.Text = "끝내기(&X)";
            this.miExit.Click += new System.EventHandler(this.miExit_Click);
            // 
            // tsHelp
            // 
            this.tsHelp.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.miAbout});
            this.tsHelp.Name = "tsHelp";
            this.tsHelp.Size = new System.Drawing.Size(87, 24);
            this.tsHelp.Text = "도움말(&H)";
            // 
            // miAbout
            // 
            this.miAbout.Name = "miAbout";
            this.miAbout.ShortcutKeys = System.Windows.Forms.Keys.F1;
            this.miAbout.Size = new System.Drawing.Size(223, 26);
            this.miAbout.Text = "프로그램 정보(&A)";
            this.miAbout.Click += new System.EventHandler(this.miAbout_Click);
            // 
            // ofBeforeExcel
            // 
            this.ofBeforeExcel.Filter = "AllFiles(*.*)|*.*";
            this.ofBeforeExcel.Title = "변환 할 엑셀 파일을 선택 하세요.";
            // 
            // panel1
            // 
            this.panel1.Dock = System.Windows.Forms.DockStyle.Top;
            this.panel1.Location = new System.Drawing.Point(0, 28);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(896, 39);
            this.panel1.TabIndex = 17;
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.tbCreated);
            this.groupBox1.Controls.Add(this.label5);
            this.groupBox1.Controls.Add(this.tbModifier);
            this.groupBox1.Controls.Add(this.label4);
            this.groupBox1.Controls.Add(this.tbAuthor);
            this.groupBox1.Controls.Add(this.label3);
            this.groupBox1.Controls.Add(this.tbBefore);
            this.groupBox1.Controls.Add(this.btBefore);
            this.groupBox1.Controls.Add(this.label1);
            this.groupBox1.Dock = System.Windows.Forms.DockStyle.Top;
            this.groupBox1.Location = new System.Drawing.Point(0, 67);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(896, 194);
            this.groupBox1.TabIndex = 18;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "변환 전 파일 정보";
            // 
            // tbCreated
            // 
            this.tbCreated.BackColor = System.Drawing.SystemColors.ButtonHighlight;
            this.tbCreated.Location = new System.Drawing.Point(141, 142);
            this.tbCreated.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.tbCreated.Name = "tbCreated";
            this.tbCreated.ReadOnly = true;
            this.tbCreated.Size = new System.Drawing.Size(361, 25);
            this.tbCreated.TabIndex = 20;
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(83, 145);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(57, 15);
            this.label5.TabIndex = 19;
            this.label5.Text = "작성일:";
            // 
            // tbModifier
            // 
            this.tbModifier.BackColor = System.Drawing.SystemColors.ButtonHighlight;
            this.tbModifier.Location = new System.Drawing.Point(141, 109);
            this.tbModifier.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.tbModifier.Name = "tbModifier";
            this.tbModifier.ReadOnly = true;
            this.tbModifier.Size = new System.Drawing.Size(361, 25);
            this.tbModifier.TabIndex = 18;
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(84, 112);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(57, 15);
            this.label4.TabIndex = 17;
            this.label4.Text = "변경자:";
            // 
            // tbAuthor
            // 
            this.tbAuthor.BackColor = System.Drawing.SystemColors.ButtonHighlight;
            this.tbAuthor.Location = new System.Drawing.Point(141, 75);
            this.tbAuthor.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.tbAuthor.Name = "tbAuthor";
            this.tbAuthor.ReadOnly = true;
            this.tbAuthor.Size = new System.Drawing.Size(361, 25);
            this.tbAuthor.TabIndex = 16;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(83, 79);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(57, 15);
            this.label3.TabIndex = 15;
            this.label3.Text = "작성자:";
            // 
            // tbBefore
            // 
            this.tbBefore.BackColor = System.Drawing.SystemColors.ButtonHighlight;
            this.tbBefore.Location = new System.Drawing.Point(141, 25);
            this.tbBefore.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.tbBefore.Name = "tbBefore";
            this.tbBefore.ReadOnly = true;
            this.tbBefore.Size = new System.Drawing.Size(647, 25);
            this.tbBefore.TabIndex = 8;
            // 
            // btBefore
            // 
            this.btBefore.Location = new System.Drawing.Point(795, 22);
            this.btBefore.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.btBefore.Name = "btBefore";
            this.btBefore.Size = new System.Drawing.Size(86, 29);
            this.btBefore.TabIndex = 7;
            this.btBefore.Text = "파일선택";
            this.btBefore.UseVisualStyleBackColor = true;
            this.btBefore.Click += new System.EventHandler(this.btBefore_Click);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(83, 28);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(52, 15);
            this.label1.TabIndex = 6;
            this.label1.Text = "파일명";
            // 
            // groupBox2
            // 
            this.groupBox2.Controls.Add(this.tbNumberOfRowPerPage);
            this.groupBox2.Controls.Add(this.tbStartRowNumber);
            this.groupBox2.Controls.Add(this.tbHeightOfRow);
            this.groupBox2.Controls.Add(this.label8);
            this.groupBox2.Controls.Add(this.label7);
            this.groupBox2.Controls.Add(this.label6);
            this.groupBox2.Controls.Add(this.cbAfterOpen);
            this.groupBox2.Controls.Add(this.btAfter);
            this.groupBox2.Controls.Add(this.tbAfter);
            this.groupBox2.Controls.Add(this.label2);
            this.groupBox2.Dock = System.Windows.Forms.DockStyle.Fill;
            this.groupBox2.Location = new System.Drawing.Point(0, 261);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Size = new System.Drawing.Size(896, 414);
            this.groupBox2.TabIndex = 19;
            this.groupBox2.TabStop = false;
            this.groupBox2.Text = "변환 후 파일 정보";
            // 
            // tbNumberOfRowPerPage
            // 
            this.tbNumberOfRowPerPage.Location = new System.Drawing.Point(400, 119);
            this.tbNumberOfRowPerPage.Name = "tbNumberOfRowPerPage";
            this.tbNumberOfRowPerPage.Size = new System.Drawing.Size(100, 25);
            this.tbNumberOfRowPerPage.TabIndex = 28;
            this.tbNumberOfRowPerPage.Text = "17";
            // 
            // tbStartRowNumber
            // 
            this.tbStartRowNumber.Location = new System.Drawing.Point(141, 119);
            this.tbStartRowNumber.Name = "tbStartRowNumber";
            this.tbStartRowNumber.Size = new System.Drawing.Size(100, 25);
            this.tbStartRowNumber.TabIndex = 27;
            this.tbStartRowNumber.Text = "5";
            // 
            // tbHeightOfRow
            // 
            this.tbHeightOfRow.Location = new System.Drawing.Point(639, 119);
            this.tbHeightOfRow.Name = "tbHeightOfRow";
            this.tbHeightOfRow.Size = new System.Drawing.Size(100, 25);
            this.tbHeightOfRow.TabIndex = 26;
            this.tbHeightOfRow.Text = "12.5";
            // 
            // label8
            // 
            this.label8.AutoSize = true;
            this.label8.Location = new System.Drawing.Point(516, 124);
            this.label8.Name = "label8";
            this.label8.Size = new System.Drawing.Size(117, 15);
            this.label8.TabIndex = 25;
            this.label8.Text = "변환 후 행 높이:";
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Location = new System.Drawing.Point(262, 124);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(132, 15);
            this.label7.TabIndex = 23;
            this.label7.Text = "페이지 당 행 갯수:";
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Location = new System.Drawing.Point(73, 124);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(62, 15);
            this.label6.TabIndex = 21;
            this.label6.Text = "시작 행:";
            // 
            // cbAfterOpen
            // 
            this.cbAfterOpen.AutoSize = true;
            this.cbAfterOpen.Checked = true;
            this.cbAfterOpen.CheckState = System.Windows.Forms.CheckState.Checked;
            this.cbAfterOpen.Location = new System.Drawing.Point(141, 83);
            this.cbAfterOpen.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.cbAfterOpen.Name = "cbAfterOpen";
            this.cbAfterOpen.Size = new System.Drawing.Size(209, 19);
            this.cbAfterOpen.TabIndex = 19;
            this.cbAfterOpen.Text = "변환 후 엑셀파일 자동실행";
            this.cbAfterOpen.UseVisualStyleBackColor = true;
            // 
            // btAfter
            // 
            this.btAfter.Location = new System.Drawing.Point(795, 33);
            this.btAfter.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.btAfter.Name = "btAfter";
            this.btAfter.Size = new System.Drawing.Size(86, 29);
            this.btAfter.TabIndex = 18;
            this.btAfter.Text = "변환시작";
            this.btAfter.UseVisualStyleBackColor = true;
            this.btAfter.Click += new System.EventHandler(this.btAfter_Click);
            // 
            // tbAfter
            // 
            this.tbAfter.BackColor = System.Drawing.SystemColors.ButtonHighlight;
            this.tbAfter.Location = new System.Drawing.Point(141, 36);
            this.tbAfter.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.tbAfter.Name = "tbAfter";
            this.tbAfter.ReadOnly = true;
            this.tbAfter.Size = new System.Drawing.Size(647, 25);
            this.tbAfter.TabIndex = 17;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(83, 40);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(52, 15);
            this.label2.TabIndex = 16;
            this.label2.Text = "파일명";
            // 
            // MainForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 15F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(896, 701);
            this.Controls.Add(this.groupBox2);
            this.Controls.Add(this.groupBox1);
            this.Controls.Add(this.panel1);
            this.Controls.Add(this.ssStatusBar);
            this.Controls.Add(this.msMainMenu);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MainMenuStrip = this.msMainMenu;
            this.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.MinimumSize = new System.Drawing.Size(912, 238);
            this.Name = "MainForm";
            this.Text = "엑셀 파일 변환 작업";
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.MainForm_FormClosing);
            this.Load += new System.EventHandler(this.MainForm_Load);
            this.ssStatusBar.ResumeLayout(false);
            this.ssStatusBar.PerformLayout();
            this.msMainMenu.ResumeLayout(false);
            this.msMainMenu.PerformLayout();
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.groupBox2.ResumeLayout(false);
            this.groupBox2.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion
        private System.Windows.Forms.StatusStrip ssStatusBar;
        private System.Windows.Forms.ToolStripStatusLabel status_message;
        private System.Windows.Forms.ToolStripProgressBar tsProgressBar;
        private System.Windows.Forms.MenuStrip msMainMenu;
        private System.Windows.Forms.ToolStripMenuItem 파일FToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem miExit;
        private System.Windows.Forms.ToolStripMenuItem tsHelp;
        private System.Windows.Forms.ToolStripMenuItem miAbout;
        private System.Windows.Forms.OpenFileDialog ofBeforeExcel;
        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.TextBox tbCreated;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.TextBox tbModifier;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.TextBox tbAuthor;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.TextBox tbBefore;
        private System.Windows.Forms.Button btBefore;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.GroupBox groupBox2;
        private System.Windows.Forms.CheckBox cbAfterOpen;
        private System.Windows.Forms.Button btAfter;
        private System.Windows.Forms.TextBox tbAfter;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.TextBox tbHeightOfRow;
        private System.Windows.Forms.Label label8;
        private System.Windows.Forms.TextBox tbNumberOfRowPerPage;
        private System.Windows.Forms.TextBox tbStartRowNumber;
    }
}


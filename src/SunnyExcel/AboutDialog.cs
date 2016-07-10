using System;
using System.Windows.Forms;

namespace SunnyExcel
{
    public partial class AboutDialog : Form
    {
        private MainForm __parent_form = null;

        public AboutDialog(Form p_parent_form)
        {
            InitializeComponent();

            __parent_form = (MainForm)p_parent_form;

            lnOdinSoft.Links.Add(0, 8, "http://www.odinsoftware.co.kr");
            lnOpenXmlSdk.Links.Add(8, 12, "https://github.com/OfficeDev/Open-XML-SDK");
        }

        private void btOK_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void AboutDialog_Load(object sender, EventArgs e)
        {
            __parent_form.RestoreFormLayout(this);
        }

        private void AboutDialog_FormClosing(object sender, FormClosingEventArgs e)
        {
            __parent_form.SaveFormLayout(this);
        }

        private void AboutDialog_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Escape)
            {
                this.DialogResult = DialogResult.Cancel;
                Close();
            }
        }

        private void WebLinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            System.Diagnostics.Process.Start(e.Link.LinkData.ToString());
        }
    }
}

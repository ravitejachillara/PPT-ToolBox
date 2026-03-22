using System;
using System.Windows.Forms;

namespace PPTToolbox
{
    public partial class AboutDialog : Form
    {
        public AboutDialog()
        {
            InitializeComponent();
        }

        private void btnClose_Click(object sender, EventArgs e) => Close();
    }
}

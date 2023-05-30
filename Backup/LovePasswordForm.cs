using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

namespace Students
{
    public partial class LovePasswordForm : Form
    {
        public LovePasswordForm()
        {
            InitializeComponent();
        }

        private void btnLove_Click(object sender, EventArgs e)
        {
            if (txbLove.Text.Trim() == "17042010")
            {
                this.DialogResult = DialogResult.OK;
            }
            else MessageBox.Show("Пароль невірний ! ! !", "П о м и л к а !", MessageBoxButtons.OK, MessageBoxIcon.Stop);
        }
    }
}

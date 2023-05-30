using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.Text;
using System.Windows.Forms;


namespace Students
{
    public partial class AddPayForm : Form
    {
        public AddPayForm()
        {
            InitializeComponent();
        }

        # region const
        const string strSqlPay = "Select ID,MNH from MNTH";
        # endregion
        
        # region var
        public string newMnthPay = string.Empty;
        public string newPay = string.Empty; 
        private OleDbCommand cmd = new OleDbCommand();
        private BindingSource bsPay = new BindingSource();
        # endregion
        private void AddPayForm_Load(object sender, EventArgs e)
        {
            try
            {
                this.Cursor = Cursors.WaitCursor; 
                (Owner as MainForm).dbConnect((Owner as MainForm).cn);
                if ((Owner as MainForm).cn.State == ConnectionState.Open)
                {
                    cmd.Connection = (Owner as MainForm).cn;                
                    cmd.CommandText = strSqlPay;
                    bsPay.DataSource = cmd.ExecuteReader();
                    cmbPay.DataSource = bsPay;
                    cmbPay.DisplayMember = "MNH";
                    cmbPay.ValueMember = "ID";
                    cmbPay.SelectedIndex = 0;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Помилка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                cmd.Connection.Close();
                this.Cursor = Cursors.Default; 
            } 
        }

        private void btnOk_Click(object sender, EventArgs e)
        {
            newMnthPay = cmbPay.SelectedValue.ToString();
            newPay = txbPay.Text.Replace("'", "`");
            this.DialogResult = DialogResult.OK;
        }
    }
}

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
    public partial class AddSubForm : Form
    {
        public AddSubForm()
        {
            InitializeComponent();
        }

         # region const
        const string strSqlSub = "Select ID,SUB from SUBJECT";
        # endregion
        
        # region var
        public string newSub = string.Empty; 
        private OleDbCommand cmd = new OleDbCommand();
        private BindingSource bsSub = new BindingSource();
        # endregion
        
        private void AddSubForm_Load(object sender, EventArgs e)
        {
            try
            {
                this.Cursor = Cursors.WaitCursor; 
                (Owner as MainForm).dbConnect((Owner as MainForm).cn);
                if ((Owner as MainForm).cn.State == ConnectionState.Open)
                {
                    cmd.Connection = (Owner as MainForm).cn;
                    cmd.CommandText = strSqlSub;
                    bsSub.DataSource = cmd.ExecuteReader();
                    cmbSub.DataSource = bsSub;
                    cmbSub.DisplayMember = "SUB";
                    cmbSub.ValueMember = "ID";
                    cmbSub.SelectedIndex = 0;
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
            newSub = cmbSub.SelectedValue.ToString();
            this.DialogResult = DialogResult.OK;
        }
    }
}

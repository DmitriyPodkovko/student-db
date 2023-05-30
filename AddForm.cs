using System;
using System.Data;
using System.Data.OleDb;
using System.Windows.Forms;

namespace Students
{
    public partial class AddForm : Form
    {
        public AddForm()
        {
            InitializeComponent();
        }

        # region const
        const string strSqlSex = "Select ID,XSEX from SEX";
        const string strSqlUkr = "Select ID,REG from UKR";
        const string strSqlEdu = "Select ID,XEDU from EDUCAT";
        const string strSqlSts = "Select ID,XSTS from STATUS";
        # endregion

        # region var
        public bool Find;
        public string curSex = string.Empty;
        public string curAddr = string.Empty;
        public string curSchl = string.Empty;
        public string curEdu = string.Empty;
        public string curSts = string.Empty;
        public string newF = string.Empty;
        public string newI = string.Empty;
        public string newO = string.Empty;
        public string newSex = string.Empty;
        public string newDB = string.Empty;
        public string newDR = string.Empty;
        public string newDoc = string.Empty;
        public string newRegAddr = string.Empty;
        public string newAddr = string.Empty;
        public string newRegSchl = string.Empty;
        public string newSchl = string.Empty;
        public string newEdu = string.Empty;
        public string newSts = string.Empty;
        public string newRprt = string.Empty;
        private OleDbCommand cmd = new OleDbCommand();
        private BindingSource bsSex = new BindingSource();
        private BindingSource bsAddr = new BindingSource();
        private BindingSource bsSchl = new BindingSource();
        private BindingSource bsEdu = new BindingSource();
        private BindingSource bsSts = new BindingSource();
        # endregion

        private void AddForm_Load(object sender, EventArgs e)
        {
            try
            {
                this.Cursor = Cursors.WaitCursor; 
                (Owner as MainForm).dbConnect((Owner as MainForm).cn);
                if ((Owner as MainForm).cn.State == ConnectionState.Open)
                {
                    cmd.Connection = (Owner as MainForm).cn;
                    cmd.CommandText = strSqlSex;
                    bsSex.DataSource = cmd.ExecuteReader();
                    cmbSex.DataSource = bsSex;
                    cmbSex.DisplayMember = "XSEX";
                    cmbSex.ValueMember = "ID";
                    if (!string.IsNullOrEmpty(curSex))
                    {
                        cmbSex.Text = curSex;
                    }
                    else cmbSex.SelectedIndex = 0;
                    cmd.Connection.Close();

                    cmd.Connection.Open();
                    cmd.CommandText = strSqlUkr;
                    bsAddr.DataSource = cmd.ExecuteReader();
                    cmbAddr.DataSource = bsAddr;
                    cmbAddr.DisplayMember = "REG";
                    cmbAddr.ValueMember = "ID";
                    if (!string.IsNullOrEmpty(curAddr))
                    {
                        cmbAddr.Text = curAddr;
                    }
                    else cmbAddr.SelectedIndex = 0;
                    cmd.Connection.Close();

                    cmd.Connection.Open();
                    bsSchl.DataSource = cmd.ExecuteReader();
                    cmbSchl.DataSource = bsSchl;
                    cmbSchl.DisplayMember = "REG";
                    cmbSchl.ValueMember = "ID";
                    if (!string.IsNullOrEmpty(curSchl))
                    {
                        cmbSchl.Text = curSchl;
                    }
                    else cmbSchl.SelectedIndex = 0;
                    cmd.Connection.Close();

                    cmd.Connection.Open();
                    cmd.CommandText = strSqlEdu;
                    bsEdu.DataSource = cmd.ExecuteReader();
                    cmbEdu.DataSource = bsEdu;
                    cmbEdu.DisplayMember = "XEDU";
                    cmbEdu.ValueMember = "ID";
                    if (!string.IsNullOrEmpty(curEdu))
                    {
                        cmbEdu.Text = curEdu;
                    }
                    else cmbEdu.SelectedIndex = 0;
                    cmd.Connection.Close();

                    cmd.Connection.Open();
                    cmd.CommandText = strSqlSts;
                    bsSts.DataSource = cmd.ExecuteReader();
                    cmbSts.DataSource = bsSts; 
                    cmbSts.DisplayMember = "XSTS";
                    cmbSts.ValueMember = "ID";
                    if (!string.IsNullOrEmpty(curSts))
                    {
                        cmbSts.Text = curSts;
                    }
                    else cmbSts.SelectedIndex = 0;
                    cmd.Connection.Close();
                    
                    //OleDbCommand myCommand = new OleDbCommand(qSelSex, (Owner as MainForm).cn);
                    //OleDbDataReader myDataReader = myCommand.ExecuteReader();
                    // while (myDataReader.Read())
                    // {
                    //  comBxSex.Items.Add(myDataReader.GetValue(1));
                    //  }
                    // comBxSex.SelectedIndex = 0;
                    // myDataReader.Close();
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
            if (txbF.Enabled) newF = txbF.Text.Trim().Replace( "'", "`");
            if (txbI.Enabled) newI = txbI.Text.Trim().Replace("'", "`");
            if (txbO.Enabled) newO = txbO.Text.Trim().Replace("'", "`");
            if (cmbSex.Enabled) newSex = cmbSex.SelectedValue.ToString();
            if (dtpDB.Enabled) newDB = dtpDB.Value.ToString().Substring(0, 10);
            if (dtpDR.Enabled) newDR = dtpDR.Value.ToString().Substring(0, 10);
            if (txbDoc.Enabled) newDoc = txbDoc.Text.Trim().Replace("'", "`");
            if (cmbAddr.Enabled) newRegAddr = cmbAddr.SelectedValue.ToString();
            if (txbAddr.Enabled) newAddr = txbAddr.Text.Trim().Replace("'", "`");
            if (cmbSchl.Enabled) newRegSchl = cmbSchl.SelectedValue.ToString();
            if (txbSchl.Enabled) newSchl = txbSchl.Text.Trim().Replace("'", "`");
            if (cmbEdu.Enabled) newEdu = cmbEdu.SelectedValue.ToString();
            if (cmbSts.Enabled) newSts = cmbSts.SelectedValue.ToString();
            if (txbRprt.Enabled) newRprt = txbRprt.Text.Trim().Replace("'", "`");
            if (Find == false) 
            {
                if (string.IsNullOrEmpty(newF) || string.IsNullOrEmpty(newI))
                {
                    MessageBox.Show("Прізвище та Ім`я повинні бути заповнені ! ! !", "Увага!",
                                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }
                else this.DialogResult = DialogResult.OK;
            } 
            else this.DialogResult = DialogResult.OK;
            
        }

        private void chbF_CheckedChanged(object sender, EventArgs e)
        {
            if (chbF.Checked) {txbF.Enabled = true;} else txbF.Enabled = false;  
        }

        private void chbI_CheckedChanged(object sender, EventArgs e)
        {
            if (chbI.Checked) {txbI.Enabled = true;} else txbI.Enabled = false;  
        }

        private void chbO_CheckedChanged(object sender, EventArgs e)
        {
            if (chbO.Checked) {txbO.Enabled = true;} else txbO.Enabled = false;  
        }

        private void chbDB_CheckedChanged(object sender, EventArgs e)
        {
            if (chbDB.Checked) {dtpDB.Enabled = true;} else dtpDB.Enabled = false;  
        }

        private void chbDR_CheckedChanged(object sender, EventArgs e)
        {
            if (chbDR.Checked) {dtpDR.Enabled = true;} else dtpDR.Enabled = false;
        }
        
        private void chbSex_CheckedChanged(object sender, EventArgs e)
        {
            if (chbSex.Checked) {cmbSex.Enabled = true;} else cmbSex.Enabled = false;  
        }

        private void chbDoc_CheckedChanged(object sender, EventArgs e)
        {
            if (chbDoc.Checked) {txbDoc.Enabled = true;} else txbDoc.Enabled = false;  
        }

        private void chbEdu_CheckedChanged(object sender, EventArgs e)
        {
            if (chbEdu.Checked) { cmbEdu.Enabled = true; } else cmbEdu.Enabled = false;
        }
        
        private void chbAddr_CheckedChanged(object sender, EventArgs e)
        {
            if (chbAddr.Checked) {cmbAddr.Enabled = true; txbAddr.Enabled = true;}
                            else {cmbAddr.Enabled = false; txbAddr.Enabled = false;};  
        }

        private void chbSchl_CheckedChanged(object sender, EventArgs e)
        {
            if (chbSchl.Checked) {cmbSchl.Enabled = true; txbSchl.Enabled = true; }
                            else {cmbSchl.Enabled = false; txbSchl.Enabled = false; };  
        }

        private void chbSts_CheckedChanged(object sender, EventArgs e)
        {
            if (chbSts.Checked) {cmbSts.Enabled = true; txbRprt.Enabled = true;}
                           else {cmbSts.Enabled = false; txbRprt.Enabled = false;};  
        }

    }
}

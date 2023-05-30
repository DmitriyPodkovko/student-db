using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace Students
{
    public partial class MainForm : Form
    {
        public MainForm()
        {
            InitializeComponent();
        }

        # region const
        const string qSelMain = "Select MAIN.SN,DR,F,I,O,XSEX,DB,DOC,REG,ADDR from (((MAIN left outer join ADDRES on MAIN.SN=ADDRES.SN_IO) left outer join SEX on MAIN.SEX=SEX.ID) left outer join UKR on ADDRES.IDREG=UKR.ID)";
        const string qSelSchl = "Select REG,SCHL from ((MAIN left outer join SCHOOL on MAIN.SN=SCHOOL.SN_IO) left outer join UKR on SCHOOL.IDREG=UKR.ID) where SN_IO=";
        const string qSelSts = "Select XEDU,XSTS,RPRT from ((MAIN left outer join STATUS on MAIN.STS=STATUS.ID) left outer join EDUCAT on MAIN.EDU=EDUCAT.ID) where SN=";
        const string qSelPay = "Select PAY.SN,MNH,PAY from ((MAIN left outer join PAY on MAIN.SN=PAY.SN_IO) left outer join MNTH on PAY.IDMNH=MNTH.ID) where SN_IO=";
        const string qSelSub = "Select SUBPERS.SN,SUB from ((MAIN left outer join SUBPERS on MAIN.SN=SUBPERS.SN_IO) left outer join SUBJECT on SUBPERS.IDSUB=SUBJECT.ID) where SN_IO=";
        const string strSqlSelMax = "Select max(SN) from Main";
        const string strSqlInsMain = "Insert into MAIN (DR,F,I,O,SEX,DB,DOC,EDU,STS,RPRT) values ('{0}','{1}','{2}','{3}',{4},'{5}','{6}',{7},{8},'{9}')";
        const string strSqlInsAddres = "Insert into ADDRES (SN,SN_IO,IDREG,ADDR) values ({0},{0},{1},'{2}')";
        const string strSqlInsSchool = "Insert into SCHOOL (SN,SN_IO,IDREG,SCHL) values ({0},{0},{1},'{2}')";
        const string strSqlInsPay = "Insert into PAY (SN_IO,IDMNH,PAY) values ({0},{1},'{2}')";
        const string strSqlInsSub = "Insert into SUBPERS (SN_IO,IDSUB) values ({0},{1})";
        const string strSqlUpdMain = "Update MAIN set DR='{1}',F='{2}',I='{3}',O='{4}',SEX={5},DB='{6}',DOC='{7}',EDU={8},STS={9},RPRT='{10}' where SN={0}";
        const string strSqlUpdAddres = "Update ADDRES set IDREG={1},ADDR='{2}' where SN_IO={0}";
        const string strSqlUpdSchool = "Update SCHOOL set IDREG={1},SCHL='{2}' where SN_IO={0}";
        const string strSqlDel = "Delete from MAIN where SN=";
        const string strSqlDelPay = "Delete from PAY where SN=";
        const string strSqlDelSub = "Delete from SUBPERS where SN=";
        const string strSqlFindSub = "Select distinct MAIN.SN,DR,F,I,O,XSEX,DB,DOC,REG,ADDR from ((((MAIN left outer join ADDRES on MAIN.SN=ADDRES.SN_IO) left outer join SEX on MAIN.SEX=SEX.ID) left outer join UKR on ADDRES.IDREG=UKR.ID) left outer join SUBPERS on MAIN.SN=SUBPERS.SN_IO) where IDSUB=";
        const string strSqlFindPay = "Select distinct MAIN.SN,DR,F,I,O,XSEX,DB,DOC,REG,ADDR from ((((MAIN left outer join ADDRES on MAIN.SN=ADDRES.SN_IO) left outer join SEX on MAIN.SEX=SEX.ID) left outer join UKR on ADDRES.IDREG=UKR.ID) left outer join PAY on MAIN.SN=PAY.SN_IO) where MAIN.SN not in (Select distinct PAY.SN_IO from PAY where IDMNH={0})";
        const string qMain = "MAIN";
        const string qSchl = "SCHL";
        const string qSts = "STS";
        const string qPay = "PAY";
        const string qSub = "SUB";
        # endregion

        # region var
        private string csACEConnStr = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source="+Application.StartupPath+"\\Students.mdb";
        private string SNMain = string.Empty;
        private OleDbCommand cmd = new OleDbCommand();
        public OleDbConnection cn = new OleDbConnection();
        private DataSet DS = new DataSet();
        private OleDbDataAdapter Adptr = new OleDbDataAdapter();
        private OleDbDataAdapter Adptr2 = new OleDbDataAdapter();
        private BindingSource bsMain = new BindingSource();
        private BindingSource bsSchl = new BindingSource();
        private BindingSource bsSts = new BindingSource();
        private BindingSource bsPay = new BindingSource();
        private BindingSource bsSub = new BindingSource();
        # endregion

        //Подключение к БД
        private void MainForm_Load(object sender, EventArgs e)
        {
            dbConnect(cn);
            ShowList(qSelMain);
            dbDisConnect(cn);
        }
        
        public void dbConnect(OleDbConnection conn)
        {
            try
            {
                conn.ConnectionString = csACEConnStr;
                conn.Open();

            }
            catch (OleDbException ex)
            {
                MessageBox.Show(ex.Message, "Помилка підключення до БД",
                                MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public void dbDisConnect(OleDbConnection conn)
        {
            try
            {
                conn.Close();

            }
            catch (OleDbException ex)
            {
                MessageBox.Show(ex.Message, "Помилка відключення від БД",
                                MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        void ShowList(string qSQL)
        {
            try
            {
                this.Cursor = Cursors.WaitCursor;
                //aMain.SelectCommand  = new OleDbCommand(string.Format(qSelAll, qMain), cn);
                Adptr.SelectCommand = new OleDbCommand(qSQL, cn);
                if (DS.Tables.Contains(qMain) == false) DS.Tables.Add(qMain); DS.Tables[qMain].Clear();
                if (DS.Tables.Contains(qSchl) == false) DS.Tables.Add(qSchl); DS.Tables[qSchl].Clear();
                if (DS.Tables.Contains(qSts) == false) DS.Tables.Add(qSts); DS.Tables[qSts].Clear();
                if (DS.Tables.Contains(qSub) == false) DS.Tables.Add(qSub); DS.Tables[qSub].Clear();
                if (DS.Tables.Contains(qPay) == false) DS.Tables.Add(qPay); DS.Tables[qPay].Clear();
                bsMain.DataSource = DS.Tables[qMain];
                Adptr.Fill(DS, qMain);
                if (bsMain.Count > 0)
                {
                    dgvMain.DataSource = bsMain;
                    dgvMain.Columns[0].Visible = false;
                    dgvMain.Columns[1].HeaderText = "Дата реєстрації";
                    dgvMain.Columns[2].HeaderText = "Прізвище";
                    dgvMain.Columns[3].HeaderText = "Ім'я";
                    dgvMain.Columns[4].HeaderText = "По-батькові";
                    dgvMain.Columns[5].HeaderText = "Стать";
                    dgvMain.Columns[6].HeaderText = "Дата народження";
                    dgvMain.Columns[7].HeaderText = "Паспорт";
                    dgvMain.Columns[8].HeaderText = "Регіон";
                    dgvMain.Columns[9].HeaderText = "Місце проживання";
                    dgvMain.Columns[1].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
                    dgvMain.Columns[2].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
                    dgvMain.Columns[3].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
                    dgvMain.Columns[4].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
                    dgvMain.Columns[5].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
                    dgvMain.Columns[6].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
                    dgvMain.Columns[7].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
                    dgvMain.Columns[8].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
                    dgvMain.Columns[9].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
                }    
                dgvMain.Select();
                statusStrip.Items[0].Text = "Всього відібрано "+dgvMain.RowCount.ToString()+" слухачів(ча) підготовчого відділення";
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Помилка відбору слухачів",
                                MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                this.Cursor = Cursors.Default;
            }
        }

        void RefreshSchl()
        {
            Adptr2.SelectCommand = new OleDbCommand(qSelSchl + SNMain, cn);
            if (DS.Tables.Contains(qSchl) == false) DS.Tables.Add(qSchl);
            DS.Tables[qSchl].Clear();
            bsSchl.DataSource = DS.Tables[qSchl];
            Adptr2.Fill(DS, qSchl);
            dgvSchl.DataSource = bsSchl;
            dgvSchl.Columns[0].HeaderText = "Регіон";
            dgvSchl.Columns[1].HeaderText = "Навчальний заклад";
            dgvSchl.Columns[0].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
            dgvSchl.Columns[1].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
        }

        void RefreshSts()
        {
            Adptr2.SelectCommand.CommandText = qSelSts + SNMain;
            if (DS.Tables.Contains(qSts) == false) DS.Tables.Add(qSts);
            DS.Tables[qSts].Clear();
            bsSts.DataSource = DS.Tables[qSts];
            Adptr2.Fill(DS, qSts);
            dgvSts.DataSource = bsSts;
            dgvSts.Columns[0].HeaderText = "Форма навчання";
            dgvSts.Columns[1].HeaderText = "Статус";
            dgvSts.Columns[2].HeaderText = "Наказ";
            dgvSts.Columns[0].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
            dgvSts.Columns[1].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
            dgvSts.Columns[2].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
        }

        void RefreshSub()
        {
            Adptr2.SelectCommand.CommandText = qSelSub + SNMain;
            if (DS.Tables.Contains(qSub) == false) DS.Tables.Add(qSub);
            DS.Tables[qSub].Clear();
            bsSub.DataSource = DS.Tables[qSub];
            Adptr2.Fill(DS, qSub);
            if (bsSub.Count > 0)
            {
                dgvSub.DataSource = bsSub;
                dgvSub.Columns[0].Visible = false;
                dgvSub.Columns[1].HeaderText = "Предмети";
                dgvSub.Columns[1].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
            } 
        }

        void RefreshPay()
        {
            Adptr2.SelectCommand.CommandText = qSelPay + SNMain;
            if (DS.Tables.Contains(qPay) == false) DS.Tables.Add(qPay);
            DS.Tables[qPay].Clear();
            bsPay.DataSource = DS.Tables[qPay];
            Adptr2.Fill(DS, qPay);
            if (bsPay.Count > 0)
            {
                dgvPay.DataSource = bsPay;
                dgvPay.Columns[0].Visible = false;
                dgvPay.Columns[1].HeaderText = "Місяць";
                dgvPay.Columns[2].HeaderText = "Оплата";
                dgvPay.Columns[1].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
                dgvPay.Columns[2].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
            }
        }
        
        private void dgvMain_RowEnter(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex > -1)
            {
                try
                {
                    this.Cursor = Cursors.WaitCursor; 
                    SNMain = dgvMain.Rows[e.RowIndex].Cells["SN"].Value.ToString();
                    RefreshSchl();
                    RefreshSts();
                    RefreshPay();
                    RefreshSub();
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message, "Помилка відбору реквізитів слухачів",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                finally
                {
                    this.Cursor = Cursors.Default;
                }
            }
        }

        void addDgvMain(string newDR, string newF, string newI, string newO, string newSex, 
                        string newDB, string newDoc, string newRegAddr, string newAddr, 
                        string newRegSchl, string newSchl, string newEdu, string newSts, string newRprt)
        {
            try
            {
                this.Cursor = Cursors.WaitCursor;       
                cmd.Connection = cn;
                cmd.Connection.Open();
                cmd.CommandText = string.Format(strSqlInsMain, newDR, newF, newI, newO, newSex, newDB, newDoc, newEdu, newSts, newRprt);
                cmd.ExecuteReader();
                cmd.Connection.Close();

                cmd.Connection.Open();
                cmd.CommandText = strSqlSelMax;
                string maxSN = cmd.ExecuteScalar().ToString();
                cmd.Connection.Close();

                cmd.Connection.Open();
                cmd.CommandText = string.Format(strSqlInsAddres, maxSN, newRegAddr, newAddr);
                cmd.ExecuteReader();
                cmd.Connection.Close();

                cmd.Connection.Open();
                cmd.CommandText = string.Format(strSqlInsSchool, maxSN, newRegSchl, newSchl);
                cmd.ExecuteReader();

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Помилка додавання слухача", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                cmd.Connection.Close();
                this.Cursor = Cursors.Default;
            }
        }

        void updDgvMain(string newDR, string newF, string newI, string newO, string newSex, 
                        string newDB, string newDoc, string newRegAddr, string newAddr, 
                        string newRegSchl, string newSchl, string newEdu, string newSts, string newRprt)
        {
            try
            {
                this.Cursor = Cursors.WaitCursor; 
                cmd.Connection = cn;
                cmd.Connection.Open();
                cmd.CommandText = string.Format(strSqlUpdMain, SNMain, newDR, newF, newI, newO, newSex, newDB, newDoc, newEdu, newSts, newRprt);
                cmd.ExecuteReader();
                cmd.Connection.Close();

                cmd.Connection.Open();
                cmd.CommandText = string.Format(strSqlUpdAddres, SNMain, newRegAddr, newAddr);
                cmd.ExecuteReader();
                cmd.Connection.Close();

                cmd.Connection.Open();
                cmd.CommandText = string.Format(strSqlUpdSchool, SNMain, newRegSchl, newSchl);
                cmd.ExecuteReader();

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Помилка зміни реквізитів слухача", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                cmd.Connection.Close();
                this.Cursor = Cursors.Default;
            }
        }

        private void btnAdd_Click(object sender, EventArgs e)
        {
            AddForm fmAdd = new AddForm();
            if (fmAdd.ShowDialog(this) == DialogResult.OK)
            {
                addDgvMain(fmAdd.newDR, fmAdd.newF, fmAdd.newI, fmAdd.newO, fmAdd.newSex,
                           fmAdd.newDB, fmAdd.newDoc, fmAdd.newRegAddr, fmAdd.newAddr,
                           fmAdd.newRegSchl, fmAdd.newSchl, fmAdd.newEdu, fmAdd.newSts, fmAdd.newRprt);
                ShowList(qSelMain);
            }
            fmAdd.Dispose();
        }

        private void btnUpdate_Click(object sender, EventArgs e)
        {
            if (dgvMain.RowCount > 0)
            {
                AddForm fmAdd = new AddForm();
                fmAdd.dtpDR.Text = dgvMain.CurrentRow.Cells[1].Value.ToString();
                fmAdd.txbF.Text = dgvMain.CurrentRow.Cells[2].Value.ToString();
                fmAdd.txbI.Text = dgvMain.CurrentRow.Cells[3].Value.ToString();
                fmAdd.txbO.Text = dgvMain.CurrentRow.Cells[4].Value.ToString();
                fmAdd.curSex = dgvMain.CurrentRow.Cells[5].Value.ToString();
                fmAdd.dtpDB.Text = dgvMain.CurrentRow.Cells[6].Value.ToString();
                fmAdd.txbDoc.Text = dgvMain.CurrentRow.Cells[7].Value.ToString();
                fmAdd.curAddr = dgvMain.CurrentRow.Cells[8].Value.ToString();
                fmAdd.txbAddr.Text = dgvMain.CurrentRow.Cells[9].Value.ToString();
                fmAdd.curSchl = dgvSchl.CurrentRow.Cells[0].Value.ToString();
                fmAdd.txbSchl.Text = dgvSchl.CurrentRow.Cells[1].Value.ToString();
                fmAdd.curEdu = dgvSts.CurrentRow.Cells[0].Value.ToString();
                fmAdd.curSts = dgvSts.CurrentRow.Cells[1].Value.ToString();
                fmAdd.txbRprt.Text = dgvSts.CurrentRow.Cells[2].Value.ToString();
                if (fmAdd.ShowDialog(this) == DialogResult.OK)
                {
                    updDgvMain(fmAdd.newDR, fmAdd.newF, fmAdd.newI, fmAdd.newO, fmAdd.newSex,
                               fmAdd.newDB, fmAdd.newDoc, fmAdd.newRegAddr, fmAdd.newAddr,
                               fmAdd.newRegSchl, fmAdd.newSchl, fmAdd.newEdu, fmAdd.newSts, fmAdd.newRprt);
                    ShowList(qSelMain);
                }
                fmAdd.Dispose();
            } else MessageBox.Show("Немає жодного слухача для редагування!", "Попередження",
                                    MessageBoxButtons.OK, MessageBoxIcon.Information);    
        }

        private void btnDelete_Click(object sender, EventArgs e)
        {
            if (dgvMain.RowCount > 0)
            {
                if (MessageBox.Show("Ви дійсно бажаєте видалити вибраного слухача?", "Попередження",
                                     MessageBoxButtons.OKCancel, MessageBoxIcon.Warning) == DialogResult.OK)
                {
                    try
                    {
                        this.Cursor = Cursors.WaitCursor;
                        cmd.Connection = cn;
                        cmd.Connection.Open();
                        cmd.CommandText = strSqlDel + SNMain;
                        cmd.ExecuteReader();
                        ShowList(qSelMain);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message, "Помилка видалення слухача", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                    finally
                    {
                        cmd.Connection.Close();
                        this.Cursor = Cursors.Default;
                    }
                }
            } else MessageBox.Show("Немає жодного слухача для видалення!", "Попередження",
                                    MessageBoxButtons.OK, MessageBoxIcon.Information);                
        }

        string Qut(string s)
        {
            return s = "'" + s + "'";
        }

        string N(string s)
        {
            if (s.StartsWith("0"))
            {
                s = s.Substring(3, 2) + "/" + s.Substring(0, 2) + "/" + s.Substring(6, 4);
            }
            else s = s.Replace(".", "/");
            return s = "#" + s + "#";
        }
        
        string BuildSql(string Col, string Val)
        {
            string s;
            if (!string.IsNullOrEmpty(Val.Trim())) 
            {
                Val = Val.Replace("*", "%").Replace("?", "_");
                if (Val.IndexOfAny(new char[] { '%', '_' }) > 0) s = Col + " LIKE " + Val;
                                                            else s = Col + " = " + Val;
            }
            else s = string.Empty;
            return s;
        }
        
        void findDgvMain(string newDR, string newF, string newI, string newO, string newSex,
                         string newDB, string newDoc, string newRegAddr, string newAddr,
                         string newRegSchl, string newSchl, string newEdu, string newSts, string newRprt)
        {
            this.Cursor = Cursors.WaitCursor;
            const string s1 = "Select MAIN.SN,DR,F,I,O,XSEX,DB,DOC,REG,ADDR from (((";
            const string s2 = "MAIN left outer join ADDRES on MAIN.SN=ADDRES.SN_IO) left outer join SEX on MAIN.SEX=SEX.ID) left outer join UKR on ADDRES.IDREG=UKR.ID)";
            const string s3 = "left outer join SCHOOL on MAIN.SN=SCHOOL.SN_IO)";
            const string where = " where ";
            string sql = string.Empty;
            
            if (!string.IsNullOrEmpty(newF)) sql = BuildSql("F", Qut(newF));
            
            if (!string.IsNullOrEmpty(sql) & !string.IsNullOrEmpty(newI)) sql = sql + " and " + BuildSql("I", Qut(newI));
            else if (string.IsNullOrEmpty(sql) & !string.IsNullOrEmpty(newI)) sql = BuildSql("I", Qut(newI));
            
            if (!string.IsNullOrEmpty(sql) & !string.IsNullOrEmpty(newO)) sql = sql + " and " + BuildSql("O", Qut(newO));
            else if (string.IsNullOrEmpty(sql) & !string.IsNullOrEmpty(newO)) sql = BuildSql("O", Qut(newO));
            
            if (!string.IsNullOrEmpty(sql) & !string.IsNullOrEmpty(newSex)) sql = sql + " and " + BuildSql("SEX", newSex);
            else if (string.IsNullOrEmpty(sql) & !string.IsNullOrEmpty(newSex)) sql = BuildSql("SEX", newSex);

            if (!string.IsNullOrEmpty(sql) & !string.IsNullOrEmpty(newDB)) sql = sql + " and " + BuildSql("DB", N(newDB));
            else if (string.IsNullOrEmpty(sql) & !string.IsNullOrEmpty(newDB)) sql = BuildSql("DB", N(newDB));

            if (!string.IsNullOrEmpty(sql) & !string.IsNullOrEmpty(newDR)) sql = sql + " and " + BuildSql("DR", N(newDR.Replace('.', '/')));
            else if (string.IsNullOrEmpty(sql) & !string.IsNullOrEmpty(newDR)) sql = BuildSql("DR", N(newDR.Replace('.', '/')));

            if (!string.IsNullOrEmpty(sql) & !string.IsNullOrEmpty(newDoc)) sql = sql + " and " + BuildSql("DOC", Qut(newDoc));
            else if (string.IsNullOrEmpty(sql) & !string.IsNullOrEmpty(newDoc)) sql = BuildSql("DOC", Qut(newDoc));
            
            if (!string.IsNullOrEmpty(sql) & !string.IsNullOrEmpty(newEdu)) sql = sql + " and " + BuildSql("EDU", newEdu);
            else if (string.IsNullOrEmpty(sql) & !string.IsNullOrEmpty(newEdu)) sql = BuildSql("EDU", newEdu);

            if (!string.IsNullOrEmpty(sql) & !string.IsNullOrEmpty(newSts)) sql = sql + " and " + BuildSql("STS", newSts);
            else if (string.IsNullOrEmpty(sql) & !string.IsNullOrEmpty(newSts)) sql = BuildSql("STS", newSts);

            if (!string.IsNullOrEmpty(sql) & !string.IsNullOrEmpty(newRprt)) sql = sql + " and " + BuildSql("RPRT", Qut(newRprt));
            else if (string.IsNullOrEmpty(sql) & !string.IsNullOrEmpty(newRprt)) sql = BuildSql("RPRT", Qut(newRprt)); 
            
            if (!string.IsNullOrEmpty(sql) & !string.IsNullOrEmpty(newRegAddr)) sql = sql + " and " + BuildSql("ADDRES.IDREG", newRegAddr);
            else if (string.IsNullOrEmpty(sql) & !string.IsNullOrEmpty(newRegAddr)) sql = BuildSql("ADDRES.IDREG", newRegAddr);

            if (!string.IsNullOrEmpty(sql) & !string.IsNullOrEmpty(newAddr)) sql = sql + " and " + BuildSql("ADDR", Qut(newAddr));
            else if (string.IsNullOrEmpty(sql) & !string.IsNullOrEmpty(newAddr)) sql = BuildSql("ADDR", Qut(newAddr));

            if (!string.IsNullOrEmpty(sql) & !string.IsNullOrEmpty(newRegSchl)) sql = sql + " and " + BuildSql("SCHOOL.IDREG", newRegSchl);
            else if (string.IsNullOrEmpty(sql) & !string.IsNullOrEmpty(newRegSchl)) sql = BuildSql("SCHOOL.IDREG", newRegSchl);

            if (!string.IsNullOrEmpty(sql) & !string.IsNullOrEmpty(newSchl)) sql = sql + " and " + BuildSql("SCHL", Qut(newSchl));
            else if (string.IsNullOrEmpty(sql) & !string.IsNullOrEmpty(newSchl)) sql = BuildSql("SCHL", Qut(newSchl));

            if (!string.IsNullOrEmpty(sql) & !string.IsNullOrEmpty(newRegSchl)) sql = s1 + "(" + s2 + s3 + where + sql;
            else if (!string.IsNullOrEmpty(sql)) sql = s1 + s2 + where + sql;
            else sql = qSelMain;
            this.Cursor = Cursors.Default;
            ShowList(sql);
        }
        
        private void btnFind_Click(object sender, EventArgs e)
        {
            AddForm fmAdd = new AddForm();
            fmAdd.txbF.Enabled = false;     fmAdd.chbF.Visible = true;
            fmAdd.txbI.Enabled = false;     fmAdd.chbI.Visible = true;
            fmAdd.txbO.Enabled = false;     fmAdd.chbO.Visible = true;
            fmAdd.dtpDB.Enabled = false;    fmAdd.chbDB.Visible = true;
            fmAdd.dtpDR.Enabled = false;    fmAdd.chbDR.Visible = true; 
            fmAdd.cmbSex.Enabled = false;   fmAdd.chbSex.Visible = true;
            fmAdd.txbDoc.Enabled = false;   fmAdd.chbDoc.Visible = true;
            fmAdd.cmbAddr.Enabled = false;
            fmAdd.txbAddr.Enabled = false;  fmAdd.chbAddr.Visible = true;
            fmAdd.cmbSchl.Enabled = false;
            fmAdd.txbSchl.Enabled = false;  fmAdd.chbSchl.Visible = true;
            fmAdd.cmbEdu.Enabled = false;   fmAdd.chbEdu.Visible = true;
            fmAdd.cmbSts.Enabled = false;
            fmAdd.txbRprt.Enabled = false;  fmAdd.chbSts.Visible = true;
            fmAdd.Find = true;
            if (fmAdd.ShowDialog(this) == DialogResult.OK)
            {
                findDgvMain(fmAdd.newDR, fmAdd.newF, fmAdd.newI, fmAdd.newO, fmAdd.newSex,
                            fmAdd.newDB, fmAdd.newDoc, fmAdd.newRegAddr, fmAdd.newAddr,
                            fmAdd.newRegSchl, fmAdd.newSchl, fmAdd.newEdu, fmAdd.newSts, fmAdd.newRprt);
            }
            fmAdd.Dispose();
        }
        
        private void btnAddSub_Click(object sender, EventArgs e)
        {
            if (dgvMain.RowCount > 0)
            {
                AddSubForm fmAddSub = new AddSubForm();
                if (fmAddSub.ShowDialog(this) == DialogResult.OK)
                {
                    try
                    {
                        this.Cursor = Cursors.WaitCursor;
                        cmd.Connection = cn;
                        cmd.Connection.Open();
                        cmd.CommandText = string.Format(strSqlInsSub, SNMain, fmAddSub.newSub);
                        cmd.ExecuteReader();
                        RefreshSub();
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message, "Помилка додавання предмета слухача", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                    finally
                    {
                        cmd.Connection.Close();
                        this.Cursor = Cursors.Default;
                    }
                }
                fmAddSub.Dispose();
            } else MessageBox.Show("Немає жодного слухача для додавання предмета!", "Попередження",
                                    MessageBoxButtons.OK, MessageBoxIcon.Information);    
        }

        private void btnAddPay_Click(object sender, EventArgs e)
        {
            if (dgvMain.RowCount > 0)
            {
                AddPayForm fmAddPay = new AddPayForm();
                if (fmAddPay.ShowDialog(this) == DialogResult.OK)
                {
                    try
                    {
                        this.Cursor = Cursors.WaitCursor;
                        cmd.Connection = cn;
                        cmd.Connection.Open();
                        cmd.CommandText = string.Format(strSqlInsPay, SNMain, fmAddPay.newMnthPay, fmAddPay.newPay);
                        cmd.ExecuteReader();
                        RefreshPay();
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message, "Помилка додавання оплати слухача", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                    finally
                    {
                        cmd.Connection.Close();
                        this.Cursor = Cursors.Default;
                    }
                }
                fmAddPay.Dispose();
            } else MessageBox.Show("Немає жодного слухача для додавання оплати!", "Попередження",
                                    MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void btnDelSub_Click(object sender, EventArgs e)
        {
            if (dgvSub.RowCount > 0)
            {
                if (MessageBox.Show("Ви дійсно бажаєте видалити вибраний предмет?", "Попередження",
                                     MessageBoxButtons.OKCancel, MessageBoxIcon.Question) == DialogResult.OK)
                {
                    try
                    {
                        this.Cursor = Cursors.WaitCursor;
                        cmd.Connection = cn;
                        cmd.Connection.Open();
                        cmd.CommandText = strSqlDelSub + dgvSub.CurrentRow.Cells[0].Value.ToString();
                        cmd.ExecuteReader();
                        RefreshSub();
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message, "Помилка видалення предмета слухача", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                    finally
                    {
                        cmd.Connection.Close();
                        this.Cursor = Cursors.Default;
                    }
                }
            } else MessageBox.Show("Немає жодного слухача для видалення предмета!", "Попередження",
                                    MessageBoxButtons.OK, MessageBoxIcon.Information);    
        }
   
        private void btnDelPay_Click(object sender, EventArgs e)
        {
            if (dgvPay.RowCount > 0)
            {
                if (MessageBox.Show("Ви дійсно бажаєте видалити вибрану оплату?", "Попередження",
                                     MessageBoxButtons.OKCancel, MessageBoxIcon.Question) == DialogResult.OK)
                {
                    try
                    {
                        this.Cursor = Cursors.WaitCursor;
                        cmd.Connection = cn;
                        cmd.Connection.Open();
                        cmd.CommandText = strSqlDelPay + dgvPay.CurrentRow.Cells[0].Value.ToString();
                        cmd.ExecuteReader();
                        RefreshPay();
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message, "Помилка видалення оплати слухача", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                    finally
                    {
                        cmd.Connection.Close();
                        this.Cursor = Cursors.Default;
                    }
                }
            } else MessageBox.Show("Немає жодного слухача для видалення оплати!", "Попередження",
                                    MessageBoxButtons.OK, MessageBoxIcon.Information);    
       }

        private void btnFindSub_Click(object sender, EventArgs e)
        {
            AddSubForm fmAddSub = new AddSubForm();
            if (fmAddSub.ShowDialog(this) == DialogResult.OK)
            {
                ShowList(strSqlFindSub + fmAddSub.newSub);
            }
            fmAddSub.Dispose();
        }

        private void btnFindPay_Click(object sender, EventArgs e)
        {
            AddPayForm fmAddPay = new AddPayForm();
            fmAddPay.txbPay.ReadOnly = true;
            fmAddPay.txbPay.Text = "Відсутня оплата за";
            if (fmAddPay.ShowDialog(this) == DialogResult.OK)
            {
                ShowList(string.Format(strSqlFindPay, fmAddPay.newMnthPay));
            }
            fmAddPay.Dispose();
        }
        
        private void releaseObject(object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch (Exception ex)
            {
                obj = null;
                MessageBox.Show("Exception Occured while releasing object " + ex.ToString());
            }
            finally
            {
                GC.Collect();
            }
        }
        
        private void btnExcel_Click(object sender, EventArgs e)
        {
            if (dgvMain.RowCount > 0)
            {
                this.Cursor = Cursors.WaitCursor; 
                Excel.Application xlApp;
                Excel.Workbook xlWorkBook;
                Excel.Worksheet xlWorkSheet;
                object misValue = System.Reflection.Missing.Value;

                xlApp = new Excel.ApplicationClass();
                xlWorkBook = xlApp.Workbooks.Add(misValue);
                xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);

                for (int i = 0; i <= dgvMain.RowCount - 1; i++)
                {   
                    for (int j = 0; j < dgvMain.ColumnCount - 1; j++)
                    {
                        DataGridViewCell cell = dgvMain[j + 1, i];
                        xlWorkSheet.Cells[i + 1, j + 1] = cell.Value;
                    }
                }

                xlWorkBook.SaveAs(Application.StartupPath + "\\Students.xls", Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue, Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
                xlWorkBook.Close(true, misValue, misValue);
                xlApp.Quit();

                releaseObject(xlWorkSheet);
                releaseObject(xlWorkBook);
                releaseObject(xlApp);

                this.Cursor = Cursors.Default; 
                MessageBox.Show("Дані слухачів експортовано у файл " + Application.StartupPath + "\\Students.xls",
                                "Інформація", MessageBoxButtons.OK, MessageBoxIcon.Information);            
            }
            else MessageBox.Show("Немає жодного слухача для експорту у файл!", "Попередження",
                         MessageBoxButtons.OK, MessageBoxIcon.Information); 
        }
        
        private void btnLove_Click(object sender, EventArgs e)
        {
            LovePasswordForm LovePasswordForm = new LovePasswordForm();
            if (LovePasswordForm.ShowDialog(this) == DialogResult.OK)
            {
                this.Cursor = Cursors.WaitCursor;
                LovePasswordForm.Dispose(); 
                LoveForm LoveForm = new LoveForm();
                LoveForm.ShowDialog(this);
                LoveForm.Dispose();
                this.Cursor = Cursors.Default;
            }
            else LovePasswordForm.Dispose();
        }
        
        private void MainForm_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (MessageBox.Show("Ви дійсно бажаєте вийти з інформаційної системи?", "Попередження",
                                 MessageBoxButtons.OKCancel, MessageBoxIcon.Question) == DialogResult.Cancel)
            {e.Cancel = true;}
        }

      }
}

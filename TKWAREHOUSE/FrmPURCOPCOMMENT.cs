using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using NPOI;
using NPOI.HPSF;
using NPOI.HSSF;
using NPOI.HSSF.UserModel;
using NPOI.POIFS;
using NPOI.Util;
using NPOI.HSSF.Util;
using NPOI.HSSF.Extractor;
using System.IO;
using System.Data.SqlClient;
using NPOI.SS.UserModel;
using System.Configuration;
using NPOI.XSSF.UserModel;
using System.Text.RegularExpressions;
using FastReport;
using FastReport.Data;
using System.Collections;

namespace TKWAREHOUSE
{
    public partial class FrmPURCOPCOMMENT : Form
    {

        SqlConnection sqlConn = new SqlConnection();
        SqlCommand sqlComm = new SqlCommand();
        string connectionString;
        StringBuilder sbSql = new StringBuilder();
        StringBuilder sbSqlQuery = new StringBuilder();
        SqlDataAdapter adapter = new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilder = new SqlCommandBuilder();
        SqlDataAdapter adapter2 = new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilder2 = new SqlCommandBuilder();
        SqlDataAdapter adapter3 = new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilder3 = new SqlCommandBuilder();
        SqlDataAdapter adapter4 = new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilder4 = new SqlCommandBuilder();
        SqlDataAdapter adapter5 = new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilder5 = new SqlCommandBuilder();
        SqlDataAdapter adapter6 = new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilder6= new SqlCommandBuilder();

        DataSet ds = new DataSet();
        DataSet ds2 = new DataSet();
        DataSet ds3 = new DataSet();
        DataSet ds4 = new DataSet();
        DataSet ds5 = new DataSet();
        DataSet ds6 = new DataSet();

        DataTable ADDDT = new DataTable();

        int result;
        string tablename = null;
        SqlTransaction tran;
        SqlCommand cmd = new SqlCommand();

        string PURTA001;
        string PURTA002;
        string COPTC001;
        string COPTC002;

        public FrmPURCOPCOMMENT()
        {
            InitializeComponent();

            ADDDT.Columns.AddRange(new DataColumn[2] {
                 new DataColumn("訂單單別", typeof(string)),
                 new DataColumn("訂單單號", typeof(string))
            });
        }

        #region FUNCTION
        private void FrmPURCOPCOMMENT_Load(object sender, EventArgs e)
        {
            SETGRIDVIEW();
        }

        public void SETGRIDVIEW()
        {
            dataGridView1.AlternatingRowsDefaultCellStyle.BackColor = Color.PaleTurquoise;      //奇數列顏色
            dataGridView2.AlternatingRowsDefaultCellStyle.BackColor = Color.PaleTurquoise;
            dataGridView3.AlternatingRowsDefaultCellStyle.BackColor = Color.PaleTurquoise;
            dataGridView4.AlternatingRowsDefaultCellStyle.BackColor = Color.PaleTurquoise;


            //先建立個 CheckBox 欄
            DataGridViewCheckBoxColumn cbCol = new DataGridViewCheckBoxColumn();
            cbCol.Width = 120;   //設定寬度
            cbCol.HeaderText = "";
            cbCol.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;   //置中
            cbCol.TrueValue = true;
            cbCol.FalseValue = false;
            dataGridView3.Columns.Insert(0, cbCol);

            #region 建立全选 CheckBox

            //建立个矩形，等下计算 CheckBox 嵌入 GridView 的位置
            Rectangle rect = dataGridView3.GetCellDisplayRectangle(0, -1, true);
            rect.X = rect.Location.X + rect.Width / 4 - 18;
            rect.Y = rect.Location.Y + (rect.Height / 2 - 9);

            CheckBox cbHeader = new CheckBox();
            cbHeader.Name = "checkboxHeader";
            cbHeader.Size = new Size(18, 18);
            cbHeader.Location = rect.Location;

            //全选要设定的事件
            cbHeader.CheckedChanged += new EventHandler(cbHeader_CheckedChanged);

            //将 CheckBox 加入到 dataGridView
            dataGridView3.Controls.Add(cbHeader);


            #endregion
        }

        private void cbHeader_CheckedChanged(object sender, EventArgs e)
        {
            dataGridView1.EndEdit();

            foreach (DataGridViewRow dr in dataGridView1.Rows)
            {
                dr.Cells[0].Value = ((CheckBox)dataGridView1.Controls.Find("checkboxHeader", true)[0]).Checked;

            }

        }
        public void Search()
        {
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sbSql.Clear();
                sbSqlQuery.Clear();

                sbSql.AppendFormat(@"  SELECT TA001 AS '請購單別',TA002 AS '請購單號',TA006 AS '備註'");
                sbSql.AppendFormat(@"  FROM [TK].dbo.PURTA");
                sbSql.AppendFormat(@"  WHERE TA003>='{0}' AND TA003<='{1}'", dateTimePicker1.Value.ToString("yyyyMMdd"), dateTimePicker2.Value.ToString("yyyyMMdd"));
                sbSql.AppendFormat(@"  ORDER BY TA001,TA002");
                sbSql.AppendFormat(@"  ");

                adapter = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder = new SqlCommandBuilder(adapter);
                sqlConn.Open();
                ds.Clear();
                adapter.Fill(ds, "ds");
                sqlConn.Close();


                if (ds.Tables["ds"].Rows.Count == 0)
                {
                    dataGridView1.DataSource = null;
                }
                else
                {
                    if (ds.Tables["ds"].Rows.Count >= 1)
                    {
                        dataGridView1.DataSource = ds.Tables["ds"];
                        dataGridView1.AutoResizeColumns();
                    }

                }

            }
            catch
            {

            }
            finally
            {

            }

        }

        private void dataGridView1_SelectionChanged(object sender, EventArgs e)
        {
            if (dataGridView1.CurrentRow != null)
            {
                int rowindex = dataGridView1.CurrentRow.Index;
                if (rowindex >= 0)
                {
                    DataGridViewRow row = dataGridView1.Rows[rowindex];

                    PURTA001 = row.Cells["請購單別"].Value.ToString().Trim();
                    PURTA002 = row.Cells["請購單號"].Value.ToString().Trim();

                    textBox1.Text = row.Cells["請購單別"].Value.ToString().Trim();
                    textBox2.Text = row.Cells["請購單號"].Value.ToString().Trim();
                    textBox4.Text = row.Cells["請購單別"].Value.ToString().Trim();
                    textBox5.Text = row.Cells["請購單號"].Value.ToString().Trim();

                    Search2(PURTA001, PURTA002);
                }
                else
                {
                    PURTA001 = null;
                    PURTA002 = null;
                    textBox1.Text = null;
                    textBox2.Text = null;
                    textBox4.Text = null;
                    textBox5.Text = null;

                }
            }

        }

        public void Search2(string PURTA001, string PURTA002)
        {
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sbSql.Clear();
                sbSqlQuery.Clear();

                sbSql.AppendFormat(@"  SELECT [COPTC001] AS '訂單單別',[COPTC002] AS '訂單單號',[COMMENT] AS '備註',[PURTA001] AS '請購單別',[PURTA002] AS '請購單號',[ID] ");
                sbSql.AppendFormat(@"  FROM [TKWAREHOUSE].[dbo].[PURCOPCOMMENT]");
                sbSql.AppendFormat(@"  WHERE [PURTA001]='{0}' AND [PURTA002]='{1}' ", PURTA001, PURTA002);
                sbSql.AppendFormat(@"  AND [VISIABLE]='Y'");
                sbSql.AppendFormat(@"  ORDER BY [ID]");
                sbSql.AppendFormat(@"  ");

                adapter2 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder2 = new SqlCommandBuilder(adapter2);
                sqlConn.Open();
                ds2.Clear();
                adapter2.Fill(ds2, "ds2");
                sqlConn.Close();


                if (ds2.Tables["ds2"].Rows.Count == 0)
                {
                    dataGridView2.DataSource = null;
                }
                else
                {
                    if (ds2.Tables["ds2"].Rows.Count >= 1)
                    {
                        dataGridView2.DataSource = ds2.Tables["ds2"];
                        dataGridView2.AutoResizeColumns();
                    }

                }

            }
            catch
            {

            }
            finally
            {

            }

        }

        public void Search3()
        {
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sbSql.Clear();
                sbSqlQuery.Clear();

                sbSql.AppendFormat(@"  SELECT TC001 AS '訂單單別',TC002 AS '訂單單號',MV002  AS '業務',TC053 AS '客戶'");
                sbSql.AppendFormat(@"  FROM [TK].dbo.COPTC,[TK].dbo.CMSMV");
                sbSql.AppendFormat(@"  WHERE TC003>='{0}' AND TC003<='{1}'", dateTimePicker3.Value.ToString("yyyyMMdd"), dateTimePicker4.Value.ToString("yyyyMMdd"));
                sbSql.AppendFormat(@"  AND TC006=MV001");
                sbSql.AppendFormat(@"  ORDER BY TC001,TC002");
                sbSql.AppendFormat(@"  ");

                adapter3 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder2 = new SqlCommandBuilder(adapter3);
                sqlConn.Open();
                ds3.Clear();
                adapter3.Fill(ds3, "ds3");
                sqlConn.Close();


                if (ds3.Tables["ds3"].Rows.Count == 0)
                {
                    dataGridView3.DataSource = null;
                }
                else
                {
                    if (ds3.Tables["ds3"].Rows.Count >= 1)
                    {
                        dataGridView3.DataSource = ds3.Tables["ds3"];
                        dataGridView3.AutoResizeColumns();
                    }

                }

            }
            catch
            {

            }
            finally
            {

            }
        }

        public void Search4()
        {
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sbSql.Clear();
                sbSqlQuery.Clear();

                sbSql.AppendFormat(@"  SELECT TC001 AS '訂單單別',TC002 AS '訂單單號',MV002  AS '業務',TC053 AS '客戶'");
                sbSql.AppendFormat(@"  FROM [TK].dbo.COPTC,[TK].dbo.CMSMV");
                sbSql.AppendFormat(@"  WHERE TC003>='{0}' AND TC003<='{1}'", dateTimePicker5.Value.ToString("yyyyMMdd"), dateTimePicker6.Value.ToString("yyyyMMdd"));
                sbSql.AppendFormat(@"  AND TC006=MV001");
                sbSql.AppendFormat(@"  ORDER BY TC001,TC002");
                sbSql.AppendFormat(@"  ");

                adapter5 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder5 = new SqlCommandBuilder(adapter5);
                sqlConn.Open();
                ds5.Clear();
                adapter5.Fill(ds5, "ds5");
                sqlConn.Close();


                if (ds5.Tables["ds5"].Rows.Count == 0)
                {
                    dataGridView5.DataSource = null;
                }
                else
                {
                    if (ds5.Tables["ds5"].Rows.Count >= 1)
                    {
                        dataGridView5.DataSource = ds5.Tables["ds5"];
                        dataGridView5.AutoResizeColumns();
                    }

                }

            }
            catch
            {

            }
            finally
            {

            }
        }


        public void SEARCHCOP()
        {
            foreach (DataGridViewRow row in dataGridView3.Rows)
            {
                DataGridViewCheckBoxCell chk = (DataGridViewCheckBoxCell)row.Cells[0];
                if (chk.Value == chk.TrueValue)
                {
                    ADDDTROWS(row.Cells["訂單單別"].Value.ToString(), row.Cells["訂單單號"].Value.ToString());
                }

            }

            if (ADDDT.Rows.Count >= 1)
            {
                dataGridView4.DataSource = ADDDT;
            }
        }

        public void ADDDTROWS(string TC001, string TC002)
        {
            string COMPARE1;
            string COMPARE2 = TC001.Trim() + TC002.Trim();
            string CHECKADD = "N"; ;

            if (dataGridView4.Rows.Count > 0)
            {
                foreach (DataGridViewRow row in dataGridView4.Rows)
                {
                    COMPARE1 = row.Cells[0].Value.ToString().Trim() + row.Cells[1].Value.ToString().Trim();

                    if (COMPARE1.Equals(COMPARE2))
                    {
                        CHECKADD = "N";
                        break;
                    }
                    else
                    {
                        CHECKADD = "Y";
                    }
                }
            }
            else
            {
                ADDDT.Rows.Add(TC001, TC002);
            }

            if (CHECKADD.Equals("Y"))
            {
                ADDDT.Rows.Add(TC001, TC002);
            }

        }
        public void CLEARCOP()
        {
            ADDDT.Clear();
        }

        public void ADDPURCOPCOMMENT()
        {
            if (!string.IsNullOrEmpty(PURTA001) && !string.IsNullOrEmpty(PURTA002) && ADDDT.Rows.Count > 0)
            {
                ADDPURCOPCOMMENTDB(PURTA001, PURTA002, ADDDT);
            }
        }

        public void ADDPURCOPCOMMENTDB(string PURTA001, string PURTA002, DataTable TEMP)
        {
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sqlConn.Close();
                sqlConn.Open();
                tran = sqlConn.BeginTransaction();

                sbSql.Clear();

                foreach (DataRow dr in TEMP.Rows)
                {
                    sbSql.AppendFormat(" INSERT INTO  [TKWAREHOUSE].[dbo].[PURCOPCOMMENT]");
                    sbSql.AppendFormat(" ([PURTA001],[PURTA002],[COPTC001],[COPTC002],[COMMENT],[VISIABLE])");
                    sbSql.AppendFormat(" VALUES ('{0}','{1}','{2}','{3}','{4}','Y')", PURTA001, PURTA002, dr["訂單單別"].ToString(), dr["訂單單號"].ToString(), textBox3.Text.ToString());
                    sbSql.AppendFormat(" ");
                }
                sbSql.AppendFormat(" ");


                cmd.Connection = sqlConn;
                cmd.CommandTimeout = 60;
                cmd.CommandText = sbSql.ToString();
                cmd.Transaction = tran;
                result = cmd.ExecuteNonQuery();

                if (result == 0)
                {
                    tran.Rollback();    //交易取消
                }
                else
                {
                    tran.Commit();      //執行交易  


                }

            }
            catch
            {

            }

            finally
            {
                sqlConn.Close();
            }
        }

        private void dataGridView2_SelectionChanged(object sender, EventArgs e)
        {
            if (dataGridView2.CurrentRow != null)
            {
                int rowindex = dataGridView2.CurrentRow.Index;
                if (rowindex >= 0)
                {
                    DataGridViewRow row = dataGridView2.Rows[rowindex];

                    COPTC001 = row.Cells["訂單單別"].Value.ToString().Trim();
                    COPTC002 = row.Cells["訂單單號"].Value.ToString().Trim();
                    textBox6.Text = row.Cells["訂單單別"].Value.ToString().Trim();
                    textBox7.Text = row.Cells["訂單單號"].Value.ToString().Trim();

                }
                else
                {
                    COPTC001 = null;
                    COPTC002 = null;
                    textBox6.Text = null;
                    textBox7.Text = null;

                }
            }
        }

        public void PURCOPCOMMENT(string PURTA001, string PURTA002,string COPTC001,string COPTC002)
        {
            if (!string.IsNullOrEmpty(PURTA001) && !string.IsNullOrEmpty(PURTA002) && !string.IsNullOrEmpty(COPTC001) && !string.IsNullOrEmpty(COPTC002))
            {
                try
                {
                    connectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
                    sqlConn = new SqlConnection(connectionString);

                    sqlConn.Close();
                    sqlConn.Open();
                    tran = sqlConn.BeginTransaction();

                    sbSql.Clear();
                    
                    sbSql.AppendFormat(" UPDATE [TKWAREHOUSE].[dbo].[PURCOPCOMMENT]");
                    sbSql.AppendFormat(" SET [VISIABLE]='N'");
                    sbSql.AppendFormat(" WHERE [PURTA001]='{0}' AND [PURTA002]='{1}' AND [COPTC001]='{2}' AND [COPTC002]='{3}'",PURTA001,PURTA002,COPTC001,COPTC002);
                    sbSql.AppendFormat(" ");


                    cmd.Connection = sqlConn;
                    cmd.CommandTimeout = 60;
                    cmd.CommandText = sbSql.ToString();
                    cmd.Transaction = tran;
                    result = cmd.ExecuteNonQuery();

                    if (result == 0)
                    {
                        tran.Rollback();    //交易取消
                    }
                    else
                    {
                        tran.Commit();      //執行交易  


                    }

                }
                catch
                {

                }

                finally
                {
                    sqlConn.Close();
                }
            }
        }

        public void PURCOPCOMMENTALL(string PURTA001,string PURTA002)
        {
            if(!string.IsNullOrEmpty(PURTA001) && !string.IsNullOrEmpty(PURTA002))
            {
                try
                {
                    connectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
                    sqlConn = new SqlConnection(connectionString);

                    sqlConn.Close();
                    sqlConn.Open();
                    tran = sqlConn.BeginTransaction();

                    sbSql.Clear();

                    sbSql.AppendFormat(" UPDATE [TKWAREHOUSE].[dbo].[PURCOPCOMMENT]");
                    sbSql.AppendFormat(" SET [VISIABLE]='N'");
                    sbSql.AppendFormat(" WHERE [PURTA001]='{0}' AND [PURTA002]='{1}' ", PURTA001, PURTA002);
                    sbSql.AppendFormat(" ");


                    cmd.Connection = sqlConn;
                    cmd.CommandTimeout = 60;
                    cmd.CommandText = sbSql.ToString();
                    cmd.Transaction = tran;
                    result = cmd.ExecuteNonQuery();

                    if (result == 0)
                    {
                        tran.Rollback();    //交易取消
                    }
                    else
                    {
                        tran.Commit();      //執行交易  


                    }

                }
                catch
                {

                }

                finally
                {
                    sqlConn.Close();
                }
            }
        }

        #endregion

        #region BUTTON

        private void button4_Click(object sender, EventArgs e)
        {
            Search();
        }
        private void button1_Click(object sender, EventArgs e)
        {
            Search3();
        }
        private void button2_Click(object sender, EventArgs e)
        {
            SEARCHCOP();
        }

        private void button5_Click(object sender, EventArgs e)
        {
            CLEARCOP();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            ADDPURCOPCOMMENT();
            Search2(PURTA001, PURTA002);
        }


        private void button6_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("是否整張請購清空?", "是否整張請購清空?", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
            {
                PURCOPCOMMENTALL(textBox4.Text.Trim(), textBox5.Text.Trim());
                Search2(PURTA001, PURTA002);
            }
           
        }

        private void button7_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("是否請購的訂單清空?", "是否請購的訂單清空?", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
            {
                PURCOPCOMMENT(textBox4.Text.Trim(), textBox5.Text.Trim(), textBox6.Text.Trim(), textBox7.Text.Trim());
                Search2(PURTA001, PURTA002);
            }

                
        }
        private void button8_Click(object sender, EventArgs e)
        {
            Search4();
        }
        private void button9_Click(object sender, EventArgs e)
        {

        }
        #endregion


    }
}

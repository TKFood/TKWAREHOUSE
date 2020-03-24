﻿using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using System.Data.SqlClient;
using NPOI.SS.UserModel;
using System.Configuration;
using NPOI.XSSF.UserModel;
using NPOI.SS.Util;
using System.Reflection;
using System.Threading;
using System.Globalization;
using Calendar.NET;

namespace TKWAREHOUSE
{
    public partial class FrmCOPPURTA : Form
    {
        private ComponentResourceManager _ResourceManager = new ComponentResourceManager();
        SqlConnection sqlConn = new SqlConnection();
        SqlCommand sqlComm = new SqlCommand();
        string connectionString;
        StringBuilder sbSql = new StringBuilder();
        StringBuilder sbSqlQuery = new StringBuilder();
        SqlDataAdapter adapter1 = new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilder1 = new SqlCommandBuilder();
        SqlDataAdapter adapter2 = new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilder2 = new SqlCommandBuilder();
        SqlDataAdapter adapter3 = new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilder3 = new SqlCommandBuilder();
        SqlDataAdapter adapter4 = new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilder4 = new SqlCommandBuilder();

        SqlTransaction tran;
        SqlCommand cmd = new SqlCommand();
        DataSet ds1 = new DataSet();
        DataSet ds2 = new DataSet();
        DataSet ds3 = new DataSet();
        DataSet ds4 = new DataSet();


        DataTable dt = new DataTable();
        string tablename = null;
        int result;

        string ID;
        string NEWID;
        string DELTD001;
        string DELTD002;
        string DELTD003;

        public FrmCOPPURTA()
        {
            InitializeComponent();
        }

        #region FUNCTION

        public void SEARCHBTACHID()
        {
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sbSql.Clear();
                sbSqlQuery.Clear();

                sbSql.AppendFormat(@"  SELECT [ID] AS '批號',CONVERT(NVARCHAR,[BACTHDATES],112) AS '日期'");
                sbSql.AppendFormat(@"  FROM [TKWAREHOUSE].[dbo].[COPPURBATCHID]");
                sbSql.AppendFormat(@"  WHERE CONVERT(NVARCHAR,[BACTHDATES],112)='{0}'", dateTimePicker1.Value.ToString("yyyyMMdd"));
                sbSql.AppendFormat(@"  ORDER BY [ID] ");
                sbSql.AppendFormat(@"  ");

                adapter1 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder1 = new SqlCommandBuilder(adapter1);
                sqlConn.Open();
                ds1.Clear();
                adapter1.Fill(ds1, "TEMPds1");
                sqlConn.Close();


                if (ds1.Tables["TEMPds1"].Rows.Count == 0)
                {
                    dataGridView1.DataSource = null;
                }
                else
                {
                    if (ds1.Tables["TEMPds1"].Rows.Count >= 1)
                    {
                        //dataGridView1.Rows.Clear();
                        dataGridView1.DataSource = ds1.Tables["TEMPds1"];
                        dataGridView1.AutoResizeColumns();
                        //dataGridView1.CurrentCell = dataGridView1[0, rownum];

                    }
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

        public string GETMAXID()
        {
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                StringBuilder sbSql = new StringBuilder();
                sbSql.Clear();
                sbSqlQuery.Clear();
                ds2.Clear();

                sbSql.AppendFormat(@"  SELECT ISNULL(MAX(ID),'00000000000') AS ID");
                sbSql.AppendFormat(@"  FROM [TKWAREHOUSE].[dbo].[COPPURBATCHID]");
                sbSql.AppendFormat(@"  WHERE CONVERT(NVARCHAR,[BACTHDATES],112)='{0}'", dateTimePicker1.Value.ToString("yyyyMMdd"));
                sbSql.AppendFormat(@"  ");


                adapter2 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder2 = new SqlCommandBuilder(adapter2);
                sqlConn.Open();
                ds2.Clear();
                adapter2.Fill(ds2, "ds2");
                sqlConn.Close();


                if (ds2.Tables["ds2"].Rows.Count == 0)
                {
                    return null;
                }
                else
                {
                    if (ds2.Tables["ds2"].Rows.Count >= 1)
                    {
                        NEWID = SETID(ds2.Tables["ds2"].Rows[0]["ID"].ToString());
                        return NEWID;

                    }
                    return null;
                }

            }
            catch
            {
                return null;
            }
            finally
            {
                sqlConn.Close();
            }
        }

        public string SETID(string ID)
        {
            DateTime dt1 = dateTimePicker1.Value;

            if (ID.Equals("00000000000"))
            {
                return dt1.ToString("yyyyMMdd") + "001";
            }

            else
            {
                int serno = Convert.ToInt16(ID.Substring(8, 3));
                serno = serno + 1;
                string temp = serno.ToString();
                temp = temp.PadLeft(3, '0');
                return dt1.ToString("yyyyMMdd") + temp.ToString();
            }

        }

        public void ADDBTACHID(string ID)
        {
            if (!string.IsNullOrEmpty(ID))
            {
                try
                {
                    connectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
                    sqlConn = new SqlConnection(connectionString);

                    sqlConn.Close();
                    sqlConn.Open();
                    tran = sqlConn.BeginTransaction();

                    sbSql.Clear();

                    sbSql.AppendFormat(" INSERT INTO [TKWAREHOUSE].[dbo].[COPPURBATCHID]");
                    sbSql.AppendFormat(" ([ID],[BACTHDATES])");
                    sbSql.AppendFormat(" VALUES ('{0}','{1}')", ID, dateTimePicker1.Value.ToString("yyyyMMdd"));
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
        private void dataGridView1_SelectionChanged(object sender, EventArgs e)
        {
            textBoxID.Text = null;

            if (dataGridView1.CurrentRow != null)
            {
                int rowindex = dataGridView1.CurrentRow.Index;
                if (rowindex >= 0)
                {
                    DataGridViewRow row = dataGridView1.Rows[rowindex];
                    textBoxID.Text = row.Cells["批號"].Value.ToString();
                    ID = row.Cells["批號"].Value.ToString();

                    SEARCHCOPPURBATCHCOPTD(ID);

                }
                else
                {
                    textBoxID.Text = null;
                    ID = null;

                }
            }
        }

        public void SEARCHCOPPURBATCHCOPTD(string ID)
        {
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sbSql.Clear();
                sbSqlQuery.Clear();

                sbSql.AppendFormat(@"  SELECT [ID] AS '批號',[TD001] AS '訂單單別',[TD002] AS '訂單單號',[TD003] AS '訂單序號',[TD004] AS '品號',[TD005] AS '品名',[TD008] AS '訂單數量',[TD009] AS '已交數量',[TD010] AS '單位',[TD024] AS '贈品量',[TD025] AS '贈品已交量'");
                sbSql.AppendFormat(@"  FROM [TKWAREHOUSE].[dbo].[COPPURBATCHCOPTD]");
                sbSql.AppendFormat(@"  WHERE  [ID]='{0}'",ID);
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
                        //dataGridView1.Rows.Clear();
                        dataGridView2.DataSource = ds2.Tables["ds2"];
                        dataGridView2.AutoResizeColumns();
                        //dataGridView1.CurrentCell = dataGridView1[0, rownum];

                    }
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
        public void ADDCOPPURBATCHCOPTD(string ID,string TD001,string TD002,string TD003)
        {
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sqlConn.Close();
                sqlConn.Open();
                tran = sqlConn.BeginTransaction();

                sbSql.Clear();

                sbSql.AppendFormat(" INSERT INTO [TKWAREHOUSE].[dbo].[COPPURBATCHCOPTD]");
                sbSql.AppendFormat(" ([ID],[TD001],[TD002],[TD003],[TD004],[TD005],[TD008],[TD009],[TD010],[TD024],[TD025])");
                sbSql.AppendFormat(" SELECT '{0}',[TD001],[TD002],[TD003],[TD004],[TD005],[TD008],[TD009],[TD010],[TD024],[TD025]",ID);
                sbSql.AppendFormat(" FROM [TK].dbo.COPTD");
                sbSql.AppendFormat(" WHERE TD001='{0}' AND TD002='{1}' AND TD003='{2}'",TD001,TD002,TD003);
                sbSql.AppendFormat(" ");
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
            DELTD001 = null;
            DELTD002 = null;
            DELTD003 = null;

            if (dataGridView2.CurrentRow != null)
            {
                int rowindex = dataGridView2.CurrentRow.Index;
                if (rowindex >= 0)
                {
                    DataGridViewRow row = dataGridView2.Rows[rowindex];
                    DELTD001 = row.Cells["訂單單別"].Value.ToString();
                    DELTD002 = row.Cells["訂單單號"].Value.ToString();
                    DELTD003 = row.Cells["訂單序號"].Value.ToString();


                }
                else
                {
                    DELTD001 = null;
                    DELTD002 = null;
                    DELTD003 = null;

                }
            }
        }

        public void DELCOPPURBATCHCOPTD(string ID, string TD001, string TD002, string TD003)
        {
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sqlConn.Close();
                sqlConn.Open();
                tran = sqlConn.BeginTransaction();

                sbSql.Clear();

                sbSql.AppendFormat(" DELETE [TKWAREHOUSE].[dbo].[COPPURBATCHCOPTD]");
                sbSql.AppendFormat(" WHERE [ID]='{0}' AND TD001='{1}' AND TD002='{2}' AND TD003='{3}'",ID, TD001, TD002, TD003);
                sbSql.AppendFormat(" ");
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

        public void ADDCOPPURBATCHUSED(string ID)
        {
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sqlConn.Close();
                sqlConn.Open();
                tran = sqlConn.BeginTransaction();

                sbSql.Clear();

                sbSql.AppendFormat(" DELETE [TKWAREHOUSE].[dbo].[COPPURBATCHUSED]");
                sbSql.AppendFormat(" WHERE [ID]='{0}'", ID);
                sbSql.AppendFormat(" ");
                sbSql.AppendFormat(" INSERT INTO [TKWAREHOUSE].[dbo].[COPPURBATCHUSED]");
                sbSql.AppendFormat(" ([ID],[TD001],[TD002],[TD003],[TD004],[TD005],[TDNUM],[TDUNIT],[MB001],[MB002],[NUM],[UNIT])");
                sbSql.AppendFormat(" SELECT '{0}',TD001,TD002,TD003,TD004,TD005,NUM,MB004,MD003,MD035,(NUM*CAL),MD004", ID);
                sbSql.AppendFormat(" FROM (");
                sbSql.AppendFormat(" SELECT   TD001,TD002,TD003,TC053 ,TD013,TD004,TD005,TD006");
                sbSql.AppendFormat(" ,((CASE WHEN MB004=TD010 THEN ((TD008-TD009)+(TD024-TD025)) ELSE ((TD008-TD009)+(TD024-TD025))*INVMD.MD004 END)-ISNULL(MOCTA.TA017,0)) AS 'NUM'");
                sbSql.AppendFormat(" ,MB004");
                sbSql.AppendFormat(" ,((TD008-TD009)+(TD024-TD025)) AS 'COPNUM'");
                sbSql.AppendFormat(" ,TD010");
                sbSql.AppendFormat(" ,(CASE WHEN ISNULL(INVMD.MD002,'')<>'' THEN INVMD.MD002 ELSE TD010 END ) AS INVMDMD002");
                sbSql.AppendFormat(" ,(CASE WHEN INVMD.MD003>0 THEN INVMD.MD003 ELSE 1 END) AS INVMDMD003");
                sbSql.AppendFormat(" ,(CASE WHEN INVMD.MD004>0 THEN INVMD.MD004 ELSE (TD008-TD009) END ) AS INVMDMD004");
                sbSql.AppendFormat(" ,ISNULL(MOCTA.TA017,0) AS TA017");
                sbSql.AppendFormat(" ,[MC001],[MC004],BOMMD.[MD003],[MD035],BOMMD.[MD006],BOMMD.[MD007],BOMMD.[MD008],BOMMD.[MD004]");
                sbSql.AppendFormat(" ,CONVERT(decimal(16,3),(1/[MC004]*BOMMD.[MD006]/BOMMD.[MD007]*(1+BOMMD.[MD008]))) AS CAL");
                sbSql.AppendFormat(" FROM [TK].dbo.BOMMC,[TK].dbo.BOMMD,[TK].dbo.INVMB,[TK].dbo.COPTC,[TK].dbo.COPTD");
                sbSql.AppendFormat(" LEFT JOIN [TK].dbo.INVMD ON MD001=TD004 AND TD010=MD002");
                sbSql.AppendFormat(" LEFT JOIN [TK].dbo.MOCTA ON TA026=TD001 AND TA027=TD002 AND TD028=TD003 AND TA006=TD004");
                sbSql.AppendFormat(" WHERE BOMMC.MC001=BOMMD.MD001");
                sbSql.AppendFormat(" AND  BOMMD.MD001=TD004");
                sbSql.AppendFormat(" AND TD004=MB001");
                sbSql.AppendFormat(" AND TC001=TD001 AND TC002=TD002");
                sbSql.AppendFormat(" AND TD001+TD002+TD003 IN (SELECT TD001+TD002+TD003 FROM [TKWAREHOUSE].[dbo].[COPPURBATCHCOPTD] WHERE ID='{0}')",ID);
                sbSql.AppendFormat(" ) AS TEMP");
                sbSql.AppendFormat(" ");
                sbSql.AppendFormat(" ");
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

        public void SEARCHCOPPURBATCHUSED(string ID)
        {
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sbSql.Clear();
                sbSqlQuery.Clear();

                sbSql.AppendFormat(@"  SELECT [ID] AS '批號',[TD001] AS '訂單單別',[TD002] AS '訂單單號',[TD003] AS '訂單序號',[TD004] AS '成品',[TD005] AS '品名',[TDNUM] AS '訂單數量',[TDUNIT] AS '成品單位',[MB001] AS '用品',[MB002] AS '用品名',[NUM] AS '需求量',[UNIT] AS '需求單位'");
                sbSql.AppendFormat(@"  FROM [TKWAREHOUSE].[dbo].[COPPURBATCHUSED]");
                sbSql.AppendFormat(@"  WHERE [ID]='{0}'",ID);
                sbSql.AppendFormat(@"  ");

                adapter3 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder3 = new SqlCommandBuilder(adapter3);
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
                        //dataGridView1.Rows.Clear();
                        dataGridView3.DataSource = ds3.Tables["ds3"];
                        dataGridView3.AutoResizeColumns();
                        //dataGridView1.CurrentCell = dataGridView1[0, rownum];

                    }
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
        #endregion

        #region BUTTON
        private void button1_Click(object sender, EventArgs e)
        {
            SEARCHBTACHID();
            //SEARCHCOPPURBATCHCOPTD(ID);
        }
        private void button2_Click(object sender, EventArgs e)
        {
            ID = GETMAXID();
            ADDBTACHID(ID);
            SEARCHBTACHID();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            if(!string.IsNullOrEmpty(textBoxID.Text)&&!string.IsNullOrEmpty(textBox1.Text)&& !string.IsNullOrEmpty(textBox2.Text)&& !string.IsNullOrEmpty(textBox3.Text))
            {
                ADDCOPPURBATCHCOPTD(textBoxID.Text.Trim(), textBox1.Text.Trim(), textBox2.Text.Trim(), textBox3.Text.Trim());
            }

            SEARCHCOPPURBATCHCOPTD(textBoxID.Text.Trim());
        }
        private void button6_Click(object sender, EventArgs e)
        {
            DialogResult dialogResult = MessageBox.Show("要刪除了?", "要刪除了?", MessageBoxButtons.YesNo);
            if (dialogResult == DialogResult.Yes)
            {
                if (!string.IsNullOrEmpty(textBoxID.Text) && !string.IsNullOrEmpty(DELTD001) && !string.IsNullOrEmpty(DELTD002) && !string.IsNullOrEmpty(DELTD003))
                {
                    DELCOPPURBATCHCOPTD(textBoxID.Text.Trim(), DELTD001.Trim(), DELTD002.Trim(), DELTD003.Trim());
                }
            }

            SEARCHCOPPURBATCHCOPTD(textBoxID.Text.Trim());

        }
        private void button4_Click(object sender, EventArgs e)
        {
            ADDCOPPURBATCHUSED(textBoxID.Text.Trim());
            SEARCHCOPPURBATCHUSED(textBoxID.Text.Trim());
        }

        #endregion


    }
}

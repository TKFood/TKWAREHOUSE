using System;
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
    public partial class frmBacthMOC : Form
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

        SqlTransaction tran;
        SqlCommand cmd = new SqlCommand();
        DataSet ds1 = new DataSet();
        DataSet ds2 = new DataSet();
        DataSet ds3 = new DataSet();

        DataTable dt = new DataTable();
        string tablename = null;
        int result;

        string ID;
        string NEWID;

        public frmBacthMOC()
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
                sbSql.AppendFormat(@"  FROM [TKWAREHOUSE].[dbo].[BTACHID]");
                sbSql.AppendFormat(@"  WHERE CONVERT(NVARCHAR,[BACTHDATES],112)='{0}'",dateTimePicker1.Value.ToString("yyyyMMdd"));
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

                    SEARCHBACTHMOCTA(ID);
                }
                else
                {
                    textBoxID.Text = null;
                    ID = null;

                }
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
                sbSql.AppendFormat(@"  FROM [TKWAREHOUSE].[dbo].[BTACHID]");
                sbSql.AppendFormat(@"  WHERE CONVERT(NVARCHAR,[BACTHDATES],112)='{0}'",dateTimePicker1.Value.ToString("yyyyMMdd"));
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
            if(!string.IsNullOrEmpty(ID))
            {
                try
                {
                    connectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
                    sqlConn = new SqlConnection(connectionString);

                    sqlConn.Close();
                    sqlConn.Open();
                    tran = sqlConn.BeginTransaction();

                    sbSql.Clear();
                    
                    sbSql.AppendFormat(" INSERT INTO [TKWAREHOUSE].[dbo].[BTACHID]");
                    sbSql.AppendFormat(" ([ID],[BACTHDATES])");
                    sbSql.AppendFormat(" VALUES ('{0}','{1}')",ID,dateTimePicker1.Value.ToString("yyyyMMdd"));
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
        
        public void SEARCHBACTHMOCTA(string ID)
        {
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sbSql.Clear();
                sbSqlQuery.Clear();

                sbSql.AppendFormat(@"  SELECT [ID] AS '批號',[TA001] AS '製令',[TA002] AS '製令號',[TA006] AS '品號',[MB002] AS '品名',[TA017] AS '生產量',[UDF007] AS '淨重',[ATA017] AS '總重'");
                sbSql.AppendFormat(@"  FROM [TKWAREHOUSE].[dbo].[BACTHMOCTA]");
                sbSql.AppendFormat(@"  WHERE [ID]='{0}' ", ID);
                sbSql.AppendFormat(@"  ");
                sbSql.AppendFormat(@"  ");

                adapter3 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder3 = new SqlCommandBuilder(adapter3);
                sqlConn.Open();
                ds3.Clear();
                adapter3.Fill(ds3, "ds3");
                sqlConn.Close();


                if (ds3.Tables["ds3"].Rows.Count == 0)
                {
                    dataGridView2.DataSource = null;
                }
                else
                {
                    if (ds3.Tables["ds3"].Rows.Count >= 1)
                    {
                        //dataGridView1.Rows.Clear();
                        dataGridView2.DataSource = ds3.Tables["ds3"];
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

        public void ADDBACTHMOCTA(string ID)
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

                    sbSql.AppendFormat(" INSERT INTO [TKWAREHOUSE].[dbo].[BACTHMOCTA]");
                    sbSql.AppendFormat(" ([ID],[TA001],[TA002],[TA006],[MB002],[TA017],[UDF007],[ATA017])");
                    sbSql.AppendFormat(" SELECT '{0}',TA001,TA002,TA006,MB002,TA017,INVMB.UDF07,TA017*INVMB.UDF07",ID);
                    sbSql.AppendFormat(" FROM [TK].dbo.MOCTA,[TK].dbo.INVMB");
                    sbSql.AppendFormat(" WHERE TA006=MB001");
                    sbSql.AppendFormat(" AND TA001='{0}' AND TA002='{1}'",textBox1.Text,textBox2.Text);
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
        public void DELBACTHMOCTA()
        {

        }


        #endregion

        #region BUTTON
        private void button1_Click(object sender, EventArgs e)
        {
            SEARCHBTACHID();
        }
        private void button2_Click(object sender, EventArgs e)
        {
            ID=GETMAXID();
            ADDBTACHID(ID);
            SEARCHBTACHID();
            //MessageBox.Show(ID.ToString());
        }

        private void button4_Click(object sender, EventArgs e)
        {
            ADDBACTHMOCTA(textBoxID.Text);
            SEARCHBACTHMOCTA(textBoxID.Text);
        }

        private void button5_Click(object sender, EventArgs e)
        {
            DELBACTHMOCTA();
        }

        #endregion


    }
}

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
using TKITDLL;

namespace TKWAREHOUSE
{
    public partial class FrmINVSTAYOVERFIND : Form
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
        SqlTransaction tran;
        SqlCommand cmd = new SqlCommand();
        DataSet ds = new DataSet();
        DataSet ds2 = new DataSet();
        int result;
        public Report report1 { get; private set; }

        public FrmINVSTAYOVERFIND()
        {
            InitializeComponent();
        }

        #region FUNCTION
        public void SETFASTREPORT()
        {

            string SQL;
            report1 = new Report();
            report1.Load(@"REPORT\呆滯品追踨.frx");

            //20210902密
            Class1 TKID = new Class1();//用new 建立類別實體
            SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

            //資料庫使用者密碼解密
            sqlsb.Password = TKID.Decryption(sqlsb.Password);
            sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

            String connectionString;
            sqlConn = new SqlConnection(sqlsb.ConnectionString);

            report1.Dictionary.Connections[0].ConnectionString = sqlsb.ConnectionString;

            TableDataSource Table = report1.GetDataSource("Table") as TableDataSource;
            SQL = SETFASETSQL();
            Table.SelectCommand = SQL;
            report1.Preview = previewControl1;
            report1.Show();

        }

        public string SETFASETSQL()
        {
            StringBuilder FASTSQL = new StringBuilder();
            StringBuilder STRQUERY = new StringBuilder();

            FASTSQL.AppendFormat(@" SELECT CONVERT(NVARCHAR,[CHECKDATE],112) AS '檢查日期',[KIND] AS '分類',[MB001] AS '品號',[MB002] AS '品名',[MB003] AS '規格',[LOTNO] AS '批號',[NUM] AS '庫存數量',[COMMEMT] AS '處理方式'");
            FASTSQL.AppendFormat(@" FROM [TKWAREHOUSE].[dbo].[INVSTAYOVER]");
            FASTSQL.AppendFormat(@" WHERE [CHECKDATE]>='{0}' AND [CHECKDATE]<='{0}'", dateTimePicker1.Value.ToString("yyyy/MM/dd"), dateTimePicker1.Value.ToString("yyyy/MM/dd"));
            FASTSQL.AppendFormat(@" ORDER BY [KIND],[LOTNO],[MB001]");
            FASTSQL.AppendFormat(@" ");
            
            

            return FASTSQL.ToString();
        }

        public void SEARCHINVSTAYOVER()
        {
            try
            {
                //20210902密
                Class1 TKID = new Class1();//用new 建立類別實體
                SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

                //資料庫使用者密碼解密
                sqlsb.Password = TKID.Decryption(sqlsb.Password);
                sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

                String connectionString;
                sqlConn = new SqlConnection(sqlsb.ConnectionString);


                sbSql.Clear();
                sbSqlQuery.Clear();


                sbSql.AppendFormat(@"  SELECT CONVERT(NVARCHAR,[CHECKDATE],112) AS '檢查日期',[KIND] AS '分類',[MB001] AS '品號',[MB002] AS '品名',[MB003] AS '規格',[LOTNO] AS '批號',[NUM] AS '庫存數量',[COMMEMT] AS '處理方式'");
                sbSql.AppendFormat(@"  FROM [TKWAREHOUSE].[dbo].[INVSTAYOVER]");
                sbSql.AppendFormat(@"  WHERE [CHECKDATE]>='{0}' and  [CHECKDATE]<='{1}'",dateTimePicker2.Value.ToString("yyyy/MM/dd"), dateTimePicker3.Value.ToString("yyyy/MM/dd"));
                sbSql.AppendFormat(@"  ORDER BY [KIND],[LOTNO],[MB001]");
                sbSql.AppendFormat(@"  ");

                adapter2 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder2 = new SqlCommandBuilder(adapter2);
                sqlConn.Open();
                ds2.Clear();
                adapter2.Fill(ds2, "ds2");
                sqlConn.Close();


                if (ds2.Tables["ds2"].Rows.Count == 0)
                {
                    dataGridView1.DataSource = null;
                }
                else
                {
                    if (ds2.Tables["ds2"].Rows.Count >= 1)
                    {
                        dataGridView1.DataSource = ds2.Tables["ds2"];
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

        public void UPDATEINVSTAYOVER()
        {
            try
            {

                //20210902密
                Class1 TKID = new Class1();//用new 建立類別實體
                SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

                //資料庫使用者密碼解密
                sqlsb.Password = TKID.Decryption(sqlsb.Password);
                sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

                String connectionString;
                sqlConn = new SqlConnection(sqlsb.ConnectionString);


                sqlConn.Close();
                sqlConn.Open();
                tran = sqlConn.BeginTransaction();

                sbSql.Clear();

                sbSql.AppendFormat(" UPDATE  [TKWAREHOUSE].[dbo].[INVSTAYOVER]");
                sbSql.AppendFormat(" SET [COMMEMT]='{0}'", textBox8.Text);
                sbSql.AppendFormat(" WHERE CONVERT(NVARCHAR,CHECKDATE,112)='{0}' AND MB001='{1}' AND LOTNO='{2}'", textBox1.Text, textBox3.Text, textBox6.Text);
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


        private void dataGridView1_SelectionChanged(object sender, EventArgs e)
        {
            if (dataGridView1.CurrentRow != null)
            {
                int rowindex = dataGridView1.CurrentRow.Index;
                if (rowindex >= 0)
                {
                    DataGridViewRow row = dataGridView1.Rows[rowindex];

                    textBox1.Text = row.Cells["檢查日期"].Value.ToString();
                    textBox2.Text = row.Cells["分類"].Value.ToString();
                    textBox3.Text = row.Cells["品號"].Value.ToString();
                    textBox4.Text = row.Cells["品名"].Value.ToString();
                    textBox5.Text = row.Cells["規格"].Value.ToString();
                    textBox6.Text = row.Cells["批號"].Value.ToString();
                    textBox7.Text = row.Cells["庫存數量"].Value.ToString();
                    textBox8.Text = row.Cells["處理方式"].Value.ToString();
           

                }
                else
                {
                    textBox1.Text = null;
                    textBox2.Text = null;
                    textBox3.Text = null;
                    textBox4.Text = null;
                    textBox5.Text = null;
                    textBox6.Text = null;
                    textBox7.Text = null;
                    textBox8.Text = null;
                   

                }
            }
        }
        #endregion

        #region BUTTON

        private void button1_Click(object sender, EventArgs e)
        {
            SETFASTREPORT();
        }
        private void button2_Click(object sender, EventArgs e)
        {
            SEARCHINVSTAYOVER();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            UPDATEINVSTAYOVER();

            SEARCHINVSTAYOVER();
            MessageBox.Show("完成");
        }

        #endregion

       
    }
}

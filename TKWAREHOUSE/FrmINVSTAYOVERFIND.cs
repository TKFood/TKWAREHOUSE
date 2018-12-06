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

            report1.Dictionary.Connections[0].ConnectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;

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
            FASTSQL.AppendFormat(@" WHERE [CHECKDATE]='{0}'",dateTimePicker1.Value.ToString("yyyy/MM/dd"));
            FASTSQL.AppendFormat(@" ORDER BY [KIND],[LOTNO],[MB001]");
            FASTSQL.AppendFormat(@" ");
            
            

            return FASTSQL.ToString();
        }

        public void SEARCHINVSTAYOVER()
        {
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

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
        }
        #endregion


    }
}

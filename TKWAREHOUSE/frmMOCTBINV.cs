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
    public partial class frmMOCTBINV : Form
    {
        SqlConnection sqlConn = new SqlConnection();
        SqlCommand sqlComm = new SqlCommand();
        string connectionString;
        StringBuilder sbSql = new StringBuilder();
        StringBuilder sbSqlQuery = new StringBuilder();
        SqlDataAdapter adapter = new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilder = new SqlCommandBuilder();
        SqlTransaction tran;
        SqlCommand cmd = new SqlCommand();
        DataSet ds = new DataSet();
        DataSet ds2 = new DataSet();
        string tablename = null;
        int rownum = 0;
        DataGridViewRow row;
        string SALSESID = null;
        int result;

        public Report report1 { get; private set; }
        public Report report2 { get; private set; }

        public frmMOCTBINV()
        {
            InitializeComponent();
        }

        #region BUTTON
        public void SETFASTREPORT()
        {

            string SQL;
            string SQL2;

            report1 = new Report();
            report1.Load(@"REPORT\查製令用量.frx");

            report1.Dictionary.Connections[0].ConnectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;

            TableDataSource Table = report1.GetDataSource("Table") as TableDataSource;
            SQL = SETFASETSQL();
            Table.SelectCommand = SQL;
            report1.Preview = previewControl1;
            report1.Show();

  
            report2 = new Report();
            report2.Load(@"REPORT\查製令用量明細.frx");

            report2.Dictionary.Connections[0].ConnectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;

            TableDataSource Table2 = report2.GetDataSource("Table") as TableDataSource;
            SQL2 = SETFASETSQL2();
            Table2.SelectCommand = SQL2;
            report2.Preview = previewControl2;
            report2.Show();

        }

        public string SETFASETSQL()
        {
            StringBuilder FASTSQL = new StringBuilder();

            FASTSQL.AppendFormat(@"  SELECT TA003 AS '生產日',TB003 AS '品號',MB002 AS '品名',SUM(TB004-TB005) AS '需求量',SUM(TB004) AS '領料量',SUM(TB005) AS '已領量'");
            FASTSQL.AppendFormat(@"  FROM [TK].dbo.MOCTB,[TK].dbo.MOCTA,[TK].dbo.INVMB");
            FASTSQL.AppendFormat(@"  WHERE TA001=TB001 AND TA002=TB002");
            FASTSQL.AppendFormat(@"  AND TB003=MB001");
            FASTSQL.AppendFormat(@"  AND (TB004-TB005)>0");
            FASTSQL.AppendFormat(@"  AND TA003>='{0}' AND TA003<='{1}'",dateTimePicker1.Value.ToString("yyyyMMdd"), dateTimePicker2.Value.ToString("yyyyMMdd"));
            FASTSQL.AppendFormat(@"  AND (TB003 LIKE '1%' OR TB003 LIKE '2%' )");
            FASTSQL.AppendFormat(@"  GROUP BY TA003,TB003,MB002");
            FASTSQL.AppendFormat(@"  ORDER BY TA003,TB003,MB002");
            FASTSQL.AppendFormat(@"  ");

            return FASTSQL.ToString();
        }

        public string SETFASETSQL2()
        {
            StringBuilder FASTSQL = new StringBuilder();

            FASTSQL.AppendFormat(@"  SELECT TA003 AS '生產日',TB001 AS '製令',TB002 AS '製令號',TB003 AS '品號',MB002 AS '品名',(TB004-TB005)  AS '需求量',TB004 AS '領料量',TB005 AS '已領量',TA026 AS '訂單',TA027 AS '訂單號',TA028 AS '訂單序',TA006 AS '製品號',TA034 AS '製品'");
            FASTSQL.AppendFormat(@"  FROM [TK].dbo.MOCTB,[TK].dbo.MOCTA,[TK].dbo.INVMB");
            FASTSQL.AppendFormat(@"  WHERE TA001=TB001 AND TA002=TB002");
            FASTSQL.AppendFormat(@"  AND TB003=MB001");
            FASTSQL.AppendFormat(@"  AND (TB004-TB005)>0");
            FASTSQL.AppendFormat(@"  AND TA003>='{0}' AND TA003<='{1}'", dateTimePicker1.Value.ToString("yyyyMMdd"), dateTimePicker2.Value.ToString("yyyyMMdd"));
            FASTSQL.AppendFormat(@"  ORDER BY TA003,TB003,TB001,TB002");
            FASTSQL.AppendFormat(@"  ");

            return FASTSQL.ToString();
        }

        #endregion

        #region BUTTON
        private void button1_Click(object sender, EventArgs e)
        {
            SETFASTREPORT();
        }

        #endregion
    }
}

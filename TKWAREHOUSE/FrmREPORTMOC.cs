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
    public partial class FrmREPORTMOC : Form
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

        public FrmREPORTMOC()
        {
            InitializeComponent();

            combobox2load();
        }

        #region FUNCTION
        public void combobox2load()
        {

            connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
            sqlConn = new SqlConnection(connectionString);
            String Sequel = "SELECT MD002,MD001 FROM [TK].dbo.CMSMD WHERE MD002 LIKE '新%' ORDER BY MD001";
            SqlDataAdapter da = new SqlDataAdapter(Sequel, sqlConn);
            DataTable dt = new DataTable();
            sqlConn.Open();

            dt.Columns.Add("MD001", typeof(string));
            dt.Columns.Add("MD002", typeof(string));
            da.Fill(dt);
            comboBox2.DataSource = dt.DefaultView;
            comboBox2.ValueMember = "MD001";
            comboBox2.DisplayMember = "MD002";
            sqlConn.Close();



        }

        public void SETFASTREPORT()
        {

            string SQL;
            report1 = new Report();
            report1.Load(@"REPORT\製令領用量.frx");

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

            FASTSQL.AppendFormat(@"  SELECT MD002 AS '線別',TA001 AS '製令單別',TA002 AS '製令單號',TA003 AS '開單日期',TA006 AS '產品品號',TA034 AS '產品品名',TA015 AS '預計產量',TA007 AS '單位1',TB003 AS '材料品號',TB012 AS '材料品名',TB007 AS '單位2',TB004 AS '需領用量',MC004 AS '標準批量',ROUND(TA015/MC004,2) AS '總桶數'");
            FASTSQL.AppendFormat(@"  ,CASE WHEN FLOOR(ROUND(TA015/MC004,2)) = ROUND(TA015/MC004,2) THEN ROUND(TA015/MC004,2) ELSE CASE WHEN (ROUND(TA015/MC004,2)-1)>0 THEN FLOOR(ROUND(TA015/MC004,2)) ELSE 0 END  END AS '整桶數'");
            FASTSQL.AppendFormat(@"  ,CASE WHEN FLOOR(ROUND(TA015/MC004,2)) != ROUND(TA015/MC004,2) THEN 1 ELSE 0  END AS '最後桶數'");
            FASTSQL.AppendFormat(@"  ,CASE WHEN FLOOR(ROUND(TA015/MC004,2)) = ROUND(TA015/MC004,2) THEN ROUND(TB004/ROUND(TA015/MC004,2),2) ELSE CASE WHEN (ROUND(TA015/MC004,2)-1)>0 THEN ROUND(TB004/ROUND(TA015/MC004,2),2) ELSE 0 END  END AS '整桶用量'");
            FASTSQL.AppendFormat(@"  ,TB004-(ROUND(TB004/ROUND(TA015/MC004,2),2)*(CASE WHEN FLOOR(ROUND(TA015/MC004,2)) = ROUND(TA015/MC004,2) THEN ROUND(TA015/MC004,2) ELSE CASE WHEN (ROUND(TA015/MC004,2)-1)>0 THEN FLOOR(ROUND(TA015/MC004,2)) ELSE 0 END  END)) AS '最後桶用量'");
            FASTSQL.AppendFormat(@"  ,ROUND(TB004/ROUND(TA015/MC004,2),2)  AS '標準用量'");
            FASTSQL.AppendFormat(@"  FROM [TK].dbo.MOCTA,[TK].dbo.MOCTB,[TK].dbo.BOMMC,[TK].dbo.CMSMD");
            FASTSQL.AppendFormat(@"  WHERE TA001=TB001 AND TA002=TB002");
            FASTSQL.AppendFormat(@"  AND TA006=MC001");
            FASTSQL.AppendFormat(@"  AND TA021=MD001");
            FASTSQL.AppendFormat(@"  AND TB003 LIKE '1%'");
            FASTSQL.AppendFormat(@"  AND TB003 NOT IN ('101001001','101001009')");
            FASTSQL.AppendFormat(@"  AND MD002 ='{0}'",comboBox2.Text.ToString());
            FASTSQL.AppendFormat(@"  AND TA003>='{0}' AND TA003<='{1}'",dateTimePicker1.Value.ToString("yyyyMMdd"), dateTimePicker2.Value.ToString("yyyyMMdd"));
            FASTSQL.AppendFormat(@"  ORDER BY TA021,TA001,TA002,TB003");
            FASTSQL.AppendFormat(@"  ");
            FASTSQL.AppendFormat(@"  ");

            return FASTSQL.ToString();
        }

        public void SETFASTREPORT2()
        {

            string SQL;
            report2 = new Report();
            report2.Load(@"REPORT\製令領用量(特).frx");

            report2.Dictionary.Connections[0].ConnectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;

            TableDataSource Table = report2.GetDataSource("Table") as TableDataSource;
            SQL = SETFASETSQL2();
            Table.SelectCommand = SQL;
            report2.Preview = previewControl2;
            report2.Show();

        }

        public string SETFASETSQL2()
        {
            StringBuilder FASTSQL = new StringBuilder();
            StringBuilder STRQUERY = new StringBuilder();

            FASTSQL.AppendFormat(@"  SELECT MD002 AS '線別',TA003 AS '開單日期',TB003 AS '材料品號',TB012 AS '材料品名',TB007 AS '單位2',SUM(TB004) AS '需領用量',ROUND(SUM(TB004)/22,0) AS '包數'");
            FASTSQL.AppendFormat(@"  FROM [TK].dbo.MOCTA,[TK].dbo.MOCTB,[TK].dbo.BOMMC,[TK].dbo.CMSMD");
            FASTSQL.AppendFormat(@"  WHERE TA001=TB001 AND TA002=TB002");
            FASTSQL.AppendFormat(@"  AND TA006=MC001");
            FASTSQL.AppendFormat(@"  AND TA021=MD001");
            FASTSQL.AppendFormat(@"  AND TB003 LIKE '1%'");
            FASTSQL.AppendFormat(@"  AND TB003 IN ('101001001','101001009')");
            FASTSQL.AppendFormat(@"  AND MD002 ='{0}'", comboBox2.Text.ToString());
            FASTSQL.AppendFormat(@"  AND TA003>='{0}' AND TA003<='{1}'", dateTimePicker1.Value.ToString("yyyyMMdd"), dateTimePicker2.Value.ToString("yyyyMMdd"));
            FASTSQL.AppendFormat(@"  GROUP BY MD002,TA003,TB003,TB012,TB007");
            FASTSQL.AppendFormat(@"  ORDER BY MD002,TA003,TB003,TB012,TB007 ");
            FASTSQL.AppendFormat(@"  ");
            FASTSQL.AppendFormat(@"  ");

            return FASTSQL.ToString();
        }

        #endregion

        #region BUTTON
        private void button1_Click(object sender, EventArgs e)
        {
            SETFASTREPORT();
            SETFASTREPORT2();
        }

        #endregion
    }
}

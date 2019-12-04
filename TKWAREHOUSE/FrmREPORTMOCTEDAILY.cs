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
    public partial class FrmREPORTMOCTEDAILY : Form
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

        public FrmREPORTMOCTEDAILY()
        {
            InitializeComponent();
        }

        #region FUNCTION
        public void SETFASTREPORT()
        {

            string SQL;
            report1 = new Report();
            report1.Load(@"REPORT\每日領退量表.frx");

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

            FASTSQL.AppendFormat(@"  SELECT TC003 AS '日期',TE004 AS '品號',TE017 AS '品名',SUM(MQ010*TE005)*-1 AS '總領退量',TE006 AS '單位'");
            FASTSQL.AppendFormat(@"  ,(SELECT ISNULL(SUM(LA005*LA011),0) FROM [TK].dbo.INVLA WHERE LA001=TE004 AND LA009 IN ('20004','20006') AND LA004<='{0}') AS '庫存量'", dateTimePicker3.Value.ToString("yyyyMMdd"));
            FASTSQL.AppendFormat(@"  FROM  [TK].dbo.CMSMQ,[TK].dbo.MOCTC,[TK].dbo.MOCTE");
            FASTSQL.AppendFormat(@"  WHERE TC001=TE001 AND TC002=TE002 AND MQ001=TC001");
            FASTSQL.AppendFormat(@"  AND (TE004 LIKE '1%' OR TE004 LIKE '2%' )");
            FASTSQL.AppendFormat(@"  AND TC009='Y'");
            FASTSQL.AppendFormat(@"  AND TC003>='{0}' AND TC003<='{1}'",dateTimePicker1.Value.ToString("yyyyMMdd"), dateTimePicker2.Value.ToString("yyyyMMdd"));
            FASTSQL.AppendFormat(@"  GROUP BY TC003,TE004,TE017,TE006");
            FASTSQL.AppendFormat(@"  ORDER BY TC003,TE004,TE017,TE006");
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

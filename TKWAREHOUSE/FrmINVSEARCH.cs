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
    public partial class FrmINVSEARCH : Form
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

        public FrmINVSEARCH()
        {
            InitializeComponent();
        }

        #region FUNCTION

        public void SETFASTREPORT()
        {

            string SQL;

            report1 = new Report();
            report1.Load(@"REPORT\查品號批號的單據.frx");

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

            FASTSQL.AppendFormat(@"  DECLARE @PRAM1 NVARCHAR(20),@PRAM2 NVARCHAR(20)");
            FASTSQL.AppendFormat(@"  SET @PRAM1='{0}' ",textBox1.Text);
            FASTSQL.AppendFormat(@"  SET @PRAM2='{0}'", textBox2.Text);
            FASTSQL.AppendFormat(@"  ");
            FASTSQL.AppendFormat(@"  SELECT TH004 AS '品號',MB002 AS '品名',TH017 AS '批號',TH008 AS '數量',KIND AS '類別',TH001 AS '單別',TH002 AS '單號',TH003 AS '序號'");
            FASTSQL.AppendFormat(@"  FROM (");
            FASTSQL.AppendFormat(@"  SELECT '銷貨' AS 'KIND',TH004,TH017,TH008,TH001,TH002,TH003 FROM [TK].dbo.COPTH WHERE TH004=@PRAM1 AND TH017=@PRAM2");
            FASTSQL.AppendFormat(@"  UNION");
            FASTSQL.AppendFormat(@"  SELECT '銷退' AS 'KIND',TJ004,TJ017,TJ007, TJ001,TJ002,TJ003 FROM [TK].dbo.COPTJ WHERE TJ004=@PRAM1 AND TJ014=@PRAM2");
            FASTSQL.AppendFormat(@"  UNION");
            FASTSQL.AppendFormat(@"  SELECT '進貨' AS 'KIND',TH004,TH010,TH007, TH001,TH002,TH003 FROM [TK].dbo.PURTH WHERE TH004=@PRAM1 AND TH010=@PRAM2");
            FASTSQL.AppendFormat(@"  UNION");
            FASTSQL.AppendFormat(@"  SELECT '退貨' AS 'KIND',TJ004,TJ012, TJ009,TJ001,TJ002,TJ003  FROM [TK].dbo.PURTJ WHERE TJ004=@PRAM1 AND TJ012=@PRAM2 ");
            FASTSQL.AppendFormat(@"  UNION");
            FASTSQL.AppendFormat(@"  SELECT '轉撥/異動' AS 'KIND',TB004,TB014, TB007, TB001,TB002,TB003 FROM [TK].dbo.INVTB WHERE TB004=@PRAM1 AND TB014=@PRAM2");
            FASTSQL.AppendFormat(@"  UNION");
            FASTSQL.AppendFormat(@"  SELECT '借出' AS 'KIND',TG004,TG017,TG009, TG001,TG002,TG003 FROM [TK].dbo.INVTG WHERE TG004=@PRAM1 AND TG017=@PRAM2");
            FASTSQL.AppendFormat(@"  UNION");
            FASTSQL.AppendFormat(@"  SELECT '歸還' AS 'KIND',TI004,TI017,TI009, TI001,TI002,TI003 FROM [TK].dbo.INVTI WHERE TI004=@PRAM1 AND TI017=@PRAM2 ");
            FASTSQL.AppendFormat(@"  ) AS TEMP   ");
            FASTSQL.AppendFormat(@"  LEFT JOIN [TK].dbo.INVMB ON MB001=TH004");
            FASTSQL.AppendFormat(@"  ORDER BY TH001,TH002,TH003");
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

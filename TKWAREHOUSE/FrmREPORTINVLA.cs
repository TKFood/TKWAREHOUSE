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
    public partial class FrmREPORTINVLA : Form
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

        DataTable dt = new DataTable();

        public Report report1 { get; private set; }
        public Report report2 { get; private set; }

        public FrmREPORTINVLA()
        {
            InitializeComponent();
        }

        #region FUNCTION
        public void SETFASTREPORT()
        {

            string SQL;
            report1 = new Report();
            report1.Load(@"REPORT\成品倉撿料表.frx");

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
            
            if(!string.IsNullOrEmpty(textBox1.Text))
            {
                STRQUERY.AppendFormat(@" AND LA001 LIKE '{0}%'",textBox1.Text.Trim());
            }

            FASTSQL.AppendFormat(@"  
                                    SELECT SERNO,KINDS AS '分類',LA004 AS '日期',LA001 AS '品號',LA009 AS '庫別', LA011 AS '數量',MB002 AS '品名',MB003 AS '規格',MB004 AS '單位'
                                    FROM (
                                    SELECT '1' AS SERNO,'銷貨單' AS 'KINDS',LA004,LA001,LA009,SUM(LA011) LA011,MB002,MB003,MB004
                                    FROM [TK].dbo.INVLA WITH(NOLOCK),[TK].dbo.CMSMQ WITH(NOLOCK),[TK].dbo.INVMB WITH(NOLOCK)
                                    WHERE LA006=MQ001
                                    AND LA001=MB001
                                    AND MQ003 IN ('23')
                                    AND LA005='-1'
                                    AND LA009='20001'
                                    GROUP BY LA004,LA001,LA009,MB002,MB003,MB004
                                    UNION ALL
                                    SELECT '2' AS SERNO,'暫出單' AS 'KINDS',LA004,LA001,LA009,SUM(LA011) LA011,MB002,MB003,MB004
                                    FROM [TK].dbo.INVLA WITH(NOLOCK),[TK].dbo.CMSMQ WITH(NOLOCK),[TK].dbo.INVMB WITH(NOLOCK)
                                    WHERE LA006=MQ001
                                    AND LA001=MB001
                                    AND MQ003 IN ('13','14')
                                    AND LA005='-1'
                                    AND LA009='20001'
                                    GROUP BY LA004,LA001,LA009,MB002,MB003,MB004
                                    UNION ALL
                                    SELECT '3' AS SERNO,'暫入單' AS 'KINDS',LA004,LA001,LA009,SUM(LA011) LA011,MB002,MB003,MB004
                                    FROM [TK].dbo.INVLA WITH(NOLOCK),[TK].dbo.CMSMQ WITH(NOLOCK),[TK].dbo.INVMB WITH(NOLOCK)
                                    WHERE LA006=MQ001
                                    AND LA001=MB001
                                    AND MQ003 IN ('15','16')
                                    AND LA005='-1'
                                    AND LA009='20001'
                                    GROUP BY LA004,LA001,LA009,MB002,MB003,MB004
                                    UNION ALL
                                    SELECT '4' AS SERNO,'庫存異動單' AS 'KINDS',LA004,LA001,LA009,SUM(LA011) LA011,MB002,MB003,MB004
                                    FROM [TK].dbo.INVLA WITH(NOLOCK),[TK].dbo.CMSMQ WITH(NOLOCK),[TK].dbo.INVMB WITH(NOLOCK)
                                    WHERE LA006=MQ001
                                    AND LA001=MB001
                                    AND MQ003 IN ('11')
                                    AND LA005='-1'
                                    AND LA009='20001'
                                    GROUP BY LA004,LA001,LA009,MB002,MB003,MB004
                                    UNION ALL
                                    SELECT '5' AS SERNO,'轉撥單' AS 'KINDS',LA004,LA001,LA009,SUM(LA011) LA011,MB002,MB003,MB004
                                    FROM [TK].dbo.INVLA WITH(NOLOCK),[TK].dbo.CMSMQ WITH(NOLOCK),[TK].dbo.INVMB WITH(NOLOCK)
                                    WHERE LA006=MQ001
                                    AND LA001=MB001
                                    AND MQ003 IN ('12','13')
                                    AND LA005='-1'
                                    AND LA009='20001'
                                    GROUP BY LA004,LA001,LA009,MB002,MB003,MB004
                                    ) AS TEMP
                                    WHERE LA004='{0}'
                                    {1}
                                    ORDER BY LA004,LA001,SERNO,KINDS

                                    ", dateTimePicker1.Value.ToString("yyyyMMdd"), STRQUERY.ToString());

            return FASTSQL.ToString();
        }

        #endregion

        #region BUTTON

        private void button2_Click(object sender, EventArgs e)
        {
            SETFASTREPORT();
        }
        #endregion

    }
}

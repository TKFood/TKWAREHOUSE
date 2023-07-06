using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
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
    public partial class FrmREPORTSTOCK : Form
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

        public FrmREPORTSTOCK()
        {
            InitializeComponent();
        }

        #region FUNCTION
        public void SETFASTREPORT()
        {

            string SQL; 
            report1 = new Report();
            report1.Load(@"REPORT\庫存及預計出貨表.frx");

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

        

            FASTSQL.AppendFormat(@" 
                                    SELECT 
                                    LA001 AS '品號'
                                    ,LA016 AS '批號'
                                    ,LA009 AS '庫別'
                                    ,NUMS AS '庫存數量'
                                    ,MB002 AS '品名'
                                    ,MB003 AS '規格'
                                    ,MB004 AS '單位'
                                    ,TC0012A
                                    ,TC0012B
                                    ,TF003 AS '入庫日'
                                    ,TG014TG015 AS '製令'
                                    ,(CASE WHEN ISNULL(TC0012A,'')<>'' THEN TC0012A ELSE TC0012B END ) AS '訂單'
                                    ,TC053 AS '客戶'
                                    ,TC006 AS '業務'
                                    ,TD013 AS '預交日'
                                    ,MV002 AS '業務員'
                                    ,DATEDIFF(day, TF003, GETDATE())  AS '存放天數'
                                    FROM 
                                    (
	                                    SELECT *
	                                    ,(SELECT TOP 1 TA026+TA027 FROM [TK].dbo.MOCTA WHERE  TA001+TA002=TG014TG015) AS TC0012A
	                                    ,(SELECT TOP 1 COPTD001+COPTD002 
	                                    FROM [TKMOC].[dbo].[MOCMANULINEMERGE],[TKMOC].dbo.[MOCMANULINE],[TK].dbo.MOCTA
	                                    WHERE [MOCMANULINEMERGE].SID=[MOCMANULINE].ID
	                                    AND TA033=[MOCMANULINEMERGE].[NO]
	                                    AND TA001+TA002=TG014TG015
	                                    ORDER BY TA015 DESC
	                                    ) AS TC0012B
	                                    FROM  
	                                    (
		                                    SELECT LA001,LA016,LA009,NUMS,MB002,MB003,MB004
		                                    ,(SELECT TOP 1 TF003 FROM [TK].dbo.MOCTG,[TK].dbo.MOCTF WHERE TG001=TF001 AND TG002=TF002 AND TG004=LA001 AND TG017=LA016 AND TG010=LA009 ) TF003
		                                    ,(SELECT TOP 1 TG014+TG015 FROM [TK].dbo.MOCTG,[TK].dbo.MOCTF WHERE TG001=TF001 AND TG002=TF002 AND TG004=LA001 AND TG017=LA016 AND TG010=LA009 ) TG014TG015
		                                    FROM 
		                                    (
		                                    SELECT LA001,LA016,LA009,SUM(LA005*LA011) AS  NUMS
		                                    FROM [TK].dbo.INVLA
		                                    WHERE LA009='20001'
		                                    GROUP BY LA001,LA016,LA009
		                                    HAVING  SUM(LA005*LA011)>0
		                                    ) AS TEMP
	                                    LEFT JOIN [TK].dbo.INVMB ON MB001=LA001
	                                    ) AS TEMP2
                                    ) AS TMEP3
                                    LEFT JOIN [TK].dbo.COPTC ON TC001+TC002=(CASE WHEN ISNULL(TC0012A,'')<>'' THEN TC0012A ELSE TC0012B END )
                                    LEFT JOIN [TK].dbo.COPTD ON TD001=TC001 AND TD002=TC002 AND TD004=LA001 
                                    LEFT JOIN [TK].dbo.CMSMV ON MV001=TC006
                                    ORDER BY LA001,LA016



                                                                        ");

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

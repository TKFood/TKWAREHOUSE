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

            //connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
            //sqlConn = new SqlConnection(connectionString);
            //String Sequel = "SELECT MD002,MD001 FROM [TK].dbo.CMSMD WHERE MD002 LIKE '新%' ORDER BY MD001";
            //SqlDataAdapter da = new SqlDataAdapter(Sequel, sqlConn);
            //DataTable dt = new DataTable();
            //sqlConn.Open();

            //dt.Columns.Add("MD001", typeof(string));
            //dt.Columns.Add("MD002", typeof(string));
            //da.Fill(dt);
            //comboBox2.DataSource = dt.DefaultView;
            //comboBox2.ValueMember = "MD001";
            //comboBox2.DisplayMember = "MD002";
            //sqlConn.Close();



        }

        public void SETFASTREPORT()
        { 
            string SQL;
            string SQL1;
            report1 = new Report();
            report1.Load(@"REPORT\製令領用量.frx");

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
            TableDataSource Table1 = report1.GetDataSource("Table1") as TableDataSource;
            SQL = SETFASETSQL();
            SQL1 = SETFASETSQL2();
            Table.SelectCommand = SQL;
            Table1.SelectCommand = SQL1;
            report1.Preview = previewControl1;
            report1.Show();

        }

        public string SETFASETSQL()
        {
            StringBuilder FASTSQL = new StringBuilder();
            StringBuilder STRQUERY = new StringBuilder();
            StringBuilder STRQUERYNOTIN = new StringBuilder();

            //限定品號
            DataTable DT = FIND_FrmREPORTMOC();
            if(DT!=null && DT.Rows.Count>=1)
            {
                // 組合 SQL 條件字串             
                STRQUERY.Append("AND (");

                for (int i = 0; i < DT.Rows.Count; i++)
                {
                    string TB003 = DT.Rows[i]["TB003"].ToString();
                    STRQUERY.AppendFormat("TB003 LIKE '{0}%'", TB003);

                    if (i < DT.Rows.Count - 1)
                    {
                        STRQUERY.Append(" OR ");
                    }
                }

                STRQUERY.Append(")");
            }
            //NOT IN 品號
            DataTable DTNOTIN = FIND_FrmREPORTMOCNOTIN();
            if (DTNOTIN != null && DTNOTIN.Rows.Count >= 1)
            {
                STRQUERYNOTIN.Append("AND TB003 NOT IN (");

                for (int i = 0; i < DTNOTIN.Rows.Count; i++)
                {
                    string TB003 = DTNOTIN.Rows[i]["TB003"].ToString();
                    STRQUERYNOTIN.AppendFormat("'{0}'", TB003);

                    if (i < DTNOTIN.Rows.Count - 1)
                    {
                        STRQUERYNOTIN.Append(", ");
                    }
                }

                STRQUERYNOTIN.Append(")");
            }

            FASTSQL.AppendFormat(@"    
                                SELECT 線別,製令單別,製令單號,開單日期,產品品號,產品品名,預計產量,單位1,材料品號,材料品名,單位2,需領用量,標準批量,總桶數,整桶數,最後桶數,整桶用量,最後桶用量,標準用量
                                ,ISNULL(BOMMD.MD004,'') AS '材料單位',ISNULL(BOMMD.MD006,0) AS '組成用量',ISNULL(BOMMD.MD007,0) AS '底數',ISNULL(BOMMD.MD008,0) AS '損耗率%'
                                ,CASE WHEN ISNULL(BOMMD.MD006,0)>0 THEN ISNULL(BOMMD.MD006,0)/ISNULL(BOMMD.MD007,0)*(1+ISNULL(BOMMD.MD008,0)) ELSE 0 END AS 'BOM用量'
                                FROM (
                                SELECT CMSMD.MD002 AS '線別',TA001 AS '製令單別',TA002 AS '製令單號',TA003 AS '開單日期',TA006 AS '產品品號',TA034 AS '產品品名',TA015 AS '預計產量',TA007 AS '單位1',TB003 AS '材料品號',TB012 AS '材料品名',TB007 AS '單位2',TB004 AS '需領用量',MC004 AS '標準批量',ROUND(TA015/MC004,3) AS '總桶數'
                                ,(CASE WHEN TA015>0 AND MC004>0 THEN  FLOOR(ROUND(TA015/MC004,3)) ELSE 0 END) AS '整桶數'
                                ,(CASE WHEN TA015>0 AND MC004>0 THEN (ROUND(TA015/MC004,3)-FLOOR(ROUND(TA015/MC004,3)))   ELSE 0 END) AS '最後桶數'
                                ,(CASE WHEN TA015>0 AND MC004>0 THEN (CASE WHEN FLOOR(ROUND(TA015/MC004,3)) = ROUND(TA015/MC004,3) THEN ROUND(TB004/ROUND(TA015/MC004,3),3) ELSE CASE WHEN (ROUND(TA015/MC004,3)-1)>0 THEN ROUND(TB004/ROUND(TA015/MC004,3),3) ELSE 0 END  END )  ELSE 0 END)  AS '整桶用量'
                                ,(CASE WHEN TA015>0 AND MC004>0 THEN (TB004-(ROUND(TB004/ROUND(TA015/MC004,3),3)*(CASE WHEN FLOOR(ROUND(TA015/MC004,3)) = ROUND(TA015/MC004,3) THEN ROUND(TA015/MC004,3) ELSE CASE WHEN (ROUND(TA015/MC004,3)-1)>0 THEN FLOOR(ROUND(TA015/MC004,3)) ELSE 0 END  END)))  ELSE 0 END)   AS '最後桶用量'
                                ,(CASE WHEN TA015>0 AND MC004>0 THEN ROUND(TB004/ROUND(TA015/MC004,3),3) ELSE 0 END )  AS '標準用量'
                                FROM [TK].dbo.MOCTA,[TK].dbo.MOCTB,[TK].dbo.BOMMC,[TK].dbo.CMSMD
                                WHERE TA001=TB001 AND TA002=TB002
                                AND TA006=MC001
                                AND TA021=CMSMD.MD001
                                {2}
                                {3}
                                AND CMSMD.MD002  IN ('製一線','製二線 ','手工線') 
                                    AND TA003>='{0}' AND TA003<='{1}'
                                ) AS TEMP
                                LEFT JOIN  [TK].dbo.BOMMD ON 產品品號=MD001 AND 材料品號=MD003
                                ORDER BY 線別,製令單別,製令單號   "
                            , dateTimePicker1.Value.ToString("yyyyMMdd"), dateTimePicker2.Value.ToString("yyyyMMdd"), STRQUERY.ToString(), STRQUERYNOTIN.ToString());
         

            return FASTSQL.ToString();
        }

        

        public string SETFASETSQL2()
        {
            StringBuilder FASTSQL = new StringBuilder();
            StringBuilder STRQUERY = new StringBuilder();
            StringBuilder STRQUERYIN = new StringBuilder();

            //限定品號
            DataTable DT = FIND_FrmREPORTMOC();
            if (DT != null && DT.Rows.Count >= 1)
            {
                // 組合 SQL 條件字串             
                STRQUERY.Append("AND (");

                for (int i = 0; i < DT.Rows.Count; i++)
                {
                    string TB003 = DT.Rows[i]["TB003"].ToString();
                    STRQUERY.AppendFormat("TB003 LIKE '{0}%'", TB003);

                    if (i < DT.Rows.Count - 1)
                    {
                        STRQUERY.Append(" OR ");
                    }
                }

                STRQUERY.Append(")");
            }
            //NOT IN 品號
            DataTable DTIN = FIND_FrmREPORTMOCNOTIN();
            if (DTIN != null && DTIN.Rows.Count >= 1)
            {
                STRQUERYIN.Append("AND TB003  IN (");

                for (int i = 0; i < DTIN.Rows.Count; i++)
                {
                    string TB003 = DTIN.Rows[i]["TB003"].ToString();
                    STRQUERYIN.AppendFormat("'{0}'", TB003);

                    if (i < DTIN.Rows.Count - 1)
                    {
                        STRQUERYIN.Append(", ");
                    }
                }

                STRQUERYIN.Append(")");
            }

            FASTSQL.AppendFormat(@"  
                                 SELECT MD002 AS '線別',TA003 AS '開單日期',TB003 AS '材料品號',TB012 AS '材料品名',TB007 AS '單位2',SUM(TB004) AS '需領用量',CEILING(SUM(TB004)/22) AS '包數'
                                 FROM [TK].dbo.MOCTA,[TK].dbo.MOCTB,[TK].dbo.BOMMC,[TK].dbo.CMSMD
                                 WHERE TA001=TB001 AND TA002=TB002
                                 AND TA006=MC001
                                 AND TA021=MD001
                                 {2}
                                    {3}
                                 AND MD002  IN ('製一線','製二線 ','手工線') 
                                 AND TA003>='{0}' AND TA003<='{1}'
                                 GROUP BY MD002,TA003,TB003,TB012,TB007
                                 ORDER BY MD002,TA003,TB003,TB012,TB007  
                                ", dateTimePicker1.Value.ToString("yyyyMMdd"), dateTimePicker2.Value.ToString("yyyyMMdd"), STRQUERY.ToString(), STRQUERYIN.ToString());

            return FASTSQL.ToString();
        }


        public DataTable FIND_FrmREPORTMOC()
        {
            DataSet DS1 = new DataSet();
            SqlDataAdapter adapter1 = new SqlDataAdapter();
            SqlCommandBuilder sqlCmdBuilder1 = new SqlCommandBuilder();

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


                sbSql.AppendFormat(@" 
                                   SELECT [TB003]
                                    FROM [TKWAREHOUSE].[dbo].[FrmREPORTMOC]
                                    ");

                adapter1 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder1 = new SqlCommandBuilder(adapter1);
                sqlConn.Open();
                DS1.Clear();
                adapter1.Fill(DS1, "DS1");
                sqlConn.Close();


                if (DS1.Tables["DS1"].Rows.Count >= 1)
                {
                    return DS1.Tables["DS1"];
                }
                else
                {
                    return null;
                }
            }
            catch(Exception EX)
            {
                return null;
            }
            finally
            {

            }
        }

        public DataTable FIND_FrmREPORTMOCNOTIN()
        {
            DataSet DS1 = new DataSet();
            SqlDataAdapter adapter1 = new SqlDataAdapter();
            SqlCommandBuilder sqlCmdBuilder1 = new SqlCommandBuilder();

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


                sbSql.AppendFormat(@" 
                                   SELECT [TB003]
                                    FROM [TKWAREHOUSE].[dbo].[FrmREPORTMOCNOTIN]
                                    ");

                adapter1 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder1 = new SqlCommandBuilder(adapter1);
                sqlConn.Open();
                DS1.Clear();
                adapter1.Fill(DS1, "DS1");
                sqlConn.Close();


                if (DS1.Tables["DS1"].Rows.Count >= 1)
                {
                    return DS1.Tables["DS1"];
                }
                else
                {
                    return null;
                }
            }
            catch (Exception EX)
            {
                return null;
            }
            finally
            {

            }
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

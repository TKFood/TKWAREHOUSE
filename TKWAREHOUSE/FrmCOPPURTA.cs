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
using TKITDLL;

namespace TKWAREHOUSE
{
    public partial class FrmCOPPURTA : Form
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
        SqlDataAdapter adapter4 = new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilder4 = new SqlCommandBuilder();
        SqlDataAdapter adapter5 = new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilder5 = new SqlCommandBuilder();

        SqlTransaction tran;
        SqlCommand cmd = new SqlCommand();
        DataSet ds1 = new DataSet();
        DataSet ds2 = new DataSet();
        DataSet ds3 = new DataSet();
        DataSet ds4 = new DataSet();
        DataSet ds5 = new DataSet();

        DataTable dt = new DataTable();
        string tablename = null;
        int result;

        string ID;
        string NEWID;
        string DELTD001;
        string DELTD002;
        string DELTD003;

        string MOCTA001;
        string MOCTA002;
        string MOCTA003;

        string MAXID;

        string MB001;
        string MB002;
        string MB003;
        decimal BAR;
        string TA026A;
        string TA027A;
        string TA028A;
        decimal SUM1;
        string TC015TD020;

        string DELCOPPURBATCHPUR_ID;
        string DELCOPPURBATCHPUR_TA001;
        string DELCOPPURBATCHPUR_TA002;

        public class PURTA
        {
            public string COMPANY;
            public string CREATOR;
            public string USR_GROUP;
            public string CREATE_DATE;
            public string MODIFIER;
            public string MODI_DATE;
            public string FLAG;
            public string CREATE_TIME;
            public string MODI_TIME;
            public string TRANS_TYPE;
            public string TRANS_NAME;
            public string sync_date;
            public string sync_time;
            public string sync_mark;
            public string sync_count;
            public string DataUser;
            public string DataGroup;
            public string TA001;
            public string TA002;
            public string TA003;
            public string TA004;
            public string TA005;
            public string TA006;
            public string TA007;
            public string TA008;
            public string TA009;
            public string TA010;
            public string TA011;
            public string TA012;
            public string TA013;
            public string TA014;
            public string TA015;
            public string TA016;
            public string TA017;
            public string TA018;
            public string TA019;
            public string TA020;
            public string TA021;
            public string TA022;
            public string TA023;
            public string TA024;
            public string TA025;
            public string TA026;
            public string TA027;
            public string TA028;
            public string TA029;
            public string TA030;
            public string TA031;
            public string TA032;
            public string TA033;
            public string TA034;
            public string TA035;
            public string TA036;
            public string TA037;
            public string TA038;
            public string TA039;
            public string TA040;
            public string TA041;
            public string TA042;
            public string TA043;
            public string TA044;
            public string TA045;
            public string TA046;
            public string UDF01;
            public string UDF02;
            public string UDF03;
            public string UDF04;
            public string UDF05;
            public string UDF06;
            public string UDF07;
            public string UDF08;
            public string UDF09;
            public string UDF10;
        }

        public class PURTB
        {
            public string COMPANY;
            public string CREATOR;
            public string USR_GROUP;
            public string CREATE_DATE;
            public string MODIFIER;
            public string MODI_DATE;
            public string FLAG;
            public string CREATE_TIME;
            public string MODI_TIME;
            public string TRANS_TYPE;
            public string TRANS_NAME;
            public string sync_date;
            public string sync_time;
            public string sync_mark;
            public string sync_count;
            public string DataUser;
            public string DataGroup;
            public string TB001;
            public string TB002;
            public string TB003;
            public string TB004;
            public string TB005;
            public string TB006;
            public string TB007;
            public string TB008;
            public string TB009;
            public string TB010;
            public string TB011;
            public string TB012;
            public string TB013;
            public string TB014;
            public string TB015;
            public string TB016;
            public string TB017;
            public string TB018;
            public string TB019;
            public string TB020;
            public string TB021;
            public string TB022;
            public string TB023;
            public string TB024;
            public string TB025;
            public string TB026;
            public string TB027;
            public string TB028;
            public string TB029;
            public string TB030;
            public string TB031;
            public string TB032;
            public string TB033;
            public string TB034;
            public string TB035;
            public string TB036;
            public string TB037;
            public string TB038;
            public string TB039;
            public string TB040;
            public string TB041;
            public string TB042;
            public string TB043;
            public string TB044;
            public string TB045;
            public string TB046;
            public string TB047;
            public string TB048;
            public string TB049;
            public string TB050;
            public string TB051;
            public string TB052;
            public string TB053;
            public string TB054;
            public string TB055;
            public string TB056;
            public string TB057;
            public string TB058;
            public string TB059;
            public string TB060;
            public string TB061;
            public string TB062;
            public string TB063;
            public string TB064;
            public string TB065;
            public string TB066;
            public string TB067;
            public string TB068;
            public string TB069;
            public string TB070;
            public string TB071;
            public string TB072;
            public string TB073;
            public string TB074;
            public string TB075;
            public string TB076;
            public string TB077;
            public string TB078;
            public string TB079;
            public string TB080;
            public string TB081;
            public string TB082;
            public string TB083;
            public string TB084;
            public string TB085;
            public string TB086;
            public string TB087;
            public string TB088;
            public string TB089;
            public string TB090;
            public string TB091;
            public string TB092;
            public string TB093;
            public string TB094;
            public string TB095;
            public string TB096;
            public string TB097;
            public string TB098;
            public string TB099;
            public string UDF01;
            public string UDF02;
            public string UDF03;
            public string UDF04;
            public string UDF05;
            public string UDF06;
            public string UDF07;
            public string UDF08;
            public string UDF09;
            public string UDF10;
        }

        public class MOCTADATA
        {
            public string COMPANY;
            public string CREATOR;
            public string USR_GROUP;
            public string CREATE_DATE;
            public string MODIFIER;
            public string MODI_DATE;
            public string FLAG;
            public string CREATE_TIME;
            public string MODI_TIME;
            public string TRANS_TYPE;
            public string TRANS_NAME;
            public string sync_count;
            public string DataGroup;
            public string TA001;
            public string TA002;
            public string TA003;
            public string TA004;
            public string TA005;
            public string TA006;
            public string TA007;
            public string TA009;
            public string TA010;
            public string TA011;
            public string TA012;
            public string TA013;
            public string TA014;
            public string TA015;
            public string TA016;
            public string TA017;
            public string TA018;
            public string TA019;
            public string TA020;
            public string TA021;
            public string TA022;
            public string TA023;
            public string TA024;
            public string TA025;
            public string TA026;
            public string TA027;
            public string TA028;
            public string TA029;
            public string TA030;
            public string TA031;
            public string TA032;
            public string TA033;
            public string TA034;
            public string TA035;
            public string TA040;
            public string TA041;
            public string TA042;
            public string TA043;
            public string TA044;
            public string TA045;
            public string TA046;
            public string TA047;
            public string TA049;
            public string TA050;
            public string TA200;
        }

        public FrmCOPPURTA()
        {
            InitializeComponent();
        }

        #region FUNCTION

        public void SEARCHBTACHID()
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

                sbSql.AppendFormat(@"  SELECT [ID] AS '批號',CONVERT(NVARCHAR,[BACTHDATES],112) AS '日期'");
                sbSql.AppendFormat(@"  FROM [TKWAREHOUSE].[dbo].[COPPURBATCHID]");
                sbSql.AppendFormat(@"  WHERE CONVERT(NVARCHAR,[BACTHDATES],112)='{0}'", dateTimePicker1.Value.ToString("yyyyMMdd"));
                sbSql.AppendFormat(@"  ORDER BY [ID] ");
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

        public string GETMAXID()
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


                StringBuilder sbSql = new StringBuilder();
                sbSql.Clear();
                sbSqlQuery.Clear();
                ds2.Clear();

                sbSql.AppendFormat(@"  SELECT ISNULL(MAX(ID),'00000000000') AS ID");
                sbSql.AppendFormat(@"  FROM [TKWAREHOUSE].[dbo].[COPPURBATCHID]");
                sbSql.AppendFormat(@"  WHERE CONVERT(NVARCHAR,[BACTHDATES],112)='{0}'", dateTimePicker1.Value.ToString("yyyyMMdd"));
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
            if (!string.IsNullOrEmpty(ID))
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

                    sbSql.AppendFormat(" INSERT INTO [TKWAREHOUSE].[dbo].[COPPURBATCHID]");
                    sbSql.AppendFormat(" ([ID],[BACTHDATES])");
                    sbSql.AppendFormat(" VALUES ('{0}','{1}')", ID, dateTimePicker1.Value.ToString("yyyyMMdd"));
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

                    SEARCHCOPPURBATCHCOPTD(ID);
                    SEARCHCOPPURBATCHUSED(ID);
                    SEARCHCOPPURBATCHPUR(ID);

                }
                else
                {
                    textBoxID.Text = null;
                    ID = null;

                }
            }
        }

        public void SEARCHCOPPURBATCHCOPTD(string ID)
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

          
                sbSql.AppendFormat(@"  
                                    SELECT [ID] AS '批號',[COPPURBATCHCOPTD].[TD001] AS '訂單單別',[COPPURBATCHCOPTD].[TD002] AS '訂單單號',[COPPURBATCHCOPTD].[TD003] AS '訂單序號',[COPPURBATCHCOPTD].[TD004] AS '品號',[COPPURBATCHCOPTD].[TD005] AS '品名',[COPPURBATCHCOPTD].[TD008] AS '訂單數量',[COPPURBATCHCOPTD].[TD009] AS '已交數量',[COPPURBATCHCOPTD].[TD010] AS '單位',[COPPURBATCHCOPTD].[TD024] AS '贈品量',[COPPURBATCHCOPTD].[TD025] AS '贈品已交量',MB001,MB002,MB003,TC015,TD020
                                    FROM [TKWAREHOUSE].[dbo].[COPPURBATCHCOPTD]
                                    LEFT JOIN [TK].dbo.COPTC ON COPTC.TC001=[COPPURBATCHCOPTD].[TD001]  AND COPTC.TC002=[COPPURBATCHCOPTD].[TD002] 
                                    LEFT JOIN [TK].dbo.COPTD ON COPTD.TD001=[COPPURBATCHCOPTD].[TD001]  AND COPTD.TD002=[COPPURBATCHCOPTD].[TD002] AND COPTD.TD003=[COPPURBATCHCOPTD].[TD003] 
                                    LEFT JOIN [TK].dbo.INVMB ON [COPPURBATCHCOPTD].[TD004]=INVMB.MB001
                                    WHERE  [ID]='{0}'
                                    ", ID);

                adapter2 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder2 = new SqlCommandBuilder(adapter2);
                sqlConn.Open();
                ds2.Clear();
                adapter2.Fill(ds2, "ds2");
                sqlConn.Close();


                if (ds2.Tables["ds2"].Rows.Count == 0)
                {
                    dataGridView2.DataSource = null;
                }
                else
                {
                    if (ds2.Tables["ds2"].Rows.Count >= 1)
                    {
                        //dataGridView1.Rows.Clear();
                        dataGridView2.DataSource = ds2.Tables["ds2"];
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
        public void ADDCOPPURBATCHCOPTD(string ID,string TD001,string TD002,string TD003)
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

                sbSql.AppendFormat(" INSERT INTO [TKWAREHOUSE].[dbo].[COPPURBATCHCOPTD]");
                sbSql.AppendFormat(" ([ID],[TD001],[TD002],[TD003],[TD004],[TD005],[TD008],[TD009],[TD010],[TD024],[TD025])");
                sbSql.AppendFormat(" SELECT '{0}',[TD001],[TD002],[TD003],[TD004],[TD005],[TD008],[TD009],[TD010],[TD024],[TD025]",ID);
                sbSql.AppendFormat(" FROM [TK].dbo.COPTD");
                sbSql.AppendFormat(" WHERE TD001='{0}' AND TD002='{1}' AND TD003='{2}'",TD001,TD002,TD003);
                sbSql.AppendFormat(" ");
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
        private void dataGridView2_SelectionChanged(object sender, EventArgs e)
        {
            //MOCTA的生產品號在dataGridView2_SelectionChanged決定
            //
            DELTD001 = null;
            DELTD002 = null;
            DELTD003 = null;

            MB001 = null;
            MB002 = null;
            MB003 = null;
            BAR = 0;
            TA026A = null;
            TA027A = null;
            TA028A = null;
            SUM1 = 0;
            TC015TD020 = null;

            if (dataGridView2.CurrentRow != null)
            {
                int rowindex = dataGridView2.CurrentRow.Index;
                if (rowindex >= 0)
                {
                    DataGridViewRow row = dataGridView2.Rows[rowindex];
                    DELTD001 = row.Cells["訂單單別"].Value.ToString();
                    DELTD002 = row.Cells["訂單單號"].Value.ToString();
                    DELTD003 = row.Cells["訂單序號"].Value.ToString();
                    MB001 = row.Cells["品號"].Value.ToString();
                    MB002 = row.Cells["MB002"].Value.ToString();
                    MB003 = row.Cells["MB003"].Value.ToString();
                    BAR = Convert.ToDecimal(row.Cells["訂單數量"].Value.ToString());
                    TA026A = row.Cells["訂單單別"].Value.ToString();
                    TA027A = row.Cells["訂單單號"].Value.ToString();
                    TA028A = row.Cells["訂單序號"].Value.ToString();
                    SUM1 = Convert.ToDecimal(row.Cells["訂單數量"].Value.ToString());
                    TC015TD020 = (row.Cells["TC015"].Value.ToString()+ row.Cells["TD020"].Value.ToString());

                }
                else
                {
                    DELTD001 = null;
                    DELTD002 = null;
                    DELTD003 = null;

                    MB001 = null;
                    MB002 = null;
                    MB003 = null;
                    BAR = 0;
                    TA026A = null;
                    TA027A = null;
                    TA028A = null;
                    SUM1 = 0;
                    TC015TD020 = null;

                }
            }
        }

        public void DELCOPPURBATCHCOPTD(string ID, string TD001, string TD002, string TD003)
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

                sbSql.AppendFormat(" DELETE [TKWAREHOUSE].[dbo].[COPPURBATCHCOPTD]");
                sbSql.AppendFormat(" WHERE [ID]='{0}' AND TD001='{1}' AND TD002='{2}' AND TD003='{3}'",ID, TD001, TD002, TD003);
                sbSql.AppendFormat(" ");
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

        public void ADDCOPPURBATCHUSED(string ID)
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

               
                sbSql.AppendFormat(" ");
                sbSql.AppendFormat(@" 
                                    DELETE [TKWAREHOUSE].[dbo].[COPPURBATCHUSED]
                                    WHERE [ID]='{0}'
 
                                    INSERT INTO [TKWAREHOUSE].[dbo].[COPPURBATCHUSED]
                                    ([ID],[TD001],[TD002],[TD003],[TD004],[TD005],[TDNUM],[TDUNIT],[MB001],[MB002],[NUM],[UNIT])
                                    SELECT '{0}',TD001,TD002,TD003,TD004,TD005,NUM,MB004,MD003,MD035,CASE WHEN [MD003] LIKE '2%' THEN ROUND((NUM*CAL),0) ELSE (NUM*CAL) END,MD004
                                    FROM (
                                    SELECT   TD001,TD002,TD003,TC053 ,TD013,TD004,TD005,TD006
                                    ,((CASE WHEN MB004=TD010 THEN ((TD008-TD009)+(TD024-TD025)) ELSE ((TD008-TD009)+(TD024-TD025))*INVMD.MD004 END)-ISNULL(MOCTA.TA017,0)) AS 'NUM'
                                    ,MB004
                                    ,((TD008-TD009)+(TD024-TD025)) AS 'COPNUM'
                                    ,TD010
                                    ,(CASE WHEN ISNULL(INVMD.MD002,'')<>'' THEN INVMD.MD002 ELSE TD010 END ) AS INVMDMD002
                                    ,(CASE WHEN INVMD.MD003>0 THEN INVMD.MD003 ELSE 1 END) AS INVMDMD003
                                    ,(CASE WHEN INVMD.MD004>0 THEN INVMD.MD004 ELSE (TD008-TD009) END ) AS INVMDMD004
                                    ,ISNULL(MOCTA.TA017,0) AS TA017
                                    ,[MC001],[MC004],BOMMD.[MD003],[MD035],BOMMD.[MD006],BOMMD.[MD007],BOMMD.[MD008],BOMMD.[MD004]
                                    ,CONVERT(decimal(16,4),(1/[MC004]*BOMMD.[MD006]/BOMMD.[MD007]*(1+BOMMD.[MD008]))) AS CAL
                                    FROM [TK].dbo.BOMMC,[TK].dbo.BOMMD,[TK].dbo.INVMB,[TK].dbo.COPTC,[TK].dbo.COPTD
                                    LEFT JOIN [TK].dbo.INVMD ON MD001=TD004 AND TD010=MD002
                                    LEFT JOIN [TK].dbo.MOCTA ON TA026=TD001 AND TA027=TD002 AND TD028=TD003 AND TA006=TD004
                                    WHERE BOMMC.MC001=BOMMD.MD001
                                    AND  BOMMD.MD001=TD004
                                    AND TD004=MB001
                                    AND TC001=TD001 AND TC002=TD002
                                    AND TD001+TD002+TD003 IN (SELECT TD001+TD002+TD003 FROM [TKWAREHOUSE].[dbo].[COPPURBATCHCOPTD] WHERE ID='{0}')
                                    ) AS TEMP
                                    WHERE (MD003 LIKE '1%' OR MD003 LIKE '2%' OR MD003 LIKE '3%' OR MD003 LIKE '4%' )  
                                    ", ID);


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

        public void SEARCHCOPPURBATCHUSED(string ID)
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

                sbSql.AppendFormat(@"  SELECT [ID] AS '批號',[TD001] AS '訂單單別',[TD002] AS '訂單單號',[TD003] AS '訂單序號',[TD004] AS '成品',[TD005] AS '品名',[TDNUM] AS '訂單數量',[TDUNIT] AS '成品單位',[MB001] AS '用品',[MB002] AS '用品名',[NUM] AS '需求量',[UNIT] AS '需求單位'");
                sbSql.AppendFormat(@"  FROM [TKWAREHOUSE].[dbo].[COPPURBATCHUSED]");
                sbSql.AppendFormat(@"  WHERE [ID]='{0}'",ID);
                sbSql.AppendFormat(@"  ");

                adapter3 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder3 = new SqlCommandBuilder(adapter3);
                sqlConn.Open();
                ds3.Clear();
                adapter3.Fill(ds3, "ds3");
                sqlConn.Close();


                if (ds3.Tables["ds3"].Rows.Count == 0)
                {
                    dataGridView3.DataSource = null;
                }
                else
                {
                    if (ds3.Tables["ds3"].Rows.Count >= 1)
                    {
                        //dataGridView1.Rows.Clear();
                        dataGridView3.DataSource = ds3.Tables["ds3"];
                        dataGridView3.AutoResizeColumns();
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

        public string GETMAXMOCTA002(string MOCTA001, string MOCTA003) // 假設 MOCTA003 是從某處傳入的參數
        {            
            string sqlQuery = @"
                                SELECT ISNULL(MAX(TA002), '00000000000') AS ID
                                FROM [TK].[dbo].[PURTA]
                                WHERE 1=1
                                AND [TA001] = @MOCTA_001
                                AND [TA003] = @MOCTA_003";

            try
            {
                Class1 TKID = new Class1();

                SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(
                    ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString
                );
                
                sqlsb.Password = TKID.Decryption(sqlsb.Password);
                sqlsb.UserID = TKID.Decryption(sqlsb.UserID);
                
                using (SqlConnection sqlConn = new SqlConnection(sqlsb.ConnectionString))
                {
                    sqlConn.Open();
                    using (SqlCommand cmd = new SqlCommand(sqlQuery, sqlConn))
                    {
                        cmd.Parameters.AddWithValue("@MOCTA_001", MOCTA001);
                        cmd.Parameters.AddWithValue("@MOCTA_003", MOCTA003);

                        object result = cmd.ExecuteScalar();
                        if (result != null && result != DBNull.Value)
                        {
                            string maxIDFromDB = result.ToString();
                            string maxID = SETIDSTRING(maxIDFromDB, MOCTA003);

                            return maxID;
                        }
                        return null;
                    } 
                } 
            }
            catch (Exception ex)
            {
               
                System.Diagnostics.Debug.WriteLine($"Error in GETMAXMOCTA002: {ex.Message}");
                return null;
            }
            
        }

        /// <summary>
        /// 根據現有的最大 ID (序列號) 和一個日期字串，產生新的編號。
        /// <para>此版本符合 C# 5.0 語法規範。</para>
        /// </summary>
        /// <param name="MAXID">從資料庫查詢到的最大 ID，例如: '20251211001' 或 '00000000000'。</param>
        /// <param name="dt">當前的日期/前綴字串，例如: '20251211'。</param>
        /// <returns>返回新的序列號字串 (例如: '20251211002')。</returns>
        /// <exception cref="FormatException">當 MAXID 格式不正確或無法解析序列號時拋出。</exception>
        public string SETIDSTRING(string MAXID, string dt)
        {          
            if (string.IsNullOrEmpty(MAXID) || MAXID.Equals("00000000000"))
            {
                // 情況 1: 新的一天或第一筆資料
                return dt + "001";
            }
            
            if (MAXID.Length < 11)
            {
                throw new FormatException("MAXID 格式不正確，長度應至少為 11。");
            }
            
            string sernoString = MAXID.Substring(MAXID.Length - 3, 3);

            int serno;
            if (Int32.TryParse(sernoString, out serno))
            {
                serno++;
                
                string sernoFormatted = String.Format("{0:D3}", serno);
                // 組合新 ID
                return dt + sernoFormatted;
            }
            else
            {
                throw new FormatException(String.Format("MAXID 的序列號部分 ('{0}') 無法轉換為數字。", sernoString));
            }
        }

        public void ADDMOCTAB(string ID,string TYPE)
        {
            try
            {
                PURTA PURTA = new PURTA();
                PURTB PURTB = new PURTB();

                PURTA = SETPURTA();
                PURTB = SETPURTB();

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

                sbSql.AppendFormat(" INSERT INTO [TK].[dbo].[PURTA]");
                sbSql.AppendFormat(" ( [COMPANY],[CREATOR],[USR_GROUP],[CREATE_DATE],[MODIFIER]");
                sbSql.AppendFormat(" ,[MODI_DATE],[FLAG],[CREATE_TIME],[MODI_TIME],[TRANS_TYPE]");
                sbSql.AppendFormat(" ,[TRANS_NAME],[sync_date],[sync_time],[sync_mark],[sync_count]");
                sbSql.AppendFormat(" ,[DataUser],[DataGroup]");
                sbSql.AppendFormat(" ,[TA001],[TA002],[TA003],[TA004],[TA005]");
                sbSql.AppendFormat(" ,[TA006],[TA007],[TA008],[TA009],[TA010]");
                sbSql.AppendFormat(" ,[TA011],[TA012],[TA013],[TA014],[TA015]");
                sbSql.AppendFormat(" ,[TA016],[TA017],[TA018],[TA019],[TA020]");
                sbSql.AppendFormat(" ,[TA021],[TA022],[TA023],[TA024],[TA025]");
                sbSql.AppendFormat(" ,[TA026],[TA027],[TA028],[TA029],[TA030]");
                sbSql.AppendFormat(" ,[TA031],[TA032],[TA033],[TA034],[TA035]");
                sbSql.AppendFormat(" ,[TA036],[TA037],[TA038],[TA039],[TA040]");
                sbSql.AppendFormat(" ,[TA041],[TA042],[TA043],[TA044],[TA045]");
                sbSql.AppendFormat(" ,[TA046],[UDF01],[UDF02],[UDF03],[UDF04]");
                sbSql.AppendFormat(" ,[UDF05],[UDF06],[UDF07],[UDF08],[UDF09]");
                sbSql.AppendFormat(" ,[UDF10]");
                sbSql.AppendFormat(" )");
                sbSql.AppendFormat(" VALUES ");
                sbSql.AppendFormat(" ('{0}','{1}','{2}','{3}','{4}',", PURTA.COMPANY, PURTA.CREATOR, PURTA.USR_GROUP, PURTA.CREATE_DATE, PURTA.MODIFIER);
                sbSql.AppendFormat(" '{0}','{1}','{2}','{3}','{4}',", PURTA.MODI_DATE, PURTA.FLAG, PURTA.CREATE_TIME, PURTA.MODI_TIME, PURTA.TRANS_TYPE);
                sbSql.AppendFormat(" '{0}','{1}','{2}','{3}','{4}',", PURTA.TRANS_NAME, PURTA.sync_date, PURTA.sync_time, PURTA.sync_mark, PURTA.sync_count);
                sbSql.AppendFormat(" '{0}','{1}',", PURTA.DataUser, PURTA.DataGroup);
                sbSql.AppendFormat(" '{0}','{1}','{2}','{3}','{4}',", PURTA.TA001, PURTA.TA002, PURTA.TA003, PURTA.TA004, PURTA.TA005);
                sbSql.AppendFormat(" '{0}','{1}','{2}','{3}','{4}',", PURTA.TA006, PURTA.TA007, PURTA.TA008, PURTA.TA009, PURTA.TA010);
                sbSql.AppendFormat(" '{0}','{1}','{2}','{3}','{4}',", PURTA.TA011, PURTA.TA012, PURTA.TA013, PURTA.TA014, PURTA.TA015);
                sbSql.AppendFormat(" '{0}','{1}','{2}','{3}','{4}',", PURTA.TA016, PURTA.TA017, PURTA.TA018, PURTA.TA019, PURTA.TA020);
                sbSql.AppendFormat(" '{0}','{1}','{2}','{3}','{4}',", PURTA.TA021, PURTA.TA022, PURTA.TA023, PURTA.TA024, PURTA.TA025);
                sbSql.AppendFormat(" '{0}','{1}','{2}','{3}','{4}',", PURTA.TA026, PURTA.TA027, PURTA.TA028, PURTA.TA029, PURTA.TA030);
                sbSql.AppendFormat(" '{0}','{1}','{2}','{3}','{4}',", PURTA.TA031, PURTA.TA032, PURTA.TA033, PURTA.TA034, PURTA.TA035);
                sbSql.AppendFormat(" '{0}','{1}','{2}','{3}','{4}',", PURTA.TA036, PURTA.TA037, PURTA.TA038, PURTA.TA039, PURTA.TA040);
                sbSql.AppendFormat(" '{0}','{1}','{2}','{3}','{4}',", PURTA.TA041, PURTA.TA042, PURTA.TA043, PURTA.TA044, PURTA.TA045);
                sbSql.AppendFormat(" '{0}','{1}','{2}','{3}','{4}',", PURTA.TA046, PURTA.UDF01, PURTA.UDF02, PURTA.UDF03, PURTA.UDF04);
                sbSql.AppendFormat(" '{0}','{1}','{2}','{3}','{4}',", PURTA.UDF05, PURTA.UDF06, PURTA.UDF07, PURTA.UDF08, PURTA.UDF09);
                sbSql.AppendFormat(" '{0}'", PURTA.UDF10);
                sbSql.AppendFormat(" )");
                sbSql.AppendFormat(" ");
                sbSql.AppendFormat(" INSERT INTO [TK].[dbo].[PURTB]");
                sbSql.AppendFormat(" ([COMPANY],[CREATOR],[USR_GROUP],[CREATE_DATE],[MODIFIER]");
                sbSql.AppendFormat(" ,[MODI_DATE],[FLAG],[CREATE_TIME],[MODI_TIME],[TRANS_TYPE]");
                sbSql.AppendFormat(" ,[TRANS_NAME],[sync_date],[sync_time],[sync_mark],[sync_count]");
                sbSql.AppendFormat(" ,[DataUser],[DataGroup]");
                sbSql.AppendFormat(" ,[TB001],[TB002],[TB003],[TB004],[TB005]");
                sbSql.AppendFormat(" ,[TB006],[TB007],[TB008],[TB009],[TB010]");
                sbSql.AppendFormat(" ,[TB011],[TB012],[TB013],[TB014],[TB015]");
                sbSql.AppendFormat(" ,[TB016],[TB017],[TB018],[TB019],[TB020]");
                sbSql.AppendFormat(" ,[TB021],[TB022],[TB023],[TB024],[TB025]");
                sbSql.AppendFormat(" ,[TB026],[TB027],[TB028],[TB029],[TB030]");
                sbSql.AppendFormat(" ,[TB031],[TB032],[TB033],[TB034],[TB035]");
                sbSql.AppendFormat(" ,[TB036],[TB037],[TB038],[TB039],[TB040]");
                sbSql.AppendFormat(" ,[TB041],[TB042],[TB043],[TB044],[TB045]");
                sbSql.AppendFormat(" ,[TB046],[TB047],[TB048],[TB049],[TB050]");
                sbSql.AppendFormat(" ,[TB051],[TB052],[TB053],[TB054],[TB055]");
                sbSql.AppendFormat(" ,[TB056],[TB057],[TB058],[TB059],[TB060]");
                sbSql.AppendFormat(" ,[TB061],[TB062],[TB063],[TB064],[TB065]");
                sbSql.AppendFormat(" ,[TB066],[TB067],[TB068],[TB069],[TB070]");
                sbSql.AppendFormat(" ,[TB071],[TB072],[TB073],[TB074],[TB075]");
                sbSql.AppendFormat(" ,[TB076],[TB077],[TB078],[TB079],[TB080]");
                sbSql.AppendFormat(" ,[TB081],[TB082],[TB083],[TB084],[TB085]");
                sbSql.AppendFormat(" ,[TB086],[TB087],[TB088],[TB089],[TB090]");
                sbSql.AppendFormat(" ,[TB091],[TB092],[TB093],[TB094],[TB095]");
                sbSql.AppendFormat(" ,[TB096],[TB097],[TB098],[TB099],[UDF01]");
                sbSql.AppendFormat(" ,[UDF02],[UDF03],[UDF04],[UDF05],[UDF06]");
                sbSql.AppendFormat(" ,[UDF07],[UDF08],[UDF09],[UDF10]");
                sbSql.AppendFormat(" )");
                sbSql.AppendFormat(" (SELECT '{0}' [COMPANY],'{1}' [CREATOR],'{2}' [USR_GROUP],'{3}' [CREATE_DATE],'{4}' [MODIFIER],", PURTB.COMPANY, PURTB.CREATOR, PURTB.USR_GROUP, PURTB.CREATE_DATE, PURTB.MODIFIER);
                sbSql.AppendFormat(" '{0}' [MODI_DATE],{1} [FLAG],'{2}' [CREATE_TIME],'{3}' [MODI_TIME],'{4}' [TRANS_TYPE],", PURTB.MODI_DATE, PURTB.FLAG, PURTB.CREATE_TIME, PURTB.MODI_TIME, PURTB.TRANS_TYPE);
                sbSql.AppendFormat(" '{0}' [TRANS_NAME],'{1}' [sync_date],'{2}' [sync_time],'{3}' [sync_mark],{4} [sync_count],", PURTB.TRANS_NAME, PURTB.sync_date, PURTB.sync_time, PURTB.sync_mark, PURTB.sync_count);
                sbSql.AppendFormat(" '{0}' [DataUser],'{1}' [DataGroup],", PURTB.DataUser, PURTB.DataGroup);
                sbSql.AppendFormat(" '{0}' [TB001],'{1}' [TB002],Right('0000' + Cast(ROW_NUMBER() OVER( ORDER BY [COPPURBATCHUSED].[MB001])  as varchar),4) AS TB003,[COPPURBATCHUSED].[MB001] AS TB004,[COPPURBATCHUSED].[MB002] AS TB005,", PURTB.TB001, PURTB.TB002);
                sbSql.AppendFormat(" MB003 AS TB006,MB004 AS TB007,MB017 AS TB008,SUM([NUM]) AS TB009,MB032 AS TB010,");
                sbSql.AppendFormat(" '{0}' [TB011],[ID]+' '+[TD001]+'-'+[TD002]+'-'+[TD003] [TB012],'{1}' [TB013],0 [TB014],'{2}' [TB015],", PURTB.TB011, PURTB.TB013, PURTB.TB015);
                sbSql.AppendFormat(" '{0}' [TB016],MB050 AS TB017,ROUND((MB050*SUM([NUM])),0) AS TB018,'{1}' [TB019],'{2}' [TB020],", PURTB.TB016, PURTB.TB019, PURTB.TB020);
                sbSql.AppendFormat(" '{0}' [TB021],'{1}' [TB022],'{2}' [TB023],'{3}' [TB024],'{4}' [TB025],", PURTB.TB021, PURTB.TB022, PURTB.TB023, PURTB.TB024, PURTB.TB025);
                sbSql.AppendFormat(" '{0}' [TB026],'{1}' [TB027],'{2}' [TB028],'{3}' [TB029],'{4}' [TB030],", PURTB.TB026, PURTB.TB027, PURTB.TB028, PURTB.TB029, PURTB.TB030);
                sbSql.AppendFormat(" '{0}' [TB031],'{1}' [TB032],'{2}' [TB033],{3} [TB034],{4} [TB035],", PURTB.TB031, PURTB.TB032, PURTB.TB033, PURTB.TB034, PURTB.TB035);
                sbSql.AppendFormat(" '{0}' [TB036],'{1}' [TB037],'{2}' [TB038],'{3}' [TB039],'{4}' [TB040],", PURTB.TB036, PURTB.TB037, PURTB.TB038, PURTB.TB039, PURTB.TB040);
                sbSql.AppendFormat(" {0} [TB041],'{1}' [TB042],'{2}' [TB043],'{3}' [TB044],'{4}' [TB045],", PURTB.TB041, PURTB.TB042, PURTB.TB043, PURTB.TB044, PURTB.TB045);
                sbSql.AppendFormat(" '{0}' [TB046],'{1}' [TB047],'{2}' [TB048],{3} [TB049],'{4}' [TB050],", PURTB.TB046, PURTB.TB047, PURTB.TB048, PURTB.TB049, PURTB.TB050);
                sbSql.AppendFormat(" {0} [TB051],{1} [TB052],{2} [TB053],'{3}' [TB054],'{4}' [TB055],", PURTB.TB051, PURTB.TB052, PURTB.TB053, PURTB.TB054, PURTB.TB055);
                sbSql.AppendFormat(" '{0}' [TB056],'{1}' [TB057],'{2}' [TB058],'{3}' [TB059],'{4}' [TB060],", PURTB.TB056, PURTB.TB057, PURTB.TB058, PURTB.TB059, PURTB.TB060);
                sbSql.AppendFormat(" '{0}' [TB061],'{1}' [TB062],{2} [TB063],'{3}' [TB064],'{4}' [TB065],", PURTB.TB061, PURTB.TB062, PURTB.TB063, PURTB.TB064, PURTB.TB065);
                sbSql.AppendFormat(" '{0}' [TB066],'{1}' [TB067],{2} [TB068],{3} [TB069],'{4}' [TB070],", PURTB.TB066, PURTB.TB067, PURTB.TB068, PURTB.TB069, PURTB.TB070);
                sbSql.AppendFormat(" '{0}' [TB071],'{1}' [TB072],'{2}' [TB073],'{3}' [TB074],{4} [TB075],", PURTB.TB071, PURTB.TB072, PURTB.TB073, PURTB.TB074, PURTB.TB075);
                sbSql.AppendFormat(" '{0}' [TB076],{1} [TB077],'{2}' [TB078],'{3}' [TB079],'{4}' [TB080],", PURTB.TB076, PURTB.TB077, PURTB.TB078, PURTB.TB079, PURTB.TB080);
                sbSql.AppendFormat(" {0} [TB081],{1} [TB082],{2} [TB083],{3} [TB084],{4} [TB085],", PURTB.TB081, PURTB.TB082, PURTB.TB083, PURTB.TB084, PURTB.TB085);
                sbSql.AppendFormat(" '{0}' [TB086],'{1}' [TB087],{2} [TB088],'{3}' [TB089],{4} [TB090],", PURTB.TB086, PURTB.TB087, PURTB.TB088, PURTB.TB089, PURTB.TB090);
                sbSql.AppendFormat(" {0} [TB091],{1} [TB092],{2} [TB093],'{3}' [TB094],'{4}' [TB095],", PURTB.TB091, PURTB.TB092, PURTB.TB093, PURTB.TB094, PURTB.TB095);
                sbSql.AppendFormat(" '{0}' [TB096],'{1}' [TB097],'{2}' [TB098],'{3}' [TB099],'{4}' [UDF01],", PURTB.TB096, PURTB.TB097, PURTB.TB098, PURTB.TB099, PURTB.UDF01);
                sbSql.AppendFormat(" '{0}' [UDF02],'{1}' [UDF03],'{2}' [UDF04],'{3}' [UDF05],{4} [UDF06],", PURTB.UDF02, PURTB.UDF03, PURTB.UDF04, PURTB.UDF05, PURTB.UDF06);
                sbSql.AppendFormat(" {0} [UDF07],{1}[UDF08],{2} [UDF09],{3} [UDF10]", PURTB.UDF07, PURTB.UDF08, PURTB.UDF09, PURTB.UDF10);
                sbSql.AppendFormat(" FROM [TKWAREHOUSE].[dbo].[COPPURBATCHUSED],[TK].dbo.INVMB");
                sbSql.AppendFormat(" WHERE [COPPURBATCHUSED].[MB001]=INVMB.[MB001]");
                sbSql.AppendFormat(" AND ([COPPURBATCHUSED].[MB001] LIKE '{0}%')", TYPE);
                sbSql.AppendFormat(" AND [ID]='{0}'", ID);
                sbSql.AppendFormat(" GROUP BY [ID]+' '+[TD001]+'-'+[TD002]+'-'+[TD003],[COPPURBATCHUSED].[MB001],[COPPURBATCHUSED].[MB002],MB003,MB004,MB017,MB032,MB050 )");
                sbSql.AppendFormat(" ");
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

                    UPDATEPURTA();
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

        public void ADDMOCTAB2(string ID)
        {
            PURTA PURTA = new PURTA();
            PURTB PURTB = new PURTB();

            PURTA = SETPURTA();
            PURTB = SETPURTB();

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

            sbSql.AppendFormat(" INSERT INTO [TK].[dbo].[PURTA]");
            sbSql.AppendFormat(" ( [COMPANY],[CREATOR],[USR_GROUP],[CREATE_DATE],[MODIFIER]");
            sbSql.AppendFormat(" ,[MODI_DATE],[FLAG],[CREATE_TIME],[MODI_TIME],[TRANS_TYPE]");
            sbSql.AppendFormat(" ,[TRANS_NAME],[sync_date],[sync_time],[sync_mark],[sync_count]");
            sbSql.AppendFormat(" ,[DataUser],[DataGroup]");
            sbSql.AppendFormat(" ,[TA001],[TA002],[TA003],[TA004],[TA005]");
            sbSql.AppendFormat(" ,[TA006],[TA007],[TA008],[TA009],[TA010]");
            sbSql.AppendFormat(" ,[TA011],[TA012],[TA013],[TA014],[TA015]");
            sbSql.AppendFormat(" ,[TA016],[TA017],[TA018],[TA019],[TA020]");
            sbSql.AppendFormat(" ,[TA021],[TA022],[TA023],[TA024],[TA025]");
            sbSql.AppendFormat(" ,[TA026],[TA027],[TA028],[TA029],[TA030]");
            sbSql.AppendFormat(" ,[TA031],[TA032],[TA033],[TA034],[TA035]");
            sbSql.AppendFormat(" ,[TA036],[TA037],[TA038],[TA039],[TA040]");
            sbSql.AppendFormat(" ,[TA041],[TA042],[TA043],[TA044],[TA045]");
            sbSql.AppendFormat(" ,[TA046],[UDF01],[UDF02],[UDF03],[UDF04]");
            sbSql.AppendFormat(" ,[UDF05],[UDF06],[UDF07],[UDF08],[UDF09]");
            sbSql.AppendFormat(" ,[UDF10]");
            sbSql.AppendFormat(" )");
            sbSql.AppendFormat(" VALUES ");
            sbSql.AppendFormat(" ('{0}','{1}','{2}','{3}','{4}',", PURTA.COMPANY, PURTA.CREATOR, PURTA.USR_GROUP, PURTA.CREATE_DATE, PURTA.MODIFIER);
            sbSql.AppendFormat(" '{0}','{1}','{2}','{3}','{4}',", PURTA.MODI_DATE, PURTA.FLAG, PURTA.CREATE_TIME, PURTA.MODI_TIME, PURTA.TRANS_TYPE);
            sbSql.AppendFormat(" '{0}','{1}','{2}','{3}','{4}',", PURTA.TRANS_NAME, PURTA.sync_date, PURTA.sync_time, PURTA.sync_mark, PURTA.sync_count);
            sbSql.AppendFormat(" '{0}','{1}',", PURTA.DataUser, PURTA.DataGroup);
            sbSql.AppendFormat(" '{0}','{1}','{2}','{3}','{4}',", PURTA.TA001, PURTA.TA002, PURTA.TA003, PURTA.TA004, PURTA.TA005);
            sbSql.AppendFormat(" '{0}','{1}','{2}','{3}','{4}',", PURTA.TA006, PURTA.TA007, PURTA.TA008, PURTA.TA009, PURTA.TA010);
            sbSql.AppendFormat(" '{0}','{1}','{2}','{3}','{4}',", PURTA.TA011, PURTA.TA012, PURTA.TA013, PURTA.TA014, PURTA.TA015);
            sbSql.AppendFormat(" '{0}','{1}','{2}','{3}','{4}',", PURTA.TA016, PURTA.TA017, PURTA.TA018, PURTA.TA019, PURTA.TA020);
            sbSql.AppendFormat(" '{0}','{1}','{2}','{3}','{4}',", PURTA.TA021, PURTA.TA022, PURTA.TA023, PURTA.TA024, PURTA.TA025);
            sbSql.AppendFormat(" '{0}','{1}','{2}','{3}','{4}',", PURTA.TA026, PURTA.TA027, PURTA.TA028, PURTA.TA029, PURTA.TA030);
            sbSql.AppendFormat(" '{0}','{1}','{2}','{3}','{4}',", PURTA.TA031, PURTA.TA032, PURTA.TA033, PURTA.TA034, PURTA.TA035);
            sbSql.AppendFormat(" '{0}','{1}','{2}','{3}','{4}',", PURTA.TA036, PURTA.TA037, PURTA.TA038, PURTA.TA039, PURTA.TA040);
            sbSql.AppendFormat(" '{0}','{1}','{2}','{3}','{4}',", PURTA.TA041, PURTA.TA042, PURTA.TA043, PURTA.TA044, PURTA.TA045);
            sbSql.AppendFormat(" '{0}','{1}','{2}','{3}','{4}',", PURTA.TA046, PURTA.UDF01, PURTA.UDF02, PURTA.UDF03, PURTA.UDF04);
            sbSql.AppendFormat(" '{0}','{1}','{2}','{3}','{4}',", PURTA.UDF05, PURTA.UDF06, PURTA.UDF07, PURTA.UDF08, PURTA.UDF09);
            sbSql.AppendFormat(" '{0}'", PURTA.UDF10);
            sbSql.AppendFormat(" )");
            sbSql.AppendFormat(" ");
            sbSql.AppendFormat(" INSERT INTO [TK].[dbo].[PURTB]");
            sbSql.AppendFormat(" ([COMPANY],[CREATOR],[USR_GROUP],[CREATE_DATE],[MODIFIER]");
            sbSql.AppendFormat(" ,[MODI_DATE],[FLAG],[CREATE_TIME],[MODI_TIME],[TRANS_TYPE]");
            sbSql.AppendFormat(" ,[TRANS_NAME],[sync_date],[sync_time],[sync_mark],[sync_count]");
            sbSql.AppendFormat(" ,[DataUser],[DataGroup]");
            sbSql.AppendFormat(" ,[TB001],[TB002],[TB003],[TB004],[TB005]");
            sbSql.AppendFormat(" ,[TB006],[TB007],[TB008],[TB009],[TB010]");
            sbSql.AppendFormat(" ,[TB011],[TB012],[TB013],[TB014],[TB015]");
            sbSql.AppendFormat(" ,[TB016],[TB017],[TB018],[TB019],[TB020]");
            sbSql.AppendFormat(" ,[TB021],[TB022],[TB023],[TB024],[TB025]");
            sbSql.AppendFormat(" ,[TB026],[TB027],[TB028],[TB029],[TB030]");
            sbSql.AppendFormat(" ,[TB031],[TB032],[TB033],[TB034],[TB035]");
            sbSql.AppendFormat(" ,[TB036],[TB037],[TB038],[TB039],[TB040]");
            sbSql.AppendFormat(" ,[TB041],[TB042],[TB043],[TB044],[TB045]");
            sbSql.AppendFormat(" ,[TB046],[TB047],[TB048],[TB049],[TB050]");
            sbSql.AppendFormat(" ,[TB051],[TB052],[TB053],[TB054],[TB055]");
            sbSql.AppendFormat(" ,[TB056],[TB057],[TB058],[TB059],[TB060]");
            sbSql.AppendFormat(" ,[TB061],[TB062],[TB063],[TB064],[TB065]");
            sbSql.AppendFormat(" ,[TB066],[TB067],[TB068],[TB069],[TB070]");
            sbSql.AppendFormat(" ,[TB071],[TB072],[TB073],[TB074],[TB075]");
            sbSql.AppendFormat(" ,[TB076],[TB077],[TB078],[TB079],[TB080]");
            sbSql.AppendFormat(" ,[TB081],[TB082],[TB083],[TB084],[TB085]");
            sbSql.AppendFormat(" ,[TB086],[TB087],[TB088],[TB089],[TB090]");
            sbSql.AppendFormat(" ,[TB091],[TB092],[TB093],[TB094],[TB095]");
            sbSql.AppendFormat(" ,[TB096],[TB097],[TB098],[TB099],[UDF01]");
            sbSql.AppendFormat(" ,[UDF02],[UDF03],[UDF04],[UDF05],[UDF06]");
            sbSql.AppendFormat(" ,[UDF07],[UDF08],[UDF09],[UDF10]");
            sbSql.AppendFormat(" )");
            sbSql.AppendFormat(" (SELECT '{0}' [COMPANY],'{1}' [CREATOR],'{2}' [USR_GROUP],'{3}' [CREATE_DATE],'{4}' [MODIFIER],", PURTB.COMPANY, PURTB.CREATOR, PURTB.USR_GROUP, PURTB.CREATE_DATE, PURTB.MODIFIER);
            sbSql.AppendFormat(" '{0}' [MODI_DATE],{1} [FLAG],'{2}' [CREATE_TIME],'{3}' [MODI_TIME],'{4}' [TRANS_TYPE],", PURTB.MODI_DATE, PURTB.FLAG, PURTB.CREATE_TIME, PURTB.MODI_TIME, PURTB.TRANS_TYPE);
            sbSql.AppendFormat(" '{0}' [TRANS_NAME],'{1}' [sync_date],'{2}' [sync_time],'{3}' [sync_mark],{4} [sync_count],", PURTB.TRANS_NAME, PURTB.sync_date, PURTB.sync_time, PURTB.sync_mark, PURTB.sync_count);
            sbSql.AppendFormat(" '{0}' [DataUser],'{1}' [DataGroup],", PURTB.DataUser, PURTB.DataGroup);
            sbSql.AppendFormat(" '{0}' [TB001],'{1}' [TB002],Right('0000' + Cast(ROW_NUMBER() OVER( ORDER BY INVMB.[MB001])  as varchar),4) AS TB003,INVMB.[MB001] AS TB004,INVMB.[MB002] AS TB005,", PURTB.TB001, PURTB.TB002);
            sbSql.AppendFormat(" MB003 AS TB006,TD010 AS TB007,MB017 AS TB008,SUM(TD008+TD024) AS TB009,MB032 AS TB010,");
            sbSql.AppendFormat(" '{0}' [TB011],[ID]+' '+[TD001]+'-'+[TD002]+'-'+[TD003] [TB012],'{1}' [TB013],0 [TB014],'{2}' [TB015],", PURTB.TB011, PURTB.TB013, PURTB.TB015);
            sbSql.AppendFormat(" '{0}' [TB016],MB050 AS TB017,ROUND((MB050*SUM([TD008])),0) AS TB018,'{1}' [TB019],'{2}' [TB020],", PURTB.TB016, PURTB.TB019, PURTB.TB020);
            sbSql.AppendFormat(" '{0}' [TB021],'{1}' [TB022],'{2}' [TB023],'{3}' [TB024],'{4}' [TB025],", PURTB.TB021, PURTB.TB022, PURTB.TB023, PURTB.TB024, PURTB.TB025);
            sbSql.AppendFormat(" '{0}' [TB026],'{1}' [TB027],'{2}' [TB028],'{3}' [TB029],'{4}' [TB030],", PURTB.TB026, PURTB.TB027, PURTB.TB028, PURTB.TB029, PURTB.TB030);
            sbSql.AppendFormat(" '{0}' [TB031],'{1}' [TB032],'{2}' [TB033],{3} [TB034],{4} [TB035],", PURTB.TB031, PURTB.TB032, PURTB.TB033, PURTB.TB034, PURTB.TB035);
            sbSql.AppendFormat(" '{0}' [TB036],'{1}' [TB037],'{2}' [TB038],'{3}' [TB039],'{4}' [TB040],", PURTB.TB036, PURTB.TB037, PURTB.TB038, PURTB.TB039, PURTB.TB040);
            sbSql.AppendFormat(" {0} [TB041],'{1}' [TB042],'{2}' [TB043],'{3}' [TB044],'{4}' [TB045],", PURTB.TB041, PURTB.TB042, PURTB.TB043, PURTB.TB044, PURTB.TB045);
            sbSql.AppendFormat(" '{0}' [TB046],'{1}' [TB047],'{2}' [TB048],{3} [TB049],'{4}' [TB050],", PURTB.TB046, PURTB.TB047, PURTB.TB048, PURTB.TB049, PURTB.TB050);
            sbSql.AppendFormat(" {0} [TB051],{1} [TB052],{2} [TB053],'{3}' [TB054],'{4}' [TB055],", PURTB.TB051, PURTB.TB052, PURTB.TB053, PURTB.TB054, PURTB.TB055);
            sbSql.AppendFormat(" '{0}' [TB056],'{1}' [TB057],'{2}' [TB058],'{3}' [TB059],'{4}' [TB060],", PURTB.TB056, PURTB.TB057, PURTB.TB058, PURTB.TB059, PURTB.TB060);
            sbSql.AppendFormat(" '{0}' [TB061],'{1}' [TB062],{2} [TB063],'{3}' [TB064],'{4}' [TB065],", PURTB.TB061, PURTB.TB062, PURTB.TB063, PURTB.TB064, PURTB.TB065);
            sbSql.AppendFormat(" '{0}' [TB066],'{1}' [TB067],{2} [TB068],{3} [TB069],'{4}' [TB070],", PURTB.TB066, PURTB.TB067, PURTB.TB068, PURTB.TB069, PURTB.TB070);
            sbSql.AppendFormat(" '{0}' [TB071],'{1}' [TB072],'{2}' [TB073],'{3}' [TB074],{4} [TB075],", PURTB.TB071, PURTB.TB072, PURTB.TB073, PURTB.TB074, PURTB.TB075);
            sbSql.AppendFormat(" '{0}' [TB076],{1} [TB077],'{2}' [TB078],'{3}' [TB079],'{4}' [TB080],", PURTB.TB076, PURTB.TB077, PURTB.TB078, PURTB.TB079, PURTB.TB080);
            sbSql.AppendFormat(" {0} [TB081],{1} [TB082],{2} [TB083],{3} [TB084],{4} [TB085],", PURTB.TB081, PURTB.TB082, PURTB.TB083, PURTB.TB084, PURTB.TB085);
            sbSql.AppendFormat(" '{0}' [TB086],'{1}' [TB087],{2} [TB088],'{3}' [TB089],{4} [TB090],", PURTB.TB086, PURTB.TB087, PURTB.TB088, PURTB.TB089, PURTB.TB090);
            sbSql.AppendFormat(" {0} [TB091],{1} [TB092],{2} [TB093],'{3}' [TB094],'{4}' [TB095],", PURTB.TB091, PURTB.TB092, PURTB.TB093, PURTB.TB094, PURTB.TB095);
            sbSql.AppendFormat(" '{0}' [TB096],'{1}' [TB097],'{2}' [TB098],'{3}' [TB099],'{4}' [UDF01],", PURTB.TB096, PURTB.TB097, PURTB.TB098, PURTB.TB099, PURTB.UDF01);
            sbSql.AppendFormat(" '{0}' [UDF02],'{1}' [UDF03],'{2}' [UDF04],'{3}' [UDF05],{4} [UDF06],", PURTB.UDF02, PURTB.UDF03, PURTB.UDF04, PURTB.UDF05, PURTB.UDF06);
            sbSql.AppendFormat(" {0} [UDF07],{1}[UDF08],{2} [UDF09],{3} [UDF10]", PURTB.UDF07, PURTB.UDF08, PURTB.UDF09, PURTB.UDF10);
            sbSql.AppendFormat(" FROM [TKWAREHOUSE].[dbo].[COPPURBATCHCOPTD],[TK].dbo.INVMB");
            sbSql.AppendFormat(" WHERE [TKWAREHOUSE].[dbo].[COPPURBATCHCOPTD].[TD004]=INVMB.[MB001]  ");
            sbSql.AppendFormat(" AND [ID]='{0}'",ID);
            sbSql.AppendFormat(" GROUP BY [ID]+' '+[TD001]+'-'+[TD002]+'-'+[TD003],INVMB.[MB001],INVMB.[MB002],MB003,TD010,MB017,MB032,MB050 )");
            sbSql.AppendFormat(" ");
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

                UPDATEPURTA();
            }
        }
        public void ADDMOCTAB3(string ID, string TYPE, string TYPE2)
        {
            try
            {
                PURTA PURTA = new PURTA();
                PURTB PURTB = new PURTB();

                PURTA = SETPURTA();
                PURTB = SETPURTB();

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

                sbSql.AppendFormat(" INSERT INTO [TK].[dbo].[PURTA]");
                sbSql.AppendFormat(" ( [COMPANY],[CREATOR],[USR_GROUP],[CREATE_DATE],[MODIFIER]");
                sbSql.AppendFormat(" ,[MODI_DATE],[FLAG],[CREATE_TIME],[MODI_TIME],[TRANS_TYPE]");
                sbSql.AppendFormat(" ,[TRANS_NAME],[sync_date],[sync_time],[sync_mark],[sync_count]");
                sbSql.AppendFormat(" ,[DataUser],[DataGroup]");
                sbSql.AppendFormat(" ,[TA001],[TA002],[TA003],[TA004],[TA005]");
                sbSql.AppendFormat(" ,[TA006],[TA007],[TA008],[TA009],[TA010]");
                sbSql.AppendFormat(" ,[TA011],[TA012],[TA013],[TA014],[TA015]");
                sbSql.AppendFormat(" ,[TA016],[TA017],[TA018],[TA019],[TA020]");
                sbSql.AppendFormat(" ,[TA021],[TA022],[TA023],[TA024],[TA025]");
                sbSql.AppendFormat(" ,[TA026],[TA027],[TA028],[TA029],[TA030]");
                sbSql.AppendFormat(" ,[TA031],[TA032],[TA033],[TA034],[TA035]");
                sbSql.AppendFormat(" ,[TA036],[TA037],[TA038],[TA039],[TA040]");
                sbSql.AppendFormat(" ,[TA041],[TA042],[TA043],[TA044],[TA045]");
                sbSql.AppendFormat(" ,[TA046],[UDF01],[UDF02],[UDF03],[UDF04]");
                sbSql.AppendFormat(" ,[UDF05],[UDF06],[UDF07],[UDF08],[UDF09]");
                sbSql.AppendFormat(" ,[UDF10]");
                sbSql.AppendFormat(" )");
                sbSql.AppendFormat(" VALUES ");
                sbSql.AppendFormat(" ('{0}','{1}','{2}','{3}','{4}',", PURTA.COMPANY, PURTA.CREATOR, PURTA.USR_GROUP, PURTA.CREATE_DATE, PURTA.MODIFIER);
                sbSql.AppendFormat(" '{0}','{1}','{2}','{3}','{4}',", PURTA.MODI_DATE, PURTA.FLAG, PURTA.CREATE_TIME, PURTA.MODI_TIME, PURTA.TRANS_TYPE);
                sbSql.AppendFormat(" '{0}','{1}','{2}','{3}','{4}',", PURTA.TRANS_NAME, PURTA.sync_date, PURTA.sync_time, PURTA.sync_mark, PURTA.sync_count);
                sbSql.AppendFormat(" '{0}','{1}',", PURTA.DataUser, PURTA.DataGroup);
                sbSql.AppendFormat(" '{0}','{1}','{2}','{3}','{4}',", PURTA.TA001, PURTA.TA002, PURTA.TA003, PURTA.TA004, PURTA.TA005);
                sbSql.AppendFormat(" '{0}','{1}','{2}','{3}','{4}',", PURTA.TA006, PURTA.TA007, PURTA.TA008, PURTA.TA009, PURTA.TA010);
                sbSql.AppendFormat(" '{0}','{1}','{2}','{3}','{4}',", PURTA.TA011, PURTA.TA012, PURTA.TA013, PURTA.TA014, PURTA.TA015);
                sbSql.AppendFormat(" '{0}','{1}','{2}','{3}','{4}',", PURTA.TA016, PURTA.TA017, PURTA.TA018, PURTA.TA019, PURTA.TA020);
                sbSql.AppendFormat(" '{0}','{1}','{2}','{3}','{4}',", PURTA.TA021, PURTA.TA022, PURTA.TA023, PURTA.TA024, PURTA.TA025);
                sbSql.AppendFormat(" '{0}','{1}','{2}','{3}','{4}',", PURTA.TA026, PURTA.TA027, PURTA.TA028, PURTA.TA029, PURTA.TA030);
                sbSql.AppendFormat(" '{0}','{1}','{2}','{3}','{4}',", PURTA.TA031, PURTA.TA032, PURTA.TA033, PURTA.TA034, PURTA.TA035);
                sbSql.AppendFormat(" '{0}','{1}','{2}','{3}','{4}',", PURTA.TA036, PURTA.TA037, PURTA.TA038, PURTA.TA039, PURTA.TA040);
                sbSql.AppendFormat(" '{0}','{1}','{2}','{3}','{4}',", PURTA.TA041, PURTA.TA042, PURTA.TA043, PURTA.TA044, PURTA.TA045);
                sbSql.AppendFormat(" '{0}','{1}','{2}','{3}','{4}',", PURTA.TA046, PURTA.UDF01, PURTA.UDF02, PURTA.UDF03, PURTA.UDF04);
                sbSql.AppendFormat(" '{0}','{1}','{2}','{3}','{4}',", PURTA.UDF05, PURTA.UDF06, PURTA.UDF07, PURTA.UDF08, PURTA.UDF09);
                sbSql.AppendFormat(" '{0}'", PURTA.UDF10);
                sbSql.AppendFormat(" )");
                sbSql.AppendFormat(" ");
                sbSql.AppendFormat(" INSERT INTO [TK].[dbo].[PURTB]");
                sbSql.AppendFormat(" ([COMPANY],[CREATOR],[USR_GROUP],[CREATE_DATE],[MODIFIER]");
                sbSql.AppendFormat(" ,[MODI_DATE],[FLAG],[CREATE_TIME],[MODI_TIME],[TRANS_TYPE]");
                sbSql.AppendFormat(" ,[TRANS_NAME],[sync_date],[sync_time],[sync_mark],[sync_count]");
                sbSql.AppendFormat(" ,[DataUser],[DataGroup]");
                sbSql.AppendFormat(" ,[TB001],[TB002],[TB003],[TB004],[TB005]");
                sbSql.AppendFormat(" ,[TB006],[TB007],[TB008],[TB009],[TB010]");
                sbSql.AppendFormat(" ,[TB011],[TB012],[TB013],[TB014],[TB015]");
                sbSql.AppendFormat(" ,[TB016],[TB017],[TB018],[TB019],[TB020]");
                sbSql.AppendFormat(" ,[TB021],[TB022],[TB023],[TB024],[TB025]");
                sbSql.AppendFormat(" ,[TB026],[TB027],[TB028],[TB029],[TB030]");
                sbSql.AppendFormat(" ,[TB031],[TB032],[TB033],[TB034],[TB035]");
                sbSql.AppendFormat(" ,[TB036],[TB037],[TB038],[TB039],[TB040]");
                sbSql.AppendFormat(" ,[TB041],[TB042],[TB043],[TB044],[TB045]");
                sbSql.AppendFormat(" ,[TB046],[TB047],[TB048],[TB049],[TB050]");
                sbSql.AppendFormat(" ,[TB051],[TB052],[TB053],[TB054],[TB055]");
                sbSql.AppendFormat(" ,[TB056],[TB057],[TB058],[TB059],[TB060]");
                sbSql.AppendFormat(" ,[TB061],[TB062],[TB063],[TB064],[TB065]");
                sbSql.AppendFormat(" ,[TB066],[TB067],[TB068],[TB069],[TB070]");
                sbSql.AppendFormat(" ,[TB071],[TB072],[TB073],[TB074],[TB075]");
                sbSql.AppendFormat(" ,[TB076],[TB077],[TB078],[TB079],[TB080]");
                sbSql.AppendFormat(" ,[TB081],[TB082],[TB083],[TB084],[TB085]");
                sbSql.AppendFormat(" ,[TB086],[TB087],[TB088],[TB089],[TB090]");
                sbSql.AppendFormat(" ,[TB091],[TB092],[TB093],[TB094],[TB095]");
                sbSql.AppendFormat(" ,[TB096],[TB097],[TB098],[TB099],[UDF01]");
                sbSql.AppendFormat(" ,[UDF02],[UDF03],[UDF04],[UDF05],[UDF06]");
                sbSql.AppendFormat(" ,[UDF07],[UDF08],[UDF09],[UDF10]");
                sbSql.AppendFormat(" )");
                sbSql.AppendFormat(" (SELECT '{0}' [COMPANY],'{1}' [CREATOR],'{2}' [USR_GROUP],'{3}' [CREATE_DATE],'{4}' [MODIFIER],", PURTB.COMPANY, PURTB.CREATOR, PURTB.USR_GROUP, PURTB.CREATE_DATE, PURTB.MODIFIER);
                sbSql.AppendFormat(" '{0}' [MODI_DATE],{1} [FLAG],'{2}' [CREATE_TIME],'{3}' [MODI_TIME],'{4}' [TRANS_TYPE],", PURTB.MODI_DATE, PURTB.FLAG, PURTB.CREATE_TIME, PURTB.MODI_TIME, PURTB.TRANS_TYPE);
                sbSql.AppendFormat(" '{0}' [TRANS_NAME],'{1}' [sync_date],'{2}' [sync_time],'{3}' [sync_mark],{4} [sync_count],", PURTB.TRANS_NAME, PURTB.sync_date, PURTB.sync_time, PURTB.sync_mark, PURTB.sync_count);
                sbSql.AppendFormat(" '{0}' [DataUser],'{1}' [DataGroup],", PURTB.DataUser, PURTB.DataGroup);
                sbSql.AppendFormat(" '{0}' [TB001],'{1}' [TB002],Right('0000' + Cast(ROW_NUMBER() OVER( ORDER BY [COPPURBATCHUSED].[MB001])  as varchar),4) AS TB003,[COPPURBATCHUSED].[MB001] AS TB004,[COPPURBATCHUSED].[MB002] AS TB005,", PURTB.TB001, PURTB.TB002);
                sbSql.AppendFormat(" MB003 AS TB006,MB004 AS TB007,MB017 AS TB008,SUM([NUM]) AS TB009,MB032 AS TB010,");
                sbSql.AppendFormat(" '{0}' [TB011],[ID]+' '+[TD001]+'-'+[TD002]+'-'+[TD003] [TB012],'{1}' [TB013],0 [TB014],'{2}' [TB015],", PURTB.TB011, PURTB.TB013, PURTB.TB015);
                sbSql.AppendFormat(" '{0}' [TB016],MB050 AS TB017,ROUND((MB050*SUM([NUM])),0) AS TB018,'{1}' [TB019],'{2}' [TB020],", PURTB.TB016, PURTB.TB019, PURTB.TB020);
                sbSql.AppendFormat(" '{0}' [TB021],'{1}' [TB022],'{2}' [TB023],'{3}' [TB024],'{4}' [TB025],", PURTB.TB021, PURTB.TB022, PURTB.TB023, PURTB.TB024, PURTB.TB025);
                sbSql.AppendFormat(" '{0}' [TB026],'{1}' [TB027],'{2}' [TB028],'{3}' [TB029],'{4}' [TB030],", PURTB.TB026, PURTB.TB027, PURTB.TB028, PURTB.TB029, PURTB.TB030);
                sbSql.AppendFormat(" '{0}' [TB031],'{1}' [TB032],'{2}' [TB033],{3} [TB034],{4} [TB035],", PURTB.TB031, PURTB.TB032, PURTB.TB033, PURTB.TB034, PURTB.TB035);
                sbSql.AppendFormat(" '{0}' [TB036],'{1}' [TB037],'{2}' [TB038],'{3}' [TB039],'{4}' [TB040],", PURTB.TB036, PURTB.TB037, PURTB.TB038, PURTB.TB039, PURTB.TB040);
                sbSql.AppendFormat(" {0} [TB041],'{1}' [TB042],'{2}' [TB043],'{3}' [TB044],'{4}' [TB045],", PURTB.TB041, PURTB.TB042, PURTB.TB043, PURTB.TB044, PURTB.TB045);
                sbSql.AppendFormat(" '{0}' [TB046],'{1}' [TB047],'{2}' [TB048],{3} [TB049],'{4}' [TB050],", PURTB.TB046, PURTB.TB047, PURTB.TB048, PURTB.TB049, PURTB.TB050);
                sbSql.AppendFormat(" {0} [TB051],{1} [TB052],{2} [TB053],'{3}' [TB054],'{4}' [TB055],", PURTB.TB051, PURTB.TB052, PURTB.TB053, PURTB.TB054, PURTB.TB055);
                sbSql.AppendFormat(" '{0}' [TB056],'{1}' [TB057],'{2}' [TB058],'{3}' [TB059],'{4}' [TB060],", PURTB.TB056, PURTB.TB057, PURTB.TB058, PURTB.TB059, PURTB.TB060);
                sbSql.AppendFormat(" '{0}' [TB061],'{1}' [TB062],{2} [TB063],'{3}' [TB064],'{4}' [TB065],", PURTB.TB061, PURTB.TB062, PURTB.TB063, PURTB.TB064, PURTB.TB065);
                sbSql.AppendFormat(" '{0}' [TB066],'{1}' [TB067],{2} [TB068],{3} [TB069],'{4}' [TB070],", PURTB.TB066, PURTB.TB067, PURTB.TB068, PURTB.TB069, PURTB.TB070);
                sbSql.AppendFormat(" '{0}' [TB071],'{1}' [TB072],'{2}' [TB073],'{3}' [TB074],{4} [TB075],", PURTB.TB071, PURTB.TB072, PURTB.TB073, PURTB.TB074, PURTB.TB075);
                sbSql.AppendFormat(" '{0}' [TB076],{1} [TB077],'{2}' [TB078],'{3}' [TB079],'{4}' [TB080],", PURTB.TB076, PURTB.TB077, PURTB.TB078, PURTB.TB079, PURTB.TB080);
                sbSql.AppendFormat(" {0} [TB081],{1} [TB082],{2} [TB083],{3} [TB084],{4} [TB085],", PURTB.TB081, PURTB.TB082, PURTB.TB083, PURTB.TB084, PURTB.TB085);
                sbSql.AppendFormat(" '{0}' [TB086],'{1}' [TB087],{2} [TB088],'{3}' [TB089],{4} [TB090],", PURTB.TB086, PURTB.TB087, PURTB.TB088, PURTB.TB089, PURTB.TB090);
                sbSql.AppendFormat(" {0} [TB091],{1} [TB092],{2} [TB093],'{3}' [TB094],'{4}' [TB095],", PURTB.TB091, PURTB.TB092, PURTB.TB093, PURTB.TB094, PURTB.TB095);
                sbSql.AppendFormat(" '{0}' [TB096],'{1}' [TB097],'{2}' [TB098],'{3}' [TB099],'{4}' [UDF01],", PURTB.TB096, PURTB.TB097, PURTB.TB098, PURTB.TB099, PURTB.UDF01);
                sbSql.AppendFormat(" '{0}' [UDF02],'{1}' [UDF03],'{2}' [UDF04],'{3}' [UDF05],{4} [UDF06],", PURTB.UDF02, PURTB.UDF03, PURTB.UDF04, PURTB.UDF05, PURTB.UDF06);
                sbSql.AppendFormat(" {0} [UDF07],{1}[UDF08],{2} [UDF09],{3} [UDF10]", PURTB.UDF07, PURTB.UDF08, PURTB.UDF09, PURTB.UDF10);
                sbSql.AppendFormat(" FROM [TKWAREHOUSE].[dbo].[COPPURBATCHUSED],[TK].dbo.INVMB");
                sbSql.AppendFormat(" WHERE [COPPURBATCHUSED].[MB001]=INVMB.[MB001]");
                sbSql.AppendFormat(" AND ([COPPURBATCHUSED].[MB001] LIKE '{0}%' OR [COPPURBATCHUSED].[MB001] LIKE '{1}%')", TYPE, TYPE2);
                sbSql.AppendFormat(" AND [ID]='{0}'", ID);
                sbSql.AppendFormat(" GROUP BY [ID]+' '+[TD001]+'-'+[TD002]+'-'+[TD003],[COPPURBATCHUSED].[MB001],[COPPURBATCHUSED].[MB002],MB003,MB004,MB017,MB032,MB050 )");
                sbSql.AppendFormat(" ");
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

                    UPDATEPURTA();
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

        public void ADDMOCTAB4(string ID)
        {
            PURTA PURTA = new PURTA();
            PURTB PURTB = new PURTB();

            PURTA = SETPURTA();
            PURTB = SETPURTB();

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

            sbSql.AppendFormat(" INSERT INTO [TK].[dbo].[PURTA]");
            sbSql.AppendFormat(" ( [COMPANY],[CREATOR],[USR_GROUP],[CREATE_DATE],[MODIFIER]");
            sbSql.AppendFormat(" ,[MODI_DATE],[FLAG],[CREATE_TIME],[MODI_TIME],[TRANS_TYPE]");
            sbSql.AppendFormat(" ,[TRANS_NAME],[sync_date],[sync_time],[sync_mark],[sync_count]");
            sbSql.AppendFormat(" ,[DataUser],[DataGroup]");
            sbSql.AppendFormat(" ,[TA001],[TA002],[TA003],[TA004],[TA005]");
            sbSql.AppendFormat(" ,[TA006],[TA007],[TA008],[TA009],[TA010]");
            sbSql.AppendFormat(" ,[TA011],[TA012],[TA013],[TA014],[TA015]");
            sbSql.AppendFormat(" ,[TA016],[TA017],[TA018],[TA019],[TA020]");
            sbSql.AppendFormat(" ,[TA021],[TA022],[TA023],[TA024],[TA025]");
            sbSql.AppendFormat(" ,[TA026],[TA027],[TA028],[TA029],[TA030]");
            sbSql.AppendFormat(" ,[TA031],[TA032],[TA033],[TA034],[TA035]");
            sbSql.AppendFormat(" ,[TA036],[TA037],[TA038],[TA039],[TA040]");
            sbSql.AppendFormat(" ,[TA041],[TA042],[TA043],[TA044],[TA045]");
            sbSql.AppendFormat(" ,[TA046],[UDF01],[UDF02],[UDF03],[UDF04]");
            sbSql.AppendFormat(" ,[UDF05],[UDF06],[UDF07],[UDF08],[UDF09]");
            sbSql.AppendFormat(" ,[UDF10]");
            sbSql.AppendFormat(" )");
            sbSql.AppendFormat(" VALUES ");
            sbSql.AppendFormat(" ('{0}','{1}','{2}','{3}','{4}',", PURTA.COMPANY, PURTA.CREATOR, PURTA.USR_GROUP, PURTA.CREATE_DATE, PURTA.MODIFIER);
            sbSql.AppendFormat(" '{0}','{1}','{2}','{3}','{4}',", PURTA.MODI_DATE, PURTA.FLAG, PURTA.CREATE_TIME, PURTA.MODI_TIME, PURTA.TRANS_TYPE);
            sbSql.AppendFormat(" '{0}','{1}','{2}','{3}','{4}',", PURTA.TRANS_NAME, PURTA.sync_date, PURTA.sync_time, PURTA.sync_mark, PURTA.sync_count);
            sbSql.AppendFormat(" '{0}','{1}',", PURTA.DataUser, PURTA.DataGroup);
            sbSql.AppendFormat(" '{0}','{1}','{2}','{3}','{4}',", PURTA.TA001, PURTA.TA002, PURTA.TA003, PURTA.TA004, PURTA.TA005);
            sbSql.AppendFormat(" '{0}','{1}','{2}','{3}','{4}',", PURTA.TA006, PURTA.TA007, PURTA.TA008, PURTA.TA009, PURTA.TA010);
            sbSql.AppendFormat(" '{0}','{1}','{2}','{3}','{4}',", PURTA.TA011, PURTA.TA012, PURTA.TA013, PURTA.TA014, PURTA.TA015);
            sbSql.AppendFormat(" '{0}','{1}','{2}','{3}','{4}',", PURTA.TA016, PURTA.TA017, PURTA.TA018, PURTA.TA019, PURTA.TA020);
            sbSql.AppendFormat(" '{0}','{1}','{2}','{3}','{4}',", PURTA.TA021, PURTA.TA022, PURTA.TA023, PURTA.TA024, PURTA.TA025);
            sbSql.AppendFormat(" '{0}','{1}','{2}','{3}','{4}',", PURTA.TA026, PURTA.TA027, PURTA.TA028, PURTA.TA029, PURTA.TA030);
            sbSql.AppendFormat(" '{0}','{1}','{2}','{3}','{4}',", PURTA.TA031, PURTA.TA032, PURTA.TA033, PURTA.TA034, PURTA.TA035);
            sbSql.AppendFormat(" '{0}','{1}','{2}','{3}','{4}',", PURTA.TA036, PURTA.TA037, PURTA.TA038, PURTA.TA039, PURTA.TA040);
            sbSql.AppendFormat(" '{0}','{1}','{2}','{3}','{4}',", PURTA.TA041, PURTA.TA042, PURTA.TA043, PURTA.TA044, PURTA.TA045);
            sbSql.AppendFormat(" '{0}','{1}','{2}','{3}','{4}',", PURTA.TA046, PURTA.UDF01, PURTA.UDF02, PURTA.UDF03, PURTA.UDF04);
            sbSql.AppendFormat(" '{0}','{1}','{2}','{3}','{4}',", PURTA.UDF05, PURTA.UDF06, PURTA.UDF07, PURTA.UDF08, PURTA.UDF09);
            sbSql.AppendFormat(" '{0}'", PURTA.UDF10);
            sbSql.AppendFormat(" )");
            sbSql.AppendFormat(" ");
            sbSql.AppendFormat(" INSERT INTO [TK].[dbo].[PURTB]");
            sbSql.AppendFormat(" ([COMPANY],[CREATOR],[USR_GROUP],[CREATE_DATE],[MODIFIER]");
            sbSql.AppendFormat(" ,[MODI_DATE],[FLAG],[CREATE_TIME],[MODI_TIME],[TRANS_TYPE]");
            sbSql.AppendFormat(" ,[TRANS_NAME],[sync_date],[sync_time],[sync_mark],[sync_count]");
            sbSql.AppendFormat(" ,[DataUser],[DataGroup]");
            sbSql.AppendFormat(" ,[TB001],[TB002],[TB003],[TB004],[TB005]");
            sbSql.AppendFormat(" ,[TB006],[TB007],[TB008],[TB009],[TB010]");
            sbSql.AppendFormat(" ,[TB011],[TB012],[TB013],[TB014],[TB015]");
            sbSql.AppendFormat(" ,[TB016],[TB017],[TB018],[TB019],[TB020]");
            sbSql.AppendFormat(" ,[TB021],[TB022],[TB023],[TB024],[TB025]");
            sbSql.AppendFormat(" ,[TB026],[TB027],[TB028],[TB029],[TB030]");
            sbSql.AppendFormat(" ,[TB031],[TB032],[TB033],[TB034],[TB035]");
            sbSql.AppendFormat(" ,[TB036],[TB037],[TB038],[TB039],[TB040]");
            sbSql.AppendFormat(" ,[TB041],[TB042],[TB043],[TB044],[TB045]");
            sbSql.AppendFormat(" ,[TB046],[TB047],[TB048],[TB049],[TB050]");
            sbSql.AppendFormat(" ,[TB051],[TB052],[TB053],[TB054],[TB055]");
            sbSql.AppendFormat(" ,[TB056],[TB057],[TB058],[TB059],[TB060]");
            sbSql.AppendFormat(" ,[TB061],[TB062],[TB063],[TB064],[TB065]");
            sbSql.AppendFormat(" ,[TB066],[TB067],[TB068],[TB069],[TB070]");
            sbSql.AppendFormat(" ,[TB071],[TB072],[TB073],[TB074],[TB075]");
            sbSql.AppendFormat(" ,[TB076],[TB077],[TB078],[TB079],[TB080]");
            sbSql.AppendFormat(" ,[TB081],[TB082],[TB083],[TB084],[TB085]");
            sbSql.AppendFormat(" ,[TB086],[TB087],[TB088],[TB089],[TB090]");
            sbSql.AppendFormat(" ,[TB091],[TB092],[TB093],[TB094],[TB095]");
            sbSql.AppendFormat(" ,[TB096],[TB097],[TB098],[TB099],[UDF01]");
            sbSql.AppendFormat(" ,[UDF02],[UDF03],[UDF04],[UDF05],[UDF06]");
            sbSql.AppendFormat(" ,[UDF07],[UDF08],[UDF09],[UDF10]");
            sbSql.AppendFormat(" )");
            sbSql.AppendFormat(" (SELECT '{0}' [COMPANY],'{1}' [CREATOR],'{2}' [USR_GROUP],'{3}' [CREATE_DATE],'{4}' [MODIFIER],", PURTB.COMPANY, PURTB.CREATOR, PURTB.USR_GROUP, PURTB.CREATE_DATE, PURTB.MODIFIER);
            sbSql.AppendFormat(" '{0}' [MODI_DATE],{1} [FLAG],'{2}' [CREATE_TIME],'{3}' [MODI_TIME],'{4}' [TRANS_TYPE],", PURTB.MODI_DATE, PURTB.FLAG, PURTB.CREATE_TIME, PURTB.MODI_TIME, PURTB.TRANS_TYPE);
            sbSql.AppendFormat(" '{0}' [TRANS_NAME],'{1}' [sync_date],'{2}' [sync_time],'{3}' [sync_mark],{4} [sync_count],", PURTB.TRANS_NAME, PURTB.sync_date, PURTB.sync_time, PURTB.sync_mark, PURTB.sync_count);
            sbSql.AppendFormat(" '{0}' [DataUser],'{1}' [DataGroup],", PURTB.DataUser, PURTB.DataGroup);
            sbSql.AppendFormat(" '{0}' [TB001],'{1}' [TB002],Right('0000' + Cast(ROW_NUMBER() OVER( ORDER BY INVMB.[MB001])  as varchar),4) AS TB003,INVMB.[MB001] AS TB004,INVMB.[MB002] AS TB005,", PURTB.TB001, PURTB.TB002);
            sbSql.AppendFormat(" MB003 AS TB006,MB004 AS TB007,MB017 AS TB008,SUM(TD008+TD024)*MD004 AS TB009,MB032 AS TB010,");
            sbSql.AppendFormat(" '{0}' [TB011],[ID]+' '+[TD001]+'-'+[TD002]+'-'+[TD003] [TB012],'{1}' [TB013],0 [TB014],'{2}' [TB015],", PURTB.TB011, PURTB.TB013, PURTB.TB015);
            sbSql.AppendFormat(" '{0}' [TB016],MB050 AS TB017,ROUND((MB050*SUM([TD008])),0) AS TB018,'{1}' [TB019],'{2}' [TB020],", PURTB.TB016, PURTB.TB019, PURTB.TB020);
            sbSql.AppendFormat(" '{0}' [TB021],'{1}' [TB022],'{2}' [TB023],'{3}' [TB024],'{4}' [TB025],", PURTB.TB021, PURTB.TB022, PURTB.TB023, PURTB.TB024, PURTB.TB025);
            sbSql.AppendFormat(" '{0}' [TB026],'{1}' [TB027],'{2}' [TB028],'{3}' [TB029],'{4}' [TB030],", PURTB.TB026, PURTB.TB027, PURTB.TB028, PURTB.TB029, PURTB.TB030);
            sbSql.AppendFormat(" '{0}' [TB031],'{1}' [TB032],'{2}' [TB033],{3} [TB034],{4} [TB035],", PURTB.TB031, PURTB.TB032, PURTB.TB033, PURTB.TB034, PURTB.TB035);
            sbSql.AppendFormat(" '{0}' [TB036],'{1}' [TB037],'{2}' [TB038],'{3}' [TB039],'{4}' [TB040],", PURTB.TB036, PURTB.TB037, PURTB.TB038, PURTB.TB039, PURTB.TB040);
            sbSql.AppendFormat(" {0} [TB041],'{1}' [TB042],'{2}' [TB043],'{3}' [TB044],'{4}' [TB045],", PURTB.TB041, PURTB.TB042, PURTB.TB043, PURTB.TB044, PURTB.TB045);
            sbSql.AppendFormat(" '{0}' [TB046],'{1}' [TB047],'{2}' [TB048],{3} [TB049],'{4}' [TB050],", PURTB.TB046, PURTB.TB047, PURTB.TB048, PURTB.TB049, PURTB.TB050);
            sbSql.AppendFormat(" {0} [TB051],{1} [TB052],{2} [TB053],'{3}' [TB054],'{4}' [TB055],", PURTB.TB051, PURTB.TB052, PURTB.TB053, PURTB.TB054, PURTB.TB055);
            sbSql.AppendFormat(" '{0}' [TB056],'{1}' [TB057],'{2}' [TB058],'{3}' [TB059],'{4}' [TB060],", PURTB.TB056, PURTB.TB057, PURTB.TB058, PURTB.TB059, PURTB.TB060);
            sbSql.AppendFormat(" '{0}' [TB061],'{1}' [TB062],{2} [TB063],'{3}' [TB064],'{4}' [TB065],", PURTB.TB061, PURTB.TB062, PURTB.TB063, PURTB.TB064, PURTB.TB065);
            sbSql.AppendFormat(" '{0}' [TB066],'{1}' [TB067],{2} [TB068],{3} [TB069],'{4}' [TB070],", PURTB.TB066, PURTB.TB067, PURTB.TB068, PURTB.TB069, PURTB.TB070);
            sbSql.AppendFormat(" '{0}' [TB071],'{1}' [TB072],'{2}' [TB073],'{3}' [TB074],{4} [TB075],", PURTB.TB071, PURTB.TB072, PURTB.TB073, PURTB.TB074, PURTB.TB075);
            sbSql.AppendFormat(" '{0}' [TB076],{1} [TB077],'{2}' [TB078],'{3}' [TB079],'{4}' [TB080],", PURTB.TB076, PURTB.TB077, PURTB.TB078, PURTB.TB079, PURTB.TB080);
            sbSql.AppendFormat(" {0} [TB081],{1} [TB082],{2} [TB083],{3} [TB084],{4} [TB085],", PURTB.TB081, PURTB.TB082, PURTB.TB083, PURTB.TB084, PURTB.TB085);
            sbSql.AppendFormat(" '{0}' [TB086],'{1}' [TB087],{2} [TB088],'{3}' [TB089],{4} [TB090],", PURTB.TB086, PURTB.TB087, PURTB.TB088, PURTB.TB089, PURTB.TB090);
            sbSql.AppendFormat(" {0} [TB091],{1} [TB092],{2} [TB093],'{3}' [TB094],'{4}' [TB095],", PURTB.TB091, PURTB.TB092, PURTB.TB093, PURTB.TB094, PURTB.TB095);
            sbSql.AppendFormat(" '{0}' [TB096],'{1}' [TB097],'{2}' [TB098],'{3}' [TB099],'{4}' [UDF01],", PURTB.TB096, PURTB.TB097, PURTB.TB098, PURTB.TB099, PURTB.UDF01);
            sbSql.AppendFormat(" '{0}' [UDF02],'{1}' [UDF03],'{2}' [UDF04],'{3}' [UDF05],{4} [UDF06],", PURTB.UDF02, PURTB.UDF03, PURTB.UDF04, PURTB.UDF05, PURTB.UDF06);
            sbSql.AppendFormat(" {0} [UDF07],{1}[UDF08],{2} [UDF09],{3} [UDF10]", PURTB.UDF07, PURTB.UDF08, PURTB.UDF09, PURTB.UDF10);
            sbSql.AppendFormat(" FROM [TKWAREHOUSE].[dbo].[COPPURBATCHCOPTD],[TK].dbo.INVMB,[TK].dbo.INVMD");
            sbSql.AppendFormat(" WHERE [TKWAREHOUSE].[dbo].[COPPURBATCHCOPTD].[TD004]=INVMB.[MB001]  ");
            sbSql.AppendFormat(" AND  MD001=MB001 AND TD010=MD002");
            sbSql.AppendFormat(" AND [ID]='{0}'", ID);
            sbSql.AppendFormat(" GROUP BY [ID]+' '+[TD001]+'-'+[TD002]+'-'+[TD003],INVMB.[MB001],INVMB.[MB002],MB003,TD010,MB017,MB032,MB050,INVMD.MD004,MB004  )");
            sbSql.AppendFormat(" ");
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

                UPDATEPURTA();
            }
        }
        public void UPDATEPURTA()
        {
            if (!string.IsNullOrEmpty(MOCTA001) && !string.IsNullOrEmpty(MOCTA002) && !string.IsNullOrEmpty(ID))
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

                //UPDATE TB039='N'
                sbSql.AppendFormat(" UPDATE  [TK].dbo.PURTB SET TB039='N' WHERE ISNULL(TB039,'')=''");
                sbSql.AppendFormat(" ");
                sbSql.AppendFormat(" UPDATE [TK].dbo.PURTB");
                sbSql.AppendFormat(" SET TB017=(SELECT TOP 1 TN008 FROM [TK].dbo.VPURTLMN WHERE  TM004=TB004 AND TL004=TB010 AND TN007<=TB009 ORDER BY TN008),TB018=ROUND((SELECT TOP 1 TN008 FROM [TK].dbo.VPURTLMN WHERE  TM004=TB004 AND TL004=TB010 AND TN007<=TB009 ORDER BY TN008)*TB009,0)");
                sbSql.AppendFormat(" FROM [TK].dbo.VPURTLMN");
                sbSql.AppendFormat(" WHERE  TL004=TB010 AND TM004=TB004");
                sbSql.AppendFormat(" AND TB001='{0}' AND TB002='{1}'", MOCTA001, MOCTA002);
                sbSql.AppendFormat(" ");
                sbSql.AppendFormat(" UPDATE  [TK].dbo.PURTA");
                sbSql.AppendFormat(" SET TA011=(SELECT SUM(TB009) FROM [TK].dbo.PURTB WHERE PURTA.TA001=PURTB.TB001 AND  PURTA.TA002=PURTB.TB002)");
                sbSql.AppendFormat(" WHERE TA001='{0}' AND TA002='{1}'", MOCTA001, MOCTA002);
                sbSql.AppendFormat(" ");
                sbSql.AppendFormat(" UPDATE [TKWAREHOUSE].[dbo].[PURTAB]");
                sbSql.AppendFormat(" SET [PURTA001]='{0}',[PURTA002]='{1}'", MOCTA001, MOCTA002);
                sbSql.AppendFormat(" WHERE [ID]='{0}'", ID);
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

        }

        public PURTA SETPURTA()
        {
            PURTA PURTA = new PURTA();

            PURTA.COMPANY = "TK";
            PURTA.CREATOR = "120025";
            PURTA.USR_GROUP = "103400";
            //MOCTA.CREATE_DATE = dt1.ToString("yyyyMMdd");
            PURTA.CREATE_DATE = DateTime.Now.ToString("yyyyMMdd");
            PURTA.MODIFIER = "160115";
            PURTA.MODI_DATE = DateTime.Now.ToString("yyyyMMdd");
            PURTA.FLAG = "0";
            PURTA.CREATE_TIME = DateTime.Now.ToString("HH:mm:dd");
            PURTA.MODI_TIME = DateTime.Now.ToString("HH:mm:dd");
            PURTA.TRANS_TYPE = "P001";
            PURTA.TRANS_NAME = "PURI05";
            PURTA.sync_date = null;
            PURTA.sync_time = null;
            PURTA.sync_mark = null;
            PURTA.sync_count = null;
            PURTA.sync_count = "0";
            PURTA.DataUser = null;
            PURTA.DataGroup = null;
            PURTA.DataGroup = "103400";
            PURTA.TA001 = MOCTA001;
            PURTA.TA002 = MOCTA002;
            PURTA.TA003 = MOCTA003;
            PURTA.TA004 = "103500";
            PURTA.TA005 = ID;
            PURTA.TA006 = null;
            PURTA.TA007 = "N";
            PURTA.TA008 = "0";
            PURTA.TA009 = "9";
            PURTA.TA010 = "20";
            PURTA.TA011 = "0";
            PURTA.TA012 = "190006";
            PURTA.TA013 = MOCTA003;
            PURTA.TA014 = null;
            PURTA.TA015 = "0";
            PURTA.TA016 = "N";
            PURTA.TA017 = "0";
            PURTA.TA018 = null;
            PURTA.TA019 = null;
            PURTA.TA020 = "0";
            PURTA.TA021 = null;
            PURTA.TA022 = null;
            PURTA.TA023 = "0";
            PURTA.TA024 = "0";
            PURTA.TA025 = null;
            PURTA.TA026 = null;
            PURTA.TA027 = null;
            PURTA.TA028 = null;
            PURTA.TA029 = null;
            PURTA.TA030 = "0";
            PURTA.TA031 = null;
            PURTA.TA032 = "0";
            PURTA.TA033 = null;
            PURTA.TA034 = null;
            PURTA.TA035 = null;
            PURTA.TA036 = "0";
            PURTA.TA037 = "0";
            PURTA.TA038 = "0";
            PURTA.TA039 = "0";
            PURTA.TA040 = "0";
            PURTA.TA041 = null;
            PURTA.TA042 = null;
            PURTA.TA043 = null;
            PURTA.TA044 = null;
            PURTA.TA045 = null;
            PURTA.TA046 = null;
            PURTA.UDF01 = null;
            PURTA.UDF02 = null;
            PURTA.UDF03 = null;
            PURTA.UDF04 = null;
            PURTA.UDF05 = null;
            PURTA.UDF06 = "0";
            PURTA.UDF07 = "0";
            PURTA.UDF08 = "0";
            PURTA.UDF09 = "0";
            PURTA.UDF10 = "0";

            return PURTA;
        }


        public PURTB SETPURTB()
        {
            PURTB PURTB = new PURTB();

            PURTB.COMPANY = "TK";
            PURTB.CREATOR = "120025";
            PURTB.USR_GROUP = "103400";
            PURTB.CREATE_DATE = DateTime.Now.ToString("yyyyMMdd");
            PURTB.MODIFIER = "160115";
            PURTB.MODI_DATE = DateTime.Now.ToString("yyyyMMdd");
            PURTB.FLAG = "0";
            PURTB.CREATE_TIME = DateTime.Now.ToString("HH:mm:dd");
            PURTB.MODI_TIME = DateTime.Now.ToString("HH:mm:dd");
            PURTB.TRANS_TYPE = "P001";
            PURTB.TRANS_NAME = "PURI05";
            PURTB.sync_count = "0";
            PURTB.TB001 = MOCTA001;
            PURTB.TB002 = MOCTA002;
            PURTB.TB003 = null;
            PURTB.TB004 = null;
            PURTB.TB005 = null;
            PURTB.TB006 = null;
            PURTB.TB007 = null;
            PURTB.TB008 = null;
            PURTB.TB009 = null;
            PURTB.TB010 = null;
            PURTB.TB011 = MOCTA003;
            PURTB.TB012 = null;
            PURTB.TB013 = null;
            PURTB.TB014 = "0";
            PURTB.TB015 = null;
            PURTB.TB016 = "NTD";
            PURTB.TB017 = null;
            PURTB.TB018 = null;
            PURTB.TB019 = MOCTA003;
            PURTB.TB020 = "N";
            PURTB.TB021 = "N";
            PURTB.TB022 = null;
            PURTB.TB023 = null;
            PURTB.TB024 = null;
            PURTB.TB025 = "N";
            PURTB.TB026 = "2";
            PURTB.TB027 = null;
            PURTB.TB028 = null;
            PURTB.TB029 = null;
            PURTB.TB030 = null;
            PURTB.TB031 = null;
            PURTB.TB032 = "N";
            PURTB.TB033 = null;
            PURTB.TB034 = "0";
            PURTB.TB035 = "0";
            PURTB.TB036 = null;
            PURTB.TB037 = null;
            PURTB.TB038 = null;
            PURTB.TB039 = "N";
            PURTB.TB040 = "0";
            PURTB.TB041 = "0";
            PURTB.TB042 = null;
            PURTB.TB043 = null;
            PURTB.TB044 = null;
            PURTB.TB045 = null;
            PURTB.TB046 = null;
            PURTB.TB047 = null;
            PURTB.TB048 = null;
            PURTB.TB049 = "0";
            PURTB.TB050 = null;
            PURTB.TB051 = "0";
            PURTB.TB052 = "0";
            PURTB.TB053 = "0";
            PURTB.TB054 = null;
            PURTB.TB055 = null;
            PURTB.TB056 = null;
            PURTB.TB057 = null;
            PURTB.TB058 = "1";
            PURTB.TB059 = null;
            PURTB.TB060 = null;
            PURTB.TB061 = null;
            PURTB.TB062 = null;
            PURTB.TB063 = "0";
            PURTB.TB064 = null;
            PURTB.TB065 = null;
            PURTB.TB066 = null;
            PURTB.TB067 = "2";
            PURTB.TB068 = "0";
            PURTB.TB069 = "0";
            PURTB.TB070 = null;
            PURTB.TB071 = null;
            PURTB.TB072 = null;
            PURTB.TB073 = null;
            PURTB.TB074 = null;
            PURTB.TB075 = "0";
            PURTB.TB076 = null;
            PURTB.TB077 = "0";
            PURTB.TB078 = null;
            PURTB.TB079 = null;
            PURTB.TB080 = null;
            PURTB.TB081 = "0";
            PURTB.TB082 = "0";
            PURTB.TB083 = "0";
            PURTB.TB084 = "0";
            PURTB.TB085 = "0";
            PURTB.TB086 = null;
            PURTB.TB087 = null;
            PURTB.TB088 = "0";
            PURTB.TB089 = "1";
            PURTB.TB090 = "0";
            PURTB.TB091 = "0";
            PURTB.TB092 = "0";
            PURTB.TB093 = "0";
            PURTB.TB094 = null;
            PURTB.TB095 = null;
            PURTB.TB096 = null;
            PURTB.TB097 = null;
            PURTB.TB098 = null;
            PURTB.TB099 = null;
            PURTB.UDF01 = null;
            PURTB.UDF02 = null;
            PURTB.UDF03 = null;
            PURTB.UDF04 = null;
            PURTB.UDF05 = null;
            PURTB.UDF06 = "0";
            PURTB.UDF07 = "0";
            PURTB.UDF08 = "0";
            PURTB.UDF09 = "0";
            PURTB.UDF10 = "0";
            return PURTB;
        }

      
        public void SEARCHCOPPURBATCHPUR(string ID)
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

                sbSql.AppendFormat(@"  SELECT [TA001] AS '單別',[TA002] AS '單號',[ID] AS '批號'");
                sbSql.AppendFormat(@"  FROM [TKWAREHOUSE].[dbo].[COPPURBATCHPUR]");
                sbSql.AppendFormat(@"  WHERE [ID]='{0}' ",ID);
                sbSql.AppendFormat(@"  ");

                adapter5 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder5 = new SqlCommandBuilder(adapter5);
                sqlConn.Open();
                ds5.Clear();
                adapter5.Fill(ds5, "ds5");
                sqlConn.Close();


                if (ds5.Tables["ds5"].Rows.Count == 0)
                {
                    dataGridView4.DataSource = null;
                }
                else
                {
                    if (ds5.Tables["ds5"].Rows.Count >= 1)
                    {
                        //dataGridView1.Rows.Clear();
                        dataGridView4.DataSource = ds5.Tables["ds5"];
                        dataGridView4.AutoResizeColumns();
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

        public void ADDCOPPURBATCHPUR(string ID,string TA001,string TA002)
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


            sbSql.AppendFormat(@" 
                                INSERT INTO  [TKWAREHOUSE].[dbo].[COPPURBATCHPUR]
                                ([ID],[TA001],[TA002])
                                VALUES ('{0}','{1}','{2}')
                                ", ID, TA001, TA002); 


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

        public void DELETECOPPURBATCHPUR(string ID, string TA001, string TA002)
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


            sbSql.AppendFormat(@"
                                DELETE [TKWAREHOUSE].[dbo].[COPPURBATCHPUR]
                                WHERE [ID]='{0}' AND [TA001]='{1}'AND [TA002]='{2}'

                                ", ID, TA001, TA002);


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

        public void ADDPURTAB(string ID, string TYPE)
        {
            try
            {
                PURTA PURTA = new PURTA();
                PURTB PURTB = new PURTB();

                PURTA = SETPURTA();
                PURTB = SETPURTB();

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

                sbSql.AppendFormat(" INSERT INTO [TK].[dbo].[PURTA]");
                sbSql.AppendFormat(" ( [COMPANY],[CREATOR],[USR_GROUP],[CREATE_DATE],[MODIFIER]");
                sbSql.AppendFormat(" ,[MODI_DATE],[FLAG],[CREATE_TIME],[MODI_TIME],[TRANS_TYPE]");
                sbSql.AppendFormat(" ,[TRANS_NAME],[sync_date],[sync_time],[sync_mark],[sync_count]");
                sbSql.AppendFormat(" ,[DataUser],[DataGroup]");
                sbSql.AppendFormat(" ,[TA001],[TA002],[TA003],[TA004],[TA005]");
                sbSql.AppendFormat(" ,[TA006],[TA007],[TA008],[TA009],[TA010]");
                sbSql.AppendFormat(" ,[TA011],[TA012],[TA013],[TA014],[TA015]");
                sbSql.AppendFormat(" ,[TA016],[TA017],[TA018],[TA019],[TA020]");
                sbSql.AppendFormat(" ,[TA021],[TA022],[TA023],[TA024],[TA025]");
                sbSql.AppendFormat(" ,[TA026],[TA027],[TA028],[TA029],[TA030]");
                sbSql.AppendFormat(" ,[TA031],[TA032],[TA033],[TA034],[TA035]");
                sbSql.AppendFormat(" ,[TA036],[TA037],[TA038],[TA039],[TA040]");
                sbSql.AppendFormat(" ,[TA041],[TA042],[TA043],[TA044],[TA045]");
                sbSql.AppendFormat(" ,[TA046],[UDF01],[UDF02],[UDF03],[UDF04]");
                sbSql.AppendFormat(" ,[UDF05],[UDF06],[UDF07],[UDF08],[UDF09]");
                sbSql.AppendFormat(" ,[UDF10]");
                sbSql.AppendFormat(" )");
                sbSql.AppendFormat(" VALUES ");
                sbSql.AppendFormat(" ('{0}','{1}','{2}','{3}','{4}',", PURTA.COMPANY, PURTA.CREATOR, PURTA.USR_GROUP, PURTA.CREATE_DATE, PURTA.MODIFIER);
                sbSql.AppendFormat(" '{0}','{1}','{2}','{3}','{4}',", PURTA.MODI_DATE, PURTA.FLAG, PURTA.CREATE_TIME, PURTA.MODI_TIME, PURTA.TRANS_TYPE);
                sbSql.AppendFormat(" '{0}','{1}','{2}','{3}','{4}',", PURTA.TRANS_NAME, PURTA.sync_date, PURTA.sync_time, PURTA.sync_mark, PURTA.sync_count);
                sbSql.AppendFormat(" '{0}','{1}',", PURTA.DataUser, PURTA.DataGroup);
                sbSql.AppendFormat(" '{0}','{1}','{2}','{3}','{4}',", PURTA.TA001, PURTA.TA002, PURTA.TA003, PURTA.TA004, PURTA.TA005);
                sbSql.AppendFormat(" '{0}','{1}','{2}','{3}','{4}',", PURTA.TA006, PURTA.TA007, PURTA.TA008, PURTA.TA009, PURTA.TA010);
                sbSql.AppendFormat(" '{0}','{1}','{2}','{3}','{4}',", PURTA.TA011, PURTA.TA012, PURTA.TA013, PURTA.TA014, PURTA.TA015);
                sbSql.AppendFormat(" '{0}','{1}','{2}','{3}','{4}',", PURTA.TA016, PURTA.TA017, PURTA.TA018, PURTA.TA019, PURTA.TA020);
                sbSql.AppendFormat(" '{0}','{1}','{2}','{3}','{4}',", PURTA.TA021, PURTA.TA022, PURTA.TA023, PURTA.TA024, PURTA.TA025);
                sbSql.AppendFormat(" '{0}','{1}','{2}','{3}','{4}',", PURTA.TA026, PURTA.TA027, PURTA.TA028, PURTA.TA029, PURTA.TA030);
                sbSql.AppendFormat(" '{0}','{1}','{2}','{3}','{4}',", PURTA.TA031, PURTA.TA032, PURTA.TA033, PURTA.TA034, PURTA.TA035);
                sbSql.AppendFormat(" '{0}','{1}','{2}','{3}','{4}',", PURTA.TA036, PURTA.TA037, PURTA.TA038, PURTA.TA039, PURTA.TA040);
                sbSql.AppendFormat(" '{0}','{1}','{2}','{3}','{4}',", PURTA.TA041, PURTA.TA042, PURTA.TA043, PURTA.TA044, PURTA.TA045);
                sbSql.AppendFormat(" '{0}','{1}','{2}','{3}','{4}',", PURTA.TA046, PURTA.UDF01, PURTA.UDF02, PURTA.UDF03, PURTA.UDF04);
                sbSql.AppendFormat(" '{0}','{1}','{2}','{3}','{4}',", PURTA.UDF05, PURTA.UDF06, PURTA.UDF07, PURTA.UDF08, PURTA.UDF09);
                sbSql.AppendFormat(" '{0}'", PURTA.UDF10);
                sbSql.AppendFormat(" )");
                sbSql.AppendFormat(" ");
                sbSql.AppendFormat(" INSERT INTO [TK].[dbo].[PURTB]");
                sbSql.AppendFormat(" ([COMPANY],[CREATOR],[USR_GROUP],[CREATE_DATE],[MODIFIER]");
                sbSql.AppendFormat(" ,[MODI_DATE],[FLAG],[CREATE_TIME],[MODI_TIME],[TRANS_TYPE]");
                sbSql.AppendFormat(" ,[TRANS_NAME],[sync_date],[sync_time],[sync_mark],[sync_count]");
                sbSql.AppendFormat(" ,[DataUser],[DataGroup]");
                sbSql.AppendFormat(" ,[TB001],[TB002],[TB003],[TB004],[TB005]");
                sbSql.AppendFormat(" ,[TB006],[TB007],[TB008],[TB009],[TB010]");
                sbSql.AppendFormat(" ,[TB011],[TB012],[TB013],[TB014],[TB015]");
                sbSql.AppendFormat(" ,[TB016],[TB017],[TB018],[TB019],[TB020]");
                sbSql.AppendFormat(" ,[TB021],[TB022],[TB023],[TB024],[TB025]");
                sbSql.AppendFormat(" ,[TB026],[TB027],[TB028],[TB029],[TB030]");
                sbSql.AppendFormat(" ,[TB031],[TB032],[TB033],[TB034],[TB035]");
                sbSql.AppendFormat(" ,[TB036],[TB037],[TB038],[TB039],[TB040]");
                sbSql.AppendFormat(" ,[TB041],[TB042],[TB043],[TB044],[TB045]");
                sbSql.AppendFormat(" ,[TB046],[TB047],[TB048],[TB049],[TB050]");
                sbSql.AppendFormat(" ,[TB051],[TB052],[TB053],[TB054],[TB055]");
                sbSql.AppendFormat(" ,[TB056],[TB057],[TB058],[TB059],[TB060]");
                sbSql.AppendFormat(" ,[TB061],[TB062],[TB063],[TB064],[TB065]");
                sbSql.AppendFormat(" ,[TB066],[TB067],[TB068],[TB069],[TB070]");
                sbSql.AppendFormat(" ,[TB071],[TB072],[TB073],[TB074],[TB075]");
                sbSql.AppendFormat(" ,[TB076],[TB077],[TB078],[TB079],[TB080]");
                sbSql.AppendFormat(" ,[TB081],[TB082],[TB083],[TB084],[TB085]");
                sbSql.AppendFormat(" ,[TB086],[TB087],[TB088],[TB089],[TB090]");
                sbSql.AppendFormat(" ,[TB091],[TB092],[TB093],[TB094],[TB095]");
                sbSql.AppendFormat(" ,[TB096],[TB097],[TB098],[TB099],[UDF01]");
                sbSql.AppendFormat(" ,[UDF02],[UDF03],[UDF04],[UDF05],[UDF06]");
                sbSql.AppendFormat(" ,[UDF07],[UDF08],[UDF09],[UDF10]");
                sbSql.AppendFormat(" )");
                sbSql.AppendFormat(" (SELECT '{0}' [COMPANY],'{1}' [CREATOR],'{2}' [USR_GROUP],'{3}' [CREATE_DATE],'{4}' [MODIFIER],", PURTB.COMPANY, PURTB.CREATOR, PURTB.USR_GROUP, PURTB.CREATE_DATE, PURTB.MODIFIER);
                sbSql.AppendFormat(" '{0}' [MODI_DATE],{1} [FLAG],'{2}' [CREATE_TIME],'{3}' [MODI_TIME],'{4}' [TRANS_TYPE],", PURTB.MODI_DATE, PURTB.FLAG, PURTB.CREATE_TIME, PURTB.MODI_TIME, PURTB.TRANS_TYPE);
                sbSql.AppendFormat(" '{0}' [TRANS_NAME],'{1}' [sync_date],'{2}' [sync_time],'{3}' [sync_mark],{4} [sync_count],", PURTB.TRANS_NAME, PURTB.sync_date, PURTB.sync_time, PURTB.sync_mark, PURTB.sync_count);
                sbSql.AppendFormat(" '{0}' [DataUser],'{1}' [DataGroup],", PURTB.DataUser, PURTB.DataGroup);
                sbSql.AppendFormat(" '{0}' [TB001],'{1}' [TB002],Right('0000' + Cast(ROW_NUMBER() OVER( ORDER BY [COPPURBATCHUSED].[MB001])  as varchar),4) AS TB003,[COPPURBATCHUSED].[MB001] AS TB004,[COPPURBATCHUSED].[MB002] AS TB005,", PURTB.TB001, PURTB.TB002);
                sbSql.AppendFormat(" MB003 AS TB006,TDUNIT AS TB007,MB017 AS TB008,SUM([TDNUM]) AS TB009,MB032 AS TB010,");
                sbSql.AppendFormat(" '{0}' [TB011],[ID]+' '+[TD001]+'-'+[TD002]+'-'+[TD003] [TB012],'{1}' [TB013],0 [TB014],'{2}' [TB015],", PURTB.TB011, PURTB.TB013, PURTB.TB015);
                sbSql.AppendFormat(" '{0}' [TB016],MB050 AS TB017,ROUND((MB050*SUM([NUM])),0) AS TB018,'{1}' [TB019],'{2}' [TB020],", PURTB.TB016, PURTB.TB019, PURTB.TB020);
                sbSql.AppendFormat(" '{0}' [TB021],'{1}' [TB022],'{2}' [TB023],'{3}' [TB024],'{4}' [TB025],", PURTB.TB021, PURTB.TB022, PURTB.TB023, PURTB.TB024, PURTB.TB025);
                sbSql.AppendFormat(" '{0}' [TB026],'{1}' [TB027],'{2}' [TB028],'{3}' [TB029],'{4}' [TB030],", PURTB.TB026, PURTB.TB027, PURTB.TB028, PURTB.TB029, PURTB.TB030);
                sbSql.AppendFormat(" '{0}' [TB031],'{1}' [TB032],'{2}' [TB033],{3} [TB034],{4} [TB035],", PURTB.TB031, PURTB.TB032, PURTB.TB033, PURTB.TB034, PURTB.TB035);
                sbSql.AppendFormat(" '{0}' [TB036],'{1}' [TB037],'{2}' [TB038],'{3}' [TB039],'{4}' [TB040],", PURTB.TB036, PURTB.TB037, PURTB.TB038, PURTB.TB039, PURTB.TB040);
                sbSql.AppendFormat(" {0} [TB041],'{1}' [TB042],'{2}' [TB043],'{3}' [TB044],'{4}' [TB045],", PURTB.TB041, PURTB.TB042, PURTB.TB043, PURTB.TB044, PURTB.TB045);
                sbSql.AppendFormat(" '{0}' [TB046],'{1}' [TB047],'{2}' [TB048],{3} [TB049],'{4}' [TB050],", PURTB.TB046, PURTB.TB047, PURTB.TB048, PURTB.TB049, PURTB.TB050);
                sbSql.AppendFormat(" {0} [TB051],{1} [TB052],{2} [TB053],'{3}' [TB054],'{4}' [TB055],", PURTB.TB051, PURTB.TB052, PURTB.TB053, PURTB.TB054, PURTB.TB055);
                sbSql.AppendFormat(" '{0}' [TB056],'{1}' [TB057],'{2}' [TB058],'{3}' [TB059],'{4}' [TB060],", PURTB.TB056, PURTB.TB057, PURTB.TB058, PURTB.TB059, PURTB.TB060);
                sbSql.AppendFormat(" '{0}' [TB061],'{1}' [TB062],{2} [TB063],'{3}' [TB064],'{4}' [TB065],", PURTB.TB061, PURTB.TB062, PURTB.TB063, PURTB.TB064, PURTB.TB065);
                sbSql.AppendFormat(" '{0}' [TB066],'{1}' [TB067],{2} [TB068],{3} [TB069],'{4}' [TB070],", PURTB.TB066, PURTB.TB067, PURTB.TB068, PURTB.TB069, PURTB.TB070);
                sbSql.AppendFormat(" '{0}' [TB071],'{1}' [TB072],'{2}' [TB073],'{3}' [TB074],{4} [TB075],", PURTB.TB071, PURTB.TB072, PURTB.TB073, PURTB.TB074, PURTB.TB075);
                sbSql.AppendFormat(" '{0}' [TB076],{1} [TB077],'{2}' [TB078],'{3}' [TB079],'{4}' [TB080],", PURTB.TB076, PURTB.TB077, PURTB.TB078, PURTB.TB079, PURTB.TB080);
                sbSql.AppendFormat(" {0} [TB081],{1} [TB082],{2} [TB083],{3} [TB084],{4} [TB085],", PURTB.TB081, PURTB.TB082, PURTB.TB083, PURTB.TB084, PURTB.TB085);
                sbSql.AppendFormat(" '{0}' [TB086],'{1}' [TB087],{2} [TB088],'{3}' [TB089],{4} [TB090],", PURTB.TB086, PURTB.TB087, PURTB.TB088, PURTB.TB089, PURTB.TB090);
                sbSql.AppendFormat(" {0} [TB091],{1} [TB092],{2} [TB093],'{3}' [TB094],'{4}' [TB095],", PURTB.TB091, PURTB.TB092, PURTB.TB093, PURTB.TB094, PURTB.TB095);
                sbSql.AppendFormat(" '{0}' [TB096],'{1}' [TB097],'{2}' [TB098],'{3}' [TB099],'{4}' [UDF01],", PURTB.TB096, PURTB.TB097, PURTB.TB098, PURTB.TB099, PURTB.UDF01);
                sbSql.AppendFormat(" '{0}' [UDF02],'{1}' [UDF03],'{2}' [UDF04],'{3}' [UDF05],{4} [UDF06],", PURTB.UDF02, PURTB.UDF03, PURTB.UDF04, PURTB.UDF05, PURTB.UDF06);
                sbSql.AppendFormat(" {0} [UDF07],{1}[UDF08],{2} [UDF09],{3} [UDF10]", PURTB.UDF07, PURTB.UDF08, PURTB.UDF09, PURTB.UDF10);
                sbSql.AppendFormat(" FROM [TKWAREHOUSE].[dbo].[COPPURBATCHUSED],[TK].dbo.INVMB");
                sbSql.AppendFormat(" WHERE [COPPURBATCHUSED].[MB001]=INVMB.[MB001]");
                sbSql.AppendFormat(" AND ([COPPURBATCHUSED].[MB001] LIKE '{0}%')", TYPE);
                sbSql.AppendFormat(" AND [ID]='{0}'", ID);
                sbSql.AppendFormat(" GROUP BY [ID]+' '+[TD001]+'-'+[TD002]+'-'+[TD003],[COPPURBATCHUSED].[MB001],[COPPURBATCHUSED].[MB002],MB003,MB004,MB017,MB032,MB050,TDUNIT )");
                sbSql.AppendFormat(" ");
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

                    UPDATEPURTA();
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

        private void textBox4_TextChanged(object sender, EventArgs e)
        {
            textBox5.Text=SEARCHMB002(textBox4.Text.Trim());
        }

        private void textBox6_TextChanged(object sender, EventArgs e)
        {
            textBox7.Text = SEARCHPURMA002(textBox6.Text.Trim());
        }

        private void textBox8_TextChanged(object sender, EventArgs e)
        {
            textBox9.Text = SEARCHCMSMC002(textBox8.Text.Trim());
        }

        public string SEARCHMB002(string MB001)
        {
            SqlDataAdapter adapter1 = new SqlDataAdapter();
            SqlCommandBuilder sqlCmdBuilder1 = new SqlCommandBuilder();
            DataSet ds1 = new DataSet();

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
                                    SELECT MB001,MB002 
                                    FROM [TK].dbo.INVMB
                                    WHERE MB001='{0}'
                                      ", MB001);

                adapter1 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder1 = new SqlCommandBuilder(adapter1);
                sqlConn.Open();
                ds1.Clear();
                adapter1.Fill(ds1, "ds1");
                sqlConn.Close();


                if (ds1.Tables["ds1"].Rows.Count >= 1)
                {
                    return ds1.Tables["ds1"].Rows[0]["MB002"].ToString().Trim();
                }
                else
                {
                    return "";
                }

            }
            catch
            {
                return "";
            }
            finally
            {
                sqlConn.Close();
            }
        }

        public string SEARCHPURMA002(string MA001)
        {
            SqlDataAdapter adapter1 = new SqlDataAdapter();
            SqlCommandBuilder sqlCmdBuilder1 = new SqlCommandBuilder();
            DataSet ds1 = new DataSet();

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
                                    SELECT MA001,MA002
                                    FROM [TK].dbo.PURMA
                                    WHERE MA001='{0}'
                                      ", MA001);

                adapter1 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder1 = new SqlCommandBuilder(adapter1);
                sqlConn.Open();
                ds1.Clear();
                adapter1.Fill(ds1, "ds1");
                sqlConn.Close();


                if (ds1.Tables["ds1"].Rows.Count >= 1)
                {
                    return ds1.Tables["ds1"].Rows[0]["MA002"].ToString().Trim(); ;
                }
                else
                {
                    return "";
                }

            }
            catch
            {
                return "";
            }
            finally
            {
                sqlConn.Close();
            }
        }

        public string SEARCHCMSMC002(string MC001)
        {
            SqlDataAdapter adapter1 = new SqlDataAdapter();
            SqlCommandBuilder sqlCmdBuilder1 = new SqlCommandBuilder();
            DataSet ds1 = new DataSet();

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
                                    SELECT MC001,MC002
                                    FROM [TK].dbo.CMSMC
                                    WHERE MC001='{0}'
                                      ", MC001);

                adapter1 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder1 = new SqlCommandBuilder(adapter1);
                sqlConn.Open();
                ds1.Clear();
                adapter1.Fill(ds1, "ds1");
                sqlConn.Close();


                if (ds1.Tables["ds1"].Rows.Count >= 1)
                {
                    return ds1.Tables["ds1"].Rows[0]["MC002"].ToString().Trim(); ;
                }
                else
                {
                    return "";
                }

            }
            catch
            {
                return "";
            }
            finally
            {
                sqlConn.Close();
            }
        }

        public void SEARCHOUTPURSET()
        {
            StringBuilder sbSql = new StringBuilder();
            StringBuilder sbSqlQuery = new StringBuilder();
            SqlDataAdapter adapter1 = new SqlDataAdapter();
            SqlCommandBuilder sqlCmdBuilder1 = new SqlCommandBuilder();
            DataSet ds1 = new DataSet();

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
                                    SELECT 
                                    [MB001] AS '品號'
                                    ,[MB002] AS '品名'
                                    ,[PURMA001] AS '廠商'
                                    ,[PURMA002] AS '廠商名'
                                    ,[MC001] AS '庫別'
                                    ,[MC002] AS '庫別名'
                                    FROM [TKWAREHOUSE].[dbo].[OUTPURSET]
                                    ORDER BY [MB001]
                                    ");

                adapter1 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder1 = new SqlCommandBuilder(adapter1);
                sqlConn.Open();
                ds1.Clear();
                adapter1.Fill(ds1, "ds1");
                sqlConn.Close();


                if (ds1.Tables["ds1"].Rows.Count == 0)
                {
                    dataGridView5.DataSource = null;
                }
                else
                {
                    if (ds1.Tables["ds1"].Rows.Count >= 1)
                    {
                        //dataGridView1.Rows.Clear();
                        dataGridView5.DataSource = ds1.Tables["ds1"];
                        dataGridView5.AutoResizeColumns();
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

        public void ADDOUTPURSET(string MB001,string MB002,string PURMA001,string PURMA002,string MC001,string MC002)
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


                sbSql.AppendFormat(@" 
                                    INSERT INTO [TKWAREHOUSE].[dbo].[OUTPURSET]
                                    ([MB001],[MB002],[PURMA001],[PURMA002],[MC001],[MC002])
                                    VALUES
                                    ('{0}','{1}','{2}','{3}','{4}','{5}')

                                    ", MB001,MB002,PURMA001,PURMA002,MC001,MC002);


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

                    MessageBox.Show("完成");
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

        public void DELETEOUTPURSET(string MB001,string PURMA001, string MC002)
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


                sbSql.AppendFormat(@" 
                                    DELETE [TKWAREHOUSE].[dbo].[OUTPURSET]
                                    WHERE [MB001]='{0}' AND [PURMA001]='{1}' AND [MC001]='{2}'

                                    ", MB001, PURMA001, MC002);


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

                    MessageBox.Show("完成");
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

        private void dataGridView5_SelectionChanged(object sender, EventArgs e)
        {
            textBox4.Text = null;
            textBox5.Text = null;
            textBox6.Text = null;
            textBox7.Text = null;
            textBox8.Text = null;
            textBox9.Text = null;

            if (dataGridView5.CurrentRow != null)
            {
                int rowindex = dataGridView5.CurrentRow.Index;
                if (rowindex >= 0)
                {
                    DataGridViewRow row = dataGridView5.Rows[rowindex];
                    textBox4.Text = row.Cells["品號"].Value.ToString();
                    textBox5.Text = row.Cells["品名"].Value.ToString();
                    textBox6.Text = row.Cells["廠商"].Value.ToString();
                    textBox7.Text = row.Cells["廠商名"].Value.ToString();
                    textBox8.Text = row.Cells["庫別"].Value.ToString();
                    textBox9.Text = row.Cells["庫別名"].Value.ToString();

                }
                else
                {
                    

                }
            }
        }

        public string GETMAXTA002(string TA001)
        {
            DateTime dt = dateTimePicker2.Value;
            string TA002;

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


                StringBuilder sbSql = new StringBuilder();
                sbSql.Clear();
                sbSqlQuery.Clear();
                ds4.Clear();

                sbSql.AppendFormat(@"  SELECT ISNULL(MAX(TA002),'00000000000') AS TA002");
                sbSql.AppendFormat(@"  FROM [TK].[dbo].[MOCTA] ");
                //sbSql.AppendFormat(@"  WHERE  TC001='{0}' AND TC003='{1}'", "A542","20170119");
                sbSql.AppendFormat(@"  WHERE  TA001='{0}' AND TA003='{1}'", TA001, dt.ToString("yyyyMMdd"));
                sbSql.AppendFormat(@"  ");
                sbSql.AppendFormat(@"  ");

                adapter4 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder4 = new SqlCommandBuilder(adapter4);
                sqlConn.Open();
                ds4.Clear();
                adapter4.Fill(ds4, "TEMPds4");
                sqlConn.Close();


                if (ds4.Tables["TEMPds4"].Rows.Count == 0)
                {
                    return null;
                }
                else
                {
                    if (ds4.Tables["TEMPds4"].Rows.Count >= 1)
                    {
                        TA002 = SETTA002(ds4.Tables["TEMPds4"].Rows[0]["TA002"].ToString());
                        return TA002;

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
        public string SETTA002(string TA002)
        {
            DateTime dt = dateTimePicker2.Value;

            try
            {
                if (TA002.Equals("00000000000"))
                {
                    return dt.ToString("yyyyMMdd") + "001";
                }

                else
                {
                    int serno = Convert.ToInt16(TA002.Substring(8, 3));
                    serno = serno + 1;
                    string temp = serno.ToString();
                    temp = temp.PadLeft(3, '0');
                    return dt.ToString("yyyyMMdd") + temp.ToString();
                }
            }
            catch
            {
                return null;
            }
            finally
            {

            }
            
            
        }
        public void ADDMOCTATB()
        {
            MOCTADATA MOCTA = new MOCTADATA();
            MOCTA = SETMOCTA();
            string MOCMB001 = null;
            decimal MOCTA004 = 0; ;
            //找出外倉，製令單身用此外倉代號
            string MAINMB001 = MB001;
            string MOCTB009 = SEARCHOUTPURSETMC001(textBoxID.Text.Trim(), MB001);
            //找出最新的加工計價
            DataTable DT_MOCTM_MOCTN = FIND_TK_MOCTM_MOCTN(MB001);
            if(DT_MOCTM_MOCTN!=null && DT_MOCTM_MOCTN.Rows.Count>=1)
            {
                MOCTA.TA022 = DT_MOCTM_MOCTN.Rows[0]["TN009"].ToString();
                MOCTA.TA023 = DT_MOCTM_MOCTN.Rows[0]["TN008"].ToString();
                MOCTA.TA032 = DT_MOCTM_MOCTN.Rows[0]["TM004"].ToString();
                MOCTA.TA042 = DT_MOCTM_MOCTN.Rows[0]["TM005"].ToString();
            }

            const int MaxLength = 100;
            
            MOCMB001 = MB001;
            MOCTA004 = BAR;
            MOCTA.TA026 = TA026A;
            MOCTA.TA027 = TA027A;
            MOCTA.TA028 = TA028A;
            //MOCTB009 = textBox77.Text;

            
            try
            {
                //check TA002=2,TA040=2
                //[TB004]的計算，如果領用倍數MB041=1且不是201開頭的箱子，就取整數、MB041=1且是201開頭的箱子，就4捨5入到整數、其他就取到小數第3位
                if (MOCTA.TA002.Substring(0, 1).Equals("2") && MOCTA.TA040.Substring(0, 1).Equals("2"))
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



                    // ,(CASE WHEN ISNULL([MC001],'')<>'' THEN [MC001] ELSE [INVMB].MB017 END ) [TB009],'1' [TB011],[INVMB].MB002 [TB012],[INVMB].MB003 [TB013],[COPPURBATCHUSED].TD004 [TB014],'N' [TB018],'0' [TB019],'0' [TB020],'2' [TB022],'0' [TB024]

                    sbSql.AppendFormat(@" 
                                        INSERT INTO [TK].[dbo].[MOCTA]
                                        ([COMPANY],[CREATOR],[USR_GROUP],[CREATE_DATE],[MODIFIER],[MODI_DATE],[FLAG],[CREATE_TIME],[MODI_TIME],[TRANS_TYPE]
                                        ,[TRANS_NAME],[sync_count],[DataGroup],[TA001],[TA002],[TA003],[TA004],[TA005],[TA006],[TA007]
                                        ,[TA009],[TA010],[TA011],[TA012],[TA013],[TA014],[TA015],[TA016],[TA017],[TA018]
                                        ,[TA019],[TA020],[TA021],[TA022],[TA024],[TA025],[TA029],[TA030],[TA031],[TA034],[TA035]
                                        ,[TA040],[TA041],[TA043],[TA044],[TA045],[TA046],[TA047],[TA049],[TA050],[TA200]
                                        ,[TA026],[TA027],[TA028],[TA032],[TA023],[TA042]
                                        )
                                        VALUES
                                        ('{0}','{1}','{2}','{3}','{4}','{5}','{6}','{7}','{8}','{9}'
                                        ,'{10}','{11}','{12}','{13}','{14}','{15}','{16}','{17}','{18}','{19}'
                                        ,'{20}','{21}','{22}','{23}','{24}','{25}','{26}','{27}','{28}','{29}'
                                        ,'{30}','{31}','{32}','{33}','{34}','{35}',N'{36}','{37}','{38}','{39}'
                                        ,'{40}','{41}','{42}','{43}','{44}','{45}','{46}','{47}','{48}','{49}'
                                        ,'{50}','{51}','{52}','{53}','{54}','{55}','{56}'
                                        )
                                        ", MOCTA.COMPANY, MOCTA.CREATOR, MOCTA.USR_GROUP, MOCTA.CREATE_DATE, MOCTA.MODIFIER, MOCTA.MODI_DATE, MOCTA.FLAG, MOCTA.CREATE_TIME, MOCTA.MODI_TIME, MOCTA.TRANS_TYPE
                                        , MOCTA.TRANS_NAME, MOCTA.sync_count, MOCTA.DataGroup, MOCTA.TA001, MOCTA.TA002, MOCTA.TA003, MOCTA.TA004, MOCTA.TA005, MOCTA.TA006, MOCTA.TA007
                                        , MOCTA.TA009, MOCTA.TA010, MOCTA.TA011, MOCTA.TA012, MOCTA.TA013, MOCTA.TA014, MOCTA.TA015, MOCTA.TA016, MOCTA.TA017, MOCTA.TA018
                                        , MOCTA.TA019, MOCTA.TA020, MOCTA.TA021, MOCTA.TA022, MOCTA.TA024, MOCTA.TA025, MOCTA.TA029, MOCTA.TA030, MOCTA.TA031, MOCTA.TA034
                                        , MOCTA.TA035, MOCTA.TA040, MOCTA.TA041, MOCTA.TA043, MOCTA.TA044, MOCTA.TA045, MOCTA.TA046, MOCTA.TA047, MOCTA.TA049, MOCTA.TA050
                                        , MOCTA.TA200, MOCTA.TA026, MOCTA.TA027, MOCTA.TA028, MOCTA.TA032, MOCTA.TA023, MOCTA.TA042);

                    sbSql.AppendFormat(@" 
                                        INSERT INTO [TK].dbo.[MOCTB]
                                        ([COMPANY],[CREATOR],[USR_GROUP],[CREATE_DATE],[MODIFIER],[MODI_DATE],[FLAG],[CREATE_TIME],[MODI_TIME],[TRANS_TYPE]
                                        ,[TRANS_NAME],[sync_count],[DataGroup],[TB001],[TB002],[TB003],[TB004],[TB005],[TB006],[TB007]
                                        ,[TB009],[TB011],[TB012],[TB013],[TB014],[TB018],[TB019],[TB020],[TB022],[TB024]
                                        ,[TB025],[TB026],[TB027],[TB029],[TB030],[TB031],[TB501],[TB554],[TB556],[TB560])  

                                        (
                                        SELECT 
                                        '{1}' [COMPANY],'{2}' [CREATOR],'{3}' [USR_GROUP],'{4}' [CREATE_DATE],'{5}' [MODIFIER],'{6}' [MODI_DATE],'{7}' [FLAG],'{8}' [CREATE_TIME],'{9}' [MODI_TIME],'{10}' [TRANS_TYPE]
                                        ,'{11}' [TRANS_NAME],{12} [sync_count],'{13}' [DataGroup],'{14}' [TB001],'{15}' [TB002]
                                        ,[COPPURBATCHUSED].MB001 [TB003],[COPPURBATCHUSED].[NUM]   [TB004],0 [TB005],'****' [TB006],[INVMB].MB004  [TB007]
                                        ,'{16}' [TB009],'1' [TB011],[INVMB].MB002 [TB012],[INVMB].MB003 [TB013],[COPPURBATCHUSED].TD004 [TB014],'N' [TB018],'0' [TB019],'0' [TB020],'2' [TB022],'0' [TB024]
                                        ,'****' [TB025],'0' [TB026],'1' [TB027],'0' [TB029],'0' [TB030],'0' [TB031],'0' [TB501],'N' [TB554],'0' [TB556],'0' [TB560]
                                        FROM [TKWAREHOUSE].[dbo].[COPPURBATCHUSED],[TK].dbo.[INVMB]
                                        LEFT JOIN [TKWAREHOUSE].[dbo].[OUTPURSET] ON LTRIM(RTRIM([OUTPURSET].[MB001]))=LTRIM(RTRIM([INVMB].[MB001]))
                                        WHERE [COPPURBATCHUSED].[MB001]=[INVMB].MB001
                                        AND [COPPURBATCHUSED].[ID]='{0}'
                                        )
                                        
                                        ", textBoxID.Text.Trim()
                                        , MOCTA.COMPANY, MOCTA.CREATOR, MOCTA.USR_GROUP, MOCTA.CREATE_DATE, MOCTA.MODIFIER, MOCTA.MODI_DATE, MOCTA.FLAG, MOCTA.CREATE_TIME, MOCTA.MODI_TIME, MOCTA.TRANS_TYPE
                                        , MOCTA.TRANS_NAME, MOCTA.sync_count, MOCTA.DataGroup, MOCTA.TA001, MOCTA.TA002
                                        , MOCTB009
                                        );



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


            }
            catch
            {

            }

            finally
            {
                sqlConn.Close();
            }
        }

        //找最新的加工計價
        public DataTable FIND_TK_MOCTM_MOCTN(string TN004)
        {
            StringBuilder sbSql = new StringBuilder();
            StringBuilder sbSqlQuery = new StringBuilder();
            SqlDataAdapter adapter1 = new SqlDataAdapter();
            SqlCommandBuilder sqlCmdBuilder1 = new SqlCommandBuilder();
            DataSet ds = new DataSet();



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
                                    SELECT  
                                    TOP 1 TN004,TM004,TM005,TN008,TN009
                                    FROM [TK].dbo.MOCTM,[TK].dbo.MOCTN
                                    WHERE TM001=TN001 AND TM002=TN002
                                    AND TM009='Y'
                                    AND TN004='{0}'
                                    ORDER BY TM002 DESC
                                    ", TN004);

                adapter1 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder1 = new SqlCommandBuilder(adapter1);
                sqlConn.Open();
                ds.Clear();
                adapter1.Fill(ds, "ds");
                sqlConn.Close();


                if (ds.Tables["ds"].Rows.Count >= 1)
                {
                    return ds.Tables["ds"];

                }
                else
                {
                    return null;
                }

            }
            catch
            {
                return null;
            }
            finally
            {

            }
        }

        public MOCTADATA SETMOCTA()
        {
            DateTime dt = dateTimePicker2.Value;

            DataTable OUTPURSET = SEARCHOUTPURSET(MB001);
            DataTable BOMMC= SEARCHBOMMC(MB001);
            DataTable DTINVMB = SEARCHINVMB(MB001);


            MOCTADATA MOCTA = new MOCTADATA();
            MOCTA.COMPANY = "TK";
            MOCTA.CREATOR = "140020";
            MOCTA.USR_GROUP = "103000";
            //MOCTA.CREATE_DATE = dt1.ToString("yyyyMMdd");
            MOCTA.CREATE_DATE = DateTime.Now.ToString("yyyyMMdd");
            MOCTA.MODIFIER = "140020";
            MOCTA.MODI_DATE = dt.ToString("yyyyMMdd");
            MOCTA.FLAG = "0";
            MOCTA.CREATE_TIME = DateTime.Now.ToString("HH:mm:dd");
            MOCTA.MODI_TIME = DateTime.Now.ToString("HH:mm:dd");
            MOCTA.TRANS_TYPE = "P001";
            MOCTA.TRANS_NAME = "MOCMI02";
            MOCTA.sync_count = "0";
            MOCTA.DataGroup = "103000";
            MOCTA.TA001 = "A512";
            MOCTA.TA002 = GETMAXTA002(MOCTA.TA001);
            MOCTA.TA003 = dt.ToString("yyyyMMdd");
            MOCTA.TA004 = dt.ToString("yyyyMMdd");
            MOCTA.TA005 = BOMMC.Rows[0]["MC009"].ToString();
            MOCTA.TA006 = MB001;
            MOCTA.TA007 = BOMMC.Rows[0]["MB004"].ToString();
            MOCTA.TA009 = dt.ToString("yyyyMMdd");
            MOCTA.TA010 = dt.ToString("yyyyMMdd");
            MOCTA.TA011 = "1";
            MOCTA.TA012 = dt.ToString("yyyyMMdd");
            MOCTA.TA013 = "N";
            //MOCTA.TA014 = dt1.ToString("yyyyMMdd");
            MOCTA.TA014 = "";
            //MOCTA.TA015 = (BAR * BOMBAR).ToString();
            MOCTA.TA015 = SUM1.ToString();
            MOCTA.TA016 = "0";
            MOCTA.TA017 = "0";
            MOCTA.TA018 = "0";
            MOCTA.TA019 = "20";
            MOCTA.TA020 = DTINVMB.Rows[0]["MB017"].ToString();
            MOCTA.TA021 = "";
            MOCTA.TA022 = "0";
            MOCTA.TA023 = BOMMC.Rows[0]["MC002"].ToString();
            MOCTA.TA024 = "A512";
            MOCTA.TA025 = MOCTA.TA002;
            MOCTA.TA029 = TC015TD020;
            MOCTA.TA030 = "2"; //1=廠內、2=託外
            MOCTA.TA031 = "0";
            MOCTA.TA032 = OUTPURSET.Rows[0]["PURMA001"].ToString();
            MOCTA.TA034 = MB002;
            MOCTA.TA035 = MB003;
            MOCTA.TA040 = dt.ToString("yyyyMMdd");
            MOCTA.TA041 = "";
            MOCTA.TA042 = "NTD";
            MOCTA.TA043 = "1";
            MOCTA.TA044 = "N";
            MOCTA.TA045 = "0";
            MOCTA.TA046 = "0";
            MOCTA.TA047 = "0";
            MOCTA.TA049 = "N";
            MOCTA.TA050 = "0";
            MOCTA.TA200 = "1";


            return MOCTA;

            

        }

        public DataTable SEARCHBOMMC(string MB001)
        {
            StringBuilder sbSql = new StringBuilder();
            StringBuilder sbSqlQuery = new StringBuilder();
            SqlDataAdapter adapter1 = new SqlDataAdapter();
            SqlCommandBuilder sqlCmdBuilder1 = new SqlCommandBuilder();
            DataSet dsBOMMC = new DataSet();

         

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
                                    SELECT 
                                    [MC001],[MC002],[MC003],[MC004],[MC005],[MC006],[MC007],[MC008],[MC009],[MC010],[MC011],[MC012],[MC013],[MC014],[MC015],[MC016],[MC017],[MC018],[MC019],[MC020],[MC021],[MC022],[MC023],[MC024],[MC025],[MC026],[MC027]
                                    ,INVMB.MB004
                                    FROM [TK].[dbo].[BOMMC]
                                    LEFT JOIN [TK].dbo.[INVMB] ON MB001=MC001
                                    WHERE  [MC001]='{0}'
                                    ", MB001);

                adapter1 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder1 = new SqlCommandBuilder(adapter1);
                sqlConn.Open();
                dsBOMMC.Clear();
                adapter1.Fill(dsBOMMC, "dsBOMMC");
                sqlConn.Close();


                if (dsBOMMC.Tables["dsBOMMC"].Rows.Count >= 1)
                {
                    return dsBOMMC.Tables["dsBOMMC"];
                }
                else
                {
                    return null;
                }

            }
            catch
            {
                return null;
            }
            finally
            {

            }


        }

        public DataTable SEARCHOUTPURSET(string MB001)
        {
            StringBuilder sbSql = new StringBuilder();
            StringBuilder sbSqlQuery = new StringBuilder();
            SqlDataAdapter adapter1 = new SqlDataAdapter();
            SqlCommandBuilder sqlCmdBuilder1 = new SqlCommandBuilder();
            DataSet dsBOMMC = new DataSet();

           

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
                                    SELECT 
                                    [MB001]
                                    ,[MB002]
                                    ,[PURMA001]
                                    ,[PURMA002]
                                    ,[MC001]
                                    ,[MC002]
                                    FROM [TKWAREHOUSE].[dbo].[OUTPURSET]
                                    WHERE [MB001]='{0}'
                                    ", MB001);

                adapter1 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder1 = new SqlCommandBuilder(adapter1);
                sqlConn.Open();
                dsBOMMC.Clear();
                adapter1.Fill(dsBOMMC, "dsBOMMC");
                sqlConn.Close();


                if (dsBOMMC.Tables["dsBOMMC"].Rows.Count >= 1)
                {
                    return dsBOMMC.Tables["dsBOMMC"];

                }
                else
                {
                    return null;
                }

            }
            catch
            {
                return null;
            }
            finally
            {

            }


        }

        public DataTable SEARCHINVMB(string MB001)
        {
            StringBuilder sbSql = new StringBuilder();
            StringBuilder sbSqlQuery = new StringBuilder();
            SqlDataAdapter adapter1 = new SqlDataAdapter();
            SqlCommandBuilder sqlCmdBuilder1 = new SqlCommandBuilder();
            DataSet ds = new DataSet();



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
                                    SELECT 
                                    *
                                    FROM [TK].dbo.INVMB
                                    WHERE MB001='{0}'
                                    ", MB001);

                adapter1 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder1 = new SqlCommandBuilder(adapter1);
                sqlConn.Open();
                ds.Clear();
                adapter1.Fill(ds, "ds");
                sqlConn.Close();


                if (ds.Tables["ds"].Rows.Count >= 1)
                {
                    return ds.Tables["ds"];

                }
                else
                {
                    return null;
                }

            }
            catch
            {
                return null;
            }
            finally
            {

            }


        }


        private void dataGridView4_SelectionChanged(object sender, EventArgs e)
        {
            //SELECT [TA001] AS '單別',[TA002] AS '單號',[ID] AS '批號'
            DELCOPPURBATCHPUR_ID = "";
            DELCOPPURBATCHPUR_TA001 = "";
            DELCOPPURBATCHPUR_TA002 = "";

            if(dataGridView4.CurrentRow != null)
            {
                int rowindex = dataGridView4.CurrentRow.Index;
                if (rowindex >= 0)
                {
                    DataGridViewRow row = dataGridView4.Rows[rowindex];
                    DELCOPPURBATCHPUR_ID = row.Cells["批號"].Value.ToString();
                    DELCOPPURBATCHPUR_TA001 = row.Cells["單別"].Value.ToString();
                    DELCOPPURBATCHPUR_TA002 = row.Cells["單號"].Value.ToString();
                   

                }
                else
                {
                    DELCOPPURBATCHPUR_ID = "";
                    DELCOPPURBATCHPUR_TA001 = "";
                    DELCOPPURBATCHPUR_TA002 = "";

                }
            }
        }

        public string SEARCHOUTPURSETMC001(string ID,string MAINMB001)
        {
            StringBuilder sbSql = new StringBuilder();
            StringBuilder sbSqlQuery = new StringBuilder();
            SqlDataAdapter adapter1 = new SqlDataAdapter();
            SqlCommandBuilder sqlCmdBuilder1 = new SqlCommandBuilder();
            DataSet ds = new DataSet();



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
                                     
                                  SELECT (CASE WHEN ISNULL(MC001,'')<>'' THEN MC001 ELSE MC001B END ) AS MC001
                                    FROM (
                                    SELECT DISTINCT MC001 MC001
                                    ,(
                                    SELECT TOP 1 MC001
                                    FROM [TKWAREHOUSE].[dbo].[COPPURBATCHUSED]
                                    LEFT JOIN [TKWAREHOUSE].[dbo].[OUTPURSET] ON (LTRIM(RTRIM([OUTPURSET].[MB001]))='{1}') 
                                    WHERE  [COPPURBATCHUSED].[ID]='{0}'
                                    AND ISNULL(MC001,'')<>''
                                    ORDER BY MC001
                                    ) MC001B
                                    FROM [TKWAREHOUSE].[dbo].[COPPURBATCHUSED],[TK].dbo.[INVMB]
                                    LEFT JOIN [TKWAREHOUSE].[dbo].[OUTPURSET] ON LTRIM(RTRIM([OUTPURSET].[MB001]))=LTRIM(RTRIM([INVMB].[MB001]))
                                    WHERE [COPPURBATCHUSED].[MB001]=[INVMB].MB001
                                    AND [COPPURBATCHUSED].[ID]='{0}'
                                   
                                    ) AS TEMP


                                    ", ID, MAINMB001);

                adapter1 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder1 = new SqlCommandBuilder(adapter1);
                sqlConn.Open();
                ds.Clear();
                adapter1.Fill(ds, "ds");
                sqlConn.Close();


                if (ds.Tables["ds"].Rows.Count >= 1)
                {
                    return ds.Tables["ds"].Rows[0]["MC001"].ToString().Trim();

                }
                else
                {
                    return null;
                }

            }
            catch
            {
                return null;
            }
            finally
            {

            }


        }

        public void SEARCH_COPTE_COPTF(string SDATES, string EDATES)
        {   
            // 使用 try-catch 區塊來處理連線和查詢錯誤
            try
            {
                // --- 1. 資料庫連線字串解密與建立 ---
                // 20210902密：使用 new Class1() 實體進行解密
                Class1 TKID = new Class1();

                // 從配置檔讀取連線字串並使用 SqlConnectionStringBuilder 處理
                SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(
                    ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString
                );

                // 資料庫使用者密碼解密
                sqlsb.Password = TKID.Decryption(sqlsb.Password);
                sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

                StringBuilder sqlQuery = new StringBuilder();
                sqlQuery.AppendFormat(@"
                                        SELECT 
                                        TE001 AS '單別'
                                        ,TE002 AS '單號'
                                        ,TE003 AS '變更版次'
                                        ,TE004 AS '變更日期'
                                        ,TE006 AS '變更原因'
                                        ,TF005 AS '品號'
                                        ,TF006 AS '品名'
                                        ,TF009 AS '數量'
                                        ,TF010 AS '單位'

                                        FROM [TK].dbo.COPTE,[TK].dbo.COPTF
                                        WHERE TE001=TF001 AND TE002=TF002
                                        AND TE004>='{0}' AND TE004<='{1}'
                                        ORDER BY TE001,TE002,TE003

                                        ", SDATES,EDATES);


                // 使用 using 確保 SqlConnection 在完成後或發生錯誤時會自動關閉和釋放
                using (SqlConnection sqlConn = new SqlConnection(sqlsb.ConnectionString))
                {
                    // --- 2. 執行查詢並填充 DataSet ---
                    // 移除不必要的 SqlCommandBuilder 和多餘的 StringBuilder 宣告/清除
                    using (SqlDataAdapter adapter = new SqlDataAdapter(sqlQuery.ToString(), sqlConn))
                    using (DataSet ds = new DataSet())
                    {
                        // Fill 方法會自動開啟連線，完成後自動關閉（如果連線最初是關閉的）
                        adapter.Fill(ds, "ds1");

                        // --- 3. 資料繫結邏輯 ---
                        if (ds.Tables.Count > 0 && ds.Tables["ds1"].Rows.Count > 0)
                        {
                            dataGridView6.DataSource = ds.Tables["ds1"];
                            dataGridView6.AutoResizeColumns();
                        }
                        else
                        {
                            // 查詢結果為空
                            dataGridView6.DataSource = null;
                        }
                    } // adapter 和 ds 會在這裡被 Dispose
                } // sqlConn 會在這裡被 Dispose 和 Close
            }
            catch (Exception ex)
            {
                // ❌ 重要的優化：避免使用空的 catch 區塊。
                // 應該記錄錯誤或提示使用者。
                System.Windows.Forms.MessageBox.Show("資料查詢失敗，請檢查配置或連線。\n錯誤訊息: " + ex.Message);

                // 發生錯誤時清空資料顯示
                if (dataGridView6 != null)
                {
                    dataGridView6.DataSource = null;
                }
            }
        }

        public void SEARCH_COPTC_COPTD(string SDATES, string EDATES)
        {
            // 使用 try-catch 區塊來處理連線和查詢錯誤
            try
            {
                // --- 1. 資料庫連線字串解密與建立 ---
                // 20210902密：使用 new Class1() 實體進行解密
                Class1 TKID = new Class1();

                // 從配置檔讀取連線字串並使用 SqlConnectionStringBuilder 處理
                SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(
                    ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString
                );

                // 資料庫使用者密碼解密
                sqlsb.Password = TKID.Decryption(sqlsb.Password);
                sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

                StringBuilder sqlQuery = new StringBuilder();
                sqlQuery.AppendFormat(@"
                                       SELECT 
                                        TC001 AS '單別'
                                        ,TC002 AS '單號'
                                        ,TC003 AS '日期'
                                        ,TD003 AS '序號'
                                        ,TD004 AS '品號'
                                        ,TD005 AS '品名'
                                        ,TD008 AS '數量'
                                        ,TD010 AS '單位'
                                        ,TD013 AS '預交日'

                                        FROM [TK].dbo.COPTC,[TK].dbo.COPTD
                                        WHERE TC001=TD001 AND TC002=TD002
                                        AND TC003>='{0}' AND TC003<='{1}'
                                        AND TD004 LIKE '5%'
                                        AND TD004 NOT LIKE '599%'
                                        ORDER BY TC001,TC002,TD003

                                        ", SDATES, EDATES);


                // 使用 using 確保 SqlConnection 在完成後或發生錯誤時會自動關閉和釋放
                using (SqlConnection sqlConn = new SqlConnection(sqlsb.ConnectionString))
                {
                    // --- 2. 執行查詢並填充 DataSet ---
                    // 移除不必要的 SqlCommandBuilder 和多餘的 StringBuilder 宣告/清除
                    using (SqlDataAdapter adapter = new SqlDataAdapter(sqlQuery.ToString(), sqlConn))
                    using (DataSet ds = new DataSet())
                    {
                        // Fill 方法會自動開啟連線，完成後自動關閉（如果連線最初是關閉的）
                        adapter.Fill(ds, "ds1");

                        // --- 3. 資料繫結邏輯 ---
                        if (ds.Tables.Count > 0 && ds.Tables["ds1"].Rows.Count > 0)
                        {
                            dataGridView7.DataSource = ds.Tables["ds1"];
                            dataGridView7.AutoResizeColumns();
                        }
                        else
                        {
                            // 查詢結果為空
                            dataGridView7.DataSource = null;
                        }
                    } // adapter 和 ds 會在這裡被 Dispose
                } // sqlConn 會在這裡被 Dispose 和 Close
            }
            catch (Exception ex)
            {
                // ❌ 重要的優化：避免使用空的 catch 區塊。
                // 應該記錄錯誤或提示使用者。
                System.Windows.Forms.MessageBox.Show("資料查詢失敗，請檢查配置或連線。\n錯誤訊息: " + ex.Message);

                // 發生錯誤時清空資料顯示
                if (dataGridView7 != null)
                {
                    dataGridView7.DataSource = null;
                }
            }
        }


        private void dataGridView7_SelectionChanged(object sender, EventArgs e)
        {
            if (dataGridView7.CurrentRow != null)
            {
                int rowindex = dataGridView7.CurrentRow.Index;
                if (rowindex >= 0)
                {
                    DataGridViewRow row = dataGridView7.Rows[rowindex];
                    textBox10.Text = row.Cells["單別"].Value.ToString();
                    textBox11.Text = row.Cells["單號"].Value.ToString();

                    SEARCH_PURTA_PURTB(textBox10.Text, textBox11.Text);
                }
                else
                {
                    textBox10.Text = null;
                    textBox11.Text = null;

                }
            }
        }

        public void SEARCH_PURTA_PURTB(string TC001, string TC002)
        {
            // 使用 try-catch 區塊來處理連線和查詢錯誤
            try
            {
                // --- 1. 資料庫連線字串解密與建立 ---
                // 20210902密：使用 new Class1() 實體進行解密
                Class1 TKID = new Class1();

                // 從配置檔讀取連線字串並使用 SqlConnectionStringBuilder 處理
                SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(
                    ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString
                );

                // 資料庫使用者密碼解密
                sqlsb.Password = TKID.Decryption(sqlsb.Password);
                sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

                StringBuilder sqlQuery = new StringBuilder();

                string TA005 = TC001.Trim() + TC002.Trim();
                sqlQuery.AppendFormat(@"
                                        SELECT 
                                        TA001 AS '請購單別'
                                        ,TA002 AS '請購單號'
                                        ,TA005 AS '訂單'

                                        FROM [TK].dbo.PURTA
                                        WHERE TA005='{0}'

                                        ", TA005);


                // 使用 using 確保 SqlConnection 在完成後或發生錯誤時會自動關閉和釋放
                using (SqlConnection sqlConn = new SqlConnection(sqlsb.ConnectionString))
                {
                    // --- 2. 執行查詢並填充 DataSet ---
                    // 移除不必要的 SqlCommandBuilder 和多餘的 StringBuilder 宣告/清除
                    using (SqlDataAdapter adapter = new SqlDataAdapter(sqlQuery.ToString(), sqlConn))
                    using (DataSet ds = new DataSet())
                    {
                        // Fill 方法會自動開啟連線，完成後自動關閉（如果連線最初是關閉的）
                        adapter.Fill(ds, "ds1");

                        // --- 3. 資料繫結邏輯 ---
                        if (ds.Tables.Count > 0 && ds.Tables["ds1"].Rows.Count > 0)
                        {
                            dataGridView8.DataSource = ds.Tables["ds1"];
                            dataGridView8.AutoResizeColumns();
                        }
                        else
                        {
                            // 查詢結果為空
                            dataGridView8.DataSource = null;
                        }
                    } // adapter 和 ds 會在這裡被 Dispose
                } // sqlConn 會在這裡被 Dispose 和 Close
            }
            catch (Exception ex)
            {
                // ❌ 重要的優化：避免使用空的 catch 區塊。
                // 應該記錄錯誤或提示使用者。
                System.Windows.Forms.MessageBox.Show("資料查詢失敗，請檢查配置或連線。\n錯誤訊息: " + ex.Message);

                // 發生錯誤時清空資料顯示
                if (dataGridView8 != null)
                {
                    dataGridView8.DataSource = null;
                }
            }
        }
        public void ADDMOCTAB_BY_COPTC_COPTD(string COPTC_TC001, string COPTC_TC002, string PURTA_TA001, string PURTA_TA002, string PURTA_TA003)
        {
            try
            {
                PURTA PURTA = new PURTA();     
                PURTA = SETPURTA();
              
                PURTA.TA001 = PURTA_TA001;
                PURTA.TA002 = PURTA_TA002;
                PURTA.TA003 = PURTA_TA003;
                PURTA.TA005 = COPTC_TC001.Trim() + COPTC_TC002.Trim();
                PURTA.TA013 = PURTA_TA003;
                PURTA.UDF01 = "Y";
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

                sbSql.AppendFormat(@" 
                                    INSERT INTO [TK].[dbo].[PURTA]
                                    (
                                        [COMPANY],[CREATOR],[USR_GROUP],[CREATE_DATE],[MODIFIER],
                                        [MODI_DATE],[FLAG],[CREATE_TIME],[MODI_TIME],[TRANS_TYPE],
                                        [TRANS_NAME],[sync_date],[sync_time],[sync_mark],[sync_count],
                                        [DataUser],[DataGroup],
                                        [TA001],[TA002],[TA003],[TA004],[TA005],
                                        [TA006],[TA007],[TA008],[TA009],[TA010],
                                        [TA011],[TA012],[TA013],[TA014],[TA015],
                                        [TA016],[TA017],[TA018],[TA019],[TA020],
                                        [TA021],[TA022],[TA023],[TA024],[TA025],
                                        [TA026],[TA027],[TA028],[TA029],[TA030],
                                        [TA031],[TA032],[TA033],[TA034],[TA035],
                                        [TA036],[TA037],[TA038],[TA039],[TA040],
                                        [TA041],[TA042],[TA043],[TA044],[TA045],
                                        [TA046],[UDF01],[UDF02],[UDF03],[UDF04],
                                        [UDF05],[UDF06],[UDF07],[UDF08],[UDF09],
                                        [UDF10]
                                    )
                                    VALUES 
                                    (
                                        '{0}','{1}','{2}','{3}','{4}',
                                        '{5}','{6}','{7}','{8}','{9}',
                                        '{10}','{11}','{12}','{13}','{14}',
                                        '{15}','{16}',
                                        '{17}','{18}','{19}','{20}','{21}',
                                        '{22}','{23}','{24}','{25}','{26}',
                                        '{27}','{28}','{29}','{30}','{31}',
                                        '{32}','{33}','{34}','{35}','{36}',
                                        '{37}','{38}','{39}','{40}','{41}',
                                        '{42}','{43}','{44}','{45}','{46}',
                                        '{47}','{48}','{49}','{50}','{51}',
                                        '{52}','{53}','{54}','{55}','{56}',
                                        '{57}','{58}','{59}','{60}','{61}',
                                        '{62}','{63}','{64}','{65}','{66}',
                                        '{67}','{68}','{69}','{70}','{71}',
                                        '{72}'
                                    )",
                                    PURTA.COMPANY, PURTA.CREATOR, PURTA.USR_GROUP, PURTA.CREATE_DATE, PURTA.MODIFIER,
                                    PURTA.MODI_DATE, PURTA.FLAG, PURTA.CREATE_TIME, PURTA.MODI_TIME, PURTA.TRANS_TYPE,
                                    PURTA.TRANS_NAME, PURTA.sync_date, PURTA.sync_time, PURTA.sync_mark, PURTA.sync_count,
                                    PURTA.DataUser, PURTA.DataGroup,

                                    // 索引 17 - 62: TA 欄位 (46 個)
                                    PURTA.TA001, PURTA.TA002, PURTA.TA003, PURTA.TA004, PURTA.TA005,
                                    PURTA.TA006, PURTA.TA007, PURTA.TA008, PURTA.TA009, PURTA.TA010,
                                    PURTA.TA011, PURTA.TA012, PURTA.TA013, PURTA.TA014, PURTA.TA015,
                                    PURTA.TA016, PURTA.TA017, PURTA.TA018, PURTA.TA019, PURTA.TA020,
                                    PURTA.TA021, PURTA.TA022, PURTA.TA023, PURTA.TA024, PURTA.TA025,
                                    PURTA.TA026, PURTA.TA027, PURTA.TA028, PURTA.TA029, PURTA.TA030,
                                    PURTA.TA031, PURTA.TA032, PURTA.TA033, PURTA.TA034, PURTA.TA035,
                                    PURTA.TA036, PURTA.TA037, PURTA.TA038, PURTA.TA039, PURTA.TA040,
                                    PURTA.TA041, PURTA.TA042, PURTA.TA043, PURTA.TA044, PURTA.TA045,
                                    PURTA.TA046,

                                    // 索引 63 - 72: UDF 欄位 (10 個)
                                    PURTA.UDF01, PURTA.UDF02, PURTA.UDF03, PURTA.UDF04,
                                    PURTA.UDF05, PURTA.UDF06, PURTA.UDF07, PURTA.UDF08, PURTA.UDF09,
                                    PURTA.UDF10
                                    );
          
                sbSql.AppendFormat(@" 
                                     INSERT INTO [TK].[dbo].[PURTB]
                                    (
                                    [COMPANY],[CREATOR],[USR_GROUP],[CREATE_DATE],[MODIFIER]
                                    ,[MODI_DATE],[FLAG],[CREATE_TIME],[MODI_TIME],[TRANS_TYPE]
                                    ,[TRANS_NAME],[sync_date],[sync_time],[sync_mark],[sync_count]
                                    ,[DataUser],[DataGroup]
                                    ,[TB001],[TB002],[TB003],[TB004],[TB005]
                                    ,[TB006],[TB007],[TB008],[TB009],[TB010]
                                    ,[TB011],[TB012],[TB013],[TB014],[TB015]
                                    ,[TB016],[TB017],[TB018],[TB019],[TB020]
                                    ,[TB021],[TB022],[TB023],[TB024],[TB025]
                                    ,[TB026],[TB027],[TB028],[TB029],[TB030]
                                    ,[TB031],[TB032],[TB033],[TB034],[TB035]
                                    ,[TB036],[TB037],[TB038],[TB039],[TB040]
                                    ,[TB041],[TB042],[TB043],[TB044],[TB045]
                                    ,[TB046],[TB047],[TB048],[TB049],[TB050]
                                    ,[TB051],[TB052],[TB053],[TB054],[TB055]
                                    ,[TB056],[TB057],[TB058],[TB059],[TB060]
                                    ,[TB061],[TB062],[TB063],[TB064],[TB065]
                                    ,[TB066],[TB067],[TB068],[TB069],[TB070]
                                    ,[TB071],[TB072],[TB073],[TB074],[TB075]
                                    ,[TB076],[TB077],[TB078],[TB079],[TB080]
                                    ,[TB081],[TB082],[TB083],[TB084],[TB085]
                                    ,[TB086],[TB087],[TB088],[TB089],[TB090]
                                    ,[TB091],[TB092],[TB093],[TB094],[TB095]
                                    ,[TB096],[TB097],[TB098],[TB099],[UDF01]
                                    ,[UDF02],[UDF03],[UDF04],[UDF05],[UDF06]
                                    ,[UDF07],[UDF08],[UDF09],[UDF10]
                                    )

                                    SELECT 
                                    'TK'[COMPANY],'120025' [CREATOR],'103400' [USR_GROUP],CONVERT(NVARCHAR,GETDATE(),112) [CREATE_DATE],'160115' [MODIFIER]
                                    ,CONVERT(NVARCHAR,GETDATE(),112) [MODI_DATE],0 [FLAG],CONVERT(NVARCHAR,GETDATE(),108) [CREATE_TIME],CONVERT(NVARCHAR,GETDATE(),108) [MODI_TIME],'P001' [TRANS_TYPE]
                                    ,'PURI05' [TRANS_NAME],'' [sync_date],'' [sync_time],'' [sync_mark],0 [sync_count]
                                    ,'' [DataUser],'103400' [DataGroup]
                                    ,'{0}' [TB001],'{1}' [TB002],RIGHT(REPLICATE('0', 4) + CONVERT(VARCHAR, ROW_NUMBER() OVER (ORDER BY TD003)), 4) [TB003],TD004 [TB004],TD005 [TB005]
                                    ,TD006 [TB006],TD010 [TB007],MB017 [TB008],(TD008+TD024) [TB009],MB032 [TB010]
                                    ,TD013 [TB011],'' [TB012],'' [TB013],0 [TB014],'' [TB015]
                                    ,'NTD' [TB016],0 [TB017],0 [TB018],'' [TB019],'N' [TB020]
                                    ,'N' [TB021],'' [TB022],'' [TB023],'' [TB024],'N' [TB025]
                                    ,(CASE WHEN MB100='Y' THEN '1' ELSE '2' END) [TB026],'' [TB027] ,'' [TB028],COPTD.TD001 [TB029],COPTD.TD002 [TB030]
                                    ,COPTD.TD003 [TB031],'N' [TB032],'0001' [TB033],0 [TB034],0 [TB035]
                                    ,'' [TB036],'' [TB037],'' [TB038],'N' [TB039],0 [TB040]
                                    ,0 [TB041],'' [TB042],'' [TB043],'' [TB044],'' [TB045]
                                    ,'' [TB046],'' [TB047],'' [TB048],0 [TB049],'' [TB050]
                                    ,0 [TB051],0 [TB052],0 [TB053],'' [TB054],'' [TB055]
                                    ,'' [TB056],'' [TB057],'1' [TB058],'' [TB059],'' [TB060]
                                    ,'' [TB061],'' [TB062],0 [TB063],'N' [TB064],'1' [TB065]
                                    ,'' [TB066],'2' [TB067],0 [TB068],0 [TB069],'' [TB070]
                                    ,'' [TB071],'' [TB072],'' [TB073],'' [TB074],0 [TB075]
                                    ,'' [TB076],0 [TB077],'' [TB078],'' [TB079],'' [TB080]
                                    ,0 [TB081],0 [TB082],0 [TB083],0 [TB084],0 [TB085]
                                    ,'' [TB086],'' [TB087],0 [TB088],'1' [TB089],0 [TB090]
                                    ,0 [TB091],0 [TB092],0 [TB093],'' [TB094],'' [TB095]
                                    ,'' [TB096],'' [TB097],'' [TB098],'' [TB099]
                                    ,'Y' [UDF01]
                                    ,'' [UDF02],'' [UDF03],'' [UDF04],'' [UDF05]
                                    ,0 [UDF06],0 [UDF07],0 [UDF08],0 [UDF09],0 [UDF10]
                                    FROM [TK].dbo.COPTC,[TK].dbo.COPTD,[TK].dbo.INVMB
                                    WHERE TC001=TD001 AND TC002=TD002
                                    AND TD004=MB001
                                    AND TC001='{2}' AND TC002='{3}'
                                    AND TD004 LIKE '5%'
                                    ", PURTA_TA001, PURTA_TA002, COPTC_TC001, COPTC_TC002);
    


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

                    UPDATEPURTA_BY_COPTC_COPTD(PURTA_TA001, PURTA_TA002);
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
        public void UPDATEPURTA_BY_COPTC_COPTD(string TA001,string TA002)
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
            //UPDATE TB039='N'
               
            sbSql.AppendFormat(@" 
                                UPDATE  [TK].dbo.PURTB SET TB039='N' WHERE ISNULL(TB039,'')=''
                                UPDATE  [TK].dbo.PURTA
                                SET TA011=(SELECT SUM(TB009) FROM [TK].dbo.PURTB WHERE PURTA.TA001=PURTB.TB001 AND  PURTA.TA002=PURTB.TB002)
                                WHERE TA001='{0}' AND TA002='{1}'
                                ", TA001, TA002);


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

        
        #endregion

        #region BUTTON
        private void button1_Click(object sender, EventArgs e)
        {
            SEARCHBTACHID();
            //SEARCHCOPPURBATCHCOPTD(ID);
        }
        private void button2_Click(object sender, EventArgs e)
        {
            ID = GETMAXID();
            ADDBTACHID(ID);
            SEARCHBTACHID();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            if(!string.IsNullOrEmpty(textBoxID.Text)&&!string.IsNullOrEmpty(textBox1.Text)&& !string.IsNullOrEmpty(textBox2.Text)&& !string.IsNullOrEmpty(textBox3.Text))
            {
                ADDCOPPURBATCHCOPTD(textBoxID.Text.Trim(), textBox1.Text.Trim(), textBox2.Text.Trim(), textBox3.Text.Trim());
            }

            SEARCHCOPPURBATCHCOPTD(textBoxID.Text.Trim());
        }
        private void button6_Click(object sender, EventArgs e)
        {
            DialogResult dialogResult = MessageBox.Show("要刪除了?", "要刪除了?", MessageBoxButtons.YesNo);
            if (dialogResult == DialogResult.Yes)
            {
                if (!string.IsNullOrEmpty(textBoxID.Text) && !string.IsNullOrEmpty(DELTD001) && !string.IsNullOrEmpty(DELTD002) && !string.IsNullOrEmpty(DELTD003))
                {
                    DELCOPPURBATCHCOPTD(textBoxID.Text.Trim(), DELTD001.Trim(), DELTD002.Trim(), DELTD003.Trim());
                }
            }

            SEARCHCOPPURBATCHCOPTD(textBoxID.Text.Trim());

        }
        private void button4_Click(object sender, EventArgs e)
        {
            ADDCOPPURBATCHUSED(textBoxID.Text.Trim());
            SEARCHCOPPURBATCHUSED(textBoxID.Text.Trim());
        }
        private void button5_Click(object sender, EventArgs e)
        {
            MOCTA001 = "A311";
            MOCTA003 = dateTimePicker1.Value.ToString("yyyyMMdd");
            MOCTA002 = GETMAXMOCTA002(MOCTA001, MOCTA003);

            ADDMOCTAB(textBoxID.Text.Trim(),"2");

            ADDCOPPURBATCHPUR(textBoxID.Text.Trim(), MOCTA001, MOCTA002);
            SEARCHCOPPURBATCHPUR(textBoxID.Text.Trim());

            MessageBox.Show("已完成請購單" + MOCTA001 + " " + MOCTA002);

        }

        private void button7_Click(object sender, EventArgs e)
        {
            MOCTA001 = "A311";
            MOCTA003 = dateTimePicker1.Value.ToString("yyyyMMdd");
            MOCTA002 = GETMAXMOCTA002(MOCTA001, MOCTA003);

            ADDMOCTAB(textBoxID.Text.Trim(),"3");

            ADDCOPPURBATCHPUR(textBoxID.Text.Trim(), MOCTA001, MOCTA002);
            SEARCHCOPPURBATCHPUR(textBoxID.Text.Trim());

            MessageBox.Show("已完成請購單" + MOCTA001 + " " + MOCTA002);
        }

        private void button8_Click(object sender, EventArgs e)
        {
            string TA001 = "A512";
            string TA002 = "";

            TA002 = GETMAXTA002(TA001);
     
            ADDMOCTATB();

            ADDCOPPURBATCHPUR(textBoxID.Text.Trim(), TA001, TA002);
            SEARCHCOPPURBATCHPUR(textBoxID.Text.Trim());

            MessageBox.Show("完成");


            //MOCTA001 = "A311";
            //MOCTA003 = dateTimePicker1.Value.ToString("yyyyMMdd");
            //MOCTA002 = GETMAXMOCTA002(MOCTA001);

            //ADDMOCTAB(textBoxID.Text.Trim(), "4");

            //ADDCOPPURBATCHPUR(textBoxID.Text.Trim(), MOCTA001, MOCTA002);
            //SEARCHCOPPURBATCHPUR(textBoxID.Text.Trim());

            //MessageBox.Show("已完成請購單" + MOCTA001 + " " + MOCTA002);
        }
        private void button10_Click(object sender, EventArgs e)
        {
            MOCTA001 = "A311";
            MOCTA003 = dateTimePicker1.Value.ToString("yyyyMMdd");
            MOCTA002 = GETMAXMOCTA002(MOCTA001, MOCTA003);

            ADDMOCTAB(textBoxID.Text.Trim(), "1");

            ADDCOPPURBATCHPUR(textBoxID.Text.Trim(), MOCTA001, MOCTA002);
            SEARCHCOPPURBATCHPUR(textBoxID.Text.Trim());

            MessageBox.Show("已完成請購單" + MOCTA001 + " " + MOCTA002);
        }

        private void button9_Click(object sender, EventArgs e)
        {
            MOCTA001 = "A311";
            MOCTA003 = dateTimePicker1.Value.ToString("yyyyMMdd");
            MOCTA002 = GETMAXMOCTA002(MOCTA001, MOCTA003);

            ADDMOCTAB2(textBoxID.Text.Trim());

            ADDCOPPURBATCHPUR(textBoxID.Text.Trim(), MOCTA001, MOCTA002);
            SEARCHCOPPURBATCHPUR(textBoxID.Text.Trim());

            MessageBox.Show("已完成請購單" + MOCTA001 + " " + MOCTA002);

        }
        private void button11_Click(object sender, EventArgs e)
        {
            MOCTA001 = "A311";
            MOCTA003 = dateTimePicker1.Value.ToString("yyyyMMdd");
            MOCTA002 = GETMAXMOCTA002(MOCTA001, MOCTA003);

            ADDMOCTAB3(textBoxID.Text.Trim(), "1", "2");

            ADDCOPPURBATCHPUR(textBoxID.Text.Trim(), MOCTA001, MOCTA002);
            SEARCHCOPPURBATCHPUR(textBoxID.Text.Trim());

            MessageBox.Show("已完成請購單" + MOCTA001 + " " + MOCTA002);
        }

        private void button12_Click(object sender, EventArgs e)
        {
            MOCTA001 = "A311";
            MOCTA003 = dateTimePicker1.Value.ToString("yyyyMMdd");
            MOCTA002 = GETMAXMOCTA002(MOCTA001, MOCTA003);

            ADDMOCTAB(textBoxID.Text.Trim(), "4");

            ADDCOPPURBATCHPUR(textBoxID.Text.Trim(), MOCTA001, MOCTA002);
            SEARCHCOPPURBATCHPUR(textBoxID.Text.Trim());

            MessageBox.Show("已完成請購單" + MOCTA001 + " " + MOCTA002);
        }

        private void button13_Click(object sender, EventArgs e)
        {
            MOCTA001 = "A311";
            MOCTA003 = dateTimePicker1.Value.ToString("yyyyMMdd");
            MOCTA002 = GETMAXMOCTA002(MOCTA001, MOCTA003);

            ADDPURTAB(textBoxID.Text.Trim(), "3");

            ADDCOPPURBATCHPUR(textBoxID.Text.Trim(), MOCTA001, MOCTA002);
            SEARCHCOPPURBATCHPUR(textBoxID.Text.Trim());

            MessageBox.Show("已完成外購的請購單" + MOCTA001 + " " + MOCTA002);
        }
        private void button14_Click(object sender, EventArgs e)
        {
            SEARCHOUTPURSET();
        }


        private void button15_Click(object sender, EventArgs e)
        {
            if(!string.IsNullOrEmpty(textBox5.Text)&& !string.IsNullOrEmpty(textBox7.Text)&& !string.IsNullOrEmpty(textBox9.Text))
            {
                ADDOUTPURSET(textBox4.Text.Trim(), textBox5.Text.Trim(), textBox6.Text.Trim(), textBox7.Text.Trim(), textBox8.Text.Trim(), textBox9.Text.Trim());

                SEARCHOUTPURSET();
            }

            
        }

        private void button16_Click(object sender, EventArgs e)
        {
            DialogResult dialogResult = MessageBox.Show("要刪除了?", "要刪除了?", MessageBoxButtons.YesNo);
            if (dialogResult == DialogResult.Yes)
            {
                if (!string.IsNullOrEmpty(textBox5.Text) && !string.IsNullOrEmpty(textBox7.Text) && !string.IsNullOrEmpty(textBox9.Text))
                {
                    DELETEOUTPURSET(textBox4.Text.Trim(), textBox6.Text.Trim(), textBox8.Text.Trim());

                    SEARCHOUTPURSET();
                }
            }
            else if (dialogResult == DialogResult.No)
            {
                //do something else
            }

            
                
        }

        private void button17_Click(object sender, EventArgs e)
        {
            DialogResult dialogResult = MessageBox.Show("要刪除了?", "要刪除了?", MessageBoxButtons.YesNo);
            if (dialogResult == DialogResult.Yes)
            {
                
                DELETECOPPURBATCHPUR(DELCOPPURBATCHPUR_ID, DELCOPPURBATCHPUR_TA001, DELCOPPURBATCHPUR_TA002);
                SEARCHCOPPURBATCHPUR(ID);
                
            }
            else if (dialogResult == DialogResult.No)
            {
                //do something else
            }
           
        }
        private void button18_Click(object sender, EventArgs e)
        {

            MOCTA001 = "A311";
            MOCTA003 = dateTimePicker1.Value.ToString("yyyyMMdd");
            MOCTA002 = GETMAXMOCTA002(MOCTA001, MOCTA003);

            ADDMOCTAB4(textBoxID.Text.Trim());

            ADDCOPPURBATCHPUR(textBoxID.Text.Trim(), MOCTA001, MOCTA002);
            SEARCHCOPPURBATCHPUR(textBoxID.Text.Trim());

            MessageBox.Show("已完成請購單" + MOCTA001 + " " + MOCTA002);
        }

        private void button19_Click(object sender, EventArgs e)
        {
            string SDATES = dateTimePicker3.Value.ToString("yyyyMMdd");
            string EDATES = dateTimePicker4.Value.ToString("yyyyMMdd");
            SEARCH_COPTE_COPTF(SDATES, EDATES);
        }
        private void button20_Click(object sender, EventArgs e)
        {
            string SDATES = dateTimePicker5.Value.ToString("yyyyMMdd");
            string EDATES = dateTimePicker6.Value.ToString("yyyyMMdd");
            SEARCH_COPTC_COPTD(SDATES, EDATES);
        }
        private void button21_Click(object sender, EventArgs e)
        {
            //轉ERP的請購單並送簽
            string COPTC_TC001 = textBox10.Text;
            string COPTC_TC002 = textBox11.Text;
            string PURTA_TA001 = "A311";
            string PURTA_TA003 = DateTime.Now.ToString("yyyyMMdd");
            string PURTA_TA002 = GETMAXMOCTA002(PURTA_TA001, PURTA_TA003);

            ADDMOCTAB_BY_COPTC_COPTD(COPTC_TC001, COPTC_TC002, PURTA_TA001, PURTA_TA002, PURTA_TA003);
            SEARCH_PURTA_PURTB(COPTC_TC001, COPTC_TC002);
            //ADDCOPPURBATCHPUR(textBoxID.Text.Trim(), MOCTA001, MOCTA002);
            //SEARCHCOPPURBATCHPUR(textBoxID.Text.Trim());

            MessageBox.Show("已完成請購單:" + PURTA_TA001 + " " + PURTA_TA002);
        }


        #endregion

       
    }
}

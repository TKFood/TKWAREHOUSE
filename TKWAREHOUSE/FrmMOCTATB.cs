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
    public partial class FrmMOCTATB : Form
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
        int result;
        DataTable dt = new DataTable();

        public class DATA_MOCTC
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
            public string TC001;
            public string TC002;
            public string TC003;
            public string TC004;
            public string TC005;
            public string TC006;
            public string TC007;
            public string TC008;
            public string TC009;
            public string TC010;
            public string TC011;
            public string TC012;
            public string TC013;
            public string TC014;
            public string TC015;
            public string TC016;
            public string TC017;
            public string TC018;
            public string TC019;
            public string TC020;
            public string TC021;
            public string TC022;
            public string TC023;
            public string TC024;
            public string TC025;
            public string TC026;
            public string TC027;
            public string TC028;
            public string TC029;
            public string TC030;
            public string TC031;
            public string TC032;
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
            public string TC200;
            public string TC201;
            public string TC202;

        }

        public Report report1 { get; private set; }

        public FrmMOCTATB()
        {
            InitializeComponent();

            comboboxload1();
            comboBox1.Text = "原料";

            comboboxload2();
            comboBox2.Text = "物料";
        }


        #region FUNCTION
        public void comboboxload1()
        {


            //20210902密
            Class1 TKID = new Class1();//用new 建立類別實體
            SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

            //資料庫使用者密碼解密
            sqlsb.Password = TKID.Decryption(sqlsb.Password);
            sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

            String connectionString;
            sqlConn = new SqlConnection(sqlsb.ConnectionString);

            StringBuilder Sequel = new StringBuilder();
            Sequel.AppendFormat(@"
                                SELECT  [ID]
                                ,[KINDS]
                                ,[NAMES]
                                ,[KEYS]
                                ,[KEYS2]
                                FROM[TKWAREHOUSE].[dbo].[TBPARAS]
                                WHERE[KINDS] = 'FrmMOCTATB' 
                                ORDER BY[KEYS]
                                ");

            SqlDataAdapter da = new SqlDataAdapter(Sequel.ToString(), sqlConn);
            DataTable dt = new DataTable();
            sqlConn.Open();

            dt.Columns.Add("NAMES", typeof(string));

            da.Fill(dt);
            comboBox1.DataSource = dt.DefaultView;
            comboBox1.ValueMember = "NAMES";
            comboBox1.DisplayMember = "NAMES";
            sqlConn.Close();           

        }
        public void comboboxload2()
        {


            //20210902密
            Class1 TKID = new Class1();//用new 建立類別實體
            SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

            //資料庫使用者密碼解密
            sqlsb.Password = TKID.Decryption(sqlsb.Password);
            sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

            String connectionString;
            sqlConn = new SqlConnection(sqlsb.ConnectionString);

            StringBuilder Sequel = new StringBuilder();
            Sequel.AppendFormat(@"
                                SELECT  [ID]
                                ,[KINDS]
                                ,[NAMES]
                                ,[KEYS]
                                ,[KEYS2]
                                FROM[TKWAREHOUSE].[dbo].[TBPARAS]
                                WHERE[KINDS] = 'FrmMOCTATB' 
                                ORDER BY[KEYS]
                                ");

            SqlDataAdapter da = new SqlDataAdapter(Sequel.ToString(), sqlConn);
            DataTable dt = new DataTable();
            sqlConn.Open();

            dt.Columns.Add("NAMES", typeof(string));

            da.Fill(dt);
            comboBox2.DataSource = dt.DefaultView;
            comboBox2.ValueMember = "NAMES";
            comboBox2.DisplayMember = "NAMES";
            sqlConn.Close();

        }

        public void SETFASTREPORT(string KINDS, string TA003, string TA003B, string TA009, string TA012)
        {
            string SQL;
            report1 = new Report();
            report1.Load(@"REPORT\烘焙製令領用表.frx");

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
            SQL = SETFASETSQL(KINDS, TA003, TA003B, TA009, TA012);

            Table.SelectCommand = SQL;
            report1.Preview = previewControl1;
            report1.Show();
        }

        public string SETFASETSQL(string KINDS, string TA003, string TA003B, string TA009, string TA012)
        {
            StringBuilder FASTSQL = new StringBuilder();
            StringBuilder STRQUERY = new StringBuilder();

            FASTSQL.Clear();
            sbSqlQuery.Clear();

            //TA012 實際開工
            if (KINDS.Equals("原料"))
            {
                sbSqlQuery.AppendFormat(@" AND TB003 NOT LIKE '2%' 
                                         AND TA009 ='{0}' 
                                        ", TA009);
            }
            //TA009 預計開工
            else if (KINDS.Equals("物料"))
            {
                sbSqlQuery.AppendFormat(@" AND TB003 LIKE '2%' 
                                           AND TA012 ='{0}'
                                        ", TA012);
            }

            FASTSQL.AppendFormat(@"  
                                SELECT 
                                TA001 AS '製令單別'
                                ,TA002 AS '製令單號'
                                ,TA003 AS '開單日期'
                                ,TA006 AS '產品品號'
                                ,TA009 AS '預計開工'
                                ,TA012 AS '實際開工'
                                ,TA034 AS '產品品名'
                                ,TA015 AS '預計產量'
                                ,TA007 AS '產品單位'
                                ,TB003 AS '材料品號'
                                ,TB012 AS '材料品名' 
                                ,(CASE WHEN TB003 LIKE '1%' OR TB003 LIKE '3%' THEN TB004 ELSE CONVERT(INT, TB004) END )AS '需領用量'
                                ,TB007 AS '材料單位'
                                ,(SELECT ISNULL(SUM(LA005*LA011),0) FROM [TK].dbo.INVLA WHERE LA009 IN ('21003') AND  LA001=TB003)AS '庫存量'
                                FROM [TK].dbo.MOCTA,[TK].dbo.MOCTB
                                WHERE TA001=TB001 AND TA002=TB002
                                AND TA001='A513'
                                AND TA013='Y'
                                AND TA003>='{0}' AND TA003<='{1}'
                                {2}

                                ORDER BY TB003,TA001,TA002 
                                        

                                        ", TA003, TA003B, sbSqlQuery.ToString());

            return FASTSQL.ToString();
        }

        public void SEARCH_TBMOCTCTDTE(string KINDS, string TA003, string TA003B, string TA009, string TA012)
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

               
                if (KINDS.Equals("物料"))
                {
                   

                }
                else
                {
                   
                }

                sbSql.AppendFormat(@"  
                                    SELECT 
                                    [ID]
                                    ,[KINDS] AS '分類'
                                    ,[TA003]  AS '開始日期'
                                    ,[TA003B] AS '結束日期'
                                    ,[TA009] AS '預計開工'
                                    ,[TA012] AS '實際開工'
                                    ,[TC003] AS '領料單日期'
                                    ,[TC001] AS '領料單別'
                                    ,[TC002] AS '領料單號'
                                    FROM [TKWAREHOUSE].[dbo].[TBMOCTCTDTE]
                                    WHERE [KINDS]='{0}'
                                    AND [TA003]='{1}' AND [TA003B]='{2}'
                                    ORDER BY [ID]
                                    ", KINDS, TA003, TA003B);

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
        public void ADD_TBMOCTCTDTE(
                        string KINDS
                        , string TA003
                        , string TA003B
                        , string TA009
                        , string TA012
                        , string TC003
                        , string TC001
                        , string TC002
                        )
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
                                        INSERT INTO [TKWAREHOUSE].[dbo].[TBMOCTCTDTE]
                                        (
                                        [KINDS]
                                        ,[TA003]
                                        ,[TA003B]
                                        ,[TA009]
                                        ,[TA012]
                                        ,[TC003]
                                        ,[TC001]
                                        ,[TC002]
                                        )
                                        VALUES
                                        (
                                        '{0}'
                                        ,'{1}'
                                        ,'{2}'
                                        ,'{3}'
                                        ,'{4}'
                                        ,'{5}'
                                        ,'{6}'
                                        ,'{7}'
                                        )"
                                        , KINDS
                                        , TA003
                                        , TA003B
                                        , TA009
                                        , TA012
                                        , TC003
                                        , TC001
                                        , TC002
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
            catch
            {

            }

            finally
            {
                sqlConn.Close();
            }
        }
        public string GETMAXTC002(string TC001,string TC003)
        {
            string TC002;

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


                StringBuilder sbSql = new StringBuilder();
                sbSql.Clear();
                sbSqlQuery.Clear();
                ds.Clear();

                sbSql.AppendFormat(@" 
                                        SELECT ISNULL(MAX(TC002),'00000000000') AS TC002
                                        FROM [TK].[dbo].[MOCTC]
                                        WHERE  TC001='{0}' AND TC002 LIKE '%{1}%' 
                                        ", TC001, TC003);

                adapter1 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder1 = new SqlCommandBuilder(adapter1);
                sqlConn.Open();
                ds1.Clear();
                adapter1.Fill(ds1, "ds1");
                sqlConn.Close();


                if (ds1.Tables["ds1"].Rows.Count == 0)
                {
                    return null;
                }
                else
                {
                    if (ds1.Tables["ds1"].Rows.Count >= 1)
                    {
                        TC002 = SETTC002(TC003, ds1.Tables["ds1"].Rows[0]["TC002"].ToString());
                        return TC002;

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
        public string SETTC002(string TC003,string TC002)
        {
            if (TC002.Equals("00000000000"))
            {
                return TC003 + "001";
            }

            else
            {
                int serno = Convert.ToInt16(TC002.Substring(8, 3));
                serno = serno + 1;
                string temp = serno.ToString();
                temp = temp.PadLeft(3, '0');
                return TC003 + temp.ToString();
            }
        }

        public void ADD_TK_MOCTC_MOCTD_MOCTE(
                        string KINDS
                        , string TA003
                        , string TA003B
                        , string TA009
                        , string TA012
                        , string TC003
                        , string TC001
                        , string TC002
                         )
        {
            DATA_MOCTC MOCTC = new DATA_MOCTC();
            MOCTC.COMPANY = "TK";
            MOCTC.CREATOR = "120024";
            MOCTC.USR_GROUP = "122000";
            MOCTC.CREATE_DATE = TC003;
            MOCTC.MODIFIER = "120024";
            MOCTC.MODI_DATE = TC003;
            MOCTC.FLAG = "0";
            MOCTC.CREATE_TIME = DateTime.Now.ToString("HH:mm:ss");
            MOCTC.MODI_TIME = DateTime.Now.ToString("HH:mm:ss");
            MOCTC.TRANS_TYPE = "P001";
            MOCTC.TRANS_NAME = "MOCMI03";
            MOCTC.sync_date = "";
            MOCTC.sync_time = "";
            MOCTC.sync_mark = "";
            MOCTC.sync_count = "0";
            MOCTC.DataUser = "";
            MOCTC.DataGroup = "122000";
            MOCTC.TC001 = TC001;
            MOCTC.TC002 = TC002;
            MOCTC.TC003 = TC003;
            MOCTC.TC004 = "20";
            MOCTC.TC005 = "08";
            MOCTC.TC006 = "";
            MOCTC.TC007 = "";
            MOCTC.TC008 = "54";
            MOCTC.TC009 = "N";
            MOCTC.TC010 = "0";
            MOCTC.TC011 = "N";
            MOCTC.TC012 = "1";
            MOCTC.TC013 = "N";
            MOCTC.TC014 = TC003;
            MOCTC.TC015 = "";
            MOCTC.TC016 = "N";
            MOCTC.TC017 = "N";
            MOCTC.TC018 = "0";
            MOCTC.TC019 = "0";
            MOCTC.TC020 = "";
            MOCTC.TC021 = "";
            MOCTC.TC022 = "0";
            MOCTC.TC023 = "0";
            MOCTC.TC024 = "";
            MOCTC.TC025 = "";
            MOCTC.TC026 = "";
            MOCTC.TC027 = "";
            MOCTC.TC028 = "0";
            MOCTC.TC029 = "";
            MOCTC.TC030 = "";
            MOCTC.TC031 = "";
            MOCTC.TC032 = "0";
            MOCTC.UDF01 = "";
            MOCTC.UDF02 = "";
            MOCTC.UDF03 = "";
            MOCTC.UDF04 = "";
            MOCTC.UDF05 = "";
            MOCTC.UDF06 = "0";
            MOCTC.UDF07 = "0";
            MOCTC.UDF08 = "0";
            MOCTC.UDF09 = "0";
            MOCTC.UDF10 = "0";
            MOCTC.TC200 = "";
            MOCTC.TC201 = "";
            MOCTC.TC202 = "";


            StringBuilder TB003LIKE = new StringBuilder();
            if (KINDS.Equals("物料"))
            {
                TB003LIKE.AppendFormat(@" TB003 LIKE '2%' ");

            }
            else
            {
                TB003LIKE.AppendFormat(@" TB003 NOT LIKE '2%' ");
            }
           
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
                                INSERT INTO [TK].[dbo].[MOCTC]
                                (
                                [COMPANY],[CREATOR],[USR_GROUP],[CREATE_DATE],[MODIFIER],[MODI_DATE],[FLAG],[CREATE_TIME],[MODI_TIME],[TRANS_TYPE]
                                ,[TRANS_NAME],[sync_date],[sync_time],[sync_mark],[sync_count],[DataUser],[DataGroup],[TC001],[TC002],[TC003]
                                ,[TC004],[TC005],[TC006],[TC007],[TC008],[TC009],[TC010],[TC011],[TC012],[TC013]
                                ,[TC014],[TC015],[TC016],[TC017],[TC018],[TC019],[TC020],[TC021],[TC022],[TC023]
                                ,[TC024],[TC025],[TC026],[TC027],[TC028],[TC029],[TC030],[TC031],[TC032],[UDF01]
                                ,[UDF02],[UDF03],[UDF04],[UDF05],[UDF06],[UDF07],[UDF08],[UDF09],[UDF10],[TC200]
                                ,[TC201],[TC202]
                                )
                                VALUES
                                (
                                '{0}','{1}','{2}','{3}','{4}','{5}','{6}','{7}','{8}','{9}'
                                ,'{10}','{11}','{12}','{13}','{14}','{15}','{16}','{17}','{18}','{19}'
                                ,'{20}','{21}','{22}','{23}','{24}','{25}','{26}','{27}','{28}','{29}'
                                ,'{30}','{31}','{32}','{33}','{34}','{35}','{36}','{37}','{38}','{39}'
                                ,'{40}','{41}','{42}','{43}','{44}','{45}','{46}','{47}','{48}','{49}'
                                ,'{50}','{51}','{52}','{53}','{54}','{55}','{56}','{57}','{58}','{59}'
                                ,'{60}','{61}'
                                )
                                "
                                , MOCTC.COMPANY, MOCTC.CREATOR, MOCTC.USR_GROUP, MOCTC.CREATE_DATE, MOCTC.MODIFIER, MOCTC.MODI_DATE, MOCTC.FLAG, MOCTC.CREATE_TIME, MOCTC.MODI_TIME, MOCTC.TRANS_TYPE
                                , MOCTC.TRANS_NAME, MOCTC.sync_date, MOCTC.sync_time, MOCTC.sync_mark, MOCTC.sync_count, MOCTC.DataUser, MOCTC.DataGroup, MOCTC.TC001, MOCTC.TC002, MOCTC.TC003
                                , MOCTC.TC004, MOCTC.TC005, MOCTC.TC006, MOCTC.TC007, MOCTC.TC008, MOCTC.TC009, MOCTC.TC010, MOCTC.TC011, MOCTC.TC012, MOCTC.TC013
                                , MOCTC.TC014, MOCTC.TC015, MOCTC.TC016, MOCTC.TC017, MOCTC.TC018, MOCTC.TC019, MOCTC.TC020, MOCTC.TC021, MOCTC.TC022, MOCTC.TC023
                                , MOCTC.TC024, MOCTC.TC025, MOCTC.TC026, MOCTC.TC027, MOCTC.TC028, MOCTC.TC029, MOCTC.TC030, MOCTC.TC031, MOCTC.TC032, MOCTC.UDF01
                                , MOCTC.UDF02, MOCTC.UDF03, MOCTC.UDF04, MOCTC.UDF05, MOCTC.UDF06, MOCTC.UDF07, MOCTC.UDF08, MOCTC.UDF09, MOCTC.UDF10, MOCTC.TC200
                                , MOCTC.TC201, MOCTC.TC202
                                );

                sbSql.AppendFormat(@"
                                    INSERT INTO [TK].dbo.MOCTD
                                    (
                                    [COMPANY]
                                    ,[CREATOR]
                                    ,[USR_GROUP]
                                    ,[CREATE_DATE]
                                    ,[MODIFIER]
                                    ,[MODI_DATE]
                                    ,[FLAG]
                                    ,[CREATE_TIME]
                                    ,[MODI_TIME]
                                    ,[TRANS_TYPE]
                                    ,[TRANS_NAME]
                                    ,[sync_date]
                                    ,[sync_time]
                                    ,[sync_mark]
                                    ,[sync_count]
                                    ,[DataUser]
                                    ,[DataGroup]
                                    ,[TD001]
                                    ,[TD002]
                                    ,[TD003]
                                    ,[TD004]
                                    ,[TD005]
                                    ,[TD006]
                                    ,[TD007]
                                    ,[TD008]
                                    ,[TD009]
                                    ,[TD010]
                                    ,[TD011]
                                    ,[TD012]
                                    ,[TD013]
                                    ,[TD014]
                                    ,[TD015]
                                    ,[TD016]
                                    ,[TD017]
                                    ,[TD018]
                                    ,[TD019]
                                    ,[TD020]
                                    ,[TD021]
                                    ,[TD022]
                                    ,[TD023]
                                    ,[TD024]
                                    ,[TD025]
                                    ,[TD026]
                                    ,[TD027]
                                    ,[TD028]
                                    ,[TD500]
                                    ,[TD501]
                                    ,[TD502]
                                    ,[TD503]
                                    ,[TD504]
                                    ,[TD505]
                                    ,[TD506]
                                    ,[UDF01]
                                    ,[UDF02]
                                    ,[UDF03]
                                    ,[UDF04]
                                    ,[UDF05]
                                    ,[UDF06]
                                    ,[UDF07]
                                    ,[UDF08]
                                    ,[UDF09]
                                    ,[UDF10]
                                    )
                                    SELECT 
                                    'TK' [COMPANY]
                                    ,'120024' [CREATOR]
                                    ,'122000' [USR_GROUP]
                                    ,'{0}' [CREATE_DATE]
                                    ,'' [MODIFIER]
                                    ,'' [MODI_DATE]
                                    ,'1' [FLAG]
                                    ,'{1}' [CREATE_TIME]
                                    ,'' [MODI_TIME]
                                    ,'P001' [TRANS_TYPE]
                                    ,'MOCMI03' [TRANS_NAME]
                                    ,'' [sync_date]
                                    ,'' [sync_time]
                                    ,'' [sync_mark]
                                    ,'0' [sync_count]
                                    ,'' [DataUser]
                                    ,'' [DataGroup]
                                    ,'{2}' [TD001]
                                    ,'{3}' [TD002]
                                    ,製令單別 [TD003]
                                    ,製令單號 [TD004]
                                    ,'1' [TD005]
                                    ,組數 [TD006]
                                    ,'' [TD007]
                                    ,'1' [TD008]
                                    ,'' [TD009]
                                    ,'' [TD010]
                                    ,'' [TD011]
                                    ,'' [TD012]
                                    ,'N' [TD013]
                                    ,'' [TD014]
                                    ,'1' [TD015]
                                    ,'' [TD016]
                                    ,'1' [TD017]
                                    ,'' [TD018]
                                    ,'' [TD019]
                                    ,'' [TD020]
                                    ,'' [TD021]
                                    ,'' [TD022]
                                    ,'' [TD023]
                                    ,'0' [TD024]
                                    ,'0' [TD025]
                                    ,'' [TD026]
                                    ,'' [TD027]
                                    ,'' [TD028]
                                    ,'' [TD500]
                                    ,'' [TD501]
                                    ,'0' [TD502]
                                    ,'' [TD503]
                                    ,'' [TD504]
                                    ,'' [TD505]
                                    ,'0' [TD506]
                                    ,'' [UDF01]
                                    ,'' [UDF02]
                                    ,'' [UDF03]
                                    ,'' [UDF04]
                                    ,'' [UDF05]
                                    ,'0' [UDF06]
                                    ,'0' [UDF07]
                                    ,'0' [UDF08]
                                    ,'0' [UDF09]
                                    ,'0' [UDF10]
                                    FROM 
                                    (
                                    SELECT 
                                    TA001 AS '製令單別'
                                    ,TA002 AS '製令單號'
                                    ,TA015 AS '組數'
                                    FROM [TK].dbo.MOCTA,[TK].dbo.MOCTB
                                    WHERE TA001=TB001 AND TA002=TB002
                                    AND TA001='A513'
                                    AND TA013='Y'
                                    AND TA003>='{4}' AND TA003<='{5}'
                                    AND {6}
                                    AND TA012 ='{7}'
                                    GROUP BY TA001,TA002,TA015

                                    ) AS TEMP

                                    ", MOCTC.TC003, MOCTC.CREATE_TIME, MOCTC.TC001, MOCTC.TC002
                                    , TA003, TA003B, TB003LIKE, TA012);

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

        public DataTable FIND_MOCTA_MOCTB(string KINDS
                        , string TA003
                        , string TA003B
                        , string TA009
                        , string TA012
                        , string TC003
                        , string TC001
                        , string TC002
            )
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


                StringBuilder sbSql = new StringBuilder();
                sbSql.Clear();
                sbSqlQuery.Clear();
                ds.Clear();


                //TA012 實際開工
                if (KINDS.Equals("原料"))
                {
                    sbSqlQuery.AppendFormat(@" AND TB003 NOT LIKE '2%' 
                                         AND TA009 ='{0}' 
                                        ", TA009);
                }
                //TA009 預計開工
                else if (KINDS.Equals("物料"))
                {
                    sbSqlQuery.AppendFormat(@" AND TB003 LIKE '2%' 
                                           AND TA012 ='{0}'
                                        ", TA012);
                }

                sbSql.AppendFormat(@"  
                                SELECT 
                                TA001 AS '製令單別'
                                ,TA002 AS '製令單號'
                                ,TA003 AS '開單日期'
                                ,TA006 AS '產品品號'
                                ,TA009 AS '預計開工'
                                ,TA012 AS '實際開工'
                                ,TA034 AS '產品品名'
                                ,TA015 AS '預計產量'
                                ,TA007 AS '產品單位'
                                ,TB003 AS '材料品號'
                                ,TB012 AS '材料品名' 
                                ,(CASE WHEN TB003 LIKE '1%' OR TB003 LIKE '3%' THEN TB004 ELSE CONVERT(INT, TB004) END )AS '需領用量'
                                ,TB007 AS '材料單位'
                                ,(SELECT ISNULL(SUM(LA005*LA011),0) FROM [TK].dbo.INVLA WHERE LA009 IN ('21003') AND  LA001=TB003)AS '庫存量'
                                FROM [TK].dbo.MOCTA,[TK].dbo.MOCTB
                                WHERE TA001=TB001 AND TA002=TB002
                                AND TA001='A513'
                                AND TA013='Y'
                                AND TA003>='{0}' AND TA003<='{1}'
                                {2}

                                ORDER BY TB003,TA001,TA002 
                                        

                                        ", TA003, TA003B, sbSqlQuery.ToString());

                adapter1 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder1 = new SqlCommandBuilder(adapter1);
                sqlConn.Open();
                ds1.Clear();
                adapter1.Fill(ds1, "ds1");
                sqlConn.Close();


                if (ds1.Tables["ds1"].Rows.Count >= 1)
                {
                    return ds1.Tables["ds1"];

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
                sqlConn.Close();
            }
        }

        #endregion

        #region BUTTON
        private void button1_Click(object sender, EventArgs e)
        {
            SETFASTREPORT(comboBox1.Text.ToString(), dateTimePicker1.Value.ToString("yyyyMMdd"), dateTimePicker2.Value.ToString("yyyyMMdd"), dateTimePicker3.Value.ToString("yyyyMMdd"), dateTimePicker4.Value.ToString("yyyyMMdd"));
        }
        private void button3_Click(object sender, EventArgs e)
        {
            SEARCH_TBMOCTCTDTE(comboBox2.Text.ToString(), dateTimePicker5.Value.ToString("yyyyMMdd"), dateTimePicker6.Value.ToString("yyyyMMdd"), dateTimePicker7.Value.ToString("yyyyMMdd"), dateTimePicker8.Value.ToString("yyyyMMdd"));
        }
        private void button2_Click(object sender, EventArgs e)
        {
            string TC001 = "A543";
            string TC002 = "";
            string TC003 = dateTimePicker9.Value.ToString("yyyyMMdd");
            TC002 = GETMAXTC002(TC001, TC003);

            string KINDS = comboBox2.Text.ToString();
            string TA003 = dateTimePicker5.Value.ToString("yyyyMMdd");
            string TA003B = dateTimePicker6.Value.ToString("yyyyMMdd");
            string TA009 = "";
            string TA012 = "";

            if (KINDS.Equals("物料"))
            {                
                TA012 = dateTimePicker8.Value.ToString("yyyyMMdd");
            }
            else
            {

                TA009 = dateTimePicker7.Value.ToString("yyyyMMdd");
            }

            if (!string.IsNullOrEmpty(TC001) && !string.IsNullOrEmpty(TC002))
            {
                //產生結果
                ADD_TBMOCTCTDTE(
                        KINDS
                        , TA003
                        , TA003B
                        , TA009
                        , TA012
                        , TC003
                        , TC001
                        , TC002
                        );
                //產生ERP的領料單
                ADD_TK_MOCTC_MOCTD_MOCTE(
                        KINDS
                        , TA003
                        , TA003B
                        , TA009
                        , TA012
                        , TC003
                        , TC001
                        , TC002
                        );
            }

            SEARCH_TBMOCTCTDTE(comboBox2.Text.ToString(), dateTimePicker5.Value.ToString("yyyyMMdd"), dateTimePicker6.Value.ToString("yyyyMMdd"), dateTimePicker7.Value.ToString("yyyyMMdd"), dateTimePicker8.Value.ToString("yyyyMMdd"));
        }

        #endregion

      
    }
}

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

        SqlTransaction tran;
        SqlCommand cmd = new SqlCommand();
        DataSet ds1 = new DataSet();
        DataSet ds2 = new DataSet();
        DataSet ds3 = new DataSet();
        DataSet ds4 = new DataSet();


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

        public FrmCOPPURTA()
        {
            InitializeComponent();
        }

        #region FUNCTION

        public void SEARCHBTACHID()
        {
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

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
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

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
                    connectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
                    sqlConn = new SqlConnection(connectionString);

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
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sbSql.Clear();
                sbSqlQuery.Clear();

                sbSql.AppendFormat(@"  SELECT [ID] AS '批號',[TD001] AS '訂單單別',[TD002] AS '訂單單號',[TD003] AS '訂單序號',[TD004] AS '品號',[TD005] AS '品名',[TD008] AS '訂單數量',[TD009] AS '已交數量',[TD010] AS '單位',[TD024] AS '贈品量',[TD025] AS '贈品已交量'");
                sbSql.AppendFormat(@"  FROM [TKWAREHOUSE].[dbo].[COPPURBATCHCOPTD]");
                sbSql.AppendFormat(@"  WHERE  [ID]='{0}'",ID);
                sbSql.AppendFormat(@"  ");

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
                connectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

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
            DELTD001 = null;
            DELTD002 = null;
            DELTD003 = null;

            if (dataGridView2.CurrentRow != null)
            {
                int rowindex = dataGridView2.CurrentRow.Index;
                if (rowindex >= 0)
                {
                    DataGridViewRow row = dataGridView2.Rows[rowindex];
                    DELTD001 = row.Cells["訂單單別"].Value.ToString();
                    DELTD002 = row.Cells["訂單單號"].Value.ToString();
                    DELTD003 = row.Cells["訂單序號"].Value.ToString();


                }
                else
                {
                    DELTD001 = null;
                    DELTD002 = null;
                    DELTD003 = null;

                }
            }
        }

        public void DELCOPPURBATCHCOPTD(string ID, string TD001, string TD002, string TD003)
        {
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

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
                connectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sqlConn.Close();
                sqlConn.Open();
                tran = sqlConn.BeginTransaction();

                sbSql.Clear();

                sbSql.AppendFormat(" DELETE [TKWAREHOUSE].[dbo].[COPPURBATCHUSED]");
                sbSql.AppendFormat(" WHERE [ID]='{0}'", ID);
                sbSql.AppendFormat(" ");
                sbSql.AppendFormat(" INSERT INTO [TKWAREHOUSE].[dbo].[COPPURBATCHUSED]");
                sbSql.AppendFormat(" ([ID],[TD001],[TD002],[TD003],[TD004],[TD005],[TDNUM],[TDUNIT],[MB001],[MB002],[NUM],[UNIT])");
                sbSql.AppendFormat(" SELECT '{0}',TD001,TD002,TD003,TD004,TD005,NUM,MB004,MD003,MD035,(NUM*CAL),MD004", ID);
                sbSql.AppendFormat(" FROM (");
                sbSql.AppendFormat(" SELECT   TD001,TD002,TD003,TC053 ,TD013,TD004,TD005,TD006");
                sbSql.AppendFormat(" ,((CASE WHEN MB004=TD010 THEN ((TD008-TD009)+(TD024-TD025)) ELSE ((TD008-TD009)+(TD024-TD025))*INVMD.MD004 END)-ISNULL(MOCTA.TA017,0)) AS 'NUM'");
                sbSql.AppendFormat(" ,MB004");
                sbSql.AppendFormat(" ,((TD008-TD009)+(TD024-TD025)) AS 'COPNUM'");
                sbSql.AppendFormat(" ,TD010");
                sbSql.AppendFormat(" ,(CASE WHEN ISNULL(INVMD.MD002,'')<>'' THEN INVMD.MD002 ELSE TD010 END ) AS INVMDMD002");
                sbSql.AppendFormat(" ,(CASE WHEN INVMD.MD003>0 THEN INVMD.MD003 ELSE 1 END) AS INVMDMD003");
                sbSql.AppendFormat(" ,(CASE WHEN INVMD.MD004>0 THEN INVMD.MD004 ELSE (TD008-TD009) END ) AS INVMDMD004");
                sbSql.AppendFormat(" ,ISNULL(MOCTA.TA017,0) AS TA017");
                sbSql.AppendFormat(" ,[MC001],[MC004],BOMMD.[MD003],[MD035],BOMMD.[MD006],BOMMD.[MD007],BOMMD.[MD008],BOMMD.[MD004]");
                sbSql.AppendFormat(" ,CONVERT(decimal(16,3),(1/[MC004]*BOMMD.[MD006]/BOMMD.[MD007]*(1+BOMMD.[MD008]))) AS CAL");
                sbSql.AppendFormat(" FROM [TK].dbo.BOMMC,[TK].dbo.BOMMD,[TK].dbo.INVMB,[TK].dbo.COPTC,[TK].dbo.COPTD");
                sbSql.AppendFormat(" LEFT JOIN [TK].dbo.INVMD ON MD001=TD004 AND TD010=MD002");
                sbSql.AppendFormat(" LEFT JOIN [TK].dbo.MOCTA ON TA026=TD001 AND TA027=TD002 AND TD028=TD003 AND TA006=TD004");
                sbSql.AppendFormat(" WHERE BOMMC.MC001=BOMMD.MD001");
                sbSql.AppendFormat(" AND  BOMMD.MD001=TD004");
                sbSql.AppendFormat(" AND TD004=MB001");
                sbSql.AppendFormat(" AND TC001=TD001 AND TC002=TD002");
                sbSql.AppendFormat(" AND TD001+TD002+TD003 IN (SELECT TD001+TD002+TD003 FROM [TKWAREHOUSE].[dbo].[COPPURBATCHCOPTD] WHERE ID='{0}')",ID);
                sbSql.AppendFormat(" ) AS TEMP");
                sbSql.AppendFormat(" ");
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

        public void SEARCHCOPPURBATCHUSED(string ID)
        {
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

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

        public string GETMAXMOCTA002(string MOCTA001)
        {
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                StringBuilder sbSql = new StringBuilder();
                sbSql.Clear();
                sbSqlQuery.Clear();
                ds4.Clear();

                sbSql.AppendFormat(@"  SELECT ISNULL(MAX(TA002),'00000000000') AS ID ");
                sbSql.AppendFormat(@"  FROM [TK].[dbo].[PURTA] ");
                //sbSql.AppendFormat(@"  WHERE  TC001='{0}' AND TC003='{1}'", "A542","20170119");
                sbSql.AppendFormat(@"  WHERE [TA003]='{0}'", MOCTA003);
                sbSql.AppendFormat(@"  ");
                sbSql.AppendFormat(@"  ");

                adapter4 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder4 = new SqlCommandBuilder(adapter4);
                sqlConn.Open();
                ds4.Clear();
                adapter4.Fill(ds4, "ds4");
                sqlConn.Close();


                if (ds4.Tables["ds4"].Rows.Count == 0)
                {
                    return null;
                }
                else
                {
                    if (ds4.Tables["ds4"].Rows.Count >= 1)
                    {
                        MAXID = SETIDSTRING(ds4.Tables["ds4"].Rows[0]["ID"].ToString(), MOCTA003);
                        return MAXID;

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

        public string SETIDSTRING(string MAXID, string dt)
        {
            if (MAXID.Equals("00000000000"))
            {
                return dt + "001";
            }

            else
            {
                int serno = Convert.ToInt16(MAXID.Substring(8, 3));
                serno = serno + 1;
                string temp = serno.ToString();
                temp = temp.PadLeft(3, '0');
                return dt + temp.ToString();
            }
        }

        public void ADDMOCTAB(string ID)
        {
            PURTA PURTA = new PURTA();
            PURTB PURTB = new PURTB();

            PURTA = SETPURTA();
            PURTB = SETPURTB();

            connectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
            sqlConn = new SqlConnection(connectionString);

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
            sbSql.AppendFormat(" AND ([COPPURBATCHUSED].[MB001] LIKE '2%')");
            sbSql.AppendFormat(" AND [ID]='{0}'",ID);
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

        public void UPDATEPURTA()
        {
            if (!string.IsNullOrEmpty(MOCTA001) && !string.IsNullOrEmpty(MOCTA002) && !string.IsNullOrEmpty(ID))
            {
                connectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

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
            MOCTA002 = GETMAXMOCTA002(MOCTA001);

            ADDMOCTAB(textBoxID.Text.Trim());

            MessageBox.Show("已完成請購單" + MOCTA001 + " " + MOCTA002);

        }


        #endregion


    }
}

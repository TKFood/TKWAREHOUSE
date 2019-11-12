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
    public partial class frmBacthMOC : Form
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
        SqlDataAdapter adapter6 = new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilder6 = new SqlCommandBuilder();
        SqlDataAdapter adapter7 = new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilder7 = new SqlCommandBuilder();
        SqlDataAdapter adapter8 = new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilder8 = new SqlCommandBuilder();
        SqlDataAdapter adapter9 = new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilder9 = new SqlCommandBuilder();

        SqlTransaction tran;
        SqlCommand cmd = new SqlCommand();
        DataSet ds1 = new DataSet();
        DataSet ds2 = new DataSet();
        DataSet ds3 = new DataSet();
        DataSet ds4 = new DataSet();
        DataSet ds5 = new DataSet();
        DataSet ds6 = new DataSet();
        DataSet ds7 = new DataSet();
        DataSet ds8 = new DataSet();
        DataSet ds9 = new DataSet();

        DataTable dt = new DataTable();
        string tablename = null;
        int result;

        string ID;
        string ID2;
        string NEWID;
        string FEEDTC001;
        string FEEDTC002;
        string TC004;
        string TC005;

        public class MOCTCDATA
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

        public class MOCTDDATA
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

            public string TD001;
            public string TD002;
            public string TD003;
            public string TD004;
            public string TD005;
            public string TD006;
            public string TD007;
            public string TD008;
            public string TD009;
            public string TD010;
            public string TD011;
            public string TD012;
            public string TD013;
            public string TD014;
            public string TD015;
            public string TD016;
            public string TD017;
            public string TD018;
            public string TD019;
            public string TD020;
            public string TD021;
            public string TD022;
            public string TD023;
            public string TD024;
            public string TD025;
            public string TD026;
            public string TD027;
            public string TD028;
            public string TD500;
            public string TD501;
            public string TD502;
            public string TD503;
            public string TD504;
            public string TD505;
            public string TD506;
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

        public class MOCTEDATA
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

            public string TE001;
            public string TE002;
            public string TE003;
            public string TE004;
            public string TE005;
            public string TE006;
            public string TE007;
            public string TE008;
            public string TE009;
            public string TE010;
            public string TE011;
            public string TE012;
            public string TE013;
            public string TE014;
            public string TE015;
            public string TE016;
            public string TE017;
            public string TE018;
            public string TE019;
            public string TE020;
            public string TE021;
            public string TE022;
            public string TE023;
            public string TE024;
            public string TE025;
            public string TE026;
            public string TE027;
            public string TE028;
            public string TE029;
            public string TE030;
            public string TE031;
            public string TE032;
            public string TE033;
            public string TE034;
            public string TE035;
            public string TE036;
            public string TE037;
            public string TE038;
            public string TE039;
            public string TE040;
            public string TE500;
            public string TE501;
            public string TE502;
            public string TE503;
            public string TE504;
            public string TE505;
            public string TE506;
            public string TE507;
            public string TE508;
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
            public string TE200;
            public string TE201;
        }

        public frmBacthMOC()
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
                sbSql.AppendFormat(@"  FROM [TKWAREHOUSE].[dbo].[BTACHID]");
                sbSql.AppendFormat(@"  WHERE CONVERT(NVARCHAR,[BACTHDATES],112)='{0}'",dateTimePicker1.Value.ToString("yyyyMMdd"));
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

        public void SEARCHBTACHID2()
        {
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sbSql.Clear();
                sbSqlQuery.Clear();

                sbSql.AppendFormat(@"  SELECT [ID] AS '批號',CONVERT(NVARCHAR,[BACTHDATES],112) AS '日期'");
                sbSql.AppendFormat(@"  FROM [TKWAREHOUSE].[dbo].[BTACHID]");
                sbSql.AppendFormat(@"  WHERE CONVERT(NVARCHAR,[BACTHDATES],112)='{0}'", dateTimePicker2.Value.ToString("yyyyMMdd"));
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

        public void SEARCHBTACHMOCTE(string ID2)
        {
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sbSql.Clear();
                sbSqlQuery.Clear();

                sbSql.AppendFormat(@"  SELECT [BACTHMOCTE].[ID] AS '批號',[BACTHMOCTE].[TE004] AS '領料品號',[BACTHMOCTE].[MB002] AS '領料品名'");
                sbSql.AppendFormat(@"  ,MOCTE.TE005 AS '第一次領料量',MOCTE.TE010 AS '第一次領料批號'");
                sbSql.AppendFormat(@"  ,ROUND((([ATE005]-[SUMTE005])*[ATA017]/(SELECT SUM(ATA017) FROM [TKWAREHOUSE].[dbo].[BACTHMOCTA] BACTHMOCTA WHERE BACTHMOCTA.ID=[BACTHMOCTE].[ID])),3) AS '預計補料量'");
                sbSql.AppendFormat(@"  ,[TA001] AS '製令',[TA002] AS '製令號',[TA006] AS '品號',[BACTHMOCTA].[MB002] AS '品名',[TA017] AS '生產量',[UDF007] AS '單位',[ATA017] AS '總重g'");
                sbSql.AppendFormat(@"  ,(SELECT SUM(ATA017) FROM [TKWAREHOUSE].[dbo].[BACTHMOCTA] BACTHMOCTA WHERE BACTHMOCTA.ID=[BACTHMOCTE].[ID]) AS '分攤總重'");
                sbSql.AppendFormat(@"  FROM [TKWAREHOUSE].[dbo].[BACTHMOCTE],[TKWAREHOUSE].[dbo].[BACTHMOCTA],[TK].dbo.MOCTE");
                sbSql.AppendFormat(@"  WHERE [BACTHMOCTE].[ID]=[BACTHMOCTA].[ID]");
                sbSql.AppendFormat(@"  AND TE011=[TA001] AND TE012=[TA002] AND [BACTHMOCTE].[TE004]= MOCTE.[TE004]");
                sbSql.AppendFormat(@"  AND ([ATE005]-[SUMTE005])>0");
                sbSql.AppendFormat(@"  AND MOCTE.TE001='A541'");
                sbSql.AppendFormat(@"  AND [BACTHMOCTE].[ID]='{0}'",ID2);
                sbSql.AppendFormat(@"  ");
                sbSql.AppendFormat(@"  ");

                adapter6 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder6 = new SqlCommandBuilder(adapter6);
                sqlConn.Open();
                ds6.Clear();
                adapter6.Fill(ds6, "ds6");
                sqlConn.Close();


                if (ds6.Tables["ds6"].Rows.Count == 0)
                {
                    dataGridView5.DataSource = null;
                }
                else
                {
                    if (ds6.Tables["ds6"].Rows.Count >= 1)
                    {
                        //dataGridView1.Rows.Clear();
                        dataGridView5.DataSource = ds6.Tables["ds6"];
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

        public void SEARCHBACTHGENMOCTE(string ID2)
        {
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sbSql.Clear();
                sbSqlQuery.Clear();

                sbSql.AppendFormat(@" SELECT [ID] AS '批號',[TE001] AS '製令',[TE002] AS '製令號' ");
                sbSql.AppendFormat(@" FROM [TKWAREHOUSE].[dbo].[BACTHGENMOCTE] ");
                sbSql.AppendFormat(@" WHERE [ID]='{0}' ",ID2);
                sbSql.AppendFormat(@"  ");

                adapter8 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder8 = new SqlCommandBuilder(adapter8);
                sqlConn.Open();
                ds8.Clear();
                adapter8.Fill(ds8, "ds8");
                sqlConn.Close();


                if (ds8.Tables["ds8"].Rows.Count == 0)
                {
                    dataGridView6.DataSource = null;
                }
                else
                {
                    if (ds8.Tables["ds8"].Rows.Count >= 1)
                    {
                        //dataGridView1.Rows.Clear();
                        dataGridView6.DataSource = ds8.Tables["ds8"];
                        dataGridView6.AutoResizeColumns();
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

                    SEARCHBACTHMOCTA(ID);
                    SEARCHBACTHMOCTE(ID);
                }
                else
                {
                    textBoxID.Text = null;
                    ID = null;

                }
            }
        }

        private void dataGridView4_SelectionChanged(object sender, EventArgs e)
        {
            textBoxID2.Text = null;

            if (dataGridView4.CurrentRow != null)
            {
                int rowindex = dataGridView4.CurrentRow.Index;
                if (rowindex >= 0)
                {
                    DataGridViewRow row = dataGridView4.Rows[rowindex];
                    textBoxID2.Text = row.Cells["批號"].Value.ToString();
                    ID2 = row.Cells["批號"].Value.ToString();

                    SEARCHBTACHMOCTE(ID2);
                    SEARCHBACTHGENMOCTE(ID2);
                }
                else
                {
                    textBoxID2.Text = null;
                    ID2 = null;

                }
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
                sbSql.AppendFormat(@"  FROM [TKWAREHOUSE].[dbo].[BTACHID]");
                sbSql.AppendFormat(@"  WHERE CONVERT(NVARCHAR,[BACTHDATES],112)='{0}'",dateTimePicker1.Value.ToString("yyyyMMdd"));
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

        public string GETMAXTC002(string FEEDTC001,string dt)
        {
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                StringBuilder sbSql = new StringBuilder();
                sbSql.Clear();
                sbSqlQuery.Clear();
                ds2.Clear();

                sbSql.AppendFormat(@"  SELECT ISNULL(MAX(TC002),'0000000000') AS TC002");
                sbSql.AppendFormat(@"  FROM [TK].[dbo].[MOCTC]");
                sbSql.AppendFormat(@"  WHERE  TC001='{0}' AND TC003='{1}'",FEEDTC001, dt);
                sbSql.AppendFormat(@"  ");

                adapter7 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder7 = new SqlCommandBuilder(adapter7);
                sqlConn.Open();
                ds7.Clear();
                adapter7.Fill(ds7, "ds7");
                sqlConn.Close();


                if (ds7.Tables["ds7"].Rows.Count == 0)
                {
                    return null;
                }
                else
                {
                    if (ds7.Tables["ds7"].Rows.Count >= 1)
                    {
                        FEEDTC002 = SETTC002(ds7.Tables["ds7"].Rows[0]["TC002"].ToString(),dt);
                        return FEEDTC002;

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

        public string SETTC002(string FEEDTC002,string dt)
        {
            if (FEEDTC002.Equals("0000000000"))
            {
                return dt+ "001";
            }

            else
            {
                int serno = Convert.ToInt16(FEEDTC002.Substring(8, 3));
                serno = serno + 1;
                string temp = serno.ToString();
                temp = temp.PadLeft(3, '0');
                return dt + temp.ToString();
            }

        }

        public void ADDBTACHID(string ID)
        {
            if(!string.IsNullOrEmpty(ID))
            {
                try
                {
                    connectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
                    sqlConn = new SqlConnection(connectionString);

                    sqlConn.Close();
                    sqlConn.Open();
                    tran = sqlConn.BeginTransaction();

                    sbSql.Clear();
                    
                    sbSql.AppendFormat(" INSERT INTO [TKWAREHOUSE].[dbo].[BTACHID]");
                    sbSql.AppendFormat(" ([ID],[BACTHDATES])");
                    sbSql.AppendFormat(" VALUES ('{0}','{1}')",ID,dateTimePicker1.Value.ToString("yyyyMMdd"));
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
        
        public void SEARCHBACTHMOCTA(string ID)
        {
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sbSql.Clear();
                sbSqlQuery.Clear();

                sbSql.AppendFormat(@"  SELECT [ID] AS '批號',[TA001] AS '製令',[TA002] AS '製令號',[TA006] AS '品號',[MB002] AS '品名',[TA017] AS '生產量',[UDF007] AS '淨重',[ATA017] AS '總重'");
                sbSql.AppendFormat(@"  FROM [TKWAREHOUSE].[dbo].[BACTHMOCTA]");
                sbSql.AppendFormat(@"  WHERE [ID]='{0}' ", ID);
                sbSql.AppendFormat(@"  ");
                sbSql.AppendFormat(@"  ");

                adapter3 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder3 = new SqlCommandBuilder(adapter3);
                sqlConn.Open();
                ds3.Clear();
                adapter3.Fill(ds3, "ds3");
                sqlConn.Close();


                if (ds3.Tables["ds3"].Rows.Count == 0)
                {
                    dataGridView2.DataSource = null;
                }
                else
                {
                    if (ds3.Tables["ds3"].Rows.Count >= 1)
                    {
                        //dataGridView1.Rows.Clear();
                        dataGridView2.DataSource = ds3.Tables["ds3"];
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

        public void ADDBACTHMOCTA(string ID,string TA001,string TA002)
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

                    sbSql.AppendFormat(" INSERT INTO [TKWAREHOUSE].[dbo].[BACTHMOCTA]");
                    sbSql.AppendFormat(" ([ID],[TA001],[TA002],[TA006],[MB002],[TA017],[UDF007],[ATA017])");
                    sbSql.AppendFormat(" SELECT '{0}',TA001,TA002,TA006,MB002,TA017,INVMB.UDF07,TA017*INVMB.UDF07",ID);
                    sbSql.AppendFormat(" FROM [TK].dbo.MOCTA,[TK].dbo.INVMB");
                    sbSql.AppendFormat(" WHERE TA006=MB001");
                    sbSql.AppendFormat(" AND TA001='{0}' AND TA002='{1}'", TA001, TA002);
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
        public void DELBACTHMOCTA(string ID, string TA001, string TA002)
        {
            if (!string.IsNullOrEmpty(TA001)&& !string.IsNullOrEmpty(TA002))
            {
                try
                {
                    connectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
                    sqlConn = new SqlConnection(connectionString);

                    sqlConn.Close();
                    sqlConn.Open();
                    tran = sqlConn.BeginTransaction();

                    sbSql.Clear();

                    sbSql.AppendFormat(" DELETE [TKWAREHOUSE].[dbo].[BACTHMOCTA]");
                    sbSql.AppendFormat(" WHERE ID='{0}' AND [TA001]='{1}' AND [TA002]='{2}'",ID,TA001,TA002);
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
        }

        public void SEARCHBACTHMOCTE(string ID)
        {
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sbSql.Clear();
                sbSqlQuery.Clear();

                
                sbSql.AppendFormat(@" SELECT [ID] AS '批號',[TE004] AS '品號',[MB002] AS '品名',[SUMTE005] AS '領用量',[ATE005] AS '實際用量'");
                sbSql.AppendFormat(@" FROM [TKWAREHOUSE].[dbo].[BACTHMOCTE] ");
                sbSql.AppendFormat(@" WHERE [ID] ='{0}' ", ID);
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
                    dataGridView3.DataSource = null;
                }
                else
                {
                    if (ds4.Tables["ds4"].Rows.Count >= 1)
                    {
                        //dataGridView1.Rows.Clear();
                        dataGridView3.DataSource = ds4.Tables["ds4"];
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

        public void SEARCHMOCTC(string ID2)
        {
            TC004 = null;
            TC005 = null;

            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sbSql.Clear();
                sbSqlQuery.Clear();


                sbSql.AppendFormat(@"  SELECT TOP 1 TC004,TC005");
                sbSql.AppendFormat(@"  FROM [TK].dbo.MOCTC, [TK].dbo.MOCTE,[TKWAREHOUSE].[dbo].[BACTHMOCTA]");
                sbSql.AppendFormat(@"  WHERE TC001=TE001 AND TC002=TE002");
                sbSql.AppendFormat(@"  AND BACTHMOCTA.TA001=TE011 AND BACTHMOCTA.TA002=TE012");
                sbSql.AppendFormat(@"  AND BACTHMOCTA.ID='{0}'",ID2);
                sbSql.AppendFormat(@"  ");


                adapter9 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder9 = new SqlCommandBuilder(adapter9);
                sqlConn.Open();
                ds9.Clear();
                adapter9.Fill(ds9, "ds9");
                sqlConn.Close();


                if (ds9.Tables["ds9"].Rows.Count == 0)
                {
                    TC004 = null;
                    TC005 = null;
                }
                else
                {
                    if (ds9.Tables["ds9"].Rows.Count >= 1)
                    {
                        TC004 = ds9.Tables["ds9"].Rows[0]["TC004"].ToString();
                        TC005 = ds9.Tables["ds9"].Rows[0]["TC005"].ToString();
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


        public void ADDBACTHMOCTE(string ID)
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

                   
                    sbSql.AppendFormat(" INSERT INTO  [TKWAREHOUSE].[dbo].[BACTHMOCTE]");
                    sbSql.AppendFormat(" ([ID],[TE004],[MB002],[SUMTE005],[ATE005])");
                    sbSql.AppendFormat(" SELECT '{0}',TE004,TE017,SUM(TE005),0", ID);
                    sbSql.AppendFormat(" FROM [TK].dbo.MOCTE");
                    sbSql.AppendFormat(" WHERE TE011+TE012 IN (SELECT TA001+TA002 FROM [TKWAREHOUSE].[dbo].[BACTHMOCTA] WHERE ID='{0}')",ID);
                    sbSql.AppendFormat(" AND TE004 LIKE '1%'");
                    sbSql.AppendFormat(" GROUP BY TE004,TE017");
                    sbSql.AppendFormat(" ORDER BY TE004,TE017  ");
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
        }

        public void ADDBACTHGENMOCTE(string ID2,string TC001,string TC002)
        {
            if (!string.IsNullOrEmpty(ID2)&& !string.IsNullOrEmpty(TC001) && !string.IsNullOrEmpty(TC002))
            {
                try
                {
                    connectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
                    sqlConn = new SqlConnection(connectionString);

                    sqlConn.Close();
                    sqlConn.Open();
                    tran = sqlConn.BeginTransaction();

                    sbSql.Clear();

                    sbSql.AppendFormat(" INSERT INTO [TKWAREHOUSE].[dbo].[BACTHGENMOCTE]");
                    sbSql.AppendFormat(" ([ID],[TE001],[TE002])");
                    sbSql.AppendFormat(" VALUES ('{0}','{1}','{2}')",ID2,TC001,TC002);
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

        public void ADDMOCTCMOCTDMOCTE(string ID2, string TC001, string TC002)
        {
            SEARCHMOCTC(ID2);

            MOCTCDATA MOCTC = new MOCTCDATA();
            MOCTC = SETMOCTC(dateTimePicker2.Value,ID2,TC001,TC002,TC004,TC005);

            MOCTDDATA MOCTD = new MOCTDDATA();
            MOCTD = SETMOCTD(dateTimePicker2.Value, ID2, TC001, TC002);

            MOCTEDATA MOCTE = new MOCTEDATA();
            MOCTE = SETMOCTE(dateTimePicker2.Value, ID2, TC001, TC002);

            if (!string.IsNullOrEmpty(ID2) && !string.IsNullOrEmpty(TC001) && !string.IsNullOrEmpty(TC002))
            {
                try
                {
                    connectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
                    sqlConn = new SqlConnection(connectionString);

                    sqlConn.Close();
                    sqlConn.Open();
                    tran = sqlConn.BeginTransaction();

                    sbSql.Clear();

                   
                    sbSql.AppendFormat(" INSERT INTO [TK].[dbo].[MOCTC]");
                    sbSql.AppendFormat(" (");
                    sbSql.AppendFormat(" [COMPANY],[CREATOR],[USR_GROUP],[CREATE_DATE],[MODIFIER],[MODI_DATE],[FLAG],[CREATE_TIME],[MODI_TIME],[TRANS_TYPE]");
                    sbSql.AppendFormat(" ,[TRANS_NAME],[sync_date],[sync_time],[sync_mark],[sync_count],[DataUser],[DataGroup]");
                    sbSql.AppendFormat(" ,[TC001],[TC002],[TC003],[TC004],[TC005],[TC006],[TC007],[TC008],[TC009],[TC010]");
                    sbSql.AppendFormat(" ,[TC011],[TC012],[TC013],[TC014],[TC015],[TC016],[TC017],[TC018],[TC019],[TC020]");
                    sbSql.AppendFormat(" ,[TC021],[TC022],[TC023],[TC024],[TC025],[TC026],[TC027],[TC028],[TC029],[TC030]");
                    sbSql.AppendFormat(" ,[TC031],[TC032]");
                    sbSql.AppendFormat(" ,[UDF01],[UDF02],[UDF03],[UDF04],[UDF05],[UDF06],[UDF07],[UDF08],[UDF09],[UDF10]");
                    sbSql.AppendFormat(" ,[TC200],[TC201],[TC202]");
                    sbSql.AppendFormat(" )");
                    sbSql.AppendFormat(" VALUES");
                    sbSql.AppendFormat(" (");
                    sbSql.AppendFormat(" '{0}','{1}','{2}','{3}','{4}','{5}','{6}','{7}','{8}','{9}',", MOCTC.COMPANY, MOCTC.CREATOR, MOCTC.USR_GROUP, MOCTC.CREATE_DATE, MOCTC.MODIFIER, MOCTC.MODI_DATE, MOCTC.FLAG, MOCTC.CREATE_TIME, MOCTC.MODI_TIME, MOCTC.TRANS_TYPE);
                    sbSql.AppendFormat(" '{0}','{1}','{2}','{3}','{4}','{5}','{6}',", MOCTC.TRANS_NAME, MOCTC.sync_date, MOCTC.sync_time, MOCTC.sync_mark, MOCTC.sync_count, MOCTC.DataUser, MOCTC.DataGroup);
                    sbSql.AppendFormat(" '{0}','{1}','{2}','{3}','{4}','{5}','{6}','{7}','{8}',{9},", MOCTC.TC001, MOCTC.TC002, MOCTC.TC003, MOCTC.TC004, MOCTC.TC005, MOCTC.TC006, MOCTC.TC007, MOCTC.TC008, MOCTC.TC009, MOCTC.TC010);
                    sbSql.AppendFormat(" '{0}','{1}','{2}','{3}','{4}','{5}','{6}',{7},'{8}','{9}',", MOCTC.TC011, MOCTC.TC012, MOCTC.TC013, MOCTC.TC014, MOCTC.TC015, MOCTC.TC016, MOCTC.TC017, MOCTC.TC018, MOCTC.TC019, MOCTC.TC020);
                    sbSql.AppendFormat(" '{0}',{1},{2},'{3}','{4}','{5}','{6}',{7},'{8}','{9}',", MOCTC.TC021, MOCTC.TC022, MOCTC.TC023, MOCTC.TC024, MOCTC.TC025, MOCTC.TC026, MOCTC.TC027, MOCTC.TC028, MOCTC.TC029, MOCTC.TC030);
                    sbSql.AppendFormat(" '{0}',{1},", MOCTC.TC031, MOCTC.TC032);
                    sbSql.AppendFormat(" '{0}','{1}','{2}','{3}','{4}','{5}','{6}','{7}','{8}','{9}',", MOCTC.UDF01, MOCTC.UDF02, MOCTC.UDF03, MOCTC.UDF04, MOCTC.UDF05, MOCTC.UDF06, MOCTC.UDF07, MOCTC.UDF08, MOCTC.UDF09, MOCTC.UDF10);
                    sbSql.AppendFormat(" '{0}','{1}','{2}'", MOCTC.TC200, MOCTC.TC201, MOCTC.TC202);
                    sbSql.AppendFormat(" )");
                    sbSql.AppendFormat(" ");
                    sbSql.AppendFormat(" INSERT INTO [TK].[dbo].[MOCTD]");
                    sbSql.AppendFormat(" (");
                    sbSql.AppendFormat(" [COMPANY],[CREATOR],[USR_GROUP],[CREATE_DATE],[MODIFIER],[MODI_DATE],[FLAG],[CREATE_TIME],[MODI_TIME],[TRANS_TYPE]");
                    sbSql.AppendFormat(" ,[TRANS_NAME],[sync_date],[sync_time],[sync_mark],[sync_count],[DataUser],[DataGroup]");
                    sbSql.AppendFormat(" ,[TD001],[TD002],[TD003],[TD004],[TD005],[TD006],[TD007],[TD008],[TD009],[TD010]");
                    sbSql.AppendFormat(" ,[TD011],[TD012],[TD013],[TD014],[TD015],[TD016],[TD017],[TD018],[TD019],[TD020]");
                    sbSql.AppendFormat(" ,[TD021],[TD022],[TD023],[TD024],[TD025],[TD026],[TD027],[TD028]");
                    sbSql.AppendFormat(" ,[TD500],[TD501],[TD502],[TD503],[TD504],[TD505],[TD506]");
                    sbSql.AppendFormat(" ,[UDF01],[UDF02],[UDF03],[UDF04],[UDF05],[UDF06],[UDF07],[UDF08],[UDF09],[UDF10]");
                    sbSql.AppendFormat(" )");
                    sbSql.AppendFormat(" (");
                    sbSql.AppendFormat(" SELECT ");
                    sbSql.AppendFormat(" '{0}','{1}','{2}','{3}','{4}','{5}','{6}','{7}','{8}','{9}',", MOCTD.COMPANY, MOCTD.CREATOR, MOCTD.USR_GROUP, MOCTD.CREATE_DATE, MOCTD.MODIFIER, MOCTD.MODI_DATE, MOCTD.FLAG, MOCTD.CREATE_TIME, MOCTD.MODI_TIME, MOCTD.TRANS_TYPE);
                    sbSql.AppendFormat(" '{0}','{1}','{2}','{3}','{4}','{5}','{6}',", MOCTD.TRANS_NAME, MOCTD.sync_date, MOCTD.sync_time, MOCTD.sync_mark, MOCTD.sync_count, MOCTD.DataUser, MOCTD.DataGroup);
                    sbSql.AppendFormat(" '{0}','{1}',[TA001],[TA002],'{2}',{3},'{4}','{5}','{6}','{7}',", MOCTD.TD001, MOCTD.TD002, MOCTD.TD005, MOCTD.TD006, MOCTD.TD007, MOCTD.TD008, MOCTD.TD009, MOCTD.TD010);
                    sbSql.AppendFormat(" '{0}','{1}','{2}','{3}','{4}','{5}','{6}','{7}','{8}','{9}',", MOCTD.TD011, MOCTD.TD012, MOCTD.TD013, MOCTD.TD014, MOCTD.TD015, MOCTD.TD016, MOCTD.TD017, MOCTD.TD018, MOCTD.TD019, MOCTD.TD020);
                    sbSql.AppendFormat(" '{0}','{1}','{2}',{3},{4},'{5}','{6}','{7}',", MOCTD.TD021, MOCTD.TD022, MOCTD.TD023, MOCTD.TD024, MOCTD.TD025, MOCTD.TD026, MOCTD.TD027, MOCTD.TD028);
                    sbSql.AppendFormat(" '{0}','{1}',{2},'{3}','{4}','{5}',{6},", MOCTD.TD500, MOCTD.TD501, MOCTD.TD502, MOCTD.TD503, MOCTD.TD504, MOCTD.TD505, MOCTD.TD506);
                    sbSql.AppendFormat(" '{0}','{1}','{2}','{3}','{4}',{5},{6},{7},{8},{9}", MOCTD.UDF01, MOCTD.UDF02, MOCTD.UDF03, MOCTD.UDF04, MOCTD.UDF05, MOCTD.UDF06, MOCTD.UDF07, MOCTD.UDF08, MOCTD.UDF09, MOCTD.UDF10);
                    sbSql.AppendFormat(" FROM [TKWAREHOUSE].[dbo].[BACTHMOCTA]");
                    sbSql.AppendFormat(" WHERE ID='{0}'",ID2);
                    sbSql.AppendFormat(" )");
                    sbSql.AppendFormat(" ");
                    sbSql.AppendFormat(" INSERT INTO [TK].[dbo].[MOCTE]");
                    sbSql.AppendFormat(" (");
                    sbSql.AppendFormat(" [COMPANY],[CREATOR],[USR_GROUP],[CREATE_DATE],[MODIFIER],[MODI_DATE],[FLAG],[CREATE_TIME],[MODI_TIME],[TRANS_TYPE]");
                    sbSql.AppendFormat(" ,[TRANS_NAME],[sync_date],[sync_time],[sync_mark],[sync_count],[DataUser],[DataGroup]");
                    sbSql.AppendFormat(" ,[TE001],[TE002],[TE003],[TE004],[TE005],[TE006],[TE007],[TE008],[TE009],[TE010]");
                    sbSql.AppendFormat(" ,[TE011],[TE012],[TE013],[TE014],[TE015],[TE016],[TE017],[TE018],[TE019],[TE020]");
                    sbSql.AppendFormat(" ,[TE021],[TE022],[TE023],[TE024],[TE025],[TE026],[TE027],[TE028],[TE029],[TE030]");
                    sbSql.AppendFormat(" ,[TE031],[TE032],[TE033],[TE034],[TE035],[TE036],[TE037],[TE038],[TE039],[TE040]");
                    sbSql.AppendFormat(" ,[TE500],[TE501],[TE502],[TE503],[TE504],[TE505],[TE506],[TE507],[TE508]");
                    sbSql.AppendFormat(" ,[UDF01],[UDF02],[UDF03],[UDF04],[UDF05],[UDF06],[UDF07],[UDF08],[UDF09],[UDF10]");
                    sbSql.AppendFormat(" ,[TE200],[TE201]");
                    sbSql.AppendFormat(" )");
                    sbSql.AppendFormat(" (");
                    sbSql.AppendFormat(" SELECT ");
                    sbSql.AppendFormat(" '{0}','{1}','{2}','{3}','{4}','{5}','{6}','{7}','{8}','{9}',", MOCTD.COMPANY, MOCTD.CREATOR, MOCTD.USR_GROUP, MOCTD.CREATE_DATE, MOCTD.MODIFIER, MOCTD.MODI_DATE, MOCTD.FLAG, MOCTD.CREATE_TIME, MOCTD.MODI_TIME, MOCTD.TRANS_TYPE);
                    sbSql.AppendFormat(" '{0}','{1}','{2}','{3}','{4}','{5}','{6}',", MOCTD.TRANS_NAME, MOCTD.sync_date, MOCTD.sync_time, MOCTD.sync_mark, MOCTD.sync_count, MOCTD.DataUser, MOCTD.DataGroup);
                    sbSql.AppendFormat(" '{0}','{1}',RIGHT(REPLICATE('0',4) + CAST(ROW_NUMBER() OVER(ORDER BY MOCTE.TE003)  as NVARCHAR),4) [TE003],[BACTHMOCTE].[TE004],ROUND((([ATE005]-[SUMTE005])*[ATA017]/(SELECT SUM(ATA017) FROM [TKWAREHOUSE].[dbo].[BACTHMOCTA] BACTHMOCTA WHERE BACTHMOCTA.ID=[BACTHMOCTE].[ID])),3) [TE005],[MOCTE].[TE006],[MOCTE].[TE007],[MOCTE].[TE008],[MOCTE].[TE009],MOCTE.TE010",MOCTE.TE001, MOCTE.TE002);
                    sbSql.AppendFormat(" ,[TA001] [TE011],[TA002] AS [TE012],'' [TE013],[TE014],[TE015],[TE016],[TE017],[TE018],[TE019],[TE020]");
                    sbSql.AppendFormat(" ,[TE021],[TE022],[TE023],[TE024],[TE025],[TE026],[TE027],[TE028],[TE029],[TE030]");
                    sbSql.AppendFormat(" ,[TE031],[TE032],[TE033],[TE034],[TE035],[TE036],[TE037],[TE038],[TE039],[TE040]");
                    sbSql.AppendFormat(" ,[TE500],[TE501],[TE502],[TE503],[TE504],[TE505],[TE506],[TE507],[TE508]");
                    sbSql.AppendFormat(" ,[UDF01],[UDF02],[UDF03],[UDF04],[UDF05],[UDF06],[UDF07],[UDF08],[UDF09],[UDF10]");
                    sbSql.AppendFormat(" ,[TE200],[TE201]");
                    sbSql.AppendFormat(" FROM [TKWAREHOUSE].[dbo].[BACTHMOCTE],[TKWAREHOUSE].[dbo].[BACTHMOCTA],[TK].dbo.MOCTE");
                    sbSql.AppendFormat(" WHERE [BACTHMOCTE].[ID]=[BACTHMOCTA].[ID]");
                    sbSql.AppendFormat(" AND TE011=[TA001] AND TE012=[TA002] AND [BACTHMOCTE].[TE004]= MOCTE.[TE004]");
                    sbSql.AppendFormat(" AND ([ATE005]-[SUMTE005])>0");
                    sbSql.AppendFormat(" AND MOCTE.TE001='A541'");
                    sbSql.AppendFormat(" AND [BACTHMOCTE].[ID]='{0}'",ID2);
                    sbSql.AppendFormat(" )");
                    sbSql.AppendFormat(" ");
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
        }

        public MOCTCDATA SETMOCTC(DateTime dt, string ID2, string TC001, string TC002,string TC004,string TC005)
        {
            MOCTCDATA MOCTC = new MOCTCDATA();

            MOCTC.COMPANY = "TK";
            MOCTC.CREATOR = "120024";
            MOCTC.USR_GROUP = "103400";
            MOCTC.CREATE_DATE = dt.ToString("yyyyMMdd");
            MOCTC.MODIFIER = "120024";
            MOCTC.MODI_DATE = dt.ToString("yyyyMMdd");
            MOCTC.FLAG = "0";
            MOCTC.CREATE_TIME = dt.ToString("HH:mm:dd");
            MOCTC.MODI_TIME = dt.ToString("HH:mm:dd");
            MOCTC.TRANS_TYPE = "P003";
            MOCTC.TRANS_NAME = "MOCMI03";
            MOCTC.sync_date = null;
            MOCTC.sync_time = null;
            MOCTC.sync_mark = null;
            MOCTC.sync_count = "0";
            MOCTC.DataUser = null;
            MOCTC.DataGroup = "103400";

            MOCTC.TC001 = TC001;
            MOCTC.TC002 = TC002;
            MOCTC.TC003 = dt.ToString("yyyyMMdd");
            MOCTC.TC004 = TC004;
            MOCTC.TC005 = TC005;
            MOCTC.TC006 = null;
            MOCTC.TC007 = ID2;
            MOCTC.TC008 = "54";
            MOCTC.TC009 = "N";
            MOCTC.TC010 = "0";
            MOCTC.TC011 = "N";
            MOCTC.TC012 = "1";
            MOCTC.TC013 = "N";
            MOCTC.TC014 = dt.ToString("yyyyMMdd");
            MOCTC.TC015 = null;
            MOCTC.TC016 = "N";
            MOCTC.TC017 = "N";
            MOCTC.TC018 = "0";
            MOCTC.TC019 = null;
            MOCTC.TC020 = null;
            MOCTC.TC021 = null;
            MOCTC.TC022 = "0";
            MOCTC.TC023 = "0";
            MOCTC.TC024 = null;
            MOCTC.TC025 = null;
            MOCTC.TC026 = null;
            MOCTC.TC027 = null;
            MOCTC.TC028 = "0";
            MOCTC.TC029 = null;
            MOCTC.TC030 = null;
            MOCTC.TC031 = null;
            MOCTC.TC032 = "0";
            MOCTC.UDF01 = null;
            MOCTC.UDF02 = null;
            MOCTC.UDF03 = null;
            MOCTC.UDF04 = null;
            MOCTC.UDF05 = null;
            MOCTC.UDF06 = "0";
            MOCTC.UDF07 = "0";
            MOCTC.UDF08 = "0";
            MOCTC.UDF09 = "0";
            MOCTC.UDF10 = "0";
            MOCTC.TC200 = null;
            MOCTC.TC201 = null;
            MOCTC.TC202 = null;

            return MOCTC;
        }

        public MOCTDDATA SETMOCTD(DateTime dt,string ID2,string TC001,string TC002)
        {
            MOCTDDATA MOCTD = new MOCTDDATA();

            MOCTD.COMPANY = "TK";
            MOCTD.CREATOR = "120024";
            MOCTD.USR_GROUP = "103400";
            MOCTD.CREATE_DATE = dt.ToString("yyyyMMdd");
            MOCTD.MODIFIER = "120024";
            MOCTD.MODI_DATE = dt.ToString("yyyyMMdd");
            MOCTD.FLAG = "0";
            MOCTD.CREATE_TIME = dt.ToString("HH:mm:dd");
            MOCTD.MODI_TIME = dt.ToString("HH:mm:dd");
            MOCTD.TRANS_TYPE = "P003";
            MOCTD.TRANS_NAME = "MOCMI03";
            MOCTD.sync_date = null;
            MOCTD.sync_time = null;
            MOCTD.sync_mark = null;
            MOCTD.sync_count = "0";
            MOCTD.DataUser = null;
            MOCTD.DataGroup = "103400";

            MOCTD.TD001 = TC001;
            MOCTD.TD002 = TC002;
            MOCTD.TD003 = null;
            MOCTD.TD004 = null;
            MOCTD.TD005 = "1";
            MOCTD.TD006 = "1";
            MOCTD.TD007 = null;
            MOCTD.TD008 = "1";
            MOCTD.TD009 = null;
            MOCTD.TD010 = null;
            MOCTD.TD011 = null;
            MOCTD.TD012 = null;
            MOCTD.TD013 = "N";
            MOCTD.TD014 = null;
            MOCTD.TD015 = "1";
            MOCTD.TD016 = null;
            MOCTD.TD017 = "1";
            MOCTD.TD018 = null;
            MOCTD.TD019 = null;
            MOCTD.TD020 = null;
            MOCTD.TD021 = null;
            MOCTD.TD022 = null;
            MOCTD.TD023 = null;
            MOCTD.TD024 = "0";
            MOCTD.TD025 = "0";
            MOCTD.TD026 = null;
            MOCTD.TD027 = null;
            MOCTD.TD028 = null;
            MOCTD.TD500 = null;
            MOCTD.TD501 = null;
            MOCTD.TD502 = "0";
            MOCTD.TD503 = null;
            MOCTD.TD504 = null;
            MOCTD.TD505 = null;
            MOCTD.TD506 = "0";
            MOCTD.UDF01 = null;
            MOCTD.UDF02 = null;
            MOCTD.UDF03 = null;
            MOCTD.UDF04 = null;
            MOCTD.UDF05 = null;
            MOCTD.UDF06 = "0";
            MOCTD.UDF07 = "0";
            MOCTD.UDF08 = "0";
            MOCTD.UDF09 = "0";
            MOCTD.UDF10 = "0";

            return MOCTD;
        }

        public MOCTEDATA SETMOCTE(DateTime dt, string ID2, string TC001, string TC002)
        {
            MOCTEDATA MOCTE = new MOCTEDATA();

            MOCTE.COMPANY = "TK";
            MOCTE.CREATOR = "120024";
            MOCTE.USR_GROUP = "103400";
            MOCTE.CREATE_DATE = dt.ToString("yyyyMMdd");
            MOCTE.MODIFIER = "120024";
            MOCTE.MODI_DATE = dt.ToString("yyyyMMdd");
            MOCTE.FLAG = "0";
            MOCTE.CREATE_TIME = dt.ToString("HH:mm:dd");
            MOCTE.MODI_TIME = dt.ToString("HH:mm:dd");
            MOCTE.TRANS_TYPE = "P003";
            MOCTE.TRANS_NAME = "MOCMI03";
            MOCTE.sync_date = null;
            MOCTE.sync_time = null;
            MOCTE.sync_mark = null;
            MOCTE.sync_count = "0";
            MOCTE.DataUser = null;
            MOCTE.DataGroup = "103400";

            MOCTE.TE001 = TC001; 
            MOCTE.TE002 = TC002;
            MOCTE.TE003 = null;
            MOCTE.TE004 = null;
            MOCTE.TE005 = null;
            MOCTE.TE006 = null;
            MOCTE.TE007 = null;
            MOCTE.TE008 = null;
            MOCTE.TE009 = "****";
            MOCTE.TE010 = null;
            MOCTE.TE011 = null;
            MOCTE.TE012 = null;
            MOCTE.TE013 = null;
            MOCTE.TE014 = null;
            MOCTE.TE015 = null;
            MOCTE.TE016 = "1";
            MOCTE.TE017 = null;
            MOCTE.TE018 = null;
            MOCTE.TE019 = null;
            MOCTE.TE020 = null;
            MOCTE.TE021 = null;
            MOCTE.TE022 = null;
            MOCTE.TE023 = null;
            MOCTE.TE024 = null;
            MOCTE.TE025 = null;
            MOCTE.TE026 = null;
            MOCTE.TE027 = null;
            MOCTE.TE028 = null;
            MOCTE.TE029 = null;
            MOCTE.TE030 = null;
            MOCTE.TE031 = null;
            MOCTE.TE032 = null;
            MOCTE.TE033 = null;
            MOCTE.TE034 = null;
            MOCTE.TE035 = null;
            MOCTE.TE036 = null;
            MOCTE.TE037 = null;
            MOCTE.TE038 = null;
            MOCTE.TE039 = null;
            MOCTE.TE040 = null;
            MOCTE.TE500 = null;
            MOCTE.TE501 = null;
            MOCTE.TE502 = null;
            MOCTE.TE503 = null;
            MOCTE.TE504 = null;
            MOCTE.TE505 = null;
            MOCTE.TE506 = null;
            MOCTE.TE507 = null;
            MOCTE.TE508 = null;
            MOCTE.UDF01 = null;
            MOCTE.UDF02 = null;
            MOCTE.UDF03 = null;
            MOCTE.UDF04 = null;
            MOCTE.UDF05 = null;
            MOCTE.UDF06 = null;
            MOCTE.UDF07 = null;
            MOCTE.UDF08 = null;
            MOCTE.UDF09 = null;
            MOCTE.UDF10 = null;
            MOCTE.TE200 = null;
            MOCTE.TE201 = null;

            return MOCTE;
        }

        public void DELETEBACTHGENMOCTE(string ID2,string TC001,string TC002)
        {
            if (!string.IsNullOrEmpty(ID2) && !string.IsNullOrEmpty(TC001) && !string.IsNullOrEmpty(TC002))
            {
                try
                {
                    connectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
                    sqlConn = new SqlConnection(connectionString);

                    sqlConn.Close();
                    sqlConn.Open();
                    tran = sqlConn.BeginTransaction();

                    sbSql.Clear();
                  
                    sbSql.AppendFormat(" DELETE [TKWAREHOUSE].[dbo].[BACTHGENMOCTE]");
                    sbSql.AppendFormat(" WHERE [ID]='{0}' AND [TE001]='{1}' AND [TE002]='{2}'",ID2,TC001,TC002);
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
        public void UPDATEBACTHMOCTE(string ID, string TE004,string ATE005)
        {
            if (!string.IsNullOrEmpty(ID)&& !string.IsNullOrEmpty(TE004) && !string.IsNullOrEmpty(ATE005))
            {
                try
                {
                    connectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
                    sqlConn = new SqlConnection(connectionString);

                    sqlConn.Close();
                    sqlConn.Open();
                    tran = sqlConn.BeginTransaction();

                    sbSql.Clear();


                    sbSql.AppendFormat(" UPDATE [TKWAREHOUSE].[dbo].[BACTHMOCTE]");
                    sbSql.AppendFormat(" SET ATE005='{0}'",ATE005);
                    sbSql.AppendFormat(" WHERE ID='{0}' AND TE004='{1}'",ID,TE004);
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

        private void dataGridView3_SelectionChanged(object sender, EventArgs e)
        {
            if (dataGridView3.CurrentRow != null)
            {
                int rowindex = dataGridView3.CurrentRow.Index;
                if (rowindex >= 0)
                {
                    DataGridViewRow row = dataGridView3.Rows[rowindex];
                    textBox4.Text = row.Cells["品號"].Value.ToString();
                   
                }
                else
                {
                    textBox4.Text = null;
                    ID = null;

                }
            }
        }

        private void dataGridView6_SelectionChanged(object sender, EventArgs e)
        {
            textBox8.Text = null;
            textBox9.Text = null;

            if (dataGridView6.CurrentRow != null)
            {
                int rowindex = dataGridView6.CurrentRow.Index;
                if (rowindex >= 0)
                {
                    DataGridViewRow row = dataGridView6.Rows[rowindex];
                    textBox8.Text = row.Cells["製令"].Value.ToString();
                    textBox9.Text = row.Cells["製令號"].Value.ToString();

                }
                else
                {
                    textBox8.Text = null;
                    textBox9.Text = null;

                }
            }
        }

        public void SETNULL()
        {
            textBox3.Text = null;
        }
        #endregion

        #region BUTTON
        private void button1_Click(object sender, EventArgs e)
        {
            SEARCHBTACHID();
            SEARCHBACTHMOCTA(textBoxID.Text);
            SEARCHBACTHMOCTE(textBoxID.Text);
        }
        private void button2_Click(object sender, EventArgs e)
        {
            ID=GETMAXID();
            ADDBTACHID(ID);
            SEARCHBTACHID();
            //MessageBox.Show(ID.ToString());
        }

        private void button4_Click(object sender, EventArgs e)
        {
            ADDBACTHMOCTA(textBoxID.Text, textBox1.Text, textBox2.Text);
            SEARCHBACTHMOCTA(textBoxID.Text);
        }

        private void button5_Click(object sender, EventArgs e)
        {
            DELBACTHMOCTA(textBoxID.Text, textBox1.Text,textBox2.Text);
            SEARCHBACTHMOCTA(textBoxID.Text);
        }

        private void button6_Click(object sender, EventArgs e)
        {
            ADDBACTHMOCTE(textBoxID.Text);
            SEARCHBACTHMOCTE(textBoxID.Text);
        }

        private void button3_Click(object sender, EventArgs e)
        {
            UPDATEBACTHMOCTE(textBoxID.Text,textBox4.Text,textBox3.Text);
            SEARCHBACTHMOCTE(textBoxID.Text);

            SETNULL();
        }
        private void button7_Click(object sender, EventArgs e)
        {
            SEARCHBTACHID2();
            SEARCHBTACHMOCTE(textBoxID2.Text);
            SEARCHBACTHGENMOCTE(textBoxID2.Text);
        }
        private void button8_Click(object sender, EventArgs e)
        {
            FEEDTC001 = textBox5.Text;
            FEEDTC002=GETMAXTC002(FEEDTC001,dateTimePicker2.Value.ToString("yyyyMMdd"));

            ADDBACTHGENMOCTE(ID2, FEEDTC001, FEEDTC002);
            ADDMOCTCMOCTDMOCTE(ID2, FEEDTC001, FEEDTC002);

            SEARCHBACTHGENMOCTE(textBoxID2.Text);
            //MessageBox.Show(FEEDTC001+" "+ FEEDTC002);
        }

        private void button11_Click(object sender, EventArgs e)
        {
            DialogResult dialogResult = MessageBox.Show("要刪除了?", "要刪除了?", MessageBoxButtons.YesNo);
            if (dialogResult == DialogResult.Yes)
            {
                DELETEBACTHGENMOCTE(textBoxID2.Text, textBox8.Text, textBox9.Text);
                SEARCHBACTHGENMOCTE(textBoxID2.Text);
            }
               
        }


        #endregion

       
    }
}

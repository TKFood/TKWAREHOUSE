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

        DataTable dt = new DataTable();
        string tablename = null;
        int result;

        string ID;
        string ID2;
        string NEWID;
        string FEEDTC001;
        string FEEDTC002;

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
            public string sync_count;
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
            public string sync_count;
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
            public string sync_count;
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

                sbSql.AppendFormat(@"  SELECT ISNULL(MAX(TC002),'00000000000') AS TC002");
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
            if (FEEDTC002.Equals("00000000000"))
            {
                return dt+ "0001";
            }

            else
            {
                int serno = Convert.ToInt16(FEEDTC002.Substring(8, 3));
                serno = serno + 1;
                string temp = serno.ToString();
                temp = temp.PadLeft(4, '0');
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

            MessageBox.Show(FEEDTC001+" "+ FEEDTC002);
        }


        #endregion


    }
}

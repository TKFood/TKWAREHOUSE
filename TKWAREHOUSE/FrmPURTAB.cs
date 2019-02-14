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
using System.Text.RegularExpressions;
using System.Globalization;

namespace TKWAREHOUSE
{
    public partial class FrmPURTAB : Form
    {
        SqlConnection sqlConn = new SqlConnection();
        SqlCommand sqlComm = new SqlCommand();
        string connectionString;
        StringBuilder sbSql = new StringBuilder();
        StringBuilder sbSqlQuery = new StringBuilder();
        SqlDataAdapter adapter = new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilder = new SqlCommandBuilder();
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
        DataSet ds = new DataSet();
        DataSet ds2 = new DataSet();
        DataSet ds3 = new DataSet();
        DataSet ds4 = new DataSet();
        DataSet ds5 = new DataSet();

        int result;
        string tablename = null;
        string ID;
        string MAXID;

        string DELID;
        string DELMOCTA001;
        string DELMOCTA002;

        string MOCTA001;
        string MOCTA002;
        string MOCTA003;

        Thread TD;

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

        public FrmPURTAB()
        {
            InitializeComponent();

            SETGRIDVIEW();
        }

        #region FUNCTION
        public void SETGRIDVIEW()
        {
            dataGridView1.AlternatingRowsDefaultCellStyle.BackColor = Color.PaleTurquoise;      //奇數列顏色



            //先建立個 CheckBox 欄
            DataGridViewCheckBoxColumn cbCol = new DataGridViewCheckBoxColumn();
            cbCol.Width = 120;   //設定寬度
            cbCol.HeaderText = "　選擇";
            cbCol.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;   //置中
            cbCol.TrueValue = true;
            cbCol.FalseValue = false;
            dataGridView1.Columns.Insert(0, cbCol);


            //建立个矩形，等下计算 CheckBox 嵌入 GridView 的位置
            Rectangle rect = dataGridView1.GetCellDisplayRectangle(0, -1, true);
            rect.X = rect.Location.X + rect.Width / 4 - 18;
            rect.Y = rect.Location.Y + (rect.Height / 2 - 9);

            CheckBox cbHeader = new CheckBox();
            cbHeader.Name = "checkboxHeader";
            cbHeader.Size = new Size(18, 18);
            cbHeader.Location = rect.Location;

            //全选要设定的事件
            cbHeader.CheckedChanged += new EventHandler(cbHeader_CheckedChanged);

            //将 CheckBox 加入到 dataGridView
            dataGridView1.Controls.Add(cbHeader);
        }

        private void cbHeader_CheckedChanged(object sender, EventArgs e)
        {
            dataGridView1.EndEdit();

            foreach (DataGridViewRow dr in dataGridView1.Rows)
            {
                dr.Cells[0].Value = ((CheckBox)dataGridView1.Controls.Find("checkboxHeader", true)[0]).Checked;

            }

        }
        public void SEARCHMOCTA()
        {
            StringBuilder SLQURY = new StringBuilder();
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sbSql.Clear();

                SLQURY.Clear();

                if (checkBox1.Checked == true)
                {
                    SLQURY.AppendFormat(@"  AND TA001+TA002 NOT IN (SELECT [MOCTA001]+[MOCTA002] FROM [TKWAREHOUSE].dbo.PURTAB)");
                }


                sbSql.AppendFormat(@"  SELECT TA001 AS '單別',TA002 AS '單號',TA003 AS '生產日',TA034 AS '品名',TA006 AS '品號',TA015 AS '生產量',TA007 AS '單位'");
                sbSql.AppendFormat(@"  FROM [TK].dbo.[MOCTA]");
                sbSql.AppendFormat(@"  WHERE TA003>='{0}' AND TA003<='{1}'", dateTimePicker1.Value.ToString("yyyyMMdd"), dateTimePicker2.Value.ToString("yyyyMMdd"));
                sbSql.AppendFormat(@"  {0}", SLQURY.ToString());
                sbSql.AppendFormat(@"  ORDER BY TA003,TA034");
                sbSql.AppendFormat(@"  ");


                adapter = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder = new SqlCommandBuilder(adapter);
                sqlConn.Open();
                ds.Clear();
                adapter.Fill(ds, "TEMPds");
                sqlConn.Close();


                if (ds.Tables["TEMPds"].Rows.Count == 0)
                {
                    dataGridView1.DataSource = null;
                }
                else
                {
                    if (ds.Tables["TEMPds"].Rows.Count >= 1)
                    {
                        dataGridView1.DataSource = ds.Tables["TEMPds"];

                        dataGridView1.AutoResizeColumns();
                        dataGridView1.FirstDisplayedScrollingRowIndex = dataGridView2.RowCount - 1;


                    }

                }

            }
            catch
            {

            }
            finally
            {

            }
        }

        public void ADDPURTAB(string ID)
        {
            foreach (DataGridViewRow dr in this.dataGridView1.Rows)
            {
                if (dr.Cells[0].Value != null && (bool)dr.Cells[0].Value)
                {
                    try
                    {
                        connectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
                        sqlConn = new SqlConnection(connectionString);

                        sqlConn.Close();
                        sqlConn.Open();
                        tran = sqlConn.BeginTransaction();

                        sbSql.Clear();
                        sbSql.AppendFormat(@" INSERT INTO [TKWAREHOUSE].[dbo].[PURTAB]");
                        sbSql.AppendFormat(@" ([ID],[IDDATES],[MOCTA001],[MOCTA002],[MOCTA003],[MOCTA006],[MOCTA007],[MOCTA015],[MOCTA034],[PURTA001],[PURTA002])");
                        sbSql.AppendFormat(@" VALUES ({0},'{1}','{2}','{3}','{4}','{5}','{6}','{7}','{8}','{9}','{10}')", ID, dateTimePicker3.Value.ToString("yyyyMMdd"), dr.Cells["單別"].Value.ToString(), dr.Cells["單號"].Value.ToString(), dr.Cells["生產日"].Value.ToString(), dr.Cells["品號"].Value.ToString(), dr.Cells["單位"].Value.ToString(), dr.Cells["生產量"].Value.ToString(), dr.Cells["品名"].Value.ToString(), "", "");
                        sbSql.AppendFormat(@" ");


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
        }

        public void SEARCHPURTAB()
        {
            StringBuilder SLQURY = new StringBuilder();
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sbSql.Clear();

                sbSql.AppendFormat(@"  SELECT [ID] AS '批號',[IDDATES] AS '請購日',[PURTA001] AS '請購單別',[PURTA002] AS '請購單號'");
                sbSql.AppendFormat(@"  FROM [TKWAREHOUSE].[dbo].[PURTAB]");
                sbSql.AppendFormat(@"  WHERE [IDDATES]='{0}'", dateTimePicker3.Value.ToString("yyyyMMdd"));
                sbSql.AppendFormat(@"  GROUP BY  [ID],[IDDATES],[PURTA001],[PURTA002] ");
                sbSql.AppendFormat(@"  ");

                adapter2 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder2 = new SqlCommandBuilder(adapter2);
                sqlConn.Open();
                ds2.Clear();
                adapter2.Fill(ds2, "TEMPds2");
                sqlConn.Close();


                if (ds2.Tables["TEMPds2"].Rows.Count == 0)
                {
                    dataGridView2.DataSource = null;
                    dataGridView3.DataSource = null;
                }
                else
                {
                    if (ds2.Tables["TEMPds2"].Rows.Count >= 1)
                    {
                        dataGridView2.DataSource = ds2.Tables["TEMPds2"];

                        dataGridView2.AutoResizeColumns();
                        dataGridView2.FirstDisplayedScrollingRowIndex = dataGridView2.RowCount - 1;


                    }

                }

            }
            catch
            {

            }
            finally
            {

            }
        }
        private void dataGridView2_SelectionChanged(object sender, EventArgs e)
        {
            if (dataGridView2.CurrentRow != null)
            {
                int rowindex = dataGridView2.CurrentRow.Index;
                if (rowindex >= 0)
                {
                    DataGridViewRow row = dataGridView2.Rows[rowindex];
                    ID = row.Cells["批號"].Value.ToString();
                    MOCTA003 = row.Cells["請購日"].Value.ToString();
                    SEARCHPURTAB2();
                }
                else
                {

                }
            }
        }

        public void SEARCHPURTAB2()
        {
            StringBuilder SLQURY = new StringBuilder();
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sbSql.Clear();

                sbSql.AppendFormat(@"  SELECT [ID] AS '批號',[PURTA001] AS '請購單別',[PURTA002] AS '請購單號',[IDDATES] AS '請購日',[MOCTA001] AS '單別',[MOCTA002] AS '單號',[MOCTA003] AS '生產日',[MOCTA006] AS '品號',[MOCTA007] AS '單位',[MOCTA015] AS '生產量',[MOCTA034] AS '品名'");
                sbSql.AppendFormat(@"  FROM [TKWAREHOUSE].[dbo].[PURTAB]");
                sbSql.AppendFormat(@"  WHERE [ID]='{0}'", ID);
                sbSql.AppendFormat(@"  ");

                adapter3 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder3 = new SqlCommandBuilder(adapter3);
                sqlConn.Open();
                ds3.Clear();
                adapter3.Fill(ds3, "TEMPds3");
                sqlConn.Close();


                if (ds3.Tables["TEMPds3"].Rows.Count == 0)
                {
                    dataGridView3.DataSource = null;
                }
                else
                {
                    if (ds3.Tables["TEMPds3"].Rows.Count >= 1)
                    {
                        dataGridView3.DataSource = ds3.Tables["TEMPds3"];

                        dataGridView3.AutoResizeColumns();
                        dataGridView3.FirstDisplayedScrollingRowIndex = dataGridView2.RowCount - 1;


                    }

                }

            }
            catch
            {

            }
            finally
            {

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
                ds4.Clear();

                sbSql.AppendFormat(@"  SELECT ISNULL(MAX(ID),'00000000000') AS ID");
                sbSql.AppendFormat(@"  FROM [TKWAREHOUSE].[dbo].[PURTAB] ");
                //sbSql.AppendFormat(@"  WHERE  TC001='{0}' AND TC003='{1}'", "A542","20170119");
                sbSql.AppendFormat(@"  WHERE [IDDATES]='{0}'", dateTimePicker3.Value.ToString("yyyyMMdd"));
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
                        MAXID = SETID(ds4.Tables["TEMPds4"].Rows[0]["ID"].ToString(),dateTimePicker3.Value);
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

        public string  SETID(string MAXID,DateTime dt)
        {
            if (MAXID.Equals("00000000000"))
            {
                return dt.ToString("yyyyMMdd") + "001";
            }

            else
            {
                int serno = Convert.ToInt16(MAXID.Substring(8, 3));
                serno = serno + 1;
                string temp = serno.ToString();
                temp = temp.PadLeft(3, '0');
                return dt.ToString("yyyyMMdd") + temp.ToString();
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

        public void SETNULL()
        {
            textBox1.Text = null;
        }

        private void dataGridView3_SelectionChanged(object sender, EventArgs e)
        {
            if (dataGridView3.CurrentRow != null)
            {
                int rowindex = dataGridView3.CurrentRow.Index;
                if (rowindex >= 0)
                {
                    DataGridViewRow row = dataGridView3.Rows[rowindex];
                    DELID = row.Cells["批號"].Value.ToString();
                    DELMOCTA001 = row.Cells["單別"].Value.ToString();
                    DELMOCTA002 = row.Cells["單號"].Value.ToString();


                }
                else
                {
                    DELID = null;
                    DELMOCTA001 = null;
                    DELMOCTA002 = null;
                }
            }
        }

        public void DELPURTAB()
        {
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sqlConn.Close();
                sqlConn.Open();
                tran = sqlConn.BeginTransaction();

                sbSql.Clear();
                sbSql.AppendFormat(@" DELETE [TKWAREHOUSE].[dbo].[PURTAB] WHERE [ID]='{0}' AND [MOCTA001]='{1}' AND [MOCTA002]='{2}'",DELID,DELMOCTA001,DELMOCTA002);
                sbSql.AppendFormat(@" ");


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

                adapter5 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder5 = new SqlCommandBuilder(adapter5);
                sqlConn.Open();
                ds5.Clear();
                adapter5.Fill(ds5, "TEMPds5");
                sqlConn.Close();


                if (ds5.Tables["TEMPds5"].Rows.Count == 0)
                {
                    return null;
                }
                else
                {
                    if (ds5.Tables["TEMPds5"].Rows.Count >= 1)
                    {
                        MAXID = SETIDSTRING(ds5.Tables["TEMPds5"].Rows[0]["ID"].ToString(), MOCTA003);
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

        public void ADDMOCTAB()
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
            sbSql.AppendFormat(" ");
            sbSql.AppendFormat(" ");
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
            PURTA.TA001=MOCTA001;
            PURTA.TA002=MOCTA002;
            PURTA.TA003=MOCTA003;
            PURTA.TA004= "103400";
            PURTA.TA005= ID;
            PURTA.TA006 = null;
            PURTA.TA007="N";
            PURTA.TA008="0";
            PURTA.TA009="9";
            PURTA.TA010="20";
            PURTA.TA011 = "0";
            PURTA.TA012= "120025";
            PURTA.TA013= MOCTA003;
            PURTA.TA014=null;
            PURTA.TA015="0";
            PURTA.TA016="N";
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
            PURTA.TA045= null;
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

            return PURTB;
        }

       


        #endregion

        #region BUTTON
        private void button1_Click(object sender, EventArgs e)
        {
            SEARCHMOCTA();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if(!string.IsNullOrEmpty(textBox1.Text))
            {
                ADDPURTAB(textBox1.Text);
                SEARCHMOCTA();
                SEARCHPURTAB();

                SETNULL();
            }
            else
            {
                MessageBox.Show("取新批號");
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            ADDPURTAB(ID);
            SEARCHMOCTA();
            SEARCHPURTAB2();
        }
        private void button4_Click(object sender, EventArgs e)
        {
            SETNULL();
            SEARCHPURTAB();
        }

        private void button5_Click(object sender, EventArgs e)
        {
            textBox1.Text = GETMAXID();
        }
        private void button6_Click(object sender, EventArgs e)
        {
            DialogResult dialogResult = MessageBox.Show("要刪除了?", "要刪除了?", MessageBoxButtons.YesNo);
            if (dialogResult == DialogResult.Yes)
            {
                DELPURTAB();
                SEARCHMOCTA();
                SEARCHPURTAB2();
            }
            else if (dialogResult == DialogResult.No)
            {
                //do something else
            }

        }

        private void button7_Click(object sender, EventArgs e)
        {
            MOCTA001 = "A311";
            MOCTA002 = GETMAXMOCTA002(MOCTA001);

            ADDMOCTAB();

            MessageBox.Show("已完成請購單"+ MOCTA001+" "+ MOCTA002);

            //MessageBox.Show(MOCTA002);
        }

        #endregion

       
    }
}

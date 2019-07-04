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

namespace TKWAREHOUSE
{
    public partial class FrmMOCINV : Form
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
        SqlDataAdapter adapter6 = new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilder6 = new SqlCommandBuilder();
        SqlDataAdapter adapter7 = new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilder7 = new SqlCommandBuilder();
        SqlTransaction tran;
        SqlCommand cmd = new SqlCommand();
        DataSet ds = new DataSet();
        DataSet ds2 = new DataSet();
        DataSet ds3 = new DataSet();
        DataSet ds4 = new DataSet();
        DataSet ds5 = new DataSet();
        DataSet ds6 = new DataSet();
        DataSet ds7 = new DataSet();

        DataTable dt = new DataTable();
        string tablename = null;
        int result;

        DataTable ADDDT = new DataTable();

        string ID =null;
        string TA001 = "A121";
        string TA002;
        string ORIGINTA001 = null;
        string ORIGINTA002 = null;

        public class INVTADATA
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
            public string DataUser;
            public string sync_count;
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
            public string TA047;
            public string TA048;
            public string TA049;
            public string TA050;
            public string TA051;
            public string TA052;
            public string TA053;
            public string TA054;
            public string TA055;
            public string TA056;
            public string TA057;
            public string TA058;
            public string TA059;
            public string TA060;
            public string TA061;
            public string TA062;
            public string TA063;
            public string TA064;
            public string TA065;
            public string TA066;
            public string TA067;
            public string TA068;
            public string TA200;
        }

        public class INVTB
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

        }

        public FrmMOCINV()
        {
            InitializeComponent();

            
            ADDDT.Columns.AddRange(new DataColumn[6] {
                 new DataColumn("庫別", typeof(string)),
                 new DataColumn("日期", typeof(string)),
                 new DataColumn("品號", typeof(string)),
                 new DataColumn("品名", typeof(string)),
                 new DataColumn("批號", typeof(string)),
                 new DataColumn("數量", typeof(decimal)) });
        }

        #region FUNCTION

        private void cbHeader_CheckedChanged(object sender, EventArgs e)
        {
            dataGridView1.EndEdit();

            foreach (DataGridViewRow dr in dataGridView1.Rows)
            {
                dr.Cells[0].Value = ((CheckBox)dataGridView1.Controls.Find("checkboxHeader", true)[0]).Checked;

            }

        }

        private void FrmMOCINV_Load(object sender, EventArgs e)
        {

            dataGridView1.AlternatingRowsDefaultCellStyle.BackColor = Color.PaleTurquoise;      //奇數列顏色

            //先建立個 CheckBox 欄
            DataGridViewCheckBoxColumn cbCol = new DataGridViewCheckBoxColumn();
            cbCol.Width = 80;   //設定寬度
            cbCol.HeaderText = "　全選";
            cbCol.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;   //置中
            cbCol.TrueValue = true;
            cbCol.FalseValue = false;
            dataGridView1.Columns.Insert(0, cbCol);

            #region 建立全选 CheckBox

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


            #endregion
        }

        public void SearchMOC()
        {
            try
            {
                if (!string.IsNullOrEmpty(dateTimePicker1.Value.ToString("yyyyMMdd")) || !string.IsNullOrEmpty(dateTimePicker2.Value.ToString("yyyyMMdd")))
                {
                    connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                    sqlConn = new SqlConnection(connectionString);

                    sbSql.Clear();
                    sbSqlQuery.Clear();
                    sbSql = SETsbSql();


                    adapter = new SqlDataAdapter(@"" + sbSql, sqlConn);
                    sqlCmdBuilder = new SqlCommandBuilder(adapter);

                    sqlConn.Open();
                    ds.Clear();
                    adapter.Fill(ds, tablename);
                    sqlConn.Close();


                    if (ds.Tables[tablename].Rows.Count == 0)
                    {
                        dataGridView1.DataSource = null;
                    }
                    else
                    {
                    
                        dataGridView1.DataSource = ds.Tables[tablename];
                        dataGridView1.AutoResizeColumns();
                    }
                }
                else
                {

                }



            }
            catch
            {

            }
            finally
            {

            }
        }

        public StringBuilder SETsbSql()
        {
            StringBuilder STR = new StringBuilder();

            STR.AppendFormat(@"  SELECT TB003 AS '品號',TB012 AS '品名',SUM(TB004) AS 數量");
            STR.AppendFormat(@"  FROM [TK].dbo.MOCTA,[TK].dbo.MOCTB");
            STR.AppendFormat(@"  WHERE TA001=TB001 AND TA002=TB002");
            STR.AppendFormat(@"  AND TA003>='{0}' AND TA003<='{1}' ", dateTimePicker1.Value.ToString("yyyyMMdd"), dateTimePicker2.Value.ToString("yyyyMMdd"));
            //STR.AppendFormat(@"  AND TB003 LIKE '101%'  ");
            STR.AppendFormat(@"  AND( TB003 LIKE '1%' OR TB003 LIKE '2%')");
            STR.AppendFormat(@"  AND( TA021 LIKE '02%' OR TA021 LIKE '03%' OR TA021 LIKE '04%' OR TA021 LIKE '09%') ");
            STR.AppendFormat(@"  GROUP BY TB003,TB012");
            STR.AppendFormat(@"  ORDER BY TB003,TB012");
            
            STR.AppendFormat(@"  ");
            


            STR.AppendFormat(@"  ");
            tablename = "TEMPds1";
      
          

            return STR;
        }

        public void  SearchINV()
        {
            try
            {
                if (!string.IsNullOrEmpty(dateTimePicker1.Value.ToString("yyyyMMdd")) || !string.IsNullOrEmpty(dateTimePicker2.Value.ToString("yyyyMMdd")))
                {
                    connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                    sqlConn = new SqlConnection(connectionString);

                    sbSql.Clear();

                    sbSql.AppendFormat(@" SELECT LA009 AS '倉庫',LA001 AS '品號',LA016  AS '批號',SUM(LA005*LA011)  AS '庫存量' ");
                    sbSql.AppendFormat(@" FROM [TK].dbo.INVLA");
                    sbSql.AppendFormat(@" WHERE (LA009='20006' OR LA009='20004')");
                    sbSql.AppendFormat(@" AND LA001 IN (SELECT TB003");
                    sbSql.AppendFormat(@" FROM [TK].dbo.INVTA,[TK].dbo.MOCTB");
                    sbSql.AppendFormat(@" WHERE TA001=TB001 AND TA002=TB002");
                    sbSql.AppendFormat(@"  AND TA003>='{0}' AND TA003<='{1}' ", dateTimePicker1.Value.ToString("yyyyMMdd"), dateTimePicker2.Value.ToString("yyyyMMdd"));
                    sbSql.AppendFormat(@"  AND( TA021 LIKE '02%' OR TA021 LIKE '03%' OR TA021 LIKE '04%' OR TA021 LIKE '09%') ");
                    sbSql.AppendFormat(@"  AND( TB003 LIKE '1%' OR TB003 LIKE '2%')");
                    //sbSql.AppendFormat(@" AND TB003='101002001'");
                    sbSql.AppendFormat(@" GROUP BY TB003");
                    sbSql.AppendFormat(@" )");
                    sbSql.AppendFormat(@" GROUP BY LA009,LA001,LA016");
                    sbSql.AppendFormat(@" HAVING  SUM(LA005*LA011)>0");
                    sbSql.AppendFormat(@" ORDER BY LA001,LA009,LA016");
                    sbSql.AppendFormat(@" ");



                    adapter = new SqlDataAdapter(@"" + sbSql, sqlConn);
                    sqlCmdBuilder = new SqlCommandBuilder(adapter);

                    sqlConn.Open();
                    ds2.Clear();
                    adapter.Fill(ds2, tablename);
                    sqlConn.Close();


                    if (ds2.Tables[tablename].Rows.Count == 0)
                    {
                        dataGridView2.DataSource = null;
                    }
                    else
                    {

                        dataGridView2.DataSource = ds2.Tables[tablename];
                        dataGridView2.AutoResizeColumns();
                    }
                }
                else
                {

                }



            }
            catch
            {

            }
            finally
            {

            }
        }


        public void SearchINVBATCH()
        {
            string ID = dateTimePicker1.Value.ToString("yyyyMMdd");
            ADDDT.Clear();

            if (dataGridView1.Rows.Count>=1)
            {
                foreach (DataGridViewRow row in dataGridView1.Rows)
                {
                    string MB001 = row.Cells["品號"].Value.ToString();
                    decimal NUM= Convert.ToDecimal(row.Cells["數量"].Value.ToString());

                    DataSet ds3 = new DataSet();
                    ds3.Clear();
                    ds3 = SearchINVNOW(MB001);

                    if(ds3!=null && ds3.Tables["ds3"].Rows.Count>=1)
                    {
                        int ROWS = ds3.Tables["ds3"].Rows.Count;
                        int NOWROWS = ds3.Tables["ds3"].Rows.Count;

                        while (NUM>0 && NOWROWS >= 1)
                        {
                            if(Convert.ToDecimal(ds3.Tables["ds3"].Rows[ROWS - NOWROWS]["庫存量"].ToString())>= NUM)
                            {
                                ADDDT.Rows.Add(ds3.Tables["ds3"].Rows[ROWS - NOWROWS]["庫別"].ToString(), ID,row.Cells["品號"].Value.ToString(), row.Cells["品名"].Value.ToString(), ds3.Tables["ds3"].Rows[ROWS - NOWROWS]["批號"].ToString(), NUM);
                            }
                            else if (Convert.ToDecimal(ds3.Tables["ds3"].Rows[ROWS - NOWROWS]["庫存量"].ToString()) < NUM)
                            {
                                ADDDT.Rows.Add(ds3.Tables["ds3"].Rows[ROWS - NOWROWS]["庫別"].ToString(), ID, row.Cells["品號"].Value.ToString(), row.Cells["品名"].Value.ToString(), ds3.Tables["ds3"].Rows[ROWS - NOWROWS]["批號"].ToString(), Convert.ToDecimal(ds3.Tables["ds3"].Rows[ROWS - NOWROWS]["庫存量"].ToString()));
                            }

                            NUM=NUM- Convert.ToDecimal(ds3.Tables["ds3"].Rows[ROWS- NOWROWS]["庫存量"].ToString());
                            NOWROWS = NOWROWS - 1;
                        }
                       
                    }
                    else
                    {

                    }

                }
            }

            if (ADDDT.Rows.Count >= 1)
            {
                dataGridView3.DataSource = ADDDT;
            }
            

        }

        public DataSet SearchINVNOW(string MB001)
        {
            try
            {
                if (!string.IsNullOrEmpty(dateTimePicker1.Value.ToString("MB001")) )
                {
                    connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                    sqlConn = new SqlConnection(connectionString);

                    sbSql.Clear();

                    sbSql.AppendFormat(@" SELECT LA009 AS '庫別',LA001 AS '品號',LA016  AS '批號',SUM(LA005*LA011)  AS '庫存量'  ");
                    sbSql.AppendFormat(@" FROM [TK].dbo.INVLA ");
                    sbSql.AppendFormat(@" WHERE (LA009='20006' OR LA009='20004')");
                    sbSql.AppendFormat(@" AND LA001 IN (SELECT TB003");
                    sbSql.AppendFormat(@" FROM [TK].dbo.MOCTA,[TK].dbo.MOCTB");
                    sbSql.AppendFormat(@" WHERE TA001=TB001 AND TA002=TB002");
                    sbSql.AppendFormat(@"  AND( TA021 LIKE '02%' OR TA021 LIKE '03%' OR TA021 LIKE '04%' OR TA021 LIKE '09%') ");
                    sbSql.AppendFormat(@"  AND( TB003 LIKE '1%' OR TB003 LIKE '2%')");
                    sbSql.AppendFormat(@"  AND TA003>='{0}' AND TA003<='{1}' ", dateTimePicker1.Value.ToString("yyyyMMdd"), dateTimePicker2.Value.ToString("yyyyMMdd"));
                    sbSql.AppendFormat(@" AND TB003='{0}'", MB001);
                    sbSql.AppendFormat(@" GROUP BY TB003");
                    sbSql.AppendFormat(@" )");
                    sbSql.AppendFormat(@" GROUP BY LA009,LA001,LA016");
                    sbSql.AppendFormat(@" HAVING  SUM(LA005*LA011)>0");
                    sbSql.AppendFormat(@" ORDER BY LA009,LA001,LA016");
                    sbSql.AppendFormat(@" ");



                    adapter2 = new SqlDataAdapter(@"" + sbSql, sqlConn);
                    sqlCmdBuilder2 = new SqlCommandBuilder(adapter2);

                    sqlConn.Open();
                    ds3.Clear();
                    adapter2.Fill(ds3, "ds3");
                    sqlConn.Close();


                    if (ds3.Tables["ds3"].Rows.Count == 0)
                    {
                        return null;
                    }
                    else
                    {
                        return ds3;

                    }
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

        public void ADDTOTKWAREHOUSE()
        {
            if (dataGridView3.Rows.Count >= 1)
            {
                try
                {
                    connectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
                    sqlConn = new SqlConnection(connectionString);

                    sqlConn.Close();
                    sqlConn.Open();
                    tran = sqlConn.BeginTransaction();

                    sbSql.Clear();

                    foreach (DataGridViewRow row in dataGridView3.Rows)
                    {
                        sbSql.AppendFormat(" INSERT INTO [TKWAREHOUSE].[dbo].[INVBATCH]");
                        sbSql.AppendFormat(" ([ID],[DATES],[WHID],[MB001],[MB002],[LOTNO],[NUM],[TA001],[TA002])");
                        sbSql.AppendFormat(" VALUES('{0}','{1}','{2}','{3}','{4}','{5}','{6}','{7}','{8}')",textBox1.Text,dateTimePicker3.Value.ToString("yyyyMMdd") ,row.Cells["庫別"].Value.ToString(), row.Cells["品號"].Value.ToString(), row.Cells["品名"].Value.ToString(), row.Cells["批號"].Value.ToString(), row.Cells["數量"].Value.ToString(),"A121", textBox1.Text);
                        sbSql.AppendFormat(" ");
                       
                    }

                    
  


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

                
                sbSql.AppendFormat(@" SELECT ISNULL(MAX(ID),'00000000000') AS ID ");
                sbSql.AppendFormat(@" FROM  [TKWAREHOUSE].[dbo].[INVBATCH] ");
                sbSql.AppendFormat(@" WHERE CONVERT(NVARCHAR,DATES,112)='{0}' ",dateTimePicker3.Value.ToString("yyyyMMdd"));
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
                        ID = SETTID(ds4.Tables["TEMPds4"].Rows[0]["ID"].ToString());
                        return ID;

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

        public string SETTID(string ID)
        {
            if (ID.Equals("00000000000"))
            {
                return dateTimePicker3.Value.ToString("yyyyMMdd") + "001";
            }

            else
            {
                int serno = Convert.ToInt16(ID.Substring(8, 3));
                serno = serno + 1;
                string temp = serno.ToString();
                temp = temp.PadLeft(3, '0');
                return dateTimePicker3.Value.ToString("yyyyMMdd") + temp.ToString();
            }
        }

        public void ADDINVTATB()
        {
            INVTADATA INVTA = new INVTADATA();
            INVTA = SETINVTA();

            try
            {
                //check TA002=2,TA040=2
                if (INVTA.TA002.Substring(0, 1).Equals("2"))
                {
                    connectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
                    sqlConn = new SqlConnection(connectionString);

                    sqlConn.Close();
                    sqlConn.Open();
                    tran = sqlConn.BeginTransaction();

                    sbSql.Clear();

                    sbSql.AppendFormat(" INSERT INTO [TK].[dbo].[INVTA]");
                    sbSql.AppendFormat(" (");
                    sbSql.AppendFormat(" [COMPANY],[CREATOR],[USR_GROUP],[CREATE_DATE],[MODIFIER],[MODI_DATE],[FLAG],[CREATE_TIME],[MODI_TIME],[TRANS_TYPE],[TRANS_NAME]");
                    sbSql.AppendFormat(" ,[sync_date],[sync_time],[sync_mark],[sync_count],[DataUser],[DataGroup]");
                    sbSql.AppendFormat(" ,[TA001],[TA002],[TA003],[TA004],[TA005],[TA006],[TA007],[TA008],[TA009],[TA010]");
                    sbSql.AppendFormat(" ,[TA011],[TA012],[TA013],[TA014],[TA015],[TA016],[TA017],[TA018],[TA019],[TA020]");
                    sbSql.AppendFormat(" ,[TA021],[TA022],[TA023],[TA024],[TA025],[TA026],[TA027],[TA028],[TA029],[TA030]");
                    sbSql.AppendFormat(" ,[TA031],[TA032],[TA033],[TA034],[TA035],[TA036],[TA037],[TA038],[TA039],[TA040]");
                    sbSql.AppendFormat(" ,[TA041],[TA042],[TA043],[TA044],[TA045],[TA046],[TA047],[TA048],[TA049],[TA050]");
                    sbSql.AppendFormat(" ,[TA051],[TA052],[TA053],[TA054],[TA055],[TA056],[TA057],[TA058],[TA059],[TA060]");
                    sbSql.AppendFormat(" ,[TA061],[TA062],[TA063],[TA064],[TA065],[TA066],[TA067],[TA068],[TA200]");
                    sbSql.AppendFormat(" )");
                    sbSql.AppendFormat(" VALUES");
                    sbSql.AppendFormat(" (");
                    sbSql.AppendFormat(" '{0}','{1}','{2}','{3}','{4}','{5}','{6}','{7}','{8}','{9}','{10}'", INVTA.COMPANY, INVTA.CREATOR, INVTA.USR_GROUP, INVTA.CREATE_DATE, INVTA.MODIFIER, INVTA.MODI_DATE, INVTA.FLAG, INVTA.CREATE_TIME, INVTA.MODI_TIME, INVTA.TRANS_TYPE, INVTA.TRANS_NAME);
                    sbSql.AppendFormat(" ,'{0}','{1}','{2}','{3}','{4}','{5}'", INVTA.sync_date, INVTA.sync_time, INVTA.sync_mark, INVTA.sync_count, INVTA.DataUser, INVTA.DataGroup);
                    sbSql.AppendFormat(" ,'{0}','{1}','{2}','{3}','{4}','{5}','{6}','{7}','{8}','{9}'", INVTA.TA001, INVTA.TA002, INVTA.TA003, INVTA.TA004, INVTA.TA005, INVTA.TA006, INVTA.TA007, INVTA.TA008, INVTA.TA009, INVTA.TA010);
                    sbSql.AppendFormat(" ,'{0}','{1}','{2}','{3}','{4}','{5}','{6}','{7}','{8}','{9}'", INVTA.TA011, INVTA.TA012, INVTA.TA013, INVTA.TA014, INVTA.TA015, INVTA.TA016, INVTA.TA017, INVTA.TA018, INVTA.TA019, INVTA.TA020);
                    sbSql.AppendFormat(" ,'{0}','{1}','{2}','{3}','{4}','{5}','{6}','{7}','{8}','{9}'", INVTA.TA021, INVTA.TA022, INVTA.TA023, INVTA.TA024, INVTA.TA025, INVTA.TA026, INVTA.TA027, INVTA.TA028, INVTA.TA029, INVTA.TA030);
                    sbSql.AppendFormat(" ,'{0}','{1}','{2}','{3}','{4}','{5}','{6}','{7}','{8}','{9}'", INVTA.TA031, INVTA.TA032, INVTA.TA033, INVTA.TA034, INVTA.TA035, INVTA.TA036, INVTA.TA037, INVTA.TA038, INVTA.TA039, INVTA.TA040);
                    sbSql.AppendFormat(" ,'{0}','{1}','{2}','{3}','{4}','{5}','{6}','{7}','{8}','{9}'", INVTA.TA041, INVTA.TA042, INVTA.TA043, INVTA.TA044, INVTA.TA045, INVTA.TA046, INVTA.TA047, INVTA.TA048, INVTA.TA049, INVTA.TA050);
                    sbSql.AppendFormat(" ,'{0}','{1}','{2}','{3}','{4}','{5}','{6}','{7}','{8}','{9}'", INVTA.TA051, INVTA.TA052, INVTA.TA053, INVTA.TA054, INVTA.TA055, INVTA.TA056, INVTA.TA057, INVTA.TA058, INVTA.TA059, INVTA.TA060);
                    sbSql.AppendFormat(" ,'{0}','{1}','{2}','{3}','{4}','{5}','{6}','{7}','{8}'", INVTA.TA061, INVTA.TA062, INVTA.TA063, INVTA.TA064, INVTA.TA065, INVTA.TA066, INVTA.TA067, INVTA.TA068, INVTA.TA200);
                    sbSql.AppendFormat(" )");
                    sbSql.AppendFormat(" ");
                    sbSql.AppendFormat(" INSERT INTO [TK].[dbo].[INVTB]");
                    sbSql.AppendFormat(" (");
                    sbSql.AppendFormat(" [COMPANY],[CREATOR],[USR_GROUP],[CREATE_DATE],[MODIFIER],[MODI_DATE],[FLAG],[CREATE_TIME],[MODI_TIME],[TRANS_TYPE],[TRANS_NAME]");
                    sbSql.AppendFormat(" ,[sync_date],[sync_time],[sync_mark],[sync_count],[DataUser],[DataGroup]");
                    sbSql.AppendFormat(" ,[TB001],[TB002],[TB003],[TB004],[TB005],[TB006],[TB007],[TB008],[TB009],[TB010]");
                    sbSql.AppendFormat(" ,[TB011],[TB012],[TB013],[TB014],[TB015],[TB016],[TB017],[TB018],[TB019],[TB020]");
                    sbSql.AppendFormat(" ,[TB021],[TB022],[TB023],[TB024],[TB025],[TB026],[TB027],[TB028],[TB029],[TB030]");
                    sbSql.AppendFormat(" ,[TB031],[TB032],[TB033],[TB034],[TB035],[TB036],[TB037],[TB038],[TB039],[TB040]");
                    sbSql.AppendFormat(" ,[TB041],[TB042],[TB043],[TB044],[TB045],[TB046],[TB047],[TB048],[TB049],[TB050]");
                    sbSql.AppendFormat(" ,[TB051],[TB052],[TB053],[TB054],[TB055],[TB056],[TB057],[TB058],[TB059],[TB060]");
                    sbSql.AppendFormat(" ,[TB061],[TB062],[TB063],[TB064],[TB065],[TB066],[TB067],[TB068],[TB069],[TB070]");
                    sbSql.AppendFormat(" ,[TB071],[TB072],[TB073]");
                    sbSql.AppendFormat(" ,[UDF01],[UDF02],[UDF03],[UDF04],[UDF05],[UDF06],[UDF07],[UDF08],[UDF09],[UDF10]");
                    sbSql.AppendFormat(" )");
                    sbSql.AppendFormat(" SELECT");
                    sbSql.AppendFormat(" '{0}' AS COMPANY,'{1}' AS CREATOR,'{2}' AS USR_GROUP,'{3}' AS CREATE_DATE,'{4}' AS MODIFIER,'{5}' AS MODI_DATE,'{6}' AS FLAG,'{7}' AS CREATE_TIME,'{8}' AS MODI_TIME,'{9}' AS TRANS_TYPE,'{10}' AS TRANS_NAME", INVTA.COMPANY, INVTA.CREATOR, INVTA.USR_GROUP, INVTA.CREATE_DATE, INVTA.MODIFIER, INVTA.MODI_DATE, INVTA.FLAG, INVTA.CREATE_TIME, INVTA.MODI_TIME, INVTA.TRANS_TYPE, INVTA.TRANS_NAME);
                    sbSql.AppendFormat(" ,'{0}' AS sync_date,'{1}' AS sync_time,'{2}' AS  sync_mark,'{3}' AS sync_count,'{4}' AS DataUser,'{5}' AS DataGroup", INVTA.sync_date, INVTA.sync_time, INVTA.sync_mark, INVTA.sync_count, INVTA.DataUser, INVTA.DataGroup);
                    sbSql.AppendFormat(" ,TA001 AS TB001,TA002 AS TB002,RIGHT(REPLICATE('0', 4) + CAST( ROW_NUMBER() OVER(ORDER BY TA002) as NVARCHAR), 4)  AS TB003,INVMB.MB001 AS TB004,INVMB.MB002 AS TB005,INVMB.MB003 AS TB006,NUM AS TB007,INVMB.MB004 AS TB008,'0' AS TB009,ROUND(INVMB.MB065/INVMB.MB064,2) AS TB010");
                    sbSql.AppendFormat(" ,ROUND(INVMB.MB065/INVMB.MB064*NUM,2)  AS TB011,WHID AS TB012,'20012' AS TB013,LOTNO AS TB014,CASE WHEN  ISNULL((SELECT TOP 1 TH036 FROM [TK].dbo.PURTH WHERE TH004=INVMB.MB001 AND TH010=LOTNO),'')<>'' THEN (SELECT TOP 1 TH036 FROM [TK].dbo.PURTH WHERE TH004=INVMB.MB001 AND TH010=LOTNO )ELSE (SELECT TOP 1 TG018 FROM [TK].dbo.MOCTG WHERE TG004=INVMB.MB001 AND TG017=LOTNO ) END  AS TB015,CASE WHEN ISNULL((SELECT TOP 1 TH037 FROM [TK].dbo.PURTH WHERE TH004=INVMB.MB001 AND TH010=LOTNO),'')<>'' THEN (SELECT TOP 1 TH037 FROM [TK].dbo.PURTH WHERE TH004=INVMB.MB001 AND TH010=LOTNO) ELSE (SELECT TOP 1 TG018 FROM [TK].dbo.MOCTG WHERE TG004=INVMB.MB001 AND TG017=LOTNO ) END  AS TB016,'' AS TB017,'N' AS TB018,'N' AS TB019,'' AS TB020 ");
                    sbSql.AppendFormat(" ,'' AS TB021,'0' AS TB022,'' AS TB023,'N' AS TB024,'0' AS TB025,'0' AS TB026,'' AS TB027,'' AS TB028,'' AS TB029,'0' AS TB030");
                    sbSql.AppendFormat(" ,'0' AS TB031,'' AS TB032,'' AS TB033,'' AS TB034,'' AS TB035,'' AS TB036,'0' AS TB037,'0' AS TB038,'0' AS TB039,'' AS TB040");
                    sbSql.AppendFormat(" ,'' AS TB041,'' AS TB042,'' AS TB043,'' AS TB044,'0' AS TB045,'' AS TB046,'0' AS TB047,'' AS TB048,'' AS TB049,'0' AS TB050");
                    sbSql.AppendFormat(" ,'' AS TB051,'N' AS TB052,'' AS TB053,'' AS TB054,'0' AS TB055,'' AS TB056,'' AS TB057,'' AS TB058,'0' AS TB059,'0' AS TB060");
                    sbSql.AppendFormat(" ,'' AS TB061,'0' AS TB062,'' AS TB063,'0' AS TB064,'0' AS TB065,'0' AS TB066,'0' AS TB067,'' AS TB068,'' AS TB069,'' AS TB070");
                    sbSql.AppendFormat(" ,'' AS TB071,'' AS TB072,'' AS TB073");
                    sbSql.AppendFormat(" ,'' AS UDF01,'' AS UDF02,'' AS UDF03,'' AS UDF04,'' AS UDF05,'0' AS UDF06,'0' AS UDF07,'0' AS UDF08,'0' AS UDF09,'0' AS UDF10");
                    sbSql.AppendFormat(" FROM [TKWAREHOUSE].dbo.INVBATCH,[TK].dbo.INVMB");
                    sbSql.AppendFormat(" WHERE INVBATCH.MB001=INVMB.MB001");
                    sbSql.AppendFormat(" AND INVBATCH.TA002='{0}'",TA002);
                    sbSql.AppendFormat(" ");
                    sbSql.AppendFormat(" ");
                    sbSql.AppendFormat(" ");
                    sbSql.AppendFormat(" ");
                    sbSql.AppendFormat(" ");
                    sbSql.AppendFormat(" ");
                    sbSql.AppendFormat(" ");
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
                        UPDATEINVTA();

                        MessageBox.Show("完成");
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
    

        public INVTADATA SETINVTA()
        {
            INVTADATA INVTA = new INVTADATA();

            INVTA.COMPANY = "TK";
            INVTA.CREATOR = "120024";
            INVTA.USR_GROUP = "103400";
            INVTA.CREATE_DATE = DateTime.Now.ToString("yyyyMMdd");
            INVTA.MODIFIER = "120024";
            INVTA.MODI_DATE = DateTime.Now.ToString("yyyyMMdd");
            INVTA.FLAG = "0";
            INVTA.CREATE_TIME = DateTime.Now.ToString("HH:mm:dd");
            INVTA.MODI_TIME = DateTime.Now.ToString("HH:mm:dd");
            INVTA.TRANS_TYPE = "P001";
            INVTA.TRANS_NAME = "INVMI08";
            INVTA.sync_date = "";
            INVTA.sync_time = "";
            INVTA.sync_mark = "";
            INVTA.sync_count = "0";
            INVTA.DataUser = "";
            INVTA.DataGroup = "103400";
            INVTA.TA001 = "A121";
            INVTA.TA002 = TA002;
            INVTA.TA003 = dateTimePicker3.Value.ToString("yyyyMMdd");
            INVTA.TA004 = "103400";
            INVTA.TA005 = "";
            INVTA.TA006 = "N";
            INVTA.TA007 = "0";
            INVTA.TA008 = "20";
            INVTA.TA009 = "12";
            INVTA.TA010 = "0";
            INVTA.TA011 = "0";
            INVTA.TA012 = "0";
            INVTA.TA013 = "N";            
            INVTA.TA014 = dateTimePicker3.Value.ToString("yyyyMMdd");
            INVTA.TA015 = "120024";
            INVTA.TA016 = "0";
            INVTA.TA017 = "N";
            INVTA.TA018 = "";
            INVTA.TA019 = "0";
            INVTA.TA020 = "6";
            INVTA.TA021 = "";
            INVTA.TA022 = "";
            INVTA.TA023 = "";
            INVTA.TA024 = "";
            INVTA.TA025 = "";
            INVTA.TA026 = "";
            INVTA.TA027 = "";
            INVTA.TA028 = "";
            INVTA.TA029 = "";
            INVTA.TA030 = "";
            INVTA.TA031 = "";
            INVTA.TA032 = "";
            INVTA.TA033 = "0";
            INVTA.TA034 = "0";
            INVTA.TA035 = "";
            INVTA.TA036 = "";
            INVTA.TA037 = "";
            INVTA.TA038 = "";
            INVTA.TA039 = "";
            INVTA.TA040 = "0";
            INVTA.TA041 = "0";
            INVTA.TA042 = "0";
            INVTA.TA043 = "";
            INVTA.TA044 = "";
            INVTA.TA045 = "";
            INVTA.TA046 = "";
            INVTA.TA047 = "";
            INVTA.TA049 = "0";
            INVTA.TA050 = "0";
            INVTA.TA051 = "";
            INVTA.TA052 = "";
            INVTA.TA053 = "";
            INVTA.TA054 = "";
            INVTA.TA055 = "0";
            INVTA.TA056 = "0";
            INVTA.TA057 = "0";
            INVTA.TA058 = "0";
            INVTA.TA059 = "";
            INVTA.TA060 = "";
            INVTA.TA061 = "";
            INVTA.TA062 = "";
            INVTA.TA063 = "";
            INVTA.TA064 = "";
            INVTA.TA065 = "";
            INVTA.TA066 = "";
            INVTA.TA067 = "";
            INVTA.TA068 = "";
            INVTA.TA200 = "";


            return INVTA;
        }
        public string GETMAXTA002(string TA001)
        {
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                StringBuilder sbSql = new StringBuilder();
                sbSql.Clear();
                sbSqlQuery.Clear();
                ds4.Clear();

                sbSql.AppendFormat(@"  SELECT ISNULL(MAX(TA002),'00000000000') AS TA002");
                sbSql.AppendFormat(@"  FROM [TK].[dbo].[INVTA] ");
                //sbSql.AppendFormat(@"  WHERE  TC001='{0}' AND TC003='{1}'", "A542","20170119");
                sbSql.AppendFormat(@"  WHERE  TA001='{0}' AND TA003='{1}'", TA001, dateTimePicker3.Value.ToString("yyyyMMdd"));
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
            if (TA002.Equals("00000000000"))
            {
                return dateTimePicker3.Value.ToString("yyyyMMdd") + "001";
            }

            else
            {
                int serno = Convert.ToInt16(TA002.Substring(8, 3));
                serno = serno + 1;
                string temp = serno.ToString();
                temp = temp.PadLeft(3, '0');
                return dateTimePicker3.Value.ToString("yyyyMMdd") + temp.ToString();
            }
        }

        public void SEARCHINVBATCH2()
        {
            try
            {
                if (!string.IsNullOrEmpty(dateTimePicker1.Value.ToString("yyyyMMdd")) || !string.IsNullOrEmpty(dateTimePicker2.Value.ToString("yyyyMMdd")))
                {
                    connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                    sqlConn = new SqlConnection(connectionString);

                    sbSql.Clear();
                    sbSqlQuery.Clear();

                    sbSql.AppendFormat(@" SELECT [TA001] AS '轉撥單別',[TA002] AS '轉撥單號'");
                    sbSql.AppendFormat(@" FROM [TKWAREHOUSE].[dbo].[INVBATCH]");
                    sbSql.AppendFormat(@" WHERE [DATES]='{0}'",dateTimePicker3.Value.ToString("yyyyMMdd"));
                    sbSql.AppendFormat(@" GROUP BY [TA001],[TA002]");
                    sbSql.AppendFormat(@" ORDER BY [TA001],[TA002]");
                    sbSql.AppendFormat(@" ");

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

                        dataGridView4.DataSource = ds5.Tables["ds5"];
                        dataGridView1.AutoResizeColumns();
                    }
                }
                else
                {

                }



            }
            catch
            {

            }
            finally
            {

            }
        }

        public void UPDATEINVTA()
        {
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sqlConn.Close();
                sqlConn.Open();
                tran = sqlConn.BeginTransaction();

                sbSql.Clear();

                sbSql.AppendFormat(" UPDATE [TK].dbo.INVTA");
                sbSql.AppendFormat(" SET TA011=(SELECT SUM(TB007) FROM [TK].dbo.INVTB WHERE TB001=TA001 AND TB002=TA002), TA012=(SELECT SUM(TB011) FROM [TK].dbo.INVTB WHERE TB001=TA001 AND TB002=TA002)");
                sbSql.AppendFormat(" WHERE TA001='{0}' AND TA002='{1}'",TA001,TA002);
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
        public void SEARCHINVBATCH3()
        {
            try
            {
                if (!string.IsNullOrEmpty(dateTimePicker4.Value.ToString("yyyyMMdd")) || !string.IsNullOrEmpty(dateTimePicker5.Value.ToString("yyyyMMdd")))
                {
                    connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                    sqlConn = new SqlConnection(connectionString);

                    sbSql.Clear();
                    sbSqlQuery.Clear();

                    sbSql.AppendFormat(@" SELECT [TA001] AS '轉撥單別',[TA002] AS '轉撥單號'");
                    sbSql.AppendFormat(@" FROM [TKWAREHOUSE].[dbo].[INVBATCH]");
                    sbSql.AppendFormat(@" WHERE SUBSTRING([TA002],1,8)>='{0}' AND SUBSTRING([TA002],1,8)<='{1}'", dateTimePicker4.Value.ToString("yyyyMMdd"), dateTimePicker5.Value.ToString("yyyyMMdd"));
                    sbSql.AppendFormat(@" GROUP BY [TA001],[TA002]");
                    sbSql.AppendFormat(@" ORDER BY [TA001],[TA002]");
                    sbSql.AppendFormat(@" ");
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

                        dataGridView5.DataSource = ds6.Tables["ds6"];
                        dataGridView5.AutoResizeColumns();
                    }
                }
                else
                {

                }



            }
            catch
            {

            }
            finally
            {

            }
        }

    

        private void dataGridView5_SelectionChanged(object sender, EventArgs e)
        {
            ORIGINTA001 = null;
            ORIGINTA002 = null;

            if (dataGridView5.CurrentRow != null)
            {
                int rowindex = dataGridView5.CurrentRow.Index;
               
                if (rowindex >= 0)
                {
                    DataGridViewRow row = dataGridView5.Rows[rowindex];
                    ORIGINTA001 = row.Cells["轉撥單別"].Value.ToString();
                    ORIGINTA002 = row.Cells["轉撥單號"].Value.ToString();
                }
                else
                {
                    ORIGINTA001 = null;
                    ORIGINTA002 = null;

                }
            }
        }
        public void DELINVBATCHRETURN()
        {
            if(!string.IsNullOrEmpty(ORIGINTA001)&& !string.IsNullOrEmpty(ORIGINTA002))
            {
                try
                {
                    connectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
                    sqlConn = new SqlConnection(connectionString);

                    sqlConn.Close();
                    sqlConn.Open();
                    tran = sqlConn.BeginTransaction();

                    sbSql.Clear();

                    sbSql.AppendFormat(" DELETE [TKWAREHOUSE].[dbo].[INVBATCHRETURN]");
                    sbSql.AppendFormat(" WHERE [TA001]='{}' AND [TA002]='{1}'", ORIGINTA001, ORIGINTA002);
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
        public void ADDINVBATCHRETURN()
        {
            if (!string.IsNullOrEmpty(ORIGINTA001) && !string.IsNullOrEmpty(ORIGINTA002))
            {
                try
                {
                    connectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
                    sqlConn = new SqlConnection(connectionString);

                    sqlConn.Close();
                    sqlConn.Open();
                    tran = sqlConn.BeginTransaction();

                    sbSql.Clear();

                    sbSql.AppendFormat(" INSERT INTO [TKWAREHOUSE].[dbo].[INVBATCHRETURN]");
                    sbSql.AppendFormat(" ([TA001],[TA002],[MB001],[LOTNO],[NUM],[USED],[RETURNED],[TA001RE],[TA002RE])");
                    sbSql.AppendFormat(" SELECT TA001 AS 'TA001',TA002 AS 'TA002',TB004 AS 'MB001',TB014 AS 'LOTNO',TB007 AS 'NUM'");
                    sbSql.AppendFormat(" ,(SELECT ISNULL(SUM(TE005),0) FROM [TK].dbo.MOCTE,[TK].dbo.MOCTC WHERE MOCTC.TC001=MOCTE.TE001 AND  MOCTC.TC002=MOCTE.TE002 AND MOCTC.TC003=INVTA.TA003 AND MOCTE.TE004=INVTB.TB004 AND MOCTE.TE010=INVTB.TB014 AND MOCTE.TE001 IN('A541','A542')) AS 'USED'");
                    sbSql.AppendFormat(" ,(SELECT ISNULL(SUM(TE005),0) FROM [TK].dbo.MOCTE,[TK].dbo.MOCTC WHERE MOCTC.TC001=MOCTE.TE001 AND  MOCTC.TC002=MOCTE.TE002 AND MOCTC.TC003=INVTA.TA003 AND MOCTE.TE004=INVTB.TB004 AND MOCTE.TE010=INVTB.TB014 AND MOCTE.TE001 IN('A561')) AS 'RETURNED'");
                    sbSql.AppendFormat(" ,NULL AS 'TA001RE',NULL AS 'TA002RE'");
                    sbSql.AppendFormat(" FROM [TK].dbo.INVTB,[TK].dbo.INVTA");
                    sbSql.AppendFormat(" WHERE TA001=TB001 AND TA002=TB002 ");
                    sbSql.AppendFormat(" AND TB001='{0}' AND TB002='{1}'", ORIGINTA001, ORIGINTA002);
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
        public void SEARCHINVBATCHRETURN()
        {
            try
            {
                if (!string.IsNullOrEmpty(dateTimePicker1.Value.ToString("yyyyMMdd")) || !string.IsNullOrEmpty(dateTimePicker2.Value.ToString("yyyyMMdd")))
                {
                    connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                    sqlConn = new SqlConnection(connectionString);

                    sbSql.Clear();
                    sbSqlQuery.Clear();

                    sbSql.AppendFormat(@" SELECT [TA001] AS '原轉撥單別' ,[TA002] AS '原轉撥單號' ,[MB001] AS '品號' ,[LOTNO] AS '批號' ,[NUM] AS '轉撥量' ,[USED] AS '領用量' ,[RETURNED] AS '退料量' ,[TA001RE] AS '回轉撥單別' ,[TA002RE] AS '回轉撥單號' ");
                    sbSql.AppendFormat(@" FROM [TKWAREHOUSE].[dbo].[INVBATCHRETURN]");
                    sbSql.AppendFormat(@" WHERE TA001='{0}' AND TA002='{1}' AND (NUM-USED+RETURNED)>0", ORIGINTA001, ORIGINTA002);
                    sbSql.AppendFormat(@" ");
                    sbSql.AppendFormat(@" ");

                    adapter7 = new SqlDataAdapter(@"" + sbSql, sqlConn);
                    sqlCmdBuilder7 = new SqlCommandBuilder(adapter7);

                    sqlConn.Open();
                    ds7.Clear();
                    adapter7.Fill(ds7, "ds7");
                    sqlConn.Close();


                    if (ds7.Tables["ds7"].Rows.Count == 0)
                    {
                        dataGridView6.DataSource = null;
                    }
                    else
                    {

                        dataGridView6.DataSource = ds7.Tables["ds7"];
                        dataGridView6.AutoResizeColumns();
                    }
                }
                else
                {

                }



            }
            catch
            {

            }
            finally
            {

            }
        }
        #endregion

        #region BUTTON

        private void button1_Click(object sender, EventArgs e)
        {
            SearchMOC();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            SearchINV();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            SearchINVBATCH();
        }

        private void button4_Click(object sender, EventArgs e)
        {
            TA002 = GETMAXTA002(TA001);
            textBox1.Text = TA002;
            //textBox1.Text = "20190702002";

            ADDTOTKWAREHOUSE();           
            ADDINVTATB();

            SEARCHINVBATCH2();
        }
        private void button5_Click(object sender, EventArgs e)
        {
            SEARCHINVBATCH3();
        }
        private void button6_Click(object sender, EventArgs e)
        {
            DELINVBATCHRETURN();
            ADDINVBATCHRETURN();

            SEARCHINVBATCHRETURN();
        }

        #endregion

        
    }
}

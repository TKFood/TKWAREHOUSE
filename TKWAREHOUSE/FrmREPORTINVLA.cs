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

            combobox1load();
        }

        #region FUNCTION
        public void combobox1load()
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
            Sequel.AppendFormat(@"SELECT MC001,MC002 FROM [TK].dbo.CMSMC WHERE MC001 IN ('20001','20017','21001') ORDER BY MC001");
            SqlDataAdapter da = new SqlDataAdapter(Sequel.ToString(), sqlConn);
            DataTable dt = new DataTable();
            sqlConn.Open();

            dt.Columns.Add("MC001", typeof(string));
            dt.Columns.Add("MC002", typeof(string));
            da.Fill(dt);
            comboBox1.DataSource = dt.DefaultView;
            comboBox1.ValueMember = "MC001";
            comboBox1.DisplayMember = "MC002";
            sqlConn.Close();



        }
        private void FrmREPORTINVLA_Load(object sender, EventArgs e)
        {
            dataGridView1.AlternatingRowsDefaultCellStyle.BackColor = Color.PaleTurquoise;      //奇數列顏色
            

            //先建立個 CheckBox 欄
            DataGridViewCheckBoxColumn cbCol = new DataGridViewCheckBoxColumn();
            cbCol.Width = 120;   //設定寬度
            cbCol.HeaderText = "選擇";
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

        private void cbHeader_CheckedChanged(object sender, EventArgs e)
        {
            dataGridView1.EndEdit();

            foreach (DataGridViewRow dr in dataGridView1.Rows)
            {
                dr.Cells[0].Value = ((CheckBox)dataGridView1.Controls.Find("checkboxHeader", true)[0]).Checked;

            }

        }

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

            string LA009 = comboBox1.SelectedValue.ToString();

            if (!string.IsNullOrEmpty(textBox1.Text))
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
                                    AND LA009 IN ('{2}')
                                    GROUP BY LA004,LA001,LA009,MB002,MB003,MB004
                                    UNION ALL
                                    SELECT '2' AS SERNO,'暫出單' AS 'KINDS',LA004,LA001,LA009,SUM(LA011) LA011,MB002,MB003,MB004
                                    FROM [TK].dbo.INVLA WITH(NOLOCK),[TK].dbo.CMSMQ WITH(NOLOCK),[TK].dbo.INVMB WITH(NOLOCK)
                                    WHERE LA006=MQ001
                                    AND LA001=MB001
                                    AND MQ003 IN ('13','14')
                                    AND LA005='-1'
                                    AND LA009 IN ('{2}')
                                    GROUP BY LA004,LA001,LA009,MB002,MB003,MB004
                                    UNION ALL
                                    SELECT '3' AS SERNO,'暫入單' AS 'KINDS',LA004,LA001,LA009,SUM(LA011) LA011,MB002,MB003,MB004
                                    FROM [TK].dbo.INVLA WITH(NOLOCK),[TK].dbo.CMSMQ WITH(NOLOCK),[TK].dbo.INVMB WITH(NOLOCK)
                                    WHERE LA006=MQ001
                                    AND LA001=MB001
                                    AND MQ003 IN ('15','16')
                                    AND LA005='-1'
                                    AND LA009 IN ('{2}')
                                    GROUP BY LA004,LA001,LA009,MB002,MB003,MB004
                                    UNION ALL
                                    SELECT '4' AS SERNO,'庫存異動單' AS 'KINDS',LA004,LA001,LA009,SUM(LA011) LA011,MB002,MB003,MB004
                                    FROM [TK].dbo.INVLA WITH(NOLOCK),[TK].dbo.CMSMQ WITH(NOLOCK),[TK].dbo.INVMB WITH(NOLOCK)
                                    WHERE LA006=MQ001
                                    AND LA001=MB001
                                    AND MQ003 IN ('11')
                                    AND LA005='-1'
                                    AND LA009 IN ('{2}')
                                    GROUP BY LA004,LA001,LA009,MB002,MB003,MB004
                                    UNION ALL
                                    SELECT '5' AS SERNO,'轉撥單' AS 'KINDS',LA004,LA001,LA009,SUM(LA011) LA011,MB002,MB003,MB004
                                    FROM [TK].dbo.INVLA WITH(NOLOCK),[TK].dbo.CMSMQ WITH(NOLOCK),[TK].dbo.INVMB WITH(NOLOCK)
                                    WHERE LA006=MQ001
                                    AND LA001=MB001
                                    AND MQ003 IN ('12','13')
                                    AND LA005='-1'
                                    AND LA009 IN ('{2}')
                                    GROUP BY LA004,LA001,LA009,MB002,MB003,MB004
                                    ) AS TEMP
                                    WHERE LA004='{0}'
                                    {1}
                                    ORDER BY LA004,LA001,SERNO,KINDS

                                    ", dateTimePicker1.Value.ToString("yyyyMMdd"), STRQUERY.ToString(), LA009);

            return FASTSQL.ToString();
        }

        public void SETFASTREPORT2()
        {

            string SQL;
            report1 = new Report();
            report1.Load(@"REPORT\成品倉撿料表明細.frx");

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
            SQL = SETFASETSQL2();
            Table.SelectCommand = SQL;
            report1.Preview = previewControl2;
            report1.Show(); 

        }

        public string SETFASETSQL2()
        {
            StringBuilder FASTSQL = new StringBuilder();
            StringBuilder STRQUERY = new StringBuilder();

            foreach (DataGridViewRow row in dataGridView1.Rows)
            {
                if (row.Cells[0].Value != null && (bool)row.Cells[0].Value&&!string.IsNullOrEmpty(row.Cells["單別"].Value.ToString())&& !string.IsNullOrEmpty(row.Cells["單號"].Value.ToString()))
                {
                    STRQUERY.AppendFormat("'"+row.Cells["單別"].Value.ToString().Trim()+ row.Cells["單號"].Value.ToString().Trim()+"'"+",");
                    
                }
            }
            STRQUERY.AppendFormat(@" ''");

          

            FASTSQL.AppendFormat(@"  
                                   SELECT LA004 AS '日期',LA001 AS '品號',LA009 AS '庫別', SUM(LA011) AS '數量',MB002 AS '品名',MB003 AS '規格',MB004 AS '單位'
                                    FROM (
                                    SELECT '1' AS SERNO,'銷貨單' AS 'KINDS',LA004,LA001,LA009,SUM(LA011) LA011,MB002,MB003,MB004,LA006,LA007
                                    FROM [TK].dbo.INVLA WITH(NOLOCK),[TK].dbo.CMSMQ WITH(NOLOCK),[TK].dbo.INVMB WITH(NOLOCK)
                                    WHERE LA006=MQ001
                                    AND LA001=MB001
                                    AND MQ003 IN ('23')
                                    AND LA005='-1'
                                    AND LA009 IN ('20001')
                                    GROUP BY LA004,LA001,LA009,MB002,MB003,MB004,LA006,LA007
                                    UNION ALL
                                    SELECT '2' AS SERNO,'暫出單' AS 'KINDS',LA004,LA001,LA009,SUM(LA011) LA011,MB002,MB003,MB004,LA006,LA007
                                    FROM [TK].dbo.INVLA WITH(NOLOCK),[TK].dbo.CMSMQ WITH(NOLOCK),[TK].dbo.INVMB WITH(NOLOCK)
                                    WHERE LA006=MQ001
                                    AND LA001=MB001
                                    AND MQ003 IN ('13','14')
                                    AND LA005='-1'
                                    AND LA009 IN ('20001')
                                    GROUP BY LA004,LA001,LA009,MB002,MB003,MB004,LA006,LA007
                                    UNION ALL
                                    SELECT '3' AS SERNO,'暫入單' AS 'KINDS',LA004,LA001,LA009,SUM(LA011) LA011,MB002,MB003,MB004,LA006,LA007
                                    FROM [TK].dbo.INVLA WITH(NOLOCK),[TK].dbo.CMSMQ WITH(NOLOCK),[TK].dbo.INVMB WITH(NOLOCK)
                                    WHERE LA006=MQ001
                                    AND LA001=MB001
                                    AND MQ003 IN ('15','16')
                                    AND LA005='-1'
                                    AND LA009 IN ('20001')
                                    GROUP BY LA004,LA001,LA009,MB002,MB003,MB004,LA006,LA007
                                    UNION ALL
                                    SELECT '4' AS SERNO,'庫存異動單' AS 'KINDS',LA004,LA001,LA009,SUM(LA011) LA011,MB002,MB003,MB004,LA006,LA007
                                    FROM [TK].dbo.INVLA WITH(NOLOCK),[TK].dbo.CMSMQ WITH(NOLOCK),[TK].dbo.INVMB WITH(NOLOCK)
                                    WHERE LA006=MQ001
                                    AND LA001=MB001
                                    AND MQ003 IN ('11')
                                    AND LA005='-1'
                                    AND LA009 IN ('20001')
                                    GROUP BY LA004,LA001,LA009,MB002,MB003,MB004,LA006,LA007
                                    UNION ALL
                                    SELECT '5' AS SERNO,'轉撥單' AS 'KINDS',LA004,LA001,LA009,SUM(LA011) LA011,MB002,MB003,MB004,LA006,LA007
                                    FROM [TK].dbo.INVLA WITH(NOLOCK),[TK].dbo.CMSMQ WITH(NOLOCK),[TK].dbo.INVMB WITH(NOLOCK)
                                    WHERE LA006=MQ001
                                    AND LA001=MB001
                                    AND MQ003 IN ('12','13')
                                    AND LA005='-1'
                                    AND LA009 IN ('20001')
                                    GROUP BY LA004,LA001,LA009,MB002,MB003,MB004,LA006,LA007
                                    ) AS TEMP
                                    WHERE LA004='{0}' 
                                    AND LTRIM(RTRIM(LA006))+LTRIM(RTRIM(LA007)) IN ({1})
                                    GROUP BY LA004,LA001,LA009,MB002,MB003,MB004
                                    ORDER BY LA001

                                    ", dateTimePicker1.Value.ToString("yyyyMMdd"), STRQUERY.ToString());

            return FASTSQL.ToString();
        }

        public void Search()
        {
            try
            {
                if (!string.IsNullOrEmpty(dateTimePicker1.Value.ToString("yyyyMMdd")) || !string.IsNullOrEmpty(dateTimePicker2.Value.ToString("yyyyMMdd")))
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
                    sbSql = SETsbSql();


                    adapter = new SqlDataAdapter(@"" + sbSql, sqlConn);
                    sqlCmdBuilder = new SqlCommandBuilder(adapter);

                    sqlConn.Open();
                    ds.Clear();
                    adapter.Fill(ds, "ds");
                    sqlConn.Close();


                    if (ds.Tables["ds"].Rows.Count == 0)
                    {
                        label14.Text = "找不到資料";
                    }
                    else
                    {

                        dataGridView1.DataSource = ds.Tables["ds"];
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


            STR.AppendFormat(@"  
                                SELECT KINDS AS '分類',LA006 AS '單別',LA007 AS '單號'
                                FROM (
                                SELECT '1' AS SERNO,'銷貨單' AS 'KINDS',LA004,LA001,LA009,SUM(LA011) LA011,MB002,MB003,MB004,LA006,LA007
                                FROM [TK].dbo.INVLA WITH(NOLOCK),[TK].dbo.CMSMQ WITH(NOLOCK),[TK].dbo.INVMB WITH(NOLOCK)
                                WHERE LA006=MQ001
                                AND LA001=MB001
                                AND MQ003 IN ('23')
                                AND LA005='-1'
                                AND LA009 IN ('20001')  
                                GROUP BY LA004,LA001,LA009,MB002,MB003,MB004,LA006,LA007
                                UNION ALL
                                SELECT '2' AS SERNO,'暫出單' AS 'KINDS',LA004,LA001,LA009,SUM(LA011) LA011,MB002,MB003,MB004,LA006,LA007
                                FROM [TK].dbo.INVLA WITH(NOLOCK),[TK].dbo.CMSMQ WITH(NOLOCK),[TK].dbo.INVMB WITH(NOLOCK)
                                WHERE LA006=MQ001
                                AND LA001=MB001
                                AND MQ003 IN ('13','14')
                                AND LA005='-1'
                                AND LA009 IN ('20001')
                                GROUP BY LA004,LA001,LA009,MB002,MB003,MB004,LA006,LA007
                                UNION ALL
                                SELECT '3' AS SERNO,'暫入單' AS 'KINDS',LA004,LA001,LA009,SUM(LA011) LA011,MB002,MB003,MB004,LA006,LA007
                                FROM [TK].dbo.INVLA WITH(NOLOCK),[TK].dbo.CMSMQ WITH(NOLOCK),[TK].dbo.INVMB WITH(NOLOCK)
                                WHERE LA006=MQ001
                                AND LA001=MB001
                                AND MQ003 IN ('15','16')
                                AND LA005='-1'
                                AND LA009 IN ('20001')
                                GROUP BY LA004,LA001,LA009,MB002,MB003,MB004,LA006,LA007
                                UNION ALL
                                SELECT '4' AS SERNO,'庫存異動單' AS 'KINDS',LA004,LA001,LA009,SUM(LA011) LA011,MB002,MB003,MB004,LA006,LA007
                                FROM [TK].dbo.INVLA WITH(NOLOCK),[TK].dbo.CMSMQ WITH(NOLOCK),[TK].dbo.INVMB WITH(NOLOCK)
                                WHERE LA006=MQ001
                                AND LA001=MB001
                                AND MQ003 IN ('11')
                                AND LA005='-1'
                                AND LA009 IN ('20001')
                                GROUP BY LA004,LA001,LA009,MB002,MB003,MB004,LA006,LA007
                                UNION ALL
                                SELECT '5' AS SERNO,'轉撥單' AS 'KINDS',LA004,LA001,LA009,SUM(LA011) LA011,MB002,MB003,MB004,LA006,LA007
                                FROM [TK].dbo.INVLA WITH(NOLOCK),[TK].dbo.CMSMQ WITH(NOLOCK),[TK].dbo.INVMB WITH(NOLOCK)
                                WHERE LA006=MQ001
                                AND LA001=MB001
                                AND MQ003 IN ('12','13')
                                AND LA005='-1'
                                AND LA009 IN ('20001') 
                                GROUP BY LA004,LA001,LA009,MB002,MB003,MB004,LA006,LA007
                                ) AS TEMP
                                WHERE LA004='{0}' 
                                GROUP BY KINDS,LA006,LA007
                                ORDER BY KINDS,LA006,LA007
                            ", dateTimePicker2.Value.ToString("yyyyMMdd"));

            return STR;
        }

       

        #endregion

        #region BUTTON

        private void button2_Click(object sender, EventArgs e)
        {
            SETFASTREPORT();
        }
        private void button3_Click(object sender, EventArgs e)
        {
            Search();
        }
        private void button1_Click(object sender, EventArgs e)
        {
            SETFASTREPORT2();
        }
        #endregion


    }
}

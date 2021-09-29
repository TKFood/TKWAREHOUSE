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
    public partial class FrmREPORTPURIN : Form
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

        public FrmREPORTPURIN()
        {
            InitializeComponent();
        }

        #region FUNCTION
        private void FrmREPORTPURIN_Load(object sender, EventArgs e)
        {
            dataGridView1.AlternatingRowsDefaultCellStyle.BackColor = Color.PaleTurquoise;      //奇數列顏色

            //先建立個 CheckBox 欄
            DataGridViewCheckBoxColumn cbCol = new DataGridViewCheckBoxColumn();
            cbCol.Width = 100;   //設定寬度
            cbCol.HeaderText = "　全選";
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
            foreach (DataGridViewRow dr in dataGridView1.Rows)
            {
                dr.Cells[0].Value = ((CheckBox)dataGridView1.Controls.Find("checkboxHeader", true)[0]).Checked;

            }



        }

        public void Search()
        {
            try
            {
                if (!string.IsNullOrEmpty(dateTimePicker1.Value.ToString("yyyyMMdd")))
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
                    adapter.Fill(ds, "ds1");
                    sqlConn.Close();


                    if (ds.Tables["ds1"].Rows.Count == 0)
                    {
                        dataGridView1.DataSource = null;
                    }
                    else
                    {
                        dataGridView1.DataSource = ds.Tables["ds1"];
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
                                 SELECT TD001 AS '採購單別',TD002 AS '採購單號',TD003 AS '序號',TD004 AS '品號',TD005 AS '品名',TD006 AS '規格',TD012 AS '預交日',TC004 AS '供應廠',MA002 AS '供應廠商'
                                    ,SUBSTRING(TD012,1,4) AS '年',SUBSTRING(TD012,5,2) AS '月',SUBSTRING(TD012,7,2) AS '日'
                                    FROM [TK].dbo.PURTC,[TK].dbo.PURTD,[TK].dbo.PURMA
                                    WHERE TC001=TD001 AND TC002=TD002
                                    AND TC004=MA001
                                    AND TD012>='{0}' AND TD012<='{0}'
                                    ORDER BY TD001,TD002,TD003
                                    ", dateTimePicker1.Value.ToString("yyyyMMdd"));

            return STR;
        }

        public void SETFASTREPORT()
        {
            string SQL;
            report1 = new Report();

            report1.Load(@"REPORT\進貨入庫單.frx");

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
            string TD001TD002TD003 = FINDCHECKED();

            StringBuilder FASTSQL = new StringBuilder();
            StringBuilder STRQUERY = new StringBuilder();

            FASTSQL.AppendFormat(@"   
                                 SELECT TD001 AS '採購單別',TD002 AS '採購單號',TD003 AS '序號',TD004 AS '品號',TD005 AS '品名',TD006 AS '規格',TD012 AS '預交日',TC004 AS '供應廠',MA002 AS '供應廠商'
                                    ,SUBSTRING(TD012,1,4) AS '年',SUBSTRING(TD012,5,2) AS '月',SUBSTRING(TD012,7,2) AS '日'
                                    FROM [TK].dbo.PURTC,[TK].dbo.PURTD,[TK].dbo.PURMA
                                    WHERE TC001=TD001 AND TC002=TD002
                                    AND TC004=MA001
                                    AND TD001+TD002+TD003 IN ({0})
                                    ORDER BY TD012,TD004  
                                    ", TD001TD002TD003);



            return FASTSQL.ToString();

        }

        public string FINDCHECKED()
        {
            StringBuilder TD001TD002TD003 = new StringBuilder();

            foreach (DataGridViewRow dr in this.dataGridView1.Rows)
            {
                String TD001 = ((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[0].ToString().Trim();
                String TD002 = ((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[1].ToString().Trim();
                String TD003 = ((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[2].ToString().Trim();

                if (dr.Cells[0].Value != null && (bool)dr.Cells[0].Value)
                {
                    TD001TD002TD003.AppendFormat(@" '{0}',", TD001+ TD002+ TD003);
                }
            }

            TD001TD002TD003.AppendFormat(@" '' ");

            //MessageBox.Show(TD001TD002TD003.ToString());

            return TD001TD002TD003.ToString();
        }
        #endregion

        #region BUTTON
        private void button1_Click(object sender, EventArgs e)
        {
            Search();
        }
        private void button2_Click(object sender, EventArgs e)
        {
            //FINDCHECKED();

            SETFASTREPORT();
        }


        #endregion

    
    }
}

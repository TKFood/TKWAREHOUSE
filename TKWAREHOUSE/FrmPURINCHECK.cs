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
using System.Configuration;
using System.Text.RegularExpressions;
using FastReport;
using FastReport.Data;
using TKITDLL;
using System.Data.OleDb;
using System.Net;
using AForge.Video;
using AForge.Video.DirectShow;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;
using System.Threading;
using System.IO.Ports;
using System.Threading;
using System.IO.Ports;


namespace TKWAREHOUSE
{
    public partial class FrmPURINCHECK : Form
    {
        int CommandTimeout = 180;
        StringBuilder sbSql = new StringBuilder();
        SqlConnection sqlConn = new SqlConnection();
        SqlDataAdapter adapter = new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilder = new SqlCommandBuilder();
        SqlTransaction tran;
        SqlCommand cmd = new SqlCommand();
        DataSet ds = new DataSet();
        DataTable dt = new DataTable();
        string tablename = null;
        int rownum = 0;
        DataGridViewRow row;
        int result;
        public FrmPURINCHECK()
        {
            InitializeComponent();
        }

        public FrmPURINCHECK(string ID)
        {
            InitializeComponent();

            textBox1.Text = ID;
        }
        private void FrmPURINCHECK_Load(object sender, EventArgs e)
        {
            SETDATE();
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
        public void SETDATE()
        {
            DateTime today = DateTime.Today;
            // 當月第一天
            DateTime firstDay = new DateTime(today.Year, today.Month, 1);
            // 當月最後一天
            DateTime lastDay = new DateTime(today.Year, today.Month, DateTime.DaysInMonth(today.Year, today.Month));

            dateTimePicker1.Value = firstDay;
            dateTimePicker2.Value = lastDay;
        }

        public void Search(string SDATES,string EDATES)
        {
            StringBuilder SLQURY = new StringBuilder();
            StringBuilder SLQURY2 = new StringBuilder();

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

                SLQURY.Clear();

               
                sbSql.AppendFormat(@" 
                                    SELECT 
                                    MA002  AS '廠商'
                                    ,TD012  AS '預計到貨日'
                                    ,TD005  AS '品名'
                                    ,TD006  AS '規格'
                                    ,TD008  AS '採購量'
                                    ,TD009  AS '單位'
                                    ,(TD008-TD015-ISNULL(TEMP.TH007,0)) AS '未到貨量'
                                    ,TD015  AS '已到貨'
                                    ,ISNULL(TEMP.TH007,0) AS '已入庫'
                                    ,TC001  AS '採購單別'
                                    ,TC002  AS '採購單號'
                                    ,TD003  AS '序號'
                                    ,TD004  AS '品號'

                                    FROM [TK].dbo.PURTC,[TK].dbo.PURTD
                                    LEFT JOIN 
                                    (SELECT TH011,TH012,TH013,TH004,SUM(TH007) AS TH007
                                    FROM [TK].dbo.PURTG,[TK].dbo.PURTH
                                    WHERE TG001=TH001 AND TG002=TH002
                                    AND TG013 IN ('Y','N')
                                    GROUP BY TH011,TH012,TH013,TH004
                                    ) AS TEMP ON TH011=TD001 AND TH012=TD002 AND TH013=TD003
                                    ,[TK].dbo.PURMA
                                    WHERE TC001=TD001 AND TC002=TD002
                                    AND MA001=TC004
                                    AND TC014='Y'
                                    AND TD016='N'
                                    AND TD008>0
                                    AND TD008-TD015-ISNULL(TEMP.TH007,0)>0
                                    AND TD012>='20250201'
                                    AND TD012<='20250228'
                                    ORDER BY MA002,TD012
                                    
                                    ");


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

                        //dataGridView1.ColumnHeadersDefaultCellStyle.Font = new Font("Tahoma", 9);
                        //dataGridView1.DefaultCellStyle.Font = new Font("Tahoma", 10);
                        dataGridView1.Columns["廠商"].Width = 100;
                        dataGridView1.Columns["預計到貨日"].Width = 100;
                        dataGridView1.Columns["品名"].Width = 200;
                        dataGridView1.Columns["規格"].Width = 100;
                        dataGridView1.Columns["採購量"].Width = 100;
                        dataGridView1.Columns["單位"].Width = 60;
                        dataGridView1.Columns["未到貨量"].Width = 100;
                        dataGridView1.Columns["已到貨"].Width = 100;
                        dataGridView1.Columns["已入庫"].Width = 100;
                        dataGridView1.Columns["採購單別"].Width = 100;
                        dataGridView1.Columns["採購單號"].Width = 100;
                        dataGridView1.Columns["序號"].Width = 100;
                        dataGridView1.Columns["品號"].Width = 100;

                        dataGridView1.Columns["採購量"].DefaultCellStyle.Format = "#,##0.000";
                        dataGridView1.Columns["採購量"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                        dataGridView1.Columns["未到貨量"].DefaultCellStyle.Format = "#,##0.000";
                        dataGridView1.Columns["未到貨量"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                        dataGridView1.Columns["已到貨"].DefaultCellStyle.Format = "#,##0.000";
                        dataGridView1.Columns["已到貨"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                        dataGridView1.Columns["已入庫"].DefaultCellStyle.Format = "#,##0.000";
                        dataGridView1.Columns["已入庫"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
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

        public void ADD()
        {
            foreach (DataGridViewRow dr in this.dataGridView1.Rows)
            {
                if (dr.Cells[0].Value != null && (bool)dr.Cells[0].Value)
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
                
                                            ");


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
        #endregion


        #region BUTTON
        private void button1_Click(object sender, EventArgs e)
        {
            Search(dateTimePicker1.Value.ToString("yyyyMMdd"), dateTimePicker2.Value.ToString("yyyyMMdd"));
        }
        private void button2_Click(object sender, EventArgs e)
        {

        }
        #endregion


    }
}

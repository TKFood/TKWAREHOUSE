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

            combobox1load();
            combobox2load();
            combobox3load();
            combobox4load();
        }
        #region FUNCTION
        public void SETGRIDVIEW()
        {
            dataGridView1.AlternatingRowsDefaultCellStyle.BackColor = Color.PaleTurquoise;      //奇數列顏色
            dataGridView2.AlternatingRowsDefaultCellStyle.BackColor = Color.PaleTurquoise;      //奇數列顏色

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
            Sequel.AppendFormat(@"
                                SELECT
                                [ID]
                                ,[KINDS]
                                ,[NAMES]
                                ,[KEYS]
                                ,[KEYS2]
                                FROM [TKWAREHOUSE].[dbo].[TBPARAS]
                                WHERE [KINDS]='FrmPURINCHECK'");
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

        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
            textBox4.Text = comboBox2.SelectedValue.ToString();
        }
        public void combobox2load()
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
            Sequel.AppendFormat(@"SELECT MC001,MC002 FROM [TK].dbo.CMSMC WHERE MC001 LIKE '2%' ORDER BY MC001");
            SqlDataAdapter da = new SqlDataAdapter(Sequel.ToString(), sqlConn);
            DataTable dt = new DataTable();
            sqlConn.Open();

            dt.Columns.Add("MC001", typeof(string));
            dt.Columns.Add("MC002", typeof(string));
            da.Fill(dt);
            comboBox2.DataSource = dt.DefaultView;
            comboBox2.ValueMember = "MC001";
            comboBox2.DisplayMember = "MC002";
            sqlConn.Close();
        }

        public void combobox3load()
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
                                SELECT
                                [ID]
                                ,[KINDS]
                                ,[NAMES]
                                ,[KEYS]
                                ,[KEYS2]
                                FROM [TKWAREHOUSE].[dbo].[TBPARAS]
                                WHERE [KINDS]='FrmPURINCHECKNAMES'
                                ORDER BY [NAMES]
                                ");
            SqlDataAdapter da = new SqlDataAdapter(Sequel.ToString(), sqlConn);
            DataTable dt = new DataTable();
            sqlConn.Open();

            dt.Columns.Add("NAMES", typeof(string));

            da.Fill(dt);
            comboBox3.DataSource = dt.DefaultView;
            comboBox3.ValueMember = "NAMES";
            comboBox3.DisplayMember = "NAMES";
            sqlConn.Close();
        }
        public void combobox4load()
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
                                SELECT
                                [ID]
                                ,[KINDS]
                                ,[NAMES]
                                ,[KEYS]
                                ,[KEYS2]
                                FROM [TKWAREHOUSE].[dbo].[TBPARAS]
                                WHERE [KINDS]='FrmPURINCHECKSTATUS'
                                ORDER BY [KEYS2]");
            SqlDataAdapter da = new SqlDataAdapter(Sequel.ToString(), sqlConn);
            DataTable dt = new DataTable();
            sqlConn.Open();

            dt.Columns.Add("NAMES", typeof(string));

            da.Fill(dt);
            comboBox4.DataSource = dt.DefaultView;
            comboBox4.ValueMember = "NAMES";
            comboBox4.DisplayMember = "NAMES";
            sqlConn.Close();
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

        public void Search(string SDATES, string EDATES, string STATUS)
        {
            DataSet ds = new DataSet();
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

                if (STATUS.Equals("未進"))
                {
                    SLQURY.AppendFormat(@"
                                        AND PURTC.[TC001]+PURTC.[TC002]+PURTD.[TD003] NOT IN (SELECT [TC001]+[TC002]+[TD003] FROM [TKWAREHOUSE].[dbo].[TBPURINCHECK])
                                        ");
                }
                else if (STATUS.Equals("已進"))
                {
                    SLQURY.AppendFormat(@"
                                        AND PURTC.[TC001]+PURTC.[TC002]+PURTD.[TD003]  IN (SELECT [TC001]+[TC002]+[TD003] FROM [TKWAREHOUSE].[dbo].[TBPURINCHECK])
                                        ");
                }
                else if (STATUS.Equals("全部"))
                {
                    SLQURY.AppendFormat(@"
                                        
                                        ");
                }

                //採購單已核或未核
                //且結案碼=N
                sbSql.AppendFormat(@" 
                                SELECT 
                                (CASE WHEN [TBPURINCHECK].NUMS>0 THEN [TBPURINCHECK].NUMS ELSE  TD008 END) AS '收貨數量'
                                ,INNAMES  AS '收貨人員'
                                ,[INVOICES] AS '發票'
                                ,[INNO] AS '貨單'
                                ,PURMA.MA002  AS '廠商'
                                ,TD012  AS '預計到貨日'
                                ,PURTD.TD005  AS '品名'
                                ,TD006  AS '規格'
                                ,TD008  AS '採購量'
                                ,TD007  AS '庫別'
                                ,TD009  AS '單位'
                                ,(TD008-ISNULL(TEMP.TH007,0)) AS '還未到貨量'
                                ,ISNULL(TEMP.TH007,0) AS '已進貨單量'
                                ,PURTC.TC001  AS '採購單別'
                                ,PURTC.TC002  AS '採購單號'
                                ,PURTD.TD003  AS '序號'
                                ,PURTD.TD004  AS '品號'
                                , [TBPURINCHECK].NUMS
                                FROM [TK].dbo.PURTC,[TK].dbo.PURTD
                                LEFT JOIN [TKWAREHOUSE].[dbo].[TBPURINCHECK] ON [TBPURINCHECK].TC001+[TBPURINCHECK].TC002+[TBPURINCHECK].TD003=PURTD.TD001+PURTD.TD002+PURTD.TD003
                                LEFT JOIN 
                                (SELECT TH011,TH012,TH013,TH004,SUM(TH007) AS TH007
                                FROM [TK].dbo.PURTG,[TK].dbo.PURTH
                                WHERE TG001=TH001 AND TG002=TH002
                                AND TG013 IN ('Y','N')
                                GROUP BY TH011,TH012,TH013,TH004
                                ) AS TEMP ON TH011=PURTD.TD001 AND TH012=PURTD.TD002 AND TH013=PURTD.TD003
                                ,[TK].dbo.PURMA

                                WHERE PURTC.TC001=PURTD.TD001 AND PURTC.TC002=PURTD.TD002
                                AND MA001=TC004
                                AND TC014='Y'
                                AND TD016='N'
                                AND TD008>0                               
                                AND TD012>='{0}'
                                AND TD012<='{1}'
                                {2}

                                ORDER BY PURMA.MA002,TD012
                                    
                                ", SDATES, EDATES, SLQURY.ToString());


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
                        dataGridView1.Columns["庫別"].Width = 60;
                        dataGridView1.Columns["還未到貨量"].Width = 100;
                        dataGridView1.Columns["已進貨單量"].Width = 100;
                        dataGridView1.Columns["採購單別"].Width = 100;
                        dataGridView1.Columns["採購單號"].Width = 100;
                        dataGridView1.Columns["序號"].Width = 100;
                        dataGridView1.Columns["品號"].Width = 100;

                        dataGridView1.Columns["採購量"].DefaultCellStyle.Format = "#,##0.000";
                        dataGridView1.Columns["採購量"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                        dataGridView1.Columns["還未到貨量"].DefaultCellStyle.Format = "#,##0.000";
                        dataGridView1.Columns["還未到貨量"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                        dataGridView1.Columns["已進貨單量"].DefaultCellStyle.Format = "#,##0.000";
                        dataGridView1.Columns["已進貨單量"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
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

        public void ADD_TBPURINCHECK(
            string TC001,
            string TC002,
            string TD003,
            string TD004,
            string TD005,
            string NUMS,
            string MA002,
            string STOCKS,
            string ISIN,
            string INVOICES,
            string INNO,
            string INNAMES
            )
        {
            foreach (DataGridViewRow dr in this.dataGridView1.Rows)
            {
                if (dr.Cells[0].Value != null && (bool)dr.Cells[0].Value)
                {
                    //MessageBox.Show(dr.Cells["收貨數量"].Value.ToString()); 
                    TC001 = dr.Cells["採購單別"].Value.ToString();
                    TC002 = dr.Cells["採購單號"].Value.ToString();
                    TD003 = dr.Cells["序號"].Value.ToString();
                    TD004 = dr.Cells["品號"].Value.ToString();
                    TD005 = dr.Cells["品名"].Value.ToString();
                    NUMS = dr.Cells["收貨數量"].Value.ToString();
                    MA002 = dr.Cells["廠商"].Value.ToString();
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
                        //dr.Cells["單別"].Value.ToString()
                        sbSql.AppendFormat(@"
                                            INSERT INTO [TKWAREHOUSE].[dbo].[TBPURINCHECK]
                                            (
                                            [TC001]
                                            ,[TC002]
                                            ,[TD003]
                                            ,[TD004]
                                            ,[TD005]
                                            ,[NUMS]
                                            ,[MA002]
                                            ,[STOCKS]
                                            ,[ISIN]
                                            ,[INVOICES]
                                            ,[INNO]
                                            ,[INNAMES]
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
                                            ,'{8}'
                                            ,'{9}'
                                            ,'{10}'
                                            ,'{11}'
                                            )

                                            ", TC001
                                            , TC002
                                            , TD003
                                            , TD004
                                            , TD005
                                            , NUMS
                                            , MA002
                                            , STOCKS
                                            , ISIN
                                            , INVOICES
                                            , INNO
                                            , INNAMES
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
            }
        }

        public void UPDATE_TBPURINCHECK(
            string TC001,
            string TC002,
            string TD003,
            string TD004,
            string TD005,
            string NUMS,
            string MA002,
            string STOCKS,
            string ISIN,
            string INVOICES,
            string INNO,
            string INNAMES
            )
        {
            foreach (DataGridViewRow dr in this.dataGridView1.Rows)
            {
                if (dr.Cells[0].Value != null && (bool)dr.Cells[0].Value)
                {
                    //MessageBox.Show(dr.Cells["收貨數量"].Value.ToString()); 
                    TC001 = dr.Cells["採購單別"].Value.ToString();
                    TC002 = dr.Cells["採購單號"].Value.ToString();
                    TD003 = dr.Cells["序號"].Value.ToString();
                    TD004 = dr.Cells["品號"].Value.ToString();
                    TD005 = dr.Cells["品名"].Value.ToString();
                    NUMS = dr.Cells["收貨數量"].Value.ToString();
                    MA002 = dr.Cells["廠商"].Value.ToString();
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
                        //dr.Cells["單別"].Value.ToString()
                        sbSql.AppendFormat(@"
                                           UPDATE [TKWAREHOUSE].[dbo].[TBPURINCHECK]
                                            SET [NUMS]='{3}',[INVOICES]='{4}' ,[INNO]='{5}',[INNAMES]='{6}'
                                            WHERE [TC001]='{0}' AND [TC002]='{1}' AND [TD003]='{2}'

                                            ", TC001
                                            , TC002
                                            , TD003
                                            , NUMS
                                            , INVOICES
                                            , INNO
                                            , INNAMES
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
            }
        }


        public void SEARCH_TBPURINCHECK(string TC002)
        {
            DataSet ds = new DataSet();
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
                                    CONVERT(NVARCHAR,[INDATES],112) AS '收貨日'
                                    ,[TC001] AS '採購單別'
                                    ,[TC002] AS '採購單號'
                                    ,[TD003] AS '採購序號'
                                    ,[TD004] AS '品號'
                                    ,[TD005] AS '品名'
                                    ,[NUMS] AS '到貨數量'
                                    ,[MA002] AS '廠商'
                                    ,[STOCKS] AS '庫別'
                                    ,[ISIN] AS '是否到貨'
                                    ,[INVOICES] AS '發票'
                                    ,[INNO] AS '貨單'
                                    ,[INNAMES] AS '收貨人員'
                                    , [ID]
                                    FROM [TKWAREHOUSE].[dbo].[TBPURINCHECK]
                                    WHERE 1=1
                                    AND [TC002] LIKE '%{0}%'
                                    ORDER BY [TC001],[TC002]
                                    
                                ", TC002);


                adapter = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder = new SqlCommandBuilder(adapter);
                sqlConn.Open();
                ds.Clear();
                adapter.Fill(ds, "TEMPds");
                sqlConn.Close();

                dataGridView2.DataSource = null;

                if (ds.Tables["TEMPds"].Rows.Count >= 1)
                {
                    dataGridView2.DataSource = ds.Tables["TEMPds"];
                    dataGridView2.AutoResizeColumns();

                    //dataGridView1.ColumnHeadersDefaultCellStyle.Font = new Font("Tahoma", 9);
                    //dataGridView1.DefaultCellStyle.Font = new Font("Tahoma", 10);

                    dataGridView2.Columns["收貨日"].Width = 100;
                    dataGridView2.Columns["採購單別"].Width = 100;
                    dataGridView2.Columns["採購單號"].Width = 100;
                    dataGridView2.Columns["採購序號"].Width = 60;
                    dataGridView2.Columns["品號"].Width = 100;
                    dataGridView2.Columns["品名"].Width = 200;
                    dataGridView2.Columns["到貨數量"].Width = 100;
                    dataGridView2.Columns["廠商"].Width = 100;
                    dataGridView2.Columns["庫別"].Width = 60;

                    dataGridView2.Columns["到貨數量"].DefaultCellStyle.Format = "#,##0.000";
                    dataGridView2.Columns["到貨數量"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;

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
            textBox6.Text = null;

            if (dataGridView2.CurrentRow != null)
            {
                int rowindex = dataGridView2.CurrentRow.Index;
                if (rowindex >= 0)
                {
                    DataGridViewRow row = dataGridView2.Rows[rowindex];

                    textBox6.Text = row.Cells["ID"].Value.ToString().Trim();


                    //SEARCH2(row.Cells["品號"].Value.ToString().Trim());
                    //SEARCH3(row.Cells["品號"].Value.ToString().Trim());

                    //SETFASTREPORT(row.Cells["品號"].Value.ToString().Trim());
                }
            }
        }

        public void DELETE_TBPURINCHECK(string ID)
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
                //dr.Cells["單別"].Value.ToString()
                sbSql.AppendFormat(@"
                                    DELETE [TKWAREHOUSE].[dbo].[TBPURINCHECK]                                            
                                    WHERE [ID]='{0}' 

                                    ", ID

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
    
        public void Search_TBPURINCHECKPLAN(string SDATES,string EDATES)
        {
            DataSet ds = new DataSet();
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
                                    CONVERT(NVARCHAR,[PLANINDATE],112 ) AS '預計到貨日'
                                    ,PURMA.MA002 AS '廠商'
                                    ,[TBPURINCHECKPLAN].[TC001] AS '採購單別'
                                    ,[TBPURINCHECKPLAN].[TC002] AS '採購單號'
                                    ,[TD003] AS '採購序號'
                                    ,[TD004] AS '品號'
                                    ,[TD005] AS '品名'
                                    ,[PLANNUMS] AS '預計到貨數量'
                                    ,[ID]

                                    FROM [TKWAREHOUSE].[dbo].[TBPURINCHECKPLAN]
                                    LEFT JOIN [TK].dbo.PURTC ON PURTC.TC001=[TBPURINCHECKPLAN].TC001 AND PURTC.TC002=[TBPURINCHECKPLAN].TC002
                                    LEFT JOIN [TK].dbo.PURMA ON PURTC.TC004=PURMA.MA001          
                                    WHERE 1=1
                                    AND CONVERT(NVARCHAR,[PLANINDATE],112 )>='{0}' AND CONVERT(NVARCHAR,[PLANINDATE],112 )<='{0}'
                                    ORDER BY [TBPURINCHECKPLAN].[TC001],[TBPURINCHECKPLAN].[TC002],[TD004]
                                    
                                ", SDATES, EDATES);


                adapter = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder = new SqlCommandBuilder(adapter);
                sqlConn.Open();
                ds.Clear();
                adapter.Fill(ds, "TEMPds");
                sqlConn.Close();

                dataGridView3.DataSource = null;

                if (ds.Tables["TEMPds"].Rows.Count >= 1)
                {
                    dataGridView3.DataSource = ds.Tables["TEMPds"];
                    dataGridView3.AutoResizeColumns();

                    //dataGridView1.ColumnHeadersDefaultCellStyle.Font = new Font("Tahoma", 9);
                    //dataGridView1.DefaultCellStyle.Font = new Font("Tahoma", 10);

                    dataGridView3.Columns["預計到貨日"].Width = 100;
                    dataGridView3.Columns["廠商"].Width = 100;
                    dataGridView3.Columns["採購單別"].Width = 100;
                    dataGridView3.Columns["採購單號"].Width = 100;
                    dataGridView3.Columns["採購序號"].Width = 60;
                    dataGridView3.Columns["品號"].Width = 100;
                    dataGridView3.Columns["品名"].Width = 200;
                    dataGridView3.Columns["預計到貨數量"].Width = 100;

                    dataGridView3.Columns["預計到貨數量"].DefaultCellStyle.Format = "#,##0.000";
                    dataGridView3.Columns["預計到貨數量"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;

                }

            }
            catch
            {

            }
            finally
            {

            }
        }
        private void textBox7_TextChanged(object sender, EventArgs e)
        {
            if(!string.IsNullOrEmpty(textBox7.Text)&& !string.IsNullOrEmpty(textBox8.Text))
            {
                string TC001 = textBox7.Text.Trim();
                string TC002 = textBox8.Text.Trim();
                string TD003 = textBox9.Text.Trim();

                FIND_MA002(TC001, TC002);
                FIND_MB002(TC001, TC002, TD003);
            }
        }
        private void textBox8_TextChanged(object sender, EventArgs e)
        {
            if (!string.IsNullOrEmpty(textBox7.Text) && !string.IsNullOrEmpty(textBox8.Text))
            {
                string TC001= textBox7.Text.Trim();
                string TC002 = textBox8.Text.Trim();
                string TD003 = textBox9.Text.Trim();

                FIND_MA002(TC001, TC002);
                FIND_MB002(TC001, TC002, TD003);
            }
        }
        private void textBox9_TextChanged(object sender, EventArgs e)
        {
            if (!string.IsNullOrEmpty(textBox7.Text) && !string.IsNullOrEmpty(textBox8.Text) && !string.IsNullOrEmpty(textBox9.Text))
            {
                string TC001 = textBox7.Text.Trim();
                string TC002 = textBox8.Text.Trim();
                string TD003 = textBox9.Text.Trim();
                FIND_MB002(TC001, TC002, TD003);
            }
        }

        public void FIND_MA002(string TC001,string TC002)
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
                                    MA002
                                    FROM [TK].dbo.PURTC,[TK].dbo.PURMA
                                    WHERE TC004=MA001
                                    AND TC001='{0}' AND TC002='{1}'
                                    
                                ", TC001, TC002);


                adapter = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder = new SqlCommandBuilder(adapter);
                sqlConn.Open();
                ds.Clear();
                adapter.Fill(ds, "TEMPds");
                sqlConn.Close();

                if (ds.Tables["TEMPds"].Rows.Count >= 1)
                {
                    textBox13.Text = ds.Tables["TEMPds"].Rows[0]["MA002"].ToString();
                }

            }
            catch
            {

            }
            finally
            {

            }
        }

        public void FIND_MB002(string TC001, string TC002,string TD003)
        {
            StringBuilder SLQURY = new StringBuilder();
            StringBuilder SLQURY2 = new StringBuilder();

            textBox10.Text = "";
            textBox11.Text = "";
            textBox12.Text = "";
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
                                    TD004,TD005,TD008          
                                    FROM [TK].dbo.PURTC,[TK].dbo.PURTD
                                    WHERE TC001=TD001 AND TC002=TD002
                                    AND TC001='{0}' AND TC002='{1}' AND TD003='{2}'
                                    
                                ", TC001, TC002, TD003);


                adapter = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder = new SqlCommandBuilder(adapter);
                sqlConn.Open();
                ds.Clear();
                adapter.Fill(ds, "TEMPds");
                sqlConn.Close();

                if (ds.Tables["TEMPds"].Rows.Count >= 1)
                {
                    textBox10.Text = ds.Tables["TEMPds"].Rows[0]["TD004"].ToString();
                    textBox11.Text = ds.Tables["TEMPds"].Rows[0]["TD005"].ToString();
                    textBox12.Text = ds.Tables["TEMPds"].Rows[0]["TD008"].ToString();
                }

            }
            catch
            {

            }
            finally
            {

            }
        }

        public void ADD_TBPURINCHECKPLAN(
            string TC001,
            string TC002,
            string TD003,
            string TD004,
            string TD005,
            string PLANINDATE,
            string PLANNUMS
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
                //dr.Cells["單別"].Value.ToString()
                sbSql.AppendFormat(@"
                                    INSERT INTO [TKWAREHOUSE].[dbo].[TBPURINCHECKPLAN]
                                    (
                                    TC001
                                    ,TC002
                                    ,TD003
                                    ,TD004
                                    ,TD005
                                    ,PLANINDATE
                                    ,PLANNUMS
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
                                    )

                                    ", TC001
                                    , TC002
                                    , TD003
                                    , TD004
                                    , TD005
                                    , PLANINDATE
                                    , PLANNUMS

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
        private void dataGridView3_SelectionChanged(object sender, EventArgs e)
        {
            textBox21.Text = null;
            textBox14.Text = null;
            textBox15.Text = null;
            textBox16.Text = null;
            textBox17.Text = null;
            textBox18.Text = null;
            textBox19.Text = null;
            textBox20.Text = null;

            if (dataGridView3.CurrentRow != null)
            {
                int rowindex = dataGridView3.CurrentRow.Index;
                if (rowindex >= 0)
                {
                    DataGridViewRow row = dataGridView3.Rows[rowindex];

                    textBox21.Text = row.Cells["ID"].Value.ToString().Trim();
                    textBox14.Text = row.Cells["採購單別"].Value.ToString().Trim();
                    textBox15.Text = row.Cells["採購單號"].Value.ToString().Trim();
                    textBox16.Text = row.Cells["採購序號"].Value.ToString().Trim();
                    textBox17.Text = row.Cells["品號"].Value.ToString().Trim();
                    textBox18.Text = row.Cells["品名"].Value.ToString().Trim();
                    textBox19.Text = row.Cells["預計到貨數量"].Value.ToString().Trim();
                    textBox20.Text = row.Cells["廠商"].Value.ToString().Trim();

                    if (row.Cells["預計到貨日"].Value != null)
                    {
                        DateTime tempDate;
                        string dateStr = row.Cells["預計到貨日"].Value.ToString().Trim();

                        if (dateStr.Length == 8 && DateTime.TryParseExact(dateStr, "yyyyMMdd", null, System.Globalization.DateTimeStyles.None, out tempDate))
                        {
                            dateTimePicker5.Value = tempDate;
                        }
                        else
                        {
                            // 若轉換失敗，處理預設值，例如設為當前日期
                            dateTimePicker5.Value = DateTime.Today;
                        }
                    }
                }
            }
        }
        public void DELETE_TBPURINCHECKPLAN(string ID)
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
                //dr.Cells["單別"].Value.ToString()
                sbSql.AppendFormat(@"
                                    DELETE [TKWAREHOUSE].[dbo].[TBPURINCHECKPLAN]
                                    where [ID]='{0}'                                    

                                    ", ID                                   

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
        public void UPDATE_TBPURINCHECKPLAN(
               string ID,
               string TC001,
               string TC002,
               string TD003,
               string TD004,
               string TD005,
               string PLANINDATE,
               string PLANNUMS
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
                //dr.Cells["單別"].Value.ToString()
                sbSql.AppendFormat(@"
                                    UPDATE [TKWAREHOUSE].[dbo].[TBPURINCHECKPLAN]
                                    SET
                                    TC001='{1}'
                                    ,TC002='{2}'
                                    ,TD003='{3}'
                                    ,TD004='{4}'
                                    ,TD005='{5}'
                                    ,PLANINDATE='{6}'
                                    ,PLANNUMS='{7}'
                                    WHERE [ID]='{0}'    
                                    "

                                    , ID
                                    , TC001
                                    , TC002
                                    , TD003
                                    , TD004
                                    , TD005
                                    , PLANINDATE
                                    , PLANNUMS

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


        #endregion


        #region BUTTON
        private void button1_Click(object sender, EventArgs e)
        {
            Search(dateTimePicker1.Value.ToString("yyyyMMdd"), dateTimePicker2.Value.ToString("yyyyMMdd"),comboBox4.Text.ToString());
        }
        private void button2_Click(object sender, EventArgs e)
        {
            string TC001="";
            string TC002 = "";
            string TD003 = "";
            string TD004 = "";
            string TD005 = "";
            string NUMS= "";           
            string MA002 = "";
            string STOCKS = comboBox2.Text.ToString();
            string ISIN = comboBox1.Text.ToString();
            string INVOICES = textBox2.Text.ToString();
            string INNO = textBox3.Text.ToString();
            string INNAMES = comboBox3.Text.ToString();

            ADD_TBPURINCHECK(
                 TC001,
                 TC002,
                 TD003,
                 TD004,
                 TD005,
                 NUMS,                 
                 MA002,
                 STOCKS,
                 ISIN,
                 INVOICES,
                 INNO,
                 INNAMES
                );

            comboBox3.Text = "";

            Search(dateTimePicker1.Value.ToString("yyyyMMdd"), dateTimePicker2.Value.ToString("yyyyMMdd"), comboBox4.Text.ToString());

            MessageBox.Show("完成");
        }

        private void button3_Click(object sender, EventArgs e)
        {
            string TC001 = "";
            string TC002 = "";
            string TD003 = "";
            string TD004 = "";
            string TD005 = "";
            string NUMS = "";
            string MA002 = "";
            string STOCKS = comboBox2.Text.ToString();
            string ISIN = comboBox1.Text.ToString();
            string INVOICES = textBox2.Text.ToString();
            string INNO = textBox3.Text.ToString();
            string INNAMES = comboBox3.Text.ToString();

            UPDATE_TBPURINCHECK(
                 TC001,
                 TC002,
                 TD003,
                 TD004,
                 TD005,
                 NUMS,
                 MA002,
                 STOCKS,
                 ISIN,
                 INVOICES,
                 INNO,
                 INNAMES
                );

            comboBox3.Text = "";

            Search(dateTimePicker1.Value.ToString("yyyyMMdd"), dateTimePicker2.Value.ToString("yyyyMMdd"), comboBox4.Text.ToString());

            MessageBox.Show("完成");
        }
        private void button4_Click(object sender, EventArgs e)
        {
            SEARCH_TBPURINCHECK(textBox5.Text);
        }

        private void button5_Click(object sender, EventArgs e)
        {

            // 顯示確認對話框
            DialogResult result = MessageBox.Show("確定要執行此操作嗎？", "確認", MessageBoxButtons.OKCancel);

            // 檢查使用者是否按下了確定按鈕
            if (result == DialogResult.OK)
            {
                // 確認後執行的動作
                // TODO: 在這裡執行您的程式碼
                // 例如：
                string ID = textBox6.Text;
                if (!string.IsNullOrEmpty(ID))
                {
                    DELETE_TBPURINCHECK(ID);

                    SEARCH_TBPURINCHECK(textBox5.Text);
                    MessageBox.Show("完成");
                }


            }
           
        }
        private void button6_Click(object sender, EventArgs e)
        {
            Search_TBPURINCHECKPLAN(dateTimePicker3.Value.ToString("yyyyMMdd"), dateTimePicker4.Value.ToString("yyyyMMdd"));
        }

        private void button7_Click(object sender, EventArgs e)
        {
            string TC001 = "";
            string TC002 = "";
            string TD003 = "";
            string TD004 = "";
            string TD005 = "";
            string PLANINDATE = "";
            string PLANNUMS = "";

            TC001 = textBox7.Text.Trim();
            TC002 = textBox8.Text.Trim();
            TD003 = textBox9.Text.Trim();
            TD004 = textBox10.Text.Trim();
            TD005 = textBox11.Text.Trim();
            PLANNUMS = textBox12.Text.Trim();
            PLANINDATE = dateTimePicker6.Value.ToString("yyyy/MM/dd");

            ADD_TBPURINCHECKPLAN(
                                TC001,
                                TC002,
                                TD003,
                                TD004,
                                TD005,
                                PLANINDATE,
                                PLANNUMS
                                );

            Search_TBPURINCHECKPLAN(dateTimePicker3.Value.ToString("yyyyMMdd"), dateTimePicker4.Value.ToString("yyyyMMdd"));
            MessageBox.Show("完成");
        }


        private void button8_Click(object sender, EventArgs e)
        {
            string ID = "";
            string TC001 = "";
            string TC002 = "";
            string TD003 = "";
            string TD004 = "";
            string TD005 = "";
            string PLANINDATE = "";
            string PLANNUMS = "";

            ID = textBox21.Text;
            TC001 = textBox14.Text.Trim();
            TC002 = textBox15.Text.Trim();
            TD003 = textBox16.Text.Trim();
            TD004 = textBox17.Text.Trim();
            TD005 = textBox18.Text.Trim();
            PLANNUMS = textBox19.Text.Trim();
            PLANINDATE = dateTimePicker5.Value.ToString("yyyy/MM/dd");

            UPDATE_TBPURINCHECKPLAN(
                                ID,
                                TC001,
                                TC002,
                                TD003,
                                TD004,
                                TD005,
                                PLANINDATE,
                                PLANNUMS
                                );

            Search_TBPURINCHECKPLAN(dateTimePicker3.Value.ToString("yyyyMMdd"), dateTimePicker4.Value.ToString("yyyyMMdd"));
            MessageBox.Show("完成");
        }
        private void button9_Click(object sender, EventArgs e)
        {        

            // 顯示確認對話框
            DialogResult result = MessageBox.Show("確定要執行此操作嗎？", "確認", MessageBoxButtons.OKCancel);

            // 檢查使用者是否按下了確定按鈕
            if (result == DialogResult.OK)
            {
                // 確認後執行的動作
                // TODO: 在這裡執行您的程式碼
                // 例如：
                string ID = textBox21.Text;

                if (!string.IsNullOrEmpty(ID))
                {
                    DELETE_TBPURINCHECKPLAN(ID);

                    Search_TBPURINCHECKPLAN(dateTimePicker3.Value.ToString("yyyyMMdd"), dateTimePicker4.Value.ToString("yyyyMMdd"));
                    MessageBox.Show("完成");
                }


            }
        }



        #endregion

     
    }
}

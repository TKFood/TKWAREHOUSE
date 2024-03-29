﻿using System;
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
using TKITDLL;

namespace TKWAREHOUSE
{
    public partial class FrmUPDATECOPTG : Form
    {
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

        SqlTransaction tran;
        SqlCommand cmd = new SqlCommand();
        DataSet ds1 = new DataSet();
        DataSet ds2 = new DataSet();
        DataSet ds3 = new DataSet();

        DataTable dt = new DataTable();
        string tablename = null;
        int result;

        public FrmUPDATECOPTG()
        {
            InitializeComponent();

            comboboxload2();
        }

        private void FrmUPDATECOPTG_Load(object sender, EventArgs e)
        {
            dataGridView1.AlternatingRowsDefaultCellStyle.BackColor = Color.PaleTurquoise;      //奇數列顏色
          
            //先建立個 CheckBox 欄
            DataGridViewCheckBoxColumn cbCol = new DataGridViewCheckBoxColumn();
            cbCol.Width = 120;   //設定寬度
            cbCol.HeaderText = "　全選";
            cbCol.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;   //置中
            cbCol.TrueValue = true;
            cbCol.FalseValue = false;
            dataGridView1.Columns.Insert(0, cbCol);

            //建立个矩形，等下计算 CheckBox 嵌入 GridView 的位置
            Rectangle rect = dataGridView1.GetCellDisplayRectangle(0, -1, true);
            rect.X = rect.Location.X + rect.Width / 4 - 30;
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
        #region FUNCTION
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

            String Sequel = "SELECT MA001,MA002 FROM [TK].dbo.PURMA WHERE MA001 LIKE '8%' ORDER BY MA001";
            SqlDataAdapter da = new SqlDataAdapter(Sequel, sqlConn);
            DataTable dt = new DataTable();
            sqlConn.Open();

            dt.Columns.Add("MA001", typeof(string));
            dt.Columns.Add("MA002", typeof(string));
            da.Fill(dt);
            comboBox2.DataSource = dt.DefaultView;
            comboBox2.ValueMember = "MA001";
            comboBox2.DisplayMember = "MA002";
            sqlConn.Close();

            comboBox2.SelectedValue = "8000001 ";

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

                    adapter1 = new SqlDataAdapter(@"" + sbSql, sqlConn);
                    sqlCmdBuilder1 = new SqlCommandBuilder(adapter1);

                    sqlConn.Open();
                    ds1.Clear();
                    adapter1.Fill(ds1, "ds1");
                    sqlConn.Close();


                    if (ds1.Tables["ds1"].Rows.Count == 0)
                    {
                        dataGridView1.DataSource = null;
                    }
                    else if (ds1.Tables["ds1"].Rows.Count >= 1)
                    {
                        dataGridView1.DataSource = ds1.Tables["ds1"];
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

            if (comboBox1.Text.ToString().Equals("未指定"))
            {
                STR.AppendFormat(@"  SELECT TG001 AS '銷貨單',TG002 AS '銷貨單號',TG007 AS '客戶',TG112 AS '貨運廠商',MA002 AS '貨運廠商名'");
                STR.AppendFormat(@"  FROM [TK].dbo.COPTG");
                STR.AppendFormat(@"  LEFT JOIN [TK].dbo.PURMA ON  TG112=MA001");
                STR.AppendFormat(@"  WHERE TG003>='{0}' AND TG003<='{1}'",dateTimePicker1.Value.ToString("yyyyMMdd"), dateTimePicker2.Value.ToString("yyyyMMdd"));
                STR.AppendFormat(@"  AND ISNULL(TG112,'')=''");
                STR.AppendFormat(@"  ORDER BY TG001,TG002");
                STR.AppendFormat(@"  ");
            }
            else if (comboBox1.Text.ToString().Equals("全部"))
            {

                STR.AppendFormat(@"  SELECT TG001 AS '銷貨單',TG002 AS '銷貨單號',TG007 AS '客戶',TG112 AS '貨運廠商',MA002 AS '貨運廠商名'");
                STR.AppendFormat(@"  FROM [TK].dbo.COPTG");
                STR.AppendFormat(@"  LEFT JOIN [TK].dbo.PURMA ON  TG112=MA001");
                STR.AppendFormat(@"  WHERE TG003>='{0}' AND TG003<='{1}'", dateTimePicker1.Value.ToString("yyyyMMdd"), dateTimePicker2.Value.ToString("yyyyMMdd"));
                STR.AppendFormat(@"  ORDER BY TG001,TG002");
                STR.AppendFormat(@"  ");
                STR.AppendFormat(@"  ");

            }
         

            return STR;
        }


        private void dataGridView1_SelectionChanged(object sender, EventArgs e)
        {
            if (dataGridView1.CurrentRow != null)
            {
                int rowindex = dataGridView1.CurrentRow.Index;
                if (rowindex >= 0)
                {
                    DataGridViewRow row = dataGridView1.Rows[rowindex];

                    SearchCOPTH(row.Cells["銷貨單"].Value.ToString(), row.Cells["銷貨單號"].Value.ToString());
                }
            }
        }
        public void SearchCOPTH(string TG001,string TG002)
        {
            StringBuilder sbSql = new StringBuilder();
            try
            {
                if (!string.IsNullOrEmpty(TG001) || !string.IsNullOrEmpty(TG002))
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

                    sbSql.AppendFormat(@" SELECT TH004 AS '品號',TH005 AS '品名',TH006 AS '規格',TH008 AS '數量',TH009 AS '單位',TH017 AS '批號',TH001 AS '銷貨單',TH002 AS '銷貨單號',TH003 AS '銷貨序號'");
                    sbSql.AppendFormat(@" FROM [TK].dbo.COPTH");
                    sbSql.AppendFormat(@" WHERE TH001='{0}' AND TH002='{1}'",TG001,TG002);
                    sbSql.AppendFormat(@" ORDER BY TH001,TH002,TH003");
                    sbSql.AppendFormat(@" ");

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
                    else if (ds3.Tables["ds3"].Rows.Count >= 1)
                    {
                        dataGridView2.DataSource = ds3.Tables["ds3"];
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
        public void UPDATECOPTG()
        {
            StringBuilder TG001TG002 = new StringBuilder();
            TG001TG002.Clear();

            foreach (DataGridViewRow row in dataGridView1.Rows)
            {
                DataGridViewCheckBoxCell chk = (DataGridViewCheckBoxCell)row.Cells[0];
                if (chk.Value == chk.TrueValue)
                {
                    TG001TG002.AppendFormat(@"'{0}',", row.Cells["銷貨單"].Value.ToString() + row.Cells["銷貨單號"].Value.ToString());

                    //MessageBox.Show(row.Cells["銷貨單"].Value.ToString()+ row.Cells["銷貨單號"].Value.ToString());
                }
            }

            TG001TG002.AppendFormat(@"''");

            SETCOPTGTG112(TG001TG002.ToString());
            //MessageBox.Show(TG001TG002.ToString());
        }

        public void SETCOPTGTG112(string TG001TG002)
        {
            if(!string.IsNullOrEmpty(TG001TG002))
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

                    sbSql.AppendFormat(" UPDATE [TK].dbo.COPTG SET TG112='{0}' WHERE TG001+TG002 IN ({1})",comboBox2.SelectedValue.ToString(),TG001TG002);
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

                        MessageBox.Show("更新完成");
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

        #endregion

        #region BUTTON
        private void button1_Click(object sender, EventArgs e)
        {
            Search();
        }
        private void button2_Click(object sender, EventArgs e)
        {
            UPDATECOPTG();

            Search();
        }


        #endregion

      
    }
}

﻿using System;
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
    public partial class FrmMOCTATB : Form
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

        public FrmMOCTATB()
        {
            InitializeComponent();

            comboboxload1();
            comboBox1.Text = "原料";
        }


        #region FUNCTION
        public void comboboxload1()
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
                                SELECT  [ID]
                                ,[KINDS]
                                ,[NAMES]
                                ,[KEYS]
                                ,[KEYS2]
                                FROM[TKWAREHOUSE].[dbo].[TBPARAS]
                                WHERE[KINDS] = 'FrmMOCTATB' ORDER BY[KEYS]");

            SqlDataAdapter da = new SqlDataAdapter(Sequel.ToString(), sqlConn);
            DataTable dt = new DataTable();
            sqlConn.Open();

            dt.Columns.Add("NAMES", typeof(string));

            da.Fill(dt);
            comboBox1.DataSource = dt.DefaultView;
            comboBox1.ValueMember = "NAMES";
            comboBox1.DisplayMember = "NAMES";
            sqlConn.Close();

            comboBox1.SelectedValue = "20001";

        }

        public void SETFASTREPORT(string KINDS, string TA009, string TA009B, string TA012, string TA012B)
        {
            string SQL;
            report1 = new Report();
            report1.Load(@"REPORT\烘焙製令領用表.frx");

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
            SQL = SETFASETSQL(KINDS, TA009, TA009B, TA012, TA012B);

            Table.SelectCommand = SQL;
            report1.Preview = previewControl1;
            report1.Show();
        }

        public string SETFASETSQL(string KINDS, string TA009, string TA009B, string TA012, string TA012B)
        {
            StringBuilder FASTSQL = new StringBuilder();
            StringBuilder STRQUERY = new StringBuilder();

            FASTSQL.Clear();
            sbSqlQuery.Clear();

            //TA012 實際開工
            if (KINDS.Equals("原料"))
            {
                sbSqlQuery.AppendFormat(@" AND TB003 NOT LIKE '2%' 
                                         AND TA009 >='{0}'   AND TA009 <='{1}' 
                                        ", TA009, TA009B);
            }
            //TA009 預計開工
            else if (KINDS.Equals("物料"))
            {
                sbSqlQuery.AppendFormat(@" AND TB003 LIKE '2%' 
                                           AND TA012 >='{0}'  AND TA012 <='{1}'  
                                        ", TA012, TA012B);
            }

            FASTSQL.AppendFormat(@"  
                                SELECT 
                                TA001 AS '製令單別'
                                ,TA002 AS '製令單號'
                                ,TA003 AS '開單日期'
                                ,TA006 AS '產品品號'
                                ,TA009 AS '預計開工'
                                ,TA012 AS '實際開工'
                                ,TA034 AS '產品品名'
                                ,TA015 AS '預計產量'
                                ,TA007 AS '產品單位'
                                ,TB003 AS '材料品號'
                                ,TB012 AS '材料品名' 
                                ,(CASE WHEN TB003 LIKE '1%' OR TB003 LIKE '3%' THEN TB004 ELSE CONVERT(INT, TB004) END )AS '需領用量'
                                ,TB007 AS '材料單位'
                                ,(SELECT ISNULL(SUM(LA005*LA011),0) FROM [TK].dbo.INVLA WHERE LA009 IN ('21003') AND  LA001=TB003)AS '庫存量'
                                FROM [TK].dbo.MOCTA,[TK].dbo.MOCTB
                                WHERE TA001=TB001 AND TA002=TB002
                                AND TA001='A513'
                                AND TA013='Y'
                              
                                {0}

                                ORDER BY TB003,TA001,TA002 
                                        

                                        ", sbSqlQuery.ToString());

            return FASTSQL.ToString();
        }

        #endregion

        #region BUTTON
        private void button1_Click(object sender, EventArgs e)
        {
            SETFASTREPORT(comboBox1.Text.ToString(), dateTimePicker1.Value.ToString("yyyyMMdd"), dateTimePicker2.Value.ToString("yyyyMMdd"), dateTimePicker3.Value.ToString("yyyyMMdd"), dateTimePicker4.Value.ToString("yyyyMMdd"));
        }
        #endregion

    }
}

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


namespace TKWAREHOUSE
{
    public partial class FrmCOPINVCHECK : Form
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
        SqlDataAdapter adapter6 = new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilder6 = new SqlCommandBuilder();
        SqlTransaction tran;
        SqlCommand cmd = new SqlCommand();
        DataSet ds = new DataSet();
        DataSet ds2 = new DataSet();
        DataSet ds3 = new DataSet();
        DataSet ds4 = new DataSet();
        DataSet ds5 = new DataSet();
        DataSet ds6 = new DataSet();
        DataTable dt = new DataTable();
        DataTable dtTemp = new DataTable();
        DataTable dtTemp2 = new DataTable();
        int result;
        string tablename = null;
        decimal COPNum = 0;
        decimal TOTALCOPNum = 0;
        double BOMNum = 0;
        double FinalNum = 0;
        decimal COOKIES = 1;
        decimal BATCH = 1;
       
        string CHECKYN = "N";
        string CHECKYNMOCPLANWEEKPUR = "N";
        decimal MOCBATCH = 1;
        string TD001 = null;
        string TD002 = null;
        string TD003 = null;
        string YEARS;
        string WEEKS;

        string TA002;
        string MOCPLANWEEKPURID;

        public FrmCOPINVCHECK()
        {
            InitializeComponent();
        }

        #region FUNCTION
        public void Search()
        {
            DateTime dt = new DateTime();
            dt = DateTime.Now;
            dt=dt.AddDays(Convert.ToDouble(numericUpDown1.Value));

            StringBuilder TD001 = new StringBuilder();
            StringBuilder TC027 = new StringBuilder();
            StringBuilder PALNQUERY = new StringBuilder();

            if (checkBox1.Checked == true)
            {
                TD001.AppendFormat("'A221',");
            }
            if (checkBox2.Checked == true)
            {
                TD001.AppendFormat("'A222',");
            }

            if (checkBox4.Checked == true)
            {
                TD001.AppendFormat("'A225',");
            }
            if (checkBox5.Checked == true)
            {
                TD001.AppendFormat("'A226',");
            }
            if (checkBox6.Checked == true)
            {
                TD001.AppendFormat("'A227',");
            }
            if (checkBox7.Checked == true)
            {
                TD001.AppendFormat("'A223',");
            }
            TD001.AppendFormat("''");

            if (comboBox1.Text.ToString().Equals("已確認"))
            {
                TC027.AppendFormat(" 'Y',");
            }
            else if (comboBox1.Text.ToString().Equals("未確認"))
            {
                TC027.AppendFormat(" 'N',");
            }
            else if (comboBox1.Text.ToString().Equals("全部"))
            {
                TC027.AppendFormat(" 'Y','N', ");
            }
            TC027.AppendFormat("''");

            

            dtTemp.Clear();
            dtTemp2.Clear();
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sbSql.Clear();
                sbSqlQuery.Clear();

                sbSql.AppendFormat(@"  SELECT 日期,品號,品名,客戶,CONVERT(INT,SUM(訂單數量)) AS 訂單數量,單位 ,單別,單號,序號  ");
                sbSql.AppendFormat(@"   ,(SELECT CONVERT(INT,ISNULL(SUM(LA005*LA011),0)) FROM [TK].dbo.INVLA WITH (NOLOCK) WHERE LA009='20001' AND LA001=品號) AS '成品倉庫存'");
                sbSql.AppendFormat(@"   ,(SELECT CONVERT(INT,ISNULL(SUM(LA005*LA011),0)) FROM [TK].dbo.INVLA WITH (NOLOCK) WHERE LA009='20002' AND LA001=品號) AS '外銷倉庫存'");
                sbSql.AppendFormat(@"   ,(SELECT CONVERT(INT,ISNULL(SUM(TA015-TA017-TA018),0)) FROM [TK].dbo.MOCTA  WITH (NOLOCK) WHERE TA011 NOT IN ('Y','y') AND TA006=品號 ) AS '未完成的製令' ");
                sbSql.AppendFormat(@"  ,(SELECT CONVERT(INT,ISNULL(MC004,0))  FROM [TK].dbo.BOMMC WHERE MC001=品號) AS 標準批量");
                sbSql.AppendFormat(@"  FROM (");
                sbSql.AppendFormat(@"  SELECT   TD001 AS '單別',TD002 AS '單號',TD003 AS '序號',TC053  AS '客戶' ,TD013 AS '日期',TD004 AS '品號',TD005 AS '品名',TD006 AS '規格'");
                sbSql.AppendFormat(@"  ,(CASE WHEN MB004=TD010 THEN ((TD008-TD009)+(TD024-TD025)) ELSE ((TD008-TD009)+(TD024-TD025))*MD004 END) AS '訂單數量'");
                sbSql.AppendFormat(@"  ,MB004 AS '單位'");
                sbSql.AppendFormat(@"  ,((TD008-TD009)+(TD024-TD025)) AS '訂單量'");
                sbSql.AppendFormat(@"  ,TD010 AS '訂單單位' ");
                sbSql.AppendFormat(@"  ,(CASE WHEN ISNULL(MD002,'')<>'' THEN MD002 ELSE TD010 END ) AS '換算單位'");
                sbSql.AppendFormat(@"  ,(CASE WHEN MD003>0 THEN MD003 ELSE 1 END) AS '分子'");
                sbSql.AppendFormat(@"  ,(CASE WHEN MD004>0 THEN MD004 ELSE (TD008-TD009) END ) AS '分母'");
                sbSql.AppendFormat(@"  FROM [TK].dbo.INVMB,[TK].dbo.COPTC,[TK].dbo.COPTD");
                sbSql.AppendFormat(@"  LEFT JOIN [TK].dbo.INVMD ON MD001=TD004 AND TD010=MD002");
                sbSql.AppendFormat(@"  WHERE TD004=MB001");
                sbSql.AppendFormat(@"  AND TC001=TD001 AND TC002=TD002");
                sbSql.AppendFormat(@"  AND TD004 LIKE '4%'");
                sbSql.AppendFormat(@"  AND TD013>='{0}' AND TD013<='{1}'", DateTime.Now.ToString("yyyyMMdd"), dt.ToString("yyyyMMdd"));
                sbSql.AppendFormat(@"  AND TC001 IN ({0}) ", TD001.ToString());
                sbSql.AppendFormat(@"  AND (TD008-TD009)>0  ");
                sbSql.AppendFormat(@" AND TC027 IN ({0})  ", TC027.ToString());
                sbSql.AppendFormat(@"  {0}", PALNQUERY.ToString());
                //sbSql.AppendFormat(@"  AND ( TD004 LIKE '40109916000740%'  ) ");
                sbSql.AppendFormat(@"  ) AS TEMP");
                sbSql.AppendFormat(@"  GROUP  BY 客戶,日期,品號,品名,規格,單位,單別,單號,序號");
                sbSql.AppendFormat(@"  ORDER BY 日期,客戶,品號,品名,規格,單位,單別,單號,序號 ");
                sbSql.AppendFormat(@"  ");

                adapter = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder = new SqlCommandBuilder(adapter);
                sqlConn.Open();
                ds.Clear();
                adapter.Fill(ds, "TEMPds1");
                sqlConn.Close();


                if (CHECKYN.Equals("N"))
                {
                    //建立一個DataGridView的Column物件及其內容
                    DataGridViewColumn dgvc = new DataGridViewCheckBoxColumn();
                    dgvc.Width = 40;
                    dgvc.Name = "選取";

                    this.dataGridView1.Columns.Insert(0, dgvc);
                    CHECKYN = "Y";
                }


                if (ds.Tables["TEMPds1"].Rows.Count == 0)
                {
                    dataGridView1.DataSource = null;
                }
                else
                {
                    if (ds.Tables["TEMPds1"].Rows.Count >= 1)
                    {                       
                        //dataGridView1.Rows.Clear();
                        dataGridView1.DataSource = ds.Tables["TEMPds1"];
                        dataGridView1.AutoResizeColumns();


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



        #endregion

        #region BUTTON
        private void button1_Click(object sender, EventArgs e)
        {
            Search();
        }

        #endregion

    }
}

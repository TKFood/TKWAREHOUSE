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
using System.Collections;
using TKITDLL;

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
        SqlDataAdapter adapter7= new SqlDataAdapter();
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
        DataTable dtTemp = new DataTable();
        DataTable dtTemp2 = new DataTable();
        int result;
        string tablename = null;
        decimal COPNum = 0;

        int rowIndexDG3 = -1;
        int rowIndexDG5 = -1;



        public FrmCOPINVCHECK()
        {
            InitializeComponent();

            NEWdtTemp();
            NEWdtTemp2();

        }


        #region FUNCTION
        public void NEWdtTemp()
        {
            dtTemp.Columns.Add("商品");

            dtTemp.Columns.Add("品號");
            dtTemp.Columns.Add("品名");
            //dtTemp.Columns.Add("數量");
            dtTemp.Columns.Add("單位");

            DataColumn colDecimal = new DataColumn("數量");
            colDecimal.DataType = System.Type.GetType("System.Decimal");
            dtTemp.Columns.Add(colDecimal);
        }

        public void NEWdtTemp2()
        {
            dtTemp2.Columns.Add("品號");
            dtTemp2.Columns.Add("品名");
            //dtTemp.Columns.Add("數量");
            dtTemp2.Columns.Add("規格");
            dtTemp2.Columns.Add("庫存單位");

            DataColumn colDecimal = new DataColumn("需求數量");
            colDecimal.DataType = System.Type.GetType("System.Decimal");
            dtTemp2.Columns.Add(colDecimal);

            DataColumn colDecimal2 = new DataColumn("庫存量");
            colDecimal2.DataType = System.Type.GetType("System.Decimal");
            dtTemp2.Columns.Add(colDecimal2);

            DataColumn colDecimal3 = new DataColumn("需求量比較");
            colDecimal2.DataType = System.Type.GetType("System.Decimal");
            dtTemp2.Columns.Add(colDecimal3);

            DataColumn colDecimal4 = new DataColumn("預計採購量");
            colDecimal2.DataType = System.Type.GetType("System.Decimal");
            dtTemp2.Columns.Add(colDecimal4);
        }
        public void Search()
        {
            DateTime dt = new DateTime();
            dt = dateTimePicker1.Value;
            //dt = DateTime.Now;
            //dt=dt.AddDays(Convert.ToDouble(numericUpDown1.Value));

            StringBuilder TD001 = new StringBuilder();
            StringBuilder TC027 = new StringBuilder();
            StringBuilder QUERY1 = new StringBuilder();

            

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


            if (comboBox2.Text.ToString().Equals("排除已製令"))
            {
                QUERY1.AppendFormat(" AND TD001+TD002+TD003 NOT IN (SELECT TA026+TA027+TA027 FROM [TK].dbo.MOCTA WITH (NOLOCK) WHERE ISNULL(TA026+TA027,'')<>'') ");
            }
            else if (comboBox1.Text.ToString().Equals("全部"))
            {
                QUERY1.AppendFormat(" ");
            }

           

            

            dtTemp.Clear();
            dtTemp2.Clear();
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
                sbSqlQuery.Clear();

                sbSql.AppendFormat(@"  SELECT 日期,品號,品名,客戶,CONVERT(INT,SUM(訂單數量)) AS 訂單數量,單位 ,單別,單號,序號,規格  ");
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
                sbSql.AppendFormat(@"  AND (TD004 LIKE '401%' OR TD004 LIKE '402%' OR TD004 LIKE '403%' OR TD004 LIKE '404%' OR TD004 LIKE '405%' OR TD004 LIKE '406%' OR TD004 LIKE '407%'   ) ");
                sbSql.AppendFormat(@"  AND TD013>='{0}' AND TD013<='{1}'", dateTimePicker1.Value.ToString("yyyyMMdd"), dateTimePicker2.Value.ToString("yyyyMMdd"));
                sbSql.AppendFormat(@"  AND ((TD008-TD009)+(TD024-TD025))>0   ");
    
                sbSql.AppendFormat(@" AND TC027 IN ({0})  ", TC027.ToString());
                sbSql.AppendFormat(@"  {0}", QUERY1.ToString());
                //sbSql.AppendFormat(@"  AND ( TD004 LIKE '40102910540200%'  ) ");
                //sbSql.AppendFormat(@"  AND ( TD002='20181211001'  ) ");
                sbSql.AppendFormat(@") AS TEMP");
                sbSql.AppendFormat(@"  GROUP  BY 客戶,日期,品號,品名,規格,單位,單別,單號,序號");
                sbSql.AppendFormat(@"  ORDER BY 日期,客戶,品號,品名,規格,單位,單別,單號,序號 ");
                sbSql.AppendFormat(@"  ");

                adapter = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder = new SqlCommandBuilder(adapter);
                sqlConn.Open();
                ds.Clear();
                adapter.Fill(ds, "TEMPds1");
                sqlConn.Close();


                //if (CHECKYN.Equals("N"))
                //{
                //    //建立一個DataGridView的Column物件及其內容
                //    DataGridViewColumn dgvc = new DataGridViewCheckBoxColumn();
                //    dgvc.Width = 40;
                //    dgvc.Name = "選取";

                //    this.dataGridView1.Columns.Insert(0, dgvc);
                //    CHECKYN = "Y";
                //}


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

        public void Search2()
        {
            DateTime dt = new DateTime();
            dt = dateTimePicker1.Value;
            //dt = DateTime.Now;
            //dt=dt.AddDays(Convert.ToDouble(numericUpDown1.Value));

            StringBuilder TD001 = new StringBuilder();
            StringBuilder TC027 = new StringBuilder();
            StringBuilder QUERY1 = new StringBuilder();



            if (comboBox3.Text.ToString().Equals("已確認"))
            {
                TC027.AppendFormat(" 'Y',");
            }
            else if (comboBox3.Text.ToString().Equals("未確認"))
            {
                TC027.AppendFormat(" 'N',");
            }
            else if (comboBox3.Text.ToString().Equals("全部"))
            {
                TC027.AppendFormat(" 'Y','N', ");
            }
            TC027.AppendFormat("''");


            if (comboBox4.Text.ToString().Equals("排除已製令"))
            {
                QUERY1.AppendFormat(" AND TD001+TD002+TD003 NOT IN (SELECT TA026+TA027+TA027 FROM [TK].dbo.MOCTA WITH (NOLOCK) WHERE ISNULL(TA026+TA027,'')<>'') ");
            }
            else if (comboBox4.Text.ToString().Equals("全部"))
            {
                QUERY1.AppendFormat(" ");
            }
            
            dtTemp.Clear();
            dtTemp2.Clear();
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
                sbSqlQuery.Clear();

                sbSql.AppendFormat(@"  SELECT 日期,品號,品名,客戶,CONVERT(INT,SUM(訂單數量)) AS 訂單數量,單位 ,單別,單號,序號,規格  ");
                sbSql.AppendFormat(@"   ,(SELECT CONVERT(INT,ISNULL(SUM(LA005*LA011),0)) FROM [TK].dbo.INVLA WITH (NOLOCK) WHERE LA009='20001' AND LA001=品號) AS '成品倉庫存'");
                sbSql.AppendFormat(@"   ,(SELECT CONVERT(INT,ISNULL(SUM(LA005*LA011),0)) FROM [TK].dbo.INVLA WITH (NOLOCK) WHERE LA009='20002' AND LA001=品號) AS '外銷倉庫存'");
                sbSql.AppendFormat(@"   ,(SELECT CONVERT(INT,ISNULL(SUM(TA015-TA017-TA018),0)) FROM [TK].dbo.MOCTA  WITH (NOLOCK) WHERE TA011 NOT IN ('Y','y') AND TA006=品號 AND TA009>='{0}' AND TA009<='{1}') AS '未完成的製令' ", dateTimePicker5.Value.ToString("yyyyMMdd"), dateTimePicker6.Value.ToString("yyyyMMdd"));
                sbSql.AppendFormat(@"  ,(SELECT CONVERT(INT,ISNULL(MC004,0))  FROM [TK].dbo.BOMMC WHERE MC001=品號) AS 標準批量");
                sbSql.AppendFormat(@"  FROM (");
                sbSql.AppendFormat(@"  SELECT   TD001 AS '單別',TD002 AS '單號',TD003 AS '序號',TC053  AS '客戶' ,TD013 AS '日期',TD004 AS '品號',TD005 AS '品名',TD006 AS '規格'");
                sbSql.AppendFormat(@"  ,((CASE WHEN MB004=TD010 THEN ((TD008-TD009)+(TD024-TD025)) ELSE ((TD008-TD009)+(TD024-TD025))*INVMD.MD004 END)-ISNULL(MOCTA.TA017,0)) AS '訂單數量'");
                sbSql.AppendFormat(@"  ,MB004 AS '單位'");
                sbSql.AppendFormat(@"  ,((TD008-TD009)+(TD024-TD025)) AS '訂單量'");
                sbSql.AppendFormat(@"  ,TD010 AS '訂單單位' ");
                sbSql.AppendFormat(@"  ,(CASE WHEN ISNULL(MD002,'')<>'' THEN MD002 ELSE TD010 END ) AS '換算單位'");
                sbSql.AppendFormat(@"  ,(CASE WHEN MD003>0 THEN MD003 ELSE 1 END) AS '分子'");
                sbSql.AppendFormat(@"  ,(CASE WHEN MD004>0 THEN MD004 ELSE (TD008-TD009) END ) AS '分母'");
                sbSql.AppendFormat(@"  ,ISNULL(MOCTA.TA017,0) AS TA017");
                sbSql.AppendFormat(@"  FROM [TK].dbo.INVMB,[TK].dbo.COPTC,[TK].dbo.COPTD");
                sbSql.AppendFormat(@"  LEFT JOIN [TK].dbo.INVMD ON MD001=TD004 AND TD010=MD002");
                sbSql.AppendFormat(@"  LEFT JOIN [TK].dbo.MOCTA ON TA026=TD001 AND TA027=TD002 AND TD028=TD003 AND TA006=TD004");
                sbSql.AppendFormat(@"  WHERE TD004=MB001");
                sbSql.AppendFormat(@"  AND TC001=TD001 AND TC002=TD002");
                sbSql.AppendFormat(@"  AND TD004 LIKE '4%'");
                sbSql.AppendFormat(@" ");
                sbSql.AppendFormat(@"  AND TD013>='{0}' AND TD013<='{1}'", dateTimePicker5.Value.ToString("yyyyMMdd"), dateTimePicker6.Value.ToString("yyyyMMdd"));
                sbSql.AppendFormat(@"  AND ((TD008-TD009)+(TD024-TD025))>0   ");

                sbSql.AppendFormat(@" AND TC027 IN ({0})  ", TC027.ToString());
                sbSql.AppendFormat(@"  {0}", QUERY1.ToString());
                //sbSql.AppendFormat(@"  AND ( TD004 LIKE '40102910540200%'  ) ");
                //sbSql.AppendFormat(@"  AND ( TD002='20181211001'  ) ");
                sbSql.AppendFormat(@") AS TEMP");
                sbSql.AppendFormat(@"  GROUP  BY 客戶,日期,品號,品名,規格,單位,單別,單號,序號");
                sbSql.AppendFormat(@"  ORDER BY 日期,客戶,品號,品名,規格,單位,單別,單號,序號 ");
                sbSql.AppendFormat(@"  ");

                adapter = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder = new SqlCommandBuilder(adapter);
                sqlConn.Open();
                ds.Clear();
                adapter.Fill(ds, "TEMPds1");
                sqlConn.Close();


                //if (CHECKYN.Equals("N"))
                //{
                //    //建立一個DataGridView的Column物件及其內容
                //    DataGridViewColumn dgvc = new DataGridViewCheckBoxColumn();
                //    dgvc.Width = 40;
                //    dgvc.Name = "選取";

                //    this.dataGridView1.Columns.Insert(0, dgvc);
                //    CHECKYN = "Y";
                //}


                if (ds.Tables["TEMPds1"].Rows.Count == 0)
                {
                   
                }
                else
                {
                    if (ds.Tables["TEMPds1"].Rows.Count >= 1)
                    {
                        SETCOPTHGROUPBY2();

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


        public void SETCOPTHGROUPBY()
        {
            try
            {
                if (ds.Tables["TEMPds1"].Rows.Count >= 1)
                {
                    var query = from t in ds.Tables["TEMPds1"].AsEnumerable()
                                group t by new { t1 = t.Field<string>("品號") } into m
                                select new
                                {
                                    MB001 = m.Key.t1,
                                    SUM = m.Sum(n => n.Field<int>("訂單數量"))
                                };
                    if (query.ToList().Count > 0)
                    {
                        string MB001 = null;
                        string MB003 = null;
                        string[] sArray = null;

                        //20210902密
                        Class1 TKID = new Class1();//用new 建立類別實體
                        SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

                        //資料庫使用者密碼解密
                        sqlsb.Password = TKID.Decryption(sqlsb.Password);
                        sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

                        String connectionString;
                        sqlConn = new SqlConnection(sqlsb.ConnectionString);

                        dtTemp.Clear();

                        query.ToList().ForEach(q =>
                        {

                            COPNum = Convert.ToDecimal(q.SUM);
                            MB001 = q.MB001.ToString();

                            //TOTALCOPNum = Convert.ToDecimal(Convert.ToDecimal(sArray[0].ToString())* COPNum);

                            sbSql.Clear();
                            sbSqlQuery.Clear();

                            sbSql.AppendFormat(@"  WITH TEMPTABLE (MD001,MD003,MD004,MD006,MD007,MD008,MC004,NUM,LV) AS");
                            sbSql.AppendFormat(@"  (");
                            sbSql.AppendFormat(@"  SELECT  MD001,MD003,MD004,MD006,MD007,MD008,MC004,CONVERT(decimal(18,5),(MD006*(1+MD008)/MD007)/MC004) AS NUM,1 AS LV FROM [TK].dbo.VBOMMD WHERE  MD001='{0}'", MB001);
                            sbSql.AppendFormat(@"  UNION ALL");
                            sbSql.AppendFormat(@"  SELECT A.MD001,A.MD003,A.MD004,A.MD006,A.MD007,A.MD008,A.MC004,CONVERT(decimal(18,5),(A.MD006*(1+A.MD008)/A.MD007/A.MC004)*(B.NUM)) AS NUM,LV+1");
                            sbSql.AppendFormat(@"  FROM [TK].dbo.VBOMMD A");
                            sbSql.AppendFormat(@"  INNER JOIN TEMPTABLE B on A.MD001=B.MD003");
                            sbSql.AppendFormat(@"  )");
                            sbSql.AppendFormat(@"  SELECT MD001,MD003,MD004,MD006,MD007,MD008,MC004,NUM,LV,MB002");
                            sbSql.AppendFormat(@"  FROM TEMPTABLE ");
                            sbSql.AppendFormat(@"  LEFT JOIN [TK].dbo.INVMB ON MB001=MD003");

                            sbSql.AppendFormat(@"  WHERE  (MD003 LIKE '1%') OR  (MD003 LIKE '2%')");
                            //sbSql.AppendFormat(@"  WHERE  MD003='203022061' ");
                            sbSql.AppendFormat(@"  ORDER BY LV,MD001,MD003");



                            adapter2 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                            sqlCmdBuilder2 = new SqlCommandBuilder(adapter2);
                            sqlConn.Open();
                            ds2.Clear();
                            adapter2.Fill(ds2, "TEMPds2");
                            sqlConn.Close();

                            if (ds2.Tables["TEMPds2"].Rows.Count >= 1)
                            {

                                foreach (DataRow od2 in ds2.Tables["TEMPds2"].Rows)
                                {
                                    DataRow row = dtTemp.NewRow();
                                    //row["MD001"] = od2["MC001"].ToString();

                                    row["商品"] =MB001.ToString();
                                    row["品號"] = od2["MD003"].ToString();
                                    row["品名"] = od2["MB002"].ToString();
                                    row["數量"] = Convert.ToDecimal(COPNum) * Convert.ToDecimal(od2["NUM"].ToString());
                                    row["單位"] = od2["MD004"].ToString();

                                    dtTemp.Rows.Add(row);


                                }

                            }

                        }
                        );
                    }



                    //query.ToList().ForEach(q =>
                    //{
                    //    //MessageBox.Show(q.MB001 + "," );
                    //    MessageBox.Show(q.MB001 + "," + q.SUM);
                    //});
                }

                if (dtTemp.Rows.Count > 0)
                {
                    dataGridView3.DataSource = dtTemp;
                    dataGridView3.AutoResizeColumns();

                    SETMOCGROUPBY();
                }
            }
            catch
            {

            }
            finally
            {

            }
           

        }

        public void SETCOPTHGROUPBY2()
        {
            try
            {
                if (ds.Tables["TEMPds1"].Rows.Count >= 1)
                {
                    var query = from t in ds.Tables["TEMPds1"].AsEnumerable()
                                group t by new { t1 = t.Field<string>("品號") } into m
                                select new
                                {
                                    MB001 = m.Key.t1,
                                    SUM = m.Sum(n => n.Field<int>("訂單數量"))
                                };
                    if (query.ToList().Count > 0)
                    {
                        string MB001 = null;
                        string MB003 = null;
                        string[] sArray = null;

                        //20210902密
                        Class1 TKID = new Class1();//用new 建立類別實體
                        SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

                        //資料庫使用者密碼解密
                        sqlsb.Password = TKID.Decryption(sqlsb.Password);
                        sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

                        String connectionString;
                        sqlConn = new SqlConnection(sqlsb.ConnectionString);

                        dtTemp.Clear();

                        query.ToList().ForEach(q =>
                        {

                            COPNum = Convert.ToDecimal(q.SUM);
                            MB001 = q.MB001.ToString();

                            //TOTALCOPNum = Convert.ToDecimal(Convert.ToDecimal(sArray[0].ToString())* COPNum);

                            sbSql.Clear();
                            sbSqlQuery.Clear();

                            sbSql.AppendFormat(@"  WITH TEMPTABLE (MD001,MD003,MD004,MD006,MD007,MD008,MC004,NUM,LV) AS");
                            sbSql.AppendFormat(@"  (");
                            sbSql.AppendFormat(@"  SELECT  MD001,MD003,MD004,MD006,MD007,MD008,MC004,CONVERT(decimal(18,5),(MD006*(1+MD008)/MD007)/MC004) AS NUM,1 AS LV FROM [TK].dbo.VBOMMD WHERE  MD001='{0}'", MB001);
                            sbSql.AppendFormat(@"  UNION ALL");
                            sbSql.AppendFormat(@"  SELECT A.MD001,A.MD003,A.MD004,A.MD006,A.MD007,A.MD008,A.MC004,CONVERT(decimal(18,5),(A.MD006*(1+A.MD008)/A.MD007/A.MC004)*(B.NUM)) AS NUM,LV+1");
                            sbSql.AppendFormat(@"  FROM [TK].dbo.VBOMMD A");
                            sbSql.AppendFormat(@"  INNER JOIN TEMPTABLE B on A.MD001=B.MD003");
                            sbSql.AppendFormat(@"  )");
                            sbSql.AppendFormat(@"  SELECT MD001,MD003,MD004,MD006,MD007,MD008,MC004,NUM,LV,MB002");
                            sbSql.AppendFormat(@"  FROM TEMPTABLE ");
                            sbSql.AppendFormat(@"  LEFT JOIN [TK].dbo.INVMB ON MB001=MD003");

                            sbSql.AppendFormat(@"  WHERE  (MD003 LIKE '1%') OR  (MD003 LIKE '2%')");
                            //sbSql.AppendFormat(@"  WHERE  MD003='203022061' ");
                            sbSql.AppendFormat(@"  ORDER BY LV,MD001,MD003");



                            adapter2 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                            sqlCmdBuilder2 = new SqlCommandBuilder(adapter2);
                            sqlConn.Open();
                            ds2.Clear();
                            adapter2.Fill(ds2, "TEMPds2");
                            sqlConn.Close();

                            if (ds2.Tables["TEMPds2"].Rows.Count >= 1)
                            {

                                foreach (DataRow od2 in ds2.Tables["TEMPds2"].Rows)
                                {
                                    DataRow row = dtTemp.NewRow();
                                    //row["MD001"] = od2["MC001"].ToString();

                                    row["商品"] = MB001.ToString();
                                    row["品號"] = od2["MD003"].ToString();
                                    row["品名"] = od2["MB002"].ToString();
                                    row["數量"] = Convert.ToDecimal(COPNum) * Convert.ToDecimal(od2["NUM"].ToString());
                                    row["單位"] = od2["MD004"].ToString();

                                    dtTemp.Rows.Add(row);


                                }

                            }

                        }
                        );
                    }



                    //query.ToList().ForEach(q =>
                    //{
                    //    //MessageBox.Show(q.MB001 + "," );
                    //    MessageBox.Show(q.MB001 + "," + q.SUM);
                    //});
                }

                if (dtTemp.Rows.Count > 0)
                {              
                    SETMOCGROUPBY2();
                }
            }
            catch
            {

            }
            finally
            {

            }


        }
        public void SETMOCGROUPBY()
        {
            DateTime dt = DateTime.Now;
            DateTime dt2 = DateTime.Now;
            dt2 = dateTimePicker2.Value;
            //dt2=dt2.AddDays(Convert.ToDouble(numericUpDown2.Value));

            if (dtTemp.Rows.Count >= 1)
            {
                var query = from t in dtTemp.AsEnumerable() 
                            group t by new { t1 = t.Field<string>("品號") } into m
                            orderby m.Key.t1
                            select new
                            {
                                MB001 = m.Key.t1,
                                SUM = m.Sum(n => n.Field<decimal>("數量"))
                                
                            }                           
                            ;

                if (query.ToList().Count > 0)
                {
                    //20210902密
                    Class1 TKID = new Class1();//用new 建立類別實體
                    SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

                    //資料庫使用者密碼解密
                    sqlsb.Password = TKID.Decryption(sqlsb.Password);
                    sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

                    String connectionString;
                    sqlConn = new SqlConnection(sqlsb.ConnectionString);

                    dtTemp2.Clear();

                    query.ToList().ForEach(q =>
                    {

                        string LA001 = q.MB001;
                        sbSql.Clear();
                        sbSqlQuery.Clear();

                        sbSql.AppendFormat(@" SELECT MB001 AS '品號',ISNULL(SUM(LA011*LA005),0) AS '庫存量',MB002 AS '品名',MB003 AS '規格',MB004 AS '庫存單位'   ");
                        sbSql.AppendFormat(@"  ,(SELECT ISNULL(SUM(TD008-TD015),0) FROM [TK].dbo.PURTD WHERE TD004=MB001 AND TD018='Y' AND TD016='N' AND TD012>='{0}' AND TD012<='{1}') AS '預計採購量' ", dateTimePicker3.Value.ToString("yyyyMMdd"), dateTimePicker4.Value.ToString("yyyyMMdd"));
                        sbSql.AppendFormat(@"  FROM [TK].dbo.INVMB   ");
                        sbSql.AppendFormat(@"  LEFT JOIN [TK].dbo.INVLA ON LA001=MB001 AND  LA009 IN ('20004','20006')");
                        sbSql.AppendFormat(@" WHERE  MB001='{0}' ", LA001);
                        sbSql.AppendFormat(@" GROUP BY MB001,LA001,MB002,MB003,MB004 ");
                        sbSql.AppendFormat(@"  ");


                        adapter3 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                        sqlCmdBuilder3= new SqlCommandBuilder(adapter3);
                        sqlConn.Open();
                        ds3.Clear();
                        adapter3.Fill(ds3, "TEMPds3");
                        sqlConn.Close();

                        if (ds3.Tables["TEMPds3"].Rows.Count >= 1)
                        {

                            foreach (DataRow od3 in ds3.Tables["TEMPds3"].Rows)
                            {
                                DataRow row = dtTemp2.NewRow();
                                //row["MD001"] = od2["MC001"].ToString();

                                row["品號"] = od3["品號"].ToString();
                                row["品名"] = od3["品名"].ToString();
                                row["規格"] = od3["規格"].ToString();
                                row["庫存單位"] = od3["庫存單位"].ToString();
                                row["需求數量"] = Convert.ToDecimal(q.SUM.ToString());
                                row["庫存量"] = Convert.ToDecimal(od3["庫存量"].ToString());
                                row["需求量比較"] =  Convert.ToDecimal(od3["庫存量"].ToString())- Convert.ToDecimal(q.SUM.ToString()) ;
                                row["預計採購量"] = Convert.ToDecimal(od3["預計採購量"].ToString());

                                dtTemp2.Rows.Add(row);
                            }

                        }

                        

                    }
                    );
                }


            }
            //dataGridView2.DataSource = query.ToList();
            //dataGridView2.AutoResizeColumns();

            dataGridView2.DataSource = dtTemp2;
            dataGridView2.AutoResizeColumns();

            //根据列表中数据不同，显示不同颜色背景
            foreach (DataGridViewRow dgRow in dataGridView2.Rows)
            {
                //判断
                if (Convert.ToDecimal(dgRow.Cells[6].Value) < 0)
                {
                    //将这行的背景色设置成Pink
                    dgRow.DefaultCellStyle.BackColor = Color.Pink;
                }
            }

        }

        public void SETMOCGROUPBY2()
        {
            DateTime dt = DateTime.Now;
            DateTime dt2 = DateTime.Now;
            dt2 = dateTimePicker2.Value;
            //dt2=dt2.AddDays(Convert.ToDouble(numericUpDown2.Value));

            if (dtTemp.Rows.Count >= 1)
            {
                var query = from t in dtTemp.AsEnumerable()
                            group t by new { t1 = t.Field<string>("品號") } into m
                            orderby m.Key.t1
                            select new
                            {
                                MB001 = m.Key.t1,
                                SUM = m.Sum(n => n.Field<decimal>("數量"))

                            }
                            ;

                if (query.ToList().Count > 0)
                {
                    //20210902密
                    Class1 TKID = new Class1();//用new 建立類別實體
                    SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

                    //資料庫使用者密碼解密
                    sqlsb.Password = TKID.Decryption(sqlsb.Password);
                    sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

                    String connectionString;
                    sqlConn = new SqlConnection(sqlsb.ConnectionString);

                    dtTemp2.Clear();

                    query.ToList().ForEach(q =>
                    {

                        string LA001 = q.MB001;
                        sbSql.Clear();
                        sbSqlQuery.Clear();

                        sbSql.AppendFormat(@" SELECT MB001 AS '品號',ISNULL(SUM(LA011*LA005),0) AS '庫存量',MB002 AS '品名',MB003 AS '規格',MB004 AS '庫存單位'   ");
                        sbSql.AppendFormat(@"  ,(SELECT ISNULL(SUM(TD008-TD015),0) FROM [TK].dbo.PURTD WHERE TD004=MB001 AND TD018='Y' AND TD016='N' AND TD012>='{0}' AND TD012<='{1}') AS '預計採購量' ", dateTimePicker7.Value.ToString("yyyyMMdd"), dateTimePicker8.Value.ToString("yyyyMMdd"));
                        sbSql.AppendFormat(@"  FROM [TK].dbo.INVMB   ");
                        sbSql.AppendFormat(@"  LEFT JOIN [TK].dbo.INVLA ON LA001=MB001 AND  LA009 IN ('20004','20006')");
                        sbSql.AppendFormat(@" WHERE  MB001='{0}' ", LA001);
                        sbSql.AppendFormat(@" GROUP BY MB001,LA001,MB002,MB003,MB004 ");
                        sbSql.AppendFormat(@"  ");


                        adapter3 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                        sqlCmdBuilder3 = new SqlCommandBuilder(adapter3);
                        sqlConn.Open();
                        ds3.Clear();
                        adapter3.Fill(ds3, "TEMPds3");
                        sqlConn.Close();

                        if (ds3.Tables["TEMPds3"].Rows.Count >= 1)
                        {

                            foreach (DataRow od3 in ds3.Tables["TEMPds3"].Rows)
                            {
                                DataRow row = dtTemp2.NewRow();
                                //row["MD001"] = od2["MC001"].ToString();

                                row["品號"] = od3["品號"].ToString();
                                row["品名"] = od3["品名"].ToString();
                                row["規格"] = od3["規格"].ToString();
                                row["庫存單位"] = od3["庫存單位"].ToString();
                                row["需求數量"] = Convert.ToDecimal(q.SUM.ToString());
                                row["庫存量"] = Convert.ToDecimal(od3["庫存量"].ToString());
                                row["需求量比較"] = Convert.ToDecimal(od3["庫存量"].ToString()) - Convert.ToDecimal(q.SUM.ToString());
                                row["預計採購量"] = Convert.ToDecimal(od3["預計採購量"].ToString());

                                dtTemp2.Rows.Add(row);
                            }

                        }



                    }
                    );
                }


            }
            //dataGridView2.DataSource = query.ToList();
            //dataGridView2.AutoResizeColumns();

            DataView dView = new DataView(dtTemp2);
            dView.Sort = "品號";
            dtTemp2 = dView.ToTable();

            //dtTemp2.DefaultView.Sort = "品號 ";
            //dtTemp2 = dtTemp2.DefaultView.ToTable();

            dataGridView3.DataSource = dtTemp2;
            
            dataGridView3.AutoResizeColumns();

            //根据列表中数据不同，显示不同颜色背景
            foreach (DataGridViewRow dgRow in dataGridView3.Rows)
            {
                //判断
                if (Convert.ToDecimal(dgRow.Cells[6].Value) < 0)
                {
                    //将这行的背景色设置成Pink
                    dgRow.DefaultCellStyle.BackColor = Color.Pink;
                }
            }

        }

        public void ExcelExport()
        {

            string NowDB = "TK";
            //建立Excel 2003檔案
            IWorkbook wb = new XSSFWorkbook();
            ISheet ws;

            ws = wb.CreateSheet("Sheet1");
            ws.CreateRow(0);
            //第一行為欄位名稱
            ws.GetRow(0).CreateCell(0).SetCellValue("品號");
            ws.GetRow(0).CreateCell(1).SetCellValue("品名");
            ws.GetRow(0).CreateCell(2).SetCellValue("規格");
            ws.GetRow(0).CreateCell(3).SetCellValue("庫存單位");
            ws.GetRow(0).CreateCell(4).SetCellValue("需求數量");
            ws.GetRow(0).CreateCell(5).SetCellValue("庫存量");
            ws.GetRow(0).CreateCell(6).SetCellValue("需求量比較");
            ws.GetRow(0).CreateCell(7).SetCellValue("預計採購量");




            int j = 0;
            foreach (DataGridViewRow dr in this.dataGridView2.Rows)
            {
                ws.CreateRow(j + 1);
                ws.GetRow(j + 1).CreateCell(0).SetCellValue(dr.Cells[0].Value.ToString());
                ws.GetRow(j + 1).CreateCell(1).SetCellValue(dr.Cells[1].Value.ToString());
                ws.GetRow(j + 1).CreateCell(2).SetCellValue(dr.Cells[2].Value.ToString());
                ws.GetRow(j + 1).CreateCell(3).SetCellValue(dr.Cells[3].Value.ToString());
                ws.GetRow(j + 1).CreateCell(4).SetCellValue(dr.Cells[4].Value.ToString());
                ws.GetRow(j + 1).CreateCell(5).SetCellValue(dr.Cells[5].Value.ToString());
                ws.GetRow(j + 1).CreateCell(6).SetCellValue(dr.Cells[6].Value.ToString());
                ws.GetRow(j + 1).CreateCell(7).SetCellValue(dr.Cells[7].Value.ToString());

                j++;
            }

           


            if (Directory.Exists(@"c:\temp\"))
            {
                //資料夾存在
            }
            else
            {
                //新增資料夾
                Directory.CreateDirectory(@"c:\temp\");
            }
            StringBuilder filename = new StringBuilder();
            filename.AppendFormat(@"c:\temp\宅配資料{0}.xlsx", DateTime.Now.ToString("yyyyMMdd"));

            FileStream file = new FileStream(filename.ToString(), FileMode.Create);//產生檔案
            wb.Write(file);
            file.Close();

            MessageBox.Show("匯出完成-EXCEL放在-" + filename.ToString());
            FileInfo fi = new FileInfo(filename.ToString());
            if (fi.Exists)
            {
                System.Diagnostics.Process.Start(filename.ToString());
            }
            else
            {
                //file doesn't exist
            }


        }
        private void dataGridView2_SelectionChanged(object sender, EventArgs e)
        {
            if (dataGridView2.CurrentRow != null)
            {
                int rowindex = dataGridView2.CurrentRow.Index;
                string MB001=null;
                if (rowindex >= 0)
                {
                    DataGridViewRow row = dataGridView2.Rows[rowindex];
                    MB001 = row.Cells["品號"].Value.ToString();

                    //MessageBox.Show(MB001.ToString());
                    Search3(MB001.ToString());
                }
                else
                {
                    MB001 = null;

                }
            }
        }

        public void Search3(string MB001)
        {
            DateTime dt = new DateTime();
            dt = dateTimePicker1.Value;
            //dt = DateTime.Now;
            //dt = dt.AddDays(Convert.ToDouble(numericUpDown1.Value));

            StringBuilder TD001 = new StringBuilder();
            StringBuilder TC027 = new StringBuilder();
            StringBuilder QUERY1 = new StringBuilder();

           

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


            QUERY1.AppendFormat(" AND ((TD004 IN (SELECT [MD001] FROM [TK].[dbo].[VBOMMD] WHERE [MD003]='{0}')) OR (TD004 IN (SELECT MD001 FROM [TK].[dbo].[VBOMMD] WHERE MD003 IN (SELECT MD001 FROM [TK].[dbo].[VBOMMD] WHERE [MD003]='{0}' ))) OR (TD004 IN (SELECT MD001 FROM [TK].[dbo].[BOMMD]  WHERE MD003 IN ( SELECT MD001 FROM [TK].[dbo].[BOMMD]  WHERE MD003 IN ( SELECT MD001 FROM [TK].[dbo].[BOMMD] WHERE [MD003]='{0}' ))) ) ) ", MB001.ToString());
            

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
                sbSqlQuery.Clear();

                sbSql.AppendFormat(@"  SELECT 日期,品號,品名,客戶,CONVERT(INT,SUM(訂單數量)) AS 訂單數量,單位 ,單別,單號,序號,規格  ");
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
                sbSql.AppendFormat(@"  AND (TD004 LIKE '401%' OR TD004 LIKE '402%' OR TD004 LIKE '403%' OR TD004 LIKE '404%' OR TD004 LIKE '405%' OR TD004 LIKE '406%' OR TD004 LIKE '407%'   ) ");
                sbSql.AppendFormat(@"  AND TD013>='{0}' AND TD013<='{1}'", DateTime.Now.ToString("yyyyMMdd"), dt.ToString("yyyyMMdd"));
    
                sbSql.AppendFormat(@"  AND ((TD008-TD009)+(TD024-TD025))>0   ");
                sbSql.AppendFormat(@"  AND ((TD008-TD009)+(TD024-TD025))>0   ");
                sbSql.AppendFormat(@"  AND ((TD008-TD009)+(TD024-TD025))>0   ");
                sbSql.AppendFormat(@"  AND ((TD008-TD009)+(TD024-TD025))>0   ");
                sbSql.AppendFormat(@" AND TC027 IN ({0})  ", TC027.ToString());
                sbSql.AppendFormat(@"  {0}", QUERY1.ToString());
                //sbSql.AppendFormat(@"  AND ( TD004 LIKE '40102910540200%'  ) ");
                //sbSql.AppendFormat(@"  AND ( TD002='20180708006'  ) ");
                sbSql.AppendFormat(@") AS TEMP");
                sbSql.AppendFormat(@"  GROUP  BY 客戶,日期,品號,品名,規格,單位,單別,單號,序號");
                sbSql.AppendFormat(@"  ORDER BY 日期,客戶,品號,品名,規格,單位,單別,單號,序號 ");
                sbSql.AppendFormat(@"  ");

                adapter7 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder7 = new SqlCommandBuilder(adapter7);
                sqlConn.Open();
                ds7.Clear();
                adapter7.Fill(ds7, "TEMPds7");
                sqlConn.Close();



                if (ds7.Tables["TEMPds7"].Rows.Count == 0)
                {
                    dataGridView4.DataSource = null;
                }
                else
                {
                    if (ds7.Tables["TEMPds7"].Rows.Count >= 1)
                    {
                        //dataGridView1.Rows.Clear();
                        dataGridView4.DataSource = ds7.Tables["TEMPds7"];
                        dataGridView4.AutoResizeColumns();


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
        private void dataGridView3_SelectionChanged(object sender, EventArgs e)
        {
            if (dataGridView3.CurrentRow != null)
            {
                int rowindex = dataGridView3.CurrentRow.Index;
                string MB001 = null;
                if (rowindex >= 0)
                {
                    DataGridViewRow row = dataGridView3.Rows[rowindex];
                    MB001 = row.Cells["品號"].Value.ToString();

                    //MessageBox.Show(MB001.ToString());
                    Search4(MB001.ToString());
                }
                else
                {
                    MB001 = null;

                }
            }

            
        }

        public void Search4(string MB001)
        {
            DateTime dt = new DateTime();
            dt = dateTimePicker1.Value;
            //dt = DateTime.Now;
            //dt = dt.AddDays(Convert.ToDouble(numericUpDown1.Value));

            StringBuilder TD001 = new StringBuilder();
            StringBuilder TC027 = new StringBuilder();
            StringBuilder QUERY1 = new StringBuilder();



            if (comboBox3.Text.ToString().Equals("已確認"))
            {
                TC027.AppendFormat(" 'Y',");
            }
            else if (comboBox3.Text.ToString().Equals("未確認"))
            {
                TC027.AppendFormat(" 'N',");
            }
            else if (comboBox3.Text.ToString().Equals("全部"))
            {
                TC027.AppendFormat(" 'Y','N', ");
            }
            TC027.AppendFormat("''");


            //QUERY1.AppendFormat(" AND ((TD004 IN (SELECT [MD001] FROM [TK].[dbo].[VBOMMD] WHERE [MD003]='{0}')) OR (TD004 IN (SELECT MD001 FROM [TK].[dbo].[VBOMMD] WHERE MD003 IN (SELECT MD001 FROM [TK].[dbo].[VBOMMD] WHERE [MD003]='{0}' ))) OR (TD004 IN (SELECT MD001 FROM [TK].[dbo].[BOMMD]  WHERE MD003 IN ( SELECT MD001 FROM [TK].[dbo].[BOMMD]  WHERE MD003 IN ( SELECT MD001 FROM [TK].[dbo].[BOMMD] WHERE [MD003]='{0}' ))) ) ) ", MB001.Trim().ToString());


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
                sbSqlQuery.Clear();

                sbSql.AppendFormat(@"  SELECT SUM(NUNC) AS '預計用量',日期,品號,品名,客戶,CONVERT(INT,SUM(訂單數量)) AS 訂單數量,單位 ,單別,單號,序號,規格  ");
                sbSql.AppendFormat(@"   ,(SELECT CONVERT(INT,ISNULL(SUM(LA005*LA011),0)) FROM [TK].dbo.INVLA WITH (NOLOCK) WHERE LA009='20001' AND LA001=品號) AS '成品倉庫存'");
                sbSql.AppendFormat(@"   ,(SELECT CONVERT(INT,ISNULL(SUM(LA005*LA011),0)) FROM [TK].dbo.INVLA WITH (NOLOCK) WHERE LA009='20002' AND LA001=品號) AS '外銷倉庫存'");
                sbSql.AppendFormat(@"   ,(SELECT CONVERT(INT,ISNULL(SUM(TA015-TA017-TA018),0)) FROM [TK].dbo.MOCTA  WITH (NOLOCK) WHERE TA011 NOT IN ('Y','y') AND  TA026=單別 AND TA027=單號 AND TA028=序號 AND TA009>='{0}' AND TA009<='{1}') AS '未完成的製令' ", dateTimePicker5.Value.ToString("yyyyMMdd"), dateTimePicker6.Value.ToString("yyyyMMdd"));
                sbSql.AppendFormat(@"  ,(SELECT CONVERT(INT,ISNULL(MC004,0))  FROM [TK].dbo.BOMMC WHERE MC001=品號) AS 標準批量");
                sbSql.AppendFormat(@"  FROM (");
                sbSql.AppendFormat(@"  SELECT   TD001 AS '單別',TD002 AS '單號',TD003 AS '序號',TC053  AS '客戶' ,TD013 AS '日期',TD004 AS '品號',TD005 AS '品名',TD006 AS '規格'");
                sbSql.AppendFormat(@"  ,(CASE WHEN MB004=TD010 THEN ((TD008-TD009)+(TD024-TD025)) ELSE ((TD008-TD009)+(TD024-TD025))*INVMD.MD004 END) AS '訂單數量'");
                sbSql.AppendFormat(@"  ,MB004 AS '單位'");
                sbSql.AppendFormat(@"  ,((TD008-TD009)+(TD024-TD025)) AS '訂單量'");
                sbSql.AppendFormat(@"  ,TD010 AS '訂單單位' ");
                sbSql.AppendFormat(@"  ,(CASE WHEN ISNULL(INVMD.MD002,'')<>'' THEN INVMD.MD002 ELSE TD010 END ) AS '換算單位'");
                sbSql.AppendFormat(@"  ,(CASE WHEN INVMD.MD003>0 THEN INVMD.MD003 ELSE 1 END) AS '分子'");
                sbSql.AppendFormat(@"  ,(CASE WHEN INVMD.MD004>0 THEN INVMD.MD004 ELSE (TD008-TD009) END ) AS '分母'");
                sbSql.AppendFormat(@"  ,BOMMD.MD001 AS 'MD001A',BOMMD.MD003 AS 'MD003A',BOMMD.MD006 AS 'MD006A',BOMMD.MD007 AS 'MD007A',BOMMD.MD008 AS 'MD008A',BOMMC.MC004 AS 'MC004A'");
                sbSql.AppendFormat(@"  ,BOMMD2.MD001 AS 'MD001B',BOMMD2.MD003 AS 'MD003B',ISNULL(BOMMD2.MD006,1) AS 'MD006B',ISNULL(BOMMD2.MD007,1) AS 'MD007B',ISNULL(BOMMD2.MD008,0) AS 'MD008B',ISNULL(BOMMC2.MC004,1) AS 'MC004B'");
                sbSql.AppendFormat(@"  ,BOMMD3.MD001 AS 'MD001C',BOMMD3.MD003 AS 'MD003C',ISNULL(BOMMD3.MD006,1) AS 'MD006C',ISNULL(BOMMD3.MD007,1) AS 'MD007C',ISNULL(BOMMD3.MD008,0) AS 'MD008C',ISNULL(BOMMC3.MC004,1) AS 'MC004C'");
                sbSql.AppendFormat(@"  ,(((CASE WHEN MB004=TD010 THEN ((TD008-TD009)+(TD024-TD025)) ELSE ((TD008-TD009)+(TD024-TD025))*INVMD.MD004 END)-ISNULL(MOCTA.TA017,0)))/BOMMC.MC004*(BOMMD.MD006*(1+BOMMD.MD008))/BOMMD.MD007 AS 'NUNA'");
                sbSql.AppendFormat(@"  ,((((CASE WHEN MB004=TD010 THEN ((TD008-TD009)+(TD024-TD025)) ELSE ((TD008-TD009)+(TD024-TD025))*INVMD.MD004 END)-ISNULL(MOCTA.TA017,0)))/BOMMC.MC004*(BOMMD.MD006*(1+BOMMD.MD008))/BOMMD.MD007)/ISNULL(BOMMC2.MC004,1)*ISNULL(BOMMD2.MD006,1)*(1+ISNULL(BOMMD2.MD008,0))/ISNULL(BOMMD2.MD007,1) AS 'NUNB'");
                sbSql.AppendFormat(@"  ,(((((CASE WHEN MB004=TD010 THEN ((TD008-TD009)+(TD024-TD025)) ELSE ((TD008-TD009)+(TD024-TD025))*INVMD.MD004 END)-ISNULL(MOCTA.TA017,0)))/BOMMC.MC004*(BOMMD.MD006*(1+BOMMD.MD008))/BOMMD.MD007)/ISNULL(BOMMC2.MC004,1)*ISNULL(BOMMD2.MD006,1)*(1+ISNULL(BOMMD2.MD008,0))/ISNULL(BOMMD2.MD007,1))/ISNULL(BOMMC3.MC004,1)*ISNULL(BOMMD3.MD006,1)*(1+ISNULL(BOMMD3.MD008,0))/ISNULL(BOMMD3.MD007,1)  AS 'NUNC'");
                sbSql.AppendFormat(@"  ,ISNULL(MOCTA.TA017,0) AS TA017");
                sbSql.AppendFormat(@"  FROM [TK].dbo.INVMB,[TK].dbo.COPTC,[TK].dbo.COPTD");                
                sbSql.AppendFormat(@"  LEFT JOIN [TK].dbo.INVMD ON MD001=TD004 AND TD010=MD002");
                sbSql.AppendFormat(@"  LEFT JOIN [TK].dbo.MOCTA ON TA026=TD001 AND TA027=TD002 AND TD028=TD003 AND TA006=TD004");
                sbSql.AppendFormat(@"  LEFT JOIN [TK].dbo.BOMMD ON BOMMD.MD001=TD004 ");
                sbSql.AppendFormat(@"  LEFT JOIN [TK].dbo.BOMMC ON BOMMC.MC001=TD004");
                sbSql.AppendFormat(@"  LEFT JOIN [TK].dbo.BOMMD  BOMMD2 ON BOMMD2.MD001=BOMMD.MD003");
                sbSql.AppendFormat(@"  LEFT JOIN [TK].dbo.BOMMC  BOMMC2 ON BOMMC2.MC001=BOMMD.MD003");
                sbSql.AppendFormat(@"  LEFT JOIN [TK].dbo.BOMMD  BOMMD3 ON BOMMD3.MD001=BOMMD2.MD003");
                sbSql.AppendFormat(@"  LEFT JOIN [TK].dbo.BOMMC  BOMMC3 ON BOMMC3.MC001=BOMMD2.MD003");
                sbSql.AppendFormat(@"  WHERE TD004=MB001");
                sbSql.AppendFormat(@"  AND TC001=TD001 AND TC002=TD002");
                sbSql.AppendFormat(@"  AND TD004 LIKE '4%'");
                
                sbSql.AppendFormat(@"  AND TD013>='{0}' AND TD013<='{1}'", dateTimePicker5.Value.ToString("yyyyMMdd"), dateTimePicker6.Value.ToString("yyyyMMdd"));
                sbSql.AppendFormat(@"  AND ((TD008-TD009)+(TD024-TD025))>0   ");
               
                sbSql.AppendFormat(@"  AND TC027 IN ({0})  ", TC027.ToString());
                sbSql.AppendFormat(@"  AND (BOMMD.MD003='{0}' OR BOMMD2.MD003='{0}' OR BOMMD3.MD003='{0}')", MB001.Trim().ToString());
                //sbSql.AppendFormat(@"  {0}", QUERY1.ToString());
                //sbSql.AppendFormat(@"  AND ( TD004 LIKE '40102910540200%'  ) ");
                //sbSql.AppendFormat(@"  AND ( TD002='20180708006'  ) ");
                sbSql.AppendFormat(@") AS TEMP");
                sbSql.AppendFormat(@"  GROUP  BY 客戶,日期,品號,品名,規格,單位,單別,單號,序號");
                sbSql.AppendFormat(@"  ORDER BY 日期,客戶,品號,品名,規格,單位,單別,單號,序號 ");
                sbSql.AppendFormat(@"  ");

                adapter7 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder7 = new SqlCommandBuilder(adapter7);
                sqlConn.Open();
                ds7.Clear();
                adapter7.Fill(ds7, "TEMPds7");
                sqlConn.Close();



                if (ds7.Tables["TEMPds7"].Rows.Count == 0)
                {
                    dataGridView5.DataSource = null;
                }
                else
                {
                    if (ds7.Tables["TEMPds7"].Rows.Count >= 1)
                    {
                        //dataGridView1.Rows.Clear();
                        dataGridView5.DataSource = ds7.Tables["TEMPds7"];
                        dataGridView5.AutoResizeColumns();


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

        public void SEARCHDG3(string SEARCHSTRING, int INDEX)
        {
            String searchValue = SEARCHSTRING;
            rowIndexDG3 = INDEX;
            int ROWS = 0;

            for (int i = INDEX; i < dataGridView3.Rows.Count; i++)
            {
                ROWS = i;

                if (dataGridView3.Rows[i].Cells[0].Value.ToString().Contains(searchValue))
                {
                    rowIndexDG3 = i;

                    dataGridView3.CurrentRow.Selected = false;
                    dataGridView3.Rows[i].Selected = true;
                    int index = rowIndexDG3;
                    dataGridView3.FirstDisplayedScrollingRowIndex = index;

                    break;
                }
                if (dataGridView3.Rows[i].Cells[1].Value.ToString().Contains(searchValue))
                {
                    rowIndexDG3 = i;

                    dataGridView3.CurrentRow.Selected = false;
                    dataGridView3.Rows[i].Selected = true;
                    int index = rowIndexDG3;
                    dataGridView3.FirstDisplayedScrollingRowIndex = index;

                    break;
                }
            }

           if(ROWS== dataGridView3.Rows.Count-1)
            {
                if (MessageBox.Show("已查到最後一筆，是否從頭開始?", "已查到最後一筆，是否從頭開始?", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                {
                    SEARCHDG3(textBox1.Text.Trim(), 0);
                }
                else
                {
                    
                }
            }
        }

        public void SEARCHDG5(string SEARCHSTRING, int INDEX)
        {
            String searchValue = SEARCHSTRING;
            rowIndexDG5 = INDEX;
            int ROWS = 0;

            for (int i = INDEX; i < dataGridView5.Rows.Count; i++)
            {
                ROWS = i;

                if (dataGridView5.Rows[i].Cells[1].Value.ToString().Contains(searchValue))
                {
                    rowIndexDG5 = i;

                    dataGridView5.CurrentRow.Selected = false;
                    dataGridView5.Rows[i].Selected = true;
                    int index = rowIndexDG5;
                    dataGridView5.FirstDisplayedScrollingRowIndex = index;

                    break;
                }
                if (dataGridView5.Rows[i].Cells[2].Value.ToString().Contains(searchValue))
                {
                    rowIndexDG5 = i;

                    dataGridView5.CurrentRow.Selected = false;
                    dataGridView5.Rows[i].Selected = true;
                    int index = rowIndexDG5;
                    dataGridView5.FirstDisplayedScrollingRowIndex = index;

                    break;
                }
                if (dataGridView5.Rows[i].Cells[3].Value.ToString().Contains(searchValue))
                {
                    rowIndexDG5 = i;

                    dataGridView5.CurrentRow.Selected = false;
                    dataGridView5.Rows[i].Selected = true;
                    int index = rowIndexDG5;
                    dataGridView5.FirstDisplayedScrollingRowIndex = index;

                    break;
                }
            }

            if (ROWS == dataGridView5.Rows.Count - 1)
            {
                if (MessageBox.Show("已查到最後一筆，是否從頭開始?", "已查到最後一筆，是否從頭開始?", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                {
                    SEARCHDG5(textBox1.Text.Trim(), 0);
                }
                else
                {

                }
            }
        }

        private void dateTimePicker5_ValueChanged(object sender, EventArgs e)
        {
            dateTimePicker7.Value = dateTimePicker5.Value;
        }

        private void dateTimePicker6_ValueChanged(object sender, EventArgs e)
        {
            dateTimePicker8.Value = dateTimePicker6.Value;
        }


        #endregion

        #region BUTTON
        private void button1_Click(object sender, EventArgs e)
        {
            Search();
        }
        private void button3_Click(object sender, EventArgs e)
        {
            Search();
            SETCOPTHGROUPBY();
            //SEARCHCOOKIES();
        }
        private void button2_Click(object sender, EventArgs e)
        {
           ExcelExport();
        }

        private void button6_Click(object sender, EventArgs e)
        {
            Search2();
        }

        private void button4_Click(object sender, EventArgs e)
        {
            //SEARCHDG3(textBox1.Text.Trim(), 0);

            if (rowIndexDG3 == -1)
            {
                SEARCHDG3(textBox1.Text.Trim(), 0);
            }
            else
            {
                SEARCHDG3(textBox1.Text.Trim(), rowIndexDG3 + 1);
            }

        }

        private void button5_Click(object sender, EventArgs e)
        {
            //SEARCHDG5(textBox1.Text.Trim(), 0);

            if (rowIndexDG5 == -1)
            {
                SEARCHDG5(textBox1.Text.Trim(), 0);
            }
            else
            {
                SEARCHDG5(textBox1.Text.Trim(), rowIndexDG5 + 1);
            }
        }



        #endregion

      
    }
}

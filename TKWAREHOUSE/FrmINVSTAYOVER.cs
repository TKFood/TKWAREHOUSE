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
    public partial class FrmINVSTAYOVER : Form
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


        SqlDataAdapter adapterCALENDAR = new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilderCALENDAR = new SqlCommandBuilder();


        DataTable dt = new DataTable();
        string tablename = null;
        DateTime StayDay;

        public Report report1 { get; private set; }

        public FrmINVSTAYOVER()
        {
            InitializeComponent();
            comboboxload();
            comboboxload2();

            dateTimePicker1.Value = DateTime.Now;
            dateTimePicker2.Value = DateTime.Now;
            dateTimePicker3.Value = DateTime.Now;
            dateTimePicker4.Value = DateTime.Now;

            DateTime StayDay = dateTimePicker1.Value;
            StayDay = StayDay.AddDays(-1 * Convert.ToDouble(textBox1.Text));

            dateTimePicker2.Value = StayDay;
            dateTimePicker4.Value = StayDay;
        }

        #region FUNCTION
        public void comboboxload()
        {


            //20210902密
            Class1 TKID = new Class1();//用new 建立類別實體
            SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

            //資料庫使用者密碼解密
            sqlsb.Password = TKID.Decryption(sqlsb.Password);
            sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

            String connectionString;
            sqlConn = new SqlConnection(sqlsb.ConnectionString);

            String Sequel = "SELECT MC001,MC001+MC002 AS MC002 FROM [TK].dbo.CMSMC WITH (NOLOCK) ORDER BY MC001";
            SqlDataAdapter da = new SqlDataAdapter(Sequel, sqlConn);
            DataTable dt = new DataTable();
            sqlConn.Open();

            dt.Columns.Add("MC001", typeof(string));
            dt.Columns.Add("MC002", typeof(string));
            da.Fill(dt);
            comboBox1.DataSource = dt.DefaultView;
            comboBox1.ValueMember = "MC001";
            comboBox1.DisplayMember = "MC002";
            sqlConn.Close();

            comboBox1.SelectedValue = "20001";

        }

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

            String Sequel = "SELECT MC001,MC001+MC002 AS MC002 FROM [TK].dbo.CMSMC WITH (NOLOCK) ORDER BY MC001";
            SqlDataAdapter da = new SqlDataAdapter(Sequel, sqlConn);
            DataTable dt = new DataTable();
            sqlConn.Open();

            dt.Columns.Add("MC001", typeof(string));
            dt.Columns.Add("MC002", typeof(string));
            da.Fill(dt);
            comboBox2.DataSource = dt.DefaultView;
            comboBox2.ValueMember = "MC001";
            comboBox2.DisplayMember = "MC002";
            sqlConn.Close();

            comboBox2.SelectedValue = "20001";

        }


        private void dateTimePicker1_ValueChanged(object sender, EventArgs e)
        {
            DateTime StayDay = dateTimePicker1.Value;
            StayDay = StayDay.AddDays(-1 * Convert.ToDouble(textBox1.Text));

            dateTimePicker2.Value = StayDay;
        }
        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            DateTime StayDay = dateTimePicker1.Value;
            StayDay = StayDay.AddDays(-1 * Convert.ToDouble(textBox1.Text));

            dateTimePicker2.Value = StayDay;
        }
        private void dateTimePicker3_ValueChanged(object sender, EventArgs e)
        {
            DateTime StayDay = dateTimePicker3.Value;
            StayDay = StayDay.AddDays(-1 * Convert.ToDouble(textBox2.Text));

            dateTimePicker4.Value = StayDay;
        }
        private void textBox2_TextChanged(object sender, EventArgs e)
        {
            DateTime StayDay = dateTimePicker3.Value;
            StayDay = StayDay.AddDays(-1 * Convert.ToDouble(textBox2.Text));

            dateTimePicker4.Value = StayDay;
        }

        public void Search()
        {
            StayDay = dateTimePicker1.Value;
            StayDay = StayDay.AddDays(-1*Convert.ToDouble(textBox1.Text));
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

                if (comboBox1.SelectedValue.ToString().Trim().Equals("20006")|| comboBox1.SelectedValue.ToString().Trim().Equals("20001") || comboBox1.SelectedValue.ToString().Trim().Equals("20005"))
                {
                    sbSql.AppendFormat(@" SELECT INVMB.MB001 AS '品號',INVMB.MB002 AS '品名',INVMB.MB003 AS '規格',INVMC.MC002 AS '庫別' ,CMSMC.MC002 AS '庫名',INVMC.MC012 AS '最近入庫日' ,INVMC.MC013 AS '最近出庫日' ");
                    sbSql.AppendFormat(@" ,MF002 AS '批號',SUM(MF008*MF010)  AS '庫存量'");
                    sbSql.AppendFormat(@"  FROM TK..INVMB INVMB ,TK..INVMC INVMC ,TK..CMSMC CMSMC ,TK.dbo.INVME, TK.dbo.INVMF");
                    sbSql.AppendFormat(@" WHERE INVMB.MB001=INVMC.MC001 AND INVMC.MC002=CMSMC.MC001 AND MB001=ME001 AND ME001=MF001 AND ME002=MF002 AND MF007=INVMC.MC002");
                    sbSql.AppendFormat(@" AND (( INVMC.MC012<='{0}') AND ( INVMC.MC013<='{0}') )", StayDay.ToString("yyyyMMdd"));
                    sbSql.AppendFormat(@" AND INVMC.MC002='{0}'", comboBox1.SelectedValue.ToString());
                    sbSql.AppendFormat(@" GROUP BY INVMB.MB001,INVMB.MB002,INVMB.MB003,INVMC.MC002,CMSMC.MC002 ,INVMC.MC007 ,INVMC.MC012 ,INVMC.MC013   ,INVMF.MF001,INVMF.MF002     ");
                    sbSql.AppendFormat(@"  HAVING SUM(MF008*MF010)>0");
                    sbSql.AppendFormat(@"  ORDER BY INVMB.MB001,INVMB.MB002,INVMB.MB003,INVMC.MC002,CMSMC.MC002      ");
                    sbSql.AppendFormat(@" ");
                    sbSql.AppendFormat(@" ");
                    sbSql.AppendFormat(@" ");
                }
                else
                {
                    sbSql.AppendFormat(@" SELECT INVMB.MB001 AS '品號',INVMB.MB002 AS '品名',INVMB.MB003 AS '規格',INVMC.MC002 AS '庫別',CMSMC.MC002 AS '庫名',INVMC.MC012 AS '最近入庫日',INVMC.MC013 AS '最近出庫日','' AS '批號',INVMC.MC007 AS '庫存量'");
                    sbSql.AppendFormat(@" FROM TK..INVMB INVMB ,TK..INVMC INVMC ,TK..CMSMC CMSMC");
                    sbSql.AppendFormat(@" WHERE INVMB.MB001=INVMC.MC001 AND INVMC.MC002=CMSMC.MC001 AND  INVMC.MC007>0 ");
                    sbSql.AppendFormat(@" AND (( INVMC.MC012<='{0}') AND ( INVMC.MC013<='{0}') )", StayDay.ToString("yyyyMMdd"));
                    sbSql.AppendFormat(@" AND INVMC.MC002='{0}'", comboBox1.SelectedValue.ToString());
                    sbSql.AppendFormat(@" ORDER BY INVMB.MB001,INVMB.MB002,INVMB.MB003,INVMC.MC002,CMSMC.MC002");
                    sbSql.AppendFormat(@" ");
                }



                tablename = "TEMPds1";
                adapter = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder = new SqlCommandBuilder(adapter);
                sqlConn.Open();
                ds.Clear();
                adapter.Fill(ds, tablename);
                sqlConn.Close();


                if (ds.Tables[tablename].Rows.Count == 0)
                {
                    label14.Text = "找不到資料";
                    //dataGridView1.Rows.Clear();
                    //dataGridView1.DataSource = null;
                }
                else
                {
                    label14.Text = "有 " + ds.Tables[tablename].Rows.Count.ToString() + " 筆";
                    //dataGridView1.Rows.Clear();
                    //dataGridView1.DataSource = null;
                    //dataGridView1.DataSource = ds.Tables[tablename];
                    //dataGridView1.AutoResizeColumns();
                }

            }
            catch
            {

            }
            finally
            {

            }
        }

        //public void ExcelExport()
        //{

        //    string NowDB = "TK";
        //    //建立Excel 2003檔案
        //    IWorkbook wb = new XSSFWorkbook();
        //    ISheet ws;

        //    XSSFCellStyle cs = (XSSFCellStyle)wb.CreateCellStyle();
        //    //框線樣式及顏色
        //    cs.BorderBottom = NPOI.SS.UserModel.BorderStyle.Double;
        //    cs.BorderLeft = NPOI.SS.UserModel.BorderStyle.Thin;
        //    cs.BorderRight = NPOI.SS.UserModel.BorderStyle.Thin;
        //    cs.BorderTop = NPOI.SS.UserModel.BorderStyle.Thin;
        //    cs.BottomBorderColor = NPOI.HSSF.Util.HSSFColor.Grey50Percent.Index;
        //    cs.LeftBorderColor = NPOI.HSSF.Util.HSSFColor.Grey50Percent.Index;
        //    cs.RightBorderColor = NPOI.HSSF.Util.HSSFColor.Grey50Percent.Index;
        //    cs.TopBorderColor = NPOI.HSSF.Util.HSSFColor.Grey50Percent.Index;

        //    Search();
        //    dt = ds.Tables[tablename];

        //    if (dt.TableName != string.Empty)
        //    {
        //        ws = wb.CreateSheet(dt.TableName);
        //    }
        //    else
        //    {
        //        ws = wb.CreateSheet("Sheet1");
        //    }

        //    ws.CreateRow(0);//第一行為欄位名稱
        //    for (int i = 0; i < dt.Columns.Count; i++)
        //    {
        //        ws.GetRow(0).CreateCell(i).SetCellValue(dt.Columns[i].ColumnName);
        //    }

        //    int j = 0;
        //    if (tablename.Equals("TEMPds1"))
        //    {
        //        foreach (DataGridViewRow dr in this.dataGridView1.Rows)
        //        {
        //            ws.CreateRow(j + 1);
        //            ws.GetRow(j + 1).CreateCell(0).SetCellValue(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[0].ToString());
        //            ws.GetRow(j + 1).CreateCell(1).SetCellValue(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[1].ToString());
        //            ws.GetRow(j + 1).CreateCell(2).SetCellValue(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[2].ToString());
        //            ws.GetRow(j + 1).CreateCell(3).SetCellValue(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[3].ToString());
        //            ws.GetRow(j + 1).CreateCell(4).SetCellValue(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[4].ToString());
        //            ws.GetRow(j + 1).CreateCell(5).SetCellValue(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[7].ToString());
        //            ws.GetRow(j + 1).CreateCell(6).SetCellValue(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[6].ToString());
        //            ws.GetRow(j + 1).CreateCell(7).SetCellValue(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[7].ToString());
        //            ws.GetRow(j + 1).CreateCell(8).SetCellValue(Convert.ToDouble(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[5].ToString()));
        //            j++;
        //        }
        //    }


        //    if (Directory.Exists(@"c:\temp\"))
        //    {
        //        //資料夾存在
        //    }
        //    else
        //    {
        //        //新增資料夾
        //        Directory.CreateDirectory(@"c:\temp\");
        //    }
        //    StringBuilder filename = new StringBuilder();
        //    filename.AppendFormat(@"c:\temp\庫存呆滯表{0}.xlsx", DateTime.Now.ToString("yyyyMMdd"));

        //    FileStream file = new FileStream(filename.ToString(), FileMode.Create);//產生檔案
        //    wb.Write(file);
        //    file.Close();

        //    MessageBox.Show("匯出完成-EXCEL放在-" + filename.ToString());
        //    FileInfo fi = new FileInfo(filename.ToString());
        //    if (fi.Exists)
        //    {
        //        System.Diagnostics.Process.Start(filename.ToString());
        //    }
        //    else
        //    {
        //        //file doesn't exist
        //    }


        //}


        public void SETFASTREPORT(string LA009,DateTime StayDay)
        {

            string SQL;
            report1 = new Report();
            report1.Load(@"REPORT\庫存呆滯表.frx");

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
            SQL = SETFASETSQL(LA009,StayDay);

            Table.SelectCommand = SQL;
            report1.Preview = previewControl1;
            report1.Show();

        }

        public string SETFASETSQL(string LA009, DateTime StayDay)
        {
            StringBuilder FASTSQL = new StringBuilder();
            StringBuilder STRQUERY = new StringBuilder();



            if (comboBox1.SelectedValue.ToString().Trim().Equals("20006") || comboBox1.SelectedValue.ToString().Trim().Equals("20001") || comboBox1.SelectedValue.ToString().Trim().Equals("20005"))
            {
                FASTSQL.AppendFormat(@" SELECT INVMB.MB001 AS '品號',INVMB.MB002 AS '品名',INVMB.MB003 AS '規格',INVMC.MC002 AS '庫別' ,CMSMC.MC002 AS '庫名',INVMC.MC012 AS '最近入庫日' ,INVMC.MC013 AS '最近出庫日' ");
                FASTSQL.AppendFormat(@" ,MF002 AS '批號',SUM(MF008*MF010)  AS '庫存量'");
                FASTSQL.AppendFormat(@"  FROM TK..INVMB INVMB ,TK..INVMC INVMC ,TK..CMSMC CMSMC ,TK.dbo.INVME, TK.dbo.INVMF");
                FASTSQL.AppendFormat(@" WHERE INVMB.MB001=INVMC.MC001 AND INVMC.MC002=CMSMC.MC001 AND MB001=ME001 AND ME001=MF001 AND ME002=MF002 AND MF007=INVMC.MC002");
                FASTSQL.AppendFormat(@" AND (( INVMC.MC012<='{0}') AND ( INVMC.MC013<='{0}') )", StayDay.ToString("yyyyMMdd"));
                FASTSQL.AppendFormat(@" AND INVMC.MC002='{0}'", LA009);
                FASTSQL.AppendFormat(@" GROUP BY INVMB.MB001,INVMB.MB002,INVMB.MB003,INVMC.MC002,CMSMC.MC002 ,INVMC.MC007 ,INVMC.MC012 ,INVMC.MC013   ,INVMF.MF001,INVMF.MF002     ");
                FASTSQL.AppendFormat(@"  HAVING SUM(MF008*MF010)>0");
                FASTSQL.AppendFormat(@"  ORDER BY INVMB.MB001,INVMB.MB002,INVMB.MB003,INVMC.MC002,CMSMC.MC002      ");
                FASTSQL.AppendFormat(@" ");
                FASTSQL.AppendFormat(@" ");
                FASTSQL.AppendFormat(@" ");
            }
            else
            {
                FASTSQL.AppendFormat(@" SELECT INVMB.MB001 AS '品號',INVMB.MB002 AS '品名',INVMB.MB003 AS '規格',INVMC.MC002 AS '庫別',CMSMC.MC002 AS '庫名',INVMC.MC012 AS '最近入庫日',INVMC.MC013 AS '最近出庫日','' AS '批號',INVMC.MC007 AS '庫存量'");
                FASTSQL.AppendFormat(@" FROM TK..INVMB INVMB ,TK..INVMC INVMC ,TK..CMSMC CMSMC");
                FASTSQL.AppendFormat(@" WHERE INVMB.MB001=INVMC.MC001 AND INVMC.MC002=CMSMC.MC001 AND  INVMC.MC007>0 ");
                FASTSQL.AppendFormat(@" AND (( INVMC.MC012<='{0}') AND ( INVMC.MC013<='{0}') )", StayDay.ToString("yyyyMMdd"));
                FASTSQL.AppendFormat(@" AND INVMC.MC002='{0}'", LA009);
                FASTSQL.AppendFormat(@" ORDER BY INVMB.MB001,INVMB.MB002,INVMB.MB003,INVMC.MC002,CMSMC.MC002");
                FASTSQL.AppendFormat(@" ");
            }





            return FASTSQL.ToString();
        }

        public void SETFASTREPORT2(string LA009,DateTime StayDay)
        {

            string SQL;
            report1 = new Report();
            report1.Load(@"REPORT\庫存呆滯表-依進貨日.frx");

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
            SQL = SETFASETSQL2(LA009,StayDay);

            Table.SelectCommand = SQL;
            report1.Preview = previewControl2;
            report1.Show();

        }

        public string SETFASETSQL2(string LA009, DateTime StayDay)
        {
            StringBuilder FASTSQL = new StringBuilder();
            StringBuilder STRQUERY = new StringBuilder();

            //庫別中的庫存總數量>0才查詢相關批號
            //AND LA001 IN (SELECT LA001 FROM [TK].dbo.INVLA WITH (NOLOCK)  WHERE LA009='{0}' GROUP BY LA001 HAVING SUM(LA005*LA011)<>0)

            FASTSQL.AppendFormat(@" 
                                    SELECT 
                                    庫別,庫名,品號 ,品名,規格,批號,庫存量,庫存金額 
                                    ,進貨製造日期
                                    ,進貨有效日期
                                    ,進貨日
                                    ,進貨單
                                    ,客供製造日期,客供有效日期
                                    ,客供進貨日
                                    ,F製造日期
                                    ,F有效日期
                                    ,F進貨日

                                    FROM (
                                    SELECT 
                                    庫別,庫名,品號 ,品名,規格,批號,庫存量,庫存金額 
                                    ,進貨製造日期
                                    ,進貨有效日期
                                    ,進貨日
                                    ,進貨單
                                    ,客供製造日期,客供有效日期
                                    ,客供進貨日
                                    ,ISNULL(CASE WHEN ISNULL(進貨製造日期,'')<>'' THEN 進貨製造日期 
                                    WHEN  ISNULL(進貨製造日期,'')='' AND ISNULL(客供製造日期,'')<>'' THEN 客供製造日期 
                                    END,'') AS 'F製造日期' 

                                    ,ISNULL(CASE WHEN ISNULL(進貨有效日期,'')<>'' THEN 進貨有效日期 
                                    WHEN  ISNULL(進貨有效日期,'')='' AND ISNULL(客供有效日期,'')<>'' THEN 客供有效日期 
                                    END,'') AS 'F有效日期' 

                                    ,ISNULL(CASE WHEN ISNULL(進貨日,'')<>'' THEN 進貨日 
                                    WHEN  ISNULL(進貨日,'')='' AND ISNULL(客供進貨日,'')<>'' THEN 客供進貨日 
                                    END,'') AS 'F進貨日' 

                                    FROM (
                                    SELECT 庫別,庫名,品號 ,品名,規格,批號,庫存量,庫存金額 
                                    ,ISNULL((SELECT TOP 1 TH117 FROM [TK].dbo.PURTH  WITH (NOLOCK) WHERE TH030='Y' AND TH004=品號 AND TH010=批號 ORDER BY TH002 DESC),'') AS '進貨製造日期'
                                    ,ISNULL((SELECT TOP 1 TH036 FROM [TK].dbo.PURTH  WITH (NOLOCK) WHERE TH030='Y' AND TH004=品號 AND TH010=批號 ORDER BY TH002 DESC),'') AS '進貨有效日期'
                                    ,ISNULL((SELECT TOP 1 TG003 FROM [TK].dbo.PURTH  WITH (NOLOCK),[TK].dbo.PURTG  WITH (NOLOCK) WHERE TG001=TH001 AND TG002=TH002 AND TH030='Y' AND TH004=品號 AND TH010=批號 ORDER BY TG003 DESC),'') AS '進貨日'
                                    ,ISNULL((SELECT TOP 1 TH001+TH002+TH003 FROM [TK].dbo.PURTH  WITH (NOLOCK),[TK].dbo.PURTG  WITH (NOLOCK) WHERE TG001=TH001 AND TG002=TH002 AND TH030='Y' AND TH004=品號 AND TH010=批號 ORDER BY TG003 DESC),'') AS '進貨單'
                                    ,ISNULL((SELECT TOP 1 TB033 FROM [TK].dbo.INVTB  WITH (NOLOCK) WHERE TB001='A11A' AND TB018='Y' AND TB004=品號 AND TB014=批號 ORDER BY TB002 DESC),'') AS '客供製造日期'
                                    ,ISNULL((SELECT TOP 1 TB015 FROM [TK].dbo.INVTB  WITH (NOLOCK) WHERE TB001='A11A' AND TB018='Y' AND TB004=品號 AND TB014=批號 ORDER BY TB002 DESC),'') AS '客供有效日期'
                                    ,ISNULL((SELECT TOP 1 TA003 FROM [TK].dbo.INVTB  WITH (NOLOCK),[TK].dbo.INVTA  WITH (NOLOCK) WHERE TA001=TB001 AND TA002=TB002 AND TB001='A11A' AND TB018='Y' AND TB004=品號 AND TB014=批號 ORDER BY TB002 DESC),'') AS '客供進貨日'

                                    FROM (
                                    SELECT  LA009 AS '庫別', MC002 AS '庫名',LA001 AS '品號' ,MB002 AS '品名',MB003 AS '規格',LA016 AS '批號'  ,CAST(SUM(LA005*LA011) AS DECIMAL(18,4)) AS '庫存量'  ,CAST(SUM(LA005*LA013) AS DECIMAL(18,4)) AS '庫存金額'  
                                    FROM [TK].dbo.INVLA WITH (NOLOCK) 
                                    LEFT JOIN [TK].dbo.INVMB WITH (NOLOCK) ON MB001=LA001 
                                    LEFT JOIN [TK].dbo.CMSMC WITH (NOLOCK) ON MC001=LA009 
                                    WHERE  (LA009='{0}') 
                                    AND LA001 IN (SELECT LA001 FROM [TK].dbo.INVLA WITH (NOLOCK)  WHERE LA009='{0}' GROUP BY LA001 HAVING SUM(LA005*LA011)<>0)

                                    GROUP BY  LA001,LA016,MB002,MB003,LA009,MC002
                                    HAVING SUM(LA005*LA011)<>0
                                    ) AS TEMP
                                    ) AS TEMP2
                                    ) AS TEMP3
                                    WHERE F進貨日<='{1}'
                                    ORDER BY  品號,批號

                                    ", LA009, StayDay.ToString("yyyyMMdd"));





            return FASTSQL.ToString();
        }

        #endregion

        #region BUTTON
        private void button1_Click(object sender, EventArgs e)
        {
            //Search();

            StayDay = dateTimePicker1.Value;
            StayDay = StayDay.AddDays(-1 * Convert.ToDouble(textBox1.Text));

            SETFASTREPORT(comboBox1.SelectedValue.ToString(),StayDay);
        }

        private void button2_Click(object sender, EventArgs e)
        {
            //ExcelExport();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            StayDay = dateTimePicker3.Value;
            StayDay = StayDay.AddDays(-1 * Convert.ToDouble(textBox2.Text));

            SETFASTREPORT2(comboBox2.SelectedValue.ToString(),StayDay);
        }






        #endregion

       
    }
}

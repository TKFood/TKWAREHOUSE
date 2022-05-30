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
    public partial class FrmINVCheck : Form
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

        public FrmINVCheck()
        {
            InitializeComponent();

            comboboxload();
            comboboxload2();
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

            String Sequel = "SELECT MC001,MC001+MC002 AS MC002 FROM [DY].dbo.CMSMC WITH (NOLOCK) ORDER BY MC001";
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

            comboBox2.SelectedValue = "10001     ";

        }

        //public void Search()
        //{
        //    try
        //    {

        //        connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
        //        sqlConn = new SqlConnection(connectionString);

        //        sbSql.Clear();
        //        sbSqlQuery.Clear();
        //        ds.Clear();


        //        if (checkBox1.Checked==true)
        //        {
        //            sbSqlQuery.AppendFormat("AND LA001 IN (SELECT LA001 FROM [TK].dbo.INVLA WITH (NOLOCK) WHERE LA004='{0}'   AND LA009='{1}')", dateTimePicker1.Value.ToString("yyyyMMdd"), comboBox1.SelectedValue.ToString());
        //            sbSqlQuery.AppendFormat("  AND LA004<='{0}'", dateTimePicker1.Value.ToString("yyyyMMdd"));
        //        }
        //        else
        //        {
        //            sbSqlQuery.Append(" ");
        //        }

        //        if (comboBox1.Text.Equals("20006     原料倉"))
        //        {
        //            sbSql.Append(@" SELECT  LA001 AS '品號' ,MB002 AS '品名',MB003 AS '規格' ,LA016 AS '批號' ,CAST(SUM(LA005*LA011) AS DECIMAL(18,4)) AS '庫存量'  ");
        //            sbSql.AppendFormat(@"  FROM [{0}].dbo.INVLA WITH (NOLOCK) LEFT JOIN  [{0}].dbo.INVMB WITH (NOLOCK) ON MB001=LA001 ", sqlConn.Database.ToString());
        //            sbSql.AppendFormat(@" WHERE  (LA009='{0}') {1}", comboBox1.SelectedValue.ToString(), sbSqlQuery.ToString());
        //            sbSql.Append(@" GROUP BY  LA001,MB002,MB003,LA016");
        //            sbSql.Append(@" HAVING SUM(LA005*LA011)<>0");
        //            sbSql.Append(@" ORDER BY  LA001,MB002,MB003,LA016");

        //            adapter = new SqlDataAdapter(@"" + sbSql, sqlConn);

        //            sqlCmdBuilder = new SqlCommandBuilder(adapter);
        //            sqlConn.Open();
        //            ds2.Clear();


        //            adapter.Fill(ds2, "TEMPds");

        //            sqlConn.Close();
        //            ds = ds2;
        //        }
        //        else if (comboBox1.Text.Equals("20004     物料倉"))
        //        {
        //            sbSql.Append(@" SELECT  LA001 AS '品號' ,MB002 AS '品名',MB003 AS '規格',LA016 AS '批號'  ,CAST(SUM(LA005*LA011) AS DECIMAL(18,4)) AS '庫存量'  ");
        //            sbSql.AppendFormat(@"  FROM [{0}].dbo.INVLA WITH (NOLOCK) LEFT JOIN  [{0}].dbo.INVMB WITH (NOLOCK) ON MB001=LA001 ", sqlConn.Database.ToString());
        //            sbSql.AppendFormat(@" WHERE  (LA009='{0}') {1}", comboBox1.SelectedValue.ToString(), sbSqlQuery.ToString());
        //            sbSql.Append(@" GROUP BY  LA001,MB002,MB003,LA016");
        //            sbSql.Append(@" HAVING SUM(LA005*LA011)<>0");
        //            sbSql.Append(@" ORDER BY  LA001,MB002,MB003,LA016");

        //            adapter = new SqlDataAdapter(@"" + sbSql, sqlConn);

        //            sqlCmdBuilder = new SqlCommandBuilder(adapter);
        //            sqlConn.Open();
        //            ds2.Clear();


        //            adapter.Fill(ds2, "TEMPds");

        //            sqlConn.Close();
        //            ds = ds2;
        //        }
        //        else if (comboBox1.Text.Equals("20005     半成品倉"))
        //        {
        //            sbSql.Append(@" SELECT  LA001 AS '品號' ,MB002 AS '品名',MB003 AS '規格',LA016 AS '批號'  ,CAST(SUM(LA005*LA011) AS DECIMAL(18,4)) AS '庫存量'  ");
        //            sbSql.AppendFormat(@"  FROM [{0}].dbo.INVLA WITH (NOLOCK) LEFT JOIN  [{0}].dbo.INVMB WITH (NOLOCK) ON MB001=LA001 ", sqlConn.Database.ToString());
        //            sbSql.AppendFormat(@" WHERE  (LA009='{0}') {1}", comboBox1.SelectedValue.ToString(), sbSqlQuery.ToString());
        //            sbSql.Append(@" GROUP BY  LA001,MB002,MB003,LA016");
        //            sbSql.Append(@" HAVING SUM(LA005*LA011)<>0");
        //            sbSql.Append(@" ORDER BY  LA001,MB002,MB003,LA016");

        //            adapter = new SqlDataAdapter(@"" + sbSql, sqlConn);

        //            sqlCmdBuilder = new SqlCommandBuilder(adapter);
        //            sqlConn.Open();
        //            ds2.Clear();


        //            adapter.Fill(ds2, "TEMPds");

        //            sqlConn.Close();
        //            ds = ds2;
        //        }
        //        else
        //        {
        //            sbSql.Append(@" SELECT  LA001 AS '品號' ,MB002 AS '品名',MB003 AS '規格',LA016 AS '批號'  ,CAST(SUM(LA005*LA011) AS INT) AS '庫存量'  ");
        //            sbSql.AppendFormat(@"  FROM [{0}].dbo.INVLA WITH (NOLOCK) LEFT JOIN  [{0}].dbo.INVMB WITH (NOLOCK) ON MB001=LA001 ", sqlConn.Database.ToString());
        //            sbSql.AppendFormat(@" WHERE  (LA009='{0}') {1}", comboBox1.SelectedValue.ToString(), sbSqlQuery.ToString());
        //            sbSql.Append(@" GROUP BY  LA001,MB002,MB003,LA016");
        //            sbSql.Append(@" HAVING SUM(LA005*LA011)<>0");
        //            sbSql.Append(@" ORDER BY  LA001,MB002,MB003,LA016");

        //            adapter = new SqlDataAdapter(@"" + sbSql, sqlConn);

        //            sqlCmdBuilder = new SqlCommandBuilder(adapter);
        //            sqlConn.Open();
        //            ds.Clear();


        //            adapter.Fill(ds, "TEMPds");

        //            sqlConn.Close();


        //        }





        //        if (ds.Tables["TEMPds"].Rows.Count == 0)
        //        {
        //            label14.Text = "找不到資料";
        //        }
        //        else
        //        {
        //            label14.Text = "有 " + ds.Tables["TEMPds"].Rows.Count.ToString() + " 筆";

        //dataGridView1.DataSource = ds.Tables["TEMPds"];
        //dataGridView1.AutoResizeColumns();
        //        }

        //    }
        //    catch
        //    {

        //    }
        //    finally
        //    {
        //        sqlConn.Close();
        //    }
        //}

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
        //    dt = ds.Tables["TEMPds"];

        //    if (dt.TableName != string.Empty)
        //    {
        //        ws = wb.CreateSheet(dt.TableName);
        //    }
        //    else
        //    {
        //        ws = wb.CreateSheet("Sheet1");
        //    }

        //    ws.CreateRow(0);
        //    //第一行為表名稱
        //    ws.GetRow(0).CreateCell(0).SetCellValue("老楊食品大林廠20001-庫存表       年     月     日");
        //    ws.AddMergedRegion(new CellRangeAddress(0, 0, 0, 5));
        //    //第一行為欄位名稱
        //    ws.CreateRow(1);
        //    ws.GetRow(1).CreateCell(0).SetCellValue("品名");
        //    ws.GetRow(1).CreateCell(1).SetCellValue("規格");
        //    ws.GetRow(1).CreateCell(2).SetCellValue("批號");
        //    ws.GetRow(1).CreateCell(3).SetCellValue("數量");
        //    ws.GetRow(1).CreateCell(4).SetCellValue("品名");
        //    ws.GetRow(1).CreateCell(5).SetCellValue("規格");
        //    ws.GetRow(1).CreateCell(6).SetCellValue("批號");
        //    ws.GetRow(1).CreateCell(7).SetCellValue("數量");
        //    ws.GetRow(1).CreateCell(8).SetCellValue("品名");
        //    ws.GetRow(1).CreateCell(9).SetCellValue("規格");
        //    ws.GetRow(1).CreateCell(10).SetCellValue("批號");
        //    ws.GetRow(1).CreateCell(11).SetCellValue("數量");


        //    //for (int i = 1; i < dt.Rows.Count; i++)
        //    //{
        //    //    ws.CreateRow(i + 1);
        //    //    for (int j = 0; j < dt.Columns.Count; j++)
        //    //    {
        //    //        ws.GetRow(i + 1).CreateCell(j).SetCellValue(dt.Rows[i][j].ToString());
        //    //    }
        //    //}

        //    if(comboBox1.Text.Equals("20006     原料倉"))
        //    {
        //        int j = 1;
        //        int k = 0;
        //        if (dt.Rows.Count <= 40)
        //        {
        //            for (int i = 0; i < dt.Rows.Count; i++)
        //            {
        //                ws.CreateRow(j + 1);
        //                ws.GetRow(j + 1).CreateCell(k).SetCellValue(dt.Rows[i][1].ToString());
        //                ws.GetRow(j + 1).CreateCell(k + 1).SetCellValue(dt.Rows[i][2].ToString());
        //                ws.GetRow(j + 1).CreateCell(k + 2).SetCellValue(dt.Rows[i][3].ToString());
        //                ws.GetRow(j + 1).CreateCell(k + 3).SetCellValue(Convert.ToDouble(dt.Rows[i][4].ToString()));

        //                j++;
        //            }

        //        }
        //        else if (dt.Rows.Count <= 80 && dt.Rows.Count >= 41)
        //        {
        //            for (int i = 0; i <= 40; i++)
        //            {
        //                ws.CreateRow(j + 1);
        //                ws.GetRow(j + 1).CreateCell(k).SetCellValue(dt.Rows[i][1].ToString());
        //                ws.GetRow(j + 1).CreateCell(k + 1).SetCellValue(dt.Rows[i][2].ToString());
        //                ws.GetRow(j + 1).CreateCell(k + 2).SetCellValue(dt.Rows[i][3].ToString());
        //                ws.GetRow(j + 1).CreateCell(k + 3).SetCellValue(Convert.ToDouble(dt.Rows[i][4].ToString()));
        //                if ((i + 41) < dt.Rows.Count)
        //                {
        //                    ws.GetRow(j + 1).CreateCell(k + 4).SetCellValue(dt.Rows[i + 41][1].ToString());
        //                    ws.GetRow(j + 1).CreateCell(k + 5).SetCellValue(dt.Rows[i + 41][2].ToString());
        //                    ws.GetRow(j + 1).CreateCell(k + 6).SetCellValue(dt.Rows[i + 41][3].ToString());
        //                    ws.GetRow(j + 1).CreateCell(k + 7).SetCellValue(Convert.ToDouble(dt.Rows[i + 41][4].ToString()));
        //                }


        //                j++;
        //            }
        //        }

        //        else if (dt.Rows.Count <= 120 && dt.Rows.Count >= 81)
        //        {
        //            for (int i = 0; i <= 40; i++)
        //            {
        //                ws.CreateRow(j + 1);
        //                ws.GetRow(j + 1).CreateCell(k).SetCellValue(dt.Rows[i][1].ToString());
        //                ws.GetRow(j + 1).CreateCell(k + 1).SetCellValue(dt.Rows[i][2].ToString());
        //                ws.GetRow(j + 1).CreateCell(k + 2).SetCellValue(dt.Rows[i][3].ToString());
        //                ws.GetRow(j + 1).CreateCell(k + 3).SetCellValue(Convert.ToDouble(dt.Rows[i][4].ToString()));
        //                if ((i + 41) < dt.Rows.Count)
        //                {
        //                    ws.GetRow(j + 1).CreateCell(k + 4).SetCellValue(dt.Rows[i + 41][1].ToString());
        //                    ws.GetRow(j + 1).CreateCell(k + 5).SetCellValue(dt.Rows[i + 41][2].ToString());
        //                    ws.GetRow(j + 1).CreateCell(k + 6).SetCellValue(dt.Rows[i + 41][3].ToString());
        //                    ws.GetRow(j + 1).CreateCell(k + 7).SetCellValue(Convert.ToDouble(dt.Rows[i + 41][4].ToString()));
        //                }
        //                if ((i + 82) < dt.Rows.Count)
        //                {
        //                    ws.GetRow(j + 1).CreateCell(k + 8).SetCellValue(dt.Rows[i + 82][1].ToString());
        //                    ws.GetRow(j + 1).CreateCell(k + 9).SetCellValue(dt.Rows[i + 82][2].ToString());
        //                    ws.GetRow(j + 1).CreateCell(k + 10).SetCellValue(dt.Rows[i + 82][3].ToString());
        //                    ws.GetRow(j + 1).CreateCell(k + 11).SetCellValue(Convert.ToDouble(dt.Rows[i + 82][4].ToString()));
        //                }
        //                j++;
        //            }
        //        }
        //        else
        //        {
        //            foreach (DataGridViewRow dr in this.dataGridView1.Rows)
        //            {
        //                ws.CreateRow(j + 1);
        //                ws.GetRow(j + 1).CreateCell(k).SetCellValue(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[1].ToString());
        //                ws.GetRow(j + 1).CreateCell(k + 1).SetCellValue(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[2].ToString());
        //                ws.GetRow(j + 1).CreateCell(k + 2).SetCellValue(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[3].ToString());
        //                ws.GetRow(j + 1).CreateCell(k + 3).SetCellValue(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[4].ToString());

        //                j++;
        //            }
        //        }
        //    }
        //    else
        //    {
        //        int j = 1;
        //        int k = 0;
        //        if (dt.Rows.Count <= 40)
        //        {
        //            for (int i = 0; i < dt.Rows.Count; i++)
        //            {
        //                ws.CreateRow(j + 1);
        //                ws.GetRow(j + 1).CreateCell(k).SetCellValue(dt.Rows[i][1].ToString());
        //                ws.GetRow(j + 1).CreateCell(k + 1).SetCellValue(dt.Rows[i][2].ToString());
        //                ws.GetRow(j + 1).CreateCell(k + 2).SetCellValue(dt.Rows[i][3].ToString());
        //                ws.GetRow(j + 1).CreateCell(k + 3).SetCellValue(dt.Rows[i][4].ToString());
        //                //ws.GetRow(j + 1).CreateCell(k + 2).SetCellValue(Convert.ToInt32(dt.Rows[i][3].ToString()));

        //                j++;
        //            }

        //        }
        //        else if (dt.Rows.Count <= 80 && dt.Rows.Count >= 41)
        //        {
        //            for (int i = 0; i <= 40; i++)
        //            {
        //                ws.CreateRow(j + 1);
        //                ws.GetRow(j + 1).CreateCell(k).SetCellValue(dt.Rows[i][1].ToString());
        //                ws.GetRow(j + 1).CreateCell(k + 1).SetCellValue(dt.Rows[i][2].ToString());
        //                //ws.GetRow(j + 1).CreateCell(k + 2).SetCellValue(Convert.ToInt32(dt.Rows[i][3].ToString()));
        //                ws.GetRow(j + 1).CreateCell(k + 2).SetCellValue(dt.Rows[i][3].ToString());
        //                ws.GetRow(j + 1).CreateCell(k + 3).SetCellValue(dt.Rows[i][4].ToString());
        //                if ((i + 41) < dt.Rows.Count)
        //                {
        //                    ws.GetRow(j + 1).CreateCell(k + 4).SetCellValue(dt.Rows[i + 41][1].ToString());
        //                    ws.GetRow(j + 1).CreateCell(k + 5).SetCellValue(dt.Rows[i + 41][2].ToString());
        //                    ws.GetRow(j + 1).CreateCell(k + 6).SetCellValue(dt.Rows[i+41][3].ToString());
        //                    ws.GetRow(j + 1).CreateCell(k + 7).SetCellValue(dt.Rows[i + 41][4].ToString());
        //                    //ws.GetRow(j + 1).CreateCell(k + 5).SetCellValue(Convert.ToInt32(dt.Rows[i + 41][3].ToString()));
        //                }


        //                j++;
        //            }
        //        }

        //        else if (dt.Rows.Count <= 120 && dt.Rows.Count >= 81)
        //        {
        //            for (int i = 0; i <= 40; i++)
        //            {
        //                ws.CreateRow(j + 1);
        //                ws.GetRow(j + 1).CreateCell(k).SetCellValue(dt.Rows[i][1].ToString());
        //                ws.GetRow(j + 1).CreateCell(k + 1).SetCellValue(dt.Rows[i][2].ToString());
        //                //ws.GetRow(j + 1).CreateCell(k + 2).SetCellValue(Convert.ToInt32(dt.Rows[i][3].ToString()));
        //                ws.GetRow(j + 1).CreateCell(k + 2).SetCellValue(dt.Rows[i][3].ToString());
        //                ws.GetRow(j + 1).CreateCell(k + 3).SetCellValue(dt.Rows[i][4].ToString());
        //                if ((i + 41) < dt.Rows.Count)
        //                {
        //                    ws.GetRow(j + 1).CreateCell(k + 4).SetCellValue(dt.Rows[i + 41][1].ToString());
        //                    ws.GetRow(j + 1).CreateCell(k + 5).SetCellValue(dt.Rows[i + 41][2].ToString());
        //                    ws.GetRow(j + 1).CreateCell(k + 6).SetCellValue(dt.Rows[i + 41][3].ToString());
        //                    ws.GetRow(j + 1).CreateCell(k + 7).SetCellValue(dt.Rows[i + 41][4].ToString());
        //                    //ws.GetRow(j + 1).CreateCell(k + 5).SetCellValue(Convert.ToInt32(dt.Rows[i + 41][3].ToString()));
        //                }
        //                if ((i + 81) < dt.Rows.Count)
        //                {
        //                    ws.GetRow(j + 1).CreateCell(k + 8).SetCellValue(dt.Rows[i + 81][1].ToString());
        //                    ws.GetRow(j + 1).CreateCell(k + 9).SetCellValue(dt.Rows[i + 81][2].ToString());
        //                    ws.GetRow(j + 1).CreateCell(k + 10).SetCellValue(dt.Rows[i + 81][3].ToString());
        //                    ws.GetRow(j + 1).CreateCell(k + 11).SetCellValue(dt.Rows[i + 81][4].ToString());
        //                    //ws.GetRow(j + 1).CreateCell(k + 8).SetCellValue(Convert.ToInt32(dt.Rows[i + 81][3].ToString()));
        //                }
        //                j++;
        //            }
        //        }
        //        else
        //        {
        //            foreach (DataGridViewRow dr in this.dataGridView1.Rows)
        //            {
        //                ws.CreateRow(j + 1);
        //                ws.GetRow(j + 1).CreateCell(k).SetCellValue(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[1].ToString());
        //                ws.GetRow(j + 1).CreateCell(k + 1).SetCellValue(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[2].ToString());
        //                ws.GetRow(j + 1).CreateCell(k + 2).SetCellValue(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[3].ToString());
        //                ws.GetRow(j + 1).CreateCell(k + 3).SetCellValue(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[4].ToString());

        //                j++;
        //            }
        //        }
        //    }





        //    //int j = 1;
        //    //int k = 0;
        //    //foreach (DataGridViewRow dr in this.dataGridView1.Rows)
        //    //{
        //    //    ws.CreateRow(j + 1);
        //    //    ws.GetRow(j + 1).CreateCell(k).SetCellValue(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[1].ToString());
        //    //    ws.GetRow(j + 1).CreateCell(k + 1).SetCellValue(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[2].ToString());
        //    //    ws.GetRow(j + 1).CreateCell(k + 2).SetCellValue(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[1].ToString());

        //    //    j++;
        //    //}

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
        //    filename.AppendFormat(@"c:\temp\老楊食品大林廠20001-庫存表{0}.xlsx", DateTime.Now.ToString("yyyyMMdd"));

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

        public void SETFASTREPORT()
        { 

            string SQL;
            report1 = new Report();

            if (comboBox1.Text.Equals("20001     成品倉"))
            {
                report1.Load(@"REPORT\每日盤點表-成品.frx");
            }
            else if (comboBox1.Text.Equals("21001     方城市銷售倉"))
            {
                report1.Load(@"REPORT\每日盤點表-成品-21001.frx");
            }
            else if(comboBox1.Text.Equals("20006     原料倉"))
            {
                report1.Load(@"REPORT\每日盤點表-原料.frx");
            }
            else if (comboBox1.Text.Equals("20004     物料倉"))
            {
                report1.Load(@"REPORT\每日盤點表-物料.frx");
            }
            else if (comboBox1.Text.Equals("20005     半成品倉"))
            {
                report1.Load(@"REPORT\每日盤點表-半成品.frx");
            }
            else
            {
                report1.Load(@"REPORT\每日盤點表.frx");
            }



            //reprot

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

            FASTSQL.Clear();
            sbSqlQuery.Clear();

            DateTime dt = DateTime.Now;
            dt = dt.AddMonths(-2);

            if (checkBox1.Checked == true)
            {
                sbSqlQuery.Clear();

                sbSqlQuery.AppendFormat(" AND LA001 IN (SELECT LA001 FROM [TK].dbo.INVLA WITH (NOLOCK) WHERE LA004='{0}'   AND LA009='{1}')", dateTimePicker1.Value.ToString("yyyyMMdd"), comboBox1.SelectedValue.ToString());
                sbSqlQuery.AppendFormat("  AND LA004<='{0}'", dateTimePicker1.Value.ToString("yyyyMMdd"));
            }
            else
            {
                sbSqlQuery.Clear();
                sbSqlQuery.Append(" ");
            }

            if (comboBox1.Text.Equals("20001     成品倉"))
            {
               
                FASTSQL.AppendFormat(@"  
                                     SELECT 品號,品名,規格,批號,庫存量,單位,效期內的訂單需求量,效期內的訂單差異量,總訂單需求量,業務
                                    ,CASE WHEN ISNULL(生產日期,'')<>'' THEN 生產日期 ELSE CONVERT(nvarchar,DATEADD(day,-1*(CASE WHEN MB198='1' THEN 1*MB023 WHEN  MB198='2' THEN 30*MB023  ELSE 0 END),CONVERT(datetime,批號)),112) END AS '生產日期'
                                    ,CASE WHEN ISNULL(在倉日期,'')<>'' THEN 在倉日期 ELSE DATEDIFF(DAY,CASE WHEN ISNULL(生產日期,'')<>'' THEN 生產日期 ELSE CONVERT(nvarchar,DATEADD(day,-1*(CASE WHEN MB198='1' THEN 1*MB023 WHEN  MB198='2' THEN 30*MB023  ELSE 0 END),CONVERT(datetime,批號)),112) END,'{0}') END AS '在倉日期'
                                    ,有效天數
                                    ,狀態
                                    ,CASE WHEN MB198='1' THEN 1*MB023 WHEN  MB198='2' THEN 30*MB023  ELSE 0 END AS 'DAYS'
                                    ,CONVERT(nvarchar,DATEADD(day,-1*(CASE WHEN MB198='1' THEN 1*MB023 WHEN  MB198='2' THEN 30*MB023  ELSE 0 END),CONVERT(datetime,批號)),112) AS '外購品的生產日'

                                    FROM (
                                    SELECT 品號,品名,規格,批號,庫存量,單位,效期內的訂單需求量,效期內的訂單差異量,總訂單需求量,業務
                                    ,生產日期
                                    ,DATEDIFF(DAY,生產日期,'{0}') AS '在倉日期'
                                    ,DATEDIFF(DAY,'{0}',有效日期NEW)  AS '有效天數'
                                    ,(CASE WHEN DATEDIFF(DAY,生產日期,'{0}')>90 THEN '在倉超過90天' ELSE (CASE WHEN DATEDIFF(DAY,生產日期,'{0}')>30 THEN '在倉超過30天' ELSE '' END) END ) AS '狀態'
                                    FROM ( 
                                    SELECT   LA001 AS '品號' ,MB002 AS '品名',MB003 AS '規格',LA016 AS '批號'
                                    ,CAST(SUM(LA005*LA011) AS INT) AS '庫存量',MB004 AS '單位'
                                    ,CAST(((SELECT ISNULL(SUM(NUM),0) FROM [TK].dbo.VCOPTDINVMD WHERE TD004=LA001 AND TD013>='20210812' AND  TD013<=CONVERT(nvarchar,DATEADD (MONTH,-1*ROUND(MB023/3,0),CAST(LA016 AS datetime)),112))) AS INT) AS '效期內的訂單需求量'     
                                    ,CAST((CAST(SUM(LA005*LA011) AS INT)-(SELECT ISNULL(SUM(NUM),0) FROM [TK].dbo.VCOPTDINVMD WHERE TD004=LA001 AND TD013>='20210812' AND  TD013<=CONVERT(nvarchar,DATEADD (MONTH,-1*ROUND(MB023/3,0),CAST(LA016 AS datetime)),112)))  AS INT) AS '效期內的訂單差異量' 
                                    ,CAST((SELECT ISNULL(SUM(NUM),0) FROM [TK].dbo.VCOPTDINVMD WHERE TD004=LA001 AND TD013>='20210812') AS INT) AS '總訂單需求量' 
                                    ,(SELECT TOP 1 TC006+' '+MV002 FROM [TK].dbo.COPTC,[TK].dbo.CMSMV WHERE TC006=MV001 AND  TC001+TC002 IN (SELECT TOP 1 TA026+TA027 FROM [TK].dbo.MOCTA WHERE TA001+TA002 IN (SELECT TOP 1 TG014+TG015 FROM [TK].dbo.MOCTG WHERE TG004=LA001 AND TG017=LA016))) AS '業務'
                                    ,(SELECT TOP 1 TG040 FROM [TK].dbo.MOCTF,[TK].dbo.MOCTG WHERE TF001=TG001 AND TF002=TG002 AND TG004=LA001 AND TG017=LA016 ORDER BY TF003 ASC) AS '生產日期'
                                    ,(SELECT TOP 1 TG018 FROM [TK].dbo.MOCTF,[TK].dbo.MOCTG WHERE TF001=TG001 AND TF002=TG002 AND TG004=LA001 AND TG017=LA016 ORDER BY TF003 ASC) AS '有效日期'
                                    ,(CASE WHEN ISNULL((SELECT TOP 1 TG040 FROM [TK].dbo.MOCTF,[TK].dbo.MOCTG WHERE TG001=TF001 AND TG002=TF002 AND TG010='20005' AND TG004=LA001 AND TG017=LA016  ORDER BY TG040),'')<>'' THEN (SELECT TOP 1 TG040 FROM [TK].dbo.MOCTF,[TK].dbo.MOCTG WHERE TG001=TF001 AND TG002=TF002 AND TG010='20005' AND TG004=LA001 AND TG017=LA016  ORDER BY TG040) ELSE LA016 END) AS '有效日期NEW'
                                    ,ISDATE(LA016) AS LA016
                                    FROM [TK].dbo.INVLA WITH (NOLOCK)  
                                    LEFT JOIN  [TK].dbo.INVMB WITH (NOLOCK) ON MB001=LA001   
                                    WHERE  (LA009='20001')   
                                    AND (LA001 LIKE '4%' OR LA001 LIKE '5%')
                                    GROUP BY  LA001,LA009,MB002,MB003,LA016,MB023,MB198,MB004    
                                    HAVING SUM(LA005*LA011)<>0 
                                    ) AS TEMP
                                    ) AS TEMP2
                                    LEFT JOIN [TK].dbo.INVMB ON MB001=品號
                                    ORDER BY 品號,批號       

                                        ", DateTime.Now.ToString("yyyyMMdd"));
            }
            else if (comboBox1.Text.Equals("21001     方城市銷售倉"))
            {
                FASTSQL.AppendFormat(@" SELECT 品號,品名,規格,批號,庫存量,單位,效期內的訂單需求量,效期內的訂單差異量,在倉日期,有效天數,總訂單需求量,業務");
                FASTSQL.AppendFormat(@" FROM (");
                FASTSQL.AppendFormat(@" ");
                FASTSQL.AppendFormat(@" SELECT   LA001 AS '品號' ,MB002 AS '品名',MB003 AS '規格',LA016 AS '批號'");
                FASTSQL.AppendFormat(@" ,CAST(SUM(LA005*LA011) AS INT) AS '庫存量',MB004 AS '單位'");
                FASTSQL.AppendFormat(@" ,0 AS '效期內的訂單需求量' ");
                FASTSQL.AppendFormat(@" ,0 AS '效期內的訂單差異量' ");
                //FASTSQL.AppendFormat(@" ,DATEDIFF(DAY,(SELECT TOP 1 TF003 FROM [TK].dbo.MOCTF,[TK].dbo.MOCTG WHERE TF001=TG001 AND TF002=TG002 AND TG004=LA001 AND TG017=LA016 AND TG010=LA009) , '{0}' ) AS '在倉日期' ", DateTime.Now.ToString("yyyyMMdd"));
                //FASTSQL.AppendFormat(@" ,CASE WHEN DATEDIFF(DAY,(SELECT TOP 1 TF003 FROM [TK].dbo.MOCTF,[TK].dbo.MOCTG WHERE TF001=TG001 AND TF002=TG002 AND TG004=LA001 AND TG017=LA016 AND TG010=LA009) , '{0}' )<> NULL THEN DATEDIFF(DAY,(SELECT TOP 1 TF003 FROM [TK].dbo.MOCTF,[TK].dbo.MOCTG WHERE TF001=TG001 AND TF002=TG002 AND TG004=LA001 AND TG017=LA016 AND TG010=LA009) , '{0}' ) ELSE DATEDIFF(DAY,(SELECT TOP 1 TF003 FROM [TK].dbo.MOCTF,[TK].dbo.MOCTG WHERE TF001=TG001 AND TF002=TG002 AND TG004=LA001 AND TG017=LA016) , '{0}' ) END  AS '在倉日期' ", DateTime.Now.ToString("yyyyMMdd"));
                FASTSQL.AppendFormat(@" ,ISNULL ( DATEDIFF(DAY,(SELECT TOP 1 TF003 FROM [TK].dbo.MOCTF,[TK].dbo.MOCTG WHERE TF001=TG001 AND TF002=TG002 AND TG004=LA001 AND TG017=LA016 AND TG010=LA009),'{0}'),DATEDIFF(DAY,(SELECT TOP 1 LA004 FROM [TK].dbo.INVLA A WHERE A.LA001=INVLA.LA001 AND A.LA016=INVLA.LA016 AND A.LA005='1') ,'{0}') ) AS '在倉日期' ", DateTime.Now.ToString("yyyyMMdd"));
                FASTSQL.AppendFormat(@" ,(CASE WHEN isdate(LA016)=1 THEN DATEDIFF(DAY, '{0}',LA016  ) ELSE 0 END )  AS '有效天數'   ", DateTime.Now.ToString("yyyyMMdd"));
                FASTSQL.AppendFormat(@" ,0 AS '總訂單需求量' ");
                FASTSQL.AppendFormat(@" ,NULL AS '業務'");
                FASTSQL.AppendFormat(@" FROM [TK].dbo.INVLA WITH (NOLOCK)  ");
                FASTSQL.AppendFormat(@" LEFT JOIN  [TK].dbo.INVMB WITH (NOLOCK) ON MB001=LA001   ");
                FASTSQL.AppendFormat(@" WHERE  (LA009='21001')   ");
                FASTSQL.AppendFormat(@" AND (LA001 LIKE '4%' OR LA001 LIKE '5%')");
                FASTSQL.AppendFormat(@" GROUP BY  LA001,LA009,MB002,MB003,LA016,MB023,MB198,MB004    ");
                FASTSQL.AppendFormat(@" HAVING SUM(LA005*LA011)<>0 ");
                FASTSQL.AppendFormat(@" ) AS TEMP");
                FASTSQL.AppendFormat(@" ORDER BY 品號,批號 ");
                FASTSQL.AppendFormat(@" ");
            }
            else if (comboBox1.Text.Equals("20006     原料倉"))
            {
                FASTSQL.AppendFormat(@"  
                                     SELECT 品號,品名,規格,批號,庫存量,單位,庫存金額,在倉日期,有效天數,業務
                                     
                                     FROM (
                                     SELECT   LA001 AS '品號' ,MB002 AS '品名',MB003 AS '規格',LA016 AS '批號'
                                     ,CONVERT(DECIMAL(16,3),SUM(LA005*LA011)) AS '庫存量',MB004 AS '單位'
                                     ,CONVERT(DECIMAL(16,3),SUM(LA005*LA013)) AS '庫存金額'
                                     ,DATEDIFF(DAY,LA016,'{0}') AS '在倉日期old' 
                                     ,DATEDIFF(DAY,(SELECT TOP 1 TH014 FROM [TK].dbo.PURTG,[TK].dbo.PURTH WHERE TG001=TH001 AND TG002=TH002 AND TH004=LA001 AND TH010=LA016 ),'{0}') AS '在倉日期'  
                                     ,DATEDIFF(DAY,'{0}',LA016) AS '有效天數'
                                     ,(SELECT TOP 1 TC006+' '+MV002 FROM [TK].dbo.COPTC,[TK].dbo.CMSMV WHERE TC006=MV001 AND  TC001+TC002 IN (SELECT TOP 1 TA026+TA027 FROM [TK].dbo.MOCTA WHERE TA001+TA002 IN (SELECT TOP 1 TG014+TG015 FROM [TK].dbo.MOCTG WHERE TG004=LA001 AND TG017=LA016))) AS '業務'
                                      ,ISDATE(LA016) AS LA016
                                     FROM [TK].dbo.INVLA WITH (NOLOCK)  
                                     LEFT JOIN  [TK].dbo.INVMB WITH (NOLOCK) ON MB001=LA001   
                                     WHERE  (LA009='20006')   
                                      {1}
                                     GROUP BY  LA001,LA009,MB002,MB003,LA016,MB023,MB198,MB004   
                                     HAVING SUM(LA005*LA011)<>0 
                                     ) AS TEMP
                                     WHERE 品號 NOT IN ('122221001','114141009')
                                     ORDER BY 品號,批號  
                                    ", dateTimePicker1.Value.ToString("yyyyMMdd"), sbSqlQuery.ToString());


                FASTSQL.AppendFormat(@" ");

            }
            else if (comboBox1.Text.Equals("20004     物料倉"))
            {
           
                FASTSQL.AppendFormat(@"                                     
                                    SELECT 品號,品名,規格,批號,庫存量,庫存金額
                                    ,製造日期,有效日期
                                    , 客供製造日期,客供有效日期
                                    ,(製造日期+客供製造日期) AS 'F製造日期',(有效日期+客供有效日期)  AS 'F有效日期'
                                    FROM ( 
                                    SELECT  LA001 AS '品號' ,MB002 AS '品名',MB003 AS '規格',LA016 AS '批號'  ,CAST(SUM(LA005*LA011) AS DECIMAL(18,4)) AS '庫存量'  ,CAST(SUM(LA005*LA013) AS DECIMAL(18,4)) AS '庫存金額'  
                                    ,ISNULL(TH117,'') AS '製造日期' ,ISNULL(TH036,'') AS '有效日期'
                                    ,ISNULL(TB033,'') AS '客供製造日期' ,ISNULL(TB015 ,'')AS '客供有效日期'
                                    FROM [TK].dbo.INVLA WITH (NOLOCK) 
                                    LEFT JOIN [TK].dbo.PURTH WITH (NOLOCK) ON TH030='Y' AND TH004=LA001 AND TH010=LA016
                                    LEFT JOIN [TK].dbo.INVTB WITH (NOLOCK) ON TB001='A11A' AND TB018='Y' AND TB004=LA001 AND TB014=LA016
                                    LEFT JOIN [TK].dbo.INVMB WITH (NOLOCK) ON MB001=LA001 
                                    WHERE  (LA009='{0}') 
                                    {1}
                                    GROUP BY  LA001,MB002,MB003,LA016,TH117,TH036, TB033,TB015
                                    HAVING SUM(LA005*LA011)<>0
                                    ) AS TEMP 
                                    ORDER BY  品號,品名,規格,批號

                                    ", comboBox1.SelectedValue.ToString(), sbSqlQuery.ToString());

                FASTSQL.AppendFormat(@" ");
            }
            else if (comboBox1.Text.Equals("20005     半成品倉"))
            {

                FASTSQL.AppendFormat(@"   
                                   SELECT 品號,品名,規格,批號,庫存量,單位,製造日期NEW
                                    , DATEDIFF(DAY,製造日期NEW,'{0}') AS '在倉日期'
                                    ,(CASE WHEN MB198='2' THEN DATEDIFF(DAY,'{0}',DATEADD(month, MB023, '{0}' )) END)-(CASE WHEN DATEDIFF(DAY,製造日期NEW,'{0}')>=0 THEN DATEDIFF(DAY,製造日期NEW,'{0}') ELSE (CASE WHEN DATEDIFF(DAY,製造日期NEW,'{0}')<0 THEN  (CASE WHEN MB198='2' THEN DATEDIFF(DAY,DATEADD(month, -1*MB023, 製造日期NEW ),'{0}') END ) END ) END)  AS '有效天數' 
                                    ,業務
                                    ,(庫存量*(SELECT MB065/MB064 FROM [TK].dbo.INVMB WHERE MB001=品號))AS 庫存金額
                                    ,(CASE WHEN DATEDIFF(DAY,製造日期NEW,'{0}')>90 THEN '在倉超過90天' ELSE (CASE WHEN DATEDIFF(DAY,製造日期NEW,'{0}')>30 THEN '在倉超過30天' ELSE '' END) END ) AS '狀態'
                                    FROM (
                                    SELECT   LA001 AS '品號' ,INVMB.MB002 AS '品名',INVMB.MB003 AS '規格',LA016 AS '批號'
                                    ,CONVERT(DECIMAL(16,3),SUM(LA005*LA011)) AS '庫存量',INVMB.MB004 AS '單位',MB198,MB023
                                    ,(SELECT TOP 1 TG040 FROM [TK].dbo.MOCTF,[TK].dbo.MOCTG WHERE TG001=TF001 AND TG002=TF002 AND TG010='20005' AND TG004=LA001 AND TG017=LA016  ORDER BY TG040) AS '製造日期'
                                    ,(SELECT TOP 1 TG018 FROM [TK].dbo.MOCTF,[TK].dbo.MOCTG WHERE TG001=TF001 AND TG002=TF002 AND TG010='20005' AND TG004=LA001 AND TG017=LA016  ORDER BY TG018) AS '有效日期'
                                    ,(SELECT TOP 1 TC006+' '+MV002 FROM [TK].dbo.COPTC,[TK].dbo.CMSMV WHERE TC006=MV001 AND  TC001+TC002 IN (SELECT TOP 1 TA026+TA027 FROM [TK].dbo.MOCTA WHERE TA001+TA002 IN (SELECT TOP 1 TG014+TG015 FROM [TK].dbo.MOCTG WHERE TG004=LA001 AND TG017=LA016))) AS '業務'
                                    ,(CASE WHEN ISNULL((SELECT TOP 1 TG040 FROM [TK].dbo.MOCTF,[TK].dbo.MOCTG WHERE TG001=TF001 AND TG002=TF002 AND TG010='20005' AND TG004=LA001 AND TG017=LA016  ORDER BY TG040),'')<>'' THEN (SELECT TOP 1 TG040 FROM [TK].dbo.MOCTF,[TK].dbo.MOCTG WHERE TG001=TF001 AND TG002=TF002 AND TG010='20005' AND TG004=LA001 AND TG017=LA016  ORDER BY TG040) ELSE LA016 END) AS '製造日期NEW'
                                      ,ISDATE(LA016) AS LA016
                                    FROM [TK].dbo.INVLA WITH (NOLOCK) 
                                    LEFT JOIN  [TK].dbo.INVMB WITH (NOLOCK) ON MB001=LA001  
                                    WHERE  (LA009='20005') 

                                    GROUP BY  LA001,LA009,MB002,MB003,LA016,MB023,MB198,MB004
                                    HAVING SUM(LA005*LA011)<>0 
                                    ) AS TEMP
                                    ORDER BY 品號,批號   
                                    ", dateTimePicker1.Value.ToString("yyyyMMdd"));
              


            }
            else
            {
                FASTSQL.AppendFormat(@" SELECT  LA001 AS '品號' ,MB002 AS '品名',MB003 AS '規格',LA016 AS '批號'  ,CAST(SUM(LA005*LA011) AS DECIMAL(18,4)) AS '庫存量',CONVERT(DECIMAL(16,3),SUM(LA005*LA013)) AS '庫存金額'  ");
                FASTSQL.AppendFormat(@"  FROM [TK].dbo.INVLA WITH (NOLOCK) LEFT JOIN  [TK].dbo.INVMB WITH (NOLOCK) ON MB001=LA001 ", sqlConn.Database.ToString());
                FASTSQL.AppendFormat(@" WHERE  (LA009='{0}') {1}", comboBox1.SelectedValue.ToString(), sbSqlQuery.ToString());
                FASTSQL.AppendFormat(@" GROUP BY  LA001,MB002,MB003,LA016");
                FASTSQL.AppendFormat(@" HAVING SUM(LA005*LA011)<>0");
                FASTSQL.AppendFormat(@" ORDER BY  LA001,MB002,MB003,LA016  ");


            }





            return FASTSQL.ToString();
        }

        public void SETFASTREPORT2()
        {
            string SQL2;
            string SQL3;
            report2 = new Report();

            report2.Load(@"REPORT\每日盤點表-成品-訂單明細.frx");

            //reprot

            //20210902密
            Class1 TKID = new Class1();//用new 建立類別實體
            SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dberp"].ConnectionString);

            //資料庫使用者密碼解密
            sqlsb.Password = TKID.Decryption(sqlsb.Password);
            sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

            String connectionString;
            sqlConn = new SqlConnection(sqlsb.ConnectionString);

            report2.Dictionary.Connections[0].ConnectionString = sqlsb.ConnectionString;

            TableDataSource Table = report2.GetDataSource("Table") as TableDataSource;
            
            SQL2 = SETFASETSQL2();
            Table.SelectCommand = SQL2;

            TableDataSource Table1 = report2.GetDataSource("Table1") as TableDataSource;
            SQL3 = SETFASETSQL3();
            Table1.SelectCommand = SQL3;
            report2.Preview = previewControl2;
            report2.Show();

        }

        public string SETFASETSQL2()
        {
            StringBuilder FASTSQL = new StringBuilder();
            StringBuilder STRQUERY = new StringBuilder();

            DateTime dt = DateTime.Now;
            dt = dt.AddMonths(-2);

            FASTSQL.AppendFormat(@" SELECT  LA001 AS '品號' ,MB002 AS '品名',MB003 AS '規格',LA016 AS '批號',CAST(SUM(LA005*LA011) AS INT) AS '庫存量' ,MB004 AS '單位'   ");
            FASTSQL.AppendFormat(@" ,(SELECT ISNULL(SUM(NUM),0) FROM [TK].dbo.VCOPTDINVMD WHERE TD004=LA001 AND TD013>='{0}') AS '訂單需求量'", dt.ToString("yyyyMMdd"));
            FASTSQL.AppendFormat(@" ,(CAST(SUM(LA005*LA011) AS INT)-(SELECT ISNULL(SUM(NUM),0) FROM [TK].dbo.VCOPTDINVMD WHERE TD004=LA001 AND TD013>='{0}')) AS '需求差異量'", dt.ToString("yyyyMMdd"));
            FASTSQL.AppendFormat(@" ,DATEDIFF(DAY,(SELECT TOP 1 LA004 FROM [TK].dbo.INVLA A WHERE A.LA001=INVLA.LA001 AND A.LA016=INVLA.LA016 AND LA005='1') , '{0}' ) AS '在倉日期'", DateTime.Now.ToString("yyyyMMdd"));
            FASTSQL.AppendFormat(@" ,DATEDIFF(DAY, '{0}',LA016  )  AS '有效天數'", DateTime.Now.ToString("yyyyMMdd"));
            FASTSQL.AppendFormat(@" FROM [TK].dbo.INVLA WITH (NOLOCK) ");
            FASTSQL.AppendFormat(@" LEFT JOIN  [TK].dbo.INVMB WITH (NOLOCK) ON MB001=LA001  ");
            FASTSQL.AppendFormat(@" WHERE  (LA009='{0}') {1}", comboBox1.SelectedValue.ToString(), sbSqlQuery.ToString());
            FASTSQL.AppendFormat(@" AND LA001 LIKE '4%'");
            FASTSQL.AppendFormat(@" GROUP BY  LA001,MB002,MB003,LA016,MB023,MB198 ,MB004 ");
            FASTSQL.AppendFormat(@" HAVING SUM(LA005*LA011)<>0 ");
            FASTSQL.AppendFormat(@" ORDER BY  LA001,MB002,MB003,LA016,MB023,MB198 ,MB004");
            FASTSQL.AppendFormat(@" ");
            

            return FASTSQL.ToString();
        }
        public string SETFASETSQL3()
        {
            StringBuilder FASTSQL = new StringBuilder();
            StringBuilder STRQUERY = new StringBuilder();

            DateTime dt = DateTime.Now;
            dt = dt.AddMonths(-2);

            
            FASTSQL.AppendFormat(@" SELECT MV002 AS '業務員',TC053 AS '客戶',TD013 AS '預交日',NUM AS '訂單需求量',TD010 AS '單位',TC001 AS '訂單',TC002 AS '訂單號',TC004 AS '客戶代號',TD004 AS '品號',TD008 AS '訂單下量',TD009 AS '已出量',TD024 AS '贈品量',TD025 AS '已出贈品',MD004 AS '換算'");
            FASTSQL.AppendFormat(@" FROM [TK].dbo.VCOPTDINVMD, [TK].dbo.COPTC");
            FASTSQL.AppendFormat(@" LEFT JOIN [TK].dbo.CMSMV ON MV001=TC006");
            FASTSQL.AppendFormat(@" WHERE TC001=TD001 AND TC002=TD002");
            FASTSQL.AppendFormat(@" AND  TD013>='{0}'", dt.ToString("yyyyMMdd"));
            FASTSQL.AppendFormat(@" ");

            return FASTSQL.ToString();
        }


        public void SETFASTREPORT3()
        {

            string SQL;
            report1 = new Report();

            if (comboBox2.Text.Equals("10001     成品倉"))
            {
                report1.Load(@"REPORT\大潁-每日盤點表-成品.frx");
            }
            else if (comboBox2.Text.Equals("11001     大榮觀音倉"))
            {
                report1.Load(@"REPORT\大潁-11001每日盤點表.frx");
            }
            else
            {
                report1.Load(@"REPORT\大潁-每日盤點表.frx");
            }



            //reprot

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
            SQL = SETFASETSQL4();
            Table.SelectCommand = SQL;
            report1.Preview = previewControl3;
            report1.Show();

        }

        public string SETFASETSQL4()
        {
            StringBuilder FASTSQL = new StringBuilder();
            StringBuilder STRQUERY = new StringBuilder();

            FASTSQL.Clear();
            sbSqlQuery.Clear();

            DateTime dt = DateTime.Now;
            dt = dt.AddMonths(-2);

            if (checkBox1.Checked == true)
            {
                sbSqlQuery.Clear();

                sbSqlQuery.AppendFormat(@" 
                                        AND LA001 IN (SELECT LA001 FROM [DY].dbo.INVLA WITH (NOLOCK) WHERE LA004='{0}'   AND LA009='{1}')
                                        AND LA004<='{2}'"
                                        , dateTimePicker3.Value.ToString("yyyyMMdd"), comboBox2.SelectedValue.ToString(), dateTimePicker3.Value.ToString("yyyyMMdd"));
            }
            else
            {
                sbSqlQuery.Clear();
                sbSqlQuery.Append(" ");
            }

            if (comboBox2.Text.Equals("10001     成品倉"))
            {

                FASTSQL.AppendFormat(@"  
                                     SELECT 品號,品名,規格,批號,庫存量,單位,效期內的訂單需求量,效期內的訂單差異量,在倉日期,有效天數,總訂單需求量,業務
                                     FROM ( 
                                     SELECT   LA001 AS '品號' ,MB002 AS '品名',MB003 AS '規格',LA016 AS '批號'
                                     ,CAST(SUM(LA005*LA011) AS INT) AS '庫存量',MB004 AS '單位'
                                     ,CAST(((SELECT ISNULL(SUM(NUM),0) FROM [DY].dbo.VCOPTDINVMD WHERE TD004=LA001 AND TD013>='{0}' AND  TD013<=CONVERT(nvarchar,DATEADD (MONTH,-1*ROUND(MB023/3,0),CAST(LA016 AS datetime)),112))) AS INT) AS '效期內的訂單需求量'     
                                     ,CAST((CAST(SUM(LA005*LA011) AS INT)-(SELECT ISNULL(SUM(NUM),0) FROM [DY].dbo.VCOPTDINVMD WHERE TD004=LA001 AND TD013>='{0}' AND  TD013<=CONVERT(nvarchar,DATEADD (MONTH,-1*ROUND(MB023/3,0),CAST(LA016 AS datetime)),112)))  AS INT) AS '效期內的訂單差異量' 
                                     ,ISNULL ( DATEDIFF(DAY,(SELECT TOP 1 TF003 FROM [DY].dbo.MOCTF,[DY].dbo.MOCTG WHERE TF001=TG001 AND TF002=TG002 AND TG004=LA001 AND TG017=LA016 AND TG010=LA009),'{0}'),DATEDIFF(DAY,(SELECT TOP 1 LA004 FROM [DY].dbo.INVLA A WHERE A.LA001=INVLA.LA001 AND A.LA016=INVLA.LA016 AND A.LA005='1') ,'{0}') ) AS '在倉日期' 
                                    , DATEDIFF(DAY, '{0}',LA016  )  AS '有效天數' 
                                     ,CAST((SELECT ISNULL(SUM(NUM),0) FROM [DY].dbo.VCOPTDINVMD WHERE TD004=LA001 AND TD013>='{0}') AS INT) AS '總訂單需求量' 
                                     ,(SELECT TOP 1 TC006+' '+MV002 FROM [DY].dbo.COPTC,[DY].dbo.CMSMV WHERE TC006=MV001 AND  TC001+TC002 IN (SELECT TOP 1 TA026+TA027 FROM [DY].dbo.MOCTA WHERE TA001+TA002 IN (SELECT TOP 1 TG014+TG015 FROM [DY].dbo.MOCTG WHERE TG004=LA001 AND TG017=LA016))) AS '業務'
                                     FROM [DY].dbo.INVLA WITH (NOLOCK)  
                                     LEFT JOIN  [DY].dbo.INVMB WITH (NOLOCK) ON MB001=LA001   
                                     WHERE  (LA009='10001')   
                                     AND (LA001 LIKE '4%' OR LA001 LIKE '5%')
                                     GROUP BY  LA001,LA009,MB002,MB003,LA016,MB023,MB198,MB004    
                                     HAVING SUM(LA005*LA011)<>0 
                                     ) AS TEMP
                                     ORDER BY 品號  
                                        ", DateTime.Now.ToString("yyyyMMdd"));
            }
            else if (comboBox2.Text.Equals("11001     大榮觀音倉"))
            {
                FASTSQL.AppendFormat(@" 
                                    SELECT  LA001 AS '品號' ,MB002 AS '品名',MB003 AS '規格',LA016 AS '批號'  ,CAST(SUM(LA005*LA011) AS DECIMAL(18,0)) AS '庫存量',CONVERT(DECIMAL(16,3),SUM(LA005*LA013)) AS '庫存金額'  
                                    ,DATEADD(day,1,DATEADD(month,-1*10,CONVERT(DATETIME,LA016)))
                                    ,CASE WHEN MB198='2' AND MB023>0 THEN DATEADD(day,1,DATEADD(month,-1*MB023,CONVERT(DATETIME,LA016))) END AS '製造日期'
                                    ,MB004 AS '庫存單位'
                                    ,(CAST(SUM(LA005*LA011) /240 AS DECIMAL(18,0))) AS '板數'
                                    ,'' AS '備註'
                                    FROM [DY].dbo.INVLA WITH (NOLOCK) LEFT JOIN  [DY].dbo.INVMB WITH (NOLOCK) ON MB001=LA001 
                                    WHERE  (LA009='11001') 
                                    GROUP BY  LA001,MB002,MB003,LA016,MB198,MB023,MB004
                                    HAVING SUM(LA005*LA011)<>0
                                    ORDER BY  LA001,MB002,MB003,LA016,MB198,MB023,MB004
                                        "
                                        );
            }
            else
            {
                FASTSQL.AppendFormat(@"

                                   SELECT  LA001 AS '品號' ,MB002 AS '品名',MB003 AS '規格',LA016 AS '批號'  ,CAST(SUM(LA005*LA011) AS DECIMAL(18,4)) AS '庫存量',CONVERT(DECIMAL(16,3),SUM(LA005*LA013)) AS '庫存金額'  
                                    FROM [DY].dbo.INVLA WITH (NOLOCK) LEFT JOIN  [DY].dbo.INVMB WITH (NOLOCK) ON MB001=LA001 
                                    WHERE  (LA009='{0}') {1}
                                    GROUP BY  LA001,MB002,MB003,LA016
                                    HAVING SUM(LA005*LA011)<>0
                                    ORDER BY  LA001,MB002,MB003,LA016  "
                                    , comboBox2.SelectedValue.ToString(), sbSqlQuery.ToString());

                 
            }





            return FASTSQL.ToString();
        }
        #endregion

        #region BUTTON
        private void button1_Click(object sender, EventArgs e)
        {
            //Search();
            
            SETFASTREPORT();
        }
        private void button2_Click(object sender, EventArgs e)
        {
            //ExcelExport();
        }
        private void button3_Click(object sender, EventArgs e)
        {
            SETFASTREPORT2();
        }

        private void button4_Click(object sender, EventArgs e)
        {
            SETFASTREPORT3();
        }


        #endregion


    }


}

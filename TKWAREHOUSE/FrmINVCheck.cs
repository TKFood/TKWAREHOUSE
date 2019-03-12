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

        public FrmINVCheck()
        {
            InitializeComponent();
            comboboxload();
        }

        #region FUNCTION
        public void comboboxload()
        {

            connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
            sqlConn = new SqlConnection(connectionString);
            String Sequel = "SELECT MC001,MC001+MC002 AS MC002 FROM CMSMC WITH (NOLOCK) ORDER BY MC001";
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

            if (comboBox1.Text.Equals("20006     原料倉"))
            {
                report1.Load(@"REPORT\每日盤點表.frx");
            }
            else if (comboBox1.Text.Equals("20004     物料倉"))
            {
                report1.Load(@"REPORT\每日盤點表.frx");
            }
            else if (comboBox1.Text.Equals("20005     半成品倉"))
            {
                report1.Load(@"REPORT\每日盤點表.frx");
            }
            else
            {
                report1.Load(@"REPORT\每日盤點表-成品.frx");
            }

               

            report1.Dictionary.Connections[0].ConnectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;

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

            DateTime dt = DateTime.Now;
            dt = dt.AddMonths(-2);

            if (checkBox1.Checked == true)
            {
                sbSqlQuery.AppendFormat(" AND LA001 IN (SELECT LA001 FROM [TK].dbo.INVLA WITH (NOLOCK) WHERE LA004='{0}'   AND LA009='{1}')", dateTimePicker1.Value.ToString("yyyyMMdd"), comboBox1.SelectedValue.ToString());
                sbSqlQuery.AppendFormat("  AND LA004<='{0}'", dateTimePicker1.Value.ToString("yyyyMMdd"));
            }
            else
            {
                sbSqlQuery.Append(" ");
            }

            if (comboBox1.Text.Equals("20006     原料倉"))
            {
                FASTSQL.AppendFormat(@" SELECT  LA001 AS '品號' ,MB002 AS '品名',MB003 AS '規格' ,LA016 AS '批號' ,CAST(SUM(LA005*LA011) AS DECIMAL(18,4)) AS '庫存量'  ");
                FASTSQL.AppendFormat(@"  FROM [{0}].dbo.INVLA WITH (NOLOCK) LEFT JOIN  [{0}].dbo.INVMB WITH (NOLOCK) ON MB001=LA001 ", sqlConn.Database.ToString());
                FASTSQL.AppendFormat(@" WHERE  (LA009='{0}') {1}", comboBox1.SelectedValue.ToString(), sbSqlQuery.ToString());
                FASTSQL.AppendFormat(@" GROUP BY  LA001,MB002,MB003,LA016");
                FASTSQL.AppendFormat(@" HAVING SUM(LA005*LA011)<>0");
                FASTSQL.AppendFormat(@" ORDER BY  LA001,MB002,MB003,LA016");
                FASTSQL.AppendFormat(@" ");

            }
            else if (comboBox1.Text.Equals("20004     物料倉"))
            {
                FASTSQL.AppendFormat(@" SELECT  LA001 AS '品號' ,MB002 AS '品名',MB003 AS '規格',LA016 AS '批號'  ,CAST(SUM(LA005*LA011) AS DECIMAL(18,4)) AS '庫存量'  ");
                FASTSQL.AppendFormat(@"  FROM [{0}].dbo.INVLA WITH (NOLOCK) LEFT JOIN  [{0}].dbo.INVMB WITH (NOLOCK) ON MB001=LA001 ", sqlConn.Database.ToString());
                FASTSQL.AppendFormat(@" WHERE  (LA009='{0}') {1}", comboBox1.SelectedValue.ToString(), sbSqlQuery.ToString());
                FASTSQL.AppendFormat(@" GROUP BY  LA001,MB002,MB003,LA016");
                FASTSQL.AppendFormat(@" HAVING SUM(LA005*LA011)<>0");
                FASTSQL.AppendFormat(@" ORDER BY  LA001,MB002,MB003,LA016");

              
            }
            else if (comboBox1.Text.Equals("20005     半成品倉"))
            {
                FASTSQL.AppendFormat(@" SELECT  LA001 AS '品號' ,MB002 AS '品名',MB003 AS '規格',LA016 AS '批號'  ,CAST(SUM(LA005*LA011) AS DECIMAL(18,4)) AS '庫存量'  ");
                FASTSQL.AppendFormat(@"  FROM [{0}].dbo.INVLA WITH (NOLOCK) LEFT JOIN  [{0}].dbo.INVMB WITH (NOLOCK) ON MB001=LA001 ", sqlConn.Database.ToString());
                FASTSQL.AppendFormat(@" WHERE  (LA009='{0}') {1}", comboBox1.SelectedValue.ToString(), sbSqlQuery.ToString());
                FASTSQL.AppendFormat(@" GROUP BY  LA001,MB002,MB003,LA016");
                FASTSQL.AppendFormat(@" HAVING SUM(LA005*LA011)<>0");
                FASTSQL.AppendFormat(@" ORDER BY  LA001,MB002,MB003,LA016");

         
            }
            else 
            {
                FASTSQL.AppendFormat(@" SELECT  LA001 AS '品號' ,MB002 AS '品名',MB003 AS '規格',LA016 AS '批號',CAST(SUM(LA005*LA011) AS INT) AS '庫存量'    ");
                FASTSQL.AppendFormat(@" ,(SELECT ISNULL(SUM(NUM),0) FROM [TK].dbo.VCOPTDINVMD WHERE TD004=LA001 AND TD013>='{0}') AS '訂單需求量'", dt.ToString("yyyyMMdd"));
                FASTSQL.AppendFormat(@" ,(CAST(SUM(LA005*LA011) AS INT)-(SELECT ISNULL(SUM(NUM),0) FROM [TK].dbo.VCOPTDINVMD WHERE TD004=LA001 AND TD013>='{0}')) AS '需求差異量'", dt.ToString("yyyyMMdd"));
                FASTSQL.AppendFormat(@" ,DATEDIFF(DAY,(SELECT TOP 1 LA004 FROM [TK].dbo.INVLA A WHERE A.LA001=INVLA.LA001 AND A.LA016=INVLA.LA016 AND LA005='1') , '{0}' ) AS '在倉日期'", DateTime.Now.ToString("yyyyMMdd"));
                FASTSQL.AppendFormat(@" ,DATEDIFF(DAY, '{0}',LA016  )  AS '有效天數'",DateTime.Now.ToString("yyyyMMdd"));
                FASTSQL.AppendFormat(@" FROM [TK].dbo.INVLA WITH (NOLOCK) ");
                FASTSQL.AppendFormat(@" LEFT JOIN  [TK].dbo.INVMB WITH (NOLOCK) ON MB001=LA001  ");
                FASTSQL.AppendFormat(@" WHERE  (LA009='{0}') {1}", comboBox1.SelectedValue.ToString(), sbSqlQuery.ToString());
                FASTSQL.AppendFormat(@" AND LA001 LIKE '4%'");
                FASTSQL.AppendFormat(@" GROUP BY  LA001,MB002,MB003,LA016,MB023,MB198 ");
                FASTSQL.AppendFormat(@" HAVING SUM(LA005*LA011)<>0 ");
                FASTSQL.AppendFormat(@" ORDER BY  LA001,MB002,MB003,LA016,MB023,MB198");
                FASTSQL.AppendFormat(@" ");
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

        #endregion




    }


}

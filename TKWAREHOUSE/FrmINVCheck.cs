using System;
using System.Collections.Generic;
using System.ComponentModel;
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
using NPOI.SS.Util;

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

        DataTable dt = new DataTable();

        public FrmINVCheck()
        {
            InitializeComponent();
        }

        #region FUNCTION
        public void Search()
        {
            try
            {
                
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sbSql.Clear();
                sbSqlQuery.Clear();
                   
                sbSql.Append(@" SELECT  LA001 AS '品號' ,MB002 AS '品名',MB003 AS '規格' ,CAST(SUM(LA005*LA011) AS INT) AS '庫存量'  ");
                sbSql.AppendFormat(@"  FROM [{0}].dbo.INVLA WITH (NOLOCK) LEFT JOIN  [{0}].dbo.INVMB WITH (NOLOCK) ON MB001=LA001 ", sqlConn.Database.ToString());
                sbSql.Append(@" WHERE LA001 LIKE '4%' AND (LA009='20001')");
                sbSql.Append(@" GROUP BY  LA001,MB002,MB003");
                sbSql.Append(@" HAVING SUM(LA005*LA011)>=1");
                sbSql.Append(@" ORDER BY  LA001,MB002,MB003");


                adapter = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder = new SqlCommandBuilder(adapter);
                sqlConn.Open();
                ds.Clear();
                adapter.Fill(ds, "TEMPds");
                sqlConn.Close();


                if (ds.Tables["TEMPds"].Rows.Count == 0)
                {
                    label14.Text = "找不到資料";
                }
                else
                {
                    label14.Text = "有 " + ds.Tables["TEMPds"].Rows.Count.ToString() + " 筆";

                    dataGridView1.DataSource = ds.Tables["TEMPds"];
                    dataGridView1.AutoResizeColumns();
                }

            }
            catch
            {

            }
            finally
            {

            }
        }

        public void ExcelExport()
        {

            string NowDB = "TK";
            //建立Excel 2003檔案
            IWorkbook wb = new XSSFWorkbook();
            ISheet ws;

            XSSFCellStyle cs = (XSSFCellStyle)wb.CreateCellStyle();
            //框線樣式及顏色
            cs.BorderBottom = NPOI.SS.UserModel.BorderStyle.Double;
            cs.BorderLeft = NPOI.SS.UserModel.BorderStyle.Thin;
            cs.BorderRight = NPOI.SS.UserModel.BorderStyle.Thin;
            cs.BorderTop = NPOI.SS.UserModel.BorderStyle.Thin;
            cs.BottomBorderColor = NPOI.HSSF.Util.HSSFColor.Grey50Percent.Index;
            cs.LeftBorderColor = NPOI.HSSF.Util.HSSFColor.Grey50Percent.Index;
            cs.RightBorderColor = NPOI.HSSF.Util.HSSFColor.Grey50Percent.Index;
            cs.TopBorderColor = NPOI.HSSF.Util.HSSFColor.Grey50Percent.Index;

            Search();
            dt = ds.Tables["TEMPds"];

            if (dt.TableName != string.Empty)
            {
                ws = wb.CreateSheet(dt.TableName);
            }
            else
            {
                ws = wb.CreateSheet("Sheet1");
            }

            ws.CreateRow(0);
            //第一行為表名稱
            ws.GetRow(0).CreateCell(0).SetCellValue("老楊食品大林廠20001-庫存表       年     月     日");
            ws.AddMergedRegion(new CellRangeAddress(0, 0, 0, 5));
            //第一行為欄位名稱
            ws.CreateRow(1);
            ws.GetRow(1).CreateCell(0).SetCellValue("品名");
            ws.GetRow(1).CreateCell(1).SetCellValue("規格");
            ws.GetRow(1).CreateCell(2).SetCellValue("數量");
            ws.GetRow(1).CreateCell(3).SetCellValue("品名");
            ws.GetRow(1).CreateCell(4).SetCellValue("規格");
            ws.GetRow(1).CreateCell(5).SetCellValue("數量");
            ws.GetRow(1).CreateCell(6).SetCellValue("品名");
            ws.GetRow(1).CreateCell(7).SetCellValue("規格");
            ws.GetRow(1).CreateCell(8).SetCellValue("數量");


            //for (int i = 1; i < dt.Rows.Count; i++)
            //{
            //    ws.CreateRow(i + 1);
            //    for (int j = 0; j < dt.Columns.Count; j++)
            //    {
            //        ws.GetRow(i + 1).CreateCell(j).SetCellValue(dt.Rows[i][j].ToString());
            //    }
            //}


            int j = 1;
            int k = 0;
            if (dt.Rows.Count <= 40)
            {
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    ws.CreateRow(j + 1);
                    ws.GetRow(j + 1).CreateCell(k).SetCellValue(dt.Rows[i][1].ToString());
                    ws.GetRow(j + 1).CreateCell(k + 1).SetCellValue(dt.Rows[i][2].ToString());
                    ws.GetRow(j + 1).CreateCell(k + 2).SetCellValue(Convert.ToInt16(dt.Rows[i][3].ToString()));

                    j++;
                }

            }
            else if (dt.Rows.Count <= 80 && dt.Rows.Count >= 41)
            {
                for (int i = 0; i <= 40; i++)
                {
                    ws.CreateRow(j + 1);
                    ws.GetRow(j + 1).CreateCell(k).SetCellValue(dt.Rows[i][1].ToString());
                    ws.GetRow(j + 1).CreateCell(k + 1).SetCellValue(dt.Rows[i][2].ToString());
                    ws.GetRow(j + 1).CreateCell(k + 2).SetCellValue(Convert.ToInt32(dt.Rows[i][3].ToString()));
                    if ((i + 41) < dt.Rows.Count)
                    {
                        ws.GetRow(j + 1).CreateCell(k + 3).SetCellValue(dt.Rows[i + 41][1].ToString());
                        ws.GetRow(j + 1).CreateCell(k + 4).SetCellValue(dt.Rows[i + 41][2].ToString());
                        ws.GetRow(j + 1).CreateCell(k + 5).SetCellValue(Convert.ToInt32(dt.Rows[i + 41][3].ToString()));
                    }


                    j++;
                }
            }

            else if (dt.Rows.Count <= 120 && dt.Rows.Count >= 81)
            {
                for (int i = 0; i <= 40; i++)
                {
                    ws.CreateRow(j + 1);
                    ws.GetRow(j + 1).CreateCell(k).SetCellValue(dt.Rows[i][1].ToString());
                    ws.GetRow(j + 1).CreateCell(k+1).SetCellValue(dt.Rows[i][2].ToString());
                    ws.GetRow(j + 1).CreateCell(k + 2).SetCellValue(Convert.ToInt32(dt.Rows[i][3].ToString()));
                    ws.GetRow(j + 1).CreateCell(k + 3).SetCellValue(dt.Rows[i + 41][1].ToString());
                    ws.GetRow(j + 1).CreateCell(k+4).SetCellValue(dt.Rows[i + 41][2].ToString());
                    ws.GetRow(j + 1).CreateCell(k + 5).SetCellValue(Convert.ToInt32(dt.Rows[i + 41][3].ToString()));
                    if ((i + 81) < dt.Rows.Count)
                    {
                        ws.GetRow(j + 1).CreateCell(k + 6).SetCellValue(dt.Rows[i + 81][1].ToString());
                        ws.GetRow(j + 1).CreateCell(k + 7).SetCellValue(dt.Rows[i + 81][2].ToString());
                        ws.GetRow(j + 1).CreateCell(k + 8).SetCellValue(Convert.ToInt32(dt.Rows[i + 81][3].ToString()));
                    }
                    j++;
                }
            }
            else
            {
                foreach (DataGridViewRow dr in this.dataGridView1.Rows)
                {
                    ws.CreateRow(j + 1);
                    ws.GetRow(j + 1).CreateCell(k).SetCellValue(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[1].ToString());
                    ws.GetRow(j + 1).CreateCell(k+1).SetCellValue(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[2].ToString());
                    ws.GetRow(j + 1).CreateCell(k + 2).SetCellValue(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[3].ToString());
                   
                    j++;
                }
            }

            


            //int j = 1;
            //int k = 0;
            //foreach (DataGridViewRow dr in this.dataGridView1.Rows)
            //{
            //    ws.CreateRow(j + 1);
            //    ws.GetRow(j + 1).CreateCell(k).SetCellValue(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[1].ToString());
            //    ws.GetRow(j + 1).CreateCell(k + 1).SetCellValue(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[2].ToString());
            //    ws.GetRow(j + 1).CreateCell(k + 2).SetCellValue(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[1].ToString());
            
            //    j++;
            //}

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
            filename.AppendFormat(@"c:\temp\老楊食品大林廠20001-庫存表{0}.xlsx", DateTime.Now.ToString("yyyyMMdd"));

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
        #endregion

        #region BUTTON
        private void button1_Click(object sender, EventArgs e)
        {
            Search();
        }
        #endregion

        private void button2_Click(object sender, EventArgs e)
        {
            ExcelExport();
        }
    }


}

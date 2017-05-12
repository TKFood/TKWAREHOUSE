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

namespace TKWAREHOUSE
{
    public partial class FrmREPORT : Form
    {
        SqlConnection sqlConn = new SqlConnection();
        SqlCommand sqlComm = new SqlCommand();
        string connectionString;
        StringBuilder sbSql = new StringBuilder();
        StringBuilder sbSqlQuery = new StringBuilder();
        SqlDataAdapter adapter1 = new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilder1 = new SqlCommandBuilder();
        SqlCommand cmd = new SqlCommand();
        DataSet ds1 = new DataSet();
        
        DataTable dt = new DataTable();
        string tablename = null;


        int rownum = 0;
        public FrmREPORT()
        {
            InitializeComponent();
        }

        #region FUNCTION
        public void Search()
        {
            try
            {
                sbSql.Clear();
                sbSql = SETsbSql();

                if (!string.IsNullOrEmpty(sbSql.ToString()))
                {
                    connectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
                    sqlConn = new SqlConnection(connectionString);

                    adapter1 = new SqlDataAdapter(sbSql.ToString(), sqlConn);
                    sqlCmdBuilder1 = new SqlCommandBuilder(adapter1);

                    sqlConn.Open();
                    //dataGridView1.Columns.Clear();
                    ds1.Clear();

                    adapter1.Fill(ds1, tablename);
                    sqlConn.Close();

                    if (ds1.Tables[tablename].Rows.Count == 0)
                    {

                    }
                    else
                    {

                        dataGridView1.DataSource = ds1.Tables[tablename];
                        dataGridView1.AutoResizeColumns();
                        //dataGridView1.CurrentCell = dataGridView1.Rows[rownum].Cells[0];

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
            DateTime dt = dateTimePicker1.Value;
            string ThisYear = null;
            string ThisMonth = null;
            string LastMonth = null;
            string LastYear = null;
            string LastYearMonth = null;

            ThisYear = dateTimePicker1.Value.ToString("yyyy");
            ThisMonth = dateTimePicker1.Value.ToString("MM");
            LastMonth = dt.AddMonths(-1).ToString("MM");
            LastYear = dt.AddYears(-1).ToString("yyyy");
            LastYearMonth = dt.AddYears(-1).AddMonths(1).ToString("MM");

            if (comboBox1.Text.ToString().Equals("入庫單統計報表"))
            {
                STR.AppendFormat(@"  SELECT TF003 AS '日期',TG001 AS '入庫單別',TG002 AS '入庫單號',TG004 AS '品號',TG005 AS '品名',TG006 AS '規格',TG007 AS '單位',");
                STR.AppendFormat(@"  TG010 AS '庫別',MC002 AS '庫名',TG011 AS '入庫量',TG014 AS '製令單別',TG015 AS '製令單號'");
                STR.AppendFormat(@"  FROM [TK].dbo.MOCTF,[TK].dbo.MOCTG,[TK].dbo.CMSMC");
                STR.AppendFormat(@"  WHERE TF001=TG001 AND TF002=TG002");
                STR.AppendFormat(@"  AND TG010=MC001");
                STR.AppendFormat(@"  AND TF003>='{0}' AND TF003<='{1}'",dateTimePicker1.Value.ToString("yyyyMMdd"), dateTimePicker2.Value.ToString("yyyyMMdd"));
                STR.AppendFormat(@"  ORDER BY TG001,TG002,TG005");
                STR.AppendFormat(@"  ");
                STR.AppendFormat(@"  ");


                tablename = "TEMPds1";
            }            
            else if (comboBox1.Text.ToString().Equals(""))
            {

            }


            return STR;
        }

        public void ExcelExport()
        {
            Search();
            string TABLENAME = "報表";

            //建立Excel 2003檔案
            IWorkbook wb = new XSSFWorkbook();
            ISheet ws;


            dt = ds1.Tables[tablename];
            if (dt.TableName != string.Empty)
            {
                ws = wb.CreateSheet(dt.TableName);
            }
            else
            {
                ws = wb.CreateSheet("Sheet1");
            }

            ws.CreateRow(0);//第一行為欄位名稱
            for (int i = 0; i < dt.Columns.Count; i++)
            {
                ws.GetRow(0).CreateCell(i).SetCellValue(dt.Columns[i].ColumnName);
            }


            int j = 0;
            if (tablename.Equals("TEMPds1"))
            {
                TABLENAME = "入庫單統計報表";
                foreach (DataGridViewRow dr in this.dataGridView1.Rows)
                {
                    ws.CreateRow(j + 1);
                    ws.GetRow(j + 1).CreateCell(0).SetCellValue(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[0].ToString());
                    ws.GetRow(j + 1).CreateCell(1).SetCellValue(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[1].ToString());
                    ws.GetRow(j + 1).CreateCell(2).SetCellValue(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[2].ToString());
                    ws.GetRow(j + 1).CreateCell(3).SetCellValue(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[3].ToString());
                    ws.GetRow(j + 1).CreateCell(4).SetCellValue(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[4].ToString());
                    ws.GetRow(j + 1).CreateCell(5).SetCellValue(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[5].ToString());
                    ws.GetRow(j + 1).CreateCell(6).SetCellValue(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[6].ToString());
                    ws.GetRow(j + 1).CreateCell(7).SetCellValue(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[7].ToString());
                    ws.GetRow(j + 1).CreateCell(8).SetCellValue(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[8].ToString());
                    ws.GetRow(j + 1).CreateCell(9).SetCellValue(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[9].ToString());
                    ws.GetRow(j + 1).CreateCell(10).SetCellValue(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[10].ToString());
                    ws.GetRow(j + 1).CreateCell(11).SetCellValue(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[11].ToString());

                    j++;
                }

            }
            
            else if (tablename.Equals(""))
            {

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
            filename.AppendFormat(@"c:\temp\{0}-{1}.xlsx", TABLENAME, DateTime.Now.ToString("yyyyMMdd"));

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

        private void button2_Click(object sender, EventArgs e)
        {
            ExcelExport();
        }

        #endregion


    }
}

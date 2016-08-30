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
using NPOI.SS.UserModel;
using System.Configuration;
using NPOI.XSSF.UserModel;
using NPOI.SS.Util;

namespace TKWAREHOUSE
{
    public partial class FrmINVCOP : Form
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
        DataTable dtTemp = new DataTable();
        DataColumn column1 = new DataColumn("MD001");
        DataColumn column2 = new DataColumn("MD003");
        DataColumn column3 = new DataColumn("NUM");
        DataColumn column4 = new DataColumn("UNIT");
        string tablename = null;
        decimal COPNum = 0;
        double BOMNum = 0;

        public FrmINVCOP()
        {
            InitializeComponent();
            dateTimePicker1.Value = DateTime.Now;

            dtTemp.Columns.Add(column1);
            dtTemp.Columns.Add(column2);
            dtTemp.Columns.Add(column3);
            dtTemp.Columns.Add(column4);

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

                sbSql.Append(@"  SELECT TD004,TD010,MD002,MD004,SUM(CASE WHEN ISNULL(MD004,0)<>0 THEN (TD008+TD024)*MD004 ELSE TD008 END )AS NUM,MC001,MC002,MC004,SUM(CASE WHEN ISNULL(MD004,0)<>0 THEN (TD008+TD024)*MD004 ELSE TD008 END )/MC004 AS BOMNum");
                sbSql.Append(@"  FROM [TK].dbo.COPTD");
                sbSql.Append(@"  LEFT JOIN [TK].dbo.INVMD ON TD004=MD001  AND MD002=TD010");
                sbSql.Append(@"  LEFT JOIN [TK].dbo.BOMMC ON TD004=MC001");
                sbSql.AppendFormat(@"  WHERE SUBSTRING(TD002,1,8)='{0}' AND TD008>0  AND TD002='20160830001'   ", dateTimePicker1.Value.ToString("yyyyMMdd"));
                sbSql.Append(@"  GROUP BY TD004,TD010,MD002,MD004,MC001,MC002,MC004");
                sbSql.Append(@"  ");
               
                adapter = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder = new SqlCommandBuilder(adapter);
                sqlConn.Open();
                ds.Clear();
                adapter.Fill(ds, "TEMPds1");
                sqlConn.Close();


                if (ds.Tables["TEMPds1"].Rows.Count == 0)
                {
                    label14.Text = "找不到資料";
                }
                else
                {
                    if (ds.Tables["TEMPds1"].Rows.Count >= 1)
                    {

                        for (int i = 0; i < ds.Tables["TEMPds1"].Rows.Count; i++)
                        {

                            COPNum = Convert.ToDecimal(ds.Tables["TEMPds1"].Rows[i]["NUM"].ToString());
                            BOMNum = Convert.ToDouble(ds.Tables["TEMPds1"].Rows[i]["BOMNum"].ToString());

                            sbSql.Clear();
                            sbSqlQuery.Clear();

                            sbSql.Append(@"  WITH TreeNode (MD001,MD002,MD003,MD004,MD006,MD007, Level)");
                            sbSql.Append(@"  AS");
                            sbSql.Append(@"  (");
                            sbSql.Append(@"  SELECT MD001,MD002,MD003,MD004,MD006,MD007, 0 AS Level");
                            sbSql.Append(@"  FROM [TK].dbo.BOMMD");
                            sbSql.AppendFormat(@"  WHERE MD001='{0}'", ds.Tables["TEMPds1"].Rows[i]["TD004"].ToString());
                            sbSql.Append(@"  UNION ALL");
                            sbSql.Append(@"  SELECT ta.MD001,ta.MD002,ta.MD003,ta.MD004,ta.MD006,ta.MD007 ,Level + 1");
                            sbSql.Append(@"  FROM [TK].dbo.BOMMD ta");
                            sbSql.Append(@"  INNER JOIN TreeNode AS tn");
                            sbSql.Append(@"  ON ta.MD001 = tn.MD003");
                            sbSql.Append(@"  )");
                            sbSql.Append(@"  SELECT MD001,MD002,MD003,MD004,MD006,MD007, Level,MB002,MB003 FROM TreeNode,[TK].dbo.INVMB");
                            sbSql.Append(@"  WHERE MD001=MB001");

                            adapter = new SqlDataAdapter(@"" + sbSql, sqlConn);

                            sqlCmdBuilder = new SqlCommandBuilder(adapter);
                            sqlConn.Open();
                            ds2.Clear();
                            adapter.Fill(ds2, "TEMPds2");
                            sqlConn.Close();

                            if (ds2.Tables["TEMPds2"].Rows.Count > 1)
                            {

                                foreach (DataRow od2 in ds2.Tables["TEMPds2"].Rows)
                                {
                                    DataRow row = dtTemp.NewRow();
                                    row["MD001"] = od2["MD001"].ToString();
                                    row["MD003"] = od2["MD003"].ToString();
                                    row["NUM"] = Convert.ToDouble(od2["MD006"].ToString()) * BOMNum;
                                    row["UNIT"] = od2["MD004"].ToString();
                                    dtTemp.Rows.Add(row);
                                }

                            }

                        }
                        
                    }

                    //dtTemp = ds.Tables["TEMPds1"];
                    //dtTemp = ds2.Tables["TEMPds2"];
                    label14.Text = "有 " + dtTemp.Rows.Count.ToString() + " 筆";

                    dataGridView1.DataSource = dtTemp;
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
            dt = ds.Tables[tablename];

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
                foreach (DataGridViewRow dr in this.dataGridView1.Rows)
                {
                    ws.CreateRow(j + 1);
                    ws.GetRow(j + 1).CreateCell(0).SetCellValue(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[0].ToString());
                    ws.GetRow(j + 1).CreateCell(1).SetCellValue(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[1].ToString());
                    ws.GetRow(j + 1).CreateCell(2).SetCellValue(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[2].ToString());
                    ws.GetRow(j + 1).CreateCell(3).SetCellValue(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[3].ToString());
                    ws.GetRow(j + 1).CreateCell(4).SetCellValue(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[4].ToString());
                    ws.GetRow(j + 1).CreateCell(5).SetCellValue(Convert.ToDouble(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[5].ToString()));
                    ws.GetRow(j + 1).CreateCell(6).SetCellValue(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[6].ToString());
                    ws.GetRow(j + 1).CreateCell(7).SetCellValue(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[7].ToString());

                    j++;
                }
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
            filename.AppendFormat(@"c:\temp\庫存呆滯表{0}.xlsx", DateTime.Now.ToString("yyyyMMdd"));

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

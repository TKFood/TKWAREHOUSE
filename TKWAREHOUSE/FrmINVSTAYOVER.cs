using System;
using System.Data;
using System.Text;
using System.Windows.Forms;
using System.IO;
using System.Data.SqlClient;
using NPOI.SS.UserModel;
using System.Configuration;
using NPOI.XSSF.UserModel;
using NPOI.SS.Util;

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

        DataTable dt = new DataTable();
        string tablename = null;
        DateTime StayDay;

        public FrmINVSTAYOVER()
        {
            InitializeComponent();
            comboboxload();
            dateTimePicker1.Value = DateTime.Now;
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
        public void Search()
        {
            StayDay = dateTimePicker1.Value;
            StayDay = StayDay.AddDays(-1*Convert.ToDouble(numericUpDown1.Value));
            try
            {


                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

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
                    dataGridView1.Rows.Clear();
                    dataGridView1.DataSource = null;
                }
                else
                {
                    label14.Text = "有 " + ds.Tables[tablename].Rows.Count.ToString() + " 筆";
                    dataGridView1.Rows.Clear();
                    dataGridView1.DataSource = null;
                    dataGridView1.DataSource = ds.Tables[tablename];
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
                    ws.GetRow(j + 1).CreateCell(5).SetCellValue(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[5].ToString());
                    ws.GetRow(j + 1).CreateCell(6).SetCellValue(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[6].ToString());
                    ws.GetRow(j + 1).CreateCell(7).SetCellValue(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[7].ToString());
                    ws.GetRow(j + 1).CreateCell(8).SetCellValue(Convert.ToDouble(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[8].ToString()));
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

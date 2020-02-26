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
using System.Configuration;
using System.Reflection;
using System.Threading;
using System.Globalization;
using Calendar.NET;
using FastReport;
using FastReport.Data;
using Excel = Microsoft.Office.Interop.Excel;

namespace TKWAREHOUSE
{
    public partial class FrmCALENDERPUR : Form
    {
        private ComponentResourceManager _ResourceManager = new ComponentResourceManager();
        SqlConnection sqlConn = new SqlConnection();
        SqlCommand sqlComm = new SqlCommand();
        string connectionString;
        StringBuilder sbSql = new StringBuilder();
        StringBuilder sbSqlQuery = new StringBuilder();
        SqlDataAdapter adapter1 = new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilder1 = new SqlCommandBuilder();
        SqlTransaction tran;
        SqlCommand cmd = new SqlCommand();
        SqlDataAdapter adapterCALENDAR = new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilderCALENDAR = new SqlCommandBuilder();
        DataSet dsCALENDAR = new DataSet();

        SqlDataAdapter adapter2= new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilder2= new SqlCommandBuilder();
        DataSet ds2 = new DataSet();

        int result;
        public Report report1 { get; private set; }
        string[] message1 = new string[31] { "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "" };
        string DATES = null;
        string strDesktopPath;
        string pathFile1;
        DateTime sdt;
        DateTime edt;
        DateTime sdt2;
        DateTime edt2;

        public FrmCALENDERPUR()
        {
            InitializeComponent();

            SETCALENDAR();
            SETDATE();
        }

        #region FUNCTION

        public void SETDATE()
        {
            DateTime SETDT = Convert.ToDateTime(dateTimePicker2.Value.ToString("yyyy/MM") + "/01");
            DateTime FirstDay = SETDT.AddDays(-SETDT.Day + 1);
            DateTime LastDay = SETDT.AddMonths(1).AddDays(-SETDT.AddMonths(1).Day);

            sdt = FirstDay;
            edt = LastDay;
        }
        public void SETCALENDAR()
        {
            string EVENT;
            DateTime dtEVENT;

            DateTime STARTTIME = DateTime.Now;
            STARTTIME = STARTTIME.AddYears(-1);

            var ce2 = new CustomEvent();


            calendar1.RemoveAllEvents();
            //calendar1.CalendarDate = new DateTime(DateTime.Now.Year, DateTime.Now.Month, DateTime.Now.Day, 0, 0, 0);
            calendar1.CalendarDate = new DateTime(dateTimePicker1.Value.Year, dateTimePicker1.Value.Month, DateTime.Now.Day, 0, 0, 0);
            calendar1.CalendarView = CalendarViews.Month;
            //calendar1.CalendarView = CalendarViews.Day;
            calendar1.AllowEditingEvents = false;




            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sbSql.Clear();
                sbSqlQuery.Clear();

                sbSql.AppendFormat(@"  SELECT TD012,CONVERT(varchar,TD004)+'-'+CONVERT(varchar,TD005)+'-採購量'+CONVERT(varchar,TD008)+''+CONVERT(varchar,TD009) AS 'EVENT'");
                sbSql.AppendFormat(@"  FROM [TK].dbo.PURTC,[TK].dbo.PURTD");
                sbSql.AppendFormat(@"  WHERE TD001=TD001 AND TC002=TD002");
                sbSql.AppendFormat(@"  AND TD012 LIKE '{0}%'",dateTimePicker1.Value.ToString("yyyyMM"));
                sbSql.AppendFormat(@"  ORDER BY TD012,TD004");
                sbSql.AppendFormat(@"  ");

                adapterCALENDAR = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilderCALENDAR = new SqlCommandBuilder(adapter1);
                sqlConn.Open();
                dsCALENDAR.Clear();
                adapterCALENDAR.Fill(dsCALENDAR, "TEMPdsCALENDAR");
                sqlConn.Close();


                if (dsCALENDAR.Tables["TEMPdsCALENDAR"].Rows.Count == 0)
                {

                }
                else
                {
                    if (dsCALENDAR.Tables["TEMPdsCALENDAR"].Rows.Count >= 1)
                    {
                        foreach (DataRow od in dsCALENDAR.Tables["TEMPdsCALENDAR"].Rows)
                        {
                            EVENT = od["EVENT"].ToString();
                            dtEVENT = Convert.ToDateTime(od["TD012"].ToString().Substring(0,4)+"/"+ od["TD012"].ToString().Substring(4, 2)+"/"+ od["TD012"].ToString().Substring(6, 2));

                            ce2 = new CustomEvent
                            {
                                IgnoreTimeComponent = false,
                                EventText = EVENT,
                                Date = new DateTime(dtEVENT.Year, dtEVENT.Month, dtEVENT.Day),
                                EventLengthInHours = 2f,
                                RecurringFrequency = RecurringFrequencies.None,
                                EventFont = new Font("Verdana", 8, FontStyle.Regular),
                                Enabled = true,
                                EventColor = Color.FromArgb(120, 255, 120),
                                EventTextColor = Color.Black,
                                ThisDayForwardOnly = true
                            };

                            calendar1.AddEvent(ce2);

                            
                            //calendar1.HighlightCurrentDay = true;
                            //calendar1.AutoSize = true;
                            //calendar1.AutoScroll = true;

                        }



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

        public void SETFASTREPORT()
        {

            string SQL;
            string SQL1;
            report1 = new Report();
            report1.Load(@"REPORT\預計採購表.frx");

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

            if(string.IsNullOrEmpty(textBox1.Text))
            {
                FASTSQL.AppendFormat(@"  SELECT TD012 AS '預交日',MA002 AS '廠商',TD004 AS '品號', TD005 AS '品名',TD006 AS '規格',TD008 AS '採購量',TD015 AS '已交量',TD009 AS '單位',TD012 AS '預交日',TD014 ");
                FASTSQL.AppendFormat(@"  ,(SELECT TOP 1 TD014 FROM [TK].dbo.PURTD A WHERE A.TD001=PURTD.TD001 AND A.TD002=PURTD.TD002 AND A.TD003='0001') AS COMMENT1");
                FASTSQL.AppendFormat(@"  ,(CASE WHEN ISNULL(TD014,'')<>'' THEN TD014 ELSE (SELECT TOP 1 TD014 FROM [TK].dbo.PURTD A WHERE A.TD001=PURTD.TD001 AND A.TD002=PURTD.TD002 AND A.TD003='0001') END )AS '備註'");
                FASTSQL.AppendFormat(@"  ,TD001 AS '採購單別',TD002 AS '採購單號',TD003 AS '採購序號'");
                FASTSQL.AppendFormat(@"  ,TD026 AS '請購單別',TD027 AS '請購單號',TD028 AS '請購序號'");
                FASTSQL.AppendFormat(@"  FROM [TK].dbo.PURTC,[TK].dbo.PURTD,[TK].dbo.PURMA ");
                FASTSQL.AppendFormat(@"  WHERE TC001=TD001 AND TC002=TD002");
                FASTSQL.AppendFormat(@"  AND MA001=TC004");
                FASTSQL.AppendFormat(@"  AND TD012>='20200226' AND TD012<='2020031'");
                FASTSQL.AppendFormat(@"  AND TD018='Y'");
                FASTSQL.AppendFormat(@"  ORDER BY TD012,TD001,TD002,TD003");
                FASTSQL.AppendFormat(@"  ");
                FASTSQL.AppendFormat(@"  ");
                FASTSQL.AppendFormat(@"  ");
            }
            else
            {
                FASTSQL.AppendFormat(@"  SELECT TD012 AS '預交日',MA002 AS '廠商',TD004 AS '品號', TD005 AS '品名',TD006 AS '規格',TD008 AS '採購量',TD015 AS '已交量',TD009 AS '單位',TD012 AS '預交日',TD014 ");
                FASTSQL.AppendFormat(@"  ,(SELECT TOP 1 TD014 FROM [TK].dbo.PURTD A WHERE A.TD001=PURTD.TD001 AND A.TD002=PURTD.TD002 AND A.TD003='0001') AS COMMENT1");
                FASTSQL.AppendFormat(@"  ,(CASE WHEN ISNULL(TD014,'')<>'' THEN TD014 ELSE (SELECT TOP 1 TD014 FROM [TK].dbo.PURTD A WHERE A.TD001=PURTD.TD001 AND A.TD002=PURTD.TD002 AND A.TD003='0001') END )AS '備註'");
                FASTSQL.AppendFormat(@"  ,TD001 AS '採購單別',TD002 AS '採購單號',TD003 AS '採購序號'");
                FASTSQL.AppendFormat(@"  ,TD026 AS '請購單別',TD027 AS '請購單號',TD028 AS '請購序號'");
                FASTSQL.AppendFormat(@"  FROM [TK].dbo.PURTC,[TK].dbo.PURTD,[TK].dbo.PURMA ");
                FASTSQL.AppendFormat(@"  WHERE TC001=TD001 AND TC002=TD002");
                FASTSQL.AppendFormat(@"  AND MA001=TC004");
                FASTSQL.AppendFormat(@"  AND TD012>='20200226' AND TD012<='2020031'");
                FASTSQL.AppendFormat(@"  AND TD018='Y'");
                FASTSQL.AppendFormat(@"  AND TD005 LIKE '%{0}%'",textBox1.Text.Trim());
                FASTSQL.AppendFormat(@"  ORDER BY TD012,TD001,TD002,TD003");
                FASTSQL.AppendFormat(@"  ");
                FASTSQL.AppendFormat(@"  ");
            }
            FASTSQL.AppendFormat(@"   ");
            FASTSQL.AppendFormat(@"  ");

            return FASTSQL.ToString();
        }

        public void RESET1()
        {
            message1 = new string[31] { "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "" };

        }
        public void SETPATH1()
        {
            DATES = DateTime.Now.ToString("yyyyMMddHHmmss");
            strDesktopPath = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory);
            pathFile1 = @"" + strDesktopPath.ToString() + @"\" + "行事曆-採購" + DATES.ToString();


            DeleteDir(pathFile1 + ".xlsx");
        }

        public void DeleteDir(string aimPath)
        {
            try
            {
                File.Delete(aimPath);
            }
            catch (Exception e)
            {
                throw e;
            }
        }

        public void SETFILE1()
        {


            // 設定儲存檔名，不用設定副檔名，系統自動判斷 excel 版本，產生 .xls 或 .xlsx 副檔名 
            Excel.Application excelApp;
            Excel._Workbook wBook;
            Excel._Worksheet wSheet;
            Excel.Range wRange;

            // 開啟一個新的應用程式
            excelApp = new Excel.Application();
            // 讓Excel文件可見
            //excelApp.Visible = true;
            // 停用警告訊息
            excelApp.DisplayAlerts = false;
            // 加入新的活頁簿
            excelApp.Workbooks.Add(Type.Missing);
            // 引用第一個活頁簿
            wBook = excelApp.Workbooks[1];
            // 設定活頁簿焦點
            wBook.Activate();

            if (!File.Exists(pathFile1 + ".xlsx"))
            {
                wBook.SaveAs(pathFile1, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            }



            //關閉Excel
            excelApp.Quit();

            //釋放Excel資源
            System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);
            wBook = null;
            wSheet = null;
            wRange = null;
            excelApp = null;
            GC.Collect();

            Console.Read();


            SEARCH1();

            //if (!File.Exists(pathFile + ".xlsx"))
            //{
            //    //SEARCH()

            //}

        }
        public void SEARCH1()
        {

            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sbSql.Clear();
                sbSqlQuery.Clear();

                StringBuilder SB = new StringBuilder();

               

                sbSql.AppendFormat(@"  ");

                adapter2 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder2 = new SqlCommandBuilder(adapter2);
                sqlConn.Open();
                ds2.Clear();
                adapter2.Fill(ds2, "ds2");
                sqlConn.Close();


                if (ds2.Tables["ds2"].Rows.Count == 0)
                {
                    //建立一筆新的DataRow，並且等於新的dt row
                    DataRow row = ds2.Tables["ds2"].NewRow();

                    //指定每個欄位要儲存的資料                   
                    row[0] = "本日無資料"; ;

                    //新增資料至DataTable的dt內
                    ds2.Tables["ds2"].Rows.Add(row);

                    //ExportDataSetToExcel2(ds5, pathFile2);
                }
                else
                {
                    if (ds2.Tables["ds2"].Rows.Count >= 1)
                    {
                        ExportDataSetToExcel1(ds2, pathFile1);
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

        public void ExportDataSetToExcel1(DataSet ds, string TopathFile)
        {
            int days = Convert.ToInt32(sdt.AddDays(-sdt.Day + 1).DayOfWeek.ToString("d"));
            //MessageBox.Show(days.ToString());
            int MONTHDAYS = DateTime.DaysInMonth(sdt.Year, sdt.Month);

            int EXCELX = 2;
            int EXCELY = 0;

            //Creae an Excel application instance
            Excel.Application excelApp = new Excel.Application();

            //Create an Excel workbook instance and open it from the predefined location
            Excel.Workbook excelWorkBook = excelApp.Workbooks.Open(TopathFile);
            Excel.Range wRange;
            Excel.Range wRangepathFile;
            Excel.Range wRangepathFilePURTA;



            foreach (DataTable table in ds.Tables)
            {
                //Add a new worksheet to workbook with the Datatable name
                Excel.Worksheet excelWorkSheet = excelWorkBook.Sheets.Add();
                excelWorkSheet.Name = table.TableName;

                for (int j = 0; j < table.Rows.Count; j++)
                {
                    for (int k = 0; k < table.Columns.Count; k++)
                    {
                        //if (table.Rows[j].ItemArray[0].ToString().Substring(6,2).Equals("01"))
                        if (Convert.ToInt32(table.Rows[j].ItemArray[0].ToString().Substring(6, 2)) == 1)
                        {
                            message1[0] = message1[0] + table.Rows[j].ItemArray[k].ToString();
                            message1[0] = message1[0] + '\n';
                        }
                        else if (Convert.ToInt32(table.Rows[j].ItemArray[0].ToString().Substring(6, 2)) == 2)
                        {
                            message1[1] = message1[1] + table.Rows[j].ItemArray[k].ToString();
                            message1[1] = message1[1] + '\n';
                        }
                        else if (Convert.ToInt32(table.Rows[j].ItemArray[0].ToString().Substring(6, 2)) == 3)
                        {
                            message1[2] = message1[2] + table.Rows[j].ItemArray[k].ToString();
                            message1[2] = message1[2] + '\n';
                        }
                        else if (Convert.ToInt32(table.Rows[j].ItemArray[0].ToString().Substring(6, 2)) == 4)
                        {
                            message1[3] = message1[3] + table.Rows[j].ItemArray[k].ToString();
                            message1[3] = message1[3] + '\n';
                        }
                        else if (Convert.ToInt32(table.Rows[j].ItemArray[0].ToString().Substring(6, 2)) == 5)
                        {
                            message1[4] = message1[4] + table.Rows[j].ItemArray[k].ToString();
                            message1[4] = message1[4] + '\n';
                        }
                        else if (Convert.ToInt32(table.Rows[j].ItemArray[0].ToString().Substring(6, 2)) == 6)
                        {
                            message1[5] = message1[5] + table.Rows[j].ItemArray[k].ToString();
                            message1[5] = message1[5] + '\n';
                        }
                        else if (Convert.ToInt32(table.Rows[j].ItemArray[0].ToString().Substring(6, 2)) == 7)
                        {
                            message1[6] = message1[6] + table.Rows[j].ItemArray[k].ToString();
                            message1[6] = message1[6] + '\n';
                        }
                        else if (Convert.ToInt32(table.Rows[j].ItemArray[0].ToString().Substring(6, 2)) == 8)
                        {
                            message1[7] = message1[7] + table.Rows[j].ItemArray[k].ToString();
                            message1[7] = message1[7] + '\n';
                        }
                        else if (Convert.ToInt32(table.Rows[j].ItemArray[0].ToString().Substring(6, 2)) == 9)
                        {
                            message1[8] = message1[8] + table.Rows[j].ItemArray[k].ToString();
                            message1[8] = message1[8] + '\n';
                        }
                        else if (Convert.ToInt32(table.Rows[j].ItemArray[0].ToString().Substring(6, 2)) == 10)
                        {
                            message1[9] = message1[9] + table.Rows[j].ItemArray[k].ToString();
                            message1[9] = message1[9] + '\n';
                        }
                        else if (Convert.ToInt32(table.Rows[j].ItemArray[0].ToString().Substring(6, 2)) == 11)
                        {
                            message1[10] = message1[10] + table.Rows[j].ItemArray[k].ToString();
                            message1[10] = message1[10] + '\n';
                        }
                        else if (Convert.ToInt32(table.Rows[j].ItemArray[0].ToString().Substring(6, 2)) == 12)
                        {
                            message1[11] = message1[11] + table.Rows[j].ItemArray[k].ToString();
                            message1[11] = message1[11] + '\n';
                        }
                        else if (Convert.ToInt32(table.Rows[j].ItemArray[0].ToString().Substring(6, 2)) == 13)
                        {
                            message1[12] = message1[12] + table.Rows[j].ItemArray[k].ToString();
                            message1[12] = message1[12] + '\n';
                        }
                        else if (Convert.ToInt32(table.Rows[j].ItemArray[0].ToString().Substring(6, 2)) == 14)
                        {
                            message1[13] = message1[13] + table.Rows[j].ItemArray[k].ToString();
                            message1[13] = message1[13] + '\n';
                        }
                        else if (Convert.ToInt32(table.Rows[j].ItemArray[0].ToString().Substring(6, 2)) == 15)
                        {
                            message1[14] = message1[14] + table.Rows[j].ItemArray[k].ToString();
                            message1[14] = message1[14] + '\n';
                        }
                        else if (Convert.ToInt32(table.Rows[j].ItemArray[0].ToString().Substring(6, 2)) == 16)
                        {
                            message1[15] = message1[15] + table.Rows[j].ItemArray[k].ToString();
                            message1[15] = message1[15] + '\n';
                        }
                        else if (Convert.ToInt32(table.Rows[j].ItemArray[0].ToString().Substring(6, 2)) == 17)
                        {
                            message1[16] = message1[16] + table.Rows[j].ItemArray[k].ToString();
                            message1[16] = message1[16] + '\n';
                        }
                        else if (Convert.ToInt32(table.Rows[j].ItemArray[0].ToString().Substring(6, 2)) == 18)
                        {
                            message1[17] = message1[17] + table.Rows[j].ItemArray[k].ToString();
                            message1[17] = message1[17] + '\n';
                        }
                        else if (Convert.ToInt32(table.Rows[j].ItemArray[0].ToString().Substring(6, 2)) == 19)
                        {
                            message1[18] = message1[18] + table.Rows[j].ItemArray[k].ToString();
                            message1[18] = message1[18] + '\n';
                        }
                        else if (Convert.ToInt32(table.Rows[j].ItemArray[0].ToString().Substring(6, 2)) == 20)
                        {
                            message1[19] = message1[19] + table.Rows[j].ItemArray[k].ToString();
                            message1[19] = message1[19] + '\n';
                        }
                        else if (Convert.ToInt32(table.Rows[j].ItemArray[0].ToString().Substring(6, 2)) == 21)
                        {
                            message1[20] = message1[20] + table.Rows[j].ItemArray[k].ToString();
                            message1[20] = message1[20] + '\n';
                        }
                        else if (Convert.ToInt32(table.Rows[j].ItemArray[0].ToString().Substring(6, 2)) == 22)
                        {
                            message1[21] = message1[21] + table.Rows[j].ItemArray[k].ToString();
                            message1[21] = message1[21] + '\n';
                        }
                        else if (Convert.ToInt32(table.Rows[j].ItemArray[0].ToString().Substring(6, 2)) == 23)
                        {
                            message1[22] = message1[22] + table.Rows[j].ItemArray[k].ToString();
                            message1[22] = message1[22] + '\n';
                        }
                        else if (Convert.ToInt32(table.Rows[j].ItemArray[0].ToString().Substring(6, 2)) == 24)
                        {
                            message1[23] = message1[23] + table.Rows[j].ItemArray[k].ToString();
                            message1[23] = message1[23] + '\n';
                        }
                        else if (Convert.ToInt32(table.Rows[j].ItemArray[0].ToString().Substring(6, 2)) == 25)
                        {
                            message1[24] = message1[24] + table.Rows[j].ItemArray[k].ToString();
                            message1[24] = message1[24] + '\n';
                        }
                        else if (Convert.ToInt32(table.Rows[j].ItemArray[0].ToString().Substring(6, 2)) == 26)
                        {
                            message1[25] = message1[25] + table.Rows[j].ItemArray[k].ToString();
                            message1[25] = message1[25] + '\n';
                        }
                        else if (Convert.ToInt32(table.Rows[j].ItemArray[0].ToString().Substring(6, 2)) == 27)
                        {
                            message1[26] = message1[26] + table.Rows[j].ItemArray[k].ToString();
                            message1[26] = message1[26] + '\n';
                        }
                        else if (Convert.ToInt32(table.Rows[j].ItemArray[0].ToString().Substring(6, 2)) == 28)
                        {
                            message1[27] = message1[27] + table.Rows[j].ItemArray[k].ToString();
                            message1[27] = message1[27] + '\n';
                        }
                        else if (Convert.ToInt32(table.Rows[j].ItemArray[0].ToString().Substring(6, 2)) == 29)
                        {
                            message1[28] = message1[28] + table.Rows[j].ItemArray[k].ToString();
                            message1[28] = message1[28] + '\n';
                        }
                        else if (Convert.ToInt32(table.Rows[j].ItemArray[0].ToString().Substring(6, 2)) == 30)
                        {
                            message1[29] = message1[29] + table.Rows[j].ItemArray[k].ToString();
                            message1[29] = message1[29] + '\n';
                        }
                        else if (Convert.ToInt32(table.Rows[j].ItemArray[0].ToString().Substring(6, 2)) == 31)
                        {
                            message1[30] = message1[30] + table.Rows[j].ItemArray[k].ToString();
                            message1[30] = message1[30] + '\n';
                        }
                    }
                    //message = message + '\n';
                }

                excelWorkSheet.Cells[1, 1] = "星期日";
                excelWorkSheet.Cells[1, 2] = "星期一";
                excelWorkSheet.Cells[1, 3] = "星期二";
                excelWorkSheet.Cells[1, 4] = "星期三";
                excelWorkSheet.Cells[1, 5] = "星期四";
                excelWorkSheet.Cells[1, 6] = "星期五";
                excelWorkSheet.Cells[1, 7] = "星期六";

                //置中
                string RangeCenter = "A1:G1";//設定範圍
                excelWorkSheet.get_Range(RangeCenter).HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

                for (int i = 1; i <= MONTHDAYS; i++)
                {

                    EXCELX = 2 + Convert.ToInt32(Math.Truncate(Convert.ToDouble((i + days - 1) / 7)));
                    EXCELY = (days + i) % 7;
                    if (EXCELY == 0)
                    {
                        EXCELY = 7;
                    }

                    //excelWorkSheet.Cells[EXCELX, EXCELY] = i;

                    excelWorkSheet.Cells[EXCELX, EXCELY] = message1[i - 1].ToString();

                    //if (!string.IsNullOrEmpty(message[i-1].ToString()))
                    //{
                    //    excelWorkSheet.Cells[EXCELX, EXCELY] = message[i - 1].ToString();
                    //}

                }
                //excelWorkSheet.Cells[1, 1] = dateTimePicker9.Value.ToString("yyyy/MM/") + "01";
                //excelWorkSheet.Cells[2, days+1] = message1;
                //message1 = null;


                //靠左
                string RangeLeft = "A2:G6";//設定範圍
                excelWorkSheet.get_Range(RangeLeft).HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;

                //設定為按照內容自動調整欄寬
                //excelWorkSheet.get_Range(RangeLeft).Columns.AutoFit();
                excelWorkSheet.get_Range(RangeLeft).ColumnWidth = 30;
                //excelWorkSheet.Columns.AutoFit();

                // 給儲存格加邊框
                excelWorkSheet.get_Range(RangeLeft).Borders.LineStyle = Excel.XlBorderWeight.xlHairline;
                //excelWorkSheet.get_Range(RangeLeft).Borders.LineStyle = Excel.XlBorderWeight.xlMedium;
            }



            excelWorkBook.Save();
            excelWorkBook.Close();
            excelApp.Quit();

        }


      
        #endregion

        #region BUTTON

        private void button6_Click(object sender, EventArgs e)
        {
            SETCALENDAR();
        }

        private void button1_Click(object sender, EventArgs e)
        {

        }

        private void button3_Click(object sender, EventArgs e)
        {
            SETFASTREPORT();
        }

        private void button4_Click(object sender, EventArgs e)
        {
            RESET1();
            SETPATH1();
            SETFILE1();

            MessageBox.Show("OK");
        }

        #endregion


    }
}

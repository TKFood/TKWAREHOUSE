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
        int result;
        public Report report1 { get; private set; }

        public FrmCALENDERPUR()
        {
            InitializeComponent();

            SETCALENDAR();
        }

        #region FUNCTION

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

        }

        #endregion


    }
}

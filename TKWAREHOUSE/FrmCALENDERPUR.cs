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
            calendar1.CalendarDate = new DateTime(DateTime.Now.Year, DateTime.Now.Month, DateTime.Now.Day, 0, 0, 0);
            calendar1.CalendarView = CalendarViews.Month;
            calendar1.AllowEditingEvents = true;




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

        #endregion

        #region BUTTON

        private void button6_Click(object sender, EventArgs e)
        {
            SETCALENDAR();
        }

        private void button1_Click(object sender, EventArgs e)
        {

        }

        #endregion
    }
}

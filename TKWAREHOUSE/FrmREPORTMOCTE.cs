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
    public partial class FrmREPORTMOCTE : Form
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
        string tablename = null;
        int rownum = 0;
        DataGridViewRow row;
        string SALSESID = null;
        int result;

        public Report report1 { get; private set; }

        public FrmREPORTMOCTE()
        {
            InitializeComponent();
            combobox2load();
        }

        #region FUNCTION
        public void combobox2load()
        {

            connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
            sqlConn = new SqlConnection(connectionString);
            String Sequel = "SELECT MD002,MD001 FROM [TK].dbo.CMSMD WHERE MD002 LIKE '新%' ORDER BY MD001";
            SqlDataAdapter da = new SqlDataAdapter(Sequel, sqlConn);
            DataTable dt = new DataTable();
            sqlConn.Open();

            dt.Columns.Add("MD001", typeof(string));
            dt.Columns.Add("MD002", typeof(string));
            da.Fill(dt);
            comboBox2.DataSource = dt.DefaultView;
            comboBox2.ValueMember = "MD001";
            comboBox2.DisplayMember = "MD002";
            sqlConn.Close();

           

        }

        public void SETFASTREPORT()
        {

            string SQL;
            report1 = new Report();
            report1.Load(@"REPORT\合併領料.frx");

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

            if(comboBox1.Text.ToString().Equals("原料"))
            {
                FASTSQL.AppendFormat(@" SELECT MD002,TE004,TE017 ,TE011,TE012,SUM(TE005) AS TE005,TE010 ");
                FASTSQL.AppendFormat(@" FROM [TK].dbo.CMSMD, [TK].dbo.MOCTC,[TK].dbo.MOCTE ");
                FASTSQL.AppendFormat(@" WHERE MD002 LIKE '新%' ");
                FASTSQL.AppendFormat(@" AND MD001=TC005 ");
                FASTSQL.AppendFormat(@" AND TC001=TE001 AND TC002=TE002 ");
                FASTSQL.AppendFormat(@" AND TE004 LIKE '1%' ");
                FASTSQL.AppendFormat(@" AND TC003>={0}  AND TC003<={1} ", dateTimePicker1.Value.ToString("yyyyMMdd"), dateTimePicker2.Value.ToString("yyyyMMdd"));
                FASTSQL.AppendFormat(@" AND MD002='{0}' ", comboBox2.Text.ToString());
                FASTSQL.AppendFormat(@" GROUP BY MD002,TE004,TE017 ,TE011,TE012,TE010 ");
                FASTSQL.AppendFormat(@" ORDER BY MD002,TE004,TE017 ,TE011,TE012,TE010 ");
            }
            else if (comboBox1.Text.ToString().Equals("物料"))
            {
                FASTSQL.AppendFormat(@" SELECT MD002,TE004,TE017 ,TE011,TE012,SUM(TE005) AS TE005,TE010 ");
                FASTSQL.AppendFormat(@" FROM [TK].dbo.CMSMD, [TK].dbo.MOCTC,[TK].dbo.MOCTE ");
                FASTSQL.AppendFormat(@" WHERE MD002 LIKE '新%' ");
                FASTSQL.AppendFormat(@" AND MD001=TC005 ");
                FASTSQL.AppendFormat(@" AND TC001=TE001 AND TC002=TE002 ");
                FASTSQL.AppendFormat(@" AND TE004 LIKE '2%' ");
                FASTSQL.AppendFormat(@" AND TC003>={0}  AND TC003<={1} ", dateTimePicker1.Value.ToString("yyyyMMdd"), dateTimePicker2.Value.ToString("yyyyMMdd"));
                FASTSQL.AppendFormat(@" AND MD002='{0}' ", comboBox2.Text.ToString());
                FASTSQL.AppendFormat(@" GROUP BY MD002,TE004,TE017 ,TE011,TE012,TE010 ");
                FASTSQL.AppendFormat(@" ORDER BY MD002,TE004,TE017 ,TE011,TE012,TE010 ");
            }


            FASTSQL.AppendFormat(@"  ");

            return FASTSQL.ToString();
        }
        #endregion

        #region BUTTON
        private void button1_Click(object sender, EventArgs e)
        {
            SETFASTREPORT();
        }


      

        #endregion


    }
}

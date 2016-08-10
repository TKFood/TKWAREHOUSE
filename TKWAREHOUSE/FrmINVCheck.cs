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
                   
                sbSql.Append(@" SELECT LA001 AS '品號' ,MB002 AS '品名' ,SUM(LA005*LA011) AS '庫存量'  ");
                sbSql.AppendFormat(@"  FROM [{0}].dbo.INVLA WITH (NOLOCK) LEFT JOIN  [{0}].dbo.INVMB WITH (NOLOCK) ON MB001=LA001 ", sqlConn.Database.ToString());
                sbSql.Append(@" WHERE LA001 LIKE '4%' AND (LA009='20001')");
                sbSql.Append(@" GROUP BY LA001,MB002");
                sbSql.Append(@" HAVING SUM(LA005*LA011)>1");
                sbSql.Append(@" ORDER BY LA001,MB002");


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
        #endregion

        #region BUTTON
        private void button1_Click(object sender, EventArgs e)
        {
            Search();
        }
        #endregion

    }


}

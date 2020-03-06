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
using System.Collections;

namespace TKWAREHOUSE
{
    public partial class FrmPURCOPCOMMENT : Form
    {

        SqlConnection sqlConn = new SqlConnection();
        SqlCommand sqlComm = new SqlCommand();
        string connectionString;
        StringBuilder sbSql = new StringBuilder();
        StringBuilder sbSqlQuery = new StringBuilder();
        SqlDataAdapter adapter = new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilder = new SqlCommandBuilder();
        SqlDataAdapter adapter2 = new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilder2 = new SqlCommandBuilder();
        SqlDataAdapter adapter3 = new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilder3 = new SqlCommandBuilder();
        SqlDataAdapter adapter4 = new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilder4 = new SqlCommandBuilder();

        DataSet ds = new DataSet();
        DataSet ds2 = new DataSet();
        DataSet ds3 = new DataSet();
        DataSet ds4 = new DataSet();

        int result;
        string tablename = null;

        string PURTA001;
        string PURTA002;

        public FrmPURCOPCOMMENT()
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

                sbSql.AppendFormat(@"  SELECT TA001 AS '請購單別',TA002 AS '請購單號',TA006 AS '備註'");
                sbSql.AppendFormat(@"  FROM [TK].dbo.PURTA");
                sbSql.AppendFormat(@"  WHERE TA003>='{0}' AND TA003<='{1}'",dateTimePicker1.Value.ToString("yyyyMMdd"), dateTimePicker2.Value.ToString("yyyyMMdd"));
                sbSql.AppendFormat(@"  ORDER BY TA001,TA002");
                sbSql.AppendFormat(@"  ");

                adapter = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder = new SqlCommandBuilder(adapter);
                sqlConn.Open();
                ds.Clear();
                adapter.Fill(ds, "ds");
                sqlConn.Close();
                

                if (ds.Tables["ds"].Rows.Count == 0)
                {
                    dataGridView1.DataSource = null;
                }
                else
                {
                    if (ds.Tables["ds"].Rows.Count >= 1)
                    {
                        dataGridView1.DataSource = ds.Tables["ds"];
                        dataGridView1.AutoResizeColumns();
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

        private void dataGridView1_SelectionChanged(object sender, EventArgs e)
        {
            if (dataGridView1.CurrentRow != null)
            {
                int rowindex = dataGridView1.CurrentRow.Index;
                if (rowindex >= 0)
                {
                    DataGridViewRow row = dataGridView1.Rows[rowindex];

                    PURTA001 = row.Cells["請購單別"].Value.ToString().Trim();
                    PURTA002 = row.Cells["請購單號"].Value.ToString().Trim();

                    Search2(PURTA001, PURTA002);
                }
                else
                {
                    PURTA001 = null;
                    PURTA002 = null;

                }
            }
           
        }

        public void Search2(string PURTA001,string PURTA002)
        {
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sbSql.Clear();
                sbSqlQuery.Clear();
                
                sbSql.AppendFormat(@"  SELECT [COPTC001] AS '訂單單別',[COPTC002] AS '訂單單號',[COMMENT] AS '備註',[PURTA001] AS '請購單別',[PURTA002] AS '請購單號',[ID] ");
                sbSql.AppendFormat(@"  FROM [TKWAREHOUSE].[dbo].[PURCOPCOMMENT]");
                sbSql.AppendFormat(@"  WHERE [PURTA001]='{0}' AND [PURTA002]='{1}' ",PURTA001,PURTA002);
                sbSql.AppendFormat(@"  ORDER BY [ID] ");
                sbSql.AppendFormat(@"  ");

                adapter2 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder2 = new SqlCommandBuilder(adapter2);
                sqlConn.Open();
                ds2.Clear();
                adapter2.Fill(ds2, "ds2");
                sqlConn.Close();


                if (ds2.Tables["ds2"].Rows.Count == 0)
                {
                    dataGridView2.DataSource = null;
                }
                else
                {
                    if (ds2.Tables["ds2"].Rows.Count >= 1)
                    {
                        dataGridView2.DataSource = ds2.Tables["ds2"];
                        dataGridView2.AutoResizeColumns();
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

        private void button4_Click(object sender, EventArgs e)
        {
            Search();
        }
        #endregion

      
    }
}

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
    public partial class FrmPURTABCD : Form
    {
        SqlConnection sqlConn = new SqlConnection();
        SqlCommand sqlComm = new SqlCommand();
        string connectionString;
        SqlTransaction tran;
        SqlCommand cmd = new SqlCommand();
        StringBuilder sbSql = new StringBuilder();
        StringBuilder sbSqlQuery = new StringBuilder();
        SqlDataAdapter adapter = new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilder = new SqlCommandBuilder();
        DataSet ds = new DataSet();
        DataSet ds2 = new DataSet();

        DataTable dt = new DataTable();


        public FrmPURTABCD()
        {
            InitializeComponent();
        }


        #region FUNCTION
        public void SEARCH()
        {
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);
                SqlDataAdapter adapter = new SqlDataAdapter();
                SqlCommandBuilder sqlCmdBuilder = new SqlCommandBuilder();
                DataSet ds = new DataSet();

                sbSql.Clear();
                sbSqlQuery.Clear();

                sbSql.AppendFormat(@"  
                                    SELECT TA001 AS '請購單別',TA002 AS '請購單號',TB003 AS '請購序號',TB004 AS '品號',TB005 AS '品名',TB006 AS '規格',TB009 AS '請購數量',TB007 AS '請購單位',TB022 AS '採購單',TD008 AS '採購數量',TD009 AS '採購單位'
                                    FROM [TK].dbo.PURTA,[TK].dbo.PURTB
                                    LEFT JOIN [TK].dbo.[PURTD] ON TD001=SUBSTRING(TB022,1,4) AND TD002=SUBSTRING(TB022,6,11) AND TD003=SUBSTRING(TB022,18,4)
                                    WHERE TA001=TB001 AND TA002=TB002
                                    AND TA003>='{0}' AND TA003<='{1}'
                                    ", dateTimePicker1.Value.ToString("yyyyMMdd"), dateTimePicker2.Value.ToString("yyyyMMdd"));

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
                        //dataGridView1.Rows.Clear();
                        dataGridView1.DataSource = ds.Tables["ds"];
                        dataGridView1.AutoResizeColumns();
                        //dataGridView1.CurrentCell = dataGridView1[0, rownum];

                        dataGridView1.AutoResizeColumns();
                        dataGridView1.ColumnHeadersDefaultCellStyle.Font = new Font("Tahoma", 9);
                        dataGridView1.DefaultCellStyle.Font = new Font("Tahoma", 10);
                        dataGridView1.Columns["請購單別"].Width = 60;
                        dataGridView1.Columns["請購單號"].Width = 100;
                        dataGridView1.Columns["請購序號"].Width = 60;
                        dataGridView1.Columns["品號"].Width = 120;
                        dataGridView1.Columns["品名"].Width = 150;
                        dataGridView1.Columns["規格"].Width = 120;
                        dataGridView1.Columns["請購數量"].Width = 60;
                        dataGridView1.Columns["請購單位"].Width = 60;
                        dataGridView1.Columns["採購單"].Width = 200;
                        dataGridView1.Columns["採購數量"].Width = 60;
                        dataGridView1.Columns["採購單位"].Width = 60;
                    }
                }

            }
            catch
            {

            }
            finally
            {
                sqlConn.Close();
            }

        }
        #endregion

        #region BUTTON
        private void button3_Click(object sender, EventArgs e)
        {
            SEARCH();
        }
        #endregion
    }
}

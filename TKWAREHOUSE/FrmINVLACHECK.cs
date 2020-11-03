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
    public partial class FrmINVLACHECK : Form
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

        public FrmINVLACHECK()
        {
            InitializeComponent();
        }

        #region FUNCTION
        public void SEARCH()
        {
            SqlDataAdapter adapter = new SqlDataAdapter();
            SqlCommandBuilder sqlCmdBuilder = new SqlCommandBuilder();
            DataSet ds = new DataSet();

            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sbSql.Clear();

                sbSql.AppendFormat(@"  
                                    SELECT MB001 AS '品號',MB002  AS '品名' 
                                    FROM [TK].dbo.INVMB
                                    WHERE (MB001 LIKE '%{0}%' OR MB002 LIKE '%{0}%')
                                    ORDER BY  MB001,MB002 
                                    ", textBox1.Text.Trim());


                adapter = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder = new SqlCommandBuilder(adapter);
                sqlConn.Open();
                ds.Clear();
                adapter.Fill(ds, "TEMPds");
                sqlConn.Close();


                if (ds.Tables["TEMPds"].Rows.Count == 0)
                {
                    dataGridView1.DataSource = null;
                }
                else
                {
                    if (ds.Tables["TEMPds"].Rows.Count >= 1)
                    {
                        dataGridView1.DataSource = ds.Tables["TEMPds"];

                        dataGridView1.AutoResizeColumns();
                        dataGridView1.ColumnHeadersDefaultCellStyle.Font = new Font("Tahoma", 9);
                        dataGridView1.DefaultCellStyle.Font = new Font("Tahoma", 10);
                        dataGridView1.Columns["品號"].Width = 140;
                        dataGridView1.Columns["品名"].Width = 140;

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
        private void dataGridView1_SelectionChanged(object sender, EventArgs e)
        {
            textBox2.Text = null;
            textBox3.Text = null;
            textBox4.Text = null;

            if (dataGridView1.CurrentRow != null)
            {
                int rowindex = dataGridView1.CurrentRow.Index;
                if (rowindex >= 0)
                {
                    DataGridViewRow row = dataGridView1.Rows[rowindex];

                    textBox2.Text = row.Cells["品號"].Value.ToString().Trim();
                    textBox3.Text = row.Cells["品號"].Value.ToString().Trim();
                    textBox4.Text = row.Cells["品號"].Value.ToString().Trim();
                    
                    //SEARCH2(row.Cells["品號"].Value.ToString().Trim());
                    //SEARCH3(row.Cells["品號"].Value.ToString().Trim());

                    //SETFASTREPORT(row.Cells["品號"].Value.ToString().Trim());
                }
            }
        }

        public void SEARCH2(string MD001)
        {
            SqlDataAdapter adapter = new SqlDataAdapter();
            SqlCommandBuilder sqlCmdBuilder = new SqlCommandBuilder();
            DataSet ds = new DataSet();

            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sbSql.Clear();

                sbSql.AppendFormat(@"  
                                   WITH NODE (MD001,MD003,LAYER,MD004,MC004,PREMC004,MD006,MD007,MD008,USEDNUM ,PREUSEDNUM) AS
                                    (
                                    SELECT MD001,MD003,0 ,[MD004],[MC004],[MC004] AS PREMC004,[MD006],[MD007],[MD008],CONVERT(DECIMAL(18,4),([MD006]/[MD007]/[MC004]*(1+MD008))),CONVERT(DECIMAL(18,4),1) AS PREUSEDNUM  FROM [TK].[dbo].[VBOMMD]
                                    UNION ALL
                                    SELECT TB1.MD001,TB2.MD003,TB2.LAYER+1,TB2.MD004,TB2.MC004,TB1.MC004,TB2.MD006,TB2.MD007,TB2.MD008,TB2.USEDNUM,CONVERT(DECIMAL(18,4),(TB1.[MD006]/TB1.[MD007]/TB1.[MC004]*(1+TB1.MD008))) AS PREUSEDNUM FROM [TK].[dbo].[VBOMMD] TB1
                                    INNER JOIN NODE TB2
                                    ON TB1.MD003 = TB2.MD001
                                    )

                                    SELECT MD001 AS '成品號',MB1.MB002 AS '成品名',MD003 AS '原物料',MB2.MB002 AS '原物料名'
                                    ,(SELECT SUM(LA011*LA005) FROM [TK].dbo.INVLA WITH(NOLOCK) WHERE LA009 LIKE '2%' AND  LA001=MD003 ) AS '總數量'
                                    FROM (
                                    SELECT DISTINCT MD001,MD003
                                    FROM NODE
                                    WHERE  MD001='{0}'
                                    ) AS TEMP,[TK].dbo.INVMB MB1,[TK].dbo.INVMB MB2
                                    WHERE  MD001=MB1.MB001 AND MD003=MB2.MB001
                                    ORDER BY MD001,MD003
                                    ", MD001);


                adapter = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder = new SqlCommandBuilder(adapter);
                sqlConn.Open();
                ds.Clear();
                adapter.Fill(ds, "TEMPds");
                sqlConn.Close();


                if (ds.Tables["TEMPds"].Rows.Count == 0)
                {
                    dataGridView2.DataSource = null;
                }
                else
                {
                    if (ds.Tables["TEMPds"].Rows.Count >= 1)
                    {
                        dataGridView2.DataSource = ds.Tables["TEMPds"];

                        dataGridView2.AutoResizeColumns();
                        dataGridView2.ColumnHeadersDefaultCellStyle.Font = new Font("Tahoma", 9);
                        dataGridView2.DefaultCellStyle.Font = new Font("Tahoma", 10);
                        dataGridView2.Columns["成品號"].Width = 100;
                        dataGridView2.Columns["成品名"].Width = 100;
                        dataGridView2.Columns["原物料"].Width = 120;
                        dataGridView2.Columns["原物料名"].Width = 120;
                        dataGridView2.Columns["總數量"].Width = 120;

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

        public void SEARCH3(string MD001)
        {
            SqlDataAdapter adapter = new SqlDataAdapter();
            SqlCommandBuilder sqlCmdBuilder = new SqlCommandBuilder();
            DataSet ds = new DataSet();

            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sbSql.Clear();

                sbSql.AppendFormat(@"  
                                  WITH NODE (MD001,MD003,LAYER,MD004,MC004,PREMC004,MD006,MD007,MD008,USEDNUM ,PREUSEDNUM) AS
                                    (
                                    SELECT MD001,MD003,0 ,[MD004],[MC004],[MC004] AS PREMC004,[MD006],[MD007],[MD008],CONVERT(DECIMAL(18,4),([MD006]/[MD007]/[MC004]*(1+MD008))),CONVERT(DECIMAL(18,4),1) AS PREUSEDNUM  FROM [TK].[dbo].[VBOMMD]
                                    UNION ALL
                                    SELECT TB1.MD001,TB2.MD003,TB2.LAYER+1,TB2.MD004,TB2.MC004,TB1.MC004,TB2.MD006,TB2.MD007,TB2.MD008,TB2.USEDNUM,CONVERT(DECIMAL(18,4),(TB1.[MD006]/TB1.[MD007]/TB1.[MC004]*(1+TB1.MD008))) AS PREUSEDNUM FROM [TK].[dbo].[VBOMMD] TB1
                                    INNER JOIN NODE TB2
                                    ON TB1.MD003 = TB2.MD001
                                    )

                                    SELECT MD001 AS '成品號',MB1.MB002 AS '成品名',MD003 AS '原物料',MB2.MB002 AS '原物料名'
                                    ,(SELECT SUM(LA011*LA005) FROM [TK].dbo.INVLA WITH(NOLOCK) WHERE LA009 LIKE '2%' AND  LA001=MD003 )  AS '總數量'
                                    ,SUM(LA011*LA005) AS '批號庫存量',LA016 AS '批號'  ,LA009 AS '庫別' 
                                    FROM (
                                    SELECT DISTINCT MD001,MD003
                                    FROM NODE
                                    WHERE  MD001='{0}'
                                    ) AS TEMP,[TK].dbo.INVMB MB1,[TK].dbo.INVMB MB2,[TK].dbo.INVLA WITH(NOLOCK)
                                    WHERE  MD001=MB1.MB001 AND MD003=MB2.MB001
                                    AND LA001=MD003
                                    AND LA009 LIKE '2%'
                                    AND ISNULL(LA016,'')<>''
                                    AND MD001='{0}'
                                    GROUP BY MD001,MB1.MB002,MD003,MB2.MB002,LA009,LA016
                                    HAVING SUM(LA011*LA005)<>0
                                    ORDER BY MD001,MB1.MB002,MD003,MB2.MB002,LA009,LA016
                                    ", MD001);


                adapter = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder = new SqlCommandBuilder(adapter);
                sqlConn.Open();
                ds.Clear();
                adapter.Fill(ds, "TEMPds");
                sqlConn.Close();


                if (ds.Tables["TEMPds"].Rows.Count == 0)
                {
                    dataGridView3.DataSource = null;
                }
                else
                {
                    if (ds.Tables["TEMPds"].Rows.Count >= 1)
                    {
                        dataGridView3.DataSource = ds.Tables["TEMPds"];

                        dataGridView3.AutoResizeColumns();
                        dataGridView3.ColumnHeadersDefaultCellStyle.Font = new Font("Tahoma", 9);
                        dataGridView3.DefaultCellStyle.Font = new Font("Tahoma", 10);
                        dataGridView3.Columns["成品號"].Width = 10;
                        dataGridView3.Columns["成品名"].Width = 10;
                        dataGridView3.Columns["原物料"].Width = 120;
                        dataGridView3.Columns["原物料名"].Width = 120;
                        dataGridView3.Columns["總數量"].Width = 120;
                        dataGridView3.Columns["批號庫存量"].Width = 120;
                        dataGridView3.Columns["批號"].Width = 120;
                        dataGridView3.Columns["庫別"].Width = 120;

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

        public void SETFASTREPORT(string MD001)
        {

            string SQL;
            string SQL1;
            report1 = new Report();
            report1.Load(@"REPORT\成品原物料追踨.frx");

            report1.Dictionary.Connections[0].ConnectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;

            TableDataSource Table = report1.GetDataSource("Table") as TableDataSource;
          
            SQL = SETFASETSQL(MD001);

            Table.SelectCommand = SQL;
            report1.Preview = previewControl1;
            report1.Show();

        }

        public string SETFASETSQL(string MD001)
        {
            StringBuilder FASTSQL = new StringBuilder();
            StringBuilder STRQUERY = new StringBuilder();

            FASTSQL.AppendFormat(@"  
                                  WITH NODE (MD001,MD003,LAYER,MD004,MC004,PREMC004,MD006,MD007,MD008,USEDNUM ,PREUSEDNUM) AS
                                    (
                                    SELECT MD001,MD003,0 ,[MD004],[MC004],[MC004] AS PREMC004,[MD006],[MD007],[MD008],CONVERT(DECIMAL(18,4),([MD006]/[MD007]/[MC004]*(1+MD008))),CONVERT(DECIMAL(18,4),1) AS PREUSEDNUM  FROM [TK].[dbo].[VBOMMD]
                                    UNION ALL
                                    SELECT TB1.MD001,TB2.MD003,TB2.LAYER+1,TB2.MD004,TB2.MC004,TB1.MC004,TB2.MD006,TB2.MD007,TB2.MD008,TB2.USEDNUM,CONVERT(DECIMAL(18,4),(TB1.[MD006]/TB1.[MD007]/TB1.[MC004]*(1+TB1.MD008))) AS PREUSEDNUM FROM [TK].[dbo].[VBOMMD] TB1
                                    INNER JOIN NODE TB2
                                    ON TB1.MD003 = TB2.MD001
                                    )

                                    SELECT MD001 AS '成品號',MB1.MB002 AS '成品名',MD003 AS '原物料',MB2.MB002 AS '原物料名'
                                    ,(SELECT SUM(LA011*LA005) FROM [TK].dbo.INVLA WITH(NOLOCK) WHERE LA009 LIKE '2%' AND  LA001=MD003 )  AS '總數量'
                                    ,SUM(LA011*LA005) AS '批號庫存量',LA016 AS '批號'  ,LA009 AS '庫別' 
                                    FROM (
                                    SELECT DISTINCT MD001,MD003
                                    FROM NODE
                                    WHERE  MD001='{0}'
                                    ) AS TEMP,[TK].dbo.INVMB MB1,[TK].dbo.INVMB MB2,[TK].dbo.INVLA WITH(NOLOCK)
                                    WHERE  MD001=MB1.MB001 AND MD003=MB2.MB001
                                    AND LA001=MD003
                                    AND LA009 LIKE '2%'
                                    AND ISNULL(LA016,'')<>''
                                    AND MD001='{0}'
                                    GROUP BY MD001,MB1.MB002,MD003,MB2.MB002,LA009,LA016
                                    HAVING SUM(LA011*LA005)<>0
                                    ORDER BY MD001,MB1.MB002,MD003,MB2.MB002,LA009,LA016
                                    ", MD001);


            return FASTSQL.ToString();
        }

        #endregion

        #region BUTTON

        private void button1_Click(object sender, EventArgs e)
        {
            SEARCH();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            SEARCH2(textBox2.Text.Trim());
        }

        private void button3_Click(object sender, EventArgs e)
        {
            SEARCH3(textBox3.Text.Trim());
        }

        private void button4_Click(object sender, EventArgs e)
        {
            SETFASTREPORT(textBox4.Text.Trim());
        }

        #endregion


    }
}

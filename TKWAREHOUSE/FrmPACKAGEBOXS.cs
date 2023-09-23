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
using System.Text.RegularExpressions;
using FastReport;
using FastReport.Data;
using TKITDLL;
using System.Data.OleDb;
using System.Net;

namespace TKWAREHOUSE
{
    public partial class FrmPACKAGEBOXS : Form
    {
        StringBuilder sbSql = new StringBuilder();
        SqlConnection sqlConn = new SqlConnection();
        SqlDataAdapter adapter = new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilder = new SqlCommandBuilder();
        SqlTransaction tran;
        SqlCommand cmd = new SqlCommand();
        DataSet ds = new DataSet();
        DataTable dt = new DataTable();
        string tablename = null;
        int rownum = 0;
        DataGridViewRow row;
        int result;
        string SortedColumn = string.Empty;
        string SortedModel = string.Empty;

        private bool isTextBox76Changing = false;
        private bool isTextBox77Changing = false;

        public FrmPACKAGEBOXS()
        {
            InitializeComponent();
        }

        #region FUNCTION
        public void SEARCH(string QUERY, DataGridView DataGridViewNew, string SortedColumn, string SortedModel)
        {
            SqlConnection sqlConn = new SqlConnection();
            SqlCommand sqlComm = new SqlCommand();
            SqlDataAdapter SqlDataAdapterNEW = new SqlDataAdapter();
            SqlCommandBuilder SqlCommandBuilderNEW = new SqlCommandBuilder();
            DataSet DataSetNEW = new DataSet();

            try
            {
                //20210902密
                Class1 TKID = new Class1();//用new 建立類別實體
                SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

                //資料庫使用者密碼解密
                sqlsb.Password = TKID.Decryption(sqlsb.Password);
                sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

                String connectionString;
                sqlConn = new SqlConnection(sqlsb.ConnectionString);

                SqlDataAdapterNEW = new SqlDataAdapter(@"" + sbSql, sqlConn);

                SqlCommandBuilderNEW = new SqlCommandBuilder(SqlDataAdapterNEW);
                sqlConn.Open();
                DataSetNEW.Clear();
                SqlDataAdapterNEW.Fill(DataSetNEW, "DataSetNEW");
                sqlConn.Close();


                DataGridViewNew.DataSource = null;

                if (DataSetNEW.Tables["DataSetNEW"].Rows.Count >= 1)
                {
                    //DataGridViewNew.Rows.Clear();
                    DataGridViewNew.DataSource = DataSetNEW.Tables["DataSetNEW"];
                    DataGridViewNew.AutoResizeColumns();
                    //DataGridViewNew.CurrentCell = dataGridView1[0, rownum];
                    //dataGridView20SORTNAME
                    //dataGridView20SORTMODE
                    if (!string.IsNullOrEmpty(SortedColumn))
                    {
                        if (SortedModel.Equals("Ascending"))
                        {
                            DataGridViewNew.Sort(DataGridViewNew.Columns["" + SortedColumn + ""], ListSortDirection.Ascending);
                        }
                        else
                        {
                            DataGridViewNew.Sort(DataGridViewNew.Columns["" + SortedColumn + ""], ListSortDirection.Descending);
                        }
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
        public void Search_COPTG(string TG002)
        {
            StringBuilder sbSqlQuery1 = new StringBuilder();
            StringBuilder sbSqlQuery2 = new StringBuilder();

            sbSql.Clear();
            sbSqlQuery1.Clear();
            sbSqlQuery2.Clear();

            if (!string.IsNullOrEmpty(TG002))
            {
                sbSqlQuery1.AppendFormat(@" AND TG002 LIKE '{0}%'", TG002);
            }
            else
            {
                sbSqlQuery1.AppendFormat(@" ");
            }
           


            sbSql.AppendFormat(@"
                                SELECT 
                                TG001 AS '銷貨單',TG002 AS '銷貨單號',TG076 AS '收貨人',TG029 AS '官網訂單'
                                ,(SELECT TOP 1 TH074 FROM [TK].dbo.COPTH WHERE TG001=TH001 AND TG002=TH002 ORDER BY  TH001)  AS '通路訂單'
                                FROM [TK].dbo.COPTG
                                WHERE 1=1
                                {0}
                                AND TG001 IN (SELECT [NAMES]  FROM [TKWAREHOUSE].[dbo].[TBPARAS] WHERE KINDS='COPTGTG001')
                                  ", sbSqlQuery1.ToString());
            sbSql.AppendFormat(@"  ");

            SEARCH(sbSql.ToString(), dataGridView1, SortedColumn, SortedModel);

        }

        private void dataGridView1_SelectionChanged(object sender, EventArgs e)
        {
            string TG001TG002 = "";

            if (dataGridView1.CurrentRow != null)
            {
                int rowindex = dataGridView1.CurrentRow.Index;
                if (rowindex >= 0)
                {
                    DataGridViewRow row = dataGridView1.Rows[rowindex];
                    TG001TG002 = row.Cells["銷貨單"].Value.ToString()+ row.Cells["銷貨單號"].Value.ToString();

                    Search_PACKAGEBOXS(TG001TG002);
                }
            }
        }

        public void Search_PACKAGEBOXS(string TG001TG002)
        {
            StringBuilder sbSqlQuery1 = new StringBuilder();
            StringBuilder sbSqlQuery2 = new StringBuilder();

            sbSql.Clear();
            sbSqlQuery1.Clear();
            sbSqlQuery2.Clear();

            sbSql.AppendFormat(@"
                                SELECT 
                                [BOXNO] AS '箱號'
                                ,[ALLWEIGHTS] AS '秤總重(A+B)'
                                ,[PACKWEIGHTS] AS '網購包材重量(KG)A'
                                ,[PRODUCTWEIGHTS] AS '商品總重量(KG)B'
                                ,[PACKRATES] AS '實際比值'
                                ,[RATECLASS] AS '商品總重量比值分類'
                                ,[CHECKRATES] AS '規定比值'
                                ,[ISVALIDS] AS '是否符合'
                                ,[PACKAGENAMES] AS '使用包材名稱/規格'
                                ,[PACKAGEFROM] AS '使用包材來源'
                                ,[TG001] AS '銷貨單'
                                ,[TG002] AS '銷貨單號'
                                ,[NO]
                                FROM [TKWAREHOUSE].[dbo].[PACKAGEBOXS]
                                WHERE TG001+TG002='{0}'
                                ORDER BY [BOXNO]
                                  ", TG001TG002);

            sbSql.AppendFormat(@"  ");

            SEARCH(sbSql.ToString(), dataGridView2, SortedColumn, SortedModel);

        }

        #endregion


        #region BUTTON
        private void button1_Click(object sender, EventArgs e)
        {
            Search_COPTG(dateTimePicker1.Value.ToString("yyyyMMdd"));
        }

        #endregion

       
    }
}

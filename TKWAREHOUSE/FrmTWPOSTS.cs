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
using TKITDLL;

namespace TKWAREHOUSE
{
    public partial class FrmTWPOSTS : Form
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
        DataSet ds1 = new DataSet();
        DataSet ds2 = new DataSet();
        string tablename = null;
        int rownum = 0;
        DataGridViewRow row;   
        int result;
        public Report report1 { get; private set; }

        public FrmTWPOSTS()
        {
            InitializeComponent();
        }

        #region FUNCTION
        public void SEARCH(string SDAYS,string EDAYS)
        {
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


                sbSql.Clear();
                sbSqlQuery.Clear();
             
                sbSql.AppendFormat(@"  

                                        SELECT 
                                        [ID] AS 'ID'
                                        ,[PAYMONEYS] AS '金額'
                                        ,CONVERT(NVARCHAR,[SENDDATES],112) AS '交寄日期'
                                        ,[WEIGHETS] AS '重量'
                                        ,[ISSINGALS] AS '單筆單件'
                                        ,[CUSTOMERNO] AS '客戶編號'
                                        ,[CUSTOMERNAMES] AS '客戶名稱'
                                        ,[PHONES] AS '電話'
                                        ,[ZIPCODE] AS '郵遞區號'
                                        ,[ADDRESS] AS '地址'
                                        ,[SENDCONTENTS] AS '內裝物品Memo'
                                        ,[SENDNUMS] AS '件數編號'
                                        ,[COMMENTS] AS '備註(出貨單編號)'
                                        ,[USEDUNITS] AS '使用單位編號'
                                        ,[SENDNO] AS '託運單編號'
                                        ,[MOBILEPHONE] AS '手機'
                                        ,[COLMONEYS] AS '代收貨價'

                                        FROM [TKWAREHOUSE].[dbo].[TWPOSTS]
                                        WHERE CONVERT(NVARCHAR,[SENDDATES],112)>='{0}' AND CONVERT(NVARCHAR,[SENDDATES],112)<='{1}'
                                        ORDER BY CONVERT(NVARCHAR,[SENDDATES],112),ID
                                        ", SDAYS, EDAYS);

                adapter = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder = new SqlCommandBuilder(adapter);
                sqlConn.Open();
                ds1.Clear();
                adapter.Fill(ds1, "ds1");
                sqlConn.Close();


                if (ds1.Tables["ds1"].Rows.Count == 0)
                {
                    dataGridView1.DataSource = null;
                }
                else
                {
                    if (ds1.Tables["ds1"].Rows.Count >= 1)
                    {
                        //dataGridView1.Rows.Clear();
                        dataGridView1.DataSource = ds1.Tables["ds1"];
                        dataGridView1.AutoResizeColumns();
                        //dataGridView1.CurrentCell = dataGridView1[0, rownum];

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

        private void dataGridView1_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            string ID = dataGridView1.Rows[e.RowIndex].Cells[0].Value.ToString();
            decimal WEIGHTS =Convert.ToDecimal(dataGridView1.Rows[e.RowIndex].Cells["重量"].Value.ToString());

            int prices = SEARCHTWPOSTSBASE(WEIGHTS);
            MessageBox.Show(prices+" "+ID + " "+e.RowIndex+" "+e.ColumnIndex);
        }

        public void SAVE()
        {
            string ID = null;

            if(ds1.Tables[0].Rows.Count>0)
            {
                foreach (DataRow DR in ds1.Tables[0].Rows)
                {                    
                    ID = ID + "," +DR["ID"].ToString();
                }
            }
           

            MessageBox.Show(ID);
        }

        public int SEARCHTWPOSTSBASE(decimal WEIGHTS)
        {
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


                sbSql.Clear();
                sbSqlQuery.Clear();

                sbSql.AppendFormat(@"  
                                    SELECT 
                                    [SWEIGHTS]
                                    ,[EWEIGHTS]
                                    ,[PRICES]
                                    FROM [TKWAREHOUSE].[dbo].[TWPOSTSBASE]
                                    WHERE [SWEIGHTS]<={0} AND [EWEIGHTS]>={0}
                                    
                                        ", WEIGHTS);

                adapter = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder = new SqlCommandBuilder(adapter);
                sqlConn.Open();
                ds1.Clear();
                adapter.Fill(ds1, "ds1");
                sqlConn.Close();


                if (ds1.Tables["ds1"].Rows.Count >= 1)
                {
                    //dataGridView1.Rows.Clear();
                    dataGridView1.DataSource = ds1.Tables["ds1"];
                    dataGridView1.AutoResizeColumns();
                    //dataGridView1.CurrentCell = dataGridView1[0, rownum];

                    return Convert.ToInt32(ds1.Tables["ds1"].Rows[0]["PRICES"].ToString());

                }
                else
                {
                    return 0;
                }

            }
            catch
            {
                return 0;
            }
            finally
            {
                sqlConn.Close();
            }
        }
        #endregion

        #region BUTTON

        private void button1_Click(object sender, EventArgs e)
        {
            SEARCH(dateTimePicker1.Value.ToString("yyyyMMdd"), dateTimePicker2.Value.ToString("yyyyMMdd"));
        }
        private void button2_Click(object sender, EventArgs e)
        {
            SAVE();
        }

        #endregion


    }
}

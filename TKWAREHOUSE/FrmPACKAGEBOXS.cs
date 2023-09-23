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

            comboBox1load();
            comboBox2load();
            comboBox3load();
        }

        #region FUNCTION
        public void LoadComboBoxData(ComboBox comboBox, string query, string valueMember, string displayMember)
        {
            //20210902密
            Class1 TKID = new Class1();//用new 建立類別實體
            SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

            //資料庫使用者密碼解密
            sqlsb.Password = TKID.Decryption(sqlsb.Password);
            sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

            using (SqlConnection connection = new SqlConnection(sqlsb.ConnectionString))
            {
                SqlCommand command = new SqlCommand(query, connection);
                connection.Open();

                SqlDataAdapter adapter = new SqlDataAdapter(command);
                DataTable dataTable = new DataTable();
                adapter.Fill(dataTable);

                comboBox.DataSource = dataTable;
                comboBox.ValueMember = valueMember;
                comboBox.DisplayMember = displayMember;
            }
        }

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

        public void comboBox1load()
        {
            LoadComboBoxData(comboBox1, "SELECT [ID],[KINDS],[NAMES],[KEYS],[KEYS2] FROM [TKWAREHOUSE].[dbo].[TBPARAS] WHERE [KINDS]='RATECLASS' ORDER BY ID", "NAMES", "NAMES");
        }
        public void comboBox2load()
        {
            LoadComboBoxData(comboBox2, "SELECT [ID],[KINDS],[NAMES],[KEYS],[KEYS2] FROM [TKWAREHOUSE].[dbo].[TBPARAS] WHERE [KINDS]='CHECKRATES' ORDER BY ID", "NAMES", "NAMES");
        }
        public void comboBox3load()
        {
            LoadComboBoxData(comboBox3, "SELECT [NAMES] FROM [TKWAREHOUSE].[dbo].[TBPARAS] WHERE [KINDS]='ISVALIDS' GROUP BY [NAMES]  ", "NAMES", "NAMES");
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

                    DataTable dt = PACKAGEBOXS_FIND(TG001TG002);
                    if(dt!=null&&dt.Rows.Count>=1)
                    {
                        Search_PACKAGEBOXS(TG001TG002);
                    }
                    else
                    {
                        SET_TEXT();
                    }
                    
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

        private void dataGridView2_SelectionChanged(object sender, EventArgs e)
        {
            SET_TEXT();

            DataGridView DV = dataGridView2;

            if (DV.CurrentRow != null)
            {
                int rowindex = DV.CurrentRow.Index;
                if (rowindex >= 0)
                {
                    DataGridViewRow row = DV.Rows[rowindex];
                    textBox1.Text = row.Cells["銷貨單"].Value.ToString()+ row.Cells["銷貨單號"].Value.ToString();
                    textBox2.Text = row.Cells["箱號"].Value.ToString();
                    textBox3.Text = row.Cells["秤總重(A+B)"].Value.ToString();
                    textBox4.Text = row.Cells["網購包材重量(KG)A"].Value.ToString();
                    textBox5.Text = row.Cells["商品總重量(KG)B"].Value.ToString();
                    textBox6.Text = row.Cells["實際比值"].Value.ToString();
                    textBox7.Text = row.Cells["使用包材名稱/規格"].Value.ToString();
                    textBox8.Text = row.Cells["使用包材來源"].Value.ToString();
                    textBox9.Text = row.Cells["NO"].Value.ToString();

                    comboBox1.Text = row.Cells["商品總重量比值分類"].Value.ToString();
                    comboBox2.Text = row.Cells["規定比值"].Value.ToString();
                    comboBox3.Text = row.Cells["是否符合"].Value.ToString();

                    DisplayImageFromFolder(row.Cells["NO"].Value.ToString());
                }
            }
           
        }

        public DataTable PACKAGEBOXS_FIND(string TG001TG002)
        {
            DataTable DT = new DataTable();
            SqlConnection sqlConn = new SqlConnection();
            SqlDataAdapter adapter1 = new SqlDataAdapter();
            SqlCommandBuilder sqlCmdBuilder1 = new SqlCommandBuilder();
            DataSet ds1 = new DataSet();
            StringBuilder QUERYS = new StringBuilder();

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
                QUERYS.Clear();


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




                adapter1 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder1 = new SqlCommandBuilder(adapter1);
                sqlConn.Open();
                ds1.Clear();
                adapter1.Fill(ds1, "TEMPds1");
                sqlConn.Close();


                if (ds1.Tables["TEMPds1"].Rows.Count > 0)
                {
                    return ds1.Tables["TEMPds1"];
                }
                else
                {
                    return null;
                }


            }
            catch
            {
                return null;
            }
            finally
            {
                sqlConn.Close();
            }
        }

        public void SET_TEXT()
        {
            textBox1.Text = "";
            textBox2.Text = "";
            textBox3.Text = "";
            textBox4.Text = "";
            textBox5.Text = "";
            textBox6.Text = "";
            textBox7.Text = "";
            textBox8.Text = "";
            textBox9.Text = "";

            // 清除 PictureBox 的图像
            pictureBox1.Image = null;
        }
        private void DisplayImageFromFolder(string NO)
        {
            string YYYY = NO.Substring(4,4);
            string folderPath = Path.Combine(Environment.CurrentDirectory, "Images", YYYY);
            string selectedImageFileName =null;
            // 檢查資料夾是否存在
            if (!Directory.Exists(folderPath))
            {
                MessageBox.Show("資料夾不存在。");
                return;
            }

            // 獲取資料夾中的所有圖片檔案
            string[] imageFiles = Directory.GetFiles(folderPath, "*.jpg"); // 只顯示 .jpg 檔案，您可以根據需要更改擴展名

            if (imageFiles.Length > 0)
            {
                // 在这里指定要显示的图像文件名
                selectedImageFileName = NO + ".jpg";

                string imagePath = Path.Combine(folderPath, selectedImageFileName);
                // 顯示圖片在 PictureBox 控制項上
                pictureBox1.Image = Image.FromFile(imagePath);
            }
            else
            {
                // 如果沒有圖片，清除 PictureBox
                pictureBox1.Image = null;
                MessageBox.Show("沒有找到圖片。");
            }
        }
        public void PACKAGEBOXS_ADD()
        {
            SqlConnection sqlConn = new SqlConnection();
            SqlCommand sqlComm = new SqlCommand();

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


                sqlConn.Close();
                sqlConn.Open();
                tran = sqlConn.BeginTransaction();

                sbSql.Clear();

                sbSql.AppendFormat(@" 
                                    
                                        
                                        "
                                        );


                cmd.Connection = sqlConn;
                cmd.CommandTimeout = 60;
                cmd.CommandText = sbSql.ToString();
                cmd.Transaction = tran;
                result = cmd.ExecuteNonQuery();

                if (result == 0)
                {
                    tran.Rollback();    //交易取消
                }
                else
                {
                    tran.Commit();      //執行交易  

                    MessageBox.Show("完成");

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
        private void button1_Click(object sender, EventArgs e)
        {
            Search_COPTG(dateTimePicker1.Value.ToString("yyyyMMdd"));
        }


        #endregion

       
    }
}

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
using AForge.Video;
using AForge.Video.DirectShow;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;
using System.Threading;
using System.IO.Ports;
using System.Threading;
using System.IO.Ports;



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

        string NO = null;
        string TG001TG002 = null;
        string TG001 = null;
        string TG002 = null;

        public FilterInfoCollection USB_Webcams = null;//FilterInfoCollection類別實體化
        public VideoCaptureDevice Cam;//攝像頭的初始化
        public VideoCaptureDevice Cam2;//攝像頭的初始化

        public Thread ReadSerialDataThread;
        public string readseroaldata;
        private SerialPort serialPortIn;
        public string CAL_TEXTBOX;
        public Report report1 { get; private set; }

        public FrmPACKAGEBOXS()
        {
            InitializeComponent();

            comboBox1load();
            comboBox2load();
            comboBox3load();
            comboBox4load();
            comboBox5load();
            comboBox6load();


            DataTable DT = SET_Btnconnect();
            if (DT != null && DT.Rows.Count >= 1)
            {
                comboBox6.Text = DT.Rows[0]["NAMES"].ToString();
            }
        }

        #region FUNCTION

        private void FrmPACKAGEBOXS_Load(object sender, EventArgs e)
        {           
            Btnconnect();
        }

        private void FrmPACKAGEBOXS_FormClosed(object sender, FormClosedEventArgs e)
        {
            Btndisconnect();
        }
        public void LoadComboBoxData(System.Windows.Forms.ComboBox comboBox, string query, string valueMember, string displayMember)
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
        public void comboBox4load()
        {
            LoadComboBoxData(comboBox4, "SELECT [ID],[NAMES] FROM [TKWAREHOUSE].[dbo].[TBPARAS] WHERE [KINDS]='PACKNAMES' GROUP BY [ID],[NAMES]  ", "NAMES", "NAMES");
        }
        public void comboBox5load()
        {
            LoadComboBoxData(comboBox5, "SELECT [ID],[NAMES] FROM [TKWAREHOUSE].[dbo].[TBPARAS] WHERE [KINDS]='REPORT1' GROUP BY [ID],[NAMES]  ", "NAMES", "NAMES");
        }

        public void comboBox6load()
        {
            LoadComboBoxData(comboBox6, "SELECT [ID],[NAMES] FROM [TKWAREHOUSE].[dbo].[TBPARAS] WHERE [KINDS]='PortNameSELECT' GROUP BY [ID],[NAMES]  ", "NAMES", "NAMES");
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
                                AND TG023 IN ('Y','N')
                                {0}
                                AND TG001 IN (SELECT [NAMES]  FROM [TKWAREHOUSE].[dbo].[TBPARAS] WHERE KINDS='COPTGTG001')
                                  ", sbSqlQuery1.ToString());
            sbSql.AppendFormat(@"  ");

            SEARCH(sbSql.ToString(), dataGridView1, SortedColumn, SortedModel);

        }

        private void dataGridView1_SelectionChanged(object sender, EventArgs e)
        {
            TG001TG002 = "";

            if (dataGridView1.CurrentRow != null)
            {
                int rowindex = dataGridView1.CurrentRow.Index;
                if (rowindex >= 0)
                {
                    DataGridViewRow row = dataGridView1.Rows[rowindex];
                    TG001TG002 = row.Cells["銷貨單"].Value.ToString()+ row.Cells["銷貨單號"].Value.ToString();
                    TG001 = row.Cells["銷貨單"].Value.ToString() ;
                    TG002 = row.Cells["銷貨單號"].Value.ToString();

                    DataTable dt = PACKAGEBOXS_FIND(TG001TG002);
                    if(dt!=null&&dt.Rows.Count>=1)
                    {
                        Search_PACKAGEBOXS(TG001TG002);
                    }
                    else
                    {                       
                        dataGridView2.DataSource = null;
                        SET_TEXT();
                        textBox1.Text = TG001TG002;
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
                                ,[PACKAGEBOXS].[NO]
                                ,A.[PHOTOS] AS '總重PHOTOS'
                                ,B.[PHOTOS] AS '箱重PHOTOS'
                                FROM [TKWAREHOUSE].[dbo].[PACKAGEBOXS]
                                LEFT JOIN  [TKWAREHOUSE].[dbo].[PACKAGEBOXSPHOTO] A ON A.NO=[PACKAGEBOXS].NO AND A.TYPES='總重'
                                LEFT JOIN  [TKWAREHOUSE].[dbo].[PACKAGEBOXSPHOTO] B ON B.NO=[PACKAGEBOXS].NO AND B.TYPES='箱重'
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
            byte[] retrievedImageBytes;
            byte[] retrievedImageBytes2;


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
                    comboBox4.Text = row.Cells["使用包材名稱/規格"].Value.ToString();
                    textBox8.Text = row.Cells["使用包材來源"].Value.ToString();
                    textBox9.Text = row.Cells["NO"].Value.ToString();

                    comboBox1.Text = row.Cells["商品總重量比值分類"].Value.ToString();
                    comboBox2.Text = row.Cells["規定比值"].Value.ToString();
                    comboBox3.Text = row.Cells["是否符合"].Value.ToString();

                    NO = row.Cells["NO"].Value.ToString();

                    try
                    {
                        retrievedImageBytes = (byte[])row.Cells["箱重PHOTOS"].Value;
                        retrievedImageBytes2 = (byte[])row.Cells["總重PHOTOS"].Value;
                        using (MemoryStream ms = new MemoryStream(retrievedImageBytes))
                        {
                            pictureBox1.Image = Image.FromStream(ms);
                        }
                        using (MemoryStream ms = new MemoryStream(retrievedImageBytes2))
                        {
                            pictureBox2.Image = Image.FromStream(ms);
                        }
                    }
                    catch
                    {

                    }
                   
                    //DisplayImageFromFolder(row.Cells["NO"].Value.ToString());
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

        public DataTable PACKAGEBOXS_FIND_MAX(string TG001TG002)
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
                                    ISNULL(COUNT(BOXNO),0) AS BOXNO
                                    FROM [TKWAREHOUSE].[dbo].[PACKAGEBOXS]
                                    WHERE TG001+TG002='{0}'
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
            //textBox7.Text = "";
            textBox8.Text = "";
            textBox9.Text = "";

            // 清除 PictureBox 的图像
            pictureBox1.Image = null;
            pictureBox2.Image = null; 
        }
        private void DisplayImageFromFolder(string NO)
        {
            if(!string.IsNullOrEmpty(NO))
            {
                string YYYY = NO.Substring(4, 4);
                string folderPath = Path.Combine(Environment.CurrentDirectory, "Images", YYYY);
                string selectedImageFileName = null;
                // 檢查資料夾是否存在
                if (!Directory.Exists(folderPath))
                {
                    MessageBox.Show("資料夾不存在。");
                    return;
                }

                // 獲取資料夾中的所有圖片檔案
                // 在这里指定要显示的图像文件名
                selectedImageFileName = NO + ".jpg";
                //string[] imageFiles = Directory.GetFiles(folderPath, selectedImageFileName); // 只顯示 .jpg 檔案，您可以根據需要更改擴展名

                string imagePath = Path.Combine(folderPath, selectedImageFileName);
                if (File.Exists(imagePath))
                {                    
                    // 顯示圖片在 PictureBox 控制項上
                    if (Image.FromFile(imagePath) != null)
                    {
                        Image img = GetCopyImage(imagePath);
                        pictureBox1.Image = img;
                        //img.Dispose();  // dispose the bitmap object

                        //System.Drawing.Image img = System.Drawing.Image.FromFile(imagePath);
                        //System.Drawing.Image bmp = new System.Drawing.Bitmap(img);
                        //img.Dispose();
                        //pictureBox1.Image = bmp;
                        //pictureBox1.Image = Image.FromFile(imagePath);



                    }

                }
                else
                {
                    // 如果沒有圖片，清除 PictureBox
                    pictureBox1.Image = null;
                    //MessageBox.Show("沒有找到圖片。");
                }
            }           
            else
            {
                // 如果沒有圖片，清除 PictureBox
                pictureBox1.Image = null;
                //MessageBox.Show("沒有找到圖片。");
            }
        }

        private Image GetCopyImage(string path)
        {
            using (Image im = Image.FromFile(path))
            {
                Bitmap bm = new Bitmap(im);
                return bm;
            }
        }

        public void PACKAGEBOXS_ADD(
                      string NO
                    , string TG001
                    , string TG002
                    , string BOXNO
                    , string ALLWEIGHTS
                    , string PACKWEIGHTS
                    , string PRODUCTWEIGHTS
                    , string PACKRATES
                    , string RATECLASS
                    , string CHECKRATES
                    , string ISVALIDS
                    , string PACKAGENAMES
                    , string PACKAGEFROM
            )
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
                                    INSERT INTO [TKWAREHOUSE].[dbo].[PACKAGEBOXS]
                                    (
                                    [NO]
                                    ,[TG001]
                                    ,[TG002]
                                    ,[BOXNO]
                                    ,[ALLWEIGHTS]
                                    ,[PACKWEIGHTS]
                                    ,[PRODUCTWEIGHTS]
                                    ,[PACKRATES]
                                    ,[RATECLASS]
                                    ,[CHECKRATES]
                                    ,[ISVALIDS]
                                    ,[PACKAGENAMES]
                                    ,[PACKAGEFROM]
                                    )
                                    VALUES
                                    (
                                    '{0}'
                                    ,'{1}'
                                    ,'{2}'
                                    ,'{3}'
                                    ,{4}
                                    ,{5}
                                    ,{6}
                                    ,'{7}'
                                    ,'{8}'
                                    ,'{9}'
                                    ,'{10}'
                                    ,'{11}'
                                    ,'{12}'
                                    )                                        
                                        "
                                    , NO
                                    , TG001
                                    , TG002
                                    , BOXNO
                                    , ALLWEIGHTS
                                    , PACKWEIGHTS
                                    , PRODUCTWEIGHTS
                                    , PACKRATES
                                    , RATECLASS
                                    , CHECKRATES
                                    , ISVALIDS
                                    , PACKAGENAMES
                                    , PACKAGEFROM
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


        public void PACKAGEBOXS_UPDATE(
                      string NO
                    , string TG001
                    , string TG002
                    , string BOXNO
                    , string ALLWEIGHTS
                    , string PACKWEIGHTS
                    , string PRODUCTWEIGHTS
                    , string PACKRATES
                    , string RATECLASS
                    , string CHECKRATES
                    , string ISVALIDS
                    , string PACKAGENAMES
                    , string PACKAGEFROM
            )
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
                                    UPDATE [TKWAREHOUSE].[dbo].[PACKAGEBOXS]
                                    SET
                                    
                                    [TG001]='{1}'
                                    ,[TG002]='{2}'
                                    ,[BOXNO]='{3}'
                                    ,[ALLWEIGHTS]={4}
                                    ,[PACKWEIGHTS]={5}
                                    ,[PRODUCTWEIGHTS]={6}
                                    ,[PACKRATES]='{7}'
                                    ,[RATECLASS]='{8}'
                                    ,[CHECKRATES]='{9}'
                                    ,[ISVALIDS]='{10}'
                                    ,[PACKAGENAMES]='{11}'
                                    ,[PACKAGEFROM]='{12}'
                                    
                                    WHERE [NO]='{0}'
                                                                     
                                        "
                                    , NO
                                    , TG001
                                    , TG002
                                    , BOXNO
                                    , ALLWEIGHTS
                                    , PACKWEIGHTS
                                    , PRODUCTWEIGHTS
                                    , PACKRATES
                                    , RATECLASS
                                    , CHECKRATES
                                    , ISVALIDS
                                    , PACKAGENAMES
                                    , PACKAGEFROM
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

        public void TPACKAGEBOXS_DELETE(string NO )
                        
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
                                    DELETE [TKWAREHOUSE].[dbo].[PACKAGEBOXS]
                                    WHERE NO='{0}'                         
                                        "
                                    , NO
                                   
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

        public void TAKE_OPEN()
        {

            USB_Webcams = new FilterInfoCollection(FilterCategory.VideoInputDevice);
        

            if (USB_Webcams.Count > 0)  // The quantity of WebCam must be more than 0.
            {
                
                Cam = new VideoCaptureDevice(USB_Webcams[0].MonikerString);
                // 取得視訊設備的所有可用解析度
                VideoCapabilities[] availableResolutions = Cam.VideoCapabilities;
                // 選擇所需的解析度，例如，選擇第一個可用的解析度
                if (availableResolutions.Length > 0)
                {
                    Cam.VideoResolution = availableResolutions[10];
                }

                Cam.NewFrame += Cam_NewFrame;//Press Tab  to   create
            }
            else
            {
                
                MessageBox.Show("No video input device is connected.");
            }
        }

        public void TAKE_CLOSE()
        {
            if (Cam != null)
            {
                if (Cam.IsRunning)  // When Form1 closes itself, WebCam must stop, too.
                {
                    Cam.Stop();   // WebCam stops capturing images.
                }
            }
        }

        void Cam_NewFrame(object sender, NewFrameEventArgs eventArgs)
        {
            // 設定 PictureBox 的大小和模式
            //pictureBox1.Size = new Size(Cam.VideoResolution.FrameSize.Width, Cam.VideoResolution.FrameSize.Height);
            pictureBox1.SizeMode = PictureBoxSizeMode.Zoom;
            //throw new NotImplementedException();
            pictureBox1.Image = (Bitmap)eventArgs.Frame.Clone();
        }

        //保存图片
        private delegate void SaveImage();
        private void SaveImageHH(string ImagePath)
        {
            if (this.pictureBox1.InvokeRequired)
            {
                SaveImage saveimage = delegate { this.pictureBox1.Image.Save(ImagePath); };
                this.pictureBox1.Invoke(saveimage);
            }
            else
            {
                this.pictureBox1.Image.Save(ImagePath);
            }

        }

        // 將 PictureBox 中的圖片轉換為位元組數組
        private byte[] ImageToByteArray(Image image)
        {
            using (MemoryStream ms = new MemoryStream())
            {
                image.Save(ms, System.Drawing.Imaging.ImageFormat.Jpeg); // 或者使用其他圖像格式
                return ms.ToArray();
            }
        }

        // 將位元組數組插入到資料庫的 BLOB 欄位中
        private void InsertImageIntoDatabase(string NO,string TYPES, string CTIMES, byte[] imageBytes)
        {
            SqlConnection sqlConn = new SqlConnection();
            SqlCommand sqlComm = new SqlCommand();
            SqlCommand cmd = new SqlCommand();

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
                                    INSERT INTO [TKWAREHOUSE].[dbo].[PACKAGEBOXSPHOTO]
                                    ([NO],[TYPES], [CTIMES], [PHOTOS])
                                    VALUES
                                    (@NO,@TYPES, @CTIMES, @PHOTOS)
                                    "
                                    );

                cmd.Parameters.AddWithValue("@NO", NO);
                cmd.Parameters.AddWithValue("@TYPES", TYPES);
                cmd.Parameters.AddWithValue("@CTIMES", CTIMES);
                cmd.Parameters.AddWithValue("@PHOTOS", imageBytes);

                cmd.Connection = sqlConn;
                cmd.CommandTimeout = 60;
                cmd.CommandText = sbSql.ToString();
                cmd.Transaction = tran;
                result = cmd.ExecuteNonQuery();

                if (result == 0)
                {
                    MessageBox.Show("圖片存儲 失敗");
                    tran.Rollback();    //交易取消
                }
                else
                {
                    tran.Commit();      //執行交易  

                    //MessageBox.Show("圖片已成功存儲到資料庫。");

                }

            }
            catch (Exception ex)
            {                
                MessageBox.Show(ex.ToString());
            }

            finally
            {
                sqlConn.Close();
            }
        }

        // 將 PictureBox 中的圖片存儲到資料庫
        private void SaveImageToDatabase(string NO)
        {

            // 替換為您的 PictureBox 控制項名稱
            Image image = pictureBox1.Image;

            if (image != null)
            {
                byte[] imageBytes = ImageToByteArray(image);
                InsertImageIntoDatabase(NO, "總重", DateTime.Now.ToString("yyyyMMdd HH:MM:ss"), imageBytes);

            }
            else
            {
                MessageBox.Show("pictureBox1是空的");
            }
        }


        private void DELETE_ImageIntoDatabase(string NO)
        {
            SqlConnection sqlConn = new SqlConnection();
            SqlCommand sqlComm = new SqlCommand();
            SqlCommand cmd = new SqlCommand();

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
                                    DELETE [TKWAREHOUSE].[dbo].[PACKAGEBOXSPHOTO]
                                    WHERE TYPES='總重'
                                    AND NO=@NO"
                                    );

                cmd.Parameters.AddWithValue("@NO", NO);
        
         

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

                    //MessageBox.Show("圖片已成功存儲到資料庫。");

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

        public void DEL_IMAGES(string ImagePath)
        {          
            //// 指定圖片的完整路徑，包括資料夾和檔案名稱
            string imagePaths = ImagePath;

            try
            {
                int maxRetryAttempts = 3;
                int retryDelayMilliseconds = 1000; // 1秒

                for (int i = 0; i < maxRetryAttempts; i++)
                {
                    try
                    {
                        File.Delete(imagePaths);
                        imagePaths = null; // 设置为 null，以释放资源

                        MessageBox.Show("完成-刪除照片 ");
                        break; // 如果删除成功，退出循环
                    }
                    catch (IOException ex)
                    {
                        if (i < maxRetryAttempts - 1)
                        {
                            // 如果删除失败，等待一段时间后重试
                            System.Threading.Thread.Sleep(retryDelayMilliseconds);
                        }
                        else
                        {
                            MessageBox.Show("失敗-刪除照片 請重開程式再刪除");
                            // 如果达到最大重试次数仍然无法删除，处理异常或显示错误消息
                            //MessageBox.Show("无法删除图像文件，因为它正在被其他进程使用。");
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                // 处理其他异常
            }
        }

        public void TAKE_OPEN2()
        {

            USB_Webcams = new FilterInfoCollection(FilterCategory.VideoInputDevice);


            if (USB_Webcams.Count > 0)  // The quantity of WebCam must be more than 0.
            {

                Cam2 = new VideoCaptureDevice(USB_Webcams[0].MonikerString);
                // 取得視訊設備的所有可用解析度
                VideoCapabilities[] availableResolutions = Cam2.VideoCapabilities;
                // 選擇所需的解析度，例如，選擇第一個可用的解析度
                if (availableResolutions.Length > 0)
                {
                    Cam2.VideoResolution = availableResolutions[10];
                }

                Cam2.NewFrame += Cam_NewFrame2;//Press Tab  to   create
            }
            else
            {

                MessageBox.Show("No video input device is connected.");
            }
        }

        public void TAKE_CLOSE2()
        {
            if (Cam2 != null)
            {
                if (Cam2.IsRunning)  // When Form1 closes itself, WebCam must stop, too.
                {
                    Cam2.Stop();   // WebCam stops capturing images.
                }
            }
        }

        void Cam_NewFrame2(object sender, NewFrameEventArgs eventArgs)
        {
            // 設定 PictureBox 的大小和模式
            //pictureBox1.Size = new Size(Cam.VideoResolution.FrameSize.Width, Cam.VideoResolution.FrameSize.Height);
            pictureBox2.SizeMode = PictureBoxSizeMode.Zoom;
            //throw new NotImplementedException();
            pictureBox2.Image = (Bitmap)eventArgs.Frame.Clone();
        }

        //保存图片
        private delegate void SaveImage2();
        private void SaveImageHH2(string ImagePath)
        {
            if (this.pictureBox2.InvokeRequired)
            {
                SaveImage saveimage = delegate { this.pictureBox2.Image.Save(ImagePath); };
                this.pictureBox2.Invoke(saveimage);
            }
            else
            {
                this.pictureBox2.Image.Save(ImagePath);
            }

        }

        // 將 PictureBox 中的圖片轉換為位元組數組
        private byte[] ImageToByteArray2(Image image)
        {
            using (MemoryStream ms = new MemoryStream())
            {
                image.Save(ms, System.Drawing.Imaging.ImageFormat.Jpeg); // 或者使用其他圖像格式
                return ms.ToArray();
            }
        }

        // 將位元組數組插入到資料庫的 BLOB 欄位中
        private void InsertImageIntoDatabase2(string NO,string TYPES, string CTIMES, byte[] imageBytes)
        {
            SqlConnection sqlConn = new SqlConnection();
            SqlCommand sqlComm = new SqlCommand();
            SqlCommand cmd = new SqlCommand();

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
                                    INSERT INTO [TKWAREHOUSE].[dbo].[PACKAGEBOXSPHOTO]
                                    ([NO],[TYPES], [CTIMES], [PHOTOS])
                                    VALUES
                                    (@NO,@TYPES, @CTIMES, @PHOTOS)
                                    "
                                    );

                cmd.Parameters.AddWithValue("@NO", NO);
                cmd.Parameters.AddWithValue("@TYPES", TYPES);
                cmd.Parameters.AddWithValue("@CTIMES", CTIMES);
                cmd.Parameters.AddWithValue("@PHOTOS", imageBytes);

                cmd.Connection = sqlConn;
                cmd.CommandTimeout = 60;
                cmd.CommandText = sbSql.ToString();
                cmd.Transaction = tran;
                result = cmd.ExecuteNonQuery();

                if (result == 0)
                {
                    MessageBox.Show("圖片存儲 失敗");
                    tran.Rollback();    //交易取消
                }
                else
                {
                    tran.Commit();      //執行交易  

                    //MessageBox.Show("圖片已成功存儲到資料庫。");

                }

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }

            finally
            {
                sqlConn.Close();
            }
        }

        // 將 PictureBox 中的圖片存儲到資料庫
        private void SaveImageToDatabase2(string NO)
        {

            // 替換為您的 PictureBox 控制項名稱
            Image image = pictureBox2.Image;

            if (image != null)
            {
                byte[] imageBytes = ImageToByteArray2(image);
                InsertImageIntoDatabase2(NO,"箱重", DateTime.Now.ToString("yyyyMMdd HH:MM:ss"), imageBytes);

            }
            else
            {
                MessageBox.Show("pictureBox1是空的");
            }
        }


        private void DELETE_ImageIntoDatabase2(string NO)
        {
            SqlConnection sqlConn = new SqlConnection();
            SqlCommand sqlComm = new SqlCommand();
            SqlCommand cmd = new SqlCommand();

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
                                    DELETE [TKWAREHOUSE].[dbo].[PACKAGEBOXSPHOTO]
                                    WHERE TYPES='箱重' AND NO=@NO"
                                    );

                cmd.Parameters.AddWithValue("@NO", NO);



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

                    //MessageBox.Show("圖片已成功存儲到資料庫。");

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

        public void DEL_IMAGES2(string ImagePath)
        {
            //// 指定圖片的完整路徑，包括資料夾和檔案名稱
            string imagePaths = ImagePath;

            try
            {
                int maxRetryAttempts = 3;
                int retryDelayMilliseconds = 1000; // 1秒

                for (int i = 0; i < maxRetryAttempts; i++)
                {
                    try
                    {
                        File.Delete(imagePaths);
                        imagePaths = null; // 设置为 null，以释放资源

                        MessageBox.Show("完成-刪除照片 ");
                        break; // 如果删除成功，退出循环
                    }
                    catch (IOException ex)
                    {
                        if (i < maxRetryAttempts - 1)
                        {
                            // 如果删除失败，等待一段时间后重试
                            System.Threading.Thread.Sleep(retryDelayMilliseconds);
                        }
                        else
                        {
                            MessageBox.Show("失敗-刪除照片 請重開程式再刪除");
                            // 如果达到最大重试次数仍然无法删除，处理异常或显示错误消息
                            //MessageBox.Show("无法删除图像文件，因为它正在被其他进程使用。");
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                // 处理其他异常
            }
        }



        private void textBox3_TextChanged(object sender, EventArgs e)
        {
            string input = textBox3.Text;
            double ALLWEIGHT;
            float result;

            if(!string.IsNullOrEmpty(input))
            {
                if (float.TryParse(input, out result))
                {
                    ALLWEIGHT = Convert.ToDouble(input);

                    if(ALLWEIGHT<0.25)
                    {
                        comboBox4.Text = "回收箱小";
                    }
                    else if(ALLWEIGHT>=0.25 && ALLWEIGHT<1)
                    {
                        comboBox4.Text = "回收箱小";
                    }
                    else if (ALLWEIGHT >= 1 && ALLWEIGHT < 3)
                    {
                        comboBox4.Text = "回收箱中";
                    }
                    else if (ALLWEIGHT >= 3 )
                    {
                        comboBox4.Text = "回收箱大";
                    }

                    CAL_RATES();

                    DataTable dt = SET_RATECLASS(input);
                    if (dt != null && dt.Rows.Count >= 1)
                    {
                        comboBox1.Text = dt.Rows[0]["NAMES"].ToString();
                    }   
                    
                                   
                }
                else
                {
                    MessageBox.Show("重量不是數字格式");
                }

                
            }
        }

        private void textBox4_TextChanged(object sender, EventArgs e)
        {
            string input = textBox4.Text;
            float result;

            if (!string.IsNullOrEmpty(input))
            {
                if (float.TryParse(input, out result))
                {
                    CAL_ALLWEIGHTS();
                    CAL_RATES();
                }
                else
                {
                    MessageBox.Show("重量不是數字格式");
                }
            }
        }

        private void textBox5_TextChanged(object sender, EventArgs e)
        {
            string input = textBox5.Text; 
            float result;
            if (!string.IsNullOrEmpty(input))
            {
                if (float.TryParse(input, out result))
                {
                   
                    CAL_ALLWEIGHTS();
                }
                else
                {
                    MessageBox.Show("重量不是數字格式");
                }
            }

        }

        public void CAL_ALLWEIGHTS()
        {
            float result;
            string input1 = textBox4.Text;
            string input2 = textBox3.Text;
            float PACKWEIGHTS = 0;
            float ALLPRODUCTWEIGHTS = 0;
            if (!string.IsNullOrEmpty(input1) && !string.IsNullOrEmpty(input2))
            {
                if (float.TryParse(input1, out result) && float.TryParse(input2, out result))
                {
                    PACKWEIGHTS = float.Parse(input1);
                    ALLPRODUCTWEIGHTS = float.Parse(input2);
                    textBox5.Text = (ALLPRODUCTWEIGHTS - PACKWEIGHTS).ToString("0.000");
                }
                else
                {
                    MessageBox.Show("重量不是數字格式");
                }
            }
        }
        public void CAL_RATES()
        {
            float result;
            string input1 = textBox3.Text;
            string input2 = textBox4.Text;
            float PACKWEIGHTS = 0;
            float ALLWEIGHTS = 0;           
            float rates = 0;

            if(!string.IsNullOrEmpty(input1)&& !string.IsNullOrEmpty(input2))
            {                
                if (float.TryParse(input1, out result)&& float.TryParse(input2, out result))
                {
                    ALLWEIGHTS = float.Parse(input1);
                    PACKWEIGHTS = float.Parse(input2);

                    decimal difference = (decimal)(ALLWEIGHTS - PACKWEIGHTS);
                    textBox5.Text = difference.ToString("0.00"); // 保留小數第二位

                    rates = (PACKWEIGHTS * 100 / ALLWEIGHTS * 100) ;
                    textBox6.Text = rates.ToString("0.00") + "%";
                }
                else
                { 
                    MessageBox.Show("重量不是數字格式");
                }
            }          
        }

        public DataTable SET_RATECLASS(string ALLWEIGHTS)
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

                //用總重去比對重量的條件
                //({0}-CONVERT(float,[KEYS])>0)，從小排到大
                sbSql.AppendFormat(@"                                      
                                   SELECT [KINDS]
                                    ,[NAMES]
                                    ,[KEYS]
                                    ,[KEYS2]
                                    ,({0}-CONVERT(float,[KEYS])) AS CONDITIONS
                                    FROM [TKWAREHOUSE].[dbo].[TBPARAS]
                                    WHERE [KINDS]='RATECLASS'
                                    AND ({0}-CONVERT(float,[KEYS])>0)
                                    ORDER BY  ({0}-CONVERT(float,[KEYS])) ASC
                                    ", ALLWEIGHTS);




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

        private void comboBox1_SelectedValueChanged(object sender, EventArgs e)
        {
            string input = comboBox1.Text.ToString();

            if(!string.IsNullOrEmpty(input))
            {
                DataTable dt = SET_CHECKRATES(input);

                if (dt!=null&&dt.Rows.Count>=1)
                {
                    comboBox2.Text = dt.Rows[0]["NAMES"].ToString();
                }
            }
        }
        public DataTable SET_CHECKRATES(string KEYS)
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

                //用總重去比對重量的條件
                //({0}-CONVERT(float,[KEYS])>0)，從小排到大
                sbSql.AppendFormat(@"                                      
                                   SELECT [KINDS]
                                    ,[NAMES]
                                    ,[KEYS]
                                    ,[KEYS2]
                                    FROM [TKWAREHOUSE].[dbo].[TBPARAS]
                                    WHERE [KINDS]='CHECKRATES'
                                    AND [KEYS]='{0}'
                                    ", KEYS);




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
        private void textBox6_TextChanged(object sender, EventArgs e)
        {
            comboBox3.Text = "不適用";
            if(!string.IsNullOrEmpty(comboBox1.Text)&&!string.IsNullOrEmpty(textBox6.Text))
            {
                string input1 = comboBox1.Text;
                string input2 = textBox6.Text.Replace("%","");
                double ALLWEIGHT = Convert.ToDouble(textBox3.Text);

                if (ALLWEIGHT < 0.25)
                {
                    comboBox3.Text = "不適用";
                }
                DataTable dt = SET_ISVALIDS(input1, input2);
                if (dt != null && dt.Rows.Count >= 1)
                {
                    comboBox3.Text = "符合";
                }
                else
                {
                    comboBox3.Text = "不符合";
                }
            }
            
        }

        public DataTable SET_ISVALIDS(string KEYS, string KEYS2)
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

                //用總重去比對重量的條件
                //({0}-CONVERT(float,[KEYS])>0)，從小排到大
                sbSql.AppendFormat(@"                                      
                                    SELECT [KINDS]
                                    ,[NAMES]
                                    ,[KEYS]
                                    ,[KEYS2]
                                    FROM [TKWAREHOUSE].[dbo].[TBPARAS]
                                    WHERE [KINDS]='ISVALIDS'
                                    AND [KEYS]='{0}'
                                    AND CONVERT(decimal,[KEYS2])>{1}
                                    ", KEYS, KEYS2);




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

        public void Btnconnect()
        {
            serialPortIn = new SerialPort();


            serialPortIn.PortName = comboBox6.Text.ToString();
            serialPortIn.BaudRate = 9600;
            serialPortIn.Parity = Parity.None;
            serialPortIn.DataBits = 8;
            serialPortIn.StopBits = StopBits.One;

           

            if (!serialPortIn.IsOpen)
            {
                try
                {
                    //serialPortIn.PortName = txtportname.Text;
                    //serialPortIn.BaudRate = int.Parse(txtbaudrate.Text);
                    //serialPortIn.Parity = (Parity)Enum.Parse(typeof(Parity), txtparity.Text);
                    //serialPortIn.DataBits = int.Parse(txtdatabits.Text);
                    //serialPortIn.StopBits = (StopBits)Enum.Parse(typeof(StopBits), txtstopbits.Text);
                    //serialPortIn.Open();

                  
                    serialPortIn.Open();

                }
                catch (Exception ee)
                {

                    MessageBox.Show(@"ERROR:" + ee);
                }


            }


            if (serialPortIn.IsOpen)
            {
                ReadSerialData();
               
            }

        }

        private void ReadSerialData()
        {
            ReadSerialDataThread = new Thread(ReadSerial);
            ReadSerialDataThread.Start();
        }

        private void ReadSerial()
        {
            while (serialPortIn.IsOpen)
            {
                try
                {
                    readseroaldata = serialPortIn.ReadLine();
                    ShowSerialData(readseroaldata);
                }
                catch (Exception)
                {


                }


                Thread.Sleep(20);
            }
        }

        public delegate void ShowSerialDatadelegate(string r);
        private void ShowSerialData(string s)
        {
            DateTime now = DateTime.Now;
            string pattern = @"[-+]?\d*\.?\d+";
            string datacon = "";

            string ymdhms = "";
            string cross = "";

            // 獲取當前年份
            int year = now.Year;

            // 獲取當前月份
            int month = now.Month;


            // 獲取當前日期
            int day = now.Day;

            // 獲取當前小時
            int hour = now.Hour;

            // 獲取當前分鐘
            int minute = now.Minute;

            // 獲取當前秒數1056.68kg


            int second = now.Second;


            ymdhms = Convert.ToString(year) + "/" + Convert.ToString(month) + "/" + Convert.ToString(day) + " " + Convert.ToString(hour) + ":" + Convert.ToString(minute) + ":" + Convert.ToString(second) + " ";


            if (textBoxCAL.InvokeRequired)
            {
                ShowSerialDatadelegate SSDD = ShowSerialData;
                Invoke(SSDD, s);
            }
            else
            {
                MatchCollection matches = Regex.Matches(s, pattern);
                foreach (Match match in matches)
                {
                    datacon += match.Value;
                }

                textBoxCAL.Text = datacon;

                //string finaldatas = ymdhms + datacon + txtUnits.Text;



                //string prev = Clipboard.GetText();
                //txtreaddata.AppendText(finaldatas.Substring(((int)numericUpDown1.Value)));

                ////SendKeys.SendWait(datacon + txtUnits.Text + Environment.NewLine);

                //Clipboard.SetText(finaldatas.Substring(((int)numericUpDown1.Value)));
                //SendKeys.Send("^v");
                //Clipboard.SetText(prev);

                //txtreaddata.Text += "\n";
                //SendKeys.SendWait("{ENTER}");

            }
        }
        private async void Btndisconnect_Click(object sender, EventArgs e)
        {
            if (serialPortIn.IsOpen)
            {
                serialPortIn.Close();
                serialPortIn.Dispose();
                
                Thread.Sleep(20);

            }

        }

        private async void Btndisconnect()
        {
            if (serialPortIn.IsOpen)
            {
                serialPortIn.Close();
                serialPortIn.Dispose();

                Thread.Sleep(20);

            }
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            if(comboBox1.Text.Equals("0.25公斤~1公斤"))
            {

            }
            else if (comboBox1.Text.Equals("1公斤~3公斤"))
            {

            }
            else if (comboBox1.Text.Equals("3公斤(KG)以上"))
            {

            }
        
        }
        public void SETFASTREPORT(string SDAYE,string EDAYS,string REPORTS)
        {
            string SQL = ""; 
            report1 = new Report();

            if(REPORTS.Equals("現場空重比值明細秤重"))
            {
                report1.Load(@"REPORT\網購包材減量應填表單-現場空重比值明細秤重.frx");

                SQL = SETFASETSQL(SDAYE,EDAYS);
              
            }
            else if (REPORTS.Equals("銷貨資料")) 
            {
                report1.Load(@"REPORT\網購包材減量應填表單-銷貨資料.frx");

                SQL = SETFASETSQL2(SDAYE,EDAYS);
               
            }


            //20210902密
            Class1 TKID = new Class1();//用new 建立類別實體
            SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

            //資料庫使用者密碼解密
            sqlsb.Password = TKID.Decryption(sqlsb.Password);
            sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

            String connectionString;
            sqlConn = new SqlConnection(sqlsb.ConnectionString);

            report1.Dictionary.Connections[0].ConnectionString = sqlsb.ConnectionString;

            TableDataSource Table = report1.GetDataSource("Table") as TableDataSource;
            Table.SelectCommand = SQL.ToString();

            report1.Preview = previewControl1;
            report1.Show();

        }

        public string SETFASETSQL(string SDAYE, string EDAYS)
        {           
            StringBuilder FASTSQL = new StringBuilder();
            StringBuilder STRQUERY = new StringBuilder();

            FASTSQL.AppendFormat(@"   
                                SELECT *
                                FROM 
                                (
                                SELECT 
                                ( CASE WHEN ISNULL(SUBSTRING(TG029,3,6),'')<>'' THEN  '20'+SUBSTRING(TG029,3,6) ELSE '' END )AS '訂單日期'
                                ,TG029 AS '訂單編號'
                                ,COPTG.TG001  AS '銷貨單別'
                                ,COPTG.TG002 AS '銷貨單號'
                                ,TG003 AS '銷貨日'
                                ,TG020 AS '購物車編號'
                                ,UDF02 AS 'UDF02'
                                ,[PACKAGEBOXS].[NO] AS '編號'
                                ,[BOXNO] AS '箱號'
                                ,[ALLWEIGHTS] AS '秤總重(A+B)'
                                ,[PACKWEIGHTS] AS '(非原箱)網購包材重量(KG)A'
                                ,[PRODUCTWEIGHTS] AS '(非原箱)商品總重量(KG)B'
                                ,[PACKRATES] AS '實際比值'
                                ,[RATECLASS] AS '商品總重量比值分類'
                                ,[CHECKRATES] AS '規定比值'
                                ,[ISVALIDS] AS '是否符合'
                                ,[PACKAGENAMES] AS '使用包材名稱/規格'
                                ,[PACKAGEFROM] AS '使用包材來源'
                                ,A.[CTIMES] AS '總重照片時間'
                                ,B.[CTIMES] AS '毛重照片時間'
                                ,A.[PHOTOS] AS '總重PHOTOS'
                                ,B.[PHOTOS] AS '箱重PHOTOS'
                                FROM [TK].dbo.COPTG
                                LEFT JOIN [TKWAREHOUSE].[dbo].[PACKAGEBOXS] ON [PACKAGEBOXS].TG001=COPTG.TG001 AND [PACKAGEBOXS].TG002=COPTG.TG002
                                LEFT JOIN  [TKWAREHOUSE].[dbo].[PACKAGEBOXSPHOTO] A ON A.NO=[PACKAGEBOXS].NO AND A.TYPES='總重'
                                LEFT JOIN  [TKWAREHOUSE].[dbo].[PACKAGEBOXSPHOTO] B ON B.NO=[PACKAGEBOXS].NO AND B.TYPES='箱重'
                                WHERE TG023='Y'
                                AND COPTG.TG001 IN ('A233')
                                AND TG003>='{0}' AND TG003<='{1}'
                                AND ISNULL(TG029,'')<>''
                                UNION ALL
                                SELECT 
                                '' AS '訂單日期'
                                ,TG029 AS '訂單編號'
                                ,COPTG.TG001  AS '銷貨單別'
                                ,COPTG.TG002 AS '銷貨單號'
                                ,TG003 AS '銷貨日'
                                ,TG020 AS '購物車編號'
                                ,UDF02 AS 'UDF02'
                                ,[PACKAGEBOXS].[NO] AS '編號'
                                ,[BOXNO] AS '箱號'
                                ,[ALLWEIGHTS] AS '秤總重(A+B)'
                                ,[PACKWEIGHTS] AS '(非原箱)網購包材重量(KG)A'
                                ,[PRODUCTWEIGHTS] AS '(非原箱)商品總重量(KG)B'
                                ,[PACKRATES] AS '實際比值'
                                ,[RATECLASS] AS '商品總重量比值分類'
                                ,[CHECKRATES] AS '規定比值'
                                ,[ISVALIDS] AS '是否符合'
                                ,[PACKAGENAMES] AS '使用包材名稱/規格'
                                ,[PACKAGEFROM] AS '使用包材來源'
                                ,A.[CTIMES] AS '總重照片時間'
                                ,B.[CTIMES] AS '毛重照片時間'
                                ,A.[PHOTOS] AS '總重PHOTOS'
                                ,B.[PHOTOS] AS '箱重PHOTOS'
                                FROM [TK].dbo.COPTG
                                LEFT JOIN [TKWAREHOUSE].[dbo].[PACKAGEBOXS] ON [PACKAGEBOXS].TG001=COPTG.TG001 AND [PACKAGEBOXS].TG002=COPTG.TG002
                                LEFT JOIN  [TKWAREHOUSE].[dbo].[PACKAGEBOXSPHOTO] A ON A.NO=[PACKAGEBOXS].NO AND A.TYPES='總重'
                                LEFT JOIN  [TKWAREHOUSE].[dbo].[PACKAGEBOXSPHOTO] B ON B.NO=[PACKAGEBOXS].NO AND B.TYPES='箱重'
                                WHERE TG023='Y'
                                AND COPTG.TG001 IN ('A23A')
                                AND TG004 IN ('A209400300')
                                AND TG003>='{0}' AND TG003<='{1}'
      
                                ) AS TEMP
                                ORDER BY 銷貨單別,銷貨單號
                                    ", SDAYE, EDAYS);



            return FASTSQL.ToString();

        }

        public string SETFASETSQL2(string SDAYE, string EDAYS)
        {
            StringBuilder FASTSQL = new StringBuilder();
            StringBuilder STRQUERY = new StringBuilder();

            FASTSQL.AppendFormat(@"                                  
                                SELECT 訂單日期,訂單號碼
                                ,TG001 AS '銷貨單別'
                                ,TG002 AS '銷貨單號'
                                ,ISNULL((SELECT TOP 1 TA016 FROM [TK].dbo.ACRTA WHERE TA015=發票號碼),'') AS 發票日期
                                ,ISNULL(發票號碼,'') AS '發票號碼'
                                ,品號
                                ,品名
                                ,銷貨數量
                                ,銷貨含稅金額
                                ,ISNULL((SELECT TOP 1 (TA017+TA018) FROM [TK].dbo.ACRTA WHERE TA015=發票號碼),0) AS 發票金額
                                FROM
                                (
                                SELECT ( CASE WHEN ISNULL(SUBSTRING(TG029,3,6),'')<>'' THEN  '20'+SUBSTRING(TG029,3,6) ELSE '' END )AS '訂單日期',TG029 AS 訂單號碼
                                ,(SELECT TOP 1 TA015 FROM [TK].dbo.ACRTA,[TK].dbo.ACRTB WHERE TA001=TB001 AND TA002=TB002 AND TB005+TB006=TG001+TG002) AS 發票號碼
                                ,TH004 AS 品號
                                ,TH005 AS 品名
                                ,(TH008+TH024) AS 銷貨數量
                                ,(TH037+TH038) AS 銷貨含稅金額
                                ,TG001,TG002,TG003,TG029
                                FROM [TK].dbo.COPTG,[TK].dbo.COPTH
                                WHERE 1=1
                                AND TG001=TH001 AND TG002=TH002
                                AND TG023='Y'
                                AND TG001 IN ('A233')
                                AND TG003>='{0}' AND TG003<='{1}'
                                AND ISNULL(TG029,'')<>''

                                UNION ALL
                                SELECT ( CASE WHEN ISNULL(SUBSTRING(TG029,3,6),'')<>'' THEN  '20'+SUBSTRING(TG029,3,6) ELSE '' END )AS '訂單日期',TG029 AS 訂單號碼
                                ,(SELECT TOP 1 TA015 FROM [TK].dbo.ACRTA,[TK].dbo.ACRTB WHERE TA001=TB001 AND TA002=TB002 AND TB005+TB006=TG001+TG002) AS 發票號碼
                                ,TH004 AS 品號
                                ,TH005 AS 品名
                                ,(TH008+TH024) AS 銷貨數量
                                ,(TH037+TH038) AS 銷貨含稅金額
                                ,TG001,TG002,TG003,TG029
                                FROM [TK].dbo.COPTG,[TK].dbo.COPTH
                                WHERE 1=1
                                AND TG001=TH001 AND TG002=TH002
                                AND TG023='Y'
                                AND TG001 IN ('A23A')
                                AND TG002 LIKE '2023%'
                                AND TG003>='{0}' AND TG003<='{1}'
                                AND TG004 IN ('A209400300')
                                ) AS TMEP 
                                ORDER BY TG001,TG002,訂單日期
                                    ", SDAYE, EDAYS);



            return FASTSQL.ToString();

        }
        public DataTable SET_Btnconnect()
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

                //用總重去比對重量的條件
                //({0}-CONVERT(float,[KEYS])>0)，從小排到大
                sbSql.AppendFormat(@"                                      
                                   SELECT 
                                    [ID]
                                    ,[KINDS]
                                    ,[NAMES]
                                    ,[KEYS]
                                    ,[KEYS2]
                                    FROM [TKWAREHOUSE].[dbo].[TBPARAS]
                                    WHERE [KINDS]='PortName'
                                    ");




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
        #endregion


        #region BUTTON
        private void button1_Click(object sender, EventArgs e)
        {
            Search_COPTG(dateTimePicker1.Value.ToString("yyyyMMdd"));
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (!string.IsNullOrEmpty(TG001TG002))
            {
                DataTable dt= PACKAGEBOXS_FIND_MAX(TG001TG002);
                if (dt != null&& dt.Rows.Count>=1)
                {
                    textBox2.Text = (Convert.ToInt32(dt.Rows[0]["BOXNO"].ToString()) + 1).ToString();
                }
                else
                {
                    textBox2.Text = "1";
                }

                string NO = TG001 + TG002 + "-" + textBox2.Text;
                TG001 = TG001;
                TG002 = TG002;
                string BOXNO = textBox2.Text;
                string ALLWEIGHTS = textBox3.Text;
                string PACKWEIGHTS = textBox4.Text;
                string PRODUCTWEIGHTS = textBox5.Text;
                string PACKRATES = textBox6.Text;
                string RATECLASS = comboBox1.Text.ToString();
                string CHECKRATES = comboBox2.Text.ToString();
                string ISVALIDS = comboBox3.Text.ToString();
                string PACKAGENAMES = comboBox4.Text;
                string PACKAGEFROM = textBox8.Text;

                PACKAGEBOXS_ADD(
                    NO
                    , TG001
                    , TG002
                    , BOXNO
                    , ALLWEIGHTS
                    , PACKWEIGHTS
                    , PRODUCTWEIGHTS
                    , PACKRATES
                    , RATECLASS
                    , CHECKRATES
                    , ISVALIDS
                    , PACKAGENAMES
                    , PACKAGEFROM
                    );

                Search_PACKAGEBOXS(TG001TG002);
            }


            
        }

        private void button3_Click(object sender, EventArgs e)
        {
            string NO = textBox9.Text;
            string BOXNO = textBox2.Text;
            string ALLWEIGHTS = textBox3.Text;
            string PACKWEIGHTS = textBox4.Text;
            string PRODUCTWEIGHTS = textBox5.Text;
            string PACKRATES = textBox6.Text;
            string RATECLASS = comboBox1.Text.ToString();
            string CHECKRATES = comboBox2.Text.ToString();
            string ISVALIDS = comboBox3.Text.ToString();
            string PACKAGENAMES = comboBox4.Text;
            string PACKAGEFROM = textBox8.Text;

            PACKAGEBOXS_UPDATE(
                NO
                , TG001
                , TG002
                , BOXNO
                , ALLWEIGHTS
                , PACKWEIGHTS
                , PRODUCTWEIGHTS
                , PACKRATES
                , RATECLASS
                , CHECKRATES
                , ISVALIDS
                , PACKAGENAMES
                , PACKAGEFROM
                );

            Search_PACKAGEBOXS(TG001TG002);
        }

        private void button4_Click(object sender, EventArgs e)
        {
            // 顯示確認對話框
            DialogResult result = MessageBox.Show("確定要執行此操作嗎？", "確認", MessageBoxButtons.OKCancel);

            // 檢查使用者是否按下了確定按鈕
            if (result == DialogResult.OK)
            {
                // 確認後執行的動作
                // TODO: 在這裡執行您的程式碼
                // 例如：
                string NO = textBox9.Text;
                if (!string.IsNullOrEmpty(NO))
                {
                    TPACKAGEBOXS_DELETE(NO);
                    Search_PACKAGEBOXS(TG001TG002);
                }


            }
        }
        private void button5_Click(object sender, EventArgs e)
        {
            NO = textBox9.Text;

            if (!string.IsNullOrEmpty(NO))
            {
                TAKE_OPEN();
                try
                {
                    Cam.Start();   // WebCam starts capturing images.     
                }
                catch { }
            }
            else
            {
                MessageBox.Show("沒有對應 箱號，不能開啟相機");
            }


        }

        private void button6_Click(object sender, EventArgs e)
        {
            NO = textBox9.Text;
            if (!string.IsNullOrEmpty(NO))
            {
                //string imagePath = System.Environment.CurrentDirectory;
                string imagePath = Path.Combine(Environment.CurrentDirectory, "Images", DateTime.Now.ToString("yyyy"));
                string imagePathNames = imagePath + "\\" + NO + "-總重.jpg";
                if (!Directory.Exists(imagePath))
                {
                    Directory.CreateDirectory(imagePath);
                }

                SaveImageToDatabase(NO);
                SaveImageHH(imagePathNames);
           

                TAKE_CLOSE();
                try
                {
                    Cam.Stop();  // WebCam stops capturing images.
                }
                catch { }

                MessageBox.Show("拍照完成");
            }
            
        }
        private void button7_Click(object sender, EventArgs e)
        {
            NO = textBox9.Text;
            if (!string.IsNullOrEmpty(NO))
            {
                //DisplayImageFromFolder("");
                //pictureBox1.Image = null;

                pictureBox1.Image.Dispose();
                pictureBox1.Image = null;
                pictureBox1.ImageLocation = null;

                string imagePath = Path.Combine(Environment.CurrentDirectory, "Images", DateTime.Now.ToString("yyyy"));
                string imagePathNames = imagePath + "\\" + NO + "-總重.jpg";

                if (File.Exists(imagePathNames))
                {
                    DELETE_ImageIntoDatabase(NO);
                    DEL_IMAGES(imagePathNames);

                   
                }               
                
            }

            //if (!string.IsNullOrEmpty(NO))
            //{
            //    TAKE_OPEN();
            //    try
            //    {
            //        Cam.Start();   // WebCam starts capturing images.     
            //    }
            //    catch { }
            //}
            //else
            //{
            //    MessageBox.Show("沒有對應 箱號，不能開啟相機");
            //}


        }



        private void button8_Click(object sender, EventArgs e)
        {
            int MAXTRY = 1;
            //Btndisconnect();
            //// 等待  秒
            ////Thread.Sleep(1000);
            //Btnconnect();
            //// 等待  秒
            ////Thread.Sleep(1000);            

            if (!string.IsNullOrEmpty(textBoxCAL.Text))
            {
                float result;
                if (float.TryParse(textBoxCAL.Text, out result))
                {
                    textBox3.Text = textBoxCAL.Text;

                    while (MAXTRY <= 10 && !textBoxCAL.Text.Equals(textBox3.Text))
                    {
                        textBox3.Text = textBoxCAL.Text;

                        MAXTRY = MAXTRY + 1;
                        Thread.Sleep(10);
                    }

                    //Btndisconnect();
                }
                else
                {
                    // textBoxCAL.Text 不是有效的浮點數
                }
            }

        }
        private void button9_Click(object sender, EventArgs e)
        {
            int MAXTRY = 1;
            //Btndisconnect();
            //// 等待  秒
            ////Thread.Sleep(1000);
            //Btnconnect();
            //// 等待  秒
            ////Thread.Sleep(1000);

            if (!string.IsNullOrEmpty(textBoxCAL.Text))
            {
                float result;
                if (float.TryParse(textBoxCAL.Text, out result))
                {
                    textBox4.Text = textBoxCAL.Text;

                    while (MAXTRY <= 10 && !textBoxCAL.Text.Equals(textBox4.Text))
                    {
                        textBox4.Text = textBoxCAL.Text;

                        MAXTRY = MAXTRY +1;
                        Thread.Sleep(10);
                    }

                    //Btndisconnect();
                }
                else
                {
                    // textBoxCAL.Text 不是有效的浮點數
                }
            }        
        }


        private void button10_Click(object sender, EventArgs e)
        {
            SETFASTREPORT(dateTimePicker2.Value.ToString("yyyyMMdd"), dateTimePicker3.Value.ToString("yyyyMMdd"), comboBox5.Text);
        }

        private void button11_Click(object sender, EventArgs e)
        {
            Btndisconnect();
            Btnconnect();
        }
        private void button12_Click(object sender, EventArgs e)
        {
            SET_TEXT();
            textBox1.Text = TG001TG002;
        }

        private void button13_Click(object sender, EventArgs e)
        {
            NO = textBox9.Text;
            if (!string.IsNullOrEmpty(NO))
            {
                //string imagePath = System.Environment.CurrentDirectory;
                string imagePath = Path.Combine(Environment.CurrentDirectory, "Images", DateTime.Now.ToString("yyyy"));
                string imagePathNames = imagePath + "\\" + NO + "-箱重.jpg";
                if (!Directory.Exists(imagePath))
                {
                    Directory.CreateDirectory(imagePath);
                }

                SaveImageToDatabase2(NO);
                SaveImageHH2(imagePathNames);


                TAKE_CLOSE2();
                try
                {
                    Cam2.Stop();  // WebCam stops capturing images.
                }
                catch { }

                MessageBox.Show("拍照完成");
            }
        }

        private void button14_Click(object sender, EventArgs e)
        {
            NO = textBox9.Text;
            if (!string.IsNullOrEmpty(NO))
            {
                //DisplayImageFromFolder("");
                //pictureBox1.Image = null;

                pictureBox2.Image.Dispose();
                pictureBox2.Image = null;
                pictureBox2.ImageLocation = null;

                string imagePath = Path.Combine(Environment.CurrentDirectory, "Images", DateTime.Now.ToString("yyyy"));
                string imagePathNames = imagePath + "\\" + NO + "-箱重.jpg";

                if (File.Exists(imagePathNames))
                {
                    DELETE_ImageIntoDatabase2(NO);
                    DEL_IMAGES2(imagePathNames);


                }
            }
        }
        private void button15_Click(object sender, EventArgs e)
        {
            NO = textBox9.Text;

            if (!string.IsNullOrEmpty(NO))
            {
                TAKE_OPEN2();
                try
                {
                    Cam2.Start();   // WebCam starts capturing images.     
                }
                catch { }
            }
            else
            {
                MessageBox.Show("沒有對應 箱號，不能開啟相機");
            }

        }

        #endregion


    }
}

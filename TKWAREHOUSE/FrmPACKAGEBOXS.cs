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

                    NO = row.Cells["NO"].Value.ToString();
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
            textBox7.Text = "";
            textBox8.Text = "";
            textBox9.Text = "";

            // 清除 PictureBox 的图像
            pictureBox1.Image = null;
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
                                    
                                    ,[TG001]='{1}'
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
                                    )
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
                button1.Enabled = true;
                Cam = new VideoCaptureDevice(USB_Webcams[0].MonikerString);

                Cam.NewFrame += Cam_NewFrame;//Press Tab  to   create
            }
            else
            {
                button1.Enabled = false;
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
        private void InsertImageIntoDatabase(string NO, string CTIMES, byte[] imageBytes)
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
                                    INSERT INTO [TKWAREHOUSE].[dbo].[PACKAGEBOXSPHOTO]
                                    ([NO], [CTIMES], [PHOTOS])
                                    VALUES
                                    (@NO, @CTIMES, @PHOTOS)
                                    "
                                    );

                cmd.Parameters.AddWithValue("@NO", NO);
                cmd.Parameters.AddWithValue("@CTIMES", CTIMES);
                cmd.Parameters.AddWithValue("@PHOTOS", imageBytes);

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

        // 將 PictureBox 中的圖片存儲到資料庫
        private void SaveImageToDatabase(string NO)
        {
            // 替換為您的 PictureBox 控制項名稱
            Image image = pictureBox1.Image;

            if (image != null)
            {
                byte[] imageBytes = ImageToByteArray(image);
                InsertImageIntoDatabase(NO, DateTime.Now.ToString("yyyyMMdd HH:MM:ss"), imageBytes);

            }
            else
            {

            }
        }


        private void DELETE_ImageIntoDatabase(string NO)
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
                                    DELETE [TKWAREHOUSE].[dbo].[PACKAGEBOXSPHOTO]
                                    WHERE NO=@NO"
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
                string PACKAGENAMES = textBox7.Text;
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
            string NO = textBox1.Text;
            string BOXNO = textBox2.Text;
            string ALLWEIGHTS = textBox3.Text;
            string PACKWEIGHTS = textBox4.Text;
            string PRODUCTWEIGHTS = textBox5.Text;
            string PACKRATES = textBox6.Text;
            string RATECLASS = comboBox1.Text.ToString();
            string CHECKRATES = comboBox2.Text.ToString();
            string ISVALIDS = comboBox3.Text.ToString();
            string PACKAGENAMES = textBox7.Text;
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
                string NO = textBox1.Text;
                if (!string.IsNullOrEmpty(NO))
                {
                    TPACKAGEBOXS_DELETE(NO);
                    Search_PACKAGEBOXS(TG001TG002);
                }


            }
        }
        private void button5_Click(object sender, EventArgs e)
        {
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
            if(!string.IsNullOrEmpty(NO))
            {
                //string imagePath = System.Environment.CurrentDirectory;
                string imagePath = Path.Combine(Environment.CurrentDirectory, "Images", DateTime.Now.ToString("yyyy"));
                string imagePathNames = imagePath + "\\" + NO + ".jpg";
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
            if (!string.IsNullOrEmpty(NO))
            {
                //DisplayImageFromFolder("");
                //pictureBox1.Image = null;

                pictureBox1.Image.Dispose();
                pictureBox1.Image = null;
                pictureBox1.ImageLocation = null;

                string imagePath = Path.Combine(Environment.CurrentDirectory, "Images", DateTime.Now.ToString("yyyy"));
                string imagePathNames = imagePath + "\\" + NO + ".jpg";

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

        #endregion


    }
}

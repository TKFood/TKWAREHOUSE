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
        string MOCTA001002 = null;

        public Report report1 { get; private set; }

        public FrmREPORTMOCTE()
        {
            InitializeComponent();
            combobox2load();
            combobox4load();
        }
        private void FrmREPORTMOCTE_Load(object sender, EventArgs e)
        {
            dataGridView1.AlternatingRowsDefaultCellStyle.BackColor = Color.PaleTurquoise;      //奇數列顏色

            //先建立個 CheckBox 欄
            DataGridViewCheckBoxColumn cbCol = new DataGridViewCheckBoxColumn();
            cbCol.Width = 120;   //設定寬度
            cbCol.HeaderText = "　選擇";
            cbCol.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;   //置中
            cbCol.TrueValue = true;
            cbCol.FalseValue = false;
            dataGridView1.Columns.Insert(0, cbCol);

            //region 建立全选 CheckBox

            //建立个矩形，等下计算 CheckBox 嵌入 GridView 的位置
            Rectangle rect = dataGridView1.GetCellDisplayRectangle(0, -1, true);
            rect.X = rect.Location.X + rect.Width / 4 - 18;
            rect.Y = rect.Location.Y + (rect.Height / 2 - 9);

            CheckBox cbHeader = new CheckBox();
            cbHeader.Name = "checkboxHeader";
            cbHeader.Size = new Size(18, 18);
            cbHeader.Location = rect.Location;

            //全选要设定的事件
            cbHeader.CheckedChanged += new EventHandler(cbHeader_CheckedChanged);

            //将 CheckBox 加入到 dataGridView
            dataGridView1.Controls.Add(cbHeader);

        }
        private void cbHeader_CheckedChanged(object sender, EventArgs e)
        {
            dataGridView1.EndEdit();

            foreach (DataGridViewRow dr in dataGridView1.Rows)
            {
                dr.Cells[0].Value = ((CheckBox)dataGridView1.Controls.Find("checkboxHeader", true)[0]).Checked;

            }

        }

        #region FUNCTION
        public void combobox2load()
        {

            //20210902密
            Class1 TKID = new Class1();//用new 建立類別實體
            SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

            //資料庫使用者密碼解密
            sqlsb.Password = TKID.Decryption(sqlsb.Password);
            sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

            String connectionString;
            sqlConn = new SqlConnection(sqlsb.ConnectionString);

            String Sequel = "SELECT MD001,MD002 FROM [TK].dbo.CMSMD    WHERE MD003 IN ('20') ORDER BY  MD001";
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

        public void combobox4load()
        {

            //20210902密
            Class1 TKID = new Class1();//用new 建立類別實體
            SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

            //資料庫使用者密碼解密
            sqlsb.Password = TKID.Decryption(sqlsb.Password);
            sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

            String connectionString;
            sqlConn = new SqlConnection(sqlsb.ConnectionString);

            String Sequel = "SELECT MD001,MD002 FROM [TK].dbo.CMSMD    WHERE MD003 IN ('20') ORDER BY  MD001";
            SqlDataAdapter da = new SqlDataAdapter(Sequel, sqlConn);
            DataTable dt = new DataTable();
            sqlConn.Open();

            dt.Columns.Add("MD001", typeof(string));
            dt.Columns.Add("MD002", typeof(string));
            da.Fill(dt);
            comboBox4.DataSource = dt.DefaultView;
            comboBox4.ValueMember = "MD001";
            comboBox4.DisplayMember = "MD002";
            sqlConn.Close();



        }

        public void SETFASTREPORT()
        {
            try
            {
                string SQL;
                report1 = new Report();
                report1.Load(@"REPORT\合併領料.frx");

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
                SQL = SETFASETSQL();
                Table.SelectCommand = SQL;
                report1.Preview = previewControl1;
                report1.Show();
            }
            catch (Exception EX)
            {

            }
            finally
            {

            }
            

        }

        public string SETFASETSQL()
        {
            StringBuilder FASTSQL = new StringBuilder();
            StringBuilder STRQUERY = new StringBuilder();

            if(comboBox1.Text.ToString().Equals("原料"))
            {
               
                FASTSQL.AppendFormat(@"    
                                    SELECT
                                    MD.MD002,
                                    T.TE004,
                                    MB.MB002 AS TE017,
                                    T.TE011,
                                    T.TE012,
                                    SUM(MQ.MQ010 * T.TE005) * -1 AS TE005,
                                    T.TE010,
                                    -- 條件彙總 (Conditional Aggregation) 優化: 取代四個相關子查詢
                                    ISNULL(SUM(CASE WHEN T.TE001 = 'A541' THEN T.TE005 ELSE 0 END), 0) AS '領料',
                                    ISNULL(SUM(CASE WHEN T.TE001 = 'A542' THEN T.TE005 ELSE 0 END), 0) AS '補料',
                                    ISNULL(SUM(CASE WHEN T.TE001 = 'A561' THEN T.TE005 ELSE 0 END), 0) AS '退料',
                                    -- 使用 LEFT JOIN 或 OUTER APPLY 取得庫存量 (請參考下方 #優化庫存量獲取)
                                    ISNULL(INV.庫存量, 0) AS '庫存量'

                                    -- **使用明確連線 (Explicit Joins) 取代逗號連線**
                                    FROM [TK].dbo.MOCTE AS T
                                    INNER JOIN [TK].dbo.CMSMQ AS MQ ON MQ.MQ001 = T.TE001
                                    INNER JOIN [TK].dbo.INVMB AS MB ON T.TE004 = MB.MB001
                                    INNER JOIN [TK].dbo.MOCTC AS TC ON TC.TC001 = T.TE001 AND TC.TC002 = T.TE002
                                    INNER JOIN [TK].dbo.CMSMD AS MD ON MD.MD001 = TC.TC005

                                    -- **將 IN 子句轉換為 INNER JOIN 以優化篩選**
                                    INNER JOIN (
                                        SELECT
                                            LTRIM(RTRIM(TA001)) AS Clean_TA001,
                                            LTRIM(RTRIM(TA002)) AS Clean_TA002
                                        FROM [TK].dbo.MOCTA AS TA
                                        INNER JOIN [TK].dbo.CMSMD AS MD_A ON TA.TA021 = MD_A.MD001
                                        -- **重要: 針對日期和 MD002 欄位建立索引**
                                        WHERE TA.TA003 >= '{0}'
                                          AND TA.TA003 <= '{1}'
                                          AND MD_A.MD002 = '{2}'
                                        GROUP BY LTRIM(RTRIM(TA001)), LTRIM(RTRIM(TA002)) -- 確保唯一性
                                    ) AS FilterData
                                        ON LTRIM(RTRIM(T.TE011)) = FilterData.Clean_TA001
                                        AND LTRIM(RTRIM(T.TE012)) = FilterData.Clean_TA002

                                    -- **優化庫存量 (Inventory) 查詢：將其預先彙總後 LEFT JOIN**
                                    LEFT JOIN (
                                        SELECT
                                            LA001,
                                            SUM(LA005 * LA011) AS 庫存量
                                        FROM [TK].dbo.INVLA
                                        -- **重要: 針對 LA009 和 LA001 建立索引**
                                        WHERE LA009 IN ('20004', '20006')
                                        GROUP BY LA001
                                    ) AS INV ON INV.LA001 = T.TE004

                                    WHERE 1=1
                                        -- 來自舊式 CMSMD 的篩選條件
                                        AND MD.MD003 IN ('20')

                                        -- TE004 模糊查詢條件 (請注意：LIKE '%...' 還是會讓索引失效)
                                        AND (T.TE004 LIKE '1%' OR (T.TE004 LIKE '301%' AND LEN(T.TE004) = 10))

                                    GROUP BY
                                    MD.MD002,
                                    T.TE004,
                                    MB.MB002,
                                    T.TE011,
                                    T.TE012,
                                    T.TE010,
                                    -- 修正：將 '庫存量' 的表達式加入 GROUP BY
                                    ISNULL(INV.庫存量, 0)

                                -- **ORDER BY 欄位**
                                ORDER BY
                                    MD.MD002,
                                    T.TE004,
                                    MB.MB002,
                                    T.TE011,
                                    T.TE012,
                                    T.TE010,
                                    -- 修正：將 '庫存量' 的表達式加入 ORDER BY
                                    ISNULL(INV.庫存量, 0);
                              
                                    ", dateTimePicker1.Value.ToString("yyyyMMdd"), dateTimePicker2.Value.ToString("yyyyMMdd"), comboBox2.Text.ToString());
                
            }
            else if (comboBox1.Text.ToString().Equals("物料"))
            {
                FASTSQL.AppendFormat(@"   
                                    SELECT
                                    MD.MD002,
                                    T.TE004,
                                    MB.MB002 AS TE017,
                                    T.TE011,
                                    T.TE012,
                                    SUM(MQ.MQ010 * T.TE005) * -1 AS TE005,
                                    T.TE010,
                                    -- 條件彙總 (Conditional Aggregation) 優化: 取代四個相關子查詢
                                    ISNULL(SUM(CASE WHEN T.TE001 = 'A541' THEN T.TE005 ELSE 0 END), 0) AS '領料',
                                    ISNULL(SUM(CASE WHEN T.TE001 = 'A542' THEN T.TE005 ELSE 0 END), 0) AS '補料',
                                    ISNULL(SUM(CASE WHEN T.TE001 = 'A561' THEN T.TE005 ELSE 0 END), 0) AS '退料',
                                    -- 使用 LEFT JOIN 或 OUTER APPLY 取得庫存量 (請參考下方 #優化庫存量獲取)
                                    ISNULL(INV.庫存量, 0) AS '庫存量'

                                    -- **使用明確連線 (Explicit Joins) 取代逗號連線**
                                    FROM [TK].dbo.MOCTE AS T
                                    INNER JOIN [TK].dbo.CMSMQ AS MQ ON MQ.MQ001 = T.TE001
                                    INNER JOIN [TK].dbo.INVMB AS MB ON T.TE004 = MB.MB001
                                    INNER JOIN [TK].dbo.MOCTC AS TC ON TC.TC001 = T.TE001 AND TC.TC002 = T.TE002
                                    INNER JOIN [TK].dbo.CMSMD AS MD ON MD.MD001 = TC.TC005

                                    -- **將 IN 子句轉換為 INNER JOIN 以優化篩選**
                                    INNER JOIN (
                                        SELECT
                                            LTRIM(RTRIM(TA001)) AS Clean_TA001,
                                            LTRIM(RTRIM(TA002)) AS Clean_TA002
                                        FROM [TK].dbo.MOCTA AS TA
                                        INNER JOIN [TK].dbo.CMSMD AS MD_A ON TA.TA021 = MD_A.MD001
                                        -- **重要: 針對日期和 MD002 欄位建立索引**
                                        WHERE TA.TA003 >= '{0}'
                                          AND TA.TA003 <= '{1}'
                                          AND MD_A.MD002 = '{2}'
                                        GROUP BY LTRIM(RTRIM(TA001)), LTRIM(RTRIM(TA002)) -- 確保唯一性
                                    ) AS FilterData
                                        ON LTRIM(RTRIM(T.TE011)) = FilterData.Clean_TA001
                                        AND LTRIM(RTRIM(T.TE012)) = FilterData.Clean_TA002

                                    -- **優化庫存量 (Inventory) 查詢：將其預先彙總後 LEFT JOIN**
                                    LEFT JOIN (
                                        SELECT
                                            LA001,
                                            SUM(LA005 * LA011) AS 庫存量
                                        FROM [TK].dbo.INVLA
                                        -- **重要: 針對 LA009 和 LA001 建立索引**
                                        WHERE LA009 IN ('20004', '20006')
                                        GROUP BY LA001
                                    ) AS INV ON INV.LA001 = T.TE004

                                    WHERE 1=1
                                        -- 來自舊式 CMSMD 的篩選條件
                                        AND MD.MD003 IN ('20')

                                        -- TE004 模糊查詢條件 (請注意：LIKE '%...' 還是會讓索引失效)
                                        AND (T.TE004 LIKE '2%' )

                                    GROUP BY
                                    MD.MD002,
                                    T.TE004,
                                    MB.MB002,
                                    T.TE011,
                                    T.TE012,
                                    T.TE010,
                                    -- 修正：將 '庫存量' 的表達式加入 GROUP BY
                                    ISNULL(INV.庫存量, 0)

                                -- **ORDER BY 欄位**
                                ORDER BY
                                    MD.MD002,
                                    T.TE004,
                                    MB.MB002,
                                    T.TE011,
                                    T.TE012,
                                    T.TE010,
                                    -- 修正：將 '庫存量' 的表達式加入 ORDER BY
                                    ISNULL(INV.庫存量, 0);
                                    ", dateTimePicker1.Value.ToString("yyyyMMdd"), dateTimePicker2.Value.ToString("yyyyMMdd"), comboBox2.Text.ToString());



            }
            else if (comboBox1.Text.ToString().Equals("原料+物料"))
            {
                FASTSQL.AppendFormat(@"   
                                    SELECT
                                    MD.MD002,
                                    T.TE004,
                                    MB.MB002 AS TE017,
                                    T.TE011,
                                    T.TE012,
                                    SUM(MQ.MQ010 * T.TE005) * -1 AS TE005,
                                    T.TE010,
                                    -- 條件彙總 (Conditional Aggregation) 優化: 取代四個相關子查詢
                                    ISNULL(SUM(CASE WHEN T.TE001 = 'A541' THEN T.TE005 ELSE 0 END), 0) AS '領料',
                                    ISNULL(SUM(CASE WHEN T.TE001 = 'A542' THEN T.TE005 ELSE 0 END), 0) AS '補料',
                                    ISNULL(SUM(CASE WHEN T.TE001 = 'A561' THEN T.TE005 ELSE 0 END), 0) AS '退料',
                                    -- 使用 LEFT JOIN 或 OUTER APPLY 取得庫存量 (請參考下方 #優化庫存量獲取)
                                    ISNULL(INV.庫存量, 0) AS '庫存量'

                                    -- **使用明確連線 (Explicit Joins) 取代逗號連線**
                                    FROM [TK].dbo.MOCTE AS T
                                    INNER JOIN [TK].dbo.CMSMQ AS MQ ON MQ.MQ001 = T.TE001
                                    INNER JOIN [TK].dbo.INVMB AS MB ON T.TE004 = MB.MB001
                                    INNER JOIN [TK].dbo.MOCTC AS TC ON TC.TC001 = T.TE001 AND TC.TC002 = T.TE002
                                    INNER JOIN [TK].dbo.CMSMD AS MD ON MD.MD001 = TC.TC005

                                    -- **將 IN 子句轉換為 INNER JOIN 以優化篩選**
                                    INNER JOIN (
                                        SELECT
                                            LTRIM(RTRIM(TA001)) AS Clean_TA001,
                                            LTRIM(RTRIM(TA002)) AS Clean_TA002
                                        FROM [TK].dbo.MOCTA AS TA
                                        INNER JOIN [TK].dbo.CMSMD AS MD_A ON TA.TA021 = MD_A.MD001
                                        -- **重要: 針對日期和 MD002 欄位建立索引**
                                        WHERE TA.TA003 >= '{0}'
                                          AND TA.TA003 <= '{1}'
                                          AND MD_A.MD002 = '{2}'
                                        GROUP BY LTRIM(RTRIM(TA001)), LTRIM(RTRIM(TA002)) -- 確保唯一性
                                    ) AS FilterData
                                        ON LTRIM(RTRIM(T.TE011)) = FilterData.Clean_TA001
                                        AND LTRIM(RTRIM(T.TE012)) = FilterData.Clean_TA002

                                    -- **優化庫存量 (Inventory) 查詢：將其預先彙總後 LEFT JOIN**
                                    LEFT JOIN (
                                        SELECT
                                            LA001,
                                            SUM(LA005 * LA011) AS 庫存量
                                        FROM [TK].dbo.INVLA
                                        -- **重要: 針對 LA009 和 LA001 建立索引**
                                        WHERE LA009 IN ('20004', '20006')
                                        GROUP BY LA001
                                    ) AS INV ON INV.LA001 = T.TE004

                                    WHERE 1=1
                                        -- 來自舊式 CMSMD 的篩選條件
                                        AND MD.MD003 IN ('20')

                                        -- TE004 模糊查詢條件 (請注意：LIKE '%...' 還是會讓索引失效)
                                        AND (T.TE004 LIKE '1%' OR T.TE004 LIKE '2%' OR (T.TE004 LIKE '301%' AND LEN(T.TE004) = 10))

                                    GROUP BY
                                    MD.MD002,
                                    T.TE004,
                                    MB.MB002,
                                    T.TE011,
                                    T.TE012,
                                    T.TE010,
                                    -- 修正：將 '庫存量' 的表達式加入 GROUP BY
                                    ISNULL(INV.庫存量, 0)

                                -- **ORDER BY 欄位**
                                ORDER BY
                                    MD.MD002,
                                    T.TE004,
                                    MB.MB002,
                                    T.TE011,
                                    T.TE012,
                                    T.TE010,
                                    -- 修正：將 '庫存量' 的表達式加入 ORDER BY
                                    ISNULL(INV.庫存量, 0);
                                    ", dateTimePicker1.Value.ToString("yyyyMMdd"), dateTimePicker2.Value.ToString("yyyyMMdd"), comboBox2.Text.ToString());


               
            }


            FASTSQL.AppendFormat(@"   ");

            return FASTSQL.ToString();
        }

        public void SETFASTREPORT2(string MOCTA001002)
        {

            string SQL;
            report1 = new Report();
            report1.Load(@"REPORT\分開領料.frx");

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
            SQL = SETFASETSQL2();
            Table.SelectCommand = SQL;
            report1.Preview = previewControl2;
            report1.Show();

        }

        public string SETFASETSQL2()
        {
            StringBuilder FASTSQL = new StringBuilder();
            StringBuilder STRQUERY = new StringBuilder();

            if (comboBox3.Text.ToString().Equals("原料"))
            {

                FASTSQL.AppendFormat(@"   
  
                                    SELECT 線別,品號,品名,製令單別,製令單號,批號
                                    ,(SELECT ISNULL(TB004,0) FROM [TK].dbo.MOCTB WHERE TB001=製令單別 AND TB002=製令單號 AND TB003=品號) AS '應領料量'
                                    ,((SELECT ISNULL(SUM(TE005),0) FROM [TK].dbo.MOCTE TE WHERE TE.TE004=品號 AND TE.TE011=製令單別 AND TE.TE012=製令單號 AND TE.TE010=批號 AND TE.TE001='A541' )+(SELECT ISNULL(SUM(TE005),0) FROM [TK].dbo.MOCTE TE WHERE TE.TE004=品號 AND TE.TE011=製令單別 AND TE.TE012=製令單號 AND TE.TE010=批號 AND TE.TE001='A542' )-(SELECT ISNULL(SUM(TE005),0) FROM [TK].dbo.MOCTE TE WHERE TE.TE004=品號 AND TE.TE011=製令單別 AND TE.TE012=製令單號 AND TE.TE010=批號 AND TE.TE001='A561' )) AS '實發數量' 
                                    ,(SELECT ISNULL(SUM(TE005),0) FROM [TK].dbo.MOCTE TE WHERE TE.TE004=品號 AND TE.TE011=製令單別 AND TE.TE012=製令單號 AND TE.TE010=批號 AND TE.TE001='A542' ) AS '補料數量'
                                    ,(SELECT ISNULL(SUM(TE005),0) FROM [TK].dbo.MOCTE TE WHERE TE.TE004=品號 AND TE.TE011=製令單別 AND TE.TE012=製令單號 AND TE.TE010=批號 AND TE.TE001='A561' ) AS '退料數量' 
                                    ,(SELECT ISNULL(SUM(LA005*LA011),0) FROM [TK].dbo.INVLA WHERE LA009 IN ('20004','20006')  AND LA001=品號 ) AS '庫存數量' 

                                    FROM 
                                    (
                                    SELECT MD002 AS '線別',TE004 AS '品號',TE017 AS '品名' ,TE011 AS '製令單別',TE012 AS '製令單號',TE010  AS '批號'

                                    FROM [TK].dbo.CMSMD, [TK].dbo.MOCTC,[TK].dbo.MOCTE,[TK].dbo.[CMSMQ]
                                    WHERE MQ001=TE001
                                    AND MD003 IN ('20') 
                                    AND MD001=TC005 
                                    AND TC001=TE001 AND TC002=TE002 
                                    AND TE001 IN ('A541','A542')

                                    AND (TE004 LIKE '1%'  OR (TE004 LIKE '301%' AND LEN(TE004)=10))  
                                    AND LTRIM(RTRIM(TE011))+ LTRIM(RTRIM(TE012)) IN (SELECT LTRIM(RTRIM(TA001))+ LTRIM(RTRIM(TA002)) FROM [TK].dbo.MOCTA,[TK].dbo.CMSMD WHERE TA021=MD001 AND LTRIM(RTRIM(TA001))+LTRIM(RTRIM(TA002)) IN ({1}) AND MD002='{0}')

                                    GROUP BY MD002,TE004,TE017  ,TE011,TE012,TE010

                                    ) AS TEMP
                                    ORDER BY 製令單別,製令單號,品號,批號
   
                                    ", comboBox4.Text.ToString(), MOCTA001002);

            }
            else if (comboBox3.Text.ToString().Equals("物料"))
            {
                FASTSQL.AppendFormat(@"   
                                    SELECT 線別,品號,品名,製令單別,製令單號,批號
                                    ,(SELECT ISNULL(TB004,0) FROM [TK].dbo.MOCTB WHERE TB001=製令單別 AND TB002=製令單號 AND TB003=品號) AS '應領料量'
                                    ,((SELECT ISNULL(SUM(TE005),0) FROM [TK].dbo.MOCTE TE WHERE TE.TE004=品號 AND TE.TE011=製令單別 AND TE.TE012=製令單號 AND TE.TE010=批號 AND TE.TE001='A541' )+(SELECT ISNULL(SUM(TE005),0) FROM [TK].dbo.MOCTE TE WHERE TE.TE004=品號 AND TE.TE011=製令單別 AND TE.TE012=製令單號 AND TE.TE010=批號 AND TE.TE001='A542' )-(SELECT ISNULL(SUM(TE005),0) FROM [TK].dbo.MOCTE TE WHERE TE.TE004=品號 AND TE.TE011=製令單別 AND TE.TE012=製令單號 AND TE.TE010=批號 AND TE.TE001='A561' )) AS '實發數量' 
                                    ,(SELECT ISNULL(SUM(TE005),0) FROM [TK].dbo.MOCTE TE WHERE TE.TE004=品號 AND TE.TE011=製令單別 AND TE.TE012=製令單號 AND TE.TE010=批號 AND TE.TE001='A542' ) AS '補料數量'
                                    ,(SELECT ISNULL(SUM(TE005),0) FROM [TK].dbo.MOCTE TE WHERE TE.TE004=品號 AND TE.TE011=製令單別 AND TE.TE012=製令單號 AND TE.TE010=批號 AND TE.TE001='A561' ) AS '退料數量' 
                                    ,(SELECT ISNULL(SUM(LA005*LA011),0) FROM [TK].dbo.INVLA WHERE LA009 IN ('20004','20006')  AND LA001=品號 ) AS '庫存數量' 

                                    FROM 
                                    (
                                    SELECT MD002 AS '線別',TE004 AS '品號',TE017 AS '品名' ,TE011 AS '製令單別',TE012 AS '製令單號',SUM((MQ010*TE005)*-1)  AS '應領料量',TE010  AS '批號'

                                    FROM [TK].dbo.CMSMD, [TK].dbo.MOCTC,[TK].dbo.MOCTE,[TK].dbo.[CMSMQ]
                                    WHERE MQ001=TE001
                                    AND MD003 IN ('20') 
                                    AND MD001=TC005 
                                    AND TC001=TE001 AND TC002=TE002 
                                    AND TE001 IN ('A541','A542')

                                    AND (TE004 LIKE '2%' )   
                                    AND LTRIM(RTRIM(TE011))+ LTRIM(RTRIM(TE012)) IN (SELECT LTRIM(RTRIM(TA001))+ LTRIM(RTRIM(TA002)) FROM [TK].dbo.MOCTA,[TK].dbo.CMSMD WHERE TA021=MD001 AND LTRIM(RTRIM(TA001))+LTRIM(RTRIM(TA002)) IN ({1}) AND MD002='{0}')

                                    GROUP BY MD002,TE004,TE017  ,TE011,TE012,TE010

                                    ) AS TEMP
                                    ORDER BY 製令單別,製令單號,品號,批號
                                    ", comboBox4.Text.ToString(), MOCTA001002);



            }
            else if (comboBox3.Text.ToString().Equals("原料+物料"))
            {
                FASTSQL.AppendFormat(@"   
                                    SELECT 線別,品號,品名,製令單別,製令單號,批號
                                    ,(SELECT ISNULL(TB004,0) FROM [TK].dbo.MOCTB WHERE TB001=製令單別 AND TB002=製令單號 AND TB003=品號) AS '應領料量'
                                    ,((SELECT ISNULL(SUM(TE005),0) FROM [TK].dbo.MOCTE TE WHERE TE.TE004=品號 AND TE.TE011=製令單別 AND TE.TE012=製令單號 AND TE.TE010=批號 AND TE.TE001='A541' )+(SELECT ISNULL(SUM(TE005),0) FROM [TK].dbo.MOCTE TE WHERE TE.TE004=品號 AND TE.TE011=製令單別 AND TE.TE012=製令單號 AND TE.TE010=批號 AND TE.TE001='A542' )-(SELECT ISNULL(SUM(TE005),0) FROM [TK].dbo.MOCTE TE WHERE TE.TE004=品號 AND TE.TE011=製令單別 AND TE.TE012=製令單號 AND TE.TE010=批號 AND TE.TE001='A561' )) AS '實發數量' 
                                    ,(SELECT ISNULL(SUM(TE005),0) FROM [TK].dbo.MOCTE TE WHERE TE.TE004=品號 AND TE.TE011=製令單別 AND TE.TE012=製令單號 AND TE.TE010=批號 AND TE.TE001='A542' ) AS '補料數量'
                                    ,(SELECT ISNULL(SUM(TE005),0) FROM [TK].dbo.MOCTE TE WHERE TE.TE004=品號 AND TE.TE011=製令單別 AND TE.TE012=製令單號 AND TE.TE010=批號 AND TE.TE001='A561' ) AS '退料數量' 
                                    ,(SELECT ISNULL(SUM(LA005*LA011),0) FROM [TK].dbo.INVLA WHERE LA009 IN ('20004','20006')  AND LA001=品號 ) AS '庫存數量' 

                                    FROM 
                                    (
                                    SELECT MD002 AS '線別',TE004 AS '品號',TE017 AS '品名' ,TE011 AS '製令單別',TE012 AS '製令單號',SUM((MQ010*TE005)*-1)  AS '應領料量',TE010  AS '批號'

                                    FROM [TK].dbo.CMSMD, [TK].dbo.MOCTC,[TK].dbo.MOCTE,[TK].dbo.[CMSMQ]
                                    WHERE MQ001=TE001
                                    AND MD003 IN ('20') 
                                    AND MD001=TC005 
                                    AND TC001=TE001 AND TC002=TE002 
                                    AND TE001 IN ('A541','A542')

                                   AND (TE004 LIKE '1%' OR TE004 LIKE '2%' OR (TE004 LIKE '301%' AND LEN(TE004)=10))   
                                   AND LTRIM(RTRIM(TE011))+ LTRIM(RTRIM(TE012)) IN (SELECT LTRIM(RTRIM(TA001))+ LTRIM(RTRIM(TA002)) FROM [TK].dbo.MOCTA,[TK].dbo.CMSMD WHERE TA021=MD001 AND LTRIM(RTRIM(TA001))+LTRIM(RTRIM(TA002)) IN ({1}) AND MD002='{0}')

                                    GROUP BY MD002,TE004,TE017  ,TE011,TE012,TE010

                                    ) AS TEMP
                                    ORDER BY 製令單別,製令單號,品號,批號
                                     ", comboBox4.Text.ToString(), MOCTA001002);



            }


            FASTSQL.AppendFormat(@"  ");

            return FASTSQL.ToString();
        }

        public void SEARCHMOCTAB()
        {
            SqlDataAdapter adapter1 = new SqlDataAdapter();
            SqlCommandBuilder sqlCmdBuilder1 = new SqlCommandBuilder();
            DataSet ds1 = new DataSet();

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
                                    SELECT MD002 AS '線別',LTRIM(RTRIM(TA001)) AS  '製令單別',LTRIM(RTRIM(TA002)) AS  '製令單號',TA006 AS  '品號',TA034 AS  '品名',TA015 AS  '預計生產量', TA007 AS  '單位'
                                    FROM [TK].dbo.MOCTA,[TK].dbo.CMSMD 
                                    WHERE TA021=MD001 AND  TA003>='{0}' AND TA003<='{1}' 
                                    ORDER BY TA001,TA002

                                    ", dateTimePicker3.Value.ToString("yyyyMMdd"), dateTimePicker4.Value.ToString("yyyyMMdd"));

                adapter1 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder1 = new SqlCommandBuilder(adapter1);
                sqlConn.Open();
                ds1.Clear();
                adapter1.Fill(ds1, "TEMPds1");
                sqlConn.Close();


                if (ds1.Tables["TEMPds1"].Rows.Count == 0)
                {
                    dataGridView1.DataSource = null;
                }
                else
                {
                    if (ds1.Tables["TEMPds1"].Rows.Count >= 1)
                    {
                        //dataGridView1.Rows.Clear();
                        dataGridView1.DataSource = ds1.Tables["TEMPds1"];
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

        public void FINDCHECK()
        {
            MOCTA001002 = null;

            foreach (DataGridViewRow dr in this.dataGridView1.Rows)
            {                
                if (dr.Cells[0].Value != null && (bool)dr.Cells[0].Value)
                {
                    MOCTA001002 = MOCTA001002 +"'"+ dr.Cells["製令單別"].Value.ToString() + dr.Cells["製令單號"].Value.ToString() + "',";
                    //MessageBox.Show(dr.Cells["製令單號"].Value.ToString());

                }
            }

            MOCTA001002 = MOCTA001002 +"''";

            if(!string.IsNullOrEmpty(MOCTA001002))
            {
                SETFASTREPORT2(MOCTA001002);
            }
            
            //MessageBox.Show(MOCTA001002.ToString());
        }

        public void SETFASTREPORT3(string MOCTA001002,string MD002)
        {
             
            string SQL;
            report1 = new Report();
            report1.Load(@"REPORT\單次合併領料.frx");

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

            SQL = SETFASETSQL3(MOCTA001002, MD002);
            Table.SelectCommand = SQL;
            report1.Preview = previewControl2;
            report1.Show();

        }

        public string SETFASETSQL3(string MOCTA001002,string MD002)
        {
            StringBuilder FASTSQL = new StringBuilder();
            StringBuilder STRQUERY = new StringBuilder();

            if (comboBox3.Text.ToString().Equals("原料"))
            {

                FASTSQL.AppendFormat(@"   
                                    SELECT MD002,TE004,TE017 ,TE011,TE012,SUM(MQ010*TE005)*-1  AS TE005,TE010 
                                    ,(SELECT ISNULL(TB004,0) FROM [TK].dbo.MOCTB WHERE TB001=TE011 AND TB002=TE012 AND TB003=TE004) AS '應領料量'
                                    ,(SELECT ISNULL(SUM(TE005),0) FROM [TK].dbo.MOCTE TE WHERE TE.TE004=MOCTE.TE004 AND TE.TE011=MOCTE.TE011 AND TE.TE012=MOCTE.TE012 AND TE.TE010=MOCTE.TE010 AND TE.TE001='A541' ) AS '領料' 
                                    ,(SELECT ISNULL(SUM(TE005),0) FROM [TK].dbo.MOCTE TE WHERE TE.TE004=MOCTE.TE004 AND TE.TE011=MOCTE.TE011 AND TE.TE012=MOCTE.TE012 AND TE.TE010=MOCTE.TE010 AND TE.TE001='A542' ) AS '補料'
                                    ,(SELECT ISNULL(SUM(TE005),0) FROM [TK].dbo.MOCTE TE WHERE TE.TE004=MOCTE.TE004 AND TE.TE011=MOCTE.TE011 AND TE.TE012=MOCTE.TE012 AND TE.TE010=MOCTE.TE010 AND TE.TE001='A561' ) AS '退料' 
                                    ,(SELECT SUM(LA005*LA011) FROM [TK].dbo.INVLA WHERE LA009 IN ('20004','20006')  AND LA001=TE004 ) AS '庫存量' 
                                    FROM [TK].dbo.CMSMD, [TK].dbo.MOCTC,[TK].dbo.MOCTE,[TK].dbo.[CMSMQ]
                                    WHERE MQ001=TE001
                                    AND MD003 IN ('20') 
                                    AND MD001=TC005 
                                    AND TC001=TE001 AND TC002=TE002 
                                    AND ((TE004 LIKE '1%' ) OR (TE004 LIKE '301%' AND LEN(TE004)=10))   
                                    AND LTRIM(RTRIM(TE011))+ LTRIM(RTRIM(TE012)) IN ({0} )
                                    AND MD002='{1}'

                                    GROUP BY MD002,TE004,TE017 ,TE011,TE012,TE010 
                                    ORDER BY MD002,TE004,TE017 ,TE011,TE012,TE010 
                                    ", MOCTA001002, MD002);

            }
            else if (comboBox3.Text.ToString().Equals("物料"))
            {
                FASTSQL.AppendFormat(@"   
                                    SELECT MD002,TE004,TE017 ,TE011,TE012,SUM(MQ010*TE005)*-1  AS TE005,TE010 
                                    ,(SELECT ISNULL(TB004,0) FROM [TK].dbo.MOCTB WHERE TB001=TE011 AND TB002=TE012 AND TB003=TE004) AS '應領料量'
                                    ,(SELECT ISNULL(SUM(TE005),0) FROM [TK].dbo.MOCTE TE WHERE TE.TE004=MOCTE.TE004 AND TE.TE011=MOCTE.TE011 AND TE.TE012=MOCTE.TE012 AND TE.TE010=MOCTE.TE010 AND TE.TE001='A541' ) AS '領料' 
                                    ,(SELECT ISNULL(SUM(TE005),0) FROM [TK].dbo.MOCTE TE WHERE TE.TE004=MOCTE.TE004 AND TE.TE011=MOCTE.TE011 AND TE.TE012=MOCTE.TE012 AND TE.TE010=MOCTE.TE010 AND TE.TE001='A542' ) AS '補料'
                                    ,(SELECT ISNULL(SUM(TE005),0) FROM [TK].dbo.MOCTE TE WHERE TE.TE004=MOCTE.TE004 AND TE.TE011=MOCTE.TE011 AND TE.TE012=MOCTE.TE012 AND TE.TE010=MOCTE.TE010 AND TE.TE001='A561' ) AS '退料' 
                                    ,(SELECT SUM(LA005*LA011) FROM [TK].dbo.INVLA WHERE LA009 IN ('20004','20006')  AND LA001=TE004 ) AS '庫存量' 
                                    FROM [TK].dbo.CMSMD, [TK].dbo.MOCTC,[TK].dbo.MOCTE,[TK].dbo.[CMSMQ]
                                    WHERE MQ001=TE001
                                    AND MD003 IN ('20') 
                                    AND MD001=TC005 
                                    AND TC001=TE001 AND TC002=TE002 
                                    AND (TE004 LIKE '2%' )   
                                    AND LTRIM(RTRIM(TE011))+ LTRIM(RTRIM(TE012)) IN ({0} )
                                    AND MD002='{1}'

                                    GROUP BY MD002,TE004,TE017 ,TE011,TE012,TE010 
                                    ORDER BY MD002,TE004,TE017 ,TE011,TE012,TE010 
                                    ", MOCTA001002, MD002);



            }
            else if (comboBox3.Text.ToString().Equals("原料+物料"))
            {
                FASTSQL.AppendFormat(@"   
                                    SELECT MD002,TE004,TE017 ,TE011,TE012,SUM(MQ010*TE005)*-1  AS TE005,TE010 
                                    ,(SELECT ISNULL(TB004,0) FROM [TK].dbo.MOCTB WHERE TB001=TE011 AND TB002=TE012 AND TB003=TE004) AS '應領料量'
                                    ,(SELECT ISNULL(SUM(TE005),0) FROM [TK].dbo.MOCTE TE WHERE TE.TE004=MOCTE.TE004 AND TE.TE011=MOCTE.TE011 AND TE.TE012=MOCTE.TE012 AND TE.TE010=MOCTE.TE010 AND TE.TE001='A541' ) AS '領料' 
                                    ,(SELECT ISNULL(SUM(TE005),0) FROM [TK].dbo.MOCTE TE WHERE TE.TE004=MOCTE.TE004 AND TE.TE011=MOCTE.TE011 AND TE.TE012=MOCTE.TE012 AND TE.TE010=MOCTE.TE010 AND TE.TE001='A542' ) AS '補料'
                                    ,(SELECT ISNULL(SUM(TE005),0) FROM [TK].dbo.MOCTE TE WHERE TE.TE004=MOCTE.TE004 AND TE.TE011=MOCTE.TE011 AND TE.TE012=MOCTE.TE012 AND TE.TE010=MOCTE.TE010 AND TE.TE001='A561' ) AS '退料' 
                                    ,(SELECT SUM(LA005*LA011) FROM [TK].dbo.INVLA WHERE LA009 IN ('20004','20006')  AND LA001=TE004 ) AS '庫存量' 
                                    FROM [TK].dbo.CMSMD, [TK].dbo.MOCTC,[TK].dbo.MOCTE,[TK].dbo.[CMSMQ]
                                    WHERE MQ001=TE001
                                    AND MD003 IN ('20') 
                                    AND MD001=TC005 
                                    AND TC001=TE001 AND TC002=TE002 
                                    AND (TE004 LIKE '1%' OR TE004 LIKE '2%' OR (TE004 LIKE '301%' AND LEN(TE004)=10))  
                                    AND LTRIM(RTRIM(TE011))+ LTRIM(RTRIM(TE012)) IN ({0} )
                                    AND MD002='{1}'

                                    GROUP BY MD002,TE004,TE017 ,TE011,TE012,TE010 
                                    ORDER BY MD002,TE004,TE017 ,TE011,TE012,TE010 
                                     ", MOCTA001002, MD002);



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


        private void button2_Click(object sender, EventArgs e)
        {
            SEARCHMOCTAB();
            //SETFASTREPORT2();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            FINDCHECK();
        }

        private void button4_Click(object sender, EventArgs e)
        {
            string MOCTA001002 = null;

            foreach (DataGridViewRow dr in this.dataGridView1.Rows)
            {
                if (dr.Cells[0].Value != null && (bool)dr.Cells[0].Value)
                {
                    MOCTA001002 = MOCTA001002 + "'" + dr.Cells["製令單別"].Value.ToString() + dr.Cells["製令單號"].Value.ToString() + "',";
                    //MessageBox.Show(dr.Cells["製令單號"].Value.ToString());

                }
            }

            MOCTA001002 = MOCTA001002 + "''";

            if (!string.IsNullOrEmpty(MOCTA001002))
            {
                //MessageBox.Show(MOCTA001002);

                SETFASTREPORT3(MOCTA001002,comboBox4.Text);
            }

        }

        #endregion


    }
}

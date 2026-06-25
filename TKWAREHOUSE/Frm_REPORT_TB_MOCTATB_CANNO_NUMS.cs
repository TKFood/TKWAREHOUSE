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
    public partial class Frm_REPORT_TB_MOCTATB_CANNO_NUMS : Form
    {
        SqlTransaction tran;
        SqlCommand cmd = new SqlCommand();      
        int result;
        public Report report1 { get; private set; }

        string TA001 = null;
        string TA002 = null;
        //找出生產說明用的品號
        string MAINMB001 = "";

        public Frm_REPORT_TB_MOCTATB_CANNO_NUMS()
        {
            InitializeComponent();
        }
        private void Frm_REPORT_TB_MOCTATB_CANNO_NUMS_Load(object sender, EventArgs e)
        {
            SetupDataGridView();
            SetupDataGridView4();
        }
        #region FUNCTION

        private void SetupDataGridView()
        {
            // 1. 建立一個 CheckBox 欄位
            DataGridViewCheckBoxColumn checkColumn = new DataGridViewCheckBoxColumn();
            checkColumn.Name = "SelectCheck";
            checkColumn.HeaderText = "選取";
            checkColumn.Width = 50;
            checkColumn.ReadOnly = false; // 確保使用者可以勾選
            checkColumn.TrueValue = true;
            checkColumn.FalseValue = false;

            // 2. 將勾選欄插入到 DataGridView 的最前面 (索引 0)
            if (!dataGridView1.Columns.Contains("SelectCheck"))
            {
                dataGridView1.Columns.Insert(0, checkColumn);
            }
        }

        private void SetupDataGridView4()
        {
            // 1. 建立一個 CheckBox 欄位
            DataGridViewCheckBoxColumn checkColumn = new DataGridViewCheckBoxColumn();
            checkColumn.Name = "SelectCheck";
            checkColumn.HeaderText = "選取";
            checkColumn.Width = 50;
            checkColumn.ReadOnly = false; // 確保使用者可以勾選
            checkColumn.TrueValue = true;
            checkColumn.FalseValue = false;

            // 2. 將勾選欄插入到 DataGridView 的最前面 (索引 0)
            if (!dataGridView4.Columns.Contains("SelectCheck"))
            {
                dataGridView4.Columns.Insert(0, checkColumn);
            }
        }
        public void SEARCH(string SDATE)
        {
            Class1 TKID = new Class1();//用new 建立類別實體
            // 1.取得原始的連線字串
            string originalConnString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;

            // 2. 使用 SqlConnectionStringBuilder 來解析與修改
            SqlConnectionStringBuilder builder = new SqlConnectionStringBuilder(originalConnString);

            // 3. 將內含的帳密解密後重新指派
            builder.UserID = TKID.Decryption(builder.UserID);
            builder.Password = TKID.Decryption(builder.Password);

            // 4. 用最後組合好的連線字串建立 SqlConnection
            SqlConnection connection = new SqlConnection(builder.ConnectionString);

            StringBuilder sb = new StringBuilder();
            sb.Append(@" 
                        SELECT 
                        線別,製令單別,製令單號,開單日期,產品品號,產品品名,預計產量,單位,總桶數
                        FROM (
	                        SELECT CMSMD.MD002 AS '線別',TA001 AS '製令單別',TA002 AS '製令單號',TA003 AS '開單日期',TA006 AS '產品品號',TA034 AS '產品品名',TA015 AS '預計產量',TA007 AS '單位',ROUND(TA015/MC004,3) AS '總桶數'
	                        FROM [TK].dbo.MOCTA,[TK].dbo.BOMMC,[TK].dbo.CMSMD
	                        WHERE 1=1
	                        AND TA006=MC001
	                        AND TA021=CMSMD.MD001
	                        AND CMSMD.MD002  IN (SELECT [MD002]  FROM [TKWAREHOUSE].[dbo].[TB_MOCTATB_CANNO_NUMS_CMSMD]) 
	                        AND TA003=@SDATE
                        ) AS TEMP
                        ORDER BY 線別,製令單別,製令單號  

                ");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.AddWithValue("@SDATE", SDATE);
            DataTable dt = new DataTable();
            try
            {
                connection.Open();
                SqlDataAdapter da = new SqlDataAdapter(command);
                da.Fill(dt);
                dataGridView1.DataSource = dt;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally
            {
                connection.Close();
            }
        }
        private void dataGridView1_SelectionChanged(object sender, EventArgs e)
        {
            string 製令單別 = null;
            string 製令單號 = null;
            textBox1.Text = null;
            textBox2.Text = null;

            if (dataGridView1.CurrentRow != null)
            {
                製令單別 = dataGridView1.CurrentRow.Cells["製令單別"].Value.ToString();
                製令單號 = dataGridView1.CurrentRow.Cells["製令單號"].Value.ToString();

                TA001 = dataGridView1.CurrentRow.Cells["製令單別"].Value.ToString();
                TA002 = dataGridView1.CurrentRow.Cells["製令單號"].Value.ToString();
                textBox1.Text = TA001;
                textBox2.Text = TA002;

                SEARCH_DETAILS(製令單別, 製令單號);
            }
        }

        public void SEARCH_DG4(string SDATE)
        {
            Class1 TKID = new Class1();//用new 建立類別實體
            // 1.取得原始的連線字串
            string originalConnString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;

            // 2. 使用 SqlConnectionStringBuilder 來解析與修改
            SqlConnectionStringBuilder builder = new SqlConnectionStringBuilder(originalConnString);

            // 3. 將內含的帳密解密後重新指派
            builder.UserID = TKID.Decryption(builder.UserID);
            builder.Password = TKID.Decryption(builder.Password);

            // 4. 用最後組合好的連線字串建立 SqlConnection
            SqlConnection connection = new SqlConnection(builder.ConnectionString);

            StringBuilder sb = new StringBuilder();
            sb.Append(@" 
                        SELECT 
                        線別,製令單別,製令單號,開單日期,產品品號,產品品名,預計產量,單位,總桶數
                        FROM (
	                        SELECT CMSMD.MD002 AS '線別',TA001 AS '製令單別',TA002 AS '製令單號',TA003 AS '開單日期',TA006 AS '產品品號',TA034 AS '產品品名',TA015 AS '預計產量',TA007 AS '單位',ROUND(TA015/MC004,3) AS '總桶數'
	                        FROM [TK].dbo.MOCTA,[TK].dbo.BOMMC,[TK].dbo.CMSMD
	                        WHERE 1=1
	                        AND TA006=MC001
	                        AND TA021=CMSMD.MD001
	                        AND CMSMD.MD002  IN (SELECT [MD002]  FROM [TKWAREHOUSE].[dbo].[TB_MOCTATB_CANNO_NUMS_CMSMD_OTHERS]) 
	                        AND TA003=@SDATE
                        ) AS TEMP
                        ORDER BY 線別,製令單別,製令單號  

                ");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.AddWithValue("@SDATE", SDATE);
            DataTable dt = new DataTable();
            try
            {
                connection.Open();
                SqlDataAdapter da = new SqlDataAdapter(command);
                da.Fill(dt);
                dataGridView4.DataSource = dt;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally
            {
                connection.Close();
            }
        }
        public void SEARCH_DETAILS(string 製令單別,string 製令單號)
        {
            Class1 TKID = new Class1();//用new 建立類別實體
            // 1.取得原始的連線字串
            string originalConnString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;

            // 2. 使用 SqlConnectionStringBuilder 來解析與修改
            SqlConnectionStringBuilder builder = new SqlConnectionStringBuilder(originalConnString);

            // 3. 將內含的帳密解密後重新指派
            builder.UserID = TKID.Decryption(builder.UserID);
            builder.Password = TKID.Decryption(builder.Password);

            // 4. 用最後組合好的連線字串建立 SqlConnection
            SqlConnection connection = new SqlConnection(builder.ConnectionString);

            StringBuilder sb = new StringBuilder();
            sb.Append(@"                       
                        SELECT 
                        線別,製令單別,製令單號,開單日期,產品品號,產品品名,預計產量,單位1,材料品號,材料品名,單位2,需領用量,總桶數,整桶數,最後桶數,整桶用量,最後桶用量,標準用量
                        ,ISNULL(BOMMD.MD004,'') AS '材料單位',ISNULL(BOMMD.MD006,0) AS '組成用量',ISNULL(BOMMD.MD007,0) AS '底數',ISNULL(BOMMD.MD008,0) AS '損耗率%'
                        ,CASE WHEN ISNULL(BOMMD.MD006,0)>0 THEN ISNULL(BOMMD.MD006,0)/ISNULL(BOMMD.MD007,0)*(1+ISNULL(BOMMD.MD008,0)) ELSE 0 END AS 'BOM用量'
                        ,標準批量
                        FROM (
	                        SELECT CMSMD.MD002 AS '線別',TA001 AS '製令單別',TA002 AS '製令單號',TA003 AS '開單日期',TA006 AS '產品品號',TA034 AS '產品品名',TA015 AS '預計產量',TA007 AS '單位1',TB003 AS '材料品號',TB012 AS '材料品名',TB007 AS '單位2',TB004 AS '需領用量',MC004 AS '標準批量',ROUND(TA015/MC004,3) AS '總桶數'
	                        ,(CASE WHEN TA015>0 AND MC004>0 THEN  FLOOR(ROUND(TA015/MC004,3)) ELSE 0 END) AS '整桶數'
	                        ,(CASE WHEN TA015>0 AND MC004>0 THEN (ROUND(TA015/MC004,3)-FLOOR(ROUND(TA015/MC004,3)))   ELSE 0 END) AS '最後桶數'
	                        ,(CASE WHEN TA015>0 AND MC004>0 THEN (CASE WHEN FLOOR(ROUND(TA015/MC004,3)) = ROUND(TA015/MC004,3) THEN ROUND(TB004/ROUND(TA015/MC004,3),3) ELSE CASE WHEN (ROUND(TA015/MC004,3)-1)>0 THEN ROUND(TB004/ROUND(TA015/MC004,3),3) ELSE 0 END  END )  ELSE 0 END)  AS '整桶用量'
	                        ,(CASE WHEN TA015>0 AND MC004>0 THEN (TB004-(ROUND(TB004/ROUND(TA015/MC004,3),3)*(CASE WHEN FLOOR(ROUND(TA015/MC004,3)) = ROUND(TA015/MC004,3) THEN ROUND(TA015/MC004,3) ELSE CASE WHEN (ROUND(TA015/MC004,3)-1)>0 THEN FLOOR(ROUND(TA015/MC004,3)) ELSE 0 END  END)))  ELSE 0 END)   AS '最後桶用量'
	                        ,(CASE WHEN TA015>0 AND MC004>0 THEN ROUND(TB004/ROUND(TA015/MC004,3),3) ELSE 0 END )  AS '標準用量'
	                        FROM [TK].dbo.MOCTA,[TK].dbo.MOCTB,[TK].dbo.BOMMC,[TK].dbo.CMSMD
	                        WHERE TA001=TB001 AND TA002=TB002
	                        AND TA006=MC001
	                        AND TA021=CMSMD.MD001
	                        AND CMSMD.MD002  IN (SELECT [MD002]  FROM [TKWAREHOUSE].[dbo].[TB_MOCTATB_CANNO_NUMS_CMSMD]) 	                       
                        ) AS TEMP
                        LEFT JOIN  [TK].dbo.BOMMD ON 產品品號=MD001 AND 材料品號=MD003
                        WHERE  MD003 LIKE '1%'
                        AND MD003 NOT LIKE '30100002%'
                        AND MD003 NOT IN ( SELECT [MB001]  FROM [TKWAREHOUSE].[dbo].[TB_MOCTATB_CANNO_NUMS_NOUSED] )
                        AND 製令單別=@製令單別 AND 製令單號=@製令單號
                        ORDER BY 線別,製令單別,製令單號  
                        ");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.AddWithValue("@製令單別", 製令單別);
            command.Parameters.AddWithValue("@製令單號", 製令單號);
            DataTable dt = new DataTable();
            try
            {
                connection.Open();
                SqlDataAdapter da = new SqlDataAdapter(command);
                da.Fill(dt);
                dataGridView2.DataSource = dt;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally
            {
                connection.Close();
            }
        }

        public void PRINTS(string TA001,string TA002)
        {
            //DataTable DT = FIND_DETAILS(TA001, TA002);
            //if(DT!=null && DT.Rows.Count > 0)
            //{
            //    ADD_TB_MOCTATB_CANNO_NUMS(DT);

            //    SETFASTREPORT();
            //}
        }

        public DataTable FIND_DETAILS(string TA001, string TA002)
        {
            DataTable DT = new DataTable();

            Class1 TKID = new Class1();//用new 建立類別實體
            // 1.取得原始的連線字串
            string originalConnString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;

            // 2. 使用 SqlConnectionStringBuilder 來解析與修改
            SqlConnectionStringBuilder builder = new SqlConnectionStringBuilder(originalConnString);

            // 3. 將內含的帳密解密後重新指派
            builder.UserID = TKID.Decryption(builder.UserID);
            builder.Password = TKID.Decryption(builder.Password);

            // 4. 用最後組合好的連線字串建立 SqlConnection
            SqlConnection connection = new SqlConnection(builder.ConnectionString);

            StringBuilder sb = new StringBuilder();
            sb.Append(@"                       
                        SELECT 
                        線別,製令單別,製令單號,開單日期,產品品號,產品品名,預計產量,單位1,材料品號,材料品名,單位2,需領用量,標準批量,總桶數,整桶數,最後桶數,整桶用量,最後桶用量,標準用量
                        ,ISNULL(BOMMD.MD004,'') AS '材料單位',ISNULL(BOMMD.MD006,0) AS '組成用量',ISNULL(BOMMD.MD007,0) AS '底數',ISNULL(BOMMD.MD008,0) AS '損耗率%'
                        ,CASE WHEN ISNULL(BOMMD.MD006,0)>0 THEN ISNULL(BOMMD.MD006,0)/ISNULL(BOMMD.MD007,0)*(1+ISNULL(BOMMD.MD008,0)) ELSE 0 END AS 'BOM用量'
                        FROM (
	                        SELECT CMSMD.MD002 AS '線別',TA001 AS '製令單別',TA002 AS '製令單號',TA003 AS '開單日期',TA006 AS '產品品號',TA034 AS '產品品名',TA015 AS '預計產量',TA007 AS '單位1',TB003 AS '材料品號',TB012 AS '材料品名',TB007 AS '單位2',TB004 AS '需領用量',MC004 AS '標準批量',ROUND(TA015/MC004,3) AS '總桶數'
	                        ,(CASE WHEN TA015>0 AND MC004>0 THEN  FLOOR(ROUND(TA015/MC004,3)) ELSE 0 END) AS '整桶數'
	                        ,(CASE WHEN TA015>0 AND MC004>0 THEN (ROUND(TA015/MC004,3)-FLOOR(ROUND(TA015/MC004,3)))   ELSE 0 END) AS '最後桶數'
	                        ,(CASE WHEN TA015>0 AND MC004>0 THEN (CASE WHEN FLOOR(ROUND(TA015/MC004,3)) = ROUND(TA015/MC004,3) THEN ROUND(TB004/ROUND(TA015/MC004,3),3) ELSE CASE WHEN (ROUND(TA015/MC004,3)-1)>0 THEN ROUND(TB004/ROUND(TA015/MC004,3),3) ELSE 0 END  END )  ELSE 0 END)  AS '整桶用量'
	                        ,(CASE WHEN TA015>0 AND MC004>0 THEN (TB004-(ROUND(TB004/ROUND(TA015/MC004,3),3)*(CASE WHEN FLOOR(ROUND(TA015/MC004,3)) = ROUND(TA015/MC004,3) THEN ROUND(TA015/MC004,3) ELSE CASE WHEN (ROUND(TA015/MC004,3)-1)>0 THEN FLOOR(ROUND(TA015/MC004,3)) ELSE 0 END  END)))  ELSE 0 END)   AS '最後桶用量'
	                        ,(CASE WHEN TA015>0 AND MC004>0 THEN ROUND(TB004/ROUND(TA015/MC004,3),3) ELSE 0 END )  AS '標準用量'
	                        FROM [TK].dbo.MOCTA,[TK].dbo.MOCTB,[TK].dbo.BOMMC,[TK].dbo.CMSMD
	                        WHERE TA001=TB001 AND TA002=TB002
	                        AND TA006=MC001
	                        AND TA021=CMSMD.MD001
	                        AND CMSMD.MD002  IN (SELECT [MD002]  FROM [TKWAREHOUSE].[dbo].[TB_MOCTATB_CANNO_NUMS_CMSMD]) 
	                       
                        ) AS TEMP
                        LEFT JOIN  [TK].dbo.BOMMD ON 產品品號=MD001 AND 材料品號=MD003
                        WHERE  MD003 LIKE '1%'
                        AND MD003 NOT LIKE '30100002%'
                        AND MD003 NOT IN ( SELECT [MB001]  FROM [TKWAREHOUSE].[dbo].[TB_MOCTATB_CANNO_NUMS_NOUSED] )
                        AND 製令單別=@製令單別 AND 製令單號=@製令單號
                        ORDER BY 線別,製令單別,製令單號  
                        ");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.AddWithValue("@製令單別", TA001);
            command.Parameters.AddWithValue("@製令單號", TA002);
            
            try
            {
                connection.Open();
                SqlDataAdapter da = new SqlDataAdapter(command);
                da.Fill(DT);

                if (DT != null && DT.Rows.Count >= 1)
                {
                    return DT;
                }
                else
                {
                    return null;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                return null;
            }
            finally
            {
                connection.Close();
            }

          
        }

        public void ADD_TB_MOCTATB_CANNO_NUMS(DataTable DT)
        {
            if (DT == null || DT.Rows.Count == 0) return;

            Class1 TKID = new Class1();
            string originalConnString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
            SqlConnectionStringBuilder builder = new SqlConnectionStringBuilder(originalConnString);
            builder.UserID = TKID.Decryption(builder.UserID);
            builder.Password = TKID.Decryption(builder.Password);

            string TA001 = DT.Rows[0]["製令單別"].ToString().Trim();
            string TA002 = DT.Rows[0]["製令單號"].ToString().Trim();
            decimal CANS = Convert.ToDecimal(DT.Rows[0]["總桶數"]);

            bool isDecimal = (CANS % 1 != 0);
            int ALLCANS = isDecimal ? Convert.ToInt32(Math.Ceiling(CANS)) : Convert.ToInt32(CANS);

            StringBuilder ADD_SQL = new StringBuilder();

            // 1. 先清除指定單別單號的舊資料 (避免全表清空影響他人)
            ADD_SQL.AppendLine(@" DELETE FROM [TKWAREHOUSE].[dbo].[TB_MOCTATB_CANNO_NUMS] 
                                 ");

            // 3. 根據您的邏輯判斷迴圈次數
            int loopCount = isDecimal ? (ALLCANS - 1) : ALLCANS;

            // 前面的桶數：取 [需領用量] (對應您原程式碼邏輯)
            for (int canNo = 1; canNo <= loopCount; canNo++)
            {
                ADD_SQL.AppendLine($@"
                                    INSERT INTO [TKWAREHOUSE].[dbo].[TB_MOCTATB_CANNO_NUMS] ([TA001],[TA002],[CANNO],[MD003],[MB002],[MB004],[NUMS],[ALLCANS])
                                    SELECT [製令單別], [製令單號], {canNo}, [材料品號], [材料品名], [單位2], [整桶用量], [總桶數]
                                    FROM  
                                            (
                                                SELECT CMSMD.MD002 AS '線別',TA001 AS '製令單別',TA002 AS '製令單號',TA003 AS '開單日期',
                                                       TA006 AS '產品品號',TA034 AS '產品品名',TA015 AS '預計產量',TA007 AS '單位1',
                                                       TB003 AS '材料品號',TB012 AS '材料品名',TB007 AS '單位2',TB004 AS '需領用量',
                                                       MC004 AS '標準批量',ROUND(TA015/MC004,3) AS '總桶數',
                                                       (CASE WHEN TA015>0 AND MC004>0 THEN FLOOR(ROUND(TA015/MC004,3)) ELSE 0 END) AS '整桶數',
                                                       (CASE WHEN TA015>0 AND MC004>0 THEN (ROUND(TA015/MC004,3)-FLOOR(ROUND(TA015/MC004,3))) ELSE 0 END) AS '最後桶數',
                                                       (CASE WHEN TA015>0 AND MC004>0 THEN (CASE WHEN FLOOR(ROUND(TA015/MC004,3)) = ROUND(TA015/MC004,3) THEN ROUND(TB004/ROUND(TA015/MC004,3),3) ELSE CASE WHEN (ROUND(TA015/MC004,3)-1)>0 THEN ROUND(TB004/ROUND(TA015/MC004,3),3) ELSE 0 END END ) ELSE 0 END) AS '整桶用量',
                                                       (CASE WHEN TA015>0 AND MC004>0 THEN (TB004-(ROUND(TB004/ROUND(TA015/MC004,3),3)*(CASE WHEN FLOOR(ROUND(TA015/MC004,3)) = ROUND(TA015/MC004,3) THEN ROUND(TA015/MC004,3) ELSE CASE WHEN (ROUND(TA015/MC004,3)-1)>0 THEN FLOOR(ROUND(TA015/MC004,3)) ELSE 0 END END))) ELSE 0 END) AS '最後桶用量',
                                                       (CASE WHEN TA015>0 AND MC004>0 THEN ROUND(TB004/ROUND(TA015/MC004,3),3) ELSE 0 END ) AS '標準用量'
                                                FROM [TK].dbo.MOCTA
                                                INNER JOIN [TK].dbo.MOCTB ON TA001=TB001 AND TA002=TB002
                                                INNER JOIN [TK].dbo.BOMMC ON TA006=MC001
                                                INNER JOIN [TK].dbo.CMSMD ON TA021=CMSMD.MD001
                                                WHERE CMSMD.MD002 IN (SELECT [MD002] FROM [TKWAREHOUSE].[dbo].[TB_MOCTATB_CANNO_NUMS_CMSMD])
                                            ) AS TEMP
                                    LEFT JOIN [TK].dbo.BOMMD ON [產品品號]=MD001 AND [材料品號]=MD003
                                    WHERE MD003 LIKE '1%' AND MD003 NOT LIKE '30100002%'
                                      AND MD003 NOT IN (SELECT [MB001] FROM [TKWAREHOUSE].[dbo].[TB_MOCTATB_CANNO_NUMS_NOUSED])
                                      AND [製令單別] = @TA001 AND [製令單號] = @TA002;
                                ");
                                    }

            // 最後一桶 (有小數的情況)：取 [最後桶用量]
            if (isDecimal)
            {
                ADD_SQL.AppendLine($@"
                                    INSERT INTO [TKWAREHOUSE].[dbo].[TB_MOCTATB_CANNO_NUMS] ([TA001],[TA002],[CANNO],[MD003],[MB002],[MB004],[NUMS],[ALLCANS])
                                        SELECT [製令單別], [製令單號], {ALLCANS}, [材料品號], [材料品名], [單位2], [最後桶用量], [總桶數]
                                        FROM  
                                            (
                                                SELECT CMSMD.MD002 AS '線別',TA001 AS '製令單別',TA002 AS '製令單號',TA003 AS '開單日期',
                                                       TA006 AS '產品品號',TA034 AS '產品品名',TA015 AS '預計產量',TA007 AS '單位1',
                                                       TB003 AS '材料品號',TB012 AS '材料品名',TB007 AS '單位2',TB004 AS '需領用量',
                                                       MC004 AS '標準批量',ROUND(TA015/MC004,3) AS '總桶數',
                                                       (CASE WHEN TA015>0 AND MC004>0 THEN FLOOR(ROUND(TA015/MC004,3)) ELSE 0 END) AS '整桶數',
                                                       (CASE WHEN TA015>0 AND MC004>0 THEN (ROUND(TA015/MC004,3)-FLOOR(ROUND(TA015/MC004,3))) ELSE 0 END) AS '最後桶數',
                                                       (CASE WHEN TA015>0 AND MC004>0 THEN (CASE WHEN FLOOR(ROUND(TA015/MC004,3)) = ROUND(TA015/MC004,3) THEN ROUND(TB004/ROUND(TA015/MC004,3),3) ELSE CASE WHEN (ROUND(TA015/MC004,3)-1)>0 THEN ROUND(TB004/ROUND(TA015/MC004,3),3) ELSE 0 END END ) ELSE 0 END) AS '整桶用量',
                                                       (CASE WHEN TA015>0 AND MC004>0 THEN (TB004-(ROUND(TB004/ROUND(TA015/MC004,3),3)*(CASE WHEN FLOOR(ROUND(TA015/MC004,3)) = ROUND(TA015/MC004,3) THEN ROUND(TA015/MC004,3) ELSE CASE WHEN (ROUND(TA015/MC004,3)-1)>0 THEN FLOOR(ROUND(TA015/MC004,3)) ELSE 0 END END))) ELSE 0 END) AS '最後桶用量',
                                                       (CASE WHEN TA015>0 AND MC004>0 THEN ROUND(TB004/ROUND(TA015/MC004,3),3) ELSE 0 END ) AS '標準用量'
                                                FROM [TK].dbo.MOCTA
                                                INNER JOIN [TK].dbo.MOCTB ON TA001=TB001 AND TA002=TB002
                                                INNER JOIN [TK].dbo.BOMMC ON TA006=MC001
                                                INNER JOIN [TK].dbo.CMSMD ON TA021=CMSMD.MD001
                                                WHERE CMSMD.MD002 IN (SELECT [MD002] FROM [TKWAREHOUSE].[dbo].[TB_MOCTATB_CANNO_NUMS_CMSMD])
                                            ) AS TEMP
                                        LEFT JOIN [TK].dbo.BOMMD ON [產品品號]=MD001 AND [材料品號]=MD003
                                        WHERE MD003 LIKE '1%' AND MD003 NOT LIKE '30100002%'
                                          AND MD003 NOT IN (SELECT [MB001] FROM [TKWAREHOUSE].[dbo].[TB_MOCTATB_CANNO_NUMS_NOUSED])
                                          AND [製令單別] = @TA001 AND [製令單號] = @TA002;
                                    ");
                                    }

            // 4. 執行資料庫寫入
            using (SqlConnection connection = new SqlConnection(builder.ConnectionString))
            {
                using (SqlCommand cmd = new SqlCommand(ADD_SQL.ToString(), connection))
                {
                    cmd.Parameters.AddWithValue("@TA001", TA001);
                    cmd.Parameters.AddWithValue("@TA002", TA002);
                    cmd.CommandTimeout = 120; // 防止計算大表時超時

                    connection.Open();
                    using (SqlTransaction tran = connection.BeginTransaction())
                    {
                        cmd.Transaction = tran;
                        try
                        {
                            cmd.ExecuteNonQuery();
                            tran.Commit();
                        }
                        catch (Exception ex)
                        {
                            tran.Rollback();
                            throw;
                        }
                    }
                }
            }
        }

        public void SETFASTREPORT(string TA001, string TA002, string currentNo)
        {
            string SQL;
            string SQL1;
            report1 = new Report();
            report1.Load(@"REPORT\製令各桶明細.frx");

            Class1 TKID = new Class1();//用new 建立類別實體
            // 1.取得原始的連線字串
            string originalConnString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;

            // 2. 使用 SqlConnectionStringBuilder 來解析與修改
            SqlConnectionStringBuilder builder = new SqlConnectionStringBuilder(originalConnString);

            // 3. 將內含的帳密解密後重新指派
            builder.UserID = TKID.Decryption(builder.UserID);
            builder.Password = TKID.Decryption(builder.Password);

            // 4. 用最後組合好的連線字串建立 SqlConnection
            SqlConnection connection = new SqlConnection(builder.ConnectionString);

            report1.Dictionary.Connections[0].ConnectionString = connection.ConnectionString;

            TableDataSource Table = report1.GetDataSource("Table") as TableDataSource;
            SQL = SETFASETSQL(TA001, TA002, currentNo);
           
            Table.SelectCommand = SQL;
          
            report1.Preview = previewControl1;
            report1.Show();

        }

        public string SETFASETSQL(string TA001, string TA002, string currentNo)
        {
            StringBuilder FASTSQL = new StringBuilder();
            StringBuilder STRQUERY = new StringBuilder();
            StringBuilder STRQUERYNOTIN = new StringBuilder();


            FASTSQL.AppendFormat(@"    
                                SELECT 
                                [TA001] AS '製令單別'
                                ,[TA002] AS '製令單號'
                                ,[CANNO] AS '桶號'
                                ,[MD003] AS '品號'
                                ,[MB002] AS '品名'
                                ,[MB004] AS '單位'
                                ,[NUMS] AS '數量'
                                ,[ALLCANS] AS '總桶數'
                                FROM [TKWAREHOUSE].[dbo].[TB_MOCTATB_CANNO_NUMS]
                                WHERE 1=1
                                AND ((TA001='{0}' AND TA002='{1}') OR (TA002='{2}'))
                                ORDER BY [TA001],[MD003] ,[CANNO]
                                 ", TA001,TA002, currentNo);


            return FASTSQL.ToString();
        }

        public void SERACH_TB_MOCTATB_CANNO_NUMS_NOUSED()
        {
            Class1 TKID = new Class1();//用new 建立類別實體
            // 1.取得原始的連線字串
            string originalConnString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;

            // 2. 使用 SqlConnectionStringBuilder 來解析與修改
            SqlConnectionStringBuilder builder = new SqlConnectionStringBuilder(originalConnString);

            // 3. 將內含的帳密解密後重新指派
            builder.UserID = TKID.Decryption(builder.UserID);
            builder.Password = TKID.Decryption(builder.Password);

            // 4. 用最後組合好的連線字串建立 SqlConnection
            SqlConnection connection = new SqlConnection(builder.ConnectionString);

            StringBuilder sb = new StringBuilder();
            sb.Append(@" 
                        SELECT 
                        [MB001] AS '品號'
                        ,[MB002] AS '品名'
                        FROM [TKWAREHOUSE].[dbo].[TB_MOCTATB_CANNO_NUMS_NOUSED]

                ");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            //command.Parameters.AddWithValue("@SDATE", SDATE);
            DataTable dt = new DataTable();
            try
            {
                connection.Open();
                SqlDataAdapter da = new SqlDataAdapter(command);
                da.Fill(dt);
                dataGridView3.DataSource = dt;

                dataGridView3.AutoResizeColumns(DataGridViewAutoSizeColumnsMode.AllCellsExceptHeader);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally
            {
                connection.Close();
            }
        }
        private void dataGridView3_SelectionChanged(object sender, EventArgs e)
        {
            if(dataGridView3.CurrentRow != null)
            {
                string MB001 = dataGridView3.CurrentRow.Cells["品號"].Value.ToString();
                string MB002 = dataGridView3.CurrentRow.Cells["品名"].Value.ToString();
               
                textBox3.Text = MB001;
                textBox4.Text = MB002;
            }
        }

        public void ADD_TB_MOCTATB_CANNO_NUMS_NOUSED(string MB001,string MB002)
        {
            Class1 TKID = new Class1();
            string originalConnString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
            SqlConnectionStringBuilder builder = new SqlConnectionStringBuilder(originalConnString);
            builder.UserID = TKID.Decryption(builder.UserID);
            builder.Password = TKID.Decryption(builder.Password);
            StringBuilder ADD_SQL = new StringBuilder();
            ADD_SQL.AppendLine($@"
                                    INSERT INTO [TKWAREHOUSE].[dbo].[TB_MOCTATB_CANNO_NUMS_NOUSED] 
                                    ([MB001],[MB002])
                                    VALUES (@MB001,@MB002);
                                ");
            using (SqlConnection connection = new SqlConnection(builder.ConnectionString))
            {
                using (SqlCommand cmd = new SqlCommand(ADD_SQL.ToString(), connection))
                {
                    cmd.Parameters.AddWithValue("@MB001", MB001);
                    cmd.Parameters.AddWithValue("@MB002", MB002);
                    cmd.CommandTimeout = 60;
                    connection.Open();
                    using (SqlTransaction tran = connection.BeginTransaction())
                    {
                        cmd.Transaction = tran;
                        try
                        {
                            cmd.ExecuteNonQuery();
                            tran.Commit();
                        }
                        catch (Exception ex)
                        {
                            tran.Rollback();
                            MessageBox.Show(ex.Message);
                        }
                    }
                }
            }
        }

        public void DELETE_TB_MOCTATB_CANNO_NUMS_NOUSED(string MB001)
        {            
            Class1 TKID = new Class1();
            string originalConnString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
            SqlConnectionStringBuilder builder = new SqlConnectionStringBuilder(originalConnString);
            builder.UserID = TKID.Decryption(builder.UserID);
            builder.Password = TKID.Decryption(builder.Password);
            StringBuilder DELETE_SQL = new StringBuilder();
            DELETE_SQL.AppendLine($@"
                                    DELETE FROM [TKWAREHOUSE].[dbo].[TB_MOCTATB_CANNO_NUMS_NOUSED] 
                                    WHERE MB001 = @MB001;
                                ");
            using (SqlConnection connection = new SqlConnection(builder.ConnectionString))
            {
                using (SqlCommand cmd = new SqlCommand(DELETE_SQL.ToString(), connection))
                {
                    cmd.Parameters.AddWithValue("@MB001", MB001);
                    cmd.CommandTimeout = 60;
                    connection.Open();
                    using (SqlTransaction tran = connection.BeginTransaction())
                    {
                        cmd.Transaction = tran;
                        try
                        {
                            cmd.ExecuteNonQuery();
                            tran.Commit();
                        }
                        catch (Exception ex)
                        {
                            tran.Rollback();
                            MessageBox.Show(ex.Message);
                        }
                    }
                }
            }
        }

        private void textBox3_TextChanged(object sender, EventArgs e)
        {
            textBox4.Text = null;
            string MB001=textBox3.Text.Trim();
            if (!string.IsNullOrEmpty(MB001))
            {
                textBox4.Text = FIND_MB002(MB001);
            }
        }

        public string FIND_MB002(string MB001)
        {
            string MB002 = null;
            Class1 TKID = new Class1();
            string originalConnString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
            SqlConnectionStringBuilder builder = new SqlConnectionStringBuilder(originalConnString);
            builder.UserID = TKID.Decryption(builder.UserID);
            builder.Password = TKID.Decryption(builder.Password);
            using (SqlConnection connection = new SqlConnection(builder.ConnectionString))
            {
                string query = "SELECT MB002 FROM [TK].dbo.INVMB WHERE MB001 = @MB001";
                using (SqlCommand cmd = new SqlCommand(query, connection))
                {
                    cmd.Parameters.AddWithValue("@MB001", MB001);
                    try
                    {
                        connection.Open();
                        object result = cmd.ExecuteScalar();
                        if (result != null)
                        {
                            MB002 = result.ToString();
                        }
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message);
                    }
                }
            }
            return MB002;
        }

        public string GetNewMergeNo()
        {
            string newMergeNo = "";
            Class1 TKID = new Class1();
            SqlConnectionStringBuilder builder = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);
            builder.UserID = TKID.Decryption(builder.UserID);
            builder.Password = TKID.Decryption(builder.Password);

            string sql = @"
                            DECLARE @Today VARCHAR(8) = CONVERT(VARCHAR(8), GETDATE(), 112);
                            SELECT '合併-'+@Today + RIGHT('000' + CAST(ISNULL(MAX(RIGHT(MERGENO, 3)), 0) + 1 AS VARCHAR), 3) AS NewNo
                            FROM [TKWAREHOUSE].[dbo].[TB_MOCTATB_CANNO_NUMS_MERGE] WITH (UPDLOCK, HOLDLOCK)
                            WHERE MERGENO LIKE '合併-'+@Today + '%';
                        ";

            using (SqlConnection conn = new SqlConnection(builder.ConnectionString))
            {
                using (SqlCommand cmd = new SqlCommand(sql, conn))
                {
                    conn.Open();
                    object result = cmd.ExecuteScalar();
                    if (result != null)
                    {
                        newMergeNo = result.ToString();
                    }
                }
            }
            return newMergeNo;
        }

        public void ADD_TB_MOCTATB_CANNO_NUMS_MERGE(string currentNo)
        {
            StringBuilder ADD_SQL = new StringBuilder();
            // 收集所有被勾選的列
            List<DataGridViewRow> selectedRows = new List<DataGridViewRow>();

            // 這裡修正為您畫面實際對應的 dataGridView1
            foreach (DataGridViewRow row in dataGridView1.Rows)
            {
                if (!row.IsNewRow && Convert.ToBoolean(row.Cells["SelectCheck"].Value) == true)
                {
                    selectedRows.Add(row);
                }
            }

            if (selectedRows.Count == 0)
            {
                MessageBox.Show("請先勾選要合併的製令單！", "提示");
                return;
            }

            SqlCommand cmd = new SqlCommand();

            // 【核心修正點 1】: 共用的編號是固定不變的，在迴圈外「只加入一次」即可！
            cmd.Parameters.AddWithValue("@NewMergeNo", currentNo);

            int paramIndex = 0;
            foreach (DataGridViewRow row in selectedRows)
            {
                string ta001 = row.Cells["製令單別"].Value?.ToString().Trim() ?? "";
                string ta002 = row.Cells["製令單號"].Value?.ToString().Trim() ?? "";

                // 參數化命名避免衝突
                string pTA001 = "@TA001_" + paramIndex;
                string pTA002 = "@TA002_" + paramIndex;

                ADD_SQL.AppendLine($@"
                                    INSERT INTO [TKWAREHOUSE].[dbo].[TB_MOCTATB_CANNO_NUMS_MERGE] ([MERGENO], [TA001], [TA002])
                                    VALUES (@NewMergeNo, {pTA001}, {pTA002});
                                ");

                // 【核心修正點 2】: 迴圈內只加入每筆資料各自專屬的單別、單號參數
                cmd.Parameters.AddWithValue(pTA001, ta001);
                cmd.Parameters.AddWithValue(pTA002, ta002);

                paramIndex++;
            }

            // 取得解密連線字串並執行
            Class1 TKID = new Class1();
            SqlConnectionStringBuilder builder = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);
            builder.UserID = TKID.Decryption(builder.UserID);
            builder.Password = TKID.Decryption(builder.Password);

            using (SqlConnection conn = new SqlConnection(builder.ConnectionString))
            {
                cmd.Connection = conn;
                cmd.CommandText = ADD_SQL.ToString();

                try
                {
                    conn.Open();
                    using (SqlTransaction tran = conn.BeginTransaction())
                    {
                        cmd.Transaction = tran;
                        cmd.ExecuteNonQuery();
                        tran.Commit();
                    }
                    //MessageBox.Show("批次存檔完成！", "成功");

                    // 重新整理 UI 或者是清除勾選
                    SetupDataGridView();
                }
                catch (Exception ex)
                {
                    MessageBox.Show("存檔失敗：" + ex.Message);
                }
            }
        }

        public DataTable CHECK_BOMMD(int COUNTS, List<string> MD001)
        {
            // 安全防護：若沒有傳入任何品號，直接回傳空表
            if (MD001 == null || MD001.Count == 0 || COUNTS < 1)
            {
                return new DataTable();
            }

            DataTable dt = new DataTable();
            StringBuilder SQL = new StringBuilder();
            int SETCOUNT = 1;

            try
            {
                Class1 TKID = new Class1();
                string originalConnString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
                SqlConnectionStringBuilder builder = new SqlConnectionStringBuilder(originalConnString);
                builder.UserID = TKID.Decryption(builder.UserID);
                builder.Password = TKID.Decryption(builder.Password);

                // 1. 開始動態組裝 SQL
                SQL.Append(@" SELECT MD003, MD006, MD007, COUNT(MD003) AS COUNTS 
                      FROM ( ");

                foreach (string MD001STR in MD001)
                {
                    if (SETCOUNT > 1)
                    {
                        SQL.Append(" UNION ALL ");
                    }

                    // 【修正點】將原本分散在 if/else 外的括號與語法整理乾淨
                    SQL.AppendFormat(@" SELECT MD001, MD003, MD006, MD007
                                FROM [TK].dbo.BOMMD
                                WHERE MD003 LIKE '1%' AND MD001 = '{0}' ", MD001STR.Trim());

                    SETCOUNT++;
                }

                // 結束子查詢，並加上您的排除與過濾條件
                SQL.AppendFormat(@" ) AS CombinedData
                            WHERE MD003 NOT IN (SELECT [MD003] FROM [TKMOC].[dbo].[REPORTMOCBOMNOSET])
                            GROUP BY MD003, MD006, MD007
                            HAVING COUNT(MD003) < {0} ", COUNTS);

                // 2. 建立連線與透過 Adapter 填充 DataTable
                using (SqlConnection connection = new SqlConnection(builder.ConnectionString))
                {
                    using (SqlCommand cmd = new SqlCommand(SQL.ToString(), connection))
                    {
                        cmd.CommandTimeout = 60;

                        // 使用 SqlDataAdapter 才能將資料庫查詢到的多筆資料整張倒進 DataTable
                        using (SqlDataAdapter da = new SqlDataAdapter(cmd))
                        {
                            connection.Open();
                            da.Fill(dt); // 填充資料
                            return dt;   // 成功回傳含有資料的 DataTable
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                // 建議此處可記錄 log：Elmah.ErrorSignal.FromCurrentContext().Raise(ex);
                return null;
            }
        }

        public void PRINT_MERGE(string TA001, string TA002, float BUCKETS, string LINK_TA001TA002, string LINK_TA006, string LINK_TA034, string MAINMB001,string currentNo)
        {
            if (BUCKETS <= 0) return;

            // 1. 處理報表中間資料 (帶入製令單別與單號)
            ProcessReportData(TA001.Trim(), TA002.Trim(), BUCKETS, currentNo);

            // 2. 加載報表與設置資料源
            SETFASTREPORT(TA001.Trim(), TA002.Trim(), currentNo);
        }

        private void ProcessReportData(string TA001, string TA002, float buckets, string currentNo)
        {
            int totalCounts = (int)Math.Ceiling(buckets);
            bool isInteger = (buckets % 1 == 0);

            // 防止傳入空參數時不小心清空或洗掉非預期的資料
            if (string.IsNullOrEmpty(TA001) || string.IsNullOrEmpty(TA002)) return;

            StringBuilder batchSql = new StringBuilder();
            SqlCommand cmd = new SqlCommand();

            // 1. 先清除舊報表資料 (【核心修正】: 務必加上 WHERE 條件，否則會清空全表！)
            batchSql.AppendLine("DELETE [TKWAREHOUSE].[dbo].[TB_MOCTATB_CANNO_NUMS] WHERE TA001=@ParamTA001 AND TA002=@ParamTA002");

            // 2. 準備動態組裝多個安全選取的資料集
            List<string> selectStatements = new List<string>();

            for (int i = 1; i <= totalCounts; i++)
            {
                string multiplierParam = $"@Multiplier_{i}";
                decimal multiplierValue = 1.0m;

                // 判斷是否為最後一桶且有小數
                if (!isInteger && i == totalCounts)
                {
                    multiplierValue = Convert.ToDecimal(buckets - (totalCounts - 1));
                }

                // 將每桶的乘數作為參數加入 SqlCommand
                cmd.Parameters.AddWithValue(multiplierParam, multiplierValue);

                // 【核心修正點】：補上 C.MB004 後方與 CONVERT 後方的「逗號」
                selectStatements.Add($@"
                                    SELECT @ParamTA001, @currentNo, {i}, B.MD003, C.MB002, C.MB004, 
                                           CONVERT(DECIMAL(16,3), (B.MD006 / B.MD007) * {multiplierParam}), @ALLCANS
                                    FROM [TK].dbo.MOCTA A
                                    INNER JOIN [TK].dbo.BOMMD B ON A.TA006 = B.MD001
                                    INNER JOIN [TK].dbo.INVMB C ON B.MD003 = C.MB001
                                    WHERE 1=1
                                    AND B.MD003 LIKE '1%' AND B.MD003 NOT LIKE '30100002%'
                                    AND B.MD003 NOT IN (SELECT [MB001] FROM [TKWAREHOUSE].[dbo].[TB_MOCTATB_CANNO_NUMS_NOUSED])
                                    AND A.TA001 = @ParamTA001 
                                    AND A.TA002 = @ParamTA002 
                                    ");
            }

            // 3. 串接整段 SQL 語法
            if (selectStatements.Count > 0)
            {
                batchSql.AppendLine("INSERT INTO [TKWAREHOUSE].[dbo].[TB_MOCTATB_CANNO_NUMS] ([TA001],[TA002],[CANNO],[MD003],[MB002],[MB004],[NUMS],[ALLCANS])");
                batchSql.AppendLine(string.Join("\n UNION ALL \n", selectStatements));
            }

            // 4. 綁定核心的製令單別、單號與總桶數參數
            cmd.Parameters.AddWithValue("@ParamTA001", TA001.Trim());
            cmd.Parameters.AddWithValue("@ParamTA002", TA002.Trim());
            cmd.Parameters.AddWithValue("@ALLCANS", buckets);
            cmd.Parameters.AddWithValue("@currentNo",  currentNo);

            // 5. 執行資料庫交易
            ExecuteSqlTransaction(batchSql.ToString(), cmd);
        }

        private void ExecuteSqlTransaction(string sqlCommandText, SqlCommand cmd)
        {
            var builder = GetDecryptedConnBuilder();
            using (SqlConnection conn = new SqlConnection(builder.ConnectionString))
            {
                conn.Open();
                using (SqlTransaction transaction = conn.BeginTransaction())
                {
                    try
                    {
                        cmd.Connection = conn;
                        cmd.Transaction = transaction;
                        cmd.CommandText = sqlCommandText;
                        cmd.CommandTimeout = 90;

                        cmd.ExecuteNonQuery();
                        transaction.Commit();
                    }
                    catch (Exception ex)
                    {
                        transaction.Rollback();
                        throw new Exception("報表暫存資料寫入失敗: " + ex.Message);
                    }
                }
            }
        }

        private SqlConnectionStringBuilder GetDecryptedConnBuilder()
        {
            Class1 TKID = new Class1();
            var sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);
            sqlsb.Password = TKID.Decryption(sqlsb.Password);
            sqlsb.UserID = TKID.Decryption(sqlsb.UserID);
            return sqlsb;
        }
        #endregion

        #region BUTTON
        private void button1_Click(object sender, EventArgs e)
        {
            string SDATE= dateTimePicker1.Value.ToString("yyyyMMdd");
            SEARCH(SDATE);
        }

        private void button2_Click(object sender, EventArgs e)
        {
            TA001 = textBox1.Text;
            TA002 = textBox2.Text;

            if (!string.IsNullOrEmpty(TA001)&& !string.IsNullOrEmpty(TA002))
            {
                PRINTS(TA001, TA002);
            }
            else
            {
                MessageBox.Show("沒有選製令");
            }
        }
        private void button3_Click(object sender, EventArgs e)
        {
            SERACH_TB_MOCTATB_CANNO_NUMS_NOUSED();
        }

        private void button4_Click(object sender, EventArgs e)
        {
            string MB001 = textBox3.Text.Trim();

            DialogResult dialogResult = MessageBox.Show($"確定要刪除 品號 {MB001} 嗎？", "確認", MessageBoxButtons.YesNo);
            if (dialogResult == DialogResult.Yes)
            {
                DELETE_TB_MOCTATB_CANNO_NUMS_NOUSED(MB001);
                SERACH_TB_MOCTATB_CANNO_NUMS_NOUSED(); // 刷新列表

                MessageBox.Show("完成");
            }
        }

        private void button5_Click(object sender, EventArgs e)
        {
            string MB001 = textBox3.Text.Trim();
            string MB002 = textBox4.Text.Trim();

            ADD_TB_MOCTATB_CANNO_NUMS_NOUSED(MB001, MB002);
            SERACH_TB_MOCTATB_CANNO_NUMS_NOUSED(); // 刷新列表
            MessageBox.Show("完成");
        }

        private void button6_Click(object sender, EventArgs e)
        {
            int COUNTS = 0;
            List<string> MD001 = new List<string>();

            // 儲存勾選列的暫存結構，避免重複讀取 UI
            var selectedRows = new List<DataGridViewRow>();

            string LINK_TA001TA002 = "";
            string LINK_TA006 = "";
            string LINK_TA034 = "";
            float BUCKETS = 0;

            string lastTA001 = "";
            string lastTA002 = "";

            // 1. 【優化：合併唯一個迴圈】收集所有勾選的資料
            foreach (DataGridViewRow dr in this.dataGridView1.Rows)
            {
                if (dr.IsNewRow) continue;

                // 確保 CheckBox 有被勾選
                bool isChecked = Convert.ToBoolean(dr.Cells[0].Value);
                if (isChecked)
                {
                    COUNTS++;
                    selectedRows.Add(dr);

                    // 收集品號供 CHECK_BOMMD 檢查
                    string prodNo = dr.Cells["產品品號"].Value?.ToString().Trim() ?? "";
                    if (!string.IsNullOrEmpty(prodNo))
                    {
                        MD001.Add(prodNo);
                    }

                    // 串接合併資訊
                    lastTA001 = dr.Cells["製令單別"].Value?.ToString().Trim() ?? "";
                    lastTA002 = dr.Cells["製令單號"].Value?.ToString().Trim() ?? "";

                    LINK_TA001TA002 += lastTA001 + lastTA002 + "*";
                    LINK_TA006 += prodNo + "*"; // 【修正】原本錯寫成 LINK_TA034
                    LINK_TA034 += dr.Cells["產品品名"].Value?.ToString().Trim() + "*";

                    BUCKETS += Convert.ToSingle(dr.Cells["總桶數"].Value ?? 0);
                }
            }

            // 防呆：如果根本沒勾選，直接結束
            if (COUNTS == 0)
            {
                MessageBox.Show("請至少勾選一筆資料！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                // 如果原本 N 的邏輯需要跑：SETREPORT(textBox1.Text.Trim(), textBox2.Text.Trim(), textBox3.Text.Trim(), MAINMB001);
                return;
            }

            // 四捨五入總桶數
            BUCKETS = (float)Math.Round(BUCKETS, 3);

            // 記錄最後一筆的品號供後續使用 (對應您原本的 MAINMB001)
            if (MD001.Count > 0)
            {
                MAINMB001 = MD001.Last();
            }

            // 2. CHECK：檢查原料單身品號與用量是否一致
            DataTable DT = CHECK_BOMMD(COUNTS, MD001);

            // 3. 判斷檢查結果
            if (DT == null || DT.Rows.Count == 0)
            {
                // 【核心修正】：檢查完全通過了，才去資料庫取號並寫入 Merge 表
                string currentNo = GetNewMergeNo();

                // 執行批次存檔
                ADD_TB_MOCTATB_CANNO_NUMS_MERGE(currentNo);

                // 執行列印或預覽
                PRINT_MERGE(lastTA001, lastTA002, BUCKETS, LINK_TA001TA002, LINK_TA006, LINK_TA034, MAINMB001, currentNo);
            }
            else
            {
                // 檢查失敗，顯示不一致的原料明細
                StringBuilder sbError = new StringBuilder();
                sbError.AppendLine("原料需單身品號元件不一致 或 組成用量不一致，不能合併：");

                foreach (DataRow row in DT.Rows)
                {
                    sbError.AppendLine($"品號: {row["MD003"]} | 用量: {row["MD006"]} | 底數: {row["MD007"]}");
                }

                MessageBox.Show(sbError.ToString(), "核心BOM檢查失敗", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

        }

        private void button7_Click(object sender, EventArgs e)
        {
            string SDATE = dateTimePicker2.Value.ToString("yyyyMMdd");
            SEARCH_DG4(SDATE);
        }

        private void button9_Click(object sender, EventArgs e)
        {

        }
        #endregion


    }
}

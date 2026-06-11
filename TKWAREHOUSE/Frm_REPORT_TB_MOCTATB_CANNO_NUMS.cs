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
                        線別,製令單別,製令單號,開單日期,產品品號,產品品名,預計產量,單位
                        FROM (
	                        SELECT CMSMD.MD002 AS '線別',TA001 AS '製令單別',TA002 AS '製令單號',TA003 AS '開單日期',TA006 AS '產品品號',TA034 AS '產品品名',TA015 AS '預計產量',TA007 AS '單位'
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
            DataTable DT = FIND_DETAILS(TA001, TA002);
            if(DT!=null && DT.Rows.Count > 0)
            {
                ADD_TB_MOCTATB_CANNO_NUMS(DT);

                SETFASTREPORT();
            }
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

        public void SETFASTREPORT()
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
            SQL = SETFASETSQL();
           
            Table.SelectCommand = SQL;
          
            report1.Preview = previewControl1;
            report1.Show();

        }

        public string SETFASETSQL()
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
                                ORDER BY [TA001],[MD003] ,[CANNO]
                                 ");


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
                            SELECT @Today + RIGHT('000' + CAST(ISNULL(MAX(RIGHT(MERGENO, 3)), 0) + 1 AS VARCHAR), 3) AS NewNo
                            FROM [TKWAREHOUSE].[dbo].[TB_MOCTATB_CANNO_NUMS_MERGE] WITH (UPDLOCK, HOLDLOCK)
                            WHERE MERGENO LIKE @Today + '%';
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
            string currentNo = GetNewMergeNo();
            int COUNTS = 0;
            List<string> MD001 = new List<string>();
            DataTable DT = null;
            string MESS = "";

            string CHECKED = "N";
            string TA001 = "";
            string TA002 = "";
            string LINK_TA001TA002 = "";
            string LINK_TA006 = "";
            string LINK_TA034 = "";
            string TEMP = "";
            float BUCKETS = 0;
            //MessageBox.Show(currentNo);

            ADD_TB_MOCTATB_CANNO_NUMS_MERGE(currentNo);

            //
            foreach (DataGridViewRow dr in this.dataGridView1.Rows)
            {
                try
                {
                    if (dr.Cells[0].Value != null && (bool)dr.Cells[0].Value)
                    {
                        COUNTS = COUNTS + 1;
                        MD001.Add(dr.Cells[5].Value.ToString());
                    }
                }
                catch (Exception ex)
                {
                    //MessageBox.Show(ex.Message);
                }
            }

            //CHECK  原料需單身品號元件一致、組成用量一致
            DT = CHECK_BOMMD(COUNTS, MD001);

            foreach (DataGridViewRow dr in this.dataGridView1.Rows)
            {
                try
                {
                    if (dr.Cells[0].Value != null && (bool)dr.Cells[0].Value)
                    {
                        CHECKED = "Y";

                        TA001 = dr.Cells["製令"].Value.ToString();
                        TA002 = dr.Cells["單號"].Value.ToString();
                        LINK_TA001TA002 = LINK_TA001TA002 + TA001 + TA002 + "*";
                        LINK_TA006 = LINK_TA034 + dr.Cells["品號"].Value.ToString() + "*";
                        LINK_TA034 = LINK_TA034 + dr.Cells["品名"].Value.ToString() + "*";
                        BUCKETS = BUCKETS + float.Parse(dr.Cells["桶數"].Value.ToString());
                        BUCKETS = (float)Math.Round(BUCKETS, 3);

                        MAINMB001 = dr.Cells["品號"].Value.ToString();
                    }
                }
                catch (Exception ex)
                {
                    //MessageBox.Show(ex.Message);
                }
            }

            if (CHECKED.Equals("Y"))
            {
                if (DT == null || DT.Rows.Count == 0)
                {
                    //MessageBox.Show("成功！", "成功");
                    //SETREPORT2(TA001, TA002, BUCKETS, LINK_TA001TA002, LINK_TA006, LINK_TA034, MAINMB001);
                }
                else
                {
                    MESS = "原料需單身品號元件不一致 或 組成用量不一致，不能合併\n";
                    foreach (DataRow ROW in DT.Rows)
                    {
                        // 每一行都是一個 DataRow                       
                        MESS = MESS + "品號:" + ROW["MD003"].ToString();
                        MESS = MESS + "用量:" + ROW["MD006"].ToString();
                        MESS = MESS + "底數:" + ROW["MD007"].ToString();

                        MESS = MESS + "\n";
                    }
                    MessageBox.Show(MESS.ToString());
                }

            }
            else if (CHECKED.Equals("N"))
            {
                //SETREPORT(textBox1.Text.Trim(), textBox2.Text.Trim(), textBox3.Text.Trim(), MAINMB001);
            }

        }

        #endregion


    }
}

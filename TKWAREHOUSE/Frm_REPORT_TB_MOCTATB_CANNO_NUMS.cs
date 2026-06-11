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

        public Frm_REPORT_TB_MOCTATB_CANNO_NUMS()
        {
            InitializeComponent();
        }
        private void Frm_REPORT_TB_MOCTATB_CANNO_NUMS_Load(object sender, EventArgs e)
        {

        }
        #region FUNCTION
        
            
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
                                 ");


            return FASTSQL.ToString();
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
        #endregion


    }
}

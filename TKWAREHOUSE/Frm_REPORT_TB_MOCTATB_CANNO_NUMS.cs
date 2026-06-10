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
using NPOI.SS.UserModel;
using System.Configuration;
using NPOI.XSSF.UserModel;
using NPOI.SS.Util;
using System.Reflection;
using System.Threading;
using System.Globalization;
using Calendar.NET;
using TKITDLL;

namespace TKWAREHOUSE
{
    public partial class Frm_REPORT_TB_MOCTATB_CANNO_NUMS : Form
    {
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

            if(dataGridView1.CurrentRow != null)
            {
                製令單別 = dataGridView1.CurrentRow.Cells["製令單別"].Value.ToString();
                製令單號 = dataGridView1.CurrentRow.Cells["製令單號"].Value.ToString();
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
	                        AND TA003>='20260610' AND TA003<='20260610'
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
        #endregion

        #region BUTTON
        private void button1_Click(object sender, EventArgs e)
        {
            string SDATE= dateTimePicker1.Value.ToString("yyyyMMdd");
            SEARCH(SDATE);
        }

        #endregion

      
    }
}

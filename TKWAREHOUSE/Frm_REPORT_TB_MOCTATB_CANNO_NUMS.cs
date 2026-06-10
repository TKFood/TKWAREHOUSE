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

using AForge.Video;
using AForge.Video.DirectShow;
using FastReport;
using FastReport.Data;
using FastReport.DevComponents.DotNetBar.Controls;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration;
using System.Data;
using System.Data.OleDb;
using System.Data.SqlClient;
using System.Drawing;
using System.IO;
using System.IO.Ports;
using System.IO.Ports;
using System.Linq;
using System.Net;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using TKITDLL;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;

namespace TKWAREHOUSE
{
    public partial class FrmPURTHADJSUTLOTNO : Form
    {
        StringBuilder sbSql = new StringBuilder();
        SqlConnection sqlConn = new SqlConnection();
        SqlDataAdapter adapter = new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilder = new SqlCommandBuilder();
        SqlTransaction tran;
        SqlCommand cmd = new SqlCommand();       
        DataSet dt = new DataSet();
        string tablename = null;
        int rownum = 0;
        DataGridViewRow row;
        int result;

        // 存储旧值用于审计
        private string old_TH010 = "";
        private string old_TH117 = "";
        private string old_TH036 = "";

        public FrmPURTHADJSUTLOTNO()
        {
            InitializeComponent();
        }
        private void FrmPURTHADJSUTLOTNO_Load(object sender, EventArgs e)
        {

        }
        #region FUNCTION
        public void SEARCH(string TH002)
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

                using (SqlConnection conn = sqlConn)
                {
                    sbSql.Clear();
                    sbSql.AppendFormat(@"
                                        SELECT 
                                        TH001 AS '單別'
                                        ,TH002 AS '單號'
                                        ,TH003 AS '序號'
                                        ,TH004 AS '品號'
                                        ,TH005 AS '品名'
                                        ,TH010 AS '批號'
                                        ,TH117 AS '製造日期'
                                        ,TH036 AS '有效日期'

                                        FROM [TK].dbo.PURTH
                                        WHERE TH002 LIKE '%{0}%'
                                        ORDER BY TH001,TH002,TH003

                                        ", TH002);
                    adapter = new SqlDataAdapter(@"" + sbSql, conn);
                    sqlCmdBuilder = new SqlCommandBuilder(adapter);
                    DataSet ds = new DataSet();
                    adapter.Fill(ds, "TEMPds");
                    dataGridView1.DataSource = ds.Tables["TEMPds"];
                }

            }
            catch (Exception ex)
            {
                //MessageBox.Show(ex.Message);
            }
            finally
            {
                sqlConn.Close();
            }
            

        }
        private void dataGridView1_SelectionChanged(object sender, EventArgs e)
        {
            SET_TEXT_NULL();

            if (dataGridView1.CurrentRow != null)
            {
                int rowindex = dataGridView1.CurrentRow.Index;
                if (rowindex >= 0)
                {
                    DataGridViewRow row = dataGridView1.Rows[rowindex];

                    textBox2.Text = row.Cells["單別"].Value.ToString().Trim();
                    textBox3.Text = row.Cells["單號"].Value.ToString().Trim();
                    textBox4.Text = row.Cells["序號"].Value.ToString().Trim();
                    textBox5.Text = row.Cells["品號"].Value.ToString().Trim();
                    textBox6.Text = row.Cells["品名"].Value.ToString().Trim();
                    textBox7.Text = row.Cells["批號"].Value.ToString().Trim();
                    textBox8.Text = row.Cells["製造日期"].Value.ToString().Trim();
                    textBox9.Text = row.Cells["有效日期"].Value.ToString().Trim();

                    // 记录旧值用于后续审计
                    old_TH010 = row.Cells["批號"].Value.ToString().Trim();
                    old_TH117 = row.Cells["製造日期"].Value.ToString().Trim();
                    old_TH036 = row.Cells["有效日期"].Value.ToString().Trim();

                    //SEARCH2(row.Cells["品號"].Value.ToString().Trim());
                    //SEARCH3(row.Cells["品號"].Value.ToString().Trim());

                    //SETFASTREPORT(row.Cells["品號"].Value.ToString().Trim());
                }
            }
        }

        //PURTH(R/W) 進貨單單身檔
        //INVME(R/W) 料件批號資料頭身
        //INVMF(R/W) 料件批號資料單身
        //INVLA(R/W) 異動明細資料檔
        public void UPDATE_PURTH_INVLA_INVME(
            string TH001
            , string TH002
            , string TH003
            , string TH004
            , string TH010
            , string TH117
            , string TH036
            , string old_TH010
            )
        {
            try
            {
                //20210902密
                Class1 TKID = new Class1();
                SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

                //資料庫使用者密碼解密
                sqlsb.Password = TKID.Decryption(sqlsb.Password);
                sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

                using (SqlConnection conn = new SqlConnection(sqlsb.ConnectionString))
                {
                    conn.Open();
                    using (SqlTransaction tran = conn.BeginTransaction())
                    {
                        using (SqlCommand cmd = new SqlCommand())
                        {
                            cmd.Connection = conn;
                            cmd.Transaction = tran;
                            cmd.CommandTimeout = 60;
                            cmd.CommandText = @"
                                                UPDATE [TK].dbo.PURTH
                                                SET TH010=@TH010,TH117=@TH117,TH036=@TH036
                                                WHERE TH001=@TH001 AND TH002=@TH002 AND TH003=@TH003 AND TH004=@TH004

                                                UPDATE  [TK].dbo.INVME
                                                SET ME002=@ME002,ME009=@ME009,ME032=@ME032
                                                WHERE ME001=@ME001 AND ME002=@old_TH010

                                                UPDATE  [TK].dbo.INVMF
                                                SET MF002=@MF002
                                                WHERE MF001=@MF001 AND MF002=@old_TH010

                                                UPDATE [TK].dbo.INVLA
                                                SET LA016=@LA016
                                                WHERE LA006=@LA006 AND LA007=@LA007 AND LA008=@LA008 AND LA001=@LA001
                                                ";
                            cmd.Parameters.AddWithValue("@TH001", TH001);
                            cmd.Parameters.AddWithValue("@TH002", TH002);
                            cmd.Parameters.AddWithValue("@TH003", TH003);
                            cmd.Parameters.AddWithValue("@TH004", TH004);
                            cmd.Parameters.AddWithValue("@TH010", TH010);
                            cmd.Parameters.AddWithValue("@TH117", TH117);
                            cmd.Parameters.AddWithValue("@TH036", TH036);

                            cmd.Parameters.AddWithValue("@old_TH010", old_TH010);
                            cmd.Parameters.AddWithValue("@ME001", TH004);
                            cmd.Parameters.AddWithValue("@ME002", TH010);
                            cmd.Parameters.AddWithValue("@ME032", TH117);
                            cmd.Parameters.AddWithValue("@ME009", TH036);

                            cmd.Parameters.AddWithValue("@MF002", TH010);
                            cmd.Parameters.AddWithValue("@MF001", TH004);                            

                            cmd.Parameters.AddWithValue("@LA016", TH010);
                            cmd.Parameters.AddWithValue("@LA006", TH001);
                            cmd.Parameters.AddWithValue("@LA007", TH002);
                            cmd.Parameters.AddWithValue("@LA008", TH003);
                            cmd.Parameters.AddWithValue("@LA001", TH004);

                            int result = cmd.ExecuteNonQuery();

                            if (result > 0)
                            {
                                tran.Commit();

                                ////保存審計日誌（使用已記錄的舊值）
                                //SAVE_AUDIT_LOG(TH001, TH002, TH003, TH004,
                                //    old_TH010, old_TH117, old_TH036,
                                //    TH010, TH117, TH036);

                                MessageBox.Show("更新成功");
                            }
                            else
                            {
                                tran.Rollback();
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("更新失敗: " + ex.Message);
            }
        }

        //private void SAVE_AUDIT_LOG(string TH001, string TH002, string TH003, string TH004,
        //    string OLD_TH010, string OLD_TH117, string OLD_TH036,
        //    string NEW_TH010, string NEW_TH117, string NEW_TH036)
        //{
        //    try
        //    {
        //        Class1 TKID = new Class1();
        //        SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

        //        sqlsb.Password = TKID.Decryption(sqlsb.Password);
        //        sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

        //        using (SqlConnection conn = new SqlConnection(sqlsb.ConnectionString))
        //        {
        //            conn.Open();
        //            using (SqlCommand cmd = new SqlCommand())
        //            {
        //                cmd.Connection = conn;
        //                cmd.CommandTimeout = 60;
        //                cmd.CommandText = @"
        //                                    INSERT INTO [TK].dbo.PURTH_AUDIT_LOG
        //                                    (TH001, TH002, TH003, TH004, OLD_TH010, OLD_TH117, OLD_TH036, 
        //                                     NEW_TH010, NEW_TH117, NEW_TH036, MODIFY_DATE, MODIFY_USER)
        //                                    VALUES
        //                                    (@TH001, @TH002, @TH003, @TH004, @OLD_TH010, @OLD_TH117, @OLD_TH036,
        //                                     @NEW_TH010, @NEW_TH117, @NEW_TH036, GETDATE(), SUSER_NAME())
        //                                    ";
        //                cmd.Parameters.AddWithValue("@TH001", TH001);
        //                cmd.Parameters.AddWithValue("@TH002", TH002);
        //                cmd.Parameters.AddWithValue("@TH003", TH003);
        //                cmd.Parameters.AddWithValue("@TH004", TH004);
        //                cmd.Parameters.AddWithValue("@OLD_TH010", OLD_TH010 ?? "");
        //                cmd.Parameters.AddWithValue("@OLD_TH117", OLD_TH117 ?? "");
        //                cmd.Parameters.AddWithValue("@OLD_TH036", OLD_TH036 ?? "");
        //                cmd.Parameters.AddWithValue("@NEW_TH010", NEW_TH010 ?? "");
        //                cmd.Parameters.AddWithValue("@NEW_TH117", NEW_TH117 ?? "");
        //                cmd.Parameters.AddWithValue("@NEW_TH036", NEW_TH036 ?? "");

        //                cmd.ExecuteNonQuery();
        //            }
        //        }
        //    }
        //    catch (Exception ex)
        //    {
        //        MessageBox.Show("保存審計日誌失敗: " + ex.Message);
        //    }
        //}

        public void SET_TEXT_NULL()
        {
            textBox2.Text = null;
            textBox3.Text = null;
            textBox4.Text = null;
            textBox5.Text = null;
            textBox6.Text = null;
            textBox7.Text = null;
            textBox8.Text = null;
            textBox9.Text = null;

            old_TH010 = null;
            old_TH117 = null;
            old_TH036 = null;

        }

        #endregion
        #region BUTTON
        private void button1_Click(object sender, EventArgs e)
        {
            string TH002= textBox1.Text.Trim();
            if(!string.IsNullOrEmpty(TH002))
            {
                SEARCH(TH002);
            }
            

        }

        private void button2_Click(object sender, EventArgs e)
        {
            // 顯示確認對話框
            DialogResult result = MessageBox.Show("確定要執行此操作嗎？", "確認", MessageBoxButtons.OKCancel);

            // 檢查使用者是否按下了確定按鈕
            if (result == DialogResult.OK)
            {
                string TH001 = textBox2.Text.Trim();
                string TH002 = textBox3.Text.Trim();
                string TH003 = textBox4.Text.Trim();
                string TH004 = textBox5.Text.Trim();
                string TH010 = textBox7.Text.Trim();
                string TH117 = textBox8.Text.Trim();
                string TH036 = textBox9.Text.Trim();
                UPDATE_PURTH_INVLA_INVME(TH001, TH002, TH003, TH004, TH010, TH117, TH036, old_TH010);

                SEARCH(TH002);
            }
        }

        #endregion


    }
}

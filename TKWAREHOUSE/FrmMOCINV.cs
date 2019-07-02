﻿using System;
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

namespace TKWAREHOUSE
{
    public partial class FrmMOCINV : Form
    {
        SqlConnection sqlConn = new SqlConnection();
        SqlCommand sqlComm = new SqlCommand();
        string connectionString;
        StringBuilder sbSql = new StringBuilder();
        StringBuilder sbSqlQuery = new StringBuilder();
        SqlDataAdapter adapter = new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilder = new SqlCommandBuilder();
        SqlDataAdapter adapter2 = new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilder2 = new SqlCommandBuilder();
        SqlDataAdapter adapter3 = new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilder3 = new SqlCommandBuilder();
        SqlDataAdapter adapter4 = new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilder4 = new SqlCommandBuilder();
        SqlTransaction tran;
        SqlCommand cmd = new SqlCommand();
        DataSet ds = new DataSet();
        DataSet ds2 = new DataSet();
        DataSet ds3 = new DataSet();
        DataSet ds4 = new DataSet();

        DataTable dt = new DataTable();
        string tablename = null;
        int result;

        DataTable ADDDT = new DataTable();

        string ID =null;


        public FrmMOCINV()
        {
            InitializeComponent();

            
            ADDDT.Columns.AddRange(new DataColumn[6] {
                 new DataColumn("庫別", typeof(string)),
                 new DataColumn("日期", typeof(string)),
                 new DataColumn("品號", typeof(string)),
                 new DataColumn("品名", typeof(string)),
                 new DataColumn("批號", typeof(string)),
                 new DataColumn("數量", typeof(decimal)) });
        }

        #region FUNCTION

        private void cbHeader_CheckedChanged(object sender, EventArgs e)
        {
            dataGridView1.EndEdit();

            foreach (DataGridViewRow dr in dataGridView1.Rows)
            {
                dr.Cells[0].Value = ((CheckBox)dataGridView1.Controls.Find("checkboxHeader", true)[0]).Checked;

            }

        }

        private void FrmMOCINV_Load(object sender, EventArgs e)
        {

            dataGridView1.AlternatingRowsDefaultCellStyle.BackColor = Color.PaleTurquoise;      //奇數列顏色

            //先建立個 CheckBox 欄
            DataGridViewCheckBoxColumn cbCol = new DataGridViewCheckBoxColumn();
            cbCol.Width = 80;   //設定寬度
            cbCol.HeaderText = "　全選";
            cbCol.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;   //置中
            cbCol.TrueValue = true;
            cbCol.FalseValue = false;
            dataGridView1.Columns.Insert(0, cbCol);

            #region 建立全选 CheckBox

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


            #endregion
        }

        public void SearchMOC()
        {
            try
            {
                if (!string.IsNullOrEmpty(dateTimePicker1.Value.ToString("yyyyMMdd")) || !string.IsNullOrEmpty(dateTimePicker2.Value.ToString("yyyyMMdd")))
                {
                    connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                    sqlConn = new SqlConnection(connectionString);

                    sbSql.Clear();
                    sbSqlQuery.Clear();
                    sbSql = SETsbSql();


                    adapter = new SqlDataAdapter(@"" + sbSql, sqlConn);
                    sqlCmdBuilder = new SqlCommandBuilder(adapter);

                    sqlConn.Open();
                    ds.Clear();
                    adapter.Fill(ds, tablename);
                    sqlConn.Close();


                    if (ds.Tables[tablename].Rows.Count == 0)
                    {
                        dataGridView1.DataSource = null;
                    }
                    else
                    {
                    
                        dataGridView1.DataSource = ds.Tables[tablename];
                        dataGridView1.AutoResizeColumns();
                    }
                }
                else
                {

                }



            }
            catch
            {

            }
            finally
            {

            }
        }

        public StringBuilder SETsbSql()
        {
            StringBuilder STR = new StringBuilder();

            STR.AppendFormat(@"  SELECT TB003 AS '品號',TB012 AS '品名',SUM(TB004) AS 數量");
            STR.AppendFormat(@"  FROM [TK].dbo.MOCTA,[TK].dbo.MOCTB");
            STR.AppendFormat(@"  WHERE TA001=TB001 AND TA002=TB002");
            STR.AppendFormat(@"  AND TA003>='{0}' AND TA003<='{1}' ", dateTimePicker1.Value.ToString("yyyyMMdd"), dateTimePicker2.Value.ToString("yyyyMMdd"));
            //STR.AppendFormat(@"  AND TB003 LIKE '201001165 %'  ");
            STR.AppendFormat(@"  AND( TB003 LIKE '1%' OR TB003 LIKE '2%')");
            STR.AppendFormat(@"  AND( TA021 LIKE '02%' OR TA021 LIKE '03%' OR TA021 LIKE '04%' OR TA021 LIKE '09%') ");
            STR.AppendFormat(@"  GROUP BY TB003,TB012");
            STR.AppendFormat(@"  ORDER BY TB003,TB012");
            
            STR.AppendFormat(@"  ");
            


            STR.AppendFormat(@"  ");
            tablename = "TEMPds1";
      
          

            return STR;
        }

        public void  SearchINV()
        {
            try
            {
                if (!string.IsNullOrEmpty(dateTimePicker1.Value.ToString("yyyyMMdd")) || !string.IsNullOrEmpty(dateTimePicker2.Value.ToString("yyyyMMdd")))
                {
                    connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                    sqlConn = new SqlConnection(connectionString);

                    sbSql.Clear();

                    sbSql.AppendFormat(@" SELECT LA009 AS '倉庫',LA001 AS '品號',LA016  AS '批號',SUM(LA005*LA011)  AS '庫存量' ");
                    sbSql.AppendFormat(@" FROM [TK].dbo.INVLA");
                    sbSql.AppendFormat(@" WHERE (LA009='20006' OR LA009='20004')");
                    sbSql.AppendFormat(@" AND LA001 IN (SELECT TB003");
                    sbSql.AppendFormat(@" FROM [TK].dbo.MOCTA,[TK].dbo.MOCTB");
                    sbSql.AppendFormat(@" WHERE TA001=TB001 AND TA002=TB002");
                    sbSql.AppendFormat(@"  AND TA003>='{0}' AND TA003<='{1}' ", dateTimePicker1.Value.ToString("yyyyMMdd"), dateTimePicker2.Value.ToString("yyyyMMdd"));
                    sbSql.AppendFormat(@"  AND( TA021 LIKE '02%' OR TA021 LIKE '03%' OR TA021 LIKE '04%' OR TA021 LIKE '09%') ");
                    sbSql.AppendFormat(@"  AND( TB003 LIKE '1%' OR TB003 LIKE '2%')");
                    //sbSql.AppendFormat(@" AND TB003='101002001'");
                    sbSql.AppendFormat(@" GROUP BY TB003");
                    sbSql.AppendFormat(@" )");
                    sbSql.AppendFormat(@" GROUP BY LA009,LA001,LA016");
                    sbSql.AppendFormat(@" HAVING  SUM(LA005*LA011)>0");
                    sbSql.AppendFormat(@" ORDER BY LA001,LA009,LA016");
                    sbSql.AppendFormat(@" ");



                    adapter = new SqlDataAdapter(@"" + sbSql, sqlConn);
                    sqlCmdBuilder = new SqlCommandBuilder(adapter);

                    sqlConn.Open();
                    ds2.Clear();
                    adapter.Fill(ds2, tablename);
                    sqlConn.Close();


                    if (ds2.Tables[tablename].Rows.Count == 0)
                    {
                        dataGridView2.DataSource = null;
                    }
                    else
                    {

                        dataGridView2.DataSource = ds2.Tables[tablename];
                        dataGridView2.AutoResizeColumns();
                    }
                }
                else
                {

                }



            }
            catch
            {

            }
            finally
            {

            }
        }


        public void SearchINVBATCH()
        {
            string ID = dateTimePicker1.Value.ToString("yyyyMMdd");
            ADDDT.Clear();

            if (dataGridView1.Rows.Count>=1)
            {
                foreach (DataGridViewRow row in dataGridView1.Rows)
                {
                    string MB001 = row.Cells["品號"].Value.ToString();
                    decimal NUM= Convert.ToDecimal(row.Cells["數量"].Value.ToString());

                    DataSet ds3 = new DataSet();
                    ds3.Clear();
                    ds3 = SearchINVNOW(MB001);

                    if(ds3!=null && ds3.Tables["ds3"].Rows.Count>=1)
                    {
                        int ROWS = ds3.Tables["ds3"].Rows.Count;
                        int NOWROWS = ds3.Tables["ds3"].Rows.Count;

                        while (NUM>0 && NOWROWS >= 1)
                        {
                            if(Convert.ToDecimal(ds3.Tables["ds3"].Rows[ROWS - NOWROWS]["庫存量"].ToString())>= NUM)
                            {
                                ADDDT.Rows.Add(ds3.Tables["ds3"].Rows[ROWS - NOWROWS]["庫別"].ToString(), ID,row.Cells["品號"].Value.ToString(), row.Cells["品名"].Value.ToString(), ds3.Tables["ds3"].Rows[ROWS - NOWROWS]["批號"].ToString(), NUM);
                            }
                            else if (Convert.ToDecimal(ds3.Tables["ds3"].Rows[ROWS - NOWROWS]["庫存量"].ToString()) < NUM)
                            {
                                ADDDT.Rows.Add(ds3.Tables["ds3"].Rows[ROWS - NOWROWS]["庫別"].ToString(), ID, row.Cells["品號"].Value.ToString(), row.Cells["品名"].Value.ToString(), ds3.Tables["ds3"].Rows[ROWS - NOWROWS]["批號"].ToString(), Convert.ToDecimal(ds3.Tables["ds3"].Rows[ROWS - NOWROWS]["庫存量"].ToString()));
                            }

                            NUM=NUM- Convert.ToDecimal(ds3.Tables["ds3"].Rows[ROWS- NOWROWS]["庫存量"].ToString());
                            NOWROWS = NOWROWS - 1;
                        }
                       
                    }
                    else
                    {

                    }

                }
            }

            if (ADDDT.Rows.Count >= 1)
            {
                dataGridView3.DataSource = ADDDT;
            }
            

        }

        public DataSet SearchINVNOW(string MB001)
        {
            try
            {
                if (!string.IsNullOrEmpty(dateTimePicker1.Value.ToString("MB001")) )
                {
                    connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                    sqlConn = new SqlConnection(connectionString);

                    sbSql.Clear();

                    sbSql.AppendFormat(@" SELECT LA009 AS '庫別',LA001 AS '品號',LA016  AS '批號',SUM(LA005*LA011)  AS '庫存量'  ");
                    sbSql.AppendFormat(@" FROM [TK].dbo.INVLA ");
                    sbSql.AppendFormat(@" WHERE (LA009='20006' OR LA009='20004')");
                    sbSql.AppendFormat(@" AND LA001 IN (SELECT TB003");
                    sbSql.AppendFormat(@" FROM [TK].dbo.MOCTA,[TK].dbo.MOCTB");
                    sbSql.AppendFormat(@" WHERE TA001=TB001 AND TA002=TB002");
                    sbSql.AppendFormat(@"  AND( TA021 LIKE '02%' OR TA021 LIKE '03%' OR TA021 LIKE '04%' OR TA021 LIKE '09%') ");
                    sbSql.AppendFormat(@"  AND( TB003 LIKE '1%' OR TB003 LIKE '2%')");
                    sbSql.AppendFormat(@"  AND TA003>='{0}' AND TA003<='{1}' ", dateTimePicker1.Value.ToString("yyyyMMdd"), dateTimePicker2.Value.ToString("yyyyMMdd"));
                    sbSql.AppendFormat(@" AND TB003='{0}'", MB001);
                    sbSql.AppendFormat(@" GROUP BY TB003");
                    sbSql.AppendFormat(@" )");
                    sbSql.AppendFormat(@" GROUP BY LA009,LA001,LA016");
                    sbSql.AppendFormat(@" HAVING  SUM(LA005*LA011)>0");
                    sbSql.AppendFormat(@" ORDER BY LA009,LA001,LA016");
                    sbSql.AppendFormat(@" ");



                    adapter2 = new SqlDataAdapter(@"" + sbSql, sqlConn);
                    sqlCmdBuilder2 = new SqlCommandBuilder(adapter2);

                    sqlConn.Open();
                    ds3.Clear();
                    adapter2.Fill(ds3, "ds3");
                    sqlConn.Close();


                    if (ds3.Tables["ds3"].Rows.Count == 0)
                    {
                        return null;
                    }
                    else
                    {
                        return ds3;

                    }
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

            }
        }

        public void ADDTOTKWAREHOUSE()
        {
            if (dataGridView3.Rows.Count >= 1)
            {
                try
                {
                    connectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
                    sqlConn = new SqlConnection(connectionString);

                    sqlConn.Close();
                    sqlConn.Open();
                    tran = sqlConn.BeginTransaction();

                    sbSql.Clear();

                    foreach (DataGridViewRow row in dataGridView3.Rows)
                    {
                        sbSql.AppendFormat(" INSERT INTO [TKWAREHOUSE].[dbo].[INVBATCH]");
                        sbSql.AppendFormat(" ([ID],[DATES],[WHID],[MB001],[MB002],[LOTNO],[NUM],[TA001],[TA002])");
                        sbSql.AppendFormat(" VALUES('{0}','{1}','{2}','{3}','{4}','{5}','{6}','{7}','{8}')",textBox1.Text,dateTimePicker3.Value.ToString("yyyyMMdd") ,row.Cells["庫別"].Value.ToString(), row.Cells["品號"].Value.ToString(), row.Cells["品名"].Value.ToString(), row.Cells["批號"].Value.ToString(), row.Cells["數量"].Value.ToString(),null,null);
                        sbSql.AppendFormat(" ");
                       
                    }

                    
  


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
        }

        public string GETMAXID()
        {
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                StringBuilder sbSql = new StringBuilder();
                sbSql.Clear();
                sbSqlQuery.Clear();
                ds4.Clear();

                
                sbSql.AppendFormat(@" SELECT ISNULL(MAX(ID),'00000000000') AS ID ");
                sbSql.AppendFormat(@" FROM  [TKWAREHOUSE].[dbo].[INVBATCH] ");
                sbSql.AppendFormat(@" WHERE CONVERT(NVARCHAR,DATES,112)='{0}' ",dateTimePicker3.Value.ToString("yyyyMMdd"));
                sbSql.AppendFormat(@"  ");

                adapter4 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder4 = new SqlCommandBuilder(adapter4);
                sqlConn.Open();
                ds4.Clear();
                adapter4.Fill(ds4, "TEMPds4");
                sqlConn.Close();


                if (ds4.Tables["TEMPds4"].Rows.Count == 0)
                {
                    return null;
                }
                else
                {
                    if (ds4.Tables["TEMPds4"].Rows.Count >= 1)
                    {
                        ID = SETTID(ds4.Tables["TEMPds4"].Rows[0]["ID"].ToString());
                        return ID;

                    }
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

        public string SETTID(string ID)
        {
            if (ID.Equals("00000000000"))
            {
                return dateTimePicker3.Value.ToString("yyyyMMdd") + "001";
            }

            else
            {
                int serno = Convert.ToInt16(ID.Substring(8, 3));
                serno = serno + 1;
                string temp = serno.ToString();
                temp = temp.PadLeft(3, '0');
                return dateTimePicker3.Value.ToString("yyyyMMdd") + temp.ToString();
            }
        }

        #endregion

        #region BUTTON

        private void button1_Click(object sender, EventArgs e)
        {
            SearchMOC();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            SearchINV();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            SearchINVBATCH();
        }

        private void button4_Click(object sender, EventArgs e)
        {
            textBox1.Text = GETMAXID();
            ADDTOTKWAREHOUSE();
        }
        #endregion

    }
}

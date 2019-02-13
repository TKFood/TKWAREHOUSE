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
using System.Text.RegularExpressions;
using System.Globalization;

namespace TKWAREHOUSE
{
    public partial class FrmPURTAB : Form
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

        SqlTransaction tran;
        SqlCommand cmd = new SqlCommand();
        DataSet ds = new DataSet();
        DataSet ds2 = new DataSet();

        int result;
        string tablename = null;

        Thread TD;
        

        public FrmPURTAB()
        {
            InitializeComponent();

            SETGRIDVIEW();
        }

        #region FUNCTION
        public void SETGRIDVIEW()
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
        public void SEARCHMOCTA()
        {
            StringBuilder SLQURY = new StringBuilder();
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sbSql.Clear();

                SLQURY.Clear();

                if(checkBox1.Checked==true)
                {
                    SLQURY.AppendFormat(@"  AND TA001+TA002 NOT IN (SELECT [MOCTA001]+[MOCTA002] FROM [TKWAREHOUSE].dbo.PURTAB)");
                }


                sbSql.AppendFormat(@"  SELECT TA001 AS '單別',TA002 AS '單號',TA003 AS '生產日',TA034 AS '品名',TA006 AS '品號',TA015 AS '生產量',TA007 AS '單位'");
                sbSql.AppendFormat(@"  FROM [TK].dbo.[MOCTA]");
                sbSql.AppendFormat(@"  WHERE TA003>='{0}' AND TA003<='{1}'",dateTimePicker1.Value.ToString("yyyyMMdd"), dateTimePicker2.Value.ToString("yyyyMMdd"));
                sbSql.AppendFormat(@"  {0}", SLQURY.ToString());
                sbSql.AppendFormat(@"  ORDER BY TA003,TA034");
                sbSql.AppendFormat(@"  ");


                adapter = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder = new SqlCommandBuilder(adapter);
                sqlConn.Open();
                ds.Clear();
                adapter.Fill(ds, "TEMPds");
                sqlConn.Close();


                if (ds.Tables["TEMPds"].Rows.Count == 0)
                {
                    dataGridView1.DataSource = null;
                }
                else
                {
                    if (ds.Tables["TEMPds"].Rows.Count >= 1)
                    {                       
                        dataGridView1.DataSource = ds.Tables["TEMPds"];

                        dataGridView1.AutoResizeColumns();
                        dataGridView1.FirstDisplayedScrollingRowIndex = dataGridView2.RowCount - 1;


                    }

                }

            }
            catch
            {

            }
            finally
            {

            }
        }

        public void ADDPURTAB()
        {
            foreach (DataGridViewRow dr in this.dataGridView1.Rows)
            {
                if (dr.Cells[0].Value != null && (bool)dr.Cells[0].Value)
                {
                    try
                    {
                        connectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
                        sqlConn = new SqlConnection(connectionString);

                        sqlConn.Close();
                        sqlConn.Open();
                        tran = sqlConn.BeginTransaction();

                        sbSql.Clear();
                        sbSql.AppendFormat(@" INSERT INTO [TKWAREHOUSE].[dbo].[PURTAB]");
                        sbSql.AppendFormat(@" ([ID],[IDDATES],[MOCTA001],[MOCTA002],[MOCTA003],[MOCTA006],[MOCTA007],[MOCTA015],[MOCTA034],[PURTA001],[PURTA002])");
                        sbSql.AppendFormat(@" VALUES ({0},'{1}','{2}','{3}','{4}','{5}','{6}','{7}','{8}','{9}','{10}')",textBox1.Text,dateTimePicker3.Value.ToString("yyyyMMdd"), dr.Cells["單別"].Value.ToString(), dr.Cells["單號"].Value.ToString(), dr.Cells["生產日"].Value.ToString(), dr.Cells["品號"].Value.ToString(), dr.Cells["單位"].Value.ToString(), dr.Cells["生產量"].Value.ToString(), dr.Cells["品名"].Value.ToString(),"","");
                        sbSql.AppendFormat(@" ");
                        
                   
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
        }

        public void SEARCHPURTAB()
        {
            StringBuilder SLQURY = new StringBuilder();
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sbSql.Clear();

                sbSql.AppendFormat(@"  SELECT [ID] AS '批號',[IDDATES] AS '請購日'");
                sbSql.AppendFormat(@"  FROM [TKWAREHOUSE].[dbo].[PURTAB]");
                sbSql.AppendFormat(@"  WHERE [IDDATES]='20190213'");
                sbSql.AppendFormat(@"  GROUP BY  [ID],[IDDATES] ");
                sbSql.AppendFormat(@"  ");

                adapter2 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder2 = new SqlCommandBuilder(adapter2);
                sqlConn.Open();
                ds2.Clear();
                adapter2.Fill(ds2, "TEMPds2");
                sqlConn.Close();


                if (ds2.Tables["TEMPds2"].Rows.Count == 0)
                {
                    dataGridView2.DataSource = null;
                }
                else
                {
                    if (ds2.Tables["TEMPds2"].Rows.Count >= 1)
                    {
                        dataGridView2.DataSource = ds2.Tables["TEMPds2"];

                        dataGridView2.AutoResizeColumns();
                        dataGridView2.FirstDisplayedScrollingRowIndex = dataGridView2.RowCount - 1;


                    }

                }

            }
            catch
            {

            }
            finally
            {

            }
        }
        #endregion

        #region BUTTON
        private void button1_Click(object sender, EventArgs e)
        {
            SEARCHMOCTA();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            ADDPURTAB();
            SEARCHMOCTA();
            SEARCHPURTAB();
        }

        private void button3_Click(object sender, EventArgs e)
        {

        }
        #endregion


    }
}

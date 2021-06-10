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

namespace TKWAREHOUSE
{
    public partial class FrmPURTABCD : Form
    {
        SqlConnection sqlConn = new SqlConnection();
        SqlCommand sqlComm = new SqlCommand();
        string connectionString;
        SqlTransaction tran;
        SqlCommand cmd = new SqlCommand();
        StringBuilder sbSql = new StringBuilder();
        StringBuilder sbSqlQuery = new StringBuilder();
        SqlDataAdapter adapter = new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilder = new SqlCommandBuilder();
        int result;
        DataSet ds = new DataSet();
        DataSet ds2 = new DataSet();

        DataTable dt = new DataTable();


        public FrmPURTABCD()
        {
            InitializeComponent();


        }


        #region FUNCTION

        private void FrmPURTABCD_Load(object sender, EventArgs e)
        {
            dataGridView1.AlternatingRowsDefaultCellStyle.BackColor = Color.PaleTurquoise;      //奇數列顏色
            dataGridView2.AlternatingRowsDefaultCellStyle.BackColor = Color.PaleTurquoise;


            //先建立個 CheckBox 欄
            DataGridViewCheckBoxColumn cbCol = new DataGridViewCheckBoxColumn();
            cbCol.Width = 120;   //設定寬度
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

        private void cbHeader_CheckedChanged(object sender, EventArgs e)
        {
            dataGridView1.EndEdit();

            foreach (DataGridViewRow dr in dataGridView1.Rows)
            {
                dr.Cells[0].Value = ((CheckBox)dataGridView1.Controls.Find("checkboxHeader", true)[0]).Checked;

            }

        }


        public void SEARCH(string STATUS)
        {
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);
                SqlDataAdapter adapter = new SqlDataAdapter();
                SqlCommandBuilder sqlCmdBuilder = new SqlCommandBuilder();
                DataSet ds = new DataSet();

                sbSql.Clear();
                sbSqlQuery.Clear();

                if(STATUS.Equals("未確認"))
                {
                    sbSqlQuery.AppendFormat(@" AND TA001+TA002+TB003 NOT IN (SELECT TA001+TA002+TB003 FROM [TKWAREHOUSE].[dbo].[CHECKPURTABPURTCD])");

                }
                else if (STATUS.Equals("已確認"))
                {
                    sbSqlQuery.AppendFormat(@" ");
                }

                sbSql.AppendFormat(@"  
                                    SELECT TA001 AS '請購單別',TA002 AS '請購單號',TB003 AS '請購序號',TB004 AS '品號',TB005 AS '品名',TB006 AS '規格',TB009 AS '請購數量',TB007 AS '請購單位',TB022 AS '採購單',TD008 AS '採購數量',TD009 AS '採購單位'
                                    FROM [TK].dbo.PURTA,[TK].dbo.PURTB
                                    LEFT JOIN [TK].dbo.[PURTD] ON TD001=SUBSTRING(TB022,1,4) AND TD002=SUBSTRING(TB022,6,11) AND TD003=SUBSTRING(TB022,18,4)
                                    WHERE TA001=TB001 AND TA002=TB002
                                    AND TA003>='{0}' AND TA003<='{1}'
                                    {2}
                                    ", dateTimePicker1.Value.ToString("yyyyMMdd"), dateTimePicker2.Value.ToString("yyyyMMdd"), sbSqlQuery.ToString());

                adapter = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder = new SqlCommandBuilder(adapter);
                sqlConn.Open();
                ds.Clear();
                adapter.Fill(ds, "ds");
                sqlConn.Close();


                if (ds.Tables["ds"].Rows.Count == 0)
                {
                    dataGridView1.DataSource = null;
                }
                else
                {
                    if (ds.Tables["ds"].Rows.Count >= 1)
                    {
                        //dataGridView1.Rows.Clear();
                        dataGridView1.DataSource = ds.Tables["ds"];
                        dataGridView1.AutoResizeColumns();
                        //dataGridView1.CurrentCell = dataGridView1[0, rownum];

                        dataGridView1.AutoResizeColumns();
                        dataGridView1.ColumnHeadersDefaultCellStyle.Font = new Font("Tahoma", 9);
                        dataGridView1.DefaultCellStyle.Font = new Font("Tahoma", 10);
                        dataGridView1.Columns["請購單別"].Width = 60;
                        dataGridView1.Columns["請購單號"].Width = 100;
                        dataGridView1.Columns["請購序號"].Width = 60;
                        dataGridView1.Columns["品號"].Width = 100;
                        dataGridView1.Columns["品名"].Width = 150;
                        dataGridView1.Columns["規格"].Width = 80;
                        dataGridView1.Columns["請購數量"].Width = 60;
                        dataGridView1.Columns["請購單位"].Width = 60;
                        dataGridView1.Columns["採購單"].Width = 200;
                        dataGridView1.Columns["採購數量"].Width = 60;
                        dataGridView1.Columns["採購單位"].Width = 60;
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

        public void ADDCHECKPURTABPURTCD()
        {
            StringBuilder ADDSQL = new StringBuilder();

            foreach (DataGridViewRow dgR in this.dataGridView1.Rows)
            {
                try
                {
                    DataGridViewCheckBoxCell cbx = (DataGridViewCheckBoxCell)dgR.Cells[0];
                    if ((bool)cbx.FormattedValue)
                    {
                        ADDSQL.AppendFormat(@" 
                                            INSERT INTO [TKWAREHOUSE].[dbo].[CHECKPURTABPURTCD]
                                            ( [TA001],[TA002],[TB003],[TB004],[TB009],[TB007],[TB022],[TD008],[TD009])
                                            VALUES ('{0}','{1}','{2}','{3}','{4}','{5}','{6}','{7}','{8}')
                                            ", dgR.Cells["請購單別"].Value.ToString(), dgR.Cells["請購單號"].Value.ToString(), dgR.Cells["請購序號"].Value.ToString(), dgR.Cells["品號"].Value.ToString(), dgR.Cells["請購數量"].Value.ToString(), dgR.Cells["請購單位"].Value.ToString(), dgR.Cells["採購單"].Value.ToString(), dgR.Cells["採購數量"].Value.ToString(), dgR.Cells["採購單位"].Value.ToString());
                    }
                }
                catch (Exception ex)
                {
                    //MessageBox.Show(ex.Message);
                }
            }

            if(!string.IsNullOrEmpty(ADDSQL.ToString()))
            {
                try
                {
                    connectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
                    sqlConn = new SqlConnection(connectionString);

                    sqlConn.Close();
                    sqlConn.Open();
                    tran = sqlConn.BeginTransaction();

                    //sbSql.Clear();
                    
                    //sbSql.AppendFormat(" ");


                    cmd.Connection = sqlConn;
                    cmd.CommandTimeout = 60;
                    cmd.CommandText = ADDSQL.ToString();
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

        #endregion

        #region BUTTON
        private void button3_Click(object sender, EventArgs e)
        {
            SEARCH(comboBox1.Text.ToString());
        }
        private void button1_Click(object sender, EventArgs e)
        {
            ADDCHECKPURTABPURTCD();

            SEARCH(comboBox1.Text.ToString());
        }
        #endregion


    }
}

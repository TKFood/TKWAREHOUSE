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

namespace TKWAREHOUSE
{
    public partial class FrmhctEDI : Form
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

        DataTable dt = new DataTable();

        string Id;
        string Name;
        string Phone;
        string Tel;
        string Addr;
        string Comment;
        string GetMoney;
        string ReceiveDay;
        string ReceiveTime;
        string Goods;
        string weight;
        string tablename = null;

        public FrmhctEDI()
        {
            InitializeComponent();
        }

        #region FUNCTION
        private void FrmhctEDI_Load(object sender, EventArgs e)
        {
            dataGridView1.AlternatingRowsDefaultCellStyle.BackColor = Color.PaleTurquoise;      //奇數列顏色

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
        public void Search()
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
                        label14.Text = "找不到資料";
                    }
                    else
                    {
                        label14.Text = "有 " + ds.Tables[tablename].Rows.Count.ToString() + " 筆";

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

            if (comboBox1.Text.ToString().Equals("銷貨單"))
            {

                STR.Append(@" SELECT TG001 AS '銷貨單別',TG002  AS '銷貨單號',TG003  AS '銷貨日',''  AS '收貨人代號',TG007  AS '客戶',TG008  AS '地址',TG066  AS '收貨人' ,TG106  AS '電話',TG113  AS '代收貸款',TG110  AS '指定日期'  ");
                STR.Append(@" ,CASE WHEN TG111='1' THEN '' WHEN TG111='2' THEN '09-13' WHEN TG111='3' THEN '13-17'  WHEN TG111='4' THEN '17-20'   WHEN TG111='5' THEN '09'  WHEN TG111='6' THEN '10'  WHEN TG111='7' THEN '11'   WHEN TG111='8' THEN '12'   WHEN TG111='9' THEN '13' WHEN TG111='A' THEN '14' WHEN TG111='B' THEN '15'  WHEN TG111='C' THEN '16'  WHEN TG111='D' THEN '17'  WHEN TG111='E' THEN '18'  WHEN TG111='F' THEN '19'  WHEN TG111='G' THEN '20' END AS '指定時間' ");
                STR.Append(@" ,TG020 AS '備註' ");
                STR.AppendFormat(@"  FROM [{0}].dbo.COPTG WITH (NOLOCK)  ", sqlConn.Database.ToString());
                STR.AppendFormat(@" WHERE TG003>='{0}' AND TG003<='{1}' ", dateTimePicker1.Value.ToString("yyyyMMdd"), dateTimePicker2.Value.ToString("yyyyMMdd"));


                STR.AppendFormat(@"  ");
                tablename = "TEMPds1";
            }
            else if (comboBox1.Text.ToString().Equals("門市宅配單"))
            {
                STR.AppendFormat(@"  SELECT  TA001 AS '單別',TA002  AS '單號',TA014  AS '出貨日' ");
                STR.AppendFormat(@"  ,''  AS '收貨人代號',TA024  AS '客戶',TA027  AS '地址',TA024  AS '收貨人'  ");
                STR.AppendFormat(@"   ,TA025  AS '電話',0  AS '代收貸款',TA014  AS '指定日期'");
                STR.AppendFormat(@"   ,'' AS '指定時間' ");
                STR.AppendFormat(@"   ,TA030 AS '備註' ");
                STR.AppendFormat(@"   FROM [TK].dbo.INVTA WITH (NOLOCK) ");
                STR.AppendFormat(@"   WHERE TA001='A124'");
                STR.AppendFormat(@"  AND TA014>='{0}' AND TA014<='{1}' ", dateTimePicker1.Value.ToString("yyyyMMdd"), dateTimePicker2.Value.ToString("yyyyMMdd"));
                STR.AppendFormat(@"  ");

                STR.AppendFormat(@"  ");
                tablename = "TEMPds2";

            }
            else if (comboBox1.Text.ToString().Equals("借出單"))
            {
                STR.AppendFormat(@"   SELECT  TF001 AS '單別',TF002  AS '單號',TF003  AS '出貨日'");
                STR.AppendFormat(@"   ,''  AS '收貨人代號',TF006  AS '客戶',TF016  AS '地址',TF006  AS '收貨人' ");
                STR.AppendFormat(@"   ,''  AS '電話',0  AS '代收貸款',TF003  AS '指定日期' ");
                STR.AppendFormat(@"   ,'' AS '指定時間'");
                STR.AppendFormat(@"   ,TF014 AS '備註'   ");
                STR.AppendFormat(@"   FROM [TK].dbo.INVTF WITH (NOLOCK) ");
                STR.AppendFormat(@"   WHERE TF001='A131' ");
                STR.AppendFormat(@"  AND TF003 >='{0}' AND TF003<='{1}' ", dateTimePicker1.Value.ToString("yyyyMMdd"), dateTimePicker2.Value.ToString("yyyyMMdd"));
                STR.AppendFormat(@"  ");
                STR.AppendFormat(@"  ");
                STR.AppendFormat(@"  ");

                STR.AppendFormat(@"  ");
                tablename = "TEMPds3";
            }
            else if (comboBox1.Text.ToString().Equals("全聯寄售單"))
            {
                STR.AppendFormat(@"  SELECT  TA001 AS '單別',TA002  AS '單號',TA014  AS '出貨日' ");
                STR.AppendFormat(@"  ,''  AS '收貨人代號',TA024  AS '客戶',TA027  AS '地址',TA024  AS '收貨人'  ");
                STR.AppendFormat(@"   ,TA025  AS '電話',0  AS '代收貸款',TA014  AS '指定日期'");
                STR.AppendFormat(@"   ,'' AS '指定時間' ");
                STR.AppendFormat(@"   ,TA030 AS '備註' ");
                STR.AppendFormat(@"   FROM [TK].dbo.INVTA WITH (NOLOCK) ");
                STR.AppendFormat(@"   WHERE TA001='A128'");
                STR.AppendFormat(@"  AND TA014>='{0}' AND TA014<='{1}' ", dateTimePicker1.Value.ToString("yyyyMMdd"), dateTimePicker2.Value.ToString("yyyyMMdd"));
                STR.AppendFormat(@"  ");

                STR.AppendFormat(@"  ");
                tablename = "TEMPds4";
            }

            return STR;
        }

        private void cbHeader_CheckedChanged(object sender, EventArgs e)
        {
            dataGridView1.EndEdit();

            foreach (DataGridViewRow dr in dataGridView1.Rows)
            {
                dr.Cells[0].Value = ((CheckBox)dataGridView1.Controls.Find("checkboxHeader", true)[0]).Checked;

            }           

        }
        #endregion

        #region BUTTON
        private void button1_Click(object sender, EventArgs e)
        {
            Search();
        }


        #endregion

       
    }
}

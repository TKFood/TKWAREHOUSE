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
    public partial class FrmhctEDIExport : Form
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

        public FrmhctEDIExport()
        {
            InitializeComponent();
        }

        #region FUNCTION
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

                    sbSqlQuery.AppendFormat("   TG003>='{0}' AND TG003<='{1}' ", dateTimePicker1.Value.ToString("yyyyMMdd"), dateTimePicker2.Value.ToString("yyyyMMdd"));
                    sbSql.Append(@" SELECT TG001 AS '銷貨單別',TG002  AS '銷貨單號',TG003  AS '銷貨日',''  AS '收貨人代號',TG007  AS '客戶',TG008  AS '地址',TG066  AS '收貨人' ,TG106  AS '電話',TG113  AS '代收貸款',TG110  AS '指定日期'  ");
                    sbSql.Append(@" ,CASE WHEN TG111='1' THEN '' WHEN TG111='2' THEN '09-13' WHEN TG111='3' THEN '13-17'  WHEN TG111='4' THEN '17-20'   WHEN TG111='5' THEN '09'  WHEN TG111='6' THEN '10'  WHEN TG111='7' THEN '11'   WHEN TG111='8' THEN '12'   WHEN TG111='9' THEN '13' WHEN TG111='A' THEN '14' WHEN TG111='B' THEN '15'  WHEN TG111='C' THEN '16'  WHEN TG111='D' THEN '17'  WHEN TG111='E' THEN '18'  WHEN TG111='F' THEN '19'  WHEN TG111='G' THEN '20' END AS '指定時間' ");
                    sbSql.Append(@" ,TG020 AS '備註' ");
                    sbSql.AppendFormat(@"  FROM [{0}].dbo.COPTG WITH (NOLOCK)  ", sqlConn.Database.ToString());
                    sbSql.AppendFormat(@" WHERE {0} ", sbSqlQuery.ToString());

                    adapter = new SqlDataAdapter(@"" + sbSql, sqlConn);
                    sqlCmdBuilder = new SqlCommandBuilder(adapter);

                    sqlConn.Open();
                    ds.Clear();
                    adapter.Fill(ds, "TEMPds");
                    sqlConn.Close();


                    if (ds.Tables["TEMPds"].Rows.Count == 0)
                    {
                        label14.Text = "找不到資料";
                    }
                    else
                    {
                        label14.Text = "銷貨單 有 " + ds.Tables["TEMPds"].Rows.Count.ToString() + " 筆";

                        dataGridView1.DataSource = ds.Tables["TEMPds"];
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

        public void ExcelExport()
        {
            
            string NowDB = "TK";
            //建立Excel 2003檔案
            IWorkbook wb = new XSSFWorkbook();
            ISheet ws;


            dt = ds.Tables["TEMPds"];
            if (dt.TableName != string.Empty)
            {
                ws = wb.CreateSheet(dt.TableName);
            }
            else
            {
                ws = wb.CreateSheet("Sheet1");
            }

            ws.CreateRow(0);
            //第一行為欄位名稱
            ws.GetRow(0).CreateCell(0).SetCellValue("收貨人代號");
            ws.GetRow(0).CreateCell(1).SetCellValue("收貨人名稱");
            ws.GetRow(0).CreateCell(2).SetCellValue("電話1");
            ws.GetRow(0).CreateCell(3).SetCellValue("電話2");
            ws.GetRow(0).CreateCell(4).SetCellValue("地址");
            ws.GetRow(0).CreateCell(5).SetCellValue("備註");
            ws.GetRow(0).CreateCell(6).SetCellValue("代收貸款");
            ws.GetRow(0).CreateCell(7).SetCellValue("指定日期");
            ws.GetRow(0).CreateCell(8).SetCellValue("指定時間");
            ws.GetRow(0).CreateCell(9).SetCellValue("件數");
            ws.GetRow(0).CreateCell(10).SetCellValue("重量");

            //for (int i = 0; i < dt.Columns.Count; i++)
            //{
            //    ws.GetRow(0).CreateCell(i).SetCellValue(dt.Columns[i].ColumnName);
            //}
            //for (int i = 0; i < dt.Rows.Count; i++)
            //{
            //    ws.CreateRow(i + 1);
            //    for (int j = 0; j < dt.Columns.Count; j++)
            //    {
            //        ws.GetRow(i + 1).CreateCell(j).SetCellValue(dt.Rows[i][j].ToString());
            //    }
            //}


            int j = 0;
            foreach (DataGridViewRow dr in this.dataGridView1.Rows)
            {
                ws.CreateRow(j + 1);
                if (dr.Cells[0].Value != null && (bool)dr.Cells[0].Value)
                {


                    Id = ((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[3].ToString();
                    //處理收貨人名稱、地址、電話
                    String value = ((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[5].ToString();
                    Char delimiter = ',';
                    String[] substrings = value.Split(delimiter);      
                    if(substrings.Length>=3)
                    {
                        Addr = substrings[0];
                        Name = substrings[1];
                        Phone = substrings[2];
                    }  
                    else if (substrings.Length == 2)
                    {
                        Addr = substrings[0];
                        Name = substrings[1];
                    }
                    else if (substrings.Length == 1)
                    {
                        Addr = substrings[0];
                    }

                    //Regex rgx = new Regex("\\d*");
                    //Addr = rgx.Replace(Addr, String.Empty);

                    Tel = null;
                    
                    try
                    {
                        if (((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[11].ToString().Substring(0, 1) == "Y")
                        {
                            Comment = ((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[11].ToString();
                            string[] sArray = Comment.Split('N');
                            Comment = sArray[0].ToString();
                        }
                        else
                        {
                            Comment = null;
                        }
                    }
                    catch
                    {
                        Comment = null;
                    }

                                 

                    GetMoney =((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[8].ToString();
                    ReceiveDay = ((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[9].ToString();
                    ReceiveTime = ((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[10].ToString();
                    Goods = null;
                    weight = null;

                    ws.GetRow(j + 1).CreateCell(0).SetCellValue(Id);
                    ws.GetRow(j + 1).CreateCell(1).SetCellValue(Name);
                    ws.GetRow(j + 1).CreateCell(2).SetCellValue(Phone);
                    ws.GetRow(j + 1).CreateCell(3).SetCellValue(Tel);
                    ws.GetRow(j + 1).CreateCell(4).SetCellValue(Addr);
                    ws.GetRow(j + 1).CreateCell(5).SetCellValue(Comment);
                    ws.GetRow(j + 1).CreateCell(6).SetCellValue(GetMoney);
                    ws.GetRow(j + 1).CreateCell(7).SetCellValue(ReceiveDay);
                    ws.GetRow(j + 1).CreateCell(8).SetCellValue(ReceiveTime);
                    ws.GetRow(j + 1).CreateCell(9).SetCellValue(Goods);
                    ws.GetRow(j + 1).CreateCell(10).SetCellValue(weight);


                    j++;
                    //MessageBox.Show("號碼 " + ((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[0] + ((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[1] + " 被選取了！");
                }


            }

            if (Directory.Exists(@"c:\temp\"))
            {
                //資料夾存在
            }
            else
            {
                //新增資料夾
                Directory.CreateDirectory(@"c:\temp\");
            }
            StringBuilder filename = new StringBuilder();
            filename.AppendFormat(@"c:\temp\宅配資料{0}.xlsx", DateTime.Now.ToString("yyyyMMdd"));

            FileStream file = new FileStream(filename.ToString(), FileMode.Create);//產生檔案
            wb.Write(file);
            file.Close();

            MessageBox.Show("匯出完成-EXCEL放在-" + filename.ToString());
            FileInfo fi = new FileInfo(filename.ToString());
            if (fi.Exists)
            {
                System.Diagnostics.Process.Start(filename.ToString());
            }
            else
            {
                //file doesn't exist
            }

            
        }

        private void cbHeader_CheckedChanged(object sender, EventArgs e)
        {
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
        private void button2_Click(object sender, EventArgs e)
        {
            ExcelExport();
        }
        #endregion

        #region ENEVT
        private void FrmhctEDIExport_Load(object sender, EventArgs e)
        {
            dataGridView1.AlternatingRowsDefaultCellStyle.BackColor = Color.PaleTurquoise;      //奇數列顏色

            //先建立個 CheckBox 欄
            DataGridViewCheckBoxColumn cbCol = new DataGridViewCheckBoxColumn();
            cbCol.Width = 50;   //設定寬度
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

        #endregion


    }
}

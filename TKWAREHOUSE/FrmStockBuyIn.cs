using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Data.SqlClient;
using System.Configuration;
using System.Reflection;

namespace TKWAREHOUSE
{
     #region DEFINE
   

    #endregion
    public partial class FrmStockBuyIn : Form
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
        DataSet dsPURTD = new DataSet();
        DataSet dsPURTH = new DataSet();
        DataTable dt = new DataTable();
        int result;
        string NowDay;

        internal class groupDtInfo
        {
            public groupDtInfo()
            {
            }

            public decimal Sum { get; set; }
            public object TH004 { get; set; }
        }

        internal class PURTDInfo
        {
            public PURTDInfo()
            {
            }

            public decimal Sum { get; set; }
            public object TD004 { get; set; }
        }


        public FrmStockBuyIn()
        {
            InitializeComponent();
            comboboxload();
            combobox2load();
            combobox3load();

            //定義
            dt.Columns.Add(new DataColumn("TH004", typeof(string)));
            dt.Columns.Add(new DataColumn("TH005", typeof(string)));
            dt.Columns.Add(new DataColumn("TH007", typeof(string)));
            dt.Columns.Add(new DataColumn("TH008", typeof(string)));
            dt.Columns.Add(new DataColumn("TH009", typeof(string)));
            dt.Columns.Add(new DataColumn("TH009CH", typeof(string)));
            dt.Columns.Add(new DataColumn("TH010", typeof(string)));
        }
        #region FUNCTION
        public void comboboxload()
        {
           
            connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
            sqlConn = new SqlConnection(connectionString);
            String Sequel = "SELECT MC001,MC002 FROM CMSMC WITH (NOLOCK) ORDER BY MC001";
            SqlDataAdapter da = new SqlDataAdapter(Sequel, sqlConn);
            DataTable dt = new DataTable();
            sqlConn.Open();

            dt.Columns.Add("MC001", typeof(string));
            dt.Columns.Add("MC002", typeof(string));
            da.Fill(dt);
            comboBox1.DataSource = dt.DefaultView;
            comboBox1.ValueMember = "MC001";
            comboBox1.DisplayMember = "MC002";
            sqlConn.Close();


        }

        public void combobox2load()
        {

            connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
            sqlConn = new SqlConnection(connectionString);
            String Sequel = "SELECT MQ001 FROM CMSMQ WITH (NOLOCK) WHERE MQ003='33'  ORDER BY MQ001";
            SqlDataAdapter da = new SqlDataAdapter(Sequel, sqlConn);
            DataTable dt = new DataTable();
            sqlConn.Open();

            dt.Columns.Add("MQ001", typeof(string));
            da.Fill(dt);
            comboBox2.DataSource = dt.DefaultView;
            comboBox2.ValueMember = "MQ001";
            comboBox2.DisplayMember = "MQ001";
            sqlConn.Close();


        }

        public void combobox3load()
        {

            connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
            sqlConn = new SqlConnection(connectionString);
            String Sequel = "SELECT MC001,MC002 FROM CMSMC WITH (NOLOCK) ORDER BY MC001";
            SqlDataAdapter da = new SqlDataAdapter(Sequel, sqlConn);
            DataTable dt = new DataTable();
            sqlConn.Open();

            dt.Columns.Add("MC001", typeof(string));
            dt.Columns.Add("MC002", typeof(string));
            da.Fill(dt);
            comboBox3.DataSource = dt.DefaultView;
            comboBox3.ValueMember = "MC001";
            comboBox3.DisplayMember = "MC002";
            sqlConn.Close();


        }
        public void Search()
        {
            try
            {

                if (!string.IsNullOrEmpty(comboBox2.Text.ToString()) || !string.IsNullOrEmpty(textBox1.Text.ToString()))
                {
                    connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                    sqlConn = new SqlConnection(connectionString);

                    sbSql.Clear();
                    sbSqlQuery.Clear();

                    sbSqlQuery.AppendFormat("  WHERE TD001='{0}' AND TD002='{1}' ", comboBox2.Text.ToString(), textBox1.Text.ToString());
                    sbSql.AppendFormat(@"SELECT TD003,TD004,TD005,TD006,TD007,TD008,TD009 FROM PURTD WITH (NOLOCK) {0} ", sbSqlQuery.ToString());

                    adapter = new SqlDataAdapter(@"" + sbSql, sqlConn);
                    sqlCmdBuilder = new SqlCommandBuilder(adapter);

                    sqlConn.Open();
                    dsPURTD.Clear();
                    adapter.Fill(dsPURTD, "TEMPdsPURTD");
                    sqlConn.Close();
                    

                    if (dsPURTD.Tables["TEMPdsPURTD"].Rows.Count == 0)
                    {
                        label14.Text = "找不到資料";
                    }
                    else
                    {
                        label14.Text = "採購品 有 "+dsPURTD.Tables["TEMPdsPURTD"].Rows.Count.ToString()+" 項";

                        dataGridView1.DataSource = dsPURTD.Tables["TEMPdsPURTD"];
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

        public void TempAdd()
        {
            
            DataRow dr;
            
            //dt.Columns.Add(new DataColumn("TH010", typeof(string)));
            // 加入第一筆銀行資料
            dr = dt.NewRow();
            dr["TH004"] = textBox4.Text.ToString();
            dr["TH005"] = dataGridView1.CurrentRow.Cells[2].Value.ToString();
            dr["TH007"] = textBox3.Text.ToString();
            dr["TH008"] = dataGridView1.CurrentRow.Cells[6].Value.ToString();
            dr["TH009"] = comboBox1.SelectedValue.ToString();
            dr["TH009CH"] = comboBox1.Text.ToString();
            dr["TH010"] = textBox2.Text.ToString();
            dt.Rows.Add(dr);




            //新增資料至DataTable的dt內
           
            dataGridView2.DataSource = dt;
            dataGridView2.AutoResizeColumns();
        }

        public void TempUpdate()
        {           
            dt.Rows[dataGridView2.CurrentCell.RowIndex]["TH007"] = textBox5.Text.ToString();
            dt.Rows[dataGridView2.CurrentCell.RowIndex]["TH009"] = comboBox3.Text.ToString();
            dt.Rows[dataGridView2.CurrentCell.RowIndex]["TH010"] = textBox6.Text.ToString();

        }

        public void TempDelete()
        {
            dt.Rows.RemoveAt(dataGridView2.CurrentCell.RowIndex);
        }
        public void TempCheck()
        {
            string group = "";
            string group1 = "";
            DataTable dt2 = new DataTable();

            var result1 = from tab in dt.AsEnumerable()
                          orderby tab["TH004"]
                          group tab by tab["TH004"]
                          into groupDt
                          select new groupDtInfo()
                          {
                              TH004 = groupDt.Key,
                              Sum = groupDt.Sum((r) => decimal.Parse(r["TH007"].ToString()))
                          };

            var result2 = from tab in dsPURTD.Tables["TEMPdsPURTD"].AsEnumerable()
                          orderby tab["TD004"]
                          group tab by tab["TD004"]
                          into PURTD
                          select new PURTDInfo()
                          {
                              TD004 = PURTD.Key,
                              Sum = PURTD.Sum((r) => decimal.Parse(r["TD008"].ToString()))
                          };

            //foreach (groupDtInfo info in result1)
            //{
            //    group = group + string.Format("{0}-{1}", info.TH004, info.Sum) + "\r\n";
            //}


            //foreach (PURTDInfo info2 in result2)
            //{
            //    group1 = group1 + string.Format("{0}-{1}", info2.TD004, info2.Sum) + "\r\n";
            //}

            group = group + "符合數量如下:" + "\r\n";
            group1 = group1 + "不符合數量如下:" + "\r\n";


            //找存在，再判斷採購量是否等罣進貨量
            foreach (PURTDInfo info2 in result2)
            {
                foreach (groupDtInfo info in result1)
                {
                    if (info2.TD004.ToString().Equals(info.TH004.ToString()))
                    {
                        if (info2.Sum == info.Sum)
                        {
                            group = group + string.Format("{0}-{1:N0}", info2.TD004, info2.Sum) + "\r\n";
                        }
                        else if (info2.Sum != info.Sum)
                        {
                            group1 = group1 + string.Format("{0}-採購量{1:N0}<>進貨量{2:N0} ", info2.TD004, info2.Sum, info.Sum) + "\r\n";
                        }
                    }

                }
            }

            //找不存在的
            foreach (PURTDInfo info2 in result2)
            {
                string NotExists = "N";
                foreach (groupDtInfo info in result1)
                {
                    if (info2.TD004.ToString().Equals(info.TH004.ToString()))
                    {
                        NotExists = "Y";
                    }

                }

                if (NotExists.Equals("N"))
                {
                    group1 = group1 + string.Format("{0}-採購量{1:N0}<>進貨量{2:N0} ", info2.TD004, info2.Sum, "0") + "\r\n";
                }
            }


            textBox9.Text = group.ToString();
            textBox10.Text = group1.ToString();


        }

        public void ADDToERP()
        {
            string TH001=null;
            string TH002 = null;

            if(comboBox2.Text.ToString().Equals("A331"))
            {
                TH001 = "A341";
            }
            else if (comboBox2.Text.ToString().Equals("A332"))
            {
                TH001 = "A343";
            }
            else if(comboBox2.Text.ToString().Equals("A333"))
            {
                TH001 = "A347";
            }

            TH002=GetMaxID(TH001);
            sbSql.Clear();
            sbSql.Append(" ");
            sbSql.Append(" INSERT INTO [test].[dbo].[PURTG]");
            sbSql.Append(" ([COMPANY],[CREATOR],[USR_GROUP],[CREATE_DATE],[MODIFIER],[MODI_DATE],[FLAG],[CREATE_TIME],[MODI_TIME],[TRANS_TYPE],[TRANS_NAME],[sync_date],[sync_time],[sync_mark],[sync_count],[DataUser],[DataGroup]");
            sbSql.Append(" ,[TG001],[TG002],[TG003],[TG004],[TG005],[TG006],[TG007],[TG008],[TG009],[TG010]");
            sbSql.Append(" ,[TG011],[TG012],[TG013],[TG014],[TG015],[TG016],[TG017],[TG018],[TG019],[TG020]");
            sbSql.Append(" ,[TG021],[TG022],[TG023],[TG024],[TG025],[TG026],[TG027],[TG028],[TG029],[TG030]");
            sbSql.Append(" ,[TG031],[TG032],[TG033],[TG034],[TG035],[TG036],[TG037],[TG038],[TG039],[TG040]");
            sbSql.Append(" ,[TG041],[TG042],[TG043],[TG044],[TG045],[TG046],[TG047],[TG048],[TG049],[TG050]");
            sbSql.Append(" ,[TG051],[TG052],[TG053],[TG054],[TG055],[TG056],[TG057],[TG058],[TG059],[TG060]");
            sbSql.Append(" ,[TG061],[TG062],[TG063],[TG064],[TG065],[TG066],[TG067],[TG068],[TG069],[TG070]");
            sbSql.Append(" ,[TG071],[TG072],[TG073],[TG074],[TG075],[TG076],[TG077],[TG078],[TG079],[TG080])");
            sbSql.Append(" SELECT  ");
            sbSql.Append(" PURTC.COMPANY,  PURTC.CREATOR, PURTC.USR_GROUP,PURTC.CREATE_DATE, PURTC.MODIFIER, PURTC.MODI_DATE, PURTC.FLAG, PURTC.CREATE_TIME, PURTC.MODI_TIME, PURTC.TRANS_TYPE, PURTC.TRANS_NAME,PURTC.sync_date, PURTC.sync_time, PURTC.sync_mark, PURTC.sync_count,PURTC.DataUser,PURTC.DataGroup");
            sbSql.AppendFormat(" , '{0}' AS TG001, '{1}' AS TG002, '{2}' AS TG003, TC010 AS TG004, TC004 AS TG005, '' AS TG006, MA021  AS TG007, '1' AS TG008, MA030  AS TG009, MA044 AS TG010",TH001,TH002, NowDay);
            sbSql.AppendFormat(" , '{0}' AS TG011, '0' AS TG012, 'N' AS TG013, '{1}' AS TG014, 'N' AS TG015, '0' AS TG016, '0' AS TG017, '0' AS TG018, '0' AS TG019, '0' AS TG020", textBox11.Text.ToString(), NowDay);
            sbSql.AppendFormat(" ,MA002 AS TG021, MA005 AS TG022, '1' AS TG023, 'N' AS TG024, '0' AS TG025, '0' AS TG026, '{0}' AS TG027, '0' AS TG028, '{1}' AS TG029, '0.05' AS TG030", dateTimePicker1.Value.ToString("yyyyMMdd"), dateTimePicker1.Value.ToString("yyyyMMdd").Substring(0,6));
            sbSql.AppendFormat(" , '0' AS TG031, '0' AS TG032, MA055  AS TG033, '{0}' AS TG034, '{1}' AS TG035, '' AS TG036, '' AS TG037, '0' AS TG038, '0' AS TG039, '0' AS TG040", comboBox2.Text.ToString(), textBox1.Text.ToString());
            sbSql.Append(" , '0' AS TG041, 'N' AS TG042, 'Y' AS TG043, 'N' AS TG044, '0' AS TG045, '0' AS TG046, '' AS TG047, '' AS TG048, '' AS TG049, '' AS TG050");
            sbSql.Append(" , NULL AS TG051, NULL AS TG052, NULL AS TG053, NULL AS TG054, NULL AS TG055, NULL AS TG056, NULL AS TG057, NULL AS TG058, NULL AS TG059, NULL AS TG060");
            sbSql.Append(" , NULL AS TG061, NULL AS TG062, NULL AS TG063, NULL AS TG064, NULL AS TG065, NULL AS TG066, NULL AS TG067, NULL AS TG068, NULL AS TG069, NULL AS TG070");
            sbSql.Append(" , NULL AS TG071, NULL AS TG072, NULL AS TG073, NULL AS TG074, NULL AS TG075, NULL AS TG076, NULL AS TG077, NULL AS TG078, NULL AS TG079, NULL AS TG080");
            sbSql.AppendFormat("  FROM TK.dbo.PURTC,TK.dbo.PURMA WHERE TC004=MA001 AND  TC001='{0}' AND TC002='{1}'", comboBox2.Text.ToString(), textBox1.Text.ToString());
            sbSql.AppendFormat(" ");

        }

        public string GetMaxID(string TG001)
        {
            string newid;
            int countid;
            NowDay = DateTime.Now.ToString("yyyyMMdd");
            StringBuilder sbSql=new StringBuilder();
            sbSql.AppendFormat(@"SELECT( CASE WHEN ISNULL(MAX(TG002),'')='' THEN '0' ELSE  MAX(TG002)  END) AS TG002  FROM PURTG WITH (NOLOCK) WHERE TG003='{0}' AND TG001='{1}' ", NowDay, TG001);

            DataSet dt = new DataSet();
            connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
            sqlConn = new SqlConnection(connectionString);
            SqlCommand cmd = new SqlCommand(sbSql.ToString(), sqlConn);

            sqlConn.Open();
            adapter = new SqlDataAdapter(cmd);
            dt.Clear();
            adapter.Fill(dt);

            newid = dt.Tables[0].Rows[0][0].ToString();
            countid = Convert.ToInt16(newid.Substring(8, 3));
            countid = countid + 1;
            newid = NowDay+countid.ToString().PadLeft(3, '0');

            return newid;
        }
        #endregion

        #region BUTTON
        private void button1_Click(object sender, EventArgs e)
        {
            Search();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            TempAdd();
        }
        private void button3_Click(object sender, EventArgs e)
        {
            TempUpdate();
        }

        private void button4_Click(object sender, EventArgs e)
        {
            TempDelete();
        }
        private void button6_Click(object sender, EventArgs e)
        {
            TempCheck();
        }

        private void button5_Click(object sender, EventArgs e)
        {
            ADDToERP();
        }
        #endregion

        #region gridview

        private void dataGridView1_SelectionChanged(object sender, EventArgs e)
        {
            if(!string.IsNullOrEmpty(dataGridView1.CurrentRow.Cells[1].Value.ToString()))
            {
                textBox4.Text = dataGridView1.CurrentRow.Cells[1].Value.ToString();
                
            }
           
        }

        private void dataGridView2_SelectionChanged(object sender, EventArgs e)
        {
            if (!string.IsNullOrEmpty(dataGridView2.CurrentRow.Cells[6].Value.ToString()))
            {
                textBox6.Text = dataGridView2.CurrentRow.Cells[6].Value.ToString();
                textBox5.Text = dataGridView2.CurrentRow.Cells[2].Value.ToString();
                comboBox3.Text = dataGridView2.CurrentRow.Cells[5].Value.ToString();
            }
        }





        #endregion

        
    }
}

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
            comboBox1.DataSource = dt.DefaultView;
            comboBox1.ValueMember = "MC001";
            comboBox1.DisplayMember = "MC002";
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
            dt.Rows[dataGridView2.CurrentCell.RowIndex]["TH009"] = comboBox3.SelectedValue.ToString();
            dt.Rows[dataGridView2.CurrentCell.RowIndex]["TH010"] = textBox6.Text.ToString();

        }

        public void TempDelete()
        {
            dt.Rows.RemoveAt(dataGridView2.CurrentCell.RowIndex);
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
        private void button3_Click(object sender, EventArgs e)
        {
            TempUpdate();
        }

        private void button4_Click(object sender, EventArgs e)
        {
            TempDelete();
        }


        #endregion


    }
}

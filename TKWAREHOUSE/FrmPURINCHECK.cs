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
using System.Configuration;
using System.Text.RegularExpressions;
using FastReport;
using FastReport.Data;
using TKITDLL;
using System.Data.OleDb;
using System.Net;
using AForge.Video;
using AForge.Video.DirectShow;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;
using System.Threading;
using System.IO.Ports;
using System.Threading;
using System.IO.Ports;


namespace TKWAREHOUSE
{
    public partial class FrmPURINCHECK : Form
    {
        int CommandTimeout = 180;
        StringBuilder sbSql = new StringBuilder();
        SqlConnection sqlConn = new SqlConnection();
        SqlDataAdapter adapter = new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilder = new SqlCommandBuilder();
        SqlTransaction tran;
        SqlCommand cmd = new SqlCommand();
        DataSet ds = new DataSet();
        DataTable dt = new DataTable();
        string tablename = null;
        int rownum = 0;
        DataGridViewRow row;
        int result;
        public FrmPURINCHECK()
        {
            InitializeComponent();
        }

        public FrmPURINCHECK(string ID)
        {
            InitializeComponent();

            textBox1.Text = ID;
        }

        #region FUNCTION

        #endregion


        #region BUTTON
        private void button1_Click(object sender, EventArgs e)
        {

        }
        #endregion
    }
}

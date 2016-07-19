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
using System.Diagnostics;

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
        string NowDB = "test";

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
            //=====不可重複開啟同一支程式(將程式碼放至於偵測是否被開啟的程式中即可)=====
            //先宣告一個process來存放系統所有開啟的程序 
            //Process[] p = Process.GetProcesses();
            ////迴圈檢查程序檔名是否相符如果是就關閉 
            //foreach (Process pp in p)
            //{
            //    if (pp.ProcessName == "TKWAREHOUSE")
            //        pp.Kill();
            //}

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
            dt.Columns.Add(new DataColumn("TH018", typeof(string)));
            dt.Columns.Add(new DataColumn("TH011", typeof(string)));
            dt.Columns.Add(new DataColumn("TH012", typeof(string)));
            dt.Columns.Add(new DataColumn("TH013", typeof(string)));
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
                    sbSql.AppendFormat(@"SELECT TD003,TD004,TD005,TD006,TD007,TD008-(SELECT ISNULL(SUM(TH007),0) FROM [{1}].[dbo].PURTH WITH (NOLOCK) WHERE TH011=TD001 AND TH012=TD002 AND TH013=TD003 ) AS TD008,TD009,TD010,TD001,TD002,TD003 FROM [{2}].[dbo].PURTD WITH (NOLOCK) {0} ", sbSqlQuery.ToString(),NowDB, NowDB);

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
            if(!string.IsNullOrEmpty(textBox2.Text.ToString())&&!textBox2.Text.ToString().Equals("") && !string.IsNullOrEmpty(textBox3.Text.ToString()) && !textBox3.Text.ToString().Equals(""))
            {
                DataRow dr;

                //dt.Columns.Add(new DataColumn("TH010", typeof(string)));
                // 加入第一筆資料
                dr = dt.NewRow();
                dr["TH004"] = textBox4.Text.ToString();
                dr["TH005"] = dataGridView1.CurrentRow.Cells[2].Value.ToString();
                dr["TH007"] = textBox3.Text.ToString();
                dr["TH008"] = dataGridView1.CurrentRow.Cells[6].Value.ToString();
                dr["TH009"] = comboBox1.SelectedValue.ToString();
                dr["TH009CH"] = comboBox1.Text.ToString();
                dr["TH010"] = textBox2.Text.ToString();
                dr["TH018"] = dataGridView1.CurrentRow.Cells[7].Value.ToString();
                dr["TH011"] = dataGridView1.CurrentRow.Cells[8].Value.ToString();
                dr["TH012"] = dataGridView1.CurrentRow.Cells[9].Value.ToString();
                dr["TH013"] = dataGridView1.CurrentRow.Cells[10].Value.ToString();
                dt.Rows.Add(dr);
            }
            




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
            textBox7.Text = "go";
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

            TH002 = GetMaxID(TH001);

            try
            {
                //add ZWAREWHOUSEPURTH
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sqlConn.Close();
                sqlConn.Open();
                tran = sqlConn.BeginTransaction();

                sbSql.Clear();
                sbSql.Append(" ");
                sbSql.AppendFormat("DELETE  [{0}].[dbo].[ZWAREWHOUSEPURTH] WHERE TH001='{1}' AND  TH002='{2}'", NowDB, TH001.ToString(), TH002.ToString());
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    
                    sbSql.AppendFormat(" INSERT INTO [{0}].[dbo].[ZWAREWHOUSEPURTH] ",NowDB);
                    sbSql.Append(" ([TH001],[TH002],[TH003],[TH004],[TH005],[TH007],[TH008],[TH009],[TH009CH],[TH010],[TH018],[TH011],[TH012],[TH013])");
                    sbSql.AppendFormat(" VALUES ('{0}','{1}','{2}','{3}','{4}','{5}','{6}','{7}','{8}','{9}','{10}','{11}','{12}','{13}')", TH001.ToString(), TH002.ToString(), (i + 1).ToString().PadLeft(4, '0'), dt.Rows[i]["TH004"].ToString(), dt.Rows[i]["TH005"].ToString(), dt.Rows[i]["TH007"].ToString(), dt.Rows[i]["TH008"].ToString(), dt.Rows[i]["TH009"].ToString(), dt.Rows[i]["TH009CH"].ToString(), dt.Rows[i]["TH010"].ToString(), dt.Rows[i]["TH018"].ToString(), dt.Rows[i]["TH011"].ToString(), dt.Rows[i]["TH012"].ToString(), dt.Rows[i]["TH013"].ToString());
                    sbSql.Append(" ");
                }
                sbSql.Append(" ");

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

                

                //add PURTH+PURTG
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);


                sqlConn.Open();
                tran = sqlConn.BeginTransaction();

                sbSql.Clear();
                sbSql.AppendFormat(" INSERT INTO [{0}].[dbo].PURTH", NowDB);
                sbSql.Append(" ([COMPANY],[CREATOR],[USR_GROUP],[CREATE_DATE],[MODIFIER],[MODI_DATE],[FLAG],[CREATE_TIME],[MODI_TIME],[TRANS_TYPE],[TRANS_NAME],[sync_date],[sync_time],[sync_mark],[sync_count],[DataUser],[DataGroup]");
                sbSql.Append(" ,[TH001],[TH002],[TH003],[TH004],[TH005],[TH006],[TH007],[TH008],[TH009],[TH010]");
                sbSql.Append(" ,[TH011],[TH012],[TH013],[TH014],[TH015],[TH016],[TH017],[TH018],[TH019],[TH020]");
                sbSql.Append(" ,[TH021],[TH022],[TH023],[TH024],[TH025],[TH026],[TH027],[TH028],[TH029],[TH030]");
                sbSql.Append(" ,[TH031],[TH032],[TH033],[TH034],[TH035],[TH036],[TH037],[TH038],[TH039],[TH040]");
                sbSql.Append(" ,[TH041],[TH042],[TH043],[TH044],[TH045],[TH046],[TH047],[TH048],[TH049],[TH050]");
                sbSql.Append(" ,[TH051],[TH052],[TH053],[TH054],[TH055],[TH056],[TH057],[TH058],[TH059],[TH060]");
                sbSql.Append(" ,[TH061],[TH062],[TH063],[TH064],[TH065],[TH066],[TH067],[TH068],[TH069],[TH070]");
                sbSql.Append(" ,[TH071],[TH072],[TH073],[TH074],[TH075],[TH076],[TH077],[TH078],[TH079],[TH080]");
                sbSql.Append(" ,[TH081],[TH082],[TH083],[TH084],[TH085],[TH086],[TH087],[TH088],[TH089],[TH090]");
                sbSql.Append(" ,[TH091],[TH092],[TH093],[TH094],[TH095],[TH096],[TH097],[TH098])");
                sbSql.Append(" SELECT ");
                sbSql.Append(" 'TK' AS [COMPANY], 'DS' AS  [CREATOR], 'DS' AS  [USR_GROUP],SUBSTRING(TH002,1,8) AS  [CREATE_DATE], NULL AS  [MODIFIER], NULL AS  [MODI_DATE], '0' AS  [FLAG],'12:00:01' AS  [CREATE_TIME], NULL AS  [MODI_TIME], 'P001' AS  [TRANS_TYPE], 'PURI09' AS  [TRANS_NAME], NULL AS  [sync_date], NULL AS  [sync_time], NULL AS  [sync_mark], '0' AS  [sync_count], 'DS' AS  [DataUser], 'DS' AS  [DataGroup]");
                sbSql.Append(" , TH001 AS  [TH001], TH002 AS  [TH002], TH003 AS  [TH003], TH004 AS  [TH004], TH005 AS  [TH005], NULL AS  [TH006], TH007 AS  [TH007], TH008 AS  [TH008], TH009 AS  [TH009], TH010 AS  [TH010]");
                sbSql.Append(" , TH011 AS  [TH011], TH012 AS  [TH012], TH013 AS  [TH013], SUBSTRING(TH002,1,8) AS  [TH014], TH007 AS  [TH015], TH007 AS  [TH016], 0 AS  [TH017], TH018 AS  [TH018], TH007*TH018 AS  [TH019], 0 AS  [TH020]");
                sbSql.Append(" , NULL AS  [TH021], NULL AS  [TH022], NULL AS  [TH023], 0 AS  [TH024], NULL AS  [TH025], 'N' AS  [TH026],  'N' AS  [TH027],  '2' AS  [TH028],  'N' AS  [TH029],  'N' AS  [TH030] ");
                sbSql.Append(" ,  'N' AS  [TH031],  'N' AS  [TH032], NULL AS  [TH033], 0 AS  [TH034], NULL AS  [TH035], NULL AS  [TH036], NULL AS  [TH037],  'DS' AS  [TH038], NULL AS  [TH039], NULL AS  [TH040]");
                sbSql.Append(" , NULL AS  [TH041], NULL AS  [TH042],  'N' AS  [TH043],  'N' AS  [TH044], TH007*TH018 AS  [TH045], ROUND(TH007*TH018*0.05,0) AS  [TH046], TH007*TH018 AS  [TH047], ROUND(TH007*TH018*0.05,0) AS  [TH048], 0 AS  [TH049], 0 AS  [TH050]");
                sbSql.Append(" , 0 AS  [TH051], 0 AS  [TH052], NULL AS  [TH053], 'N' AS  [TH054], 0 AS  [TH055], TH008 AS  [TH056], NULL AS  [TH057], 'N' AS  [TH058], 00 AS  [TH059], 0 AS  [TH060]");
                sbSql.Append(" , NULL AS  [TH061], NULL AS  [TH062], NULL AS  [TH063], NULL AS  [TH064], NULL AS  [TH065], NULL AS  [TH066], 0 AS  [TH067], 0 AS  [TH068], NULL AS  [TH069], NULL AS  [TH070]");
                sbSql.Append(" , NULL AS  [TH071], NULL AS  [TH072], 0 AS  [TH073], 'N' AS  [TH074], NULL AS  [TH075], NULL AS  [TH076], 0 AS  [TH077], '2' AS  [TH078], NULL AS  [TH079], NULL AS  [TH080]");
                sbSql.Append(" , NULL AS  [TH081], NULL AS  [TH082], NULL AS  [TH083], NULL AS  [TH084], NULL AS  [TH085], NULL AS  [TH086], 'N' AS  [TH087], NULL AS  [TH088], NULL AS  [TH089], NULL AS  [TH090]");
                sbSql.Append(" , NULL AS  [TH091], 'N' AS  [TH092], NULL AS  [TH093], 0 AS  [TH094], NULL AS  [TH095], NULL AS  [TH096], NULL AS  [TH097], 0 AS  [TH098]");
                sbSql.AppendFormat(" FROM  [{0}].[dbo].[ZWAREWHOUSEPURTH] ", NowDB);
                sbSql.AppendFormat(" WHERE  TH001='{0}' AND TH002='{1}'",TH001,TH002);
                sbSql.Append(" ");

                sbSql.Append(" ");
                sbSql.AppendFormat(" INSERT INTO [{0}].[dbo].[PURTG]", NowDB);
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
                sbSql.AppendFormat(" , '{0}' AS TG001, '{1}' AS TG002, '{2}' AS TG003, TC010 AS TG004, TC004 AS TG005, '' AS TG006, MA021  AS TG007, '1' AS TG008, MA030  AS TG009, MA044 AS TG010", TH001, TH002, NowDay);
                sbSql.AppendFormat(" , '{0}' AS TG011, '0' AS TG012, 'N' AS TG013, '{1}' AS TG014, 'N' AS TG015, '0' AS TG016, '0' AS TG017, '0' AS TG018, '0' AS TG019, '0' AS TG020", textBox11.Text.ToString(), NowDay);
                sbSql.AppendFormat(" ,MA002 AS TG021, MA005 AS TG022, '1' AS TG023, 'N' AS TG024, '0' AS TG025, '0' AS TG026, '{0}' AS TG027, '0' AS TG028, '{1}' AS TG029, '0.05' AS TG030", dateTimePicker1.Value.ToString("yyyyMMdd"), dateTimePicker1.Value.ToString("yyyyMMdd").Substring(0, 6));
                sbSql.AppendFormat(" , '0' AS TG031, '0' AS TG032, MA055  AS TG033, '{0}' AS TG034, '{1}' AS TG035, '' AS TG036, '' AS TG037, '0' AS TG038, '0' AS TG039, '0' AS TG040", comboBox2.Text.ToString(), textBox1.Text.ToString());
                sbSql.Append(" , '0' AS TG041, 'N' AS TG042, 'Y' AS TG043, 'N' AS TG044, '0' AS TG045, '0' AS TG046, '' AS TG047, '' AS TG048, '' AS TG049, '' AS TG050");
                sbSql.Append(" , NULL AS TG051, NULL AS TG052, NULL AS TG053, NULL AS TG054, NULL AS TG055, NULL AS TG056, NULL AS TG057, NULL AS TG058, NULL AS TG059, NULL AS TG060");
                sbSql.Append(" , NULL AS TG061, NULL AS TG062, NULL AS TG063, NULL AS TG064, NULL AS TG065, NULL AS TG066, NULL AS TG067, NULL AS TG068, NULL AS TG069, NULL AS TG070");
                sbSql.Append(" , NULL AS TG071, NULL AS TG072, NULL AS TG073, NULL AS TG074, NULL AS TG075, NULL AS TG076, NULL AS TG077, NULL AS TG078, NULL AS TG079, NULL AS TG080");
                sbSql.AppendFormat("  FROM [{2}].dbo.PURTC,TK.dbo.PURMA WHERE TC004=MA001 AND  TC001='{0}' AND TC002='{1}'", comboBox2.Text.ToString(), textBox1.Text.ToString(), NowDB);
                sbSql.Append(" ");

                sbSql.AppendFormat(" UPDATE [{0}].dbo.PURTG SET ",NowDB);
                sbSql.Append(" TG017=(SELECT SUM(TH045) FROM  [test].dbo.PURTH WHERE TG001=TH001 AND TG002=TH002)");
                sbSql.Append(" ,TG019=(SELECT SUM(TH046) FROM  [test].dbo.PURTH WHERE TG001=TH001 AND TG002=TH002)  ");
                sbSql.Append(" ,TG025=(SELECT SUM(TH007) FROM  [test].dbo.PURTH WHERE TG001=TH001 AND TG002=TH002) ");
                sbSql.Append(" ,TG026=(SELECT SUM(TH007) FROM  [test].dbo.PURTH WHERE TG001=TH001 AND TG002=TH002)");
                sbSql.Append(" ,TG031=(SELECT SUM(TH045) FROM  [test].dbo.PURTH WHERE TG001=TH001 AND TG002=TH002)");
                sbSql.Append(" ,TG032=(SELECT SUM(TH046) FROM  [test].dbo.PURTH WHERE TG001=TH001 AND TG002=TH002) ");
                sbSql.Append(" ,TG040=(SELECT SUM(TH007) FROM  [test].dbo.PURTH WHERE TG001=TH001 AND TG002=TH002) ");
                sbSql.AppendFormat(" WHERE TG001='{0}' AND TG002='{1}'",TH001,TH002);

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
                    textBox7.Text = TH001.ToString();
                    textBox8.Text = TH002.ToString();
                }

                sqlConn.Close();

                //UPDATE PURTG
               

            }
            catch
            {

            }

            finally
            {
                sqlConn.Close();
            }

            



        }

        public string GetMaxID(string TG001)
        {
            string newid;
            int countid;
            NowDay = DateTime.Now.ToString("yyyyMMdd");
            StringBuilder sbSql=new StringBuilder();
            sbSql.AppendFormat(@"SELECT( CASE WHEN ISNULL(MAX(TG002),'')='' THEN '0' ELSE  MAX(TG002)  END) AS TG002  FROM  [{2}].dbo.PURTG WITH (NOLOCK) WHERE TG003='{0}' AND TG001='{1}' ", NowDay, TG001,NowDB);

            DataSet dt = new DataSet();
            connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
            sqlConn = new SqlConnection(connectionString);
            SqlCommand cmd = new SqlCommand(sbSql.ToString(), sqlConn);

            sqlConn.Open();
            adapter = new SqlDataAdapter(cmd);
            dt.Clear();
            adapter.Fill(dt);

            newid = dt.Tables[0].Rows[0][0].ToString();
            if(newid.ToString().Equals("0"))
            {
                countid = 0;
            }
            else
            {
                countid = Convert.ToInt16(newid.Substring(8, 3));
            }
            

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
            var curRow = dataGridView1.CurrentRow;
            if (curRow != null)
            {
                textBox4.Text = dataGridView1.CurrentRow.Cells[1].Value.ToString();
            }
                       
           
        }

        private void dataGridView2_SelectionChanged(object sender, EventArgs e)
        {
            var curRow = dataGridView2.CurrentRow;
            if (curRow != null)
            {
                textBox6.Text = dataGridView2.CurrentRow.Cells[6].Value.ToString();
                textBox5.Text = dataGridView2.CurrentRow.Cells[2].Value.ToString();
                comboBox3.Text = dataGridView2.CurrentRow.Cells[5].Value.ToString();
            }
          
        }





        #endregion

        
    }
}

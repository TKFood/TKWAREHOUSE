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
using TKITDLL;

namespace TKWAREHOUSE
{
    public partial class FrmINVCHECKLOT : Form
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
        DataSet ds2 = new DataSet();
        string tablename = null;
        int rownum = 0;
        DataGridViewRow row;
        string SALSESID = null;
        int result;

        public Report report1 { get; private set; }

        public FrmINVCHECKLOT()
        {
            InitializeComponent();

            comboboxload1();
        }

        #region FUNCTION

        public void comboboxload1()
        {


            //20210902密
            Class1 TKID = new Class1();//用new 建立類別實體
            SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

            //資料庫使用者密碼解密
            sqlsb.Password = TKID.Decryption(sqlsb.Password);
            sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

            String connectionString;
            sqlConn = new SqlConnection(sqlsb.ConnectionString);

            string Sequel = "SELECT MC001,MC002 FROM[TK].dbo.CMSMC WHERE MC001 NOT LIKE '1%' ORDER BY MC001";
            
            SqlDataAdapter da = new SqlDataAdapter(Sequel, sqlConn);
            DataTable dt = new DataTable();
            sqlConn.Open();

            dt.Columns.Add("MC001", typeof(string));
            dt.Columns.Add("MC002", typeof(string));
            da.Fill(dt);
            comboBox1.DataSource = dt.DefaultView;
            comboBox1.ValueMember = "MC001";
            comboBox1.DisplayMember = "MC001";
            sqlConn.Close();

            comboBox1.SelectedValue = "20001";

        }
        public void SETFASTREPORT(string MC001, string MB001, string LOTNO)
        {

            string SQL;

            report1 = new Report();
            report1.Load(@"REPORT\庫別未核準的品號+批號.frx");

            //20210902密
            Class1 TKID = new Class1();//用new 建立類別實體
            SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

            //資料庫使用者密碼解密
            sqlsb.Password = TKID.Decryption(sqlsb.Password);
            sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

            String connectionString;
            sqlConn = new SqlConnection(sqlsb.ConnectionString);

            report1.Dictionary.Connections[0].ConnectionString = sqlsb.ConnectionString;

            TableDataSource Table = report1.GetDataSource("Table") as TableDataSource;

            SQL = SETFASETSQL(MC001,MB001,LOTNO);

            Table.SelectCommand = SQL;

            report1.Preview = previewControl1;
            report1.Show();

        }

        public string SETFASETSQL(string MC001,string MB001,string LOTNO)
        {
            StringBuilder FASTSQL = new StringBuilder();
            StringBuilder STRQUERY = new StringBuilder();

            //MB001
            if (!string.IsNullOrEmpty(MB001))
            {
                STRQUERY.AppendFormat(@" AND TH004 LIKE '%{0}%' ", MB001);
            }
            else
            {
                STRQUERY.AppendFormat(@" ");
            }

            //LOTNO
            if (!string.IsNullOrEmpty(LOTNO))
            {
                STRQUERY.AppendFormat(@" AND TH017 LIKE '%{0}%' ", LOTNO);
            }
            else
            {
                STRQUERY.AppendFormat(@" ");
            }


            FASTSQL.AppendFormat(@" 
                                    SELECT TH004 AS '品號',MB002 AS '品名',TH017  AS '批號', Type AS '分類',Key1 AS '單別',Key2 AS '單號',CONVERT(NVARCHAR,Key3)  AS '數量',M_MF002 AS '申請人'
                                    FROM ( 
                                    SELECT TH004,TH017,'銷貨' As Type, TG001 As Key1, TG002 As Key2 ,(TH008+TH024) AS Key3
                                    ,(CASE WHEN (COPTG.MODIFIER <> '') THEN B.MF002 ELSE A.MF002 END) As M_MF002 
                                    FROM TK..COPTG AS COPTG
                                    Left Join TK..COPTH AS COPTH ON TH001=TG001 AND TH002=TG002
                                    Left Join TK..ADMMF As A On A.MF001=COPTG.CREATOR
                                    Left Join TK..ADMMF As B On B.MF001=COPTG.MODIFIER
                                    Where  TH007='{0}' AND TG023='N'
                                    UNION ALL 
                                    SELECT TB004,TB014, (CASE WHEN (MQ003 = '11') THEN '轉撥' ELSE '庫存異動' END) As Type 
                                    ,TA001 As Key1, TA002 As Key2  ,(TB007) AS Key3
                                    ,(CASE WHEN (INVTA.MODIFIER <> '') THEN B.MF002 ELSE A.MF002 END) As M_MF002 
                                    FROM TK..INVTA AS INVTA
                                    LEFT JOIN TK..INVTB AS INVTB ON TB001=TA001 AND TB002=TA002
                                    LEFT JOIN TK..CMSMQ AS CMSMQ ON MQ001=TA001
                                    Left Join TK..ADMMF As A On A.MF001=INVTA.CREATOR
                                    Left Join TK..ADMMF As B On B.MF001=INVTA.MODIFIER
                                    WHERE  TB012='{0}' AND TA006='N' AND MQ010=-1 
                                    UNION ALL 
                                    SELECT TG004,TG017, '借出入轉撥' As Type 
                                    ,TF001 As Key1, TF002 As Key2 ,(TG009) AS Key3
                                    ,(CASE WHEN (INVTF.MODIFIER <> '') THEN B.MF002 ELSE A.MF002 END) As M_MF002 
                                    FROM TK..INVTF AS INVTF
                                    LEFT JOIN TK..INVTG AS INVTG ON TG001=TF001 AND TG002=TF002
                                    Left Join TK..ADMMF As A On A.MF001=INVTF.CREATOR
                                    Left Join TK..ADMMF As B On B.MF001=INVTF.MODIFIER
                                    WHERE  TG007='{0}' AND TF020='N' 
                                    UNION ALL 
                                    SELECT TI004,TI017, '借出入歸還' As Type 
                                    ,TH001 As Key1, TH002 As Key2 ,(TI009) AS Key3
                                    ,(CASE WHEN (INVTH.MODIFIER <> '') THEN B.MF002 ELSE A.MF002 END) As M_MF002 
                                    FROM TK..INVTH AS INVTH
                                    LEFT JOIN TK..INVTI AS INVTI ON TI001=TH001 AND TI002=TH002
                                    Left Join TK..ADMMF As A On A.MF001=INVTH.CREATOR
                                    Left Join TK..ADMMF As B On B.MF001=INVTH.MODIFIER
                                    WHERE  TI007='{0}' AND TH020='N' 
                                    UNION ALL 
                                    SELECT TB007,TB019, '出貨通知' As Type 
                                    ,TA001 As Key1, TA002 As Key2 ,(TB009+TB011) AS Key3
                                    ,(CASE WHEN (EPSTA.MODIFIER <> '') THEN B.MF002 ELSE A.MF002 END) As M_MF002 
                                    FROM TK..EPSTA AS EPSTA
                                    LEFT JOIN TK..EPSTB AS EPSTB ON TB001=TA001 AND TB002=TA002
                                    Left Join TK..ADMMF As A On A.MF001=EPSTA.CREATOR
                                    Left Join TK..ADMMF As B On B.MF001=EPSTA.MODIFIER
                                    WHERE  TB018='{0}' AND TA034<>'V' 
                                    AND TB021+TB022+TB023='''' AND TB042+TB043+TB044='' 
                                    UNION ALL 
                                    SELECT TE004,TE013, '組合單' As Type 
                                    ,TD001 As Key1, TD002 As Key2 ,(TE008) AS Key3
                                    ,(CASE WHEN (BOMTD.MODIFIER <> '') THEN B.MF002 ELSE A.MF002 END) As M_MF002 
                                    FROM TK..BOMTD AS BOMTD
                                    LEFT JOIN TK..BOMTE AS BOMTE ON TE001=TD001 AND TE002=TD002
                                    Left Join TK..ADMMF As A On A.MF001=BOMTD.CREATOR
                                    Left Join TK..ADMMF As B On B.MF001=BOMTD.MODIFIER
                                    WHERE  TE007='{0}' AND TD012='N' 
                                    UNION ALL 
                                    SELECT TF004,TF015, '拆解單' As Type 
                                    ,TF001 As Key1, TF002 As Key2 ,(TF007) AS Key3
                                    ,(CASE WHEN (BOMTF.MODIFIER <> '') THEN B.MF002 ELSE A.MF002 END) As M_MF002 
                                    FROM TK..BOMTF AS BOMTF
                                    Left Join TK..ADMMF As A On A.MF001=BOMTF.CREATOR
                                    Left Join TK..ADMMF As B On B.MF001=BOMTF.MODIFIER
                                    WHERE  TF008='{0}' AND TF010='N' 
                                    UNION ALL 
                                    SELECT TE004,TE010, '領/退料單' As Type 
                                    ,TC001 As Key1, TC002 As Key2 ,(TE005) AS Key3
                                    ,(CASE WHEN (MOCTC.MODIFIER <> '') THEN B.MF002 ELSE A.MF002 END) As M_MF002 
                                    FROM TK..MOCTC AS MOCTC
                                    LEFT JOIN TK..MOCTE AS MOCTE ON TE001=TC001 AND TE002=TC002
                                    Left Join TK..ADMMF As A On A.MF001=MOCTC.CREATOR
                                    Left Join TK..ADMMF As B On B.MF001=MOCTC.MODIFIER
                                    LEFT JOIN TK..CMSMQ AS CMSMQ ON MQ001=TC001
                                    WHERE TE008='{0}' AND TC009='N' 
                                    AND MQ010=-1 
                                    UNION ALL 
                                    SELECT TK004,TK018, '成本開帳/調整單' As Type 
                                    ,TJ001 As Key1, TJ002 As Key2 ,SUM(TK007) AS Key3
                                    ,(CASE WHEN (INVTJ.MODIFIER <> '') THEN B.MF002 ELSE A.MF002 END) As M_MF002 
                                    FROM [TK].dbo.INVTJ As INVTJ
                                    LEFT JOIN [TK].dbo.INVTK AS INVTK ON TK001=TJ001 AND TK002=TJ002
                                    Left Join [TK].dbo.ADMMF As A On A.MF001=INVTJ.CREATOR
                                    Left Join [TK].dbo.ADMMF As B On B.MF001=INVTJ.MODIFIER
                                    WHERE  TK017='{0}' AND TJ010='N'  
                                    GROUP BY TK004,TK018,TJ001, TJ002, INVTJ.MODIFIER,A.MF002,B.MF002
                                    HAVING (SUM(ISNULL(TK007,0)) < 0) 
                                    ) AS MoidA 
                                    LEFT JOIN [TK].dbo.INVMB ON MB001=TH004
                                    WHERE 1=1
                                    {1}

                                    ORDER BY TH004,TH017 

                                ",MC001, STRQUERY.ToString());

            return FASTSQL.ToString();
        }


        #endregion

        #region BUTTON
        private void button1_Click(object sender, EventArgs e)
        {
            string MC001 = comboBox1.Text.ToString();
            string MB001 = textBox1.Text.ToString();
            string LOTNO = textBox2.Text.ToString();


            SETFASTREPORT(MC001, MB001, LOTNO);
        }

        #endregion


    }
}

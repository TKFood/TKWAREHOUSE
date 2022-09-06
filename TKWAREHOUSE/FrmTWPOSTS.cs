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
using System.Data.OleDb;

namespace TKWAREHOUSE
{
    public partial class FrmTWPOSTS : Form
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
        DataSet ds1 = new DataSet();
        DataSet ds2 = new DataSet();
        DataSet MAINds = new DataSet();

        string tablename = null;
        int rownum = 0;
        DataGridViewRow row;
        int result;

        string _path = null;
        DataTable EXCEL = null;


        public Report report1 { get; private set; }

        public FrmTWPOSTS()
        {
            InitializeComponent();
        }

        #region FUNCTION
        public void SEARCH(string SDAYS, string EDAYS)
        {
            SqlDataAdapter adapter = new SqlDataAdapter();
            SqlCommandBuilder sqlCmdBuilder = new SqlCommandBuilder();
            DataSet ds1 = new DataSet();
            MAINds.Clear();

            try
            {
                //20210902密
                Class1 TKID = new Class1();//用new 建立類別實體
                SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

                //資料庫使用者密碼解密
                sqlsb.Password = TKID.Decryption(sqlsb.Password);
                sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

                String connectionString;
                sqlConn = new SqlConnection(sqlsb.ConnectionString);


                sbSql.Clear();
                sbSqlQuery.Clear();

                sbSql.AppendFormat(@"  

                                        SELECT 
                                        [ID] AS 'ID'
                                        ,[PAYMONEYS] AS '金額'
                                        ,CONVERT(NVARCHAR,[SENDDATES],112) AS '交寄日期'
                                        ,[WEIGHETS] AS '重量'
                                        ,[ISSINGALS] AS '單筆單件'
                                        ,[CUSTOMERNO] AS '客戶編號'
                                        ,[CUSTOMERNAMES] AS '客戶名稱'
                                        ,[PHONES] AS '電話'
                                        ,[ZIPCODE] AS '郵遞區號'
                                        ,[ADDRESS] AS '地址'
                                        ,[SENDCONTENTS] AS '內裝物品Memo'
                                        ,[SENDNUMS] AS '件數編號'
                                        ,[COMMENTS] AS '備註(出貨單編號)'
                                        ,[USEDUNITS] AS '使用單位編號'
                                        ,[SENDNO] AS '託運單編號'
                                        ,[MOBILEPHONE] AS '手機'
                                        ,[COLMONEYS] AS '代收貨價'

                                        FROM [TKWAREHOUSE].[dbo].[TWPOSTS]
                                        WHERE CONVERT(NVARCHAR,[SENDDATES],112)>='{0}' AND CONVERT(NVARCHAR,[SENDDATES],112)<='{1}'
                                        ORDER BY CONVERT(NVARCHAR,[SENDDATES],112),ID
                                        ", SDAYS, EDAYS);

                adapter = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder = new SqlCommandBuilder(adapter);
                sqlConn.Open();
                ds1.Clear();
                adapter.Fill(ds1, "ds1");
                sqlConn.Close();


                if (ds1.Tables["ds1"].Rows.Count == 0)
                {
                    dataGridView1.DataSource = null;
                }
                else
                {
                    if (ds1.Tables["ds1"].Rows.Count >= 1)
                    {
                        //dataGridView1.Rows.Clear();
                        dataGridView1.DataSource = ds1.Tables["ds1"];
                        dataGridView1.AutoResizeColumns();
                        //dataGridView1.CurrentCell = dataGridView1[0, rownum];
                        MAINds = ds1;
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

        private void dataGridView1_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            dataGridView1.EndEdit();
            int prices = 0;

            string ID = dataGridView1.Rows[e.RowIndex].Cells[0].Value.ToString();
            decimal WEIGHTS = Convert.ToDecimal(dataGridView1.Rows[e.RowIndex].Cells["重量"].Value.ToString());

            prices = SEARCHTWPOSTSBASE(WEIGHTS);

            if (dataGridView1.Rows[e.RowIndex].Cells["單筆單件"].Value.ToString().Equals("y"))
            {
                prices = prices - 10;
                dataGridView1.Rows[e.RowIndex].Cells["單筆單件"].Value = "Y";
                dataGridView1.Rows[e.RowIndex].Cells["金額"].Value = prices;
            }
            else if (dataGridView1.Rows[e.RowIndex].Cells["單筆單件"].Value.ToString().Equals("n"))
            {
                dataGridView1.Rows[e.RowIndex].Cells["單筆單件"].Value = "N";
                dataGridView1.Rows[e.RowIndex].Cells["金額"].Value = prices;
            }
            else if (dataGridView1.Rows[e.RowIndex].Cells["單筆單件"].Value.ToString().Equals("Y"))
            {
                prices = prices - 10;
                dataGridView1.Rows[e.RowIndex].Cells["金額"].Value = prices;
            }
            else if (dataGridView1.Rows[e.RowIndex].Cells["單筆單件"].Value.ToString().Equals("N"))
            {
                dataGridView1.Rows[e.RowIndex].Cells["金額"].Value = prices;
            }


            //MessageBox.Show(prices+" "+ID + " "+e.RowIndex+" "+e.ColumnIndex);
        }

        public void SAVE(DataSet MAINds)
        {
            string ID = null;

            sbSql.Clear();

            if (MAINds.Tables[0].Rows.Count > 0)
            {
                foreach (DataRow DR in MAINds.Tables[0].Rows)
                {
                    sbSql.AppendFormat(@" 
                                        UPDATE [TKWAREHOUSE].[dbo].[TWPOSTS]
                                        SET [WEIGHETS]={1},[PAYMONEYS]={2},[ISSINGALS]='{3}'
                                        WHERE [ID]='{0}'

                                       ", DR["ID"].ToString(), DR["重量"].ToString(), DR["金額"].ToString(), DR["單筆單件"].ToString());

                    //ID = ID + "," +DR["ID"].ToString();
                }
            }

            UPDATETWPOSTS(sbSql.ToString());
            //MessageBox.Show(ID);
        }

        public void UPDATETWPOSTS(string SQLCOMMAND)
        {
            try
            {
                //20210902密
                Class1 TKID = new Class1();//用new 建立類別實體
                SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

                //資料庫使用者密碼解密
                sqlsb.Password = TKID.Decryption(sqlsb.Password);
                sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

                String connectionString;
                sqlConn = new SqlConnection(sqlsb.ConnectionString);

                sqlConn.Close();
                sqlConn.Open();
                tran = sqlConn.BeginTransaction();

                sbSql.Clear();

                sbSql.AppendFormat(@" 
                                    {0}
                                    ", SQLCOMMAND.ToString());


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

        public int SEARCHTWPOSTSBASE(decimal WEIGHTS)
        {
            SqlDataAdapter adapter = new SqlDataAdapter();
            SqlCommandBuilder sqlCmdBuilder = new SqlCommandBuilder();
            DataSet ds1 = new DataSet();


            try
            {
                //20210902密
                Class1 TKID = new Class1();//用new 建立類別實體
                SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

                //資料庫使用者密碼解密
                sqlsb.Password = TKID.Decryption(sqlsb.Password);
                sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

                String connectionString;
                sqlConn = new SqlConnection(sqlsb.ConnectionString);


                sbSql.Clear();
                sbSqlQuery.Clear();

                sbSql.AppendFormat(@"  
                                    SELECT 
                                    [SWEIGHTS]
                                    ,[EWEIGHTS]
                                    ,[PRICES]
                                    FROM [TKWAREHOUSE].[dbo].[TWPOSTSBASE]
                                    WHERE [SWEIGHTS]<={0} AND [EWEIGHTS]>={0}
                                    
                                        ", WEIGHTS);

                adapter = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder = new SqlCommandBuilder(adapter);
                sqlConn.Open();
                ds1.Clear();
                adapter.Fill(ds1, "ds1");
                sqlConn.Close();


                if (ds1.Tables["ds1"].Rows.Count >= 1)
                {
                    return Convert.ToInt32(ds1.Tables["ds1"].Rows[0]["PRICES"].ToString());

                }
                else
                {
                    return 0;
                }

            }
            catch
            {
                return 0;
            }
            finally
            {
                sqlConn.Close();
            }
        }

        public void CHECKADDDATA()
        {
            //IEnumerable<DataRow> tempExcept = null;

            DataTable DT1 = SEARCHTWPOSTS();
            DataTable DT2 = IMPORTEXCEL();
         
            //找DataTable差集
            //要有相同的欄位名稱
            //找DataTable差集
            //如果兩個datatable中有部分欄位相同，可以使用Contains比較　
        
              var  tempExcept = from r in DT2.AsEnumerable()
                                 where
                                 !(from rr in DT1.AsEnumerable() select rr.Field<string>("託運單編號")).Contains(
                                 r.Field<string>("託運單編號"))
                                 select r;
          


            //var tempExcept = DT2.AsEnumerable();

            if (tempExcept.Count() > 0)
            {
                //差集集合
                DataTable dt3 = tempExcept.CopyToDataTable();

                INSERTINTOTWPOSTS(dt3);
            }
            else
            {
                MessageBox.Show("沒有新資料可匯入");
            }
        }

        public DataTable SEARCHTWPOSTS()
        {
            SqlDataAdapter adapter1 = new SqlDataAdapter();
            SqlCommandBuilder sqlCmdBuilder1 = new SqlCommandBuilder();
            DataSet ds1 = new DataSet();

            //THISYEARS = "21";

            try
            {
                //connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                //sqlConn = new SqlConnection(connectionString);

                //20210902密
                Class1 TKID = new Class1();//用new 建立類別實體
                SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

                //資料庫使用者密碼解密
                sqlsb.Password = TKID.Decryption(sqlsb.Password);
                sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

                String connectionString;
                sqlConn = new SqlConnection(sqlsb.ConnectionString);

                sbSql.Clear();
                sbSqlQuery.Clear();



                //核準過TASK_RESULT='0'
                //AND DOC_NBR  LIKE 'QC1002{0}%'

                sbSql.AppendFormat(@"  
                                   SELECT 
                                    [SENDNO] AS '託運單編號'
                                    FROM [TKWAREHOUSE].[dbo].[TWPOSTS]
                                    UNION ALL
                                    SELECT '託運單編號'  AS '託運單編號'


                                    ");


                adapter1 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder1 = new SqlCommandBuilder(adapter1);
                sqlConn.Open();
                ds1.Clear();
                adapter1.Fill(ds1, "ds1");
                sqlConn.Close();

                if (ds1.Tables["ds1"].Rows.Count >= 1)
                {
                    return ds1.Tables["ds1"];

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
                sqlConn.Close();
            }
        }

        public DataTable IMPORTEXCEL()
        {
            //記錄選到的檔案路徑
            _path = null;

            OpenFileDialog od = new OpenFileDialog();
            od.Filter = "Excell|*.xls;*.xlsx;";

            DialogResult dr = od.ShowDialog();
            if (dr == DialogResult.Abort)
            {
                return null;
            }
            if (dr == DialogResult.Cancel)
            {
                return null;
            }
            
           
            _path = od.FileName.ToString();

            try
            {
                //  ExcelConn(_path);
                //找出不同excel的格式，設定連接字串
                //xls跟非xls
                string constr = null;
                string CHECKEXCELFORMAT = _path.Substring(_path.Length - 4, 4);

                if (CHECKEXCELFORMAT.CompareTo("xlsx") == 0)
                {
                    constr = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + _path + ";Extended Properties='Excel 12.0;HDR=NO';"; //for above excel 2007  
                }
                else
                {
                    
                    constr = @"provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + _path + ";Extended Properties='Excel 8.0;HRD=Yes;IMEX=1';"; //for below excel 2007  
                }

                //找出excel的第1張分頁名稱，用query中                
                OleDbConnection Econ = new OleDbConnection(constr);
                Econ.Open();



                DataTable excelShema = Econ.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
                string firstSheetName = excelShema.Rows[0]["TABLE_NAME"].ToString();

                string Query = string.Format("Select * FROM [{0}]", firstSheetName);
                OleDbCommand Ecom = new OleDbCommand(Query, Econ);


                DataTable dtExcelData = new DataTable();

                OleDbDataAdapter oda = new OleDbDataAdapter(Query, Econ);
                Econ.Close();
                oda.Fill(dtExcelData);
                DataTable Exceldt = dtExcelData;

                //如果xlsx要另外處理欄位名
                if (CHECKEXCELFORMAT.CompareTo("xlsx") == 0)
                {
                    //把第一列的欄位名移除，並重設欄位名
                    //Exceldt.Rows[0].Delete();
                    Exceldt.Columns[0].ColumnName = "交寄日期";
                    Exceldt.Columns[1].ColumnName = "客戶編號";
                    Exceldt.Columns[2].ColumnName = "客戶名稱";
                    Exceldt.Columns[3].ColumnName = "電話";
                    Exceldt.Columns[4].ColumnName = "郵遞區號";
                    Exceldt.Columns[5].ColumnName = "地址";
                    Exceldt.Columns[6].ColumnName = "內裝物品Memo";
                    Exceldt.Columns[7].ColumnName = "件數編號";
                    Exceldt.Columns[8].ColumnName = "備註(出貨單編號)";
                    Exceldt.Columns[9].ColumnName = "使用單位編號";
                    Exceldt.Columns[10].ColumnName = "託運單編號";
                    Exceldt.Columns[11].ColumnName = "手機";
                    //Exceldt.Columns[12].ColumnName = "代收貨價";
                    //Exceldt.Columns[13].ColumnName = "重量";
                    //Exceldt.Columns[14].ColumnName = "金額";
                    //Exceldt.Columns[15].ColumnName = "單筆單件";
                }



                if (Exceldt.Rows.Count > 0)
                {
                    return Exceldt;
                }
                else
                {
                    return null;
                }


            }
            catch (Exception ex)
            {
                return null;
                //MessageBox.Show(string.Format("錯誤:{0}", ex.Message), "Not Imported", MessageBoxButtons.OK, MessageBoxIcon.Warning);

            }
        }

        public void INSERTINTOTWPOSTS(DataTable DT)
        {
            try
            {
                //20210902密
                Class1 TKID = new Class1();//用new 建立類別實體
                SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

                //資料庫使用者密碼解密
                sqlsb.Password = TKID.Decryption(sqlsb.Password);
                sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

                String connectionString;
                sqlConn = new SqlConnection(sqlsb.ConnectionString);


                sqlConn.Close();
                sqlConn.Open();
                tran = sqlConn.BeginTransaction();

                sbSql.Clear();

                foreach (DataRow DR in DT.Rows)
                {
                    sbSql.AppendFormat(@" 
                                        
                                        INSERT INTO [TKWAREHOUSE].[dbo].[TWPOSTS]
                                        (
                                        [SENDDATES]
                                        ,[CUSTOMERNO]
                                        ,[CUSTOMERNAMES]
                                        ,[PHONES]
                                        ,[ZIPCODE]
                                        ,[ADDRESS]
                                        ,[SENDCONTENTS]
                                        ,[SENDNUMS]
                                        ,[COMMENTS]
                                        ,[USEDUNITS]
                                        ,[SENDNO]
                                        ,[MOBILEPHONE]
                                        ,[COLMONEYS]
                                        ,[WEIGHETS]
                                        ,[PAYMONEYS]
                                        ,[ISSINGALS]
                                        )
                                        VALUES
                                        (
                                        '{0}'
                                        ,'{1}'
                                        ,'{2}'
                                        ,'{3}'
                                        ,'{4}'
                                        ,'{5}'
                                        ,'{6}'
                                        ,'{7}'
                                        ,'{8}'
                                        ,'{9}'
                                        ,'{10}'
                                        ,'{11}'
                                        ,'{12}'
                                        ,'{13}'
                                        ,'{14}'
                                        ,'{15}'
                                        )

                                        
                                           
                                         ", Convert.ToDateTime(DR["交寄日期"].ToString().Replace("'", "").Replace("下午", "PM").Replace("上午", "AM")).ToString("yyyy/MM/dd HH:mm")
                                            , DR["客戶編號"].ToString().Replace("'", "").Replace(" ", "")
                                            , DR["客戶名稱"].ToString().Replace("'", "").Replace(" ", "")
                                            , DR["電話"].ToString().Replace("'", "").Replace(" ", "")
                                            , DR["郵遞區號"].ToString().Replace("'", "").Replace(" ", "")
                                            , DR["地址"].ToString().Replace("'", "").Replace(" ", "")
                                            , DR["內裝物品Memo"].ToString().Replace("'", "").Replace(" ", "")
                                            , DR["件數編號"].ToString().Replace("'", "").Replace(" ", "")
                                            , DR["備註(出貨單編號)"].ToString().Replace("'", "").Replace(" ", "")
                                            , DR["使用單位編號"].ToString().Replace("'", "").Replace(" ", "")
                                            , DR["託運單編號"].ToString().Replace("'", "").Replace(" ", "")
                                            , DR["手機"].ToString().Replace("'", "").Replace(" ", "")
                                            , 0
                                            , 0
                                            , 0
                                            , 'N'



                                        );

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

                    MessageBox.Show("完成");
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

        #endregion

        #region BUTTON

        private void button1_Click(object sender, EventArgs e)
        {
            SEARCH(dateTimePicker1.Value.ToString("yyyyMMdd"), dateTimePicker2.Value.ToString("yyyyMMdd"));

        }
        private void button2_Click(object sender, EventArgs e)
        {
            SAVE(MAINds);

            SEARCH(dateTimePicker1.Value.ToString("yyyyMMdd"), dateTimePicker2.Value.ToString("yyyyMMdd"));

        }


        #endregion

        private void button4_Click(object sender, EventArgs e)
        {
            CHECKADDDATA();

            //MessageBox.Show("完成");
        }
    }
}

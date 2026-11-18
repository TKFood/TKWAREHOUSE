using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using FastReport;
using FastReport.Data;
using FastReport.Preview;
using TKITDLL;


namespace TKWAREHOUSE
{
    public partial class FrmINVCheck : Form
    {
        // 報表 SQL 執行逾時秒數
        private const int REPORT_TIMEOUT = 180;
        // 批號錯誤訊息最多列出的筆數
        private const int LA016_ERROR_SHOW_MAX = 30;

        public Report report1 { get; private set; }
        public Report report2 { get; private set; }
        public Report report3 { get; private set; }

        public FrmINVCheck()
        {
            InitializeComponent();

            comboboxload();
            comboboxload2();
        }

        #region 共用

        /// <summary>
        /// 取得解密後的連線字串 (dbconn / dberp)
        /// </summary>
        private string GetConnectionString(string name)
        {
            Class1 TKID = new Class1();
            SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings[name].ConnectionString);
            sqlsb.Password = TKID.Decryption(sqlsb.Password);
            sqlsb.UserID = TKID.Decryption(sqlsb.UserID);
            return sqlsb.ConnectionString;
        }

        /// <summary>
        /// 執行查詢並回傳 DataTable (連線用完即關閉)
        /// </summary>
        private DataTable QueryTable(string sql, params SqlParameter[] parameters)
        {
            DataTable dt = new DataTable();

            using (SqlConnection conn = new SqlConnection(GetConnectionString("dbconn")))
            using (SqlCommand cmd = new SqlCommand(sql, conn))
            using (SqlDataAdapter da = new SqlDataAdapter(cmd))
            {
                cmd.CommandTimeout = REPORT_TIMEOUT;
                if (parameters != null && parameters.Length > 0)
                {
                    cmd.Parameters.AddRange(parameters);
                }
                da.Fill(dt);
            }

            return dt;
        }

        /// <summary>
        /// 載入 frx、設定連線與資料來源 SQL 後顯示在指定的預覽元件。
        /// 失敗時顯示錯誤訊息並保留原本的報表。
        /// </summary>
        private Report ShowReport(Report oldReport, string frxPath, string connName, Dictionary<string, string> dataSources, PreviewControl preview)
        {
            Report report = new Report();

            try
            {
                report.Load(frxPath);

                report.Dictionary.Connections[0].ConnectionString = GetConnectionString(connName);
                report.Dictionary.Connections[0].CommandTimeout = REPORT_TIMEOUT;

                foreach (KeyValuePair<string, string> ds in dataSources)
                {
                    TableDataSource table = report.GetDataSource(ds.Key) as TableDataSource;
                    if (table != null)
                    {
                        table.SelectCommand = ds.Value;
                    }
                }

                report.Preview = preview;
                report.Show();
            }
            catch (Exception ex)
            {
                report.Dispose();
                MessageBox.Show("產生報表失敗：" + ex.Message, "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return oldReport;
            }

            return report;
        }

        /// <summary>
        /// 載入庫別下拉選單
        /// </summary>
        private void LoadWarehouseCombo(ComboBox comboBox, string dbName, string defaultValue)
        {
            string sql = string.Format("SELECT MC001,MC001+MC002 AS MC002 FROM [{0}].dbo.CMSMC WITH (NOLOCK) ORDER BY MC001", dbName);

            DataTable dt = new DataTable();
            dt.Columns.Add("MC001", typeof(string));
            dt.Columns.Add("MC002", typeof(string));

            using (SqlConnection conn = new SqlConnection(GetConnectionString("dbconn")))
            using (SqlDataAdapter da = new SqlDataAdapter(sql, conn))
            {
                da.Fill(dt);
            }

            comboBox.DataSource = dt.DefaultView;
            comboBox.ValueMember = "MC001";
            comboBox.DisplayMember = "MC002";
            comboBox.SelectedValue = defaultValue;
        }

        /// <summary>
        /// 「只看當日有異動的品號」篩選條件 (勾選 checkBox 時才有)
        /// </summary>
        private string BuildDateFilter(string dbName, CheckBox checkBox, DateTimePicker picker, string LA009)
        {
            if (!checkBox.Checked)
            {
                return " ";
            }

            string day = picker.Value.ToString("yyyyMMdd");
            return string.Format(@"
                                    AND LA001 IN (SELECT LA001 FROM [{0}].dbo.INVLA WITH (NOLOCK) WHERE LA004='{1}' AND LA009='{2}')
                                    AND LA004<='{1}'", dbName, day, LA009);
        }

        #endregion

        #region FUNCTION

        public void comboboxload()
        {
            LoadWarehouseCombo(comboBox1, "TK", "20001");
        }

        public void comboboxload2()
        {
            LoadWarehouseCombo(comboBox2, "DY", "10001     ");
        }

        public void SETFASTREPORT(string KINDS)
        {
            // 先檢查批號 LA016 前 8 碼是否為 yyyyMMdd 日期，有錯先告知使用者
            if (!CHECK_LA016(KINDS))
            {
                return;
            }

            Dictionary<string, string> dataSources = new Dictionary<string, string>();
            dataSources.Add("Table", SETFASETSQL(KINDS));

            report1 = ShowReport(report1, GetReportFile(KINDS), "dbconn", dataSources, previewControl1);
        }

        /// <summary>
        /// 依庫別取得報表檔
        /// </summary>
        private string GetReportFile(string KINDS)
        {
            switch (KINDS)
            {
                case "20001":
                case "20028":
                    return @"REPORT\每日盤點表-成品.frx";
                case "21001":
                    return @"REPORT\每日盤點表-成品-21001V2.frx";
                case "20006":
                    return @"REPORT\每日盤點表-原料.frx";
                case "20004":
                    return @"REPORT\每日盤點表-物料.frx";
                case "20005":
                    return @"REPORT\每日盤點表-半成品.frx";
                default:
                    return @"REPORT\每日盤點表.frx";
            }
        }

        /// <summary>
        /// 檢查指定庫別有庫存的批號(LA016)，前 8 碼是否為 yyyyMMdd 有效日期
        /// (第 9 碼之後不檢查，例：20261118、20261118A、20261118- 都算正常；空白批號也不警告)。
        /// 有錯誤時列出錯誤的品號/批號給使用者，並詢問是否仍要產生報表。
        /// 回傳 true = 繼續產生報表；false = 中止。
        /// </summary>
        public bool CHECK_LA016(string KINDS)
        {
            StringBuilder sbSql = new StringBuilder();
            sbSql.Append(@"
                            SELECT LA001 AS '品號', MB002 AS '品名', LA016 AS '批號', CAST(SUM(LA005*LA011) AS DECIMAL(18,3)) AS '庫存量'
                            FROM [TK].dbo.INVLA WITH (NOLOCK)
                            LEFT JOIN [TK].dbo.INVMB WITH (NOLOCK) ON MB001=LA001
                            WHERE LA009=@LA009
                            -- 空白批號不檢查
                            AND LTRIM(RTRIM(ISNULL(LA016,'')))<>''
                            -- 合格批號：前 8 碼為 yyyyMMdd 有效日期
                            AND NOT (
                                    LEN(LTRIM(ISNULL(LA016,'')))>=8
                                    AND LEFT(LTRIM(LA016),8) NOT LIKE '%[^0-9]%'
                                    AND ISDATE(LEFT(LTRIM(LA016),8))=1
                                    )
                            ");

            // 與各報表 SQL 的篩選條件一致，避免檢查到報表本來就不會出現的資料
            if (KINDS.Equals("20001"))
            {
                sbSql.Append(@" AND (LA001 LIKE '4%' OR LA001 LIKE '5%') ");
            }
            else if (KINDS.Equals("21001"))
            {
                sbSql.Append(@" AND (LA001 LIKE '4%' OR LA001 LIKE '5%') AND LA016 NOT IN ('LotError') ");
            }

            sbSql.Append(@"
                            GROUP BY LA001, MB002, LA016
                            HAVING SUM(LA005*LA011)<>0
                            ORDER BY LA001, LA016
                            ");

            DataTable dtErr;
            try
            {
                dtErr = QueryTable(sbSql.ToString(), new SqlParameter("@LA009", KINDS));
            }
            catch (Exception ex)
            {
                MessageBox.Show("檢查批號(LA016)格式時發生錯誤：" + ex.Message, "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }

            if (dtErr.Rows.Count == 0)
            {
                return true;
            }

            StringBuilder sbMsg = new StringBuilder();
            sbMsg.AppendLine("庫別 " + KINDS + " 有 " + dtErr.Rows.Count + " 筆批號(LA016)前 8 碼不是 yyyyMMdd 日期：");
            sbMsg.AppendLine();

            for (int i = 0; i < dtErr.Rows.Count && i < LA016_ERROR_SHOW_MAX; i++)
            {
                DataRow dr = dtErr.Rows[i];
                sbMsg.AppendLine("品號: " + dr["品號"].ToString().Trim()
                                 + " | 品名: " + dr["品名"].ToString().Trim()
                                 + " | 批號: [" + dr["批號"].ToString().Trim() + "]"
                                 + " | 庫存量: " + dr["庫存量"].ToString());
            }
            if (dtErr.Rows.Count > LA016_ERROR_SHOW_MAX)
            {
                sbMsg.AppendLine("...(另有 " + (dtErr.Rows.Count - LA016_ERROR_SHOW_MAX) + " 筆未列出)");
            }

            sbMsg.AppendLine();
            sbMsg.AppendLine("請先修正以上批號。是否仍要產生報表？");

            DialogResult result = MessageBox.Show(sbMsg.ToString(), "批號格式錯誤", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
            return (result == DialogResult.Yes);
        }

        public string SETFASETSQL(string KINDS)
        {
            string today = DateTime.Now.ToString("yyyyMMdd");
            string dateFilter = BuildDateFilter("TK", checkBox1, dateTimePicker1, comboBox1.SelectedValue.ToString());

            switch (KINDS)
            {
                case "20001":
                    // 成品倉：只看 4、5 開頭的品號
                    return SQL_FINISHED_GOODS("20001", " AND (LA.LA001 LIKE '4%' OR LA.LA001 LIKE '5%') ");
                case "20028":
                    return SQL_FINISHED_GOODS("20028", " ");
                case "21001":
                    return string.Format(@"                                         
                                    SELECT 
                                    品號,品名,規格,批號,庫存量,單位
                                    ,CASE WHEN ISNULL(生產日期,'')<>'' THEN 生產日期 ELSE CONVERT(nvarchar,DATEADD(day,-1*(CASE WHEN MB198='1' THEN 1*MB023 WHEN  MB198='2' THEN 30*MB023  ELSE 0 END),CONVERT(datetime,LEFT(批號,8))),112) END AS '生產日期'
                                    ,CASE WHEN ISNULL(在倉日期,'')<>'' THEN 在倉日期 ELSE DATEDIFF(DAY,CASE WHEN ISNULL(生產日期,'')<>'' THEN 生產日期 ELSE CONVERT(nvarchar,DATEADD(day,-1*(CASE WHEN MB198='1' THEN 1*MB023 WHEN  MB198='2' THEN 30*MB023  ELSE 0 END),CONVERT(datetime,LEFT(批號,8))),112) END,'{0}') END AS '在倉日期'
                                    ,有效天數
                                    ,狀態
                                    ,CASE WHEN MB198='1' THEN 1*MB023 WHEN  MB198='2' THEN 30*MB023  ELSE 0 END AS 'DAYS'                 
                                    FROM (
                                    SELECT 
                                    品號,品名,規格,批號,庫存量,單位
                                    ,生產日期
                                    ,DATEDIFF(DAY,生產日期,'{0}') AS '在倉日期'
                                    ,DATEDIFF(DAY,'{0}',有效日期)  AS '有效天數'
                                    ,(CASE WHEN DATEDIFF(DAY,生產日期,'{0}')>90 THEN '在倉超過90天' ELSE (CASE WHEN DATEDIFF(DAY,生產日期,'{0}')>30 THEN '在倉超過30天' ELSE '' END) END ) AS '狀態'
                                    FROM ( 

                                    SELECT  
                                    LA001 AS '品號' ,MB002 AS '品名',MB003 AS '規格',LA016 AS '批號'
                                    ,CAST(SUM(LA005*LA011) AS INT) AS '庫存量',MB004 AS '單位'

                                    ,(SELECT TOP 1 ME032
                                    FROM [TK].dbo.INVME
                                    WHERE ME001=LA001 AND ME002=LA016) AS '生產日期'
                                    ,(SELECT TOP 1 ME009
                                    FROM [TK].dbo.INVME
                                    WHERE ME001=LA001 AND ME002=LA016) AS '有效日期'
                                    ,ISDATE(LA016) AS LA016
                                    FROM [TK].dbo.INVLA WITH (NOLOCK)  
                                    LEFT JOIN  [TK].dbo.INVMB WITH (NOLOCK) ON MB001=LA001   
                                    WHERE  (LA009='21001')   
                                    AND (LA001 LIKE '4%' OR LA001 LIKE '5%')

                                    AND LA016  NOT IN ('LotError')

                                    GROUP BY  LA001,LA009,MB002,MB003,LA016,MB023,MB198,MB004    
                                    HAVING SUM(LA005*LA011)<>0 
                                    ) AS TEMP
                                    ) AS TEMP2
                                    LEFT JOIN [TK].dbo.INVMB ON MB001=品號
                                    WHERE 品號 IN (SELECT LA001 FROM [TK].dbo.INVLA WHERE LA009 IN ('21001') GROUP BY LA001 HAVING SUM(LA005*LA011)>0 )
                                    AND 品名 NOT LIKE '%試吃%'
                                    ORDER BY 品號,批號
                                    ", today);
                case "20006":
                    return string.Format(@"  
                                     SELECT 
                                     品號,品名,規格,批號,庫存量,單位,庫存金額,在倉日期,有效天數,業務
                                     
                                     FROM (
                                     SELECT   LA001 AS '品號' ,MB002 AS '品名',MB003 AS '規格',LA016 AS '批號'
                                     ,CONVERT(DECIMAL(16,3),SUM(LA005*LA011)) AS '庫存量',MB004 AS '單位'
                                     ,CONVERT(DECIMAL(16,3),SUM(LA005*LA013)) AS '庫存金額'
                                     ,DATEDIFF(DAY,(SELECT TOP 1 TH014 FROM [TK].dbo.PURTG,[TK].dbo.PURTH WHERE TG001=TH001 AND TG002=TH002 AND TH004=LA001 AND TH010=LA016 ),'{0}') AS '在倉日期'  
                                     ,CASE WHEN ISDATE(LEFT(LA016,8))=1 THEN DATEDIFF(DAY,'{0}',LEFT(LA016,8)) ELSE 0 END  AS '有效天數'
                                     ,(SELECT TOP 1 TC006+' '+MV002 FROM [TK].dbo.COPTC,[TK].dbo.CMSMV WHERE TC006=MV001 AND  TC001+TC002 IN (SELECT TOP 1 TA026+TA027 FROM [TK].dbo.MOCTA WHERE TA001+TA002 IN (SELECT TOP 1 TG014+TG015 FROM [TK].dbo.MOCTG WHERE TG004=LA001 AND TG017=LA016))) AS '業務'

                                     FROM [TK].dbo.INVLA WITH (NOLOCK)  
                                     LEFT JOIN  [TK].dbo.INVMB WITH (NOLOCK) ON MB001=LA001   
                                     WHERE  (LA009='20006')   
                                      {1}
                                     GROUP BY  LA001,LA009,MB002,MB003,LA016,MB023,MB198,MB004   
                                     HAVING SUM(LA005*LA011)<>0 
                                     ) AS TEMP
                                     WHERE 品號 NOT IN ('122221001','114141009')
                                     ORDER BY 品號,批號
                                    ", dateTimePicker1.Value.ToString("yyyyMMdd"), dateFilter);
                case "20004":
                    return string.Format(@"                                     
                                        SELECT 
                                        品號,品名,規格,批號,庫存量,庫存金額
                                        ,(SELECT TOP 1 ME032
                                        FROM [TK].dbo.INVME
                                        WHERE ME001=品號 AND ME002=批號) AS '生產日期'
                                        ,(SELECT TOP 1 ME009
                                        FROM [TK].dbo.INVME
                                        WHERE ME001=品號 AND ME002=批號) AS '有效日期'
                                        FROM ( 
                                        SELECT  LA001 AS '品號' ,MB002 AS '品名',MB003 AS '規格',LA016 AS '批號'  ,CAST(SUM(LA005*LA011) AS DECIMAL(18,4)) AS '庫存量'  ,CAST(SUM(LA005*LA013) AS DECIMAL(18,4)) AS '庫存金額'  
                                        FROM [TK].dbo.INVLA WITH (NOLOCK) 
                                        LEFT JOIN [TK].dbo.INVMB WITH (NOLOCK) ON MB001=LA001 
                                        WHERE  (LA009='20004') 
                                        {1}

                                        GROUP BY  LA001,MB002,MB003,LA016
                                        HAVING SUM(LA005*LA011)<>0
                                        ) AS TEMP 
                                        ORDER BY  品號,品名,規格,批號
                                    ", comboBox1.SelectedValue.ToString(), dateFilter);
                case "20005":
                    return string.Format(@"   
                                    SELECT 
                                    品號,品名,規格,批號,庫存量,單位,生產日期
                                    , DATEDIFF(DAY,生產日期,'{0}') AS '在倉日期'
                                    ,(CASE WHEN MB198='2' THEN DATEDIFF(DAY,'{0}',DATEADD(month, MB023, '{0}' )) END)-(CASE WHEN DATEDIFF(DAY,生產日期,'{0}')>=0 THEN DATEDIFF(DAY,生產日期,'{0}') ELSE (CASE WHEN DATEDIFF(DAY,生產日期,'{0}')<0 THEN  (CASE WHEN MB198='2' THEN DATEDIFF(DAY,DATEADD(month, -1*MB023, 生產日期 ),'{0}') END ) END ) END)  AS '有效天數' 
                                    ,業務
                                    ,ISNULL((庫存量*(SELECT MB065/NULLIF(MB064,0) FROM [TK].dbo.INVMB WHERE MB001=品號)),0) AS 庫存金額
                                    ,(CASE WHEN DATEDIFF(DAY,生產日期,'{0}')>90 THEN '在倉超過90天' ELSE (CASE WHEN DATEDIFF(DAY,生產日期,'{0}')>30 THEN '在倉超過30天' ELSE '' END) END ) AS '狀態'
                                    ,有效日期

                                    FROM (
                                    SELECT   
                                    LA001 AS '品號' ,INVMB.MB002 AS '品名',INVMB.MB003 AS '規格',LA016 AS '批號'
                                    ,CONVERT(DECIMAL(16,3),SUM(LA005*LA011)) AS '庫存量',INVMB.MB004 AS '單位',MB198,MB023

                                    ,(SELECT TOP 1 TC006+' '+MV002 FROM [TK].dbo.COPTC,[TK].dbo.CMSMV WHERE TC006=MV001 AND  TC001+TC002 IN (SELECT TOP 1 TA026+TA027 FROM [TK].dbo.MOCTA WHERE TA001+TA002 IN (SELECT TOP 1 TG014+TG015 FROM [TK].dbo.MOCTG WHERE TG004=LA001 AND TG017=LA016))) AS '業務'
                                    ,(SELECT TOP 1 ME032
                                    FROM [TK].dbo.INVME
                                    WHERE ME001=LA001 AND ME002=LA016) AS '生產日期'
                                    ,(SELECT TOP 1 ME009
                                    FROM [TK].dbo.INVME
                                    WHERE ME001=LA001 AND ME002=LA016) AS '有效日期'
                                    ,ISDATE(LA016) AS LA016
                                    FROM [TK].dbo.INVLA WITH (NOLOCK) 
                                    LEFT JOIN  [TK].dbo.INVMB WITH (NOLOCK) ON MB001=LA001  
                                    WHERE  (LA009='20005') 

                                    GROUP BY  LA001,LA009,MB002,MB003,LA016,MB023,MB198,MB004
                                    HAVING SUM(LA005*LA011)<>0 
                                    ) AS TEMP
                                    ORDER BY 品號,批號
                                    ", dateTimePicker1.Value.ToString("yyyyMMdd"));
                default:
                    return string.Format(@"  
                                        SELECT 
                                        品號,品名,規格,批號,庫存量,庫存金額
                                        ,(SELECT TOP 1 ME032
                                        FROM [TK].dbo.INVME
                                        WHERE ME001=品號 AND ME002=批號) AS '生產日期'
                                        ,(SELECT TOP 1 ME009
                                        FROM [TK].dbo.INVME
                                        WHERE ME001=品號 AND ME002=批號) AS '有效日期'
                                        FROM ( 
                                        SELECT  LA001 AS '品號' ,MB002 AS '品名',MB003 AS '規格',LA016 AS '批號'  ,CAST(SUM(LA005*LA011) AS DECIMAL(18,4)) AS '庫存量'  ,CAST(SUM(LA005*LA013) AS DECIMAL(18,4)) AS '庫存金額'  
                                        FROM [TK].dbo.INVLA WITH (NOLOCK) 
                                        LEFT JOIN [TK].dbo.INVMB WITH (NOLOCK) ON MB001=LA001 
                                        WHERE  (LA009='{0}') 
                                        {1}

                                        GROUP BY  LA001,MB002,MB003,LA016
                                        HAVING SUM(LA005*LA011)<>0
                                        ) AS TEMP 

                                        ORDER BY  品號,品名,規格,批號
                                    ", comboBox1.SelectedValue.ToString(), dateFilter);
            }
        }

        /// <summary>
        /// 成品類每日盤點表 SQL (20001、20028 共用，報表檔都是 每日盤點表-成品.frx)
        /// </summary>
        /// <param name="LA009">庫別</param>
        /// <param name="extraFilter">額外的 INVLA 篩選條件 (欄位請用 LA. 前綴)</param>
        private string SQL_FINISHED_GOODS(string LA009, string extraFilter)
        {
            return string.Format(@"  
                                        -- 1. 設定基準日變數與基礎轉型，釋放運算效能
                                            DECLARE @BaseDate DATETIME = '{0}';
                                            DECLARE @BaseDateStr NVARCHAR(8) = '{1}';

                                            WITH 
                                            -- 2. 提前彙整訂單總需求量
                                            Order_Total_Demand AS (
                                                SELECT TD004 AS ItemNo, ISNULL(SUM(NUM), 0) AS TotalDemand
                                                FROM [TK].dbo.VCOPTDINVMD WITH(NOLOCK)
                                                WHERE TD013 >= @BaseDateStr
                                                GROUP BY TD004
                                            ),
                                            -- 3. 基礎庫存資料與基礎欄位加總 (消除內層的子查詢)
                                            Base_Inventory AS (
                                                SELECT  
                                                    LA.LA001 AS [品號],
                                                    MB.MB002 AS [品名],
                                                    MB.MB003 AS [規格],
                                                    LA.LA016 AS [批號],
                                                    CAST(SUM(LA.LA005 * LA.LA011) AS INT) AS [庫存量],
                                                    MB.MB004 AS [單位],
                                                    MB.MB023,
                                                    MB.MB198,
                                                    LA.LA009 AS [庫別],
                                                    ROUND(MB.MB023 / 3.0, 0) AS [MonthOffset],
                                                    -- 批號前 8 碼轉日期；不是有效 yyyyMMdd 時為 NULL，避免整張報表轉換失敗
                                                    CASE WHEN LEFT(LA.LA016,8) NOT LIKE '%[^0-9]%' AND LEN(LA.LA016)>=8 AND ISDATE(LEFT(LA.LA016,8))=1
                                                         THEN CONVERT(DATETIME, LEFT(LA.LA016,8)) END AS [BatchDate],
                                                    -- 提前帶出製令/業務鏈需要的末端製令單號
                                                    (SELECT TOP 1 TG014+TG015 FROM [TK].dbo.MOCTG WITH(NOLOCK) WHERE TG004 = LA.LA001 AND TG017 = LA.LA016) AS MoctggKey
                                                FROM [TK].dbo.INVLA AS LA WITH (NOLOCK)  
                                                LEFT JOIN [TK].dbo.INVMB AS MB WITH (NOLOCK) ON MB.MB001 = LA.LA001   
                                                WHERE LA.LA009 = '{2}'   
                                                {3}
                                                GROUP BY LA.LA001, LA.LA009, MB.MB002, MB.MB003, LA.LA016, MB.MB023, MB.MB198, MB.MB004  
                                                HAVING SUM(LA.LA005 * LA.LA011) <> 0 
                                            ),
                                            -- 4. 串接生產/有效日期與業務名稱 (收斂巢狀 Loop)
                                            Inventory_With_Dates AS (
                                                SELECT 
                                                    I.*,
                                                    ISNULL(ME.ME032, '') AS [原始生產日期],
                                                    ME.ME009 AS [原始有效日期],
                                                    -- 計算動態生產日期
                                                    CASE 
                                                        WHEN ISNULL(ME.ME032, '') <> '' THEN ME.ME032 
                                                        ELSE CONVERT(NVARCHAR(8), DATEADD(DAY, -1 * (CASE WHEN I.MB198 = '1' THEN 1 * I.MB023 WHEN I.MB198 = '2' THEN 30 * I.MB023 ELSE 0 END), I.[BatchDate]), 112) 
                                                    END AS [修正生產日期],
                                                    -- 業務人員名稱
                                                    (SELECT TOP 1 C.TC006 + ' ' + V.MV002 
                                                     FROM [TK].dbo.MOCTA AS A WITH(NOLOCK)
                                                     INNER JOIN [TK].dbo.COPTC AS C WITH(NOLOCK) ON C.TC001 = A.TA026 AND C.TC002 = A.TA027
                                                     INNER JOIN [TK].dbo.CMSMV AS V WITH(NOLOCK) ON C.TC006 = V.MV001
                                                     WHERE A.TA001 + A.TA002 = I.MoctggKey) AS [業務]
                                                FROM Base_Inventory AS I
                                                LEFT JOIN [TK].dbo.INVME AS ME WITH(NOLOCK) ON ME.ME001 = I.[品號] AND ME.ME002 = I.[批號]
                                            )

                                            -----------------------------------------------------------------------------------
                                            -- 主查詢開始 (引入大表需求量與動態 FOR XML 單據串接)
                                            -----------------------------------------------------------------------------------
                                            SELECT 
                                                D.[品號], D.[品名], D.[規格], D.[批號], D.[庫存量], D.[單位],
                                                CONVERT(INT,V_Demand.EffDemand) AS [效期內的訂單需求量],
                                                CONVERT(INT,(D.[庫存量] - ISNULL(V_Demand.EffDemand, 0))) AS [效期內的訂單差異量],
                                                CONVERT(INT,ISNULL(T_Demand.TotalDemand, 0)) AS [總訂單需求量],
                                                D.[業務],
                                                D.[修正生產日期] AS [生產日期],
    
                                                -- 在倉日期計算
                                                CASE 
                                                    WHEN ISNULL(D.[原始生產日期], '') <> '' THEN DATEDIFF(DAY, D.[原始生產日期], @BaseDate) 
                                                    ELSE DATEDIFF(DAY, D.[修正生產日期], @BaseDate) 
                                                END AS [在倉日期],
    
                                                DATEDIFF(DAY, @BaseDate, D.[原始有效日期]) AS [有效天數],
    
                                                -- 狀態分流
                                                CASE 
                                                    WHEN DATEDIFF(DAY, CASE WHEN ISNULL(D.[原始生產日期], '') <> '' THEN D.[原始生產日期] ELSE D.[修正生產日期] END, @BaseDate) > 90 THEN '在倉超過90天' 
                                                    WHEN DATEDIFF(DAY, CASE WHEN ISNULL(D.[原始生產日期], '') <> '' THEN D.[原始生產日期] ELSE D.[修正生產日期] END, @BaseDate) > 30 THEN '在倉超過30天' 
                                                    ELSE '' 
                                                END AS [狀態],
    
                                                CASE WHEN D.MB198 = '1' THEN 1 * D.MB023 WHEN D.MB198 = '2' THEN 30 * D.MB023 ELSE 0 END AS 'DAYS',
                                                CONVERT(NVARCHAR(8), DATEADD(DAY, -1 * (CASE WHEN D.MB198 = '1' THEN 1 * D.MB023 WHEN D.MB198 = '2' THEN 30 * D.MB023 ELSE 0 END), D.[BatchDate]), 112) AS '外購品的生產日',
                                                D.[庫別],
                                                NOS.NOS AS [NOS]

                                            FROM Inventory_With_Dates AS D
                                            -- 關聯總需求量
                                            LEFT JOIN Order_Total_Demand AS T_Demand ON T_Demand.ItemNo = D.[品號]
                                            -- 關聯效期內需求量 (動態計算日期區間，減少對大表的掃描)
                                            CROSS APPLY (
                                                SELECT ISNULL(SUM(NUM), 0) AS EffDemand
                                                FROM [TK].dbo.VCOPTDINVMD WITH(NOLOCK)
                                                WHERE TD004 = D.[品號] 
                                                  AND TD013 >= @BaseDateStr 
                                                  AND TD013 <= CONVERT(NVARCHAR(8), DATEADD(MONTH, -1 * D.MonthOffset, D.[BatchDate]), 112)
                                            ) AS V_Demand

                                            -- 延遲加載：最後才為每一列庫存生成 10 表 UNION 的單據號碼 (大幅釋放系統 memory 與 I/O)
                                            CROSS APPLY (
                                                SELECT (
                                                    SELECT Key1 + '-' + Key2 + ': ' + CONVERT(NVARCHAR(50), Key3) + ' ' + M_MF002 + CHAR(10)
                                                    FROM ( 
                                                        SELECT DISTINCT '1' As Type, TG001 As Key1, TG002 As Key2, (TH008 + TH024) AS Key3, (CASE WHEN (COPTG.MODIFIER <> '') THEN B.MF002 ELSE A.MF002 END) As M_MF002 
                                                        FROM TK..COPTG AS COPTG WITH(NOLOCK)
                                                        INNER JOIN TK..COPTH AS COPTH WITH(NOLOCK) ON TH001 = TG001 AND TH002 = TG002
                                                        LEFT JOIN TK..ADMMF As A WITH(NOLOCK) On A.MF001 = COPTG.CREATOR
                                                        LEFT JOIN TK..ADMMF As B WITH(NOLOCK) On B.MF001 = COPTG.MODIFIER
                                                        WHERE TH004 = D.[品號] And TH017 = D.[批號] AND TH007 = '{2}' AND TG023 = 'N'
            
                                                        UNION ALL 
            
                                                        SELECT DISTINCT (CASE WHEN (MQ003 = '11') THEN '2' ELSE '3' END) As Type, TA001 As Key1, TA002 As Key2, (TB007) AS Key3, (CASE WHEN (INVTA.MODIFIER <> '') THEN B.MF002 ELSE A.MF002 END) As M_MF002 
                                                        FROM TK..INVTA AS INVTA WITH(NOLOCK)
                                                        INNER JOIN TK..INVTB AS INVTB WITH(NOLOCK) ON TB001 = TA001 AND TB002 = TA002
                                                        LEFT JOIN TK..CMSMQ AS CMSMQ WITH(NOLOCK) ON MQ001 = TA001
                                                        LEFT JOIN TK..ADMMF As A WITH(NOLOCK) On A.MF001 = INVTA.CREATOR
                                                        LEFT JOIN TK..ADMMF As B WITH(NOLOCK) On B.MF001 = INVTA.MODIFIER
                                                        WHERE TB004 = D.[品號] AND TB014 = D.[批號] AND TB012 = '{2}' AND TA006 = 'N' AND MQ010 = -1 
            
                                                        UNION ALL 
            
                                                        SELECT DISTINCT '4' As Type, TF001 As Key1, TF002 As Key2, (TG009) AS Key3, (CASE WHEN (INVTF.MODIFIER <> '') THEN B.MF002 ELSE A.MF002 END) As M_MF002 
                                                        FROM TK..INVTF AS INVTF WITH(NOLOCK)
                                                        INNER JOIN TK..INVTG AS INVTG WITH(NOLOCK) ON TG001 = TF001 AND TG002 = TF002
                                                        LEFT JOIN TK..ADMMF As A WITH(NOLOCK) On A.MF001 = INVTF.CREATOR
                                                        LEFT JOIN TK..ADMMF As B WITH(NOLOCK) On B.MF001 = INVTF.MODIFIER
                                                        WHERE TG004 = D.[品號] AND TG017 = D.[批號] AND TG007 = '{2}' AND TF020 = 'N' 
            
                                                        UNION ALL 
            
                                                        SELECT DISTINCT '5' As Type, TH001 As Key1, TH002 As Key2, (TI009) AS Key3, (CASE WHEN (INVTH.MODIFIER <> '') THEN B.MF002 ELSE A.MF002 END) As M_MF002 
                                                        FROM TK..INVTH AS INVTH WITH(NOLOCK)
                                                        INNER JOIN TK..INVTI AS INVTI WITH(NOLOCK) ON TI001 = TH001 AND TI002 = TH002
                                                        LEFT JOIN TK..ADMMF As A WITH(NOLOCK) On A.MF001 = INVTH.CREATOR
                                                        LEFT JOIN TK..ADMMF As B WITH(NOLOCK) On B.MF001 = INVTH.MODIFIER
                                                        WHERE TI004 = D.[品號] AND TI017 = D.[批號] AND TI007 = '{2}' AND TH020 = 'N' 
            
                                                        UNION ALL 
            
                                                        SELECT DISTINCT '6' As Type, TA001 As Key1, TA002 As Key2, (TB009 + TB011) AS Key3, (CASE WHEN (EPSTA.MODIFIER <> '') THEN B.MF002 ELSE A.MF002 END) As M_MF002 
                                                        FROM TK..EPSTA AS EPSTA WITH(NOLOCK)
                                                        INNER JOIN TK..EPSTB AS EPSTB WITH(NOLOCK) ON TB001 = TA001 AND TB002 = TA002
                                                        LEFT JOIN TK..ADMMF As A WITH(NOLOCK) On A.MF001 = EPSTA.CREATOR
                                                        LEFT JOIN TK..ADMMF As B WITH(NOLOCK) On B.MF001 = EPSTA.MODIFIER
                                                        WHERE TB007 = D.[品號] AND TB019 = D.[批號] AND TB018 = '{2}' AND TA034 <> 'V' AND TB021 + TB022 + TB023 = '' AND TB042 + TB043 + TB044 = '' 
            
                                                        UNION ALL 
            
                                                        SELECT DISTINCT '7' As Type, TD001 As Key1, TD002 As Key2, (TE008) AS Key3, (CASE WHEN (BOMTD.MODIFIER <> '') THEN B.MF002 ELSE A.MF002 END) As M_MF002 
                                                        FROM TK..BOMTD AS BOMTD WITH(NOLOCK)
                                                        INNER JOIN TK..BOMTE AS BOMTE WITH(NOLOCK) ON TE001 = TD001 AND TE002 = TD002
                                                        LEFT JOIN TK..ADMMF As A WITH(NOLOCK) On A.MF001 = BOMTD.CREATOR
                                                        LEFT JOIN TK..ADMMF As B WITH(NOLOCK) On B.MF001 = BOMTD.MODIFIER
                                                        WHERE TE004 = D.[品號] AND TE013 = D.[批號] AND TE007 = '{2}' AND TD012 = 'N' 
            
                                                        UNION ALL 
            
                                                        SELECT DISTINCT '8' As Type, TF001 As Key1, TF002 As Key2, (TF007) AS Key3, (CASE WHEN (BOMTF.MODIFIER <> '') THEN B.MF002 ELSE A.MF002 END) As M_MF002 
                                                        FROM TK..BOMTF AS BOMTF WITH(NOLOCK)
                                                        LEFT JOIN TK..ADMMF As A WITH(NOLOCK) On A.MF001 = BOMTF.CREATOR
                                                        LEFT JOIN TK..ADMMF As B WITH(NOLOCK) On B.MF001 = BOMTF.MODIFIER
                                                        WHERE TF004 = D.[品號] AND TF015 = D.[批號] AND TF008 = '{2}' AND TF010 = 'N' 
            
                                                        UNION ALL 
            
                                                        SELECT DISTINCT '9' As Type, TC001 As Key1, TC002 As Key2, (TE005) AS Key3, (CASE WHEN (MOCTC.MODIFIER <> '') THEN B.MF002 ELSE A.MF002 END) As M_MF002 
                                                        FROM TK..MOCTC AS MOCTC WITH(NOLOCK)
                                                        INNER JOIN TK..MOCTE AS MOCTE WITH(NOLOCK) ON TE001 = TC001 AND TE002 = TC002
                                                        LEFT JOIN TK..ADMMF As A WITH(NOLOCK) On A.MF001 = MOCTC.CREATOR
                                                        LEFT JOIN TK..ADMMF As B WITH(NOLOCK) On B.MF001 = MOCTC.MODIFIER
                                                        LEFT JOIN TK..CMSMQ AS CMSMQ WITH(NOLOCK) ON MQ001 = TC001
                                                        WHERE TE004 = D.[品號] AND TE010 = D.[批號] AND TE008 = '{2}' AND TC009 = 'N' AND MQ010 = -1 
            
                                                        UNION ALL 
            
                                                        SELECT DISTINCT 'A' As Type, TJ001 As Key1, TJ002 As Key2, SUM(TK007) AS Key3, (CASE WHEN (INVTJ.MODIFIER <> '') THEN B.MF002 ELSE A.MF002 END) As M_MF002 
                                                        FROM [TK].dbo.INVTJ As INVTJ WITH(NOLOCK)
                                                        INNER JOIN [TK].dbo.INVTK AS INVTK WITH(NOLOCK) ON TK001 = TJ001 AND TK002 = TJ002
                                                        LEFT JOIN [TK].dbo.ADMMF As A WITH(NOLOCK) On A.MF001 = INVTJ.CREATOR
                                                        LEFT JOIN [TK].dbo.ADMMF As B WITH(NOLOCK) On B.MF001 = INVTJ.MODIFIER
                                                        WHERE TK004 = D.[品號] AND TK018 = D.[批號] AND TK017 = '{2}' AND TJ010 = 'N'  
                                                        GROUP BY TJ001, TJ002, INVTJ.MODIFIER, A.MF002, B.MF002
                                                        HAVING (SUM(ISNULL(TK007, 0)) < 0) 
                                                    ) AS MoidA 
                                                    FOR XML PATH('')
                                                ) AS NOS
                                            ) AS NOS

                                            ORDER BY D.[品號], D.[批號];
                                    ", DateTime.Now.ToString("yyyy-MM-dd"), DateTime.Now.ToString("yyyyMMdd"), LA009, extraFilter);
        }

        public void SETFASTREPORT2()
        {
            Dictionary<string, string> dataSources = new Dictionary<string, string>();
            dataSources.Add("Table", SETFASETSQL2());
            dataSources.Add("Table1", SETFASETSQL3());

            report2 = ShowReport(report2, @"REPORT\每日盤點表-成品-訂單明細.frx", "dberp", dataSources, previewControl2);
        }

        public string SETFASETSQL2()
        {
            string twoMonthsAgo = DateTime.Now.AddMonths(-2).ToString("yyyyMMdd");
            string today = DateTime.Now.ToString("yyyyMMdd");
            string LA009 = comboBox1.SelectedValue.ToString();
            string dateFilter = BuildDateFilter("TK", checkBox1, dateTimePicker1, LA009);

            return string.Format(@"
                                    SELECT  LA001 AS '品號' ,MB002 AS '品名',MB003 AS '規格',LA016 AS '批號',CAST(SUM(LA005*LA011) AS INT) AS '庫存量' ,MB004 AS '單位'
                                    ,(SELECT ISNULL(SUM(NUM),0) FROM [TK].dbo.VCOPTDINVMD WHERE TD004=LA001 AND TD013>='{0}') AS '訂單需求量'
                                    ,(CAST(SUM(LA005*LA011) AS INT)-(SELECT ISNULL(SUM(NUM),0) FROM [TK].dbo.VCOPTDINVMD WHERE TD004=LA001 AND TD013>='{0}')) AS '需求差異量'
                                    ,DATEDIFF(DAY,(SELECT TOP 1 LA004 FROM [TK].dbo.INVLA A WHERE A.LA001=INVLA.LA001 AND A.LA016=INVLA.LA016 AND LA005='1') , '{1}' ) AS '在倉日期'
                                    ,DATEDIFF(DAY, '{1}',LEFT(LA016,8))  AS '有效天數'
                                    FROM [TK].dbo.INVLA WITH (NOLOCK)
                                    LEFT JOIN  [TK].dbo.INVMB WITH (NOLOCK) ON MB001=LA001
                                    WHERE  (LA009='{2}') {3}
                                    AND LA001 LIKE '4%'
                                    GROUP BY  LA001,MB002,MB003,LA016,MB023,MB198 ,MB004
                                    HAVING SUM(LA005*LA011)<>0
                                    ORDER BY  LA001,MB002,MB003,LA016,MB023,MB198 ,MB004
                                    ", twoMonthsAgo, today, LA009, dateFilter);
        }

        public string SETFASETSQL3()
        {
            string twoMonthsAgo = DateTime.Now.AddMonths(-2).ToString("yyyyMMdd");

            return string.Format(@"
                                    SELECT MV002 AS '業務員',TC053 AS '客戶',TD013 AS '預交日',NUM AS '訂單需求量',TD010 AS '單位',TC001 AS '訂單',TC002 AS '訂單號',TC004 AS '客戶代號',TD004 AS '品號',TD008 AS '訂單下量',TD009 AS '已出量',TD024 AS '贈品量',TD025 AS '已出贈品',MD004 AS '換算'
                                    FROM [TK].dbo.VCOPTDINVMD, [TK].dbo.COPTC
                                    LEFT JOIN [TK].dbo.CMSMV ON MV001=TC006
                                    WHERE TC001=TD001 AND TC002=TD002
                                    AND  TD013>='{0}'
                                    ", twoMonthsAgo);
        }

        public void SETFASTREPORT3()
        {
            string frxPath;
            if (comboBox2.Text.Equals("10001     成品倉"))
            {
                frxPath = @"REPORT\大潁-每日盤點表-成品.frx";
            }
            else if (comboBox2.Text.Equals("11001     大榮觀音倉"))
            {
                frxPath = @"REPORT\大潁-11001每日盤點表.frx";
            }
            else
            {
                frxPath = @"REPORT\大潁-每日盤點表.frx";
            }

            Dictionary<string, string> dataSources = new Dictionary<string, string>();
            dataSources.Add("Table", SETFASETSQL4());

            report3 = ShowReport(report3, frxPath, "dbconn", dataSources, previewControl3);
        }

        public string SETFASETSQL4()
        {
            // 大潁頁籤用自己的 checkBox2 / dateTimePicker3
            string dateFilter = BuildDateFilter("DY", checkBox2, dateTimePicker3, comboBox2.SelectedValue.ToString());

            if (comboBox2.Text.Equals("10001     成品倉"))
            {
                return string.Format(@"  
                                     SELECT 品號,品名,規格,批號,庫存量,單位,效期內的訂單需求量,效期內的訂單差異量,在倉日期,有效天數,總訂單需求量,業務
                                     FROM ( 
                                     SELECT   LA001 AS '品號' ,MB002 AS '品名',MB003 AS '規格',LA016 AS '批號'
                                     ,CAST(SUM(LA005*LA011) AS INT) AS '庫存量',MB004 AS '單位'
                                     ,CAST(((SELECT ISNULL(SUM(NUM),0) FROM [DY].dbo.VCOPTDINVMD WHERE TD004=LA001 AND TD013>='{0}' AND  TD013<=CONVERT(nvarchar,DATEADD (MONTH,-1*ROUND(MB023/3,0),CAST(LEFT(LA016,8) AS datetime)),112))) AS INT) AS '效期內的訂單需求量'     
                                     ,CAST((CAST(SUM(LA005*LA011) AS INT)-(SELECT ISNULL(SUM(NUM),0) FROM [DY].dbo.VCOPTDINVMD WHERE TD004=LA001 AND TD013>='{0}' AND  TD013<=CONVERT(nvarchar,DATEADD (MONTH,-1*ROUND(MB023/3,0),CAST(LEFT(LA016,8) AS datetime)),112)))  AS INT) AS '效期內的訂單差異量' 
                                     ,ISNULL ( DATEDIFF(DAY,(SELECT TOP 1 TF003 FROM [DY].dbo.MOCTF,[DY].dbo.MOCTG WHERE TF001=TG001 AND TF002=TG002 AND TG004=LA001 AND TG017=LA016 AND TG010=LA009),'{0}'),DATEDIFF(DAY,(SELECT TOP 1 LA004 FROM [DY].dbo.INVLA A WHERE A.LA001=INVLA.LA001 AND A.LA016=INVLA.LA016 AND A.LA005='1') ,'{0}') ) AS '在倉日期' 
                                    , DATEDIFF(DAY, '{0}',LEFT(LA016,8))  AS '有效天數' 
                                     ,CAST((SELECT ISNULL(SUM(NUM),0) FROM [DY].dbo.VCOPTDINVMD WHERE TD004=LA001 AND TD013>='{0}') AS INT) AS '總訂單需求量' 
                                     ,(SELECT TOP 1 TC006+' '+MV002 FROM [DY].dbo.COPTC,[DY].dbo.CMSMV WHERE TC006=MV001 AND  TC001+TC002 IN (SELECT TOP 1 TA026+TA027 FROM [DY].dbo.MOCTA WHERE TA001+TA002 IN (SELECT TOP 1 TG014+TG015 FROM [DY].dbo.MOCTG WHERE TG004=LA001 AND TG017=LA016))) AS '業務'
                                     FROM [DY].dbo.INVLA WITH (NOLOCK)  
                                     LEFT JOIN  [DY].dbo.INVMB WITH (NOLOCK) ON MB001=LA001   
                                     WHERE  (LA009='10001')   
                                     AND (LA001 LIKE '4%' OR LA001 LIKE '5%')
                                     GROUP BY  LA001,LA009,MB002,MB003,LA016,MB023,MB198,MB004    
                                     HAVING SUM(LA005*LA011)<>0 
                                     ) AS TEMP
                                     ORDER BY 品號
                                    ", DateTime.Now.ToString("yyyyMMdd"));
            }
            else if (comboBox2.Text.Equals("11001     大榮觀音倉"))
            {
                return @" 
                                    SELECT  LA001 AS '品號' ,MB002 AS '品名',MB003 AS '規格',LA016 AS '批號'  ,CAST(SUM(LA005*LA011) AS DECIMAL(18,0)) AS '庫存量',CONVERT(DECIMAL(16,3),SUM(LA005*LA013)) AS '庫存金額'  
                                    ,DATEADD(day,1,DATEADD(month,-1*10,CONVERT(DATETIME,LEFT(LA016,8))))
                                    ,CASE WHEN MB198='2' AND MB023>0 THEN DATEADD(day,1,DATEADD(month,-1*MB023,CONVERT(DATETIME,LEFT(LA016,8)))) END AS '製造日期'
                                    ,MB004 AS '庫存單位'
                                    ,(CAST(SUM(LA005*LA011) /240 AS DECIMAL(18,0))) AS '板數'
                                    ,'' AS '備註'
                                    FROM [DY].dbo.INVLA WITH (NOLOCK) LEFT JOIN  [DY].dbo.INVMB WITH (NOLOCK) ON MB001=LA001 
                                    WHERE  (LA009='11001') 
                                    GROUP BY  LA001,MB002,MB003,LA016,MB198,MB023,MB004
                                    HAVING SUM(LA005*LA011)<>0
                                    ORDER BY  LA001,MB002,MB003,LA016,MB198,MB023,MB004
                                    ";
            }
            else
            {
                return string.Format(@"

                                   SELECT  LA001 AS '品號' ,MB002 AS '品名',MB003 AS '規格',LA016 AS '批號'  ,CAST(SUM(LA005*LA011) AS DECIMAL(18,4)) AS '庫存量',CONVERT(DECIMAL(16,3),SUM(LA005*LA013)) AS '庫存金額'  
                                    FROM [DY].dbo.INVLA WITH (NOLOCK) LEFT JOIN  [DY].dbo.INVMB WITH (NOLOCK) ON MB001=LA001 
                                    WHERE  (LA009='{0}') {1}
                                    GROUP BY  LA001,MB002,MB003,LA016
                                    HAVING SUM(LA005*LA011)<>0
                                    ORDER BY  LA001,MB002,MB003,LA016
                                    ", comboBox2.SelectedValue.ToString(), dateFilter);
            }
        }

        #endregion

        #region BUTTON
        private void button1_Click(object sender, EventArgs e)
        {
            SETFASTREPORT(comboBox1.SelectedValue.ToString().Trim());
        }

        private void button2_Click(object sender, EventArgs e)
        {
            // 原 Excel 匯出功能已停用
        }

        private void button3_Click(object sender, EventArgs e)
        {
            SETFASTREPORT2();
        }

        private void button4_Click(object sender, EventArgs e)
        {
            SETFASTREPORT3();
        }
        #endregion
    }
}

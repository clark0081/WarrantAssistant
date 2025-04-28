using System;
using System.Collections.Generic;
using System.Data;
using System.Windows.Forms;
using System.Threading;
using System.Runtime.InteropServices;
using Infragistics.Win;
using Infragistics.Win.UltraWinGrid;
using System.Data.SqlClient;
using Microsoft.Office.Interop.Excel;
using System.IO;
using System.Text;
using EDLib.SQL;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium;
using OpenQA.Selenium.Support.UI;
using System.Net;
using System.Drawing;


namespace WarrantAssistant
{
    public partial class MainForm:Form
    {
        //private delegate void ShowHandler();  
        
        public SqlConnection conn = new SqlConnection(GlobalVar.loginSet.warrantassistant45);

        public MainForm() {
            InitializeComponent();
        }

        //private SafeQueue workQueue = new SafeQueue();
        private Thread workThread;
        private Thread workThread2;
        /*
        public void AddWork(Work work)
        {
            workQueue.Enqueue(work);
        }
        */

        private void MainForm_Load(object sender, EventArgs e) {
            GlobalVar.mainForm = this;
            Text += System.Reflection.Assembly.GetExecutingAssembly().GetName().Version;

            //Text += System.Windows.Forms.Application.ProductVersion;
            FrmLogIn frmLogin = new FrmLogIn();
            if (!frmLogin.TryIPLogin()) {
                MessageBox.Show("Auto login failed. Please e-mail your IP address to clark.lin@kgi.com");

                Close();
                //frmLogin.ShowDialog();
            }
            if (!frmLogin.loginOK)
                Close();
           
        }

        public void Start() {//在FrmLogIn 登入成功後會執行Start()
            
            toolStripComboBox1.Visible = false;
            if (GlobalVar.globalParameter.userGroup == "TR") {
                行政ToolStripMenuItem.Visible = false;
                財工ToolStripMenuItem.Visible = false;            
            }

            if (GlobalVar.globalParameter.userGroup == "AD") {
                traderToolStripMenuItem.Visible = false;
                財工ToolStripMenuItem.Visible = false;
            }
            
       
            代理人發行條件輸入ToolStripMenuItem.Visible = false;
            代理人增額條件輸入ToolStripMenuItem.Visible = false;
            SetUltraGrid(dtInfo, ultraGrid1);
            SetUltraGrid(dtAnnounce, ultraGrid2);
            if (GlobalVar.globalParameter.userGroup != "AD")
            {
                toolStripComboBox1.Visible = true;
                foreach (var item in GlobalVar.globalParameter.traders)
                    toolStripComboBox1.Items.Add(item);
                toolStripComboBox1.Items.Add("All");
                toolStripComboBox1.Text = GlobalVar.globalParameter.userID;
            }
            GlobalVar.autoWork = new AutoWork();

            workThread = new Thread(new ThreadStart(RoutineWork));
            workThread.Start();
            
        }
        private void RoutineWork() {
            try {
                for (; ; ) {
                    try {
                        if (ultraGrid1.InvokeRequired)
                            ultraGrid1.Invoke(new System.Action(LoadUltraGrid1));
                        else
                            LoadUltraGrid1();

                        if (ultraGrid2.InvokeRequired)
                            ultraGrid2.Invoke(new System.Action(LoadUltraGrid2));
                        else
                            LoadUltraGrid2();

                    } catch (ThreadAbortException) {
                        return;
                    } catch (Exception ex) {
                        MessageBox.Show("In main form routine work in for loop " + ex.Message);
                    }
                    Thread.Sleep(10000);
                }

            } catch (Exception ex) {
                //MessageBox.Show("In main form routine work "+ex.Message);
            }
        }
        /*
        private void RoutineWork()
        {
            try
            {
                for (; ; )
                {
                    while (workQueue.Count > 0)
                    {
                        try
                        {
                            object obj = workQueue.Dequeue();
                            if (obj != null)
                            {
                                Work work = (Work)obj;
                                WorkState workstate = work.DoWork();
                                work.Close();
                            }
                        }
                        catch (ThreadAbortException tex)
                        {
                            //MessageBox.Show(tex.Message);
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show(ex.Message);
                        }

                    }
                    Thread.Sleep(1000);
                }
            }
            catch (Exception ex)
            {
            }
        }
        */
        public System.Data.DataTable dtInfo = new System.Data.DataTable();
        public System.Data.DataTable dtAnnounce = new System.Data.DataTable();
        public static int dv_Rows_Count = 0;
        public static int dv_issueDate_Rows_Count = 0;


        public void SetUltraGrid(System.Data.DataTable dt, UltraGrid grid) {
            dt.Columns.Add("時間", typeof(string));
            dt.Columns.Add("內容", typeof(string));
            dt.Columns.Add("人員", typeof(string));
            grid.DataSource = dt;

            UltraGridBand band0 = grid.DisplayLayout.Bands[0];
            band0.Columns["時間"].Width = 60;
            band0.Columns["人員"].Width = 30;
            //ultraGrid1.DisplayLayout.Bands[0].Override.HeaderAppearance.TextHAlign = Infragistics.Win.HAlign.Left;
            band0.ColHeadersVisible = false;
            grid.DisplayLayout.AutoFitStyle = AutoFitStyle.ResizeAllColumns;
            grid.DisplayLayout.Override.CellAppearance.BorderAlpha = Alpha.Transparent;
            grid.DisplayLayout.Override.RowAppearance.BorderAlpha = Alpha.Transparent;
            band0.Columns[2].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right;
            band0.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
            band0.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
            band0.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.False;

            band0.Columns["時間"].CellActivation = Activation.NoEdit;
            band0.Columns["內容"].CellActivation = Activation.NoEdit;
            band0.Columns["人員"].CellActivation = Activation.NoEdit;
        }
        //清空再刷新
        private void LoadUltraGrid(System.Data.DataTable dt, string infoOrAnnounce) {
            try {
                string sql = "";
                if (infoOrAnnounce == "Announce")//不同排序方式
                {
                    if (toolStripComboBox1.Text.TrimStart('0') == "All")
                    {
                        sql = $@"SELECT [MDate]
                                  ,[InformationContent]
                                  ,[MUser]
                              FROM [WarrantAssistant].[dbo].[InformationLog]
                              WHERE InformationType='{infoOrAnnounce}' AND [InformationContent] NOT LIKE '%刪除%'";
                        sql += $"AND CONVERT(VARCHAR,Date,112) >'{GlobalVar.globalParameter.lastTradeDate.ToString("yyyyMMdd")}' ORDER BY (CASE WHEN [InformationContent] like '%可以發行%' THEN 1 WHEN [InformationContent] like '%可以發行(除權息)%' THEN 2 WHEN [InformationContent] like '%不可發(除權息)%' THEN 4 ELSE 3 END )";

                    }
                    else
                    {
                        sql = $@"SELECT [MDate]
                                  ,[InformationContent]
                                  ,[MUser]
                              FROM [WarrantAssistant].[dbo].[InformationLog]
                              WHERE InformationType='{infoOrAnnounce}' AND [InformationContent] NOT LIKE '%刪除%'";
                        sql += $"AND CONVERT(VARCHAR,Date,112) >'{GlobalVar.globalParameter.lastTradeDate.ToString("yyyyMMdd")}'ORDER BY (CASE WHEN [InformationContent] like '%可以發行%' THEN 1 WHEN [InformationContent] like '%可以發行(除權息)%' THEN 2 WHEN [InformationContent] like '%不可發(除權息)%' THEN 4 ELSE 3 END )";
                    }
                }
                else
                {
                    if (toolStripComboBox1.Text.TrimStart('0') == "All")
                    {
                        sql = $@"SELECT [MDate]
                                  ,[InformationContent]
                                  ,[MUser]
                              FROM [WarrantAssistant].[dbo].[InformationLog]
                              WHERE InformationType='{infoOrAnnounce}' AND [InformationContent] NOT LIKE '%刪除%'";
                        sql += $"AND CONVERT(VARCHAR,Date,112) >'{GlobalVar.globalParameter.lastTradeDate.ToString("yyyyMMdd")}' ORDER BY MDate DESC";

                    }
                    else
                    {
                        sql = $@"SELECT [MDate]
                                  ,[InformationContent]
                                  ,[MUser]
                              FROM [WarrantAssistant].[dbo].[InformationLog]
                              WHERE InformationType='{infoOrAnnounce}' AND [InformationContent] NOT LIKE '%刪除%'";
                        sql += $"AND CONVERT(VARCHAR,Date,112) >'{GlobalVar.globalParameter.lastTradeDate.ToString("yyyyMMdd")}'ORDER BY MDate DESC";
                    }
                }
                System.Data.DataTable dv = MSSQL.ExecSqlQry(sql, GlobalVar.loginSet.warrantassistant45);
                if (infoOrAnnounce == "Announce")
                {
                    try
                    {
                        string sql_issueDate = "";
                        if (toolStripComboBox1.Text.TrimStart('0') == "All")
                            sql_issueDate = $@"SELECT [TraderID], [DateContent]
                                                 FROM[WarrantAssistant].[dbo].[WarrantIssueDate]
                                                   ORDER BY (CASE [DateType] WHEN 'conferenceDate' THEN '01'
						                            WHEN 'CashDividendDate' THEN '02'
							                        WHEN 'StockDividendDate' THEN '02'
							                        WHEN 'shrHoldersDate' THEN '03'
                                                    WHEN 'ReductionCapital' THEN '04' END), [Date]";
                        else
                            sql_issueDate = $@"SELECT [TraderID], [DateContent]
                                                 FROM[WarrantAssistant].[dbo].[WarrantIssueDate]
                                                   WHERE [TraderID] ='{toolStripComboBox1.Text.TrimStart('0')}'
                                                   ORDER BY (CASE [DateType] WHEN 'conferenceDate' THEN '01'
						                            WHEN 'CashDividendDate' THEN '02'
							                        WHEN 'StockDividendDate' THEN '02'
							                        WHEN 'shrHoldersDate' THEN '03'
                                                    WHEN 'ReductionCapital' THEN '04' END), [Date]";
                        
                        System.Data.DataTable dv_issueDate = MSSQL.ExecSqlQry(sql_issueDate, GlobalVar.loginSet.warrantassistant45);
                        //MessageBox.Show($@"{dv_Rows_Count} {dv.Rows.Count} {dv_issueDate_Rows_Count} {dv_issueDate.Rows.Count}");
                        if ((dv_Rows_Count == dv.Rows.Count) && (dv_issueDate_Rows_Count == dv_issueDate.Rows.Count) )
                        {
                            //MessageBox.Show($@"{dv_Rows_Count} {dv.Rows.Count} {dv_issueDate_Rows_Count} {dv_issueDate.Rows.Count}");
                            return;
                        }
                        dv_Rows_Count = dv.Rows.Count;
                        dv_issueDate_Rows_Count = dv_issueDate.Rows.Count;
                        dt.Rows.Clear();
                        string sql_StopTradeTemp = $@"SELECT [股票代號]
                                                              ,[股票名稱]
                                                              ,CONVERT(VARCHAR,[暫停交易日期],112) AS 暫停交易日期
                                                              ,CONVERT(VARCHAR,[恢復交易日期],112) AS 恢復交易日期
                                                          FROM [WarrantAssistant].[dbo].[暫停交易標的]
                                                          WHERE [日期] = '{DateTime.Today.ToString("yyyyMMdd")}' AND [恢復交易日期] > '{DateTime.Today.ToString("yyyyMMdd")}' AND [暫停交易日期] <= '{DateTime.Today.ToString("yyyyMMdd")}'
                                                          AND [股票代號] IN ( 
                                                            SELECT  [UnderlyingID] FROM [WarrantAssistant].[dbo].[WarrantUnderlyingSummary]
                                                            UNION
                                                            SELECT  DISTINCT [UID]  FROM [TwData].[dbo].[V_WarrantTrading] WHERE [TDate] = (SELECT Max([TDate]) FROM [TwData].[dbo].[WarrantFlow]) AND [IssuerName] = '9200')";
                        System.Data.DataTable dv_StopTradeTemp = MSSQL.ExecSqlQry(sql_StopTradeTemp, GlobalVar.loginSet.warrantassistant45);
                        foreach (DataRow drv in dv_StopTradeTemp.Rows)
                        {
                            DataRow dr = dt.NewRow();
                            string uid = drv["股票代號"].ToString();
                            string uname = drv["股票名稱"].ToString();
                            string t1 = drv["暫停交易日期"].ToString();
                            string t2 = drv["恢復交易日期"].ToString();
                            dr["時間"] = DateTime.Today.ToString("yyyy/MM/dd HH:mm:ss");
                            dr["內容"] = uid + " " + uname + " 暫停交易，恢復交易日期:" + t2;
                            dr["人員"] = "";
                            dt.Rows.Add(dr);
                        }
                        //處置股票
                        /*
                        string sql_DisposeTemp = $@" SELECT * FROM (
                                 SELECT  A.[UnderlyingID],B.[TraderAccount] AS TraderID,RTrim(A.[UnderlyingName]) AS UnderlyingName,CONVERT(VARCHAR, A.[DisposeEndDate],112) AS D
                                                                      FROM [WarrantAssistant].[dbo].[WarrantIssueCheck] AS A
                                                                      LEFT JOIN [TwData].[dbo].[Underlying_Trader] AS B ON A.[UnderlyingID] =B.[UID]
                                                                      WHERE A.[DisposeEndDate] >= '{DateTime.Today.ToString("yyyyMMdd")}' AND A.[IsQuaterUnderlying] = 'Y'
                                 UNION 
                                 SELECT  A.[UnderlyingID],B.[TraderAccount] AS TraderID,RTrim(A.[UnderlyingName]) AS UnderlyingName,CONVERT(VARCHAR, A.[DisposeEndDate],112) AS D
                                                                      FROM [WarrantAssistant].[dbo].[WarrantIssueCheckNoUnderlying] AS A
                                                                      LEFT JOIN [TwData].[dbo].[Underlying_Trader] AS B ON A.[UnderlyingID] =B.[UID]
                                                                      WHERE A.[DisposeEndDate] >= '{DateTime.Today.ToString("yyyyMMdd")}'
                                ) AS A WHERE A.UnderlyingID IN (SELECT DISTINCT [UID] FROM [TwData].[dbo].[V_WarrantTrading] WHERE [TDate] = (SELECT MAX(TDate) FROM [TwData].[dbo].[WarrantFlow])) ORDER BY D"; 
                        */

                        string sql_DisposeTemp = $@"SELECT * FROM (
                                    SELECT A.UnderlyingID,B.TraderID,A.UnderlyingName,A.D
                                    ,Row_Number() over (Partition by A.UnderlyingID ORDER by D DESC) AS RowN FROM (
                                 SELECT  A.[UnderlyingID],B.[TraderAccount] AS TraderID,RTrim(A.[UnderlyingName]) AS UnderlyingName,CONVERT(VARCHAR, A.[DisposeEndDate],112) AS D
                                                                      FROM [WarrantAssistant].[dbo].[WarrantIssueCheck] AS A
                                                                      LEFT JOIN [TwData].[dbo].[Underlying_Trader] AS B ON A.[UnderlyingID] =B.[UID]
                                                                      WHERE A.[DisposeEndDate] >= '{DateTime.Today.ToString("yyyyMMdd")}' AND A.[IsQuaterUnderlying] = 'Y'
                                 UNION 
                                 SELECT  A.[UnderlyingID],B.[TraderAccount] AS TraderID,RTrim(A.[UnderlyingName]) AS UnderlyingName,CONVERT(VARCHAR, A.[DisposeEndDate],112) AS D
                                                                      FROM [WarrantAssistant].[dbo].[WarrantIssueCheckNoUnderlying] AS A
                                                                      LEFT JOIN [TwData].[dbo].[Underlying_Trader] AS B ON A.[UnderlyingID] =B.[UID]
                                                                      WHERE A.[DisposeEndDate] >= '{DateTime.Today.ToString("yyyyMMdd")}'
								 UNION
								 (SELECT A.StockNo AS [UnderlyingID],'' AS TraderID,B.股票名稱 AS UnderlyingName,CONVERT(VARCHAR, GETDATE(),112) AS D FROM 
                                (
								    SELECT DISTINCT [STKNO] AS StockNo FROM (SELECT [STKNO] FROM [TWCMData].[dbo].[HHPT30M] WHERE ((SETTYPE=2 AND MARKW=0) OR MARKW >= 1) AND [MODIFYTIME] = CONVERT(VARCHAR,GETDATE(),112) ) AS A
							        UNION (SELECT [STKNO] FROM [TWCMData].[dbo].[OCPT30M] WHERE ((SETTYPE=2 AND MARKW=0) or MARKW >= 1) AND [MODIFYTIME] = CONVERT(VARCHAR,GETDATE(),112))

                                ) AS A
								LEFT JOIN
								(SELECT  [股票代號],[股票名稱] FROM [TwCMData].[dbo].[上市櫃公司基本資料] WHERE [年度] = (SELECT MAX(年度) FROM [TwCMData].[dbo].[上市櫃公司基本資料])) AS B ON　A.StockNo = B.股票代號)
                                ) AS A 
								LEFT JOIN [10.101.10.5].[WMM3].[dbo].[UnderlyingLinkTrader] AS B ON A.UnderlyingID = B.[UnderlyingID]
								WHERE A.UnderlyingID IN (SELECT DISTINCT [標的代號] FROM [TwCMData].[dbo].[Warrant總表] WHERE [日期] = (SELECT MAX(TDate) FROM [TwData].[dbo].[WarrantFlow])) ) AS A WHERE RowN = 1 ORDER By A.D";

                        System.Data.DataTable dv_DisposeTemp = MSSQL.ExecSqlQry(sql_DisposeTemp, GlobalVar.loginSet.warrantassistant45);
                        foreach (DataRow drv in dv_DisposeTemp.Rows)
                        {
                            DataRow dr = dt.NewRow();
                            string trader  = drv["TraderID"].ToString();
                            string uid = drv["UnderlyingID"].ToString();
                            string uname = drv["UnderlyingName"].ToString();
                            string t1 = drv["D"].ToString();
                            dr["時間"] = DateTime.Today.ToString("yyyy/MM/dd HH:mm:ss");
                            dr["內容"] = trader + " " + uid + " " + uname + " 處置，處置結束日期:" + t1;
                            dr["人員"] = "";
                            dt.Rows.Add(dr);
                        }

                        foreach (DataRow drv in dv.Rows)
                        {
                            DataRow dr = dt.NewRow();
                            DateTime md = Convert.ToDateTime(drv["MDate"]);
                            dr["時間"] = md.ToString("yyyy/MM/dd HH:mm:ss");
                            dr["內容"] = drv["InformationContent"].ToString();
                            dr["人員"] = drv["MUser"].ToString();
                            dt.Rows.Add(dr);
                        }

                        //被注意的日期

                        Dictionary<string, DateTime> watchDate = new Dictionary<string, DateTime>();
                        string sqlWatchCount = $@"SELECT CONVERT(VARCHAR,[日期],112) AS 日期,[股票代號] FROM [WarrantAssistant].[dbo].[注意股票] WHERE [日期] >= '{EDLib.TradeDate.LastNTradeDate(6).ToString("yyyyMMdd")}'";
                        System.Data.DataTable dtWatchCount = MSSQL.ExecSqlQry(sqlWatchCount, GlobalVar.loginSet.warrantassistant45);

                        foreach(DataRow dr in dtWatchCount.Rows)
                        {
                            string stockID = dr["股票代號"].ToString();
                            

                            string d = dr["日期"].ToString();
                            DateTime dd = DateTime.ParseExact(d, "yyyyMMdd", null);
                            if (!watchDate.ContainsKey(stockID))
                                watchDate.Add(stockID, dd);
                            else
                            {
                                if (watchDate[stockID] <= dd)
                                    watchDate[stockID] = dd;
                            }
                            
                        }
                        
                        string sql_WatchCount = "";
                        if (toolStripComboBox1.Text.TrimStart('0') == "All")
                            sql_WatchCount = $@"SELECT  A.[UnderlyingID], A.[MDate], B.[TraderAccount]    
                                                FROM [WarrantAssistant].[dbo].[WarrantIssueCheck] as A
                                                LEFT JOIN [TwData].[dbo].[Underlying_Trader] as B on A.[UnderlyingID] = B.[UID]
                                                LEFT JOIN [WarrantAssistant].[dbo].[WarrantUnderlyingSummary] as C ON A.UnderlyingID = C.UnderlyingID
                                                WHERE A.[WatchCount] = 1 AND C.Issuable = 'Y' AND A.[IsQuaterUnderlying] = 'Y'";
                        else
                            sql_WatchCount = $@"SELECT  A.[UnderlyingID], A.[MDate], B.[TraderAccount]    
                                                FROM [WarrantAssistant].[dbo].[WarrantIssueCheck] as A
                                                LEFT JOIN [TwData].[dbo].[Underlying_Trader] as B on A.[UnderlyingID] = B.[UID]
                                                LEFT JOIN [WarrantAssistant].[dbo].[WarrantUnderlyingSummary] as C ON A.UnderlyingID = C.UnderlyingID
                                                WHERE B.[TraderAccount] = '{toolStripComboBox1.Text.TrimStart('0')}'
                                                AND A.[WatchCount] = 1 AND C.Issuable = 'Y' AND A.[IsQuaterUnderlying] = 'Y'";
                        System.Data.DataTable dv_WatchCount = MSSQL.ExecSqlQry(sql_WatchCount, GlobalVar.loginSet.warrantassistant45);

                       
                        foreach (DataRow dr_w in dv_WatchCount.Rows)
                        {
                            DataRow dr = dt.NewRow();
                            DateTime md = Convert.ToDateTime(dr_w["MDate"]);
                            dr["時間"] = md.ToString("yyyy/MM/dd HH:mm:ss");
                            DateTime dd = watchDate[dr_w["UnderlyingID"].ToString()];
                            dr["內容"] = dr_w["UnderlyingID"].ToString() + " " + "注意次數 1" + $@"({dd.ToString("yyyyMMdd")})";
                            dr["人員"] = toolStripComboBox1.Text.ToString();
                            dt.Rows.Add(dr);
                        }
                    
              

                        string sql_Transfer = "";
                        if (toolStripComboBox1.Text.TrimStart('0') == "All")
                            sql_Transfer = $@"SELECT A.[UID], SUM([TransferAmount]) AS Amt
                                                ,SUM([TransferShares]) AS Shr, B.[TraderAccount]
                                              FROM [WarrantAssistant].[dbo].[TransferUnderlying] AS A
                                              LEFT JOIN [TwData].[dbo].[Underlying_Trader] as B on A.[UID] = B.[UID]
                                              GROUP BY A.[UID], B.[TraderAccount]";
                        else
                            sql_Transfer = $@"SELECT A.[UID], SUM([TransferAmount]) AS Amt
                                                ,SUM([TransferShares]) AS Shr, B.[TraderAccount]
                                              FROM [WarrantAssistant].[dbo].[TransferUnderlying] AS A
                                              LEFT JOIN [TwData].[dbo].[Underlying_Trader] as B on A.[UID] = B.[UID]
                                              WHERE B.[TraderAccount] ='{toolStripComboBox1.Text.TrimStart('0')}'
                                              GROUP BY A.[UID], B.[TraderAccount]";
                        System.Data.DataTable dv_Transfer = MSSQL.ExecSqlQry(sql_Transfer, GlobalVar.loginSet.warrantassistant45);

                        foreach (DataRow dr in dv_Transfer.Rows)
                        {
                            string uid = dr["UID"].ToString();
                            double Amt = Convert.ToDouble(dr["Amt"].ToString());
                            double Shr = Convert.ToDouble(dr["Shr"].ToString());
                            if(Amt >= 20000)
                            {

                                string sql_mindate = $@"SELECT MIN(DataDate) AS MinDate
                                                          FROM [WarrantAssistant].[dbo].[TransferUnderlying]
                                                          WHERE [UID] = '{uid}'";
                                System.Data.DataTable dt_mindate = MSSQL.ExecSqlQry(sql_mindate, GlobalVar.loginSet.warrantassistant45);

                                bool b1 = DateTime.TryParse(dt_mindate.Rows[0]["MinDate"].ToString(), out DateTime t1);
                                string str1 = t1.ToString("yyyyMMdd");
                                DataRow dr2 = dt.NewRow();
                                
                                dr2["時間"] = DateTime.Today.ToString("yyyy/MM/dd HH:mm:ss");
                                dr2["內容"] = $@"{uid} 申報轉讓(自{str1}) : {Math.Round(Amt,1)}千 / {Math.Round(Shr,1)}張";
                                dr2["人員"] = toolStripComboBox1.Text.ToString();
                                dt.Rows.Add(dr2);
                            }
                        }
                        foreach (DataRow drv in dv_issueDate.Rows)
                        {
                            DataRow dr = dt.NewRow();
             
                            dr["時間"] = DateTime.Today.ToString("yyyy/MM/dd HH:mm:ss");
                            dr["內容"] = drv["DateContent"].ToString();
                            dr["人員"] = "";
                            dt.Rows.Add(dr);
                        }
                        
                        string sql_StopTrade = $@"SELECT [股票代號]
                                                  ,[股票名稱]
                                                  ,CONVERT(VARCHAR,[停止買賣開始日期],112) AS 停止買賣開始日期 
                                              FROM [WarrantAssistant].[dbo].[停止買賣標的] WHERE [日期] = '{DateTime.Today.ToString("yyyyMMdd")}'
                                              AND [股票代號] IN (SELECT  [UnderlyingID] FROM [WarrantAssistant].[dbo].[WarrantUnderlyingSummary])";
                        System.Data.DataTable dv_StopTrade = MSSQL.ExecSqlQry(sql_StopTrade, GlobalVar.loginSet.warrantassistant45);
                        foreach (DataRow drv in dv_StopTrade.Rows)
                        {
                            DataRow dr = dt.NewRow();
                            string uid = drv["股票代號"].ToString();
                            string uname = drv["股票名稱"].ToString();
                            string t = drv["停止買賣開始日期"].ToString();
                            dr["時間"] = DateTime.Today.ToString("yyyy/MM/dd HH:mm:ss");
                            dr["內容"] = uid + " " + uname + " 停止買賣開始日期:" + t;
                            dr["人員"] = "";
                            dt.Rows.Add(dr);
                        }

                        string sql_CreditError = $@"SELECT  [MDate]
                                                      ,[InformationContent]
                                                      ,[MUser]
                                                      ,[Date]
                                                  FROM [WarrantAssistant].[dbo].[InformationLog]
                                                  WHERE [InformationType] = 'CreditError' AND [MDate] > '{DateTime.Today.ToString("yyyyMMdd")}'";
                        System.Data.DataTable dv_CreditError = MSSQL.ExecSqlQry(sql_CreditError, GlobalVar.loginSet.warrantassistant45);
                        foreach (DataRow drv in dv_CreditError.Rows)
                        {
                            DataRow dr = dt.NewRow();
                            string content = drv["InformationContent"].ToString();
                            dr["時間"] = DateTime.Today.ToString("yyyy/MM/dd HH:mm:ss");
                            dr["內容"] = content;
                            dr["人員"] = "";
                            dt.Rows.Add(dr);
                        }

                    }
                    catch (Exception ex)
                    {
                        //MessageBox.Show("in Announce"+ex.Message);
                    }
                }
                else
                {
                    /*
                    if (dt.Rows.Count == dv.Rows.Count)
                        return;
                    */
                    //記錄要搶標的
                    List<string> Compete = new List<string>();
                    dt.Rows.Clear();
                    foreach (DataRow drv in dv.Rows)
                    {
                        DataRow dr = dt.NewRow();

                        DateTime md = Convert.ToDateTime(drv["MDate"]);
                        dr["時間"] = md.ToString("yyyy/MM/dd HH:mm:ss");
                        string content = drv["InformationContent"].ToString();
                        dr["內容"] = content;
                        int l = content.Length;
                        if (content.Contains("為要搶的標的"))
                            Compete.Add(content.Substring(0,l-17));
                        dr["人員"] = drv["MUser"].ToString();
                        dt.Rows.Add(dr);
                    }

                    string sqltemp = $@"SELECT B.[UpdateTime],C.[WarrantName], A.[DataName], B.[FromValue],B.[ToValue], B.[TraderID]
                                    FROM (SELECT [SerialNumber], [DataName], MAX([UpdateCount]) AS UpdateCount
                                          FROM [WarrantAssistant].[dbo].[ApplyTotalRecord]
                                          WHERE [UpdateTime] >= CONVERT(varchar, getdate(), 112) 
                                          GROUP BY [SerialNumber] ,[DataName]) AS A
                                    LEFT JOIN [WarrantAssistant].[dbo].[ApplyTotalRecord] AS B ON A.[SerialNumber] = B.[SerialNumber] AND A.[DataName] =B.[DataName] AND A.[UpdateCount] =B.[UpdateCount]
									LEFT JOIN [WarrantAssistant].[dbo].[ApplyTotalList] AS C ON A.[SerialNumber] = C.[SerialNum]
                                    
                                    WHERE A.[UpdateCount] > 0 AND A.[UpdateCount] < 9999
                                    ORDER BY B.[UpdateTime]";
                    System.Data.DataTable dttemp = MSSQL.ExecSqlQry(sqltemp, GlobalVar.loginSet.warrantassistant45);

                    foreach (DataRow dr in dttemp.Rows)
                    {
                        DateTime md = Convert.ToDateTime(dr["UpdateTime"]);
                        
                        string wname = dr["WarrantName"].ToString();
                        string dataname = dr["DataName"].ToString();
                        string fromvalue = dr["FromValue"].ToString();
                        string tovalue = dr["ToValue"].ToString();
                        string trader = dr["TraderID"].ToString();
                        if(!(dataname=="IssueNum" && Compete.Contains(wname)))
                        {
                            DataRow drv = dt.NewRow();
                            drv["時間"] = md.ToString("yyyy/MM/dd HH:mm:ss");
                            drv["內容"] = $@"{wname}改{dataname} {fromvalue}>{tovalue}";
                            drv["人員"] = trader;
                            dt.Rows.Add(drv);
                        }
                    }

                    string sqltemp2 = $@"SELECT  [MDate], [InformationContent], [MUser]
                                  FROM [WarrantAssistant].[dbo].[InformationLog]
                                  WHERE [MDate] > '{DateTime.Today.ToString("yyyyMMdd")}' AND [InformationContent] LIKE '%刪除%'
                                  ORDER BY [MDate]";
                    System.Data.DataTable dttemp2 = MSSQL.ExecSqlQry(sqltemp2, GlobalVar.loginSet.warrantassistant45);

                    foreach (DataRow dr in dttemp2.Rows)
                    {
                        DateTime md = Convert.ToDateTime(dr["MDate"]);

                        string content = dr["InformationContent"].ToString();
                        string user = dr["MUser"].ToString();
                        DataRow drv = dt.NewRow();
                        drv["時間"] = md.ToString("yyyy/MM/dd HH:mm:ss");
                        drv["內容"] = content;
                        drv["人員"] = user;
                        dt.Rows.Add(drv);
                    }
                    
                }
            } catch (Exception ex) {
                MessageBox.Show(ex.Message);
            }
        }
        public void LoadUltraGrid1() {
            LoadUltraGrid(dtInfo, "Info");
        }
        public void LoadUltraGrid2() {
            LoadUltraGrid(dtAnnounce, "Announce");
        }

        private void 標的SummaryToolStripMenuItem_Click(object sender, EventArgs e) {
            GlobalUtility.MenuItemClick<FrmUnderlyingSummary>();
        }

        private void 標的發行檢查ToolStripMenuItem_Click(object sender, EventArgs e) {
            GlobalUtility.MenuItemClick<FrmIssueCheck>();
        }

        private void put發行檢查ToolStripMenuItem_Click(object sender, EventArgs e) {
            //GlobalUtility.MenuItemClick<FrmIssueCheckPut>();
        }

        private void 已發行權證ToolStripMenuItem_Click(object sender, EventArgs e) {
            GlobalUtility.MenuItemClick<FrmWarrant>();
        }

        private void 可增額列表ToolStripMenuItem1_Click(object sender, EventArgs e) {
            GlobalUtility.MenuItemClick<FrmReIssueInput>();
        }

        private void 可增額列表ToolStripMenuItem_Click(object sender, EventArgs e) {
            GlobalUtility.MenuItemClick<FrmReIssuable>();
        }

        private void 試算表ToolStripMenuItem_Click(object sender, EventArgs e) {
            GlobalUtility.MenuItemClick<Frm71>();
        }

        private void 發行條件輸入ToolStripMenuItem_Click(object sender, EventArgs e) {
            GlobalUtility.MenuItemClick<FrmApply>();
        }

        private void MainForm_FormClosed(object sender, FormClosedEventArgs e) {
            if (workThread != null && workThread.IsAlive) { workThread.Abort(); }
            GlobalUtility.Close();
        }

        private void toolStripButton1_Click(object sender, EventArgs e) {
            string info = "";
            info = toolStripTextBox1.Text;
            if (info != "") {
                GlobalUtility.LogInfo("Announce", info);
                toolStripTextBox1.Text = "";
                LoadUltraGrid2();//
            }
        }

        private void 增額條件輸入ToolStripMenuItem_Click(object sender, EventArgs e) {
            GlobalUtility.MenuItemClick<FrmReIssue>();
        }

        private void 搶額度總表含增額ToolStripMenuItem_Click(object sender, EventArgs e) {
            GlobalUtility.MenuItemClick<FrmApplyTotalList>();
        }

        private void 發行總表ToolStripMenuItem_Click(object sender, EventArgs e) {
            GlobalUtility.MenuItemClick<FrmIssueTotal>();
        }

        private void 增額總表ToolStripMenuItem_Click(object sender, EventArgs e) {
            GlobalUtility.MenuItemClick<FrmReIssueTotal>();
        }
        private void 修改檔案名稱ToolStripMenuItem_Click(object sender, EventArgs e) {
            GlobalUtility.MenuItemClick<FrmRename>();
        }

        private void 代理人發行條件輸入ToolStripMenuItem_Click(object sender, EventArgs e) {
            foreach (Form iForm in System.Windows.Forms.Application.OpenForms) {
                if (iForm.GetType() == typeof(FrmApply)) {
                    iForm.BringToFront();
                    return;
                }
            }
            FrmApply frmApplyDeputy = new FrmApply();
            frmApplyDeputy.userID = GlobalVar.globalParameter.userDeputy;
            frmApplyDeputy.StartPosition = FormStartPosition.CenterScreen;
            frmApplyDeputy.Show();

        }

        private void 代理人增額條件輸入ToolStripMenuItem_Click(object sender, EventArgs e) {
            foreach (Form iForm in System.Windows.Forms.Application.OpenForms) {
                if (iForm.GetType() == typeof(FrmReIssue)) {
                    iForm.BringToFront();
                    return;
                }
            }

            FrmReIssue frmReIssue = new FrmReIssue();
            frmReIssue.userID = GlobalVar.globalParameter.userDeputy;
            frmReIssue.StartPosition = FormStartPosition.CenterScreen;
            frmReIssue.Show();

        }

        private void 已發權證條件發行ToolStripMenuItem_Click(object sender, EventArgs e) {
            GlobalUtility.MenuItemClick<FrmIssueByCurrent>();
        }

        private void 詳細LOGToolStripMenuItem_Click(object sender, EventArgs e) {
            GlobalUtility.MenuItemClick<FrmLog>();
        }

        private void 轉申請發行TXTToolStripMenuItem_Click(object sender, EventArgs e) {
            /*
            string sql5 = "SELECT [SerialNumber], [TempName] FROM [EDIS].[dbo].[ApplyOfficial]";
            System.Data.DataTable dv = MSSQL.ExecSqlQry(sql5, GlobalVar.loginSet.edisSqlConnString);

            string cmdText = "UPDATE [ApplyTotalList] SET WarrantName=@WarrantName WHERE SerialNum=@SerialNum";
            List<System.Data.SqlClient.SqlParameter> pars = new List<System.Data.SqlClient.SqlParameter>();
            pars.Add(new SqlParameter("@WarrantName", SqlDbType.VarChar));
            pars.Add(new SqlParameter("@SerialNum", SqlDbType.VarChar));
            SQLCommandHelper h = new SQLCommandHelper(GlobalVar.loginSet.edisSqlConnString, cmdText, pars);

            bool updated = false;
            foreach (DataRow dr in dv.Rows)
            {
                //string serialNum = dr["SerialNum"].ToString();
                //string warrantName = dr["WarrantName"].ToString();
                string serialNum = dr["SerialNumber"].ToString();
                string warrantName = dr["TempName"].ToString();

                string sqlTemp = "select top (1) WarrantName from (SELECT [WarrantName] FROM [EDIS].[dbo].[WarrantBasic] WHERE SUBSTRING(WarrantName,1,(len(WarrantName)-3))='" + warrantName.Substring(0, warrantName.Length - 1) + "' union ";
                sqlTemp += " SELECT [WarrantName] FROM [EDIS].[dbo].[ApplyTotalList] WHERE [ApplyKind]='1' AND [SerialNum]<" + serialNum + " AND SUBSTRING(WarrantName,1,(len(WarrantName)-3))='" + warrantName.Substring(0, warrantName.Length - 1) + "') as tb1 ";
                sqlTemp += " order by SUBSTRING(WarrantName,len(WarrantName)-1,len(WarrantName)) desc";

                System.Data.DataTable dvTemp = MSSQL.ExecSqlQry(sqlTemp, GlobalVar.loginSet.edisSqlConnString);
                int count = 0;
                if (dvTemp.Rows.Count > 0)
                {
                    string lastWarrantName = dvTemp.Rows[0][0].ToString();
                    if (!int.TryParse(lastWarrantName.Substring(lastWarrantName.Length - 2, 2), out count))
                        MessageBox.Show("parse failed " + lastWarrantName);
                }
                
                warrantName = warrantName + (count + 1).ToString("0#");
                h.SetParameterValue("@WarrantName", warrantName);
                h.SetParameterValue("@SerialNum", serialNum);
                h.ExecuteCommand();
            }
            h.Dispose();
            */
            string fileTSE = "D:\\權證發行_相關Excel\\上傳檔\\TSE申請上傳檔.txt";
            string fileOTC = "D:\\權證發行_相關Excel\\上傳檔\\OTC申請上傳檔.txt";

            //TXTFileWriter tseWriter = new TXTFileWriter(fileTSE);
            //TXTFileWriter otcWriter = new TXTFileWriter(fileOTC);           

            int tseCount = 0;
            int otcCount = 0;

            int tseReissue = 0;
            int otcReissue = 0;

            int tseReward = 0;
            int otcReward = 0;
            try {
                using (StreamWriter tseWriter = new StreamWriter(fileTSE, false, Encoding.GetEncoding("Big5")))
                using (StreamWriter otcWriter = new StreamWriter(fileOTC, false, Encoding.GetEncoding("Big5"))) {

                    string sql = @"SELECT a.ApplyKind
                                      ,a.Market
	                                  ,a.WarrantName
                                      ,a.UnderlyingID
                                      ,a.IssueNum
                                      ,a.CR
                                      ,IsNull(CASE WHEN a.ApplyKind='2' THEN c.T ELSE b.T END,6) T
                                      ,a.Type
                                      ,a.CP
                                      ,CASE WHEN a.UseReward='Y' THEN '1' ELSE '0' END UseReward
                                      ,CASE WHEN a.MarketTmr='Y' THEN '1' Else '0' END MarketTmr
                                  FROM [WarrantAssistant].[dbo].[ApplyTotalList] a
                                  LEFT JOIN [WarrantAssistant].[dbo].[ApplyOfficial] b ON a.SerialNum=b.SerialNumber
                                  LEFT JOIN [WarrantAssistant].[dbo].[WarrantBasic] c ON a.WarrantName=c.WarrantName
                                  ORDER BY a.Market desc, a.Type, a.CP, a.UnderlyingID, SUBSTRING(a.SerialNum,9,7),CONVERT(INT,SUBSTRING(a.SerialNum,18,LEN(a.SerialNum)-17))";//a.SerialNum
                    //DataView dv = DeriLib.Util.ExecSqlQry(sql, GlobalVar.loginSet.edisSqlConnString);
                    
                    System.Data.DataTable dv = MSSQL.ExecSqlQry(sql, GlobalVar.loginSet.warrantassistant45);
                    if (dv.Rows.Count > 0) {
                        foreach (DataRow dr in dv.Rows) {
                            string applyKind = dr["ApplyKind"].ToString();
                            string market = dr["Market"].ToString();
                            string warrantName = dr["WarrantName"].ToString();
                            string underlyingID = dr["UnderlyingID"].ToString();
                            double issueNum = Convert.ToDouble(dr["IssueNum"]);
                            double cr = Convert.ToDouble(dr["CR"]);
                            int t = Convert.ToInt32(dr["T"]);
                            string type = dr["Type"].ToString();
                            string cp = dr["CP"].ToString();
                            string useReward = dr["UseReward"].ToString();
                            string marketTmr = dr["MarketTmr"].ToString();
                            string markup = "                                     ";
                            int byteLen = System.Text.Encoding.Default.GetBytes(warrantName).Length;
                            warrantName = warrantName + markup.Substring(0, 16 - byteLen);
                            underlyingID = underlyingID.PadRight(12, ' ');
                            string issueNumS = issueNum.ToString().PadLeft(7, '0');
                            //issueNumS = issueNumS.PadLeft(7, '0');
                            string crS = (cr * 10000).ToString();
                            crS = crS.Substring(0, Math.Min(5, crS.Length));
                            crS = crS.PadLeft(5, '0');
                            string tS = t.ToString().PadLeft(2, '0');
                            //tS = tS.PadLeft(2, '0');
                            string tempType = "1";
                            if (type == "牛熊證") {
                                if (cp == "P")
                                    tempType = "4";
                                else
                                    tempType = "3";
                            } else {
                                if (cp == "P")
                                    tempType = "2";
                                else
                                    tempType = "1";
                            }
                                                       
                            string writestr = warrantName + underlyingID + issueNumS + crS + tS + tempType + useReward + marketTmr;
                            if (market == "TSE") {
                                //tseWriter.WriteFile(writestr);
                                tseWriter.WriteLine(writestr);
                                tseCount++;
                                if (useReward == "1")
                                    tseReward++;
                                if (applyKind == "2")
                                    tseReissue++;
                            } else if (market == "OTC") {
                                //otcWriter.WriteFile(writestr);
                                otcWriter.WriteLine(writestr);
                                otcCount++;
                                if (useReward == "1")
                                    otcReward++;
                                if (applyKind == "2")
                                    otcReissue++;
                            }
                        }
                    }
                }
                string pair = "";
                int nonmatch = 0;

                string sql2 = $@"SELECT  A.[SerialNumber], A.[TempName], B.[WarrantName]
                                  FROM [WarrantAssistant].[dbo].[ApplyOfficial] AS A
                                  LEFT JOIN [WarrantAssistant].[dbo].[ApplyTotalList] AS B ON A.[SerialNumber] =B.SerialNum";
                
                System.Data.DataTable dv2 = MSSQL.ExecSqlQry(sql2, GlobalVar.loginSet.warrantassistant45);

                foreach (DataRow dr in dv2.Rows)
                {
                    string serialNum = dr["SerialNumber"].ToString();
                    string tempName = dr["TempName"].ToString();
                    string warrantName = dr["WarrantName"].ToString();
                    string newWarrantName = "";

                    string sqlTemp = $@"SELECT top (1) WarrantName from (SELECT [WarrantName] FROM [WarrantAssistant].[dbo].[WarrantBasic] WHERE SUBSTRING(WarrantName,1,(len(WarrantName)-3))='{tempName.Substring(0, tempName.Length - 1)}' union 
                      SELECT [WarrantName] FROM [WarrantAssistant].[dbo].[ApplyTotalList] WHERE [ApplyKind]='1' AND CONVERT(INT, SUBSTRING([SerialNum], 18, LEN([SerialNum])-18 + 1)) <  CONVERT(INT, SUBSTRING('{serialNum}', 18, LEN('{serialNum}')-18 + 1)) AND SUBSTRING(WarrantName,1,(len(WarrantName)-3))='{tempName.Substring(0, tempName.Length - 1)}') as tb1 
                      order by SUBSTRING(WarrantName,len(WarrantName)-1,len(WarrantName)) desc";
                   
                    System.Data.DataTable dvTemp = MSSQL.ExecSqlQry(sqlTemp, GlobalVar.loginSet.warrantassistant45);

                    int count = 0;
                    if (dvTemp.Rows.Count > 0)
                    {
                        string lastWarrantName = dvTemp.Rows[0][0].ToString();
                        if (!int.TryParse(lastWarrantName.Substring(lastWarrantName.Length - 2, 2), out count))
                            MessageBox.Show("parse failed " + lastWarrantName);
                    }
                    newWarrantName = tempName + (count + 1).ToString("0#");
                    if (newWarrantName != warrantName)
                    {
                        pair += $"{warrantName} 應為 {newWarrantName}\n";
                        nonmatch++;
                    }
                }
                if (nonmatch > 0)
                    MessageBox.Show(pair);
                string infoStr = $"TSE 共{tseCount}檔，增額{tseReissue}檔，獎勵{tseReward}檔。\nOTC共{otcCount}檔，增額{otcReissue}檔，獎勵{otcReward}檔。";

                GlobalUtility.LogInfo("Info", $"今日共申請{(tseCount + otcCount)}檔權證發行/增額");

                MessageBox.Show("轉TXT檔完成!\n" + infoStr);

            } catch (Exception ex) {
                MessageBox.Show(ex.Message);
            }

        }

        private void 權證系統上傳檔ToolStripMenuItem_Click(object sender, EventArgs e) {
            string fileName = "D:\\權證發行_相關Excel\\上傳檔\\權證發行匯入檔.xls";
            Microsoft.Office.Interop.Excel.Application app = new Microsoft.Office.Interop.Excel.Application();
            Workbook workBook = null;

            try {
                string sql = @"SELECT a.UnderlyingID
	                                  ,c.TraderID
                                      ,a.WarrantName
                                      ,c.Type
                                      ,c.CP
                                      ,IsNull(b.MPrice,0) MPrice
                                      ,c.K
                                      ,c.ResetR
                                      ,c.BarrierR
                                      ,c.T
                                      ,a.CR
                                      ,c.HV
                                      ,c.IV
                                      ,a.IssueNum
                                      ,c.FinancialR
                                      ,a.UseReward
                                      ,c.Apply1500W
                                      ,c.SerialNumber
                                      ,ISNULL(c.說明,'') 說明
                                  FROM [WarrantAssistant].[dbo].[ApplyTotalList] a
                                  LEFT JOIN [WarrantAssistant].[dbo].[WarrantPrices] b ON a.UnderlyingID=b.CommodityID
                                  LEFT JOIN [WarrantAssistant].[dbo].[ApplyOfficial] c ON a.SerialNum=c.SerialNumber
                                  WHERE a.ApplyKind='1' AND a.Result+0.00001 >= a.EquivalentNum
                                  ORDER BY a.Market desc, a.Type, a.CP, a.UnderlyingID, SUBSTRING(a.SerialNum,9,7), CONVERT(INT,SUBSTRING(a.SerialNum,18,LEN(a.SerialNum)-17))"; //a.SerialNum
                System.Data.DataTable dv = MSSQL.ExecSqlQry(sql, GlobalVar.loginSet.warrantassistant45);
                
                //寫入EXCEL
                if (dv.Rows.Count > 0) {
                    int i = 3;
                    app.Visible = true;
                    workBook = app.Workbooks.Open(fileName);
                    //workBook.EnvelopeVisible = false;
                    Worksheet workSheet = (Worksheet) workBook.Sheets[1];
                    workSheet.get_Range("A3:BZ1000").ClearContents();
                    //workSheet.UsedRange.
                    bool havingReset = false;
                    foreach (DataRow dr in dv.Rows) {
                        string date = DateTime.Today.ToString("yyyyMMdd");
                        string underlyingID = dr["UnderlyingID"].ToString();
                        string traderID = dr["TraderID"].ToString().TrimStart('0');
                        //if (traderID == "10120")
                           // traderID = "7643";
                        //if (traderID == "10120" || traderID == "10329")
                        //    traderID = "7643";
                        string warrantName = dr["WarrantName"].ToString();
                        string type = dr["Type"].ToString();
                        string cp = dr["CP"].ToString();
                        string isReset = "N";
                        //if (type == "重設型" || type == "牛熊證")
                        if (type == "重設型" || type == "牛熊證" || type == "展延型")
                            isReset = "Y";
                        double stockPrice = Convert.ToDouble(dr["MPrice"]);
                        double k = Convert.ToDouble(dr["K"]);
                        double resetR = Convert.ToDouble(dr["ResetR"]);
                        double barrierR = Convert.ToDouble(dr["BarrierR"]);
                        if (isReset == "Y")
                        {
                            //if (type == "重設型")
                            if (type == "重設型"|| type == "展延型")
                            {
                                havingReset = true;
                                //k = Math.Round(resetR / 100 * stockPrice, 2);
                                if (cp == "C")
                                    k = Math.Round(1.5 * stockPrice, 2);
                                else
                                    k = Math.Round(0.5 * stockPrice, 2);
                            }
                            if(type == "牛熊證")
                            {
                                k = Math.Round(resetR / 100 * stockPrice, 2);
                            }
                        }
                            
                        double barrierP = Math.Round(barrierR / 100 * stockPrice, 2);
                        if (type == "牛熊證") {
                            if (cp == "C")
                            {
                                barrierP = Math.Round(Math.Floor(barrierR * stockPrice) / 100, 2);
                                //barrierP = Math.Round((barrierR * stockPrice) / 100, 2);
                            }
                            else if (cp == "P")
                                barrierP = Math.Round(Math.Ceiling(barrierR * stockPrice) / 100, 2);
                        }

                        //Check for moneyness constraint
                        if (type != "牛熊證") {
                            if ((cp == "C" && k / stockPrice >= 1.5) || (cp == "P" && k / stockPrice <= 0.5)) {
                                if (!havingReset)
                                {
                                    MessageBox.Show(warrantName + " 超過價外50%限制");
                                }
                                
                                // continue;
                            }
                        }

                        int t = Convert.ToInt32(dr["T"]);
                        double cr = Convert.ToDouble(dr["CR"]);
                        double r = GlobalVar.globalParameter.interestRate * 100;

                        //if(underlyingID == "IX0001" || underlyingID =="IX0027" || underlyingID == "IX0039" || underlyingID == "IX0118")
                        //    r = GlobalVar.globalParameter.interestRate_Index * 100;
                        if (underlyingID.Length > 4 && underlyingID.Substring(0,2) != "00")
                            r = GlobalVar.globalParameter.interestRate_Index * 100;
                        double hv = Convert.ToDouble(dr["HV"]);
                        double iv = Convert.ToDouble(dr["IV"]);
                        double initial_iv = iv;
                        string explain = dr["說明"].ToString();
                        double issueNum = Convert.ToDouble(dr["IssueNum"]);
                        double price = 0.0;
                        double financialR = Convert.ToDouble(dr["FinancialR"]);
                        string isReward = dr["UseReward"].ToString();

                        string is1500W = dr["Apply1500W"].ToString();
                        string serialNum = dr["SerialNumber"].ToString();
                        double p = 0.0;
                        double vol = iv / 100;
                        if (is1500W == "Y") {
                            CallPutType cpType = CallPutType.Call;
                            if (cp == "P")
                                cpType = CallPutType.Put;
                            if (underlyingID.Length > 4 && underlyingID.Substring(0, 2) != "00")
                            {
                                if (type == "牛熊證")
                                    p = Pricing.BullBearWarrantPrice(cpType, stockPrice, resetR, GlobalVar.globalParameter.interestRate_Index, vol, t, financialR, cr);
                                //else if (type == "重設型")
                                else if (type == "重設型" || type == "展延型")
                                    p = Pricing.ResetWarrantPrice(cpType, stockPrice, resetR, GlobalVar.globalParameter.interestRate_Index, vol, t, cr);
                                else
                                    p = Pricing.NormalWarrantPrice(cpType, stockPrice, k, GlobalVar.globalParameter.interestRate_Index, vol, t, cr);
                            }
                            else
                            {
                                if (type == "牛熊證")
                                    p = Pricing.BullBearWarrantPrice(cpType, stockPrice, resetR, GlobalVar.globalParameter.interestRate, vol, t, financialR, cr);
                                //else if (type == "重設型")
                                else if (type == "重設型" || type == "展延型")
                                    p = Pricing.ResetWarrantPrice(cpType, stockPrice, resetR, GlobalVar.globalParameter.interestRate, vol, t, cr);
                                else
                                    p = Pricing.NormalWarrantPrice(cpType, stockPrice, k, GlobalVar.globalParameter.interestRate, vol, t, cr);
                            }
                            double totalValue = p * issueNum * 1000;
                            double volUpperLimmit = vol * 2;
                            while (totalValue < 15000000 && vol < volUpperLimmit) {
                                vol += 0.01;
                                if (underlyingID.Length > 4 && underlyingID.Substring(0, 2) != "00")
                                {
                                    if (type == "牛熊證")
                                        p = Pricing.BullBearWarrantPrice(cpType, stockPrice, resetR, GlobalVar.globalParameter.interestRate_Index, vol, t, financialR, cr);
                                    //else if (type == "重設型")
                                    else if (type == "重設型" || type == "展延型")
                                        p = Pricing.ResetWarrantPrice(cpType, stockPrice, resetR, GlobalVar.globalParameter.interestRate_Index, vol, t, cr);
                                    else
                                        p = Pricing.NormalWarrantPrice(cpType, stockPrice, k, GlobalVar.globalParameter.interestRate_Index, vol, t, cr);
                                }
                                else
                                {
                                    if (type == "牛熊證")
                                        p = Pricing.BullBearWarrantPrice(cpType, stockPrice, resetR, GlobalVar.globalParameter.interestRate, vol, t, financialR, cr);
                                    //else if (type == "重設型")
                                    else if (type == "重設型" || type == "展延型")
                                        p = Pricing.ResetWarrantPrice(cpType, stockPrice, resetR, GlobalVar.globalParameter.interestRate, vol, t, cr);
                                    else
                                        p = Pricing.NormalWarrantPrice(cpType, stockPrice, k, GlobalVar.globalParameter.interestRate, vol, t, cr);
                                }
                                totalValue = p * issueNum * 1000;
                            }

                            if (vol < volUpperLimmit) {
                                iv = vol * 100;
                                string cmdText = "UPDATE [ApplyOfficial] SET IVNew=@IVNew WHERE SerialNumber=@SerialNumber";
                                List<System.Data.SqlClient.SqlParameter> pars = new List<System.Data.SqlClient.SqlParameter>();
                                pars.Add(new SqlParameter("@IVNew", SqlDbType.Float));
                                pars.Add(new SqlParameter("@SerialNumber", SqlDbType.VarChar));

                                SQLCommandHelper h = new SQLCommandHelper(GlobalVar.loginSet.warrantassistant45, cmdText, pars);

                                h.SetParameterValue("@IVNew", iv);
                                h.SetParameterValue("@SerialNumber", serialNum);
                                h.ExecuteCommand();
                                h.Dispose();
                            }
                        }

                        //if (type == "重設型")
                        if (type == "重設型" || type == "展延型")
                            type = "一般型";
                        if (cp == "P")
                            cp = "認售";
                        else
                            cp = "認購";
                        try {
                            // workSheet.Cells[1][i] = date;
                            
                            workSheet.get_Range("A" + i.ToString(), "A" + i.ToString()).Value = date;
                            workSheet.get_Range("B" + i.ToString(), "B" + i.ToString()).Value = underlyingID;
                            workSheet.get_Range("C" + i.ToString(), "C" + i.ToString()).Value = traderID;
                            workSheet.get_Range("D" + i.ToString(), "D" + i.ToString()).Value = warrantName;
                            workSheet.get_Range("E" + i.ToString(), "E" + i.ToString()).Value = type;
                            workSheet.get_Range("F" + i.ToString(), "F" + i.ToString()).Value = cp;
                            workSheet.get_Range("G" + i.ToString(), "G" + i.ToString()).Value = isReset;
                            workSheet.get_Range("H" + i.ToString(), "H" + i.ToString()).Value = stockPrice;
                            workSheet.get_Range("I" + i.ToString(), "I" + i.ToString()).Value = k;
                            workSheet.get_Range("J" + i.ToString(), "J" + i.ToString()).Value = resetR;
                            workSheet.get_Range("K" + i.ToString(), "K" + i.ToString()).Value = barrierP;
                            workSheet.get_Range("L" + i.ToString(), "L" + i.ToString()).Value = barrierR;
                            workSheet.get_Range("M" + i.ToString(), "M" + i.ToString()).Value = t;
                            workSheet.get_Range("N" + i.ToString(), "N" + i.ToString()).Value = cr;
                            workSheet.get_Range("O" + i.ToString(), "O" + i.ToString()).Value = r;
                            workSheet.get_Range("P" + i.ToString(), "P" + i.ToString()).Value = hv;
                            workSheet.get_Range("Q" + i.ToString(), "Q" + i.ToString()).Value = iv;
                            workSheet.get_Range("R" + i.ToString(), "R" + i.ToString()).Value = issueNum;
                            workSheet.get_Range("S" + i.ToString(), "S" + i.ToString()).Value = price;
                            workSheet.get_Range("T" + i.ToString(), "T" + i.ToString()).Value = financialR;
                            workSheet.get_Range("Y" + i.ToString(), "Y" + i.ToString()).Value = isReward;
                            
                            i++;
                        } catch (Exception ex) {
                            MessageBox.Show("write" + ex.Message);
                        }
                        try
                        {

                            string sql_InitialIV = $@"IF EXISTS(SELECT [WarrantName] FROM [WarrantAssistant].[dbo].[WarrantBasic_InitialIV] where [WarrantName] ='{warrantName}')
	                                                    UPDATE [WarrantAssistant].[dbo].[WarrantBasic_InitialIV] SET [InitialIV] ={initial_iv},[ApplyDate]='{date}',[說明]='{explain}',[財務費用率] = {financialR},[UID] = '{underlyingID}',[T] = {t},[利率] = {r},[重設比]={resetR} ,[行使比例] ={cr} , [CP] = '{cp}'  WHERE [WarrantName] ='{warrantName}'
                                                      ELSE 
                                                        INSERT [WarrantAssistant].[dbo].[WarrantBasic_InitialIV]
	                                                    VALUES ('{warrantName}',{initial_iv},'{date}','{explain}',{financialR},'{underlyingID}',{t},{r},{resetR},{cr},'{cp}')";
                            int bo = MSSQL.ExecSqlCmd(sql_InitialIV, GlobalVar.loginSet.warrantassistant45);
                        }
                        catch(Exception ex)
                        {

                        }
                        //存發行時的備註
                        
                    }

                    string sql2 = "SELECT [UnderlyingID] FROM [WarrantAssistant].[dbo].[ApplyOfficial] as A "
                                       + " left join (Select CS8010, count(1) as count from [VOLDB].[dbo].[ED_RelationUnderlying] "
                                                  + $" where RecordDate = (select top 1 RecordDate from [VOLDB].[dbo].[ED_RelationUnderlying])"
                                                   + " group by CS8010) as B on A.UnderlyingID = B.CS8010 "
                                       + " left join (SELECT stkid, MAX([IssueVol]) as MAX, min(IssueVol) as min FROM [10.101.10.5].[WMM3].[dbo].[Warrants]"
                                                  + " where kgiwrt = '他家' and marketdate <= GETDATE() and lasttradedate >= GETDATE() and IssueVol<> 0 "
                                                  + " group by stkid ) as C on A.UnderlyingID = C.stkid "
                                        + " WHERE B.count > 0 and (((IVNew > C.MAX or IVNew < C.min) and Apply1500W = 'Y') or ((IV > C.MAX or IV < C.min) and Apply1500W = 'N'))";
                    System.Data.DataTable badParam = MSSQL.ExecSqlQry(sql2, "Data Source=10.60.0.39;Initial Catalog=VOLDB;User ID=voldbuser;Password=voldbuser");

                    foreach (DataRow Row in badParam.Rows) {
                        //WindowState = FormWindowState.Minimized;
                        //Show();
                        //WindowState = FormWindowState.Normal;
                        Activate();
                        MessageBox.Show(Row["UnderlyingID"] + " 為關係人標的，波動度超過可發範圍，會被稽核該該叫，請修改條件。");
                    }
                    
                    GlobalUtility.LogInfo("Log", GlobalVar.globalParameter.userID + "產發行上傳檔");
                    app.Visible = false;
                    MessageBox.Show("發行上傳檔完成!");
                } else
                    MessageBox.Show("無可發行權證");

            } catch (Exception ex) {
                MessageBox.Show(ex.Message);
            } finally {
                if (workBook != null) {
                    workBook.Save();
                    workBook.Close();
                }
                if (app != null)
                    app.Quit();
            }
        }

        private void 增額上傳檔ToolStripMenuItem_Click(object sender, EventArgs e) {
            string fileName = "D:\\權證發行_相關Excel\\上傳檔\\增額作業匯入資料.xls";
            Microsoft.Office.Interop.Excel.Application app = new Microsoft.Office.Interop.Excel.Application();
            Workbook workBook = null;
            bool noPrice = false;

            try {

                string sql = $@"SELECT a.WarrantName
                                      ,a.IssueNum
                                      ,IsNull(c.MPrice,IsNull(c.BPrice,IsNull(c.APrice,0))) MPrice
                                      ,b.WarrantID
                                      ,ISNULL(d.WClosePrice, 0) AS WClosePrice
                                  FROM [WarrantAssistant].[dbo].[ApplyTotalList] a
                                  LEFT JOIN [WarrantAssistant].[dbo].[ReIssueOfficial] b ON a.SerialNum=b.SerialNum
                                  LEFT JOIN [WarrantAssistant].[dbo].[WarrantPrices] c ON b.WarrantID=c.CommodityID
                                  LEFT JOIN (SELECT  [WID],[WClosePrice] FROM [TwData].[dbo].[V_WarrantTrading] WHERE [TDate] = '{EDLib.TradeDate.LastNTradeDate(1).ToString("yyyyMMdd")}' and [TtoM] > 2) AS d on b.WarrantID = d.WID
                                  WHERE a.ApplyKind='2' AND a.Result + 0.00001 >=a.EquivalentNum
                                  ORDER BY a.Market desc, SUBSTRING(a.SerialNum,9,7) ,CONVERT(INT,SUBSTRING(a.SerialNum,18,LEN(a.SerialNum)-17))";
               
                System.Data.DataTable dv = MSSQL.ExecSqlQry(sql, GlobalVar.loginSet.warrantassistant45);

                if (dv.Rows.Count > 0) {
                    int i = 3;
                    app.Visible = true;
                    workBook = app.Workbooks.Open(fileName);
                    //workBook.EnvelopeVisible = false;
                    Worksheet workSheet = (Worksheet) workBook.Sheets[1];
                    workSheet.get_Range("A3:Z1000").ClearContents();

                    foreach (DataRow dr in dv.Rows) {
                        double warrantPrice = Convert.ToDouble(dr["MPrice"]);
                        double bidPrice = 0;
                        double askPrice = 0;
                        double refPrice = 0;
                        double wClosePrice = Convert.ToDouble(dr["WClosePrice"].ToString());
                        string wid = dr["WarrantID"].ToString();
                        /*
                        string sql_match_str = $@"SELECT [CommodityId], [MatchTotalQty],[BuyPrice], [SellPrice],[ReferencePrice]
                                                  FROM [TsQuote].[dbo].[HClose]
                                                  WHERE [CloseDate] = '{DateTime.Today.ToString("yyyyMMdd")}' AND [CommodityId] = '{wid}'";
                        */
                        string sql_match_str = $@"SELECT  [CommodityId],[MatchTotalQty],[BuyPriceBest1],[SellPriceBest1],[referenceprice]
                                                  FROM [TsQuote].[dbo].[vwprice2]
                                                  WHERE [tradedate] = '{DateTime.Today.ToString("yyyyMMdd")}' and [CommodityId] = '{wid}'";
                        System.Data.DataTable dt_match = MSSQL.ExecSqlQry(sql_match_str, GlobalVar.loginSet.tsquoteSqlConnString);
                        
                        if (dt_match.Rows.Count > 0)
                        {
                            double matchQty = Convert.ToDouble(dt_match.Rows[0][1]);
                            if (matchQty <= 0)
                            {
                                MessageBox.Show($@"{wid} 當日沒有成交");
                                //warrantPrice = Convert.ToDouble(dt_match.Rows[0][2]);
                                bidPrice = Convert.ToDouble(dt_match.Rows[0][2]);
                                askPrice = Convert.ToDouble(dt_match.Rows[0][3]);
                                refPrice = Convert.ToDouble(dt_match.Rows[0][4]);
                                if (refPrice < bidPrice)
                                    warrantPrice = bidPrice;
                                else if (refPrice > bidPrice && askPrice > refPrice)
                                    warrantPrice = refPrice;
                                else
                                    warrantPrice = askPrice;
                                
                            }
                        }
                        if (warrantPrice == 0.0)
                        {
                            noPrice = true;
                            warrantPrice = wClosePrice;
                        }

                        workSheet.get_Range("A" + i.ToString(), "A" + i.ToString()).Value = dr["WarrantName"].ToString();
                        workSheet.get_Range("B" + i.ToString(), "B" + i.ToString()).Value = "權證增額";
                        workSheet.get_Range("C" + i.ToString(), "C" + i.ToString()).Value = DateTime.Today.ToString("yyyyMMdd");
                        workSheet.get_Range("D" + i.ToString(), "D" + i.ToString()).Value = "增額發行";
                        workSheet.get_Range("E" + i.ToString(), "E" + i.ToString()).Value = Convert.ToDouble(dr["IssueNum"]) * 1000;
                        workSheet.get_Range("F" + i.ToString(), "F" + i.ToString()).Value = warrantPrice;
                        i++;
                    }
                    if (noPrice)
                        MessageBox.Show("注意!有權證價格為零!，代昨日收盤價");

                    GlobalUtility.LogInfo("Log", GlobalVar.globalParameter.userID + "產增額上傳檔");
                    app.Visible = false;
                    MessageBox.Show("增額上傳檔完成!");
                } else
                    MessageBox.Show("無可增額權證");

            } catch (Exception ex) {
                MessageBox.Show(ex.Message);
            } finally {
                if (workBook != null) {
                    workBook.Save();
                    workBook.Close();
                }
                if (app != null)
                    app.Quit();
            }
        }

        private void 關係人列表ToolStripMenuItem_Click(object sender, EventArgs e) {
            string fileName = "D:\\權證發行_相關Excel\\上傳檔\\利害關係人整批查詢上傳格式範例.xls";
            Microsoft.Office.Interop.Excel.Application app = new Microsoft.Office.Interop.Excel.Application();
            Workbook workBook = null;

            try {

                string sql = @"SELECT DISTINCT a.UnderlyingID
                                      ,b.UnifiedID
                                      ,b.FullName
                                  FROM [WarrantAssistant].[dbo].[ApplyTotalList] a
                                  LEFT JOIN [WarrantAssistant].[dbo].[WarrantUnderlying] b ON a.UnderlyingID=b.UnderlyingID
                                  WHERE a.Result>=a.EquivalentNum AND (b.StockType='DS' OR b.StockType='DR')";
                DataView dv = DeriLib.Util.ExecSqlQry(sql, GlobalVar.loginSet.warrantassistant45);

                int i = 2;
                if (dv.Count > 0) {
                    app.Visible = true;
                    workBook = app.Workbooks.Open(fileName);
                    //workBook.EnvelopeVisible = false;
                    Worksheet workSheet = (Worksheet) workBook.Sheets[1];
                    workSheet.get_Range("A3:Z1000").ClearContents();

                    foreach (DataRowView dr in dv) {
                        workSheet.get_Range("A" + i.ToString(), "A" + i.ToString()).Value = i - 1;
                        workSheet.get_Range("B" + i.ToString(), "B" + i.ToString()).Value = dr["UnifiedID"].ToString();
                        workSheet.get_Range("C" + i.ToString(), "C" + i.ToString()).Value = dr["FullName"].ToString();
                        i++;
                    }

                    GlobalUtility.LogInfo("Log", GlobalVar.globalParameter.userID + "產關係人上傳檔");
                    app.Visible = false;
                    MessageBox.Show("關係人上傳檔完成!");
                } else
                    MessageBox.Show("無關係人需查詢");

            } catch (Exception ex) {
                MessageBox.Show(ex.Message);
            } finally {
                if (workBook != null) {
                    workBook.Save();
                    workBook.Close();
                }
                if (app != null)
                    app.Quit();
            }
        }

        private void 修正權證名稱ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            //string sql5 = "SELECT [SerialNum], [WarrantName] FROM [EDIS].[dbo].[ApplyTotalList] WHERE [ApplyKind]='1'";
            string sql5 = "SELECT [SerialNumber], [TempName] FROM [WarrantAssistant].[dbo].[ApplyOfficial] ORDER BY CONVERT(INT, SUBSTRING([SerialNumber], 18, LEN([SerialNumber])-18 + 1))";
            System.Data.DataTable dv = MSSQL.ExecSqlQry(sql5, GlobalVar.loginSet.warrantassistant45);

            string cmdText = "UPDATE [ApplyTotalList] SET WarrantName=@WarrantName WHERE SerialNum=@SerialNum";
            List<System.Data.SqlClient.SqlParameter> pars = new List<System.Data.SqlClient.SqlParameter>();
            pars.Add(new SqlParameter("@WarrantName", SqlDbType.VarChar));
            pars.Add(new SqlParameter("@SerialNum", SqlDbType.VarChar));
            SQLCommandHelper h = new SQLCommandHelper(GlobalVar.loginSet.warrantassistant45, cmdText, pars);

            bool updated = false;
            foreach (DataRow dr in dv.Rows)
            {
                string serialNum = dr["SerialNumber"].ToString();
                string warrantName = dr["TempName"].ToString();

                string sqlTemp = $@"select top (1) WarrantName from (SELECT [WarrantName] FROM [WarrantAssistant].[dbo].[WarrantBasic] WHERE SUBSTRING(WarrantName,1,(len(WarrantName)-3))='{warrantName.Substring(0, warrantName.Length - 1)}' union 
                SELECT [WarrantName] FROM [WarrantAssistant].[dbo].[ApplyTotalList] WHERE [ApplyKind]='1' AND CONVERT(INT, SUBSTRING([SerialNum], 18 ,LEN([SerialNum])-18 + 1)) <  CONVERT(INT, SUBSTRING('{serialNum}', 18, LEN('{serialNum}')-18 + 1)) AND SUBSTRING(WarrantName,1,(len(WarrantName)-3))='{warrantName.Substring(0, warrantName.Length - 1)}') as tb1 
                order by SUBSTRING(WarrantName,len(WarrantName)-1,len(WarrantName)) desc";

                System.Data.DataTable dvTemp = MSSQL.ExecSqlQry(sqlTemp, GlobalVar.loginSet.warrantassistant45);
                int count = 0;
                if (dvTemp.Rows.Count > 0)
                {
                    string lastWarrantName = dvTemp.Rows[0][0].ToString();
                    if (!int.TryParse(lastWarrantName.Substring(lastWarrantName.Length - 2, 2), out count))
                        MessageBox.Show("parse failed " + lastWarrantName);
                }
                /*
                if (warrantName.Substring(warrantName.Length - 2, 2) != (count + 1).ToString("0#")) {
                    updated = true;
                    warrantName = warrantName.Substring(0, warrantName.Length - 2) + (count + 1).ToString("0#");
                    h.SetParameterValue("@WarrantName", warrantName);
                    h.SetParameterValue("@SerialNum", serialNum);
                    h.ExecuteCommand();
                }
                */
                warrantName = warrantName + (count + 1).ToString("0#");
                h.SetParameterValue("@WarrantName", warrantName);
                h.SetParameterValue("@SerialNum", serialNum);
                h.ExecuteCommand();
            }
            h.Dispose();
            /*
            if (updated)
                MessageBox.Show("Magic!");
            else
                MessageBox.Show("No magic.");
            */
            string sql2 = "UPDATE [WarrantAssistant].[dbo].[ApplyTotalList] SET Result=0";
            string sql3 = @"UPDATE [WarrantAssistant].[dbo].[ApplyTotalList] 
                                SET Result= CASE WHEN ApplyKind='1' THEN B.Result ELSE B.ReIssueResult END
                                FROM [WarrantAssistant].[dbo].[Apply_71] B
                                WHERE [ApplyTotalList].[WarrantName]=B.WarrantName";
            string sql4 = @"UPDATE [WarrantAssistant].[dbo].[ApplyTotalList]
                                SET Result= CASE WHEN [RewardCredit]>=[EquivalentNum] THEN [EquivalentNum] ELSE [RewardCredit] END
                               WHERE [UseReward]='Y'";

            conn.Open();
            MSSQL.ExecSqlCmd(sql2, conn);
            MSSQL.ExecSqlCmd(sql3, conn);
            MSSQL.ExecSqlCmd(sql4, conn);
            conn.Close();
            MessageBox.Show("已重新更新權證名稱");
        }

        private void ultraGrid1_InitializeRow(object sender, InitializeRowEventArgs e)
        {
            string content = e.Row.Cells["內容"].Value.ToString();
            if (content.Contains("改"))
            {
                e.Row.Appearance.ForeColor = Color.White;
                e.Row.Appearance.BackColor = Color.Teal;
            }
            if (content.Contains("刪除"))
            {
                e.Row.Appearance.ForeColor = Color.White;
                e.Row.Appearance.BackColor = Color.IndianRed;
            }
        }
        private void ultraGrid2_InitializeRow(object sender, InitializeRowEventArgs e)
        {
            string content = e.Row.Cells["內容"].Value.ToString();
          
            if (content.Contains("有誤"))
            {
                e.Row.Appearance.ForeColor = Color.White;
                e.Row.Appearance.BackColor = Color.IndianRed;
            }
        }
        private void toolStripComboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            LoadUltraGrid(dtAnnounce, "Announce");
        }
        //自動申請，一樣會產生發行檔，只是會自動透過selenium模擬瀏覽器上傳

        private bool Apply_first(string file,int TSEorOTC,string id,string key)
        {
            //TSEorOTC :0 上傳上市
            //TSEorOTC :1 上傳櫃買
            
            IWebDriver driver = new ChromeDriver("R:\\工具\\WarrantAssistant");
            if (TSEorOTC == 0)
                //driver.Navigate().GoToUrl("file:///C:/Users/0010327/Desktop/上傳發行檔新增.html");
                driver.Navigate().GoToUrl($"https://siis.twse.com.tw/server-java/t150sa02?step=0&id=9200pd{id}&TYPEK=sii&key={key}");
            //點新增
            else
            {//gotoOTC
             //driver.Navigate().GoToUrl($"https://siis.twse.com.tw/server-java/t150sa02?step=0&id=9200pd{id}&TYPEK=sii&key={key}");
            }
            //點新增
            IWebElement Addnew = driver.FindElement(By.XPath("//*[@id=\"form1\"]/input[7]"));
            Addnew.Click();
            //點選擇檔案
            IWebElement upFile = driver.FindElement(By.XPath("//*[@id=\"form1\"]/table/tbody/tr[3]/td/input[1]"));
            upFile.SendKeys(file);
            //點匯入
            IWebElement importFile = driver.FindElement(By.XPath("//*[@id=\"form1\"]/table/tbody/tr[3]/td/input[2]"));
            importFile.Click();
            //點確定匯入
            IWebElement send_confirm = driver.FindElement(By.XPath("//*[@id=\"content_d\"]/center/center/input[1]"));
            send_confirm.Click();
            //選取全部
            IWebElement clickall = driver.FindElement(By.XPath("//*[@id=\"sa02\"]/table/tbody/tr[1]/th[1]/input"));
            clickall.Click();
            //點送出
            IWebElement send = driver.FindElement(By.XPath("//*[@id=\"form1\"]/input[10]"));
            send.Click();
     
            try
            {
                //重覆匯入
                IWebElement backTolast = driver.FindElement(By.XPath("//*[@id=\"content_d\"]/center/h4/font"));
                return false;
            }
            catch
            {
                //回首頁
                IWebElement backToMain = driver.FindElement(By.XPath("//*[@id=\"content_d\"]/center/form/center/input"));
                backToMain.Click();
                driver.Quit();
                return true;
            }

            
            
            //IAlert alert = driver.SwitchTo().Alert();
            //alert.Accept();

            
        }

        //自動申請，一樣會產生發行檔，只是會自動透過selenium模擬瀏覽器上傳
        /*private void 自動申請發行TXTToolStripMenuItem_Click(object sender, EventArgs e)
        {
            MessageBox.Show("test");
        }*/
        
        private void 自動申請發行TXTToolStripMenuItem_Click(object sender, EventArgs e)
        { 

            string fileTSE = "D:\\權證發行_相關Excel\\上傳檔\\TSE申請上傳檔.txt";
            string fileOTC = "D:\\權證發行_相關Excel\\上傳檔\\OTC申請上傳檔.txt";

            //TXTFileWriter tseWriter = new TXTFileWriter(fileTSE);
            //TXTFileWriter otcWriter = new TXTFileWriter(fileOTC);           

            int tseCount = 0;
            int otcCount = 0;

            int tseReissue = 0;
            int otcReissue = 0;

            int tseReward = 0;
            int otcReward = 0;

            try
            {
                using (StreamWriter tseWriter = new StreamWriter(fileTSE, false, Encoding.GetEncoding("Big5")))
                using (StreamWriter otcWriter = new StreamWriter(fileOTC, false, Encoding.GetEncoding("Big5")))
                {

                    string sql = @"SELECT a.ApplyKind
                                      ,a.Market
	                                  ,a.WarrantName
                                      ,a.UnderlyingID
                                      ,a.IssueNum
                                      ,a.CR
                                      ,IsNull(CASE WHEN a.ApplyKind='2' THEN c.T ELSE b.T END,6) T
                                      ,a.Type
                                      ,a.CP
                                      ,CASE WHEN a.UseReward='Y' THEN '1' ELSE '0' END UseReward
                                      ,CASE WHEN a.MarketTmr='Y' THEN '1' Else '0' END MarketTmr
                                  FROM [WarrantAssistant].[dbo].[ApplyTotalList] a
                                  LEFT JOIN [WarrantAssistant].[dbo].[ApplyOfficial] b ON a.SerialNum=b.SerialNumber
                                  LEFT JOIN [WarrantAssistant].[dbo].[WarrantBasic] c ON a.WarrantName=c.WarrantName
                                  ORDER BY a.Market desc, a.Type, a.CP, a.UnderlyingID, SUBSTRING(a.SerialNum,9,7),CONVERT(INT,SUBSTRING(a.SerialNum,18,LEN(a.SerialNum)-17))";//a.SerialNum
                    //DataView dv = DeriLib.Util.ExecSqlQry(sql, GlobalVar.loginSet.edisSqlConnString);
                    System.Data.DataTable dv = MSSQL.ExecSqlQry(sql, GlobalVar.loginSet.warrantassistant45);

                    if (dv.Rows.Count > 0)
                    {
                        foreach (DataRow dr in dv.Rows)
                        {
                            string applyKind = dr["ApplyKind"].ToString();
                            string market = dr["Market"].ToString();
                            string warrantName = dr["WarrantName"].ToString();
                            string underlyingID = dr["UnderlyingID"].ToString();
                            double issueNum = Convert.ToDouble(dr["IssueNum"]);
                            double cr = Convert.ToDouble(dr["CR"]);
                            int t = Convert.ToInt32(dr["T"]);
                            string type = dr["Type"].ToString();
                            string cp = dr["CP"].ToString();
                            string useReward = dr["UseReward"].ToString();
                            string marketTmr = dr["MarketTmr"].ToString();

                            string markup = "                                     ";
                            int byteLen = System.Text.Encoding.Default.GetBytes(warrantName).Length;
                            warrantName = warrantName + markup.Substring(0, 16 - byteLen);
                            underlyingID = underlyingID.PadRight(12, ' ');
                            string issueNumS = issueNum.ToString().PadLeft(7, '0');
                            //issueNumS = issueNumS.PadLeft(7, '0');
                            string crS = (cr * 10000).ToString();
                            crS = crS.Substring(0, Math.Min(5, crS.Length));
                            crS = crS.PadLeft(5, '0');
                            string tS = t.ToString().PadLeft(2, '0');
                            //tS = tS.PadLeft(2, '0');

                            string tempType = "1";
                            if (type == "牛熊證")
                            {
                                if (cp == "P")
                                    tempType = "4";
                                else
                                    tempType = "3";
                            }
                            else
                            {
                                if (cp == "P")
                                    tempType = "2";
                                else
                                    tempType = "1";
                            }

                            string writestr = warrantName + underlyingID + issueNumS + crS + tS + tempType + useReward + marketTmr;

                            if (market == "TSE")
                            {
                                //tseWriter.WriteFile(writestr);
                                tseWriter.WriteLine(writestr);
                                tseCount++;
                                if (useReward == "1")
                                    tseReward++;
                                if (applyKind == "2")
                                    tseReissue++;
                            }
                            else if (market == "OTC")
                            {
                                //otcWriter.WriteFile(writestr);
                                otcWriter.WriteLine(writestr);
                                otcCount++;
                                if (useReward == "1")
                                    otcReward++;
                                if (applyKind == "2")
                                    otcReissue++;
                            }
                        }
                    }
                }
                string infoStr = $"TSE 共{tseCount}檔，增額{tseReissue}檔，獎勵{tseReward}檔。\nOTC共{otcCount}檔，增額{otcReissue}檔，獎勵{otcReward}檔。";
                GlobalUtility.LogInfo("Info", $"今日自動申請{(tseCount + otcCount)}檔權證發行/增額");
                string key = null;
                //Get key and id  
                try
                {
                    key = GlobalUtility.GetKey();
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"GetKey from BSSDB failed {ex}");
                    return;
                }
                string id = GlobalUtility.GetID();
                if (Apply_first(fileTSE, 0, id, key))
                {
                    MessageBox.Show("轉TXT檔完成AND自動申請完成!\n" + infoStr);
                }
                else
                {
                    MessageBox.Show("轉TXT檔完成，但有重覆權證，請手動檢查\n", "TopMostMessageBox", MessageBoxButtons.OK,
                    MessageBoxIcon.Information, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly);
                    
                }
                
            }
            catch(FileNotFoundException e1)
            {
                MessageBox.Show($"{e1.Message}   請在小幫手資料夾裡安裝chromedriver!");
            }
            catch (WebException)
            {
                MessageBox.Show("可能要更新Key，或是網頁有問題");
            }
            catch(Exception e2)
            {
                MessageBox.Show(e2.Message);
            }
            
        }

        private void 特殊標的ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            GlobalUtility.MenuItemClick<FrmSpecialUnderlying>();
        }

        private void 注意股票ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            //GlobalUtility.MenuItemClick<FrmSpecialUnderlying>();
            GlobalUtility.MenuItemClick<FrmWatchStock>();
        }


        private void 到期釋出額度ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            GlobalUtility.MenuItemClick<FrmExpireRelease>();
        }

        private void Manual_Click(object sender, EventArgs e)
        {
            GlobalUtility.MenuItemClick<FrmManual>();
        }

        private void 發行參數設定ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            GlobalUtility.MenuItemClick<FrmAutoSelect>();
        }

        

        private void 發行篩選ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            GlobalUtility.MenuItemClick<FrmAutoSelectResult>();
        }

        private void 發行矩陣設定ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            GlobalUtility.MenuItemClick<FrmAutoSelectMatrix>();
        }

        private void 原始參數查詢ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            GlobalUtility.MenuItemClick<FrmAutoSelectData>();
        }

        
        private void 註銷檔案匯入ToolStripMenuItem_Click(object sender, EventArgs e)
        {

            try
            {

                OpenFileDialog openFileDialog1 = new OpenFileDialog();
                openFileDialog1.Title = "打開excel文件";
                openFileDialog1.Filter = "Excel Files|*.csv;*.xls;*.xlsx;*.xlsm";
                openFileDialog1.InitialDirectory = @"D:\\";
                openFileDialog1.RestoreDirectory = true;
                
                if (openFileDialog1.ShowDialog() == DialogResult.OK)
                {

                    string path = openFileDialog1.FileName;
                    Microsoft.Office.Interop.Excel.Workbook book = null;
                    Microsoft.Office.Interop.Excel.Worksheet sheet = null;
                    Microsoft.Office.Interop.Excel.Range range = null;
                    Microsoft.Office.Interop.Excel.Application app = new Microsoft.Office.Interop.Excel.Application();

                    book = app.Workbooks.Open(path);
                    //sheet從1開始
                    sheet = book.Sheets["註銷清單"];
                    range = sheet.UsedRange;
                    int rs = range.Rows.Count;
                    MSSQL.ExecSqlCmd($@"DELETE FROM [WarrantAssistant].[dbo].[盤中即時註銷增額報表]
                                         WHERE[UpdateTime] > CONVERT(VARCHAR, GETDATE(), 112) AND [Type] = '註銷'", GlobalVar.loginSet.warrantassistant45);


                    string sql = @"INSERT INTO [WarrantAssistant].[dbo].[盤中即時註銷增額報表] ([UpdateTime],[WID],[Lots],[Type],[UserID],[說明]) "
                    + "VALUES(GETDATE(),@WID,@Lots,@Type,@UserID,@說明)";
                    List<SqlParameter> ps = new List<SqlParameter> {
                    new SqlParameter("@WID", SqlDbType.VarChar),
                    new SqlParameter("@Lots", SqlDbType.Float),
                    new SqlParameter("@Type", SqlDbType.VarChar),
                    new SqlParameter("@UserID", SqlDbType.VarChar),
                    new SqlParameter("@說明", SqlDbType.VarChar)};

                    SQLCommandHelper h = new SQLCommandHelper(GlobalVar.loginSet.warrantassistant45, sql, ps);

                    int i = 1;
                    int havingReset = 0;


                    for (int j = 2; j <= rs; j++)
                    {

                        string wid = Convert.ToString(sheet.get_Range("A" + j.ToString(), "A" + j.ToString()).Value);
                        //檔案是註銷股數，要轉成註銷張數
                        double lots = Convert.ToDouble(sheet.get_Range("B" + j.ToString(), "B" + j.ToString()).Value) / 1000;

                        string ex = Convert.ToString(sheet.get_Range("C" + j.ToString(), "C" + j.ToString()).Value);

                        h.SetParameterValue("@WID", wid);
                        h.SetParameterValue("@Lots", lots);
                        h.SetParameterValue("@Type", "註銷");
                        h.SetParameterValue("@UserID", GlobalVar.globalParameter.userID);
                        h.SetParameterValue("@說明", ex);
                        h.ExecuteCommand();
                        i++;
                    }
                    h.Dispose();
                    if (sheet != null)
                    {
                        Marshal.FinalReleaseComObject(range);
                        Marshal.FinalReleaseComObject(sheet);
                    }

                    if (book != null)
                    {
                        book.Close(false); //忽略尚未存檔內容，避免跳出提示卡住
                        Marshal.FinalReleaseComObject(book);
                    }
                    if (app != null)
                    {
                        app.Workbooks.Close();
                        app.Quit();
                        Marshal.FinalReleaseComObject(app);
                        app = null;
                        GC.Collect();
                    }
                    GlobalUtility.LogInfo("Log", GlobalVar.globalParameter.userID + " 註銷" + (i - 1) + "檔權證");
                    MessageBox.Show($@"成功匯入 {i - 1} 檔權證!");
                }
                
                

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            
        }
        
    }
}

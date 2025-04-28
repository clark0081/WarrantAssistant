#define To39
using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.Windows.Forms;
using Infragistics.Win;
using Infragistics.Win.UltraWinGrid;
using System.Data.SqlClient;
using EDLib.SQL;
using System.Threading;
using System.IO;
using System.Diagnostics;
using System.Configuration;

namespace WarrantAssistant
{
    
    public partial class FrmReIssue:Form
    {
        DateTime today = DateTime.Today;

        public SqlConnection conn = new SqlConnection(GlobalVar.loginSet.warrantassistant45);
        private DataTable dt = new DataTable();
        private bool isEdit = false;
        public string userID = GlobalVar.globalParameter.userID;
        public string userName = GlobalVar.globalParameter.userName;
        private int applyCount = 0;
        Dictionary<string, UidPutCallDeltaOne> UidDeltaOne = new Dictionary<string, UidPutCallDeltaOne>();//紀錄昨日的DeltaOne
        Dictionary<string, UidPutCallDeltaOne> UidDeltaOne_Temp = new Dictionary<string, UidPutCallDeltaOne>();//紀錄昨日DeltaOne加上擬發行的DeltaOne
        List<string> IsSpecial = new List<string>();//特殊標的
        List<string> IsIndex = new List<string>();//臺灣50指數,臺灣中型100指數,櫃買富櫃50指
        List<string> Market30 = new List<string>();//市值前30大
        public List<string> Iscompete = new List<string>();//要搶的標的
        private Thread thread;
        public double NonSpecialCallPutRatio = Convert.ToDouble(ConfigurationManager.AppSettings["NonSpecialCallPutRatio"].ToString());
        public double SpecialCallPutRatio = Convert.ToDouble(ConfigurationManager.AppSettings["SpecialCallPutRatio"].ToString());
        public double SpecialKGIALLPutRatio = Convert.ToDouble(ConfigurationManager.AppSettings["SpecialKGIALLPutRatio"].ToString());
        public double ISTOP30MaxIssue = Convert.ToDouble(ConfigurationManager.AppSettings["ISTOP30MaxIssue"].ToString());
        public double NonTOP30MaxIssue = Convert.ToDouble(ConfigurationManager.AppSettings["NonTOP30MaxIssue"].ToString());
        public FrmReIssue() {
            InitializeComponent();
            //thread = new Thread(new ThreadStart(Load_10min));
            //thread.Start();
        }
        private void Load_10min()
        {
            while (true)
            {
                Thread.Sleep(2000);
                LoadData();
                Thread.Sleep(600000);//10min
            }
        }

        private void FrmReIssue_Load(object sender, EventArgs e) {
            toolStripLabel1.Text = "使用者: " + userName;
            toolStripLabel2.Text = "";
            LoadData();
            InitialGrid();
        }

        private void InitialGrid() {
            
            /*dt.Columns.Add("WarrantID", typeof(string));
            dt.Columns.Add("增額張數", typeof(double));
            dt.Columns.Add("明日上市", typeof(string));
            dt.Columns.Add("確認", typeof(bool));
            dt.Columns["確認"].ReadOnly = false;
            dt.Columns.Add("獎勵", typeof(string));
            dt.Columns.Add("權證價格", typeof(double));
            dt.Columns.Add("標的代號", typeof(string));
            dt.Columns.Add("權證名稱", typeof(string));
            dt.Columns.Add("行使比例", typeof(double));
            dt.Columns.Add("約當張數", typeof(double));
            dt.Columns.Add("今日額度(%)", typeof(double));
            dt.Columns.Add("獎勵額度", typeof(double));
            dt.Columns.Add("交易員", typeof(string));
            ultraGrid1.DataSource = dt;*/

            UltraGridBand band0 = ultraGrid1.DisplayLayout.Bands[0];

            band0.Columns["TraderID"].Hidden = true;

            band0.Columns["ReIssueNum"].Format = "N0";
            band0.Columns["EquivalentNum"].Format = "N0";
            band0.Columns["RewardIssueCredit"].Format = "N0";

            band0.Columns["MarketTmr"].Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownList;
            band0.Columns["ConfirmChecked"].Style = Infragistics.Win.UltraWinGrid.ColumnStyle.CheckBox;

            band0.Columns["WarrantID"].Width = 100;
            band0.Columns["ReIssueNum"].Width = 100;
            band0.Columns["MarketTmr"].Width = 100;
            band0.Columns["ConfirmChecked"].Width = 70;
            band0.Columns["isReward"].Width = 70;
            band0.Columns["MPrice"].Width = 70;
            band0.Columns["UnderlyingID"].Width = 70;
            band0.Columns["WarrantName"].Width = 150;
            band0.Columns["exeRatio"].Width = 73;
            band0.Columns["說明"].Width = 200;


            for (int i = 4; i < 12; i++)
                band0.Columns[i].CellAppearance.BackColor = Color.LightGray;

            band0.Columns["MarketTmr"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center;
            band0.Override.HeaderAppearance.TextHAlign = Infragistics.Win.HAlign.Left;
            SetButton();
        }

        private void LoadData() {
            
            try {
                
                dt.Rows.Clear();
                UidDeltaOne.Clear();
                UidDeltaOne_Temp.Clear();
                IsSpecial.Clear();
                IsIndex.Clear();
                Market30.Clear();
                Iscompete.Clear();
                LoadIsIndex();
                //要搶的標的

                string sql_compete = $@"SELECT [UnderlyingID]
                          FROM [WarrantAssistant].[dbo].[Apply_71]
                          WHERE len([OriApplyTime])> 0";
                DataTable dt_compete = MSSQL.ExecSqlQry(sql_compete, GlobalVar.loginSet.warrantassistant45);
                foreach (DataRow dr in dt_compete.Rows)
                {
                    Iscompete.Add(dr["UnderlyingID"].ToString());
                }

                string sql_deltaone = $@"SELECT  [UnderlyingID], [KgiCallDeltaOne], [KgiPutDeltaOne], [KgiCallPutRatio]
                                        ,[AllCallDeltaOne], [AllPutDeltaOne], [KgiAllPutRatio], [YuanPutDeltaOne], [KgiPutNum]
                                        FROM [WarrantAssistant].[dbo].[WarrantIssueDeltaOne]
                                        WHERE [DateTime]='{DateTime.Today.ToString("yyyyMMdd")}'";
                DataTable dv_deltaone = MSSQL.ExecSqlQry(sql_deltaone, GlobalVar.loginSet.warrantassistant45);

                foreach (DataRow dr in dv_deltaone.Rows)
                {
                    string uid_ = dr["UnderlyingID"].ToString();
                    double Kgi_CallDeltaOne = Math.Round(Convert.ToDouble(dr["KgiCallDeltaOne"].ToString()), 2);
                    double Kgi_PutDeltaOne = Math.Round(Convert.ToDouble(dr["KgiPutDeltaOne"].ToString()), 2);
                    double Kgi_CallPutRatio = Convert.ToDouble(dr["KgiCallPutRatio"].ToString());
                    double All_CallDeltaOne = Math.Round(Convert.ToDouble(dr["AllCallDeltaOne"].ToString()), 2);
                    double All_PutDeltaOne = Math.Round(Convert.ToDouble(dr["AllPutDeltaOne"].ToString()), 2);
                    double KgiAll_PutRatio = Convert.ToDouble(dr["KgiAllPutRatio"].ToString());
                    double Yuan_PutDeltaOne = Convert.ToDouble(dr["YuanPutDeltaOne"].ToString());
                    int Kgi_PutNum = Convert.ToInt32(dr["KgiPutNum"].ToString());
                    if (!UidDeltaOne.ContainsKey(uid_))
                        UidDeltaOne.Add(uid_, new UidPutCallDeltaOne(uid_, Kgi_CallDeltaOne, Kgi_PutDeltaOne, Kgi_CallPutRatio, All_CallDeltaOne, All_PutDeltaOne, KgiAll_PutRatio, Yuan_PutDeltaOne, Kgi_PutNum));
                    if (!UidDeltaOne_Temp.ContainsKey(uid_))
                        UidDeltaOne_Temp.Add(uid_, new UidPutCallDeltaOne(uid_, Kgi_CallDeltaOne, Kgi_PutDeltaOne, Kgi_CallPutRatio, All_CallDeltaOne, All_PutDeltaOne, KgiAll_PutRatio, Yuan_PutDeltaOne, Kgi_PutNum));
                }

                string sql_reissue = $@"SELECT  A.ReIssueNum AS IssueNum, B.UnderlyingID AS UnderlyingID, IsNull(B.exeRatio,0) CR, B.WarrantName AS WarrantName ,C.RewardIssueCredit     
                              FROM [WarrantAssistant].[dbo].[ReIssueTempList] A
                              LEFT JOIN [WarrantAssistant].[dbo].[WarrantBasic] B ON A.WarrantID = B.WarrantID
                                LEFT JOIN (SELECT  [UID],IsNull(Floor(A.[WarrantAvailableShares] * 0.01 - IsNull(B.[UsedRewardNum],0)), 0) AS RewardIssueCredit  FROM [WarrantAssistant].[dbo].[WarrantUnderlyingCreditNew] AS A
                                                                              LEFT JOIN [WarrantAssistant].[dbo].[WarrantReward] AS B ON A.UID = B.UnderlyingID
                                                                              WHERE [UpdateTime] > CONVERT(VARCHAR,GETDATE(),112)) AS C ON B.UnderlyingID = C.UID
							  WHERE  A.ConfirmChecked='Y'";
                DataTable dv_reissue = MSSQL.ExecSqlQry(sql_reissue, GlobalVar.loginSet.warrantassistant45);

                foreach (DataRow dr in dv_reissue.Rows)
                {
                    string uid = dr["UnderlyingID"].ToString();
                    string wname = dr["WarrantName"].ToString();
                    double issuenum = Convert.ToDouble(dr["IssueNum"].ToString());
                    double cr = Convert.ToDouble(dr["CR"].ToString());
                    if (UidDeltaOne_Temp.ContainsKey(uid))
                    {
                        int wlength = wname.Length;
                        string subwname = wname.Substring(0, wlength - 2);

                        if (subwname.EndsWith("購") || subwname.EndsWith("牛"))//如果是Call  要考慮Put=0的情況 
                        {
                            UidDeltaOne_Temp[uid].KgiCallDeltaOne += issuenum * cr;
                            if (UidDeltaOne_Temp[uid].KgiPutNum == 0)
                                UidDeltaOne_Temp[uid].KgiCallPutRatio = 100;
                            else
                                UidDeltaOne_Temp[uid].KgiCallPutRatio = Math.Round((double)UidDeltaOne_Temp[uid].KgiCallDeltaOne / (double)UidDeltaOne_Temp[uid].KgiPutDeltaOne, 4);
                            UidDeltaOne_Temp[uid].AllCallDeltaOne += issuenum * cr;
                            if (UidDeltaOne_Temp[uid].AllPutDeltaOne == 0)
                                UidDeltaOne_Temp[uid].KgiAllPutRatio = 100;
                            else
                                UidDeltaOne_Temp[uid].KgiAllPutRatio = Math.Round((double)UidDeltaOne_Temp[uid].KgiPutDeltaOne / (double)UidDeltaOne_Temp[uid].AllPutDeltaOne, 4);
                        }
                        if (subwname.EndsWith("售") || subwname.EndsWith("熊"))
                        {
                            UidDeltaOne_Temp[uid].KgiPutDeltaOne += issuenum * cr;
                            UidDeltaOne_Temp[uid].KgiCallPutRatio = Math.Round((double)UidDeltaOne_Temp[uid].KgiCallDeltaOne / (double)UidDeltaOne_Temp[uid].KgiPutDeltaOne, 4);
                            UidDeltaOne_Temp[uid].AllPutDeltaOne += issuenum * cr;
                            UidDeltaOne_Temp[uid].KgiAllPutRatio = Math.Round((double)UidDeltaOne_Temp[uid].KgiPutDeltaOne / (double)UidDeltaOne_Temp[uid].AllPutDeltaOne, 4);
                            //UidDeltaOne_Temp[uid].KgiPutNum++;增額不用增加put
                        }
                    }
                }

                string sql_issue = $@"SELECT [SerialNum], [UnderlyingID], [R], [IssueNum], [CP], [ConfirmChecked]
                                    FROM [WarrantAssistant].[dbo].[ApplyTempList]
                                    WHERE [MDate] >='{today.ToString("yyyyMMdd")}' AND [ConfirmChecked] ='Y' AND LEN([UnderlyingID]) < 5";
                DataTable dv_issue = MSSQL.ExecSqlQry(sql_issue, GlobalVar.loginSet.warrantassistant45);

                foreach (DataRow dr in dv_issue.Rows)
                {
                    string serialNum = dr["SerialNum"].ToString();
                    string uid = dr["UnderlyingID"].ToString();
                    string cp = dr["CP"].ToString();
                    double issuenum = Convert.ToDouble(dr["IssueNum"].ToString());
                    double cr = Convert.ToDouble(dr["R"].ToString());
                    
                    if (UidDeltaOne_Temp.ContainsKey(uid))
                    {

                        bool deltaone = true;

                        string sqlTemp = $@"SELECT [ApplyTime], [OriApplyTime] FROM [WarrantAssistant].[dbo].[Apply_71] WHERE SerialNum = '{serialNum}'";
                        DataView dvTemp = DeriLib.Util.ExecSqlQry(sqlTemp, GlobalVar.loginSet.warrantassistant45);

                        string applyTime = "";
                        string apytime = "";
                        string oriapplyTime = "";
                        foreach (DataRowView drTemp in dvTemp)
                        {
                            applyTime = drTemp["ApplyTime"].ToString().Substring(0, 2);//時間幾點
                            apytime = drTemp["ApplyTime"].ToString();//時間的全部字串
                            oriapplyTime = drTemp["OriApplyTime"].ToString();
                        }
                        if (applyTime == "22" || (apytime.Length == 0) && Iscompete.Contains(uid))
                            deltaone = false;
                        
                        if (deltaone)
                        {
                           
                            if (cp == "C")//如果是Call  要考慮Put=0的情況 
                            {
                                
                                UidDeltaOne_Temp[uid].KgiCallDeltaOne += issuenum * cr;

                                if (UidDeltaOne_Temp[uid].KgiPutNum == 0)
                                    UidDeltaOne_Temp[uid].KgiCallPutRatio = 100;
                                else
                                    UidDeltaOne_Temp[uid].KgiCallPutRatio = Math.Round((double)UidDeltaOne_Temp[uid].KgiCallDeltaOne / (double)UidDeltaOne_Temp[uid].KgiPutDeltaOne, 4);
                               
                                UidDeltaOne_Temp[uid].AllCallDeltaOne += issuenum * cr;
                                if (UidDeltaOne_Temp[uid].AllPutDeltaOne == 0)
                                    UidDeltaOne_Temp[uid].KgiAllPutRatio = 100;
                                else
                                    UidDeltaOne_Temp[uid].KgiAllPutRatio = Math.Round((double)UidDeltaOne_Temp[uid].KgiPutDeltaOne / (double)UidDeltaOne_Temp[uid].AllPutDeltaOne, 4);
                               
                            }
                            if (cp == "P")
                            {
                               
                                UidDeltaOne_Temp[uid].KgiPutDeltaOne += issuenum * cr;
                                UidDeltaOne_Temp[uid].KgiCallPutRatio = Math.Round((double)UidDeltaOne_Temp[uid].KgiCallDeltaOne / (double)UidDeltaOne_Temp[uid].KgiPutDeltaOne, 4);
                                UidDeltaOne_Temp[uid].AllPutDeltaOne += issuenum * cr;
                                UidDeltaOne_Temp[uid].KgiAllPutRatio = Math.Round((double)UidDeltaOne_Temp[uid].KgiPutDeltaOne / (double)UidDeltaOne_Temp[uid].AllPutDeltaOne, 4);
                                UidDeltaOne_Temp[uid].KgiPutNum++;
                            }
                        }
                    }
                }


                string sql_special = $@"SELECT  DISTINCT  [UID]
                                      FROM [WarrantAssistant].[dbo].[SpecialStock]
                                      WHERE [Type] in ('B','R','KY','O')
                                      AND [DataDate] >=(select Max([Datadate]) FROM [WarrantAssistant].[dbo].[SpecialStock])";
                DataTable dv_special = MSSQL.ExecSqlQry(sql_special, GlobalVar.loginSet.warrantassistant45);

                foreach (DataRow dr in dv_special.Rows)
                {
                    string uid = dr["UID"].ToString();
                    if (!IsSpecial.Contains(uid))
                        IsSpecial.Add(uid);
                }

                string sql_M = $@"SELECT  DISTINCT  [UID]
                              FROM [WarrantAssistant].[dbo].[SpecialStock]
                              WHERE [Type] = 'M' AND [IsTop30] = 'Y'
                              AND [DataDate] >=(select Max([Datadate]) FROM [WarrantAssistant].[dbo].[SpecialStock])";
                DataTable dv_M = MSSQL.ExecSqlQry(sql_M, GlobalVar.loginSet.warrantassistant45);

                foreach (DataRow dr in dv_M.Rows)
                {
                    string uid = dr["UID"].ToString();
                    if (!Market30.Contains(uid))
                        Market30.Add(uid);
                }

                string sql = @"SELECT a.WarrantID
                                  ,a.ReIssueNum
                                  ,a.MarketTmr
                                  ,CASE WHEN a.ConfirmChecked='Y' THEN 1 ELSE 0 END ConfirmChecked                                  
                                  ,CASE WHEN b.isReward='1' THEN 'Y' ELSE 'N' END isReward
                                  ,IsNull(c.MPrice,ISNull(c.BPrice,IsNull(c.APrice,0))) MPrice
                                  ,b.UnderlyingID
                                  ,b.WarrantName
                                  ,IsNull(b.exeRatio,0) exeRatio
                                  ,(a.ReIssueNum*IsNull(b.exeRatio,0)) as EquivalentNum
                                  ,IsNull(d.IssuedPercent,0) IssuedPercent
                                  ,IsNull(d.RewardIssueCredit,0) RewardIssueCredit
                                  ,a.TraderID
                                    ,ROUND(([TwData].[dbo].[Gamma] (e.OptionPrice,e.OptionPrice,0.025,f.MMVol, 2 / 252.0) * e.[CR] * e.OptionPrice * e.OptionPrice / 100) * f.OI * 1000,0) AS 目前OI最大Gamma
                              FROM [WarrantAssistant].[dbo].[ReIssueTempList] a
                              LEFT JOIN [WarrantAssistant].[dbo].[WarrantBasic] b ON a.WarrantID=b.WarrantID
                              LEFT JOIN [WarrantAssistant].[dbo].[WarrantPrices] c ON a.WarrantID=c.CommodityID
                              LEFT JOIN [WarrantAssistant].[dbo].[WarrantUnderlyingSummary] d ON b.UnderlyingID=d.UnderlyingID 
							  LEFT JOIN (SELECT  [StockNo],[OptionPrice],[Scale] / 1000 AS CR FROM [TwCMData].[dbo].[RTD1001] WHERE [TDate] = CONVERT(VARCHAR,GETDATE(),112)) as e on A.WarrantID = e.StockNo
							  LEFT JOIN (SELECT [WID],[AccReleasingLots] * -1 AS OI,[MMVol] FROM [TwData].[dbo].[V_WarrantTrading] WHERE [TDate] = (SELECT MAX(TDate) FROM [TwData].[dbo].[WarrantFlow]) AND [IssuerName] = '9200') as f on a.WarrantID = f.WID ";     
                sql += "WHERE a.UserID='" + userID + "' ";
                sql += "ORDER BY a.MDate";

                dt = MSSQL.ExecSqlQry(sql, GlobalVar.loginSet.warrantassistant45);

                dt.Columns.Add("總張數", typeof(double));
                dt.Columns.Add("即時庫存", typeof(double));
                dt.Columns.Add("說明", typeof(string));
                dt.Columns.Add("昨日KgiCall DeltaOne", typeof(double));
                dt.Columns.Add("昨日KgiPut DeltaOne", typeof(double));
                dt.Columns.Add("今日KgiCall DeltaOne", typeof(double));
                dt.Columns.Add("今日KgiPut DeltaOne", typeof(double));
                dt.Columns.Add("自家 Call/Put DeltaOne比例", typeof(double));
                dt.Columns.Add("自家/市場Put DeltaOne比例>25%", typeof(string));
                dt.Columns.Add("Put大於元大", typeof(string));
                
                ultraGrid1.DataSource = dt;

                dt.Columns[0].Caption = "權證代號";
                dt.Columns[1].Caption = "增額張數";
                dt.Columns[2].Caption = "明日上市";
                dt.Columns[3].Caption = "確認";
                dt.Columns[4].Caption = "獎勵";
                dt.Columns[5].Caption = "權證價格";
                dt.Columns[6].Caption = "標的代號";
                dt.Columns[7].Caption = "權證名稱";
                dt.Columns[8].Caption = "行使比例";
                dt.Columns[9].Caption = "約當張數";
                dt.Columns[10].Caption = "今日額度(%)";
                dt.Columns[11].Caption = "獎勵額度";
                dt.Columns[12].Caption = "交易員";
                dt.Columns[12].Caption = "最大Gamma";


                foreach (DataRow row in dt.Rows) {
                    row["IssuedPercent"] = Math.Round((double) row["IssuedPercent"], 2);
                    row["RewardIssueCredit"] = Math.Floor((double) row["RewardIssueCredit"]);
                }
                /*DataView dv = DeriLib.Util.ExecSqlQry(sql, GlobalVar.loginSet.edisSqlConnString);

                foreach (DataRowView drv in dv) {
                    DataRow dr = dt.NewRow();

                    string warrantID = drv["WarrantID"].ToString();
                    dr["WarrantID"] = warrantID;
                    double reIssueNum = Convert.ToDouble(drv["ReIssueNum"]);
                    dr["增額張數"] = reIssueNum;
                    string marketTmr = drv["MarketTmr"].ToString();
                    dr["明日上市"] = marketTmr;
                    dr["確認"] = drv["ConfirmChecked"];
                    dr["獎勵"] = drv["isReward"].ToString();
                    double warrantPrice = 0.0;
                    warrantPrice = Convert.ToDouble(drv["MPrice"]);
                    dr["權證價格"] = warrantPrice;
                    dr["標的代號"] = drv["UnderlyingID"].ToString();
                    dr["權證名稱"] = drv["WarrantName"].ToString();
                    double cr = Convert.ToDouble(drv["exeRatio"]);
                    dr["exeRatio"] = cr;
                    dr["約當張數"] = Convert.ToDouble(drv["EquivalentNum"]);
                    dr["IssuedPercent"] = Math.Round(Convert.ToDouble(drv["IssuedPercent"]), 2);
                    double rewardCredit = (double) drv["RewardIssueCredit"];
                    rewardCredit = Math.Floor(rewardCredit);
                    dr["獎勵額度"] = rewardCredit;
                    dr["交易員"] = drv["TraderID"].ToString();

                    dt.Rows.Add(dr);
                }*/
        
            } catch (Exception ex) {
                MessageBox.Show(ex.Message);
                
            }
        }
        private void LoadIsIndex()
        {
#if !To39
            string sql_indexTW50 = $@"SELECT [UID] FROM [newEDIS].[dbo].[IndexUnderlying]
                                      WHERE [IndexType] = 'TW50 Index' 
                                      AND [DataDate] = (SELECT MAX(DataDate) FROM [newEDIS].[dbo].[IndexUnderlying]
                                      WHERE [IndexType] = 'TW50 Index')";
            string sql_indexTWMC = $@"SELECT [UID] FROM [newEDIS].[dbo].[IndexUnderlying]
                                      WHERE [IndexType] = 'TWMC Index' 
                                      AND [DataDate] = (SELECT MAX(DataDate) FROM [newEDIS].[dbo].[IndexUnderlying]
                                      WHERE [IndexType] = 'TWMC Index')";
            string sql_indexGTSM50 = $@"SELECT [UID] FROM [newEDIS].[dbo].[IndexUnderlying]
                                      WHERE [IndexType] = 'GTSM50 Index' 
                                      AND [DataDate] = (SELECT MAX(DataDate) FROM [newEDIS].[dbo].[IndexUnderlying]
                                      WHERE [IndexType] = 'GTSM50 Index')";
            DataTable dv_indexTW50 = MSSQL.ExecSqlQry(sql_indexTW50, GlobalVar.loginSet.newEDIS);
            DataTable dv_indexTWMC = MSSQL.ExecSqlQry(sql_indexTWMC, GlobalVar.loginSet.newEDIS);
            DataTable dv_indexGTSM50 = MSSQL.ExecSqlQry(sql_indexGTSM50, GlobalVar.loginSet.newEDIS);
#else
            string sql_indexTW50 = $@"SELECT [UID] FROM [TwData].[dbo].[IndexUnderlying]
                                      WHERE [IndexType] = 'TW50 Index' 
                                      AND [DataDate] = (SELECT MAX(DataDate) FROM [TwData].[dbo].[IndexUnderlying]
                                      WHERE [IndexType] = 'TW50 Index')";
            string sql_indexTWMC = $@"SELECT [UID] FROM [TwData].[dbo].[IndexUnderlying]
                                      WHERE [IndexType] = 'TWMC Index' 
                                      AND [DataDate] = (SELECT MAX(DataDate) FROM [TwData].[dbo].[IndexUnderlying]
                                      WHERE [IndexType] = 'TWMC Index')";
            string sql_indexGTSM50 = $@"SELECT [UID] FROM [TwData].[dbo].[IndexUnderlying]
                                      WHERE [IndexType] = 'GTSM50 Index' 
                                      AND [DataDate] = (SELECT MAX(DataDate) FROM [TwData].[dbo].[IndexUnderlying]
                                      WHERE [IndexType] = 'GTSM50 Index')";
            DataTable dv_indexTW50 = MSSQL.ExecSqlQry(sql_indexTW50, GlobalVar.loginSet.twData);
            DataTable dv_indexTWMC = MSSQL.ExecSqlQry(sql_indexTWMC, GlobalVar.loginSet.twData);
            DataTable dv_indexGTSM50 = MSSQL.ExecSqlQry(sql_indexGTSM50, GlobalVar.loginSet.twData);
#endif
            foreach (DataRow dr in dv_indexTW50.Rows)
            {
                string uid = dr["UID"].ToString();
                if (!IsIndex.Contains(uid))
                    IsIndex.Add(uid);
            }
            foreach (DataRow dr in dv_indexTWMC.Rows)
            {
                string uid = dr["UID"].ToString();
                if (!IsIndex.Contains(uid))
                    IsIndex.Add(uid);
            }
            foreach (DataRow dr in dv_indexGTSM50.Rows)
            {
                string uid = dr["UID"].ToString();
                if (!IsIndex.Contains(uid))
                    IsIndex.Add(uid);
            }
        }
        private void UpdateData() {
            try {
                MSSQL.ExecSqlCmd("DELETE FROM [ReIssueTempList] WHERE UserID='" + userID + "'", conn);

                string sql = @"INSERT INTO [ReIssueTempList] (SerialNum, WarrantID, ReIssueNum, MarketTmr, ConfirmChecked, TraderID, MDate, UserID) ";
                sql += "VALUES(@SerialNum, @WarrantID, @ReIssueNum, @MarketTmr, @ConfirmChecked, @TraderID, @MDate, @UserID)";
                List<SqlParameter> ps = new List<SqlParameter>();
                ps.Add(new SqlParameter("@SerialNum", SqlDbType.VarChar));
                ps.Add(new SqlParameter("@WarrantID", SqlDbType.VarChar));
                ps.Add(new SqlParameter("@ReIssueNum", SqlDbType.Float));
                ps.Add(new SqlParameter("@MarketTmr", SqlDbType.VarChar));
                ps.Add(new SqlParameter("@ConfirmChecked", SqlDbType.VarChar));
                ps.Add(new SqlParameter("@TraderID", SqlDbType.VarChar));
                ps.Add(new SqlParameter("@MDate", SqlDbType.DateTime));
                ps.Add(new SqlParameter("@UserID", SqlDbType.VarChar));

                SQLCommandHelper h = new SQLCommandHelper(GlobalVar.loginSet.warrantassistant45, sql, ps);

                int i = 1;
                applyCount = 0;
                foreach (UltraGridRow r in ultraGrid1.Rows) {
                    string serialNumber = DateTime.Today.ToString("yyyyMMdd") + userID + "02" + i.ToString("0#");
                    string warrantID = r.Cells["WarrantID"].Value.ToString();
                    double reIssueNum = Convert.ToDouble(r.Cells["ReIssueNum"].Value);
                    string marketTmr = r.Cells["MarketTmr"].Value == DBNull.Value ? "Y" : r.Cells["MarketTmr"].Value.ToString();
                    string traderID = r.Cells["TraderID"].Value == DBNull.Value ? userID : r.Cells["TraderID"].Value.ToString();
                    bool confirmed = false;
                    confirmed = r.Cells["ConfirmChecked"].Value == DBNull.Value ? false : Convert.ToBoolean(r.Cells["ConfirmChecked"].Value);
                    string confirmChecked = "N";
                    if (confirmed) {
                        confirmChecked = "Y";
                        applyCount++;
                    } else
                        confirmChecked = "N";

                    h.SetParameterValue("@SerialNum", serialNumber);
                    h.SetParameterValue("@WarrantID", warrantID);
                    h.SetParameterValue("@ReIssueNum", reIssueNum);
                    h.SetParameterValue("@MarketTmr", marketTmr);
                    h.SetParameterValue("@ConfirmChecked", confirmChecked);
                    h.SetParameterValue("@TraderID", traderID);
                    h.SetParameterValue("@MDate", DateTime.Now);
                    h.SetParameterValue("@UserID", userID);

                    h.ExecuteCommand();
                    i++;
                }

                h.Dispose();
                GlobalUtility.LogInfo("Log", GlobalVar.globalParameter.userID + " 編輯/更新" + (i - 1) + "檔增額");

            } catch (Exception ex) {
                MessageBox.Show(ex.Message);
            }
        }

        private void OfficiallyApply() {
            try {
                UpdateData();
                LoadData();
                bool dataOK = true;
                //計算獎勵額度有沒有重複
                Dictionary<string, double> accCredit = new Dictionary<string, double>();
                foreach (Infragistics.Win.UltraWinGrid.UltraGridRow dr in ultraGrid1.Rows)
                {
                    /*
                    dt.Columns[0].Caption = "權證代號";
                    dt.Columns[1].Caption = "增額張數";
                    dt.Columns[2].Caption = "明日上市";
                    dt.Columns[3].Caption = "確認";
                    dt.Columns[4].Caption = "獎勵";
                    dt.Columns[5].Caption = "權證價格";
                    dt.Columns[6].Caption = "標的代號";
                    dt.Columns[7].Caption = "權證名稱";
                    dt.Columns[8].Caption = "行使比例";
                    dt.Columns[9].Caption = "約當張數";
                    dt.Columns[10].Caption = "今日額度(%)";
                    dt.Columns[11].Caption = "獎勵額度";
                    dt.Columns[12].Caption = "交易員";
                    */

                    string underlyingID = dr.Cells["UnderlyingID"].Value.ToString();
                    string wname = dr.Cells["WarrantName"].Value.ToString();
                    //double spot = dr.Cells["股價"].Value == DBNull.Value ? 0.0 : Convert.ToDouble(dr.Cells["股價"].Value);
                    double cr = dr.Cells["exeRatio"].Value == DBNull.Value ? 0 : Convert.ToDouble(dr.Cells["exeRatio"].Value);
                    double shares = dr.Cells["ReIssueNum"].Value == DBNull.Value ? 10000 : Convert.ToDouble(dr.Cells["ReIssueNum"].Value);
                    bool confirmed = dr.Cells["ConfirmChecked"].Value == DBNull.Value ? false : Convert.ToBoolean(dr.Cells["ConfirmChecked"].Value);
                    string isReward = dr.Cells["isReward"].Value.ToString();
                    if (!accCredit.ContainsKey(underlyingID))
                        accCredit.Add(underlyingID, 0);
                    if (isReward == "Y" && confirmed)
                        accCredit[underlyingID] += cr * shares;

                    string sqlTemp = @"SELECT a.[UnderlyingName]
	                                      ,IsNull(IsNull(b.MPrice, IsNull(b.BPrice,b.APrice)),0) MPrice
                                          ,a.TraderID TraderID
                                      FROM [WarrantAssistant].[dbo].[WarrantUnderlying] a
                                      LEFT JOIN [WarrantAssistant].[dbo].[WarrantPrices] b ON a.UnderlyingID=b.CommodityID ";
                    sqlTemp += $"WHERE  CAST(UnderlyingID as varbinary(100)) = CAST('{underlyingID}' as varbinary(100))";

                    //double shares = dr.Cells["ReIssueNum"].Value == DBNull.Value ? 10000 : Convert.ToDouble(dr.Cells["ReIssueNum"].Value);
                    //band0.Columns["EquivalentNum"].Format = "N0";
                    //band0.Columns["RewardIssueCredit"].Format = "N0";



                    double spot = 0;
                    DataTable dvTemp = MSSQL.ExecSqlQry(sqlTemp, GlobalVar.loginSet.warrantassistant45);

                    foreach (DataRow drTemp in dvTemp.Rows) {
                        //underlyingName = drTemp["UnderlyingName"].ToString();
                        //traderID = drTemp["TraderID"].ToString().PadLeft(7, '0');
                        spot = Convert.ToDouble(drTemp["MPrice"]);
                    }
                    string sqlTemp2 = $@"SELECT [UnderlyingName], [WarningScore]
                                          FROM [WarrantAssistant].[dbo].[WarrantIssueCheck]
                                          WHERE [UnderlyingID] = '{underlyingID}' and [WarningScore] > 0 AND [IsQuaterUnderlying] = 'Y'";
                    DataTable dvTemp2 = MSSQL.ExecSqlQry(sqlTemp2, GlobalVar.loginSet.warrantassistant45);
                    if (dvTemp2.Rows.Count > 0)
                    {
                        string uname = dvTemp2.Rows[0][0].ToString();
                        MessageBox.Show($@"{underlyingID} {uname}為警示股");
                        dataOK = false;
                        continue;
                    }
                    if (UidDeltaOne_Temp.ContainsKey(underlyingID) && confirmed)
                    {
                        int wlength = wname.Length;
                        string subwname = wname.Substring(0, wlength - 2);
                        if (subwname.EndsWith("購") || subwname.EndsWith("牛"))
                        {
                            if (IsSpecial.Contains(underlyingID))
                            {
                                if (Market30.Contains(underlyingID))//市值前30  DeltaOne*股價<5億
                                {
                                    if (cr * shares * spot > ISTOP30MaxIssue)
                                    {
                                        MessageBox.Show($@"{underlyingID} 為風險標的且為市值前30大標的，DeltaOne市值已超過{(int)(ISTOP30MaxIssue / 100000)}億");
                                        dataOK = false;
                                        continue;
                                    }
                                }
                                else
                                {
                                    if (cr * shares * spot > NonTOP30MaxIssue)
                                    {
                                        MessageBox.Show($@"{underlyingID} 為風險標的 DeltaOne市值已超過{(int)(NonTOP30MaxIssue / 100000)}億");
                                        dataOK = false;
                                        continue;
                                    }
                                }
                            }
                        }
                        if (subwname.EndsWith("售") || subwname.EndsWith("熊"))
                        {
                            if (IsSpecial.Contains(underlyingID))
                            {
                                if (Market30.Contains(underlyingID))//市值前30  DeltaOne*股價<5億
                                {
                                    if (cr * shares * spot > ISTOP30MaxIssue)
                                    {
                                        MessageBox.Show($@"{underlyingID} 為風險標的且為市值前30大標的，DeltaOne市值已超過{(int)(ISTOP30MaxIssue / 100000)}億");
                                        dataOK = false;
                                        continue;
                                    }
                                }
                                else
                                {
                                    if (cr * shares * spot > NonTOP30MaxIssue)
                                    {
                                        MessageBox.Show($@"{underlyingID} 為風險標的 DeltaOne市值已超過{(int)(NonTOP30MaxIssue / 100000)}億");
                                        dataOK = false;
                                        continue;
                                    }
                                }
                            }
                            //考慮發Put的時候不能把今天要發的Call算進來
                            if(IsSpecial.Contains(underlyingID) && (double)UidDeltaOne_Temp[underlyingID].KgiCallDeltaOne / (double)UidDeltaOne_Temp[underlyingID].KgiPutDeltaOne < SpecialCallPutRatio)
                            {
                                MessageBox.Show($@"{underlyingID} 為風險標的，自家權證 Call/Put DeltaOne比例 < {SpecialCallPutRatio}");
                                dataOK = false;
                                continue;
                            }
                            else if ((double)UidDeltaOne_Temp[underlyingID].KgiCallDeltaOne / (double)UidDeltaOne_Temp[underlyingID].KgiPutDeltaOne < NonSpecialCallPutRatio)
                            {
                                if (!IsIndex.Contains(underlyingID))
                                {
                                    MessageBox.Show($"{underlyingID} 自家權證 Call/Put DeltaOne比例 < {NonSpecialCallPutRatio}");
                                    dataOK = false;
                                    continue;
                                }
                            }
                            if (IsSpecial.Contains(underlyingID) && UidDeltaOne_Temp[underlyingID].AllPutDeltaOne > 0 && UidDeltaOne_Temp[underlyingID].KgiAllPutRatio > SpecialKGIALLPutRatio)
                            {
                                //若之前這檔標的沒發過Put可以跳過，可是要考慮今天發超過一檔
                                if (UidDeltaOne[underlyingID].KgiPutNum > 0 || (UidDeltaOne[underlyingID].KgiPutNum == 0 && UidDeltaOne_Temp[underlyingID].KgiPutNum > 1))
                                {
                                    MessageBox.Show($"{underlyingID} 自家/市場 Put DeltaOne比例 > {SpecialKGIALLPutRatio}");
                                    dataOK = false;
                                    continue;
                                }
                            }
                            //特殊標的要Follow元大
                            /*
                            if (IsSpecial.Contains(underlyingID))
                            {
                                if (UidDeltaOne_Temp[underlyingID].KgiPutDeltaOne > UidDeltaOne_Temp[underlyingID].YuanPutDeltaOne)
                                {
                                    MessageBox.Show($"{underlyingID} 為風險標的，自家Put DeltaOne 超過元大");
                                    dataOK = false;
                                    continue;
                                }
                            }
                            */
                        }
                    }

                }
                string sqlReissueCredit = $@"SELECT [UID] ,IsNull(Floor(A.[WarrantAvailableShares] * 0.01 - IsNull(B.[UsedRewardNum], 0)), 0) AS RewardIssueCredit  FROM[WarrantAssistant].[dbo].[WarrantUnderlyingCreditNew] AS A
                                              LEFT JOIN[WarrantAssistant].[dbo].[WarrantReward] AS B ON A.UID = B.UnderlyingID
                                              WHERE[UpdateTime] > CONVERT(VARCHAR, GETDATE(), 112)";
                DataTable dtReissueCredit = MSSQL.ExecSqlQry(sqlReissueCredit, conn);
                foreach(string key in accCredit.Keys)
                {
                    DataRow[] accCredit_Select = dtReissueCredit.Select($@"UID = '{key}'");
                    if(accCredit_Select.Length > 0)
                    {
                        double acc = Convert.ToDouble(accCredit_Select[0][1].ToString());
                        if(accCredit[key] > acc)
                        {
                            MessageBox.Show($@"{key} 獎勵額度不夠");
                            dataOK = false;
                        }
                    }
                }
                

                if (!dataOK)
                    return;

                string sql1 = "DELETE FROM [WarrantAssistant].[dbo].[ReIssueOfficial] WHERE [UserID]='" + userID + "'";
                string sql2 = @"INSERT INTO [WarrantAssistant].[dbo].[ReIssueOfficial] ([SerialNum],[UnderlyingID],[WarrantID],[WarrantName],[exeRatio],[ReIssueNum],[UseReward],[MarketTmr],[TraderID],[MDate],UserID)
                                SELECT a.SerialNum, b.UnderlyingID, a.WarrantID, b.WarrantName, b.exeRatio, a.ReIssueNum, CASE WHEN b.isReward='1' THEN 'Y' ELSE 'N' END isReward, a.MarketTmr, a.TraderID, a.MDate, a.UserID
                                  FROM [WarrantAssistant].[dbo].[ReIssueTempList] a
                                  LEFT JOIN [WarrantAssistant].[dbo].[WarrantBasic] b ON a.WarrantID=b.WarrantID";
                sql2 += " WHERE a.[UserID]='" + userID + "' AND a.[ConfirmChecked]='Y'";
                string sql3 = "DELETE FROM [WarrantAssistant].[dbo].[ApplyTotalList] WHERE [UserID]='" + userID + "' AND [ApplyKind]='2'";
                string sql4 = @"INSERT INTO [WarrantAssistant].[dbo].[ApplyTotalList] ([ApplyKind],[SerialNum],[Market],[UnderlyingID],[WarrantName],[CR] ,[IssueNum],[EquivalentNum],[Credit],[RewardCredit],[Type],[CP],[UseReward],[MarketTmr],[TraderID],[MDate],UserID)
                                SELECT '2',a.SerialNum, b.Market, a.UnderlyingID, a.WarrantName, a.exeRatio, a.ReIssueNum, (a.exeRatio*a.ReIssueNum), b.IssueCredit, b.RewardIssueCredit, CASE WHEN SUBSTRING(c.WarrantType,1,2)='浮動' THEN '重設型' ELSE (CASE WHEN SUBSTRING(c.WarrantType,1,2)='重設' THEN '牛熊證' ELSE '一般型' END) END, CASE WHEN SUBSTRING(c.WarrantType,LEN(c.WarrantType)-3,4)='熊證認售' OR SUBSTRING(c.WarrantType,LEN(c.WarrantType)-3,4)='認售權證' THEN 'P' ELSE 'C' END, a.UseReward,a.MarketTmr, a.TraderID, GETDATE(), a.UserID
                                FROM [WarrantAssistant].[dbo].[ReIssueOfficial] a
                                LEFT JOIN [WarrantAssistant].[dbo].[WarrantUnderlyingSummary] b ON a.UnderlyingID=b.UnderlyingID
                                LEFT JOIN [WarrantAssistant].[dbo].[WarrantBasic] c ON a.WarrantID=c.WarrantID";
                sql4 += " WHERE a.[UserID]='" + userID + "'";
                //盤中即時增額註銷報表
                string sql5 = $@"DELETE FROM [WarrantAssistant].[dbo].[盤中即時註銷增額報表] WHERE [UpdateTime] > CONVERT(VARCHAR,GETDATE(),112) AND [Type] = '增額' AND [UserID]= '{userID}'";
                string sql6 = $@"INSERT INTO [WarrantAssistant].[dbo].[盤中即時註銷增額報表] ([UpdateTime],[WID],[Lots],[Type],[UserID],[說明])
                        SELECT GETDATE(),WarrantID, ReIssueNum, '增額', UserID,'' FROM [WarrantAssistant].[dbo].[ReIssueTempList]  WHERE [UserID]='{userID}' AND [ConfirmChecked]='Y'";

                conn.Open();
                MSSQL.ExecSqlCmd(sql1, conn);
                MSSQL.ExecSqlCmd(sql2, conn);
                MSSQL.ExecSqlCmd(sql3, conn);
                MSSQL.ExecSqlCmd(sql4, conn);
                MSSQL.ExecSqlCmd(sql5, conn);
                MSSQL.ExecSqlCmd(sql6, conn);
                conn.Close();

                toolStripLabel2.Text = DateTime.Now + "申請成功";
                GlobalUtility.LogInfo("Info", GlobalVar.globalParameter.userID + " 申請" + applyCount + "檔權證增額");
                MessageBox.Show("增額申請成功!");

            } catch (Exception ex) {
                MessageBox.Show(ex.Message);
            }
        }

        private void SetButton() {
            UltraGridBand band0 = ultraGrid1.DisplayLayout.Bands[0];
            if (isEdit) {
                band0.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.TemplateOnBottom;
                band0.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
                band0.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.True;

                band0.Columns["WarrantID"].CellActivation = Activation.AllowEdit;
                band0.Columns["ReIssueNum"].CellActivation = Activation.AllowEdit;
                band0.Columns["MarketTmr"].CellActivation = Activation.AllowEdit;
                band0.Columns["isReward"].CellActivation = Activation.AllowEdit;
                band0.Columns["MPrice"].CellActivation = Activation.AllowEdit;
                band0.Columns["WarrantName"].CellActivation = Activation.AllowEdit;
                band0.Columns["exeRatio"].CellActivation = Activation.AllowEdit;

                band0.Columns["昨日KgiCall DeltaOne"].CellActivation = Activation.NoEdit;
                band0.Columns["昨日KgiPut DeltaOne"].CellActivation = Activation.NoEdit;
                band0.Columns["今日KgiCall DeltaOne"].CellActivation = Activation.NoEdit;
                band0.Columns["今日KgiPut DeltaOne"].CellActivation = Activation.NoEdit;
                band0.Columns["自家 Call/Put DeltaOne比例"].CellActivation = Activation.NoEdit;
                band0.Columns["自家/市場Put DeltaOne比例>25%"].CellActivation = Activation.NoEdit;
                band0.Columns["Put大於元大"].CellActivation = Activation.NoEdit;

                buttonEdit.Visible = false;
                buttonConfirm.Visible = true;
                buttonDelete.Visible = true;
                buttonCancel.Visible = true;
                toolStripButton1.Visible = false;
                toolStripSeparator2.Visible = false;

                band0.Columns["ConfirmChecked"].Hidden = true;
                band0.Columns["UnderlyingID"].Hidden = true;
                band0.Columns["EquivalentNum"].Hidden = true;
                band0.Columns["IssuedPercent"].Hidden = true;
                band0.Columns["RewardIssueCredit"].Hidden = true;

            } else {
                band0.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
                band0.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
                band0.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;

                band0.Columns["ConfirmChecked"].CellActivation = Activation.AllowEdit;

                band0.Columns["WarrantID"].CellActivation = Activation.NoEdit;
                band0.Columns["ReIssueNum"].CellActivation = Activation.NoEdit;
                band0.Columns["MarketTmr"].CellActivation = Activation.NoEdit;
                band0.Columns["isReward"].CellActivation = Activation.NoEdit;
                band0.Columns["MPrice"].CellActivation = Activation.NoEdit;
                band0.Columns["UnderlyingID"].CellActivation = Activation.NoEdit;
                band0.Columns["WarrantName"].CellActivation = Activation.NoEdit;
                band0.Columns["exeRatio"].CellActivation = Activation.NoEdit;
                band0.Columns["EquivalentNum"].CellActivation = Activation.NoEdit;
                band0.Columns["IssuedPercent"].CellActivation = Activation.NoEdit;
                band0.Columns["RewardIssueCredit"].CellActivation = Activation.NoEdit;
                band0.Columns["昨日KgiCall DeltaOne"].CellActivation = Activation.NoEdit;
                band0.Columns["昨日KgiPut DeltaOne"].CellActivation = Activation.NoEdit;
                band0.Columns["今日KgiCall DeltaOne"].CellActivation = Activation.NoEdit;
                band0.Columns["今日KgiPut DeltaOne"].CellActivation = Activation.NoEdit;
                band0.Columns["自家 Call/Put DeltaOne比例"].CellActivation = Activation.NoEdit;
                band0.Columns["自家/市場Put DeltaOne比例>25%"].CellActivation = Activation.NoEdit;
                band0.Columns["Put大於元大"].CellActivation = Activation.NoEdit;
                buttonEdit.Visible = true;
                buttonConfirm.Visible = false;
                buttonDelete.Visible = false;
                buttonCancel.Visible = false;
                toolStripButton1.Visible = true;
                toolStripSeparator2.Visible = true;

                band0.Columns["ConfirmChecked"].Hidden = false;
                band0.Columns["UnderlyingID"].Hidden = false;
                band0.Columns["EquivalentNum"].Hidden = false;
                band0.Columns["IssuedPercent"].Hidden = false;
                band0.Columns["RewardIssueCredit"].Hidden = false;
            }
        }

        private void ultraGrid1_InitializeLayout(object sender, InitializeLayoutEventArgs e) {
            ultraGrid1.DisplayLayout.Override.RowSelectorHeaderStyle = RowSelectorHeaderStyle.ColumnChooserButton;

            ValueList v;
            if (!e.Layout.ValueLists.Exists("MyValueList")) {
                v = e.Layout.ValueLists.Add("MyValueList");
                v.ValueListItems.Add("Y", "Y");
                v.ValueListItems.Add("N", "N");
            }
            e.Layout.Bands[0].Columns["MarketTmr"].ValueList = e.Layout.ValueLists["MyValueList"];
        }

        private void buttonEdit_Click(object sender, EventArgs e) {
            isEdit = true;
            SetButton();
        }

        private void buttonConfirm_Click(object sender, EventArgs e) {
            ultraGrid1.PerformAction(Infragistics.Win.UltraWinGrid.UltraGridAction.ExitEditMode);
            isEdit = false;
            UpdateData();
            SetButton();
            LoadData();
        }

        private void buttonDelete_Click(object sender, EventArgs e) {
            isEdit = true;

            DialogResult result = MessageBox.Show("將全部刪除，確定?", "刪除資料", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
            if (result == DialogResult.Yes)
                MSSQL.ExecSqlCmd("DELETE FROM [ReIssueTempList] WHERE UserID='" + userID + "'", conn);

            LoadData();
            SetButton();
        }

        private void buttonCancel_Click(object sender, EventArgs e) {
            isEdit = false;
            LoadData();
            SetButton();
        }

        private void toolStripButton1_Click(object sender, EventArgs e) {
            if (GlobalVar.globalParameter.userGroup == "FE") {
                OfficiallyApply();
                LoadData();
            } else {
                if (DateTime.Now.TimeOfDay.TotalMinutes > 1200) { 
                //if (DateTime.Now.TimeOfDay.TotalMinutes > 750) { 
                    MessageBox.Show("超過交易所申報時間，欲改條件請洽行政組");
                /*
                else if (DateTime.Now.TimeOfDay.TotalMinutes > 570) {
                    DialogResult result = MessageBox.Show("超過約定的9:30了，已經告知組長及行政?", "逾時申請", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
                    if (result == DialogResult.Yes) {
                        OfficiallyApply();
                        LoadData();
                        GlobalUtility.LogInfo("Announce", GlobalVar.globalParameter.userID + " 逾時申請" + applyCount + "檔權證增額");

                    } else
                        LoadData();
                */
                } else {
                    OfficiallyApply();
                    LoadData();
                }
            }
        }

        private void ultraGrid1_CellChange(object sender, CellEventArgs e) {
            if (e.Cell.Column.Key == "ConfirmChecked")
                ultraGrid1.PerformAction(UltraGridAction.ExitEditMode);
        }

        private void ultraGrid1_AfterCellUpdate(object sender, CellEventArgs e) {
            if (e.Cell.Column.Key == "WarrantID") {
                string warrantID = e.Cell.Row.Cells["WarrantID"].Value.ToString();

                string sqlTemp = @"SELECT CASE WHEN a.isReward='1' THEN 'Y' ELSE 'N' END isReward
		                                ,IsNull(b.MPrice,ISNull(b.BPrice,IsNull(b.APrice,0))) MPrice
		                                ,a.WarrantName
		                                ,a.exeRatio
                                        ,a.TraderID
                                  FROM [WarrantAssistant].[dbo].[WarrantBasic] a
                                  LEFT JOIN [WarrantAssistant].[dbo].[WarrantPrices] b ON a.WarrantID=b.CommodityID";
                sqlTemp += " WHERE a.WarrantID='" + warrantID + "'";
                DataTable dtTemp = MSSQL.ExecSqlQry(sqlTemp, GlobalVar.loginSet.warrantassistant45);

                if (dtTemp.Rows.Count == 0) {
                    MessageBox.Show("Wrong WarrantID!");
                } else {
                    e.Cell.Row.Cells["isReward"].Value = dtTemp.Rows[0]["isReward"];
                    e.Cell.Row.Cells["MPrice"].Value = dtTemp.Rows[0]["MPrice"];
                    e.Cell.Row.Cells["WarrantName"].Value = dtTemp.Rows[0]["WarrantName"];
                    e.Cell.Row.Cells["exeRatio"].Value = dtTemp.Rows[0]["exeRatio"];
                    e.Cell.Row.Cells["TraderID"].Value = dtTemp.Rows[0]["TraderID"];
                }
            }
        }

        private void ultraGrid1_MouseDown(object sender, MouseEventArgs e) {
            if (e.Button == MouseButtons.Right) {
                contextMenuStrip1.Show();
            }
        }

        private void ultraGrid1_BeforeRowsDeleted(object sender, BeforeRowsDeletedEventArgs e) {
            e.DisplayPromptMsg = false;
        }

        private void 刪除ToolStripMenuItem_Click(object sender, EventArgs e) {
            DialogResult result = MessageBox.Show("刪除此檔，確定?", "刪除資料", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
            if (result == DialogResult.Yes) {
                ultraGrid1.ActiveRow.Delete();
                UpdateData();
            }
            LoadData();
        }

        private void ultraGrid1_InitializeRow(object sender, InitializeRowEventArgs e) {
            string warrantID = "";
            string underlyingID = "";
            string warrantName = "";
            warrantID = e.Row.Cells["WarrantID"].Value.ToString();
            underlyingID = e.Row.Cells["UnderlyingID"].Value.ToString();
            warrantName = e.Row.Cells["WarrantName"].Value.ToString();

            string traderID = "NA";
            string issuable = "NA";
            string reissuable = "NA";

            string toolTip1 = "今日未達增額標準";
            string toolTip2 = "非本季標的";
            string toolTip3 = "標的發行檢查=N";
            string toolTip4 = "非此使用者所屬標的";

            string sqlTemp = "SELECT TraderID, IsNull(Issuable,'NA') Issuable FROM [WarrantAssistant].[dbo].[WarrantUnderlyingSummary] WHERE UnderlyingID = '" + underlyingID + "'";
            DataView dvTemp = DeriLib.Util.ExecSqlQry(sqlTemp, GlobalVar.loginSet.warrantassistant45);

            if (dvTemp.Count > 0) {
                foreach (DataRowView drTemp in dvTemp) {
                    traderID = drTemp["TraderID"].ToString().PadLeft(7, '0');
                    issuable = drTemp["Issuable"].ToString();
                }
            }

            string sqlTemp2 = "SELECT IsNull([ReIssuable],'NA') ReIssuable FROM [WarrantAssistant].[dbo].[WarrantReIssuable] WHERE WarrantID = '" + warrantID + "'";
            DataView dvTemp2 = DeriLib.Util.ExecSqlQry(sqlTemp2, GlobalVar.loginSet.warrantassistant45);

            if (dvTemp2.Count > 0) {
                foreach (DataRowView drTemp2 in dvTemp2) {
                    reissuable = drTemp2["ReIssuable"].ToString();
                }
            }

            string sqlTemp5 = $@"SELECT  [Symbol],SUM([Qty]/1000) AS SellQTY FROM [10.101.10.5].[WMM3].[dbo].[TradeFeed]
							  WHERE [Date] = CONVERT(VARCHAR,GETDATE(),112) and [Source] = 'FIXGW85' and [Side] = '2' AND [Symbol] = '{warrantID}' GROUP BY [Symbol]";
            DataTable dvTemp5 = EDLib.SQL.MSSQL.ExecSqlQry(sqlTemp5, GlobalVar.loginSet.twData);
            double sellQty = 0;
            if (dvTemp5.Rows.Count > 0)
            {
                sellQty = Convert.ToDouble(dvTemp5.Rows[0][1].ToString());
            }
            double buyQty = 0;
            string sqlTemp6 = $@"SELECT  [Symbol],SUM([Qty]/1000) AS BuyQTY FROM [10.101.10.5].[WMM3].[dbo].[TradeFeed]
							  WHERE [Date] = CONVERT(VARCHAR,GETDATE(),112) and [Source] = 'FIXGW85' and [Side] = '1' AND [Symbol] = '{warrantID}' GROUP BY [Symbol]";
            DataTable dvTemp6 = EDLib.SQL.MSSQL.ExecSqlQry(sqlTemp6, GlobalVar.loginSet.twData);
            if (dvTemp6.Rows.Count > 0)
            {
                buyQty = Convert.ToDouble(dvTemp6.Rows[0][1].ToString());
            }
            double newFL = 0;
            string sqlTemp7 = "SELECT [WName],[DelLastNum] * 1000 AS FLnew FROM [10.101.10.5].[WMM3].[dbo].[Warrants] WHERE [WName] = '" + warrantName + "'";
            DataTable dvTemp7 = EDLib.SQL.MSSQL.ExecSqlQry(sqlTemp7, GlobalVar.loginSet.twData);

            if (dvTemp7.Rows.Count > 0)
            {
                newFL = Convert.ToDouble(dvTemp7.Rows[0][1].ToString());
            }

            double acc = 0;
            string sqlTemp8 = $@"SELECT  [WID],[AccReleasingLots] FROM [TwData].[dbo].[V_WarrantTrading] WHERE [TDate] = '{TradeDate.LastNTradeDate(1)}' AND [WID] = '{warrantID}'";
            DataTable dvTemp8 = EDLib.SQL.MSSQL.ExecSqlQry(sqlTemp8, GlobalVar.loginSet.twData);

            if (dvTemp8.Rows.Count > 0)
            {
                acc = Convert.ToDouble(dvTemp8.Rows[0][1].ToString());
            }



            e.Row.Cells["總張數"].Value = newFL;
            e.Row.Cells["即時庫存"].Value = newFL + acc - sellQty + buyQty;

            if (!isEdit) {

                if (issuable == "NA") {
                    e.Row.ToolTipText = toolTip2;
                    e.Row.Appearance.ForeColor = Color.Red;
                } else if (issuable == "N") {
                    e.Row.Cells["UnderlyingID"].ToolTipText = toolTip3;
                    e.Row.Cells["UnderlyingID"].Appearance.ForeColor = Color.Red;
                }

                if (reissuable == "NA") {
                    e.Row.Cells["WarrantID"].ToolTipText = toolTip1;
                    e.Row.Cells["WarrantID"].Appearance.ForeColor = Color.Red;
                }

                if (issuable != "NA" && traderID != userID) {
                    e.Row.Appearance.BackColor = Color.LightYellow;
                    e.Row.Cells["WarrantID"].ToolTipText = toolTip4;
                }

                //string sqlTemp3 = $@"SELECT  CONVERT(VARCHAR,[日期],112) AS 日期,[註銷張數] FROM [WarrantAssistant].[dbo].[權證註銷報表] WHERE [權證代號] = '{warrantID}' AND [日期] >= '{TradeDate.LastNTradeDateDT(10).ToString("yyyyMMdd")}' ORDER BY [日期] DESC";
                //DataView dvTemp3 = DeriLib.Util.ExecSqlQry(sqlTemp3, GlobalVar.loginSet.warrantassistant45);
                string log = "";
                //有通報公會的權證，若增額要解除通報
                string sqlDeviation = $@"SELECT [WID] FROM [TwData].[dbo].[WPriceDeviation] WHERE [RemoveDate] IS NULL AND [UID] = '{underlyingID}' AND [WID] = '{warrantID}'";
                DataView dvDeviation = DeriLib.Util.ExecSqlQry(sqlDeviation, GlobalVar.loginSet.twData);
                if(dvDeviation.Count > 0)
                {
                    log += "已通報公會 ";
                }
                /*
                if (dvTemp3.Count > 0)
                {
                    foreach (DataRowView drTemp3 in dvTemp3)
                    {
                        log += "(" + drTemp3["日期"].ToString()+" 註銷"+drTemp3["註銷張數"].ToString() +"張)";
                    }
                }
                */
                //string sqlTemp4 = $@"SELECT  CONVERT(VARCHAR,[TDate],112) AS 日期,[ReIssueLots] FROM [WarrantAssistant].[dbo].[WarrantReIssueLog] WHERE [WID] = '{warrantID}' AND [TDate] >= '{TradeDate.LastNTradeDateDT(10).ToString("yyyyMMdd")}' ORDER BY [TDate] DESC";
                //DataView dvTemp4 = DeriLib.Util.ExecSqlQry(sqlTemp4, GlobalVar.loginSet.warrantassistant45);
                /*
                if (dvTemp4.Count > 0)
                {
                    foreach (DataRowView drTemp4 in dvTemp4)
                    {
                        log += "(" + drTemp4["日期"].ToString() + " 增額" + drTemp4["ReIssueLots"].ToString() + "張)";
                    }
                }
                */
                e.Row.Cells["說明"].Value = log;
            }

            if (UidDeltaOne_Temp.ContainsKey(underlyingID))
            {
                e.Row.Cells["昨日KgiCall DeltaOne"].Value = UidDeltaOne[underlyingID].KgiCallDeltaOne;
                e.Row.Cells["昨日KgiPut DeltaOne"].Value = UidDeltaOne[underlyingID].KgiPutDeltaOne;
                e.Row.Cells["今日KgiCall DeltaOne"].Value = UidDeltaOne_Temp[underlyingID].KgiCallDeltaOne;
                e.Row.Cells["今日KgiPut DeltaOne"].Value = UidDeltaOne_Temp[underlyingID].KgiPutDeltaOne;
                e.Row.Cells["自家 Call/Put DeltaOne比例"].Value = UidDeltaOne_Temp[underlyingID].KgiCallPutRatio;
                if (IsSpecial.Contains(underlyingID) && UidDeltaOne_Temp[underlyingID].KgiCallPutRatio < SpecialCallPutRatio)
                    e.Row.Cells["自家 Call/Put DeltaOne比例"].Appearance.ForeColor = Color.Red;
                else if (UidDeltaOne_Temp[underlyingID].KgiCallPutRatio < NonSpecialCallPutRatio)
                {
                    if(!IsIndex.Contains(underlyingID))
                        e.Row.Cells["自家 Call/Put DeltaOne比例"].Appearance.ForeColor = Color.Red;
                }
                if (IsSpecial.Contains(underlyingID) && UidDeltaOne_Temp[underlyingID].KgiAllPutRatio > SpecialKGIALLPutRatio && UidDeltaOne_Temp[underlyingID].KgiPutNum > 0)
                {
                    e.Row.Cells["自家/市場Put DeltaOne比例>25%"].Value = "Y";
                    e.Row.Cells["自家/市場Put DeltaOne比例>25%"].Appearance.ForeColor = Color.Red;
                }
                else
                {
                    e.Row.Cells["自家/市場Put DeltaOne比例>25%"].Value = "N";
                    e.Row.Cells["自家/市場Put DeltaOne比例>25%"].Appearance.ForeColor = Color.Black;
                }
                if (UidDeltaOne_Temp[underlyingID].KgiPutDeltaOne > UidDeltaOne_Temp[underlyingID].YuanPutDeltaOne)
                {
                    e.Row.Cells["Put大於元大"].Value = "Y";
                    e.Row.Cells["Put大於元大"].Appearance.ForeColor = Color.Red;
                }
                else
                {
                    e.Row.Cells["Put大於元大"].Value = "N";
                    e.Row.Cells["Put大於元大"].Appearance.ForeColor = Color.Black;
                }
            }
        }

        private void ultraGrid1_DoubleClickCell(object sender, DoubleClickCellEventArgs e) {
            if (e.Cell.Column.Key == "WarrantID")
                GlobalUtility.MenuItemClick<FrmReIssuable>();
        }

        private void ultraGrid1_DoubleClickHeader(object sender, DoubleClickHeaderEventArgs e) {
            if (e.Header.Column.Key == "ConfirmChecked") {
                foreach (Infragistics.Win.UltraWinGrid.UltraGridRow r in ultraGrid1.Rows) {
                    r.Cells["ConfirmChecked"].Value = true;
                    UpdateData();
                    LoadData();
                }
            }
        }

        private void FrmReIssue_FormClosed(object sender, FormClosedEventArgs e) {
            UpdateData();
        }
        private void FrmReIssue_FormClosing(object sender, FormClosingEventArgs e)
        {
            //UpdateData();
            if (thread != null && thread.IsAlive)
                thread.Abort();
            //MessageBox.Show($"stop");
        }
    }
}

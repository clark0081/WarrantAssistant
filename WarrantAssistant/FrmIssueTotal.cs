#define To39
using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.Windows.Forms;
using Infragistics.Win.UltraWinGrid;
using System.Data.SqlClient;
using EDLib.SQL;
using System.IO;
using System.Threading.Tasks;
using System.Configuration;
using System.Threading;
namespace WarrantAssistant
{
    public partial class FrmIssueTotal:Form
    {
        private DataTable dt = new DataTable();
        private DataTable back_dt = new DataTable();//UpdateRecord()時用來比對更改資料
        private DataTable dt_LongTerm = new DataTable();//Vol穿需要發長天期，跳出提示
        private Dictionary<string,Dictionary<string,string>> dt_record = new Dictionary<string, Dictionary<string, string>>();
        private Dictionary<string, string> dt_recordByTrader = new Dictionary<string, string>();//顯示上一筆的更改紀錄
        private DataTable updaterecord_dt = new DataTable();//當有更改資料時，用來記錄那些欄位更改
        private Dictionary<string, double> UidMaxCR_C = new Dictionary<string, double>();
        private Dictionary<string, double> UidMaxCR_P = new Dictionary<string, double>();
        private List<string> ColumnName = new List<string>();
        Dictionary<string, string> sqlTogrid = new Dictionary<string, string>();
        private string userID = GlobalVar.globalParameter.userID;
        private bool isEdit = false;

        DateTime lastDate = EDLib.TradeDate.LastNTradeDate(1);

        public SqlConnection conn = new SqlConnection(GlobalVar.loginSet.warrantassistant45);

        public Dictionary<string, double> wname2result = new Dictionary<string, double>();
        //在10:40左右額度結果出來後，用來記錄要搶額度的權證，哪些改為用10000張發行
        public List<string> changeTo_1w = new List<string>();
        //用來記需要搶額度的權證
        int applyhour_compete = 0;
        int applymin_compete = 0;
        public List<string> Iscompete = new List<string>();
        //計算每個標的發行幾檔
        public Dictionary<string, int> UidCnt = new Dictionary<string, int>();
        List<string> Market30 = new List<string>();//市值前20大
        List<string> IsSpecial = new List<string>();//特殊標的
        List<string> IsIndex = new List<string>();//臺灣50指數,臺灣中型100指數,櫃買富櫃50指
        public double NonSpecialCallPutRatio = Convert.ToDouble(ConfigurationManager.AppSettings["NonSpecialCallPutRatio"].ToString());
        public double SpecialCallPutRatio = Convert.ToDouble(ConfigurationManager.AppSettings["SpecialCallPutRatio"].ToString());
        public double SpecialKGIALLPutRatio = Convert.ToDouble(ConfigurationManager.AppSettings["SpecialKGIALLPutRatio"].ToString());
        public double ISTOP30MaxIssue = Convert.ToDouble(ConfigurationManager.AppSettings["ISTOP30MaxIssue"].ToString());
        public double NonTOP30MaxIssue = Convert.ToDouble(ConfigurationManager.AppSettings["NonTOP30MaxIssue"].ToString());
        Dictionary<string, UidPutCallDeltaOne> UidDeltaOne = new Dictionary<string, UidPutCallDeltaOne>();
        private static Dictionary<int, string> reasonString = new Dictionary<int, string> {
            { 0," "},
            { 1,"技術面偏多，股價持續看好，因此發行認購權證吸引投資人。" },
            { 2,"基本面良好，股價具有漲升的條件，因此發行認購權證吸引投資人。"},
            { 3, "營運動能具提升潛力，因此發行認購權證吸引投資人。"},
            { 4, "提供投資人槓桿避險工具"},
            { 5, "持續針對不同的履約條件、存續期間及認購認售等發行新條件，提供投資人更多元投資選擇"}
        };
        //計算個股類權證發行單價Q1~Q3
        List<double> PQs = new List<double>();
        double PQ1 = 0;
        double PQ2 = 0;
        double PQ3 = 0;
        double P85 = 0;

        double PQ1_yuan = 0;
        double PQ2_yuan = 0;
        double PQ3_yuan = 0;

        //計算自家 90~120/120~150/150~180/180~ 四個期間發行平均單價
        List<double> T1AvgP = new List<double>();
        List<double> T2AvgP = new List<double>();
        List<double> T3AvgP = new List<double>();
        List<double> T4AvgP = new List<double>();
        private Dictionary<int, int> T2TtoM = new Dictionary<int, int>();

        DataTable HVs;
        public FrmIssueTotal() {
            InitializeComponent();
        }

        private void FrmIssueTotal_Load(object sender, EventArgs e) {
            InitialGrid();
            //sqlTogrid.Add("WarrantName", "權證名稱");
            //sqlTogrid.Add("1500W", "Apply1500W" );
            sqlTogrid.Add("類型","Type");
            sqlTogrid.Add("CP", "CP");
            //sqlTogrid.Add("履約價", "K");
            sqlTogrid.Add("期間", "T");
            sqlTogrid.Add("行使比例","CR");
            //sqlTogrid.Add("IV", "IV");
            sqlTogrid.Add("重設比", "ResetR");
            sqlTogrid.Add("界限比", "BarrierR");
            sqlTogrid.Add("張數", "IssueNum");
            sqlTogrid.Add("獎勵", "UseReward");
            LoadMaxCR();
            
            
            LoadT2TtoM();
            LoadData();
            SetUpdateRecord();
            LoadIsSpecial();
 
            LoadMarket30();

            LoadIsIndex();

            Thread thCheckStrike = new Thread(new ThreadStart(CheckK));
            thCheckStrike.Start();

        }



        private void CheckK()
        {
            try
            {
                for (; ; )
                {
                    try
                    {
                        DateTime now = DateTime.Now;


                        if (now.TimeOfDay.TotalMinutes > 420 && now.TimeOfDay.TotalMinutes <= 480)
                        {
                            //早上都睡覺50分鐘，1330再啟動就好
                            Thread.Sleep(3000000);
                        }
                        if (now.TimeOfDay.TotalMinutes > 480 && now.TimeOfDay.TotalMinutes <= 540)
                        {
                            //早上都睡覺，1330再啟動就好
                            Thread.Sleep(3000000);
                        }
                        if (now.TimeOfDay.TotalMinutes > 540 && now.TimeOfDay.TotalMinutes <= 600)
                        {
                            //早上都睡覺，1330再啟動就好
                            Thread.Sleep(3000000);
                        }
                        if (now.TimeOfDay.TotalMinutes > 600 && now.TimeOfDay.TotalMinutes <= 660)
                        {
                            //早上都睡覺，1330再啟動就好
                            Thread.Sleep(3000000);
                        }
                        if (now.TimeOfDay.TotalMinutes > 660 && now.TimeOfDay.TotalMinutes <= 720)
                        {
                            //早上都睡覺，1330再啟動就好
                            Thread.Sleep(3000000);
                        }
                        if (now.TimeOfDay.TotalMinutes > 720 && now.TimeOfDay.TotalMinutes <= 780)
                        {
                            //早上都睡覺，1330再啟動就好
                            Thread.Sleep(1500000);
                        }


                        if (now.TimeOfDay.TotalMinutes > 720 && now.TimeOfDay.TotalMinutes <= 780)
                        {
                            //早上都睡覺，1330再啟動就好
                            Thread.Sleep(1500000);
                        }
                        if(now.TimeOfDay.TotalMinutes > 811)
                        {

                            string sql = $@"SELECT  [UnderlyingID]
                                      FROM [WarrantAssistant].[dbo].[ApplyOfficial] AS A
                                      LEFT JOIN [WarrantAssistant].[dbo].[WarrantPrices] B on A.UnderlyingID=B.CommodityID
                                      WHERE [CP] IN ('C','P')　AND CASE WHEN [CP] = 'C' THEN (CASE WHEN [K] > IsNull(B.MPrice,0) * 1.5 + 0.00001 THEN '1' ELSE '0' END) ELSE (CASE WHEN [K] < IsNull(B.MPrice,0) * 0.5 - 0.00001 THEN '1' ELSE '0' END) END = '1'";
                            DataTable dt = MSSQL.ExecSqlQry(sql, GlobalVar.loginSet.warrantassistant45);
                            if(dt.Rows.Count > 0)
                            {
                                MessageBox.Show("有履約價沒更改!");
                            }

                            string sql2 = $@"SELECT A.WID
		                                    ,CASE WHEN WClass = 'cdo' THEN (CASE WHEN UID = 'TWA00' THEN TWAIndex ELSE bidUID END - BarrierPrice) / (CASE WHEN UID = 'TWA00' THEN TWAIndex ELSE bidUID END)
				                                    WHEN WClass = 'puo' THEN (BarrierPrice - CASE WHEN UID = 'TWA00' THEN TWAIndex ELSE bidUID END) / (CASE WHEN UID = 'TWA00' THEN TWAIndex ELSE bidUID END)
				                                    END AS MoneynessB
                                    FROM(
	                                    SELECT [TDate],CAST(GETDATE() AS TIME) AS Time,[WID],[UID],[WClass],[ListedDays],[StrikePrice],[Moneyness],[BarrierPrice],[TtoM],[CR],UBidPrice_1330,HV_60D AS HV60D
	                                    FROM [TwData].[dbo].[V_WarrantTrading] where WClass in ('cdo','puo') and TDate = TWData.dbo.previousTrade(CAST(GETDATE() AS DATE),1) and IssuerName = '9200' 
	                                    )AS A
	                                    LEFT JOIN (SELECT [symbol],[bidPrz1] AS bidUID,[askPrz1] AS askUID 
				                                    FROM [WMM3].[dbo].[tblQuote] where LEFT(symbol,2) = '00') AS ASKBID on ASKBID.symbol = A.UID
	                                    LEFT JOIN (SELECT Top 1 LASTINDEX AS TWAIndex 
				                                    FROM [TWTickEquity].[dbo].[TAIEXCalculated] ORDER BY TRADEDATE DESC, UPDATETIME DESC) AS TAIEX ON A.UID = 'TWA00'
	                                    LEFT JOIN (SELECT [symbol] AS symbolWID,[bidPrz1] AS bidWID,[askPrz1] AS askWID 
				                                    FROM [WMM3].[dbo].[tblQuote] where LEN(symbol) = 6) AS ASKBID_WID on ASKBID_WID.symbolWID = A.WID
	                                    LEFT JOIN [WarrantAssistant].[dbo].[ReIssueOfficial] AS 增額列表 ON A.WID = 增額列表.WarrantID
	                                    where (CASE WHEN WClass = 'cdo' THEN (CASE WHEN UID = 'TWA00' THEN TWAIndex ELSE bidUID END - BarrierPrice) / (CASE WHEN UID = 'TWA00' THEN TWAIndex ELSE bidUID END)
				                                    WHEN WClass = 'puo' THEN (BarrierPrice - CASE WHEN UID = 'TWA00' THEN TWAIndex ELSE bidUID END) / (CASE WHEN UID = 'TWA00' THEN TWAIndex ELSE bidUID END)
				                                    END <= 0.015 OR ABS((CASE WHEN UID = 'TWA00' THEN TWAIndex ELSE bidUID END - BarrierPrice)) <= 200)
				                                    AND 增額列表.WarrantID IS NOT NULL";
                            DataTable dt2 = MSSQL.ExecSqlQry(sql2, GlobalVar.loginSet.twData);
                            if (dt2.Rows.Count > 0)
                            {
                                MessageBox.Show("有牛熊證快觸價，取消增額");
                            }

                            Thread.Sleep(30000);

                        }

                        if (now.TimeOfDay.TotalMinutes > 840)
                            break;

                    }
                    catch (ThreadAbortException)
                    {
                    }
                    catch (Exception)
                    {
                    }
                    Thread.Sleep(20000);
                }
            }
            catch (ThreadAbortException)
            {
            }
            catch (Exception)
            {
            }
        }
        private void InitialGrid() {
            dt.Columns.Add("市場", typeof(string));
            dt.Columns.Add("序號", typeof(string));
            dt.Columns.Add("交易員", typeof(string));
            dt.Columns.Add("標的代號", typeof(string));
            dt.Columns.Add("權證名稱", typeof(string));
            dt.Columns.Add("發行價格", typeof(double));
            dt.Columns.Add("跳動價差", typeof(double));
            dt.Columns.Add("1500W", typeof(string));
            dt.Columns.Add("類型", typeof(string));
            dt.Columns.Add("CP", typeof(string));
            dt.Columns.Add("股價", typeof(double));
            dt.Columns.Add("履約價", typeof(double));
            dt.Columns.Add("期間", typeof(int));
            dt.Columns.Add("行使比例", typeof(double));
            dt.Columns.Add("平均Theta", typeof(double));
            dt.Columns.Add("Ratio", typeof(double));
            dt.Columns.Add("HV", typeof(double));
            dt.Columns.Add("IV", typeof(double));
            dt.Columns.Add("IVOri", typeof(double));
            dt.Columns.Add("重設比", typeof(double));
            dt.Columns.Add("張數", typeof(double));
            dt.Columns.Add("約當張數", typeof(double));
            dt.Columns.Add("額度結果", typeof(double));
            dt.Columns.Add("獎勵", typeof(string));
            dt.Columns.Add("說明", typeof(string));
            dt.Columns.Add("額度", typeof(double));
            dt.Columns.Add("可發行檔數", typeof(double));
            dt.Columns.Add("獎勵額度", typeof(double));
            dt.Columns.Add("可發", typeof(double));
            dt.Columns.Add("已發", typeof(string));
            dt.Columns.Add("狀態", typeof(string));
            dt.Columns.Add("順位", typeof(string));
            dt.Columns.Add("發行原因", typeof(string));
            dt.Columns.Add("界限比", typeof(double));
            dt.Columns.Add("財務費用", typeof(double));
            ultraGrid1.DataSource = dt;
            UltraGridBand band0 = ultraGrid1.DisplayLayout.Bands[0];

            band0.Columns["張數"].Format = "N0";
            band0.Columns["約當張數"].Format = "N0";
            band0.Columns["額度結果"].Format = "N0";
            band0.Columns["額度"].Format = "N0";
            band0.Columns["獎勵額度"].Format = "N0";
            band0.Columns["可發"].Format = "N0";
            band0.Columns["平均Theta"].Format = "N0";

            band0.Columns["交易員"].Width = 60;
            band0.Columns["權證名稱"].Width = 150;
            band0.Columns["發行價格"].Width = 70;
            band0.Columns["跳動價差"].Width = 70;
            band0.Columns["標的代號"].Width = 70;
            band0.Columns["1500W"].Width = 70;
            band0.Columns["市場"].Width = 50;
            band0.Columns["類型"].Width = 70;
            band0.Columns["CP"].Width = 40;
            band0.Columns["股價"].Width = 70;
            band0.Columns["履約價"].Width = 70;
            band0.Columns["期間"].Width = 40;
            band0.Columns["行使比例"].Width = 70;
            band0.Columns["平均Theta"].Width = 70;
            band0.Columns["Ratio"].Width = 40;
            band0.Columns["HV"].Width = 40;
            band0.Columns["IV"].Width = 40;
            band0.Columns["重設比"].Width = 70;
            band0.Columns["界限比"].Width = 70;
            band0.Columns["財務費用"].Width = 70;
            band0.Columns["張數"].Width = 80;
            band0.Columns["約當張數"].Width = 80;
            band0.Columns["額度結果"].Width = 60;
            band0.Columns["說明"].Width = 80;
            band0.Columns["額度"].Width = 80;
            band0.Columns["可發行檔數"].Width = 60;
            band0.Columns["獎勵額度"].Width = 80;
            band0.Columns["獎勵"].Width = 40;
            band0.Columns["順位"].Width = 40;
            band0.Columns["狀態"].Width = 200;
            band0.Columns["發行原因"].Width = 80;
            band0.Columns["可發"].Width = 80;
            band0.Columns["已發"].Width = 120;

            band0.Columns["1500W"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center;
            band0.Columns["標的代號"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center;
            band0.Columns["類型"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center;
            band0.Columns["CP"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center;
            band0.Columns["期間"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center;
            band0.Columns["發行價格"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right;
            band0.Columns["跳動價差"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right;
            band0.Columns["股價"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right;
            band0.Columns["履約價"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right;
            band0.Columns["HV"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right;
            band0.Columns["IV"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right;
            band0.Columns["重設比"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right;
            band0.Columns["界限比"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right;
            band0.Columns["財務費用"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right;
            band0.Columns["張數"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right;
            band0.Columns["約當張數"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right;
            band0.Columns["額度結果"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right;
            band0.Columns["額度"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right;
            band0.Columns["獎勵額度"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right;
            band0.Columns["獎勵"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center;
            band0.Columns["順位"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center;
            band0.Columns["平均Theta"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right;

            band0.Columns["發行原因"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Left;
            band0.Columns["可發"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right;
            band0.Columns["已發"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right;
            band0.Override.HeaderAppearance.TextHAlign = Infragistics.Win.HAlign.Left;

            band0.Columns["序號"].Hidden = true;
            band0.Columns["IVOri"].Hidden = true;
            band0.Columns["發行原因"].Hidden = true;
            band0.Columns["狀態"].Hidden = true;
            // To sort multi-column using SortedColumns property
            // This enables multi-column sorting
            this.ultraGrid1.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti;
            
            // It is good practice to clear the sorted columns collection
            band0.SortedColumns.Clear();

            SetButton();
        }
        private void SetUpdateRecord()
        {
            updaterecord_dt.Columns.Add("serialNum", typeof(string));
            updaterecord_dt.Columns.Add("dataname", typeof(string));
            updaterecord_dt.Columns.Add("fromvalue", typeof(string));
            updaterecord_dt.Columns.Add("tovalue", typeof(string));
        }
        private void LoadIsSpecial()
        {

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
        }
        private void LoadT2TtoM()
        {
            T2TtoM.Clear();
            string listedDate = EDLib.TradeDate.NextNTradeDate(3).ToString("yyyyMMdd");

            string sql = $@"SELECT CONVERT(VARCHAR,TradeDate,112) AS TDate ,ROW_NUMBER() OVER(ORDER BY TradeDate) AS TtoM　FROM [DeriPosition].[dbo].[Calendar] WHERE IsTrade='Y' AND [CountryId] = 'TWN' AND [TradeDate] >= '{listedDate}'";

            DataTable dt = MSSQL.ExecSqlQry(sql, GlobalVar.loginSet.tsquoteSqlConnString);
            
            for (int i = 6; i <= 30; i++)
            {
                
                DateTime tempDate = GlobalVar.globalParameter.nextTradeDate3.AddMonths(i);

               
                if (tempDate.Day == GlobalVar.globalParameter.nextTradeDate3.Day)
                    tempDate = tempDate.AddDays(-1);
                
                while (true)
                {
                    if (EDLib.TradeDate.IsTradeDay(tempDate))
                        break;
                    tempDate = tempDate.AddDays(1);
                }
      
                DataRow[] dr = dt.Select($@"TDate = '{tempDate.ToString("yyyyMMdd")}'");
                if (dr.Length > 0)
                {
                    int ttom = Convert.ToInt32(dr[0][1].ToString());
                    if (!T2TtoM.ContainsKey(i))
                        T2TtoM.Add(i, ttom);

                }
            }
        }
        private void LoadMarket30()
        {
            Market30.Clear();
            //市值前30大的表

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
        }
        private void LoadIsIndex()
        {

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
        private void LoadMaxCR()
        {
            UidMaxCR_C.Clear();
            UidMaxCR_P.Clear();
            string sql = $@"SELECT 
                          [標的代號] AS [UID]
	                      ,[權證類別代號] AS [CP]
                          ,MAX([最新執行比例]) AS [CR] 
                      FROM [TwCMData].[dbo].[Warrant總表]
                      WHERE [日期] = '{lastDate.ToString("yyyyMMdd")}' AND [券商代號] = '9200' AND [上市日期] <= '{DateTime.Today.ToString("yyyyMMdd")}' AND [最後交易日] >= '{DateTime.Today.ToString("yyyyMMdd")}' AND [權證類別代號] IN ('c','p') AND LEFT([標的代號],1) <> 'T' AND  LEFT([標的代號],2) <> '00' 
                      GROUP BY [標的代號],[權證類別代號]";
            DataTable dt = MSSQL.ExecSqlQry(sql, GlobalVar.loginSet.twCMData);
            foreach (DataRow dr in dt.Rows)
            {
                string uid = dr["UID"].ToString();

                string cp = dr["CP"].ToString();
                double cr = Convert.ToDouble(dr["CR"].ToString());
                if (cp == "c")
                    UidMaxCR_C.Add(uid, cr);
                if (cp == "p")
                    UidMaxCR_P.Add(uid, cr);
            }
        }
        private void LoadData() {
            try {
                UidDeltaOne.Clear();
                UidCnt.Clear();
                dt.Rows.Clear();
                dt_record.Clear();
                dt_recordByTrader.Clear();
                dt_LongTerm.Clear();
                PQs.Clear();

                T1AvgP.Clear();
                T2AvgP.Clear();
                T3AvgP.Clear();
                T4AvgP.Clear();

                string sqlHVs = $@"SELECT [UID],[HV20],[HV60]
                          FROM [TwData].[dbo].[UnderlyingHitoricalVol]
                          WHERE [TDate] = '{EDLib.TradeDate.LastNTradeDate(1).ToString("yyyyMMdd")}'　AND LEN([UID]) < 5 AND LEFT([UID],2) <> '00'";
                HVs = MSSQL.ExecSqlQry(sqlHVs, GlobalVar.loginSet.twData);


                //抓元大發行中位數
                string sql_yuanIssuePrice = $@"SELECT distinct CONVERT(VARCHAR,[發行日期],112) AS 發行日期,[券商名稱] 
                   ,ROUND(PERCENTILE_CONT(0.25) WITHIN GROUP (ORDER BY 發行價格) OVER (PARTITION BY [券商名稱],[發行日期]),2) AS 個股Q1 
                     ,ROUND(PERCENTILE_CONT(0.5) WITHIN GROUP (ORDER BY 發行價格 ) OVER (PARTITION BY [券商名稱],[發行日期]),2) AS 個股Q2 
                     ,ROUND(PERCENTILE_CONT(0.75) WITHIN GROUP (ORDER BY 發行價格) OVER (PARTITION BY [券商名稱],[發行日期]),2) AS 個股Q3 
                     ,ROUND(PERCENTILE_CONT(0.85) WITHIN GROUP (ORDER BY 發行價格) OVER (PARTITION BY [券商名稱],[發行日期]),2) AS 個股P85 
                  FROM [TwCMData].[dbo].[Warrant總表] WHERE [日期] = '{EDLib.TradeDate.LastNTradeDate(1).ToString("yyyyMMdd")}' and [發行日期] = '{DateTime.Today.ToString("yyyyMMdd")}' and [一般證/牛熊證] = 'False' AND LEN([標的代號]) < 5 AND LEFT([標的代號],2) <> '00' AND [券商名稱] = '元大'";
                DataTable dt_yuanIssuePrice = MSSQL.ExecSqlQry(sql_yuanIssuePrice, GlobalVar.loginSet.twCMData);
                if (dt_yuanIssuePrice.Rows.Count > 0)
                {
                    PQ1_yuan = Convert.ToDouble(dt_yuanIssuePrice.Rows[0][2].ToString());
                    PQ2_yuan = Convert.ToDouble(dt_yuanIssuePrice.Rows[0][3].ToString());
                    PQ3_yuan = Convert.ToDouble(dt_yuanIssuePrice.Rows[0][4].ToString());
                }


                //Load DeltaOne

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
                }

                string StrAvgTheta = $@"(SELECT A.[UID],A.[WClass],ROUND(SUM(A.[AccReleasingLots] * [Theta_IV] * 1000) / COUNT(*),0)  AS 平均Theta
                                          FROM [TwData].[dbo].[V_WarrantTrading] AS A
                                          LEFT JOIN [TwData].[dbo].[URank] AS B ON A.TDate = B.TDate AND A.UID = B.UID AND A.WClass = B.WClass
                                          WHERE A.[TDate] = (SELECT MAX(TDate) FROM [TwData].[dbo].[WarrantFlow])  AND A.[WClass] IN ('c','p') AND LEN(A.[UID]) < 5 AND LEFT(A.[UID],2) <> '00' AND [TtoM] >= 90 
                                          GROUP BY A.TDate,A.UID,A.WClass,B.class HAVING SUM(A.[AccReleasingLots] * [Theta_IV] * 1000) > 0) ";
                DataTable dtAvgTheta = MSSQL.ExecSqlQry(StrAvgTheta, GlobalVar.loginSet.twData);


                //20231207 改穿VOL邏輯:HV20 > HV20(T-1)，HV20 > HV60 近5日漲跌幅(含今日) > 20
                string sql_LongTerm = $@"SELECT  [UnderlyingID],B.HV20Dif,B.HV20,B.HV60,C.LevelPrice,P.MPrice,  ROUND((P.MPrice - C.LevelPrice) * 100 /C.LevelPrice,1) AS 今日漲跌幅,R.累積漲跌幅
  FROM [WarrantAssistant].[dbo].[WarrantUnderlyingSummary]　AS A
  LEFT JOIN (SELECT T1.UID,ROUND((T1.[HV20] - T2.[HV20]),1) AS HV20Dif ,T1.HV20,T1.HV60 FROM (
SELECT  [UID],[HV20],[HV60] FROM [TwData].[dbo].[UnderlyingHitoricalVol]
  WHERE [TDate] = [TwData].[dbo].[PreviousTrade] (GETDATE(),1)) AS T1
  LEFT JOIN (　SELECT  [UID],[HV20],[HV60] FROM [TwData].[dbo].[UnderlyingHitoricalVol]
  WHERE [TDate] = [TwData].[dbo].[PreviousTrade] (GETDATE(),2)) AS T2 ON T1.UID = T2.UID
  WHERE T2.UID IS NOT NULL) AS B ON A.UnderlyingID = B.UID
  LEFT JOIN (SELECT  [StockNo],[LevelPrice]　FROM [TwCMData].[dbo].[RTD1001]　WHERE [TDate] = CONVERT(VARCHAR,GETDATE(),112)) AS C ON A.UnderlyingID = C.StockNo
  LEFT JOIN [WarrantAssistant].[dbo].[WarrantPrices] AS P ON A.UnderlyingID = P.CommodityID
  LEFT JOIN (SELECT  [股票代號],ROUND(SUM([漲幅(%)]),1) AS 累積漲跌幅 FROM [TwCMData].[dbo].[日收盤表排行] WHERE [日期]　>= [TwData].[dbo].[PreviousTrade] (GETDATE(),4)　GROUP BY [股票代號]) AS R ON A.UnderlyingID = R.股票代號
WHERE (B.HV20Dif > 0 AND B.HV20 > B.HV60 AND (ROUND((P.MPrice - C.LevelPrice) * 100 /C.LevelPrice,1) + R.累積漲跌幅) > 20) OR (B.HV20Dif > 0 AND B.HV20 > B.HV60 * 1.2)";


                dt_LongTerm = MSSQL.ExecSqlQry(sql_LongTerm, GlobalVar.loginSet.warrantassistant45);

                string sql = $@"SELECT  a.SerialNumber
                                      ,a.TraderID
	                                  ,b.WarrantName
	                                  ,a.UnderlyingID
                                      ,a.Apply1500W
	                                  ,b.Market
	                                  ,a.Type
	                                  ,a.CP
	                                  ,IsNull(c.MPrice,0) MPrice
                                      ,a.K
                                      ,a.T
                                      ,a.R
                                      ,CASE WHEN (LEN(a.UnderlyingID) < 5 and LEFT(a.UnderlyingID,2) <> '00') THEN ROUND(ISNULL(I.HV60,a.HV),0) ELSE a.HV END AS HV
                                      ,CASE WHEN a.Apply1500W='Y' THEN a.IVNew ELSE a.IV END IVNew
                                      ,a.ResetR
                                      ,a.BarrierR
                                      ,a.FinancialR
                                      ,a.IssueNum
                                      ,b.EquivalentNum
                                      ,b.Result
                                      ,a.IV
                                      ,CASE WHEN a.CP='C' THEN d.Reason ELSE d.ReasonP END Reason
                                      ,a.UseReward
                                      ,IsNull(Floor(E.[WarrantAvailableShares] * {GlobalVar.globalParameter.givenRewardPercent} - IsNull(F.[UsedRewardNum],0)), 0) AS RewardCredit
                                      ,ISNULL(FLOOR(E.CanIssue),0) AS CanIssue
                                      ,CASE WHEN LEN(a.UnderlyingID) < 5 AND LEFT(a.UnderlyingID,2) <> '00' THEN 'N' ELSE 'Y' END AS IndexOrNot
                                      ,ISNULL(G.[說明],'') AS 說明
                                      ,ROUND(ISNULL(CASE WHEN LEN(a.[UnderlyingID]) < 5 AND LEFT(a.[UnderlyingID],2) <> '00' THEN (H.AvailableShares / 1000) * (0.22 - ((CASE WHEN H.[AvailableShares] > 0 THEN (CONVERT(float,REPLACE(H.[AccUsedShares],'%',''))) ELSE 0 END) / 100)) ELSE (H.AvailableShares / 1000) * (1 - ((CASE WHEN H.[AvailableShares] > 0 THEN (CONVERT(float,REPLACE(H.[AccUsedShares],'%',''))) ELSE 0 END) / 100)) END,0),1) AS 可發張數
									  ,H.AccUsedShares
                                  FROM [WarrantAssistant].[dbo].[ApplyOfficial] a
                                  LEFT JOIN [WarrantAssistant].[dbo].[ApplyTotalList] b ON a.SerialNumber=b.SerialNum
                                  LEFT JOIN [WarrantAssistant].[dbo].[WarrantPrices] c on a.UnderlyingID=c.CommodityID
                                  LEFT JOIN Underlying_TraderIssue d on a.UnderlyingID=d.UID 
                                  LEFT JOIN (SELECT [UID], [CanIssue], [WarrantAvailableShares] FROM [WarrantAssistant].[dbo].[WarrantUnderlyingCreditNew] WHERE [UpdateTime] > '{DateTime.Today.ToString("yyyyMMdd")}' ) as E on a.UnderlyingID = E.[UID]
                                  LEFT JOIN [WarrantAssistant].[dbo].[WarrantReward] F on a.UnderlyingID=F.UnderlyingID
                                  LEFT JOIN [WarrantAssistant].[dbo].[WarrantBasic_InitialIV] G on b.WarrantName=G.[WarrantName]
                                  LEFT JOIN [WarrantAssistant].[dbo].[Apply_71] H on b.WarrantName=H.[WarrantName]
                                  LEFT JOIN (SELECT  [UID],[HV60] FROM [TwData].[dbo].[UnderlyingHitoricalVol] WHERE [TDate] = (SELECT MAX([TDate]) FROM [TwData].[dbo].[UnderlyingHitoricalVol])) AS I ON A.UnderlyingID = I.UID
                                  ORDER BY b.Market desc, a.Type, a.CP, a.UnderlyingID, SUBSTRING(a.SerialNumber,9,7),CONVERT(INT,SUBSTRING(a.SerialNumber,18,LEN(a.SerialNumber)-17))"; //or (a.UnderlyingID = 'IX0001' and d.UID ='TWA00')

                DataView dv = DeriLib.Util.ExecSqlQry(sql, GlobalVar.loginSet.warrantassistant45);

                foreach (DataRowView drv in dv) {
                    DataRow dr = dt.NewRow();

                    dr["序號"] = drv["SerialNumber"].ToString();
                    dr["交易員"] = drv["TraderID"].ToString().TrimStart('0');
                    dr["權證名稱"] = drv["WarrantName"].ToString();
                    string uid = drv["UnderlyingID"].ToString(); 
                    dr["標的代號"] = drv["UnderlyingID"].ToString();
                    string underlyingID = drv["UnderlyingID"].ToString();
                    if (!UidCnt.ContainsKey(underlyingID))
                        UidCnt.Add(underlyingID, 0);
                    UidCnt[underlyingID]++;
                    //wclass為小寫  cp為大寫
                    CallPutType cp = drv["CP"].ToString() == "C" ? CallPutType.Call : CallPutType.Put;
                    dr["CP"] = drv["CP"].ToString();
                    dr["1500W"] = drv["Apply1500W"].ToString();
                    dr["市場"] = drv["Market"].ToString();
                    dr["張數"] = drv["IssueNum"];
                    dr["約當張數"] = drv["EquivalentNum"];
                    dr["額度結果"] = drv["Result"];
                    double canIssue = Convert.ToDouble(drv["CanIssue"]);
                    dr["額度"] = canIssue;
                    double canIssueNum = 0;//可發行檔數
                    if ((double)(canIssue / 5000) < 1)
                        canIssueNum = 1;
                    else
                    {
                        double maxCR = 0;
                        if(drv["CP"].ToString() == "C")
                        {
                            if (UidMaxCR_C.ContainsKey(uid))
                                maxCR = UidMaxCR_C[uid];
                        }
                        if (drv["CP"].ToString() == "P")
                        {
                            if (UidMaxCR_P.ContainsKey(uid))
                                maxCR = UidMaxCR_P[uid];
                        }
                        if (maxCR > 0)
                            canIssueNum = Math.Round((double)canIssue / (5000 * maxCR),1);

                    }

                    dr["可發行檔數"] = canIssueNum;
                    dr["IVOri"] = drv["IV"];
                    dr["獎勵"] = drv["UseReward"];
                   
                    double underlyingPrice = 0.0;
                    underlyingPrice = Convert.ToDouble(drv["MPrice"]);
                    dr["股價"] = underlyingPrice;
                    double k = Convert.ToDouble(drv["K"]);
                    dr["履約價"] = k;
                    int t = Convert.ToInt32(drv["T"]);
                    dr["期間"] = t;
                    double cr = Convert.ToDouble(drv["R"]);
                    dr["行使比例"] = cr;
                    dr["HV"] = Convert.ToDouble(drv["HV"]);
                    double vol = Convert.ToDouble(drv["IVNew"]) / 100;
                    dr["IV"] = Convert.ToDouble(drv["IVNew"]);
                   
                    dr["Ratio"] = Math.Round(Convert.ToDouble(drv["IVNew"]) / Convert.ToDouble(drv["HV"]), 2);
                    DataRow[] rAvgTheta = dtAvgTheta.Select($@"UID = '{underlyingID}' AND WClass = '{drv["CP"].ToString()}'");
                    if(rAvgTheta.Length > 0)
                    {
                        dr["平均Theta"] = Convert.ToDouble(rAvgTheta[0][2].ToString());
                    }
                    double resetR = Convert.ToDouble(drv["ResetR"]) / 100;
                    dr["重設比"] = Convert.ToDouble(drv["ResetR"]);
                    //double barrierR = Convert.ToDouble(drv["BarrierR"]);
                    dr["界限比"] = Convert.ToDouble(drv["BarrierR"]);
                    double financialR = Convert.ToDouble(drv["FinancialR"]) / 100;
                    dr["財務費用"] = Convert.ToDouble(drv["FinancialR"]);
                    string warrantType = drv["Type"].ToString();
                    dr["類型"] = warrantType;
                    
                    dr["發行原因"] = drv["Reason"] == DBNull.Value ? " " : reasonString[Convert.ToInt32(drv["Reason"])];


                    double rewardcredit = Convert.ToDouble(drv["RewardCredit"].ToString());
                    dr["獎勵額度"] = rewardcredit;
                    double price = 0.0;
                    double delta = 0.0;
                    double adj = 0.0;
                    string indexOrnot = drv["IndexOrNot"].ToString();
                    //20230218更新
                    //上傳到交易所時發行價格不含ADJ，所以計算時不把ADJ拿進去算
                    /*
                    if(indexOrnot == "Y")
                    {
                        string sqlAdjStr = $@"SELECT [Adj]
                                              FROM [WarrantAssistant].[dbo].[ApplyTempList]
                                              WHERE [UnderlyingID] = '{drv["UnderlyingID"].ToString()}' AND [K] = {k} AND [T] = {t} AND IV = {vol} AND [R] = {cr} AND [IssueNum] = {drv["IssueNum"].ToString()}";
                        DataTable dtAdj = MSSQL.ExecSqlQry(sqlAdjStr, GlobalVar.loginSet.warrantassistant45);
                        if(dtAdj.Rows.Count > 0)
                        {
                            adj = Convert.ToDouble(dtAdj.Rows[0][0].ToString());
                        }
                    }
                    */

                    if (underlyingPrice != 0) {
                        if (underlyingID.Length > 4 && underlyingID.Substring(0, 2) != "00")
                        {
                            if (warrantType == "牛熊證")
                            {
                                price = Pricing.BullBearWarrantPrice(cp, underlyingPrice + adj, resetR, GlobalVar.globalParameter.interestRate_Index, vol, t, financialR, cr);
                                delta = 1.0;
                            }
                            else if (warrantType == "重設型")
                            {
                                price = Pricing.ResetWarrantPrice(cp, underlyingPrice + adj, resetR, GlobalVar.globalParameter.interestRate_Index, vol, t, cr);
                                delta = Pricing.Delta(cp, underlyingPrice + adj, k, GlobalVar.globalParameter.interestRate_Index, vol, (t * 30.0) / GlobalVar.globalParameter.dayPerYear, GlobalVar.globalParameter.interestRate_Index) * cr;
                            }
                            else
                            {
                                price = Pricing.NormalWarrantPrice(cp, underlyingPrice + adj, k, GlobalVar.globalParameter.interestRate_Index, vol, t, cr);
                                delta = Pricing.Delta(cp, underlyingPrice + adj, k, GlobalVar.globalParameter.interestRate_Index, vol, (t * 30.0) / GlobalVar.globalParameter.dayPerYear, GlobalVar.globalParameter.interestRate_Index) * cr;
                            }
                        }
                        else
                        {
                            if (warrantType == "牛熊證")
                            {
                                price = Pricing.BullBearWarrantPrice(cp, underlyingPrice + adj, resetR, GlobalVar.globalParameter.interestRate, vol, t, financialR, cr);
                                delta = 1.0;
                            }
                            else if (warrantType == "重設型")
                            {
                                price = Pricing.ResetWarrantPrice(cp, underlyingPrice + adj, resetR, GlobalVar.globalParameter.interestRate, vol, t, cr);
                                delta = Pricing.Delta(cp, underlyingPrice + adj, k, GlobalVar.globalParameter.interestRate, vol, (t * 30.0) / GlobalVar.globalParameter.dayPerYear, GlobalVar.globalParameter.interestRate) * cr;
                            }
                            else
                            {
                                price = Pricing.NormalWarrantPrice(cp, underlyingPrice + adj, k, GlobalVar.globalParameter.interestRate, vol, t, cr);
                                delta = Pricing.Delta(cp, underlyingPrice + adj, k, GlobalVar.globalParameter.interestRate, vol, (t * 30.0) / GlobalVar.globalParameter.dayPerYear, GlobalVar.globalParameter.interestRate) * cr;
                            }
                        }
                    }
                    double jumpSize = 0.0;
                    double multiplier = EDLib.Tick.UpTickSize(drv["UnderlyingID"].ToString(), underlyingPrice + adj);
                    
                    jumpSize = delta * multiplier;
                    dr["跳動價差"] = Math.Round(jumpSize, 4);
                    dr["發行價格"] = Math.Round(price, 2);
                  
                    if(underlyingID.Length < 5 && underlyingID.Substring(0,2) != "00" && warrantType != "牛熊證")
                    {
                        PQs.Add(Math.Round(price, 2));

                        if(drv["CP"].ToString() == "C")
                        {
                            if (T2TtoM[t] >= 90 && T2TtoM[t] < 120)
                                T1AvgP.Add(price);
                            else if (T2TtoM[t] >= 120 && T2TtoM[t] < 150)
                                T2AvgP.Add(price);
                            else if (T2TtoM[t] >= 150 && T2TtoM[t] < 180)
                                T3AvgP.Add(price);
                            else
                                T4AvgP.Add(price);
                        }
                    }
                    dr["說明"] = drv["說明"].ToString();
                    dr["可發"] = Convert.ToDouble(drv["可發張數"]);
                    if(drv["AccUsedShares"].ToString().Length > 0)
                        dr["已發"] = Math.Round(Convert.ToDouble(drv["AccUsedShares"].ToString().Replace("%","")),1).ToString();
                    else
                        dr["已發"] = drv["AccUsedShares"].ToString();

                    dt.Rows.Add(dr);
                }
                int count = PQs.Count;
                PQs.Sort();
                if (count % 4 == 0)
                {
                    PQ1 = (PQs[count / 4 - 1] + PQs[count / 4]) / 2;
                    if (count % 2 == 0)
                    {
                        PQ2 = (PQs[count / 2 - 1] + PQs[count / 2]) / 2;
                    }
                    else
                    {
                        PQ2 = PQs[count / 2];
                    }
                    PQ3 = (PQs[3 * count / 4 - 1] + PQs[3 * count / 4]) / 2;
                }
                else
                {
                    PQ1 = PQs[count / 4];
                    if (count % 2 == 0)
                    {
                        PQ2 = (PQs[count / 2 - 1] + PQs[count / 2]) / 2;
                    }
                    else
                    {
                        PQ2 = PQs[count / 2];
                    }
                    PQ3 = PQs[3 * count / 4];
                }
                double realIndex = 0.85 * (count - 1);
                int index = (int)realIndex;
                double frac = realIndex - index;
                if (index + 1 < count)
                    P85 = Math.Round(PQs[index] * (1 - frac) + PQs[index + 1] * frac, 2);
                else
                    P85 = Math.Round(PQs[index], 2);
                double t1PSum = 0;
                double t2PSum = 0;
                double t3PSum = 0;
                double t4PSum = 0;
                foreach(double d in T1AvgP)
                {
                    t1PSum = t1PSum + d;
                }
                if (t1PSum > 0)
                    t1PSum = t1PSum / T1AvgP.Count;
                foreach (double d in T2AvgP)
                {
                    t2PSum = t2PSum + d;
                }
                if (t2PSum > 0)
                    t2PSum = t2PSum / T2AvgP.Count;
                foreach (double d in T3AvgP)
                {
                    t3PSum = t3PSum + d;
                }
                if (t3PSum > 0)
                    t3PSum = t3PSum / T3AvgP.Count;
                foreach (double d in T4AvgP)
                {
                    t4PSum = t4PSum + d;
                }
                if (t4PSum > 0)
                    t4PSum = t4PSum / T4AvgP.Count;

                toolStripLabel3.Text = $@"Q1[{Math.Round(PQ1,2)}]元[{Math.Round(PQ1_yuan,2)}]   Q2[{Math.Round(PQ2,2)}]元[{Math.Round(PQ2_yuan, 2)}]   Q3[{Math.Round(PQ3,2)}]元[{Math.Round(PQ3_yuan, 2)}]   P85[{P85}]";
                toolStripLabel4.Text = $@"[90~119]:[{Math.Round(t1PSum  , 2)}]元,[120~149]:[{Math.Round(t2PSum, 2)}]元,[150~179]:[{Math.Round(t3PSum, 2)}]元,[180~]:[{Math.Round(t4PSum, 2)}]元 [6]:{T2TtoM[6]}[7]:{T2TtoM[7]}[8]:{T2TtoM[8]}[9]:{T2TtoM[9]}[10]:{T2TtoM[10]}[11]:{T2TtoM[11]}";
                foreach (DataColumn dc in dt.Columns)
                {
                    if(!ColumnName.Contains(dc.ColumnName))
                    ColumnName.Add(dc.ColumnName);
                }
                
               
                //抓更新紀錄的資料
                //ApplyKind = 4 代表更新資料已確認，不須再顯示
                string sql_record = $@"SELECT A.[SerialNumber], A.[DataName], A.[UpdateCount], B.[FromValue], B.[TraderID]
                                    FROM (SELECT [SerialNumber], [DataName], MAX([UpdateCount]) AS UpdateCount
                                          FROM [WarrantAssistant].[dbo].[ApplyTotalRecord]
                                          WHERE [UpdateTime] >= CONVERT(varchar, getdate(), 112) 
                                          GROUP BY [SerialNumber] ,[DataName]) AS A
                                    LEFT JOIN [WarrantAssistant].[dbo].[ApplyTotalRecord] AS B 
                                    ON A.[SerialNumber] = B.[SerialNumber] AND A.[DataName] =B.[DataName] AND A.[UpdateCount] =B.[UpdateCount]
                                    WHERE A.[UpdateCount] > 0 AND A.[UpdateCount] < 9999 AND B.UpdateTime > '{DateTime.Today.ToString("yyyyMMdd")}' AND B.[ApplyKind] <> '4'";
                DataTable dv_record = MSSQL.ExecSqlQry(sql_record, GlobalVar.loginSet.warrantassistant45);
                //存進record table
                foreach (DataRow dr in dv_record.Rows)
                {
                    string serialNum = dr["SerialNumber"].ToString();
                    string dataName = dr["DataName"].ToString();
                    string FromValue = dr["FromValue"].ToString();
                    string TraderID = dr["TraderID"].ToString().TrimStart('0');
                    if (!dt_record.ContainsKey(serialNum))
                        dt_record.Add(serialNum, new Dictionary<string, string>());
                    dt_record[serialNum].Add(dataName, FromValue);
                    if(!dt_recordByTrader.ContainsKey(serialNum))
                        dt_recordByTrader.Add(serialNum, TraderID);
                }
                Iscompete.Clear();

                string sql2 = $@"SELECT [UnderlyingID]
                          FROM [WarrantAssistant].[dbo].[Apply_71]
                          WHERE len([OriApplyTime])> 0";
                DataTable dt2 = MSSQL.ExecSqlQry(sql2, GlobalVar.loginSet.warrantassistant45);
                foreach (DataRow dr in dt2.Rows)
                {
                    Iscompete.Add(dr["UnderlyingID"].ToString());
                }
                changeTo_1w.Clear();
                string sql_changeTo1w = $@"SELECT [SerialNumber] 
                                    FROM [WarrantAssistant].[dbo].[ApplyTotalRecord]
                                    WHERE [UpdateTime] >= CONVERT(VARCHAR, GETDATE(), 112) AND [ApplyKind] = '3'";
                DataTable dt_changeTo1w = MSSQL.ExecSqlQry(sql_changeTo1w, GlobalVar.loginSet.warrantassistant45);
                if (DateTime.Now.TimeOfDay.TotalMinutes >= GlobalVar.globalParameter.resultTime)
                {
                    string sql_maxOri = $@"SELECT Max([OriApplyTime]) AS OriTime
                                    FROM [WarrantAssistant].[dbo].[Apply_71]";
                    DataTable dt_maxOri = MSSQL.ExecSqlQry(sql_maxOri, GlobalVar.loginSet.warrantassistant45);

                    foreach (DataRow dr in dt_maxOri.Rows)
                    {
                        if (dr["OriTime"].ToString().Length > 0)
                        {
                            applyhour_compete = Convert.ToInt32(dr["OriTime"].ToString().Substring(0, 2));
                            applymin_compete = Convert.ToInt32(dr["OriTime"].ToString().Substring(3, 2));
                        }
                    }
                    if ((applyhour_compete == 0) && (applymin_compete == 0))
                    {

                        string sql_maxApp = $@"SELECT MAX([ApplyTime]) AS AppTime
                                              FROM [WarrantAssistant].[dbo].[Apply_71]
                                              where [ApplyTime] <='10:30:00'";
                        DataTable dt_maxApp = MSSQL.ExecSqlQry(sql_maxApp, GlobalVar.loginSet.warrantassistant45);

                        foreach (DataRow dr in dt_maxApp.Rows)
                        {
                            applyhour_compete = Convert.ToInt32(dr["AppTime"].ToString().Substring(0, 2));
                            applymin_compete = Convert.ToInt32(dr["AppTime"].ToString().Substring(3, 2));
                        }
                    }
                }
                foreach (DataRow dr in dt_changeTo1w.Rows)
                {
                    string serialnum = dr["SerialNumber"].ToString();
                    changeTo_1w.Add(serialnum);
                }
                string sqltemp = $@"SELECT A.[SerialNum]
                              ,CASE WHEN A.[ApplyKind]='1' THEN '新發' ELSE '增額' END ApplyKind
                              ,A.[UnderlyingID]
                              ,A.[CR]
                              ,A.[CP]
                              ,A.[IssueNum]
                              ,A.[UseReward] 
                              ,B.[ApplyTime]
                              ,B.[OriApplyTime]
                          FROM [WarrantAssistant].[dbo].[ApplyTotalList] A
                          LEFT JOIN [WarrantAssistant].[dbo].[Apply_71] B on A.[SerialNum] = B.[SerialNum] ";
                DataTable dvtemp = MSSQL.ExecSqlQry(sqltemp, GlobalVar.loginSet.warrantassistant45);

                foreach (DataRow dr in dvtemp.Rows)
                {
                    bool needadd = true;
                    string uid = dr["UnderlyingID"].ToString();
                    string cp = dr["UnderlyingID"].ToString();
                    string isreward = dr["UseReward"].ToString();
                    string apytime = dr["ApplyTime"].ToString();//時間的全部字串
                    string applyTime = "";
                    if (apytime != string.Empty)
                        applyTime = dr["ApplyTime"].ToString().Substring(0, 2);//時間幾點
                    string oriapplyTime = dr["OriApplyTime"].ToString();
                    double cr = Convert.ToDouble(dr["CR"].ToString());
                    double issueNum = Convert.ToDouble(dr["IssueNum"].ToString());
                    if (UidDeltaOne.ContainsKey(uid))
                    {
                        if (applyTime == "22" || (apytime.Length == 0) && Iscompete.Contains(uid) && isreward == "N")
                            needadd = false;
                        if (needadd)
                        {
                            if (cp == "C")
                            {
                                UidDeltaOne[uid].KgiCallDeltaOne += issueNum * cr;
                                if (UidDeltaOne[uid].KgiPutNum == 0)
                                    UidDeltaOne[uid].KgiCallPutRatio = 100;
                                else
                                    UidDeltaOne[uid].KgiCallPutRatio = Math.Round((double)UidDeltaOne[uid].KgiCallDeltaOne / (double)UidDeltaOne[uid].KgiPutDeltaOne, 4);
                                UidDeltaOne[uid].AllCallDeltaOne += issueNum * cr;
                                if (UidDeltaOne[uid].AllPutDeltaOne == 0)
                                    UidDeltaOne[uid].KgiAllPutRatio = 100;
                                else
                                    UidDeltaOne[uid].KgiAllPutRatio = Math.Round((double)UidDeltaOne[uid].KgiPutDeltaOne / (double)UidDeltaOne[uid].AllPutDeltaOne, 4);
                            }
                            else
                            {
                                UidDeltaOne[uid].KgiPutDeltaOne += issueNum * cr;
                                UidDeltaOne[uid].KgiCallPutRatio = Math.Round((double)UidDeltaOne[uid].KgiCallDeltaOne / (double)UidDeltaOne[uid].KgiPutDeltaOne, 4);
                                UidDeltaOne[uid].AllPutDeltaOne += issueNum * cr;
                                UidDeltaOne[uid].KgiAllPutRatio = Math.Round((double)UidDeltaOne[uid].KgiPutDeltaOne / (double)UidDeltaOne[uid].AllPutDeltaOne, 4);
                                UidDeltaOne[uid].KgiPutNum++;
                            }
                        }
                    }
                }
            } catch (Exception ex) {
                MessageBox.Show(ex.Message);
            }
        }

        private void UpdateData() {
            try {

                string cmdText = "UPDATE [ApplyOfficial] SET K=@K, T=@T, R=@R, HV=@HV, IV=@IV, ResetR=@ResetR, BarrierR=@BarrierR, FinancialR=@FinancialR, Type=@Type, CP=@CP, Apply1500W=@Apply1500W, MDate=@MDate, TempName=@TempName WHERE SerialNumber=@SerialNumber";
                List<SqlParameter> pars = new List<SqlParameter>();
                pars.Add(new SqlParameter("@K", SqlDbType.Float));
                pars.Add(new SqlParameter("@T", SqlDbType.Int));
                pars.Add(new SqlParameter("@R", SqlDbType.Float));
                pars.Add(new SqlParameter("@HV", SqlDbType.Float));
                pars.Add(new SqlParameter("@IV", SqlDbType.Float));
                pars.Add(new SqlParameter("@ResetR", SqlDbType.Float));
                pars.Add(new SqlParameter("@BarrierR", SqlDbType.Float));
                pars.Add(new SqlParameter("@FinancialR", SqlDbType.Float));
                pars.Add(new SqlParameter("@Type", SqlDbType.VarChar));
                pars.Add(new SqlParameter("@CP", SqlDbType.VarChar));
                pars.Add(new SqlParameter("@Apply1500W", SqlDbType.VarChar));
                //pars.Add(new SqlParameter("@TraderID", SqlDbType.VarChar));
                pars.Add(new SqlParameter("@MDate", SqlDbType.DateTime));
                pars.Add(new SqlParameter("@TempName", SqlDbType.VarChar));
                pars.Add(new SqlParameter("@SerialNumber", SqlDbType.VarChar));

                SQLCommandHelper h = new SQLCommandHelper(GlobalVar.loginSet.warrantassistant45, cmdText, pars);

                foreach (Infragistics.Win.UltraWinGrid.UltraGridRow r in ultraGrid1.Rows) {
                    double k = r.Cells["履約價"].Value == DBNull.Value ? 0 : Convert.ToDouble(r.Cells["履約價"].Value);
                    double t = r.Cells["期間"].Value == DBNull.Value ? 0 : Convert.ToDouble(r.Cells["期間"].Value);
                    double cr = r.Cells["行使比例"].Value == DBNull.Value ? 0 : Convert.ToDouble(r.Cells["行使比例"].Value);
                    double hv = r.Cells["HV"].Value == DBNull.Value ? 0 : Convert.ToDouble(r.Cells["HV"].Value);
                    double iv = 0.0;
                    double resetR = r.Cells["重設比"].Value == DBNull.Value ? 0 : Convert.ToDouble(r.Cells["重設比"].Value);
                    double barrierR = r.Cells["界限比"].Value == DBNull.Value ? 0 : Convert.ToDouble(r.Cells["界限比"].Value);
                    double financialR = r.Cells["財務費用"].Value == DBNull.Value ? 0 : Convert.ToDouble(r.Cells["財務費用"].Value);
                    string type = r.Cells["類型"].Value.ToString();
                    string cp = r.Cells["CP"].Value.ToString();
                    string apply1500w = r.Cells["1500W"].Value.ToString();
                    string serialNumber = r.Cells["序號"].Value.ToString();
                   
                    string warrantType = "";

                    DateTime expiryDate = GlobalVar.globalParameter.nextTradeDate3.AddMonths(Convert.ToInt32(t));
                    if (expiryDate.Day == GlobalVar.globalParameter.nextTradeDate3.Day)
                        expiryDate = expiryDate.AddDays(-1);
                    string sqlTemp = $"SELECT TOP 1 TradeDate from TradeDate WHERE IsTrade='Y' AND TradeDate >= '{expiryDate.ToString("yyyy-MM-dd")}'";
                    //DataView dvTemp = DeriLib.Util.ExecSqlQry(sqlTemp, GlobalVar.loginSet.tsquoteSqlConnString);
                    DataTable dvTemp = MSSQL.ExecSqlQry(sqlTemp, GlobalVar.loginSet.tsquoteSqlConnString);
                    foreach (DataRow drTemp in dvTemp.Rows)
                    {
                        expiryDate = Convert.ToDateTime(drTemp["TradeDate"]);
                    }
                    int month = expiryDate.Month;
                    string expiryMonth = month.ToString();
                    if (month >= 10)
                    {
                        if (month == 10)
                            expiryMonth = "A";
                        if (month == 11)
                            expiryMonth = "B";
                        if (month == 12)
                            expiryMonth = "C";
                    }

                    string expiryYear = expiryDate.AddYears(-1).ToString("yyyy");
                    expiryYear = expiryYear.Substring(expiryYear.Length - 1, 1);
                    if (type == "牛熊證")
                    {
                        if (cp == "P")
                            warrantType = "熊";                 
                        else
                            warrantType = "牛";
                    }
                    else
                    {
                        if (cp == "P")
                            warrantType = "售";       
                        else
                            warrantType = "購";      
                    }
                    string tempName = r.Cells["權證名稱"].Value.ToString();
                    int length = tempName.Length;
                    tempName = tempName.Substring(0, length - 5);
                    tempName = tempName + expiryYear + expiryMonth + warrantType;
                    //MessageBox.Show(tempName);
                    if (apply1500w == "Y")
                        iv = r.Cells["IVOri"].Value == DBNull.Value ? 0 : Convert.ToDouble(r.Cells["IVOri"].Value);
                    else
                        iv = r.Cells["IV"].Value == DBNull.Value ? 0 : Convert.ToDouble(r.Cells["IV"].Value);

                    h.SetParameterValue("@K", k);
                    h.SetParameterValue("@T", t);
                    h.SetParameterValue("@R", cr);
                    h.SetParameterValue("@HV", hv);
                    h.SetParameterValue("@IV", iv);
                    h.SetParameterValue("@ResetR", resetR);
                    h.SetParameterValue("@BarrierR", barrierR);
                    h.SetParameterValue("@FinancialR", financialR);
                    h.SetParameterValue("@Type", type);
                    h.SetParameterValue("@CP", cp);
                    h.SetParameterValue("@Apply1500W", apply1500w);
                    //h.SetParameterValue("@TraderID", traderID);
                    h.SetParameterValue("@MDate", DateTime.Now);
                    h.SetParameterValue("@TempName", tempName);
                    h.SetParameterValue("@SerialNumber", serialNumber);
                    h.ExecuteCommand();
                }
                h.Dispose();

                string sql = $@"DELETE FROM [WarrantAssistant].[dbo].[ApplyTotalList] WHERE [ApplyKind]='1'";
                string sql2 = $@"INSERT INTO [WarrantAssistant].[dbo].[ApplyTotalList] ([ApplyKind],[SerialNum],[Market],[UnderlyingID],[WarrantName],[CR] ,[IssueNum],[EquivalentNum],[Credit],[RewardCredit],[Type],[CP],[UseReward],[MarketTmr],[TraderID],[MDate],UserID)
                                SELECT '1',a.SerialNumber, isnull(b.Market, 'TSE'), a.UnderlyingID, a.TempName, a.R, a.IssueNum, ROUND(a.R*a.IssueNum, 2), b.IssueCredit, b.RewardIssueCredit, a.Type, a.CP, a.UseReward,'N', a.TraderID, GETDATE(), a.UserID
                                FROM [WarrantAssistant].[dbo].[ApplyOfficial] a
                                LEFT JOIN [WarrantAssistant].[dbo].[WarrantUnderlyingSummary] b ON a.UnderlyingID=b.UnderlyingID";
                string sql3 = "UPDATE [WarrantAssistant].[dbo].[ApplyTotalList] SET Result=0";
                string sql4 = @"UPDATE [WarrantAssistant].[dbo].[ApplyTotalList] 
                                SET Result= B.Result 
                                FROM [WarrantAssistant].[dbo].[Apply_71] B
                                WHERE [ApplyTotalList].[WarrantName]=B.WarrantName
                                AND [ApplyTotalList].[ApplyKind] = '1'";
                string sql5 = @"UPDATE [WarrantAssistant].[dbo].[ApplyTotalList]
                                SET Result= CASE WHEN [RewardCredit]>=[EquivalentNum] THEN [EquivalentNum] ELSE [RewardCredit] END
                               WHERE [UseReward]='Y'";


                conn.Open();
                MSSQL.ExecSqlCmd(sql, conn);
                MSSQL.ExecSqlCmd(sql2, conn);
                conn.Close();

                //------------------------------------------------------
                string sql6 = "SELECT [SerialNum], [WarrantName] FROM [WarrantAssistant].[dbo].[ApplyTotalList] WHERE [ApplyKind]='1'";
                
                DataTable dv = MSSQL.ExecSqlQry(sql6, GlobalVar.loginSet.warrantassistant45);

                string cmdText_2 = "UPDATE [ApplyTotalList] SET WarrantName=@WarrantName WHERE SerialNum=@SerialNum";
                List<SqlParameter> pars_2 = new List<SqlParameter> {
                    new SqlParameter("@WarrantName", SqlDbType.VarChar),
                    new SqlParameter("@SerialNum", SqlDbType.VarChar)
                };
                SQLCommandHelper h_2 = new SQLCommandHelper(GlobalVar.loginSet.warrantassistant45, cmdText_2, pars_2);

                foreach (DataRow dr in dv.Rows)
                {
                    string serialNum = dr["SerialNum"].ToString();
                    string warrantName = dr["WarrantName"].ToString();

                    string sqlTemp = $"select top (1) WarrantName from (SELECT [WarrantName] FROM [WarrantAssistant].[dbo].[WarrantBasic] WHERE SUBSTRING(WarrantName,1,(len(WarrantName)-3))='{warrantName.Substring(0, warrantName.Length - 1)}' union "
                     + $" SELECT [WarrantName] FROM [WarrantAssistant].[dbo].[ApplyTotalList] WHERE [ApplyKind]='1' AND CONVERT(INT, SUBSTRING([SerialNum], 18 ,LEN([SerialNum])-18 + 1)) <  CONVERT(INT, SUBSTRING('{serialNum}', 18, LEN('{serialNum}')-18 + 1)) AND SUBSTRING(WarrantName,1,(len(WarrantName)-3))='{warrantName.Substring(0, warrantName.Length - 1)}') as tb1 "
                     + " order by SUBSTRING(WarrantName,len(WarrantName)-1,len(WarrantName)) desc";
                    
                    DataTable dvTemp = MSSQL.ExecSqlQry(sqlTemp, GlobalVar.loginSet.warrantassistant45);
                    int count = 0;
                    if (dvTemp.Rows.Count > 0)
                    {
                        string lastWarrantName = dvTemp.Rows[0][0].ToString();
                        if (!int.TryParse(lastWarrantName.Substring(lastWarrantName.Length - 2, 2), out count))
                            MessageBox.Show("parse failed " + lastWarrantName);
                    }

                    //if (dvTemp.Count > 0)
                    //   count += dvTemp.Count;

                    /*sqlTemp = "SELECT [WarrantName] FROM [EDIS].[dbo].[ApplyTotalList] WHERE [ApplyKind]='1' AND [SerialNum]<" + serialNum + " AND SUBSTRING(WarrantName,1,(len(WarrantName)-3))='" + warrantName.Substring(0, warrantName.Length - 1) + "'";
                    dvTemp = DeriLib.Util.ExecSqlQry(sqlTemp, GlobalVar.loginSet.edisSqlConnString);
                    if (dvTemp.Count > 0)
                        count += dvTemp.Count;*/

                    warrantName = warrantName + (count + 1).ToString("0#");

                    h_2.SetParameterValue("@WarrantName", warrantName);
                    h_2.SetParameterValue("@SerialNum", serialNum);
                    h_2.ExecuteCommand();
                }
                h_2.Dispose();
                conn.Open();
                MSSQL.ExecSqlCmd(sql3, conn);
                MSSQL.ExecSqlCmd(sql4, conn);
                MSSQL.ExecSqlCmd(sql5, conn);
                conn.Close();

                UpdateRecord();
                LoadData();
                GlobalUtility.LogInfo("Info", GlobalVar.globalParameter.userID + " 更新發行總表");
            } catch (Exception ex) {
                MessageBox.Show(ex.Message);
            }
        }

        private void UpdateRecord()//紀錄任何權證發行與修改，存在ApplyTotalRecord
        {
            updaterecord_dt.Rows.Clear();
            int length = ultraGrid1.Rows.Count;
            for(int i = length - 1; i >= 0; i--)
            {
                string serialNum = ultraGrid1.Rows[i].Cells[1].Value.ToString();
                try
                {
                    DataRow back_row = back_dt.Select($"序號= '{serialNum}'")[0];
                    foreach (string colname in sqlTogrid.Keys)
                    {
                        if (back_row[colname].ToString() != ultraGrid1.Rows[i].Cells[colname].Value.ToString())
                        {
                            DataRow dr = updaterecord_dt.NewRow();
                            dr["serialNum"] = serialNum;
                            dr["dataname"] = sqlTogrid[colname];
                            dr["fromvalue"] = back_row[colname].ToString();
                            dr["tovalue"] = ultraGrid1.Rows[i].Cells[colname].Value.ToString();
                            updaterecord_dt.Rows.Add(dr);
                        }

                    }
                }
                catch
                {
                    continue;
                }
            }
            
            foreach(DataRow dr in updaterecord_dt.Rows)
            {
                //MessageBox.Show($"{dr["serialNum"].ToString()}  {dr["dataname"].ToString()}  {dr["fromvalue"].ToString()}  {dr["tovalue"].ToString()}");
                
                int count = -1;
                string serialNum = dr["serialNum"].ToString();
                string dataname = dr["dataname"].ToString();
                string fromvalue = dr["fromvalue"].ToString();
                string tovalue = dr["tovalue"].ToString();

                string sql_count = $@"SELECT A.[UpdateCount]
                                    FROM (SELECT [SerialNumber], [DataName], MAX([UpdateCount]) AS UpdateCount
                                          FROM [WarrantAssistant].[dbo].[ApplyTotalRecord]
                                          WHERE [UpdateTime] >= CONVERT(varchar, getdate(), 112)
                                          GROUP BY[SerialNumber],[DataName]) AS A
                                    WHERE A.[UpdateCount] < 9999 AND A.[SerialNumber] ='{serialNum}' AND A.DataName= '{dataname}'";
                DataTable dv_count = MSSQL.ExecSqlQry(sql_count, GlobalVar.loginSet.warrantassistant45);

                foreach (DataRow drr in dv_count.Rows)
                {
                     count = Convert.ToInt32(drr["UpdateCount"].ToString());
                }

                string sql3 = $@"INSERT INTO [WarrantAssistant].[dbo].[ApplyTotalRecord] ([UpdateTime], [UpdateType], [TraderID], [SerialNumber]
                                            , [ApplyKind], [DataName] ,[FromValue], [ToValue], [UpdateCount])
                                           VALUES(GETDATE(), 'UPDATE', '{userID}', {serialNum}, '1','{dataname}','{fromvalue}','{tovalue}',{count + 1})";
                if (count >= 0)
                    MSSQL.ExecSqlCmd(sql3, GlobalVar.loginSet.warrantassistant45);

            }
        }
        private void SetButton() {
            UltraGridBand band0 = ultraGrid1.DisplayLayout.Bands[0];
            band0.Columns["狀態"].CellActivation = Activation.NoEdit;
            band0.Columns["Ratio"].CellActivation = Activation.NoEdit;
            if (isEdit) {
                band0.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.Default;
                band0.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
                band0.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.True;

                band0.Columns["1500W"].CellActivation = Activation.AllowEdit;
                band0.Columns["類型"].CellActivation = Activation.AllowEdit;
                band0.Columns["CP"].CellActivation = Activation.AllowEdit;
                band0.Columns["履約價"].CellActivation = Activation.AllowEdit;
                band0.Columns["期間"].CellActivation = Activation.AllowEdit;
                band0.Columns["HV"].CellActivation = Activation.AllowEdit;
                band0.Columns["IV"].CellActivation = Activation.AllowEdit;
                band0.Columns["重設比"].CellActivation = Activation.AllowEdit;
                band0.Columns["界限比"].CellActivation = Activation.AllowEdit;
                band0.Columns["財務費用"].CellActivation = Activation.AllowEdit;
                band0.Columns["行使比例"].CellActivation = Activation.AllowEdit;
                //band0.Columns["獎勵"].CellActivation = Activation.AllowEdit;

                band0.Columns["交易員"].CellAppearance.BackColor = Color.LightGray;
                band0.Columns["發行價格"].CellAppearance.BackColor = Color.LightGray;
                band0.Columns["跳動價差"].CellAppearance.BackColor = Color.LightGray;
                band0.Columns["標的代號"].CellAppearance.BackColor = Color.LightGray;
                band0.Columns["市場"].CellAppearance.BackColor = Color.LightGray;
                //ultraGrid1.DisplayLayout.Bands[0].Columns["類型"].CellAppearance.BackColor = Color.LightGray;
                band0.Columns["股價"].CellAppearance.BackColor = Color.LightGray;
                //band0.Columns["行使比例"].CellAppearance.BackColor = Color.LightGray;

                toolStripButtonReload.Visible = false;
                toolStripButtonEdit.Visible = false;
                toolStripButtonConfirm.Visible = true;
                toolStripButtonCancel.Visible = true;

                band0.Columns["張數"].Hidden = true;
                band0.Columns["約當張數"].Hidden = true;
                band0.Columns["額度結果"].Hidden = true;
                band0.Columns["可發"].Hidden = true;
                band0.Columns["已發"].Hidden = true;
                band0.Columns["順位"].Hidden = true;

            } else {
                band0.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
                band0.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
                band0.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;

                band0.Columns["市場"].CellActivation = Activation.NoEdit;
                band0.Columns["交易員"].CellActivation = Activation.NoEdit;
                band0.Columns["權證名稱"].CellActivation = Activation.NoEdit;
                band0.Columns["發行價格"].CellActivation = Activation.NoEdit;
                band0.Columns["跳動價差"].CellActivation = Activation.NoEdit;
                band0.Columns["1500W"].CellActivation = Activation.NoEdit;
                band0.Columns["標的代號"].CellActivation = Activation.NoEdit;
                band0.Columns["類型"].CellActivation = Activation.NoEdit;
                band0.Columns["CP"].CellActivation = Activation.NoEdit;
                band0.Columns["股價"].CellActivation = Activation.NoEdit;
                band0.Columns["履約價"].CellActivation = Activation.NoEdit;
                band0.Columns["期間"].CellActivation = Activation.NoEdit;
                band0.Columns["行使比例"].CellActivation = Activation.NoEdit;
                band0.Columns["HV"].CellActivation = Activation.NoEdit;
                band0.Columns["IV"].CellActivation = Activation.NoEdit;
                band0.Columns["重設比"].CellActivation = Activation.NoEdit;
                band0.Columns["界限比"].CellActivation = Activation.NoEdit;
                band0.Columns["財務費用"].CellActivation = Activation.NoEdit;
                band0.Columns["張數"].CellActivation = Activation.NoEdit;
                band0.Columns["約當張數"].CellActivation = Activation.NoEdit;
                band0.Columns["額度結果"].CellActivation = Activation.NoEdit;
                band0.Columns["獎勵"].CellActivation = Activation.NoEdit;
                band0.Columns["可發"].CellActivation = Activation.NoEdit;
                band0.Columns["已發"].CellActivation = Activation.NoEdit;

                band0.Columns["交易員"].CellAppearance.BackColor = Color.White;
                band0.Columns["發行價格"].CellAppearance.BackColor = Color.White;
                band0.Columns["標的代號"].CellAppearance.BackColor = Color.White;
                band0.Columns["市場"].CellAppearance.BackColor = Color.White;
                //ultraGrid1.DisplayLayout.Bands[0].Columns["類型"].CellAppearance.BackColor = Color.White;
                band0.Columns["股價"].CellAppearance.BackColor = Color.White;
                band0.Columns["行使比例"].CellAppearance.BackColor = Color.White;

                band0.Columns["張數"].Hidden = false;
                band0.Columns["約當張數"].Hidden = false;
                band0.Columns["額度結果"].Hidden = false;

                band0.Columns["可發"].Hidden = true;
                band0.Columns["已發"].Hidden = true;
                band0.Columns["順位"].Hidden = true;

                toolStripButtonReload.Visible = true;
                toolStripButtonEdit.Visible = true;
                toolStripButtonConfirm.Visible = false;
                toolStripButtonCancel.Visible = false;

                if (GlobalVar.globalParameter.userGroup == "TR") {
                    toolStripButtonEdit.Visible = false;
                }

            }
        }

        private void toolStripButtonReload_Click(object sender, EventArgs e) {
            LoadData();
            toolStripLabel2.Text = DateTime.Now + "重新整理";
            GlobalUtility.LogInfo("Info", GlobalVar.globalParameter.userID + "重新整理");
        }

        private void toolStripButtonEdit_Click(object sender, EventArgs e) {
            isEdit = true;
            LoadData();
            back_dt.Rows.Clear();
            back_dt = dt.Copy();
            SetButton();
        }

        private void toolStripButtonConfirm_Click(object sender, EventArgs e) {
            ultraGrid1.PerformAction(Infragistics.Win.UltraWinGrid.UltraGridAction.ExitEditMode);
            isEdit = false;
            UpdateData();
            SetButton();
            LoadData();
        }

        private void toolStripButtonCancel_Click(object sender, EventArgs e) {
            isEdit = false;
            LoadData();
            SetButton();
        }

        private void ultraGrid1_InitializeLayout(object sender, InitializeLayoutEventArgs e) {
            ultraGrid1.DisplayLayout.Override.RowSelectorHeaderStyle = RowSelectorHeaderStyle.ColumnChooserButton;
        }

        private void ultraGrid1_InitializeRow(object sender, InitializeRowEventArgs e) {

            string serialNum = e.Row.Cells["序號"].Value.ToString();
            string is1500W = "N";
            is1500W = e.Row.Cells["1500W"].Value.ToString();
            string useReward = "N";
            useReward = e.Row.Cells["獎勵"].Value.ToString();
            if (is1500W == "Y")
                e.Row.Cells["1500W"].Appearance.ForeColor = Color.Blue;
            if (useReward == "Y")
            {
                e.Row.Cells["獎勵"].Appearance.BackColor = Color.PapayaWhip;
                e.Row.Cells["獎勵"].Appearance.ForeColor = Color.MediumBlue;
                e.Row.Cells["獎勵"].Appearance.FontData.Bold = Infragistics.Win.DefaultableBoolean.True;
            }
            string underlyingID = "";
            string cp = "C";
            underlyingID = e.Row.Cells["標的代號"].Value.ToString();
            cp = e.Row.Cells["CP"].Value.ToString();
            double vol = Convert.ToDouble(e.Row.Cells["IV"].Value.ToString());
            string issuable = "Y";
            string putIssuable = "Y";
            string toolTip1 = "發行檢查=N";

            string sqlTemp2 = "SELECT [Issuable], [PutIssuable] FROM [WarrantAssistant].[dbo].[WarrantUnderlyingSummary] WHERE UnderlyingID = '" + underlyingID + "'";
            DataView dvTemp2 = DeriLib.Util.ExecSqlQry(sqlTemp2, GlobalVar.loginSet.warrantassistant45);

            foreach (DataRowView drTemp2 in dvTemp2) {
                issuable = drTemp2["Issuable"].ToString();
                putIssuable = drTemp2["PutIssuable"].ToString();
            }
            if (underlyingID != "") {

                if (issuable == "N") {
                    e.Row.ToolTipText = toolTip1;
                    e.Row.Cells["標的代號"].Appearance.ForeColor = Color.Red;
                }

                if (cp == "P" && putIssuable == "N") {
                    e.Row.Cells["CP"].Appearance.ForeColor = Color.Red;
                    e.Row.Cells["CP"].ToolTipText = "Put not issuable";
                }

            }


            DataRow[] dt_LongTermSelect = dt_LongTerm.Select($@"UnderlyingID = '{underlyingID}'");

            if(dt_LongTermSelect.Length > 0)
            {
                e.Row.Cells["標的代號"].Appearance.BackColor = Color.Coral;
                e.Row.Cells["標的代號"].ToolTipText = "此標的穿VOL，要發長天期";

            }
            double issuePrice = e.Row.Cells["發行價格"].Value == DBNull.Value ? 0 : Convert.ToDouble(e.Row.Cells["發行價格"].Value);
            if (issuePrice <= 0.6 || issuePrice > 3) {
                e.Row.Cells["發行價格"].Appearance.ForeColor = Color.Red;
                e.Row.Cells["發行價格"].ToolTipText = " <= 0.6 or > 3";
            }

            string warrantType = e.Row.Cells["類型"].Value == DBNull.Value ? "一般型" : e.Row.Cells["類型"].Value.ToString();
            double k = e.Row.Cells["履約價"].Value == DBNull.Value ? 0 : Convert.ToDouble(e.Row.Cells["履約價"].Value);
            double underlyingPrice = e.Row.Cells["股價"].Value == DBNull.Value ? 0 : Convert.ToDouble(e.Row.Cells["股價"].Value);

            //Check for moneyness constraint
            e.Row.Cells["履約價"].Appearance.ForeColor = Color.Black;
            if (warrantType != "牛熊證") {
                if ((cp == "C" && k / underlyingPrice >= 1.5) || (cp == "P" && k / underlyingPrice <= 0.5)) {
                    e.Row.Cells["履約價"].Appearance.ForeColor = Color.Red;
                    e.Row.Cells["履約價"].ToolTipText = "履約價超過價外50%";
                }
            }
            DataRow[] HVsSelect = HVs.Select($@"UID = '{underlyingID}'");
            if (HVsSelect.Length > 0)
            {
                double hv20 = Convert.ToDouble(HVsSelect[0][1].ToString());
                double hv60 = Convert.ToDouble(HVsSelect[0][2].ToString());
                if (vol < hv20 * 0.7 || vol < hv60 * 0.7)
                {
                    e.Row.Cells["IV"].Appearance.ForeColor = Color.Red;
                    e.Row.Cells["IV"].ToolTipText = "VOL低於HV的0.7倍";
                }
            }
            if (!isEdit && DateTime.Now.TimeOfDay.TotalMinutes >= GlobalVar.globalParameter.resultTime) {
                string warrantName = e.Row.Cells["權證名稱"].Value.ToString();
                string traderID = e.Row.Cells["交易員"].Value.ToString();
                string applyStatus = "";
                double issueNum = 0.0;
                string applyTime = "";
                string apytime = "";
                string rank = "";
                string oriapplyTime = "";
                double result = 0;
                issueNum = Convert.ToDouble(e.Row.Cells["張數"].Value);
                double rewardNum = Convert.ToDouble(e.Row.Cells["獎勵額度"].Value);
                double equivalentNum = Convert.ToDouble(e.Row.Cells["約當張數"].Value);
                //double result = e.Row.Cells["額度結果"].Value == DBNull.Value ? 0 : Convert.ToDouble(e.Row.Cells["額度結果"].Value);
                double cr = Convert.ToDouble(e.Row.Cells["行使比例"].Value);

                string sqlTemp = "SELECT [ApplyStatus],[ApplyTime], [OriApplyTime],[Result] FROM [WarrantAssistant].[dbo].[Apply_71] WHERE SerialNum = '" + serialNum + "'";
                DataView dvTemp = DeriLib.Util.ExecSqlQry(sqlTemp, GlobalVar.loginSet.warrantassistant45);

                foreach (DataRowView drTemp in dvTemp) {
                    applyStatus = drTemp["ApplyStatus"].ToString();
                    applyTime = drTemp["ApplyTime"].ToString().Substring(0, 2);//時間幾點
                    apytime = drTemp["ApplyTime"].ToString();//時間的全部字串
                    rank = drTemp["ApplyTime"].ToString().Substring(6, 2);
                    oriapplyTime = drTemp["OriApplyTime"].ToString();
                    result = Convert.ToDouble(drTemp["Result"].ToString());
                }
                if (applyStatus == "排隊中" && issueNum != 10000) {
                    //e.Row.Cells["張數"].Appearance.ForeColor = Color.Red;
                    e.Row.Cells["張數"].ToolTipText = "排隊中";
                }

                if (applyStatus == "X 沒額度") {
                    e.Row.Cells["權證名稱"].Appearance.BackColor = Color.LightGray;
                }
                
                if ((result + 0.00001).CompareTo(equivalentNum) >=0) {
                    e.Row.Cells["權證名稱"].Appearance.BackColor = Color.PaleGreen;
                    e.Row.Cells["權證名稱"].ToolTipText = "額度OK";
                }
                if(useReward == "Y" && equivalentNum <= rewardNum)
                {
                    e.Row.Cells["權證名稱"].Appearance.BackColor = Color.LightPink;
                    e.Row.Cells["權證名稱"].ToolTipText = "額度OK";
                }
                /*
                if ((result + 0.00001).CompareTo(equivalentNum) < 0 && result > 0) {
                    e.Row.Cells["權證名稱"].Appearance.BackColor = Color.PaleTurquoise;
                    e.Row.Cells["權證名稱"].ToolTipText = "部分額度";
                }
                */
                if (apytime.Length > 0)
                {
                    //要搶
                    if (oriapplyTime.Length > 0 && applyTime != "22")
                    {
                        e.Row.Cells["順位"].Appearance.BackColor = Color.Bisque;
                        e.Row.Cells["順位"].Value = apytime.Substring(6, 2);
                        e.Row.Cells["順位"].ToolTipText = "需搶額度";
                      
                        if ((result + 0.00001).CompareTo(equivalentNum) >= 0)
                        {
                            e.Row.Cells["狀態"].Value = "需搶額度，有搶到";
                            e.Row.Cells["權證名稱"].Appearance.BackColor = Color.Gold;
                        }
                        else
                        {
                            if (UidCnt[underlyingID] > 1)
                            {
                                e.Row.Cells["狀態"].Value = "需搶額度，只有部分額度，有另一檔搶到，小心誤刪";
                                e.Row.Cells["權證名稱"].Appearance.BackColor = Color.LightGray;
                            }
                            else
                            {
                                e.Row.Cells["狀態"].Value = "需搶額度，只有部分額度";
                                e.Row.Cells["權證名稱"].Appearance.BackColor = Color.LightGray;
                            }
                        }
                        if (!wname2result.ContainsKey(warrantName))
                            wname2result.Add(warrantName, result);
                        
                        if((wname2result[warrantName]==0) && (result != 0)&& (userID.TrimStart('0')==traderID))
                        {
                            MessageBox.Show($"{warrantName} 有額度 {result.ToString()} 張");
                        }
                        

                    }
                    else if (applyTime == "22")
                    {
                        if (UidCnt[underlyingID] > 1)
                        {
                            e.Row.Cells["順位"].Appearance.BackColor = Color.Aquamarine;
                            e.Row.Cells["順位"].Value = "X";
                            e.Row.Cells["順位"].ToolTipText = "沒額度";
                            e.Row.Cells["狀態"].Value = "需搶額度，沒搶到，有另一檔搶到，小心誤刪";
                            e.Row.Cells["權證名稱"].Appearance.BackColor = Color.LightGray;
                        }
                        else
                        {
                            e.Row.Cells["順位"].Appearance.BackColor = Color.Aquamarine;
                            e.Row.Cells["順位"].Value = "X";
                            e.Row.Cells["順位"].ToolTipText = "沒額度";
                            e.Row.Cells["狀態"].Value = "需搶額度，沒搶到";
                            e.Row.Cells["權證名稱"].Appearance.BackColor = Color.LightGray;
                        }
                    }
                    else
                    {
                        e.Row.Cells["順位"].Appearance.BackColor = Color.Aquamarine;
                        e.Row.Cells["順位"].Value = "-";
                        e.Row.Cells["順位"].ToolTipText = "不用搶";
                        e.Row.Cells["狀態"].Value = "不用搶";
                    }
                }
                else
                {
                    
                    if (Iscompete.Contains(underlyingID) && (useReward == "N"))
                    {
                        e.Row.Cells["狀態"].ToolTipText = "搶發沒額度，從7-1表拿掉";
                    }
                    else
                    {
                        int h1 = 0;
                        int m1 = 0;
                       
                        string sql_MDate = $@"SELECT Min([UpdateTime]) AS MDate
                                            FROM [WarrantAssistant].[dbo].[ApplyTotalRecord]
                                            WHERE [SerialNumber] = '{serialNum}' AND [UpdateCount] = 0";
                        //申請時間如果在10:30申請前的最後一刻之後，則視為後來加發
                        DataTable dt_MDate = MSSQL.ExecSqlQry(sql_MDate, GlobalVar.loginSet.warrantassistant45);

                        foreach (DataRow dr in dt_MDate.Rows)
                        {
                            bool f = DateTime.TryParse(dr["MDate"].ToString(), out DateTime t1);
                            h1 = Convert.ToInt32(t1.ToString("yyyyMMddHHmmss").Substring(8, 2));
                            m1 = Convert.ToInt32(t1.ToString("yyyyMMddHHmmss").Substring(10, 2));
                        }
                        if ((h1 > applyhour_compete) || (h1 == applyhour_compete) && (m1 > applymin_compete))
                        {
                            e.Row.Cells["狀態"].ToolTipText = "加發，沒在7-1表中";
                        }
                        else if (useReward == "Y")//在早上第一批就有發，但後來改成獎勵
                        {
                            e.Row.Cells["狀態"].ToolTipText = "改用獎勵發，沒在7-1表中";
                        }
                        else
                        {
                            e.Row.Cells["狀態"].ToolTipText = "後來不發";
                        }
                    }
                    e.Row.Cells["順位"].Appearance.BackColor = Color.Aquamarine;
                    e.Row.Cells["順位"].Value = "-";
                    e.Row.Cells["順位"].ToolTipText = "不用搶";
                }
                /*
                if (applyTime == "10" && oriapplyTime.Length > 0 && issueNum != 10000 && useReward == "N")
                {
                    e.Row.Cells["張數"].Appearance.BackColor = Color.Bisque;
                    e.Row.Cells["張數"].ToolTipText = "要搶，更改為10000張";
                    e.Row.Cells["張數"].Value = 10000;
                }
                */
                if ((changeTo_1w.Contains(serialNum) && useReward == "N") || (Iscompete.Contains(underlyingID) && useReward == "Y"))
                {
                    if (UidDeltaOne.ContainsKey(underlyingID))
                    {
                        int wlength = warrantName.Length;
                        string subwname = warrantName.Substring(0, wlength - 2);

                        string sql_price = $@"SELECT [MPrice]
                                            FROM [WarrantAssistant].[dbo].[WarrantPrices]
                                            WHERE [CommodityID] = '{underlyingID}'";
                        DataTable dt_price = MSSQL.ExecSqlQry(sql_price, GlobalVar.loginSet.warrantassistant45);

                        double spot = 0;
                        foreach (DataRow dr in dt_price.Rows)
                        {
                            spot = Convert.ToDouble(dr["MPrice"].ToString());
                        }

                        if (subwname.EndsWith("購") || subwname.EndsWith("牛"))//如果是Call  要考慮Put=0的情況 
                        {
                            if (IsSpecial.Contains(underlyingID))
                            {
                                if (Market30.Contains(underlyingID))//市值前30  DeltaOne*股價<5億
                                {
                                    if (cr * issueNum * spot > 500000)
                                    {

                                        e.Row.Cells["行使比例"].Appearance.BackColor = Color.DimGray;
                                        e.Row.Cells["行使比例"].Appearance.ForeColor = Color.Gold;
                                        e.Row.Cells["行使比例"].ToolTipText = "為風險標的且為市值前30大標的，DeltaOne市值已超過5億\n";
                                    }
                                }
                                else
                                {
                                    if (cr * issueNum * spot > 200000)
                                    {
                                        e.Row.Cells["行使比例"].Appearance.BackColor = Color.DimGray;
                                        e.Row.Cells["行使比例"].Appearance.ForeColor = Color.Gold;
                                        e.Row.Cells["行使比例"].ToolTipText = "為風險標的 DeltaOne市值已超過2億\n";
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
                                    if (cr * issueNum * spot > ISTOP30MaxIssue)
                                    {
                                        e.Row.Cells["行使比例"].Appearance.BackColor = Color.DimGray;
                                        e.Row.Cells["行使比例"].Appearance.ForeColor = Color.Gold;
                                        e.Row.Cells["行使比例"].ToolTipText = $@"為風險標的且為市值前30大標的，DeltaOne市值已超過{(int)(ISTOP30MaxIssue / 100000)}億\n";
                                    }
                                }
                                else
                                {
                                    if (cr * issueNum * spot > NonTOP30MaxIssue)
                                    {
                                        e.Row.Cells["行使比例"].Appearance.BackColor = Color.DimGray;
                                        e.Row.Cells["行使比例"].Appearance.ForeColor = Color.Gold;
                                        e.Row.Cells["行使比例"].ToolTipText = $@"為風險標的 DeltaOne市值已超過{(int)(NonTOP30MaxIssue / 100000)}億\n";
                                    }
                                }
                            }
                            if(IsSpecial.Contains(underlyingID) && (double)UidDeltaOne[underlyingID].KgiCallDeltaOne / (double)UidDeltaOne[underlyingID].KgiPutDeltaOne < SpecialCallPutRatio)
                            {
                                e.Row.Cells["行使比例"].Appearance.BackColor = Color.DimGray;
                                e.Row.Cells["行使比例"].Appearance.ForeColor = Color.Gold;
                                e.Row.Cells["行使比例"].ToolTipText += $@"為風險標的，自家權證 Call/Put DeltaOne比例 < {SpecialCallPutRatio}\n";
                            }
                            else if ((double)UidDeltaOne[underlyingID].KgiCallDeltaOne / (double)UidDeltaOne[underlyingID].KgiPutDeltaOne < NonSpecialCallPutRatio)
                            {
                                if (!IsIndex.Contains(underlyingID))
                                {
                                    e.Row.Cells["行使比例"].Appearance.BackColor = Color.DimGray;
                                    e.Row.Cells["行使比例"].Appearance.ForeColor = Color.Gold;
                                    e.Row.Cells["行使比例"].ToolTipText += $@"自家權證 Call/Put DeltaOne比例 < {NonSpecialCallPutRatio}\n";
                                }

                            }
                            if (IsSpecial.Contains(underlyingID) && UidDeltaOne[underlyingID].AllPutDeltaOne > 0 && UidDeltaOne[underlyingID].KgiAllPutRatio > SpecialKGIALLPutRatio)
                            {
                                //若之前這檔標的沒發過Put可以跳過，可是要考慮今天發超過一檔
                                if (UidDeltaOne[underlyingID].KgiPutNum > 1)
                                {
                                    e.Row.Cells["行使比例"].Appearance.BackColor = Color.DimGray;
                                    e.Row.Cells["行使比例"].Appearance.ForeColor = Color.Gold;
                                    e.Row.Cells["行使比例"].ToolTipText += $@"自家/市場 Put DeltaOne比例 > {SpecialKGIALLPutRatio}\n";
                                }
                            }
                        }
                    }

                }
            }
            
            if ((DateTime.Now.TimeOfDay.TotalMinutes >= 660) && dt_record.ContainsKey(serialNum))
            {//11點後再變顏色
                
                foreach (string key in sqlTogrid.Keys)
                {
                    if (dt_record[serialNum].ContainsKey(sqlTogrid[key]))
                    {
                        //if (key != "行使比例" && e.Row.Cells[key].Appearance.BackColor != Color.DimGray)
                        if (key != "重設比" && e.Row.Cells[key].Appearance.BackColor != Color.LightPink)
                        {
                            e.Row.Cells[key].Appearance.BackColor = Color.LightPink;
                            e.Row.Cells[key].ToolTipText += "(" + dt_record[serialNum][sqlTogrid[key]] + " BY " + dt_recordByTrader[serialNum] + ")";
                        }
                    }
                }
                
            }

        }

        private void ultraGrid1_AfterCellUpdate(object sender, CellEventArgs e) {
            if (e.Cell.Column.Key != "交易員" && e.Cell.Column.Key != "權證名稱" && e.Cell.Column.Key != "發行價格" && e.Cell.Column.Key != "標的代號" && e.Cell.Column.Key != "市場" && e.Cell.Column.Key != "1500W") {
                double price = 0.0;

                double underlyingPrice = 0.0;
                underlyingPrice = e.Cell.Row.Cells["股價"].Value == DBNull.Value ? 0 : Convert.ToDouble(e.Cell.Row.Cells["股價"].Value);
                double k = 0.0;
                k = e.Cell.Row.Cells["履約價"].Value == DBNull.Value ? 0 : Convert.ToDouble(e.Cell.Row.Cells["履約價"].Value);
                int t = 0;
                t = e.Cell.Row.Cells["期間"].Value == DBNull.Value ? 0 : Convert.ToInt32(e.Cell.Row.Cells["期間"].Value);
                double cr = 0.0;
                cr = e.Cell.Row.Cells["行使比例"].Value == DBNull.Value ? 0 : Convert.ToDouble(e.Cell.Row.Cells["行使比例"].Value);
                double vol = 0.0;
                vol = e.Cell.Row.Cells["IV"].Value == DBNull.Value ? 0 : Convert.ToDouble(e.Cell.Row.Cells["IV"].Value) / 100;
                double resetR = 0.0;
                resetR = e.Cell.Row.Cells["重設比"].Value == DBNull.Value ? 0 : Convert.ToDouble(e.Cell.Row.Cells["重設比"].Value) / 100;
                double financialR = 0.0;
                financialR = e.Cell.Row.Cells["財務費用"].Value == DBNull.Value ? 0 : Convert.ToDouble(e.Cell.Row.Cells["財務費用"].Value) / 100;
                string warrantType = "一般型";
                warrantType = e.Cell.Row.Cells["類型"].Value == DBNull.Value ? "一般型" : e.Cell.Row.Cells["類型"].Value.ToString();
                string cpType = "C";
                cpType = e.Cell.Row.Cells["CP"].Value == DBNull.Value ? "C" : e.Cell.Row.Cells["CP"].Value.ToString();
                string underlyingID = e.Cell.Row.Cells["標的代號"].Value.ToString();
                if (warrantType != "一般型" && warrantType != "牛熊證" && warrantType != "重設型") {
                    if (warrantType == "2")
                        warrantType = "牛熊證";
                    else if (warrantType == "3")
                        warrantType = "重設型";
                    else
                        warrantType = "一般型";
                }

                if (cpType != "C" && cpType != "P") {
                    if (cpType == "2")
                        cpType = "P";
                    else
                        cpType = "C";
                }

                CallPutType cp = CallPutType.Call;
                if (cpType == "P")
                    cp = CallPutType.Put;
                else
                    cp = CallPutType.Call;
                if (underlyingID.Length > 4 && underlyingID.Substring(0, 2) != "00")
                {
                    if (warrantType == "牛熊證")
                        price = Pricing.BullBearWarrantPrice(cp, underlyingPrice, resetR, GlobalVar.globalParameter.interestRate_Index, vol, t, financialR, cr);
                    else if (warrantType == "重設型")
                        price = Pricing.ResetWarrantPrice(cp, underlyingPrice, resetR, GlobalVar.globalParameter.interestRate_Index, vol, t, cr);
                    else
                        price = Pricing.NormalWarrantPrice(cp, underlyingPrice, k, GlobalVar.globalParameter.interestRate_Index, vol, t, cr);
                }
                else
                {
                    if (warrantType == "牛熊證")
                        price = Pricing.BullBearWarrantPrice(cp, underlyingPrice, resetR, GlobalVar.globalParameter.interestRate, vol, t, financialR, cr);
                    else if (warrantType == "重設型")
                        price = Pricing.ResetWarrantPrice(cp, underlyingPrice, resetR, GlobalVar.globalParameter.interestRate, vol, t, cr);
                    else
                        price = Pricing.NormalWarrantPrice(cp, underlyingPrice, k, GlobalVar.globalParameter.interestRate, vol, t, cr);
                }
                /*e.Cell.Row.Cells["履約價"].Appearance.ForeColor = Color.Black;
                if (warrantType != "牛熊證") {
                    if (cpType == "C" && k / underlyingPrice >= 1.5) {
                        e.Cell.Row.Cells["履約價"].Appearance.ForeColor = Color.Red;
                    } else if (cpType == "P" && k / underlyingPrice <= 0.5) {
                        e.Cell.Row.Cells["履約價"].Appearance.ForeColor = Color.Red;
                    }
                }*/

                double shares = 0.0;
                shares = e.Cell.Row.Cells["張數"].Value == DBNull.Value ? 10000 : Convert.ToDouble(e.Cell.Row.Cells["張數"].Value);
                /*
                string is1500W = "N";
                is1500W = e.Cell.Row.Cells["1500W"].Value == DBNull.Value ? "N" : (string)e.Cell.Row.Cells["1500W"].Value;
                if (e.Cell.Column.Key == "1500W" && is1500W=="Y")
                {
                    double totalValue = 0.0;
                    totalValue = price * shares * 1000;
                    while (totalValue < 15000000)
                    {
                        vol += 0.01;
                        if (warrantType == "牛熊證")
                            price = Pricing.BullBearWarrantPrice(cp, underlyingPrice, resetR, GlobalVar.globalParameter.interestRate, vol, t, financialR, cr);
                        else if (warrantType == "重設型")
                            price = Pricing.ResetWarrantPrice(cp, underlyingPrice, resetR, GlobalVar.globalParameter.interestRate, vol, t, cr);
                        else
                            price = Pricing.NormalWarrantPrice(cp, underlyingPrice, k, GlobalVar.globalParameter.interestRate, vol, t, cr);
                        totalValue = price * shares * 1000;
                    }
                    e.Cell.Row.Cells["IV"].Value = Math.Round(vol * 100, 0);
                }
                 * */

                e.Cell.Row.Cells["發行價格"].Value = Math.Round(price, 2);
            }

            string is1500W = "N";
            is1500W = e.Cell.Row.Cells["1500W"].Value.ToString();
            if (e.Cell.Column.Key == "1500W") {
                if (is1500W == "N")
                    e.Cell.Row.Cells["IV"].Value = e.Cell.Row.Cells["IVOri"].Value;
            }
        }

        private void ultraGrid1_DoubleClickCell(object sender, DoubleClickCellEventArgs e) {
            if (e.Cell.Column.Key == "標的代號")
                GlobalUtility.MenuItemClick<FrmIssueCheck>().SelectUnderlying((string) e.Cell.Value);

            if (e.Cell.Column.Key == "CP")
                GlobalUtility.MenuItemClick<FrmIssueCheckPut>().SelectUnderlying((string) e.Cell.Row.Cells["標的代號"].Value);


            if (e.Cell.Column.Key == "1500W" || e.Cell.Column.Key == "類型" || e.Cell.Column.Key == "CP" || e.Cell.Column.Key == "期間" || e.Cell.Column.Key == "行使比例" || e.Cell.Column.Key == "重設比" || e.Cell.Column.Key == "界限比" || e.Cell.Column.Key == "張數")
            {
                string serialNum = e.Cell.Row.Cells["序號"].Value.ToString();
                string key = e.Cell.Column.Key.ToString();
                if (dt_record.ContainsKey(serialNum))
                {
                    if (dt_record[serialNum].ContainsKey(sqlTogrid[key]))
                    {
                        string fromValue = dt_record[serialNum][sqlTogrid[key]];
                        string toValue = e.Cell.Value.ToString();
                        DialogResult dialogResult = MessageBox.Show($@"確認變更 {key} {fromValue}至{toValue}?", "確認", MessageBoxButtons.YesNo);
                        if (dialogResult == DialogResult.Yes)
                        {
                            //[ApplyKind] = 4 代表已確認資料變更
                            string sqlUpdate = $@"UPDATE [WarrantAssistant].[dbo].[ApplyTotalRecord] SET [ApplyKind] = '4' WHERE [DataName] = '{sqlTogrid[key]}' AND [UpdateTime] >= CONVERT(varchar, getdate(), 112) AND [SerialNumber] = '{serialNum}' AND [UpdateCount] > 0";
                            MSSQL.ExecSqlCmd(sqlUpdate, GlobalVar.loginSet.warrantassistant45);
                            LoadData();

                        }
                    }
                }
            }


            /*
            sqlTogrid.Add("1500W", "Apply1500W");
            sqlTogrid.Add("類型", "Type");
            sqlTogrid.Add("CP", "CP");
            
            sqlTogrid.Add("期間", "T");
            sqlTogrid.Add("行使比例", "CR");
            
            sqlTogrid.Add("重設比", "ResetR");
            sqlTogrid.Add("界限比", "BarrierR");
            sqlTogrid.Add("張數", "IssueNum");
            */
        }

        private void toolStripUpDateCheck_Click(object sender, EventArgs e)
        {
            DialogResult dialogResult = MessageBox.Show($@"確認全部變更?", "確認", MessageBoxButtons.YesNo);
            if (dialogResult == DialogResult.Yes)
            {
                //[ApplyKind] = 4 代表已確認資料變更
                foreach(string serialNum in dt_record.Keys)
                {
                    foreach(string key in dt_record[serialNum].Keys)
                    {
                        string sqlUpdate = $@"UPDATE [WarrantAssistant].[dbo].[ApplyTotalRecord] SET [ApplyKind] = '4' WHERE [DataName] = '{key}' AND [UpdateTime] >= CONVERT(varchar, getdate(), 112) AND [SerialNumber] = '{serialNum}' AND [UpdateCount] > 0";
                        MSSQL.ExecSqlCmd(sqlUpdate, GlobalVar.loginSet.warrantassistant45);
                    }
                }
                LoadData();
            }
        }
    }
}

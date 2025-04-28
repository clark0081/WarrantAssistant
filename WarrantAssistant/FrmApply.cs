#define deletelog
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
using System.Text.RegularExpressions;
using System.Threading;
using System.IO;
using System.Configuration;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop;



namespace WarrantAssistant
{
    
    public partial class FrmApply:Form
    {
        //HeaderUIElement
        DateTime today = DateTime.Today;
        DateTime lastday = EDLib.TradeDate.LastNTradeDate(1);


        public SqlConnection conn = new SqlConnection(GlobalVar.loginSet.warrantassistant45);
        private DataTable dt = new DataTable();
        private DataTable updaterecord_dt = new DataTable();//當有更改資料時，用來記錄那些欄位更改
        Dictionary<string, string> sqlTogrid = new Dictionary<string, string>();
        private bool isEdit = false;
        public string userID = GlobalVar.globalParameter.userID;
        public string userName = GlobalVar.globalParameter.userName;
        private int applyCount = 0;
        Dictionary<string, UidPutCallDeltaOne> UidDeltaOne = new Dictionary<string, UidPutCallDeltaOne>();//紀錄昨日的DeltaOne
        Dictionary<string, UidPutCallDeltaOne> UidDeltaOne_Temp = new Dictionary<string, UidPutCallDeltaOne>();//紀錄昨日DeltaOne加上擬發行的DeltaOne
        //Dictionary<string, UidPutCallDeltaOne> UidDeltaOne_Temp;
        List<string> IsSpecial = new List<string>();//特殊標的
        List<string> IsIndex = new List<string>();//臺灣50指數,臺灣中型100指數,櫃買富櫃50指
        List<string> Market30 = new List<string>();//市值前30大
        List<string> WMarket30 = new List<string>();//市值前30大標的的權證
        Dictionary<string, DateTime> Reduction = new Dictionary<string, DateTime>();//減資
        List<int> DeletedSerialNum = new List<int>();
        public List<string> Iscompete = new List<string>();//要搶的標的
        //Dictionary<string, Dictionary<string, double>> SuggestVol = new Dictionary<string, Dictionary<string, double>>();
        //Dictionary<string, Dictionary<string, string>> SuggestVolResult = new Dictionary<string, Dictionary<string, string>>();
        Dictionary<string, double> canIssue = new Dictionary<string, double>();//額度
        //DataTable SuggestVol_C;
        //DataTable SuggestVol_P;
        //DataTable VolManager;
        DataTable dtKK;
        DataTable dtTT;
        DataTable dtPL20D;
        DataTable dtPLYear;
        DataTable dt_LongTerm;
        private static object _thisLock = new object();
        private Thread thread1;
        public double NonSpecialCallPutRatio = Convert.ToDouble(ConfigurationManager.AppSettings["NonSpecialCallPutRatio"].ToString());
        public double SpecialCallPutRatio = Convert.ToDouble(ConfigurationManager.AppSettings["SpecialCallPutRatio"].ToString());
        public double SpecialKGIALLPutRatio = Convert.ToDouble(ConfigurationManager.AppSettings["SpecialKGIALLPutRatio"].ToString());
        public double ISTOP30MaxIssue = Convert.ToDouble(ConfigurationManager.AppSettings["ISTOP30MaxIssue"].ToString());
        public double NonTOP30MaxIssue = Convert.ToDouble(ConfigurationManager.AppSettings["NonTOP30MaxIssue"].ToString());
        private Dictionary<string,string> IssuedSerialNum = new Dictionary<string, string>();//方便行政知道那些權證被刪掉
        private Dictionary<string, string> IssuedSerialNum_UID = new Dictionary<string, string>();//方便行政知道那些權證的標的代號被刪掉

        private Dictionary<string, double> AvailableShares = new Dictionary<string, double>();
        private Dictionary<string, double> RewardShares = new Dictionary<string, double>();

        // 區域標的
        private List<string> IsTWUid = new List<string>();
        //紀錄發行天數對應TtoM
        private Dictionary<int, int> T2TtoM = new Dictionary<int, int>();

        DataTable dtVolRatio = new DataTable();
        public FrmApply() {
            InitializeComponent();
        }

        private void FrmApply_Load(object sender, EventArgs e) {
            
            toolStripLabel1.Text = "使用者: " + userName;
            //預設重複權證檢查值
            toolStripTextBox1.Text = "20";
            toolStripTextBox2.Text = "1.5";
            InitialGrid();
            //紀錄資料庫對照datagrid欄位
            //sqlTogrid.Add("1500W", "Apply1500W");
            sqlTogrid.Add("類型", "Type");
            sqlTogrid.Add("CP", "CP");
            //sqlTogrid.Add("履約價", "K");
            sqlTogrid.Add("期間(月)", "T");
            sqlTogrid.Add("行使比例", "CR");
            //sqlTogrid.Add("IV", "IV");
            sqlTogrid.Add("重設比", "ResetR");
            sqlTogrid.Add("界限比", "BarrierR");
            sqlTogrid.Add("張數", "IssueNum");
            sqlTogrid.Add("獎勵", "UseReward");
            LoadT2TtoM();
            LoadData();
            SetUpdateRecord();
            //thread1 = new Thread(new ThreadStart(Load_10min));
            //thread1.Start();
            //load 建議Vol
            LoadCanIssue();
            //LoadSuggestVol();
            LoadIsSpecial();
            LoadIsIndex();
            LoadMarket30();
            LoadDeletedSerial();
            LoadIssuedSerial();
            LoadReduction();
        }
        #region   Load_10min() 一支thread負責每十分鐘更新資料
        private void Load_10min()
        {
            while (DateTime.Now.TimeOfDay.TotalMinutes >= GlobalVar.globalParameter.resultTime)
            {
                Thread.Sleep(600000);
                LoadData();
                Thread.Sleep(600000);//10min
               
            }
        }


        
        #endregion

        private void SetUpdateRecord()
        {
            updaterecord_dt.Columns.Add("serialNum", typeof(string));
            updaterecord_dt.Columns.Add("dataname", typeof(string));
            updaterecord_dt.Columns.Add("fromvalue", typeof(string));
            updaterecord_dt.Columns.Add("tovalue", typeof(string));
        }
        private void LoadIssuedSerial()
        {
            IssuedSerialNum.Clear();
            IssuedSerialNum_UID.Clear();
            string sqlIssued = $@"SELECT [SerialNum], [WarrantName],[UnderlyingID]  
                                  FROM [WarrantAssistant].[dbo].[ApplyTotalList]
                                  WHERE [UserID] = '{userID}' AND [ApplyKind] = '1'";
            DataTable dt = MSSQL.ExecSqlQry(sqlIssued, GlobalVar.loginSet.warrantassistant45);

            foreach (DataRow dr in dt.Rows)
            {
                string serial = dr["SerialNum"].ToString();
                string wname = dr["WarrantName"].ToString();
                string uid = dr["UnderlyingID"].ToString();
                IssuedSerialNum.Add(serial,wname);
                IssuedSerialNum_UID.Add(serial, uid);
            }
        }

        private void LoadCanIssue()
        {
            canIssue.Clear();
            string sql = $@"SELECT  [UID],Floor([CanIssue]) AS CanIssue
                                      FROM [WarrantAssistant].[dbo].[WarrantUnderlyingCreditNew]
                                      WHERE [UpdateTime] > '{DateTime.Today.ToString("yyyyMMdd")}'";
            DataTable dt = MSSQL.ExecSqlQry(sql, GlobalVar.loginSet.warrantassistant45);
            foreach(DataRow dr in dt.Rows)
            {
                string uid = dr["UID"].ToString();
                double canissue = Convert.ToDouble(dr["CanIssue"].ToString());
                canIssue.Add(uid, canissue);
            }
        }
        private void LoadT2TtoM()
        {
            T2TtoM.Clear();
            string listedDate = EDLib.TradeDate.NextNTradeDate(3).ToString("yyyyMMdd");

            string sql = $@"SELECT CONVERT(VARCHAR,TradeDate,112) AS TDate ,ROW_NUMBER() OVER(ORDER BY TradeDate) AS TtoM　FROM [DeriPosition].[dbo].[Calendar] WHERE IsTrade='Y' AND [CountryId] = 'TWN' AND [TradeDate] >= '{listedDate}'";
            DataTable dt = MSSQL.ExecSqlQry(sql, GlobalVar.loginSet.tsquoteSqlConnString);
            for(int i  = 6 ; i <= 30 ; i++)
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
                if(dr.Length > 0)
                {
                    int ttom = Convert.ToInt32(dr[0][1].ToString());
                    if (!T2TtoM.ContainsKey(i))
                        T2TtoM.Add(i, ttom);
                    
                }
            }
        }
        private void LoadReduction()
        {
            Reduction.Clear();
            string sql = $@"SELECT  [UID], CONVERT(VARCHAR,[StartTradingDate],112) AS [StartTradingDate]
                                    FROM [WarrantAssistant].[dbo].[ReductionCapital]";
            DataTable dt = MSSQL.ExecSqlQry(sql, GlobalVar.loginSet.warrantassistant45);

            foreach (DataRow dr in dt.Rows)
            {

                string uid = dr["UID"].ToString();

                DateTime st = DateTime.ParseExact(dr["StartTradingDate"].ToString(), "yyyyMMdd", null);
                if(!Reduction.ContainsKey(uid))
                    Reduction.Add(uid, st);
            }

        }
        /*
        private void LoadSuggestVol()
        {

            string sql = $@"SELECT  [UID], [CP], [IV_Rec], [Result]
                          FROM [WarrantAssistant].[dbo].[RecommendVol]
                          WHERE [DateTime] >='{today.ToString("yyyyMMdd")}'
                          ORDER BY [UID]";
            DataTable dt = MSSQL.ExecSqlQry(sql, GlobalVar.loginSet.warrantassistant45);
            string sql_SuggestVolC = $@"SELECT  [UID], [CP], [IV_Rec], [Result], [HV_60D_Spread], [HV_60D] 
                          FROM [WarrantAssistant].[dbo].[RecommendVol]
                          WHERE [DateTime] >='{today.ToString("yyyyMMdd")}' AND [CP] ='C'
                          ORDER BY [UID]";
            SuggestVol_C = MSSQL.ExecSqlQry(sql_SuggestVolC, GlobalVar.loginSet.warrantassistant45);
            string sql_SuggestVolP = $@"SELECT  [UID], [CP], [IV_Rec], [Result], [HV_60D_Spread], [HV_60D] 
                          FROM [WarrantAssistant].[dbo].[RecommendVol]
                          WHERE [DateTime] >='{today.ToString("yyyyMMdd")}' AND [CP] ='P'
                          ORDER BY [UID]";
            SuggestVol_P = MSSQL.ExecSqlQry(sql_SuggestVolP, GlobalVar.loginSet.warrantassistant45);
            string sql_VolManager = $@"SELECT  DISTINCT [STKID],[HV_60D]
                                        FROM [10.101.10.5].[WMM3].[dbo].[VolManagerDetail]
			                            where [UPDATETIME] >='{lastday.ToString("yyyyMMdd")} 14:00' AND [UPDATETIME] <'{today.ToString("yyyyMMdd")}'";
            VolManager = MSSQL.ExecSqlQry(sql_VolManager, GlobalVar.loginSet.warrantassistant45);

            foreach (DataRow dr in dt.Rows)
            {
                string uid = dr["UID"].ToString();
                string cp = dr["CP"].ToString();
                double vol = Convert.ToDouble(dr["IV_Rec"].ToString());
                string result = dr["Result"].ToString();
                if (!SuggestVol.ContainsKey(uid))
                {
                    SuggestVol.Add(uid, new Dictionary<string, double>());
                    SuggestVolResult.Add(uid, new Dictionary<string, string>());
                }
                if (!SuggestVol[uid].ContainsKey(cp))
                {
                    SuggestVol[uid].Add(cp, vol);
                    SuggestVolResult[uid].Add(cp, result);
                }
            }
        }
        */
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
        private void LoadIsCompete()
        {
            Iscompete.Clear();

            string sql_compete = $@"SELECT [UnderlyingID]
                          FROM [WarrantAssistant].[dbo].[Apply_71]
                          WHERE len([OriApplyTime])> 0";
            DataTable dt_compete = MSSQL.ExecSqlQry(sql_compete, GlobalVar.loginSet.warrantassistant45);
            foreach (DataRow dr in dt_compete.Rows)
            {
                Iscompete.Add(dr["UnderlyingID"].ToString());
            }
        }
        private void LoadDeletedSerial()
        {
            DeletedSerialNum.Clear();

            string sqlDelSerial = $@"SELECT [SerialNum] FROM [WarrantAssistant].[dbo].[TempListDeleteLog]
                                    WHERE [Trader] ='{userID}' AND [DateTime] >='{today.ToString("yyyyMMdd")}'";
            DataTable dtDelSerial = MSSQL.ExecSqlQry(sqlDelSerial, GlobalVar.loginSet.warrantassistant45);

            foreach (DataRow dr in dtDelSerial.Rows)
            {
                string serialNum = dr["SerialNum"].ToString();
                if (serialNum.Length == 0)
                    continue;
                int index = Convert.ToInt32(serialNum.Substring(17, serialNum.Length - 17));
                string d = (serialNum.Substring(0, 8));
                if (d == DateTime.Today.ToString("yyyyMMdd")) 
                    DeletedSerialNum.Add(index);
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
        private void InitialGrid() {
            dt.Columns.Add("刪除", typeof(bool));
            dt.Columns["刪除"].ReadOnly = false;
           
            
            dt.Columns.Add("編號", typeof(string));
            dt.Columns.Add("標的代號", typeof(string));
            dt.Columns.Add("履約價", typeof(double));
            dt.Columns.Add("行使比例", typeof(double));
            dt.Columns.Add("建議行使比例", typeof(double));
            dt.Columns.Add("HV", typeof(double));
            dt.Columns.Add("IV", typeof(double));
            //dt.Columns.Add("建議Vol", typeof(double));
            //dt.Columns.Add("下限Vol", typeof(double));
            dt.Columns.Add("期間(月)", typeof(int));
            dt.Columns.Add("張數", typeof(double));
            dt.Columns.Add("類型", typeof(string));
            dt.Columns.Add("CP", typeof(string));
            dt.Columns.Add("交易員", typeof(string));
            dt.Columns.Add("重設比", typeof(double));
            dt.Columns.Add("界限比", typeof(double));
            dt.Columns.Add("財務費用", typeof(double));
            dt.Columns.Add("獎勵", typeof(bool));//bool直接變checkedbox
            dt.Columns["獎勵"].ReadOnly = false;
            dt.Columns.Add("1500W", typeof(bool));
            dt.Columns["1500W"].ReadOnly = false;
            dt.Columns.Add("跌停價*", typeof(double));
            dt.Columns.Add("確認", typeof(bool));
            dt.Columns["確認"].ReadOnly = false;
            dt.Columns.Add("Adj", typeof(double));
            //dt.Columns.Add("發行原因", typeof(string));
            dt.Columns.Add("發行價格", typeof(double));
            dt.Columns.Add("標的名稱", typeof(string));
            dt.Columns.Add("分級", typeof(string));
            dt.Columns.Add("覆", typeof(bool));
            dt.Columns["覆"].ReadOnly = false;
            dt.Columns.Add("說明", typeof(string));
            dt.Columns.Add("今日額度", typeof(double));
            dt.Columns.Add("平均Theta", typeof(double));
            dt.Columns.Add("HV20Ratio", typeof(string));
            dt.Columns.Add("HV60Ratio", typeof(string));
            dt.Columns.Add("PL20日", typeof(double));
            dt.Columns.Add("PL年", typeof(double));
            dt.Columns.Add("股價", typeof(double));
            dt.Columns.Add("Delta", typeof(double));
            //joufan
            dt.Columns.Add("Theta", typeof(double));
            dt.Columns.Add("Vega", typeof(double));
            dt.Columns.Add("跳動價差", typeof(double));
            dt.Columns.Add("利率", typeof(double));


            dt.Columns.Add("獎勵額度", typeof(double));
            dt.Columns.Add("約當張數", typeof(double));
            dt.Columns.Add("約當張數(5000張)", typeof(double));
           

            dt.Columns.Add("昨日KgiCall DeltaOne", typeof(double));
            dt.Columns.Add("昨日KgiPut DeltaOne", typeof(double));
            dt.Columns.Add("今日KgiCall DeltaOne", typeof(double));
            dt.Columns.Add("今日KgiPut DeltaOne", typeof(double));
            dt.Columns.Add("自家 Call/Put DeltaOne比例", typeof(double));
            dt.Columns.Add("自家/市場Put DeltaOne比例>25%", typeof(string));
            dt.Columns.Add("Put大於元大", typeof(string));
            dt.Columns.Add("IV*", typeof(double));
            dt.Columns.Add("發行價格*", typeof(double));
            
            dt.Columns.Add("市場", typeof(string));
            

            ultraGrid1.DataSource = dt;
            UltraGridBand band0 = ultraGrid1.DisplayLayout.Bands[0];

            //ultraGrid1.DisplayLayout.Bands[0].Columns["確認"].Header.Appearance.
            band0.Columns["張數"].Format = "N0";
            band0.Columns["約當張數"].Format = "N0";
            band0.Columns["約當張數(5000張)"].Format = "N0";
            band0.Columns["今日額度"].Format = "N0";
            band0.Columns["獎勵額度"].Format = "N0";
            band0.Columns["平均Theta"].Format = "N0";
            band0.Columns["HV20Ratio"].Format = "N0";
            band0.Columns["HV60Ratio"].Format = "N0";
            band0.Columns["PL20日"].Format = "N0";
            band0.Columns["PL年"].Format = "N0";

            band0.Columns["類型"].Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownList;
            band0.Columns["CP"].Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownList;
            band0.Columns["交易員"].Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownList;
            //band0.Columns["發行原因"].Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownList;
            //ultraGrid1.DisplayLayout.Bands[0].Columns["確認"].Style = Infragistics.Win.UltraWinGrid.ColumnStyle.CheckBox;
            //ultraGrid1.DisplayLayout.Bands[0].Columns["刪除"].Style = Infragistics.Win.UltraWinGrid.ColumnStyle.CheckBox;

            //ultraGrid1.DisplayLayout.Bands[0].Columns["編號"].Width = 75;

            band0.Columns["標的代號"].Width = 60;
            band0.Columns["履約價"].Width = 60;
            band0.Columns["期間(月)"].Width = 55;
            band0.Columns["行使比例"].Width = 60;
            band0.Columns["建議行使比例"].Width = 90;
            band0.Columns["HV"].Width = 50;
            band0.Columns["IV"].Width = 50;
            //band0.Columns["建議Vol"].Width = 60;
            //band0.Columns["下限Vol"].Width = 60;
            band0.Columns["張數"].Width = 60;
            band0.Columns["重設比"].Width = 60;
            band0.Columns["界限比"].Width = 60;
            band0.Columns["財務費用"].Width = 60;

            band0.Columns["類型"].Width = 60;
            band0.Columns["CP"].Width = 30;
            band0.Columns["交易員"].Width = 70;
            band0.Columns["獎勵"].Width = 40;
            band0.Columns["確認"].Width = 40;
            band0.Columns["1500W"].Width = 50;
            band0.Columns["發行價格"].Width = 60;
            band0.Columns["Adj"].Width = 60;
            //band0.Columns["發行原因"].Width = 50;
            band0.Columns["標的名稱"].Width = 70;
            band0.Columns["分級"].Width = 40;
            band0.Columns["股價"].Width = 60;
            band0.Columns["Delta"].Width = 70;
            //joufan
            band0.Columns["Theta"].Width = 70;
            band0.Columns["Vega"].Width = 70;
            band0.Columns["跳動價差"].Width = 70;
            band0.Columns["利率"].Width = 50;
            band0.Columns["市場"].Width = 40;
            band0.Columns["IV*"].Width = 50;
            band0.Columns["發行價格*"].Width = 70;
            band0.Columns["跌停價*"].Width = 60;
            band0.Columns["約當張數"].Width = 60;
            band0.Columns["約當張數(5000張)"].Width = 130;
            band0.Columns["今日額度"].Width = 60;
            band0.Columns["獎勵額度"].Width = 60;
            band0.Columns["刪除"].Width = 40;
            band0.Columns["覆"].Width = 40;
            band0.Columns["平均Theta"].Width = 80;
            band0.Columns["HV20Ratio"].Width = 80;
            band0.Columns["HV60Ratio"].Width = 80;
            band0.Columns["PL20日"].Width = 60;
            band0.Columns["PL年"].Width = 60;

            band0.Columns["發行價格"].CellAppearance.BackColor = Color.LightGray;
            band0.Columns["標的名稱"].CellAppearance.BackColor = Color.LightGray;
            band0.Columns["股價"].CellAppearance.BackColor = Color.LightGray;
            band0.Columns["Delta"].CellAppearance.BackColor = Color.LightGray;
            //joufan
            band0.Columns["Theta"].CellAppearance.BackColor = Color.LightGray;
            band0.Columns["Vega"].CellAppearance.BackColor = Color.LightGray;
            band0.Columns["跳動價差"].CellAppearance.BackColor = Color.LightGray;
            band0.Columns["IV*"].CellAppearance.BackColor = Color.LightBlue;
            band0.Columns["發行價格*"].CellAppearance.BackColor = Color.LightBlue;
            band0.Columns["跌停價*"].CellAppearance.BackColor = Color.LightBlue;
            band0.Columns["市場"].CellAppearance.BackColor = Color.LightGray;
            band0.Columns["約當張數"].CellAppearance.BackColor = Color.LightGray;
            band0.Columns["約當張數(5000張)"].CellAppearance.BackColor = Color.LightGray;
            band0.Columns["今日額度"].CellAppearance.BackColor = Color.LightGray;
            band0.Columns["獎勵額度"].CellAppearance.BackColor = Color.LightGray;
            band0.Columns["刪除"].CellAppearance.BackColor = Color.Coral;


            //band0.Columns["標的代號"].SortIndicator = SortIndicator.None;

            // To sort multi-column using SortedColumns property
            // This enables multi-column sorting
            //如果開放sort功能，在更新templist時serialnum會亂掉
            //this.ultraGrid1.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti;

            // It is good practice to clear the sorted columns collection
            band0.SortedColumns.Clear();

            // You can sort (as well as group rows by) columns by using SortedColumns 
            // property off the band
            //band0.SortedColumns.Add("ContactName", false, false);

            //ultraGrid1.DisplayLayout.Bands[0].Columns["可發行股數"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right;
            //ultraGrid1.DisplayLayout.Bands[0].Columns["截至前一日"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right;
            //ultraGrid1.DisplayLayout.Bands[0].Columns["本日累積發行"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right;
            //ultraGrid1.DisplayLayout.Bands[0].Columns["累計%"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center;


            ultraGrid1.DisplayLayout.Bands[0].Override.HeaderAppearance.TextHAlign = Infragistics.Win.HAlign.Left;
            ultraGrid1.DisplayLayout.Bands[0].Columns["平均Theta"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right;
            ultraGrid1.DisplayLayout.Bands[0].Columns["HV20Ratio"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right;
            ultraGrid1.DisplayLayout.Bands[0].Columns["HV60Ratio"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right;
            ultraGrid1.DisplayLayout.Bands[0].Columns["PL20日"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right;
            ultraGrid1.DisplayLayout.Bands[0].Columns["PL年"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right;
            //ultraGrid1.DisplayLayout.Bands[0].Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
            //ultraGrid1.DisplayLayout.Bands[0].Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
            //ultraGrid1.DisplayLayout.Bands[0].Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.False;
            //ultraGrid1.DisplayLayout.Bands[0].Columns["確認"].CellActivation = Activation.AllowEdit;


            SetButton();
        }
        
        private void WarrantBasicRefreshTime()
        {

            string sql_warrantbasic = $@"SELECT  TOP(1) [MDate]
                            FROM [WarrantAssistant].[dbo].[InformationLog]
                            WHERE [InformationType] ='WarrantBasicRefresh' AND [MDate] >='{today.ToString("yyyyMMdd")}'
                            ORDER BY [MDate] DESC";
            DataTable dv_warrantbasic = MSSQL.ExecSqlQry(sql_warrantbasic, conn);
            string t1 = "";
            if (dv_warrantbasic.Rows.Count > 0)
            {
                bool isTime = DateTime.TryParse(dv_warrantbasic.Rows[0]["MDate"].ToString(), out DateTime temp);
                if (isTime)
                {
                    t1 = temp.Hour.ToString("0#") + ":" + temp.Minute.ToString("0#");
                }
            }
            if (t1.Length > 0)
                toolStripLabel3.Text = "WarrantBasic更新時間: " + t1;
            else
                toolStripLabel3.Text = "今日未更新WarrantBasic";
        }

        private void CreditRefreshTime()
        {

            string sql_credit = $@"SELECT  TOP(1) [UpdateTime]
                            FROM [WarrantAssistant].[dbo].[WarrantUnderlyingCreditNew]
                            WHERE [UpdateTime] >= '{today.ToString("yyyyMMdd")}'
                            ORDER BY [UpdateTime] DESC";
            DataTable dv_credit = MSSQL.ExecSqlQry(sql_credit, conn);

            string t2 = "";
            if (dv_credit.Rows.Count > 0)
            {
                bool isTime = DateTime.TryParse(dv_credit.Rows[0]["UpdateTime"].ToString(), out DateTime temp);
                if (isTime)
                {
                    t2 = temp.Hour.ToString("0#") + ":" + temp.Minute.ToString("0#");
                }
            }
            if (t2.Length > 0)
                toolStripLabel4.Text = "額度更新時間: " + t2;
            else
                toolStripLabel4.Text = "今日尚未更新額度";
        }

        private void LoadWMarket30()
        {
            WMarket30.Clear();

            string sql_WM = $@"SELECT [UnderlyingID], [WarrantName]
                                    FROM [WarrantAssistant].[dbo].[ApplyTotalList]
                                    WHERE [ApplyKind] ='1' AND [UserID] = '{userID}'";
            DataTable dv_WM = MSSQL.ExecSqlQry(sql_WM, GlobalVar.loginSet.warrantassistant45);

            foreach (DataRow dr in dv_WM.Rows)
            {
                string uid = dr["UnderlyingID"].ToString();
                string wname = dr["WarrantName"].ToString();
                if (!WMarket30.Contains(wname) && Market30.Contains(uid))
                    WMarket30.Add(wname);
            }
        }

        private void LoadUidDeltaOne()
        {

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
        }

        private void UpdateReIssueDeltaOne()
        {
            
            string sql_reissue = $@"SELECT  A.ReIssueNum AS IssueNum, B.UnderlyingID AS UnderlyingID, IsNull(B.exeRatio,0) CR, B.WarrantName AS WarrantName      
                              FROM [WarrantAssistant].[dbo].[ReIssueTempList] A
                              LEFT JOIN [WarrantAssistant].[dbo].[WarrantBasic] B ON A.WarrantID = B.WarrantID
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
        }
        private void LoadData() {
            try {
                
                dt.Rows.Clear();
                UidDeltaOne.Clear();
                UidDeltaOne_Temp.Clear();
                //確認權證基本資料是否有更新
                WarrantBasicRefreshTime();
                //確認額度是否有更新
                CreditRefreshTime();
                //高風險標的PutdeltaOne
                LoadUidDeltaOne();
                //11:40後搶發標的清單
                LoadIsCompete();
                //今日可發額度
                AvailableShares.Clear();
                //獎勵額度
                RewardShares.Clear();
                //台股標的
                IsTWUid.Clear();


                string sqlIsTw = $@"SELECT  [股票代號] FROM [TwCMData].[dbo].[ETF基本資料表] WHERE [年度] = (SELECT MAX(年度) FROM [TwCMData].[dbo].[ETF基本資料表]) AND [標的區域代號] <> 'TW'";
                DataTable dtIsTw = MSSQL.ExecSqlQry(sqlIsTw, GlobalVar.loginSet.twCMData);
                foreach(DataRow dr in dtIsTw.Rows)
                {
                    string uid = dr["股票代號"].ToString();
                    if (!IsTWUid.Contains(uid))
                        IsTWUid.Add(uid);
                }
                //抓發行標的的VolRatio
                string sqlVolRatio = $@"SELECT [UID]
                                        ,[WClass]
                                        ,SUM([AccReleasingLots] * [Theta_IV]) AS Theta
                                        ,SUM([AccReleasingLots] *  -1 *[Gamma_IV]) AS Gamma
                                        ,COUNT(*) AS 檔數
                                        ,AVG(HV_20D) * 100 AS HV20
                                        ,AVG(HV_60D) * 100 AS HV60
                                      FROM [TwData].[dbo].[V_WarrantTrading]
                                      WHERE [TDate] = '{lastday.ToString("yyyyMMdd")}' AND　[WClass] IN ('c','p') AND [TtoM] >= 90  AND LEN([UID]) < 5 AND LEFT([UID],2) <> '00'
                                      GROUP BY [UID],[WClass]";
                dtVolRatio = MSSQL.ExecSqlQry(sqlVolRatio, GlobalVar.loginSet.twData);

                //抓發行額度
                string sqlCredit = $@"SELECT  [UID],IsNull(A.[CanIssue],0) AS CanIssue,IsNull(Floor(A.[WarrantAvailableShares] * {GlobalVar.globalParameter.givenRewardPercent} - IsNull(B.[UsedRewardNum],0)), 0) AS RewardIssueCredit  
                                              FROM [WarrantAssistant].[dbo].[WarrantUnderlyingCreditNew] AS A
                                              LEFT JOIN [WarrantAssistant].[dbo].[WarrantReward] AS B ON A.UID = B.UnderlyingID
                                              WHERE [UpdateTime] > '{DateTime.Today.ToString("yyyyMMdd")}'";
                DataTable dtCredit = MSSQL.ExecSqlQry(sqlCredit, GlobalVar.loginSet.warrantassistant45);
                foreach (DataRow dr in dtCredit.Rows)
                {
                    string uid = dr["UID"].ToString();
                    double canIssue = Math.Floor(Convert.ToDouble(dr["CanIssue"].ToString()));
                    double canIssueReward = Math.Floor(Convert.ToDouble(dr["RewardIssueCredit"].ToString()));
                    if (!AvailableShares.ContainsKey(uid))
                        AvailableShares.Add(uid, canIssue);
                    if (!RewardShares.ContainsKey(uid))
                        RewardShares.Add(uid, canIssueReward);
                }
                
                //抓取自家近20個交易日掛牌的權證，發行時檢查是否有重複K
                string sqlKK = $@"SELECT [wid],[stkid],[strike_now],CASE WHEN [type] LIKE '%購%' THEN 'c' ELSE 'p' END AS CP
                                FROM [HEDGE].[dbo].[WARRANTS] 
                                WHERE ISSUECOMNAME = '凱基' AND ([type] IN ('一般型認購權證','一般型認售權證') OR ([type] IN ('浮動重設認購權證','浮動重設認售權證') AND [marketdate] <= '{DateTime.Today.ToString("yyyyMMdd")}')) AND [issuedate] >= '{TradeDate.LastNTradeDate(20)}'";
                dtKK = MSSQL.ExecSqlQry(sqlKK, GlobalVar.loginSet.HEDGE);

                //抓取自家TtoM差20天內的權證，發行時檢查是否有重複K
                string sqlTT = $@"SELECT A.UID,B.最新履約價 AS [StrikePrice],A.WClass,A.TtoM FROM
(SELECT  [UID],[WID],[WClass],[StrikePrice],[TtoM]
                                  FROM [TwData].[dbo].[V_WarrantTrading] WHERE [TDate] = (SELECT MAX(TDate) FROM [TwData].[dbo].[WarrantFlow]) AND [WClass] IN ('c','p') AND [IssuerName] = '9200'
                                  AND [TtoM] > {T2TtoM[6] - 20}) AS A
LEFT JOIN (SELECT  [代號],[最新履約價]　FROM [TwCMData].[dbo].[Warrant總表]　WHERE [日期] = (SELECT MAX(TDate) FROM [TwData].[dbo].[WarrantFlow])) AS B ON A.WID = B.代號";
                dtTT = MSSQL.ExecSqlQry(sqlTT, GlobalVar.loginSet.twData);

                //抓取該標的近20個交易日損益
                string sqlPL20D = $@"SELECT  [UID],[OptionType],ROUND(SUM([MTH] + [SimUPL_IV])/1000,0) AS PL
                                      FROM [TwPL].[dbo].[PLSplit_All]
                                      WHERE [TDate] >= '{EDLib.TradeDate.LastNTradeDate(20).ToString("yyyyMMdd")}' AND [IssuerName] = '9200'　AND [ProductType] = 'Warrant'　AND LEN([UID]) < 5 AND LEFT([UID],2) <> '00' AND [OptionType] IN ('c','p')
                                      GROUP BY [UID],[OptionType]";
                dtPL20D = MSSQL.ExecSqlQry(sqlPL20D,  "Data Source=10.60.0.39;Initial Catalog=TwPL;User ID=WarrantWeb;Password=WarrantWeb");
                //抓取該標的年損益
                string sqlPLYear = $@"SELECT  [UnderlyingId],ROUND([YMTMNetTradePL] / 1000,0) AS PL
                                      FROM [PositionReport].[dbo].[UnderlyingPL]
                                      WHERE [TradeDate] = '{EDLib.TradeDate.LastNTradeDate(1).ToString("yyyyMMdd")}'　AND [strategymainnm] = '權證造市' AND [strategymainid] = 'WMM' AND [strategysubid] = 'WMM'　AND [strategysubnm] = '權證業務'　AND LEN([UnderlyingId]) < 5 AND LEFT([UnderlyingId],2) <> '00'";
                dtPLYear = MSSQL.ExecSqlQry(sqlPLYear, "Data Source=10.60.0.37;Initial Catalog=PositionReport;User ID=WarrantWeb;Password=WarrantWeb");


                //20231207 改穿VOL邏輯:HV20 > HV20(T-1)，HV20 > HV60 近5日漲跌幅(含今日) > 20，符合條件會跳出提示，要發長天期權證或獨立檔數
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
                                WHERE B.HV20Dif >0 AND B.HV20 >= B.HV60";

                dt_LongTerm = MSSQL.ExecSqlQry(sql_LongTerm, GlobalVar.loginSet.warrantassistant45);



                //抓取交易所7-1表格時間
                string sql_apply71 = $@"SELECT [ApplyTime], [OriApplyTime],[SerialNum] FROM [WarrantAssistant].[dbo].[Apply_71]";
                DataTable dv_apply71 = MSSQL.ExecSqlQry(sql_apply71, GlobalVar.loginSet.warrantassistant45);
                //抓取發行資料
                string sql = $@"SELECT A.SerialNum
                                  ,A.UnderlyingID
                                  ,A.K
                                  ,A.T
                                  ,A.R
                                  ,A.HV
                                  ,A.IV
                                  ,A.IssueNum
                                  ,A.ResetR
                                  ,A.BarrierR
                                  ,A.FinancialR
                                  ,A.Type
                                  ,A.CP
                                  ,A.TraderID
                                  ,CASE WHEN A.UseReward='Y' THEN 1 ELSE 0 END UseReward
                                  ,CASE WHEN A.ConfirmChecked='Y' THEN 1 ELSE 0 END ConfirmChecked
                                  ,CASE WHEN A.Apply1500W='Y' THEN 1 ELSE 0 END Apply1500W
	                              ,B.UnderlyingName
	                              ,C.MPrice
	                              ,B.Market
	                              ,(A.IssueNum*A.R) AS EquivalentNum
                                  ,5000 * A.R AS EquivalentNum5000
                                  ,IsNull(E.[CanIssue],0) AS IssueCredit
	                              ,IsNull(Floor(E.[WarrantAvailableShares] * {GlobalVar.globalParameter.givenRewardPercent} - IsNull(F.[UsedRewardNum],0)), 0) AS RewardIssueCredit                                 
                                  ,IsNull(A.[Adj],0) Adj
                                  ,IsNull(A.[說明],'') 說明
                                  ,CASE WHEN A.[Delete]='Y' THEN 1 ELSE 0 END AS 刪除
                                  ,G.[標的分級]
                              FROM [WarrantAssistant].[dbo].[ApplyTempList] A 
                    LEFT JOIN [WarrantAssistant].[dbo].[WarrantUnderlyingSummary] B ON A.UnderlyingID=B.UnderlyingID
                    LEFT JOIN [WarrantAssistant].[dbo].[WarrantPrices] C ON A.UnderlyingID=C.CommodityID 
                    LEFT JOIN [Underlying_TraderIssue] D on D.UID = A.UnderlyingID 
                    LEFT JOIN (SELECT [UID], [CanIssue], [WarrantAvailableShares] FROM [WarrantAssistant].[dbo].[WarrantUnderlyingCreditNew] WHERE [UpdateTime] > '{DateTime.Today.ToString("yyyyMMdd")}' ) as E on A.UnderlyingID = E.[UID]
                    LEFT JOIN [WarrantAssistant].[dbo].[WarrantReward] F on A.UnderlyingID=F.UnderlyingID
                    LEFT JOIN (SELECT [UID],[WClass],[標的分級] FROM [WarrantIssue].[dbo].[UIDClassification] WHERE [TDate] = (SELECT MAX(TDate) FROM [TwData].[dbo].[WarrantFlow])) AS G ON A.[UnderlyingID] = G.[UID] AND A.[CP] = G.[WClass]
                    WHERE A.UserID='{userID}' 
                    ORDER BY A.MDate, SUBSTRING(A.SerialNum,9,7),CONVERT(INT,SUBSTRING(A.SerialNum,18,LEN(A.SerialNum)-17))";

                DataTable dv = MSSQL.ExecSqlQry(sql, GlobalVar.loginSet.warrantassistant45);
                if (dv.Rows.Count > 0) {
                    try
                    {
                        foreach (DataRow drv in dv.Rows)
                        {
                            DataRow dr = dt.NewRow();
                            //判斷是否今天申請的權證是否要加deltaone，若是沒搶到的權證就不用加
                            bool deltaone = true;
                            //只要在發行頁面key新權證，就會給一個編號，主要是程式內部用的
                            string serialNum = drv["SerialNum"].ToString();
                            string underlyingID = drv["UnderlyingID"].ToString();
                            dr["編號"] = serialNum;
                            dr["標的代號"] = underlyingID;
                            double k = Convert.ToDouble(drv["K"]);
                            dr["履約價"] = k;
                            int t = Convert.ToInt32(drv["T"]);
                            dr["期間(月)"] = t;
                            double cr = Convert.ToDouble(drv["R"]);
                            dr["行使比例"] = cr;
                            dr["HV"] = Convert.ToDouble(drv["HV"]);
                            
                            double vol = Convert.ToDouble(drv["IV"]) / 100;
                            dr["IV"] = Convert.ToDouble(drv["IV"]);
                           
                            double shares = Convert.ToDouble(drv["IssueNum"]);
                            dr["張數"] = shares;
                            
                            double resetR = Convert.ToDouble(drv["ResetR"]) / 100;
                            dr["重設比"] = Convert.ToDouble(drv["ResetR"]);
                            
                            double barrierR = Convert.ToDouble(drv["BarrierR"]);
                            dr["界限比"] = barrierR;
                            
                            double financialR = Convert.ToDouble(drv["FinancialR"]) / 100;
                            dr["財務費用"] = Convert.ToDouble(drv["FinancialR"]);
                            
                            string warrantType = drv["Type"].ToString();
                            dr["類型"] = warrantType;
                            
                            CallPutType cp = drv["CP"].ToString() == "C" ? CallPutType.Call : CallPutType.Put;
                            dr["CP"] = drv["CP"].ToString();
                            dr["交易員"] = drv["TraderID"].ToString();
                            
                            dr["獎勵"] = drv["UseReward"];
                            
                            dr["確認"] = drv["ConfirmChecked"];
                            
                            //dr["發行原因"] = drv["Reason"] == DBNull.Value ? 0 : Convert.ToInt32(drv["Reason"]);
                            dr["1500W"] = drv["Apply1500W"];

                            dr["標的名稱"] = drv["UnderlyingName"].ToString();
                            double underlyingPrice = Convert.ToDouble(drv["MPrice"]);
                            dr["股價"] = underlyingPrice;
                            
                            dr["市場"] = drv["Market"].ToString();
                            dr["刪除"] = drv["刪除"];
                            dr["約當張數"] = Convert.ToDouble(drv["EquivalentNum"]);
                            dr["約當張數(5000張)"] = Convert.ToDouble(drv["EquivalentNum5000"]);
                            double credit = Math.Floor((double)drv["IssueCredit"]);
                            
                            double rewardCredit = Math.Floor((double)drv["RewardIssueCredit"]);
                            
                            dr["今日額度"] = credit;
                            
                            dr["獎勵額度"] = rewardCredit;
                           
                            double adj = (double)drv["Adj"];
                            
                            dr["Adj"] = adj;
                            string ex = Convert.ToString(drv["說明"]);
                            dr["說明"] = ex;
                            if (ex.Contains("對"))
                                dr["覆"] = true;
                            else
                                dr["覆"] = false;

                            string uRank = drv["標的分級"].ToString();
                            dr["分級"] = uRank;

                            double price = 0.0;
                            double delta = 0.0;
                            double theta = 0.0; //joufan
                            double vega = 0.0;
                            if (underlyingPrice != 0)
                            {
                                if (underlyingID.Length > 4 && underlyingID.Substring(0, 2) != "00")
                                {
                                    if (warrantType == "牛熊證")
                                        price = Pricing.BullBearWarrantPrice(cp, underlyingPrice + adj, resetR, GlobalVar.globalParameter.interestRate_Index, vol, t, financialR, cr);
                                    //else if (warrantType == "重設型")
                                    else if (warrantType == "重設型" || warrantType == "展延型")
                                        price = Pricing.ResetWarrantPrice(cp, underlyingPrice + adj, resetR, GlobalVar.globalParameter.interestRate_Index, vol, t, cr);
                                    else
                                        price = Pricing.NormalWarrantPrice(cp, underlyingPrice + adj, k, GlobalVar.globalParameter.interestRate_Index, vol, t, cr);
                                }
                                else
                                {
                                    if (warrantType == "牛熊證")
                                        price = Pricing.BullBearWarrantPrice(cp, underlyingPrice + adj, resetR, GlobalVar.globalParameter.interestRate, vol, t, financialR, cr);
                                    //else if (warrantType == "重設型")
                                    else if (warrantType == "重設型" || warrantType == "展延型")
                                        price = Pricing.ResetWarrantPrice(cp, underlyingPrice + adj, resetR, GlobalVar.globalParameter.interestRate, vol, t, cr);
                                    else
                                        price = Pricing.NormalWarrantPrice(cp, underlyingPrice + adj, k, GlobalVar.globalParameter.interestRate, vol, t, cr);
                                }
                                if (warrantType == "牛熊證")
                                {
                                    delta = 1.0;
                                    theta = -k * financialR * cr / 365.0;
                                }
                                else
                                {
                                    if (underlyingID.Length > 4 && underlyingID.Substring(0, 2) != "00")
                                    {
                                        delta = Pricing.Delta(cp, underlyingPrice + adj, k, GlobalVar.globalParameter.interestRate_Index, vol, (t * 30.0) / GlobalVar.globalParameter.dayPerYear, GlobalVar.globalParameter.interestRate_Index) * cr;
                                        theta = Pricing.Theta(cp, underlyingPrice + adj, k, GlobalVar.globalParameter.interestRate_Index, vol, (t * 30.0) / GlobalVar.globalParameter.dayPerYear, GlobalVar.globalParameter.interestRate_Index) * cr;
                                        if (cp == CallPutType.Call)
                                            vega = EDLib.Pricing.Option.PlainVanilla.CallVega(underlyingPrice + adj, k, GlobalVar.globalParameter.interestRate_Index, vol, (t * 30.0) / GlobalVar.globalParameter.dayPerYear) * cr / 100;
                                        if (cp == CallPutType.Put)
                                            vega = EDLib.Pricing.Option.PlainVanilla.PutVega(underlyingPrice + adj, k, GlobalVar.globalParameter.interestRate_Index, vol, (t * 30.0) / GlobalVar.globalParameter.dayPerYear) * cr / 100;
                                    }
                                    else
                                    {
                                        delta = Pricing.Delta(cp, underlyingPrice + adj, k, GlobalVar.globalParameter.interestRate, vol, (t * 30.0) / GlobalVar.globalParameter.dayPerYear, GlobalVar.globalParameter.interestRate) * cr;
                                        theta = Pricing.Theta(cp, underlyingPrice + adj, k, GlobalVar.globalParameter.interestRate, vol, (t * 30.0) / GlobalVar.globalParameter.dayPerYear, GlobalVar.globalParameter.interestRate) * cr;
                                        if (cp == CallPutType.Call)
                                            vega = EDLib.Pricing.Option.PlainVanilla.CallVega(underlyingPrice + adj, k, GlobalVar.globalParameter.interestRate, vol, (t * 30.0) / GlobalVar.globalParameter.dayPerYear) * cr / 100;
                                        if (cp == CallPutType.Put)
                                            vega = EDLib.Pricing.Option.PlainVanilla.PutVega(underlyingPrice + adj, k, GlobalVar.globalParameter.interestRate, vol, (t * 30.0) / GlobalVar.globalParameter.dayPerYear) * cr / 100;
                                    }
                                }
                                
                            }
                            
                            dr["發行價格"] = Math.Round(price, 2);
                            double jumpSize = 0.0;
                            double multiplier = EDLib.Tick.UpTickSize(underlyingID, underlyingPrice + adj);

                            jumpSize = delta * multiplier;
                            

                            //計算發行湊1500萬所需的VOL
                            double vol_ = vol;
                            double price_ = price;
                            double lowerLimit = 0.0;
                            double totalValue = price_ * shares * 1000;
                            double volLimit = 2 * vol_;


                            if (underlyingID.Length > 4 && underlyingID.Substring(0, 2) != "00")
                            {
                                dr["利率"] = GlobalVar.globalParameter.interestRate_Index * 100;
                            }
                            else
                                dr["利率"] = GlobalVar.globalParameter.interestRate * 100;

                            
                            while (totalValue < 15000000 && vol_ < volLimit)
                            {
                                vol_ += 0.01;
                                if (underlyingID.Length > 4 && underlyingID.Substring(0, 2) != "00")
                                {
                                    if (warrantType == "牛熊證")
                                        price_ = Pricing.BullBearWarrantPrice(cp, underlyingPrice, resetR, GlobalVar.globalParameter.interestRate_Index, vol_, t, financialR, cr);
                                    //else if (warrantType == "重設型")
                                    else if (warrantType == "重設型" || warrantType == "展延型")
                                        price_ = Pricing.ResetWarrantPrice(cp, underlyingPrice, resetR, GlobalVar.globalParameter.interestRate_Index, vol_, t, cr);
                                    else
                                        price_ = Pricing.NormalWarrantPrice(cp, underlyingPrice, k, GlobalVar.globalParameter.interestRate_Index, vol_, t, cr);
                                }
                                else
                                {
                                    if (warrantType == "牛熊證")
                                        price_ = Pricing.BullBearWarrantPrice(cp, underlyingPrice, resetR, GlobalVar.globalParameter.interestRate, vol_, t, financialR, cr);
                                    //else if (warrantType == "重設型")
                                    else if (warrantType == "重設型" || warrantType == "展延型")
                                        price_ = Pricing.ResetWarrantPrice(cp, underlyingPrice, resetR, GlobalVar.globalParameter.interestRate, vol_, t, cr);
                                    else
                                        price_ = Pricing.NormalWarrantPrice(cp, underlyingPrice, k, GlobalVar.globalParameter.interestRate, vol_, t, cr);
                                }
                                totalValue = price_ * shares * 1000;
                            }

                            //計算完沒有ADJ的VOL後，發行要顯示含ADJ的價格
                            if (underlyingID.Length > 4 && underlyingID.Substring(0, 2) != "00")
                            {
                                if (warrantType == "牛熊證")
                                    price_ = Pricing.BullBearWarrantPrice(cp, underlyingPrice + adj, resetR, GlobalVar.globalParameter.interestRate_Index, vol_, t, financialR, cr);
                                //else if (warrantType == "重設型")
                                else if (warrantType == "重設型" || warrantType == "展延型")
                                    price_ = Pricing.ResetWarrantPrice(cp, underlyingPrice + adj, resetR, GlobalVar.globalParameter.interestRate_Index, vol_, t, cr);
                                else
                                    price_ = Pricing.NormalWarrantPrice(cp, underlyingPrice + adj, k, GlobalVar.globalParameter.interestRate_Index, vol_, t, cr);
                            }
                            else
                            {
                                if (warrantType == "牛熊證")
                                    price_ = Pricing.BullBearWarrantPrice(cp, underlyingPrice + adj, resetR, GlobalVar.globalParameter.interestRate, vol_, t, financialR, cr);
                                //else if (warrantType == "重設型")
                                else if (warrantType == "重設型" || warrantType == "展延型")
                                    price_ = Pricing.ResetWarrantPrice(cp, underlyingPrice + adj, resetR, GlobalVar.globalParameter.interestRate, vol_, t, cr);
                                else
                                    price_ = Pricing.NormalWarrantPrice(cp, underlyingPrice + adj, k, GlobalVar.globalParameter.interestRate, vol_, t, cr);
                            }

                            if (!IsTWUid.Contains(underlyingID))
                                lowerLimit = Math.Max(0.01, price_ - (underlyingPrice + adj) * 0.1 * cr);
                            else
                                lowerLimit = 0.01;

                            dr["IV*"] = vol_ * 100;
                            dr["發行價格*"] = Math.Round(price_, 2);
                            dr["跌停價*"] = Math.Round(lowerLimit, 2);
                            dr["Delta"] = Math.Round(delta, 4);
                            dr["Theta"] = Math.Round(theta, 4); //joufan
                            dr["Vega"] = Math.Round(vega, 4); //joufan
                            dr["跳動價差"] = Math.Round(jumpSize, 4);
                            dt.Rows.Add(dr);
                            
                            if ((int)drv["ConfirmChecked"] == 1 && UidDeltaOne_Temp.ContainsKey(underlyingID))//有點確認的要加上DeltaOne
                            {
                                string sqlTemp = "SELECT [ApplyTime], [OriApplyTime] FROM [WarrantAssistant].[dbo].[Apply_71] WHERE SerialNum = '" + serialNum + "'";
                                DataView dvTemp = DeriLib.Util.ExecSqlQry(sqlTemp, GlobalVar.loginSet.warrantassistant45);
                                DataRow[] select = dv_apply71.Select($@"SerialNum ='{serialNum}'");
                                string applyTime = "";
                                string apytime = "";
                                string oriapplyTime = "";
                               
                                if (select.Length > 0)
                                {
                                    applyTime = select[0][0].ToString().Substring(0, 2);
                                    apytime = select[0][0].ToString();
                                    oriapplyTime = select[0][1].ToString();
                                }
                                
                                //if (applyTime == "22" || (apytime.Length == 0) && Iscompete.Contains(underlyingID) && (int)drv["UseReward"] == 0)
                                //  deltaone = false;
                                //如果是搶發權證中其他沒搶到的，就不用加deltaone
                                if (deltaone)
                                {
                                    
                                    if (drv["CP"].ToString() == "C")//如果是Call  要考慮Put=0的情況 
                                    {
                                        UidDeltaOne_Temp[underlyingID].KgiCallDeltaOne += shares * cr;
                                        if (UidDeltaOne_Temp[underlyingID].KgiPutNum == 0)
                                            UidDeltaOne_Temp[underlyingID].KgiCallPutRatio = 100;
                                        else
                                            UidDeltaOne_Temp[underlyingID].KgiCallPutRatio = Math.Round((double)UidDeltaOne_Temp[underlyingID].KgiCallDeltaOne / (double)UidDeltaOne_Temp[underlyingID].KgiPutDeltaOne, 4);
                                        UidDeltaOne_Temp[underlyingID].AllCallDeltaOne += shares * cr;
                                        if (UidDeltaOne_Temp[underlyingID].AllPutDeltaOne == 0)
                                            UidDeltaOne_Temp[underlyingID].KgiAllPutRatio = 100;
                                        else
                                            UidDeltaOne_Temp[underlyingID].KgiAllPutRatio = Math.Round((double)UidDeltaOne_Temp[underlyingID].KgiPutDeltaOne / (double)UidDeltaOne_Temp[underlyingID].AllPutDeltaOne, 4);
                                    }
                                    else
                                    {
                                        
                                        UidDeltaOne_Temp[underlyingID].KgiPutDeltaOne += shares * cr;
                                        UidDeltaOne_Temp[underlyingID].KgiCallPutRatio = Math.Round((double)UidDeltaOne_Temp[underlyingID].KgiCallDeltaOne / (double)UidDeltaOne_Temp[underlyingID].KgiPutDeltaOne, 4);
                                        UidDeltaOne_Temp[underlyingID].AllPutDeltaOne += shares * cr;
                                        UidDeltaOne_Temp[underlyingID].KgiAllPutRatio = Math.Round((double)UidDeltaOne_Temp[underlyingID].KgiPutDeltaOne / (double)UidDeltaOne_Temp[underlyingID].AllPutDeltaOne, 4);
                                        UidDeltaOne_Temp[underlyingID].KgiPutNum++;
                                    }
                                    
                                }
                            }
                        }
                        
                    }
                    catch(Exception ex) 
                    {
                        MessageBox.Show(ex.Message);
                    }
                }
                //高風險標的Putdeltaone檢核會包含今天預計發行的額度，所以新發與增額都要算進去
                UpdateReIssueDeltaOne();
                //標的為市值前30大的權證
                LoadWMarket30();

            } catch (Exception ex) {
                
                MessageBox.Show(ex.Message);
            }
        }
       
        private bool CheckData() {
            bool dataOK = true;
            bool differentTrader = false;
            string differentTraderUid = "";


            //20230525 更新 重設型權證上市日不可與除權息日同一天
            string sqlexList = $@" SELECT  [股票代號]
                                  FROM [TwCMData].[dbo].[停止過戶及停止資券表]
                                  WHERE [原因] IN ('除權','除息') AND [日期] IN ('{EDLib.TradeDate.NextNTradeDate(3).ToString("yyyyMMdd")}','{EDLib.TradeDate.NextNTradeDate(2).ToString("yyyyMMdd")}') AND　[年月] IN ('{DateTime.Today.Year.ToString("0#")}{DateTime.Today.Month.ToString("0#")}','{DateTime.Today.Year.ToString("0#")}{DateTime.Today.AddMonths(1).Month.ToString("0#")}','{DateTime.Today.Year.ToString("0#")}{DateTime.Today.AddMonths(2).Month.ToString("0#")}')";

            DataTable dtexList = MSSQL.ExecSqlQry(sqlexList, GlobalVar.loginSet.twCMData);
            
            //有申請放寬CallPut Ratio的標的
            Dictionary<string, double> dicAuthorize = new Dictionary<string, double>();
            string sqlauthorize = $@"SELECT [UID],[截止日],[CPRatio]　FROM [WarrantAssistant].[dbo].[CPRatioMonitor]　WHERE [截止日] >= GETDATE()";
            DataTable dtauthorize = MSSQL.ExecSqlQry(sqlauthorize, conn);
            foreach(DataRow draut in dtauthorize.Rows)
            {
                string uid = draut["UID"].ToString();
                double r = Convert.ToDouble(draut["CPRatio"].ToString());
                if (!dicAuthorize.ContainsKey(uid))
                    dicAuthorize.Add(uid, r);
            }

            string sql2 = $@"SELECT A.[UnderlyingID]
             FROM [WarrantAssistant].[dbo].[ApplyTempList] AS A
             LEFT JOIN (SELECT * FROM [WarrantAssistant].[dbo].[WarrantIssueCheck] WHERE [IsQuaterUnderlying] = 'Y') AS B ON A.UnderlyingID = B.UnderlyingID
             WHERE A.[UserID]='{userID}' AND A.[ConfirmChecked]='Y' and (A.HV = 0 or A.IV = 0 or A.IssueNum = 0 or A.T = 0 or A.K = 0 or B.[WarningScore] > 0)";

            DataTable badParam = MSSQL.ExecSqlQry(sql2, conn);// new DataTable("noReason");            

            foreach (DataRow Row in badParam.Rows)
            {
                MessageBox.Show(Row["UnderlyingID"] + " 發行條件輸入有誤OR為警示股，會被後臺某些人罵，避免他們該該叫，請修改條件。");
                dataOK = false;
            }
            sql2 = "SELECT [UnderlyingID] FROM [WarrantAssistant].[dbo].[ApplyTempList] as A "
                + " left join (Select CS8010, count(1) as count from [VOLDB].[dbo].[ED_RelationUnderlying] "
                          + $" where RecordDate = (select top 1 RecordDate from [VOLDB].[dbo].[ED_RelationUnderlying])"
                           + " group by CS8010) as B on A.UnderlyingID = B.CS8010 "
                 + " left join (SELECT stkid, MAX([IssueVol]) as MAX, min(IssueVol) as min FROM[10.101.10.5].[WMM3].[dbo].[Warrants]"
                            + " where kgiwrt = '他家' and marketdate <= GETDATE() and lasttradedate >= GETDATE() and IssueVol<> 0 "
                            + " group by stkid ) as C on A.UnderlyingID = C.stkid "
                + $" WHERE [UserID] = '{userID}' AND [ConfirmChecked] = 'Y' and B.count > 0 and (IV > C.MAX or IV < C.min)";
            badParam = MSSQL.ExecSqlQry(sql2, "Data Source=10.60.0.39;Initial Catalog=VOLDB;User ID=voldbuser;Password=voldbuser");

            foreach (DataRow Row in badParam.Rows)
            {
                MessageBox.Show(Row["UnderlyingID"] + " 為關係人標的，波動度超過可發範圍，會被稽核稽稽歪歪，請修改條件。");
                dataOK = false;
            }
            //避免發行在長假後到期權證 EX 除夕
            List<string> Holidays = new List<string>();
            string sqlGetLongHoliday = $@"SELECT CONVERT(VARCHAR,B.TradeDate,112) AS D FROM (
                                    SELECT  ROW_NUMBER() OVER (ORDER BY [TradeDate]) AS NUM,[TradeDate] FROM [DeriPosition].[dbo].[Calendar]
                                      WHERE [TradeDate] >= 　CONVERT(VARCHAR,GETDATE(),112)  AND [CountryId] = 'TWN'　AND [IsTrade] = 'Y') AS A
                                      LEFT JOIN
                                    (SELECT  ROW_NUMBER() OVER (ORDER BY [TradeDate]) -1 AS NUM,[TradeDate] FROM [DeriPosition].[dbo].[Calendar]
                                      WHERE [TradeDate] >= 　CONVERT(VARCHAR,GETDATE(),112)  AND [CountryId] = 'TWN'　AND [IsTrade] = 'Y') AS B ON　A.NUM = B.NUM
                                      WHERE B.TradeDate IS NOT NULL AND DATEDIFF(DAY,A.TradeDate,B.TradeDate) > 5";
            DataTable dtGetLongHoliday = MSSQL.ExecSqlQry(sqlGetLongHoliday, GlobalVar.loginSet.tsquoteSqlConnString);
            foreach(DataRow dr in dtGetLongHoliday.Rows)
            {
                string h = dr["D"].ToString();

                if (!Holidays.Contains(h))
                    Holidays.Add(h);
                DateTime dh = DateTime.ParseExact(h, "yyyyMMdd", null);
                int i = 0;
                while(i < 10)
                {
                    dh = dh.AddDays(1);
                    if (EDLib.TradeDate.IsTradeDay(dh))
                    {
                        string h2 = dh.ToString("yyyyMMdd");
                        if (!Holidays.Contains(h2))
                            Holidays.Add(h2);
                        i++;
                    }

                }
                
            }


            string sqlListedUID = $@"SELECT  [建立日期],[標的代號],[說明]　,ISNULL([解除日期],'19110101') AS 解除日期　FROM [WarrantAssistant].[dbo].[正面表列標的]";
            DataTable dtListedUID = MSSQL.ExecSqlQry(sqlListedUID, GlobalVar.loginSet.warrantassistant45);


            //用來計算今天發行put使用的PutDeltaOne，有些標的會控管Put市佔
            Dictionary<string, double> PutDeltaOne = new Dictionary<string, double>();
            foreach (Infragistics.Win.UltraWinGrid.UltraGridRow dr in ultraGrid1.Rows)
            {
                //string serial = dr.Cells["編號"].Value.ToString();
                
                string cp = dr.Cells["CP"].Value.ToString();
                string underlyingID = dr.Cells["標的代號"].Value.ToString();
                
                DataRow[] dtListedUID_Select = dtListedUID.Select($@"標的代號 = '{underlyingID}'");
                if(dtListedUID_Select.Length > 0)
                {
                    string ex = dtListedUID_Select[0][2].ToString();
                    string removeDate = dtListedUID_Select[0][2].ToString();
                    if(removeDate == "19110101")
                    {
                        dataOK = false;
                        MessageBox.Show($@"{underlyingID} 為{ex}，未解除");
                    }
                    if(ex != "關係人標的")
                    {
                        MessageBox.Show($@"{underlyingID} {ex}");
                    }
                }


                double spot = dr.Cells["股價"].Value == DBNull.Value ? 0.0 : Convert.ToDouble(dr.Cells["股價"].Value);
                double cr = dr.Cells["行使比例"].Value == DBNull.Value ? 0 : Convert.ToDouble(dr.Cells["行使比例"].Value);
                double shares = dr.Cells["張數"].Value == DBNull.Value ? 10000 : Convert.ToDouble(dr.Cells["張數"].Value);
                bool confirmed = dr.Cells["確認"].Value == DBNull.Value ? false : Convert.ToBoolean(dr.Cells["確認"].Value);
                int t = dr.Cells["期間(月)"].Value == DBNull.Value ? 6 : Convert.ToInt32(dr.Cells["期間(月)"].Value);
                string type = dr.Cells["類型"].Value.ToString();
                double resetR = dr.Cells["重設比"].Value == DBNull.Value ? 0 : Convert.ToDouble(dr.Cells["重設比"].Value);

                if (cp == "P")
                {
                    if (!PutDeltaOne.ContainsKey(underlyingID))
                        PutDeltaOne.Add(underlyingID, 0);
                    PutDeltaOne[underlyingID] += shares * cr;
                }
               
                if(underlyingID == "3089" && confirmed)
                {
                    MessageBox.Show($@"高風險分數大於10分，本月暫緩發行");
                    dataOK = false;
                }
                if (underlyingID == "1312" && confirmed)
                {
                    MessageBox.Show($@"國喬為高風險標的，不可發行");
                    dataOK = false;
                }

                DateTime expiryDate = GlobalVar.globalParameter.nextTradeDate3.AddMonths(t);
                if (expiryDate.Day == GlobalVar.globalParameter.nextTradeDate3.Day)
                    expiryDate = expiryDate.AddDays(-1);
                string sqlTemp = $"SELECT TOP 1 TradeDate from [DeriPosition].[dbo].[Calendar] WHERE IsTrade='Y' AND [CountryId] = 'TWN' AND TradeDate >= '{expiryDate.ToString("yyyy-MM-dd")}'";
                
                //DataView dvTemp = DeriLib.Util.ExecSqlQry(sqlTemp, GlobalVar.loginSet.tsquoteSqlConnString);
                DataTable dvTemp = MSSQL.ExecSqlQry(sqlTemp, GlobalVar.loginSet.tsquoteSqlConnString);
                foreach (DataRow drTemp in dvTemp.Rows)
                {
                    expiryDate = Convert.ToDateTime(drTemp["TradeDate"]);
                }
                int month = expiryDate.Month;


                if (Holidays.Contains(expiryDate.ToString("yyyyMMdd")) && confirmed)
                {

                   
                    MessageBox.Show($@"不可發行在長假後到期之權證，日期{expiryDate.ToString("yyyyMMdd")}，請將T多加1個月，目前發行期間為{t}月");
                    dataOK = false;
                  

                }
                if (month >= 6 && month <= 9)
                {
                    
                    if(type == "牛熊證"  && underlyingID.Contains("IX") && confirmed)
                    {
                        MessageBox.Show($@"t = {t}的牛熊證到期日不可介於6~9月!");
                        dataOK = false;
                    }
                    if (type != "牛熊證" && underlyingID.Contains("IX"))
                    {
                        MessageBox.Show($@"t = {t}的指數權證到期日介於6~9月，仍要發行?");
                        //dataOK = false;
                    }

                }
                if(type == "重設型" && confirmed)
                {
                    DataRow[] dtexListSelect = dtexList.Select($@"[股票代號] = '{underlyingID}'");
                    if(dtexListSelect.Length > 0)
                    {
                        MessageBox.Show($@"{underlyingID} 重設型上市日不可與除權息日同一天!");
                        dataOK = false;
                            
                    }
                    if(cp == "C")
                    {
                        if(resetR < 85)
                        {
                            MessageBox.Show($@"Call重設比 不得低於85");
                            dataOK = false;
                        }
                    }
                    if (cp == "P")
                    {
                        if (resetR > 115)
                        {
                            MessageBox.Show($@"Put重設比 不得高於115");
                            dataOK = false;
                        }
                    }
                }
                if (type == "牛熊證" && confirmed)
                {
                    if (cp == "C")
                    {
                        if (resetR >= 100)
                        {
                            MessageBox.Show($@"牛證重設比 不得高於100");
                            dataOK = false;
                        }
                    }
                    if (cp == "P")
                    {
                        if (resetR <= 100)
                        {
                            MessageBox.Show($@"熊證重設比 不得低於100");
                            dataOK = false;
                        }
                    }
                }


                if (UidDeltaOne_Temp.ContainsKey(underlyingID) && confirmed)
                {
                    if (cp == "C")
                    {
                        if (IsSpecial.Contains(underlyingID))
                        {
                            if (Market30.Contains(underlyingID))//市值前30  DeltaOne*股價<5億
                            {
                                if (cr * shares * spot > ISTOP30MaxIssue)
                                {
                                    MessageBox.Show($"{underlyingID} 為風險標的且市值前30大標的，DeltaOne市值已超過{(int)(ISTOP30MaxIssue/100000)}億");
                                    dataOK = false;
                                    continue;
                                }
                            }
                            else
                            {
                                if (cr * shares * spot > NonTOP30MaxIssue)
                                {
                                    MessageBox.Show($"{underlyingID} 為風險標的 DeltaOne市值已超過{(int)(NonTOP30MaxIssue / 100000)}億");
                                    dataOK = false;
                                    continue;
                                }
                            }
                        }
                    }
                    
                    if (cp == "P")
                    {
                        if (IsSpecial.Contains(underlyingID))
                        {
                            if (Market30.Contains(underlyingID))//市值前30  DeltaOne*股價<5億
                            {
                                if (cr * shares * spot > ISTOP30MaxIssue)
                                {
                                    MessageBox.Show($"{underlyingID} 為風險標的且市值前30大標的，DeltaOne市值已超過{(int)(ISTOP30MaxIssue / 100000)}億");
                                    dataOK = false;
                                    continue;
                                }
                            }
                            else
                            {
                                if (cr * shares * spot > NonTOP30MaxIssue)
                                {
                                    MessageBox.Show($"{underlyingID} 為風險標的 DeltaOne市值已超過{(int)(NonTOP30MaxIssue / 100000)}億");
                                    dataOK = false;
                                    continue;
                                }
                            }
                        }
                        if (dicAuthorize.ContainsKey(underlyingID))
                        {
                            if (IsSpecial.Contains(underlyingID) && (double)UidDeltaOne_Temp[underlyingID].KgiCallDeltaOne / (double)UidDeltaOne_Temp[underlyingID].KgiPutDeltaOne < dicAuthorize[underlyingID])
                            {
                                MessageBox.Show($"{underlyingID} 為風險標的，自家權證 Call/Put DeltaOne比例 < {dicAuthorize[underlyingID]}");
                                dataOK = false;
                                continue;
                            }
                            else if ((double)UidDeltaOne_Temp[underlyingID].KgiCallDeltaOne / (double)UidDeltaOne_Temp[underlyingID].KgiPutDeltaOne < dicAuthorize[underlyingID])
                            {
                                if (!IsIndex.Contains(underlyingID))
                                {
                                    MessageBox.Show($"{underlyingID} 自家權證 Call/Put DeltaOne比例 < {dicAuthorize[underlyingID]}");
                                    dataOK = false;
                                    continue;
                                }
                            }
                        }
                        else
                        {
                            //考慮發Put的時候不能把今天要發的Call算進來
                            if (IsSpecial.Contains(underlyingID) && (double)UidDeltaOne_Temp[underlyingID].KgiCallDeltaOne / (double)UidDeltaOne_Temp[underlyingID].KgiPutDeltaOne < SpecialCallPutRatio)
                            {
                                MessageBox.Show($"{underlyingID} 為風險標的，自家權證 Call/Put DeltaOne比例 < {SpecialCallPutRatio}");
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
                    }
                }
            }

            string sql_checkPutDeltaOne = $@"SELECT [UID],[Ratio],B.KgiPutDeltaOne,B.AllPutDeltaOne
                                              FROM [WarrantAssistant].[dbo].[PutDeltaOneRatioMonitor]  AS A
                                              LEFT JOIN (SELECT [UnderlyingID],[KgiPutDeltaOne],[AllPutDeltaOne] FROM [WarrantAssistant].[dbo].[WarrantIssueDeltaOne]
                                              WHERE [DateTime] = (SELECT MAX([DateTime]) FROM [WarrantAssistant].[dbo].[WarrantIssueDeltaOne])) AS B ON A.[UID] = B.UnderlyingID";
            DataTable dt_checkPutDeltaOne = MSSQL.ExecSqlQry(sql_checkPutDeltaOne, GlobalVar.loginSet.warrantassistant45);

            foreach(DataRow dr_checkPutDeltaOne in dt_checkPutDeltaOne.Rows)
            {
                string uid = dr_checkPutDeltaOne["UID"].ToString();
                double kgiPutDeltaOne = Convert.ToDouble(dr_checkPutDeltaOne["KgiPutDeltaOne"].ToString());
                double allPutDeltaOne = Convert.ToDouble(dr_checkPutDeltaOne["AllPutDeltaOne"].ToString());
                double ratio = Convert.ToDouble(dr_checkPutDeltaOne["Ratio"].ToString());
                if (PutDeltaOne.ContainsKey(uid))
                {
                    if(((kgiPutDeltaOne + PutDeltaOne[uid]) / allPutDeltaOne) * 100 > ratio)
                    {
                        MessageBox.Show($@"{uid} PutdeltaOne 市佔超過{ratio}!");
                        dataOK = false;
                    }
                }
            }


            if (!dataOK)
                return false;

            sql2 = $@"SELECT [UnderlyingID], TraderID
            FROM [WarrantAssistant].[dbo].[ApplyTempList]
            WHERE  [UserID]='{userID}' and [ConfirmChecked]='Y' and [UserID] <> TraderID ";
            badParam = MSSQL.ExecSqlQry(sql2, conn);
            foreach (DataRow Row in badParam.Rows)
            {
                string uid = Row["UnderlyingID"].ToString();
                differentTraderUid += uid + "/";
                differentTrader = true;
            }
            if (differentTrader)
            {
                if (DialogResult.No == MessageBox.Show($@"與使用者不同{differentTraderUid}，是否確認發行?", "確認發行", MessageBoxButtons.YesNo))
                    return false;
            }
            return dataOK;
        }
#region UpdateData()
        private void UpdateData() {//會有編輯紀錄
            try {

                int availables = 0;
                int i = 1;
                //更新權證編號，被刪除的代號會跳過
                foreach (Infragistics.Win.UltraWinGrid.UltraGridRow r in ultraGrid1.Rows)
                {
                    string serialNum = r.Cells["編號"].Value.ToString();
                    if (serialNum != "")
                    {
                        int index = Convert.ToInt32(serialNum.Substring(17, serialNum.Length - 17));
                        if (i <= index)
                            i = index;
                    }
                }
                foreach(int ii in DeletedSerialNum)
                {
                    if (i <= ii)
                        i = ii;
                }
                i++;
                applyCount = 0;
                int havingReset = 0;
                //SQLCommandHelper h = new SQLCommandHelper(GlobalVar.loginSet.warrantassistant45, sql, ps);
                List<string> sqlInserts = new List<string>();
                sqlInserts.Clear();
                foreach (Infragistics.Win.UltraWinGrid.UltraGridRow r in ultraGrid1.Rows) {
                    string underlyingID = r.Cells["標的代號"].Value.ToString();
                
                    if (underlyingID != "") {
#if deletelog
                        /*
                        while (true)
                        {
                            if (!DeletedSerialNum.Contains(i))
                                break;
                            i++;
                        }
                        */
#endif

                        //string serialNumber = DateTime.Today.ToString("yyyyMMdd") + userID + "01" + i.ToString("0#");
                        string serialNumber = r.Cells["編號"].Value.ToString();
                        if (serialNumber == "")
                        {
                            serialNumber = DateTime.Today.ToString("yyyyMMdd") + userID + "01" + i.ToString("0#");
                            i++;
                        }
                        string underlyingName = r.Cells["標的名稱"].Value.ToString();
                        double k = r.Cells["履約價"].Value == DBNull.Value ? 0 : Convert.ToDouble(r.Cells["履約價"].Value);
                      
                        int t = r.Cells["期間(月)"].Value == DBNull.Value ? 6 : Convert.ToInt32(r.Cells["期間(月)"].Value);
                        double cr = r.Cells["行使比例"].Value == DBNull.Value ? 0 : Convert.ToDouble(r.Cells["行使比例"].Value);
                        double hv = r.Cells["HV"].Value == DBNull.Value ? 0 : Convert.ToDouble(r.Cells["HV"].Value);
                        double iv = r.Cells["IV"].Value == DBNull.Value ? 0 : Convert.ToDouble(r.Cells["IV"].Value);
                        double issueNum = r.Cells["張數"].Value == DBNull.Value ? 10000 : Convert.ToDouble(r.Cells["張數"].Value);
                        double resetR = r.Cells["重設比"].Value == DBNull.Value ? 0 : Convert.ToDouble(r.Cells["重設比"].Value);
                        double barrierR = r.Cells["界限比"].Value == DBNull.Value ? 0 : Convert.ToDouble(r.Cells["界限比"].Value);
                        double financialR = r.Cells["財務費用"].Value == DBNull.Value ? 0 : Convert.ToDouble(r.Cells["財務費用"].Value);
                        double adj = r.Cells["Adj"].Value == DBNull.Value ? 0 : Convert.ToDouble(r.Cells["Adj"].Value);
                        string type = r.Cells["類型"].Value.ToString();
                        double underlyingPrice = r.Cells["股價"].Value == DBNull.Value ? 0 : Convert.ToDouble(r.Cells["股價"].Value);
                        string ex = r.Cells["說明"].Value == DBNull.Value ? "" : r.Cells["說明"].Value.ToString();
                        //if (type != "一般型" && type != "牛熊證" && type != "重設型") {
                        if (type != "一般型" && type != "牛熊證" && type != "重設型" && type != "展延型")
                        {
                            if (type == "2")
                                type = "牛熊證";
                            else if (type == "3")
                                type = "重設型";
                            else if (type == "4")
                                type = "展延型";
                            else
                                type = "一般型";
                        }

                        string cp = r.Cells["CP"].Value.ToString();
                        if (cp != "C" && cp != "P") {
                            if (cp == "2")
                                cp = "P";
                            else
                                cp = "C";
                        }
                        //if(type == "重設型")
                        
                        if (type == "重設型"|| type == "展延型")
                        {
                            havingReset = 1;
                            if (cp == "P")
                                k = Math.Round(underlyingPrice * (resetR / 100), 2);
                            else
                                k = Math.Round(underlyingPrice * (resetR / 100), 2);
                        }
                        
                        bool isReward = r.Cells["獎勵"].Value == DBNull.Value ? false : Convert.ToBoolean(r.Cells["獎勵"].Value);
                        string useReward = "N";
                        if (isReward)
                            useReward = "Y";

                        bool confirmed = r.Cells["確認"].Value == DBNull.Value ? false : Convert.ToBoolean(r.Cells["確認"].Value);
                        string confirmChecked = "N";
                        if (confirmed) {
                            confirmChecked = "Y";
                            applyCount++;
                        }
                        bool deleted = r.Cells["刪除"].Value == DBNull.Value ? false : Convert.ToBoolean(r.Cells["刪除"].Value);
                       
                        string deletedStr = "N";
                        if (deleted)
                        {
                            deletedStr = "Y";
                        }
                        bool apply1500Wbool = r.Cells["1500W"].Value == DBNull.Value ? false : Convert.ToBoolean(r.Cells["1500W"].Value);
                        string apply1500W = "N";
                        if (apply1500Wbool)
                            apply1500W = "Y";


                        DateTime expiryDate = GlobalVar.globalParameter.nextTradeDate3.AddMonths(t);
                        if (expiryDate.Day == GlobalVar.globalParameter.nextTradeDate3.Day)
                            expiryDate = expiryDate.AddDays(-1);
                        string sqlTemp = $"SELECT TOP 1 TradeDate from [DeriPosition].[dbo].[Calendar] WHERE IsTrade='Y' AND [CountryId] = 'TWN' AND TradeDate >= '{expiryDate.ToString("yyyy-MM-dd")}'";
                        //DataView dvTemp = DeriLib.Util.ExecSqlQry(sqlTemp, GlobalVar.loginSet.tsquoteSqlConnString);
                        DataTable dvTemp = MSSQL.ExecSqlQry(sqlTemp, GlobalVar.loginSet.tsquoteSqlConnString);
                        foreach (DataRow drTemp in dvTemp.Rows) {
                            expiryDate = Convert.ToDateTime(drTemp["TradeDate"]);
                        }
                        int month = expiryDate.Month;
                        string expiryMonth = month.ToString();
                        if (month >= 10) {
                            if (month == 10)
                                expiryMonth = "A";
                            if (month == 11)
                                expiryMonth = "B";
                            if (month == 12)
                                expiryMonth = "C";
                        }
                        string expiryYear = expiryDate.AddYears(-1).ToString("yyyy");
                        expiryYear = expiryYear.Substring(expiryYear.Length - 1, 1);

                        string warrantType = "";
                        string tempType = "";

                        if (type == "牛熊證") {
                            if (cp == "P") {
                                warrantType = "熊";
                                tempType = "4";
                            } else {
                                warrantType = "牛";
                                tempType = "3";
                            }
                        } else {
                            if (cp == "P") {
                                warrantType = "售";
                                tempType = "2";
                            } else {
                                warrantType = "購";
                                tempType = "1";
                            }
                        }

                        string tempName = underlyingName + "凱基" + expiryYear + expiryMonth + warrantType;

                        string traderID = r.Cells["交易員"].Value == DBNull.Value ? userID : r.Cells["交易員"].Value.ToString();

                        double ivNew = r.Cells["IV*"].Value == DBNull.Value ? 0.0 : (double) r.Cells["IV*"].Value;
                      
           
                        
                        string sqlInsert = $@"INSERT INTO [ApplyTempList] (SerialNum, UnderlyingID, K, T, R, HV, IV, IssueNum, ResetR, BarrierR, FinancialR, Type, CP, UseReward, ConfirmChecked, Apply1500W, UserID, MDate, TempName, TempType, TraderID, IVNew, Adj, 說明, [Delete]) 
                                    VALUES('{serialNumber}', '{underlyingID}', {k}, {t}, {cr}, {hv}, {iv}, {issueNum}, {resetR}, {barrierR},{financialR}, '{type}', '{cp}', '{useReward}', '{confirmChecked}', '{apply1500W}', '{userID}', GETDATE(), '{tempName}' , '{tempType}', '{traderID}', {ivNew}, {adj}, '{ex}', '{deletedStr}')";
                        sqlInserts.Add(sqlInsert);

                    }
                }
                MessageBox.Show("發行成功，將刪除資料表");
                MSSQL.ExecSqlCmd($"DELETE FROM [ApplyTempList] WHERE UserID='{userID}'", conn);
                foreach (string str in sqlInserts)
                {
                    try
                    {
                        MSSQL.ExecSqlCmd(str, conn);
                    }
                    catch(Exception ex2)
                    {
                        MessageBox.Show(ex2.Message + " " + str);
                        continue;
                    }
                }
                MessageBox.Show("已更新資料表");
                //h.Dispose();
                GlobalUtility.LogInfo("Log", GlobalVar.globalParameter.userID + " 編輯/更新" + (i - 1) + "檔發行");
                if(havingReset == 1)
                    MessageBox.Show("有重設型OR展延型，履約價已調整為(C / P)股價的(150% / 50%) 會以申請當日收盤價為準");
            } catch (Exception ex) {
                MessageBox.Show(ex.Message);
            }
        }
#endregion
        private void OfficiallyApply() {
            try {
                
                UpdateData();
                LoadData();//可能其他表有更改資料，先Load同步資料
               
                if (!CheckData())
                    return;
                

                string sql1 = $"DELETE FROM [WarrantAssistant].[dbo].[ApplyOfficial] WHERE [UserID]='{userID}'";
                string sql2 = @"INSERT INTO [WarrantAssistant].[dbo].[ApplyOfficial] ([SerialNumber],[UnderlyingID],[K],[T],[R],[HV],[IV],[IssueNum],[ResetR],[BarrierR],[FinancialR],[Type],[CP],[UseReward],[Apply1500W],[TempName],[TraderID],[MDate],UserID, IVNew, [說明])
                                SELECT [SerialNum],[UnderlyingID],[K],[T],[R],[HV],[IV],[IssueNum],[ResetR],[BarrierR],[FinancialR],[Type],[CP],[UseReward],[Apply1500W],[TempName],[TraderID],[MDate],UserID, IVNew ,[說明] "
                //sql2 += "'"+userID + "', [MDate]" ;
                 + " FROM [WarrantAssistant].[dbo].[ApplyTempList]"
                 + $" WHERE [UserID]='{userID}' AND [ConfirmChecked]='Y'";

                string sql3 = $"DELETE FROM [WarrantAssistant].[dbo].[ApplyTotalList] WHERE [UserID]='{userID}' AND [ApplyKind]='1'";
                string sql4 = $@"INSERT INTO [WarrantAssistant].[dbo].[ApplyTotalList] ([ApplyKind],[SerialNum],[Market],[UnderlyingID],[WarrantName],[CR] ,[IssueNum],[EquivalentNum],[Credit],[RewardCredit],[Type],[CP],[UseReward],[MarketTmr],[TraderID],[MDate],UserID)
                                SELECT '1',A.SerialNumber, isnull(B.Market, 'TSE'), A.UnderlyingID, A.TempName, A.R, A.IssueNum, ROUND(A.R*A.IssueNum, 2), ISNULL(E.CanIssue,0), Floor(E.[WarrantAvailableShares] * {GlobalVar.globalParameter.givenRewardPercent} - IsNull(F.[UsedRewardNum],0)), A.Type, A.CP, A.UseReward,'N', A.TraderID, GETDATE(), A.UserID
                                FROM [WarrantAssistant].[dbo].[ApplyOfficial] A
                                LEFT JOIN [WarrantAssistant].[dbo].[WarrantUnderlyingSummary] B ON A.UnderlyingID=B.UnderlyingID
                                LEFT JOIN (SELECT [UID], [CanIssue], [WarrantAvailableShares] FROM [WarrantAssistant].[dbo].[WarrantUnderlyingCreditNew] WHERE [UpdateTime] > '{DateTime.Today.ToString("yyyyMMdd")}' ) as E on A.UnderlyingID = E.[UID]
                                LEFT JOIN [WarrantAssistant].[dbo].[WarrantReward] F on A.UnderlyingID=F.UnderlyingID
                                WHERE a.[UserID]='{userID}'";
                conn.Open();
                MSSQL.ExecSqlCmd(sql1, conn);
                MSSQL.ExecSqlCmd(sql2, conn);
                MSSQL.ExecSqlCmd(sql3, conn);
                MSSQL.ExecSqlCmd(sql4, conn);
                conn.Close();
                //------------------------------------------------------

                //string sql5 = "SELECT [SerialNum], [WarrantName] FROM [WarrantAssistant].[dbo].[ApplyTotalList] WHERE [ApplyKind]='1' AND UserID='" + userID + "' ORDER BY CONVERT(INT, SUBSTRING([SerialNum], 18, LEN([SerialNum])-18 + 1))";
                string sql5 = $@"SELECT A.[SerialNum], A.[WarrantName],ISNULL(B.說明,'') 說明 ,B.[IV] ,B.[FinancialR]
                                  FROM [WarrantAssistant].[dbo].[ApplyTotalList] AS A
                                  LEFT JOIN [WarrantAssistant].[dbo].[ApplyOfficial] AS B ON A.SerialNum = B.SerialNumber
                                  WHERE [ApplyKind]='1' AND A.UserID='{userID}' ORDER BY CONVERT(INT, SUBSTRING([SerialNum], 18, LEN([SerialNum])-18 + 1))";
                DataTable dv = MSSQL.ExecSqlQry(sql5, GlobalVar.loginSet.warrantassistant45);

                string cmdText = "UPDATE [ApplyTotalList] SET WarrantName=@WarrantName WHERE SerialNum=@SerialNum";
                List<SqlParameter> pars = new List<SqlParameter> {
                    new SqlParameter("@WarrantName", SqlDbType.VarChar),
                    new SqlParameter("@SerialNum", SqlDbType.VarChar)
                };

                SQLCommandHelper h = new SQLCommandHelper(GlobalVar.loginSet.warrantassistant45, cmdText, pars);
                int where = 0;
      
                try
                {
                    foreach (DataRow dr in dv.Rows)
                    {
                        string serialNum = dr["SerialNum"].ToString();
                        string warrantName = dr["WarrantName"].ToString();
                        string log = dr["說明"].ToString();
                        string initial_iv = dr["IV"].ToString();
                        string financialR = dr["FinancialR"].ToString();
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
                        h.SetParameterValue("@WarrantName", warrantName);
                        h.SetParameterValue("@SerialNum", serialNum);
                        h.ExecuteCommand();

                        string sql_IssueLog = $@"IF EXISTS(SELECT [WarrantName] FROM [WarrantAssistant].[dbo].[WarrantBasic_InitialIV] where [WarrantName] ='{warrantName}')
	                                                    UPDATE [WarrantAssistant].[dbo].[WarrantBasic_InitialIV] SET [InitialIV] ={initial_iv},[ApplyDate]='{DateTime.Today.ToString("yyyyMMdd")}',[說明]='{log}',[財務費用率] = {financialR}  WHERE [WarrantName] ='{warrantName}'
                                                      ELSE 
                                                        INSERT [WarrantAssistant].[dbo].[WarrantBasic_InitialIV]
	                                                    VALUES ('{warrantName}',{initial_iv},'{DateTime.Today.ToString("yyyyMMdd")}','{log}',{financialR},'',0,0,0,0,'')";
                        int bo = MSSQL.ExecSqlCmd(sql_IssueLog, GlobalVar.loginSet.warrantassistant45);
                    }
                }
                catch(Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
                h.Dispose();
                //UpdateRecord();
                toolStripLabel2.Text = DateTime.Now + "申請" + applyCount + "檔權證發行成功";

                string sql_WM = $@"SELECT A.[UnderlyingID], A.[WarrantName], A.[CR], A.[IssueNum], B.[MPrice]
                                FROM [WarrantAssistant].[dbo].[ApplyTotalList] AS A
                                LEFT JOIN [WarrantAssistant].[dbo].[WarrantPrices] AS B ON A.[UnderlyingID] = B.[CommodityID]
                                WHERE [ApplyKind] ='1' AND [UserID] = '{userID}'";
                DataTable dv_WM = MSSQL.ExecSqlQry(sql_WM, GlobalVar.loginSet.warrantassistant45);

                foreach (DataRow dr in dv_WM.Rows)
                {
                    double cr = Convert.ToDouble(dr["CR"].ToString());
                    double issueNum = Convert.ToDouble(dr["IssueNum"].ToString());
                    double spot = Convert.ToDouble(dr["MPrice"].ToString());
                    string uid = dr["UnderlyingID"].ToString();
                    string wname = dr["WarrantName"].ToString();
                    if (!WMarket30.Contains(wname) && Market30.Contains(uid)&&IsSpecial.Contains(uid))
                    {
                        if(cr * issueNum * spot > 200000)
                            GlobalUtility.LogInfo("TOP30", $"{GlobalVar.globalParameter.userID}申請{wname}為風險標的且為市值前30大標的");
                    }
                }
                GlobalUtility.LogInfo("Info", GlobalVar.globalParameter.userID + " 申請" + applyCount + "檔權證發行");
                MessageBox.Show($"申請 {applyCount} 檔權證發行成功!");
                LoadData();
                
                foreach(string serial in IssuedSerialNum.Keys)
                {

                    string sqlSearch = $@"SELECT [SerialNum]
                                          FROM [WarrantAssistant].[dbo].[ApplyTotalList]
                                          WHERE [SerialNum] = '{serial}'";
                    DataTable dtSearch = MSSQL.ExecSqlQry(sqlSearch, GlobalVar.loginSet.warrantassistant45);
                    if (dtSearch.Rows.Count < 1)
                    {
                        GlobalUtility.LogInfo("Info", $@"{GlobalVar.globalParameter.userID} 刪除 {IssuedSerialNum_UID[serial]} {IssuedSerialNum[serial]}");
                    }
                }
                
                LoadIssuedSerial();
                Thread th = new Thread(new ThreadStart(UpdateRecord));
                th.Start();
                
            } catch (Exception ex) {
                MessageBox.Show(ex.Message);
            }
        }
        private void UpdateRecord()//紀錄任何權證發行與修改，存在ApplyTotalRecord
        {
            //抓更新紀錄的資料，扣除DELETE
            lock (_thisLock)
            {
                //MessageBox.Show("start record");
                Dictionary<string, Dictionary<string, string>> dt_record_value = new Dictionary<string, Dictionary<string, string>>();
                Dictionary<string, Dictionary<string, int>> dt_record_count = new Dictionary<string, Dictionary<string, int>>();
 
                string sql_record = $@"SELECT A.[SerialNumber], A.[DataName], A.[UpdateCount], B.[ToValue]
                                FROM (SELECT [SerialNumber], [DataName], MAX([UpdateCount]) AS UpdateCount
                                        FROM [WarrantAssistant].[dbo].[ApplyTotalRecord]
                                        WHERE [UpdateTime] >= CONVERT(varchar, getdate(), 112) 
                                        GROUP BY [SerialNumber] ,[DataName]) AS A
                                LEFT JOIN [WarrantAssistant].[dbo].[ApplyTotalRecord] AS B 
                                ON A.[SerialNumber] = B.[SerialNumber] AND A.[DataName] =B.[DataName] AND A.[UpdateCount] =B.[UpdateCount]
                                WHERE A.[UpdateCount] < 9999 AND A.[DataName]!='' AND (B.[TraderID] ='{userID}' OR B.[TraderID] ='0004260' OR B.[TraderID] ='0009148') AND B.UpdateTime > '{DateTime.Today.ToString("yyyyMMdd")}'";
                
                DataTable dv_record = MSSQL.ExecSqlQry(sql_record, GlobalVar.loginSet.warrantassistant45);
                
                foreach (DataRow dr in dv_record.Rows)
                {
                    string serialNum = dr["SerialNumber"].ToString();

                    string dataName = dr["DataName"].ToString();//EX CR
                    int updateCount = Convert.ToInt32(dr["UpdateCount"].ToString());
                    string toValue = dr["ToValue"].ToString();
                    if (!dt_record_value.ContainsKey(serialNum))
                    {
                        dt_record_value.Add(serialNum, new Dictionary<string, string>());
                        dt_record_count.Add(serialNum, new Dictionary<string, int>());
                    }
                    dt_record_value[serialNum].Add(dataName, toValue);//EX. CR 0.08
                    dt_record_count[serialNum].Add(dataName, updateCount);//EX CR 0
                }


                updaterecord_dt.Rows.Clear();
             
                int length = dt.Rows.Count;
                try
                {
                    
                    for (int i = length - 1; i >= 0; i--)
                    {
                        string serialNum = dt.Rows[i][1].ToString();//抓序號

                        //string serialNum = dt.Rows[i].Cells[0].Value.ToString();//抓序號
                        /*
                        if (ultraGrid1.Rows[i].Cells[18].Value.ToString() == "False")
                        {
                            MessageBox.Show($"dttt {dt.Rows[i][18].ToString()}");
                            continue;
                        }
                        */
                        if (dt.Rows[i][19].ToString() == "False")
                        {
                            //MessageBox.Show($"dttt {dt.Rows[i]["編號"].ToString()}");
                            continue;
                        }
                        if (!dt_record_value.ContainsKey(serialNum))//沒有在備份表中，要新增
                        {
                            //MessageBox.Show($"new  {serialNum}");
                            //insert ADD 新權證
       
                            foreach (string colname in sqlTogrid.Keys)//11種欄位
                            {
    
                                string sql4 = $@"INSERT INTO [WarrantAssistant].[dbo].[ApplyTotalRecord] ([UpdateTime], [UpdateType], [TraderID], [SerialNumber]
                                , [ApplyKind], [DataName] ,[FromValue], [ToValue], [UpdateCount])
                                VALUES(GETDATE(), 'ADD', '{userID}', {serialNum}, '1','{sqlTogrid[colname]}','{ultraGrid1.Rows[i].Cells[colname].Value.ToString()}','{ultraGrid1.Rows[i].Cells[colname].Value.ToString()}',0)";
                                MSSQL.ExecSqlCmd(sql4, GlobalVar.loginSet.warrantassistant45);
                            }
    
                        }
                        else
                        {
                            foreach (string colname in sqlTogrid.Keys)
                            {
                                //if (dt_record_value[serialNum][sqlTogrid[colname]].ToString() != ultraGrid1.Rows[i].Cells[colname].Value.ToString())
                                if (dt_record_value[serialNum][sqlTogrid[colname]].ToString() != dt.Rows[i][colname].ToString())
                                {
                                    DataRow dr = updaterecord_dt.NewRow();
                                    dr["serialNum"] = serialNum;
                                    dr["dataname"] = colname;//EX 期間(月)
                                    dr["fromvalue"] = dt_record_value[serialNum][sqlTogrid[colname]].ToString();
                                    dr["tovalue"] = ultraGrid1.Rows[i].Cells[colname].Value.ToString();
                                    updaterecord_dt.Rows.Add(dr);
                                }
                            }
                        }
                    }
                }
                catch(Exception ex)
                {
                    MessageBox.Show($@"{ex.Message}  in 更新修正權證");
                }
                //MessageBox.Show($"finish compare {updaterecord_dt.Rows.Count}");

                try
                {
                   
                    foreach (DataRow dr in updaterecord_dt.Rows)
                    {
                        //MessageBox.Show($"{dr["serialNum"].ToString()}  {dr["dataname"].ToString()}  {dr["fromvalue"].ToString()}  {dr["tovalue"].ToString()}");

                        int count = -1;
                        string serialNum = dr["serialNum"].ToString();
                        string dataname = dr["dataname"].ToString();//期間(月)
                        string fromvalue = dr["fromvalue"].ToString();
                        string tovalue = dr["tovalue"].ToString();
                        //MessageBox.Show($"{serialNum}  {dataname}  {fromvalue}  {tovalue}");
                        /*
                        string sql_count = $@"SELECT A.[UpdateCount]
                                            FROM (SELECT [SerialNumber], [DataName], MAX([UpdateCount]) AS UpdateCount
                                                  FROM [EDIS].[dbo].[ApplyTotalRecord]
                                                  WHERE [UpdateTime] >= CONVERT(varchar, getdate(), 112)
                                                  GROUP BY[SerialNumber],[DataName]) AS A
                                            WHERE A.[UpdateCount] < 9999 AND A.[SerialNumber] ='{serialNum}' AND A.DataName= '{dataname}'";
                        DataTable dv_count = MSSQL.ExecSqlQry(sql_count, GlobalVar.loginSet.edisSqlConnString);
                        foreach(DataRow drr in dv_count.Rows)
                        {
                             count = Convert.ToInt32(drr["UpdateCount"].ToString());
                        }
                        */
                        string sql3 = $@"INSERT INTO [WarrantAssistant].[dbo].[ApplyTotalRecord] ([UpdateTime], [UpdateType], [TraderID], [SerialNumber]
                                        , [ApplyKind], [DataName] ,[FromValue], [ToValue], [UpdateCount])
                                        VALUES(GETDATE(), 'UPDATE', '{userID}', {serialNum}, '1','{sqlTogrid[dataname]}','{fromvalue}','{tovalue}',{dt_record_count[serialNum][sqlTogrid[dataname]] + 1})";
                        MSSQL.ExecSqlCmd(sql3, GlobalVar.loginSet.warrantassistant45);
                    }
                }
                catch(Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }
        }


        private void SetButton() {
            UltraGridBand bands0 = ultraGrid1.DisplayLayout.Bands[0];
            //this.ultraGrid1.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti;
            if (isEdit) {
                bands0.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.TemplateOnBottom;
                bands0.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
                bands0.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.True;

                bands0.Columns["編號"].CellActivation = Activation.NoEdit;
                bands0.Columns["標的代號"].CellActivation = Activation.AllowEdit;
                bands0.Columns["履約價"].CellActivation = Activation.AllowEdit;
                bands0.Columns["期間(月)"].CellActivation = Activation.AllowEdit;
                bands0.Columns["行使比例"].CellActivation = Activation.AllowEdit;
                
                bands0.Columns["HV"].CellActivation = Activation.AllowEdit;
                bands0.Columns["IV"].CellActivation = Activation.AllowEdit;
                //bands0.Columns["建議Vol"].CellActivation = Activation.NoEdit;
                //bands0.Columns["下限Vol"].CellActivation = Activation.NoEdit;
                bands0.Columns["張數"].CellActivation = Activation.AllowEdit;
                bands0.Columns["重設比"].CellActivation = Activation.AllowEdit;
                bands0.Columns["界限比"].CellActivation = Activation.AllowEdit;
                bands0.Columns["財務費用"].CellActivation = Activation.AllowEdit;
                bands0.Columns["類型"].CellActivation = Activation.AllowEdit;
                bands0.Columns["CP"].CellActivation = Activation.AllowEdit;
                //bands0.Columns["發行原因"].CellActivation = Activation.AllowEdit;
                bands0.Columns["交易員"].CellActivation = Activation.AllowEdit;
                bands0.Columns["1500W"].CellActivation = Activation.AllowEdit;
                bands0.Columns["發行價格"].CellActivation = Activation.AllowEdit;
                bands0.Columns["說明"].CellActivation = Activation.AllowEdit;
                //ultraGrid1.DisplayLayout.Bands[0].Columns["標的名稱"].CellActivation = Activation.AllowEdit;
                //ultraGrid1.DisplayLayout.Bands[0].Columns["股價"].CellActivation = Activation.AllowEdit;
                //ultraGrid1.DisplayLayout.Bands[0].Columns["市場"].CellActivation = Activation.AllowEdit;
                //ultraGrid1.DisplayLayout.Bands[0].Columns["約當張數"].CellActivation = Activation.AllowEdit;
                //ultraGrid1.DisplayLayout.Bands[0].Columns["今日額度"].CellActivation = Activation.AllowEdit;
                //ultraGrid1.DisplayLayout.Bands[0].Columns["獎勵額度"].CellActivation = Activation.AllowEdit;


                bands0.Columns["昨日KgiCall DeltaOne"].CellActivation = Activation.NoEdit;
                bands0.Columns["昨日KgiPut DeltaOne"].CellActivation = Activation.NoEdit;
                bands0.Columns["今日KgiCall DeltaOne"].CellActivation = Activation.NoEdit;
                bands0.Columns["今日KgiPut DeltaOne"].CellActivation = Activation.NoEdit;
                bands0.Columns["自家 Call/Put DeltaOne比例"].CellActivation = Activation.NoEdit;
                bands0.Columns["自家/市場Put DeltaOne比例>25%"].CellActivation = Activation.NoEdit;
                bands0.Columns["Put大於元大"].CellActivation = Activation.NoEdit;
                

                bands0.Columns["編號"].Hidden = true;
                bands0.Columns["覆"].Hidden = true;

                buttonEdit.Visible = false;
                buttonUpdatePrice.Visible = false;
                buttonConfirm.Visible = true;
                buttonDelete.Visible = true;
                buttonCancel.Visible = true;
                toolStripButton1.Visible = false;
                toolStripSeparator2.Visible = false;
                toolStripButton2.Visible = false;
                toolStripSeparator3.Visible = false;
                toolStripButton3.Visible = false;

                ultraGrid1.DisplayLayout.Bands[0].Columns["確認"].Hidden = true;
                //ultraGrid1.DisplayLayout.Bands[0].Columns["獎勵"].Hidden = true;
                //ultraGrid1.DisplayLayout.Bands[0].Columns["1500W"].Hidden = true;

                ultraGrid1.DisplayLayout.Bands[0].Columns["市場"].Hidden = true;
                ultraGrid1.DisplayLayout.Bands[0].Columns["約當張數"].Hidden = true;
                //ultraGrid1.DisplayLayout.Bands[0].Columns["今日額度"].Hidden = true;
                //ultraGrid1.DisplayLayout.Bands[0].Columns["獎勵額度"].Hidden = true;
                //ultraGrid1.DisplayLayout.Bands[0].Columns["HV"].Hidden = true;


                bands0.Columns["昨日KgiCall DeltaOne"].Hidden = true;
                bands0.Columns["昨日KgiPut DeltaOne"].Hidden = true;
                bands0.Columns["今日KgiCall DeltaOne"].Hidden = true;
                bands0.Columns["今日KgiPut DeltaOne"].Hidden = true;
                bands0.Columns["自家 Call/Put DeltaOne比例"].Hidden = true;
                bands0.Columns["自家/市場Put DeltaOne比例>25%"].Hidden = true;
                bands0.Columns["Put大於元大"].Hidden = true;


            } else {
                bands0.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
                bands0.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
                bands0.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
                bands0.Columns["編號"].CellActivation = Activation.NoEdit;
                bands0.Columns["確認"].CellActivation = Activation.AllowEdit;
                bands0.Columns["1500W"].CellActivation = Activation.AllowEdit;
                bands0.Columns["標的代號"].CellActivation = Activation.NoEdit;
                bands0.Columns["履約價"].CellActivation = Activation.NoEdit;
                bands0.Columns["期間(月)"].CellActivation = Activation.NoEdit;
                bands0.Columns["行使比例"].CellActivation = Activation.NoEdit;
                bands0.Columns["建議行使比例"].CellActivation = Activation.NoEdit;
                bands0.Columns["HV"].CellActivation = Activation.NoEdit;
                bands0.Columns["IV"].CellActivation = Activation.NoEdit;
                //bands0.Columns["建議Vol"].CellActivation = Activation.NoEdit;
                //bands0.Columns["下限Vol"].CellActivation = Activation.NoEdit;
                bands0.Columns["張數"].CellActivation = Activation.NoEdit;
                bands0.Columns["重設比"].CellActivation = Activation.NoEdit;
                bands0.Columns["界限比"].CellActivation = Activation.NoEdit;
                bands0.Columns["財務費用"].CellActivation = Activation.NoEdit;
                bands0.Columns["類型"].CellActivation = Activation.NoEdit;
                bands0.Columns["CP"].CellActivation = Activation.NoEdit;
                bands0.Columns["交易員"].CellActivation = Activation.NoEdit;
                bands0.Columns["獎勵"].CellActivation = Activation.AllowEdit;
                bands0.Columns["發行價格"].CellActivation = Activation.NoEdit;
                //bands0.Columns["發行原因"].CellActivation = Activation.NoEdit;
                bands0.Columns["標的名稱"].CellActivation = Activation.NoEdit;
                bands0.Columns["股價"].CellActivation = Activation.NoEdit;
                bands0.Columns["Delta"].CellActivation = Activation.NoEdit;
                //joufan
                bands0.Columns["Theta"].CellActivation = Activation.NoEdit;
                bands0.Columns["跳動價差"].CellActivation = Activation.NoEdit;
                bands0.Columns["市場"].CellActivation = Activation.NoEdit;
                bands0.Columns["約當張數"].CellActivation = Activation.NoEdit;
                bands0.Columns["約當張數(5000張)"].CellActivation = Activation.NoEdit;
                bands0.Columns["今日額度"].CellActivation = Activation.NoEdit;
                bands0.Columns["獎勵額度"].CellActivation = Activation.NoEdit;
                bands0.Columns["IV*"].CellActivation = Activation.NoEdit;
                bands0.Columns["發行價格*"].CellActivation = Activation.NoEdit;
                bands0.Columns["跌停價*"].CellActivation = Activation.NoEdit;

                bands0.Columns["昨日KgiCall DeltaOne"].CellActivation = Activation.NoEdit;
                bands0.Columns["昨日KgiPut DeltaOne"].CellActivation = Activation.NoEdit;
                bands0.Columns["今日KgiCall DeltaOne"].CellActivation = Activation.NoEdit;
                bands0.Columns["今日KgiPut DeltaOne"].CellActivation = Activation.NoEdit;
                bands0.Columns["自家 Call/Put DeltaOne比例"].CellActivation = Activation.NoEdit;
                bands0.Columns["自家/市場Put DeltaOne比例>25%"].CellActivation = Activation.NoEdit;
                bands0.Columns["Put大於元大"].CellActivation = Activation.NoEdit;
                bands0.Columns["說明"].CellActivation = Activation.NoEdit;
                this.ultraGrid1.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti;

                buttonEdit.Visible = true;
                buttonConfirm.Visible = false;
                buttonDelete.Visible = false;
                buttonCancel.Visible = false;
                toolStripButton1.Visible = true;
                toolStripSeparator2.Visible = true;
                toolStripButton2.Visible = true;
                toolStripSeparator3.Visible = true;
                toolStripButton3.Visible = true;
                buttonUpdatePrice.Visible = true;
                bands0.Columns["編號"].Hidden = true;
                bands0.Columns["確認"].Hidden = false;
                bands0.Columns["獎勵"].Hidden = false;
                bands0.Columns["1500W"].Hidden = false;

                bands0.Columns["市場"].Hidden = false;
                bands0.Columns["約當張數"].Hidden = false;
                bands0.Columns["今日額度"].Hidden = false;
                bands0.Columns["獎勵額度"].Hidden = false;
                bands0.Columns["覆"].Hidden = false;
                //bands0.Columns["HV"].Hidden = true;

                bands0.Columns["昨日KgiCall DeltaOne"].Hidden = true;
                bands0.Columns["昨日KgiPut DeltaOne"].Hidden = true;
                bands0.Columns["今日KgiCall DeltaOne"].Hidden = true;
                bands0.Columns["今日KgiPut DeltaOne"].Hidden = true;
                bands0.Columns["自家 Call/Put DeltaOne比例"].Hidden = true;
                bands0.Columns["自家/市場Put DeltaOne比例>25%"].Hidden = true;
                bands0.Columns["Put大於元大"].Hidden = true;
            }
            //排除指數類權證交易員
            if(userID != "0011135")
            {
                //bands0.Columns["下限Vol"].Hidden = true;
                bands0.Columns["界限比"].Hidden = true;
                bands0.Columns["1500W"].Hidden = true;
                bands0.Columns["Adj"].Hidden = true;
            }
            
            //bands0.Columns["刪除"].Hidden = true;
            bands0.Columns["刪除"].CellActivation = Activation.AllowEdit;
        }

        private void UltraGrid1_InitializeLayout(object sender, InitializeLayoutEventArgs e) {
            ultraGrid1.DisplayLayout.Override.RowSelectorHeaderStyle = RowSelectorHeaderStyle.ColumnChooserButton;

            if (!e.Layout.ValueLists.Exists("MyValueList")) {
                ValueList v;
                v = e.Layout.ValueLists.Add("MyValueList");
                v.ValueListItems.Add(1, "一般型");
                v.ValueListItems.Add(2, "牛熊證");
                v.ValueListItems.Add(3, "重設型");
                v.ValueListItems.Add(4, "展延型");
            }
            e.Layout.Bands[0].Columns["類型"].ValueList = e.Layout.ValueLists["MyValueList"];

            if (!e.Layout.ValueLists.Exists("MyValueList2")) {
                ValueList v2;
                v2 = e.Layout.ValueLists.Add("MyValueList2");
                v2.ValueListItems.Add(1, "C");
                v2.ValueListItems.Add(2, "P");
            }
            e.Layout.Bands[0].Columns["CP"].ValueList = e.Layout.ValueLists["MyValueList2"];

            if (!e.Layout.ValueLists.Exists("MyValueList3")) {
                ValueList v3;
                v3 = e.Layout.ValueLists.Add("MyValueList3");
                foreach (var item in GlobalVar.globalParameter.traders)
                    v3.ValueListItems.Add(item, item);
            }
            e.Layout.Bands[0].Columns["交易員"].ValueList = e.Layout.ValueLists["MyValueList3"];

            if (!e.Layout.ValueLists.Exists("MyValueList4")) {
                ValueList v;
                v = e.Layout.ValueLists.Add("MyValueList4");
                v.ValueListItems.Add(0, " ");
                v.ValueListItems.Add(1, "技術面偏多，股價持續看好，因此發行認購權證吸引投資人。");
                v.ValueListItems.Add(2, "基本面良好，股價具有漲升的條件，因此發行認購權證吸引投資人。");
                v.ValueListItems.Add(3, "營運動能具提升潛力，因此發行認購權證吸引投資人。");
                v.ValueListItems.Add(4, "提供投資人槓桿避險工具");
                v.ValueListItems.Add(5, "持續針對不同的履約條件、存續期間及認購認售等發行新條件，提供投資人更多元投資選擇");
            }
            // e.Layout.Bands[0].Columns["發行原因"].ValueList = e.Layout.ValueLists["MyValueList4"];

        }
        //按編輯時複製一個備份的table
        private void ButtonEdit_Click(object sender, EventArgs e) {
            //LoadData();
            isEdit = true;
            SetButton();
        }
        //確認
        private void ButtonConfirm_Click(object sender, EventArgs e) {
            
            ultraGrid1.PerformAction(Infragistics.Win.UltraWinGrid.UltraGridAction.ExitEditMode);
            isEdit = false;
            //if (!CheckData())
            //   return;
            
            UpdateData();
           
            SetButton();
            MSSQL.ExecSqlCmd($@"INSERT INTO [WarrantAssistant].[dbo].[TempListDeleteLog] 
                                  SELECT CONVERT(VARCHAR,GETDATE(),112),[SerialNum],[UserID]
                                  FROM [WarrantAssistant].[dbo].[ApplyTempList]
                                  WHERE [Delete] = 'Y'", conn);
            MSSQL.ExecSqlCmd($@"DELETE from [WarrantAssistant].[dbo].[ApplyTempList] WHERE [UserID] = '{userID}' AND [Delete] = 'Y'", conn);
            LoadData();
        }

        private void ButtonDelete_Click(object sender, EventArgs e) {
            isEdit = true;

            DialogResult result = MessageBox.Show("將全部刪除，確定?", "刪除資料", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
            if (result == DialogResult.Yes) {
                int length = ultraGrid1.Rows.Count;
                for(int i = length - 1; i >= 0; i--)
                {
                    string serial = ultraGrid1.Rows[i].Cells["編號"].Value.ToString();
                    if (serial.Length > 0)
                    {
                        if (ultraGrid1.Rows[i].Cells["確認"].Value.ToString() == "True")
                        {

                            foreach (var key in sqlTogrid.Keys)
                            {

                                string sql_delete2 = $@"INSERT INTO [WarrantAssistant].[dbo].[ApplyTotalRecord] ([UpdateTime], [UpdateType], [TraderID], [SerialNumber]
                                           , [ApplyKind], [DataName] ,[FromValue], [ToValue], [UpdateCount])
                                          VALUES(GETDATE(), 'DELETE', '{userID}', {serial}, '1','{sqlTogrid[key]}','','',9999)";
                                MSSQL.ExecSqlCmd(sql_delete2, GlobalVar.loginSet.warrantassistant45);

                            }

                        }

                        string sql2 = $@"INSERT INTO [WarrantAssistant].[dbo].[TempListDeleteLog] ([DateTime],[SerialNum],[Trader])
                                  VALUES ('{today.ToString("yyyyMMdd")}','{serial}','{userID}')";
                        MSSQL.ExecSqlCmd(sql2, GlobalVar.loginSet.warrantassistant45);

                        int index = Convert.ToInt32(serial.Substring(serial.Length - 2, 2));
                        string d = serial.Substring(0, 8);
                        if(d == DateTime.Today.ToString("yyyyMMdd"))
                            DeletedSerialNum.Add(index);
                    }
                    
                }
                MSSQL.ExecSqlCmd($"DELETE FROM [ApplyTempList] WHERE UserID='{userID}'", conn);
               
            }
            LoadData();
            SetButton();
        }

        private void ButtonCancel_Click(object sender, EventArgs e) {
            isEdit = false;
            LoadData();
            SetButton();
        }
        private void ButtonUpdatePrice_Click(object sender, EventArgs e)
        {
            MSSQL.ExecSqlCmd("DELETE FROM [WarrantPrices]", GlobalVar.loginSet.warrantassistant45);
            MSSQL.ExecSqlCmd(@"INSERT INTO WarrantAssistant.dbo.WarrantPrices 
                                SELECT DISTINCT CASE WHEN ([CommodityId]='1000') THEN 'IX0001' WHEN ([CommodityId]='2300') THEN 'IX0027' WHEN ([CommodityId]='2800') THEN 'IX0039' ELSE [CommodityId] END
                                            ,CASE WHEN ISNULL([MatchPrice],0) = 0 THEN isnull([LastPrice],0) ELSE [MatchPrice] END
                                             ,[tradedate]
                                             ,isnull([BuyPriceBest1],0)
                                             ,isnull([SellPriceBest1],0)
                                             ,tradedate
                               FROM [10.60.0.37].[TsQuote].[dbo].[vwprice2] ", GlobalVar.loginSet.warrantassistant45);
            MessageBox.Show("股價更新完成!");
        }

        private void UltraGrid1_InitializeRow(object sender, InitializeRowEventArgs e) {
            
            
            string sqlSummary = $"SELECT [UnderlyingID], [TraderID], [Issuable], [PutIssuable] FROM [WarrantAssistant].[dbo].[WarrantUnderlyingSummary]";
            DataTable dvSummary = MSSQL.ExecSqlQry(sqlSummary, GlobalVar.loginSet.warrantassistant45);
            try
            {
                string cp = e.Row.Cells["CP"].Value == DBNull.Value ? "" : e.Row.Cells["CP"].Value.ToString();
                string underlyingID = e.Row.Cells["標的代號"].Value == DBNull.Value ? "" : e.Row.Cells["標的代號"].Value.ToString();
                string underlyingName = e.Row.Cells["標的名稱"].Value == DBNull.Value ? "" : e.Row.Cells["標的名稱"].Value.ToString();
                double price = e.Row.Cells["發行價格"].Value == DBNull.Value ? 0.0 : Convert.ToDouble(e.Row.Cells["發行價格"].Value);
                double price_ = e.Row.Cells["發行價格*"].Value == DBNull.Value ? 0.0 : Convert.ToDouble(e.Row.Cells["發行價格*"].Value);
                double vol_ = e.Row.Cells["IV*"].Value == DBNull.Value ? 0.0 : Convert.ToDouble(e.Row.Cells["IV*"].Value);
                double vol = e.Row.Cells["IV"].Value == DBNull.Value ? 0.0 : Convert.ToDouble(e.Row.Cells["IV"].Value);
                double lowerLimit = e.Row.Cells["跌停價*"].Value == DBNull.Value ? 0.0 : Convert.ToDouble(e.Row.Cells["跌停價*"].Value);
                double strike = e.Row.Cells["履約價"].Value == DBNull.Value ? 0.0 : Convert.ToDouble(e.Row.Cells["履約價"].Value);
                double spot = e.Row.Cells["股價"].Value == DBNull.Value ? 0.0 : Convert.ToDouble(e.Row.Cells["股價"].Value);
                double shares = e.Row.Cells["張數"].Value == DBNull.Value ? 10000 : Convert.ToDouble(e.Row.Cells["張數"].Value);
                double cr = e.Row.Cells["行使比例"].Value == DBNull.Value ? 0 : Convert.ToDouble(e.Row.Cells["行使比例"].Value);
                double equivalent5000 = e.Row.Cells["約當張數(5000張)"].Value == DBNull.Value ? 0 : Convert.ToDouble(e.Row.Cells["約當張數(5000張)"].Value);
                double canissue = e.Row.Cells["今日額度"].Value == DBNull.Value ? 0 : Convert.ToDouble(e.Row.Cells["今日額度"].Value);
                double resetR = e.Row.Cells["重設比"].Value == DBNull.Value ? 0 : Convert.ToDouble(e.Row.Cells["重設比"].Value);
                string type = e.Row.Cells["類型"].Value.ToString();
                int t = e.Row.Cells["期間(月)"].Value == DBNull.Value ? 0 : Convert.ToInt32(e.Row.Cells["期間(月)"].Value.ToString());
                string traderID = "NA";
                string issuable = "Y";
                string putIssuable = "Y";
                string toolTip1 = "非本季標的";
                string toolTip2 = "發行檢查=N";
                string toolTip3 = "非此使用者所屬標的";
                string toolTip4 = "此檔Put須告知主管";
                string toolTip5 = $@"高風險且為市值前30大標的 DeltaOne市值 > {(int)(ISTOP30MaxIssue / 100000)}億";
                string toolTip6 = $@"高風險標的 DeltaOne市值 > {(int)(NonTOP30MaxIssue / 100000)}億";
                string toolTip7 = $@"自家Call/Put DeltaOne比例<{NonSpecialCallPutRatio}";
                string toolTip7_2 = $@"為風險標的，自家Call/Put DeltaOne比例<{SpecialCallPutRatio}";
                string toolTip8 = $@"自家/全市場 Put DeltaOne比例>{SpecialKGIALLPutRatio}";
                string toolTip9 = "自家Put DeltaOne>元大Put DeltaOne";
                string toolTip10 = "發行10000張額度不夠";

                DataRow[] dvSummarySelect = dvSummary.Select($@"UnderlyingID = '{underlyingID}'");
                if (dvSummarySelect.Length > 0)
                {
                    traderID = dvSummarySelect[0][1].ToString().PadLeft(7, '0');
                    issuable = dvSummarySelect[0][2].ToString();
                    putIssuable = dvSummarySelect[0][3].ToString();
                }
                if (underlyingID != "")
                {
                    //判斷是否可發，若沒有代入標的名稱，代表非本季標的
                    if (underlyingName == "" & !isEdit)
                    {
                        e.Row.ToolTipText = toolTip1;
                        e.Row.Appearance.ForeColor = Color.Red;
                    }
                    else
                    {
                        if (cp == "C")
                        {
                            DataRow[] volRatio = dtVolRatio.Select($@"UID = '{underlyingID}' AND WClass = 'c'");
                            
                            if (volRatio.Length > 0)
                            {
                                double thetaAmt = Convert.ToDouble(volRatio[0][2].ToString());
                                double gamma = Convert.ToDouble(volRatio[0][3].ToString());
                                double cnt = Convert.ToDouble(volRatio[0][4].ToString());
                                double hv20 = Convert.ToDouble(volRatio[0][5].ToString());
                                double hv60 = Convert.ToDouble(volRatio[0][6].ToString());
                                double r20 = 0;
                                double r60 = 0;
                                double avgTheta = 0;
                                
                                if (thetaAmt > 0 && gamma > 0)
                                {
                                    r20 = Math.Round((Math.Sqrt(thetaAmt * 200 / gamma) * 16) / hv20, 2);
                                    r60 = Math.Round((Math.Sqrt(thetaAmt * 200 / gamma) * 16) / hv60, 2);
                                }
                                if(cnt > 0)
                                {
                                    avgTheta = Math.Round(thetaAmt * 1000 / cnt, 0);
                                }
                                
                                e.Row.Cells["平均Theta"].Value = avgTheta;
                                e.Row.Cells["HV20Ratio"].Value = r20;
                                e.Row.Cells["HV60Ratio"].Value = r60;
                                if (vol < hv20 * 0.7 || vol < hv60 * 0.7)
                                {
                                    //e.Row.Cells["IV"].Appearance.ForeColor = Color.Red;
                                    e.Row.Cells["IV"].ToolTipText = "Vol低於HV的0.7倍";
                                    e.Row.Cells["IV"].Appearance.BackColor = Color.Coral;
                                }
                                if (vol >= 110 || vol > hv20 * 2)
                                {
                                    //e.Row.Cells["IV"].Appearance.ForeColor = Color.Red;
                                    e.Row.Cells["IV"].ToolTipText = "提醒VOL>=110 OR VOL >HV20的2倍，可能KEY錯";
                                    e.Row.Cells["IV"].Appearance.BackColor = Color.Coral;
                                }

                            }
                            DataRow[] dtPL20DSelect = dtPL20D.Select($@"UID = '{underlyingID}' AND OptionType = 'c'");
                            if(dtPL20DSelect.Length > 0)
                            {
                                double PL20 = Convert.ToDouble(dtPL20DSelect[0][2].ToString());
                                e.Row.Cells["PL20日"].Value = PL20;
                                if(PL20 < 0)
                                    e.Row.Cells["PL20日"].Appearance.ForeColor = Color.Red;
                            }
                            
                            DataRow[] dtPLYearSelect = dtPLYear.Select($@"UnderlyingId = '{underlyingID}'");
                            if (dtPLYearSelect.Length > 0)
                            {
                                double PLYear = Convert.ToDouble(dtPLYearSelect[0][1].ToString());
                                e.Row.Cells["PL年"].Value = PLYear;
                                if (PLYear < 0)
                                    e.Row.Cells["PL年"].Appearance.ForeColor = Color.Red;
                            }
                            

                        }
                        if (cp == "P")
                        {
                            DataRow[] volRatio = dtVolRatio.Select($@"UID = '{underlyingID}' AND WClass = 'p'");
                            
                            if (volRatio.Length > 0)
                            {
                                double thetaAmt = Convert.ToDouble(volRatio[0][2].ToString());
                                double gamma = Convert.ToDouble(volRatio[0][3].ToString());
                                double cnt = Convert.ToDouble(volRatio[0][4].ToString());
                                double hv20 = Convert.ToDouble(volRatio[0][5].ToString());
                                double hv60 = Convert.ToDouble(volRatio[0][6].ToString());
                                double r20 = 0;
                                double r60 = 0;
                                double avgTheta = 0;
                                if (thetaAmt > 0 && gamma > 0)
                                {
                                    r20 = Math.Round((Math.Sqrt(thetaAmt * 200 / gamma) * 16) / hv20, 2);
                                    r60 = Math.Round((Math.Sqrt(thetaAmt * 200 / gamma) * 16) / hv60, 2);
                                }
                                if (cnt > 0)
                                {
                                    avgTheta = Math.Round(thetaAmt  * 1000/ cnt, 0);
                                }

                                e.Row.Cells["平均Theta"].Value = avgTheta;
                                e.Row.Cells["HV20Ratio"].Value = r20;
                                e.Row.Cells["HV60Ratio"].Value = r60;
                                if (vol < hv20 * 0.7 || vol < hv60 * 0.7)
                                {
                                   // e.Row.Cells["IV"].Appearance.ForeColor = Color.Red;
                                    e.Row.Cells["IV"].ToolTipText = "Vol低於HV的0.7倍";
                                    e.Row.Cells["IV"].Appearance.BackColor = Color.Coral;
                                }
                                if (vol >= 110 || vol > hv20 * 2)
                                {
                                   // e.Row.Cells["IV"].Appearance.ForeColor = Color.Red;
                                    e.Row.Cells["IV"].ToolTipText = "提醒VOL>=110 OR VOL >HV20的2倍，可能KEY錯";
                                    e.Row.Cells["IV"].Appearance.BackColor = Color.Coral;
                                }
                            }
                            DataRow[] dtPL20DSelect = dtPL20D.Select($@"UID = '{underlyingID}' AND OptionType = 'p'");
                            if (dtPL20DSelect.Length > 0)
                            {
                                double PL20 = Convert.ToDouble(dtPL20DSelect[0][2].ToString());
                                e.Row.Cells["PL20日"].Value = PL20;
                                if (PL20 < 0)
                                    e.Row.Cells["PL20日"].Appearance.ForeColor = Color.Red;
                            }
                            
                            DataRow[] dtPLYearSelect = dtPLYear.Select($@"UnderlyingId = '{underlyingID}'");
                            if (dtPLYearSelect.Length > 0)
                            {
                                double PLYear = Convert.ToDouble(dtPLYearSelect[0][1].ToString());
                                e.Row.Cells["PL年"].Value = PLYear;
                                if (PLYear < 0)
                                    e.Row.Cells["PL年"].Appearance.ForeColor = Color.Red;
                            }
                            

                        }
                        if (UidDeltaOne_Temp.ContainsKey(underlyingID))
                        {
                            if (cp == "C")
                            {
                                if (IsSpecial.Contains(underlyingID))//高風險標的才要控管市值
                                {
                                    if (Market30.Contains(underlyingID))//市值前30  DeltaOne*股價<5億
                                    {
                                        if (cr * shares * spot > ISTOP30MaxIssue)
                                        {
                                            e.Row.ToolTipText = toolTip5;
                                            e.Row.Appearance.ForeColor = Color.Red;
                                        }
                                    }
                                    else
                                    {
                                        if (cr * shares * spot > NonTOP30MaxIssue)
                                        {
                                            e.Row.ToolTipText = toolTip6;
                                            e.Row.Appearance.ForeColor = Color.Red;
                                        }
                                    }
                                }
                            }
                            if (cp == "P")
                            {
                                
                                if (IsSpecial.Contains(underlyingID))
                                {
                                    if (Market30.Contains(underlyingID))//市值前30  DeltaOne*股價<5億
                                    {
                                        if (cr * shares * spot > ISTOP30MaxIssue)
                                        {
                                            e.Row.ToolTipText = toolTip5;
                                            e.Row.Appearance.ForeColor = Color.Red;
                                        }
                                    }
                                    else
                                    {
                                        if (cr * shares * spot > NonTOP30MaxIssue)
                                        {
                                            e.Row.ToolTipText = toolTip6;
                                            e.Row.Appearance.ForeColor = Color.Red;
                                        }
                                    }
                                }
                                //考慮發Put的時候不能把今天要發的Call算進來
                                if (IsSpecial.Contains(underlyingID) && ((double)UidDeltaOne_Temp[underlyingID].KgiCallDeltaOne / (double)UidDeltaOne_Temp[underlyingID].KgiPutDeltaOne < SpecialCallPutRatio))
                                {
                                    e.Row.ToolTipText = toolTip7_2;
                                    e.Row.Appearance.ForeColor = Color.Red;
                                }

                                else if ((double)UidDeltaOne_Temp[underlyingID].KgiCallDeltaOne / (double)UidDeltaOne_Temp[underlyingID].KgiPutDeltaOne < NonSpecialCallPutRatio)
                                {
                                    if (!IsIndex.Contains(underlyingID))
                                    {
                                        e.Row.ToolTipText = toolTip7;
                                        e.Row.Appearance.ForeColor = Color.Red;
                                    }
                                }
                                if (IsSpecial.Contains(underlyingID) && UidDeltaOne_Temp[underlyingID].AllPutDeltaOne > 0 && UidDeltaOne_Temp[underlyingID].KgiAllPutRatio > SpecialKGIALLPutRatio)
                                {

                                    //若之前這檔標的沒發過Put可以跳過，可是要考慮今天發超過一檔
                                    if (UidDeltaOne[underlyingID].KgiPutNum > 0 || (UidDeltaOne[underlyingID].KgiPutNum == 0 && UidDeltaOne_Temp[underlyingID].KgiPutNum > 1))
                                    {
                                        e.Row.ToolTipText = toolTip8;
                                        e.Row.Appearance.ForeColor = Color.Red;
                                    }
                                }
                                //特殊標的要Follow元大
                                /*
                                if (IsSpecial.Contains(underlyingID))
                                {
                                    if (UidDeltaOne_Temp[underlyingID].KgiPutDeltaOne > UidDeltaOne_Temp[underlyingID].YuanPutDeltaOne)
                                    {
                                        e.Row.ToolTipText = toolTip9;
                                        e.Row.Appearance.ForeColor = Color.Red;
                                    }
                                }
                                */
                            }
                        }

                        if (issuable == "N")
                        {
                            e.Row.Cells["標的代號"].ToolTipText = toolTip2;
                            e.Row.Cells["標的代號"].Appearance.ForeColor = Color.Red;
                        }

                        if (cp == "P" && putIssuable == "N")
                        {
                            e.Row.Cells["CP"].Appearance.ForeColor = Color.Red;
                            e.Row.Cells["CP"].ToolTipText = toolTip4;
                        }

                        if (traderID != userID)
                        {
                            e.Row.Appearance.BackColor = Color.LightYellow;
                            e.Row.ToolTipText = toolTip3;
                        }
                    }
                    
                    if (price != 0.0 && (price <= 0.6 || price > 3.0))
                        e.Row.Cells["發行價格"].Appearance.ForeColor = Color.Red;
                    else
                        e.Row.Cells["發行價格"].Appearance.ForeColor = Color.Black;
                    
                    //Check for moneyness constraint
                    e.Row.Cells["履約價"].Appearance.ForeColor = Color.Black;
                    if (type != "牛熊證")
                    {
                        
                        if (cp == "C" && strike / spot >= 1.5 || cp == "P" && strike / spot <= 0.5)
                        {
                            e.Row.Cells["履約價"].Appearance.ForeColor = Color.Red;
                            e.Row.Cells["履約價"].ToolTipText = "履約價超過價外50%";
                        }
                        if (cp == "C" && strike / spot <= 0.9 || cp == "P" && strike / spot >= 1.1)
                        {
                            e.Row.Cells["履約價"].Appearance.ForeColor = Color.Red;
                            e.Row.Cells["履約價"].ToolTipText = "履約價超過價內10%";
                        }
                       
                        if (cp == "C")
                        {
                            
                            DataRow[] dtKKSelect = dtKK.Select($@"stkid = '{underlyingID}' AND [strike_now] = {strike} AND [CP] = 'c'");
                            if (dtKKSelect.Length> 0)
                            {
                                e.Row.Cells["履約價"].Appearance.BackColor = Color.Salmon;
                                e.Row.Cells["履約價"].ToolTipText += "有重複履約價";
                            }
                            DataRow[] dtTTSelect = dtTT.Select($@"UID = '{underlyingID}' AND [StrikePrice] = {strike} AND [WClass] = 'c' AND [TtoM] >= {T2TtoM[t]-20} AND [TtoM] <= {T2TtoM[t] + 20}");
                            if (dtTTSelect.Length > 0)
                            {
                                e.Row.Cells["期間(月)"].Appearance.BackColor = Color.Salmon;
                                e.Row.Cells["期間(月)"].ToolTipText += "TtoM有20天內重複權證";
                            }
                        }
                        else
                        {
                            
                            DataRow[] dtKKSelect = dtKK.Select($@"stkid = '{underlyingID}' AND [strike_now] = {strike} AND [CP] = 'p'");
        
                            //if (dtK.Rows.Count > 0)
                            if (dtKKSelect.Length > 0)
                            {
                                e.Row.Cells["履約價"].Appearance.BackColor = Color.Salmon;
                                e.Row.Cells["履約價"].ToolTipText += "有重複履約價";
                            }
                            DataRow[] dtTTSelect = dtTT.Select($@"UID = '{underlyingID}' AND [StrikePrice] = {strike} AND [WClass] = 'p' AND [TtoM] >= {T2TtoM[t] - 20} AND [TtoM] <= {T2TtoM[t] + 20}");
                            if (dtTTSelect.Length > 0)
                            {
                                e.Row.Cells["期間(月)"].Appearance.BackColor = Color.Salmon;
                                e.Row.Cells["期間(月)"].ToolTipText += "TtoM有20天內重複權證";
                            }
                        }
                        
                    }
                    if(type == "重設型")
                    {
                        if (cp == "C")
                        {
                            if (resetR < 85)
                            {
                                e.Row.Cells["重設比"].Appearance.BackColor = Color.Red;
                                e.Row.Cells["重設比"].ToolTipText += "Call重設比不得低於85";
                            }
                        }
                        else
                        {
                            if (resetR > 115)
                            {
                                e.Row.Cells["重設比"].Appearance.BackColor = Color.Red;
                                e.Row.Cells["重設比"].ToolTipText += "Put重設比不得大於115";
                            }
                        }
                    }
                    if (type == "牛熊證")
                    {
                        if (cp == "C")
                        {
                            string last20 = TradeDate.LastNTradeDate(20);
                            /*
                            string sqlK = $@"SELECT  [代號]
                                      FROM [TwCMData].[dbo].[Warrant總表]
                                      WHERE [日期] = '{lastday.ToString("yyyyMMdd")}' AND [標的代號] = '{underlyingID}' AND [最新履約價] = {strike} AND [券商代號] = '9200' AND [一般證/牛熊證] = 'True' AND [認購/認售] = '認購' AND [發行日期] >= '{last20}'";
                            */
                            string sqlK = $@"SELECT wid
                                FROM [HEDGE].[dbo].[WARRANTS] 
								WHERE ISSUECOMNAME = '凱基' AND [stkid] = '{underlyingID}' AND [type] = '重設型牛證認購' AND [issuedate] >= '{last20}' AND [strike_now] = {strike}
								order by [stkid]";
                            DataTable dtK = MSSQL.ExecSqlQry(sqlK, GlobalVar.loginSet.HEDGE);
                            if (dtK.Rows.Count > 0)
                            {
                                e.Row.Cells["履約價"].Appearance.BackColor = Color.Salmon;
                                e.Row.Cells["履約價"].ToolTipText += "有重複履約價";
                            }
                            if(resetR > 100)
                            {
                                e.Row.Cells["重設比"].Appearance.BackColor = Color.Red;
                                e.Row.Cells["重設比"].ToolTipText += "牛證重設比不得大於100";
                            }
                        }
                        else
                        {
                            string last20 = TradeDate.LastNTradeDate(20);
                            /*
                            string sqlK = $@"SELECT  [代號]
                                      FROM [TwCMData].[dbo].[Warrant總表]
                                      WHERE [日期] = '{lastday.ToString("yyyyMMdd")}' AND [標的代號] = '{underlyingID}' AND [最新履約價] = {strike} AND [券商代號] = '9200' AND [一般證/牛熊證] = 'True' AND [認購/認售] = '認售' AND [發行日期] >= '{last20}'";
                            */
                            string sqlK = $@"SELECT wid
                                FROM [HEDGE].[dbo].[WARRANTS] 
								WHERE ISSUECOMNAME = '凱基' AND [stkid] = '{underlyingID}' AND [type] = '重設型熊證認售' AND [issuedate] >= '{last20}' AND [strike_now] = {strike}
								order by [stkid]";
                            DataTable dtK = MSSQL.ExecSqlQry(sqlK, GlobalVar.loginSet.HEDGE);
                            if (dtK.Rows.Count > 0)
                            {
                                e.Row.Cells["履約價"].Appearance.BackColor = Color.Salmon;
                                e.Row.Cells["履約價"].ToolTipText += "有重複履約價";
                            }
                            if (resetR < 100)
                            {
                                e.Row.Cells["重設比"].Appearance.BackColor = Color.Red;
                                e.Row.Cells["重設比"].ToolTipText += "熊證重設比不得小於100";
                            }
                        }
                    }
                    
                    if (price != 0.0 && (price <= lowerLimit))
                    {
                        e.Row.Cells["IV*"].Appearance.ForeColor = Color.Red;
                        e.Row.Cells["跌停價*"].Appearance.ForeColor = Color.Red;
                    }
                    else
                    {
                        e.Row.Cells["IV*"].Appearance.ForeColor = Color.Black;
                        e.Row.Cells["跌停價*"].Appearance.ForeColor = Color.Black;
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
                    
                    if(equivalent5000 > canissue)
                    {
                        e.Row.Cells["約當張數(5000張)"].Appearance.ForeColor = Color.Red;
                        e.Row.Cells["約當張數(5000張)"].ToolTipText = toolTip10;
                        e.Row.Cells["行使比例"].Appearance.ForeColor = Color.Red;
                        e.Row.Cells["行使比例"].ToolTipText = toolTip10;
                    }
                    if(0.5 * equivalent5000 < canissue)
                    {
                        if (canissue < 5000)
                        {
                            double crTemp = canissue / 5000;
                            e.Row.Cells["建議行使比例"].Value = Math.Floor(crTemp * 1000) / 1000;
                        }
                        else
                            e.Row.Cells["建議行使比例"].Value = 1;
                    }
                    if (underlyingID != "")
                    {
                        if (Reduction.ContainsKey(underlyingID))
                        {
                            e.Row.Cells["標的代號"].ToolTipText += $@"/減資股票，買賣開始日為{Reduction[underlyingID].ToString("yyyyMMdd")}";
                            e.Row.Cells["標的代號"].Appearance.ForeColor = Color.Red;
                        }
                    }
                }
            }
            catch (Exception e1)
            {
                MessageBox.Show($@"{e1.Message}");
            }
        }

        private void UltraGrid1_DoubleClickCell(object sender, DoubleClickCellEventArgs e) {
            if (e.Cell.Row.Cells[1].Value == DBNull.Value)
                return;
            string target = (string) e.Cell.Row.Cells[1].Value;
            if (e.Cell.Row.Cells["CP"].Value.ToString() == "C") {
                FrmIssueCheck frmIssueCheck = GlobalUtility.MenuItemClick<FrmIssueCheck>();
                frmIssueCheck.SelectUnderlying(target);
            }
            if (e.Cell.Row.Cells["CP"].Value.ToString() == "P") {
                //FrmIssueCheckPut frmIssueCheckPut = GlobalUtility.MenuItemClick<FrmIssueCheckPut>();
                //frmIssueCheckPut.SelectUnderlying(target);
            }
        }

        private void 刪除ToolStripMenuItem_Click(object sender, EventArgs e) {

            DialogResult result = MessageBox.Show("刪除此檔，標的:" + ultraGrid1.ActiveRow.Cells["標的代號"].Value + " 履約價:" + ultraGrid1.ActiveRow.Cells["履約價"].Value + "，確定?", "刪除資料", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
            if (result == DialogResult.Yes) {
                
                string serial = ultraGrid1.ActiveRow.Cells["編號"].Value.ToString();
                if (serial.Length > 0)
                {
                    if (ultraGrid1.ActiveRow.Cells["確認"].Value.ToString() == "True")
                    {

                        foreach (var key in sqlTogrid.Keys)
                        {

                            string sql_delete2 = $@"INSERT INTO [WarrantAssistant].[dbo].[ApplyTotalRecord] ([UpdateTime], [UpdateType], [TraderID], [SerialNumber]
                                           , [ApplyKind], [DataName] ,[FromValue], [ToValue], [UpdateCount])
                                          VALUES(GETDATE(), 'DELETE', '{userID}', {serial}, '1','{sqlTogrid[key]}','','',9999)";
                            MSSQL.ExecSqlCmd(sql_delete2, GlobalVar.loginSet.warrantassistant45);
                        }

                    }

                    string sql2 = $@"INSERT INTO [WarrantAssistant].[dbo].[TempListDeleteLog] ([DateTime],[SerialNum],[Trader])
                                  VALUES ('{today.ToString("yyyyMMdd")}','{serial}','{userID}')";
                    MSSQL.ExecSqlCmd(sql2, GlobalVar.loginSet.warrantassistant45);

                    int index = Convert.ToInt32(serial.Substring(serial.Length - 2, 2));
                    DeletedSerialNum.Add(index);
                }
                ultraGrid1.ActiveRow.Delete();
                UpdateData();
            }
            LoadData();
        }

        private void UltraGrid1_MouseDown(object sender, MouseEventArgs e) {
            if (e.Button == MouseButtons.Right) {
                contextMenuStrip1.Show();
            }
        }

        private void UltraGrid1_BeforeRowsDeleted(object sender, BeforeRowsDeletedEventArgs e) {
            e.DisplayPromptMsg = false;
        }

        private void UltraGrid1_AfterCellUpdate(object sender, CellEventArgs e) {

            
            try
            {
               
                if (e.Cell.Column.Key == "標的代號")
                {
                    string underlyingID = e.Cell.Row.Cells["標的代號"].Value.ToString();
                    string underlyingName = "";
                    string traderID = "";
                    double underlyingPrice = 0.0;
                    string sqlTemp;
                    double cr = 0;
                    double canissue = 0;
                    if (canIssue.ContainsKey(underlyingID))
                        canissue = canIssue[underlyingID];
                    if ((double)(canissue / 5000) > 1)
                        cr = 1;
                    else
                        cr = Math.Round((double)(canissue / 5000), 3);
                    DataTable dvTemp;
                    //if (char.IsDigit(underlyingID[0]) || underlyingID == "IX0001") { 
                    //if (Regex.IsMatch(underlyingID, @"^\d")) {

                    sqlTemp = @"SELECT a.[UnderlyingName]
	                                      ,IsNull(IsNull(b.MPrice, IsNull(b.BPrice,b.APrice)),0) MPrice
                                          ,a.TraderID TraderID
                                      FROM [WarrantAssistant].[dbo].[WarrantUnderlying] a
                                      LEFT JOIN [WarrantAssistant].[dbo].[WarrantPrices] b ON a.UnderlyingID=b.CommodityID ";
                    sqlTemp += $"WHERE  CAST(UnderlyingID as varbinary(100)) = CAST('{underlyingID}' as varbinary(100))";


                    dvTemp = MSSQL.ExecSqlQry(sqlTemp, GlobalVar.loginSet.warrantassistant45);

                    foreach (DataRow drTemp in dvTemp.Rows)
                    {
                        underlyingName = drTemp["UnderlyingName"].ToString();
                        traderID = drTemp["TraderID"].ToString().PadLeft(7, '0');
                        underlyingPrice = Convert.ToDouble(drTemp["MPrice"]);
                    }
                    e.Cell.Row.Cells["行使比例"].Value = cr;
                    e.Cell.Row.Cells["張數"].Value = 5000;
                    e.Cell.Row.Cells["期間(月)"].Value = 6;
                    e.Cell.Row.Cells["類型"].Value = "一般型";
                    //e.Cell.Row.Cells["今日額度"].Value = cr;
                    e.Cell.Row.Cells["獎勵"].Value = false;
                    e.Cell.Row.Cells["1500W"].Value = false;
                    e.Cell.Row.Cells["標的名稱"].Value = underlyingName;
                    e.Cell.Row.Cells["交易員"].Value = traderID;
                    e.Cell.Row.Cells["股價"].Value = underlyingPrice;
                    e.Cell.Row.Cells["刪除"].Value = false;
                    e.Cell.Row.Cells["覆"].Value = false;


                    if (underlyingID.Length > 4 && underlyingID.Substring(0, 2) != "00")
                    {
                        e.Cell.Row.Cells["利率"].Value = GlobalVar.globalParameter.interestRate_Index * 100;
                    }
                    else
                        e.Cell.Row.Cells["利率"].Value = GlobalVar.globalParameter.interestRate * 100;



                    DataRow[] dt_LongTermSelect = dt_LongTerm.Select($@"UnderlyingID = '{underlyingID}'");

                    if (dt_LongTermSelect.Length > 0)
                    {
                        double hv20Diff = Convert.ToDouble(dt_LongTermSelect[0][1].ToString());
                        double hv20 = Convert.ToDouble(dt_LongTermSelect[0][2].ToString());
                        double hv60 = Convert.ToDouble(dt_LongTermSelect[0][3].ToString());
                        double levelPrice = Convert.ToDouble(dt_LongTermSelect[0][4].ToString());
                        double accRice = Convert.ToDouble(dt_LongTermSelect[0][7].ToString());
                        if (levelPrice > 0) {
                            if ((Math.Round((underlyingPrice - levelPrice) *100 / levelPrice,1) + accRice > 20) || hv20 > hv60 * 1.2)
                            {
                                MessageBox.Show($@"{underlyingID} 要發長天期, HV20:{hv20} HV20兩天差異:{hv20Diff} 累積漲跌幅:{Math.Round((underlyingPrice - levelPrice) * 100 / levelPrice, 1) + accRice} 要發長天期!");
                            }
                        }

                    }

                    if (AvailableShares.ContainsKey(underlyingID))
                        e.Cell.Row.Cells["今日額度"].Value = AvailableShares[underlyingID];
                    if (RewardShares.ContainsKey(underlyingID))
                        e.Cell.Row.Cells["獎勵額度"].Value = RewardShares[underlyingID];
                    // Check Relation
                    sqlTemp = "Select count(1) from [VOLDB].[dbo].[ED_RelationUnderlying]"
                        + $" where RecordDate = (select top 1 RecordDate from [VOLDB].[dbo].[ED_RelationUnderlying]) and CS8010 = '{underlyingID}'";
                    dvTemp = MSSQL.ExecSqlQry(sqlTemp, "Data Source=10.60.0.39;Initial Catalog=VOLDB;User ID=voldbuser;Password=voldbuser");
                    if (dvTemp.Rows[0][0].ToString() != "0")
                    {
                        sqlTemp = "SELECT MAX([IssueVol]), min(IssueVol) FROM[dbo].[WARRANTS] where kgiwrt = '他家' "
                            + $" and stkid = '{underlyingID}' and marketdate <= GETDATE() and lasttradedate >= GETDATE() and IssueVol<> 0";
                        dvTemp = MSSQL.ExecSqlQry(sqlTemp, "Data Source=10.101.10.5;Initial Catalog=WMM3;User ID=hedgeuser;Password=hedgeuser");
                        if (dvTemp.Rows[0][1] != DBNull.Value)
                            MessageBox.Show($"此為關係人標的，波動度需介於 {dvTemp.Rows[0][1]} 與 {dvTemp.Rows[0][0]} 之間，不然雞盒會該該叫。");
                        else
                            MessageBox.Show("此為關係人標的，須注意波動度，不然雞盒會靠邀。");
                    }
                }
            }
            catch(Exception ex1)
            {
                MessageBox.Show("ex1" + ex1.Message);
            }
            try
            {
                if (e.Cell.Column.Key == "重設比")
                {//改完重設比後執行
                    string cpType = e.Cell.Row.Cells["CP"].Value == DBNull.Value ? "C" : e.Cell.Row.Cells["CP"].Value.ToString();
                    double underlyingPrice = e.Cell.Row.Cells["股價"].Value == DBNull.Value ? 0 : Convert.ToDouble(e.Cell.Row.Cells["股價"].Value);
                    double resetR = e.Cell.Row.Cells["重設比"].Value == DBNull.Value ? 0 : Convert.ToDouble(e.Cell.Row.Cells["重設比"].Value) / 100;
                    if (resetR != 0)
                    {
                        if (cpType == "P")
                            e.Cell.Row.Cells["履約價"].Value = Math.Round((Math.Ceiling(underlyingPrice * resetR * 100)) / 100, 2);
                        else
                            e.Cell.Row.Cells["履約價"].Value = Math.Round((Math.Floor(underlyingPrice * resetR * 100)) / 100, 2);
                    }
                }
            }
            catch (Exception ex2)
            {
                MessageBox.Show("ex2" + ex2.Message);
            }
            try
            {
                if (e.Cell.Column.Key == "財務費用")
                {
                    
                    string warrantType = e.Cell.Row.Cells["類型"].Value == DBNull.Value ? "1" : e.Cell.Row.Cells["類型"].Value.ToString();
                    
                    if (warrantType != "2")
                        return;

                    double price = 0.0;
                    double jumpSize = 0.0;

                    string underlyingID = e.Cell.Row.Cells["標的代號"].Value == DBNull.Value ? "" : e.Cell.Row.Cells["標的代號"].Value.ToString();
                    double underlyingPrice = e.Cell.Row.Cells["股價"].Value == DBNull.Value ? 0 : Convert.ToDouble(e.Cell.Row.Cells["股價"].Value);
                    double k = e.Cell.Row.Cells["履約價"].Value == DBNull.Value ? 0 : Convert.ToDouble(e.Cell.Row.Cells["履約價"].Value);
                    double financialR = e.Cell.Row.Cells["財務費用"].Value == DBNull.Value ? 0 : Convert.ToDouble(e.Cell.Row.Cells["財務費用"].Value) / 100;
                    int t = e.Cell.Row.Cells["期間(月)"].Value == DBNull.Value ? 0 : Convert.ToInt32(e.Cell.Row.Cells["期間(月)"].Value);
                    double cr = e.Cell.Row.Cells["行使比例"].Value == DBNull.Value ? 0 : Convert.ToDouble(e.Cell.Row.Cells["行使比例"].Value);
                    double vol = e.Cell.Row.Cells["IV"].Value == DBNull.Value ? 0 : Convert.ToDouble(e.Cell.Row.Cells["IV"].Value) / 100;
                    double adj = e.Cell.Row.Cells["Adj"].Value == DBNull.Value ? 0 : Convert.ToDouble(e.Cell.Row.Cells["Adj"].Value);

                    double resetR = Math.Round(k / underlyingPrice, 2);
                    string cpType = e.Cell.Row.Cells["CP"].Value == DBNull.Value ? "1" : e.Cell.Row.Cells["CP"].Value.ToString();
                    CallPutType cp = cpType == "2" ? CallPutType.Put : CallPutType.Call;

                    if (underlyingPrice != 0.0 && underlyingID != "")
                    {
                        e.Cell.Row.Cells["重設比"].Value = resetR * 100;
                        if (underlyingID.Length > 4 && underlyingID.Substring(0, 2) != "00")
                        {
                            price = Pricing.BullBearWarrantPrice(cp, underlyingPrice + adj, resetR, GlobalVar.globalParameter.interestRate_Index, vol, t, financialR, cr);
                        }
                        else
                        {
                            price = Pricing.BullBearWarrantPrice(cp, underlyingPrice + adj, resetR, GlobalVar.globalParameter.interestRate, vol, t, financialR, cr);
                        }

                        jumpSize = EDLib.Tick.UpTickSize(underlyingID, underlyingPrice + adj);
                    }

                    e.Cell.Row.Cells["發行價格"].Value = Math.Round(price, 2);
                    e.Cell.Row.Cells["Delta"].Value = 1;
                    e.Cell.Row.Cells["跳動價差"].Value = Math.Round(jumpSize, 4);

                    double shares = e.Cell.Row.Cells["張數"].Value == DBNull.Value ? 10000 : Convert.ToDouble(e.Cell.Row.Cells["張數"].Value);
                    double vol_ = vol;
                    double price_ = price;
                    double lowerLimit = 0.0;
                    double totalValue = price_ * shares * 1000;
                    double volLimit = 2 * vol_;
                    while (totalValue < 15000000 && vol_ != 0.0 && price != 0.0 && vol_ < volLimit)
                    {
                        vol_ += 0.01;
                        if (warrantType == "牛熊證") {
                            if (underlyingID.Length > 4 && underlyingID.Substring(0, 2) != "00")
                            {
                                price_ = Pricing.BullBearWarrantPrice(cp, underlyingPrice + adj, resetR, GlobalVar.globalParameter.interestRate_Index, vol_, t, financialR, cr);
                            }
                            else
                            {
                                price_ = Pricing.BullBearWarrantPrice(cp, underlyingPrice + adj, resetR, GlobalVar.globalParameter.interestRate, vol_, t, financialR, cr);
                            }
                        }
                        totalValue = price_ * shares * 1000;
                    }
                    lowerLimit = price_ - (underlyingPrice + adj) * 0.1 * cr;
                    if (!IsTWUid.Contains(underlyingID))
                        lowerLimit = Math.Max(0.01, lowerLimit);
                    else
                        lowerLimit = 0.01;

                    e.Cell.Row.Cells["IV*"].Value = vol_ * 100;
                    e.Cell.Row.Cells["發行價格*"].Value = Math.Round(price_, 2);
                    e.Cell.Row.Cells["跌停價*"].Value = Math.Round(lowerLimit, 2);

                }
            }
            catch (Exception ex3)
            {
                MessageBox.Show("ex3" + ex3.Message);
            }
            try
            {
                if (e.Cell.Column.Key == "履約價" || e.Cell.Column.Key == "期間(月)" || e.Cell.Column.Key == "行使比例" || e.Cell.Column.Key == "IV"
                    || e.Cell.Column.Key == "類型" || e.Cell.Column.Key == "CP" || e.Cell.Column.Key == "張數" || e.Cell.Column.Key == "Adj")
                {
                    
                    double price = 0.0;
                    double delta = 0.0;
                    double theta = 0.0; //joufan
                    double jumpSize = 0.0;
                    double multiplier = 0.0;
                   
                    string underlyingID = e.Cell.Row.Cells["標的代號"].Value == DBNull.Value ? "" : e.Cell.Row.Cells["標的代號"].Value.ToString();
                    double underlyingPrice = e.Cell.Row.Cells["股價"].Value == DBNull.Value ? 0 : Convert.ToDouble(e.Cell.Row.Cells["股價"].Value);
                    double k = e.Cell.Row.Cells["履約價"].Value == DBNull.Value ? 0 : Convert.ToDouble(e.Cell.Row.Cells["履約價"].Value);
                    int t = e.Cell.Row.Cells["期間(月)"].Value == DBNull.Value ? 0 : Convert.ToInt32(e.Cell.Row.Cells["期間(月)"].Value);
                    int t2m = 0;
                    if (t >= 6)
                        t2m = T2TtoM[t];
                    double cr = e.Cell.Row.Cells["行使比例"].Value == DBNull.Value ? 0 : Convert.ToDouble(e.Cell.Row.Cells["行使比例"].Value);
                    double vol = e.Cell.Row.Cells["IV"].Value == DBNull.Value ? 0 : Convert.ToDouble(e.Cell.Row.Cells["IV"].Value) / 100;
                    double resetR = e.Cell.Row.Cells["重設比"].Value == DBNull.Value ? 0 : Convert.ToDouble(e.Cell.Row.Cells["重設比"].Value) / 100;
                    double financialR = e.Cell.Row.Cells["財務費用"].Value == DBNull.Value ? 0 : Convert.ToDouble(e.Cell.Row.Cells["財務費用"].Value) / 100;
                    string warrantType = e.Cell.Row.Cells["類型"].Value == DBNull.Value ? "一般型" : e.Cell.Row.Cells["類型"].Value.ToString();
                    string cpType = e.Cell.Row.Cells["CP"].Value == DBNull.Value ? "C" : e.Cell.Row.Cells["CP"].Value.ToString();
                    double shares = e.Cell.Row.Cells["張數"].Value == DBNull.Value ? 10000 : Convert.ToDouble(e.Cell.Row.Cells["張數"].Value);
                    bool is1500W = e.Cell.Row.Cells["1500W"].Value == DBNull.Value ? false : (bool)e.Cell.Row.Cells["1500W"].Value;
                    double adj = e.Cell.Row.Cells["Adj"].Value == DBNull.Value ? 0 : Convert.ToDouble(e.Cell.Row.Cells["Adj"].Value);
                    if((underlyingID == "IX0001" || underlyingID == "IX0027" || underlyingID == "IX0039")){
                        if (t == 0)
                            t2m = T2TtoM[6];
                        else
                            t2m = T2TtoM[t];
                        string uidTemp = "";
                        if (underlyingID == "IX0001")
                            uidTemp = "TWA00";
                        if (underlyingID == "IX0027")
                            uidTemp = "TWB23";
                        if (underlyingID == "IX0039")
                            uidTemp = "TWB28";
                        string sqlgetADJ = $@"SELECT Top 1 [ADJ], ABS({t2m} - TtoM) absTtoM
                                             FROM [TwData].[dbo].[V_WarrantTrading] WHERE  UID = '{uidTemp}' AND [WClass] IN ('c','p') AND TDate = (SELECT MAX(TDate) FROM [TwData].[dbo].[WarrantFlow]) ORDER BY absTtoM";
                        DataTable dtgetADJ = MSSQL.ExecSqlQry(sqlgetADJ, GlobalVar.loginSet.twData);
                        if(dtgetADJ.Rows.Count > 0)
                        {
                            adj = Convert.ToDouble(dtgetADJ.Rows[0][0].ToString());
                            e.Cell.Row.Cells["ADJ"].Value = adj;
                        }

                    }
                    //if (warrantType != "一般型" && warrantType != "牛熊證" && warrantType != "重設型")
                    if (warrantType != "一般型" && warrantType != "牛熊證" && warrantType != "重設型" && warrantType != "展延型")
                    {
                        if (warrantType == "2")
                            warrantType = "牛熊證";
                        else if (warrantType == "3")
                            warrantType = "重設型";
                        else if (warrantType == "4")
                            warrantType = "展延型";
                        else
                            warrantType = "一般型";
                    }

                    if (cpType != "C" && cpType != "P")
                    {
                        if (cpType == "2")
                            cpType = "P";
                        else
                            cpType = "C";
                    }
                    if(underlyingID!="" && cpType == "C")
                    {
                        
                        DataRow[] volRatio = dtVolRatio.Select($@"UID = '{underlyingID}' AND WClass = 'c'");
                        if (volRatio.Length > 0)
                        {
                            double thetaAmt = Convert.ToDouble(volRatio[0][2].ToString());
                            double gamma = Convert.ToDouble(volRatio[0][3].ToString());
                            double cnt = Convert.ToDouble(volRatio[0][4].ToString());
                            double hv20 = Convert.ToDouble(volRatio[0][5].ToString());
                            double hv60 = Convert.ToDouble(volRatio[0][6].ToString());
                            double r20 = 0;
                            double r60 = 0;
                            double avgTheta = 0;
                            if (thetaAmt > 0 && gamma > 0)
                            {
                                r20 = Math.Round((Math.Sqrt(thetaAmt * 200 / gamma) * 16) / hv20, 2);
                                r60 = Math.Round((Math.Sqrt(thetaAmt * 200 / gamma) * 16) / hv60, 2);
                            }
                            if (cnt > 0)
                            {
                                avgTheta = Math.Round(thetaAmt  * 1000/ cnt, 0);
                            }

                            e.Cell.Row.Cells["平均Theta"].Value = avgTheta;
                            e.Cell.Row.Cells["HV20Ratio"].Value = r20;
                            e.Cell.Row.Cells["HV60Ratio"].Value = r60;
                        }
                        
                    }
                    if (underlyingID != "" && cpType == "P")
                    {
                        
                        DataRow[] volRatio = dtVolRatio.Select($@"UID = '{underlyingID}' AND WClass = 'p'");
                        if (volRatio.Length > 0)
                        {
                            double thetaAmt = Convert.ToDouble(volRatio[0][2].ToString());
                            double gamma = Convert.ToDouble(volRatio[0][3].ToString());
                            double cnt = Convert.ToDouble(volRatio[0][4].ToString());
                            double hv20 = Convert.ToDouble(volRatio[0][5].ToString());
                            double hv60 = Convert.ToDouble(volRatio[0][6].ToString());
                            double r20 = 0;
                            double r60 = 0;
                            double avgTheta = 0;
                            if (thetaAmt > 0 && gamma > 0)
                            {
                                r20 = Math.Round((Math.Sqrt(thetaAmt * 200 / gamma) * 16) / hv20, 2);
                                r60 = Math.Round((Math.Sqrt(thetaAmt * 200 / gamma) * 16) / hv60, 2);
                            }
                            if (cnt > 0)
                            {
                                avgTheta = Math.Round(thetaAmt *1000/ cnt, 0);
                            }

                            e.Cell.Row.Cells["平均Theta"].Value = avgTheta;
                            e.Cell.Row.Cells["HV20Ratio"].Value = r20;
                            e.Cell.Row.Cells["HV60Ratio"].Value = r60;
                        }
                        
                    }
                    CallPutType cp = CallPutType.Call;
                    if (cpType == "P")
                        cp = CallPutType.Put;
                    else
                        cp = CallPutType.Call;

                    if (underlyingPrice != 0.0 && underlyingID != "")
                    {
                        if (warrantType == "牛熊證")
                        {
                            resetR = Math.Round(k / underlyingPrice, 2);
                            e.Cell.Row.Cells["重設比"].Value = resetR * 100;
                            if (underlyingID.Length > 4 && underlyingID.Substring(0, 2) != "00")
                            {
                                price = Pricing.BullBearWarrantPrice(cp, underlyingPrice + adj, resetR, GlobalVar.globalParameter.interestRate_Index, vol, t, financialR, cr);
                            }
                            else
                            {
                                price = Pricing.BullBearWarrantPrice(cp, underlyingPrice + adj, resetR, GlobalVar.globalParameter.interestRate, vol, t, financialR, cr);
                            }
                        }
                        //else if (warrantType == "重設型")
                        else if (warrantType == "重設型" || warrantType == "展延型")
                        {
                            resetR = Math.Round(k / underlyingPrice, 2);
                            e.Cell.Row.Cells["重設比"].Value = resetR * 100;
                            if (underlyingID.Length > 4 && underlyingID.Substring(0, 2) != "00")
                            {
                                price = Pricing.ResetWarrantPrice(cp, underlyingPrice + adj, resetR, GlobalVar.globalParameter.interestRate_Index, vol, t, cr);
                            }
                            else
                            {
                                price = Pricing.ResetWarrantPrice(cp, underlyingPrice + adj, resetR, GlobalVar.globalParameter.interestRate, vol, t, cr);
                            }
                        }
                        else
                        {
                            if (underlyingID.Length > 4 && underlyingID.Substring(0, 2) != "00")
                            {
                                price = Pricing.NormalWarrantPrice(cp, underlyingPrice + adj, k, GlobalVar.globalParameter.interestRate_Index, vol, t, cr);
                            }
                            else
                            {
                                price = Pricing.NormalWarrantPrice(cp, underlyingPrice + adj, k, GlobalVar.globalParameter.interestRate, vol, t, cr);
                            }
                            e.Cell.Row.Cells["重設比"].Value = 0;
                            e.Cell.Row.Cells["界限比"].Value = 0;
                            e.Cell.Row.Cells["財務費用"].Value = 0;
                        }
                        if (warrantType == "牛熊證")
                        {
                            delta = 1.0;
                            theta = -k * financialR * cr / 365.0;
                        }
                        else
                        {
                            if (underlyingID.Length > 4 && underlyingID.Substring(0, 2) != "00")
                            {
                                delta = Pricing.Delta(cp, underlyingPrice + adj, k, GlobalVar.globalParameter.interestRate_Index, vol, (t * 30.0) / GlobalVar.globalParameter.dayPerYear, GlobalVar.globalParameter.interestRate_Index) * cr;
                                theta = Pricing.Theta(cp, underlyingPrice + adj, k, GlobalVar.globalParameter.interestRate_Index, vol, (t * 30.0) / GlobalVar.globalParameter.dayPerYear, GlobalVar.globalParameter.interestRate_Index) * cr;
                            }
                            else
                            {
                                delta = Pricing.Delta(cp, underlyingPrice + adj, k, GlobalVar.globalParameter.interestRate, vol, (t * 30.0) / GlobalVar.globalParameter.dayPerYear, GlobalVar.globalParameter.interestRate) * cr;
                                theta = Pricing.Theta(cp, underlyingPrice + adj, k, GlobalVar.globalParameter.interestRate, vol, (t * 30.0) / GlobalVar.globalParameter.dayPerYear, GlobalVar.globalParameter.interestRate) * cr;
                            }
                        }

                        multiplier = EDLib.Tick.UpTickSize(underlyingID, underlyingPrice + adj);
                        /*
                        if (e.Cell.Row.Cells["CP"].Value != DBNull.Value)
                        {
                            if (SuggestVol.ContainsKey(underlyingID))
                            {
                                if (SuggestVol[underlyingID].ContainsKey(cpType))
                                {
             
                                    e.Cell.Row.Cells["建議Vol"].Value = Math.Round(SuggestVol[underlyingID][cpType], 2);
                                    if (vol <= 0)
                                    {
                                        e.Cell.Row.Cells["HV"].Value = Math.Round(SuggestVol[underlyingID][cpType], 2);
                                        e.Cell.Row.Cells["IV"].Value = Math.Round(SuggestVol[underlyingID][cpType], 2);
                                    }
                                    e.Cell.Row.Cells["下限Vol"].Value = Math.Round(SuggestVol[underlyingID][cpType] * 0.9, 2);
                                    e.Cell.Row.Cells["建議Vol"].ToolTipText = SuggestVolResult[underlyingID][cpType];
                                }
                                else
                                {
                                    
                                    string sql = $@"SELECT [HV_60D_Spread] FROM [WarrantAssistant].[dbo].[RecommendVol]
                                          WHERE [UID] = '{underlyingID}'";

                                    double hv130 = 0;
                                    DataTable dt = MSSQL.ExecSqlQry(sql, GlobalVar.loginSet.warrantassistant45);
                                    foreach (DataRow dr in dt.Rows)
                                    {
                                        hv130 = Convert.ToDouble(dr["HV_60D_Spread"].ToString());
                                    }
                                    e.Cell.Row.Cells["建議Vol"].Value = Math.Round(hv130, 2);
                                    e.Cell.Row.Cells["建議Vol"].Appearance.BackColor = Color.Bisque;
                                    e.Cell.Row.Cells["建議Vol"].ToolTipText = $@"沒有發{cpType}，用HV*1.3";
                                    e.Cell.Row.Cells["下限Vol"].Value = Math.Round(hv130 * 0.9, 2);
                                    e.Cell.Row.Cells["下限Vol"].Appearance.BackColor = Color.Bisque;
                                    e.Cell.Row.Cells["下限Vol"].ToolTipText = $@"沒有發{cpType}，用HV*1.3";
                                }
                            }
                            else
                            {
                                MessageBox.Show($@"{underlyingID} 沒有建議Vol");
                                double hv = 0;
                                e.Cell.Row.Cells["建議Vol"].Value = hv;
                                e.Cell.Row.Cells["建議Vol"].Appearance.BackColor = Color.Bisque;
                                e.Cell.Row.Cells["建議Vol"].ToolTipText = $@"沒有發權證，建議Vol為0";
                                e.Cell.Row.Cells["下限Vol"].Value = hv;
                                e.Cell.Row.Cells["下限Vol"].Appearance.BackColor = Color.Bisque;
                                e.Cell.Row.Cells["下限Vol"].ToolTipText = $@"沒有發權證，建議Vol為0";
                            }
                        }
                        */
                    }
                    
                    jumpSize = delta * multiplier;

                    e.Cell.Row.Cells["發行價格"].Value = Math.Round(price, 2);
                    e.Cell.Row.Cells["Delta"].Value = Math.Round(delta, 4);
                    e.Cell.Row.Cells["Theta"].Value = Math.Round(theta, 4); //joufan
                    e.Cell.Row.Cells["跳動價差"].Value = Math.Round(jumpSize, 4);
                   
                    double vol_ = vol;
                    double price_ = price;
                    double lowerLimit = 0.0;
                    double totalValue = price_ * shares * 1000;
                    double volLimit = 2 * vol_;
                    while (totalValue < 15000000 && vol_ != 0.0 && price != 0.0 && vol_ < volLimit)
                    {
                        vol_ += 0.01;
                        if (underlyingID.Length > 4 && underlyingID.Substring(0, 2) != "00")
                        {
                            if (warrantType == "牛熊證")
                                price_ = Pricing.BullBearWarrantPrice(cp, underlyingPrice + adj, resetR, GlobalVar.globalParameter.interestRate_Index, vol_, t, financialR, cr);
                            //else if (warrantType == "重設型")
                            else if (warrantType == "重設型" || warrantType == "展延型")
                                price_ = Pricing.ResetWarrantPrice(cp, underlyingPrice + adj, resetR, GlobalVar.globalParameter.interestRate_Index, vol_, t, cr);
                            else
                                price_ = Pricing.NormalWarrantPrice(cp, underlyingPrice + adj, k, GlobalVar.globalParameter.interestRate_Index, vol_, t, cr);
                        }
                        else
                        {
                            if (underlyingID.Length > 4 && underlyingID.Substring(0, 2) != "00")
                            {
                                if (warrantType == "牛熊證")
                                    price_ = Pricing.BullBearWarrantPrice(cp, underlyingPrice + adj, resetR, GlobalVar.globalParameter.interestRate_Index, vol_, t, financialR, cr);
                                //else if (warrantType == "重設型")
                                else if (warrantType == "重設型" || warrantType == "展延型")
                                    price_ = Pricing.ResetWarrantPrice(cp, underlyingPrice + adj, resetR, GlobalVar.globalParameter.interestRate_Index, vol_, t, cr);
                                else
                                    price_ = Pricing.NormalWarrantPrice(cp, underlyingPrice + adj, k, GlobalVar.globalParameter.interestRate_Index, vol_, t, cr);
                            }
                            else
                            {
                                if (warrantType == "牛熊證")
                                    price_ = Pricing.BullBearWarrantPrice(cp, underlyingPrice + adj, resetR, GlobalVar.globalParameter.interestRate, vol_, t, financialR, cr);
                                //else if (warrantType == "重設型")
                                else if (warrantType == "重設型" || warrantType == "展延型")
                                    price_ = Pricing.ResetWarrantPrice(cp, underlyingPrice + adj, resetR, GlobalVar.globalParameter.interestRate, vol_, t, cr);
                                else
                                    price_ = Pricing.NormalWarrantPrice(cp, underlyingPrice + adj, k, GlobalVar.globalParameter.interestRate, vol_, t, cr);
                            }
                        }
                        totalValue = price_ * shares * 1000;
                    }
                    lowerLimit = price_ - (underlyingPrice + adj) * 0.1 * cr;
                    lowerLimit = Math.Max(0.01, lowerLimit);
                    
                    e.Cell.Row.Cells["IV*"].Value = vol_ * 100;
                    e.Cell.Row.Cells["發行價格*"].Value = Math.Round(price_, 2);
                    e.Cell.Row.Cells["跌停價*"].Value = Math.Round(lowerLimit, 2);
                }
            }
            catch (Exception ex4)
            {
                MessageBox.Show("ex4" + ex4.Message);
            }
           
        }
        private void ToolStripButton1_Click(object sender, EventArgs e) {
            //讓交易員可以在任何時間更新權證，不過更改要留紀錄
            if ((GlobalVar.globalParameter.userGroup == "FE")) {
                OfficiallyApply();
                //LoadData();
            }
            else
            {
                
                if (DateTime.Now.TimeOfDay.TotalMinutes < 485)//8:05前會有WarrantBasic Update
                    MessageBox.Show("權證資料更新中，請於08:05後再送出");
                else if (DateTime.Now.TimeOfDay.TotalMinutes > 1200)
                //else if (DateTime.Now.TimeOfDay.TotalMinutes > 750)//12:30
                    MessageBox.Show("超過交易所申報時間，欲改條件請洽管理組");
                /*
                else if (DateTime.Now.TimeOfDay.TotalMinutes > 570) {
                    DialogResult result = MessageBox.Show("超過約定的9:30了，已經告知組長及管理組?", "逾時申請", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
                    if (result == DialogResult.Yes) {
                        OfficiallyApply();
                        LoadData();
                        GlobalUtility.LogInfo("Announce", GlobalVar.globalParameter.userID + " 逾時申請" + applyCount + "檔權證發行");

                    } else
                        LoadData();
                } */
                else {
                    OfficiallyApply();
                    //LoadData();
                }
            }
            
        }
        
        private void UltraGrid1_CellChange(object sender, CellEventArgs e) {
            if (e.Cell.Column.Key == "確認" || e.Cell.Column.Key == "1500W" || e.Cell.Column.Key == "獎勵")
            {
                ultraGrid1.PerformAction(Infragistics.Win.UltraWinGrid.UltraGridAction.ExitEditMode);
            }
        }

        private void UltraGrid1_DoubleClickHeader(object sender, DoubleClickHeaderEventArgs e) {
            if (e.Header.Column.Key == "確認") {
                foreach (Infragistics.Win.UltraWinGrid.UltraGridRow r in ultraGrid1.Rows) {
                    r.Cells["確認"].Value = true;
                }
                UpdateData();
                LoadData();
            }
        }
        private void FrmApply_FormClosed(object sender, FormClosedEventArgs e) {
            //UpdateData();
        }
        private void FrmApply_FormClosing(object sender, FormClosingEventArgs e)
        {
            //UpdateData();
            if(thread1!=null && thread1.IsAlive)
                thread1.Abort();
            //MessageBox.Show($"stop");
        }

        private void UltraGrid1_AfterRowInsert(object sender, RowEventArgs e) {
            //UpdateData();
        }

        private void ToolStripButton2_Click(object sender, EventArgs e) {
            LoadData();
        }
        //存檔
        private void ToolStripButton3_Click(object sender, EventArgs e) {
            UpdateData();
            LoadData();
        }

        private void buttonUpload_Click(object sender, EventArgs e)
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
                    sheet = book.Sheets["Summary"];
                    range = sheet.UsedRange;
                    int rs = range.Rows.Count;
                    MSSQL.ExecSqlCmd($"DELETE FROM [ApplyTempList] WHERE UserID='{userID}'", conn);
                    string sql = @"INSERT INTO [ApplyTempList] (SerialNum, UnderlyingID, K, T, R, HV, IV, IssueNum, ResetR, BarrierR, FinancialR, Type, CP, UseReward, ConfirmChecked, Apply1500W, UserID, MDate, TempName, TempType, TraderID, IVNew, Adj, 說明, [Delete]) "
                    + "VALUES(@SerialNum, @UnderlyingID, @K, @T, @R, @HV, @IV, @IssueNum, @ResetR, @BarrierR, @FinancialR, @Type, @CP, @UseReward, @ConfirmChecked, @Apply1500W, @UserID, @MDate, @TempName ,@TempType, @TraderID, @IVNew, @Adj, @說明, @刪除)";
                    List<SqlParameter> ps = new List<SqlParameter> {
                    new SqlParameter("@SerialNum", SqlDbType.VarChar),
                    new SqlParameter("@UnderlyingID", SqlDbType.VarChar),
                    new SqlParameter("@K", SqlDbType.Float),
                    new SqlParameter("@T", SqlDbType.Int),
                    new SqlParameter("@R", SqlDbType.Float),
                    new SqlParameter("@HV", SqlDbType.Float),
                    new SqlParameter("@IV", SqlDbType.Float),
                    new SqlParameter("@IssueNum", SqlDbType.Float),
                    new SqlParameter("@ResetR", SqlDbType.Float),
                    new SqlParameter("@BarrierR", SqlDbType.Float),
                    new SqlParameter("@FinancialR", SqlDbType.Float),
                    new SqlParameter("@Type", SqlDbType.VarChar),
                    new SqlParameter("@CP", SqlDbType.VarChar),
                    new SqlParameter("@UseReward", SqlDbType.VarChar),
                    new SqlParameter("@ConfirmChecked", SqlDbType.VarChar),
                    new SqlParameter("@Apply1500W", SqlDbType.VarChar),
                    new SqlParameter("@UserID", SqlDbType.VarChar),
                    new SqlParameter("@MDate", SqlDbType.DateTime),
                    new SqlParameter("@TempName", SqlDbType.VarChar),
                    new SqlParameter("@TempType", SqlDbType.VarChar),
                    new SqlParameter("@TraderID", SqlDbType.VarChar),
                    new SqlParameter("@IVNew", SqlDbType.Float),
                    new SqlParameter("@Adj", SqlDbType.Float),
                    new SqlParameter("@說明", SqlDbType.VarChar),
                    new SqlParameter("@刪除", SqlDbType.VarChar)};

                    SQLCommandHelper h = new SQLCommandHelper(GlobalVar.loginSet.warrantassistant45, sql, ps);

                    int i = 1;
                    applyCount = 0;
                    int havingReset = 0;
                    
                    foreach (Infragistics.Win.UltraWinGrid.UltraGridRow r in ultraGrid1.Rows)
                    {
                        string underlyingID = r.Cells["標的代號"].Value.ToString();

                        if (underlyingID != "")
                        {
#if deletelog
                            while (true)
                            {
                                if (!DeletedSerialNum.Contains(i))
                                    break;
                                i++;
                            }
#endif
                            string serialNumber = DateTime.Today.ToString("yyyyMMdd") + userID + "01" + i.ToString("0#");
                            string underlyingName = r.Cells["標的名稱"].Value.ToString();
                            double k = r.Cells["履約價"].Value == DBNull.Value ? 0 : Convert.ToDouble(r.Cells["履約價"].Value);
                            int t = r.Cells["期間(月)"].Value == DBNull.Value ? 6 : Convert.ToInt32(r.Cells["期間(月)"].Value);
                            double cr = r.Cells["行使比例"].Value == DBNull.Value ? 0 : Convert.ToDouble(r.Cells["行使比例"].Value);
                            double hv = r.Cells["HV"].Value == DBNull.Value ? 0 : Convert.ToDouble(r.Cells["HV"].Value);
                            double iv = r.Cells["IV"].Value == DBNull.Value ? 0 : Convert.ToDouble(r.Cells["IV"].Value);
                            double issueNum = r.Cells["張數"].Value == DBNull.Value ? 10000 : Convert.ToDouble(r.Cells["張數"].Value);
                            double resetR = r.Cells["重設比"].Value == DBNull.Value ? 0 : Convert.ToDouble(r.Cells["重設比"].Value);
                            double barrierR = r.Cells["界限比"].Value == DBNull.Value ? 0 : Convert.ToDouble(r.Cells["界限比"].Value);
                            double financialR = r.Cells["財務費用"].Value == DBNull.Value ? 0 : Convert.ToDouble(r.Cells["財務費用"].Value);
                            double adj = r.Cells["Adj"].Value == DBNull.Value ? 0 : Convert.ToDouble(r.Cells["Adj"].Value);
                            string type = r.Cells["類型"].Value.ToString();
                            double underlyingPrice = r.Cells["股價"].Value == DBNull.Value ? 0 : Convert.ToDouble(r.Cells["股價"].Value);
                            //if (type != "一般型" && type != "牛熊證" && type != "重設型") {
                            if (type != "一般型" && type != "牛熊證" && type != "重設型" && type != "展延型")
                            {
                                if (type == "2")
                                    type = "牛熊證";
                                else if (type == "3")
                                    type = "重設型";
                                else if (type == "4")
                                    type = "展延型";
                                else
                                    type = "一般型";
                            }

                            string cp = r.Cells["CP"].Value.ToString();
                            if (cp != "C" && cp != "P")
                            {
                                if (cp == "2")
                                    cp = "P";
                                else
                                    cp = "C";
                            }
                            //if(type == "重設型")
                            if (type == "重設型" || type == "展延型")
                            {
                                havingReset = 1;
                                if (cp == "P")
                                    k = Math.Round(underlyingPrice * 0.5, 2);
                                else
                                    k = Math.Round(underlyingPrice * 1.5, 2);
                            }
                            bool isReward = r.Cells["獎勵"].Value == DBNull.Value ? false : Convert.ToBoolean(r.Cells["獎勵"].Value);
                            string useReward = "N";
                            if (isReward)
                                useReward = "Y";

                            bool confirmed = r.Cells["確認"].Value == DBNull.Value ? false : Convert.ToBoolean(r.Cells["確認"].Value);
                            string confirmChecked = "N";
                            if (confirmed)
                            {
                                confirmChecked = "Y";
                                applyCount++;
                            }

                            bool deleted = r.Cells["刪除"].Value == DBNull.Value ? false : Convert.ToBoolean(r.Cells["刪除"].Value);
                            string deletedStr = "N";
                            if (deleted)
                            {
                                deletedStr = "Y";
                            }
                            bool apply1500Wbool = r.Cells["1500W"].Value == DBNull.Value ? false : Convert.ToBoolean(r.Cells["1500W"].Value);
                            string apply1500W = "N";
                            if (apply1500Wbool)
                                apply1500W = "Y";

                            DateTime expiryDate = GlobalVar.globalParameter.nextTradeDate3.AddMonths(t);


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

                            string warrantType = "";
                            string tempType = "";

                            if (type == "牛熊證")
                            {
                                if (cp == "P")
                                {
                                    warrantType = "熊";
                                    tempType = "4";
                                }
                                else
                                {
                                    warrantType = "牛";
                                    tempType = "3";
                                }
                            }
                            else
                            {
                                if (cp == "P")
                                {
                                    warrantType = "售";
                                    tempType = "2";
                                }
                                else
                                {
                                    warrantType = "購";
                                    tempType = "1";
                                }
                            }

                            string tempName = underlyingName + "凱基" + expiryYear + expiryMonth + warrantType;

                            string traderID = r.Cells["交易員"].Value == DBNull.Value ? userID : r.Cells["交易員"].Value.ToString();

                            double ivNew = r.Cells["IV*"].Value == DBNull.Value ? 0.0 : (double)r.Cells["IV*"].Value;
                            string ex = r.Cells["說明"].Value == DBNull.Value ? "" : r.Cells["說明"].Value.ToString();
                            h.SetParameterValue("@SerialNum", serialNumber);
                            h.SetParameterValue("@UnderlyingID", underlyingID);
                            h.SetParameterValue("@K", k);
                            h.SetParameterValue("@T", t);
                            h.SetParameterValue("@R", cr);
                            h.SetParameterValue("@HV", hv);
                            h.SetParameterValue("@IV", iv);
                            h.SetParameterValue("@IssueNum", issueNum);
                            h.SetParameterValue("@ResetR", resetR);
                            h.SetParameterValue("@BarrierR", barrierR);
                            h.SetParameterValue("@FinancialR", financialR);
                            h.SetParameterValue("@Type", type);
                            h.SetParameterValue("@CP", cp);
                            h.SetParameterValue("@UseReward", useReward);
                            h.SetParameterValue("@ConfirmChecked", confirmChecked);
                            h.SetParameterValue("@Apply1500W", apply1500W);
                            h.SetParameterValue("@UserID", userID);
                            h.SetParameterValue("@MDate", DateTime.Now);
                            h.SetParameterValue("@TempName", tempName);
                            h.SetParameterValue("@TempType", tempType);
                            h.SetParameterValue("@TraderID", traderID);
                            h.SetParameterValue("@IVNew", ivNew);
                            h.SetParameterValue("@Adj", adj);
                            h.SetParameterValue("@說明", ex);
                            h.SetParameterValue("@刪除", deletedStr);
                            h.ExecuteCommand();
                            i++;
                        }
                    }
   
                    for (int j = 2; j <= rs; j++)
                    {
                        
                        string underlyingID = Convert.ToString(sheet.get_Range("A" + j.ToString(), "A" + j.ToString()).Value);
                 
                        string serialNumber = DateTime.Today.ToString("yyyyMMdd") + userID + "01" + i.ToString("0#");
                        //string underlyingName = r.Cells["標的名稱"].Value.ToString();
                        double k = Convert.ToDouble(sheet.get_Range("B" + j.ToString(), "B" + j.ToString()).Value);
                        double cr = Convert.ToDouble(sheet.get_Range("C" + j.ToString(), "C" + j.ToString()).Value);
                        double hv = Convert.ToDouble(sheet.get_Range("D" + j.ToString(), "D" + j.ToString()).Value);
                        double iv = Convert.ToDouble(sheet.get_Range("D" + j.ToString(), "D" + j.ToString()).Value);
                        int t = Convert.ToInt32(sheet.get_Range("E" + j.ToString(), "E" + j.ToString()).Value);
                        
                       
                        double issueNum = Convert.ToDouble(sheet.get_Range("F" + j.ToString(), "F" + j.ToString()).Value);
                        double resetR = 0;
                        double barrierR = 0;
                        double financialR = 0;
                        double adj = 0;
                        string type = "一般型";
                        string cp = Convert.ToString(sheet.get_Range("G" + j.ToString(), "G" + j.ToString()).Value);
                        if (cp == "c" || cp == "C")
                            cp = "C";
                        else
                            cp = "P";
                        string ex = Convert.ToString(sheet.get_Range("H" + j.ToString(), "H" + j.ToString()).Value);

                        string underlyingName = "";
                        string traderID = "";
                        double underlyingPrice = 0.0;
                        string useReward = "N";
                        string confirmChecked = "N";
                        string deletedStr = "N";
                        string sqlTemp;
                        DataTable dvTemp;
             
                        sqlTemp = @"SELECT a.[UnderlyingName]
	                                      ,IsNull(IsNull(b.MPrice, IsNull(b.BPrice,b.APrice)),0) MPrice
                                          ,a.TraderID TraderID
                                      FROM [WarrantAssistant].[dbo].[WarrantUnderlying] a
                                      LEFT JOIN [WarrantAssistant].[dbo].[WarrantPrices] b ON a.UnderlyingID=b.CommodityID ";
                        sqlTemp += $"WHERE  CAST(UnderlyingID as varbinary(100)) = CAST('{underlyingID}' as varbinary(100))";


                        dvTemp = MSSQL.ExecSqlQry(sqlTemp, GlobalVar.loginSet.warrantassistant45);

                        foreach (DataRow drTemp in dvTemp.Rows)
                        {
                            underlyingName = drTemp["UnderlyingName"].ToString();
                            traderID = drTemp["TraderID"].ToString().PadLeft(7, '0');
                            underlyingPrice = Convert.ToDouble(drTemp["MPrice"]);
                        }

                        string apply1500W = "N";
                       

                        DateTime expiryDate = GlobalVar.globalParameter.nextTradeDate3.AddMonths(t);


                        if (expiryDate.Day == GlobalVar.globalParameter.nextTradeDate3.Day)
                            expiryDate = expiryDate.AddDays(-1);
                        sqlTemp = $"SELECT TOP 1 TradeDate from TradeDate WHERE IsTrade='Y' AND TradeDate >= '{expiryDate.ToString("yyyy-MM-dd")}'";
                        
                        DataTable dvTemp2 = MSSQL.ExecSqlQry(sqlTemp, GlobalVar.loginSet.tsquoteSqlConnString);
                        foreach (DataRow drTemp2 in dvTemp2.Rows)
                        {
                            expiryDate = Convert.ToDateTime(drTemp2["TradeDate"]);
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

                        string warrantType = "";
                        string tempType = "";

                        
                            
                        
                        if (cp == "P")
                        {
                            warrantType = "售";
                            tempType = "2";
                        }
                        else
                        {
                            warrantType = "購";
                            tempType = "1";
                        }
                        

                        string tempName = underlyingName + "凱基" + expiryYear + expiryMonth + warrantType;
                        double ivNew = iv;
                        h.SetParameterValue("@SerialNum", serialNumber);
                        h.SetParameterValue("@UnderlyingID", underlyingID);
                        h.SetParameterValue("@K", k);
                        h.SetParameterValue("@T", t);
                        h.SetParameterValue("@R", cr);
                        h.SetParameterValue("@HV", hv);
                        h.SetParameterValue("@IV", iv);
                        h.SetParameterValue("@IssueNum", issueNum);
                        h.SetParameterValue("@ResetR", resetR);
                        h.SetParameterValue("@BarrierR", barrierR);
                        h.SetParameterValue("@FinancialR", financialR);
                        h.SetParameterValue("@Type", type);
                        h.SetParameterValue("@CP", cp);
                        h.SetParameterValue("@UseReward", useReward);
                        h.SetParameterValue("@ConfirmChecked", confirmChecked);
                        h.SetParameterValue("@Apply1500W", apply1500W);
                        h.SetParameterValue("@UserID", userID);
                        h.SetParameterValue("@MDate", DateTime.Now);
                        h.SetParameterValue("@TempName", tempName);
                        h.SetParameterValue("@TempType", tempType);
                        h.SetParameterValue("@TraderID", traderID);
                        h.SetParameterValue("@IVNew", ivNew);
                        h.SetParameterValue("@Adj", adj);
                        h.SetParameterValue("@說明", ex);
                        h.SetParameterValue("@刪除", deletedStr);
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
                    GlobalUtility.LogInfo("Log", GlobalVar.globalParameter.userID + " 編輯/更新" + (i - 1) + "檔發行");
                    if (havingReset == 1)
                        MessageBox.Show("有重設型OR展延型，履約價已調整為(C / P)股價的(150% / 50%) 會以申請當日收盤價為準");
                    LoadData();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

            /*
            OpenFileDialog openFileDialog1 = new OpenFileDialog();
            openFileDialog1.Title = "打開excel文件";
            openFileDialog1.Filter = "Excel Files|*.csv;*.xls;*.xlsx;*.xlsm";
            openFileDialog1.InitialDirectory = @"D:\\";
            openFileDialog1.RestoreDirectory = true;
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                string path = openFileDialog1.FileName;
                try
                {
                    Microsoft.Office.Interop.Excel.Workbook book = null;
                    Microsoft.Office.Interop.Excel.Worksheet sheet = null;
                    Microsoft.Office.Interop.Excel.Range range = null;
                    Microsoft.Office.Interop.Excel.Application app = new Microsoft.Office.Interop.Excel.Application();
                    
                    book = app.Workbooks.Open(path);
                   
                    //sheet從1開始
                    sheet = book.Sheets["Summary"];
                    range = sheet.UsedRange;
                    int rs = range.Rows.Count;
                    int cs = range.Columns.Count;
                    for(int i = 2; i < rs; i++)
                    {
                        try
                        {
                            string uid = Convert.ToString(sheet.get_Range("A" + i.ToString(), "A" + i.ToString()).Value);
                            string cp = Convert.ToString(sheet.get_Range("P" + i.ToString(), "P" + i.ToString()).Value);
                            double k = Convert.ToDouble(sheet.get_Range("Y" + i.ToString(), "Y" + i.ToString()).Value);
                            double vol = Convert.ToDouble(sheet.get_Range("Z" + i.ToString(), "Z" + i.ToString()).Value);
                            double t = Convert.ToDouble(sheet.get_Range("AA" + i.ToString(), "AA" + i.ToString()).Value);
                            double cr = Convert.ToDouble(sheet.get_Range("AB" + i.ToString(), "AB" + i.ToString()).Value);
                            MessageBox.Show($@"{i}  {uid}  {cp}  {k}  {vol}  {t}  {cr}");
                        }
                        catch(Exception ex)
                        {
                            MessageBox.Show(i.ToString());
                        }
                        
                    }
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
                        GC.Collect();
                    }
                    
                }
                catch(Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }
            */
        }

        private void buttonGeneralIssue_Click(object sender, EventArgs e)
        {
            GlobalUtility.MenuItemClick<FrmGeneralIssueV2>();
        }

        private void buttonSaveExcel_Click(object sender, EventArgs e)
        {
            SaveFileDialog sFileDialog = new SaveFileDialog();
            sFileDialog.Title = "匯出Excel";
            sFileDialog.Filter = "EXCEL檔 (*.xlsx)|*.xlsx";
            sFileDialog.InitialDirectory = "D:\\";
            Microsoft.Office.Interop.Excel.Application app = new Microsoft.Office.Interop.Excel.Application();
            if (sFileDialog.ShowDialog() == System.Windows.Forms.DialogResult.OK && sFileDialog.FileName != null)
            {
                Microsoft.Office.Interop.Excel.Workbook workbook = app.Workbooks.Add(1);
                Microsoft.Office.Interop.Excel.Worksheet worksheets = workbook.Sheets[1];
                int colnum = ultraGrid1.DisplayLayout.Bands[0].Columns.Count;
                int rownum = ultraGrid1.Rows.Count;
                int ii = 0;
                for (int i = 0; i < colnum; i++)
                {
                    if (ultraGrid1.DisplayLayout.Bands[0].Columns[i].ToString() == "標的代號" || ultraGrid1.DisplayLayout.Bands[0].Columns[i].ToString() == "履約價" || ultraGrid1.DisplayLayout.Bands[0].Columns[i].ToString() == "期間(月)" || ultraGrid1.DisplayLayout.Bands[0].Columns[i].ToString() == "行使比例" || ultraGrid1.DisplayLayout.Bands[0].Columns[i].ToString() == "IV" || ultraGrid1.DisplayLayout.Bands[0].Columns[i].ToString() == "張數" || ultraGrid1.DisplayLayout.Bands[0].Columns[i].ToString() == "CP" || ultraGrid1.DisplayLayout.Bands[0].Columns[i].ToString() == "說明" || ultraGrid1.DisplayLayout.Bands[0].Columns[i].ToString() == "跳動價差")
                    {
                        worksheets.get_Range($"{(char)(65 + ii) + "1"}", $"{(char)(65 + ii) + "1"}").Value = ultraGrid1.DisplayLayout.Bands[0].Columns[i].ToString();
                        ii++;
                    }
                }
                Microsoft.Office.Interop.Excel.Range range = worksheets.get_Range("A2", $"A{rownum + 1}");
                range.NumberFormat = "@";
                for (int i = 0; i < rownum; i++)
                {
                    int jj = 0;
                    for (int j = 0; j < colnum; j++)
                    {
                        if (ultraGrid1.DisplayLayout.Bands[0].Columns[j].ToString() == "標的代號" || ultraGrid1.DisplayLayout.Bands[0].Columns[j].ToString() == "履約價" || ultraGrid1.DisplayLayout.Bands[0].Columns[j].ToString() == "期間(月)" || ultraGrid1.DisplayLayout.Bands[0].Columns[j].ToString() == "行使比例" || ultraGrid1.DisplayLayout.Bands[0].Columns[j].ToString() == "IV" || ultraGrid1.DisplayLayout.Bands[0].Columns[j].ToString() == "張數" || ultraGrid1.DisplayLayout.Bands[0].Columns[j].ToString() == "CP" || ultraGrid1.DisplayLayout.Bands[0].Columns[j].ToString() == "說明" || ultraGrid1.DisplayLayout.Bands[0].Columns[j].ToString() == "跳動價差")
                        {
                            worksheets.get_Range($"{(char)(65 + jj) + (i + 2).ToString()}", $"{(char)(65 + jj) + (i + 2).ToString()}").Value = ultraGrid1.Rows[i].Cells[j].Value;
                            jj++;
                        }
                        //MessageBox.Show($"{(char)(65 + j) + (i + 2).ToString()} {(char)(65 + j) + (i + 2).ToString()}  { dataGridView1.Rows[i].Cells[j].Value}");
                    }
                }
                workbook.SaveAs(sFileDialog.FileName);
                workbook.Close();
                app.Quit();
                MessageBox.Show("匯出完成");
            }
        }

        
        private int CheckData_SameIssue()
        {

            Dictionary<string, double> Kdiff_C = new Dictionary<string, double>();
            Dictionary<string, double> Kdiff_P = new Dictionary<string, double>();
            int sameCnt = 0;
            string date = EDLib.TradeDate.LastNTradeDate(1).ToString("yyyyMMdd");
            //string date = "20230216";
            string toDate = DateTime.Today.ToString("yyyyMMdd");
            //string toDate = "20230217";
            double diff_T = Convert.ToDouble(toolStripTextBox1.Text);
            double KRatio = Convert.ToDouble(toolStripTextBox2.Text);

            string sql_Kdiff = $@"SELECT  [UID],[WClass],CEILING(AVG([UClosePrice]) *0.6 * {KRatio} / COUNT(*)) AS 價格間距 
                                  FROM [TwData].[dbo].[V_WarrantTrading]
                                  WHERE [TDate] = '{date}' AND [IssuerName] IN ('9800','9200','9100','7790') AND [WClass] IN ('c','p')
                                  AND (((ISNULL(BidAskMed/NULLIF(JumpSize,0),3) < 2.5 AND BidAskMed <> 0) OR BidAskMed = 0.01) OR ((JumpSize > 0 AND JumpSize < 0.01 AND BidAskMed = 0.02)))
                                  AND [Moneyness] >= -0.5 AND [Moneyness] <= 0.1　AND [TtoM] >= 90 AND LEN([UID]) < 5 AND LEFT([UID],2) <> '00'
                                  GROUP BY [UID],[WClass]";
            DataTable dt_Kdiff = MSSQL.ExecSqlQry(sql_Kdiff, GlobalVar.loginSet.twData);
            foreach (Infragistics.Win.UltraWinGrid.UltraGridRow dr in ultraGrid1.Rows)
            {
                string cp = dr.Cells["CP"].Value.ToString();
                string underlyingID = dr.Cells["標的代號"].Value.ToString();
                if(cp == "C")
                {
                    if (!Kdiff_C.ContainsKey(underlyingID))
                    {
                        DataRow[] KdiffSelect = dt_Kdiff.Select($@"UID = '{underlyingID}' AND WClass = 'c'");
                        if(KdiffSelect.Length > 0)
                        {
                            double k = Convert.ToDouble(KdiffSelect[0][2]);
                            Kdiff_C.Add(underlyingID, k);
                        }
                    }
                }
                if (cp == "P")
                {
                    if (!Kdiff_P.ContainsKey(underlyingID))
                    {
                        DataRow[] KdiffSelect = dt_Kdiff.Select($@"UID = '{underlyingID}' AND WClass = 'p'");
                        if (KdiffSelect.Length > 0)
                        {
                            double k = Convert.ToDouble(KdiffSelect[0][2]);
                            Kdiff_P.Add(underlyingID, k);
                        }
                    }
                }
            }

            string sql_checkList = $@"SELECT A.[標的代號] AS UID,A.代號 AS WID,A.名稱 AS WName,A.WClass,A.最新履約價 AS StrikePrice,B.[WarrantTradeDays] AS TtoM FROM
                                        (SELECT  [代號],[名稱],[標的代號],CASE WHEN [名稱] like '%購%' THEN 'c' ELSE 'p' END AS WClass,[最新履約價],[上市日期]
                                          FROM [TwCMData].[dbo].[Warrant總表]　WHERE [日期] = '{date}'　AND [券商名稱] = '凱基' AND ([名稱] LIKE '%購%' OR [名稱] LIKE '%售%')　
                                          AND LEN([標的代號]) < 5 AND LEFT([標的代號],2) <> '00') AS A
                                          LEFT JOIN
                                          (SELECT  SUBSTRING([WarrantKey],1,LEN([WarrantKey]) -5) AS WARRANTKEY,[WId],[WarrantTradeDays]
                                          FROM [10.101.10.5].[WMM3].[dbo].[TtoM]　WHERE [TDate] = '{toDate}') AS B ON A.代號 = B.[WId] AND A.名稱 = B.WARRANTKEY
                                          WHERE B.[WarrantTradeDays] > 0";
            DataTable dt_checkList = MSSQL.ExecSqlQry(sql_checkList, GlobalVar.loginSet.twCMData);
            foreach (Infragistics.Win.UltraWinGrid.UltraGridRow dr in ultraGrid1.Rows)
            {
                //string serial = dr.Cells["編號"].Value.ToString();
                string cp = dr.Cells["CP"].Value.ToString();
                string underlyingID = dr.Cells["標的代號"].Value.ToString();
                int t = dr.Cells["期間(月)"].Value == DBNull.Value ? 0 : Convert.ToInt32(dr.Cells["期間(月)"].Value.ToString());
                double strike = dr.Cells["履約價"].Value == DBNull.Value ? 0.0 : Convert.ToDouble(dr.Cells["履約價"].Value);
                bool confirmed = dr.Cells["確認"].Value == DBNull.Value ? false : Convert.ToBoolean(dr.Cells["確認"].Value);
                if (!confirmed)
                    continue;
                else
                {
                    //MessageBox.Show(underlyingID);
                    if (cp == "C") {
                        //MessageBox.Show($@"UID = '{underlyingID}' AND [StrikePrice] <= {strike + Kdiff_C[underlyingID]} AND [StrikePrice] >= {strike - Kdiff_C[underlyingID]} AND [WClass] = 'c' AND [TtoM] >= {T2TtoM[t] - 20} AND [TtoM] <= {T2TtoM[t] + 20}");
                        DataRow[] checkListSelect = dt_checkList.Select($@"UID = '{underlyingID}' AND [StrikePrice] <= {strike + Kdiff_C[underlyingID]} AND [StrikePrice] >= {strike - Kdiff_C[underlyingID]} AND [WClass] = 'c' AND [TtoM] >= {T2TtoM[t] - 20} AND [TtoM] <= {T2TtoM[t] + 20}");
                        if(checkListSelect.Length > 0)
                        {
                            string wid = checkListSelect[0][1].ToString();
                            dr.Cells["標的代號"].Appearance.BackColor = Color.MediumPurple;
                            dr.Cells["標的代號"].ToolTipText += $@"與{wid}重複";
                            sameCnt++;
                        }
                    }
                    if (cp == "P")
                    {
                        DataRow[] checkListSelect = dt_checkList.Select($@"UID = '{underlyingID}' AND [StrikePrice] <= {strike + Kdiff_P[underlyingID]} AND [StrikePrice] >= {strike - Kdiff_P[underlyingID]} AND [WClass] = 'p' AND [TtoM] >= {T2TtoM[t] - 20} AND [TtoM] <= {T2TtoM[t] + 20}");
                        if (checkListSelect.Length > 0)
                        {
                            string wid = checkListSelect[0][1].ToString();
                            dr.Cells["標的代號"].Appearance.BackColor = Color.MediumPurple;
                            dr.Cells["標的代號"].ToolTipText += $@"與{wid}重複";
                            sameCnt++;
                        }
                    }
                }
            }
            return sameCnt;
           
        }

        private int CheckData_OverlapIssue()
        {

            
            int sameCnt = 0;
            string date = EDLib.TradeDate.LastNTradeDate(1).ToString("yyyyMMdd");
            //string date = "20230216";
            string toDate = DateTime.Today.ToString("yyyyMMdd");
            //string toDate = "20230217";
            //double diff_T = Convert.ToDouble(toolStripTextBox1.Text);
            //double KRatio = Convert.ToDouble(toolStripTextBox2.Text);
            double diff_T_Call = 20;
            double diff_T_Put = 25;
            double K_Ratio_Call = 0.05;
            double K_Ratio_Put = 0.07;
            double volRatio = 0.95;


            Dictionary<string, double> UClosePrice = new Dictionary<string, double>();
            string sql_U = $@"SELECT  [股票代號],[收盤價]　FROM [TwCMData].[dbo].[日收盤表排行]　WHERE [日期] = '{date}' AND [收盤價] IS NOT NULL";
            DataTable dt_U = MSSQL.ExecSqlQry(sql_U, GlobalVar.loginSet.twCMData);
            
            foreach(DataRow dr_U in dt_U.Rows)
            {
                string uid = dr_U["股票代號"].ToString();
                double p = Convert.ToDouble(dr_U["收盤價"].ToString());
                if (!UClosePrice.ContainsKey(uid))
                    UClosePrice.Add(uid, p);
            }
            
            Dictionary<string, double> HV20 = new Dictionary<string, double>();
            Dictionary<string, double> HV60 = new Dictionary<string, double>();
            string sql_HV = $@"SELECT [UID],CASE WHEN [HV20] = 0 THEN 100 ELSE [HV20] END AS HV20, CASE WHEN [HV60] = 0 THEN 100 ELSE [HV60] END AS HV60　FROM [TwData].[dbo].[UnderlyingHitoricalVol] WHERE [TDate] = '{date}'";
            DataTable dt_HV = MSSQL.ExecSqlQry(sql_HV, GlobalVar.loginSet.twData);
            
            foreach (DataRow dr_HV in dt_HV.Rows)
            {
                string uid = dr_HV["UID"].ToString();
                double hv20 = Convert.ToDouble(dr_HV["HV20"].ToString());
                double hv60 = Convert.ToDouble(dr_HV["HV60"].ToString());
                if (!HV20.ContainsKey(uid))
                    HV20.Add(uid, hv20);
                if (!HV60.ContainsKey(uid))
                    HV60.Add(uid, hv60);
            }
            
      

            string sql_checkList = $@"SELECT [WID],[UID],[IssuerName],[WClass],[ListedDays],[StrikePrice],[TtoM],[IV] *100 AS [IV]
                                    FROM [TwData].[dbo].[V_WarrantTrading]　WHERE [TDate] = '{date}' AND [IssuerName] IN ('9800','9200','9100','7790') AND [WClass] IN ('c','p')
                                  AND (((ISNULL(BidAskMed/NULLIF(JumpSize,0),3) < 2.5 AND BidAskMed <> 0) OR BidAskMed = 0.01) OR ((JumpSize > 0 AND JumpSize < 0.01 AND BidAskMed = 0.02)))
                                  AND [Moneyness] >= -0.5 AND [Moneyness] <= 0.1";
            DataTable dt_checkList = MSSQL.ExecSqlQry(sql_checkList, GlobalVar.loginSet.twData);
            foreach (Infragistics.Win.UltraWinGrid.UltraGridRow dr in ultraGrid1.Rows)
            {
                //string serial = dr.Cells["編號"].Value.ToString();
                string cp = dr.Cells["CP"].Value.ToString();
                string underlyingID = dr.Cells["標的代號"].Value.ToString();
                double vol = Convert.ToDouble(dr.Cells["IV"].Value.ToString());
                int t = dr.Cells["期間(月)"].Value == DBNull.Value ? 0 : Convert.ToInt32(dr.Cells["期間(月)"].Value.ToString());
                double strike = dr.Cells["履約價"].Value == DBNull.Value ? 0.0 : Convert.ToDouble(dr.Cells["履約價"].Value);
                bool overlap = dr.Cells["覆"].Value == DBNull.Value ? false : Convert.ToBoolean(dr.Cells["覆"].Value);
                if (!overlap)
                    continue;
                bool confirmed = dr.Cells["確認"].Value == DBNull.Value ? false : Convert.ToBoolean(dr.Cells["確認"].Value);
                if (!confirmed)
                    continue;
                else
                {
                    //MessageBox.Show(underlyingID);
                    if (cp == "C")
                    {
                        //MessageBox.Show($@"UID = '{underlyingID}' AND [StrikePrice] <= {strike + Math.Round(UClosePrice[underlyingID] * K_Ratio_Call, 1)} AND [StrikePrice] >= {strike - Math.Round(UClosePrice[underlyingID] * K_Ratio_Call, 1)} AND [WClass] = 'c' AND [TtoM] >= {T2TtoM[t] - diff_T_Call} AND [TtoM] <= {T2TtoM[t] + diff_T_Call} AND [IssuerName] <> '9200'");
                        DataRow[] checkListSelect = dt_checkList.Select($@"UID = '{underlyingID}' AND [StrikePrice] <= {strike + Math.Round(UClosePrice[underlyingID] * K_Ratio_Call,1)} AND [StrikePrice] >= {strike - Math.Round(UClosePrice[underlyingID] * K_Ratio_Call, 1)} AND [WClass] = 'c' AND [TtoM] >= {T2TtoM[t] - diff_T_Call} AND [TtoM] <= {T2TtoM[t] + diff_T_Call} AND [IssuerName] <> '9200'");
                        
                        if (checkListSelect.Length > 0)
                        {
                            double MinVol = 999;
                            foreach(DataRow dr_check in checkListSelect)
                            {
                                string wid = dr_check["WID"].ToString();
                                double iv = Convert.ToDouble(dr_check["IV"].ToString());
                                if (iv <= MinVol)
                                    MinVol = iv;
                            }
                            DataRow[] checkListSelect_9200 = dt_checkList.Select($@"UID = '{underlyingID}' AND [StrikePrice] <= {strike + Math.Round(UClosePrice[underlyingID] * K_Ratio_Call, 1)} AND [StrikePrice] >= {strike - Math.Round(UClosePrice[underlyingID] * K_Ratio_Call, 1)} AND [WClass] = 'c' AND [TtoM] >= {T2TtoM[t] - diff_T_Call} AND [TtoM] <= {T2TtoM[t] + diff_T_Call} AND [IssuerName] = '9200'");
                            //MessageBox.Show($@"UID = '{underlyingID}' AND [StrikePrice] <= {strike + Math.Round(UClosePrice[underlyingID] * K_Ratio_Call, 1)} AND [StrikePrice] >= {strike - Math.Round(UClosePrice[underlyingID] * K_Ratio_Call, 1)} AND [WClass] = 'c' AND [TtoM] >= {T2TtoM[t] - diff_T_Call} AND [TtoM] <= {T2TtoM[t] + diff_T_Call} AND [IssuerName] = '9200'");
                            if(checkListSelect_9200.Length > 0)
                            {
                                double MinVol_9200 = 999;
                                foreach (DataRow dr_check in checkListSelect_9200)
                                {
                                    string wid = dr_check["WID"].ToString();
                                    double iv = Convert.ToDouble(dr_check["IV"].ToString());
                                    MessageBox.Show($@"{wid} {iv}");
                                    if (iv <= MinVol_9200)
                                        MinVol_9200 = iv;
                                }
                                MessageBox.Show(MinVol_9200.ToString());
                                if(MinVol_9200 <= MinVol * volRatio)
                                {
                                    dr.Cells["覆"].Appearance.BackColor = Color.MediumPurple;
                                    dr.Cells["覆"].ToolTipText += $@"已有覆蓋權證";
                                    sameCnt++;
                                }
                                else
                                {
                                    if (vol > MinVol * 0.95)
                                    {
                                        dr.Cells["覆"].Appearance.BackColor = Color.MediumPurple;
                                        dr.Cells["覆"].ToolTipText += $@"Vol應為{Math.Floor(MinVol * volRatio)}";
                                        sameCnt++;
                                    }
                                    if (Math.Floor(MinVol * volRatio) <= HV20[underlyingID])
                                    {
                                        dr.Cells["覆"].Appearance.BackColor = Color.MediumPurple;
                                        dr.Cells["覆"].ToolTipText += $@"覆蓋率的VOL被HV20穿";
                                    }
                                    if (Math.Floor(MinVol * volRatio) <= HV60[underlyingID])
                                    {
                                        dr.Cells["覆"].Appearance.BackColor = Color.MediumPurple;
                                        dr.Cells["覆"].ToolTipText += $@"覆蓋率的VOL被HV60穿";
                                    }
                                }
                            }
                            else
                            {
                                if(vol > MinVol * 0.95)
                                {
                                    dr.Cells["覆"].Appearance.BackColor = Color.MediumPurple;
                                    dr.Cells["覆"].ToolTipText += $@"Vol應為{Math.Floor(MinVol * volRatio)}";
                                    sameCnt++;
                                }
                                if(Math.Floor(MinVol * volRatio) <= HV20[underlyingID])
                                {
                                    dr.Cells["覆"].Appearance.BackColor = Color.MediumPurple;
                                    dr.Cells["覆"].ToolTipText += $@"覆蓋率的VOL被HV20穿";
                                }
                                if (Math.Floor(MinVol * volRatio) <= HV60[underlyingID])
                                {
                                    dr.Cells["覆"].Appearance.BackColor = Color.MediumPurple;
                                    dr.Cells["覆"].ToolTipText += $@"覆蓋率的VOL被HV60穿";
                                }

                            }
                            
                        }
                        else
                        {
                            dr.Cells["覆"].Appearance.BackColor = Color.MediumPurple;
                            dr.Cells["覆"].ToolTipText += $@"其他三家沒發";
                            sameCnt++;
                        }
                        
                    }
                    
                    if (cp == "P")
                    {
                        //MessageBox.Show($@"UID = '{underlyingID}' AND [StrikePrice] <= {strike + Kdiff_C[underlyingID]} AND [StrikePrice] >= {strike - Kdiff_C[underlyingID]} AND [WClass] = 'c' AND [TtoM] >= {T2TtoM[t] - 20} AND [TtoM] <= {T2TtoM[t] + 20}");
                        DataRow[] checkListSelect = dt_checkList.Select($@"UID = '{underlyingID}' AND [StrikePrice] <= {strike + Math.Round(UClosePrice[underlyingID] * K_Ratio_Put, 1)} AND [StrikePrice] >= {strike - Math.Round(UClosePrice[underlyingID] * K_Ratio_Put, 1)} AND [WClass] = 'p' AND [TtoM] >= {T2TtoM[t] - diff_T_Put} AND [TtoM] <= {T2TtoM[t] + diff_T_Put} AND [IssuerName] <> '9200'");
                        if (checkListSelect.Length > 0)
                        {
                            double MinVol = 999;
                            foreach (DataRow dr_check in checkListSelect)
                            {
                                double iv = Convert.ToDouble(dr_check["IV"].ToString());
                                if (iv <= MinVol)
                                    MinVol = iv;
                            }
                            DataRow[] checkListSelect_9200 = dt_checkList.Select($@"UID = '{underlyingID}' AND [StrikePrice] <= {strike + Math.Round(UClosePrice[underlyingID] * K_Ratio_Put, 1)} AND [StrikePrice] >= {strike - Math.Round(UClosePrice[underlyingID] * K_Ratio_Put, 1)} AND [WClass] = 'p' AND [TtoM] >= {T2TtoM[t] - diff_T_Put} AND [TtoM] <= {T2TtoM[t] + diff_T_Put} AND [IssuerName] = '9200'");
                            if (checkListSelect_9200.Length > 0)
                            {
                                double MinVol_9200 = 999;
                                foreach (DataRow dr_check in checkListSelect_9200)
                                {
                                    double iv = Convert.ToDouble(dr_check["IV"].ToString());
                                    if (iv <= MinVol_9200)
                                        MinVol_9200 = iv;
                                }
                                if (MinVol_9200 <= MinVol * volRatio)
                                {
                                    dr.Cells["覆"].Appearance.BackColor = Color.MediumPurple;
                                    dr.Cells["覆"].ToolTipText += $@"自家已有覆蓋權證";
                                    sameCnt++;
                                }
                            }
                            else
                            {
                                if (vol > MinVol * 0.95)
                                {
                                    dr.Cells["覆"].Appearance.BackColor = Color.MediumPurple;
                                    dr.Cells["覆"].ToolTipText += $@"Vol應為{Math.Floor(MinVol * volRatio)}";
                                    sameCnt++;
                                }
                                if (Math.Floor(MinVol * volRatio) <= HV20[underlyingID])
                                {
                                    dr.Cells["覆"].Appearance.BackColor = Color.MediumPurple;
                                    dr.Cells["覆"].ToolTipText += $@"覆蓋率的VOL被HV20穿";
                                }
                                if (Math.Floor(MinVol * volRatio) <= HV60[underlyingID])
                                {
                                    dr.Cells["覆"].Appearance.BackColor = Color.MediumPurple;
                                    dr.Cells["覆"].ToolTipText += $@"覆蓋率的VOL被HV60穿";
                                }
                            }

                        }
                        else
                        {
                            dr.Cells["覆"].Appearance.BackColor = Color.MediumPurple;
                            dr.Cells["覆"].ToolTipText += $@"其他三家沒發";
                            sameCnt++;
                        }
                    }
                }
            }
            return sameCnt;

        }

        private void toolStripButton4_Click(object sender, EventArgs e)
        {
            int cnt = CheckData_SameIssue();
            MessageBox.Show($@"{cnt}權證重複!");
        }

        private void toolStripButton5_Click(object sender, EventArgs e)
        {
            int cnt = CheckData_OverlapIssue();
            MessageBox.Show($@"{cnt}權證重複!");
        }
    }
}

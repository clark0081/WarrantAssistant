#define To39
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using EDLib.SQL;
using Infragistics.Win;
using Infragistics.Win.UltraWinGrid;
using System.Data.SqlClient;

namespace WarrantAssistant
{
    public partial class FrmAutoSelect : Form
    {
        DateTime today = DateTime.Today;
        public string userID = GlobalVar.globalParameter.userID;
        private DataTable dt = new DataTable();
        private DataTable dt_CMnoey = new DataTable();
        private bool isEdit = false;
        private System.Data.DataTable UnderlyingLiquidity = new System.Data.DataTable();//標的分級大中小
        private System.Data.DataTable UnderlyingRank = new System.Data.DataTable();//標的分級ABCD，以Call為主
        public FrmAutoSelect()
        {
            InitializeComponent();
        }

        private void FrmAutoSelect_Load(object sender, EventArgs e)
        {
            InitialGrid();
            string sql = $@"SELECT  CASE WHEN([UID] ='TWA00') THEN 'IX0001' ELSE [UID] END AS [UID], [Class]
                          FROM [newEDIS].[dbo].[UnderlyingRank]
                          WHERE [TDate] = (SELECT Max(TDate) FROM [newEDIS].[dbo].[UnderlyingRank]) AND [CallPut] ='C'";
            UnderlyingRank = MSSQL.ExecSqlQry(sql, "SERVER=10.60.0.37;DATABASE=newEDIS;UID=WarrantWeb;PWD=WarrantWeb");

            sql = $@"SELECT  [UID], [UnderlyingSize]
                          FROM [newEDIS].[dbo].[UnderlyingLiquidity]
                          WHERE [TDate] = (SELECT Max(TDate) FROM [newEDIS].[dbo].[UnderlyingLiquidity]) ";
            UnderlyingLiquidity = MSSQL.ExecSqlQry(sql, "SERVER=10.60.0.37;DATABASE=newEDIS;UID=WarrantWeb;PWD=WarrantWeb");
            LoadData();
            LoadCMoney();
        }

        

        private void InitialGrid()
        {
            dt.Columns.Add("標的代號",typeof(string));
            dt.Columns.Add("標的名稱", typeof(string));
            dt.Columns.Add("標的分級", typeof(string));
            dt.Columns.Add("發Put", typeof(bool));
            dt.Columns.Add("篩選", typeof(bool));
            dt.Columns.Add("一個月BrokerPL(%)", typeof(float));
            //dt.Columns.Add("標的月損益(仟)", typeof(float));
            dt.Columns.Add("市場剩餘額度(檔)", typeof(int));
            dt.Columns.Add("有額度釋出", typeof(int));
            dt.Columns.Add("今日解禁可發行", typeof(string));
            dt.Columns.Add("3日漲幅(%)", typeof(float));
            dt.Columns.Add("3日跌幅(%)", typeof(float));
            dt.Columns.Add("市場Theta IV", typeof(float));
            dt.Columns.Add("市場部位近五日變化(萬)", typeof(float));
            //dt.Columns.Add("市場Theta IV 金額週變化(部位)", typeof(float));
            dt.Columns.Add("市場 Med_Vol / HV_60D", typeof(float));
            dt.Columns.Add("市場超額利潤(仟)", typeof(float));
            dt.Columns.Add("平均每檔權證超額利潤(元)", typeof(float));
            dt.Columns.Add("Theta天數", typeof(float));
            dt.Columns.Add("評鑑權證比重排名", typeof(float));
            dt.Columns.Add("融資使用率(%)", typeof(float));
            dt.Columns.Add("C有效檔數市佔(%)", typeof(float));
            dt.Columns.Add("P有效檔數市佔(%)", typeof(float));
            dt.Columns.Add("C發行密度", typeof(float));
            dt.Columns.Add("P發行密度", typeof(float));
            dt.Columns.Add("凱基-元大C發行密度", typeof(float));
            dt.Columns.Add("凱基-元大P發行密度", typeof(float));
            dt.Columns.Add("履約價重覆檢查(%)", typeof(float));
            dt.Columns.Add("到期日重覆檢查(月)", typeof(float));

            this.ultraGrid1.DisplayLayout.Override.DefaultRowHeight = 30;
            this.ultraGrid1.DisplayLayout.AutoFitStyle = Infragistics.Win.UltraWinGrid.AutoFitStyle.ResizeAllColumns;
            this.ultraGrid1.DisplayLayout.Override.WrapHeaderText = Infragistics.Win.DefaultableBoolean.True;
            this.ultraGrid1.DisplayLayout.Override.CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center;
            this.ultraGrid1.DisplayLayout.Override.CellAppearance.TextVAlign = Infragistics.Win.VAlign.Middle;
            this.ultraGrid1.DisplayLayout.Override.SelectTypeCell = Infragistics.Win.UltraWinGrid.SelectType.Extended;
            //this.ultraGrid1.DisplayLayout.Override.SelectTypeRow = Infragistics.Win.UltraWinGrid.SelectType.Extended;
            ultraGrid1.DataSource = dt;
            SetButton();
        }

        private void LoadData()
        {
            dt.Clear();
#if !To39
            string sql = $@"SELECT  [UID], [UName], [Trader], [Checked], [BrokerPL_Month], [Profit_Month]
                        , [OptionAvailable], [OptionRelease], [RiseUp_3Days], [DropDown_3Days], [ThetaIV_WeekDelta]
                        , [Med_HV60D_VolRatio], [Theta_Days], [AppraisalRank], [FinancingRatio], [CallMarketShare]
                        , [PutMarketShare], [CallDensity], [PutDensity], [K_OverLap], [T_Overlap], [IssuePut], [Theta_EndDate]
                        , [KgiYuanCallDensity], [KgiYuanPutDensity]
                        , [MarketAmtChg5Days], [AlphaTheta], [AvgAlphaThetaCost]
                        FROM [newEDIS].[dbo].[OptionAutoSelect]
                        WHERE [Trader] = '{userID.TrimStart('0')}'
                        ORDER BY [UID]";
            DataTable dttemp = MSSQL.ExecSqlQry(sql, GlobalVar.loginSet.newEDIS);
#else
            string sql = $@"SELECT  [UID], [UName], [Trader], [Checked], [BrokerPL_Month], [Profit_Month]
                        , [OptionAvailable], [OptionRelease], [RiseUp_3Days], [DropDown_3Days], [ThetaIV_WeekDelta]
                        , [Med_HV60D_VolRatio], [Theta_Days], [AppraisalRank], [FinancingRatio], [CallMarketShare]
                        , [PutMarketShare], [CallDensity], [PutDensity], [K_OverLap], [T_Overlap], [IssuePut], [Theta_EndDate]
                        , [KgiYuanCallDensity], [KgiYuanPutDensity]
                        , [MarketAmtChg5Days], [AlphaTheta], [AvgAlphaThetaCost]
                        FROM [WarrantAssistant].[dbo].[OptionAutoSelect]
                        WHERE [Trader] = '{userID.TrimStart('0')}'
                        ORDER BY [UID]";
            DataTable dttemp = MSSQL.ExecSqlQry(sql, GlobalVar.loginSet.warrantassistant45);
#endif
            foreach (DataRow dr in dttemp.Rows)
            {
                DataRow drv = dt.NewRow();
                
                string uid = dr["UID"].ToString();
                drv["標的代號"] = uid;

                string uname = dr["UName"].ToString();
                drv["標的名稱"] = uname;

                string Big3 = "X";
                string Big4 = "X";
                DataRow[] tempRow = UnderlyingLiquidity.Select($@"UID='{uid}'");
                if (tempRow.Length > 0)
                {
                    string X = tempRow[0][1].ToString();
                    if (X == "A")
                        Big3 = "大";
                    else if (X == "B")
                        Big3 = "中";
                    else
                        Big3 = "小";
                }
                DataRow[] tempRow2 = UnderlyingRank.Select($@"UID='{uid}'");
                if (tempRow2.Length > 0)
                {
                    Big4 = tempRow2[0][1].ToString();
                }
                drv["標的分級"] = Big3 + Big4;

                string trader = dr["Trader"].ToString();

                string issueput = dr["IssuePut"].ToString();
                drv["發Put"] = Convert.ToInt32(issueput);

                string check = dr["Checked"].ToString();
                drv["篩選"] = Convert.ToInt32(check);
                //drv["篩選"] = 1;

                string brokerPL_month = dr["BrokerPL_Month"].ToString();
                drv["一個月BrokerPL(%)"] = Convert.ToDouble(brokerPL_month);
                /*
                string profit_month = dr["Profit_Month"].ToString();
                drv["標的月損益(仟)"] = Convert.ToDouble(profit_month);
                */
                string optionavailable = dr["OptionAvailable"].ToString();
                drv["市場剩餘額度(檔)"] = Convert.ToInt32(optionavailable);

                string optionrelease = dr["OptionRelease"].ToString();
                drv["有額度釋出"] = Convert.ToInt32(optionrelease);

                drv["今日解禁可發行"] = "";

                string riseup_3days = dr["RiseUp_3Days"].ToString();
                drv["3日漲幅(%)"] = Convert.ToDouble(riseup_3days);

                string dropdown_3days = dr["DropDown_3Days"].ToString();
                drv["3日跌幅(%)"] = Convert.ToDouble(dropdown_3days);
                /*
                string thetaiv_weekdelta = dr["ThetaIV_WeekDelta"].ToString();
                drv["市場Theta IV 金額週變化(部位)"] = Convert.ToDouble(thetaiv_weekdelta);
                */
                string theta_enddate = dr["Theta_EndDate"].ToString();
                drv["市場Theta IV"] = Convert.ToDouble(theta_enddate);

                string marketamtchg = dr["MarketAmtChg5Days"].ToString();
                drv["市場部位近五日變化(萬)"] = Convert.ToDouble(marketamtchg);

                string med_hv60d_volratio = dr["Med_HV60D_VolRatio"].ToString();
                drv["市場 Med_Vol / HV_60D"] = Convert.ToDouble(med_hv60d_volratio);

                string alphatheta = dr["AlphaTheta"].ToString();
                drv["市場超額利潤(仟)"] = Convert.ToDouble(alphatheta);

                string avgalphathetacost = dr["AvgAlphaThetaCost"].ToString();
                drv["平均每檔權證超額利潤(元)"] = Convert.ToDouble(avgalphathetacost);

                string theta_days = dr["Theta_Days"].ToString();
                drv["Theta天數"] = Convert.ToDouble(theta_days);

                string appraisalrank = dr["AppraisalRank"].ToString();
                drv["評鑑權證比重排名"] = Convert.ToDouble(appraisalrank);

                string financingratio = dr["FinancingRatio"].ToString();
                drv["融資使用率(%)"] = Convert.ToDouble(financingratio);

                string callmarketshare = dr["CallMarketShare"].ToString();
                drv["C有效檔數市佔(%)"] = Convert.ToDouble(callmarketshare);

                string putmarketshare = dr["PutMarketShare"].ToString();
                drv["P有效檔數市佔(%)"] = Convert.ToDouble(putmarketshare);

                string calldensity = dr["CallDensity"].ToString();
                drv["C發行密度"] = Convert.ToDouble(calldensity);

                string putdensity = dr["PutDensity"].ToString();
                drv["P發行密度"] = Convert.ToDouble(putdensity);

                string kgiyuancalldensity = dr["KgiYuanCallDensity"].ToString();
                drv["凱基-元大C發行密度"] = Convert.ToDouble(kgiyuancalldensity);

                string kgiyuanputdensity = dr["KgiYuanPutDensity"].ToString();
                drv["凱基-元大P發行密度"] = Convert.ToDouble(kgiyuanputdensity);

                string k_overLap = dr["K_OverLap"].ToString();
                drv["履約價重覆檢查(%)"] = Convert.ToDouble(k_overLap);

                string t_overLap = dr["T_OverLap"].ToString();
                drv["到期日重覆檢查(月)"] = Convert.ToDouble(t_overLap);

                dt.Rows.Add(drv);
            }

        }
        private void LoadCMoney()
        {
            dt_CMnoey.Clear();
#if !To39
            string sql = $@"SELECT [UID], CONVERT(varchar,[DisposeEndDate],112) AS DisposeEndDate 
                          , [LastWatchCount], [WatchCount], [LastWarningScore], [WarningScore]
                            FROM [newEDIS].[dbo].[OptionAutoSelectData]";
            dt_CMnoey = MSSQL.ExecSqlQry(sql, GlobalVar.loginSet.newEDIS);
#else
            string sql = $@"SELECT [UID], CONVERT(varchar,[DisposeEndDate],112) AS DisposeEndDate 
                          , [LastWatchCount], [WatchCount], [LastWarningScore], [WarningScore]
                            FROM [WarrantAssistant].[dbo].[OptionAutoSelectData]";
            dt_CMnoey = MSSQL.ExecSqlQry(sql, GlobalVar.loginSet.warrantassistant45);
#endif

        }
        private void LoadDataByUID()
        {
            string str = textBox1.Text;
            if (!isEdit)
            {
                if (str == "")
                {
                    LoadData();
                    return;
                }
                dt.Clear();
#if !To39
                string sql = $@"SELECT  [UID], [UName], [Trader], [Checked], [BrokerPL_Month], [Profit_Month]
                        , [OptionAvailable], [OptionRelease], [RiseUp_3Days], [DropDown_3Days], [ThetaIV_WeekDelta]
                        , [Med_HV60D_VolRatio], [Theta_Days], [AppraisalRank], [FinancingRatio], [CallMarketShare]
                        , [PutMarketShare], [CallDensity], [PutDensity], [K_OverLap], [T_Overlap], [IssuePut], [Theta_EndDate]
                        , [KgiYuanCallDensity], [KgiYuanPutDensity]
                        , [MarketAmtChg5Days], [AlphaTheta], [AvgAlphaThetaCost]
                        FROM [newEDIS].[dbo].[OptionAutoSelect]
                        WHERE [Trader] = '{userID.TrimStart('0')}' AND [UID] ='{str}'";
                DataTable dttemp = MSSQL.ExecSqlQry(sql, GlobalVar.loginSet.newEDIS);
#else
                string sql = $@"SELECT  [UID], [UName], [Trader], [Checked], [BrokerPL_Month], [Profit_Month]
                        , [OptionAvailable], [OptionRelease], [RiseUp_3Days], [DropDown_3Days], [ThetaIV_WeekDelta]
                        , [Med_HV60D_VolRatio], [Theta_Days], [AppraisalRank], [FinancingRatio], [CallMarketShare]
                        , [PutMarketShare], [CallDensity], [PutDensity], [K_OverLap], [T_Overlap], [IssuePut], [Theta_EndDate]
                        , [KgiYuanCallDensity], [KgiYuanPutDensity]
                        , [MarketAmtChg5Days], [AlphaTheta], [AvgAlphaThetaCost]
                        FROM [WarrantAssistant].[dbo].[OptionAutoSelect]
                        WHERE [Trader] = '{userID.TrimStart('0')}' AND [UID] ='{str}'";
                DataTable dttemp = MSSQL.ExecSqlQry(sql, GlobalVar.loginSet.warrantassistant45);
#endif
                foreach (DataRow dr in dttemp.Rows)
                {
                    DataRow drv = dt.NewRow();

                    string uid = dr["UID"].ToString();
                    drv["標的代號"] = uid;

                    string uname = dr["UName"].ToString();
                    drv["標的名稱"] = uname;

                    string Big3 = "X";
                    string Big4 = "X";
                    DataRow[] tempRow = UnderlyingLiquidity.Select($@"UID='{uid}'");
                    if (tempRow.Length > 0)
                    {
                        string X = tempRow[0][1].ToString();
                        if (X == "A")
                            Big3 = "大";
                        else if (X == "B")
                            Big3 = "中";
                        else
                            Big3 = "小";
                    }
                    DataRow[] tempRow2 = UnderlyingRank.Select($@"UID='{uid}'");
                    if (tempRow2.Length > 0)
                    {
                        Big4 = tempRow2[0][1].ToString();
                    }
                    drv["標的分級"] = Big3 + Big4;

                    string trader = dr["Trader"].ToString();

                    string issueput = dr["IssuePut"].ToString();
                    drv["發Put"] = Convert.ToInt32(issueput);

                    string check = dr["Checked"].ToString();
                    drv["篩選"] = Convert.ToInt32(check);
                    //drv["篩選"] = 1;

                    string brokerPL_month = dr["BrokerPL_Month"].ToString();
                    drv["一個月BrokerPL(%)"] = Convert.ToDouble(brokerPL_month);
                    /*
                    string profit_month = dr["Profit_Month"].ToString();
                    drv["標的月損益(仟)"] = Convert.ToDouble(profit_month);
                    */
                    string optionavailable = dr["OptionAvailable"].ToString();
                    drv["市場剩餘額度(檔)"] = Convert.ToInt32(optionavailable);

                    string optionrelease = dr["OptionRelease"].ToString();
                    drv["有額度釋出"] = Convert.ToInt32(optionrelease);

                    drv["今日解禁可發行"] = "";

                    string riseup_3days = dr["RiseUp_3Days"].ToString();
                    drv["3日漲幅(%)"] = Convert.ToDouble(riseup_3days);

                    string dropdown_3days = dr["DropDown_3Days"].ToString();
                    drv["3日跌幅(%)"] = Convert.ToDouble(dropdown_3days);
                    /*
                    string thetaiv_weekdelta = dr["ThetaIV_WeekDelta"].ToString();
                    drv["市場Theta IV 金額週變化(部位)"] = Convert.ToDouble(thetaiv_weekdelta);
                    */
                    string theta_enddate = dr["Theta_EndDate"].ToString();
                    drv["市場Theta IV"] = Convert.ToDouble(theta_enddate);

                    string marketamtchg = dr["MarketAmtChg5Days"].ToString();
                    drv["市場部位近五日變化(萬)"] = Convert.ToDouble(marketamtchg);

                    string med_hv60d_volratio = dr["Med_HV60D_VolRatio"].ToString();
                    drv["市場 Med_Vol / HV_60D"] = Convert.ToDouble(med_hv60d_volratio);

                    string alphatheta = dr["AlphaTheta"].ToString();
                    drv["市場超額利潤(仟)"] = Convert.ToDouble(alphatheta);

                    string avgalphathetacost = dr["AvgAlphaThetaCost"].ToString();
                    drv["平均每檔權證超額利潤(元)"] = Convert.ToDouble(avgalphathetacost);

                    string theta_days = dr["Theta_Days"].ToString();
                    drv["Theta天數"] = Convert.ToDouble(theta_days);

                    string appraisalrank = dr["AppraisalRank"].ToString();
                    drv["評鑑權證比重排名"] = Convert.ToDouble(appraisalrank);

                    string financingratio = dr["FinancingRatio"].ToString();
                    drv["融資使用率(%)"] = Convert.ToDouble(financingratio);

                    string callmarketshare = dr["CallMarketShare"].ToString();
                    drv["C有效檔數市佔(%)"] = Convert.ToDouble(callmarketshare);

                    string putmarketshare = dr["PutMarketShare"].ToString();
                    drv["P有效檔數市佔(%)"] = Convert.ToDouble(putmarketshare);

                    string calldensity = dr["CallDensity"].ToString();
                    drv["C發行密度"] = Convert.ToDouble(calldensity);

                    string putdensity = dr["PutDensity"].ToString();
                    drv["P發行密度"] = Convert.ToDouble(putdensity);

                    string kgiyuancalldensity = dr["KgiYuanCallDensity"].ToString();
                    drv["凱基-元大C發行密度"] = Convert.ToDouble(kgiyuancalldensity);

                    string kgiyuanputdensity = dr["KgiYuanPutDensity"].ToString();
                    drv["凱基-元大P發行密度"] = Convert.ToDouble(kgiyuanputdensity);

                    string k_overLap = dr["K_OverLap"].ToString();
                    drv["履約價重覆檢查(%)"] = Convert.ToDouble(k_overLap);

                    string t_overLap = dr["T_OverLap"].ToString();
                    drv["到期日重覆檢查(月)"] = Convert.ToDouble(t_overLap);

                    dt.Rows.Add(drv);
                }
            }
            else
            {
                MessageBox.Show("編輯模式無法搜尋");
            }

        }
        private void SetButton()
        {
            UltraGridBand bands0 = ultraGrid1.DisplayLayout.Bands[0];
            if (isEdit)
            {
                bands0.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
                bands0.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
                bands0.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;


                bands0.Columns["標的代號"].CellActivation = Activation.NoEdit;
                bands0.Columns["標的名稱"].CellActivation = Activation.NoEdit;
                bands0.Columns["標的分級"].CellActivation = Activation.NoEdit;
                bands0.Columns["發Put"].CellActivation = Activation.AllowEdit;
                bands0.Columns["篩選"].CellActivation = Activation.AllowEdit;
                bands0.Columns["一個月BrokerPL(%)"].CellActivation = Activation.AllowEdit;
                //bands0.Columns["標的月損益(仟)"].CellActivation = Activation.AllowEdit;
                bands0.Columns["市場剩餘額度(檔)"].CellActivation = Activation.AllowEdit;
                bands0.Columns["有額度釋出"].CellActivation = Activation.AllowEdit;
                bands0.Columns["今日解禁可發行"].CellActivation = Activation.NoEdit;
                bands0.Columns["3日漲幅(%)"].CellActivation = Activation.AllowEdit;
                bands0.Columns["3日跌幅(%)"].CellActivation = Activation.AllowEdit;
                bands0.Columns["市場Theta IV"].CellActivation = Activation.AllowEdit;
                bands0.Columns["市場部位近五日變化(萬)"].CellActivation = Activation.AllowEdit;
                //bands0.Columns["市場Theta IV 金額週變化(部位)"].CellActivation = Activation.AllowEdit;
                bands0.Columns["市場 Med_Vol / HV_60D"].CellActivation = Activation.AllowEdit;
                bands0.Columns["市場超額利潤(仟)"].CellActivation = Activation.AllowEdit;
                bands0.Columns["平均每檔權證超額利潤(元)"].CellActivation = Activation.AllowEdit;
                bands0.Columns["Theta天數"].CellActivation = Activation.AllowEdit;
                bands0.Columns["評鑑權證比重排名"].CellActivation = Activation.AllowEdit;
                bands0.Columns["融資使用率(%)"].CellActivation = Activation.AllowEdit;
                bands0.Columns["C有效檔數市佔(%)"].CellActivation = Activation.AllowEdit;
                bands0.Columns["P有效檔數市佔(%)"].CellActivation = Activation.AllowEdit;
                bands0.Columns["C發行密度"].CellActivation = Activation.AllowEdit;
                bands0.Columns["P發行密度"].CellActivation = Activation.AllowEdit;
                bands0.Columns["凱基-元大C發行密度"].CellActivation = Activation.AllowEdit;
                bands0.Columns["凱基-元大P發行密度"].CellActivation = Activation.AllowEdit;
                bands0.Columns["履約價重覆檢查(%)"].CellActivation = Activation.AllowEdit;
                bands0.Columns["到期日重覆檢查(月)"].CellActivation = Activation.AllowEdit;
                ultraGrid1.DisplayLayout.Override.CellAppearance.BackColor = Color.White;
                button1.Visible = false;
                button2.Visible = true;
                button3.Visible = true;

            }
            else
            {
                bands0.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
                bands0.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
                bands0.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
                bands0.Columns["標的代號"].CellActivation = Activation.NoEdit;
                bands0.Columns["標的名稱"].CellActivation = Activation.NoEdit;
                bands0.Columns["標的分級"].CellActivation = Activation.NoEdit;
                bands0.Columns["發Put"].CellActivation = Activation.NoEdit;
                bands0.Columns["篩選"].CellActivation = Activation.NoEdit;
                bands0.Columns["一個月BrokerPL(%)"].CellActivation = Activation.NoEdit;
                //bands0.Columns["標的月損益(仟)"].CellActivation = Activation.NoEdit;
                bands0.Columns["市場剩餘額度(檔)"].CellActivation = Activation.NoEdit;
                bands0.Columns["有額度釋出"].CellActivation = Activation.NoEdit;
                bands0.Columns["今日解禁可發行"].CellActivation = Activation.NoEdit;
                bands0.Columns["3日漲幅(%)"].CellActivation = Activation.NoEdit;
                bands0.Columns["3日跌幅(%)"].CellActivation = Activation.NoEdit;
                bands0.Columns["市場Theta IV"].CellActivation = Activation.NoEdit;
                bands0.Columns["市場部位近五日變化(萬)"].CellActivation = Activation.NoEdit;
                //bands0.Columns["市場Theta IV 金額週變化(部位)"].CellActivation = Activation.NoEdit;
                bands0.Columns["市場 Med_Vol / HV_60D"].CellActivation = Activation.NoEdit;
                bands0.Columns["市場超額利潤(仟)"].CellActivation = Activation.NoEdit;
                bands0.Columns["平均每檔權證超額利潤(元)"].CellActivation = Activation.NoEdit;
                bands0.Columns["Theta天數"].CellActivation = Activation.NoEdit;
                bands0.Columns["評鑑權證比重排名"].CellActivation = Activation.NoEdit;
                bands0.Columns["融資使用率(%)"].CellActivation = Activation.NoEdit;
                bands0.Columns["C有效檔數市佔(%)"].CellActivation = Activation.NoEdit;
                bands0.Columns["P有效檔數市佔(%)"].CellActivation = Activation.NoEdit;
                bands0.Columns["C發行密度"].CellActivation = Activation.NoEdit;
                bands0.Columns["P發行密度"].CellActivation = Activation.NoEdit;
                bands0.Columns["凱基-元大C發行密度"].CellActivation = Activation.NoEdit;
                bands0.Columns["凱基-元大P發行密度"].CellActivation = Activation.NoEdit;
                bands0.Columns["履約價重覆檢查(%)"].CellActivation = Activation.NoEdit;
                bands0.Columns["到期日重覆檢查(月)"].CellActivation = Activation.NoEdit;
                ultraGrid1.DisplayLayout.Override.CellAppearance.BackColor = Color.Moccasin;
                button1.Visible = true;
                button2.Visible = false;
                button3.Visible = false;
            }
        }
        private void ButtonEdit_Click(object sender, EventArgs e)
        {
            isEdit = true;
            SetButton();
        }

        private void buttonConfirm_Click(object sender, EventArgs e)
        {
            
            int rownum = ultraGrid1.Rows.Count;
            if (rownum == 1)
            {

            }
            else
            {
#if !To39
                string sqldel = $@"DELETE [newEDIS].[dbo].[OptionAutoSelect]
                                WHERE [Trader] ='{userID.TrimStart('0')}'";
                MSSQL.ExecSqlCmd(sqldel, GlobalVar.loginSet.newEDIS);
                string sql = $@"INSERT INTO [newEDIS].[dbo].[OptionAutoSelect] ( [UID], [UName], [Trader], [Checked], [BrokerPL_Month], [Profit_Month]
                              , [OptionAvailable], [OptionRelease], [RiseUp_3Days], [DropDown_3Days], [ThetaIV_WeekDelta], [Med_HV60D_VolRatio]
                              , [Theta_Days], [AppraisalRank], [FinancingRatio], [CallMarketShare], [PutMarketShare], [CallDensity]
                              , [PutDensity], [K_OverLap], [T_Overlap], [IssuePut], [Theta_EndDate], [KgiYuanCallDensity], [KgiYuanPutDensity]
                              , [MarketAmtChg5Days], [AlphaTheta], [AvgAlphaThetaCost]) 
                                VALUES( @UID, @UName, @Trader, @Checked, @BrokerPL_Month, 0
                              , @OptionAvailable, @OptionRelease, @RiseUp_3Days, @DropDown_3Days, 0, @Med_HV60D_VolRatio, @Theta_Days
                              , @AppraisalRank, @FinancingRatio, @CallMarketShare, @PutMarketShare, @CallDensity, @PutDensity, @K_OverLap, @T_Overlap,@IssuePut,@Theta_EndDate, @KgiYuanCallDensity, @KgiYuanPutDensity
                              , @MarketAmtChg5Days, @AlphaTheta, @AvgAlphaThetaCost)";
#else
                string sqldel = $@"DELETE [WarrantAssistant].[dbo].[OptionAutoSelect]
                                WHERE [Trader] ='{userID.TrimStart('0')}'";
                MSSQL.ExecSqlCmd(sqldel, GlobalVar.loginSet.warrantassistant45);
                string sql = $@"INSERT INTO [WarrantAssistant].[dbo].[OptionAutoSelect] ( [UID], [UName], [Trader], [Checked], [BrokerPL_Month], [Profit_Month]
                              , [OptionAvailable], [OptionRelease], [RiseUp_3Days], [DropDown_3Days], [ThetaIV_WeekDelta], [Med_HV60D_VolRatio]
                              , [Theta_Days], [AppraisalRank], [FinancingRatio], [CallMarketShare], [PutMarketShare], [CallDensity]
                              , [PutDensity], [K_OverLap], [T_Overlap], [IssuePut], [Theta_EndDate], [KgiYuanCallDensity], [KgiYuanPutDensity]
                              , [MarketAmtChg5Days], [AlphaTheta], [AvgAlphaThetaCost]) 
                                VALUES( @UID, @UName, @Trader, @Checked, @BrokerPL_Month, 0
                              , @OptionAvailable, @OptionRelease, @RiseUp_3Days, @DropDown_3Days, 0, @Med_HV60D_VolRatio, @Theta_Days
                              , @AppraisalRank, @FinancingRatio, @CallMarketShare, @PutMarketShare, @CallDensity, @PutDensity, @K_OverLap, @T_Overlap,@IssuePut,@Theta_EndDate, @KgiYuanCallDensity, @KgiYuanPutDensity
                              , @MarketAmtChg5Days, @AlphaTheta, @AvgAlphaThetaCost)";
#endif
                List<SqlParameter> ps = new List<SqlParameter> {
                    new SqlParameter("@UID", SqlDbType.VarChar),
                    new SqlParameter("@UName", SqlDbType.VarChar),
                    new SqlParameter("@Trader", SqlDbType.VarChar),
                    new SqlParameter("@Checked", SqlDbType.Int),
                    new SqlParameter("@BrokerPL_Month", SqlDbType.Float),
                    new SqlParameter("@OptionAvailable", SqlDbType.Int),
                    new SqlParameter("@OptionRelease", SqlDbType.Int),
                    new SqlParameter("@RiseUp_3Days", SqlDbType.Float),
                    new SqlParameter("@DropDown_3Days", SqlDbType.Float),
                    new SqlParameter("@Med_HV60D_VolRatio", SqlDbType.Float),
                    new SqlParameter("@Theta_Days", SqlDbType.Float),
                    new SqlParameter("@AppraisalRank", SqlDbType.Float),
                    new SqlParameter("@FinancingRatio", SqlDbType.Float),
                    new SqlParameter("@CallMarketShare", SqlDbType.Float),
                    new SqlParameter("@PutMarketShare", SqlDbType.Float),
                    new SqlParameter("@CallDensity", SqlDbType.Float),
                    new SqlParameter("@PutDensity", SqlDbType.Float),
                    new SqlParameter("@K_OverLap", SqlDbType.Float),
                    new SqlParameter("@T_Overlap", SqlDbType.Float),
                    new SqlParameter("@IssuePut", SqlDbType.Int),
                    new SqlParameter("@Theta_EndDate", SqlDbType.Float),
                    new SqlParameter("@KgiYuanCallDensity", SqlDbType.Float),
                    new SqlParameter("@KgiYuanPutDensity", SqlDbType.Float),
                    new SqlParameter("@MarketAmtChg5Days", SqlDbType.Float),
                    new SqlParameter("@AlphaTheta", SqlDbType.Float),
                    new SqlParameter("@AvgAlphaThetaCost", SqlDbType.Float)
                };
#if !To39
                SQLCommandHelper h = new SQLCommandHelper(GlobalVar.loginSet.newEDIS, sql, ps);
#else
                SQLCommandHelper h = new SQLCommandHelper(GlobalVar.loginSet.warrantassistant45, sql, ps);
#endif
                foreach (Infragistics.Win.UltraWinGrid.UltraGridRow r in ultraGrid1.Rows)
                {

                    try
                    {
                        string userid = userID.TrimStart('0');
                        h.SetParameterValue("@Trader", userid);
                        string uid = r.Cells["標的代號"].Value.ToString();
                        h.SetParameterValue("@UID", uid);
                        string uname = r.Cells["標的名稱"].Value.ToString();
                        h.SetParameterValue("@UName", uname);
                        int check = Convert.ToInt32(r.Cells["篩選"].Value);
                        h.SetParameterValue("@Checked", check);
                        double brokerPL_month = r.Cells["一個月BrokerPL(%)"].Value == DBNull.Value ? 0 : Convert.ToDouble(r.Cells["一個月BrokerPL(%)"].Value);
                        h.SetParameterValue("@BrokerPL_Month", brokerPL_month);
                        /*
                        double profit_month = r.Cells["標的月損益(仟)"].Value == DBNull.Value ? 0 : Convert.ToDouble(r.Cells["標的月損益(仟)"].Value);
                        h.SetParameterValue("@Profit_Month", profit_month);
                        */
                        int optionavailable = r.Cells["市場剩餘額度(檔)"].Value == DBNull.Value ? 0 : Convert.ToInt32(r.Cells["市場剩餘額度(檔)"].Value);
                        h.SetParameterValue("@OptionAvailable", optionavailable);
                        int optionrelease = r.Cells["有額度釋出"].Value == DBNull.Value ? 0 : Convert.ToInt32(r.Cells["有額度釋出"].Value);
                        h.SetParameterValue("@OptionRelease", optionrelease);
                        double riseup_3days = r.Cells["3日漲幅(%)"].Value == DBNull.Value ? 0 : Convert.ToDouble(r.Cells["3日漲幅(%)"].Value);
                        h.SetParameterValue("@RiseUp_3Days", riseup_3days);
                        double dropdown_3days = r.Cells["3日跌幅(%)"].Value == DBNull.Value ? 0 : Convert.ToDouble(r.Cells["3日跌幅(%)"].Value);
                        h.SetParameterValue("@DropDown_3Days", dropdown_3days);
                        double theta_enddate = r.Cells["市場Theta IV"].Value == DBNull.Value ? 0 : Convert.ToDouble(r.Cells["市場Theta IV"].Value);
                        h.SetParameterValue("@Theta_EndDate", theta_enddate);
                        /*
                        double thetaiv_weekdelta = r.Cells["市場Theta IV 金額週變化(部位)"].Value == DBNull.Value ? 0 : Convert.ToDouble(r.Cells["市場Theta IV 金額週變化(部位)"].Value);
                        h.SetParameterValue("@ThetaIV_WeekDelta", thetaiv_weekdelta);
                        */
                        double med_hv60d_volratio = r.Cells["市場 Med_Vol / HV_60D"].Value == DBNull.Value ? 0 : Convert.ToDouble(r.Cells["市場 Med_Vol / HV_60D"].Value);
                        h.SetParameterValue("@Med_HV60D_VolRatio", med_hv60d_volratio);
                        double theta_days = r.Cells["Theta天數"].Value == DBNull.Value ? 0 : Convert.ToDouble(r.Cells["Theta天數"].Value);
                        h.SetParameterValue("@Theta_Days", theta_days);
                        double appraisalrank = r.Cells["評鑑權證比重排名"].Value == DBNull.Value ? 0 : Convert.ToDouble(r.Cells["評鑑權證比重排名"].Value);
                        h.SetParameterValue("@AppraisalRank", appraisalrank);
                        double financingratio = r.Cells["融資使用率(%)"].Value == DBNull.Value ? 0 : Convert.ToDouble(r.Cells["融資使用率(%)"].Value);
                        h.SetParameterValue("@FinancingRatio", financingratio);
                        double callmarketshare = r.Cells["C有效檔數市佔(%)"].Value == DBNull.Value ? 0 : Convert.ToDouble(r.Cells["C有效檔數市佔(%)"].Value);
                        h.SetParameterValue("@CallMarketShare", callmarketshare);
                        double putmarketshare = r.Cells["P有效檔數市佔(%)"].Value == DBNull.Value ? 0 : Convert.ToDouble(r.Cells["P有效檔數市佔(%)"].Value);
                        h.SetParameterValue("@PutMarketShare", putmarketshare);
                        double calldensity = r.Cells["C發行密度"].Value == DBNull.Value ? 0 : Convert.ToDouble(r.Cells["C發行密度"].Value);
                        h.SetParameterValue("@CallDensity", calldensity);
                        double putdensity = r.Cells["P發行密度"].Value == DBNull.Value ? 0 : Convert.ToDouble(r.Cells["P發行密度"].Value);
                        h.SetParameterValue("@PutDensity", putdensity);
                        double kgiyuancalldensity = r.Cells["凱基-元大C發行密度"].Value == DBNull.Value ? 0 : Convert.ToDouble(r.Cells["凱基-元大C發行密度"].Value);
                        h.SetParameterValue("@KgiYuanCallDensity", kgiyuancalldensity);
                        double kgiyuanputdensity = r.Cells["凱基-元大P發行密度"].Value == DBNull.Value ? 0 : Convert.ToDouble(r.Cells["凱基-元大P發行密度"].Value);
                        h.SetParameterValue("@KgiYuanPutDensity", kgiyuanputdensity);
                        double k_overLap = r.Cells["履約價重覆檢查(%)"].Value == DBNull.Value ? 0 : Convert.ToDouble(r.Cells["履約價重覆檢查(%)"].Value);
                        h.SetParameterValue("@K_OverLap", k_overLap);
                        double t_overLap = r.Cells["到期日重覆檢查(月)"].Value == DBNull.Value ? 0 : Convert.ToDouble(r.Cells["到期日重覆檢查(月)"].Value);
                        h.SetParameterValue("@T_Overlap", t_overLap);
                        int issueput = Convert.ToInt32(r.Cells["發Put"].Value);
                        h.SetParameterValue("@IssuePut", issueput);
                        double marketamtchg = r.Cells["市場部位近五日變化(萬)"].Value == DBNull.Value ? 0 : Convert.ToDouble(r.Cells["市場部位近五日變化(萬)"].Value);
                        h.SetParameterValue("@MarketAmtChg5Days", marketamtchg);
                        double alphatheta = r.Cells["市場超額利潤(仟)"].Value == DBNull.Value ? 0 : Convert.ToDouble(r.Cells["市場超額利潤(仟)"].Value);
                        h.SetParameterValue("@AlphaTheta", alphatheta);
                        double avgalphathetacost = r.Cells["平均每檔權證超額利潤(元)"].Value == DBNull.Value ? 0 : Convert.ToDouble(r.Cells["平均每檔權證超額利潤(元)"].Value);
                        h.SetParameterValue("@AvgAlphaThetaCost", avgalphathetacost);

                        h.ExecuteCommand();
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message);
                    }
                }
            }
            isEdit = false;
            SetButton();
            LoadData();
        }

        private void buttonCancel_Click(object sender, EventArgs e)
        {
            isEdit = false;
            SetButton();
            LoadData();
        }

        private void ultraGrid1_InitializeLayout(object sender, InitializeLayoutEventArgs e)
        {
            e.Layout.ScrollBounds = ScrollBounds.ScrollToFill;
        }

        private void textBox1_KeyDown(object sender, KeyEventArgs e)
        {
            try
            {
                if (e.KeyCode == Keys.Enter)
                {
                    LoadDataByUID();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
        private void ultraGrid1_MouseClick(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Right)
            {
                contextMenuStrip1.Show();
            }
        }
        private void ultraGrid1_ClickCell(object sender, Infragistics.Win.UltraWinGrid.DoubleClickCellEventArgs e)
        {
            if (isEdit)
            {
                e.Cell.Selected = true;
            }
        }
        private void Setting(object sender, EventArgs e)
        {
            if (!isEdit)
            {
                MessageBox.Show("非編輯模式!");
                return;
            }
            int colnum = ultraGrid1.DisplayLayout.Bands[0].Columns.Count;
            int rownum = ultraGrid1.Rows.Count;
            if (rownum < 1)
            {
                return;
            }
            int colindex = -1;
            int mode = 0;
            int head = 0;
            int tail = 0;
            for(int i = 0; i < rownum; i++)
            {
                for (int j = 0; j < colnum; j++)
                {
                    if (ultraGrid1.Rows[i].Cells[j].Selected == true)
                    {
                        //MessageBox.Show($@"{ultraGrid1.Rows[i].Cells[j].Value.ToString()} mode {mode}");
                        //initial colindex
                        if (colindex < 0)
                            colindex = j;
                        else
                        {
                            if (colindex != j)
                            {
                                MessageBox.Show($@"不能設定多欄");
                                return;
                            }
                        }
                        if(mode == 0)
                        {
                            mode = 1;
                            head = i;
                        }
                        if(mode == 1)
                        {
                            //如果最下面一格也被選到
                            if(i == rownum - 1)
                            {
                                mode = 0;
                                tail = i;
                            }

                        }
                    }
                    else
                    {
                        //抓tail mode代表上一格有被選到
                        if(mode == 1 && colindex >= 0 && colindex == j)
                        {
                            mode = 0;
                            if (i >= head && i > 0)
                                tail = i-1;
                        }
                    }
                }
            }
            //MessageBox.Show($@"head:{head}  tail:{tail}  rowindex:{rowindex}");
            var headvalue = ultraGrid1.Rows[head].Cells[colindex].Value;
            var tailvalue = ultraGrid1.Rows[tail].Cells[colindex].Value;
            if (colindex <= 2)//名稱跟代號不能改
                return;
            if (headvalue.ToString() == tailvalue.ToString())
            {
                //MessageBox.Show("equal");
                for (int i = head; i <= tail; i++)
                {
                    if (colindex >= 0)
                    {
                        //MessageBox.Show($@"{ultraGrid1.Rows[i].Cells[rowindex].Value}");
                        ultraGrid1.Rows[i].Cells[colindex].Value = headvalue;
                    }
                }
            }
            else
            {
                MessageBox.Show("設定參數不同!");
            }
        }
        private void UltraGrid1_InitializeRow(object sender, InitializeRowEventArgs e)
        {
            try
            {
                string uid = e.Row.Cells["標的代號"].Value.ToString();
                string tip = "";
                DataRow[] drs = dt_CMnoey.Select($@"UID = '{uid}'");
                if (drs.Length > 0)
                {
                    bool result = false;
                    string disposedatestr = drs[0][1].ToString();
                    int lastwatch = Convert.ToInt32(drs[0][2].ToString());
                    int watch = Convert.ToInt32(drs[0][3].ToString());
                    int lastwarning = Convert.ToInt32(drs[0][4].ToString());
                    int warning = Convert.ToInt32(drs[0][5].ToString());
                    if (disposedatestr != "19101231")
                    {
                        DateTime disposedate = DateTime.ParseExact(disposedatestr, "yyyyMMdd", null);
                        DateTime next3month = disposedate.AddMonths(3);
                        while (!EDLib.TradeDate.IsTradeDay(next3month))
                            next3month = next3month.AddDays(1);
                        if (next3month == DateTime.Today)
                        {
                            tip += "超過處置結束日3個月\n";
                            result = result | true;
                        }
                    }
                    if ((lastwatch == 2 && watch == 1)) 
                    {
                        tip += "6個營業日注意次數由(2->1)\n";
                        result = result | true;
                    }
                    if ((lastwatch == 0 && watch == 1)) 
                    {
                        tip += "6個營業日注意次數由(0->1)\n";
                        result = result | true;
                    }
                    if (lastwarning > 0 && warning == 0)
                    {
                        tip += "警示分數由非0變0\n";
                        result = result | true;
                    }
                    if (result)
                    {
                        e.Row.Cells["今日解禁可發行"].Value = "1";
                        e.Row.Cells["今日解禁可發行"].ToolTipText = tip;
                        e.Row.Cells["今日解禁可發行"].Appearance.BackColor = Color.IndianRed;
                    }
                }
            }
            catch(Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

        }
    }
}

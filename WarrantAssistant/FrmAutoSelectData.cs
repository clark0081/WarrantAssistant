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
using Infragistics.Win.UltraWinGrid;
using Microsoft.Office.Interop.Excel;

namespace WarrantAssistant
{
    public partial class FrmAutoSelectData : Form
    {
        DateTime today = DateTime.Today;
        public string userID = GlobalVar.globalParameter.userID;
        private System.Data.DataTable dt = new System.Data.DataTable();
        private System.Data.DataTable UnderlyingLiquidity = new System.Data.DataTable();//標的分級大中小
        private System.Data.DataTable UnderlyingRank = new System.Data.DataTable();//標的分級ABCD，以Call為主
        public FrmAutoSelectData()
        {
            InitializeComponent();
        }

        private void FrmAutoSelectData_Load(object sender, EventArgs e)
        {
            InitialGrid();
            this.comboBox2.Items.Add("C");
            this.comboBox2.Items.Add("P");
            this.comboBox2.SelectedItem = "C";
            string sql = $@"SELECT  CASE WHEN([UID] ='TWA00') THEN 'IX0001' ELSE [UID] END AS [UID], [Class]
                          FROM [newEDIS].[dbo].[UnderlyingRank]
                          WHERE [TDate] = (SELECT Max(TDate) FROM [newEDIS].[dbo].[UnderlyingRank]) AND [CallPut] ='C'";
            UnderlyingRank = MSSQL.ExecSqlQry(sql, "SERVER=10.60.0.37;DATABASE=newEDIS;UID=WarrantWeb;PWD=WarrantWeb");

            sql = $@"SELECT  [UID], [UnderlyingSize]
                          FROM [newEDIS].[dbo].[UnderlyingLiquidity]
                          WHERE [TDate] = (SELECT Max(TDate) FROM [newEDIS].[dbo].[UnderlyingLiquidity]) ";
            UnderlyingLiquidity = MSSQL.ExecSqlQry(sql, "SERVER=10.60.0.37;DATABASE=newEDIS;UID=WarrantWeb;PWD=WarrantWeb");
            LoadData();
            foreach (var item in GlobalVar.globalParameter.traders)
                comboBox1.Items.Add(item);
            comboBox1.Items.Add("All");
            comboBox1.Text = userID;
        }
        private void InitialGrid()
        {
            dt.Columns.Add("標的代號", typeof(string));
            dt.Columns.Add("標的名稱", typeof(string));
            dt.Columns.Add("標的分級", typeof(string));
            dt.Columns.Add("一個月BrokerPL(%)", typeof(float));
            //dt.Columns.Add("標的月損益(仟)", typeof(float));
            dt.Columns.Add("市場剩餘額度(檔)", typeof(int));
            dt.Columns.Add("前日市場剩餘額度(檔)", typeof(int));
            dt.Columns.Add("今日市場釋出額度(檔)", typeof(string));
            dt.Columns.Add("3日漲跌幅(%)", typeof(float));
            //dt.Columns.Add("3日跌幅(%)", typeof(float));
            dt.Columns.Add("市場Theta IV(仟)", typeof(float));
            dt.Columns.Add("市場部位近五日變化(萬)", typeof(float));
            //dt.Columns.Add("市場Theta IV 金額週變化(部位)(仟)", typeof(float));
            dt.Columns.Add("市場 Med_Vol / HV_60D", typeof(float));
            dt.Columns.Add("市場超額利潤(仟)", typeof(float));
            dt.Columns.Add("平均每檔權證超額利潤(元)", typeof(float));
            dt.Columns.Add("Theta天數", typeof(float));
            dt.Columns.Add("評鑑權證比重排名", typeof(float));
            dt.Columns.Add("融資使用率(%)", typeof(float));
            dt.Columns.Add("C有效檔數市佔(%)", typeof(float));
            dt.Columns.Add("P有效檔數市佔(%)", typeof(float));
            dt.Columns.Add("自家C發行密度", typeof(string));
            dt.Columns.Add("自家P發行密度", typeof(string));
            dt.Columns.Add("元大C發行密度", typeof(string));
            dt.Columns.Add("元大P發行密度", typeof(string));
            dt.Columns.Add("凱基-元大C發行密度", typeof(string));
            dt.Columns.Add("凱基-元大P發行密度", typeof(string));


            this.ultraGrid1.DisplayLayout.Override.DefaultRowHeight = 30;
            this.ultraGrid1.DisplayLayout.AutoFitStyle = Infragistics.Win.UltraWinGrid.AutoFitStyle.ResizeAllColumns;
            this.ultraGrid1.DisplayLayout.Override.WrapHeaderText = Infragistics.Win.DefaultableBoolean.True;
            this.ultraGrid1.DisplayLayout.Override.CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center;
            this.ultraGrid1.DisplayLayout.Override.CellAppearance.TextVAlign = Infragistics.Win.VAlign.Middle;

            this.ultraGrid1.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti;
            ultraGrid1.DataSource = dt;
            SetButton();
        }

        private void LoadData()
        {
            dt.Clear();
#if !To39
            string sql = $@"SELECT  [UID], [UName], [Trader], [BrokerPL_Month]
                          , [Profit_Month], [OptionAvailable], [LastOptionAvailable], [DiffOptionAvailable], [RiseUp_3Days]
                          , [ThetaIV_WeekDelta], [Med_HV60D_VolRatio], [Theta_Days]
                          , [AppraisalRank], [FinancingRatio], [CallMarketShare], [PutMarketShare]
                          , [CallDensity], [PutDensity], [Theta_EndDate]
                          , [ThetaIV_WeekDelta_P], [Med_HV60D_VolRatio_P], [Theta_Days_P], [AppraisalRank_P], [Theta_EndDate_P]
                          , [YuanCallDensity], [YuanPutDensity], [KgiYuanCallDensity], [KgiYuanPutDensity]
                          , [MarketAmtChg5Days], [MarketAmtChg5Days_P], [AlphaTheta], [AlphaTheta_P], [AvgAlphaThetaCost], [AvgAlphaThetaCost_P]
                            FROM [newEDIS].[dbo].[OptionAutoSelectData]
                           WHERE [Trader] = '{userID.TrimStart('0')}'
                          ORDER BY [UID]";
            System.Data.DataTable dttemp = MSSQL.ExecSqlQry(sql, GlobalVar.loginSet.newEDIS);
#else
            string sql = $@"SELECT  [UID], [UName], [Trader], [BrokerPL_Month]
                          , [Profit_Month], [OptionAvailable], [LastOptionAvailable], [DiffOptionAvailable], [RiseUp_3Days]
                          , [ThetaIV_WeekDelta], [Med_HV60D_VolRatio], [Theta_Days]
                          , [AppraisalRank], [FinancingRatio], [CallMarketShare], [PutMarketShare]
                          , [CallDensity], [PutDensity], [Theta_EndDate]
                          , [ThetaIV_WeekDelta_P], [Med_HV60D_VolRatio_P], [Theta_Days_P], [AppraisalRank_P], [Theta_EndDate_P]
                          , [YuanCallDensity], [YuanPutDensity], [KgiYuanCallDensity], [KgiYuanPutDensity]
                          , [MarketAmtChg5Days], [MarketAmtChg5Days_P], [AlphaTheta], [AlphaTheta_P], [AvgAlphaThetaCost], [AvgAlphaThetaCost_P]
                            FROM [WarrantAssistant].[dbo].[OptionAutoSelectData]
                           WHERE [Trader] = '{userID.TrimStart('0')}'
                          ORDER BY [UID]";
            System.Data.DataTable dttemp = MSSQL.ExecSqlQry(sql, GlobalVar.loginSet.warrantassistant45);
#endif
            string CP = comboBox2.Text;
            foreach (DataRow dr in dttemp.Rows)
            {
                DataRow drv = dt.NewRow();

                string uid = dr["UID"].ToString();
                drv["標的代號"] = uid;

                string uname = dr["UName"].ToString();
                drv["標的名稱"] = uname;

                string trader = dr["Trader"].ToString();

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

                string brokerPL_month = dr["BrokerPL_Month"].ToString();
                drv["一個月BrokerPL(%)"] = Convert.ToDouble(brokerPL_month);
                /*
                string profit_month = dr["Profit_Month"].ToString();
                drv["標的月損益(仟)"] = Convert.ToDouble(profit_month);
                */
                string optionavailable = dr["OptionAvailable"].ToString();
                drv["市場剩餘額度(檔)"] = Convert.ToInt32(optionavailable);

                string lastoptionavailable = dr["LastOptionAvailable"].ToString();
                drv["前日市場剩餘額度(檔)"] = Convert.ToInt32(lastoptionavailable);

                
                string diffoptionavailable = dr["DiffOptionAvailable"].ToString();
                //drv["今日市場釋出額度(檔)"] = Convert.ToInt32(diffoptionavailable);
                drv["今日市場釋出額度(檔)"] = (diffoptionavailable);

                string riseup_3days = dr["RiseUp_3Days"].ToString();
                drv["3日漲跌幅(%)"] = Convert.ToDouble(riseup_3days);

                /*
                string dropdown_3days = dr["DropDown_3Days"].ToString();
                drv["3日跌幅(%)"] = Convert.ToDouble(dropdown_3days);
                */

                string theta_enddate = "";
                if (CP == "C")
                {
                    theta_enddate = dr["Theta_EndDate"].ToString();
                }
                else
                {
                    theta_enddate = dr["Theta_EndDate_P"].ToString();
                }
                drv["市場Theta IV(仟)"] = Math.Round(Convert.ToDouble(theta_enddate));

                string marketamtchg = "";
                if (CP == "C")
                {
                    marketamtchg = dr["MarketAmtChg5Days"].ToString();
                }
                else
                {
                    marketamtchg = dr["MarketAmtChg5Days_P"].ToString();
                }
                drv["市場部位近五日變化(萬)"] = Math.Round(Convert.ToDouble(marketamtchg));
                /*
                string thetaiv_weekdelta = "";
                
                if (CP == "C")
                {
                    thetaiv_weekdelta = dr["ThetaIV_WeekDelta"].ToString();
                }
                else
                {
                    thetaiv_weekdelta = dr["ThetaIV_WeekDelta_P"].ToString();
                }
                drv["市場Theta IV 金額週變化(部位)(仟)"] = Math.Round(Convert.ToDouble(thetaiv_weekdelta));
                */
                string med_hv60d_volratio = "";
                if (CP == "C")
                {
                    med_hv60d_volratio = dr["Med_HV60D_VolRatio"].ToString();
                }
                else
                {
                    med_hv60d_volratio = dr["Med_HV60D_VolRatio_P"].ToString();
                }
                drv["市場 Med_Vol / HV_60D"] = Convert.ToDouble(med_hv60d_volratio);

                string alphatheta = "";
                if (CP == "C")
                {
                    alphatheta = dr["AlphaTheta"].ToString();
                }
                else
                {
                    alphatheta = dr["AlphaTheta_P"].ToString();
                }
                drv["市場超額利潤(仟)"] = Convert.ToDouble(alphatheta);

                string avgalphathetacost = "";
                if (CP == "C")
                {
                    avgalphathetacost = dr["AvgAlphaThetaCost"].ToString();
                }
                else
                {
                    avgalphathetacost = dr["AvgAlphaThetaCost_P"].ToString();
                }
                drv["平均每檔權證超額利潤(元)"] = Convert.ToDouble(avgalphathetacost);
                string theta_days = "";
                if (CP == "C")
                {
                    theta_days = dr["Theta_Days"].ToString();
                }
                else
                {
                    theta_days = dr["Theta_Days_P"].ToString();
                }
                drv["Theta天數"] = Math.Round(Convert.ToDouble(theta_days));

                string appraisalrank = "";
                if (CP == "C")
                {
                    appraisalrank = dr["AppraisalRank"].ToString();
                }
                else
                {
                    appraisalrank = dr["AppraisalRank_P"].ToString();
                }
                drv["評鑑權證比重排名"] = Math.Round(Convert.ToDouble(appraisalrank));

                string financingratio = dr["FinancingRatio"].ToString();
                drv["融資使用率(%)"] = Convert.ToDouble(financingratio);

                string callmarketshare = dr["CallMarketShare"].ToString();
                drv["C有效檔數市佔(%)"] = Convert.ToDouble(callmarketshare);

                string putmarketshare = dr["PutMarketShare"].ToString();
                drv["P有效檔數市佔(%)"] = Convert.ToDouble(putmarketshare);

                string calldensity = dr["CallDensity"].ToString();
                
                if (Convert.ToDouble(calldensity) < 0)
                    drv["自家C發行密度"] = "";
                else
                    drv["自家C發行密度"] = calldensity;

                string putdensity = dr["PutDensity"].ToString();
                
                if (Convert.ToDouble(putdensity) < 0)
                    drv["自家P發行密度"] = "";
                else
                    drv["自家P發行密度"] = putdensity;

                string yuancalldensity = dr["YuanCallDensity"].ToString();

                if (Convert.ToDouble(yuancalldensity) < 0)
                    drv["元大C發行密度"] = "";
                else
                    drv["元大C發行密度"] = yuancalldensity;

                string yuanputdensity = dr["YuanPutDensity"].ToString();

                if (Convert.ToDouble(yuanputdensity) < 0)
                    drv["元大P發行密度"] = "";
                else
                    drv["元大P發行密度"] = yuanputdensity;

                string kgiyuancalldensity = dr["KgiYuanCallDensity"].ToString();

                if (Convert.ToDouble(kgiyuancalldensity) < 0)
                    drv["凱基-元大C發行密度"] = "";
                else
                    drv["凱基-元大C發行密度"] = kgiyuancalldensity;

                string kgiyuanputdensity = dr["KgiYuanPutDensity"].ToString();

                if (Convert.ToDouble(kgiyuanputdensity) < 0)
                    drv["凱基-元大P發行密度"] = "";
                else
                    drv["凱基-元大P發行密度"] = kgiyuanputdensity;
                dt.Rows.Add(drv);
            }

        }
        private void LoadDataByUID()
        {
            string str = textBox1.Text;
            string traderID = comboBox1.Text;
            string CP = comboBox2.Text;
            
            
            string sql = "";
            if (str == "")
            {
#if !To39
                if(traderID=="All")
                    sql = $@"SELECT  [UID], [UName], [Trader], [BrokerPL_Month]
                              , [Profit_Month], [OptionAvailable], [LastOptionAvailable], [DiffOptionAvailable], [RiseUp_3Days]
                              , [ThetaIV_WeekDelta], [Med_HV60D_VolRatio], [Theta_Days]
                              , [AppraisalRank], [FinancingRatio], [CallMarketShare], [PutMarketShare]
                              , [CallDensity], [PutDensity], [Theta_EndDate]
                              , [ThetaIV_WeekDelta_P], [Med_HV60D_VolRatio_P], [Theta_Days_P], [AppraisalRank_P], [Theta_EndDate_P]
                              , [YuanCallDensity], [YuanPutDensity], [KgiYuanCallDensity], [KgiYuanPutDensity]
                              , [MarketAmtChg5Days], [MarketAmtChg5Days_P], [AlphaTheta], [AlphaTheta_P], [AvgAlphaThetaCost], [AvgAlphaThetaCost_P]
                                FROM [newEDIS].[dbo].[OptionAutoSelectData]
                                ORDER BY [UID]";
                else
                    sql = $@"SELECT  [UID], [UName], [Trader], [BrokerPL_Month]
                              , [Profit_Month], [OptionAvailable], [LastOptionAvailable], [DiffOptionAvailable], [RiseUp_3Days]
                              , [ThetaIV_WeekDelta], [Med_HV60D_VolRatio], [Theta_Days]
                              , [AppraisalRank], [FinancingRatio], [CallMarketShare], [PutMarketShare]
                              , [CallDensity], [PutDensity], [Theta_EndDate]
                              , [ThetaIV_WeekDelta_P], [Med_HV60D_VolRatio_P], [Theta_Days_P], [AppraisalRank_P], [Theta_EndDate_P]
                              , [YuanCallDensity], [YuanPutDensity], [KgiYuanCallDensity], [KgiYuanPutDensity]
                              , [MarketAmtChg5Days], [MarketAmtChg5Days_P], [AlphaTheta], [AlphaTheta_P], [AvgAlphaThetaCost], [AvgAlphaThetaCost_P]
                                FROM [newEDIS].[dbo].[OptionAutoSelectData]
                               WHERE [Trader] = '{traderID.TrimStart('0')}'
                                ORDER BY [UID]";
#else
                if (traderID == "All")
                    sql = $@"SELECT  [UID], [UName], [Trader], [BrokerPL_Month]
                              , [Profit_Month], [OptionAvailable], [LastOptionAvailable], [DiffOptionAvailable], [RiseUp_3Days]
                              , [ThetaIV_WeekDelta], [Med_HV60D_VolRatio], [Theta_Days]
                              , [AppraisalRank], [FinancingRatio], [CallMarketShare], [PutMarketShare]
                              , [CallDensity], [PutDensity], [Theta_EndDate]
                              , [ThetaIV_WeekDelta_P], [Med_HV60D_VolRatio_P], [Theta_Days_P], [AppraisalRank_P], [Theta_EndDate_P]
                              , [YuanCallDensity], [YuanPutDensity], [KgiYuanCallDensity], [KgiYuanPutDensity]
                              , [MarketAmtChg5Days], [MarketAmtChg5Days_P], [AlphaTheta], [AlphaTheta_P], [AvgAlphaThetaCost], [AvgAlphaThetaCost_P]
                                FROM [WarrantAssistant].[dbo].[OptionAutoSelectData]
                                ORDER BY [UID]";
                else
                    sql = $@"SELECT  [UID], [UName], [Trader], [BrokerPL_Month]
                              , [Profit_Month], [OptionAvailable], [LastOptionAvailable], [DiffOptionAvailable], [RiseUp_3Days]
                              , [ThetaIV_WeekDelta], [Med_HV60D_VolRatio], [Theta_Days]
                              , [AppraisalRank], [FinancingRatio], [CallMarketShare], [PutMarketShare]
                              , [CallDensity], [PutDensity], [Theta_EndDate]
                              , [ThetaIV_WeekDelta_P], [Med_HV60D_VolRatio_P], [Theta_Days_P], [AppraisalRank_P], [Theta_EndDate_P]
                              , [YuanCallDensity], [YuanPutDensity], [KgiYuanCallDensity], [KgiYuanPutDensity]
                              , [MarketAmtChg5Days], [MarketAmtChg5Days_P], [AlphaTheta], [AlphaTheta_P], [AvgAlphaThetaCost], [AvgAlphaThetaCost_P]
                                FROM [WarrantAssistant].[dbo].[OptionAutoSelectData]
                               WHERE [Trader] = '{traderID.TrimStart('0')}'
                                ORDER BY [UID]";
#endif
            }
            else
            {
#if !To39
                if (traderID == "All")
                    sql = $@"SELECT  [UID], [UName], [Trader], [BrokerPL_Month]
                          , [Profit_Month], [OptionAvailable], [LastOptionAvailable], [DiffOptionAvailable], [RiseUp_3Days]
                          , [ThetaIV_WeekDelta], [Med_HV60D_VolRatio], [Theta_Days]
                          , [AppraisalRank], [FinancingRatio], [CallMarketShare], [PutMarketShare]
                          , [CallDensity], [PutDensity], [Theta_EndDate]
                          , [ThetaIV_WeekDelta_P], [Med_HV60D_VolRatio_P], [Theta_Days_P], [AppraisalRank_P], [Theta_EndDate_P]
                          , [YuanCallDensity], [YuanPutDensity], [KgiYuanCallDensity], [KgiYuanPutDensity]
                          , [MarketAmtChg5Days], [MarketAmtChg5Days_P], [AlphaTheta], [AlphaTheta_P], [AvgAlphaThetaCost], [AvgAlphaThetaCost_P]
                            FROM [newEDIS].[dbo].[OptionAutoSelectData]
                           WHERE [UID] ='{str}' ORDER BY [UID]";
                else
                    sql = $@"SELECT  [UID], [UName], [Trader], [BrokerPL_Month]
                          , [Profit_Month], [OptionAvailable], [LastOptionAvailable], [DiffOptionAvailable], [RiseUp_3Days]
                          , [ThetaIV_WeekDelta], [Med_HV60D_VolRatio], [Theta_Days]
                          , [AppraisalRank], [FinancingRatio], [CallMarketShare], [PutMarketShare]
                          , [CallDensity], [PutDensity], [Theta_EndDate]
                          , [ThetaIV_WeekDelta_P], [Med_HV60D_VolRatio_P], [Theta_Days_P], [AppraisalRank_P], [Theta_EndDate_P]
                          , [YuanCallDensity], [YuanPutDensity], [KgiYuanCallDensity], [KgiYuanPutDensity]
                          , [MarketAmtChg5Days], [MarketAmtChg5Days_P], [AlphaTheta], [AlphaTheta_P], [AvgAlphaThetaCost], [AvgAlphaThetaCost_P]
                            FROM [newEDIS].[dbo].[OptionAutoSelectData]
                           WHERE [Trader] = '{traderID.TrimStart('0')}' AND [UID] ='{str}'
                            ORDER BY [UID]";
#else
                if (traderID == "All")
                    sql = $@"SELECT  [UID], [UName], [Trader], [BrokerPL_Month]
                          , [Profit_Month], [OptionAvailable], [LastOptionAvailable], [DiffOptionAvailable], [RiseUp_3Days]
                          , [ThetaIV_WeekDelta], [Med_HV60D_VolRatio], [Theta_Days]
                          , [AppraisalRank], [FinancingRatio], [CallMarketShare], [PutMarketShare]
                          , [CallDensity], [PutDensity], [Theta_EndDate]
                          , [ThetaIV_WeekDelta_P], [Med_HV60D_VolRatio_P], [Theta_Days_P], [AppraisalRank_P], [Theta_EndDate_P]
                          , [YuanCallDensity], [YuanPutDensity], [KgiYuanCallDensity], [KgiYuanPutDensity]
                          , [MarketAmtChg5Days], [MarketAmtChg5Days_P], [AlphaTheta], [AlphaTheta_P], [AvgAlphaThetaCost], [AvgAlphaThetaCost_P]
                            FROM [WarrantAssistant].[dbo].[OptionAutoSelectData]
                           WHERE [UID] ='{str}' ORDER BY [UID]";
                else
                    sql = $@"SELECT  [UID], [UName], [Trader], [BrokerPL_Month]
                          , [Profit_Month], [OptionAvailable], [LastOptionAvailable], [DiffOptionAvailable], [RiseUp_3Days]
                          , [ThetaIV_WeekDelta], [Med_HV60D_VolRatio], [Theta_Days]
                          , [AppraisalRank], [FinancingRatio], [CallMarketShare], [PutMarketShare]
                          , [CallDensity], [PutDensity], [Theta_EndDate]
                          , [ThetaIV_WeekDelta_P], [Med_HV60D_VolRatio_P], [Theta_Days_P], [AppraisalRank_P], [Theta_EndDate_P]
                          , [YuanCallDensity], [YuanPutDensity], [KgiYuanCallDensity], [KgiYuanPutDensity]
                          , [MarketAmtChg5Days], [MarketAmtChg5Days_P], [AlphaTheta], [AlphaTheta_P], [AvgAlphaThetaCost], [AvgAlphaThetaCost_P]
                            FROM [WarrantAssistant].[dbo].[OptionAutoSelectData]
                           WHERE [Trader] = '{traderID.TrimStart('0')}' AND [UID] ='{str}'
                            ORDER BY [UID]";
#endif
            }
            dt.Clear();
#if !To39
            System.Data.DataTable dttemp = MSSQL.ExecSqlQry(sql, GlobalVar.loginSet.newEDIS);
#else
            System.Data.DataTable dttemp = MSSQL.ExecSqlQry(sql, GlobalVar.loginSet.warrantassistant45);
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


                string brokerPL_month = dr["BrokerPL_Month"].ToString();
                drv["一個月BrokerPL(%)"] = Convert.ToDouble(brokerPL_month);
                /*
                string profit_month = dr["Profit_Month"].ToString();
                drv["標的月損益(仟)"] = Convert.ToDouble(profit_month);
                */
                string optionavailable = dr["OptionAvailable"].ToString();
                drv["市場剩餘額度(檔)"] = Convert.ToInt32(optionavailable);

                string lastoptionavailable = dr["LastOptionAvailable"].ToString();
                drv["前日市場剩餘額度(檔)"] = Convert.ToInt32(lastoptionavailable);

                string diffoptionavailable = dr["DiffOptionAvailable"].ToString();
                //drv["今日市場釋出額度(檔)"] = Convert.ToInt32(diffoptionavailable);
                drv["今日市場釋出額度(檔)"] = (diffoptionavailable);

                string riseup_3days = dr["RiseUp_3Days"].ToString();
                drv["3日漲跌幅(%)"] = Convert.ToDouble(riseup_3days);

                /*
                string dropdown_3days = dr["DropDown_3Days"].ToString();
                drv["3日跌幅(%)"] = Convert.ToDouble(dropdown_3days);
                */
                /*
                string theta_enddate = dr["Theta_EndDate"].ToString();
                drv["市場Theta IV(仟)"] = Math.Round(Convert.ToDouble(theta_enddate));

                string thetaiv_weekdelta = dr["ThetaIV_WeekDelta"].ToString();
                drv["市場Theta IV 金額週變化(部位)(仟)"] = Math.Round(Convert.ToDouble(thetaiv_weekdelta));

                string med_hv60d_volratio = dr["Med_HV60D_VolRatio"].ToString();
                drv["市場 Med_Vol / HV_60D"] = Convert.ToDouble(med_hv60d_volratio);

                string theta_days = dr["Theta_Days"].ToString();
                drv["Theta天數"] = Math.Round(Convert.ToDouble(theta_days));

                string optweightinprice1to2dot5 = dr["OptWeightInPrice1To2dot5"].ToString();
                drv["庫存權證價格在1~2.5元的比重"] = Math.Round(Convert.ToDouble(optweightinprice1to2dot5));
                */
                string theta_enddate = "";
                if (CP == "C")
                {
                    theta_enddate = dr["Theta_EndDate"].ToString();
                }
                else
                {
                    theta_enddate = dr["Theta_EndDate_P"].ToString();
                }
                drv["市場Theta IV(仟)"] = Math.Round(Convert.ToDouble(theta_enddate));

                string marketamtchg = "";
                if (CP == "C")
                {
                    marketamtchg = dr["MarketAmtChg5Days"].ToString();
                }
                else
                {
                    marketamtchg = dr["MarketAmtChg5Days_P"].ToString();
                }
                drv["市場部位近五日變化(萬)"] = Math.Round(Convert.ToDouble(marketamtchg));
                /*
                string thetaiv_weekdelta = "";
                if (CP == "C")
                {
                    thetaiv_weekdelta = dr["ThetaIV_WeekDelta"].ToString();
                }
                else
                {
                    thetaiv_weekdelta = dr["ThetaIV_WeekDelta_P"].ToString();
                }
                drv["市場Theta IV 金額週變化(部位)(仟)"] = Math.Round(Convert.ToDouble(thetaiv_weekdelta));
                */
                string med_hv60d_volratio = "";
                if (CP == "C")
                {
                    med_hv60d_volratio = dr["Med_HV60D_VolRatio"].ToString();
                }
                else
                {
                    med_hv60d_volratio = dr["Med_HV60D_VolRatio_P"].ToString();
                }
                drv["市場 Med_Vol / HV_60D"] = Convert.ToDouble(med_hv60d_volratio);


                string alphatheta = "";
                if (CP == "C")
                {
                    alphatheta = dr["AlphaTheta"].ToString();
                }
                else
                {
                    alphatheta = dr["AlphaTheta_P"].ToString();
                }
                drv["市場超額利潤(仟)"] = Convert.ToDouble(alphatheta);

                string avgalphathetacost = "";
                if (CP == "C")
                {
                    avgalphathetacost = dr["AvgAlphaThetaCost"].ToString();
                }
                else
                {
                    avgalphathetacost = dr["AvgAlphaThetaCost_P"].ToString();
                }
                drv["平均每檔權證超額利潤(元)"] = Convert.ToDouble(avgalphathetacost);
                string theta_days = "";
                if (CP == "C")
                {
                    theta_days = dr["Theta_Days"].ToString();
                }
                else
                {
                    theta_days = dr["Theta_Days_P"].ToString();
                }
                drv["Theta天數"] = Math.Round(Convert.ToDouble(theta_days));

                string appraisalrank = "";
                if (CP == "C")
                {
                    appraisalrank = dr["AppraisalRank"].ToString();
                }
                else
                {
                    appraisalrank = dr["AppraisalRank_P"].ToString();
                }
                drv["評鑑權證比重排名"] = Math.Round(Convert.ToDouble(appraisalrank));
                string financingratio = dr["FinancingRatio"].ToString();
                drv["融資使用率(%)"] = Convert.ToDouble(financingratio);

                string callmarketshare = dr["CallMarketShare"].ToString();
                drv["C有效檔數市佔(%)"] = Convert.ToDouble(callmarketshare);

                string putmarketshare = dr["PutMarketShare"].ToString();
                drv["P有效檔數市佔(%)"] = Convert.ToDouble(putmarketshare);

                string calldensity = dr["CallDensity"].ToString();
                
                if (Convert.ToDouble(calldensity) < 0)
                    drv["自家C發行密度"] = "";
                else
                    drv["自家C發行密度"] = calldensity;
                

                string putdensity = dr["PutDensity"].ToString();
                
                if (Convert.ToDouble(putdensity) < 0)
                    drv["自家P發行密度"] = "";
                else
                    drv["自家P發行密度"] = putdensity;

                string yuancalldensity = dr["YuanCallDensity"].ToString();

                if (Convert.ToDouble(yuancalldensity) < 0)
                    drv["元大C發行密度"] = "";
                else
                    drv["元大C發行密度"] = yuancalldensity;

                string yuanputdensity = dr["YuanPutDensity"].ToString();

                if (Convert.ToDouble(yuanputdensity) < 0)
                    drv["元大P發行密度"] = "";
                else
                    drv["元大P發行密度"] = yuanputdensity;

                string kgiyuancalldensity = dr["KgiYuanCallDensity"].ToString();

                if (Convert.ToDouble(kgiyuancalldensity) < 0)
                    drv["凱基-元大C發行密度"] = "";
                else
                    drv["凱基-元大C發行密度"] = kgiyuancalldensity;

                string kgiyuanputdensity = dr["KgiYuanPutDensity"].ToString();

                if (Convert.ToDouble(kgiyuanputdensity) < 0)
                    drv["凱基-元大P發行密度"] = "";
                else
                    drv["凱基-元大P發行密度"] = kgiyuanputdensity;
                dt.Rows.Add(drv);
                
            }

        }
        private void LoadDataByTrader()
        {
            string str = textBox1.Text;
            string traderID = comboBox1.Text;
            string sql = "";
            string CP = comboBox2.Text;
#if !To39
            if (str == "")
            {
                if (traderID == "All")
                    sql = $@"SELECT  [UID], [UName], [Trader], [BrokerPL_Month]
                              , [Profit_Month], [OptionAvailable], [LastOptionAvailable], [DiffOptionAvailable], [RiseUp_3Days]
                              , [ThetaIV_WeekDelta], [Med_HV60D_VolRatio], [Theta_Days]
                              , [AppraisalRank], [FinancingRatio], [CallMarketShare], [PutMarketShare]
                              , [CallDensity], [PutDensity], [Theta_EndDate]
                              , [ThetaIV_WeekDelta_P], [Med_HV60D_VolRatio_P], [Theta_Days_P], [AppraisalRank_P], [Theta_EndDate_P]
                              , [YuanCallDensity], [YuanPutDensity], [KgiYuanCallDensity], [KgiYuanPutDensity]
                              , [MarketAmtChg5Days], [MarketAmtChg5Days_P], [AlphaTheta], [AlphaTheta_P], [AvgAlphaThetaCost], [AvgAlphaThetaCost_P]
                                FROM [newEDIS].[dbo].[OptionAutoSelectData] ORDER BY [UID]";
                else
                    sql = $@"SELECT  [UID], [UName], [Trader], [BrokerPL_Month]
                              , [Profit_Month], [OptionAvailable], [LastOptionAvailable], [DiffOptionAvailable], [RiseUp_3Days]
                              , [ThetaIV_WeekDelta], [Med_HV60D_VolRatio], [Theta_Days]
                              , [AppraisalRank], [FinancingRatio], [CallMarketShare], [PutMarketShare]
                              , [CallDensity], [PutDensity], [Theta_EndDate]
                              , [ThetaIV_WeekDelta_P], [Med_HV60D_VolRatio_P], [Theta_Days_P], [AppraisalRank_P], [Theta_EndDate_P]
                              , [YuanCallDensity], [YuanPutDensity], [KgiYuanCallDensity], [KgiYuanPutDensity]
                              , [MarketAmtChg5Days], [MarketAmtChg5Days_P], [AlphaTheta], [AlphaTheta_P], [AvgAlphaThetaCost], [AvgAlphaThetaCost_P]
                                FROM [newEDIS].[dbo].[OptionAutoSelectData]
                               WHERE [Trader] = '{traderID.TrimStart('0')}'
                                ORDER BY [UID]";
            }
            else
            {
                if (traderID == "All")
                    sql = $@"SELECT  [UID], [UName], [Trader], [BrokerPL_Month]
                          , [Profit_Month], [OptionAvailable], [LastOptionAvailable], [DiffOptionAvailable], [RiseUp_3Days]
                          , [ThetaIV_WeekDelta], [Med_HV60D_VolRatio], [Theta_Days]
                          , [AppraisalRank], [FinancingRatio], [CallMarketShare], [PutMarketShare]
                          , [CallDensity], [PutDensity], [Theta_EndDate]
                          , [ThetaIV_WeekDelta_P], [Med_HV60D_VolRatio_P], [Theta_Days_P], [AppraisalRank_P], [Theta_EndDate_P]
                          , [YuanCallDensity], [YuanPutDensity], [KgiYuanCallDensity], [KgiYuanPutDensity]
                          , [MarketAmtChg5Days], [MarketAmtChg5Days_P], [AlphaTheta], [AlphaTheta_P], [AvgAlphaThetaCost], [AvgAlphaThetaCost_P]
                            FROM [newEDIS].[dbo].[OptionAutoSelectData]
                           WHERE [UID] ='{str}' ORDER BY [UID]";
                else
                    sql = $@"SELECT  [UID], [UName], [Trader], [BrokerPL_Month]
                          , [Profit_Month], [OptionAvailable], [LastOptionAvailable], [DiffOptionAvailable], [RiseUp_3Days]
                          , [ThetaIV_WeekDelta], [Med_HV60D_VolRatio], [Theta_Days]
                          , [AppraisalRank], [FinancingRatio], [CallMarketShare], [PutMarketShare]
                          , [CallDensity], [PutDensity], [Theta_EndDate]
                          , [ThetaIV_WeekDelta_P], [Med_HV60D_VolRatio_P], [Theta_Days_P], [AppraisalRank_P], [Theta_EndDate_P]
                          , [YuanCallDensity], [YuanPutDensity], [KgiYuanCallDensity], [KgiYuanPutDensity]
                          , [MarketAmtChg5Days], [MarketAmtChg5Days_P], [AlphaTheta], [AlphaTheta_P], [AvgAlphaThetaCost], [AvgAlphaThetaCost_P]
                            FROM [newEDIS].[dbo].[OptionAutoSelectData]
                           WHERE [Trader] = '{traderID.TrimStart('0')}' AND [UID] ='{str}'
                            ORDER BY [UID]";
            }
#else
            if (str == "")
            {
                if (traderID == "All")
                    sql = $@"SELECT  [UID], [UName], [Trader], [BrokerPL_Month]
                              , [Profit_Month], [OptionAvailable], [LastOptionAvailable], [DiffOptionAvailable], [RiseUp_3Days]
                              , [ThetaIV_WeekDelta], [Med_HV60D_VolRatio], [Theta_Days]
                              , [AppraisalRank], [FinancingRatio], [CallMarketShare], [PutMarketShare]
                              , [CallDensity], [PutDensity], [Theta_EndDate]
                              , [ThetaIV_WeekDelta_P], [Med_HV60D_VolRatio_P], [Theta_Days_P], [AppraisalRank_P], [Theta_EndDate_P]
                              , [YuanCallDensity], [YuanPutDensity], [KgiYuanCallDensity], [KgiYuanPutDensity]
                              , [MarketAmtChg5Days], [MarketAmtChg5Days_P], [AlphaTheta], [AlphaTheta_P], [AvgAlphaThetaCost], [AvgAlphaThetaCost_P]
                                FROM [WarrantAssistant].[dbo].[OptionAutoSelectData] ORDER BY [UID]";
                else
                    sql = $@"SELECT  [UID], [UName], [Trader], [BrokerPL_Month]
                              , [Profit_Month], [OptionAvailable], [LastOptionAvailable], [DiffOptionAvailable], [RiseUp_3Days]
                              , [ThetaIV_WeekDelta], [Med_HV60D_VolRatio], [Theta_Days]
                              , [AppraisalRank], [FinancingRatio], [CallMarketShare], [PutMarketShare]
                              , [CallDensity], [PutDensity], [Theta_EndDate]
                              , [ThetaIV_WeekDelta_P], [Med_HV60D_VolRatio_P], [Theta_Days_P], [AppraisalRank_P], [Theta_EndDate_P]
                              , [YuanCallDensity], [YuanPutDensity], [KgiYuanCallDensity], [KgiYuanPutDensity]
                              , [MarketAmtChg5Days], [MarketAmtChg5Days_P], [AlphaTheta], [AlphaTheta_P], [AvgAlphaThetaCost], [AvgAlphaThetaCost_P]
                                FROM [WarrantAssistant].[dbo].[OptionAutoSelectData]
                               WHERE [Trader] = '{traderID.TrimStart('0')}'
                                ORDER BY [UID]";
            }
            else
            {
                if (traderID == "All")
                    sql = $@"SELECT  [UID], [UName], [Trader], [BrokerPL_Month]
                          , [Profit_Month], [OptionAvailable], [LastOptionAvailable], [DiffOptionAvailable], [RiseUp_3Days]
                          , [ThetaIV_WeekDelta], [Med_HV60D_VolRatio], [Theta_Days]
                          , [AppraisalRank], [FinancingRatio], [CallMarketShare], [PutMarketShare]
                          , [CallDensity], [PutDensity], [Theta_EndDate]
                          , [ThetaIV_WeekDelta_P], [Med_HV60D_VolRatio_P], [Theta_Days_P], [AppraisalRank_P], [Theta_EndDate_P]
                          , [YuanCallDensity], [YuanPutDensity], [KgiYuanCallDensity], [KgiYuanPutDensity]
                          , [MarketAmtChg5Days], [MarketAmtChg5Days_P], [AlphaTheta], [AlphaTheta_P], [AvgAlphaThetaCost], [AvgAlphaThetaCost_P]
                            FROM [WarrantAssistant].[dbo].[OptionAutoSelectData]
                           WHERE [UID] ='{str}' ORDER BY [UID]";
                else
                    sql = $@"SELECT  [UID], [UName], [Trader], [BrokerPL_Month]
                          , [Profit_Month], [OptionAvailable], [LastOptionAvailable], [DiffOptionAvailable], [RiseUp_3Days]
                          , [ThetaIV_WeekDelta], [Med_HV60D_VolRatio], [Theta_Days]
                          , [AppraisalRank], [FinancingRatio], [CallMarketShare], [PutMarketShare]
                          , [CallDensity], [PutDensity], [Theta_EndDate]
                          , [ThetaIV_WeekDelta_P], [Med_HV60D_VolRatio_P], [Theta_Days_P], [AppraisalRank_P], [Theta_EndDate_P]
                          , [YuanCallDensity], [YuanPutDensity], [KgiYuanCallDensity], [KgiYuanPutDensity]
                          , [MarketAmtChg5Days], [MarketAmtChg5Days_P], [AlphaTheta], [AlphaTheta_P], [AvgAlphaThetaCost], [AvgAlphaThetaCost_P]
                            FROM [WarrantAssistant].[dbo].[OptionAutoSelectData]
                           WHERE [Trader] = '{traderID.TrimStart('0')}' AND [UID] ='{str}'
                            ORDER BY [UID]";
            }
#endif
            dt.Clear();
#if !To39
            System.Data.DataTable dttemp = MSSQL.ExecSqlQry(sql, GlobalVar.loginSet.newEDIS);
#else
            System.Data.DataTable dttemp = MSSQL.ExecSqlQry(sql, GlobalVar.loginSet.warrantassistant45);
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


                string brokerPL_month = dr["BrokerPL_Month"].ToString();
                drv["一個月BrokerPL(%)"] = Convert.ToDouble(brokerPL_month);
                /*
                string profit_month = dr["Profit_Month"].ToString();
                drv["標的月損益(仟)"] = Convert.ToDouble(profit_month);
                */
                string optionavailable = dr["OptionAvailable"].ToString();
                drv["市場剩餘額度(檔)"] = Convert.ToInt32(optionavailable);

                string lastoptionavailable = dr["LastOptionAvailable"].ToString();
                drv["前日市場剩餘額度(檔)"] = Convert.ToInt32(lastoptionavailable);

                string diffoptionavailable = dr["DiffOptionAvailable"].ToString();
                //drv["今日市場釋出額度(檔)"] = Convert.ToInt32(diffoptionavailable);
                drv["今日市場釋出額度(檔)"] = (diffoptionavailable);

                string riseup_3days = dr["RiseUp_3Days"].ToString();
                drv["3日漲跌幅(%)"] = Convert.ToDouble(riseup_3days);

                string marketamtchg = "";
                if (CP == "C")
                {
                    marketamtchg = dr["MarketAmtChg5Days"].ToString();
                }
                else
                {
                    marketamtchg = dr["MarketAmtChg5Days_P"].ToString();
                }
                drv["市場部位近五日變化(萬)"] = Math.Round(Convert.ToDouble(marketamtchg));

                /*
                string dropdown_3days = dr["DropDown_3Days"].ToString();
                drv["3日跌幅(%)"] = Convert.ToDouble(dropdown_3days);
                */
                /*
                string theta_enddate = dr["Theta_EndDate"].ToString();
                drv["市場Theta IV(仟)"] = Math.Round(Convert.ToDouble(theta_enddate));

                string thetaiv_weekdelta = dr["ThetaIV_WeekDelta"].ToString();
                drv["市場Theta IV 金額週變化(部位)(仟)"] = Math.Round(Convert.ToDouble(thetaiv_weekdelta));

                string med_hv60d_volratio = dr["Med_HV60D_VolRatio"].ToString();
                drv["市場 Med_Vol / HV_60D"] = Convert.ToDouble(med_hv60d_volratio);

                string theta_days = dr["Theta_Days"].ToString();
                drv["Theta天數"] = Math.Round(Convert.ToDouble(theta_days));

                string optweightinprice1to2dot5 = dr["OptWeightInPrice1To2dot5"].ToString();
                drv["庫存權證價格在1~2.5元的比重"] = Math.Round(Convert.ToDouble(optweightinprice1to2dot5));
                */
                string alphatheta = "";
                if (CP == "C")
                {
                    alphatheta = dr["AlphaTheta"].ToString();
                }
                else
                {
                    alphatheta = dr["AlphaTheta_P"].ToString();
                }
                drv["市場超額利潤(仟)"] = Convert.ToDouble(alphatheta);

                string avgalphathetacost = "";
                if (CP == "C")
                {
                    avgalphathetacost = dr["AvgAlphaThetaCost"].ToString();
                }
                else
                {
                    avgalphathetacost = dr["AvgAlphaThetaCost_P"].ToString();
                }
                drv["平均每檔權證超額利潤(元)"] = Convert.ToDouble(avgalphathetacost);

                string theta_enddate = "";
                if (CP == "C")
                {
                    theta_enddate = dr["Theta_EndDate"].ToString();
                }
                else
                {
                    theta_enddate = dr["Theta_EndDate_P"].ToString();
                }
                drv["市場Theta IV(仟)"] = Math.Round(Convert.ToDouble(theta_enddate));
                /*
                string thetaiv_weekdelta = "";
                if (CP == "C")
                {
                    thetaiv_weekdelta = dr["ThetaIV_WeekDelta"].ToString();
                }
                else
                {
                    thetaiv_weekdelta = dr["ThetaIV_WeekDelta_P"].ToString();
                }
                drv["市場Theta IV 金額週變化(部位)(仟)"] = Math.Round(Convert.ToDouble(thetaiv_weekdelta));
                */
                string med_hv60d_volratio = "";
                if (CP == "C")
                {
                    med_hv60d_volratio = dr["Med_HV60D_VolRatio"].ToString();
                }
                else
                {
                    med_hv60d_volratio = dr["Med_HV60D_VolRatio_P"].ToString();
                }
                drv["市場 Med_Vol / HV_60D"] = Convert.ToDouble(med_hv60d_volratio);

                string theta_days = "";
                if (CP == "C")
                {
                    theta_days = dr["Theta_Days"].ToString();
                }
                else
                {
                    theta_days = dr["Theta_Days_P"].ToString();
                }
                drv["Theta天數"] = Math.Round(Convert.ToDouble(theta_days));

                string appraisalrank = "";
                if (CP == "C")
                {
                   appraisalrank = dr["AppraisalRank"].ToString();
                }
                else
                {
                    appraisalrank = dr["AppraisalRank_P"].ToString();
                }
                drv["評鑑權證比重排名"] = Math.Round(Convert.ToDouble(appraisalrank));
                string financingratio = dr["FinancingRatio"].ToString();
                drv["融資使用率(%)"] = Convert.ToDouble(financingratio);

                string callmarketshare = dr["CallMarketShare"].ToString();
                drv["C有效檔數市佔(%)"] = Convert.ToDouble(callmarketshare);

                string putmarketshare = dr["PutMarketShare"].ToString();
                drv["P有效檔數市佔(%)"] = Convert.ToDouble(putmarketshare);

                string calldensity = dr["CallDensity"].ToString();
               
                if (Convert.ToDouble(calldensity) < 0)
                    drv["自家C發行密度"] = "";
                else
                    drv["自家C發行密度"] = calldensity;


                string putdensity = dr["PutDensity"].ToString();
                
                if (Convert.ToDouble(putdensity) < 0)
                    drv["自家P發行密度"] = "";
                else
                    drv["自家P發行密度"] = putdensity;

                string yuancalldensity = dr["YuanCallDensity"].ToString();

                if (Convert.ToDouble(yuancalldensity) < 0)
                    drv["元大C發行密度"] = "";
                else
                    drv["元大C發行密度"] = yuancalldensity;

                string yuanputdensity = dr["YuanPutDensity"].ToString();

                if (Convert.ToDouble(yuanputdensity) < 0)
                    drv["元大P發行密度"] = "";
                else
                    drv["元大P發行密度"] = yuanputdensity;

                string kgiyuancalldensity = dr["KgiYuanCallDensity"].ToString();

                if (Convert.ToDouble(kgiyuancalldensity) >=50)
                    drv["凱基-元大C發行密度"] = "";
                else
                    drv["凱基-元大C發行密度"] = kgiyuancalldensity;

                string kgiyuanputdensity = dr["KgiYuanPutDensity"].ToString();

                if (Convert.ToDouble(kgiyuanputdensity) >=50)
                    drv["凱基-元大P發行密度"] = "";
                else
                    drv["凱基-元大P發行密度"] = kgiyuanputdensity;
                dt.Rows.Add(drv);

            }

        }
        private void SetButton()
        {
            try
            {
                UltraGridBand bands0 = ultraGrid1.DisplayLayout.Bands[0];

                bands0.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
                bands0.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
                bands0.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
                bands0.Columns["標的代號"].CellActivation = Activation.NoEdit;
                bands0.Columns["標的名稱"].CellActivation = Activation.NoEdit;
                bands0.Columns["標的分級"].CellActivation = Activation.NoEdit;
                bands0.Columns["一個月BrokerPL(%)"].CellActivation = Activation.NoEdit;
                //bands0.Columns["標的月損益(仟)"].CellActivation = Activation.NoEdit;
                bands0.Columns["市場剩餘額度(檔)"].CellActivation = Activation.NoEdit;
                bands0.Columns["前日市場剩餘額度(檔)"].CellActivation = Activation.NoEdit;
                //bands0.Columns["前日市場剩餘額度(檔)"].Hidden = true;
                bands0.Columns["今日市場釋出額度(檔)"].CellActivation = Activation.NoEdit;
                bands0.Columns["3日漲跌幅(%)"].CellActivation = Activation.NoEdit;
                //bands0.Columns["3日跌幅(%)"].CellActivation = Activation.NoEdit;
                bands0.Columns["市場Theta IV(仟)"].CellActivation = Activation.NoEdit;
                bands0.Columns["市場部位近五日變化(萬)"].CellActivation = Activation.NoEdit;
                //bands0.Columns["市場Theta IV 金額週變化(部位)(仟)"].CellActivation = Activation.NoEdit;
                bands0.Columns["市場超額利潤(仟)"].CellActivation = Activation.NoEdit;
                bands0.Columns["平均每檔權證超額利潤(元)"].CellActivation = Activation.NoEdit;
                bands0.Columns["市場 Med_Vol / HV_60D"].CellActivation = Activation.NoEdit;
                bands0.Columns["Theta天數"].CellActivation = Activation.NoEdit;
                bands0.Columns["評鑑權證比重排名"].CellActivation = Activation.NoEdit;
                bands0.Columns["融資使用率(%)"].CellActivation = Activation.NoEdit;
                bands0.Columns["C有效檔數市佔(%)"].CellActivation = Activation.NoEdit;
                bands0.Columns["P有效檔數市佔(%)"].CellActivation = Activation.NoEdit;
                bands0.Columns["自家C發行密度"].CellActivation = Activation.NoEdit;
                bands0.Columns["自家P發行密度"].CellActivation = Activation.NoEdit;
                bands0.Columns["元大C發行密度"].CellActivation = Activation.NoEdit;
                bands0.Columns["元大P發行密度"].CellActivation = Activation.NoEdit;
                bands0.Columns["凱基-元大C發行密度"].CellActivation = Activation.NoEdit;
                bands0.Columns["凱基-元大P發行密度"].CellActivation = Activation.NoEdit;
                
                
                ultraGrid1.DisplayLayout.Override.CellAppearance.BackColor = Color.Moccasin;
            }
            catch(Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

        }

        private void ultraGrid1_InitializeLayout(object sender, InitializeLayoutEventArgs e)
        {
            e.Layout.ScrollBounds = ScrollBounds.ScrollToFill;
        }
        private void UltraGrid1_InitializeRow(object sender, InitializeRowEventArgs e)
        {
            
            double num = Convert.ToDouble(e.Row.Cells["前日市場剩餘額度(檔)"].Value.ToString());
            if (num > 0)
                e.Row.Cells["今日市場釋出額度(檔)"].Value = "";
            
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
        
        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            LoadDataByTrader();
        }

        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
            LoadDataByTrader();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            SaveFileDialog sFileDialog = new SaveFileDialog();
            sFileDialog.Title = "匯出Excel";
            sFileDialog.Filter = "EXCEL檔 (*.xlsx)|*.xlsx";
            sFileDialog.InitialDirectory = "D:\\";
            Microsoft.Office.Interop.Excel.Application app = new Microsoft.Office.Interop.Excel.Application();
            if (sFileDialog.ShowDialog() == System.Windows.Forms.DialogResult.OK && sFileDialog.FileName != null)
            {

                Workbook workbook = app.Workbooks.Add(1);
                Worksheet worksheets = workbook.Sheets[1];
                int colnum = ultraGrid1.DisplayLayout.Bands[0].Columns.Count;
                int rownum = ultraGrid1.Rows.Count;
                for (int i = 0; i < colnum; i++)
                {
                    worksheets.get_Range($"{(char)(65 + i) + "1"}", $"{(char)(65 + i) + "1"}").Value = ultraGrid1.DisplayLayout.Bands[0].Columns[i].ToString();
                }
                Microsoft.Office.Interop.Excel.Range range = worksheets.get_Range("A2", $"A{rownum + 1}");
                range.NumberFormat = "@";
                for (int i = 0; i < rownum; i++)
                {
                    for (int j = 0; j < colnum; j++)
                    {
                        worksheets.get_Range($"{(char)(65 + j) + (i + 2).ToString()}", $"{(char)(65 + j) + (i + 2).ToString()}").Value = ultraGrid1.Rows[i].Cells[j].Value;
                        //MessageBox.Show($"{(char)(65 + j) + (i + 2).ToString()} {(char)(65 + j) + (i + 2).ToString()}  { dataGridView1.Rows[i].Cells[j].Value}");
                    }
                }
                workbook.SaveAs(sFileDialog.FileName);
                workbook.Close();
                app.Quit();
                MessageBox.Show("匯出完成");
            }
        }
    }
}

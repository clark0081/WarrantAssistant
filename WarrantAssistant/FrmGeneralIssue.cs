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
using EDLib;

namespace WarrantAssistant
{

    public partial class FrmGeneralIssue : Form
    {
        //20210906利率改版，個股從2.5改成1，指數(IX0001、IX0027)改成0
        //public static double interestR = 0.025;
        public static double interestR_Index = 0.0;
        public static double interestR = 0.01;
        public int duration = 6;
        public int tradeDays = 252;
        public int calendarDays = 365;
        public double vspTotal = 0;
        public double vspAvg = 0;
        public static double roundpriceK(double x)
        {
            if (x <= 100)
            {
                double tick = 0.5;
                int up = (int)Math.Floor(x) + 1;//49
                int down = (int)Math.Floor(x);//48
                if ((double)x % 1 > tick)//ex 48.6
                {
                    if ((double)(up - x) > (double)(x - down - tick))
                        return (double)(down + tick);
                    else
                        return (double)up;
                }
                else//ex3.4
                {
                    if ((double)(down + tick - x) > (double)(x - down))
                        return (double)(down);
                    else
                        return (double)(down + tick);
                }
            }
            else if (x <= 500)
                return Math.Round(x);
            else if (x <= 1000)
            {
                double tick = 5;
                int up = (int)Math.Floor((double)(x / 10)) * 10 + 10;
                int down = (int)Math.Floor((double)(x / 10)) * 10;
                if ((double)x % 10 > tick)//ex 498
                {
                    if ((double)(up - x) > (double)(x - down - tick))
                        return (double)(down + tick);
                    else
                        return (double)up;
                }
                else//ex 
                {
                    if ((double)(down + tick - x) > (double)(x - down))
                        return (double)(down);
                    else
                        return (double)(down + tick);
                }
            }
            else
            {
                return Math.Round(x / 10) * 10;
            }
        }

        public static double uJumpSize(double s)
        {
            //每股市價未滿10元者，股價升降單位為0.01元，10元至未滿50元者為0.05元、50元至未滿100元者為0.1元、100元至未滿500元者為0.5元、500元至未滿1000元者為1元、1000元以上者為5元。
            if (s < 10)
                return 0.01;
            else if (s < 50)
                return 0.05;
            else if (s < 100)
                return 0.1;
            else if (s < 500)
                return 0.5;
            else if (s < 1000)
                return 1;
            else
                return 5;
        }
        private DataTable dt = new DataTable();
        //加減碼錶
        private DataTable dt_adj = new DataTable();
        //發行共通設定
        private DataTable dt_globalvar = new DataTable();
        //跳動價差預設值
        private DataTable dt_jumpsize = new DataTable();

        private DataTable uid_classify;

        public DateTime lastTradeDate = EDLib.TradeDate.LastNTradeDate(1);
        //public DateTime lastTradeDate = EDLib.TradeDate.LastNTradeDate(1);

        public Dictionary<string, double> ADJ = new Dictionary<string, double>();

        private DataTable selectResult = new DataTable();

        //昨日收盤價
        private Dictionary<string, double> ucloseP = new Dictionary<string, double>();

        //標的名稱
        private Dictionary<string, string> uName = new Dictionary<string, string>();

        //將流動性分級為C
        private List<string> liquidityC = new List<string>();

        //建議Vol
        private Dictionary<string, Dictionary<string, double>> suggestVol_CP = new Dictionary<string, Dictionary<string, double>>();


        private Dictionary<string, double> underlying_hv60 = new Dictionary<string, double>();

        private Dictionary<string, Dictionary<string, string>> Issuable_CP = new Dictionary<string, Dictionary<string, string>>();

        private Dictionary<string, Dictionary<string, double>> ThetaShare_CP = new Dictionary<string, Dictionary<string, double>>();

        public string userID = GlobalVar.globalParameter.userID;

        private Dictionary<string, List<string>> underlying_trader = new Dictionary<string, List<string>>();
        private Dictionary<string, string> underlying2trader = new Dictionary<string, string>();

        private Dictionary<string, double> underlying_issueCredit = new Dictionary<string, double>();

        private Dictionary<string, double> underlying_issuePercent = new Dictionary<string, double>();

        private Dictionary<string, double> underlying_issueCreditReward = new Dictionary<string, double>();

        private Dictionary<string, Dictionary<string, int>> k_seperate = new Dictionary<string, Dictionary<string, int>>();

        List<int> DeletedSerialNum = new List<int>();


        //private DataTable lastTradeDate = EDLib.TradeDate.LastNTradeDate(1);
        public FrmGeneralIssue()
        {
            InitializeComponent();
        }

        private void FrmGeneralIssue_Load(object sender, EventArgs e)
        {
            InitialGrid();
            foreach (var item in GlobalVar.globalParameter.traders)
            {
                comboBox1.Items.Add(item);
                underlying_trader.Add(item.TrimStart('0'), new List<string>());
            }
            //comboBox1.Items.Add("All");
            //underlying_trader.Add("All", new List<string>());
            string getUnderlyingTraderStr = $@"SELECT  [UID]
                                                      ,[TraderAccount]
                                                  FROM [TwData].[dbo].[Underlying_Trader]
                                                  WHERE LEN(UID) < 5 AND LEFT(UID, 2) <> '00'";
            DataTable underlyingTrader = MSSQL.ExecSqlQry(getUnderlyingTraderStr, GlobalVar.loginSet.twData);
            foreach(DataRow dr in underlyingTrader.Rows)
            {
                string uid = dr["UID"].ToString();
                string trader = dr["TraderAccount"].ToString();
                underlying2trader.Add(uid, trader.PadLeft(7,'0'));
                if (underlying_trader.ContainsKey(trader))
                    underlying_trader[trader].Add(uid);
                //underlying_trader["All"].Add(uid);
            }
            comboBox1.Text = userID;
            try
            {
                LoadAll();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            
        }
        private void LoadAll()
        {
            LoadData();
            LoadPara();//traderID
            LoadAdj();//traderID
            LoadGlobalVar();//traderID
            LoadJumpSize();//traderID
            LoadCredit();
            LoadRewardCredit();
            LoadDeletedSerial();
            LoadIssuable();
            LoadThetaShare();
            LoadLiquidityC();
        }
        private void InitialGrid()
        {
            dt.Columns.Add("組別", typeof(string));
            dt.Columns.Add("分類", typeof(string));
            dt.Columns.Add("CP", typeof(string));
            dt.Columns.Add("TtoM", typeof(int));
            dt.Columns.Add("基本市佔", typeof(double));
            dt.Columns.Add("價內外起始", typeof(double));
            dt.Columns.Add("價內外長度", typeof(double));
            dt.Columns.Add("交易員", typeof(string));
            dt.Columns.Add("使用", typeof(bool));//bool直接變checkedbox
            dt.Columns["使用"].ReadOnly = false;

            ultraGrid1.DataSource = dt;
            UltraGridBand band0 = ultraGrid1.DisplayLayout.Bands[0];
            band0.Columns["組別"].Format = "N0";
            band0.Columns["分類"].Format = "N0";
            band0.Columns["CP"].Format = "N0";
            band0.Columns["TtoM"].Format = "N0";
            band0.Columns["基本市佔"].Format = "N0";
            band0.Columns["價內外起始"].Format = "N0";

            // band0.Columns["類型"].Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownList;
            //band0.Columns["CP"].Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownList;
            //band0.Columns["交易員"].Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownList;
            //band0.Columns["發行原因"].Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownList;
            //ultraGrid1.DisplayLayout.Bands[0].Columns["確認"].Style = Infragistics.Win.UltraWinGrid.ColumnStyle.CheckBox;
            //ultraGrid1.DisplayLayout.Bands[0].Columns["刪除"].Style = Infragistics.Win.UltraWinGrid.ColumnStyle.CheckBox;

            //ultraGrid1.DisplayLayout.Bands[0].Columns["編號"].Width = 75;
            band0.Columns["組別"].Width = 30;
            band0.Columns["分類"].Width = 40;
            band0.Columns["CP"].Width = 30;
            band0.Columns["TtoM"].Width = 30;
            band0.Columns["基本市佔"].Width = 40;
            band0.Columns["價內外起始"].Width = 40;
            band0.Columns["價內外長度"].Width = 40;
            band0.Columns["交易員"].Width = 50;
            band0.Columns["使用"].Width = 40;

            band0.Columns["組別"].CellAppearance.BackColor = Color.LightGray;
            band0.Columns["分類"].CellAppearance.BackColor = Color.LightGray;
            band0.Columns["CP"].CellAppearance.BackColor = Color.LightGray;
            band0.Columns["TtoM"].CellAppearance.BackColor = Color.LightGray;
            band0.Columns["基本市佔"].CellAppearance.BackColor = Color.LightGray;
            band0.Columns["價內外起始"].CellAppearance.BackColor = Color.LightGray;
            band0.Columns["價內外長度"].CellAppearance.BackColor = Color.LightGray;
            band0.Columns["交易員"].CellAppearance.BackColor = Color.LightBlue;



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
            ultraGrid1.DisplayLayout.Override.RowSelectorNumberStyle = RowSelectorNumberStyle.None;
            //ultraGrid1.DisplayLayout.Bands[0].Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
            //ultraGrid1.DisplayLayout.Bands[0].Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
            //ultraGrid1.DisplayLayout.Bands[0].Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.False;
            //ultraGrid1.DisplayLayout.Bands[0].Columns["確認"].CellActivation = Activation.AllowEdit;


            //設定加減碼參數
            dt_adj.Columns.Add("分類", typeof(string));
            dt_adj.Columns.Add("加減碼", typeof(double));
            ultraGrid2.DataSource = dt_adj;
            UltraGridBand band_adj = ultraGrid2.DisplayLayout.Bands[0];
            band_adj.Columns["分類"].Format = "N0";
            band_adj.Columns["加減碼"].Format = "N0";

            band_adj.Columns["分類"].Width = 60;
            band_adj.Columns["加減碼"].Width = 60;
            


            band_adj.Columns["分類"].CellAppearance.BackColor = Color.LightGray;
            band_adj.Columns["加減碼"].CellAppearance.BackColor = Color.LightGray;

            band_adj.SortedColumns.Clear();

            ultraGrid2.DisplayLayout.Bands[0].Override.HeaderAppearance.TextHAlign = Infragistics.Win.HAlign.Left;
            ultraGrid2.DisplayLayout.Override.RowSelectorNumberStyle = RowSelectorNumberStyle.None;
            //設定發行共通參數
            dt_globalvar.Columns.Add("設定", typeof(string));
            dt_globalvar.Columns.Add("參數", typeof(double));
            ultraGrid3.DataSource = dt_globalvar;
            UltraGridBand band_globalvar = ultraGrid3.DisplayLayout.Bands[0];
            band_globalvar.Columns["設定"].Format = "N0";
            //band_globalvar.Columns["參數"].Format = "N0";

            band_globalvar.Columns["設定"].Width = 60;
            band_globalvar.Columns["參數"].Width = 60;



            band_globalvar.Columns["設定"].CellAppearance.BackColor = Color.LightGray;
            band_globalvar.Columns["參數"].CellAppearance.BackColor = Color.LightGray;

            band_globalvar.SortedColumns.Clear();

            ultraGrid3.DisplayLayout.Bands[0].Override.HeaderAppearance.TextHAlign = Infragistics.Win.HAlign.Left;
            ultraGrid3.DisplayLayout.Override.RowSelectorNumberStyle = RowSelectorNumberStyle.None;

            //設定跳動價差
            dt_jumpsize.Columns.Add("價內外起(含)", typeof(double));
            dt_jumpsize.Columns.Add("價內外終", typeof(double));
            dt_jumpsize.Columns.Add("跳動價差", typeof(int));
            ultraGrid4.DataSource = dt_jumpsize;
            UltraGridBand band_jumpsize = ultraGrid4.DisplayLayout.Bands[0];
            band_jumpsize.Columns["價內外起(含)"].Format = "N0";
            band_jumpsize.Columns["價內外終"].Format = "N0";
            band_jumpsize.Columns["跳動價差"].Format = "N0";

            band_jumpsize.Columns["價內外起(含)"].Width = 60;
            band_jumpsize.Columns["價內外終"].Width = 60;
            band_jumpsize.Columns["跳動價差"].Width = 60;


            band_jumpsize.Columns["價內外起(含)"].CellAppearance.BackColor = Color.LightGray;
            band_jumpsize.Columns["價內外終"].CellAppearance.BackColor = Color.LightGray;
            band_jumpsize.Columns["跳動價差"].CellAppearance.BackColor = Color.LightGray;

            band_jumpsize.SortedColumns.Clear();

            ultraGrid4.DisplayLayout.Bands[0].Override.HeaderAppearance.TextHAlign = Infragistics.Win.HAlign.Left;
            ultraGrid4.DisplayLayout.Override.RowSelectorNumberStyle = RowSelectorNumberStyle.None;
            
            //篩選結果Table

            selectResult.Columns.Add("標的代號", typeof(string));
            selectResult.Columns.Add("標的名稱", typeof(string));
            selectResult.Columns.Add("確認", typeof(bool));
            selectResult.Columns.Add("行使比例", typeof(double));
            selectResult.Columns.Add("履約價", typeof(double));
            selectResult.Columns.Add("標的收盤價", typeof(double));
            selectResult.Columns.Add("推薦價內外", typeof(double));
            selectResult.Columns.Add("建議Vol", typeof(double));
            selectResult.Columns.Add("HV", typeof(double));
            selectResult.Columns.Add("HV倍數", typeof(double));
            selectResult.Columns.Add("權證跳動價差", typeof(double));
            selectResult.Columns.Add("試算價格", typeof(double));
            selectResult.Columns.Add("額度試算", typeof(double));
            selectResult.Columns.Add("額度", typeof(double));
            selectResult.Columns.Add("已發行", typeof(double));
            selectResult.Columns.Add("獎勵額度", typeof(double));
            selectResult.Columns.Add("CP", typeof(string));
            selectResult.Columns.Add("KGI檔數", typeof(int));
            selectResult.Columns.Add("基本檔數", typeof(double));
            selectResult.Columns.Add("總檔數", typeof(int));
            selectResult.Columns.Add("建議Moneyness間隔", typeof(double));
            //selectResult.Columns.Add("格子數", typeof(double));
            selectResult.Columns.Add("檔數市佔", typeof(double));
            selectResult.Columns.Add("目標市佔", typeof(double));
            selectResult.Columns.Add("Theta市佔", typeof(double));
            selectResult.Columns.Add("Theta分類", typeof(string));
            selectResult.Columns.Add("Margin分類", typeof(string));
            //selectResult.Columns.Add("現股跳動價差", typeof(double));
            selectResult.Columns.Add("標的VSP", typeof(double));
            selectResult.Columns.Add("單檔VSP", typeof(double));




            ultraGrid5.DataSource = selectResult;
            UltraGridBand band_selectResult = ultraGrid5.DisplayLayout.Bands[0];
            /*
            band_selectResult.Columns["標的代號"].Format = "N0";
            band_selectResult.Columns["標的名稱"].Format = "N0";
            band_selectResult.Columns["KGI檔數"].Format = "N0";
            band_selectResult.Columns["總檔數"].Format = "N0";
            band_selectResult.Columns["基本檔數"].Format = "N0";
            band_selectResult.Columns["建議Moneyness間隔"].Format = "N0";
            band_selectResult.Columns["格子數"].Format = "N0";
            band_selectResult.Columns["檔數市佔"].Format = "N0";
            band_selectResult.Columns["建議Vol"].Format = "N0";
            band_selectResult.Columns["標的收盤價"].Format = "N0";
            band_selectResult.Columns["Theta分類"].Format = "N0";
            band_selectResult.Columns["Margin分類"].Format = "N0";
            band_selectResult.Columns["CP"].Format = "N0";
            band_selectResult.Columns["推薦價內外"].Format = "N0";
            band_selectResult.Columns["現股跳動價差"].Format = "N0";
            band_selectResult.Columns["履約價"].Format = "N0";
            band_selectResult.Columns["權證跳動價差"].Format = "N0";
            band_selectResult.Columns["行使比例"].Format = "N0";
            band_selectResult.Columns["試算價格"].Format = "N0";
            */

            band_selectResult.Columns["標的代號"].Width = 40;
            band_selectResult.Columns["標的名稱"].Width = 60;
            band_selectResult.Columns["確認"].Width = 30;
            band_selectResult.Columns["行使比例"].Width = 60;
            band_selectResult.Columns["試算價格"].Width = 60;
            band_selectResult.Columns["額度試算"].Width = 60;
            band_selectResult.Columns["額度"].Width = 60;
            band_selectResult.Columns["已發行"].Width = 40;
            band_selectResult.Columns["獎勵額度"].Width = 60;
            band_selectResult.Columns["KGI檔數"].Width = 60;
            band_selectResult.Columns["總檔數"].Width = 40;
            band_selectResult.Columns["基本檔數"].Width = 60;
            band_selectResult.Columns["建議Moneyness間隔"].Width = 60;
            //band_selectResult.Columns["格子數"].Width = 60;
            band_selectResult.Columns["檔數市佔"].Width = 60;
            band_selectResult.Columns["目標市佔"].Width = 40;
            band_selectResult.Columns["Theta市佔"].Width = 40;
            band_selectResult.Columns["建議Vol"].Width = 60;
            band_selectResult.Columns["HV"].Width = 40;
            band_selectResult.Columns["HV倍數"].Width = 40;
            band_selectResult.Columns["標的收盤價"].Width = 60;
            band_selectResult.Columns["Theta分類"].Width = 35;
            band_selectResult.Columns["Margin分類"].Width = 35;
            band_selectResult.Columns["CP"].Width = 30;
            band_selectResult.Columns["推薦價內外"].Width = 60;
            //band_selectResult.Columns["現股跳動價差"].Width = 60;
            band_selectResult.Columns["履約價"].Width = 40;
            band_selectResult.Columns["權證跳動價差"].Width = 60;
            band_selectResult.Columns["標的VSP"].Width = 50;
            //band_selectResult.Columns["單檔VSP"].Width = 50;




            band_selectResult.Columns["標的代號"].CellAppearance.BackColor = Color.Wheat;
            band_selectResult.Columns["標的名稱"].CellAppearance.BackColor = Color.Wheat;
            band_selectResult.Columns["確認"].CellAppearance.BackColor = Color.Wheat;
            band_selectResult.Columns["行使比例"].CellAppearance.BackColor = Color.Aquamarine;
            band_selectResult.Columns["試算價格"].CellAppearance.BackColor = Color.Wheat;
            band_selectResult.Columns["額度試算"].CellAppearance.BackColor = Color.Wheat;

            band_selectResult.Columns["額度"].CellAppearance.BackColor = Color.Wheat;
            band_selectResult.Columns["已發行"].CellAppearance.BackColor = Color.Wheat;
            band_selectResult.Columns["獎勵額度"].CellAppearance.BackColor = Color.Wheat;
            band_selectResult.Columns["KGI檔數"].CellAppearance.BackColor = Color.Wheat;
            band_selectResult.Columns["總檔數"].CellAppearance.BackColor = Color.Wheat;
            band_selectResult.Columns["基本檔數"].CellAppearance.BackColor = Color.Wheat;
            band_selectResult.Columns["建議Moneyness間隔"].CellAppearance.BackColor = Color.Wheat;
            //band_selectResult.Columns["格子數"].CellAppearance.BackColor = Color.Wheat;
            band_selectResult.Columns["檔數市佔"].CellAppearance.BackColor = Color.Wheat;
            band_selectResult.Columns["目標市佔"].CellAppearance.BackColor = Color.Wheat;
            band_selectResult.Columns["Theta市佔"].CellAppearance.BackColor = Color.Wheat;
            band_selectResult.Columns["建議Vol"].CellAppearance.BackColor = Color.Aquamarine;
            band_selectResult.Columns["標的收盤價"].CellAppearance.BackColor = Color.Wheat;
            band_selectResult.Columns["HV"].CellAppearance.BackColor = Color.Wheat;
            band_selectResult.Columns["HV倍數"].CellAppearance.BackColor = Color.Wheat;
            band_selectResult.Columns["Theta分類"].CellAppearance.BackColor = Color.Wheat;
            band_selectResult.Columns["Margin分類"].CellAppearance.BackColor = Color.Wheat;
            band_selectResult.Columns["CP"].CellAppearance.BackColor = Color.Wheat;
            band_selectResult.Columns["推薦價內外"].CellAppearance.BackColor = Color.Wheat;
            //band_selectResult.Columns["現股跳動價差"].CellAppearance.BackColor = Color.Wheat;
            band_selectResult.Columns["履約價"].CellAppearance.BackColor = Color.Aquamarine;
            band_selectResult.Columns["權證跳動價差"].CellAppearance.BackColor = Color.Wheat;
            band_selectResult.Columns["標的VSP"].CellAppearance.BackColor = Color.Wheat;
            band_selectResult.Columns["單檔VSP"].CellAppearance.BackColor = Color.Wheat;


            band_selectResult.Columns["標的代號"].CellActivation = Activation.NoEdit;
            band_selectResult.Columns["標的名稱"].CellActivation = Activation.NoEdit;
            band_selectResult.Columns["確認"].CellAppearance.BackColor = Color.LightGray;
            //band_selectResult.Columns["行使比例"].CellAppearance.BackColor = Color.LightBlue;
            band_selectResult.Columns["試算價格"].CellActivation = Activation.NoEdit;
            band_selectResult.Columns["額度試算"].CellActivation = Activation.NoEdit;
            band_selectResult.Columns["額度"].CellActivation = Activation.NoEdit;
            band_selectResult.Columns["已發行"].CellActivation = Activation.NoEdit;
            band_selectResult.Columns["獎勵額度"].CellActivation = Activation.NoEdit;
            band_selectResult.Columns["KGI檔數"].CellActivation = Activation.NoEdit;
            band_selectResult.Columns["總檔數"].CellActivation = Activation.NoEdit;
            band_selectResult.Columns["基本檔數"].CellActivation = Activation.NoEdit;
            band_selectResult.Columns["建議Moneyness間隔"].CellActivation = Activation.NoEdit;
            //band_selectResult.Columns["格子數"].CellActivation = Activation.NoEdit;
            band_selectResult.Columns["檔數市佔"].CellActivation = Activation.NoEdit;
            band_selectResult.Columns["目標市佔"].CellActivation = Activation.NoEdit;
            band_selectResult.Columns["Theta市佔"].CellActivation = Activation.NoEdit;
            //band_selectResult.Columns["建議Vol"].CellActivation = Activation.NoEdit;
            band_selectResult.Columns["標的收盤價"].CellActivation = Activation.NoEdit;
            band_selectResult.Columns["Theta分類"].CellActivation = Activation.NoEdit;
            band_selectResult.Columns["Margin分類"].CellActivation = Activation.NoEdit;
            band_selectResult.Columns["CP"].CellActivation = Activation.NoEdit;
            band_selectResult.Columns["推薦價內外"].CellActivation = Activation.NoEdit;
            //band_selectResult.Columns["現股跳動價差"].CellActivation = Activation.NoEdit;
            //band_selectResult.Columns["履約價"].CellAppearance.BackColor = Color.LightBlue;
            band_selectResult.Columns["權證跳動價差"].CellActivation = Activation.NoEdit;
            band_selectResult.Columns["標的VSP"].CellActivation = Activation.NoEdit;
            band_selectResult.Columns["單檔VSP"].CellActivation = Activation.NoEdit;
            band_selectResult.Columns["HV"].CellActivation = Activation.NoEdit;
            band_selectResult.Columns["HV倍數"].CellActivation = Activation.NoEdit;


            band_selectResult.SortedColumns.Clear();

            ultraGrid5.DisplayLayout.Bands[0].Override.HeaderAppearance.TextHAlign = Infragistics.Win.HAlign.Left;
            //ultraGrid5.DisplayLayout.Override.RowSelectorNumberStyle = RowSelectorNumberStyle.None;
            this.ultraGrid2.DisplayLayout.AutoFitStyle = Infragistics.Win.UltraWinGrid.AutoFitStyle.ResizeAllColumns;
            this.ultraGrid3.DisplayLayout.AutoFitStyle = Infragistics.Win.UltraWinGrid.AutoFitStyle.ResizeAllColumns;
            this.ultraGrid4.DisplayLayout.AutoFitStyle = Infragistics.Win.UltraWinGrid.AutoFitStyle.ResizeAllColumns;

            this.ultraGrid5.DisplayLayout.AutoFitStyle = Infragistics.Win.UltraWinGrid.AutoFitStyle.ResizeAllColumns;
            this.ultraGrid6.DisplayLayout.AutoFitStyle = Infragistics.Win.UltraWinGrid.AutoFitStyle.ResizeAllColumns;
            this.ultraGrid6.DisplayLayout.Bands[0].Override.HeaderAppearance.TextHAlign = Infragistics.Win.HAlign.Center;
            //SetButton();
        }

        private void LoadThetaShare()
        {
            ThetaShare_CP.Clear();
            string sql_C = $@"SELECT A.UID, CASE WHEN A.AllTheta > 0  THEN ISNULL(B.KGITheta,0) / A.AllTheta ELSE 0 END AS SHARE FROM (
                                      SELECT  [UID], SUM(AccReleasingLots * Theta_IV) AS AllTheta
                                      FROM [TwData].[dbo].[V_WarrantTrading]
                                      WHERE [TDate] = '{lastTradeDate.ToString("yyyyMMdd")}' AND [WClass] = 'c' AND [TtoM] >= 70
                                      GROUP BY [UID]) AS A

                                      LEFT JOIN 
                                      (
                                      SELECT  [UID], SUM(AccReleasingLots * Theta_IV) AS KGITheta
                                      FROM [TwData].[dbo].[V_WarrantTrading]
                                      WHERE [TDate] = '{lastTradeDate.ToString("yyyyMMdd")}' AND [WClass] = 'c' AND [TtoM] >= 70 AND [IssuerName] = '9200'
                                      GROUP BY [UID]) AS B ON A.UID = B.UID";
            DataTable dt_C = MSSQL.ExecSqlQry(sql_C, GlobalVar.loginSet.twData);
            ThetaShare_CP.Add("c", new Dictionary<string, double>());
            foreach(DataRow dr_C in dt_C.Rows)
            {
                string uid = dr_C["UID"].ToString();
                double share = Math.Round(Convert.ToDouble(dr_C["SHARE"].ToString()) * 100, 1);
                ThetaShare_CP["c"].Add(uid, share);
            }
            string sql_P = $@"SELECT A.UID, CASE WHEN A.AllTheta > 0  THEN ISNULL(B.KGITheta,0) / A.AllTheta ELSE 0 END AS SHARE FROM (
                                      SELECT  [UID], SUM(AccReleasingLots * Theta_IV) AS AllTheta
                                      FROM [TwData].[dbo].[V_WarrantTrading]
                                      WHERE [TDate] = '{lastTradeDate.ToString("yyyyMMdd")}' AND [WClass] = 'p' AND [TtoM] >= 70
                                      GROUP BY [UID]) AS A

                                      LEFT JOIN 
                                      (
                                      SELECT  [UID], SUM(AccReleasingLots * Theta_IV) AS KGITheta
                                      FROM [TwData].[dbo].[V_WarrantTrading]
                                      WHERE [TDate] = '{lastTradeDate.ToString("yyyyMMdd")}' AND [WClass] = 'p' AND [TtoM] >= 70 AND [IssuerName] = '9200'
                                      GROUP BY [UID]) AS B ON A.UID = B.UID";
            DataTable dt_P = MSSQL.ExecSqlQry(sql_P, GlobalVar.loginSet.twData);
            ThetaShare_CP.Add("p", new Dictionary<string, double>());
            foreach (DataRow dr_P in dt_P.Rows)
            {
                string uid = dr_P["UID"].ToString();
                double share = Math.Round(Convert.ToDouble(dr_P["SHARE"].ToString()) * 100, 1);
                ThetaShare_CP["p"].Add(uid, share);
            }
        }
        private void LoadIssuable()
        {
            Issuable_CP.Clear();
            Issuable_CP.Add("c", new Dictionary<string, string>());
            Issuable_CP.Add("p", new Dictionary<string, string>());
            string sql_C = $@"SELECT  [UnderlyingID]
                              ,[Issuable]
                          FROM [WarrantAssistant].[dbo].[WarrantUnderlyingSummary]
                          WHERE LEN([UnderlyingID]) < 5 AND LEFT([UnderlyingID],2) <> '00'";
            DataTable dt_C = MSSQL.ExecSqlQry(sql_C, GlobalVar.loginSet.warrantassistant45);
            foreach(DataRow dr in dt_C.Rows)
            {
                string uid = dr["UnderlyingID"].ToString();
                string issuable = dr["Issuable"].ToString();
                Issuable_CP["c"].Add(uid, issuable);
            }
            string sql_P = $@"SELECT  [UnderlyingID]
                              ,[PutIssuable]
                          FROM [WarrantAssistant].[dbo].[WarrantUnderlyingSummary]
                          WHERE LEN([UnderlyingID]) < 5 AND LEFT([UnderlyingID],2) <> '00'";
            DataTable dt_P = MSSQL.ExecSqlQry(sql_P, GlobalVar.loginSet.warrantassistant45);
            foreach (DataRow dr in dt_P.Rows)
            {
                string uid = dr["UnderlyingID"].ToString();
                string issuable = dr["PutIssuable"].ToString();
                Issuable_CP["p"].Add(uid, issuable);
            }
        }
        private void LoadDeletedSerial()
        {
            DeletedSerialNum.Clear();
            string sqlDelSerial = $@"SELECT [SerialNum] FROM [WarrantAssistant].[dbo].[TempListDeleteLog]
                                    WHERE [Trader] ='{userID}' AND [DateTime] >='{DateTime.Today.ToString("yyyyMMdd")}'";
            DataTable dtDelSerial = MSSQL.ExecSqlQry(sqlDelSerial, GlobalVar.loginSet.warrantassistant45);

            foreach (DataRow dr in dtDelSerial.Rows)
            {
                string serialNum = dr["SerialNum"].ToString();
                if (serialNum.Length == 0)
                    continue;
                int index = Convert.ToInt32(serialNum.Substring(17, serialNum.Length - 17));
                DeletedSerialNum.Add(index);

            }
            DeletedSerialNum.Sort();
        }
        private void LoadData()
        {
            ucloseP.Clear();
            uName.Clear();
            suggestVol_CP.Clear();
            underlying_hv60.Clear();

            //抓標的分類
            string sql = $@"SELECT  
                          [UID]
                          ,[WClass]
                          ,[class]
                          ,[ProfitRank]
                          ,[VSP70D]
						  ,CASE WHEN [權證檔數70D] > 0 THEN ROUND([VSP70D] / [權證檔數70D],0) ELSE 0 END AS [VSP70D_AVG]
                          ,[目標市佔]
                          ,[目標TtoM]
                      FROM [TwData].[dbo].[URank]
                      WHERE [TDate] = '{lastTradeDate.ToString("yyyyMMdd")}' AND [IndexOrNot] = 'N' AND [class] IS NOT NULL AND [ProfitRank] IS NOT NULL";
            uid_classify = MSSQL.ExecSqlQry(sql, GlobalVar.loginSet.twData);

            string sql2 = $@"SELECT  [股票代號]
                                      ,[股票名稱]
                                      ,[收盤價]
                                  FROM [TwCMData].[dbo].[日收盤表排行]
                                  WHERE [日期] = '{lastTradeDate.ToString("yyyyMMdd")}' AND LEN([股票代號]) < 5 AND LEFT([股票代號], 2 ) <> '00'";
            DataTable dt2 = MSSQL.ExecSqlQry(sql2, GlobalVar.loginSet.twCMData);
            foreach(DataRow dr in dt2.Rows)
            {
                string uid = dr["股票代號"].ToString();
                string uname = dr["股票名稱"].ToString();
                double closeP = Convert.ToDouble(dr["收盤價"].ToString());
                ucloseP.Add(uid, closeP);
                uName.Add(uid, uname);
            }

            suggestVol_CP.Add("c", new Dictionary<string, double>());
            string sql3 = $@"SELECT  
                              [UID]
                              ,[RecommendVol]
                          FROM [TwData].[dbo].[RecommendVolAll]
                          WHERE [DateTime] = '{DateTime.Today.ToString("yyyyMMdd")}' AND [CP] = 'C' AND [AllCounts] > 0  AND LEN([UID]) < 5 AND LEFT([UID], 2 ) <> '00'";
            DataTable dt3 = MSSQL.ExecSqlQry(sql3, GlobalVar.loginSet.twData);
            foreach(DataRow dr3 in dt3.Rows)
            {
                string uid = dr3["UID"].ToString();
                double vol = Convert.ToDouble(dr3["RecommendVol"].ToString());
                suggestVol_CP["c"].Add(uid, vol);
            }
            suggestVol_CP.Add("p", new Dictionary<string, double>());
            string sql4 = $@"SELECT  
                              [UID]
                              ,[RecommendVol]
                          FROM [TwData].[dbo].[RecommendVolAll]
                          WHERE [DateTime] = '{DateTime.Today.ToString("yyyyMMdd")}' AND [CP] = 'P' AND [AllCounts] > 0  AND LEN([UID]) < 5 AND LEFT([UID], 2 ) <> '00'";
            DataTable dt4 = MSSQL.ExecSqlQry(sql4, GlobalVar.loginSet.twData);
            foreach (DataRow dr4 in dt4.Rows)
            {
                string uid = dr4["UID"].ToString();
                double vol = Convert.ToDouble(dr4["RecommendVol"].ToString());
                suggestVol_CP["p"].Add(uid, vol);
            }
            string sql5 = $@"SELECT  
                                  [UID]
                                  ,[HV60]
                              FROM [TwData].[dbo].[RecommendVolAll]
                              WHERE [DateTime] = '{DateTime.Today.ToString("yyyyMMdd")}'";
            DataTable dt5 = MSSQL.ExecSqlQry(sql5, GlobalVar.loginSet.twData);
            foreach (DataRow dr5 in dt5.Rows)
            {
                string uid = dr5["UID"].ToString();
                double hv = Convert.ToDouble(dr5["HV60"].ToString());
                if (!underlying_hv60.ContainsKey(uid))
                    underlying_hv60.Add(uid, hv);
            }
        }
        private void LoadCredit()
        {
            underlying_issueCredit.Clear();
            underlying_issuePercent.Clear();
            string sql = $@"SELECT [UID], ROUND([CanIssue],0) AS CanIssue, ROUND([IssuedPercent],1) AS IssuedPercent
                              FROM [WarrantAssistant].[dbo].[WarrantUnderlyingCreditNew]
                              WHERE [UpdateTime] > '{DateTime.Today.ToString("yyyyMMdd")}'";
            DataTable dt = MSSQL.ExecSqlQry(sql, GlobalVar.loginSet.warrantassistant45);
            foreach(DataRow dr in dt.Rows)
            {
                string uid = dr["UID"].ToString();
                double credit = Convert.ToDouble(dr["CanIssue"].ToString());
                double issuedPercent = Convert.ToDouble(dr["IssuedPercent"].ToString());
                underlying_issueCredit.Add(uid, credit);
                underlying_issuePercent.Add(uid, issuedPercent);
            }
        }

        private void LoadLiquidityC()
        {
            liquidityC.Clear();
            string sql = $@"SELECT  [UID] FROM [newEDIS].[dbo].[UnderlyingLiquidity] WHERE [TDate] = '{lastTradeDate.ToString("yyyyMMdd")}' AND [UnderlyingSize] = 'C'";
            DataTable dt = MSSQL.ExecSqlQry(sql, "SERVER=10.60.0.37;DATABASE=newEDIS;UID=WarrantWeb;PWD=WarrantWeb");
            foreach(DataRow dr in dt.Rows)
            {
                string uid = dr["UID"].ToString();
                if (!liquidityC.Contains(uid))
                    liquidityC.Add(uid);
            }
        }
        private void LoadRewardCredit()
        {//IsNull(Floor(E.[WarrantAvailableShares] * {GlobalVar.globalParameter.givenRewardPercent} - IsNull(F.[UsedRewardNum],0)), 0) AS RewardIssueCredit 
            underlying_issueCreditReward.Clear();
            string sql = $@"SELECT A.UID, IsNull(Floor(A.[WarrantAvailableShares] * {GlobalVar.globalParameter.givenRewardPercent} - IsNull(B.[UsedRewardNum],0)), 0) AS RewardIssueCredit  FROM (
                            SELECT [UID], [WarrantAvailableShares]
                              FROM [WarrantAssistant].[dbo].[WarrantUnderlyingCreditNew]
                              WHERE [UpdateTime] > '{DateTime.Today.ToString("yyyyMMdd")}' AND LEN(UID) < 5 AND LEFT(UID, 2) <> '00') AS A
                              LEFT JOIN 
                              (SELECT  [UnderlyingID], [UsedRewardNum]
                                FROM [WarrantAssistant].[dbo].[WarrantReward]
                                WHERE LEN(UnderlyingID) <5 AND LEFT(UnderlyingID,2) <> '00') AS B
                              ON A.UID = B.UnderlyingID";
            DataTable dt = MSSQL.ExecSqlQry(sql, GlobalVar.loginSet.warrantassistant45);
            foreach (DataRow dr in dt.Rows)
            {
                string uid = dr["UID"].ToString();
                double credit = Convert.ToDouble(dr["RewardIssueCredit"].ToString());
                underlying_issueCreditReward.Add(uid, credit);
            }
        }
        //每個交易員的設定參數
        private void LoadPara()
        {
            
            string traderID = comboBox1.Text;
            string sql = $@"SELECT  [GroupNum]
                              ,[Rank]
                              ,[WClass]
                              ,[TtoM]
                              ,[MarketShare]
                              ,[MoneynessLength]
                              ,[MoneynessStart]
                              ,[TraderID]
                              ,[Used]
                          FROM [WarrantAssistant].[dbo].[GeneralIssueGroup]
                          WHERE [TraderID] = '{traderID}'";
            dt.Clear();
            DataTable dtTemp = MSSQL.ExecSqlQry(sql, GlobalVar.loginSet.warrantassistant45);
            try { 
                foreach (DataRow drw in dtTemp.Rows) {

                    DataRow dr = dt.NewRow();
                    dr["組別"] = drw["GroupNum"].ToString();
                    dr["分類"] = drw["Rank"].ToString();
                    dr["CP"] = drw["WClass"].ToString();
                    dr["TtoM"] = Convert.ToInt32(drw["TtoM"].ToString());
                    dr["基本市佔"] = Convert.ToDouble(drw["MarketShare"].ToString());
                    dr["價內外起始"] = Convert.ToDouble(drw["MoneynessStart"].ToString());
                    dr["價內外長度"] = Convert.ToDouble(drw["MoneynessLength"].ToString());
                    dr["交易員"] = traderID;
                    dr["使用"] = drw["Used"].ToString() == "1" ? true : false;
                    dt.Rows.Add(dr);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($@"{ex.Message}");
            }
        }
        private void LoadAdj()
        {
            string traderID = comboBox1.Text;
            string sql = $@"SELECT  [TableType]
                                      ,[RowKey]
                                      ,[RowValue]
                                      ,[TraderID]
                                  FROM [WarrantAssistant].[dbo].[GeneralIssueSetting]
                                  WHERE [TraderID] = '{traderID}' AND [TableType] = 'TableADJ'";
            DataTable dtTemp = MSSQL.ExecSqlQry(sql, GlobalVar.loginSet.warrantassistant45);

            ADJ.Clear();
            dt_adj.Clear();
            foreach(DataRow drw in dtTemp.Rows)
            {
                string type = drw["RowKey"].ToString();
                double value = Convert.ToDouble(drw["RowValue"].ToString());
                DataRow dr = dt_adj.NewRow();
                dr["分類"] = type;
                dr["加減碼"] = value;
                dt_adj.Rows.Add(dr);
                ADJ.Add(type, value);
                //ADJ.Add(type, );
            }
        }
        private void LoadGlobalVar()
        {
           
            string traderID = comboBox1.Text;
            string sql = $@"SELECT  [TableType]
                                      ,[RowKey]
                                      ,[RowValue]
                                      ,[TraderID]
                                  FROM [WarrantAssistant].[dbo].[GeneralIssueSetting]
                                  WHERE [TraderID] = '{traderID}' AND [TableType] = 'TableSetting'";
            DataTable dtTemp = MSSQL.ExecSqlQry(sql, GlobalVar.loginSet.warrantassistant45);

            dt_globalvar.Clear();
            foreach(DataRow drw in dtTemp.Rows)
            {
                string type = drw["RowKey"].ToString();
                double value = Convert.ToDouble(drw["RowValue"].ToString());
                if (type == "全年交易日")
                    tradeDays = Convert.ToInt32(drw["RowValue"].ToString());
                if (type == "全年日曆日")
                    calendarDays = Convert.ToInt32(drw["RowValue"].ToString());
                if (type == "利率")
                    interestR = Convert.ToDouble(drw["RowValue"].ToString()) / 100;
                if (type == "期間(月)")
                    duration = Convert.ToInt32(drw["RowValue"].ToString());
                if (type == "標的VSP")
                    vspTotal = Convert.ToInt32(drw["RowValue"].ToString());
                if (type == "單檔VSP")
                    vspAvg = Convert.ToInt32(drw["RowValue"].ToString());
                DataRow dr = dt_globalvar.NewRow();
                dr["設定"] = type;
                dr["參數"] = value;
                dt_globalvar.Rows.Add(dr);
            }
           
        }
        private void LoadJumpSize()
        {
            string traderID = comboBox1.Text;
            string sql = $@"SELECT  [TableType]
                                      ,[RowKey]
                                      ,[RowValue]
                                      ,[TraderID]
                                  FROM [WarrantAssistant].[dbo].[GeneralIssueSetting]
                                  WHERE [TraderID] = '{traderID}' AND [TableType] = 'TableMoneyness'";
            DataTable dtTemp = MSSQL.ExecSqlQry(sql, GlobalVar.loginSet.warrantassistant45);

            dt_jumpsize.Clear();
            foreach (DataRow drw in dtTemp.Rows)
            {
                string type = drw["RowKey"].ToString();
                int index = type.IndexOf("_");
                double m1 = 99999;
                if(type.Substring(0, index) != "")
                    m1 = Convert.ToDouble(type.Substring(0, index));
                double m2 = -99999;
                if (type.Substring(index + 1, type.Length - index - 1) != "") 
                    m2 = Convert.ToDouble(type.Substring(index + 1, type.Length - index - 1));
                double value = Convert.ToInt32(drw["RowValue"].ToString());
                DataRow dr = dt_jumpsize.NewRow();
                dr["價內外起(含)"] = m1;
                dr["價內外終"] = m2;
                dr["跳動價差"] = value;
                dt_jumpsize.Rows.Add(dr);
            }
        }
        private void Select()
        {
            
            selectResult.Clear();
            string traderID = comboBox1.Text;
            using (StreamWriter sw = new StreamWriter("發行.txt"))
            {
                k_seperate.Clear();
                foreach (Infragistics.Win.UltraWinGrid.UltraGridRow dr in ultraGrid1.Rows)
                {

                    string type = dr.Cells["分類"].Value.ToString();
                   
                    string cp = dr.Cells["CP"].Value.ToString();
                    bool used = dr.Cells["使用"].Value == DBNull.Value ? false : Convert.ToBoolean(dr.Cells["使用"].Value);
                    double t2m = dr.Cells["TtoM"].Value == DBNull.Value ? -1 : Convert.ToDouble(dr.Cells["TtoM"].Value);
                    if (used == false)
                        continue;
                    //價內外起始
                    double moneyness_start = dr.Cells["價內外起始"].Value == DBNull.Value ? -1 : Convert.ToDouble(dr.Cells["價內外起始"].Value);
                    //價內外長度
                    double moneyness_length = dr.Cells["價內外長度"].Value == DBNull.Value ? -1 : Convert.ToDouble(dr.Cells["價內外長度"].Value);

                    double moneyness_end = moneyness_start - moneyness_length;
                   
                    string sql = $@"SELECT C.WID,C.UID,C.IssuerName,C.Moneyness FROM (
                                        SELECT  A.[WID]
		                                    ,A.[UID]
		                                    ,A.[IssuerName]
		                                    ,CASE WHEN A.[WClass] = 'c' THEN (A.StrikePrice / B.收盤價) - 1 ELSE 1-(A.StrikePrice / B.收盤價) END AS [Moneyness]
                                        FROM [TwData].[dbo].[V_WarrantTrading] AS A
                                        LEFT JOIN
	                                    (SELECT [股票代號]
                                          ,[收盤價]
                                      FROM [TwCMData].[dbo].[日收盤表排行]
                                      WHERE [日期] = '{lastTradeDate.ToString("yyyyMMdd")}' AND LEN([股票代號]) < 5 AND LEFT([股票代號], 2) <> '00') AS B ON A.[UID] = B.股票代號
                                      WHERE [TDate] = '{lastTradeDate.ToString("yyyyMMdd")}' AND [WClass] = '{cp}' AND LEN(UID) < 5 AND LEFT(UID, 2) <> '00' AND [TtoM] > {t2m}) AS C
                                      WHERE C.Moneyness >= {-moneyness_start / 100} AND C.Moneyness <= 0.5";
                    DataTable wTable = MSSQL.ExecSqlQry(sql, GlobalVar.loginSet.twData);

                    string sql2 = "";
                    if (cp == "c")
                    {
                        sql2 = $@"SELECT C.代號,C.券商代號,C.標的代號,C.Moneyness FROM(
                                  SELECT  A.[代號]
	                              ,A.[券商代號]
                                  ,A.[標的代號]
                                  ,(A.[最新履約價] / B.[收盤價]) -1 AS Moneyness 
                              FROM [TwCMData].[dbo].[Warrant總表] AS A
                                 LEFT JOIN
	                                    (SELECT [股票代號]
                                          ,[收盤價]
                                      FROM [TwCMData].[dbo].[日收盤表排行]
                                      WHERE [日期] = '{lastTradeDate.ToString("yyyyMMdd")}' AND LEN([股票代號]) < 5 AND LEFT([股票代號], 2) <> '00') AS B ON A.[標的代號] = B.[股票代號]
                              WHERE [日期] = '{lastTradeDate.ToString("yyyyMMdd")}' AND ([上市日期] >= '{lastTradeDate.ToString("yyyyMMdd")}' OR [上市日期] = '19110101') AND [名稱] LIKE '%購%' AND LEN(標的代號) < 5 AND LEFT(標的代號, 2) <> '00') AS C
                              WHERE C.Moneyness > {-moneyness_start / 100} AND C.Moneyness <= 0.5";
                    }
                    else if (cp == "p")
                    {
                        sql2 = $@"SELECT C.代號,C.券商代號,C.標的代號,C.Moneyness FROM(
                                  SELECT  A.[代號]
	                              ,A.[券商代號]
                                  ,A.[標的代號]
                                  ,1 - (A.[最新履約價] / B.[收盤價]) AS Moneyness 
                              FROM [TwCMData].[dbo].[Warrant總表] AS A
                               LEFT JOIN
	                                    (SELECT [股票代號]
                                          ,[收盤價]
                                      FROM [TwCMData].[dbo].[日收盤表排行]
                                      WHERE [日期] = '{lastTradeDate.ToString("yyyyMMdd")}' AND LEN([股票代號]) < 5 AND LEFT([股票代號], 2) <> '00') AS B ON A.[標的代號] = B.[股票代號]
                              WHERE [日期] = '{lastTradeDate.ToString("yyyyMMdd")}' AND ([上市日期] >= '{lastTradeDate.ToString("yyyyMMdd")}' OR [上市日期] = '19110101') AND [名稱] LIKE '%售%' AND LEN(標的代號) < 5 AND LEFT(標的代號, 2) <> '00') AS C
                                WHERE C.Moneyness > {-moneyness_start / 100} AND C.Moneyness <= 0.5"; 
                    }
                    else
                    {
                        MessageBox.Show("CP錯誤");
                        continue;
                    }
                    DataTable wTable_unlisted = MSSQL.ExecSqlQry(sql2, GlobalVar.loginSet.twCMData);

                    double marketshare_basic = dr.Cells["基本市佔"].Value == DBNull.Value ? -1 : Convert.ToDouble(dr.Cells["基本市佔"].Value);
                    if (marketshare_basic < 0)
                        MessageBox.Show("基本市佔數字怪怪的");
                    DataRow[] uid_classifySelect = uid_classify.Select($@"[WClass] = '{cp}' AND [class] = '{type}'");
                   
                    try
                    {
                        if (uid_classifySelect.Length > 0)
                        {
                            foreach (DataRow dr2 in uid_classifySelect)
                            {
                                string uid = dr2[0].ToString();
                                if (Issuable_CP[cp].ContainsKey(uid))
                                {
                                    if (Issuable_CP[cp][uid] == "N")
                                        continue;
                                }
                                else
                                    continue;

                                if (!underlying_issueCredit.ContainsKey(uid))
                                    continue;

                                //C2及D若流動性分級為C則不發
                                if (type == "C2" && liquidityC.Contains(uid))
                                    continue;

                                if (type == "D" && liquidityC.Contains(uid))
                                    continue;
                                double uid_vspTotal = Convert.ToDouble(dr2[4].ToString());
                                double uid_vspAvg = Convert.ToDouble(dr2[5].ToString());
                                if (uid_vspTotal < vspTotal || uid_vspAvg < vspAvg)
                                    continue;
                                ///////////////////////////////
                                //用最小行使比例檢查是否有額度
                                double test_vol = -1;
                                if (suggestVol_CP[cp].ContainsKey(uid)) 
                                    test_vol = suggestVol_CP[cp][uid];
                                if (test_vol <= 0)
                                    test_vol = underlying_hv60[uid];
                                double test_cr = 0;
                                double test_p = 0;
                                if (test_vol >= 0)
                                {
                                    if (uid == "IX0001" || uid == "IX0027")
                                    {
                                        if (cp == "c")
                                            test_p = Math.Abs(EDLib.Pricing.Option.PlainVanilla.CallPrice(ucloseP[uid], ucloseP[uid], interestR_Index, test_vol / 100, (double)duration / 12));
                                        if (cp == "p")
                                            test_p = Math.Abs(EDLib.Pricing.Option.PlainVanilla.PutPrice(ucloseP[uid], ucloseP[uid], interestR_Index, test_vol / 100, (double)duration / 12));
                                    }
                                    else
                                    {
                                        if (cp == "c")
                                            test_p = Math.Abs(EDLib.Pricing.Option.PlainVanilla.CallPrice(ucloseP[uid], ucloseP[uid], interestR, test_vol / 100, (double)duration / 12));
                                        if (cp == "p")
                                            test_p = Math.Abs(EDLib.Pricing.Option.PlainVanilla.PutPrice(ucloseP[uid], ucloseP[uid], interestR, test_vol / 100, (double)duration / 12));
                                    }
                                    test_cr = Math.Round(0.6 / test_p, 3);
                                    if (test_cr * 5000 > underlying_issueCredit[uid] && test_cr * 5000 > underlying_issueCreditReward[uid])
                                        continue;
                                }
                                ///////////////////////////////////////////
                                string k_seperate_key = uid + "_" + cp;
                                if (!k_seperate.ContainsKey(k_seperate_key))
                                {
                                    k_seperate.Add(k_seperate_key, new Dictionary<string, int>());
                                }

                                if (!underlying_trader[traderID.TrimStart('0')].Contains(uid))
                                    continue;
                                string margin_type = dr2[3].ToString();

                                //double marketshare_target = marketshare_basic + ADJ[margin_type];
                                double marketshare_target = Convert.ToDouble(dr2[6].ToString());
                                
                                DataRow[] wTableSelect = wTable.Select($@"UID = '{uid}'");
                                DataRow[] wTableSelect_KGI = wTable.Select($@"UID = '{uid}' AND IssuerName = '9200'");
                                DataRow[] wTable_unlistedSelect = wTable_unlisted.Select($@"標的代號 = '{uid}'");
                                DataRow[] wTable_unlistedSelect_KGI = wTable_unlisted.Select($@"標的代號 = '{uid}' AND 券商代號 = '9200'");
                                int wAllCount = wTableSelect.Length + wTable_unlistedSelect.Length;
                                int wKgiCount = wTableSelect_KGI.Length + wTable_unlistedSelect_KGI.Length;
                                //應該要發幾檔
                                double targetCount = Math.Ceiling((double)wAllCount * marketshare_target / 100);
                                //double moneyness_unit = moneyness_length / targetCount;
                                //基本檔數
                                double basicCount = Math.Round((double)wAllCount * marketshare_target / 100, 1) < 1 ? 1 : Math.Round((double)wAllCount * marketshare_target / 100, 1);
 
                                double moneyness_unit = moneyness_length / basicCount;
                                //sw.Write($@"{uid},{type},{margin_type},{cp},{Math.Round(moneyness_unit, 2)}");
                                for (int i = 0; i < targetCount; i++)
                                {
                                    double m1 = moneyness_start - i * Math.Round(moneyness_unit, 2);
                                    
                                    double m2 = moneyness_start - (i + 1) * Math.Round(moneyness_unit, 2);
                                    
                                    DataRow[] wTableSelect_KGI_M = wTable.Select($@"UID = '{uid}' AND IssuerName = '9200' AND Moneyness >= {-Math.Round(m1 / 100, 3)} AND Moneyness < {Math.Round(-m2 / 100, 3)}");
                                   
                                    DataRow[] wTable_unlistedSelect_KGI_M = wTable_unlisted.Select($@"標的代號 = '{uid}' AND 券商代號 = '9200' AND Moneyness >= {-Math.Round(m1 / 100, 3)} AND Moneyness < {Math.Round(-m2 / 100, 3)}");
                                  
                                    int c = wTableSelect_KGI_M.Length + wTable_unlistedSelect_KGI_M.Length;
                                
                                    if (cp == "c")
                                        k_seperate[k_seperate_key].Add(roundpriceK(ucloseP[uid] * (1 - m1 / 100)).ToString() + "-" + roundpriceK(ucloseP[uid] * (1 - m2 / 100)).ToString(), c);
                                    if (cp == "p")
                                        k_seperate[k_seperate_key].Add(roundpriceK(ucloseP[uid] * (1 + m1 / 100)).ToString() + "-" + roundpriceK(ucloseP[uid] * (1 + m2 / 100)).ToString(), c);
                                    //這個區間沒有權證
                                    if (c <= 0)
                                    {
                                        double suggest_moneyness = (m1 + m2) / 2;
                                        double suggest_k = (cp == "c") ? roundpriceK(ucloseP[uid] * (1 - suggest_moneyness / 100)) : roundpriceK(ucloseP[uid] * (1 + suggest_moneyness / 100));
                                        
                                        double wJumpSize = 0;

                                        foreach (Infragistics.Win.UltraWinGrid.UltraGridRow dr4 in ultraGrid4.Rows)
                                        {
                                            double m3 = Convert.ToDouble(dr4.Cells["價內外起(含)"].Value.ToString());
                                            double m4 = Convert.ToDouble(dr4.Cells["價內外終"].Value.ToString());
                                            double j = Convert.ToDouble(dr4.Cells["跳動價差"].Value.ToString());
                                            if (suggest_moneyness <= m3 && suggest_moneyness > m4)
                                                wJumpSize = Math.Round(j / 100, 2);
                                        }

                                        double suggest_vol = -1;
                                        if (suggestVol_CP[cp].ContainsKey(uid))
                                            suggest_vol = suggestVol_CP[cp][uid];


                                        double delta = 0;
                                        double suggest_cr = 0;
                                        if (suggest_vol > 0)
                                        {
                                            if (uid == "IX0001" || uid == "IX0027")
                                            {
                                                if (cp == "c")
                                                    delta = Math.Abs(EDLib.Pricing.Option.PlainVanilla.CallDelta(ucloseP[uid], suggest_k, interestR_Index, suggest_vol / 100, (double)duration / 12));
                                                if (cp == "p")
                                                    delta = Math.Abs(EDLib.Pricing.Option.PlainVanilla.PutDelta(ucloseP[uid], suggest_k, interestR_Index, suggest_vol / 100, (double)duration / 12));
                                            }
                                            else
                                            {
                                                if (cp == "c")
                                                    delta = Math.Abs(EDLib.Pricing.Option.PlainVanilla.CallDelta(ucloseP[uid], suggest_k, interestR, suggest_vol / 100, (double)duration / 12));
                                                if (cp == "p")
                                                    delta = Math.Abs(EDLib.Pricing.Option.PlainVanilla.PutDelta(ucloseP[uid], suggest_k, interestR, suggest_vol / 100, (double)duration / 12));
                                            }
                                            suggest_cr = Math.Round((wJumpSize / delta) / uJumpSize(ucloseP[uid]), 3);
                                            //suggest_cr = (wJumpSize / delta) / uJumpSize(ucloseP[uid]);
                                        }
                                        //sw.WriteLine($@"{uid} CP:{cp} 基本市佔:{marketshare_basic} 毛利率分類:{margin_type} 調整後市佔:{marketshare_target} 價內{m1}% 價外{m2}% 檔數:{c} 應該發");
                                        DataRow drNew = selectResult.NewRow();
                                        drNew["標的代號"] = uid;
                                        drNew["標的名稱"] = uName[uid];
                                        drNew["KGI檔數"] = wKgiCount;
                                        drNew["總檔數"] = wAllCount;
                                        drNew["基本檔數"] = Math.Round((double)wAllCount * marketshare_target / 100, 1);
                                        drNew["建議Moneyness間隔"] = Math.Round(moneyness_unit, 2);
                                        //drNew["格子數"] = targetCount;
                                        drNew["檔數市佔"] = Math.Round((double)wKgiCount / (double)wAllCount * 100, 2);
                                        drNew["目標市佔"] = marketshare_target;
                                        if (ThetaShare_CP[cp].ContainsKey(uid))
                                            drNew["Theta市佔"] = ThetaShare_CP[cp][uid];
                                        drNew["建議Vol"] = Math.Round(suggest_vol, 1);
                                        if (underlying_hv60.ContainsKey(uid))
                                            drNew["HV"] = Math.Round(underlying_hv60[uid], 1);
                                        else
                                            drNew["HV"] = -1;

                                        drNew["標的收盤價"] = ucloseP[uid];
                                        drNew["Theta分類"] = type;
                                        drNew["Margin分類"] = margin_type;
                                        drNew["CP"] = cp;
                                        drNew["推薦價內外"] = Math.Round(suggest_moneyness, 2);
                                        //drNew["現股跳動價差"] = uJumpSize(ucloseP[uid]);
                                        drNew["履約價"] = suggest_k;
                                        drNew["權證跳動價差"] = wJumpSize;
                                        drNew["行使比例"] = suggest_cr;
                                        drNew["標的VSP"] = uid_vspTotal;
                                        drNew["單檔VSP"] = uid_vspAvg;
                                        if (suggest_vol > 0)
                                        {
                                            if (uid == "IX0001" || uid == "IX0027")
                                            {
                                                if (cp == "c")
                                                    drNew["試算價格"] = Math.Round(EDLib.Pricing.Option.PlainVanilla.CallPrice(ucloseP[uid], suggest_k, interestR_Index, suggest_vol / 100, (double)duration / 12) * suggest_cr, 2);
                                                if (cp == "p")
                                                    drNew["試算價格"] = Math.Round(EDLib.Pricing.Option.PlainVanilla.PutPrice(ucloseP[uid], suggest_k, interestR_Index, suggest_vol / 100, (double)duration / 12) * suggest_cr, 2);
                                            }
                                            else
                                            {
                                                if (cp == "c")
                                                    drNew["試算價格"] = Math.Round(EDLib.Pricing.Option.PlainVanilla.CallPrice(ucloseP[uid], suggest_k, interestR, suggest_vol / 100, (double)duration / 12) * suggest_cr, 2);
                                                if (cp == "p")
                                                    drNew["試算價格"] = Math.Round(EDLib.Pricing.Option.PlainVanilla.PutPrice(ucloseP[uid], suggest_k, interestR, suggest_vol / 100, (double)duration / 12) * suggest_cr, 2);
                                            }
                                            drNew["HV倍數"] = Math.Round(suggest_vol / underlying_hv60[uid], 2);
                                        }
                                        else
                                        {
                                            drNew["試算價格"] = -1;
                                            drNew["HV倍數"] = -1;
                                        }

                                        drNew["額度試算"] = Math.Ceiling(5000 * suggest_cr);
                                        drNew["已發行"] = underlying_issuePercent[uid];
                                        drNew["額度"] = underlying_issueCredit[uid];
                                        if (underlying_issueCreditReward.ContainsKey(uid))
                                            drNew["獎勵額度"] = underlying_issueCreditReward[uid];
                                        else
                                            drNew["獎勵額度"] = 0;
                                        drNew["確認"] = false;
                                        selectResult.Rows.Add(drNew);
                                    }
                                }

                            }
                        }
                    }
                    catch(Exception ex)
                    {
                        MessageBox.Show($@"{ex.Message}");
                    }

                }
            }
            MessageBox.Show("篩選完成!");
        }

        private void button_Select_Click(object sender, EventArgs e)
        {
            Select();
        }

        private void button_Confirm_Click(object sender, EventArgs e)
        {
            string trader = comboBox1.Text;
            MSSQL.ExecSqlQry($@"DELETE FROM [WarrantAssistant].[dbo].[GeneralIssueGroup]
                                WHERE [TraderID] = '{trader}'", GlobalVar.loginSet.warrantassistant45);
            foreach (Infragistics.Win.UltraWinGrid.UltraGridRow dr in ultraGrid1.Rows)
            {
                string groupNum = dr.Cells["組別"].Value.ToString();
                string type = dr.Cells["分類"].Value.ToString();
                string cp = dr.Cells["CP"].Value.ToString();
                //bool used = dr.Cells["使用"].Value == DBNull.Value ? false : Convert.ToBoolean(dr.Cells["使用"].Value);
                double t2m = dr.Cells["TtoM"].Value == DBNull.Value ? -1 : Convert.ToDouble(dr.Cells["TtoM"].Value);
                double marketShare = dr.Cells["基本市佔"].Value == DBNull.Value ? 0 : Convert.ToDouble(dr.Cells["基本市佔"].Value);
                double moneynessLength = dr.Cells["價內外長度"].Value == DBNull.Value ? 0 : Convert.ToDouble(dr.Cells["價內外長度"].Value);
                double moneynessStart = dr.Cells["價內外起始"].Value == DBNull.Value ? 0 : Convert.ToDouble(dr.Cells["價內外起始"].Value);
                string traderID = dr.Cells["交易員"].Value.ToString();
                string used = dr.Cells["使用"].Value.ToString() == "False" ? "0" : "1";
                string sqlInsert = $@"INSERT INTO [WarrantAssistant].[dbo].[GeneralIssueGroup] ([GroupNum], [Rank], [WClass], [TtoM], [MarketShare], [MoneynessLength], [MoneynessStart], [TraderID], [Used])
                                                                                                VALUES({groupNum}, '{type}', '{cp}', {t2m}, {marketShare}, {moneynessLength}, {moneynessStart}, '{traderID}', '{used}')";
                MSSQL.ExecSqlQry(sqlInsert, GlobalVar.loginSet.warrantassistant45);
            }
            MSSQL.ExecSqlQry($@"DELETE FROM [WarrantAssistant].[dbo].[GeneralIssueSetting]
                                WHERE [TableType] = 'TableADJ' AND [TraderID] = '{trader}'", GlobalVar.loginSet.warrantassistant45);
            foreach (Infragistics.Win.UltraWinGrid.UltraGridRow dr in ultraGrid2.Rows)
            {
                
                string type = dr.Cells["分類"].Value.ToString();
                string adj = dr.Cells["加減碼"].Value.ToString();
                string traderID = comboBox1.Text; 

                string sqlInsert = $@"INSERT INTO [WarrantAssistant].[dbo].[GeneralIssueSetting] ([TableType], [RowKey], [RowValue], [TraderID])
                                                                                                VALUES('TableADJ', '{type}', '{adj}', '{traderID}')";
                MSSQL.ExecSqlQry(sqlInsert, GlobalVar.loginSet.warrantassistant45);
            }
            MSSQL.ExecSqlQry($@"DELETE FROM [WarrantAssistant].[dbo].[GeneralIssueSetting]
                                WHERE [TableType] = 'TableSetting' AND [TraderID] = '{trader}'", GlobalVar.loginSet.warrantassistant45);
            foreach (Infragistics.Win.UltraWinGrid.UltraGridRow dr in ultraGrid3.Rows)
            {

                string key = dr.Cells["設定"].Value.ToString();
                string value = dr.Cells["參數"].Value.ToString();
                string traderID = comboBox1.Text;

                string sqlInsert = $@"INSERT INTO [WarrantAssistant].[dbo].[GeneralIssueSetting] ([TableType], [RowKey], [RowValue], [TraderID])
                                                                                                VALUES('TableSetting', '{key}', '{value}', '{traderID}')";
                MSSQL.ExecSqlQry(sqlInsert, GlobalVar.loginSet.warrantassistant45);
            }
            MSSQL.ExecSqlQry($@"DELETE FROM [WarrantAssistant].[dbo].[GeneralIssueSetting]
                                WHERE [TableType] = 'TableMoneyness' AND [TraderID] = '{trader}'", GlobalVar.loginSet.warrantassistant45);
            foreach (Infragistics.Win.UltraWinGrid.UltraGridRow dr in ultraGrid4.Rows)
            {
                string key = "";
                double m1 = Convert.ToDouble(dr.Cells["價內外起(含)"].Value.ToString());
                if (m1 < 99999)
                    key += m1.ToString()+"_";
                double m2 = Convert.ToDouble(dr.Cells["價內外終"].Value.ToString());
                if (m2 > -99999)
                    key += m2.ToString();
                string value = dr.Cells["跳動價差"].Value.ToString();
                string traderID = comboBox1.Text;

                string sqlInsert = $@"INSERT INTO [WarrantAssistant].[dbo].[GeneralIssueSetting] ([TableType], [RowKey], [RowValue], [TraderID])
                                                                                                VALUES('TableMoneyness', '{key}', '{value}', '{traderID}')";
                MSSQL.ExecSqlQry(sqlInsert, GlobalVar.loginSet.warrantassistant45);
            }
            LoadAll();
            MessageBox.Show("OK!");
        }

        private void UltraGrid5_ClickCell(object sender, ClickCellEventArgs e)
        {
            //MessageBox.Show("QQ");
            string uid = e.Cell.Row.Cells[0].Value.ToString();
            string cp = e.Cell.Row.Cells[16].Value.ToString();
            string key = uid + "_" + cp;
            if (k_seperate.ContainsKey(key))
            {
                DataTable dt = new DataTable();
                foreach(string k in k_seperate[key].Keys)
                {
                    dt.Columns.Add(k, typeof(int));
                }
                DataRow dr = dt.NewRow();
                foreach (string k in k_seperate[key].Keys)
                {
                    dr[k] = k_seperate[key][k];
                }
                dt.Rows.Add(dr);
                ultraGrid6.DataSource = dt;
            }
        }

        private void button_Issue_Click(object sender, EventArgs e)
        {
            //string traderID = comboBox1.Text;
           
            string sql = @"INSERT INTO [ApplyTempList] (SerialNum, UnderlyingID, K, T, R, HV, IV, IssueNum, ResetR, BarrierR, FinancialR, Type, CP, UseReward, ConfirmChecked, Apply1500W, UserID, MDate, TempName, TempType, TraderID, IVNew, Adj) "
                    + "VALUES(@SerialNum, @UnderlyingID, @K, @T, @R, @HV, @IV, @IssueNum, @ResetR, @BarrierR, @FinancialR, @Type, @CP, @UseReward, @ConfirmChecked, @Apply1500W, @UserID, @MDate, @TempName ,@TempType, @TraderID, @IVNew, @Adj)";
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
            new SqlParameter("@Adj", SqlDbType.Float)};

            SQLCommandHelper h = new SQLCommandHelper(GlobalVar.loginSet.warrantassistant45, sql, ps);
            
            try
            {
                int deleteNumMax = 0;
                if (DeletedSerialNum.Count > 0)
                    deleteNumMax = DeletedSerialNum[DeletedSerialNum.Count - 1];
                
                string sqlTempListStr = $@"SELECT MAX(A.SerialNum) AS MaxNum FROM
                                        (SELECT CAST(SUBSTRING([SerialNum],18,LEN([SerialNum])-17) AS INT) AS SerialNum
                                          FROM [WarrantAssistant].[dbo].[ApplyTempList]
                                          WHERE [MDate] > '{DateTime.Today.ToString("yyyyMMdd")}' AND [UserID] ='{userID}') AS A";

                int tempListNumMax = 0;
                DataTable dtTempList = MSSQL.ExecSqlQry(sqlTempListStr, GlobalVar.loginSet.warrantassistant45);
                if (dtTempList.Rows.Count > 0)
                {
                    if(dtTempList.Rows[0][0].ToString() !="")
                      tempListNumMax = Convert.ToInt32(dtTempList.Rows[0][0].ToString());
                }
                
                string sqlOfficialStr = $@"SELECT MAX(A.SerialNum) AS MaxNum FROM
                                        (SELECT CAST(SUBSTRING([SerialNumber],18,LEN([SerialNumber])-17) AS INT) AS SerialNum
                                          FROM [WarrantAssistant].[dbo].[ApplyOfficial]
                                          WHERE [MDate] > '{DateTime.Today.ToString("yyyyMMdd")}' AND [UserID] ='{userID}') AS A";
                int officialNumMax = 0;
                DataTable dtOfficialList = MSSQL.ExecSqlQry(sqlOfficialStr, GlobalVar.loginSet.warrantassistant45);
                if (dtOfficialList.Rows.Count > 0)
                {
                    if(dtOfficialList.Rows[0][0].ToString()!="")
                        officialNumMax = Convert.ToInt32(dtOfficialList.Rows[0][0].ToString());
                }
                
                int max = 0;
                if (deleteNumMax >= tempListNumMax)
                    max = deleteNumMax;
                else
                    max = tempListNumMax;
                if (max <= officialNumMax)
                    max = officialNumMax;
               
                int i = 1;
                foreach (Infragistics.Win.UltraWinGrid.UltraGridRow dr in ultraGrid5.Rows)
                {
                    string confirm = dr.Cells["確認"].Value.ToString() == "False" ? "0" : "1";
                    if (confirm == "0")
                        continue;
                    string underlyingID = Convert.ToString(dr.Cells["標的代號"].Value);
                    string traderID = "";
                    try
                    {
                        traderID = underlying2trader[underlyingID];
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show($@"{underlyingID}沒有對照交易員");
                        continue;
                    }
                    string serialNumber = DateTime.Today.ToString("yyyyMMdd") + userID + "01" + (max + i).ToString("0#");
                    double k = Convert.ToDouble(dr.Cells["履約價"].Value);
                    
                    int t = duration;
                    double cr = Convert.ToDouble(dr.Cells["行使比例"].Value);
                    double hv = Convert.ToDouble(dr.Cells["建議Vol"].Value);
                    double iv = Convert.ToDouble(dr.Cells["建議Vol"].Value);
                    double issueNum = 5000;
                    double resetR = 0;
                    double barrierR = 0;
                    double financialR = 0;
                    double adj = 0;
                    string type = "一般型";
                    string cp = Convert.ToString(dr.Cells["CP"].Value) == "c" ? "C" : "P";
                    string underlyingName = Convert.ToString(dr.Cells["標的名稱"].Value);
                    double underlyingPrice = Convert.ToDouble(dr.Cells["標的收盤價"].Value);
                    string useReward = "N";
                    string confirmChecked = "N";
                    string apply1500W = "N";
                    
                    DateTime expiryDate = GlobalVar.globalParameter.nextTradeDate3.AddMonths(t);


                    if (expiryDate.Day == GlobalVar.globalParameter.nextTradeDate3.Day)
                        expiryDate = expiryDate.AddDays(-1);
                    string sqlTemp = $"SELECT TOP 1 TradeDate from TradeDate WHERE IsTrade='Y' AND TradeDate >= '{expiryDate.ToString("yyyy-MM-dd")}'";

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
                    h.ExecuteCommand();
                    i++;
                }
            }
            catch(Exception ex)
            {
                MessageBox.Show($@"{ex.Message}");
            }
            h.Dispose();
            MessageBox.Show("發行完成!");
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            
            LoadPara();
            LoadAdj();
            LoadGlobalVar();
            LoadJumpSize();
            
        }


        private void UltraGrid5_AfterCellUpdate(object sender, CellEventArgs e)
        {
            try
            {
                if (e.Cell.Column.Key == "行使比例" || e.Cell.Column.Key == "履約價" || e.Cell.Column.Key == "建議Vol")
                {
                    double price = 0.0;
                    double delta = 0.0;
                    double theta = 0.0; //joufan
                    double jumpSize = 0.0;
                    double multiplier = 0.0;

                    double underlyingPrice = e.Cell.Row.Cells["標的收盤價"].Value == DBNull.Value ? 0 : Convert.ToDouble(e.Cell.Row.Cells["標的收盤價"].Value);

                    double k = e.Cell.Row.Cells["履約價"].Value == DBNull.Value ? 0 : Convert.ToDouble(e.Cell.Row.Cells["履約價"].Value);
                    double cr = e.Cell.Row.Cells["行使比例"].Value == DBNull.Value ? 0 : Convert.ToDouble(e.Cell.Row.Cells["行使比例"].Value);
                    double vol = e.Cell.Row.Cells["建議Vol"].Value == DBNull.Value ? 0 : Convert.ToDouble(e.Cell.Row.Cells["建議Vol"].Value);
                    double hv = e.Cell.Row.Cells["HV"].Value == DBNull.Value ? 0 : Convert.ToDouble(e.Cell.Row.Cells["HV"].Value);
                    string cpType = e.Cell.Row.Cells["CP"].Value.ToString();

                    string uid = e.Cell.Row.Cells["標的代號"].Value.ToString();


                    if (vol >= 0)
                    {
                        if (uid == "IX0001" || uid == "IX0027")
                        {
                            if (cpType == "c")
                            {
                                e.Cell.Row.Cells["試算價格"].Value = Math.Round(EDLib.Pricing.Option.PlainVanilla.CallPrice(underlyingPrice, k, interestR_Index, vol / 100, (double)duration / 12) * cr, 2);
                                delta = Math.Abs(EDLib.Pricing.Option.PlainVanilla.CallDelta(underlyingPrice, k, interestR_Index, vol / 100, (double)duration / 12));
                            }
                            if (cpType == "p")
                            {
                                e.Cell.Row.Cells["試算價格"].Value = Math.Round(EDLib.Pricing.Option.PlainVanilla.PutPrice(underlyingPrice, k, interestR_Index, vol / 100, (double)duration / 12) * cr, 2);
                                delta = Math.Abs(EDLib.Pricing.Option.PlainVanilla.PutDelta(underlyingPrice, k, interestR_Index, vol / 100, (double)duration / 12));
                            }
                        }
                        else
                        {
                            if (cpType == "c")
                            {
                                e.Cell.Row.Cells["試算價格"].Value = Math.Round(EDLib.Pricing.Option.PlainVanilla.CallPrice(underlyingPrice, k, interestR, vol / 100, (double)duration / 12) * cr, 2);
                                delta = Math.Abs(EDLib.Pricing.Option.PlainVanilla.CallDelta(underlyingPrice, k, interestR, vol / 100, (double)duration / 12));
                            }
                            if (cpType == "p")
                            {
                                e.Cell.Row.Cells["試算價格"].Value = Math.Round(EDLib.Pricing.Option.PlainVanilla.PutPrice(underlyingPrice, k, interestR, vol / 100, (double)duration / 12) * cr, 2);
                                delta = Math.Abs(EDLib.Pricing.Option.PlainVanilla.PutDelta(underlyingPrice, k, interestR, vol / 100, (double)duration / 12));
                            }
                        }
                        e.Cell.Row.Cells["權證跳動價差"].Value = Math.Round(uJumpSize(underlyingPrice) * cr * delta, 3);
                        e.Cell.Row.Cells["HV倍數"].Value = Math.Round(vol / hv, 2);
                    }
                    else
                    {
                        e.Cell.Row.Cells["試算價格"].Value = -1;
                        e.Cell.Row.Cells["HV倍數"].Value = -1;
                    }
                    e.Cell.Row.Cells["額度試算"].Value = Math.Ceiling(5000 * cr);

                    if (cpType == "c")
                    {
                        e.Cell.Row.Cells["推薦價內外"].Value = Math.Round(-(k / underlyingPrice - 1) * 100, 2);
                    }
                    if (cpType == "p")
                    {
                        e.Cell.Row.Cells["推薦價內外"].Value = Math.Round((k / underlyingPrice - 1) * 100, 2);
                    }

                    

                }
            }
            catch (Exception ex4)
            {
                MessageBox.Show("ex4" + ex4.Message);
            }
        }


        
        private void UltraGrid5_InitializeRow(object sender, InitializeRowEventArgs e)
        {
            try
            {
                
                double cr = e.Row.Cells["行使比例"].Value == DBNull.Value ? 0 : Convert.ToDouble(e.Row.Cells["行使比例"].Value);
                double credit = e.Row.Cells["額度"].Value == DBNull.Value ? 0 : Convert.ToDouble(e.Row.Cells["額度"].Value);
                double rewardcredit = e.Row.Cells["獎勵額度"].Value == DBNull.Value ? 0 : Convert.ToDouble(e.Row.Cells["獎勵額度"].Value);
                double creditcal = e.Row.Cells["額度試算"].Value == DBNull.Value ? 0 : Convert.ToDouble(e.Row.Cells["額度試算"].Value);
                double vol = e.Row.Cells["建議Vol"].Value == DBNull.Value ? 0.0 : Convert.ToDouble(e.Row.Cells["建議Vol"].Value);
                double hv = e.Row.Cells["HV"].Value == DBNull.Value ? 0.0 : Convert.ToDouble(e.Row.Cells["HV"].Value);
                double wprice = e.Row.Cells["試算價格"].Value == DBNull.Value ? 0.0 : Convert.ToDouble(e.Row.Cells["試算價格"].Value);
                double issuedPercent = e.Row.Cells["已發行"].Value == DBNull.Value ? 0.0 : Convert.ToDouble(e.Row.Cells["已發行"].Value);

                if (vol <= 0)
                {
                    e.Row.Cells["建議Vol"].Appearance.BackColor = Color.LightCoral;
                    e.Row.Cells["HV倍數"].Appearance.BackColor = Color.LightCoral;
                }
                else
                {
                    e.Row.Cells["建議Vol"].Appearance.BackColor = Color.Aquamarine;
                    e.Row.Cells["HV倍數"].Appearance.BackColor = Color.Wheat;
                }
                if (wprice <= 0.6)
                    e.Row.Cells["試算價格"].Appearance.BackColor = Color.LightCoral;
                else
                    e.Row.Cells["試算價格"].Appearance.BackColor = Color.Wheat;
                if (creditcal > credit)
                {
                    e.Row.Cells["額度"].Appearance.BackColor = Color.LightCoral;
                    e.Row.Cells["額度試算"].Appearance.BackColor = Color.LightCoral;
                }
                else
                {
                    e.Row.Cells["額度"].Appearance.BackColor = Color.Wheat;
                    e.Row.Cells["額度試算"].Appearance.BackColor = Color.Wheat;
                }
                if (creditcal > rewardcredit)
                {
                    e.Row.Cells["獎勵額度"].Appearance.BackColor = Color.LightCoral;
                }
                else
                {
                    e.Row.Cells["獎勵額度"].Appearance.BackColor = Color.Wheat;
                }
                if (cr > 1 || cr == 0)
                    e.Row.Cells["行使比例"].Appearance.BackColor = Color.LightCoral;
                else
                    e.Row.Cells["行使比例"].Appearance.BackColor = Color.Aquamarine;
                if(hv < 0)
                    e.Row.Cells["HV"].Appearance.BackColor = Color.LightCoral;
                if(issuedPercent > 18)
                    e.Row.Cells["已發行"].Appearance.BackColor = Color.LightCoral;
            }
            catch (Exception e1)
            {
                MessageBox.Show($@"{e1.Message}");
            }
        }
        
    }   
}

using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
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
    public partial class FrmGeneralIssueV2 : Form
    {

        private DataTable dt = new DataTable();


        //前一日Call好分點買進金額比例
        private Dictionary<string, double> Call_GoodAmtRatio_LastDay = new Dictionary<string, double>();
        //前一日Put好分點買進金額比例
        private Dictionary<string, double> Put_GoodAmtRatio_LastDay = new Dictionary<string, double>();
        //累積五日Call好分點買進金額比例
        private Dictionary<string, double> Call_GoodAmtRatio_Last5Day = new Dictionary<string, double>();
        //累積五日Put好分點買進金額比例
        private Dictionary<string, double> Put_GoodAmtRatio_Last5Day = new Dictionary<string, double>();
        //現股3日累積漲跌幅
        private Dictionary<string, double> UID_PriceUpDown = new Dictionary<string, double>();
        //現股開盤參考價，計算第四日漲跌幅
        private Dictionary<string, double> UID_RefPrice = new Dictionary<string, double>();


        private Dictionary<string, double> OptionCanIssue = new Dictionary<string, double>();

        public DateTime lastTradeDate = EDLib.TradeDate.LastNTradeDate(1);
        public DateTime last3TradeDate = EDLib.TradeDate.LastNTradeDate(3);
        public DateTime last5TradeDate = EDLib.TradeDate.LastNTradeDate(5);
        public FrmGeneralIssueV2()
        {
            InitializeComponent();
        }

        private void FrmGeneralIssueV2_Load(object sender, EventArgs e)
        {
            InitialGrid();
            LoadData();
        }

        private void InitialGrid()
        {
            dt.Columns.Add("標的代號-名稱", typeof(string));
            dt.Columns.Add("WClass", typeof(string));
            dt.Columns.Add("分級", typeof(string));
            dt.Columns.Add("Theta金額", typeof(double));
            dt.Columns.Add("平均VSP", typeof(double));
            dt.Columns.Add("檔數市佔", typeof(double));
            dt.Columns.Add("Theta市佔率", typeof(double));
            dt.Columns.Add("股價累計漲幅", typeof(string));
            dt.Columns.Add("好分點買進", typeof(string));
            dt.Columns.Add("搶發標的", typeof(string));
            dt.Columns.Add("價外25以下", typeof(double));
            dt.Columns.Add("價外25以上", typeof(double));


            ultraGrid1.DataSource = dt;
            UltraGridBand band0 = ultraGrid1.DisplayLayout.Bands[0];
            /*
            band0.Columns["標的代號-名稱"].Format = "N0";
            band0.Columns["WClass"].Format = "N0";
            band0.Columns["分級"].Format = "N0";
            band0.Columns["Theta金額"].Format = "N0";
            band0.Columns["平均VSP"].Format = "N0";
            band0.Columns["檔數市佔"].Format = "N0";
            band0.Columns["Theta市佔率"].Format = "N0";
            band0.Columns["股價累計漲幅"].Format = "N0";
            band0.Columns["好分點買進"].Format = "N0";
            band0.Columns["搶發標的"].Format = "N0";
            band0.Columns["價外25以下"].Format = "N0";
            band0.Columns["價外25以上"].Format = "N0";
            */
            band0.Columns["Theta金額"].Format = "N0";
            band0.Columns["平均VSP"].Format = "N0";
            // band0.Columns["類型"].Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownList;
            //band0.Columns["CP"].Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownList;
            //band0.Columns["交易員"].Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownList;
            //band0.Columns["發行原因"].Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownList;
            //ultraGrid1.DisplayLayout.Bands[0].Columns["確認"].Style = Infragistics.Win.UltraWinGrid.ColumnStyle.CheckBox;
            //ultraGrid1.DisplayLayout.Bands[0].Columns["刪除"].Style = Infragistics.Win.UltraWinGrid.ColumnStyle.CheckBox;

            //ultraGrid1.DisplayLayout.Bands[0].Columns["編號"].Width = 75;
            /*
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
            */


            // To sort multi-column using SortedColumns property
            // This enables multi-column sorting
            //如果開放sort功能，在更新templist時serialnum會亂掉
            this.ultraGrid1.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti;

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
        }


        private void LoadData()
        {
            LoadAmtRatioLastDay();
            LoadAmtRatioLast5Day();
            Load_UIDPrice();
            CanIssue();
            string sql_underlying_s = $@"SELECT A.[UnderlyingName]
                                               ,A.[UnderlyingID]
	                                           ,IsNull(IsNull(B.MPrice, IsNull(B.BPrice,B.APrice)),0) MPrice
                                               ,A.TraderID TraderID
                                      FROM [WarrantAssistant].[dbo].[WarrantUnderlying] A
                                      LEFT JOIN [WarrantAssistant].[dbo].[WarrantPrices] B ON A.UnderlyingID=B.CommodityID";

            DataTable Underlying_S = MSSQL.ExecSqlQry(sql_underlying_s, GlobalVar.loginSet.warrantassistant45);

            /*
            string sql = $@"SELECT A.[UID]
                          ,B.[UnderlyingName]
                          ,A.class
                          ,C.[WClass]
						  ,ROUND(C.Theta_IV金額*1000, 0) AS ThetaIV金額
	                      ,ROUND(CASE WHEN C.Acc = 0 THEN 0 ELSE (C.Theta_IV金額 - C.Theta_HV金額) * 1000 / (C.檔數) END, 0) AS 平均VSP
	                      ,CASE WHEN ISNULL(C.檔數,0) = 0 THEN 0 ELSE CAST(ISNULL(D.檔數,0) as float)/CAST( C.檔數 as float) END AS 檔數市佔
	                      ,CASE WHEN ISNULL(C.Theta_IV金額,0) = 0 THEN 0 ELSE ISNULL(D.Theta_IV金額,0) / C.Theta_IV金額 END AS ThetaIV市佔
	                      ,CASE WHEN ISNULL(E.檔數,0) = 0 THEN 0 ELSE CAST(ISNULL(F.檔數,0) as float)/CAST( E.檔數 as float) END AS 價外25以下檔數市佔
	                      ,CASE WHEN ISNULL(G.檔數,0) = 0 THEN 0 ELSE CAST(ISNULL(H.檔數,0) as float)/CAST( G.檔數 as float) END AS 價外25以上檔數市佔
                      FROM [TwData].[dbo].[URank] AS A
                      LEFT JOIN (SELECT  [UnderlyingID],[UnderlyingName],[Issuable] FROM [WarrantAssistant].[dbo].[WarrantUnderlyingSummary]) AS B ON A.UID = B.UnderlyingID
                      LEFT JOIN (SELECT  [UID],[WClass],COUNT(WID) AS 檔數,SUM(Theta_IV * AccReleasingLots) AS Theta_IV金額,SUM(Theta_HV60D * AccReleasingLots) AS Theta_HV金額, SUM(AccReleasingLots) AS Acc
			                      FROM [TwData].[dbo].[V_WarrantTrading]
			                      WHERE [TDate] = '{lastTradeDate.ToString("yyyyMMdd")}' AND [WClass] IN　('c','p') AND [TtoM] > 80 AND [WTheoPrice_IV] > 0.6 AND [IV] > [HV_60D]
			                      GROUP BY [UID],[WClass]) AS C ON A.UID = C.UID AND A.WClass = C.WClass
                      LEFT JOIN (SELECT  [UID],[WClass],COUNT(WID) AS 檔數,SUM(Theta_IV * AccReleasingLots) AS Theta_IV金額,SUM(Theta_HV60D * AccReleasingLots) AS Theta_HV金額, SUM(AccReleasingLots) AS Acc
			                      FROM [TwData].[dbo].[V_WarrantTrading]
			                      WHERE [TDate] = '{lastTradeDate.ToString("yyyyMMdd")}' AND [WClass] IN　('c','p') AND [TtoM] > 80 AND [IssuerName] = '9200' AND [WTheoPrice_IV] > 0.6 AND [IV] > [HV_60D]
			                      GROUP BY [UID],[WClass]) AS D ON A.UID = D.UID AND A.WClass = D.WClass
                      LEFT JOIN (SELECT  [UID],[WClass],COUNT(WID) AS 檔數
			                      FROM [TwData].[dbo].[V_WarrantTrading]
			                      WHERE [TDate] = '{lastTradeDate.ToString("yyyyMMdd")}' AND [WClass] IN　('c','p') AND [TtoM] > 80 AND [WTheoPrice_IV] > 0.6 AND [IV] > [HV_60D] AND CASE WHEN [WClass] = 'c' THEN 1- (StrikePrice / UClosePrice) ELSE  (StrikePrice / UClosePrice) -1 END > -0.25 AND CASE WHEN [WClass] = 'c' THEN 1- (StrikePrice / UClosePrice) ELSE  (StrikePrice / UClosePrice) -1 END < 0.05
			                      GROUP BY [UID],[WClass]) AS E ON A.UID = E.UID AND A.WClass = E.WClass
                       LEFT JOIN (SELECT  [UID],[WClass],COUNT(WID) AS 檔數
			                      FROM [TwData].[dbo].[V_WarrantTrading]
			                      WHERE [TDate] = '{lastTradeDate.ToString("yyyyMMdd")}' AND [WClass] IN　('c','p') AND [TtoM] > 80 AND [WTheoPrice_IV] > 0.6 AND [IV] > [HV_60D] AND CASE WHEN [WClass] = 'c' THEN 1- (StrikePrice / UClosePrice) ELSE  (StrikePrice / UClosePrice) -1 END > -0.25 AND CASE WHEN [WClass] = 'c' THEN 1- (StrikePrice / UClosePrice) ELSE  (StrikePrice / UClosePrice) -1 END < 0.05 AND [IssuerName] = '9200'
			                      GROUP BY [UID],[WClass]) AS F ON A.UID = F.UID AND A.WClass = F.WClass
                      LEFT JOIN (SELECT  [UID],[WClass],COUNT(WID) AS 檔數
			                      FROM [TwData].[dbo].[V_WarrantTrading]
			                      WHERE [TDate] = '{lastTradeDate.ToString("yyyyMMdd")}' AND [WClass] IN　('c','p') AND [TtoM] > 80 AND [WTheoPrice_IV] > 0.6 AND [IV] > [HV_60D] AND CASE WHEN [WClass] = 'c' THEN 1- (StrikePrice / UClosePrice) ELSE  (StrikePrice / UClosePrice) -1 END <= -0.25 
			                      GROUP BY [UID],[WClass]) AS G ON A.UID = G.UID AND A.WClass = G.WClass
                       LEFT JOIN (SELECT  [UID],[WClass],COUNT(WID) AS 檔數
			                      FROM [TwData].[dbo].[V_WarrantTrading]
			                      WHERE [TDate] = '{lastTradeDate.ToString("yyyyMMdd")}' AND [WClass] IN　('c','p') AND [TtoM] > 80 AND [WTheoPrice_IV] > 0.6 AND [IV] > [HV_60D] AND CASE WHEN [WClass] = 'c' THEN 1- (StrikePrice / UClosePrice) ELSE  (StrikePrice / UClosePrice) -1 END <= -0.25  AND [IssuerName] = '9200'
			                      GROUP BY [UID],[WClass]) AS H ON A.UID = H.UID AND A.WClass = H.WClass
                      WHERE [TDate] = '{lastTradeDate.ToString("yyyyMMdd")}' AND [IndexOrNot] = 'N' AND B.Issuable = 'Y' AND C.WClass is not null
                      ORDER BY A.class"; 
            */
            string sql = $@"SELECT A.[UID]
                          ,B.[UnderlyingName]
                          ,A.class
                          ,C.[WClass]
						  ,ROUND(C.Theta_IV金額*1000, 0) AS ThetaIV金額
				          ,ISNULL(C.檔數,0) AS All檔數
						  ,ISNULL(D.檔數,0) AS KGI檔數
	                      ,CASE WHEN ISNULL(C.Theta_IV金額,0) = 0 THEN 0 ELSE ISNULL(D.Theta_IV金額,0) / C.Theta_IV金額 END AS ThetaIV市佔
	                      ,CASE WHEN ISNULL(E.檔數,0) = 0 THEN 0 ELSE CAST(ISNULL(F.檔數,0) as float)/CAST( E.檔數 as float) END AS 價外25以下檔數市佔
	                      ,CASE WHEN ISNULL(G.檔數,0) = 0 THEN 0 ELSE CAST(ISNULL(H.檔數,0) as float)/CAST( G.檔數 as float) END AS 價外25以上檔數市佔
                      FROM [TwData].[dbo].[URank] AS A
                      LEFT JOIN (SELECT  [UnderlyingID],[UnderlyingName],[Issuable] FROM [WarrantAssistant].[dbo].[WarrantUnderlyingSummary]) AS B ON A.UID = B.UnderlyingID
                      LEFT JOIN (SELECT  [UID],[WClass],COUNT(WID) AS 檔數,SUM(Theta_IV * AccReleasingLots) AS Theta_IV金額,SUM(Theta_HV60D * AccReleasingLots) AS Theta_HV金額, SUM(AccReleasingLots) AS Acc
			                      FROM [TwData].[dbo].[V_WarrantTrading]
			                      WHERE [TDate] = '{lastTradeDate.ToString("yyyyMMdd")}' AND [WClass] IN　('c','p') AND [TtoM] > 80 AND [WTheoPrice_IV] > 0.6 AND [IV] > [HV_60D]
			                      GROUP BY [UID],[WClass]) AS C ON A.UID = C.UID AND A.WClass = C.WClass
                      LEFT JOIN (SELECT  [UID],[WClass],COUNT(WID) AS 檔數,SUM(Theta_IV * AccReleasingLots) AS Theta_IV金額,SUM(Theta_HV60D * AccReleasingLots) AS Theta_HV金額, SUM(AccReleasingLots) AS Acc
			                      FROM [TwData].[dbo].[V_WarrantTrading]
			                      WHERE [TDate] = '{lastTradeDate.ToString("yyyyMMdd")}' AND [WClass] IN　('c','p') AND [TtoM] > 80 AND [IssuerName] = '9200' AND [WTheoPrice_IV] > 0.6 AND [IV] > [HV_60D]
			                      GROUP BY [UID],[WClass]) AS D ON A.UID = D.UID AND A.WClass = D.WClass
                      LEFT JOIN (SELECT  [UID],[WClass],COUNT(WID) AS 檔數
			                      FROM [TwData].[dbo].[V_WarrantTrading]
			                      WHERE [TDate] = '{lastTradeDate.ToString("yyyyMMdd")}' AND [WClass] IN　('c','p') AND [TtoM] > 80 AND [WTheoPrice_IV] > 0.6 AND [IV] > [HV_60D] AND CASE WHEN [WClass] = 'c' THEN 1- (StrikePrice / UClosePrice) ELSE  (StrikePrice / UClosePrice) -1 END > -0.25 AND CASE WHEN [WClass] = 'c' THEN 1- (StrikePrice / UClosePrice) ELSE  (StrikePrice / UClosePrice) -1 END < 0.05
			                      GROUP BY [UID],[WClass]) AS E ON A.UID = E.UID AND A.WClass = E.WClass
                       LEFT JOIN (SELECT  [UID],[WClass],COUNT(WID) AS 檔數
			                      FROM [TwData].[dbo].[V_WarrantTrading]
			                      WHERE [TDate] = '{lastTradeDate.ToString("yyyyMMdd")}' AND [WClass] IN　('c','p') AND [TtoM] > 80 AND [WTheoPrice_IV] > 0.6 AND [IV] > [HV_60D] AND CASE WHEN [WClass] = 'c' THEN 1- (StrikePrice / UClosePrice) ELSE  (StrikePrice / UClosePrice) -1 END > -0.25 AND CASE WHEN [WClass] = 'c' THEN 1- (StrikePrice / UClosePrice) ELSE  (StrikePrice / UClosePrice) -1 END < 0.05 AND [IssuerName] = '9200'
			                      GROUP BY [UID],[WClass]) AS F ON A.UID = F.UID AND A.WClass = F.WClass
                      LEFT JOIN (SELECT  [UID],[WClass],COUNT(WID) AS 檔數
			                      FROM [TwData].[dbo].[V_WarrantTrading]
			                      WHERE [TDate] = '{lastTradeDate.ToString("yyyyMMdd")}' AND [WClass] IN　('c','p') AND [TtoM] > 80 AND [WTheoPrice_IV] > 0.6 AND [IV] > [HV_60D] AND CASE WHEN [WClass] = 'c' THEN 1- (StrikePrice / UClosePrice) ELSE  (StrikePrice / UClosePrice) -1 END <= -0.25 
			                      GROUP BY [UID],[WClass]) AS G ON A.UID = G.UID AND A.WClass = G.WClass
                       LEFT JOIN (SELECT  [UID],[WClass],COUNT(WID) AS 檔數
			                      FROM [TwData].[dbo].[V_WarrantTrading]
			                      WHERE [TDate] = '{lastTradeDate.ToString("yyyyMMdd")}' AND [WClass] IN　('c','p') AND [TtoM] > 80 AND [WTheoPrice_IV] > 0.6 AND [IV] > [HV_60D] AND CASE WHEN [WClass] = 'c' THEN 1- (StrikePrice / UClosePrice) ELSE  (StrikePrice / UClosePrice) -1 END <= -0.25  AND [IssuerName] = '9200'
			                      GROUP BY [UID],[WClass]) AS H ON A.UID = H.UID AND A.WClass = H.WClass
                      WHERE [TDate] = '{lastTradeDate.ToString("yyyyMMdd")}' AND [IndexOrNot] = 'N' AND B.Issuable = 'Y' AND C.WClass is not null
                      ORDER BY A.class";
            SqlConnection conn = new SqlConnection(GlobalVar.loginSet.twData);
            SqlCommand cmd = new SqlCommand(sql, conn);
            cmd.CommandTimeout = 0;
            DataTable dtLoad = new DataTable();
            using (SqlDataAdapter adp = new SqlDataAdapter(cmd))
            {
                adp.Fill(dtLoad);
            }
            //要把未掛牌權證也納進來算
            string sql_unlisted = $@"SELECT  
                                      [標的代號] AS UID
	                                  ,CASE WHEN [名稱] like '%購%' THEN　'c' ELSE 'p' END AS WClass
                                      ,COUNT([代號]) AS Cnt
                                  FROM [TwCMData].[dbo].[Warrant總表]
                                  WHERE [日期] = '{lastTradeDate.ToString("yyyyMMdd")}' AND ([上市日期] = '19110101' OR [上市日期] > '{lastTradeDate.ToString("yyyyMMdd")}') AND ([名稱] like '%購%' OR [名稱] like '%售%')
                                  AND LEN([標的代號]) < 5 AND LEFT([標的代號],2) <> '00'
                                  GROUP BY [標的代號],CASE WHEN [名稱] like '%購%' THEN　'c' ELSE 'p' END";
            DataTable dt_unlisted = MSSQL.ExecSqlQry(sql_unlisted, GlobalVar.loginSet.twCMData);

            string sql_unlistedkgi = $@"SELECT  
                                      [標的代號] AS UID
	                                  ,CASE WHEN [名稱] like '%購%' THEN　'c' ELSE 'p' END AS WClass
                                      ,COUNT([代號]) AS Cnt
                                  FROM [TwCMData].[dbo].[Warrant總表]
                                  WHERE [日期] = '{lastTradeDate.ToString("yyyyMMdd")}' AND ([上市日期] = '19110101' OR [上市日期] > '{lastTradeDate.ToString("yyyyMMdd")}') AND ([名稱] like '%購%' OR [名稱] like '%售%')
                                  AND LEN([標的代號]) < 5 AND LEFT([標的代號],2) <> '00' AND [券商代號] = '9200'
                                  GROUP BY [標的代號],CASE WHEN [名稱] like '%購%' THEN　'c' ELSE 'p' END";
            DataTable dt_unlistedkgi = MSSQL.ExecSqlQry(sql_unlistedkgi, GlobalVar.loginSet.twCMData);
            foreach (DataRow drLoad in dtLoad.Rows)
            {
                DataRow dr = dt.NewRow();
                string uid = drLoad["UID"].ToString();
                double underlyingPrice = 0;
                
                DataRow[] Underlying_S_Select = Underlying_S.Select($@"UnderlyingID = '{uid}'");
                if (Underlying_S_Select.Length > 0)
                {
                    underlyingPrice = Convert.ToDouble(Underlying_S_Select[0][2].ToString());
                }
                
                string uname = drLoad["UnderlyingName"].ToString();
                string catagory = drLoad["class"].ToString();
                string wclass = drLoad["WClass"].ToString();

                DataRow[] dt_unlistedSelect = dt_unlisted.Select($@"UID = '{uid}' AND WClass = '{wclass}'");
                double cnt_unlisted = 0;
                if(dt_unlistedSelect.Length > 0)
                {
                    cnt_unlisted = Convert.ToDouble(dt_unlistedSelect[0][2].ToString());
                }

                DataRow[] dt_unlistedkgiSelect = dt_unlistedkgi.Select($@"UID = '{uid}' AND WClass = '{wclass}'");
                double cnt_unlistedkgi= 0;
                if (dt_unlistedkgiSelect.Length > 0)
                {
                    cnt_unlistedkgi = Convert.ToDouble(dt_unlistedkgiSelect[0][2].ToString());
                }

                double theta = Convert.ToDouble(drLoad["ThetaIV金額"].ToString());
                double allcnt = Convert.ToDouble(drLoad["All檔數"].ToString());
                double kgicnt = Convert.ToDouble(drLoad["KGI檔數"].ToString());
                double avgASP = 0;
                double cntShare = 0;
                if (allcnt + cnt_unlisted > 0)
                {
                    avgASP = Math.Round(theta / (allcnt + cnt_unlisted), 2);
                    cntShare = Math.Round((kgicnt + cnt_unlistedkgi) / (allcnt + cnt_unlisted) * 100, 1);
                }
                double thetaShare = Math.Round(Convert.ToDouble(drLoad["ThetaIV市佔"].ToString()) * 100, 1);
                double moneyness25DownShare = Math.Round(Convert.ToDouble(drLoad["價外25以下檔數市佔"].ToString()) * 100, 1);
                double moneyness25UpShare = Math.Round(Convert.ToDouble(drLoad["價外25以上檔數市佔"].ToString()) * 100, 1);


                dr["標的代號-名稱"] = uid + "-" + uname;
                
                dr["WClass"] = wclass;
                
                dr["分級"] = catagory;
                
                dr["Theta金額"] = theta;
                
                dr["平均VSP"] = avgASP;

                dr["檔數市佔"] = cntShare;
                dr["Theta市佔率"] = thetaShare;


                try
                {
                    double updown = UID_PriceUpDown[uid];
                    double refP = UID_RefPrice[uid];
                    double todayUpDown = 0;
                    if (refP > 0)
                        todayUpDown = Math.Round(((underlyingPrice - refP) / refP) * 100, 2);
                    if(updown > 10 || updown + todayUpDown > 10)
                        dr["股價累計漲幅"] = "V";
                }
                catch(Exception ex) 
                {
                    MessageBox.Show($@"{ex.Message} 股價累積漲幅計算有誤");
                }
                try
                {
                    if(wclass == "c")
                    {
                        if (Call_GoodAmtRatio_LastDay.ContainsKey(uid))
                        {
                            if (Call_GoodAmtRatio_LastDay[uid] >= 0.7)
                                dr["好分點買進"] = $@"V";
                        }
                        if (Call_GoodAmtRatio_Last5Day.ContainsKey(uid))
                        {
                            if (Call_GoodAmtRatio_Last5Day[uid] >= 0.7)
                                dr["好分點買進"] = $@"V";
                        }
                    }
                    if (wclass == "p")
                    {
                        if (Put_GoodAmtRatio_LastDay.ContainsKey(uid))
                        {
                            if (Put_GoodAmtRatio_LastDay[uid] >= 0.7)
                                dr["好分點買進"] = $@"V";
                        }
                        if (Put_GoodAmtRatio_Last5Day.ContainsKey(uid))
                        {
                            if (Put_GoodAmtRatio_Last5Day[uid] >= 0.7)
                                dr["好分點買進"] = $@"V";
                        }
                    }
                    
                }
                catch (Exception ex2)
                {
                    MessageBox.Show($@"{ex2.Message} 計算好分點買進有誤");
                }

                try
                {
                    if(OptionCanIssue.ContainsKey(uid))
                        if(OptionCanIssue[uid] < 10)
                            dr["搶發標的"] = "V";
                }
                catch (Exception ex3)
                {
                    MessageBox.Show($@"{ex3.Message} 計算搶發標的有誤");
                }



                dr["價外25以下"] = moneyness25DownShare;
                dr["價外25以上"] = moneyness25UpShare;
                
                dt.Rows.Add(dr);
            }

        }

        private void Load_UIDPrice()
        {
            UID_PriceUpDown.Clear();
            UID_RefPrice.Clear();
            try
            {
                string sql1 = $@"SELECT 
                              [股票代號]
                              ,SUM([漲幅(%)]) AS 漲幅
                          FROM [TwCMData].[dbo].[日收盤表排行]
                          WHERE [日期] >= '{last3TradeDate.ToString("yyyyMMdd")}' AND [日期] <= '{lastTradeDate.ToString("yyyyMMdd")}' AND LEN([股票代號]) < 5 AND LEFT('股票代號',2) <> '00'
                          GROUP BY [股票代號]";
                DataTable dt1 = MSSQL.ExecSqlQry(sql1, GlobalVar.loginSet.twCMData);
                foreach (DataRow dr in dt1.Rows)
                {
                    string uid = dr["股票代號"].ToString();
                    double updown = Convert.ToDouble(dr["漲幅"].ToString());
                    if (!UID_PriceUpDown.ContainsKey(uid))
                        UID_PriceUpDown.Add(uid, updown);
                }

                string sql2 = $@"  SELECT 
                              [股票代號]
                              ,ISNULL([自算參考價],0) AS [自算參考價]
                          FROM [TwCMData].[dbo].[日收盤表排行]
                          WHERE [日期] = '{lastTradeDate.ToString("yyyyMMdd")}' AND LEN([股票代號]) < 5 AND LEFT('股票代號',2) <> '00'";
                DataTable dt2 = MSSQL.ExecSqlQry(sql2, GlobalVar.loginSet.twCMData);
                foreach (DataRow dr in dt2.Rows)
                {
                    string uid = dr["股票代號"].ToString();
                    double refP = Convert.ToDouble(dr["自算參考價"].ToString());
                    if (!UID_RefPrice.ContainsKey(uid))
                        UID_RefPrice.Add(uid, refP);
                }
            }
            catch(Exception e)
            {
                MessageBox.Show(e.Message);
            }
        }

        private void LoadAmtRatioLastDay()
        {
            Call_GoodAmtRatio_LastDay.Clear();
            Put_GoodAmtRatio_LastDay.Clear();
            string sql = $@" SELECT P.UID,P.WClass,Round(ISNULL(Q.Amt,0) / P.Amt,4) AS GoodAmtRatio FROM (

                            SELECT [UID],[WClass],SUM(Amt) AS Amt FROM (
                              SELECT  [UID],[WClass],[LastPx] * [LastQty] AS Amt,[BrokerBuyID],[BrokerSellID]
                              FROM [TwData].[dbo].[WarrantTradingBrokerNew] 
                              WHERE [TradeDate] = '{lastTradeDate.ToString("yyyyMMdd")}' AND [BrokerBuyFlag] <> 'MM' AND [BrokerSellFlag] = 'MM') AS A
                              LEFT JOIN (SELECT  [BrokerID],[Class] FROM [10.60.0.37].[newEDIS].[dbo].[BrokerClassification]
                                         WHERE [TDate] = (SELECT MAX(TDate) FROM [10.60.0.37].[newEDIS].[dbo].[BrokerClassification])) AS B ON A.BrokerBuyID = B.[BrokerID] Collate SQL_Latin1_General_CP1_CS_AS
                              GROUP BY [UID],[WClass]) AS P

                              LEFT JOIN 

                            (SELECT [UID],[WClass],SUM(Amt) AS Amt FROM (
                              SELECT  [UID],[WClass],[LastPx] * [LastQty] AS Amt,[BrokerBuyID],[BrokerSellID]
                              FROM [TwData].[dbo].[WarrantTradingBrokerNew] 
                              WHERE [TradeDate] = '{lastTradeDate.ToString("yyyyMMdd")}' AND [BrokerBuyFlag] <> 'MM' AND [BrokerSellFlag] = 'MM') AS A
                              LEFT JOIN (SELECT  [BrokerID],[Class] FROM [10.60.0.37].[newEDIS].[dbo].[BrokerClassification]
                                         WHERE [TDate] = (SELECT MAX(TDate) FROM [10.60.0.37].[newEDIS].[dbo].[BrokerClassification])) AS B ON A.BrokerBuyID = B.[BrokerID] Collate SQL_Latin1_General_CP1_CS_AS
                              WHERE B.[Class] = 'Good'
                              GROUP BY [UID],[WClass]) AS Q ON P.UID = Q.UID AND P.WClass = Q.WClass";
            
            SqlConnection conn = new SqlConnection(GlobalVar.loginSet.twData);
            SqlCommand cmd = new SqlCommand(sql, conn);
            cmd.CommandTimeout = 0;
            DataTable dt = new DataTable();
            using (SqlDataAdapter adp = new SqlDataAdapter(cmd))
            {
                adp.Fill(dt);
            }
            foreach (DataRow dr in dt.Rows)
            {
                string uid = dr["UID"].ToString().Trim();
                string cp = dr["WClass"].ToString();
                double r = Convert.ToDouble(dr["GoodAmtRatio"].ToString());
                if (cp == "c")
                {
                    if (!Call_GoodAmtRatio_LastDay.ContainsKey(uid))
                        Call_GoodAmtRatio_LastDay.Add(uid, r);
                }
                else
                {
                    if (!Put_GoodAmtRatio_LastDay.ContainsKey(uid))
                        Put_GoodAmtRatio_LastDay.Add(uid, r);
                }
            }
        }

        private void LoadAmtRatioLast5Day()
        {
            Call_GoodAmtRatio_Last5Day.Clear();
            Put_GoodAmtRatio_Last5Day.Clear();
            string sql = $@" SELECT P.UID,P.WClass,Round(ISNULL(Q.Amt,0) / P.Amt,4) AS GoodAmtRatio FROM (

                            SELECT [UID],[WClass],SUM(Amt) AS Amt FROM (
                              SELECT  [UID],[WClass],[LastPx] * [LastQty] AS Amt,[BrokerBuyID],[BrokerSellID]
                              FROM [TwData].[dbo].[WarrantTradingBrokerNew] 
                              WHERE [TradeDate] <= '{lastTradeDate.ToString("yyyyMMdd")}' AND [TradeDate] >= '{last5TradeDate.ToString("yyyyMMdd")}' AND [BrokerBuyFlag] <> 'MM' AND [BrokerSellFlag] = 'MM') AS A
                              LEFT JOIN (SELECT  [BrokerID],[Class] FROM [10.60.0.37].[newEDIS].[dbo].[BrokerClassification]
                                         WHERE [TDate] = (SELECT MAX(TDate) FROM [10.60.0.37].[newEDIS].[dbo].[BrokerClassification])) AS B ON A.BrokerBuyID = B.[BrokerID] Collate SQL_Latin1_General_CP1_CS_AS
                              GROUP BY [UID],[WClass]) AS P

                              LEFT JOIN 

                            (SELECT [UID],[WClass],SUM(Amt) AS Amt FROM (
                              SELECT  [UID],[WClass],[LastPx] * [LastQty] AS Amt,[BrokerBuyID],[BrokerSellID]
                              FROM [TwData].[dbo].[WarrantTradingBrokerNew] 
                              WHERE [TradeDate] <= '{lastTradeDate.ToString("yyyyMMdd")}' AND [TradeDate] >= '{last5TradeDate.ToString("yyyyMMdd")}' AND [BrokerBuyFlag] <> 'MM' AND [BrokerSellFlag] = 'MM') AS A
                              LEFT JOIN (SELECT  [BrokerID],[Class] FROM [10.60.0.37].[newEDIS].[dbo].[BrokerClassification]
                                         WHERE [TDate] = (SELECT MAX(TDate) FROM [10.60.0.37].[newEDIS].[dbo].[BrokerClassification])) AS B ON A.BrokerBuyID = B.[BrokerID] Collate SQL_Latin1_General_CP1_CS_AS
                              WHERE B.[Class] = 'Good'
                              GROUP BY [UID],[WClass]) AS Q ON P.UID = Q.UID AND P.WClass = Q.WClass";

            SqlConnection conn = new SqlConnection(GlobalVar.loginSet.twData);
            SqlCommand cmd = new SqlCommand(sql, conn);
            cmd.CommandTimeout = 0;
            DataTable dt = new DataTable();
            using (SqlDataAdapter adp = new SqlDataAdapter(cmd))
            {
                adp.Fill(dt);
            }
            foreach (DataRow dr in dt.Rows)
            {
                string uid = dr["UID"].ToString();
                string cp = dr["WClass"].ToString();
                double r = Convert.ToDouble(dr["GoodAmtRatio"].ToString());
                if (cp == "c")
                {
                    if (!Call_GoodAmtRatio_Last5Day.ContainsKey(uid))
                        Call_GoodAmtRatio_Last5Day.Add(uid, r);
                }
                else
                {
                    if (!Put_GoodAmtRatio_Last5Day.ContainsKey(uid))
                        Put_GoodAmtRatio_Last5Day.Add(uid, r);
                }
            }
        }


        //用0.6元，價平權證計算有多少檔數可發
        private void CanIssue()
        {
            OptionCanIssue.Clear();
            string sql = $@"SELECT  [UID]
                                      ,[OptionAvailable]
                                        FROM [WarrantAssistant].[dbo].[OptionAutoSelectData]";
            DataTable dt = MSSQL.ExecSqlQry(sql, GlobalVar.loginSet.warrantassistant45);
            foreach(DataRow dr in dt.Rows)
            {
                string uid = dr["UID"].ToString();
                double canIssue = Convert.ToDouble(dr["OptionAvailable"].ToString());
                if (!OptionCanIssue.ContainsKey(uid))
                    OptionCanIssue.Add(uid, canIssue);
            }
        }
        private void UltraGrid1_DoubleClickCell(object sender, DoubleClickCellEventArgs e)
        {

            string uid = e.Cell.Row.Cells[0].Value.ToString();
            string cp = e.Cell.Row.Cells[1].Value.ToString();
            bool formExist = false;
            FrmGeneralIssueV2Sub f1 = new FrmGeneralIssueV2Sub();
            f1.textBox1.Text = uid + " : " + cp;
            f1.Show();
            /*
            FrmGeneralIssueV2Sub f1;
            foreach (Form iForm in System.Windows.Forms.Application.OpenForms)
            {//只出現一個頁面，不允許開多個視窗
                if (iForm.GetType() == typeof(FrmGeneralIssueV2Sub))
                {
                    f1 = (FrmGeneralIssueV2Sub)iForm;
                    formExist = true;
                    f1.textBox1.Text = uid + " : " + cp;
                    //f1.InitialGrid();
                    f1.LoadData();
                    f1.BringToFront();
                    f1.Show();
                }
            }
            if(formExist == false)
            {
                f1 = new FrmGeneralIssueV2Sub();
                f1.textBox1.Text = uid + " : " + cp;
                f1.Show();
            }
            */
            /*
            string uid = e.Cell.Row.Cells[0].Value.ToString();
            string cp = e.Cell.Row.Cells[16].Value.ToString();
            string key = uid + "_" + cp;
            if (k_seperate.ContainsKey(key))
            {
                DataTable dt = new DataTable();
                foreach (string k in k_seperate[key].Keys)
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
            */
        }
    }
}

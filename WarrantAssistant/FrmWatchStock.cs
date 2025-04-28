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

namespace WarrantAssistant
{
    public partial class FrmWatchStock : Form
    {

        DataTable dt = new DataTable();
        public static CMADODB5.CMConnection cn = new CMADODB5.CMConnection();
        public static string arg = "5"; //%
        public static string srvLocation = "10.60.0.191";
        public static string cnPort = "";
        public FrmWatchStock()
        {
            try
            {
                InitializeComponent();
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message.ToString());
            }
        }

        private void FrmWatchStock_Load(object sender, EventArgs e)
        {
            InitialGrid();
            
            LoadData2();
        }
        private void InitialGrid()
        {
            dt.Columns.Add("標的代號", typeof(string));
            dt.Columns.Add("標的名稱", typeof(string));
            
            dt.Columns.Add("累積漲跌異常", typeof(string));
            dt.Columns.Add("起訖日漲跌異常", typeof(string));
            dt.Columns.Add("成交量過高", typeof(string));
            dt.Columns.Add("周轉率過高", typeof(string));
            dt.Columns.Add("券資比過高", typeof(string));
            dt.Columns.Add("平均成交量過高", typeof(string));
            dt.Columns.Add("累積周轉率過高", typeof(string));
            dt.Columns.Add("股價價差異常", typeof(string));
            dt.Columns.Add("借券賣出過高", typeof(string));
            dt.Columns.Add("當沖比過高", typeof(string));
            //dt.Columns.Add("注意11", typeof(string));
            //dt.Columns.Add("注意12", typeof(string));
            //dt.Columns.Add("注意13", typeof(string));
            


            ultraGrid1.DataSource = dt;
            UltraGridBand band0 = ultraGrid1.DisplayLayout.Bands[0];

            
            band0.Columns["標的代號"].Format = "N0";
            band0.Columns["標的名稱"].Format = "N0";
            
            
            band0.Columns["標的代號"].Width = 60;
            band0.Columns["標的名稱"].Width = 150;
            band0.Columns["累積漲跌異常"].Width = 100;
            band0.Columns["起訖日漲跌異常"].Width = 110;
            band0.Columns["成交量過高"].Width = 90;
            band0.Columns["周轉率過高"].Width = 90;
            band0.Columns["券資比過高"].Width = 90;
            band0.Columns["平均成交量過高"].Width = 110;
            band0.Columns["累積周轉率過高"].Width = 110;
            band0.Columns["股價價差異常"].Width = 100;
            band0.Columns["借券賣出過高"].Width = 100;
            band0.Columns["當沖比過高"].Width = 90;
            //band0.Columns["注意11"].Width = 70;
            //band0.Columns["注意12"].Width = 70;
            //band0.Columns["注意13"].Width = 70;
            


            band0.Override.HeaderAppearance.TextHAlign = Infragistics.Win.HAlign.Left;

           
            this.ultraGrid1.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti;

        }
        private void ultraGrid1_InitializeLayout(object sender, InitializeLayoutEventArgs e)
        {
            ultraGrid1.DisplayLayout.Override.RowSelectorHeaderStyle = RowSelectorHeaderStyle.ColumnChooserButton;
        }

        private void ultraGrid1_InitializeRow(object sender, InitializeRowEventArgs e)
        {

            string uid = e.Row.Cells["標的代號"].Value.ToString();
            string W1 = e.Row.Cells["累積漲跌異常"].Value.ToString();
            if(W1 == "1")
                e.Row.Cells["累積漲跌異常"].Appearance.BackColor = Color.LightPink;

            string W2 = e.Row.Cells["起訖日漲跌異常"].Value.ToString();
            if (W2 == "1")
                e.Row.Cells["起訖日漲跌異常"].Appearance.BackColor = Color.LightPink;

            string W3 = e.Row.Cells["成交量過高"].Value.ToString();
            if (W3 == "1")
                e.Row.Cells["成交量過高"].Appearance.BackColor = Color.LightPink;

            string W4 = e.Row.Cells["周轉率過高"].Value.ToString();
            if (W4 == "1")
                e.Row.Cells["周轉率過高"].Appearance.BackColor = Color.LightPink;
            string W5 = e.Row.Cells["券資比過高"].Value.ToString();
            if (W5 == "1")
                e.Row.Cells["券資比過高"].Appearance.BackColor = Color.LightPink;
            string W6 = e.Row.Cells["平均成交量過高"].Value.ToString();
            if (W6 == "1")
                e.Row.Cells["平均成交量過高"].Appearance.BackColor = Color.LightPink;
            string W7 = e.Row.Cells["累積周轉率過高"].Value.ToString();
            if (W7 == "1")
                e.Row.Cells["累積周轉率過高"].Appearance.BackColor = Color.LightPink;
            string W8 = e.Row.Cells["股價價差異常"].Value.ToString();
            if (W8 == "1")
                e.Row.Cells["股價價差異常"].Appearance.BackColor = Color.LightPink;
            string W9 = e.Row.Cells["借券賣出過高"].Value.ToString();
            if (W9 == "1")
                e.Row.Cells["借券賣出過高"].Appearance.BackColor = Color.LightPink;
            string W10 = e.Row.Cells["當沖比過高"].Value.ToString();
            if (W10 == "1")
                e.Row.Cells["當沖比過高"].Appearance.BackColor = Color.LightPink;
        }
        private void LoadData()
        {

            Dictionary<string, string> TSEOTC = new Dictionary<string, string>();

            string sqlTSEOTC = $@"SELECT [股票代號],[上市櫃] FROM [權證標的證券(季)] WHERE [年季] = (SELECT MAX(年季) FROM [權證標的證券(季)]) AND [上市櫃] IS NOT NULL";
            
            ADODB.Recordset rs = cn.CMExecute(ref arg, srvLocation, cnPort, sqlTSEOTC);
            for (; !rs.EOF; rs.MoveNext())
            {
                string uid = rs.Fields["股票代號"].Value;
                string ex = rs.Fields["上市櫃"].Value;
                if (ex.ToString() == "上市")
                    ex = "1";
                else
                    ex = "2";
                if (!TSEOTC.ContainsKey(uid))
                    TSEOTC.Add(uid, ex);
            }

            //上市上櫃累積漲跌幅
            string sqlPriceUpDown = $@"(SELECT C.[上市上櫃],ROUND(AVG(A.漲跌幅+B.盤中漲跌幅),2) AS 平均漲跌幅 FROM
                                      (SELECT  [股票代號],SUM([漲幅(%)]) AS 漲跌幅 FROM [TwCMData].[dbo].[日收盤表排行]
                                      WHERE [日期] >= '{EDLib.TradeDate.LastNTradeDate(5).ToString("yyyyMMdd")}' AND LEN([股票代號]) < 5 AND LEFT([股票代號],2) <> '00' GROUP BY [股票代號]) AS A 
                                      LEFT JOIN (SELECT  [Symbol],CASE WHEN [TrdPrz] = 0 THEN [RefPrz] ELSE [TrdPrz] END AS [TrdPrz],[RefPrz],((CASE WHEN [TrdPrz] = 0 THEN [RefPrz] ELSE [TrdPrz] END) - [RefPrz]) / [RefPrz] * 100 AS 盤中漲跌幅
                                      FROM [TwQuoteAddin].[dbo].[STKQUOTE] WHERE [UpdateTime] > CONVERT(VARCHAR,GETDATE(),112) AND [RefPrz] > 0) AS B ON A.股票代號 = B.Symbol
                                      LEFT JOIN( SELECT [股票代號],[產業代號],[上市上櫃] FROM [TwCMData].[dbo].[上市櫃公司基本資料] WHERE [年度] = (SELECT MAX(年度) FROM [TwCMData].[dbo].[上市櫃公司基本資料])) AS C ON A.股票代號 = C.股票代號
                                      GROUP BY C.[上市上櫃]) ";
            DataTable dtPriceUpDown = MSSQL.ExecSqlQry(sqlPriceUpDown, GlobalVar.loginSet.twCMData);
            double TSEPriceUpDown6 = 0;
            double OTCPriceUpDown6 = 0;
            DataRow[] drPriceUpDown_TSE_Select = dtPriceUpDown.Select($@"上市上櫃 = '1'");
            if (drPriceUpDown_TSE_Select.Length > 0)
            {
                TSEPriceUpDown6 = Convert.ToDouble(drPriceUpDown_TSE_Select[0][1].ToString());
            }
            DataRow[] drPriceUpDown_OTC_Select = dtPriceUpDown.Select($@"上市上櫃 = '2'");
            if (drPriceUpDown_OTC_Select.Length > 0)
            {
                OTCPriceUpDown6 = Convert.ToDouble(drPriceUpDown_OTC_Select[0][1].ToString());
            }

            string sqlUID = $@"SELECT  A.[UnderlyingID],A.[UnderlyingName], A.[MDate], B.[TraderAccount] ,C.[Market] 
                                                FROM [WarrantAssistant].[dbo].[WarrantIssueCheck] as A
                                                LEFT JOIN [TwData].[dbo].[Underlying_Trader] as B on A.[UnderlyingID] = B.[UID]
                                                LEFT JOIN [WarrantAssistant].[dbo].[WarrantUnderlyingSummary] as C ON A.UnderlyingID = C.UnderlyingID
                                                WHERE A.[WatchCount] = 1 AND C.Issuable = 'Y' AND A.[IsQuaterUnderlying] = 'Y'";
            DataTable dtUID = MSSQL.ExecSqlQry(sqlUID, GlobalVar.loginSet.warrantassistant45);
            string lastTradeDate89 = EDLib.TradeDate.LastNTradeDate(89).ToString("yyyyMMdd");
            string lastTradeDate60 = EDLib.TradeDate.LastNTradeDate(60).ToString("yyyyMMdd");
            string lastTradeDate59 = EDLib.TradeDate.LastNTradeDate(59).ToString("yyyyMMdd");
            string lastTradeDate29 = EDLib.TradeDate.LastNTradeDate(29).ToString("yyyyMMdd");
            string lastTradeDate6 = EDLib.TradeDate.LastNTradeDate(6).ToString("yyyyMMdd");
            string lastTradeDate5 = EDLib.TradeDate.LastNTradeDate(5).ToString("yyyyMMdd");
            string lastTradeDate1 = EDLib.TradeDate.LastNTradeDate(1).ToString("yyyyMMdd");
            
            foreach (DataRow drW in dtUID.Rows)
            {
                string W1 = "0";
                string W2 = "0";
                string W3 = "0";
                string W4 = "0";
                string W5 = "0";
                string W6 = "0";
                string W7 = "0";
                string W8 = "0";
                string W9 = "0";
                string W10 = "0";
                DataRow dr = dt.NewRow();
                
                string uid = drW["UnderlyingID"].ToString();
                string uname = drW["UnderlyingName"].ToString();
                //MessageBox.Show(uid);
                /*
                if (uid == "5443")
                    MessageBox.Show("5443");
                else
                    continue;
                */
                dr["標的代號"] = uid;
                dr["標的名稱"] = uname;
                //注意1
                //一、	有價證券最近一段期間累積之收盤價漲跌百分比異常者(二選一)
                //   累積收盤價漲幅 = 六天的漲跌幅相加
                //    1.最近六個營業日（含當日）累積之收盤價漲跌百分比超過32 %，且其漲跌百分比與全體及同類依本項規定計算之平均值的差幅均在20 % 以上者。
                //    2.最近六個營業日（含當日）累積之收盤價漲跌百分比超過25 %，且其漲跌百分比與全體及同類依本項規定計算之平均值的差幅均在20 % 以上，及最近六個營業日（含當日）起迄兩個營業日之收盤價價差達50元以上者。
                /*
                string sqlW1_1 = $@" SELECT A.股票代號,
                (ROUND((A.漲跌幅+B.盤中漲跌幅),2)) AS 累積盤中漲跌幅, 
                ROUND((A.價差+B.盤中價差),2) AS 累積盤中價差,D.產業平均漲跌幅,
                ROUND(ABS((A.漲跌幅+B.盤中漲跌幅) - D.產業平均漲跌幅),2) AS 與平均差 FROM
                                      (SELECT  [股票代號],SUM([漲跌]) AS 價差,SUM([漲幅(%)]) AS 漲跌幅
                                      FROM [TwCMData].[dbo].[日收盤表排行]
                                      WHERE [日期] >= '{lastTradeDate5}'  AND LEN([股票代號]) < 5 AND LEFT([股票代號],2) <> '00'
                                      GROUP BY [股票代號]) AS A 
                                      LEFT JOIN 
                                      (SELECT  [Symbol],[TrdPrz],[RefPrz],([TrdPrz] - [RefPrz]) / [RefPrz] * 100 AS 盤中漲跌幅,[TrdPrz] - [RefPrz] AS 盤中價差
                                      FROM [TwQuoteAddin].[dbo].[STKQUOTE] WHERE [UpdateTime] > CONVERT(VARCHAR,GETDATE(),112) AND [RefPrz] > 0) AS B ON A.股票代號 = B.Symbol
                                      LEFT JOIN
                                      (SELECT  [股票代號],[產業代號],[上市上櫃]
                                      FROM [TwCMData].[dbo].[上市櫃公司基本資料]
                                      WHERE [年度] = (SELECT MAX(年度) FROM [TwCMData].[dbo].[上市櫃公司基本資料])) AS C ON A.股票代號 = C.股票代號
                                      LEFT JOIN 
                                      (SELECT C.[產業代號],C.[上市上櫃],ROUND(AVG(A.漲跌幅+B.盤中漲跌幅),2) AS 產業平均漲跌幅 FROM
                                      (SELECT  [股票代號],SUM([漲幅(%)]) AS 漲跌幅 FROM [TwCMData].[dbo].[日收盤表排行]
                                      WHERE [日期] >= '{lastTradeDate5}'  AND LEN([股票代號]) < 5 AND LEFT([股票代號],2) <> '00' GROUP BY [股票代號]) AS A 
                                      LEFT JOIN (SELECT  [Symbol],[TrdPrz],[RefPrz],([TrdPrz] - [RefPrz]) / [RefPrz] * 100 AS 盤中漲跌幅
                                      FROM [TwQuoteAddin].[dbo].[STKQUOTE] WHERE [UpdateTime] > CONVERT(VARCHAR,GETDATE(),112) AND [RefPrz] > 0) AS B ON A.股票代號 = B.Symbol
                                      LEFT JOIN( SELECT [股票代號],[產業代號],[上市上櫃] FROM [TwCMData].[dbo].[上市櫃公司基本資料] WHERE [年度] = (SELECT MAX(年度) FROM [TwCMData].[dbo].[上市櫃公司基本資料])) AS C ON A.股票代號 = C.股票代號
                                      GROUP BY C.[產業代號],C.[上市上櫃]) AS D ON C.產業代號 = D.產業代號 AND C.上市上櫃 = D.上市上櫃
                                      WHERE B.Symbol IS NOT NULL";
                */

                string sqlW1_1 = $@"SELECT A.股票代號
                                        ,(ROUND((A.漲跌幅+B.盤中漲跌幅),2)) AS 累積盤中漲跌幅
                                        , ROUND((A.價差+B.盤中價差),2) AS 累積盤中價差
                                        ,D.產業平均漲跌幅,ROUND((A.漲跌幅+B.盤中漲跌幅) - D.產業平均漲跌幅,2) AS 與平均差
                                        FROM
                                        (SELECT  [股票代號],SUM([漲跌]) AS 價差,SUM([漲幅(%)]) AS 漲跌幅
                                      FROM [TwCMData].[dbo].[日收盤表排行]
                                      WHERE [日期] >= '{lastTradeDate5}' AND LEN([股票代號]) < 5 AND LEFT([股票代號],2) <> '00'
                                      GROUP BY [股票代號]) AS A 
                                      LEFT JOIN 
                                      (SELECT  [Symbol],CASE WHEN [TrdPrz] = 0 THEN [RefPrz] ELSE [TrdPrz] END AS [TrdPrz],[RefPrz],((CASE WHEN [TrdPrz] = 0 THEN [RefPrz] ELSE [TrdPrz] END) - [RefPrz]) / [RefPrz] * 100 AS 盤中漲跌幅,(CASE WHEN [TrdPrz] = 0 THEN [RefPrz] ELSE [TrdPrz] END) - [RefPrz] AS 盤中價差
                                      FROM [TwQuoteAddin].[dbo].[STKQUOTE] WHERE [UpdateTime] > CONVERT(VARCHAR,GETDATE(),112) AND [RefPrz] > 0) AS B ON A.股票代號 = B.Symbol
                                      LEFT JOIN
                                      (SELECT  [股票代號],[產業代號],[上市上櫃]
                                      FROM [TwCMData].[dbo].[上市櫃公司基本資料]
                                      WHERE [年度] = (SELECT MAX(年度) FROM [TwCMData].[dbo].[上市櫃公司基本資料])) AS C ON A.股票代號 = C.股票代號
                                      LEFT JOIN 
                                      (SELECT C.[產業代號],C.[上市上櫃],ROUND(AVG(A.漲跌幅+B.盤中漲跌幅),2) AS 產業平均漲跌幅 FROM
                                      (SELECT  [股票代號],SUM([漲幅(%)]) AS 漲跌幅 FROM [TwCMData].[dbo].[日收盤表排行]
                                      WHERE [日期] >= '{lastTradeDate5}' AND LEN([股票代號]) < 5 AND LEFT([股票代號],2) <> '00' GROUP BY [股票代號]) AS A 
                                      LEFT JOIN (SELECT  [Symbol],CASE WHEN [TrdPrz] = 0 THEN [RefPrz] ELSE [TrdPrz] END AS [TrdPrz],[RefPrz],((CASE WHEN [TrdPrz] = 0 THEN [RefPrz] ELSE [TrdPrz] END) - [RefPrz]) / [RefPrz] * 100 AS 盤中漲跌幅
                                      FROM [TwQuoteAddin].[dbo].[STKQUOTE] WHERE [UpdateTime] > CONVERT(VARCHAR,GETDATE(),112) AND [RefPrz] > 0) AS B ON A.股票代號 = B.Symbol
                                      LEFT JOIN( SELECT [股票代號],[產業代號],[上市上櫃] FROM [TwCMData].[dbo].[上市櫃公司基本資料] WHERE [年度] = (SELECT MAX(年度) FROM [TwCMData].[dbo].[上市櫃公司基本資料])) AS C ON A.股票代號 = C.股票代號
                                      GROUP BY C.[產業代號],C.[上市上櫃]) AS D ON C.產業代號 = D.產業代號 AND C.上市上櫃 = D.上市上櫃
                                      WHERE B.Symbol IS NOT NULL";
                
                DataTable dtW1_1 = MSSQL.ExecSqlQry(sqlW1_1, GlobalVar.loginSet.twCMData);

                string sqlW1_2 = $@"SELECT  [股票代號],
                                    [收盤價],
                                    CASE WHEN B.[TrdPrz] = 0 THEN B.[RefPrz] ELSE B.[TrdPrz] END AS [TrdPrz],
                                    ROUND(ABS([收盤價] - (CASE WHEN B.[TrdPrz] = 0 THEN B.[RefPrz] ELSE B.[TrdPrz] END)),2) AS DIFF　
                                    FROM [TwCMData].[dbo].[日收盤表排行]　AS A
                                      LEFT JOIN [TwQuoteAddin].[dbo].[STKQUOTE] AS B ON A.股票代號 = B.Symbol
                                      WHERE [日期] = '{lastTradeDate5}' AND LEN([股票代號]) < 5 AND LEFT([股票代號],2) <> '00'";
                DataTable dtW1_2 = MSSQL.ExecSqlQry(sqlW1_2, GlobalVar.loginSet.twCMData);
                DataRow[] drW1_1_Select = dtW1_1.Select($@"股票代號 = '{uid}'");
                if(drW1_1_Select.Length > 0)
                {
                    double acc = Convert.ToDouble(drW1_1_Select[0][1].ToString());
                    double diff =Convert.ToDouble(drW1_1_Select[0][3].ToString());
                    if (TSEOTC[uid] == "1")
                    {
                        if (Math.Abs(acc) > 32 && Math.Abs(diff) > 20 && Math.Abs(acc-TSEPriceUpDown6) > 20)
                            W1 = "1";
                        if (Math.Abs(acc) > 25 && Math.Abs(diff) > 20 && Math.Abs(acc - TSEPriceUpDown6) > 20)
                        {
                            DataRow[] drW1_2_Select = dtW1_2.Select($@"股票代號 = '{uid}'");
                            if (drW1_2_Select.Length > 0)
                            {
                                double pDiff = Convert.ToDouble(drW1_2_Select[0][3].ToString());
                                if (pDiff > 50)
                                    W1 = "1";
                            }
                        }
                    }
                    else
                    {
                        if (Math.Abs(acc) > 30 && Math.Abs(diff) > 20 && Math.Abs(acc - TSEPriceUpDown6) > 20)
                            W1 = "1";
                        if (Math.Abs(acc) > 23 && Math.Abs(diff) > 20 && Math.Abs(acc - TSEPriceUpDown6) > 20)
                        {
                            DataRow[] drW1_2_Select = dtW1_2.Select($@"股票代號 = '{uid}'");
                            if (drW1_2_Select.Length > 0)
                            {
                                double pDiff = Convert.ToDouble(drW1_2_Select[0][3].ToString());
                                if (pDiff > 40)
                                    W1 = "1";
                            }
                        }
                    }
                }
                //注意2
                //有價證券最近一段期間起迄兩個營業日之收盤價漲跌百分比異常者(三選一)
                //1.最近三十個營業日（含當日）起迄兩個營業日之收盤價漲跌百分比 > 100 %，且符合下列二項條件之一
                //       (1)其漲幅百分比與全體及同類依本款規定計算之平均值的差幅均在85 % 以上，及收盤價 > 開盤參考價者。
                //	(2)其跌幅百分比與全體及同類依本款規定計算之平均值的差幅均在85 % 以上，及收盤價 < 開盤參考價者。
                //2.最近六十個營業日（含當日）起迄兩個營業日之收盤價漲跌百分比 > 130 %，且符合下列二項條件之一
                //       (1)其漲幅百分比與全體及同類依本款規定計算之平均值的差幅均在110 % 以上，及收盤價 > 開盤參考價者。
                //	(2)其跌幅百分比與全體及同類依本款規定計算之平均值的差幅均在110 % 以上，及收盤價 < 開盤參考價者。
                //3.最近九十個營業日（含當日）起迄兩個營業日之收盤價漲跌百分比 > 160 %，且符合下列二項條件之一：
                //	(1)其漲幅百分比與全體及同類依本款規定計算之平均值的差幅均在135 % 以上，及收盤價 > 開盤參考價者。
                //	(2)其跌幅百分比與全體及同類依本款規定計算之平均值的差幅均在135 % 以上，及收盤價 < 開盤參考價者。

                //近30個營業日
                string W2_30 = "0";
                string W2_60 = "0";
                string W2_90 = "0";

                string sqlW2_30 = $@"SELECT A.股票代號,
                                        A.盤中價格,
                                        A.開盤參考價,
                                        A.漲跌幅,
                                        B.平均漲跌幅,
										A.收盤價,
										C.指數漲跌幅
									　FROM (
                                      SELECT  A.股票代號,A.收盤價, C.產業代號,C.上市上櫃,B.RefPrz AS 開盤參考價,B.TrdPrz AS 盤中價格,ROUND(ABS(CASE WHEN [收盤價] > 0 THEN (B.TrdPrz - [收盤價]) / [收盤價] * 100 ELSE 0 END),2) AS 漲跌幅
                                      FROM [TwCMData].[dbo].[日收盤表排行] AS A
                                      LEFT JOIN (SELECT  [Symbol],[RefPrz],CASE WHEN [TrdPrz] = 0 THEN [RefPrz] ELSE [TrdPrz] END AS [TrdPrz] FROM [TwQuoteAddin].[dbo].[STKQUOTE] WHERE [UpdateTime] > CONVERT(VARCHAR,GETDATE(),112)) AS B ON A.股票代號 = B.Symbol
                                      LEFT JOIN( SELECT [股票代號],[產業代號],[上市上櫃] FROM [TwCMData].[dbo].[上市櫃公司基本資料] WHERE [年度] = (SELECT MAX(年度) FROM [TwCMData].[dbo].[上市櫃公司基本資料])) AS C ON A.股票代號 = C.股票代號
                                      WHERE [日期] = '{lastTradeDate29}'  AND LEN(A.[股票代號]) < 5 AND LEFT(A.[股票代號],2) <> '00' AND C.產業代號 IS NOT NULL AND B.RefPrz > 0 ) AS A
                                      LEFT JOIN 
                                      (SELECT  C.產業代號,ROUND(AVG(ABS(CASE WHEN [收盤價] > 0 THEN (B.TrdPrz - [收盤價]) / [收盤價] * 100 ELSE 0 END)),2) AS 平均漲跌幅
                                      FROM [TwCMData].[dbo].[日收盤表排行] AS A
                                      LEFT JOIN (SELECT  [Symbol],CASE WHEN [TrdPrz] = 0 THEN [RefPrz] ELSE [TrdPrz] END AS [TrdPrz] FROM [TwQuoteAddin].[dbo].[STKQUOTE] WHERE [UpdateTime] > CONVERT(VARCHAR,GETDATE(),112)) AS B ON A.股票代號 = B.Symbol
                                      LEFT JOIN( SELECT [股票代號],[產業代號],[上市上櫃] FROM [TwCMData].[dbo].[上市櫃公司基本資料] WHERE [年度] = (SELECT MAX(年度) FROM [TwCMData].[dbo].[上市櫃公司基本資料])) AS C ON A.股票代號 = C.股票代號
                                      WHERE [日期] = '{lastTradeDate29}'  AND LEN(A.[股票代號]) < 5 AND LEFT(A.[股票代號],2) <> '00' AND C.產業代號 IS NOT NULL
                                      GROUP BY C.[產業代號]) AS B ON A.產業代號 = B.產業代號

									  LEFT JOIN 
                                      (SELECT  C.上市上櫃,ROUND((AVG(ABS(CASE WHEN [收盤價] > 0 THEN (B.TrdPrz - [收盤價]) / [收盤價] * 100 ELSE 0 END) * 100)),2) AS 指數漲跌幅
                                      FROM [TwCMData].[dbo].[日收盤表排行] AS A
                                      LEFT JOIN (SELECT  [Symbol],CASE WHEN [TrdPrz] = 0 THEN [RefPrz] ELSE [TrdPrz] END AS [TrdPrz] FROM [TwQuoteAddin].[dbo].[STKQUOTE] WHERE [UpdateTime] > CONVERT(VARCHAR,GETDATE(),112)) AS B ON A.股票代號 = B.Symbol
                                      LEFT JOIN( SELECT [股票代號],[產業代號],[上市上櫃] FROM [TwCMData].[dbo].[上市櫃公司基本資料] WHERE [年度] = (SELECT MAX(年度) FROM [TwCMData].[dbo].[上市櫃公司基本資料])) AS C ON A.股票代號 = C.股票代號
                                      WHERE [日期] = '{lastTradeDate29}'  AND LEN(A.[股票代號]) < 5 AND LEFT(A.[股票代號],2) <> '00' 
                                      GROUP BY C.上市上櫃) AS C ON A.上市上櫃 = C.上市上櫃";
                DataTable drW2_30 = MSSQL.ExecSqlQry(sqlW2_30, GlobalVar.loginSet.twCMData);
                DataRow[] drW2_30_Select = drW2_30.Select($@"股票代號 = '{uid}'");
                if (drW2_30_Select.Length > 0)
                {
                    double P = Convert.ToDouble(drW2_30_Select[0][1].ToString());
                    double refP = Convert.ToDouble(drW2_30_Select[0][2].ToString());
                    double R = Convert.ToDouble(drW2_30_Select[0][3].ToString());
                    double avgR = Convert.ToDouble(drW2_30_Select[0][4].ToString());
                    double PClose = Convert.ToDouble(drW2_30_Select[0][5].ToString());
                    double avgIndexR = Convert.ToDouble(drW2_30_Select[0][6].ToString());
                    if (TSEOTC[uid] == "1")
                    {
                        if (R > 100 && Math.Abs(R - avgR) > 85 && Math.Abs(R - avgIndexR) > 85 && ((P > refP && P > PClose) || (P < refP && P < PClose)))
                        {
                            W2_30 = "1";
                        }
                    }
                    else
                    {
                        if(P >= 5)
                        {
                            if (R > 100 && Math.Abs(R - avgR) > 80 && Math.Abs(R - avgIndexR) > 80 && ((P > refP && P > PClose) || (P < refP && P < PClose)))
                            {
                                W2_30 = "1";
                            }
                        }
                        else
                        {
                            if (R > 120 && Math.Abs(R - avgR) > 80 && Math.Abs(R - avgIndexR) > 80 && ((P > refP && P > PClose) || (P < refP && P < PClose)))
                            {
                                W2_30 = "1";
                            }
                        }
                    }
                        
                    
                }
                string sqlW2_60 = $@"SELECT A.股票代號,
                                        A.盤中價格,
                                        A.開盤參考價,
                                        A.漲跌幅,
                                        B.平均漲跌幅,
										A.收盤價,
										C.指數漲跌幅
									　FROM (
                                      SELECT  A.股票代號,A.收盤價, C.產業代號,C.上市上櫃,B.RefPrz AS 開盤參考價,B.TrdPrz AS 盤中價格,ROUND(ABS(CASE WHEN [收盤價] > 0 THEN (B.TrdPrz - [收盤價]) / [收盤價] * 100 ELSE 0 END),2) AS 漲跌幅
                                      FROM [TwCMData].[dbo].[日收盤表排行] AS A
                                      LEFT JOIN (SELECT  [Symbol],[RefPrz],CASE WHEN [TrdPrz] = 0 THEN [RefPrz] ELSE [TrdPrz] END AS [TrdPrz] FROM [TwQuoteAddin].[dbo].[STKQUOTE] WHERE [UpdateTime] > CONVERT(VARCHAR,GETDATE(),112)) AS B ON A.股票代號 = B.Symbol
                                      LEFT JOIN( SELECT [股票代號],[產業代號],[上市上櫃] FROM [TwCMData].[dbo].[上市櫃公司基本資料] WHERE [年度] = (SELECT MAX(年度) FROM [TwCMData].[dbo].[上市櫃公司基本資料])) AS C ON A.股票代號 = C.股票代號
                                      WHERE [日期] = '{lastTradeDate59}'  AND LEN(A.[股票代號]) < 5 AND LEFT(A.[股票代號],2) <> '00' AND C.產業代號 IS NOT NULL AND B.RefPrz > 0 ) AS A
                                      LEFT JOIN 
                                      (SELECT  C.產業代號,ROUND(AVG(ABS(CASE WHEN [收盤價] > 0 THEN (B.TrdPrz - [收盤價]) / [收盤價] * 100 ELSE 0 END)),2) AS 平均漲跌幅
                                      FROM [TwCMData].[dbo].[日收盤表排行] AS A
                                      LEFT JOIN (SELECT  [Symbol],CASE WHEN [TrdPrz] = 0 THEN [RefPrz] ELSE [TrdPrz] END AS [TrdPrz] FROM [TwQuoteAddin].[dbo].[STKQUOTE] WHERE [UpdateTime] > CONVERT(VARCHAR,GETDATE(),112)) AS B ON A.股票代號 = B.Symbol
                                      LEFT JOIN( SELECT [股票代號],[產業代號],[上市上櫃] FROM [TwCMData].[dbo].[上市櫃公司基本資料] WHERE [年度] = (SELECT MAX(年度) FROM [TwCMData].[dbo].[上市櫃公司基本資料])) AS C ON A.股票代號 = C.股票代號
                                      WHERE [日期] = '{lastTradeDate59}'  AND LEN(A.[股票代號]) < 5 AND LEFT(A.[股票代號],2) <> '00' AND C.產業代號 IS NOT NULL
                                      GROUP BY C.[產業代號]) AS B ON A.產業代號 = B.產業代號

									  LEFT JOIN 
                                      (SELECT  C.上市上櫃,ROUND((AVG(ABS(CASE WHEN [收盤價] > 0 THEN (B.TrdPrz - [收盤價]) / [收盤價] * 100 ELSE 0 END) * 100)),2) AS 指數漲跌幅
                                      FROM [TwCMData].[dbo].[日收盤表排行] AS A
                                      LEFT JOIN (SELECT  [Symbol],CASE WHEN [TrdPrz] = 0 THEN [RefPrz] ELSE [TrdPrz] END AS [TrdPrz] FROM [TwQuoteAddin].[dbo].[STKQUOTE] WHERE [UpdateTime] > CONVERT(VARCHAR,GETDATE(),112)) AS B ON A.股票代號 = B.Symbol
                                      LEFT JOIN( SELECT [股票代號],[產業代號],[上市上櫃] FROM [TwCMData].[dbo].[上市櫃公司基本資料] WHERE [年度] = (SELECT MAX(年度) FROM [TwCMData].[dbo].[上市櫃公司基本資料])) AS C ON A.股票代號 = C.股票代號
                                      WHERE [日期] = '{lastTradeDate59}'  AND LEN(A.[股票代號]) < 5 AND LEFT(A.[股票代號],2) <> '00' 
                                      GROUP BY C.上市上櫃) AS C ON A.上市上櫃 = C.上市上櫃";
                DataTable drW2_60 = MSSQL.ExecSqlQry(sqlW2_60, GlobalVar.loginSet.twCMData);
                DataRow[] drW2_60_Select = drW2_60.Select($@"股票代號 = '{uid}'");
                if (drW2_60_Select.Length > 0)
                {
                    double P = Convert.ToDouble(drW2_60_Select[0][1].ToString());
                    double refP = Convert.ToDouble(drW2_60_Select[0][2].ToString());
                    double R = Convert.ToDouble(drW2_60_Select[0][3].ToString());
                    double avgR = Convert.ToDouble(drW2_60_Select[0][4].ToString());
                    double PClose = Convert.ToDouble(drW2_60_Select[0][5].ToString());
                    double avgIndexR = Convert.ToDouble(drW2_60_Select[0][6].ToString());
                    if (TSEOTC[uid] == "1")
                    {
                        if (R > 130 && Math.Abs(R - avgR) > 110 && Math.Abs(R - avgIndexR) > 110 && ((P > refP && P > PClose) || (P < refP && P < PClose)))
                        {
                            W2_60 = "1";
                        }
                    }
                    else
                    {
                        if (P >= 5)
                        {
                            if (R > 140 && Math.Abs(R - avgR) > 80 && Math.Abs(R - avgIndexR) > 80 && ((P > refP && P > PClose) || (P < refP && P < PClose)))
                            {
                                W2_60 = "1";
                            }
                        }
                        else
                        {
                            if (R > 160 && Math.Abs(R - avgR) > 80 && Math.Abs(R - avgIndexR) > 80 && ((P > refP && P > PClose) || (P < refP && P < PClose)))
                            {
                                W2_60 = "1";
                            }
                        }
                    }
                }

                string sqlW2_90 = $@"SELECT A.股票代號,
                                        A.盤中價格,
                                        A.開盤參考價,
                                        A.漲跌幅,
                                        B.平均漲跌幅,
										A.收盤價,
										C.指數漲跌幅
									　FROM (
                                      SELECT  A.股票代號,A.收盤價, C.產業代號,C.上市上櫃,B.RefPrz AS 開盤參考價,B.TrdPrz AS 盤中價格,ROUND(ABS(CASE WHEN [收盤價] > 0 THEN (B.TrdPrz - [收盤價]) / [收盤價] * 100 ELSE 0 END),2) AS 漲跌幅
                                      FROM [TwCMData].[dbo].[日收盤表排行] AS A
                                      LEFT JOIN (SELECT  [Symbol],[RefPrz],CASE WHEN [TrdPrz] = 0 THEN [RefPrz] ELSE [TrdPrz] END AS [TrdPrz] FROM [TwQuoteAddin].[dbo].[STKQUOTE] WHERE [UpdateTime] > CONVERT(VARCHAR,GETDATE(),112)) AS B ON A.股票代號 = B.Symbol
                                      LEFT JOIN( SELECT [股票代號],[產業代號],[上市上櫃] FROM [TwCMData].[dbo].[上市櫃公司基本資料] WHERE [年度] = (SELECT MAX(年度) FROM [TwCMData].[dbo].[上市櫃公司基本資料])) AS C ON A.股票代號 = C.股票代號
                                      WHERE [日期] = '{lastTradeDate89}'  AND LEN(A.[股票代號]) < 5 AND LEFT(A.[股票代號],2) <> '00' AND C.產業代號 IS NOT NULL AND B.RefPrz > 0 ) AS A
                                      LEFT JOIN 
                                      (SELECT  C.產業代號,ROUND(AVG(ABS(CASE WHEN [收盤價] > 0 THEN (B.TrdPrz - [收盤價]) / [收盤價] * 100 ELSE 0 END)),2) AS 平均漲跌幅
                                      FROM [TwCMData].[dbo].[日收盤表排行] AS A
                                      LEFT JOIN (SELECT  [Symbol],CASE WHEN [TrdPrz] = 0 THEN [RefPrz] ELSE [TrdPrz] END AS [TrdPrz] FROM [TwQuoteAddin].[dbo].[STKQUOTE] WHERE [UpdateTime] > CONVERT(VARCHAR,GETDATE(),112)) AS B ON A.股票代號 = B.Symbol
                                      LEFT JOIN( SELECT [股票代號],[產業代號],[上市上櫃] FROM [TwCMData].[dbo].[上市櫃公司基本資料] WHERE [年度] = (SELECT MAX(年度) FROM [TwCMData].[dbo].[上市櫃公司基本資料])) AS C ON A.股票代號 = C.股票代號
                                      WHERE [日期] = '{lastTradeDate89}'  AND LEN(A.[股票代號]) < 5 AND LEFT(A.[股票代號],2) <> '00' AND C.產業代號 IS NOT NULL
                                      GROUP BY C.[產業代號]) AS B ON A.產業代號 = B.產業代號

									  LEFT JOIN 
                                      (SELECT  C.上市上櫃,ROUND((AVG(ABS(CASE WHEN [收盤價] > 0 THEN (B.TrdPrz - [收盤價]) / [收盤價] * 100 ELSE 0 END) * 100)),2) AS 指數漲跌幅
                                      FROM [TwCMData].[dbo].[日收盤表排行] AS A
                                      LEFT JOIN (SELECT  [Symbol],CASE WHEN [TrdPrz] = 0 THEN [RefPrz] ELSE [TrdPrz] END AS [TrdPrz] FROM [TwQuoteAddin].[dbo].[STKQUOTE] WHERE [UpdateTime] > CONVERT(VARCHAR,GETDATE(),112)) AS B ON A.股票代號 = B.Symbol
                                      LEFT JOIN( SELECT [股票代號],[產業代號],[上市上櫃] FROM [TwCMData].[dbo].[上市櫃公司基本資料] WHERE [年度] = (SELECT MAX(年度) FROM [TwCMData].[dbo].[上市櫃公司基本資料])) AS C ON A.股票代號 = C.股票代號
                                      WHERE [日期] = '{lastTradeDate89}'  AND LEN(A.[股票代號]) < 5 AND LEFT(A.[股票代號],2) <> '00' 
                                      GROUP BY C.上市上櫃) AS C ON A.上市上櫃 = C.上市上櫃";
                DataTable drW2_90 = MSSQL.ExecSqlQry(sqlW2_90, GlobalVar.loginSet.twCMData);
                DataRow[] drW2_90_Select = drW2_90.Select($@"股票代號 = '{uid}'");
                if (drW2_90_Select.Length > 0)
                {
                    double P = Convert.ToDouble(drW2_90_Select[0][1].ToString());
                    double refP = Convert.ToDouble(drW2_90_Select[0][2].ToString());
                    double R = Convert.ToDouble(drW2_90_Select[0][3].ToString());
                    double avgR = Convert.ToDouble(drW2_90_Select[0][4].ToString());
                    double PClose = Convert.ToDouble(drW2_90_Select[0][5].ToString());
                    double avgIndexR = Convert.ToDouble(drW2_90_Select[0][6].ToString());
                    if (TSEOTC[uid] == "1")
                    {
                        if (R > 160 && Math.Abs(R - avgR) > 135 && Math.Abs(R - avgIndexR) > 135 && ((P > refP && P > PClose) || (P < refP && P < PClose)))
                        {
                            W2_90 = "1";
                        }
                    }
                    else
                    {
                        if (P >= 5)
                        {
                            if (R > 160 && Math.Abs(R - avgR) > 80 && Math.Abs(R - avgIndexR) > 80 && ((P > refP && P > PClose) || (P < refP && P < PClose)))
                            {
                                W2_90 = "1";
                            }
                        }
                        else
                        {
                            if (R > 240 && Math.Abs(R - avgR) > 80 && Math.Abs(R - avgIndexR) > 80 && ((P > refP && P > PClose) || (P < refP && P < PClose)))
                            {
                                W2_90 = "1";
                            }
                        }
                    }
                }
                if (W2_30 == "1" || W2_60 == "1" || W2_90 == "1")
                    W2 = "1";
               

                //注意3
                //三、有價證券最近一段期間累積之收盤價漲跌百分比異常，且其當日之成交量較最近一段期間之日平均成交量異常放大者
                //最近六個營業日（含當日）累積之收盤價漲跌百分比超過25 %，且其漲跌百分比與全體及同類依本款規定計算之平均值的差幅均在20 % 以上。
                //且當日之成交量較最近六十個營業日（含當日）之日平均成交量放大為五倍以上，且其放大倍數與全體依本款規定計算之平均值相差四倍以上。



                string sqlW3 = $@"SELECT A.股票代號,
                                            (A.累積成交量 + ISNULL(B.AccTrdQty,0)) / 60 AS 平均成交量,
                                            ISNULL(B.AccTrdQty,0) AS 當日成交量,
                                            CASE WHEN ((A.累積成交量 + ISNULL(B.AccTrdQty,0)) / 60) = 0 THEN 0 ELSE ROUND(ISNULL(CAST(B.AccTrdQty as float),0) / ((A.累積成交量 + ISNULL(B.AccTrdQty,0)) / 60),2) END AS 倍數,
                                            D.平均倍數
                                    FROM (SELECT [股票代號],SUM([成交量]) AS 累積成交量　FROM [TwCMData].[dbo].[日收盤表排行] WHERE [日期] >= '{lastTradeDate59}' AND LEN([股票代號]) < 5 AND LEFT([股票代號],2) <> '00' GROUP BY [股票代號]) AS A
                                    LEFT JOIN [TwQuoteAddin].[dbo].[STKQUOTE] AS B ON A.股票代號 = B.Symbol
                                    LEFT JOIN( SELECT [股票代號],[產業代號],[上市上櫃] FROM [TwCMData].[dbo].[上市櫃公司基本資料] WHERE [年度] = (SELECT MAX(年度) FROM [TwCMData].[dbo].[上市櫃公司基本資料])) AS C ON A.股票代號 = C.股票代號
                                    LEFT JOIN 
                                    (SELECT [上市上櫃],ROUND(AVG(倍數),2) AS 平均倍數 FROM
                                    (SELECT A.股票代號,(A.累積成交量 + ISNULL(B.AccTrdQty,0)) / 60 AS 平均成交量,ISNULL(B.AccTrdQty,0) AS 當日成交量,C.上市上櫃,CASE WHEN ((A.累積成交量 + ISNULL(B.AccTrdQty,0)) / 60) = 0 THEN 0 ELSE ISNULL(CAST(B.AccTrdQty as float),0) / ((A.累積成交量 + ISNULL(B.AccTrdQty,0)) / 60) END AS 倍數
                                    FROM (SELECT [股票代號],SUM([成交量]) AS 累積成交量　FROM [TwCMData].[dbo].[日收盤表排行] WHERE [日期] >= '{lastTradeDate59}' AND LEN([股票代號]) < 5 AND LEFT([股票代號],2) <> '00' GROUP BY [股票代號]) AS A
                                    LEFT JOIN [TwQuoteAddin].[dbo].[STKQUOTE] AS B ON A.股票代號 = B.Symbol
                                    LEFT JOIN( SELECT [股票代號],[產業代號],[上市上櫃] FROM [TwCMData].[dbo].[上市櫃公司基本資料] WHERE [年度] = (SELECT MAX(年度) FROM [TwCMData].[dbo].[上市櫃公司基本資料])) AS C ON A.股票代號 = C.股票代號
                                    WHERE C.產業代號 IS NOT NULL) AS A GROUP BY A.[上市上櫃]) AS D ON C.上市上櫃 = D.上市上櫃
                                    WHERE C.產業代號 IS NOT NULL";
                DataTable dtW3 = MSSQL.ExecSqlQry(sqlW3, GlobalVar.loginSet.twCMData);
                DataRow[] drW3_Select = dtW3.Select($@"股票代號 = '{uid}'");
                if (drW1_1_Select.Length > 0)
                {
                   
                    double acc = Convert.ToDouble(drW1_1_Select[0][1].ToString());
                    double diff = Convert.ToDouble(drW1_1_Select[0][3].ToString());

                    if (drW3_Select.Length > 0)
                    {
                        double R = Convert.ToDouble(drW3_Select[0][3].ToString());
                        double avgR = Convert.ToDouble(drW3_Select[0][4].ToString());
                       
                        if (TSEOTC[uid] == "1")
                        {
                            if (Math.Abs(acc) > 25 && Math.Abs(diff) > 20 && Math.Abs(acc - TSEPriceUpDown6) > 20 && R > 5 && Math.Abs(R - avgR) > 4)
                            {
                                W3 = "1";
                            }
                        }
                        else
                        {
                            if (Math.Abs(acc) > 27 && Math.Abs(diff) > 20 && Math.Abs(acc - TSEPriceUpDown6) > 20 && R > 5 && Math.Abs(R - avgR) > 4)
                            {
                                W3 = "1";
                            }
                        }
                    }
                }

                //注意4
                //四、有價證券最近一段期間累積之收盤價漲跌百分比異常，且其當日之週轉率過高者
                //成交金額(千)] / [總市值(億)] / 1000
                //總市值(億元) = [收盤價] * [股本(百萬)] /[普通股每股面額(元)] / 100
                //最近六個營業日（含當日）累積之收盤價漲跌百分比超過25 %，且其漲跌百分比與全體及同類依本款規定計算之平均值的差幅均在20 % 以上。
                //且當日週轉率10 % 以上，且其週轉率與全體依本款規定計算之平均值的差幅在5 % 以上。

                //周轉率已經是%了 EX 3 代表 3%
                string sqlW4 = $@"SELECT A.股票代號,
                                        ROUND(ISNULL(A.周轉率,0),2) AS 周轉率,
                                        ROUND(ISNULL(B.平均周轉率,0),2) AS 平均周轉率  FROM　(
                                  SELECT  [股票代號],[上市上櫃],CASE WHEN B.TrdPrz = 0 THEN 0 ELSE B.AccTrdAmtK / ((([交易所公告股本(千)]/1000) / [普通股每股面額(元)]) * B.TrdPrz  / 100) / 1000 END AS 周轉率,產業代號
                                  FROM [TwCMData].[dbo].[上市櫃公司基本資料]　AS A
                                  LEFT JOIN　[TwQuoteAddin].[dbo].[STKQUOTE] AS B ON A.股票代號 = B.Symbol
                                  WHERE [年度] = (SELECT MAX(年度)　FROM [TwCMData].[dbo].[上市櫃公司基本資料])) AS A 
                                  LEFT JOIN
                                  (SELECT  [上市上櫃],AVG(CASE WHEN B.TrdPrz = 0 THEN 0 ELSE B.AccTrdAmtK / ((([交易所公告股本(千)]/1000) / [普通股每股面額(元)]) * B.TrdPrz  / 100) / 1000 END) AS 平均周轉率
                                  FROM [TwCMData].[dbo].[上市櫃公司基本資料]　AS A
                                  LEFT JOIN　[TwQuoteAddin].[dbo].[STKQUOTE] AS B ON A.股票代號 = B.Symbol
                                  WHERE [年度] = (SELECT MAX(年度)　FROM [TwCMData].[dbo].[上市櫃公司基本資料])
                                   GROUP BY [上市上櫃]) AS B ON A.上市上櫃 = B.上市上櫃";
                DataTable dtW4 = MSSQL.ExecSqlQry(sqlW4, GlobalVar.loginSet.twCMData);
                DataRow[] drW4_Select = dtW4.Select($@"股票代號 = '{uid}'");
                if (drW1_1_Select.Length > 0)
                {
                   
                    double acc = Convert.ToDouble(drW1_1_Select[0][1].ToString());
                    double diff = Convert.ToDouble(drW1_1_Select[0][3].ToString());

                    if (drW4_Select.Length > 0)
                    {
                        double R = Convert.ToDouble(drW4_Select[0][1].ToString());
                        double avgR = Convert.ToDouble(drW4_Select[0][2].ToString());

                        if (TSEOTC[uid] == "1")
                        {
                            if (Math.Abs(acc) > 25 && Math.Abs(diff) > 20 && Math.Abs(acc - TSEPriceUpDown6) > 20 && R > 10 && Math.Abs(R - avgR) > 5)
                            {
                                W4 = "1";
                            }
                        }
                        else
                        {
                            if (Math.Abs(acc) > 27 && Math.Abs(diff) > 20 && Math.Abs(acc - TSEPriceUpDown6) > 20 && R > 5 && Math.Abs(R - avgR) > 3)
                            {
                                W4 = "1";
                            }
                        }
                    }
                }
                //注意5 日融資券排行
                //五、有價證券最近一段期間累積之收盤價漲跌百分比異常，且券資比明顯放大者(都須滿足)
                //1.最近六個營業日（含當日）累積之收盤價漲跌百分比超過25 %，且其漲跌百分比與全體及同類依本款規定計算之平均值的差幅均在20 % 以上。
                //2.當日之前一個營業日之券資比20 % 以上，且同時符合下列二項條件：融資使用率25 % 以上且融券使用率15 % 以上
                //3.當日之前一個營業日之券資比較最近六個營業日（從當日之前一個營業日起）之最低券資比放大四倍以上
                string sqlW5 = $@"SELECT A.股票代號,
                                        A.券資比,
                                        A.融資使用率,
                                        A.融券使用率,
                                        B.最低券資比 FROM
                                        (SELECT  [股票代號],ISNULL([券資比],0) AS 券資比, ISNULL([資使用率],0) AS 融資使用率,ISNULL([券使用率],0) AS 融券使用率
                                          FROM [TwCMData].[dbo].[日融資券排行]　WHERE [日期] = '{lastTradeDate1}') AS A
                                          LEFT JOIN
                                        (SELECT  [股票代號],　MIN(ISNULL([券資比],0)) AS 最低券資比
                                          FROM [TwCMData].[dbo].[日融資券排行]　WHERE [日期] >= '{lastTradeDate6}' 
                                          GROUP BY [股票代號]) AS B ON A.股票代號 = B.股票代號";
                DataTable dtW5 = MSSQL.ExecSqlQry(sqlW5, GlobalVar.loginSet.twCMData);
                DataRow[] drW5_Select = dtW5.Select($@"股票代號 = '{uid}'");
                if (drW1_1_Select.Length > 0)
                {
                   
                    double acc = Convert.ToDouble(drW1_1_Select[0][1].ToString());
                    double diff = Convert.ToDouble(drW1_1_Select[0][3].ToString());

                    if (drW5_Select.Length > 0)
                    {
                        double R = Convert.ToDouble(drW5_Select[0][1].ToString());
                        double lbR = Convert.ToDouble(drW5_Select[0][2].ToString());//leverage buy
                        double lsR = Convert.ToDouble(drW5_Select[0][3].ToString());//leverage short
                        double minR = Convert.ToDouble(drW5_Select[0][4].ToString());

                        if (TSEOTC[uid] == "1")
                        {
                            if (Math.Abs(acc) > 27 && Math.Abs(diff) > 20 && Math.Abs(acc - TSEPriceUpDown6) > 20 && R > 20 && lbR > 25 && lsR > 15 && (R > minR * 4))
                            {
                                W5 = "1";
                            }
                        }
                        else
                        {
                            if (Math.Abs(acc) > 27 && Math.Abs(diff) > 20 && Math.Abs(acc - TSEPriceUpDown6) > 20 && R > 10 && lbR > 20 && lsR > 10 && (R > minR * 4))
                            {
                                W5 = "1";
                            }
                        }
                    }
                   
                }
                //注意6
                //六、有價證券當日及最近數日之日平均成交量較最近一段期間之日平均成交量明顯放大者
                //最近六個營業日（含當日）之日平均成交量較最近六十個營業日（含當日）之日平均成交量放大為五倍以上，且其放大倍數與全體依本款規定計算之平均值相差四倍以上。
                //且當日之成交量較最近六十個營業日（含當日）之日平均成交量放大為五倍以上，且其放大倍數與全體依本款規定計算之平均值相差四倍以上。
                string sqlW6 = $@"SELECT A.股票代號,
                                        ROUND(A.倍數6除60,2) AS 倍數6除60,
                                        ROUND(A.倍數1除60,2) AS 倍數1除60,
                                        ROUND(B.平均倍數6除60,2) AS 平均倍數6除60數,
                                        ROUND(B.平均倍數1除60,2) AS 平均倍數1除60數

                                        FROM 
									(SELECT T6.股票代號,(T6.累積成交量 + ISNULL(T1.AccTrdQty,0)) / 6 AS T6平均成交量,(T60.累積成交量 + ISNULL(T1.AccTrdQty,0)) / 60 AS T60平均成交量,B.上市上櫃
									,CASE WHEN (T6.累積成交量 + ISNULL(T1.AccTrdQty,0)) = 0 OR (T60.累積成交量 + ISNULL(T1.AccTrdQty,0)) = 0 THEN 0 ELSE CAST((T6.累積成交量 + ISNULL(T1.AccTrdQty,0)) AS float) / (T60.累積成交量 + ISNULL(T1.AccTrdQty,0))　* 10 END AS 倍數6除60
									,CASE WHEN (ISNULL(T1.AccTrdQty,0)) = 0 OR (T60.累積成交量 + ISNULL(T1.AccTrdQty,0)) = 0 THEN 0 ELSE CAST((ISNULL(T1.AccTrdQty,0)) AS float) / (T60.累積成交量 + ISNULL(T1.AccTrdQty,0))　* 60 END AS 倍數1除60
                                    FROM (SELECT [股票代號],SUM([成交量]) AS 累積成交量　FROM [TwCMData].[dbo].[日收盤表排行] WHERE [日期] >= '{lastTradeDate5}'  AND LEN([股票代號]) < 5 AND LEFT([股票代號],2) <> '00' GROUP BY [股票代號]) AS T6
									LEFT JOIN (SELECT [股票代號],SUM([成交量]) AS 累積成交量　FROM [TwCMData].[dbo].[日收盤表排行] WHERE [日期] >= '{lastTradeDate59}'  AND LEN([股票代號]) < 5 AND LEFT([股票代號],2) <> '00' GROUP BY [股票代號]) AS T60 ON T6.股票代號 = T60.股票代號
                                    LEFT JOIN [TwQuoteAddin].[dbo].[STKQUOTE] AS T1 ON T6.股票代號 = T1.Symbol
                                    LEFT JOIN( SELECT [股票代號],[產業代號],[上市上櫃] FROM [TwCMData].[dbo].[上市櫃公司基本資料] WHERE [年度] = (SELECT MAX(年度) FROM [TwCMData].[dbo].[上市櫃公司基本資料])) AS B ON T6.股票代號 = B.股票代號
                                    WHERE B.產業代號 IS NOT NULL) AS A
									LEFT JOIN 
									(SELECT B.上市上櫃,AVG(CASE WHEN (T6.累積成交量 + ISNULL(T1.AccTrdQty,0)) = 0 OR (T60.累積成交量 + ISNULL(T1.AccTrdQty,0)) = 0 THEN 0 ELSE CAST((T6.累積成交量 + ISNULL(T1.AccTrdQty,0)) AS float) / (T60.累積成交量 + ISNULL(T1.AccTrdQty,0))　* 10 END) AS 平均倍數6除60
									,AVG(CASE WHEN (ISNULL(T1.AccTrdQty,0)) = 0 OR (T60.累積成交量 + ISNULL(T1.AccTrdQty,0)) = 0 THEN 0 ELSE CAST((ISNULL(T1.AccTrdQty,0)) AS float) / (T60.累積成交量 + ISNULL(T1.AccTrdQty,0))　* 60 END) AS 平均倍數1除60
                                    FROM (SELECT [股票代號],SUM([成交量]) AS 累積成交量　FROM [TwCMData].[dbo].[日收盤表排行] WHERE [日期] >= '{lastTradeDate5}'  AND LEN([股票代號]) < 5 AND LEFT([股票代號],2) <> '00' GROUP BY [股票代號]) AS T6
									LEFT JOIN (SELECT [股票代號],SUM([成交量]) AS 累積成交量　FROM [TwCMData].[dbo].[日收盤表排行] WHERE [日期] >= '{lastTradeDate59}' AND LEN([股票代號]) < 5 AND LEFT([股票代號],2) <> '00' GROUP BY [股票代號]) AS T60 ON T6.股票代號 = T60.股票代號
                                    LEFT JOIN [TwQuoteAddin].[dbo].[STKQUOTE] AS T1 ON T6.股票代號 = T1.Symbol
                                    LEFT JOIN( SELECT [股票代號],[產業代號],[上市上櫃] FROM [TwCMData].[dbo].[上市櫃公司基本資料] WHERE [年度] = (SELECT MAX(年度) FROM [TwCMData].[dbo].[上市櫃公司基本資料])) AS B ON T6.股票代號 = B.股票代號
                                    WHERE B.產業代號 IS NOT NULL GROUP BY B.上市上櫃) AS B ON A.上市上櫃 = B.上市上櫃";
                DataTable dtW6 = MSSQL.ExecSqlQry(sqlW6, GlobalVar.loginSet.twCMData);
                DataRow[] drW6_Select = dtW6.Select($@"股票代號 = '{uid}'");
                if (drW6_Select.Length > 0)
                {
                    double R6_60 = Convert.ToDouble(drW6_Select[0][1].ToString());
                    double R1_60 = Convert.ToDouble(drW6_Select[0][2].ToString());
                    double avgR6_60 = Convert.ToDouble(drW6_Select[0][3].ToString());
                    double avgR1_60 = Convert.ToDouble(drW6_Select[0][4].ToString());
                    if(TSEOTC[uid] == "1")
                    {
                        if(R6_60 > 5 && Math.Abs(R6_60 - avgR6_60) > 4 && R1_60 > 5 && Math.Abs(R1_60-avgR1_60) > 4)
                        {
                            W6 = "1";
                        }
                    }
                    else
                    {
                        if (R6_60 > 5 && Math.Abs(R6_60 - avgR6_60) > 4 && R1_60 > 5 && Math.Abs(R1_60 - avgR1_60) > 4)
                        {
                            W6 = "1";
                        }
                    }
                        
                }
                
                //注意7
                //七、有價證券最近一段期間之累積週轉率明顯過高者
                //最近六個營業日（含當日）之累積週轉率超過百分之八十，且其累積週轉率與全體依本款規定計算之平均值的差幅在百分之五十以上。
                //且當日週轉率百分之五以上，且其週轉率與全體依本款規定計算之平均值的差幅在百分之三以上。

                string sqlW7_1 = $@"SELECT A.股票代號,
                                    ROUND(A.周轉率,2) AS 周轉率,
                                    ROUND(A.累加周轉率,2) AS 累加周轉率,
                                    ROUND(B.平均周轉率,2) AS 平均周轉率,
                                    ROUND(B.平均累加周轉率,2) AS 平均累加周轉率 FROM (
                                SELECT A.股票代號,(A.累加周轉率 + ISNULL(B.周轉率,0)) AS 累加周轉率,(ISNULL(B.周轉率,0)) AS 周轉率,B.上市上櫃 FROM (
                                SELECT  [股票代號],SUM([週轉率(%)])　AS 累加周轉率
                                  FROM [TwCMData].[dbo].[日收盤表排行] WHERE [日期] >= '{lastTradeDate5}'  AND LEN([股票代號]) < 5 AND LEFT([股票代號],2) <> '00'
                                  GROUP BY [股票代號])　AS A LEFT JOIN 
                                  (SELECT  [股票代號],CASE WHEN B.TrdPrz = 0 THEN 0 ELSE B.AccTrdAmtK / ((([交易所公告股本(千)]/1000) / [普通股每股面額(元)]) * B.TrdPrz  / 100) / 1000 END AS 周轉率,上市上櫃
                                   FROM [TwCMData].[dbo].[上市櫃公司基本資料]　AS A LEFT JOIN　[TwQuoteAddin].[dbo].[STKQUOTE] AS B ON A.股票代號 = B.Symbol WHERE [年度] = (SELECT MAX(年度)　FROM [TwCMData].[dbo].[上市櫃公司基本資料])) AS B ON A.股票代號 = B.股票代號) AS A
                                   LEFT JOIN
                                   (SELECT B.上市上櫃,AVG(A.累加周轉率 + B.周轉率) AS 平均累加周轉率,AVG(ISNULL(B.周轉率,0)) AS 平均周轉率 FROM (
　　                                SELECT  [股票代號],SUM([週轉率(%)])　AS 累加周轉率　FROM [TwCMData].[dbo].[日收盤表排行]
  　                                WHERE [日期] >= '{lastTradeDate5}'  AND LEN([股票代號]) < 5 AND LEFT([股票代號],2) <> '00'　GROUP BY [股票代號])　AS A
  　                                LEFT JOIN 
  　                                (SELECT  [股票代號],CASE WHEN B.TrdPrz = 0 THEN 0 ELSE B.AccTrdAmtK / ((([交易所公告股本(千)]/1000) / [普通股每股面額(元)]) * B.TrdPrz  / 100) / 1000 END AS 周轉率,上市上櫃
                                   FROM [TwCMData].[dbo].[上市櫃公司基本資料]　AS A LEFT JOIN　[TwQuoteAddin].[dbo].[STKQUOTE] AS B ON A.股票代號 = B.Symbol WHERE [年度] = (SELECT MAX(年度)　FROM [TwCMData].[dbo].[上市櫃公司基本資料])) AS B ON A.股票代號 = B.股票代號
                                   GROUP BY B.上市上櫃) AS B ON A.上市上櫃 = B.上市上櫃";
                DataTable dtW7_1 = MSSQL.ExecSqlQry(sqlW7_1, GlobalVar.loginSet.twCMData);
                /*
                string sqlW7_2 = $@"SELECT A.股票代號,B.產業代號,A.周轉率,B.平均周轉率 FROM
                                   (SELECT [股票代號],[產業代號],ROUND(ISNULL(CASE WHEN B.TrdPrz = 0 THEN 0 ELSE B.AccTrdAmtK / ((([交易所公告股本(千)]/1000) / [普通股每股面額(元)]) * B.TrdPrz  / 100) / 1000 END,0),2) AS 周轉率
                                   FROM [TwCMData].[dbo].[上市櫃公司基本資料] AS A
                                   LEFT JOIN　[TwQuoteAddin].[dbo].[STKQUOTE] AS B ON A.[股票代號] = B.[Symbol]
                                   WHERE [年度] = (SELECT MAX(年度)　FROM [TwCMData].[dbo].[上市櫃公司基本資料])) AS A
                                   LEFT JOIN
                                   (SELECT [產業代號],ROUND(AVG(ISNULL(CASE WHEN B.TrdPrz = 0 THEN 0 ELSE B.AccTrdAmtK / ((([交易所公告股本(千)]/1000) / [普通股每股面額(元)]) * B.TrdPrz  / 100) / 1000 END,0)),2) AS 平均周轉率
                                   FROM [TwCMData].[dbo].[上市櫃公司基本資料] AS A
                                   LEFT JOIN　[TwQuoteAddin].[dbo].[STKQUOTE] AS B ON A.[股票代號] = B.[Symbol]
                                   WHERE [年度] = (SELECT MAX(年度)　FROM [TwCMData].[dbo].[上市櫃公司基本資料])
                                   GROUP BY A.產業代號) AS B ON A.產業代號 = B.產業代號";
                DataTable dtW7_2 = MSSQL.ExecSqlQry(sqlW7_2, GlobalVar.loginSet.twCMData);
                */


                DataRow[] drW7_1_Select = dtW7_1.Select($@"股票代號 = '{uid}'");
                //DataRow[] drW7_2_Select = dtW7_2.Select($@"股票代號 = '{uid}'");
                if (drW7_1_Select.Length > 0)
                {
                    double T6R = Convert.ToDouble(drW7_1_Select[0][2].ToString());
                    double avgT6R = Convert.ToDouble(drW7_1_Select[0][4].ToString());
                    double R = Convert.ToDouble(drW7_1_Select[0][1].ToString());
                    double avgR = Convert.ToDouble(drW7_1_Select[0][3].ToString());
                    if (TSEOTC[uid] == "1")
                    {
                        if (T6R > 50 && Math.Abs(T6R - avgT6R) > 40 && R > 10 && Math.Abs(R - avgR) > 5)
                            W7 = "1";
                    }
                    else
                    {
                        if (T6R > 80 && Math.Abs(T6R - avgT6R) > 50 && R > 5 && Math.Abs(R - avgR) > 3)
                            W7 = "1";
                    }
                }


                //注意8
                //八、有價證券最近一段期間起迄兩個營業日之收盤價價差異常者(二選一)

                string sqlW8 = $@"SELECT A.股票代號,
                                    A.起日收盤價,
                                    B.近5日累積成交量,
                                    CASE WHEN B.近五日最大收盤價 = 0 OR B.近五日最大收盤價 IS NULL THEN C.RefPrz ELSE B.近五日最大收盤價 END AS 近五日最大收盤價,
                                    CASE WHEN B.近五日最小收盤價 = 0 OR B.近五日最小收盤價 IS NULL THEN C.RefPrz ELSE B.近五日最小收盤價 END AS 近五日最小收盤價,
                                    C.RefPrz,
                                    C.TrdPrz FROM (
                                    SELECT [股票代號],[收盤價] AS 起日收盤價
                                      FROM [TwCMData].[dbo].[日收盤表排行]  WHERE [日期] = '{lastTradeDate5}' AND LEN([股票代號]) < 5 AND LEFT([股票代號],2) <> '00') AS A LEFT JOIN
                                      (SELECT  [股票代號],SUM(成交量) AS 近5日累積成交量,MAX(收盤價) AS 近五日最大收盤價,MIN(收盤價) AS 近五日最小收盤價
                                      FROM [TwCMData].[dbo].[日收盤表排行]
                                      WHERE [日期] >= '{lastTradeDate5}' AND [日期] < CONVERT(VARCHAR,GETDATE(),112) AND LEN([股票代號]) < 5 AND LEFT([股票代號],2) <> '00'
                                      GROUP BY [股票代號]) AS B ON A.股票代號 = B.股票代號
                                      LEFT JOIN [TwQuoteAddin].[dbo].[STKQUOTE] AS C ON A.股票代號 = C.Symbol";
                DataTable dtW8 = MSSQL.ExecSqlQry(sqlW8, GlobalVar.loginSet.twCMData);
                DataRow[] drW8_Select = dtW8.Select($@"股票代號 = '{uid}'");
                if (drW8_Select.Length > 0)
                {
                    double p_now = Convert.ToDouble(drW8_Select[0][6].ToString());
                    double p_start = Convert.ToDouble(drW8_Select[0][1].ToString());
                    double p_max = Convert.ToDouble(drW8_Select[0][3].ToString());
                    double p_min = Convert.ToDouble(drW8_Select[0][4].ToString());
                    if(TSEOTC[uid] == "1")
                    {
                        if (Math.Abs(p_now - p_start) > 100 && (p_now < p_min || p_now > p_max))
                            W8 = "1";
                    }
                    else
                    {
                        if (Math.Abs(p_now - p_start) > 70 && (p_now < p_min || p_now > p_max))
                            W8 = "1";
                    }
                }
                //1.最近六個營業日（含當日）起迄兩個營業日之收盤價價差達100以上，且當日收盤價須為最近六個營業日（含當日）收盤價最高者。
                //但最近五個營業日（不含當日）無收盤價時，則當日收盤價須高於開盤參考價。
                //2.最近六個營業日（含當日）起迄兩個營業日之收盤價價差達100以上，且當日收盤價須為最近六個營業日（含當日）收盤價最低者。
                //但最近五個營業日（不含當日）無收盤價時，則當日收盤價須低於開盤參考價。

                //注意9
                //九、最近一段期間之借券賣出成交量占總成交量比率明顯過高者
                //最近六個營業日（自當日之前一個營業日起）之借券賣出成交量占最近六個營業日（自當日之前一個營業日起）總成交量比率超過百分之九。
                //且當日之前一個營業日借券賣出成交量較最近六十個營業日（自當日之前一個營業日起）之日平均借券賣出成交量放大為四倍以上。


                string sqlW9_1 = $@"SELECT  A.[股票代號],
                                            SUM([成交量]) AS 成交量,
                                            ISNULL(SUM(B.借券賣出),0) AS 借券賣出,
                                            CASE WHEN SUM([成交量]) = 0 THEN 0 ELSE CAST(ISNULL(SUM(B.借券賣出),0) AS float) / SUM([成交量]) END AS 比率
                                      FROM [TwCMData].[dbo].[日收盤表排行]　AS A LEFT JOIN[TwCMData].[dbo].[日融資券排行] AS B ON　A.日期 = B.日期 AND A.股票代號 = B.股票代號
                                      WHERE A.[日期] >= '{lastTradeDate6}' AND A.[日期] < CONVERT(VARCHAR,GETDATE(),112) AND LEN(A.[股票代號]) < 5 AND LEFT(A.[股票代號],2) <> '00' 
                                      GROUP BY A.股票代號";
                DataTable dtW9_1 = MSSQL.ExecSqlQry(sqlW9_1, GlobalVar.loginSet.twCMData);
                DataRow[] drW9_1_Select = dtW9_1.Select($@"股票代號 = '{uid}'");
                string sqlW9_2 = $@"SELECT A.股票代號,
                                    A.借券賣出,
                                    B.平均借券賣出60  FROM
                                    (SELECT  [股票代號],ISNULL([借券賣出],0) AS [借券賣出]
                                      FROM [TwCMData].[dbo].[日融資券排行] WHERE [日期] = '{lastTradeDate1}' AND LEN([股票代號]) < 5 AND LEFT([股票代號],2) <> '00') AS A
                                    LEFT JOIN
                                      (SELECT  [股票代號],AVG(ISNULL([借券賣出],0)) AS 平均借券賣出60
                                      FROM [TwCMData].[dbo].[日融資券排行] WHERE [日期] >= '{lastTradeDate60}' AND LEN([股票代號]) < 5 AND LEFT([股票代號],2) <> '00'
                                      GROUP BY [股票代號]) AS B ON A.股票代號 = B.股票代號";
                DataTable dtW9_2 = MSSQL.ExecSqlQry(sqlW9_2, GlobalVar.loginSet.twCMData);
                DataRow[] drW9_2_Select = dtW9_2.Select($@"股票代號 = '{uid}'");
                if(drW9_1_Select.Length>0 && drW9_2_Select.Length > 0)
                {
                    double W9R1 = Math.Round(Convert.ToDouble(drW9_1_Select[0][3].ToString()) * 100, 2);
                    double W9R2_1 = Convert.ToDouble(drW9_2_Select[0][1].ToString());
                    double W9R2_60 = Convert.ToDouble(drW9_2_Select[0][2].ToString());
                    double W9R2 = 0;
                    if(W9R2_1 > 0)
                    {
                        W9R2 = Math.Round(W9R2_1 / W9R2_60, 2);
                    }
                    if(TSEOTC[uid] == "1")
                    {
                        if (W9R1 > 12 && W9R2 > 5)
                            W9 = "1";
                    }
                    else
                    {
                        if (W9R1 > 9 && W9R2 > 4)
                            W9 = "1";
                    }
                }
                //注意10
                
                 string sqlW10_1 = $@"SELECT  A.[股票代號],
                                            SUM([成交量]) AS 成交量,
                                            ISNULL(SUM(B.當沖成交量),0) AS 當沖量,
                                            CASE WHEN SUM([成交量]) = 0 THEN 0 ELSE CAST(ISNULL(SUM(B.當沖成交量),0) AS float) / SUM([成交量]) END AS 比率
                                      FROM [TwCMData].[dbo].[日收盤表排行]　AS A 
									  LEFT JOIN　[TwCMData].[dbo].[個股當日沖銷交易] AS B ON　A.日期 = B.日期 AND A.股票代號 = B.股票代號
                                      WHERE A.[日期] >= '{lastTradeDate6}' AND A.[日期] < CONVERT(VARCHAR,GETDATE(),112) AND LEN(A.[股票代號]) < 5 AND LEFT(A.[股票代號],2) <> '00' 
                                      GROUP BY A.股票代號";
                DataTable dtW10_1 = MSSQL.ExecSqlQry(sqlW10_1, GlobalVar.loginSet.twCMData);
                DataRow[] drW10_1_Select = dtW10_1.Select($@"股票代號 = '{uid}'");
                string sqlW10_2 = $@"SELECT  A.[股票代號],
                                            SUM([成交量]) AS 成交量,
                                            ISNULL(SUM(B.當沖成交量),0) AS 當沖量,
                                            CASE WHEN SUM([成交量]) = 0 THEN 0 ELSE CAST(ISNULL(SUM(B.當沖成交量),0) AS float) / SUM([成交量]) END AS 比率
                                      FROM [TwCMData].[dbo].[日收盤表排行]　AS A 
									  LEFT JOIN　[TwCMData].[dbo].[個股當日沖銷交易] AS B ON　A.日期 = B.日期 AND A.股票代號 = B.股票代號
                                      WHERE A.[日期] >= '{lastTradeDate1}' AND A.[日期] < CONVERT(VARCHAR,GETDATE(),112) AND LEN(A.[股票代號]) < 5 AND LEFT(A.[股票代號],2) <> '00' 
                                      GROUP BY A.股票代號";
                DataTable dtW10_2 = MSSQL.ExecSqlQry(sqlW10_2, GlobalVar.loginSet.twCMData);
                DataRow[] drW10_2_Select = dtW10_2.Select($@"股票代號 = '{uid}'");
                if (drW10_1_Select.Length>0 && drW10_2_Select.Length > 0)
                {
                    double W10_D1 = Math.Round(Convert.ToDouble(drW10_1_Select[0][3].ToString()), 2);
                    double W10_D6 = Math.Round(Convert.ToDouble(drW10_2_Select[0][3].ToString()), 2);

                    //MessageBox.Show($@"{uid} {W10_D1} {W10_D6} {W10}");
                    if ((W10_D1 > 0.6) && (W10_D6  > 0.6))
                    {
                        
                        W10 = "1";
                    }
                    //MessageBox.Show($@"{uid} {W10_D1} {W10_D6} {W10}");
                    //MessageBox.Show(W10.ToString());  
                }


              

                dr["累積漲跌異常"] = W1;
                dr["起訖日漲跌異常"] = W2;
                dr["成交量過高"] = W3;
                dr["周轉率過高"] = W4;
                dr["券資比過高"] = W5;
                dr["平均成交量過高"] = W6;
                dr["累積周轉率過高"] = W7;
                dr["股價價差異常"] = W8;
                dr["借券賣出過高"] = W9;
                dr["當沖比過高"] = W10;
                dt.Rows.Add(dr);
            }
        }



        //原本的LoadData 是直接計算邏輯，現在盤中會有排程計算邏輯然後轉進DB，從DB撈資料即可
        private void LoadData2()
        {

            string sql = $@"SELECT  [TDate],[UID],[UName],[W1],[W2],[W3],[W4],[W5],[W6],[W7],[W8],[W9],[W10],[WCnt]
  FROM [WarrantAssistant].[dbo].[盤中注意股票警示]　where [TDate] >　CONVERT(VARCHAR,GETDATE(),112)";

            DataTable dtWatch = MSSQL.ExecSqlQry(sql, GlobalVar.loginSet.warrantassistant45);

            foreach (DataRow drW in dtWatch.Rows)
            {
                string W1 = "0";
                
                string W2 = "0";
                string W3 = "0";
                string W4 = "0";
                string W5 = "0";
                string W6 = "0";
                string W7 = "0";
                string W8 = "0";
                string W9 = "0";
                string W10 = "0";
                DataRow dr = dt.NewRow();

                string uid = drW["UID"].ToString();
                string uname = drW["UName"].ToString();

                W1 = drW["W1"].ToString();
                W2 = drW["W2"].ToString();
                W3 = drW["W3"].ToString();
                W4 = drW["W4"].ToString();
                W5 = drW["W5"].ToString();
                W6 = drW["W6"].ToString();
                W7 = drW["W7"].ToString();
                W8 = drW["W8"].ToString();
                W9 = drW["W9"].ToString();
                W10 = drW["W10"].ToString();

                dr["標的代號"] = uid;
                dr["標的名稱"] = uname;
                dr["累積漲跌異常"] = W1;
                dr["起訖日漲跌異常"] = W2;
                dr["成交量過高"] = W3;
                dr["周轉率過高"] = W4;
                dr["券資比過高"] = W5;
                dr["平均成交量過高"] = W6;
                dr["累積周轉率過高"] = W7;
                dr["股價價差異常"] = W8;
                dr["借券賣出過高"] = W9;
                dr["當沖比過高"] = W10;
                dt.Rows.Add(dr);
            }
        }


    }
}

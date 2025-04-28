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
    public partial class FrmGeneralIssueV2Sub2 : Form
    {

        public string userID = GlobalVar.globalParameter.userID;
        private Dictionary<string, string> underlying_trader = new Dictionary<string, string>();
        DataTable dt;
        DataTable dt2;
        DataTable Data;
        DataTable Data_unlisted;
        Dictionary<string, double> SuggectVol_C = new Dictionary<string, double>();
        Dictionary<string, double> SuggectVol_P = new Dictionary<string, double>();
        Dictionary<string, double> HV = new Dictionary<string, double>();
        List<int> DeletedSerialNum = new List<int>();
        DataTable Underlying_S;

        public FrmGeneralIssueV2Sub2()
        {
            InitializeComponent();
        }

        private void FrmGeneralIssueV2Sub2_Load(object sender, EventArgs e)
        {

           
            SuggectVol_C.Clear();
            SuggectVol_P.Clear();
            HV.Clear();
            
            string sql_suggestVol_call = $@"SELECT  
                                              [UID]
	                                          ,[AllCounts]
                                              ,[RecommendVol]
	                                          ,[HV60]
                                          FROM [TwData].[dbo].[RecommendVolAll]
                                          WHERE [DateTime] = '{DateTime.Today.ToString("yyyyMMdd")}' AND [CP] = 'C' ";

            DataTable dt_suggectVol_call = MSSQL.ExecSqlQry(sql_suggestVol_call, GlobalVar.loginSet.twData);
            
            foreach (DataRow dr in dt_suggectVol_call.Rows)
            {
                string uid = dr["UID"].ToString();
                double allcount = Convert.ToDouble(dr["AllCounts"].ToString());
                double hv = Convert.ToDouble(dr["HV60"].ToString());
                double suggestVol = Convert.ToDouble(dr["RecommendVol"].ToString());
                if (allcount > 0)
                {
                    if (!SuggectVol_C.ContainsKey(uid))
                        SuggectVol_C.Add(uid, suggestVol);
                }
                if (!HV.ContainsKey(uid))
                    HV.Add(uid, hv);
            }

            string sql_suggestVol_put = $@"SELECT  
                                              [UID]
	                                          ,[AllCounts]
                                              ,[RecommendVol]
	                                          ,[HV60]
                                          FROM [TwData].[dbo].[RecommendVolAll]
                                          WHERE [DateTime] = '{DateTime.Today.ToString("yyyyMMdd")}' AND [CP] = 'P' ";
            DataTable dt_suggectVol_put = MSSQL.ExecSqlQry(sql_suggestVol_put, GlobalVar.loginSet.twData);
           
            foreach (DataRow dr in dt_suggectVol_put.Rows)
            {
                string uid = dr["UID"].ToString();
                double allcount = Convert.ToDouble(dr["AllCounts"].ToString());
                double suggestVol = Convert.ToDouble(dr["RecommendVol"].ToString());
                if (allcount > 0)
                {
                    if (!SuggectVol_P.ContainsKey(uid))
                        SuggectVol_P.Add(uid, suggestVol);
                }
            }

            underlying_trader.Clear();
            string getUnderlyingTraderStr = $@"SELECT  [UID]
                                                      ,[TraderAccount]
                                                  FROM [TwData].[dbo].[Underlying_Trader]
                                                  WHERE LEN(UID) < 5 AND LEFT(UID, 2) <> '00'";
            DataTable underlyingTrader = MSSQL.ExecSqlQry(getUnderlyingTraderStr, GlobalVar.loginSet.twData);
            foreach (DataRow dr in underlyingTrader.Rows)
            {
                string uid = dr["UID"].ToString();
                string trader = dr["TraderAccount"].ToString();
                if (!underlying_trader.ContainsKey(uid))
                    underlying_trader.Add(uid,trader);
            }
            
            LoadData();
        }
        public void LoadData()
        {
            LoadDeletedSerial();
            dt = new DataTable();
            dt2 = new DataTable();
            Data = new DataTable();
            Data_unlisted = new DataTable();
           
            dt.Columns.Add("欄位名稱", typeof(string));
            dt.Columns.Add("值", typeof(string));

            dt2.Columns.Add("現有履約價", typeof(string));
            //欲發行權證履約價
            double k = 0;
            //欲發行權證行使比例
            double cr = 0;
            //欲發行權證建議vol
            double iv = 0;
            //欲發行權證期間(預設6個月)
            double t = 6;
            //20210906利率改版，個股從2.5改為1，指數(IX0001、IX0027)改為0
            //double r = 0.025;
            double r_Index = 0.0;
            double r = 0.01;
            double s = 0;
            double p = 0;
            double vega = 0;
            double delta = 0;
            double gamma = 0;
            double theta = 0;
            double jumpsize = 0;
            string cp = this.textBox1.Text.Substring(textBox1.Text.Length - 1, 1);
            string[] temp = this.textBox1.Text.Split('-');
            string uid = temp[0];
            
            string sql_underlying_s = $@"SELECT A.[UnderlyingName]
                                               ,A.[UnderlyingID]
	                                           ,IsNull(IsNull(B.MPrice, IsNull(B.BPrice,B.APrice)),0) MPrice
                                               ,A.TraderID TraderID
                                      FROM [WarrantAssistant].[dbo].[WarrantUnderlying] A
                                      LEFT JOIN [WarrantAssistant].[dbo].[WarrantPrices] B ON A.UnderlyingID=B.CommodityID";

            Underlying_S = MSSQL.ExecSqlQry(sql_underlying_s, GlobalVar.loginSet.warrantassistant45);
            DataRow[] Underlying_S_Select = Underlying_S.Select($@"UnderlyingID = '{uid}'");
            if(Underlying_S_Select.Length > 0)
            {
                s = Convert.ToDouble(Underlying_S_Select[0][2].ToString());
            }



            double price_low = Convert.ToDouble(this.textBox2.Text);
            double price_high = Convert.ToDouble(this.textBox3.Text);
            if (price_high > 3)
                price_high = 999;

            double moneyness_low = Convert.ToDouble(this.textBox4.Text) / 100;
            double moneyness_high = Convert.ToDouble(this.textBox5.Text) / 100;
            string sql = $@"SELECT  DISTINCT [StrikePrice]
                            ,((CASE WHEN [WClass] = 'c' THEN 1- (StrikePrice / UClosePrice) ELSE  (StrikePrice / UClosePrice) -1 END) * -1) AS Moneyness
                          FROM [TwData].[dbo].[V_WarrantTrading]
                          WHERE [TDate] = '{EDLib.TradeDate.LastNTradeDate(1).ToString("yyyyMMdd")}' AND [UID] = '{uid}' and [WClass] = '{cp}' and [TtoM] > 80
                          and ((CASE WHEN [WClass] = 'c' THEN 1- (StrikePrice / UClosePrice) ELSE  (StrikePrice / UClosePrice) -1 END) * -1)  < {moneyness_high} and ((CASE WHEN [WClass] = 'c' THEN 1- (StrikePrice / UClosePrice) ELSE  (StrikePrice / UClosePrice) -1 END) * -1) >= {moneyness_low}  and [WTheoPrice_IV] >= 0.6 and [IV] > [HV_60D] and [IssuerName] = '9200'
                          ORDER BY ((CASE WHEN [WClass] = 'c' THEN 1- (StrikePrice / UClosePrice) ELSE  (StrikePrice / UClosePrice) -1 END) * -1)";
            Data = MSSQL.ExecSqlQry(sql, GlobalVar.loginSet.twData);
            foreach(DataRow dr in Data.Rows)
            {
                string strK = dr["StrikePrice"].ToString();
                DataRow dr2 = dt2.NewRow();
                dr2[0] = strK;
                dt2.Rows.Add(dr2);
            }

            //要加入未掛牌的權證
            string cp2 = "";
            if (cp == "c")
                cp2 = "購";
            else
                cp2 = "售";

            string sql2 = $@"SELECT  
                              DISTINCT [最新履約價]
	                          ,(CASE WHEN [名稱] like '%購%' THEN 1- ([最新履約價] / [標的收盤價]) ELSE  ([最新履約價] / [標的收盤價]) -1 END) * -1 AS Moneyness
                              FROM [TwCMData].[dbo].[Warrant總表]
                              WHERE [日期] = '{EDLib.TradeDate.LastNTradeDate(1).ToString("yyyyMMdd")}' AND ([上市日期] = '19110101' OR [上市日期] > '{EDLib.TradeDate.LastNTradeDate(1).ToString("yyyyMMdd")}') AND ([名稱] like '%{cp2}%')
                              AND (CASE WHEN [名稱] like '%購%' THEN 1- ([最新履約價] / [標的收盤價]) ELSE  ([最新履約價] / [標的收盤價]) -1 END) * -1 < {moneyness_high} AND (CASE WHEN [名稱] like '%購%' THEN 1- ([最新履約價] / [標的收盤價]) ELSE  ([最新履約價] / [標的收盤價]) -1 END) * -1 >= {moneyness_low}
                              AND LEN([標的代號]) < 5 AND LEFT([標的代號],2) <> '00' AND [券商代號] = '9200' AND [標的代號] = '{uid}'
                              ORDER BY (CASE WHEN [名稱] like '%購%' THEN 1- ([最新履約價] / [標的收盤價]) ELSE  ([最新履約價] / [標的收盤價]) -1 END) * -1";

            Data_unlisted = MSSQL.ExecSqlQry(sql2, GlobalVar.loginSet.twCMData);
            foreach (DataRow dr in Data_unlisted.Rows)
            {
                string strK = dr["最新履約價"].ToString();
                DataRow dr2 = dt2.NewRow();
                dr2[0] = strK+"*";
                dt2.Rows.Add(dr2);
            }

            if (price_high == 1)
                p = 0.8;
            else if (price_high == 2)
                p = 1.5;
            else if (price_high == 3)
                p = 2.5;
            else
                p = 3.2;
            //實際算出的價格，再反推行使比例
            double warrantP = 0;
            if (uid == "IX0001" || uid == "IX0027")
            {
                if (cp == "c")
                {
                    k = Math.Round(s * (1 + moneyness_high), 1);
                    if (SuggectVol_C.ContainsKey(uid))
                        iv = SuggectVol_C[uid] / 100;
                    else
                    {
                        if (HV.ContainsKey(uid))
                            iv = HV[uid] / 100;
                    }
                    warrantP = EDLib.Pricing.Option.PlainVanilla.CallPrice(s, k, r_Index, iv, t / 12);
                    vega = EDLib.Pricing.Option.PlainVanilla.CallVega(s, k, r_Index, iv, t / 12);
                    delta = EDLib.Pricing.Option.PlainVanilla.CallDelta(s, k, r_Index, iv, t / 12);
                    gamma = EDLib.Pricing.Option.PlainVanilla.CallGamma(s, k, r_Index, iv, t / 12);
                    theta = EDLib.Pricing.Option.PlainVanilla.CallTheta(s, k, r_Index, iv, t / 12);


                }
                if (cp == "p")
                {
                    k = Math.Round(s * (1 - moneyness_high), 1);
                    if (SuggectVol_P.ContainsKey(uid))
                        iv = SuggectVol_P[uid] / 100;
                    else
                    {
                        if (HV.ContainsKey(uid))
                            iv = HV[uid] / 100;
                    }
                    warrantP = EDLib.Pricing.Option.PlainVanilla.PutPrice(s, k, r_Index, iv, t / 12);
                    vega = EDLib.Pricing.Option.PlainVanilla.PutVega(s, k, r_Index, iv, t / 12);
                    delta = EDLib.Pricing.Option.PlainVanilla.PutDelta(s, k, r_Index, iv, t / 12);
                    gamma = EDLib.Pricing.Option.PlainVanilla.PutGamma(s, k, r_Index, iv, t / 12);
                    theta = EDLib.Pricing.Option.PlainVanilla.PutTheta(s, k, r_Index, iv, t / 12);
                }
            }
            else
            {
                if (cp == "c")
                {
                    k = Math.Round(s * (1 + moneyness_high), 1);
                    if (SuggectVol_C.ContainsKey(uid))
                        iv = SuggectVol_C[uid] / 100;
                    else
                    {
                        if (HV.ContainsKey(uid))
                            iv = HV[uid] / 100;
                    }
                    warrantP = EDLib.Pricing.Option.PlainVanilla.CallPrice(s, k, r, iv, t / 12);
                    vega = EDLib.Pricing.Option.PlainVanilla.CallVega(s, k, r, iv, t / 12);
                    delta = EDLib.Pricing.Option.PlainVanilla.CallDelta(s, k, r, iv, t / 12);
                    gamma = EDLib.Pricing.Option.PlainVanilla.CallGamma(s, k, r, iv, t / 12);
                    theta = EDLib.Pricing.Option.PlainVanilla.CallTheta(s, k, r, iv, t / 12);


                }
                if (cp == "p")
                {
                    k = Math.Round(s * (1 - moneyness_high), 1);
                    if (SuggectVol_P.ContainsKey(uid))
                        iv = SuggectVol_P[uid] / 100;
                    else
                    {
                        if (HV.ContainsKey(uid))
                            iv = HV[uid] / 100;
                    }
                    warrantP = EDLib.Pricing.Option.PlainVanilla.PutPrice(s, k, r, iv, t / 12);
                    vega = EDLib.Pricing.Option.PlainVanilla.PutVega(s, k, r, iv, t / 12);
                    delta = EDLib.Pricing.Option.PlainVanilla.PutDelta(s, k, r, iv, t / 12);
                    gamma = EDLib.Pricing.Option.PlainVanilla.PutGamma(s, k, r, iv, t / 12);
                    theta = EDLib.Pricing.Option.PlainVanilla.PutTheta(s, k, r, iv, t / 12);
                }
            }

            double p2 = 0;
            if (warrantP > 0)
            {
                cr = Math.Floor(p * 1000 / warrantP) / 1000;
            }
            if(cr > 0)
            {
                if (uid == "IX0001" || uid == "IX0027")
                {
                    if (cp == "c")
                        p2 = Math.Round(EDLib.Pricing.Option.PlainVanilla.CallPrice(s, k, r_Index, iv, t / 12) * cr, 2);
                    if (cp == "p")
                        p2 = Math.Round(EDLib.Pricing.Option.PlainVanilla.PutPrice(s, k, r_Index, iv, t / 12) * cr, 2);
                }
                else
                {
                    if (cp == "c")
                        p2 = Math.Round(EDLib.Pricing.Option.PlainVanilla.CallPrice(s, k, r, iv, t / 12) * cr, 2);
                    if (cp == "p")
                        p2 = Math.Round(EDLib.Pricing.Option.PlainVanilla.PutPrice(s, k, r, iv, t / 12) * cr, 2);
                }
                vega = Math.Round(vega * cr, 4);
                theta = Math.Round(theta * cr / 252.0, 4);
                gamma = Math.Round(gamma * cr * (s) * (s) / 100.0, 4);
                jumpsize = Math.Round(delta * cr * EDLib.Tick.UpTickSize(uid,s),4);
            }
            

            DataRow dr_k = dt.NewRow();
            dr_k[0] = "履約價";
            dr_k[1] = k.ToString();
            dt.Rows.Add(dr_k);

            DataRow dr_cr = dt.NewRow();
            dr_cr[0] = "行使比例";
            dr_cr[1] = cr.ToString();
            dt.Rows.Add(dr_cr);

            DataRow dr_iv = dt.NewRow();
            dr_iv[0] = "IV";
            dr_iv[1] = (iv * 100).ToString();
            dt.Rows.Add(dr_iv);

            DataRow dr_t = dt.NewRow();
            dr_t[0] = "期間(月)";
            dr_t[1] = t.ToString();
            dt.Rows.Add(dr_t);

            DataRow dr_cp = dt.NewRow();
            dr_cp[0] = "CP";
            dr_cp[1] = cp;
            dt.Rows.Add(dr_cp);

            DataRow dr_s = dt.NewRow();
            dr_s[0] = "股價";
            dr_s[1] = s.ToString();
            dt.Rows.Add(dr_s);

            DataRow dr_p = dt.NewRow();
            dr_p[0] = "價格(應發/實際)";
            dr_p[1] = p.ToString() + "/" + p2.ToString();
            dt.Rows.Add(dr_p);

            DataRow dr_vega = dt.NewRow();
            dr_vega[0] = "Vega";
            dr_vega[1] = vega.ToString();
            dt.Rows.Add(dr_vega);

            DataRow dr_jumpsize = dt.NewRow();
            dr_jumpsize[0] = "跳動價差";
            dr_jumpsize[1] = jumpsize.ToString();
            dt.Rows.Add(dr_jumpsize);

            DataRow dr_theta = dt.NewRow();
            dr_theta[0] = "Theta";
            dr_theta[1] = theta.ToString();
            dt.Rows.Add(dr_theta);

            DataRow dr_gamma = dt.NewRow();
            dr_gamma[0] = "Gamma";
            dr_gamma[1] = gamma.ToString();
            dt.Rows.Add(dr_gamma);

            this.ultraGrid1.DataSource = dt;
            this.ultraGrid2.DataSource = dt2;
            InitialGrid();
        }

        private void InitialGrid()
        {
            this.ultraGrid1.DisplayLayout.Override.RowSelectors = DefaultableBoolean.False;
            this.ultraGrid1.DisplayLayout.AutoFitStyle = AutoFitStyle.ResizeAllColumns;
            this.ultraGrid1.DisplayLayout.Override.RowSizing = RowSizing.AutoFree;

            this.ultraGrid2.DisplayLayout.Override.RowSelectors = DefaultableBoolean.False;
            this.ultraGrid2.DisplayLayout.AutoFitStyle = AutoFitStyle.ResizeAllColumns;
            

            foreach (UltraGridRow dr in ultraGrid1.Rows)
            {
                dr.Cells[0].Appearance.BackColor = Color.Orange;
                dr.Cells[0].Activation = Activation.NoEdit;
            }
            foreach (UltraGridRow dr in ultraGrid1.Rows)
            {
                if(dr.Cells[0].Text == "履約價" || dr.Cells[0].Text == "行使比例" || dr.Cells[0].Text == "IV"|| dr.Cells[0].Text == "期間(月)"|| dr.Cells[0].Text == "CP")
                    dr.Cells[1].Appearance.BackColor = Color.Yellow;
                else
                {
                    dr.Cells[1].Appearance.BackColor = Color.Gray;
                    dr.Cells[1].Activation = Activation.NoEdit;
                }
            }
        }


        private void Calculate()
        {
            //欲發行權證履約價
            double k = Convert.ToDouble(dt.Rows[0][1].ToString());
            //欲發行權證行使比例
            double cr = Convert.ToDouble(dt.Rows[1][1].ToString());
            //欲發行權證建議vol
            double iv = Convert.ToDouble(dt.Rows[2][1].ToString());
            //欲發行權證期間(預設6個月)
            double t = Convert.ToDouble(dt.Rows[3][1].ToString());
            //20210906利率改版，個股從2.5改成1，指數改成0
            //double r = 0.025;
            double r_Index = 0.0; 
            double r = 0.01;
            double s = 0;
            double p = 0;
            double vega = 0;
            double delta = 0;
            double jumpsize = 0;
            string cp = this.textBox1.Text.Substring(textBox1.Text.Length - 1, 1);
            string[] temp = this.textBox1.Text.Split('-');
            string uid = temp[0];

            string sql_underlying_s = $@"SELECT A.[UnderlyingName]
                                               ,A.[UnderlyingID]
	                                           ,IsNull(IsNull(B.MPrice, IsNull(B.BPrice,B.APrice)),0) MPrice
                                               ,A.TraderID TraderID
                                      FROM [WarrantAssistant].[dbo].[WarrantUnderlying] A
                                      LEFT JOIN [WarrantAssistant].[dbo].[WarrantPrices] B ON A.UnderlyingID=B.CommodityID";

            Underlying_S = MSSQL.ExecSqlQry(sql_underlying_s, GlobalVar.loginSet.warrantassistant45);
            DataRow[] Underlying_S_Select = Underlying_S.Select($@"UnderlyingID = '{uid}'");
            if (Underlying_S_Select.Length > 0)
            {
                s = Convert.ToDouble(Underlying_S_Select[0][2].ToString());
            }



            double price_low = Convert.ToDouble(this.textBox2.Text);
            double price_high = Convert.ToDouble(this.textBox3.Text);
            if (price_high > 3)
                price_high = 999;

            double moneyness_low = Convert.ToDouble(this.textBox4.Text) / 100;
            double moneyness_high = Convert.ToDouble(this.textBox4.Text) / 100;
            if (price_high == 1)
                p = 0.8;
            else if (price_high == 2)
                p = 1.5;
            else if (price_high == 3)
                p = 2.5;
            else
                p = 3.2;
            //實際算出的價格，再反推行使比例
            double warrantP = 0;

            if (uid == "IX0001" || uid == "IX0027")
            {
                if (cp == "c")
                {
                    warrantP = EDLib.Pricing.Option.PlainVanilla.CallPrice(s, k, r_Index, iv / 100, t / 12);
                    vega = EDLib.Pricing.Option.PlainVanilla.CallVega(s, k, r_Index, iv / 100, t / 12);
                    delta = EDLib.Pricing.Option.PlainVanilla.CallDelta(s, k, r_Index, iv / 100, t / 12);

                }
                if (cp == "p")
                {
                    warrantP = EDLib.Pricing.Option.PlainVanilla.PutPrice(s, k, r_Index, iv / 100, t / 12);
                    vega = EDLib.Pricing.Option.PlainVanilla.PutVega(s, k, r_Index, iv / 100, t / 12);
                    delta = EDLib.Pricing.Option.PlainVanilla.PutDelta(s, k, r_Index, iv / 100, t / 12);
                }
            }
            else
            {
                if (cp == "c")
                {
                    warrantP = EDLib.Pricing.Option.PlainVanilla.CallPrice(s, k, r, iv / 100, t / 12);
                    vega = EDLib.Pricing.Option.PlainVanilla.CallVega(s, k, r, iv / 100, t / 12);
                    delta = EDLib.Pricing.Option.PlainVanilla.CallDelta(s, k, r, iv / 100, t / 12);

                }
                if (cp == "p")
                {
                    warrantP = EDLib.Pricing.Option.PlainVanilla.PutPrice(s, k, r, iv / 100, t / 12);
                    vega = EDLib.Pricing.Option.PlainVanilla.PutVega(s, k, r, iv / 100, t / 12);
                    delta = EDLib.Pricing.Option.PlainVanilla.PutDelta(s, k, r, iv / 100, t / 12);
                }
            }

            double p2 = 0;
           
            if (cr > 0)
            {
                p2 = Math.Round(warrantP * cr, 2);
                vega = Math.Round(vega * cr / 100, 4);
                jumpsize = Math.Round(delta * cr * EDLib.Tick.UpTickSize(uid, s), 4);
            }


           

            dt.Rows[5][1] = s.ToString();
            dt.Rows[6][1] = p2.ToString();
            dt.Rows[7][1] = vega.ToString();
            dt.Rows[8][1] = jumpsize.ToString();
     
            InitialGrid();
        }

        private void button_EDIT_Click(object sender, EventArgs e)
        {
            Calculate();
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

        private void Issue()
        {
            LoadDeletedSerial();

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
                    if (dtTempList.Rows[0][0].ToString() != "")
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
                    if (dtOfficialList.Rows[0][0].ToString() != "")
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
                //foreach (Infragistics.Win.UltraWinGrid.UltraGridRow dr in ultraGrid5.Rows)
                {
                    string cp = this.textBox1.Text.Substring(textBox1.Text.Length - 1, 1);
                    string[] temp = this.textBox1.Text.Split('-');
                    string underlyingID = temp[0];
                    string[] underlying = temp[1].Split(':');
                    string underlyingName = underlying[0];
                    
                    string traderID = "";
                    try
                    {
                        traderID = underlying_trader[underlyingID].PadLeft(7, '0');
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show($@"{underlyingID}沒有對照交易員");
                        return;
                    }
                    string serialNumber = DateTime.Today.ToString("yyyyMMdd") + userID + "01" + (max + i).ToString("0#");
                    
                    double k = Convert.ToDouble(dt.Rows[0][1].ToString());
                    //欲發行權證行使比例
                    double cr = Convert.ToDouble(dt.Rows[1][1].ToString());
                    //欲發行權證建議vol
                    double iv = Convert.ToDouble(dt.Rows[2][1].ToString());
                    //欲發行權證期間(預設6個月)
                    int t = Convert.ToInt32(dt.Rows[3][1].ToString());

                    double hv = Convert.ToDouble(dt.Rows[2][1].ToString());
                    double issueNum = 5000;
                    double resetR = 0;
                    double barrierR = 0;
                    double financialR = 0;
                    double adj = 0;
                    string type = "一般型";
                    cp = cp == "c" ? "C" : "P";
                    double underlyingPrice = 0;
                    string sql_underlying_s = $@"SELECT A.[UnderlyingName]
                                               ,A.[UnderlyingID]
	                                           ,IsNull(IsNull(B.MPrice, IsNull(B.BPrice,B.APrice)),0) MPrice
                                               ,A.TraderID TraderID
                                      FROM [WarrantAssistant].[dbo].[WarrantUnderlying] A
                                      LEFT JOIN [WarrantAssistant].[dbo].[WarrantPrices] B ON A.UnderlyingID=B.CommodityID";

                    Underlying_S = MSSQL.ExecSqlQry(sql_underlying_s, GlobalVar.loginSet.warrantassistant45);
                    DataRow[] Underlying_S_Select = Underlying_S.Select($@"UnderlyingID = '{underlyingID}'");
                    if (Underlying_S_Select.Length > 0)
                    {
                        underlyingPrice = Convert.ToDouble(Underlying_S_Select[0][2].ToString());
                    }
                  
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
            catch (Exception ex)
            {
                MessageBox.Show($@"{ex.Message}");
            }
            h.Dispose();
            MessageBox.Show("發行完成!");
        }

        private void button_CONFIRM_Click(object sender, EventArgs e)
        {
            Issue();
        }
    }
}

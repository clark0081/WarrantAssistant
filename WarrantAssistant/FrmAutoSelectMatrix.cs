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
using System.Data;
using System.Data.SqlClient;
using EDLib.SQL;
using Infragistics.Win.UltraWinGrid;
using System.IO;

namespace WarrantAssistant
{
    public partial class FrmAutoSelectMatrix : Form
    {

        #region 計算履約價間距近似價格
        public static double roundpriceK(double s,double k)
        {
            if (s < 100)
            {
                return Math.Round(k, 1);
            }
            else if (s < 500)
            {
                return Math.Round(k);
            }
            else//s>=500  ex 503 502
            {
                double tick = 5;//ex 26
                int up = (int)Math.Floor((double)(k / 10)) * 10 + 10;//30
                int down = (int)Math.Floor((double)(k / 10)) * 10;//20
                if ((double)k % 10 > tick)
                {
                    if ((double)(up - k) > (double)(k - down - tick))
                        return (double)(down + tick);
                    else
                        return (double)up;
                }
                else
                {
                    if ((double)(down + tick - k) > (double)(k - down))
                        return (double)(down);
                    else
                        return (double)(down + tick);
                }
            }
        }
        #endregion
        DateTime today = DateTime.Today;
        DateTime lastday = EDLib.TradeDate.LastNTradeDate(1);
        public string userID = GlobalVar.globalParameter.userID;
        private System.Data.DataTable dataTable1 = new System.Data.DataTable();
        Dictionary<string,string> IssuersNum2Char = new Dictionary<string, string>();//券商代號轉券商名稱
        Dictionary<string, string> IssuersChar2Num = new Dictionary<string, string>();//券商名稱轉券商代號

        private bool isEdit = false;
        public FrmAutoSelectMatrix()
        {
            InitializeComponent();
        }

        private void FrmAutoSelectMatrix_Load(object sender, EventArgs e)
        {

            InitialGrid();
            this.comboBox1.Items.Add("C");
            this.comboBox1.Items.Add("P");
            this.comboBox1.SelectedItem = "C";
            LoadData();
            LoadIssuer();
            foreach (var item in IssuersChar2Num.Keys)
                comboBox2.Items.Add(item);
            comboBox2.Items.Add("All");
            comboBox2.Text = "All";
            
        }

        private void InitialGrid()
        {
            dataTable1.Columns.Add("標的代號", typeof(string));
            dataTable1.Columns.Add("標的名稱", typeof(string));
            dataTable1.Columns.Add("CP", typeof(string));
            dataTable1.Columns.Add("確認", typeof(bool));
            dataTable1.Columns.Add("履約價中點", typeof(double));
            dataTable1.Columns.Add("履約價間距(%)", typeof(double));
            dataTable1.Columns.Add("價內監控區間", typeof(int));
            dataTable1.Columns.Add("價外監控區間", typeof(int));
            dataTable1.Columns.Add("距到期日起始月", typeof(double));
            dataTable1.Columns.Add("到期日間距(月)", typeof(double));
            dataTable1.Columns.Add("期間監控區", typeof(int));

            ultraGrid1.DataSource = dataTable1;
            this.ultraGrid1.DisplayLayout.AutoFitStyle = Infragistics.Win.UltraWinGrid.AutoFitStyle.ResizeAllColumns;
            SetButton();
            
        }
        private void LoadIssuer()
        {
            string lastDay = TradeDate.LastNTradeDate(1);
            string sql = $@"SELECT  DISTINCT [IssuerName] AS IssuerCode,SUBSTRING([WName],Len([WName])-6,2) AS IssuerName
                            FROM [TwData].[dbo].[V_WarrantTrading]
                            WHERE [TDate] = '{lastday}' AND SUBSTRING([WName],Len([WName])-6,2) <> ''";
            System.Data.DataTable dt = MSSQL.ExecSqlQry(sql, GlobalVar.loginSet.twData);
            foreach(DataRow dr in dt.Rows)
            {
                string issuername = dr["IssuerName"].ToString();
                string issuer = dr["IssuerCode"].ToString();
                if (issuername != "")
                {
                    if (!IssuersNum2Char.ContainsKey(issuer))
                        IssuersNum2Char.Add(issuer, issuername);
                    if (!IssuersChar2Num.ContainsKey(issuername))
                        IssuersChar2Num.Add(issuername, issuer);
                }
            }
        }
        private void LoadData()
        {

            string getprice = $@"SELECT DISTINCT CASE WHEN ([CommodityId]='1000') THEN 'IX0001' ELSE [CommodityId] END AS CommodityID
                                             ,ISNULL([LastPrice],0) AS LastPrice
                                             ,[tradedate] AS TradeDate
                                             ,isnull([BuyPriceBest1],0) AS Bid1
                                             ,isnull([SellPriceBest1],0) AS Ask1
                               FROM [TsQuote].[dbo].[vwprice2]
							   WHERE [kind] in ('ETF','Index','Stock');";
            DataTable vprice = EDLib.SQL.MSSQL.ExecSqlQry(getprice, GlobalVar.loginSet.tsquoteSqlConnString);
            dataTable1.Clear();
            string CP = comboBox1.Text;
#if !To39
            string sql = $@"SELECT  [UID], [UName], [CP],[Interval], [StartMonth]
                        , [IntervalMonth], [Interval_InTheMoney], [Interval_OutOfMoney], [Num_IntervalMonth], [Checked]
                          FROM [newEDIS].[dbo].[OptionAutoSelectMatrix]
                          WHERE [TraderID] ='{userID.TrimStart('0')}' and [CP] ='{CP}'
                          ORDER BY [UID]";
            System.Data.DataTable dt = EDLib.SQL.MSSQL.ExecSqlQry(sql, GlobalVar.loginSet.newEDIS);
#else
            string sql = $@"SELECT  [UID], [UName], [CP],[Interval], [StartMonth]
                        , [IntervalMonth], [Interval_InTheMoney], [Interval_OutOfMoney], [Num_IntervalMonth], [Checked]
                          FROM [WarrantAssistant].[dbo].[OptionAutoSelectMatrix]
                          WHERE [TraderID] ='{userID.TrimStart('0')}' and [CP] ='{CP}'
                          ORDER BY [UID]";
            System.Data.DataTable dt = EDLib.SQL.MSSQL.ExecSqlQry(sql, GlobalVar.loginSet.warrantassistant45);
#endif
            foreach (DataRow dr in dt.Rows)
            {
                string uid = dr["UID"].ToString();
                DataRow[] select = vprice.Select($@"CommodityID='{uid}'");
                if (select.Length == 0)
                    continue;
                //取得履約價中點
                double get_quote;
                if (Convert.ToDouble(select[0][1].ToString()) == 0)
                {
                    if (Convert.ToDouble(select[0][3].ToString()) == 0)
                    {
                        if (Convert.ToDouble(select[0][4].ToString()) == 0)
                            get_quote = 0;
                        else
                            get_quote = roundprice(Convert.ToDouble(select[0][4].ToString()));
                    }
                    else
                    {
                        get_quote = roundprice(Convert.ToDouble(select[0][3].ToString()));
                    }

                }
                else
                {
                    get_quote = roundprice(Convert.ToDouble(select[0][1].ToString()));
                }
                DataRow drv = dataTable1.NewRow();
                
                
               
                drv["標的代號"] = uid;

                string uname = dr["UName"].ToString();
                drv["標的名稱"] = uname;
                int check = Convert.ToInt32(dr["Checked"].ToString());
                drv["確認"] = check;
                string cp = dr["CP"].ToString();
                drv["CP"] = cp;
                //履約價中點
                drv["履約價中點"] = get_quote;
                //履約價間隔
                double interval = Convert.ToDouble(dr["Interval"].ToString());
                drv["履約價間距(%)"] = interval;
                //價內監控區間
                int num_itm = Convert.ToInt32(dr["Interval_InTheMoney"].ToString());
                drv["價內監控區間"] = num_itm;
                //價外監控區間
                int num_otm = Convert.ToInt32(dr["Interval_OutOfMoney"].ToString());
                drv["價外監控區間"] = num_otm;
                //距到期日起始月
                double maturity_start = Convert.ToDouble(dr["StartMonth"].ToString());
                drv["距到期日起始月"] = maturity_start;
                //到期日間距
                double maturity_interval = Convert.ToDouble(dr["IntervalMonth"].ToString());
                drv["到期日間距(月)"] = maturity_interval;
                //期間監控區
                int num_maturity = Convert.ToInt32(dr["Num_IntervalMonth"].ToString());
                drv["期間監控區"] = num_maturity;

                dataTable1.Rows.Add(drv);
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


                bands0.Columns["標的代號"].CellActivation = Activation.AllowEdit;
                bands0.Columns["標的名稱"].CellActivation = Activation.AllowEdit;
                bands0.Columns["CP"].CellActivation = Activation.AllowEdit;
                bands0.Columns["確認"].CellActivation = Activation.AllowEdit;
                bands0.Columns["履約價中點"].CellActivation = Activation.AllowEdit;
                bands0.Columns["履約價間距(%)"].CellActivation = Activation.AllowEdit;
                bands0.Columns["價內監控區間"].CellActivation = Activation.AllowEdit;
                bands0.Columns["價外監控區間"].CellActivation = Activation.AllowEdit;
                bands0.Columns["距到期日起始月"].CellActivation = Activation.AllowEdit;
                bands0.Columns["到期日間距(月)"].CellActivation = Activation.AllowEdit;
                bands0.Columns["期間監控區"].CellActivation = Activation.AllowEdit;
              
                ultraGrid1.DisplayLayout.Override.CellAppearance.BackColor = Color.White;
                button3.Visible = false;
                button4.Visible = true;
                button5.Visible = true;
            }
            else
            {
                bands0.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
                bands0.Columns["標的代號"].CellActivation = Activation.NoEdit;
                bands0.Columns["標的名稱"].CellActivation = Activation.NoEdit;
                bands0.Columns["CP"].CellActivation = Activation.NoEdit;
                bands0.Columns["確認"].CellActivation = Activation.NoEdit;
                bands0.Columns["履約價中點"].CellActivation = Activation.NoEdit;
                bands0.Columns["履約價間距(%)"].CellActivation = Activation.NoEdit;
                bands0.Columns["價內監控區間"].CellActivation = Activation.NoEdit;
                bands0.Columns["價外監控區間"].CellActivation = Activation.NoEdit;
                bands0.Columns["距到期日起始月"].CellActivation = Activation.NoEdit;
                bands0.Columns["到期日間距(月)"].CellActivation = Activation.NoEdit;
                bands0.Columns["期間監控區"].CellActivation = Activation.NoEdit;
                ultraGrid1.DisplayLayout.Override.CellAppearance.BackColor = Color.Moccasin;
                button3.Visible = true;
                button4.Visible = false;
                button5.Visible = false;
            }
        }
        private void ComboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (!isEdit)
            {
                LoadData();
                
            }
            else
            {
                isEdit = false;
                LoadData();
                SetButton();
            }
        }
        private void ComboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
            string uid = textBox2.Text;
            if(uid!="")
                ShowMatrix_NonKGI();
        }
        private void button3_EditClick(object sender, EventArgs e)
        {
            isEdit = true;
            SetButton();
        }

        private void button5_CancelClick(object sender, EventArgs e)
        {
            isEdit = false;
            SetButton();
            LoadData();
        }

        private void button4_ConfirmClick(object sender, EventArgs e)
        {
            if (!CheckData())//先確認資料是否有效
                return;
            string CP = comboBox1.Text;
#if !To39
            string sqldel = $@"DELETE [newEDIS].[dbo].[OptionAutoSelectMatrix]
                                WHERE [TraderID] ='{userID.TrimStart('0')}' AND [CP] ='{CP}'";
            MSSQL.ExecSqlCmd(sqldel, GlobalVar.loginSet.newEDIS);
            string sql = $@"INSERT INTO [newEDIS].[dbo].[OptionAutoSelectMatrix] ([TraderID], [UID], [UName], [CP]
                , [Interval], [StartMonth], [IntervalMonth], [Interval_InTheMoney], [Interval_OutOfMoney], [Num_IntervalMonth], [Checked])
                VALUES (@TraderID, @UID,@UName,@CP,@Interval,@StartMonth,@IntervalMonth,@Interval_InTheMoney,@Interval_OutOfMoney,@Num_IntervalMonth,@Checked )";

            List<SqlParameter> ps = new List<SqlParameter> {
                    new SqlParameter("@TraderID", SqlDbType.VarChar),
                    new SqlParameter("@UID", SqlDbType.VarChar),
                    new SqlParameter("@UName", SqlDbType.VarChar),
                    new SqlParameter("@CP", SqlDbType.VarChar),
                    new SqlParameter("@Interval", SqlDbType.Float),
                    new SqlParameter("@StartMonth", SqlDbType.Float),
                    new SqlParameter("@IntervalMonth", SqlDbType.Float),
                    new SqlParameter("@Interval_InTheMoney", SqlDbType.Float),
                    new SqlParameter("@Interval_OutOfMoney", SqlDbType.Float),
                    new SqlParameter("@Num_IntervalMonth", SqlDbType.Float),
                    new SqlParameter("@Checked", SqlDbType.Int)
                };
            SQLCommandHelper h = new SQLCommandHelper(GlobalVar.loginSet.newEDIS, sql, ps);
#else
            string sqldel = $@"DELETE [WarrantAssistant].[dbo].[OptionAutoSelectMatrix]
                                WHERE [TraderID] ='{userID.TrimStart('0')}' AND [CP] ='{CP}'";
            MSSQL.ExecSqlCmd(sqldel, GlobalVar.loginSet.warrantassistant45);
            string sql = $@"INSERT INTO [WarrantAssistant].[dbo].[OptionAutoSelectMatrix] ([TraderID], [UID], [UName], [CP]
                , [Interval], [StartMonth], [IntervalMonth], [Interval_InTheMoney], [Interval_OutOfMoney], [Num_IntervalMonth], [Checked])
                VALUES (@TraderID, @UID,@UName,@CP,@Interval,@StartMonth,@IntervalMonth,@Interval_InTheMoney,@Interval_OutOfMoney,@Num_IntervalMonth,@Checked )";

            List<SqlParameter> ps = new List<SqlParameter> {
                    new SqlParameter("@TraderID", SqlDbType.VarChar),
                    new SqlParameter("@UID", SqlDbType.VarChar),
                    new SqlParameter("@UName", SqlDbType.VarChar),
                    new SqlParameter("@CP", SqlDbType.VarChar),
                    new SqlParameter("@Interval", SqlDbType.Float),
                    new SqlParameter("@StartMonth", SqlDbType.Float),
                    new SqlParameter("@IntervalMonth", SqlDbType.Float),
                    new SqlParameter("@Interval_InTheMoney", SqlDbType.Float),
                    new SqlParameter("@Interval_OutOfMoney", SqlDbType.Float),
                    new SqlParameter("@Num_IntervalMonth", SqlDbType.Float),
                    new SqlParameter("@Checked", SqlDbType.Int)
                };
            SQLCommandHelper h = new SQLCommandHelper(GlobalVar.loginSet.warrantassistant45, sql, ps);
#endif
            foreach (Infragistics.Win.UltraWinGrid.UltraGridRow r in ultraGrid1.Rows)
            {
                try
                {
                    string userid = userID.TrimStart('0');
                    h.SetParameterValue("@TraderID", userid);
                    string uid = r.Cells["標的代號"].Value.ToString();
                    h.SetParameterValue("@UID", uid);
                    string uname = r.Cells["標的名稱"].Value.ToString();
                    h.SetParameterValue("@UName", uname);
                    string cp = r.Cells["CP"].Value.ToString();
                    h.SetParameterValue("@CP", cp);

                    double interval = r.Cells["履約價間距(%)"].Value == DBNull.Value ? 0 : Convert.ToDouble(r.Cells["履約價間距(%)"].Value);
                    h.SetParameterValue("@Interval", interval);

                    double maturity_start = r.Cells["距到期日起始月"].Value == DBNull.Value ? 0 : Convert.ToDouble(r.Cells["距到期日起始月"].Value);
                    h.SetParameterValue("@StartMonth", maturity_start);
                    
                    double maturity_interval = r.Cells["到期日間距(月)"].Value == DBNull.Value ? 0 : Convert.ToDouble(r.Cells["到期日間距(月)"].Value);
                    h.SetParameterValue("@IntervalMonth", maturity_interval);

                    
                    double num_itm = r.Cells["價內監控區間"].Value == DBNull.Value ? 0 : Convert.ToDouble(r.Cells["價內監控區間"].Value);
                    h.SetParameterValue("@Interval_InTheMoney", num_itm);

                    double num_otm = r.Cells["價外監控區間"].Value == DBNull.Value ? 0 : Convert.ToDouble(r.Cells["價外監控區間"].Value);
                    h.SetParameterValue("@Interval_OutOfMoney", num_otm);

                    double num_maturity = r.Cells["期間監控區"].Value == DBNull.Value ? 0 : Convert.ToDouble(r.Cells["期間監控區"].Value);
                    h.SetParameterValue("@Num_IntervalMonth", num_maturity);
                    
                    int check = Convert.ToInt32(r.Cells["確認"].Value);
                    h.SetParameterValue("@Checked", check);
                    
                    h.ExecuteCommand();
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }
            isEdit = false;
            SetButton();
            LoadData();
            ReCalCallPutDensity();
            ReCalYuanCallPutDensity();
            ReCalKgiYuanCallPutDensity();
        }

#region 計算履約價中點
        public static double roundprice(double x)
        {
            if (x < 10)
            {
                double tick = 0.5;
                int up = (int)Math.Floor(x) + 1;
                int down = (int)Math.Floor(x);//8
                if ((double)x % 1 > tick)//ex 3.6
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
            else if (x < 50)
                return Math.Round(x);
            else if (x < 100)
            {
                double tick = 5;
                int up = (int)Math.Floor((double)(x / 10)) * 10 + 10;
                int down = (int)Math.Floor((double)(x / 10)) * 10;
                if ((double)x % 10 > tick)//ex 88
                {
                    if ((double)(up - x) > (double)(x - down - tick))
                        return (double)(down + tick);
                    else
                        return (double)up;
                }
                else//ex 84
                {
                    if ((double)(down + tick - x) > (double)(x - down))
                        return (double)(down);
                    else
                        return (double)(down + tick);
                }
            }
            else if (x < 500)
            {
                return Math.Round(x / 10) * 10;
            }
            else if (x < 1000)
            {
                double tick = 50;
                int up = (int)Math.Floor((double)(x / 100)) * 100 + 100;
                int down = (int)Math.Floor((double)(x / 100)) * 100;
                if ((double)x % 100 > tick)//ex 980
                {
                    if ((double)(up - x) > (double)(x - down - tick))
                        return (double)(down + tick);
                    else
                        return (double)up;
                }
                else//ex 84
                {
                    if ((double)(down + tick - x) > (double)(x - down))
                        return (double)(down);
                    else
                        return (double)(down + tick);
                }
            }
            else
                return Math.Round(x / 100) * 100;

        }

#endregion

        private void button1_ShowMatrixClick(object sender, EventArgs e)
        {
            //先更新即時價格
            LoadData();
            ShowMatrix();
            
        }
        private void ShowMatrix()
        {
            string todaystr = today.ToString("yyyyMMdd");
            string lastdaystr = lastday.ToString("yyyyMMdd");
            string uid = textBox1.Text;
            string CP = comboBox1.Text;
            string cp;
            if (CP == "C")
                cp = "c";
            else
                cp = "p";

            DataTable warrantbasic = EDLib.SQL.MSSQL.ExecSqlQry($@" SELECT * FROM (SELECT  [stkid] AS [UID]
                                              , CASE WHEN (SUBSTRING([type],4,4) ='認購權證') THEN 'c' ELSE 'p' END AS WClass
	                                          , [issuedate] AS IssuedDate
                                              , [strike_now] AS StrikePrice
                                              , [maturitydate] AS MaturityDate
	                                          , [TraderID] AS TraderAccount
                                          FROM [HEDGE].[dbo].[WARRANTS]
                                          WHERE [kgiwrt] = '自家' AND [maturitydate] >'{DateTime.Today.ToString("yyyyMMdd")}' AND ([type] LIKE '%認購權證%'  OR [type] LIKE '%認售權證%')) AS A
                                            WHERE A.[WClass] = '{cp}' AND A.[UID] = '{uid}'", "Data Source=10.101.10.5;Initial Catalog=HEDGE;User ID=hedgeuser;Password=hedgeuser");
            DataTable dtMatrix = new System.Data.DataTable();
        
            //取得履約價中點
            DataRow[] select = dataTable1.Select($@"標的代號='{uid}'");
            if (select.Length > 0)
            {
                int check = Convert.ToInt32(select[0][3]);
                if (check == 0)
                {
                    MessageBox.Show("標的未設定參數");
                    return;
                }
                double get_quote = Convert.ToDouble(select[0][4].ToString());
                //履約價間隔(%)
                double intervalPercent = Convert.ToDouble(select[0][5].ToString());
                if (intervalPercent <= 0)
                {
                    MessageBox.Show("未設定履約價間距");
                    return;
                }
                double interval = roundpriceK(get_quote, get_quote * intervalPercent / 100);
                //價內監控區間
                int num_itm = (cp == "c" ? Convert.ToInt32(select[0][6].ToString()) : Convert.ToInt32(select[0][7].ToString()));
                //價內監控區間若小於4格，設4格
                int min_itm = num_itm >= 4 ? num_itm : 4;
                //價外監控區間
                int num_otm = (cp == "c" ? Convert.ToInt32(select[0][7].ToString()) : Convert.ToInt32(select[0][6].ToString()));
                //價外監控區間若小於4格，設4格
                int max_otm = num_otm >= 4 ? num_otm : 4;
                //距到期日起始月
                double maturity_start = Convert.ToDouble(select[0][8].ToString());
                //到期日間距
                double maturity_interval = Convert.ToDouble(select[0][9].ToString());
                //期間監控區
                int num_maturity = Convert.ToInt32(select[0][10].ToString());
                //最大期間監控區，若小於5格，設5格
                int max_num_maturity = num_maturity >= 5 ? num_maturity : 5;
                if (warrantbasic.Rows.Count > 0)
                {
                    //建表
                    dtMatrix.Columns.Add("Maturity", typeof(string));
                    for (int i = min_itm; i > 0; i--)
                        dtMatrix.Columns.Add($@"{(get_quote - i * interval).ToString()}~{(get_quote - (i - 1) * interval).ToString()}", typeof(int));
                    for (int i = 1; i <= max_otm; i++)
                        dtMatrix.Columns.Add($@"{(get_quote + (i - 1) * interval).ToString()}~{(get_quote + i * interval).ToString()}", typeof(int));

                    for (int i = 1; i <= max_num_maturity; i++)
                    {
                        DataRow dr = dtMatrix.NewRow();
                        //用掛牌日算，才能有效把前一天的0補成1
                        DateTime ListedDat = EDLib.TradeDate.NextNTradeDate(2);
                        DateTime t1 = ListedDat.AddMonths((int)(maturity_start + maturity_interval * (i - 1))).AddDays(-1);
                        DateTime t2 = ListedDat.AddMonths((int)(maturity_start + maturity_interval * i)).AddDays(-1);
                        string t1str = t1.ToString("yyyy-MM-dd HH:mm:ss.fff");
                        string t2str = t2.ToString("yyyy-MM-dd HH:mm:ss.fff");

                        /*
                        string t1 = DateTime.Today.AddDays((int)Math.Round((maturity_start + maturity_interval * (i - 1)) * 30)).ToString("yyyy-MM-dd HH:mm:ss.fff");
                        DateTime t11 = DateTime.Today.AddDays((int)Math.Round((maturity_start + maturity_interval * (i - 1)) * 30));
                        string t2 = DateTime.Today.AddDays((int)Math.Round((maturity_start + maturity_interval * i) * 30)).ToString("yyyy-MM-dd HH:mm:ss.fff");
                        DateTime t22 = DateTime.Today.AddDays((int)Math.Round((maturity_start + maturity_interval * i) * 30));
                        dr[0] = (maturity_start + (i - 1) * maturity_interval).ToString() + "~" + (maturity_start + i * maturity_interval).ToString();
                        dr[0] = t1 + "~" + t2;
                        */

                        dr[0] = (maturity_start + (i - 1) * maturity_interval).ToString() + "~" + (maturity_start + i * maturity_interval).ToString();
                        //dr[0] = t1str + "~" + t2str;
                        int w = 1;
                        for (int j = min_itm; j > 0; j--)
                        {
                            //DataRow[] temp = warrantbasic.Select($@"[MaturityDate] > '{t11}' AND [MaturityDate] <= '{t22}' AND [StrikePrice] > {get_quote - j * interval} AND [StrikePrice] <={get_quote - (j - 1) * interval}");
                            DataRow[] temp = warrantbasic.Select($@"[MaturityDate] > '{t1}' AND [MaturityDate] <= '{t2}' AND [StrikePrice] > {get_quote - j * interval} AND [StrikePrice] <={get_quote - (j - 1) * interval}");

                            dr[w] = temp.Length;
                            w++;
                        }
                        for (int j = 1; j <= max_otm; j++)
                        {
                            //DataRow[] temp = warrantbasic.Select($@"[MaturityDate] > '{t11}' AND [MaturityDate] <= '{t22}' AND [StrikePrice] >{get_quote + (j - 1) * interval} AND [StrikePrice] <={get_quote + (j) * interval}");
                            DataRow[] temp = warrantbasic.Select($@"[MaturityDate] > '{t1}' AND [MaturityDate] <= '{t2}' AND [StrikePrice] >{get_quote + (j - 1) * interval} AND [StrikePrice] <={get_quote + (j) * interval}");
                            dr[w] = temp.Length;
                            w++;
                        }
                        dtMatrix.Rows.Add(dr);
                    }
                }
                else//如果warrantbasics為空，建一張全都是0的table
                {
                    //建表
                    dtMatrix.Columns.Add("Maturity", typeof(string));
               
                    for (int i = min_itm; i > 0; i--)
                        dtMatrix.Columns.Add($@"{(get_quote - i * interval).ToString()}~{(get_quote - (i - 1) * interval).ToString()}", typeof(int));
                    for (int i = 1; i <= max_otm; i++)
                        dtMatrix.Columns.Add($@"{(get_quote + (i - 1) * interval).ToString()}~{(get_quote + i * interval).ToString()}", typeof(int));
          

                    for (int i = 1; i <= max_num_maturity; i++)
                    {
                        DataRow dr = dtMatrix.NewRow();
                        DateTime ListedDat = EDLib.TradeDate.NextNTradeDate(2);
                        DateTime t1 = ListedDat.AddMonths((int)(maturity_start + maturity_interval * (i - 1))).AddDays(-1);
                        DateTime t2 = ListedDat.AddMonths((int)(maturity_start + maturity_interval * i)).AddDays(-1);
                        string t1str = t1.ToString("yyyy-MM-dd HH:mm:ss.fff");
                        string t2str = t2.ToString("yyyy-MM-dd HH:mm:ss.fff");
                        /*
                        string t1 = DateTime.Today.AddDays((int)Math.Round((maturity_start + maturity_interval * (i - 1)) * 30)).ToString("yyyy-MM-dd HH:mm:ss.fff");
                        DateTime t11 = DateTime.Today.AddDays((int)Math.Round((maturity_start + maturity_interval * (i - 1)) * 30));
                        string t2 = DateTime.Today.AddDays((int)Math.Round((maturity_start + maturity_interval * i) * 30)).ToString("yyyy-MM-dd HH:mm:ss.fff");
                        DateTime t22 = DateTime.Today.AddDays((int)Math.Round((maturity_start + maturity_interval * i) * 30));
                        dr[0] = t1 + "~" + t2;
                        */
                        
                        dr[0] = (maturity_start + (i - 1) * maturity_interval).ToString() + "~" + (maturity_start + i * maturity_interval).ToString();
                        //dr[0] = t1str + "~" + t2str;
                        int w = 1;
                        for (int j = min_itm; j > 0; j--)
                        {
                            dr[w] = 0;
                            w++;
                        }
                        for (int j = 1; j <= max_otm; j++)
                        {
                            dr[w] = 0;
                            w++;
                        }
                        dtMatrix.Rows.Add(dr);
                    }
                }
                ultraGrid2.DataSource = dtMatrix;
                SetMatrix(min_itm, max_otm, num_itm, num_otm, num_maturity);
            }
            else
            {
                MessageBox.Show("標的錯誤或無設定此標的");
            }
        }
        private void SetMatrix(int min_itm, int max_otm, int num_itm,int num_otm,int num_maturity)
        {
            UltraGridBand bands0 = ultraGrid2.DisplayLayout.Bands[0];
            bands0.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
            this.ultraGrid2.DisplayLayout.AutoFitStyle = Infragistics.Win.UltraWinGrid.AutoFitStyle.ResizeAllColumns;
            this.ultraGrid2.DisplayLayout.Override.DefaultRowHeight = 50;
            this.ultraGrid2.DisplayLayout.Override.CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center;
            this.ultraGrid2.DisplayLayout.Override.CellAppearance.TextVAlign = Infragistics.Win.VAlign.Middle;
            /*
            bands0.Columns["標的代號"].CellActivation = Activation.NoEdit;
            bands0.Columns["標的名稱"].CellActivation = Activation.NoEdit;
            bands0.Columns["CP"].CellActivation = Activation.NoEdit;
            bands0.Columns["確認"].CellActivation = Activation.NoEdit;
            bands0.Columns["履約價中點"].CellActivation = Activation.NoEdit;
            bands0.Columns["履約價間距"].CellActivation = Activation.NoEdit;
            bands0.Columns["價內監控區間"].CellActivation = Activation.NoEdit;
            bands0.Columns["價外監控區間"].CellActivation = Activation.NoEdit;
            bands0.Columns["距到期日起始月"].CellActivation = Activation.NoEdit;
            bands0.Columns["到期日間距(月)"].CellActivation = Activation.NoEdit;
            bands0.Columns["期間監控區"].CellActivation = Activation.NoEdit;
            */
            foreach (var col in bands0.Columns)
            {
                col.CellActivation = Activation.NoEdit;
            }
            //ultraGrid2.DisplayLayout.Override.CellAppearance.BackColor = Color.Moccasin;
            for(int i = 0; i < num_maturity; i++)
            {
                for(int j = (min_itm + 1) - num_itm; j < (min_itm + 1); j++)
                {
                    ultraGrid2.Rows[i].Cells[j].Appearance.BackColor = Color.Orange;
                }
                for (int j = (min_itm + 1); j < (min_itm + 1) + num_otm; j++)
                {
                    ultraGrid2.Rows[i].Cells[j].Appearance.BackColor = Color.SkyBlue;
                }

            }
        }

        //印出他家矩陣
        private void ShowMatrix_NonKGI()
        {

            string todaystr = today.ToString("yyyyMMdd");
            string lastdaystr = lastday.ToString("yyyyMMdd");
            string uid = textBox2.Text;
            string CP = comboBox1.Text;
            string cp;
            if (CP == "C")
                cp = "c";
            else
                cp = "p";
            
            DataTable warrantbasic;
            if (comboBox2.Text == "All")
            {
             
                warrantbasic = EDLib.SQL.MSSQL.ExecSqlQry($@"SELECT * FROM(SELECT CASE [UID] WHEN 'TWA00' THEN 'IX0001' ELSE [UID] END AS [UID]
      ,[WClass]
      ,[IssuedDate]
      ,CASE WHEN([MaturityDate]='19110101') THEN DATEADD(MONTH,[Duration(Month)],[ListedDate]) ELSE [MaturityDate] END AS MaturityDate
      ,[StrikePrice]
      ,[TraderAccount]
  FROM [newEDIS].[dbo].[WarrantBasics]
  WHERE ([MaturityDate] >'{DateTime.Today.ToString("yyyyMMdd")}' OR [MaturityDate] = '19110101') and [WClass] ='{cp}' and [IssuerName] != '9200') AS A
  WHERE A.[UID] = '{uid}'", GlobalVar.loginSet.newEDIS);
            }
            else
            {
                
                string issuer = IssuersChar2Num[comboBox2.Text];
                /*
                warrantbasic = EDLib.SQL.MSSQL.ExecSqlQry($@"SELECT A.[WID], A.[UID], A.[WClass], A.[IssuedDate],  A.[StrikePrice], A.[TraderAccount], ISNULL(B.[maturitydate],DATEADD(day,A.[Duration(Month)]*30,A.[IssuedDate])) AS MaturityDate
     FROM (SELECT [WID],[UID],[WName],[WClass], [IssuedDate], [MaturityDate], [StrikePrice], [TraderAccount], [Duration(Month)]
        FROM [newEDIS].[dbo].[WarrantBasics]
  WHERE ([MaturityDate] >'{DateTime.Today.ToString("yyyyMMdd")}' OR [MaturityDate] ='19110101') AND [WClass] ='{cp}' AND [IssuerName] ='{issuer}' ) AS A
  LEFT JOIN (SELECT [wname], [maturitydate]
  FROM [10.101.10.5].[HEDGE].[dbo].[WARRANTS]
  WHERE [kgiwrt] = '他家' and [maturitydate] > '{DateTime.Today.ToString("yyyyMMdd")}') as B on A.[WName] = B.[wname]
  WHERE A.[UID] = '{uid}'", GlobalVar.loginSet.newEDIS);
  */
                warrantbasic = EDLib.SQL.MSSQL.ExecSqlQry($@"SELECT * FROM(SELECT CASE [UID] WHEN 'TWA00' THEN 'IX0001' ELSE [UID] END AS [UID]
      ,[WClass]
      ,[IssuedDate]
      ,CASE WHEN([MaturityDate]='19110101') THEN DATEADD(MONTH,[Duration(Month)],[ListedDate]) ELSE [MaturityDate] END AS MaturityDate
      ,[StrikePrice]
      ,[TraderAccount]
  FROM [newEDIS].[dbo].[WarrantBasics]
  WHERE ([MaturityDate] >'{DateTime.Today.ToString("yyyyMMdd")}' OR [MaturityDate] = '19110101') and [WClass] ='{cp}' and [IssuerName] = '{issuer}') AS A
  WHERE A.[UID] = '{uid}'", GlobalVar.loginSet.newEDIS);
            }

       
            DataTable dtMatrix = new System.Data.DataTable();

            //just try Call
            //foreach (var uid in UID)

            //取得履約價中點
            DataRow[] select = dataTable1.Select($@"標的代號='{uid}'");
            if (select.Length > 0)
            {
                int check = Convert.ToInt32(select[0][3]);
                if (check == 0)
                {
                    MessageBox.Show("標的未設定參數");
                    return;
                }
                double get_quote = Convert.ToDouble(select[0][4].ToString());
                //履約價間隔(%)
                double intervalPercent = Convert.ToDouble(select[0][5].ToString());
                if (intervalPercent <= 0)
                {
                    MessageBox.Show("未設定履約價間距");
                    return;
                }
                double interval = roundpriceK(get_quote, get_quote * intervalPercent / 100);

                //價內監控區間
                int num_itm = (cp == "c" ? Convert.ToInt32(select[0][6].ToString()) : Convert.ToInt32(select[0][7].ToString()));
                //價內監控區間若小於4格，設4格
                int min_itm = num_itm >= 4 ? num_itm : 4;
                //價外監控區間
                int num_otm = (cp == "c" ? Convert.ToInt32(select[0][7].ToString()) : Convert.ToInt32(select[0][6].ToString()));
                //價外監控區間若小於4格，設4格
                int max_otm = num_otm >= 4 ? num_otm : 4;
                //距到期日起始月
                double maturity_start = Convert.ToDouble(select[0][8].ToString());
                //到期日間距
                double maturity_interval = Convert.ToDouble(select[0][9].ToString());
                //期間監控區
                int num_maturity = Convert.ToInt32(select[0][10].ToString());
                //最大期間監控區，若小於5格，設5格
                int max_num_maturity = num_maturity >= 5 ? num_maturity : 5;
                if (warrantbasic.Rows.Count > 0)
                {

                    //建表
                    dtMatrix.Columns.Add("Maturity", typeof(string));
                   
                    for (int i = min_itm; i > 0; i--)
                        dtMatrix.Columns.Add($@"{(get_quote - i * interval).ToString()}~{(get_quote - (i - 1) * interval).ToString()}", typeof(int));
                    for (int i = 1; i <= max_otm; i++)
                        dtMatrix.Columns.Add($@"{(get_quote + (i - 1) * interval).ToString()}~{(get_quote + i * interval).ToString()}", typeof(int));
                    

                    for (int i = 1; i <= max_num_maturity; i++)
                    {
                        DataRow dr = dtMatrix.NewRow();
                        DateTime ListedDat = EDLib.TradeDate.NextNTradeDate(2);
                        DateTime t1 = ListedDat.AddMonths((int)(maturity_start + maturity_interval * (i - 1))).AddDays(-1);
                        DateTime t2 = ListedDat.AddMonths((int)(maturity_start + maturity_interval * i)).AddDays(-1);
                        string t1str = t1.ToString("yyyy-MM-dd HH:mm:ss.fff");
                        string t2str = t2.ToString("yyyy-MM-dd HH:mm:ss.fff");
                        /*
                        string t1 = DateTime.Today.AddDays((int)Math.Round((maturity_start + maturity_interval * (i - 1)) * 30)).ToString("yyyy-MM-dd HH:mm:ss.fff");
                        DateTime t11 = DateTime.Today.AddDays((int)Math.Round((maturity_start + maturity_interval * (i - 1)) * 30));
                        string t2 = DateTime.Today.AddDays((int)Math.Round((maturity_start + maturity_interval * i) * 30)).ToString("yyyy-MM-dd HH:mm:ss.fff");
                        DateTime t22 = DateTime.Today.AddDays((int)Math.Round((maturity_start + maturity_interval * i) * 30));
                        dr[0] = (maturity_start + (i - 1) * maturity_interval).ToString() + "~" + (maturity_start + i * maturity_interval).ToString();
                        dr[0] = t1 + "~" + t2;
                        */
                        dr[0] = (maturity_start + (i - 1) * maturity_interval).ToString() + "~" + (maturity_start + i * maturity_interval).ToString();
                        //dr[0] = t1str + "~" + t2str;
                        int w = 1;
                        for (int j = min_itm; j > 0; j--)
                        {
                            DataRow[] temp = warrantbasic.Select($@"[MaturityDate] > '{t1}' AND [MaturityDate] <= '{t2}' AND [StrikePrice] >{get_quote - j * interval} AND [StrikePrice] <={get_quote - (j - 1) * interval}");
                            //MessageBox.Show($@"[MaturityDate] >= '{t11}' AND [MaturityDate] < '{t22}' AND [StrikePrice] >={get_quote - j * interval} AND [StrikePrice] <{get_quote - (j - 1) * interval}");
                            //DataRow[] temp = warrantbasic.Select($@"[MaturityDate] >= '{t1}' AND [MaturityDate] < '{t2}' AND [StrikePrice] >={get_quote - j * interval} AND [StrikePrice] <{get_quote - (j - 1) * interval}");
                            //MessageBox.Show($@"[MaturityDate] >= '{t1}' AND [MaturityDate] < '{t2}' AND [StrikePrice] >={get_quote - j * interval} AND [StrikePrice] <{get_quote - (j - 1) * interval}");
            
                            dr[w] = temp.Length;
                            w++;
                        }
                        for (int j = 1; j <= max_otm; j++)
                        {
                            DataRow[] temp = warrantbasic.Select($@"[MaturityDate] > '{t1}' AND [MaturityDate] <= '{t2}' AND [StrikePrice] >{get_quote + (j - 1) * interval} AND [StrikePrice] <={get_quote + (j) * interval}");
                            dr[w] = temp.Length;
                            w++;
                        }
                        //temp = warrantbasic.Select($@"[MaturityDate] >= '{t11}' AND [MaturityDate] < '{t22}' AND [StrikePrice] >={max}");
                        //dr[w] = temp.Length;
                        dtMatrix.Rows.Add(dr);
                    }
                }
                else
                {
                    //建表
                    dtMatrix.Columns.Add("Maturity", typeof(string));
                    
                    for (int i = min_itm; i > 0; i--)
                        dtMatrix.Columns.Add($@"{(get_quote - i * interval).ToString()}~{(get_quote - (i - 1) * interval).ToString()}", typeof(int));
                    for (int i = 1; i <= max_otm; i++)
                        dtMatrix.Columns.Add($@"{(get_quote + (i - 1) * interval).ToString()}~{(get_quote + i * interval).ToString()}", typeof(int));
        

                    for (int i = 1; i <= max_num_maturity; i++)
                    {
                        DataRow dr = dtMatrix.NewRow();
                        DateTime ListedDat = EDLib.TradeDate.NextNTradeDate(2);
                        DateTime t1 = ListedDat.AddMonths((int)(maturity_start + maturity_interval * (i - 1))).AddDays(-1);
                        DateTime t2 = ListedDat.AddMonths((int)(maturity_start + maturity_interval * i)).AddDays(-1);
                        string t1str = t1.ToString("yyyy-MM-dd HH:mm:ss.fff");
                        string t2str = t2.ToString("yyyy-MM-dd HH:mm:ss.fff");
                        /*
                        string t1 = DateTime.Today.AddDays((int)Math.Round((maturity_start + maturity_interval * (i - 1)) * 30)).ToString("yyyy-MM-dd HH:mm:ss.fff");
                        DateTime t11 = DateTime.Today.AddDays((int)Math.Round((maturity_start + maturity_interval * (i - 1)) * 30));
                        string t2 = DateTime.Today.AddDays((int)Math.Round((maturity_start + maturity_interval * i) * 30)).ToString("yyyy-MM-dd HH:mm:ss.fff");
                        DateTime t22 = DateTime.Today.AddDays((int)Math.Round((maturity_start + maturity_interval * i) * 30));
                        */
                        dr[0] = (maturity_start + (i - 1) * maturity_interval).ToString() + "~" + (maturity_start + i * maturity_interval).ToString();
                        //dr[0] = t1str + "~" + t2str;
                       
                        int w = 1;
                        for (int j = min_itm; j > 0; j--)
                        {
                            dr[w] = 0;
                            w++;
                        }
                        for (int j = 1; j <= max_otm; j++)
                        {
                            dr[w] = 0;
                            w++;
                        }
                        //temp = warrantbasic.Select($@"[MaturityDate] >= '{t11}' AND [MaturityDate] < '{t22}' AND [StrikePrice] >={max}");
                        //dr[w] = temp.Length;
                        dtMatrix.Rows.Add(dr);
                    }
                }
                ultraGrid3.DataSource = dtMatrix;
                SetMatrix_NonKGI(min_itm, max_otm, num_itm, num_otm, num_maturity);
            }
            else
            {
                MessageBox.Show("標的錯誤或無設定此標的");
            }
        }

        private void SetMatrix_NonKGI(int min_itm, int max_otm, int num_itm, int num_otm, int num_maturity)
        {
            UltraGridBand bands0 = ultraGrid3.DisplayLayout.Bands[0];
            bands0.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
            this.ultraGrid3.DisplayLayout.AutoFitStyle = Infragistics.Win.UltraWinGrid.AutoFitStyle.ResizeAllColumns;
            this.ultraGrid3.DisplayLayout.Override.DefaultRowHeight = 50;
            this.ultraGrid3.DisplayLayout.Override.CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center;
            this.ultraGrid3.DisplayLayout.Override.CellAppearance.TextVAlign = Infragistics.Win.VAlign.Middle;
            
            foreach (var col in bands0.Columns)
            {
                col.CellActivation = Activation.NoEdit;
            }
            //ultraGrid2.DisplayLayout.Override.CellAppearance.BackColor = Color.Moccasin;
            for (int i = 0; i < num_maturity; i++)
            {
                for (int j = (min_itm + 1) - num_itm; j < (min_itm + 1); j++)
                {
                    ultraGrid3.Rows[i].Cells[j].Appearance.BackColor = Color.Orange;
                }
                for (int j = (min_itm + 1); j < (min_itm + 1) + num_otm; j++)
                {
                    ultraGrid3.Rows[i].Cells[j].Appearance.BackColor = Color.SkyBlue;
                }

            }
        }
        private void textBox1_KeyDown(object sender, KeyEventArgs e)
        {
            try
            {
                if (e.KeyCode == Keys.Enter)
                {
                    ShowMatrix();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void textBox2_KeyDown(object sender, KeyEventArgs e)
        {
            try
            {
                if (e.KeyCode == Keys.Enter)
                {
                    ShowMatrix_NonKGI();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
        private void ultraGrid1_InitializeRow(object sender, InitializeRowEventArgs e)
        {
            string uid = e.Row.Cells["標的代號"].Value.ToString();
            int num_itm = Convert.ToInt32(e.Row.Cells["價內監控區間"].Value.ToString());
            int num_otm = Convert.ToInt32(e.Row.Cells["價外監控區間"].Value.ToString());
            int num_maturity = Convert.ToInt32(e.Row.Cells["期間監控區"].Value.ToString());
            double interval = Convert.ToDouble(e.Row.Cells["履約價間距(%)"].Value);
            int check = Convert.ToInt32(e.Row.Cells["確認"].Value);
            double maturity_start = Convert.ToDouble(e.Row.Cells["距到期日起始月"].Value);
            double maturity_interval = Convert.ToDouble(e.Row.Cells["到期日間距(月)"].Value);
            double maturity_start_ceiling = Math.Ceiling(maturity_start);
            double maturity_interval_ceiling = Math.Ceiling(maturity_interval);
            if (interval <= 0 && check == 1)
            {
                MessageBox.Show($@"{uid} 未設定履約價間距");
            }
            if(((maturity_interval_ceiling-maturity_interval)>0 || (maturity_start_ceiling - maturity_start) > 0) && check == 1)
            {
                MessageBox.Show($@"{uid} 距到期日起始月 / 到期日間距須為整數");
            }
            /*
            if (((num_itm > 4) || (num_otm > 4) || (num_maturity > 5)) && !isEdit && check == 1)
            {
                MessageBox.Show($@"{uid} 資料輸入有誤");
            }
            */
        }
        private bool CheckData()
        {
            foreach (Infragistics.Win.UltraWinGrid.UltraGridRow r in ultraGrid1.Rows)
            {
                int check = Convert.ToInt32(r.Cells["確認"].Value);
                if (check == 0)
                    continue;
                double interval = Convert.ToDouble(r.Cells["履約價間距(%)"].Value);
                
                string uid = r.Cells["標的代號"].Value.ToString();
                int num_itm = Convert.ToInt32(r.Cells["價內監控區間"].Value.ToString());
                int num_otm = Convert.ToInt32(r.Cells["價外監控區間"].Value.ToString());
                int num_maturity = Convert.ToInt32(r.Cells["期間監控區"].Value.ToString());
                double maturity_start = Convert.ToDouble(r.Cells["距到期日起始月"].Value.ToString());
                double maturity_interval = Convert.ToDouble(r.Cells["到期日間距(月)"].Value.ToString());
                double maturity_start_ceiling = Math.Ceiling(maturity_start);
                double maturity_interval_ceiling = Math.Ceiling(maturity_interval);
                if (interval <= 0)
                {
                    MessageBox.Show($@"{uid} 未設定履約價間距");
                    return false;
                }
                if ((maturity_interval_ceiling - maturity_interval) > 0 || (maturity_start_ceiling - maturity_start) > 0)
                {
                    MessageBox.Show($@"{uid} 距到期日起始月 / 到期日間距不為整數");
                    return false;
                }
                /*
                if (num_itm > 4 || num_otm > 4)
                {
                    MessageBox.Show($@"{uid} 價內外範圍超過4個區間");
                    return false;
                }
                if (num_maturity>5)
                {
                    MessageBox.Show($@"{uid} 期間監控區超過5個區間");
                    return false;
                }
                */
            }
            return true;
        }

        //更新完設定後要重新算矩陣
        public static void ReCalCallPutDensity()
        {
            DateTime today = DateTime.Today;
            DateTime lastday = EDLib.TradeDate.LastNTradeDate(1);
            string todaystr = today.ToString("yyyyMMdd");
            string lastdaystr = lastday.ToString("yyyyMMdd");
            List<string> UID = new List<string>();
#if !To39
            string sql = $@"SELECT [UID]
            FROM [newEDIS].[dbo].[OptionAutoSelectData]";

            DataTable dt = EDLib.SQL.MSSQL.ExecSqlQry(sql, GlobalVar.loginSet.newEDIS);
#else
            string sql = $@"SELECT [UID]
            FROM [WarrantAssistant].[dbo].[OptionAutoSelectData]";

            DataTable dt = EDLib.SQL.MSSQL.ExecSqlQry(sql, GlobalVar.loginSet.warrantassistant45);
#endif
            foreach (DataRow dr in dt.Rows)
            {
                string uid = dr["UID"].ToString();
                UID.Add(uid);
            }
            //先算Call
            string getprice = $@"SELECT DISTINCT CASE WHEN ([CommodityId]='1000') THEN 'IX0001' ELSE [CommodityId] END AS CommodityID
                                             ,ISNULL([LastPrice],0) AS LastPrice
                                             ,[tradedate] AS TradeDate
                                             ,isnull([BuyPriceBest1],0) AS Bid1
                                             ,isnull([SellPriceBest1],0) AS Ask1
                               FROM [TsQuote].[dbo].[vwprice2]
							   WHERE [kind] in ('ETF','Index','Stock');";
            DataTable vprice = EDLib.SQL.MSSQL.ExecSqlQry(getprice, GlobalVar.loginSet.tsquoteSqlConnString);
            //從WarrantPrice取及時價格
#if !To39
            string Cmatrixsetting = $@"SELECT [TraderID],[UID],[CP],[Interval],[StartMonth],[IntervalMonth],[Interval_InTheMoney]
                                        ,[Interval_OutOfMoney],[Num_IntervalMonth],[Checked]
                                     FROM [newEDIS].[dbo].[OptionAutoSelectMatrix]
                                     WHERE [Checked] =1 AND [CP] ='C'";
            string Pmatrixsetting = $@"SELECT [TraderID],[UID],[CP],[Interval],[StartMonth],[IntervalMonth],[Interval_InTheMoney]
                                        ,[Interval_OutOfMoney],[Num_IntervalMonth],[Checked]
                                     FROM [newEDIS].[dbo].[OptionAutoSelectMatrix]
                                     WHERE [Checked] =1 AND [CP] ='P'";
            DataTable Cmatrix = EDLib.SQL.MSSQL.ExecSqlQry(Cmatrixsetting, GlobalVar.loginSet.newEDIS);
            DataTable Pmatrix = EDLib.SQL.MSSQL.ExecSqlQry(Pmatrixsetting, GlobalVar.loginSet.newEDIS);
#else
            string Cmatrixsetting = $@"SELECT [TraderID],[UID],[CP],[Interval],[StartMonth],[IntervalMonth],[Interval_InTheMoney]
                                        ,[Interval_OutOfMoney],[Num_IntervalMonth],[Checked]
                                     FROM [WarrantAssistant].[dbo].[OptionAutoSelectMatrix]
                                     WHERE [Checked] =1 AND [CP] ='C'";
            string Pmatrixsetting = $@"SELECT [TraderID],[UID],[CP],[Interval],[StartMonth],[IntervalMonth],[Interval_InTheMoney]
                                        ,[Interval_OutOfMoney],[Num_IntervalMonth],[Checked]
                                     FROM [WarrantAssistant].[dbo].[OptionAutoSelectMatrix]
                                     WHERE [Checked] =1 AND [CP] ='P'";
            DataTable Cmatrix = EDLib.SQL.MSSQL.ExecSqlQry(Cmatrixsetting, GlobalVar.loginSet.warrantassistant45);
            DataTable Pmatrix = EDLib.SQL.MSSQL.ExecSqlQry(Pmatrixsetting, GlobalVar.loginSet.warrantassistant45);
#endif
            DataTable warrantbasic = EDLib.SQL.MSSQL.ExecSqlQry($@"SELECT  [stkid] AS [UID]
                                              , CASE WHEN (SUBSTRING([type],4,4) ='認購權證') THEN 'c' ELSE 'p' END AS WClass
	                                          , [issuedate] AS IssuedDate
                                              , [strike_now] AS StrikePrice
                                              , [maturitydate] AS MaturityDate
	                                          , [TraderID] AS TraderAccount
                                          FROM [HEDGE].[dbo].[WARRANTS]
                                          WHERE [kgiwrt] = '自家' AND [maturitydate] >'{DateTime.Today.ToString("yyyyMMdd")}' AND ([type] LIKE '%認購權證%'  OR [type] LIKE '%認售權證%')", "Data Source=10.101.10.5;Initial Catalog=HEDGE;User ID=hedgeuser;Password=hedgeuser");
            //just try Call
            foreach (var uid in UID)
            {
                DataTable dtMatrix = new DataTable();
                DataRow[] select = vprice.Select($@"CommodityID='{uid}'");
                if (select.Length == 0)
                    continue;
                //取得履約價中點
                double get_quote;
                if (Convert.ToDouble(select[0][1].ToString()) == 0)
                {
                    if (Convert.ToDouble(select[0][3].ToString()) == 0)
                    {
                        if (Convert.ToDouble(select[0][4].ToString()) == 0)
                            get_quote = 0;
                        else
                            get_quote = roundprice(Convert.ToDouble(select[0][4].ToString()));
                    }
                    else
                    {
                        get_quote = roundprice(Convert.ToDouble(select[0][3].ToString()));
                    }

                }
                else
                {
                    get_quote = roundprice(Convert.ToDouble(select[0][1].ToString()));
                }

                DataRow[] select2 = Cmatrix.Select($@"UID='{uid}'");
                //using (StreamWriter sw = new StreamWriter("D:\\density.txt", true))
                {
                    //try
                    {
                        if (select2.Length > 0)
                        {
                            //履約價間隔
                            double intervalPercent = Convert.ToDouble(select2[0][3].ToString());
                            if (intervalPercent <= 0)
                                continue;
                            //價內監控區間
                            double interval = roundpriceK(get_quote, get_quote * intervalPercent / 100);
                            int num_itm = Convert.ToInt32(select2[0][6].ToString());
                            //價外監控區間
                            int num_otm = Convert.ToInt32(select2[0][7].ToString());
                            //距到期日起始月
                            double maturity_start = Convert.ToDouble(select2[0][4].ToString());
                            //到期日間距
                            double maturity_interval = Convert.ToDouble(select2[0][5].ToString());
                            //期間監控區
                            int num_maturity = Convert.ToInt32(select2[0][8].ToString());
                            DataRow[] warrantbasicselect = warrantbasic.Select($@"UID='{uid}'");
                            
                            if (warrantbasicselect.Length > 0)
                            {
                                //建表

                                for (int i = num_itm; i > 0; i--)
                                    dtMatrix.Columns.Add($@"{(get_quote - i * interval).ToString()}~{(get_quote - (i - 1) * interval).ToString()}", typeof(int));
                                for (int i = 1; i <= num_otm; i++)
                                    dtMatrix.Columns.Add($@"{(get_quote + (i - 1) * interval).ToString()}~{(get_quote + i * interval).ToString()}", typeof(int));
                                //double max = get_quote + interval * num_otm;
                                //dtMatrix.Columns.Add($@">{max.ToString()}", typeof(int));

                                for (int i = 1; i <= num_maturity; i++)
                                {
                                    DataRow dr = dtMatrix.NewRow();
                                    /*
                                    string t1 = DateTime.Today.AddDays((int)Math.Round((maturity_start + maturity_interval * (i - 1)) * 30)).ToString("yyyyMMdd");
                                    DateTime t11 = DateTime.Today.AddDays((int)Math.Round((maturity_start + maturity_interval * (i - 1)) * 30));
                                    string t2 = DateTime.Today.AddDays((int)Math.Round((maturity_start + maturity_interval * i) * 30)).ToString("yyyyMMdd");
                                    DateTime t22 = DateTime.Today.AddDays((int)Math.Round((maturity_start + maturity_interval * i) * 30));
                                    */
                                    DateTime ListedDat = EDLib.TradeDate.NextNTradeDate(2);
                                    DateTime t1 = ListedDat.AddMonths((int)(maturity_start + maturity_interval * (i - 1))).AddDays(-1);
                                    DateTime t2 = ListedDat.AddMonths((int)(maturity_start + maturity_interval * i)).AddDays(-1);
                                    string t1str = t1.ToString("yyyy-MM-dd HH:mm:ss.fff");
                                    string t2str = t2.ToString("yyyy-MM-dd HH:mm:ss.fff");


                                    int w = 0;
                                    for (int j = num_itm; j > 0; j--)
                                    {
                                        DataRow[] temp = warrantbasic.Select($@"[UID]='{uid}' AND [WClass]='c' AND [MaturityDate] > '{t1}' AND [MaturityDate] <= '{t2}' AND [StrikePrice] >{get_quote - j * interval} AND [StrikePrice] <={get_quote - (j - 1) * interval}");
                                        dr[w] = temp.Length;
                                        w++;
                                    }

                                    for (int j = 1; j <= num_otm; j++)
                                    {
                                        DataRow[] temp = warrantbasic.Select($@"[UID]='{uid}' AND [WClass]='c' AND [MaturityDate] > '{t1}' AND [MaturityDate] <= '{t2}' AND [StrikePrice] >{get_quote + (j - 1) * interval} AND [StrikePrice] <={get_quote + (j) * interval}");
                                        dr[w] = temp.Length;
                                        w++;
                                    }
                                    dtMatrix.Rows.Add(dr);

                                }
                            }
                            
                            int density = 0;
                            if (dtMatrix.Rows.Count > 0)
                            {
                                for (int i = 0; i < num_maturity; i++)
                                {
                                    for (int j = 0; j < (num_itm + num_otm); j++)
                                    {
                                        /*
                                        if (Convert.ToInt16(dtMatrix.Rows[i][j].ToString()) == 0)
                                            density++;
                                        */
                                        if (Convert.ToInt16(dtMatrix.Rows[i][j].ToString()) > 0)
                                            density++;
                                    }
                                }
                            }
                            /*
                            else
                            {
                                density = (num_itm + num_otm) * num_maturity;
                            }
                            */
                            

                            string sqlInsert = $@"UPDATE [newEDIS].[dbo].[OptionAutoSelectData]
                                    SET [CallDensity] = {density}
                                    WHERE [UID] ='{uid}'";
                            MSSQL.ExecSqlCmd(sqlInsert, GlobalVar.loginSet.newEDIS);

                        }
                    }
                    /*
                    catch (Exception ex)
                    {
                        MessageBox.Show($@"C {uid}   {ex.Message}");
                        //sw.WriteLine($@"C {uid}   {ex.Message}");
                    }
                    */

                }

            }
            //just try Put
            foreach (var uid in UID)
            {
                DataTable dtMatrix = new DataTable();
                DataRow[] select = vprice.Select($@"CommodityID='{uid}'");
                if (select.Length == 0)
                    continue;
                //取得履約價中點
                
                double get_quote;
                if (Convert.ToDouble(select[0][1].ToString()) == 0)
                {
                    if (Convert.ToDouble(select[0][3].ToString()) == 0)
                    {
                        if (Convert.ToDouble(select[0][4].ToString()) == 0)
                            get_quote = 0;
                        else
                            get_quote = roundprice(Convert.ToDouble(select[0][4].ToString()));
                    }
                    else
                    {
                        get_quote = roundprice(Convert.ToDouble(select[0][3].ToString()));
                    }

                }
                else
                {
                    get_quote = roundprice(Convert.ToDouble(select[0][1].ToString()));
                }
                
                DataRow[] select2 = Pmatrix.Select($@"UID='{uid}'");
                //using (StreamWriter sw = new StreamWriter("D:\\density.txt", true))
                {
                    //try
                    {
                        if (select2.Length > 0)
                        {

                            //履約價間隔
                            double intervalPercent = Convert.ToDouble(select2[0][3].ToString());
                            if (intervalPercent <= 0)
                                continue;
                            //價內監控區間
                            double interval = roundpriceK(get_quote, get_quote * intervalPercent / 100);
                            //Put價內外要顛倒
                            int num_itm = Convert.ToInt32(select2[0][7].ToString());
                            //價外監控區間
                            int num_otm = Convert.ToInt32(select2[0][6].ToString());
                            //距到期日起始月
                            double maturity_start = Convert.ToDouble(select2[0][4].ToString());
                            //到期日間距
                            double maturity_interval = Convert.ToDouble(select2[0][5].ToString());
                            //期間監控區
                            int num_maturity = Convert.ToInt32(select2[0][8].ToString());
                            DataRow[] warrantbasicselect = warrantbasic.Select($@"UID='{uid}'");
                            
                            if (warrantbasicselect.Length > 0)
                            {

                                //建表

                                for (int i = num_itm; i > 0; i--)
                                    dtMatrix.Columns.Add($@"{(get_quote - i * interval).ToString()}~{(get_quote - (i - 1) * interval).ToString()}", typeof(int));
                                for (int i = 1; i <= num_otm; i++)
                                    dtMatrix.Columns.Add($@"{(get_quote + (i - 1) * interval).ToString()}~{(get_quote + i * interval).ToString()}", typeof(int));
                                //double max = get_quote + interval * num_otm;
                                //dtMatrix.Columns.Add($@">{max.ToString()}", typeof(int));

                                for (int i = 1; i <= num_maturity; i++)
                                {
                                    DataRow dr = dtMatrix.NewRow();
                                    /*
                                    string t1 = DateTime.Today.AddDays((int)Math.Round((maturity_start + maturity_interval * (i - 1)) * 30)).ToString("yyyyMMdd");
                                    DateTime t11 = DateTime.Today.AddDays((int)Math.Round((maturity_start + maturity_interval * (i - 1)) * 30));
                                    string t2 = DateTime.Today.AddDays((int)Math.Round((maturity_start + maturity_interval * i) * 30)).ToString("yyyyMMdd");
                                    DateTime t22 = DateTime.Today.AddDays((int)Math.Round((maturity_start + maturity_interval * i) * 30));
                                    */
                                    DateTime ListedDat = EDLib.TradeDate.NextNTradeDate(2);
                                    DateTime t1 = ListedDat.AddMonths((int)(maturity_start + maturity_interval * (i - 1))).AddDays(-1);
                                    DateTime t2 = ListedDat.AddMonths((int)(maturity_start + maturity_interval * i)).AddDays(-1);
                                    string t1str = t1.ToString("yyyy-MM-dd HH:mm:ss.fff");
                                    string t2str = t2.ToString("yyyy-MM-dd HH:mm:ss.fff");
                                    int w = 0;
                                    for (int j = num_itm; j > 0; j--)
                                    {
                                        DataRow[] temp = warrantbasic.Select($@"[UID]='{uid}' AND [WClass]='p' AND [MaturityDate] > '{t1}' AND [MaturityDate] <= '{t2}' AND [StrikePrice] >{get_quote - j * interval} AND [StrikePrice] <={get_quote - (j - 1) * interval}");
                                        dr[w] = temp.Length;
                                        w++;
                                    }

                                    for (int j = 1; j <= num_otm; j++)
                                    {
                                        DataRow[] temp = warrantbasic.Select($@"[UID]='{uid}' AND [WClass]='p' AND [MaturityDate] > '{t1}' AND [MaturityDate] <= '{t2}' AND [StrikePrice] >{get_quote + (j - 1) * interval} AND [StrikePrice] <={get_quote + (j) * interval}");
                                        dr[w] = temp.Length;
                                        w++;
                                    }
                                    dtMatrix.Rows.Add(dr);

                                }
                            }
                            
                            int density = 0;
                          
                            if (dtMatrix.Rows.Count > 0)
                            {
                                for (int i = 0; i < num_maturity; i++)
                                {
                                    
                                    for (int j = 0; j < (num_itm + num_otm); j++)
                                    {
                                        /*
                                        if (Convert.ToInt16(dtMatrix.Rows[i][j].ToString()) == 0)
                                            density++;
                                        */
                                        if (Convert.ToInt16(dtMatrix.Rows[i][j].ToString()) > 0)
                                            density++;
                                    }

                                }
                            }
                            /*
                            else
                            {
                                density = (num_itm + num_otm) * num_maturity;
                            }
                            */


#if !To39
                            string sqlInsert = $@"UPDATE [newEDIS].[dbo].[OptionAutoSelectData]
                                        SET [PutDensity] = {density}
                                        WHERE [UID] ='{uid}'";
                            MSSQL.ExecSqlCmd(sqlInsert, GlobalVar.loginSet.newEDIS);
#else
                            string sqlInsert = $@"UPDATE [WarrantAssistant].[dbo].[OptionAutoSelectData]
                                        SET [PutDensity] = {density}
                                        WHERE [UID] ='{uid}'";
                            MSSQL.ExecSqlCmd(sqlInsert, GlobalVar.loginSet.warrantassistant45);
#endif

                        }
                    }
                    /*
                    catch (Exception ex)
                    {
                        sw.WriteLine($@"P ii:{ii} {uid}   {ex.Message}");
                    }
                    */
                }

            }
        }

        public static void ReCalYuanCallPutDensity()
        {
            DateTime today = DateTime.Today;
            DateTime lastday = EDLib.TradeDate.LastNTradeDate(1);
            string todaystr = today.ToString("yyyyMMdd");
            string lastdaystr = lastday.ToString("yyyyMMdd");
            List<string> UID = new List<string>();
#if !To39
            string sql = $@"SELECT [UID]
            FROM [newEDIS].[dbo].[OptionAutoSelectData]";

            DataTable dt = EDLib.SQL.MSSQL.ExecSqlQry(sql, GlobalVar.loginSet.newEDIS);
#else
            string sql = $@"SELECT [UID]
            FROM [WarrantAssistant].[dbo].[OptionAutoSelectData]";

            DataTable dt = EDLib.SQL.MSSQL.ExecSqlQry(sql, GlobalVar.loginSet.warrantassistant45);
#endif
            foreach (DataRow dr in dt.Rows)
            {
                string uid = dr["UID"].ToString();
                UID.Add(uid);
            }
            //先算Call
            string getprice = $@"SELECT DISTINCT CASE WHEN ([CommodityId]='1000') THEN 'IX0001' ELSE [CommodityId] END AS CommodityID
                                             ,ISNULL([LastPrice],0) AS LastPrice
                                             ,[tradedate] AS TradeDate
                                             ,isnull([BuyPriceBest1],0) AS Bid1
                                             ,isnull([SellPriceBest1],0) AS Ask1
                               FROM [TsQuote].[dbo].[vwprice2]
							   WHERE [kind] in ('ETF','Index','Stock');";
            DataTable vprice = EDLib.SQL.MSSQL.ExecSqlQry(getprice, GlobalVar.loginSet.tsquoteSqlConnString);
            //從WarrantPrice取及時價格
#if !To39
            string Cmatrixsetting = $@"SELECT [TraderID],[UID],[CP],[Interval],[StartMonth],[IntervalMonth],[Interval_InTheMoney]
                                        ,[Interval_OutOfMoney],[Num_IntervalMonth],[Checked]
                                     FROM [newEDIS].[dbo].[OptionAutoSelectMatrix]
                                     WHERE [Checked] =1 AND [CP] ='C'";
            string Pmatrixsetting = $@"SELECT [TraderID],[UID],[CP],[Interval],[StartMonth],[IntervalMonth],[Interval_InTheMoney]
                                        ,[Interval_OutOfMoney],[Num_IntervalMonth],[Checked]
                                     FROM [newEDIS].[dbo].[OptionAutoSelectMatrix]
                                     WHERE [Checked] =1 AND [CP] ='P'";
            DataTable Cmatrix = EDLib.SQL.MSSQL.ExecSqlQry(Cmatrixsetting, GlobalVar.loginSet.newEDIS);
            DataTable Pmatrix = EDLib.SQL.MSSQL.ExecSqlQry(Pmatrixsetting, GlobalVar.loginSet.newEDIS);
#else
            string Cmatrixsetting = $@"SELECT [TraderID],[UID],[CP],[Interval],[StartMonth],[IntervalMonth],[Interval_InTheMoney]
                                        ,[Interval_OutOfMoney],[Num_IntervalMonth],[Checked]
                                     FROM [WarrantAssistant].[dbo].[OptionAutoSelectMatrix]
                                     WHERE [Checked] =1 AND [CP] ='C'";
            string Pmatrixsetting = $@"SELECT [TraderID],[UID],[CP],[Interval],[StartMonth],[IntervalMonth],[Interval_InTheMoney]
                                        ,[Interval_OutOfMoney],[Num_IntervalMonth],[Checked]
                                     FROM [WarrantAssistant].[dbo].[OptionAutoSelectMatrix]
                                     WHERE [Checked] =1 AND [CP] ='P'";
            DataTable Cmatrix = EDLib.SQL.MSSQL.ExecSqlQry(Cmatrixsetting, GlobalVar.loginSet.warrantassistant45);
            DataTable Pmatrix = EDLib.SQL.MSSQL.ExecSqlQry(Pmatrixsetting, GlobalVar.loginSet.warrantassistant45);
#endif
           
            DataTable warrantbasic = EDLib.SQL.MSSQL.ExecSqlQry($@"SELECT  CASE [標的代號] WHEN 'TWA00' THEN 'IX0001' ELSE [標的代號] END AS [UID]
                       ,CASE WHEN [名稱] LIKE '%購%' THEN 'c' ELSE 'p' END AS [WClass]
	                   ,[發行日期] AS [IssuedDate]
	                   ,CASE WHEN([到期日期]='19110101') THEN DATEADD(MONTH,[存續期間(月)],[發行日期]) ELSE [到期日期] END AS MaturityDate
                      ,[最新履約價] AS [StrikePrice]
                  FROM [TwCMData].[dbo].[Warrant總表]
                  WHERE [日期] = '{lastdaystr}' AND ([到期日期] >'{todaystr}' OR [到期日期] = '19110101') AND ([名稱] LIKE '%購%' OR [名稱] LIKE '%售%') AND [券商代號] ='9800'",GlobalVar.loginSet.twCMData);
            //just try Call
            foreach (var uid in UID)
            {
                DataTable dtMatrix = new DataTable();
                DataRow[] select = vprice.Select($@"CommodityID='{uid}'");
                if (select.Length == 0)
                    continue;
                //取得履約價中點
                double get_quote;
                if (Convert.ToDouble(select[0][1].ToString()) == 0)
                {
                    if (Convert.ToDouble(select[0][3].ToString()) == 0)
                    {
                        if (Convert.ToDouble(select[0][4].ToString()) == 0)
                            get_quote = 0;
                        else
                            get_quote = roundprice(Convert.ToDouble(select[0][4].ToString()));
                    }
                    else
                    {
                        get_quote = roundprice(Convert.ToDouble(select[0][3].ToString()));
                    }

                }
                else
                {
                    get_quote = roundprice(Convert.ToDouble(select[0][1].ToString()));
                }

                DataRow[] select2 = Cmatrix.Select($@"UID='{uid}'");
                using (StreamWriter sw = new StreamWriter("D:\\density.txt", true))
                {
                    try
                    {
                        if (select2.Length > 0)
                        {

                            //履約價間隔
                            double intervalPercent = Convert.ToDouble(select2[0][3].ToString());
                            if (intervalPercent <= 0)
                                continue;
                            //價內監控區間
                            double interval = roundpriceK(get_quote, get_quote * intervalPercent / 100);
                            int num_itm = Convert.ToInt32(select2[0][6].ToString());
                            //價外監控區間
                            int num_otm = Convert.ToInt32(select2[0][7].ToString());
                            //距到期日起始月
                            double maturity_start = Convert.ToDouble(select2[0][4].ToString());
                            //到期日間距
                            double maturity_interval = Convert.ToDouble(select2[0][5].ToString());
                            //期間監控區
                            int num_maturity = Convert.ToInt32(select2[0][8].ToString());
                            DataRow[] warrantbasicselect = warrantbasic.Select($@"UID='{uid}'");

                            if (warrantbasicselect.Length > 0)
                            {

                                //建表

                                for (int i = num_itm; i > 0; i--)
                                    dtMatrix.Columns.Add($@"{(get_quote - i * interval).ToString()}~{(get_quote - (i - 1) * interval).ToString()}", typeof(int));
                                for (int i = 1; i <= num_otm; i++)
                                    dtMatrix.Columns.Add($@"{(get_quote + (i - 1) * interval).ToString()}~{(get_quote + i * interval).ToString()}", typeof(int));
                                //double max = get_quote + interval * num_otm;
                                //dtMatrix.Columns.Add($@">{max.ToString()}", typeof(int));

                                for (int i = 1; i <= num_maturity; i++)
                                {
                                    DataRow dr = dtMatrix.NewRow();
                                    /*
                                    string t1 = DateTime.Today.AddDays((int)Math.Round((maturity_start + maturity_interval * (i - 1)) * 30)).ToString("yyyyMMdd");
                                    DateTime t11 = DateTime.Today.AddDays((int)Math.Round((maturity_start + maturity_interval * (i - 1)) * 30));
                                    string t2 = DateTime.Today.AddDays((int)Math.Round((maturity_start + maturity_interval * i) * 30)).ToString("yyyyMMdd");
                                    DateTime t22 = DateTime.Today.AddDays((int)Math.Round((maturity_start + maturity_interval * i) * 30));
                                    */
                                    DateTime ListedDat = EDLib.TradeDate.NextNTradeDate(2);
                                    DateTime t1 = ListedDat.AddMonths((int)(maturity_start + maturity_interval * (i - 1))).AddDays(-1);
                                    DateTime t2 = ListedDat.AddMonths((int)(maturity_start + maturity_interval * i)).AddDays(-1);
                                    string t1str = t1.ToString("yyyy-MM-dd HH:mm:ss.fff");
                                    string t2str = t2.ToString("yyyy-MM-dd HH:mm:ss.fff");

                                    //小於等於上界
                                    int w = 0;
                                    for (int j = num_itm; j > 0; j--)
                                    {
                                        DataRow[] temp = warrantbasic.Select($@"[UID]='{uid}' AND [WClass]='c' AND [MaturityDate] > '{t1}' AND [MaturityDate] <= '{t2}' AND [StrikePrice] >{get_quote - j * interval} AND [StrikePrice] <={get_quote - (j - 1) * interval}");
                                        dr[w] = temp.Length;
                                        w++;
                                    }

                                    for (int j = 1; j <= num_otm; j++)
                                    {
                                        DataRow[] temp = warrantbasic.Select($@"[UID]='{uid}' AND [WClass]='c' AND [MaturityDate] > '{t1}' AND [MaturityDate] <= '{t2}' AND [StrikePrice] >{get_quote + (j - 1) * interval} AND [StrikePrice] <={get_quote + (j) * interval}");
                                        dr[w] = temp.Length;
                                        w++;
                                    }
                                    dtMatrix.Rows.Add(dr);

                                }
                            }

                            int density = 0;
                            if (dtMatrix.Rows.Count > 0)
                            {
                                for (int i = 0; i < num_maturity; i++)
                                {
                                    for (int j = 0; j < (num_itm + num_otm); j++)
                                    {
                                        /*
                                        if (Convert.ToInt16(dtMatrix.Rows[i][j].ToString()) == 0)
                                            density++;
                                        */
                                        if (Convert.ToInt16(dtMatrix.Rows[i][j].ToString()) > 0)
                                            density++;
                                    }

                                }
                            }
                            /*
                            else
                            {
                                density = (num_itm + num_otm) * num_maturity;
                            }
                            */
#if !To39
                            string sqlInsert = $@"UPDATE [newEDIS].[dbo].[OptionAutoSelectData]
                                        SET [YuanCallDensity] = {density}
                                        WHERE [UID] ='{uid}'";
                            MSSQL.ExecSqlCmd(sqlInsert, GlobalVar.loginSet.newEDIS);
#else
                            string sqlInsert = $@"UPDATE [WarrantAssistant].[dbo].[OptionAutoSelectData]
                                        SET [YuanCallDensity] = {density}
                                        WHERE [UID] ='{uid}'";
                            MSSQL.ExecSqlCmd(sqlInsert, GlobalVar.loginSet.warrantassistant45);
#endif
                        }
                        else
                        {
#if !To39
                            string sqlInsert = $@"UPDATE [newEDIS].[dbo].[OptionAutoSelectData]
                                        SET [YuanCallDensity] = -1
                                        WHERE [UID] ='{uid}'";
                            MSSQL.ExecSqlCmd(sqlInsert, GlobalVar.loginSet.newEDIS);
#else
                            string sqlInsert = $@"UPDATE [WarrantAssistant].[dbo].[OptionAutoSelectData]
                                        SET [YuanCallDensity] = -1
                                        WHERE [UID] ='{uid}'";
                            MSSQL.ExecSqlCmd(sqlInsert, GlobalVar.loginSet.warrantassistant45);
#endif
                        }
                    }
                    catch (Exception ex)
                    {
                        sw.WriteLine($@"C {uid}   {ex.Message}");
                    }

                }

            }
            //just try Put
            foreach (var uid in UID)
            {
                DataTable dtMatrix = new DataTable();
                DataRow[] select = vprice.Select($@"CommodityID='{uid}'");
                if (select.Length == 0)
                    continue;
                //取得履約價中點
                double get_quote;
                if (Convert.ToDouble(select[0][1].ToString()) == 0)
                {
                    if (Convert.ToDouble(select[0][3].ToString()) == 0)
                    {
                        if (Convert.ToDouble(select[0][4].ToString()) == 0)
                            get_quote = 0;
                        else
                            get_quote = roundprice(Convert.ToDouble(select[0][4].ToString()));
                    }
                    else
                    {
                        get_quote = roundprice(Convert.ToDouble(select[0][3].ToString()));
                    }

                }
                else
                {
                    get_quote = roundprice(Convert.ToDouble(select[0][1].ToString()));
                }

                DataRow[] select2 = Pmatrix.Select($@"UID='{uid}'");
                using (StreamWriter sw = new StreamWriter("D:\\density.txt", true))
                {

                    try
                    {
                        if (select2.Length > 0)
                        {

                            //履約價間隔
                            double intervalPercent = Convert.ToDouble(select2[0][3].ToString());
                            if (intervalPercent <= 0)
                                continue;
                            double interval = roundpriceK(get_quote, get_quote * intervalPercent / 100);
                            //Put價內外要顛倒
                            //價內監控區間
                            //int num_itm = Convert.ToInt32(select2[0][6].ToString());
                            int num_itm = Convert.ToInt32(select2[0][7].ToString());
                            //價外監控區間
                            //int num_otm = Convert.ToInt32(select2[0][7].ToString());
                            int num_otm = Convert.ToInt32(select2[0][6].ToString());
                            //距到期日起始月
                            double maturity_start = Convert.ToDouble(select2[0][4].ToString());
                            //到期日間距
                            double maturity_interval = Convert.ToDouble(select2[0][5].ToString());
                            //期間監控區
                            int num_maturity = Convert.ToInt32(select2[0][8].ToString());
                            DataRow[] warrantbasicselect = warrantbasic.Select($@"UID='{uid}'");

                            if (warrantbasicselect.Length > 0)
                            {

                                //建表

                                for (int i = num_itm; i > 0; i--)
                                    dtMatrix.Columns.Add($@"{(get_quote - i * interval).ToString()}~{(get_quote - (i - 1) * interval).ToString()}", typeof(int));
                                for (int i = 1; i <= num_otm; i++)
                                    dtMatrix.Columns.Add($@"{(get_quote + (i - 1) * interval).ToString()}~{(get_quote + i * interval).ToString()}", typeof(int));
                                //double max = get_quote + interval * num_otm;
                                //dtMatrix.Columns.Add($@">{max.ToString()}", typeof(int));

                                for (int i = 1; i <= num_maturity; i++)
                                {
                                    DataRow dr = dtMatrix.NewRow();
                                    /*
                                    string t1 = DateTime.Today.AddDays((int)Math.Round((maturity_start + maturity_interval * (i - 1)) * 30)).ToString("yyyyMMdd");
                                    DateTime t11 = DateTime.Today.AddDays((int)Math.Round((maturity_start + maturity_interval * (i - 1)) * 30));
                                    string t2 = DateTime.Today.AddDays((int)Math.Round((maturity_start + maturity_interval * i) * 30)).ToString("yyyyMMdd");
                                    DateTime t22 = DateTime.Today.AddDays((int)Math.Round((maturity_start + maturity_interval * i) * 30));
                                    */
                                    DateTime ListedDat = EDLib.TradeDate.NextNTradeDate(2);
                                    DateTime t1 = ListedDat.AddMonths((int)(maturity_start + maturity_interval * (i - 1))).AddDays(-1);
                                    DateTime t2 = ListedDat.AddMonths((int)(maturity_start + maturity_interval * i)).AddDays(-1);
                                    string t1str = t1.ToString("yyyy-MM-dd HH:mm:ss.fff");
                                    string t2str = t2.ToString("yyyy-MM-dd HH:mm:ss.fff");
                                    int w = 0;
                                    for (int j = num_itm; j > 0; j--)
                                    {
                                        DataRow[] temp = warrantbasic.Select($@"[UID]='{uid}' AND [WClass]='p' AND [MaturityDate] > '{t1}' AND [MaturityDate] <='{t2}' AND [StrikePrice] >{get_quote - j * interval} AND [StrikePrice] <={get_quote - (j - 1) * interval}");
                                        dr[w] = temp.Length;
                                        w++;
                                    }

                                    for (int j = 1; j <= num_otm; j++)
                                    {
                                        DataRow[] temp = warrantbasic.Select($@"[UID]='{uid}' AND [WClass]='p' AND [MaturityDate] > '{t1}' AND [MaturityDate] <= '{t2}' AND [StrikePrice] >{get_quote + (j - 1) * interval} AND [StrikePrice] <={get_quote + (j) * interval}");
                                        dr[w] = temp.Length;
                                        w++;
                                    }
                                    dtMatrix.Rows.Add(dr);

                                }
                            }

                            int density = 0;
                            if (dtMatrix.Rows.Count > 0)
                            {
                                for (int i = 0; i < num_maturity; i++)
                                {
                                    for (int j = 0; j < (num_itm + num_otm); j++)
                                    {
                                        /*
                                        if (Convert.ToInt16(dtMatrix.Rows[i][j].ToString()) == 0)
                                            density++;
                                        */
                                        if (Convert.ToInt16(dtMatrix.Rows[i][j].ToString()) > 0)
                                            density++;
                                    }

                                }
                            }
                            /*
                            else
                            {
                                density = (num_itm + num_otm) * num_maturity;
                            }
                            */
#if !To39
                            string sqlInsert = $@"UPDATE [newEDIS].[dbo].[OptionAutoSelectData]
                                          SET [YuanPutDensity] = {density}
                                          WHERE [UID] ='{uid}'";
                            MSSQL.ExecSqlCmd(sqlInsert, GlobalVar.loginSet.newEDIS);
#else
                            string sqlInsert = $@"UPDATE [WarrantAssistant].[dbo].[OptionAutoSelectData]
                                          SET [YuanPutDensity] = {density}
                                          WHERE [UID] ='{uid}'";
                            MSSQL.ExecSqlCmd(sqlInsert, GlobalVar.loginSet.warrantassistant45);
#endif
                        }
                        else
                        {
#if !To39
                            string sqlInsert = $@"UPDATE [newEDIS].[dbo].[OptionAutoSelectData]
                                          SET [YuanPutDensity] = -1
                                          WHERE [UID] ='{uid}'";
                            MSSQL.ExecSqlCmd(sqlInsert, GlobalVar.loginSet.newEDIS);
#else
                            string sqlInsert = $@"UPDATE [WarrantAssistant].[dbo].[OptionAutoSelectData]
                                          SET [YuanPutDensity] = -1
                                          WHERE [UID] ='{uid}'";
                            MSSQL.ExecSqlCmd(sqlInsert, GlobalVar.loginSet.warrantassistant45);
#endif
                        }

                    }
                    catch (Exception ex)
                    {
                        sw.WriteLine($@"P {uid}   {ex.Message}");
                    }
                }

            }

        }

        public static void ReCalKgiYuanCallPutDensity()
        {
#if !To39
            string sqlC = $@"SELECT [UID],[Checked]
              FROM [newEDIS].[dbo].[OptionAutoSelectMatrix]
              WHERE [CP] = 'C'";
            DataTable dtC = MSSQL.ExecSqlQry(sqlC, GlobalVar.loginSet.newEDIS);
            foreach (DataRow dr in dtC.Rows)
            {
                string uid = dr["UID"].ToString();
                string check = dr["Checked"].ToString();
                if (check == "1")
                {
                    string sqlUpdate = $@"UPDATE [newEDIS].[dbo].[OptionAutoSelectData]
                                            SET [KgiYuanCallDensity] = [CallDensity] - [YuanCallDensity]
                                            WHERE [UID] = '{uid}'";
                    MSSQL.ExecSqlCmd(sqlUpdate, GlobalVar.loginSet.newEDIS);
                }
                else
                {
                    string sqlUpdate = $@"UPDATE [newEDIS].[dbo].[OptionAutoSelectData]
                                            SET [KgiYuanCallDensity] = -50
                                            WHERE [UID] = '{uid}'";
                    MSSQL.ExecSqlCmd(sqlUpdate, GlobalVar.loginSet.newEDIS);
                }
            }

            string sqlP = $@"SELECT [UID],[Checked]
              FROM [newEDIS].[dbo].[OptionAutoSelectMatrix]
              WHERE [CP] = 'P'";
            DataTable dtP = MSSQL.ExecSqlQry(sqlP, GlobalVar.loginSet.newEDIS);
            foreach (DataRow dr in dtP.Rows)
            {
                string uid = dr["UID"].ToString();
                string check = dr["Checked"].ToString();
                if (check == "1")
                {
                    string sqlUpdate = $@"UPDATE [newEDIS].[dbo].[OptionAutoSelectData]
                                            SET [KgiYuanPutDensity] = [PutDensity] - [YuanPutDensity]
                                            WHERE [UID] = '{uid}'";
                    MSSQL.ExecSqlCmd(sqlUpdate, GlobalVar.loginSet.newEDIS);
                }
                else
                {
                    string sqlUpdate = $@"UPDATE [newEDIS].[dbo].[OptionAutoSelectData]
                                            SET [KgiYuanPutDensity] = -50
                                            WHERE [UID] = '{uid}'";
                    MSSQL.ExecSqlCmd(sqlUpdate, GlobalVar.loginSet.newEDIS);
                }
            }
#else
            string sqlC = $@"SELECT [UID],[Checked]
              FROM [WarrantAssistant].[dbo].[OptionAutoSelectMatrix]
              WHERE [CP] = 'C'";
            DataTable dtC = MSSQL.ExecSqlQry(sqlC, GlobalVar.loginSet.warrantassistant45);
            foreach (DataRow dr in dtC.Rows)
            {
                string uid = dr["UID"].ToString();
                string check = dr["Checked"].ToString();
                if (check == "1")
                {
                    string sqlUpdate = $@"UPDATE [WarrantAssistant].[dbo].[OptionAutoSelectData]
                                            SET [KgiYuanCallDensity] = [CallDensity] - [YuanCallDensity]
                                            WHERE [UID] = '{uid}'";
                    MSSQL.ExecSqlCmd(sqlUpdate, GlobalVar.loginSet.warrantassistant45);
                }
                else
                {
                    string sqlUpdate = $@"UPDATE [WarrantAssistant].[dbo].[OptionAutoSelectData]
                                            SET [KgiYuanCallDensity] = -50
                                            WHERE [UID] = '{uid}'";
                    MSSQL.ExecSqlCmd(sqlUpdate, GlobalVar.loginSet.warrantassistant45);
                }
            }

            string sqlP = $@"SELECT [UID],[Checked]
              FROM [WarrantAssistant].[dbo].[OptionAutoSelectMatrix]
              WHERE [CP] = 'P'";
            DataTable dtP = MSSQL.ExecSqlQry(sqlP, GlobalVar.loginSet.warrantassistant45);
            foreach (DataRow dr in dtP.Rows)
            {
                string uid = dr["UID"].ToString();
                string check = dr["Checked"].ToString();
                if (check == "1")
                {
                    string sqlUpdate = $@"UPDATE [WarrantAssistant].[dbo].[OptionAutoSelectData]
                                            SET [KgiYuanPutDensity] = [PutDensity] - [YuanPutDensity]
                                            WHERE [UID] = '{uid}'";
                    MSSQL.ExecSqlCmd(sqlUpdate, GlobalVar.loginSet.warrantassistant45);
                }
                else
                {
                    string sqlUpdate = $@"UPDATE [WarrantAssistant].[dbo].[OptionAutoSelectData]
                                            SET [KgiYuanPutDensity] = -50
                                            WHERE [UID] = '{uid}'";
                    MSSQL.ExecSqlCmd(sqlUpdate, GlobalVar.loginSet.warrantassistant45);
                }
            }
#endif
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

            int rowindex = -1;
            int mode = 0;
            int head = 0;
            int tail = 0;


            for (int i = 0; i < rownum; i++)
            {
                for (int j = 0; j < colnum; j++)
                {
                    if (ultraGrid1.Rows[i].Cells[j].Selected == true)
                    {
                        //MessageBox.Show($@"{ultraGrid1.Rows[i].Cells[j].Value.ToString()} mode {mode}");
                        if (rowindex < 0)
                            rowindex = j;
                        else
                        {
                            if (rowindex != j)
                            {
                                MessageBox.Show($@"不能設定多欄");
                                return;
                            }
                        }
                        if (mode == 0)
                        {
                            mode = 1;
                            head = i;
                        }
                    }
                    else
                    {
                        if (mode == 1 && rowindex >= 0 && rowindex == j)
                        {
                            mode = 0;
                            if (i >= head && i > 0)
                                tail = i - 1;
                        }
                    }
                }

            }
            //MessageBox.Show($@"head:{head}  tail:{tail}  rowindex:{rowindex}");
            var headvalue = ultraGrid1.Rows[head].Cells[rowindex].Value;
            var tailvalue = ultraGrid1.Rows[tail].Cells[rowindex].Value;
            if (rowindex <= 1)//名稱跟代號不能改
                return;
            if (headvalue.ToString() == tailvalue.ToString())
            {
                //MessageBox.Show("equal");
                for (int i = head; i <= tail; i++)
                {
                    if (rowindex >= 0)
                    {
                        //MessageBox.Show($@"{ultraGrid1.Rows[i].Cells[rowindex].Value}");
                        ultraGrid1.Rows[i].Cells[rowindex].Value = headvalue;
                    }
                }
            }
            else
            {
                MessageBox.Show("設定參數不同!");
            }
        }

        private void ultraGrid2_InitializeLayout(object sender, InitializeLayoutEventArgs e)
        {
            e.Layout.ScrollBounds = ScrollBounds.ScrollToFill;
        }

        private void ultraGrid3_InitializeLayout(object sender, InitializeLayoutEventArgs e)
        {
            e.Layout.ScrollBounds = ScrollBounds.ScrollToFill;
        }
    }
}

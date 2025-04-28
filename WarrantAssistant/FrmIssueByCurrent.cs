#define To39
using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.Windows.Forms;
using Infragistics.Win.UltraWinGrid;
using System.Data.SqlClient;
using EDLib.SQL;

namespace WarrantAssistant
{
    public partial class FrmIssueByCurrent:Form
    {
        private DataTable dt = new DataTable();
        string userID = GlobalVar.globalParameter.userID;
        string sql = "";
        string trader = "";
        string underlyingID = "";

        public FrmIssueByCurrent() {
            InitializeComponent();
        }

        private void FrmIssueByCurrent_Load(object sender, EventArgs e) {
            InitialGrid();
            foreach (var item in GlobalVar.globalParameter.traders)
                toolStripComboBox1.Items.Add(item.TrimStart('0'));                       
                       
            toolStripComboBox1.Text = userID.TrimStart('0');
            //LoadTraderSql();
            //LoadData();
        }
        private void InitialGrid() {
            dt.Columns.Add("標的代號", typeof(string));
            dt.Columns.Add("標的名稱", typeof(string));
            dt.Columns.Add("權證代號", typeof(string));
            dt.Columns.Add("今日理論價", typeof(double));
            dt.Columns.Add("Delta", typeof(double));
            dt.Columns.Add("型態", typeof(string));
            dt.Columns.Add("CP", typeof(string));
            dt.Columns.Add("S", typeof(double));
            dt.Columns.Add("K", typeof(double));
            dt.Columns.Add("T", typeof(int));
            dt.Columns.Add("行使比例", typeof(double));
            dt.Columns.Add("HV", typeof(double));
            dt.Columns.Add("IV", typeof(double));
            dt.Columns.Add("Initial_IV", typeof(double));
            dt.Columns.Add("重設比", typeof(double));
            dt.Columns.Add("界限比", typeof(double));
            dt.Columns.Add("財務費用", typeof(double));
            dt.Columns.Add("發行日", typeof(DateTime));
            dt.Columns.Add("到期日", typeof(DateTime));
            dt.Columns.Add("交易員", typeof(string));
            dt.Columns.Add("說明", typeof(string));

            ultraGrid1.DataSource = dt;

            ultraGrid1.DisplayLayout.Bands[0].Columns["交易員"].Hidden = true;

            ultraGrid1.DisplayLayout.Bands[0].Columns["標的名稱"].Width = 130;
            ultraGrid1.DisplayLayout.Bands[0].Columns["型態"].Width = 80;
            ultraGrid1.DisplayLayout.Bands[0].Columns["CP"].Width = 70;
            ultraGrid1.DisplayLayout.Bands[0].Columns["發行日"].Width = 120;
            ultraGrid1.DisplayLayout.Bands[0].Columns["到期日"].Width = 120;
            ultraGrid1.DisplayLayout.AutoFitStyle = AutoFitStyle.ResizeAllColumns;

            ultraGrid1.DisplayLayout.Bands[0].Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
            ultraGrid1.DisplayLayout.Bands[0].Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
            ultraGrid1.DisplayLayout.Bands[0].Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.False;
            ultraGrid1.DisplayLayout.Bands[0].Columns[0].CellActivation = Activation.NoEdit;
            ultraGrid1.DisplayLayout.Bands[0].Columns[1].CellActivation = Activation.NoEdit;
            ultraGrid1.DisplayLayout.Bands[0].Columns[2].CellActivation = Activation.NoEdit;
            ultraGrid1.DisplayLayout.Bands[0].Columns[3].CellActivation = Activation.NoEdit;
            ultraGrid1.DisplayLayout.Bands[0].Columns[4].CellActivation = Activation.NoEdit;
            ultraGrid1.DisplayLayout.Bands[0].Columns[5].CellActivation = Activation.NoEdit;
            ultraGrid1.DisplayLayout.Bands[0].Columns[6].CellActivation = Activation.NoEdit;
            ultraGrid1.DisplayLayout.Bands[0].Columns[7].CellActivation = Activation.NoEdit;
            ultraGrid1.DisplayLayout.Bands[0].Columns[8].CellActivation = Activation.NoEdit;
            ultraGrid1.DisplayLayout.Bands[0].Columns[9].CellActivation = Activation.NoEdit;
            ultraGrid1.DisplayLayout.Bands[0].Columns[10].CellActivation = Activation.NoEdit;
            ultraGrid1.DisplayLayout.Bands[0].Columns[11].CellActivation = Activation.NoEdit;
            ultraGrid1.DisplayLayout.Bands[0].Columns[12].CellActivation = Activation.NoEdit;
            ultraGrid1.DisplayLayout.Bands[0].Columns[13].CellActivation = Activation.NoEdit;
            ultraGrid1.DisplayLayout.Bands[0].Columns[14].CellActivation = Activation.NoEdit;
            ultraGrid1.DisplayLayout.Bands[0].Columns[15].CellActivation = Activation.NoEdit;
            ultraGrid1.DisplayLayout.Bands[0].Columns[16].CellActivation = Activation.NoEdit;
            ultraGrid1.DisplayLayout.Bands[0].Columns[17].CellActivation = Activation.NoEdit;
            ultraGrid1.DisplayLayout.Bands[0].Columns[18].CellActivation = Activation.NoEdit;
            ultraGrid1.DisplayLayout.Bands[0].Columns[19].CellActivation = Activation.NoEdit;
            ultraGrid1.DisplayLayout.Bands[0].Columns[20].CellActivation = Activation.NoEdit;

            ultraGrid1.DisplayLayout.Bands[0].Columns["發行日"].Format = "yyyy/MM/dd";
            ultraGrid1.DisplayLayout.Bands[0].Columns["到期日"].Format = "yyyy/MM/dd";
            //ultraGrid1.DisplayLayout.Bands[0].Columns["明日上市"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center;
            //ultraGrid1.DisplayLayout.Bands[0].Override.HeaderAppearance.TextHAlign = Infragistics.Win.HAlign.Left;
            ultraGrid1.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti;
        }

        private void LoadTraderSql() {
            trader = toolStripComboBox1.Text;

            sql = @"SELECT a.[UnderlyingID]
                                  ,c.[UnderlyingName]
                                  ,a.[WarrantID]
                                  ,a.[WarrantType]
                                  ,IsNull(b.[MPrice],0) MPrice
                                  ,a.[K]
                                  ,a.[T]
                                  ,a.[exeRatio]
                                  ,a.[HV]
                                  ,a.[IV]
                                  ,a.[ResetR]
                                  ,a.[BarrierR]
                                  ,a.[IssuePrice]
                                  ,a.[IssueDate]
                                  ,a.[ExpiryDate]
                                  ,c.[TraderID] TraderID
                                  ,ISNULL(d.InitialIV,0) InitialIV
                                  ,ISNULL(d.說明,'') 說明
                                  ,ISNULL(d.[財務費用率],0) AS 財務費用率
                              FROM [WarrantAssistant].[dbo].[WarrantBasic] a
                              LEFT JOIN [WarrantAssistant].[dbo].[WarrantPrices] b ON a.UnderlyingID=b.CommodityID
                              LEFT JOIN [WarrantAssistant].[dbo].[WarrantUnderlyingSummary] c ON a.UnderlyingID=c.UnderlyingID
                              LEFT JOIN [WarrantAssistant].[dbo].[WarrantBasic_InitialIV] d ON a.WarrantName =d.WarrantName";
            sql += " WHERE (c.[TraderID])='" + trader + "' ORDER BY a.UnderlyingID, a.[WarrantType] desc, a.IssueDate desc";

        }

        private void LoadUnderSql() {
            underlyingID = toolStripTextBox1.Text;

            sql = @"SELECT a.[UnderlyingID]
                                  ,c.[UnderlyingName]
                                  ,a.[WarrantID]
                                  ,a.[WarrantType]
                                  ,IsNull(b.[MPrice],0) MPrice
                                  ,a.[K]
                                  ,a.[T]
                                  ,a.[exeRatio]
                                  ,a.[HV]
                                  ,a.[IV]
                                  ,a.[ResetR]
                                  ,a.[BarrierR]
                                  ,a.[IssuePrice]
                                  ,a.[IssueDate]
                                  ,a.[ExpiryDate]
                                  ,c.[TraderID] TraderID
                                  ,ISNULL(d.InitialIV,0) InitialIV
                                  ,ISNULL(d.說明,'') 說明
                                  ,ISNULL(d.[財務費用率],0) AS 財務費用率
                              FROM [WarrantAssistant].[dbo].[WarrantBasic] a
                              LEFT JOIN [WarrantAssistant].[dbo].[WarrantPrices] b ON a.UnderlyingID=b.CommodityID
                              LEFT JOIN [WarrantAssistant].[dbo].[WarrantUnderlying] c ON a.UnderlyingID=c.UnderlyingID
                              LEFT JOIN [WarrantAssistant].[dbo].[WarrantBasic_InitialIV] d ON a.WarrantName =d.WarrantName";
            sql += " WHERE a.[UnderlyingID]='" + underlyingID + "' ORDER BY a.[WarrantType] desc, a.IssueDate desc";

        }

        private void SortTrader() {
            ultraGrid1.DisplayLayout.Bands[0].Columns["標的代號"].SortIndicator = SortIndicator.Ascending;
        }

        private void LoadData() {

            DataView dv = DeriLib.Util.ExecSqlQry(sql, GlobalVar.loginSet.warrantassistant45);


            if (dv.Count > 0) {
                dt.Rows.Clear();
                foreach (DataRowView drv in dv) {
                    DataRow dr = dt.NewRow();

                    string warrantType = drv["WarrantType"].ToString();
                    string type = "";
                    string cp = "";
                    CallPutType cpType = CallPutType.Call;
                    double financialR = 0.0;
                    financialR  = Convert.ToDouble(drv["財務費用率"]);

                    if (warrantType == "浮動重設認購權證") {
                        type = "重設型";
                        cp = "C";
                        cpType = CallPutType.Call;
                    } else if (warrantType == "浮動重設認售權證") {
                        type = "重設型";
                        cp = "P";
                        cpType = CallPutType.Put;
                    } else if (warrantType == "重設型牛證認購") {
                        type = "牛熊證";
                        cp = "C";
                        cpType = CallPutType.Call;
                        //financialR = 5.0;
                    } else if (warrantType == "重設型熊證認售") {
                        type = "牛熊證";
                        cp = "P";
                        cpType = CallPutType.Put;
                        //financialR = 5.0;
                    } else if (warrantType == "一般型認售權證") {
                        type = "一般型";
                        cp = "P";
                        cpType = CallPutType.Put;
                    } else {
                        type = "一般型";
                        cp = "C";
                        cpType = CallPutType.Call;
                    }

                    double s = Convert.ToDouble(drv["MPrice"]);
                    double k = Convert.ToDouble(drv["K"]);
                    int t = Convert.ToInt32(drv["T"]);
                    double cr = Convert.ToDouble(drv["exeRatio"]);
                    double hv = Convert.ToDouble(drv["HV"]);
                    double iv = Convert.ToDouble(drv["IV"]);
                    double initial_iv = Convert.ToDouble(drv["InitialIV"]);
                    double resetR = Convert.ToDouble(drv["ResetR"]);
                    double barrierR = Convert.ToDouble(drv["BarrierR"]);

                    DateTime issueDate = Convert.ToDateTime(drv["IssueDate"]);
                    string issueDateStr = issueDate.ToShortDateString();
                    DateTime expiry = Convert.ToDateTime(drv["ExpiryDate"]);
                    string expiryStr = expiry.ToShortDateString();
                    double price = 0.0;
                    double delta = 0.0;
                    string underlyingID = drv["UnderlyingID"].ToString();
                    if (s != 0.0) {
                        if (underlyingID.Length > 4 && underlyingID.Substring(0, 2) != "00")
                        {
                            if (type == "牛熊證")
                                price = Pricing.BullBearWarrantPrice(cpType, s, (resetR / 100), GlobalVar.globalParameter.interestRate_Index, (iv / 100), t, (financialR / 100), cr);
                            else if (type == "重設型")
                                price = Pricing.ResetWarrantPrice(cpType, s, (resetR / 100), GlobalVar.globalParameter.interestRate_Index, (iv / 100), t, cr);
                            else
                                price = Pricing.NormalWarrantPrice(cpType, s, k, GlobalVar.globalParameter.interestRate_Index, (iv / 100), t, cr);
                        }
                        else
                        {
                            if (type == "牛熊證")
                                price = Pricing.BullBearWarrantPrice(cpType, s, (resetR / 100), GlobalVar.globalParameter.interestRate, (iv / 100), t, (financialR / 100), cr);
                            else if (type == "重設型")
                                price = Pricing.ResetWarrantPrice(cpType, s, (resetR / 100), GlobalVar.globalParameter.interestRate, (iv / 100), t, cr);
                            else
                                price = Pricing.NormalWarrantPrice(cpType, s, k, GlobalVar.globalParameter.interestRate, (iv / 100), t, cr);
                        }
                        if (warrantType == "牛熊證")
                            delta = 1.0;
                        else
                        {
                            if (underlyingID.Length > 4 && underlyingID.Substring(0, 2) != "00")
                                delta = Pricing.Delta(cpType, s, k, GlobalVar.globalParameter.interestRate_Index, (iv / 100), (t * 30.0) / GlobalVar.globalParameter.dayPerYear, GlobalVar.globalParameter.interestRate_Index) * cr;
                            else
                                delta = Pricing.Delta(cpType, s, k, GlobalVar.globalParameter.interestRate, (iv / 100), (t * 30.0) / GlobalVar.globalParameter.dayPerYear, GlobalVar.globalParameter.interestRate) * cr;
                        }
                    }

                    dr["標的代號"] = drv["UnderlyingID"].ToString();
                    dr["標的名稱"] = drv["UnderlyingName"].ToString();
                    dr["權證代號"] = drv["WarrantID"].ToString();
                    dr["型態"] = type;
                    dr["CP"] = cp;
                    dr["S"] = s;
                    dr["K"] = k;
                    dr["T"] = t;
                    dr["行使比例"] = cr;
                    dr["HV"] = hv;
                    dr["IV"] = iv;
                    dr["Initial_IV"] = initial_iv;
                    dr["重設比"] = resetR;
                    dr["界限比"] = barrierR;
                    dr["財務費用"] = financialR;
                    dr["發行日"] = issueDate;
                    dr["到期日"] = expiry;
                    dr["今日理論價"] = Math.Round(price, 2);
                    dr["Delta"] = Math.Round(delta, 4);
                    dr["交易員"] = drv["TraderID"].ToString().PadLeft(7, '0');
                    dr["說明"] = drv["說明"].ToString();

                    dt.Rows.Add(dr);
                }

            }

        }

        private void toolStripComboBox1_KeyDown(object sender, KeyEventArgs e) {
            if (e.KeyCode == Keys.Enter && toolStripComboBox1.Text != "") {
                LoadTraderSql();
                LoadData();
                //toolStripComboBox1.Text = "";
            }
        }

        private void toolStripComboBox1_SelectedIndexChanged(object sender, EventArgs e) {
            LoadTraderSql();
            LoadData();
        }

        private void toolStripTextBox1_KeyDown(object sender, KeyEventArgs e) {
            if (e.KeyCode == Keys.Enter && toolStripTextBox1.Text != "") {
                LoadUnderSql();
                LoadData();
                toolStripTextBox1.Text = "";
            }
        }

        private void ultraGrid1_InitializeRow(object sender, InitializeRowEventArgs e) {
            DateTime expiry;
            string expiryStr = e.Row.Cells["到期日"].Value.ToString();
            expiry = Convert.ToDateTime(expiryStr);
            if (expiry <= DateTime.Today)
                e.Row.Cells["到期日"].Appearance.BackColor = Color.LightYellow;

            double price = 0.0;
            price = Convert.ToDouble(e.Row.Cells["今日理論價"].Value);
            if (price < 0.6 || price > 3)
                e.Row.Cells["今日理論價"].Appearance.ForeColor = Color.Red;
        }

        private void ultraGrid1_MouseClick(object sender, MouseEventArgs e) {
            
            if (e.Button == MouseButtons.Right) {
                contextMenuStrip1.Show();
            }
        }

        private void 加到申請編輯表ToolStripMenuItem_Click(object sender, EventArgs e) {
            //string sqlTemp = "SELECT * FROM [EDIS].[dbo].[ApplyTempList] WHERE UserID='" + userID + "'";
            //DataView dvTemp = DeriLib.Util.ExecSqlQry(sqlTemp, GlobalVar.loginSet.edisSqlConnString);
            //int count = dvTemp.Count;
            ///////////////

            string sqlTemp = $@"SELECT  [SerialNum]
                              FROM [WarrantAssistant].[dbo].[ApplyTempList]
                              WHERE [MDate] >='{DateTime.Today.ToString("yyyyMMdd")}' AND [TraderID] = '{userID}'";
            DataTable dtTemp = MSSQL.ExecSqlQry(sqlTemp, GlobalVar.loginSet.warrantassistant45);

            int ApplyTempListMax = 0;
            int TempListDeleteLogMax = 0;
            foreach (DataRow dr in dtTemp.Rows)
            {
                string serial = dr["SerialNum"].ToString();
                int seriallength = serial.Length;
                int temp = Convert.ToInt32(serial.Substring(seriallength - 2, 2));
                ApplyTempListMax = temp > ApplyTempListMax ? temp : ApplyTempListMax;
            }

            string sqlTempdel = $@"SELECT [SerialNum]
                          FROM [WarrantAssistant].[dbo].[TempListDeleteLog]
                          WHERE [DateTime] >='{DateTime.Today.ToString("yyyyMMdd")}' and [Trader] ='{userID}'";
            DataTable dtTempdel = MSSQL.ExecSqlQry(sqlTempdel, GlobalVar.loginSet.warrantassistant45);

            foreach (DataRow dr in dtTempdel.Rows)
            {
                string serial = dr["SerialNum"].ToString();
                int seriallength = serial.Length;
                int temp = Convert.ToInt32(serial.Substring(seriallength - 2, 2));
                TempListDeleteLogMax = temp > TempListDeleteLogMax ? temp : TempListDeleteLogMax;
            }
            int count = ApplyTempListMax >= TempListDeleteLogMax ? ApplyTempListMax : TempListDeleteLogMax;
            ///////////////
            string serialNum = DateTime.Today.ToString("yyyyMMdd") + userID + "01" + (count + 1).ToString("0#");
            string underlyingID = ultraGrid1.ActiveRow.Cells["標的代號"].Value.ToString();
            string underlyingName = ultraGrid1.ActiveRow.Cells["標的名稱"].Value.ToString();
            string WID = ultraGrid1.ActiveRow.Cells["權證代號"].Value.ToString();
            double k = Convert.ToDouble(ultraGrid1.ActiveRow.Cells["K"].Value);
            int t = Convert.ToInt32(ultraGrid1.ActiveRow.Cells["T"].Value);
            double cr = Convert.ToDouble(ultraGrid1.ActiveRow.Cells["行使比例"].Value);
            double hv = Convert.ToDouble(ultraGrid1.ActiveRow.Cells["HV"].Value);
            double iv = Convert.ToDouble(ultraGrid1.ActiveRow.Cells["IV"].Value);
            double issueNum = 10000.0;
            double resetR = Convert.ToDouble(ultraGrid1.ActiveRow.Cells["重設比"].Value);
            double barrierR = Convert.ToDouble(ultraGrid1.ActiveRow.Cells["界限比"].Value);
            double financialR = Convert.ToDouble(ultraGrid1.ActiveRow.Cells["財務費用"].Value);
            string type = ultraGrid1.ActiveRow.Cells["型態"].Value.ToString();
            string cp = ultraGrid1.ActiveRow.Cells["CP"].Value.ToString();
            string useReward = "N";
            string confirm = "N";
            string is1500W = "N";
            string tempName = "";
            string traderID = ultraGrid1.ActiveRow.Cells["交易員"].Value.ToString().PadLeft(7, '0');

            DateTime expiryDate;
            expiryDate = GlobalVar.globalParameter.nextTradeDate3.AddMonths(t);
            expiryDate = expiryDate.AddDays(-1);
            string sqlTemp2 = "SELECT TOP 1 TradeDate from TradeDate WHERE IsTrade='Y' AND TradeDate >= '" + expiryDate.ToString("yyyy-MM-dd") + "'";
            DataView dvTemp2 = DeriLib.Util.ExecSqlQry(sqlTemp2, GlobalVar.loginSet.tsquoteSqlConnString);
            foreach (DataRowView drTemp in dvTemp2) {
                expiryDate = Convert.ToDateTime(drTemp["TradeDate"]);
            }
            string expiryMonth = "";
            int month = expiryDate.Month;
            if (month >= 10) {
                if (month == 10)
                    expiryMonth = "A";
                if (month == 11)
                    expiryMonth = "B";
                if (month == 12)
                    expiryMonth = "C";
            } else
                expiryMonth = month.ToString();

            string expiryYear = "";
            expiryYear = expiryDate.AddYears(-1).ToString("yyyy");
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

            tempName = underlyingName + "凱基" + expiryYear + expiryMonth + warrantType;

            string sqlTemp3 = @"INSERT INTO [ApplyTempList] (SerialNum, UnderlyingID, K, T, R, HV, IV, IssueNum, ResetR, BarrierR, FinancialR, Type, CP, UseReward, ConfirmChecked, Apply1500W, UserID, MDate, TempName, TempType, TraderID) ";
            sqlTemp3 += "VALUES(@SerialNum, @UnderlyingID, @K, @T, @R, @HV, @IV, @IssueNum, @ResetR, @BarrierR, @FinancialR, @Type, @CP, @UseReward, @ConfirmChecked, @Apply1500W, @UserID, @MDate, @TempName ,@TempType, @TraderID)";
            List<SqlParameter> ps = new List<SqlParameter>();
            ps.Add(new SqlParameter("@SerialNum", SqlDbType.VarChar));
            ps.Add(new SqlParameter("@UnderlyingID", SqlDbType.VarChar));
            ps.Add(new SqlParameter("@K", SqlDbType.Float));
            ps.Add(new SqlParameter("@T", SqlDbType.Int));
            ps.Add(new SqlParameter("@R", SqlDbType.Float));
            ps.Add(new SqlParameter("@HV", SqlDbType.Float));
            ps.Add(new SqlParameter("@IV", SqlDbType.Float));
            ps.Add(new SqlParameter("@IssueNum", SqlDbType.Float));
            ps.Add(new SqlParameter("@ResetR", SqlDbType.Float));
            ps.Add(new SqlParameter("@BarrierR", SqlDbType.Float));
            ps.Add(new SqlParameter("@FinancialR", SqlDbType.Float));
            ps.Add(new SqlParameter("@Type", SqlDbType.VarChar));
            ps.Add(new SqlParameter("@CP", SqlDbType.VarChar));
            ps.Add(new SqlParameter("@UseReward", SqlDbType.VarChar));
            ps.Add(new SqlParameter("@ConfirmChecked", SqlDbType.VarChar));
            ps.Add(new SqlParameter("@Apply1500W", SqlDbType.VarChar));
            ps.Add(new SqlParameter("@UserID", SqlDbType.VarChar));
            ps.Add(new SqlParameter("@MDate", SqlDbType.DateTime));
            ps.Add(new SqlParameter("@TempName", SqlDbType.VarChar));
            ps.Add(new SqlParameter("@TempType", SqlDbType.VarChar));
            ps.Add(new SqlParameter("@TraderID", SqlDbType.VarChar));

            SQLCommandHelper h = new SQLCommandHelper(GlobalVar.loginSet.warrantassistant45, sqlTemp3, ps);

            h.SetParameterValue("@SerialNum", serialNum);
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
            h.SetParameterValue("@ConfirmChecked", confirm);
            h.SetParameterValue("@Apply1500W", is1500W);
            h.SetParameterValue("@UserID", userID);
            h.SetParameterValue("@MDate", DateTime.Now);
            h.SetParameterValue("@TempName", tempName);
            h.SetParameterValue("@TempType", tempType);
            h.SetParameterValue("@TraderID", traderID);
           
            h.ExecuteCommand();
            
           
            h.Dispose();

            GlobalUtility.LogInfo("Log", GlobalVar.globalParameter.userID + " " + WID + " 右鍵新增 " + underlyingID + underlyingName + " 一檔權證");

            MessageBox.Show("Done!");
        }


    }
}

using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.Windows.Forms;
using System.Data.SqlClient;
using Infragistics.Win.UltraWinGrid;
using HtmlAgilityPack;
using EDLib.SQL;
using System.Threading.Tasks;
using System.Net;

namespace WarrantAssistant
{
    public partial class Frm71:Form
    {

        public SqlConnection conn = new SqlConnection(GlobalVar.loginSet.warrantassistant45);

        private DataTable dt;// = new DataTable();
        private bool isEdit = false;

        public Frm71() {
            InitializeComponent();
        }

        private void Frm71_Load(object sender, EventArgs e) {
            //toolStripLabel1.Text = "";
            LoadData();
            InitialGrid();
        }

        private void InitialGrid() {
            //dt.PrimaryKey = new DataColumn[] { dt.Columns["權證名稱"] };
            //ultraGrid1.DataSource = dt;

            UltraGridBand bands0 = ultraGrid1.DisplayLayout.Bands[0];
            bands0.Columns["IssueNum"].Format = "N0";
            bands0.Columns["AvailableShares"].Format = "N0";
            bands0.Columns["LastDayUsedShares"].Format = "N0";
            bands0.Columns["TodayApplyShares"].Format = "N0";
            bands0.Columns["Issuer"].Width = 80;
            bands0.Columns["WarrantName"].Width = 130;
            bands0.Columns["UnderlyingID"].Width = 80;
            bands0.Columns["IssueNum"].Width = 80;
            bands0.Columns["exeRatio"].Width = 80;
            bands0.Columns["ApplyTime"].Width = 110;
            bands0.Columns["AvailableShares"].Width = 120;
            bands0.Columns["LastDayUsedShares"].Width = 120;
            bands0.Columns["TodayApplyShares"].Width = 120;
            bands0.Columns["AccUsedShares"].Width = 80;
            bands0.Columns["SameUnderlying"].Width = 80;
          
            //ultraGrid1.DisplayLayout.Bands[0].Columns["OriApplyTime"].Width = 110;
            ultraGrid1.DisplayLayout.AutoFitStyle = AutoFitStyle.ResizeAllColumns;

            bands0.Columns["AvailableShares"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right;
            bands0.Columns["LastDayUsedShares"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right;
            bands0.Columns["TodayApplyShares"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right;
            bands0.Columns["AccUsedShares"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center;

            bands0.Override.HeaderAppearance.TextHAlign = Infragistics.Win.HAlign.Left;
            bands0.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
            bands0.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
            bands0.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.False;

            SetButton();
        }

        private void SetButton() {
            if (isEdit) {
                ultraGrid1.DisplayLayout.Bands[0].Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.Yes;
                ultraGrid1.DisplayLayout.Bands[0].Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
                ultraGrid1.DisplayLayout.Bands[0].Override.AllowDelete = Infragistics.Win.DefaultableBoolean.True;
            } else {
                ultraGrid1.DisplayLayout.Bands[0].Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
                ultraGrid1.DisplayLayout.Bands[0].Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.False;
                ultraGrid1.DisplayLayout.Bands[0].Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
            }
            toolStripButtonEdit.Visible = !isEdit;
            toolStripButtonConfirm.Visible = isEdit;
            toolStripButtonCancel.Visible = isEdit;
            Edit2.Visible = isEdit;
        }

        private void LoadData() {

            string sql = @"SELECT [Issuer]
                                  ,[WarrantName]
                                  ,[UnderlyingID]
                                  ,[IssueNum]
                                  ,[exeRatio]
                                  ,[ApplyTime]
                                  ,[AvailableShares]
                                  ,[LastDayUsedShares]
                                  ,[TodayApplyShares]
                                  ,[AccUsedShares]
                                  ,[SameUnderlying]
                                  ,[OriApplyTime]
                              FROM [WarrantAssistant].[dbo].[Apply_71]";

            dt = MSSQL.ExecSqlQry(sql, GlobalVar.loginSet.warrantassistant45);

            ultraGrid1.DataSource = dt;

            dt.Columns[0].Caption = "發行人";
            dt.Columns[1].Caption = "權證名稱";
            dt.Columns[2].Caption = "標的代號";
            dt.Columns[3].Caption = "發行張數";
            dt.Columns[4].Caption = "行使比例";
            dt.Columns[5].Caption = "申報時間";
            dt.Columns[6].Caption = "可發行股數";
            dt.Columns[7].Caption = "截至前一日";
            dt.Columns[8].Caption = "本日累積發行";
            dt.Columns[9].Caption = "累計%";
            dt.Columns[10].Caption = "同標的2檔";
            dt.Columns[11].Caption = "原始申報時間";
            //dt.Columns[12].Caption = "順位";
        }

        private void UpdateDB() {
            for (int x = ultraGrid1.Rows.Count - 1; x >= 0; x--) {
                try {
                    if (ultraGrid1.Rows[x].Cells[0].Value.ToString() == "")
                        ultraGrid1.Rows[x].Delete(false);

                } catch (Exception ex) {
                    MessageBox.Show(ex.Message);
                }
            }

            MSSQL.ExecSqlCmd("DELETE FROM [Apply_71]", conn);

            try {
                string sql = "INSERT INTO [Apply_71] values(@Issuer,@WarrantName,@UnderlyingID,@IssueNum,@exeRatio,@ApplyTime,@AvailableShares,@LastDayUsedShares,@TodayApplyShares,@AccUsedShares,@SameUnderlying,@OriApplyTime,@Result, @ApplyStatus, @ReIssueResult, @SerialNum)";
                List<SqlParameter> ps = new List<SqlParameter>();
                ps.Add(new SqlParameter("@Issuer", SqlDbType.VarChar));
                ps.Add(new SqlParameter("@WarrantName", SqlDbType.VarChar));
                ps.Add(new SqlParameter("@UnderlyingID", SqlDbType.VarChar));
                ps.Add(new SqlParameter("@IssueNum", SqlDbType.Float));
                ps.Add(new SqlParameter("@exeRatio", SqlDbType.Float));
                ps.Add(new SqlParameter("@ApplyTime", SqlDbType.VarChar));
                ps.Add(new SqlParameter("@AvailableShares", SqlDbType.Float));
                ps.Add(new SqlParameter("@LastDayUsedShares", SqlDbType.Float));
                ps.Add(new SqlParameter("@TodayApplyShares", SqlDbType.Float));
                ps.Add(new SqlParameter("@AccUsedShares", SqlDbType.VarChar));
                ps.Add(new SqlParameter("@SameUnderlying", SqlDbType.VarChar));
                ps.Add(new SqlParameter("@OriApplyTime", SqlDbType.VarChar));
                ps.Add(new SqlParameter("@Result", SqlDbType.Float));
                ps.Add(new SqlParameter("@ApplyStatus", SqlDbType.VarChar));
                ps.Add(new SqlParameter("@ReIssueResult", SqlDbType.Float));
                ps.Add(new SqlParameter("@SerialNum", SqlDbType.VarChar));

                SQLCommandHelper h = new SQLCommandHelper(GlobalVar.loginSet.warrantassistant45, sql, ps);

                foreach (UltraGridRow r in ultraGrid1.Rows) {
                    string issuer = r.Cells["Issuer"].Value.ToString();
                    string warrantName = r.Cells["WarrantName"].Value.ToString();
                    string underlyingID = r.Cells["UnderlyingID"].Value.ToString();
                    double issueNum = Convert.ToDouble(r.Cells["IssueNum"].Value);
                    double exeRatio = Convert.ToDouble(r.Cells["exeRatio"].Value);
                    string applyTime = r.Cells["ApplyTime"].Value.ToString();
                    double availableShares = 0.0;
                    double lastDayUsedShares = 0.0;
                    double todayApplyShares = 0.0;
                    //if (underlyingID == "IX0001") {
                    if (char.IsLetter(underlyingID[0])) {
                        availableShares = 0.0;
                        lastDayUsedShares = 0.0;
                        todayApplyShares = 0.0;
                    } else {
                        availableShares = Convert.ToDouble(r.Cells["AvailableShares"].Value);//可發行股數
                        lastDayUsedShares = Convert.ToDouble(r.Cells["LastDayUsedShares"].Value);//截至前一日
                        todayApplyShares = Convert.ToDouble(r.Cells["TodayApplyShares"].Value);//本日累積發行
                    }
                    string accUsedShares = r.Cells["AccUsedShares"].Value.ToString();
                    string sameUnderlying = r.Cells["SameUnderlying"].Value.ToString();
                    string oriApplyTime = r.Cells["OriApplyTime"].Value.ToString();

                    //string underlyingName = warrantName.Substring(1, warrantName.Length - 7);//需考慮以前的短權證名稱

                    double multiplier = 1.0;

                    string sqlTemp = "SELECT CASE WHEN [StockType]='DS' OR [StockType]='DR' THEN 0.22 ELSE 1 END AS Multiplier FROM [WarrantAssistant].[dbo].[WarrantUnderlying] WHERE UnderlyingID = '" + underlyingID + "'";
                    DataTable dv = MSSQL.ExecSqlQry(sqlTemp, GlobalVar.loginSet.warrantassistant45);

                    foreach (DataRow dr in dv.Rows)
                        multiplier = Convert.ToDouble(dr["Multiplier"]);

                    double todayAvailable = Math.Round(((availableShares * multiplier - lastDayUsedShares) / 1000), 1);
                    double attempShares = issueNum * exeRatio;
                    double result = 0.0;
                    double tempAvailable = 0.0;
                    string applyStatus = "";
                    tempAvailable = todayAvailable - todayApplyShares / 1000 + attempShares;

                    //if (underlyingID == "IX0001") {
                    if (char.IsLetter(underlyingID[0])) {
                        result = attempShares;
                        applyStatus = "Y";
                    } else if (applyTime.Substring(0, 2) == "09" || applyTime.Substring(0, 2) == "10") {
                        if (tempAvailable >= attempShares) {
                            result = attempShares;
                            applyStatus = "Y";
                        } else if (tempAvailable > 0) {
                            result = tempAvailable;
                            applyStatus = "排隊中";
                        } else {
                            result = 0;
                            if (todayAvailable >= 0.6 * attempShares)
                                applyStatus = "排隊中";
                            else
                                applyStatus = "X 沒額度";
                        }
                    }else if (applyTime.Substring(0, 2) == "11" || applyTime.Substring(0, 2) == "12"|| applyTime.Substring(0, 2) == "13" || applyTime.Substring(0, 2) == "14") {//搶額度結果出來後再增發
                        if (tempAvailable >= attempShares)//交易所開放上傳時間為1400前
                        {
                            result = attempShares;
                            applyStatus = "Y";
                        }
                        else if (tempAvailable > 0)
                        {
                            result = tempAvailable;
                            applyStatus = "排隊中";
                        }
                        else
                        {
                            result = 0;
                            if (todayAvailable >= 0.6 * attempShares)
                                applyStatus = "排隊中";
                            else
                                applyStatus = "X 沒額度";
                        }
                    }
                    else if (applyTime.Substring(0, 2) == "22") {
                        result = 0;
                        applyStatus = "X 沒額度";
                    }

                    double accUsed = (lastDayUsedShares + todayApplyShares) / availableShares;
                    double reIssueResult = 0.0;
                    if (accUsed <= 0.3 || underlyingID.StartsWith("00") || (char.IsLetter(underlyingID[0])))
                        reIssueResult = attempShares;

                    h.SetParameterValue("@Issuer", issuer);
                    h.SetParameterValue("@WarrantName", warrantName);
                    h.SetParameterValue("@UnderlyingID", underlyingID);
                    h.SetParameterValue("@IssueNum", issueNum);
                    h.SetParameterValue("@exeRatio", exeRatio);
                    h.SetParameterValue("@ApplyTime", applyTime);
                    h.SetParameterValue("@AvailableShares", availableShares);
                    h.SetParameterValue("@LastDayUsedShares", lastDayUsedShares);
                    h.SetParameterValue("@TodayApplyShares", todayApplyShares);
                    h.SetParameterValue("@AccUsedShares", accUsedShares);
                    h.SetParameterValue("@SameUnderlying", sameUnderlying);
                    h.SetParameterValue("@OriApplyTime", oriApplyTime);
                    h.SetParameterValue("@Result", result);
                    h.SetParameterValue("@ApplyStatus", applyStatus);
                    h.SetParameterValue("@ReIssueResult", reIssueResult);
                    h.SetParameterValue("@SerialNum", "0");

                    h.ExecuteCommand();
                }

                h.Dispose();

                string sql5 = "UPDATE [WarrantAssistant].[dbo].[Apply_71] SET SerialNum = B.SerialNum FROM [WarrantAssistant].[dbo].[ApplyTotalList] B WHERE [Apply_71].[WarrantName]=B.WarrantName";
                string sql2 = "UPDATE [WarrantAssistant].[dbo].[ApplyTotalList] SET Result=0";
                string sql3 = @"UPDATE [WarrantAssistant].[dbo].[ApplyTotalList] 
                                SET Result= CASE WHEN ApplyKind='1' THEN B.Result ELSE B.ReIssueResult END
                                FROM [WarrantAssistant].[dbo].[Apply_71] B
                                WHERE [ApplyTotalList].[WarrantName]=B.WarrantName";
                string sql4 = @"UPDATE [WarrantAssistant].[dbo].[ApplyTotalList]
                                SET Result= CASE WHEN [RewardCredit]>=[EquivalentNum] THEN [EquivalentNum] ELSE [RewardCredit] END
                               WHERE [UseReward]='Y'";

                conn.Open();
                MSSQL.ExecSqlCmd(sql5, conn);
                MSSQL.ExecSqlCmd(sql2, conn);
                MSSQL.ExecSqlCmd(sql3, conn);
                MSSQL.ExecSqlCmd(sql4, conn);
                conn.Close();

                toolStripLabel1.Text = DateTime.Now + "更新成功";

                GlobalUtility.LogInfo("Info", GlobalVar.globalParameter.userID + " 更新7-1試算表");


            } catch (Exception ex) {
                GlobalUtility.LogInfo("Exception", GlobalVar.globalParameter.userID + "7-1試算表" + ex.Message);

                MessageBox.Show(ex.Message);
            }
        }

        private async void toolStripButtonEdit_Click(object sender, EventArgs e) {
            toolStripLabel1.Text = "";
            toolStrip1.Enabled = false;
            dt.Rows.Clear();
            isEdit = true;
            SetButton();

            string key = null;
            //Get key and id  
            try {
                key = GlobalUtility.GetKey();
            } catch (Exception ex) {
                MessageBox.Show($"GetKey from BSSDB failed {ex}");
                return;
            }
            //員編
            string id = GlobalUtility.GetID();
            //parse TWSE and OTC 7-1 html
            if (!await ParseHtml($"https://siis.twse.com.tw/server-java/t150sa10?step=0&id=9200pd{id}&TYPEK=sii&key={key}", true)
                & !await ParseHtml($"https://siis.twse.com.tw/server-java/o_t150sa10?step=0&id=9200pd{id}&TYPEK=otc&key={key}", false))
                return;
            

            GlobalUtility.LogInfo("Info", GlobalVar.globalParameter.userID + " 下載7-1試算表");
            toolStrip1.Enabled = true;
        }

        private void toolStripButtonConfirm_Click(object sender, EventArgs e) {
            isEdit = false;
            UpdateDB();
            LoadData();
            SetButton();
        }

        private void toolStripButtonCancel_Click(object sender, EventArgs e) {
            isEdit = false;
            LoadData();
            SetButton();
        }

        private void ultraGrid1_InitializeRow(object sender, Infragistics.Win.UltraWinGrid.InitializeRowEventArgs e) {
            string applyTime = e.Row.Cells["ApplyTime"].Value == DBNull.Value ? "" : e.Row.Cells["ApplyTime"].Value.ToString();
            double issueNum = e.Row.Cells["IssueNum"].Value == DBNull.Value ? 0 : Convert.ToDouble(e.Row.Cells["IssueNum"].Value);
            string underlyingID = e.Row.Cells["UnderlyingID"].Value.ToString();


            if (applyTime.Length > 0) {
                applyTime = applyTime.Substring(0, 2);

                if (applyTime == "22")
                    e.Row.Cells["ApplyTime"].Appearance.ForeColor = Color.Red;

                if (applyTime == "09")
                    e.Row.Cells["ApplyTime"].Appearance.ForeColor = Color.Green;

                if (applyTime == "10" && issueNum != 10000)
                    e.Row.Cells["IssueNum"].Appearance.ForeColor = Color.Red;
            }
        }

        private void ultraGrid1_Error(object sender, Infragistics.Win.UltraWinGrid.ErrorEventArgs e) {
            if (isEdit)
                e.Cancel = true;
        }

        private async Task<bool> ParseHtml(string url, bool isTwse) {
            try {
                string firstResponse = await GlobalUtility.GetHtmlAsync(url, System.Text.Encoding.Default);//EDLib.Utility.GetHtml(url, System.Text.Encoding.Default);
                HtmlAgilityPack.HtmlDocument doc = new HtmlAgilityPack.HtmlDocument();
                doc.LoadHtml(firstResponse);
                HtmlNodeCollection navNodeChild = doc.DocumentNode.SelectSingleNode("//table[1]").ChildNodes; // /td[1]/table[1]/tr[2]

                int loopend = navNodeChild.Count;

                for (int i = 5; i < loopend; i += 2) {
                    //MessageBox.Show(navNodeChild[i].InnerText);

                    string[] split = navNodeChild[i].InnerText.Split(new string[] { "\n","\r" }, StringSplitOptions.RemoveEmptyEntries); //" ", "\t", "&nbsp;",
                    if (split.Length != 12)
                        continue;
                    DataRow dr = dt.NewRow();

                    dr["Issuer"] = split[0];
                    dr["WarrantName"] = split[1];
                    dr["UnderlyingID"] = split[2];
                    dr["IssueNum"] = split[3];
                    dr["exeRatio"] = split[4];
                    dr["ApplyTime"] = split[5];
                    if (double.TryParse(split[6], out double result)) {
                        dr["AvailableShares"] = split[6];
                        dr["LastDayUsedShares"] = split[7];
                        dr["TodayApplyShares"] = split[8];
                        dr["AccUsedShares"] = split[9];
                    } else {
                        dr["AvailableShares"] = 0;
                        dr["LastDayUsedShares"] = 0;
                        dr["TodayApplyShares"] = 0;
                    }

                    if (!split[10].StartsWith("&nbsp"))
                        dr["SameUnderlying"] = split[10];
                    if (!split[11].StartsWith("&nbsp"))
                        dr["OriApplyTime"] = split[11];

                    dt.Rows.Add(dr);
                }
                return true;
            } catch (WebException) {
                if (isTwse)
                    MessageBox.Show("可能要更新Key，或是網頁有問題");
                return false;
            } catch (NullReferenceException) {
                if (isTwse)
                    MessageBox.Show("TWSE 沒有資料");
                else
                    MessageBox.Show("OTC 沒有資料");
                return false;
            } catch (Exception e) {
                MessageBox.Show(e.ToString());
                return false;
            }
        }

        private void Edit2_Click(object sender, EventArgs e) {
            ultraGrid1.DisplayLayout.Bands[0].Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.Yes;
            ultraGrid1.DisplayLayout.Bands[0].Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
            ultraGrid1.DisplayLayout.Bands[0].Override.AllowDelete = Infragistics.Win.DefaultableBoolean.True;
            dt.Rows.Clear();
            for (int x = 0; x < 100; x++)
                ultraGrid1.DisplayLayout.Bands[0].AddNew();

            ultraGrid1.ActiveRowScrollRegion.ScrollRowIntoView(ultraGrid1.Rows[0]);
        }
    }
}

using System;
using System.Data;
using System.Drawing;
using System.Windows.Forms;
using Infragistics.Win.UltraWinGrid;
using EDLib.SQL;

namespace WarrantAssistant
{
    public partial class FrmReIssueTotal:Form
    {
        private DataTable dt;

        public FrmReIssueTotal() {
            InitializeComponent();
        }

        private void FrmReIssueTotal_Load(object sender, EventArgs e) {
            LoadData();
            InitialGrid();
        }

        private void InitialGrid() {           
            UltraGridBand band0 = ultraGrid1.DisplayLayout.Bands[0];
            band0.Columns["ReIssueNum"].Format = "N0";
            band0.Columns["EquivalentNum"].Format = "N0";
            band0.Columns["Result"].Format = "N0";
            band0.Columns["RewardIssueCredit"].Format = "N0";

            band0.Columns["TraderID"].Width = 70;
            band0.Columns["UnderlyingID"].Width = 70;
            band0.Columns["WarrantID"].Width = 70;
            band0.Columns["WarrantName"].Width = 150;
            band0.Columns["exeRatio"].Width = 70;
            band0.Columns["ReIssueNum"].Width = 80;
            band0.Columns["EquivalentNum"].Width = 80;
            band0.Columns["Result"].Width = 80;
            //ultraGrid1.DisplayLayout.Bands[0].Columns["IssuedPercent"].Width = 100;
            //ultraGrid1.DisplayLayout.Bands[0].Columns["RewardIssueCredit"].Width = 100;
            band0.Columns["UseReward"].Width = 70;
            band0.Columns["MarketTmr"].Width = 70;
            ultraGrid1.DisplayLayout.AutoFitStyle = AutoFitStyle.ResizeAllColumns;

            band0.Columns["ReIssueNum"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right;
            band0.Columns["EquivalentNum"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right;
            band0.Columns["Result"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right;
            band0.Columns["IssuedPercent"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right;
            band0.Columns["RewardIssueCredit"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right;
            band0.Columns["UseReward"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center;
            band0.Columns["MarketTmr"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center;
            band0.Override.HeaderAppearance.TextHAlign = Infragistics.Win.HAlign.Left;

            band0.Columns["TraderID"].CellActivation = Activation.NoEdit;
            band0.Columns["UnderlyingID"].CellActivation = Activation.NoEdit;
            band0.Columns["WarrantID"].CellActivation = Activation.NoEdit;
            band0.Columns["WarrantName"].CellActivation = Activation.NoEdit;
            band0.Columns["exeRatio"].CellActivation = Activation.NoEdit;
            band0.Columns["ReIssueNum"].CellActivation = Activation.NoEdit;
            band0.Columns["EquivalentNum"].CellActivation = Activation.NoEdit;
            band0.Columns["Result"].CellActivation = Activation.NoEdit;
            band0.Columns["IssuedPercent"].CellActivation = Activation.NoEdit;
            band0.Columns["RewardIssueCredit"].CellActivation = Activation.NoEdit;
            band0.Columns["UseReward"].CellActivation = Activation.NoEdit;
            band0.Columns["MarketTmr"].CellActivation = Activation.NoEdit;

            band0.Columns["SerialNum"].Hidden = true;

        }

        private void LoadData() {
            try {

                //dt.Rows.Clear();
                string sql = @"SELECT a.SerialNum
                                      ,a.TraderID
                                      ,a.UnderlyingID
                                      ,a.WarrantID
                                      ,a.WarrantName
                                      ,a.exeRatio
                                      ,a.ReIssueNum
                                      ,c.EquivalentNum
                                      ,c.Result
                                      ,IsNull(b.IssuedPercent,0) IssuedPercent
                                      ,IsNull(b.RewardIssueCredit,0) RewardIssueCredit
                                      ,a.UseReward
                                      ,a.MarketTmr
                                  FROM [WarrantAssistant].[dbo].[ReIssueOfficial] a
                                  LEFT JOIN [WarrantAssistant].[dbo].[WarrantUnderlyingSummary] b ON a.UnderlyingID=b.UnderlyingID
                                  LEFT JOIN [WarrantAssistant].[dbo].[ApplyTotalList] c ON a.SerialNum=c.SerialNum";

                dt = MSSQL.ExecSqlQry(sql, GlobalVar.loginSet.warrantassistant45);

                ultraGrid1.DataSource = dt;

                foreach (DataRow row in dt.Rows) {
                    row["IssuedPercent"] = Math.Round((double) row["IssuedPercent"], 2);
                    row["TraderID"] = row["TraderID"].ToString().TrimStart('0');
                }

                dt.Columns[0].Caption = "序號";
                dt.Columns[1].Caption = "交易員";
                dt.Columns[2].Caption = "標的代號";
                dt.Columns[3].Caption = "權證代號";
                dt.Columns[4].Caption = "權證名稱";
                dt.Columns[5].Caption = "行使比例";
                dt.Columns[6].Caption = "張數";
                dt.Columns[7].Caption = "約當張數";
                dt.Columns[8].Caption = "額度結果";
                dt.Columns[9].Caption = "今日額度(%)";
                dt.Columns[10].Caption = "獎勵額度";
                dt.Columns[11].Caption = "使用獎勵";
                dt.Columns[12].Caption = "明日上市";

                /*DataView dv = DeriLib.Util.ExecSqlQry(sql, GlobalVar.loginSet.edisSqlConnString);
                if (dv.Count > 0) {
                    foreach (DataRowView drv in dv) {
                        DataRow dr = dt.NewRow();

                        dr["序號"] = drv["SerialNum"].ToString();
                        dr["交易員"] = drv["TraderID"].ToString();
                        dr["標的代號"] = drv["UnderlyingID"].ToString();
                        dr["權證代號"] = drv["WarrantID"].ToString();
                        dr["權證名稱"] = drv["WarrantName"].ToString();
                        dr["行使比例"] = Convert.ToDouble(drv["exeRatio"]);
                        dr["張數"] = Convert.ToDouble(drv["ReIssueNum"]);
                        dr["約當張數"] = Convert.ToDouble(drv["EquivalentNum"]);
                        dr["額度結果"] = drv["Result"];
                        dr["今日額度(%)"] = Math.Round(Convert.ToDouble(drv["IssuedPercent"]), 2);
                        //double rewardCredit = (double) drv["RewardIssueCredit"];
                        //rewardCredit = Math.Floor((double) drv["RewardIssueCredit"]);
                        dr["獎勵額度"] = Math.Floor((double) drv["RewardIssueCredit"]);
                        dr["使用獎勵"] = drv["UseReward"].ToString();
                        dr["明日上市"] = drv["MarketTmr"].ToString();

                        dt.Rows.Add(dr);
                    }
                }*/
            } catch (Exception ex) {
                MessageBox.Show(ex.Message);
            }
        }

        private void toolStripButtonReload_Click(object sender, EventArgs e) {
            LoadData();
        }

        private void ultraGrid1_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e) {
            ultraGrid1.DisplayLayout.Override.RowSelectorHeaderStyle = RowSelectorHeaderStyle.ColumnChooserButton;
            
        }

        private void ultraGrid1_InitializeRow(object sender, InitializeRowEventArgs e) {
            if (DateTime.Now.TimeOfDay.TotalMinutes >= GlobalVar.globalParameter.resultTime) {
                double equivalentNum = Convert.ToDouble(e.Row.Cells["EquivalentNum"].Value);
                double result = e.Row.Cells["Result"].Value == DBNull.Value ? 0 : Convert.ToDouble(e.Row.Cells["Result"].Value);

                if (result + 0.00001 >= equivalentNum) {
                    e.Row.Cells["WarrantName"].Appearance.BackColor = Color.PaleGreen;
                }
            }
            string warrantID = e.Row.Cells["WarrantID"].Value.ToString();
            string underlyingID = e.Row.Cells["UnderlyingID"].Value.ToString();

            string issuable = "NA";
            string reissuable = "NA";

            string toolTip1 = "今日未達增額標準";
            string toolTip2 = "非本季標的";
            string toolTip3 = "標的發行檢查=N";

            string sqlTemp = "SELECT IsNull(Issuable,'NA') Issuable FROM [WarrantAssistant].[dbo].[WarrantUnderlyingSummary] WHERE UnderlyingID = '" + underlyingID + "'";
            //DataView dvTemp = DeriLib.Util.ExecSqlQry(sqlTemp, GlobalVar.loginSet.edisSqlConnString);
            DataTable dtTemp = MSSQL.ExecSqlQry(sqlTemp, GlobalVar.loginSet.warrantassistant45);
            if (dtTemp.Rows.Count > 0)
                issuable = dtTemp.Rows[0]["Issuable"].ToString();

            string sqlTemp2 = "SELECT IsNull([ReIssuable],'NA') ReIssuable FROM [WarrantAssistant].[dbo].[WarrantReIssuable] WHERE WarrantID = '" + warrantID + "'";
            //DataView dvTemp2 = DeriLib.Util.ExecSqlQry(sqlTemp2, GlobalVar.loginSet.edisSqlConnString);
            dtTemp = MSSQL.ExecSqlQry(sqlTemp2, GlobalVar.loginSet.warrantassistant45);
            if (dtTemp.Rows.Count > 0)
                reissuable = dtTemp.Rows[0]["ReIssuable"].ToString();


            if (issuable == "NA") {
                e.Row.ToolTipText = toolTip2;
                e.Row.Appearance.ForeColor = Color.Red;
            } else if (issuable == "N") {
                e.Row.Cells["UnderlyingID"].ToolTipText = toolTip3;
                e.Row.Cells["UnderlyingID"].Appearance.ForeColor = Color.Red;
            }

            if (reissuable == "NA") {
                e.Row.Cells["WarrantID"].ToolTipText = toolTip1;
                e.Row.Cells["WarrantID"].Appearance.ForeColor = Color.Red;
            }
        }
    }
}

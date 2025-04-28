#define To39
using System;
using System.Data;
using System.Drawing;
using System.Windows.Forms;

namespace WarrantAssistant
{
    public partial class FrmReIssuable:Form
    {
        private DataTable dt;

        public FrmReIssuable() {
            InitializeComponent();
        }

        private void FrmReIssuable_Load(object sender, EventArgs e) {
            LoadData();
            InitialGrid();
        }

        private void InitialGrid() {
            dataGridView1.Columns[0].HeaderText = "權證代號";
            dataGridView1.Columns[1].HeaderText = "權證名稱";
            dataGridView1.Columns[2].HeaderText = "標的代號";
            dataGridView1.Columns[3].HeaderText = "標的名稱";
            dataGridView1.Columns[4].HeaderText = "是否可發";
            dataGridView1.Columns[5].HeaderText = "獎勵額度";
            dataGridView1.Columns[6].HeaderText = "交易員";
            dataGridView1.Columns[7].HeaderText = "發行張數";
            dataGridView1.Columns[8].HeaderText = "流通在外";
            dataGridView1.Columns[9].HeaderText = "前1日(%)";
            dataGridView1.Columns[10].HeaderText = "前2日(%)";
            dataGridView1.Columns[11].HeaderText = "前3日(%)";
            dataGridView1.Columns[12].HeaderText = "到期日";

            dataGridView1.Columns[0].Width = 80;
            dataGridView1.Columns[1].Width = 120;
            dataGridView1.Columns[4].Width = 80;
            dataGridView1.Columns[5].Width = 80;
            dataGridView1.Columns[6].Width = 80;
            dataGridView1.Columns[7].Width = 80;
            dataGridView1.Columns[8].Width = 80;
            dataGridView1.Columns[9].Width = 80;
            dataGridView1.Columns[10].Width = 80;
            dataGridView1.Columns[11].Width = 80;

            dataGridView1.Columns[7].DefaultCellStyle.Format = "N0";
            dataGridView1.Columns[8].DefaultCellStyle.Format = "N0";

            dataGridView1.AllowUserToAddRows = false;
            dataGridView1.AllowUserToDeleteRows = false;
            dataGridView1.AllowUserToResizeRows = false;
            dataGridView1.SelectionMode = DataGridViewSelectionMode.RowHeaderSelect;
        }

        private void LoadData() {
            try {
#if !To39
                string sql = @"SELECT a.WarrantID
                                  ,a.WarrantName
                                  ,IsNull(b.UnderlyingID,'NA') UnderlyingID
                                  ,IsNull(b.UnderlyingName,'NA') UnderlyingName
                                  ,IsNull(c.Issuable,'NA') Issuable
                                  ,CASE WHEN b.isReward=1 THEN 'Y' ELSE 'N' END isReward
                                  ,IsNull(d.TraderAccount,'NA') TraderAccount
                                  ,a.IssueNum/1000 as IssueNum
                                  ,a.SoldNum/1000 as SoldNum
                                  ,a.Last1Sold
                                  ,a.Last2Sold
                                  ,a.Last3Sold
                                  ,IsNull(b.ExpiryDate,'') ExpiryDate
                              FROM [WarrantAssistant].[dbo].[WarrantReIssuable] a
                              LEFT JOIN [WarrantAssistant].[dbo].[WarrantBasic] b ON a.WarrantID=b.WarrantID
                              LEFT JOIN [WarrantAssistant].[dbo].[WarrantUnderlyingSummary] c ON c.UnderlyingID=b.UnderlyingID
                              LEFT JOIN [10.19.1.20].[EDIS].[dbo].[Underlying_Trader] d ON d.UID=b.UnderlyingID
                              ORDER BY b.UnderlyingID, a.WarrantID";

                dt = EDLib.SQL.MSSQL.ExecSqlQry(sql, GlobalVar.loginSet.warrantassistant45);
#else
                string sql = @"SELECT a.WarrantID
                                  ,a.WarrantName
                                  ,IsNull(b.UnderlyingID,'NA') UnderlyingID
                                  ,IsNull(b.UnderlyingName,'NA') UnderlyingName
                                  ,IsNull(c.Issuable,'NA') Issuable
                                  ,CASE WHEN b.isReward=1 THEN 'Y' ELSE 'N' END isReward
                                  ,IsNull(d.TraderAccount,'NA') TraderAccount
                                  ,a.IssueNum/1000 as IssueNum
                                  ,a.SoldNum/1000 as SoldNum
                                  ,a.Last1Sold
                                  ,a.Last2Sold
                                  ,a.Last3Sold
                                  ,IsNull(b.ExpiryDate,'') ExpiryDate
                              FROM [WarrantAssistant].[dbo].[WarrantReIssuable] a
                              LEFT JOIN [WarrantAssistant].[dbo].[WarrantBasic] b ON a.WarrantID=b.WarrantID
                              LEFT JOIN [WarrantAssistant].[dbo].[WarrantUnderlyingSummary] c ON c.UnderlyingID=b.UnderlyingID
                              LEFT JOIN [TwData].[dbo].[Underlying_Trader] d ON d.UID=b.UnderlyingID
                              ORDER BY b.UnderlyingID, a.WarrantID";

                dt = EDLib.SQL.MSSQL.ExecSqlQry(sql, GlobalVar.loginSet.warrantassistant45);

#endif
                dataGridView1.DataSource = dt;
            } catch (Exception ex) {
                MessageBox.Show(ex.Message);
            }
        }

        private void dataGridView1_CellFormatting(object sender, DataGridViewCellFormattingEventArgs e) {
            switch (dataGridView1.Columns[e.ColumnIndex].Name) {
                case "Issuable":
                    string cellValue = (string) e.Value;
                    if (cellValue == "N")
                        e.CellStyle.BackColor = Color.Coral;
                    if (cellValue == "NA")
                        e.CellStyle.BackColor = Color.LightYellow;
                    break;
                case "isReward":
                    if ((string) e.Value == "Y")
                        e.CellStyle.BackColor = Color.Coral;
                    break;
                case "ExpiryDate":
                    if ((DateTime) e.Value < DateTime.Today.AddDays(7))
                        e.CellStyle.BackColor = Color.LightYellow;
                    break;
            }          
        }
    }
}

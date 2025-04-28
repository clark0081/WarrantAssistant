
using System;
using System.Data;
using System.Drawing;
using System.Windows.Forms;
using System.Data.SqlClient;
using System.Collections.Generic;
namespace WarrantAssistant
{
    public partial class FrmUnderlyingSummary:Form
    {
        //public SqlConnection conn = new SqlConnection(GlobalVar.loginSet.edisSqlConnString);
        private DataTable dataTable = new DataTable();
        private string enteredKey = "";

        public FrmUnderlyingSummary() {
            InitializeComponent();
        }

        private void InitialGrid() {

            dataGridView1.Columns[0].HeaderText = "標的代號";
            dataGridView1.Columns[1].HeaderText = "標的名稱";
            dataGridView1.Columns[2].HeaderText = "交易員";
            dataGridView1.Columns[3].HeaderText = "市場";
            dataGridView1.Columns[4].HeaderText = "是否可發";
            dataGridView1.Columns[5].HeaderText = "Put發行檢查";
            dataGridView1.Columns[6].HeaderText = "已發行(%)";
            dataGridView1.Columns[7].HeaderText = "已發行張數";
            dataGridView1.Columns[8].HeaderText = "總發行額度(張)";
            dataGridView1.Columns[9].HeaderText = "今日額度";
            dataGridView1.Columns[10].HeaderText = "獎勵額度";
            dataGridView1.Columns[11].HeaderText = "是否虧損";
            dataGridView1.Columns[12].HeaderText = "額度變化";

            dataGridView1.Columns[2].Width = 80;
            dataGridView1.Columns[3].Width = 80;
            dataGridView1.Columns[4].Width = 80;
            dataGridView1.Columns[5].Width = 80;
            dataGridView1.Columns[9].Width = 80;
            dataGridView1.Columns[10].Width = 80;
            dataGridView1.Columns[11].Width = 80;
            dataGridView1.Columns[12].Width = 80;

            dataGridView1.Columns[9].DefaultCellStyle.Format = "N0";            
            dataGridView1.Columns[10].DefaultCellStyle.Format = "N0";
            dataGridView1.Columns[12].DefaultCellStyle.Format = "N0";
            dataGridView1.AllowUserToAddRows = false;
            dataGridView1.AllowUserToDeleteRows = false;
            dataGridView1.AllowUserToResizeRows = false;
            dataGridView1.SelectionMode = DataGridViewSelectionMode.RowHeaderSelect;
           
            //dataGridView1.DefaultCellStyle.SelectionBackColor = Color.White;
            //dataGridView1.DefaultCellStyle.SelectionForeColor = Color.Red;
        }

        private void LoadData(string ID = "") {
            dataTable.Clear();
            bool autocrawl = false;//資訊部撈的標的額度比較晚更新，所以有另一張表會比較早撈額度，如果早上有撈則autocrawl為true

            string sql_credit = $@"SELECT  TOP(1) [UpdateTime]
                            FROM [WarrantAssistant].[dbo].[WarrantUnderlyingCreditNew]
                            WHERE [UpdateTime] >= '{DateTime.Today.ToString("yyyyMMdd")}'
                            ORDER BY [UpdateTime] DESC";
            DataTable dv_credit = EDLib.SQL.MSSQL.ExecSqlQry(sql_credit, GlobalVar.loginSet.warrantassistant45);

            if (dv_credit.Rows.Count > 0)
            {
                autocrawl = true;
            }
            string canissuepercent = $@"";
            string sql;

            if (ID == "")
                /*
                sql = $@"SELECT A.[UnderlyingID], A.[UnderlyingName], A.[TraderID], A.[Market], A.[Issuable], A.[PutIssuable], IsNull(A.[IssuedPercent],0) [IssuedPercent], IsNull(A.[IssueCredit],0) [IssueCredit],  IsNull(A.[RewardIssueCredit],0) [RewardIssueCredit], CASE WHEN A.[AccNetIncome]<0 THEN 'Y' ELSE 'N' END AccNetIncome, A.IssueCreditDelta 
                        , CASE WHEN B.[WRTCAN_STOCKTYPE] ='DE' THEN 100 ELSE 22 END AS CanIssuePercent
                        FROM [EDIS].[dbo].[WarrantUnderlyingSummary] AS A
                        LEFT JOIN (SELECT * 
			                        FROM [10.100.10.131].[WAFT].[dbo].[CANDIDATE] WHERE WRTCAN_DATE = 
			                        (SELECT MAX(WRTCAN_DATE) FROM [10.100.10.131].[WAFT].[dbo].[CANDIDATE])) AS B on B.[WRTCAN_STKID] = A.[UnderlyingID] COLLATE Chinese_Taiwan_Stroke_CI_AS";
                                    */
                sql = $@"SELECT A.[UnderlyingID], A.[UnderlyingName], A.[TraderID], A.[Market], A.[Issuable], A.[PutIssuable], IsNull(A.[IssuedPercent],0) [IssuedPercent], IsNull(A.[IssueCredit],0) [IssueCredit],  IsNull(A.[RewardIssueCredit],0) [RewardIssueCredit], CASE WHEN A.[AccNetIncome]<0 THEN 'Y' ELSE 'N' END AccNetIncome, A.IssueCreditDelta 
                        FROM [WarrantAssistant].[dbo].[WarrantUnderlyingSummary] AS A
                        ";
            else
                /*
                sql = $@"SELECT A.[UnderlyingID] AS UnderlyingID, A.[UnderlyingName] AS UnderlyingName, A.[TraderID] AS TraderID, A.[Market] AS Market, A.[Issuable] AS Issuable, A.[PutIssuable] AS PutIssuable, IsNull(A.[IssuedPercent],0) [IssuedPercent], IsNull(A.[IssueCredit],0) [IssueCredit],  IsNull(A.[RewardIssueCredit],0) [RewardIssueCredit], CASE WHEN A.[AccNetIncome]<0 THEN 'Y' ELSE 'N' END AccNetIncome, A.IssueCreditDelta
                        , CASE WHEN B.[WRTCAN_STOCKTYPE] ='DE' THEN 100 ELSE 22 END AS CanIssuePercent
                        FROM [EDIS].[dbo].[WarrantUnderlyingSummary]  AS A
                        LEFT JOIN (SELECT DISTINCT [WRTCAN_STKID], [WRTCAN_NAME], [WRTCAN_STOCKTYPE]  
			                        FROM [10.100.10.131].[WAFT].[dbo].[CANDIDATE] WHERE WRTCAN_DATE = 
			                        (SELECT MAX(WRTCAN_DATE) FROM [10.100.10.131].[WAFT].[dbo].[CANDIDATE])) AS B on B.[WRTCAN_STKID] = A.[UnderlyingID] COLLATE Chinese_Taiwan_Stroke_CI_AS
                        WHERE TraderID = '{TraderID.Text.TrimStart('0')}'";
                            */
                sql = $@"SELECT A.[UnderlyingID] AS UnderlyingID, A.[UnderlyingName] AS UnderlyingName, A.[TraderID] AS TraderID, A.[Market] AS Market, A.[Issuable] AS Issuable, A.[PutIssuable] AS PutIssuable, IsNull(A.[IssuedPercent],0) [IssuedPercent], IsNull(A.[IssueCredit],0) [IssueCredit],  IsNull(A.[RewardIssueCredit],0) [RewardIssueCredit], CASE WHEN A.[AccNetIncome]<0 THEN 'Y' ELSE 'N' END AccNetIncome, A.IssueCreditDelta
         
                        FROM [WarrantAssistant].[dbo].[WarrantUnderlyingSummary]  AS A
                        
                        WHERE TraderID = '{TraderID.Text.TrimStart('0')}'";
            DataTable dt = EDLib.SQL.MSSQL.ExecSqlQry(sql, GlobalVar.loginSet.warrantassistant45);

            foreach (DataRow dr in dt.Rows)
            {
                try
                {
                    DataRow drv = dataTable.NewRow();
                    
                    drv["UnderlyingID"] = dr["UnderlyingID"].ToString();                    
                    drv["UnderlyingName"] = dr["UnderlyingName"].ToString();                   
                    drv["TraderID"] = dr["TraderID"].ToString();                    
                    drv["Market"] = dr["Market"].ToString();                   
                    drv["Issuable"] = dr["Issuable"].ToString();                   
                    drv["PutIssuable"] = dr["PutIssuable"].ToString();
                    drv["IssuedPercent"] = Convert.ToDouble(dr["IssuedPercent"].ToString());
                    drv["IssueCredit"] = Convert.ToDouble(dr["IssueCredit"].ToString());
                    if (autocrawl)
                    {

                        string sql1 = $@"SELECT  [IssuedShares], [WarrantAvailableShares], [IssuedPercent], [CanIssue]
                                        FROM [WarrantAssistant].[dbo].[WarrantUnderlyingCreditNew]
                                        WHERE [UID] ='{dr["UnderlyingID"].ToString()}' AND [UpdateTime] >='{DateTime.Today.ToString("yyyyMMdd")}'";
                        DataTable dt1 = EDLib.SQL.MSSQL.ExecSqlQry(sql1, GlobalVar.loginSet.warrantassistant45);

                        if (dt1.Rows.Count > 0)
                        {
                            drv["IssuedShares"] = Convert.ToDouble(dt1.Rows[0]["IssuedShares"].ToString());
                            //drv["AvailableShares"] = Convert.ToDouble(dt1.Rows[0]["WarrantAvailableShares"].ToString()) * Convert.ToDouble(dr["CanIssuePercent"].ToString()) / 100;
                            drv["AvailableShares"] = Convert.ToDouble(dt1.Rows[0]["WarrantAvailableShares"].ToString());
                            drv["IssuedPercent"] = Convert.ToDouble(dt1.Rows[0]["IssuedPercent"].ToString());
                            drv["IssueCredit"] = Convert.ToDouble(dt1.Rows[0]["CanIssue"].ToString());
                        }
                        else
                        {
                            drv["IssuedShares"] = 0;
                            drv["AvailableShares"] = 0;
                        }
                    }
                    else
                    {
                        drv["IssuedShares"] = 0;
                        drv["AvailableShares"] = 0;
                    }
                    
                    
                    drv["RewardIssueCredit"] = Convert.ToDouble(dr["RewardIssueCredit"].ToString());
                    drv["AccNetIncome"] = dr["AccNetIncome"].ToString();
                    string icd = dr["IssueCreditDelta"].ToString();
                    double ICD = (icd == "") ? 0 : Convert.ToDouble(icd);
                    drv["IssueCreditDelta"] = ICD;
                    dataTable.Rows.Add(drv);
                }
                catch(Exception ex)
                {
                    MessageBox.Show($"{ex.Message}");
                }
            }
            dataGridView1.DataSource = dataTable;
            foreach (DataRow row in dataTable.Rows) {
                row["IssuedPercent"] = Math.Round((double) row["IssuedPercent"], 2);
                row["AvailableShares"] = Math.Round((double)row["AvailableShares"], 2);
                row["IssueCredit"] = Math.Round((double)row["IssueCredit"], 2);
            }
        }

        private void FrmUnderlyingSummary_Load(object sender, EventArgs e) {
            dataTable.Columns.Add("UnderlyingID", typeof(string));
            dataTable.Columns.Add("UnderlyingName", typeof(string));
            dataTable.Columns.Add("TraderID", typeof(string));
            dataTable.Columns.Add("Market", typeof(string));
            dataTable.Columns.Add("Issuable", typeof(string));
            dataTable.Columns.Add("PutIssuable", typeof(string));
            dataTable.Columns.Add("IssuedPercent", typeof(double));
            dataTable.Columns.Add("IssuedShares", typeof(double));
            dataTable.Columns.Add("AvailableShares", typeof(double));
            dataTable.Columns.Add("IssueCredit", typeof(double));
            dataTable.Columns.Add("RewardIssueCredit", typeof(double));
            dataTable.Columns.Add("AccNetIncome", typeof(string));
            dataTable.Columns.Add("IssueCreditDelta", typeof(double));
            LoadData();
            InitialGrid();
            foreach (var item in GlobalVar.globalParameter.traders)
                TraderID.Items.Add(item);
            TraderID.Items.Add("");
        }

        private void dataGridView1_CellFormatting(object sender, DataGridViewCellFormattingEventArgs e) {
            switch (dataGridView1.Columns[e.ColumnIndex].Name) {
                case "Issuable":
                case "PutIssuable":
                    if ((string) e.Value == "N")
                        e.CellStyle.BackColor = Color.Coral;
                    break;
                case "IssueCredit":
                    if ((double) e.Value < 0)
                        e.CellStyle.BackColor = Color.Coral;
                    break;
                case "AccNetIncome":
                    if ((string) e.Value == "Y")
                        e.CellStyle.BackColor = Color.Coral;
                    break;
            }
        }
       
        public void SelectUnderlying(string underlyingID) {
            GlobalUtility.SelectUnderlying(underlyingID, dataGridView1);
        }

        private void dataGridView1_KeyDown(object sender, KeyEventArgs e) {
            try {
                /*
                if(e.Control == true && e.KeyCode == Keys.C)
                {
                    MessageBox.Show("copy");
                    foreach(DataGridViewCell cell in dataGridView1.SelectedCells)
                    {
                        MessageBox.Show(cell.Value.ToString());
                    }
                        
                }
                */
                if (e.KeyCode == Keys.Enter) {
                    
                    SelectUnderlying(enteredKey);
                    e.Handled = true;
                    enteredKey = "";
                } else
                    GlobalUtility.KeyDecoder(e, ref enteredKey);
            } catch (Exception ex) {
                MessageBox.Show(ex.Message);
            }
        }

        private void dataGridView1_CellDoubleClick(object sender, DataGridViewCellEventArgs e) {
            if (e.RowIndex == -1) return;//header有時候header要排序點太快會變雙擊
            string target = (string) dataGridView1.Rows[e.RowIndex].Cells[0].Value;
            switch (dataGridView1.Columns[e.ColumnIndex].Name) {
                case "Issuable":
                    GlobalUtility.MenuItemClick<FrmIssueCheck>().SelectUnderlying(target);
                    break;
                case "PutIssuable":
                    GlobalUtility.MenuItemClick<FrmIssueCheckPut>().SelectUnderlying(target);
                    break;
            }
        }

        private void TraderID_SelectedIndexChanged(object sender, EventArgs e) {
            LoadData(TraderID.Text);            
        }
    }
}

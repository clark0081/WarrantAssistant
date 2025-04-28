#define To39
using System;
using System.Data;
using System.Drawing;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using System.Collections.Generic;
using System.Configuration;

namespace WarrantAssistant
{
    public partial class FrmIssueCheckPut:Form
    {
        private System.Data.DataTable dataTable;
        private string enteredKey = "";
        public double NonSpecialCallPutRatio = Convert.ToDouble(ConfigurationManager.AppSettings["NonSpecialCallPutRatio"].ToString());
        public double SpecialCallPutRatio = Convert.ToDouble(ConfigurationManager.AppSettings["SpecialCallPutRatio"].ToString());
        public double SpecialKGIALLPutRatio = Convert.ToDouble(ConfigurationManager.AppSettings["SpecialKGIALLPutRatio"].ToString());
        public double ISTOP30MaxIssue = Convert.ToDouble(ConfigurationManager.AppSettings["ISTOP30MaxIssue"].ToString());
        public double NonTOP30MaxIssue = Convert.ToDouble(ConfigurationManager.AppSettings["NonTOP30MaxIssue"].ToString());
        public FrmIssueCheckPut() {
            InitializeComponent();
        }

        private void FrmIssueCheckPut_Load(object sender, EventArgs e) {
            LoadData();
            InitialGrid();

            comboBox1.Text = "All";
            foreach (var item in GlobalVar.globalParameter.traders)
            {
                comboBox1.Items.Add(item.TrimStart('0'));
            }
            comboBox1.Items.Add("SYS");
            comboBox1.Items.Add("All");
        }

        private void InitialGrid() {
            dataGridView1.Columns[0].HeaderText = "標的代號";
            dataGridView1.Columns[1].HeaderText = "標的名稱";
            dataGridView1.Columns[2].HeaderText = "台灣50成分股";
            dataGridView1.Columns[3].HeaderText = "本益比";
            dataGridView1.Columns[4].HeaderText = "過去一年損益";
            dataGridView1.Columns[5].HeaderText = "股價";
            dataGridView1.Columns[6].HeaderText = "前一季股價";
            dataGridView1.Columns[7].HeaderText = "前一年股價";
            dataGridView1.Columns[8].HeaderText = "季報酬";
            dataGridView1.Columns[9].HeaderText = "年報酬";
            dataGridView1.Columns[10].HeaderText = "自家Call_DeltaOne";
            dataGridView1.Columns[11].HeaderText = "自家Put_DeltaOne";
            dataGridView1.Columns[12].HeaderText = "自家Call/Put 比例";
            dataGridView1.Columns[13].HeaderText = "全市場Call_DeltaOne";
            dataGridView1.Columns[14].HeaderText = "全市場Put_DeltaOne";
            dataGridView1.Columns[15].HeaderText = "自家/全市場Put_DeltaOne比例";
            dataGridView1.Columns[16].HeaderText = "元大Put_DeltaOne";
            dataGridView1.Columns[17].HeaderText = "自家已發行Put";
            dataGridView1.Columns[18].HeaderText = "特殊標的";

            dataGridView1.Columns[4].DefaultCellStyle.Format = "N0";
            dataGridView1.ColumnHeadersHeight = 40;
            
            dataGridView1.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
            dataGridView1.Columns[0].Width = 60;
            dataGridView1.Columns[1].Width = 60;
            dataGridView1.Columns[2].Width = 70;
            dataGridView1.Columns[2].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dataGridView1.Columns[3].Width = 80;
            dataGridView1.Columns[4].Width = 120;
            dataGridView1.Columns[5].Width = 60;
            dataGridView1.Columns[6].Width = 70;
            dataGridView1.Columns[7].Width = 70;
            dataGridView1.Columns[8].Width = 70;
            dataGridView1.Columns[9].Width = 70;
            dataGridView1.Columns[10].Width = 80;
            dataGridView1.Columns[11].Width = 80;
            dataGridView1.Columns[12].Width = 80;
            dataGridView1.Columns[13].Width = 100;
            dataGridView1.Columns[14].Width = 100;
            dataGridView1.Columns[15].Width = 120;
            dataGridView1.Columns[16].Width = 80;
            dataGridView1.Columns[17].Width = 110;
            dataGridView1.Columns[17].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dataGridView1.Columns[18].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dataGridView1.Columns[19].Visible = false;

            dataGridView1.AllowUserToAddRows = false;
            dataGridView1.AllowUserToDeleteRows = false;
            dataGridView1.AllowUserToResizeRows = false;
            dataGridView1.SelectionMode = DataGridViewSelectionMode.RowHeaderSelect;
        }

        private void LoadData() {

#if !To39
            string sql = $@"SELECT A.[UnderlyingID], A.[UnderlyingName], A.[IsTW50Stocks], A.[PERatio], A.[SumEarning], A.[Price], A.[PriceQuarter]
                                 ,A.[PriceYear], A.[ReturnQuarter], A.[ReturnYear], B.[KgiCallDeltaOne], B.[KgiPutDeltaOne], B.[KgiCallPutRatio], B.[AllCallDeltaOne]
								 ,B.[AllPutDeltaOne], B.[KgiAllPutRatio], B.[YuanPutDeltaOne], B.[KgiPutNum], B.[IsSpecial], C.[TraderAccount] 
                                FROM [WarrantAssistant].[dbo].[WarrantIssueCheckPut] AS A
							    LEFT JOIN  (SELECT * FROM [WarrantAssistant].[dbo].[WarrantIssueDeltaOne] 
								WHERE [DateTime]='{DateTime.Today.ToString("yyyyMMdd")}') AS B on A.[UnderlyingID] = B.[UnderlyingID] 
                                LEFT JOIN  [10.19.1.20].[EDIS].[dbo].[Underlying_Trader] as C on A.UnderlyingID = C.UID
                                WHERE [DateTime] >='{DateTime.Today.ToString("yyyyMMdd")}'";
#else
            string sql = $@"SELECT A.[UnderlyingID], A.[UnderlyingName], A.[IsTW50Stocks], A.[PERatio], A.[SumEarning], A.[Price], A.[PriceQuarter]
                                 ,A.[PriceYear], A.[ReturnQuarter], A.[ReturnYear], B.[KgiCallDeltaOne], B.[KgiPutDeltaOne], B.[KgiCallPutRatio], B.[AllCallDeltaOne]
								 ,B.[AllPutDeltaOne], B.[KgiAllPutRatio], B.[YuanPutDeltaOne], B.[KgiPutNum], B.[IsSpecial], C.[TraderAccount] 
                                FROM [WarrantAssistant].[dbo].[WarrantIssueCheckPut] AS A
							    LEFT JOIN  (SELECT * FROM [WarrantAssistant].[dbo].[WarrantIssueDeltaOne] 
								WHERE [DateTime]='{DateTime.Today.ToString("yyyyMMdd")}') AS B on A.[UnderlyingID] = B.[UnderlyingID] 
                                LEFT JOIN  [TwData].[dbo].[Underlying_Trader] as C on A.UnderlyingID = C.UID
                                WHERE [DateTime] >='{DateTime.Today.ToString("yyyyMMdd")}'";
#endif
            dataTable = EDLib.SQL.MSSQL.ExecSqlQry(sql, GlobalVar.loginSet.warrantassistant45);

            dataGridView1.DataSource = dataTable;
            foreach (DataRow row in dataTable.Rows) {
                row["ReturnQuarter"] = Math.Round((double) row["ReturnQuarter"], 2);
                row["ReturnYear"] = Math.Round((double) row["ReturnYear"], 2);
                row["KgiCallDeltaOne"] = Math.Round((double)row["KgiCallDeltaOne"],2);
                row["KgiPutDeltaOne"] = Math.Round((double)row["KgiPutDeltaOne"], 2);
                row["AllCallDeltaOne"] = Math.Round((double)row["AllCallDeltaOne"], 2);
                row["AllPutDeltaOne"] = Math.Round((double)row["AllPutDeltaOne"], 2);
            }
        }

        private void dataGridView1_CellFormatting(object sender, DataGridViewCellFormattingEventArgs e) {
            switch (dataGridView1.Columns[e.ColumnIndex].Name) {
                case "PERatio":
                    if ((double) e.Value > 40)
                        e.CellStyle.BackColor = Color.LightPink;
                    break;
                case "SumEarning":
                    if ((double) e.Value < 0)
                        e.CellStyle.BackColor = Color.LightPink;
                    break;
                case "ReturnQuarter":
                    if ((double) e.Value > 0.5)
                        e.CellStyle.BackColor = Color.LightPink;
                    break;
                case "ReturnYear":
                    if ((double) e.Value > 1)
                        e.CellStyle.BackColor = Color.LightPink;
                    break;
                case "KgiCallPutRatio":
                    if ((double)e.Value < NonSpecialCallPutRatio)
                    {
                        if(Convert.ToInt32(dataGridView1[17,e.RowIndex].Value) >0)
                            e.CellStyle.BackColor = Color.LightPink;
                    }
                    break;
                case "KgiAllPutRatio":
                    if(((double)e.Value > SpecialKGIALLPutRatio) && (Convert.ToDouble(dataGridView1[14, e.RowIndex].Value) > 0))
                    {
                        e.CellStyle.BackColor = Color.LightPink;
                    }
                    break;
                case "KgiPutNum":
                    if((int)e.Value == 0)
                    {
                        e.CellStyle.BackColor = Color.MediumAquamarine;
                    }
                    break;
                case "YuanPutDeltaOne":
                    if(dataGridView1[18, e.RowIndex].Value.ToString()=="Y")
                    {
                        if (Convert.ToDouble(dataGridView1[11, e.RowIndex].Value) > Convert.ToDouble(e.Value))
                            e.CellStyle.BackColor = Color.LightPink;
                    }
                    break;
            }
        }

        public void SelectUnderlying(string underlyingID) {
            GlobalUtility.SelectUnderlying(underlyingID, dataGridView1);           
        }

        private void dataGridView1_KeyDown(object sender, KeyEventArgs e) {
            try {
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
        private void ComboBox1_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter && comboBox1.Text != "")
            {
                LoadTraderSql();
                //toolStripComboBox1.Text = "";
            }
        }

        private void ComboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            LoadTraderSql();
        }
        private void LoadTraderSql()
        {
            string sql = "";
            string trader = comboBox1.Text;
            if (trader == "All")
            {
                LoadData();
                return;
            }
#if !To39
            if (trader == "SYS")
            {
                sql = $@"SELECT A.[UnderlyingID], A.[UnderlyingName], A.[IsTW50Stocks], A.[PERatio], A.[SumEarning], A.[Price], A.[PriceQuarter]
                                 ,A.[PriceYear], A.[ReturnQuarter], A.[ReturnYear], B.[KgiCallDeltaOne], B.[KgiPutDeltaOne], B.[KgiCallPutRatio], B.[AllCallDeltaOne]
								 ,B.[AllPutDeltaOne], B.[KgiAllPutRatio], B.[YuanPutDeltaOne], B.[KgiPutNum], B.[IsSpecial], C.[TraderAccount] 
                                FROM [WarrantAssistant].[dbo].[WarrantIssueCheckPut] AS A
							    LEFT JOIN  (SELECT * FROM [WarrantAssistant].[dbo].[WarrantIssueDeltaOne] 
								WHERE [DateTime]='{DateTime.Today.ToString("yyyyMMdd")}') AS B on A.[UnderlyingID] = B.[UnderlyingID] 
                                LEFT JOIN  [10.19.1.20].[EDIS].[dbo].[Underlying_Trader] as C on A.UnderlyingID = C.UID
                                WHERE [TraderAccount] IS NULL";
            }

            else
            {
                sql = $@"SELECT A.[UnderlyingID], A.[UnderlyingName], A.[IsTW50Stocks], A.[PERatio], A.[SumEarning], A.[Price], A.[PriceQuarter]
                                 ,A.[PriceYear], A.[ReturnQuarter], A.[ReturnYear], B.[KgiCallDeltaOne], B.[KgiPutDeltaOne], B.[KgiCallPutRatio], B.[AllCallDeltaOne]
								 ,B.[AllPutDeltaOne], B.[KgiAllPutRatio], B.[YuanPutDeltaOne], B.[KgiPutNum], B.[IsSpecial], C.[TraderAccount] 
                                FROM [WarrantAssistant].[dbo].[WarrantIssueCheckPut] AS A
							    LEFT JOIN  (SELECT * FROM [WarrantAssistant].[dbo].[WarrantIssueDeltaOne] 
								WHERE [DateTime]='{DateTime.Today.ToString("yyyyMMdd")}') AS B on A.[UnderlyingID] = B.[UnderlyingID] 
                                LEFT JOIN  [10.19.1.20].[EDIS].[dbo].[Underlying_Trader] as C on A.UnderlyingID = C.UID
                                WHERE [TraderAccount] ='{trader}'";
            }
            dataTable = EDLib.SQL.MSSQL.ExecSqlQry(sql, GlobalVar.loginSet.warrantassistant45);
#else
            if (trader == "SYS")
            {
                sql = $@"SELECT A.[UnderlyingID], A.[UnderlyingName], A.[IsTW50Stocks], A.[PERatio], A.[SumEarning], A.[Price], A.[PriceQuarter]
                                 ,A.[PriceYear], A.[ReturnQuarter], A.[ReturnYear], B.[KgiCallDeltaOne], B.[KgiPutDeltaOne], B.[KgiCallPutRatio], B.[AllCallDeltaOne]
								 ,B.[AllPutDeltaOne], B.[KgiAllPutRatio], B.[YuanPutDeltaOne], B.[KgiPutNum], B.[IsSpecial], C.[TraderAccount] 
                                FROM [WarrantAssistant].[dbo].[WarrantIssueCheckPut] AS A
							    LEFT JOIN  (SELECT * FROM [WarrantAssistant].[dbo].[WarrantIssueDeltaOne] 
								WHERE [DateTime]='{DateTime.Today.ToString("yyyyMMdd")}') AS B on A.[UnderlyingID] = B.[UnderlyingID] 
                                LEFT JOIN  [TwData].[dbo].[Underlying_Trader] as C on A.UnderlyingID = C.UID
                                WHERE [TraderAccount] IS NULL";
            }

            else
            {
                sql = $@"SELECT A.[UnderlyingID], A.[UnderlyingName], A.[IsTW50Stocks], A.[PERatio], A.[SumEarning], A.[Price], A.[PriceQuarter]
                                 ,A.[PriceYear], A.[ReturnQuarter], A.[ReturnYear], B.[KgiCallDeltaOne], B.[KgiPutDeltaOne], B.[KgiCallPutRatio], B.[AllCallDeltaOne]
								 ,B.[AllPutDeltaOne], B.[KgiAllPutRatio], B.[YuanPutDeltaOne], B.[KgiPutNum], B.[IsSpecial], C.[TraderAccount] 
                                FROM [WarrantAssistant].[dbo].[WarrantIssueCheckPut] AS A
							    LEFT JOIN  (SELECT * FROM [WarrantAssistant].[dbo].[WarrantIssueDeltaOne] 
								WHERE [DateTime]='{DateTime.Today.ToString("yyyyMMdd")}') AS B on A.[UnderlyingID] = B.[UnderlyingID] 
                                LEFT JOIN  [TwData].[dbo].[Underlying_Trader] as C on A.UnderlyingID = C.UID
                                WHERE [TraderAccount] ='{trader}'";
            }
            dataTable = EDLib.SQL.MSSQL.ExecSqlQry(sql, GlobalVar.loginSet.warrantassistant45);
#endif
            dataGridView1.DataSource = dataTable;
            foreach (DataRow row in dataTable.Rows)
            {
                row["ReturnQuarter"] = Math.Round((double)row["ReturnQuarter"], 2);
                row["ReturnYear"] = Math.Round((double)row["ReturnYear"], 2);
                row["KgiCallDeltaOne"] = Math.Round((double)row["KgiCallDeltaOne"], 2);
                row["KgiPutDeltaOne"] = Math.Round((double)row["KgiPutDeltaOne"], 2);
                row["AllCallDeltaOne"] = Math.Round((double)row["AllCallDeltaOne"], 2);
                row["AllPutDeltaOne"] = Math.Round((double)row["AllPutDeltaOne"], 2);
            }

        }

        private void dataGridView1_DataBindingComplete(object sender, DataGridViewBindingCompleteEventArgs e)
        {
            foreach (DataGridViewRow row in dataGridView1.Rows)
            {
                row.HeaderCell.Value = (row.Index + 1).ToString();
            }
        }
        private void button1_Click(object sender, EventArgs e)
        {
            SaveFileDialog sFileDialog = new SaveFileDialog();
            sFileDialog.Title = "匯出Excel";
            sFileDialog.Filter = "EXCEL檔 (*.xlsx)|*.xlsx";
            sFileDialog.InitialDirectory = "D:\\";
            Microsoft.Office.Interop.Excel.Application app = new Microsoft.Office.Interop.Excel.Application();
            if (sFileDialog.ShowDialog() == System.Windows.Forms.DialogResult.OK && sFileDialog.FileName != null)
            {

                Workbook workbook = app.Workbooks.Add(1);
                Worksheet worksheets = workbook.Sheets[1];
                
                for(int i = 0; i < dataGridView1.ColumnCount; i++)
                {
                    worksheets.get_Range($"{(char)(65 + i) + "1"}", $"{(char)(65 + i) + "1"}").Value = dataGridView1.Columns[i].HeaderText;
                }
                Microsoft.Office.Interop.Excel.Range range = worksheets.get_Range("A2", $"A{dataGridView1.RowCount + 1}");
                range.NumberFormat = "@";
                for (int i = 0; i < dataGridView1.RowCount; i++)
                {
                    for(int j = 0; j < dataGridView1.ColumnCount; j++)
                    {
                        worksheets.get_Range($"{(char)(65 + j) + (i + 2).ToString()}", $"{(char)(65 + j) + (i + 2).ToString()}").Value = dataGridView1.Rows[i].Cells[j].Value;
                        //MessageBox.Show($"{(char)(65 + j) + (i + 2).ToString()} {(char)(65 + j) + (i + 2).ToString()}  { dataGridView1.Rows[i].Cells[j].Value}");
                    }
                }
                workbook.SaveAs(sFileDialog.FileName);
                workbook.Close();
                app.Quit();
                MessageBox.Show("匯出完成");
            }
        }
    }
}

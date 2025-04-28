#define To45
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
using Microsoft.Office.Interop.Excel;
using Infragistics.Win;
using Infragistics.Win.UltraWinGrid;

namespace WarrantAssistant
{
    public partial class FrmExpireRelease : Form
    {
        private System.Data.DataTable dataTable = new System.Data.DataTable();
        public FrmExpireRelease()
        {
            InitializeComponent();
        }

        private void FrmExpireRelease_Load(object sender, EventArgs e)
        {
            comboBox1.Text = "All";
            foreach (var item in GlobalVar.globalParameter.traders)
            {
                comboBox1.Items.Add(item.TrimStart('0'));
            }
            
            comboBox1.Items.Add("SYS");
            comboBox1.Items.Add("All");
            dataTable.Columns.Add("股票代號", typeof(string));
            dataTable.Columns.Add("強制註銷(估)", typeof(double));
            dataTable.Columns.Add("Today", typeof(double));
            dataTable.Columns.Add("Day1", typeof(double));
            dataTable.Columns.Add("Day2", typeof(double));
            dataTable.Columns.Add("Day3", typeof(double));
            dataTable.Columns.Add("Day4", typeof(double));
            LoadData();
            InitialGrid();

        }
        private void LoadData()
        {
            dataTable.Clear();
#if !To45
            string sql = $@"SELECT  [UID]
                          ,[LogOut]
                          ,[Day0]
                          ,[Day1]
                          ,[Day2]
                          ,[Day3]
                          ,[Day4]
                      FROM [EDIS].[dbo].[ExpireRelease]";

            System.Data.DataTable dt = EDLib.SQL.MSSQL.ExecSqlQry(sql, GlobalVar.loginSet.edisSqlConnString);
#else
            string sql = $@"SELECT  [UID]
                          ,[LogOut]
                          ,[Day0]
                          ,[Day1]
                          ,[Day2]
                          ,[Day3]
                          ,[Day4]
                      FROM [WarrantAssistant].[dbo].[ExpireRelease]";

            System.Data.DataTable dt = EDLib.SQL.MSSQL.ExecSqlQry(sql, GlobalVar.loginSet.warrantassistant45);
#endif
            foreach (DataRow dr in dt.Rows)
            {
                DataRow drtemp = dataTable.NewRow();
                drtemp["股票代號"] = dr["UID"].ToString();
                drtemp["強制註銷(估)"] = Math.Round(Convert.ToDouble(dr["LogOut"].ToString()), 2);
                
                drtemp["Today"] = Math.Round(Convert.ToDouble(dr["Day0"].ToString()), 2);
                drtemp["Day1"] = Math.Round(Convert.ToDouble(dr["Day1"].ToString()), 2);
                drtemp["Day2"] = Math.Round(Convert.ToDouble(dr["Day2"].ToString()), 2);
                drtemp["Day3"] = Math.Round(Convert.ToDouble(dr["Day3"].ToString()), 2);
                drtemp["Day4"] = Math.Round(Convert.ToDouble(dr["Day4"].ToString()), 2);
                dataTable.Rows.Add(drtemp);
            }
            ultraGrid1.DataSource = dataTable;
        }
        private void InitialGrid()
        {
            UltraGridBand band0 = ultraGrid1.DisplayLayout.Bands[0];
            this.ultraGrid1.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti;

            // It is good practice to clear the sorted columns collection
            band0.SortedColumns.Clear();
            band0.Columns["股票代號"].CellActivation = Activation.NoEdit;
            band0.Columns["強制註銷(估)"].CellActivation = Activation.NoEdit;
            
            band0.Columns["Today"].CellActivation = Activation.NoEdit;
            band0.Columns["Day1"].CellActivation = Activation.NoEdit;
            band0.Columns["Day2"].CellActivation = Activation.NoEdit;
            band0.Columns["Day3"].CellActivation = Activation.NoEdit;
            band0.Columns["Day4"].CellActivation = Activation.NoEdit;

            band0.Columns["Day1"].Hidden = true;
            band0.Columns["Day2"].Hidden = true;
            band0.Columns["Day3"].Hidden = true;
            band0.Columns["Day4"].Hidden = true;

        }
        private void UltraGrid1_InitializeLayout(object sender, InitializeLayoutEventArgs e)
        {
            ultraGrid1.DisplayLayout.Override.RowSelectorHeaderStyle = RowSelectorHeaderStyle.ColumnChooserButton;



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

            if (trader == "SYS")
            {
                sql = $@"SELECT  A.[UID], A.[LogOut], A.[Day0], A.[Day1], A.[Day2], A.[Day3], A.[Day4]
                        FROM [WarrantAssistant].[dbo].[ExpireRelease] AS A
                        LEFT JOIN [TwData].[dbo].[Underlying_Trader] as B on A.[UID] = B.[UID]
                        WHERE  B.[TraderAccount] IS NULL";
            }

            else
            {
                sql = $@"SELECT  A.[UID], A.[LogOut], A.[Day0], A.[Day1], A.[Day2], A.[Day3], A.[Day4]
                        FROM [WarrantAssistant].[dbo].[ExpireRelease] AS A
                        LEFT JOIN [TwData].[dbo].[Underlying_Trader] as B on A.[UID] = B.[UID]
                        WHERE B.[TraderAccount] ='{trader}'";
            }
            dataTable.Clear();
            System.Data.DataTable dt = EDLib.SQL.MSSQL.ExecSqlQry(sql, GlobalVar.loginSet.warrantassistant45);

            foreach (DataRow dr in dt.Rows)
            {

                DataRow drtemp = dataTable.NewRow();
                drtemp["股票代號"] = dr["UID"].ToString();
                drtemp["強制註銷(估)"] = Math.Round(Convert.ToDouble(dr["LogOut"].ToString()), 2);

                drtemp["Today"] = Math.Round(Convert.ToDouble(dr["Day0"].ToString()), 2);
                drtemp["Day1"] = Math.Round(Convert.ToDouble(dr["Day1"].ToString()), 2);
                drtemp["Day2"] = Math.Round(Convert.ToDouble(dr["Day2"].ToString()), 2);
                drtemp["Day3"] = Math.Round(Convert.ToDouble(dr["Day3"].ToString()), 2);
                drtemp["Day4"] = Math.Round(Convert.ToDouble(dr["Day4"].ToString()), 2);
                dataTable.Rows.Add(drtemp);
            }
            ultraGrid1.DataSource = dataTable;

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
                UltraGridBand band0 = ultraGrid1.DisplayLayout.Bands[0];
                for (int i = 0; i < band0.Columns.Count; i++)
                {
                    worksheets.get_Range($"{(char)(65 + i) + "1"}", $"{(char)(65 + i) + "1"}").Value = dataTable.Columns[i].ColumnName;
                }
                //設成文字格式，不然0050會變50
                Microsoft.Office.Interop.Excel.Range range = worksheets.get_Range("A2", $"A{ultraGrid1.Rows.Count + 1}");
                range.NumberFormat = "@";
                for (int i = 0; i < ultraGrid1.Rows.Count; i++)
                {
                    for (int j = 0; j < band0.Columns.Count; j++)
                    {

                        worksheets.get_Range($"{(char)(65 + j) + (i + 2).ToString()}", $"{(char)(65 + j) + (i + 2).ToString()}").Value = ultraGrid1.Rows[i].Cells[j].Value;

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

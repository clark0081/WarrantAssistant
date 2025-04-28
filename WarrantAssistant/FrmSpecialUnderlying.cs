using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Infragistics.Win.UltraWinGrid;
using EDLib.SQL;
namespace WarrantAssistant
{
    public partial class FrmSpecialUnderlying : Form
    {
        DataTable dt = new DataTable();
        DataTable dt_R = new DataTable();//量化指標風險級數>=5
        DataTable dt_B = new DataTable();//生技文創
        DataTable dt_KY = new DataTable();//資本額小於10億之KY公司
        DataTable dt_M = new DataTable();//市值前30大
        public FrmSpecialUnderlying()
        {
            InitializeComponent();
        }

        private void FrmSpecialUnderlying_Load(object sender, EventArgs e)
        {
            LoadData();
            InitialGrid();
        }
        private void LoadData()
        {

            string sql = $@"SELECT [UID], [UName], [Type] FROM [WarrantAssistant].[dbo].[SpecialStock]
                            WHERE [DataDate] >=(select Max([Datadate]) FROM [WarrantAssistant].[dbo].[SpecialStock])";
            dt = MSSQL.ExecSqlQry(sql, GlobalVar.loginSet.warrantassistant45);

        }
        private void InitialGrid()
        {
            LoadGrid(ultraGrid1, "M", dt_M);
            LoadGrid(ultraGrid2, "R", dt_R);
            LoadGrid(ultraGrid3, "B", dt_B);
            LoadGrid(ultraGrid4, "KY", dt_KY);
            ultraGrid1.Select();
        }
        private void LoadGrid(UltraGrid grid, string type, DataTable dt_kind)
        {
            dt_kind.Columns.Add("股票代號",typeof(string));
            dt_kind.Columns.Add("股票名稱", typeof(string));
            DataRow[] dt_type = dt.Select($@"[Type] = '{type}'");
            foreach(DataRow dr in dt_type)
            {
                DataRow dr_temp = dt_kind.NewRow();
                dr_temp["股票代號"] = dr["UID"].ToString();
                dr_temp["股票名稱"] = dr["UName"].ToString();
                dt_kind.Rows.Add(dr_temp);
            }
            grid.DataSource = dt_kind;
            UltraGridBand band0 = grid.DisplayLayout.Bands[0];
            band0.Columns["股票代號"].CellActivation = Activation.NoEdit;
            band0.Columns["股票名稱"].CellActivation = Activation.NoEdit;


        }
        private void SelectGrid(object sender, EventArgs e)
        {
            switch (this.tabControl1.SelectedTab.Name.ToString())
            {
                case "tabPage1":
                    ultraGrid1.Select();
                    break;
                case "tabPage2":
                    ultraGrid2.Select();
                    break;
                case "tabPage3":
                    ultraGrid3.Select();
                    break;
                case "tabPage4":
                    ultraGrid4.Select();
                    break;
            }
            
        }
    }
}

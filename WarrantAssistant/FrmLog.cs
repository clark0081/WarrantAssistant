using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Threading;
using Infragistics.Shared;
using Infragistics.Win;
using Infragistics.Win.UltraWinGrid;
using System.Data.SqlClient;


namespace WarrantAssistant
{
    public partial class FrmLog : Form
    {
        public DataTable dt = new DataTable();

        public FrmLog()
        {
            InitializeComponent();
        }

        private void FrmLog_Load(object sender, EventArgs e)
        {
            SetUltraGrid1();
            LoadUltraGrid1();
        }

        private void SetUltraGrid1()
        {
            dt.Columns.Add("時間", typeof(string));
            dt.Columns.Add("類型", typeof(string));
            dt.Columns.Add("內容", typeof(string));
            dt.Columns.Add("人員", typeof(string));

            ultraGrid1.DataSource = dt;

            ultraGrid1.DisplayLayout.Bands[0].Columns["時間"].Width = 60;
            ultraGrid1.DisplayLayout.Bands[0].Columns["類型"].Width = 30;
            ultraGrid1.DisplayLayout.Bands[0].Columns["人員"].Width = 30;
            //ultraGrid1.DisplayLayout.Bands[0].Override.HeaderAppearance.TextHAlign = Infragistics.Win.HAlign.Left;
            ultraGrid1.DisplayLayout.Bands[0].ColHeadersVisible = false;
            ultraGrid1.DisplayLayout.AutoFitStyle = AutoFitStyle.ResizeAllColumns;
            ultraGrid1.DisplayLayout.Override.CellAppearance.BorderAlpha = Alpha.Transparent;
            ultraGrid1.DisplayLayout.Override.RowAppearance.BorderAlpha = Alpha.Transparent;
            ultraGrid1.DisplayLayout.Bands[0].Columns[3].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right;
            ultraGrid1.DisplayLayout.Bands[0].Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
            ultraGrid1.DisplayLayout.Bands[0].Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
            ultraGrid1.DisplayLayout.Bands[0].Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.False;

            ultraGrid1.DisplayLayout.Bands[0].Columns["時間"].CellActivation = Activation.NoEdit;
            ultraGrid1.DisplayLayout.Bands[0].Columns["類型"].CellActivation = Activation.NoEdit;
            ultraGrid1.DisplayLayout.Bands[0].Columns["內容"].CellActivation = Activation.NoEdit;
            ultraGrid1.DisplayLayout.Bands[0].Columns["人員"].CellActivation = Activation.NoEdit;
        }

        private void LoadUltraGrid1()
        {
            try
            {
                dt.Rows.Clear();

                string sql = @"SELECT [MDate]
                                  ,[InformationType]
                                  ,[InformationContent]
                                  ,[MUser]
                              FROM [WarrantAssistant].[dbo].[InformationLog] ";
                sql += "WHERE CONVERT(VARCHAR,Date,112) >='" + GlobalVar.globalParameter.lastTradeDate.ToString("yyyy-MM-dd") + "' ORDER BY MDate DESC";
                DataView dv = DeriLib.Util.ExecSqlQry(sql, GlobalVar.loginSet.warrantassistant45);

                foreach (DataRowView drv in dv)
                {
                    DataRow dr = dt.NewRow();

                    DateTime md = Convert.ToDateTime(drv["MDate"]);
                    dr["時間"] = md.ToString("yyyy/MM/dd HH:mm:ss");
                    dr["類型"] = drv["InformationType"].ToString();
                    dr["內容"] = drv["InformationContent"].ToString();
                    dr["人員"] = drv["MUser"].ToString();

                    dt.Rows.Add(dr);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void toolStripButtonReload_Click(object sender, EventArgs e)
        {
            LoadUltraGrid1();
        }
        private void UltraGrid1_InitializeRow(object sender, InitializeRowEventArgs e)
        {
            string content = e.Row.Cells["內容"].Value.ToString();
            if (content.Contains("申請")|| content.Contains("自動篩選"))
            {
                e.Row.Appearance.ForeColor = Color.White;
                e.Row.Appearance.BackColor = Color.Teal;
            }
            if (content.Contains("刪除"))
            {
                e.Row.Appearance.ForeColor = Color.White;
                e.Row.Appearance.BackColor = Color.IndianRed;
            }


        }


    }
}

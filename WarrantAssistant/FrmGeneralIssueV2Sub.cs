using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Infragistics.Win;
using Infragistics.Win.UltraWinGrid;
using System.Data.SqlClient;
using EDLib.SQL;
using System.Text.RegularExpressions;
using System.Threading;
using System.IO;
using System.Configuration;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop;
using EDLib;

namespace WarrantAssistant
{
    public partial class FrmGeneralIssueV2Sub : Form
    {
        DataTable dt = new DataTable();
        DataTable Data = new DataTable();
        public DateTime lastTradeDate = EDLib.TradeDate.LastNTradeDate(1);

        double[] Moneyness = new double[] { -0.05, 0.1, 0.2,0.3,0.4,0.5};
        double[] Price = new double[] { 0, 1, 2, 3, 99999 };
        public FrmGeneralIssueV2Sub()
        {
            InitializeComponent();
        }

        private void FrmGeneralIssueV2Sub_Load(object sender, EventArgs e)
        {
            LoadData();
            InitialGrid();
        }

        private void LoadData()
        {
            
            string cp = textBox1.Text.Substring(textBox1.Text.Length - 1, 1);
            string[] temp = textBox1.Text.Split('-');
            string uid = temp[0];

            string sql = $@"SELECT  [WID]
                              ,(CASE WHEN [WClass] = 'c' THEN 1- (StrikePrice / UClosePrice) ELSE  (StrikePrice / UClosePrice) -1 END) * -1 AS Moneyness
                              ,[IssuerName]
                              ,[WTheoPrice_IV]
                              ,[TtoM]
	                          ,[IV]
	                          ,[Theta_IV] * [AccReleasingLots] AS Theta
                          FROM [TwData].[dbo].[V_WarrantTrading]
                          WHERE [TDate] = '{lastTradeDate.ToString("yyyyMMdd")}' AND [UID] = '{uid}' and [WClass] = '{cp}' and [TtoM] > 80 
                          and [WTheoPrice_IV] >= 0.6 and [IV] > [HV_60D]";
            Data = MSSQL.ExecSqlQry(sql, GlobalVar.loginSet.twData);

            

        }
        private void InitialGrid()
        {
            dt.Columns.Add("價格/價內外區間", typeof(string));
            dt.Columns.Add("-5~10", typeof(string));
            dt.Columns.Add("10~20", typeof(string));
            dt.Columns.Add("20~30", typeof(string));
            dt.Columns.Add("30~40", typeof(string));
            dt.Columns.Add("40~50", typeof(string));
            dt.Columns.Add("總計", typeof(string));

            ultraGrid1.DataSource = dt;
            UltraGridBand band0 = ultraGrid1.DisplayLayout.Bands[0];


            for (int i = 0; i < Price.Length - 1; i++)
            {
                DataRow dr = dt.NewRow();
                //先給Row Header
                if (i == 0)
                    dr[0] = "<1";
                else if (i == Price.Length - 2)
                    dr[0] = ">3";
                else
                    dr[0] = Price[i].ToString() + "-" + Price[i + 1].ToString();

                double price_low = Price[i];
                double price_high = Price[i + 1];


                for (int j = 0; j < Moneyness.Length - 1; j++)
                {
                    double m_low = Moneyness[j];// 價內往價外 -0.05~0.1
                    double m_high = Moneyness[j + 1];
                    double thetaShare;
                    double warrantShare;
                    DataRow[] DataSelect = Data.Select($@"Moneyness < {m_high} AND Moneyness >= {m_low} AND WTheoPrice_IV < {price_high} AND [WTheoPrice_IV] >= {price_low}");
                    DataRow[] DataSelectKGI = Data.Select($@"Moneyness < {m_high} AND Moneyness >= {m_low} AND WTheoPrice_IV < {price_high} AND [WTheoPrice_IV] >= {price_low} AND IssuerName = '9200'");

                    var thetaAll = DataSelect.AsEnumerable().Sum(x => x.Field<double>("Theta"));
                    var thetaKGI = DataSelectKGI.AsEnumerable().Sum(x => x.Field<double>("Theta"));

                    if (Convert.ToDouble(thetaAll) > 0)
                        thetaShare = Math.Round(Convert.ToDouble(thetaKGI) / Convert.ToDouble(thetaAll), 2) * 100;
                    else
                        thetaShare = 0;

                    if (DataSelect.Length > 0)
                        warrantShare = Math.Round((double)DataSelectKGI.Length / (double)DataSelect.Length, 2) * 100;
                    else
                        warrantShare = 0;

                    dr[j + 1] = thetaShare.ToString() + " / " + warrantShare.ToString();
                }
                DataRow[] DataSelectSum = Data.Select($@" WTheoPrice_IV < {price_high} AND [WTheoPrice_IV] >= {price_low}");
                DataRow[] DataSelectSumKGI = Data.Select($@"WTheoPrice_IV < {price_high} AND [WTheoPrice_IV] >= {price_low} AND IssuerName = '9200'");

                double thetaSumShare;
                double warrantSumShare;
                var thetaSumAll = DataSelectSum.AsEnumerable().Sum(x => x.Field<double>("Theta"));
                var thetaSumKGI = DataSelectSumKGI.AsEnumerable().Sum(x => x.Field<double>("Theta"));

                if (Convert.ToDouble(thetaSumAll) > 0)
                    thetaSumShare = Math.Round(Convert.ToDouble(thetaSumKGI) / Convert.ToDouble(thetaSumAll), 2) * 100;
                else
                    thetaSumShare = 0;

                if (DataSelectSum.Length > 0)
                    warrantSumShare = Math.Round((double)DataSelectSumKGI.Length / (double)DataSelectSum.Length, 2) * 100;
                else
                    warrantSumShare = 0;

                dr[6] = thetaSumShare.ToString() + " / " + warrantSumShare.ToString();
                dt.Rows.Add(dr);

            }

            //建立總計的row
            DataRow dr2 = dt.NewRow();

            dr2[0] = "總計";
            for (int j = 0; j < Moneyness.Length - 1; j++)
            {
                double m_low = Moneyness[j];// 價內往價外 -0.05~0.1
                double m_high = Moneyness[j + 1];
                double thetaShare;
                double warrantShare;
                DataRow[] DataSelect = Data.Select($@"Moneyness < {m_high} AND Moneyness >= {m_low}");
                DataRow[] DataSelectKGI = Data.Select($@"Moneyness < {m_high} AND Moneyness >= {m_low} AND IssuerName = '9200'");

                var thetaAll = DataSelect.AsEnumerable().Sum(x => x.Field<double>("Theta"));
                var thetaKGI = DataSelectKGI.AsEnumerable().Sum(x => x.Field<double>("Theta"));

                if (Convert.ToDouble(thetaAll) > 0)
                    thetaShare = Math.Round(Convert.ToDouble(thetaKGI) / Convert.ToDouble(thetaAll), 2) * 100;
                else
                    thetaShare = 0;

                if (DataSelect.Length > 0)
                    warrantShare = Math.Round((double)DataSelectKGI.Length / (double)DataSelect.Length, 2) * 100;
                else
                    warrantShare = 0;

                dr2[j + 1] = thetaShare.ToString() + " / " + warrantShare.ToString();
            }
            dt.Rows.Add(dr2);

            //建立每個價內外程度的set中Theta最大的權證的 權證理論價/TtoM/IV

            DataRow dr3 = dt.NewRow();
            dr3[0] = "Theta最大的P/TtoM/IV";
            for (int j = 0; j < Moneyness.Length - 1; j++)
            {
                double m_low = Moneyness[j];// 價內往價外 -0.05~0.1
                double m_high = Moneyness[j + 1];

                DataRow[] DataSelect = Data.Select($@"Moneyness < {m_high} AND Moneyness >= {m_low}", "Theta DESC");
                //DataRow[] DataSelect = DataTemp.Select($@"Theta = MAX(Theta)");
                //var DataSelect = DataTemp.AsEnumerable().Max(x => x.Field<DataRow>("Theta"));
                if (DataSelect.Length > 0)
                {
                    double theoP = Math.Round(Convert.ToDouble(DataSelect[0][3].ToString()), 1);
                    double t2m = Convert.ToDouble(DataSelect[0][4].ToString());
                    double iv = Math.Round(Convert.ToDouble(DataSelect[0][5].ToString()), 3) * 100;
                    dr3[j + 1] = theoP.ToString() + "/" + t2m.ToString() + "/" + iv.ToString();
                }
            }
            dt.Rows.Add(dr3);

            this.ultraGrid1.DisplayLayout.Override.RowSelectors = DefaultableBoolean.False;
            this.ultraGrid1.DisplayLayout.AutoFitStyle = AutoFitStyle.ResizeAllColumns;
            this.ultraGrid1.DisplayLayout.Override.RowSizing = RowSizing.AutoFree;

          
            foreach (UltraGridRow dr in ultraGrid1.Rows)
            {
                dr.Activation = Activation.NoEdit;
                if(dr.Cells[0].Text != "總計" && !dr.Cells[0].Text.Contains("Theta"))
                {
                    for(int i = 1; i <= 5; i++)
                    {
                        dr.Cells[i].Appearance.BackColor = Color.Yellow;
                    }
                    dr.Cells[6].Appearance.BackColor = Color.Red;
                }
            }
           
            /*
            band0.Columns["標的代號-名稱"].Format = "N0";
            band0.Columns["WClass"].Format = "N0";
            band0.Columns["分級"].Format = "N0";
            band0.Columns["Theta金額"].Format = "N0";
            band0.Columns["平均VSP"].Format = "N0";
            band0.Columns["檔數市佔"].Format = "N0";
            band0.Columns["Theta市佔率"].Format = "N0";
            band0.Columns["股價累計漲幅"].Format = "N0";
            band0.Columns["好分點買進"].Format = "N0";
            band0.Columns["搶發標的"].Format = "N0";
            band0.Columns["價外25以下"].Format = "N0";
            band0.Columns["價外25以上"].Format = "N0";
            */


            //ultraGrid1.DisplayLayout.Bands[0].Columns["編號"].Width = 75;
            /*
            band0.Columns["組別"].Width = 30;
            band0.Columns["分類"].Width = 40;
            band0.Columns["CP"].Width = 30;
            band0.Columns["TtoM"].Width = 30;
            band0.Columns["基本市佔"].Width = 40;
            band0.Columns["價內外起始"].Width = 40;
            band0.Columns["價內外長度"].Width = 40;
            band0.Columns["交易員"].Width = 50;
            band0.Columns["使用"].Width = 40;

            band0.Columns["組別"].CellAppearance.BackColor = Color.LightGray;
            band0.Columns["分類"].CellAppearance.BackColor = Color.LightGray;
            band0.Columns["CP"].CellAppearance.BackColor = Color.LightGray;
            band0.Columns["TtoM"].CellAppearance.BackColor = Color.LightGray;
            band0.Columns["基本市佔"].CellAppearance.BackColor = Color.LightGray;
            band0.Columns["價內外起始"].CellAppearance.BackColor = Color.LightGray;
            band0.Columns["價內外長度"].CellAppearance.BackColor = Color.LightGray;
            band0.Columns["交易員"].CellAppearance.BackColor = Color.LightBlue;
            */


            // To sort multi-column using SortedColumns property
            // This enables multi-column sorting
            //如果開放sort功能，在更新templist時serialnum會亂掉
            //this.ultraGrid1.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti;

            // It is good practice to clear the sorted columns collection
            //band0.SortedColumns.Clear();

            // You can sort (as well as group rows by) columns by using SortedColumns 
            // property off the band
            //band0.SortedColumns.Add("ContactName", false, false);

            //ultraGrid1.DisplayLayout.Bands[0].Columns["可發行股數"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right;
            //ultraGrid1.DisplayLayout.Bands[0].Columns["截至前一日"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right;
            //ultraGrid1.DisplayLayout.Bands[0].Columns["本日累積發行"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right;
            //ultraGrid1.DisplayLayout.Bands[0].Columns["累計%"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center;


            ultraGrid1.DisplayLayout.Bands[0].Override.HeaderAppearance.TextHAlign = Infragistics.Win.HAlign.Left;
            ultraGrid1.DisplayLayout.Override.RowSelectorNumberStyle = RowSelectorNumberStyle.None;
            //ultraGrid1.DisplayLayout.Bands[0].Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
            //ultraGrid1.DisplayLayout.Bands[0].Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
            //ultraGrid1.DisplayLayout.Bands[0].Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.False;
            //ultraGrid1.DisplayLayout.Bands[0].Columns["確認"].CellActivation = Activation.AllowEdit;
        }

        private void UltraGrid1_DoubleClickCell(object sender, DoubleClickCellEventArgs e)
        {
            int row = e.Cell.Row.Index;
            int column = e.Cell.Column.Index;
            bool formExist = false;
            FrmGeneralIssueV2Sub2 f1;
            if (row < 4 && column > 0 && column < 6)
            {
                foreach (Form iForm in System.Windows.Forms.Application.OpenForms)
                {//只出現一個頁面，不允許開多個視窗
                    if (iForm.GetType() == typeof(FrmGeneralIssueV2Sub2))
                    {
                        f1 = (FrmGeneralIssueV2Sub2)iForm;
                        
                        formExist = true;
                        //MessageBox.Show("存在");
                        f1.textBox1.Text = this.textBox1.Text;
                        //MessageBox.Show($@"ROW:{row}  COLUMN:{column}");
                        //對應價格
                        if (row == 0)
                        {
                            f1.textBox2.Text = Price[0].ToString();
                            f1.textBox3.Text = Price[1].ToString();
                        }
                        else if (row == 1)
                        {
                            f1.textBox2.Text = Price[1].ToString();
                            f1.textBox3.Text = Price[2].ToString();
                        }
                        else if (row == 2)
                        {
                            f1.textBox2.Text = Price[2].ToString();
                            f1.textBox3.Text = Price[3].ToString();
                        }
                        else if (row == 3)
                        {
                            f1.textBox2.Text = Price[3].ToString();
                            f1.textBox3.Text = Price[4].ToString();
                        }
                        else
                        {
                            f1.textBox2.Text = "0";
                            f1.textBox3.Text = "999";
                        }
                        //對應價內外
                        if (column == 0)
                            return;
                        else if (column == 1)
                        {
                            f1.textBox4.Text = "-5";
                            f1.textBox5.Text = "10";
                        }
                        else if (column == 2)
                        {
                            f1.textBox4.Text = "10";
                            f1.textBox5.Text = "20";
                        }
                        else if (column == 3)
                        {
                            f1.textBox4.Text = "20";
                            f1.textBox5.Text = "30";
                        }
                        else if (column == 4)
                        {
                            f1.textBox4.Text = "30";
                            f1.textBox5.Text = "40";
                        }
                        else if (column == 5)
                        {
                            f1.textBox4.Text = "40";
                            f1.textBox5.Text = "50";
                        }
                        else
                        {
                            f1.textBox4.Text = "-5";
                            f1.textBox5.Text = "50";
                        }
                        f1.BringToFront();
                        f1.Show();
                        f1.LoadData();
                    }
                }
                if(formExist == false)
                {
                    f1 = new FrmGeneralIssueV2Sub2();
                    f1.textBox1.Text = this.textBox1.Text;
                    
                    //對應價格
                    if (row == 0)
                    {
                        f1.textBox2.Text = Price[0].ToString();
                        f1.textBox3.Text = Price[1].ToString();
                    }
                    else if (row == 1)
                    {
                        f1.textBox2.Text = Price[1].ToString();
                        f1.textBox3.Text = Price[2].ToString();
                    }
                    else if (row == 2)
                    {
                        f1.textBox2.Text = Price[2].ToString();
                        f1.textBox3.Text = Price[3].ToString();
                    }
                    else if (row == 3)
                    {
                        f1.textBox2.Text = Price[3].ToString();
                        f1.textBox3.Text = Price[4].ToString();
                    }
                    else
                    {
                        f1.textBox2.Text = "0";
                        f1.textBox3.Text = "999";
                    }
                    //對應價內外
                    if (column == 0)
                        return;
                    else if (column == 1)
                    {
                        f1.textBox4.Text = "-5";
                        f1.textBox5.Text = "10";
                    }
                    else if (column == 2)
                    {
                        f1.textBox4.Text = "10";
                        f1.textBox5.Text = "20";
                    }
                    else if (column == 3)
                    {
                        f1.textBox4.Text = "20";
                        f1.textBox5.Text = "30";
                    }
                    else if (column == 4)
                    {
                        f1.textBox4.Text = "30";
                        f1.textBox5.Text = "40";
                    }
                    else if (column == 5)
                    {
                        f1.textBox4.Text = "40";
                        f1.textBox5.Text = "50";
                    }
                    else
                    {
                        f1.textBox4.Text = "-5";
                        f1.textBox5.Text = "50";
                    }
                    f1.Show();
                }
            }
        }

       
    }
}

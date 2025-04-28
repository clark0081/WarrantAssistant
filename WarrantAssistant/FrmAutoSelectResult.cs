#define To39
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Infragistics;
using Infragistics.Win.UltraWinGrid;
using EDLib.SQL;


namespace WarrantAssistant
{

    
    public partial class FrmAutoSelectResult : Form
    {
        #region 機率密度函數
        public const double PI = 3.1415926535897932384626433;
        public static double StandardGaussianProbability(double x)
        {
            double n = 0.0;
            double L = Math.Abs(x);
            const double a1 = 0.31938153;
            const double a2 = -0.356563782;
            const double a3 = 1.781477937;
            const double a4 = -1.821255978;
            const double a5 = 1.330274429;
            const double a6 = 0.2316479;
            double k = 1.0 / (1.0 + a6 * L);
            n = 1 - Math.Exp(-1.0 * L * L * 0.5) * (a1 * k + a2 * k * k + a3 * k * k * k + a4 * k * k * k * k + a5 * k * k * k * k * k) / Math.Sqrt(2.0 * PI);
            if (x < 0)
                n = 1.0 - n;
            return n;
        }
        public static double BlackSholeFormula(string cp,
                                               double spotPrice,
                                               double strikePrice,
                                               double interestRate,
                                               double volatility,
                                               double timeToExpiry,
                                               double costOfCarry)
        {
            double bsprice = 0.0;

            double d1 = (Math.Log(spotPrice / strikePrice) + (costOfCarry + volatility * volatility * 0.5) * timeToExpiry) / (volatility * Math.Sqrt(timeToExpiry));
            double d2 = d1 - volatility * Math.Sqrt(timeToExpiry);
            if (cp == "C")
                bsprice = spotPrice * Math.Exp((costOfCarry - interestRate) * timeToExpiry) * StandardGaussianProbability(d1) - strikePrice * Math.Exp(-1.0 * interestRate * timeToExpiry) * StandardGaussianProbability(d2);
            else
                bsprice = strikePrice * Math.Exp(-1.0 * interestRate * timeToExpiry) * StandardGaussianProbability(-1.0 * d2) - spotPrice * Math.Exp((costOfCarry - interestRate) * timeToExpiry) * StandardGaussianProbability(-1.0 * d1);

            return bsprice;
        }

        public static double WarrantPrice(string cp, double underlyingPrice, double k, double interestRate, double vol, double t, double cr)
        {
            //MessageBox.Show($@"{underlyingPrice} {k} {interestRate} {vol} {t}  {cr}");
            double price = 0.0;
            //double timeToExpiry = t / dayPerYear;
            double timeToExpiry = t;
            price = BlackSholeFormula(cp, underlyingPrice, k, interestRate, vol, timeToExpiry, interestRate) * cr;
            return price;
        }
        #endregion

        #region 計算近似價格
        public static double roundprice(double x)
        {
            if (x < 10)
            {
                double tick = 0.5;
                int up = (int)Math.Floor(x) + 1;
                int down = (int)Math.Floor(x);//8
                if ((double)x % 1 > tick)//ex 3.6
                {
                    if ((double)(up - x) > (double)(x - down - tick))
                        return (double)(down + tick);
                    else
                        return (double)up;
                }
                else//ex3.4
                {
                    if ((double)(down + tick - x) > (double)(x - down))
                        return (double)(down);
                    else
                        return (double)(down + tick);
                }
            }
            else if (x < 50)
                return Math.Round(x);
            else if (x < 100)
            {
                double tick = 5;
                int up = (int)Math.Floor((double)(x / 10)) * 10 + 10;
                int down = (int)Math.Floor((double)(x / 10)) * 10;
                if ((double)x % 10 > tick)//ex 88
                {
                    if ((double)(up - x) > (double)(x - down - tick))
                        return (double)(down + tick);
                    else
                        return (double)up;
                }
                else//ex 84
                {
                    if ((double)(down + tick - x) > (double)(x - down))
                        return (double)(down);
                    else
                        return (double)(down + tick);
                }
            }
            else if (x < 500)
            {
                return Math.Round(x / 10) * 10;
            }
            else if (x < 1000)
            {
                double tick = 50;
                int up = (int)Math.Floor((double)(x / 100)) * 100 + 100;
                int down = (int)Math.Floor((double)(x / 100)) * 100;
                if ((double)x % 100 > tick)//ex 980
                {
                    if ((double)(up - x) > (double)(x - down - tick))
                        return (double)(down + tick);
                    else
                        return (double)up;
                }
                else//ex 84
                {
                    if ((double)(down + tick - x) > (double)(x - down))
                        return (double)(down);
                    else
                        return (double)(down + tick);
                }
            }
            else
                return Math.Round(x / 100) * 100;

        }
        #endregion

        #region 計算履約價近似價格
        public static double roundpriceK(double x)
        {
            if (x < 50)
            {
                double tick = 0.5;
                int up = (int)Math.Floor(x) + 1;//49
                int down = (int)Math.Floor(x);//48
                if ((double)x % 1 > tick)//ex 48.6
                {
                    if ((double)(up - x) > (double)(x - down - tick))
                        return (double)(down + tick);
                    else
                        return (double)up;
                }
                else//ex3.4
                {
                    if ((double)(down + tick - x) > (double)(x - down))
                        return (double)(down);
                    else
                        return (double)(down + tick);
                }
            }
            else if (x < 100)
                return Math.Round(x);
            else if (x < 500)
            {
                double tick = 5;
                int up = (int)Math.Floor((double)(x / 10)) * 10 + 10;
                int down = (int)Math.Floor((double)(x / 10)) * 10;
                if ((double)x % 10 > tick)//ex 498
                {
                    if ((double)(up - x) > (double)(x - down - tick))
                        return (double)(down + tick);
                    else
                        return (double)up;
                }
                else//ex 
                {
                    if ((double)(down + tick - x) > (double)(x - down))
                        return (double)(down);
                    else
                        return (double)(down + tick);
                }
            }
            else if (x < 1000)
            {
                return Math.Round(x / 10) * 10;
            }
            else 
            {
                double tick = 50;
                int up = (int)Math.Floor((double)(x / 100)) * 100 + 100;
                int down = (int)Math.Floor((double)(x / 100)) * 100;
                if ((double)x % 100 > tick)//ex 980
                {
                    if ((double)(up - x) > (double)(x - down - tick))
                        return (double)(down + tick);
                    else
                        return (double)up;
                }
                else//ex 84
                {
                    if ((double)(down + tick - x) > (double)(x - down))
                        return (double)(down);
                    else
                        return (double)(down + tick);
                }
            }
            

        }
        #endregion
        public static CMADODB5.CMConnection cn = new CMADODB5.CMConnection();
        public static string arg = "5"; //%
        public static string srvLocation = "10.60.0.191";
        public static string cnPort = "";
        DateTime today = DateTime.Today;
        public string userID = GlobalVar.globalParameter.userID;
        public Dictionary<string, string> EngToChinese = new Dictionary<string, string>();
        public Dictionary<string, List<string>> ShowLogicResult = new Dictionary<string, List<string>>();
        public Dictionary<string, double> IssueCredit = new Dictionary<string, double>();

        public bool LogicAND(DataRow data, DataRow setting, List<string> And)
        {
            bool result = true;
            foreach(string column in And)
            {
                try
                {
                    int BigOrless = CompareTable[column];
                    bool temp;
                    
                    if (BigOrless == 1)//>
                    {
                        double d = Convert.ToDouble(data[column].ToString());//data
                        double s = Convert.ToDouble(setting[column].ToString());//setting
                        if (d > s)
                            temp = true;
                        else
                            temp = false;
                    }
                    else if (BigOrless == 2)
                    {
                        temp = false;
                        double lastOptionAvailable = Convert.ToDouble(data["LastOptionAvailable"].ToString());
                        double optionAvailable = Convert.ToDouble(data["OptionAvailable"].ToString());
                        double min = lastOptionAvailable <= optionAvailable ? lastOptionAvailable : optionAvailable;
                        if (min <= 10 && optionAvailable > Convert.ToDouble(setting["OptionRelease"].ToString()))
                            temp = true;
                        /*
                        if (Convert.ToDouble(data["LastOptionAvailable"].ToString()) < 1)
                        {
                            if (Convert.ToDouble(data["OptionAvailable"].ToString()) > Convert.ToDouble(setting["OptionRelease"].ToString()))
                                temp = true;
                        }
                        */
                        
                    }
                    else if (BigOrless == 3)
                    {
                        temp = false;
                        //為可解禁發行
                        double d = Convert.ToDouble(data[column].ToString());
                        if (d == 1)
                        {
                            temp = true;
                        }
                    }
                    else
                    {
                        double d = Convert.ToDouble(data[column].ToString());//data
                        double s = Convert.ToDouble(setting[column].ToString());//setting
                        if (d < s)
                            temp = true;
                        else
                            temp = false;
                    }
                    result = result && temp;
                }
                catch
                {
                    MessageBox.Show("and  "+column);
                    result = true;
                }
            }
            if (And.Count == 0)
                return false;
            
            return result;
        }
        public bool LogicOR(DataRow data, DataRow setting, List<string> Or)
        {
            //MessageBox.Show($@"in logicOR and lenth{Or.Count}");
            bool result = false;
            foreach (string column in Or)
            {
                try
                {
                    int BigOrless = CompareTable[column];
                    bool temp;
                    if (BigOrless == 1)//>
                    {
                        double d = Convert.ToDouble(data[column].ToString());//data
                        double s = Convert.ToDouble(setting[column].ToString());//setting
                        if (d > s)
                            temp = true;
                        else
                            temp = false;
                    }
                    else if (BigOrless == 2)
                    {
                        temp = false;
                        double lastOptionAvailable = Convert.ToDouble(data["LastOptionAvailable"].ToString());
                        double optionAvailable = Convert.ToDouble(data["OptionAvailable"].ToString());
                        double min = lastOptionAvailable <= optionAvailable ? lastOptionAvailable : optionAvailable;
                        if (min <= 10 && optionAvailable > Convert.ToDouble(setting["OptionRelease"].ToString()))
                            temp = true;
                        /*
                        if (Convert.ToDouble(data["LastOptionAvailable"].ToString()) < 1)
                        {
                            if (Convert.ToDouble(data["OptionAvailable"].ToString()) > Convert.ToDouble(setting["OptionRelease"].ToString()))
                                temp = true;
                        }
                        */
                    }
                    else if (BigOrless == 3)
                    {
                        temp = false;
                        //為可解禁發行
                        double d = Convert.ToDouble(data[column].ToString());
                        if (d == 1)
                        {
                            temp = true;
                        }
                        
                    }
                    else
                    {
                        double d = Convert.ToDouble(data[column].ToString());//data
                        double s = Convert.ToDouble(setting[column].ToString());//setting
                        if (d < s)
                            temp = true;
                        else
                            temp = false;
                    }
                    result = result || temp;
                }
                catch
                {
                    MessageBox.Show("or  "+column);
                    result = false;
                }
            }
            if (Or.Count == 0)
                return false;
            return result;
        }
        public List<string> LogicORANDResult(DataRow data, DataRow setting, List<string> Or, List<string> And)
        {
            List<string> result = new List<string>();
            foreach (string column in Or)
            {
                
                int BigOrless = CompareTable[column];
                    
                if (BigOrless == 1)//>
                {
                    double d = Convert.ToDouble(data[column].ToString());//data
                    double s = Convert.ToDouble(setting[column].ToString());//setting
                    if (d > s)
                    {
                        string str = "Y," + "(OR)" + EngToChinese[column] + "," + data[column].ToString() + ","+">"+"," + setting[column].ToString();
                        result.Add(str);
                    }
                    else
                    {
                        string str = "N," + "(OR)" + EngToChinese[column] + "," + data[column].ToString() + "," + ">" + "," + setting[column].ToString();
                        result.Add(str);
                    }
                        
                }
                else if (BigOrless == 2)
                {
                    double lastOptionAvailable = Convert.ToDouble(data["LastOptionAvailable"].ToString());
                    double optionAvailable = Convert.ToDouble(data["OptionAvailable"].ToString());
                    double min = lastOptionAvailable <= optionAvailable ? lastOptionAvailable : optionAvailable;
                    if (min <= 10 && optionAvailable > Convert.ToDouble(setting["OptionRelease"].ToString()))
                    {
                        string str = "Y," + "(OR)" + EngToChinese["OptionRelease"] + "," + min + "," + "<=" + "," + "10";
                        string str2 = "Y," + "(OR)" + "今日可發行" + "," + optionAvailable + "," + ">" + "," + Convert.ToDouble(setting["OptionRelease"].ToString());
                        result.Add(str);
                        result.Add(str2);
                    }
                    else if (min<=10 && optionAvailable <= Convert.ToDouble(setting["OptionRelease"].ToString()))
                    {
                        string str = "Y," + "(OR)" + EngToChinese["OptionRelease"] + "," + min + "," + "<=" + "," + "10";
                        string str2 = "N," + "(OR)" + "今日可發行" + "," + optionAvailable + "," + ">" + "," + Convert.ToDouble(setting["OptionRelease"].ToString());
                        result.Add(str);
                        result.Add(str2);
                    }
                    else if (min > 10)
                    {
                        string str = "N," + "(OR)" + EngToChinese["OptionRelease"] + "," + min + "," + "<=" + "," + "10";
                        result.Add(str);
                    }
                    /*
                    if (Convert.ToDouble(data["LastOptionAvailable"].ToString()) < 1)
                    {
                        if (Convert.ToDouble(data["DiffOptionAvailable"].ToString()) >= Convert.ToDouble(setting["OptionRelease"].ToString()))
                        {
                            string str = "Y," + "(OR)" + EngToChinese["OptionRelease"] + "," + data["DiffOptionAvailable"].ToString() + "," + ">=" + "," + setting["OptionRelease"].ToString();
                            result.Add(str);
                            //result.Add(EngToChinese[column]);
                        }
                        else
                        {
                            string str = "N," + "(OR)" + EngToChinese["OptionRelease"] + "," + data["DiffOptionAvailable"].ToString() + "," + ">" + "," + setting["OptionRelease"].ToString();
                            result.Add(str);
                        }
                    }
                    */

                }
                else if (BigOrless == 3)
                {
                    double d = Convert.ToDouble(data[column].ToString());
                    if(d == 1)
                    {
                        string str = "Y," + "(OR)" + EngToChinese[column] + "," + data[column].ToString() + "," + ">=" + "," + "1";
                        result.Add(str);
                    }
                    else
                    {
                        string str = "N," + "(OR)" + EngToChinese[column] + "," + data[column].ToString() + "," + ">=" + "," + "1";
                        result.Add(str);
                    }
                    
                }
                else
                {
                    double d = Convert.ToDouble(data[column].ToString());//data
                    double s = Convert.ToDouble(setting[column].ToString());//setting
                    if (d < s) {
                        string str = "Y," + "(OR)" + EngToChinese[column] + "," + data[column].ToString() + "," + "<" + "," + setting[column].ToString();
                        result.Add(str);
                    }
                    else
                    {
                        string str = "N," + "(OR)" +EngToChinese[column] + "," + data[column].ToString() + "," + "<" + "," + setting[column].ToString();
                        result.Add(str);
                    }
                }

            }
            foreach (string column in And)
            {

                int BigOrless = CompareTable[column];

                if (BigOrless == 1)//>
                {
                    double d = Convert.ToDouble(data[column].ToString());//data
                    double s = Convert.ToDouble(setting[column].ToString());//setting
                    if (d > s)
                    {
                        string str = "Y," + "(AND)" + EngToChinese[column] + "," + data[column].ToString() + "," + ">" + "," + setting[column].ToString();
                        result.Add(str);
                    }
                    else
                    {
                        string str = "N," + "(AND)" + EngToChinese[column] + "," + data[column].ToString() + "," + ">" + "," + setting[column].ToString();
                        result.Add(str);
                    }

                }
                else if (BigOrless == 2)
                {
                    /*
                    if (Convert.ToDouble(data["LastOptionAvailable"].ToString()) < 1)
                    {
                        if (Convert.ToDouble(data["OptionAvailable"].ToString()) >= Convert.ToDouble(setting["OptionRelease"].ToString()))
                        {
                            string str = "Y," + "(AND)" + EngToChinese["OptionRelease"] + "," + data["OptionAvailable"].ToString() + "," + ">=" + "," + setting["OptionRelease"].ToString();
                            result.Add(str);
                            //result.Add(EngToChinese[column]);
                        }
                        else
                        {
                            string str = "N," + "(AND)" + EngToChinese["OptionRelease"] + "," + data["OptionAvailable"].ToString() + "," + ">" + "," + setting["OptionRelease"].ToString();
                            result.Add(str);
                        }
                    }
                    */
                    double lastOptionAvailable = Convert.ToDouble(data["LastOptionAvailable"].ToString());
                    double optionAvailable = Convert.ToDouble(data["OptionAvailable"].ToString());
                    double min = lastOptionAvailable <= optionAvailable ? lastOptionAvailable : optionAvailable;
                    if (min <= 10 && optionAvailable > Convert.ToDouble(setting["OptionRelease"].ToString()))
                    {
                        string str = "Y," + "(AND)" + EngToChinese["OptionRelease"] + "," + min + "," + "<=" + "," + "10";
                        string str2 = "Y," + "(AND)" + "今日可發行" + "," + optionAvailable + "," + ">" + "," + Convert.ToDouble(setting["OptionRelease"].ToString());
                        result.Add(str);
                        result.Add(str2);
                    }
                    else if (min <= 10 && optionAvailable <= Convert.ToDouble(setting["OptionRelease"].ToString()))
                    {
                        string str = "Y," + "(AND)" + EngToChinese["OptionRelease"] + "," + min + "," + "<=" + "," + "10";
                        string str2 = "N," + "(AND)" + "今日可發行" + "," + optionAvailable + "," + ">" + "," + Convert.ToDouble(setting["OptionRelease"].ToString());
                        result.Add(str);
                        result.Add(str2);
                    }
                    else if (min > 10)
                    {
                        string str = "N," + "(OR)" + EngToChinese["OptionRelease"] + "," + min + "," + "<=" + "," + "10";
                        result.Add(str);
                    }
                }
                else if (BigOrless == 3)
                {
                    double d = Convert.ToDouble(data[column].ToString());
                    if (d == 1)
                    {
                        string str = "Y," + "(AND)" + EngToChinese[column] + "," + data[column].ToString() + "," + ">=" + "," + "1";
                        result.Add(str);
                    }
                    else
                    {
                        string str = "N," + "(AND)" + EngToChinese[column] + "," + data[column].ToString() + "," + ">=" + "," + "1";
                        result.Add(str);
                    }

                }
                else
                {
                    double d = Convert.ToDouble(data[column].ToString());//data
                    double s = Convert.ToDouble(setting[column].ToString());//setting
                    if (d < s)
                    {
                        string str = "Y," + "(AND)" + EngToChinese[column] + "," + data[column].ToString() + "," + "<" + "," + setting[column].ToString();
                        result.Add(str);
                    }
                    else
                    {
                        string str = "N," + "(AND)" + EngToChinese[column] + "," + data[column].ToString() + "," + "<" + "," + setting[column].ToString();
                        result.Add(str);
                    }
                }

            }
            return result;
        }

        public bool Result(DataRow data,DataRow setting,List<string> And,List<string> Or,int AndOrSet)
        {
            if (AndOrSet == 1)
            {
                bool and = LogicAND(data, setting, And);
                bool or = LogicOR(data, setting, Or);
                bool result = (and && or);
                return result;
            }
            else
            {
                bool and = LogicAND(data, setting, And);
                bool or = LogicOR(data, setting, Or);
                bool result = (and || or);
                return result;
            }
        }
        private System.Data.DataTable dataTable1 = new System.Data.DataTable();
        private System.Data.DataTable dataTable2 = new System.Data.DataTable();
        private System.Data.DataTable dataTable3 = new System.Data.DataTable();
        private System.Data.DataTable dataTable4 = new System.Data.DataTable();//1、4為CP分開的表
        //用來儲存參數要比大或比小
        private Dictionary<string, int> CompareTable = new Dictionary<string, int>();
        private Dictionary<string, string> underlying2trader = new Dictionary<string, string>();
        List<string> VarCAnd = new List<string>();
        List<string> VarCOr = new List<string>();

        List<string> VarPAnd = new List<string>();
        List<string> VarPOr = new List<string>();

        private bool isEdit1 = false;
        private bool isEdit2 = false;
        private bool isEdit3 = false;
        private bool isEdit4 = false;
        public FrmAutoSelectResult()
        {
            InitializeComponent();
        }

        private void FrmAutoSelectResult_Load(object sender, EventArgs e)
        {
            foreach (var item in GlobalVar.globalParameter.traders)
                comboBox1.Items.Add(item);
            comboBox1.Items.Add("All");
            comboBox1.Text = GlobalVar.globalParameter.userID;
            //0:參數比較小於
            //1:參數表較大於
            //其他:其他情況
            CompareTable.Add("BrokerPL_Month",0);
            //CompareTable.Add("Profit_Month",0);
            CompareTable.Add("OptionAvailable",0);
            CompareTable.Add("OptionRelease",2);
            CompareTable.Add("AllowIssue", 3);
            CompareTable.Add("RiseUp_3Days",1);
            CompareTable.Add("DropDown_3Days",0);
            CompareTable.Add("Theta_EndDate", 1);
            //CompareTable.Add("ThetaIV_WeekDelta",1);
            CompareTable.Add("Med_HV60D_VolRatio",1);
            CompareTable.Add("Theta_Days",0);
            CompareTable.Add("AppraisalRank",0);
            CompareTable.Add("FinancingRatio",0);
            CompareTable.Add("CallMarketShare",0);
            CompareTable.Add("PutMarketShare",0);
            CompareTable.Add("CallDensity",0);
            CompareTable.Add("PutDensity",0);
            CompareTable.Add("KgiYuanCallDensity", 0);
            CompareTable.Add("KgiYuanPutDensity", 0);
            CompareTable.Add("MarketAmtChg5Days", 1);
            CompareTable.Add("AlphaTheta", 1);
            CompareTable.Add("AvgAlphaThetaCost", 1);
            //datatable1
            dataTable1.Columns.Add("Logic", typeof(string));
            dataTable1.Columns.Add("BrokerPL_Month",typeof(string));
            //dataTable1.Columns.Add("Profit_Month", typeof(string));
            dataTable1.Columns.Add("OptionAvailable", typeof(string));
            dataTable1.Columns.Add("OptionRelease", typeof(string));
            dataTable1.Columns.Add("AllowIssue", typeof(string));
            dataTable1.Columns.Add("RiseUp_3Days", typeof(string));
            dataTable1.Columns.Add("DropDown_3Days", typeof(string));
            dataTable1.Columns.Add("Theta_EndDate", typeof(string));
            dataTable1.Columns.Add("MarketAmtChg5Days", typeof(string));
            //dataTable1.Columns.Add("ThetaIV_WeekDelta", typeof(string));
            dataTable1.Columns.Add("Med_HV60D_VolRatio", typeof(string));

            
            dataTable1.Columns.Add("AlphaTheta", typeof(string));
            dataTable1.Columns.Add("AvgAlphaThetaCost", typeof(string));

            dataTable1.Columns.Add("Theta_Days", typeof(string));
            dataTable1.Columns.Add("AppraisalRank", typeof(string));
            dataTable1.Columns.Add("FinancingRatio", typeof(string));
            dataTable1.Columns.Add("CallMarketShare", typeof(string));
            //dataTable1.Columns.Add("PutMarketShare", typeof(string));
            dataTable1.Columns.Add("CallDensity", typeof(string));
            //dataTable1.Columns.Add("PutDensity", typeof(string));
            dataTable1.Columns.Add("KgiYuanCallDensity", typeof(string));
            //dataTable1.Columns.Add("KgiYuanPutDensity", typeof(string));
            //datatable2
            dataTable2.Columns.Add("Rank", typeof(string));
            dataTable2.Columns.Add("K",typeof(string));
            dataTable2.Columns.Add("T", typeof(string));
            dataTable2.Columns.Add("HV", typeof(string));
            dataTable2.Columns.Add("IV", typeof(string));
            dataTable2.Columns.Add("IssueLots", typeof(string));
            dataTable2.Columns.Add("CP", typeof(string));
            dataTable2.Columns.Add("Type", typeof(string));
            dataTable2.Columns.Add("Price", typeof(string));
            dataTable2.Columns.Add("ResetR", typeof(string));
            dataTable2.Columns.Add("CR", typeof(string));
            //datatable3
            dataTable3.Columns.Add("編號", typeof(int));
            dataTable3.Columns.Add("確認", typeof(bool));
            dataTable3.Columns.Add("標的代號",typeof(string));
            dataTable3.Columns.Add("標的名稱", typeof(string));
            dataTable3.Columns.Add("昨日收盤價", typeof(double));
            dataTable3.Columns.Add("CP", typeof(string));
            dataTable3.Columns.Add("履約價", typeof(double));
            dataTable3.Columns.Add("期間(月)", typeof(int));
            dataTable3.Columns.Add("Vol(建議Vol)", typeof(double));
            dataTable3.Columns.Add("發行張數", typeof(int));
            dataTable3.Columns.Add("類型", typeof(string));
            dataTable3.Columns.Add("重設比", typeof(string));
            dataTable3.Columns.Add("發行價", typeof(double));
            dataTable3.Columns.Add("行使比例", typeof(double));
            dataTable3.Columns.Add("額度", typeof(double));
            dataTable3.Columns.Add("Theta/VSP", typeof(string));
            dataTable3.Columns.Add("備註", typeof(string));
            //datatable4
            dataTable4.Columns.Add("Logic", typeof(string));
            dataTable4.Columns.Add("BrokerPL_Month", typeof(string));
            //dataTable4.Columns.Add("Profit_Month", typeof(string));
            dataTable4.Columns.Add("OptionAvailable", typeof(string));
            dataTable4.Columns.Add("OptionRelease", typeof(string));
            dataTable4.Columns.Add("AllowIssue", typeof(string));
            dataTable4.Columns.Add("RiseUp_3Days", typeof(string));
            dataTable4.Columns.Add("DropDown_3Days", typeof(string));
            dataTable4.Columns.Add("Theta_EndDate", typeof(string));
            dataTable4.Columns.Add("MarketAmtChg5Days", typeof(string));
            //dataTable4.Columns.Add("ThetaIV_WeekDelta", typeof(string));
            dataTable4.Columns.Add("Med_HV60D_VolRatio", typeof(string));
            dataTable4.Columns.Add("AlphaTheta", typeof(string));
            dataTable4.Columns.Add("AvgAlphaThetaCost", typeof(string));
            dataTable4.Columns.Add("Theta_Days", typeof(string));
            dataTable4.Columns.Add("AppraisalRank", typeof(string));
            dataTable4.Columns.Add("FinancingRatio", typeof(string));
            //dataTable4.Columns.Add("CallMarketShare", typeof(string));
            dataTable4.Columns.Add("PutMarketShare", typeof(string));
            //dataTable4.Columns.Add("CallDensity", typeof(string));
            dataTable4.Columns.Add("PutDensity", typeof(string));
            //dataTable4.Columns.Add("KgiYuanCallDensity", typeof(string));
            dataTable4.Columns.Add("KgiYuanPutDensity", typeof(string));
            //
            LoadEngToChinese();


            LoadDataUlt1();
            LoadDataUlt4();
            LoadDataUlt2();
            LoadDataUlt3();
            InitialGrid1();
            InitialGrid4();  
            InitialGrid2();
            InitialGrid3();

            LoadIssueCredit();
            LoadUnderlying_Trader();
           
            //AutoSelect();
        }
        private void LoadAll()
        {
            LoadDataUlt1();
            LoadDataUlt4();
            LoadDataUlt2();
            //LoadDataUlt3();
        }
        public void LoadUnderlying_Trader()
        {
            underlying2trader.Clear();
            string getUnderlyingTraderStr = $@"SELECT  [UID]
                                                      ,[TraderAccount]
                                                  FROM [TwData].[dbo].[Underlying_Trader]
                                                  WHERE LEN(UID) < 5 AND LEFT(UID, 2) <> '00'";
            DataTable underlyingTrader = MSSQL.ExecSqlQry(getUnderlyingTraderStr, GlobalVar.loginSet.twData);
            foreach (DataRow dr in underlyingTrader.Rows)
            {
                string uid = dr["UID"].ToString();
                string trader = dr["TraderAccount"].ToString();
                underlying2trader.Add(uid, trader.PadLeft(7, '0'));
            }
        }
        public void LoadIssueCredit()
        {
            IssueCredit.Clear();
            string sql  = $@"SELECT [UID], [CanIssue]
                          FROM [WarrantAssistant].[dbo].[WarrantUnderlyingCreditNew]
                          WHERE [UpdateTime] > '{DateTime.Today.ToString("yyyyMMdd")}'";
            DataTable dt = MSSQL.ExecSqlQry(sql, GlobalVar.loginSet.warrantassistant45);
            foreach(DataRow dr in dt.Rows)
            {
                string uid = dr["UID"].ToString();
                double credit = Math.Round(Convert.ToDouble(dr["CanIssue"].ToString()));
                IssueCredit.Add(uid, credit);
            }
        }
        private void LoadEngToChinese()
        {
            EngToChinese.Add("BrokerPL_Month", "一個月brokerPL(%)");
            //EngToChinese.Add("Profit_Month", "標的月損益(仟)");
            EngToChinese.Add("OptionAvailable", "市場剩餘額度(檔)");
            EngToChinese.Add("OptionRelease", "有額度釋出");
            EngToChinese.Add("AllowIssue", "解禁可發行");
            EngToChinese.Add("RiseUp_3Days", "股票3日漲幅");
            EngToChinese.Add("DropDown_3Days", "股票3日跌幅");
            EngToChinese.Add("Theta_EndDate", "市場Theta IV");
            //EngToChinese.Add("ThetaIV_WeekDelta", "市場Theta IV 金額週變化(部位)");
            EngToChinese.Add("Med_HV60D_VolRatio", "市場Med vol/ HV_60D");
            EngToChinese.Add("Theta_Days", "Theta天數");
            EngToChinese.Add("AppraisalRank", "評鑑權證比重排名");
            EngToChinese.Add("FinancingRatio", "融資使用率(%)");
            EngToChinese.Add("CallMarketShare", "C有效檔數市佔(%)");
            EngToChinese.Add("PutMarketShare", "P有效檔數市佔(%)");
            EngToChinese.Add("CallDensity", "C發行密度");
            EngToChinese.Add("PutDensity", "P發行密度");
            EngToChinese.Add("KgiYuanCallDensity", "凱基-元大C發行密度");
            EngToChinese.Add("KgiYuanPutDensity", "凱基-元大P發行密度");
            EngToChinese.Add("MarketAmtChg5Days", "市場部位近五日變化(萬)");
            EngToChinese.Add("AlphaTheta", "市場超額利潤(仟)");
            EngToChinese.Add("AvgAlphaThetaCost", "平均每檔權證超額利潤(元)");
        }
        private void LoadDataUlt1()
        {
            dataTable1.Clear();
            VarCOr.Clear();
            VarCAnd.Clear();
            string traderID = comboBox1.Text;

#if !To39
            string sql = $@"SELECT [MatchValue]
                          FROM [newEDIS].[dbo].[OptionAutoSelectSettings]
                          WHERE [TraderID] ='{traderID}' AND [TableType] ='C_MatchAndOr' AND [MatchKey] ='AND'";
            DataTable dt = EDLib.SQL.MSSQL.ExecSqlQry(sql, GlobalVar.loginSet.newEDIS);
#else
            string sql = $@"SELECT [MatchValue]
                          FROM [WarrantAssistant].[dbo].[OptionAutoSelectSettings]
                          WHERE [TraderID] ='{traderID}' AND [TableType] ='C_MatchAndOr' AND [MatchKey] ='AND'";
            DataTable dt = EDLib.SQL.MSSQL.ExecSqlQry(sql, GlobalVar.loginSet.warrantassistant45);
#endif
            DataRow drtemp = dataTable1.NewRow();
            drtemp["Logic"] = "AND(V)";
            foreach (DataRow dr in dt.Rows)
            {
                string columnname = dr["MatchValue"].ToString();
                drtemp[columnname] = 'V';
                VarCAnd.Add(columnname);
            }
            
            dataTable1.Rows.Add(drtemp);
#if !To39
            string sql2 = $@"SELECT [MatchValue]
                          FROM [newEDIS].[dbo].[OptionAutoSelectSettings]
                          WHERE [TraderID] ='{traderID}' AND [TableType] ='C_MatchAndOr' AND [MatchKey] ='OR'";
            dt = EDLib.SQL.MSSQL.ExecSqlQry(sql2, GlobalVar.loginSet.newEDIS);
#else
            string sql2 = $@"SELECT [MatchValue]
                          FROM [WarrantAssistant].[dbo].[OptionAutoSelectSettings]
                          WHERE [TraderID] ='{traderID}' AND [TableType] ='C_MatchAndOr' AND [MatchKey] ='OR'";
            dt = EDLib.SQL.MSSQL.ExecSqlQry(sql2, GlobalVar.loginSet.warrantassistant45);
#endif
            drtemp = dataTable1.NewRow();
            drtemp["Logic"] = "OR(O)";
            foreach (DataRow dr in dt.Rows)
            {
                string columnname = dr["MatchValue"].ToString();
                drtemp[columnname] = 'O';
                VarCOr.Add(columnname);
            }
            dataTable1.Rows.Add(drtemp);
            ultraGrid1.DataSource = dataTable1;
#if !To39
            string sql3 = $@"SELECT [MatchValue]
                          FROM [newEDIS].[dbo].[OptionAutoSelectSettings]
                          WHERE [TraderID] ='{traderID}' AND [TableType] ='C_MatchAndOr' AND [MatchKey] ='IsClicked'";
            dt = EDLib.SQL.MSSQL.ExecSqlQry(sql3, GlobalVar.loginSet.newEDIS);
#else
            string sql3 = $@"SELECT [MatchValue]
                          FROM [WarrantAssistant].[dbo].[OptionAutoSelectSettings]
                          WHERE [TraderID] ='{traderID}' AND [TableType] ='C_MatchAndOr' AND [MatchKey] ='IsClicked'";
            dt = EDLib.SQL.MSSQL.ExecSqlQry(sql3, GlobalVar.loginSet.warrantassistant45);
#endif
            foreach (DataRow dr in dt.Rows)
            {
                if (dr["MatchValue"].ToString() == "AND")
                    radioButton1.Checked = true;
                else if(dr["MatchValue"].ToString() == "OR")
                    radioButton2.Checked = true;
            }
        }
        private void LoadDataUlt4()
        {
            dataTable4.Clear();
            VarPAnd.Clear();
            VarPOr.Clear();
            string traderID = comboBox1.Text;
#if !To39
            string sql = $@"SELECT [MatchValue]
                          FROM [newEDIS].[dbo].[OptionAutoSelectSettings]
                          WHERE [TraderID] ='{traderID}' AND [TableType] ='P_MatchAndOr' AND [MatchKey] ='AND'";
            DataTable dt = EDLib.SQL.MSSQL.ExecSqlQry(sql, GlobalVar.loginSet.newEDIS);
#else
            string sql = $@"SELECT [MatchValue]
                          FROM [WarrantAssistant].[dbo].[OptionAutoSelectSettings]
                          WHERE [TraderID] ='{traderID}' AND [TableType] ='P_MatchAndOr' AND [MatchKey] ='AND'";
            DataTable dt = EDLib.SQL.MSSQL.ExecSqlQry(sql, GlobalVar.loginSet.warrantassistant45);
#endif
            DataRow drtemp = dataTable4.NewRow();
            drtemp["Logic"] = "AND(V)";
            foreach (DataRow dr in dt.Rows)
            {
                string columnname = dr["MatchValue"].ToString();
                drtemp[columnname] = 'V';
                VarPAnd.Add(columnname);
            }
            dataTable4.Rows.Add(drtemp);
#if !To39
            string sql2 = $@"SELECT [MatchValue]
                          FROM [newEDIS].[dbo].[OptionAutoSelectSettings]
                          WHERE [TraderID] ='{traderID}' AND [TableType] ='P_MatchAndOr' AND [MatchKey] ='OR'";
            dt = EDLib.SQL.MSSQL.ExecSqlQry(sql2, GlobalVar.loginSet.newEDIS);
#else
            string sql2 = $@"SELECT [MatchValue]
                          FROM [WarrantAssistant].[dbo].[OptionAutoSelectSettings]
                          WHERE [TraderID] ='{traderID}' AND [TableType] ='P_MatchAndOr' AND [MatchKey] ='OR'";
            dt = EDLib.SQL.MSSQL.ExecSqlQry(sql2, GlobalVar.loginSet.warrantassistant45);
#endif
            drtemp = dataTable4.NewRow();
            drtemp["Logic"] = "OR(O)";
            foreach (DataRow dr in dt.Rows)
            {
                string columnname = dr["MatchValue"].ToString();
                drtemp[columnname] = 'O';
                VarPOr.Add(columnname);
            }
            dataTable4.Rows.Add(drtemp);
            ultraGrid4.DataSource = dataTable4;
#if !To39
            string sql3 = $@"SELECT [MatchValue]
                          FROM [newEDIS].[dbo].[OptionAutoSelectSettings]
                          WHERE [TraderID] ='{traderID}' AND [TableType] ='P_MatchAndOr' AND [MatchKey] ='IsClicked'";
            dt = EDLib.SQL.MSSQL.ExecSqlQry(sql3, GlobalVar.loginSet.newEDIS);
#else
            string sql3 = $@"SELECT [MatchValue]
                          FROM [WarrantAssistant].[dbo].[OptionAutoSelectSettings]
                          WHERE [TraderID] ='{traderID}' AND [TableType] ='P_MatchAndOr' AND [MatchKey] ='IsClicked'";
            dt = EDLib.SQL.MSSQL.ExecSqlQry(sql3, GlobalVar.loginSet.warrantassistant45);
#endif
            foreach (DataRow dr in dt.Rows)
            {
                if (dr["MatchValue"].ToString() == "AND")
                    radioButton3.Checked = true;
                else if (dr["MatchValue"].ToString() == "OR")
                    radioButton4.Checked = true;
            }
        }
        private void LoadDataUlt2()
        {
            try
            {
                dataTable2.Clear();
                string traderID = comboBox1.Text;
                for (int i = 1; i < 5; i++)
                {
                    string strmatch = "C_MatchIssue" + i.ToString();
#if !To39
                    string sql = $@"SELECT [TraderID], [TableType], [MatchKey], [MatchValue]
                          FROM [newEDIS].[dbo].[OptionAutoSelectSettings]
                          WHERE [TraderID] ='{traderID}' AND [TableType] ='{strmatch}'";
                    DataRow drtemp = dataTable2.NewRow();
                    DataTable dt = EDLib.SQL.MSSQL.ExecSqlQry(sql, GlobalVar.loginSet.newEDIS);
#else
                    string sql = $@"SELECT [TraderID], [TableType], [MatchKey], [MatchValue]
                          FROM [WarrantAssistant].[dbo].[OptionAutoSelectSettings]
                          WHERE [TraderID] ='{traderID}' AND [TableType] ='{strmatch}'";
                    DataRow drtemp = dataTable2.NewRow();
                    DataTable dt = EDLib.SQL.MSSQL.ExecSqlQry(sql, GlobalVar.loginSet.warrantassistant45);
#endif
                    foreach (DataRow dr in dt.Rows)
                    {
                        string columnname = dr["MatchKey"].ToString();
                        string columnvalue = dr["MatchValue"].ToString();
                        drtemp[columnname] = columnvalue;
                    }
                    drtemp["Rank"] = i.ToString();
                    drtemp["HV"] = "建議Vol";
                    drtemp["IV"] = "建議Vol";
                    dataTable2.Rows.Add(drtemp);
                }
                for (int i = 1; i < 5; i++)
                {
                    string strmatch = "P_MatchIssue" + i.ToString();
                    DataRow drtemp = dataTable2.NewRow();
#if !To39
                    string sql2 = $@"SELECT [TraderID], [TableType], [MatchKey], [MatchValue]
                          FROM [newEDIS].[dbo].[OptionAutoSelectSettings]
                          WHERE [TraderID] ='{traderID}' AND [TableType] ='{strmatch}'";
                    DataTable dt2 = EDLib.SQL.MSSQL.ExecSqlQry(sql2, GlobalVar.loginSet.newEDIS);
#else
                    string sql2 = $@"SELECT [TraderID], [TableType], [MatchKey], [MatchValue]
                          FROM [WarrantAssistant].[dbo].[OptionAutoSelectSettings]
                          WHERE [TraderID] ='{traderID}' AND [TableType] ='{strmatch}'";
                    DataTable dt2 = EDLib.SQL.MSSQL.ExecSqlQry(sql2, GlobalVar.loginSet.warrantassistant45);
#endif
                    foreach (DataRow dr in dt2.Rows)
                    {
                        string columnname = dr["MatchKey"].ToString();
                        string columnvalue = dr["MatchValue"].ToString();
                        drtemp[columnname] = columnvalue;
                    }
                    drtemp["Rank"] = i.ToString();
                    drtemp["HV"] = "建議Vol";
                    drtemp["IV"] = "建議Vol";
                    dataTable2.Rows.Add(drtemp);
                }
                ultraGrid2.DataSource = dataTable2;
            }
            catch(Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
        private void LoadDataUlt3()
        {
            ultraGrid3.DataSource = dataTable3;
        }
        private void InitialGrid1()
        {
            dataTable1.Columns[0].Caption = "邏輯";
            dataTable1.Columns[1].Caption = "一個月brokerPL(%)";
            dataTable1.Columns[2].Caption = "市場剩餘額度(檔)";
            dataTable1.Columns[3].Caption = "有額度釋出";
            dataTable1.Columns[4].Caption = "解禁可發行";
            dataTable1.Columns[5].Caption = "股票3日漲幅";
            dataTable1.Columns[6].Caption = "股票3日跌幅";
            dataTable1.Columns[7].Caption = "市場Theta IV";
            dataTable1.Columns[8].Caption = "市場部位近五日變化(萬)";
            dataTable1.Columns[9].Caption = "市場Med vol/ HV_60D";
            dataTable1.Columns[10].Caption = "市場超額利潤(仟)";
            dataTable1.Columns[11].Caption = "平均每檔權證超額利潤(元)";
            dataTable1.Columns[12].Caption = "Theta天數";
            dataTable1.Columns[13].Caption = "評鑑權證比重排名";
            dataTable1.Columns[14].Caption = "融資使用率(%)";
            dataTable1.Columns[15].Caption = "C有效檔數市佔(%)";      
            dataTable1.Columns[16].Caption = "C發行密度";
            dataTable1.Columns[17].Caption = "凱基-元大C發行密度";
            

            this.ultraGrid1.DisplayLayout.Override.DefaultRowHeight = 40;
            this.ultraGrid1.DisplayLayout.AutoFitStyle = Infragistics.Win.UltraWinGrid.AutoFitStyle.ResizeAllColumns;
            this.ultraGrid1.DisplayLayout.Override.WrapHeaderText = Infragistics.Win.DefaultableBoolean.True;
            this.ultraGrid1.DisplayLayout.Override.CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center;
            this.ultraGrid1.DisplayLayout.Override.CellAppearance.TextVAlign = Infragistics.Win.VAlign.Middle;
            
            Infragistics.Win.UltraWinGrid.UltraGridBand band0 = ultraGrid1.DisplayLayout.Bands[0];
            SetButton1();
            
        }
        private void InitialGrid4()
        {
           
            dataTable4.Columns[0].Caption = "邏輯";
            dataTable4.Columns[1].Caption = "一個月brokerPL(%)";
            dataTable4.Columns[2].Caption = "市場剩餘額度(檔)";
            dataTable4.Columns[3].Caption = "有額度釋出";
            dataTable4.Columns[4].Caption = "解禁可發行";
            dataTable4.Columns[5].Caption = "股票3日漲幅";
            dataTable4.Columns[6].Caption = "股票3日跌幅";
            dataTable4.Columns[7].Caption = "市場Theta IV";
            dataTable4.Columns[8].Caption = "市場部位近五日變化(萬)";
            dataTable4.Columns[9].Caption = "市場Med vol/ HV_60D";
            dataTable4.Columns[10].Caption = "市場超額利潤(仟)";
            dataTable4.Columns[11].Caption = "平均每檔權證超額利潤(元)";
            dataTable4.Columns[12].Caption = "Theta天數";
            dataTable4.Columns[13].Caption = "評鑑權證比重排名";
            dataTable4.Columns[14].Caption = "融資使用率(%)";
            dataTable4.Columns[15].Caption = "P有效檔數市佔(%)";
            dataTable4.Columns[16].Caption = "P發行密度";
            dataTable4.Columns[17].Caption = "凱基-元大P發行密度";
            this.ultraGrid4.DisplayLayout.Override.DefaultRowHeight = 40;
            this.ultraGrid4.DisplayLayout.AutoFitStyle = Infragistics.Win.UltraWinGrid.AutoFitStyle.ResizeAllColumns;
            this.ultraGrid4.DisplayLayout.Override.WrapHeaderText = Infragistics.Win.DefaultableBoolean.True;
            this.ultraGrid4.DisplayLayout.Override.CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center;
            this.ultraGrid4.DisplayLayout.Override.CellAppearance.TextVAlign = Infragistics.Win.VAlign.Middle;

            Infragistics.Win.UltraWinGrid.UltraGridBand band0 = ultraGrid4.DisplayLayout.Bands[0];
            SetButton4();

        }
        private void InitialGrid2()
        {
            dataTable2.Columns[0].Caption = "順位";
            dataTable2.Columns[1].Caption = "履約價(%)";
            dataTable2.Columns[2].Caption = "期間(月)";
            dataTable2.Columns[3].Caption = "HV";
            dataTable2.Columns[4].Caption = "IV";
            dataTable2.Columns[5].Caption = "張數";
            dataTable2.Columns[6].Caption = "CP";
            dataTable2.Columns[7].Caption = "類型";
            dataTable2.Columns[8].Caption = "發行價";
            dataTable2.Columns[9].Caption = "重設比";
            dataTable2.Columns[10].Caption = "行使比例";
            Infragistics.Win.UltraWinGrid.UltraGridBand band0 = ultraGrid2.DisplayLayout.Bands[0];

            this.ultraGrid2.DisplayLayout.AutoFitStyle = Infragistics.Win.UltraWinGrid.AutoFitStyle.ResizeAllColumns;
            this.ultraGrid2.DisplayLayout.Override.WrapHeaderText = Infragistics.Win.DefaultableBoolean.True;
            this.ultraGrid2.DisplayLayout.Override.CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center;
            SetButton2();
        }
        private void InitialGrid3()
        {
            Infragistics.Win.UltraWinGrid.UltraGridBand band0 = ultraGrid3.DisplayLayout.Bands[0];
            band0.Columns["備註"].Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button;
            
            this.ultraGrid3.DisplayLayout.AutoFitStyle = Infragistics.Win.UltraWinGrid.AutoFitStyle.ResizeAllColumns;
            this.ultraGrid3.DisplayLayout.Override.WrapHeaderText = Infragistics.Win.DefaultableBoolean.True;
            band0.Columns["編號"].CellActivation = Activation.NoEdit;
            band0.Columns["確認"].CellActivation = Activation.AllowEdit;
            band0.Columns["標的代號"].CellActivation = Activation.NoEdit;
            band0.Columns["標的名稱"].CellActivation = Activation.NoEdit;
            band0.Columns["昨日收盤價"].CellActivation = Activation.NoEdit;
            band0.Columns["CP"].CellActivation = Activation.NoEdit;
            band0.Columns["履約價"].CellActivation = Activation.NoEdit;
            band0.Columns["期間(月)"].CellActivation = Activation.NoEdit;
            band0.Columns["Vol(建議Vol)"].CellActivation = Activation.NoEdit;
            band0.Columns["發行張數"].CellActivation = Activation.NoEdit;
            band0.Columns["類型"].CellActivation = Activation.NoEdit;
            band0.Columns["重設比"].CellActivation = Activation.NoEdit;
            band0.Columns["發行價"].CellActivation = Activation.NoEdit;
            band0.Columns["行使比例"].CellActivation = Activation.NoEdit;
            band0.Columns["額度"].CellActivation = Activation.NoEdit;
            band0.Columns["Theta/VSP"].CellActivation = Activation.NoEdit;
            band0.Columns["備註"].CellActivation = Activation.NoEdit;

        }
        private void ultraGrid3_ClickCellButton(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
        {
            
            DataTable dt = new DataTable();
            dt.Columns.Add("符合", typeof(string));
            dt.Columns.Add("欄位",typeof(string));
            dt.Columns.Add("原始參數", typeof(string));
            dt.Columns.Add("比較", typeof(string));
            dt.Columns.Add("設定參數", typeof(string));
            // The UltraGrid fires CellClickButton when the user clicks the button in a cell
            // you a chance to handle it.
            //string str = "Y," + EngToChinese[column] + "," + data[column].ToString() + "," + "<" + "," + setting[column].ToString();
            int number = Convert.ToInt32(ultraGrid3.ActiveRow.Cells["編號"].Value);
            string uid = ultraGrid3.ActiveRow.Cells["標的代號"].Value.ToString();
            string cp = ultraGrid3.ActiveRow.Cells["CP"].Value.ToString();
            string UID_CP = uid + cp;
            //List<string> result = ShowLogicResult[number - 1];
            List<string> result = ShowLogicResult[UID_CP];
            foreach (string q in result)
            {
                string[] linesplit = q.Split(',');
                DataRow dr = dt.NewRow();
                dr["符合"] = linesplit[0];
                dr["欄位"] = linesplit[1];
                dr["原始參數"] = linesplit[2];
                dr["比較"] = linesplit[3];
                dr["設定參數"] = linesplit[4];
                dt.Rows.Add(dr);

            }
            ultraGrid5.DataSource = dt;
            //this.ultraGrid5.DisplayLayout.AutoFitStyle = Infragistics.Win.UltraWinGrid.AutoFitStyle.ResizeAllColumns;
            UltraGridBand band0 = ultraGrid5.DisplayLayout.Bands[0];
            band0.Columns["符合"].Width = 20;
            band0.Columns["欄位"].Width = 180;
            band0.Columns["原始參數"].Width = 65;
            band0.Columns["比較"].Width = 40;
            band0.Columns["設定參數"].Width = 65;
        }
        private void UltraGrid5_InitializeRow(object sender, InitializeRowEventArgs e)
        {
            string match = e.Row.Cells["符合"].Value.ToString();
            if (match == "Y")
                e.Row.Appearance.BackColor = Color.Salmon;
        }
        

        private void ultraGrid1_ClickCell(object sender, Infragistics.Win.UltraWinGrid.DoubleClickCellEventArgs e)
        {
            if (isEdit1)
            {
                string l = e.Cell.Row.Cells["Logic"].Value.ToString();
                if (l == "AND(V)")
                {
                    if (e.Cell.Value.ToString() == "")
                        e.Cell.Value = "V";
                    else
                        e.Cell.Value = "";
                }
                else if (l == "OR(O)")
                {
                    if (e.Cell.Value.ToString() == "")
                        e.Cell.Value = "O";
                    else
                        e.Cell.Value = "";
                }
            }
        }
        private void ultraGrid4_ClickCell(object sender, Infragistics.Win.UltraWinGrid.DoubleClickCellEventArgs e)
        {

            if (isEdit4)
            {
                string l = e.Cell.Row.Cells["Logic"].Value.ToString();
                if (l == "AND(V)")
                {
                    if (e.Cell.Value.ToString() == "")
                        e.Cell.Value = "V";
                    else
                        e.Cell.Value = "";
                }
                else if (l == "OR(O)")
                {
                    if (e.Cell.Value.ToString() == "")
                        e.Cell.Value = "O";
                    else
                        e.Cell.Value = "";
                }
            }
            /*
            ultraLabel1.Text = "";
            // The UltraGrid fires CellClickButton when the user clicks the button in a cell
            // you a chance to handle it.
            int number = Convert.ToInt32(ultraGrid3.ActiveRow.Cells["編號"].Value);
            List<string> result = ShowLogicResult[number - 1];

            foreach (var q in result)
            {
                ultraLabel1.Text += q + "\n";
            }
            */
        }
        private void SetButton1()
        {
            Infragistics.Win.UltraWinGrid.UltraGridBand band0 = ultraGrid1.DisplayLayout.Bands[0];
            if (isEdit1)
            {
                foreach (var c in band0.Columns)
                    c.CellActivation = Infragistics.Win.UltraWinGrid.Activation.AllowEdit;
                    
                band0.Columns["Logic"].CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
                ultraGrid1.DisplayLayout.Override.CellAppearance.BackColor = Color.White;
                button1.Visible = false;
                button2.Visible = true;
                button3.Visible = true;
                radioButton1.Enabled = true;
                radioButton2.Enabled = true;
            }
            else
            {
                foreach (var c in band0.Columns)
                    c.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
                ultraGrid1.DisplayLayout.Override.CellAppearance.BackColor = Color.Moccasin;
                button1.Visible = true;
                button2.Visible = false;
                button3.Visible = false;
                radioButton1.Enabled = false;
                radioButton2.Enabled = false;
            }
        }
        private void SetButton4()
        {
            Infragistics.Win.UltraWinGrid.UltraGridBand band0 = ultraGrid4.DisplayLayout.Bands[0];
            if (isEdit4)
            {
                foreach (var c in band0.Columns)
                    c.CellActivation = Infragistics.Win.UltraWinGrid.Activation.AllowEdit;

                band0.Columns["Logic"].CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
                ultraGrid4.DisplayLayout.Override.CellAppearance.BackColor = Color.White;
                button11.Visible = false;
                button12.Visible = true;
                button13.Visible = true;
                radioButton3.Enabled = true;
                radioButton4.Enabled = true;
            }
            else
            {
                foreach (var c in band0.Columns)
                    c.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
                ultraGrid4.DisplayLayout.Override.CellAppearance.BackColor = Color.Moccasin;
                button11.Visible = true;
                button12.Visible = false;
                button13.Visible = false;
                radioButton3.Enabled = false;
                radioButton4.Enabled = false;
            }
        }
        private void SetButton2()
        {
            Infragistics.Win.UltraWinGrid.UltraGridBand band0 = ultraGrid2.DisplayLayout.Bands[0];
            if (isEdit2)
            {
                foreach (var c in band0.Columns)
                    c.CellActivation = Infragistics.Win.UltraWinGrid.Activation.AllowEdit;
                ultraGrid2.DisplayLayout.Override.CellAppearance.BackColor = Color.White;
                button4.Visible = false;
                button5.Visible = true;
                button6.Visible = true;
            }
            else
            {
                foreach (var c in band0.Columns)
                    c.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
                ultraGrid2.DisplayLayout.Override.CellAppearance.BackColor = Color.Moccasin;
                button4.Visible = true;
                button5.Visible = false;
                button6.Visible = false;
            }
        }
        private void SetButton3()
        {
            
        }
        private void AutoSelect()
        {
            //LoadAll();
            ShowLogicResult.Clear();
            string traderID = comboBox1.Text;
            DateTime today = DateTime.Today;
            DateTime lastday = EDLib.TradeDate.LastNTradeDate(1);
            int CAndOr = 0;
            int PAndOr = 0;
            if (radioButton1.Checked == true)
                CAndOr = 1;
            else if(radioButton2.Checked == true)
                CAndOr = 0;

            if (radioButton3.Checked == true)
                PAndOr = 1;
            else if (radioButton4.Checked == true)
                PAndOr = 0;

            
            string todaystr = today.ToString("yyyyMMdd");
            string lastdaystr = lastday.ToString("yyyyMMdd");
#if !To39
            string sqldata = $@"SELECT [UID], [UName], [Trader], [BrokerPL_Month]
                            , [OptionAvailable], [LastOptionAvailable], [DiffOptionAvailable], [RiseUp_3Days], [DropDown_3Days]
                            , [Med_HV60D_VolRatio], [Theta_Days], [AppraisalRank], [FinancingRatio]
                            , [CallMarketShare], [PutMarketShare], [CallDensity], [PutDensity], [Theta_EndDate]
                            , [KgiYuanCallDensity], [KgiYuanPutDensity], [AllowIssue]
                            , [MarketAmtChg5Days], [AlphaTheta], [AvgAlphaThetaCost]
                            FROM [newEDIS].[dbo].[OptionAutoSelectData] WHERE [Trader] ='{traderID.TrimStart('0')}'";

            string sqldata_P = $@"SELECT [UID], [UName], [Trader], [BrokerPL_Month]
                            , [OptionAvailable], [LastOptionAvailable], [DiffOptionAvailable], [RiseUp_3Days], [DropDown_3Days]
                            , [Med_HV60D_VolRatio_P] AS Med_HV60D_VolRatio, [Theta_Days_P] AS Theta_Days, [AppraisalRank_P] AS AppraisalRank, [FinancingRatio]
                            , [CallMarketShare], [PutMarketShare], [CallDensity], [PutDensity], [Theta_EndDate_P] AS Theta_EndDate
                            , [KgiYuanCallDensity], [KgiYuanPutDensity], [AllowIssue]
                            , [MarketAmtChg5Days_P] AS MarketAmtChg5Days, [AlphaTheta_P] AS AlphaTheta, [AvgAlphaThetaCost_P] AS AvgAlphaThetaCost
                            FROM [newEDIS].[dbo].[OptionAutoSelectData] WHERE [Trader] ='{traderID.TrimStart('0')}'";

            string sqlURank = $@"SELECT  [UID],[WClass],[class],ISNULL([ProfitRank],'X') AS [ProfitRank]
                                  FROM [TwData].[dbo].[URank]
                                  where [TDate] = '{lastdaystr}'";
            DataTable Data = MSSQL.ExecSqlQry(sqldata, GlobalVar.loginSet.newEDIS);
            DataTable Data_P = MSSQL.ExecSqlQry(sqldata_P, GlobalVar.loginSet.newEDIS);
            DataTable URank = MSSQL.ExecSqlQry(sqlURank, GlobalVar.loginSet.twData);
#else
            string sqldata = $@"SELECT [UID], [UName], [Trader], [BrokerPL_Month]
                            , [OptionAvailable], [LastOptionAvailable], [DiffOptionAvailable], [RiseUp_3Days], [DropDown_3Days]
                            , [Med_HV60D_VolRatio], [Theta_Days], [AppraisalRank], [FinancingRatio]
                            , [CallMarketShare], [PutMarketShare], [CallDensity], [PutDensity], [Theta_EndDate]
                            , [KgiYuanCallDensity], [KgiYuanPutDensity], [AllowIssue]
                            , [MarketAmtChg5Days], [AlphaTheta], [AvgAlphaThetaCost]
                            FROM [WarrantAssistant].[dbo].[OptionAutoSelectData] WHERE [Trader] ='{traderID.TrimStart('0')}'";

            string sqldata_P = $@"SELECT [UID], [UName], [Trader], [BrokerPL_Month]
                            , [OptionAvailable], [LastOptionAvailable], [DiffOptionAvailable], [RiseUp_3Days], [DropDown_3Days]
                            , [Med_HV60D_VolRatio_P] AS Med_HV60D_VolRatio, [Theta_Days_P] AS Theta_Days, [AppraisalRank_P] AS AppraisalRank, [FinancingRatio]
                            , [CallMarketShare], [PutMarketShare], [CallDensity], [PutDensity], [Theta_EndDate_P] AS Theta_EndDate
                            , [KgiYuanCallDensity], [KgiYuanPutDensity], [AllowIssue]
                            , [MarketAmtChg5Days_P] AS MarketAmtChg5Days, [AlphaTheta_P] AS AlphaTheta, [AvgAlphaThetaCost_P] AS AvgAlphaThetaCost
                            FROM [WarrantAssistant].[dbo].[OptionAutoSelectData] WHERE [Trader] ='{traderID.TrimStart('0')}'";

            string sqlURank = $@"SELECT  [UID],[WClass],[class],ISNULL([ProfitRank],'X') AS [ProfitRank]
                                  FROM [TwData].[dbo].[URank]
                                  where [TDate] = '{lastdaystr}'";
            DataTable Data = MSSQL.ExecSqlQry(sqldata, GlobalVar.loginSet.warrantassistant45);
            DataTable Data_P = MSSQL.ExecSqlQry(sqldata_P, GlobalVar.loginSet.warrantassistant45);
            DataTable URank = MSSQL.ExecSqlQry(sqlURank, GlobalVar.loginSet.twData);
#endif
            //try Call
#if !To39
            string sqlsetting = $@"SELECT  [UID], [UName], [BrokerPL_Month]
                          ,[OptionAvailable], [OptionRelease], [RiseUp_3Days], [DropDown_3Days]
                          , [Med_HV60D_VolRatio], [Theta_Days], [AppraisalRank]
                          , [FinancingRatio], [CallMarketShare], [PutMarketShare], [CallDensity]
                          , [PutDensity], [K_OverLap], [T_Overlap], [IssuePut], [Theta_EndDate]
                          , [KgiYuanCallDensity], [KgiYuanPutDensity]
                          , [MarketAmtChg5Days], [AlphaTheta], [AvgAlphaThetaCost]
                          FROM [newEDIS].[dbo].[OptionAutoSelect]
                          WHERE [Trader] ='{traderID.TrimStart('0')}' and [checked] =1
                          ORDER BY [UID]";
            
            
             DataTable Setting = MSSQL.ExecSqlQry(sqlsetting, GlobalVar.loginSet.newEDIS);
#else
             string sqlsetting = $@"SELECT  [UID], [UName], [BrokerPL_Month]
                          ,[OptionAvailable], [OptionRelease], [RiseUp_3Days], [DropDown_3Days]
                          , [Med_HV60D_VolRatio], [Theta_Days], [AppraisalRank]
                          , [FinancingRatio], [CallMarketShare], [PutMarketShare], [CallDensity]
                          , [PutDensity], [K_OverLap], [T_Overlap], [IssuePut], [Theta_EndDate]
                          , [KgiYuanCallDensity], [KgiYuanPutDensity]
                          , [MarketAmtChg5Days], [AlphaTheta], [AvgAlphaThetaCost]
                          FROM [WarrantAssistant].[dbo].[OptionAutoSelect]
                          WHERE [Trader] ='{traderID.TrimStart('0')}' and [checked] =1
                          ORDER BY [UID]";


            DataTable Setting = MSSQL.ExecSqlQry(sqlsetting, GlobalVar.loginSet.warrantassistant45);
#endif


            DataTable SuggestVol = EDLib.SQL.MSSQL.ExecSqlQry($@"SELECT [UID],[IV_Rec]  
                              FROM [WarrantAssistant].[dbo].[RecommendVol]
                              WHERE CP='C' AND [DateTime] >= '{todaystr}'", GlobalVar.loginSet.warrantassistant45);

            DataTable SuggestVol_P = EDLib.SQL.MSSQL.ExecSqlQry($@"SELECT [UID],[IV_Rec]  
                              FROM [WarrantAssistant].[dbo].[RecommendVol]
                              WHERE CP='P' AND [DateTime] >= '{todaystr}'", GlobalVar.loginSet.warrantassistant45);

            bool[] Available_CParameter = { true, true, true, true };
            bool[] Available_PParameter = { true, true, true, true };
            double[] C_K = { 0, 0, 0, 0 };
            int[] C_T = { 0, 0, 0, 0 };
            int[] C_IssueLots = { 0, 0, 0, 0 };
            string[] C_Type = {"","","", ""};
            double[] C_Price = { 0, 0, 0, 0 };
            double[] C_ResetR = { 0, 0, 0, 0 };

            double[] P_K = { 0, 0, 0, 0 };
            int[] P_T = { 0, 0, 0, 0 };
            int[] P_IssueLots = { 0, 0, 0, 0 };
            string[] P_Type = { "", "", "", "" };
            double[] P_Price = { 0, 0, 0, 0 };
            double[] P_ResetR = { 0, 0, 0, 0 };
            for (int i = 1; i < 5; i++)
            {
                string strmatch = "C_MatchIssue" + i.ToString();
#if !To39
                DataTable dtCIssue = MSSQL.ExecSqlQry($@"SELECT [MatchKey],[MatchValue]
                  FROM [newEDIS].[dbo].[OptionAutoSelectSettings]
                  WHERE [TraderID] ='{traderID}' and [TableType] ='{strmatch}'", GlobalVar.loginSet.newEDIS);
#else
                DataTable dtCIssue = MSSQL.ExecSqlQry($@"SELECT [MatchKey],[MatchValue]
                  FROM [WarrantAssistant].[dbo].[OptionAutoSelectSettings]
                  WHERE [TraderID] ='{traderID}' and [TableType] ='{strmatch}'", GlobalVar.loginSet.warrantassistant45);
#endif
                foreach (DataRow dr in dtCIssue.Rows)
                {
                    if (dr["MatchKey"].ToString() == "K")
                    {
                        C_K[i - 1] = Convert.ToDouble(dr["MatchValue"].ToString());
                        if (C_K[i - 1] <= 0)
                            Available_CParameter[i - 1] = false;
                    }
                    if (dr["MatchKey"].ToString() == "T")
                    {
                        C_T[i - 1] = Convert.ToInt32(dr["MatchValue"].ToString());
                        if(C_T[i - 1] <= 0)
                            Available_CParameter[i - 1] = false;
                    }
                    if (dr["MatchKey"].ToString() == "IssueLots")
                    {
                        C_IssueLots[i - 1] = Convert.ToInt32(dr["MatchValue"].ToString());
                        if(C_IssueLots[i - 1] <= 0)
                            Available_CParameter[i - 1] = false;
                    }
                    if (dr["MatchKey"].ToString() == "Type")
                    {
                        C_Type[i - 1] = dr["MatchValue"].ToString();
                        if(C_Type[i - 1] != "一般型" && C_Type[i - 1] != "重設型")
                            Available_CParameter[i - 1] = false;
                    }
                    if (dr["MatchKey"].ToString() == "Price")
                    {
                        C_Price[i - 1] = Convert.ToDouble(dr["MatchValue"].ToString());
                        if(C_Price[i - 1] <= 0)
                            Available_CParameter[i - 1] = false;
                    }
                    if (dr["MatchKey"].ToString() == "ResetR")
                    {
                        C_ResetR[i - 1] = Convert.ToDouble(dr["MatchValue"].ToString());
                        if(C_ResetR[i - 1] < 0)
                            Available_CParameter[i - 1] = false;
                    }
                }
            }



            if (!Available_CParameter[0] || !Available_PParameter[0])
            {
                MessageBox.Show("參數設定有誤!");
                return;
            }
            string getprice = $@"SELECT DISTINCT CASE WHEN ([CommodityId]='1000') THEN 'IX0001' ELSE [CommodityId] END AS CommodityID
                                             ,ISNULL([LastPrice],0) AS LastPrice
                                             ,[tradedate] AS TradeDate
                                             ,isnull([BuyPriceBest1],0) AS Bid1
                                             ,isnull([SellPriceBest1],0) AS Ask1
                               FROM [TsQuote].[dbo].[vwprice2]
							   WHERE [kind] in ('ETF','Index','Stock');";
            DataTable vprice = EDLib.SQL.MSSQL.ExecSqlQry(getprice, GlobalVar.loginSet.tsquoteSqlConnString);
            
            for (int i = 1; i < 5; i++)
            {
                string strmatch = "P_MatchIssue" + i.ToString();
#if !To39
                DataTable dtPIssue = MSSQL.ExecSqlQry($@"SELECT [MatchKey],[MatchValue]
                  FROM [newEDIS].[dbo].[OptionAutoSelectSettings]
                  WHERE [TraderID] ='{traderID}' and [TableType] ='{strmatch}'", GlobalVar.loginSet.newEDIS);
#else
                DataTable dtPIssue = MSSQL.ExecSqlQry($@"SELECT [MatchKey],[MatchValue]
                  FROM [WarrantAssistant].[dbo].[OptionAutoSelectSettings]
                  WHERE [TraderID] ='{traderID}' and [TableType] ='{strmatch}'", GlobalVar.loginSet.warrantassistant45);
#endif
                foreach (DataRow dr in dtPIssue.Rows)
                {
                    if (dr["MatchKey"].ToString() == "K")
                    {
                        P_K[i - 1] = Convert.ToDouble(dr["MatchValue"].ToString());
                        if (P_K[i - 1] <= 0)
                            Available_PParameter[i - 1] = false;
                    }
                    if (dr["MatchKey"].ToString() == "T")
                    {
                        P_T[i - 1] = Convert.ToInt32(dr["MatchValue"].ToString());
                        if (P_T[i - 1] <= 0)
                            Available_PParameter[i - 1] = false;
                    }
                    if (dr["MatchKey"].ToString() == "IssueLots")
                    {
                        P_IssueLots[i - 1] = Convert.ToInt32(dr["MatchValue"].ToString());
                        if (P_IssueLots[i - 1] <= 0)
                            Available_PParameter[i - 1] = false;
                    }
                    if (dr["MatchKey"].ToString() == "Type")
                    {
                        P_Type[i - 1] = dr["MatchValue"].ToString();
                        if (P_Type[i - 1] != "一般型" && P_Type[i - 1] != "重設型")
                            Available_PParameter[i - 1] = false;
                    }
                    if (dr["MatchKey"].ToString() == "Price")
                    {
                        P_Price[i - 1] = Convert.ToDouble(dr["MatchValue"].ToString());
                        if (P_Price[i - 1] <= 0)
                            Available_PParameter[i - 1] = false;
                    }
                    if (dr["MatchKey"].ToString() == "ResetR")
                    {
                        P_ResetR[i - 1] = Convert.ToDouble(dr["MatchValue"].ToString());
                        if (P_ResetR[i - 1] < 0)
                            Available_PParameter[i - 1] = false;
                    }
                }
            }
            
            dataTable3.Clear();
            int number = 1;
            double timeToExpiry = 0.5;
            double costOfCarry = 0.01;
            double interestRate_Index = 0.0;
            //20230721利率從0.01 改成0.015
            double interestRate = 0.015;
            double dayPerYear = 365;
            
            foreach (DataRow dr in Setting.Rows)
            {
                DataRow setting = dr;
                string uid = dr["UID"].ToString();
                string uname = dr["UName"].ToString();
                int issueput = Convert.ToInt32(dr["IssuePut"].ToString());
                DataRow[] select = Data.Select($@"UID='{uid}'");
                DataRow[] select_P = Data_P.Select($@"UID='{uid}'");
                DataRow data;
                DataRow data_P;
                double closeP = 0;
                DataRow[] select2 = vprice.Select($@"CommodityID='{uid}'");
                if (select2.Length == 0)
                    continue;
                //取得履約價中點
                if (Convert.ToDouble(select2[0][1].ToString()) == 0)
                {
                    if (Convert.ToDouble(select2[0][3].ToString()) == 0)
                    {
                        if (Convert.ToDouble(select2[0][4].ToString()) == 0)
                            closeP = 0;
                        else
                            closeP = Convert.ToDouble(select2[0][4].ToString());
                    }
                    else
                    {
                        closeP = Convert.ToDouble(select2[0][3].ToString());
                    }
                }
                else
                {
                    closeP = Convert.ToDouble(select2[0][1].ToString());
                }

             
                double K_overlap = 0;
                double T_overlap = 0;
                K_overlap = Convert.ToDouble(setting["K_OverLap"].ToString());
                T_overlap = Convert.ToDouble(setting["T_OverLap"].ToString());
               
                if (select.Length > 0)//如果select有值，那select_P也會有值
                {
                    data = select[0];
                    data_P = select_P[0];
                    if (Result(data, setting, VarCAnd, VarCOr, CAndOr) == true)
                    {
                        int CRank = -1;
                            
                        for(int i = 0; i < 4; i++)
                        {
                            if (!Available_CParameter[i])
                                break;
                            int updays = (int)Math.Round((C_T[i] + T_overlap) * 30);
                            int downdays = (C_T[i] - T_overlap) > 0 ? (int)Math.Round((C_T[i] - T_overlap) * 30) : 0;
                            DateTime update = DateTime.Today.AddDays(updays);
                            DateTime downdate = DateTime.Today.AddDays(downdays);
                            double strikePrice = 0;
                            if (C_Type[i] == "一般型")
                                strikePrice = roundpriceK(closeP * C_K[i] / 100);
                            else//重設型
                                strikePrice = roundpriceK(closeP * C_ResetR[i] / 100);
                            
                            double upk = Math.Round(strikePrice * (100 + K_overlap) / 100, 2);
                            double downk = Math.Round(strikePrice * (100 - K_overlap) / 100, 2);
                            DataTable warrantbasic = EDLib.SQL.MSSQL.ExecSqlQry($@"SELECT * FROM (SELECT  [stkid] AS [UID]
                                              , CASE WHEN (SUBSTRING([type],4,4) ='認購權證') THEN 'c' ELSE 'p' END AS WClass
                                              , [strike_now] AS StrikePrice
                                              , [maturitydate] AS MaturityDate
                                          FROM [HEDGE].[dbo].[WARRANTS]
                                          WHERE [kgiwrt] = '自家' AND [maturitydate] >'{DateTime.Today.ToString("yyyyMMdd")}' AND ([type] LIKE '%認購權證%'  OR [type] LIKE '%認售權證%')) AS A
                                          WHERE A.[WClass] = 'c' AND A.[UID] = '{uid}' AND [MaturityDate] >= '{downdate.ToString("yyyyMMdd")}' AND [MaturityDate] <= '{update.ToString("yyyyMMdd")}'
                                          AND [StrikePrice] >= {downk} AND [StrikePrice] <= {upk}", "Data Source=10.101.10.5;Initial Catalog=HEDGE;User ID=hedgeuser;Password=hedgeuser");
                            

                            if(warrantbasic.Rows.Count > 0)
                            {
                                /*
                                if(i == 0)
                                {
                                    if (Convert.ToDouble(data["CallDensity"].ToString()) < Convert.ToDouble(setting["CallDensity"].ToString()))
                                        continue;
                                    else
                                        break;
                                }
                                */
                                continue;
                            }
                            else
                            {
                                CRank = i;
                                break;
                            }
                        }
                        if (CRank > -1)
                        {
                            List<string> resultstr = LogicORANDResult(data, setting, VarCOr, VarCAnd);
                            string UID_CP = uid + "C";
                            ShowLogicResult.Add(UID_CP, resultstr);
                            double hv = 0;
                            DataRow[] vol = SuggestVol.Select($@"UID='{uid}'");
                            if (vol.Length > 0)
                            {
                                hv = Convert.ToDouble(vol[0][1].ToString());
                            }
                            else
                            {

                                string sql_VolManager = $@"SELECT [HV_60D]
                                    FROM [10.101.10.5].[WMM3].[dbo].[VolManagerDetail]
			                        where [UPDATETIME] >='{lastday.ToString("yyyyMMdd")} 14:00' AND [UPDATETIME] <'{today.ToString("yyyyMMdd")}' AND [STKID] ='{uid}'";
                                DataTable dt2 = MSSQL.ExecSqlQry(sql_VolManager, GlobalVar.loginSet.warrantassistant45);

                                if (dt2.Rows.Count > 0)
                                {
                                    hv = Convert.ToDouble(dt2.Rows[0]["HV_60D"].ToString()) * 1.3;
                                }
                                else
                                {
                                    continue;
                                }
                            }
                            double Mprice = 0;
                            hv = hv / 100;
                            if (C_Type[CRank] == "一般型")
                            {
                                double strikePrice = closeP * C_K[CRank] / 100;
                                timeToExpiry = ((double)C_T[CRank] * 30 / dayPerYear);
                                if (uid.Length > 4 && uid.Substring(0, 2) != "00")
                                {
                                    Mprice = WarrantPrice("C", closeP, strikePrice, interestRate_Index, hv, timeToExpiry, 1);
                                }
                                else
                                {
                                    Mprice = WarrantPrice("C", closeP, strikePrice, interestRate, hv, timeToExpiry, 1);
                                }
                            }
                            else//重設型
                            {
                                double strikePrice = closeP * C_ResetR[CRank] / 100;
                                timeToExpiry = (((double)C_T[CRank] * 30 + 3.0) / dayPerYear);
                                if (uid.Length > 4 && uid.Substring(0, 2) != "00")
                                {
                                    Mprice = WarrantPrice("C", closeP, strikePrice, interestRate_Index, hv, timeToExpiry, 1);
                                }
                                else
                                {
                                    Mprice = WarrantPrice("C", closeP, strikePrice, interestRate, hv, timeToExpiry, 1);
                                }
                            }
                            double cr = C_Price[CRank] / Mprice;
                            //若反推行使比例大於1，代表權證價格小於1.5，給原先的權證價格
                            if (cr >= 1)
                            {
                                cr = 1;
                            }
                            if (closeP > 200)
                            {
                                cr = Math.Round(cr, 3);
                            }
                            else
                                cr = Math.Round(cr, 2);

                            try
                            {
                                DataRow drnew = dataTable3.NewRow();
                                drnew["編號"] = number;
                                drnew["確認"] = 0;
                                drnew["標的代號"] = uid;
                                drnew["標的名稱"] = uname;
                                drnew["昨日收盤價"] = closeP;

                                drnew["CP"] = "C";
                                if (C_Type[CRank] == "一般型")
                                    drnew["履約價"] = roundpriceK(closeP * C_K[CRank] / 100);
                                else if (C_Type[0] == "重設型")
                                    drnew["履約價"] = roundpriceK(closeP * C_ResetR[CRank] / 100);

                                drnew["期間(月)"] = C_T[CRank];
                                drnew["Vol(建議Vol)"] = Math.Round(hv * 100);
                                drnew["發行張數"] = C_IssueLots[CRank];
                                drnew["類型"] = C_Type[CRank];
                                if (C_Type[CRank] == "一般型")
                                    drnew["重設比"] = 0;
                                else if (C_Type[CRank] == "重設型")
                                    drnew["重設比"] = C_ResetR[CRank];
                                if (cr == 1)
                                    drnew["發行價"] = Math.Round(Mprice, 3);
                                else
                                    drnew["發行價"] = C_Price[CRank];
                                drnew["行使比例"] = cr;
                                if(IssueCredit.ContainsKey(uid))
                                    drnew["額度"] = IssueCredit[uid];
                                drnew["備註"] = $@"順位{CRank+1}";

                                string urank = "";
                                DataRow[] URankSelect = URank.Select($@"UID = '{uid}' AND WClass = 'c'");
                                if(URankSelect.Length > 0)
                                {
                                    string theta = URankSelect[0][2].ToString();
                                    string vsp = URankSelect[0][3].ToString();
                                    urank = theta + " / " + vsp;
                                }
                                drnew["Theta/VSP"] = urank;
                                dataTable3.Rows.Add(drnew);
                                number++;
                            }
                            catch
                            {
                                MessageBox.Show("Insert Table3 Fail");
                            }
                        }
                            
                    }
                    if (issueput == 1)
                    {

                        if (Result(data_P, setting, VarPAnd, VarPOr, PAndOr) == true)
                        {
                            int PRank = -1;

                            for (int i = 0; i < 4; i++)
                            {
                                if (!Available_PParameter[i])
                                    break;
                                int updays = (int)Math.Round((P_T[i] + T_overlap) * 30);
                                int downdays = (P_T[i] - T_overlap) > 0 ? (int)Math.Round((P_T[i] - T_overlap) * 30) : 0;
                                DateTime update = DateTime.Today.AddDays(updays);
                                DateTime downdate = DateTime.Today.AddDays(downdays);
                                double strikePrice = 0;
                                if (P_Type[i] == "一般型")
                                    strikePrice = roundpriceK(closeP * P_K[i] / 100);
                                else//重設型
                                    strikePrice = roundpriceK(closeP * P_ResetR[i] / 100);
                                double upk = Math.Round(strikePrice * (100 + K_overlap) / 100, 2);
                                double downk = Math.Round(strikePrice * (100 - K_overlap) / 100, 2);
                                DataTable warrantbasic = EDLib.SQL.MSSQL.ExecSqlQry($@"SELECT * FROM (SELECT  [stkid] AS [UID]
                                              , CASE WHEN (SUBSTRING([type],4,4) ='認購權證') THEN 'c' ELSE 'p' END AS WClass
                                              , [strike] AS StrikePrice
                                              , [maturitydate] AS MaturityDate
                                          FROM [HEDGE].[dbo].[WARRANTS]
                                          WHERE [kgiwrt] = '自家' AND [maturitydate] >'{DateTime.Today.ToString("yyyyMMdd")}' AND ([type] LIKE '%認購權證%'  OR [type] LIKE '%認售權證%')) AS A
                                          WHERE A.[WClass] = 'p' AND A.[UID] = '{uid}' AND [MaturityDate] >= '{downdate.ToString("yyyyMMdd")}' AND [MaturityDate] <= '{update.ToString("yyyyMMdd")}'
                                          AND [StrikePrice] >= {downk} AND [StrikePrice] <= {upk}", "Data Source=10.101.10.5;Initial Catalog=HEDGE;User ID=hedgeuser;Password=hedgeuser");
                                if (warrantbasic.Rows.Count > 0)
                                {
                                    /*
                                    if (i == 0)
                                    {
                                        if (Convert.ToDouble(data_P["PutDensity"].ToString()) < Convert.ToDouble(setting["PutDensity"].ToString()))
                                            continue;
                                        else
                                            break;
                                    }
                                    */
                                    continue;
                                }
                                else
                                {
                                    PRank = i;
                                    break;
                                }
                            }
                            if (PRank > -1)
                            {
                                List<string> resultstr = LogicORANDResult(data_P, setting, VarPOr, VarPAnd);
                                string UID_CP = uid + "P";
                                ShowLogicResult.Add(UID_CP, resultstr);
                                double hv = 0;
                                DataRow[] vol = SuggestVol_P.Select($@"UID='{uid}'");

                                if (vol.Length > 0)
                                {
                                    hv = Convert.ToDouble(vol[0][1].ToString());

                                }
                                else
                                {

                                    string sql_VolManager = $@"SELECT [HV_60D]
                                    FROM [10.101.10.5].[WMM3].[dbo].[VolManagerDetail]
			                        where [UPDATETIME] >='{lastday.ToString("yyyyMMdd")} 14:00' AND [UPDATETIME] <'{today.ToString("yyyyMMdd")}' AND [STKID] ='{uid}'";
                                    DataTable dt2 = MSSQL.ExecSqlQry(sql_VolManager, GlobalVar.loginSet.warrantassistant45);
                                    if (dt2.Rows.Count > 0)
                                    {
                                        hv = Convert.ToDouble(dt2.Rows[0]["HV_60D"].ToString()) * 1.3;
                                        //sw.Write($@"VolManager: {Math.Round(hv, 2)}  ");
                                    }
                                    else
                                    {
                                        continue;
                                        //sw.Write($@"沒有建議Vol: {Math.Round(hv, 2)}  ");
                                    }

                                }
                                double Mprice = 0;
                                hv = hv / 100;
                                if (P_Type[PRank] == "一般型")
                                {
                                    double strikePrice = closeP * P_K[PRank] / 100;
                                    timeToExpiry = ((double)P_T[PRank] * 30 / dayPerYear);
                                    if (uid.Length > 4 && uid.Substring(0, 2) != "00")
                                    {
                                        Mprice = WarrantPrice("P", closeP, strikePrice, interestRate_Index, hv, timeToExpiry, 1);
                                    }
                                    else
                                    {
                                        Mprice = WarrantPrice("P", closeP, strikePrice, interestRate, hv, timeToExpiry, 1);
                                    }
                                }
                                else//重設型
                                {
                                    double strikePrice = closeP * P_ResetR[PRank] / 100;
                                    timeToExpiry = (((double)P_T[PRank] * 30 + 3.0) / dayPerYear);
                                    if (uid.Length > 4 && uid.Substring(0, 2) != "00")
                                    {
                                        Mprice = WarrantPrice("P", closeP, strikePrice, interestRate_Index, hv, timeToExpiry, 1);
                                    }
                                    else
                                    {
                                        Mprice = WarrantPrice("P", closeP, strikePrice, interestRate, hv, timeToExpiry, 1);
                                    }
                                }
                                
                                double cr = P_Price[PRank] / Mprice;
                                
                                //若反推行使比例大於1，代表權證價格小於1.5，給原先的權證價格
                                if (cr >= 1)
                                {
                                    cr = 1;
                                }
                                if (closeP > 200)
                                {
                                    cr = Math.Round(cr, 3);
                                }
                                else
                                    cr = Math.Round(cr, 2);
                                try
                                {
                                    DataRow drnew = dataTable3.NewRow();
                                    drnew["編號"] = number;
                                    drnew["確認"] = 0;
                                    drnew["標的代號"] = uid;
                                    drnew["標的名稱"] = uname;
                                    drnew["昨日收盤價"] = closeP;

                                    drnew["CP"] = "P";
                                    if (P_Type[PRank] == "一般型")
                                        drnew["履約價"] = roundpriceK(closeP * P_K[PRank] / 100);
                                    else if (P_Type[PRank] == "重設型")
                                        drnew["履約價"] = (closeP * P_ResetR[PRank] / 100);
                                    drnew["期間(月)"] = P_T[PRank];
                                    drnew["Vol(建議Vol)"] = Math.Round(hv * 100);
                                    
                                    drnew["發行張數"] = P_IssueLots[PRank];
                                    drnew["類型"] = P_Type[PRank];
                                   
                                    if (P_Type[PRank] == "一般型")
                                        drnew["重設比"] = 0;
                                    else if (P_Type[0] == "重設型")
                                        drnew["重設比"] = P_ResetR[PRank];

                                   
                                    if (cr == 1)
                                        drnew["發行價"] = Math.Round(Mprice, 3);
                                    else
                                        drnew["發行價"] = P_Price[PRank];
                                    drnew["行使比例"] = cr;
                                    drnew["額度"] = IssueCredit[uid];
                                    drnew["備註"] = $@"順位{PRank+1}";
                                    string urank = "";
                                    DataRow[] URankSelect = URank.Select($@"UID = '{uid}' AND WClass = 'p'");
                                    if (URankSelect.Length > 0)
                                    {
                                        string theta = URankSelect[0][2].ToString();
                                        string vsp = URankSelect[0][3].ToString();
                                        urank = theta + " / " + vsp;
                                    }
                                    drnew["Theta/VSP"] = urank;
                                    if (IssueCredit.ContainsKey(uid))
                                        dataTable3.Rows.Add(drnew);
                                    number++;
                                }
                                catch(Exception ex)
                                {
                                    MessageBox.Show($@"insert fail {ex.Message}");
                                }
                            }
                        }
                    }
                }
                else
                    continue;

            }
            MessageBox.Show($@"篩選完成");
        }

        private void button1_EditClick(object sender, EventArgs e)
        {
            isEdit1 = true;
            SetButton1();
        }
        private void button3_CancelClick(object sender, EventArgs e)
        {
            isEdit1 = false;
            LoadDataUlt1();
            SetButton1();
        }

        private void button4_EditClick(object sender, EventArgs e)
        {
            isEdit2 = true;
            SetButton2();
        }
        private void button6_CancelClick(object sender, EventArgs e)
        {
            isEdit2 = false;
            LoadDataUlt2();
            SetButton2();
        }
        private void button9_ClearClick(object sender, EventArgs e)
        {
            /*
            foreach (Infragistics.Win.UltraWinGrid.UltraGridRow r in ultraGrid3.Rows)
            {
                r.Cells["確認"].Value = 0;
            }
            */
            dataTable3.Clear();
        }

        private void button10_SelectClick(object sender, EventArgs e)
        {
            AutoSelect();
        }

        private void button11_EditClick(object sender, EventArgs e)
        {
            isEdit4 = true;
            SetButton4();
        }

        private void button13_CancelClick(object sender, EventArgs e)
        {
            isEdit4 = false;
            LoadDataUlt4();
            SetButton4();
        }

        private void button2_ConfirmClick(object sender, EventArgs e)
        {
            string traderID = comboBox1.Text;
            DataRow V = dataTable1.Rows[0];
            DataRow O = dataTable1.Rows[1];
            int num = dataTable1.Columns.Count;
            bool checkAndOr = true;
            bool checkRadioButton = true;
            for(int i = 1; i < num; i++)
            {
                if ((V[i].ToString() != "V" && V[i].ToString() != "") || (O[i].ToString() != "O" && O[i].ToString() != ""))
                {
                    checkAndOr = false;
                }
                
            }
            if ((radioButton1.Checked == true && radioButton2.Checked == true) || (radioButton1.Checked == false && radioButton2.Checked == false))
                checkRadioButton = false;
            if (checkAndOr&checkRadioButton)
            {
                try
                {
#if !To39
                    string sql = $@"DELETE FROM [newEDIS].[dbo].[OptionAutoSelectSettings]
                                WHERE [TableType] ='C_MatchAndOr' AND [TraderID] ='{traderID}'";
                    MSSQL.ExecSqlCmd(sql, GlobalVar.loginSet.newEDIS);
#else
                    string sql = $@"DELETE FROM [WarrantAssistant].[dbo].[OptionAutoSelectSettings]
                                WHERE [TableType] ='C_MatchAndOr' AND [TraderID] ='{traderID}'";
                    MSSQL.ExecSqlCmd(sql, GlobalVar.loginSet.warrantassistant45);
#endif
                    for (int i = 1; i < num; i++)
                    {
                        if (V[i].ToString() == "V")
                        {
#if !To39
                            string sqlInsert = $@"INSERT INTO [newEDIS].[dbo].[OptionAutoSelectSettings] ([TraderID]
                                     , [TableType], [MatchKey], [MatchValue])
	                                VALUES ('{traderID}','C_MatchAndOr','AND','{dataTable1.Columns[i].ToString()}')";
                            MSSQL.ExecSqlCmd(sqlInsert, GlobalVar.loginSet.newEDIS);
#else
                            string sqlInsert = $@"INSERT INTO [WarrantAssistant].[dbo].[OptionAutoSelectSettings] ([TraderID]
                                     , [TableType], [MatchKey], [MatchValue])
	                                VALUES ('{traderID}','C_MatchAndOr','AND','{dataTable1.Columns[i].ToString()}')";
                            MSSQL.ExecSqlCmd(sqlInsert, GlobalVar.loginSet.warrantassistant45);
#endif
                        }
                        if (O[i].ToString() == "O")
                        {
#if !To39
                            string sqlInsert = $@"INSERT INTO [newEDIS].[dbo].[OptionAutoSelectSettings] ([TraderID]
                                     , [TableType], [MatchKey], [MatchValue])
	                                VALUES ('{traderID}','C_MatchAndOr','OR','{dataTable1.Columns[i].ToString()}')";
                            MSSQL.ExecSqlCmd(sqlInsert, GlobalVar.loginSet.newEDIS);
#else
                            string sqlInsert = $@"INSERT INTO [WarrantAssistant].[dbo].[OptionAutoSelectSettings] ([TraderID]
                                     , [TableType], [MatchKey], [MatchValue])
	                                VALUES ('{traderID}','C_MatchAndOr','OR','{dataTable1.Columns[i].ToString()}')";
                            MSSQL.ExecSqlCmd(sqlInsert, GlobalVar.loginSet.warrantassistant45);
#endif
                        }
                    }
                    if (radioButton1.Checked == true)
                    {
#if !To39
                        string sqlInsert = $@"INSERT INTO [newEDIS].[dbo].[OptionAutoSelectSettings] ([TraderID]
                                     , [TableType], [MatchKey], [MatchValue])
	                                VALUES ('{traderID}','C_MatchAndOr','IsClicked','AND')";
                        MSSQL.ExecSqlCmd(sqlInsert, GlobalVar.loginSet.newEDIS);
#else
                        string sqlInsert = $@"INSERT INTO [WarrantAssistant].[dbo].[OptionAutoSelectSettings] ([TraderID]
                                     , [TableType], [MatchKey], [MatchValue])
	                                VALUES ('{traderID}','C_MatchAndOr','IsClicked','AND')";
                        MSSQL.ExecSqlCmd(sqlInsert, GlobalVar.loginSet.warrantassistant45);
#endif
                    }
                    else
                    {
#if !To39
                        string sqlInsert = $@"INSERT INTO [newEDIS].[dbo].[OptionAutoSelectSettings] ([TraderID]
                                     , [TableType], [MatchKey], [MatchValue])
	                                VALUES ('{traderID}','C_MatchAndOr','IsClicked','OR')";
                        MSSQL.ExecSqlCmd(sqlInsert, GlobalVar.loginSet.newEDIS);
#else
                        string sqlInsert = $@"INSERT INTO [WarrantAssistant].[dbo].[OptionAutoSelectSettings] ([TraderID]
                                     , [TableType], [MatchKey], [MatchValue])
	                                VALUES ('{traderID}','C_MatchAndOr','IsClicked','OR')";
                        MSSQL.ExecSqlCmd(sqlInsert, GlobalVar.loginSet.warrantassistant45);
#endif
                    }
                    isEdit1 = false;
                    LoadDataUlt1();
                    SetButton1();
                }
                catch(Exception ex)
                {
                    MessageBox.Show($@"{ex.Message}");
                }
            }
            else if (!checkAndOr && checkRadioButton)
            {
                MessageBox.Show("邏輯參數設定錯誤!");
            }
            else if (checkAndOr && !checkRadioButton)
            {
                MessageBox.Show("請點擊AND或OR");
            }
            else
            {
                MessageBox.Show("邏輯參數設定錯誤!\n請點擊AND或OR");
            }
        }

        private void button12_ConfirmClick(object sender, EventArgs e)
        {
            string traderID = comboBox1.Text;
            DataRow V = dataTable4.Rows[0];
            DataRow O = dataTable4.Rows[1];
            int num = dataTable4.Columns.Count;
            bool checkAndOr = true;
            bool checkRadioButton = true;
            for (int i = 1; i < num; i++)
            {
                if ((V[i].ToString() != "V" && V[i].ToString() != "") || (O[i].ToString() != "O" && O[i].ToString() != ""))
                {
                    checkAndOr = false;
                }

            }
            if ((radioButton3.Checked == true && radioButton4.Checked == true) || (radioButton3.Checked == false && radioButton4.Checked == false))
                checkRadioButton = false;
            if (checkAndOr & checkRadioButton)
            {
                try
                {
#if !To39
                    string sql = $@"DELETE FROM [newEDIS].[dbo].[OptionAutoSelectSettings]
                                WHERE [TableType] ='P_MatchAndOr' AND [TraderID] ='{traderID}'";
                    MSSQL.ExecSqlCmd(sql, GlobalVar.loginSet.newEDIS);
#else
                    string sql = $@"DELETE FROM [WarrantAssistant].[dbo].[OptionAutoSelectSettings]
                                WHERE [TableType] ='P_MatchAndOr' AND [TraderID] ='{traderID}'";
                    MSSQL.ExecSqlCmd(sql, GlobalVar.loginSet.warrantassistant45);
#endif
                    for (int i = 1; i < num; i++)
                    {
                        if (V[i].ToString() == "V")
                        {
#if !To39
                            string sqlInsert = $@"INSERT INTO [newEDIS].[dbo].[OptionAutoSelectSettings] ([TraderID]
                                     , [TableType], [MatchKey], [MatchValue])
	                                VALUES ('{traderID}','P_MatchAndOr','AND','{dataTable4.Columns[i].ToString()}')";
                            MSSQL.ExecSqlCmd(sqlInsert, GlobalVar.loginSet.newEDIS);
#else
                            string sqlInsert = $@"INSERT INTO [WarrantAssistant].[dbo].[OptionAutoSelectSettings] ([TraderID]
                                     , [TableType], [MatchKey], [MatchValue])
	                                VALUES ('{traderID}','P_MatchAndOr','AND','{dataTable4.Columns[i].ToString()}')";
                            MSSQL.ExecSqlCmd(sqlInsert, GlobalVar.loginSet.warrantassistant45);
#endif
                        }
                        if (O[i].ToString() == "O")
                        {
#if !To39
                            string sqlInsert = $@"INSERT INTO [newEDIS].[dbo].[OptionAutoSelectSettings] ([TraderID]
                                     , [TableType], [MatchKey], [MatchValue])
	                                VALUES ('{traderID}','P_MatchAndOr','OR','{dataTable4.Columns[i].ToString()}')";
                            MSSQL.ExecSqlCmd(sqlInsert, GlobalVar.loginSet.newEDIS);
#else
                            string sqlInsert = $@"INSERT INTO [WarrantAssistant].[dbo].[OptionAutoSelectSettings] ([TraderID]
                                     , [TableType], [MatchKey], [MatchValue])
	                                VALUES ('{traderID}','P_MatchAndOr','OR','{dataTable4.Columns[i].ToString()}')";
                            MSSQL.ExecSqlCmd(sqlInsert, GlobalVar.loginSet.warrantassistant45);
#endif
                        }
                    }
                    if (radioButton3.Checked == true)
                    {
#if !To39
                        string sqlInsert = $@"INSERT INTO [newEDIS].[dbo].[OptionAutoSelectSettings] ([TraderID]
                                     , [TableType], [MatchKey], [MatchValue])
	                                VALUES ('{traderID}','P_MatchAndOr','IsClicked','AND')";
                        MSSQL.ExecSqlCmd(sqlInsert, GlobalVar.loginSet.newEDIS);
#else
                        string sqlInsert = $@"INSERT INTO [WarrantAssistant].[dbo].[OptionAutoSelectSettings] ([TraderID]
                                     , [TableType], [MatchKey], [MatchValue])
	                                VALUES ('{traderID}','P_MatchAndOr','IsClicked','AND')";
                        MSSQL.ExecSqlCmd(sqlInsert, GlobalVar.loginSet.warrantassistant45);
#endif
                    }
                    else
                    {
#if !To39
                        string sqlInsert = $@"INSERT INTO [newEDIS].[dbo].[OptionAutoSelectSettings] ([TraderID]
                                     , [TableType], [MatchKey], [MatchValue])
	                                VALUES ('{traderID}','P_MatchAndOr','IsClicked','OR')";
                        MSSQL.ExecSqlCmd(sqlInsert, GlobalVar.loginSet.newEDIS);
#else
                        string sqlInsert = $@"INSERT INTO [WarrantAssistant].[dbo].[OptionAutoSelectSettings] ([TraderID]
                                     , [TableType], [MatchKey], [MatchValue])
	                                VALUES ('{traderID}','P_MatchAndOr','IsClicked','OR')";
                        MSSQL.ExecSqlCmd(sqlInsert, GlobalVar.loginSet.warrantassistant45);
#endif
                    }
                    isEdit4 = false;
                    LoadDataUlt4();
                    SetButton4();
                }
                catch (Exception ex)
                {
                    MessageBox.Show($@"{ex.Message}");
                }
            }
            else if (!checkAndOr && checkRadioButton)
            {
                MessageBox.Show("邏輯參數設定錯誤!");
            }
            else if (checkAndOr && !checkRadioButton)
            {
                MessageBox.Show("請點擊AND或OR");
            }
            else
            {
                MessageBox.Show("邏輯參數設定錯誤!\n請點擊AND或OR");
            }
        }

        private void button8_AutoIssueClick(object sender, EventArgs e)
        {
            DateTime today = DateTime.Today;
            DateTime lastday = EDLib.TradeDate.LastNTradeDate(1);
            string todaystr = today.ToString("yyyyMMdd");
            string lastdaystr = lastday.ToString("yyyyMMdd");
            //Set Insert SQL Parameters 
            string sqlTemp3 = @"INSERT INTO [ApplyTempList] (SerialNum, UnderlyingID, K, T, R, HV, IV, IssueNum, ResetR, BarrierR, FinancialR, Type, CP, UseReward, ConfirmChecked, Apply1500W, UserID, MDate, TempName, TempType, TraderID, IVNew) ";
            sqlTemp3 += "VALUES(@SerialNum, @UnderlyingID, @K, @T, @R, @HV, @IV, @IssueNum, @ResetR, @BarrierR, @FinancialR, @Type, @CP, @UseReward, @ConfirmChecked, @Apply1500W, @UserID, @MDate, @TempName ,@TempType, @TraderID, @IVNew)";
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
            ps.Add(new SqlParameter("@IVNew", SqlDbType.Float));

            SQLCommandHelper h = new SQLCommandHelper(GlobalVar.loginSet.warrantassistant45, sqlTemp3, ps);


          

            string sqlTemp = $@"SELECT  [SerialNum]
                              FROM [WarrantAssistant].[dbo].[ApplyTempList]
                              WHERE [MDate] >='{today.ToString("yyyyMMdd")}' AND [TraderID] = '{userID}'";
            DataTable dtTemp = MSSQL.ExecSqlQry(sqlTemp, GlobalVar.loginSet.warrantassistant45);

            int ApplyTempListMax = 0;
            int TempListDeleteLogMax = 0;
            foreach(DataRow dr in dtTemp.Rows)
            {
                string serial = dr["SerialNum"].ToString();
                int seriallength = serial.Length;
                int temp = Convert.ToInt32(serial.Substring(17, seriallength - 17));
                ApplyTempListMax = temp > ApplyTempListMax ? temp : ApplyTempListMax;
            }

            string sqlTemp2 = $@"SELECT [SerialNum]
                          FROM [WarrantAssistant].[dbo].[TempListDeleteLog]
                          WHERE [DateTime] >='{today.ToString("yyyyMMdd")}' and [Trader] ='{userID}'";
            DataTable dtTemp2 = MSSQL.ExecSqlQry(sqlTemp2, GlobalVar.loginSet.warrantassistant45);

            foreach (DataRow dr in dtTemp2.Rows)
            {
                string serial = dr["SerialNum"].ToString();
                int seriallength = serial.Length;
                int temp = Convert.ToInt32(serial.Substring(17, seriallength-17));
                //MessageBox.Show(temp.ToString());
                TempListDeleteLogMax = temp > TempListDeleteLogMax ? temp : TempListDeleteLogMax;
            }
            //MessageBox.Show($@"{ApplyTempListMax} {TempListDeleteLogMax}");
            int count = ApplyTempListMax >= TempListDeleteLogMax ? ApplyTempListMax : TempListDeleteLogMax;
            int issuecount = 0;
            foreach (Infragistics.Win.UltraWinGrid.UltraGridRow r in ultraGrid3.Rows)
            {
                string serialNum = DateTime.Today.ToString("yyyyMMdd") + userID + "01" + (count + 1).ToString("0#");
                string underlyingID = r.Cells["標的代號"].Value.ToString();
                int check = Convert.ToInt32(r.Cells["確認"].Value);
                if (check == 1)
                {
                    count++;
                    issuecount++;
                }
                else
                    continue;
                    
                
                string underlyingName = r.Cells["標的名稱"].Value.ToString();
                double k = Convert.ToDouble(r.Cells["履約價"].Value);
                int t = Convert.ToInt32(r.Cells["期間(月)"].Value);
                double cr = Convert.ToDouble(r.Cells["行使比例"].Value);
                double hv = Convert.ToDouble(r.Cells["Vol(建議Vol)"].Value);
                double iv = Convert.ToDouble(r.Cells["Vol(建議Vol)"].Value);
                double issueNum = Convert.ToDouble(r.Cells["發行張數"].Value);
                double resetR = Convert.ToDouble(r.Cells["重設比"].Value);
                double barrierR = 0;
                double financialR = 0;
                string type = r.Cells["類型"].Value.ToString();
                string cp = r.Cells["CP"].Value.ToString();
                string useReward = "N";
                string confirm = "N";
                string is1500W = "N";
                string tempName = "";
                string traderID = userID;

                DateTime expiryDate;
                expiryDate = GlobalVar.globalParameter.nextTradeDate3.AddMonths(t);
                expiryDate = expiryDate.AddDays(-1);
                string sqlTemp4 = "SELECT TOP 1 TradeDate from TradeDate WHERE IsTrade='Y' AND TradeDate >= '" + expiryDate.ToString("yyyy-MM-dd") + "'";
                DataView dvTemp4 = DeriLib.Util.ExecSqlQry(sqlTemp4, GlobalVar.loginSet.tsquoteSqlConnString);
                foreach (DataRowView drTemp in dvTemp4)
                {
                    expiryDate = Convert.ToDateTime(drTemp["TradeDate"]);
                }
                string expiryMonth = "";
                int month = expiryDate.Month;
                if (month >= 10)
                {
                    if (month == 10)
                        expiryMonth = "A";
                    if (month == 11)
                        expiryMonth = "B";
                    if (month == 12)
                        expiryMonth = "C";
                }
                else
                    expiryMonth = month.ToString();

                string expiryYear = "";
                expiryYear = expiryDate.AddYears(-1).ToString("yyyy");
                expiryYear = expiryYear.Substring(expiryYear.Length - 1, 1);

                string warrantType = "";
                string tempType = "";

                if (cp == "P")
                {
                    warrantType = "售";
                    tempType = "2";
                }
                else
                {
                    warrantType = "購";
                    tempType = "1";
                }
                

                tempName = underlyingName + "凱基" + expiryYear + expiryMonth + warrantType;



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
                h.SetParameterValue("@IVNew", iv);

                h.ExecuteCommand();
                
            }
            h.Dispose();

            GlobalUtility.LogInfo("Log", GlobalVar.globalParameter.userID + $@"發行自動篩選 {issuecount} 檔");

            MessageBox.Show("Done!");
        }

        private void button5_ConfirmClick(object sender, EventArgs e)
        {
            string traderID = comboBox1.Text;
            int C_checkenum1 = Convert.ToInt32(dataTable2.Rows[0][0].ToString());
            int C_checkenum2 = Convert.ToInt32(dataTable2.Rows[1][0].ToString());
            int C_checkenum3 = Convert.ToInt32(dataTable2.Rows[2][0].ToString());
            int C_checkenum4 = Convert.ToInt32(dataTable2.Rows[3][0].ToString());
            int P_checkenum1 = Convert.ToInt32(dataTable2.Rows[4][0].ToString());
            int P_checkenum2 = Convert.ToInt32(dataTable2.Rows[5][0].ToString());
            int P_checkenum3 = Convert.ToInt32(dataTable2.Rows[6][0].ToString());
            int P_checkenum4 = Convert.ToInt32(dataTable2.Rows[7][0].ToString());
            if ((C_checkenum1 != 1) && (C_checkenum2 != 1) && (C_checkenum3 != 1) && (C_checkenum4 != 1) && (C_checkenum1 + C_checkenum2 + C_checkenum3 + C_checkenum4) != 10)
            {
                MessageBox.Show("C順位設定有誤");
                return;
            }
            if ((P_checkenum1 != 1) && (P_checkenum2 != 1) && (P_checkenum3 != 1) && (P_checkenum4 != 1) && (P_checkenum1 + P_checkenum2 + P_checkenum3 + P_checkenum4) != 10)
            {
                MessageBox.Show("P順位設定有誤");
                return;
            }
#if !To39
            string sql1 = $@"DELETE [newEDIS].[dbo].[OptionAutoSelectSettings]
                        WHERE [TraderID] ='{traderID}'
                        AND [TableType] LIKE '%C_MatchIssue%'";
            string sql2 = $@"DELETE [newEDIS].[dbo].[OptionAutoSelectSettings]
                        WHERE [TraderID] ='{traderID}'
                        AND [TableType]  LIKE '%P_MatchIssue%'";
            MSSQL.ExecSqlCmd(sql1, GlobalVar.loginSet.newEDIS);
            MSSQL.ExecSqlCmd(sql2, GlobalVar.loginSet.newEDIS);
#else
            string sql1 = $@"DELETE [WarrantAssistant].[dbo].[OptionAutoSelectSettings]
                        WHERE [TraderID] ='{traderID}'
                        AND [TableType] LIKE '%C_MatchIssue%'";
            string sql2 = $@"DELETE [WarrantAssistant].[dbo].[OptionAutoSelectSettings]
                        WHERE [TraderID] ='{traderID}'
                        AND [TableType]  LIKE '%P_MatchIssue%'";
            MSSQL.ExecSqlCmd(sql1, GlobalVar.loginSet.warrantassistant45);
            MSSQL.ExecSqlCmd(sql2, GlobalVar.loginSet.warrantassistant45);
#endif

            int num = dataTable2.Columns.Count;
            try
            {
                for (int index = 0; index < 4; index++)
                {
                    DataRow Crow = dataTable2.Rows[index];
                    string strmatch = "C_MatchIssue" + Crow[0].ToString();
                    for (int i = 0; i < num; i++)
                    {
                        if (i == 0 || i == 3 || i == 4 || i == 10)
                            continue;
                        else if (i == 7)
                        {
                            string Ctype = "";
                            if (Crow[i].ToString() == "1" || Crow[i].ToString() == "一般型")
                                Ctype = "一般型";
                            else
                                Ctype = "重設型";
#if !To39
                            string sqlInsert = $@"INSERT INTO [newEDIS].[dbo].[OptionAutoSelectSettings] ([TraderID]
                                    , [TableType], [MatchKey], [MatchValue])
	                            VALUES ('{traderID}','{strmatch}','{dataTable2.Columns[i].ToString()}','{Ctype}')";
                            MSSQL.ExecSqlCmd(sqlInsert, GlobalVar.loginSet.newEDIS);
#else
                            string sqlInsert = $@"INSERT INTO [WarrantAssistant].[dbo].[OptionAutoSelectSettings] ([TraderID]
                                    , [TableType], [MatchKey], [MatchValue])
	                            VALUES ('{traderID}','{strmatch}','{dataTable2.Columns[i].ToString()}','{Ctype}')";
                            MSSQL.ExecSqlCmd(sqlInsert, GlobalVar.loginSet.warrantassistant45);
#endif
                        }
                        else
                        {
                            if (Crow[i].ToString() != "")
                            {
#if !To39
                                string sqlInsert = $@"INSERT INTO [newEDIS].[dbo].[OptionAutoSelectSettings] ([TraderID]
                                    , [TableType], [MatchKey], [MatchValue])
	                            VALUES ('{traderID}','{strmatch}','{dataTable2.Columns[i].ToString()}','{Crow[i].ToString()}')";
                                MSSQL.ExecSqlCmd(sqlInsert, GlobalVar.loginSet.newEDIS);
#else
                                string sqlInsert = $@"INSERT INTO [WarrantAssistant].[dbo].[OptionAutoSelectSettings] ([TraderID]
                                    , [TableType], [MatchKey], [MatchValue])
	                            VALUES ('{traderID}','{strmatch}','{dataTable2.Columns[i].ToString()}','{Crow[i].ToString()}')";
                                MSSQL.ExecSqlCmd(sqlInsert, GlobalVar.loginSet.warrantassistant45);
#endif
                            }
                            else
                            {
#if !To39
                                string sqlInsert = $@"INSERT INTO [newEDIS].[dbo].[OptionAutoSelectSettings] ([TraderID]
                                    , [TableType], [MatchKey], [MatchValue])
	                            VALUES ('{traderID}','{strmatch}','{dataTable2.Columns[i].ToString()}','0')";
                                MSSQL.ExecSqlCmd(sqlInsert, GlobalVar.loginSet.newEDIS);
#else
                                string sqlInsert = $@"INSERT INTO [WarrantAssistant].[dbo].[OptionAutoSelectSettings] ([TraderID]
                                    , [TableType], [MatchKey], [MatchValue])
	                            VALUES ('{traderID}','{strmatch}','{dataTable2.Columns[i].ToString()}','0')";
                                MSSQL.ExecSqlCmd(sqlInsert, GlobalVar.loginSet.warrantassistant45);
#endif
                            }
                        }

                    }
                }
                for (int index = 4; index < 8; index++)
                {
                    DataRow Prow = dataTable2.Rows[index];
                    string strmatch = "P_MatchIssue" + Prow[0].ToString();
                    for (int i = 0; i < num; i++)
                    {
                        if (i == 0 || i == 3 || i == 4 || i == 10)
                            continue;
                        else if (i == 7)
                        {
                            
                            string Ptype = "";
                            
                            
                            if (Prow[i].ToString() == "1" || Prow[i].ToString() == "一般型")
                                Ptype = "一般型";
                            else
                                Ptype = "重設型";



#if !To39
                            string sqlInsert2 = $@"INSERT INTO [newEDIS].[dbo].[OptionAutoSelectSettings] ([TraderID]
                                    , [TableType], [MatchKey], [MatchValue])
	                            VALUES ('{traderID}','{strmatch}','{dataTable2.Columns[i].ToString()}','{Ptype}')";
                            MSSQL.ExecSqlCmd(sqlInsert2, GlobalVar.loginSet.newEDIS);
#else
                            string sqlInsert2 = $@"INSERT INTO [WarrantAssistant].[dbo].[OptionAutoSelectSettings] ([TraderID]
                                    , [TableType], [MatchKey], [MatchValue])
	                            VALUES ('{traderID}','{strmatch}','{dataTable2.Columns[i].ToString()}','{Ptype}')";
                            MSSQL.ExecSqlCmd(sqlInsert2, GlobalVar.loginSet.warrantassistant45);
#endif
                        }
                        else
                        {
                            
                            if (Prow[i].ToString() != "")
                            {
#if !To39
                                string sqlInsert2 = $@"INSERT INTO [newEDIS].[dbo].[OptionAutoSelectSettings] ([TraderID]
                                    , [TableType], [MatchKey], [MatchValue])
	                            VALUES ('{traderID}','{strmatch}','{dataTable2.Columns[i].ToString()}','{Prow[i].ToString()}')";
                                MSSQL.ExecSqlCmd(sqlInsert2, GlobalVar.loginSet.newEDIS);
#else
                                string sqlInsert2 = $@"INSERT INTO [WarrantAssistant].[dbo].[OptionAutoSelectSettings] ([TraderID]
                                    , [TableType], [MatchKey], [MatchValue])
	                            VALUES ('{traderID}','{strmatch}','{dataTable2.Columns[i].ToString()}','{Prow[i].ToString()}')";
                                MSSQL.ExecSqlCmd(sqlInsert2, GlobalVar.loginSet.warrantassistant45);
#endif
                            }
                            else
                            {
                                string sqlInsert2 = $@"INSERT INTO [newEDIS].[dbo].[OptionAutoSelectSettings] ([TraderID]
                                    , [TableType], [MatchKey], [MatchValue])
	                            VALUES ('{traderID}','{strmatch}','{dataTable2.Columns[i].ToString()}','0')";
                                MSSQL.ExecSqlCmd(sqlInsert2, GlobalVar.loginSet.newEDIS);
                            }
                        }

                    }
                }

                isEdit2 = false;
                
                LoadDataUlt2();
                SetButton2();
            }
            catch (Exception ex)
            {
                MessageBox.Show($@"{ex.Message}");
            }
        }
        private void UltraGrid3_InitializeRow(object sender, InitializeRowEventArgs e)
        {
            double cr = Convert.ToDouble(e.Row.Cells["行使比例"].Value);
            if (cr >= 1)
            {
                e.Row.Cells["發行價"].Appearance.BackColor = Color.PaleVioletRed;
                e.Row.Cells["行使比例"].ToolTipText = "CR = 1，仍未達目標價格";
            }
            string uid = e.Row.Cells["標的代號"].Value.ToString();

            string CP = e.Row.Cells["CP"].Value.ToString();
            if (CP == "C")
                CP = "c";
            else
                CP = "p";

            string rank = e.Row.Cells["備註"].Value.ToString();
            if(rank == "順位4")
            {
                e.Row.Cells["備註"].Appearance.BackColor = Color.PaleVioletRed;
            }

            //MessageBox.Show(e.Row.Cells["履約價"].ToString());
            double k = Convert.ToDouble(e.Row.Cells["履約價"].Value.ToString());
            double t = Convert.ToDouble(e.Row.Cells["期間(月)"].Value.ToString());
                
            string sql = $@"SELECT [K_OverLap], [T_Overlap]
                        FROM [newEDIS].[dbo].[OptionAutoSelect]
                        WHERE [UID] ='{uid}'";
                
            double K_overlap = 0;
            double T_overlap = 0;
            DataTable dt = MSSQL.ExecSqlQry(sql, GlobalVar.loginSet.newEDIS);
            foreach (DataRow dr in dt.Rows)
            {
                K_overlap = Convert.ToDouble(dr["K_OverLap"].ToString());
                T_overlap = Convert.ToDouble(dr["T_OverLap"].ToString());
            }
            int updays = (int)Math.Round((t + T_overlap) * 30);
            int downdays = (t - T_overlap) > 0 ? (int)Math.Round((t - T_overlap) * 30) : 0;
            DateTime update = DateTime.Today.AddDays(updays);
            DateTime downdate = DateTime.Today.AddDays(downdays);
            double upk = Math.Round(k * (100 + K_overlap) / 100, 2);
            double downk = Math.Round(k * (100 - K_overlap) / 100, 2);
            try
            {

                string sql2 = $@"SELECT A.[WID]
                        FROM (SELECT [WID], CASE [UID] WHEN 'TWA00' THEN 'IX0001' ELSE [UID] END AS [UID], [WNAME], [WClass], [IssuedDate], [MaturityDate], [StrikePrice]
                                     FROM [newEDIS].[dbo].[WarrantBasics]
                                    WHERE ([MaturityDate] >'{DateTime.Today.ToString("yyyyMMdd")}' OR [MaturityDate] = '19110101') AND [WClass] = '{CP}' AND [IssuerName] ='9200' AND [UID] ='{uid}') AS A
                         LEFT JOIN (SELECT [wname], [maturitydate]
                      FROM [10.101.10.5].[HEDGE].[dbo].[WARRANTS]
                      WHERE [kgiwrt] = '自家' and [maturitydate] > '{DateTime.Today.ToString("yyyyMMdd")}') AS B on A.[WName] = B.[wname]
                        WHERE B.[maturitydate]<='{update.ToString("yyyyMMdd")}' AND B.[maturitydate]>='{downdate.ToString("yyyyMMdd")}' AND A.[StrikePrice]>={downk} AND A.[StrikePrice]<={upk}";
                DataTable dt2 = MSSQL.ExecSqlQry(sql2, GlobalVar.loginSet.newEDIS);
                
               
                

                if (dt2.Rows.Count > 0)
                {
                    e.Row.Cells["履約價"].Appearance.ForeColor = Color.Crimson;
                    e.Row.Cells["履約價"].ToolTipText = "K、T重複\n";
                    e.Row.Cells["履約價"].ToolTipText += $@"到期日:{downdate.ToString("yyyyMMdd")}-{update.ToString("yyyyMMdd")} 履約價:{downk}-{upk}";
                }
            }
            catch(Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void ultraGrid2_InitializeLayout(object sender, InitializeLayoutEventArgs e)
        {
            //ultraGrid2.DisplayLayout.Override.RowSelectorHeaderStyle = RowSelectorHeaderStyle.ColumnChooserButton;
            if (!e.Layout.ValueLists.Exists("MyValueList"))
            {
                Infragistics.Win.ValueList v;
                v = e.Layout.ValueLists.Add("MyValueList");
                v.ValueListItems.Add(1, "一般型");
                v.ValueListItems.Add(2, "重設型");
                /*
                Infragistics.Win.ValueList v2;
                v2 = e.Layout.ValueLists.Add("MyValueList2");
                v2.ValueListItems.Add(1, "C");
                v2.ValueListItems.Add(2, "P");
                */
            }
            e.Layout.Bands[0].Columns["Type"].ValueList = e.Layout.ValueLists["MyValueList"];
            //e.Layout.Bands[0].Columns["CP"].ValueList = e.Layout.ValueLists["MyValueList2"];
        }
        private void ultraGrid3_InitializeLayout(object sender, InitializeLayoutEventArgs e)
        {
            e.Layout.ScrollBounds = ScrollBounds.ScrollToFill;
        }
        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            string traderID = comboBox1.Text;
            if (traderID != userID)
            {
                LoadDataUlt1();
                LoadDataUlt4();
                LoadDataUlt2();
                LoadDataUlt3();
            }
        }

    }
}

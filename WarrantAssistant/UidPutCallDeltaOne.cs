using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WarrantAssistant
{
    class UidPutCallDeltaOne
    {
        public string uid;
        public double KgiCallDeltaOne;
        public double KgiPutDeltaOne;
        public double KgiCallPutRatio;
        public double AllCallDeltaOne;
        public double AllPutDeltaOne;
        public double KgiAllPutRatio;
        public double YuanPutDeltaOne;
        public int KgiPutNum;

        public UidPutCallDeltaOne(string uid, double KgiCallDeltaOne, double KgiPutDeltaOne, double KgiCallPutRatio, double AllCallDeltaOne, double AllPutDeltaOne, double KgiAllPutRatio, double YuanPutDeltaOne, int KgiPutNum)
        {
            this.uid = uid;
            this.KgiCallDeltaOne = KgiCallDeltaOne;
            this.KgiPutDeltaOne = KgiPutDeltaOne;
            this.KgiCallPutRatio = KgiCallPutRatio;
            this.AllCallDeltaOne = AllCallDeltaOne;
            this.AllPutDeltaOne = AllPutDeltaOne;
            this.YuanPutDeltaOne = YuanPutDeltaOne;
            this.KgiAllPutRatio = KgiAllPutRatio;
            this.KgiPutNum = KgiPutNum;
        }

    }

    public class UID
    {
        string uid;
        string uname;
        double brokerPL_month;
        double profit_month;
        int optionavailable;
        int optionrelease;
        double riseup_3days;
        double dropdown_3days;
        double thetaiv_weekdelta;
        double med_hv60d_volratio;
        double theta_days;
        double optweightinprice1to2dot5;
        double financingratio;
        double callmarketshare;
        double putmarketshare;
        double calldensity;     
        double putdensity;
        double k_overLap;
        double t_overLap;
        public UID(string uid,string uname,double brokerPL_month,double profit_month,int optionavailable,int optionrelease,
                    double riseup_3days,double dropdown_3days,double thetaiv_weekdelta,double med_hv60d_volratio,double theta_days,double optweightinprice1to2dot5,
                    double financingratio,double callmarketshare,double putmarketshare,double calldensity,double putdensity,double k_overLap,double t_overLap)
        {

        }

        public bool LogicAND()
        {
            return true;
        }
        public bool LogicOR()
        {
            return true;
        }
        public bool Result(bool and,bool or)
        {
            return true;
        }
        
    }
    //15個比較參數
    public enum AutoSelectSettings
    {
        BrokerPL_Month = 0,
        Profit_Month = 1,
        OptionAvailable = 2,
        OptionRelease = 3,
        RiseUp_3Days = 4,
        DropDown_3Days = 5,
        ThetaIV_WeekDelta = 6,
        Med_HV60D_VolRatio = 7,
        Theta_Days = 8,
        OptWeightInPrice1To2dot5 = 9,
        FinancingRatio = 10,
        CallMarketShare = 11,
        PutMarketShare = 12,
        CallDensity = 13,
        PutDensity = 14
    }
}

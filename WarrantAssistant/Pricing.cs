using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace WarrantAssistant
{
    public enum CallPutType { Call = 1, Put = 2 }
    public enum WarrantType { Normal = 1, BullBear = 2, Reset = 3 }

    class Pricing
    {
        public const double PI = 3.1415926535897932384626433;
        public static double StandardGaussianDensity(double x)
        {
            return Math.Exp(-1 * x * x * 0.5) / Math.Sqrt(2 * PI);
        }
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
        public static double RoundUp(double x, int d)
        {
            double x1 = Math.Round(x, d);
            if (x1 < x) { x1 = x1 + 1.0 / Math.Pow(10.0, d); }
            return x1;
        }
        public static double RoundDown(double x, int d)
        {
            double x1 = Math.Round(x, d);
            if (x1 > x) { x1 = x1 - 1.0 / Math.Pow(10.0, d); }
            return x1;
        }
        public static double BlackSholeFormula(CallPutType cp,
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
            if (cp == CallPutType.Call)
                bsprice = spotPrice * Math.Exp((costOfCarry - interestRate) * timeToExpiry) * StandardGaussianProbability(d1) - strikePrice * Math.Exp(-1.0 * interestRate * timeToExpiry) * StandardGaussianProbability(d2);
            else
                bsprice = strikePrice * Math.Exp(-1.0 * interestRate * timeToExpiry) * StandardGaussianProbability(-1.0 * d2) - spotPrice * Math.Exp((costOfCarry - interestRate) * timeToExpiry) * StandardGaussianProbability(-1.0 * d1);

            return bsprice;
        }

        public static double NormalWarrantPrice(CallPutType cp, double underlyingPrice, double k, double interestRate, double vol, int t, double cr)
        {
            double price = 0.0;
            double timeToExpiry = (t * 30.0) / GlobalVar.globalParameter.dayPerYear;
            price = BlackSholeFormula(cp, underlyingPrice, k, interestRate, vol, timeToExpiry, interestRate) * cr;
            return price;
        }

        public static double BullBearWarrantPrice(CallPutType cp,
                                              double underlyingPrice,
                                              double resetR,
                                              double interestR,
                                              double vol,
                                              int t,
                                              double financialR,
                                              double cr)
        {
            double price = 0.0;
            price = (cp == CallPutType.Call ? 1.0 : -1.0) * (underlyingPrice - Math.Round(underlyingPrice * resetR, 2)) * cr + Math.Round(underlyingPrice * resetR, 2) * t * (30.0 / GlobalVar.globalParameter.dayPerYear) * cr * financialR;
            price = RoundUp(price, 2);
            return price;
        }

        public static double ResetWarrantPrice(CallPutType cp,
                                              double underlyingPrice,
                                              double resetR,
                                              double interestRate,
                                              double vol,
                                              int t,
                                              double cr)
        {
            double price = 0.0;
            double timetoexpiry = (t * 30.0 + 3.0) / GlobalVar.globalParameter.dayPerYear;
            double k = 0.0;
            k = underlyingPrice * resetR;
            price = BlackSholeFormula(cp, underlyingPrice, k, interestRate, vol, timetoexpiry, interestRate) * cr;
            return price;
        }

        public static double Delta(CallPutType cp,
                                  double spotPrice,
                                  double strikePrice,
                                  double interestRate,
                                  double volatility,
                                  double timeToExpiry,
                                  double costOfCarry)
        {
            double delta=0.0;
            double d1 = (Math.Log(spotPrice / strikePrice) + (costOfCarry + volatility * volatility * 0.5) * timeToExpiry) / (volatility * Math.Sqrt(timeToExpiry));
            double d2 = d1 - volatility * Math.Sqrt(timeToExpiry);
            if (cp == CallPutType.Call)
                delta = Math.Exp((costOfCarry - interestRate) * timeToExpiry) * StandardGaussianProbability(d1);
            else
                delta = Math.Exp((costOfCarry - interestRate) * timeToExpiry) * (StandardGaussianProbability(d1)-1);

            return delta;
        }

        public static double Theta(CallPutType cp,
                                  double spotPrice,
                                  double strikePrice,
                                  double interestRate,
                                  double volatility,
                                  double timeToExpiry,
                                  double costOfCarry)
        {
            double theta = 0.0;
            double d1 = (Math.Log(spotPrice / strikePrice) + (costOfCarry + volatility * volatility * 0.5) * timeToExpiry) / (volatility * Math.Sqrt(timeToExpiry));
            double d2 = d1 - volatility * Math.Sqrt(timeToExpiry);
            if (cp == CallPutType.Call)
            {
                theta = spotPrice * volatility * Math.Exp((costOfCarry - interestRate) * timeToExpiry) * StandardGaussianDensity(d1) / (2 * Math.Sqrt(timeToExpiry))
                    - (interestRate - costOfCarry) * spotPrice * Math.Exp((costOfCarry - interestRate) * timeToExpiry) * StandardGaussianProbability(d1)
                    + interestRate * strikePrice * Math.Exp((-interestRate * timeToExpiry) * StandardGaussianProbability(d2));
            }
            else
            {
                theta = spotPrice * volatility * Math.Exp((costOfCarry - interestRate) * timeToExpiry) * StandardGaussianDensity(d1) / (2 * Math.Sqrt(timeToExpiry))
                    + (interestRate - costOfCarry) * spotPrice * Math.Exp((costOfCarry - interestRate) * timeToExpiry) * StandardGaussianProbability(-d1)
                    - interestRate * strikePrice * Math.Exp((-interestRate * timeToExpiry) * StandardGaussianProbability(-d2));
            }

            return theta/252;
        }
        /*
        public static double Vega(CallPutType cp,
                                  double spotPrice,
                                  double strikePrice,
                                  double interestRate,
                                  double volatility,
                                  double timeToExpiry,
                                  double costOfCarry)
        {
            double vega = 0.0;
            double d1 = (Math.Log(spotPrice / strikePrice) + (costOfCarry + volatility * volatility * 0.5) * timeToExpiry) / (volatility * Math.Sqrt(timeToExpiry));
            double d2 = d1 - volatility * Math.Sqrt(timeToExpiry);

            return vega;
        }
        */
    }
}

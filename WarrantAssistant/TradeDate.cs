using System;
using System.Data.SqlClient;

/// <summary>
/// TradeDate
/// </summary>
public class TradeDate
{	
    static public string LastNTradeDate(int N) {        
        return LastNTradeDateDT(N).ToString("yyyyMMdd");
    }

    static public DateTime LastNTradeDateDT(int N) {
        //Get Last Trading Date
        int nDays = -1;
        DateTime retDate = DateTime.Today.AddDays(nDays);
        string date = retDate.ToString("yyyyMMdd");

        SqlConnection conn2 = new SqlConnection("Data Source=10.101.10.5;Initial Catalog=HEDGE;User ID=hedgeuser;Password=hedgeuser");
        conn2.Open();
        SqlCommand cmd2 = new SqlCommand("Select * from HOLIDAY where CCY='TWD' and HOL_DATE='" + date + "'", conn2);
        SqlDataReader holiday = cmd2.ExecuteReader();
        if (!holiday.HasRows)
            N--;

        while (holiday.HasRows || N != 0) {
            nDays--;

            retDate = DateTime.Today.AddDays(nDays);
            date = retDate.ToString("yyyyMMdd");

            cmd2 = new SqlCommand("Select * from HOLIDAY where CCY='TWD' and HOL_DATE='" + date + "'", conn2);
            holiday.Close();
            holiday = cmd2.ExecuteReader();
            if (!holiday.HasRows)
                N--;
        }

        conn2.Close();
        cmd2.Dispose();
        holiday.Close();
        return retDate;

    }
}

#define To39
using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.Windows.Forms;
using Infragistics.Win.UltraWinGrid;
using System.Data.SqlClient;
using EDLib.SQL;
using System.Configuration;

namespace WarrantAssistant
{
    public partial class FrmApplyTotalList:Form
    {

        public SqlConnection conn = new SqlConnection(GlobalVar.loginSet.warrantassistant45);

        //在10:40左右額度結果出來後，用來記錄要搶額度的權證，哪些改為用10000張發行
        public List<string> changeTo_1w = new List<string>();
        //用來記需要搶額度的權證
        public List<string> Iscompete = new List<string>();
        private DataTable dt;// = new DataTable();
        private string userID = GlobalVar.globalParameter.userID;
        //private string userID = GlobalVar.globalParameter.userID;
        private bool isEdit = false;
        int applyhour_compete = 0;
        int applymin_compete = 0;
        Dictionary<string, UidPutCallDeltaOne> UidDeltaOne = new Dictionary<string, UidPutCallDeltaOne>();
        List<string> Market30 = new List<string>();//市值前20大
        List<string> IsSpecial = new List<string>();//特殊標的
        List<string> IsIndex = new List<string>();//臺灣50指數,臺灣中型100指數,櫃買富櫃50指
        public double NonSpecialCallPutRatio = Convert.ToDouble(ConfigurationManager.AppSettings["NonSpecialCallPutRatio"].ToString());
        public double SpecialCallPutRatio = Convert.ToDouble(ConfigurationManager.AppSettings["SpecialCallPutRatio"].ToString());
        public double SpecialKGIALLPutRatio = Convert.ToDouble(ConfigurationManager.AppSettings["SpecialKGIALLPutRatio"].ToString());
        public double ISTOP30MaxIssue = Convert.ToDouble(ConfigurationManager.AppSettings["ISTOP30MaxIssue"].ToString());
        public double NonTOP30MaxIssue = Convert.ToDouble(ConfigurationManager.AppSettings["NonTOP30MaxIssue"].ToString());
        public FrmApplyTotalList() {
            InitializeComponent();
        }

        private void FrmApplyTotalList_Load(object sender, EventArgs e) {
            //LoadApplyTime();
            LoadData();
            InitialGrid();
            LoadIsIndex();
            LoadIsSpecial();
        }

        //記上傳申報窗口的時間
        private void InitialGrid() {
            
            UltraGridBand band0 = ultraGrid1.DisplayLayout.Bands[0];
            band0.Columns["IssueNum"].Format = "N0";
            //band0.Columns["EquivalentNum"].Format = "N0";
            //band0.Columns["Result"].Format = "N0";
            band0.Columns["Credit"].Format = "N0";
            band0.Columns["RewardCredit"].Format = "N0";
            band0.Columns["ApplyTime"].Width = 50;
            band0.Columns["IssueNum"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right;
            band0.Columns["EquivalentNum"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right;
            band0.Columns["Result"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right;
            band0.Columns["Credit"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right;
            band0.Columns["RewardCredit"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right;
            band0.Columns["UseReward"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center;
            band0.Columns["ApplyTime"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center;



            ultraGrid1.DisplayLayout.Bands[0].Override.HeaderAppearance.TextHAlign = Infragistics.Win.HAlign.Left;

            ultraGrid1.DisplayLayout.Bands[0].Columns["SerialNum"].Hidden = true;
            this.ultraGrid1.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti;
            band0.SortedColumns.Clear();
            SetButton();
        }

        private void LoadData() {
            try {
                //dt.Rows.Clear();
                UidDeltaOne.Clear();
                Market30.Clear();

                //市值前30大
                string sql_M = $@"SELECT  DISTINCT  [UID]
                              FROM [WarrantAssistant].[dbo].[SpecialStock]
                              WHERE [Type] = 'M' AND [IsTop30] = 'Y'
                              AND [DataDate] >=(select Max([Datadate]) FROM [WarrantAssistant].[dbo].[SpecialStock])";
                DataTable dv_M = MSSQL.ExecSqlQry(sql_M, GlobalVar.loginSet.warrantassistant45);
                foreach (DataRow dr in dv_M.Rows)
                {
                    string uid = dr["UID"].ToString();
                    if (!Market30.Contains(uid))
                        Market30.Add(uid);
                }

                string sql_deltaone = $@"SELECT  [UnderlyingID], [KgiCallDeltaOne], [KgiPutDeltaOne], [KgiCallPutRatio]
                                        ,[AllCallDeltaOne], [AllPutDeltaOne], [KgiAllPutRatio], [YuanPutDeltaOne], [KgiPutNum]
                                        FROM [WarrantAssistant].[dbo].[WarrantIssueDeltaOne]
                                        WHERE [DateTime]='{DateTime.Today.ToString("yyyyMMdd")}'";
                DataTable dv_deltaone = MSSQL.ExecSqlQry(sql_deltaone, GlobalVar.loginSet.warrantassistant45);

                foreach (DataRow dr in dv_deltaone.Rows)
                {
                    string uid_ = dr["UnderlyingID"].ToString();
                    double Kgi_CallDeltaOne = Math.Round(Convert.ToDouble(dr["KgiCallDeltaOne"].ToString()), 2);
                    double Kgi_PutDeltaOne = Math.Round(Convert.ToDouble(dr["KgiPutDeltaOne"].ToString()), 2);
                    double Kgi_CallPutRatio = Convert.ToDouble(dr["KgiCallPutRatio"].ToString());
                    double All_CallDeltaOne = Math.Round(Convert.ToDouble(dr["AllCallDeltaOne"].ToString()), 2);
                    double All_PutDeltaOne = Math.Round(Convert.ToDouble(dr["AllPutDeltaOne"].ToString()), 2);
                    double KgiAll_PutRatio = Convert.ToDouble(dr["KgiAllPutRatio"].ToString());
                    double Yuan_PutDeltaOne = Convert.ToDouble(dr["YuanPutDeltaOne"].ToString());
                    int Kgi_PutNum = Convert.ToInt32(dr["KgiPutNum"].ToString());
                    if (!UidDeltaOne.ContainsKey(uid_))
                        UidDeltaOne.Add(uid_, new UidPutCallDeltaOne(uid_, Kgi_CallDeltaOne, Kgi_PutDeltaOne, Kgi_CallPutRatio, All_CallDeltaOne, All_PutDeltaOne, KgiAll_PutRatio, Yuan_PutDeltaOne, Kgi_PutNum));
                }
 
                string sql = $@"SELECT a.[SerialNum]
                              ,CASE WHEN a.[ApplyKind]='1' THEN '新發' ELSE '增額' END ApplyKind
                              ,a.[TraderID]
                              ,a.[Market]
                              ,a.[UnderlyingID]
                              ,a.[WarrantName]
                              ,a.[CR]
                              ,a.[IssueNum]
                              ,a.[EquivalentNum]
                              ,IsNull(a.[Result], 0) Result
                              ,IsNull(E.[CanIssue],0) Credit
                              ,IsNull(Floor(E.[WarrantAvailableShares] * {GlobalVar.globalParameter.givenRewardPercent} - IsNull(F.[UsedRewardNum],0)), 0) AS RewardCredit
                              ,a.[UseReward]   
                              ,d.[ApplyTime]
                          FROM [WarrantAssistant].[dbo].[ApplyTotalList] a 
                          LEFT JOIN [WarrantAssistant].[dbo].[WarrantUnderlyingSummary] b ON a.UnderlyingID=b.UnderlyingID
                          LEFT JOIN [Underlying_TraderIssue] c on a.UnderlyingID=c.UID 
                          LEFT JOIN [WarrantAssistant].[dbo].[Apply_71] d on a.[SerialNum] = d.[SerialNum] 
                          LEFT JOIN (SELECT [UID], [CanIssue], [WarrantAvailableShares] FROM [WarrantAssistant].[dbo].[WarrantUnderlyingCreditNew] WHERE [UpdateTime] > '{DateTime.Today.ToString("yyyyMMdd")}' ) as E on A.UnderlyingID = E.[UID]
                          LEFT JOIN [WarrantAssistant].[dbo].[WarrantReward] F on A.UnderlyingID=F.UnderlyingID
                          ORDER BY  a.Market desc, a.ApplyKind, a.SerialNum";

                dt = EDLib.SQL.MSSQL.ExecSqlQry(sql, GlobalVar.loginSet.warrantassistant45);
                
               
                ultraGrid1.DataSource = dt;

                dt.Columns[0].Caption = "序號";
                dt.Columns[1].Caption = "類型";
                dt.Columns[2].Caption = "交易員";
                dt.Columns[3].Caption = "市場";
                dt.Columns[4].Caption = "標的代號";
                dt.Columns[5].Caption = "權證名稱";
                dt.Columns[6].Caption = "行使比例";
                dt.Columns[7].Caption = "張數";
                dt.Columns[8].Caption = "約當張數";
                dt.Columns[9].Caption = "額度結果";
                dt.Columns[10].Caption = "今日額度";
                dt.Columns[11].Caption = "獎勵額度";
                dt.Columns[12].Caption = "使用獎勵";
                dt.Columns[13].Caption = "順位";

                foreach (DataRow row in dt.Rows) {
                    row["Result"] = Math.Round((double) row["Result"],3);
                    row["Credit"] = Math.Round((double) row["Credit"]);
                    row["RewardCredit"] = Math.Round((double) row["RewardCredit"]);
                    row["TraderID"] = row["TraderID"].ToString().TrimStart('0');
                }
                Iscompete.Clear();
      
                string sql2 = $@"SELECT [UnderlyingID]
                          FROM [WarrantAssistant].[dbo].[Apply_71]
                          WHERE len([OriApplyTime])> 0";
                DataTable dt2 = MSSQL.ExecSqlQry(sql2, GlobalVar.loginSet.warrantassistant45);

                foreach (DataRow dr in dt2.Rows)
                {
                    Iscompete.Add(dr["UnderlyingID"].ToString());
                }
   
                if (DateTime.Now.TimeOfDay.TotalMinutes >= GlobalVar.globalParameter.resultTime)
                {

                    string sql_maxOri = $@"SELECT Max([OriApplyTime]) AS OriTime
                                    FROM [WarrantAssistant].[dbo].[Apply_71]";
                    DataTable dt_maxOri = MSSQL.ExecSqlQry(sql_maxOri, GlobalVar.loginSet.warrantassistant45);

                    foreach (DataRow dr in dt_maxOri.Rows)
                    {
                        if (dr["OriTime"].ToString().Length > 0)
                        {
                            applyhour_compete = Convert.ToInt32(dr["OriTime"].ToString().Substring(0, 2));
                            applymin_compete = Convert.ToInt32(dr["OriTime"].ToString().Substring(3, 2));
                        }
                    }
                  
                    if ((applyhour_compete == 0) && (applymin_compete == 0))
                    {

                        string sql_maxApp = $@"SELECT MAX([ApplyTime]) AS AppTime
                                              FROM [WarrantAssistant].[dbo].[Apply_71]
                                              where [ApplyTime] <='10:30:00'";
                        DataTable dt_maxApp = MSSQL.ExecSqlQry(sql_maxApp, GlobalVar.loginSet.warrantassistant45);

                        foreach (DataRow dr in dt_maxApp.Rows)
                        {
                            applyhour_compete = Convert.ToInt32(dr["AppTime"].ToString().Substring(0, 2));
                            applymin_compete = Convert.ToInt32(dr["AppTime"].ToString().Substring(3, 2));
                        }
                    }
                   
                    changeTo_1w.Clear();

                    string sql_changeTo1w = $@"SELECT [SerialNumber] 
                                    FROM [WarrantAssistant].[dbo].[ApplyTotalRecord]
                                    WHERE [UpdateTime] >= CONVERT(VARCHAR, GETDATE(), 112) AND [ApplyKind] = '3'";
                    DataTable dt_changeTo1w = MSSQL.ExecSqlQry(sql_changeTo1w, GlobalVar.loginSet.warrantassistant45);
                    foreach (DataRow dr in dt_changeTo1w.Rows)
                    {
                        string serialnum = dr["SerialNumber"].ToString();
                        changeTo_1w.Add(serialnum);
                    }
  

                    string sqltemp = $@"SELECT A.[SerialNum]
                            ,CASE WHEN A.[ApplyKind]='1' THEN '新發' ELSE '增額' END ApplyKind
                            ,A.[UnderlyingID]
                            ,A.[CR]
                            ,A.[CP]
                            ,A.[IssueNum]
                            ,A.[UseReward] 
                            ,B.[ApplyTime]
                            ,B.[OriApplyTime]
                        FROM [WarrantAssistant].[dbo].[ApplyTotalList] A
                        LEFT JOIN [WarrantAssistant].[dbo].[Apply_71] B on A.[SerialNum] = B.[SerialNum] ";
                    DataTable dvtemp = MSSQL.ExecSqlQry(sqltemp, GlobalVar.loginSet.warrantassistant45);

                    foreach (DataRow dr in dvtemp.Rows)
                    {
                       
                        bool needadd = true;
                        string uid = dr["UnderlyingID"].ToString();
                        string cp = dr["CP"].ToString();
                        string isreward = dr["UseReward"].ToString();
                        string apytime = dr["ApplyTime"].ToString();//時間的全部字串
                        string applyTime = "";
                        if(apytime!=string.Empty)
                            applyTime  = dr["ApplyTime"].ToString().Substring(0, 2);//時間幾點
                        string oriapplyTime = dr["OriApplyTime"].ToString();
                        double cr = Convert.ToDouble(dr["CR"].ToString());
                        double issueNum = Convert.ToDouble(dr["IssueNum"].ToString());
                            
                        
                        if (UidDeltaOne.ContainsKey(uid))
                        {
                            if (applyTime == "22" || ((apytime.Length ==0) && Iscompete.Contains(uid) && isreward == "N"))
                                needadd = false;


                            if (needadd)
                            {
                                if (cp == "C")
                                {
                                    
                                    UidDeltaOne[uid].KgiCallDeltaOne += issueNum * cr;
                                    if (UidDeltaOne[uid].KgiPutNum == 0)
                                        UidDeltaOne[uid].KgiCallPutRatio = 100;
                                    else
                                        UidDeltaOne[uid].KgiCallPutRatio = Math.Round((double)UidDeltaOne[uid].KgiCallDeltaOne / (double)UidDeltaOne[uid].KgiPutDeltaOne, 4);
                                    UidDeltaOne[uid].AllCallDeltaOne += issueNum * cr;
                                    if (UidDeltaOne[uid].AllPutDeltaOne == 0)
                                        UidDeltaOne[uid].KgiAllPutRatio = 100;
                                    else
                                        UidDeltaOne[uid].KgiAllPutRatio = Math.Round((double)UidDeltaOne[uid].KgiPutDeltaOne / (double)UidDeltaOne[uid].AllPutDeltaOne, 4);
                                }
                                else
                                {
                                    UidDeltaOne[uid].KgiPutDeltaOne += issueNum * cr;
                                    UidDeltaOne[uid].KgiCallPutRatio = Math.Round((double)UidDeltaOne[uid].KgiCallDeltaOne / (double)UidDeltaOne[uid].KgiPutDeltaOne, 4);
                                    UidDeltaOne[uid].AllPutDeltaOne += issueNum * cr;
                                    UidDeltaOne[uid].KgiAllPutRatio = Math.Round((double)UidDeltaOne[uid].KgiPutDeltaOne / (double)UidDeltaOne[uid].AllPutDeltaOne, 4);
                                    UidDeltaOne[uid].KgiPutNum++;
                                }
                            }
                        }
                    }
                }
            } catch (Exception ex) {
                MessageBox.Show(ex.Message);
            }
        }
        private void LoadIsIndex()
        {
#if !To39
            string sql_indexTW50 = $@"SELECT [UID] FROM [newEDIS].[dbo].[IndexUnderlying]
                                      WHERE [IndexType] = 'TW50 Index' 
                                      AND [DataDate] = (SELECT MAX(DataDate) FROM [newEDIS].[dbo].[IndexUnderlying]
                                      WHERE [IndexType] = 'TW50 Index')";
            string sql_indexTWMC = $@"SELECT [UID] FROM [newEDIS].[dbo].[IndexUnderlying]
                                      WHERE [IndexType] = 'TWMC Index' 
                                      AND [DataDate] = (SELECT MAX(DataDate) FROM [newEDIS].[dbo].[IndexUnderlying]
                                      WHERE [IndexType] = 'TWMC Index')";
            string sql_indexGTSM50 = $@"SELECT [UID] FROM [newEDIS].[dbo].[IndexUnderlying]
                                      WHERE [IndexType] = 'GTSM50 Index' 
                                      AND [DataDate] = (SELECT MAX(DataDate) FROM [newEDIS].[dbo].[IndexUnderlying]
                                      WHERE [IndexType] = 'GTSM50 Index')";
            DataTable dv_indexTW50 = MSSQL.ExecSqlQry(sql_indexTW50, GlobalVar.loginSet.newEDIS);
            DataTable dv_indexTWMC = MSSQL.ExecSqlQry(sql_indexTWMC, GlobalVar.loginSet.newEDIS);
            DataTable dv_indexGTSM50 = MSSQL.ExecSqlQry(sql_indexGTSM50, GlobalVar.loginSet.newEDIS);
#else
            string sql_indexTW50 = $@"SELECT [UID] FROM [TwData].[dbo].[IndexUnderlying]
                                      WHERE [IndexType] = 'TW50 Index' 
                                      AND [DataDate] = (SELECT MAX(DataDate) FROM [TwData].[dbo].[IndexUnderlying]
                                      WHERE [IndexType] = 'TW50 Index')";
            string sql_indexTWMC = $@"SELECT [UID] FROM [TwData].[dbo].[IndexUnderlying]
                                      WHERE [IndexType] = 'TWMC Index' 
                                      AND [DataDate] = (SELECT MAX(DataDate) FROM [TwData].[dbo].[IndexUnderlying]
                                      WHERE [IndexType] = 'TWMC Index')";
            string sql_indexGTSM50 = $@"SELECT [UID] FROM [TwData].[dbo].[IndexUnderlying]
                                      WHERE [IndexType] = 'GTSM50 Index' 
                                      AND [DataDate] = (SELECT MAX(DataDate) FROM [TwData].[dbo].[IndexUnderlying]
                                      WHERE [IndexType] = 'GTSM50 Index')";
            DataTable dv_indexTW50 = MSSQL.ExecSqlQry(sql_indexTW50, GlobalVar.loginSet.twData);
            DataTable dv_indexTWMC = MSSQL.ExecSqlQry(sql_indexTWMC, GlobalVar.loginSet.twData);
            DataTable dv_indexGTSM50 = MSSQL.ExecSqlQry(sql_indexGTSM50, GlobalVar.loginSet.twData);
#endif


            foreach (DataRow dr in dv_indexTW50.Rows)
            {
                string uid = dr["UID"].ToString();
                if (!IsIndex.Contains(uid))
                    IsIndex.Add(uid);
            }
            foreach (DataRow dr in dv_indexTWMC.Rows)
            {
                string uid = dr["UID"].ToString();
                if (!IsIndex.Contains(uid))
                    IsIndex.Add(uid);
            }
            foreach (DataRow dr in dv_indexGTSM50.Rows)
            {
                string uid = dr["UID"].ToString();
                if (!IsIndex.Contains(uid))
                    IsIndex.Add(uid);
            }
        }

        private void LoadIsSpecial()
        {

            string sql_special = $@"SELECT  DISTINCT  [UID]
                                      FROM [WarrantAssistant].[dbo].[SpecialStock]
                                      WHERE [Type] in ('B','R','KY','O')
                                      AND [DataDate] >=(select Max([Datadate]) FROM [WarrantAssistant].[dbo].[SpecialStock])";
            DataTable dv_special = MSSQL.ExecSqlQry(sql_special, GlobalVar.loginSet.warrantassistant45);

            foreach (DataRow dr in dv_special.Rows)
            {
                string uid = dr["UID"].ToString();
                if (!IsSpecial.Contains(uid))
                    IsSpecial.Add(uid);
            }
        }
        private void UpdateData() {
            try {
                string cmdText = "UPDATE [ApplyTotalList] SET WarrantName=@WarrantName, CR=@CR, IssueNum=@IssueNum, EquivalentNum=@EquivalentNum, Result=@Result, UseReward=@UseReward WHERE SerialNum=@SerialNum";
                List<System.Data.SqlClient.SqlParameter> pars = new List<SqlParameter>();
                pars.Add(new SqlParameter("@WarrantName", SqlDbType.VarChar));
                pars.Add(new SqlParameter("@CR", SqlDbType.Float));
                pars.Add(new SqlParameter("@IssueNum", SqlDbType.Float));
                pars.Add(new SqlParameter("@EquivalentNum", SqlDbType.Float));
                pars.Add(new SqlParameter("@Result", SqlDbType.Float));
                pars.Add(new SqlParameter("@SerialNum", SqlDbType.VarChar));
                pars.Add(new SqlParameter("@UseReward", SqlDbType.VarChar));

                SQLCommandHelper h = new SQLCommandHelper(GlobalVar.loginSet.warrantassistant45, cmdText, pars);
                foreach (Infragistics.Win.UltraWinGrid.UltraGridRow r in ultraGrid1.Rows) {
                    string warrantName = r.Cells["WarrantName"].Value.ToString();
                    double cr = r.Cells["CR"].Value == DBNull.Value ? 0 : Convert.ToDouble(r.Cells["CR"].Value);
                    double issueNum = r.Cells["IssueNum"].Value == DBNull.Value ? 0 : Convert.ToDouble(r.Cells["IssueNum"].Value);
                    double equivalentNum = cr * issueNum;
                    double result = r.Cells["Result"].Value == DBNull.Value ? 0 : Convert.ToDouble(r.Cells["Result"].Value);
                    string useReward = r.Cells["UseReward"].Value.ToString();
                    string serialNum = r.Cells["SerialNum"].Value.ToString();

                    h.SetParameterValue("@WarrantName", warrantName);
                    h.SetParameterValue("@CR", cr);
                    h.SetParameterValue("@IssueNum", issueNum);
                    h.SetParameterValue("@EquivalentNum", equivalentNum);
                    h.SetParameterValue("@Result", result);
                    h.SetParameterValue("@UseReward", useReward);
                    h.SetParameterValue("@SerialNum", serialNum);
                    h.ExecuteCommand();
                    //------------------------

                    string sql = $@"SELECT TOP(1) [ToValue], [UpdateCount]
                                    FROM [WarrantAssistant].[dbo].[ApplyTotalRecord]
                                    WHERE[SerialNumber] ='{serialNum}' 
                                    AND[DataName] ='CR'
                                    ORDER BY [UpdateCount] desc";

                    conn.Open();
                    DataTable dv_temp = MSSQL.ExecSqlQry(sql, conn);
                    if (dv_temp.Rows.Count > 0)
                    {
                        DataRow dr_temp = dv_temp.Rows[0];
                        string str_temp = dr_temp["ToValue"].ToString();
                        if (!str_temp.Equals(cr.ToString()))//有資料更新
                        {
                            int count = Int32.Parse(dr_temp["UpdateCount"].ToString());

                            string sql2 = $@"INSERT INTO [WarrantAssistant].[dbo].[ApplyTotalRecord] ([UpdateTime], [UpdateType], [TraderID], [SerialNumber]
                                            , [ApplyKind], [DataName] ,[FromValue], [ToValue], [UpdateCount])
                                           VALUES(GETDATE(), 'UPDATE', '{userID}', {serialNum}, '1','CR','{str_temp}','{cr.ToString()}',{count + 1})";
                            MSSQL.ExecSqlCmd(sql2, conn);

                        }
                    }
                    //------------------------
 
                    conn.Close();
                }
                h.Dispose();

                string cmdText2 = "UPDATE [ApplyOfficial] SET R=@R, IssueNum=@IssueNum, UseReward=@UseReward WHERE SerialNumber=@SerialNumber";
                List<SqlParameter> pars2 = new List<SqlParameter>();

                pars2.Add(new SqlParameter("@R", SqlDbType.Float));
                pars2.Add(new SqlParameter("@IssueNum", SqlDbType.Float));
                pars2.Add(new SqlParameter("@SerialNumber", SqlDbType.VarChar));
                pars2.Add(new SqlParameter("@UseReward", SqlDbType.VarChar));

                SQLCommandHelper h2 = new SQLCommandHelper(GlobalVar.loginSet.warrantassistant45, cmdText2, pars2);

                foreach (Infragistics.Win.UltraWinGrid.UltraGridRow r in ultraGrid1.Rows) {
                    double cr = r.Cells["CR"].Value == DBNull.Value ? 0 : Convert.ToDouble(r.Cells["CR"].Value);
                    double issueNum = r.Cells["IssueNum"].Value == DBNull.Value ? 0 : Convert.ToDouble(r.Cells["IssueNum"].Value);
                    string serialNumber = r.Cells["SerialNum"].Value.ToString();
                    string useReward = r.Cells["UseReward"].Value.ToString();

                    h2.SetParameterValue("@R", cr);
                    h2.SetParameterValue("@IssueNum", issueNum);
                    h2.SetParameterValue("@SerialNumber", serialNumber);
                    h2.SetParameterValue("@UseReward", useReward);
                    h2.ExecuteCommand();
                }
                h2.Dispose();
                
                GlobalUtility.LogInfo("Info", GlobalVar.globalParameter.userID + " 更新搶額度總表");

            } catch (Exception ex) {
                MessageBox.Show(ex.Message);
            }
        }

        private void SetButton() {
            UltraGridBand band0 = ultraGrid1.DisplayLayout.Bands[0];
            if (isEdit) {
                band0.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.Default;
                band0.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
                band0.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.True;

                band0.Columns["ApplyKind"].CellAppearance.BackColor = Color.LightGray;
                band0.Columns["Market"].CellAppearance.BackColor = Color.LightGray;
                band0.Columns["UnderlyingID"].CellAppearance.BackColor = Color.LightGray;
                band0.Columns["EquivalentNum"].CellAppearance.BackColor = Color.LightGray;

                band0.Columns["WarrantName"].CellActivation = Activation.AllowEdit;
                band0.Columns["CR"].CellActivation = Activation.AllowEdit;
                band0.Columns["IssueNum"].CellActivation = Activation.AllowEdit;
                band0.Columns["Result"].CellActivation = Activation.AllowEdit;
                band0.Columns["UseReward"].CellActivation = Activation.AllowEdit;

                toolStripButtonReload.Visible = false;
                toolStripButtonEdit.Visible = false;
                toolStripButtonConfirm.Visible = true;
                toolStripButtonCancel.Visible = true;

                //ultraGrid1.DisplayLayout.Bands[0].Columns["ApplyKind"].Hidden = true;
                band0.Columns["TraderID"].Hidden = true;
                band0.Columns["Credit"].Hidden = true;
                band0.Columns["RewardCredit"].Hidden = true;

                ultraGrid1.DisplayLayout.AutoFitStyle = AutoFitStyle.ResizeAllColumns;

            } else {
                //ultraGrid1.DisplayLayout.Bands[0].Columns["ApplyKind"].Hidden = false;
                band0.Columns["TraderID"].Hidden = false;
                band0.Columns["Credit"].Hidden = false;
                band0.Columns["RewardCredit"].Hidden = false;

                band0.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
                band0.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
                band0.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;

                band0.Columns["ApplyKind"].CellAppearance.BackColor = Color.White;
                band0.Columns["Market"].CellAppearance.BackColor = Color.White;
                band0.Columns["UnderlyingID"].CellAppearance.BackColor = Color.White;
                band0.Columns["EquivalentNum"].CellAppearance.BackColor = Color.White;

                band0.Columns["ApplyKind"].CellActivation = Activation.NoEdit;
                band0.Columns["Market"].CellActivation = Activation.NoEdit;
                band0.Columns["TraderID"].CellActivation = Activation.NoEdit;
                band0.Columns["UnderlyingID"].CellActivation = Activation.NoEdit;
                band0.Columns["WarrantName"].CellActivation = Activation.NoEdit;
                band0.Columns["CR"].CellActivation = Activation.NoEdit;
                band0.Columns["IssueNum"].CellActivation = Activation.NoEdit;
                band0.Columns["EquivalentNum"].CellActivation = Activation.NoEdit;
                band0.Columns["Result"].CellActivation = Activation.NoEdit;
                band0.Columns["Credit"].CellActivation = Activation.NoEdit;
                band0.Columns["RewardCredit"].CellActivation = Activation.NoEdit;
                band0.Columns["UseReward"].CellActivation = Activation.NoEdit;
                band0.Columns["ApplyTime"].CellActivation = Activation.NoEdit;

                band0.Columns["ApplyKind"].Width = 70;
                band0.Columns["TraderID"].Width = 70;
                band0.Columns["Market"].Width = 70;
                band0.Columns["UnderlyingID"].Width = 70;
                band0.Columns["WarrantName"].Width = 150;
                band0.Columns["CR"].Width = 70;
                band0.Columns["IssueNum"].Width = 80;
                band0.Columns["EquivalentNum"].Width = 80;
                band0.Columns["Result"].Width = 80;
                band0.Columns["UseReward"].Width = 70;
                band0.Columns["ApplyTime"].Width = 60;
                ultraGrid1.DisplayLayout.AutoFitStyle = AutoFitStyle.ResizeAllColumns;

                toolStripButtonReload.Visible = true;
                toolStripButtonEdit.Visible = true;
                toolStripButtonConfirm.Visible = false;
                toolStripButtonCancel.Visible = false;

                if (GlobalVar.globalParameter.userGroup == "TR") {
                    toolStripButtonEdit.Visible = false;
                    刪除ToolStripMenuItem.Visible = false;
                }
            }
        }


        private void ultraGrid1_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e) {
            ultraGrid1.DisplayLayout.Override.RowSelectorHeaderStyle = RowSelectorHeaderStyle.ColumnChooserButton;
        }

        private void toolStripButtonEdit_Click(object sender, EventArgs e) {
            LoadData();//畫面可能是舊的資料，要先load新資料再改
            isEdit = true;
            SetButton();
        }

        private void toolStripButtonConfirm_Click(object sender, EventArgs e) {
            ultraGrid1.PerformAction(Infragistics.Win.UltraWinGrid.UltraGridAction.ExitEditMode);
            isEdit = false;
            UpdateData();
            SetButton();
            LoadData();
        }

        private void toolStripButtonCancel_Click(object sender, EventArgs e) {
            isEdit = false;
            LoadData();
            SetButton();
        }

        private void ultraGrid1_InitializeRow(object sender, InitializeRowEventArgs e) { 
            string applyKind = e.Row.Cells["ApplyKind"].Value.ToString();
            string warrantName = e.Row.Cells["WarrantName"].Value.ToString();
            string serialNum = e.Row.Cells["SerialNum"].Value.ToString();
            string applyStatus = "";
            string applyTime = "";
            string apytime = "";
            string rank = "";
            string oriapplyTime = "";
            double issueNum = Convert.ToDouble(e.Row.Cells["IssueNum"].Value);
            
            double equivalentNum = Convert.ToDouble(e.Row.Cells["EquivalentNum"].Value);
            double result = e.Row.Cells["Result"].Value == DBNull.Value ? 0 : Convert.ToDouble(e.Row.Cells["Result"].Value);
            double cr = e.Row.Cells["CR"].Value == DBNull.Value ? 0 : Convert.ToDouble(e.Row.Cells["CR"].Value);
            string isReward = e.Row.Cells["UseReward"].Value.ToString();
            

            if (isReward == "Y")
                e.Row.Cells["UseReward"].Appearance.ForeColor = Color.Blue;

            if (!isEdit && DateTime.Now.TimeOfDay.TotalMinutes >= GlobalVar.globalParameter.resultTime) {//10:40

                string sqlTemp = "SELECT [ApplyStatus],[ApplyTime], [OriApplyTime] FROM [WarrantAssistant].[dbo].[Apply_71] WHERE SerialNum = '" + serialNum + "'";
                DataTable dtTemp = EDLib.SQL.MSSQL.ExecSqlQry(sqlTemp, GlobalVar.loginSet.warrantassistant45);


                foreach (DataRow drTemp in dtTemp.Rows) {
                    applyStatus = drTemp["ApplyStatus"].ToString();
                    applyTime = drTemp["ApplyTime"].ToString().Substring(0, 2);//時間幾點
                    apytime = drTemp["ApplyTime"].ToString();//時間的全部字串
                    rank = drTemp["ApplyTime"].ToString().Substring(6, 2);
                    oriapplyTime = drTemp["OriApplyTime"].ToString();
                }
                /*
                if (applyTime == "10" && applyStatus != "X 沒額度" && issueNum != 10000) {
                    e.Row.Cells["IssueNum"].Appearance.ForeColor = Color.Red;
                    e.Row.Cells["IssueNum"].ToolTipText = "applyTime == 10 && applyStatus != 沒額度 && issueNum != 10000";
                }
                */
                string useReward = e.Row.Cells["UseReward"].Value.ToString();
                string UID = e.Row.Cells["UnderlyingID"].Value.ToString();
                if (apytime.Length >0)//7-1表中有的權證
                {
                    int hour = Convert.ToInt32(apytime.Substring(0, 2));
                    int min = Convert.ToInt32(apytime.Substring(3, 2));
                    //要搶
                    if (applyTime == "10" && oriapplyTime.Length > 0)
                    {
                        e.Row.Cells["ApplyTime"].Appearance.BackColor = Color.Bisque;
                        e.Row.Cells["ApplyTime"].Value = apytime.Substring(6, 2);
                        e.Row.Cells["ApplyTime"].ToolTipText = "需搶額度";
                    }
                    else if (applyTime == "22")
                    {
                        e.Row.Cells["ApplyTime"].Appearance.BackColor = Color.Aquamarine;
                        e.Row.Cells["ApplyTime"].Value = "X";
                        e.Row.Cells["ApplyTime"].ToolTipText = "沒額度";
                    }
                    else//不用搶的
                    {
                        e.Row.Cells["ApplyTime"].Appearance.BackColor = Color.Aquamarine;
                        e.Row.Cells["ApplyTime"].Value = "-";
                        e.Row.Cells["ApplyTime"].ToolTipText = "不用搶";
                    }
                    if(applyKind =="增額" && oriapplyTime.Length > 0)
                    {
                        e.Row.Cells["ApplyKind"].Appearance.BackColor = Color.Gold;
                        e.Row.Cells["Market"].Appearance.BackColor = Color.Gold;
                        e.Row.Cells["UnderlyingID"].Appearance.BackColor = Color.Gold;
                        e.Row.Cells["TraderID"].Appearance.BackColor = Color.Gold;
                        e.Row.Cells["WarrantName"].Appearance.BackColor = Color.Gold;
                        e.Row.Cells["CR"].Appearance.BackColor = Color.Gold;
                        e.Row.Cells["IssueNum"].Appearance.BackColor = Color.Gold;
                        e.Row.Cells["ApplyTime"].Appearance.BackColor = Color.Gold;
                        e.Row.Cells["ApplyTime"].Value = "搶增";
                        e.Row.Cells["ApplyTime"].ToolTipText = "搶發增額標的";
                    }
                }
                else
                {
                    
                    if (Iscompete.Contains(UID) && (isReward=="N"))
                    {
                        e.Row.Cells["ApplyTime"].Appearance.BackColor = Color.Aquamarine;
                        e.Row.Cells["ApplyTime"].Value = "X";
                        e.Row.Cells["ApplyTime"].ToolTipText = "沒額度";
                    }
                    else//要判斷是加發，還是後來選擇不發，用申請時間判斷
                    {
                        int h1 = 0;
                        int m1 = 0;
                        //string sql_MDate = $@"SELECT [MDate] 
                        //                    FROM [EDIS].[dbo].[ApplyTotalList]
                        //                    WHERE [SerialNum] = '{serialNum}'";

                        string sql_MDate = $@"SELECT Min([UpdateTime]) AS MDate
                                            FROM [WarrantAssistant].[dbo].[ApplyTotalRecord]
                                            WHERE [SerialNumber] = '{serialNum}' AND [UpdateCount] = 0";
                        //申請時間如果在10:30申請前的最後一刻之後，則視為後來加發
                        DataTable dt_MDate = MSSQL.ExecSqlQry(sql_MDate, GlobalVar.loginSet.warrantassistant45);

                        foreach (DataRow dr in dt_MDate.Rows)
                        {
                            bool f = DateTime.TryParse(dr["MDate"].ToString(), out DateTime t1);
                            h1 = Convert.ToInt32(t1.ToString("yyyyMMddHHmmss").Substring(8, 2));
                            m1 = Convert.ToInt32(t1.ToString("yyyyMMddHHmmss").Substring(10, 2));
                        }
                        if ((h1 > applyhour_compete) || (h1 == applyhour_compete) && (m1 > applymin_compete))
                        {
                            e.Row.Cells["ApplyKind"].Appearance.BackColor = Color.LightCoral;
                            e.Row.Cells["Market"].Appearance.BackColor = Color.LightCoral;
                            e.Row.Cells["UnderlyingID"].Appearance.BackColor = Color.LightCoral;
                            e.Row.Cells["TraderID"].Appearance.BackColor = Color.LightCoral;
                            e.Row.Cells["WarrantName"].Appearance.BackColor = Color.LightCoral;
                            e.Row.Cells["CR"].Appearance.BackColor = Color.LightCoral;
                            e.Row.Cells["IssueNum"].Appearance.BackColor = Color.LightCoral;
                            e.Row.Cells["ApplyTime"].Appearance.BackColor = Color.Aquamarine;
                            e.Row.Cells["ApplyTime"].Value = "加";
                            e.Row.Cells["ApplyTime"].ToolTipText = "加發，沒在7-1表中";
                        }
                        else if (isReward=="Y")//在早上第一批就有發，但後來改成獎勵
                        {
                            e.Row.Cells["ApplyKind"].Appearance.BackColor = Color.LightCoral;
                            e.Row.Cells["Market"].Appearance.BackColor = Color.LightCoral;
                            e.Row.Cells["UnderlyingID"].Appearance.BackColor = Color.LightCoral;
                            e.Row.Cells["TraderID"].Appearance.BackColor = Color.LightCoral;
                            e.Row.Cells["WarrantName"].Appearance.BackColor = Color.LightCoral;
                            e.Row.Cells["CR"].Appearance.BackColor = Color.LightCoral;
                            e.Row.Cells["IssueNum"].Appearance.BackColor = Color.LightCoral;
                            e.Row.Cells["ApplyTime"].Appearance.BackColor = Color.Aquamarine;
                            e.Row.Cells["ApplyTime"].Value = "加";
                            e.Row.Cells["ApplyTime"].ToolTipText = "改用獎勵發，沒在7-1表中";
                        }
                        else
                        {
                            e.Row.Cells["ApplyTime"].Appearance.BackColor = Color.Aquamarine;
                            e.Row.Cells["ApplyTime"].Value = "X";
                            e.Row.Cells["ApplyTime"].ToolTipText = "後來不發";
                        }
                            
                    }
                    //兩種情況1.加發但還沒有進7-1 2.沒搶到的權證也不會在7-1
                }
                
                if((changeTo_1w.Contains(serialNum) && isReward == "N") || (Iscompete.Contains(UID) && isReward == "Y"))
                {
                    if (UidDeltaOne.ContainsKey(UID))
                    {
                        int wlength = warrantName.Length;
                        string subwname = warrantName.Substring(0, wlength - 2);

                        string sql_price = $@"SELECT [MPrice]
                                            FROM [WarrantAssistant].[dbo].[WarrantPrices]
                                            WHERE [CommodityID] = '{UID}'";
                        DataTable dt_price = MSSQL.ExecSqlQry(sql_price, GlobalVar.loginSet.warrantassistant45);

                        double spot = 0;
                        foreach(DataRow dr in dt_price.Rows)
                        {
                            spot = Convert.ToDouble(dr["MPrice"].ToString());
                        }
                        issueNum = Convert.ToDouble(e.Row.Cells["IssueNum"].Value);

                        if (IsSpecial.Contains(UID) && (subwname.EndsWith("購") || subwname.EndsWith("牛")))//如果是Call  要考慮Put=0的情況 
                        {
                            if (Market30.Contains(UID) )//市值前30  DeltaOne*股價<5億
                            {
                                if (cr * issueNum * spot > ISTOP30MaxIssue)
                                {
                                    e.Row.Cells["CR"].Appearance.BackColor = Color.DimGray;
                                    e.Row.Cells["CR"].Appearance.ForeColor = Color.Gold;
                                    e.Row.Cells["CR"].ToolTipText = $@"為市值前30大標的，DeltaOne市值已超過{(int)(ISTOP30MaxIssue / 100000)}億\n";
                                }
                            }
                            else
                            {
                                if (cr * issueNum * spot > NonTOP30MaxIssue)
                                {
                                    e.Row.Cells["CR"].Appearance.BackColor = Color.DimGray;
                                    e.Row.Cells["CR"].Appearance.ForeColor = Color.Gold;
                                    e.Row.Cells["CR"].ToolTipText = $@"DeltaOne市值已超過{(int)(NonTOP30MaxIssue / 100000)}億\n";
                                }
                            }
                        }
                        if (subwname.EndsWith("售") || subwname.EndsWith("熊"))
                        {
                            if (IsSpecial.Contains(UID))
                            {
                                if (Market30.Contains(UID))//市值前30  DeltaOne*股價<5億
                                {
                                    if (cr * issueNum * spot > ISTOP30MaxIssue)
                                    {
                                        e.Row.Cells["CR"].Appearance.BackColor = Color.DimGray;
                                        e.Row.Cells["CR"].Appearance.ForeColor = Color.Gold;
                                        e.Row.Cells["CR"].ToolTipText = $@"為市值前30大標的，DeltaOne市值已超過{(int)(ISTOP30MaxIssue / 100000)}億\n";
                                    }
                                }
                                else
                                {
                                    if (cr * issueNum * spot > NonTOP30MaxIssue)
                                    {
                                        e.Row.Cells["CR"].Appearance.BackColor = Color.DimGray;
                                        e.Row.Cells["CR"].Appearance.ForeColor = Color.Gold;
                                        e.Row.Cells["CR"].ToolTipText = $@"DeltaOne市值已超過{(int)(NonTOP30MaxIssue / 100000)}億\n";
                                    }
                                }
                            }
                            if(IsSpecial.Contains(UID) && (double)UidDeltaOne[UID].KgiCallDeltaOne / (double)UidDeltaOne[UID].KgiPutDeltaOne < SpecialCallPutRatio)
                            {
                                e.Row.Cells["CR"].Appearance.BackColor = Color.DimGray;
                                e.Row.Cells["CR"].Appearance.ForeColor = Color.Gold;
                                e.Row.Cells["CR"].ToolTipText += $@"為風險標的，自家權證 Call/Put DeltaOne比例 < {SpecialCallPutRatio}\n";
                            }
                            else if ((double)UidDeltaOne[UID].KgiCallDeltaOne / (double)UidDeltaOne[UID].KgiPutDeltaOne < NonSpecialCallPutRatio)
                            {
                                if (!IsIndex.Contains(UID))
                                {
                                    e.Row.Cells["CR"].Appearance.BackColor = Color.DimGray;
                                    e.Row.Cells["CR"].Appearance.ForeColor = Color.Gold;
                                    e.Row.Cells["CR"].ToolTipText += $@"自家權證 Call/Put DeltaOne比例 < {NonSpecialCallPutRatio}\n";
                                }
                            }
                            if (IsSpecial.Contains(UID) && UidDeltaOne[UID].AllPutDeltaOne > 0 && UidDeltaOne[UID].KgiAllPutRatio > SpecialKGIALLPutRatio)
                            {
                                //若之前這檔標的沒發過Put可以跳過，可是要考慮今天發超過一檔
                                if (UidDeltaOne[UID].KgiPutNum > 1)
                                {
                                    e.Row.Cells["CR"].Appearance.BackColor = Color.DimGray;
                                    e.Row.Cells["CR"].Appearance.ForeColor = Color.Gold;
                                    e.Row.Cells["CR"].ToolTipText += $@"自家/市場 Put DeltaOne比例 > {SpecialKGIALLPutRatio}\n";
                                }
                            }
                        }
                    }
                    
                }
                if (applyStatus == "X 沒額度") {
                    e.Row.Cells["EquivalentNum"].Appearance.BackColor = Color.LightGray;
                    e.Row.Cells["Result"].Appearance.BackColor = Color.LightGray;
                    e.Row.Cells["EquivalentNum"].ToolTipText = "沒額度";
                    e.Row.Cells["Result"].ToolTipText = "沒額度";
                }

                //double precision issue
                if (result + 0.00001 >= equivalentNum) {
                    e.Row.Cells["EquivalentNum"].Appearance.BackColor = Color.PaleGreen;
                    e.Row.Cells["Result"].Appearance.BackColor = Color.PaleGreen;
                    e.Row.Cells["EquivalentNum"].ToolTipText = "額度OK";
                    e.Row.Cells["Result"].ToolTipText = "額度OK";
                }

                if (result + 0.00001 < equivalentNum && result > 0) {
                    e.Row.Cells["EquivalentNum"].Appearance.BackColor = Color.PaleTurquoise;
                    e.Row.Cells["Result"].Appearance.BackColor = Color.PaleTurquoise;
                    e.Row.Cells["EquivalentNum"].ToolTipText = "部分額度";
                    e.Row.Cells["Result"].ToolTipText = "部分額度";
                }
            } else
                e.Row.Appearance.BackColor = Color.White;

            string underlyingID = e.Row.Cells["UnderlyingID"].Value.ToString();           
            string issuable = "NA";
            string accNI = "N";
            string reissuable = "NA";

            string toolTip1 = "標的發行檢查=N";
            string toolTip2 = "非本季標的";
            string toolTip3 = "標的虧損";
            string toolTip4 = "今日未達增額標準";

            string sqlTemp2 = "SELECT IsNull(Issuable,'NA') Issuable, CASE WHEN [AccNetIncome]<0 THEN 'Y' ELSE 'N' END AccNetIncome FROM [WarrantAssistant].[dbo].[WarrantUnderlyingSummary] WHERE UnderlyingID = '" + underlyingID + "'";
            DataTable dtTemp2 = EDLib.SQL.MSSQL.ExecSqlQry(sqlTemp2, GlobalVar.loginSet.warrantassistant45);

            if (dtTemp2.Rows.Count > 0) {
                issuable = dtTemp2.Rows[0]["Issuable"].ToString();
                accNI = dtTemp2.Rows[0]["AccNetIncome"].ToString();
                /*foreach (DataRow drTemp2 in dtTemp2.Rows) {
                    issuable = drTemp2["Issuable"].ToString();
                    accNI = drTemp2["AccNetIncome"].ToString();
                }*/
                }


                if (!isEdit) {
                if (accNI == "Y" && result >= equivalentNum) {
                    e.Row.Cells["UnderlyingID"].ToolTipText = toolTip3;
                    e.Row.Cells["UnderlyingID"].Appearance.ForeColor = Color.Blue;
                }

                if (issuable == "NA") {
                    e.Row.ToolTipText = toolTip2;
                    e.Row.Appearance.ForeColor = Color.Red;
                } else if (issuable == "N") {
                    e.Row.ToolTipText = toolTip1;
                    e.Row.Cells["UnderlyingID"].Appearance.ForeColor = Color.Red;
                }
            }

            if (applyKind == "增額") {

                string sqlTemp3 = "SELECT IsNull([ReIssuable],'NA') ReIssuable FROM [WarrantAssistant].[dbo].[WarrantReIssuable] WHERE WarrantName = '" + warrantName + "'";
                DataTable dtTemp3 = EDLib.SQL.MSSQL.ExecSqlQry(sqlTemp3, GlobalVar.loginSet.warrantassistant45);

                if (dtTemp3.Rows.Count > 0)
                    reissuable = dtTemp3.Rows[0]["ReIssuable"].ToString();
                //foreach (DataRow drTemp3 in dtTemp3.Rows)
                //    reissuable = drTemp3["ReIssuable"].ToString();

                if (!isEdit && reissuable == "NA") {
                    e.Row.Cells["WarrantName"].ToolTipText = toolTip4;
                    e.Row.Cells["WarrantName"].Appearance.ForeColor = Color.Red;
                }
            }
        }

        private void toolStripButtonReload_Click(object sender, EventArgs e) {
            LoadData();
        }

        private void ultraGrid1_AfterCellUpdate(object sender, CellEventArgs e) {
            if (e.Cell.Column.Key == "CR" || e.Cell.Column.Key == "IssueNum") {
                double cr = e.Cell.Row.Cells["CR"].Value == DBNull.Value ? 0 : Convert.ToDouble(e.Cell.Row.Cells["CR"].Value);
                double issueNum = e.Cell.Row.Cells["IssueNum"].Value == DBNull.Value ? 0 : Convert.ToDouble(e.Cell.Row.Cells["IssueNum"].Value);
                double equivalentNum = cr * issueNum;
                e.Cell.Row.Cells["EquivalentNum"].Value = equivalentNum;
            }
        }

        private void ultraGrid1_DoubleClickCell(object sender, DoubleClickCellEventArgs e) {
            if (e.Cell.Column.Key == "UnderlyingID")
                GlobalUtility.MenuItemClick<FrmIssueCheck>().SelectUnderlying((string) e.Cell.Value);
        }

        private void ultraGrid1_MouseDown(object sender, MouseEventArgs e) {
            if (e.Button == MouseButtons.Right)
                contextMenuStrip1.Show();
        }

        private void 刪除ToolStripMenuItem_Click(object sender, EventArgs e) {
            string applyKind = ultraGrid1.ActiveRow.Cells["ApplyKind"].Value.ToString();
            string warrantName = ultraGrid1.ActiveRow.Cells["WarrantName"].Value.ToString();

            DialogResult result = MessageBox.Show("刪除此檔 " + applyKind + warrantName + "，確定?", "刪除資料", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
            if (result == DialogResult.Yes) {
                string serialNum = ultraGrid1.ActiveRow.Cells["SerialNum"].Value.ToString();               

                conn.Open();
                EDLib.SQL.MSSQL.ExecSqlCmd("DELETE FROM [ApplyTotalList] WHERE SerialNum='" + serialNum + "'", conn);
                if (applyKind == "新發")
                    EDLib.SQL.MSSQL.ExecSqlCmd("DELETE FROM [ApplyOfficial] WHERE SerialNumber='" + serialNum + "'", conn);
                else if (applyKind == "增額")
                    EDLib.SQL.MSSQL.ExecSqlCmd("DELETE FROM [ReIssueOfficial] WHERE SerialNum='" + serialNum + "'", conn);
                conn.Close();

                LoadData();

                GlobalUtility.LogInfo("Info", GlobalVar.globalParameter.userID + " 刪除一檔" + applyKind + ": " + warrantName);
            }
        }
    }
}

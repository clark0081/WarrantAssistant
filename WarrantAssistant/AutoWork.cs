using System;
using System.Threading;


namespace WarrantAssistant
{
    public class AutoWork : IDisposable
    {
        private Thread workThread;
        private bool globalDataOK = true;
        

        public AutoWork()
        {
            workThread = new Thread(new ThreadStart(Working));
            workThread.Start();
        }
    
        private void Working()
        {
            try
            {
                for (; ; )
                {
                    DateTime now = DateTime.Now;

                    if (now.TimeOfDay.TotalSeconds > 10 && now.TimeOfDay.TotalSeconds < 30)
                        globalDataOK = false;

                    if (now.TimeOfDay.TotalSeconds > 60 && (!globalDataOK))
                    {
                        GlobalUtility.LoadGlobalParameters();
                        globalDataOK = true;
                    }
                    /*
                    if (GlobalVar.globalParameter.isTodayTradeDate)
                    {
                         GlobalVar.mainForm.AddWork(new InfoWork("資訊更新"));
                         GlobalVar.mainForm.AddWork(new AnnounceWork("公告更新"));
                    }
                     * */
                    Thread.Sleep(5000);

                }
            }
            catch (Exception ex)
            {
                //MessageBox.Show(ex.Message);
            }
        }

        #region IDisposable成員

        public void Dispose()
        {
            if (workThread != null && workThread.IsAlive) { workThread.Abort(); }
        }

        #endregion
    
    }
}

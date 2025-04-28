using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
//using Microsoft.Office.Interop.Excel;
//using Microsoft.Office.Interop;
using System.Data.SqlClient;
using System.Data;
using System.IO;
//using Oracle.DataAccess.Client;
using System.Data.OleDb;
using System.Windows.Forms;
using System.Configuration;
using System.ComponentModel;
using System.Threading;
using System.Drawing;

namespace WarrantAssistant
{
    public enum WorkState { Successful = 0, Exception = 1, Failed = 2 }

    public class Work
    {
        public string workName = "";
        public DateTime doWorkTime;

        public Work(string workName)
        {
            this.workName = workName;
        }

        public virtual WorkState DoWork() { return WorkState.Successful; }
        public virtual void Close() { }
    }

    public class InfoWork : Work
    {
        //private MainForm mainform;
        public InfoWork(string workName)
            : base(workName)
        {
            //mainform = new MainForm();
        }

        public override WorkState DoWork()
        {
            try
            {
                if (GlobalVar.mainForm != null)
                {
                    //GlobalVar.mainForm.SetUltraGrid1();
                    GlobalVar.mainForm.LoadUltraGrid1();
                    return WorkState.Successful;
                }
                else
                    return WorkState.Failed;
            }
            catch (Exception ex)
            {
                GlobalUtility.LogInfo("Error", "InfoWork Error: " + ex.Message);              

                return WorkState.Exception;
            }
        }

        public override void Close()
        {
            base.Close();
        }
    }

    public class AnnounceWork : Work
    {
        //private MainForm mainform;
        public AnnounceWork(string workName)
            : base(workName)
        {
            //mainform = new MainForm();
        }

        public override WorkState DoWork()
        {
            try
            {
                if (GlobalVar.mainForm != null)
                {
                    //GlobalVar.mainForm.SetUltraGrid2();
                    GlobalVar.mainForm.LoadUltraGrid2();
                    return WorkState.Successful;
                }
                else
                    return WorkState.Failed;
            }
            catch (Exception ex)
            {
                GlobalUtility.LogInfo("Error", "AnnounceWork Error: " + ex.Message);             
                return WorkState.Exception;
            }
        }

        public override void Close()
        {
            base.Close();
        }
    }
}

using System;
using System.Data;
using System.Drawing;
using System.Windows.Forms;
using EDLib.SQL;
using HtmlAgilityPack;
using System.IO;
using System.Linq;
using System.Text;
using System.Net;
using System.Threading.Tasks;

namespace WarrantAssistant
{
    public partial class FrmRename:Form
    {
        public FrmRename() {
            InitializeComponent();
        }
        private void ultraGrid1_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e) {
            e.Layout.Override.CellAppearance.BackColor = Color.LightCyan;
        }

        string path;// = "C:\\WarrantDocuments";

        private async void FrmRename_Load(object sender, EventArgs e) {
            //toolStripLabel1.Text = "";
            toolStrip1.Enabled = true;
            RenameFiles.Enabled = false;
            RenameFilesB.Enabled = false;
            RenameXML.Enabled = false;
            toolStripButton1.Enabled = false;
            toolStripButton2.Enabled = true;
            DataTable dt = MSSQL.ExecSqlQry("SELECT FLGDAT_FLGDSC FROM EDAISYS.dbo.FLAGDATAS WHERE FLGDAT_FLGNAM = 'DOCUMENT_SETTING' and FLGDAT_FLGVAR='EXPORT_PATH'"
               , GlobalVar.loginSet.warrantSysKeySqlConnString);
            path = dt.Rows[0]["FLGDAT_FLGDSC"].ToString().TrimEnd('\\');
            //Get key and id          
            string key = GlobalUtility.GetKey();
            string id = GlobalUtility.GetID();
            if (id == "10387")
            {
                if (!await ParseHtml($"https://siis.twse.com.tw/server-java/t150sa03?step=0&id=9200h{id}&TYPEK=sii&key={key}", true)
                    & !await ParseHtml($"https://siis.twse.com.tw/server-java/o_t150sa03?step=0&id=9200h{id}&TYPEK=otc&key={key}", false))
                    return;
            }
            else
            {
                if (!await ParseHtml($"https://siis.twse.com.tw/server-java/t150sa03?step=0&id=9200pd{id}&TYPEK=sii&key={key}", true)
                    & !await ParseHtml($"https://siis.twse.com.tw/server-java/o_t150sa03?step=0&id=9200pd{id}&TYPEK=otc&key={key}", false))
                    return;
            }
            // Get document path
            
            RenameFiles.Enabled = true;
            RenameFilesB.Enabled = true;
            RenameXML.Enabled = true;
            toolStripButton1.Enabled = true;
            //toolStrip1.Enabled = true;
        }

        private async Task<bool> ParseHtml(string url, bool isTwse) {
            
            try {
                
                string firstResponse = await GlobalUtility.GetHtmlAsync(url, Encoding.Default);
                HtmlAgilityPack.HtmlDocument doc = new HtmlAgilityPack.HtmlDocument();
                doc.LoadHtml(firstResponse);
 
                HtmlNodeCollection navNodeChild = doc.DocumentNode.SelectSingleNode("//table[1]/tr[1]/td/table").ChildNodes;
                
                for (int i = 3; i < navNodeChild.Count; i += 2) {
                    string[] split = navNodeChild[i].InnerText.Split(new string[] { " ", "\t", "&nbsp;", "\n" ,"\r"}, StringSplitOptions.RemoveEmptyEntries);
                    ultraDataSource1.Rows.Add();
                   
                    ultraDataSource1.Rows[ultraDataSource1.Rows.Count - 1]["WName"] = split[1];
                   
                    ultraDataSource1.Rows[ultraDataSource1.Rows.Count - 1]["SerialNumber"] = split[0];
                    
                    ultraDataSource1.Rows[ultraDataSource1.Rows.Count - 1]["Market"] = isTwse ? "TWSE" : "OTC";
                }
                
                return true;
               
            } catch (WebException) {
                if (isTwse)
                    MessageBox.Show("可能要更新Key，或是網頁有問題");
                return false;
            } catch (NullReferenceException) {
                if (isTwse)
                    MessageBox.Show($@"TWSE 沒有資料");
                else
                    MessageBox.Show("OTC 沒有資料");
                return false;
            } catch (Exception e) {
                MessageBox.Show(e.ToString()); //
                return false;
            }
        }

        private void RenameFiles_Click(object sender, EventArgs e) {
            Rename("Renamed");
        }
        private void RenameFilesB_Click(object sender, EventArgs e) {
            Rename("RenamedB");
        }

        private void RenameXML_Click(object sender, EventArgs e) {
            Rename("Xml");
        }

        private void Rename(string type) {
            string now = DateTime.Now.ToString("yyyyMMdd-HHmmss");
           
            MessageBox.Show(path);
            if (Directory.Exists($"{path}\\{type}{now}"))
                Directory.Delete($"{path}\\{type}{now}", true);
            Directory.CreateDirectory($"{path}\\{type}{now}");
            if (Directory.Exists($"{path}\\{type}OTC{now}"))
                Directory.Delete($"{path}\\{type}OTC{now}", true);
            Directory.CreateDirectory($"{path}\\{type}OTC{now}");
            if(type == "SalesAnnounce")
            {
                MSSQL.ExecSqlQry($@"DELETE FROM [WarrantAssistant].[dbo].[WarrantIssueAnnounce]",GlobalVar.loginSet.warrantassistant45);
            }
            if (type == "ListedAnnounce")
            {
                string sql = "SELECT TOP(1) CONVERT(VARCHAR,TDate,112) AS TDate FROM [WarrantAssistant].[dbo].[WarrantIssueAnnounce]";
                DataTable dt = MSSQL.ExecSqlQry(sql, GlobalVar.loginSet.warrantassistant45);
                if(dt.Rows.Count <= 0)
                {
                    MessageBox.Show("沒有存續號");
                    return;
                }
                else
                {
                    string t = dt.Rows[0][0].ToString();
                    DateTime rt = DateTime.ParseExact(t, "yyyyMMdd", null);
                    if(rt!=DateTime.Today && rt != EDLib.TradeDate.LastNTradeDate(1))
                    {
                        MessageBox.Show("序號時間不對!");
                        return;
                    }
                }
                
                string listedsql = $@"SELECT [WName], [SerialNum], [Market]
                                                    FROM [WarrantAssistant].[dbo].[WarrantIssueAnnounce]";
                DataTable listeddt = MSSQL.ExecSqlQry(listedsql, GlobalVar.loginSet.warrantassistant45);
                foreach(DataRow dr in listeddt.Rows)
                {
                    string wname = dr["WName"].ToString();
                    string serial = dr["SerialNum"].ToString();
                    string market = dr["Market"].ToString();
                    
                    if (Directory.Exists($"{path}\\{wname}"))
                    {
                        foreach (string file in Directory.GetFiles($"{path}\\{wname}"))
                        {
                            string toFile = null;
                            string fileName = Path.GetFileName(file);
                            if (fileName.Contains("上市公告"))
                            {
                                int index = fileName.IndexOf('_');
                                string wname2 = fileName.Substring(index + 1, fileName.Length - index - 1);
                                if (market == "TWSE")
                                {
                                    toFile = $"{path}\\ListedAnnounce{now}\\{serial}-2-{wname2}";
                                }
                                else if (market == "OTC" )
                                {
                                    //toFile = $"{path}\\ListedAnnounceOTC{now}\\{serial}-2-{wname2}";
                                    //2021/5/14 OTC開放整批上傳，有改檔名
                                    toFile = $"{path}\\ListedAnnounceOTC{now}\\O-{serial}-2-{wname2}";
                                }
                                File.Copy(file, toFile, true);
                            }
                        }
                    }
                }
                
                if (!Directory.EnumerateFileSystemEntries($"{path}\\{type}{now}").Any())
                    Directory.Delete($"{path}\\{type}{now}");
                if (!Directory.EnumerateFileSystemEntries($"{path}\\{type}OTC{now}").Any())
                    Directory.Delete($"{path}\\{type}OTC{now}");
                    
                toolStripTextBox1.Text = DateTime.Now + "修改上市公告完成";
                return;
            }
            for (int i = 0; i < ultraDataSource1.Rows.Count; i++)
                if (Directory.Exists($"{path}\\{ultraDataSource1.Rows[i]["WName"]}")) {
                    //string[] files = Directory.GetFiles($"{path}\\{ultraDataSource1.Rows[i]["WName"]}");
                    if(type == "SalesAnnounce")
                    {
                        using (StreamWriter sw = new StreamWriter($@"{path}\\{DateTime.Today.ToString("yyyyMMdd")}序號.txt", true))
                        {
                            sw.WriteLine($@"{ultraDataSource1.Rows[i]["WName"]},{ultraDataSource1.Rows[i]["SerialNumber"]}");
                        }
                        string sql = $@"INSERT INTO [WarrantAssistant].[dbo].[WarrantIssueAnnounce] ([TDate], [WName], [SerialNum], [Market]) VALUES ('{DateTime.Today.ToString("yyyyMMdd")}', '{ultraDataSource1.Rows[i]["WName"]}', '{ultraDataSource1.Rows[i]["SerialNumber"]}', '{ultraDataSource1.Rows[i]["Market"].ToString()}')";

                        MSSQL.ExecSqlQry(sql, GlobalVar.loginSet.warrantassistant45);
                    }
                    foreach (string file in Directory.GetFiles($"{path}\\{ultraDataSource1.Rows[i]["WName"]}")) {
                        string toFile = null;
                        string fileName = Path.GetFileName(file);
                        string fileExtension = Path.GetExtension(file).ToLower();
                        switch (type) {
                            case "Xml":
                                if (fileExtension == ".xml") {
                                    if (ultraDataSource1.Rows[i]["Market"].ToString() == "TWSE" && !fileName.Contains("OTC"))
                                        toFile = $"{path}\\Xml{now}\\{ultraDataSource1.Rows[i]["SerialNumber"]}.xml";
                                    else if (ultraDataSource1.Rows[i]["Market"].ToString() == "OTC" && fileName.Contains("OTC"))
                                        toFile = $"{path}\\XmlOTC{now}\\{ultraDataSource1.Rows[i]["SerialNumber"]}.xml";
                                }
                                break;
                            case "Renamed":
                                if (fileExtension != ".xml" && !fileName.Contains("銷售辦法公告")) {
                                    if (ultraDataSource1.Rows[i]["Market"].ToString() == "TWSE" && !fileName.Contains("OTC"))
                                        toFile = $"{path}\\Renamed{now}\\{ultraDataSource1.Rows[i]["SerialNumber"]}-{fileName}";
                                    else if (ultraDataSource1.Rows[i]["Market"].ToString() == "OTC" && fileName.Contains("OTC"))
                                        toFile = $"{path}\\RenamedOTC{now}\\{ultraDataSource1.Rows[i]["SerialNumber"]}-{fileName}";
                                }
                                break;
                                /*
                            case "RenamedB":
                                if (fileExtension != ".xml" && !fileName.Contains("銷售辦法公告") && !fileName.StartsWith("02") && !fileName.StartsWith("08")
                                && !fileName.StartsWith("14") && !fileName.StartsWith("15") && !fileName.StartsWith("16")
                                && !fileName.StartsWith("19") && !fileName.StartsWith("20") && !fileName.StartsWith("21")) {
                                    if (ultraDataSource1.Rows[i]["Market"].ToString() == "TWSE" && !fileName.Contains("OTC"))
                                        toFile = $"{path}\\RenamedB{now}\\{ultraDataSource1.Rows[i]["SerialNumber"]}-{Path.GetFileName(file)}";
                                    else if (ultraDataSource1.Rows[i]["Market"].ToString() == "OTC" && fileName.Contains("OTC"))
                                        toFile = $"{path}\\RenamedBOTC{now}\\{ultraDataSource1.Rows[i]["SerialNumber"]}-{Path.GetFileName(file)}";
                                }
                                break;
                                */
                                //2023/11/1 上線測試
                            case "RenamedB":
                                if (fileExtension != ".xml" && !fileName.Contains("銷售辦法公告") && !fileName.StartsWith("02") && !fileName.StartsWith("04")
                                && !fileName.StartsWith("07") && !fileName.StartsWith("08") && !fileName.StartsWith("09")
                                && !fileName.StartsWith("12") && !fileName.StartsWith("13") && !fileName.StartsWith("14"))
                                {
                                    if (ultraDataSource1.Rows[i]["Market"].ToString() == "TWSE" && !fileName.Contains("OTC"))
                                        toFile = $"{path}\\RenamedB{now}\\{ultraDataSource1.Rows[i]["SerialNumber"]}-{Path.GetFileName(file)}";
                                    else if (ultraDataSource1.Rows[i]["Market"].ToString() == "OTC" && fileName.Contains("OTC"))
                                        toFile = $"{path}\\RenamedBOTC{now}\\{ultraDataSource1.Rows[i]["SerialNumber"]}-{Path.GetFileName(file)}";
                                }
                                break;
                            case "SalesAnnounce":
                                if (fileName.Contains("銷售辦法公告"))
                                {
                                    if (ultraDataSource1.Rows[i]["Market"].ToString() == "TWSE" && !fileName.Contains("OTC"))
                                    {
                                        int index = fileName.IndexOf('_');
                                        toFile = $"{path}\\SalesAnnounce{now}\\{ultraDataSource1.Rows[i]["SerialNumber"]}-1-{fileName.Substring(index+1,fileName.Length-index-1)}";
                                    }
                                    else if (ultraDataSource1.Rows[i]["Market"].ToString() == "OTC" && fileName.Contains("OTC"))
                                    {
                                        int index = fileName.IndexOf('_');
                                        //toFile = $"{path}\\SalesAnnounceOTC{now}\\{ultraDataSource1.Rows[i]["SerialNumber"]}-1-{fileName.Substring(index + 1, fileName.Length - index - 1)}";
                                        //2021/5/14 OTC開放整批上傳，有改檔名
                                        toFile = $"{path}\\SalesAnnounceOTC{now}\\O-{ultraDataSource1.Rows[i]["SerialNumber"]}-1-{fileName.Substring(index + 1, fileName.Length - index - 1)}";
                                    }
                                }
                                break;
                        }
                        if (toFile != null)
                            File.Copy(file, toFile, true);
                    }
                }
            
            if (!Directory.EnumerateFileSystemEntries($"{path}\\{type}{now}").Any())
                Directory.Delete($"{path}\\{type}{now}");
            if (!Directory.EnumerateFileSystemEntries($"{path}\\{type}OTC{now}").Any())
                Directory.Delete($"{path}\\{type}OTC{now}");
            
            switch (type) {
                case "Xml":
                    toolStripTextBox1.Text = DateTime.Now + "修改XML檔名完成";
                    break;
                case "Renamed":
                    toolStripTextBox1.Text = DateTime.Now + "修改發行檔名A完成";
                    break;
                case "RenamedB":
                    toolStripTextBox1.Text = DateTime.Now + "修改發行檔名B完成";
                    break;
                case "SalesAnnounce":
                    toolStripTextBox1.Text = DateTime.Now + "修改銷售公告完成";
                    break;
            }
        }

        //修改銷售公告
        private void SalesAnnounce_Click(object sender, EventArgs e)
        {
            Rename("SalesAnnounce");
        }

        //修改上市公告
        private void ListedAnnounce_Click(object sender, EventArgs e)
        {
            Rename("ListedAnnounce");
        }

    }
}

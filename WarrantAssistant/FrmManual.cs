using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace WarrantAssistant
{
    public partial class FrmManual : Form
    {
        public FrmManual()
        {
            InitializeComponent();
            this.pic1.Select();
        }
        private void Paint(object sender, System.Windows.Forms.PaintEventArgs e)
        {
            Graphics g = e.Graphics;
            Font fnt_title = new Font("微軟正黑體", 18);
            Font fnt = new Font("微軟正黑體", 12);
            //Font fnt = new Font("Arial", 14);
            g.DrawString(@"基本資料",
            fnt_title, System.Drawing.Brushes.Black, new Point(10, 10));
            g.DrawString(@"發行及增額總表",
            fnt_title, System.Drawing.Brushes.Black, new Point(10, 740));
            g.DrawString(@"交易員",
           fnt_title, System.Drawing.Brushes.Black, new Point(10, 1020));
            g.DrawString(@"行政",
           fnt_title, System.Drawing.Brushes.Black, new Point(10, 1470));
            g.DrawString(@"
*標的Summary 
公開資訊觀測站上公布的本季可發標的，於每日上午8:15更新每日額度資料(已發額度、可發額度等)

*特殊標的	
每月月底更新高風險標的，包含量化指標風險級數、生技文創類股、以及資本額低於10億的KY公司

*標的發行檢查
每日更新標的日個股事件(法說會、股東會、股利、注意次數等)

*Put發行檢查
每日更新是否可發Put，以及控管機制計算指標(自家Call DeltaOne、自家Put DeltaOne、全市場Call DeltaOne、
全市場Put DeltaOne、自家Call/Put DeltaOne比例、自家/全市場 Put DeltaOne 比例、是否為特殊標底等)
DeltaOne計算公式：發行張數*行使比例

*已發行權證
登錄自家已發權證

*可增額列表
登錄可增額發行的權證列表

*建議Vol
提供交易員發行權證時的建議Vol
基本精神為確保一定的Vol spread和不要在市場太低水位，且不要賣比元大和市場PR60IV貴
公式為 Max(Max(HV *1.3 , 市場PR20 IV),Min(市場PR60 IV , 元大PR -10 IV))

*到期釋出額度
計算各標的近五日內到期釋出額度及預估當日強制註銷額度(不一定會強制註銷，只要符合強制註銷條件就會
納入計算，每日強制註銷額度依當日收到的通知信件為主)。
強制註銷條件為:
標準1：7-1百分比達20%(含)以上，且距到期日2個月內且流通在外數量低於發行數量5%，須強制註銷至發行數量之20%。
標準2：7-1百分比達20%(含)以上，且距到期日2個月內且流通在外數量低於發行數量10%，須強制註銷至發行數量之30%。
",
            fnt, System.Drawing.Brushes.Black, new Point(10,40));
            g.DrawString(@"
*搶額度總表(包含新發行和增額發行)
於每日上午10:40後顯示當日申請發行權證(包含新發及增額)是否要搶額度及搶發順位，
並提供後台更改權證名稱、行使比例、張數、以及發行張數權限。如果是需要搶額度的權證，
需強制發行10000張，發行張數改變的權證會在搶額度總表及發行總表上留下紀錄。

*發行總表
顯示當日所有申請新發行的權證，並提供後台修改權證發行條件的權限

*增額總表
顯示當日所有增額發行的權證

",
            fnt, System.Drawing.Brushes.Black, new Point(10, 770));
            g.DrawString(@"
*已發權證條件發行
顯示所有自家已發權證，並依照交易員分類，可以在這張表新增想發行的標的。
右鍵點擊[加到申請編輯表]後該標的就會出現在發行條件輸入的頁面中，接著更改發行條件即可。


*發行條件輸入
交易員在此輸入當日欲新發行權證，想發行的權證請點擊[確認]。權證發行作業須在上午9:50分前確認發行送出，
以便後台作業。點擊確認發行後，小幫手會審核權證條件，發行條件皆無誤會顯示成功申請。若發行條件有誤，
代表發行未成功，請修改權證條件或取消發行該檔權證，修改後再次點擊確認發行即可。上午10:40後搶額度總表
和發行總表會顯示申請權證的額度及順位，交易員可以修改權證條件以符合額度。
交易員點下確認送出後，小幫手會記錄每檔權證發行資料，每當有權證條件被修改，
再次點下確認送出後，在發行總表會顯示該檔權證的修改紀錄。

*增額條件輸入
交易員在此輸入當日欲增額發行權證，想發行的權證請點擊[確認]。權證發行作業須在上午9:50分前確認發行送出，
以便後台作業。點擊確認發行後，小幫手會審核權證條件，發行條件皆無誤會顯示成功申請。若發行條件有誤，
代表發行未成功，請修改權證條件或取消發行該檔權證，修改後再次點擊確認發行即可。上午10:40後搶額度總表
和發行總表會顯示申請權證的額度及順位，交易員可以修改權證條件以符合額度。
",
            fnt, System.Drawing.Brushes.Black, new Point(10, 1050));

            g.DrawString(@"
*可增額列表
登錄符合可增額發行權證

*7-1試算表
點擊[取得資料及編輯]能從交易所取得申請權證的額度，若原始申報時間不為空，代表該檔權證要搶額度，
申報時間會顯示交易所釋出結果的最終時間，此時申報時間的秒數即為該檔權證發行的順位(EX : 10:35:04:222 順位為04)，
若前面順位發完後有剩餘額度，該檔權證才有額度。若原始申報時間為空，則申報時間則為原始申報時間，
該檔權證不用搶額度，不受額度限制。由於每家發行商的順位不同，因此要搶發權證的額度會不斷變化，
7-1試算表也要定期點擊[取得資料及編輯]來更新。點擊[確認上傳]後，會更新7-1表的結果以及搶額度總表的額度。

*修正權證名稱
早上九點前權證基本資料可能更新不完全，會影響新發行權證名稱的序號。點擊[修正權證名稱]能更新當日新發權證名稱的序號

*轉申請發行TXT
將申請權證(包含新發與增額)的資料轉成txt檔，此txt檔用來上傳至交易所申報窗口。

*發行上傳檔
將申請新發行權證的資料轉成csv檔，此檔用來上傳至權證系統。

*增額上傳檔
將申請增額權證的資料轉成csv檔，此檔用來上傳至權證系統。

*關係人列表
產出關係人的csv檔

*修改檔案名稱
用於整批上傳至申報窗口時，能夠修改權證資料
",
           fnt, System.Drawing.Brushes.Black, new Point(10, 1500));



            //g.DrawString("QQ");
        }
    }
}

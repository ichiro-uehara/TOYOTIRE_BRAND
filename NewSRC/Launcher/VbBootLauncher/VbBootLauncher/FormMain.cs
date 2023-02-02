using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using DDEServer;
using SpecInfo;

namespace VbBootLauncher
{
    /// <summary>
    /// メニュータイプ
    /// </summary>
    public enum MenuType
    {
        None,
        PrimitiveSearch,
        EditingSearch,
        PrimitiveRegist,
        EditingRegist,
    }

    /// <summary>
    /// 登録モード
    /// </summary>
    public enum RegistMode
    {
        None,
        FRESH,
        MODIFY,
    }

    public delegate void AddListDelegate(string data);

    public partial class FormMain : Form
    {
        #region 定数

        private const string RecieveLabel = "<受信> ";
        private const string SendLabel = "<送信> ";

        private List<string> GmCodeList = new List<string> { "KO9001A0AT", "KO9002A0AO", "KO9003A0AY", "KO9002A0AO" };

        private List<string> GmCodeList2 = new List<string> { "KO9001A0AT", "KO9002A0AO", "KO9003A0AY", "KO9002A0AO", "KO9001A0AT" };

        #endregion

        /// <summary>
        /// DDE通信クラス
        /// </summary>
        private DDECom DdeCom = new DDECom();

        /// <summary>
        /// 刻印文字スペッククラス
        /// </summary>
        SpecPrimitiveChar SpecPrimitive = new SpecPrimitiveChar();

        /// <summary>
        /// 編集文字スペッククラス
        /// </summary>
        SpecEditingChar SpecEditing = new SpecEditingChar();

        private MenuType SelectMenu = MenuType.None;

        /// <summary>
        /// 登録モード
        /// </summary>
        private RegistMode RegistMode = RegistMode.None;

        /// <summary>
        /// SPECADDリクエスト応答
        /// </summary>
        private ReqResponsType SpecAddResp = ReqResponsType.SpecAddOK;

        public FormMain()
        {
            InitializeComponent();
        }

        /// <summary>
        /// リクエスト受信イベント
        /// </summary>
        /// <param name="item">リクエストアイテム</param>
        public void OnRequest(RequestItem item)
        {
            string mes = RecieveLabel;
            mes += "[リクエスト]  ";
            mes += item.ToString();
            object[] objArray = new object[1];
            objArray[0] = mes;
            listBoxCom.BeginInvoke(new AddListDelegate(AddListEvent), objArray);

            switch (SelectMenu)
            {
                case MenuType.PrimitiveSearch:
                    OnRequestPrimitiveSearch(item);
                    break;
                case MenuType.EditingSearch:
                    OnRequestEditingSearch(item);
                    break;
                case MenuType.PrimitiveRegist:
                    OnRequestPrimitiveRegist(item);
                    break;
                case MenuType.EditingRegist:
                    OnRequestEditingRegist(item);
                    break;
            }
        }

        /// <summary>
        /// Poke受信イベント
        /// </summary>
        /// <param name="item">ポークアイテム</param>
        /// <param name="data">データ</param>
        public void OnPoke(PokeItem item, string data)
        {
            string mes = RecieveLabel;
            mes = mes + "[ポーク]      " + "アイテム：" + item.ToString() + "   データ：" + data;
            listBoxCom.BeginInvoke((MethodInvoker)delegate () { listBoxCom.Items.Add(mes); });

            switch (SelectMenu)
            {
                case MenuType.PrimitiveSearch:
                    OnPokePrimitiveSearch(item, data);
                    break;
                case MenuType.EditingSearch:
                    OnPokeEditingSearch(item, data);
                    break;
                case MenuType.PrimitiveRegist:
                    OnPokePrimitiveRegist(item, data);
                    break;
                case MenuType.EditingRegist:
                    OnPokeEditingRegist(item, data);
                    break;
            }
        }

        /// <summary>
        /// クライアント通信切断イベント
        /// </summary>
        public void OnDisconnected()
        {
            object[] objArray = new object[1];
            objArray[0] = "[クライアント切断通知]";
            listBoxCom.BeginInvoke(new AddListDelegate(AddListEvent), objArray);

            DdeCom.TerminateServer();

            listBoxCom.BeginInvoke((MethodInvoker)delegate () { listBoxCom.Items.Add("[DDEサーバー終了]"); });
        }

        #region リクエスト

        /// <summary>
        /// リクエスト受信イベント(刻印文字検索用)
        /// </summary>
        /// <param name="item">リクエストアイテム</param>
        private void OnRequestPrimitiveSearch(RequestItem item)
        {
            string mes = SendLabel;

            if (item == RequestItem.WINNAME)
            {
                DdeCom.WinNameRequestRespons(WinType.GMSEARCH);

                mes += "[応答]        GMSEARCH";
                listBoxCom.BeginInvoke((MethodInvoker)delegate () { listBoxCom.Items.Add(mes); });
            }
            else if (item == RequestItem.PICEMPTY)
            {
                DdeCom.PicEmptyRequestRespons(99);

                mes += "[応答]        PICEMPTY099";
                listBoxCom.BeginInvoke((MethodInvoker)delegate () { listBoxCom.Items.Add(mes); });
            }
            else if (item == RequestItem.ACADREAD)
            {
                DdeCom.RequestRespons(ReqResponsType.CadReadOK);

                mes += "[応答]        OK-DATA";
                listBoxCom.BeginInvoke((MethodInvoker)delegate () { listBoxCom.Items.Add(mes); });
            }
        }

        /// <summary>
        /// リクエスト受信イベント(編集文字検索用)
        /// </summary>
        /// <param name="item">リクエストアイテム</param>
        private void OnRequestEditingSearch(RequestItem item)
        {
            string mes = SendLabel;

            if (item == RequestItem.WINNAME)
            {
                DdeCom.WinNameRequestRespons(WinType.HMSEARCH1);

                mes += "[応答]        HMSEARCH1";
                listBoxCom.BeginInvoke((MethodInvoker)delegate () { listBoxCom.Items.Add(mes); });
            }
            else if (item == RequestItem.PICEMPTY)
            {
                DdeCom.PicEmptyRequestRespons(99);

                mes += "[応答]        PICEMPTY099";
                listBoxCom.BeginInvoke((MethodInvoker)delegate () { listBoxCom.Items.Add(mes); });
            }
            else if (item == RequestItem.ACADREAD)
            {
                DdeCom.RequestRespons(ReqResponsType.CadReadOK);

                mes += "[応答]        OK-DATA";
                listBoxCom.BeginInvoke((MethodInvoker)delegate () { listBoxCom.Items.Add(mes); });
            }
        }

        /// <summary>
        /// リクエスト受信イベント(刻印文字登録用)
        /// </summary>
        /// <param name="item">リクエストアイテム</param>
        private void OnRequestPrimitiveRegist(RequestItem item)
        {
            string mes = SendLabel;

            if (item == RequestItem.WINNAME)
            {
                DdeCom.WinNameRequestRespons(WinType.GMSAVE);

                mes += "[応答]        GMSAVE";
            }
            else if (item == RequestItem.SAVEMODE)
            {
                DdeCom.DummyRequestRespons();

                mes += @"[応答]        「改行（￥ｒ￥ｎ）」";
            }
            else if (item == RequestItem.SPECADD)
            {
                DdeCom.RequestRespons(SpecAddResp);

                mes += "[応答]        " + SpecAddResp.ToString();
            }
            else if (item == RequestItem.ACADSAVE)
            {
                DdeCom.RequestRespons(ReqResponsType.CadSaveOK);

                mes += "[応答]        ZUMEN SAVE OK !!";
            }
            else if (item == RequestItem.SPECDATA)
            {
                // フォント名、フォント区分１・２、文字名１・２
                if (RegistMode == RegistMode.MODIFY)
                {
                    SpecPrimitive.font_name = "KO9901";
                    SpecPrimitive.font_class1 = "A";
                    SpecPrimitive.font_class2 = "0";
                    SpecPrimitive.name1 = "A";
                    SpecPrimitive.name2 = "T";
                }

                // 高さ
                SpecPrimitive.high = 20.14;
                if (RegistMode == RegistMode.MODIFY)
                {
                    SpecPrimitive.high = 21.5;
                }

                // 幅
                SpecPrimitive.width = 22.0;
                if (RegistMode == RegistMode.MODIFY)
                {
                    SpecPrimitive.width = 23.5;
                }

                // 角度
                SpecPrimitive.ang = 0.0;
                if (RegistMode == RegistMode.MODIFY)
                {
                    SpecPrimitive.ang = 4.0;
                }

                // 実高さ
                SpecPrimitive.moji_high = 20.0;
                if (RegistMode == RegistMode.MODIFY)
                {
                    SpecPrimitive.moji_high = 21.5;
                }

                // ずれ量
                SpecPrimitive.moji_shift = 0.0;

                // 水平原点位置

                // 垂直原点位置

                // 文字原点Ｘ
                SpecPrimitive.org_x = 11.21;
                if (RegistMode == RegistMode.MODIFY)
                {
                    SpecPrimitive.org_x = 12.35;
                }

                // 文字原点Ｙ
                SpecPrimitive.org_y = 0.0;

                // 枠左下Ｘ
                SpecPrimitive.left_bottom_x = 0.0;

                // 枠左下Ｙ
                SpecPrimitive.left_bottom_y = 0.0;

                // 枠右下Ｘ
                SpecPrimitive.right_bottom_x = 22.43;
                if (RegistMode == RegistMode.MODIFY)
                {
                    SpecPrimitive.right_bottom_x = 23.15;
                }

                // 枠右下Ｙ
                SpecPrimitive.right_bottom_y = 0.0;

                // 枠右上Ｘ
                SpecPrimitive.right_top_x = 22.43;
                if (RegistMode == RegistMode.MODIFY)
                {
                    SpecPrimitive.right_top_x = 23.15;
                }

                // 枠右上Ｙ
                SpecPrimitive.right_top_y = 20.81;
                if (RegistMode == RegistMode.MODIFY)
                {
                    SpecPrimitive.right_top_y = 21.45;
                }

                // 枠左上Ｘ
                SpecPrimitive.left_top_x = 0.0;

                // 枠左上Ｙ
                SpecPrimitive.left_top_y = 20.81;
                if (RegistMode == RegistMode.MODIFY)
                {
                    SpecPrimitive.left_top_y = 21.45;
                }

                // 縁取り幅
                SpecPrimitive.hem_width = 0.0;

                // ハッチング角度
                SpecPrimitive.hatch_ang = 0.0;

                // ハッチング幅
                SpecPrimitive.hatch_width = 0.0;

                // ハッチング間隔
                SpecPrimitive.hatch_space = 0.0;

                // ハッチング始点Ｘ
                SpecPrimitive.hatch_x = 0.0;

                // ハッチング始点Ｙ
                SpecPrimitive.hatch_y = 0.0;

                // 基準Ｒ
                SpecPrimitive.base_r = 0.0;

                // 旧フォント名

                // 旧フォント区分

                // 旧文字名

                // 配置ＰＩＣ
                SpecPrimitive.haiti_pic = 11;
                if (RegistMode == RegistMode.MODIFY)
                {
                    SpecPrimitive.haiti_pic = 12;
                }

                // DDE通信用の特性101データ取得
                string resp = "";

                if (RegistMode == RegistMode.MODIFY)
                {
                    resp = SpecPrimitive.GetSpec101ForDDE(false);
                }
                else
                {
                    resp = SpecPrimitive.GetSpec101ForDDE();
                }

                DdeCom.RequestRespons(resp + "\n");

                mes += ("[応答]        " + resp);
            }

            listBoxCom.BeginInvoke((MethodInvoker)delegate () { listBoxCom.Items.Add(mes); });
        }

        /// <summary>
        /// リクエスト受信イベント(編集文字登録用)
        /// </summary>
        /// <param name="item">リクエストアイテム</param>
        private void OnRequestEditingRegist(RequestItem item)
        {
            string mes = SendLabel;

            if (item == RequestItem.WINNAME)
            {
                DdeCom.WinNameRequestRespons(WinType.HMSAVE);

                mes += "[応答]        HMSAVE";
            }
            else if (item == RequestItem.SAVEMODE)
            {
                DdeCom.DummyRequestRespons();

                mes += @"[応答]        「改行（￥ｒ￥ｎ）」";
            }
            else if (item == RequestItem.SPEC2011)
            {
                string resp = SpecEditing.GetSpec2011ForDDE();
                DdeCom.RequestRespons(resp + "\n");

                mes += ("[応答]        " + resp);
            }
            else if (item == RequestItem.SPECADD)
            {
                DdeCom.RequestRespons(SpecAddResp);

                mes += "[応答]        " + SpecAddResp.ToString();
            }
            else if (item == RequestItem.ACADSAVE)
            {
                DdeCom.RequestRespons(ReqResponsType.CadSaveOK);

                mes += "[応答]        ZUMEN SAVE OK !!";
            }
            else if (item == RequestItem.SPECDATA)
            {
                // ID

                // フォント名
                if (RegistMode == RegistMode.MODIFY)
                {
                    SpecEditing.font_name = "HE9901";
                }

                // no
                if (RegistMode == RegistMode.MODIFY)
                {
                    SpecEditing.no = "00";
                }

                // Spell
                if (RegistMode == RegistMode.MODIFY)
                {
                    SpecEditing.SetSpell("TOYOT");
                }
                else
                {
                    SpecEditing.SetSpell("TOYO");
                }

                // 配置指定
                SpecEditing.haiti_sitei = 1;

                // 刻印文字数
                SpecEditing.gm_num = 4;
                if (RegistMode == RegistMode.MODIFY)
                {
                    SpecEditing.gm_num = 5;
                }

                // 幅
                SpecEditing.width = 90.0;
                if (RegistMode == RegistMode.MODIFY)
                {
                    SpecEditing.width = 91.5;
                }

                // 高さ
                SpecEditing.high = 20.0;
                if (RegistMode == RegistMode.MODIFY)
                {
                    SpecEditing.high = 21.5;
                }

                // 角度
                SpecEditing.ang = 0.0;
                if (RegistMode == RegistMode.MODIFY)
                {
                    SpecEditing.ang = 4.0;
                }

                // 配置ＰＩＣ
                SpecEditing.haiti_pic = 11;
                if (RegistMode == RegistMode.MODIFY)
                {
                    SpecEditing.haiti_pic = 12;
                    SpecEditing.AddGmList(GmCodeList2);
                }
                else
                {
                    SpecEditing.AddGmList(GmCodeList);
                }

                // DDE通信用の特性201データ取得
                string resp = "";

                if (RegistMode == RegistMode.MODIFY)
                {
                    resp = SpecEditing.GetSpec201ForDDE(false);
                }
                else
                {
                    resp = SpecEditing.GetSpec201ForDDE();
                }

                DdeCom.RequestRespons(resp + "\n");

                mes += ("[応答]        " + resp);
            }

            listBoxCom.BeginInvoke((MethodInvoker)delegate () { listBoxCom.Items.Add(mes); });
        }

        #endregion

        #region Poke

        /// <summary>
        /// Poke受信イベント(刻印文字検索用)
        /// </summary>
        /// <param name="item">ポークアイテム</param>
        /// <param name="data">データ</param>
        public void OnPokePrimitiveSearch(PokeItem item, string data)
        {
        }

        /// <summary>
        /// Poke受信イベント(編集文字検索用)
        /// </summary>
        /// <param name="item">ポークアイテム</param>
        /// <param name="data">データ</param>
        public void OnPokeEditingSearch(PokeItem item, string data)
        {
        }

        /// <summary>
        /// Poke受信イベント(刻印文字登録用)
        /// </summary>
        /// <param name="item">ポークアイテム</param>
        /// <param name="data">データ</param>
        public void OnPokePrimitiveRegist(PokeItem item, string data)
        {
            if (item == PokeItem.SAVEMODE)
            {
                if (data.Equals("FRESH"))
                {
                    RegistMode = RegistMode.FRESH;
                }
                else if (data.Equals("MODIFY"))
                {
                    RegistMode = RegistMode.MODIFY;
                }
            }
            else if (item == PokeItem.SPECADD)
            {
                // 受信データより刻印文字スペックを作成
                SpecPrimitiveChar spec = new SpecPrimitiveChar();
                spec.SetSpecDataByDDE(data);

                if (SpecPrimitive == spec)
                {
                    SpecAddResp = ReqResponsType.SpecAddOK;
                }
                else
                {
                    SpecAddResp = ReqResponsType.Error;
                }
            }
        }

        /// <summary>
        /// Poke受信イベント(編集文字登録用)
        /// </summary>
        /// <param name="item">ポークアイテム</param>
        /// <param name="data">データ</param>
        public void OnPokeEditingRegist(PokeItem item, string data)
        {
            if (item == PokeItem.SAVEMODE)
            {
                if (data.Equals("FRESH"))
                {
                    RegistMode = RegistMode.FRESH;
                }
                else if (data.Equals("MODIFY"))
                {
                    RegistMode = RegistMode.MODIFY;
                }
            }
            else if (item == PokeItem.SPECADD)
            {
                // 受信データより編集文字スペックを作成
                SpecEditingChar spec = new SpecEditingChar();
                spec.SetSpecDataByDDE(data);

                if (SpecEditing == spec)
                {
                    SpecAddResp = ReqResponsType.SpecAddOK;
                }
                else
                {
                    SpecAddResp = ReqResponsType.Error;
                }
            }
        }

        #endregion

        /// <summary>
        /// DDEサーバー開始
        /// </summary>
        private void StartServer()
        {
            DdeCom.StartServer();
            DdeCom.RegistRequestEvent(this.OnRequest);
            DdeCom.RegistPokeEvent(this.OnPoke);
            DdeCom.RegistDisconnectedEvent(this.OnDisconnected);
        }

        /// <summary>
        /// リストアイテム追加イベント
        /// </summary>
        /// <param name="data">データ</param>
        private void AddListEvent(string data)
        {
            listBoxCom.Items.Add(data);
        }

        /// <summary>
        /// 「刻印文字検索」メニュー
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void MenuItemPrimitiveSearch_Click(object sender, EventArgs e)
        {
            SelectMenu = MenuType.PrimitiveSearch;
            StartServer();
            listBoxCom.Items.Clear();
            AddListEvent("[DDEサーバー開始]");

            System.Diagnostics.Process.Start(@"C:\ACAD19\exe\BrandVB1.exe");
        }

        /// <summary>
        /// 「編集文字検索」メニュー
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void MenuItemEditingSearch_Click(object sender, EventArgs e)
        {
            SelectMenu = MenuType.EditingSearch;
            StartServer();
            listBoxCom.Items.Clear();
            AddListEvent("[DDEサーバー開始]");

            System.Diagnostics.Process.Start(@"C:\ACAD19\exe\BrandVB1.exe");
        }

        /// <summary>
        /// 「刻印文字登録」メニュー
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void MenuItemPrimitiveRegist_Click(object sender, EventArgs e)
        {
            SelectMenu = MenuType.PrimitiveRegist;
            StartServer();
            listBoxCom.Items.Clear();
            AddListEvent("[DDEサーバー開始]");

            //System.Diagnostics.Process.Start(@"C:\ACAD19\exe\BrandVB1.exe");
        }

        /// <summary>
        /// 「編集文字登録」メニュー
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void MenuItemEditingRegist_Click(object sender, EventArgs e)
        {
            SelectMenu = MenuType.EditingRegist;
            StartServer();
            listBoxCom.Items.Clear();
            AddListEvent("[DDEサーバー開始]");

            System.Diagnostics.Process.Start(@"C:\ACAD19\exe\BrandVB1.exe");
        }
    }
}

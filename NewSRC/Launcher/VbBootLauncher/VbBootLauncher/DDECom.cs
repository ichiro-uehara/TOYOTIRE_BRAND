using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace DDEServer
{
    #region デリゲート

    /// <summary>
    /// リクエスト受信デリゲート
    /// </summary>
    /// <param name="item">リクエストアイテム</param>
    public delegate void OnRequestDelegate(RequestItem item);

    /// <summary>
    /// ポーク受信デリゲート
    /// </summary>
    /// <param name="item">ポークアイテム</param>
    /// <param name="data">データ</param>
    public delegate void OnPokeDelegate(PokeItem item, string data);

    #endregion

    /// <summary>
    /// リクエストアイテム
    /// </summary>
    public enum RequestItem
    {
        None,
        WINNAME,        // 画面名リクエスト
        PICEMPTY,       // 空ピクチャー数（残数）リクエスト
        ACADREAD,       // CADデータ(DWG)読込リクエスト
        SAVEMODE,       // 保存モード送信通知(新規／更新)
        SPECADD,        // 保存スペックデータ送信通知
        ACADSAVE,       // CADデータ(DWG)保存リクエスト
        SPECDATA,       // スペックデータリクエスト
        SPEC2011,       // スペック2011リクエスト
    }

    /// <summary>
    /// ポークアイテム
    /// </summary>
    public enum PokeItem
    {
        None,
        ACADREAD,       // CADデータ(DWG)読込ファイル名送信
        SAVEMODE,       // 保存モード送信(新規／更新)
        SPECADD,        // 保存スペックデータ送信
        ACADSAVE,       // CADデータ(DWG)保存ァイル名送信
        ERROR,          // エラー通知
    }

    /// <summary>
    /// リクエスト応答タイプ
    /// </summary>
    public enum ReqResponsType
    {
        CadReadOK,      // CADデータ(DWG)読込結果OK
        SpecAddOK,      // 保存スペックデータ結果OK
        CadSaveOK,      // CADデータ(DWG)保存結果OK
        Error,          // エラー
    }

    /// <summary>
    /// 画面タイプ
    /// </summary>
    public enum WinType
    {
        None,
        GMSEARCH,       // 刻印文字検索画面
        GMSAVE,         // 刻印文字登録画面
        HMSEARCH1,      // 編集文字検索画面1
        HMSAVE,         // 編集文字登録画面
    }

    /// <summary>
    /// DDE通信クラス
    /// </summary>
    public class DDECom
    {
        #region 定数

        #region リクエストアイテム

        /// <summary>
        /// 画面名リクエスト
        /// </summary>
        private const string RequestWinName = "WINNAME";

        /// <summary>
        /// 空ピクチャー数（残数）リクエスト
        /// </summary>
        private const string RequestPicEmpty = "PICEMPTY";

        /// <summary>
        /// CADデータ(DWG)読込リクエスト
        /// </summary>
        private const string RequestAcadRead = "ACADREAD";

        /// <summary>
        /// 保存モード通知(新規／更新)
        /// </summary>
        private const string RequestSaveMode = "SAVEMODE";

        /// <summary>
        /// 保存スペックデータ送信通知
        /// </summary>
        private const string RequestSpecAdd = "SPECADD";

        /// <summary>
        /// CADデータ(DWG)保存リクエスト
        /// </summary>
        private const string RequestAcadSave = "ACADSAVE";

        /// <summary>
        /// スペックデータリクエスト
        /// </summary>
        private const string RequestSpecData = "SPECDATA";

        /// <summary>
        /// スペック2011リクエスト
        /// </summary>
        private const string RequestSpec2011 = "SPEC2011";

        #endregion

        #region リクエスト応答

        /// <summary>
        /// CADデータ(DWG)読込結果OK
        /// </summary>
        private const string ResCadReadOK = "OK-DATA";

        /// <summary>
        /// 保存スペックデータ結果OK
        /// </summary>
        private const string ResSpecAddOK = "SPECADD OK";

        /// <summary>
        /// CADデータ(DWG)保存結果OK
        /// </summary>
        private const string ResCadSaveOK = "ZUMEN SAVE OK !!";

        /// <summary>
        /// エラー
        /// </summary>
        private const string ResError = "ERROR";

        #endregion

        #region ポークアイテム

        /// <summary>
        /// CADデータ(DWG)読込ファイル名送信
        /// </summary>
        private const string PokeAcadRead = "ACADREAD";

        /// <summary>
        /// 保存モード送信(新規／更新)
        /// </summary>
        private const string PokeSaveMode = "SAVEMODE";

        /// <summary>
        /// 保存スペックデータ送信
        /// </summary>
        private const string PokeSpecAdd = "SPECADD";

        /// <summary>
        /// CADデータ(DWG)保存ァイル名送信
        /// </summary>
        private const string PokeAcadSave = "ACADSAVE";

        /// <summary>
        /// エラー通知
        /// </summary>
        private const string PokeError = "ERROR";

        #endregion

        #endregion

        #region publicプロパティ

        /// <summary>
        /// DDEサービス名
        /// </summary>
        public string DDEServiceName { get; set; }  = "TOYO";

        #endregion

        #region privateプロパティ

        /// <summary>
        /// DDEサーバークラス
        /// </summary>
        private DDEServer Server = null;

        /// <summary>
        /// リクエスト受信イベント
        /// </summary>
        private OnRequestDelegate OnRequestEvent = null;

        /// <summary>
        /// Poke受信イベント
        /// </summary>
        private OnPokeDelegate OnPokeEvent = null;

        /// <summary>
        /// クライアント通信切断イベント
        /// </summary>
        private OnDisconnectedDelegate OnDisconnectedEvent = null;

        #endregion

        #region publicメソッド

        /// <summary>
        /// コンストラクタ
        /// </summary>
        public DDECom()
        {
        }

        /// <summary>
        /// DDEサーバーを開始する
        /// </summary>
        /// <returns>true:成功 false:失敗</returns>
        public bool StartServer()
        {
            try
            {
                Server = new DDEServer(DDEServiceName);
                Server.Register();
                Server.RegistRequestEvent(this.OnRequest);
                Server.RegistPokeEvent(this.OnPoke);
                Server.RegistDisconnectedEvent(this.OnDisconnected);
                return true;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message + "\n" + ex.StackTrace, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }
        }

        /// <summary>
        /// DDEサーバーを終了する
        /// </summary>
        public void TerminateServer()
        {
            Server.Terminate();
            Server.Dispose();

            OnRequestEvent = null;
            OnPokeEvent = null;
            OnDisconnectedEvent = null;
        }

        #region リクエスト応答

        /// <summary>
        /// リクエスト応答
        /// </summary>
        /// <param name="respons">応答データ</param>
        public void RequestRespons(string respons)
        {
            Server.RequestRespons(respons);
        }

        /// <summary>
        /// リクエスト応答
        /// </summary>
        /// <param name="type">リクエスト応答タイプ</param>
        public void RequestRespons(ReqResponsType type)
        {
            if (type == ReqResponsType.CadReadOK)
            {
                this.RequestRespons(ResCadReadOK + "\n");
            }
            else if (type == ReqResponsType.SpecAddOK)
            {
                this.RequestRespons(ResSpecAddOK + "\n");
            }
            else if (type == ReqResponsType.CadSaveOK)
            {
                this.RequestRespons(ResCadSaveOK + "\n");
            }
            else if (type == ReqResponsType.Error)
            {
                this.RequestRespons(ResError + "\n");
            }
        }

        /// <summary>
        /// 画面名リクエスト応答
        /// </summary>
        /// <param name="item">画面タイプ</param>
        public void WinNameRequestRespons(WinType type)
        {
            this.RequestRespons(type.ToString() + "\n");
        }

        /// <summary>
        /// 空ピクチャー数（残数）リクエスト応答
        /// </summary>
        /// <param name="num">空ピクチャー数</param>
        public void PicEmptyRequestRespons(int num)
        {
            this.RequestRespons("PICEMPTY" + num.ToString("000") + "\n");
        }

        /// <summary>
        /// リクエスト(内容なし)応答
        /// </summary>
        public void DummyRequestRespons()
        {
            this.RequestRespons("\r\n\n");
        }

        #endregion

        #region イベント登録

        /// <summary>
        /// リクエスト受信イベント登録
        /// </summary>
        /// <param name="eventMethod">リクエスト受信イベントメソッド</param>
        /// <remarks>登録した受信イベントメソッドに通知される</remarks>
        public void RegistRequestEvent(OnRequestDelegate eventMethod)
        {
            OnRequestEvent = new OnRequestDelegate(eventMethod);
        }

        /// <summary>
        /// Poke受信イベント登録
        /// </summary>
        /// <param name="eventMethod">Poke受信イベントメソッド</param>
        /// <remarks>登録した受信イベントメソッドに通知される</remarks>
        public void RegistPokeEvent(OnPokeDelegate eventMethod)
        {
            OnPokeEvent = new OnPokeDelegate(eventMethod);
        }

        /// <summary>
        /// クライアント通信切断イベント登録
        /// </summary>
        /// <param name="eventMethod"></param>
        public void RegistDisconnectedEvent(OnDisconnectedDelegate eventMethod)
        {
            OnDisconnectedEvent = new OnDisconnectedDelegate(eventMethod);
        }

        #endregion

        #endregion

        #region privateメソッド

        /// <summary>
        /// リクエスト受信イベント
        /// </summary>
        /// <param name="item">アイテム</param>
        private void OnRequest(string item)
        {
            if (OnRequestEvent != null)
            {
                RequestItem reqItem = RequestItem.None;

                if (item.Equals(RequestWinName))
                {
                    reqItem = RequestItem.WINNAME;
                }
                else if (item.Equals(RequestPicEmpty))
                {
                    reqItem = RequestItem.PICEMPTY;
                }
                else if (item.Equals(RequestAcadRead))
                {
                    reqItem = RequestItem.ACADREAD;
                }
                else if (item.Equals(RequestSaveMode))
                {
                    reqItem = RequestItem.SAVEMODE;
                }
                else if (item.Equals(RequestSpecAdd))
                {
                    reqItem = RequestItem.SPECADD;
                }
                else if (item.Equals(RequestAcadSave))
                {
                    reqItem = RequestItem.ACADSAVE;
                }
                else if (item.Equals(RequestSpecData))
                {
                    reqItem = RequestItem.SPECDATA;
                }
                else if (item.Equals(RequestSpec2011))
                {
                    reqItem = RequestItem.SPEC2011;
                }

                // リクエスト受信イベント発行
                OnRequestEvent(reqItem);
            }
        }

        /// <summary>
        /// Poke受信イベント
        /// </summary>
        /// <param name="item">アイテム</param>
        /// <param name="data">データ</param>
        private void OnPoke(string item, string data)
        {
            if (OnRequestEvent != null)
            {
                PokeItem reqItem = PokeItem.None;

                if (item.Equals(PokeAcadRead))
                {
                    reqItem = PokeItem.ACADREAD;
                }
                else if (item.Equals(PokeSaveMode))
                {
                    reqItem = PokeItem.SAVEMODE;
                }
                else if (item.Equals(PokeSpecAdd))
                {
                    reqItem = PokeItem.SPECADD;
                }
                else if (item.Equals(PokeAcadSave))
                {
                    reqItem = PokeItem.ACADSAVE;
                }
                else if (item.Equals(PokeError))
                {
                    reqItem = PokeItem.ERROR;
                }

                // Poke受信イベント発行
                OnPokeEvent(reqItem, data);
            }
        }

        /// <summary>
        /// クライアント通信切断イベント
        /// </summary>
        private void OnDisconnected()
        {
            if (OnDisconnectedEvent != null)
            {
                // クライアント通信切断イベント発行
                OnDisconnectedEvent();
            }
        }

        #endregion
    }
}

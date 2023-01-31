using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using NDde.Server;

namespace DDEServer
{
    #region デリゲート

    /// <summary>
    /// クライアントからのリクエスト受信デリゲート
    /// </summary>
    /// <param name="item">アイテム</param>
    public delegate void OnClientRequestDelegate(string item);

    /// <summary>
    /// クライアントからのポーク受信デリゲート
    /// </summary>
    /// <param name="item">アイテム</param>
    /// <param name="data">データ</param>
    public delegate void OnClientPokeDelegate(string item, string data);

    /// <summary>
    /// クライアント通信切断デリゲート
    /// </summary>
    public delegate void OnDisconnectedDelegate();

    #endregion

    /// <summary>
    /// DDEサーバークラス
    /// </summary>
    public class DDEServer : DdeServer
    {
        #region privateプロパティ

        /// <summary>
        /// DDE Client Conversation
        /// </summary>
        private DdeConversation Conversation { get; set; } = null;

        /// <summary>
        /// リクエスト応答構造体
        /// </summary>
        private RequestResult _RequestResult;

        /// <summary>
        /// リクエスト受信イベント
        /// </summary>
        private OnClientRequestDelegate OnRequestEvent = null;

        /// <summary>
        /// Poke受信イベント
        /// </summary>
        private OnClientPokeDelegate OnPokeEvent = null;

        /// <summary>
        /// クライアント通信切断イベント
        /// </summary>
        private OnDisconnectedDelegate OnDisconnectedEvent = null;

        #endregion

        #region コンストラクタ

        /// <summary>
        /// コンストラクタ
        /// </summary>
        /// <param name="service">サービス名</param>
        public DDEServer(string service) : base(service)
        {
        }

        #endregion

        #region publicメソッド

        /// <summary>
        /// DDEサーバー終了
        /// </summary>
        public void Terminate()
        {
            base.Dispose();

            OnRequestEvent = null;
            OnPokeEvent = null;
            OnDisconnectedEvent = null;
        }

        /// <summary>
        /// リクエスト受信イベント登録
        /// </summary>
        /// <param name="eventMethod">リクエスト受信イベントメソッド</param>
        /// <remarks>登録した受信イベントメソッドに通知される</remarks>
        public void RegistRequestEvent(OnClientRequestDelegate eventMethod)
        {
            OnRequestEvent = new OnClientRequestDelegate(eventMethod);
        }

        /// <summary>
        /// Poke受信イベント登録
        /// </summary>
        /// <param name="eventMethod">Poke受信イベントメソッド</param>
        /// <remarks>登録した受信イベントメソッドに通知される</remarks>
        public void RegistPokeEvent(OnClientPokeDelegate eventMethod)
        {
            OnPokeEvent = new OnClientPokeDelegate(eventMethod);
        }

        /// <summary>
        /// クライアント通信切断イベント登録
        /// </summary>
        /// <param name="eventMethod"></param>
        public void RegistDisconnectedEvent(OnDisconnectedDelegate eventMethod)
        {
            OnDisconnectedEvent = new OnDisconnectedDelegate(eventMethod);
        }

        /// <summary>
        /// クライアントとの通信切断
        /// </summary>
        public void DisconnectClient()
        {
            try
            {
                if (this.Conversation != null)
                {
                    base.Disconnect(this.Conversation);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message + "\n" + ex.StackTrace, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// リクエスト応答
        /// </summary>
        /// <param name="respons">応答データ</param>
        public void RequestRespons(string respons)
        {
            _RequestResult = new RequestResult(System.Text.Encoding.ASCII.GetBytes(respons));
        }

        #endregion

        #region overrideメソッド

        /// <summary>
        /// レジスター
        /// </summary>
        public override void Register()
        {
            try
            {
                base.Register();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message + "\n" + ex.StackTrace, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// クライアントからのリクエスト受信
        /// </summary>
        /// <param name="conversation">Client Conversation</param>
        /// <param name="item">アイテム</param>
        /// <param name="format">フォーマット</param>
        /// <returns></returns>
        protected override RequestResult OnRequest(DdeConversation conversation, string item, int format)
        {
            if (OnRequestEvent != null)
            {
                // リクエスト受信イベント発行
                OnRequestEvent(item);
            }

            return _RequestResult;
        }

        public override void Unregister()
        {
            base.Unregister();
        }

        protected override bool OnBeforeConnect(string topic)
        {
            return true;
        }

        /// <summary>
        /// クライアント接続後イベント
        /// </summary>
        /// <param name="conversation"></param>
        protected override void OnAfterConnect(DdeConversation conversation)
        {
            this.Conversation = conversation;
        }

        /// <summary>
        /// クライアント通信切断イベント
        /// </summary>
        /// <param name="conversation">Client Conversation</param>
        protected override void OnDisconnect(DdeConversation conversation)
        {
            this.DisconnectClient();
            if (OnDisconnectedEvent != null)
            {
                // クライアント通信切断イベント発行
                OnDisconnectedEvent();
            }
        }

        protected override bool OnStartAdvise(DdeConversation conversation, string item, int format)
        {
            return format == 1;
        }

        protected override void OnStopAdvise(DdeConversation conversation, string item)
        {
        }

        protected override ExecuteResult OnExecute(DdeConversation conversation, string command)
        {
            return ExecuteResult.Processed;
        }

        /// <summary>
        /// クライアントからのPoke受信イベント
        /// </summary>
        /// <param name="conversation">Client Conversation</param>
        /// <param name="item">アイテム</param>
        /// <param name="data">データ</param>
        /// <param name="format">フォーマット</param>
        /// <returns></returns>
        protected override PokeResult OnPoke(DdeConversation conversation, string item, byte[] data, int format)
        {
            try
            {
                if (OnPokeEvent != null && data.Length > 0)
                {
                    string strData = System.Text.Encoding.ASCII.GetString(data);
                    if (data[data.Length - 1] == 0x0)
                    {
                        if (data.Length == 1)
                        {
                            strData = "";
                        }
                        else
                        {
                            strData = strData.Substring(0, data.Length - 1);
                        }
                    }

                    // Poke受信イベント発行
                    OnPokeEvent(item, strData);
                }
                return PokeResult.Processed;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message + "\n" + ex.StackTrace, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return PokeResult.Processed;
            }
        }

        protected override byte[] OnAdvise(string topic, string item, int format)
        {
            byte[] output = new byte[0];
            return output;
        }

        #endregion
    }
}

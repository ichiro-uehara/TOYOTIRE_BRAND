using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DDEServer;

namespace SpecInfo
{
    /// <summary>
    /// 編集文字スペッククラス
    /// </summary>
    public class SpecEditingChar
    {
        #region 定数

        /// <summary>
        /// スペック201
        /// </summary>
        private const string Spec201 = "SPEC201-";

        /// <summary>
        /// スペック2011
        /// </summary>
        private const string Spec2011 = "SPEC2011";

        #endregion

        #region publicプロパティ

        /// <summary>
        /// 削除フラグ
        /// </summary>
        public short flag_delete { get; set; } = 0;

        /// <summary>
        /// ＩＤ(HM固定)
        /// </summary>
        public string id { get; set; } = "HM";

        /// <summary>
        /// フォント名
        /// </summary>
        public string font_name { get; set; } = "";

        /// <summary>
        /// 区分番号（00～99の自動連番）
        /// </summary>
        public string no { get; set; } = "";

        /// <summary>
        /// スペル
        /// </summary>
        public string spell { get; set; } = new String(' ', 255);

        /// <summary>
        /// 配置方法
        /// </summary>
        public short haiti_sitei { get; set; } = 0;

        /// <summary>
        /// 原始文字数
        /// </summary>
        public short gm_num { get; set; } = 0;

        /// <summary>
        /// 幅
        /// </summary>
        public double width { get; set; } = 0.0;

        /// <summary>
        /// 高さ
        /// </summary>
        public double high { get; set; } = 0.0;

        /// <summary>
        /// 角度
        /// </summary>
        public double ang { get; set; } = 0.0;

        /// <summary>
        /// 配置PIC番号
        /// </summary>
        public short haiti_pic { get; set; } = 0;

        /// <summary>
        /// 刻印図面ID
        /// </summary>
        public string hz_id { get; set; } = "";

        /// <summary>
        /// 刻印図面番号1
        /// </summary>
        public string hz_no1 { get; set; } = "";

        /// <summary>
        /// 刻印図面番号2
        /// </summary>
        public string hz_no2 { get; set; } = "";

        /// <summary>
        /// コメント
        /// </summary>
        public string comment { get; set; } = "";

        /// <summary>
        /// 部署コード
        /// </summary>
        public string dep_name { get; set; } = "";

        /// <summary>
        /// 登録者
        /// </summary>
        public string entry_name { get; set; } = "";

        /// <summary>
        /// 登録日時
        /// </summary>
        public DateTime entry_date { get; set; } = DateTime.MinValue;

        /// <summary>
        /// 原始文字リスト
        /// </summary>
        public List<string> gm_list = new List<string>();

        #endregion

        #region publicメソッド

        /// <summary>
        /// スペル文字列設定
        /// </summary>
        /// <param name="spellStr">スペル文字列</param>
        public void SetSpell(string spellStr)
        {
            spell = spellStr + spell.Substring(spellStr.Length);
        }

        /// <summary>
        /// 原始文字リスト追加
        /// </summary>
        /// <param name="gmCode">原始文字コード</param>
        public void AddGmList(string gmCode)
        {
            gm_list.Add(gmCode);
        }

        /// <summary>
        /// 原始文字リスト追加
        /// </summary>
        /// <param name="gmCodeList">原始文字コードリスト</param>
        public void AddGmList(List<string> gmCodeList)
        {
            gm_list.Clear();

            foreach (string code in gmCodeList)
            {
                gm_list.Add(code);
            }
        }

        /// <summary>
        /// DDE通信用のスペック201データ取得
        /// </summary>
        /// <param name="isCreate">刻印文字の新規登録か否か</param>
        /// <returns>スペック201データ文字列</returns>
        public string GetSpec201ForDDE(bool isCreate = true)
        {
            string ret = Spec201 + id;

            if (isCreate)
            {
                ret += "      ";
                ret += "  ";

            }
            else
            {
                ret += font_name;
                ret += no;
            }

            ret += spell;

            ret += DDEHexConv.ShortToHex(haiti_sitei);

            ret += DDEHexConv.ShortToHex(gm_num);

            ret += DDEHexConv.DoubleToHex(width);
            ret += DDEHexConv.DoubleToHex(high);
            ret += DDEHexConv.DoubleToHex(ang);

            ret += DDEHexConv.ShortToHex(haiti_pic);

            return ret;
        }

        /// <summary>
        /// DDE通信用のスペック2011データ取得
        /// </summary>
        /// <returns>スペック2011データ文字列</returns>
        public string GetSpec2011ForDDE()
        {
            string ret = Spec2011;

            foreach (string code in gm_list)
            {
                ret += code;
            }

            return ret;
        }

        /// <summary>
        /// DDE通信データからスペックを設定
        /// </summary>
        /// <param name="recieveData">DDE通信データ</param>
        public void SetSpecDataByDDE(string recieveData)
        {
            font_name = recieveData.Substring(2, 6);
            no = recieveData.Substring(8, 2);
            haiti_sitei = DDEHexConv.HexToShort(recieveData.Substring(10, 4));
            gm_num = DDEHexConv.HexToShort(recieveData.Substring(14, 4));

            width = DDEHexConv.HexToDouble(recieveData.Substring(18, 16));
            high = DDEHexConv.HexToDouble(recieveData.Substring(34, 16));
            ang = DDEHexConv.HexToDouble(recieveData.Substring(50, 16));

            haiti_pic = DDEHexConv.HexToShort(recieveData.Substring(66, 4));

            spell = recieveData.Substring(70, 255);
        }

        /// <summary>
        /// ＝演算子
        /// </summary>
        /// <param name="left">左辺</param>
        /// <param name="right">右辺</param>
        /// <returns>左辺＝右辺</returns>
        public static bool operator == (SpecEditingChar left, SpecEditingChar right)
        {
            if (left is null || right is null)
            {
                return false;
            }

            if (left.haiti_sitei != right.haiti_sitei)
            {
                return false;
            }

            if (left.gm_num != right.gm_num)
            {
                return false;
            }

            if (left.width != right.width)
            {
                return false;
            }
            if (left.high != right.high)
            {
                return false;
            }
            if (left.ang != right.ang)
            {
                return false;
            }

            //if (left.haiti_pic != right.haiti_pic)
            //{
            //    return false;
            //}

            return true;
        }

        /// <summary>
        /// ≠演算子
        /// </summary>
        /// <param name="left">左辺</param>
        /// <param name="right">右辺</param>
        /// <returns>左辺≠右辺</returns>
        public static bool operator != (SpecEditingChar left, SpecEditingChar right)
        {
            if (left is null || right is null)
            {
                return true;
            }

            return !(left == right);
        }

        /// <summary>
        /// クローン
        /// </summary>
        /// <returns>編集文字スペッククラスのクローン</returns>
        public SpecEditingChar Clone()
        {
            return (SpecEditingChar)MemberwiseClone();
        }

        #endregion
    }
}

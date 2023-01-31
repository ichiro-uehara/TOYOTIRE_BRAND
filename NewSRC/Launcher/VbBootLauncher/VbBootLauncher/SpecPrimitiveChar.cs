using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DDEServer;

namespace SpecInfo
{
    /// <summary>
    /// 刻印文字スペッククラス
    /// </summary>
    public class SpecPrimitiveChar
    {
        #region 定数

        /// <summary>
        /// スペック101
        /// </summary>
        private const string Spec101 = "SPEC101-";

        #endregion

        #region publicプロパティ

        /// <summary>
        /// 削除フラグ
        /// </summary>
        public short flag_delete { get; set; } = 0;

        /// <summary>
        /// ＩＤ(GM固定)
        /// </summary>
        public string id { get; set; } = "GM";

        /// <summary>
        /// フォント名
        /// </summary>
        public string font_name { get; set; } = "";

        /// <summary>
        /// フォント区分1(A,F,H,B,D,P,N）
        /// </summary>
        public string font_class1 { get; set; } = "";

        /// <summary>
        /// フォント区分2(0～9: 自動連番）
        /// </summary>
        public string font_class2 { get; set; } = "";

        /// <summary>
        /// 文字名1
        /// </summary>
        public string name1 { get; set; } = "";

        /// <summary>
        /// 文字名2
        /// </summary>
        public string name2 { get; set; } = "";

        /// <summary>
        /// 高さ
        /// </summary>
        public double high { get; set; } = 0.0;

        /// <summary>
        /// 幅
        /// </summary>
        public double width { get; set; } = 0.0;

        /// <summary>
        /// 角度
        /// </summary>
        public double ang { get; set; } = 0.0;

        /// <summary>
        /// 実高さ
        /// </summary>
        public double moji_high { get; set; } = 0.0;

        /// <summary>
        /// ずれ量
        /// </summary>
        public double moji_shift { get; set; } = 0.0;

        /// <summary>
        /// 水平原点位置(現在Cに固定)
        /// </summary>
        public string org_hor { get; set; } = "C";

        /// <summary>
        /// 垂直原点位置(現在Bに固定)
        /// </summary>
        public string org_ver { get; set; } = "B";

        /// <summary>
        /// 文字原点座標X
        /// </summary>
        public double org_x { get; set; } = 0.0;

        /// <summary>
        /// 文字原点座標Y
        /// </summary>
        public double org_y { get; set; } = 0.0;

        /// <summary>
        /// 枠左下座標X
        /// </summary>
        public double left_bottom_x { get; set; } = 0.0;

        /// <summary>
        /// 枠左下座標Y
        /// </summary>
        public double left_bottom_y { get; set; } = 0.0;

        /// <summary>
        /// 枠右下座標X
        /// </summary>
        public double right_bottom_x { get; set; } = 0.0;

        /// <summary>
        /// 枠右下座標Y
        /// </summary>
        public double right_bottom_y { get; set; } = 0.0;

        /// <summary>
        /// 枠右上座標X
        /// </summary>
        public double right_top_x { get; set; } = 0.0;

        /// <summary>
        /// 枠右上座標Y
        /// </summary>
        public double right_top_y { get; set; } = 0.0;

        /// <summary>
        /// 枠左上座標X
        /// </summary>
        public double left_top_x { get; set; } = 0.0;

        /// <summary>
        /// 枠左上座標Y
        /// </summary>
        public double left_top_y { get; set; } = 0.0;

        /// <summary>
        /// 縁取り幅
        /// </summary>
        public double hem_width { get; set; } = 0.0;

        /// <summary>
        /// ハッチング角度
        /// </summary>
        public double hatch_ang { get; set; } = 0.0;

        /// <summary>
        /// ハッチング幅
        /// </summary>
        public double hatch_width { get; set; } = 0.0;

        /// <summary>
        /// ハッチング間隔
        /// </summary>
        public double hatch_space { get; set; } = 0.0;

        /// <summary>
        /// ハッチング始点X
        /// </summary>
        public double hatch_x { get; set; } = 0.0;

        /// <summary>
        /// ハッチング始点Y
        /// </summary>
        public double hatch_y { get; set; } = 0.0;

        /// <summary>
        /// 基準Ｒ
        /// </summary>
        public double base_r { get; set; } = 0.0;

        /// <summary>
        /// 旧フォント名
        /// </summary>
        public string old_font_name { get; set; } = "";

        /// <summary>
        /// 旧フォント区分
        /// </summary>
        public string old_font_class { get; set; } = "";

        /// <summary>
        /// 旧フォント文字名
        /// </summary>
        public string old_name { get; set; } = "";

        /// <summary>
        /// 配置PIC番号
        /// </summary>
        public short haiti_pic { get; set; } = 0;

        /// <summary>
        /// 刻印図面ID
        /// </summary>
        public string gz_id { get; set; } = "";

        /// <summary>
        /// 刻印図面番号1
        /// </summary>
        public string gz_no1 { get; set; } = "";

        /// <summary>
        /// 刻印図面番号2
        /// </summary>
        public string gz_no2 { get; set; } = "";

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

        #endregion

        #region publicメソッド

        /// <summary>
        /// DDE通信用のスペック101データ取得
        /// </summary>
        /// <param name="isCreate">刻印文字の新規登録か否か</param>
        /// <returns>スペック101データ文字列</returns>
        public string GetSpec101ForDDE(bool isCreate = true)
        {
            string ret = Spec101 + id;

            if (isCreate)
            {
                ret += "          ";
            }
            else
            {
                ret += (font_name + font_class1 + font_class2 + name1 + name2);
            }

            ret += DDEHexConv.DoubleToHex(high);
            ret += DDEHexConv.DoubleToHex(width);
            ret += DDEHexConv.DoubleToHex(ang);

            ret += DDEHexConv.DoubleToHex(moji_high);
            ret += DDEHexConv.DoubleToHex(moji_shift);

            ret += org_hor;
            ret += org_ver;

            ret += DDEHexConv.DoubleToHex(org_x);
            ret += DDEHexConv.DoubleToHex(org_y);
            ret += DDEHexConv.DoubleToHex(left_bottom_x);
            ret += DDEHexConv.DoubleToHex(left_bottom_y);
            ret += DDEHexConv.DoubleToHex(right_bottom_x);
            ret += DDEHexConv.DoubleToHex(right_bottom_y);
            ret += DDEHexConv.DoubleToHex(right_top_x);
            ret += DDEHexConv.DoubleToHex(right_top_y);
            ret += DDEHexConv.DoubleToHex(left_top_x);
            ret += DDEHexConv.DoubleToHex(left_top_y);

            ret += DDEHexConv.DoubleToHex(hem_width);

            ret += DDEHexConv.DoubleToHex(hatch_ang);
            ret += DDEHexConv.DoubleToHex(hatch_width);
            ret += DDEHexConv.DoubleToHex(hatch_space);
            ret += DDEHexConv.DoubleToHex(hatch_x);
            ret += DDEHexConv.DoubleToHex(hatch_y);

            ret += DDEHexConv.DoubleToHex(base_r);

            // 旧フォント名
            ret += "      ";

            // 旧フォント区分
            ret += "  ";

            // 旧文字名
            ret += "  ";

            ret += DDEHexConv.ShortToHex(haiti_pic);

            return ret;
        }

        /// <summary>
        /// DDE通信データからスペックを設定
        /// </summary>
        /// <param name="recieveData">DDE通信データ</param>
        public void SetSpecDataByDDE(string recieveData)
        {
            font_name = recieveData.Substring(2, 6);
            font_class1 = recieveData.Substring(8, 1);
            font_class2 = recieveData.Substring(9, 1);
            name1 = recieveData.Substring(10, 1);
            name2 = recieveData.Substring(11, 1);

            high = DDEHexConv.HexToDouble(recieveData.Substring(12, 16));
            width = DDEHexConv.HexToDouble(recieveData.Substring(28, 16));
            ang = DDEHexConv.HexToDouble(recieveData.Substring(44, 16));

            moji_high = DDEHexConv.HexToDouble(recieveData.Substring(60, 16));
            moji_shift = DDEHexConv.HexToDouble(recieveData.Substring(76, 16));

            org_hor = recieveData.Substring(92, 1);
            org_ver = recieveData.Substring(93, 1);

            org_x = DDEHexConv.HexToDouble(recieveData.Substring(94, 16));
            org_y = DDEHexConv.HexToDouble(recieveData.Substring(110, 16));

            left_bottom_x = DDEHexConv.HexToDouble(recieveData.Substring(126, 16));
            left_bottom_y = DDEHexConv.HexToDouble(recieveData.Substring(142, 16));
            right_bottom_x = DDEHexConv.HexToDouble(recieveData.Substring(158, 16));
            right_bottom_y = DDEHexConv.HexToDouble(recieveData.Substring(174, 16));

            right_top_x = DDEHexConv.HexToDouble(recieveData.Substring(190, 16));
            right_top_y = DDEHexConv.HexToDouble(recieveData.Substring(206, 16));
            left_top_x = DDEHexConv.HexToDouble(recieveData.Substring(222, 16));
            left_top_y = DDEHexConv.HexToDouble(recieveData.Substring(238, 16));

            hem_width = DDEHexConv.HexToDouble(recieveData.Substring(254, 16));

            hatch_ang = DDEHexConv.HexToDouble(recieveData.Substring(270, 16));
            hatch_width = DDEHexConv.HexToDouble(recieveData.Substring(286, 16));
            hatch_space = DDEHexConv.HexToDouble(recieveData.Substring(302, 16));
            hatch_x = DDEHexConv.HexToDouble(recieveData.Substring(318, 16));
            hatch_y = DDEHexConv.HexToDouble(recieveData.Substring(334, 16));

            base_r = DDEHexConv.HexToDouble(recieveData.Substring(350, 16));

            old_font_name = recieveData.Substring(366, 6);
            old_font_class = recieveData.Substring(372, 2);
            old_name = recieveData.Substring(374, 2);

            haiti_pic = DDEHexConv.HexToShort(recieveData.Substring(376, 4));
        }

        /// <summary>
        /// クローン
        /// </summary>
        /// <returns>刻印文字スペッククラスのクローン</returns>
        public SpecPrimitiveChar Clone()
        {
            return (SpecPrimitiveChar)MemberwiseClone();
        }

        #endregion
    }
}

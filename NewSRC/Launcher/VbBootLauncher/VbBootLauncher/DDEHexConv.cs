using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace DDEServer
{
    /// <summary>
    /// DDE通信用16進数変換クラス
    /// </summary>
    public static class DDEHexConv
    {
        #region 数値 ⇒ Hex文字列

        /// <summary>
        /// byte型数値(1byte) ⇒ Hex文字列(2文字)
        /// </summary>
        /// <param name="value">変換値</param>
        /// <returns>Hex文字列</returns>
        public static string ByteToHex(byte value)
        {
            string ret = "";

            try
            {
                ret = value.ToString("X2");

                return ret;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message + "\n" + ex.StackTrace, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return ret;
            }
        }

        /// <summary>
        /// short型数値(2byte) ⇒ Hex文字列(4文字)
        /// </summary>
        /// <param name="value">変換値</param>
        /// <returns>Hex文字列</returns>
        public static string ShortToHex(short value)
        {
            string ret = "";

            try
            {
                byte[] byteArray = BitConverter.GetBytes(value);

                for (int i = 1; i >= 0; i--)
                {
                    ret += byteArray[i].ToString("X2");
                }

                return ret;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message + "\n" + ex.StackTrace, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return ret;
            }
        }

        /// <summary>
        /// int型数値(4byte) ⇒ Hex文字列(8文字)
        /// </summary>
        /// <param name="value">変換値</param>
        /// <returns>Hex文字列</returns>
        public static string IntToHex(int value)
        {
            string ret = "";

            try
            {
                byte[] byteArray = BitConverter.GetBytes(value);

                for (int i = 3; i >= 0; i--)
                {
                    ret += byteArray[i].ToString("X2");
                }

                return ret;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message + "\n" + ex.StackTrace, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return ret;
            }
        }

        /// <summary>
        /// float型数値(4byte) ⇒ Hex文字列(8文字)
        /// </summary>
        /// <param name="value">変換値</param>
        /// <returns>Hex文字列</returns>
        public static string FloatToHex(float value)
        {
            string ret = "";

            try
            {
                byte[] byteArray = BitConverter.GetBytes(value);
                
                for (int i = 3; i >= 0; i--)
                {
                    ret += byteArray[i].ToString("X2");
                }

                return ret;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message + "\n" + ex.StackTrace, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return ret;
            }
        }

        /// <summary>
        /// double型数値(8byte) ⇒ Hex文字列(16文字)
        /// </summary>
        /// <param name="value">変換値</param>
        /// <returns>Hex文字列</returns>
        public static string DoubleToHex(double value)
        {
            string ret = "";

            try
            {
                byte[] byteArray = BitConverter.GetBytes(value);
                
                for (int i = 3; i >= 0; i--)
                {
                    ret += byteArray[i].ToString("X2");
                }

                for (int i = 7; i >= 4; i--)
                {
                    ret += byteArray[i].ToString("X2");
                }

                return ret;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message + "\n" + ex.StackTrace, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return ret;
            }
        }

        #endregion

        #region Hex文字列 ⇒ 数値

        /// <summary>
        /// Hex文字列(2文字) ⇒ byte型数値(1byte)
        /// </summary>
        /// <param name="hex">Hex文字列</param>
        /// <returns>変換値</returns>
        public static byte HexToByte(string hex)
        {
            byte ret = 0;

            try
            {
                ret = Convert.ToByte(hex.Substring(0, 2), 16);

                return ret;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message + "\n" + ex.StackTrace, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return ret;
            }
        }

        /// <summary>
        /// Hex文字列(4文字) ⇒ short型数値(2byte)
        /// </summary>
        /// <param name="hex">Hex文字列</param>
        /// <returns>変換値</returns>
        public static short HexToShort(string hex)
        {
            short ret = 0;

            try
            {
                byte[] bytes = new byte[2];

                int index = 2;
                for (int i = 0; i < 2; i++)
                {
                    bytes[i] = Convert.ToByte(hex.Substring(index, 2), 16);
                    index -= 2;
                }

                ret = BitConverter.ToInt16(bytes, 0);

                return ret;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message + "\n" + ex.StackTrace, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return ret;
            }
        }

        /// <summary>
        /// Hex文字列(8文字) ⇒ int型数値(4byte)
        /// </summary>
        /// <param name="hex">Hex文字列</param>
        /// <returns>変換値</returns>
        public static int HexToInt(string hex)
        {
            int ret = 0;

            try
            {
                byte[] bytes = new byte[4];

                int index = 6;
                for (int i = 0; i < 4; i++)
                {
                    bytes[i] = Convert.ToByte(hex.Substring(index, 2), 16);
                    index -= 2;
                }

                ret = BitConverter.ToInt32(bytes, 0);

                return ret;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message + "\n" + ex.StackTrace, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return ret;
            }
        }

        /// <summary>
        /// Hex文字列(8文字) ⇒ float型数値(4byte)
        /// </summary>
        /// <param name="hex">Hex文字列</param>
        /// <returns>変換値</returns>
        public static float HexToFloat(string hex)
        {
            float ret = 0.0F;

            try
            {
                byte[] bytes = new byte[4];

                int index = 6;
                for (int i = 0; i < 4; i++)
                {
                    bytes[i] = Convert.ToByte(hex.Substring(index, 2), 16);
                    index -= 2;
                }

                ret = BitConverter.ToSingle(bytes, 0);

                return ret;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message + "\n" + ex.StackTrace, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return ret;
            }
        }

        /// <summary>
        /// Hex文字列(16文字) ⇒ double型数値(8byte)
        /// </summary>
        /// <param name="hex">Hex文字列</param>
        /// <returns>変換値</returns>
        public static double HexToDouble(string hex)
        {
            double ret = 0.0;

            try
            {
                byte[] bytes = new byte[8];

                int index = 6;
                for (int i = 0; i < 4; i++)
                {
                    bytes[i] = Convert.ToByte(hex.Substring(index, 2), 16);
                    index -= 2;
                }

                index = 14;
                for (int i = 4; i < 8; i++)
                {
                    bytes[i] = Convert.ToByte(hex.Substring(index, 2), 16);
                    index -= 2;
                }

                ret = BitConverter.ToDouble(bytes, 0);

                return ret;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message + "\n" + ex.StackTrace, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return ret;
            }
        }

        #endregion
    }
}

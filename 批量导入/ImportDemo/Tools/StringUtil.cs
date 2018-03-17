using Microsoft.International.Converters.PinYinConverter;
using Microsoft.VisualBasic;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace Tools
{
    /// <summary>
    /// 字符处理拓展类
    /// </summary>
    public static class StringUtil
    {
        /// <summary>
        /// 将字符串转换为简体中文
        /// </summary>
        /// <param name="s">字符串</param>
        /// <returns>转换后的</returns>
        public static string ToSimplifiedChinese(this string s)
        {
            return Strings.StrConv(s, VbStrConv.SimplifiedChinese, 0);
        }

        /// <summary>
        /// 将字符串转换为繁体中文
        /// </summary>
        /// <param name="s">字符串</param>
        /// <returns>转换后的</returns>
        public static string ToTraditionalChinese(this string s)
        {
            return Strings.StrConv(s, VbStrConv.TraditionalChinese, 0);
        }

        /// <summary>
        /// 将单个单词首字母转换成大写
        /// </summary>
        /// <param name="s">字符串</param>
        /// <returns>转换后的</returns>
        public static string ToTitleCase(this string s)
        {
            return s.Substring(0, 1).ToUpper() + s.Substring(1).ToLower();
        }


        /// <summary>
        /// 将一段语句首字母转换成大写
        /// </summary>
        /// <param name="s">字符串</param>
        /// <returns>转换后的</returns>
        public static string StrConv(this string s)
        {
            return Strings.StrConv(s, VbStrConv.ProperCase, System.Globalization.CultureInfo.CurrentCulture.LCID);
        }

        /// <summary>
        /// 判断字符串中是否包含中文
        /// </summary>
        /// <param name="str">需要判断的字符串</param>
        /// <returns>判断结果</returns>
        public static bool HasChinese(this string str)
        {
            return Regex.IsMatch(str, @"[\u4e00-\u9fa5]");
        }

        /// <summary>
        /// 判断字符是否为中文
        /// </summary>
        /// <param name="str">需要判断的字符串</param>
        /// <returns>判断结果</returns>
        public static bool IsChinese(this char c)
        {
            return c >= 0x4e00 && c <= 0x9fa5;
        }

        /// <summary> 
        /// 汉字转化为拼音
        /// </summary> 
        /// <param name="str">汉字</param> 
        /// <returns>全拼</returns> 
        public static string GetPinyin(this string str)
        {
            if (!str.HasChinese())
            {
                return str;
            }
            string r = string.Empty;
            string t = string.Empty;
            ChineseChar chineseChar = null;
            foreach (char c in str)
            {
                if (c.IsChinese())
                {
                    chineseChar = new ChineseChar(c);
                    t = chineseChar.Pinyins[0].ToString();
                    r += t.Substring(0, t.Length - 1);
                }
                else
                {
                    r += c;
                }
            }
            return r;
        }

        /// <summary> 
        /// 汉字转化为拼音首字母
        /// </summary> 
        /// <param name="str">汉字</param> 
        /// <returns>首字母</returns> 
        public static string GetFirstPinyin(this string str)
        {
            if (!str.HasChinese())
            {
                return str;
            }
            string r = string.Empty;
            string t = string.Empty;
            ChineseChar chineseChar = null;
            foreach (char c in str)
            {
                if (c.IsChinese())
                {
                    chineseChar = new ChineseChar(c);
                    t = chineseChar.Pinyins[0].ToString();
                    r += t.Substring(0, 1);
                }
                else
                {
                    r += c;
                }
            }
            return r;
        }

        /// <summary>
        /// 获取汉字对应的拼音首字母 和全拼组合中间空格隔开
        /// </summary>
        /// <example> 测试路口  cclk ceshilukou</example>
        /// <param name="str">汉字</param>
        /// <returns>拼音首字母 和全拼组合中间空格隔开</returns>
        /// <remarks>
        /// 2016-06-24 杜冬军修改计算全拼速度慢的问题，慢的原因由于字符串不包含中文或者中英文混合 这种情况速度很慢
        /// </remarks>
        public static string GetBopomofo(this string str)
        {
            //参数为空 或者不包含中文则直接返回
            if (string.IsNullOrEmpty(str) || !str.HasChinese())
            {
                return str;
            }
            string rfirst = string.Empty;
            string rfull = string.Empty;
            string t = string.Empty;
            ChineseChar chineseChar = null;
            foreach (char c in str)
            {
                if (c.IsChinese())
                {
                    chineseChar = new ChineseChar(c);
                    t = chineseChar.Pinyins[0].ToString();
                    rfirst += t.Substring(0, 1);
                    rfull += t.Substring(0, t.Length - 1);
                }
                else
                {
                    rfirst += c;
                    rfull += c;
                }
            }

            string py = rfirst + " " + rfull;
            if (!string.IsNullOrEmpty(py))
            {
                return py.Substring(0, py.Length < 200 ? py.Length : 200);
            }
            return py;
        }

        /// <summary>
        /// 判断字符串是否为中文
        /// </summary>
        /// <param name="str">字符串</param>
        /// <returns>bool</returns>
        public static bool IsChinese(this string str)
        {
            return Regex.IsMatch(str, @"^[\u4e00-\u9fa5]+$");
        }
    }
}

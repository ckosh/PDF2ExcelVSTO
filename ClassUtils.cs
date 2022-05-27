using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using static iText.Kernel.Pdf.Colorspace.PdfSpecialCs;
using System.Web;
using System.Net;
using Newtonsoft.Json.Linq;
using System.Drawing;

namespace PDF2ExcelVsto
{
    static public class ClassUtils
    {
        
        public static string ColumnLabel(this int col)
        {
            var dividend = col;
            var columnLabel = string.Empty;
            int modulo;

            while (dividend > 0)
            {
                modulo = (dividend - 1) % 26;
                columnLabel = Convert.ToChar(65 + modulo).ToString() + columnLabel;
                dividend = (int)((dividend - modulo) / 26);
            }

            return columnLabel;
        }
        public static int ColumnIndex(this string colLabel)
        {
            // "AD" (1 * 26^1) + (4 * 26^0) ...
            var colIndex = 0;
            for (int ind = 0, pow = colLabel.Count() - 1; ind < colLabel.Count(); ++ind, --pow)
            {
                var cVal = Convert.ToInt32(colLabel[ind]) - 64; //col A is index 1
                colIndex += cVal * ((int)Math.Pow(26, pow));
            }
            return colIndex;
        }

        public static List<string> ConvertToHebrew0(List<string> row)
        {
            char char10 = (char)10;
            var dataRow = new List<string>();
            List<string> newDataRow = new List<string>();
            dataRow = row;

            for (int j = 0; j < dataRow.Count; j++)
            {
                string value;
                value = dataRow[j];
                if (value == "") continue;

                if (isreverible(value))
                {
                    value = convertInternalNumber(value);
                    value = stringReverseString1(value);
                    char[] characters = value.ToCharArray();
                }
                else if (value.IndexOf(char10) > -1)
                {
                    string[] words = value.Split(char10);
                    value = words[1] + char10 + words[0];
                }
                newDataRow.Add(value);

            }
            return newDataRow;
        }
        private static string convertInternalNumber(string val)
        {
            string ss;
            string retval = "";
            char char10 = (char)10;
            char char12 = (char)12;
            val = val.Replace(char10, ' ');
            val = val.Replace(char12, ' ');
            //            val = val.Replace('-', ' ');
            string[] Results = val.Split(' ');
            for (int i = 0; i < Results.Length; i++)
            {
                Results[i] = dealWithBrakets(Results[i]);
                ss = Results[i];
                if (!isreverible(ss))
                {
                    ss = stringReverseString1(ss);
                    Results[i] = ss;
                }
            }
            retval = string.Join(" ", Results);
            return retval;
        }
        public static string stringReverseString1(string str)
        {
            char[] chars = str.ToCharArray();
            char[] result = new char[chars.Length];
            for (int i = 0, j = str.Length - 1; i < str.Length; i++, j--)
            {
                result[i] = chars[j];
            }
            return new string(result);
        }

        private static string dealWithBrakets(string val)
        {
            string retval = "";

            int pOpen;
            int pClos;
            pOpen = val.IndexOf('(');
            pClos = val.IndexOf(')');
            if (pOpen > -1) val = val.Replace('(', '&');
            if (pClos > -1) val = val.Replace(')', '#');
            val = val.Replace('&', ')');
            val = val.Replace('#', '(');
            retval = val;
            return retval;
        }

        public static bool isreverible(string astring)
        {
            bool rev = false;
            if (!containHebrew(astring)) return rev;
            char achar;
            for (int i = 0; i < astring.Length; i++)
            {
                achar = astring[i];
                if ((int)achar == 164) return rev;
                if (char.IsDigit(achar) ||
                     achar.CompareTo('\\') == 0 ||
                     achar.CompareTo(':') == 0 ||
                     achar.CompareTo(' ') == 0 ||
                     achar.CompareTo('/') == 0 ||
                     achar.CompareTo('-') == 0 ||
                     achar.CompareTo('.') == 0 ||
                     achar.CompareTo(',') == 0) continue;
                rev = true;
                return rev;
            }
            return rev;
        }

        public static string ConvertToHebrew(string s)
        {
            string combindedString;
            string nl = "\n";
            List<string> results = new List<string>();
            List<string> raw = s.Split('\n').ToList();
            foreach (string ss in raw)
            {
                List<string> s0 = ss.Split(' ').ToList();
                foreach (string s1 in s0)
                {
                    if (s1 == "") continue;
                    if (s1.IndexOf(':') != -1)
                    {
                        if (s1.IndexOf(':') == 0 || s1.IndexOf(':') == s1.Length - 1)
                        {
                            results.Add(ReverseExcludeChar(s1, ':'));
                            continue;
                        }
                        else
                        {
                            results.Add(s1);
                            continue;
                        }
                    }
                    if (s1.All(Char.IsDigit) || isDate(s1) || isShtarNumber(s1) || isFloatingNumber(s1))
                    {
                        results.Add(s1);
                    }
                    else
                    {
                        results.Add(Reverse(s1));
                    }
                }
                results.Add(nl);
            }
            combindedString = string.Join(" ", results.ToArray());
            return combindedString;
        }

        public static string ReverseWordsInString(string s1)
        {
            string res = "";
            string[] subs = s1.Split(' ');
            for (int i = subs.Length - 1; i >= 0; i--)
            {
                res = res + " " + subs[i];
            }
            return res;
        }
        public static string Reverse(string s)
        {
            char[] charArray = s.ToCharArray();
            Array.Reverse(charArray);
            return new string(charArray);
        }
        public static string ReverseExcludeChar(string s1, char cr)
        {
            string ret;
            int ind = s1.IndexOf(cr);
            if (ind == 0)
            {
                ret = s1.Substring(1);
            }
            else
            {
                ret = s1.Substring(0, s1.Length - 2);
            }
            ret = ":" + Reverse(ret);
            return ret;
        }
        public static bool isDate(string sss)
        {
            bool bret = false;
            Regex rx = new Regex(@"^([0-2][0-9]|(3)[0-1])(\/)(((0)[0-9])|((1)[0-2]))(\/)\d{4}$", RegexOptions.Compiled | RegexOptions.IgnoreCase);
            MatchCollection matches = rx.Matches(sss);
            if (matches.Count > 0)
            {
                bret = true;
            }
            return bret;
        }
        public static bool isIdType(string sss)
        {
            bool bret = false;
            if (sss == "חברה" || sss == "ת.ז" || sss == "דרכון") bret = true;
            return bret;
        }
        public static bool isFloating(string sss)
        {
            bool bret = false;
            Regex rx = new Regex(@"^[0-9]{1,3}(,[0-9]{3})*(\.[0-9]+)?$", RegexOptions.Compiled | RegexOptions.IgnoreCase);
            MatchCollection matches = rx.Matches(sss);
            if (matches.Count > 0)
            {
                bret = true;
            }
            return bret;
        }
        public static bool isShtarNumber(string sss)
        {
            bool bret = false;

            string[] fff = sss.Split('/');
            if (fff.Length == 3 || fff.Length == 2)
            {
                if ( Int32.Parse(fff[0]) > 31 || Int32.Parse(fff[1]) > 12)
                {
                    bret = true;
                } 
                
            }
            return bret;
        }
        public static bool isFloatingNumber(string sss)
        {
            bool bret = false;
            ///   ^([1-9]\d{0,2}(,?\d{3})*|0)(\.\d+[1-9])?$
            ///   ^[-+]?[0-9]*\.[0-9]+$
            Regex rx = new Regex(@"^(?!0|\.00)[0-9]+(,\d{3})*(.[0-9]{0,2})$", RegexOptions.Compiled | RegexOptions.IgnoreCase);
            MatchCollection matches = rx.Matches(sss);
            if (matches.Count > 0)
            {
                bret = true;
            }
            return bret;
        }
        static public bool isStringOneOfParams(string val, params string[] list)
        {
            bool bret = false;
            for (int i = 0; i < list.Length; i++)
            {
                if (val == list[i])
                {
                    bret = true;
                    break;
                }
            }
            return bret;
        }
        static public int isArrayIncludString(List<string> array, string sss)
        {
            int ret = -1; ;
            int ind;
            for (int i = 0; i < array.Count; i++)
            {
                if ((ind = array[i].IndexOf(sss)) > -1)
                {
                    ret = ind;
                    break;
                }
            }
            return ret;
        }
        public static string buildCombinedline(List<string> aLine)
        {
            string retstr = "";
            for (int i = aLine.Count - 1; i > -1; i--)
            {
                retstr += aLine[i] + " ";
            }
            retstr = retstr.Substring(0, retstr.Length - 1);
            return retstr;
        }

        public static string buildReverseCombinedLine(List<string> aLine)
        {
            string retstr = "";
            if (aLine.Count > 0)
            {
                for (int i = 0; i < aLine.Count; i++)
                {
                    retstr += aLine[i] + " ";
                }
                retstr = retstr.Substring(0, retstr.Length - 1);
            }
            return retstr;
        }
        static public List<List<string>> RemoveHeaderSection(List<List<string>> ddd, string foot, string topHeader)
        {
            List<List<string>> retData = new List<List<string>>();
            int j = 0;
            for (int i = 0; i < ddd.Count; i++)
            {
                string ppp = ClassUtils.buildCombinedline(ddd[i]);
                if (ddd[i].Count == 0 || ppp == topHeader)
                {
                    for (j = i + 1; j < ddd.Count; j++)
                    {
                        string sss = ClassUtils.buildCombinedline(ddd[j]);
                        if (sss == foot)
                        {
                            i = j;
                            break;
                        }
                    }
                }
                else
                {
                    if (ClassUtils.isArrayIncludString(ddd[i], "עמוד") > -1 && ClassUtils.isArrayIncludString(ddd[i], "מתוך") > -1)
                    {
                    }
                    else if (ddd[i].Count == 1 && ddd[i][0] == "")
                    {
                    }
                    else
                    {
                        retData.Add(ddd[i]);
                    }

                }
            }
            return retData;
        }
        static public bool isArrayIsUniqueInLine(List<string> array, string sectionName)
        {
            bool bret = false;
            string[] subs = sectionName.Split(' ');
            if (array.Count == subs.Length)
            {
                string ss = buildCombinedline(array);
                if (ss == sectionName)
                {
                    bret = true;
                }
            }
            return bret;
        }
        static public bool isArrayIncludeOneOfStringParam(List<string> array, params string[] list)
        {
            bool bret = false;
            for (int i = 0; i < list.Length; i++)
            {
                if ((isArrayIncludString(array, list[i])) > -1)
                {
                    bret = true;
                    break;
                }
            }
            return bret;
        }
        static public bool isArrayIncludeAllStringsParamFromBeggining(List<string> array, params string[] list)
        {
            bool bret = true;
            bret = isArrayIncludeAllStringsParam(array, list);
            if ( bret )
            {
                array.Reverse();
                if (array[0] != list[0]) bret = false;
                array.Reverse();

            }
            return bret;
        }
        static public bool isArrayIncludeAllStringsParam(List<string> array, params string[] list)
        {
            bool bret = true;
            for (int i = 0; i < list.Length; i++)
            {
                if ((isArrayIncludString(array, list[i])) == -1)
                {
                    bret = false;
                    break;
                }
            }
            return bret;
        }
        static public bool isArrayIncludeAllStrings(List<string> array, List<string> subArray)
        {
            bool bret = true;
            int instring = 0;
            for (int i = 0; i < subArray.Count; i++)
            {
                if ((instring = isArrayIncludString(array, subArray[i])) == -1)
                {
                    bret = false;
                    break;
                }
            }
            return bret;
        }
        static public List<string> reverseOrder(List<string> array)
        {
            List<string> sss = new List<string>();
            for (int i = array.Count - 1; i > -1; i--)
            {
                sss.Add(array[i]);
            }
            return sss;
        }
        public static bool isStringIncludesSubstring(string s0, string s1)
        {
            bool bret = false;
            string[] subs = s0.Split(' ');
            foreach (var z in subs)
            {
                if (z == s1)
                {
                    bret = true;
                    break;
                }
            }

            return bret;
        }
        static public bool isMatchSequence(List<string> lineData, int jstart, params string[] list)
        {
            bool bret = true;
            for (int i = 0; i < list.Length; i++)
            {
                if (lineData[jstart - i] != list[i])    /// was jstart -1
                {
                    bret = false;
                    break;
                }
            }
            return bret;
        }
        
        static public bool isMatchSequenceNormal(List<string> lineData, int jstart, params string[] list)
        {
            bool bret = true;
            for (int i = 0; i < list.Length; i++)
            {
                if (lineData[jstart+i] != list[i])   
                {
                    bret = false;
                    break;
                }
            }
            return bret;
        }

        static public bool isMatchSequenceStright(List<string> lineData, int jstart, params string[] list)
        {
            bool bret = true;
            if (list.Length > lineData.Count - jstart + 1)
            {
                bret = false;
            }
            else
            {
                for (int i = 0; i < list.Length; i++)
                {
                    if (lineData[jstart - 1 + i] != list[i])    /// was jstart -1
                    {
                        bret = false;
                        break;
                    }
                }
            }
            return bret;
        }

        public static bool isForeignID(string sss)
        {
            bool bret = true;
            Regex re = new Regex(@"([A-Za-z]+)");
            Match m = re.Match(sss);
            bret = m.Success;
            return bret;

        }
        public static bool getNumberOfdigits(string sss)
        {
            bool bret = true;
            Regex re = new Regex(@"(?<=^| )\d+(\/\d+)?(?=$| )");
            Match m = re.Match(sss);
            bret = m.Success;
            return bret;
        }

        public static bool isAllDigit(string sss)
        {
            bool bret = false;
            bret = sss.All(Char.IsDigit);
            return bret;
        }
        public static string buildCombinedLineSelected(List<string> aLine, int start, int end, bool rev)
        {
            string retstr = "";
            if (!rev)
            {
                for (int i = start; i < end; i++)
                {
                    retstr += aLine[i] + " ";
                }
            }
            else
            {
                for (int i = end - 1; i > start - 1; i--)
                {
                    retstr += aLine[i] + " ";
                }
            }
            return retstr;
        }
        public static List<string> removeAllBlancs(string[] s)
        {
            List<string> ret = new List<string>();
            for (int i = 0; i < s.Length; i++)
            {
                if (s[i] == "") continue;
                ret.Add(s[i]);
            }
            return ret;
        }
        public static bool IsIDNumber(string sss)
        {
            // min 5 , max 10, no blank , min 5 digits , 
            bool bret = false;

            if (sss.Length < 5 || sss.Length > 10) return bret;
            if (!sss.Any(ch => Char.IsDigit(ch)))
            {
                return bret;
            }
            bret = !sss.Any(ch => !Char.IsLetterOrDigit(ch));
            return bret;
        }
        public static bool containHebrew(string sss)
        {
            char FirstHebChar = (char)1488; //א
                                            //           char LastHebChar = (char)1514; //ת
            bool bret = false;
            char[] arr;
            arr = sss.ToCharArray();

            for (int i = 0; i < arr.Length; i++)
            {
                if (arr[i] >= FirstHebChar)
                {
                    bret = true;
                    break;
                }
            }
            return bret;
        }
        public static string convertPartToPercent(string parts)
        {
            string ret = "";
            if (parts != null )
            {
                
                NumberFormatInfo nfi = new CultureInfo("en-US", false).NumberFormat;
                nfi.PercentDecimalDigits = 4;
                string[] sub = parts.Split('/');
                if (parts.IndexOf("בשלמות") > -1)
                {
                    Double myval = 1.0;
                    ret = myval.ToString("P", nfi);
                }
                else if (sub.Length == 1)
                {
                    ret = parts;
                }
                else
                {
                    //string[] sub = parts.Split('/');
                    double a1 = Convert.ToDouble(sub[0]);
                    double a2 = Convert.ToDouble(sub[1]);
                    Double myval = a1 / a2;
                    ret = myval.ToString("P", nfi);
                }

            }
            return ret;
        }
        public static string convertPartToFraction(string parts)
        {
            string ret = "";
            if (parts == "בשלמות")
            {
                Double myval = 1.0;
                ret = myval.ToString("0.####");
            }
            else
            {
                string[] sub = parts.Split('/');
                double a1 = Convert.ToDouble(sub[0]);
                double a2 = Convert.ToDouble(sub[1]);
                Double myval = a1 / a2;
                ret = myval.ToString("0.####");
            }
            return ret;
        }
        public static List<List<string>> File2Data(string csvfile)
        {
            List<List<string>> datas = new List<List<string>>();
            List<string> sss = File.ReadAllLines(csvfile).ToList();
            for (int i = 0; i < sss.Count; i++)
            {
                string[] subs = sss[i].Split(' ');
                List<string> lll = subs.ToList();
                datas.Add(lll);
            }
            return datas;
        }
        
        public static Dictionary<string, int> GetLeasingAction(List<string> rawValue, int iv, bool reverse)
        {
            Dictionary<string, int> result = new Dictionary<string, int>();
            List<string> temp = new List<string>(rawValue);
            if (reverse) temp = reverseOrder(temp);
            string sss;

            if (ClassUtils.isMatchSequenceNormal(temp, iv, "בעברת", "שכירות", "עפ\"י", "צו", "בימ\"ש"))
            {
                sss = temp[iv] + " " + temp[iv + 1] + " " + temp[iv + 2] + " " + temp[iv + 3] + temp[iv + 4];
                result.Add(sss, 5);
            }
 
            else if (ClassUtils.isMatchSequenceNormal(temp, iv, "עדכון", "פרטי", "זיהוי", "-", "חוכר"))
            {
                sss = temp[iv] + " " + temp[iv + 1] + " " + temp[iv + 2] + " " + temp[iv + 3] + temp[iv + 4];
                result.Add(sss, 5);
            }
            else if (ClassUtils.isMatchSequenceNormal(temp, iv, "העברת", "שכירות", "בצוואה", "עפ\"י"))
            {
                sss = temp[iv] + " " + temp[iv + 1] + " " + temp[iv + 2] + " " + temp[iv + 3];
                result.Add(sss, 4);
                result.Add("הסכם", 1);
            }

            else if (ClassUtils.isMatchSequenceNormal(temp, iv, "העברת", "שכירות", "חלקית", "ללא"))
            {
                sss = temp[iv] + " " + temp[iv + 1] + " " + temp[iv + 2] + " " + temp[iv + 3];
                result.Add(sss, 4);
                result.Add("תמורה", 1);
            }

            else if (ClassUtils.isMatchSequenceNormal(temp, iv, "עודף", "לאחר", "העברת", "שכירות"))
            {
                sss = temp[iv] + " " + temp[iv + 1] + " " + temp[iv + 2] + " " + temp[iv + 3];
                result.Add(sss, 4);
                result.Add("חלקית", 1);
            }
            else if (ClassUtils.isMatchSequenceNormal(temp, iv, "העברת", "שכירות", "בירושה", "עפ\"י"))
            {
                sss = temp[iv] + " " + temp[iv + 1] + " " + temp[iv + 2] + " " + temp[iv + 3];
                result.Add(sss, 4);
                result.Add("הסכם", 1);
            }


            else if (ClassUtils.isMatchSequenceNormal(temp, iv, "העברת", "שכירות", "ללא", "תמורה"))
            {
                sss = temp[iv] + " " + temp[iv + 1] + " " + temp[iv + 2] + " " + temp[iv + 3];
                result.Add(sss, 4);
            }
            else if (ClassUtils.isMatchSequenceNormal(temp, iv, "העברת", "שכירות", "חלקית", "בירושה"))
            {
                sss = temp[iv] + " " + temp[iv + 1] + " " + temp[iv + 2] + " " + temp[iv + 3];
                result.Add(sss, 4);
            }
            else if (ClassUtils.isMatchSequenceNormal(temp, iv, "העברת", "שכירות", "ותיקון", "תנאיה"))
            {
                sss = temp[iv] + " " + temp[iv + 1] + " " + temp[iv + 2] + " " + temp[iv + 3];
                result.Add(sss, 4);
            }

            else if (ClassUtils.isMatchSequenceNormal(temp, iv, "תיקון", "טעות", "סופר", "בחוכר"))
            {
                sss = temp[iv] + " " + temp[iv + 1] + " " + temp[iv + 2] + " " + temp[iv + 3];
                result.Add(sss, 4);
            }
            else if (ClassUtils.isMatchSequenceNormal(temp, iv, "תיקון", "טעות", "סופר", "בשכירות"))
            {
                sss = temp[iv] + " " + temp[iv + 1] + " " + temp[iv + 2] + " " + temp[iv + 3];
                result.Add(sss, 4);
            }
            else if (ClassUtils.isMatchSequenceNormal(temp, iv, "העברת", "שכירות", "בירושה"))
            {
                sss = temp[iv] + " " + temp[iv + 1] + " " + temp[iv + 2] ;
                result.Add(sss, 3);
            }
            else if (ClassUtils.isMatchSequenceNormal(temp, iv, "העברת", "שכירות", "בצוואה"))
            {
                sss = temp[iv] + " " + temp[iv + 1] + " " + temp[iv + 2];
                result.Add(sss, 3);
            }
            else if (ClassUtils.isMatchSequenceNormal(temp, iv, "שינוי", "שם", "בחוכר"))
            {
                sss = temp[iv] + " " + temp[iv + 1] + " " + temp[iv + 2];
                result.Add(sss, 3);
            }
            else if (ClassUtils.isMatchSequenceNormal(temp, iv, "תיקון", "תנאים", "בשכירות"))
            {
                sss = temp[iv] + " " + temp[iv + 1] + " " + temp[iv + 2];
                result.Add(sss, 3);
            }
            else if (ClassUtils.isMatchSequenceNormal(temp, iv, "העברת", "שכירות"))
            {
                sss = temp[iv] + " " + temp[iv + 1] ;
                result.Add(sss, 2);
            }
            else if (ClassUtils.isMatchSequenceNormal(temp, iv, "שכירות") || ClassUtils.isMatchSequenceNormal(temp, iv, "חכירות"))
            {
                sss = temp[iv] ;
                result.Add(sss, 1);
            }
            return result;
        }
        public static Dictionary<string,int> GetOwnershipAction(List<string> rawValue, int iv, bool reverse)
        {
            Dictionary<string, int> result = new Dictionary<string, int>();
            List<string> temp = new List<string>(rawValue);
            if (reverse) temp = reverseOrder(temp);
            string sss;

            if (ClassUtils.isStringOneOfParams(temp[iv], "העברת", "שכירות", "ללא", "תמורה", "תיקון", "תנאים", "בשכירות", "בצוואה", "בירושה", "טעות", "סופר", "עודף", "חלקית", "בחוכר", "הסכם", "עפ\"י", "פיצול", "ירושה", "מכר", "רישום", "בעלות", "לאחר", "הסדר", "שנוי", "שם", "חלוקה", "איחוד", "רכישה", "לפי", "חק", "רכישת", "צוואה", "עדכון", "התאמת" ,"משותף", "צו", "בית", "על", "פי","שינוי","משפט"))
            {
                if (ClassUtils.isMatchSequenceNormal(temp, iv, "רישום", "לפי", "צו", "בית","משפט"))
                {
                    sss = temp[iv] + " " + temp[iv + 1] + " " + temp[iv + 2] + " " + temp[iv + 3] + " " + temp[iv+4]  ;
                    result.Add(sss, 5);
                }
                else if (ClassUtils.isMatchSequenceNormal(temp, iv, "העברת", "שכירות", "ללא", "תמורה"))
                {
                    sss = temp[iv] + " " + temp[iv + 1] + " " + temp[iv + 2] + " " + temp[iv + 3];
                    result.Add(sss, 4);
                }
                if (ClassUtils.isMatchSequenceNormal(temp, iv, "תיקון", "צו", "בית", "משותף"))
                {
                    sss = temp[iv] + " " + temp[iv + 1] + " " + temp[iv + 2] + " " + temp[iv + 3];
                    result.Add(sss, 4);
                }
                else if (ClassUtils.isMatchSequenceNormal(temp, iv, "תיקון", "טעות", "סופר", "בשכירות"))
                {
                    sss = temp[iv] + " " + temp[iv + 1] + " " + temp[iv + 2] + " " + temp[iv + 3];
                    result.Add(sss, 4);
                }
                else if (ClassUtils.isMatchSequenceNormal(temp, iv, "תיקון", "טעות", "סופר", "בחוכר"))
                {
                    sss = temp[iv] + " " + temp[iv + 1] + " " + temp[iv + 2] + " " + temp[iv + 3];
                    result.Add(sss, 4);
                 }
                else if (ClassUtils.isMatchSequenceNormal(temp, iv, "רישום", "בעלות", "לאחר", "הסדר"))
                {
                    sss = temp[iv] + " " + temp[iv + 1] + " " + temp[iv + 2] + " " + temp[iv + 3];
                    result.Add(sss, 4);
                }
                else if (ClassUtils.isMatchSequenceNormal(temp, iv, "רכישה", "לפי", "חק", "רכישת"))
                {
                    sss = temp[iv] + " " + temp[iv + 1] + " " + temp[iv + 2] + " " + temp[iv + 3];
                    result.Add(sss, 4);
                }
                else if (ClassUtils.isMatchSequenceNormal(temp, iv, "צוואה", "על", "פי", "הסכם"))
                {
                    sss = temp[iv] + " " + temp[iv + 1] + " " + temp[iv + 2] + " " + temp[iv + 3];
                    result.Add(sss, 4);
                }
                else if (ClassUtils.isMatchSequenceNormal(temp, iv, "ירושה", "על", "פי", "הסכם"))
                {
                    sss = temp[iv] + " " + temp[iv + 1] + " " + temp[iv + 2] + " " + temp[iv + 3];
                    result.Add(sss, 4);
                }

                else if (ClassUtils.isMatchSequenceNormal(temp, iv, "עדכון", "פרטי", "זיהוי"))
                {
                    sss = temp[iv] + " " + temp[iv + 1] + " " + temp[iv + 2];
                    result.Add(sss, 3);
                }
                else if (ClassUtils.isMatchSequenceNormal(temp, iv, "העברת", "שכירות", "בירושה"))
                {
                    sss = temp[iv] + " " + temp[iv + 1] + " " + temp[iv + 2];
                    result.Add(sss, 3);
                }
                else if (ClassUtils.isMatchSequenceNormal(temp, iv, "העברת", "שכירות", "בצוואה"))
                {
                    sss = temp[iv] + " " + temp[iv + 1] + " " + temp[iv + 2];
                    result.Add(sss, 3);
                }
                else if (ClassUtils.isMatchSequenceNormal(temp, iv, "רישום", "בית", "משותף"))
                {
                    sss = temp[iv] + " " + temp[iv + 1] + " " + temp[iv + 2];
                    result.Add(sss, 3);
                }
                else if (ClassUtils.isMatchSequenceNormal(temp, iv, "העברת", "שכירות", "חלקית"))
                {
                    sss = temp[iv] + " " + temp[iv + 1] + " " + temp[iv + 2];
                    result.Add(sss, 3);
                }
                else if (ClassUtils.isMatchSequenceNormal(temp, iv, "תיקון", "תנאים", "בשכירות"))
                {
                    sss = temp[iv] + " " + temp[iv + 1] + " " + temp[iv + 2];
                    result.Add(sss, 3);
                }
                else if (ClassUtils.isMatchSequenceNormal(temp, iv, "תיקון", "טעות", "סופר"))
                {
                    sss = temp[iv] + " " + temp[iv + 1] + " " + temp[iv + 2];
                    result.Add(sss, 3);
                }

                else if (ClassUtils.isMatchSequenceNormal(temp, iv, "מכר", "ללא", "תמורה"))
                {
                    sss = temp[iv] + " " + temp[iv + 1] + " " + temp[iv + 2];
                    result.Add(sss, 3);
                }
                else if (ClassUtils.isMatchSequenceNormal(temp, iv, "התאמת", "רישום"))
                {
                    sss = temp[iv] + " " + temp[iv + 1];
                    result.Add(sss, 2);
                }
                else if (ClassUtils.isMatchSequenceNormal(temp, iv, "תיקון", "רישום"))
                {
                    sss = temp[iv] + " " + temp[iv + 1];
                    result.Add(sss, 2);
                }
                else if (ClassUtils.isMatchSequenceNormal(temp, iv, "העברת", "שכירות"))
                {
                    sss = temp[iv] + " " + temp[iv + 1];
                    result.Add(sss, 2);
                }
                else if (ClassUtils.isMatchSequenceNormal(temp, iv, "שנוי", "שם"))
                {
                    sss = temp[iv] + " " + temp[iv + 1];
                    result.Add(sss, 2);
                }
                else if (ClassUtils.isMatchSequenceNormal(temp, iv, "שינוי", "שם"))
                {
                    sss = temp[iv] + " " + temp[iv + 1];
                    result.Add(sss, 2);
                }

                else if (ClassUtils.isMatchSequenceNormal(temp, iv, "צירוף", "חלקים"))
                {
                    sss = temp[iv] + " " + temp[iv + 1];
                    result.Add(sss, 2);
                }

                else if (ClassUtils.isMatchSequenceNormal(temp, iv, "שכירות"))
                {
                    sss = temp[iv];
                    result.Add(sss, 1);
                }
                else if (ClassUtils.isMatchSequenceNormal(temp, iv, "עודף"))
                {
                    sss = temp[iv];
                    result.Add(sss, 1);
                }
                else if (ClassUtils.isMatchSequenceNormal(temp, iv, "פיצול"))
                {
                    sss = temp[iv];
                    result.Add(sss, 1);
                }
                else if (ClassUtils.isMatchSequenceNormal(temp, iv, "ירושה"))
                {
                    sss = temp[iv];
                    result.Add(sss, 1);
                }
                else if (ClassUtils.isMatchSequenceNormal(temp, iv, "מכר"))
                {
                    sss = temp[iv];
                    result.Add(sss, 1);
                }
                else if (ClassUtils.isMatchSequenceNormal(temp, iv, "חלוקה"))
                {
                    sss = temp[iv];
                    result.Add(sss, 1);
                }
                else if (ClassUtils.isMatchSequenceNormal(temp, iv, "איחוד"))
                {
                    sss = temp[iv];
                    result.Add(sss, 1);
                }
                else if (ClassUtils.isMatchSequenceNormal(temp, iv, "צוואה"))
                {
                    sss = temp[iv];
                    result.Add(sss, 1);
                }
                else if (ClassUtils.isMatchSequenceNormal(temp, iv, "עדכון"))
                {
                    sss = temp[iv];
                    result.Add(sss, 1);
                }
            }
            return result;
        }

        public static bool isItARealNesach(List<string> aline, string key)
        {
            bool bret = false;
            bret = (ClassUtils.isArrayIncludString(aline, key) > -1) ;
            return bret;
        }

        public static bool verifyMailAddress(string mail)
        {
            bool bret = true;
            string apiKey = "c5ab5405b5f709990f8bbaaa0d4a94ddf27b6066c252368cc3d5ea917f91"; // Replace API_KEY with your API Key
            string emailToValidate = mail;
            string responseString = "";
            string apiURL = "https://api.quickemailverification.com/v1/verify?email=" + HttpUtility.UrlEncode(emailToValidate) + "&apikey=" + apiKey;
            try
            {
                HttpWebRequest request = (HttpWebRequest)WebRequest.Create(apiURL);
                request.Method = "GET";
                using (WebResponse response = request.GetResponse())
                {
                    using (StreamReader ostream = new StreamReader(response.GetResponseStream()))
                    {
                        responseString = ostream.ReadToEnd();
                        JObject json = JObject.Parse(responseString);
                        JToken value;
                        json.TryGetValue("result", out value);
                        if ( value.ToString() == "unknown")
                        {
                            bret = false;
                        }
                        else if ( value.ToString() == "invalid")
                        {
                            bret = false;
                        }
                    }
                }
            }
            catch(Exception ex)
            {

            }
            return bret;
        }
        public static readonly Random rand = new Random();

        public static Color GetRandomColour()
        {
            return Color.FromArgb(rand.Next(200,255), rand.Next(150,255), rand.Next(150,255));
        }
    }
}

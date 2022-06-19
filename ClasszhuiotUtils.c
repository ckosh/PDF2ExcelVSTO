using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using static iText.Kernel.Pdf.Colorspace.PdfSpecialCs;

namespace PDF2ExcelVsto
{
    static public class ClasszhuiotUtils
    {
        public static List<string> rawlineToKeyPropDescription(List<string> raw)
        {
            List<string> results = new List<string>();

            for (int i = raw.Count - 1; i >= 0; i--)
            {
                if (raw[i] == "רשויות")
                {
                    results.Add(raw[i]);
                }
                else if (raw[i] == "שטח")
                {
                    results.Add(raw[i] + " " + raw[i - 1]);
                    i--;
                }
                else if (raw[i] == "סוג")
                {
                    results.Add(raw[i] + " " + raw[i - 1]);
                    i--;
                }
                else if (raw[i] == "יעוד")
                {
                    results.Add(raw[i] + " " + raw[i - 1]);
                    i--;
                }
            }
            return results;
        }
        public static List<string> rawlineToKeyMortgage(List<string> raw)
        {
            List<string> results = new List<string>();
            for (int i = raw.Count - 1; i >= 0; i--)
            {
                if (raw[i] == "דרגה")
                {
                    results.Add(raw[i]);
                }
                else if (raw[i] == "סכום")
                {
                    results.Add(raw[i]);
                }
                else if (raw[i] == "בתנאי" && raw[i - 1] == "שטר" && raw[i - 2] == "מקורי")
                {
                    results.Add(raw[i] + " " + raw[i - 1] + " " + raw[i - 2]);
                    i--; i--;
                }
                else if (raw[i] == "החלק" && raw[i - 1] == "בנכס")
                {
                    results.Add(raw[i] + " " + raw[i - 1]);
                    i--;

                }
            }
            return results;
        }
        public static List<string> rawlineToKeystage(List<string> raw)
        {
            List<string> results = new List<string>();
            for (int i = raw.Count - 1; i >= 0; i--)
            {
                if (raw[i] == "דרגה")
                {
                    results.Add(raw[i]);
                    continue;
                }
                else if (raw[i] == "סכום")
                {
                    results.Add(raw[i]);
                    continue;
                }
                else if (raw[i] == "בתנאי" && raw[i - 1] == "שטר" && raw[i - 2] == "מקורי")
                {
                    results.Add(raw[i] + " " + raw[i - 1] + " " + raw[i - 2]);
                    i--;
                    i--;
                    continue;
                }
                else if (raw[i] == "החלק" && raw[i - 1] == "בנכס")
                {
                    results.Add(raw[i] + " " + raw[i - 1]);
                    i--;
                    continue;
                }
            }
            return results;
        }
        public static List<string> rawlineToKeyPropborrow(List<string> raw)
        {
            List<string> results = new List<string>();
            for (int i = raw.Count - 1; i >= 0; i--)
            {
                if (raw[i] == "שם" && raw[i - 1] == "הלווה")
                {
                    results.Add(raw[i] + " " + raw[i - 1]);
                    i--;
                }
                else if (raw[i] == "סוג" && raw[i - 1] == "זיהוי")
                {
                    results.Add(raw[i] + " " + raw[i - 1]);
                    i--;
                }
                else if (raw[i] == "מס'" && raw[i - 1] == "זיהוי")
                {
                    results.Add(raw[i] + " " + raw[i - 1]);
                    i--;
                }

            }
            return results;
        }
        public static List<string> rawlineToKeyPropOwners(List<string> raw)
        {
            List<string> results = new List<string>();
            for (int i = raw.Count - 1; i >= 0; i--)
            {
                if (raw[i] == "מס'" && raw[i - 1] == "שטר")
                {
                    results.Add(raw[i] + " " + raw[i - 1]);
                    i--;
                }
                else if (raw[i] == "תאריך")
                {
                    results.Add(raw[i]);
                }
                else if (raw[i] == "מהות" && raw[i - 1] == "פעולה")
                {
                    results.Add(raw[i] + " " + raw[i - 1]);
                    i--;
                }
                else if (raw[i] == "הבעלים")
                {
                    results.Add(raw[i]);
                }
                else if (raw[i] == "שם" && raw[i - 1] == "המוטב")
                {
                    results.Add(raw[i] + " " + raw[i - 1]);
                    i--;
                }
                else if (raw[i] == "שם" && raw[i - 1] == "החוכר")
                {
                    results.Add(raw[i] + " " + raw[i - 1]);
                    i--;
                }
                else if (raw[i] == "בעלי" && raw[i - 1] == "המשכנתה")
                {
                    results.Add(raw[i] + " " + raw[i - 1]);
                    i--;
                }
                else if (raw[i] == "סוג" && raw[i - 1] == "זיהוי")
                {
                    results.Add(raw[i] + " " + raw[i - 1]);
                    i--;
                }
                else if (raw[i] == "מס'" && raw[i - 1] == "זיהוי")
                {
                    results.Add(raw[i] + " " + raw[i - 1]);
                    i--;
                }
                else if (raw[i] == "החלק" && raw[i - 1] == "בנכס")
                {
                    results.Add(raw[i] + " " + raw[i - 1]);
                    i--;
                }
            }

            return results;
        }
        public static List<string> rawlineToValuesPropOwners0(List<string> rawKey0, List<string> rawValue0, ref string cont0)
        {
            List<string> results = new List<string>();
            List<string> rawKey = new List<string>(rawKey0);
            List<string> rawValue = new List<string>(rawValue0);
            rawValue.Reverse();
            int[] markKey = new int[rawKey.Count];
            int[] markValue = new int[rawValue.Count];
            string[] result = new string[rawKey.Count];
            int top = rawValue.Count - 1;
            for (int i = 0; i < rawKey.Count; i++)
            {
                markKey[i] = -1;
                result[i] = "";
            }
            for (int i = 0; i < rawValue.Count; i++) markValue[i] = -1;
            if (rawKey[0] == "מס' שטר")
            {
                if (ClassUtils.isShtarNumber(rawValue[0]))
                {
                    markKey[0] = 0;
                    markValue[0] = 0;
                    result[0] = rawValue[0];
                }
            }
            if (rawKey[1] == "תאריך")
            {
                if (ClassUtils.isDate(rawValue[1]))
                {
                    markKey[1] = 0;
                    markValue[1] = 0;
                    result[1] = rawValue[1];
                }
            }
            if (rawKey.Count > 5)
            {
                if (rawKey[5] == "מס' זיהוי")
                {
                    result[5] = rawValue[top];
                    markKey[5] = 0;
                    markValue[top] = 0;
                    top--;
                }
            }
            if (rawKey.Count > 4)
            {
                if (rawKey[4] == "סוג זיהוי")
                {
                    if (rawValue[top] == "חברה" || rawValue[top] == "ת.ז" || rawValue[top] == "דרכון" || rawValue[top] == "עמותה")
                    {
                        result[4] = rawValue[top];
                        markKey[4] = 0;
                        markValue[top] = 0;
                        top--;
                    }
                    else if (rawValue[top - 1] == "חברה" || rawValue[top - 1] == "ת.ז" || rawValue[top - 1] == "דרכון")
                    {
                        result[4] = rawValue[top - 1] + rawValue[top];
                        markKey[4] = 0;
                        markValue[top] = 0;
                        markValue[top - 1] = 0;
                        top--;
                        top--;
                    }
                    else
                    {
                        result[4] = "";
                        markKey[4] = 0;
                    }
                }
            }
            if (rawKey[2] == "מהות פעולה")
            {
                int lineSize;
                lineSize = analyzeActionOwner(rawValue, 3, ref cont0);
                for (int jj = 0; jj < lineSize; jj++)
                {
                    result[2] = result[2] + " " + rawValue[jj + 2];
                    markValue[jj + 2] = 0;
                }
                markKey[2] = 0;
            }// מהות פעולה 
            if (markKey.Sum() == -1) // last 
            {
                for (int ik = 0; ik < rawKey.Count; ik++)
                {
                    if (markKey[ik] == -1)
                    {
                        markKey[ik] = 0;
                        for (int iv = 0; iv < rawValue.Count; iv++)
                        {
                            if (markValue[iv] == -1)
                            {
                                result[ik] = result[ik] + rawValue[iv] + " ";
                            }
                        }
                    }
                }
            }
            for (int i = 0; i < result.Length; i++)
            {
                results.Add(result[i]);
            }
            return results;
        }

        public static List<string> rawlineToValuesMortgageboroer(List<string> rawKey0, List<string> rawValue0, ref string cont0, ref string cont1, ref bool checkList)
        {
            List<string> results = new List<string>();
            List<string> rawKey = new List<string>(rawKey0);
            List<string> rawValue = new List<string>(rawValue0);
            rawKey.Reverse();
            int[] markKey = new int[rawKey.Count];
            int[] markValue = new int[rawValue.Count];
            string[] result = new string[rawKey.Count];
            int top = rawValue.Count - 1;
            for (int i = 0; i < rawKey.Count; i++)
            {
                markKey[i] = -1;
                result[i] = "";
            }
            for (int i = 0; i < rawValue.Count; i++) markValue[i] = -1;
            if (rawKey[0] == "מס' זיהוי")
            {
                if (ClassUtils.IsIDNumber(rawValue[0]))
                {
                    markKey[0] = 0;
                    markValue[0] = 0;
                    result[0] = rawValue[0];
                }
            }
            if (rawKey[1] == "סוג זיהוי")
            {
                result[1] = rawValue[1];
                markKey[1] = 0;
                markValue[1] = 0;
            }
            for (int i = 2; i < markValue.Length; i++)
            {
                result[2] = rawValue[i] + " " + result[2];
            }
            for (int i = 0; i < result.Length; i++)
            {
                results.Add(result[i]);
            }

            return results;
        }


        public static List<string> rawlineToValuesMortgage(List<string> rawKey0, List<string> rawValue0, ref string cont0, ref string cont1, ref bool checkList)
        {
            List<string> results = new List<string>();
            List<string> rawKey = new List<string>(rawKey0);
            List<string> rawValue = new List<string>(rawValue0);
            rawValue.Reverse();
            int[] markKey = new int[rawKey.Count];
            int[] markValue = new int[rawValue.Count];
            string[] result = new string[rawKey.Count];
            int top = rawValue.Count - 1;
            for (int i = 0; i < rawKey.Count; i++)
            {
                markKey[i] = -1;
                result[i] = "";
            }
            for (int i = 0; i < rawValue.Count; i++) markValue[i] = -1;
            if (rawKey[0] == "מס' שטר")
            {
                if (ClassUtils.isShtarNumber(rawValue[0]))
                {
                    markKey[0] = 0;
                    markValue[0] = 0;
                    result[0] = rawValue[0];
                }
            }
            if (rawKey[1] == "תאריך")
            {
                if (ClassUtils.isDate(rawValue[1]))
                {
                    markKey[1] = 0;
                    markValue[1] = 0;
                    result[1] = rawValue[1];
                }
            }
            if (rawKey.Count > 5)
            {
                if (rawKey[5] == "מס' זיהוי")
                {
                    result[5] = rawValue[top];
                    markKey[5] = 0;
                    markValue[top] = 0;
                    top--;
                }
            }
            if (rawKey.Count > 4)
            {
                if (rawKey[4] == "סוג זיהוי")
                {
                    if (ClassUtils.isArrayIncludString(rawValue, "דרכון") > -1)
                    {
                        if (!ClassUtils.isIdType(rawValue[top]))
                        {
                            result[4] = rawValue[top];
                            markValue[top] = 0;
                            top--;
                        }
                    }
                    result[4] = rawValue[top] + " " + result[4];
                    markKey[4] = 0;
                    markValue[top] = 0;
                    top--;
                }
            }
            if (rawKey[2] == "מהות פעולה")
            {
                int lineSize;
                lineSize = analyzeActionMortgage(rawValue, 3, ref cont0, ref cont1, ref checkList);
                for (int jj = 0; jj < lineSize; jj++)
                {
                    result[2] = result[2] + " " + rawValue[jj + 2];
                    markValue[jj + 2] = 0;
                }
                markKey[2] = 0;
            }// מהות פעולה 
            if (markKey.Sum() == -1) // last 
            {
                for (int ik = 0; ik < rawKey.Count; ik++)
                {
                    if (markKey[ik] == -1)
                    {
                        markKey[ik] = 0;
                        for (int iv = 0; iv < rawValue.Count; iv++)
                        {
                            if (markValue[iv] == -1)
                            {
                                if (!ClassUtils.containHebrew(rawValue[iv]))
                                {
                                    result[ik] = rawValue[iv] + " " + result[ik];
                                }
                                else
                                {
                                    result[ik] = result[ik] + rawValue[iv] + " ";
                                }
                            }
                        }
                    }
                }
            }
            for (int i = 0; i < result.Length; i++)
            {
                results.Add(result[i]);
            }


            return results;
        }
        public static List<string> rawlineToValuesRemarks0(List<string> rawKey0, List<string> rawValue0, ref string cont0, ref string cont1, ref bool checkList)
        {
            List<string> results = new List<string>();
            List<string> rawKey = new List<string>(rawKey0);
            List<string> rawValue = new List<string>(rawValue0);

            //          rawKey.Reverse();
            rawValue.Reverse();
            if (rawValue[rawValue.Count - 1] == "X")
            {
                rawValue.RemoveAt(rawValue.Count - 1);
            }
            int[] markKey = new int[rawKey.Count];
            int[] markValue = new int[rawValue.Count];
            string[] result = new string[rawKey.Count];
            int top = rawValue.Count - 1;
            for (int i = 0; i < rawKey.Count; i++)
            {
                markKey[i] = -1;
                result[i] = "";
            }
            for (int i = 0; i < rawValue.Count; i++) markValue[i] = -1;

            if (rawKey[0] == "מס' שטר")
            {
                if (ClassUtils.isShtarNumber(rawValue[0]))
                {
                    markKey[0] = 0;
                    markValue[0] = 0;
                    result[0] = rawValue[0];
                }
            }
            if (rawKey[1] == "תאריך")
            {
                if (ClassUtils.isDate(rawValue[1]))
                {
                    markKey[1] = 0;
                    markValue[1] = 0;
                    result[1] = rawValue[1];
                }
            }
            if (rawKey.Count > 5)
            {
                if (rawKey[5] == "מס' זיהוי")
                {
                    result[5] = rawValue[top];
                    markKey[5] = 0;
                    markValue[top] = 0;
                    top--;
                }
            }
            if (rawKey.Count > 4)
            {
                if (rawKey[4] == "סוג זיהוי")
                {
                    if (ClassUtils.isArrayIncludString(rawValue, "דרכון") > -1)
                    {
                        if (!ClassUtils.isIdType(rawValue[top]))
                        {
                            result[4] = rawValue[top];
                            markValue[top] = 0;
                            top--;
                        }
                    }
                    result[4] = rawValue[top] + " " + result[4];
                    markKey[4] = 0;
                    markValue[top] = 0;
                    top--;
                }
            }
            if (rawKey[2] == "מהות פעולה")
            {
                int lineSize;
                lineSize = analyzeAction(rawValue, 3, ref cont0, ref cont1, ref checkList);
                for (int jj = 0; jj < lineSize; jj++)
                {
                    result[2] = result[2] + " " + rawValue[jj + 2];
                    markValue[jj + 2] = 0;
                }
                markKey[2] = 0;
            }// מהות פעולה 

            if (markKey.Sum() == -1) // last 
            {
                for (int ik = 0; ik < rawKey.Count; ik++)
                {
                    if (markKey[ik] == -1)
                    {
                        markKey[ik] = 0;
                        for (int iv = 0; iv < rawValue.Count; iv++)
                        {
                            if (markValue[iv] == -1)
                            {
                                if (!ClassUtils.containHebrew(rawValue[iv]))
                                {
                                    result[ik] = rawValue[iv] + " " + result[ik];
                                }
                                else
                                {
                                    result[ik] = result[ik] + rawValue[iv] + " ";
                                }
                            }
                        }
                    }
                }
            }
            for (int i = 0; i < result.Length; i++)
            {
                results.Add(result[i]);
            }
            return results;
        }
        public static List<string> rawlineToValuesLeasing0(List<string> rawKey0, List<string> rawValue)
        {
            List<string> results = new List<string>();
            List<string> rawKey = new List<string>(rawKey0);
            rawKey.Reverse();
            int[] markKey = new int[rawKey.Count];
            int[] markValue = new int[rawValue.Count];
            string[] result = new string[rawKey.Count];
            bool ZihuiDone = false;
            for (int i = 0; i < rawKey.Count; i++)
            {
                markKey[i] = -1;
                result[i] = "";
            }
            for (int i = 0; i < rawValue.Count; i++) markValue[i] = -1;

            do
            {
                for (int ik = rawKey.Count - 1; ik >= 0; ik--)
                {
                    if (markKey[ik] < 0)
                    {
                        for (int iv = rawValue.Count - 1; iv >= 0; iv--)
                        {
                            if (markValue[iv] < 0)
                            {
                                if (rawKey[ik] == "מס' שטר")
                                {
                                    if (ClassUtils.isShtarNumber(rawValue[iv]))
                                    {
                                        markKey[ik] = 0;
                                        markValue[iv] = 0;
                                        result[ik] = rawValue[iv];

                                        break;
                                    }
                                }
                                else if (rawKey[ik] == "תאריך")
                                {
                                    if (ClassUtils.isDate(rawValue[iv]))
                                    {
                                        markKey[ik] = 0;
                                        markValue[iv] = 0;
                                        result[ik] = rawValue[iv];

                                        break;
                                    }
                                }
                                else if (rawKey[ik] == "מהות פעולה")
                                {
                                    if (ClassUtils.isMatchSequence(rawValue, iv, "העברת", "שכירות", "בירושה", "עפ\"י", "הסכם"))
                                    {
                                        result[ik] = rawValue[iv] + " " + rawValue[iv - 1] + " " + rawValue[iv - 2] + " " + rawValue[iv - 3] + " " + rawValue[iv - 4];
                                        markKey[ik] = 0;
                                        markValue[iv] = 0;
                                        markValue[iv - 1] = 0;
                                        markValue[iv - 2] = 0;
                                        markValue[iv - 3] = 0;
                                        markValue[iv - 4] = 0;

                                        break;
                                    }

                                    if (ClassUtils.isStringOneOfParams(rawValue[iv], "העברת", "שכירות", "ללא", "תמורה", "תיקון", "תנאים", "בשכירות", "בצוואה", "בירושה", "טעות", "סופר", "עודף", "חלקית", "בחוכר", "הסכם", "עפ\"י", "פיצול", "ירושה", "מכר", "רישום", "בעלות", "לאחר", "הסדר", "שנוי", "שם", "חלוקה", "איחוד", "רכישה", "לפי", "חק", "רכישת", "צוואה", "עדכון", "התאמת","שינוי"))
                                    {
                                        if (ClassUtils.isMatchSequence(rawValue, iv, "העברת", "שכירות", "ללא", "תמורה"))
                                        {
                                            result[ik] = rawValue[iv] + " " + rawValue[iv - 1] + " " + rawValue[iv - 2] + " " + rawValue[iv - 3];
                                            markKey[ik] = 0;
                                            markValue[iv] = 0;
                                            markValue[iv - 1] = 0;
                                            markValue[iv - 2] = 0;
                                            markValue[iv - 3] = 0;

                                            break;
                                        }
                                        else if (ClassUtils.isMatchSequence(rawValue, iv, "תיקון", "טעות", "סופר", "בשכירות"))
                                        {
                                            result[ik] = rawValue[iv] + " " + rawValue[iv - 1] + " " + rawValue[iv - 2] + " " + rawValue[iv - 3];
                                            markKey[ik] = 0;
                                            markValue[iv] = 0;
                                            markValue[iv - 1] = 0;
                                            markValue[iv - 2] = 0;
                                            markValue[iv - 3] = 0;

                                            break;
                                        }
                                        else if (ClassUtils.isMatchSequence(rawValue, iv, "תיקון", "טעות", "סופר", "בחוכר"))
                                        {
                                            result[ik] = rawValue[iv] + " " + rawValue[iv - 1] + " " + rawValue[iv - 2] + " " + rawValue[iv - 3];
                                            markKey[ik] = 0;
                                            markValue[iv] = 0;
                                            markValue[iv - 1] = 0;
                                            markValue[iv - 2] = 0;
                                            markValue[iv - 3] = 0;

                                            break;
                                        }
                                        else if (ClassUtils.isMatchSequence(rawValue, iv, "רישום", "בעלות", "לאחר", "הסדר"))
                                        {
                                            result[ik] = rawValue[iv] + " " + rawValue[iv - 1] + " " + rawValue[iv - 2] + " " + rawValue[iv - 3];
                                            markKey[ik] = 0;
                                            markValue[iv] = 0;
                                            markValue[iv - 1] = 0;
                                            markValue[iv - 2] = 0;
                                            markValue[iv - 3] = 0;
                                            break;
                                        }
                                        else if (ClassUtils.isMatchSequence(rawValue, iv, "רכישה", "לפי", "חק", "רכישת"))
                                        {
                                            result[ik] = rawValue[iv] + " " + rawValue[iv - 1] + " " + rawValue[iv - 2] + " " + rawValue[iv - 3];
                                            markKey[ik] = 0;
                                            markValue[iv] = 0;
                                            markValue[iv - 1] = 0;
                                            markValue[iv - 2] = 0;
                                            markValue[iv - 3] = 0;
                                            break;
                                        }
                                        else if (ClassUtils.isMatchSequence(rawValue, iv, "העברת", "שכירות", "בירושה"))
                                        {
                                            result[ik] = rawValue[iv] + " " + rawValue[iv - 1] + " " + rawValue[iv - 2];
                                            markKey[ik] = 0;
                                            markValue[iv] = 0;
                                            markValue[iv - 1] = 0;
                                            markValue[iv - 2] = 0;

                                            break;
                                        }
                                        else if (ClassUtils.isMatchSequence(rawValue, iv, "העברת", "שכירות", "בצוואה"))
                                        {
                                            result[ik] = rawValue[iv] + " " + rawValue[iv - 1] + " " + rawValue[iv - 2];
                                            markKey[ik] = 0;
                                            markValue[iv] = 0;
                                            markValue[iv - 1] = 0;
                                            markValue[iv - 2] = 0;

                                            break;
                                        }
                                        else if (ClassUtils.isMatchSequence(rawValue, iv, "העברת", "שכירות", "חלקית"))
                                        {
                                            result[ik] = rawValue[iv] + " " + rawValue[iv - 1] + " " + rawValue[iv - 2];
                                            markKey[ik] = 0;
                                            markValue[iv] = 0;
                                            markValue[iv - 1] = 0;
                                            markValue[iv - 2] = 0;

                                            break;
                                        }
                                        else if (ClassUtils.isMatchSequence(rawValue, iv, "תיקון", "תנאים", "בשכירות"))
                                        {
                                            result[ik] = rawValue[iv] + " " + rawValue[iv - 1] + " " + rawValue[iv - 2];
                                            markKey[ik] = 0;
                                            markValue[iv] = 0;
                                            markValue[iv - 1] = 0;
                                            markValue[iv - 2] = 0;

                                            break;
                                        }
                                        else if (ClassUtils.isMatchSequence(rawValue, iv, "תיקון", "טעות", "סופר"))
                                        {
                                            result[ik] = rawValue[iv] + " " + rawValue[iv - 1] + " " + rawValue[iv - 2];
                                            markKey[ik] = 0;
                                            markValue[iv] = 0;
                                            markValue[iv - 1] = 0;
                                            markValue[iv - 2] = 0;

                                            break;
                                        }
                                        else if (ClassUtils.isMatchSequence(rawValue, iv, "מכר", "ללא", "תמורה"))
                                        {
                                            result[ik] = rawValue[iv] + " " + rawValue[iv - 1] + " " + rawValue[iv - 2];
                                            markKey[ik] = 0;
                                            markValue[iv] = 0;
                                            markValue[iv - 1] = 0;
                                            markValue[iv - 2] = 0;

                                            break;
                                        }
                                        else if (ClassUtils.isMatchSequence(rawValue, iv, "התאמת", "רישום"))
                                        {
                                            result[ik] = rawValue[iv] + " " + rawValue[iv - 1];
                                            markKey[ik] = 0;
                                            markValue[iv] = 0;
                                            markValue[iv - 1] = 0;

                                            break;
                                        }
                                        else if (ClassUtils.isMatchSequence(rawValue, iv, "תיקון", "רישום"))
                                        {
                                            result[ik] = rawValue[iv] + " " + rawValue[iv - 1];
                                            markKey[ik] = 0;
                                            markValue[iv] = 0;
                                            markValue[iv - 1] = 0;
                                            break;
                                        }
                                        else if (ClassUtils.isMatchSequence(rawValue, iv, "העברת", "שכירות"))
                                        {
                                            result[ik] = rawValue[iv] + " " + rawValue[iv - 1];
                                            markKey[ik] = 0;
                                            markValue[iv] = 0;
                                            markValue[iv - 1] = 0;

                                            break;
                                        }
                                        else if (ClassUtils.isMatchSequence(rawValue, iv, "שנוי", "שם"))
                                        {
                                            result[ik] = rawValue[iv] + " " + rawValue[iv - 1];
                                            markKey[ik] = 0;
                                            markValue[iv] = 0;
                                            markValue[iv - 1] = 0;
                                            break;
                                        }
                                        else if (ClassUtils.isMatchSequence(rawValue, iv, "שינוי", "שם"))
                                        {
                                        result[ik] = rawValue[iv] + " " + rawValue[iv - 1];
                                        markKey[ik] = 0;
                                        markValue[iv] = 0;
                                        markValue[iv - 1] = 0;
                                        break;
                                        }
                                        else if (ClassUtils.isMatchSequence(rawValue, iv, "צירוף", "חלקים"))
                                        {
                                            result[ik] = rawValue[iv] + " " + rawValue[iv - 1];
                                            markKey[ik] = 0;
                                            markValue[iv] = 0;
                                            markValue[iv - 1] = 0;
                                            break;
                                        }
                                        else if (ClassUtils.isMatchSequence(rawValue, iv, "שכירות"))
                                        {
                                            result[ik] = rawValue[iv];
                                            markKey[ik] = 0;
                                            markValue[iv] = 0;

                                            break;
                                        }
                                        else if (ClassUtils.isMatchSequence(rawValue, iv, "עודף"))
                                        {
                                            result[ik] = rawValue[iv];
                                            markKey[ik] = 0;
                                            markValue[iv] = 0;

                                            break;
                                        }
                                        else if (ClassUtils.isMatchSequence(rawValue, iv, "פיצול"))
                                        {
                                            result[ik] = rawValue[iv];
                                            markKey[ik] = 0;
                                            markValue[iv] = 0;

                                            break;
                                        }
                                        else if (ClassUtils.isMatchSequence(rawValue, iv, "ירושה"))
                                        {
                                            result[ik] = rawValue[iv];
                                            markKey[ik] = 0;
                                            markValue[iv] = 0;

                                            break;
                                        }
                                        else if (ClassUtils.isMatchSequence(rawValue, iv, "מכר"))
                                        {
                                            result[ik] = rawValue[iv];
                                            markKey[ik] = 0;
                                            markValue[iv] = 0;

                                            break;
                                        }
                                        else if (ClassUtils.isMatchSequence(rawValue, iv, "חלוקה"))
                                        {
                                            result[ik] = rawValue[iv];
                                            markKey[ik] = 0;
                                            markValue[iv] = 0;
                                            break;
                                        }
                                        else if (ClassUtils.isMatchSequence(rawValue, iv, "איחוד"))
                                        {
                                            result[ik] = rawValue[iv];
                                            markKey[ik] = 0;
                                            markValue[iv] = 0;
                                            break;
                                        }
                                        else if (ClassUtils.isMatchSequence(rawValue, iv, "צוואה"))
                                        {
                                            result[ik] = rawValue[iv];
                                            markKey[ik] = 0;
                                            markValue[iv] = 0;
                                            break;
                                        }
                                        else if (ClassUtils.isMatchSequence(rawValue, iv, "עדכון"))
                                        {
                                            result[ik] = rawValue[iv];
                                            markKey[ik] = 0;
                                            markValue[iv] = 0;
                                            break;
                                        }

                                    }
                                    return results;
                                }
                            }
                        }
                        // break;
                    }
                    continue;
                }
                for (int ik = 0; ik < rawKey.Count; ik++)
                {
                    if (ZihuiDone) break;
                    if (markKey[ik] < 0)
                    {
                        for (int iv = 0; iv < rawValue.Count; iv++)
                        {
                            if (markValue[iv] < 0)
                            {
                                if (rawKey[ik] == "מס' זיהוי")
                                {
                                    result[ik] = rawValue[iv];
                                    markKey[ik] = 0;
                                    markValue[iv] = 0;
                                    break;
                                }
                                else if (rawKey[ik] == "סוג זיהוי")
                                {
                                    if (ClassUtils.isStringOneOfParams(rawValue[iv], "ת.ז", "דרכון", "חברה"))
                                    {
                                        ZihuiDone = true;
                                        result[ik] = rawValue[iv];
                                        markKey[ik] = 0;
                                        markValue[iv] = 0;
                                        break;
                                    }
                                }
                            }
                        }
                    }
                }
                if (markKey.Sum() == -1) // last 
                {
                    for (int ik = 0; ik < rawKey.Count; ik++)
                    {
                        if (markKey[ik] == -1)
                        {
                            markKey[ik] = 0;
                            for (int iv = rawValue.Count - 1; iv >= 0; iv--)
                            {
                                if (markValue[iv] == -1)
                                {
                                    result[ik] = result[ik] + rawValue[iv] + " ";
                                }
                            }
                        }
                    }
                }
            } while (!markKey.AsQueryable().All(val => val > -1));

            for (int i = 0; i < result.Length; i++)
            {
                results.Add(result[i]);
            }
            return results;
        }
        public static List<string> rawlineToKeyPropLeasing(List<string> raw)
        {
            List<string> results = new List<string>();
            for (int i = raw.Count - 1; i >= 0; i--)
            {
                if (raw[i] == "רמת" && raw[i - 1] == "חכירה")
                {
                    results.Add(raw[i] + " " + raw[i - 1]);
                    i--;
                }
                else if (raw[i] == "בתנאי" && raw[i - 1] == "שטר" && raw[i - 2] == "מקורי")
                {
                    results.Add(raw[i] + " " + raw[i - 1] + " " + raw[i - 2]);
                    i--; i--;
                }
                else if (raw[i] == "תאריך" && raw[i - 1] == "סיום")
                {
                    results.Add(raw[i] + " " + raw[i - 1]);
                    i--;
                }
                else if (raw[i] == "תקופה" && raw[i - 1] == "בשנים")
                {
                    results.Add(raw[i] + " " + raw[i - 1]);
                    i--;
                }
                else if (raw[i] == "החלק" && raw[i - 1] == "בנכס")
                {
                    results.Add(raw[i] + " " + raw[i - 1]);
                    i--;
                }
            }
            return results;
        }
        public static List<string> rawlineToValuesLeasing1(List<string> raw)
        {
            List<string> results = new List<string>();
            for (int i = raw.Count - 1; i >= 0; i--)
            {
                if (ClassUtils.isShtarNumber(raw[i]))
                {
                    results.Add(raw[i]); // shtar
                }
                else if (ClassUtils.isDate(raw[i]))
                {
                    results.Add(raw[i]);
                }
                else
                {
                    results.Add(raw[i]);
                }
            }
            return results;
        }
        public static List<string> rawlineToValuesLeasing3(List<string> raw)
        {
            List<string> results = new List<string>();
            if (raw[0] == "בשלמות")
            {
                results.Add(raw[0]);
                results.Add(ClassUtils.buildCombinedLineSelected(raw, 1, raw.Count, true));
            }
            else
            {
                string ss = ClassUtils.buildCombinedLineSelected(raw, 0, 3, false);
                results.Add(ss);
                ss = ClassUtils.buildCombinedLineSelected(raw, 2, raw.Count, true);
                results.Add(ss);
            }
            return results;
        }
        public static int analyzeActionOwner(List<string> rawValue, int iv, ref string cont0)
        {
            int retVal = 0;
            cont0 = "";
            if (ClassUtils.isMatchSequenceStright(rawValue, iv, "העברת", "שכירות", "בירושה", "עפ\"י", "הסכם"))
            {
                retVal = 5;
            }
            else if (ClassUtils.isMatchSequenceStright(rawValue, iv, "מכר", "לפי", "צו", "בית", "משפט"))
            {
                retVal = 5;
            }
            else if (ClassUtils.isMatchSequenceStright(rawValue, iv, "מכר", "לפי", "סעיף", "5", "לחוק"))
            {
                retVal = 5;
                cont0 = "שיכונים ציבוריים";
            }
            else if (ClassUtils.isMatchSequenceStright(rawValue, iv, "העברת", "שכירות", "ללא", "תמורה"))
            {
                retVal = 4;
            }
            else if (ClassUtils.isMatchSequenceStright(rawValue, iv, "תיקון", "טעות", "סופר", "בשכירות"))
            {
                retVal = 4;
            }
            else if (ClassUtils.isMatchSequenceStright(rawValue, iv, "תיקון", "טעות", "סופר", "בחוכר"))
            {
                retVal = 4;
            }
            else if (ClassUtils.isMatchSequenceStright(rawValue, iv, "תיקון", "בעלות", "לאחר", "הסדר"))
            {
                retVal = 4;
            }
            else if (ClassUtils.isMatchSequenceStright(rawValue, iv, "רישום", "בעלות", "לאחר", "הסדר"))
            {
                retVal = 4;
            }
            else if (ClassUtils.isMatchSequenceStright(rawValue, iv, "רכישה", "לפי", "חק", "רכישת"))
            {
                retVal = 4;
                cont0 = "מקרקעין 1953";
            }
            else if (ClassUtils.isMatchSequenceStright(rawValue, iv, "הפקעה", "לפי", "סעיף", "19"))
            {
                retVal = 4;
            }
            else if (ClassUtils.isMatchSequenceStright(rawValue, iv, "העברת", "שכירות", "בירושה"))
            {
                retVal = 3;
            }
            else if (ClassUtils.isMatchSequenceStright(rawValue, iv, "העברת", "שכירות", "בצוואה"))
            {
                retVal = 3;
            }
            else if (ClassUtils.isMatchSequenceStright(rawValue, iv, "העברת", "שכירות", "חלקית"))
            {
                retVal = 3;
            }
            else if (ClassUtils.isMatchSequenceStright(rawValue, iv, "תיקון", "תנאים", "בשכירות"))
            {
                retVal = 3;
            }
            else if (ClassUtils.isMatchSequenceStright(rawValue, iv, "תיקון", "טעות", "סופר"))
            {
                retVal = 3;
            }
            else if (ClassUtils.isMatchSequenceStright(rawValue, iv, "מכר", "ללא", "תמורה"))
            {
                retVal = 3;
            }
            else if (ClassUtils.isMatchSequenceStright(rawValue, iv, "עדכון", "פרטי", "זיהוי"))
            {
                retVal = 3;
            }
            else if (ClassUtils.isMatchSequenceStright(rawValue, iv, "העברה", "מגוש", "לגוש"))
            {
                retVal = 3;
            }

            else if (ClassUtils.isMatchSequenceStright(rawValue, iv, "התאמת", "רישום"))
            {
                retVal = 2;
            }
            else if (ClassUtils.isMatchSequenceStright(rawValue, iv, "תיקון", "רישום"))
            {
                retVal = 2;
            }
            else if (ClassUtils.isMatchSequenceStright(rawValue, iv, "העברת", "שכירות"))
            {
                retVal = 2;
            }
            else if (ClassUtils.isMatchSequenceStright(rawValue, iv, "שנוי", "שם"))
            {
                retVal = 2;
            }
            else if (ClassUtils.isMatchSequenceStright(rawValue, iv, "שינוי", "שם"))
            {
                retVal = 2;
            }
            else if (ClassUtils.isMatchSequenceStright(rawValue, iv, "צירוף", "חלקים"))
            {
                retVal = 2;
            }
            else if (ClassUtils.isMatchSequenceStright(rawValue, iv, "שכירות"))
            {
                retVal = 1;
            }
            else if (ClassUtils.isMatchSequenceStright(rawValue, iv, "עודף"))
            {
                retVal = 1;
            }
            else if (ClassUtils.isMatchSequenceStright(rawValue, iv, "פיצול"))
            {
                retVal = 1;
            }
            else if (ClassUtils.isMatchSequenceStright(rawValue, iv, "ירושה"))
            {
                retVal = 1;
            }
            else if (ClassUtils.isMatchSequenceStright(rawValue, iv, "מכר"))
            {
                retVal = 1;
            }
            else if (ClassUtils.isMatchSequenceStright(rawValue, iv, "חלוקה"))
            {
                retVal = 1;
            }
            else if (ClassUtils.isMatchSequenceStright(rawValue, iv, "איחוד"))
            {
                retVal = 1;
            }
            else if (ClassUtils.isMatchSequenceStright(rawValue, iv, "צוואה"))
            {
                retVal = 1;
            }
            else if (ClassUtils.isMatchSequenceStright(rawValue, iv, "עדכון"))
            {
                retVal = 1;
            }
            return retVal;
        }
        public static int analyzeActionMortgage(List<string> rawValue, int iv, ref string cont, ref string cont1, ref bool checkList)
        {
            cont = "";
            cont1 = "";
            checkList = true;
            int retVal = 0;
            if (ClassUtils.isMatchSequenceStright(rawValue, iv, "פדיון", "משכנתה", "חלקי"))
            {
                retVal = 3;
            }
            else if (ClassUtils.isMatchSequenceStright(rawValue, iv, "תיקון", "משכנתה"))
            {
                retVal = 2;
            }
            else if (ClassUtils.isMatchSequenceStright(rawValue, iv, "משכנתה"))
            {
                retVal = 1;
            }

            return retVal;
        }
        public static int analyzeAction(List<string> rawValue, int iv, ref string cont, ref string cont1, ref bool checkList)
        {
            cont = "";
            cont1 = "";
            checkList = false;
            int retVal = 0;
            if (ClassUtils.isMatchSequenceStright(rawValue, iv, "הערה", "לפי", "סעיף", "11(א),", "(1)"))
            {
                retVal = 5;
                cont = "(2) לפקודת המיסים";
            }
            if (ClassUtils.isMatchSequenceStright(rawValue, iv, "הערה", "לפי", "סעיף", "11(א),", "(2)(1)"))
            {
                retVal = 5;
                cont = "(2) לפקודת המיסים";
            }

            else if (ClassUtils.isMatchSequenceStright(rawValue, iv, "הערה", "לפי", "סעיף", "11(א),", "12"))
            {
                retVal = 5;
                cont = "(2) לפקודת המיסים";
            }

            else if (ClassUtils.isMatchSequenceStright(rawValue, iv, "הערה", "על", "הפקעה", "סעיפים", "5"))
            {
                retVal = 5;
                cont = "ו- 7";
            }
            else if (ClassUtils.isMatchSequenceStright(rawValue, iv, "הערה", "על", "אי", "התאמה", "תקנה"))
            {
                retVal = 5;
                cont = "29";
            }
            else if (ClassUtils.isMatchSequenceStright(rawValue, iv, "הערה", "על", "מינוי", "כונס", "נכסים"))
            {
                retVal = 5;
            }
            else if (ClassUtils.isMatchSequenceStright(rawValue, iv, "הערה", "לפי", "צו", "בימ\"ש", "סעיף"))
            {
                retVal = 5;
                cont = "130";
            }
            else if (ClassUtils.isMatchSequenceStright(rawValue, iv, "הערה", "לפי", "סעיף", "11א(2)"))
            {
                retVal = 4;
                cont = "לפקודת המסים (גביה)";
            }

            else if (ClassUtils.isMatchSequenceStright(rawValue, iv, "הערה", "על", "הפקעת", "חלק"))
            {
                retVal = 4;
                cont = "מחלקה - סעיף 19";
            }
            else if (ClassUtils.isMatchSequenceStright(rawValue, iv, "הערת", "אזהרה", "סעיף", "126"))
            {
                retVal = 4;
                checkList = true;
            }
            else if (ClassUtils.isMatchSequenceStright(rawValue, iv, "תיקון", "הערת", "אזהרה", "סעיף"))
            {
                retVal = 4;
                cont = "126";
                checkList = true;
            }
            else if (ClassUtils.isMatchSequenceStright(rawValue, iv, "הערה", "על", "יעוד", "מקרקעין"))
            {
                retVal = 4;
                cont = "תקנה 27";
            }
            else if (ClassUtils.isMatchSequenceStright(rawValue, iv, "הערה", "על", "צורך", "בהסכמה"))
            {
                retVal = 4;
                cont = "סעיף 128";
            }
            else if (ClassUtils.isMatchSequenceStright(rawValue, iv, "הערה", "בדבר", "אתר", "עתיקות"))
            {
                retVal = 4;
            }
            else if (ClassUtils.isMatchSequenceStright(rawValue, iv, "הערה", "בדבר", "העברה", "לזרים"))
            {
                retVal = 4;
            }
            else if (ClassUtils.isMatchSequenceStright(rawValue, iv, "הערה", "על", "הפקדת", "תכנית"))
            {
                retVal = 4;
                cont = "סעיף 123 לחוק התכנון";
                cont1 = "והבניה";
            }
            else if (ClassUtils.isMatchSequenceStright(rawValue, iv, "הערה", "לפי", "פקודת", "הדרכים"))
            {
                retVal = 4;
            }
            else if (ClassUtils.isMatchSequenceStright(rawValue, iv, "הערה", "על", "ביטול", "הרשאה"))
            {
                retVal = 4;
                cont = "תקנה 26";
            }
            else if (ClassUtils.isMatchSequenceStright(rawValue, iv, "תיקונים", "שונים", "בהערה"))
            {
                retVal = 3;
            }
            else if (ClassUtils.isMatchSequenceStright(rawValue, iv, "תיקון", "טעות", "סופר"))
            {
                retVal = 3;
            }
            else if (ClassUtils.isMatchSequenceStright(rawValue, iv, "הערת", "אפוטרופוס", "(כללי,"))
            {
                retVal = 3;
                cont = "נפקדים)";
            }
            else if (ClassUtils.isMatchSequenceStright(rawValue, iv, "הסכם", "שיתוף"))
            {
                retVal = 2;
            }
            else if (ClassUtils.isMatchSequenceStright(rawValue, iv, "צו", "עקול"))
            {
                retVal = 2;
            }
            else if (ClassUtils.isMatchSequenceStright(rawValue, iv, "צו", "מניעה"))
            {
                retVal = 2;
            }
            else if (ClassUtils.isMatchSequenceStright(rawValue, iv, "הערה"))
            {
                retVal = 1;
            }
            return retVal;
        }
        public static int parseMortgageOwnerCont(List<string> rawKey, List<string> rawValue, string textToIgnore, ref string idNum, ref string idType, ref string Name)
        {
            List<string> nrawKey = new List<string>(rawKey);
            List<string> nrawValue = new List<string>(rawValue);
            nrawKey.Reverse();
            //            nrawValue.Reverse();

            int ret = 0;
            idNum = " ";
            idType = " ";
            Name = " ";
            int top = 0;
            if (textToIgnore != "")
            {

            }

            for (int i = 0; i < nrawKey.Count - 3; i++)
            {
                if (nrawKey[i] == "מס' זיהוי")
                {
                    if (ClassUtils.containHebrew(nrawValue[top])) continue;

                    if (ClassUtils.IsIDNumber(nrawValue[top]))
                    {
                        bool foreign = ClassUtils.isForeignID(nrawValue[top]);
                        if (foreign)
                        {
                            nrawValue[top] = ClassUtils.Reverse(nrawValue[top]);
                        }
                        idNum = nrawValue[top];
                        top++;
                        ret++;
                    }
                    continue;
                }
                if (nrawKey[i] == "סוג זיהוי" && idType == " ")
                {
                    if (ClassUtils.isArrayIncludString(nrawValue, "ת.ז") == -1 &&
                        ClassUtils.isArrayIncludString(nrawValue, "דרכון") == -1 &&
                        ClassUtils.isArrayIncludString(nrawValue, "חברה") == -1)
                    {
                        continue;
                    }

                    if (ClassUtils.isArrayIncludString(nrawValue, "דרכון") > -1)
                    {
                        if (!ClassUtils.isIdType(nrawValue[top]))
                        {
                            idType = nrawValue[top];
                            top++;
                        }
                    }
                    idType = nrawValue[top] + " " + idType;
                    top++;
                    ret++;
                    continue;
                }
                if (nrawKey[i] == "בעלי המשכנתה")
                {
                    for (int j = top; j < nrawValue.Count; j++)
                    {
                        if (!ClassUtils.containHebrew(nrawValue[j]))
                        {
                            Name = Name + " " + nrawValue[j];
                        }
                        else
                        {
                            Name = nrawValue[j] + " " + Name;
                        }
                    }
                    ret++;
                }
            }
            return ret;

        }
        public static int parse126Worning(List<string>rawKey, List<string> rawValue, string textToIgnore, ref string idNum, ref string idType, ref string Name)
        {
            List<string> nrawKey = new List<string>(rawKey);
            List<string> nrawValue = new List<string>(rawValue);
            nrawKey.Reverse();
            //            nrawValue.Reverse();

            int ret = 0;
            idNum = " ";
            idType = " ";
            Name = " ";
            int top = 0;
            if (textToIgnore != "")
            {

            }

            for (int i = 0; i < nrawKey.Count - 3; i++)
            {
                if (nrawKey[i] == "מס' זיהוי")
                {
                    if (ClassUtils.containHebrew(nrawValue[top])) continue;

                    if (ClassUtils.IsIDNumber(nrawValue[top]))
                    {
                        bool foreign = ClassUtils.isForeignID(nrawValue[top]);
                        if (foreign)
                        {
                            nrawValue[top] = ClassUtils.Reverse(nrawValue[top]);
                        }
                        idNum = nrawValue[top];
                        top++;
                        ret++;
                    }
                    continue;
                }
                if (nrawKey[i] == "סוג זיהוי" && idType == " ")
                {
                    if (ClassUtils.isArrayIncludString(nrawValue, "ת.ז") == -1 &&
                        ClassUtils.isArrayIncludString(nrawValue, "דרכון") == -1 &&
                        ClassUtils.isArrayIncludString(nrawValue, "חברה") == -1)
                    {
                        continue;
                    }

                    if (ClassUtils.isArrayIncludString(nrawValue, "דרכון") > -1)
                    {
                        if (!ClassUtils.isIdType(nrawValue[top]))
                        {
                            idType = nrawValue[top];
                            top++;
                        }
                    }
                    idType = nrawValue[top] + " " + idType;
                    top++;
                    ret++;
                    continue;
                }
                if (nrawKey[i] == "שם המוטב")
                {
                    for (int j = top; j < nrawValue.Count; j++)
                    {
                        if (!ClassUtils.containHebrew(nrawValue[j]))
                        {
                            Name = Name + " " + nrawValue[j];
                        }
                        else
                        {
                            Name = nrawValue[j] + " " + Name;
                        }
                    }
                    ret++;
                }
            }
            return ret;
        }
    }
}

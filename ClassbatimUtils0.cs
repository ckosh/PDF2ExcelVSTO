using System;
using System.Collections.Generic;
using System.Linq;

namespace PDF2ExcelVsto
{
    static public class ClassbatimUtils0
    {
        public static List<string> parseNozar(List<string> arr)
        {
            List<string> results = new List<string>();
            List<string> temp = new List<string>(arr);
            temp = ClassUtils.reverseOrder(temp);
            if (ClassUtils.isArrayIncludeAllStringsParam(temp, "הנכס", "נוצר", "ע\"י", "שטר:"))
            {
                results.Add(temp[4]);
            }
            if (ClassUtils.isArrayIncludeAllStringsParam(temp, "מיום:"))
            {
                results.Add(temp[6]);
            }
            if (ClassUtils.isArrayIncludeAllStringsParam(temp, "סוג", "שטר:"))
            {
                results.Add(ClassUtils.buildCombinedLineSelected(temp, 9, temp.Count, false));
            }
            return results;
        }
        public static List<string> ParseTopProperty(List<string> argKey)
        {
            List<string> results = new List<string>();
            List<string> temp = new List<string>(argKey);
            temp = ClassUtils.reverseOrder(temp);
            int top = 0;
            if (temp[top] == "רשויות")
            {
                results.Add("רשויות");
                top++;
            }
            if (temp[top] == "שטח" && temp[top+1] ==  "במ\"ר")
            {
                results.Add(temp[top]+ " " + temp[top+1]);
                top++; top++;
            }
            if (temp[top] == "תת" && temp[top + 1] == "חלקות")
            {
                results.Add(temp[top] + " " + temp[top + 1]);
                top++; top++;
            }
            if (temp[top] == "תקנון")
            {
                results.Add(temp[top]);
                top++; 
            }
            if (temp[top] == "שטר" && temp[top + 1] == "יוצר")
            {
                results.Add(temp[top] + " " + temp[top + 1]);
                top++; top++;
                if (top > temp.Count - 1) return results;
            }
            if (temp[top] == "תיק" && temp[top + 1] == "יוצר")
            {
                results.Add(temp[top] + " " + temp[top + 1]);
                top++; top++;
                if (top > temp.Count - 1) return results;
            }
            if (temp[top] == "תיק" && temp[top + 1] == "בית" && temp[top + 2] == "משותף")
            {
                results.Add(temp[top] + " " + temp[top + 1] + " " + temp[top + 2]);
                top++; top++; top++;
            }
            return results;
        }
        public static List<string> ParseSecondProperty(List<string> argKey , List<string> argVal)
        {
            List<string> results = new List<string>();
            List<string> temp = new List<string>(argVal);
            temp = ClassUtils.reverseOrder(temp);
            int top = 0;
            int point = 0;
            if ( argKey[point] == "רשויות")
            {
                string ss = "";
                while (ClassUtils.containHebrew(temp[top]))
                {
                    ss = ss + temp[top] + " ";
                    top++;
                }
                results.Add(ss);
                point++;
            }
            if (argKey[point] == "שטח במ\"ר")
            {
                results.Add(temp[top]);
                point++;
                top++;
            }
            if ( argKey[point] == "תת חלקות")
            {
                results.Add(temp[top]);
                point++;
                top++;
            }
            if (argKey[point] == "תקנון")
            {
                string ss = "";
                while (ClassUtils.containHebrew(temp[top]))
                {
                    ss = ss + temp[top] + " ";
                    top++;
                }
                results.Add(ss);
                point++;
            }
            if (argKey[point] == "שטר יוצר")
            {
                results.Add(temp[top]);
                point++;
                top++;
                if (point > argKey.Count - 1) return results;
            }
            if (argKey[point] == "תיק יוצר")
            {
                results.Add(temp[top]);
                point++;
                top++;
                if (point > argKey.Count - 1) return results;
            }
            if (argKey[point] == "תיק בית משותף")
            {
                results.Add(temp[top]);
                point++;
                top++;
            }
            return results;
        }
        public static List<string> ParseTopTatHelka(List<string> argKey)
        {
            List<string> results = new List<string>();
            List<string> temp = new List<string>(argKey);
            temp = ClassUtils.reverseOrder(temp);
            int top = 0;
            if (temp[top] == "שטח" && temp[top + 1] == "במ\"ר")
            {
                results.Add(temp[top] + " " + temp[top + 1]);
                top++; top++;
            }
            if (temp[top] == "תיאור" && temp[top + 1] == "קומה")
            {
                results.Add(temp[top] + " " + temp[top + 1]);
                top++; top++;
            }
            if (temp[top] == "כניסה" )
            {
                results.Add(temp[top] );
                top++; 
            }
            if (temp[top] == "אגף")
            {
                results.Add(temp[top]);
                top++;
            }
            if (temp[top] == "מבנה")
            {
                results.Add(temp[top]);
                top++;
            }
            if (temp[top] == "החלק" && temp[top + 1] == "ברכוש" && temp[top + 2] == "המשותף")
            {
                results.Add(temp[top] + " " + temp[top + 1] + " " + temp[top + 2]);
                top++; top++;
            }
            return results;
        }
        public static List<string> ParseSecondTatHelka(List<string> argKey, List<string> argVal)
        {
            List<string> results = new List<string>(new string[argKey.Count]);
            List<string> temp = new List<string>(argVal);
            temp = ClassUtils.reverseOrder(temp);
            int top = 0;
            int point = 0;
            if (argKey[argKey.Count - 1] == "החלק ברכוש המשותף")
            {
                results[argKey.Count-1] =  (temp[temp.Count-1]);
                //point++;
                //top++;
            }

            if (argKey[point] == "שטח במ\"ר")
            {
                results[0] = (temp[top]);
                point++;
                top++;
            }
            if (argKey[point] == "תיאור קומה")
            {
                if ( temp[top] == "קומה")
                {
                    results[point] = (temp[top] + " " + temp[top+1]);
                    point++;
                    top++;top++;
                }
                else
                {
                    results[point]= (temp[top]);
                    point++;
                    top++;
                }
            }
            if (argKey[point] == "כניסה")
            {
                results[point] = (temp[top]);
                point++;
                top++;
            }
            if (argKey[point] == "אגף")
            {
                results[point]= (temp[top]);
                point++;
                top++;
            }
            if (argKey[point] == "מבנה")
            {
                results[point] = (temp[top]);
                point++;
                top++;
            }
            return results;
        }
        public static List<string> parseLastLeasRow(List<string> argVal)
        {
            int j;
            List<int> keynum = new List<int>();
            for (int i = 0; i < argVal.Count; i++) keynum.Add(-1);
            List<string> results = new List<string>();
            for (int i = 0; i < 4; i++) results.Add(null);

            j = ClassbatimUtils0.findkeyinLeasing(argVal,"רמה:");
            if (j > 0)
            { 
                results[0] = argVal[j];
                keynum[j-1] = 1;
                keynum[j] = 1;
            }

            j = ClassbatimUtils0.findkeyinLeasing(argVal, "סיום:");
            if (j > 0)
            {
                keynum[j - 2] = 1;
                keynum[j - 1] = 1;
                keynum[j] = 1;
                results[1] = argVal[j];
            }
            j = ClassbatimUtils0.findkeyinLeasing(argVal, "בנכס:");
            if (j > 0)
            {
                keynum[j - 1] = 1;
                keynum[j - 2] = 1;
                for (int k = j ; k < argVal.Count; k++)
                {
                    results[3] = results[3] + " " + argVal[k];
                    keynum[k] = 1;
                }
            }
            for ( int k = 0; k < argVal.Count; k++)
            {
                if (keynum[k] < 0) results[2] = results[2] + " " + argVal[k];
            }
            return results;
        }
        public static List<string> parseMortgage(List<string> argVal, ref bool skipNextline)
        {
            List<string> results = new List<string>();
            for (int i = 0; i < 6; i++) results.Add(null);
            int[] markVal = new int[argVal.Count];
            int offset = 0;
            for (int i = 0; i < argVal.Count; i++) markVal[i] = -1;
             
            if (argVal[0] == "משכנתה")
            {
                results[0] = "משכנתה";
                markVal[0] = 1;
            }
            else if (argVal[0] == "שינוי" && argVal[1] == "בתנאי" && argVal[2] == "המשכנתה" )
            {
                results[0] = "שינוי בתנאי המשכנתה";
                markVal[0] = 1;
                markVal[1] = 1;
                markVal[2] = 1;
            }
            else if (argVal[0] == "תיקון" && argVal[1] == "טעות" && argVal[2] == "סופר" && argVal[3] == "במשכנתה")
            {
                results[0] = "תיקון טעות סופר במשכנתה";
                markVal[0] = 1;
                markVal[1] = 1;
                markVal[2] = 1;
                markVal[3] = 1;
            }
            if (ClassUtils.isShtarNumber(argVal[argVal.Count - 1]))
            {
                results[5] = argVal[argVal.Count - 1];
                markVal[argVal.Count - 1] = 1;
            }
            if (argVal[argVal.Count - 2] == "בשלמות")
            {
                results[4] = "בשלמות";
                markVal[argVal.Count - 2] = 1;
            }
            else if (argVal[argVal.Count - 3] == "חלק" && argVal[argVal.Count - 2] == "במקרקעין")
            {
                results[4] = "חלק במקרקעין";
                markVal[argVal.Count - 2] = 1;
                markVal[argVal.Count - 3] = 1;

                offset = 1;
            }
            else if ( int.TryParse(argVal[argVal.Count - 2], out _) && argVal[argVal.Count - 3] == "/" && int.TryParse(argVal[argVal.Count - 4], out _))
            {
                results[4] = argVal[argVal.Count - 2] + " / " + argVal[argVal.Count - 4];
                markVal[argVal.Count - 2] = 1;
                markVal[argVal.Count - 3] = 1;
                markVal[argVal.Count - 4] = 1;
                offset = 2;
            }
            if (ClassUtils.IsIDNumber( argVal[argVal.Count - 3 - offset] ))
            {
                results[3] = argVal[argVal.Count - 3 - offset];
                markVal[argVal.Count - 3 - offset] = 1;
            }
            if (ClassUtils.isIdType(argVal[argVal.Count - 4 - offset]))
            {
                results[2] = argVal[argVal.Count - 4 - offset];
                markVal[argVal.Count - 4 - offset] = 1;
            }
            string sss = "";
            for ( int i = 0; i < argVal.Count; i++)
            {
                if ( markVal[i] == -1)
                {
                    sss = sss + " " + argVal[i];
                }
            }
            results[1] = sss;
            return results;
        }
        public static List<string> parseLeaser(List<string> argVal, ref bool skipNextline )
        {
            List<string> results = new List<string>();
            for (int i = 0; i < 6; i++) results.Add(null);
            int[] markVal = new int[argVal.Count];
            for (int i = 0; i < argVal.Count; i++) markVal[i] = -1;
            int top = -1;
            Dictionary<string, int> ret;
            ret = ClassUtils.GetLeasingAction(argVal, 0, false);
            skipNextline = false;
            if ( ret.Count > 0)
            {
                results[0] = ret.ElementAt(0).Key;
                for (int i = 0; i < ret.ElementAt(0).Value; i++)
                {
                    markVal[i] = 1;
                }
                if ( ret.Count > 1)
                {
                    results[0] = results[0] + " " + ret.ElementAt(1).Key;
                    skipNextline = true;
                }
            }
            // shtar number
            if (ClassUtils.isShtarNumber(argVal[argVal.Count - 1]))
            {
                results[5] = argVal[argVal.Count - 1];
                markVal[argVal.Count - 1] = 1;
            }
            // part
            if (argVal[argVal.Count - 2] == "בשלמות")
            {
                results[4] = argVal[argVal.Count - 2];
                markVal[argVal.Count - 2] = 1;
                top = argVal.Count - 3;
            }
            else if (ClassUtils.isAllDigit(argVal[argVal.Count - 2]) && argVal[argVal.Count - 3] == "/" && ClassUtils.isAllDigit(argVal[argVal.Count - 4]))
            {
                results[4] = argVal[argVal.Count - 2] + " " + argVal[argVal.Count - 3] + " " + argVal[argVal.Count - 4];
                markVal[argVal.Count - 2] = 1;
                markVal[argVal.Count - 3] = 1;
                markVal[argVal.Count - 4] = 1;
                top = argVal.Count - 5;
            }
            // id type 
            int idtypelocation = findIDTypeLocation(argVal);
            if (idtypelocation > -1)
            {
                if (ClassUtils.containHebrew(argVal[idtypelocation + 1]))
                {
                    results[2] = argVal[idtypelocation] + " " + argVal[idtypelocation + 1];
                    markVal[idtypelocation] = 1;
                    markVal[idtypelocation + 1] = 1;
                    results[3] = argVal[idtypelocation + 2];
                    markVal[idtypelocation + 2] = 1;
                }
                else
                {
                    results[2] = argVal[idtypelocation];
                    markVal[idtypelocation] = 1;
                    results[3] = argVal[idtypelocation + 1];
                    markVal[idtypelocation + 1] = 1;
                }
            }
            string sss = "";
            for (int i = 0; i < markVal.Length; i++)
            {
                if (markVal[i] == -1) sss = sss + argVal[i] + " ";
            }
            results[1] = sss;


            return results;
        }
        public static int analyzeBatimAction(List<string> rawValue, int iv, ref string cont, ref string cont1, ref bool checkList)
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
            else if (ClassUtils.isMatchSequenceStright(rawValue, iv, "הערה", "על", "אי", "התאמה", "תקנה","29"))
            {
                retVal = 6;
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
            else if (ClassUtils.isMatchSequenceStright(rawValue, iv, "הערת", "אזהרה", "תמ\"א", "38"))
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
            else if (ClassUtils.isMatchSequenceStright(rawValue, iv, "צו", "ניהול", "ע\"י", "אפוטרופוס"))
            {
                retVal = 4;
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
            else if (ClassUtils.isMatchSequenceStright(rawValue, iv, "צו", "הריסה"))
            {
                retVal = 2;
            }

            else if (ClassUtils.isMatchSequenceStright(rawValue, iv, "הערה"))
            {
                retVal = 1;
            }
            return retVal;
        }
        public static List<string> ParseRemarkSecondLine(List<string> argVal)
        {
            List<string> results = new List<string>(6);
            for (int i = 0; i < 6; i++) results.Add(null);
            int[] markVal = new int[argVal.Count];
            for (int i = 0; i < argVal.Count; i++) markVal[i] = -1;

            if ( markVal[markVal.Length-1] == -1  && ClassUtils.IsIDNumber(argVal[argVal.Count - 1]))
            {
                results[3] = argVal[argVal.Count-1];
                markVal[markVal.Length-1] = 1;
                results[2] = argVal[argVal.Count - 2];
                markVal[markVal.Length - 2] = 1;
            }

            string ss = "";
            for (int k = 0; k < argVal.Count - 1; k++)
            {
                if (markVal[k] == -1) ss = ss + argVal[k] + " ";
            }
            results[1] = ss;

            return results;
        }
        public static List<string> ParseAttachmentsValues(List<string> l4,List<string>l3)        
        {
            List<string> results = new List<string>(5) { "", "", "", "", "" };
            List<int> l4int = new List<int>();
            for (int i = 0; i < l4.Count; i++) l4int.Add(-1);

            results[0] = l4[0]; //סימון בתשריט
            l4int[0] = 0;

            if ( l3[1] != "") // צבע בתשריט
            {
                for (int i = 1; i < 3; i++)
                {
                    if (ClassUtils.isStringOneOfParams(l4[i], "אפור", "סגול", "ורוד", "חום", "אדום", "כחול", "צהוב", "ירןק", "כתום", "תכלת", "זית", "חום"))
                    {
                        results[1] = results[1] + l4[i];
                        l4int[i] = 0;
                        continue;
                    }
                    else
                    {
                        break;
                    }
                }
            }

            if ( l3[4] != "") // שטח במ"ר
            {
                if ( ClassUtils.isFloating(l4[l4.Count - 1]))
                {
                    results[4] = l4[l4.Count - 1];
                    l4int[l4.Count - 1] = 0;
                }
            }
            if ( l3[3] != "")
            {
                for ( int j = l4int.Count-1; j > 0; j--)
                {
                    if (l4int[j] == -1 )
                    {
                        if (checkNumOrminusOrComma(l4[j]))
                        {
                            results[3] = l4[j];
                            l4int[j] = 0;
                            break;
                        }
                    }
                }
            }
            ///// fill extra charachters
            for ( int i = 0; i < l4int.Count; i++)
            {
                if (l4int[i] == -1)
                {
                    results[2] = results[2] + l4[i] + " ";
                    l4int[i] = 0;
                }
            }

            return results;
        }
        public static bool checkNumOrminusOrComma(string ss)
        {
            bool bret = false;
            if (ss.All(c => (c >= 48 && c <= 57 || c == 45  || c == 44 )))
            {
                bret = true;
            }
            return bret;
        }
        public static List<string> ParseAttachments(List<string> argVal)
        {
            List<string> results = new List<string>(5);
            for (int i = 0; i < 5; i++) results.Add(null);
            int[] markVal = new int[argVal.Count];
            for (int i = 0; i < argVal.Count; i++) markVal[i] = -1;
            for (int i = 0; i < argVal.Count; i++)
            {
                if (argVal[i] == "סימון" && argVal[i + 1] == "בתשריט")
                {
                    i++;
                    results[0] = "סימון בתשריט";
                }
                else if (argVal[i] == "צבע" && argVal[i + 1] == "בתשריט")
                {
                    i++;
                    results[1] = "צבע בתשריט";
                }
                else if (argVal[i] == "תיאור" && argVal[i + 1] == "הצמדה")
                {
                    i++;
                    results[2] = "תיאור הצמדה";
                }
                else if (argVal[i] == "משותפת" && argVal[i + 1] == "ל")
                {
                    i++;
                    results[3] = "משותפת ל";
                }
                else if (argVal[i] == "שטח" && argVal[i + 1] == "במ\"ר")
                {
                    i++;
                    results[4] = "שטח במ\"ר";
                }

            }

            return results;
        }
        public static List<string> ParseRemark(List<string> argVal, ref string conti)
        {
            List<string> results = new List<string>(6);
            for (int i = 0; i < 6; i++) results.Add(null);
            int[] markVal = new int[argVal.Count];
            for (int i = 0; i < argVal.Count; i++) markVal[i] = -1;
            string cont = "";
            string cont1 = "";
            bool checklist = false;
            if (ClassUtils.isShtarNumber(argVal[argVal.Count - 1]))
            {
                // set shtar value
                results[5] = argVal[argVal.Count - 1];
                markVal[argVal.Count - 1] = 1;

                int num = ClassbatimUtils0.analyzeBatimAction(argVal, 1, ref cont, ref cont1, ref checklist);
                if (num == 0)
                {
                    throw new Exception("הערה לא מזוהה");
                }
                conti = cont;
                // set remark type 
                string ss = "";
                for (int k = 0; k < num; k++)
                {
                    ss = ss + argVal[k] + " ";
                    markVal[k] = 1;
                }
                results[0] = ss + cont;
                // find id number and id type 
                for ( int k = 0; k < argVal.Count-1; k++)
                {
                    if (markVal[k] == -1 && ClassUtils.IsIDNumber(argVal[k]))
                    {
                        results[3] = argVal[k];
                        markVal[k] = 1;

                        results[2] = argVal[k - 1];
                        markVal[k - 1] = 1;

                        if (argVal[k + 1] == "בשלמות")
                        {
                            results[4] = "בשלמות";
                            markVal[k + 1] = 1;
                        }
                        break;
                    }
                    
                }
                // extract name
                ss = "";
                for ( int k = 0; k < argVal.Count-1; k++)
                {
                    if (markVal[k] == -1) ss = ss + argVal[k] + " ";
                }
                results[1] = ss;
            }
            return results;
        }
        public static List<string> parseOwners(List<string> argVal)
        {
            List<string> results = new List<string>(6);
            for (int i = 0; i < 6; i++) results.Add(null);

            int[] markVal = new int[argVal.Count];
            for (int i = 0; i < argVal.Count; i++) markVal[i] = -1;
            int top = -1;
            Dictionary<string, int> ret;
            
            // transaction type 
            ret = ClassUtils.GetOwnershipAction(argVal, 0, false);
            if ( ret.Count > 0 )
            {
                results[0] = ret.ElementAt(0).Key;
                for ( int i = 0; i < ret.ElementAt(0).Value; i++)
                {
                    markVal[i] = 1;
                }
            }

            // shtar number
            if (ClassUtils.isShtarNumber(argVal[argVal.Count-1]))
            {
                results[5] = argVal[argVal.Count - 1];
                markVal[argVal.Count - 1] = 1;
            }
            // part
            if (argVal[argVal.Count - 2] == "בשלמות")
            {
                results[4] = argVal[argVal.Count - 2];
                markVal[argVal.Count - 2] = 1;
                top = argVal.Count - 3;
            }
            else if (ClassUtils.isAllDigit(argVal[argVal.Count - 2]) && argVal[argVal.Count - 3] == "/" && ClassUtils.isAllDigit(argVal[argVal.Count - 4]))
            {
                results[4] = argVal[argVal.Count - 2] + " " + argVal[argVal.Count - 3] + " " + argVal[argVal.Count - 4];
                markVal[argVal.Count - 2] = 1;
                markVal[argVal.Count - 3] = 1;
                markVal[argVal.Count - 4] = 1;
                top = argVal.Count - 5;
            }
            // id type 
            int idtypelocation = findIDTypeLocation(argVal);
            if (idtypelocation > -1)
            {
                if (ClassUtils.containHebrew(argVal[idtypelocation + 1]))
                {
                    results[2] = argVal[idtypelocation] + " " + argVal[idtypelocation+1];
                    markVal[idtypelocation] = 1;
                    markVal[idtypelocation+1] = 1;
                    results[3] = argVal[idtypelocation + 2];
                    markVal[idtypelocation + 2] = 1;
                }
                else
                {
                    results[2] = argVal[idtypelocation];
                    markVal[idtypelocation] = 1;
                    results[3] = argVal[idtypelocation + 1];
                    markVal[idtypelocation + 1] = 1;
                }
            }

            //// id number
            //if ( ClassUtils.IsIDNumber(argVal[top]))
            //{
            //    results[3] = argVal[top];
            //    markVal[top] = 1;
            //    top--;
            //}
            //// idtype
            //if (ClassUtils.isIdType(argVal[top]))
            //{
            //    results[2] = argVal[top];
            //    markVal[top] = 1;
            //    top--;
            //}
            //Name
            string sss = "";
            for ( int i = 0; i < markVal.Length; i++)
            {
                if (markVal[i] == -1) sss = sss + argVal[i] + " ";
            }
            results[1] = sss;
                
             return results;
        }
        public static int findIDTypeLocation(List<string> argVal)
        {
            int ret = -1;
            for (int i = 0; i < argVal.Count - 1; i++)
            {
                if (argVal[i] == "ת.ז" || argVal[i] == "חברה" || argVal[i] == "דרכון" || argVal[i] == "עמותה")
                {
                    ret = i;
                    continue;
                }
            }
            return ret;
        }
        public static int findkeyinLeasing(List<string> argVal,string ss)
        {
            int ret = -1;
            for (int i = 0; i < argVal.Count - 1; i++)
            {
                if (argVal[i] == ss )
                {
                    ret = i+1;
                    break;
                }
            }

            return ret;
        }
    }
}

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using static PDF2ExcelVsto.ClassBatim;


namespace PDF2ExcelVsto
{
    class ClassMapBatim
    {
        List<List<string>> NesachTaboo;
        public ClassNozar nozar;
        public BatimCommonProperty property;
        List<TatHelka> tatHelkot = new List<TatHelka>();
        int endOfdata;

        public ClassMapBatim(List<List<string>> Nesach)
        {
            NesachTaboo = Nesach;
            MapMainSections();
            MapSubSections();
        }
        private void MapMainSections()
        {
            int next = 0;
            int row = next;
            while (!ClassUtils.isArrayIncludeAllStringsParam(NesachTaboo[row], "הרכוש", "המשותף"))
            {
                if (ClassUtils.isArrayIncludeAllStringsParam(NesachTaboo[row], "הנכס", "נוצר"))
                {
                    nozar = new ClassNozar();
                    nozar.line = row;
                    next = row;
                    break;

                }
                row++;
            }

            while (!ClassUtils.isArrayIncludeAllStringsParam(NesachTaboo[row], "תת", "חלקה"))
            {
                if (ClassUtils.isArrayIncludeAllStringsParam(NesachTaboo[row], "הרכוש", "המשותף"))
                {
                    property = new BatimCommonProperty();
                    property.line = row;
                    next = row;
                    break;
                }
                row++;
            }

            for (row = next + 1; row < NesachTaboo.Count; row++)
            {                
                while (!ClassUtils.isArrayIncludeAllStringsParam(NesachTaboo[row], "תת", "חלקה"))
                {
                    row++;
                    if (ClassUtils.isArrayIncludeAllStringsParam(NesachTaboo[row], "סוף", "נתונים"))
                    {
                        endOfdata = row;
                        return;
                    }
                    continue;                    
                }
                TatHelka tathelka = new TatHelka();
                tathelka.line = row;
                tatHelkot.Add(tathelka);
                continue;
            }
        }
        private void MapSubSections()
        {
            for ( int i = 0; i < tatHelkot.Count; i++)
            {
                TatHelka tat = tatHelkot[i];
                int lastline;
                if ( i == tatHelkot.Count-1)
                {
                    lastline = endOfdata;
                }
                else
                {
                    lastline = tatHelkot[i + 1].line;
                }
                for (int j = tatHelkot[i].line; j < lastline; j++)
                {
                    
                    if (ClassUtils.isArrayIncludeAllStringsParam(NesachTaboo[j], "הצמדות"))
                    {
                        Attachment att = new Attachment();
                        att.line = j;
                        
                        tat.attachments.Add(att);
                        continue;
                    }
                    else if (ClassUtils.isArrayIncludeAllStringsParam(NesachTaboo[j], "בעלויות"))
                    {
                        Owner own = new Owner();
                        own.line = j;
                        tat.owners.Add(own);
                        continue;
                    }
                    else if (ClassUtils.isArrayIsUniqueInLine(NesachTaboo[j], "משכנתאות"))
                    {
                        MortgageTatHelka mort = new MortgageTatHelka();
                        mort.line = j;
                        tat.mortgageTatHelkas.Add(mort);
                        continue;
                    }
                    else if (ClassUtils.isArrayIsUniqueInLine(NesachTaboo[j], "הערות"))
                    {
                        Remark rem = new Remark();
                        rem.line = j;
                        tat.remarks.Add(rem);
                        continue;
                    }
                }
            }

        }
    }
}

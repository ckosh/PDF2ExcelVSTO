using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PDF2ExcelVsto
{
    class ClassMapTaboo
    {
        int PropCreation;
        int PropDescription;
        int PropOwners;
        int Remarks;
        int Zikot;
        int Leasing;
        int Mortgage;
        int EndOdData;
        int Attachments;
        List<List<string>> NesachTaboo;
        Dictionary<string, int> tempMap = new Dictionary<string, int>();
        List<string> sections = new List<string>();
        List<int> sectionsRow = new List<int>();
        /// <summary>
        /// PropDescription

        List<string> description = new List<string>();
        List<int> descriptionRow = new List<int>();
        List<string> Owners = new List<string>();
        List<int> OwnersRow = new List<int>();
        List<string> Remark = new List<string>();
        List<int> RemarkRow = new List<int>();
        List<string> Leasingg = new List<string>();
        List<int> LeasingRow = new List<int>();
        List<string> MortgageRowg = new List<string>();
        List<int> MortgageRow = new List<int>();
        List<int> linesToDelete = new List<int>();
        public ClassMapTaboo(List<List<string>> Nesach)
        {
            PropCreation = -1;
            PropDescription = -1;
            PropOwners = -1;
            Remarks = -1;
            Zikot = -1;
            Leasing = -1;
            Mortgage = -1;
            EndOdData = -1;
            Attachments = -1;
            NesachTaboo = Nesach;
            MapMainSections();
            SortMap();
            linesToDelete = SortSubSections();
            if (linesToDelete.Count > 0)
            {
                linesToDelete.Sort();
                for ( int i = linesToDelete.Count -1; i >= 0; i--)
                {
                    NesachTaboo.RemoveAt(linesToDelete[i]);
                    // correct all sections accordinglly
                    shiftDown(descriptionRow, linesToDelete[i]);
                    shiftDown(OwnersRow, linesToDelete[i]);
                    shiftDown(RemarkRow, linesToDelete[i]);
                    shiftDown(LeasingRow, linesToDelete[i]);
                    shiftDown(MortgageRow, linesToDelete[i]);
                    shiftDown(sectionsRow, linesToDelete[i]);
                }
            }
        }

        private void MapMainSections()
        {
            for (int i = 0; i < NesachTaboo.Count; i++)
            {

                if (PropCreation == -1 && Attachments  == -1 && ClassUtils.isArrayIsUniqueInLine(NesachTaboo[i], "הנכס נוצר"))
                {
                    PropCreation = i;
                    tempMap.Add("PropCreation", PropCreation);
                    continue;
                }
                else if (PropDescription == -1 && Attachments == -1 && ClassUtils.isArrayIsUniqueInLine(NesachTaboo[i], "תיאור הנכס"))
                {
                    PropDescription = i;
                    tempMap.Add("PropDescription", PropDescription);
                    continue;
                }
                else if (PropOwners == -1 && Attachments == -1 && ClassUtils.isArrayIsUniqueInLine(NesachTaboo[i], "בעלויות"))
                {
                    PropOwners = i;
                    tempMap.Add("PropOwners", PropOwners);
                    continue;
                }
                else if (Remarks == -1 && Attachments == -1 && ClassUtils.isArrayIsUniqueInLine(NesachTaboo[i], "הערות"))
                {
                    Remarks = i;
                    tempMap.Add("Remarks", Remarks);
                    continue;
                }
                else if (Zikot == -1 && Attachments == -1 && ClassUtils.isArrayIsUniqueInLine(NesachTaboo[i], "זיקות הנאה"))
                {
                    Zikot = i;
                    tempMap.Add("Zikot", Zikot);
                    continue;
                }
                else if (Leasing == -1 && Attachments == -1 && ClassUtils.isArrayIsUniqueInLine(NesachTaboo[i], "חכירות"))
                {
                    Leasing = i;
                    tempMap.Add("Leasing", Leasing);
                    continue;
                }
                else if (Mortgage == -1 && Attachments == -1 && ClassUtils.isArrayIsUniqueInLine(NesachTaboo[i], "משכנתאות"))
                {
                    Mortgage = i;
                    tempMap.Add("Mortgage", Mortgage);
                    continue;
                }
                else if (EndOdData == -1 && ClassUtils.isArrayIsUniqueInLine(NesachTaboo[i], "סוף נתונים"))
                {
                    EndOdData = i;
                    tempMap.Add("EndOdData", EndOdData);
                    continue;
                }
                else if (Attachments == -1 && ClassUtils.isArrayIsUniqueInLine(NesachTaboo[i], "מחובר 1"))
                {
                    Attachments = i;
                    tempMap.Add("Attachments", Attachments);
                    continue;
                }
            }
        }
        private void SortMap()
        {
            foreach (KeyValuePair<string, int> map in tempMap.OrderBy(key => key.Value))
            {
                sections.Add(map.Key);
                sectionsRow.Add(map.Value);
            }
        }

        private List<int>  SortSubSections()
        {
            List<int> deleteLines = new List<int>();
            for (int i = 0; i < sections.Count - 1; i++)
            {
                if (sections[i] == "PropDescription")
                {
                    int startrow = sectionsRow[i] + 1;
                    int endRow = sectionsRow[i + 1] - 1;
                    for (int j = startrow; j < endRow; j++)
                    {
                        if (ClassUtils.isArrayIncludeAllStringsParam(NesachTaboo[j], "רשויות") || ClassUtils.isArrayIncludeAllStringsParam(NesachTaboo[j], "שטח"))
                        {
                            description.Add("רשויות");
                            descriptionRow.Add(j);
                        }
                        else if (ClassUtils.isArrayIncludeAllStringsParam(NesachTaboo[j], "המספרים"))
                        {
                            description.Add("המספרים הישנים של החלקה");
                            descriptionRow.Add(j);
                        }
                        else if (ClassUtils.isArrayIncludeAllStringsParam(NesachTaboo[j], "הערות"))
                        {
                            description.Add("הערות רשם המקרקעין");
                            descriptionRow.Add(j);
                        }
                        else if (ClassUtils.isArrayIncludeAllStringsParam(NesachTaboo[j], "מספר"))
                        {
                            description.Add("מספר המחוברים");
                            descriptionRow.Add(j);
                        }
                    }
                }
                else if (sections[i] == "PropOwners")
                {
                    int startrow = sectionsRow[i] + 1;
                    int endRow = sectionsRow[i + 1] - 1;
                    for (int j = startrow; j < endRow; j++)
                    {
                        if (ClassUtils.isArrayIncludeAllStringsParam(NesachTaboo[j], "מס'", "שטר"))
                        {
                            Owners.Add("מס' שטר");
                            OwnersRow.Add(j);
                        }
                    }
                }
                else if (sections[i] == "Remarks")
                {
                    int startrow = sectionsRow[i]+1;
                    int endRow = sectionsRow[i + 1]-1;
                    int lineToDelete;
                    bool withinremarks = false;
                    for (int j = startrow; j < endRow; j++)
                    {
                        if (ClassUtils.isArrayIncludeAllStringsParam(NesachTaboo[j], "מס'", "שטר") || 
                            ClassUtils.isArrayIncludeAllStringsParam(NesachTaboo[j], "סוף", "נתונים"))
                        {
                            if (!withinremarks)
                            {
                                withinremarks = true;
                                Remark.Add("מס' שטר");
                                RemarkRow.Add(j);
                            }
                            else
                            {
                                lineToDelete = j;
                                deleteLines.Add(lineToDelete);
                            }
                        }
                        else if (ClassUtils.isArrayIncludeAllStringsParam(NesachTaboo[j], "על", "הבעלות", "של") || 
                                 ClassUtils.isArrayIncludeAllStringsParam(NesachTaboo[j], "על", "כל", "הבעלים"))
                        {
                            withinremarks = false;
                        }
                    }
                    RemarkRow.Add(endRow+1);
                }
                else if (sections[i] == "Leasing")
                {
                    //
                    // line to delete was added to detect pass to a new page with a new header 
                    // from the top of the section.
                    //
                    int startrow = sectionsRow[i];
                    int endRow = sectionsRow[i + 1];
                    int lineToDelete;
                    bool withinLeasing = false;
                    for (int j = startrow; j < endRow; j++)
                    {                       
                        if (ClassUtils.isArrayIncludeAllStringsParam(NesachTaboo[j], "מס'", "שטר")  )
                        {
                            if (!withinLeasing)
                            {
                                withinLeasing = true;
                                Leasingg.Add("מס' שטר");
                                LeasingRow.Add(j);
                            }
                            else
                            {
                                lineToDelete = j;
                                deleteLines.Add(lineToDelete);
                            }
                        }
                        else if (ClassUtils.isArrayIncludeAllStringsParam(NesachTaboo[j], "על","כל","הבעלים"))
                        {
                            withinLeasing = false;
                        }
                    }
                    LeasingRow.Add(endRow);
                }
                else if (sections[i] == "Mortgage")
                {
                    int startrow = sectionsRow[i];
                    int endRow = sectionsRow[i + 1];
                    for (int j = startrow; j < endRow; j++)
                    {
                        if (ClassUtils.isArrayIncludeAllStringsParam(NesachTaboo[j], "מס'", "שטר"))
                        {
                            MortgageRowg.Add("מס' שטר");
                            MortgageRow.Add(j);
                        }
                    }
                    MortgageRow.Add(endRow);
                }
            }
            return deleteLines;
        }
        public bool isSectionExists(string section)
        {
            bool bret = false;
            for (int i = 0; i < sections.Count; i++)
            {
                if (sections[i] == section)
                {
                    bret = true;
                    break;
                }
            }
            return bret;
        }
        public List<int> getRowsofSection(string section)
        {
            List<int> rows = new List<int>();
            for (int i = 0; i < sections.Count; i++)
            {
                if (sections[i] == section)
                {
                    rows.Add(sectionsRow[i] + 1);
                    rows.Add(sectionsRow[i + 1] - 1);
                    break;
                }
            }
            return rows;
        }

        public List<int> getDescriptionRows()
        {
            return descriptionRow;
        }
        public List<int> GetOwnersRows()
        {
            return OwnersRow;
        }
        public List<int> GetLeasingRows()
        {
            return LeasingRow;
        }
        public List<int> GetMortGageRows()
        {
            return MortgageRow;
        }
        public List<int> GetRemarksRows()
        {
            return RemarkRow;
        }
        public void shiftDown(List<int> sec0, int jjj)
        {
            for (int j = 0; j < sec0.Count; j++)
            {
                if (sec0[j] > jjj)
                {
                    sec0[j] = sec0[j] - 1;
                }
            }

        }
    }
}

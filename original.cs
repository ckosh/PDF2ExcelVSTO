        public int buildHeader()
        {
            int retRow = 0;
            int i;
            List<string> dataRow = new List<string>();
            List<string> temp = new List<string>();


            header.dateCalendar = ClassUtils.buildCombinedline(slExcelData.DataRows[0]);
            header.dateHebrew = ClassUtils.buildCombinedline(slExcelData.DataRows[1]);
            header.time = ClassUtils.buildCombinedline(slExcelData.DataRows[2]);
            header.nesachNumber = ClassUtils.buildCombinedline(slExcelData.DataRows[3]);

            for (i = 4; i < slExcelData.DataRows.Count; i++)
            {
                dataRow = slExcelData.DataRows[i];
                if (ClassUtils.isArrayIncludString(dataRow, "הזכויות") > -1)
                {
                    tabooType = TabooType.Zehuiot;
                }
                else if (ClassUtils.isArrayIncludString(dataRow, "משותפים") > -1)
                {
                    tabooType = TabooType.MeshutafAll;
                }
                else if (ClassUtils.isArrayIncludString(dataRow, "גוש") > -1)
                {

                    if (dataRow.Count == 4) // Tat Chelka 
                    {
                        header.gush = dataRow[2];
                        header.tatHelka = null;
                        header.helka = dataRow[0];
                    }
                    else
                    {
                        //string[] helk = words[1].Split(' ');
                        //header.gush = helk[1];
                        //header.helka = words[2];
                        //header.tatHelka = "";
                    }
                    retRow = i + 1;
                    temp = (slExcelData.DataRows[retRow]);
                    if (ClassUtils.isArrayIncludeAllStringsParam(temp, "משותף", "עם", "חלקות"))
                    {
                        header.headerFoot = ClassUtils.buildCombinedline(temp);
                    }
                    else
                    {
                        header.headerFoot = ClassUtils.buildCombinedline(dataRow);
                    }
                    
                    header.tabooHeader.Add(header.headerFoot);
                    break;
                }
                header.tabooHeader.Add(ClassUtils.buildCombinedline(dataRow));
            }
            slExcelData.DataRows = ClassUtils.RemoveHeaderSection(slExcelData.DataRows, header.headerFoot, header.dateCalendar);

            return retRow;
        }

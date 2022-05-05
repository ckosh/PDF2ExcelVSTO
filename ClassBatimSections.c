using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PDF2ExcelVsto
{
    class ClassNozar
    {
        public string shtar;
        public string date;
        public string shtarType;

        public ClassNozar() 
        {
            shtar = "";
            date = "";
            shtarType = "";
        }
    }

    class CommonProperty 
    {
        public string rashuiot;
        public string areasqmr;
        public string numOfTatHelkot;
        public string takanon;
        public string shtarYozer;
        public string tikbaitMeshutaf;
        public string addtress;
        public List<string> remarks;

        public CommonProperty() 
        {
            rashuiot = "";
            areasqmr = "";
            numOfTatHelkot = "";
            takanon = "";
            shtarYozer = "";
            tikbaitMeshutaf = "";
            addtress = "";
            remarks = new List<string>();
        }
    }

    class TatHelka 
    {
        public int number;
        public string shetah;
        public string floor;
        public string partincommon;
        List<Owner> owners;
        List<Attachment> attachments;
        List<MortgageTatHelka> mortgageTatHelkas;

        public TatHelka() 
        {
            number = 0;
            shetah = "";
            floor = "";
            partincommon = "";
            owners = new List<Owner>();
            attachments = new List< Attachment>();
            mortgageTatHelkas = new List<MortgageTatHelka>();
        }
    };

    class Owner 
    {
        public string transaction;
        public string name;
        public string idType;
        public string idNumber;
        public string part;
        public string shtar;
        public Owner() 
        {
            transaction = "";
            name = "";
            idType = "";
            idNumber = "";
            part = "";
            shtar = "";
        }
    };

    class Attachment
    {
        public string mark;
        public string color;
        public string description;
        public string commonWith;
        public string area;
        public Attachment()
        {
            mark = "";
            color = "";
            description = "";
            commonWith = "";
            area = "";
        }
    };
    class MortgageTatHelka
    {
        public string type;
        public string Name;
        public string idType;
        public string idNumber;
        public string part;
        public string shtar;
        public string grade;

        public MortgageTatHelka()
        {
            type = "";
            Name = "";
            idType = "";
            idNumber = "";
            part = "";
            shtar = "";
            grade = "";
        }
    };
    class Remark
    {
        public string remarkType;
        public string name;
        public string idType;
        public string idNumber;
        public string shtar;
        public List<string> remarks;
        public Remark()
        {
            remarkType = "";
            name = "";
            idType = "";
            idNumber = "";
            shtar = "";
            remarks = new List<string>();
        }
    };
}

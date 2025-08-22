using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WizMes_WellMade
{
    class Win_dvl_MoldRegularInspect_U_CodeView : BaseView
    {
        public int Num { get; set; }
       
        public string MoldInspectID { get; set; } //금형점검번호
        public string MoldID { get; set; } //금형번호
        public string MoldInspectBasisID { get; set; } //기준번호 
        public string InspectCycle { get; set; } //점검주기 
        public string FileName { get; set; } 
        public string FilePath { get; set; } 
        public string MoldInspectDate { get; set; } //점검일자
        public string MoldInspectPersonID { get; set; } 
        public string MoldInspectPerson { get; set; } //점검자
        public string Comments { get; set; } //비고
        public string Article { get; set; } //비고

       
    }

    class Win_dvl_MoldRegularInspect_U_Sub_CodeView : BaseView
    {
        public override string ToString()
        {
            return (this.ReportAllProperties());
        }

        public string MoldInspectID { get; set; }
        public int Num { get; set; }
        public string MoldInspectBasisID { get; set; }
        public string MoldID { get; set; }
        public string MoldInspectBasisDate { get; set; }  
        public int MoldInspectSeq { get; set; }

        public string MoldInspectItemName { get; set; } 
        public string MoldInspectContent { get; set; }  
        public string MoldInspectCheckGbn { get; set; } 
        public string MoldInspectCheckName { get; set; }
        public string MoldInspectCycleGbn { get; set; } 
        public string MoldInspectCycleName { get; set; }
        public string MoldInspectCycleDate { get; set; }
        public string MoldInspectRecordGbn { get; set; }
        public string MoldInspectRecordName { get; set; }   
        public string MldInspectLegend { get; set; }   
        public double MldValue { get; set; }   
        public string Comments { get; set; }   


    }
}

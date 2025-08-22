namespace WizMes_WellMade
{
    class Win_Qul_InspectCode_U_Defect_CodeView : BaseView
    {
        public int Num { get; set; }

        public string DefectID { get; set; }
        public string Display1 { get; set; }
        public string Display2 { get; set; }
        public string Display3 { get; set; }
        public string KDefect { get; set; }

        public string Edefect { get; set; }
        public string TagName { get; set; }
        public string DefectClss { get; set; }
        public string DefectClssSub { get; set; }

        public string ButtonSeq { get; set; }
    }

    class Win_Qul_InspectCode_U_Basic_CodeView : BaseView
    {
        public int Num { get; set; }

        public string BasisID { get; set; }
        public string Basis { get; set; }
    }

    class Win_Qul_InspectCode_U_Grade_CodeView : BaseView
    {
        public int Num { get; set; }

        public string GradeID { get; set; }
        public string Grade { get; set; }
    }




    class Win_Qul_InspectCode_U_DefectProcess_All_CodeView : BaseView
    {
        public int AllProcessNum { get; set; }
        public int SelectProcessNum { get; set; }

        public bool chkFlag { get; set; }
        public string Process { get; set; }
        public string ProcessID { get; set; }
    }

    class Win_Qul_InspectCode_U_DefectProcess_Select_CodeView : BaseView
    {
        public int SelectProcessNum { get; set; }
        public int AllProcessNum { get; set; }

        public bool chkFlag { get; set; }
        public string Process { get; set; }
        public string ProcessID { get; set; }
    }

}

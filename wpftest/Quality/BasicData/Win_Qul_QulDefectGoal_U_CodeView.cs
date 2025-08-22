namespace WizMes_WellMade
{
    class Win_Qul_QulDefectGoal_U_CodeView : BaseView
    {
        public int Num { get; set; }

        public string YYYY { get; set; }
        public string DefectGoalAvg { get; set; }
    }

    class Win_Qul_QulDefectGoal_U_Sub_CodeView : BaseView
    {
        public int Num { get; set; }

        public string YYYY { get; set; }
        public string MM { get; set; }
        public string DefectGoal { get; set; }
        public string sortMM { get; set; }
        public string InspectGubun { get; set; }

        public string AvgDefectGoal { get; set; }
    }
}

namespace WizMes_WellMade
{
    class Win_dvl_InOutWare_Mold_U_CodeView : BaseView
    {
        public override string ToString()
        {
            return (this.ReportAllProperties());
        }

        public int Num { get; set; }
        public string InOutID { get; set; }
        public string InOutGbn { get; set; }
        public string InOutDate { get; set; }
        public string MoldID { get; set; }
        public string MoldNo { get; set; }
        public string InOutPlace { get; set; }
        public string InOutQty { get; set; }
        public string InOutPerson { get; set; }
        public string PersonName { get; set; }
        public string Comments { get; set; }
        public string MoldName { get; set; }
        public string Place { get; set; }
        public string InOutName { get; set; }
        public string ArticleSabun { get; set; }
        public string ArticleID { get; set; }
    }
}

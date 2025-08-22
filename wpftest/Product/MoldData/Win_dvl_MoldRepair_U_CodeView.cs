

namespace WizMes_WellMade
{
    class Win_dvl_MoldRepair_U_CodeView : BaseView
    {
        public override string ToString()
        {
            return (this.ReportAllProperties());
        }

        public string RepairID { get; set; }
        public string repairdate { get; set; }
        public string RepairGubun { get; set; }
        public string RepairGubunname { get; set; }
        public string MoldID { get; set; }
        public string RepairCustom { get; set; }
        public string Repairremark { get; set; }
        public string MoldKind { get; set; }
        public string MoldQuality { get; set; }
        public string Weight { get; set; }
        public string Spec { get; set; }
        public string MoldNo { get; set; }
        public string ProdCustomName { get; set; }
        public string Article { get; set; }
        public string ArticleID { get; set; }
        public string Article_Sabun { get; set; }
    }

    class Win_dvl_MoldRepair_U_Sub_CodeView :BaseView
    {
        public override string ToString()
        {
            return (this.ReportAllProperties());
        }

        public string RepairID { get; set; }
        public string RepairSubSeq { get; set; }
        public string McPartid { get; set; }
        public string customid { get; set; }
        public string partcnt { get; set; }
        public string partprice { get; set; }
        public string reason { get; set; }
        public string partremark { get; set; }
        public string MCPartName { get; set; }
        public string InCustomName { get; set; }

        public bool flagMcPart { get; set; }
        public bool flagCustom { get; set; }
        public bool flagPartcnt { get; set; }
        public bool flagPartprice { get; set; }
        public bool flagReason { get; set; }
        public bool flagPartremark { get; set; }
    }

    class GetAritcleByMoldID
    {
        public override string ToString()
        {
            return (this.ReportAllProperties());
        }

        public string dvlYNName { get; set; }
        public string dvlYN { get; set; }
        public string MoldQuality { get; set; }
        public string Weight { get; set; }
        public string Spec { get; set; }
        public string ProdCustomName { get; set; }
    }
}

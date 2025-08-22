using System.Collections.ObjectModel;
using System.Windows.Media.Imaging;

namespace WizMes_WellMade
{
    class Win_Qul_InspectAutoBasis_U_CodeView : BaseView
    {
        public int Num { get; set; }

        public string InspectBasisID { get; set; }
        public string Seq { get; set; }
        public string ArticleID { get; set; }
        public string Article { get; set; }
        public string EcoNo { get; set; }

        public string Comments { get; set; }
        public string CreateDate { get; set; }
        public string CreateUserID { get; set; }
        public string LastUpdateDate { get; set; }
        public string LastUpdateUserID { get; set; }

        public string BuyerArticleNo { get; set; }
        public string BuyerModelID { get; set; }
        public string Model { get; set; }
        public string InspectPoint { get; set; }
        public string MoldNo { get; set; }

        public string ProcessID { get; set; }
        public string Process { get; set; }
        public string InsCycleID { get; set; }
        public string InsCycle { get; set; }

        public string sketch1FilePath { get; set; }
        public string sketch1FileName { get; set; }
        public string sketch1FileAlias { get; set; }
    }

    class Win_Qul_InspectAutoBasis_U_Sub_CodeView : BaseView
    {
        public int Num { get; set; }

        public string InspectBasisID { get; set; }
        public string Seq { get; set; }
        public string SubSeq { get; set; }
        public string insType { get; set; }
        public string insItemName { get; set; }

        public string insRaSpec { get; set; }
        public string insRASpecMax { get; set; }
        public string InsRaSpecMin { get; set; }
        public string InsTPSpec { get; set; }
        public string InsTPSpecMax { get; set; }

        public string InsTPSpecMin { get; set; }
        public string InsSampleQty { get; set; }
        public string CreateDate { get; set; }
        public string CreateUserID { get; set; }
        public string LastUpdateDate { get; set; }

        public string LastUpdateUserID { get; set; }
        public string ManageGubun { get; set; }
        public string ManageGubunname { get; set; }

        public string InspectGageName { get; set; }
        public string InspectGage { get; set; }

        public string InspectCycleGubun { get; set; }

        public string InspectCycleGubunName { get; set; }
        public string InspectCycle { get; set; }
        public string Comments { get; set; }

        public string insTypeText { get; set; }
        public string Spec { get; set; }
        public string SpecMax { get; set; }
        public string SpecMin { get; set; }

        public string stringFlag { get; set; }

        public string InsImageFile { get; set; }
        public string InsImagePath { get; set; }

        public BitmapImage ImageView { get; set; }
        public bool imageFlag { get; set; }

        public ObservableCollection<CodeView> ovcType { get; set; }
        public ObservableCollection<CodeView> ovcManage { get; set; }
        public ObservableCollection<CodeView> ovcCycle { get; set; }
        public string ExcelCoordinates { get; set; }
    }
}

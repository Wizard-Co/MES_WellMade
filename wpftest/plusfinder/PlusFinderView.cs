using System.Collections.ObjectModel;

namespace WizMes_WellMade
{
    public class PlusFinderView : BaseView
    {
        public PlusFinderView()
        {
        }

        public ObservableCollection<CodeView> cboTrade { get; set; }

        public string m_sCodeField { get; set; }
        public string m_sNameField { get; set; }
        public string key { get; set; }
        public string value { get; set; }
    }
}

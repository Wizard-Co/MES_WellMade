using System.Collections.Generic;
using System.Data;
using System.Windows;
using System.Windows.Controls;

namespace WizMes_WellMade.PopUp
{
    /// <summary>
    /// MuniChoice.xaml에 대한 상호 작용 논리
    /// </summary>
    public partial class MuniChoice : Window
    {
        Lib lib = new Lib();

        string TargetDate;
        public bool MuniDataCountZero = false;
        public string SelectTextFileName = string.Empty;
        public string SelectM04PlusData = string.Empty;

        public MuniChoice()
        {
            InitializeComponent();
        }

        public MuniChoice(string targetdate)
        {
            TargetDate = targetdate;
            InitializeComponent();
        }

        private void MuniChoice_Loaded(object sender, RoutedEventArgs e)
        {
            Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
            sqlParameter.Clear();
            sqlParameter.Add("YYYYMMDD", TargetDate);

            DataSet ds = DataStore.Instance.ProcedureToDataSet("xp_Muni_sMuniDataTable", sqlParameter, false);

            if (ds != null && ds.Tables.Count > 0)
            {
                DataTable dt = null;
                dt = ds.Tables[0];

                if (dt.Rows.Count == 0)
                {
                    MuniDataCountZero = true;
                    DialogResult = DialogResult.HasValue;
                }
                else
                {
                    MuniDataCountZero = false;

                    dgdMuni.Items.Clear();
                    DataRowCollection drc = dt.Rows;
                    int num = 1;
                    foreach (DataRow item in drc)
                    {
                        var MuniChoice_View_Insert = new MuniChoice_View()
                        {
                            NUM = num.ToString(),

                            YYYYMMDD = item["YYYYMMDD"].ToString(),
                            TextFileName = item["TextFileName"].ToString(),
                            TestName = item["TestName"].ToString(),
                            TestNO = item["TestNO"].ToString(),
                            SampleName = item["SampleName"].ToString(),

                            Operator = item["Operator"].ToString(),
                            StartTime = item["StartTime"].ToString().Substring(8, 2) + "시" + item["StartTime"].ToString().Substring(10, 2) + "분",
                            EndTime = item["EndTime"].ToString().Substring(8, 2) + "시" + item["EndTime"].ToString().Substring(10, 2) + "분",
                            Temperature = item["Temperature"].ToString(),
                            InitialData = item["InitialData"].ToString(),

                            MinData = item["MinData"].ToString(),
                            ML01Plus4Data = item["ML01Plus4Data"].ToString(),
                            InterfaceYN = item["InterfaceYN"].ToString()
                        };
                        dgdMuni.Items.Add(MuniChoice_View_Insert);
                        num++;
                    }
                }
            }
        }


        // 적용버튼 클릭.
        private void btnConfirm_Click(object sender, RoutedEventArgs e)
        {
            int selectrow = dgdMuni.SelectedIndex;
            DataGridRow dgr = lib.GetRow(selectrow, dgdMuni);
            var ViewReceiver = dgr.Item as MuniChoice_View;

            SelectM04PlusData = ViewReceiver.ML01Plus4Data;
            SelectTextFileName = ViewReceiver.TextFileName;
            DialogResult = DialogResult.HasValue;
        }

        // 취소버튼 클릭.
        private void btnCancel_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }


    }


    public class MuniChoice_View
    {
        public string NUM { get; set; }

        public string YYYYMMDD { get; set; }
        public string TextFileName { get; set; }
        public string TestName { get; set; }
        public string TestNO { get; set; }
        public string SampleName { get; set; }

        public string Operator { get; set; }
        public string StartTime { get; set; }
        public string EndTime { get; set; }
        public string Temperature { get; set; }
        public string InitialData { get; set; }

        public string MinData { get; set; }
        public string ML01Plus4Data { get; set; }
        public string InterfaceYN { get; set; }

    }

}

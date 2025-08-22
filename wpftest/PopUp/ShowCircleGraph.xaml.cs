using LiveCharts;
using LiveCharts.Wpf;
using System;
using System.Collections.Generic;
using System.Data;
using System.Windows;

namespace WizMes_WellMade.PopUp
{
    /// <summary>
    /// ShowCircleGraph.xaml에 대한 상호 작용 논리
    /// </summary>
    public partial class ShowCircleGraph : Window
    {
        public ShowCircleGraph()
        {
            InitializeComponent();
        }

        DataTable GetDT;
        int TargetPoint;


        public ShowCircleGraph(DataTable dt, int sender)
        {
            InitializeComponent();

            GetDT = dt;
            TargetPoint = sender;
        }

        // 로드 시.
        private void ShowCircleGraph_Loaded(object sender, RoutedEventArgs e)
        {
            PieData pd = new PieData();

            DataRowCollection drc = GetDT.Rows;
            foreach (DataRow item in drc)
            {
                double value = 0;
                if ((TargetPoint == 1) || (TargetPoint == 2))
                {
                    value = Convert.ToDouble(item["DefectCountRate"].ToString());
                }
                else if (TargetPoint == 3)      //defect repair _ Total 부분
                {
                    value = Convert.ToDouble(item["RepairRate1"].ToString());
                }

                if (TargetPoint == 1)
                {
                    pd.AddSlice(item["DefectSymtom"].ToString(), value);
                }
                else if (TargetPoint == 2)
                {
                    pd.AddSlice(item["Article"].ToString(), value);
                }
                else if (TargetPoint == 3)      //defect repair _ total 부분
                {

                    pd.AddSlice(item["GroupingName"].ToString(), value);
                }
            }

            foreach (var n in pd.Slice)
            {
                PieChart.Series.Add(new PieSeries
                {
                    Title = n.Key,
                    Values = new ChartValues<double> { n.Value }
                });
            }

        }


        //닫기버튼
        private void btnClose_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }
    }




    public class PieData
    {
        private Dictionary<string, double> slice = new Dictionary<string, double>();

        public Dictionary<string, double> Slice
        {
            get { return slice; }
            set { slice = value; }
        }

        public void AddSlice(string slicename, double slicevalue)
        {
            try
            {
                slice.Add(slicename, slicevalue);
            }
            catch (Exception ex)
            {
                MessageBox.Show("오류 발생, 오류 내용 : " + ex.ToString());
            }
            finally
            {
                DataStore.Instance.CloseConnection();
            }
        }

    }


}

/*
 * 
 * @author 정승학
 * @remark 자동 온도 수집 자료 조회
 * @date 2021.01.25
 * @file Win_prd_TemperWorkLog_Q
 * @version
 * 
 * Create By Doxygen Form
 */


using LiveCharts;
using LiveCharts.Wpf;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Data;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using WizMes_WellMade.PopUp;
using WizMes_WellMade.PopUP;

namespace WizMes_WellMade
{
    /// <summary>
    /// Win_prd_TemperWorkLog_Q.xaml에 대한 상호 작용 논리
    /// </summary>
    public partial class Win_prd_TemperWorkLog_Q : UserControl
    {
        Lib lib = new Lib();
        PlusFinder pf = new PlusFinder();

        public Win_prd_TemperWorkLog_Q()
        {
            InitializeComponent();
        }

        private void UserControl_Loaded(object sender, RoutedEventArgs e)
        {
            lib.UiLoading(sender);
            SetComboBox();
            DatePickerStartDateSearch.SelectedDate = DateTime.Today;
            DatePickerEndDateSearch.SelectedDate = DateTime.Today;
            ComboBoxInspectResultSearch.SelectedIndex = 0;
        }

        #region 텍스트박스 입력방식 수정

        #endregion

        #region 콤보박스
        private void SetComboBox()
        {
            try
            {
                List<string[]> listPass = new List<string[]>();
                string[] strPass1 = { "Y", "합격" };
                string[] strPass2 = { "N", "불합격" };
                listPass.Add(strPass1);
                listPass.Add(strPass2);

                ObservableCollection<CodeView> ovcPass = ComboBoxUtil.Instance.Direct_SetComboBox(listPass);
                this.ComboBoxInspectResultSearch.ItemsSource = ovcPass;
                this.ComboBoxInspectResultSearch.DisplayMemberPath = "code_name";
                this.ComboBoxInspectResultSearch.SelectedValuePath = "code_id";

                ObservableCollection<CodeView> ovcLoc = ComboBoxUtil.Instance.Gf_DB_CM_GetComCodeDataset(null, "TMPRLOC", "Y", "", "");
                this.ComboBoxLocSearch.ItemsSource = ovcLoc;
                this.ComboBoxLocSearch.DisplayMemberPath = "code_name";
                this.ComboBoxLocSearch.SelectedValuePath = "code_id";
                this.ComboBoxLocSearch.SelectedIndex = 0;

            }
            catch(Exception ee)
            {
                MessageBox.Show("오류지점 - " + ee.ToString());
            }
        }
        #endregion

        #region 상단 레이아웃 조건 모음
        private void LabelDateSearch_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            try
            {
                if(CheckBoxDateSearch.IsChecked == true)
                {
                    CheckBoxDateSearch.IsChecked = false;
                }
                else
                {
                    CheckBoxDateSearch.IsChecked = true;
                }
            }
            catch(Exception ee)
            {
                MessageBox.Show("오류지점 - " + ee.ToString());
            }
        }

        private void CheckBoxDateSearch_Checked(object sender, RoutedEventArgs e)
        {
            try
            {
                if(DatePickerStartDateSearch != null )
                {
                    DatePickerStartDateSearch.IsEnabled = true;
                    DatePickerEndDateSearch.IsEnabled = true;
                }
            }
            catch (Exception ee)
            {
                MessageBox.Show("오류지점 - " + ee.ToString());
            }
        }

        private void CheckBoxDateSearch_Unchecked(object sender, RoutedEventArgs e)
        {
            try
            {
                DatePickerStartDateSearch.IsEnabled = false;
                DatePickerEndDateSearch.IsEnabled = false;
            }
            catch (Exception ee)
            {
                MessageBox.Show("오류지점 - " + ee.ToString());
            }
        }

        private void LabelInspectResultSearch_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            try
            {
                if(CheckBoxInspectResultSearch.IsChecked == true)
                {
                    CheckBoxInspectResultSearch.IsChecked = false;
                }
                else
                {
                    CheckBoxInspectResultSearch.IsChecked = true;
                }
            }
            catch (Exception ee)
            {
                MessageBox.Show("오류지점 - " + ee.ToString());
            }
        }

        private void CheckBoxInspectResultSearch_Checked(object sender, RoutedEventArgs e)
        {
            try
            {
                ComboBoxInspectResultSearch.IsEnabled = true;
            }
            catch (Exception ee)
            {
                MessageBox.Show("오류지점 - " + ee.ToString());
            }
        }

        private void CheckBoxInspectResultSearch_Unchecked(object sender, RoutedEventArgs e)
        {
            try
            {
                ComboBoxInspectResultSearch.IsEnabled = false;
            }
            catch (Exception ee)
            {
                MessageBox.Show("오류지점 - " + ee.ToString());
            }
        }
        #endregion

        #region 버튼 모음
        private void btnSearch_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                using(Loading lw = new Loading(beSearch))
                {
                    lw.ShowDialog();
                }
            }
            catch (Exception ee)
            {
                MessageBox.Show("오류지점 - " + ee.ToString());
            }
        }

        private void btnClose_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                lib.ChildMenuClose(this.ToString());
            }
            catch (Exception ee)
            {
                MessageBox.Show("오류지점 - " + ee.ToString());
            }
        }

        private void btnExcel_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                try
                {
                    DataTable dt = null;
                    string Name = string.Empty;

                    string[] lst = new string[2];
                    lst[0] = "온도 수집 자료";
                    lst[1] = DataGridMain.Name;

                    ExportExcelxaml ExpExc = new ExportExcelxaml(lst);

                    ExpExc.ShowDialog();

                    if (ExpExc.DialogResult.HasValue)
                    {
                        if (ExpExc.choice.Equals(DataGridMain.Name))
                        {
                            if (ExpExc.Check.Equals("Y"))
                                dt = Lib.Instance.DataGridToDTinHidden(DataGridMain);
                            else
                                dt = Lib.Instance.DataGirdToDataTable(DataGridMain);

                            Name = DataGridMain.Name;

                            if (Lib.Instance.GenerateExcel(dt, Name))
                                Lib.Instance.excel.Visible = true;
                            else
                                return;
                        }
                        else
                        {
                            if (dt != null)
                            {
                                dt.Clear();
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }
            catch (Exception ee)
            {
                MessageBox.Show("오류지점 - " + ee.ToString());
            }
        }
        #endregion

        #region 버튼 활성화(Enable) & 비활성화(Disable)

        #endregion

        #region 로드 딜레이타임 기능
        private void beSearch()
        {
            FillGrid();

            if(DataGridMain.Items.Count > 0)
            {
                FillChartGraph(DataGridMain);
                DataGridMain.SelectedIndex = 0;
            }
            else
            {
                if(lvcTotalChart != null && lvcTotalChart.Series != null)
                {
                    lvcTotalChart.Series.Clear();
                }

                MessageBox.Show("조회된 데이터가 없습니다.");
                return;
            }
        }
        #endregion

        #region 입력공간 이벤트 모음

        #endregion

        #region 입력공간 데이터 클리어

        #endregion

        #region 달력 및 콤보박스 등 기본 값 설정

        #endregion

        #region 데이터그리드 선택 변경

        #endregion

        #region Re_Search

        #endregion

        #region 조회
        private void FillGrid()
        {
            try
            {
                DataGridMain.Items.Clear();

                if(lvcTotalChart.Series != null)
                {
                    lvcTotalChart.Series.Clear();
                }

                Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                sqlParameter.Clear();
                sqlParameter.Add("nchkDate", CheckBoxDateSearch.IsChecked == true ? 1 : 0);
                sqlParameter.Add("sFromDate", DatePickerStartDateSearch.SelectedDate == null ? "" : DatePickerStartDateSearch.SelectedDate.Value.ToString("yyyyMMdd"));
                sqlParameter.Add("sToDate", DatePickerEndDateSearch.SelectedDate == null ? "" : DatePickerEndDateSearch.SelectedDate.Value.ToString("yyyyMMdd"));
                sqlParameter.Add("nchkLocID", CheckBoxLocSearch.IsChecked == true ? 1 : 0);
                sqlParameter.Add("sLocID", ComboBoxLocSearch.SelectedValue == null ? "" : ComboBoxLocSearch.SelectedValue);
                sqlParameter.Add("nchkPass", CheckBoxInspectResultSearch.IsChecked == true ? (ComboBoxInspectResultSearch.SelectedValue.ToString() == "Y" ? 1 : 2) : 0);

                DataSet ds = DataStore.Instance.ProcedureToDataSet("xp_WizWork_sTmprGathering_WPF", sqlParameter, false);

                if(ds != null && ds.Tables.Count > 0)
                {
                    DataTable dt = ds.Tables[0];

                    if(dt.Rows.Count > 0)
                    {
                        int i = 0;

                        DataRowCollection drc = dt.Rows;

                        foreach(DataRow dr in drc)
                        {
                            i++;

                            var WPTQC = new Win_prd_TemperWorkLog_Q_CodeView()
                            {
                                Num = i,

                                WorkDate = dr["WorkDate"].ToString(),
                                WorkTime = dr["WorkTime"].ToString(),
                                LOCID = dr["LOCID"].ToString(),
                                LOCName = dr["LOCName"].ToString(),
                                MinTemp = dr["MinTemp"].ToString(),
                                MaxTemp = dr["MaxTemp"].ToString(),
                                Temper = dr["Temper"].ToString(),
                                Humi = dr["Humi"].ToString(),
                                Comments = dr["Comments"].ToString(),
                                SaveOK = dr["SaveOK"].ToString(),
                                MinTempError = dr["MinTempError"].ToString(),
                                MaxTempError = dr["MaxTempError"].ToString(),

                                //tmprInspectBasisID = dr["tmprInspectBasisID"].ToString(),
                                //tmprBasisDate = dr["tmprBasisDate"].ToString(),
                                //LocID = dr["LocID"].ToString(),
                                //LocName = dr["LocName"].ToString(),
                                //tmprInsCycleGbn = dr["tmprInsCycleGbn"].ToString(),
                                //tmprInsCycleGbnName = dr["tmprInsCycleGbnName"].ToString(),
                                //tmprInsCondition = dr["tmprInsCondition"].ToString(),
                                //tmprInsCheckGbn = dr["tmprInsCheckGbn"].ToString(),
                                //tmprInsCheckGbnName = dr["tmprInsCheckGbnName"].ToString(),
                                //tmprInsSpec = dr["tmprInsSpec"].ToString(),
                                //tmprInsMin = dr["tmprInsMin"].ToString(),
                                //tmprInsMax = dr["tmprInsMax"].ToString(),
                                //tmprInspectID = dr["tmprInspectID"].ToString(),
                                //InspectDate = dr["InspectDate"].ToString(),
                                //InspectTime = dr["InspectTime"].ToString(),
                                //tmprInspectValue = dr["tmprInspectValue"].ToString(),
                                //Comments = dr["Comments"].ToString(),
                            };

                            if(WPTQC.MinTempError == "True" && WPTQC.MaxTempError == "True")
                            {
                                WPTQC.TempErrorColor = "True";
                            }
                            else
                            {
                                WPTQC.TempErrorColor = "False";
                            }

                            DataGridMain.Items.Add(WPTQC);
                        }
                    }
                }

            }
            catch(Exception ee)
            {
                MessageBox.Show("오류지점 - " + ee.ToString());
            }
            finally
            {
                DataStore.Instance.CloseConnection();
            }
        }
        #endregion

        #region 그래프 조회
        private void FillChartGraph(DataGrid dataGrid)
        {
            try
            {
                if(lvcTotalChart.Series != null)
                {
                    lvcTotalChart.Series.Clear();
                }

                ChartInfoTemperRate chartInfoTemp = new ChartInfoTemperRate();
                chartInfoTemp.seriesCollection = new SeriesCollection();
                chartInfoTemp.chartRunningRate = new ChartValues<double>();
                chartInfoTemp.chartGoalRate = new ChartValues<double>();
                chartInfoTemp.Labels = new string[dataGrid.Items.Count];

                chartInfoTemp.chartMinValue = new ChartValues<double>();
                chartInfoTemp.chartMaxValue = new ChartValues<double>();
                

                int index = 0;

                for(int i = 0; i < dataGrid.Items.Count; i++)
                {
                    var Rating = dataGrid.Items[i] as Win_prd_TemperWorkLog_Q_CodeView;

                    //chartInfoTemp.chartMinValue = new ChartValues<double> { ConvertDouble(Rating.MinTemp) };
                    //chartInfoTemp.chartMaxValue = new ChartValues<double> { ConvertDouble(Rating.MaxTemp) };

                    if (Rating != null)
                    {
                        chartInfoTemp.Labels[index] = Rating.WorkTime;

                        index++;

                        if(Rating.Temper != null && CheckConvertDouble(Rating.Temper))
                        {
                            chartInfoTemp.chartRunningRate.Add(ConvertDouble(Rating.Temper));
                            
                        }
                        else
                        {
                            chartInfoTemp.chartRunningRate.Add(0);
                        }

                        if (Rating.MinTemp != null & CheckConvertDouble(Rating.MinTemp))
                        {
                            chartInfoTemp.chartMinValue.Add(ConvertDouble(Rating.MinTemp));
                            chartInfoTemp.chartMaxValue.Add(ConvertDouble(Rating.MaxTemp));
                        }
                        else
                        {
                            chartInfoTemp.chartMinValue.Add(0);
                            chartInfoTemp.chartMaxValue.Add(0);
                        }
                    }
                }

                //SeriesCollection sc = new SeriesCollection
                //{
                //    new LineSeries
                //    {
                //        Values = chartInfoTemp.chartRunningRate,
                //        Title = "현재온도"
                //    },
                //    new LineSeries
                //    {
                //        Values = new ChartValues<double> { 29 },
                //        Title = "최대"
                //    }
                //};


                chartInfoTemp.seriesCollection.Add(new LineSeries
                {
                    Stroke = new SolidColorBrush(Colors.Blue),
                    //Fill = new SolidColorBrush(Colors.White),
                    Values = chartInfoTemp.chartRunningRate,
                    PointGeometry = null,
                    Title = "현재온도"
                });

                chartInfoTemp.seriesCollection.Add(new LineSeries
                {
                    Stroke = new SolidColorBrush(Colors.Green),
                    //Fill = new SolidColorBrush(Colors.Green),
                    Values = chartInfoTemp.chartMaxValue,
                    PointGeometry = null,
                    ToolTip = null,
                    
                    Title = "최대"
                });

                chartInfoTemp.seriesCollection.Add(new LineSeries
                {
                    Stroke = new SolidColorBrush(Colors.Red),
                    //Fill = new SolidColorBrush(Colors.red),
                    Values = chartInfoTemp.chartMinValue,
                    PointGeometry = null,
                    Title = "최소"
                });

                chartInfoTemp.Formatter = value => value + "( ℃)";
                this.DataContext = chartInfoTemp;

            }
            catch(Exception ee)
            {
                MessageBox.Show("오류지점 - " + ee.ToString());
            }
        }

        #endregion

        #region 저장

        #endregion

        #region 삭제

        #endregion

        #region 데이터체크

        #endregion

        #region 데이터그리드 입력 수정 방향키 포커스

        #endregion

        #region etc
        // 소수로 변환
        private double ConvertDouble(string str)
        {
            double result = 0;
            double chkDouble = 0;

            if (!str.Trim().Equals(""))
            {
                str = str.Replace(",", "");

                if (Double.TryParse(str, out chkDouble) == true)
                {
                    result = Double.Parse(str);
                }
            }

            return result;
        }

        // 소수로 변환 가능한지 체크 이벤트
        private bool CheckConvertDouble(string str)
        {
            bool flag = false;
            double chkDouble = 0;

            if (!str.Trim().Equals(""))
            {
                if (Double.TryParse(str, out chkDouble) == true)
                {
                    flag = true;
                }
            }

            return flag;
        }



        #endregion

        private void LabelLocSearch_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {

        }

        private void CheckBoxLocSearch_Checked(object sender, RoutedEventArgs e)
        {

        }

        private void CheckBoxLocSearch_Unchecked(object sender, RoutedEventArgs e)
        {

        }
    }

    class Win_prd_TemperWorkLog_Q_CodeView : BaseView
    {
        public override string ToString()
        {
            return (this.ReportAllProperties());
        }

        public int Num { get; set; }
        
        public string WorkDate { get; set; }
        public string WorkTime { get; set; }
        public string LOCID { get; set; }
        public string LOCName { get; set; }
        public string MinTemp { get; set; }
        public string MaxTemp { get; set; }
        public string Temper { get; set; }
        public string Humi { get; set; }
        public string Comments { get; set; }
        public string SaveOK { get; set; }
        public string MinTempError { get; set; }
        public string MaxTempError { get; set; }
        public string TempErrorColor { get; set; }


        //public string tmprInspectBasisID { get; set; }
        //public string tmprBasisDate { get; set; }
        //public string LocID { get; set; }
        //public string LocName { get; set; }
        //public string tmprInsCycleGbn { get; set; }
        //public string tmprInsCycleGbnName { get; set; }
        //public string tmprInsCondition { get; set; }
        //public string tmprInsCheckGbn { get; set; }
        //public string tmprInsCheckGbnName { get; set; }
        //public string tmprInsSpec { get; set; }
        //public string tmprInsMin { get; set; }
        //public string tmprInsMax { get; set; }
        //public string tmprInspectID { get; set; }
        //public string InspectDate { get; set; }
        //public string InspectTime { get; set; }
        //public string tmprInspectValue { get; set; }
        //public string Comments { get; set; }
    }

    class ChartInfoTemperRate
    {
        public SeriesCollection seriesCollection { get; set; }
        public string[] Labels { get; set; }
        public Func<double, string> Formatter { get; set; }
        public ColumnSeries columnSeries { get; set; }
        public ChartValues<double> chartRunningRate { get; set; }
        public ChartValues<double> chartGoalRate { get; set; }
        public ChartValues<double> chartMaxValue { get; set; }
        public ChartValues<double> chartMinValue { get; set; }
    }

}

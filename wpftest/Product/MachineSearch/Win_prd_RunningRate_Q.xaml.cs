using LiveCharts;
using LiveCharts.Wpf;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Data;
using System.Windows;
using System.Windows.Controls;
using WizMes_ANT.PopUp;
using WizMes_ANT.PopUP;
using System.Windows.Input;
using System.Threading;

namespace WizMes_ANT
{
    /**************************************************************************************************
    '** System 명 : WizMES
    '** Author    : Wizard
    '** 작성자    : 최준호
    '** 내용      : 설비가동률 조회
    '** 생성일자  : 2018.10~2019.2 월 사이
    '** 변경일자  : 
    '**------------------------------------------------------------------------------------------------
    ''*************************************************************************************************
    ' 변경일자  , 변경자, 요청자    , 요구사항ID  , 요청 및 작업내용
    '**************************************************************************************************
    ' ex) 2015.11.09, 박진성, 오영      ,S_201510_AFT_03 , 월별집계(가로) 순서 변경 : 합계/10월/9월/8월 순으로
    ' 2019.06.25 , 최준호 , 최규환   , 일자선택에서 기간선택으로, 비가동시간 나오게(특정 동작시 나오도록)
    '**************************************************************************************************/

    /// <summary>
    /// Win_prd_RunningRate_Q.xaml에 대한 상호 작용 논리
    /// </summary>
    public partial class Win_prd_RunningRate_Q : UserControl
    {
        string stDate = string.Empty;
        string stTime = string.Empty;

        Lib lib = new Lib();
        public DataGrid FilterGrid { get; set; }
        public DataTable FilterTable { get; set; }

        ObservableCollection<Win_prd_RunningRate_Q_CodeView> ovcCollection = new ObservableCollection<Win_prd_RunningRate_Q_CodeView>();

        public Win_prd_RunningRate_Q()
        {
            InitializeComponent();
        }

        private void UserControl_Loaded(object sender, RoutedEventArgs e)
        {
            stDate = DateTime.Now.ToString("yyyyMMdd");
            stTime = DateTime.Now.ToString("HHmm");

            DataStore.Instance.InsertLogByFormS(this.GetType().Name, stDate, stTime, "S");

            Lib.Instance.UiLoading(sender);

            dtpSDate.SelectedDate = DateTime.Today;
            dtpEDate.SelectedDate = DateTime.Today;
        }

        #region Header - 검색조건

        //전일
        private void btnYesterDay_Click(object sender, RoutedEventArgs e)
        {
            DateTime[] SearchDate = lib.BringLastDayDateTimeContinue(dtpEDate.SelectedDate.Value);

            dtpSDate.SelectedDate = SearchDate[0];
            dtpEDate.SelectedDate = SearchDate[1];
        }

        //금일
        private void btnToday_Click(object sender, RoutedEventArgs e)
        {
            dtpSDate.SelectedDate = DateTime.Today;
            dtpEDate.SelectedDate = DateTime.Today;
        }

        // 전월 버튼 클릭 이벤트
        private void btnLastMonth_Click(object sender, RoutedEventArgs e)
        {
            DateTime[] SearchDate = lib.BringLastMonthContinue(dtpSDate.SelectedDate.Value);

            dtpSDate.SelectedDate = SearchDate[0];
            dtpEDate.SelectedDate = SearchDate[1];
        }

        // 금월 버튼 클릭 이벤트
        private void btnThisMonth_Click(object sender, RoutedEventArgs e)
        {
            dtpSDate.SelectedDate = lib.BringThisMonthDatetimeList()[0];
            dtpEDate.SelectedDate = lib.BringThisMonthDatetimeList()[1];
        }

        #endregion

        #region Header - 오른쪽 상단 버튼

        //검색
        private void btnSearch_Click(object sender, RoutedEventArgs e)
        {
            //검색버튼 비활성화
            btnSearch.IsEnabled = false;

            Dispatcher.BeginInvoke(new Action(() =>

            {
                //로직
                using (Loading lw = new Loading(beSearch))
                {
                    lw.ShowDialog();
                }

            }), System.Windows.Threading.DispatcherPriority.Background);



            Dispatcher.BeginInvoke(new Action(() =>

            {
                btnSearch.IsEnabled = true;

            }), System.Windows.Threading.DispatcherPriority.Background);
            
        }

        private void beSearch()
        {
            if (dtpSDate.SelectedDate == null
                || dtpSDate.SelectedDate.ToString() == ""
                || dtpEDate.SelectedDate == null
                || dtpEDate.SelectedDate.ToString() == null)
            {
                MessageBox.Show("날짜를 정확히 입력해주세요.");
                //검색 다 되면 활성화
                btnSearch.IsEnabled = true;
                return;
            }

            FillGrid();

            if (dgdMain.Items.Count > 0)
            {
                FillChartGraph(dgdMain);
                dgdMain.SelectedIndex = 0;
            }
            else
            {
                if (lvcChart != null
                    && lvcChart.Series != null)
                {
                    lvcChart.Series.Clear();
                }

                MessageBox.Show("조회된 데이터가 없습니다.");
                return;
            }
        }

        //닫기
        private void btnClose_Click(object sender, RoutedEventArgs e)
        {
            DataStore.Instance.InsertLogByFormS(this.GetType().Name, stDate, stTime, "E");
            Lib.Instance.ChildMenuClose(this.ToString());
        }

        //엑셀
        private void btnExcel_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                DataTable dt = null;
                string Name = string.Empty;

                string[] lst = new string[2];
                lst[0] = "설비 가동률";
                lst[1] = dgdMain.Name;

                ExportExcelxaml ExpExc = new ExportExcelxaml(lst);

                ExpExc.ShowDialog();

                if (ExpExc.DialogResult.HasValue)
                {
                    if (ExpExc.choice.Equals(dgdMain.Name))
                    {
                        DataStore.Instance.InsertLogByForm(this.GetType().Name, "E");
                        if (ExpExc.Check.Equals("Y"))
                            dt = Lib.Instance.DataGridToDTinHidden(dgdMain);
                        else
                            dt = Lib.Instance.DataGirdToDataTable(dgdMain);

                        Name = dgdMain.Name;

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
            catch(Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }


        // 품번
        private void chkArticle_Click(object sender, MouseButtonEventArgs e)
        {
            if (chkArticle.IsChecked == true)
            {
                chkArticle.IsChecked = false;
                txtArticle.IsEnabled = false;
                btnArticle.IsEnabled = false;
            }
            else
            {
                chkArticle.IsChecked = true;
                txtArticle.IsEnabled = true;
                btnArticle.IsEnabled = true;
                txtArticle.Focus();
            }
        }
        // 품번
        private void chkArticle_Click(object sender, RoutedEventArgs e)
        {
            if (chkArticle.IsChecked == true)
            {
                txtArticle.IsEnabled = true;
                txtArticle.Focus();
                btnArticle.IsEnabled = true;
            }
            else
            {
                txtArticle.IsEnabled = false;
                btnArticle.IsEnabled = false;
            }
        }

        //플러스파인더 _ 품명_클릭.
        private void btnArticle_Click(object sender, RoutedEventArgs e)
        {
            MainWindow.pf.ReturnCode(txtArticle, 76, "");
        }

        private void TxtArticle_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                MainWindow.pf.ReturnCode(txtArticle, 76, "");
            }
        }

        #endregion

        #region 주요 메서드 - FillGrid

        //조회
        private void FillGrid()
        {
            try
            {
                dgdMain.ItemsSource = null;
                ovcCollection.Clear();

                Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                sqlParameter.Add("sFromDate", dtpSDate.SelectedDate != null ? dtpSDate.SelectedDate.Value.ToString("yyyyMMdd") : "");
                sqlParameter.Add("sToDate", dtpSDate.SelectedDate != null ? dtpEDate.SelectedDate.Value.ToString("yyyyMMdd") : "");

                sqlParameter.Add("nChkArticleID", chkArticle.IsChecked == true ? 1 : 0);
                sqlParameter.Add("sArticleID", chkArticle.IsChecked == true && txtArticle.Tag != null ? txtArticle.Tag.ToString() : "");

                //DataSet ds = DataStore.Instance.ProcedureToDataSet("xp_prd_sMCRunningRate_WPF_20201012", sqlParameter, false);
                DataSet ds = DataStore.Instance.ProcedureToDataSet_LogWrite("xp_prd_sMCRunningRate_WPF_20210517", sqlParameter, true, "R");
                if (ds != null && ds.Tables.Count > 0)
                {
                    DataTable dt = ds.Tables[0];

                    //dgdPNMSubul.ItemsSource = dt.DefaultView;
                    if (dt.Rows.Count > 0)
                    {
                        int i = 0;
                        //var DataTemplateHeader = new ProdNMtrSubulHeaderItem("원재료", "이월량", "생산량", "입고량", "사용량", "재고량");
                        //dgdPNMSubul.ItemsSource = dt.DefaultView;
                        DataRowCollection drc = dt.Rows;

                        foreach (DataRow dr in drc)
                        {
                            i++;
                            var dgdWorkRate = new Win_prd_RunningRate_Q_CodeView()
                            {
                                Num = i,
                                MCID = dr["MCID"].ToString(),
                                MCName = dr["MCName"].ToString(),

                                CT = stringFormatN1(dr["CT"]),

                                MonthWorkHour = stringFormatN1(dr["MonthWorkHour"]),
                                MonthWorkQty = stringFormatN0(dr["MonthWorkQty"]),
                                MonthGoalQty = stringFormatN0(dr["MonthGoalQty"]),
                                MonthWorkRate = stringFormatN1(dr["MonthWorkRate"]),

                                DayWorkHour = stringFormatN1(dr["DayWorkHour"]),
                                DayNonWorkHour = stringFormatN1(dr["DayNonWorkHour"]),
                                DayWorkQty = stringFormatN0(dr["DayWorkQty"]),
                                DayGoalQty = stringFormatN0(dr["DayGoalQty"]),
                                DayWorkRate = stringFormatN1(dr["DayWorkRate"]),
                            };

                            ovcCollection.Add(dgdWorkRate);
                        }

                        dgdMain.ItemsSource = ovcCollection;
                    }
                }
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

        #endregion

        #region 주요 메서드 그래프 조회 FillChartGraph

        private void FillChartGraph(DataGrid dataGrid)
        {
            try
            {
                if (lvcChart.Series != null)
                {
                    lvcChart.Series.Clear();
                }

                ChartInfoMCRunningRate chartRunningRate = new ChartInfoMCRunningRate();
                chartRunningRate.seriesCollection = new SeriesCollection();
                chartRunningRate.chartRunningRate = new ChartValues<double>();
                chartRunningRate.chartGoalRate = new ChartValues<double>();
                chartRunningRate.Labels = new string[dataGrid.Items.Count];

                int index = 0;
                for (int i = 0; i < dataGrid.Items.Count; i++)
                {
                    var Rating = dataGrid.Items[i] as Win_prd_RunningRate_Q_CodeView;

                    
                    if (Rating != null)
                    {
                        chartRunningRate.Labels[index] = Rating.MCName;
                        index++;

                        if (Rating.MonthWorkRate != null
                            && CheckConvertDouble(Rating.MonthWorkRate))
                        {
                            chartRunningRate.chartRunningRate.Add(ConvertDouble(Rating.MonthWorkRate));
                        }
                        else
                        {
                            chartRunningRate.chartRunningRate.Add(0);
                        }

                        if (Rating.DayWorkRate != null
                            && CheckConvertDouble(Rating.DayWorkRate))
                        {
                            chartRunningRate.chartGoalRate.Add(ConvertDouble(Rating.DayWorkRate));
                        }
                        else
                        {
                            chartRunningRate.chartGoalRate.Add(0);
                        }
                    }
                }

                chartRunningRate.seriesCollection.Add(new ColumnSeries
                {
                    Values = chartRunningRate.chartRunningRate,
                    //StackMode = StackMode.Values,
                    DataLabels = true,
                    Title = "월별 가동률"
                });

                chartRunningRate.seriesCollection.Add(new LineSeries
                {
                    Values = chartRunningRate.chartGoalRate,
                    Title = "일별 가동률"
                });

                chartRunningRate.Formatter = value => value + "(%)";
                this.DataContext = chartRunningRate;
            }
            catch(Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        #endregion

        private void BtnNoWorking_Click(object sender, RoutedEventArgs e)
        {
            var NoWorkingCode = dgdMain.SelectedItem as Win_prd_RunningRate_Q_CodeView;

            string sDate = dtpSDate.SelectedDate.Value.ToString("yyyyMMdd");
            string eDate = dtpEDate.SelectedDate.Value.ToString("yyyyMMdd");

            NoWorkInfo NoWorking = null;

            if (NoWorkingCode != null)
            {
                if (NoWorkingCode.DayNonWorkHour == null
                    || ConvertDouble(NoWorkingCode.DayNonWorkHour) == 0)
                    MessageBox.Show("선택된 자료의 비가동 시간을 확인해보세요.");
                else
                    NoWorking = new NoWorkInfo(sDate, eDate, NoWorkingCode.MCID);
            }
            else
            {
                NoWorking = new NoWorkInfo(sDate, eDate, "");
            }

            if (NoWorking != null)
                NoWorking.ShowDialog();
        }

        private void DgdMain_MouseLeftButtonDown(object sender, System.Windows.Input.MouseButtonEventArgs e)
        {
            if (e.ClickCount == 2)
            {
                var NoWorkingCode = dgdMain.SelectedItem as Win_prd_RunningRate_Q_CodeView;

                string sDate = dtpSDate.SelectedDate.Value.ToString("yyyyMMdd");
                string eDate = dtpEDate.SelectedDate.Value.ToString("yyyyMMdd");

                NoWorkInfo NoWorking = null;

                if (NoWorkingCode != null)
                {
                    if (NoWorkingCode.DayNonWorkHour == null
                        || ConvertDouble(NoWorkingCode.DayNonWorkHour) == 0)
                        MessageBox.Show("선택된 자료의 비가동 시간을 확인해보세요.");
                    else
                        NoWorking = new NoWorkInfo(sDate, eDate, NoWorkingCode.MCID);
                }
                else
                {
                    NoWorking = new NoWorkInfo(sDate, eDate, "");
                }

                if (NoWorking != null)
                    NoWorking.ShowDialog();
            }
        }

        private void dgdMain_ScrollChanged(object sender, ScrollChangedEventArgs e)
        {
            //svrHeader.ScrollToVerticalOffset(e.VerticalOffset);
            svrHeader.ScrollToHorizontalOffset(e.HorizontalOffset);
        }


        #region 기타 메서드 모음

        // 천마리 콤마, 소수점 버리기
        private string stringFormatN0(object obj)
        {
            return string.Format("{0:N0}", obj);
        }

        private string stringFormatN1(object obj)
        {
            return string.Format("{0:N1}", obj);
        }

        // 천마리 콤마, 소수점 두자리
        private string stringFormatN2(object obj)
        {
            return string.Format("{0:N2}", obj);
        }

        // 데이터피커 포맷으로 변경
        private string DatePickerFormat(string str)
        {
            string result = "";

            if (str.Length == 8)
            {
                if (!str.Trim().Equals(""))
                {
                    result = str.Substring(0, 4) + "-" + str.Substring(4, 2) + "-" + str.Substring(6, 2);
                }
            }

            return result;
        }

        // Int로 변환
        private int ConvertInt(string str)
        {
            int result = 0;
            int chkInt = 0;

            if (!str.Trim().Equals(""))
            {
                str = str.Replace(",", "");

                if (Int32.TryParse(str, out chkInt) == true)
                {
                    result = Int32.Parse(str);
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

        // 숫자로 변환 가능한지 체크 이벤트
        private bool CheckConvertInt(string str)
        {
            bool flag = false;
            int chkInt = 0;

            if (!str.Trim().Equals(""))
            {
                str = str.Trim().Replace(",", "");

                if (Int32.TryParse(str, out chkInt) == true)
                {
                    flag = true;
                }
            }

            return flag;
        }

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

        #endregion

    }

    public class Win_prd_RunningRate_Q_CodeView : BaseView
    {
        public override string ToString()
        {
            return (this.ReportAllProperties());
        }

        public int Num { get; set; }
        public string cls { get; set; }
        public string ProcessID { get; set; }
        public string Process { get; set; }
        public string Machineid { get; set; }
        public string Machine { get; set; }
        public string MachineNo { get; set; }
        public string CT { get; set; }

        public string MonthGoalQty { get; set; }
        public string DayGoalQty { get; set; }

        public string MonthWorkQty { get; set; }
        public string DayWorkQty { get; set; }

        public string MonthWorkRate { get; set; }
        public string DayWorkRate { get; set; }

        public string MonthWorkHour { get; set; }
        public string MonthNonWorkHour { get; set; }

        public string DayWorkHour { get; set; }
        public string DayNonWorkHour { get; set; }

        public string MCID { get; set; }
        public string MCName { get; set; }

    }

    public class ChartInfoMCRunningRate
    {
        public SeriesCollection seriesCollection { get; set; }
        public string[] Labels { get; set; }
        public Func<double, string> Formatter { get; set; }
        public ColumnSeries columnSeries { get; set; }
        public ChartValues<double> chartRunningRate { get; set; }
        public ChartValues<double> chartGoalRate { get; set; }
    }
}

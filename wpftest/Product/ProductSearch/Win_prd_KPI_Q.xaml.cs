using System;
using System.Collections.Generic;
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
using WizMes_WellMade.PopUP;
using WizMes_WellMade.PopUp;
using LiveCharts.Wpf;
using LiveCharts;
using System.Windows.Markup;


namespace WizMes_WellMade
{
    /// <summary>
    /// Win_prd_KPI_Q.xaml에 대한 상호 작용 논리
    /// </summary>
    public partial class Win_prd_KPI_Q : UserControl
    {
        string stDate = string.Empty;
        string stTime = string.Empty;

        Lib lib = new Lib();
        PlusFinder pf = new PlusFinder();

        int rowNum = 0;

        public Win_prd_KPI_Q()
        {
            InitializeComponent();
        }

        private void UserControl_Loaded(object sender, RoutedEventArgs e)
        {
            stDate = DateTime.Now.ToString("yyyyMMdd");
            stTime = DateTime.Now.ToString("HHmm");

            DataStore.Instance.InsertLogByFormS(this.GetType().Name, stDate, stTime, "S");

            lib.UiLoading(sender);
            dtpSDate.SelectedDate = DateTime.Today;
            dtpEDate.SelectedDate = DateTime.Today;
        }

        #region 상단 검색조건
        //전월
        private void btnLastMonth_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (dtpSDate.SelectedDate != null)
                {
                    DateTime ThatMonth1 = dtpSDate.SelectedDate.Value.AddDays(-(dtpSDate.SelectedDate.Value.Day - 1)); // 선택한 일자 달의 1일!

                    DateTime LastMonth1 = ThatMonth1.AddMonths(-1); // 저번달 1일
                    DateTime LastMonth31 = ThatMonth1.AddDays(-1); // 저번달 말일

                    dtpSDate.SelectedDate = LastMonth1;
                    dtpEDate.SelectedDate = LastMonth31;
                }
                else
                {
                    DateTime ThisMonth1 = DateTime.Today.AddDays(-(DateTime.Today.Day - 1)); // 이번달 1일

                    DateTime LastMonth1 = ThisMonth1.AddMonths(-1); // 저번달 1일
                    DateTime LastMonth31 = ThisMonth1.AddDays(-1); // 저번달 말일

                    dtpSDate.SelectedDate = LastMonth1;
                    dtpEDate.SelectedDate = LastMonth31;
                }
            }
            catch (Exception ee)
            {
                MessageBox.Show("오류지점 - btnLastMonth_Click : " + ee.ToString());
            }
        }

        //전일
        private void btnYesterday_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (dtpSDate.SelectedDate != null)
                {
                    dtpSDate.SelectedDate = dtpSDate.SelectedDate.Value.AddDays(-1);
                    dtpEDate.SelectedDate = dtpSDate.SelectedDate;
                }
                else
                {
                    dtpSDate.SelectedDate = DateTime.Today.AddDays(-1);
                    dtpEDate.SelectedDate = DateTime.Today.AddDays(-1);
                }
            }
            catch (Exception ee)
            {
                MessageBox.Show("오류지점 - btnYesterday_Click : " + ee.ToString());
            }
        }
        //금일
        private void btnToday_Click(object sender, RoutedEventArgs e)
        {
            dtpSDate.SelectedDate = DateTime.Today;
            dtpEDate.SelectedDate = DateTime.Today;
        }
        //금월
        private void btnThisMonth_Click(object sender, RoutedEventArgs e)
        {
            dtpSDate.SelectedDate = Lib.Instance.BringThisMonthDatetimeList()[0];
            dtpEDate.SelectedDate = Lib.Instance.BringThisMonthDatetimeList()[1];
        }

        #endregion

        #region Re_Search
        private void re_Search(int selectedIndex)
        {
            try
            {
                if (dgdOut.Items.Count > 0)
                {
                    dgdOut.Items.Clear();
                }

                if (dgdGonsu.Items.Count > 0)
                {
                    dgdGonsu.Items.Clear();
                }

                FillGrid();

            }
            catch (Exception ee)
            {
                MessageBox.Show("오류지점 - " + ee.ToString());
            }
        }

        #endregion

        #region 공수조회
        private void FillGrid()
        {
            try
            {
                string ArticleID = null;

                if (dgdOut.Items.Count > 0)
                {
                    dgdOut.Items.Clear();
                }
                if (dgdGonsu.Items.Count > 0)
                {
                    dgdGonsu.Items.Clear();
                }

                if (chkBuyerArticleNo.IsChecked == true)
                {
                    ArticleID = txtBuyerArticleNoSearch.Tag.ToString();
                }
                else if (CheckBoxArticleSearch.IsChecked == true)
                {
                    ArticleID = TextBoxArticleSearch.Tag.ToString();
                }


                DataSet ds = null;
                Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                sqlParameter.Clear();
                sqlParameter.Add("FromDate", dtpSDate.SelectedDate == null ? "" : dtpSDate.SelectedDate.Value.ToString().Replace("-", ""));
                sqlParameter.Add("ToDate", dtpEDate.SelectedDate == null ? "" : dtpEDate.SelectedDate.Value.ToString().Replace("-", ""));
                sqlParameter.Add("ArticleNo", ArticleID != null ? ArticleID : ""); //품번
                ds = DataStore.Instance.ProcedureToDataSet("xp_prd_sKPI_KPI", sqlParameter, false);

                if (ds != null && ds.Tables.Count > 0)
                {
                    DataTable dt = ds.Tables[0];
                    int i = 0;

                    if (dt.Rows.Count == 0)
                    {
                        MessageBox.Show("조회된 데이터가 없습니다.");
                    }
                    else
                    {
                        DataRowCollection drc = dt.Rows;

                        foreach (DataRow dr in drc)
                        {
                            var WPKQC = new Win_prd_KPI_Q_CodeView()
                            {
                                Num = i + 1,

                                //WorkDate = DatePickerFormat(dr["WorkDate"].ToString()),
                                WorkDate = dr["WorkDate"].ToString(),

                                WorkQty = stringFormatN0(dr["WorkQty"]),
                                WorkTime = stringFormatN1(dr["WorkTime"]),
                                WorkQtyPerHour = stringFormatN0(dr["WorkQtyPerHour"]),
                                //Daily_WorkQty = stringFormatN0(dr["Daily_WorkQty"]),
                                DefectQty = stringFormatN0(dr["DefectQty"]),
                                DefectWorkQty = stringFormatN0(dr["DefectWorkQty"]),
                                DefectRate = stringFormatN2(dr["DefectRate"]),
                                DefectUpRate = stringFormatN1(dr["DefectUpRate"]),
                                DefectGoalRate = stringFormatN1(dr["DefectGoalRate"]),
                                gbn = dr["gbn"].ToString(),
                                
                                Sort = dr["Sort"].ToString(),
                                WorkUpRate = stringFormatN1(dr["WorkUpRate"]),
                                WorkGoalRate = stringFormatN1(dr["WorkGoalRate"])
                            };
                            if (WPKQC.gbn == "P")
                            {
                                dgdGonsu.Items.Add(WPKQC);
                            }
                            if (WPKQC.gbn == "Q")
                            {
                                dgdOut.Items.Add(WPKQC);
                            }
                            i++;
                        }
                    }
                }
            }
            catch (Exception ee)
            {
                MessageBox.Show("오류지점 - " + ee.ToString());
            }
            finally
            {
                DataStore.Instance.CloseConnection();
            }
        }

        private void setGraphP(DataGrid dataGrid)
        {
            try
            {
                if (lvcChartP.Series != null)
                {
                    lvcChartP.Series.Clear();
                }

                ChartDTO chart = new ChartDTO();
                chart.seriesCollection = new SeriesCollection();
                chart.lineQty = new ChartValues<double>();
                chart.varQty = new ChartValues<double>();
                chart.Labels = new string[dataGrid.Items.Count];

                int index = 0;
                for (int i = 0; i < dataGrid.Items.Count; i++)
                {
                    var Rating = dataGrid.Items[i] as Win_prd_KPI_Q_CodeView;

                    if (Rating != null)
                    {
                        chart.Labels[index] = Rating.WorkDate;
                        index++;

                        if (Rating.WorkQtyPerHour != null)
                        {
                            chart.lineQty.Add(Convert.ToInt32(Rating.WorkQtyPerHour.Replace("," , "")));
                            //chart.lineQty.Add(Convert.ToInt32(Rating.Daily_WorkQty.Replace("," , "")));
                        }
                        else
                        {
                            chart.lineQty.Add(0);
                        }

                        //if (Rating.WorkQty != null)
                        //{
                        //    chart.varQty.Add(Convert.ToInt32(Rating.WorkQty.Replace(",", "")));
                        //}
                        //else
                        //{
                        //    chart.varQty.Add(0);
                        //}
                    }
                }

                chart.Formatter = data => data.Y.ToString("N0") + "(개)";

                chart.seriesCollection.Add(new LineSeries
                {
                    Values = chart.lineQty,
                    DataLabels = true,
                    Title = "시간당 생산량",
                    LabelPoint = chart.Formatter
                });

                //chart.seriesCollection.Add(new ColumnSeries
                //{
                //    Values = chart.varQty,
                //    DataLabels = true,
                //    Title = "생산량",
                //    LabelPoint = chart.Formatter
                //});

                lvcChartP.DataContext = chart;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void setGraphQ(DataGrid dataGrid)
        {
            try
            {
                if (lvcChartQ.Series != null)
                {
                    lvcChartQ.Series.Clear();
                }

                ChartDTO chart2 = new ChartDTO();
                chart2.seriesCollection = new SeriesCollection();
                chart2.lineQty = new ChartValues<double>();
                chart2.varQty = new ChartValues<double>();
                chart2.Labels = new string[dataGrid.Items.Count];
      

                int index = 0;
                for (int i = 0; i < dataGrid.Items.Count; i++)
                {
                    var Rating = dataGrid.Items[i] as Win_prd_KPI_Q_CodeView;

                    if (Rating != null)
                    {
                        chart2.Labels[index] = Rating.WorkDate;
                        index++;

                        if (Rating.DefectRate != null)
                        {
                            chart2.lineQty.Add(Convert.ToDouble(Rating.DefectRate));
                        }
                        else
                        {
                            chart2.lineQty.Add(0);
                        }

                        //if (Rating.DefectQty != null)
                        //{
                        //    chart2.varQty.Add(Convert.ToDouble(Rating.DefectQty));
                        //}
                        //else
                        //{
                        //    chart2.varQty.Add(0);
                        //}
                    }
                }

                chart2.PercentFormatter = data => data.Y + "(%)";
                chart2.Formatter = data => data.Y + "(개)";

                chart2.seriesCollection.Add(new LineSeries
                {
                    Values = chart2.lineQty,
                    DataLabels = true,
                    Title = "불량",
                    LabelPoint = chart2.PercentFormatter
                });

                //chart2.seriesCollection.Add(new ColumnSeries
                //{
                //    Values = chart2.varQty,
                //    DataLabels = true,
                //    Title = "불량량",
                //    LabelPoint = chart2.Formatter
                //});

                

                lvcChartQ.DataContext = chart2;

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
        #endregion

        private void btnSearch_Click(object sender, RoutedEventArgs e)
        {
            //검색버튼 비활성화
            btnSearch.IsEnabled = false;

            Dispatcher.BeginInvoke(new Action(() =>

            {
                try
                {
                    rowNum = 0;
                    using (Loading lw = new Loading(FillGrid))
                    {
                        lw.ShowDialog();

                        if (dgdGonsu.Items.Count > 0)
                        {
                            setGraphP(dgdGonsu);
                        }
                        if (dgdOut.Items.Count > 0)
                        {

                            setGraphQ(dgdOut);
                        }

                        if (dgdGonsu.Items.Count <= 0 || dgdOut.Items.Count <= 0)
                        {
                            MessageBox.Show("조회된 내용이 없습니다.");
                        }
                        
                       

                        btnSearch.IsEnabled = true;
                        
                    }
                }
                catch (Exception ee)
                {
                    MessageBox.Show("예외처리 - " + ee.ToString());
                }

            }), System.Windows.Threading.DispatcherPriority.Background);


        }

        private void btiClose_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                DataStore.Instance.InsertLogByFormS(this.GetType().Name, stDate, stTime, "E");
                lib.ChildMenuClose(this.ToString());
            }
            catch (Exception ee)
            {
                MessageBox.Show("오류지점 - " + ee.ToString());
            }
        }

        private void btiExcel_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                //if(dgdOut.Items.Count == 0 && dgdGonsu.Items.Count == 0)
                //{
                //    MessageBox.Show("먼저 검색해 주세요.");
                //    return;
                //}

                DataTable dt = null;
                string Name = string.Empty;

                string[] lst = new string[4];
                lst[0] = "생산성 향상";
                lst[1] = "품질 향상";
                lst[2] = dgdGonsu.Name;
                lst[3] = dgdOut.Name;

                ExportExcelxaml ExpExc = new ExportExcelxaml(lst);
                ExpExc.ShowDialog();

                if (ExpExc.DialogResult.HasValue)
                {
                    DataStore.Instance.InsertLogByForm(this.GetType().Name, "E");
                    if (ExpExc.choice.Equals(dgdGonsu.Name))
                    {
                        if (ExpExc.Check.Equals("Y"))
                            dt = Lib.Instance.DataGridToDTinHidden(dgdGonsu);
                        else
                            dt = Lib.Instance.DataGirdToDataTable(dgdGonsu);

                        Name = dgdGonsu.Name;
                        Lib.Instance.GenerateExcel(dt, Name);
                        Lib.Instance.excel.Visible = true;
                    }
                    else if (ExpExc.choice.Equals(dgdOut.Name))
                    {
                        if (ExpExc.Check.Equals("Y"))
                            dt = Lib.Instance.DataGridToDTinHidden(dgdOut);
                        else
                            dt = Lib.Instance.DataGirdToDataTable(dgdOut);

                        Name = dgdOut.Name;
                        Lib.Instance.GenerateExcel(dt, Name);
                        Lib.Instance.excel.Visible = true;
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
            catch (Exception ee)
            {
                MessageBox.Show("오류지점 - " + ee.ToString());
            }
        }

        private void lblBuyerArticleNo_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            if (chkBuyerArticleNo.IsChecked == true)
            {
                chkBuyerArticleNo.IsChecked = false;
            }
            else
            {
                chkBuyerArticleNo.IsChecked = true;
            }
        }
        // 거래처 체크박스 이벤트
        private void chkBuyerArticleNo_Checked(object sender, RoutedEventArgs e)
        {
            chkBuyerArticleNo.IsChecked = true;
            txtBuyerArticleNoSearch.IsEnabled = true;
            btnBuyerArticleNoSearch.IsEnabled = true;
        }
        private void chkBuyerArticleNo_UnChecked(object sender, RoutedEventArgs e)
        {
            chkBuyerArticleNo.IsChecked = false;
            txtBuyerArticleNoSearch.IsEnabled = false;
            btnBuyerArticleNoSearch.IsEnabled = false;
        }
        // 거래처 텍스트박스 엔터 → 플러스파인더
        private void txtBuyerArticleNoSearch_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                MainWindow.pf.ReturnCode(txtBuyerArticleNoSearch, 76, "");
            }
        }
        // 거래처 플러스파인더 이벤트
        private void btnBuyerArticleNoSearch_Click(object sender, RoutedEventArgs e)
        {
            // 거래처 : 0
            MainWindow.pf.ReturnCode(txtBuyerArticleNoSearch, 76, "");
        }

        //품명 라벨 클릭
        private void LabelArticleSearch_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            if (CheckBoxArticleSearch.IsChecked == true)
            {
                CheckBoxArticleSearch.IsChecked = false;
            }
            else
            {
                CheckBoxArticleSearch.IsChecked = true;
            }
        }

        private void CheckBoxArticleSearch_Checked(object sender, RoutedEventArgs e)
        {
            TextBoxArticleSearch.IsEnabled = true;
            ButtonArticleSearch.IsEnabled = true;
        }

        private void CheckBoxArticleSearch_Unchecked(object sender, RoutedEventArgs e)
        {
            TextBoxArticleSearch.IsEnabled = false;
            ButtonArticleSearch.IsEnabled = false;
        }

        private void TextBoxArticleSearch_KeyDown(object sender, KeyEventArgs e)
        {
            try
            {
                if (e.Key == Key.Enter)
                {
                    pf.ReturnCode(TextBoxArticleSearch, 77, TextBoxArticleSearch.Text);
                }
            }
            catch (Exception ee)
            {
                MessageBox.Show("예외처리 - " + ee.ToString());
            }
        }

        private void ButtonArticleSearch_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                pf.ReturnCode(TextBoxArticleSearch, 77, TextBoxArticleSearch.Text);
            }
            catch (Exception ee)
            {
                MessageBox.Show("예외처리 - " + ee.ToString());
            }
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

        // 천마리 콤마, 소수점 버리기
        private string stringFormatN0(object obj)
        {
            return string.Format("{0:N0}", obj);
        }

        // 천마리 콤마, 소수점 두자리
        private string stringFormatN2(object obj)
        {
            return string.Format("{0:N2}", obj);
        }
        // 천마리 콤마, 소수점 한자리
        private string stringFormatN1(object obj)
        {
            return string.Format("{0:N1}", obj);
        }
    }

    #region CodeView
    class Win_prd_KPI_Q_CodeView : BaseView
    {
        public override string ToString()
        {
            return (this.ReportAllProperties());
        }

        public int Num { get; set; }
        public string gbn
        {
            get; set;
        }
        public string GbnName { get; set; }
        public string WorkDate { get; set; }
        public string ArticleNo { get; internal set; }
        public string Article { get; internal set; }
        public string WorkQty { get; internal set; }
        public string WorkTime { get; internal set; }
        public string WorkQtyPerHour { get; internal set; }
        public string Daily_WorkQty { get; internal set; }
        public string WorkUpRate { get; set; }
        public string WorkGoalRate { get; set; }
        public string DefectQty { get; set; }
        public string DefectWorkQty { get; set; }
        public string DefectRate { get; set; }
        public string DefectUpRate { get; set; }
        public string DefectGoalRate { get; set; }
        public string Sort { get; set; }

    }

    public class ChartDTO
    {
        public SeriesCollection seriesCollection { get; set; }
        public string[] Labels { get; set; }
        public Func<ChartPoint, string> Formatter { get; set; }
        public Func<ChartPoint, string> PercentFormatter { get; set; }
        public ColumnSeries columnSeries { get; set; }
        public ChartValues<double> lineQty { get; set; }
        public ChartValues<double> varQty { get; set; }
    }


    #endregion

}
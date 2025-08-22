using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading;
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
    /// Win_mtr_Subul_Q_New.xaml에 대한 상호 작용 논리
    /// </summary>
    public partial class Win_prd_DailyWorkerProcessResult_Q : UserControl
    {
        int rowNum = 0;
        Lib lib = new Lib();
        ScrollViewer scrollView = null;
        ScrollViewer scrollView2 = null;

        public Win_prd_DailyWorkerProcessResult_Q()
        {
            InitializeComponent();
        }

        private void UserControl_Loaded(object sender, RoutedEventArgs e)
        {
            // 스크롤 동기화
            scrollView = dgdMainHeader;
            scrollView2 = getScrollbar(dgdMain);

            if (null != scrollView)
            {
                scrollView.ScrollChanged += new ScrollChangedEventHandler(scrollView_ScrollChanged);
            }
            if (null != scrollView2)
            {
                scrollView2.ScrollChanged += new ScrollChangedEventHandler(scrollView_ScrollChanged);
            }

            chkDateSrh.IsChecked = true;

            dtpSDate.SelectedDate = DateTime.Today;
            dtpEDate.SelectedDate = DateTime.Today;
        }

        #region Header 부분 - 검색조건

        // 일자
        private void lblDateSrh_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            if (chkDateSrh.IsEnabled == true)
            {
                if (chkDateSrh.IsChecked == true)
                {
                    chkDateSrh.IsChecked = false;
                }
                else
                {
                    chkDateSrh.IsChecked = true;
                }
            }
        }
        private void chkDateSrh_Checked(object sender, RoutedEventArgs e)
        {
            if (chkDateSrh.IsEnabled == true)
            {
                chkDateSrh.IsChecked = true;
                dtpSDate.IsEnabled = true;
                dtpEDate.IsEnabled = true;

                btnYesterDay.IsEnabled = true;
                btnToday.IsEnabled = true;
                btnLastMonth.IsEnabled = true;
                btnThisMonth.IsEnabled = true;
            }
        }
        private void chkDateSrh_Unchecked(object sender, RoutedEventArgs e)
        {
            if (chkDateSrh.IsEnabled == true)
            {
                chkDateSrh.IsChecked = false;
                dtpSDate.IsEnabled = false;
                dtpEDate.IsEnabled = false;

                btnYesterDay.IsEnabled = false;
                btnToday.IsEnabled = false;
                btnLastMonth.IsEnabled = false;
                btnThisMonth.IsEnabled = false;
            }
        }

        // 전일 금일 전월 금월 버튼

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

        // 공정 검색
        private void lblProcess_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            if (chkProcess.IsChecked == true)
            {
                chkProcess.IsChecked = false;
            }
            else
            {
                chkProcess.IsChecked = true;
            }
        }
        private void chkProcess_Checked(object sender, RoutedEventArgs e)
        {
            chkProcess.IsChecked = true;
            txtProcess.IsEnabled = true;
            btnPfProcess.IsEnabled = true;
        }
        private void chkProcess_Unchecked(object sender, RoutedEventArgs e)
        {
            chkProcess.IsChecked = false;
            txtProcess.IsEnabled = false;
            btnPfProcess.IsEnabled = false;
        }
        private void txtProcess_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                MainWindow.pf.ReturnCode(txtProcess, (int)Defind_CodeFind.DCF_PROCESS, "");
            }
        }
        private void btnPfProcess_Click(object sender, RoutedEventArgs e)
        {
            MainWindow.pf.ReturnCode(txtProcess, (int)Defind_CodeFind.DCF_PROCESS, "");
        }

        // 작업자 검색
        private void lblPerson_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            if (chkPerson.IsChecked == true)
            {
                chkPerson.IsChecked = false;
            }
            else
            {
                chkPerson.IsChecked = true;
            }
        }
        private void chkPerson_Checked(object sender, RoutedEventArgs e)
        {
            chkPerson.IsChecked = true;
            txtPerson.IsEnabled = true;
            btnPfPerson.IsEnabled = true;
        }
        private void chkPerson_Unchecked(object sender, RoutedEventArgs e)
        {
            chkPerson.IsChecked = false;
            txtPerson.IsEnabled = false;
            btnPfPerson.IsEnabled = false;
        }
        private void txtPerson_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                //MainWindow.pf.ReturnCode(txtPerson, (int)Defind_CodeFind.DCF_PERSON, "");
                re_search(rowNum);
            }
        }
        private void btnPfPerson_Click(object sender, RoutedEventArgs e)
        {
            MainWindow.pf.ReturnCode(txtPerson, (int)Defind_CodeFind.DCF_PERSON, "");
        }

        //// 품명 검색
        //private void lblArticle_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        //{
        //    if (chkArticle.IsChecked == true)
        //    {
        //        chkArticle.IsChecked = false;
        //    }
        //    else
        //    {
        //        chkArticle.IsChecked = true;
        //    }
        //}
        //private void chkArticle_Checked(object sender, RoutedEventArgs e)
        //{
        //    chkArticle.IsChecked = true;
        //    txtArticle.IsEnabled = true;
        //    btnPfArticle.IsEnabled = true;
        //}
        //private void chkArticle_Unchecked(object sender, RoutedEventArgs e)
        //{
        //    chkArticle.IsChecked = false;
        //    txtArticle.IsEnabled = false;
        //    btnPfArticle.IsEnabled = false;
        //}
        //private void txtArticle_KeyDown(object sender, KeyEventArgs e)
        //{
        //    if (e.Key == Key.Enter)
        //    {
        //        MainWindow.pf.ReturnCode(txtArticle, (int)Defind_CodeFind.DCF_Article, "");
        //    }
        //}
        //private void btnPfArticle_Click(object sender, RoutedEventArgs e)
        //{
        //    MainWindow.pf.ReturnCode(txtArticle, (int)Defind_CodeFind.DCF_Article, "");
        //}

        // 품번
        private void lblBuyerArticleNo_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
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
        private void chkBuyerArticleNo_Checked(object sender, RoutedEventArgs e)
        {
            chkBuyerArticleNo.IsChecked = true;
            txtBuyerArticleNo.IsEnabled = true;
            //btnPfBuyerArticleNo.IsEnabled = true;
        }
        private void chkBuyerArticleNo_Unchecked(object sender, RoutedEventArgs e)
        {
            chkBuyerArticleNo.IsChecked = false;
            txtBuyerArticleNo.IsEnabled = false;
            //btnPfBuyerArticleNo.IsEnabled = false;
        }
        private void txtBuyerArticleNo_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                //MainWindow.pf.ReturnCode(txtBuyerArticleNo, (int)Defind_CodeFind.DCF_BuyerArticleNo, "");
                re_search(rowNum);
            }
        }
        private void btnPfBuyerArticleNoClick(object sender, RoutedEventArgs e)
        {
            MainWindow.pf.ReturnCode(txtBuyerArticleNo, (int)Defind_CodeFind.DCF_BuyerArticleNo, "");
        }

        #endregion // Header 부분 - 검색조건

        #region Header 부분 - 오른쪽 버튼 모음 (검색, 닫기, 엑셀)

        // 검색버튼
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
            if (dtpEDate.SelectedDate != null)
            {
                DateTime FromDate = dtpEDate.SelectedDate.Value.AddMonths(-11);
                dtpSDate.SelectedDate = FromDate;
            }

            re_search(rowNum);
        }

        // 닫기버튼
        private void btnClose_Click(object sender, RoutedEventArgs e)
        {
            Lib.Instance.ChildMenuClose(this.ToString());
        }

        // 엑셀버튼
        private void btnExcel_Click(object sender, RoutedEventArgs e)
        {
            DataTable dt = null;
            string Name = string.Empty;

            string[] lst = new string[2];
            lst[0] = "기간별 작업자 실적";
            lst[1] = dgdMain.Name;

            ExportExcelxaml ExpExc = new ExportExcelxaml(lst);

            ExpExc.ShowDialog();

            if (ExpExc.DialogResult.HasValue)
            {
                if (ExpExc.choice.Equals(dgdMain.Name))
                {
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

        #endregion // Header 부분 - 오른쪽 버튼 모음 (검색, 닫기, 엑셀)

        private void re_search(int selectedIndex)
        {
            //2021-04-23 검색조건이 있어야 검색되도록 수정(공정)
            if((chkProcess.IsChecked == true && txtProcess.Text == "") || chkProcess.IsChecked == false)
            {
                MessageBox.Show("검색조건(공정)을 입력해주세요. 또는 검색조건(공정)의 체크박스에 체크를 해주세요.");
                return;
            }
            FillGrid();

            if (dgdMain.Items.Count > 0)
            {
                dgdMain.SelectedIndex = selectedIndex;
            }
            else
            {
                MessageBox.Show("조회된 데이터가 없습니다.");
                return;
            }
        }

        #region 조회 메서드

        private void FillGrid()
        {
            string[] Header = new string[12];

            DateTime ToDate = dtpEDate.SelectedDate.Value;
            ToDate = ToDate.AddMonths(1);
            dgdTotal.Items.Clear();
            if (dgdMain.Items.Count > 0)
            {
                dgdMain.Items.Clear();
            }

            try
            {
                Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                sqlParameter.Clear();

                //sqlParameter.Add("BasisMonth", dtpEDate.SelectedDate != null ? dtpEDate.SelectedDate.Value.ToString("yyyyMM") : "");
                sqlParameter.Add("BasisMonth", ToDate != null ? ToDate.ToString("yyyyMM").Trim() : "");

                sqlParameter.Add("chkProcessID", chkProcess.IsChecked == true ? 1 : 0);
                sqlParameter.Add("ProcessID", txtProcess.Tag != null ? txtProcess.Tag.ToString() : "");

                sqlParameter.Add("chkWorkerName", chkPerson.IsChecked == true ? 1 : 0);
                sqlParameter.Add("WorkerName", txtPerson.Text);

                sqlParameter.Add("chkBuyerArticleNo", chkBuyerArticleNo.IsChecked == true ? 1 : 0);
                sqlParameter.Add("BuyerArticleNo", !txtBuyerArticleNo.Text.Trim().Equals("") ? txtBuyerArticleNo.Text : "");

                DataSet ds = DataStore.Instance.ProcedureToDataSet("xp_prd_sProdSumPersonMonthSpread", sqlParameter, false);
                if (ds != null && ds.Tables.Count > 0)
                {
                    DataTable dt = ds.Tables[0];
                    int i = 0;

                    bool firstFlag = false;

                    //GLS에서 요청 2021-10-21
                    if (dt.Rows.Count == 0)
                    {
                        MessageBox.Show("조회된 데이터가 없습니다.");
                    }
                    if (dt.Rows.Count > 1)
                    {
                        DataRowCollection drc = dt.Rows;

                        foreach (DataRow dr in drc)
                        {
                            i++;

                            var WinR = new Win_prd_DailyWorkerProcessResult_Q_CodeView()
                            {
                                Num = i.ToString(),

                                cls = dr["cls"].ToString(),
                                ProcessID = dr["ProcessID"].ToString(),
                                Process = dr["Process"].ToString(),
                                AutoMcYNName = dr["AutoMcYNName"].ToString(),

                                WorkerName = dr["WorkerName"].ToString(),
                                ipSaDate = DatePickerFormat(dr["ipSaDate"].ToString()),

                                totProdQty = stringFormatN0(dr["totProdQty"]),
                                totProcessRate = stringFormatN0(dr["totProcessRate"]) + "%",

                                Month01 = dr["Month01"].ToString(),
                                ProdQty1 = stringFormatN0(dr["ProdQty1"]),
                                ProcessRate1 = stringFormatN0(dr["ProcessRate1"]) + "%",

                                Month02 = dr["Month02"].ToString(),
                                ProdQty2 = stringFormatN0(dr["ProdQty2"]),
                                ProcessRate2 = stringFormatN0(dr["ProcessRate2"]) + "%",

                                Month03 = dr["Month03"].ToString(),
                                ProdQty3 = stringFormatN0(dr["ProdQty3"]),
                                ProcessRate3 = stringFormatN0(dr["ProcessRate3"]) + "%",

                                Month04 = dr["Month04"].ToString(),
                                ProdQty4 = stringFormatN0(dr["ProdQty4"]),
                                ProcessRate4 = stringFormatN0(dr["ProcessRate4"]) + "%",

                                Month05 = dr["Month05"].ToString(),
                                ProdQty5 = stringFormatN0(dr["ProdQty5"]),
                                ProcessRate5 = stringFormatN0(dr["ProcessRate5"]) + "%",

                                Month06 = dr["Month06"].ToString(),
                                ProdQty6 = stringFormatN0(dr["ProdQty6"]),
                                ProcessRate6 = stringFormatN0(dr["ProcessRate6"]) + "%",

                                Month07 = dr["Month07"].ToString(),
                                ProdQty7 = stringFormatN0(dr["ProdQty7"]),
                                ProcessRate7 = stringFormatN0(dr["ProcessRate7"]) + "%",

                                Month08 = dr["Month08"].ToString(),
                                ProdQty8 = stringFormatN0(dr["ProdQty8"]),
                                ProcessRate8 = stringFormatN0(dr["ProcessRate8"]) + "%",

                                Month09 = dr["Month09"].ToString(),
                                ProdQty9 = stringFormatN0(dr["ProdQty9"]),
                                ProcessRate9 = stringFormatN0(dr["ProcessRate9"]) + "%",

                                Month10 = dr["Month10"].ToString(),
                                ProdQty10 = stringFormatN0(dr["ProdQty10"]),
                                ProcessRate10 = stringFormatN0(dr["ProcessRate10"]) + "%",

                                Month11 = dr["Month11"].ToString(),
                                ProdQty11 = stringFormatN0(dr["ProdQty11"]),
                                ProcessRate11 = stringFormatN0(dr["ProcessRate11"]) + "%",

                                Month12 = dr["Month12"].ToString(),
                                ProdQty12 = stringFormatN0(dr["ProdQty12"]),
                                ProcessRate12 = stringFormatN0(dr["ProcessRate12"]) + "%",
                            };

                            // 헤더 값 세팅
                            if (firstFlag == false)
                            {
                                Header[0] = getYearMonth(WinR.Month01);
                                Header[1] = getYearMonth(WinR.Month02);
                                Header[2] = getYearMonth(WinR.Month03);
                                Header[3] = getYearMonth(WinR.Month04);
                                Header[4] = getYearMonth(WinR.Month05);
                                Header[5] = getYearMonth(WinR.Month06);
                                Header[6] = getYearMonth(WinR.Month07);
                                Header[7] = getYearMonth(WinR.Month08);
                                Header[8] = getYearMonth(WinR.Month09);
                                Header[9] = getYearMonth(WinR.Month10);
                                Header[10] = getYearMonth(WinR.Month11);
                                Header[11] = getYearMonth(WinR.Month12);

                                firstFlag = true;
                            }

                            if (WinR.cls.Trim().Equals("9"))
                            {
                               
                                WinR.Process = "총계";
                                WinR.Total_Color = true;
                                dgdTotal.Items.Add(WinR);
                            }
                            else
                            {
                                dgdMain.Items.Add(WinR);
                            }

                            

                        }
                    }

                    // 헤더 세팅!!!!!
                    dgdHeader1.Content = Header[0];
                    dgdHeader2.Content = Header[1];
                    dgdHeader3.Content = Header[2];
                    dgdHeader4.Content = Header[3];
                    dgdHeader5.Content = Header[4];
                    dgdHeader6.Content = Header[5];
                    dgdHeader7.Content = Header[6];
                    dgdHeader8.Content = Header[7];
                    dgdHeader9.Content = Header[8];
                    dgdHeader10.Content = Header[9];
                    dgdHeader11.Content = Header[10];
                    dgdHeader12.Content = Header[11];

                    // 헤더 세팅!!!!!
                    dgdTotalHeader1.Content = Header[0];
                    dgdTotalHeader2.Content = Header[1];
                    dgdTotalHeader3.Content = Header[2];
                    dgdTotalHeader4.Content = Header[3];
                    dgdTotalHeader5.Content = Header[4];
                    dgdTotalHeader6.Content = Header[5];
                    dgdTotalHeader7.Content = Header[6];
                    dgdTotalHeader8.Content = Header[7];
                    dgdTotalHeader9.Content = Header[8];
                    dgdTotalHeader10.Content = Header[9];
                    dgdTotalHeader11.Content = Header[10];
                    dgdTotalHeader12.Content = Header[11];
                    
                    //tbkCount.Text = " ▶ 검색 결과 : " + i + " 건";
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

        #endregion // 조회 메서드

        #region 기타 메서드 모음

        // 월 만들기
        private string getYearMonth(string str)
        {
            str = str.Trim();

            if (str.Length == 6)
            {
                string Y = str.Substring(0, 4);
                string M = str.Substring(4, 2);

                if (M.Substring(0, 1).Equals("0"))
                {
                    M = M.Substring(1, 1);
                }

                str = Y + "년 " + M + "월";
            }
            else if (str.Length == 8)
            {
                string Y = str.Substring(0, 4);
                string M = str.Substring(4, 2);

                if (M.Substring(0, 1).Equals("0"))
                {
                    M = M.Substring(1, 1);
                }

                str = Y + "년 " + M + "월";
            }

            return str;
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

        // 시간 형식 6글자라면! 11:11:11
        private string DateTimeFormat(string str)
        {
            str = str.Replace(":", "").Trim();

            if (str.Length == 6)
            {
                string Hour = str.Substring(0, 2);
                string Min = str.Substring(2, 2);
                string Sec = str.Substring(4, 2);

                str = Hour + ":" + Min + ":" + Sec;
            }

            return str;
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

        #region 스크롤 Scroll 메서드 모음

        void scrollView_ScrollChanged(object sender, ScrollChangedEventArgs e)
        {
            var newOffset = e.HorizontalOffset;

            if ((null != scrollView) && (null != scrollView2))
            {
                scrollView.ScrollToHorizontalOffset(newOffset);
                scrollView2.ScrollToHorizontalOffset(newOffset);
            }
        }

        private ScrollViewer getScrollbar(DependencyObject dep)
        {
            for (int i = 0; i < VisualTreeHelper.GetChildrenCount(dep); i++)
            {
                var child = VisualTreeHelper.GetChild(dep, i);
                if ((null != child) && child is ScrollViewer)
                {
                    return (ScrollViewer)child;
                }
                else
                {
                    ScrollViewer sub = getScrollbar(child);
                    if (sub != null)
                    {
                        return sub;
                    }
                }
            }
            return null;
        }

        #endregion // 스크롤 Scroll 메서드 모음

    }

    class Win_prd_DailyWorkerProcessResult_Q_CodeView
    {
        public string Num { get; set; }


        public string cls { get; set; }
        public string ProcessID { get; set; }
        public string Process { get; set; }
        public string AutoMcYNName { get; set; }
        public string Machine { get; set; }

        public string Machineno { get; set; }
        public string WorkerName { get; set; }
        public string ipSaDate { get; set; }

        public string totProdQty { get; set; }
        public string totProcessRate { get; set; }

        public string Month01 { get; set; }
        public string ProdQty1 { get; set; }
        public string ProcessRate1 { get; set; }

        public string Month02 { get; set; }
        public string ProdQty2 { get; set; }
        public string ProcessRate2 { get; set; }

        public string Month03 { get; set; }
        public string ProdQty3 { get; set; }
        public string ProcessRate3 { get; set; }

        public string Month04 { get; set; }
        public string ProdQty4 { get; set; }
        public string ProcessRate4 { get; set; }

        public string Month05 { get; set; }
        public string ProdQty5 { get; set; }
        public string ProcessRate5 { get; set; }

        public string Month06 { get; set; }
        public string ProdQty6 { get; set; }
        public string ProcessRate6 { get; set; }

        public string Month07 { get; set; }
        public string ProdQty7 { get; set; }
        public string ProcessRate7 { get; set; }

        public string Month08 { get; set; }
        public string ProdQty8 { get; set; }
        public string ProcessRate8 { get; set; }

        public string Month09 { get; set; }
        public string ProdQty9 { get; set; }
        public string ProcessRate9 { get; set; }

        public string Month10 { get; set; }
        public string ProdQty10 { get; set; }
        public string ProcessRate10 { get; set; }

        public string Month11 { get; set; }
        public string ProdQty11 { get; set; }
        public string ProcessRate11 { get; set; }

        public string Month12 { get; set; }
        public string ProdQty12 { get; set; }
        public string ProcessRate12 { get; set; }
        public bool ArticleTotal_Color { get; set; }
        public bool Total_Color { get; set; }
    }
}

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
    public partial class Win_prd_MonthlyWorkerProcessResult_Q : UserControl
    {
        int rowNum = 0;
        Lib lib = new Lib(); //2021-04-23
        ScrollViewer scrollView = null;
        ScrollViewer scrollView2 = null;

        public Win_prd_MonthlyWorkerProcessResult_Q()
        {
            InitializeComponent();
        }

        private void UserControl_Loaded(object sender, RoutedEventArgs e)
        {
            // 스크롤 동기화
            //scrollView = dgdMainHeader;
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
        //전월
        private void btnLastMonth_Click(object sender, RoutedEventArgs e)
        {
            //dtpSDate.SelectedDate = Lib.Instance.BringLastMonthDatetimeList()[0];
            //dtpEDate.SelectedDate = Lib.Instance.BringLastMonthDatetimeList()[1];

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
        //전일
        private void btnYesterday_Click(object sender, RoutedEventArgs e)
        {
            //dtpSDate.SelectedDate = DateTime.Today.AddDays(-1);
            //dtpEDate.SelectedDate = DateTime.Today.AddDays(-1);

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

        //공정 검색
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
                re_search(0);
            }
        }
        private void btnPfPerson_Click(object sender, RoutedEventArgs e)
        {
            MainWindow.pf.ReturnCode(txtPerson, (int)Defind_CodeFind.DCF_PERSON, "");
        }
        #region 품명 검색 주석
        //// 품명 검색
        //private void lblArticle_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        //{
        //    if (chkBuyerArticleNo.IsChecked == true)
        //    {
        //        chkBuyerArticleNo.IsChecked = false;
        //    }
        //    else
        //    {
        //        chkBuyerArticleNo.IsChecked = true;
        //    }
        //}
        //private void chkArticle_Checked(object sender, RoutedEventArgs e)
        //{
        //    chkBuyerArticleNo.IsChecked = true;
        //    txtBuyerArticleNo.IsEnabled = true;
        //    btnPfBuyerArticleNo.IsEnabled = true;
        //}
        //private void chkArticle_Unchecked(object sender, RoutedEventArgs e)
        //{
        //    chkBuyerArticleNo.IsChecked = false;
        //    txtBuyerArticleNo.IsEnabled = false;
        //    btnPfBuyerArticleNo.IsEnabled = false;
        //}
        //private void txtArticle_KeyDown(object sender, KeyEventArgs e)
        //{
        //    if (e.Key == Key.Enter)
        //    {
        //        MainWindow.pf.ReturnCode(txtBuyerArticleNo, (int)Defind_CodeFind.DCF_BuyerArticleNo, "");
        //    }
        //}
        //private void btnPfArticle_Click(object sender, RoutedEventArgs e)
        //{
        //    MainWindow.pf.ReturnCode(txtBuyerArticleNo, (int)Defind_CodeFind.DCF_BuyerArticleNo, "");
        //}
        #endregion
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
            btnPfBuyerArticleNo.IsEnabled = true;
        }
        private void chkBuyerArticleNo_Unchecked(object sender, RoutedEventArgs e)
        {
            chkBuyerArticleNo.IsChecked = false;
            txtBuyerArticleNo.IsEnabled = false;
            btnPfBuyerArticleNo.IsEnabled = false;
        }
        private void txtBuyerArticleNo_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                MainWindow.pf.ReturnCode(txtBuyerArticleNo, (int)Defind_CodeFind.DCF_BuyerArticleNo, "");
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
            //2021-04-26 <DataGridTemplateColumn Header="No" x:Name="Num" MinWidth="40"> Header 추가하면 엑셀로 넘겼을 경우 컬럼명이 나타남(디자인 부분 수정)
            DataTable dt = null;
            string Name = string.Empty;
            string[] lst = new string[2];
            lst[0] = "월별 작업자 실적 집계";
            lst[1] = dgdMain.Name;

            ExportExcelxaml ExpExc = new ExportExcelxaml(lst);

            ExpExc.ShowDialog();

            if (ExpExc.DialogResult.HasValue)
            {
                if (ExpExc.choice.Equals(dgdMain.Name))
                {
                    if (ExpExc.Check.Equals("Y"))
                    {
                        dt = Lib.Instance.DataGridToDTinHidden(dgdMain);
                    }                        
                    else
                    {
                        dt = Lib.Instance.DataGirdToDataTable(dgdMain);
                    } 

                     Name = dgdMain.Name; 

                    if (lib.GenerateExcel(dt, Name)) //Lib.Instance.GenerateExcel(dt, Name)
                        lib.excel.Visible = true; //Lib.Instance.excel.Visible = true;
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
            if (dgdMain.Items.Count > 0)
            {
                dgdMain.Items.Clear();
            }

            dgdTotal.Items.Clear();
            
            try
            {
                Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                sqlParameter.Clear();

                sqlParameter.Add("BasisMonth", dtpEDate.SelectedDate != null ? dtpEDate.SelectedDate.Value.ToString("yyyyMM") : "");

                sqlParameter.Add("chkProcessID", chkProcess.IsChecked == true ? 1 : 0);
                sqlParameter.Add("ProcessID", txtProcess.Tag != null ? txtProcess.Tag.ToString() : "");

                sqlParameter.Add("chkWorkerName", chkPerson.IsChecked == true ? 1 : 0);
                sqlParameter.Add("WorkerName", !txtPerson.Text.Trim().Equals("") ? txtPerson.Text : "");

                sqlParameter.Add("chkBuyerArticleNo", chkBuyerArticleNo.IsChecked == true ? 1 : 0);
                sqlParameter.Add("BuyerArticleNo", !txtBuyerArticleNo.Text.Trim().Equals("") ? txtBuyerArticleNo.Text : "");

                DataSet ds = DataStore.Instance.ProcedureToDataSet("xp_prd_sProdSumPersonMonth", sqlParameter, false);

                if (ds != null && ds.Tables.Count > 0)
                {
                    DataTable dt = ds.Tables[0];
                    int i = 0;

                    if (dt.Rows.Count > 1)
                    {
                        DataRowCollection drc = dt.Rows;

                        foreach (DataRow dr in drc)
                        {
                            i++;

                            var WinR = new Win_prd_MonthlyWorkerProcessResult_Q_CodeView()
                            {
                                Num = i.ToString(),

                                cls = dr["cls"].ToString(),
                                ProcessID = dr["ProcessID"].ToString(),
                                Process = dr["Process"].ToString(),
                                AutoMcYNName = dr["AutoMcYNName"].ToString(),
                               
                                WorkerName = dr["WorkerName"].ToString(),
                                ipSaDate = DatePickerFormat(dr["ipSaDate"].ToString()),
                                BaseProdQty = stringFormatN0(dr["BaseProdQty"]),
                                BaseProcessRate = stringFormatN0(dr["BaseProcessRate"]) + "%",

                                ProdQty1 = stringFormatN0(dr["ProdQty1"]),
                                ProcessRate1 = stringFormatN0(dr["ProcessRate1"]) + "%",
                                AdvancedRate1 = stringFormatN0(dr["AdvancedRate1"]) + "%",

                                ProcessAmount = stringFormatN0(dr["ProcessAmount"]),
                            };

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

    class Win_prd_MonthlyWorkerProcessResult_Q_CodeView
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
        public string BaseProdQty { get; set; }

        public string BaseProcessRate { get; set; }
        public string ProdQty1 { get; set; }
        public string ProcessRate1 { get; set; }
        public string AdvancedRate1 { get; set; }

        public string Article { get; set; }
        public string ProcessAmount { get; set; }

        public bool Total_Color { get; set; }
    }
}

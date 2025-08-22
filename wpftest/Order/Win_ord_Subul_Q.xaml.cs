using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Data;
using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using WizMes_WellMade.PopUP;
using WizMes_WellMade.PopUp;

namespace WizMes_WellMade
{
    /// <summary>
    /// Win_ord_Subul_Q_New.xaml에 대한 상호 작용 논리
    /// </summary>
    public partial class Win_ord_Subul_Q : UserControl
    {
        string stDate = string.Empty;
        string stTime = string.Empty;
        PlusFinder pf = new PlusFinder();
        Lib lib = new Lib();
        int rowNum = 0;

        // 엑셀 활용 용도 (프린트)
        private Microsoft.Office.Interop.Excel.Application excelapp;
        private Microsoft.Office.Interop.Excel.Workbook workbook;
        private Microsoft.Office.Interop.Excel.Worksheet worksheet;
        private Microsoft.Office.Interop.Excel.Range workrange;
        private Microsoft.Office.Interop.Excel.Worksheet copysheet;
        private Microsoft.Office.Interop.Excel.Worksheet pastesheet;
        WizMes_WellMade.PopUp.NoticeMessage msg = new PopUp.NoticeMessage();

        // 인쇄를 위한 데이터 저장소
        List<Win_ord_Subul_Q_CodeView> lstSubul = new List<Win_ord_Subul_Q_CodeView>();

        ScrollViewer scrollView = null;
        ScrollViewer scrollView2 = null;

        int SearchCheckCount = 0;       // 검색 시 체크된 검색조건 갯수 파악. (2개 이상 체크시에만 )
        //2021-10-07
        double SumStockQty = 0; //현재 재고수량 
        double SumStuffinQty = 0; //입고 수량
        double SumOutQty = 0; //출고 수량

        public Win_ord_Subul_Q()
        {
            InitializeComponent();
        }

        private void Win_sbl_Subul_Q_Loaded(object sender, RoutedEventArgs e)
        {
            stDate = DateTime.Now.ToString("yyyyMMdd");
            stTime = DateTime.Now.ToString("HHmm");

            DataStore.Instance.InsertLogByFormS(this.GetType().Name, stDate, stTime, "S");

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

            // 초기세팅
            chkInOutDate.IsChecked = true;
            dtpFromDate.SelectedDate = DateTime.Today;
            dtpToDate.SelectedDate = DateTime.Today;

            rbnManageNumSrh.IsChecked = true; // 관리번호 : OrderID

            chkIn_NotApprovedIncloudSrh.IsChecked = true; // 입고미승인건 포함
            chkAutoInOutItemsIncloudSrh.IsChecked = true; // 자동입출건 포함

            // 콤보박스 세팅
            ComboBoxSetting();
        }

        #region 콤보박스 세팅

        // 콤보박스 세팅.
        private void ComboBoxSetting()
        {
            // 품명 그룹
            ObservableCollection<CodeView> cbArticleGroup = ComboBoxUtil.Instance.Gf_DB_MT_sArticleGrp();
            this.cboArticleGroupSrh.ItemsSource = cbArticleGroup;
            this.cboArticleGroupSrh.DisplayMemberPath = "code_name";
            this.cboArticleGroupSrh.SelectedValuePath = "code_id";
            this.cboArticleGroupSrh.SelectedIndex = 3; // 영업의 수불내역서는 제품만 보이면 된대서 일단 제품을 기본값으로 설정

            // 창고
            ObservableCollection<CodeView> cbWareHouse = ComboBoxUtil.Instance.Gf_DB_CM_GetComCodeDataset(null, "LOC", "Y", "", "");
            this.cboWareHouseSrh.ItemsSource = cbWareHouse;
            this.cboWareHouseSrh.DisplayMemberPath = "code_name";
            this.cboWareHouseSrh.SelectedValuePath = "code_id";
            this.cboWareHouseSrh.SelectedIndex = 0;

            // 입고구분
            ObservableCollection<CodeView> cbInGbn = ComboBoxUtil.Instance.Gf_DB_CM_GetComCodeDataset(null, "ICD", "Y", "", "");
            this.cboInGbnSrh.ItemsSource = cbInGbn;
            this.cboInGbnSrh.DisplayMemberPath = "code_id_plus_code_name";
            this.cboInGbnSrh.SelectedValuePath = "code_id";
            this.cboInGbnSrh.SelectedIndex = 0;

            // 출고구분
            ObservableCollection<CodeView> cbOutGbn = ComboBoxUtil.Instance.Gf_DB_CM_GetComCodeDataset(null, "OCD", "Y", "", "");
            this.cboOutGbnSrh.ItemsSource = cbOutGbn;
            this.cboOutGbnSrh.DisplayMemberPath = "code_id_plus_code_name";
            this.cboOutGbnSrh.SelectedValuePath = "code_id";
            this.cboOutGbnSrh.SelectedIndex = 0;

            // 공급유형 : 구매, MIP, 외주가공
            ObservableCollection<CodeView> cbSupplyType = ComboBoxUtil.Instance.Gf_DB_CM_GetComCodeDataset(null, "CMMASPLTYPE", "Y", "", "");
            this.cboSupplyTypeSrh.ItemsSource = cbSupplyType;
            this.cboSupplyTypeSrh.DisplayMemberPath = "code_name";
            this.cboSupplyTypeSrh.SelectedValuePath = "code_id";
            this.cboSupplyTypeSrh.SelectedIndex = 0;

        }

        #endregion 콤보박스 세팅

        #region Header 부분 - 검색조건

        // 입출일자 라벨 이벤트
        private void chkInOutDate_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            if (chkInOutDate.IsChecked == true)
            {
                chkInOutDate.IsChecked = false;
            }
            else
            {
                chkInOutDate.IsChecked = true;
            }
        }

        private void chkInOutDate_Checked(object sender, RoutedEventArgs e)
        {
            chkInOutDate.IsChecked = true;
            dtpFromDate.IsEnabled = true;
            dtpToDate.IsEnabled = true;

            btnYesterday.IsEnabled = true;
            btnToday.IsEnabled = true;
            btnLastMonth.IsEnabled = true;
            btnThisMonth.IsEnabled = true;
        }

        private void chkInOutDate_Unchecked(object sender, RoutedEventArgs e)
        {
            chkInOutDate.IsChecked = false;
            dtpFromDate.IsEnabled = false;
            dtpToDate.IsEnabled = false;

            btnYesterday.IsEnabled = false;
            btnToday.IsEnabled = false;
            btnLastMonth.IsEnabled = false;
            btnThisMonth.IsEnabled = false;
        }

        // 전일
        private void btnYesterday_Click(object sender, RoutedEventArgs e)
        {
            //dtpFromDate.SelectedDate = DateTime.Today.AddDays(-1);
            //dtpToDate.SelectedDate = DateTime.Today.AddDays(-1);

            if (dtpFromDate.SelectedDate != null)
            {
                dtpFromDate.SelectedDate = dtpFromDate.SelectedDate.Value.AddDays(-1);
                dtpToDate.SelectedDate = dtpFromDate.SelectedDate;
            }
            else
            {
                dtpFromDate.SelectedDate = DateTime.Today.AddDays(-1);
                dtpToDate.SelectedDate = DateTime.Today.AddDays(-1);
            }
        }
        // 금일
        private void btnToday_Click(object sender, RoutedEventArgs e)
        {
            dtpFromDate.SelectedDate = DateTime.Today;
            dtpToDate.SelectedDate = DateTime.Today;
        }
        // 전월
        private void btnLastMonth_Click(object sender, RoutedEventArgs e)
        {
            //dtpFromDate.SelectedDate = Lib.Instance.BringLastMonthDatetimeList()[0];
            //dtpToDate.SelectedDate = Lib.Instance.BringLastMonthDatetimeList()[1];

            if (dtpFromDate.SelectedDate != null)
            {
                DateTime ThatMonth1 = dtpFromDate.SelectedDate.Value.AddDays(-(dtpFromDate.SelectedDate.Value.Day - 1)); // 선택한 일자 달의 1일!

                DateTime LastMonth1 = ThatMonth1.AddMonths(-1); // 저번달 1일
                DateTime LastMonth31 = ThatMonth1.AddDays(-1); // 저번달 말일

                dtpFromDate.SelectedDate = LastMonth1;
                dtpToDate.SelectedDate = LastMonth31;
            }
            else
            {
                DateTime ThisMonth1 = DateTime.Today.AddDays(-(DateTime.Today.Day - 1)); // 이번달 1일

                DateTime LastMonth1 = ThisMonth1.AddMonths(-1); // 저번달 1일
                DateTime LastMonth31 = ThisMonth1.AddDays(-1); // 저번달 말일

                dtpFromDate.SelectedDate = LastMonth1;
                dtpToDate.SelectedDate = LastMonth31;
            }

        }
        // 금월
        private void btnThisMonth_Click(object sender, RoutedEventArgs e)
        {
            dtpFromDate.SelectedDate = Lib.Instance.BringThisMonthDatetimeList()[0];
            dtpToDate.SelectedDate = Lib.Instance.BringThisMonthDatetimeList()[1];
        }


      

        //검색조건 - 거래처 - 키다운
        private void txtCustomIDSrh_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                e.Handled = true;
                MainWindow.pf.ReturnCode(txtCustomIDSrh, 0, "");
            }
        }

        //검색조건 - 거래처 - 클릭
        private void btnCustomIDSrh_Click(object sender, RoutedEventArgs e)
        {
            MainWindow.pf.ReturnCode(txtCustomIDSrh, 0, "");
        }

    
        private void txtArticleIDSrh_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                e.Handled = true;
                MainWindow.pf.ReturnCode(txtArticleIDSrh, 77, ""); //2021-07-16
            }
        }
        private void btnArticleIDSrh_Click(object sender, RoutedEventArgs e)
        {
            MainWindow.pf.ReturnCode(txtArticleIDSrh, 77, "");
        }

   
  

   
        private void txtManageNumSrh_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                e.Handled = true;
                MainWindow.pf.ReturnCode(txtManageNumSrh, (int)Defind_CodeFind.DCF_ORDER, "");
            }
        }
        private void btnManageNumSrh_Click(object sender, RoutedEventArgs e)
        {
            MainWindow.pf.ReturnCode(txtManageNumSrh, (int)Defind_CodeFind.DCF_ORDER, "");
        }

     
        private void txtOrderNum_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                e.Handled = true;
                MainWindow.pf.ReturnCode(txtOrderNumSrh, (int)Defind_CodeFind.DCF_ORDER, "");
            }
        }
        private void btnOrderNum_Click(object sender, RoutedEventArgs e)
        {
            MainWindow.pf.ReturnCode(txtOrderNumSrh, (int)Defind_CodeFind.DCF_ORDER, "");
        }
            

        #endregion // Header 부분 - 검색조건

        #region Header 부분 - 입출고 근거번호 숨김 기능

        // 입출고 근거번호 숨김
        private void chkHideInOutReasonNumber_Click(object sender, RoutedEventArgs e)
        {
            if (chkHideInOutReasonNumberSrh.IsChecked == true)
            {
                // 숨김버튼 체크 → 발주번호, 오더번호 안보이게 설정
                Req_ID.Visibility = Visibility.Hidden;
                OrderID.Visibility = Visibility.Hidden;

                dgdReqID_With.Width = new GridLength(0);
                dgdOrderID_With.Width = new GridLength(0);
            }
            else
            {
                // 숨김버튼 체크 해제 → 발주번호, 오더번호 보이게 설정
                Req_ID.Visibility = Visibility.Visible;
                OrderID.Visibility = Visibility.Visible;

                dgdReqID_With.Width = new GridLength(Req_ID.ActualWidth);
                dgdOrderID_With.Width = new GridLength(OrderID.ActualWidth);

                //for (int i = 0; i < 10; i++)
                //{
                //    Binding reqBinding = new Binding("ElementName=Req_ID, Path=ActualWidth");
                //    Binding orderBinding = new Binding("ElementName=OrderID, Path=ActualWidth");

                //    dgdReqID_With.SetBinding(ColumnDefinition.WidthProperty, reqBinding);
                //    dgdOrderID_With.SetBinding(ColumnDefinition.WidthProperty, orderBinding);
                //}


                //MessageBox.Show(dgdReqID_With.GetBindingExpression(ColumnDefinition.WidthProperty).ParentBinding.Path.Path);
            }
        }

        #endregion // Header 부분 - 입출고 근거번호 숨김 기능

        #region Header 부분 - 상단 오른쪽 버튼

        // 검색 버튼 이벤트
        private void btnSearch_Click(object sender, RoutedEventArgs e)
        {
            if(lib.DatePickerCheck(dtpFromDate, dtpToDate, chkInOutDate))
            {
                using (Loading ld = new Loading(beSearch))
                {
                    ld.ShowDialog();
                }
            }

         
        }

        private void beSearch()
        {
            //검색버튼 비활성화
            btnSearch.IsEnabled = false;

            Dispatcher.BeginInvoke(new Action(() =>
            {
                //로직
                rowNum = 0;
                re_Search(rowNum);
            }), System.Windows.Threading.DispatcherPriority.Background);

            btnSearch.IsEnabled = true;
        }

        private void btnClose_Click(object sender, RoutedEventArgs e)
        {
            DataStore.Instance.InsertLogByFormS(this.GetType().Name, stDate, stTime, "E");
            Lib.Instance.ChildMenuClose(this.ToString());
        }

        #region 인쇄

        // 인쇄 기능.
        private void btnPrint_Click(object sender, RoutedEventArgs e)
        {
            ContextMenu menu = btnPrint.ContextMenu;
            menu.StaysOpen = true;
            menu.IsOpen = true;
        }

        // 인쇄 서브메뉴1. 미리보기
        private void menuSeeAhead_Click(object sender, RoutedEventArgs e)
        {
            if (dgdMain.Items.Count < 1)
            {
                MessageBox.Show("먼저 검색해 주세요.");
                return;
            }

            msg.Show();
            msg.Topmost = true;
            msg.Refresh();

            PrintWork(true);
        }

        // 인쇄 서브메뉴2. 바로인쇄
        private void menuRighPrint_Click(object sender, RoutedEventArgs e)
        {
            if (dgdMain.Items.Count < 1)
            {
                MessageBox.Show("먼저 검색해 주세요.");
                return;
            }
            DataStore.Instance.InsertLogByForm(this.GetType().Name, "P");
            msg.Show();
            msg.Topmost = true;
            msg.Refresh();

            PrintWork(false);
        }
        //인쇄 서브메뉴3. 그냥 닫기
        private void menuClose_Click(object sender, RoutedEventArgs e)
        {
            ContextMenu menu = btnPrint.ContextMenu;
            menu.StaysOpen = false;
            menu.IsOpen = false;
        }

        private void PrintWork(bool previewYN)
        {
            Lib lib = new Lib();
            try
            {
                if (lstSubul.Count > 0)
                {
                    excelapp = new Microsoft.Office.Interop.Excel.Application();

                    string MyBookPath = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetEntryAssembly().Location) + "\\Report\\자재수불내역(영업관리).xls";
                    workbook = excelapp.Workbooks.Add(MyBookPath);
                    worksheet = workbook.Sheets["Form"];
                    pastesheet = workbook.Sheets["Print"];

                    // 창고로 검색했다면 창고 입력 : D3
                    workrange = worksheet.get_Range("D3");
                    workrange.Value2 = chkWareHouseSrh.IsChecked == true && cboWareHouseSrh.SelectedValue != null ? cboWareHouseSrh.SelectedValue.ToString() : "";

                    // 기간 설정후에 검색했다면 검색일자 2019.10.01 ~ 2019.10.31 입력 : D4
                    workrange = worksheet.get_Range("D4");
                    workrange.Value2 = chkInOutDate.IsChecked == true ? dtpFromDate.SelectedDate.Value.ToString("yyyy-MM-dd") + " ~ " + dtpToDate.SelectedDate.Value.ToString("yyyy-MM-dd") : "";

                    // 일자 입력(오늘일자) : AE4
                    workrange = worksheet.get_Range("AO4");
                    workrange.Value2 = DateTime.Today.ToString("yyyy-MM-dd");

                    // 페이지 계산 등
                    int rowCount = lstSubul.Count;
                    int excelStartRow = 7;

                    // 총 데이터를 입력할수 있는 갯수
                    int totalDataInput = 39;

                    // 카피할 다음페이지 인덱스
                    int nextCopyLine = 51;

                    int copyLine = 0;
                    int Page = 0;
                    int PageAll = (int)Math.Ceiling(1.0 * rowCount / totalDataInput);
                    int DataCount = 0;

                    for (int k = 0; k < PageAll; k++)
                    {
                        Page++;
                        copyLine = ((Page - 1) * (nextCopyLine - 1));

                        int excelNum = 0;

                        // 기존에 있는 데이터 지우기 "A7", "AP45"
                        worksheet.Range["A7", "AP45"].EntireRow.ClearContents();

                        for (int i = DataCount; i < rowCount; i++)
                        {
                            if (i == totalDataInput * Page)
                            {
                                break;
                            }

                            var Subul = lstSubul[i];
                            int excelRow = excelStartRow + excelNum;

                            if (Subul != null)
                            {
                                // 품명
                                workrange = worksheet.get_Range("A" + excelRow);
                                workrange.Value2 = Subul.Article;

                                // 일자
                                workrange = worksheet.get_Range("F" + excelRow);
                                workrange.Value2 = "'" + MonthSlashDay(Subul.ioDate);

                                // 창고
                                workrange = worksheet.get_Range("H" + excelRow);
                                workrange.Value2 = Subul.LocName;

                                // 발주번호
                                workrange = worksheet.get_Range("K" + excelRow);
                                workrange.Value2 = Subul.Req_ID;

                                // 수량
                                workrange = worksheet.get_Range("P" + excelRow);
                                workrange.Value2 = Subul.StuffQty;

                                // 오더번호
                                workrange = worksheet.get_Range("U" + excelRow);
                                workrange.Value2 = Subul.OrderNo;

                                // 수량
                                workrange = worksheet.get_Range("Z" + excelRow);
                                workrange.Value2 = Subul.OutQty;

                                // 입출고처
                                workrange = worksheet.get_Range("AE" + excelRow);
                                workrange.Value2 = Subul.RelLocName;

                                // 재고량
                                workrange = worksheet.get_Range("AL" + excelRow);
                                workrange.Value2 = Subul.StockQty;

                                // 비고
                                workrange = worksheet.get_Range("AP" + excelRow);
                                workrange.Value2 = Subul.Remark;


                                excelNum++;
                                DataCount = i;
                            }
                        }

                        // 2장 이상 넘어가면 페이지 넘버 입력
                        if (PageAll > 1)
                        {
                            pastesheet.PageSetup.CenterFooter = "&P / &N";
                        }

                        // Form 시트 내용 Print 시트에 복사 붙여넣기
                        worksheet.Select();
                        worksheet.UsedRange.EntireRow.Copy();
                        pastesheet.Select();
                        workrange = pastesheet.Cells[copyLine + 1, 1];
                        workrange.Select();
                        pastesheet.Paste();

                        DataCount++;
                    }

                    pastesheet.UsedRange.EntireRow.Select();

                    //
                    excelapp.Visible = true;
                    msg.Hide();

                    if (previewYN == true)
                    {
                        pastesheet.PrintPreview();
                    }
                    else
                    {
                        pastesheet.PrintOutEx();
                    }

                }
                else
                {
                    msg.Hide();

                    //lstsubul에 인쇄할 데이터가 없을 때는 메세지를 띄워주는데, 품번계, 총계, 이월 등은 출력이안돼..
                    MessageBox.Show("출력할 데이터가 없습니다. (품명계, 총계, 이월만 있는 경우 출력이 되지 않습니다.");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("오류 발생, 오류 내용 : " + ex.ToString());
            }
            finally
            {
                lib.ReleaseExcelObject(workbook);
                lib.ReleaseExcelObject(worksheet);
                lib.ReleaseExcelObject(pastesheet);
                lib.ReleaseExcelObject(excelapp);
                lib = null;
            }
        }

        #endregion

        // 엑셀 버튼 이벤트
        private void btnExcel_Click(object sender, RoutedEventArgs e)
        {
            DataTable dt = null;
            string Name = string.Empty;
            Lib lib = new Lib();

            string[] lst = new string[2];
            lst[0] = "수불내역";
            lst[1] = dgdMain.Name;

            ExportExcelxaml ExpExc = new ExportExcelxaml(lst);

            ExpExc.ShowDialog();

            if (ExpExc.DialogResult.HasValue)
            {
                if (ExpExc.choice.Equals(dgdMain.Name))
                {
                    DataStore.Instance.InsertLogByForm(this.GetType().Name, "E");
                    if (ExpExc.Check.Equals("Y"))
                        dt = lib.DataGridToDTinHidden(dgdMain);
                    else
                        dt = lib.DataGirdToDataTable(dgdMain);

                    Name = dgdMain.Name;

                    if (lib.GenerateExcel(dt, Name))
                    {
                        lib.excel.Visible = true;
                        lib.ReleaseExcelObject(lib.excel);
                    }
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

            lib = null;
        }

        #endregion // Header 부분 - 상단 오른쪽 버튼

        #region 주요 메서드 모음

        private void re_Search(int selectedIndex)
        {
            //if (chkManageNum.IsChecked == true && txtManageNum.Text == "")
            //{
            //    MessageBox.Show("관리번호를 입력한 후 검색을 하거나 관리번호 체크를 해제 후 검색 하세요");
            //    return;
            //}
            //else if (chkOrderNum.IsChecked == true && txtOrderNum.Text == "")
            //{
            //    MessageBox.Show("발주번호를 입력한 후 검색을 하거나 발주번호 체크를 해제 후 검색 하세요");
            //    return;
            //}
            //else if (chkCustomer.IsChecked == true && txtCustomer.Text == "")
            //{
            //    MessageBox.Show("거래처를 입력한 후 검색을 하거나 거래처 체크를 해제 후 검색 하세요");
            //    return;
            //}
            //else if (chkArticle.IsChecked == true && txtArticle.Text == "")
            //{
            //    MessageBox.Show("품번를 입력한 후 검색을 하거나 품번 체크를 해제 후 검색 하세요");
            //    return;
            //}


            FillGrid();

            if (dgdMain.Items.Count > 0)
            {
                dgdMain.SelectedIndex = rowNum;
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
            lstSubul.Clear();

            if (dgdMain.Items.Count > 0)
            {
                dgdMain.Items.Clear();
            }

            try
            {
                //관리번호 1 오더번호 2 // 아무것도 체크안하면 0 ssw
                Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                sqlParameter.Clear();

                sqlParameter.Add("nChkDate", chkInOutDate.IsChecked == true ? 1 : 0);
                sqlParameter.Add("sSDate", chkInOutDate.IsChecked == true && dtpFromDate.SelectedDate != null ? dtpFromDate.SelectedDate.Value.ToString("yyyyMMdd") : "");
                sqlParameter.Add("sEDate", chkInOutDate.IsChecked == true && dtpToDate.SelectedDate != null ? dtpToDate.SelectedDate.Value.ToString("yyyyMMdd") : "");
                sqlParameter.Add("nChkCustom", chkCustomIDSrh.IsChecked == true ? 1 : 0);
                sqlParameter.Add("sCustomID", chkCustomIDSrh.IsChecked == true ? txtCustomIDSrh.Tag?.ToString() ?? string.Empty : string.Empty);

                sqlParameter.Add("nChkArticleID", chkArticleIDSrh.IsChecked == true ? 1 : 0);
                sqlParameter.Add("sArticleID", chkArticleIDSrh.IsChecked ==true ? txtArticleIDSrh.Tag?.ToString() ?? string.Empty : string.Empty);
                sqlParameter.Add("nChkOrder", chkManageNumSrh.IsChecked == true ? 2 : 0); //이 친구는 주석이라서 무슨 값이 들어가도...

                sqlParameter.Add("sOrder", txtManageNumSrh.Text);
                sqlParameter.Add("ArticleGrpID", chkArticleGroupSrh.IsChecked == true && cboArticleGroupSrh.SelectedValue != null ? cboArticleGroupSrh.SelectedValue.ToString() : "");

                sqlParameter.Add("sFromLocID", chkWareHouseSrh.IsChecked == true ? cboWareHouseSrh.SelectedValue?.ToString() ?? string.Empty : string.Empty);
                sqlParameter.Add("sToLocID", "");
                sqlParameter.Add("nChkOutClss", chkOutGbnSrh.IsChecked == true ? 1 : 0);
                sqlParameter.Add("sOutClss", chkOutGbnSrh.IsChecked == true ? cboOutGbnSrh.SelectedValue?.ToString() ?? string.Empty : string.Empty);
                
                sqlParameter.Add("nChkInClss", chkInGbnSrh.IsChecked == true ? 1 : 0);
                sqlParameter.Add("sInClss", chkInGbnSrh.IsChecked == true ? cboInGbnSrh.SelectedValue?.ToString() ?? string.Empty : string.Empty);
                sqlParameter.Add("nChkReqID", chkOrderNumSrh.IsChecked == true ? 1 : 0);
                sqlParameter.Add("sReqID", chkOrderNumSrh.IsChecked == true && txtOrderNumSrh.Tag != null ? txtOrderNumSrh.Tag.ToString() : "");
                sqlParameter.Add("incNotApprovalYN", chkIn_NotApprovedIncloudSrh.IsChecked == true ? "Y" : "N");
                sqlParameter.Add("incAutoInOutYN", chkAutoInOutItemsIncloudSrh.IsChecked == true ? "Y" : "N");

                sqlParameter.Add("sProductYN", chkArticleGroupSrh.IsChecked == true &&  cboArticleGroupSrh.SelectedValue.Equals("3") ? "Y" : "");    // 제품으로 조회시 Y 그게 아니면 빈값
                sqlParameter.Add("nMainItem", chkMainInterestItemsSeeSrh.IsChecked == true ? 1 : 0);
                sqlParameter.Add("nCustomItem", chkRegistItemsByCustomerSrh.IsChecked == true ? 1 : 0);
                sqlParameter.Add("nSupplyType", chkSupplyTypeSrh.IsChecked == true ? 1 : 0);
                sqlParameter.Add("sSupplyType", chkSupplyTypeSrh.IsChecked == true ? cboSupplyTypeSrh.SelectedValue?.ToString() ?? string.Empty: string.Empty);

                sqlParameter.Add("nBuyerArticleNo", chkBuyerArticleNoSrh.IsChecked == true ? 1 : 0);
                sqlParameter.Add("BuyerArticleNo", chkBuyerArticleNoSrh.IsChecked == true ? txtBuyerArticleNoSrh.Tag?.ToString()  ?? string.Empty : string.Empty);

                DataSet ds = DataStore.Instance.ProcedureToDataSet_LogWrite("xp_Subul_sSubul", sqlParameter, true, "R");

                if (ds != null && ds.Tables.Count > 0)
                {
                    DataTable dt = ds.Tables[0];

                    if (dt.Rows.Count > 0)
                    {
                        int i = 0;
                        int totalStuffQty = 0;
                        int totalOutQty = 0;
                        int totalStockQty = 0;
                        DataRowCollection drc = dt.Rows;

                        foreach (DataRow dr in drc)
                        {
                            i++;
                            var Subul = new Win_ord_Subul_Q_CodeView()
                            {
                                Num = i,

                                cls = dr["cls"].ToString(),

                                BuyerArticleNo = dr["BuyerArticleNo"].ToString(),
                                ArticleID = dr["ArticleID"].ToString(),
                                Article = dr["Article"].ToString(),
                                ioDate = DatePickerFormat(dr["ioDate"].ToString()),
                                LocID = dr["LocID"].ToString(),
                                LocName = dr["LocName"].ToString(),

                                Req_ID = dr["req_ID"].ToString(),
                                //ReqName = dr["ReqName "].ToString(),
                                StuffRoll = dr["StuffRoll"].ToString(),
                                StuffQty = stringFormatN0AndZeroEmpty(dr["StuffQty"]),
                                UnitClss = dr["UnitClss"].ToString(),

                                UnitClssName = dr["UnitClssName"].ToString(),
                                CustomID = dr["CustomID"].ToString(),
                                CustomName = dr["CustomName"].ToString(),
                                OrderID = dr["OrderID"].ToString(),
                                OrderNo = dr["OrderNo"].ToString(),

                                OutRoll = dr["OutRoll"].ToString(),
                                OutQty = stringFormatN0AndZeroEmpty(dr["OutQty"]),
                                RelLocID = dr["RelLocID"].ToString(),
                                RelLocName = dr["RelLocName"].ToString(),
                                InOutClssName = dr["inoutClssname"].ToString(),

                                Remark = dr["Remark"].ToString(),
                                StockQty = stringFormatN0AndZeroEmpty(dr["StockQty"]),

                            };

                            if (dr["cls"].ToString().Trim().Equals("0")) // 이월
                            {
                                Subul.ioDate = "이월";
                                Subul.StockQty = stringFormatN0(dr["StuffQty"]);                      //2021-10-08 재고량이 바로 보이게 하기 위해 추가
                                SumStockQty = SumStockQty + ConvertDouble(dr["StuffQty"].ToString()); //2021-10-08 재고량을 사용하기 위해 변수에 저장
                            }
                            else if (dr["cls"].ToString().Trim().Equals("6")) // 품번계
                            {
                                Subul.Article = "";
                                Subul.BuyerArticleNo = "품번계";
                                Subul.ioDate = "";
                                Subul.LocName = "";
                                Subul.ArticleTotal_Color = true;
                                SumStockQty = 0; //2021-10-08 0으로 변경해야 다른 품명의 재고가 맞게 나옴
                            }
                            else if (dr["cls"].ToString().Trim().Equals("7")) // 총계
                            {
                                Subul.BuyerArticleNo = "";
                                Subul.Article = "총계";
                                Subul.ioDate = "";
                                Subul.LocName = "";
                                Subul.Total_Color = true;
                                totalOutQty += lib.RemoveComma(dr["OutQty"].ToString(), 0);
                                totalStuffQty += lib.RemoveComma(dr["StuffQty"].ToString(), 0);
                                totalStockQty += lib.RemoveComma(dr["StockQty"].ToString(), 0);
                            }
                            else if (dr["cls"].ToString().Trim().Equals("1"))
                            {
                                SumStuffinQty = SumStockQty + ConvertDouble(dr["StuffQty"].ToString());     //2021-10-08 재고량 수정 후 저장  //b == 0 ? SumStockQty + ConvertDouble(dr["StuffQty"].ToString()) : b + ConvertDouble(dr["StuffQty"].ToString());
                                Subul.StockQty = stringFormatN0(SumStuffinQty);                                  //2021-10-08 재고량을 바로 보기 위해 추가
                                SumStockQty = SumStuffinQty;                                                //2021-10-08 재고량을 수정 후 다시 재고량변수에 저장
                            }
                            else if (dr["cls"].ToString().Trim().Equals("2"))
                            {
                                SumOutQty = SumStockQty - ConvertDouble(dr["OutQty"].ToString());   //2021-10-08 재고량 수정 후 저장 //SumStockQty == 0 ? SumStuffinQty - ConvertDouble(dr["OutQty"].ToString()) : SumStockQty - ConvertDouble(dr["OutQty"].ToString());
                                Subul.StockQty = stringFormatN0(SumOutQty);                              //2021-10-08 재고량을 바로 보기 위해 추가
                                SumStockQty = SumOutQty;                                            //2021-10-08 재고량을 수정 후 다시 재고량변수에 저장
                            }
                            else // 그냥 입고, 출고 데이터를 인쇄를 위해 따로 저장
                            {
                                lstSubul.Add(Subul);
                            }

                            dgdMain.Items.Add(Subul);
                        }

                        if(dgdMain.Items.Count > 0)
                        {
                            var subulTotal = new Win_ord_Subul_Q_CodeView_Total
                            {
                                TotalOutQty = stringFormatN0(totalOutQty),
                                TotalStuffQty = stringFormatN0(totalStuffQty),
                                TotalStockQty = stringFormatN0(totalStockQty),
                            };

                            dgdSubulQTotal.Items.Add(subulTotal);   
                        }

                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
            finally
            {
                DataStore.Instance.CloseConnection();
            }
        }

        #endregion // 조회 메서드

        #endregion // 주요 메서드 모음

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

        #region 기타 메서드

        // 천마리 콤마, 소수점 버리기
        private string stringFormatN0(object obj)
        {
            return string.Format("{0:N0}", obj);
        }

        // 0이면 빈칸, 아니면, 천자리 콤마, 소수점 버리기
        private string stringFormatN0AndZeroEmpty(object obj)
        {
            string result = string.Format("{0:N0}", obj);

            if (result.Trim().Equals("0"))
            {
                result = "";
            }

            return result;
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

        private string MonthSlashDay(string str)
        {
            string result = "";

            if (str.Length == 8)
            {
                if (!str.Trim().Equals(""))
                {
                    result = str.Substring(4, 2) + "/" + str.Substring(6, 2);
                }
            }

            return result;
        }


        //특수문자 포함 검색
        private string Escape(string str)
        {
            string result = "";

            for (int i = 0; i < str.Length; i++)
            {
                string txt = str.Substring(i, 1);

                bool isSpecial = Regex.IsMatch(txt, @"[^a-zA-Z0-9가-힣]");

                if (isSpecial == true)
                {
                    result += (@"/" + txt);
                }
                else
                {
                    result += txt;
                }
            }
            return result;
        }



        #endregion // 기타 메서드

        private void DataGrid_SizeChange(object sender, SizeChangedEventArgs e)
        {
            DataGrid dgs = sender as DataGrid;
            if (dgs.ColumnHeaderHeight == 0)
            {
                dgs.ColumnHeaderHeight = 1;
            }
            double a = e.NewSize.Height / 100;
            double b = e.PreviousSize.Height / 100;
            double c = a / b;

            if (c != double.PositiveInfinity && c != 0 && double.IsNaN(c) == false)
            {
                dgs.ColumnHeaderHeight = dgs.ColumnHeaderHeight * c;
                dgs.FontSize = dgs.FontSize * c;
            }
        }


        private void btnBuyerArticleNoSrh_Click(object sender, RoutedEventArgs e)
        {
            pf.ReturnCode(txtBuyerArticleNoSrh, 84, txtBuyerArticleNoSrh.Text);
        }

        private void txtBuyerArticleNoSrh_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                pf.ReturnCode(txtBuyerArticleNoSrh, 84, txtBuyerArticleNoSrh.Text);
            }
        }


        private void CommonControl_Click(object sender, MouseButtonEventArgs e)
        {
            lib.CommonControl_Click(sender, e);
        }

        private void CommonControl_Click(object sender, RoutedEventArgs e)
        {
            lib.CommonControl_Click(sender, e);
        }
    }

    class Win_ord_Subul_Q_CodeView
    {
        public int Num { get; set; }

        public string cls { get; set; }

        public string BuyerArticleNo { get; set; }
        public string ArticleID { get; set; }
        public string Article { get; set; }
        public string ioDate { get; set; }
        public string LocID { get; set; }
        public string LocName { get; set; }

        public string Req_ID { get; set; }
        public string ReqName { get; set; }
        public string StuffRoll { get; set; }
        public string StuffQty { get; set; }
        public string UnitClss { get; set; }

        public string UnitClssName { get; set; }
        public string CustomID { get; set; }
        public string CustomName { get; set; }
        public string OrderID { get; set; }
        public string OrderNo { get; set; }

        public string OutRoll { get; set; }
        public string OutQty { get; set; }
        public string RelLocID { get; set; }
        public string RelLocName { get; set; }
        public string InOutClssName { get; set; }

        public string Remark { get; set; }
        public string StockQty { get; set; }

        public bool ArticleTotal_Color { get; set; }
        public bool Total_Color { get; set; }
    }

    public class Win_ord_Subul_Q_CodeView_Total : BaseView
    {
        public string TotalInitStockQty { get; set; }
        public string TotalStuffQty { get; set; }
        public string TotalOutQty { get; set; }
        public string TotalStockQty { get; set; }
    }
}

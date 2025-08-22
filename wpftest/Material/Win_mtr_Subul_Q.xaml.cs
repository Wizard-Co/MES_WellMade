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

namespace WizMes_WellMade
{
    /// <summary>
    /// Win_mtr_Subul_Q_New.xaml에 대한 상호 작용 논리
    /// </summary>
    public partial class Win_mtr_Subul_Q : UserControl
    {
        string stDate = string.Empty;
        string stTime = string.Empty;

        int rowNum = 0;
        Lib lib = new Lib();

        // 엑셀 활용 용도 (프린트)
        private Microsoft.Office.Interop.Excel.Application excelapp;
        private Microsoft.Office.Interop.Excel.Workbook workbook;
        private Microsoft.Office.Interop.Excel.Worksheet worksheet;
        private Microsoft.Office.Interop.Excel.Range workrange;
        private Microsoft.Office.Interop.Excel.Worksheet copysheet;
        private Microsoft.Office.Interop.Excel.Worksheet pastesheet;
        WizMes_WellMade.PopUp.NoticeMessage msg = new PopUp.NoticeMessage();

        // 인쇄를 위한 데이터 저장소
        List<Win_Subul> lstSubul = new List<Win_Subul>();

        ScrollViewer scrollView = null;
        ScrollViewer scrollView2 = null;

        int SearchCheckCount = 0;       // 검색 시 체크된 검색조건 갯수 파악. (2개 이상 체크시에만 )
        //2021-10-07
        double SumStockQty = 0; //현재 재고수량 
        double SumStuffinQty = 0; //입고 수량
        double SumOutQty = 0; //출고 수량

        //엑셀 타이틀 및 조건들어갈 배열
        string[] title = new string[11];

        public Win_mtr_Subul_Q()
        {
            InitializeComponent();
        }

        private void Win_sbl_Subul_Q_Loaded(object sender, RoutedEventArgs e)
        {
            stDate = DateTime.Now.ToString("yyyyMMdd");
            stTime = DateTime.Now.ToString("HHmm");

            DataStore.Instance.InsertLogByFormS(this.GetType().Name, stDate, stTime, "S");

            // 초기세팅
            chkInOutDate.IsChecked = true;
            dtpFromDate.SelectedDate = DateTime.Today;
            dtpToDate.SelectedDate = DateTime.Today;

            rbnManageNum.IsChecked = true; // 관리번호 : OrderID

            chkIn_NotApprovedIncloud.IsChecked = true; // 입고미승인건 포함
            chkAutoInOutItemsIncloud.IsChecked = true; // 자동입출건 포함

            // 콤보박스 세팅
            // 콤보박스 세팅
            ComboBoxSetting();
        }

        #region 콤보박스 세팅

        // 콤보박스 세팅.
        private void ComboBoxSetting()
        {
            cboArticleGroup.Items.Clear();
            cboWareHouse.Items.Clear();
            cboInGbn.Items.Clear();
            cboOutGbn.Items.Clear();
            cboSupplyType.Items.Clear();

            ObservableCollection<CodeView> cbArticleGroup = ComboBoxUtil.Instance.Gf_DB_MT_sArticleGrp();
            ObservableCollection<CodeView> cbWareHouse = ComboBoxUtil.Instance.Gf_DB_CM_GetComCodeDataset(null, "LOC", "Y", "", "");
            ObservableCollection<CodeView> cbInGbn = ComboBoxUtil.Instance.Gf_DB_CM_GetComCodeDataset(null, "ICD", "Y", "", "");
            ObservableCollection<CodeView> cbOutGbn = ComboBoxUtil.Instance.Gf_DB_CM_GetComCodeDataset(null, "OCD", "Y", "", "");
            ObservableCollection<CodeView> cbSupplyType = ComboBoxUtil.Instance.Gf_DB_CM_GetComCodeDataset(null, "CMMASPLTYPE", "Y", "", "");

            this.cboArticleGroup.ItemsSource = cbArticleGroup;
            this.cboArticleGroup.DisplayMemberPath = "code_name";
            this.cboArticleGroup.SelectedValuePath = "code_id";
            this.cboArticleGroup.SelectedIndex = 0;

            this.cboWareHouse.ItemsSource = cbWareHouse;
            this.cboWareHouse.DisplayMemberPath = "code_name";
            this.cboWareHouse.SelectedValuePath = "code_id";
            this.cboWareHouse.SelectedIndex = 0;

            this.cboInGbn.ItemsSource = cbInGbn;
            this.cboInGbn.DisplayMemberPath = "code_id_plus_code_name";
            this.cboInGbn.SelectedValuePath = "code_id";
            this.cboInGbn.SelectedIndex = 0;

            this.cboOutGbn.ItemsSource = cbOutGbn;
            this.cboOutGbn.DisplayMemberPath = "code_id_plus_code_name";
            this.cboOutGbn.SelectedValuePath = "code_id";
            this.cboOutGbn.SelectedIndex = 0;

            this.cboSupplyType.ItemsSource = cbSupplyType;
            this.cboSupplyType.DisplayMemberPath = "code_name";
            this.cboSupplyType.SelectedValuePath = "code_id";
            this.cboSupplyType.SelectedIndex = 0;

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
            //dtpFromDate.SelectedDate = DateTime.Today.AddDays(-1);
            //dtpToDate.SelectedDate = DateTime.Today.AddDays(-1);
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
            //dtpFromDate.SelectedDate = Lib.Instance.BringLastMonthDatetimeList()[0];
            //dtpToDate.SelectedDate = Lib.Instance.BringLastMonthDatetimeList()[1];
        }
        // 금월
        private void btnThisMonth_Click(object sender, RoutedEventArgs e)
        {
            dtpFromDate.SelectedDate = Lib.Instance.BringThisMonthDatetimeList()[0];
            dtpToDate.SelectedDate = Lib.Instance.BringThisMonthDatetimeList()[1];
        }

        // 제품그룹 
        private void chkArticleGroup_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            if (chkArticleGroup.IsChecked == true)
            {
                chkArticleGroup.IsChecked = false;
            }
            else
            {
                chkArticleGroup.IsChecked = true;
            }
        }
        private void chkArticleGroup_Checked(object sender, RoutedEventArgs e)
        {
            chkArticleGroup.IsChecked = true;
            cboArticleGroup.IsEnabled = true;
        }
        private void chkArticleGroup_Unchecked(object sender, RoutedEventArgs e)
        {
            chkArticleGroup.IsChecked = false;
            cboArticleGroup.IsEnabled = false;
        }

        // 거래처
        private void chkCustomer_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            if (chkCustomer.IsChecked == true)
            {
                chkCustomer.IsChecked = false;
            }
            else
            {
                chkCustomer.IsChecked = true;
            }
        }
        private void chkCustomer_Checked(object sender, RoutedEventArgs e)
        {
            chkCustomer.IsChecked = true;
            txtCustomer.IsEnabled = true;
            btnCustomer.IsEnabled = true;
        }
        private void chkCustomer_UnChecked(object sender, RoutedEventArgs e)
        {
            chkCustomer.IsChecked = false;
            txtCustomer.IsEnabled = false;
            btnCustomer.IsEnabled = false;
        }

        private void txtCustomer_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                e.Handled = true;
                MainWindow.pf.ReturnCode(txtCustomer, 0, "");
                //MainWindow.pf.ReturnCode(txtCustomer, 68, "");
            }
        }
        private void btnCustomer_Click(object sender, RoutedEventArgs e)
        {
            MainWindow.pf.ReturnCode(txtCustomer, 0, "");

        }

        // 품명
        private void chkArticle_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            if (chkArticle.IsChecked == true)
            {
                chkArticle.IsChecked = false;
            }
            else
            {
                chkArticle.IsChecked = true;
            }
        }
        private void chkArticle_Checked(object sender, RoutedEventArgs e)
        {
            chkArticle.IsChecked = true;
            txtArticle.IsEnabled = true;
            btnArticle.IsEnabled = true;
        }
        private void chkArticle_Unchecked(object sender, RoutedEventArgs e)
        {
            chkArticle.IsChecked = false;
            txtArticle.IsEnabled = false;
            btnArticle.IsEnabled = false;
        }
        private void txtArticle_KeyDown(object sender, KeyEventArgs e)
        {

            //if (e.Key == Key.Enter)
            //{
            //    rowNum = 0;
            //    re_Search(rowNum);
            //}
            if (e.Key == Key.Enter)
            {
                MainWindow.pf.ReturnCode(txtArticle, 76, "");
            }
        }
        private void btnArticle_Click(object sender, RoutedEventArgs e)
        {
            MainWindow.pf.ReturnCode(txtArticle, 76, "");
        }

        // 창고
        private void chkWareHouse_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            if (chkWareHouse.IsChecked == true)
            {
                chkWareHouse.IsChecked = false;
            }
            else
            {
                chkWareHouse.IsChecked = true;
            }
        }
        private void chkWareHouse_Checked(object sender, RoutedEventArgs e)
        {
            chkWareHouse.IsChecked = true;
            //cboWareHouse.IsEnabled = true;
        }
        private void chkWareHouse_Unchecked(object sender, RoutedEventArgs e)
        {
            chkWareHouse.IsChecked = false;
            cboWareHouse.IsEnabled = false;
        }

        // 관리번호
        private void chkManageNum_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            if (chkManageNum.IsChecked == true)
            {
                chkManageNum.IsChecked = false;
            }
            else
            {
                chkManageNum.IsChecked = true;
            }
        }
        private void chkManageNum_Checked(object sender, RoutedEventArgs e)
        {
            chkManageNum.IsChecked = true;
            txtManageNum.IsEnabled = true;
            btnManageNum.IsEnabled = true;
        }
        private void chkManageNum_Unchecked(object sender, RoutedEventArgs e)
        {
            chkManageNum.IsChecked = false;
            txtManageNum.IsEnabled = false;
            btnManageNum.IsEnabled = false;
        }
        private void txtManageNum_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                e.Handled = true;
                MainWindow.pf.ReturnCode(txtManageNum, (int)Defind_CodeFind.DCF_ORDER, "");
            }
        }
        private void btnManageNum_Click(object sender, RoutedEventArgs e)
        {
            MainWindow.pf.ReturnCode(txtManageNum, (int)Defind_CodeFind.DCF_ORDER, "");
        }

        // 발주번호
        private void chkOrderNum_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            if (chkOrderNum.IsChecked == true)
            {
                chkOrderNum.IsChecked = false;
            }
            else
            {
                chkOrderNum.IsChecked = true;
            }
        }
        private void chkOrderNum_Checked(object sender, RoutedEventArgs e)
        {
            chkOrderNum.IsChecked = true;
            txtOrderNum.IsEnabled = true;
            btnOrderNum.IsEnabled = true;
        }
        private void chkOrderNum_Unchecked(object sender, RoutedEventArgs e)
        {
            chkOrderNum.IsChecked = false;
            txtOrderNum.IsEnabled = false;
            btnOrderNum.IsEnabled = false;
        }
        private void txtOrderNum_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                e.Handled = true;
                MainWindow.pf.ReturnCode(txtOrderNum, (int)Defind_CodeFind.DCF_ORDER, "");
            }
        }
        private void btnOrderNum_Click(object sender, RoutedEventArgs e)
        {
            MainWindow.pf.ReturnCode(txtOrderNum, (int)Defind_CodeFind.DCF_ORDER, "");
        }

        // 입고구분
        private void chkInGbn_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            if (chkInGbn.IsChecked == true)
            {
                chkInGbn.IsChecked = false;
            }
            else
            {
                chkInGbn.IsChecked = true;
            }
        }
        private void chkInGbn_Checked(object sender, RoutedEventArgs e)
        {
            chkInGbn.IsChecked = true;
            cboInGbn.IsEnabled = true;
        }
        private void chkInGbn_Unchecked(object sender, RoutedEventArgs e)
        {
            chkInGbn.IsChecked = false;
            cboInGbn.IsEnabled = false;
        }

        // 출고구분
        private void chkOutGbn_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            if (chkOutGbn.IsChecked == true)
            {
                chkOutGbn.IsChecked = false;
            }
            else
            {
                chkOutGbn.IsChecked = true;
            }
        }
        private void chkOutGbn_Checked(object sender, RoutedEventArgs e)
        {
            chkOutGbn.IsChecked = true;
            cboOutGbn.IsEnabled = true;
        }
        private void chkOutGbn_Unchecked(object sender, RoutedEventArgs e)
        {
            chkOutGbn.IsChecked = false;
            cboOutGbn.IsEnabled = false;
        }

        // 공급유형
        private void chkSupplyType_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            if (chkSupplyType.IsChecked == true)
            {
                chkSupplyType.IsChecked = false;
            }
            else
            {
                chkSupplyType.IsChecked = true;
            }
        }
        private void chkSupplyType_Checked(object sender, RoutedEventArgs e)
        {
            chkSupplyType.IsChecked = true;
            cboSupplyType.IsEnabled = true;
        }
        private void chkSupplyType_Unchecked(object sender, RoutedEventArgs e)
        {
            chkSupplyType.IsChecked = false;
            cboSupplyType.IsEnabled = true;
        }

        #endregion // Header 부분 - 검색조건

        #region Header 부분 - 입출고 근거번호 숨김 기능

        // 입출고 근거번호 숨김
        private void chkHideInOutReasonNumber_Click(object sender, RoutedEventArgs e)
        {
            if (chkHideInOutReasonNumber.IsChecked == true)
            {
                // 숨김버튼 체크 → 발주번호, 오더번호 안보이게 설정
                Req_ID.Visibility = Visibility.Hidden;
                OrderID.Visibility = Visibility.Hidden;

            }
            else
            {
                // 숨김버튼 체크 해제 → 발주번호, 오더번호 보이게 설정
                Req_ID.Visibility = Visibility.Visible;
                OrderID.Visibility = Visibility.Visible;

            }
        }

        #endregion // Header 부분 - 입출고 근거번호 숨김 기능

        #region Header 부분 - 상단 오른쪽 버튼

        // 검색 버튼 이벤트
        private void btnSearch_Click(object sender, RoutedEventArgs e)
        {
            rowNum = 0;
            re_Search(rowNum);
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
            excelapp = new Microsoft.Office.Interop.Excel.Application();

            string MyBookPath = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetEntryAssembly().Location) + "\\Report\\자재수불내역.xls";
            workbook = excelapp.Workbooks.Add(MyBookPath);
            worksheet = workbook.Sheets["Form"];
            pastesheet = workbook.Sheets["Print"];

            // 창고로 검색했다면 창고 입력 : D3
            workrange = worksheet.get_Range("D3");
            workrange.Value2 = chkWareHouse.IsChecked == true && cboWareHouse.SelectedValue != null ? cboWareHouse.SelectedValue.ToString() : "";

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
            int nextCopyLine = 46;

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
                        workrange.Value2 = Subul.BuyerArticleNo;

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

        #endregion

        // 엑셀 버튼 이벤트
        private void btnExcel_Click(object sender, RoutedEventArgs e)
        {
            DataTable dt = null;
            //string Name = string.Empty; Name 대신 title 사용

            string[] lst = new string[2];
            lst[0] = "수불내역";
            lst[1] = dgdMain.Name;

            dgdMain.Tag = lst[0];

            //검색조건대로 타이틀 넣어줌
            SearchCheckToTitle();

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

                    //Name = dgdMain.Name; 대신 title 사용

                    //타이틀  0번은 엑셀 제목 또는 시트이름           
                    title[0] = dgdMain.Tag.ToString();

                    if (Lib.Instance.TitleAddGenerateExcel(dt, title))
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

        #endregion // Header 부분 - 상단 오른쪽 버튼

        #region 주요 메서드 모음

        private void re_Search(int selectedIndex)
        {
            if (chkManageNum.IsChecked == true && txtManageNum.Text == "")
            {
                MessageBox.Show("관리번호를 입력한 후 검색을 하거나 관리번호 체크를 해제 후 검색 하세요");
                return;
            }
            else if (chkOrderNum.IsChecked == true && txtOrderNum.Text == "")
            {
                MessageBox.Show("발주번호를 입력한 후 검색을 하거나 발주번호 체크를 해제 후 검색 하세요");
                return;
            }
            else if (chkCustomer.IsChecked == true && txtCustomer.Text == "")
            {
                MessageBox.Show("거래처를 입력한 후 검색을 하거나 거래처 체크를 해제 후 검색 하세요");
                return;
            }
            else if (chkArticle.IsChecked == true && txtArticle.Text == "")
            {
                MessageBox.Show("품번를 입력한 후 검색을 하거나 품번 체크를 해제 후 검색 하세요");
                return;
            }

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

        #region 조회 메서드

        private void FillGrid()
        {
            SumStockQty = 0;
            SumStuffinQty = 0;
            SumOutQty = 0;

            lstSubul.Clear();

            if (dgdMain.Items.Count > 0)
            {
                dgdMain.Items.Clear();
            }

            try
            {
                //관리번호 1 오더번호 2 // 아무것도 체크안하면 0  ssw
                Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                sqlParameter.Clear();

                sqlParameter.Add("nChkDate", chkInOutDate.IsChecked == true ? 1 : 0);
                sqlParameter.Add("sSDate", chkInOutDate.IsChecked == true && dtpFromDate.SelectedDate != null ? dtpFromDate.SelectedDate.Value.ToString("yyyyMMdd") : "");
                sqlParameter.Add("sEDate", chkInOutDate.IsChecked == true && dtpToDate.SelectedDate != null ? dtpToDate.SelectedDate.Value.ToString("yyyyMMdd") : "");
                sqlParameter.Add("nChkCustom", chkCustomer.IsChecked == true ? 1 : 0);
                sqlParameter.Add("sCustomID", chkCustomer.IsChecked == true && txtCustomer.Tag != null ? txtCustomer.Tag.ToString() : "");

                sqlParameter.Add("nChkArticleID", 0);
                sqlParameter.Add("sArticleID", "");
                sqlParameter.Add("nChkOrder", chkManageNum.IsChecked == true ? 2 : 0); //이 친구는 주석이라서 무슨 값이 들어가도...

                sqlParameter.Add("sOrder", txtManageNum.Text);
                sqlParameter.Add("ArticleGrpID", chkArticleGroup.IsChecked == true && cboArticleGroup.SelectedValue != null ? cboArticleGroup.SelectedValue.ToString() : "");

                sqlParameter.Add("sFromLocID", chkWareHouse.IsChecked == true && cboWareHouse.SelectedValue != null ? cboWareHouse.SelectedValue.ToString() : "");
                sqlParameter.Add("sToLocID", "");
                sqlParameter.Add("nChkOutClss", chkOutGbn.IsChecked == true ? 1 : 0);
                sqlParameter.Add("sOutClss", chkOutGbn.IsChecked == true && cboOutGbn.SelectedValue != null ? cboOutGbn.SelectedValue.ToString() : "");
                sqlParameter.Add("nChkInClss", chkInGbn.IsChecked == true ? 1 : 0);

                sqlParameter.Add("sInClss", chkInGbn.IsChecked == true && cboInGbn.SelectedValue != null ? cboInGbn.SelectedValue.ToString() : "");
                sqlParameter.Add("nChkReqID", chkOrderNum.IsChecked == true ? 1 : 0);
                sqlParameter.Add("sReqID", chkOrderNum.IsChecked == true && txtOrderNum.Tag != null ? txtOrderNum.Tag.ToString() : "");
                sqlParameter.Add("incNotApprovalYN", chkIn_NotApprovedIncloud.IsChecked == true ? "Y" : "N");
                sqlParameter.Add("incAutoInOutYN", chkAutoInOutItemsIncloud.IsChecked == true ? "Y" : "N");

                sqlParameter.Add("sProductYN", "");    // 이유불문 무조건 Y인JaturiNoYN데..;;
                sqlParameter.Add("nMainItem", chkMainInterestItemsSee.IsChecked == true ? 1 : 0);
                sqlParameter.Add("nCustomItem", chkRegistItemsByCustomer.IsChecked == true ? 1 : 0);
                sqlParameter.Add("nSupplyType", chkSupplyType.IsChecked == true ? 1 : 0);
                sqlParameter.Add("sSupplyType", chkSupplyType.IsChecked == true && cboSupplyType.SelectedValue != null ? cboSupplyType.SelectedValue.ToString() : "");

                sqlParameter.Add("nBuyerArticleNo", chkArticle.IsChecked == true ? 1 : 0);
                sqlParameter.Add("BuyerArticleNo", chkArticle.IsChecked == true && !txtArticle.Text.Trim().Equals("") ? @Escape(txtArticle.Text) : "");

                DataSet ds = DataStore.Instance.ProcedureToDataSet("xp_Subul_sSubul_MTR", sqlParameter, true);

                if (ds != null && ds.Tables.Count > 0)
                {
                    DataTable dt = ds.Tables[0];

                    if (dt.Rows.Count > 0)
                    {
                        int i = 0;
                        DataRowCollection drc = dt.Rows;

                        foreach (DataRow dr in drc)
                        {
                            i++;
                            var Subul = new Win_Subul()
                            {
                                Num = i,

                                cls = dr["cls"].ToString(),

                                ArticleID = dr["ArticleID"].ToString(),
                                Article = dr["Article"].ToString(),
                                BuyerArticleNo = dr["BuyerArticleNo"].ToString(),
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
                                StockQty = stringFormatN0(dr["StockQty"]),

                            };

                            if (dr["cls"].ToString().Trim().Equals("0")) // 이월
                            {
                                Subul.ioDate = "이월";
                                Subul.StockQty = stringFormatN0(dr["StuffQty"]);                      //2021-10-08 재고량이 바로 보이게 하기 위해 추가
                                SumStockQty = SumStockQty + ConvertDouble(dr["StuffQty"].ToString()); //2021-10-08 재고량을 사용하기 위해 변수에 저장

                            }
                            else if (dr["cls"].ToString().Trim().Equals("6")) // 품명계
                            {
                                Subul.Article = "";
                                Subul.BuyerArticleNo = "품명계";
                                Subul.ioDate = "";
                                Subul.LocName = "";
                                Subul.ArticleTotal_Color = true;
                                SumStockQty = 0; //2021-10-08 0으로 변경해야 다른 품명의 재고가 맞게 나옴
                            }
                            else if (dr["cls"].ToString().Trim().Equals("7")) // 총계
                            {
                                Subul.Article = "";
                                Subul.BuyerArticleNo = "총계";
                                Subul.ioDate = "";
                                Subul.LocName = "";
                                Subul.Total_Color = true;
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
                            //else // 그냥 입고, 출고 데이터를 인쇄를 위해 따로 저장
                            //{
                            //    lstSubul.Add(Subul);
                            //}
                            lstSubul.Add(Subul); // 엑셀 출력 할 떄 리스트에 못 담아서 오류떠서 다 뿌리고나면 넣어줘야함
                            dgdMain.Items.Add(Subul);
                        }

                    }
                }

                //검색조건대로 타이틀 넣어줌
                //SearchCheckToTitle();
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

        //엑셀에 들어갈 조건 내용들
        private void SearchCheckToTitle()
        {
            if (chkInOutDate.IsChecked == true && dtpFromDate.SelectedDate != null && dtpToDate.SelectedDate != null)
            {
                title[1] = "   ▶ 기준일자 : " + dtpFromDate.SelectedDate.ToString().Substring(0, 10) + " ~ " + dtpToDate.SelectedDate.ToString().Substring(0, 10);
            }

            if (chkArticleGroup.IsChecked == true && cboArticleGroup.SelectedValue != null)
            {
                title[2] = "   ▶ 제품그룹 : " + cboArticleGroup.Text.ToString();
            }

            if (chkCustomer.IsChecked == true && txtCustomer.Tag != null)
            {
                title[3] = "   ▶ 거래처 : " + txtCustomer.Text.ToString();
            }

            if (chkArticle.IsChecked == true && !txtArticle.Text.Trim().Equals(""))
            {
                title[4] = "   ▶ 품번 : " + @Escape(txtArticle.Text);
            }

            if (chkWareHouse.IsChecked == true && cboWareHouse.SelectedValue != null)
            {
                title[5] = "   ▶ 창고 : " + cboWareHouse.Text.ToString();
            }

            if (chkOrderNum.IsChecked == true && txtOrderNum.Tag != null)
            {
                title[6] = "   ▶ 발주번호 : " + txtOrderNum.Tag.ToString();
            }

            if (chkInGbn.IsChecked == true && cboInGbn.SelectedValue != null)
            {
                title[7] = "   ▶ 입고구분 : " + cboInGbn.Text.ToString();
            }

            if (chkOutGbn.IsChecked == true && cboOutGbn.SelectedValue != null)
            {
                title[8] = "   ▶ 출고구분 : " + cboOutGbn.Text.ToString();
            }
            if (chkSupplyType.IsChecked == true && cboSupplyType.SelectedValue != null)
            {
                title[9] = "   ▶ 공급유형 : " + cboSupplyType.Text.ToString();
            }

        }
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

        #endregion // 기타 메서드
    }

    class Win_Subul
    {
        public int Num { get; set; }

        public string cls { get; set; }

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
        public string BuyerArticleNo { get; set; }

    }
}

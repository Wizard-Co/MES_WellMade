using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Data;
using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using WizMes_WellMade.PopUP;

namespace WizMes_WellMade
{
    /// <summary>
    /// Win_mtr_Stuffin_Q_New.xaml에 대한 상호 작용 논리
    /// </summary>
    public partial class Win_mtr_Stuffin_Q : UserControl
    {
        string stDate = string.Empty;
        string stTime = string.Empty;

        private int rowNum = 0;
        Lib lib = new Lib();

        // 엑셀 활용 용도 (프린트)
        private Microsoft.Office.Interop.Excel.Application excelapp;
        private Microsoft.Office.Interop.Excel.Workbook workbook;
        private Microsoft.Office.Interop.Excel.Worksheet worksheet;
        private Microsoft.Office.Interop.Excel.Range workrange;
        private Microsoft.Office.Interop.Excel.Worksheet copysheet;
        private Microsoft.Office.Interop.Excel.Worksheet pastesheet;
        //(기다림 알림 메시지창)
        WizMes_WellMade.PopUp.NoticeMessage msg = new PopUp.NoticeMessage();

        public Win_mtr_Stuffin_Q()
        {
            InitializeComponent();
        }

        // 폼 로드
        private void UserControl_Loaded(object sender, RoutedEventArgs e)
        {
            stDate = DateTime.Now.ToString("yyyyMMdd");
            stTime = DateTime.Now.ToString("HHmm");

            DataStore.Instance.InsertLogByFormS(this.GetType().Name, stDate, stTime, "S");

            // 발주별 토글버튼 클릭
            tgnCustomSort.IsChecked = true;

            // 입고일자 오늘 세팅
            chkDateSrh.IsChecked = true;
            dtpSDateSrh.SelectedDate = DateTime.Today;
            dtpEDateSrh.SelectedDate = DateTime.Today;

            setComboBox();
        }

        // 콤보박스 세팅
        private void setComboBox()
        {
            // 제품그룹
            ObservableCollection<CodeView> ovcArticleGrpID = ComboBoxUtil.Instance.GetArticleCode_SetComboBox("", 0);
            cboArticleGrpSrh.ItemsSource = ovcArticleGrpID;
            cboArticleGrpSrh.DisplayMemberPath = "code_name";
            cboArticleGrpSrh.SelectedValuePath = "code_id";

            // 입고구분
            ObservableCollection<CodeView> ovcStuff = ComboBoxUtil.Instance.Gf_DB_CM_GetComCodeDataset(null, "ICD", "Y", "", "MTR");
            cboStuffClssSrh.ItemsSource = ovcStuff;
            cboStuffClssSrh.DisplayMemberPath = "code_name";
            cboStuffClssSrh.SelectedValuePath = "code_id";

            // 입고창고
            ObservableCollection<CodeView> cbFromToLoc = ComboBoxUtil.Instance.Gf_DB_CM_GetComCodeDataset(null, "LOC", "Y", "", "");
            this.cboToLocSrh.ItemsSource = cbFromToLoc;
            this.cboToLocSrh.DisplayMemberPath = "code_name";
            this.cboToLocSrh.SelectedValuePath = "code_id";
            this.cboToLocSrh.SelectedIndex = 0;

            // 입고검수구분
            List<string[]> strValueYN = new List<string[]>();
            strValueYN.Add(new string[] { "Y", "Y" });
            strValueYN.Add(new string[] { "N", "N" });

            ObservableCollection<CodeView> ovcYN = ComboBoxUtil.Instance.Direct_SetComboBox(strValueYN);
            this.cbosInspectApprovalYN.ItemsSource = ovcYN;
            this.cbosInspectApprovalYN.DisplayMemberPath = "code_name";
            this.cbosInspectApprovalYN.SelectedValuePath = "code_id";
        }

        #region Header 부분 - 검색 조건

        // 일자 검색
        private void lblDateSrh_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
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
        // 일자 검색 체크박스
        private void chkDateSrh_Checked(object sender, RoutedEventArgs e)
        {
            chkDateSrh.IsChecked = true;

            dtpSDateSrh.IsEnabled = true;
            dtpEDateSrh.IsEnabled = true;
        }
        private void chkDateSrh_Unchecked(object sender, RoutedEventArgs e)
        {
            chkDateSrh.IsChecked = false;

            dtpSDateSrh.IsEnabled = false;
            dtpEDateSrh.IsEnabled = false;
        }
        // 전일
        private void btnYesterday_Click(object sender, RoutedEventArgs e)
        {
            //dtpSDateSrh.SelectedDate = DateTime.Today.AddDays(-1);
            //dtpEDateSrh.SelectedDate = DateTime.Today.AddDays(-1);
            try
            {
                if (dtpSDateSrh.SelectedDate != null)
                {
                    dtpSDateSrh.SelectedDate = dtpSDateSrh.SelectedDate.Value.AddDays(-1);
                    dtpEDateSrh.SelectedDate = dtpSDateSrh.SelectedDate;
                }
                else
                {
                    dtpSDateSrh.SelectedDate = DateTime.Today.AddDays(-1);
                    dtpEDateSrh.SelectedDate = DateTime.Today.AddDays(-1);
                }
            }
            catch (Exception ee)
            {
                MessageBox.Show("오류지점 - btnYesterday_Click : " + ee.ToString());
            }
        }
        // 금일
        private void btnToday_Click(object sender, RoutedEventArgs e)
        {
            dtpSDateSrh.SelectedDate = DateTime.Today;
            dtpEDateSrh.SelectedDate = DateTime.Today;
        }
        // 전월
        private void btnLastMonth_Click(object sender, RoutedEventArgs e)
        {
            //dtpSDateSrh.SelectedDate = Lib.Instance.BringLastMonthDatetimeList()[0];
            //dtpEDateSrh.SelectedDate = Lib.Instance.BringLastMonthDatetimeList()[1];
            try
            {
                if (dtpSDateSrh.SelectedDate != null)
                {
                    DateTime ThatMonth1 = dtpSDateSrh.SelectedDate.Value.AddDays(-(dtpSDateSrh.SelectedDate.Value.Day - 1)); // 선택한 일자 달의 1일!

                    DateTime LastMonth1 = ThatMonth1.AddMonths(-1); // 저번달 1일
                    DateTime LastMonth31 = ThatMonth1.AddDays(-1); // 저번달 말일

                    dtpSDateSrh.SelectedDate = LastMonth1;
                    dtpEDateSrh.SelectedDate = LastMonth31;
                }
                else
                {
                    DateTime ThisMonth1 = DateTime.Today.AddDays(-(DateTime.Today.Day - 1)); // 이번달 1일

                    DateTime LastMonth1 = ThisMonth1.AddMonths(-1); // 저번달 1일
                    DateTime LastMonth31 = ThisMonth1.AddDays(-1); // 저번달 말일

                    dtpSDateSrh.SelectedDate = LastMonth1;
                    dtpEDateSrh.SelectedDate = LastMonth31;
                }
            }
            catch (Exception ee)
            {
                MessageBox.Show("오류지점 - btnLastMonth_Click : " + ee.ToString());
            }

        }
        // 금월
        private void btnThisMonth_Click(object sender, RoutedEventArgs e)
        {
            dtpSDateSrh.SelectedDate = Lib.Instance.BringThisMonthDatetimeList()[0];
            dtpEDateSrh.SelectedDate = Lib.Instance.BringThisMonthDatetimeList()[1];
        }

        // 품명그룹
        private void lblArticleGrpSrh_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            if (chkArticleGrpSrh.IsChecked == true)
            {
                chkArticleGrpSrh.IsChecked = false;
            }
            else
            {
                chkArticleGrpSrh.IsChecked = true;
            }
        }
        private void chkArticleGrpSrh_Checked(object sender, RoutedEventArgs e)
        {
            //chkArticleGrpSrh.IsChecked = true;

            //cboArticleGrpSrh.IsEnabled = true;
        }
        private void chkArticleGrpSrh_Unchecked(object sender, RoutedEventArgs e)
        {
            //chkArticleGrpSrh.IsChecked = false;

            //cboArticleGrpSrh.IsEnabled = false;
        }

        // 품명
        private void lblArticleSrh_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            if (chkArticleSrh.IsChecked == true)
            {
                chkArticleSrh.IsChecked = false;
            }
            else
            {
                chkArticleSrh.IsChecked = true;
            }
        }

        private void chkArticleSrh_Checked(object sender, RoutedEventArgs e)
        {
            chkArticleSrh.IsChecked = true;

            txtArticleSrh.IsEnabled = true;
            btnPfArticleSrh.IsEnabled = true;
        }

        private void chkArticleSrh_Unchecked(object sender, RoutedEventArgs e)
        {
            chkArticleSrh.IsChecked = false;

            txtArticleSrh.IsEnabled = false;
            btnPfArticleSrh.IsEnabled = false;
        }
        // 품명 검색 엔터 → 플러스 파인더 이벤트
        private void txtArticleSrh_KeyDown(object sender, KeyEventArgs e)
        {

            //if (e.Key == Key.Enter)
            //{
            //    rowNum = 0;
            //    re_Search(rowNum);
            //}

            if (e.Key == Key.Enter)
            {
                e.Handled = true;
                MainWindow.pf.ReturnCode(txtArticleSrh, (int)Defind_CodeFind.DCF_BuyerArticleNo, "");
            }
        }
        private void btnPfArticleSrh_Click(object sender, RoutedEventArgs e)
        {
            MainWindow.pf.ReturnCode(txtArticleSrh, (int)Defind_CodeFind.DCF_BuyerArticleNo, "");
        }

        // 구매거래처
        private void lblBuyCustomSrh_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            if (chkBuyCustomSrh.IsChecked == true)
            {
                chkBuyCustomSrh.IsChecked = false;
            }
            else
            {
                chkBuyCustomSrh.IsChecked = true;
            }
        }
        private void chkBuyCustomSrh_Checked(object sender, RoutedEventArgs e)
        {
            chkBuyCustomSrh.IsChecked = true;

            txtBuyCustomSrh.IsEnabled = true;
            btnPfBuyCustomSrh.IsEnabled = true;
        }
        private void chkBuyCustomSrh_Unchecked(object sender, RoutedEventArgs e)
        {
            chkBuyCustomSrh.IsChecked = false;

            txtBuyCustomSrh.IsEnabled = false;
            btnPfBuyCustomSrh.IsEnabled = false;
        }
        // 거래처 검색 엔터 → 플러스 파인더 이벤트
        private void txtBuyCustomSrh_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                e.Handled = true;
                MainWindow.pf.ReturnCode(txtBuyCustomSrh, (int)Defind_CodeFind.DCF_CUSTOM, "");
            }
        }
        private void btnPfBuyCustomSrh_Click(object sender, RoutedEventArgs e)
        {
            MainWindow.pf.ReturnCode(txtBuyCustomSrh, (int)Defind_CodeFind.DCF_CUSTOM, "");
        }

        // 입고구분
        private void lblStuffClssSrh_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            if (chkStuffClssSrh.IsChecked == true)
            {
                chkStuffClssSrh.IsChecked = false;
            }
            else
            {
                chkStuffClssSrh.IsChecked = true;
            }
        }
        // 입고구분 검색 체크박스 이벤트
        private void chkStuffClssSrh_Checked(object sender, RoutedEventArgs e)
        {
            chkStuffClssSrh.IsChecked = true;

            cboStuffClssSrh.IsEnabled = true;
        }
        private void chkStuffClssSrh_Unchecked(object sender, RoutedEventArgs e)
        {
            chkStuffClssSrh.IsChecked = false;

            cboStuffClssSrh.IsEnabled = false;
        }

        // 입고창고
        private void lblToLocSrh_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            if (chkToLocSrh.IsChecked == true)
            {
                chkToLocSrh.IsChecked = false;
            }
            else
            {
                chkToLocSrh.IsChecked = true;
            }
        }
        private void chkToLocSrh_Checked(object sender, RoutedEventArgs e)
        {
            chkToLocSrh.IsChecked = true;

            cboToLocSrh.IsEnabled = true;
        }
        private void chkToLocSrh_Unchecked(object sender, RoutedEventArgs e)
        {
            chkToLocSrh.IsChecked = false;

            cboToLocSrh.IsEnabled = false;
        }

        // 입고검수구분
        private void lblsInspectApprovalYN_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            if (chksInspectApprovalYN.IsChecked == true)
            {
                chksInspectApprovalYN.IsChecked = false;
            }
            else
            {
                chksInspectApprovalYN.IsChecked = true;
            }
        }
        private void chksInspectApprovalYN_Checked(object sender, RoutedEventArgs e)
        {
            chksInspectApprovalYN.IsChecked = true;

            cbosInspectApprovalYN.IsEnabled = true;
        }

        private void chksInspectApprovalYN_Unchecked(object sender, RoutedEventArgs e)
        {
            chksInspectApprovalYN.IsChecked = false;

            cbosInspectApprovalYN.IsEnabled = false;
        }


        #endregion // Header 부분 - 검색조건

        #region 상단 오른쪽 버튼 모음

        // 검색
        private void btnSearch_Click(object sender, RoutedEventArgs e)
        {


            btnSearch.IsEnabled = false;

            //Dispatcher.BeginInvoke(new Action(() =>
            //{
            //    Thread.Sleep(2000);
            //    MessageBox.Show("조회중");
            //}), System.Windows.Threading.DispatcherPriority.Background);

            //Dispatcher.BeginInvoke(new Action(() =>
            //{

            //    btnSearch.IsEnabled = true;
            //}), System.Windows.Threading.DispatcherPriority.Background);


            if (CheckData())
            {
                rowNum = 0;
                re_Search(rowNum);
            }


        }
        // 닫기
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
            if (dgdMainReq.Items.Count < 1)
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
            if (dgdMainReq.Items.Count < 1)
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

            string MyBookPath = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetEntryAssembly().Location) + "\\Report\\자재입고명세서.xls";
            workbook = excelapp.Workbooks.Add(MyBookPath);
            worksheet = workbook.Sheets["Form"];
            pastesheet = workbook.Sheets["Print"];

            // 구매거래처로 검색했다면 구매거래처 입력 : D3
            workrange = worksheet.get_Range("D3");
            workrange.Value2 = chkBuyCustomSrh.IsChecked == true ? txtBuyCustomSrh.Text : "";

            // 기간 설정후에 검색했다면 검색일자 2019.10.01 ~ 2019.10.31 입력 : D4
            workrange = worksheet.get_Range("D4");
            workrange.Value2 = chkDateSrh.IsChecked == true ? dtpSDateSrh.SelectedDate.Value.ToString("yyyy-MM-dd") + " ~ " + dtpEDateSrh.SelectedDate.Value.ToString("yyyy-MM-dd") : "";

            // 일자 입력(오늘일자) : AE4
            workrange = worksheet.get_Range("AB4");
            workrange.Value2 = DateTime.Today.ToString("yyyy-MM-dd");

            // 페이지 계산 등
            int rowCount = dgdMainReq.Items.Count;
            int excelStartRow = 7;

            // 총 데이터를 입력할수 있는 갯수
            int totalDataInput = 37;

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

                // 기존에 있는 데이터 지우기 "A7", "AG43"
                worksheet.Range["A7", "AG43"].EntireRow.ClearContents();

                for (int i = DataCount; i < rowCount; i++)
                {
                    if (i == totalDataInput * Page)
                    {
                        break;
                    }

                    var Stuffin = dgdMainReq.Items[i] as Win_mtr_ocStuffIN_U_CodeView;
                    int excelRow = excelStartRow + excelNum;

                    if (Stuffin != null)
                    {
                        //일자
                        workrange = worksheet.get_Range("A" + excelRow);
                        workrange.Value2 = Stuffin.StuffDate_CV;

                        // 거래처명
                        workrange = worksheet.get_Range("E" + excelRow);
                        workrange.Value2 = Stuffin.CustomName;

                        // 품명
                        workrange = worksheet.get_Range("K" + excelRow);
                        workrange.Value2 = Stuffin.Article;

                        // 모델
                        workrange = worksheet.get_Range("Q" + excelRow);
                        workrange.Value2 = Stuffin.Req_ID;

                        // 수량
                        workrange = worksheet.get_Range("V" + excelRow);
                        workrange.Value2 = Stuffin.StuffQty;

                        // LotID
                        workrange = worksheet.get_Range("Z" + excelRow);
                        workrange.Value2 = DatePickerFormat(Stuffin.InspectDate);

                        // 비고
                        workrange = worksheet.get_Range("AD" + excelRow);
                        workrange.Value2 = Stuffin.Remark;

                        //// 검사자
                        //workrange = worksheet.get_Range("AF" + excelRow);
                        //workrange.Value2 = Stuffin.Inspector;

                        //SumAmount += ConvertDouble(OcReqSub.Amount);

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

            //// 총금액 입력하기 : 10, 50, 90
            //for (int i = 0; i < PageAll; i++)
            //{
            //    int sumAmount_Index = 10 + (40 * i);

            //    workrange = pastesheet.get_Range("E" + sumAmount_Index);
            //    workrange.Value2 = SumAmount;
            //}

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
        // 엑셀
        private void btnExcel_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                DataGrid dgd = new DataGrid();
                if (dgdMainReq.Visibility == Visibility.Visible)
                {
                    dgd = dgdMainReq;
                }
                else
                {
                    dgd = dgdMainCustom;
                }

                DataTable dt = null;
                string Name = string.Empty;

                string[] lst = new string[2];
                lst[0] = "자재 입고 목록";
                lst[1] = dgd.Name;

                ExportExcelxaml ExpExc = new ExportExcelxaml(lst);

                ExpExc.ShowDialog();

                if (ExpExc.DialogResult.HasValue)
                {
                    if (ExpExc.choice.Equals(dgd.Name))
                    {
                        DataStore.Instance.InsertLogByForm(this.GetType().Name, "E");
                        if (ExpExc.Check.Equals("Y"))
                            dt = Lib.Instance.DataGridToDTinHidden(dgd);
                        else
                            dt = Lib.Instance.DataGirdToDataTable(dgd);

                        Name = dgd.Name;

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

        #endregion // 상단 오른쪽 버튼 모음


        #region Content 부분

        // 발주별 버튼 이벤트
        private void tgnReqSort_Click(object sender, RoutedEventArgs e)
        {
            // 발주별 버튼 Checked
            tgnReqSort.IsChecked = true;
            // 발주 그리드 활성화
            dgdMainReq.Visibility = Visibility.Visible;

            // 거래처별 버튼 UnChecked
            tgnCustomSort.IsChecked = false;
            // 거래처 그리드 비활성화
            dgdMainCustom.Visibility = Visibility.Hidden;

            // 그리드 헤더가 안맞아서, 일단 다시 검색 되도록
            //rowNum = 0;
            //re_Search(rowNum);
        }

        // 거래처별 버튼 이벤트
        private void tgnCustomSort_Click(object sender, RoutedEventArgs e)
        {
            // 거래처별 버튼 Checked
            tgnCustomSort.IsChecked = true;
            // 거래처 그리드 활성화
            dgdMainCustom.Visibility = Visibility.Visible;

            // 발주별 버튼 UnChecked
            tgnReqSort.IsChecked = false;
            // 발주 그리드 비활성화
            dgdMainReq.Visibility = Visibility.Hidden;

            // 그리드 헤더가 안맞아서, 일단 다시 검색 되도록
            //rowNum = 0;
            //re_Search(rowNum);
        }

        #endregion // Content 부분

        #region 주요 메서드

        // 조회
        private void re_Search(int rowNum)
        {
            FillGrid();

            if (dgdMainReq.Items.Count > 0 || dgdMainCustom.Items.Count > 0)
            {
                dgdMainReq.SelectedIndex = rowNum;
                dgdMainCustom.SelectedIndex = rowNum;
            }
            else
            {
                this.DataContext = null;
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


        #region 조회

        // 조회
        private void FillGrid()
        {
            // 입고량, 입고건수 - 합계 구하기
            var SumStuffIN = new Win_mtr_ocStuffIN_Sum();

            if (dgdMainReq.Items.Count > 0 || dgdMainCustom.Items.Count > 0)
            {
                dgdMainReq.Items.Clear();
                dgdMainCustom.Items.Clear();
                dgdTotal.Items.Clear();
            }

            try
            {
                Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                sqlParameter.Clear();

                sqlParameter.Add("nChkDate", chkDateSrh.IsChecked == true ? 1 : 0);
                sqlParameter.Add("sSDate", dtpSDateSrh.SelectedDate != null ? dtpSDateSrh.SelectedDate.Value.ToString("yyyyMMdd") : "");
                sqlParameter.Add("sEDate", dtpEDateSrh.SelectedDate != null ? dtpEDateSrh.SelectedDate.Value.ToString("yyyyMMdd") : "");
                sqlParameter.Add("nChkCustom", 0);
                sqlParameter.Add("sCustom", "");

                sqlParameter.Add("nChkArticleID", chkArticleNoSrh.IsChecked == true ? 1 : 0);
                sqlParameter.Add("sArticleID", txtArticleNo.Tag != null && !txtArticleNo.Text.Trim().Equals("") ? txtArticleNo.Tag.ToString() : "");
                //sqlParameter.Add("sArticleID", chkArticleSrh.IsChecked == true && txtArticleNo.Tag != null ? txtArticleNo.Tag.ToString() : ""); //품명코드 검색
                sqlParameter.Add("nChkStuffClss", chkStuffClssSrh.IsChecked == true ? 1 : 0);
                sqlParameter.Add("sStuffClss", cboStuffClssSrh.SelectedValue != null ? cboStuffClssSrh.SelectedValue.ToString() : "");
                sqlParameter.Add("nChkIncStuffIN", 0);

                sqlParameter.Add("nChkArticleGrp", chkArticleGrpSrh.IsChecked == true ? 1 : 0);
                sqlParameter.Add("ArticleGrpID", cboArticleGrpSrh.SelectedValue != null ? cboArticleGrpSrh.SelectedValue.ToString() : "");
                sqlParameter.Add("chkInspect", chksInspectApprovalYN.IsChecked == true ? 1 : 0);      // 입고 검수
                sqlParameter.Add("sInspect", cbosInspectApprovalYN.SelectedValue != null ? cbosInspectApprovalYN.SelectedValue.ToString() : "");

                sqlParameter.Add("nChkBuyCustom", chkBuyCustomSrh.IsChecked == true ? 1 : 0); // 발주별, 거래처별 정렬
                sqlParameter.Add("sBuyCustom", chkBuyCustomSrh.IsChecked == true && txtBuyCustomSrh.Tag != null ? txtBuyCustomSrh.Tag.ToString() : ""); // 발주별, 거래처별 정렬

                sqlParameter.Add("OrderrByClss", ""); // 발주별, 거래처별 정렬
                sqlParameter.Add("sInspectBasisID", ""); // 발주별, 거래처별 정렬

                sqlParameter.Add("sToLocID", chkToLocSrh.IsChecked == true && cboToLocSrh.SelectedValue != null ? cboToLocSrh.SelectedValue.ToString() : "");

                sqlParameter.Add("nBuyArticleNo", chkArticleSrh.IsChecked == true ? 1 : 0);
                sqlParameter.Add("BuyArticleNo", chkArticleSrh.IsChecked == true && !txtArticleSrh.Text.Trim().Equals("") ? txtArticleSrh.Text : "");

                DataSet ds = DataStore.Instance.ProcedureToDataSet_LogWrite("xp_StuffIN_sStuffIN", sqlParameter, true, "R");

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
                            var OcStuffIn = new Win_mtr_ocStuffIN_U_CodeView()
                            {
                                Num = i,
                                Req_ID = dr["REQ_ID"].ToString(),
                                ReqName = dr["ReqName"].ToString(),
                                CompanyID = dr["companyID"].ToString(),
                                kCompany = dr["kCompany"].ToString(),

                                StuffDate = dr["StuffDate"].ToString(),
                                CustomID = dr["CustomID"].ToString(),
                                CustomName = dr["CustomName"].ToString(),

                                CustomChief = dr["CustomChief"].ToString(),
                                CustomAddr1 = dr["CustomAddr1"].ToString(),
                                CustomAddr2 = dr["CustomAddr2"].ToString(),
                                CustomAddr3 = dr["CustomAddr3"].ToString(),
                                CustomPhone = dr["CustomPhone"].ToString(),
                                CustomFaxNo = dr["CustomFaxNo"].ToString(),

                                ArticleID = dr["ArticleID"].ToString(),
                                Article = dr["Article"].ToString(),
                                Spec = dr["Spec"].ToString(),
                                StuffClss = dr["StuffClss"].ToString(),
                                StuffClssName = dr["StuffClssName"].ToString(),

                                //FromLocID = dr["fromLocID"].ToString(),
                                // FromLocName = dr["fromLocName"].ToString(),
                                ToLocID = dr["ToLocID"].ToString(),
                                ToLocName = dr["ToLocName"].ToString(),
                                Custom = dr["Custom"].ToString(),

                                StuffRoll = dr["StuffRoll"].ToString(),
                                StuffQty = stringFormatN0(ConvertDouble(dr["StuffQty"].ToString())), // 입고수량 → 소수점 버림 + 천 단위
                                UnitClss = dr["UnitClss"].ToString(),
                                UnitClssName = dr["UnitClssName"].ToString(),
                                PriceClss = dr["PriceClss"].ToString(),

                                PriceClssName = dr["PriceClssName"].ToString(),
                                UnitPrice = dr["UnitPrice"].ToString(),

                                Vat_Ind_YN = dr["Vat_Ind_YN"].ToString(),
                                //ExchRate = stringFormatN2(dr["ExchRate"]),

                                StuffInID = dr["StuffInID"].ToString(),

                                Remark = dr["Remark"].ToString(),
                                Lotid = dr["Lotid"].ToString(),

                                Inspector = dr["Inspector"].ToString(),
                                Inspector1 = dr["Inspector1"].ToString(),
                                InspectDate = dr["InspectDate"].ToString(),
                                InspectApprovalYN = dr["InspectApprovalYN"].ToString(),
                                Amount = stringFormatN0(dr["Amount"]), // 금액 → 소수점 버림 + 천 단위

                                ArticleGrpID = dr["ArticleGrpID"].ToString(),
                                ScrapQty = stringFormatN0(ConvertDouble(dr["ScrapQty"].ToString())), // 잔량 → 소수점 버림 + 천 단위
                                MilSheetNo = dr["MilSheetNo"].ToString(),
                                //QtyPerBox = dr["QtyPerBox"].ToString(),
                                CustomInspector = dr["CustomInspector"].ToString(),
                                BuyerArticleNo = dr["BuyerArticleNo"].ToString()
                                //ProductGrpID = dr["ProductGrpID"].ToString()
                            };

                            if (OcStuffIn.InspectApprovalYN.Trim().Equals("")) { OcStuffIn.InspectApprovalYN = "N"; }

                            // 입고일자
                            OcStuffIn.StuffDate_CV = DatePickerFormat(OcStuffIn.StuffDate);
                            OcStuffIn.InspectDate_CV = DatePickerFormat(OcStuffIn.InspectDate);

                            SumStuffIN.SumStuffInQty += ConvertDouble(OcStuffIn.StuffQty);

                            dgdMainReq.Items.Add(OcStuffIn);
                            dgdMainCustom.Items.Add(OcStuffIn);
                        }

                        SumStuffIN.SumStuffInCount = i;

                        dgdTotal.Items.Add(SumStuffIN);
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
            //끝나면 트루
            btnSearch.IsEnabled = true;
        }

        #endregion // 조회

        #region 유효성 검사
        // 유효성 검사
        private bool CheckData()
        {
            bool flag = true;

            // 입고일자
            if (chkDateSrh.IsChecked == true)
            {
                if (dtpEDateSrh.SelectedDate == null
                    || dtpSDateSrh.SelectedDate == null)
                {
                    MessageBox.Show("입고일자를 선택해주세요.");
                    flag = false;
                    return flag;
                }
            }

            // 제품그룹
            if (chkArticleGrpSrh.IsChecked == true)
            {
                if (cboArticleGrpSrh.SelectedValue == null
                    || cboArticleGrpSrh.SelectedValue.ToString().Trim().Equals(""))
                {
                    MessageBox.Show("제품그룹을 선택해주세요.");
                    flag = false;
                    return flag;
                }
            }

            //// 품명
            //if (chkArticleSrh.IsChecked == true)
            //{
            //    if (txtArticleSrh.Tag == null || txtArticleSrh.Tag.ToString().Trim().Equals("")
            //            || txtArticleSrh.Text.Trim().Equals(""))
            //    {
            //        MessageBox.Show("품명을 입력해주세요.");
            //        flag = false;
            //        return flag;
            //    }
            //}

            // 거래처
            if (chkBuyCustomSrh.IsChecked == true)
            {
                if (txtBuyCustomSrh.Tag == null || txtBuyCustomSrh.Tag.ToString().Trim().Equals("")
                        || txtBuyCustomSrh.Text.Trim().Equals(""))
                {
                    MessageBox.Show("구매거래처를 입력해주세요.");
                    flag = false;
                    return flag;
                }
            }

            // 입고구분
            if (chkStuffClssSrh.IsChecked == true)
            {
                if (cboStuffClssSrh.SelectedValue == null
                    || cboStuffClssSrh.SelectedValue.ToString().Trim().Equals(""))
                {
                    MessageBox.Show("입고구분을 선택해주세요.");
                    flag = false;
                    return flag;
                }
            }

            // 입고창고
            if (chkToLocSrh.IsChecked == true)
            {
                if (cboToLocSrh.SelectedValue == null
                    || cboToLocSrh.SelectedValue.ToString().Trim().Equals(""))
                {
                    MessageBox.Show("입고창고를 선택해주세요.");
                    flag = false;
                    return flag;
                }
            }

            // 입고검수구분
            if (chksInspectApprovalYN.IsChecked == true)
            {
                if (cbosInspectApprovalYN.SelectedValue == null
                    || cbosInspectApprovalYN.SelectedValue.ToString().Trim().Equals(""))
                {
                    MessageBox.Show("입고검수구분을 선택해주세요.");
                    flag = false;
                    return flag;
                }
            }

            return flag;
        }

        #endregion // 유효성 검사

        #endregion // 주요 메서드

        #region 기타메서드

        // 천마리 콤마, 소수점 버리기
        private string stringFormatN0(object obj)
        {
            return string.Format("{0:N0}", obj);
        }

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

        #endregion // 기타메서드

        //품명 라벨체크
        private void lblArticleNo_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            if (chkArticleNoSrh.IsChecked == true)
            {
                chkArticleNoSrh.IsChecked = false;
            }
            else
            {
                chkArticleNoSrh.IsChecked = true;
            }
        }

        //품명 체크 
        private void chkArticleNo_Checked(object sender, RoutedEventArgs e)
        {
            chkArticleNoSrh.IsChecked = true;

            txtArticleNo.IsEnabled = true;

            btnPfArticleNoSrh.IsEnabled = true;
        }
        //품명 안체크
        private void chkArticleNo_Unchecked(object sender, RoutedEventArgs e)
        {
            chkArticleNoSrh.IsChecked = false;

            txtArticleNo.IsEnabled = false;

            btnPfArticleNoSrh.IsEnabled = false;
        }
        //품명 키다운 
        private void txtArticleNo_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                e.Handled = true;
                MainWindow.pf.ReturnCode(txtArticleNo, 77, "");
            }
        }

        //품명 플러스파인더
        private void btnPfArticleNoSrh_Click(object sender, RoutedEventArgs e)
        {
            MainWindow.pf.ReturnCode(txtArticleNo, 77, "");
        }
    }
}

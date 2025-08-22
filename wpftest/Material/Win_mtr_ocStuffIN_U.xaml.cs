using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Data;
using System.Drawing.Printing;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using WizMes_WellMade.PopUp;
using WizMes_WellMade.PopUP;
//*******************************************************************************
//프로그램명    Win_mtr_ocStuffIN_U.cs
//메뉴ID        Win_mtr_ocStuffIN_U
//설명          데이터 처리 클래스
//작성일        2023.04.11
//개발자        J
//*******************************************************************************
// 변경일자     변경자      요청자      요구사항ID          요청 및 작업내용
//*******************************************************************************
//2023.04.11    HD                      Win_mtr_ocStuffIN_U     검사필요 Y일 때 검사, N 이면 자동검사
//
//*******************************************************************************


namespace WizMes_WellMade
{
    public partial class Win_mtr_ocStuffIN_U : UserControl
    {
        string stDate = string.Empty;
        string stTime = string.Empty;

        Lib lib = new Lib();
        string strBasisID = string.Empty;
        string InspectName = string.Empty;
        string AASS = string.Empty;

        string strInspectYN = string.Empty; //무검사품 입고자동여부를 위해서 추가

        string strFlag = string.Empty;
        int rowNum = 0;


        // 스크랩처리를 위해 모아놓는 곳이외다.
        ObservableCollection<Win_mtr_ocStuffIN_U_CodeView> ovcStuffIN = new ObservableCollection<Win_mtr_ocStuffIN_U_CodeView>();
        ObservableCollection<Win_mtr_ocStuffIN_U_CodeView> ovcStuffIN_Scrap = new ObservableCollection<Win_mtr_ocStuffIN_U_CodeView>();
        ObservableCollection<LabelPrint> ovcLabelPrint = new ObservableCollection<LabelPrint>();

        // 인쇄 활용 객체
        private Microsoft.Office.Interop.Excel.Application excelapp;
        private Microsoft.Office.Interop.Excel.Workbook workbook;
        private Microsoft.Office.Interop.Excel.Worksheet worksheet;
        private Microsoft.Office.Interop.Excel.Range workrange;
        private Microsoft.Office.Interop.Excel.Worksheet copysheet;
        private Microsoft.Office.Interop.Excel.Worksheet pastesheet;
        WizMes_WellMade.PopUp.NoticeMessage msg = new WizMes_WellMade.PopUp.NoticeMessage();
        bool printYN = true;
        bool doPass = true;
        public Win_mtr_ocStuffIN_U()
        {
            InitializeComponent();
        }

        // 폼 로드 됬을때
        private void UserControl_Loaded(object sender, RoutedEventArgs e)
        {
            stDate = DateTime.Now.ToString("yyyyMMdd");
            stTime = DateTime.Now.ToString("HHmm");

            DataStore.Instance.InsertLogByFormS(this.GetType().Name, stDate, stTime, "S");

            // 입고일자 체크하기
            chkDateSrh.IsChecked = true;
            dtpSDateSrh.SelectedDate = DateTime.Today;
            dtpEDateSrh.SelectedDate = DateTime.Today;

            SetComboBox();
        }

        #region 추가, 수정 / 저장 후, 취소 메서드

        // 추가, 수정 시
        private void SaveUpdateMode()
        {
            // Header 부분
            Header.IsEnabled = false;

            // 추가, 수정 메세지
            if (strFlag.Equals("I"))
            {
                lblMsg.Content = "자료 추가 중";
            }
            else
            {
                lblMsg.Content = "자료 수정 중";
            }
            lblMsg.Visibility = Visibility.Visible;

            // Content 왼쪽 데이터 그리드
            dgdMain.IsEnabled = false;
            dgdTotal.IsEnabled = false;

            // Content 버튼모음
            btnSave.Visibility = Visibility.Visible;
            btnCancel.Visibility = Visibility.Visible;
            btnAdd.IsEnabled = false;
            btnUpdate.IsEnabled = false;
            btnDelete.IsEnabled = false;

            // Content 오른쪽 입력란
            gbxInput.IsHitTestVisible = true;
            gbxInput2.IsHitTestVisible = true;

            // 일자수정 체크박스 해제
            chkDateUpdate.IsChecked = false;
        }
        // 저장, 취소 시
        private void CompleteCancelMode()
        {
            // Header 부분
            Header.IsEnabled = true;

            // 추가, 수정 메세지
            lblMsg.Visibility = Visibility.Hidden;

            // Content 왼쪽 데이터 그리드
            dgdMain.IsEnabled = true;
            dgdTotal.IsEnabled = true;

            // Content 버튼모음
            btnSave.Visibility = Visibility.Hidden;
            btnCancel.Visibility = Visibility.Hidden;
            btnAdd.IsEnabled = true;
            btnUpdate.IsEnabled = true;
            btnDelete.IsEnabled = true;

            // Content 오른쪽 입력란
            gbxInput.IsHitTestVisible = false;
            gbxInput2.IsHitTestVisible = false;

            // 일자수정 체크박스 해제
            chkDateUpdate.IsChecked = false;

            // 수정시 거래처, 발주번호, 품명 바꾸지 못하게 한거 원상복귀
            txtCustom.IsHitTestVisible = true;
            txtReqID.IsHitTestVisible = true;
            //txtArticle.IsHitTestVisible = true;
        }

        #endregion

        #region 콤보박스 세팅

        private void SetComboBox()
        {

            List<string[]> strValueInOut = new List<string[]>();
            string[] strValueOneInOut = { "0", "내수" };
            string[] strValueTwoInOut = { "1", "수출" };
            strValueInOut.Add(strValueOneInOut);
            strValueInOut.Add(strValueTwoInOut);
            //ObservableCollection<CodeView> ovcFlag = ComboBoxUtil.Instance.Direct_SetComboBox(strValueInOut);
            //cboOrderFlag.ItemsSource = ovcFlag;
            //cboOrderFlag.DisplayMemberPath = "code_name";
            //cboOrderFlag.SelectedValuePath = "code_id";

            ObservableCollection<CodeView> ovcArticleGrpID = ComboBoxUtil.Instance.GetArticleCode_SetComboBox("", 0);
            cboArticleGrp.ItemsSource = ovcArticleGrpID;
            cboArticleGrp.DisplayMemberPath = "code_name";
            cboArticleGrp.SelectedValuePath = "code_id";

            cboArticleGrpSrh.ItemsSource = ovcArticleGrpID;
            cboArticleGrpSrh.DisplayMemberPath = "code_name";
            cboArticleGrpSrh.SelectedValuePath = "code_id";

            ObservableCollection<CodeView> ovcStuff = ComboBoxUtil.Instance.Gf_DB_CM_GetComCodeDataset(null, "ICD", "Y", "", "MTR");
            cboStuffClss.ItemsSource = ovcStuff;
            cboStuffClss.DisplayMemberPath = "code_name";
            cboStuffClss.SelectedValuePath = "code_id";

            cboStuffClssSrh.ItemsSource = ovcStuff;
            cboStuffClssSrh.DisplayMemberPath = "code_name";
            cboStuffClssSrh.SelectedValuePath = "code_id";

            //ObservableCollection<CodeView> ovcFLOC = ComboBoxUtil.Instance.Gf_DB_CM_GetComCodeDataset(null, "LOC", "N", "", "NONE");
            //cboFromLoc.ItemsSource = ovcFLOC;
            //cboFromLoc.DisplayMemberPath = "code_name";
            //cboFromLoc.SelectedValuePath = "code_id";

            ObservableCollection<CodeView> ovcTLOC = ComboBoxUtil.Instance.Gf_DB_CM_GetComCodeDataset(null, "LOC", "Y", "", "INSIDE");
            //ObservableCollection<CodeView> ovcTLOC = ComboBoxUtil.Instance.Gf_DB_CM_GetComCodeDataset(null, "LOC", "Y", "", "MTR");
            cboToLoc.ItemsSource = ovcTLOC;
            cboToLoc.DisplayMemberPath = "code_name";
            cboToLoc.SelectedValuePath = "code_id";

            ObservableCollection<CodeView> ovcUnit = ComboBoxUtil.Instance.Gf_DB_CM_GetComCodeDataset(null, "MTRUNIT", "Y", "", "");
            cboUnit.ItemsSource = ovcUnit;
            cboUnit.DisplayMemberPath = "code_name";
            cboUnit.SelectedValuePath = "code_id";

            ObservableCollection<CodeView> ovcPriceClss = ComboBoxUtil.Instance.Gf_DB_CM_GetComCodeDataset(null, "CRU", "Y", "", "");
            cboPriceClss.ItemsSource = ovcPriceClss;
            cboPriceClss.DisplayMemberPath = "code_name";
            cboPriceClss.SelectedValuePath = "code_id";

            List<string[]> strValueYN0 = new List<string[]>();
            strValueYN0.Add(new string[] { "Y", "Y" });
            strValueYN0.Add(new string[] { "N", "N" });
            strValueYN0.Add(new string[] { "0", "0" });

            ObservableCollection<CodeView> ovcYN0 = ComboBoxUtil.Instance.Direct_SetComboBox(strValueYN0);
            this.cboVatInd.ItemsSource = ovcYN0;
            this.cboVatInd.DisplayMemberPath = "code_name";
            this.cboVatInd.SelectedValuePath = "code_id";

            List<string[]> strValueYN = new List<string[]>();
            strValueYN.Add(new string[] { "Y", "Y" });
            strValueYN.Add(new string[] { "N", "N" });

            ObservableCollection<CodeView> ovcYN = ComboBoxUtil.Instance.Direct_SetComboBox(strValueYN);
            this.cboInspectApprovalYN.ItemsSource = ovcYN;
            this.cboInspectApprovalYN.DisplayMemberPath = "code_name";
            this.cboInspectApprovalYN.SelectedValuePath = "code_id";

            // 검색 입고검수 승인 cbosInspectApprovalYN
            this.cbosInspectApprovalYN.ItemsSource = ovcYN;
            this.cbosInspectApprovalYN.DisplayMemberPath = "code_name";
            this.cbosInspectApprovalYN.SelectedValuePath = "code_id";

            // 검사필요여부cboInspectYN
            this.cboFreeStuffinYN.ItemsSource = ovcYN;
            this.cboFreeStuffinYN.DisplayMemberPath = "code_name";
            this.cboFreeStuffinYN.SelectedValuePath = "code_id";


            cbosInspectApprovalYN.SelectedIndex = 1; // 검색 입곡검수승인 N을 기본으로 해놓기

            // 관리사업장
            ObservableCollection<CodeView> ovcCompany = ComboBoxUtil.Instance.Get_CompanyID();
            this.cboCompanySite.ItemsSource = ovcCompany;
            this.cboCompanySite.DisplayMemberPath = "code_name";
            this.cboCompanySite.SelectedValuePath = "code_id";

            // 제품분류 : 일반, 샤프트반제품(구매품) 
            //List<string[]> strProductGrp = new List<string[]>();
            //strProductGrp.Add(new string[] { "02", "일반" });
            //strProductGrp.Add(new string[] { "01", "샤프트반제품(구매품)" });

            //ObservableCollection<CodeView> ovcProductGrp = SetComboBox_ProductGrp();
            //ObservableCollection<CodeView> ovcProductGrp = ComboBoxUtil.Instance.Direct_SetComboBox(strProductGrp);
            //this.cboProductGrp.ItemsSource = ovcProductGrp;
            //this.cboProductGrp.DisplayMemberPath = "code_name";
            //this.cboProductGrp.SelectedValuePath = "code_id";
        }

        // 제품분류 콤보박스 세팅
        //private ObservableCollection<CodeView> SetComboBox_ProductGrp()
        //{
        //    ObservableCollection<CodeView> ovcProductGrp = new ObservableCollection<CodeView>();

        //    string sql = "";
        //    sql += "SELECT PartGBNID as Code_ID, Code_Name = dbo.fn_cm_sCodeInfo('PARTGBNID', PartGBNID, 'N')";
        //    sql += " FROM [mt_Article]";
        //    sql += " where PartGBNID is not null";

        //    DataSet ds = DataStore.Instance.QueryToDataSet(sql);
        //    if (ds != null && ds.Tables.Count > 0)
        //    {
        //        DataTable dt = ds.Tables[0];
        //        if (dt.Rows.Count != 0)
        //        {
        //            DataRowCollection drc = dt.Rows;

        //            foreach (DataRow dr in drc)
        //            {
        //                CodeView codeView = new CodeView()
        //                {
        //                    code_id = dr["Code_ID"].ToString().Trim(),
        //                    code_name = dr["CODE_NAME"].ToString().Trim()
        //                };
        //                ovcProductGrp.Add(codeView);
        //            }
        //        }
        //    }

        //    return ovcProductGrp;
        //}

        #endregion // 콤보박스 세팅

        #region Header 부분 - 검색조건


        // 입고 일자
        // 입고 일자 검색 라벨 왼쪽 클릭 이벤트
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
        // 입고 일자 검색 체크박스 이벤트
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

        // 제품그룹
        // 제품그룹 검색 라벨 왼쪽 클릭 이벤트
        private void LblArticleGrpSrh_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
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
        // 제품그룹 검색 체크박스 이벤트
        private void chkArticleGrpSrh_Checked(object sender, RoutedEventArgs e)
        {
            chkArticleGrpSrh.IsChecked = true;

            cboArticleGrpSrh.IsEnabled = true;
        }
        private void chkArticleGrpSrh_Unchecked(object sender, RoutedEventArgs e)
        {
            chkArticleGrpSrh.IsChecked = false;

            cboArticleGrpSrh.IsEnabled = false;
        }

        // 품명
        // 품명 검색 라벨 왼쪽 클릭 이벤트
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
        // 품명 검색 체크박스 이벤트
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
            //    re_Search();
            //}

            if (e.Key == Key.Enter)
            {
                e.Handled = true;
                MainWindow.pf.ReturnCode(txtArticleSrh, 76, "");
            }
        }
        // 품명 검색 플러스파인더 이벤트
        private void btnPfArticleSrh_Click(object sender, RoutedEventArgs e)
        {
            MainWindow.pf.ReturnCode(txtArticleSrh, 76, "");
        }

        // 거래처
        // 거래처 검색 라벨 왼쪽 클릭 이벤트
        private void lblCustomSrh_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            if (chkCustomSrh.IsChecked == true)
            {
                chkCustomSrh.IsChecked = false;
            }
            else
            {
                chkCustomSrh.IsChecked = true;
            }
        }
        // 거래처 검색 체크박스 이벤트
        private void chkCustomSrh_Checked(object sender, RoutedEventArgs e)
        {
            chkCustomSrh.IsChecked = true;

            txtCustomSrh.IsEnabled = true;
            btnPfCustomSrh.IsEnabled = true;
        }
        private void chkCustomSrh_Unchecked(object sender, RoutedEventArgs e)
        {
            chkCustomSrh.IsChecked = false;

            txtCustomSrh.IsEnabled = false;
            btnPfCustomSrh.IsEnabled = false;
        }
        // 거래처 검색 엔터 → 플러스 파인더 이벤트
        private void txtCustomSrh_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                e.Handled = true;
                MainWindow.pf.ReturnCode(txtCustomSrh, (int)Defind_CodeFind.DCF_CUSTOM, "");
            }
        }
        // 거래처 검색 플러스파인더 이벤트
        private void btnPfCustomSrh_Click(object sender, RoutedEventArgs e)
        {
            MainWindow.pf.ReturnCode(txtCustomSrh, (int)Defind_CodeFind.DCF_CUSTOM, "");
        }

        // 입고구분
        // 입고구분 검색 라벨 왼쪽 클릭 이벤트
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


        // 입고검수승인
        // 입고검수승인 검색 라벨 왼쪽 클릭 이벤트
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
        // 입고검수승인 검색 체크박스 이벤트
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

        // 스크랩처리
        // 스크랩처리 체크박스 이벤트
        private void chkScrap_Checked(object sender, RoutedEventArgs e)
        {
            btnScrap.Visibility = Visibility.Visible;
        }

        private void chkScrap_Unchecked(object sender, RoutedEventArgs e)
        {
            btnScrap.Visibility = Visibility.Hidden;
        }
        // 사용불가 LotID 검색 버튼 클릭 이벤트
        private void btnScrapSearch_Click(object sender, RoutedEventArgs e)
        {
            FillGrid_Scrap();

            if (dgdMain.Items.Count > 0)
            {
                dgdMain.SelectedIndex = 0;
            }
            else
            {
                MessageBox.Show("자료가 존재하지 않습니다.");
                return;
            }
        }
        // 잔량스크랩처리 버튼 이벤트
        private void btnScrap_Click(object sender, RoutedEventArgs e)
        {
            ovcStuffIN_Scrap.Clear();

            for (int i = 0; i < ovcStuffIN.Count; i++)
            {
                var Scrap = ovcStuffIN[i] as Win_mtr_ocStuffIN_U_CodeView;
                if (Scrap.Chk)
                {
                    ovcStuffIN_Scrap.Add(Scrap);
                }
            }

            if (ovcStuffIN_Scrap.Count > 0)
            {
                if (Scrap(ovcStuffIN_Scrap))
                {
                    MessageBox.Show("스크랩처리가 완료되었습니다.");

                    rowNum = 0;
                    re_Search();
                }
            }
            else
            {
                MessageBox.Show("스크랩 처리 할 대상을 선택해주세요.");
                return;
            }
        }

        #endregion // Header 부분 - 검색조건

        #region 상단 오른쪽 버튼 모음

        // 검색 버튼 이벤트
        private void btnSearch_Click(object sender, RoutedEventArgs e)
        {
            rowNum = 0;

            using (Loading lw = new Loading(re_Search))
            {
                lw.ShowDialog();
            }

            if (dgdMain.Items.Count == 0)
                this.DataContext = null;
        }
        // 닫기 버튼 이벤트
        private void btnClose_Click(object sender, RoutedEventArgs e)
        {
            DataStore.Instance.InsertLogByFormS(this.GetType().Name, stDate, stTime, "E");
            Lib.Instance.ChildMenuClose(this.ToString());
        }

        //public string GetDefaultPrinter()
        //{
        //    PrinterSettings settings = new PrinterSettings();
        //    foreach (string printer in PrinterSettings.InstalledPrinters)
        //    {
        //        settings.PrinterName = printer;

        //        // 지금은 기본프린터 상태가 아니라서 임시로 막아놓음
        //        //if (settings.IsDefaultPrinter && (printer.Contains("TSC"))) //기본 프린트일때
        //        //{
        //        //    return printer;
        //        //}
        //        if (settings.IsDefaultPrinter || (printer.Contains("TSC"))) //기본 프린트일때
        //        {
        //            return printer;
        //        }
        //    }
        //    return string.Empty;
        //}

        // 거래명세서 인쇄 버튼 이벤트
        private void btnPrint_Click(object sender, RoutedEventArgs e)
        {
            string KCustom = "";

            if (ovcStuffIN.Count > 0)
            {
                // 거래처가 돌일한것만 출력이 가능 → 거래처가 동일한것만 인쇄가 가능합니다.
                for (int i = 0; i < ovcStuffIN.Count; i++)
                {
                    var OcStuffin = ovcStuffIN[i];

                    if (OcStuffin != null)
                    {
                        if (i == 0)
                        {
                            KCustom = OcStuffin.Custom;
                        }
                        else
                        {
                            if (!KCustom.Trim().Equals(OcStuffin.Custom.Trim()))
                            {
                                MessageBox.Show("서로 다른 거래처의 입고건을 동시 발행할 수 없습니다.");
                                return;
                            }
                        }
                    }
                }
            }

            // 인쇄 메서드
            ContextMenu menu = btnPrint.ContextMenu;
            menu.StaysOpen = true;
            menu.IsOpen = true;
        }

        #region 거래명세서 인쇄 메서드

        // 미리보기
        private void menuSeeAhead_Click(object sender, RoutedEventArgs e)
        {
            if (ovcStuffIN.Count < 1)
            {
                MessageBox.Show("해당 자료가 존재하지 않습니다.");
                return;
            }

            msg.Show();
            msg.Topmost = true;
            msg.Refresh();

            PrintWork(true);

            msg.Visibility = Visibility.Hidden;
        }

        // 바로 인쇄
        private void menuRightPrint_Click(object sender, RoutedEventArgs e)
        {
            if (ovcStuffIN.Count < 1)
            {
                MessageBox.Show("해당 자료가 존재하지 않습니다.");
                return;
            }
            DataStore.Instance.InsertLogByForm(this.GetType().Name, "P");
            msg.Show();
            msg.Topmost = true;
            msg.Refresh();

            PrintWork(false);

            msg.Visibility = Visibility.Hidden;
        }
        // 닫기
        private void menuClose_Click(object sender, RoutedEventArgs e)
        {
            ContextMenu menu = btnPrint.ContextMenu;
            menu.StaysOpen = false;
            menu.IsOpen = false;
        }

        // 프린터 엑셀 작업
        // 실제 엑셀작업
        private void PrintWork(bool previewYN)
        {
            try
            {
                // 공급받는자
                // 상호
                string CustomName = "";
                // 사업장 주소
                string CustomAddr = "";
                // 성명
                string CustomChief = "";
                //자회사명
                string kCompany = "";
                //자회사 대표
                string Chief = "";
                // 합계 금액 변수

                var Stuffin = dgdMain.SelectedItem as Win_mtr_ocStuffIN_U_CodeView;

                if (Stuffin != null)
                {

                    CustomName = Stuffin.CustomName;
                    CustomChief = Stuffin.CustomChief;
                    CustomAddr = Stuffin.CustomAddr1 + " " + Stuffin.CustomAddr2 + " " + Stuffin.CustomAddr3;
                    kCompany = Stuffin.kCompany;
                    Chief = Stuffin.Chief;
                }

                excelapp = new Microsoft.Office.Interop.Excel.Application();


                string MyBookPath = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetEntryAssembly().Location) + "\\Report\\자재거래명세서.xls";
                workbook = excelapp.Workbooks.Add(MyBookPath);
                worksheet = workbook.Sheets["Form"];
                pastesheet = workbook.Sheets["Print"];

                // 워크시트 이름 변경
                //worksheet.Name = "에너지-" + (year - 1) + "," + year;

                // 거래일자는 어떻게 해야 하는가
                workrange = worksheet.get_Range("C5");
                workrange.Value2 = DateTime.Today.ToString("yyyy.MM.dd");

                // 상호
                workrange = worksheet.get_Range("G6");
                workrange.Value2 = CustomName;

                // 사업장 주소
                workrange = worksheet.get_Range("G8");
                workrange.Value2 = CustomAddr;

                // 성명
                workrange = worksheet.get_Range("G10");
                workrange.Value2 = CustomChief;

                // 회사명
                workrange = worksheet.get_Range("W8");
                workrange.Value2 = kCompany;

                workrange = worksheet.get_Range("AE8");
                workrange.Value2 = Chief;

                // 페이지 계산 등
                int rowCount = ovcStuffIN.Count;
                int excelStartRow = 15;

                int copyLine = 0;
                int Page = 0;
                int PageAll = (int)Math.Ceiling(rowCount / 10.0);
                int DataCount = 0;


                // 총 금액 계산하기
                //double SumAmount = 0;

                for (int k = 0; k < PageAll; k++)
                {
                    Page++;
                    copyLine = ((Page - 1) * 53);

                    int excelNum = 0;

                    // 기존에 있는 데이터 지우기 "A7", "W41"
                    worksheet.Range["C15", "AG24"].EntireRow.ClearContents();

                    for (int i = DataCount; i < rowCount; i++)
                    {
                        if (i == 11 * Page)
                        {
                            break;
                        }

                        var OcStuffin = ovcStuffIN[i];

                        int excelRow = excelStartRow + excelNum;


                        if (OcStuffin != null)
                        {
                            // 월
                            workrange = worksheet.get_Range("C" + excelRow);
                            workrange.Value2 = getDateMonth(OcStuffin.StuffDate);

                            // 일
                            workrange = worksheet.get_Range("D" + excelRow);
                            workrange.Value2 = getDateDay(OcStuffin.StuffDate);

                            // 품명
                            workrange = worksheet.get_Range("E" + excelRow);
                            workrange.Value2 = OcStuffin.BuyerArticleNo;

                            //수량
                            workrange = worksheet.get_Range("P" + excelRow);
                            workrange.Value2 = OcStuffin.StuffQty.Replace(",", "");

                            //단가
                            workrange = worksheet.get_Range("R" + excelRow);
                            workrange.Value2 = OcStuffin.UnitPrice;

                            //공급가액
                            workrange = worksheet.get_Range("W" + excelRow);
                            workrange.Value2 = OcStuffin.Amount;
                        }

                        //SumAmount += ConvertDouble(OcStuffin.Amount);

                        excelNum++;
                        DataCount = i;
                    }


                    // 2장 이상 넘어가면 페이지 넘버 입력
                    //if (PageAll > 1)
                    //{
                    //    pastesheet.PageSetup.CenterFooter = "&P / &N";
                    //}

                    //Form 시트 내용 Print 시트에 복사 붙여넣기
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
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally
            {
                // Clean up 백그라운드에서 엑셀을 지우자 - 달달

                ReleaseExcelObject(workbook);
                ReleaseExcelObject(worksheet);
                ReleaseExcelObject(pastesheet);
            }
        }

        #endregion // 거래명세서 인쇄 메서드


        #region 라벨 인쇄 메서드


        // Lot 라벨 인쇄 버튼 이벤트
        private void btnLotPrint_Click(object sender, RoutedEventArgs e)
        {

            if (ovcStuffIN.Count < 1)
            {
                MessageBox.Show("선택된 데이터가 없습니다.");
                return;
            }

            if (cboFreeStuffinYN.SelectedValue == null ||
                cboFreeStuffinYN.SelectedValue.Equals("N") ||
                cboFreeStuffinYN.SelectedValue.Equals(""))
            {
                MessageBox.Show("검사필요 여부를 확인해주세요");
                return;
            }

            // 인쇄 메서드
            ContextMenu menu = btnLotPrint.ContextMenu;
            menu.StaysOpen = true;
            menu.IsOpen = true;

        }
        // 라벨 인쇄 메뉴 닫기
        private void menuLabelClose_Click(object sender, RoutedEventArgs e)
        {
            ContextMenu menu = btnLotPrint.ContextMenu;
            menu.StaysOpen = false;
            menu.IsOpen = false;
        }
        // 라벨 인쇄 미리보기 버튼 이벤트
        private void menuLabelSeeAhead_Click(object sender, RoutedEventArgs e)
        {
            if (ovcStuffIN.Count < 1)
            {
                MessageBox.Show("해당 자료가 존재하지 않습니다.");
                return;
            }

            printYN = false;
            using (Loading lw = new Loading("excel", LabelPrintWork))
            {
                lw.ShowDialog();
            }

        }
        // 라벨 인쇄 버튼 이벤트
        private void menuLabelRightPrint_Click(object sender, RoutedEventArgs e)
        {
            if (ovcStuffIN.Count < 1)
            {
                MessageBox.Show("해당 자료가 존재하지 않습니다.");
                return;
            }
            DataStore.Instance.InsertLogByForm(this.GetType().Name, "P");
            //msg.Show();
            //msg.Topmost = true;
            //msg.Refresh();

            //Lib.Instance.Delay(1000);

            printYN = true;
            using (Loading lw = new Loading("excel", LabelPrintWork))
            {
                lw.ShowDialog();
            }

            //msg.Visibility = Visibility.Hidden;
        }

        private void setLabePrint(LabelPrint labelPrint)
        {
            int QtyPerBox = ConvertInt(labelPrint.QtyPerBox);
            int Qty = ConvertInt(labelPrint.Qty);

            if (QtyPerBox == 0 || Qty <= QtyPerBox) // QtyPerBox 가 0인 경우 대비
            {
                ovcLabelPrint.Add(labelPrint);
            }
            else if (QtyPerBox != 0 && Qty > QtyPerBox) // QtyPerBox 가 0인 경우 대비
            {
                var RealLabel = new LabelPrint()
                {
                    Custom = labelPrint.Custom,
                    Article = labelPrint.Article,
                    Spec = labelPrint.Spec,
                    StuffDate = labelPrint.StuffDate,
                    CustomInspector = labelPrint.CustomInspector,

                    Qty = stringFormatN0(QtyPerBox),
                    QtyPerBox = labelPrint.QtyPerBox,
                    LotID = labelPrint.LotID,
                    UnitClssName = labelPrint.UnitClssName,
                    BuyerArticleNo = labelPrint.BuyerArticleNo,

                    kCompany = labelPrint.kCompany,
                };

                ovcLabelPrint.Add(RealLabel);

                labelPrint.Qty = stringFormatN0(Qty - QtyPerBox);
                setLabePrint(labelPrint);
            }
        }

        // 라벨 프린터 엑셀 작업
        // 실제 엑셀작업
        private void LabelPrintWork()
        {
            try
            {

                //PrinterSettings settings = new PrinterSettings();
                //string Print = "";
                //foreach (string printer in PrinterSettings.InstalledPrinters)
                //{
                //    settings.PrinterName = printer;
                //    if (printer.Contains("TSC")) //기본 프린트일때
                //    {
                //        Print = printer;
                //    }
                //}
                //PrintDialog pd = new PrintDialog();
                //PrintQueue pq = new PrintQueue(new PrintServer(), Print);
                //pq.Refresh();
                //new PrintServer();

                ////pd.PrintQueue.QueueDriver = new PrintDriver(Print);

                //LocalPrintServer lps = new LocalPrintServer();
                //lps.DefaultPrintQueue = new PrintQueue(new PrintServer(), Print);

                // 엑셀 작업 해봅시다
                excelapp = new Microsoft.Office.Interop.Excel.Application();

                string MyBookPath = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetEntryAssembly().Location) + "\\Report\\자재입고등록라벨_최종본_임시.xls";
                workbook = excelapp.Workbooks.Add(MyBookPath);
                worksheet = workbook.Sheets["Form"];
                pastesheet = workbook.Sheets["Print"];

                #region 실제 엑셀 작업

                ovcLabelPrint.Clear();

                // 실제 라벨 객체에 담기
                for (int i = 0; i < ovcStuffIN.Count; i++)
                {
                    var OcStuffin = ovcStuffIN[i];

                    var LabelPrint = new LabelPrint()
                    {
                        Custom = OcStuffin.Custom,
                        Article = OcStuffin.Article,
                        Spec = OcStuffin.Spec,
                        StuffDate = OcStuffin.StuffDate,
                        CustomInspector = OcStuffin.CustomInspector,

                        Qty = OcStuffin.StuffQty,
                        QtyPerBox = OcStuffin.ProdQtyPerBox,
                        LotID = OcStuffin.Lotid,
                        UnitClssName = OcStuffin.UnitClssName,
                        BuyerArticleNo = OcStuffin.BuyerArticleNo
                    };

                    if (OcStuffin.LabelPrintYN.Trim().Equals("N"))
                    {
                        ovcLabelPrint.Add(LabelPrint);
                    }
                    else
                    {
                        setLabePrint(LabelPrint);
                    }
                }

                // 페이지 계산 등
                int rowCount = ovcLabelPrint.Count;

                int copyLine = 0;
                int Page = 0;
                int PageAll = (int)Math.Ceiling(rowCount / 2.0);
                int DataCount = 0;


                // 총 금액 계산하기
                //double SumAmount = 0;

                for (int k = 0; k < PageAll; k++)
                {
                    Page++;
                    copyLine = ((Page - 1) * 9);

                    // 기존에 있는 데이터 지우기 "A7", "W41"
                    // 왼쪽 정보(거래처, 품명, 수량) 초기화
                    worksheet.Range["B1", "D5"].ClearContents();
                    // 오른쪽 정보(입고일, 품번, 검수자) 초기화
                    worksheet.Range["G1", "I5"].ClearContents();
                    // 라벨 초기화
                    worksheet.Range["A6", "I6"].ClearContents();
                    //worksheet.Range["A4", "E5"].ClearContents();

                    // 왼쪽 라벨 입력
                    var LeftLabel = ovcLabelPrint[DataCount];

                    // 거래처
                    workrange = worksheet.get_Range("B1");
                    workrange.Value2 = LeftLabel.Custom;

                    // 품번
                    workrange = worksheet.get_Range("B2");
                    workrange.Value2 = LeftLabel.BuyerArticleNo;

                    // 입고일
                    workrange = worksheet.get_Range("B4");
                    workrange.Value2 = DatePickerFormat(LeftLabel.StuffDate);

                    // 수량
                    workrange = worksheet.get_Range("B5");
                    workrange.Value2 = LeftLabel.Qty;

                    // 단위
                    workrange = worksheet.get_Range("D5");
                    workrange.Value2 = LeftLabel.UnitClssName;

                    // 스펙 
                    workrange = worksheet.get_Range("B3");
                    workrange.Value2 = LeftLabel.Spec;

                    // 바코드
                    workrange = worksheet.get_Range("A6");
                    workrange.Value2 = "*" + LeftLabel.LotID + "*";

                    workrange = worksheet.get_Range("A7");
                    workrange.Value2 = "'" + LeftLabel.LotID;

                    DataCount++;


                    if (DataCount <= ovcLabelPrint.Count - 1)
                    {
                        var rightLabel = ovcLabelPrint[DataCount];

                        // 거래처
                        workrange = worksheet.get_Range("G1");
                        workrange.Value2 = rightLabel.Custom;

                        // 품명
                        workrange = worksheet.get_Range("G2");
                        workrange.Value2 = rightLabel.BuyerArticleNo;

                        // 입고일
                        workrange = worksheet.get_Range("G4");
                        workrange.Value2 = DatePickerFormat(rightLabel.StuffDate);

                        // 수량
                        workrange = worksheet.get_Range("G5");
                        workrange.Value2 = rightLabel.Qty;

                        // 단위
                        workrange = worksheet.get_Range("I5");
                        workrange.Value2 = rightLabel.UnitClssName;

                        // 스펙 
                        workrange = worksheet.get_Range("G3");
                        workrange.Value2 = rightLabel.Spec;

                        // 바코드
                        workrange = worksheet.get_Range("F6");
                        workrange.Value2 = "*" + rightLabel.LotID + "*";

                        workrange = worksheet.get_Range("F7");
                        workrange.Value2 = "'" + rightLabel.LotID;

                        DataCount++;


                    }

                    //Form 시트 내용 Print 시트에 복사 붙여넣기
                    worksheet.Select();
                    worksheet.UsedRange.EntireRow.Copy();
                    pastesheet.Select();
                    workrange = pastesheet.Cells[copyLine + 1, 1];
                    workrange.Select();
                    pastesheet.Paste();
                    //pastesheet.Paste();

                }

                pastesheet.UsedRange.EntireRow.Select();

                #endregion // 실제 엑셀 작업


                excelapp.Visible = true;

                // 바로 인쇄
                if (printYN == true)
                {
                    pastesheet.PrintOutEx();
                }
                else // 미리보기
                {
                    pastesheet.PrintPreview();
                }


            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally
            {
                // Clean up 백그라운드에서 엑셀을 지우자 - 달달

                ReleaseExcelObject(workbook);
                ReleaseExcelObject(worksheet);
                ReleaseExcelObject(pastesheet);
            }
        }


        //엑셀 백그라운드 증발 - 달달
        private static void ReleaseExcelObject(object obj)
        {
            try
            {
                if (obj != null)
                {
                    Marshal.ReleaseComObject(obj);
                    obj = null;
                }
            }
            catch (Exception ex)
            {
                obj = null;
                throw ex;
            }
            finally
            {
                GC.Collect();
            }
        }

        #endregion // 라벨 인쇄 메서드


        // 엑셀 버튼 이벤트
        private void btnExcel_Click(object sender, RoutedEventArgs e)
        {
            DataTable dt = null;
            string Name = string.Empty;

            string[] lst = new string[2];
            lst[0] = "자재입고내역";
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


        #endregion // 상단 오른쪽 버튼 모음

        #region Content 부분

        // 메인 그리드 셀렉션 체인지 이벤트
        private void dgdMain_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            var OcStuffIN = dgdMain.SelectedItem as Win_mtr_ocStuffIN_U_CodeView;

            if (OcStuffIN != null)
            {
                ShowData(OcStuffIN.StuffInID);

                //// 체크
                //if (OcStuffIN.Chk == true)
                //{
                //    OcStuffIN.Chk = false;

                //    if (ovcStuffIN.Contains(OcStuffIN) == true)
                //    {
                //        ovcStuffIN.Remove(OcStuffIN);
                //    }
                //}
                //else
                //{
                //    OcStuffIN.Chk = true;

                //    if (ovcStuffIN.Contains(OcStuffIN) == false)
                //    {
                //        ovcStuffIN.Add(OcStuffIN);
                //    }
                //}
            }
        }

        // 메인 그리드 전체선택 체크박스 이벤트
        private void AllCheck_Checked(object sender, RoutedEventArgs e)
        {
            ovcStuffIN.Clear();

            for (int i = 0; i < dgdMain.Items.Count; i++)
            {
                var OcStuffIN = dgdMain.Items[i] as Win_mtr_ocStuffIN_U_CodeView;
                OcStuffIN.Chk = true;

                ovcStuffIN.Add(OcStuffIN);
            }
        }
        private void AllCheck_Unchecked(object sender, RoutedEventArgs e)
        {
            ovcStuffIN.Clear();

            for (int i = 0; i < dgdMain.Items.Count; i++)
            {
                var OcStuffIN = dgdMain.Items[i] as Win_mtr_ocStuffIN_U_CodeView;
                OcStuffIN.Chk = false;
            }
        }

        // 메인 그리드 체크박스 이벤트
        // ovcStuffIN 에 추가하기!
        private void chkReq_Click(object sender, RoutedEventArgs e)
        {
            CheckBox chkSender = sender as CheckBox;
            var OcStuffIN = chkSender.DataContext as Win_mtr_ocStuffIN_U_CodeView;

            if (OcStuffIN != null)
            {
                if (chkSender.IsChecked == true)
                {
                    OcStuffIN.Chk = true;

                    if (ovcStuffIN.Contains(OcStuffIN) == false)
                    {
                        ovcStuffIN.Add(OcStuffIN);
                    }
                }
                else
                {
                    OcStuffIN.Chk = false;

                    if (ovcStuffIN.Contains(OcStuffIN) == true)
                    {
                        ovcStuffIN.Remove(OcStuffIN);
                    }
                }

            }
        }

        // 오른쪽 거래처 엔터 → 플러스 파인더 이벤트
        private void txtCustomID_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                e.Handled = true;
                MainWindow.pf.ReturnCode(txtCustom, (int)Defind_CodeFind.DCF_CUSTOM, "");

                // 입고처명에 거래처 집어넣기
                if (!txtCustom.Text.Trim().Equals("") && txtCustom.Tag != null)
                {
                    txtCustomName.Text = txtCustom.Text;
                }
            }
        }
        // 오른쪽 거래처 플러스 파인더 버튼 이벤트
        private void btnPfCustomID_Click(object sender, RoutedEventArgs e)
        {
            if (strFlag.Trim().Equals("U")
                && !txtReqID.Text.Trim().Equals(""))
            {
                MessageBox.Show("거래처는 해당 발주번호를 지운 후에 수정이 가능합니다.");
                return;
            }
            else
            {
                MainWindow.pf.ReturnCode(txtCustom, (int)Defind_CodeFind.DCF_CUSTOM, "");

                // 입고처명에 거래처 집어넣기
                if (!txtCustom.Text.Trim().Equals("") && txtCustom.Tag != null)
                {
                    txtCustomName.Text = txtCustom.Text;
                }
            }
        }

        // 오른쪽 발주 번호 엔터 → 플러스 파인더 이벤트 : 거래처에 해당하는 발주번호 리스트 출력
        private void txtReqID_KeyDown(object sender, KeyEventArgs e)
        {
            if (txtCustom.Tag == null || txtCustom.Tag.ToString().Trim().Equals("")
                || txtCustom.Text.Trim().Equals(""))
            {
                MessageBox.Show("거래처를 먼저 선택해주세요.");
                return;
            }

            if (e.Key == Key.Enter)
            {
                e.Handled = true;

                MainWindow.pf.ReturnCode(txtReqID, 74, txtCustom.Tag.ToString());

                if (!txtReqID.Text.Equals("") && txtReqID.Tag != null && !txtReqID.Tag.ToString().Equals(""))
                {
                    string[] req = txtReqID.Tag.ToString().Trim().Split('/');
                    txtReqID.Text = req[0];
                    txtReqID.Tag = req[0];
                    if (req.Length == 2)
                    {
                        getArticleIdByReq(req[0], ConvertInt(req[1]));
                    }
                }
            }
        }

        // 오른쪽 발주 번호 플러스 파인더 버튼 이벤트
        private void btnPfReqID_Click(object sender, RoutedEventArgs e)
        {
            if (txtCustom.Tag == null || txtCustom.Tag.ToString().Trim().Equals("")
            || txtCustom.Text.Trim().Equals(""))
            {
                MessageBox.Show("거래처를 먼저 선택해주세요.");
                return;
            }

            MainWindow.pf.ReturnCode(txtReqID, 74, txtCustom.Tag.ToString());

            if (!txtReqID.Text.Equals("") && txtReqID.Tag != null && !txtReqID.Tag.ToString().Equals(""))
            {
                string[] req = txtReqID.Tag.ToString().Trim().Split('/');
                txtReqID.Text = req[0];
                txtReqID.Tag = req[0];
                if (req.Length == 2)
                {
                    getArticleIdByReq(req[0], ConvertInt(req[1]));
                }

            }
        }

        // 발주번호로 ArticleID를 가져오기
        private void getArticleIdByReq(string req, int seq)
        {
            try
            {
                Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                sqlParameter.Add("REQ_ID", req);
                sqlParameter.Add("Seq", seq);
                DataSet ds = DataStore.Instance.ProcedureToDataSet("xp_DyeAuxReq_sDyeAuxReqSubOne", sqlParameter, false);

                if (ds != null && ds.Tables.Count > 0)
                {
                    DataTable dt = ds.Tables[0];

                    if (dt.Rows.Count > 0)
                    {
                        txtArticle.Text = dt.Rows[0]["Item_Name"].ToString();
                        txtArticle.Tag = dt.Rows[0]["Item_ID"];

                        // 가져온 ArticleID로 세팅
                        getArticleInfo(txtArticle.Tag.ToString());
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

        // 위의 getArticleIdByReq 로 가져온
        // ArticleID 로 Article 정보 가져오기
        private void getArticleInfo(string setArticleID)
        {
            try
            {
                Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                sqlParameter.Add("ArticleID", setArticleID);

                DataSet ds = DataStore.Instance.ProcedureToDataSet("xp_Order_sArticleData", sqlParameter, false);

                if (ds != null && ds.Tables.Count > 0)
                {
                    DataTable dt = ds.Tables[0];

                    if (dt.Rows.Count > 0)
                    {
                        DataRow dr = dt.Rows[0];

                        var getArticleInfo = new ArticleInfo
                        {
                            ArticleGrpID = dr["ArticleGrpID"].ToString(),
                            UnitPrice = dr["UnitPrice"].ToString(),
                            UnitPriceClss = dr["UnitPriceClss"].ToString(),
                            UnitClss = dr["UnitClss"].ToString(),
                            PartGBNID = dr["PartGBNID"].ToString(),
                            FreeStuffinYN = dr["FreeStuffinYN"].ToString(), // 품명에서 무검사입고품여부 정보들고옴
                            InspectYN = dr["InspectYN"].ToString()
                        };

                        cboArticleGrp.SelectedValue = getArticleInfo.ArticleGrpID;
                        cboPriceClss.SelectedValue = getArticleInfo.UnitPriceClss;
                        cboUnit.SelectedValue = getArticleInfo.UnitClss;
                        cboFreeStuffinYN.SelectedValue = getArticleInfo.FreeStuffinYN; // 품명에서 무검사입고품여부 정보들고옴
                        //cboProductGrp.SelectedValue = getArticleInfo.PartGBNID;
                        txtUnitPrice.Text = getArticleInfo.UnitPrice;
                        cboFreeStuffinYN.SelectedIndex = 0; //검사필요여부 Y면 검사, N이면 자동검사
                        strInspectYN = getArticleInfo.InspectYN; //무검사품 입고자동여부를 위해서 

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

        // 오른쪽 품명 엔터 → 플러스 파인더 이벤트
        private void txtArticle_KeyDown(object sender, KeyEventArgs e)
        {
            //if (e.Key == Key.Enter)
            //{
            //    e.Handled = true;
            //    MainWindow.pf.ReturnCode(txtArticle, (int)Defind_CodeFind.DCF_BuyerArticleNo, "");

            //    if (txtArticle.Tag != null)
            //    {
            //        getArticleInfo(txtArticle.Tag.ToString());
            //    }
            //}
            try
            {
                if (e.Key == Key.Enter)
                {

                    //   if (txtCustom.Tag == null || txtCustom.Tag.ToString().Trim().Equals("")
                    //|| txtCustom.Text.Trim().Equals(""))
                    //   {
                    //       MessageBox.Show("거래처를 먼저 선택해주세요.");
                    //       return;
                    //   }


                    if (txtCustom != null && txtCustom.Text != "")
                    {   //선택된 납품거래처에 따른 품명만 보여주게
                        //MainWindow.pf.ReturnCode(txtArticle, 57, txtCustom.Tag.ToString().Trim());

                        // 품번을 조회하도록 
                        MainWindow.pf.ReturnCodeGLS(txtArticle, 7070, txtCustom.Tag.ToString().Trim());
                        //txtBuyerarticleNo.Tag = txtArticle.Tag;
                        //txtBuyerarticleNo.Text = txtBuyerarticleNo.Text;
                    }
                    else
                    {   //선택된 납품거래처가 없다면 전체 품명 다 보여주게
                        //MainWindow.pf.ReturnCode(txtArticle, (int)Defind_CodeFind.DCF_Article, "");

                        // 품번을 조회하도록 
                        MainWindow.pf.ReturnCodeGLS(txtArticle, 7071, "");
                        //txtBuyerarticleNo.Tag = txtArticle.Tag;
                    }


                    if (txtArticle.Tag != null)
                    {
                        getArticleInfo(txtArticle.Tag.ToString());
                        //////품명그룹 대입(ex.제품 등)
                        ////cboArticleGrp.SelectedValue = ArticleInfo.ArticleGrpID;

                        //////단위
                        ////cboUnit.SelectedValue = ArticleInfo.UnitClss;


                        //cboArticleGrp.SelectedValue = ArticleInfo.ArticleGrpID;
                        //cboPriceClss.SelectedValue = ArticleInfo.UnitPriceClss;
                        //cboUnit.SelectedValue = ArticleInfo.UnitClss;

                    }
                }

                cboPriceClss.SelectedIndex = 0;

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

        // 오른쪽 품명 플러스 파인더 버튼 이벤트
        private void btnPfArticle_Click(object sender, RoutedEventArgs e)
        {
            //MainWindow.pf.ReturnCode(txtArticle, (int)Defind_CodeFind.DCF_BuyerArticleNo, "");

            //if (txtArticle.Tag != null)
            //{
            //    getArticleInfo(txtArticle.Tag.ToString());
            //}
            try
            {
                if (txtCustom != null && txtCustom.Text != "")
                {   //선택된 납품거래처에 따른 품명만 보여주게
                    MainWindow.pf.ReturnCodeGLS(txtArticle, 7070, txtCustom.Tag.ToString().Trim());
                }
                else
                {   //선택된 납품거래처가 없다면 전체 품명 다 보여주게
                    MainWindow.pf.ReturnCodeGLS(txtArticle, 7071, "");
                }

                if (txtArticle.Tag != null)
                {
                    getArticleInfo(txtArticle.Tag.ToString());

                    //cboArticleGrp.SelectedValue = ArticleInfo.ArticleGrpID;
                    //cboPriceClss.SelectedValue = ArticleInfo.UnitPriceClss;
                    //cboUnit.SelectedValue = ArticleInfo.UnitClss;
                }

                ////플러스 파인더 작동 후 규격으로 커서 이동
                //txtSpec.Focus();
                cboPriceClss.SelectedIndex = 0;

            }
            catch (Exception ex)
            {
                //MessageBox.Show("오류 발생, 오류 내용 : " + ex.ToString());
            }
            finally
            {
                DataStore.Instance.CloseConnection();
            }

        }

        // 발주번호를 지운다면, 거래처를 수정할수 있고, 변경이 가능하도록
        private void txtReqID_TextChanged(object sender, TextChangedEventArgs e)
        {
            if (txtReqID.Text.Trim().Equals(""))
            {
                txtCustom.IsHitTestVisible = true;
                txtReqID.Tag = "";
                txtReqID.Text = "";
            }
        }

        #region 오른쪽 상단 추가, 삭제 등등 버튼 이벤트

        //// 저장 버튼 이벤트
        //private void btnSave_Click(object sender, RoutedEventArgs e)
        //{
        //    if (SaveData(strFlag))
        //    {
        //        CompleteCancelMode();

        //        strFlag = string.Empty;

        //        rowNum = 0;
        //        re_Search(rowNum);
        //    }
        //}



        //저장
        private void btnSave_Click(object sender, RoutedEventArgs e)
        {

            using (Loading lw = new Loading(beSave))
            {

                lw.ShowDialog();
            }


        }

        #region 저장 메서드 묶음 beSave()
        private void beSave()
        {
            if (SaveData())
            {
                CompleteCancelMode();
                strBasisID = string.Empty;
                lblMsg.Visibility = Visibility.Hidden;

                if (strFlag.Equals("I"))
                {
                    InspectName = txtStuffInID.ToString();
                    //InspectName = txtKCustom.ToString();
                    //InspectDate = dtpInspectDate.SelectedDate.ToString().Substring(0, 10);

                    rowNum = 0;
                    re_Search();
                    return;
                }
                else
                {
                    rowNum = dgdMain.SelectedIndex;
                }
            }

            int i = 0;

            //foreach (Win_mtr_ocStuffIN_U_CodeView WMRIC in dgdMain.Items)
            //{

            //    string a = WMRIC.StuffInID.ToString();
            //    string b = AASS;


            //    if (a == b)
            //    {
            //        System.Diagnostics.Debug.WriteLine("데이터 같음");

            //        break;
            //    }
            //    else
            //    {
            //        System.Diagnostics.Debug.WriteLine("다름");
            //    }

            //    i++;
            //}

            //rowNum = i;
            //re_Search();
        }
        #endregion

        // 취소 버튼 이벤트
        private void btnCancel_Click(object sender, RoutedEventArgs e)
        {
            using (Loading lw = new Loading(beCancel))
            {
                lw.ShowDialog();
            }
        }

        private void beCancel()
        {
            strFlag = string.Empty;
            CompleteCancelMode();

            re_Search();
        }

        // 추가 버튼 이벤트
        private void btnAdd_Click(object sender, RoutedEventArgs e)
        {


            if (chkRemainAddSrh.IsChecked == true) { txtStuffInID.Text = null; }
            else { this.DataContext = null; }



            strFlag = "I";
            SaveUpdateMode();

            // 추가 세팅
            // 1. 일자수정 오늘날짜
            dtpDate.SelectedDate = DateTime.Today;
            // 2. 입고구분 - 자재입고
            cboStuffClss.SelectedIndex = 0;
            // 3. 관리사업장 - 자회사
            cboCompanySite.SelectedIndex = 0;
            // 4. 사용구분 - 내수
            //cboOrderFlag.SelectedIndex = 0;
            // 5. 입고단위, 화폐, 부가세 - 첫번째 걸로
            cboUnit.SelectedIndex = 0;
            cboPriceClss.SelectedIndex = 0;
            cboVatInd.SelectedIndex = 0;
            // 6. 검수자 - 로그인한 아이디
            //txtCustomInspector.Text = MainWindow.CurrentUser;
            txtCustomInspector.Text = MainWindow.CurrentName;
            // 7. 전창고, 후창고 - 두번째 선택
            //cboFromLoc.SelectedIndex = 0;
            cboToLoc.SelectedIndex = 0;
            cboFreeStuffinYN.SelectedIndex = 0; //검사필요여부 기본값 Y


            // 검수일자 오늘날짜
            dtpCustomInspectDate.SelectedDate = DateTime.Today;

            rowNum = dgdMain.SelectedIndex;
        }
        // 수정 버튼 이벤트
        private void btnUpdate_Click(object sender, RoutedEventArgs e)
        {
            var OcStuffIN = dgdMain.SelectedItem as Win_mtr_ocStuffIN_U_CodeView;

            if (OcStuffIN != null)
            {
                strFlag = "U";
                SaveUpdateMode();

                //txtCustom.IsHitTestVisible = false;
                //txtReqID.IsHitTestVisible = false;
                //txtArticle.IsHitTestVisible = false;

                // 2020.03.16 수정
                // 발주번호가 입력되지 않았다면, 수정 가능.
                if (!OcStuffIN.Req_ID.Trim().Equals(""))
                {
                    txtCustom.IsHitTestVisible = false;
                }

                txtArticle.Tag = OcStuffIN.ArticleID;

                rowNum = dgdMain.SelectedIndex;
            }
            else
            {
                MessageBox.Show("수정할 자료를 선택해주세요.");
            }
        }
        // 삭제 버튼 이벤트
        private void btnDelete_Click(object sender, RoutedEventArgs e)
        {

            var OcStuffIN = dgdMain.SelectedItem as Win_mtr_ocStuffIN_U_CodeView;

            if (OcStuffIN != null)
            {
                if (chkWk_Result(OcStuffIN.Lotid) == false)
                {
                    return;
                }

                // 생산 이력이 있는 입고건은 삭제가 불가능 합니다.         
                if (MessageBox.Show("선택한 항목을 삭제하시겠습니까?", "삭제 전 확인", MessageBoxButton.YesNo) == MessageBoxResult.Yes)
                    beDelete();
            }
            else
            {
                MessageBox.Show("삭제할 데이터를 선택해주세요.");
            }
        }

        private void beDelete()
        {
            //var OcStuffIN = dgdMain.SelectedItem as Win_mtr_ocStuffIN_U_CodeView;
            doPass = false;
            foreach (Win_mtr_ocStuffIN_U_CodeView OcStuffIN in ovcStuffIN)
            {

                if (OcStuffIN != null)
                {
                    if (DeleteData(OcStuffIN.StuffInID))
                    {
                        rowNum = 0;
                        re_Search();
                    }
                }
            }
            doPass = true;
        }

        private bool chkWk_Result(string LotID)
        {
            try
            {
                bool flag = true;

                string[] result = new string[2];
                //2021-06-21 Wk_result -> OutwareSub로 수정(출고내역이 있을 경우 입고내역 삭제를 막기 위해)
                string sql = "SELECT Count(*) as num FROM wk_Result"
                     + " WHERE LabelID = '" + LotID + "'";
                result = DataStore.Instance.ExecuteQuery(sql, true);

                if (result[0].Equals("success"))
                {
                    int count = ConvertInt(result[1]);


                    if (count > 0)
                    {
                        MessageBox.Show("생산 이력이 있는 입고건은 삭제가 불가능 합니다.");
                        return false;
                    }
                }
                else
                {
                    MessageBox.Show("에러 : " + result[1]);
                    return false;
                }

                return flag;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
                return false;
            }
            finally
            {
                DataStore.Instance.CloseConnection();
            }
        }



        #endregion // 오른쪽 상단 추가, 삭제 등등 버튼 이벤트

        #endregion

        #region 주요 메서드

        private void re_Search()
        {
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

        // 조회 검색 메서드
        // 일단 조회 할때마다, ovcStuffIN 초기화 시키기
        private void FillGrid()
        {
            if (doPass)
            {
                ovcStuffIN.Clear();
            }

            // 입고량, 입고건수 - 합계 구하기
            var SumStuffIN = new Win_mtr_ocStuffIN_Sum();

            if (dgdMain.Items.Count > 0)
            {
                dgdMain.Items.Clear();
                dgdTotal.Items.Clear();
            }

            try
            {
                Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                sqlParameter.Clear();

                sqlParameter.Add("nChkDate", chkDateSrh.IsChecked == true ? 1 : 0);
                sqlParameter.Add("sSDate", dtpSDateSrh.SelectedDate != null ? dtpSDateSrh.SelectedDate.Value.ToString("yyyyMMdd") : "");
                sqlParameter.Add("sEDate", dtpEDateSrh.SelectedDate != null ? dtpEDateSrh.SelectedDate.Value.ToString("yyyyMMdd") : "");
                sqlParameter.Add("nChkCustom", chkCustomSrh.IsChecked == true ? 1 : 0);
                sqlParameter.Add("sCustom", txtCustomSrh.Tag != null ? txtCustomSrh.Tag.ToString() : "");

                sqlParameter.Add("nChkArticleID", chkArticleSrh.IsChecked == true ? 1 : 0);
                sqlParameter.Add("sArticleID", txtArticleSrh.Tag != null && !txtArticleSrh.Text.Trim().Equals("") ? txtArticleSrh.Tag.ToString() : "");
                sqlParameter.Add("nChkStuffClss", chkStuffClssSrh.IsChecked == true ? 1 : 0);
                sqlParameter.Add("sStuffClss", cboStuffClssSrh.SelectedValue != null ? cboStuffClssSrh.SelectedValue.ToString() : "");

                sqlParameter.Add("nChkIncStuffIN", 0);

                sqlParameter.Add("nChkArticleGrp", chkArticleGrpSrh.IsChecked == true ? 1 : 0);
                sqlParameter.Add("ArticleGrpID", cboArticleGrpSrh.SelectedValue != null ? cboArticleGrpSrh.SelectedValue.ToString() : "");
                sqlParameter.Add("chkInspect", chksInspectApprovalYN.IsChecked == true ? 1 : 0);      // 입고 검수
                sqlParameter.Add("sInspect", cbosInspectApprovalYN.SelectedValue != null ? cbosInspectApprovalYN.SelectedValue.ToString() : "");

                // [자재입고검사등록] 검색조건 - 나머진 공통 
                sqlParameter.Add("nChkBuyCustom", 0); // 구매거래처 
                sqlParameter.Add("sBuyCustom", "");

                sqlParameter.Add("OrderrByClss", ""); // 발주별, 거래처별 정렬
                sqlParameter.Add("sInspectBasisID", ""); // 입고명세서번호 → 검사성적관리 테이블의 InspectBasisID 속성 검색

                // [자재입고명세서] 검색조건 - 입고창고
                sqlParameter.Add("sToLocID", "");
                sqlParameter.Add("nBuyArticleNo", 0);//chkArticleSrh.IsChecked == true ? 1 : 0);
                sqlParameter.Add("BuyArticleNo", "");//chkArticleSrh.IsChecked == true && !txtArticleSrh.Text.Trim().Equals("") ? txtArticleSrh.Text : "");

                sqlParameter.Add("nChkLotID", chkMtrLOTIDSrh.IsChecked == true ? 1 : 0);
                sqlParameter.Add("sLotID", chkMtrLOTIDSrh.IsChecked == true && !txtMtrLOTIDSrh.Text.Trim().Equals("") ? txtMtrLOTIDSrh.Text : "");  //@escape함수제거


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
                                FromLocID = dr["fromLocID"].ToString(),
                                FromLocName = dr["fromLocName"].ToString(),

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
                                //ExchRate = dr["ExchRate"].ToString(),
                                StuffInID = dr["StuffInID"].ToString(),
                                Remark = dr["Remark"].ToString(),
                                Lotid = dr["Lotid"].ToString(),
                                mtrCustomLotno = dr["mtrCustomLotno"].ToString(),

                                Inspector = dr["Inspector"].ToString(),
                                Inspector1 = dr["Inspector1"].ToString(),
                                InspectDate = dr["InspectDate"].ToString(),
                                InspectApprovalYN = dr["InspectApprovalYN"].ToString(),
                                Amount = stringFormatN0(dr["Amount"]), // 금액 → 소수점 버림 + 천 단위

                                ArticleGrpID = dr["ArticleGrpID"].ToString(),
                                ScrapQty = stringFormatN0(ConvertDouble(dr["ScrapQty"].ToString())), // 잔량 → 소수점 버림 + 천 단위
                                MilSheetNo = dr["MilSheetNo"].ToString(),
                                ProdQtyPerBox = dr["ProdQtyPerBox"].ToString(),
                                CustomInspector = dr["CustomInspector"].ToString(),
                                BuyerArticleNo = dr["BuyerArticleNo"].ToString(),
                                LabelPrintYN = dr["LabelPrintYN"].ToString(),
                                FreeStuffinYN = dr["FreeStuffinYN"].ToString(),
                            };

                            // 입고일자
                            OcStuffIn.StuffDate_CV = DatePickerFormat(OcStuffIn.StuffDate);

                            SumStuffIN.SumStuffInQty += ConvertDouble(OcStuffIn.StuffQty);

                            dgdMain.Items.Add(OcStuffIn);
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
        }

        // 메인 그리드 행 선택시 오른쪽에 데이터 출력
        private void ShowData(string StuffINID)
        {

            try
            {
                this.DataContext = null;

                Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                sqlParameter.Clear();

                sqlParameter.Add("StuffINID", StuffINID);

                DataSet ds = DataStore.Instance.ProcedureToDataSet("xp_StuffIN_sStuffINONE", sqlParameter, false);

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
                            var OcStuffInSub = new Win_mtr_ocStuffIN_U_CodeViewSub()
                            {
                                StuffInID = dr["StuffINID"].ToString(),
                                StuffDate = dr["StuffDate"].ToString(),
                                StuffClss = dr["StuffClss"].ToString(),
                                CustomID = dr["CustomID"].ToString(),
                                CompanyID = dr["CompanyID"].ToString(),

                                kCompany = dr["kCustom"].ToString(),
                                BuyCustomID = dr["BuyCustomID"].ToString(),
                                kBuyCustom = dr["kBuyCustom"].ToString(),
                                BuyerID = dr["BuyerID"].ToString(),
                                BuyerName = dr["BuyerName"].ToString(),

                                Custom = dr["Custom"].ToString(),
                                ArticleID = dr["ArticleID"].ToString(),
                                Article = dr["Article"].ToString(),

                                UnitClss = dr["UnitClss"].ToString(),
                                UnitName = dr["UnitName"].ToString(),
                                TotRoll = dr["TotRoll"].ToString(),
                                TotQty = stringFormatN0(dr["TotQty"]),

                                UnitPrice = dr["UnitPrice"].ToString(),
                                Priceclss = dr["Priceclss"].ToString(),
                                //ExchRate = dr["ExchRate"].ToString(),
                                VAT_IND_YN = dr["VAT_IND_YN"].ToString(),
                                Remark = dr["Remark"].ToString(),

                                InsStuffInYN = dr["InsStuffInYN"].ToString(),
                                OutSeq = dr["OutSeq"].ToString(),

                                Req_ID = dr["Req_ID"].ToString(),
                                REQName = dr["REQName"].ToString(),

                                FromLocID = dr["FromLocID"].ToString(),
                                TOLocID = dr["TOLocID"].ToString(),
                                ArticleGrpID = dr["ArticleGrpID"].ToString(),
                                CustomInspector = dr["CustomInspector"].ToString(),
                                CustomInspectDate = dr["CustomInspectDate"].ToString(),

                                inspector = dr["inspector"].ToString(),
                                InspectApprovalYN = dr["InspectApprovalYN"].ToString(),
                                InspectDate = dr["InspectDate"].ToString(),
                                Lotid = dr["Lotid"].ToString(),
                                inspector1 = dr["inspector1"].ToString(),

                                //mtrHeatingNo = dr["mtrHeatingNo"].ToString(),
                                ProdAutoStuffinYN = dr["ProdAutoStuffinYN"].ToString(),
                                //mtrProdDate = dr["mtrProdDate"].ToString(),
                                //mtrBonsu = dr["mtrBonsu"].ToString(),
                                mtrWeightPerBonsu = dr["mtrWeightPerBonsu"].ToString(),

                                mtrWeight = dr["mtrWeight"].ToString(),
                                //mtrCoilNo = dr["mtrCoilNo"].ToString(),
                                //mtrBNo = dr["mtrBNo"].ToString(),


                                //OrderFlag = dr["OrderFlag"].ToString(),
                                PartGbnID = dr["PartGBNID"].ToString(),


                                BuyerArticleNo = dr["BuyerArticleNo"].ToString(),
                                mtrCustomLotno = dr["mtrCustomLotno"].ToString(),
                                FreeStuffinYN = dr["FreeStuffinYN"].ToString(),



                            };

                            OcStuffInSub.StuffDate_CV = DatePickerFormat(OcStuffInSub.StuffDate);
                            OcStuffInSub.InspectDate_CV = DatePickerFormat(OcStuffInSub.InspectDate);
                            OcStuffInSub.CustomInspectDate_CV = DatePickerFormat(OcStuffInSub.CustomInspectDate);
                            txtCustom.Tag = OcStuffInSub.CustomID;


                            this.DataContext = OcStuffInSub;

                            // 전창고가 빈값이면, 첫번째 행 선택하기
                            //if (OcStuffInSub.FromLocID.Trim().Equals(""))
                            //{
                            //    cboFromLoc.SelectedIndex = 0;
                            //}
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

        // 사용불가 LotID 조회 ← 스크랩처리
        private void FillGrid_Scrap()
        {
            // 입고량, 입고건수 - 합계 구하기
            var SumStuffIN = new Win_mtr_ocStuffIN_Sum();

            if (dgdMain.Items.Count > 0)
            {
                dgdMain.Items.Clear();
                dgdTotal.Items.Clear();
            }

            try
            {
                Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                sqlParameter.Add("sLotID", txtScrapLotID.Text != null ? txtScrapLotID.Text.Trim() : "");
                DataSet ds = DataStore.Instance.ProcedureToDataSet("xp_StuffIN_sStuffIN_Scrap", sqlParameter, true);

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
                                ArticleID = dr["ArticleID"].ToString(),
                                Article = dr["Article"].ToString(),
                                Spec = dr["Spec"].ToString(),

                                StuffClss = dr["StuffClss"].ToString(),
                                StuffClssName = dr["StuffClssName"].ToString(),

                                FromLocID = dr["fromLocID"].ToString(),
                                FromLocName = dr["fromLocName"].ToString(),
                                ToLocID = dr["ToLocID"].ToString(),
                                ToLocName = dr["ToLocName"].ToString(),

                                Custom = dr["Custom"].ToString(),
                                StuffRoll = dr["StuffRoll"].ToString(),
                                StuffQty = stringFormatN0(dr["StuffQty"]), // 입고수량 → 소수점 버림 + 천 단위

                                UnitClss = dr["UnitClss"].ToString(),
                                UnitClssName = dr["UnitClssName"].ToString(),
                                PriceClss = dr["PriceClss"].ToString(),
                                PriceClssName = dr["PriceClssName"].ToString(),
                                UnitPrice = dr["UnitPrice"].ToString(),

                                Vat_Ind_YN = dr["Vat_Ind_YN"].ToString(),
                                //ExchRate = dr["ExchRate"].ToString(),
                                StuffInID = dr["StuffInID"].ToString(),
                                Remark = dr["Remark"].ToString(),
                                Lotid = dr["Lotid"].ToString(),

                                Inspector = dr["Inspector"].ToString(),
                                InspectDate = dr["InspectDate"].ToString(),
                                InspectApprovalYN = dr["InspectApprovalYN"].ToString(),
                                Amount = stringFormatN0(dr["Amount"]), // 금액 → 소수점 버림 + 천 단위

                                ArticleGrpID = dr["ArticleGrpID"].ToString(),
                                ScrapQty = stringFormatN0(dr["ScrapQty"]), // 잔량 → 소수점 버림 + 천 단위
                                MilSheetNo = dr["MilSheetNo"].ToString(),

                                OutUnitPrice = dr["OutUnitPrice"].ToString() // 출고가
                            };

                            // 입고일자
                            OcStuffIn.StuffDate_CV = DatePickerFormat(OcStuffIn.StuffDate);

                            SumStuffIN.SumStuffInQty += ConvertDouble(OcStuffIn.StuffQty);

                            dgdMain.Items.Add(OcStuffIn);
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
        }

        // 삭제 메서드
        private bool DeleteData(string StuffINID)
        {
            bool flag = false;

            List<Procedure> Prolist = new List<Procedure>();
            List<Dictionary<string, object>> ListParameter = new List<Dictionary<string, object>>();

            Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
            sqlParameter.Clear();
            sqlParameter.Add("sStuffINID", StuffINID);

            try
            {

                Procedure pro3 = new Procedure();
                pro3.Name = "xp_StuffIN_dStuffIN";
                pro3.OutputUseYN = "Y";
                pro3.OutputName = "StuffInID";
                pro3.OutputLength = "12";

                Prolist.Add(pro3);
                ListParameter.Add(sqlParameter);

                List<KeyValue> list_Result = new List<KeyValue>();
                list_Result = DataStore.Instance.ExecuteAllProcedureOutputGetCS(Prolist, ListParameter);
                string[] result = DataStore.Instance.ExecuteProcedure_NewLog("xp_StuffIN_dStuffIN", sqlParameter, "D");

                if (list_Result[0].Equals("success"))
                {
                    MessageBox.Show("삭제 실패");
                    flag = false;
                }
                else
                {
                    //MessageBox.Show("성공적으로 삭제되었습니다.");



                    flag = true;
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

            return flag;
        }

        #region 유효성 검사

        // 유효성 검사
        private bool CheckData()
        {
            bool flag = true;


            // 검색할때 유효성 검사
            if (!strFlag.Equals("I") && !strFlag.Equals("U"))
            {
                // 입고일자 검색 조건
                if (chkDateSrh.IsChecked == true
                    && (dtpSDateSrh.SelectedDate == null
                    || dtpEDateSrh.SelectedDate == null))
                {
                    MessageBox.Show("입고일자를 선택해주세요.");
                    flag = false;
                    return flag;
                }

                // 제품그룹 검색 조검
                if (chkArticleGrpSrh.IsChecked == true
                    && (cboArticleGrp.SelectedValue == null
                    || cboArticleGrp.SelectedValue.ToString().Trim().Equals("")))
                {
                    MessageBox.Show("제품그룹을 선택해주세요");
                    flag = false;
                    return flag;
                }

                // 품명 검색 조건
                if (chkArticleSrh.IsChecked == true
                    && (txtArticleSrh.Tag == null
                    || txtArticleSrh.Tag.ToString().Trim().Equals("")))
                {
                    MessageBox.Show("품명이 입력되지 않았습니다.");
                    flag = false;
                    return flag;
                }

                // 거래처 검색 조건
                if (chkCustomSrh.IsChecked == true
                    && (txtCustomSrh.Tag == null
                    || txtCustomSrh.Tag.ToString().Trim().Equals("")))
                {
                    MessageBox.Show("거래처가 입력되지 않았습니다.");
                    flag = false;
                    return flag;
                }

                // 입고구분 검색조건
                if (chkStuffClssSrh.IsChecked == true
                    && (cboStuffClssSrh.SelectedValue == null
                        || cboStuffClssSrh.SelectedValue.ToString().Trim().Equals("")))
                {
                    MessageBox.Show("입고구분을 선택해주세요");
                    flag = false;
                    return flag;
                }

                // 입고검수승인 검색조건
                if (chksInspectApprovalYN.IsChecked == true
                    && (cbosInspectApprovalYN.SelectedValue == null
                       || cbosInspectApprovalYN.SelectedValue.ToString().Trim().Equals("")))
                {
                    MessageBox.Show("입고검수승인을 선택해주세요");
                    flag = false;
                    return flag;
                }
            }

            // 저장, 수정시 유효성 검사
            if (strFlag.Equals("I") || strFlag.Equals("U"))
            {
                // 관리사업장
                if (cboCompanySite.SelectedValue == null
                       || cboCompanySite.SelectedValue.ToString().Trim().Equals(""))
                {
                    MessageBox.Show("관리사업장을 선택해주세요");
                    flag = false;
                    return flag;
                }
                // 거래처
                if (txtCustom.Tag == null
                   || txtCustom.Tag.ToString().Trim().Equals(""))
                {
                    MessageBox.Show("거래처가 입력되지 않았습니다.");
                    flag = false;
                    return flag;
                }
                // 품명
                if (txtArticle.Tag == null
                   || txtArticle.Tag.ToString().Trim().Equals(""))
                {
                    MessageBox.Show("품명이 입력되지 않았습니다.");
                    flag = false;
                    return flag;
                }
                // 품명그룹
                if (cboArticleGrp.SelectedValue == null
                      || cboArticleGrp.SelectedValue.ToString().Trim().Equals(""))
                {
                    MessageBox.Show("품명그룹을 선택해주세요");
                    flag = false;
                    return flag;
                }
                // 제품분류
                //if (cboProductGrp.SelectedValue == null
                //      || cboProductGrp.SelectedValue.ToString().Trim().Equals(""))
                //{
                //    MessageBox.Show("제품분류를 선택해주세요");
                //    flag = false;
                //    return flag;
                //}
                // 입고수량
                if (txInQty.Text == null
                   || txInQty.Text.Trim().Equals(""))
                {
                    MessageBox.Show("입고수량이 입력되지 않았습니다.");
                    flag = false;
                    return flag;
                }
                // 입고단위
                if (cboUnit.SelectedValue == null
                      || cboUnit.SelectedValue.ToString().Trim().Equals(""))
                {
                    MessageBox.Show("입고단위를 선택해주세요");
                    flag = false;
                    return flag;
                }
                if (CheckConvertDouble(txInQty.Text) == false)
                {
                    MessageBox.Show("입고수량은 숫자만 입력 가능 합니다.");
                    flag = false;
                    return flag;
                }
                // 부가세별도 cboVatInd
                if (cboVatInd.SelectedValue == null
                      || cboVatInd.SelectedValue.ToString().Trim().Equals(""))
                {
                    MessageBox.Show("부과세별도를 선택해주세요");
                    flag = false;
                    return flag;
                }
                // 전창고
                //if (cboFromLoc.SelectedValue == null
                //     || cboFromLoc.SelectedValue.ToString().Trim().Equals(""))
                //{
                //    MessageBox.Show("전창고를 선택해주세요");
                //    flag = false;
                //    return flag;
                //}
                // 후창고
                if (cboToLoc.SelectedValue == null
                     || cboToLoc.SelectedValue.ToString().Trim().Equals(""))
                {
                    MessageBox.Show("후창고를 선택해주세요");
                    flag = false;
                    return flag;
                }

                // 필수는 아니지만, 입력했을때 → [환율, 본수, 본딩중량, 중량] 숫자 변환 가능한지 체크
                // 환율
                //if (txtExchRate.Text != null && !txtExchRate.Text.Trim().Equals("")
                //    && CheckConvertDouble(txtExchRate.Text) == false)
                //{
                //    MessageBox.Show("환율은 숫자만 입력 가능 합니다.");
                //    flag = false;
                //    return flag;
                //}
                // 본수
                //if (txtBonsu.Text != null && !txtBonsu.Text.Trim().Equals("")
                //    && CheckConvertDouble(txtBonsu.Text) == false)
                //{
                //    MessageBox.Show("본수는 숫자만 입력 가능 합니다.");
                //    flag = false;
                //    return flag;
                //}
                // 본딩중량
                if (txtWeightPerBonsu.Text != null && !txtWeightPerBonsu.Text.Trim().Equals("")
                    && CheckConvertDouble(txtWeightPerBonsu.Text) == false)
                {
                    MessageBox.Show("본딩중량은 숫자만 입력 가능 합니다.");
                    flag = false;
                    return flag;
                }
                // 중량
                if (txtWeight.Text != null && !txtWeight.Text.Trim().Equals("")
                    && CheckConvertDouble(txtWeight.Text) == false)
                {
                    MessageBox.Show("중량은 숫자만 입력 가능 합니다.");
                    flag = false;
                    return flag;
                }
            }




            return flag;
        }

        #endregion

        // 저장
        private bool SaveData()
        {
            bool flag = false;

            List<Procedure> Prolist = new List<Procedure>();
            List<Dictionary<string, object>> ListParameter = new List<Dictionary<string, object>>();

            try
            {
                if (CheckData())
                {
                    Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                    sqlParameter.Clear();

                    sqlParameter.Add("JobFlag", strFlag);
                    sqlParameter.Add("StuffInID", txtStuffInID.Text != null && !txtStuffInID.Text.Trim().Equals("") ? txtStuffInID.Text : "");
                    sqlParameter.Add("StuffDate", dtpDate.SelectedDate != null ? dtpDate.SelectedDate.Value.ToString("yyyyMMdd") : "");
                    sqlParameter.Add("CompanyID", cboCompanySite.SelectedValue != null ? cboCompanySite.SelectedValue.ToString() : MainWindow.CompanyID);
                    sqlParameter.Add("StuffClss", cboStuffClss.SelectedValue != null ? cboStuffClss.SelectedValue.ToString() : "");

                    sqlParameter.Add("CustomID", txtCustom.Tag != null && !txtCustom.Tag.ToString().Trim().Equals("") ? txtCustom.Tag.ToString() : "");
                    // 이게 뭐임?? 구매자를 적어놓는거 같은데, 거래처를 적으면 되는건가????
                    sqlParameter.Add("BuyCustomID", txtCustom.Tag != null && !txtCustom.Tag.ToString().Trim().Equals("") ? txtCustom.Tag.ToString() : "");
                    sqlParameter.Add("sReqID", txtReqID.Tag != null && !txtReqID.Tag.ToString().Trim().Equals("") ? txtReqID.Tag.ToString() : "");
                    sqlParameter.Add("Custom", txtCustomName.Text != null && !txtCustomName.Text.Trim().Equals("") ? txtCustomName.Text : "");


                    sqlParameter.Add("BuyerID", ""); // 이건 또 뭐여

                    sqlParameter.Add("ArticleID", txtArticle.Tag != null && !txtArticle.Tag.ToString().Trim().Equals("") ? txtArticle.Tag.ToString() : "");
                    sqlParameter.Add("ModelID", ""); // 얜 또 뭐여
                    sqlParameter.Add("UnitClss", cboUnit.SelectedValue != null ? cboUnit.SelectedValue.ToString() : ""); // 입고 단위

                    sqlParameter.Add("TotRoll", 0);
                    sqlParameter.Add("TotQty", ConvertDouble(txInQty.Text)); // 입고수량

                    sqlParameter.Add("UnitPrice", ConvertDouble(txtUnitPrice.Text)); // 단가?? 입력란 없음
                    sqlParameter.Add("PriceClss", cboPriceClss.SelectedValue != null ? cboPriceClss.SelectedValue.ToString() : ""); // 화폐단위
                    sqlParameter.Add("Vat_Ind_YN", cboVatInd.SelectedValue != null ? cboVatInd.SelectedValue.ToString() : ""); // 부과세별도
                    sqlParameter.Add("Remark", txtRemark.Text != null && !txtRemark.Text.Trim().Equals("") ? txtRemark.Text : ""); // 비고 txtRemark

                    sqlParameter.Add("InsStuffINYN", ""); // 동시 입고여부
                    sqlParameter.Add("OutSeq", 0); // 이거는 왜 다 0으로 넣느냐!!!!!!!!!!!!!
                    sqlParameter.Add("sFromLocid", "");
                    sqlParameter.Add("sToLocid", cboToLoc.SelectedValue != null ? cboToLoc.SelectedValue.ToString() : "");

                    sqlParameter.Add("ProdAutoStuffinYN", "");
                    sqlParameter.Add("mtrWeightPerBonsu", txtWeightPerBonsu.Text != null && !txtWeightPerBonsu.Text.Trim().Equals("") ? ConvertDouble(txtWeightPerBonsu.Text) : 0);
                    sqlParameter.Add("mtrWeight", txtWeight.Text != null && !txtWeight.Text.Trim().Equals("") ? ConvertDouble(txtWeight.Text) : 0);

                    sqlParameter.Add("FreeStuffinYN", cboFreeStuffinYN.SelectedValue != null ? cboFreeStuffinYN.SelectedValue.ToString() : ""); // 검사필요 여부 Y/N
                    //sqlParameter.Add("PartGBNID", cboProductGrp.SelectedValue != null ? cboProductGrp.SelectedValue.ToString() : "");
                    sqlParameter.Add("sUserID", MainWindow.CurrentUser);

                    if (strFlag.Equals("I"))
                    {
                        sqlParameter.Add("sLOTID", ""); //2021-08-06 수정을 위해 LOTID 추가
                        Procedure pro1 = new Procedure();
                        pro1.Name = "xp_StuffIN_iuStuffIN";
                        pro1.OutputUseYN = "Y";
                        pro1.OutputName = "StuffInID";
                        pro1.OutputLength = "12";

                        Prolist.Add(pro1);
                        ListParameter.Add(sqlParameter);

                        List<KeyValue> list_Result = new List<KeyValue>();
                        list_Result = DataStore.Instance.ExecuteAllProcedureOutputGetCS_NewLog(Prolist, ListParameter, "C");
                        string sGetID = string.Empty;

                        if (list_Result[0].key.ToLower() == "success")
                        {
                            list_Result.RemoveAt(0);
                            for (int i = 0; i < list_Result.Count; i++)
                            {
                                KeyValue kv = list_Result[i];
                                if (kv.key == "StuffInID")
                                {
                                    sGetID = kv.value;

                                    AASS = kv.value;

                                    flag = true;
                                }
                            }
                        }
                        else
                        {
                            MessageBox.Show("[저장실패]\r\n" + list_Result[0].value.ToString());
                            flag = false;
                            //return false;
                        }

                        // 위에서 실행한 저장 메인그리드 저장 프로시저 삭제
                        Prolist.Clear();
                        ListParameter.Clear();

                        // DB : StuffINSub에 등록
                        sqlParameter = new Dictionary<string, object>();
                        sqlParameter.Clear();

                        sqlParameter.Add("StuffInID", sGetID);

                        sqlParameter.Add("StuffinSubseq", 0); // 얘는 뭔데 다 0인거냐
                        sqlParameter.Add("Qty", ConvertDouble(txInQty.Text));
                        sqlParameter.Add("sCustomInspector", txtCustomInspector.Text != null && !txtCustomInspector.Text.Trim().Equals("") ? txtCustomInspector.Text : "");
                        sqlParameter.Add("sCustomInspectDate", dtpCustomInspectDate.SelectedDate != null ? dtpCustomInspectDate.SelectedDate.Value.ToString("yyyyMMdd") : "");
                        sqlParameter.Add("sLOTID", txtLotID.Text != null && !txtLotID.Text.Trim().Equals("") ? txtLotID.Text : "");
                        sqlParameter.Add("mtrCustomLotno", txtmtrCustomLotno.Text != null && !txtmtrCustomLotno.Text.Trim().Equals("") ? txtmtrCustomLotno.Text : "");

                        sqlParameter.Add("InspectYN", strInspectYN); // 2023.04.11 검사필요 Y = 검사 O, N이면 자동검사
                        //sqlParameter.Add("InspectYN", cboFreeStuffinYN.SelectedValue == null || cboFreeStuffinYN.SelectedValue.Equals("N") ? "Y" : "N"); // 2023.04.11 검사필요 Y = 검사 O, N이면 자동검사

                        //sqlParameter.Add("InspectYN", "Y"); // 2020.02.14 삼주 요청사항 : 입고할때 입고검수도 알아서 되게 해달라!!! 
                        sqlParameter.Add("sUserID", MainWindow.CurrentUser);


                        Procedure pro2 = new Procedure();
                        pro2.Name = "xp_StuffIN_iuStuffINSub";
                        pro2.OutputUseYN = "N";
                        pro2.OutputName = "REQ_ID";
                        pro2.OutputLength = "10";

                        Prolist.Add(pro2);
                        ListParameter.Add(sqlParameter);
                    }
                    else // 수정시
                    {
                        sqlParameter.Add("sLOTID", txtLotID.Text); //2021-08-06 수정을 위해 LOTID 추가
                        Procedure pro1 = new Procedure();
                        pro1.Name = "xp_StuffIN_iuStuffIN";
                        pro1.OutputUseYN = "N";
                        pro1.OutputName = "StuffInID";
                        pro1.OutputLength = "12";

                        Prolist.Add(pro1);
                        ListParameter.Add(sqlParameter);

                        sqlParameter = new Dictionary<string, object>();
                        sqlParameter.Clear();

                        sqlParameter.Add("StuffInID", txtStuffInID.Text);

                        sqlParameter.Add("StuffinSubseq", 0); // 얘는 뭔데 다 0인거냐
                        sqlParameter.Add("Qty", ConvertDouble(txInQty.Text));
                        sqlParameter.Add("sCustomInspector", txtCustomInspector.Text != null && !txtCustomInspector.Text.Trim().Equals("") ? txtCustomInspector.Text : "");
                        sqlParameter.Add("sCustomInspectDate", dtpCustomInspectDate.SelectedDate != null ? dtpCustomInspectDate.SelectedDate.Value.ToString("yyyyMMdd") : "");
                        sqlParameter.Add("sLOTID", txtLotID.Text != null && !txtLotID.Text.Trim().Equals("") ? txtLotID.Text : "");
                        sqlParameter.Add("mtrCustomLotno", txtmtrCustomLotno.Text != null && !txtmtrCustomLotno.Text.Trim().Equals("") ? txtmtrCustomLotno.Text : "");

                        //sqlParameter.Add("sInspector", txtinspector.Text);      //검수자
                        //sqlParameter.Add("sInspectDate", dtpInspectDate.SelectedDate != null ? dtpInspectDate.SelectedDate.Value.ToString("yyyyMMdd") : "");
                        //sqlParameter.Add("sInspectApprovalYN", "N");
                        sqlParameter.Add("sUserID", MainWindow.CurrentUser);


                        Procedure pro2 = new Procedure();
                        pro2.Name = "xp_StuffIN_iuStuffINSub";
                        pro2.OutputUseYN = "N";
                        pro2.OutputName = "REQ_ID";
                        pro2.OutputLength = "10";

                        Prolist.Add(pro2);
                        ListParameter.Add(sqlParameter);
                    }

                    string[] Confirm = new string[2];
                    Confirm = DataStore.Instance.ExecuteAllProcedureOutputNew_NewLog(Prolist, ListParameter, "U");
                    if (Confirm[0] != "success")
                    {
                        MessageBox.Show("[저장실패]\r\n" + Confirm[1].ToString());
                        //MessageBox.Show("[저장실패]\r\n" + "수정할 입고수량보다 출고수량이 많아 수정이 불가합니다.");
                        flag = false;
                        //return false;
                    }
                    else
                    {
                        flag = true;
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
            FillGrid();
            return flag;
        }

        // 스크랩처리
        private bool Scrap(ObservableCollection<Win_mtr_ocStuffIN_U_CodeView> ovcScrap)
        {
            bool flag = true;
            List<Procedure> Prolist = new List<Procedure>();
            List<Dictionary<string, object>> ListParameter = new List<Dictionary<string, object>>();

            try
            {
                Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                for (int i = 0; i < ovcScrap.Count; i++)
                {
                    double SumVatAmount = 0;
                    SumVatAmount += Math.Round(ConvertDouble(ovcScrap[i].ScrapQty) * ConvertDouble(ovcScrap[i].UnitPrice) * 0.1);

                    sqlParameter.Clear();

                    sqlParameter.Add("OrderID", "");
                    sqlParameter.Add("CompanyID", ovcScrap[i].CompanyID);
                    sqlParameter.Add("OutSeq", 0);
                    sqlParameter.Add("OutwareNo", "");
                    sqlParameter.Add("OutClss", "08");

                    sqlParameter.Add("CustomID", ovcScrap[i].CompanyID);
                    sqlParameter.Add("BuyerDirectYN", "N");
                    sqlParameter.Add("WorkID", "0001");
                    sqlParameter.Add("ExchRate", 0);
                    sqlParameter.Add("UnitPriceClss", "0");
                    sqlParameter.Add("InsStuffInYN", "N");

                    sqlParameter.Add("OutcustomID", ovcScrap[i].CompanyID);
                    sqlParameter.Add("Outcustom", ovcScrap[i].kCompany);
                    sqlParameter.Add("LossRate", 0);
                    sqlParameter.Add("LossQty", 0);
                    sqlParameter.Add("OutRoll", 1);

                    sqlParameter.Add("OutQty", ConvertDouble(ovcScrap[i].ScrapQty));
                    sqlParameter.Add("OutRealQty", ConvertDouble(ovcScrap[i].ScrapQty));
                    sqlParameter.Add("OutDate", DateTime.Today.ToString("yyyyMMdd"));
                    sqlParameter.Add("ResultDate", DateTime.Today.ToString("yyyyMMdd"));
                    //sqlParameter.Add("BoOutClss", "");

                    sqlParameter.Add("Remark", "잔량 스크랩 처리");
                    sqlParameter.Add("OutType", "2");
                    sqlParameter.Add("Amount", 0);

                    sqlParameter.Add("VatAmount", SumVatAmount);
                    sqlParameter.Add("VatINDYN", "Y");

                    sqlParameter.Add("FromLocID", "A0001");
                    sqlParameter.Add("ToLocID", "");
                    sqlParameter.Add("UnitClss", ovcScrap[i].UnitClss);
                    sqlParameter.Add("ArticleID", ovcScrap[i].ArticleID);
                    sqlParameter.Add("UserID", MainWindow.CurrentUser);

                    DataStore.Instance.TransactionBegin();
                    Dictionary<string, int> outputParam = new Dictionary<string, int>();
                    outputParam.Add("OutwareNo", 12);
                    outputParam.Add("OutSeq", 10);

                    Dictionary<string, string> dicResult = DataStore.Instance.ExecuteProcedureOutputNoTran("xp_Outware_iOutware", sqlParameter, outputParam, true);
                    string result = dicResult["OutwareNo"];
                    string resultSeq = dicResult["OutSeq"];

                    if ((result != string.Empty) || (result != "9999"))
                    {
                        sqlParameter = new Dictionary<string, object>();
                        sqlParameter.Clear();

                        sqlParameter.Add("OutwareID", result);
                        sqlParameter.Add("OrderID", "");
                        sqlParameter.Add("OutSeq", resultSeq);
                        sqlParameter.Add("OutSubSeq", ConvertInt(resultSeq));
                        sqlParameter.Add("OrderSeq", 0);

                        sqlParameter.Add("LineSeq", 0);
                        sqlParameter.Add("LineSubSeq", 0);
                        sqlParameter.Add("RollSeq", 0);
                        sqlParameter.Add("LabelID", ovcScrap[i].Lotid);
                        sqlParameter.Add("LabelGubun", "");

                        sqlParameter.Add("LotNo", "");
                        sqlParameter.Add("Gubun", "");
                        sqlParameter.Add("StuffQty", 0);
                        sqlParameter.Add("OutQty", ConvertDouble(ovcScrap[i].ScrapQty));
                        sqlParameter.Add("OutRoll", 1);

                        sqlParameter.Add("UnitPrice", ConvertDouble(ovcScrap[i].UnitPrice));
                        sqlParameter.Add("UserID", MainWindow.CurrentUser);
                        sqlParameter.Add("CustomBoxID", "");
                        //sqlParameter.Add("PackingID", "");
                        sqlParameter.Add("StuffinID", ovcScrap[i].StuffInID);

                        sqlParameter.Add("StuffInSubSeq", 1);
                        sqlParameter.Add("ArticleID", ovcScrap[i].ArticleID);

                        string[] SubResult = DataStore.Instance.ExecuteProcedureWithoutTransaction("xp_Outware_iOutwareSub_Excpt", sqlParameter, false);
                        if (SubResult[0].Equals("success"))
                        {
                            sqlParameter = new Dictionary<string, object>();
                            sqlParameter.Clear();

                            sqlParameter.Add("LotID", ovcScrap[i].Lotid);
                            sqlParameter.Add("StuffinID", ovcScrap[i].StuffInID);
                            sqlParameter.Add("UpdateUserID", MainWindow.CurrentUser);

                            string[] FinalResult = DataStore.Instance.ExecuteProcedureWithoutTransaction("xp_Stuffin_uStuffinSub_OutwareYN", sqlParameter, false);

                            if (FinalResult[0].Equals("success"))
                            {
                                DataStore.Instance.TransactionCommit();
                            }
                            else
                            {
                                flag = false;
                                DataStore.Instance.TransactionRollBack();
                                MessageBox.Show("저장실패 , 원인 : " + FinalResult[1]);
                                break;
                            }
                        }
                        else
                        {
                            flag = false;
                            DataStore.Instance.TransactionRollBack();
                            MessageBox.Show("저장실패 , 원인 : " + SubResult[1]);
                            break;
                        }
                    }
                    else
                    {
                        flag = false;
                        DataStore.Instance.TransactionRollBack();
                        MessageBox.Show("저장실패 , 원인 : " + result);
                        break;
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

            return flag;
        }

        #endregion // 주요 메서드

        #region 기타 메서드

        // 천마리 콤마, 소수점 버리기
        private string stringFormatN0(object obj)
        {
            return string.Format("{0:N0}", obj);
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
                str = str.Replace(",", "");

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

        // 월만 가져오기 > 앞에 0 없애기
        private string getDateMonth(string str)
        {
            string month = "";

            if (!str.Trim().Equals(""))
            {
                str = str.Trim().Replace("-", "").Replace(".", "");

                if (str.Length == 8)
                {
                    month = str.Substring(4, 2);

                    if (month.Substring(0, 1).Equals("0"))
                    {
                        month = month.Substring(0, 1);
                    }
                }
            }

            return month;
        }

        // 일만 가져오기 > 앞에 0 없애기
        private string getDateDay(string str)
        {
            string day = "";

            if (!str.Trim().Equals(""))
            {
                str = str.Trim().Replace("-", "").Replace(".", "");

                if (str.Length == 8)
                {
                    day = str.Substring(6, 2);

                    if (day.Substring(0, 1).Equals("0"))
                    {
                        day = day.Substring(1, 1);
                    }
                }
            }

            return day;
        }

        #endregion // 기타 메서드

        private void btnTest_Click(object sender, RoutedEventArgs e)
        {
            string message = "";

            for (int i = 0; i < ovcStuffIN.Count; i++)
            {
                message += "발주번호 : " + ovcStuffIN[i].Req_ID + " / 거래처 : " + ovcStuffIN[i].Custom + " / 품명 : " + ovcStuffIN[i].Article + " \r";
            }

            MessageBox.Show(message);
        }

        private void lblRemainAddSrh_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            if (chkRemainAddSrh.IsChecked == true) { chkRemainAddSrh.IsChecked = false; }
            else { chkRemainAddSrh.IsChecked = true; }
        }

        private void txtMtrLOTIDSrh_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                rowNum = 0;
                re_Search();

            }
        }

        //lot 췌크
        private void chkMtrLOTIDSrh_Checked(object sender, RoutedEventArgs e)
        {
            chkMtrLOTIDSrh.IsChecked = true;
            txtMtrLOTIDSrh.IsEnabled = true;
        }


        //lot 노췌크
        private void chkMtrLOTIDSrh_UnChecked(object sender, RoutedEventArgs e)
        {
            chkMtrLOTIDSrh.IsChecked = false;
            txtMtrLOTIDSrh.IsEnabled = false;
        }

        private void lblSrhMtrLOTID_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            if (chkMtrLOTIDSrh.IsChecked == true)
            {
                chkMtrLOTIDSrh.IsChecked = false;
            }
            else
            {
                chkMtrLOTIDSrh.IsChecked = true;
            }
        }

        //바코드 출력
        private void Button_Click(object sender, RoutedEventArgs e)
        {
            List<TagPrint_CodeView> lstTagPrint = new List<TagPrint_CodeView>();

            // 실제 라벨 객체에 담기
            for (int i = 0; i < ovcStuffIN.Count; i++)
            {
                var OcStuffin = ovcStuffIN[i];

                var Tag = new TagPrint_CodeView()
                {
                    Custom = OcStuffin.Custom,
                    //Article = OcStuffin.Article,
                    BuyerArticle = OcStuffin.Article,
                    Date = OcStuffin.StuffDate_CV,
                    Qty = OcStuffin.StuffQty,
                    UnitClssName = OcStuffin.UnitClssName,
                    Barcode = OcStuffin.Lotid
                };

                lstTagPrint.Add(Tag);
            }


            if (MessageBox.Show("총 " + lstTagPrint.Count + "장의 라벨을 출력하시겠습니까?", "라벨 인쇄 전 확인", MessageBoxButton.YesNo) == MessageBoxResult.Yes)
            {
                string PrintName = GetDefaultPrinter();
                TSCLIB_DLL.openport(PrintName);


                if (SendWindowDllCommand(lstTagPrint, "013", 1, 0))
                {

                    MessageBox.Show("성공");
                }
                else
                {

                    MessageBox.Show("실패");
                }
                TSCLIB_DLL.clearbuffer();
                TSCLIB_DLL.closeport();

            }


        }

        // 기본 프린터 이름 가져오기
        private string GetDefaultPrinter()
        {
            PrinterSettings settings = new PrinterSettings();
            foreach (string printer in PrinterSettings.InstalledPrinters)
            {
                settings.PrinterName = printer;
                if (settings.IsDefaultPrinter && (printer.Contains("TSC"))) //기본 프린트일때
                {
                    return printer;
                }

            }
            return string.Empty;
        }

        // 인쇄
        private bool SendWindowDllCommand(List<TagPrint_CodeView> lstTagPrint, string sTagID, int nPrintCount, int nDefectCnt)
        {

            bool flag = false;

            for (int m = 0; m < lstTagPrint.Count; m++)
            {
                List<string> vData = new List<string>();
                vData.Add(lstTagPrint[m].Barcode);

                vData.Add(lstTagPrint[m].Custom);
                //vData.Add(lstTagPrint[m].Article);
                vData.Add(lstTagPrint[m].BuyerArticle);
                vData.Add(lstTagPrint[m].Qty);
                vData.Add(lstTagPrint[m].UnitClssName);
                vData.Add(lstTagPrint[m].Date);
                //vData.Add(lstTagPrint[m].Barcode);
                var Tag = new mt_Tag_CodeView();
                var TagSub = new mt_TagSub_CodeView();

                List<mt_TagSub_CodeView> lstTagSub = new List<mt_TagSub_CodeView>();

                try
                {
                    // mt_Tag
                    Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                    sqlParameter.Clear();

                    sqlParameter.Add("TagID", sTagID);

                    DataSet ds = DataStore.Instance.ProcedureToDataSet("xp_WizWork_sMtTag", sqlParameter, false);

                    if (ds != null && ds.Tables.Count > 0)
                    {
                        DataTable dt = ds.Tables[0];

                        if (dt.Rows.Count > 0)
                        {
                            foreach (DataRow dr in dt.Rows)
                            {
                                Tag.TagID = dr["TagID"].ToString();
                                Tag.Tag = dr["Tag"].ToString();
                                Tag.Width = ConvertInt(dr["Width"].ToString());
                                Tag.Height = ConvertInt(dr["Height"].ToString());

                                Tag.DefHeight = ConvertInt(dr["DefHeight"].ToString());
                                Tag.DefBaseY = ConvertInt(dr["DefBaseY"].ToString());
                                Tag.DefBaseX1 = ConvertInt(dr["DefBaseX1"].ToString());
                                Tag.DefBaseX2 = ConvertInt(dr["DefBaseX2"].ToString());
                                Tag.DefBaseX3 = ConvertInt(dr["DefBaseX3"].ToString());

                                Tag.DefGapY = ConvertInt(dr["DefGapY"].ToString());
                                Tag.DefGapX1 = ConvertInt(dr["DefGapX1"].ToString());
                                Tag.DefGapX2 = ConvertInt(dr["DefGapX2"].ToString());
                                Tag.DefLength = ConvertInt(dr["DefLength"].ToString());
                                Tag.DefHCount = ConvertInt(dr["DefHCount"].ToString());

                                Tag.DefBarClss = ConvertInt(dr["DefBarClss"].ToString());
                                Tag.Direct = dr["Direct"].ToString();
                            }
                        }
                    }

                    // mt_TagSub
                    sqlParameter = new Dictionary<string, object>();
                    sqlParameter.Clear();

                    sqlParameter.Add("TagID", sTagID);

                    DataSet ds2 = DataStore.Instance.ProcedureToDataSet("xp_WizWork_sMtTagSub", sqlParameter, false);

                    if (ds2 != null && ds.Tables.Count > 0)
                    {
                        DataTable dt = ds2.Tables[0];

                        if (dt.Rows.Count > 0)
                        {
                            int j = 0;
                            foreach (DataRow dr in dt.Rows)
                            {
                                lstTagSub.Add(new mt_TagSub_CodeView());

                                lstTagSub[j].Name = dr["Name"].ToString();
                                lstTagSub[j].Type = ConvertInt(dr["Type"].ToString());
                                lstTagSub[j].Align = ConvertInt(dr["Align"].ToString());
                                lstTagSub[j].x = ConvertInt(dr["x"].ToString());
                                lstTagSub[j].y = ConvertInt(dr["y"].ToString());
                                lstTagSub[j].Font = ConvertInt(dr["Font"].ToString());
                                lstTagSub[j].Length = ConvertInt(dr["Length"].ToString());
                                lstTagSub[j].HMulti = ConvertInt(dr["HMulti"].ToString());
                                lstTagSub[j].VMulti = ConvertInt(dr["VMulti"].ToString());
                                lstTagSub[j].Relation = ConvertInt(dr["Relation"].ToString());
                                lstTagSub[j].Rotation = ConvertInt(dr["Rotation"].ToString());
                                lstTagSub[j].Space = ConvertInt(dr["Space"].ToString());

                                lstTagSub[j].PrevItem = ConvertInt(dr["PrevItem"].ToString());
                                lstTagSub[j].BarType = ConvertInt(dr["BarType"].ToString());
                                lstTagSub[j].BarHeight = ConvertInt(dr["BarHeight"].ToString());
                                lstTagSub[j].FigureWidth = ConvertInt(dr["FigureWidth"].ToString());
                                lstTagSub[j].FigureHeight = ConvertInt(dr["FigureHeight"].ToString());
                                lstTagSub[j].Thickness = ConvertInt(dr["Thickness"].ToString());
                                lstTagSub[j].ImageFile = dr["ImageFile"].ToString();
                                lstTagSub[j].Width = ConvertInt(dr["Width"].ToString());
                                lstTagSub[j].Height = ConvertInt(dr["Height"].ToString());
                                lstTagSub[j].Visible = ConvertInt(dr["Visible"].ToString());

                                lstTagSub[j].FontName = dr["FontName"].ToString();
                                lstTagSub[j].FontStyle = dr["FontStyle"].ToString();
                                lstTagSub[j].FontUnderLine = dr["FontUnderLine"].ToString();

                                //20171011 김종영 수정 type 변경
                                //if (list_m_tItem[i].nType == 1 && list_m_tItem[i].sName.Substring(0, 1).ToUpper() == "D")
                                if ((lstTagSub[j].Type < 2 || lstTagSub[j].Type == 1) && lstTagSub[j].Name.Substring(0, 1).ToUpper() == "D")
                                {
                                    if (lstTagSub[j].Relation == 0 && lstTagSub[j].Type == 1)//바코드
                                    {
                                        lstTagSub[j].Text = vData[0];
                                    }

                                    else if (lstTagSub[j].Relation > 0 && lstTagSub[j].Type == 0)
                                    {
                                        if (vData.Count > lstTagSub[j].Relation)
                                        {
                                            lstTagSub[j].Text = vData[lstTagSub[j].Relation];
                                        }
                                        else
                                        {
                                            lstTagSub[j].Text = "";
                                        }
                                    }
                                }
                                else
                                {
                                    lstTagSub[j].Text = dr["Text"].ToString();
                                }

                                //if (lstTagSub[j].nType < 2 && lstTagSub[j].sName.Substring(0, 1).ToUpper() == "D")
                                //{
                                //    if (lstTagSub[j].nRelation == 0 && lstTagSub[i].nType == 1)//바코드
                                //    {
                                //        lstTagSub[i].sText = vData[0];
                                //    }

                                //    else if (lstTagSub[i].nRelation > 0 && lstTagSub[i].nType == 0)
                                //    {
                                //        if (vData.Count > lstTagSub[i].nRelation)
                                //        {
                                //            lstTagSub[i].sText = vData[lstTagSub[i].nRelation];
                                //        }
                                //        else
                                //        {
                                //            lstTagSub[i].sText = "";
                                //        }
                                //    }
                                //}
                                //else
                                //{
                                //    lstTagSub[i].sText = dr["Text"].ToString();
                                //}

                                j++;
                            } // foreach 끝

                            string strWidth = "";
                            string strHeight = "";
                            try
                            {
                                if (Tag.Width.ToString() != "0")
                                {
                                    strWidth = (Tag.Width / 10F).ToString();
                                }
                                if (Tag.Height.ToString() != "0")
                                {
                                    strHeight = (Tag.Height / 10F).ToString();
                                }
                            }
                            catch
                            {
                                strWidth = "0";
                                strHeight = "0";
                            }

                            TSCLIB_DLL.setup(strWidth, strHeight, "8", "15", "0", "3", "0");//기존소스
                                                                                            //TSCLIB_DLL.setup(strWidth, strHeight, "8", "15", "0", "0", "0");//감열지 테스트용
                            TSCLIB_DLL.sendcommand("DIRECTION " + Tag.Direct);

                            TSCLIB_DLL.clearbuffer();
                            string sText = "";
                            string[] sBarType = new string[2];

                            for (int i = 0; i < lstTagSub.Count; i++)
                            {
                                if (lstTagSub[i].Visible > 0)//출력여부
                                {
                                    ////'QR CODE
                                    //if (lstTagSub[i].Type == EnumItem.IO_QRcode)
                                    //{
                                    //    //QRCODE x, y, ECC Level,cell width, mode, rotation,[model, mask,]"content"

                                    //    string qr_command = "QRCODE " + lstTagSub[i].x.ToString() + "," +
                                    //                lstTagSub[i].y.ToString() + "," +
                                    //                "L" + "," +     // ECC Level (L,M,Q,H)
                                    //                lstTagSub[i].FigureWidth.ToString() + "," +
                                    //                "M" + "," +         // MODE (A,M)
                                    //                "0" + "," +
                                    //                "M2" + "," +
                                    //                "S1" + "," +
                                    //                "\"A" + lstTagSub[0].Text + "\"";


                                    //    //string qr_command = "QRCODE 100,80,L,7,M,0,M2,S1," + "\"A" + list_m_tItem[0].sText + "\"";

                                    //    TSCLIB_DLL.sendcommand(qr_command);

                                    //    //TSCLIB_DLL.windowsfont(lstTagSub[i].x, lstTagSub[i].y, fontheight, rotation, fontstyle, fontunderline, szFaceName, content);
                                    //}
                                    ////'바코드
                                    //else if (lstTagSub[i].Type == EnumItem.IO_BARCODE)
                                    //{
                                    //    if (lstTagSub[i].PrevItem == 0)
                                    //    {
                                    //        if (lstTagSub[i].BarType == 0)// 1:1 Code
                                    //        {
                                    //            sBarType[0] = "1";
                                    //            sBarType[1] = "1";
                                    //        }
                                    //        else                            // 2:5 Code
                                    //        {
                                    //            sBarType[0] = "2";
                                    //            sBarType[1] = "5";
                                    //        }
                                    //        TSCLIB_DLL.barcode(lstTagSub[i].x.ToString(), // x
                                    //                           lstTagSub[i].y.ToString(), // x
                                    //                           "39", // type
                                    //                           lstTagSub[i].BarHeight.ToString(), // 높이
                                    //                           "1", // ReadAble
                                    //                           lstTagSub[i].Rotation.ToString(), // Rotation
                                    //                           sBarType[0], // Narrow
                                    //                           sBarType[1], // Wide
                                    //                           lstTagSub[0].Text
                                    //                           );
                                    //    }
                                    //}
                                    //'바코드
                                    if (lstTagSub[i].Type == EnumItem.IO_BARCODE)
                                    {
                                        if (lstTagSub[i].PrevItem == 0)
                                        {
                                            if (lstTagSub[i].BarType == 0)// 1:1 Code
                                            {
                                                sBarType[0] = "1";
                                                sBarType[1] = "1";
                                            }
                                            else                            // 2:5 Code
                                            {
                                                sBarType[0] = "2";
                                                sBarType[1] = "5";
                                            }
                                            TSCLIB_DLL.barcode(lstTagSub[i].x.ToString(),
                                                               lstTagSub[i].y.ToString(),
                                                               "39",
                                                               lstTagSub[i].BarHeight.ToString(),
                                                               "1",
                                                               lstTagSub[i].Rotation.ToString(),
                                                               sBarType[0],
                                                               sBarType[1],
                                                               lstTagSub[i].Text
                                                               );
                                        }
                                    }
                                    //데이터 OR 문자
                                    else if (lstTagSub[i].Type == EnumItem.IO_DATA || lstTagSub[i].Type == EnumItem.IO_TEXT)
                                    {
                                        sText = lstTagSub[i].Text;
                                        int intx = lstTagSub[i].x;
                                        int inty = lstTagSub[i].y;
                                        int fontheight = ConvertInt((lstTagSub[i].Font).ToString());
                                        int rotation = lstTagSub[i].Rotation;
                                        int fontstyle = ConvertInt(lstTagSub[i].FontStyle);
                                        int fontunderline = ConvertInt(lstTagSub[i].FontUnderLine);
                                        string szFaceName = lstTagSub[i].FontName;
                                        string content = sText.Trim();

                                        TSCLIB_DLL.windowsfont(intx, inty, fontheight, rotation, fontstyle, fontunderline, szFaceName, content);
                                    }
                                    //'선(Line)-5이하
                                    else if (lstTagSub[i].Type == EnumItem.IO_LINE)// && (list_m_tItem[i].nFigureHeight <= 5 || list_m_tItem[i].nFigureWidth <= 5))
                                    {
                                        int x1 = 0;
                                        int x2 = 0;
                                        int y1 = 0;
                                        int y2 = 0;
                                        int.TryParse(lstTagSub[i].x.ToString(), out x1);
                                        int.TryParse(lstTagSub[i].y.ToString(), out y1);
                                        int.TryParse(lstTagSub[i].FigureWidth.ToString(), out x2);
                                        int.TryParse(lstTagSub[i].FigureHeight.ToString(), out y2);

                                        string IsDllStr = "BAR " + x1.ToString() + ", " + y1.ToString() + ", " + x2.ToString() + ", " + y2.ToString();

                                        TSCLIB_DLL.sendcommand(IsDllStr);
                                    }
                                    else if (lstTagSub[i].Type == EnumItem.IO_BOX)
                                    {
                                        int x1 = 0;
                                        int x2 = 0;
                                        int y1 = 0;
                                        int y2 = 0;
                                        int nTh = 0;
                                        int.TryParse(lstTagSub[i].x.ToString(), out x1);
                                        int.TryParse(lstTagSub[i].y.ToString(), out y1);
                                        int.TryParse(lstTagSub[i].FigureWidth.ToString(), out x2);
                                        int.TryParse(lstTagSub[i].FigureHeight.ToString(), out y2);
                                        int.TryParse(lstTagSub[i].Thickness.ToString(), out nTh);

                                        string IsDllStr = "BOX " + x1.ToString() + ", " + y1.ToString() + ", " + x2.ToString() + ", " + y2.ToString() + ", " + nTh.ToString();

                                        TSCLIB_DLL.sendcommand(IsDllStr);
                                    }

                                }
                            }
                            //if (m_ProcessID == "0405")
                            //{
                            //    nPrintCount = 2;
                            //}

                            TSCLIB_DLL.printlabel("1", nPrintCount.ToString());

                            //list_m_tItem = new List<TTagSub>();
                            vData = new List<string>();
                            flag = true;
                        }
                    }
                }
                catch (Exception e)
                {
                    MessageBox.Show(e.Message);
                    //WizCommon.Popup.MyMessageBox.ShowBox(string.Format("오류! 관리자에게 문의<SendWindowDllCommand>\r\n{0}", excpt.Message), "[오류]", 0, 1);
                    flag = false;
                }
                finally
                {
                    DataStore.Instance.CloseConnection();
                }
            }



            return flag;
        }

    }

    #region CodeView(코드뷰)

    class Win_mtr_ocStuffIN_U_CodeView : BaseView
    {
        public bool Chk { get; set; }

        public int Num { get; set; }
        public string Req_ID { get; set; }          // 발주번호
        public string ReqName { get; set; }       // 발주명
        public string CompanyID { get; set; }    // 사업장
        public string kCompany { get; set; }

        public string StuffDate { get; set; }        // 입고일자
        public string StuffDate_CV { get; set; }        // 입고일자

        public string CustomID { get; set; }       // 거래처
        public string CustomName { get; set; }

        public string CustomChief { get; set; }
        public string CustomAddr1 { get; set; }
        public string CustomAddr2 { get; set; }
        public string CustomAddr3 { get; set; }
        public string CustomPhone { get; set; }
        public string CustomFaxNo { get; set; }

        public string ArticleID { get; set; }
        public string Article { get; set; }
        public string Spec { get; set; }

        public string StuffClss { get; set; }          // 입고구분
        public string StuffClssName { get; set; }

        public string FromLocID { get; set; }       // 입고 이전 창고
        public string FromLocName { get; set; }
        public string ToLocID { get; set; }           // 입고 이후 창고
        public string ToLocName { get; set; }

        public string Custom { get; set; }            // 입고처명
        public string StuffRoll { get; set; }
        public string StuffQty { get; set; }
        public string UnitClss { get; set; }           // 입고 단위
        public string UnitClssName { get; set; }

        public string PriceClss { get; set; }
        public string PriceClssName { get; set; }
        public string UnitPrice { get; set; }
        public string Vat_Ind_YN { get; set; }
        //public string ExchRate { get; set; }

        public string StuffInID { get; set; }
        public string Remark { get; set; }
        public string Lotid { get; set; }
        public string mtrCustomLotno { get; set; }
        public string Inspector { get; set; }
        public string Inspector1 { get; set; } // 검수자

        public string InspectDate { get; set; } // 입고 검수 일자
        public string InspectDate_CV { get; set; }
        public string InspectApprovalYN { get; set; }
        public string Amount { get; set; }
        public string ArticleGrpID { get; set; }

        public string ScrapQty { get; set; }    // 잔량
        public string MilSheetNo { get; set; } // 자주검사, 검사성적관리 테이블 속성 (Ins_InspectAuto)

        // 라벨 발행시 쓰는 박스당 수량 - mt_article의 prodQtyPerBox
        public string ProdQtyPerBox { get; set; }
        public string CustomInspector { get; set; }

        // 스크랩 프로시저에서 추가
        public string OutUnitPrice { get; set; } // 출하단가

        public string BuyerArticleNo { get; set; } // 출하단가
        public string LabelPrintYN { get; set; } // 출하단가
        public string FreeStuffinYN { get; set; } // 검사필요여부
        public string Chief { get; set; } // 업체사장 
    }

    class Win_mtr_ocStuffIN_Sum : BaseView
    {
        public double SumStuffInQty { get; set; }
        public double SumStuffInCount { get; set; }
    }

    class Win_mtr_ocStuffIN_U_CodeViewSub : BaseView
    {
        public string StuffInID { get; set; }
        public string StuffDate { get; set; }
        public string StuffDate_CV { get; set; }

        public string StuffClss { get; set; }
        //public string StuffSeq { get; set; }
        public string CustomID { get; set; }
        public string CompanyID { get; set; }
        public string kCompany { get; set; }

        public string BuyCustomID { get; set; }
        public string kBuyCustom { get; set; }
        public string BuyerID { get; set; }
        public string BuyerName { get; set; }
        public string Custom { get; set; }

        public string Article { get; set; }

        public string ArticleID { get; set; }
        public string UnitClss { get; set; }
        public string UnitName { get; set; }
        public string TotRoll { get; set; }
        public string TotQty { get; set; }

        public string UnitPrice { get; set; }
        public string Priceclss { get; set; }
        //public string ExchRate { get; set; }
        public string VAT_IND_YN { get; set; }
        public string Remark { get; set; }

        public string InsStuffInYN { get; set; }
        public string OutSeq { get; set; }

        public string BrandClss { get; set; }
        public string OrderForm { get; set; }
        public string Req_ID { get; set; }
        public string REQName { get; set; }

        public string FromLocID { get; set; }
        public string TOLocID { get; set; }
        public string ArticleGrpID { get; set; }
        public string CustomInspector { get; set; }

        public string CustomInspectDate { get; set; }
        public string CustomInspectDate_CV { get; set; }

        public string inspector { get; set; }
        public string InspectApprovalYN { get; set; }

        public string InspectDate { get; set; }
        public string InspectDate_CV { get; set; }

        public string Lotid { get; set; }
        public string mtrCustomLotno { get; set; }
        public string FreeStuffinYN { get; set; }

        public string inspector1 { get; set; }

        //public string mtrHeatingNo { get; set; }    // 히팅넘버
        public string ProdAutoStuffinYN { get; set; }   // 이건 뭐여

        //public string mtrProdDate { get; set; } // 소재생산일?
        //public string mtrProdDate_CV { get; set; }

        //public string mtrBonsu { get; set; }    // 본수
        public string mtrWeightPerBonsu { get; set; }   // 본딩중량
        public string mtrWeight { get; set; }   // 중량
        //public string mtrCoilNo { get; set; }   // 코일넘버
        //public string mtrBNo { get; set; }      // B/No

        // 제품분류
        public string PartGbnID { get; set; }

        public string BuyerArticleNo { get; set; }
        //public string PartGbnName { get; set; }

        // 사용구분
        //public string OrderFlag { get; set; }

        // 스크랩용
        public string ScrapQty { get; set; }    // 잔량
        public string MilSheetNo { get; set; } // 이건 뭐여
        public string OutUnitPrice { get; set; }
    }

    // 라벨 인쇄에 사용되는 클래스
    class LabelPrint
    {
        public string Custom { get; set; }
        public string Article { get; set; }
        public string Spec { get; set; } // 품번이 뭐여!!!!!?!?!?!?!?
        public string StuffDate { get; set; }
        public string CustomInspector { get; set; }

        public string Qty { get; set; }
        public string LotID { get; set; }
        public string UnitClssName { get; set; }
        public string QtyPerBox { get; set; }
        public string BuyerArticleNo { get; set; }
        public string mtrCustomLotno { get; set; }

        public string kCompany { get; set; }
    }

    class ArticleInfo : BaseView
    {
        public string PartGBNID { get; set; }
        public string ProductGrpID { get; set; }
        public string ArticleGrpID { get; set; }
        public string UnitPrice { get; set; }
        public string UnitPriceClss { get; set; }
        public string UnitClss { get; set; }
        public string ArticleID { get; set; }
        public string Article { get; set; }
        public string ThreadID { get; set; }
        public string thread { get; set; }
        public string StuffWidth { get; set; }
        public string DyeingID { get; set; }
        public string Weight { get; set; }
        public string Spec { get; set; }
        //public string ArticleGrpID { get; set; }
        public string BuyerArticleNo { get; set; }
        //public string UnitPrice { get; set; }
        //public string UnitPriceClss { get; set; }
        //public string UnitClss { get; set; }
        public string ProcessName { get; set; }
        public string HSCode { get; set; }
        public string OutUnitPrice { get; set; }
        public string FreeStuffinYN { get; set; }
        public string InspectYN { get; set; }


    }


    // Tag 코드 뷰
    class TagPrint_CodeView : BaseView
    {
        public string Num { get; set; }
        public string Custom { get; set; }
        public string Article { get; set; }
        public string Date { get; set; }
        public string Spec { get; set; }
        public string UnitClssName { get; set; }
        public string Qty { get; set; }
        public string Barcode { get; set; }
        public string BuyerArticle { get; set; }
    }



    #endregion
}

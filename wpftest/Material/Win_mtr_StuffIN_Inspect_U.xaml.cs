using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Data;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using WizMes_WellMade.PopUp;
using WizMes_WellMade.PopUP;

namespace WizMes_WellMade
{
    /// <summary>
    /// Win_mtr_StuffIN_Inspect_U.xaml에 대한 상호 작용 논리
    /// </summary>
    public partial class Win_mtr_StuffIN_Inspect_U : UserControl
    {
        string stDate = string.Empty;
        string stTime = string.Empty;

        int rowNum = 0;
        ObservableCollection<Win_mtr_ocStuffIN_U_CodeView> ovcStuffIN = new ObservableCollection<Win_mtr_ocStuffIN_U_CodeView>();
        ObservableCollection<Win_mtr_ocStuffIN_U_CodeView> ovcInspect = new ObservableCollection<Win_mtr_ocStuffIN_U_CodeView>();

        public Win_mtr_StuffIN_Inspect_U()
        {
            InitializeComponent();
        }

        // 폼 로드
        private void UserControl_Loaded(object sender, RoutedEventArgs e)
        {
            stDate = DateTime.Now.ToString("yyyyMMdd");
            stTime = DateTime.Now.ToString("HHmm");

            DataStore.Instance.InsertLogByFormS(this.GetType().Name, stDate, stTime, "S");

            // 입고일자 세팅
            chkDateSrh.IsChecked = true;
            dtpSDateSrh.SelectedDate = DateTime.Today;
            dtpEDateSrh.SelectedDate = DateTime.Today;

            // 제품그룹 세팅
            cboArticleGrp.SelectedIndex = 0;

            // 콤보박스 세팅
            SetComboBox();

            // 입고검수 취소 버튼 비활성화
            btnInspectApprovalMinus.IsEnabled = false;
            // 입고검수구분 체크박스 Checked + N으로 기본 세팅
            chksInspectApprovalYN.IsChecked = true;
            cbosInspectApprovalYN.SelectedValue = "N";

            // 발주별 버튼 활성화 및 검색
            tgnReqSort.IsChecked = true;
        }

        #region 추가, 수정 모드 / 완료, 취소 모드

        // 수정,추가 진행 중
        private void SaveUpdateMode()
        {
            //Lib.Instance.UiButtonEnableChange_IUControl(this);
            //dgdMain_REQ.IsHitTestVisible = true;
            //dgdMain_Custom.IsHitTestVisible = true;
            //gbxInput.IsHitTestVisible = false;
        }

        // 수정, 추가 완료
        private void CompleteCancelMode()
        {
            //Lib.Instance.UiButtonEnableChange_SCControl(this);
            //dgdMain_REQ.IsHitTestVisible = false;
            //dgdMain_Custom.IsHitTestVisible = false;
            //gbxInput.IsHitTestVisible = true;
        }

        #endregion // 추가, 수정 모드 / 완료, 취소 모드

        #region 콤보박스 세팅

        // 콤보박스 세팅
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

            this.cbosInspectApprovalYN.ItemsSource = ovcYN;
            this.cbosInspectApprovalYN.DisplayMemberPath = "code_name";
            this.cbosInspectApprovalYN.SelectedValuePath = "code_id";

            // 관리사업장
            ObservableCollection<CodeView> ovcCompany = ComboBoxUtil.Instance.Get_CompanyID();
            this.cboCompanySite.ItemsSource = ovcCompany;
            this.cboCompanySite.DisplayMemberPath = "code_name";
            this.cboCompanySite.SelectedValuePath = "code_id";

            //// 제품분류 : 일반, 샤프트반제품(구매품) 
            //List<string[]> strProductGrp = new List<string[]>();
            //strProductGrp.Add(new string[] { "02", "일반" });
            //strProductGrp.Add(new string[] { "01", "샤프트반제품(구매품)" });

            ////ObservableCollection<CodeView> ovcProductGrp = SetComboBox_ProductGrp();
            //ObservableCollection<CodeView> ovcProductGrp = ComboBoxUtil.Instance.Direct_SetComboBox(strProductGrp);
            //this.cboProductGrp.ItemsSource = ovcProductGrp;
            //this.cboProductGrp.DisplayMemberPath = "code_name";
            //this.cboProductGrp.SelectedValuePath = "code_id";
        }

        #endregion // 콤보박스 세팅

        #region Header 부분 - 검색조건

        // 입고일자 검색
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
                    dtpSDateSrh.SelectedDate = LastMonth31;
                }
                else
                {
                    DateTime ThisMonth1 = DateTime.Today.AddDays(-(DateTime.Today.Day - 1)); // 이번달 1일

                    DateTime LastMonth1 = ThisMonth1.AddMonths(-1); // 저번달 1일
                    DateTime LastMonth31 = ThisMonth1.AddDays(-1); // 저번달 말일

                    dtpSDateSrh.SelectedDate = LastMonth1;
                    dtpSDateSrh.SelectedDate = LastMonth31;
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

        // 제품그룹 검색 라벨 클릭 이벤트
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
        // 품명 검색 라벨 클릭 이벤트
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
        // 품명 검색 텍스트 박스 엔터 이벤트
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
                MainWindow.pf.ReturnCode(txtArticleSrh, 76, "");
            }

        }
        // 품명 검색 플러스파인더 버튼 이벤트
        private void btnPfArticleSrh_Click(object sender, RoutedEventArgs e)
        {
            MainWindow.pf.ReturnCode(txtArticleSrh, (int)Defind_CodeFind.DCF_BuyerArticleNo, "");
        }

        // 거래처 검색 라벨 클릭 이벤트
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
        // 거래처 검색 텍스트박스 엔터 이벤트
        private void txtCustomSrh_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                e.Handled = true;
                MainWindow.pf.ReturnCode(txtCustomSrh, (int)Defind_CodeFind.DCF_CUSTOM, "");
            }
        }
        // 거래처 검색 플러스파인더 버튼 이벤트
        private void btnPfCustomSrh_Click(object sender, RoutedEventArgs e)
        {
            MainWindow.pf.ReturnCode(txtCustomSrh, (int)Defind_CodeFind.DCF_CUSTOM, "");
        }

        // 입고 구분 검색 라벨 클릭 이벤트
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

        // 입고검수구분 검색 라벨 클릭 이벤트
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

        // 입고검수구분 Y 일때는, 입고처리 버튼 활성화 / N 일경우, 입고취소 버튼 활성화
        private void cbosInspectApprovalYN_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (cbosInspectApprovalYN.SelectedValue != null &&
                cbosInspectApprovalYN.SelectedValue.ToString().Equals("N"))
            {
                btnInspectApprovalPlus.IsEnabled = true;
                btnInspectApprovalMinus.IsEnabled = false;
            }
            else if (cbosInspectApprovalYN.SelectedValue != null &&
                cbosInspectApprovalYN.SelectedValue.ToString().Equals("Y"))
            {
                btnInspectApprovalPlus.IsEnabled = false;
                btnInspectApprovalMinus.IsEnabled = true;
            }
        }

        #endregion // Header 부분 - 검색조건

        #region 상단 오른쪽 버튼 모음

        // 검색 버튼
        private void btnSearch_Click(object sender, RoutedEventArgs e)
        {
            rowNum = 0;
            re_Search(rowNum);

            AllCheck.IsChecked = false;
        }
        // 닫기 버튼
        private void btnClose_Click(object sender, RoutedEventArgs e)
        {
            DataStore.Instance.InsertLogByFormS(this.GetType().Name, stDate, stTime, "E");
            Lib.Instance.ChildMenuClose(this.ToString());
        }
        // 입고 승인 처리 버튼
        private void btnInspectApprovalPlus_Click(object sender, RoutedEventArgs e)
        {
            List<string> StuffinID = new List<string>();

            if (ovcStuffIN.Count > 0)
            {
                for (int i = 0; i < ovcStuffIN.Count; i++)
                {
                    var StuffIN = ovcStuffIN[i];

                    if (StuffIN.Chk)
                    {
                        StuffinID.Add(StuffIN.StuffInID);
                    }
                }

                Stuffin_Inspect inspect = new Stuffin_Inspect(StuffinID);
                //inspect.setStuffinID(StuffinID);
                //var location = btnInspectApprovalPlus.PointToScreen(new Point(0, 0));
                //inspect.Left = location.X - 100;
                //inspect.Top = location.Y + 30;
                inspect.WindowStartupLocation = WindowStartupLocation.CenterScreen;
                inspect.ShowDialog();

                if (inspect.DialogResult == true)
                {
                    MessageBox.Show("입고검수처리가 완료되었습니다.");

                    rowNum = 0;
                    re_Search(rowNum);
                }
            }
            else
            {
                MessageBox.Show("처리할 건이 선택되지 않았습니다.");
                return;
            }

        }
        // 입고 취소 버튼
        private void btnInspectApprovalMinus_Click(object sender, RoutedEventArgs e)
        {
            ovcInspect.Clear();

            if (ovcStuffIN != null && ovcStuffIN.Count > 0)
            {

                if (MessageBox.Show("선택한 항목들을 입고검수취소 하시겠습니까?", "입고검수취소 전 확인", MessageBoxButton.YesNo) == MessageBoxResult.Yes)
                {
                    for (int i = 0; i < ovcStuffIN.Count; i++)
                    {
                        var Stuffin = ovcStuffIN[i];

                        if (Stuffin.Chk)
                        {
                            ovcInspect.Add(Stuffin);
                        }
                    }

                    if (ovcInspect != null && ovcInspect.Count > 0)
                    {
                        if (CancelStuffinInspect(ovcInspect))
                        {
                            MessageBox.Show("입고검수취소가 완료되었습니다.");

                            rowNum = 0;
                            re_Search(rowNum);
                        }
                    }
                }
            }
            else
            {
                MessageBox.Show("처리할 건이 선택되지 않았습니다.");
                return;
            }

        }
        // 엑셀 버튼
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

        // 발주별 토글 버튼 이벤트
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

        }
        // 거래처별 토글 버튼 이벤트
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


        }

        // 발주별 메인그리드 선택 이벤트 → Visible 상태일때만 발동 되도록 하기
        private void dgdMainReq_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (dgdMainReq.Visibility == Visibility.Visible)
            {
                var OcStuffIN = dgdMainReq.SelectedItem as Win_mtr_ocStuffIN_U_CodeView;

                if (OcStuffIN != null)
                {
                    ShowData(OcStuffIN.StuffInID);

                    //// 체크박스 체크하고, ovcStuffin 에 넣기
                    //if (OcStuffIN.Chk == false)
                    //{
                    //    OcStuffIN.Chk = true;

                    //    if (ovcStuffIN.Contains(OcStuffIN) == false)
                    //    {
                    //        ovcStuffIN.Add(OcStuffIN);
                    //    }
                    //}
                    //else
                    //{
                    //    OcStuffIN.Chk = false;

                    //    if (ovcStuffIN.Contains(OcStuffIN) == true)
                    //    {
                    //        ovcStuffIN.Remove(OcStuffIN);
                    //    }
                    //}
                }
            }
        }
        // 거래처별 메인그리드 선택 이벤트 → Visible 상태일때만 발동 되도록 하기
        private void dgdMainCustom_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (dgdMainCustom.Visibility == Visibility.Visible)
            {
                var OcStuffIN = dgdMainCustom.SelectedItem as Win_mtr_ocStuffIN_U_CodeView;

                if (OcStuffIN != null)
                {
                    ShowData(OcStuffIN.StuffInID);

                    //// 체크박스 체크하고, ovcStuffin 에 넣기
                    //if (OcStuffIN.Chk == false)
                    //{
                    //    OcStuffIN.Chk = true;

                    //    if (ovcStuffIN.Contains(OcStuffIN) == false)
                    //    {
                    //        ovcStuffIN.Add(OcStuffIN);
                    //    }
                    //}
                    //else
                    //{
                    //    OcStuffIN.Chk = false;

                    //    if (ovcStuffIN.Contains(OcStuffIN) == true)
                    //    {
                    //        ovcStuffIN.Remove(OcStuffIN);
                    //    }
                    //}
                }
            }
        }

        // 발주별 메인 그리드 체크박스 이벤트
        private void CHK_Click_REQ(object sender, RoutedEventArgs e)
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

        // 거래처별 메인 그리드 체크박스 이벤트
        private void CHK_Click_Custom(object sender, RoutedEventArgs e)
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

        // 전체 선택 체크박스 체크 이벤트
        private void AllCheck_Checked(object sender, RoutedEventArgs e)
        {
            ovcStuffIN.Clear();

            if (dgdMainReq.Visibility == Visibility.Visible)
            {
                for (int i = 0; i < dgdMainReq.Items.Count; i++)
                {
                    var OcStuffIN = dgdMainReq.Items[i] as Win_mtr_ocStuffIN_U_CodeView;
                    OcStuffIN.Chk = true;

                    ovcStuffIN.Add(OcStuffIN);
                }
            }
            else if (dgdMainCustom.Visibility == Visibility.Visible)
            {
                for (int i = 0; i < dgdMainCustom.Items.Count; i++)
                {
                    var OcStuffIN = dgdMainCustom.Items[i] as Win_mtr_ocStuffIN_U_CodeView;
                    OcStuffIN.Chk = true;

                    ovcStuffIN.Add(OcStuffIN);
                }
            }
        }

        // 전체 선택 체크박스 언체크 이벤트
        private void AllCheck_Unchecked(object sender, RoutedEventArgs e)
        {
            ovcStuffIN.Clear();

            if (dgdMainReq.Visibility == Visibility.Visible)
            {
                for (int i = 0; i < dgdMainReq.Items.Count; i++)
                {
                    var OcStuffIN = dgdMainReq.Items[i] as Win_mtr_ocStuffIN_U_CodeView;
                    OcStuffIN.Chk = false;
                }
            }
            else if (dgdMainCustom.Visibility == Visibility.Visible)
            {
                for (int i = 0; i < dgdMainCustom.Items.Count; i++)
                {
                    var OcStuffIN = dgdMainCustom.Items[i] as Win_mtr_ocStuffIN_U_CodeView;
                    OcStuffIN.Chk = false;
                }
            }
        }

        #endregion // Content 부분

        #region 주요 메서드 모음 

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

        // 조회 검색 메서드
        // 일단 조회 할때마다, ovcStuffIN 초기화 시키기
        private void FillGrid()
        {
            ovcStuffIN.Clear();

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
                sqlParameter.Add("nChkCustom", chkCustomSrh.IsChecked == true ? 1 : 0);
                sqlParameter.Add("sCustom", txtCustomSrh.Tag != null ? txtCustomSrh.Tag.ToString() : "");

                sqlParameter.Add("nChkArticleID", 0);// chkArticleSrh.IsChecked == true ? 1 : 0);
                sqlParameter.Add("sArticleID", "");// txtArticleSrh.Tag != null ? txtArticleSrh.Tag.ToString() : "");
                sqlParameter.Add("nChkStuffClss", chkStuffClssSrh.IsChecked == true ? 1 : 0);
                sqlParameter.Add("sStuffClss", cboStuffClssSrh.SelectedValue != null ? cboStuffClssSrh.SelectedValue.ToString() : "");

                sqlParameter.Add("nChkIncStuffIN", 0);

                sqlParameter.Add("nChkArticleGrp", chkArticleGrpSrh.IsChecked == true ? 1 : 0);
                sqlParameter.Add("ArticleGrpID", cboArticleGrpSrh.SelectedValue != null ? cboArticleGrpSrh.SelectedValue.ToString() : "");
                sqlParameter.Add("chkInspect", chksInspectApprovalYN.IsChecked == true ? 1 : 0);      // 입고 검수
                sqlParameter.Add("sInspect", cbosInspectApprovalYN.SelectedValue != null ? cbosInspectApprovalYN.SelectedValue.ToString() : "");

                // [자재입고검사등록] 검색조건 - 나머진 공통 
                sqlParameter.Add("nChkBuyCustom", 0); // 구매거래처 ?????????????
                sqlParameter.Add("sBuyCustom", "");

                sqlParameter.Add("OrderrByClss", ""); // 발주별, 거래처별 정렬
                sqlParameter.Add("sInspectBasisID", ""); // 입고명세서번호 → 검사성적관리 테이블의 InspectBasisID 속성 검색????????????????????

                // [자재입고명세서] 검색조건 - 입고창고
                sqlParameter.Add("sToLocID", "");
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

                                Inspector = dr["Inspector"].ToString(),
                                Inspector1 = dr["Inspector1"].ToString(),
                                InspectDate = dr["InspectDate"].ToString(),
                                InspectApprovalYN = dr["InspectApprovalYN"].ToString(),
                                Amount = stringFormatN0(dr["Amount"]), // 금액 → 소수점 버림 + 천 단위
                                BuyerArticleNo = dr["BuyerArticleNo"].ToString(),

                                ArticleGrpID = dr["ArticleGrpID"].ToString(),
                                ScrapQty = stringFormatN0(ConvertDouble(dr["ScrapQty"].ToString())), // 잔량 → 소수점 버림 + 천 단위
                                MilSheetNo = dr["MilSheetNo"].ToString()
                            };

                            // 입고일자
                            OcStuffIn.StuffDate_CV = DatePickerFormat(OcStuffIn.StuffDate);

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
                                BuyerArticleNo = dr["BuyerArticleNo"].ToString(),
                                //OrderFlag = dr["OrderFlag"].ToString(),
                                PartGbnID = dr["PartGBNID"].ToString()


                            };

                            OcStuffInSub.StuffDate_CV = DatePickerFormat(OcStuffInSub.StuffDate);
                            OcStuffInSub.InspectDate_CV = DatePickerFormat(OcStuffInSub.InspectDate);
                            OcStuffInSub.CustomInspectDate_CV = DatePickerFormat(OcStuffInSub.CustomInspectDate);

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

        private bool CancelStuffinInspect(ObservableCollection<Win_mtr_ocStuffIN_U_CodeView> ovcInspect)
        {
            bool flag = false;
            List<Procedure> Prolist = new List<Procedure>();
            List<Dictionary<string, object>> ListParameter = new List<Dictionary<string, object>>();

            try
            {
                Dictionary<string, object> sqlParameter = new Dictionary<string, object>();

                for (int i = 0; i < ovcInspect.Count; i++)
                {
                    sqlParameter = new Dictionary<string, object>();
                    sqlParameter.Clear();
                    sqlParameter.Add("StuffInID", ovcInspect[i].StuffInID);
                    sqlParameter.Add("StuffInSubSeq", "1");
                    sqlParameter.Add("sInspector", "");
                    sqlParameter.Add("sInspectDate", "");
                    sqlParameter.Add("sInspectApprovalYN", "N");
                    sqlParameter.Add("sUserID", MainWindow.CurrentUser);
                    sqlParameter.Add("sInspector1", "");

                    Procedure pro2 = new Procedure();
                    pro2.Name = "xp_StuffIN_uStuffINSub_Inspect";
                    pro2.OutputUseYN = "N";
                    pro2.OutputName = "StuffInID";
                    pro2.OutputLength = "12";

                    Prolist.Add(pro2);
                    ListParameter.Add(sqlParameter);
                }

                string[] Confirm = new string[2];
                Confirm = DataStore.Instance.ExecuteAllProcedureOutputNew(Prolist, ListParameter);
                if (Confirm[0] != "success")
                {
                    MessageBox.Show("[저장실패]\r\n" + Confirm[1].ToString());
                    return flag = false;
                }
                else
                {
                    return flag = true;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }

            return flag;
        }

        #endregion // 주요 메서드 모음

        #region 기타 메서드 모음

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


        #endregion // 기타 메서드 모음

        private void btnTest_Click(object sender, RoutedEventArgs e)
        {
            string message = "";

            for (int i = 0; i < ovcStuffIN.Count; i++)
            {
                message += "발주번호 : " + ovcStuffIN[i].Req_ID + " / 거래처 : " + ovcStuffIN[i].Custom + " / 품명 : " + ovcStuffIN[i].Article + " \r";
            }

            MessageBox.Show(message);
        }

        private void chkDate_Checked(object sender, RoutedEventArgs e)
        {
            chkDateSrh.IsChecked = true;

            dtpSDateSrh.IsEnabled = true;
            dtpEDateSrh.IsEnabled = true;

            btnYesterday.IsEnabled = true;
            btnToday.IsEnabled = true;
            btnLastMonth.IsEnabled = true;
            btnThisMonth.IsEnabled = true;
        }

        private void chkDate_Unchecked(object sender, RoutedEventArgs e)
        {
            chkDateSrh.IsChecked = false;

            dtpSDateSrh.IsEnabled = false;
            dtpEDateSrh.IsEnabled = false;

            btnYesterday.IsEnabled = false;
            btnToday.IsEnabled = false;
            btnLastMonth.IsEnabled = false;
            btnThisMonth.IsEnabled = false;
        }

        private void btnPfBuyerArticleNoSrh_Click(object sender, RoutedEventArgs e)
        {
            MainWindow.pf.ReturnCode(txtArticleSrh, 76, "");
        }
    }

    #region

    class Win_mtr_StuffIN_Inspect_U_CodeView_Old
    {
    }

    #endregion
}

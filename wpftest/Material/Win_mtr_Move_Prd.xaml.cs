using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Data;
using System.Threading;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using WizMes_WellMade.PopUP;

namespace WizMes_WellMade
{
    /// <summary>
    /// Win_mtr_Move_Prd_New.xaml에 대한 상호 작용 논리
    /// </summary>
    public partial class Win_mtr_Move_Prd : UserControl
    {
        string strBasisID = string.Empty;
        string InspectName = string.Empty;
        string AASS = string.Empty;
        Lib lib = new Lib();

        private int rowNum = 0;
        private string strFlag = string.Empty;

        public Win_mtr_Move_Prd()
        {
            InitializeComponent();
        }

        // 폼 로드
        private void UserControl_Loaded(object sender, RoutedEventArgs e)
        {
            // 이동일자 오늘 날짜로 세팅
            chkDateSrh.IsChecked = true;
            dtpSDateSrh.SelectedDate = DateTime.Today;
            dtpEDateSrh.SelectedDate = DateTime.Today;

            SetComboBox();

            // 이동구분 세팅 > 외주이동처리로 고정된 상태 > 변경 불가
            chkOutClssSrh.IsChecked = true;
            cboOutClssSrh.SelectedIndex = 0;

            tgnMoveByID_Click(sender, e); //수량기준이동 띄우기
        }


        #region 추가, 수정 / 저장 후, 취소 메서드

        // 추가, 수정 시
        private void SaveUpdateMode()
        {
            // 상단 저장, 취소 이외에 버튼 비활성화
            lblDateSrh.IsEnabled = false;
            dtpSDateSrh.IsEnabled = false;
            dtpEDateSrh.IsEnabled = false;
            btnYesterDay.IsEnabled = false;
            btnToday.IsEnabled = false;
            btnLastMonth.IsEnabled = false;
            btnThisMonth.IsEnabled = false;

            lblBuyerArticleNoSrh.IsEnabled = false;
            txtBuyerArticleNo.IsEnabled = false;
            btnPfBuyerArticleNo.IsEnabled = false;

            lblFromLocSrh.IsEnabled = false;
            cboFromLocSrh.IsEnabled = false;

            lblToLocSrh.IsEnabled = false;
            cboToLocSrh.IsEnabled = false;

            btnAdd.IsEnabled = false;
            btnUpdate.IsEnabled = false;
            btnDelete.IsEnabled = false;
            btnSearch.IsEnabled = false;
            btnExcel.IsEnabled = false;

            btnSave.Visibility = Visibility.Visible;
            btnCancel.Visibility = Visibility.Visible;

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

            // Content
            // 왼쪽 데이터 그리드
            dgdMain.IsEnabled = false;

            // 토글버튼
            //tgnMoveByID.IsHitTestVisible = true;
            //tgnMoveByQty.IsHitTestVisible = true; 2021-08-06 수량으로 이동 주석 처리로 인한 수정
            tgnMovePartial.IsHitTestVisible = true;

            // 오른쪽 입력란
            gbxInput.IsHitTestVisible = true;

            // 바코드 입력
            txtBarCode.IsHitTestVisible = true;

            // 서브그리드 추가, 삭제 버튼 활성화
            //부분처리 때만 활성화 시키도록 - ID, 수량 기준은 바코드를 통해서만 새로운 행이 추가 가능
            if (tgnMovePartial.IsChecked == true)
            {
                //btnAddSub.IsEnabled = true;
            }
            btnDeleteSub.IsEnabled = true;

        }
        // 저장, 취소 시
        private void CompleteCancelMode()
        {
            // 상단 저장, 취소 이외에 버튼 활성화
            lblDateSrh.IsEnabled = true;
            if (chkDateSrh.IsChecked == true)
            {
                dtpSDateSrh.IsEnabled = true;
                dtpEDateSrh.IsEnabled = true;
            }

            btnYesterDay.IsEnabled = true;
            btnToday.IsEnabled = true;
            btnLastMonth.IsEnabled = true;
            btnThisMonth.IsEnabled = true;

            lblBuyerArticleNoSrh.IsEnabled = true;
            if (chkBuyerArticleNo.IsChecked == true)
            {
                txtBuyerArticleNo.IsEnabled = true;
                btnPfBuyerArticleNo.IsEnabled = true;
            }

            lblFromLocSrh.IsEnabled = true;
            if (chkFromLocSrh.IsChecked == true)
            {
                cboFromLocSrh.IsEnabled = true;
            }

            lblToLocSrh.IsEnabled = true;
            if (chkFromLocSrh.IsChecked == true)
            {
                cboToLocSrh.IsEnabled = true;
            }

            btnAdd.IsEnabled = true;
            btnUpdate.IsEnabled = true;
            btnDelete.IsEnabled = true;
            btnSearch.IsEnabled = true;
            btnExcel.IsEnabled = true;

            btnSave.Visibility = Visibility.Hidden;
            btnCancel.Visibility = Visibility.Hidden;

            // 추가, 수정 메세지
            lblMsg.Visibility = Visibility.Hidden;

            // Content
            // 왼쪽 데이터 그리드
            dgdMain.IsEnabled = true;

            // 토글버튼
            //tgnMoveByID.IsHitTestVisible = false;
            //tgnMoveByQty.IsHitTestVisible = false; 2021-08-06 수량으로 이동 주석 처리로 인한 수정
            tgnMovePartial.IsHitTestVisible = false;

            // 오른쪽 입력란
            gbxInput.IsHitTestVisible = false;

            // 바코드 입력
            txtBarCode.IsHitTestVisible = false;

            // 서브그리드 추가, 삭제 버튼 비활성화
            btnAddSub.IsEnabled = false;
            btnDeleteSub.IsEnabled = false;
        }

        #endregion // 추가, 수정 / 저장 후, 취소 메서드

        #region 콤보박스 세팅

        private void SetComboBox()
        {
            // 검색 전 창고
            // 전 창고
            //ObservableCollection<CodeView> ovcFLOC = ComboBoxUtil.Instance.Gf_DB_CM_GetComCodeDataset(null, "LOC", "Y", "", "");
            //cboFromLocSrh.ItemsSource = ovcFLOC;
            //cboFromLocSrh.DisplayMemberPath = "code_name";
            //cboFromLocSrh.SelectedValuePath = "code_id";

            //cboFromLoc.ItemsSource = ovcFLOC;
            //cboFromLoc.DisplayMemberPath = "code_name";
            //cboFromLoc.SelectedValuePath = "code_id";

            //// 검색 후 창고
            //// 후 창고
            //ObservableCollection<CodeView> ovcTLOC = ComboBoxUtil.Instance.Gf_DB_CM_GetComCodeDataset(null, "LOC", "Y", "", "");
            //cboToLocSrh.ItemsSource = ovcTLOC;
            //cboToLocSrh.DisplayMemberPath = "code_name";
            //cboToLocSrh.SelectedValuePath = "code_id";

            //cboToLoc.ItemsSource = ovcTLOC;
            //cboToLoc.DisplayMemberPath = "code_name";
            //cboToLoc.SelectedValuePath = "code_id";

            //// 검색 이동구분
            //// 이동구분
            //ObservableCollection<CodeView> ovcOut = ComboBoxUtil.Instance.Gf_DB_CM_GetComCodeDataset(null, "OCD", "Y", "", "MOVE");
            //cboOutClssSrh.ItemsSource = ovcOut;
            //cboOutClssSrh.DisplayMemberPath = "code_name";
            //cboOutClssSrh.SelectedValuePath = "code_id";

            //cboOutClss.ItemsSource = ovcOut;
            //cboOutClss.DisplayMemberPath = "code_name";
            //cboOutClss.SelectedValuePath = "code_id";

            //// 단위
            //ObservableCollection<CodeView> ovcUnit = ComboBoxUtil.Instance.Gf_DB_CM_GetComCodeDataset(null, "MTRUNIT", "Y", "", "");
            //cboUnitClss.ItemsSource = ovcUnit;
            //cboUnitClss.DisplayMemberPath = "code_name";
            //cboUnitClss.SelectedValuePath = "code_id";
        }

        #endregion // 콤보박스 세팅

        #region Header 부분

        #region 상단 왼쪽 검색조건 모음

        // 검색 이동일자 라벨 이벤트
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
        // 검색 이동일자 체크박스 이벤트
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

        //전일
        private void btnYesterDay_Click(object sender, RoutedEventArgs e)
        {
            DateTime[] SearchDate = lib.BringLastDayDateTimeContinue(dtpEDateSrh.SelectedDate.Value);

            dtpSDateSrh.SelectedDate = SearchDate[0];
            dtpEDateSrh.SelectedDate = SearchDate[1];
        }

        //금일
        private void btnToday_Click(object sender, RoutedEventArgs e)
        {
            dtpSDateSrh.SelectedDate = DateTime.Today;
            dtpEDateSrh.SelectedDate = DateTime.Today;
        }

        // 전월 버튼 클릭 이벤트
        private void btnLastMonth_Click(object sender, RoutedEventArgs e)
        {
            DateTime[] SearchDate = lib.BringLastMonthContinue(dtpEDateSrh.SelectedDate.Value);

            dtpSDateSrh.SelectedDate = SearchDate[0];
            dtpEDateSrh.SelectedDate = SearchDate[1];
        }

        // 금월 버튼 클릭 이벤트
        private void btnThisMonth_Click(object sender, RoutedEventArgs e)
        {
            dtpSDateSrh.SelectedDate = lib.BringThisMonthDatetimeList()[0];
            dtpEDateSrh.SelectedDate = lib.BringThisMonthDatetimeList()[1];
        }

        // 검색 품명 라벨 이벤트
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
        // 검색 품명 체크박스 이벤트
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
        // 검색 품명 키 이벤트 → 엔터 → 플러스파인더
        private void txtBuyerArticleNo_KeyDown(object sender, KeyEventArgs e)
        {

            //if (e.Key == Key.Enter)
            //{
            //    rowNum = 0;
            //    re_Search(rowNum);
            //}
            if (e.Key == Key.Enter)
            {
                e.Handled = true;
                MainWindow.pf.ReturnCode(txtBuyerArticleNo, 76, "");
            }

        }
        //// 검색 품명 플러스파인더 버튼 이벤트
        //private void btnPfBuyerArticleNo_Click(object sender, RoutedEventArgs e)
        //{
        //    MainWindow.pf.ReturnCode(txtBuyerArticleNo, (int)Defind_CodeFind.DCF_Article, "");
        //}

        // 검색 전 창고 라벨 이벤트
        private void lblFromLocSrh_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            if (chkFromLocSrh.IsChecked == true)
            {
                chkFromLocSrh.IsChecked = false;
            }
            else
            {
                chkFromLocSrh.IsChecked = true;
            }
        }
        // 검색 전 창고 체크박스 이벤트
        private void chkFromLocSrh_Checked(object sender, RoutedEventArgs e)
        {
            chkFromLocSrh.IsChecked = true;

            cboFromLocSrh.IsEnabled = true;
        }
        private void chkFromLocSrh_Unchecked(object sender, RoutedEventArgs e)
        {
            chkFromLocSrh.IsChecked = false;

            cboFromLocSrh.IsEnabled = false;
        }

        // 검색 후 창고 라벨 이벤트
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
        // 검색 후 창고 체크박스 이벤트 
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

        // 검색 이동구분 라벨 버튼 이벤트 → 막음
        private void lblOutClssSrh_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            if (chkOutClssSrh.IsChecked == true)
            {
                chkOutClssSrh.IsChecked = false;
            }
            else
            {
                chkOutClssSrh.IsChecked = true;
            }
        }
        // 검색 이동구분 체크박스 이벤트 → 막음
        private void chkOutClssSrh_Checked(object sender, RoutedEventArgs e)
        {
            chkOutClssSrh.IsChecked = true;

            cboOutClssSrh.IsEnabled = true;
        }
        private void chkOutClssSrh_Unchecked(object sender, RoutedEventArgs e)
        {
            chkOutClssSrh.IsChecked = false;

            cboOutClssSrh.IsEnabled = false;
        }

        #endregion // 상단 왼쪽 검색조건 모음

        #region 상단 오른쪽 버튼 이벤트

        // 추가 버튼 이벤트
        private void btnAdd_Click(object sender, RoutedEventArgs e)
        {
            this.DataContext = null;

            if (dgdSub.Items.Count > 0
                || dgdPart.Items.Count > 0)
            {
                dgdSub.Items.Clear();
                dgdPart.Items.Clear();
            }

            strFlag = "I";
            SaveUpdateMode();

            // 1. 작성일자 오늘날짜
            dtpOutDate.SelectedDate = DateTime.Today;

            // 2. 이동구분 외주이동 선택
            //cboOutClss.SelectedIndex = 0; //이동구문
            //cboFromLoc.SelectedIndex = 0; //전창고
            //cboToLoc.SelectedIndex = 1; // 후창고
            //cboUnitClss.SelectedIndex = 0; //단위


        }
        // 수정 버튼 이벤트
        private void btnUpdate_Click(object sender, RoutedEventArgs e)
        {
            var WinMove = dgdMain.SelectedItem as Win_mtr_Move_Prd_CodeView2;



            if (WinMove != null)
            {
                rowNum = dgdMain.SelectedIndex;
                SaveUpdateMode();
                strFlag = "U";

                txtArticle.Tag = WinMove.ArticleID;
            }
            else
            {
                MessageBox.Show("수정할 자료를 선택해주세요.");
                return;
            }
        }

        // 삭제 버튼 이벤트
        private void btnDelete_Click(object sender, RoutedEventArgs e)
        {
            var WinMove = dgdMain.SelectedItem as Win_mtr_Move_Prd_CodeView2;

            if (WinMove == null)
            {
                MessageBox.Show("삭제할 데이터가 지정되지 않았습니다. 삭제 데이터를 지정하고 눌러주세요.");
            }
            else
            {
                if (MessageBox.Show("선택하신 항목을 삭제하시겠습니까?", "삭제 전 확인", MessageBoxButton.YesNo) == MessageBoxResult.Yes)
                {

                    if (DeleteData(WinMove.OutwareID))
                    {
                        rowNum = 0;
                        re_Search(rowNum);
                    }
                }
            }
        }
        // 닫기 버튼 이벤트
        private void btnClose_Click(object sender, RoutedEventArgs e)
        {
            Lib.Instance.ChildMenuClose(this.ToString());
        }
        // 검색 버튼 이벤트
        private void btnSearch_Click(object sender, RoutedEventArgs e)
        {
            //검색버튼 비활성화
            btnSearch.IsEnabled = false;

            Dispatcher.BeginInvoke(new Action(() =>

            {
                Thread.Sleep(2000);

                //로직
                rowNum = 0;
                re_Search(rowNum);

            }), System.Windows.Threading.DispatcherPriority.Background);



            Dispatcher.BeginInvoke(new Action(() =>

            {
                btnSearch.IsEnabled = true;

            }), System.Windows.Threading.DispatcherPriority.Background);
        }

        // 저장 버튼 이벤트
        private void btnSave_Click(object sender, RoutedEventArgs e)
        {
            if (SaveData())
            {
                CompleteCancelMode();

                //if (cboOutClss.SelectedValue != null)
                //{
                //    cboOutClssSrh.SelectedValue = cboOutClss.SelectedValue;
                //}

                //rowNum = 0;
                re_Search(rowNum);
                strFlag = string.Empty;
            }
        }


        ////저장
        //private void btnSave_Click(object sender, RoutedEventArgs e)
        //{
        //    if (SaveData())
        //    {
        //        CompleteCancelMode();
        //        strBasisID = string.Empty;
        //        lblMsg.Visibility = Visibility.Hidden;

        //        if (strFlag.Equals("I"))
        //        {
        //            InspectName = txtArticle.ToString();
        //            //InspectName = txtKCustom.ToString();
        //            //InspectDate = dtpInspectDate.SelectedDate.ToString().Substring(0, 10);

        //            rowNum = 0;
        //            re_Search(rowNum);
        //        }
        //        else
        //        {
        //            rowNum = dgdMain.SelectedIndex;
        //        }
        //    }

        //    int i = 0;

        //    foreach (Win_mtr_Move_Prd_CodeView2 WMRIC in dgdMain.Items)
        //    {

        //        string a = WMRIC.OrderID.ToString();
        //        string b = InspectName;


        //        if (a == b)
        //        {
        //            System.Diagnostics.Debug.WriteLine("데이터 같음");

        //            break;
        //        }
        //        else
        //        {
        //            System.Diagnostics.Debug.WriteLine("다름");
        //        }

        //        i++;
        //    }

        //    rowNum = i;
        //    re_Search(rowNum);
        //}

        // 취소 버튼 이벤트
        private void btnCancel_Click(object sender, RoutedEventArgs e)
        {
            strFlag = string.Empty;
            CompleteCancelMode();

            //rowNum = 0;
            re_Search(rowNum);
        }
        // 엑셀 버튼 이벤트
        private void btnExcel_Click(object sender, RoutedEventArgs e)
        {
            DataTable dt = null;
            string Name = string.Empty;

            // 부분처리인지 아닌지
            DataGrid dgd = new DataGrid();

            if (dgdPart.Visibility == Visibility.Visible)
            {
                dgd = dgdPart;
            }
            else
            {
                dgd = dgdSub;
            }

            string[] lst = new string[4];
            lst[0] = "외주이동 내역";
            lst[1] = "외주이동 상세내역";
            lst[2] = dgdMain.Name;
            lst[3] = dgd.Name;

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
                else if (ExpExc.choice.Equals(dgd.Name))
                {
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


        #endregion // 상단 오른쪽 버튼 모음

        #endregion // Header 부분

        #region Content 부분

        #region 메인 그리드 모음

        // 메인 그리드 셀렉션 체인지
        private void dgdMain_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            var WinMove = dgdMain.SelectedItem as Win_mtr_Move_Prd_CodeView2;

            if (WinMove != null)
            {
                this.DataContext = WinMove;

                // 1 : ID기준처리, 2 : 수량기준처리, 3 : 부분처리
                // 부분처리일때
                if (WinMove.OutSubType.Trim().Equals("3"))
                {
                    //tgnMoveByID.IsChecked = false;
                    //tgnMoveByQty.IsChecked = false; 2021-08-06 수량으로 이동 주석 처리로 인한 수정
                    tgnMovePartial.IsChecked = true;

                    dgdSub.Visibility = Visibility.Hidden;
                    dgdPart.Visibility = Visibility.Visible;

                    FillGridSub(WinMove);
                }
                else
                {
                    if (WinMove.OutSubType.Trim().Equals("2"))
                    {
                        //tgnMoveByID.IsChecked = false;
                        //tgnMoveByQty.IsChecked = true; 2021-08-06 수량으로 이동 주석 처리로 인한 수정
                        tgnMovePartial.IsChecked = false;
                    }
                    else // if (WinMove.OutSubType.Equals("1"))
                    {
                        //tgnMoveByID.IsChecked = true;
                        //tgnMoveByQty.IsChecked = false; 2021-08-06 수량으로 이동 주석 처리로 인한 수정
                        tgnMovePartial.IsChecked = false;
                    }

                    dgdSub.Visibility = Visibility.Visible;
                    dgdPart.Visibility = Visibility.Hidden;

                    FillGridSub(WinMove);
                }

                // OutRoll : 박스수, 서브그리드 갯수 / OutQty : 총 개수 - 구하기 
                SetOutRollAndOutQty();
            }
        }

        #endregion // 메인 그리드 모음

        #region Content 오른쪽 상세내역 + 바코드

        // ID 기준 이동 토글 버튼
        private void tgnMoveByID_Click(object sender, RoutedEventArgs e)
        {
            //tgnMoveByID.IsChecked = true;
            //tgnMoveByQty.IsChecked = false; 2021-08-06 수량으로 이동 주석 처리로 인한 수정
            tgnMovePartial.IsChecked = false;

            // 수량 입력 안되도록 → 수량기준이동 토글버튼이 활성화 됬을때만 입력 가능하도록
            //txtOutRoll.IsHitTestVisible = false;
            //txtOutQty.IsHitTestVisible = false;

            // 바코드 활성화
            txtBarCode.IsHitTestVisible = true;

            // 그리드 변경
            dgdSub.Visibility = Visibility.Visible;
            dgdPart.Visibility = Visibility.Hidden;

            // OutRoll : 박스수, 서브그리드 갯수 / OutQty : 총 개수 - 구하기 
            SetOutRollAndOutQty();

            // btnAddSub(서브 그리드 추가 버튼) 부분처리 때만 활성화 시키도록 - ID, 수량 기준은 바코드를 통해서만 새로운 행이 추가 가능
            btnAddSub.IsEnabled = false;

            if (!strFlag.Trim().Equals(""))
            {
                txtArticle.IsEnabled = true;
                //btnPfArticle.IsEnabled = true;
            }
        }
        // 수량 기준 이동 토글 버튼
        private void tgnMoveByQty_Click(object sender, RoutedEventArgs e)
        {
            //tgnMoveByID.IsChecked = false;
            //tgnMoveByQty.IsChecked = true; 2021-08-06 수량으로 이동 주석 처리로 인한 수정
            tgnMovePartial.IsChecked = false;

            // 수량 입력 되도록 → 바코드로 입력하도록 막아놓자.
            //txtOutRoll.IsHitTestVisible = false;
            //txtOutQty.IsHitTestVisible = false;

            // 바코드 입력 안되도록 → 수량기준이동은 바코드가 아닌 수량으로 관리
            //txtBarCode.IsHitTestVisible = false;

            // 바코드 활성화
            txtBarCode.IsHitTestVisible = true;

            // 그리드 변경
            dgdSub.Visibility = Visibility.Visible;
            dgdPart.Visibility = Visibility.Hidden;

            // OutRoll : 박스수, 서브그리드 갯수 / OutQty : 총 개수 - 구하기 
            SetOutRollAndOutQty();

            // btnAddSub(서브 그리드 추가 버튼) 부분처리 때만 활성화 시키도록 - ID, 수량 기준은 바코드를 통해서만 새로운 행이 추가 가능
            btnAddSub.IsEnabled = false;

            if (!strFlag.Trim().Equals(""))
            {
                txtArticle.IsEnabled = true;
                //btnPfArticle.IsEnabled = true;
            }
        }
        // 부분 처리 토글 버튼
        private void tgnMovePartial_Click(object sender, RoutedEventArgs e)
        {
            //tgnMoveByID.IsChecked = false;
            //tgnMoveByQty.IsChecked = false; 2021-08-06 수량으로 이동 주석 처리로 인한 수정
            tgnMovePartial.IsChecked = true;

            // 수량 입력 안되도록 → 수량기준이동 토글버튼이 활성화 됬을때만 입력 가능하도록
            //txtOutRoll.IsHitTestVisible = false;
            //txtOutQty.IsHitTestVisible = false;

            // 바코드 활성화
            txtBarCode.IsHitTestVisible = true;

            // 그리드 변경
            dgdSub.Visibility = Visibility.Hidden;
            dgdPart.Visibility = Visibility.Visible;

            // OutRoll : 박스수, 서브그리드 갯수 / OutQty : 총 개수 - 구하기 
            SetOutRollAndOutQty();

            // btnAddSub(서브 그리드 추가 버튼) 부분처리 때만 활성화 시키도록 - ID, 수량 기준은 바코드를 통해서만 새로운 행이 추가 가능
            btnAddSub.IsEnabled = false;

            if (!strFlag.Trim().Equals(""))
            {
                txtArticle.IsEnabled = true;
                //btnPfArticle.IsEnabled = true;
            }
        }

        // 품명 엔터 → 플러스파인더 이벤트
        private void txtArticle_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                e.Handled = true;
                MainWindow.pf.ReturnCode(txtArticle, (int)Defind_CodeFind.DCF_BuyerArticleNo, "");

                if (txtArticle.Tag != null)
                {
                    getArticleInfo(txtArticle.Tag.ToString());
                }
            }
        }
        // 품명 플러스파인더 버튼 이벤트
        private void btnPfArticle_Click(object sender, RoutedEventArgs e)
        {
            MainWindow.pf.ReturnCode(txtArticle, (int)Defind_CodeFind.DCF_BuyerArticleNo, "");

            if (txtArticle.Tag != null)
            {
                getArticleInfo(txtArticle.Tag.ToString());
            }
        }

        #region ArticleID 로 Article 정보 가져오기

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
                            ProductGrpID = dr["ProductGrpID"].ToString()
                        };

                        //cboUnitClss.SelectedValue = getArticleInfo.UnitClss;
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
        #endregion

        // 바코드 클릭시, 바코드 입력할 수 있도록
        private void borderBarcode_MouseLeftButtonUp(object sender, RoutedEventArgs e)
        {
            if (strFlag.Equals("I") || strFlag.Equals("U"))
            {
                txtBarCode.Focus();
            }
        }
        // 바코드 입력 이벤트
        private void txtBarCode_KeyDown(object sender, KeyEventArgs e)
        {
            // Q. ID기준관리는 : 바코드로 관리하는거고 / 수량기준은 수량으로 관리하는거면 > 바코드에 수량 후 엔터키 > 라벨없이 서브그리드에 추가

            // 바코드 체크 > 잘못된 바코드 입니다. (바코드 길이 :  12 또는 14 자리) > 기존에 입력한 바코드도 지우기
            // Enter 키를 눌렀을때 이벤트 발생 되도록 (어차피 바코드 쓰면, 입력하고 엔터처리 되지 않나??)
            if (e.Key == Key.Enter)
            {
                // 여기서 ID 기준 / 수량기준 / 부분처리 구분
                #region ID기준 바코드 유효성 검사 (xp_Outware_sLabelIDOne - 존재하는 데이터만 입력 가능)
                //if (tgnMoveByID.IsChecked == true)
                //{
                if (txtBarCode.Text.Trim().Length < 10)
                {
                    //MessageBox.Show("잘못된 바코드 입니다. (바코드 길이 : 10자리 이상)");
                    MessageBox.Show("ID기준 이동만 가능합니다. 바코드를 입력해주세요 (바코드 길이 : 10자리 이상).");
                    return;
                }

                // 라벨 리스트 가져오는 프로시저로 체크
                ObservableCollection<LabelList3> ovcLableList = new ObservableCollection<LabelList3>();

                try
                {
                    Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                    sqlParameter.Clear();

                    sqlParameter.Add("LabelID", txtBarCode.Text);

                    DataSet ds = DataStore.Instance.ProcedureToDataSet("xp_Outware_sLabelIDOne", sqlParameter, false);

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
                                var labelList = new LabelList3()
                                {
                                    Num = i,

                                    LabelID = dr["LabelID"].ToString(),
                                    LabelGubun = dr["LabelGubun"].ToString(), // 2 - ?, 3 - ?, 7 - 공정이동ID : wk_LabelPrint (라벨 발행 테이블) 에서 가져오는듯??
                                    ArticleID = dr["ArticleID"].ToString(),
                                    Qty = stringFormatN0(dr["QTY"]),
                                    OutSideQTY = stringFormatN0(dr["OutSideQTY"]),


                                    LabelGubunName = dr["LabelGubunName"].ToString(),

                                    InspectApprovalYN = dr["InspectApprovalYN"].ToString(),
                                    Inspector = dr["Inspector"].ToString(),
                                    Article = dr["Article"].ToString(),
                                    ProcessID = dr["ProcessID"].ToString(),
                                    CustomID = dr["CustomID"].ToString(),

                                    Custom = dr["Custom"].ToString(),
                                    UnitClss = dr["UnitClss"].ToString(),
                                    OutClss = dr["OutClss"].ToString()

                                };

                                //if (cboOutClss.SelectedValue.Equals("15"))
                                //{
                                //    labelList.Qty = labelList.OutSideQTY;
                                //}
                                //else if (cboOutClss.SelectedValue.Equals("05"))
                                //{
                                //    labelList.Qty = labelList.Qty;
                                //}

                                ovcLableList.Add(labelList);
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

                // 존재하지 않는 바코드 입니다.
                if (ovcLableList.Count <= 0)
                {
                    MessageBox.Show("존재하지 않는 바코드 입니다.");
                    txtBarCode.Text = "";
                    return;
                }
                else
                {
                    if (ovcLableList[0].InspectApprovalYN == null
                        || !ovcLableList[0].InspectApprovalYN.Equals("Y"))
                    {
                        MessageBox.Show("승인되지 않은 건은 이동처리 불가합니다.");
                        txtBarCode.Text = "";
                        return;
                    }

                    // 이때 품명을 넣어주자. 2019-11-25 바코드를 입력했을때 품명 자동으로 넣어줘! 라고 요청함
                    // 서브 그리드에 아무것도 없는 경우에 최초로 등록[텍스트박스의 Text : 품명 / Tag : 품명ID] 후
                    // → 서브 그리드에 값이 있는 경우, TextBox 의 걸로 비교하기에는 문제가 (만약에 수량기준으로 서브그리드에 하나를 등록했다면, 품명 정보가 없기에 비교가 불가능)
                    // → 서브 그리드에 품명 컬럼을 추가 + for문을 돌면서 품명있는 것(하나 나오면 break)과 비교 하여, 다른 품목은 입력 못하도록 막는걸로 시도해봅세 

                    // Q 2 : 만약에 바코드를 하나 등록했는데, 품명을 변경한다면?
                    // 1) 바코드 성공, 품명 등록시 TextBox를 isEnabled = false 처리 
                    //   → 그럼 삭제시에도 for문을 돌아서, 품명 정보가 없으면 TextBox를 isEnabled = true 처리 해줘야됨.
                    //       그럼 수정 시에도 for문 돌면서 품명을 못쓰게 해야 겠네.
                    // 2) 
                    if (dgdSub.Items.Count < 1)
                    {
                        txtArticle.Text = ovcLableList[0].Article;
                        txtArticle.Tag = ovcLableList[0].ArticleID;

                        // 바코드가 입력되면, 품명 수정은 불가능 하도록.
                        txtArticle.IsEnabled = false;
                        //btnPfArticle.IsEnabled = false;
                    }
                    else
                    {
                        string ArticleID = "";

                        for (int k = 0; k < dgdSub.Items.Count; k++)
                        {
                            var MoveSub = dgdSub.Items[k] as LabelList3;

                            if (MoveSub != null)
                            {
                                // 일단 품명 ID 체크
                                if (MoveSub.ArticleID != null && !MoveSub.ArticleID.Trim().Equals(""))
                                {
                                    ArticleID = MoveSub.ArticleID;
                                }

                                if (MoveSub.LabelID != null && MoveSub.LabelID.Trim().Equals(txtBarCode.Text))
                                {
                                    MessageBox.Show("이미 스캔된 바코드입니다.");
                                    txtBarCode.Text = "";
                                    return;
                                }
                            }
                        }

                        if (ArticleID.Equals("")) // 품명이 없다면, 최초등록 → TextBox에 세팅
                        {
                            txtArticle.Text = ovcLableList[0].Article;
                            txtArticle.Tag = ovcLableList[0].ArticleID;

                            // 바코드가 입력되면, 품명 수정은 불가능 하도록.
                            txtArticle.IsEnabled = false;
                            //btnPfArticle.IsEnabled = false;

                            //cboUnitClss.SelectedValue = ovcLableList[0].UnitClss;
                        }
                        else // 품명이 있다면 → 지금 입력하려는 바코드 품명과 비교
                        {
                            if (!ArticleID.Trim().Equals(ovcLableList[0].ArticleID))
                            {
                                MessageBox.Show("서로 다른 품명을 동시에 출고처리 할 수 없습니다.");
                                txtBarCode.Text = "";
                                return;
                            }
                        }
                    }

                    for (int i = 0; i < ovcLableList.Count; i++)
                    {
                        // 순번을 순서대로 넣기 위해서 사용
                        int index = dgdSub.Items.Count + 1;
                        ovcLableList[i].Num = index;
                        dgdSub.Items.Add(ovcLableList[i]);
                    }

                    txtBarCode.Text = "";


                    //if (txtArticle.Tag == null || txtArticle.Tag.ToString().Equals("")
                    //    || ovcLableList[0].ArticleID == null || !txtArticle.Tag.ToString().Trim().Equals(ovcLableList[0].ArticleID))
                    //{
                    //    MessageBox.Show("서로 다른 품명을 동시에 출고처리 할 수 없습니다.");
                    //    txtBarCode.Text = "";
                    //    return;
                    //}

                    //bool isAddFlag = true;
                    //for (int i = 0; i < dgdSub.Items.Count; i++)
                    //{
                    //    var compareLabel = dgdSub.Items[i] as Win_mtr_Move_Prd_CodeViewSub;

                    //    if (compareLabel.LabelID.Trim().Equals(txtBarCode.Text))
                    //    {
                    //        MessageBox.Show("이미 스캔된 바코드입니다.");
                    //        txtBarCode.Text = "";
                    //        isAddFlag = false;
                    //        break;
                    //    }
                    //}

                    //if (isAddFlag)
                    //{
                    //    for (int i = 0; i < ovcLableList.Count; i++)
                    //    {
                    //        dgdSub.Items.Add(ovcLableList[i]);
                    //    }

                    //    txtBarCode.Text = "";
                    //}
                }
                //}
                #endregion // ID기준 바코드 유효성 검사 (xp_Outware_sLabelIDOne - 존재하는 데이터만 입력 가능)
                //else if (tgnMoveByQty.IsChecked == true) 2021-08-06 수량으로 이동 주석 처리로 인한 수정
                //{
                //    // 바코드에 수량을 입력 → 숫자만 입력 가능하도록 유효성 검사
                //    if (CheckConvertInt(txtBarCode.Text))
                //    {
                //        // 수량 입력시 라벨 없이 입력됨
                //        LabelList3 label = new LabelList3();

                //        int num = dgdSub.Items.Count + 1;
                //        label.Num = num;

                //        label.Qty = stringFormatN0(txtBarCode.Text);
                //        dgdSub.Items.Add(label);

                //        // 데이터 그리드 등록 후 바코드 초기화
                //        txtBarCode.Text = "";
                //    }
                //    else
                //    {
                //        MessageBox.Show("수량 등록에는 숫자만 입력 가능합니다.");
                //    }
                //}
                //else // 부분처리
                //    {
                //if (txtBarCode.Text.Trim().Length != 12 && txtBarCode.Text.Trim().Length != 14)
                //{
                //    MessageBox.Show("잘못된 바코드 입니다. (바코드 길이 : 12 또는 14자리)");
                //    return;
                //}


                #region 바코드 유효성 검사
                if (txtBarCode.Text.Trim().Length < 10)
                {
                    MessageBox.Show("잘못된 바코드 입니다. (바코드 길이 : 10자리 이상)");
                    return;
                }

                // 라벨 리스트 가져오는 프로시저로 체크
                //ObservableCollection<LabelList3> ovcLableList = new ObservableCollection<LabelList3>();

                try
                {
                    Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                    sqlParameter.Clear();

                    sqlParameter.Add("LabelID", txtBarCode.Text);

                    DataSet ds = DataStore.Instance.ProcedureToDataSet("xp_Outware_sLabelIDOne", sqlParameter, false);

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
                                var labelList = new LabelList3()
                                {
                                    Num = i,

                                    LabelID = dr["LabelID"].ToString(),
                                    LabelGubun = dr["LabelGubun"].ToString(), // 2 - ?, 3 - ?, 7 - 공정이동ID : wk_LabelPrint (라벨 발행 테이블) 에서 가져오는듯??
                                    ArticleID = dr["ArticleID"].ToString(),
                                    Qty = stringFormatN0(dr["QTY"]),
                                    LabelGubunName = dr["LabelGubunName"].ToString(),

                                    InspectApprovalYN = dr["InspectApprovalYN"].ToString(),
                                    Inspector = dr["Inspector"].ToString(),
                                    Article = dr["Article"].ToString(),
                                    ProcessID = dr["ProcessID"].ToString(),
                                    CustomID = dr["CustomID"].ToString(),

                                    Custom = dr["Custom"].ToString(),
                                    UnitClss = dr["UnitClss"].ToString(),

                                    OutClss = dr["OutClss"].ToString()
                                };

                                //if (cboOutClss.SelectedValue.Equals("15"))
                                //{
                                //    labelList.Qty = labelList.OutSideQTY;
                                //}
                                //else if (cboOutClss.SelectedValue.Equals("05"))
                                //{
                                //    labelList.Qty = labelList.Qty;
                                //}

                                ovcLableList.Add(labelList);
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

                // 존재하지 않는 바코드 입니다.
                if (ovcLableList.Count <= 0)
                {
                    MessageBox.Show("존재하지 않는 바코드 입니다.");
                    txtBarCode.Text = "";
                    return;
                }
                else
                {
                    if (ovcLableList[0].InspectApprovalYN == null
                        || !ovcLableList[0].InspectApprovalYN.Equals("Y"))
                    {
                        MessageBox.Show("승인되지 않은 건은 이동처리 불가합니다.");
                        txtBarCode.Text = "";
                        return;
                    }

                    // 이때 품명을 넣어주자. 2019-11-25 바코드를 입력했을때 품명 자동으로 넣어줘! 라고 요청함
                    // 서브 그리드에 아무것도 없는 경우에 최초로 등록[텍스트박스의 Text : 품명 / Tag : 품명ID] 후
                    // → 서브 그리드에 값이 있는 경우, TextBox 의 걸로 비교하기에는 문제가 (만약에 수량기준으로 서브그리드에 하나를 등록했다면, 품명 정보가 없기에 비교가 불가능)
                    // → 서브 그리드에 품명 컬럼을 추가 + for문을 돌면서 품명있는 것(하나 나오면 break)과 비교 하여, 다른 품목은 입력 못하도록 막는걸로 시도해봅세 

                    // Q 2 : 만약에 바코드를 하나 등록했는데, 품명을 변경한다면?
                    // 1) 바코드 성공, 품명 등록시 TextBox를 isEnabled = false 처리 
                    //   → 그럼 삭제시에도 for문을 돌아서, 품명 정보가 없으면 TextBox를 isEnabled = true 처리 해줘야됨.
                    //       그럼 수정 시에도 for문 돌면서 품명을 못쓰게 해야 겠네.
                    // 2) 
                    if (dgdPart.Items.Count < 1)
                    {
                        txtArticle.Text = ovcLableList[0].Article;
                        txtArticle.Tag = ovcLableList[0].ArticleID;

                        // 바코드가 입력되면, 품명 수정은 불가능 하도록.
                        txtArticle.IsEnabled = false;
                        //btnPfArticle.IsEnabled = false;
                    }
                    else
                    {
                        string ArticleID = "";

                        for (int k = 0; k < dgdPart.Items.Count; k++)
                        {
                            var MoveSub = dgdPart.Items[k] as LabelList3;

                            if (MoveSub != null)
                            {
                                // 일단 품명 ID 체크
                                if (MoveSub.ArticleID != null && !MoveSub.ArticleID.Trim().Equals(""))
                                {
                                    ArticleID = MoveSub.ArticleID;
                                }

                                if (MoveSub.LabelID != null && MoveSub.LabelID.Trim().Equals(txtBarCode.Text))
                                {
                                    MessageBox.Show("이미 스캔된 바코드입니다.");
                                    txtBarCode.Text = "";
                                    return;
                                }
                            }
                        }

                        if (ArticleID.Equals("")) // 품명이 없다면, 최초등록 → TextBox에 세팅
                        {
                            txtArticle.Text = ovcLableList[0].Article;
                            txtArticle.Tag = ovcLableList[0].ArticleID;

                            // 바코드가 입력되면, 품명 수정은 불가능 하도록.
                            txtArticle.IsEnabled = false;
                            //btnPfArticle.IsEnabled = false;

                            //cboUnitClss.SelectedValue = ovcLableList[0].UnitClss;
                        }
                        else // 품명이 있다면 → 지금 입력하려는 바코드 품명과 비교
                        {
                            if (!ArticleID.Trim().Equals(ovcLableList[0].ArticleID))
                            {
                                MessageBox.Show("서로 다른 품명을 동시에 출고처리 할 수 없습니다.");
                                txtBarCode.Text = "";
                                return;
                            }
                        }
                    }

                    for (int i = 0; i < ovcLableList.Count; i++)
                    {
                        // 순번을 순서대로 넣기 위해서 사용
                        int index = dgdPart.Items.Count + 1;
                        ovcLableList[i].Num = index;
                        dgdPart.Items.Add(ovcLableList[i]);
                    }

                    txtBarCode.Text = "";

                    //if (txtArticle.Tag == null || txtArticle.Tag.ToString().Equals("")
                    //    || ovcLableList[0].ArticleID == null || !txtArticle.Tag.ToString().Trim().Equals(ovcLableList[0].ArticleID))
                    //{
                    //    MessageBox.Show("서로 다른 품명을 동시에 출고처리 할 수 없습니다.");
                    //    txtBarCode.Text = "";
                    //    return;
                    //}

                    //bool isAddFlag = true;
                    //for (int i = 0; i < dgdSub.Items.Count; i++)
                    //{
                    //    var compareLabel = dgdSub.Items[i] as Win_mtr_Move_Prd_CodeViewSub;

                    //    if (compareLabel.LabelID.Trim().Equals(txtBarCode.Text))
                    //    {
                    //        MessageBox.Show("이미 스캔된 바코드입니다.");
                    //        txtBarCode.Text = "";
                    //        isAddFlag = false;
                    //        break;
                    //    }
                    //}

                    //if (isAddFlag)
                    //{
                    //    for (int i = 0; i < ovcLableList.Count; i++)
                    //    {
                    //        dgdSub.Items.Add(ovcLableList[i]);
                    //    }

                    //    txtBarCode.Text = "";
                    //}
                }

                #endregion


                // 부분 처리도 바코드 검색 안해도 되나?? ??? 
                // 바코드 입력 및 엔터 > 그 라벨로 새로운 행 생성
                //var label = new LabelList3();
                //label.LabelID = txtBarCode.Text;
                //label.Num = dgdPart.Items.Count + 1;

                //dgdPart.Items.Add(label);

                // 데이터 그리드 등록 후 바코드 초기화
                txtBarCode.Text = "";
                //}

                SetOutRollAndOutQty();
                SetDdgNum();
            }
        }
        private void SetDdgNum()
        {
            for (int i = 0; i < dgdSub.Items.Count; i++)
            {
                var MoveSub = dgdSub.Items[i] as LabelList3;

                if (MoveSub != null)
                {
                    MoveSub.Num = i + 1;
                }
            }

            for (int i = 0; i < dgdPart.Items.Count; i++)
            {
                var MoveSub = dgdPart.Items[i] as LabelList3;

                if (MoveSub != null)
                {
                    MoveSub.Num = i + 1;
                }
            }
        }

        // OutRoll : 박스수, 서브그리드 갯수 / OutQty : 총 개수 - 구하기 
        private void SetOutRollAndOutQty()
        {
            int OutRoll = 0;
            double OutQty = 0;

            if (tgnMovePartial.IsChecked == true)
            {
                OutRoll = dgdPart.Items.Count;

                for (int i = 0; i < dgdPart.Items.Count; i++)
                {
                    var label = dgdPart.Items[i] as LabelList3;
                    if (label.Qty != null)
                        OutQty += ConvertDouble(label.Qty.ToString());
                }
            }
            else
            {
                OutRoll = dgdSub.Items.Count;

                for (int i = 0; i < dgdSub.Items.Count; i++)
                {
                    var label = dgdSub.Items[i] as LabelList3;
                    if (label.Qty != null)
                        OutQty += ConvertDouble(label.Qty.ToString());
                }
            }

            //txtOutRoll.Text = stringFormatN0(OutRoll);
            //txtOutQty.Text = stringFormatN0(OutQty);
        }

        private void txtBarCode_KeyUp(object sender, KeyEventArgs e)
        {
            // 수량 입력시에만
            //if (tgnMoveByQty.IsChecked == true) 2021-08-06 수량으로 이동 주석 처리로 인한 수정
            //{
            //    // 바코드에 입력이 되있을 때만 실행 되도록 → 품명을 입력하지 않았을때, 바코드를 비워주기 위해서. 
            //    if (!txtBarCode.Text.Trim().Equals(""))
            //    {
            //        // 품명이 입력되지 않았을때는 바코드 입력 못함 > 품명을 먼저 입력해주세요.
            //        if (txtArticle.Tag == null || txtArticle.Text.Trim().Equals(""))
            //        {
            //            MessageBox.Show("품명을 먼저 입력해주세요.");
            //            txtBarCode.Text = "";
            //            return;
            //        }
            //    }
            //}
        }

        #endregion //  Content 오른쪽 상세내역 + 바코드

        #region 서브 그리드 모음

        // 서브그리드 추가, 삭제 이벤트
        private void btnAddSub_Click(object sender, RoutedEventArgs e)
        {
            // 부분처리 라면
            if (tgnMovePartial.IsChecked == true)
            {
                var label = new LabelList3();
                label.Num = dgdPart.Items.Count + 1;

                dgdPart.Items.Add(label);
            }
            else // 그 외
            {
                var label = new LabelList3();
                label.Num = dgdSub.Items.Count + 1;

                dgdSub.Items.Add(label);
            }
        }
        private void btnDeleteSub_Click(object sender, RoutedEventArgs e)
        {
            // 새로운 셀을 추가할건지 메시지를 띄우고, 추가
            if (MessageBox.Show("해당 데이터를 삭제하시겠습니까?", "삭제 전 확인", MessageBoxButton.YesNo) == MessageBoxResult.Yes)
            {
                // 부분처리 라면
                if (tgnMovePartial.IsChecked == true)
                {
                    int sIndex = dgdPart.SelectedIndex - 1;
                    if (sIndex < 0) { sIndex = 0; }

                    var label = dgdPart.SelectedItem as LabelList3;
                    dgdPart.Items.Remove(label);

                    if (sIndex < dgdPart.Items.Count)
                    {
                        dgdPart.SelectedIndex = sIndex;
                    }
                }
                else // 그 외
                {
                    int sIndex = dgdSub.SelectedIndex - 1;
                    if (sIndex < 0) { sIndex = 0; }

                    var label = dgdSub.SelectedItem as LabelList3;
                    dgdSub.Items.Remove(label);

                    if (sIndex < dgdSub.Items.Count)
                    {
                        dgdSub.SelectedIndex = sIndex;
                    }
                }
            }

            //SaveData();
        }

        // 부분 처리 그리드 엔터 → 플러스 파인더 이벤트
        private void txtDefect_KeyDown(object sender, KeyEventArgs e)
        {
            var WinMoveSub = dgdPart.CurrentItem as LabelList3;

            if (WinMoveSub != null)
            {
                if (lblMsg.Visibility == Visibility.Visible)
                {
                    if (e.Key == Key.Enter)
                    {
                        e.Handled = true;
                        TextBox tb1 = sender as TextBox;

                        PlusFinder pf = new PlusFinder();

                        pf.ReturnCode(tb1, (int)Defind_CodeFind.DCF_DEFECT, "");

                        if (tb1.Tag != null)
                        {
                            WinMoveSub.DefectID = tb1.Tag.ToString();
                            WinMoveSub.DefectName = tb1.Text;
                        }
                    }
                    else if ((sender as TextBox).Text.Trim().Equals(""))
                    {
                        WinMoveSub.DefectID = "";
                        WinMoveSub.DefectName = "";
                    }
                }
            }
        }

        #region 데이터그리드 이벤트 → 입고 수량(Column : 2)은 패스 

        // 2019.08.27 PreviewKeyDown 는 key 다운과 같은것 같음
        private void DataGird_PreviewKeyDown(object sender, KeyEventArgs e)
        {
            try
            {
                if (e.Key == Key.Down || e.Key == Key.Up || e.Key == Key.Left || e.Key == Key.Right)
                {
                    DataGird_KeyDown(sender, e);
                }
            }
            catch (Exception ex)
            {

            }
        }

        // KeyDown 이벤트
        private void DataGird_KeyDown(object sender, KeyEventArgs e)
        {
            int currRow = dgdPart.Items.IndexOf(dgdPart.CurrentItem);
            int currCol = dgdPart.Columns.IndexOf(dgdPart.CurrentCell.Column);
            int startCol = 3;
            int endCol = 6;

            if (e.Key == Key.Enter)
            {
                e.Handled = true;
                (sender as DataGridCell).IsEditing = false;

                // 마지막 열, 마지막 행 아님
                if (endCol == currCol && dgdPart.Items.Count - 1 > currRow)
                {
                    dgdPart.SelectedIndex = currRow + 1; // 이건 한줄 파란색으로 활성화 된 걸 조정하는 것입니다.
                    dgdPart.CurrentCell = new DataGridCellInfo(dgdPart.Items[currRow + 1], dgdPart.Columns[startCol]);
                } // 마지막 열 아님
                else if (endCol > currCol && dgdPart.Items.Count - 1 >= currRow)
                {
                    //if (currCol == 1) currCol++; // 2는 건너뛰기
                    dgdPart.CurrentCell = new DataGridCellInfo(dgdPart.Items[currRow], dgdPart.Columns[currCol + 1]);
                } // 마지막 열, 마지막 행
                else if (endCol == currCol && dgdPart.Items.Count - 1 == currRow)
                {
                    //// 새로운 셀을 추가할건지 메시지를 띄우고, 추가
                    //if (MessageBox.Show("새로운 행을 추가 하시겠습니까?", "추가 전 확인", MessageBoxButton.YesNo) == MessageBoxResult.Yes)
                    //{
                    //    var label = new LabelList3();
                    //    label.Num = dgdPart.Items.Count + 1;

                    //    dgdPart.Items.Add(label);
                    //}
                }
                else
                {
                    //MessageBox.Show("나머지가 있나..");
                }
            }
            else if (e.Key == Key.Down)
            {
                e.Handled = true;
                (sender as DataGridCell).IsEditing = false;

                if (dgdPart.Items.Count - 1 > currRow)
                {
                    dgdPart.SelectedIndex = currRow + 1;
                    dgdPart.CurrentCell = new DataGridCellInfo(dgdPart.Items[currRow + 1], dgdPart.Columns[currCol]);
                }
            }
            else if (e.Key == Key.Up)
            {
                e.Handled = true;
                (sender as DataGridCell).IsEditing = false;

                if (currRow > 0)
                {
                    dgdPart.SelectedIndex = currRow - 1;
                    dgdPart.CurrentCell = new DataGridCellInfo(dgdPart.Items[currRow - 1], dgdPart.Columns[currCol]);
                }
            }
            else if (e.Key == Key.Left)
            {
                e.Handled = true;
                (sender as DataGridCell).IsEditing = false;

                // 
                if (startCol < currCol)
                {
                    dgdPart.CurrentCell = new DataGridCellInfo(dgdPart.Items[currRow], dgdPart.Columns[currCol - 1]);
                }
                else if (startCol == currCol)
                {
                    if (0 < currRow)
                    {
                        dgdPart.SelectedIndex = currRow - 1;
                        dgdPart.CurrentCell = new DataGridCellInfo(dgdPart.Items[currRow - 1], dgdPart.Columns[endCol]);
                    }
                    else
                    {
                        btnSave.Focus();
                    }
                }
            }
            else if (e.Key == Key.Right)
            {
                e.Handled = true;
                (sender as DataGridCell).IsEditing = false;

                if (endCol > currCol)
                {
                    dgdPart.CurrentCell = new DataGridCellInfo(dgdPart.Items[currRow], dgdPart.Columns[currCol + 1]);
                }
                else if (endCol == currCol)
                {
                    if (dgdPart.Items.Count - 1 > currRow)
                    {
                        dgdPart.SelectedIndex = currRow + 1;
                        dgdPart.CurrentCell = new DataGridCellInfo(dgdPart.Items[currRow + 1], dgdPart.Columns[startCol]);
                    }
                    else
                    {
                        btnSave.Focus();
                    }
                }
            }
            else if (e.Key == Key.Delete)
            {
                e.Handled = true;
                (sender as DataGridCell).IsEditing = false;

                // 새로운 셀을 추가할건지 메시지를 띄우고, 추가
                if (MessageBox.Show("해당 데이터를 삭제하시겠습니까?", "삭제 전 확인", MessageBoxButton.YesNo) == MessageBoxResult.Yes)
                {
                    var label = dgdPart.SelectedItem as LabelList3;
                    dgdPart.Items.Remove(label);
                }

            }
        }
        // KeyUp 이벤트
        private void DatagridIn_TextFocus(object sender, KeyEventArgs e)
        {
            // 엔터 → 포커스 = true → cell != null → 해당 텍스트박스가 null이 아니라면 
            // → 해당 텍스트박스가 포커스가 안되있음 SelectAll() or 포커스
            Lib.Instance.DataGridINTextBoxFocus(sender, e);
        }
        // GotFocus 이벤트
        private void DataGridCell_GotFocus(object sender, RoutedEventArgs e)
        {
            if (lblMsg.Visibility == Visibility.Visible)
            {
                int currRow = dgdPart.Items.IndexOf(dgdPart.CurrentItem);

                DataGridCell cell = sender as DataGridCell;
                cell.IsEditing = true;
            }
        }

        // 2019.08.27 MouseUp 이벤트
        private void DataGridCell_MouseUp(object sender, MouseButtonEventArgs e)
        {
            Lib.Instance.DataGridINTextBoxFocusByMouseUP(sender, e);
        }


        //private void DataGridCell_LostFocus(object sender, RoutedEventArgs e)
        //{
        //    // Qty 세팅하기 (정상, 샘플, 불량 합계 구하기)
        //    var labelSum = dgdPart.SelectedItem as LabelList3;
        //    if (labelSum != null)
        //    {
        //        labelSum.Qty = stringFormatN0(ConvertDouble(labelSum.NQty == null ? "" : labelSum.NQty)
        //            + ConvertDouble(labelSum.SQty == null ? "" : labelSum.SQty)
        //            + ConvertDouble(labelSum.DQty == null ? "" : labelSum.DQty));

        //        // OutRoll : 박스수, 서브그리드 갯수 / OutQty : 총 개수 - 구하기 
        //        SetOutRollAndOutQty();
        //    }

        //}

        #endregion // 데이터 그리드 키 이벤트

        #endregion // 서브 그리드 모음

        #endregion // Content 부분

        #region 주요 메서드 모음

        private void re_Search(int rowNum)
        {
            FillGrid();

            if (dgdMain.Items.Count > 0)
            {
                dgdMain.SelectedIndex = strFlag.Trim().Equals("I") ? dgdMain.Items.Count - 1 : rowNum;
            }
            else
            {
                this.DataContext = null;
                MessageBox.Show("조회된 데이터가 없습니다.");
                return;
            }
        }

        #region 조회 메서드 FillGrid()

        // 조회 검색 메서드
        private void FillGrid()
        {
            //var LabelIst = new LabelList3();

            // 입고량, 입고건수 - 합계 구하기
            var SumStuffIN = new Win_mtr_ocStuffIN_Sum();

            if (dgdMain.Items.Count > 0)
            {
                dgdMain.Items.Clear();

                dgdSub.Items.Clear();
                dgdPart.Items.Clear();
            }


            try
            {
                Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                sqlParameter.Clear();

                sqlParameter.Add("ChkDate", chkDateSrh.IsChecked == true ? 1 : 0);
                sqlParameter.Add("SDate", dtpSDateSrh.SelectedDate != null ? dtpSDateSrh.SelectedDate.Value.ToString("yyyyMMdd") : "");
                sqlParameter.Add("EDate", dtpEDateSrh.SelectedDate != null ? dtpEDateSrh.SelectedDate.Value.ToString("yyyyMMdd") : "");
                sqlParameter.Add("ChkCustomID", 0);
                sqlParameter.Add("CustomID", "");

                sqlParameter.Add("Custom", "");
                sqlParameter.Add("ChkArticleID", 0); // chkBuyerArticleNo.IsChecked == true ? 1 : 0);
                sqlParameter.Add("ArticleID", ""); // chkBuyerArticleNo.IsChecked == true ? (txtBuyerArticleNo.Tag != null ? txtBuyerArticleNo.Tag.ToString() : "") : "");
                sqlParameter.Add("Article", "");// chkBuyerArticleNo.IsChecked == true ? txtArticle.Text : "");

                sqlParameter.Add("ChkOrder", 0);

                sqlParameter.Add("Order", "");
                sqlParameter.Add("OutFlag", 0);
                sqlParameter.Add("OutClss", chkOutClssSrh.IsChecked == true ? (cboOutClssSrh.SelectedValue != null ? cboOutClssSrh.SelectedValue.ToString() : "") : "");      // 이동구분
                sqlParameter.Add("FromLocID", chkFromLocSrh.IsChecked == true ? (cboFromLocSrh.SelectedValue != null ? cboFromLocSrh.SelectedValue.ToString() : "") : "");
                sqlParameter.Add("ToLocID", chkToLocSrh.IsChecked == true ? (cboToLocSrh.SelectedValue != null ? cboToLocSrh.SelectedValue.ToString() : "") : ""); // 후 창고

                sqlParameter.Add("BuyerDirectYN", "");
                sqlParameter.Add("nBuyerArticleNo", chkBuyerArticleNo.IsChecked == true ? 1 : 0);
                sqlParameter.Add("BuyerArticleNo", chkBuyerArticleNo.IsChecked == true && !txtBuyerArticleNo.Text.Trim().Equals("") ? txtBuyerArticleNo.Text : "");

                DataSet ds = DataStore.Instance.ProcedureToDataSet("xp_Outware_sOrder", sqlParameter, false);

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
                            var WinMove = new Win_mtr_Move_Prd_CodeView2()
                            {
                                Num = i,

                                OutwareID = dr["OutwareID"].ToString(),
                                OutSeq = dr["OutSeq"].ToString(),

                                OrderID = dr["OrderID"].ToString(),
                                //OrderSeq = dr["OrderSeq"].ToString(),
                                OrderNo = dr["OrderNo"].ToString(),
                                CustomID = dr["CustomID"].ToString(),
                                CustomName = dr["CustomName"].ToString(),

                                KCustom = dr["KCustom"].ToString(),
                                OutDate = dr["OutDate"].ToString(),
                                ArticleID = dr["ArticleID"].ToString(),
                                Article = dr["Article"].ToString(),
                                OutClss = dr["OutClss"].ToString(),

                                WorkID = dr["WorkID"].ToString(),
                                OutRoll = stringFormatN0(dr["OutRoll"]),
                                OutQty = stringFormatN0(dr["OutQty"]),
                                OutRealQty = dr["OutRealQty"].ToString(),
                                ResultDate = dr["ResultDate"].ToString(),

                                OrderQty = dr["OrderQty"].ToString(),
                                UnitClss = dr["UnitClss"].ToString().Trim(),
                                //UnitClss = dr["UnitClss"].ToString(),
                                WorkName = dr["WorkName"].ToString(),
                                OutType = dr["OutType"].ToString(),
                                Remark = dr["Remark"].ToString(),

                                BuyerModel = dr["BuyerModel"].ToString(),
                                OutSumQty = dr["OutSumQty"].ToString(),
                                OutQtyY = dr["OutQtyY"].ToString(),
                                StuffInQty = stringFormatN0(dr["StuffInQty"]),
                                OutWeight = dr["OutWeight"].ToString(),

                                OutRealWeight = dr["OutRealWeight"].ToString(),
                                UnitPriceClss = dr["UnitPriceClss"].ToString(),
                                BuyerDirectYN = dr["BuyerDirectYN"].ToString(),
                                Vat_Ind_YN = dr["Vat_Ind_YN"].ToString(),
                                InsStuffINYN = dr["InsStuffINYN"].ToString(),

                                ExchRate = dr["ExchRate"].ToString(),
                                FromLocID = dr["FromLocID"].ToString(),
                                ToLocID = dr["ToLocID"].ToString(),
                                UnitClssName = dr["UnitClssName"].ToString().Trim(),
                                FromLocName = dr["FromLocName"].ToString(),

                                TOLocname = dr["TOLocname"].ToString(),
                                OutClssname = dr["OutClssname"].ToString(),
                                UnitPrice = dr["UnitPrice"].ToString(),
                                Amount = dr["Amount"].ToString(),
                                VatAmount = dr["VatAmount"].ToString(),

                                BuyerArticleNo = dr["BuyerArticleNo"].ToString(),
                                OutCustomID = dr["OutCustomID"].ToString(),
                                BuyerID = dr["BuyerID"].ToString(),
                                BuyerName = dr["BuyerName"].ToString(),
                                OutCustom = dr["OutCustom"].ToString(),

                                OutSubType = dr["OutSubType"].ToString() // OutSubType - 1:ID기준, 2:수량기준, 3:부분처리
                            };




                            //if (cboOutClss.SelectedValue.Equals("15"))
                            //{
                            //    WinMove.OutQty = LabelIst.Qty;


                            //}
                            //else if (cboOutClss.SelectedValue.Equals("05"))
                            //{
                            //    WinMove.OutQty = LabelIst.OutSideQTY;
                            //}


                            // 입고일자
                            WinMove.OutDate_CV = DatePickerFormat(WinMove.OutDate);
                            WinMove.ResultDate_CV = DatePickerFormat(WinMove.ResultDate);

                            dgdMain.Items.Add(WinMove);
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

        #endregion // 조회 메서드 FillGrid()

        #region 서브 그리드 조회 메서드 FillGridSub()

        private void FillGridSub(Win_mtr_Move_Prd_CodeView2 WinMove)
        {
            LabelList3 tempLabel = new LabelList3();
            tempLabel.LabelID = "";
            tempLabel.Gubun = "";
            tempLabel.NQty = "";
            tempLabel.SQty = "";
            tempLabel.DQty = "";

            List<LabelList3> lstLabel = new List<LabelList3>();

            // 입고량, 입고건수 - 합계 구하기
            var SumStuffIN = new Win_mtr_ocStuffIN_Sum();

            if (dgdSub.Items.Count > 0
                || dgdPart.Items.Count > 0)
            {
                dgdSub.Items.Clear();
                dgdPart.Items.Clear();
            }

            try
            {
                Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                sqlParameter.Clear();

                sqlParameter.Add("OutwareID", WinMove.OutwareID);

                DataSet ds = DataStore.Instance.ProcedureToDataSet("xp_Outware_sOutwareSubGroup", sqlParameter, false);

                if (ds != null && ds.Tables.Count > 0)
                {
                    DataTable dt = ds.Tables[0];

                    if (dt.Rows.Count > 0)
                    {
                        int i = 0;
                        int q = 0;
                        DataRowCollection drc = dt.Rows;

                        foreach (DataRow dr in drc)
                        {
                            i++;
                            var WinMoveSub = new LabelList3()
                            {
                                Num = i,
                                LabelID = dr["LabelID"].ToString(),
                                LabelGubun = dr["LabelGubun"].ToString(), // 2 - ?, 3 - ?, 7 - 공정이동ID : wk_LabelPrint (라벨 발행 테이블) 에서 가져오는듯??
                                ArticleID = dr["ArticleID"].ToString(),
                                Qty = stringFormatN0(dr["OutQty"]),
                                LabelGubunName = dr["LabelGubunName"].ToString(),

                                //InspectApprovalYN = dr["InspectApprovalYN"].ToString(),
                                //Inspector = dr["Inspector"].ToString(),
                                Article = dr["Article"].ToString(),
                                //ProcessID = dr["ProcessID"].ToString(),
                                //CustomID = dr["CustomID"].ToString(),

                                //Custom = dr["Custom"].ToString(),
                                UnitClss = dr["UnitClss"].ToString(),
                                OutClss = dr["OutClss"].ToString(),

                                DefectID = dr["DefectID"].ToString(), // 불량코드
                                DefectName = dr["DefectName"].ToString(),
                                Gubun = dr["Gubun"].ToString()

                            };

                            // OutSubType - 1:ID기준, 2:수량기준, 3:부분처리
                            // 부분처리일때, 
                            if (WinMove.OutSubType.Trim().Equals("3"))
                            {
                                // 정상, 샘플, 불량 한곳에 넣기 - 프로시저에서 Order by - LabelID
                                if (tempLabel.LabelID.Equals(WinMoveSub.LabelID))
                                {

                                    // N:정상, S:Sample, D:Defect
                                    if (WinMoveSub.Gubun == null
                                        || WinMoveSub.Gubun.Trim().Equals("N")
                                        || WinMoveSub.Gubun.Trim().Equals(""))
                                    {
                                        tempLabel.NQty = WinMoveSub.Qty;
                                    }
                                    else if (WinMoveSub.Gubun.Trim().Equals("S"))
                                    {
                                        tempLabel.SQty = WinMoveSub.Qty;
                                    }
                                    else if (WinMoveSub.Gubun.Trim().Equals("D"))
                                    {
                                        tempLabel.DefectID = WinMoveSub.DefectID;
                                        tempLabel.DefectName = WinMoveSub.DefectName;
                                        tempLabel.DQty = WinMoveSub.Qty;
                                    }

                                    if (i == dt.Rows.Count)
                                    {
                                        tempLabel.Qty = stringFormatN0(ConvertDouble(tempLabel.NQty == null ? "" : tempLabel.NQty.ToString())
                                                + ConvertDouble(tempLabel.SQty == null ? "" : tempLabel.SQty.ToString())
                                                + ConvertDouble(tempLabel.DQty == null ? "" : tempLabel.DQty.ToString()));
                                        q++;
                                        tempLabel.Num = q;
                                        dgdPart.Items.Add(tempLabel);
                                    }

                                }
                                else // 라벨이 다르다면, 
                                {
                                    if (dt.Rows.Count != 1)
                                    {
                                        if (i != 1) // 첫번째는 패스
                                        {
                                            tempLabel.Qty = stringFormatN0(ConvertDouble(tempLabel.NQty == null ? "" : tempLabel.NQty.ToString())
                                                + ConvertDouble(tempLabel.SQty == null ? "" : tempLabel.SQty.ToString())
                                                + ConvertDouble(tempLabel.DQty == null ? "" : tempLabel.DQty.ToString()));
                                            q++;
                                            tempLabel.Num = q;
                                            dgdPart.Items.Add(tempLabel);
                                        }
                                    }

                                    tempLabel = WinMoveSub;

                                    // N:정상, S:Sample, D:Defect
                                    if (tempLabel.Gubun == null
                                        || tempLabel.Gubun.Trim().Equals("N")
                                        || tempLabel.Gubun.Trim().Equals(""))
                                    {
                                        tempLabel.NQty = tempLabel.Qty;
                                    }
                                    else if (tempLabel.Gubun.Trim().Equals("S"))
                                    {
                                        tempLabel.SQty = tempLabel.Qty;
                                    }
                                    else if (tempLabel.Gubun.Trim().Equals("D"))
                                    {
                                        tempLabel.DefectID = WinMoveSub.DefectID;
                                        tempLabel.DefectName = WinMoveSub.DefectName;
                                        tempLabel.DQty = tempLabel.Qty;
                                    }

                                    if (dt.Rows.Count == 1 || i == dt.Rows.Count)
                                    {
                                        q++;
                                        tempLabel.Num = q;
                                        dgdPart.Items.Add(tempLabel);
                                    }
                                }

                            }
                            else // 나머진 ID, 수량기준
                            {
                                dgdSub.Items.Add(WinMoveSub);
                            }
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

        #endregion // 서브 그리드 조회 메서드 FillGridSub()

        #region 유효성 검사 CheckData()

        // 유효성 검사
        private bool CheckData()
        {
            bool flag = true;

            // 검색
            if (!strFlag.Equals("I") && !strFlag.Equals("U"))
            {
                // 이동일자 체크 > 날짜 선택 안됬을 때
                if (chkDateSrh.IsChecked == true
                    && (dtpSDateSrh.SelectedDate == null || dtpEDateSrh.SelectedDate == null))
                {
                    MessageBox.Show("이동일자를 선택해주세요.");
                    flag = false;
                    return flag;
                }

                // 품명 체크 > 품명 입력 안됬을 때
                if (chkBuyerArticleNo.IsChecked == true
                    && (txtArticle.Tag == null || txtArticle.Text.Trim().Equals("")))
                {
                    MessageBox.Show("품명을 입력해주세요.");
                    flag = false;
                    return flag;
                }

                // 전 창고 체크 > 선택 안됬을 때
                if (chkFromLocSrh.IsChecked == true
                    && cboFromLocSrh.SelectedValue == null)
                {
                    MessageBox.Show("전 창고를 선택해주세요.");
                    flag = false;
                    return flag;
                }

                // 후 창고 체크 > 선택 안됬을 때
                if (chkToLocSrh.IsChecked == true
                    && cboToLocSrh.SelectedValue == null)
                {
                    MessageBox.Show("후 창고를 선택해주세요.");
                    flag = false;
                    return flag;
                }
            }

            // 추가, 수정
            if (strFlag.Equals("I") || strFlag.Equals("U"))
            {
                // 품명 입력 안했을 때
                if (txtArticle.Tag == null
                    || txtArticle.Tag.ToString().Trim().Equals(""))
                {
                    MessageBox.Show("품명을 입력해주세요.");
                    flag = false;
                    return flag;
                }

                // 전 창고 선택 안했을 때
                //if (cboFromLoc.SelectedValue == null
                //    || cboFromLoc.SelectedValue.ToString().Trim().Equals(""))
                //{
                //    MessageBox.Show("전 창고를 선택해주세요.");
                //    flag = false;
                //    return flag;
                //}

                //// 후 창고 선택 안했을 때
                //if (cboToLoc.SelectedValue == null
                //    || cboToLoc.SelectedValue.ToString().Trim().Equals(""))
                //{
                //    MessageBox.Show("후 창고를 선택해주세요.");
                //    flag = false;
                //    return flag;
                //}

                // 전창고 후창고 다르게 > 이동전 창고와 이후 창고가 동일합니다. \r 서로 다른 창고로 선택해 주세요.
                //if (cboFromLoc.SelectedValue != null && cboToLoc.SelectedValue !=null
                //    && cboFromLoc.SelectedValue.ToString().Trim().Equals(cboToLoc.SelectedValue.ToString().Trim()))
                //{
                //    MessageBox.Show("이동전 창고와 이후 창고가 동일합니다.\r서로 다른 창고로 선택해 주세요.");
                //    flag = false;
                //    return flag;
                //}

                //// 이동구분 선택 안했을 때
                //if (cboOutClss.SelectedValue == null
                //    || cboOutClss.SelectedValue.ToString().Trim().Equals(""))
                //{
                //    MessageBox.Show("이동구분을 선택해주세요.");
                //    flag = false;
                //    return flag;
                //}

                //// 단위 선택 안했을 때
                //if (cboUnitClss.SelectedValue == null
                //    || cboOutClss.SelectedValue.ToString().Trim().Equals(""))
                //{
                //    MessageBox.Show("단위를 선택해주세요.");
                //    flag = false;
                //    return flag;
                //}

                // 수량이 입력되지 않았습니다. 바코드를 통해 데이터를 입력해주세요.
                // 부분처리 일 경우
                if (tgnMovePartial.IsChecked == true)
                {
                    // 데이터 없음
                    if (dgdPart.Items.Count < 1)
                    {
                        MessageBox.Show("수량이 입력되지 않았습니다.\r바코드를 통해 데이터를 입력해주세요.");
                        flag = false;
                        return flag;
                    }

                    // 불량코드 입력 → 불량 수량 입력하지 않았을때
                    // 불량수량 입력 → 불량 코드 입력하지 않았을때
                    // 숫자 유효성 체크
                    for (int i = 0; i < dgdPart.Items.Count; i++)
                    {
                        bool numFlag = true;

                        var label = dgdPart.Items[i] as LabelList3;

                        // 정상수량
                        if (label.NQty == null || CheckConvertInt(label.NQty) == false)
                        {
                            numFlag = false;
                        }
                        // 샘플수량
                        if (label.NQty == null || CheckConvertInt(label.NQty) == false)
                        {
                            numFlag = false;
                        }
                        // 불량수량
                        if (label.NQty == null || CheckConvertInt(label.NQty) == false)
                        {
                            numFlag = false;
                        }

                        if (numFlag == false)
                        {
                            MessageBox.Show("수량은 숫자만 입력이 가능 합니다.");
                            flag = false;

                            dgdPart.SelectedIndex = i;

                            return flag;
                        }
                        else
                        {
                            if ((label.DefectName != null && !label.DefectName.Trim().Equals(""))
                                && label.DefectID != null && !label.DefectID.ToString().Trim().Equals("")
                                && (label.DQty == null || label.DQty.Trim().Equals("")))
                            {
                                MessageBox.Show("불량 수량을 입력해주세요.");
                                flag = false;
                                return flag;
                            }

                            if (label.DQty != null && !label.DQty.ToString().Trim().Equals("")
                                && (label.DefectID == null || label.DefectID.Trim().Equals("")))
                            {
                                MessageBox.Show("불량 코드를 입력해주세요.");
                                flag = false;
                                return flag;
                            }
                        }
                    }
                }
                else // ID 기준, 수량 기준 일 경우
                {
                    // 데이터 없음
                    if (dgdSub.Items.Count < 1)
                    {
                        MessageBox.Show("수량이 입력되지 않았습니다.\r바코드를 통해 데이터를 입력해주세요.");
                        flag = false;
                        return flag;
                    }

                    // 숫자 유효성 체크
                    for (int i = 0; i < dgdSub.Items.Count; i++)
                    {

                        var label = dgdSub.Items[i] as LabelList3;

                        if (label != null)
                        {
                            // 정상수량
                            if (label.Qty == null || CheckConvertInt(label.Qty) == false)
                            {
                                MessageBox.Show("수량은 숫자만 입력이 가능 합니다.");
                                flag = false;

                                dgdSub.SelectedIndex = i;

                                return flag;
                            }
                        }

                    }
                }

            }


            return flag;
        }

        #endregion // 유효성 검사 CheckData()

        #region 저장 메서드 SaveData()

        // 외주 이동 > outware, outwaresub + 동시입고처리!! (stuffin에만!)
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

                    sqlParameter.Add("OrderID", "");
                    sqlParameter.Add("CompanyID", MainWindow.CompanyID);
                    //sqlParameter.Add("OutClss", cboOutClss.SelectedValue != null ? cboOutClss.SelectedValue.ToString() : "");

                    //sqlParameter.Add("CustomID", MainWindow.CompanyID); // 얘는 어떻게?? >  일단 자사껄로
                    //sqlParameter.Add("CustomID", txtCustom.Tag != null ? txtCustom.Tag.ToString() : ""); // 얘는 거래처냐

                    sqlParameter.Add("CustomID", txtCustom.Tag != null && !txtCustom.Tag.ToString().Trim().Equals("") ? txtCustom.Tag.ToString() : "");

                    sqlParameter.Add("BuyerDirectYN", "N"); // 이건 무조건 N
                    sqlParameter.Add("WorkID", "0001");
                    sqlParameter.Add("ExchRate", 0);
                    sqlParameter.Add("UnitPriceClss", "");

                    sqlParameter.Add("InsStuffInYN", "Y");
                    //sqlParameter.Add("OutcustomID", MainWindow.CompanyID);  // 이동의 경우에는 거래처가 없으므로 해당 시스템이 설치된 업체의 코드를 가져옴
                    sqlParameter.Add("OutcustomID", txtCustom.Tag != null ? txtCustom.Tag.ToString() : "");  // 아놔암ㄴ이라ㅓㅁㅇ니라ㅓ 출고처

                    //sqlParameter.Add("Outcustom", txtOutCustom.Text);
                    sqlParameter.Add("LossRate", 0);
                    sqlParameter.Add("LossQty", 0);

                    //sqlParameter.Add("OutRoll", ConvertInt(txtOutRoll.Text));
                    //sqlParameter.Add("OutQty", ConvertDouble(txtOutQty.Text)); 
                    //sqlParameter.Add("OutRealQty", ConvertDouble(txtOutQty.Text));
                    sqlParameter.Add("OutDate", dtpOutDate.SelectedDate != null ? dtpOutDate.SelectedDate.Value.ToString("yyyyMMdd") : "");
                    sqlParameter.Add("ResultDate", dtpOutDate.SelectedDate != null ? dtpOutDate.SelectedDate.Value.ToString("yyyyMMdd") : "");

                    //sqlParameter.Add("Remark", txtRemark.Text);
                    sqlParameter.Add("OutType", 0); // CM_Code 테이블의 OUTTYP - 출고방식 : 0 : 수동 / 1 : 검사기준 자동 / 2: 기타출고 / 3 : PDA출고 > 사무실에서 하는거니까 0!!!!
                    //sqlParameter.Add("OutSubType", tgnMoveByID.IsChecked == false ? (tgnMoveByQty.IsChecked == true ? "2" : "3") : "1"); // 1 : ID기준, 2 : 수량기준, 3 : 부분처리 2021-08-06 수량으로 이동 주석 처리로 인한 수정
                    //sqlParameter.Add("OutSubType", tgnMoveByID.IsChecked == false ?  "3" : "1"); // 1 : ID기준, 2 : 수량기준, 3 : 부분처리

                    sqlParameter.Add("Amount", 0);
                    sqlParameter.Add("VatAmount", 0);
                    sqlParameter.Add("VatINDYN", "");

                    //sqlParameter.Add("FromLocID", cboFromLoc.SelectedValue != null ? cboFromLoc.SelectedValue.ToString() : "");
                    //sqlParameter.Add("ToLocID", cboToLoc.SelectedValue != null ? cboToLoc.SelectedValue.ToString() : "");
                    //sqlParameter.Add("UnitClss", cboUnitClss.SelectedValue != null ? cboUnitClss.SelectedValue.ToString().Trim() : "");
                    sqlParameter.Add("ArticleID", txtArticle.Tag != null ? txtArticle.Tag.ToString() : "");
                    sqlParameter.Add("UserID", MainWindow.CurrentUser);

                    // OutSubType - 1 : ID기준 / 2 : 수량 기준 / 3 : 부분처리 => 추가하기

                    if (strFlag.Equals("I"))
                    {
                        sqlParameter.Add("OutSeq", 0); // output > ioutware 프로시저에서 새로 설정됨. 
                        sqlParameter.Add("OutwareNo", ""); // output > OutwareID 임


                        Dictionary<string, int> outputParam = new Dictionary<string, int>();
                        outputParam.Add("OutwareNo", 12);
                        outputParam.Add("OutSeq", 10);

                        Dictionary<string, string> dicResult = DataStore.Instance.ExecuteProcedureOutputNoTran("xp_Outware_iOutware", sqlParameter, outputParam, true);
                        string result = dicResult["OutwareNo"];
                        string resultSeq = dicResult["OutSeq"];

                        if ((result != string.Empty) || (result != "9999"))
                        {

                            // outwareSub 에 데이터를 넣어줘야 하는데..
                            // 라벨 관리 일경우에는 dgdSub 를 가져오지만.
                            // 수량관리, 부분처리는 별개로 처리

                            // 1. ID기준이동, 2. 수량기준
                            //if (tgnMoveByID.IsChecked == true || tgnMoveByQty.IsChecked == true)  2021-08-06 수량으로 이동 주석 처리로 인한 수정
                            //if (tgnMoveByID.IsChecked == true)
                            //{
                            //    for (int i = 0; i < dgdSub.Items.Count; i++)
                            //    {
                            //        var WinMoveSub = dgdSub.Items[i] as LabelList3;

                            //        // OutwareSub 에 등록
                            //        sqlParameter = new Dictionary<string, object>();
                            //        sqlParameter.Clear();

                            //        sqlParameter.Add("OutwareID", result);
                            //        sqlParameter.Add("OrderID", "");
                            //        sqlParameter.Add("OutSeq", ConvertInt(resultSeq));
                            //        sqlParameter.Add("OutSubSeq", i + 1);
                            //        sqlParameter.Add("OrderSeq", "");

                            //        sqlParameter.Add("LineSeq", 0);
                            //        sqlParameter.Add("LineSubSeq", 0);
                            //        sqlParameter.Add("RollSeq", 0);
                            //        sqlParameter.Add("LabelID", WinMoveSub.LabelID == null ? "" : WinMoveSub.LabelID);
                            //        sqlParameter.Add("LabelGubun", WinMoveSub.LabelGubun == null ? "" : WinMoveSub.LabelGubun); // 2 : BoxID / 3: LotID

                            //        sqlParameter.Add("LotNo", ""); // 얘는 도대체 뭐여
                            //        sqlParameter.Add("Gubun", "N"); // N : 정상 / S : 샘플 / D : Defect(결함, 불량) > 부분처리 일경우에는 !!!!! 적용 되는 것들
                            //        sqlParameter.Add("StuffQty", 0);
                            //        //sqlParameter.Add("OutQty", ConvertDouble(WinMoveSub.Qty));
                            //        sqlParameter.Add("OutQty", ConvertDouble(WinMoveSub.Qty));
                            //        sqlParameter.Add("OutRoll", 1); // 박스 갯수 - 라벨 하나당 박스 1개로 처리 하니, 1로 저장

                            //        sqlParameter.Add("UnitPrice", 0);
                            //        sqlParameter.Add("UserID", MainWindow.CurrentUser);
                            //        sqlParameter.Add("CustomBoxID", "");
                            //        sqlParameter.Add("BoxID", "");


                            //        Procedure pro2 = new Procedure();
                            //        pro2.Name = "xp_Outware_iOutwareSub";
                            //        pro2.OutputUseYN = "N";
                            //        pro2.OutputName = "REQ_ID";
                            //        pro2.OutputLength = "10";

                            //        Prolist.Add(pro2);
                            //        ListParameter.Add(sqlParameter);
                            //    }
                            //}
                            //else // 부분처리
                            //{
                            int q = 0;
                            for (int i = 0; i < dgdPart.Items.Count; i++)
                            {
                                var WinMoveSub = dgdPart.Items[i] as LabelList3;

                                if (WinMoveSub != null)
                                {


                                    if (WinMoveSub.NQty != null && !WinMoveSub.NQty.Trim().Equals("")) // 정상 제품이 존재한다면 = 정상 개수가 존재한다면
                                    {
                                        sqlParameter = new Dictionary<string, object>();
                                        sqlParameter.Clear();

                                        q++;
                                        sqlParameter.Add("OutQty", ConvertInt(WinMoveSub.NQty));
                                        sqlParameter.Add("OutSubSeq", q);
                                        sqlParameter.Add("Gubun", "N");

                                        sqlParameter.Add("OutwareID", result);
                                        sqlParameter.Add("OrderID", "");
                                        sqlParameter.Add("OutSeq", ConvertInt(resultSeq));
                                        //sqlParameter.Add("OutSubSeq", i + 1);
                                        sqlParameter.Add("OrderSeq", "");

                                        sqlParameter.Add("LineSeq", 0);
                                        sqlParameter.Add("LineSubSeq", 0);
                                        sqlParameter.Add("RollSeq", 0);
                                        sqlParameter.Add("LabelID", WinMoveSub.LabelID == null ? "" : WinMoveSub.LabelID);
                                        sqlParameter.Add("LabelGubun", ""); // 2 : BoxID / 3: LotID

                                        sqlParameter.Add("LotNo", ""); // 얘는 도대체 뭐여
                                                                       //sqlParameter.Add("Gubun", "N"); // N : 정상 / S : 샘플 / D : Defect(결함, 불량) > 부분처리 일경우에는 !!!!! 적용 되는 것들
                                        sqlParameter.Add("StuffQty", 0);
                                        sqlParameter.Add("OutQty", ConvertDouble(WinMoveSub.Qty));
                                        sqlParameter.Add("OutRoll", 1); // 박스 갯수 - 라벨 하나당 박스 1개로 처리 하니, 1로 저장

                                        sqlParameter.Add("UnitPrice", 0);
                                        sqlParameter.Add("UserID", MainWindow.CurrentUser);
                                        sqlParameter.Add("CustomBoxID", "");
                                        sqlParameter.Add("BoxID", "");

                                        Procedure pro2 = new Procedure();
                                        pro2.Name = "xp_Outware_iOutwareSub";
                                        pro2.OutputUseYN = "N";
                                        pro2.OutputName = "REQ_ID";
                                        pro2.OutputLength = "10";

                                        Prolist.Add(pro2);
                                        ListParameter.Add(sqlParameter);
                                    }

                                    if (WinMoveSub.SQty != null && !WinMoveSub.SQty.Trim().Equals("")) // 샘플 제품이 존재한다면 = 샘플 개수가 존재한다면
                                    {
                                        //// 기존에 들어가 있는 OutQty 제거
                                        //if (sqlParameter.ContainsKey("OutQty") == true)
                                        //    sqlParameter.Remove("OutQty");
                                        //if (sqlParameter.ContainsKey("OutSubSeq") == true)
                                        //    sqlParameter.Remove("OutSubSeq");
                                        //if (sqlParameter.ContainsKey("Gubun") == true)
                                        //    sqlParameter.Remove("Gubun");

                                        sqlParameter = new Dictionary<string, object>();
                                        sqlParameter.Clear();

                                        q++;
                                        sqlParameter.Add("OutQty", ConvertInt(WinMoveSub.SQty));
                                        sqlParameter.Add("OutSubSeq", q);
                                        sqlParameter.Add("Gubun", "S");

                                        sqlParameter.Add("OutwareID", result);
                                        sqlParameter.Add("OrderID", "");
                                        sqlParameter.Add("OutSeq", ConvertInt(resultSeq));
                                        //sqlParameter.Add("OutSubSeq", i + 1);
                                        sqlParameter.Add("OrderSeq", "");

                                        sqlParameter.Add("LineSeq", 0);
                                        sqlParameter.Add("LineSubSeq", 0);
                                        sqlParameter.Add("RollSeq", 0);
                                        sqlParameter.Add("LabelID", WinMoveSub.LabelID == null ? "" : WinMoveSub.LabelID);
                                        sqlParameter.Add("LabelGubun", ""); // 2 : BoxID / 3: LotID

                                        sqlParameter.Add("LotNo", ""); // 얘는 도대체 뭐여
                                                                       //sqlParameter.Add("Gubun", "N"); // N : 정상 / S : 샘플 / D : Defect(결함, 불량) > 부분처리 일경우에는 !!!!! 적용 되는 것들
                                        sqlParameter.Add("StuffQty", 0);
                                        sqlParameter.Add("OutQty", ConvertDouble(WinMoveSub.Qty));
                                        sqlParameter.Add("OutRoll", 1); // 박스 갯수 - 라벨 하나당 박스 1개로 처리 하니, 1로 저장

                                        sqlParameter.Add("UnitPrice", 0);
                                        sqlParameter.Add("UserID", MainWindow.CurrentUser);
                                        sqlParameter.Add("CustomBoxID", "");
                                        sqlParameter.Add("BoxID", "");

                                        Procedure pro2 = new Procedure();
                                        pro2.Name = "xp_Outware_iOutwareSub";
                                        pro2.OutputUseYN = "N";
                                        pro2.OutputName = "REQ_ID";
                                        pro2.OutputLength = "10";

                                        Prolist.Add(pro2);
                                        ListParameter.Add(sqlParameter);
                                    }


                                    if (WinMoveSub.DefectName != null
                                        && !WinMoveSub.DefectName.Trim().Equals("")
                                        && WinMoveSub.DefectID != null
                                        && !WinMoveSub.DefectID.Trim().Equals("")) // 불량 제품이 존재한다면 = 불량코드가 존재한다면
                                    {
                                        //// 기존에 들어가 있는 OutQty 제거
                                        //if (sqlParameter.ContainsKey("OutQty") == true)
                                        //    sqlParameter.Remove("OutQty");
                                        //if (sqlParameter.ContainsKey("OutSubSeq") == true)
                                        //    sqlParameter.Remove("OutSubSeq");
                                        //if (sqlParameter.ContainsKey("Gubun") == true)
                                        //    sqlParameter.Remove("Gubun");

                                        sqlParameter = new Dictionary<string, object>();
                                        sqlParameter.Clear();

                                        q++;
                                        sqlParameter.Add("DefectID", WinMoveSub.DefectID);
                                        sqlParameter.Add("OutQty", ConvertInt(WinMoveSub.DQty));
                                        sqlParameter.Add("OutSubSeq", q);
                                        sqlParameter.Add("Gubun", "D");

                                        sqlParameter.Add("OutwareID", result);
                                        sqlParameter.Add("OrderID", "");
                                        sqlParameter.Add("OutSeq", ConvertInt(resultSeq));
                                        //sqlParameter.Add("OutSubSeq", i + 1);
                                        sqlParameter.Add("OrderSeq", "");

                                        sqlParameter.Add("LineSeq", 0);
                                        sqlParameter.Add("LineSubSeq", 0);
                                        sqlParameter.Add("RollSeq", 0);
                                        sqlParameter.Add("LabelID", WinMoveSub.LabelID == null ? "" : WinMoveSub.LabelID);
                                        sqlParameter.Add("LabelGubun", ""); // 2 : BoxID / 3: LotID

                                        sqlParameter.Add("LotNo", ""); // 얘는 도대체 뭐여
                                                                       //sqlParameter.Add("Gubun", "N"); // N : 정상 / S : 샘플 / D : Defect(결함, 불량) > 부분처리 일경우에는 !!!!! 적용 되는 것들
                                        sqlParameter.Add("StuffQty", 0);
                                        sqlParameter.Add("OutQty", ConvertDouble(WinMoveSub.Qty));
                                        sqlParameter.Add("OutRoll", 1); // 박스 갯수 - 라벨 하나당 박스 1개로 처리 하니, 1로 저장

                                        sqlParameter.Add("UnitPrice", 0);
                                        sqlParameter.Add("UserID", MainWindow.CurrentUser);
                                        sqlParameter.Add("CustomBoxID", "");
                                        sqlParameter.Add("BoxID", "");

                                        Procedure pro2 = new Procedure();
                                        pro2.Name = "xp_Outware_iOutwareSub";
                                        pro2.OutputUseYN = "N";
                                        pro2.OutputName = "REQ_ID";
                                        pro2.OutputLength = "10";

                                        Prolist.Add(pro2);
                                        ListParameter.Add(sqlParameter);
                                    }
                                }
                                // }
                            }


                            // 입고처리 하기 > 수량기준, 부분처리를 했을 때, StuffinSub에 라벨이 생성 되어야 하는가 > 라벨은 모르겠고, StuffinSub에도 들어감
                            // [xp_StuffIN_iStuffINByOutware ]
                            sqlParameter = new Dictionary<string, object>();
                            sqlParameter.Clear();

                            sqlParameter.Add("OutwareID", result);
                            sqlParameter.Add("sUserID", MainWindow.CurrentUser);
                            sqlParameter.Add("sOutmsg", "");

                            Procedure pro3 = new Procedure();
                            pro3.Name = "xp_StuffIN_iStuffINByOutware";
                            pro3.OutputUseYN = "N";
                            pro3.OutputName = "REQ_ID";
                            pro3.OutputLength = "10";

                            Prolist.Add(pro3);
                            ListParameter.Add(sqlParameter);
                        }
                    }
                    #region 수정
                    else // 수정시
                    {
                        // 1. outware 는 [xp_Outware_uOutware] : outware 수정 후 outwaresub, stuffin 도 같이 지우는 프로시저 
                        // 2. outwaresub 다시 등록
                        // 3. stuffin 다시 등록

                        var WinMove = dgdMain.SelectedItem as Win_mtr_Move_Prd_CodeView2;
                        string OutwareID = WinMove.OutwareID;
                        string OutSeq = WinMove.OutSeq;

                        // 1. 
                        sqlParameter.Add("OutwareID", OutwareID);

                        Procedure pro1 = new Procedure();
                        pro1.Name = "xp_Outware_uOutware_MTR";
                        pro1.OutputUseYN = "N";
                        pro1.OutputName = "REQ_ID";
                        pro1.OutputLength = "10";

                        Prolist.Add(pro1);
                        ListParameter.Add(sqlParameter);

                        // 2. 
                        // 1. ID기준이동, 2. 수량기준
                        //if (tgnMoveByID.IsChecked == true || tgnMoveByQty.IsChecked == true) 2021-08-06 수량으로 이동 주석 처리로 인한 수정
                        //if (tgnMoveByID.IsChecked == true)
                        //{
                        //    for (int i = 0; i < dgdSub.Items.Count; i++)
                        //    {
                        //        var WinMoveSub = dgdSub.Items[i] as LabelList3;

                        //        // OutwareSub 에 등록
                        //        sqlParameter = new Dictionary<string, object>();
                        //        sqlParameter.Clear();

                        //        sqlParameter.Add("OutwareID", OutwareID);
                        //        sqlParameter.Add("OrderID", "");
                        //        sqlParameter.Add("OutSeq", ConvertInt(OutSeq));
                        //        sqlParameter.Add("OutSubSeq", i + 1);
                        //        sqlParameter.Add("OrderSeq", "");

                        //        sqlParameter.Add("LineSeq", 0);
                        //        sqlParameter.Add("LineSubSeq", 0);
                        //        sqlParameter.Add("RollSeq", 0);
                        //        sqlParameter.Add("LabelID", WinMoveSub.LabelID == null ? "" : WinMoveSub.LabelID);
                        //        sqlParameter.Add("LabelGubun", WinMoveSub.LabelGubun == null ? "" : WinMoveSub.LabelGubun); // 2 : BoxID / 3: LotID

                        //        sqlParameter.Add("LotNo", ""); // 얘는 도대체 뭐여
                        //        sqlParameter.Add("Gubun", "N"); // N : 정상 / S : 샘플 / D : Defect(결함, 불량) > 부분처리 일경우에는 !!!!! 적용 되는 것들
                        //        sqlParameter.Add("StuffQty", 0);
                        //        sqlParameter.Add("OutQty", ConvertDouble(WinMoveSub.Qty));
                        //        sqlParameter.Add("OutRoll", 1); // 박스 갯수 - 라벨 하나당 박스 1개로 처리 하니, 1로 저장

                        //        sqlParameter.Add("UnitPrice", 0);
                        //        sqlParameter.Add("UserID", MainWindow.CurrentUser);
                        //        sqlParameter.Add("CustomBoxID", "");
                        //        sqlParameter.Add("BoxID", "");


                        //        Procedure pro2 = new Procedure();
                        //        pro2.Name = "xp_Outware_iOutwareSub";
                        //        pro2.OutputUseYN = "N";
                        //        pro2.OutputName = "REQ_ID";
                        //        pro2.OutputLength = "10";

                        //        Prolist.Add(pro2);
                        //        ListParameter.Add(sqlParameter);
                        //    }
                        //}
                        //else // 부분처리
                        //{
                        int q = 0;
                        for (int i = 0; i < dgdPart.Items.Count; i++)
                        {
                            var WinMoveSub = dgdPart.Items[i] as LabelList3;

                            if (WinMoveSub != null)
                            {


                                if (WinMoveSub.NQty != null && !WinMoveSub.NQty.Trim().Equals("")) // 정상 제품이 존재한다면 = 정상 개수가 존재한다면
                                {
                                    sqlParameter = new Dictionary<string, object>();
                                    sqlParameter.Clear();

                                    q++;
                                    sqlParameter.Add("OutQty", ConvertInt(WinMoveSub.NQty));
                                    sqlParameter.Add("OutSubSeq", q);
                                    sqlParameter.Add("Gubun", "N");

                                    sqlParameter.Add("OutwareID", OutwareID);
                                    sqlParameter.Add("OrderID", "");
                                    sqlParameter.Add("OutSeq", ConvertInt(OutSeq));
                                    //sqlParameter.Add("OutSubSeq", i + 1);
                                    sqlParameter.Add("OrderSeq", "");

                                    sqlParameter.Add("LineSeq", 0);
                                    sqlParameter.Add("LineSubSeq", 0);
                                    sqlParameter.Add("RollSeq", 0);
                                    sqlParameter.Add("LabelID", WinMoveSub.LabelID == null ? "" : WinMoveSub.LabelID);
                                    sqlParameter.Add("LabelGubun", ""); // 2 : BoxID / 3: LotID

                                    sqlParameter.Add("LotNo", ""); // 얘는 도대체 뭐여
                                                                   //sqlParameter.Add("Gubun", "N"); // N : 정상 / S : 샘플 / D : Defect(결함, 불량) > 부분처리 일경우에는 !!!!! 적용 되는 것들
                                    sqlParameter.Add("StuffQty", 0);
                                    sqlParameter.Add("OutQty", ConvertDouble(WinMoveSub.Qty));
                                    sqlParameter.Add("OutRoll", 1); // 박스 갯수 - 라벨 하나당 박스 1개로 처리 하니, 1로 저장

                                    sqlParameter.Add("UnitPrice", 0);
                                    sqlParameter.Add("UserID", MainWindow.CurrentUser);
                                    sqlParameter.Add("CustomBoxID", "");
                                    sqlParameter.Add("BoxID", "");

                                    Procedure pro2 = new Procedure();
                                    pro2.Name = "xp_Outware_iOutwareSub";
                                    pro2.OutputUseYN = "N";
                                    pro2.OutputName = "REQ_ID";
                                    pro2.OutputLength = "10";

                                    Prolist.Add(pro2);
                                    ListParameter.Add(sqlParameter);
                                }

                                if (WinMoveSub.SQty != null && !WinMoveSub.SQty.Trim().Equals("")) // 샘플 제품이 존재한다면 = 샘플 개수가 존재한다면
                                {
                                    //// 기존에 들어가 있는 OutQty 제거
                                    //if (sqlParameter.ContainsKey("OutQty") == true)
                                    //    sqlParameter.Remove("OutQty");
                                    //if (sqlParameter.ContainsKey("OutSubSeq") == true)
                                    //    sqlParameter.Remove("OutSubSeq");
                                    //if (sqlParameter.ContainsKey("Gubun") == true)
                                    //    sqlParameter.Remove("Gubun");

                                    sqlParameter = new Dictionary<string, object>();
                                    sqlParameter.Clear();

                                    q++;
                                    sqlParameter.Add("OutQty", ConvertInt(WinMoveSub.SQty));
                                    sqlParameter.Add("OutSubSeq", q);
                                    sqlParameter.Add("Gubun", "S");

                                    sqlParameter.Add("OutwareID", OutwareID);
                                    sqlParameter.Add("OrderID", "");
                                    sqlParameter.Add("OutSeq", ConvertInt(OutSeq));
                                    //sqlParameter.Add("OutSubSeq", i + 1);
                                    sqlParameter.Add("OrderSeq", "");

                                    sqlParameter.Add("LineSeq", 0);
                                    sqlParameter.Add("LineSubSeq", 0);
                                    sqlParameter.Add("RollSeq", 0);
                                    sqlParameter.Add("LabelID", WinMoveSub.LabelID == null ? "" : WinMoveSub.LabelID);
                                    sqlParameter.Add("LabelGubun", ""); // 2 : BoxID / 3: LotID

                                    sqlParameter.Add("LotNo", ""); // 얘는 도대체 뭐여
                                                                   //sqlParameter.Add("Gubun", "N"); // N : 정상 / S : 샘플 / D : Defect(결함, 불량) > 부분처리 일경우에는 !!!!! 적용 되는 것들
                                    sqlParameter.Add("StuffQty", 0);
                                    sqlParameter.Add("OutQty", ConvertDouble(WinMoveSub.Qty));
                                    sqlParameter.Add("OutRoll", 1); // 박스 갯수 - 라벨 하나당 박스 1개로 처리 하니, 1로 저장

                                    sqlParameter.Add("UnitPrice", 0);
                                    sqlParameter.Add("UserID", MainWindow.CurrentUser);
                                    sqlParameter.Add("CustomBoxID", "");
                                    sqlParameter.Add("BoxID", "");

                                    Procedure pro2 = new Procedure();
                                    pro2.Name = "xp_Outware_iOutwareSub";
                                    pro2.OutputUseYN = "N";
                                    pro2.OutputName = "REQ_ID";
                                    pro2.OutputLength = "10";

                                    Prolist.Add(pro2);
                                    ListParameter.Add(sqlParameter);
                                }


                                if (WinMoveSub.DefectName != null
                                        && !WinMoveSub.DefectName.Trim().Equals("")
                                        && WinMoveSub.DefectID != null
                                        && !WinMoveSub.DefectID.Trim().Equals("")) // 불량 제품이 존재한다면 = 불량코드가 존재한다면
                                {
                                    //// 기존에 들어가 있는 OutQty 제거
                                    //if (sqlParameter.ContainsKey("OutQty") == true)
                                    //    sqlParameter.Remove("OutQty");
                                    //if (sqlParameter.ContainsKey("OutSubSeq") == true)
                                    //    sqlParameter.Remove("OutSubSeq");
                                    //if (sqlParameter.ContainsKey("Gubun") == true)
                                    //    sqlParameter.Remove("Gubun");

                                    sqlParameter = new Dictionary<string, object>();
                                    sqlParameter.Clear();

                                    q++;
                                    sqlParameter.Add("DefectID", WinMoveSub.DefectID);
                                    sqlParameter.Add("OutQty", ConvertInt(WinMoveSub.DQty));
                                    sqlParameter.Add("OutSubSeq", q);
                                    sqlParameter.Add("Gubun", "D");

                                    sqlParameter.Add("OutwareID", OutwareID);
                                    sqlParameter.Add("OrderID", "");
                                    sqlParameter.Add("OutSeq", ConvertInt(OutSeq));
                                    //sqlParameter.Add("OutSubSeq", i + 1);
                                    sqlParameter.Add("OrderSeq", "");

                                    sqlParameter.Add("LineSeq", 0);
                                    sqlParameter.Add("LineSubSeq", 0);
                                    sqlParameter.Add("RollSeq", 0);
                                    sqlParameter.Add("LabelID", WinMoveSub.LabelID == null ? "" : WinMoveSub.LabelID);
                                    sqlParameter.Add("LabelGubun", ""); // 2 : BoxID / 3: LotID

                                    sqlParameter.Add("LotNo", ""); // 얘는 도대체 뭐여
                                                                   //sqlParameter.Add("Gubun", "N"); // N : 정상 / S : 샘플 / D : Defect(결함, 불량) > 부분처리 일경우에는 !!!!! 적용 되는 것들
                                    sqlParameter.Add("StuffQty", 0);
                                    sqlParameter.Add("OutQty", ConvertDouble(WinMoveSub.Qty));
                                    sqlParameter.Add("OutRoll", 1); // 박스 갯수 - 라벨 하나당 박스 1개로 처리 하니, 1로 저장

                                    sqlParameter.Add("UnitPrice", 0);
                                    sqlParameter.Add("UserID", MainWindow.CurrentUser);
                                    sqlParameter.Add("CustomBoxID", "");
                                    sqlParameter.Add("BoxID", "");

                                    Procedure pro2 = new Procedure();
                                    pro2.Name = "xp_Outware_iOutwareSub";
                                    pro2.OutputUseYN = "N";
                                    pro2.OutputName = "REQ_ID";
                                    pro2.OutputLength = "10";

                                    Prolist.Add(pro2);
                                    ListParameter.Add(sqlParameter);
                                }
                            }
                            //}
                        }

                        // 3. 
                        // 입고처리 하기 > 수량기준, 부분처리를 했을 때, StuffinSub에 라벨이 생성 되어야 하는가 > 라벨은 모르겠고, StuffinSub에도 들어감
                        // [xp_StuffIN_iStuffINByOutware ]
                        sqlParameter = new Dictionary<string, object>();
                        sqlParameter.Clear();

                        sqlParameter.Add("OutwareID", OutwareID);
                        sqlParameter.Add("sUserID", MainWindow.CurrentUser);
                        sqlParameter.Add("sOutmsg", "");

                        Procedure pro3 = new Procedure();
                        pro3.Name = "xp_StuffIN_iStuffINByOutware";
                        pro3.OutputUseYN = "N";
                        pro3.OutputName = "REQ_ID";
                        pro3.OutputLength = "10";

                        Prolist.Add(pro3);
                        ListParameter.Add(sqlParameter);
                    }
                    #endregion // 수정

                    string[] Confirm = new string[2];
                    Confirm = DataStore.Instance.ExecuteAllProcedureOutputNew(Prolist, ListParameter);
                    if (Confirm[0] != "success")
                    {
                        MessageBox.Show("[저장실패]\r\n" + Confirm[1].ToString());
                        flag = false;
                        //return false;
                    }
                    else
                    {
                        //MessageBox.Show("성공");
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

            return flag;
        }

        #endregion // 저장 메서드 SaveData()

        #region 삭제 메서드 Delete()

        private bool DeleteData(string OutwareID)
        {
            bool flag = false;

            try
            {
                Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                sqlParameter.Clear();
                sqlParameter.Add("OutwareID", OutwareID);

                string[] result = DataStore.Instance.ExecuteProcedure("xp_Outware_dOutware", sqlParameter, false);

                if (result[0].Equals("success"))
                {
                    //MessageBox.Show("성공 *^^*");
                    flag = true;
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

            return flag;
        }

        #endregion // 삭제 메서드 Delete()

        #endregion // 주요 메서드 모음

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


        #endregion // 기타 메서드

        private void btnTest_Click(object sender, RoutedEventArgs e)
        {
            var WinMove = dgdMain.SelectedItem as Win_mtr_Move_Prd_CodeView2;

            MessageBox.Show(WinMove.ToString());

            //MessageBox.Show(cboFromLoc.SelectedValue.ToString());
        }

        private void txtCustom_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                MainWindow.pf.ReturnCode(txtCustom, (int)Defind_CodeFind.DCF_CUSTOM, "");
            }
        }

        private void btnCustom_Click(object sender, RoutedEventArgs e)
        {
            MainWindow.pf.ReturnCode(txtCustom, (int)Defind_CodeFind.DCF_CUSTOM, "");
        }


        private void btnPfBuyerArticleNoSrh_Click(object sender, RoutedEventArgs e)
        {
            MainWindow.pf.ReturnCode(txtBuyerArticleNo, 76, "");
        }

        private void DataGridCell_PreviewMouseDown(object sender, MouseButtonEventArgs e)
        {

        }


        private void DataGird_KeyDown02(object sender, KeyEventArgs e)
        {
            int currRow = dgdSub.Items.IndexOf(dgdSub.CurrentItem);
            int currCol = dgdSub.Columns.IndexOf(dgdSub.CurrentCell.Column);
            int startCol = 3;
            int endCol = 3;

            if (e.Key == Key.Enter)
            {
                e.Handled = true;
                (sender as DataGridCell).IsEditing = false;

                // 마지막 열, 마지막 행 아님
                if (endCol == currCol && dgdSub.Items.Count - 1 > currRow)
                {
                    dgdSub.SelectedIndex = currRow + 1; // 이건 한줄 파란색으로 활성화 된 걸 조정하는 것입니다.
                    dgdSub.CurrentCell = new DataGridCellInfo(dgdSub.Items[currRow + 1], dgdSub.Columns[startCol]);
                } // 마지막 열 아님
                else if (endCol > currCol && dgdSub.Items.Count - 1 >= currRow)
                {
                    //if (currCol == 1) currCol++; // 2는 건너뛰기
                    dgdSub.CurrentCell = new DataGridCellInfo(dgdSub.Items[currRow], dgdSub.Columns[currCol + 1]);
                } // 마지막 열, 마지막 행
                else if (endCol == currCol && dgdSub.Items.Count - 1 == currRow)
                {

                }
                else
                {
                    //MessageBox.Show("나머지가 있나..");
                }
            }


        }

        private void TxtQty_TextChanged(object sender, TextChangedEventArgs e)
        {
            SetOutRollAndOutQty();
        }
    }

    #region 코드뷰 CodeView

    class Win_mtr_Move_Prd_CodeView2 : BaseView
    {
        public int Num { get; set; }

        public string OutwareID { get; set; }
        public string OutSeq { get; set; }

        public string OrderID { get; set; }
        public string OrderSeq { get; set; }
        public string OrderNo { get; set; }
        public string CustomID { get; set; }

        public string KCustom { get; set; }
        public string OutDate { get; set; }
        public string ArticleID { get; set; }
        public string Article { get; set; }
        public string OutClss { get; set; }

        public string OutDate_CV { get; set; }

        public string WorkID { get; set; }
        public string OutRoll { get; set; }
        public string OutQty { get; set; }
        public string OutRealQty { get; set; }
        public string ResultDate { get; set; }

        public string ResultDate_CV { get; set; }

        public string OrderQty { get; set; }
        public string UnitClss { get; set; }
        public string WorkName { get; set; }
        public string OutType { get; set; }
        public string Remark { get; set; }

        public string BuyerModel { get; set; }
        public string OutSumQty { get; set; }
        public string OutQtyY { get; set; }
        public string StuffInQty { get; set; }
        public string OutWeight { get; set; }

        public string OutRealWeight { get; set; }
        public string UnitPriceClss { get; set; }
        public string BuyerDirectYN { get; set; }
        public string Vat_Ind_YN { get; set; }
        public string InsStuffINYN { get; set; }

        public string ExchRate { get; set; }
        public string FromLocID { get; set; }
        public string ToLocID { get; set; }
        public string UnitClssName { get; set; }
        public string FromLocName { get; set; }

        public string TOLocname { get; set; }
        public string OutClssname { get; set; }
        public string UnitPrice { get; set; }
        public string Amount { get; set; }
        public string VatAmount { get; set; }

        public string BuyerArticleNo { get; set; }
        public string OutCustomID { get; set; }
        public string BuyerID { get; set; }
        public string BuyerName { get; set; }
        public string OutCustom { get; set; }

        public string CustomName { get; set; }
        public string OutSubType { get; set; }
    }

    class Win_mtr_Move_Prd_CodeView2Sub : BaseView
    {
        public int Num { get; set; }

        public string OutwareID { get; set; }
        public string OutSubSeq { get; set; }
        public string LabelID { get; set; }
        public string LabelGubun { get; set; } // 2 : 박스ID, 3 : LotID
        public string LabelGubunName { get; set; }

        public string OutQty { get; set; }
        public string OutCnt { get; set; }
        public string OutRoll { get; set; }
        public string LotNo { get; set; }
        public string Weight { get; set; }

        public string UnitPrice { get; set; }
        public string Vat_IND_YN { get; set; }
        public string Orderseq { get; set; }
        public string Amount { get; set; }
        public string CustomBoxID { get; set; }

        public string FromLocID { get; set; }
        public string ToLocID { get; set; }
        public string UnitClss { get; set; }
        public string ArticleID { get; set; }
        public string Article { get; set; }

        public string OutClss { get; set; }
    }

    class LabelList3 : BaseView
    {
        public int Num { get; set; }

        public string LabelID { get; set; }
        public string LabelGubun { get; set; } // 2 - ?, 3 - ?, 7 - 공정이동ID : wk_LabelPrint (라벨 발행 테이블) 에서 가져오는듯??
        public string ArticleID { get; set; }
        public string Qty { get; set; }
        public string LabelGubunName { get; set; }

        public string InspectApprovalYN { get; set; }
        public string Inspector { get; set; }
        public string Article { get; set; }
        public string ProcessID { get; set; }
        public string CustomID { get; set; }

        public string Custom { get; set; }
        public string UnitClss { get; set; }
        public string OutClss { get; set; }



        public string OutQty { get; set; }

        public string OutSideQTY { get; set; }


        // 부분처리용
        public string Gubun { get; set; } // N:정상, S:Sample, D:Defect
        public string NQty { get; set; } // 정상 수량
        public string SQty { get; set; } // 샘플 수량

        public string DefectID { get; set; } // 불량코드
        public string DefectName { get; set; }
        public string DQty { get; set; } // 불량 수량

    }

    #endregion 코드뷰 CodeView
}

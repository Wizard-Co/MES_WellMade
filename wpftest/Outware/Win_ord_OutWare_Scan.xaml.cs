using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Data;
using System.Linq;
using System.Threading;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using WizMes_WellMade.PopUP;
using WizMes_WellMade.PopUp;
using System.Text.RegularExpressions;
using System.Windows.Data;
using System.Windows.Media;
using System.Windows.Controls.Primitives;

/**************************************************************************************************
'** 프로그램명 : Win_ord_OutWare_Scan
'** 설명       : 출고지시(스캔)
'** 작성일자   : 
'** 작성자     : 장시영
'**------------------------------------------------------------------------------------------------
'**************************************************************************************************
' 변경일자  , 변경자, 요청자    , 요구사항ID      , 요청 및 작업내용
'**************************************************************************************************
' 2023.03.30, 장시영, 삼익SDT에서 가져옴
' 2023.04.28, 장시영, 라벨 스캔 시 선출고 체크
'**************************************************************************************************/

namespace WizMes_WellMade
{
    /// <summary>
    /// Win_ord_OutWare_Scan.xaml에 대한 상호 작용 논리
    /// </summary>
    public partial class Win_ord_OutWare_Scan : UserControl
    {
        string stDate = string.Empty;
        string stTime = string.Empty;

        Lib lib = new Lib();
        PlusFinder pf = new PlusFinder();

        // 인쇄 활용 용도 (프린트)
        private Microsoft.Office.Interop.Excel.Application excelapp;
        private Microsoft.Office.Interop.Excel.Workbook workbook;
        private Microsoft.Office.Interop.Excel.Worksheet worksheet;
        private Microsoft.Office.Interop.Excel.Range workrange;
        private Microsoft.Office.Interop.Excel.Worksheet pastesheet;
        PopUp.NoticeMessage msg = new PopUp.NoticeMessage();

        Dictionary<string, object> lstCheck = new Dictionary<string, object>();

        List<Win_ord_OutWare_Scan_CodeView> lstOutwarePrint = new List<Win_ord_OutWare_Scan_CodeView>();

        // 수정 정보를 보관하기 위한 변수
        List<Win_ord_OutWare_Scan_Sub_CodeView> lstBoxID = new List<Win_ord_OutWare_Scan_Sub_CodeView>();
        List<Win_ord_OutWare_Scan_Sub_CodeView> ListOutwareSub = new List<Win_ord_OutWare_Scan_Sub_CodeView>();

        int rowNum = 0;                          // 조회시 데이터 줄 번호 저장용도
        string strFlag = string.Empty;           // 추가, 수정 구분 
        string GetKey = "";

        string orderSeq = "";
        int ReqQty = 0;
        int OrderQty = 0;
        int RestOrderQty = 0;
        int tmpRestQty = 0;
        int cnt = 0;

        Dictionary<string, object> LabelGroupInfo = new Dictionary<string, object>();
        List<string> LabelGroupList = new List<string>();   // packing ID 스캔에 따른 LabelID를 모아 담을 리스트 그릇입니다.       
        bool EventStatus = false;                           // 추가 / 수정 상태확인을 위한 이벤트 bool
        bool preview_click = false;                         // 인쇄 미리보기 인지 아닌지

        public Win_ord_OutWare_Scan()
        {
            InitializeComponent();
        }

        private void UserControl_Loaded(object sender, RoutedEventArgs e)
        {
            try
            {
                stDate = DateTime.Now.ToString("yyyyMMdd");
                stTime = DateTime.Now.ToString("HHmm");

                DataStore.Instance.InsertLogByFormS(this.GetType().Name, stDate, stTime, "S");

                chkOutwareDay.IsChecked = true; //출고일자 IsCheked
                dtpFromDate.Text = DateTime.Now.ToString("yyyy-MM-dd");
                dtpToDate.Text = DateTime.Now.ToString("yyyy-MM-dd");   // 오늘 날짜 자동 반영

                rbnOrderNumber.IsChecked = true;

                rbnOrderNumber_Click(null, null);

                CantBtnControl();
                SetComboBox();
            }
            catch (Exception ee)
            {
                MessageBox.Show("오류지점 - UserControl_Loaded : " + ee.ToString());
            }
        }

        #region 콤보박스
        private void SetComboBox()
        {
            try
            {
                ObservableCollection<CodeView> cbOutClss = ComboBoxUtil.Instance.Gf_DB_CM_GetComCodeDataset(null, "OCD", "Y", "", "PROD");
                this.cboOutClss.ItemsSource = cbOutClss;
                this.cboOutClss.DisplayMemberPath = "code_name";
                this.cboOutClss.SelectedValuePath = "code_id";
                this.cboOutClss.SelectedIndex = 0;

                ObservableCollection<CodeView> cbFromLoc = ComboBoxUtil.Instance.Gf_DB_CM_GetComCodeDataset(null, "LOC", "Y", "", "INSIDE");
                this.cboFromLoc.ItemsSource = cbFromLoc;
                this.cboFromLoc.DisplayMemberPath = "code_name";
                this.cboFromLoc.SelectedValuePath = "code_id";
                this.cboFromLoc.SelectedIndex = 0;

                ObservableCollection<CodeView> cbToLoc = ComboBoxUtil.Instance.Gf_DB_CM_GetComCodeDataset(null, "LOC", "Y", "", "NONE");
                this.cboToLoc.ItemsSource = cbToLoc;
                this.cboToLoc.DisplayMemberPath = "code_name";
                this.cboToLoc.SelectedValuePath = "code_id";
                this.cboToLoc.SelectedIndex = 0;
            }
            catch (Exception ee)
            {
                MessageBox.Show("오류지점 - SetComboBox : " + ee.ToString());
            }
        }
        #endregion 콤보박스

        #region 상단 레이아웃 조건 모음
        //출고일자 라벨 클릭시
        private void lblOutwareDay_Click(object sender, MouseButtonEventArgs e)
        {
            if (chkOutwareDay.IsChecked == true)
            {
                chkOutwareDay.IsChecked = false;

                dtpFromDate.IsEnabled = false;
                dtpToDate.IsEnabled = false;
            }
            else
            {
                chkOutwareDay.IsChecked = true;

                dtpFromDate.IsEnabled = true;
                dtpToDate.IsEnabled = true;
            }
        }

        //출고일자 체크 
        private void ChkOutwareDay_Checked(object sender, RoutedEventArgs e)
        {
            chkOutwareDay.IsChecked = true;

            dtpFromDate.IsEnabled = true;
            dtpToDate.IsEnabled = true;

        }

        //출고일자 체크해제
        private void ChkOutwareDay_Unchecked(object sender, RoutedEventArgs e)
        {
            chkOutwareDay.IsChecked = false;

            dtpFromDate.IsEnabled = false;
            dtpToDate.IsEnabled = false;
        }

        //전일
        private void btnYesterday_Click(object sender, RoutedEventArgs e)
        {
            //dtpFromDate.SelectedDate = DateTime.Today.AddDays(-1);
            //dtpToDate.SelectedDate = DateTime.Today.AddDays(-1);

            try
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
            }
            catch (Exception ee)
            {
                MessageBox.Show("오류지점 - btnYesterday_Click : " + ee.ToString());
            }
        }

        //금일
        private void btnToday_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                dtpFromDate.SelectedDate = DateTime.Today;
                dtpToDate.SelectedDate = DateTime.Today;
            }
            catch (Exception ee)
            {
                MessageBox.Show("오류지점 - btnToday_Click : " + ee.ToString());
            }
        }

        //전월
        private void btnLastMonth_Click(object sender, RoutedEventArgs e)
        {
            //dtpFromDate.SelectedDate = Lib.Instance.BringLastMonthDatetimeList()[0];
            //dtpToDate.SelectedDate = Lib.Instance.BringLastMonthDatetimeList()[1];

            try
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
            }
            catch (Exception ee)
            {
                MessageBox.Show("오류지점 - btnLastMonth_Click : " + ee.ToString());
            }
        }

        //금월
        private void btnThisMonth_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                dtpFromDate.SelectedDate = Lib.Instance.BringThisMonthDatetimeList()[0];
                dtpToDate.SelectedDate = Lib.Instance.BringThisMonthDatetimeList()[1];
            }
            catch (Exception ee)
            {
                MessageBox.Show("오류지점 - btnThisMonth_Click : " + ee.ToString());
            }
        }

  

 
        //거래처-조건 텍스트박스 키다운 이벤트
        private void txtCustomerIDSrh_KeyDown(object sender, KeyEventArgs e)
        {
            try
            {
                if (e.Key == Key.Enter)
                {
                    pf.ReturnCode(txtCustomIDSrh, 0, "");
                }
            }
            catch (Exception ee)
            {
                MessageBox.Show("오류지점 - txtCustomer_KeyDown : " + ee.ToString());
            }
        }

        //거래처-조건 플러스파인더 버튼
        private void btnCustomerIDSrh_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                pf.ReturnCode(txtCustomIDSrh, 0, "");
            }
            catch (Exception ee)
            {
                MessageBox.Show("오류지점 - btnCustomer_Click : " + ee.ToString());
            }
        }



        //최종고객사 라벨 클릭시
        private void lblInCustomer_Click(object sender, MouseButtonEventArgs e)
        {
            if (chkInCustomer.IsChecked == true)
            {
                chkInCustomer.IsChecked = false;
                txtInCustomer.IsEnabled = false;
                btnInCustomer.IsEnabled = false;
            }
            else
            {
                chkInCustomer.IsChecked = true;
                txtInCustomer.IsEnabled = true;
                btnInCustomer.IsEnabled = true;
                txtInCustomer.Focus();
            }
        }

        //최종고객사 체크
        private void ChkInCustomer_Checked(object sender, RoutedEventArgs e)
        {
            chkInCustomer.IsChecked = true;
            txtInCustomer.IsEnabled = true;
            btnInCustomer.IsEnabled = true;
            txtInCustomer.Focus();
        }

        //최종고객사 체크 해제
        private void ChkInCustomer_Unchecked(object sender, RoutedEventArgs e)
        {
            chkInCustomer.IsChecked = false;
            txtInCustomer.IsEnabled = false;
            btnInCustomer.IsEnabled = false;
        }

        //거래처-조건 텍스트박스 키다운 이벤트
        private void txtInCustomer_KeyDown(object sender, KeyEventArgs e)
        {
            try
            {
                if (e.Key == Key.Enter)
                {
                    pf.ReturnCode(txtInCustomer, 0, "");
                }
            }
            catch (Exception ee)
            {
                MessageBox.Show("오류지점 - txtInCustomer_KeyDown : " + ee.ToString());
            }
        }

        //거래처-조건 플러스파인더 버튼
        private void btnInCustomer_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                pf.ReturnCode(txtInCustomer, 0, "");
            }
            catch (Exception ee)
            {
                MessageBox.Show("오류지점 - btnInCustomer_Click : " + ee.ToString());
            }
        }



    

        //품명 텍스트박스 키다운 이벤트
        private void txtArticleIDSrh_KeyDown(object sender, KeyEventArgs e)
        {
            try
            {
                if (e.Key == Key.Enter)
                {
                    pf.ReturnCode(txtArticleIDSrh, 77, txtArticleIDSrh.Text);
                }
            }
            catch (Exception ee)
            {
                MessageBox.Show("오류지점 - txtArticle_KeyDown : " + ee.ToString());
            }
        }

        //품명 플러스파인더 버튼
        private void btnArticleIDSrh_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                pf.ReturnCode(txtArticleIDSrh, 77, txtArticleIDSrh.Text);
            }
            catch (Exception ee)
            {
                MessageBox.Show("오류지점 - btnArticle_Click : " + ee.ToString());
            }
        }

    

        //관리번호 라벨 클릭시
        private void lblRadioOptionNum_Click(object sender, MouseButtonEventArgs e)
        {
            if (chkRadioOptionNum.IsChecked == true)
            {
                chkRadioOptionNum.IsChecked = false;
                txtRadioOptionNum.IsEnabled = false;
            }
            else
            {
                chkRadioOptionNum.IsChecked = true;
                txtRadioOptionNum.IsEnabled = true;
                txtRadioOptionNum.Focus();
            }
        }

        //관리번호 체크
        private void ChkRadioOptionNum_Checked(object sender, RoutedEventArgs e)
        {
            chkRadioOptionNum.IsChecked = true;
            txtRadioOptionNum.IsEnabled = true;
            txtRadioOptionNum.Focus();
        }

        //관리번호 체크 해제
        private void ChkRadioOptionNum_Unchecked(object sender, RoutedEventArgs e)
        {
            chkRadioOptionNum.IsChecked = false;
            txtRadioOptionNum.IsEnabled = false;
        }

        //라디오버튼 OrderNo 버튼 클릭
        private void rbnOrderNo_Click(object sender, RoutedEventArgs e)
        {
            Check_bdrOrder();
        }

        //라디오버튼 OrderID 버튼 클릭
        private void rbnManageNum_Click(object sender, RoutedEventArgs e)
        {
            Check_bdrOrder();
        }

        private void Check_bdrOrder()
        {
            if (rbnManageNum.IsChecked == true)
            {
                tbkRadioOptionNum.Text = " 관리번호";
                dgdtxtcol_ManageNum.Visibility = Visibility.Visible;
                dgdtxtcol_OrderNo.Visibility = Visibility.Hidden;
            }
            else if (rbnOrderNo.IsChecked == true)
            {
                tbkRadioOptionNum.Text = " 발주번호";
                dgdtxtcol_ManageNum.Visibility = Visibility.Hidden;
                dgdtxtcol_OrderNo.Visibility = Visibility.Visible;
            }
        }

        private void Check_SwitchOrderNoAndReqID()
        {
            if (rbnOrderNumber.IsChecked == true) //관리번호 클릭
            {
        
              
                txtOrderID.Visibility = Visibility.Visible;
                btnOrderID.Visibility = Visibility.Visible;
                lblOrderID.Visibility = Visibility.Visible;

                dgdtxtcol_ManageNum.Visibility = Visibility.Visible;

            }
            else if (rbnReqID.IsChecked == true) //지시번호 클릭
            {
               
                txtOrderID.Visibility = Visibility.Hidden;
                btnOrderID.Visibility = Visibility.Hidden;
                lblOrderID.Visibility = Visibility.Hidden;

                dgdtxtcol_ManageNum.Visibility = Visibility.Hidden;
            }
        }


        #endregion

        #region 버튼 모음
        //추가버튼 클릭
        private void btnAdd_Click(object sender, RoutedEventArgs e)
        {
            //2021-06-02
            EventStatus = true;
            TextBoxClear(); // 추가누르면 다시 클리어 해 줘야지
            try
            {
                strFlag = "I";

                this.DataContext = null;
                CanBtnControl();                             //버튼 컨트롤
                dtpOutDate.SelectedDate = DateTime.Today;

                txtOrderID.Focus();                          //관리번호에 포커스 이동

                cboOutClss.SelectedIndex = 0;
                cboFromLoc.SelectedIndex = 0; //사내제품창고가 기본값이 되게 설정
                cboToLoc.SelectedIndex = 0;

                dgdOutwareSub.Items.Clear();            

            }
            catch (Exception ee)
            {
                MessageBox.Show("오류지점 - btnAdd_Click : " + ee.ToString());
            }
        }

        //수정버튼 클릭
        private void btnUpdate_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var OutwareItem = dgdOutware.SelectedItem as Win_ord_OutWare_Scan_CodeView;

                if (OutwareItem != null)
                {
                    strFlag = "U";

                    rowNum = dgdOutware.SelectedIndex;
                    CanBtnControl();
                }
            }
            catch (Exception ee)
            {
                MessageBox.Show("오류지점 - btnUpdate_Click : " + ee.ToString());
            }
        }

        //삭제버튼 클릭
        private void btnDelete_Click(object sender, RoutedEventArgs e)
        {
            using (Loading ld = new Loading(beDelete))
            {
                ld.ShowDialog();
            }
        }

        private void beDelete()
        {
            btnDelete.IsEnabled = false;

            Dispatcher.BeginInvoke(new Action(() =>
            {
                if (lstOutwarePrint.Count == 0)
                {
                    MessageBox.Show("삭제할 데이터가 지정되지 않았습니다. 삭제 데이터를 지정하고 눌러주세요.");
                }
                else
                {
                    if (MessageBox.Show("선택하신 항목을 삭제하시겠습니까?", "삭제 전 확인", MessageBoxButton.YesNo) == MessageBoxResult.Yes)
                    {
                        foreach (Win_ord_OutWare_Scan_CodeView RemoveData in lstOutwarePrint)
                            DeleteData(RemoveData.OutwareID);

                        rowNum = 0;
                        re_Search(rowNum);
                    }
                }
            }), System.Windows.Threading.DispatcherPriority.Background);

            btnDelete.IsEnabled = true;
        }

        //닫기버튼 클릭
        private void btnClose_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                DataStore.Instance.InsertLogByFormS(this.GetType().Name, stDate, stTime, "E");
                Lib.Instance.ChildMenuClose(this.ToString());
            }
            catch (Exception ee)
            {
                MessageBox.Show("오류지점 - btnClose_Click : " + ee.ToString());
            }
        }

        //검색버튼 클릭
        private void btnSearch_Click(object sender, RoutedEventArgs e)
        {
            if(lib.DatePickerCheck(dtpFromDate, dtpToDate, chkOutwareDay))
            {
                using (Loading ld = new Loading(beSearch))
                {
                    ld.ShowDialog();
                }
            } 
        }

        private void beSearch()
        {
            btnSearch.IsEnabled = false;

            Dispatcher.BeginInvoke(new Action(() =>
            {
                //로직
                rowNum = 0;
                re_Search(rowNum);
            }), System.Windows.Threading.DispatcherPriority.Background);

            btnSearch.IsEnabled = true;
        }

        //저장버튼 클릭
        private void btnSave_Click(object sender, RoutedEventArgs e)
        {
            using (Loading ld = new Loading(beSave))
            {
                ld.ShowDialog();
            }
        }

        private void beSave()
        {
            btnSave.IsEnabled = false;

            Dispatcher.BeginInvoke(new Action(() =>
            {
              //버튼 컨트롤

                if (SaveData(strFlag))
                {
                    CantBtnControl();

                    if (strFlag.Equals("I"))
                    {
                        var outwareCount = dgdOutware.Items.Count;
                        rowNum = outwareCount;
                    }
                    
                    lstCheck.Clear();
                    strFlag = string.Empty;
                    TextBoxClear(); // 저장했으면 클리어 해야지
                    re_Search(rowNum);
                    EventStatus = false;
                }
            }), System.Windows.Threading.DispatcherPriority.Background);

            btnSave.IsEnabled = true;
        }

        //취소버튼 클릭
        private void btnCancel_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                EventStatus = false;
                CantBtnControl();           //버튼 컨트롤
                TextBoxClear();
                lstCheck.Clear();

                if (strFlag.Equals("I"))
                {
                    re_Search(0);
                }
                else if (strFlag.Equals("U"))
                {
                    re_Search(rowNum);
                }

                strFlag = string.Empty;
            }
            catch (Exception ee)
            {
                MessageBox.Show("오류지점 - btnCancel_Click : " + ee.ToString());
            }
        }

        //엑셀버튼 클릭
        private void btnExcel_Click(object sender, RoutedEventArgs e)
        {
            Lib lib2 = new Lib();
            try
            {
                if (dgdOutware.Items.Count < 1)
                {
                    MessageBox.Show("먼저 검색해 주세요.");
                    return;
                }
                DataTable dt = null;
                string Name = string.Empty;

                string[] lst = new string[4];
                lst[0] = "메인그리드";
                lst[1] = "서브그리드";
                lst[2] = dgdOutware.Name;
                lst[3] = dgdOutwareSub.Name;

                ExportExcelxaml ExpExc = new ExportExcelxaml(lst);

                ExpExc.ShowDialog();

                if (ExpExc.DialogResult.HasValue)
                {
                    if (ExpExc.choice.Equals(dgdOutware.Name))
                    {
                        DataStore.Instance.InsertLogByForm(this.GetType().Name, "E");
                        //MessageBox.Show("대분류");
                        if (ExpExc.Check.Equals("Y"))
                            dt = lib2.DataGridToDTinHidden(dgdOutware);
                        else
                            dt = lib2.DataGirdToDataTable(dgdOutware);

                        Name = dgdOutware.Name;
                        if (lib2.GenerateExcel(dt, Name))
                        {
                            lib2.excel.Visible = true;
                            lib2.ReleaseExcelObject(lib2.excel);
                        }
                    }
                    else if (ExpExc.choice.Equals(dgdOutwareSub.Name))
                    {
                        //MessageBox.Show("정성류");
                        if (ExpExc.Check.Equals("Y"))
                            dt = lib2.DataGridToDTinHidden(dgdOutwareSub);
                        else
                            dt = lib2.DataGirdToDataTable(dgdOutwareSub);
                        Name = dgdOutwareSub.Name;
                        if (lib2.GenerateExcel(dt, Name))
                        {
                            lib2.excel.Visible = true;
                            lib2.ReleaseExcelObject(lib2.excel);
                        }
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
                MessageBox.Show("오류지점 - btnExcel_Click : " + ee.ToString());
            }
            finally
            {
                lib2 = null;
            }
        }

        //인쇄버튼 클릭
        private void btnPrint_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                ContextMenu menu = btnPrint.ContextMenu;
                menu.StaysOpen = true;
                menu.IsOpen = true;
            }
            catch (Exception ee)
            {
                MessageBox.Show("오류지점 - btnPrint_Click : " + ee.ToString());
            }
        }

        //인쇄-미리보기 클릭
        private void menuSeeAhead_Click(object sender, RoutedEventArgs e)
        {
            menuPrint_Click(true);
        }

        //인쇄-바로인쇄 클릭
        private void menuRighPrint_Click(object sender, RoutedEventArgs e)
        {
            menuPrint_Click(false);
        }

        private void menuPrint_Click(bool Ahead)
        {
            try
            {
                if (dgdOutware.Items.Count == 0)
                {
                    MessageBox.Show("먼저 검색해 주세요.");
                    return;
                }

                var OBJ = dgdOutware.SelectedItem as Win_ord_OutWare_Scan_CodeView;
                if (OBJ == null)
                {
                    MessageBox.Show("거래명세표 항목이 정확히 선택되지 않았습니다.");
                    return;
                }

                List<Win_ord_OutWare_Scan_CodeView> find = lstOutwarePrint.FindAll(
                    delegate (Win_ord_OutWare_Scan_CodeView a)
                    {
                        return a.CustomID == lstOutwarePrint[0].CustomID &&
                               a.OutCustomID == lstOutwarePrint[0].OutCustomID;
                    }
                );

                if (lstOutwarePrint.Count != find.Count)
                {
                    MessageBox.Show("동일한 거래처, 최종고객사만 선택할 수 있습니다.");
                    return;
                }

                preview_click = Ahead;

                DataStore.Instance.InsertLogByForm(this.GetType().Name, "P");
                /*msg.Show();
                msg.Topmost = true;
                msg.Refresh();
                msg.Visibility = Visibility.Hidden;*/

                using (Loading ld = new Loading("excel", PrintWork))
                {
                    ld.ShowDialog();
                }
            }
            catch (Exception ee)
            {
                MessageBox.Show("오류지점 - menuRighPrint_Click : " + ee.ToString());
            }
        }

        //인쇄-닫기 클릭
        private void menuClose_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                ContextMenu menu = btnPrint.ContextMenu;
                menu.StaysOpen = false;
                menu.IsOpen = false;
            }
            catch (Exception ee)
            {
                MessageBox.Show("오류지점 - menuClose_Click : " + ee.ToString());
            }
        }

        // 실제 엑셀작업 스타트.
        private void PrintWork()
        {
            Lib lib2 = new Lib();
            try
            {
                if (lstOutwarePrint.Count == 0)
                {
                    MessageBox.Show("인쇄할 거래명세표를 체크 선택하세요.");
                    lib2 = null;
                    return;
                }

                excelapp = new Microsoft.Office.Interop.Excel.Application();

                string MyBookPath = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetEntryAssembly().Location) + "\\Report\\거래명세표.xlsx";
                workbook = excelapp.Workbooks.Add(MyBookPath);
                worksheet = workbook.Sheets["Form"];
                pastesheet = workbook.Sheets["Print"];

                // 거래일자
                workrange = worksheet.get_Range("C4", "H4");
                workrange.Value2 = lstOutwarePrint[0].OutDate.Replace("-", ".");

                // 공급받는 사업자번호
                workrange = worksheet.get_Range("G5", "R6");
                workrange.Value2 = lstOutwarePrint[0].CustomNo;

                // 공급받는 상호
                workrange = worksheet.get_Range("G7", "L8");
                workrange.Value2 = lstOutwarePrint[0].KCustom;

                // 공급받는 사업장 주소
                workrange = worksheet.get_Range("G9", "R10");
                workrange.Value2 = lstOutwarePrint[0].Buyer_Address1 + lstOutwarePrint[0].Buyer_Address2 + lstOutwarePrint[0].Buyer_Address3;

                // 공급받는 성명
                workrange = worksheet.get_Range("O7", "R8");
                workrange.Value2 = lstOutwarePrint[0].Buyer_Chief;

                // 공급자 정보 구해오기.
                DataTable DT = Fill_DS_CompanyInfo();
                DataRow DR = DT.Rows[0];

                // 공급자 등록번호 (사업자등록번호)
                string companyNo = DR["CompanyNo"].ToString();
                workrange = worksheet.get_Range("W5", "AH6");
                workrange.Value2 = companyNo.Substring(0, 3) + "-" + companyNo.Substring(3, 2) + "-" + companyNo.Substring(5, 5);

                // 공급자 상호
                workrange = worksheet.get_Range("W7", "AC8");
                workrange.Value2 = DR["KCompany"].ToString();

                // 공급자 성명
                workrange = worksheet.get_Range("AF7", "AH8");
                workrange.Value2 = DR["Chief"].ToString();

                // 공급자 사업장 주소
                workrange = worksheet.get_Range("W9", "AH10");
                workrange.Value2 = DR["Address1"].ToString();

                // 공급자 전화
                //workrange = worksheet.get_Range("W11", "AB12");
                //workrange.Value2 = DR["Phone1"].ToString();

                // 공급자 팩스
                //workrange = worksheet.get_Range("AD11", "AH12");
                //workrange.Value2 = DR["FaxNo"].ToString();

                int copyLine = 1;
                int copyRow = 54;

                int inputPossibleRowCnt = 10;   // 내역 입력 가능한 갯수
                int startRowNum = 14;           // 내역 입력 시작점
                int endCnt = 0;                 // 엑셀 입력 종료 갯수

                int cnt = 0;
                int totCnt = 0;

                int pageCnt = 1;
                int totPageCnt = (lstOutwarePrint.Count / inputPossibleRowCnt) + 1;

                double totalSumAmount = 0, totalSumVatAmount = 0, totalAmount = 0;

                // key : 출고일, value : 출고 항목
                Dictionary<string, List<Win_ord_OutWare_Scan_Sub_CodeView>> dic = new Dictionary<string, List<Win_ord_OutWare_Scan_Sub_CodeView>>();

                lstOutwarePrint.Sort((x, y) => x.OutDate.CompareTo(y.OutDate));

                // 합계 먼저 계산
                foreach (Win_ord_OutWare_Scan_CodeView outware in lstOutwarePrint)
                {
                    string outwareID = outware.OutwareID;
                    string UnitClssName = outware.UnitClssName;

                    Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                    sqlParameter.Add("OutwareID", outwareID);
                    DataSet ds = DataStore.Instance.ProcedureToDataSet("xp_Outware_sOutwareSubGroup", sqlParameter, false);
                    if (ds != null && ds.Tables.Count > 0)
                    {
                        DataTable dt = ds.Tables[0];
                        if (dt.Rows.Count > 0)
                        {
                            double sumOutQty = 0, sumUnitPrice = 0;

                            DataRowCollection drc = dt.Rows;
                            foreach (DataRow item in drc)
                            {
                                double outQty = ConvertDouble(item["OutQty"].ToString());
                                double unitPrice = ConvertDouble(item["UnitPrice"].ToString());
                                double calcAmount = outQty * unitPrice;
                                double calcVatAmount = calcAmount * .1;

                                sumOutQty += outQty;
                                sumUnitPrice += unitPrice;

                                totalSumAmount += calcAmount;
                                totalSumVatAmount += calcVatAmount;
                                totalAmount += totalSumAmount + totalSumVatAmount;
                            }

                            Win_ord_OutWare_Scan_Sub_CodeView sub = new Win_ord_OutWare_Scan_Sub_CodeView();
                            sub.ArticleID = outware.ArticleID;
                            sub.Article = outware.Article;
                            sub.dOutQty = sumOutQty;
                            sub.UnitClssName = "EA";
                            sub.dUnitPrice = sumUnitPrice;

                            string outDate = outware.OutDate.Replace("-", "");
                            if (dic.ContainsKey(outDate))
                            {
                                int findIdx = dic[outDate].FindIndex(x => x.ArticleID == outware.ArticleID);
                                if (findIdx > -1)
                                {
                                    dic[outDate][findIdx].dOutQty += sumOutQty;
                                    dic[outDate][findIdx].dUnitPrice += sumUnitPrice;
                                }
                                else
                                {
                                    dic[outDate].Add(sub);
                                    endCnt++;
                                }
                            }
                            else
                            {
                                List<Win_ord_OutWare_Scan_Sub_CodeView> listSub = new List<Win_ord_OutWare_Scan_Sub_CodeView>() { sub };
                                dic.Add(outDate, listSub);
                                endCnt++;
                            }
                        }
                    }
                }

                foreach (KeyValuePair<string, List<Win_ord_OutWare_Scan_Sub_CodeView>> pair in dic)
                {
                    string outDate = pair.Key;
                    List<Win_ord_OutWare_Scan_Sub_CodeView> listSub = pair.Value;

                    string month = outDate.Substring(4, 2);
                    string day = outDate.Substring(6, 2);

                    for (int i = 0; i < listSub.Count; i++)
                    {
                        int rowNum = startRowNum + (cnt % (inputPossibleRowCnt + 1));

                        // 월
                        workrange = worksheet.get_Range("C" + rowNum.ToString());
                        workrange.Value2 = i + 1;

                        workrange = worksheet.get_Range("M" + rowNum.ToString());
                        workrange.Value2 = listSub[i].UnitClssName;

                        //// 일
                        //workrange = worksheet.get_Range("D" + rowNum.ToString());
                        //workrange.Value2 = day;

                        // 품명
                        workrange = worksheet.get_Range("E" + rowNum.ToString(), "L" + rowNum.ToString());
                        workrange.Value2 = listSub[i].Article;

                        //workrange = worksheet.get_Range("E" + rowNum.ToString(), "L" + rowNum.ToString());
                        //workrange.Value2 = listSub[i].UnitClss;

                        string strOutQty = stringFormatN0(listSub[i].dOutQty);
                        string strUnitPrice = stringFormatN0(listSub[i].dUnitPrice);
                        double calcAmount = ConvertDouble(strOutQty) * ConvertDouble(strUnitPrice);
                        double calcValAmount = calcAmount * .1;
                        double calcSumAmount = calcAmount + calcValAmount;

                        // 수량
                        workrange = worksheet.get_Range("P" + rowNum.ToString(), "Q" + rowNum.ToString());
                        workrange.Value2 = strOutQty;

                        // 단가
                        workrange = worksheet.get_Range("R" + rowNum.ToString(), "V" + rowNum.ToString());
                        workrange.Value2 = strUnitPrice;

                        // 공급가액
                        workrange = worksheet.get_Range("W" + rowNum.ToString(), "Z" + rowNum.ToString());
                        workrange.Value2 = lib.returnNumStringTargetNum(calcAmount.ToString(), 3);

                        // 세액
                        workrange = worksheet.get_Range("AA" + rowNum.ToString(), "AD" + rowNum.ToString());
                        workrange.Value2 = lib.returnNumStringTargetNum(calcValAmount.ToString(), 3);

                        //합계
                        workrange = worksheet.get_Range("AE" + rowNum.ToString(), "AH" + rowNum.ToString());
                        workrange.Value2 = lib.returnNumStringTargetNum(calcSumAmount.ToString(), 3);

                        // 비고
                        /*workrange = worksheet.get_Range("AG" + rowNum.ToString(), "AH" + rowNum.ToString());
                        workrange.Value2 = outware.Remark;*/

                        cnt++;
                        totCnt++;

                        // 거래명세표 다음 및 종료 조건
                        if (totCnt == endCnt || cnt == inputPossibleRowCnt)
                        {
                            // 페이지수
                            workrange = worksheet.get_Range("AC3", "AH4");
                            workrange.NumberFormat = "@";
                            workrange.Value2 = pageCnt.ToString() + "/" + totPageCnt.ToString();

                            // 합계 공급가액 
                            workrange = worksheet.get_Range("G24", "L24");
                            workrange.Value2 = totalSumAmount.ToString();

                            // 합계 세액 
                            workrange = worksheet.get_Range("Q24", "W24");
                            workrange.Value2 = totalSumVatAmount.ToString();

                            // 합계 총액 
                            workrange = worksheet.get_Range("AB24", "AH24");
                            workrange.Value2 = totalAmount.ToString();


                            // 붙여넣기
                            worksheet.Select();
                            worksheet.UsedRange.EntireRow.Copy();
                            pastesheet.Select();
                            workrange = pastesheet.Rows[copyLine];
                            workrange.Select();
                            pastesheet.Paste();

                            // 내역 삭제
                            workrange = worksheet.get_Range("C" + startRowNum.ToString(), "AH" + (startRowNum + inputPossibleRowCnt - 1).ToString());
                            workrange.ClearContents();

                            copyLine += copyRow;

                            cnt = 0;
                            pageCnt++;
                        }
                    }
                }

                excelapp.Visible = true;

                if (preview_click)
                    pastesheet.PrintPreview();
                else
                    pastesheet.PrintOutEx();
            }
            catch (Exception ee)
            {
                MessageBox.Show("오류지점 = PrintWork : " + ee.ToString());
            }

      
            lib2.ReleaseExcelObject(workrange);
     
            lib2.ReleaseExcelObject(pastesheet);
            lib2.ReleaseExcelObject(worksheet);
            lib2.ReleaseExcelObject(workbook);
            lib2.ReleaseExcelObject(excelapp);

            GC.Collect();
            GC.WaitForPendingFinalizers();
            GC.Collect();
            GC.WaitForPendingFinalizers();

            lib2 = null;
        }

        // 거래명세표 인쇄시 공급자 정보 구해오기
        private DataTable Fill_DS_CompanyInfo()
        {
            try
            {
                Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                sqlParameter.Add("nChkCompany", 0);
                sqlParameter.Add("sCompanyID", "");
                sqlParameter.Add("sKCompany", "");

                DataSet ds = DataStore.Instance.ProcedureToDataSet("xp_Info_GetCompanyInfo", sqlParameter, false);

                if (ds != null && ds.Tables.Count > 0)
                {
                    DataTable ddt = null;
                    ddt = ds.Tables[0];

                    if (ddt.Rows.Count == 0)
                    {
                        MessageBox.Show("공급자 정보를 구하지 못했습니다.");
                        return ddt;
                    }
                    else
                    {
                        return ddt;
                    }
                }
                else
                {
                    return null;
                }
            }
            catch (Exception ee)
            {
                MessageBox.Show("오류지점 - Fill_DS_CompanyInfo : " + ee.ToString());
                return null;
            }
        }
        #endregion


        // 플러스 파인더 품명 이벤트
        string replyArticle = "";
        string replyOrderQty = "";
        string replyReqQty = "";
        string replyReqRemainQty = "";

        //private void plusFinder_replyArticle(string article)
        //{
        //    replyArticle = article;
        //    pf.refEvent -= new PlusFinder.RefEventHandler(plusFinder_replyArticle);
        //}

        private void plusFinder_replyArticle(string data)
        {
            string[] values = data.Split(',');
            if (values.Length > 0)
                replyArticle = values[0].Trim();
            else
                replyArticle = string.Empty;
        }

        //플러스파인더에서 대리자 이벤트로 가져온 수주량
        #region ...
        //private void plusFinder_replyOrderQty(string data)
        //{
        //    if(rbnReqID.IsChecked == true)
        //    {


        //        string[] values = data.Split(','); 

        //        if (values.Length > 1)
        //            replyOrderQty = values[1].Trim();
        //        else
        //            replyOrderQty = string.Empty;
        //    }
        //    else
        //    {
        //        string[] values = data.Split(',');
        //        if (values.Length > 1)
        //            replyOrderQty = values[0].Trim();
        //        else
        //            replyOrderQty = string.Empty;
        //    }

        //}
        #endregion

        private void plusFinder_replyOrderQty(string data)
        {
            string[] values = data.Split(',');          
            if (values.Length > 0)
                replyOrderQty = values[0].Trim();
            else
                replyOrderQty = string.Empty;  
        }

        //플러스파인더에서 대리자 이벤트로 가져온 지시량
        private void plusFinder_replyReqQty(string data)
        {
            string[] values = data.Split(',');
            if (values.Length > 1)
                replyReqQty = values[2].Trim();
            else
                replyReqQty = string.Empty;
        }

        //플러스파인더에서 대리자 이벤트로 가져온 남은지시량
        private void plusFinder_replyReqRemainQty(string data)
        {
            string[] values = data.Split(',');
            if (values.Length > 1)
                replyReqRemainQty = values[3].Trim();
            else
                replyReqRemainQty = string.Empty;
        }

        #region 키다운 이동 모음
   

        //관리번호 텍스트박스 키다운 이벤트
        private void txtOrderID_KeyDown(object sender, KeyEventArgs e)
        {
            try
            {
                if (e.Key == Key.Enter)
                {                       

                    pf.refEvent += new PlusFinder.RefEventHandler(plusFinder_replyOrderQty);
                    pf.ReturnCode(txtOrderID, 99, txtOrderID.Text);

                    if (txtOrderID.Text.Length > 0)
                    {
                  
                        OrderID_OtherSearch(txtOrderID.Text, "");

                
                    }                   
                }
            }
            catch (Exception ee)
            {
                MessageBox.Show("오류지점 - txtOrderID_KeyDown : " + ee.ToString());
            }
        }

        private void OutwareSubInput()
        {

            if(dgdOutwareSub.Items.Count > 0)
            {
                dgdOutwareSub.Items.Clear();
            }

            // 수량 입력시 라벨 없이 입력됨
            Win_ord_OutWare_Scan_Sub_CodeView label = new Win_ord_OutWare_Scan_Sub_CodeView();

            int num = dgdOutwareSub.Items.Count + 1;
            label.Num = num;
            label.LabelID = "";
            label.Spec = "";
            label.Orderseq = orderSeq;
            label.OutQty = stringFormatN0(replyOrderQty);
            dgdOutwareSub.Items.Add(label);

            SumScanQty();
        }

        //품명 텍스트박스 키다운 이벤트
        private void EnterMove_KeyDown(object sender, KeyEventArgs e)
        {
            //품명도 땡겨와서 텍스트 박스 막음
        }

        //차종 텍스트박스 키다운 이벤트
        private void txtBuyerModel_KeyDown(object sender, KeyEventArgs e)
        {
            //차종은 땡겨와서 텍스트 박스 막음
        }

        //출고구분 콤보박스 닫힘
        private void cboOutClss_DropDownClosed(object sender, EventArgs e)
        {
            dtpOutDate.IsDropDownOpen = true;
        }

        //작성일자 달력 닫힘
        private void DtpOutDate_CalendarClosed(object sender, RoutedEventArgs e)
        {
            txtBuyerName.Focus();
        }

        //박스 키다운 이벤트
        private void TxtOutRoll_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                txtOutQty.Focus();
            }
        }

        //수량 키다운 이벤트
        private void TxtOutQty_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                txtKCustom.Focus();
            }
        }

        //전창고 콤보박스 닫힘
        private void cboFromLoc_DropDownClosed(object sender, EventArgs e)
        {
            txtRemark.Focus();
            //cboToLoc.IsDropDownOpen = true;
        }

        //수주거래처 키다운 이벤트
        private void TxtKCustom_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                txtBuyerName.Focus();
            }
        }

        //후창고 콤보박스 닫힘
        private void cboToLoc_DropDownClosed(object sender, EventArgs e)
        {
            txtRemark.Focus();
        }

        //납품거래처 텍스트박스 키다운 이벤트
        private void txtBuyerName_KeyDown(object sender, KeyEventArgs e)
        {
            try
            {
                if (e.Key == Key.Enter)
                {
                    pf.ReturnCode(txtBuyerName, 0, "");

                    if (txtBuyerName.Text.Length > 0)
                    {
                        txtOutCustom.Text = txtBuyerName.Text;
                    }

                    txtOutCustom.Focus();
                }
            }
            catch (Exception ee)
            {
                MessageBox.Show("오류지점 - txtBuyerName_KeyDown : " + ee.ToString());
            }
        }

        //비고 키다운 이벤트
        private void txtRemark_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                //비고에서 엔터를 누르면 바코드 스캔하는 위치로 이동
                txtScanData.Focus();
            }
        }

        //출고처 키다운 이벤트
        private void TxtOutCustom_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                pf.ReturnCode(txtOutCustom, 0, "");

                if (txtOutCustom.Text.Length > 0)
                    txtBuyerName.Text = txtOutCustom.Text;
                    txtBuyerName.Tag = txtOutCustom.Tag;

                txtBuyerName.Focus();
            }
        }
        #endregion

        #region 플러스파인더 및 데이터그리드 선택 변경

        //메인 데이터그리드 선택 변경
        private void dgdOutware_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            try
            {
                var OutwareInfo = dgdOutware.SelectedItem as Win_ord_OutWare_Scan_CodeView;

                if (OutwareInfo != null)
                {
                    this.DataContext = OutwareInfo;                           

                    FillGridSub(OutwareInfo.OutwareID);
                }
            }
            catch (Exception ee)
            {
                MessageBox.Show("오류지점 - dgdOutware_SelectionChanged : " + ee.ToString());
            }
        }

  

        //관리번호 플러스파인더 버튼 클릭
        private void btnOrderID_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                pf.refEvent += new PlusFinder.RefEventHandler(plusFinder_replyOrderQty);
                pf.ReturnCode(txtOrderID, 99, txtOrderID.Text);
                if (txtOrderID.Text.Length > 0)
                {
                    //관리번호 기반_ 항목 뿌리기 작업.
                    OrderID_OtherSearch(txtOrderID.Text, "");

                    //항목 뿌린후 Subgrid에 수량 자동입력
                    //박스수량 계산
                    OutwareSubInput();

                    cboOutClss.IsDropDownOpen = true;
                }               
            }
            catch (Exception ee)
            {
                MessageBox.Show("오류지점 - btnOrderID_Click : " + ee.ToString());
            }
        }

        //납품거래처 플러스파인더 버튼 
        private void btnOutCustom_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                pf.ReturnCode(txtOutCustom, 0, "");

                if (txtOutCustom.Text.Length > 0)
                    txtBuyerName.Text = txtOutCustom.Text;

                txtBuyerName.Focus();
            }
            catch (Exception ee)
            {
                MessageBox.Show("오류지점 - btnOutCustom_Click : " + ee.ToString());
            }
        }

        //라벨스캔 텍스트박스 키다운 이벤트
        private void txtScanData_KeyDown(object sender, KeyEventArgs e)
        {
            try
            {
                if (e.Key == Key.Enter)
                {
                    if (btnCustomerLabelScanYN.IsChecked == true)
                    {
                        if (dgdOutwareSub.Items.Count > 0)
                        {
                            for (int i = 0; i < dgdOutwareSub.Items.Count; i++)
                            {
                                var outwareSubitem = dgdOutwareSub.Items[i] as Win_ord_OutWare_Scan_Sub_CodeView;

                                if (string.IsNullOrEmpty(outwareSubitem.CustomBoxID))
                                {
                                    outwareSubitem.CustomBoxID = txtScanData.Text;
                                    txtScanData.Text = string.Empty;
                                    return;
                                }
                                else
                                {
                                    MessageBox.Show("모든 행에 고객라벨을 전부 입력하였습니다.", "확인");
                                    txtScanData.Text = string.Empty;
                                    return;
                                }
                            };
                        }
                        else
                        {
                            MessageBox.Show("먼저 스캔하여 행을 추가하세요", "확인");
                            txtScanData.Text = string.Empty;
                            return;
                        }
                    }


                    if (tgnMoveByID.IsChecked == true)
                    {
                        //1. 일반 케이스 (사내라벨 스캔시)
                        if (txtScanData.Text.Trim().Length != 11)   // 삼주테크 바코드 길이 13자리로 확정
                        {
                            MessageBox.Show("잘못된 바코드 입니다.");
                            txtScanData.Text = string.Empty;
                            return;
                        }


                        if (txtScanData.Text.Substring(0, 1) == "P")
                        {
                            //2018.07.05 PACKINGID SCAN 과정 추가._허윤구.
                            // 지금 스캔된 녀석은 PACKING이다.
                            // 성공적으로 Packing List를 가져왔을 때,
                            //if (FindPackingLabelID(txtScanData.Text) == true)
                            //{
                            //    string InsideLabelID = string.Empty;

                            //    // 리스트 내부 LabelID를 돌면서 박스 스캔. > SUBGRID 추가(여러개)
                            //    for (int j = 0; j < LabelGroupList.Count; j++)
                            //    {
                            //        InsideLabelID = LabelGroupList[j].ToString();

                            //        FindBoxScanData(InsideLabelID);
                            //    }
                            //}
                        }
                        else
                        {
                            //부품식별표 박스ID 스캔 > SUBGRID 추가
                            FindBoxScanData(txtScanData.Text);
                        }

                        txtScanData.Text = string.Empty;
                    }

                    if (tgnMoveByQty.IsChecked == true)
                    {
                        // 바코드에 수량을 입력 → 숫자만 입력 가능하도록 유효성 검사
                        if (txtScanData.Text != "" && CheckConvertInt(txtScanData.Text))
                        {
                            // 수량 입력시 라벨 없이 입력됨
                            Win_ord_OutWare_Scan_Sub_CodeView label = new Win_ord_OutWare_Scan_Sub_CodeView();

                            int num = dgdOutwareSub.Items.Count + 1;
                            label.Num = num;
                            label.LabelID = "";
                            label.Spec = "";
                            label.Orderseq = orderSeq;
                            label.OutQty = stringFormatN0(txtScanData.Text);
                            dgdOutwareSub.Items.Add(label);

                            // 데이터 그리드 등록 후 바코드 초기화
                            txtScanData.Text = "";
                        }
                        else
                        {
                            MessageBox.Show("수량 등록에는 숫자만 입력 가능합니다.");
                        }
                    }
                }
                SumScanQty();
            }
            catch (Exception ee)
            {
                MessageBox.Show("오류지점 - txtScanData_KeyDown : " + ee.ToString());
            }
        }

        //PACKINGID SCAN 과정 > LABELID LIST 담기.
        private bool FindPackingLabelID(string PackingLabelID)
        {
            try
            {
                LabelGroupList.Clear();

                Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                sqlParameter.Clear();
                sqlParameter.Add("PackingLabelID", PackingLabelID);

                DataSet ds = DataStore.Instance.ProcedureToDataSet("xp_Outware_sPackingIDList", sqlParameter, false);

                if (ds != null && ds.Tables.Count > 0)
                {
                    DataTable dt = null;
                    dt = ds.Tables[0];

                    if (dt.Rows.Count == 0)
                    {
                        MessageBox.Show("PackingID를 포함하고 있는 LabelID를 찾을 수 없습니다.");
                        return false;
                    }
                    else
                    {
                        LabelGroupList.Clear();
                        LabelGroupInfo.Clear();
                        for (int i = 0; i < dt.Rows.Count; i++)
                        {
                            LabelGroupList.Add(dt.Rows[i]["InBoxID"].ToString());
                            LabelGroupInfo.Add(dt.Rows[i]["InBoxID"].ToString(), dt.Rows[i]["PackQty"].ToString());
                        }
                        return true;
                    }
                }
                else
                {
                    return false;
                }
            }
            catch (Exception ee)
            {
                MessageBox.Show("오류지점 - FindPackingLabelID : " + ee.ToString());
                return false;
            }
        }
      


        // 부품식별표 박스ID 스캔 > SUBGRID 추가
        private void FindBoxScanData(string ScanData)
        {
            try
            {
                ScanData = ScanData.Trim().ToUpper();
                LabelGroupList.Clear();

                Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                sqlParameter.Clear();
                sqlParameter.Add("BoxID", ScanData);

                DataSet ds = DataStore.Instance.ProcedureToDataSet("xp_Outware_sBoxIDOne", sqlParameter, false); ////// 2020.01.20 장가빈, wk_packing 테이블
                if (ds != null && ds.Tables.Count > 0)
                {
                    DataTable dt = null;
                    dt = ds.Tables[0];

                    if (dt.Rows.Count == 0)
                    {
                        MessageBox.Show("존재하지 않거나, 생산, 검사, 또는 수주등록 되지 않은 바코드 입니다.\r\n" +
                            "바코드 번호 :" + ScanData);
                        return;
                    }
                    else
                    {
                        DataRow DR = dt.Rows[0];

                        //세부작업 1. 스캔값에 대한 각종검증작업. > 리턴처리

                        /* if (DR["OutDate"].ToString() != string.Empty) // OutDate 컬럼에 값이 들어가 있으면 
                         {
                             MessageBox.Show(ScanData + " : 이미 출고된 바코드 번호입니다.");
                             return;
                         }*/

                        // key : labelID , value : outQty
                        Dictionary<string, double> dicCheck = new Dictionary<string, double>();
                        for (int i = 0; i < dgdOutwareSub.Items.Count; i++)     //이미 스캔한 바코드인지 체크, 
                        {
                            var OutSub = dgdOutwareSub.Items[i] as Win_ord_OutWare_Scan_Sub_CodeView;
                            string outSub_LabelID = OutSub.LabelID.Trim().ToUpper();
                            if (outSub_LabelID == ScanData)
                            {
                                MessageBox.Show((i + 1) + "줄에 이미 스캔된 바코드 입니다.");
                                return;
                            }
                            else
                            {
                                double outQty = ConvertDouble(OutSub.OutQty);
                                if (dicCheck.ContainsKey(outSub_LabelID))
                                    dicCheck[outSub_LabelID] += outQty;
                                else
                                    dicCheck.Add(outSub_LabelID, outQty);
                            }
                        }

                        string scanDate = DR["ScanDate"].ToString();




                        // 선출고ID의 존재여부
                        string foLotID = DR["FOLotID"].ToString();
                        if (string.IsNullOrEmpty(foLotID) == false)
                        {
                            string foLotDate = DR["FOStuffDate"].ToString();

                            DateTime dFo = string.IsNullOrEmpty(foLotDate) ? DateTime.Today : DateTime.Parse(DatePickerFormat(foLotDate));
                            DateTime dScan = string.IsNullOrEmpty(scanDate) ? DateTime.Today : DateTime.Parse(DatePickerFormat(scanDate));

                            // 출고할려고 그리드에 스캔한 정보 및 수량 확인 후 선출고건 수량에서 마이너스
                            // 그리드에 포함된 정보는 선출고건에서 제외시키기 위해
                            double foRemainQty = ConvertDouble(DR["FORemainQty"].ToString());
                            if (dicCheck.ContainsKey(foLotID))
                                foRemainQty = Math.Max(0, foRemainQty - dicCheck[foLotID]);

                            // 선출고ID와 스캔ID가 같지않고, 선출고일이 스캔일보다 이전이고, 선출고수량이 존재할때
                            if (foLotID.Trim().ToUpper() != ScanData && dFo.CompareTo(dScan) < 0 && foRemainQty > 0)
                            {
                                string foLotQty = stringFormatN0(foRemainQty);
                                string desc = "-------------------------------------\n" +
                                              "선출고시켜야할 출고건이 존재합니다.\n" +
                                              "-------------------------------------\n\n" +
                                              "[선출고 라벨 ID] : {0}\n" +
                                              "[선출고 라벨 수량] : {1} 개\n" +
                                              "[선출고 라벨 생성일] : {2}\n";

                                MessageBox.Show(string.Format(desc, foLotID, foLotQty, DatePickerFormat(foLotDate)));
                                return;
                            }
                        }


                        if (lib.returnNumStringZero(DR["StuffQty"].ToString()) == "0")
                        {
                            MessageBox.Show("출고가능한 수량이 없습니다.");
                            return;
                        }
                        else if (scanDate == string.Empty) //ScanDate 컬럼에 값이 비어있으면 / ScanDate는 PackDate와 같다
                        {
                            MessageBox.Show("생산이력이 없는 바코드 번호입니다.");
                            return;
                        }
                        else if (DR["inspectDate"].ToString() == string.Empty)   //wk_PackingCardList 테이블의 InspectDate / 검사일자가 비어있다면
                        {
                            MessageBox.Show("검사이력이 없는 바코드 번호입니다.");
                            return;
                        }

                        if (txtArticle_InGroupBox.Tag != null) //품명 텍스트 박스에 값이 있고,
                        {
                            if (txtArticle_InGroupBox.Tag.ToString() != DR["ArticleID"].ToString()) //품명 텍스트 박스에 기재된 품명과 받아온 품명이 다르면
                            {
                                System.Windows.MessageBox.Show("서로 다른 품명을 동시에 출고처리 할 수 없습니다.\r\n바코드를 확인해주세요.\n" +
                                    "바코드   품명 :" + DR["Article"].ToString() + " \r\n" +
                                    "출고희망 품명 :" + txtArticle_InGroupBox.Text + "","확인",MessageBoxButton.OK, MessageBoxImage.Error);
                                return;
                            }
                        }
                        if (txtKCustom.Tag != null) //거래처 텍스트 박스에 값이 있고, 
                        {
                            if (txtKCustom.Tag.ToString() != DR["CustomID"].ToString())         //거래처 텍스트 박스에 기재된 거래처와 받아온 거래처가 다르면
                            {
                                System.Windows.MessageBox.Show("서로 다른 거래처를 동시에 출고처리 할 수 없습니다. \r\n" +
                                    "바코드 거래처 :" + DR["CustomName"].ToString() + ". \r\n" +
                                    "출고 거래처 :" + txtKCustom.Text + ".","확인",MessageBoxButton.OK, MessageBoxImage.Error);
                                return;
                            }
                        }

                        //세부작업 2. 관리번호가 비어있다면 > 스캔항목을 통해 관리번호 자동유추 > 관리번호 값 입력.
                        if (txtOrderID.Text == string.Empty)
                        {
                            txtOrderID.Tag = DR["OrderID"].ToString();
                            txtOrderID.Text = DR["OrderID"].ToString();

                            // 관리번호 기반_ 항목 뿌리기 작업.
                            //OrderID_OtherSearch(txtOrderID.Text);

                            txtKCustom.Text = DR["CustomName"].ToString();
                            txtKCustom.Tag = DR["CustomID"].ToString();
                            txtOutCustom.Text = DR["OutCustom"].ToString();
                            txtOutCustom.Tag = DR["OutCustomID"].ToString();
                            txtBuyerName.Text = DR["CustomName"].ToString();
                            txtBuyerName.Tag = DR["CustomID"].ToString();
                            if (txtArticle_InGroupBox.Text == string.Empty) { txtArticle_InGroupBox.Text = DR["Article"].ToString(); }
                            if (txtArticle_InGroupBox.Tag == null)
                            {
                                txtArticle_InGroupBox.Tag = DR["ArticleID"].ToString();
                                txtArticleID_InGroupBox.Text = DR["ArticleID"].ToString();
                            }

                            if (txtArticleID_InGroupBox.Text == string.Empty)
                            {
                                txtArticleID_InGroupBox.Text = DR["ArticleID"].ToString();
                            }

                            txtBuyerArticleNo.Text = DR["BuyerArticleNo"].ToString();
                        }
                        //else
                        //{
                        //    txtOrderID.Tag = DR["OrderID"].ToString();
                        //    txtOrderID.Text = DR["OrderID"].ToString();

                        //    OrderID_OtherSearch(txtOrderID.Text);
                        //}

                        //세부작업 3. dgdOutwareSub에 ScanData Box DR 값 추가. (+ 1 Row)
                        var Win_ord_OutWare_Scan_Insert = new Win_ord_OutWare_Scan_Sub_CodeView()
                        {
                            LabelID = ScanData,                         //바코드 번호
                            OutQty = Lib.Instance.returnNumStringZero(DR["StuffQty"].ToString()),        //수량
                            OutRealQty = Lib.Instance.returnNumStringZero(DR["StuffQty"].ToString()),
                            UnitPrice = DR["UNITPRICE"].ToString(),     //단가
                            Orderseq = DR["OrderSeq"].ToString(),       //수주순서?
                            Amount = DR["Amount"].ToString(),           //금액
                            Vat_IND_YN = DR["VAT_IND_YN"].ToString(),   //부가세별도여부
                            LabelGubun = DR["labelGubun"].ToString(),   //라벨구분
                            Article = DR["Article"].ToString(),         //품명     

                            ArticleID = DR["ArticleID"].ToString(),     //품명ID         

                            DeleteYN = "Y",

                        };

                        lstCheck.Add(Win_ord_OutWare_Scan_Insert.LabelID, Win_ord_OutWare_Scan_Insert.OutQty);

                        //dgdOutwareSub.Items.Add(Win_ord_OutWare_Scan_Insert);
                        dgdOutwareSub.Items.Insert(0, Win_ord_OutWare_Scan_Insert); //2021-05-21 최근에 스캔 한 것이 위로 오게 수정

                    }
                }
            }
            catch (Exception ee)
            {
                MessageBox.Show("오류지점 - FindBoxScanData : " + ee.ToString());
            }
        }

        //서브 데이터 그리드 변경 이벤트
        private void dgdOutwareSub_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            try
            {
                if ((btnSave.Visibility == Visibility.Visible) && (btnCancel.Visibility == Visibility.Visible))
                {
                    //추가 / 수정 이벤트가 진행중인 경우,
                    var deleteControl = dgdOutwareSub.SelectedItem as Win_ord_OutWare_Scan_Sub_CodeView;
                    if (deleteControl != null)
                    {
                        deleteControl.DeleteYN = "Y";
                    }
                }
            }
            catch (Exception ee)
            {
                MessageBox.Show("오류지점 - dgdOutwareSub_SelectionChanged : " + ee.ToString());
            }
        }

        //서브 데이터 그리드 키다운 이벤트
        private void dgdOutwareSub_KeyDown(object sender, KeyEventArgs e)
        {
            try
            {
                if (e.Key == Key.Delete)
                {
                    //추가 / 수정 이벤트가 진행중인 경우,
                    if ((btnSave.Visibility == Visibility.Visible) && (btnCancel.Visibility == Visibility.Visible))
                    {
                        var OutwareSub = dgdOutwareSub.SelectedItem as Win_ord_OutWare_Scan_Sub_CodeView;
                        if (OutwareSub != null)
                        {
                            dgdOutwareSub.Items.Remove(OutwareSub);
                        }
                    }
                }
            }
            catch (Exception ee)
            {
                MessageBox.Show("오류지점 - dgdOutwareSub_KeyDown : " + ee.ToString());
            }
        }

        #endregion

        #region Research
        private void re_Search(int rowNum)
        {
            try
            {
                lstOutwarePrint.Clear();
                dgdOutware.Items.Clear();
                dgdOutwareSub.Items.Clear();

                FillGrid();

                if (dgdOutware.Items.Count > 0)
                    dgdOutware.SelectedIndex = rowNum;
                else
                    this.DataContext = null;
            }
            catch (Exception ee)
            {
                MessageBox.Show("오류지점 - re_Search : " + ee.ToString());
            }
        }

        #endregion

        #region 조회
        private void FillGrid()
        {
            dgdOutware.Items.Clear();
            dgdTotal.Items.Clear();

            try
            {
                DataSet ds = null;
                Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                sqlParameter.Add("ChkDate", chkOutwareDay.IsChecked == true ? 1 : 0);
                sqlParameter.Add("SDate", chkOutwareDay.IsChecked == true ? dtpFromDate.ToString().Substring(0, 10).Replace("-", "") : "");
                sqlParameter.Add("EDate", chkOutwareDay.IsChecked == true ? dtpToDate.ToString().Substring(0, 10).Replace("-", "") : "");

                // 거래처
                sqlParameter.Add("ChkCustomID", chkCustomIDSrh.IsChecked == true ? 1 : 0);
                sqlParameter.Add("CustomID", chkCustomIDSrh.IsChecked == true ? (txtCustomIDSrh.Tag != null ? txtCustomIDSrh.Tag.ToString() : "") : "");
                // 최종고객사
                sqlParameter.Add("ChkInCustom", chkInCustomer.IsChecked == true ? 1 : 0);
                sqlParameter.Add("InCustomID", chkInCustomer.IsChecked == true ? (txtInCustomer.Tag != null ? txtInCustomer.Tag.ToString() : "") : "");

                // 품명
                sqlParameter.Add("ChkArticleID", chkArticleIDSrh.IsChecked == true ? 1 : 0);
                sqlParameter.Add("ArticleID", chkArticleIDSrh.IsChecked == true ? (txtArticleIDSrh.Tag != null ? txtArticleIDSrh.Tag.ToString() : "") : "");

                // 품번
                sqlParameter.Add("ChkBuyerArticleNo", chkpfBuyerArticleNo.IsChecked == true ? 1 : 0);
                sqlParameter.Add("BuyerArticleNo", chkpfBuyerArticleNo.IsChecked == true ? (txtpfBuyerArticleNo.Tag != null ? txtpfBuyerArticleNo.Tag.ToString() : "") : "");  
          
                sqlParameter.Add("ChkOrder", rbnOrderNo.IsChecked == true ? 1 : 2);
                sqlParameter.Add("Order", txtRadioOptionNum.Text);                

                ds = DataStore.Instance.ProcedureToDataSet_LogWrite("xp_Outware_sOrder", sqlParameter, true, "R");

                if (ds != null && ds.Tables.Count > 0)
                {
                    DataTable dt = ds.Tables[0];

                    if (dt.Rows.Count == 0)
                    {
                        MessageBox.Show("조회결과가 없습니다. 검색조건을 확인해 주세요.");
                        return;
                    }
                    else
                    {
                        int i = 0;
                        int outqtySum = 0;
                        DataRowCollection drc = dt.Rows;
                        foreach (DataRow dr in drc)
                        {
                            i++;
                            var OutWardInfo = new Win_ord_OutWare_Scan_CodeView()
                            {
                                num = i,
                                OutwareID = dr["OutwareID"].ToString(),       //출고번호
                                OrderID = dr["OrderID"].ToString(),
                                OrderNo = dr["OrderNo"].ToString(),
                                OutClss = dr["OutClss"].ToString(),
                                OutDate = DateTypeHyphen(dr["Outdate"].ToString()),
                                OutTime = TimeTypeColon(dr["OutTime"].ToString()),
                                OutRoll = stringFormatN0(dr["OutRoll"]),
                                OutQty = stringFormatN0(dr["OutQty"]),
                                CustomID = dr["CustomID"].ToString(),
                                KCustom = dr["KCustom"].ToString(),
                                OutCustomID = dr["OutCustomID"].ToString(),
                                OutCustom = dr["OutCustom"].ToString(),
                                DvlyCustomID = dr["DvlyCustomID"].ToString(),
                                DvlyCustom = dr["DvlyCustom"].ToString(),

                                Article = dr["Article"].ToString(),
                                ArticleID = dr["ArticleID"].ToString(),
                                BuyerArticleNo = dr["BuyerArticleNo"].ToString(),
                                FromLocID = dr["FromLocID"].ToString(),
                                Remark = dr["Remark"].ToString(),

                                BuyerModel = dr["BuyerModel"].ToString(),

                                Buyer_Address1 = dr["Address1"].ToString(),
                                Buyer_Address2 = dr["Address2"].ToString(),
                                Buyer_Address3 = string.Empty,
                                Chief = dr["Chief"].ToString(),
                                CustomNo = dr["CustomNo"].ToString(),                                

                            };

                            outqtySum += (int)RemoveComma(dr["OutQty"].ToString(), true);

                            dgdOutware.Items.Add(OutWardInfo);
     
                        }

                        if(dgdOutware.Items.Count > 0)
                        {
                            var total = new Win_ord_OutWare_Scan_dgdTotal_CodeView
                            {
                                count = i,
                                totQty = stringFormatN0(outqtySum)
                            };

                            dgdTotal.Items.Add(total);
                        }
                      
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("오류지점 - FillGrid : " + ex.ToString());
            }
            finally
            {
                DataStore.Instance.CloseConnection();
            }
        }
        #endregion

        #region Sub조회
        private void FillGridSub(string OutwareID)
        {
            try
            {
                if (dgdOutwareSub.Items.Count > 0)
                {
                    dgdOutwareSub.Items.Clear();
                }

                Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                sqlParameter.Add("OutwareID", OutwareID);
                DataSet ds = DataStore.Instance.ProcedureToDataSet("xp_Outware_sOutwareSubGroup", sqlParameter, false);

                if (ds != null && ds.Tables.Count > 0)
                {
                    DataTable dt = null;
                    dt = ds.Tables[0];

                    if (dt.Rows.Count == 0)
                    {
                        return;
                    }
                    else
                    {
                        DataRowCollection drc = dt.Rows;
                        foreach (DataRow item in drc)
                        {
                            var Win_ord_OutWareSub_Scan_Insert = new Win_ord_OutWare_Scan_Sub_CodeView()
                            {
                                OutwareID = item["OutwareID"].ToString(),
                                OutSubSeq = item["OutSubSeq"].ToString(),
                                LabelID = item["LabelID"].ToString(),
                                LabelGubun = item["LabelGubun"].ToString(),
                                LabelGubunName = item["LabelGubunName"].ToString(),

                                OutQty = item["OutQty"].ToString(),
                                OutCnt = item["OutCnt"].ToString(),
                                OutRoll = item["OutRoll"].ToString(),
                                LotNo = item["LotNo"].ToString(),
                                Weight = item["Weight"].ToString(),

                                UnitPrice = item["UnitPrice"].ToString(),
                                Vat_IND_YN = item["Vat_IND_YN"].ToString(),
                                Orderseq = item["Orderseq"].ToString(),
                                Amount = item["Amount"].ToString(),
                                CustomBoxID = item["CustomBoxID"].ToString(),

                                FromLocID = item["FromLocID"].ToString(),
                                TOLocID = item["TOLocID"].ToString(),
                                UnitClss = item["UnitClss"].ToString(),
                                UnitClssName = item["UnitClssName"].ToString(),
                                ArticleID = item["ArticleID"].ToString(),
                                Article = item["Article"].ToString(),

                                OutClss = item["OutClss"].ToString(),
                                Gubun = item["Gubun"].ToString(),
                                DefectID = item["DefectID"].ToString(),
                                DefectName = item["DefectName"].ToString(),

                                SubRemark = item["SubRemark"].ToString(),
                                //Spec = item["Spec"].ToString(),

                                DeleteYN = "N",

                                OutRealQty = item["OutRealQty"].ToString()

                            };

                            Win_ord_OutWareSub_Scan_Insert.OutQty = lib.returnNumStringZero(Win_ord_OutWareSub_Scan_Insert.OutQty);
                            dgdOutwareSub.Items.Add(Win_ord_OutWareSub_Scan_Insert);
                        }
                    }
                }
            }
            catch (Exception ee)
            {
                MessageBox.Show("오류지점 - FillGridSub : " + ee.ToString());
            }
        }

        #endregion Sub조회

        #region 저장
        private bool SaveData(string strFlag)
        {
            bool flag = false;

            List<Procedure> Prolist = new List<Procedure>();
            List<Dictionary<string, object>> ListParameter = new List<Dictionary<string, object>>();

            try
            {
                if (CheckData())
                {               

                    #region 추가

                    if (strFlag == "I")
                    {
                        double cnt = 0;

                        Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                        sqlParameter.Clear();
                        sqlParameter.Add("OrderID", txtOrderID.Text);                  
                        sqlParameter.Add("CompanyID", MainWindow.CompanyID);    
                        sqlParameter.Add("OutSeq", "");
                        sqlParameter.Add("OutwareNo", "");
                        sqlParameter.Add("OutClss", cboOutClss.SelectedValue != null? cboOutClss.SelectedValue.ToString() : "");

                        sqlParameter.Add("CustomID", txtKCustom.Tag != null ? txtKCustom.Tag.ToString() : "");
                        sqlParameter.Add("BuyerDirectYN", "Y");
                        sqlParameter.Add("WorkID", "0001");                 
                        sqlParameter.Add("ExchRate", 0);
                        sqlParameter.Add("UnitPriceClss", "0");

                        sqlParameter.Add("InsStuffInYN", "N");             
                        sqlParameter.Add("OutcustomID", txtOutCustom.Tag != null ? txtOutCustom.Tag.ToString() : "");               
                        sqlParameter.Add("Outcustom", txtOutCustom.Text);
                        sqlParameter.Add("LossRate", 0);
                        sqlParameter.Add("LossQty", 0);

                        sqlParameter.Add("OutRoll", RemoveComma(txtOutRoll.Text,true));
                        sqlParameter.Add("OutQty", RemoveComma(txtOutQty.Text, true));
                        sqlParameter.Add("OutRealQty", ConvertDouble(txtOutQty.Text.Replace(",", ""))); 
                        sqlParameter.Add("OutDate", dtpOutDate.Text.ToString().Substring(0, 10).Replace("-", ""));
                        sqlParameter.Add("ResultDate", dtpOutDate.Text.ToString().Substring(0, 10).Replace("-", ""));

                        sqlParameter.Add("Remark", txtRemark.Text.Equals("") ? "사무실에서 출고" : txtRemark.Text);
                        sqlParameter.Add("OutType", "3");                
                        sqlParameter.Add("OutSubType", "");              
                        sqlParameter.Add("Amount", 0);                   
                        sqlParameter.Add("VatAmount", 0);                

                        sqlParameter.Add("VatINDYN", "Y");               
                        sqlParameter.Add("FromLocID", cboFromLoc.SelectedValue != null ? cboFromLoc.SelectedValue.ToString() : "");
                        sqlParameter.Add("ToLocID", cboToLoc.SelectedValue != null ? cboToLoc.SelectedValue.ToString() : "");
                        sqlParameter.Add("UnitClss", 0);                       

                        sqlParameter.Add("DvlyCustomID", txtBuyerName.Tag != null ?  txtBuyerName.Tag.ToString() :""); 
                        sqlParameter.Add("UserID", MainWindow.CurrentUser);
                 

                        Procedure pro1 = new Procedure();
                        pro1.Name = "xp_Outware_iOutware";
                        pro1.OutputUseYN = "Y";
                        pro1.OutputName = "OutwareNo";      //OutwareNo = OutwareID
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
                                if (kv.key == "OutwareNo")
                                {
                                    sGetID = kv.value;

                                    GetKey = kv.value;

                                    Prolist.RemoveAt(0);
                                    ListParameter.Clear();
                                }
                            }
                        }
                        else
                        {
                            MessageBox.Show("[저장실패]\r\n" + list_Result[0].value.ToString());
                        }


                        //sub그리드 아이템 수만큼 반복되어야 하므로
                        for (int i = 0; i < dgdOutwareSub.Items.Count; i++)
                        {
                            var OutwareSub = dgdOutwareSub.Items[i] as Win_ord_OutWare_Scan_Sub_CodeView;

                            sqlParameter = new Dictionary<string, object>();
                            sqlParameter.Clear();
                            sqlParameter.Add("OutwareID", GetKey);
                            sqlParameter.Add("OrderID", txtOrderID.Text);
                            sqlParameter.Add("OutSeq", "");
                            sqlParameter.Add("OutSubSeq", i + 1);
                            sqlParameter.Add("OrderSeq", OutwareSub.Orderseq);

                            sqlParameter.Add("LineSeq", 0);
                            sqlParameter.Add("LineSubSeq", 0);
                            sqlParameter.Add("RollSeq", i);
                            sqlParameter.Add("LabelID", OutwareSub.LabelID);
                            sqlParameter.Add("LabelGubun", "2");        //박스라벨출고는 2번

                            sqlParameter.Add("LotNo", "0");
                            sqlParameter.Add("Gubun", "");             
                            sqlParameter.Add("StuffQty", 0);
                            sqlParameter.Add("OutQty", RemoveComma(OutwareSub.OutQty, true));
                            sqlParameter.Add("OutRoll", 1); 

                            sqlParameter.Add("UnitPrice", RemoveComma(OutwareSub.UnitPrice, true));

                        
                            sqlParameter.Add("CustomBoxID", OutwareSub.CustomBoxID);
                            sqlParameter.Add("DefectID", "");            
                            sqlParameter.Add("BoxID", OutwareSub.LabelID);
                            //sqlParameter.Add("ArticleID", OutwareSub.ArticleID);
                            sqlParameter.Add("ArticleID", txtArticleID_InGroupBox.Text != null ? txtArticleID_InGroupBox.Text : "");
                            sqlParameter.Add("SubRemark", "사무실에서 출고(출하스캔)");
                            //sqlParameter.Add("Spec", OutwareSub.Spec);

                            sqlParameter.Add("UserID", MainWindow.CurrentUser);


                            Procedure pro2 = new Procedure();
                            pro2.Name = "xp_Outware_iOutwareSub";
                            pro2.OutputUseYN = "N";
                            pro2.OutputName = "REQ_ID";
                            pro2.OutputLength = "10";

                            //cnt += (Double.Parse(OutwareSub.OutQty.Replace(",", "")) * Double.Parse(OutwareSub.UnitPrice.Replace(",", "")));

                            Prolist.Add(pro2);
                            ListParameter.Add(sqlParameter);

                        }
          
                    }

                    #endregion   추가

                    #region 수정

                    else if (strFlag == "U")
                    {    

                        Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                        sqlParameter.Clear();
                        sqlParameter.Add("OutwareID", txtOutwareID.Text);
                        sqlParameter.Add("OrderID", txtOrderID.Text);
                        sqlParameter.Add("CompanyID", MainWindow.CompanyID);
                        sqlParameter.Add("OutClss", cboOutClss.SelectedValue.ToString());

                        sqlParameter.Add("CustomID", txtKCustom.Tag != null ? txtKCustom.Tag.ToString() : "");
                        sqlParameter.Add("BuyerDirectYN", "Y");
                        sqlParameter.Add("WorkID", "0001");
                        sqlParameter.Add("ExchRate", 0);
                        sqlParameter.Add("UnitPriceClss", "0");

                        sqlParameter.Add("InsStuffInYN", "N");
                        sqlParameter.Add("OutcustomID", txtOutCustom.Tag != null ? txtOutCustom.Tag.ToString() : ""); //20210526
                        sqlParameter.Add("Outcustom", txtOutCustom.Text);
                        sqlParameter.Add("LossRate", 0);
                        sqlParameter.Add("LossQty", 0);

                        sqlParameter.Add("OutRoll", RemoveComma(txtOutRoll.Text,true));
                        sqlParameter.Add("OutQty", RemoveComma(txtOutQty.Text, true));
                        sqlParameter.Add("OutRealQty", RemoveComma(txtOutQty.Text, true));
                        sqlParameter.Add("OutDate", !IsDatePickerNull(dtpOutDate) ? ConvertDate(dtpOutDate) : "");
                        sqlParameter.Add("ResultDate", !IsDatePickerNull(dtpOutDate) ? ConvertDate(dtpOutDate) : "");

                        sqlParameter.Add("Remark", string.IsNullOrEmpty(txtRemark.Text) ? "사무실에서 출고" : txtRemark.Text);
                        sqlParameter.Add("OutType", "3");
                        sqlParameter.Add("OutSubType", "");
                        sqlParameter.Add("Amount", 0);
                        sqlParameter.Add("VatAmount", 0);

                        sqlParameter.Add("VatINDYN", 'Y');
                        sqlParameter.Add("FromLocID", cboFromLoc.SelectedValue != null ? cboFromLoc.SelectedValue.ToString() :"");
                        sqlParameter.Add("ToLocID", cboToLoc.SelectedValue != null ? cboToLoc.SelectedValue.ToString() :"");
                        sqlParameter.Add("UnitClss", 0);

                        sqlParameter.Add("DvlyCustomID", txtBuyerName.Tag != null ?  txtBuyerName.Tag.ToString() : "" ); //20210526
                        sqlParameter.Add("UserID", MainWindow.CurrentUser);
                        

                        Procedure pro3 = new Procedure();
                        pro3.Name = "xp_Outware_uOutware";
                        pro3.OutputUseYN = "N";
                        pro3.OutputName = "";
                        pro3.OutputLength = "15";

                        Prolist.Add(pro3);
                        ListParameter.Add(sqlParameter);

                        // 모든것을 삭제한 후에, 새롭게 추가
                        sqlParameter = new Dictionary<string, object>();
                        sqlParameter.Clear();
                        sqlParameter.Add("OutWareID", txtOutwareID.Text);
                        //sqlParameter.Add("Seq", "");

                        Procedure pro4 = new Procedure();
                        pro4.Name = "xp_Outware_dOutwareSubAll";
                        pro4.OutputUseYN = "N";
                        pro4.OutputName = "OrderID";
                        pro4.OutputLength = "10";

                        Prolist.Add(pro4);
                        ListParameter.Add(sqlParameter);


                        //sub그리드 아이템 수만큼 반복되어야 하므로 
                        for (int i = 0; i < dgdOutwareSub.Items.Count; i++)
                        {
                            var OutwareSub = dgdOutwareSub.Items[i] as Win_ord_OutWare_Scan_Sub_CodeView;

                            sqlParameter = new Dictionary<string, object>();
                            sqlParameter.Clear();
                            sqlParameter.Add("OutwareID", txtOutwareID.Text);
                            sqlParameter.Add("OrderID", txtOrderID.Text);
                            sqlParameter.Add("OutSeq", "");
                            sqlParameter.Add("OutSubSeq", i + 1);
                            sqlParameter.Add("OrderSeq", OutwareSub.Orderseq);

                            sqlParameter.Add("LineSeq", 0);
                            sqlParameter.Add("LineSubSeq", 0);
                            sqlParameter.Add("RollSeq", i);
                            sqlParameter.Add("LabelID", OutwareSub.LabelID);
                            sqlParameter.Add("LabelGubun", "2");        

                            sqlParameter.Add("LotNo", "0");
                            sqlParameter.Add("Gubun", "");              
                            sqlParameter.Add("StuffQty", 0);
                            sqlParameter.Add("OutQty", RemoveComma(OutwareSub.OutQty,true));
                            sqlParameter.Add("OutRoll", 1); // 하나당 박스 1개로 처리 하니, 1로 저장한다고 함

                            sqlParameter.Add("UnitPrice", RemoveComma(OutwareSub.UnitPrice, true));
                            sqlParameter.Add("CustomBoxID", OutwareSub.CustomBoxID);
                            sqlParameter.Add("DefectID", "");          
                            sqlParameter.Add("BoxID", OutwareSub.LabelID);
                            sqlParameter.Add("ArticleID", txtArticleID_InGroupBox.Text != null ? txtArticleID_InGroupBox.Text : "");
                            sqlParameter.Add("SubRemark", "사무실에서 출고(출하스캔)");

                            sqlParameter.Add("UserID", MainWindow.CurrentUser);

                            Procedure pro2 = new Procedure();
                            pro2.Name = "xp_Outware_iOutwareSub";
                            pro2.OutputUseYN = "N";
                            pro2.OutputName = "REQ_ID";
                            pro2.OutputLength = "10";

                            //cnt += (Double.Parse(OutwareSub.OutQty.Replace(",", "")) * Double.Parse(OutwareSub.UnitPrice.Replace(",", "")));

                            Prolist.Add(pro2);
                            ListParameter.Add(sqlParameter);
                        }

                        //ListParameter[0]["Amount"] = cnt.ToString();
                        //ListParameter[0]["VatAmount"] = (cnt * 0.1).ToString();
                    }

                    #endregion 수정

                    string[] Confirm = new string[2];
                    Confirm = DataStore.Instance.ExecuteAllProcedureOutputNew_NewLog(Prolist, ListParameter, "U");
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
                else
                {
                    
                    txtScanData.Focus();
                    flag = false;
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show("오류지점 - SaveData : " + ex.ToString());
            }
            finally
            {
                DataStore.Instance.CloseConnection();
                tmpRestQty = 0;
            }

            return flag;
        }

        #endregion 저장

        #region 데이터 체크
        // 그룹박스 데이터 기입체크
        private bool CheckData()
        {
            try
            {
                if (txtOrderID.Text == "")
                {
                    MessageBox.Show("관리번호를 반드시 입력하세요.");
                    return false;
                }
              
                if (cboOutClss.SelectedIndex < 0)
                {
                    MessageBox.Show("출고구분은 반드시 선택하세요.");
                    return false;
                }
                if (cboFromLoc.SelectedIndex < 0)
                {
                    MessageBox.Show("전 창고는 반드시 선택하세요.");
                    return false;
                }
                if (dgdOutwareSub.Items.Count == 0)
                {
                    MessageBox.Show("스캔된 라벨 정보가 없습니다.");
                    return false;
                }
             
                // 박스에 담긴 수량 체크
                //if (tgnMoveByID.IsChecked == true)
                //{
                //    for (int i = 0; i < dgdOutwareSub.Items.Count; i++)
                //    {
                //        var OutwareSub = dgdOutwareSub.Items[i] as Win_ord_OutWare_Scan_Sub_CodeView;
                //        foreach (KeyValuePair<string, object> kvp in lstCheck)
                //        {
                //            if (OutwareSub.LabelID == kvp.Key)
                //            {
                //                if (Convert.ToInt32(OutwareSub.OutQty) > Convert.ToInt32(kvp.Value))
                //                {
                //                    MessageBox.Show("박스ID(" + OutwareSub.LabelID + ")에 남은 수량을 초과하셨습니다");
                //                    return false;
                //                }
                //            }
                //        }
                //    }
                //}
              

                return true;
            }
            catch (Exception ee)
            {
                MessageBox.Show("오류지점 - CheckData : " + ee.ToString());
                return false;
            }
        }
        #endregion

        #region 삭제
        private bool DeleteData(string OutwareID)
        {
            bool flag = false;

            try
            {
                Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                sqlParameter.Clear();
                sqlParameter.Add("OutwareID", OutwareID);

                string[] result = DataStore.Instance.ExecuteProcedure_NewLog("xp_Outware_dOutware", sqlParameter, "D");

                if (result[0].Equals("success"))
                {
                    //MessageBox.Show("성공 *^^*");
                    flag = true;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("오류지점 - DeleteData : " + ex.ToString());
            }
            finally
            {
                DataStore.Instance.CloseConnection();
            }



            return flag;
        }


        #endregion 삭제

        //라벨스캔 토글버튼 클릭
        private void btnCustomerLabelScanYN_Click(object sender, RoutedEventArgs e)
        {
            //안쓸 듯
        }

        //서브 데이터 그리드 삭제컬럼 버튼 클릭
        private void dgdOutwareSub_btnDelete_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var OutwareSub = dgdOutwareSub.SelectedItem as Win_ord_OutWare_Scan_Sub_CodeView;
                if (OutwareSub != null)
                {
                    dgdOutwareSub.Items.Remove(OutwareSub);

                    // lstCheck에서 해당 항목 삭제
                    if (lstCheck.ContainsKey(OutwareSub.LabelID))
                    {
                        lstCheck.Remove(OutwareSub.LabelID);
                    }
                }

                SumScanQty();
            }
            catch (Exception ee)
            {
                MessageBox.Show("오류지점 - dgdOutwareSub_btnDelete_Click : " + ee.ToString());
            }
        }

        // 출고지시 기반_ 항목 뿌리기 작업.
        private void OrderID_OtherSearch(string orderID, string OutwareReqID)
        {
            try
            {
                Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                sqlParameter.Clear();
                sqlParameter.Add("OrderID", orderID);
                sqlParameter.Add("OutwareReqID", OutwareReqID);
                sqlParameter.Add("Article", replyArticle);

                DataSet ds = DataStore.Instance.ProcedureToDataSet("xp_Outware_sOrderOne", sqlParameter, false);

                if (ds != null && ds.Tables.Count > 0)
                {
                    DataTable dt = null;
                    dt = ds.Tables[0];

                    if (dt.Rows.Count == 0)
                    {
                        return;
                    }
                    else
                    {
                        DataRow DR = dt.Rows[0];

                        // 거래처
                        txtKCustom.Text = DR["KCustom"].ToString();
                        txtKCustom.Tag = DR["CustomID"].ToString();
                        // 최종고객사
                        txtOutCustom.Text = DR["KInCustom"].ToString();
                        txtOutCustom.Tag = DR["InCustomID"].ToString();
                        // 납품장소
                        txtBuyerName.Text = DR["KInCustom"].ToString();
                        txtBuyerName.Tag = DR["InCustomID"].ToString();

                        if (txtArticle_InGroupBox.Text == string.Empty) { txtArticle_InGroupBox.Text = DR["Article"].ToString(); }
                        if (txtArticle_InGroupBox.Tag == null)
                        {
                            txtArticle_InGroupBox.Tag = DR["ArticleID"].ToString();
                            txtArticleID_InGroupBox.Text = DR["ArticleID"].ToString();
                        }

                        if (txtArticleID_InGroupBox.Text == string.Empty)
                            txtArticleID_InGroupBox.Text = DR["ArticleID"].ToString();

                        txtBuyerArticleNo.Text = DR["BuyerArticleNo"].ToString();
                        orderSeq = DR["OrderSeq"].ToString();
                    
                    }
                }

                /*if (txtOrderID.Text != string.Empty)
                    FillOrderID(txtOrderID.Text, (txtKCustom.Tag == null ? "" : txtKCustom.Tag.ToString()));*/
            }
            catch (Exception ee)
            {
                MessageBox.Show("오류지점 - OrderID_OtherSearch : " + ee.ToString());
            }
        }

        //관리번호기반 품명뿌려주기
        private void FillOrderID(String OrderID, string CustomID)
        {
            try
            {
                double OutQty = 0;
                int OutRoll = 0;

                Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                sqlParameter.Clear();
                sqlParameter.Add("OrderID", OrderID);
                //sqlParameter.Add("CustomID", CustomID); //2020.12.02 정승학 수정

                DataSet ds = DataStore.Instance.ProcedureToDataSet("xp_OutWare_sGrid", sqlParameter, false);

                if (ds != null && ds.Tables.Count > 0)
                {
                    DataTable dt = null;
                    dt = ds.Tables[0];

                    if (dt.Rows.Count == 0)
                    {
                        return;
                    }
                    else
                    {
                        DataRowCollection drc = dt.Rows;
                        foreach (DataRow item in drc)
                        {
                            var Win_ord_OutWare_Scan_Insert = new Win_ord_OutWare_Scan_Sub_CodeView()
                            {

                                OutQty = item["OutQty"].ToString(),        //수량
                                LabelID = item["LabelID"].ToString(),        //라벨(빈값)
                                UnitPrice = item["UnitPrice"].ToString(),     //단가
                                Orderseq = item["OrderSeq"].ToString(),       //수주순서?                                
                                //Amount = item["Amount"].ToString(),           //금액
                                Vat_IND_YN = item["VAT_IND_YN"].ToString(),    //부가세별도여부
                                LabelGubun = item["labelGubun"].ToString(),    //라벨구분
                                Article = item["Article"].ToString(),          //품명           
                                ArticleID = item["ArticleID"].ToString(),  //품명코드
                                SubRemark = item["SubRemark"].ToString(),   //품명비고
                                Spec = item["Spec"].ToString(),       //규격

                                DeleteYN = "Y",
                            };
                            //Win_ord_OutWare_Scan_Insert.UnitPrice = lib.returnNumStringZero(Win_ord_OutWare_Scan_Insert.UnitPrice);
                            Win_ord_OutWare_Scan_Insert.OutQty = lib.returnNumStringZero(Win_ord_OutWare_Scan_Insert.OutQty);
                            //Win_ord_OutWare_Scan_Insert.Amount = lib.returnNumStringZero(Win_ord_OutWare_Scan_Insert.Amount).ToString();

                            dgdOutwareSub.Items.Add(Win_ord_OutWare_Scan_Insert);

                            OutQty += lib.returnDouble(Win_ord_OutWare_Scan_Insert.OutQty);

                            txtOutQty.Text = lib.returnNumStringZero(OutQty.ToString());

                            OutRoll = dgdOutwareSub.Items.Count;
                            txtOutRoll.Text = stringFormatN0(OutRoll);

                        }

                        //SumOutQty();
                        //SumAmout();

                    }
                }
            }
            catch (Exception ee)
            {
                MessageBox.Show("오류지점 - FillOrderID : " + ee.ToString());
            }
            finally
            {
                DataStore.Instance.CloseConnection();
            }
        }

        //추가, 수정일 때 
        private void CanBtnControl()
        {
            btnAdd.IsEnabled = false;               //추가
            btnUpdate.IsEnabled = false;            //수정
            btnDelete.IsEnabled = false;            //삭제
            btnClose.IsEnabled = true;              //닫기
            btnSearch.IsEnabled = false;            //검색
            btnSave.Visibility = Visibility.Visible;             //저장
            btnCancel.Visibility = Visibility.Visible;             //취소
            btnExcel.IsEnabled = false;             //엑셀
            btnPrint.IsEnabled = false;             //인쇄

            txtBuyerModel.IsHitTestVisible = false;  //차종은 땡겨오니까
            txtOutwareID.IsHitTestVisible = false;   //출고번호는 자동으로 생성되니까
            txtScanData.IsEnabled = true;           //바코드 스캔
            EventLabel.Visibility = Visibility.Visible; //자료입력중
            grbOutwareDetailBox.IsEnabled = true;       //DataContext Box
            dgdOutware.IsHitTestVisible = false;        //데이터그리드 클릭 안되게

            // 토글버튼
            tgnMoveByID.IsHitTestVisible = true;
            tgnMoveByQty.IsHitTestVisible = true;

            
        }
        //저장, 취소일 때
        private void CantBtnControl()
        {
            btnAdd.IsEnabled = true;               //추가
            btnUpdate.IsEnabled = true;            //수정
            btnDelete.IsEnabled = true;            //삭제
            btnClose.IsEnabled = true;             //닫기
            btnSearch.IsEnabled = true;            //검색
            btnSave.Visibility = Visibility.Hidden;             //저장
            btnCancel.Visibility = Visibility.Hidden;             //취소
            btnExcel.IsEnabled = true;             //엑셀
            btnPrint.IsEnabled = true;             //인쇄

            txtBuyerModel.IsHitTestVisible = false;  //차종은 땡겨오니까
            txtScanData.IsEnabled = false;         //바코드 스캔
            EventLabel.Visibility = Visibility.Hidden; //자료입력중
            grbOutwareDetailBox.IsEnabled = false;       //DataContext Box
            dgdOutware.IsHitTestVisible = true;        //데이터그리드 클릭되게

            // 토글버튼
            tgnMoveByID.IsHitTestVisible = false;
            tgnMoveByQty.IsHitTestVisible = false;
        }

        private void TextBoxClear()
        {
            txtOrderID.Text = string.Empty;
            txtArticleID_InGroupBox.Text = string.Empty;
            txtArticle_InGroupBox.Text = string.Empty;
            txtArticle_InGroupBox.Tag = null;
            txtBuyerArticleNo.Text = string.Empty;
            cboOutClss.SelectedIndex = 0;
            txtBuyerModel.Text = string.Empty;
            txtOutwareID.Text = string.Empty;
            txtOutRoll.Text = string.Empty;
            txtOutQty.Text = string.Empty;
            cboFromLoc.SelectedIndex = 0;
            cboToLoc.SelectedIndex = 0;
            txtKCustom.Text = string.Empty;
            txtKCustom.Tag = null;
            txtBuyerName.Text = string.Empty;
            txtBuyerName.Tag = null;
            txtRemark.Text = string.Empty;
            txtOutCustom.Text = string.Empty;
            txtScanData.Text = string.Empty;

        }

        private void SumScanQty()
        {
            try
            {
                int OutRoll = 0;
                double OutQty = 0;

                OutRoll = dgdOutwareSub.Items.Count;

                for (int i = 0; i < dgdOutwareSub.Items.Count; i++)
                {
                    var label = dgdOutwareSub.Items[i] as Win_ord_OutWare_Scan_Sub_CodeView;
                    if (label.OutQty != null)
                        OutQty += ConvertDouble(label.OutQty.ToString());
                }

                txtOutRoll.Text = dgdOutwareSub.Items.Count.ToString();
                txtOutQty.Text = stringFormatN0(OutQty);
            }
            catch (Exception ee)
            {
                MessageBox.Show("오류지점 - SumScanQty : " + ee.ToString());
            }
        }

        // 천자리 콤마, 소수점 버리기
        private string stringFormatN0(object obj)
        {
            return string.Format("{0:N0}", obj);
        }

        //더블로 형식 변환
        private double ConvertDouble(string str)
        {
            double result = 0;
            double chkDouble = 0;

            try
            {
                if (!str.Trim().Equals(""))
                {
                    str = str.Trim().Replace(",", "");

                    if (double.TryParse(str, out chkDouble) == true)
                    {
                        result = double.Parse(str);
                    }
                }
                return result;
            }
            catch (Exception ee)
            {
                MessageBox.Show("오류지점 - ConvertDouble : " + ee.ToString());
                return result;
            }
        }

        // 데이터피커 포맷으로 변경
        private string DatePickerFormat(string str)
        {
            string result = "";

            try
            {
                if (str.Length == 8)
                {
                    if (!str.Trim().Equals(""))
                    {
                        result = str.Substring(0, 4) + "-" + str.Substring(4, 2) + "-" + str.Substring(6, 2);
                    }
                }

                return result;
            }
            catch (Exception ee)
            {
                MessageBox.Show("오류지점 - DatePickerFormat : " + ee.ToString());
                return result;
            }
        }

        //출고지시번호 숫자만 입력
        private void txtOutwareReqID_PreviewTextInput(object sender, TextCompositionEventArgs e)
        {
            try
            {
                lib.CheckIsNumeric((TextBox)sender, e);
            }
            catch (Exception ee)
            {
                MessageBox.Show("오류지점 - txtOutwareReqID_PreviewTextInput : " + ee.ToString());
            }
        }

        //관리번호 숫자만 입력
        private void txtOrderID_PreviewTextInput(object sender, TextCompositionEventArgs e)
        {
            try
            {
                lib.CheckIsNumeric((TextBox)sender, e);
            }
            catch (Exception ee)
            {
                MessageBox.Show("오류지점 - txtOrderID_PreviewTextInput : " + ee.ToString());
            }
        }

        //박스에 숫자만 입력
        private void txtOutRoll_PreviewTextInput(object sender, TextCompositionEventArgs e)
        {
            try
            {
                lib.CheckIsNumeric((TextBox)sender, e);
            }
            catch (Exception ee)
            {
                MessageBox.Show("오류지점 - txtOutRoll_PreviewTextInput : " + ee.ToString());
            }
        }

        //수량에 숫자만 입력
        private void txtOutQty_PreviewTextInput(object sender, TextCompositionEventArgs e)
        {
            try
            {
                lib.CheckIsNumeric((TextBox)sender, e);
            }
            catch (Exception ee)
            {
                MessageBox.Show("오류지점 - txtOutQty_PreviewTextInput : " + ee.ToString());
            }
        }

        //검색조건 - 관리번호에 숫자만 입력
        private void txtRadioOptionNum_PreviewTextInput(object sender, TextCompositionEventArgs e)
        {
            try
            {
                lib.CheckIsNumeric((TextBox)sender, e);
            }
            catch (Exception ee)
            {
                MessageBox.Show("오류지점 - txtRadioOptionNum_PreviewTextInput : " + ee.ToString());
            }
        }


        private void chkReq_Click(object sender, RoutedEventArgs e)
        {
            CheckBox chkSender = sender as CheckBox;
            var Outware = chkSender.DataContext as Win_ord_OutWare_Scan_CodeView;

            if (Outware != null)
            {
                if (chkSender.IsChecked == true)
                {
                    Outware.Chk = true;

                    if (lstOutwarePrint.Contains(Outware) == false)
                        lstOutwarePrint.Add(Outware);
                }
                else
                {
                    Outware.Chk = false;

                    if (lstOutwarePrint.Contains(Outware) == true)
                        lstOutwarePrint.Remove(Outware);
                }

            }
        }

        private void txtQty_KeyDown(object sender, KeyEventArgs e)
        {
            try
            {
                //박스포장량 검증 추가
                TextBox textbox = sender as TextBox;
                if(textbox != null  && decimal.TryParse(textbox.Text, out decimal OutQty))
                {
                    var selectedRow = dgdOutwareSub.CurrentItem as Win_ord_OutWare_Scan_Sub_CodeView;
                    if(selectedRow != null)
                    {
                        if (!string.IsNullOrEmpty(selectedRow.LabelID))
                        {
                            int boxQty = lib.GetBoxQty(selectedRow.LabelID);
                            if(boxQty > 0 && OutQty > boxQty)
                            {
                                e.Handled = true;
                                textbox.Text = stringFormatN0(boxQty);
                                MessageBox.Show($"입력된 수량이 남은 수량 ({boxQty}) 보다 많이 입력 되었습니다.", "확인");
                                return;
                            }
                        }
                    }
                }
                else
                {
                    SumColorQty();
                }
            }
            catch (Exception ee)
            {
                MessageBox.Show("오류지점 - txtQty_KeyDown : " + ee.ToString());
            }

        
        }



        private void dgdOutwareSubRequest_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            // 추가 상태로 들어와야 하고
            if (EventStatus == true)
            {
                var ViewReceiver = dgdOutwareSub.CurrentCell.Item as Win_ord_OutWare_Scan_Sub_CodeView;   //dgdOutRequest.SelectedItem as Win_out_OutwareReq_U_View;
                if (ViewReceiver != null)
                {
                    string eventer = ((DataGridCell)sender).Column.Header.ToString();

                    if (eventer == "수량")//(((eventer == "수량")) || (ButtonTag == "2") && (eventer == "Comments"))
                    {
                        List<TextBox> list = new List<TextBox>();
                        lib.FindChildGroup<TextBox>(dgdOutwareSub, "txtQty", ref list);
                        int target = dgdOutwareSub.Items.IndexOf(dgdOutwareSub.CurrentCell.Item);  //dgdOutRequest.SelectedIndex;
                        TextBox TextBoxComments = list[target];

                        TextBoxComments.IsReadOnly = false;
                        TextBoxComments.Focus();

                        Dispatcher.BeginInvoke((ThreadStart)delegate
                        {
                            TextBoxComments.Focus();
                        });
                    }
                }
            }
        }

        private void dgdOutwareSubRequest_KeyDown(object sender, KeyEventArgs e)
        {

        }

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

        //박스선택을 만들어보자
        //private void BtnAddbox_Click(object sender, RoutedEventArgs e)
        //{

        //    if (CheckdateBox())
        //    {
        //        Win_pop_Box_LotNo BoxID = new Win_pop_Box_LotNo(lstBoxID);

        //        BoxID.ShowDialog();

        //        //   // 중복되는 라벨이 있을 경우 메시지 띄워주기 위한 변수
        //        string Msg = "";

        //        if (BoxID.DialogResult == true)
        //        {
        //            int count = 0;

        //            string InsideLabelID = string.Empty; // 추가해봄

        //            for (int i = 0; i < BoxID.lstBoxID.Count; i++)
        //            {
        //                var main = BoxID.lstBoxID[i] as Win_ord_OutWare_Scan_Sub_CodeView;

        //                //txtScanData.Text = main.LabelID;

        //                try
        //                {
        //                    InsideLabelID = main.LabelID.ToString(); // 박스선택으로 받아온 라벨들을 저장
        //                    FindBoxScanData(InsideLabelID);                 //저장된 라벨로 바코드 칸 엔터신공
        //                }
        //                catch (Exception ex)
        //                {
        //                    MessageBox.Show(ex.ToString());
        //                }

        //                txtScanData.Text = string.Empty;

        //                SumScanQty();

        //            }
        //            if (Msg.Length > 0)
        //            {
        //                Msg += "위의 라벨은 이미 등록되어 있습니다.";
        //                if (count != 0) { Msg += "\r(위의 라벨을 제외하고 추가되었습니다.)"; }
        //                MessageBox.Show(Msg);
        //            }

        //        }

        //        setNumSubDgd();
        //    }



        //}


        private bool CheckdateBox()
        {
            bool flag = true;

            if (txtOrderID.Text.Trim().Equals(""))
            {
                MessageBox.Show("관리번호가 입력되지 않았습니다.");
                flag = false;
                return flag;
            }


            return flag;
        }

        // 서브그리드 삭제 시 → Num 재정렬
        private void setNumSubDgd()
        {
            int index = 0;
            for (int i = 0; i < dgdOutwareSub.Items.Count; i++)
            {
                var Sub = dgdOutwareSub.Items[i] as Win_ord_OutWare_Scan_Sub_CodeView;
                if (Sub != null)
                {
                    index++;
                    Sub.Num = index;
                }
            }
        }

        private void ButtonDataGridSubRowAdd_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                SubRowAdd();
            }
            catch (Exception ee)
            {
                MessageBox.Show("오류지점 - ButtonDataGridSubRowAdd_Click : " + ee.ToString());
            }
        }

        private void ButtonDataGridSubRowDel_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                SubRowDel();
                //SumOutQty();
                //SumAmout();
            }
            catch (Exception ee)
            {
                MessageBox.Show("오류지점 - ButtonDataGridSubRowDel_Click : " + ee.ToString());
            }
        }

        //서브 그리드 추가
        private void SubRowAdd()
        {
            try
            {
                int index = dgdOutwareSub.Items.Count;

                var WOOSSC = new Win_ord_OutWare_Scan_Sub_CodeView()
                {
                    OutwareID = "",
                    OutSubSeq = "",
                    LabelID = "",
                    LabelGubun = "",
                    LabelGubunName = "",

                    OutQty = "",
                    OutCnt = "",
                    OutRoll = "",
                    LotNo = "",
                    Weight = "",

                    UnitPrice = "",
                    Vat_IND_YN = "",
                    Orderseq = "",
                    Amount = "",
                    CustomBoxID = "",

                    FromLocID = "",
                    TOLocID = "",
                    UnitClss = "",
                    ArticleID = "",
                    Article = "",

                    OutClss = "",
                    Gubun = "",
                    DefectID = "",
                    DefectName = "",

                    DeleteYN = "N",
                };
                dgdOutwareSub.Items.Add(WOOSSC);




            }
            catch (Exception ee)
            {
                MessageBox.Show("오류지점 - ButtonDataGridSubRowDel_Click : " + ee.ToString());
            }
        }

        //서브 그리드 삭제
        private void SubRowDel()
        {
            try
            {
                if (dgdOutwareSub.Items.Count > 0)
                {
                    if (dgdOutwareSub.SelectedItem != null)
                    {
                        if (dgdOutwareSub.CurrentItem != null)
                        {
                            dgdOutwareSub.Items.Remove(dgdOutwareSub.CurrentItem as Win_ord_OutWare_Scan_Sub_CodeView);
                        }
                        else
                        {
                            ListOutwareSub.Add(dgdOutwareSub.SelectedItem as Win_ord_OutWare_Scan_Sub_CodeView);
                            dgdOutwareSub.Items.Remove((dgdOutwareSub.Items[dgdOutwareSub.SelectedIndex]) as Win_ord_OutWare_Scan_Sub_CodeView);
                        }

                        dgdOutwareSub.Refresh();
                    }
                    else
                    {
                        MessageBox.Show("삭제할 데이터를 먼저 선택하세요.");
                    }
                }
            }
            catch (Exception ee)
            {
                MessageBox.Show("오류지점 - ButtonDataGridSubRowDel_Click : " + ee.ToString());
            }
        }

        #region 서브 데이터그리드 방향키 이동 및 셀 포커스
        private void DataGridSub_PreviewKeyDown(object sender, KeyEventArgs e)
        {
            try
            {
                if (e.Key == Key.Down || e.Key == Key.Up || e.Key == Key.Left || e.Key == Key.Right)
                {
                    DataGridSub_KeyDown(sender, e);
                }
            }
            catch (Exception ee)
            {
                MessageBox.Show("오류지점 - DataGridSub_PreviewKeyDown " + ee.ToString());
            }
        }

        private void DataGridSub_KeyDown(object sender, KeyEventArgs e)
        {
            try
            {
                var SubItem = dgdOutwareSub.CurrentItem as Win_ord_OutWare_Scan_Sub_CodeView;
                int rowCount = dgdOutwareSub.Items.IndexOf(dgdOutwareSub.CurrentItem);
                int colCount = dgdOutwareSub.Columns.IndexOf(dgdOutwareSub.CurrentCell.Column);
                int StartColumnCount = 1; //DataGridSub.Columns.IndexOf(dgdtpeMCoperationRateScore);
                int EndColumnCount = 7; //DataGridSub.Columns.IndexOf(dgdtpeComments);

                if (e.Key == Key.Enter)
                {
                    e.Handled = true;
                    (sender as DataGridCell).IsEditing = false;

                    if (EndColumnCount == colCount && dgdOutwareSub.Items.Count - 1 > rowCount)
                    {
                        dgdOutwareSub.SelectedIndex = rowCount + 1;
                        dgdOutwareSub.CurrentCell = new DataGridCellInfo(dgdOutwareSub.Items[rowCount + 1], dgdOutwareSub.Columns[StartColumnCount]);
                    }
                    else if (EndColumnCount > colCount && dgdOutwareSub.Items.Count - 1 > rowCount)
                    {
                        dgdOutwareSub.CurrentCell = new DataGridCellInfo(dgdOutwareSub.Items[rowCount], dgdOutwareSub.Columns[colCount + 1]);
                    }
                    else if (EndColumnCount == colCount && dgdOutwareSub.Items.Count - 1 == rowCount)
                    {
                        btnSave.Focus();
                    }
                    else if (EndColumnCount > colCount && dgdOutwareSub.Items.Count - 1 == rowCount)
                    {
                        dgdOutwareSub.CurrentCell = new DataGridCellInfo(dgdOutwareSub.Items[rowCount], dgdOutwareSub.Columns[colCount + 1]);
                    }
                    else { }
                }
                else if (e.Key == Key.Down)
                {
                    e.Handled = true;
                    (sender as DataGridCell).IsEditing = false;

                    if (dgdOutwareSub.Items.Count - 1 > rowCount)
                    {
                        dgdOutwareSub.SelectedIndex = rowCount + 1;
                        dgdOutwareSub.CurrentCell = new DataGridCellInfo(dgdOutwareSub.Items[rowCount + 1], dgdOutwareSub.Columns[colCount]);
                    }
                    else if (dgdOutwareSub.Items.Count - 1 == rowCount)
                    {
                        if (EndColumnCount > colCount)
                        {
                            dgdOutwareSub.SelectedIndex = 0;
                            dgdOutwareSub.CurrentCell = new DataGridCellInfo(dgdOutwareSub.Items[0], dgdOutwareSub.Columns[colCount + 1]);
                        }
                        else
                        {
                            btnSave.Focus();
                        }
                    }
                }
                else if (e.Key == Key.Up)
                {
                    e.Handled = true;
                    (sender as DataGridCell).IsEditing = false;

                    if (rowCount > 0)
                    {
                        dgdOutwareSub.SelectedIndex = rowCount - 1;
                        dgdOutwareSub.CurrentCell = new DataGridCellInfo(dgdOutwareSub.Items[rowCount - 1], dgdOutwareSub.Columns[colCount]);
                    }
                }
                else if (e.Key == Key.Left)
                {
                    e.Handled = true;
                    (sender as DataGridCell).IsEditing = false;

                    if (colCount > 0)
                    {
                        dgdOutwareSub.CurrentCell = new DataGridCellInfo(dgdOutwareSub.Items[rowCount], dgdOutwareSub.Columns[colCount - 1]);
                    }
                }
                else if (e.Key == Key.Right)
                {
                    e.Handled = true;
                    (sender as DataGridCell).IsEditing = false;

                    if (EndColumnCount > colCount)
                    {
                        dgdOutwareSub.CurrentCell = new DataGridCellInfo(dgdOutwareSub.Items[rowCount], dgdOutwareSub.Columns[colCount + 1]);
                    }
                    else if (EndColumnCount == colCount)
                    {
                        if (dgdOutwareSub.Items.Count - 1 > rowCount)
                        {
                            dgdOutwareSub.SelectedIndex = rowCount + 1;
                            dgdOutwareSub.CurrentCell = new DataGridCellInfo(dgdOutwareSub.Items[rowCount + 1], dgdOutwareSub.Columns[StartColumnCount]);
                        }
                        else
                        {
                            btnSave.Focus();
                        }
                    }
                }
            }
            catch (Exception ee)
            {
                MessageBox.Show("오류지점 - DataGridSub_KeyDown " + ee.ToString());
            }
        }

        private void DataGridSub_TextFocus(object sender, KeyEventArgs e)
        {
            try
            {
                Lib.Instance.DataGridINControlFocus(sender, e);
            }
            catch (Exception ee)
            {
                MessageBox.Show("오류지점 - DataGridSub_TextFocus " + ee.ToString());
            }
        }

        private void DataGridSub_GotFocus(object sender, RoutedEventArgs e)
        {
            try
            {
                if (strFlag.Equals("I") || strFlag.Equals("U"))
                {
                    DataGridCell cell = sender as DataGridCell;
                    cell.IsEditing = true;
                }
            }
            catch (Exception ee)
            {
                MessageBox.Show("오류지점 - DataGridSub_GotFocus " + ee.ToString());
            }
        }

        private void DataGridSub_MouseUp(object sender, MouseButtonEventArgs e)
        {
            try
            {
                Lib.Instance.DataGridINBothByMouseUP(sender, e);
            }
            catch (Exception ee)
            {
                MessageBox.Show("오류지점 - DataGridSub_MouseUp " + ee.ToString());
            }
        }
        #endregion

        //서브 데이터 그리드 수량 숫자만 입력
        private void DataGridTextBoxColorQty_PreviewTextInput(object sender, TextCompositionEventArgs e)
        {
            try
            {
                Lib.Instance.CheckIsNumeric((TextBox)sender, e);             
               
            }
            catch (Exception ee)
            {
                MessageBox.Show("오류지점 - DataGridTextBoxColorQty_PreviewTextInput : " + ee.ToString());
            }
        }


     


        //서브 데이터 그리드 수량 변경 이벤트
        private void DataGridTextBoxColorQty_TextChanged(object sender, TextChangedEventArgs e)
        {
            try
            {   
                SumColorQty();
                
            }
            catch (Exception ee)
            {
                MessageBox.Show("오류지점 - DataGridTextBoxColorQty_TextChanged : " + ee.ToString());
            }
        }

        // 서브 데이터 그리드 단가 변경 이벤트
        private void DataGridTextBoxUnitPrice_PreviewTextInput(object sender, TextCompositionEventArgs e)
        {
            try
            {
                Lib.Instance.CheckIsNumeric((TextBox)sender, e);
               
            }
            catch (Exception ee)
            {
                MessageBox.Show("오류지점 - DataGridTextBoxColorQty_PreviewTextInput : " + ee.ToString());
            }
        }

        #region 서브 그리드 수량 합계
        private void SumColorQty()
        {
            try
            {                
                double OutQty = 0;

                for (int i = 0; i < dgdOutwareSub.Items.Count; i++)
                {
                    var label = dgdOutwareSub.Items[i] as Win_ord_OutWare_Scan_Sub_CodeView;
                    if (label.OutQty != null)
                    {
                        OutQty += lib.returnDouble(label.OutQty.ToString());
                    }
                }

                txtOutRoll.Text = dgdOutwareSub.Items.Count.ToString();
                txtOutQty.Text = lib.returnNumStringZero(OutQty.ToString());

            }
            catch (Exception ee)
            {
                MessageBox.Show("오류지점 - SumQty : " + ee.ToString());
            }
        }

        #endregion

        #region 이동 토글 버튼
        // ID 기준 이동 토글 버튼
        private void tgnMoveByID_Click(object sender, RoutedEventArgs e)
        {
            tgnMoveByID.IsChecked = true;
            tgnMoveByQty.IsChecked = false;

            // 수량 입력 안되도록 → 수량기준이동 토글버튼이 활성화 됬을때만 입력 가능하도록
            txtOutRoll.IsHitTestVisible = false;
            txtOutQty.IsHitTestVisible = false;

            // 바코드 활성화
            txtScanData.IsHitTestVisible = true;

            // 그리드 변경
            dgdOutwareSub.Visibility = Visibility.Visible;

            // OutRoll : 박스수, 서브그리드 갯수 / OutQty : 총 개수 - 구하기 
            //SetOutRollAndOutQty();

        }
        // 수량 기준 이동 토글 버튼
        private void tgnMoveByQty_Click(object sender, RoutedEventArgs e)
        {
            tgnMoveByID.IsChecked = false;
            tgnMoveByQty.IsChecked = true;

            // 수량 입력 되도록 → 바코드로 입력하도록 막아놓자.
            txtOutRoll.IsHitTestVisible = false;
            txtOutQty.IsHitTestVisible = false;

            // 바코드 입력 안되도록 → 수량기준이동은 바코드가 아닌 수량으로 관리
            //txtBarCode.IsHitTestVisible = false;

            // 바코드 활성화
            txtScanData.IsHitTestVisible = true;

            // 그리드 변경
            dgdOutwareSub.Visibility = Visibility.Visible;

            // OutRoll : 박스수, 서브그리드 갯수 / OutQty : 총 개수 - 구하기 
            //SetOutRollAndOutQty();

        }
        #endregion

        // 숫자로 변환 가능한지 체크 이벤트
        private bool CheckConvertInt(string str)
        {
            bool flag = false;
            int chkInt = 0;

            if (!str.Trim().Equals(""))
            {
                str = str.Trim().Replace(",", "");

                if (Int32.TryParse(str, out chkInt) == true)
                    flag = true;
            }

            return flag;
        }

        // Int로 변환
        private int ConvertInt(string str)
        {
            int result = 0;
            int chkInt = 0;

            if (!str.Trim().Equals(""))
            {
                str = str.Replace(",", "");

                if (int.TryParse(str, out chkInt) == true)
                    result = int.Parse(str);
            }

            return result;
        }
              
       

        private void lblpfBuyerArticleNo_Click(object sender, MouseButtonEventArgs e)
        {
            if (chkpfBuyerArticleNo.IsChecked == true)
            {
                txtpfBuyerArticleNo.IsEnabled = false;
                btnpfBuyerArticleNo.IsEnabled = false;
                chkpfBuyerArticleNo.IsChecked = false;
            }
            else
            {
                txtpfBuyerArticleNo.IsEnabled = true;
                btnpfBuyerArticleNo.IsEnabled = true;
                chkpfBuyerArticleNo.IsChecked = true;
            }
        }

        private void ChkpfBuyerArticleNo_Checked(object sender, RoutedEventArgs e)
        {
            if(chkpfBuyerArticleNo.IsChecked == true)
            {
                txtpfBuyerArticleNo.IsEnabled = true;
                btnpfBuyerArticleNo.IsEnabled = true;
                chkpfBuyerArticleNo.IsChecked = true;
            }
        }

        private void ChkpfBuyerArticleNo_Unchecked(object sender, RoutedEventArgs e)
        {
            if(chkpfBuyerArticleNo.IsChecked == false)
            {
                txtpfBuyerArticleNo.IsEnabled = false;
                btnpfBuyerArticleNo.IsEnabled = false;
                chkpfBuyerArticleNo.IsChecked = false;
            }
        }

        private void txtpfBuyerArticleNo_KeyDown(object sender, KeyEventArgs e)
        {
            if(e.Key == Key.Enter)
            {
                pf.ReturnCode(txtpfBuyerArticleNo, 76, txtpfBuyerArticleNo.Text);


            }

        }

        private void btnpfBuyerArticleNo_Click(object sender, RoutedEventArgs e)
        {
            pf.ReturnCode(txtInCustomer, 76, txtpfBuyerArticleNo.Text);

        }
    

        private void rbnOrderNumber_Click(object sender, RoutedEventArgs e)
        {
            Check_SwitchOrderNoAndReqID();
        }

        private void rbnReqID_Click(object sender, RoutedEventArgs e)
        {
            Check_SwitchOrderNoAndReqID();
        }

        #region 기타 메서드 모음 ADD

        //그리드 클리어
        private void ClearGrdInput()
        {
            List<Grid> grids = new List<Grid> { grdInput };

            foreach (Grid grid in grids)
            {
                FindUiObject(grid, child =>
                {
                    if (child is TextBox txtbox)
                    {
                        txtbox.Text = string.Empty;
                        txtbox.Tag = null;
                    }
                    else if (child is DatePicker dtp)
                    {
                        dtp.SelectedDate = null;
                    }
                    else if (child is ComboBox cb)
                    {
                        cb.SelectedValue = null;
                    }
                    //else if (child is DataGrid dgd)
                    //{
                    //    if (dgd.ItemsSource != null)
                    //    {
                    //        var originalCollection = dgd.ItemsSource;
                    //        dgd.ItemsSource = null;

                    //        if (originalCollection is IList list)
                    //        {
                    //            list.Clear();
                    //            dgd.ItemsSource = originalCollection;
                    //        }
                    //        else if (originalCollection is ObservableCollection<object> ovc)
                    //        {
                    //            ovc.Clear();
                    //            dgd.ItemsSource = originalCollection;
                    //        }

                    //    }
                    //    else
                    //    {
                    //        dgd.Items.Clear();
                    //    }
                    //}

                });
            }
        }

        //라벨클릭, 체크박스 토글 코드가 너무 반복 되어서 작성
        private void CommonControl_Click(object sender, EventArgs e)
        {
            CheckBox checkBox = null;
            DependencyObject parentGrid = null;

            if (sender is Label label)
            {
                // 라벨의 부모 그리드 찾기
                parentGrid = FindVisualParent<Grid>(label);
                if (parentGrid != null)
                {
                    // 같은 그리드 내에서 체크박스 찾기
                    checkBox = FindChild<CheckBox>(parentGrid);
                    if (checkBox != null)
                    {
                        // 체크박스 상태 토글
                        checkBox.IsChecked = !checkBox.IsChecked;
                    }
                }
            }
            else if (sender is CheckBox clickedCheckBox)
            {
                // 클릭된 것이 체크박스인 경우
                checkBox = clickedCheckBox;
                parentGrid = FindVisualParent<Grid>(checkBox);
            }

            // 체크박스와 부모 그리드가 있으면 컨트롤 활성화/비활성화 처리
            if (checkBox != null && parentGrid != null)
            {
                List<Control> controlsToToggle = new List<Control>();

                // 그리드 내 모든 Control 찾기 (체크박스 제외)
                FindUiObject(parentGrid, obj => {
                    if (obj is Control control && obj != checkBox && !(obj is Label) && !(obj is CheckBox))
                    {
                        controlsToToggle.Add(control);
                    }
                });

                // 컨트롤 활성화/비활성화
                foreach (var control in controlsToToggle)
                {
                    control.IsEnabled = checkBox.IsChecked == true;
                }
            }
        }

        //UI컨트롤 요소찾기
        private void FindUiObject(DependencyObject parent, Action<DependencyObject> action)
        {
            int childCount = VisualTreeHelper.GetChildrenCount(parent);
            for (int i = 0; i < childCount; i++)
            {
                var child = VisualTreeHelper.GetChild(parent, i);

                action?.Invoke(child);

                FindUiObject(child, action);
            }
        }

        //컨트롤 안 특정 타입의 자식 컨트롤을 찾는 함수 (그리드내에서)
        //var parentContainer = VisualTreeHelper.GetParent(checkbox);
        //var datePicker = FindChild<DatePicker>(parentContainer);
        private T FindChild<T>(DependencyObject parent) where T : DependencyObject
        {
            int childCount = VisualTreeHelper.GetChildrenCount(parent);
            for (int i = 0; i < childCount; i++)
            {
                var child = VisualTreeHelper.GetChild(parent, i);
                if (child is T typedChild)
                {
                    return typedChild;
                }

                // 재귀적으로 자식의 자식들도 검색
                var result = FindChild<T>(child);
                if (result != null)
                    return result;
            }
            return null;
        }


        // 자식요소 안에서 부모요소 찾기
        //DataGridRow row = FindVisualParent<DataGridRow>(checkBox); 데이터그리드안의 행속 체크박스의 부모행 찾기
        //DataGrid parentGrid = FindVisualParent<DataGrid>(row); 데이터그리드 행의 부모 데이터그리드 찾기
        private T FindVisualParent<T>(DependencyObject child) where T : DependencyObject
        {
            DependencyObject parentObject = VisualTreeHelper.GetParent(child);

            if (parentObject == null)
                return null;

            T parent = parentObject as T;
            if (parent != null)
                return parent;
            else
                return FindVisualParent<T>(parentObject);
        }

        //8자리 char형태 날짜 년도-월-일 하이픈 삽입
        //16자리 일경우 8자리 사이에 ~ 삽입
        private string DateTypeHyphen(string DigitsDate)
        {
            string pattern1 = @"(\d{4})(\d{2})(\d{2})";
            string pattern2 = @"(\d{4})(\d{2})(\d{2})(\d{4})(\d{2})(\d{2})";

            if (DigitsDate.Length == 8)
            {
                DigitsDate = Regex.Replace(DigitsDate, pattern1, "$1-$2-$3");
            }
            else if (DigitsDate.Length == 16)
            {
                DigitsDate = Regex.Replace(DigitsDate, pattern2, "$1-$2-$3 ~ $4-$5-$6");
            }
            else if (DigitsDate.Length == 0)
            {
                DigitsDate = string.Empty;
            }

            return DigitsDate;
        }


        private string TimeTypeColon(string DigitsTime)
        {
            string pattern1 = @"(\d{2})(\d{2})";

            if (DigitsTime.Length == 4)
            {
                DigitsTime = Regex.Replace(DigitsTime, pattern1, "$1:$2");
            }

            return DigitsTime;
        }
        private object RemoveComma(object obj, bool returnAsNumber = false, Type returnType = null)
        {
            //파라미터가 만약 null일때
            if (obj == null)
            {
                //숫자타입이 false면 string으로 내보내기
                if (!returnAsNumber) return "0";

                // 만약 숫자타입을 써야되면 returnType파라미터의 받은 형태로 전달
                // null일 때도 returnType에 따라 적절한 타입의 0 반환
                switch (returnType?.Name)
                {
                    case "Decimal": return (object)0m;  //monetary
                    case "Double": return (object)0d;   //double
                    case "Int64": return (object)0L;    //long
                    default: return (object)0;          //int
                }
            }

            string digits = obj.ToString()
                              .Trim()
                              .Replace(",", "");

            //만약 빈공백(blank)이더라도 0으로 내보내야한다.
            if (string.IsNullOrEmpty(digits))
            {
                if (!returnAsNumber) return "0";

                // returnType을 활용해서 적절한 타입으로 반환
                switch (returnType?.Name)
                {
                    case "Decimal": return (object)0m;
                    case "Double": return (object)0d;
                    case "Int64": return (object)0L;
                    default: return (object)0;
                }
            }


            try
            {
                Type targetType = returnType ?? typeof(int);

                //혹시나 하는 예외처리
                //입력 컨트롤간에 LostFocus나 TextChanged같은 걸로 계산을 할 때
                //처리 가능한 숫자 범위를 초과하면 오류가 발생하므로
                //초과하면 해당 자료형타입이 처리할 수 있는 최대 숫자를 표시해줌
                switch (targetType.Name)
                {
                    case "Int32":
                        if (decimal.TryParse(digits, out decimal intParsed))
                        {
                            if (intParsed > int.MaxValue) return int.MaxValue;
                            if (intParsed < int.MinValue) return int.MinValue;
                            return (int)intParsed;
                        }
                        return int.MaxValue;

                    case "Int64":
                        if (decimal.TryParse(digits, out decimal longParsed))
                        {
                            if (longParsed > long.MaxValue) return long.MaxValue;
                            if (longParsed < long.MinValue) return long.MinValue;
                            return (long)longParsed;
                        }
                        return long.MaxValue;

                    case "Double":
                        if (double.TryParse(digits, out double doubleParsed))
                        {
                            return doubleParsed;
                        }
                        return double.MaxValue;

                    case "Decimal":
                        if (decimal.TryParse(digits, out decimal decimalParsed))
                        {
                            return decimalParsed;
                        }
                        return decimal.MaxValue;

                    default:
                        return int.MaxValue;
                }
            }
            catch
            {

                if (returnType != null)
                {
                    switch (returnType.Name)
                    {
                        case "Int32":
                            return int.MaxValue;
                        case "Int64":
                            return long.MaxValue;
                        case "Double":
                            return double.MaxValue;
                        case "Decimal":
                            return decimal.MaxValue;
                        default:
                            return int.MaxValue;
                    }
                }
                return int.MaxValue;
            }
        }

        private string ConvertDate(DatePicker datePicker)
        {
            if (datePicker.SelectedDate != null)
                return datePicker.SelectedDate.Value.ToString("yyyyMMdd");
            else
                return string.Empty;
        }

        private bool IsDatePickerNull(DatePicker datePicker)
        {
            if (datePicker.SelectedDate == null)
                return true;
            else
                return false;
        }



        //텍스트박스 , DatePicker, 콤보박스의 바인딩 값과 넘겨주는 오브젝트 value가 일치하는 곳에
        //자동으로 바인딩
        //사용하려하면 바인딩하려는 UI개체에 updateSourceTrigger를 propertyChange, Tag값도 변경하려면 mode=TwoWay를 작성하세요
        private void AutoBindDataToControls(object dataObject, DependencyObject parent)
        {
            var properties = dataObject.GetType().GetProperties()
                .ToDictionary(p => p.Name.ToLower(), p => p);

            // TextBox 처리
            var textBoxes = FindAllControls<TextBox>(parent);
            foreach (var textBox in textBoxes)
            {
                // Text 바인딩 처리
                var textBinding = BindingOperations.GetBinding(textBox, TextBox.TextProperty);
                if (textBinding != null && !string.IsNullOrEmpty(textBinding.Path.Path))
                {
                    var textPropertyName = textBinding.Path.Path.ToLower();
                    if (properties.TryGetValue(textPropertyName, out var textProperty))
                    {
                        var textValue = textProperty.GetValue(dataObject)?.ToString();
                        if (decimal.TryParse(textValue, out _))
                            textBox.Text = stringFormatN0(textValue);
                        else
                            textBox.Text = textValue;
                    }
                }

                // Tag 바인딩 처리
                var tagBinding = BindingOperations.GetBinding(textBox, TextBox.TagProperty);
                if (tagBinding != null && !string.IsNullOrEmpty(tagBinding.Path.Path))
                {
                    var tagPropertyName = tagBinding.Path.Path.ToLower();
                    if (properties.TryGetValue(tagPropertyName, out var tagProperty))
                    {
                        textBox.Tag = tagProperty.GetValue(dataObject)?.ToString();
                    }
                }
            }

            // DatePicker 처리
            var datePickers = FindAllControls<DatePicker>(parent);
            foreach (var datePicker in datePickers)
            {
                var binding = BindingOperations.GetBinding(datePicker, DatePicker.SelectedDateProperty);
                if (binding != null && !string.IsNullOrEmpty(binding.Path.Path))
                {
                    var propertyName = binding.Path.Path.ToLower();
                    if (properties.TryGetValue(propertyName, out var property))
                    {
                        datePicker.SelectedDate = ConvertToDateTime(property.GetValue(dataObject)?.ToString());
                    }
                }
            }

            // ComboBox 처리
            var comboBoxes = FindAllControls<ComboBox>(parent);
            foreach (var comboBox in comboBoxes)
            {
                var binding = BindingOperations.GetBinding(comboBox, ComboBox.SelectedValueProperty);
                if (binding != null && !string.IsNullOrEmpty(binding.Path.Path))
                {
                    var propertyName = binding.Path.Path.ToLower();
                    if (properties.TryGetValue(propertyName, out var property))
                    {
                        comboBox.SelectedValue = property.GetValue(dataObject)?.ToString();
                    }
                }
            }
        }

        private IEnumerable<T> FindAllControls<T>(DependencyObject parent) where T : DependencyObject
        {
            var count = VisualTreeHelper.GetChildrenCount(parent);
            for (int i = 0; i < count; i++)
            {
                var child = VisualTreeHelper.GetChild(parent, i);

                if (child is T control)
                    yield return control;

                foreach (var descendant in FindAllControls<T>(child))
                    yield return descendant;
            }
        }

        // 단일 컨트롤을 찾는 메서드도 필요할 수 있습니다
        private T FindControl<T>(DependencyObject parent, string name) where T : FrameworkElement
        {
            if (parent == null) return null;

            var count = VisualTreeHelper.GetChildrenCount(parent);
            for (int i = 0; i < count; i++)
            {
                var child = VisualTreeHelper.GetChild(parent, i);

                if (child is T control && control.Name == name)
                    return control;

                var result = FindControl<T>(child, name);
                if (result != null)
                    return result;
            }

            return null;
        }

        private DateTime? ConvertToDateTime(string dateStr)
        {
            if (string.IsNullOrEmpty(dateStr?.Trim()))
                return null;

            // 특수문자 제거 (숫자만 남김)
            string cleanDate = new string(dateStr.Where(char.IsDigit).ToArray());

            // 8자리가 아닌 경우 null 반환
            if (cleanDate.Length != 8)
                return null;

            try
            {
                return DateTime.ParseExact(cleanDate, "yyyyMMdd", null);
            }
            catch
            {
                return null;
            }
        }







        #endregion

        private void btnCustomLabelScanYN_Checked(object sender, RoutedEventArgs e)
        {
            ToolTip tooltip = new ToolTip();
            tooltip.Content = "입력한 바코드가 고객라벨에 입력됩니다.";

            ToggleButton tbn = sender as ToggleButton;
            tbn.ToolTip = tooltip;
            tooltip.IsOpen = true;

            System.Windows.Threading.DispatcherTimer timer = new System.Windows.Threading.DispatcherTimer();

            timer.Interval = TimeSpan.FromSeconds(2);
            timer.Tick += (s, args) => {
                tooltip.IsOpen = false;
                timer.Stop();
            };
            timer.Start();
        }
    }


    class Win_ord_OutWare_Scan_CodeView : BaseView
    {

        public bool Chk { get; set; }
        public int num { get; set; }
        public string OutwareID { get; set; }
        public string OrderID { get; set; }
        public string OutwareReqID { get; set; }
        public string OutSeq { get; set; }
        public string OrderNo { get; set; }
        public string CustomID { get; set; }
        public string KCustom { get; set; }
        public string OutDate { get; set; }
        public string ArticleID { get; set; }
        public string Article { get; set; }
        public string OutClss { get; set; }
        public string WorkID { get; set; }
        public string OutRoll { get; set; }
        public string OutQty { get; set; }
        public string OutRealQty { get; set; }
        public string ResultDate { get; set; }
        public string OrderQty { get; set; }
        public string UnitClss { get; set; }
        public string WorkName { get; set; }
        public string OutType { get; set; }
        public string Remark { get; set; }
        public string BuyerModel { get; set; }
        public string BuyerModelID { get; set; }
        public string OutSumQty { get; set; }
        public string OutQtyY { get; set; }
        public string StuffinQty { get; set; }
        public string OutWeight { get; set; }
        public string OutRealWeight { get; set; }
        public string UnitPriceClss { get; set; }
        public string BuyerDirectYN { get; set; }
        public string Vat_Ind_YN { get; set; }
        public string workID { get; set; }
        public string InsStuffINYN { get; set; }
        public string ExchRate { get; set; }
        public string FromLocID { get; set; }
        public string TOLocID { get; set; }
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
        public string Buyer_Chief { get; set; }
        public string Buyer_Address1 { get; set; }
        public string Buyer_Address2 { get; set; }
        public string Buyer_Address3 { get; set; }
        public string CustomNo { get; set; }
        public string Chief { get; set; }
        public string Address1 { get; set; }
        public string Address2 { get; set; }
        public string Address3 { get; set; }
        public string OutCustom { get; set; }
        public string OutSubType { get; set; }

        public string RemainQty { get; set; }
        public string DvlyCustomID { get; set; }
        public string DvlyCustom { get; set; }

        //2021-05-31
        public string Category { get; set; }
        public string Condition { get; set; }

        public string OutTime { get; set; }

    }

    public class Win_ord_OutWare_Scan_Sub_CodeView : BaseView
    {
        public int Num { get; set; }

        public bool Chk { get; set; }

        public string OutwareID { get; set; }
        public string OutSubSeq { get; set; }
        public string LabelID { get; set; }
        public string LabelGubun { get; set; }
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
        public string TOLocID { get; set; }
        public string UnitClss { get; set; }
        public string UnitClssName { get; set; }
        public string ArticleID { get; set; }
        public string Article { get; set; }
        public string BuyerArticleNo { get; set; }

        public string OutClss { get; set; }
        public string Gubun { get; set; }
        public string DefectID { get; set; }
        public string DefectName { get; set; }

        public string DeleteYN { get; set; }

        public string OutRealQty { get; set; }
        public string CustomName { get; set; }

        public string SubRemark { get; set; }
        public string Spec { get; set; }


        public bool UDFlag { get; set; }
        public double dOutQty { get; set; }
        public double dUnitPrice { get; set; }
    }

    public class Win_ord_OutWare_Scan_dgdTotal_CodeView : BaseView
    {
        public int count { get; set; }
        public string totQty { get; set; }
    
    }

    }

using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
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
using System.Collections.ObjectModel;
using WizMes_WellMade;
using WizMes_WellMade.PopUP;

namespace WizMes_WellMade
{
    /// <summary>
    /// Win_mtr_LotStockControl_U.xaml에 대한 상호 작용 논리
    /// </summary>
    public partial class Win_mtr_LotStockControl_U : UserControl
    {
        public Win_mtr_LotStockControl_U()
        {
            InitializeComponent();
        }

        Lib lib = new Lib();
        PlusFinder pf = new PlusFinder();

        // 수정 정보를 보관하기 위한 변수
        List<Win_mtr_LotStockControl_U_CodeView> lstLotStock = new List<Win_mtr_LotStockControl_U_CodeView>();

        List<String> LabelID = new List<String>();

        String LabelIDList = "";

        int rowNum = 0;
        string strFlag = string.Empty;

        private void UserControl_Loaded(object sender, RoutedEventArgs e)
        {
            CanBtnControl();
            ComboBoxSetting();

            dtpFromDate.SelectedDate = DateTime.Today;
            dtpToDate.SelectedDate = DateTime.Today;


        }



        // 콤보박스셋팅
        private void ComboBoxSetting()
        {
            cboWareHouse.Items.Clear();

            ObservableCollection<CodeView> cbWareHouse = ComboBoxUtil.Instance.Gf_DB_CM_GetComCodeDataset(null, "LOC", "Y", "", "");

            this.cboWareHouse.ItemsSource = cbWareHouse;
            this.cboWareHouse.DisplayMemberPath = "code_name";
            this.cboWareHouse.SelectedValuePath = "code_id";
            this.cboWareHouse.SelectedIndex = 0;

        }

        #region CRUD 버튼

        //추가
        private void btnAdd_Click(object sender, RoutedEventArgs e)
        {
            lstLotStock.Clear();

            CantBtnControl();
            
            strFlag = "I";
            this.DataContext = null;

            //서브그리드 비워주기
            if (dgdSub.Items.Count > 0)
            {
                dgdSub.Items.Clear();
            }

            //조정일자는 오늘날짜 기본 셋팅
            dtpAdjustDate.SelectedDate = DateTime.Today;

            //작업자는 로그인한 아이디로 기본셋팅
            //txtWorker.Tag = MainWindow.CurrentUser; //로그인 할때 ID

            txtWorker.Tag = MainWindow.CurrentPersonID; //2021-10-28 로 수정 CurrentUser(로그인 아이디)와 CurrentPersonID(로그인 아이디의 PersonID)가 다름
            txtWorker.Text = MainWindow.CurrentPerson; //mt_person , Name
        }

        //수정
        private void BtnUpdate_Click(object sender, RoutedEventArgs e)
        {
            var UpdateData = dgdMain.SelectedItem as Win_mtr_LotStockControl_U_CodeView;

            if (UpdateData != null)
            {
                CantBtnControl();

                strFlag = "U";
                //2021-10-28 수정 시 작업자 수정 안하면 해당 작업자 그대로 저장하기
                txtWorker.Tag = UpdateData.PersonID; 
                txtWorker.Text = UpdateData.Name;
            }

            // 임시 : 오늘 날짜 아닌건 추가 못하도록 설정
            if (dtpAdjustDate.SelectedDate != null
                && dtpAdjustDate.SelectedDate.Value.ToString("yyyyMMdd").Equals(DateTime.Today.ToString("yyyyMMdd")) == false)
            {
                btnChoice.IsEnabled = false;
                btnPlus.IsEnabled = false;
            }
            else
            {
                btnChoice.IsEnabled = true;
                btnPlus.IsEnabled = true;
            }
        }

        //삭제
        private void btnDelete_Click(object sender, RoutedEventArgs e)
        {
            var DeleteItem = dgdMain.SelectedItem as Win_mtr_LotStockControl_U_CodeView;

            if (DeleteItem != null)
            {
                string Msg = "";
                // 삭제 전에 체크하기 → 이전 기록이면 삭제 못하도록
                for (int i = 0; i < dgdSub.Items.Count; i++)
                {
                    var Sub = dgdSub.Items[i] as Win_mtr_LotStockControl_U_CodeView;
                    if (Sub != null
                        && Sub.UDFlag == false)
                    {
                        Msg += "\r " + Sub.LotID + " : " + DatePickerFormat(Sub.LastDate);
                    }
                }
                if (Msg.Equals("") == false)
                {
                    Msg = "재고조정 삭제는 가장 최근 날짜로 등록된 기록만 삭제 가능합니다." + Msg;

                    MessageBox.Show(Msg);
                    return;
                }

                if (MessageBox.Show("선택한 항목을 삭제하시겠습니까?", "삭제 전 확인", MessageBoxButton.YesNo) == MessageBoxResult.Yes)
                {
                    //2021-10-18 MainGrid에서 선택 후 삭제할 경우 체크 로직
                    //삭제하면 마이너스 재고 생성되는 라벨을 리스트에 담기
                    for(int i = 0; i < dgdSub.Items.Count; i++)
                    {
                        var SubCheck = dgdSub.Items[i] as Win_mtr_LotStockControl_U_CodeView;
                        if(!DeleteDataMainCheck(SubCheck.ArticleID, SubCheck.LotID, DeleteItem.ControlID))
                        {
                            LabelID.Add(SubCheck.LotID);
                        }
                    }

                    //삭제하면 안되는 라벨이 존재한다면, 나열 한 뒤에 메세지에 띄우기
                    if(LabelID.Count > 0)
                    {
                        for(int i = 0; i < LabelID.Count; i++)
                        {
                            if (1 < LabelID.Count && i < LabelID.Count - 1)
                            {
                                LabelIDList += LabelID[i] + ", ";
                            }
                            else if(i == LabelID.Count - 1 || LabelID.Count == 1)
                            {
                                LabelIDList += LabelID[i];
                            }                          
                        }
                        //YES 누르면 무시하고 삭제 됨
                        if(MessageBox.Show("삭제 시 마이너스 재고가 발생하는 라벨이 존재합니다. 삭제하시겠습니까?  \r" + LabelID.Count + "개 발생 " + LabelIDList ,"삭제 전 확인", MessageBoxButton.YesNo) == MessageBoxResult.Yes)
                        {
                            LabelID.Clear();
                            if (DeleteData(DeleteItem.ControlID))
                            {
                                re_Search();
                            }
                        }
                        else
                        {
                            LabelID.Clear();
                            return;
                        }
                    }
                    //삭제하면 안되는 라벨이 없다면 삭제
                    else
                    {
                        LabelID.Clear();
                        if (DeleteData(DeleteItem.ControlID))
                        {
                            re_Search();
                        }
                    }

                }
            }
            else
            {
                MessageBox.Show("삭제할 데이터를 선택해주세요.");
            }
        }

        //닫기
        private void btnClose_Click(object sender, RoutedEventArgs e)
        {
            Lib.Instance.ChildMenuClose(this.ToString());
        }

        //검색
        private void btnSearch_Click(object sender, RoutedEventArgs e)
        {
            //검색버튼 비활성화
            btnSearch.IsEnabled = false;

            Dispatcher.BeginInvoke(new Action(() =>

            {
                Thread.Sleep(2000);

                //로직
                re_Search();

            }), System.Windows.Threading.DispatcherPriority.Background);



            Dispatcher.BeginInvoke(new Action(() =>

            {
                btnSearch.IsEnabled = true;

            }), System.Windows.Threading.DispatcherPriority.Background);
        }

        //저장
        private void btnSave_Click(object sender, RoutedEventArgs e)
        {
            if (SaveData(strFlag, txtControlID.Text))
            {
                //Lib.Instance.DBReIndex(); //2021-11-10 재고 실사 후 DBReIndex

                CanBtnControl();

                re_Search();

                //저장이 성공 했다면, FillGrid 후에 strFlag 비워주기 
                strFlag = string.Empty;
            }
        }

        //취소
        private void btnCancel_Click(object sender, RoutedEventArgs e)
        {
            strFlag = string.Empty;

            CanBtnControl();

            re_Search();
        }

        //엑셀
        private void btnExcel_Click(object sender, RoutedEventArgs e)
        {
            DataTable dt = null;
            string Name = string.Empty;

            string[] lst = new string[4];
            lst[0] = "재고조정";
            lst[1] = "재고조정세부내역";
            lst[2] = dgdMain.Name;
            lst[3] = dgdSub.Name;

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
                }
                else if (ExpExc.choice.Equals(dgdSub.Name))
                {
                    if (ExpExc.Check.Equals("Y"))
                        dt = Lib.Instance.DataGridToDTinHidden(dgdSub);
                    else
                        dt = Lib.Instance.DataGirdToDataTable(dgdSub);

                    Name = dgdSub.Name;
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

        #endregion CRUD 버튼

        #region 상단조건 

        //조정일자 체크
        private void ChkAdjustDate_Checked(object sender, RoutedEventArgs e)
        {
            if (dtpFromDate != null && dtpToDate != null)
            {
                dtpFromDate.IsEnabled = true;
                dtpToDate.IsEnabled = true;
            }
        }

        //조정일자 체크해제
        private void ChkAdjustDate_Unchecked(object sender, RoutedEventArgs e)
        {
            dtpFromDate.IsEnabled = false;
            dtpToDate.IsEnabled = false;
        }

        //조정일자 라벨 
        private void Label_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            if (chkAdjustDate.IsChecked == true)
            {
                chkAdjustDate.IsChecked = false;
                dtpFromDate.IsEnabled = false;
                dtpToDate.IsEnabled = false;
            }
            else
            {
                chkAdjustDate.IsChecked = true;
                dtpFromDate.IsEnabled = true;
                dtpToDate.IsEnabled = true;
            }
        }

        //전일
        private void btnYesterDay_Click(object sender, RoutedEventArgs e)
        {
            DateTime[] SearchDate = lib.BringLastDayDateTimeContinue(dtpToDate.SelectedDate.Value);

            dtpFromDate.SelectedDate = SearchDate[0];
            dtpToDate.SelectedDate = SearchDate[1];
        }

        //금일
        private void btnToday_Click(object sender, RoutedEventArgs e)
        {
            dtpFromDate.SelectedDate = DateTime.Today;
            dtpToDate.SelectedDate = DateTime.Today;
        }

        // 전월 버튼 클릭 이벤트
        private void btnLastMonth_Click(object sender, RoutedEventArgs e)
        {
            DateTime[] SearchDate = lib.BringLastMonthContinue(dtpFromDate.SelectedDate.Value);

            dtpFromDate.SelectedDate = SearchDate[0];
            dtpToDate.SelectedDate = SearchDate[1];
        }

        // 금월 버튼 클릭 이벤트
        private void btnThisMonth_Click(object sender, RoutedEventArgs e)
        {
            dtpFromDate.SelectedDate = lib.BringThisMonthDatetimeList()[0];
            dtpToDate.SelectedDate = lib.BringThisMonthDatetimeList()[1];
        }


        //품명 라벨
        private void Article_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            if (chkArticle.IsChecked == true)
            {
                chkArticle.IsChecked = false;
                txtArticle.IsEnabled = false;
                btnArticle.IsEnabled = false;
            }
            else
            {
                chkArticle.IsChecked = true;
                txtArticle.IsEnabled = true;
                btnArticle.IsEnabled = true;
            }
        }

        //품명 체크 
        private void ChkArticle_Checked(object sender, RoutedEventArgs e)
        {
            txtArticle.IsEnabled = true;
            btnArticle.IsEnabled = true;
        }

        //품명 체크해제
        private void ChkArticle_Unchecked(object sender, RoutedEventArgs e)
        {
            txtArticle.IsEnabled = false;
            btnArticle.IsEnabled = false;
        }

        //품명 키다운
        private void TxtArticleSrh_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                pf.ReturnCode(txtArticle, 76, txtArticle.Text);
            }
        }

        //품명 플러스 파인더
        private void btnArticleSrh_Click(object sender, RoutedEventArgs e)
        {
            pf.ReturnCode(txtArticle, 76, txtArticle.Text);
        }

        // 라벨 검색
        private void lblLotIDSrh_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            if (chkLotIDSrh.IsChecked == true)
            {
                chkLotIDSrh.IsChecked = false;
            }
            else
            {
                chkLotIDSrh.IsChecked = true;
            }
        }
        private void chkLotIDSrh_Checked(object sender, RoutedEventArgs e)
        {
            chkLotIDSrh.IsChecked = true;
            txtLotIDSrh.IsEnabled = true;
        }
        private void chkLotIDSrh_Unchecked(object sender, RoutedEventArgs e)
        {
            chkLotIDSrh.IsChecked = false;
            txtLotIDSrh.IsEnabled = false;
        }

        private void txtLotIDSrh_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                re_Search();
            }

        }

        // 가장 최근에 조정한 라벨들만
        private void lblRecency_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            if (chkRecency.IsChecked == true)
            {
                chkRecency.IsChecked = false;
            }
            else
            {
                chkRecency.IsChecked = true;
            }
        }

        #endregion 상단조건 

        #region 텍스트 박스 이벤트, 서브그리드 버튼

        //??
        private void dtpAdjustDate_PreviewKeyUp(object sender, KeyEventArgs e)
        {

        }

        //조정일자 
        private void dtpAdjustDate_CalendarClosed(object sender, RoutedEventArgs e)
        {

        }

        //작업자 키다운
        private void txtWorker_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                MainWindow.pf.ReturnCode(txtWorker, (int)Defind_CodeFind.DCF_PERSON, "");
            }
        }

        //작업자 플러스 파인더
        private void btnWorker_Click(object sender, RoutedEventArgs e)
        {
            MainWindow.pf.ReturnCode(txtWorker, (int)Defind_CodeFind.DCF_PERSON, "");
        }


        //현재고 조회
        private void BtnStockCheck_Click(object sender, RoutedEventArgs e)
        {
            if (dgdSub.Items.Count > 0)
            {
                dgdSub.Items.Clear();
            }

            try
            {
                Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                sqlParameter.Clear();

                sqlParameter.Add("sDate", dtpAdjustDate.SelectedDate.Value.ToString("yyyyMMdd"));
                sqlParameter.Add("ChkArticle", chkArticle.IsChecked == true ? 1 : 0);
                sqlParameter.Add("ArticleID", chkArticle.IsChecked == true && txtArticle.Tag != null ? txtArticle.Tag.ToString() : "");

                DataSet ds = DataStore.Instance.ProcedureToDataSet("xp_sbStock_sLotStockControl_LotStock", sqlParameter, false);

                if (ds != null && ds.Tables.Count > 0)
                {
                    DataTable dt = ds.Tables[0];

                    if (dt.Rows.Count == 0)
                    {
                        MessageBox.Show("조회된 데이터가 없습니다.");
                    }
                    else
                    {
                        int i = 0;
                        DataRowCollection drc = dt.Rows;

                        foreach (DataRow dr in drc)
                        {
                            i++;
                            var NowStockData = new Win_mtr_LotStockControl_U_CodeView
                            {
                                Num = i,
                                ArticleID = dr["ArticleID"].ToString(),
                                LotID = dr["LotID"].ToString(),
                                UnitClss = dr["UnitClss"].ToString(),
                                StuffinQty = dr["StuffinQty"].ToString(),
                                OutQty = dr["Outqty"].ToString(),
                                StockQty = stringFormatN0(dr["StockQty"]),
                                Article = dr["Article"].ToString(),
                                UnitClssName = dr["UnitClssName"].ToString(),
                                BuyerArticleNo = dr["BuyerArticleNo"].ToString(),
                            };

                            dgdSub.Items.Add(NowStockData);
                        }
                    }
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

        //서브 그리드 행 추가
        private void btnPlus_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                #region 꺼졍
                ////0. 마우스를 쓰고있을 당신을 위해, + 혹시모르니까. 추가 누를때마다, sub 그리드 view upload.
                //int upgradePoint = dgdSub.Items.Count;
                //for (int i = 0; i < upgradePoint; i++)
                //{
                //    DataGridRow dgr = lib.GetRow(i, dgdSub);
                //    var ViewReceiver = dgr.Item as Win_mtr_LotStockControl_U_CodeView;
                //    if (ViewReceiver != null)
                //    {
                //        DataGridCell cell1 = lib.GetCell(i, 6, dgdSub);
                //        TextBox tb1 = lib.GetVisualChild<TextBox>(cell1);
                //        DataGridCell cell2 = lib.GetCell(i, 7, dgdSub);
                //        TextBox tb2 = lib.GetVisualChild<TextBox>(cell2);

                //        ViewReceiver.ControlQty = tb1.Text;
                //        ViewReceiver.Comments = tb2.Text;
                //    }
                //}
                //dgdSub.Items.Refresh();

                ////1.  추가되어 지는 새 항목을 넣는 작업.
                //var Win_mtr_LotStockControl_U_Insert = new Win_mtr_LotStockControl_U_CodeView()
                //{
                //    ControlID = string.Empty,
                //    ControlSeq = string.Empty,
                //    ControlQty = string.Empty,
                //    UnitClssName = string.Empty,
                //    UnitClss = string.Empty,
                //    Comments = string.Empty,
                //    ArticleID = string.Empty,
                //    Article = string.Empty,
                //    ArticleGrp = string.Empty

                //};
                //dgdSub.Items.Add(Win_mtr_LotStockControl_U_Insert);


                //if (dgdSub.Items.Count == 1)
                //{
                //    dgdSub.Focus();
                //    dgdSub.CurrentCell = new DataGridCellInfo(dgdSub.Items[0], dgdSub.Columns[0]);
                //}

                #endregion

                int i = 1;

                if (dgdSub.Items.Count > 0)
                    i = dgdSub.Items.Count + 1;

                var LotStockSub = new Win_mtr_LotStockControl_U_CodeView()
                {
                    Num = i,
                    LotID = "",
                    LotName = "",
                    BuyerArticleNo = "",
                    Article = "",
                    ArticleID = "",
                    UnitClssName = "",
                    UnitClss = "",
                    StockQty = "0",
                    ControlQty = "0",
                    Comments = "",
                    UDFlag = true,
                    //IsEnabled = true
                };

                dgdSub.Items.Add(LotStockSub);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }


        }

        // 서브 그리드 행 삭제
        private void btnMinus_Click(object sender, RoutedEventArgs e)
        {

            var Sub = dgdSub.SelectedItem as Win_mtr_LotStockControl_U_CodeView;
            if (Sub != null)
            {
                // 삭제전 체크
                if (Sub.UDFlag == false)
                {
                    MessageBox.Show("재고조정 삭제는 가장 최근 날짜로 등록된 기록만 삭제 가능합니다.\r(" + Sub.LotID + " : " + DatePickerFormat(Sub.LastDate) + ")");
                    return;
                }

                int selIndex = dgdSub.SelectedIndex - 1;
                if (selIndex < 0) { selIndex = 0; }
                //2021-10-15 삭제하기 전에 삭제해도 되는지 체크 함수 생성
                if(!DeleteDataCheck(Sub.ArticleID, Sub.LotID, txtControlID.Text))
                {
                    return;
                }

                dgdSub.Items.Remove(Sub);

                if (dgdSub.Items.Count > 0)
                {
                    dgdSub.SelectedIndex = selIndex;
                }

                setNumSubDgd();
            }
        }

        // 서브그리드 삭제 시 → Num 재정렬
        private void setNumSubDgd()
        {
            int index = 0;
            for(int i = 0; i < dgdSub.Items.Count; i++)
            {
                var Sub = dgdSub.Items[i] as Win_mtr_LotStockControl_U_CodeView;
                if (Sub != null)
                {
                    index++;
                    Sub.Num = index;
                }
            }
        }

        #endregion 텍스트 박스 이벤트, 서브그리드 버튼

        #region 버튼 컨트롤

        //수정, 추가 저장 후
        private void CanBtnControl()
        {
            btnAdd.IsEnabled = true;
            btnUpdate.IsEnabled = true;
            btnSearch.IsEnabled = true;
            btnSave.Visibility = Visibility.Hidden;
            btnCancel.Visibility = Visibility.Hidden;
            btnExcel.Visibility = Visibility.Visible;
            EventLabel.Visibility = Visibility.Hidden;
            dgdMain.IsHitTestVisible = true;
            dgdSub.IsHitTestVisible = true;

            dtpAdjustDate.IsHitTestVisible = false;
            txtWorker.IsHitTestVisible = false;
            btnWorker.IsHitTestVisible = false;
            txtReason.IsHitTestVisible = false;
            //btnStockCheck.IsEnabled = false;
            btnMinus.IsEnabled = false;

            btnChoice.IsEnabled = false;
            btnPlus.IsEnabled = false;
        }

        //수정, 추가 진행 중
        private void CantBtnControl()
        {
            btnAdd.IsEnabled = false;
            btnUpdate.IsEnabled = false;
            btnSearch.IsEnabled = false;
            btnSave.Visibility = Visibility.Visible;
            btnCancel.Visibility = Visibility.Visible;
            btnExcel.Visibility = Visibility.Hidden;
            EventLabel.Visibility = Visibility.Visible;
            EventLabel.Content = "자료 입력 중";
            dgdMain.IsHitTestVisible = false;
            dgdSub.IsHitTestVisible = true;

            dtpAdjustDate.IsHitTestVisible = true;
            txtWorker.IsHitTestVisible = true;
            btnWorker.IsHitTestVisible = true;
            txtReason.IsHitTestVisible = true;
            //btnStockCheck.IsEnabled = true;
            btnMinus.IsEnabled = true;

            btnChoice.IsEnabled = true;
            btnPlus.IsEnabled = true;

        }
        #endregion 버튼 컨트롤

        #region re_search, FillGrid, SelectionChanged

        //재조회
        private void re_Search()
        {
            if (strFlag.Equals(string.Empty))
            {
                rowNum = 0;
            }
            else if (strFlag.Equals("I"))
            {
                rowNum = dgdMain.Items.Count;
            }
            else if (strFlag.Equals("U"))
            {
                rowNum = dgdMain.SelectedIndex;
            }

            FillGrid(rowNum);
        }

        //조회
        private void FillGrid(int rowNum)
        {
            if (dgdMain.Items.Count > 0)
            {
                dgdMain.Items.Clear();
            }
            if (dgdSub.Items.Count > 0)
            {
                dgdSub.Items.Clear();
            }

            try
            {
                Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                sqlParameter.Clear();

                sqlParameter.Add("chkDate", chkAdjustDate.IsChecked == true ? 1 : 0);
                sqlParameter.Add("sFromDate", chkAdjustDate.IsChecked == true && dtpFromDate.SelectedDate != null ? dtpFromDate.SelectedDate.Value.ToString("yyyyMMdd") : "");
                sqlParameter.Add("sTODate", chkAdjustDate.IsChecked == true && dtpToDate.SelectedDate != null ? dtpToDate.SelectedDate.Value.ToString("yyyyMMdd") : "");
                sqlParameter.Add("ChkArticle", chkArticle.IsChecked == true ? 1 : 0);
                sqlParameter.Add("ArticleID", chkArticle.IsChecked == true && txtArticle.Tag != null && txtArticle.Text.Trim().Equals("") == false? txtArticle.Tag.ToString() : "");

                sqlParameter.Add("ChkLabelID", chkLotIDSrh.IsChecked == true ? 1 : 0);
                sqlParameter.Add("LabelID", chkLotIDSrh.IsChecked == true && !txtLotIDSrh.Text.Trim().Equals("") ? txtLotIDSrh.Text : "");

                sqlParameter.Add("ChkRecency", chkRecency.IsChecked == true ? 1 : 0);

                sqlParameter.Add("sToLocID", chkToLocSrh.IsChecked == true ? (cboWareHouse.SelectedValue != null ? cboWareHouse.SelectedValue.ToString() : "") : ""); // 창고

                DataSet ds = DataStore.Instance.ProcedureToDataSet("xp_sbStock_sLotStockControl", sqlParameter, false);

                if (ds != null && ds.Tables.Count > 0)
                {
                    DataTable dt = ds.Tables[0];

                    if (dt.Rows.Count == 0)
                    {
                        MessageBox.Show("조회된 데이터가 없습니다.");
                        Win_mtr_LotStockControl_U_CodeView Empty = new Win_mtr_LotStockControl_U_CodeView();
                        this.DataContext = Empty;
                        dgdMain.Items.Clear();
                        dgdSub.Items.Clear();
                    }
                    else
                    {
                        int i = 0;
                        DataRowCollection drc = dt.Rows;

                        foreach (DataRow dr in drc)
                        {
                            i++;
                            var LotStockControl = new Win_mtr_LotStockControl_U_CodeView
                            {
                                Num = i,
                                ControlID = dr["ControlID"].ToString(),
                                ControlDate = dr["ControlDate"].ToString(),
                                PersonID = dr["PersonID"].ToString(),
                                Name = dr["Name"].ToString(),
                                Comments = dr["Comments"].ToString(),
                            };

                            //조정일자 날짜서식
                            LotStockControl.ControlDate = DatePickerFormat(LotStockControl.ControlDate);

                            dgdMain.Items.Add(LotStockControl);
                        }
                    }
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
            dgdMain.SelectedIndex = rowNum;
        }

        //메인 그리드 
        private void DgdMain_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            var ControlItem = dgdMain.SelectedItem as Win_mtr_LotStockControl_U_CodeView;

            if (ControlItem != null
                && ControlItem.ControlID != null)
            {
                this.DataContext = ControlItem;
                FillGridSub(ControlItem.ControlID);
            }
            else
            {
                this.DataContext = null;
                dgdSub.Items.Clear();
            }
        }


        #endregion re_search, FillGrid, SelectionChanged

        #region 서브그리드 조회 메서드 - FillGridSub
        
        private void FillGridSub(string strID)
        {
            lstLotStock.Clear();

            if (dgdSub.Items.Count > 0)
            {
                dgdSub.Items.Clear();
            }

            try
            {
                Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                sqlParameter.Clear();

                sqlParameter.Add("chkDate", chkAdjustDate.IsChecked == true ? 1 : 0);
                sqlParameter.Add("sFromDate", chkAdjustDate.IsChecked == true && dtpFromDate.SelectedDate != null ?
                                              dtpFromDate.SelectedDate.Value.ToString("yyyyMMdd") : "");
                sqlParameter.Add("sTODate", chkAdjustDate.IsChecked == true && dtpToDate.SelectedDate != null ?
                                            dtpToDate.SelectedDate.Value.ToString("yyyyMMdd") : "");
                //sqlParameter.Add("ChkArticle", chkArticle.IsChecked == true ? 1 : 0);
                //sqlParameter.Add("ArticleID", chkArticle.IsChecked == true && txtArticle.Tag != null ?
                //                              txtArticle.Tag.ToString() : "");
                sqlParameter.Add("ChkArticle", 0);
                sqlParameter.Add("ArticleID", "");
                sqlParameter.Add("ChkControlID", 1); //ControlID는 무조건 있을 거니까
                sqlParameter.Add("ControlID", strID);


                DataSet ds = DataStore.Instance.ProcedureToDataSet("xp_sbStock_sLotStockControlSub", sqlParameter, false);

                if (ds != null && ds.Tables.Count > 0)
                {
                    DataTable dt = ds.Tables[0];

                    if (dt.Rows.Count == 0)
                    {
                        //MessageBox.Show("조회된 데이터가 없습니다.");
                        dgdSub.Items.Clear();
                    }
                    else
                    {
                        int i = 0;
                        DataRowCollection drc = dt.Rows;

                        foreach (DataRow dr in drc)
                        {
                            i++;
                            var StockControlSub = new Win_mtr_LotStockControl_U_CodeView
                            {
                                Num = i,
                                ControlID = dr["ControlID"].ToString(),
                                ControlSeq = dr["ControlSeq"].ToString(),
                                LotID = dr["LOTID"].ToString(),
                                LotName = dr["LOTID"].ToString(),
                                ArticleID = dr["ArticleID"].ToString(),
                                Article = dr["Article"].ToString(),
                                BuyerArticleNo = dr["BuyerArticleNo"].ToString(),
                                UnitClss = dr["UnitClss"].ToString(),
                                UnitClssName = dr["UnitclssName"].ToString(),
                                StockQty = stringFormatN0(dr["StockQty"]),
                                ControlQty = stringFormatN0(dr["ControlQty"]),
                                TOLocID = dr["TOLocID"].ToString(),
                                ToLocName = dr["ToLocName"].ToString(),
                                Comments = dr["Comments"].ToString(),
                                UDFlag = dr["UDFlag"].ToString().Trim().Equals("Y") ? true : false,
                                LastDate = dr["LastDate"].ToString(),
                            };

                            dgdSub.Items.Add(StockControlSub);

                            lstLotStock.Add(StockControlSub.Clone());
                        }
                    }
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

        #endregion

        #region 저장구문 SaveData

        //실제 저장
        private bool SaveData(string strFlag, string strID)
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

                    sqlParameter.Add("sControlID", txtControlID.Text);
                    sqlParameter.Add("ControlDate", dtpAdjustDate.SelectedDate != null ? dtpAdjustDate.SelectedDate.Value.ToString("yyyyMMdd") : "");
                    sqlParameter.Add("PersonID", txtWorker.Tag != null ? txtWorker.Tag.ToString() : "");
                    sqlParameter.Add("Comments", txtReason.Text);

                    #region 추가

                    if (strFlag.Equals("I"))
                    {
                        sqlParameter.Add("CreateUserID", MainWindow.CurrentUser);

                        Procedure pro1 = new Procedure();
                        pro1.Name = "xp_sbStock_iLotStockControl";
                        pro1.OutputUseYN = "Y";
                        pro1.OutputName = "sControlID";
                        pro1.OutputLength = "12";

                        Prolist.Add(pro1);
                        ListParameter.Add(sqlParameter);


                        //for문 타면서 서브그리드 저장
                        for (int i = 0; i < dgdSub.Items.Count; i++)
                        {
                            var LotStockControl = dgdSub.Items[i] as Win_mtr_LotStockControl_U_CodeView;

                            if (LotStockControl != null
                                && LotStockControl.LotID != null
                                && LotStockControl.LotID.Trim().Equals("") == false
                                && LotStockControl.StockQty != null)
                            {
                                //클리어만 해서는 소용이 없어, 이건 무조건 한 세트야
                                sqlParameter = new Dictionary<string, object>();
                                sqlParameter.Clear();

                                sqlParameter.Add("sControlID", "");
                                sqlParameter.Add("LotID", LotStockControl.LotID);
                                sqlParameter.Add("ArticleID", LotStockControl.ArticleID);
                                sqlParameter.Add("nStockQty", ConvertDouble(LotStockControl.StockQty));
                                sqlParameter.Add("nControlQty", ConvertDouble(LotStockControl.ControlQty));

                                sqlParameter.Add("ControlDate", dtpAdjustDate.SelectedDate != null ? dtpAdjustDate.SelectedDate.Value.ToString("yyyyMMdd") : "");
                                sqlParameter.Add("UnitClss", LotStockControl.UnitClss);
                                sqlParameter.Add("LocID", LotStockControl.TOLocID);
                                sqlParameter.Add("Comments", LotStockControl.Comments == null ? "실재고조사를 통해 재고조정" : LotStockControl.Comments); //2021-07-05 비고 추가
                                sqlParameter.Add("CreateUserID", MainWindow.CurrentUser);


                                Procedure pro2 = new Procedure();
                                pro2.Name = "xp_sbStock_iLotStockControlsub";
                                pro2.OutputUseYN = "N";
                                pro2.OutputName = "sControlID";
                                pro2.OutputLength = "12";

                                Prolist.Add(pro2);
                                ListParameter.Add(sqlParameter);
                            }
                        }

                        //동운씨가 만든 아웃풋 값 찾는 방법
                        List<KeyValue> list_Result = new List<KeyValue>();
                        list_Result = DataStore.Instance.ExecuteAllProcedureOutputGetCS(Prolist, ListParameter);

                        string sGetID = string.Empty;

                        if (list_Result[0].key.ToLower() == "success")
                        {
                            flag = true;
                        }
                        else
                        {
                            MessageBox.Show("[저장실패]\r\n" + list_Result[0].value.ToString());
                            flag = false;
                        }

                    }
                    #endregion 추가

                    #region 수정

                    else if (strFlag.Equals("U"))
                    {
                        sqlParameter.Add("UpdateUserID", MainWindow.CurrentUser);

                        Procedure pro1 = new Procedure();
                        pro1.Name = "xp_sbStock_uLotStockControl";
                        pro1.OutputUseYN = "N";
                        pro1.OutputName = "sControlID";
                        pro1.OutputLength = "12";

                        Prolist.Add(pro1);
                        ListParameter.Add(sqlParameter);

                        //for문 타면서 서브그리드 저장(메인 데이터 수정시, sub는 삭제 하므로 재추가함)
                        for (int i = 0; i < dgdSub.Items.Count; i++)
                        {
                            var LotStockControl = dgdSub.Items[i] as Win_mtr_LotStockControl_U_CodeView;

                            if (LotStockControl != null
                               && LotStockControl.LotID != null
                               && LotStockControl.LotID.Trim().Equals("") == false
                                && LotStockControl.StockQty != null)
                            {
                                //클리어만 해서는 소용이 없어, 이건 무조건 한 세트야
                                sqlParameter = new Dictionary<string, object>();
                                sqlParameter.Clear();

                                sqlParameter.Add("sControlID", txtControlID.Text);
                                sqlParameter.Add("LotID", LotStockControl.LotID);
                                sqlParameter.Add("ArticleID", LotStockControl.ArticleID);
                                sqlParameter.Add("nStockQty", ConvertDouble(LotStockControl.StockQty));
                                sqlParameter.Add("nControlQty", ConvertDouble(LotStockControl.ControlQty));

                                sqlParameter.Add("ControlDate", dtpAdjustDate.SelectedDate != null ? dtpAdjustDate.SelectedDate.Value.ToString("yyyyMMdd") : "");
                                sqlParameter.Add("UnitClss", LotStockControl.UnitClss);
                                sqlParameter.Add("LocID", LotStockControl.TOLocID);
                                sqlParameter.Add("Comments", LotStockControl.Comments == null ? "실재고조사를 통해 재고조정" : LotStockControl.Comments); //2021-07-05
                                sqlParameter.Add("CreateUserID", MainWindow.CurrentUser);

                                Procedure pro2 = new Procedure();
                                pro2.Name = "xp_sbStock_iLotStockControlsub";
                                pro2.OutputUseYN = "N";
                                pro2.OutputName = "sControlID";
                                pro2.OutputLength = "12";

                                Prolist.Add(pro2);
                                ListParameter.Add(sqlParameter);
                            }
                        }

                        string[] confirm = new string[2];
                        confirm = DataStore.Instance.ExecuteAllProcedureOutputNew(Prolist, ListParameter);

                        if (confirm[0] == "success")
                        {
                            //MessageBox.Show("성공");
                            flag = true;
                        }
                        else
                        {
                            MessageBox.Show("실패 : " + confirm[1]);
                            flag = false;
                        }
                    }
                    #endregion 수정
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

        #endregion

        #region 유효성 검사 CheckData

        //저장 전 입력 데이터 체크
        private bool CheckData()
        {
            bool flag = true;

            //조정일자
            if (dtpAdjustDate.SelectedDate == null)
            {
                MessageBox.Show("조정일자는 필수입력 항목입니다. 조정일자를 입력해주세요.");
                flag = false;
                return flag;
            }

            if (txtWorker.Text.Length <= 0 || txtWorker.Tag == null)
            {
                MessageBox.Show("작업자가 입력되지 않았습니다. 작업자를 입력해주세요.");
                flag = false;
                return flag;
            }

            if (txtReason.Text.Length <= 0)
            {
                MessageBox.Show("재고조정 사유는 필수입력 항목입니다. 재고조정 사유를 입력해주세요.");
                flag = false;
                return flag;
            }

            // 
            if (dgdSub.Items.Count == 0)
            {
                MessageBox.Show("오른쪽 하단의 재고조정 품목이 등록되지 않았습니다.");
                flag = false;
                return flag;
            }

            // 서브그리드 조정재고는 숫자만 입력 가능하도록
            for (int i = 0; i < dgdSub.Items.Count; i++)
            {
                var Sub = dgdSub.Items[i] as Win_mtr_LotStockControl_U_CodeView;
                if (Sub != null
                    && Sub.ControlQty != null
                    && Sub.LotID != null)
                {
                    if (CheckConvertDouble(Sub.ControlQty) == false)
                    {
                        MessageBox.Show("조정재고는 숫자만 입력 가능합니다.\r" + Sub.LotID + " 라벨을 확인해주세요.");
                        dgdSub.SelectedIndex = i;
                        flag = false;
                        return flag;
                    }
                }
            }

            return flag;
        }

        #endregion SaveData

        #region DeleteData

        private bool DeleteData(string OrderID)
        {
            bool flag = false;

            Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
            sqlParameter.Clear();
            sqlParameter.Add("sControlID", txtControlID.Text);

            try
            {
                string[] result = DataStore.Instance.ExecuteProcedure("xp_sbStock_dLotStockControl", sqlParameter, true);

                if (!result[0].Equals("success"))
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

        #endregion DeleteData

        //2021-10-15 삭제 전에 현재고를 체크하여 삭제 시 마이너스 재고만 생성되는 지 확인 하기 위해 생성
        #region DeleteDataCheck
        private bool DeleteDataCheck(string ArticleID, string LabelID, string ControlID)
        {
            bool flag = false;

            Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
            sqlParameter.Clear();
            sqlParameter.Add("sControlID", ControlID);
            sqlParameter.Add("sArticleID", ArticleID);
            sqlParameter.Add("sLabelID", LabelID);

            try
            {
                string[] result = DataStore.Instance.ExecuteProcedure("xp_sbStock_dLotStockControlCheck", sqlParameter, true);

                if (!result[0].Equals("success") || !result[1].Equals(""))
                {
                    if (MessageBox.Show(result[1], "로트별 재고 조정 삭제", MessageBoxButton.YesNo) == MessageBoxResult.Yes)
                    {
                        flag = true;
                    }
                    else
                    {
                        flag = false;
                    }
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
        #endregion

        //2021-10-18 MainGrid를 선택 후 삭제 시 체크 함수 생성
        #region DeleteDataMainCheck
        private bool DeleteDataMainCheck(string ArticleID, string LabelID, string ControlID)
        {
            bool flag = false;

            Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
            sqlParameter.Clear();
            sqlParameter.Add("sControlID", ControlID);
            sqlParameter.Add("sArticleID", ArticleID);
            sqlParameter.Add("sLabelID", LabelID);

            try
            {
                string[] result = DataStore.Instance.ExecuteProcedure("xp_sbStock_dLotStockControlCheck", sqlParameter, true);

                if (!result[0].Equals("success") || !result[1].Equals(""))
                {
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
        #endregion


        #region Content 부분 - 데이터 그리드 키 이벤트

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
            int currRow = dgdSub.Items.IndexOf(dgdSub.CurrentItem);
            int currCol = dgdSub.Columns.IndexOf(dgdSub.CurrentCell.Column);
            int startCol = 2;
            int endCol = 10;

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
                    dgdSub.CurrentCell = new DataGridCellInfo(dgdSub.Items[currRow], dgdSub.Columns[currCol + 1]);

                } // 마지막 열, 마지막 행
                else if (endCol == currCol && dgdSub.Items.Count - 1 == currRow)
                {

                }
                else
                {
                    MessageBox.Show("나머지가 있나..");
                }
            }
            else if (e.Key == Key.Down)
            {
                e.Handled = true;
                (sender as DataGridCell).IsEditing = false;

                // 마지막 행 아님
                if (dgdSub.Items.Count - 1 > currRow)
                {
                    dgdSub.SelectedIndex = currRow + 1;
                    dgdSub.CurrentCell = new DataGridCellInfo(dgdSub.Items[currRow + 1], dgdSub.Columns[currCol]);
                } // 마지막 행일때
                else if (dgdSub.Items.Count - 1 == currRow)
                {
                    if (endCol > currCol) // 마지막 열이 아닌 경우, 열을 오른쪽으로 이동
                    {
                        //dgdMain.SelectedIndex = 0;
                        dgdSub.CurrentCell = new DataGridCellInfo(dgdSub.Items[currRow], dgdSub.Columns[currCol + 1]);
                    }
                    else
                    {
                        //btnSave.Focus();
                    }
                }
            }
            else if (e.Key == Key.Up)
            {
                e.Handled = true;
                (sender as DataGridCell).IsEditing = false;

                // 첫행 아님
                if (currRow > 0)
                {
                    dgdSub.SelectedIndex = currRow - 1;
                    dgdSub.CurrentCell = new DataGridCellInfo(dgdSub.Items[currRow - 1], dgdSub.Columns[currCol]);
                } // 첫 행
                else if (dgdSub.Items.Count - 1 == currRow)
                {
                    if (0 < currCol) // 첫 열이 아닌 경우, 열을 왼쪽으로 이동
                    {
                        //dgdMain.SelectedIndex = 0;
                        dgdSub.CurrentCell = new DataGridCellInfo(dgdSub.Items[currRow], dgdSub.Columns[currCol - 1]);
                    }
                    else
                    {
                        //btnSave.Focus();
                    }
                }
            }
            else if (e.Key == Key.Left)
            {
                e.Handled = true;
                (sender as DataGridCell).IsEditing = false;

                if (startCol < currCol)
                {
                    dgdSub.CurrentCell = new DataGridCellInfo(dgdSub.Items[currRow], dgdSub.Columns[currCol - 1]);
                }
                else if (startCol == currCol)
                {
                    if (0 < currRow)
                    {
                        dgdSub.SelectedIndex = currRow - 1;
                        dgdSub.CurrentCell = new DataGridCellInfo(dgdSub.Items[currRow - 1], dgdSub.Columns[endCol]);
                    }
                    else
                    {
                        //btnSave.Focus();
                    }
                }
            }
            else if (e.Key == Key.Right)
            {
                e.Handled = true;
                (sender as DataGridCell).IsEditing = false;

                if (endCol > currCol)
                {

                    dgdSub.CurrentCell = new DataGridCellInfo(dgdSub.Items[currRow], dgdSub.Columns[currCol + 1]);
                }
                else if (endCol == currCol)
                {
                    if (dgdSub.Items.Count - 1 > currRow)
                    {
                        dgdSub.SelectedIndex = currRow + 1;
                        dgdSub.CurrentCell = new DataGridCellInfo(dgdSub.Items[currRow + 1], dgdSub.Columns[startCol]);
                    }
                    else
                    {
                        //btnSave.Focus();
                    }
                }
            }

        }
        // KeyUp 이벤트
        private void DataGridIn_TextFocus(object sender, KeyEventArgs e)
        {
            // 엔터 → 포커스 = true → cell != null → 해당 텍스트박스가 null이 아니라면 
            // → 해당 텍스트박스가 포커스가 안되있음 SelectAll() or 포커스
            Lib.Instance.DataGridINTextBoxFocus(sender, e);
        }
        // GotFocus 이벤트
        private void DataGridCell_GotFocus(object sender, RoutedEventArgs e)
        {
            if (btnSave.Visibility == Visibility.Visible)
            {
                int currCol = dgdSub.Columns.IndexOf(dgdSub.CurrentCell.Column);

                DataGridCell cell = sender as DataGridCell;

                if ((currCol == 2)
                    || currCol == 9
                    || currCol == 10)
                {
                    // 수정 시, 이전 기록이면 수정이 불가능 하도록
                    if (strFlag.Trim().Equals("U"))
                    {
                        var Sub = dgdSub.SelectedItem as Win_mtr_LotStockControl_U_CodeView;
                        if (Sub != null
                            && Sub.UDFlag == false)
                        {
                            
                            cell.IsEditing = false;
                            return;
                        }
                    }

                    //DataGridCell cell = sender as DataGridCell;
                    cell.IsEditing = true;
                }
            }
        }
        // 2019.08.27 MouseUp 이벤트
        private void DataGridCell_MouseUp(object sender, MouseButtonEventArgs e)
        {
            Lib.Instance.DataGridINTextBoxFocusByMouseUP(sender, e);
        }

        #endregion // Content 부분 - 데이터 그리드 키 이벤트

        #region 기타 매서드

        // 천 단위 콤마, 소수점 버리기
        private string stringFormatN0(object obj)
        {
            return string.Format("{0:N0}", obj);
        }

        // 천 단위 콤마, 소수점 두자리
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

        //숫자 외에 다른 문자열 못들어오도록
        public bool IsNumeric(string source)
        {

            Regex regex = new Regex("[^0-9.-]+");
            return !regex.IsMatch(source);
        }

        //나눗셈, 분모가 0이면 0값 반환
        private double division(double a, double b)
        {
            if (b == 0)
            {
                return 0;
            }
            else
            {
                return a / b;
            }
        }


        #endregion 기타 매서드

        #region Content - 서브그리드 관련 메서드

        // 대상선택 클릭 
        private void btnChoice_Click(object sender, RoutedEventArgs e)
        {
            Win_pop_Stock_LotNo_2 Lot = new Win_pop_Stock_LotNo_2(lstLotStock);

            Lot.date = dtpAdjustDate.SelectedDate != null ? dtpAdjustDate.SelectedDate.Value.ToString("yyyyMMdd") : "";
            

            Lot.ShowDialog();

            // 중복되는 라벨이 있을 경우 메시지 띄워주기 위한 변수
            string Msg = "";

            string MsgAll = "";

            if (Lot.DialogResult == true)
            {
                // 중복을 제외하고 몇개가 들어가는지 확인하는 변수
                int count = 0;
                string ChgDate = dtpAdjustDate.SelectedDate != null ? dtpAdjustDate.SelectedDate.Value.ToString("yyyy-MM-dd") : "";

                for (int i = 0; i < Lot.lstLotStock.Count; i++)
                {
                    var main = Lot.lstLotStock[i] as Win_mtr_LotStockControl_U_CodeView;
                    
                    if (main != null
                        && main.LotID != null)
                    {
                        


                        if (CheckIsLabel(main.LotID,main.ArticleID, false) == false)
                        {
                            Msg += main.LotID + "\r";
                            continue;
                        }
                        else
                        {
                            //2021-07-05 최근 작업 내역 안보이게
                            //if (main.LastDate != null
                            //    && !main.LastDate.Trim().Equals(""))
                            //{
                            //    MsgAll += main.LotID + "라벨, 최근 조정내역 : " + DatePickerFormat(main.LastDate) + "\r";
                            //}



                            main.Num = dgdSub.Items.Count;
                            dgdSub.Items.Add(main);
                            count++;
                        }
                    }

                }


                if (Msg.Length > 0)
                {
                    Msg += "위의 라벨은 이미 등록되어 있습니다.";
                    if (count != 0) { Msg += "\r(위의 라벨을 제외하고 추가되었습니다.)"; }
                    MessageBox.Show(Msg);
                }

                if(MsgAll.Length >0)
                {
                    MsgAll += "해당 라벨을 재고 조정 시 "+ ChgDate+ " 이후의 데이터는 재고 수량이 정확하지 않으니\r기존에 조정한 재고를 다시 작업하시기 바랍니다.";
                    MessageBox.Show(MsgAll);

                }
            }

            setNumSubDgd();
        }

        // 서브 그리드 라벨 엔터 → 해당 정보 가져오기
        private void txtLotID_KeyDown(object sender, KeyEventArgs e)
        {

            if (e.Key == Key.Enter)
            {
                var Sub = dgdSub.SelectedItem as Win_mtr_LotStockControl_U_CodeView;
                if (Sub != null)
                {
                    if (CheckIsLabel(Sub.LotID, true) == false)
                    {
                        MessageBox.Show("해당 라벨은 이미 등록되어 있습니다.");
                        return;
                    }

                    Win_pop_Stock_LotNoPF Lot = new Win_pop_Stock_LotNoPF(Sub.LotName, lstLotStock);

                    Lot.date = dtpAdjustDate.SelectedDate != null ? dtpAdjustDate.SelectedDate.Value.ToString("yyyyMMdd") : "";

                    Lot.ShowDialog();


                    if (Lot.DialogResult == true)
                    {
                        if (CheckIsLabel(Lot.Stock.LotID, true) == false)
                        {
                            MessageBox.Show(Lot.Stock.LotID + " 라벨은 이미 등록되어 있습니다.");
                            Sub.LotName = "";
                            return;
                        }
                        else
                        {
                            if(Lot.Stock.LastDate != null && !Lot.Stock.LastDate.Trim().Equals(""))
                            {
                                string ChgDate = dtpAdjustDate.SelectedDate != null ? dtpAdjustDate.SelectedDate.Value.ToString("yyyy-MM-dd") : "";


                                MessageBox.Show(Lot.Stock.LotID + " 라벨은" + ChgDate + " 이후 \r최근에 " + DatePickerFormat(Lot.Stock.LastDate) + "에 조정한 내역이 있습니다."
                                    + "\r해당 라벨을 재고 조정 시 " + ChgDate + " 이후의 데이터는 재고 수량이 정확하지 않으니 기존에 조정한 재고를 다시 작업하시기 바랍니다."

                                    , "재작업 요청", MessageBoxButton.OK);


                                
                             }

                            Sub.Copy(Lot.Stock);

                        }
                    }
                }
            }


        }

        // 중복으로 라벨 등록하는걸 막기 위한 체크 이벤트
        // → 선택된 그 라벨은 제외 하고 검색을 해야 됨
        // ExcptSelLot : true (지금 서브 그리드 선택된 행의 LotID 를 제외 하고)
        // ExcptSelLot : false (지금 서브 그리드 선택된 행의 LotID 를 포함 해서)
        private bool CheckIsLabel(string LableID, bool ExcptSelLot)
        {
            bool flag = true;

            string SelLotID = "";

            // 지금 활성화된 라벨
            var LotSub = dgdSub.SelectedItem as Win_mtr_LotStockControl_U_CodeView;
            if (LotSub != null)
            {
                SelLotID = LotSub.LotID;
            }

            for (int i = 0; i < dgdSub.Items.Count; i++)
            {
                var Sub = dgdSub.Items[i] as Win_mtr_LotStockControl_U_CodeView;
                if (Sub != null
                    && Sub.LotID != null
                    && !Sub.LotID.Trim().Equals(""))
                {
                    if (ExcptSelLot == true
                        && SelLotID.Equals("") == false
                        && Sub.LotID.Equals(SelLotID))
                    {
                        continue;
                    }

                    if (Sub.LotID.ToUpper().Trim().Equals(LableID.ToUpper().Trim()))
                    {
                        flag = false;
                        break;
                    }
                }
            }

            return flag;
        }

        #endregion
        //2021-06-26 라벨은 같은 품명이 다를 경우를 대비하여 추가
        private bool CheckIsLabel(string LableID, string ArticleID ,bool ExcptSelLot)
        {
            bool flag = true;

            string SelLotID = "";

            // 지금 활성화된 라벨
            var LotSub = dgdSub.SelectedItem as Win_mtr_LotStockControl_U_CodeView;
            if (LotSub != null)
            {
                SelLotID = LotSub.LotID;
            }

            for (int i = 0; i < dgdSub.Items.Count; i++)
            {
                var Sub = dgdSub.Items[i] as Win_mtr_LotStockControl_U_CodeView;
                if (Sub != null
                    && Sub.LotID != null
                    && !Sub.LotID.Trim().Equals(""))
                {
                    if (ExcptSelLot == true
                        && SelLotID.Equals("") == false
                        && Sub.LotID.Equals(SelLotID))
                    {
                        continue;
                    }

                    if (Sub.LotID.ToUpper().Trim().Equals(LableID.ToUpper().Trim()) && Sub.ArticleID.ToUpper().Trim().Equals(ArticleID.ToUpper().Trim()))
                    {
                        flag = false;
                        break;
                    }
                }
            }

            return flag;
        }


        private void dgdControlQty_PreviewTextInput(object sender, TextCompositionEventArgs e)
        {
            Lib.Instance.CheckIsNumeric((TextBox)sender, e);
        }



        private void DtpAdjustDate_SelectedDateChanged(object sender, SelectionChangedEventArgs e)
        {
            if(EventLabel.Visibility == Visibility.Visible)
            {
                if (dgdSub.Items.Count > 0 && dtpAdjustDate.SelectedDate != null )
                {
                    if (MessageBox.Show("조정일자 변경시 해당 조정일자로 넣은 현재고 수량이 일치하지 않습니다.\r조정품목을 모두 초기화 하시겠습니까?", "초기화 전 확인", MessageBoxButton.YesNo) == MessageBoxResult.Yes)
                    {
                        dgdSub.Items.Clear();
                    }
                }
            }
            
        }


        //창고 체크
        private void chkToLocSrh_Click(object sender, RoutedEventArgs e)
        {
            if (chkToLocSrh.IsChecked == true)
            {
                cboWareHouse.IsEnabled = true;
                cboWareHouse.Focus();
            }
            else
            {
                cboWareHouse.IsEnabled = false;
            }
        }

        private void chkToLocSrh_Click(object sender, MouseButtonEventArgs e)
        {
            if (chkToLocSrh.IsChecked == true)
            {
                chkToLocSrh.IsChecked = false;
                cboWareHouse.IsEnabled = false;

            }
            else
            {
                chkToLocSrh.IsChecked = true;
                cboWareHouse.IsEnabled = true;
                cboWareHouse.Focus();
            }
        }
    }

    #region CodeView 

    public class Win_mtr_LotStockControl_U_CodeView : BaseView
    {
        public int Num { get; set; }

        public bool Chk { get; set; }

        public string ControlID { get; set; }
        public string ControlDate { get; set; }
        public string PersonID { get; set; }
        public string Name { get; set; }
        public string Comments { get; set; }
        public string StuffinQty { get; set; }
        public string OutQty { get; set; }

        public string ControlSeq { get; set; }
        public string LotID { get; set; }
        public string LotName { get; set; }
        public string ArticleID { get; set; }
        public string Article { get; set; }
        public string BuyerArticleNo { get; set; }
        public string StockQty { get; set; }
        public string ControlQty { get; set; }
        public string UnitClss { get; set; }
        public string UnitClssName { get; set; }
        public string StuffINID { get; set; }

        public string CreateDate { get; set; }
        public string CreateUserID { get; set; }
        public string LastUpdateDate { get; set; }
        public string LastUpdateUserID { get; set; }

        public string ArticleGrpID { get; set; }
        public string ArticleGrp { get; set; }
        public string TOLocID { get; set; }
        public string ToLocName { get; set; }

        // 수정, 삭제가 가능한지 체크하는 변수
        public bool UDFlag { get; set; }
        public string LastDate { get; set; } // 이 라벨로 마지막에 조정재고 등록한 날짜가 언제인지

        public Win_mtr_LotStockControl_U_CodeView Clone ()
        {
            return (Win_mtr_LotStockControl_U_CodeView)this.MemberwiseClone();
        }

        public void Copy(Win_mtr_LotStockControl_U_CodeView LotStock)
        {
            this.TOLocID = LotStock.TOLocID;
            this.ToLocName = LotStock.ToLocName;
            this.LotID = LotStock.LotID;
            this.LotName = LotStock.LotID;
            this.BuyerArticleNo = LotStock.BuyerArticleNo;
            this.Article = LotStock.Article;
            this.ArticleID = LotStock.ArticleID;
            this.UnitClss = LotStock.UnitClss;
            this.UnitClssName = LotStock.UnitClssName;
            this.StockQty = LotStock.StockQty;
        }

    }




    #endregion CodeView 
}

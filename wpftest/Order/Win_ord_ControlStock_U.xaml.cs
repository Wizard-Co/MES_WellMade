using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Data;
using System.Text.RegularExpressions;
using System.Threading;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using WizMes_WellMade.PopUP;
using WPF.MDI;

namespace WizMes_WellMade
{
    /// <summary>
    /// Win_ord_ControlStock_U.xaml에 대한 상호 작용 논리
    /// </summary>
    public partial class Win_ord_ControlStock_U : UserControl
    {
        Lib lib = new Lib();
        PlusFinder pf = new PlusFinder();

        // 추가저장인지 / 수정저장인지 구별하는 용도입니다.
        string ButtonTag = string.Empty;
        // 추가/수정 시 저장취소와 연계되어 value를 자동으로 찾아 select해 줄 value list 값입니다.
        List<string> lstCompareValue = new List<string>();


        int TotalStockQty = 0;



        public Win_ord_ControlStock_U()
        {
            InitializeComponent();
        }

        // 로딩.
        private void Win_ord_ControlStock_Loaded(object sender, RoutedEventArgs e)
        {
            First_Step();
            ComboBoxSetting();
        }

        #region 첫단계 / 날짜버튼 컨트롤 / 체크박스클릭 컨트롤.

        // 첫 단계.
        private void First_Step()
        {
            //GLS의 기초재고 등록 기준일은 1월 31일로 설정
            dtpFromDate.Text = "2020-01-31";
            dtpToDate.Text = "2020-01-31";

            //dtpToDate.Text = DateTime.Now.ToString("yyyy-MM-dd");

            txtCustomer.IsEnabled = false;
            txtArticle.IsEnabled = false;
            btnCustomer.IsEnabled = false;
            btnArticle.IsEnabled = false;

            btnSave.Visibility = Visibility.Hidden;
            btnCancel.Visibility = Visibility.Hidden;

            grbStockBox.IsEnabled = false;
            grbStockDatePickerBox.Visibility = Visibility.Hidden;
            EventLabel.Visibility = Visibility.Hidden;

        }

        //오늘
        private void btn_Today(object sender, RoutedEventArgs e)
        {
            dtpFromDate.Text = DateTime.Now.ToString("yyyy-MM-dd");
            dtpToDate.Text = DateTime.Now.ToString("yyyy-MM-dd");
        }

        //이번 달.
        private void btn_ThisMonth_Click(object sender, RoutedEventArgs e)
        {
            string[] receiver = lib.BringThisMonthDatetime();

            dtpFromDate.Text = receiver[0];
            dtpToDate.Text = receiver[1];
        }

        //라디오 버튼 (월)말일 클릭 시.
        private void rbnLastDay_Click(object sender, RoutedEventArgs e)
        {
            string[] receiver = lib.BringThisMonthLastDatetime();

            dtpStandardDate.Text = receiver[0];
        }

        // 라디오 버튼 현재일 클릭 시.
        private void rbnToDay_Click(object sender, RoutedEventArgs e)
        {
            dtpStandardDate.Text = DateTime.Now.ToString("yyyy-MM-dd");
        }


        //거래처 클릭.
        private void chkCustomer_Click(object sender, RoutedEventArgs e)
        {
            if (chkCustomer.IsChecked == true)
            {
                txtCustomer.IsEnabled = true;
                txtCustomer.Focus();
                btnCustomer.IsEnabled = true;
            }
            else
            {
                txtCustomer.IsEnabled = false;
                btnCustomer.IsEnabled = false;
            }
        }
        //거래처 클릭.
        private void chkCustomer_Click(object sender, MouseButtonEventArgs e)
        {
            if (chkCustomer.IsChecked == true)
            {
                chkCustomer.IsChecked = false;
                txtCustomer.IsEnabled = false;
                btnCustomer.IsEnabled = false;
            }
            else
            {
                chkCustomer.IsChecked = true;
                txtCustomer.IsEnabled = true;
                txtCustomer.Focus();
                btnCustomer.IsEnabled = true;
            }
        }
        // 품명 클릭.
        private void chkArticle_Click(object sender, RoutedEventArgs e)
        {
            if (chkArticle.IsChecked == true)
            {
                txtArticle.IsEnabled = true;
                txtArticle.Focus();
                btnArticle.IsEnabled = true;
            }
            else
            {
                txtArticle.IsEnabled = false;
                btnArticle.IsEnabled = false;
            }
        }
        // 품명 클릭.
        private void chkArticle_Click(object sender, MouseButtonEventArgs e)
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
                txtArticle.Focus();
                btnArticle.IsEnabled = true;
            }
        }

        //그룹박스 내 거래처 클릭
        private void chkGbCustomer_Click(object sender, RoutedEventArgs e)
        {
            if (chkGbCustomer.IsChecked == true)
            {
                txtGbCustomer.IsEnabled = true;
                txtGbCustomer.Focus();
                btnGbCustomer.IsEnabled = true;
            }
            else
            {
                txtGbCustomer.IsEnabled = false;
                btnGbCustomer.IsEnabled = false;
            }
        }
        // 그룹박스 내 거래처 클릭.
        private void chkGbCustomer_Click(object sender, MouseButtonEventArgs e)
        {
            if (chkGbCustomer.IsChecked == true)
            {
                chkGbCustomer.IsChecked = false;
                txtGbCustomer.IsEnabled = false;
                btnGbCustomer.IsEnabled = false;
            }
            else
            {
                chkGbCustomer.IsChecked = true;
                txtGbCustomer.IsEnabled = true;
                txtGbCustomer.Focus();
                btnGbCustomer.IsEnabled = true;
            }
        }

        #endregion


        #region 콤보박스 세팅

        // 콤보박스 세팅.
        private void ComboBoxSetting()
        {
            cboWareHouse.Items.Clear();
            cboUnitClss.Items.Clear();

            ObservableCollection<CodeView> cbWareHouse = ComboBoxUtil.Instance.Gf_DB_CM_GetComCodeDataset(null, "LOC", "Y", "", "");
            ObservableCollection<CodeView> cbUnitClss = ComboBoxUtil.Instance.Gf_DB_CM_GetComCodeDataset(null, "MTRUNIT", "Y", "", "");

            this.cboWareHouse.ItemsSource = cbWareHouse;
            this.cboWareHouse.DisplayMemberPath = "code_name";
            this.cboWareHouse.SelectedValuePath = "code_id";
            this.cboWareHouse.SelectedIndex = 0;

            this.cboUnitClss.ItemsSource = cbUnitClss;
            this.cboUnitClss.DisplayMemberPath = "code_name";
            this.cboUnitClss.SelectedValuePath = "code_id";
            this.cboUnitClss.SelectedIndex = 0;

        }

        #endregion


        #region 플러스파인더

        // 플러스 파인더 세팅_ 거래처
        private void btnCustomer_Click(object sender, RoutedEventArgs e)
        {
            pf.ReturnCode(txtCustomer, 0, "");
        }

        // 플러스 파인더 엔터 _ 거래처
        private void TxtCustomer_KeyDown(object sender, KeyEventArgs e)
        {
            //pf.ReturnCode(txtCustomer, 0, "");

            if (e.Key == Key.Enter)
            {
                FillGrid();

                if (dgdStock.Items.Count > 0)
                {
                    dgdStock.SelectedIndex = 0;
                }
            }
        }

        // 플러스 파인더 세팅_ 품명
        private void btnArticle_Click(object sender, RoutedEventArgs e)
        {
            pf.ReturnCode(txtArticle, 1, "");
        }

        // 플러스 파인더 엔터_ 거래처
        private void TxtArticle_KeyDown(object sender, KeyEventArgs e)
        {
            //pf.ReturnCode(txtArticle, 1, "");

            if (e.Key == Key.Enter)
            {
                FillGrid();

                if (dgdStock.Items.Count > 0)
                {
                    dgdStock.SelectedIndex = 0;
                }
            }
        }

        // 플러스 파인더 세팅. _ 그룹박스 내 거래처
        private void btnGbCustomer_Click(object sender, RoutedEventArgs e)
        {
            if ((txtGbArticle.Tag != null) && (txtGbArticle.Text.Length > 0))
            {
                pf.ReturnCode(txtGbCustomer, 65, txtGbArticle.Tag.ToString());
            }
            else { pf.ReturnCode(txtGbCustomer, 0, ""); }

            lib.SendK(Key.Tab, this);
        }

        // 플러스 파인더 세팅. _ 그룹박스 내 품명
        private void btnGbArticle_Click(object sender, RoutedEventArgs e)
        {
            if ((txtGbCustomer.Tag != null) && (txtGbCustomer.Text.Length > 0) && (chkGbCustomer.IsChecked == true))
            {
                pf.ReturnCode(txtGbArticle, 85, txtGbCustomer.Tag.ToString());  //거래처별 품번 검색
            }
            else
            {
                pf.ReturnCode(txtGbArticle, 84, "");        //그냥 품번 검색
            }

            if (txtGbArticle.Text.Length > 0)
            {
                txtGbArticleID.Text = txtGbArticle.Tag.ToString();
            }
            lib.SendK(Key.Tab, this);
        }

        #endregion


        #region 공통 버튼이벤트

        //공통 사용가능
        private void PublicEnableTrue()
        {
            btnSave.Visibility = Visibility.Hidden;
            btnCancel.Visibility = Visibility.Hidden;
            EventLabel.Visibility = Visibility.Hidden;

            btnSearch.IsEnabled = true;
            btnAdd.IsEnabled = true;
            btnEdit.IsEnabled = true;
            btnDelete.IsEnabled = true;
            btnExcel.IsEnabled = true;

            grbStockBox.IsEnabled = false;
            grbStockDatePickerBox.Visibility = Visibility.Hidden;
            //그룹박스 내 거래처 체크했을지도, 풀어버려.
            chkGbCustomer.IsChecked = false;

            // 수정시 못건드리게 했던 그룹박스 내부 컨트롤, 쓸수있게 변경.
            dtpStandardDate.IsEnabled = true;
            chkGbCustomer.IsEnabled = true;
            txtGbArticle.IsEnabled = true;
            btnGbArticle.IsEnabled = true;

        }

        // 공통 버튼이벤트.
        private void PublicEnableFalse()
        {
            btnSave.Visibility = Visibility.Visible;
            btnCancel.Visibility = Visibility.Visible;
            EventLabel.Visibility = Visibility.Visible;

            btnSearch.IsEnabled = false;
            btnAdd.IsEnabled = false;
            btnEdit.IsEnabled = false;
            btnDelete.IsEnabled = false;
            btnExcel.IsEnabled = false;

            grbStockBox.IsEnabled = true;
            grbStockDatePickerBox.Visibility = Visibility.Visible;
            // 그룹박스 내 거레처는 여전히 false.
            txtGbCustomer.IsEnabled = false;
            btnGbCustomer.IsEnabled = false;

        }

        #endregion


        #region 조회 / 조회용 프로시저 

        // 검색. // 조회.
        private void btnSearch_Click(object sender, RoutedEventArgs e)
        {
            //검색버튼 비활성화
            btnSearch.IsEnabled = false;

            Dispatcher.BeginInvoke(new Action(() =>

            {
                Thread.Sleep(2000);

                //로직
                FillGrid();

            }), System.Windows.Threading.DispatcherPriority.Background);



            Dispatcher.BeginInvoke(new Action(() =>

            {
                btnSearch.IsEnabled = true;

            }), System.Windows.Threading.DispatcherPriority.Background);



        }

        private void FillGrid()
        {
            int totalCnt = 0;
            int totalQty = 0;

            if (dgdStock.Items.Count > 0)
            {
                dgdStock.Items.Clear();
            }
            if (dgdSum.Items.Count > 0)
            {
                dgdSum.Items.Clear();
            }

            try
            {

                Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                sqlParameter.Add("StartDate", dtpFromDate.ToString().Substring(0, 10).Replace("-", ""));
                sqlParameter.Add("EndDate", dtpToDate.ToString().Substring(0, 10).Replace("-", ""));
                sqlParameter.Add("ChkCustom", chkCustomer.IsChecked == true ? 1 : 0);
                sqlParameter.Add("Custom", txtCustomer.Text);
                sqlParameter.Add("ChkBuyerArticleNo", chkArticle.IsChecked == true ? 1 : 0);
                //sqlParameter.Add("BuyerArticleNo", @Escape(txtArticle.Text));
                sqlParameter.Add("BuyerArticleNo", txtArticle.Text);
                DataSet ds = DataStore.Instance.ProcedureToDataSet("xp_Subul_sbStock", sqlParameter, false);

                if (ds != null && ds.Tables.Count > 0)
                {
                    DataTable dt = null;
                    dt = ds.Tables[0];

                    if (dt.Rows.Count == 0)
                    {
                        MessageBox.Show("조회결과가 없습니다.");

                        this.DataContext = null;

                        return;
                    }

                    else
                    {
                        DataRowCollection drc = dt.Rows;
                        int i = 1;
                        foreach (DataRow item in drc)
                        {
                            var Win_ord_ControlStock_U_Insert = new Win_ord_ControlStock_U_View()
                            {
                                BasisDate = item["BasisDate"].ToString().Substring(4, 2) + "/" + item["BasisDate"].ToString().Substring(6, 2),
                                Full_BasisDate = DateTime.ParseExact(item["BasisDate"].ToString(), "yyyyMMdd", null).ToString("yyyy-MM-dd"),
                                CustomID = item["CustomID"].ToString(),
                                ArticleID = item["ArticleID"].ToString(),
                                ProcInClss = item["ProcInClss"].ToString(),
                                StockClss = item["StockClss"].ToString(),
                                StockQty = item["StockQty"].ToString(),
                                StockUnitClss = item["StockUnitClss"].ToString(),
                                WeightPerYard = item["WeightPerYard"].ToString(),
                                LocID = item["LocID"].ToString(),
                                KCustom = item["KCustom"].ToString(),
                                BuyerArticleNo = item["BuyerArticleNo"].ToString(),
                                Article = item["Article"].ToString(),
                                LocName = item["LocName"].ToString(),
                                UnitClssName = item["UnitClssName"].ToString()

                            };

                            Win_ord_ControlStock_U_Insert.StockQty = stringFormatN0(item["StockQty"]);


                            totalCnt = dt.Rows.Count;
                            totalQty += ConvertInt(item["StockQty"].ToString());

                            //MessageBox.Show(": " + TotalStockQty);

                            dgdStock.Items.Add(Win_ord_ControlStock_U_Insert);
                        }

                        var Total = new Win_ord_ControlStock_U_ViewSum()
                        {
                            TotalCnt = stringFormatN0(totalCnt),
                            TotalQty = stringFormatN0(totalQty),
                        };

                        dgdSum.Items.Clear();
                        dgdSum.Items.Add(Total);

                        txtTotalStockQty.Text = "총 수량 합계 : " + stringFormatN0(TotalStockQty) + "개";
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


        // 그리드 셀 클릭 시.
        private void dgdStock_SeletionChanged(object sender, SelectionChangedEventArgs e)
        {
            this.DataContext = dgdStock.SelectedItem as Win_ord_ControlStock_U_View;

            // 날짜.
            var ViewReceiver = dgdStock.SelectedItem as Win_ord_ControlStock_U_View;

            if (ViewReceiver != null) { dtpStandardDate.Text = ViewReceiver.Full_BasisDate; }
        }


        #region (추가 / 수정 / 삭제 / 저장 / 취소)  버튼 이벤트모음

        // 추가 버튼 선택시.
        private void btnAdd_Click(object sender, RoutedEventArgs e)
        {
            ButtonTag = ((Button)sender).Tag.ToString();

            // 1. 검색을 먼저 했을수도 있자나.. 일단 패널 클리어.
            GBStockBoxClear();

            // 2. 그리드 영향은 더 받지 마.
            //dgdStock.IsEnabled = false;
            dgdStock.IsHitTestVisible = false;

            //3. (추가)버튼 이벤트 시작.
            PublicEnableFalse();
            EventLabel.Content = "자료 입력(추가) 중..";

            //4. 기준일.
            dtpStandardDate.Text = DateTime.Now.ToString("yyyy-MM-dd");
            dtpStandardDate.Focus();
        }


        // 수정 버튼 클릭 시.
        private void btnEdit_Click(object sender, RoutedEventArgs e)
        {
            // 1. 수정할 자격은 있는거야? 조회? 데이터 선택??
            if (dgdStock.Items.Count < 1)
            {
                MessageBox.Show("먼저 검색해 주세요.");
                return;
            }
            var OBJ = dgdStock.SelectedItem as Win_ord_ControlStock_U_View;
            if (OBJ == null)
            {
                MessageBox.Show("수정할 항목이 정확히 선택되지 않았습니다.");
                return;
            }

            ButtonTag = ((Button)sender).Tag.ToString();
            EventLabel.Content = "자료입력(수정) 중..";

            // 2. 그리드 영향은 더 받지 마.
            //dgdStock.IsEnabled = false;
            dgdStock.IsHitTestVisible = false;

            //3. (수정)버튼 이벤트 시작.
            PublicEnableFalse();

            // 4. 여기서! 수정은 못건드리게 하는 부분이 더 있으니깐!!. (초재고니깐.)
            dtpStandardDate.IsEnabled = false;
            grbStockDatePickerBox.Visibility = Visibility.Hidden;
            chkGbCustomer.IsEnabled = false;
            txtGbArticle.IsEnabled = false;
            btnGbArticle.IsEnabled = false;
        }


        // 삭제 버튼 클릭.
        private void btnDelete_Click(object sender, RoutedEventArgs e)
        {
            // 1. 삭제할 자격은 있는거야? 조회? 데이터 선택??
            if (dgdStock.Items.Count < 1)
            {
                MessageBox.Show("먼저 검색해 주세요.");
                return;
            }
            var OBJ = dgdStock.SelectedItem as Win_ord_ControlStock_U_View;
            if (OBJ == null)
            {
                MessageBox.Show("삭제할 항목이 정확히 선택되지 않았습니다.");
                return;
            }

            MessageBoxResult msgresult = MessageBox.Show("선택하신 항목을 삭제하시겠습니까?", "삭제 확인", MessageBoxButton.YesNo);
            if (msgresult == MessageBoxResult.Yes)
            {
                // 2.  삭제용
                DeleteData();
                dgdStock.Refresh();

                // 3. 화면정리.
                btnCancel_Click(null, null);
            }
        }


        // 저장 버튼 클릭 시.
        private void btnSave_Click(object sender, RoutedEventArgs e)
        {
            // 1. 데이터 기입체크.(항목을 제대로 모두 똑바로 넣고 저장버튼을 누르는 거야??) 
            if (GBStockBoxDataCheck() == false) { return; }

            // 2. 저장.
            SaveData(ButtonTag);

            // 3. 화면정리.
            btnCancel_Click(null, null);
        }


        // 취소 버튼 클릭 시.
        private void btnCancel_Click(object sender, RoutedEventArgs e)
        {


            //1. 취소했으니 그룹박스 내용 클리어.
            GBStockBoxClear();

            //2. 그리드 다시사용 가능
            //dgdStock.IsEnabled = true;
            dgdStock.IsHitTestVisible = true;

            //3. 공통버튼 사용 가능.
            PublicEnableTrue();

            //4. 바인딩 값 엉망이 됬을지도.. 재 조회.
            FillGrid();

            //5. 추가/수정 이후 저장.취소시 Target 자동 세팅.
            if (ButtonTag != string.Empty)
            {
                int ReturnCount = lib.reTrunIndex(dgdStock, lstCompareValue);
                if (dgdStock.Items.Count > 0)
                {
                    object item = dgdStock.Items[ReturnCount];
                    dgdStock.SelectedItem = item;
                    dgdStock.ScrollIntoView(item);
                    DataGridRow row = dgdStock.ItemContainerGenerator.ContainerFromIndex(ReturnCount) as DataGridRow;
                    if (row != null)
                    {
                        row.MoveFocus(new TraversalRequest(FocusNavigationDirection.Next));
                    }
                }

            }
            ButtonTag = string.Empty;
            this.Focusable = true;
            this.Focus();
        }

        //전일
        private void btnYesterDay_Click(object sender, RoutedEventArgs e)
        {
            DateTime[] SearchDate = lib.BringLastDayDateTimeContinue(dtpFromDate.SelectedDate.Value);

            dtpFromDate.SelectedDate = SearchDate[0];
            dtpToDate.SelectedDate = SearchDate[1];
        }

        //금일
        private void btnToday_Click(object sender, RoutedEventArgs e)
        {
            dtpFromDate.SelectedDate = DateTime.Today;
            dtpFromDate.SelectedDate = DateTime.Today;
        }

        // 전월 버튼 클릭 이벤트
        private void btnLastMonth_Click(object sender, RoutedEventArgs e)
        {
            DateTime[] SearchDate = lib.BringLastMonthContinue(dtpFromDate.SelectedDate.Value);

            dtpFromDate.SelectedDate = SearchDate[0];
            dtpFromDate.SelectedDate = SearchDate[1];
        }

        // 금월 버튼 클릭 이벤트
        private void btnThisMonth_Click(object sender, RoutedEventArgs e)
        {
            dtpFromDate.SelectedDate = lib.BringThisMonthDatetimeList()[0];
            dtpFromDate.SelectedDate = lib.BringThisMonthDatetimeList()[1];
        }
        #endregion



        #region 그룹박스 클리어 / 저장작업 전, 그룹박스 데이터 체크.

        // 그룹박스 항목 클리어.
        private void GBStockBoxClear()
        {
            dtpStandardDate.Text = string.Empty;
            txtGbCustomer.Text = string.Empty;
            txtGbArticleID.Text = string.Empty;
            txtGbArticle.Text = string.Empty;
            cboWareHouse.SelectedIndex = -1;
            txtStockQty.Text = string.Empty;
            cboUnitClss.SelectedIndex = -1;
        }


        private bool GBStockBoxDataCheck()
        {
            // 1. 기준일.
            if (lib.IsNullOrWhiteSpace(dtpStandardDate.Text) == true)
            {
                MessageBox.Show("기준일자를 기입해야 합니다.");
                return false;
            }
            // 2. 거래처.
            if (chkGbCustomer.IsChecked == true)
            {
                if (lib.IsNullOrWhiteSpace(txtGbCustomer.Text) == true)
                {
                    MessageBox.Show("거래처를 기입하거나, 거래처 선택을 해제하세요.");
                    return false;
                }
            }
            // 3. 품명
            if (lib.IsNullOrWhiteSpace(txtGbArticle.Text) == true)
            {
                MessageBox.Show("품명을 기입해야 합니다.");
                return false;
            }
            // 4. 창고
            if (cboWareHouse.SelectedValue == null)
            {
                MessageBox.Show("창고를 선택해야 합니다.");
                return false;
            }
            // 5. 재고량
            if (lib.IsIntOrAnother(txtStockQty.Text) == false)
            {
                MessageBox.Show("재고량은 숫자로만 기입해야 합니다.");
                return false;
            }
            // 6. 단위.
            if (cboUnitClss.SelectedValue == null)
            {
                MessageBox.Show("단위를 선택해야 합니다.");
                return false;
            }

            return true;
        }

        #endregion


        #region CRUD / SQL 파라미터 모음

        // 저장.
        private void SaveData(string TagNUM)
        {

            try
            {

                if (TagNUM == "1")      // 신규추가입니다.
                {
                    // 신규추가 저장 insert.
                    string CustomID = string.Empty;
                    if (chkGbCustomer.IsChecked == true) { CustomID = txtGbCustomer.Tag.ToString(); }


                    // 추가저장시 > findTargetValue 설정.
                    lstCompareValue.Add(dtpStandardDate.Text.Substring(5, 2) + "/" + dtpStandardDate.Text.Substring(8, 2));
                    lstCompareValue.Add(txtGbArticle.Tag.ToString());

                    // 1단계. 신규저장 하고자 하는 내용이 기 등록건은 아닌지 체크해야 한다. (초재고이기 때문에, 등록에 엄격해야 한다.)
                    Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                    sqlParameter.Clear();
                    sqlParameter.Add("sDate", dtpStandardDate.Text.Substring(0, 10).Replace("-", ""));
                    sqlParameter.Add("CustomID", CustomID);
                    sqlParameter.Add("ArticleID", txtGbArticle.Tag.ToString());
                    sqlParameter.Add("sLocID", cboWareHouse.SelectedValue.ToString());     //창고ID.

                    DataSet ds = DataStore.Instance.ProcedureToDataSet("xp_Subul_sStockDataOne", sqlParameter, false);

                    if (ds != null && ds.Tables.Count > 0)
                    {
                        DataTable dt = null;
                        dt = ds.Tables[0];

                        if (dt.Rows.Count > 0)
                        {
                            MessageBox.Show("해당 기준일, 거래처, 품명의 데이터가 이미 존재합니다." + "\r\n"
                                 + "확인후 다시 작업해 주시기 바랍니다.", "기 등록 존재");
                            return;
                        }
                    }


                    //2단계. 인서트.
                    sqlParameter = new Dictionary<string, object>();
                    sqlParameter.Clear();
                    sqlParameter.Add("sDate", dtpStandardDate.Text.Substring(0, 10).Replace("-", ""));
                    sqlParameter.Add("CustomID", CustomID);
                    sqlParameter.Add("ArticleID", txtGbArticle.Tag.ToString());
                    sqlParameter.Add("StockQty", Convert.ToInt32(txtStockQty.Text));
                    sqlParameter.Add("StockUnitClss", cboUnitClss.SelectedValue.ToString());
                    sqlParameter.Add("LocID", cboWareHouse.SelectedValue.ToString());     //창고ID.
                    sqlParameter.Add("CreateUserID", MainWindow.CurrentUser);

                    string[] result = DataStore.Instance.ExecuteProcedure("xp_Subul_iStock", sqlParameter, false);
                    if (!result[0].Equals("success"))
                    {
                        MessageBox.Show("이상발생, 관리자에게 문의하세요.");
                        return;
                    }
                }


                else if (TagNUM == "2")         // 수정 저장입니다.
                {
                    // 수정 저장 update.
                    string CustomID = string.Empty;
                    if (txtGbCustomer.Text != null) { CustomID = txtGbCustomer.Tag.ToString(); }

                    // 수정저장시 > findTargetValue 설정.
                    lstCompareValue.Add(dtpStandardDate.Text.Substring(5, 2) + "/" + dtpStandardDate.Text.Substring(8, 2));
                    lstCompareValue.Add(txtGbArticle.Tag.ToString());

                    Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                    sqlParameter.Clear();
                    sqlParameter.Add("sDate", dtpStandardDate.Text.Substring(0, 10).Replace("-", ""));
                    sqlParameter.Add("CustomID", CustomID);
                    sqlParameter.Add("ArticleID", txtGbArticle.Tag.ToString());
                    sqlParameter.Add("StockQty", Convert.ToInt32(txtStockQty.Text));
                    sqlParameter.Add("StockUnitClss", cboUnitClss.SelectedValue.ToString());
                    sqlParameter.Add("LocID", cboWareHouse.SelectedValue.ToString());     //창고ID.
                    sqlParameter.Add("LastUpdateUserID", MainWindow.CurrentUser);

                    string[] result = DataStore.Instance.ExecuteProcedure("xp_Subul_uStock", sqlParameter, false);
                    if (!result[0].Equals("success"))
                    {
                        MessageBox.Show("이상발생, 관리자에게 문의하세요.");
                        return;
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


        // 삭제용.
        private void DeleteData()
        {
            try
            {

                string CustomID = string.Empty;
                if (txtGbCustomer.Text != null) { CustomID = txtGbCustomer.Tag.ToString(); }

                Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                sqlParameter.Clear();
                sqlParameter.Add("sDate", dtpStandardDate.Text.Substring(0, 10).Replace("-", ""));
                sqlParameter.Add("CustomID", CustomID);
                sqlParameter.Add("ArticleID", txtGbArticle.Tag.ToString());
                sqlParameter.Add("LocID", cboWareHouse.SelectedValue.ToString());     //창고ID.

                string[] result = DataStore.Instance.ExecuteProcedure("xp_Subul_dStock", sqlParameter, false);
                if (!result[0].Equals("success"))
                {
                    MessageBox.Show("이상발생, 관리자에게 문의하세요.");
                    return;
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


        // 닫기 버튼 (나가기)
        private void btnClose_Click(object sender, RoutedEventArgs e)
        {
            int i = 0;
            foreach (MenuViewModel mvm in MainWindow.mMenulist)
            {
                if (mvm.subProgramID.ToString().Contains("MDI"))
                {
                    if (this.ToString().Equals((mvm.subProgramID as MdiChild).Content.ToString()))
                    {
                        (MainWindow.mMenulist[i].subProgramID as MdiChild).Close();
                        break;
                    }
                }
                i++;
            }
        }


        // 천단위 콤마, 소수점 버리기
        private string stringFormatN0(object obj)
        {
            return string.Format("{0:N0}", obj);
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



        #region 엑셀
        //엑셀.
        private void btnExcel_Click(object sender, RoutedEventArgs e)
        {
            DataTable dt = null;
            string Name = string.Empty;
            Lib lib2 = new Lib();

            string[] lst = new string[2];
            lst[0] = "메인 그리드";
            lst[1] = dgdStock.Name;

            ExportExcelxaml ExpExc = new ExportExcelxaml(lst);

            ExpExc.ShowDialog();

            if (ExpExc.DialogResult.HasValue)
            {
                if (ExpExc.Check.Equals("Y"))
                    dt = lib2.DataGridToDTinHidden(dgdStock);
                else
                    dt = lib2.DataGirdToDataTable(dgdStock);

                Name = dgdStock.Name;

                if (lib2.GenerateExcel(dt, Name))
                {
                    lib2.excel.Visible = true;
                    lib2.ReleaseExcelObject(lib2.excel);
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

            lib2 = null;

            //if (dgdStock.Items.Count < 1)
            //{
            //    MessageBox.Show("먼저 검색해 주세요.");
            //    return;
            //}

            //Lib lib = new Lib();
            //DataTable dt = null;
            //string Name = string.Empty;

            //string[] lst = new string[2];
            //lst[0] = "메인 그리드";
            //lst[1] = dgdStock.Name;

            //ExportExcelxaml ExpExc = new ExportExcelxaml(lst);

            //ExpExc.ShowDialog();

            //if (ExpExc.DialogResult.HasValue)
            //{
            //    if (ExpExc.choice.Equals(dgdStock.Name))
            //    {
            //        //MessageBox.Show("대분류");
            //        if (ExpExc.Check.Equals("Y"))
            //            dt = lib.DataGridToDTinHidden(dgdStock);
            //        else
            //            dt = lib.DataGirdToDataTable(dgdStock);

            //        Name = dgdStock.Name;
            //        lib.GenerateExcel(dt, Name);
            //        lib.excel.Visible = true;
            //    }
            //    else
            //    {
            //        if (dt != null)
            //        {
            //            dt.Clear();
            //        }
            //    }
            //}
        }

        #endregion


        #region 기타 메서드 모음

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


        #region 텍스트박스 엔터 키 이동

        private void dtpStandardDate_PreviewKeyUp(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                dtpStandardDate.IsDropDownOpen = true;
            }
        }
        private void dtpStandardDate_CalendarClosed(object sender, RoutedEventArgs e)
        {
            lib.SendK(Key.Tab, this);
        }
        private void chkGbCustomer_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                txtGbArticle.Focus();
            }
        }
        private void txtGbArticle_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                btnGbArticle_Click(null, null);
                cboWareHouse.Focus();
                cboWareHouse.IsDropDownOpen = true;
            }
        }
        private void txtGbCustomer_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                btnGbCustomer_Click(null, null);
            }
        }
        private void cboWareHouse_DropDownClosed(object sender, EventArgs e)
        {
            lib.SendK(Key.Tab, this);
        }
        private void txtStockQty_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                lib.SendK(Key.Tab, this);
                cboUnitClss.IsDropDownOpen = true;
            }
        }
        private void cboUnitClss_DropDownClosed(object sender, EventArgs e)
        {
            dtpStandardDate.Focus();
        }

        // 엔터 키를 통한 탭 인덱스 키 이동.
        private void EnterMove_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                lib.SendK(Key.Tab, this);
            }
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
    }


    #endregion




    class Win_ord_ControlStock_U_View : BaseView
    {
        public override string ToString()
        {
            return (this.ReportAllProperties());
        }

        public ObservableCollection<CodeView> cboTrade { get; set; }

        // 조회 값.    
        public string BasisDate { get; set; }
        public string CustomID { get; set; }
        public string ArticleID { get; set; }
        public string ProcInClss { get; set; }
        public string StockClss { get; set; }

        public string StockQty { get; set; }
        public string StockUnitClss { get; set; }
        public string WeightPerYard { get; set; }
        public string LocID { get; set; }
        public string KCustom { get; set; }

        public string BuyerArticleNo { get; set; }
        public string Article { get; set; }
        public string LocName { get; set; }
        public string UnitClssName { get; set; }
        public string Sabun { get; set; }

        public string Full_BasisDate { get; set; }
    }

    class Win_ord_ControlStock_U_ViewSum
    {
        public string TotalCnt { get; set; }
        public string TotalQty { get; set; }
    }
}

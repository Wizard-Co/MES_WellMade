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
using WPF.MDI;

namespace WizMes_WellMade
{
    /// <summary>
    /// Win_prd_PlanInput_U.xaml에 대한 상호 작용 논리
    /// </summary>
    public partial class Win_prd_PlanInput_U : UserControl
    {
        string stDate = string.Empty;
        string stTime = string.Empty;

        int rowNum = 0;

        bool saveComplete = false;
        string InstID = ""; // 작업지시 PK

        System.Windows.Threading.DispatcherTimer dispatcherTimer = new System.Windows.Threading.DispatcherTimer();

        Lib lib = new Lib();

        public Win_prd_PlanInput_U()
        {
            InitializeComponent();
        }

        private void UserControl_Loaded(object sender, RoutedEventArgs e)
        {
            stDate = DateTime.Now.ToString("yyyyMMdd");
            stTime = DateTime.Now.ToString("HHmm");

            DataStore.Instance.InsertLogByFormS(this.GetType().Name, stDate, stTime, "S");

            Lib.Instance.UiLoading(sender);

            //btnToday_Click(null, null);
            chkOrderDay.IsChecked = true;
            dtpSDate.SelectedDate = Lib.Instance.BringThisMonthDatetimeList()[0];
            dtpEDate.SelectedDate = Lib.Instance.BringThisMonthDatetimeList()[1];

            dtpInstDate.SelectedDate = DateTime.Today; // 지시일자
            dtpInstCompleteDate.SelectedDate = DateTime.Today; // 작업완료일
        }

        #region 추가, 수정모드 / 저장완료, 취소 시

        private void SaveUpdateMode()
        {
            dgdPlanInput.Visibility = Visibility.Visible; // 공정 상세 내역 보이고
            dgdMain.IsHitTestVisible = false; // 메인 그리드 노터치 플리즈

            // 버튼 세팅
            btnAdd.IsEnabled = false;
            btnCancel.IsEnabled = true;
            btnSave.IsEnabled = true;
        }

        // 저장완료, 취소 시
        private void CompleteCancelMode()
        {
            // 하위 그리드를 비워줌
            if (dgdPlanInput.Items.Count > 0)
            {
                dgdPlanInput.Items.Clear();
            }

            dgdPlanInput.Visibility = Visibility.Hidden; // 하위그리드 숨김
            dgdMain.IsHitTestVisible = true;

            // 버튼 정리
            btnAdd.IsEnabled = true;
            btnCancel.IsEnabled = false;
            btnSave.IsEnabled = false;
        }

        #endregion // 추가, 수정모드 / 저장완료, 취소 시

        #region 공정패턴 콤보박스
        private void setCboPattern(string ArticleGrpID)
        {
            ObservableCollection<CodeView> cboPattern = new ObservableCollection<CodeView>();

            List<string> CbView = new List<string>();
            List<string> PatternID = new List<string>();

            string strCompare1 = string.Empty;
            string strCompare2 = string.Empty;
            string TheView = string.Empty;

            try
            {
                Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                sqlParameter.Add("sArticleGrpID", ArticleGrpID);

                DataSet ds = DataStore.Instance.ProcedureToDataSet("xp_PlanInput_sPatternByArticleGrpID", sqlParameter, false);

                if (ds != null && ds.Tables.Count > 0)
                {
                    DataTable dt = ds.Tables[0];

                    if (dt.Rows.Count > 0)
                    { 
                        DataRowCollection drc = dt.Rows;
                        int i = 0;

                        foreach (DataRow item in drc)
                        {
                            CodeView codeView = new CodeView();

                            strCompare1 = item["PatternID"].ToString().Trim();
                            strCompare2 = item["Pattern"].ToString().Trim();

                            TheView = strCompare1 + "." + strCompare2 + " : ";

                            foreach (DataRow items in drc)
                            {
                                if (items["PatternID"].ToString().Equals(strCompare1))
                                {
                                    TheView += " [" + items["Process"].ToString() + "] →";
                                }
                            }
                            if (TheView != null && !TheView.Equals(""))
                            {
                                TheView = TheView.Substring(0, TheView.Length - 1);
                            }

                            if (CbView.Count > 0)
                            {
                                if (!CbView[i].Substring(0, 2).Equals(strCompare1))
                                {
                                    codeView.code_id = strCompare1;
                                    codeView.code_name = TheView;

                                    CbView.Add(TheView);
                                    cboPattern.Add(codeView);
                                    i++;
                                }
                            }
                            else
                            {
                                codeView.code_id = strCompare1;
                                codeView.code_name = TheView;

                                CbView.Add(TheView);
                                cboPattern.Add(codeView);
                            }
                        }
                        drc.Clear();
                    }
                    dt.Clear();
                }
                ds.Clear();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
            finally
            {
                DataStore.Instance.CloseConnection();
            }

            cboProcessPattern.ItemsSource = cboPattern;
            cboProcessPattern.DisplayMemberPath = "code_name";
            cboProcessPattern.SelectedValuePath = "code_id";
        }
        #endregion

        #region Header 검색조건 - 일자
        //수주 일자
        private void lblOrderDay_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            if (chkOrderDay.IsChecked == true)
            {
                chkOrderDay.IsChecked = false;
            }
            else
            {
                chkOrderDay.IsChecked = true;
            }
        }
        //수주 일자
        private void chkOrderDay_Checked(object sender, RoutedEventArgs e)
        {
            dtpSDate.IsEnabled = true;
            dtpEDate.IsEnabled = true;
        }
        //수주 일자
        private void chkOrderDay_Unchecked(object sender, RoutedEventArgs e)
        {
            dtpSDate.IsEnabled = false;
            dtpEDate.IsEnabled = false;
        }

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
        #endregion

        #region Header 검색조건 - 기타

        
        //거래처
        private void lblCustom_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            if (chkCustom.IsChecked == true)
            {
                chkCustom.IsChecked = false;
            }
            else
            {
                chkCustom.IsChecked = true;
            }
        }

        //거래처
        private void chkCustom_Checked(object sender, RoutedEventArgs e)
        {
            txtCustom.IsEnabled = true;
            btnPfCustom.IsEnabled = true;
            txtCustom.Focus();
        }

        //거래처
        private void chkCustom_Unchecked(object sender, RoutedEventArgs e)
        {
            txtCustom.IsEnabled = false;
            btnPfCustom.IsEnabled = false;
        }

        //거래처
        private void txtCustom_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                MainWindow.pf.ReturnCode(txtCustom, (int)Defind_CodeFind.DCF_CUSTOM, "");
            }
        }

        //거래처
        private void btnPfCustom_Click(object sender, RoutedEventArgs e)
        {
            MainWindow.pf.ReturnCode(txtCustom, (int)Defind_CodeFind.DCF_CUSTOM, "");
        }

        //품번
        private void LabelBuyerArticleNoSearch_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            if (CheckBoxBuyerArticleNoSearch.IsChecked == true)
            {
                CheckBoxBuyerArticleNoSearch.IsChecked = false;
            }
            else
            {
                CheckBoxBuyerArticleNoSearch.IsChecked = true;
            }
        }

        
        //품번
        private void CheckBoxBuyerArticleNoSearch_Checked(object sender, RoutedEventArgs e)
        {
            TextBoxBuyerArticleNoSearch.IsEnabled = true;
            ButtonBuyerArticleNoSearch.IsEnabled = true;
            TextBoxBuyerArticleNoSearch.Focus();
        }

        //품번
        private void CheckBoxBuyerArticleNoSearch_Unchecked(object sender, RoutedEventArgs e)
        {
            TextBoxBuyerArticleNoSearch.IsEnabled = false;
            ButtonBuyerArticleNoSearch.IsEnabled = false;
        }

        //품번
        private void TextBoxBuyerArticleNoSearch_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                MainWindow.pf.ReturnCode(TextBoxBuyerArticleNoSearch, 76, TextBoxBuyerArticleNoSearch.Text);
            }
        }

        //품번
        private void ButtonBuyerArticleNoSearch_Click(object sender, RoutedEventArgs e)
        {
            MainWindow.pf.ReturnCode(TextBoxBuyerArticleNoSearch, 76, TextBoxBuyerArticleNoSearch.Text);
        }

        //품명
        private void lblArticle_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            if (chkArticle.IsChecked == true) { chkArticle.IsChecked = false; }
            else { chkArticle.IsChecked = true; }
        }

        //품명
        private void chkArticle_Checked(object sender, RoutedEventArgs e)
        {
            txtArticle.IsEnabled = true;
            btnPfArticle.IsEnabled = true;
            txtArticle.Focus();
        }

        //품명
        private void chkArticle_Unchecked(object sender, RoutedEventArgs e)
        {
            txtArticle.IsEnabled = false;
            btnPfArticle.IsEnabled = false;
        }

        //품명
        private void txtArticle_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                MainWindow.pf.ReturnCode(txtArticle, 77, "");
            }
        }

        //품명
        private void btnPfArticle_Click(object sender, RoutedEventArgs e)
        {
            MainWindow.pf.ReturnCode(txtArticle, 77, "");
        }

        //관리번호
        private void lblOrderID_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            if (chkOrderID.IsChecked == true) { chkOrderID.IsChecked = false; }
            else { chkOrderID.IsChecked = true; }
        }

        //관리번호
        private void chkOrderID_Checked(object sender, RoutedEventArgs e)
        {
            txtOrderID.IsEnabled = true;
            txtOrderID.Focus();
        }

        //관리번호
        private void chkOrderID_Unchecked(object sender, RoutedEventArgs e)
        {
            txtOrderID.IsEnabled = false;
        }

        //OrderNo
        private void txtOrderID_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                if (e.Key == Key.Enter)
                {
                    MainWindow.pf.ReturnCode(txtOrderID, (int)Defind_CodeFind.DCF_ORDER, "");
                }
            }
        }

        //해상도가 낮아지면 체크박스 클릭이 어려워지므로 라벨 클릭으로 대체할수 있게 한다.
        private void lblCompleted_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            if (chkCompleted.IsChecked == true) { chkCompleted.IsChecked = false; }
            else { chkCompleted.IsChecked = true; }
        }

        //해상도가 낮아지면 체크박스 클릭이 어려워지므로 라벨 클릭으로 대체할수 있게 한다.
        private void lblOrderClosed_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            if (chkOrderClosed.IsChecked == true) { chkOrderClosed.IsChecked = false; }
            else { chkOrderClosed.IsChecked = true; }
        }

        //OrderNo
        private void rbnOrderNo_Click(object sender, RoutedEventArgs e)
        {
            Check_bdrOrder();
        }

        //관리번호
        private void rbnOrderID_Click(object sender, RoutedEventArgs e)
        {
            Check_bdrOrder();
        }

        private void Check_bdrOrder()
        {
            if (rbnOrderID.IsChecked == true)
            {
                tbkOrderID.Text = "관리번호";
            }
            else if (rbnOrderNo.IsChecked == true)
            {
                tbkOrderID.Text = "Order No.";
            }
        }

        #endregion

        #region Header 우측 상단 버튼 - 검색, 닫기

        //검색
        private void btnSearch_Click(object sender, RoutedEventArgs e)
        {
            // 추가 모드일때는 취소 후 검색 되도록
            if (btnCancel.IsEnabled == true)
            {
                if (MessageBox.Show("작업중인 항목을 취소하시겠습니까?", "취소 확인", MessageBoxButton.YesNo) != MessageBoxResult.Yes)
                {
                    return;
                }

                CompleteCancelMode();
            }

            //검색버튼 비활성화
            btnSearch.IsEnabled = false;

            Dispatcher.BeginInvoke(new Action(() =>

            {
                //로직
                using (Loading lw = new Loading(re_Search))
                {
                    lw.ShowDialog();
                }

                if(dgdMain.Items.Count == 0)
                {
                    InsertDisable();
                    return;
                }
                else
                {
                    dgdMain.SelectedIndex = rowNum;
                }

            }), System.Windows.Threading.DispatcherPriority.Background);



            Dispatcher.BeginInvoke(new Action(() =>

            {
                btnSearch.IsEnabled = true;

            }), System.Windows.Threading.DispatcherPriority.Background);
            
        }

        //닫기
        private void btnClose_Click(object sender, RoutedEventArgs e)
        {
            DataStore.Instance.InsertLogByFormS(this.GetType().Name, stDate, stTime, "E");
            Lib.Instance.ChildMenuClose(this.ToString());
        }

        #endregion

        // 재조회
        private void re_Search()
        {
            // 서브 그리드는 무조건 초기화
            if (dgdPlanInput.Items.Count > 0)
            {
                dgdPlanInput.Items.Clear();
            }

            FillGrid();

            if (dgdMain.Items.Count > 0)
            {
                dgdMain.SelectedItem = rowNum;
            }
            // 데이터가 없다면, 메시지를 여기서 출력
            else if (dgdMain.Items.Count == 0)
            {
                MessageBox.Show("조회된 데이터가 없습니다.");
                return;
            }
            else
            {
                dgdMain.SelectedIndex = rowNum;
            }
        }

        #region 조회 FillGrid

        private void FillGrid()
        {
            if (dgdMain.Items.Count > 0)
            {
                dgdMain.Items.Clear();
            }

            try
            {
                Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                sqlParameter.Clear();
                sqlParameter.Add("ChkDate", chkOrderDay.IsChecked == true ? 1 : 0);
                sqlParameter.Add("SDate", chkOrderDay.IsChecked == true && dtpSDate.SelectedDate != null ? dtpSDate.SelectedDate.Value.ToString("yyyyMMdd") : "");
                sqlParameter.Add("EDate", chkOrderDay.IsChecked == true && dtpEDate.SelectedDate != null ? dtpEDate.SelectedDate.Value.ToString("yyyyMMdd") : "");
                sqlParameter.Add("ChkCustomID", chkCustom.IsChecked == true ? 1 : 0);
                sqlParameter.Add("CustomID", chkCustom.IsChecked == true ?
                    (txtCustom.Tag != null ? txtCustom.Tag.ToString() : "") : "");

                sqlParameter.Add("ChkArticleID", chkArticle.IsChecked == true ? 1 : 0);
                sqlParameter.Add("ArticleID", chkArticle.IsChecked == true ?
                    (txtArticle.Tag != null ? txtArticle.Tag.ToString() : "") : "");
                sqlParameter.Add("ChkOrder", chkOrderID.IsChecked == true ? (rbnOrderID.IsChecked == true ? 1 : 2) : 0);
                sqlParameter.Add("Order", chkOrderID.IsChecked == true ? txtOrderID.Text : "");
                sqlParameter.Add("ChkCompleted", chkCompleted.IsChecked == true ? 1 : 0);
                sqlParameter.Add("ChkOrderClosed", chkOrderClosed.IsChecked == true ? 1 : 0);
                sqlParameter.Add("ChkBuyerArticleNo", CheckBoxBuyerArticleNoSearch.IsChecked == true ? 1 : 0);
                sqlParameter.Add("BuyerArticleNoID", CheckBoxBuyerArticleNoSearch.IsChecked == true ? (TextBoxBuyerArticleNoSearch.Tag != null ? TextBoxBuyerArticleNoSearch.Tag.ToString() : "") : "");

                DataSet ds = DataStore.Instance.ProcedureToDataSet_LogWrite("xp_PlanInput_sOrder", sqlParameter, true, "R");

                if (ds != null && ds.Tables.Count > 0)
                {
                    DataTable dt = ds.Tables[0];
                    int i = 0;

                    if (dt.Rows.Count > 0)
                    { 
                        DataRowCollection drc = dt.Rows;

                        foreach (DataRow dr in drc)
                        {
                            i++;

                            var WinPlanOrder = new Win_prd_PlanInput_U_CodeView()
                            {
                                Num = i + 1,
                                KCustom = dr["KCustom"].ToString(),
                                Article = dr["Article"].ToString(),
                                OrderID = dr["OrderID"].ToString(),
                                OrderNo = dr["OrderNo"].ToString(),
                                OrderQty = Convert.ToDouble(dr["OrderQty"]),
                                notOrderInstQty = Convert.ToDouble(dr["notOrderInstQty"]),
                                OrderInstQy = Convert.ToDouble(dr["OrderInstQy"]),
                                p1WorkQty = Convert.ToDouble(dr["p1WorkQty"]),
                                p1ProcessID = dr["p1ProcessID"].ToString(),
                                p1ProcessName = dr["p1ProcessName"].ToString(),
                                InspectQty = Convert.ToDouble(dr["InspectQty"]),
                                OutQty = Convert.ToDouble(dr["OutQty"]),
                                PatternID = dr["PatternID"].ToString(),
                                ArticleGrpID = dr["ArticleGrpID"].ToString(),
                                BuyerModel = dr["BuyerModel"].ToString(),
                                BuyerModelID = dr["BuyerModelID"].ToString(),
                                BuyerArticleNo = dr["BuyerArticleNo"].ToString(),
                                Remark = dr["Remark"].ToString(),
                                ArticleID = dr["ArticleID"].ToString(),
                                OrderClosed = dr["OrderClosed"].ToString(),
                                Completed = dr["Completed"].ToString(),
                                ArticleGrpName = dr["ArticleGrpName"].ToString(),
                                //InstID = dr["InstID"].ToString(),
                                AcptDate = dr["AcptDate"].ToString(),
                                AcptDate_CV = DatePickerFormat(dr["AcptDate"].ToString()),
                                OrderSeq = dr["OrderSeq"].ToString(),
                            };

                            dgdMain.Items.Add(WinPlanOrder);
                        }
                        tbkCount.Text = "▶ 검색결과 : " + i.ToString() + " 건";
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

        #region 하단 그리드 조회 FillGridPlanInput

        private void FillGridPlanInput(string strPatternID, string strArticleID, string strOrderID)
        {
            List<Win_prd_PlanArticleOne_CodeView> lstPl = new List<Win_prd_PlanArticleOne_CodeView>();

            if (dgdPlanInput.Items.Count > 0)
            {
                dgdPlanInput.Items.Clear();
            }

            try
            {
                Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                sqlParameter.Clear();

                sqlParameter.Add("sPatternID", strPatternID);
                sqlParameter.Add("sArticleID", strArticleID);
                sqlParameter.Add("sOutMessage", "");

                DataSet ds = DataStore.Instance.ProcedureToDataSet("xp_PlanInput_sPatternArticleOne", sqlParameter, false);

                if (ds != null && ds.Tables.Count > 0)
                {
                    DataTable dt = ds.Tables[0];
                    int i = 0;

                    if (dt.Rows.Count > 0)
                    {
                        // 문제가 있을 경우에는, OutMessage 하나만 가져오기 때문에, 
                        // 컬럼이 1개인 경우에 해당 메시지를 출력
                        if (dt.Columns.Count == 1)
                        {
                            MessageBox.Show(dt.Rows[0].ItemArray[0].ToString());
                        }
                        else
                        {
                            DataRowCollection drc = dt.Rows;

                            foreach (DataRow dr in drc)
                            {
                                var WinPlanArticle = new Win_prd_PlanArticleOne_CodeView()
                                {
                                     Num = i + 1,
                                    PatternSeq = dr["PatternSeq"].ToString(),
                                    ProcessID = dr["ProcessID"].ToString(),
                                    Process = dr["Process"].ToString(),
                                    Qty = dr["Qty"].ToString(),
                                    Article = dr["Article"].ToString(),
                                    ArticleID = dr["ArticleID"].ToString(),
                                    BuyerArticleNo = dr["BuyerArticleNo"].ToString(),
                                    LVL = dr["LVL"].ToString(),
                                    StartDate = dtpInstDate.SelectedDate.Value.ToString("yyyyMMdd"),
                                    EndDate = dtpInstCompleteDate.SelectedDate.Value.ToString("yyyyMMdd"),
                                    ChildBuyerArticleNo = dr["ChildBuyerArticleNo"].ToString(),
                                };

                                //지시수량
                                WinPlanArticle.InstQty = stringFormatN2(ConvertDouble(WinPlanArticle.Qty) * ConvertDouble(txtQty.Text));

                                // 날짜 세팅
                                WinPlanArticle.StartDate_CV = DatePickerFormat(WinPlanArticle.StartDate);
                                WinPlanArticle.EndDate_CV = DatePickerFormat(WinPlanArticle.EndDate);

                                if (WinPlanArticle.ArticleID == null
                                    || WinPlanArticle.ArticleID.Trim().Equals(""))
                                {
                                    MessageBox.Show("해당 품명의 생산공정이 공정패턴과 일치하지 않습니다.\r해당 품명정보를 확인해주세요.");
                                    return;
                                }

                                lstPl.Add(WinPlanArticle);

                                //dgdPlanInput.Items.Add(WinPlanArticle);
                                i++;
                            }

                            tbkCount2.Text = "▶ 검색결과 : " + i.ToString() + " 건";

                            if (lstPl.Count > 0)
                            {
                                for (int k = 0; k < lstPl.Count; k++)
                                {
                                    dgdPlanInput.Items.Add(lstPl[k]);
                                }
                            }
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

        #region Content 메인그리드 선택 시 - dgdMain_SelectionChanged 

        private void dgdMain_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            chkMtrExceptYN.IsChecked = false;
            chkOutwareExceptYN.IsChecked = false;
            chkRemainData.IsChecked = true;
            chkAutoPrint.IsChecked = false;

            var Plan = dgdMain.SelectedItem as Win_prd_PlanInput_U_CodeView;

            if (Plan != null
                && Plan.ArticleGrpID != null)
            {
                /********************
                 * DataContext 로 바인딩 해서 넣게 되면, 값이 바뀐 상태로 유지 되기 때문에, 바인딩을 제외하고 수동으로 값 입력
                 ********************/
                
                // 공정패턴
                setCboPattern(Plan.ArticleGrpID); // 공정패턴 콤보박스 세팅!!
                cboProcessPattern.SelectedValue = Plan.PatternID;

                // 지시일자
                dtpInstDate.SelectedDate = DateTime.Today;
                // 작업 완료일은 이달 말로 설정
                var lastDay = DateTime.Today.AddMonths(1).AddDays(-DateTime.Today.Day);
                dtpInstCompleteDate.SelectedDate = lastDay;

                // 지시수량
                txtQty.Text = Plan.notOrderInstQty.ToString();

                // 비고
                txtRemark.Text = Plan.Remark;

                InsertEnable();
            }
        }
        #endregion

        #region Content 버튼 모음 - 추가, 취소, 작업지시

        //추가
        private void btnAdd_Click(object sender, RoutedEventArgs e)
        {
            var WinPlan = dgdMain.SelectedItem as Win_prd_PlanInput_U_CodeView;

            if (WinPlan != null)
            {
                rowNum = dgdMain.SelectedIndex;

                if (AddCheckData(WinPlan))
                {
                    FillGridPlanInput(cboProcessPattern.SelectedValue.ToString(), WinPlan.ArticleID, WinPlan.OrderID);

                    if (dgdPlanInput.Items.Count > 0)
                    {
                        SaveUpdateMode();

                        InsertDisable();
                    }
                }
            }
        }

        #region 추가 전 체크 AddCheckData
        private bool AddCheckData(Win_prd_PlanInput_U_CodeView WinPlan) 
        {
            bool flag = true;

            // 공정패턴이 없다면
            if (cboProcessPattern.SelectedValue == null)
            {
                MessageBox.Show("공정패턴을 선택해주세요.");
                return false;
            }

            // 품명이 없다면
            if (WinPlan.ArticleID == null)
            {
                MessageBox.Show("해당 수주에 품명 정보가 없습니다.");
                return false;
            }

            // 지시일자
            if (dtpInstDate.SelectedDate == null)
            {
                MessageBox.Show("지시일자를 선택해주세요.");
                return false;
            }

            // 작업완료일
            if (dtpInstCompleteDate.SelectedDate == null)
            {
                MessageBox.Show("작업완료일을 선택해주세요.");
                return false;
            }

            return flag;
        }
        #endregion

        //취소
        private void btnCancel_Click(object sender, RoutedEventArgs e)
        {
            if (btnAdd.IsEnabled == false)
            {
                if (MessageBox.Show("선택하신 항목을 취소하시겠습니까?", "취소 확인", MessageBoxButton.YesNo) == MessageBoxResult.Yes)
                {
                    CompleteCancelMode();
                    InsertDisable();

                    using (Loading lw = new Loading(re_Search))
                    {
                        lw.ShowDialog();
                    }
                }
            }
        }  

        //작업지시
        private void btnSave_Click(object sender, RoutedEventArgs e)
        {
            saveComplete = false;

            using (Loading lw = new Loading(beSave))
            {
                lw.ShowDialog();
            }

            InsertDisable();

            // 저장과 동시에 작업지시 안내 체크를 한 경우에는..?
            // 1. 생산계획 관리 화면을 띄워야 하고,
            // 2. 지시 일자를 가지고, 해당 지시일자로 검색을 해야 하고 → 날짜로만 하자
            // 3. 해당 작지에 체크만 하자.
            if (saveComplete == true)
            {
                //dispatcherTimer.Tick += new EventHandler(PopUpComponent);
                //dispatcherTimer.Interval = new TimeSpan(0, 0, 1);
                //dispatcherTimer.Start();
                MainWindow.plInputFlag_SavePrint = true;
                PopUpComponent(null, null);
            }
        }

        // 저장하는 메서드 묶음
        private void beSave()
        {
            if (SaveData())
            {
                if (chkAutoPrint.IsChecked == true)
                {
                    saveComplete = true;

                    MainWindow.plInput.Clear();
                    MainWindow.plInput.Add("Date", dtpInstDate.SelectedDate);
                    MainWindow.plInput.Add("InstID", InstID);
                }

                re_Search();

                CompleteCancelMode();
            }
        }

        // 화면을 띄우는 메서드
        private void PopUpComponent(object sender, System.EventArgs e)
        {
            int i = 0;
            foreach (MenuViewModel mvm in MainWindow.mMenulist)
            {
                if (mvm.Menu.Equals("생산계획 관리"))
                {
                    break;
                }
                i++;
            }
            try
            {
                if (MainWindow.MainMdiContainer.Children.Contains(MainWindow.mMenulist[i].subProgramID as MdiChild))
                {
                    double actualWidth = (MainWindow.mMenulist[i].subProgramID as MdiChild).ActualWidth;
                    (MainWindow.mMenulist[i].subProgramID as MdiChild).Width = actualWidth + 0.1;
                    (MainWindow.mMenulist[i].subProgramID as MdiChild).Focus();
                }
                else
                {
                    Type type = Type.GetType("WizMes_WellMade." + MainWindow.mMenulist[i].ProgramID.Trim(), true);
                    object uie = Activator.CreateInstance(type);

                    MainWindow.mMenulist[i].subProgramID = new MdiChild()
                    {
                        Title = "DAEWON [" + MainWindow.mMenulist[i].MenuID.Trim() + "] " + MainWindow.mMenulist[i].Menu.Trim() +
                                " (→" + MainWindow.mMenulist[i].ProgramID + ")",
                        Height = SystemParameters.PrimaryScreenHeight * 0.8,
                        MaxHeight = SystemParameters.PrimaryScreenHeight * 0.85,
                        Width = SystemParameters.WorkArea.Width * 0.85,
                        MaxWidth = SystemParameters.WorkArea.Width,
                        Content = uie as UIElement,
                        Tag = MainWindow.mMenulist[i]
                    };

                    MainWindow.plInputFlag_SavePrint = true;
                    MainWindow.plInput.Add("FirstFlag", true);

                    Lib.Instance.AllMenuLogInsert(MainWindow.mMenulist[i].MenuID, MainWindow.mMenulist[i].Menu, MainWindow.mMenulist[i].subProgramID);
                    MainWindow.MainMdiContainer.Children.Add(MainWindow.mMenulist[i].subProgramID as MdiChild);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("해당 화면이 존재하지 않습니다.");
            }
        }

        #endregion

        #region 하단 조건 모음

        private void txtQty_PreviewTextInput(object sender, TextCompositionEventArgs e)
        {
            Lib.Instance.CheckIsNumeric((TextBox)sender, e);
        }


        private void cboProcessPattern_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (btnAdd.IsEnabled == false
                 && cboProcessPattern.SelectedValue != null)
            {
                var WinPlan = dgdMain.SelectedItem as Win_prd_PlanInput_U_CodeView;

                if (WinPlan != null)
                {
                    dgdPlanInput.Visibility = Visibility.Visible;
                    FillGridPlanInput(cboProcessPattern.SelectedValue.ToString(), WinPlan.ArticleID, WinPlan.OrderID);
                    //dgdMain.IsEnabled = false;
                    dgdMain.IsHitTestVisible = false;
                }
                else
                {
                    return;
                }
            }
        }
        #endregion

        #region 저장
        private bool SaveData()
        {
            bool flag = false;
            List<Procedure> Prolist = new List<Procedure>();
            List<Dictionary<string, object>> ListParameter = new List<Dictionary<string, object>>();
            String devideYN = string.Empty;

            var WinPlan = dgdMain.SelectedItem as Win_prd_PlanInput_U_CodeView;
            if (WinPlan != null)
            {
                try
                {
                    if (CheckData())
                    {
                        Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                        sqlParameter.Clear();

                        sqlParameter.Add("InstID", "");
                        sqlParameter.Add("InstDate", dtpInstDate.SelectedDate.Value.ToString("yyyyMMdd"));
                        sqlParameter.Add("OrderID", WinPlan.OrderID);
                        sqlParameter.Add("OrderSeq", WinPlan.OrderSeq);
                        sqlParameter.Add("InstRoll", 0);

                        sqlParameter.Add("InstQty", txtQty.Text.Replace(",", ""));
                        sqlParameter.Add("ExpectDate", dtpInstCompleteDate.SelectedDate.Value.ToString("yyyyMMdd"));
                        sqlParameter.Add("PersonID", MainWindow.CurrentPersonID);
                        sqlParameter.Add("Remark", txtRemark.Text);
                        sqlParameter.Add("MtrExceptYN", chkMtrExceptYN.IsChecked == true ? "Y" : "N");
                        sqlParameter.Add("OutwareExceptYN", chkOutwareExceptYN.IsChecked == true ? "Y" : "N");     //단위 선택
                        sqlParameter.Add("CreateUserID", MainWindow.CurrentUser);
                        sqlParameter.Add("DevideYN", "N");

                        Procedure pro1 = new Procedure();
                        pro1.Name = "xp_PlanInput_iPlanInput";
                        pro1.OutputUseYN = "Y";
                        pro1.OutputName = "InstID";
                        pro1.OutputLength = "12";

                        Prolist.Add(pro1);
                        ListParameter.Add(sqlParameter);

                        for (int i = 0; i < dgdPlanInput.Items.Count; i++)
                        {
                            var WinPlanArticleOne = dgdPlanInput.Items[i] as Win_prd_PlanArticleOne_CodeView;
                            sqlParameter = new Dictionary<string, object>();
                            sqlParameter.Add("InstID", "");
                            sqlParameter.Add("InstDate", dtpInstDate.SelectedDate.Value.ToString("yyyyMMdd"));
                            sqlParameter.Add("ProcSeq", WinPlanArticleOne.PatternSeq);
                            sqlParameter.Add("ArticleID", WinPlanArticleOne.ArticleID);
                            sqlParameter.Add("ProcessID", WinPlanArticleOne.ProcessID);

                            sqlParameter.Add("InstRemark", WinPlanArticleOne.InstRemark == null ? "" : WinPlanArticleOne.InstRemark);
                            sqlParameter.Add("InstQty", WinPlanArticleOne.InstQty.Replace(",", ""));
                            sqlParameter.Add("StartDate", WinPlanArticleOne.StartDate);
                            sqlParameter.Add("EndDate", WinPlanArticleOne.EndDate);
                            sqlParameter.Add("Remark", WinPlanArticleOne.Remark == null ? "" : WinPlanArticleOne.Remark);

                            sqlParameter.Add("MachineID", WinPlanArticleOne.MachineID == null ? "" : WinPlanArticleOne.MachineID);
                            sqlParameter.Add("MtrExceptYN", WinPlanArticleOne.MtrExceptYN == null ? "" : WinPlanArticleOne.MtrExceptYN);
                            sqlParameter.Add("FirstInFirstOutYN", WinPlanArticleOne.FirstInFirstOutYN == null ? "" : WinPlanArticleOne.FirstInFirstOutYN);
                            sqlParameter.Add("CreateUserID", MainWindow.CurrentUser);
                            sqlParameter.Add("DevideYN", "N");

                            Procedure pro2 = new Procedure();
                            pro2.Name = "xp_PlanInput_iPlanInputSub";
                            pro2.OutputUseYN = "N";
                            pro2.OutputName = "InstID";
                            pro2.OutputLength = "12";

                            Prolist.Add(pro2);
                            ListParameter.Add(sqlParameter);
                        }

                        List<KeyValue> list_Result = new List<KeyValue>();
                        list_Result = DataStore.Instance.ExecuteAllProcedureOutputGetCS_NewLog(Prolist, ListParameter,"C");
                        string sGetID = string.Empty;

                        if (list_Result[0].key.ToLower() == "success")
                        {
                            list_Result.RemoveAt(0);
                            for (int i = 0; i < list_Result.Count; i++)
                            {
                                KeyValue kv = list_Result[i];
                                if (kv.key == "InstID")
                                {
                                    InstID = kv.value;
                                    sGetID = kv.value;
                                    flag = true;
                                }
                            }

                            if (flag)
                            {
                                UpdatePattern(WinPlan.OrderID, WinPlan.ArticleID);
                            }
                        }
                        else
                        {
                            MessageBox.Show("[저장실패]\r\n" + list_Result[0].value.ToString());
                            flag = false;
                            
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

            return flag;
        }
        #endregion

        #region 유효성 검사 CheckData()

        private bool CheckData()
        {
            bool flag = true;

            // pl_inputdet 에 데이터가 무조건 들어가야 됨!!!
            // == dgdPlanInput 에 데이터가 들어가 있어야 저장이 되야됨.
            if (dgdPlanInput.Items.Count == 0)
            {
                MessageBox.Show("하단의 공정별 작업지시 데이터가 없습니다.\r(공정패턴을 변경해주시거나, 취소 후 다시 작업지시를 내려주세요.)", "저장 전 확인");
                flag = false;
                return flag;
            }

            // 검사 ArticleID 와 마지막 공정 ArticleID 가 같지 않으면 작업지시 내리지 못하도록.
            // 무조건 검사 ArticleID = 마지막 공정 ArticleID

            return flag;
        }

        #endregion

        #region UpdatePattern
        //
        private bool UpdatePattern(string strOrderID, string strArticleID)
        {
            bool flag = false;
            List<Procedure> Prolist = new List<Procedure>();
            List<Dictionary<string, object>> ListParameter = new List<Dictionary<string, object>>();

            try
            {
                Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                sqlParameter.Clear();
                sqlParameter.Add("OrderID", strOrderID);
                sqlParameter.Add("ArticleID", strArticleID);
                sqlParameter.Add("PatternID", cboProcessPattern.SelectedValue.ToString());
                sqlParameter.Add("StuffOrderClosed", chkStuffClose.IsChecked == true ? "*" : "");
                sqlParameter.Add("LastUpdateUserID", MainWindow.CurrentUser);

                Procedure pro1 = new Procedure();
                pro1.Name = "xp_PlanInput_uOrderPatternID";
                pro1.OutputUseYN = "N";
                pro1.OutputName = "OrderID";
                pro1.OutputLength = "10";

                Prolist.Add(pro1);
                ListParameter.Add(sqlParameter);

                string[] Confirm = new string[2];
                Confirm = DataStore.Instance.ExecuteAllProcedureOutputNew(Prolist, ListParameter);
                if (Confirm[0] != "success")
                {
                    MessageBox.Show("[저장실패]\r\n" + Confirm[1].ToString());
                    flag = false;
                    //return false;
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
            int currRow = dgdPlanInput.Items.IndexOf(dgdPlanInput.CurrentItem);
            int currCol = dgdPlanInput.Columns.IndexOf(dgdPlanInput.CurrentCell.Column);
            int startCol = 7;
            int endCol = 12;

            if (e.Key == Key.Enter)
            {
                e.Handled = true;
                (sender as DataGridCell).IsEditing = false;

                // 마지막 열, 마지막 행 아님
                if (endCol == currCol && dgdPlanInput.Items.Count - 1 > currRow)
                {
                    dgdPlanInput.SelectedIndex = currRow + 1; // 이건 한줄 파란색으로 활성화 된 걸 조정하는 것입니다.
                    dgdPlanInput.CurrentCell = new DataGridCellInfo(dgdPlanInput.Items[currRow + 1], dgdPlanInput.Columns[startCol]);

                } // 마지막 열 아님
                else if (endCol > currCol && dgdPlanInput.Items.Count - 1 >= currRow)
                {
                    dgdPlanInput.CurrentCell = new DataGridCellInfo(dgdPlanInput.Items[currRow], dgdPlanInput.Columns[currCol + 1]);

                } // 마지막 열, 마지막 행
                else if (endCol == currCol && dgdPlanInput.Items.Count - 1 == currRow)
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
                if (dgdPlanInput.Items.Count - 1 > currRow)
                {
                    dgdPlanInput.SelectedIndex = currRow + 1;
                    dgdPlanInput.CurrentCell = new DataGridCellInfo(dgdPlanInput.Items[currRow + 1], dgdPlanInput.Columns[currCol]);
                } // 마지막 행일때
                else if (dgdPlanInput.Items.Count - 1 == currRow)
                {
                    if (endCol > currCol) // 마지막 열이 아닌 경우, 열을 오른쪽으로 이동
                    {
                        //dgdMain.SelectedIndex = 0;
                        dgdPlanInput.CurrentCell = new DataGridCellInfo(dgdPlanInput.Items[currRow], dgdPlanInput.Columns[currCol + 1]);
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
                    dgdPlanInput.SelectedIndex = currRow - 1;
                    dgdPlanInput.CurrentCell = new DataGridCellInfo(dgdPlanInput.Items[currRow - 1], dgdPlanInput.Columns[currCol]);
                } // 첫 행
                else if (dgdPlanInput.Items.Count - 1 == currRow)
                {
                    if (0 < currCol) // 첫 열이 아닌 경우, 열을 왼쪽으로 이동
                    {
                        //dgdMain.SelectedIndex = 0;
                        dgdPlanInput.CurrentCell = new DataGridCellInfo(dgdPlanInput.Items[currRow], dgdPlanInput.Columns[currCol - 1]);
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
                    dgdPlanInput.CurrentCell = new DataGridCellInfo(dgdPlanInput.Items[currRow], dgdPlanInput.Columns[currCol - 1]);
                }
                else if (startCol == currCol)
                {
                    if (0 < currRow)
                    {
                        dgdPlanInput.SelectedIndex = currRow - 1;
                        dgdPlanInput.CurrentCell = new DataGridCellInfo(dgdPlanInput.Items[currRow - 1], dgdPlanInput.Columns[endCol]);
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

                    dgdPlanInput.CurrentCell = new DataGridCellInfo(dgdPlanInput.Items[currRow], dgdPlanInput.Columns[currCol + 1]);
                }
                else if (endCol == currCol)
                {
                    if (dgdPlanInput.Items.Count - 1 > currRow)
                    {
                        dgdPlanInput.SelectedIndex = currRow + 1;
                        dgdPlanInput.CurrentCell = new DataGridCellInfo(dgdPlanInput.Items[currRow + 1], dgdPlanInput.Columns[startCol]);
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
            if (btnSave.IsEnabled == true)
            {
                int currCol = dgdPlanInput.Columns.IndexOf(dgdPlanInput.CurrentCell.Column);

                if (currCol > 6)
                {
                    DataGridCell cell = sender as DataGridCell;
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


        #region 지시수량 - 숫자만 입력 가능 하도록 / 시작일, 종료일 날짜 변경시

        private void dgdtpetxtOrderQty_PreviewTextInput(object sender, TextCompositionEventArgs e)
        {
            Lib.Instance.CheckIsNumeric((TextBox)sender, e);
        }

        //시작일
        private void dtpStartDate_SelectedDateChanged(object sender, SelectionChangedEventArgs e)
        {
            //Win_prd_PlanArticleOne_CodeView;
            DatePicker dtpSender = sender as DatePicker;
            var pInput = dtpSender.DataContext as Win_prd_PlanArticleOne_CodeView;

            if (pInput != null)
            {
                pInput.StartDate_CV = dtpSender.SelectedDate.Value.ToString("yyyy-MM-dd");
                pInput.StartDate = pInput.StartDate_CV.Replace("-", "");
            }
        }

        //종료일
        private void dtpEndDate_SelectedDateChanged(object sender, SelectionChangedEventArgs e)
        {
            //Win_prd_PlanArticleOne_CodeView;
            DatePicker dtpSender = sender as DatePicker;
            var pInput = dtpSender.DataContext as Win_prd_PlanArticleOne_CodeView;

            if (pInput != null)
            {
                pInput.EndDate_CV = dtpSender.SelectedDate.Value.ToString("yyyy-MM-dd");
                pInput.EndDate = pInput.StartDate_CV.Replace("-", "");
            }
        }

        #endregion

        #region 호기에서 키, 마우스 이벤트
        //호기
        private void dgdtpetxtMachine_KeyDown(object sender, KeyEventArgs e)
        {
            var WinPlanArticleOne = dgdPlanInput.CurrentItem as Win_prd_PlanArticleOne_CodeView;

            if (e.Key == Key.Enter)
            {
                TextBox tb1 = sender as TextBox;
                MainWindow.pf.ReturnCode(tb1, 79, WinPlanArticleOne.ProcessID);

                if (tb1.Tag != null)
                {
                    WinPlanArticleOne.Machine = tb1.Text;
                    WinPlanArticleOne.MachineID = tb1.Tag.ToString();
                }

                sender = tb1;
            }
        }

        //호기
        private void dgdtpetxtMachine_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            var WinPlanArticleOne = dgdPlanInput.CurrentItem as Win_prd_PlanArticleOne_CodeView;

            TextBox tb1 = sender as TextBox;
            MainWindow.pf.ReturnCode(tb1, 66, WinPlanArticleOne.ProcessID);

            if (tb1.Tag != null)
            {
                WinPlanArticleOne.Machine = tb1.Text;
                WinPlanArticleOne.MachineID = tb1.Tag.ToString();
            }

            sender = tb1;
        }
        #endregion

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

        //전월
        private void btnLastMonth_Click(object sender, RoutedEventArgs e)
        {
            try
            {
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
            catch (Exception ee)
            {
                MessageBox.Show("오류지점 - btnLastMonth_Click : " + ee.ToString());
            }
        }

        //전일
        private void btnYesterday_Click(object sender, RoutedEventArgs e)
        {
            try
            {
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
            catch (Exception ee)
            {
                MessageBox.Show("오류지점 - btnYesterday_Click : " + ee.ToString());
            }
        }

        private void txtQty_KeyDown(object sender, KeyEventArgs e)
        {
            if(e.Key == Key.Enter)
            {
                cboProcessPattern.Focus();
                cboProcessPattern.IsDropDownOpen = true;
            }
        }

        private void InsertEnable()
        {
            GridInputArea1.IsEnabled = true;
        }

        private void InsertDisable()
        {
            GridInputArea1.IsEnabled = false;
        }

        #region 입력창 이동 이벤트
        //텍스트박스
        private void EnterMoveTextBox_KeyDown(object sender, KeyEventArgs e)
        {
            try
            {
                if (e.Key == Key.Enter)
                {
                    lib.SendK(Key.Tab, this);
                }
            }
            catch (Exception ee)
            {
                MessageBox.Show("예외처리 - " + ee.ToString());
            }
        }

        //콤보박스일때
        private void EnterMoveComboBox_DropDownOrderClosed(object sender, EventArgs e)
        {
            try
            {
                lib.SendK(Key.Tab, this);
            }
            catch (Exception ee)
            {
                MessageBox.Show("예외처리 - " + ee.ToString());
            }
        }
        #endregion

        private void DataGridComboBoxMtrExceptYN_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            var WPPC = dgdPlanInput.SelectedItem as Win_prd_PlanArticleOne_CodeView;
            ComboBox comboBoxMtrExceptYN = sender as ComboBox;

            if(WPPC.MtrExceptYN != null)
            {
                WPPC.MtrExceptYN = comboBoxMtrExceptYN.SelectedValue.ToString();
            }

        }

        private void DataGridComboBoxMtrExceptYN_Loaded(object sender, RoutedEventArgs e)
        {
            ComboBox comboBoxMtrExceptYN = sender as ComboBox;

            comboBoxMtrExceptYN.Items.Add("N");
            comboBoxMtrExceptYN.Items.Add("Y");
        }

        private void DataGridComboBoxMtrExceptYN_KeyDown(object sender, KeyEventArgs e)
        {
            if(e.Key == Key.Enter)
            {
                DataGridCell cell = lib.GetParent<DataGridCell>(sender as ComboBox);
                cell.Focus();
            }
        }

        private void DataGridComboBoxFirstInFirstOutYN_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            var WPPC = dgdPlanInput.SelectedItem as Win_prd_PlanArticleOne_CodeView;
            ComboBox comboBoxFirstInFirstOutYN = sender as ComboBox;

            if(WPPC.FirstInFirstOutYN != null)
            {
                WPPC.FirstInFirstOutYN = comboBoxFirstInFirstOutYN.SelectedValue.ToString();
            }
        }

        private void DataGridComboBoxFirstInFirstOutYN_Loaded(object sender, RoutedEventArgs e)
        {
            ComboBox comboBoxFirstInFirstOutYN = sender as ComboBox;

            comboBoxFirstInFirstOutYN.Items.Add("N");
            comboBoxFirstInFirstOutYN.Items.Add("Y");
        }

        private void DataGridComboBoxFirstInFirstOutYN_KeyDown(object sender, KeyEventArgs e)
        {
            if(e.Key == Key.Enter)
            {
                DataGridCell cell = lib.GetParent<DataGridCell>(sender as ComboBox);
                cell.Focus();
            }
        }

        private void EnterMoveComboBox_DropDownClosed(object sender, EventArgs e)
        {
            try
            {
                lib.SendK(Key.Tab, this);
            }
            catch (Exception ee)
            {
                MessageBox.Show("예외처리 - " + ee.ToString());
            }
        }
    }

    #region CodeView
    class Win_prd_PlanInput_U_CodeView
    {
        public override string ToString()
        {
            return (this.ReportAllProperties());
        }

        public int Num { get; set; }
        public string KCustom { get; set; }
        public string Article { get; set; }
        public string OrderID { get; set; }
        public string OrderNo { get; set; }
        public double OrderQty { get; set; }
        public double notOrderInstQty { get; set; }
        public double OrderInstQy { get; set; }
        public double p1WorkQty { get; set; }
        public string p1ProcessID { get; set; }
        public string p1ProcessName { get; set; }
        public double InspectQty { get; set; }
        public double OutQty { get; set; }
        public string PatternID { get; set; }
        public string ArticleGrpID { get; set; }
        public string BuyerModel { get; set; }
        public string BuyerModelID { get; set; }
        public string BuyerArticleNo { get; set; }
        public string Remark { get; set; }
        public string ArticleID { get; set; }
        public string OrderClosed { get; set; }
        public string Completed { get; set; }

        public string ArticleGrpName { get; set; }
        public string InstID { get; set; }

        public string AcptDate { get; set; } // 수주일자
        public string AcptDate_CV { get; set; }
        public string OrderSeq { get; set; }

        //public string Remark { get; set; }
        //public string ArticleID { get; set; }
        //public string Article_Sabun { get; set; }
        //public string OrderClosed { get; set; }
        //public string Completed { get; set; }
        //public string subCompleted { get; set; }
        //public string cboPatternID { get; set; }
        //public string PartGBNID { get; set; }
        //public string AcptDate { get; set; }

    }

    class Win_prd_PlanArticleOne_CodeView
    {
        public int Num { get; set; }
        public string PatternSeq { get; set; }
        public string ProcessID { get; set; }
        public string Process { get; set; }
        public string Qty { get; set; }
        public string ArticleID { get; set; }
        public string Article { get; set; }
        public string BuyerArticleNo { get; set; }
        public string LVL { get; set; }

        public string InstQty { get; set; }
        public string InstRemark { get; set; }

        public string StartDate { get; set; }
        public string StartDate_CV { get; set; }

        public string EndDate { get; set; }
        public string EndDate_CV { get; set; }

        public string Remark { get; set; }

        public string MachineID { get; set; }
        public string Machine { get; set; }

        public string InstDate { get; set; }
        public string InstDate_CV { get; set; }
        public string ChildBuyerArticleNo { get; set; }
        public string MtrExceptYN { get; set; }
        public string FirstInFirstOutYN { get; set; }
        
    }

    #endregion
}

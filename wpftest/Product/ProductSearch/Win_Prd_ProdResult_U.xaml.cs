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
using WizMes_WellMade.PopUP;
using WPF.MDI;

namespace WizMes_WellMade
{
    /// <summary>
    /// Win_Prd_ProdResult_U.xaml에 대한 상호 작용 논리
    /// </summary>
    public partial class Win_Prd_ProdResult_U : UserControl
    {
        string stDate = string.Empty;
        string stTime = string.Empty;

        Win_Prd_ProdResult_U_CodeView prodResult = new Win_Prd_ProdResult_U_CodeView();
        Lib lib = new Lib();
        int rowNum = 0;
        // 수정 → 저장 시 JobId 로 수정한 행 찾아가기 위한 변수
        string jobID = "";

        string strFlag = "";

        public Win_Prd_ProdResult_U()
        {
            InitializeComponent();
        }

        private void UserControl_Loaded(object sender, RoutedEventArgs e)
        {
            stDate = DateTime.Now.ToString("yyyyMMdd");
            stTime = DateTime.Now.ToString("HHmm");

            DataStore.Instance.InsertLogByFormS(this.GetType().Name, stDate, stTime, "S");

            Lib.Instance.UiLoading(sender);

            chkDay.IsChecked = true;
            dtpSDate.SelectedDate = DateTime.Today;
            dtpEDate.SelectedDate = DateTime.Today;

            SetCombox();

            cboProcessSearch.SelectedIndex = 0;
            cboMachineSearch.SelectedIndex = 0;
        }

        #region 추가, 수정모드 / 저장완료, 취소 모드

        // 추가, 수정모드 
        private void SaveUpdateMode()
        {
            Lib.Instance.UiButtonEnableChange_SCControl(this);

            chkDay.IsEnabled = false;
            dtpSDate.IsEnabled = false;
            dtpEDate.IsEnabled = false;

            //생산일자 
            dtpProdDate.IsEnabled = true;
            txtProdScanTime.IsEnabled = true;

            //작업일시
            dtpWorkStartDate.IsEnabled = true;
            txtStartTime.IsEnabled = true;
            dtpWorkEndDate.IsEnabled = true;
            txtEndTime.IsEnabled = true;
            txtWorkMinute.IsEnabled = true;

            // 작업조 
            cboDayOrNight.IsEnabled = true;

            // 작업자, 호기 수정 가능 하도록
            txtWorker.IsEnabled = true;
            cboMachine.IsEnabled = true;
          
            SaveUpdateHeaderFalseMode();
        }

        private void SaveUpdateHeaderFalseMode()
        {
            lblWorkerName.IsEnabled = false;
            lblDay.IsEnabled = false;
            dtpEDate.IsEnabled = false;
            btnLastMonth.IsEnabled = false;
            btnThisMonth.IsEnabled = false;
            btnYesterDay.IsEnabled = false;
            btnToday.IsEnabled = false;
            lblCustom.IsEnabled = false;
            lblArticle.IsEnabled = false;
            lblModel.IsEnabled = false;
            lblProcess.IsEnabled = false;
            lblMachine.IsEnabled = false;
            lblGubun.IsEnabled = false;
            lblDefectWork.IsEnabled = false;
            dtpSDate.IsEnabled = false;
            cboProcessSearch.IsEnabled = false;
            cboMachineSearch.IsEnabled = false;
        }

        private void SaveUpdateHeaderMode()
        {
            lblWorkerName.IsEnabled = true;
            lblDay.IsEnabled = true;
            dtpEDate.IsEnabled = true;
            btnLastMonth.IsEnabled = true;
            btnThisMonth.IsEnabled = true;
            btnYesterDay.IsEnabled = true;
            btnToday.IsEnabled = true;
            lblCustom.IsEnabled = true;
            lblArticle.IsEnabled = true;
            lblModel.IsEnabled = true;
            lblProcess.IsEnabled = true;
            lblMachine.IsEnabled = true;
            lblGubun.IsEnabled = true;
            lblDefectWork.IsEnabled = true;
            dtpSDate.IsEnabled = true;
            cboProcessSearch.IsEnabled = true;
            cboMachineSearch.IsEnabled = true;
        }


        // 저장완료, 취소 모드
        private void CompleteCancelMode()
        {
            Lib.Instance.UiButtonEnableChange_IUControl(this);

            dtpProdDate.IsEnabled = false;
            txtProdScanTime.IsEnabled = false;

            dtpWorkStartDate.IsEnabled = false;
            txtStartTime.IsEnabled = false;
            dtpWorkEndDate.IsEnabled = false;
            txtEndTime.IsEnabled = false;
            txtWorkMinute.IsEnabled = false;

            //dgdResult.IsEnabled = true;
            dgdResult.IsHitTestVisible = true;

            // 작업조 비활성화
            cboDayOrNight.IsEnabled = false;
            txtCT.IsEnabled = false;

            // 작업자, 호기 수정 가능 하도록
            txtWorker.IsEnabled = false;
            btnPfWorker.IsEnabled = false;
            cboMachine.IsEnabled = false;

            SaveUpdateHeaderMode();

        }

        #endregion

        #region 콤보박스 세팅 SetCombox()

        private void SetCombox()
        {
            List<string> strCombo1 = new List<string>();
            strCombo1.Add("전체");
            strCombo1.Add("정상");
            strCombo1.Add("무작업");
            strCombo1.Add("재작업");
            ObservableCollection<CodeView> ovcGugunSearch = ComboBoxUtil.Instance.Direct_SetComboBox(strCombo1);
            this.cboGubunSearch.ItemsSource = ovcGugunSearch;
            this.cboGubunSearch.DisplayMemberPath = "code_name";
            this.cboGubunSearch.SelectedValuePath = "code_id";

            // 구분
            List<string[]> strCombo2 = new List<string[]>();
            string[] strVal1 = { "1", "정상" };
            string[] strVal2 = { "2", "무작업" };
            string[] strVal3 = { "3", "재작업" };
            strCombo2.Add(strVal1);
            strCombo2.Add(strVal2);
            strCombo2.Add(strVal3);
            ObservableCollection<CodeView> ovcGugun = ComboBoxUtil.Instance.Direct_SetComboBox(strCombo2);
            this.cbGubun.ItemsSource = ovcGugun;
            this.cbGubun.DisplayMemberPath = "code_name";
            this.cbGubun.SelectedValuePath = "code_id";


            // 공정
            ObservableCollection<CodeView> ovcProcess = ComboBoxUtil.Instance.GetWorkProcess(0, "");
            this.cboProcessSearch.ItemsSource = ovcProcess;
            this.cboProcessSearch.DisplayMemberPath = "code_name";
            this.cboProcessSearch.SelectedValuePath = "code_id";

            this.cboProcess.ItemsSource = ovcProcess;
            this.cboProcess.DisplayMemberPath = "code_name";
            this.cboProcess.SelectedValuePath = "code_id";


            // 설비
            ObservableCollection<CodeView> ovcMachine = ComboBoxUtil.Instance.GetMachine("");
            this.cboMachine.ItemsSource = ovcMachine;
            this.cboMachine.DisplayMemberPath = "code_name";
            this.cboMachine.SelectedValuePath = "code_id";

            ObservableCollection<CodeView> ovcMachineS = GetMachineByProcessID("");
            this.cboMachineSearch.ItemsSource = ovcMachineS;
            this.cboMachineSearch.DisplayMemberPath = "code_name";
            this.cboMachineSearch.SelectedValuePath = "code_id";

            // 주간 / 야간
            ObservableCollection<CodeView> ovcDayOrNight = ComboBoxUtil.Instance.Gf_DB_CM_GetComCodeDataset(null, "DayOrNight", "Y", "", "");
            cboDayOrNight.ItemsSource = ovcDayOrNight;
            cboDayOrNight.DisplayMemberPath = "code_name";
            cboDayOrNight.SelectedValuePath = "code_id";

        }
        #endregion

        #region mt_Machine - 호기 세팅

        /// <summary>
        /// 호기ID 가져오기
        /// </summary>
        /// <param name="value"></param>
        /// <returns></returns>
        public ObservableCollection<CodeView> GetMachineByProcessID(string value)
        {
            //2021-10-25 공정 콤보박스에 전체가 선택되면 호기 공정 콤보박스 안되게 막기
            if (value.Equals(""))
            {
                cboMachineSearch.IsEnabled = false;
            }
            else
            {
                cboMachineSearch.IsEnabled = true;
            }

            ObservableCollection<CodeView> ovcMachine = new ObservableCollection<CodeView>();

            Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
            sqlParameter.Add("sProcessID", value);

            DataSet ds = DataStore.Instance.ProcedureToDataSet("xp_Process_sMachineForComboBoxAndUsing", sqlParameter, false);

            if (ds != null && ds.Tables.Count > 0)
            {
                DataTable dt = ds.Tables[0];
                if (dt.Rows.Count > 0)
                {
                    CodeView CV = new CodeView();
                    CV.code_id = "";
                    CV.code_name = "전체";
                    ovcMachine.Add(CV);

                    DataRowCollection drc = dt.Rows;

                    foreach (DataRow dr in drc)
                    {
                        CodeView mCodeView = new CodeView()
                        {
                            code_id = dr["Code"].ToString().Trim(),
                            code_name = dr["Name"].ToString().Trim()
                        };

                        ovcMachine.Add(mCodeView);
                    }
                }
            }

            return ovcMachine;
        }

        #endregion // mt_Machine - 호기 세팅

        private void dgdResultCount_LoadingRow(object sender, DataGridRowEventArgs e)
        {
            int index = e.Row.GetIndex();
            var Data = dgdResultCount.Items[index] as Win_Prd_ProdResult_U_CodeView;
            if (Data != null)
            {
                e.Row.Header = Data.JobGbnname;
            }

            //e.Row.Header = "정상생산";
        }

        #region Header 검색조건

        //수주일자
        private void lblDay_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            if (chkDay.IsEnabled == true)
            {
                if (chkDay.IsChecked == true) { chkDay.IsChecked = false; }
                else { chkDay.IsChecked = true; }
                chkDayClick();
            }
        }

        //수주일자 클릭시
        private void chkDay_Click(object sender, RoutedEventArgs e)
        {
            chkDayClick();
        }

        //수주일자 이벤트
        private void chkDayClick()
        {
            if (chkDay.IsEnabled == true)
            {
                if (chkDay.IsChecked == true)
                {
                    dtpSDate.IsEnabled = true;
                    dtpEDate.IsEnabled = true;
                }
                else
                {
                    dtpSDate.IsEnabled = false;
                    dtpEDate.IsEnabled = false;
                }
            }
        }


        //거래처
        private void lblCustom_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            if (chkCustom.IsChecked == true) { chkCustom.IsChecked = false; }
            else { chkCustom.IsChecked = true; }

            chkCustomClick();
        }

        //거래처 클릭시
        private void chkCustom_Click(object sender, RoutedEventArgs e)
        {
            chkCustomClick();
        }

        //거래처 이벤트
        private void chkCustomClick()
        {
            if (chkCustom.IsChecked == true)
            {
                txtCustomSeach.IsEnabled = true;
                btnPfCustom.IsEnabled = true;
                txtCustomSeach.Focus();
            }
            else
            {
                txtCustomSeach.IsEnabled = false;
                btnPfCustom.IsEnabled = false;
            }
        }

        //거래처
        private void txtCustomSeach_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                //pf.ReturnCode(txtCustomSeach, 0, "");
                MainWindow.pf.ReturnCode(txtCustomSeach, (int)Defind_CodeFind.DCF_CUSTOM, "");
            }
        }

        //거래처
        private void btnPfCustom_Click(object sender, RoutedEventArgs e)
        {
            //pf.ReturnCode(txtCustomSeach, 0, "");
            MainWindow.pf.ReturnCode(txtCustomSeach, (int)Defind_CodeFind.DCF_CUSTOM, "");
        }

        //품명
        private void lblArticle_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            if (chkArticle.IsChecked == true) { chkArticle.IsChecked = false; }
            else { chkArticle.IsChecked = true; }

            chkArticleClick();
        }

        //품명 클릭시
        private void chkArticle_Click(object sender, RoutedEventArgs e)
        {
            chkArticleClick();
        }

        //품명 이벤트
        private void chkArticleClick()
        {
            if (chkArticle.IsChecked == true)
            {
                txtArticleSearch.IsEnabled = true;
                btnPfArticle.IsEnabled = true;
                txtArticleSearch.Focus();
            }
            else
            {
                txtArticleSearch.IsEnabled = false;
                btnPfArticle.IsEnabled = false;
            }
        }

        //품명
        private void txtArticleSearch_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                MainWindow.pf.ReturnCode(txtArticleSearch, 77, "");
                //MainWindow.pf.ReturnCode(txtArticleSearch, (int)Defind_CodeFind.DCF_Article, "");
            }
        }

        //품명
        private void btnPfArticle_Click(object sender, RoutedEventArgs e)
        {
            MainWindow.pf.ReturnCode(txtArticleSearch, 77, "");
        }

        //관리번호
        private void lblOrder_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            if (chkOrder.IsChecked == true) { chkOrder.IsChecked = false; }
            else { chkOrder.IsChecked = true; }

            chkOrderClick();
        }

        //관리번호 클릭시
        private void chkOrder_Click(object sender, RoutedEventArgs e)
        {
            chkOrderClick();
        }

        //관리번호 이벤트
        private void chkOrderClick()
        {
            if (chkOrder.IsChecked == true)
            {
                txtOrderSearch.IsEnabled = true;
                txtOrderSearch.Focus();
            }
            else
            {
                txtOrderSearch.IsEnabled = false;
            }
        }

        //모델
        private void lblModel_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            if (chkModel.IsChecked == true) { chkModel.IsChecked = false; }
            else { chkModel.IsChecked = true; }

            chkModelClick(); //2021-04-23 검색조건 차종을 이동전표로 바뀌어서 PlusFind 안써도 됨
        }

        //모델 클릭시
        private void chkModel_Click(object sender, RoutedEventArgs e)
        {
            chkModelClick(); //2021-04-23 검색조건 차종을 이동전표로 바뀌어서 PlusFind 안써도 됨
        }

        //모델 이벤트 2021-04-23 검색조건 차종을 이동전표로 바뀌어서 PlusFind 안써도 됨
        //2022.03.28 이동전표 -> 모델 복원
        private void chkModelClick()
        {
            if (chkModel.IsChecked == true)
            {
                txtModelSearch.IsEnabled = true;
                btnPfModel.IsEnabled = true;
                txtModelSearch.Focus();
            }
            else
            {
                txtModelSearch.IsEnabled = false;
                btnPfModel.IsEnabled = false;
            }
        }

        //모델 2021-04-23 검색조건 차종을 이동전표로 바뀌어서 PlusFind 안써도 됨
        //2022.03.28 이동전표 -> 모델 복원
        private void txtModelSearch_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                rowNum = 0;
                //using (Loading lw = new Loading(re_Search))
                //{
                //    lw.ShowDialog();
                //}

                MainWindow.pf.ReturnCode(txtModelSearch, (int)Defind_CodeFind.DCF_BUYERMODEL, ""); 

            }
        }

        //모델
        private void btnPfModel_Click(object sender, RoutedEventArgs e)
        {
            MainWindow.pf.ReturnCode(txtModelSearch, (int)Defind_CodeFind.DCF_BUYERMODEL, "");
        }

        //품번
        private void lblArticleNo_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            if (chkArticleNo.IsChecked == true) { chkArticleNo.IsChecked = false; }
            else { chkArticleNo.IsChecked = true; }

            chkArticleNoClick();
        }

        //품번 클릭시
        private void chkArticleNo_Click(object sender, RoutedEventArgs e)
        {
            chkArticleNoClick();
        }

        //품번 이벤트
        private void chkArticleNoClick()
        {
            if (chkArticleNo.IsChecked == true)
            {
                txtArticleNoSearch.IsEnabled = true;
                txtArticleNoSearch.Focus();
            }
            else
            {
                txtArticleNoSearch.IsEnabled = false;
            }
        }

        //공정명
        private void lblProcess_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            if (chkProcess.IsChecked == true) { chkProcess.IsChecked = false; }
            else { chkProcess.IsChecked = true; }

            chkProcessClick();
        }

        //공정명 클릭시
        private void chkProcess_Click(object sender, RoutedEventArgs e)
        {
            chkProcessClick();
        }

        //공정명 이벤트
        private void chkProcessClick()
        {
            if (chkProcess.IsChecked == true)
            {
                cboProcessSearch.IsEnabled = true;
                cboProcessSearch.Focus();
            }
            else
            {
                cboProcessSearch.IsEnabled = false;
            }
        }

        private void cboProcessSearch_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (cboProcessSearch.SelectedValue != null)
            {
                ObservableCollection<CodeView> ovcMachine = GetMachineByProcessID(cboProcessSearch.SelectedValue.ToString());
                this.cboMachineSearch.ItemsSource = ovcMachine;
                this.cboMachineSearch.DisplayMemberPath = "code_name";
                this.cboMachineSearch.SelectedValuePath = "code_id";

                cboMachineSearch.SelectedIndex = 0;
            }
        }

        //호기
        private void lblMachine_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            if (chkMachine.IsChecked == true) { chkMachine.IsChecked = false; }
            else { chkMachine.IsChecked = true; }

            chkMachineClick();
        }

        //호기 클릭시
        private void chkMachine_Click(object sender, RoutedEventArgs e)
        {
            chkMachineClick();
        }

        //호기 이벤트
        private void chkMachineClick()
        {
            if (chkMachine.IsChecked == true)
            {
                cboMachineSearch.IsEnabled = true;
                cboMachineSearch.Focus();
            }
            else
            {
                cboMachineSearch.IsEnabled = false;
            }
        }

        //작업구분
        private void lblGubun_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            if (chkGubun.IsChecked == true) { chkGubun.IsChecked = false; }
            else { chkGubun.IsChecked = true; }

            GubunClick();
        }

        //작업구분 클릭시
        private void chkGubun_Click(object sender, RoutedEventArgs e)
        {
            GubunClick();
        }

        //작업구분 이벤트
        private void GubunClick()
        {
            if (chkGubun.IsChecked == true)
            {
                cboGubunSearch.IsEnabled = true;
                cboGubunSearch.Focus();
            }
            else
            {
                cboGubunSearch.IsEnabled = false;
            }
        }

        //불량발생건
        private void brDefectWork_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            if (chkDefectWork.IsChecked == true) { chkDefectWork.IsChecked = false; }
            else { chkDefectWork.IsChecked = true; }
        }

        //불량발생건
        private void lblDefectWork_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            if (chkDefectWork.IsChecked == true) { chkDefectWork.IsChecked = false; }
            else { chkDefectWork.IsChecked = true; }
        }

        #endregion

        #region Header 전일 금일 전월 금월, 날짜관련 클릭 이벤트

        //전일
        private void btnYesterDay_Click(object sender, RoutedEventArgs e)
        {
            DateTime[] SearchDate = lib.BringLastDayDateTimeContinue(dtpEDate.SelectedDate.Value);

            dtpSDate.SelectedDate = SearchDate[0];
            dtpEDate.SelectedDate = SearchDate[1];
        }

        //금일
        private void btnToday_Click(object sender, RoutedEventArgs e)
        {
            dtpSDate.SelectedDate = DateTime.Today;
            dtpEDate.SelectedDate = DateTime.Today;
        }

        // 전월 버튼 클릭 이벤트
        private void btnLastMonth_Click(object sender, RoutedEventArgs e)
        {
            DateTime[] SearchDate = lib.BringLastMonthContinue(dtpSDate.SelectedDate.Value);

            dtpSDate.SelectedDate = SearchDate[0];
            dtpEDate.SelectedDate = SearchDate[1];
        }

        // 금월 버튼 클릭 이벤트
        private void btnThisMonth_Click(object sender, RoutedEventArgs e)
        {
            dtpSDate.SelectedDate = lib.BringThisMonthDatetimeList()[0];
            dtpEDate.SelectedDate = lib.BringThisMonthDatetimeList()[1];
        }


        #endregion

        #region Header 오른상단 버튼 이벤트

        // 검색
        private void btnSearch_Click(object sender, RoutedEventArgs e)
        {

            //검색버튼 비활성화
            btnSearch.IsEnabled = false;

            Dispatcher.BeginInvoke(new Action(() =>

            {
                //로직
                rowNum = 0;
                using (Loading lw = new Loading(re_Search))
                {
                    lw.ShowDialog();
                }

            }), System.Windows.Threading.DispatcherPriority.Background);



            Dispatcher.BeginInvoke(new Action(() =>

            {
                btnSearch.IsEnabled = true;

            }), System.Windows.Threading.DispatcherPriority.Background);
            

        }

        // 수정
        private void btnUpdate_Click(object sender, RoutedEventArgs e)
        {
            if (dgdResult.Items.Count == 0)
            {
                MessageBox.Show("검색을 먼저 해주세요.");
                return;
            }

            var Main = dgdResult.SelectedItem as Win_Prd_ProdResult_U_CodeView;

            if (Main != null)
            {
                SaveUpdateMode();

                // 정렬이 생산일자 이기 때문에, 생산일자가 변경되면 행순서도 변경 되어 수정시 찾아갈 수가 없음.
                // JobId 로 찾아가도록 수정
                //rowNum = dgdResult.SelectedIndex;
                jobID = Main.JobID.Trim();
                strFlag = "U";

                cboDayOrNight.IsDropDownOpen = true;
            }
            else
            {
                MessageBox.Show("수정할 작업을 선택해주세요.");
            }
        }

        /// 삭제 버튼 클릭 이벤트
        private void btnDelete_Click(object sender, RoutedEventArgs e)
        {
            var WinProdU = dgdResult.SelectedItem as Win_Prd_ProdResult_U_CodeView;

            if (WinProdU == null)
            {
                MessageBox.Show("삭제할 데이터가 지정되지 않았습니다. 삭제데이터를 지정하고 눌러주세요");
                return;
            }
            else
            {
                if (MessageBox.Show("선택하신 항목을 삭제하시겠습니까?", "삭제 전 확인", MessageBoxButton.YesNo) == MessageBoxResult.Yes)
                {
                    using (Loading lw = new Loading(beDelete))
                    {
                        lw.ShowDialog();
                    }
                }
            }
        }

        // 삭제 메서드 모음
        private void beDelete()
        {
            var Main = dgdResult.SelectedItem as Win_Prd_ProdResult_U_CodeView;

            if (Main != null
                && Main.JobID != null
                && Main.SplitSeq != null)
            {
                //if (Main.IsSplitYN == true)
                //{
                //    MessageBox.Show("해당 잔량이동처리 건 작업 실적은 삭제가 불가능 합니다.");
                //    return;
                //}
                string msg = CheckIsNextWorkData(ConvertDouble(Main.JobID));

                if (msg.Length > 0) { MessageBox.Show(msg); return; }

                string prodmsg = CheckIsProd(ConvertDouble(Main.JobID)); //2021-10-28 재고 체크

                if (prodmsg.Length > 0) { MessageBox.Show(prodmsg); return; }

                if (DeleteData(Main.JobID, Main.SplitSeq))
                {
                    re_Search();
                }
            }
        }

        #region 해당 라벨의 후공정이 있는지 체크 하기, 있으면 삭제가 안되도록!

        string ChkNextWorkData_Msg = "";

        private string CheckIsNextWorkData(double JobID)
        {
            string result = "";

            try
            {
                Dictionary<string, object> sqlParameter = new Dictionary<string, object>();

                sqlParameter.Add("JobID", JobID);

                DataTable dt = DataStore.Instance.ProcedureToDataSet("xp_prdWork_CheckIsNextWorkData", sqlParameter, false).Tables[0];

                if (dt != null
                    && dt.Rows.Count > 0
                    && dt.Columns.Count == 1)
                {
                    string Msg = dt.Rows[0]["Result"].ToString().Trim();

                    if (Msg.ToUpper().Equals("PASS"))
                    {
                        //flag = true;
                    }
                    else
                    {
                        result += "해당 작업 이력은 후공정 작업 이력이 있어 삭제가 불가능합니다.\r\n" + Msg;
                    }
                }

                return result;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                return result;
            }
        }

        #endregion

        #region 해당 라벨의 현재고를 체크하여 0이면 삭제가 안되도록 막기
        private string CheckIsProd(double JobID)
        {
            string result = "";

            try
            {
                Dictionary<string, object> sqlParameter = new Dictionary<string, object>();

                sqlParameter.Add("JobID", JobID);

                DataTable dt = DataStore.Instance.ProcedureToDataSet("xp_prdWork_DeleteCheckProdCapa", sqlParameter, false).Tables[0];

                if (dt != null
                    && dt.Rows.Count > 0
                    && dt.Columns.Count == 1)
                {
                    string Msg = dt.Rows[0]["Result"].ToString().Trim();

                    if (Msg.ToUpper().Equals("PASS"))
                    {
                        //flag = true;
                    }
                    else
                    {
                        result += "해당 작업 라벨의 현 재고가 0입니다. 삭제 시 마이너스 재고가 발생합니다.\r\n";
                    }
                }

                return result;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                return result;
            }
        }
        #endregion

        /// 닫기
        private void btnClose_Click(object sender, RoutedEventArgs e)
        {
            DataStore.Instance.InsertLogByFormS(this.GetType().Name, stDate, stTime, "E");
            Lib.Instance.ChildMenuClose(this.ToString());
        }

        /// 저장
        private void btnSave_Click(object sender, RoutedEventArgs e)
        {
            using (Loading lw = new Loading(beSave))
            {
                lw.ShowDialog();
            }
        }

        private void beSave()
        {
            if (SaveData())
            {
                CompleteCancelMode();

                re_Search();
            }
        }

        /// 취소
        private void btnCancel_Click(object sender, RoutedEventArgs e)
        {
            CompleteCancelMode();

            using (Loading lw = new Loading(re_Search))
            {
                lw.ShowDialog();
            }
        }

        /// 엑셀
        private void btnExcel_Click(object sender, RoutedEventArgs e)
        {
            DataTable dt = null;
            string Name = string.Empty;

            string[] dgdStr = new string[2];
            dgdStr[0] = "일생산 상세현황";
            dgdStr[1] = dgdResult.Name;

            ExportExcelxaml ExpExc = new ExportExcelxaml(dgdStr);
            ExpExc.ShowDialog();

            if (ExpExc.DialogResult.HasValue)
            {
                if (ExpExc.choice.Equals(dgdResult.Name))
                {
                    DataStore.Instance.InsertLogByForm(this.GetType().Name, "E");
                    if (ExpExc.Check.Equals("Y"))
                        dt = Lib.Instance.DataGridToDTinHidden(dgdResult);
                    else
                        dt = Lib.Instance.DataGirdToDataTable(dgdResult);

                    Name = dgdResult.Name;
                    if (lib.GenerateExcel(dt, Name))
                        lib.excel.Visible = true;
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

        // 재조회
        private void re_Search()
        {
            FillGrid();

            if (dgdResult.Items.Count > 0)
            {
                if (strFlag.Trim().Equals("U")
                    && !jobID.Equals(""))
                {
                    for (int i = 0; i < dgdResult.Items.Count; i++)
                    {
                        var Main = dgdResult.Items[i] as Win_Prd_ProdResult_U_CodeView;
                        if (Main != null)
                        {
                            if (Main.JobID.Trim().Equals(jobID))
                            {
                                dgdResult.SelectedIndex = i;
                                break;
                            }
                        }
                    }
                }
                else
                {
                    dgdResult.SelectedIndex = rowNum;
                }

                if (dgdResult.SelectedIndex == -1)
                {
                    dgdResult.SelectedIndex = 0;
                }
            }
            else
            {
                MessageBox.Show("조회된 데이터가 없습니다.");
                return;
            }
        }

        #endregion


        #region 주요 메서드 - 메인그리드 조회 FillGrid()

        private void FillGrid()
        {
            if (dgdResult.Items.Count > 0)
            {
                dgdResult.Items.Clear();
            }

            if (dgdResultCount.Items.Count > 0)
            {
                dgdResultCount.Items.Clear();
            }

            try
            {
                // 공정 호기 세팅
                string ProcessID = "";
                string MachineID = "";

                // 공정을 전체나 선택하지 않았을시 → 호기는 공정 + 호기로 출력 → 공정과 호기를 검색하기 위해서
                if (chkMachine.IsChecked == true
                    && cboMachineSearch.SelectedValue != null
                    && cboMachineSearch.SelectedValue.ToString().Trim().Length == 6)
                {
                    ProcessID = cboMachineSearch.SelectedValue.ToString().Trim().Substring(0, 4);
                    MachineID = cboMachineSearch.SelectedValue.ToString().Trim().Substring(4, 2);
                }
                else
                {
                    ProcessID = chkProcess.IsChecked == true && cboProcessSearch.SelectedValue != null ? cboProcessSearch.SelectedValue.ToString() : "";
                    MachineID = chkMachine.IsChecked == true && cboMachineSearch.SelectedValue != null ? cboMachineSearch.SelectedValue.ToString() : "";
                }


                Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                sqlParameter.Add("ChkDate", chkDay.IsChecked == true ? 1 : 0);
                sqlParameter.Add("sFromDate", chkDay.IsChecked == true && dtpSDate.SelectedDate != null ? dtpSDate.SelectedDate.Value.ToString("yyyyMMdd") : "");
                sqlParameter.Add("sToDate", chkDay.IsChecked == true && dtpEDate.SelectedDate != null ? dtpEDate.SelectedDate.Value.ToString("yyyyMMdd") : "");
                sqlParameter.Add("sProcessID", ProcessID);
                sqlParameter.Add("sMachineID", MachineID);
                sqlParameter.Add("ArticleID", chkArticle.IsChecked == true && txtArticleSearch.Tag != null ? txtArticleSearch.Tag.ToString() : "");

                sqlParameter.Add("CustomID", chkCustom.IsChecked == true && txtCustomSeach.Tag != null ? txtCustomSeach.Tag.ToString() : "");
                sqlParameter.Add("nOrderID", (chkOrder.IsChecked == true ? (tbkOrder.Text.Equals("관리번호") ? 1 : 2) : 0));
                sqlParameter.Add("sOrderID", (chkOrder.IsChecked == true ? txtOrderSearch.Text : ""));
                sqlParameter.Add("nJobGbn", (chkGubun.IsChecked == true ? 1 : 0));
                sqlParameter.Add("sJobGubun", chkGubun.IsChecked == true && cboGubunSearch.SelectedValue != null ? cboGubunSearch.SelectedValue.ToString() : "");

                sqlParameter.Add("nBuyerModel", ((chkModel.IsChecked == true) ? 1 : 0));
                sqlParameter.Add("sBuyerModel", chkModel.IsChecked == true ? txtModelSearch.Text : ""); //2021-04-23 PlusFinder 안써서 수정 ("sBuyerModel", chkModel.IsChecked == true && txtModelSearch.Tag != null ? txtModelSearch.Tag.ToString() : "")
                sqlParameter.Add("nBuyerArticleNo", chkArticleNo.IsChecked == true ? 1 : 0);
                sqlParameter.Add("sBuyerArticleNo", chkArticleNo.IsChecked == true ? txtArticleNoSearch.Text : "");
                sqlParameter.Add("ndefect", chkDefectWork.IsChecked == true ? 1 : 0);
                sqlParameter.Add("nWorkerName", chkWorkerName.IsChecked == true ? 1 : 0);
                sqlParameter.Add("sWorkerName", chkWorkerName.IsChecked == true ? txtWorkerNameSearch.Text : "");

                DataSet ds = DataStore.Instance.ProcedureToDataSet_LogWrite("xp_prd_sWKResult_WPF", sqlParameter, true, "R");

                if (ds != null && ds.Tables.Count > 0)
                {
                    DataTable dt = ds.Tables[0];
                    int i = 1;

                    if (dt.Rows.Count > 0)
                    {
                        DataRowCollection drc = dt.Rows;

                        foreach (DataRow dr in drc)
                        {
                            if (dr["JobGbn"].ToString() != "2") //2021-05-06 비가동일때는 작업지시가 없어 품명, 품번 등이 없어 빈값으로 가야되서 경우의 수를 나눴음
                            {
                                var Grid_DTO = new Win_Prd_ProdResult_U_CodeView()
                                {
                                    Num = i,
                                    cls = (int)dr["cls"],
                                    WorkDate = dr["WorkDate"].ToString(),
                                    WorkDate_CV = DatePickerFormat(dr["WorkDate"].ToString()),
                                    ProcessID = dr["ProcessID"].ToString(),
                                    Process = dr["Process"].ToString(),
                                    OrderID = dr["OrderID"].ToString(), //20210504
                                    OrderNo = dr["OrderNo"].ToString(), //20210504
                                    AcptDate = dr["AcptDate"].ToString(),
                                    OrderQty = Convert.ToDouble(dr["OrderQty"]), //20210504

                                    MachineID = dr["MachineID"].ToString(),
                                    InstDate = dr["InstDate"].ToString(),
                                    InstQty = Convert.ToDouble(dr["InstQty"]),
                                    WorkQty = Convert.ToDouble(dr["WorkQty"]),
                                    WorkPersonID = dr["WorkPersonID"].ToString(),

                                    BuyerModelID = dr["BuyerModelID"].ToString(),
                                    BuyerModel = dr["BuyerModel"].ToString(),
                                    CustomID = dr["CustomID"].ToString(), //20210504
                                    KCustom = dr["KCustom"].ToString(), //20210504
                                    Worker = dr["Worker"].ToString(),

                                    Article = dr["Article"].ToString(), //20210504
                                    LabelID = dr["LabelID"].ToString(),
                                    JobGbn = dr["JobGbn"].ToString(),
                                    JobGbnname = dr["JobGbnname"].ToString(),
                                    WorkStartDate = dr["WorkStartDate"].ToString(),

                                    WorkStartDate_CV = DatePickerFormat(dr["WorkStartDate"].ToString()),
                                    WorkStartTime = (dr["WorkStartTime"].ToString() != null ? Lib.Instance.SixLengthTime(dr["WorkStartTime"].ToString().Replace(" ", "")) : ""),
                                    WorkEndDate = dr["WorkEndDate"].ToString(),
                                    WorkEndDate_CV = DatePickerFormat(dr["WorkEndDate"].ToString()),
                                    WorkEndTime = (dr["WorkEndTime"].ToString() != null ? Lib.Instance.SixLengthTime(dr["WorkEndTime"].ToString().Replace(" ", "")) : ""),
                                    WorkHour = stringFormatN0(dr["WorkHour"]),
                                    WorkMinute = stringFormatN0(dr["WorkMinute"]),

                                    JobID = dr["JobID"].ToString(),
                                    Articleid = dr["Articleid"].ToString(),
                                    WorkCnt = stringFormatN0(dr["WorkCnt"]),
                                    NoReworkCode = dr["NoReworkCode"].ToString(),
                                    NoReworkName = dr["NoReworkName"].ToString(),

                                    FourMID = dr["4MID"].ToString(),
                                    FourMSubject = dr["4MSubject"].ToString(),
                                    DayOrNightID = dr["DayOrNightID"].ToString(),
                                    CycleTime = dr["CycleTime"].ToString(),
                                    BuyerArticleNo = dr["BuyerArticleNo"].ToString(), //20210504

                                    WorkTime = stringFormatN0(dr["WorkTime"]),
                                    ScanDate = dr["ScanDate"].ToString(),
                                    ScanDate_CV = DatePickerFormat(dr["ScanDate"].ToString()),

                                    ScanTime = (dr["ScanTime"].ToString() != null ? Lib.Instance.SixLengthTime(dr["ScanTime"].ToString().Replace(" ", "")) : ""),
                                    IsSplitYN = dr["IsSplitYN"] != null && dr["IsSplitYN"].ToString().Trim().Equals("Y") ? true : false,
                                    SplitSeq = dr["SplitSeq"].ToString(),
                                    //gbn = dr["gbn"].ToString(),

                                    WorkTimeMinute = Convert.ToDouble(dr["WorkTimeMinute"]),

                                };

                                if (Grid_DTO.cls == 1)
                                {
                                    dgdResult.Items.Add(Grid_DTO);
                                    i++;
                                }
                                else if (Grid_DTO.cls == 4)
                                {
                                    dgdResultCount.Items.Add(Grid_DTO);
                                }
                                else if(Grid_DTO.cls == 8)
                                {
                                    dgdResult.Items.Add(Grid_DTO);
                                }
                                else if (Grid_DTO.cls == 9)
                                {
                                    dgdResult.Items.Add(Grid_DTO);
                                }

                                Grid_DTO.WorkStartDate = DatePickerFormat(Grid_DTO.WorkStartDate);
                                Grid_DTO.WorkEndDate = DatePickerFormat(Grid_DTO.WorkEndDate);
                            }
                            else
                            {
                                var Grid_DTO = new Win_Prd_ProdResult_U_CodeView()
                                {
                                    Num = i,
                                    cls = (int)dr["cls"],
                                    WorkDate = dr["WorkDate"].ToString(),
                                    WorkDate_CV = DatePickerFormat(dr["WorkDate"].ToString()),
                                    ProcessID = dr["ProcessID"].ToString(),
                                    Process = dr["Process"].ToString(),
                                    OrderID = "", //dr["OrderID"].ToString(), //20210504
                                    OrderNo = "", //dr["OrderNo"].ToString(), //20210504
                                    AcptDate = dr["AcptDate"].ToString(),
                                    OrderQty = 0,//stringFormatN0(dr["OrderQty"]), //20210504

                                    MachineID = dr["MachineID"].ToString(),
                                    InstDate = dr["InstDate"].ToString(),
                                    InstQty = Convert.ToDouble(dr["InstQty"]),
                                    WorkQty = Convert.ToDouble(dr["WorkQty"]),
                                    WorkPersonID = dr["WorkPersonID"].ToString(),

                                    BuyerModelID = dr["BuyerModelID"].ToString(),
                                    BuyerModel = dr["BuyerModel"].ToString(),
                                    CustomID = "",//dr["CustomID"].ToString(), //20210504
                                    KCustom = "",//dr["KCustom"].ToString(), //20210504
                                    Worker = dr["Worker"].ToString(),

                                    Article = "",//dr["Article"].ToString(), //20210504
                                    LabelID = dr["LabelID"].ToString(),
                                    JobGbn = dr["JobGbn"].ToString(),
                                    JobGbnname = dr["JobGbnname"].ToString(),
                                    WorkStartDate = dr["WorkStartDate"].ToString(),

                                    WorkStartDate_CV = DatePickerFormat(dr["WorkStartDate"].ToString()),
                                    WorkStartTime = (dr["WorkStartTime"].ToString() != null ? Lib.Instance.SixLengthTime(dr["WorkStartTime"].ToString().Replace(" ", "")) : ""),
                                    WorkEndDate = dr["WorkEndDate"].ToString(),
                                    WorkEndDate_CV = DatePickerFormat(dr["WorkEndDate"].ToString()),
                                    WorkEndTime = (dr["WorkEndTime"].ToString() != null ? Lib.Instance.SixLengthTime(dr["WorkEndTime"].ToString().Replace(" ", "")) : ""),
                                    WorkHour = stringFormatN0(dr["WorkHour"]),
                                    WorkMinute = stringFormatN0(dr["WorkMinute"]),

                                    JobID = dr["JobID"].ToString(),
                                    Articleid = dr["Articleid"].ToString(),
                                    WorkCnt = stringFormatN0(dr["WorkCnt"]),
                                    NoReworkCode = dr["NoReworkCode"].ToString(),
                                    NoReworkName = dr["NoReworkName"].ToString(),

                                    FourMID = dr["4MID"].ToString(),
                                    FourMSubject = dr["4MSubject"].ToString(),
                                    DayOrNightID = dr["DayOrNightID"].ToString(),
                                    CycleTime = dr["CycleTime"].ToString(),
                                    BuyerArticleNo = "",//dr["BuyerArticleNo"].ToString(), //20210504

                                    WorkTime = stringFormatN0(dr["WorkTime"]),
                                    ScanDate = dr["ScanDate"].ToString(),
                                    ScanDate_CV = DatePickerFormat(dr["ScanDate"].ToString()),

                                    ScanTime = (dr["ScanTime"].ToString() != null ? Lib.Instance.SixLengthTime(dr["ScanTime"].ToString().Replace(" ", "")) : ""),
                                    IsSplitYN = dr["IsSplitYN"] != null && dr["IsSplitYN"].ToString().Trim().Equals("Y") ? true : false,
                                    SplitSeq = dr["SplitSeq"].ToString(),

                                    WorkTimeMinute = Convert.ToDouble(dr["WorkTimeMinute"]),

                                };

                                if (Grid_DTO.cls == 1)
                                {
                                    dgdResult.Items.Add(Grid_DTO);
                                    i++;
                                }
                                else if (Grid_DTO.cls == 4)
                                {
                                    dgdResultCount.Items.Add(Grid_DTO);
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

        #endregion // 주요 메서드 - 메인그리드 조회 FillGrid()

        #region 주요메서드 - 합계 구하기? FillGridCount

        private void FillGridCount()
        {
            if (dgdResultCount.Items.Count > 0)
            {
                dgdResultCount.Items.Clear();
            }

            double sumQty = 0.0;
            int sumhour = 0;
            int sumCount = 0;

            if (dgdResult != null && dgdResult.Items.Count > 0)
            {
                for (int i = 0; i < dgdResult.Items.Count; i++)
                {
                    var prodResult = dgdResult.Items[i] as Win_Prd_ProdResult_U_CodeView;

                    if (Lib.Instance.IsNumOrAnother(prodResult.WorkQty.ToString()))
                    {
                        double qty = double.Parse(prodResult.WorkQty.ToString());
                        sumQty += qty;
                    }

                    if (Lib.Instance.IsNumOrAnother(prodResult.WorkHour))
                    {
                        int hour = int.Parse(prodResult.WorkHour);
                        sumhour += hour;
                    }

                    sumCount++;
                }

                var Count_DTO = new ProdResult_Count()
                {
                    ctGubun = prodResult.JobGbnname,
                    ctCount = sumCount.ToString(),
                    ctQty = sumQty.ToString(),
                    ctWorkHour = sumhour.ToString()
                };

                dgdResultCount.Items.Add(Count_DTO);
            }
        }

        #endregion // 주요메서드 - 합계 구하기? FillGridCount

        #region 주요메서드 - 하위 정보 FillGridChild

        private void FillGridChild(int jobID)
        {
            if (dgdResultChild.Items.Count > 0)
            {
                dgdResultChild.Items.Clear();
            }

            try
            {
                Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                sqlParameter.Add("nJobID", jobID);

                DataSet ds = DataStore.Instance.ProcedureToDataSet("xp_wkResult_sWKResultOneChild", sqlParameter, false);

                if (ds != null && ds.Tables.Count > 0)
                {
                    DataTable dt = ds.Tables[0];

                    if (dt.Rows.Count > 0)
                    {
                        DataRowCollection drc = dt.Rows;

                        foreach (DataRow dr in drc)
                        {
                            var Child_DTO = new ProdResult_Child()
                            {
                                ChildLabelID = dr["ChildLabelID"].ToString(),
                                ChildArticleID = dr["ChildArticleID"].ToString(),
                                Article = dr["Article"].ToString()
                            };

                            dgdResultChild.Items.Add(Child_DTO);
                        }
                        drc.Clear();
                    }
                    dt.Clear();
                }
                ds.Clear();
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

        #endregion // 주요메서드 - 하위 정보 FillGridChild

        #region 주요 메서드 - 불량정보 조회 FillGridDefect

        private void FillGridDefect(int jobID)
        {
            if (dgdResultDefect.Items.Count > 0)
            {
                dgdResultDefect.Items.Clear();
            }

            try
            {
                Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                sqlParameter.Add("nJobID", jobID);

                DataSet ds = DataStore.Instance.ProcedureToDataSet("xp_wkResult_sWKResultOneDefect", sqlParameter, false);

                if (ds != null && ds.Tables.Count > 0)
                {
                    DataTable dt = ds.Tables[0];

                    if (dt.Rows.Count > 0)
                    {
                        DataRowCollection drc = dt.Rows;

                        foreach (DataRow dr in drc)
                        {
                            var Defect_DTO = new ProdResult_Defect()
                            {
                                DefectID = dr["DefectID"].ToString(),
                                KDefect = dr["KDefect"].ToString(),
                                DefectCount = dr["DefectCount"].ToString()
                            };

                            Defect_DTO.DefectCount = Lib.Instance.returnNumStringZero(Defect_DTO.DefectCount);
                            dgdResultDefect.Items.Add(Defect_DTO);
                        }
                        drc.Clear();
                    }
                    dt.Clear();
                }
                ds.Clear();
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

        #endregion // 주요 메서드 - 불량정보 조회 FillGridDefect

        #region 주요 메서드 - 작업 정보 수정
        private bool SaveData()
        {
            bool flag = false;

            try
            {
                if (CheckData())
                {
                    Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                    sqlParameter.Add("JobID", prodResult.JobID);
                    sqlParameter.Add("ScanDate", dtpProdDate.SelectedDate.Value.ToString("yyyyMMdd"));
                    sqlParameter.Add("ScanTime", DateTimeFormat(txtProdScanTime.Text));
                    sqlParameter.Add("WorkStartDate", dtpWorkStartDate.SelectedDate != null ? dtpWorkStartDate.SelectedDate.Value.ToString("yyyyMMdd") : "");
                    sqlParameter.Add("WorkStartTime", DateTimeFormat(txtStartTime.Text));
                    sqlParameter.Add("WorkEndDate", dtpWorkEndDate.SelectedDate != null ? dtpWorkEndDate.SelectedDate.Value.ToString("yyyyMMdd") : "");
                    sqlParameter.Add("WorkEndTime", dtpWorkEndDate.SelectedDate != null ? DateTimeFormat(txtEndTime.Text) : "");
                    sqlParameter.Add("WorkQty", (Lib.Instance.IsIntOrAnother(txtQty.Text) ? int.Parse(txtQty.Text) : 0));
                    sqlParameter.Add("DayOrNightID", cboDayOrNight.SelectedValue != null ? cboDayOrNight.SelectedValue.ToString() : "");
                    sqlParameter.Add("CycleTime", ConvertDouble(txtCT.Text));
                    sqlParameter.Add("MachineID", cboMachine.SelectedValue.ToString());
                    sqlParameter.Add("WorkPersonID", txtWorker.Tag.ToString());
                    sqlParameter.Add("IsSplit", prodResult.IsSplitYN == true ? 1 : 0); // 이것이 스플릿인가 아닌가
                    sqlParameter.Add("SplitSeq", ConvertInt(prodResult.SplitSeq)); // 이것이 스플릿인가 아닌가
                    sqlParameter.Add("WorkTimeMinute", txtWorkMinute.Text == string.Empty ? 0 : Convert.ToDouble(txtWorkMinute.Text.Replace(",", "")));
                    sqlParameter.Add("UpdateUserID", MainWindow.CurrentUser);

                    string[] result = DataStore.Instance.ExecuteProcedure_NewLog("xp_prd_uWkResultOne_WPF", sqlParameter, "U");

                    if (!result[0].Equals("success"))
                    {
                        //MessageBox.Show("실패 ㅠㅠ");
                        MessageBox.Show("저장실패 : " + result[1]);
                        flag = false;
                    }
                    else
                    {
                        //MessageBox.Show("성공 *^^*");
                        flag = true;
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

            return flag;
        }

        #region 유효성 검사 CheckData
        private bool CheckData()
        {
            bool flag = true;

            // 유효성 검사!!!!
            // 1. 작업 시작 날짜가 없다면? 생산 데이터를 고치는데, 작업 시작날짜가 있어야지.
            if (dtpWorkStartDate.SelectedDate == null)
            {
                MessageBox.Show("작업 시작 날짜를 선택해주세요.");
                return false;
            }

            // 시간이 맞지 않을때
            if (dtpWorkEndDate.SelectedDate != null)
            {
                // 1. 작업완료날짜가 시작날짜보다 적을 때
                if (dtpWorkEndDate.SelectedDate < dtpWorkStartDate.SelectedDate)
                {
                    MessageBox.Show("[작업 완료 일자]가 [작업 시작 일자]보다 작을 수 없습니다.");
                    return false;
                }

                // 2. 같은 날짜에 작업완료 시간이 적을 때
                if (dtpWorkEndDate.SelectedDate == dtpWorkStartDate.SelectedDate)
                {
                    if (ConvertInt(DateTimeFormat(txtEndTime.Text)) < ConvertInt(DateTimeFormat(txtStartTime.Text)))
                    {
                        MessageBox.Show("같은 날짜에 [작업완료 시간]이 [작업시작 시간]보다 작을 수 없습니다.");
                        return false;
                    }
                }

                // 3. 날짜 시간으로 계산
                DateTime SDate = new DateTime();
                DateTime EDate = new DateTime();

                DateTime.TryParse(dtpWorkStartDate.SelectedDate.Value.ToString("yyyy-MM-dd") + " " + getDateTime_Colon(DateTimeFormat(txtStartTime.Text.Trim())), out SDate);
                DateTime.TryParse(dtpWorkEndDate.SelectedDate.Value.ToString("yyyy-MM-dd") + " " + getDateTime_Colon(DateTimeFormat(txtEndTime.Text.Trim())), out EDate);

                if (SDate > EDate)
                {
                    MessageBox.Show("작업 시작 날짜 : " + SDate.ToString("yyyy-MM-dd HH:mm:ss") + "\r작업 완료 날짜 : " + EDate.ToString("yyyy-MM-dd HH:mm:ss") + "\r[작업 완료 날짜]가 [작업 시작 날짜]보다 작을 수 없습니다.");
                    return false;
                }
            }

            // 호기를 입력하도록
            if (cboMachine.SelectedValue == null)
            {
                MessageBox.Show("호기를 선택해주세요.");
                return false;
            }

            // 작업자를 입력하도록
            if (txtWorker.Tag == null)
            {
                MessageBox.Show("작업자를 입력해주세요.");
                return false;
            }


            return flag;
        }
        #endregion

        private string DateTimeFormat(string Time)
        {
            Time = Time.Trim().Replace(":", "");

            if (Time.Length <= 6)
            {
                // 8 하나만 입력했을 시 08 00 00 으로 저장 되도록 설정
                if (Time.Length == 1)
                {
                    Time = "0" + Time;
                }

                // 6글자 보다 작다면 0으로 채워줌
                for (int i = Time.Length; i < 6; i++)
                {
                    Time += "0";
                }
            }
            else
            {
                Time = Time.Substring(0, 6);
            }

            return Time;
        }

        private string getDateTime_Colon(string str)
        {
            if (str.Length == 6)
            {
                string h = str.Substring(0, 2);
                string m = str.Substring(2, 2);
                string d = str.Substring(4, 2);

                str = h + ":" + m + ":" + d;
            }

            return str;
        }

        #endregion // 주요 메서드 - 작업 정보 수정

        #region 주요 메서드 - 삭제 DeleteData

        private bool DeleteData(string strJobID, string SplitSeq)
        {
            bool flag = false;

            try
            {
                Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                sqlParameter.Clear();
                sqlParameter.Add("JobID", ConvertDouble(strJobID));
                sqlParameter.Add("SplitSeq", ConvertInt(SplitSeq));
                sqlParameter.Add("CreateUserID", MainWindow.CurrentUser);
                sqlParameter.Add("sRtnMsg", "");

                string[] result = DataStore.Instance.ExecuteProcedure_NewLog("xp_prdWork_dWkResult", sqlParameter, "D");

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

        #endregion // 주요 메서드 - 삭제 DeleteData

        // 메인 그리드 선택 시
        private void dgdResult_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            prodResult = dgdResult.SelectedItem as Win_Prd_ProdResult_U_CodeView;

            if (prodResult != null)
            {
                if (dgdResultChild != null) { dgdResultChild.Items.Clear(); }
                if (dgdResultDefect != null) { dgdResultDefect.Items.Clear(); }

                if (Lib.Instance.IsIntOrAnother(prodResult.JobID))
                {
                    int jid = int.Parse(prodResult.JobID);
                    FillGridChild(jid);
                    FillGridDefect(jid);
                }

                // 아래에 다른거 선택했을때 이전 MachineID 를 미리 백업 시켜놓고, 공정 콤보박스를 다시 세팅하면서 없어진 MachineID 복원 시키기
                string MachineID = prodResult.MachineID;
                this.DataContext = prodResult;

                // 설비 콤보박스 재 새팅 → 다른거 선택 시에 이전 행의 MachineID 값이 사라짐.
                ObservableCollection<CodeView> ovcMachine = ComboBoxUtil.Instance.GetMachine(prodResult.ProcessID == null ? "" : prodResult.ProcessID);
                this.cboMachine.ItemsSource = ovcMachine;
                this.cboMachine.DisplayMemberPath = "code_name";
                this.cboMachine.SelectedValuePath = "code_id";

                prodResult.MachineID = MachineID;
            }
        }

        #region 관리 번호 or OrderNo

        private void rbnOrderNo_Click(object sender, RoutedEventArgs e)
        {
            RbClick();
        }

        private void rbnOrderID_Click(object sender, RoutedEventArgs e)
        {
            RbClick();
        }

        private void RbClick()
        {
            if (rbnOrderID.IsChecked == true)
            {
                tbkOrder.Text = "관리번호";
                dgtOrderID.Visibility = Visibility.Visible;
                dgtOrderNo.Visibility = Visibility.Hidden;
            }
            else if (rbnOrderNo.IsChecked == true)
            {
                tbkOrder.Text = "Order No.";
                dgtOrderID.Visibility = Visibility.Hidden;
                dgtOrderNo.Visibility = Visibility.Visible;
            }
        }

        #endregion

        #region 작업시간 자동계산
        private void txtStartTime_TextChanged(object sender, TextChangedEventArgs e)
        {
            calcurateTime();
        }

        private void txtEndTime_TextChanged(object sender, TextChangedEventArgs e)
        {
            calcurateTime();
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

        private void txtQty_PreviewTextInput(object sender, TextCompositionEventArgs e)
        {
            Lib.Instance.CheckIsNumeric((TextBox)sender, e);
        }

        private void txtCT_PreviewTextInput(object sender, TextCompositionEventArgs e)
        {
            Lib.Instance.CheckIsNumeric((TextBox)sender, e);
        }

        private void txtStartTime_PreviewTextInput(object sender, TextCompositionEventArgs e)
        {
            Lib.Instance.CheckIsNumeric((TextBox)sender, e);
        }

        private void txtEndTime_PreviewTextInput(object sender, TextCompositionEventArgs e)
        {
            Lib.Instance.CheckIsNumeric((TextBox)sender, e);
        }

        private void txtProdScanTime_PreviewTextInput(object sender, TextCompositionEventArgs e)
        {
            Lib.Instance.CheckIsNumeric((TextBox)sender, e);
        }


        // 검색조건 - 작업자 이름
        private void lblWorkerName_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            if (chkWorkerName.IsChecked == true)
            {
                chkWorkerName.IsChecked = false;
                txtWorkerNameSearch.IsEnabled = false;
            }
            else
            {
                chkWorkerName.IsChecked = true;
                txtWorkerNameSearch.IsEnabled = true;
            }
        }

        private void chkWorkerName_Click(object sender, RoutedEventArgs e)
        {
            if (chkWorkerName.IsChecked == true)
            {
                txtWorkerNameSearch.IsEnabled = true;
            }
            else
            {
                txtWorkerNameSearch.IsEnabled = false;
            }
        }

        private void txtWorkerNameSearch_KeyDown(object sender, KeyEventArgs e)
        {
            if (txtWorkerNameSearch.IsEnabled == true)
            {
                if (e.Key == Key.Enter)
                {
                    rowNum = 0;
                    using (Loading lw = new Loading(re_Search))
                    {
                        lw.ShowDialog();
                    }
                }
            }
        }

        private void txtArticleNoSearch_KeyDown(object sender, KeyEventArgs e)
        {
            //if (txtArticleNoSearch.IsEnabled == true
            //    && e.Key == Key.Enter)
            //{
            //    rowNum = 0;
            //    using (Loading lw = new Loading(re_Search))
            //    {
            //        lw.ShowDialog();
            //    }
            //}

            if(e.Key == Key.Enter)
            {
                MainWindow.pf.ReturnCode(txtArticleNoSearch, 76, txtArticleNoSearch.Text);
            }
        }

        private void btnBuyerArticleNoSearch_Click(object sender, RoutedEventArgs e)
        {
            MainWindow.pf.ReturnCode(txtArticleNoSearch, 76, txtArticleNoSearch.Text);
        }

        // 작업자도 수정 되도록
        private void txtWorker_KeyDown(object sender, KeyEventArgs e)
        {
            if (txtWorker.IsEnabled == true
                && e.Key == Key.Enter)
            {
                MainWindow.pf.ReturnCode(txtWorker, (int)Defind_CodeFind.DCF_PERSON, "");
            }
        }

        private void btnPfWorker_Click(object sender, RoutedEventArgs e)
        {
            MainWindow.pf.ReturnCode(txtWorker, (int)Defind_CodeFind.DCF_PERSON, "");
        }

        private void btnRefresh_Click(object sender, RoutedEventArgs e) //2021-11-02 새로고침 클릭 시 팝업창 띄우기
        {
                      
            try
            {
                ReFresh ReFresh = null;
                ReFresh = new ReFresh();                
                ReFresh.ShowDialog();

            }
            catch(Exception Exception)
            {
                MessageBox.Show(Exception.Message);
                return;
            }

            //Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
            //sqlParameter.Clear();
            //string[] result = DataStore.Instance.ExecuteProcedure("xp_Batch_iWorkTime", sqlParameter, false);
            //if (!result[0].Equals("success"))
            //{
            //    MessageBox.Show("이상발생, 관리자에게 문의하세요.");
            //    return;
            //}
        }

        #region 입력창 이동 이벤트

        private void cboDayOrNight_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                txtCT.Focus();
            }
        }

        private void cboDayOrNight_DropDownClosed(object sender, EventArgs e)
        {
            txtCT.Focus();
        }

        private void txtCT_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                dtpProdDate.Focus();
            }
        }

        private void dtpProdDate_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                txtProdScanTime.Focus();
            }
        }

        private void dtpProdDate_PreviewKeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                dtpProdDate.IsDropDownOpen = true;
            }
        }

        private void dtpProdDate_CalendarClosed(object sender, RoutedEventArgs e)
        {
            txtProdScanTime.Focus();
        }

        private void txtProdScanTime_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                cboMachine.Focus();
            }
        }

        private void cboMachine_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                dtpWorkStartDate.Focus();
            }
        }

        private void cboMachine_DropDownClosed(object sender, EventArgs e)
        {
            try
            {
                dtpWorkStartDate.Focus();
            }
            catch (Exception ee)
            {
                MessageBox.Show("예외처리 - " + ee.ToString());
            }
        }

        private void dtpWorkStartDate_KeyDown(object sender, EventArgs e)
        {
            try
            {
                txtStartTime.Focus();
            }
            catch (Exception ee)
            {
                MessageBox.Show("예외처리 - " + ee.ToString());
            }
        }

        private void dtpWorkStartDate_PreviewKeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                dtpWorkStartDate.IsDropDownOpen = true;
            }
        }

        private void dtpWorkStartDate_CalendarClosed(object sender, RoutedEventArgs e)
        {
            txtStartTime.Focus();
        }

        private void txtStartTime_KeyDown(object sender, EventArgs e)
        {
            try
            {
                //dtpWorkEndDate.Focus();
            }
            catch (Exception ee)
            {
                MessageBox.Show("예외처리 - " + ee.ToString());
            }
        }

        private void dtpWorkEndDate_KeyDown(object sender, EventArgs e)
        {
            try
            {
                txtEndTime.Focus();
            }
            catch (Exception ee)
            {
                MessageBox.Show("예외처리 - " + ee.ToString());
            }
        }

        private void dtpWorkEndDate_PreviewKeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                dtpWorkEndDate.IsDropDownOpen = true;
            }
        }
  
        private void dtpWorkEndDate_CalendarClosed(object sender, RoutedEventArgs e)
        {
            txtEndTime.Focus();
        }

        private void txtEndTime_KeyDown(object sender, EventArgs e)
        {
            try
            {
                //txtWorkMinute.Focus();
            }
            catch (Exception ee)
            {
                MessageBox.Show("예외처리 - " + ee.ToString());
            }
        }

        private void txtWorkMinute_KeyDown(object sender, EventArgs e)
        {
            try
            {
                txtWorker.Focus();
            }
            catch (Exception ee)
            {
                MessageBox.Show("예외처리 - " + ee.ToString());
            }
        }


        #endregion

        private void dtpWorkStartDate_SelectedDateChanged(object sender, SelectionChangedEventArgs e)
        {
            calcurateTime();
        }

        private void dtpWorkEndDate_SelectedDateChanged(object sender, SelectionChangedEventArgs e)
        {
            calcurateTime();
        }

        public void calcurateTime()
        {
            try
            {
                if (txtStartTime.Text != string.Empty && txtEndTime.Text != string.Empty && txtStartTime.Text.Length == 8)
                {
                    string SDate = dtpWorkStartDate.SelectedDate.Value.ToString("yyyy-MM-dd") + " " + txtStartTime.Text;

                    if (dtpWorkEndDate.SelectedDate.Value != null)
                    {
                        string EDate = dtpWorkEndDate.SelectedDate.Value.ToString("yyyy-MM-dd") + " " + txtEndTime.Text;
                        DateTime StartDate = Convert.ToDateTime(SDate);
                        DateTime EndDate = Convert.ToDateTime(EDate);

                        TimeSpan dateDiff = EndDate - StartDate;

                        double totalMinutes = dateDiff.TotalMinutes;

                        txtWorkMinute.Text = Math.Round(totalMinutes).ToString();
                    }
                    else
                    {
                        MessageBox.Show("작업 종료 시간이 없습니다");
                    }

                }
            }
            catch (Exception ex)
            {

            }
        }
    }

    class Win_Prd_ProdResult_U_CodeView : BaseView
    {
        public override string ToString()
        {
            return (this.ReportAllProperties());
        }

        public int cls { get; set; }
        public string WorkDate { get; set; }
        public string ProcessID { get; set; }
        public string Process { get; set; }
        public string OrderID { get; set; }

        public string OrderNo { get; set; }
        public string AcptDate { get; set; }
        public double OrderQty { get; set; }
        public string MachineID { get; set; }
        public string InstDate { get; set; }

        public double InstQty { get; set; }
        public double WorkQty { get; set; }
        public string WorkPersonID { get; set; }
        public string ScanTime { get; set; }
        public string BuyerModelID { get; set; }

        public string BuyerModel { get; set; }
        public string CustomID { get; set; }
        public string KCustom { get; set; }
        public string Worker { get; set; }
        public string Article { get; set; }

        public string LabelID { get; set; }
        public string JobGbn { get; set; }
        public string JobGbnname { get; set; }
        public string WorkStartDate { get; set; }
        public string WorkStartTime { get; set; }

        public string WorkEndDate { get; set; }
        public string WorkEndTime { get; set; }
        public string WorkHour { get; set; }
        public string WorkMinute { get; set; }
        public string JobID { get; set; }

        public string Articleid { get; set; }
        public string WorkCnt { get; set; }
        public string NoReworkCode { get; set; }
        public string NoReworkName { get; set; }
        public string FourMID { get; set; }

        public string FourMSubject { get; set; }
        public int Num { get; set; }

        public string WorkDate_CV { get; set; }
        public string AcptDate_CV { get; set; }
        public string InstDate_CV { get; set; }
        public string WorkStartDate_CV { get; set; }
        public string WorkEndDate_CV { get; set; }

        public string DayOrNightID { get; set; } // 작업조 주간(01) / 야간(02) 
        public string CycleTime { get; set; } // 사이클 타임 
        public string BuyerArticleNo { get; set; } // 품번

        public string WorkTime { get; set; }

        public string ScanDate { get; set; }
        public string ScanDate_CV { get; set; }
        public bool IsSplitYN { get; set; }
        public string SplitSeq { get; set; }
        public string gbn { get; set; }
        public double WorkTimeMinute { get; set; }
    }

    class ProdResult_Count : BaseView
    {
        public override string ToString()
        {
            return (this.ReportAllProperties());
        }

        public string ctGubun { get; set; }
        public string ctCount { get; set; }
        public string ctQty { get; set; }
        public string ctWorkHour { get; set; }
    }

    class ProdResult_Child : BaseView
    {
        public override string ToString()
        {
            return (this.ReportAllProperties());
        }

        public string ChildLabelID { get; set; }
        public string ChildArticleID { get; set; }
        public string Article { get; set; }
    }

    class ProdResult_Defect : BaseView
    {
        public override string ToString()
        {
            return (this.ReportAllProperties());
        }

        public string DefectID { get; set; }
        public string KDefect { get; set; }
        public string DefectCount { get; set; }
    }
}

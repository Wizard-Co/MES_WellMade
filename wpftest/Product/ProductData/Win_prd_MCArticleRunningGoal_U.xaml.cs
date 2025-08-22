using MahApps.Metro.Controls;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Data;
using System.Drawing;
using System.IO;
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
using static System.Net.WebRequestMethods;

namespace WizMes_WellMade
{
    /// <summary>
    /// Win_prd_MCRunningGoal_U.xaml에 대한 상호 작용 논리
    /// </summary>
    public partial class Win_prd_MCArticleRunningGoal_U : UserControl
    {
        string stDate = string.Empty;
        string stTime = string.Empty;

        string strFlag = string.Empty;
        Lib lib = new Lib();
        int rowNum = 0;
        PlusFinder pf = new PlusFinder();
        string strMachine = string.Empty;
        string strProcess = string.Empty;

        Win_prd_MCRunningGoal_U_CodeView WinMachine = new Win_prd_MCRunningGoal_U_CodeView();
        Win_prd_MCRunningGoal_U_Sub_CodeView WinMachineGoal = new Win_prd_MCRunningGoal_U_Sub_CodeView();

        List<Win_prd_MCRunningGoal_U_Sub_CodeView> lstMG = new List<Win_prd_MCRunningGoal_U_Sub_CodeView>();

        // 수정 후 찾아가기 위한 변수!
        string ProcessIDS = "";
        string MachineIDS = "";

        //FTP 활용모음
        string strImagePath = string.Empty;
        string strFullPath = string.Empty;
        string strDelFileName = string.Empty;
        List<string[]> listFtpFile = new List<string[]>();
        string FTP_ADDRESS = "ftp://" + LoadINI.FileSvr + ":" + LoadINI.FTPPort + LoadINI.FtpImagePath + "/";
        string ForderName = "ProdGoal";
        private FTP_EX _ftp = null;
        private const string FTP_ID = "wizuser";
        private const string FTP_PASS = "wiz9999";

        public Win_prd_MCArticleRunningGoal_U()
        {
            InitializeComponent();
        }

        private void UserContrl_Loaded(object sender, RoutedEventArgs e)
        {
            stDate = DateTime.Now.ToString("yyyyMMdd");
            stTime = DateTime.Now.ToString("HHmm");

            DataStore.Instance.InsertLogByFormS(this.GetType().Name, stDate, stTime, "S");
            dtpStartDate.SelectedDate = DateTime.Today;
            dtpEndDate.SelectedDate = DateTime.Today;
            SetComboBox();
        }
    

        #region 콤보박스
        private void SetComboBox()
        {
            ObservableCollection<CodeView> ovcProcess = ComboBoxUtil.Instance.GetWorkProcess_only(0, "");
            this.cboProcess.ItemsSource = ovcProcess;
            this.cboProcess.DisplayMemberPath = "code_name";
            this.cboProcess.SelectedValuePath = "code_id";

            ObservableCollection<CodeView> ovcMachine = SetMachine("");
            this.cboMachine.ItemsSource = ovcMachine;
            this.cboMachine.DisplayMemberPath = "code_name";
            this.cboMachine.SelectedValuePath = "code_id";

            // 수동 자동 세팅
            ObservableCollection<CodeView> ovcAutoPassive = ComboBoxUtil.Instance.GetCMCode_SetComboBox("AUTOPASSIVE", "");
            cboAutoPassive.ItemsSource = ovcAutoPassive;
            cboAutoPassive.DisplayMemberPath = "code_name";
            cboAutoPassive.SelectedValuePath = "code_id";
        }

        private ObservableCollection<CodeView> SetMachine(string value)
        {
            ObservableCollection<CodeView> retunCollection = new ObservableCollection<CodeView>();

            Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
            sqlParameter.Add("@sProcessID", value);

            DataSet ds = DataStore.Instance.ProcedureToDataSet("xp_Process_sMachine", sqlParameter, false);

            if (ds != null && ds.Tables.Count > 0)
            {
                DataTable dt = ds.Tables[0];

                if (dt.Rows.Count > 0)
                { 
                    DataRowCollection drc = dt.Rows;

                    foreach (DataRow item in drc)
                    {
                        CodeView mCodeView = new CodeView()
                        {
                            code_id = item["ProcessID"].ToString().Trim() + "/" + item["MachineID"].ToString().Trim(),
                            //code_id = item["MachineID"].ToString().Trim(),
                            //code_name = item["Process"].ToString().Trim() + "  " + item["Machine"].ToString().Trim()
                            code_name = item["MachineNo"].ToString().Trim()
                        };
                        retunCollection.Add(mCodeView);
                    }
                }
            }

            return retunCollection;
        }
        #endregion

        #region 상단 검색 조건
        //기간
        private void lblTermSrh_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            if (chkTermSrh.IsChecked == true) { chkTermSrh.IsChecked = false; }
            else { chkTermSrh.IsChecked = true; }

        }
        //기간 체크박스 체크
        private void chkTermSrh_Checked(object sender, RoutedEventArgs e)
        {
            if (dtpStartDate != null && dtpEndDate != null)
            {
                dtpStartDate.IsEnabled = true;
                dtpEndDate.IsEnabled = true;
            }
        }

        //기간 체크박스 체크해제
        private void chkTermSrh_Unchecked(object sender, RoutedEventArgs e)
        {
            dtpStartDate.IsEnabled = false;
            dtpEndDate.IsEnabled = false;
        }
        #endregion

        #region 상단 버튼과 버튼 상태
        /// <summary>
        /// 수정,추가 저장 후
        /// </summary>
        private void CanBtnControl()
        {
            Lib.Instance.UiButtonEnableChange_IUControl(this);
            grdOne.IsEnabled = false;
            grdTwo.IsEnabled = false;
            grdThree.IsEnabled = false;

            btnAddSub.IsEnabled = false;
            btnDelSub.IsEnabled = false;

            btnSelectSave.Visibility = Visibility.Hidden;
        }

        /// <summary>
        /// 수정,추가 진행 중
        /// </summary>
        private void CantBtnControl()
        {
            Lib.Instance.UiButtonEnableChange_SCControl(this);
            grdOne.IsEnabled = true;
            grdTwo.IsEnabled = true;
            grdThree.IsEnabled = true;

            btnAddSub.IsEnabled = true;
            btnDelSub.IsEnabled = true;

            btnSelectSave.Visibility = Visibility.Visible;
        }

        //복사
        private void btnCopy_Click(object sender, RoutedEventArgs e)
        {
            Win_comMCRunningTarget_PopUP TargetCopy = new Win_comMCRunningTarget_PopUP();
            TargetCopy.ShowDialog();

            if (TargetCopy.DialogResult == true)
            {
                rowNum = 0;
                re_Search(rowNum);
            }
        }

        //추가
        private void btnAdd_Click(object sender, RoutedEventArgs e)
        {
            CantBtnControl();
            strFlag = "I";
            //dgdMain.IsEnabled = false;
            dgdMain.IsHitTestVisible = false;

            if (dgdSub.Items.Count > 0)
            {
                dgdSub.Items.Clear();
            }

            lblMsg.Visibility = Visibility.Visible;
            tbkMsg.Text = "자료 입력 중";
            rowNum = dgdMain.SelectedIndex;
            this.DataContext = null;
            txtYYYY.Text = DateTime.Today.ToString("yyyy");

            ObservableCollection<CodeView> ovcMachine = SetMachine("");
            this.cboMachine.ItemsSource = ovcMachine;
            this.cboMachine.DisplayMemberPath = "code_name";
            this.cboMachine.SelectedValuePath = "code_id";
        }

        //수정
        private void btnUpdate_Click(object sender, RoutedEventArgs e)
        {
            WinMachine = dgdMain.SelectedItem as Win_prd_MCRunningGoal_U_CodeView;

            if (WinMachine != null)
            {

                txtArticle.Text = "";
                FillGridSub(WinMachine.YYYY, WinMachine.ProcessID, WinMachine.MachineID);

                rowNum = dgdMain.SelectedIndex;
                //dgdMain.IsEnabled = false;
                dgdMain.IsHitTestVisible = false;
                tbkMsg.Text = "자료 수정 중";
                lblMsg.Visibility = Visibility.Visible;
                CantBtnControl();
                strFlag = "U";

                // 수정 저장 후 인덱스 찾아가기
                ProcessIDS = WinMachine.ProcessID;
                MachineIDS = WinMachine.MachineID;
            }
        }

        //삭제ㅇ
        private void btnDelete_Click(object sender, RoutedEventArgs e)
        {
            WinMachine = dgdMain.SelectedItem as Win_prd_MCRunningGoal_U_CodeView;

            if (WinMachineGoal == null)
            {
                MessageBox.Show("삭제할 데이터가 지정되지 않았습니다. 삭제데이터를 지정하고 눌러주세요");
            }
            else
            {
                if (MessageBox.Show("선택하신 항목을 삭제하시겠습니까?", "삭제 전 확인", MessageBoxButton.YesNo) == MessageBoxResult.Yes)
                {
                    if (dgdMain.Items.Count > 0 && dgdMain.SelectedItem != null)
                    {
                        rowNum = dgdMain.SelectedIndex;
                    }

                    if (DeleteData(WinMachine.YYYY, WinMachine.ProcessID, WinMachine.MachineID))
                    {
                        rowNum -= 1;
                        re_Search(rowNum);
                    }
                }
            }
        }

        //닫기
        private void btnClose_Click(object sender, RoutedEventArgs e)
        {
            DataStore.Instance.InsertLogByFormS(this.GetType().Name, stDate, stTime, "E");
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

        //검색
        private void btnSearch_Click(object sender, RoutedEventArgs e)
        {
            //검색버튼 비활성화
            btnSearch.IsEnabled = false;

            Dispatcher.BeginInvoke(new Action(() =>

            {
                //로직
                using (Loading lw = new Loading(beSearch))
                {
                    lw.ShowDialog();
                }

            }), System.Windows.Threading.DispatcherPriority.Background);



            Dispatcher.BeginInvoke(new Action(() =>

            {
                btnSearch.IsEnabled = true;

            }), System.Windows.Threading.DispatcherPriority.Background);
                        
        }

        private void beSearch()
        {
            rowNum = 0;
            re_Search(rowNum);
        }

        //저장
        private void btnSave_Click(object sender, RoutedEventArgs e)
        {
            using (Loading lw = new Loading(beSave))
            {
                lw.ShowDialog();
            }
        }

        private void beSave()
        {
            if (chkProcessAll.IsChecked == true)
            {
                if (MessageBox.Show("공정 전체로 추가하는 경우, 해당 공정의 이전 데이터는 사라집니다.\r이대로 진행 하시겠습니까??", "저장 전 확인", MessageBoxButton.YesNo) == MessageBoxResult.Yes)
                {
                    if (SaveData())
                    {
                        CanBtnControl();
                        lblMsg.Visibility = Visibility.Hidden;
                        rowNum = 0;
                        //dgdMain.IsEnabled = true;
                        dgdMain.IsHitTestVisible = true;

                        if (dgdSub.Items.Count > 0)
                        {
                            dgdSub.Items.Clear();
                        }

                        strProcess = string.Empty;
                        strMachine = string.Empty;

                        chkProcessAll.IsChecked = false;

                        SetComboBox();
                        re_Search(rowNum);
                    }
                }
            }
            else
            {
                if (SaveData())
                {
                    CanBtnControl();
                    lblMsg.Visibility = Visibility.Hidden;
                    rowNum = 0;
                    //dgdMain.IsEnabled = true;
                    dgdMain.IsHitTestVisible = true;

                    if (dgdSub.Items.Count > 0)
                    {
                        dgdSub.Items.Clear();
                    }

                    strProcess = string.Empty;
                    strMachine = string.Empty;

                    chkProcessAll.IsChecked = false;

                    SetComboBox();
                    re_Search(rowNum);
                }
            }
        }

        //취소
        private void btnCancel_Click(object sender, RoutedEventArgs e)
        {
            using (Loading lw = new Loading(beCancel))
            {
                lw.ShowDialog();
            }
        }

        private void beCancel()
        {
            CanBtnControl();
            strFlag = string.Empty;
            //dgdMain.IsEnabled = true;
            dgdMain.IsHitTestVisible = true;

            if (dgdSub.Items.Count > 0)
            {
                dgdSub.Items.Clear();
            }

            strProcess = string.Empty;
            strMachine = string.Empty;

            chkProcessAll.IsChecked = false;

            SetComboBox();
            re_Search(rowNum);
        }

        //엑셀
        private void btnExcel_Click(object sender, RoutedEventArgs e)
        {
            DataTable dt = null;
            string Name = string.Empty;

            string[] lst = new string[4];
            lst[0] = "설비";
            lst[1] = "설비 가동률 목표";
            lst[2] = dgdMain.Name;
            lst[3] = dgdSub.Name;

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
                    Lib.Instance.GenerateExcel(dt, Name);
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

        // 호기 선택해서 저장하기
        private void btnSelectSave_Click(object sender, RoutedEventArgs e)
        {
            if (CheckData())
            {
                Win_pop_CycleTime CT = new Win_pop_CycleTime();

                CT.ReceiveData = new Win_pop_CycleTime_CodeView()
                {
                    ProcessID = cboProcess.SelectedValue.ToString(),
                    MachineID = cboMachine.SelectedValue.ToString().Split('/')[1],
                    AutoPassive = cboAutoPassive.SelectedValue.ToString(),
                    YYYY = txtYYYY.Text.Trim(),
                };

                for (int i = 0; i < dgdSub.Items.Count; i++)
                {
                    var goalSub = dgdSub.Items[i] as Win_prd_MCRunningGoal_U_Sub_CodeView;
                    if (goalSub != null)
                    {
                        CT.lstSub.Add(goalSub);
                    }
                }

                CT.ShowDialog();

                if (CT.DialogResult == true)
                {
                    CanBtnControl();
                    lblMsg.Visibility = Visibility.Hidden;
                    rowNum = 0;
                    //dgdMain.IsEnabled = true;
                    dgdMain.IsHitTestVisible = true;

                    if (dgdSub.Items.Count > 0)
                    {
                        dgdSub.Items.Clear();
                    }

                    strProcess = string.Empty;
                    strMachine = string.Empty;

                    chkProcessAll.IsChecked = false;

                    SetComboBox();
                    re_Search(rowNum);
                }
            }
        }

        #endregion

        #region 우측 레이아웃 검색 조건
        //년도의 ▲ 버튼
        private void btnNextYear_Click(object sender, RoutedEventArgs e)
        {
            if (Lib.Instance.IsNumOrAnother(txtYYYY.Text))
            {
                txtYYYY.Text = (int.Parse(txtYYYY.Text) + 1).ToString();
            }
        }

        //년도의 ▼ 버튼
        private void btnPreYear_Click(object sender, RoutedEventArgs e)
        {
            if (Lib.Instance.IsNumOrAnother(txtYYYY.Text))
            {
                txtYYYY.Text = (int.Parse(txtYYYY.Text) - 1).ToString();
            }
        }

        //공정선택
        private void cboProcess_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (lblMsg.Visibility == Visibility.Visible)
            {
                if (cboProcess.SelectedValue != null 
                    && !cboProcess.SelectedValue.ToString().Equals(string.Empty))
                {
                    strProcess = cboProcess.SelectedValue.ToString();
                    this.cboMachine.ItemsSource = null;

                    ObservableCollection<CodeView> ovcMachine = SetMachine(cboProcess.SelectedValue.ToString());
                    this.cboMachine.ItemsSource = ovcMachine;
                    this.cboMachine.DisplayMemberPath = "code_name";
                    this.cboMachine.SelectedValuePath = "code_id";
                }
            }
        }

        #endregion

        #region 재검색
        /// <summary>
        /// 재검색(수정,삭제,추가 저장후에 자동 재검색)
        /// </summary>
        /// <param name="selectedIndex"></param>
        private void re_Search(int selectedIndex)
        {
            FillGrid();

            if (dgdMain.Items.Count > 0)
            {
                if (strFlag.Trim().Equals("U")
                    && !ProcessIDS.Trim().Equals("")
                    && !MachineIDS.Trim().Equals(""))
                {
                    bool flag = false;
                    for(int i = 0; i < dgdMain.Items.Count; i++)
                    {
                        var Main = dgdMain.Items[i] as Win_prd_MCRunningGoal_U_CodeView;
                        if (Main != null
                            && Main.ProcessID != null
                            && Main.MachineID != null)
                        {
                            if (Main.ProcessID.Trim().Equals(ProcessIDS)
                                && Main.MachineID.Trim().Equals(MachineIDS))
                            {
                                dgdMain.SelectedIndex = i;
                                flag = true;
                                break;
                            }
                        }
                    }

                    if (flag == false)
                    {
                        dgdMain.SelectedIndex = 0;
                    }
                }
                else
                {
                    dgdMain.SelectedIndex = selectedIndex;
                }
            }
            else
            {
                MessageBox.Show("조회된 데이터가 없습니다.");
                return;
            }
        }

        #endregion

        #region 조회 main
        /// <summary>
        /// 실조회
        /// </summary>
        private void FillGrid()
        {
            if (dgdMain.Items.Count > 0)
            {
                dgdMain.Items.Clear();
            }

            try
            {
                DataSet ds = null;
                Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                sqlParameter.Clear();
                sqlParameter.Add("ChkDate", chkTermSrh.IsChecked == true ? 1 : 0);
                sqlParameter.Add("sYYYY", chkTermSrh.IsChecked == true ? dtpStartDate.SelectedDate.Value.ToString("yyyy") : "");
                sqlParameter.Add("eYYYY", chkTermSrh.IsChecked == true ? dtpEndDate.SelectedDate.Value.ToString("yyyy") : "");

                ds = DataStore.Instance.ProcedureToDataSet_LogWrite("xp_MachineGoal_sMachineGoalList", sqlParameter, true, "R");

                if (ds != null && ds.Tables.Count > 0)
                {
                    DataTable dt = ds.Tables[0];
                    int i = 0;

                    if (dt.Rows.Count == 0)
                    {
                        //MessageBox.Show("조회된 데이터가 없습니다.");
                        this.DataContext = null;
                        if (dgdSub.Items.Count > 0)
                        {
                            dgdSub.Items.Clear();
                        }
                    }
                    else
                    {
                        DataRowCollection drc = dt.Rows;

                        foreach (DataRow dr in drc)
                        {
                            i++;
                            var WinMain = new Win_prd_MCRunningGoal_U_CodeView()
                            {
                                Num = i,
                                YYYY = dr["YYYY"].ToString(),
                                ProcessID = dr["ProcessID"].ToString(),
                                Process = dr["Process"].ToString(),
                                MachineID = dr["MachineID"].ToString(),
                                MachineNo = dr["MachineNo"].ToString(),
                                AutoPassive = dr["AutoPassive"].ToString(),
                                AutoPassiveName = dr["AutoPassiveName"].ToString(),
                            };

                            WinMain.ProcessMachineID = WinMain.ProcessID + "/" + WinMain.MachineID;

                            dgdMain.Items.Add(WinMain);
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

        #region 조회 sub
        private void FillGridSub(string strYYYY, string strProcessID, string strMachineID)
        {
            lstMG.Clear();
            
            if (dgdSub.Items.Count > 0)
            {
                dgdSub.Items.Clear();
            }

            try
            {
                DataSet ds = null;
                Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                sqlParameter.Clear();
                sqlParameter.Add("sYYYY", strYYYY);
                sqlParameter.Add("sProcessID", strProcessID);
                sqlParameter.Add("sMachineID", strMachineID);
                sqlParameter.Add("sArticleNo", txtArticle.Text);
                ds = DataStore.Instance.ProcedureToDataSet("xp_MachineGoal_sMachineGoal", sqlParameter, false);

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
                            var WinSub = new Win_prd_MCRunningGoal_U_Sub_CodeView()
                            {
                                Num = i,
                                ArticleID = dr["ArticleID"].ToString(),
                                Article = dr["Article"].ToString(),
                                BuyerArticleNo = dr["BuyerArticleNo"].ToString(),
                                CT = stringFormatN1(dr["CT"]),
                                AutoPassive = dr["AutoPassive"].ToString(),
                                AutoPassiveName = dr["AutoPassiveName"].ToString(),
                            };

                            dgdSub.Items.Add(WinSub);
                            lstMG.Add(WinSub);
                        }

                        tblCnt.Text = "▶ 검색 건수 : " + i + "건";
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

        #region 데이터그리드 SelectionChanged
        private void dgdMain_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            WinMachine = dgdMain.SelectedItem as Win_prd_MCRunningGoal_U_CodeView;

            if (WinMachine != null)
            {
                //string MachineID = WinMachine.MachineID;
                //ObservableCollection<CodeView> ovcMachine = SetMachine(WinMachine.ProcessID);
                //this.cboMachine.ItemsSource = ovcMachine;
                //this.cboMachine.DisplayMemberPath = "code_name";
                //this.cboMachine.SelectedValuePath = "code_id";
                //WinMachine.MachineID = MachineID;

                this.DataContext = WinMachine;
                txtArticle.Text = "";
                FillGridSub(WinMachine.YYYY, WinMachine.ProcessID, WinMachine.MachineID);
                
            }
        }
        #endregion

        #region 삭제
        /// <summary>
        /// 실삭제
        /// </summary>
        /// <param name="WinMcRunning"></param>
        /// <returns></returns>
        private bool DeleteData(string strYYYY, string strProcessID, string strMachineID)
        {
            bool flag = false;

            Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
            sqlParameter.Clear();
            sqlParameter.Add("YYYY", strYYYY);
            sqlParameter.Add("ProcessID", strProcessID);
            sqlParameter.Add("MachineID", strMachineID);

            string[] result = DataStore.Instance.ExecuteProcedure_NewLog("xp_MachineGoal_dMachineGoalAll", sqlParameter, "D");
            DataStore.Instance.CloseConnection();

            if (result[0].Equals("success"))
            {
                //MessageBox.Show("성공 *^^*");
                flag = true;
            }

            return flag;
        }
        #endregion

        #region 저장
        /// <summary>
        /// 실저장
        /// </summary>
        /// <param name="strFlag"></param>
        /// <param name="strArticleID"></param>
        /// <returns></returns>
        private bool SaveData()
        {
            bool flag = false;
            List<Procedure> Prolist = new List<Procedure>();
            List<Dictionary<string, object>> ListParameter = new List<Dictionary<string, object>>();

            Dictionary<string, object> sqlParameter = new Dictionary<string, object>();

            try
            {
                if (CheckData())
                {
                    #region 추가 또는 수정

                    if (strFlag.Equals("I")
                        || strFlag.Equals("U"))
                    {
                        sqlParameter = new Dictionary<string, object>();
                        sqlParameter.Clear();

                        sqlParameter.Add("YYYY", txtYYYY.Text);
                        sqlParameter.Add("ProcessID", cboProcess.SelectedValue != null ? cboProcess.SelectedValue.ToString() : "");
                        sqlParameter.Add("MachineID", cboMachine.SelectedValue != null ? cboMachine.SelectedValue.ToString().Split('/')[1] : "");
                        sqlParameter.Add("nProcessAllYN", chkProcessAll.IsChecked == true ? 1 : 0);

                        Procedure pro0 = new Procedure();
                        pro0.Name = "xp_MachineGoal_dMachineGoalAll";
                        pro0.OutputUseYN = "N";
                        pro0.OutputName = "sArticleID";
                        pro0.OutputLength = "5";

                        Prolist.Add(pro0);
                        ListParameter.Add(sqlParameter);

                        for (int i = 0; i < dgdSub.Items.Count; i++)
                        {
                            var WinSub = dgdSub.Items[i] as Win_prd_MCRunningGoal_U_Sub_CodeView;

                            if (WinSub != null
                                && WinSub.ArticleID != null
                                && !WinSub.ArticleID.Trim().Equals(""))
                            {
                                sqlParameter = new Dictionary<string, object>();
                                sqlParameter.Clear();

                                sqlParameter.Add("YYYY", txtYYYY.Text);
                                sqlParameter.Add("ProcessID", cboProcess.SelectedValue != null ? cboProcess.SelectedValue.ToString() : "");
                                sqlParameter.Add("MachineID", cboMachine.SelectedValue != null ? cboMachine.SelectedValue.ToString().Split('/')[1] : "");
                                sqlParameter.Add("ArticleID", WinSub.ArticleID);
                                sqlParameter.Add("CT", ConvertDouble(WinSub.CT));
                                sqlParameter.Add("AutoPassive", cboAutoPassive.SelectedValue != null ? cboAutoPassive.SelectedValue.ToString() : "");
                                sqlParameter.Add("nProcessAllYN", chkProcessAll.IsChecked == true ? 1 : 0);
                            
                                sqlParameter.Add("CreateUserID", MainWindow.CurrentUser);

                                Procedure pro1 = new Procedure();
                                pro1.Name = "xp_MachineGoal_iMachineGoal";
                                pro1.OutputUseYN = "N";
                                pro1.OutputName = "sArticleID";
                                pro1.OutputLength = "5";

                                Prolist.Add(pro1);
                                ListParameter.Add(sqlParameter);

                               
                            }
                        }

                        string[] Confirm = new string[2];
                        Confirm = DataStore.Instance.ExecuteAllProcedureOutputNew_NewLog(Prolist, ListParameter,"C");
                        if (Confirm[0] != "success")
                        {
                            MessageBox.Show("[저장실패]\r\n" + Confirm[1].ToString());
                            flag = false;
                            //return false;
                        }
                        else
                        {
                            flag = true;
                        }
                    }

                    if (flag)
                    {
                        if (listFtpFile.Count > 0)
                        {
                            FTP_Save_File(listFtpFile, txtYYYY.Text.ToString());
                        }
                    } 

                    listFtpFile.Clear();

                    #endregion
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

        private void FTP_Save_File(List<string[]> listStrArrayFileInfo, string MakeFolderName)
        {
            _ftp = new FTP_EX(FTP_ADDRESS, FTP_ID, FTP_PASS);

            List<string[]> UpdateFilesInfo = new List<string[]>();
            string[] fileListSimple;
            string[] fileListDetail = null;
            fileListSimple = _ftp.directoryListSimple("", Encoding.Default);

            // 기존 폴더 확인작업.
            bool MakeFolder = false;
            MakeFolder = FolderInfoAndFlag(fileListSimple, MakeFolderName);


            if (MakeFolder == false)        // 같은 아이를 찾지 못한경우,
            {
                if (_ftp.createDirectory("ProdGoal") == false)
                {
                    MessageBox.Show("업로드를 위한 폴더를 생성할 수 없습니다.");
                    return;
                }

                //MIL 폴더에 InspectionID로 저장
                if (_ftp.createDirectory(MakeFolderName) == false)
                {
                    MessageBox.Show("업로드를 위한 폴더를 생성할 수 없습니다.");
                    return;
                }
            }
            else
            {
                fileListDetail = _ftp.directoryListSimple(MakeFolderName, Encoding.Default);
            }

            for (int i = 0; i < listStrArrayFileInfo.Count; i++)
            {
                bool flag = true;

                if (fileListDetail != null)
                {
                    foreach (string compare in fileListDetail)
                    {
                        if (compare.Equals(listStrArrayFileInfo[i][0]))
                        {
                            flag = false;
                            break;
                        }
                    }
                }

                if (flag)
                {
                    listStrArrayFileInfo[i][0] = MakeFolderName + "/" + listStrArrayFileInfo[i][0];
                    UpdateFilesInfo.Add(listStrArrayFileInfo[i]);
                }
            }

            if (!_ftp.UploadTempFilesToFTP(UpdateFilesInfo))
            {
                MessageBox.Show("파일업로드에 실패하였습니다.");
                return;
            }
        }

        /// <summary>
        /// 해당영역에 폴더가 있는지 확인
        /// </summary>
        bool FolderInfoAndFlag(string[] strFolderList, string FolderName)
        {
            bool flag = false;
            foreach (string FolderList in strFolderList)
            {
                if (FolderList == FolderName)
                {
                    flag = true;
                    break;
                }
            }
            return flag;
        }

        #endregion

        #region 입력 체크
        /// <summary>
        /// 입력 데이터 체크
        /// </summary>
        /// <returns></returns>
        private bool CheckData()
        {
            bool flag = true;

            if (txtYYYY.Text.Length <= 0 || txtYYYY.Text.Equals(string.Empty))
            {
                MessageBox.Show("년도가 입력되지 않았습니다.");
                flag = false;
                return flag;
            }

            // 공정을 선택해주세요.
            if (cboProcess.SelectedValue == null
                || cboProcess.SelectedValue.ToString().Trim().Equals(""))
            {
                MessageBox.Show("공정을 선택해주세요.");
                flag = false;
                return flag;
            }

            // 호기를 선택해주세요.
            if (chkProcessAll.IsChecked == false
                && (cboMachine.SelectedValue == null
                || cboMachine.SelectedValue.ToString().Trim().Equals("")))
            {
                MessageBox.Show("호기를 선택해주세요.");
                flag = false;
                return flag;
            }

            // 수동 / 자동을 선택해주세요.
            if (cboAutoPassive.SelectedValue == null
              || cboAutoPassive.SelectedValue.ToString().Trim().Equals(""))
            {
                MessageBox.Show("수동 / 자동을 선택해주세요.");
                flag = false;
                return flag;
            }

            if (dgdSub.Items.Count <= 0)
            {
                MessageBox.Show("품명별 CT가 입력되지 않았습니다.\r[추가]버튼을 통해서 데이터를 입력해주세요.");
                flag = false;
                return flag;
            }

            return flag;
        }
        #endregion

        #region 데이터그리드 키조작
        private void DataGridCell_KeyDown(object sender, KeyEventArgs e)
        {
            WinMachineGoal = dgdSub.CurrentItem as Win_prd_MCRunningGoal_U_Sub_CodeView;
            int rowCount = dgdSub.Items.IndexOf(dgdSub.CurrentItem);
            int colCountOne = dgdSub.Columns.IndexOf(dgdtpeBuyArticleNo);
            int colCountTwo = dgdSub.Columns.IndexOf(dgdtpeArticle);
            int colCountThree = dgdSub.Columns.IndexOf(dgdtpeGoalRunRate);
            int colCount = dgdSub.Columns.IndexOf(dgdSub.CurrentCell.Column);

            if (e.Key == Key.Enter)
            {
                e.Handled = true;
                (sender as DataGridCell).IsEditing = false;

                if (dgdSub.Items.Count - 1 >= rowCount && colCount == colCountOne)
                {
                    dgdSub.CurrentCell = new DataGridCellInfo(dgdSub.Items[rowCount], dgdSub.Columns[colCountTwo]);
                }
                else if (dgdSub.Items.Count - 1 >= rowCount && colCount == colCountTwo)
                {
                    dgdSub.CurrentCell = new DataGridCellInfo(dgdSub.Items[rowCount], dgdSub.Columns[colCountThree]);
                }
                else if (dgdSub.Items.Count - 1 >= rowCount && colCount == colCountThree)
                {
                    if (MessageBox.Show("선택한 행을 추가하시겠습니까?", "추가 전 확인", MessageBoxButton.YesNo) == MessageBoxResult.Yes)
                    {
                        AddDgdSub();
                    }
                }

            }
            else if (e.Key == Key.Delete)
            {
                var WinMachineGoal = dgdSub.SelectedItem as Win_prd_MCRunningGoal_U_Sub_CodeView;
                if (WinMachineGoal != null)
                {
                    if (MessageBox.Show("선택한 행을 삭제하시겠습니까?", "삭제 전 확인", MessageBoxButton.YesNo) == MessageBoxResult.Yes)
                    {
                        dgdSub.Items.Remove(WinMachineGoal);
                    }
                }
            }
        }

        //
        private void TextBoxFocusInDataGrid(object sender, KeyEventArgs e)
        {
            Lib.Instance.DataGridINControlFocus(sender, e);
        }

        //
        private void TextBoxFocusInDataGrid_MouseUp(object sender, MouseButtonEventArgs e)
        {
            Lib.Instance.DataGridINBothByMouseUP(sender, e);
        }

        //
        private void DataGridCell_GotFocus(object sender, RoutedEventArgs e)
        {
            if (lblMsg.Visibility == Visibility.Visible)
            {
                DataGridCell cell = sender as DataGridCell;

                if (cell.Column == dgdtpeBuyArticleNo
                    || cell.Column == dgdtpeArticle
                    || cell.Column == dgdtpeGoalRunRate)
                {
                    cell.IsEditing = true;
                }
            }
        }
        #endregion

        #region 데이터그리드 텍스트 변경

        private void TextBoxBuyArticleNo_KeyDown(object sender, KeyEventArgs e)
        {
            // 품번 76
            if (e.Key == Key.Enter)
            {
                TextBox txtSender = sender as TextBox;
                var MCSub = dgdSub.CurrentItem as Win_prd_MCRunningGoal_U_Sub_CodeView;

                if (MCSub != null)
                {
                    MainWindow.pf.ReturnCode(txtSender, 76, txtSender.Text);

                    if (txtSender.Tag != null)
                    {
                        for (int i = 0; i < dgdSub.Items.Count; i++)
                        {
                            var Compare = dgdSub.Items[i] as Win_prd_MCRunningGoal_U_Sub_CodeView;
                            if (Compare != null
                                && Compare.ArticleID != null)
                            {
                                if (txtSender.Tag != null
                                    && !txtSender.Tag.ToString().Equals("")
                                    && Compare.ArticleID.Trim().Equals(txtSender.Tag.ToString().Trim()))
                                {
                                    MessageBox.Show("이미 등록이 된 품목입니다.\r등록시 마지막에 등록한 데이터로 저장이 됩니다.");
                                    break;
                                }
                            }
                        }

                        MCSub.BuyerArticleNo = txtSender.Text;
                        MCSub.ArticleID = txtSender.Tag.ToString();
                        MCSub.Article = getArticleInfo(txtSender.Tag.ToString()).Article;
                    }
                }
            }
        }

        //
        private void TextBoxArticle_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                #region 품번 플러스파인더 주석처리
                //TextBox txtSender = sender as TextBox;
                //var MCSub = dgdSub.CurrentItem as Win_prd_MCRunningGoal_U_Sub_CodeView;

                //if (MCSub != null)
                //{
                //    MainWindow.pf.ReturnCode(txtSender, 76, txtSender.Text);

                //    if (txtSender.Tag != null)
                //    {
                //        for (int i = 0; i < dgdSub.Items.Count; i++)
                //        {
                //            var Compare = dgdSub.Items[i] as Win_prd_MCRunningGoal_U_Sub_CodeView;
                //            if (Compare != null
                //                && Compare.ArticleID != null)
                //            {
                //                if (Compare.ArticleID.Trim().Equals(txtSender.Tag.ToString().Trim()))
                //                {
                //                    MessageBox.Show("이미 등록이 된 품목입니다.\r등록시 마지막에 등록한 데이터로 저장이 됩니다.");
                //                    break;
                //                }
                //            }
                //        }

                //        MCSub.Article = txtSender.Text;
                //        MCSub.ArticleID = txtSender.Tag.ToString();
                //        MCSub.BuyerArticleNo = getArticleInfo(txtSender.Tag.ToString()).BuyerArticleNo;
                //    }
                //}
                #endregion
            }
        }

        //
        private void TextBox_PreviewTextInput(object sender, TextCompositionEventArgs e)
        {
            decimal result;
            if (!(Decimal.TryParse(e.Text, out result)) && (!e.Text.Equals(".")))
            {
                e.Handled = true;
            }
        }

        //
        private void dgdtxtGoalRunRate_TextChanged(object sender, TextChangedEventArgs e)
        {
            if (lblMsg.Visibility == Visibility.Visible)
            {
                WinMachineGoal = dgdSub.CurrentItem as Win_prd_MCRunningGoal_U_Sub_CodeView;

                if (WinMachineGoal != null)
                {
                    TextBox tb1 = sender as TextBox;

                    if (tb1 != null)
                    {
                        WinMachineGoal.CT = tb1.Text;
                    }

                    sender = tb1;
                }
            }
        }

        //
        

        #endregion

        // 서브 그리드 추가
        private void btnAddSub_Click(object sender, RoutedEventArgs e)
        {
            AddDgdSub();
        }

        // 서브그리드 새로운 행 추가
        private void AddDgdSub()
        {
            var MCRunningGoal = new Win_prd_MCRunningGoal_U_Sub_CodeView()
            {
                Num = dgdSub.Items.Count + 1,
                Article = "",
                CT = "",
                BuyerArticleNo = "",
                AutoPassive = "",
                AutoPassiveName = "",
                ImageFile ="",
                ImageFilePath = "",
            };

            lstMG.Add(MCRunningGoal);
            dgdSub.Items.Add(MCRunningGoal);

            SettingDgdSubNum();

            dgdSub.SelectedIndex = dgdSub.Items.Count - 1;
            dgdSub.CurrentCell = new DataGridCellInfo(dgdSub.Items[dgdSub.SelectedIndex], dgdSub.Columns[1]);
        }

        // 서브 그리드 삭제
        private void btnDelSub_Click(object sender, RoutedEventArgs e)
        {
            var MCRunningGoal = dgdSub.SelectedItem as Win_prd_MCRunningGoal_U_Sub_CodeView;
            int index = dgdSub.SelectedIndex;

            if (MCRunningGoal != null)
            {
                dgdSub.Items.Remove(MCRunningGoal);
                lstMG.Remove(MCRunningGoal);

                if (dgdSub.Items.Count > 0)
                {
                    if (index < dgdSub.Items.Count)
                    {
                        dgdSub.SelectedIndex = index;
                    }
                    else
                    {
                        dgdSub.SelectedIndex = --index;
                    }
                }
            }

            SettingDgdSubNum();
        }

        // 서브그리드 Num 세팅
        private void SettingDgdSubNum()
        {
            for (int i = 0; i < dgdSub.Items.Count; i++)
            {
                var MCRunningGoal = dgdSub.Items[i] as Win_prd_MCRunningGoal_U_Sub_CodeView;
                if (MCRunningGoal != null)
                {
                    MCRunningGoal.Num = i + 1;
                }
            }
        }

        #region 기타 메서드 모음

        // 천마리 콤마, 소수점 버리기
        private string stringFormatN0(object obj)
        {
            return string.Format("{0:N0}", obj);
        }

        private string stringFormatN1(object obj)
        {
            return string.Format("{0:N1}", obj);
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

 
        // 서브그리드 설비 자동 수동 콤보박스 세팅
        private void cboAutoPassive_Loaded(object sender, RoutedEventArgs e)
        {
            ComboBox cboSender = sender as ComboBox;
            if (cboSender.ItemsSource == null)
            {
                ObservableCollection<CodeView> ovcAutoPassive = ComboBoxUtil.Instance.GetCMCode_SetComboBox("AUTOPASSIVE", "");
                cboSender.ItemsSource = ovcAutoPassive;
                cboSender.DisplayMemberPath = "code_name";
                cboSender.SelectedValuePath = "code_id";
            }
        }

        // ArticleID 로 Article 정보 가져오기
        private ArticleInfo getArticleInfo(string setArticleID)
        {
            var getArticleInfo = new ArticleInfo();

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

                        getArticleInfo = new ArticleInfo
                        {
                            ArticleGrpID = dr["ArticleGrpID"].ToString(),
                            UnitPrice = dr["UnitPrice"].ToString(),
                            UnitPriceClss = dr["UnitPriceClss"].ToString(),
                            UnitClss = dr["UnitClss"].ToString(),
                            PartGBNID = dr["PartGBNID"].ToString(),
                            BuyerArticleNo = dr["BuyerArticleNo"].ToString(),
                            Article = dr["Article"].ToString(),
                            ArticleID = dr["ArticleID"].ToString(),
                        };
                    }
                }

                return getArticleInfo;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                return null;
            }
        }

        private void cboAutoPassive_DropDownClosed(object sender, EventArgs e)
        {
            ComboBox cboSender = sender as ComboBox;
            var MCSub = cboSender.DataContext as Win_prd_MCRunningGoal_U_Sub_CodeView;

            if (MCSub != null
                && cboSender.SelectedValue != null)
            {
                MCSub.AutoPassive = cboSender.SelectedValue.ToString();
                MCSub.AutoPassiveName = cboSender.Text;
            }
        }

        private void lblProcessAll_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            if (chkProcessAll.IsChecked == true)
            {
                chkProcessAll.IsChecked = false;
            }
            else
            {
                chkProcessAll.IsChecked = true;
            }
        }

        private void chkProcessAll_Checked(object sender, RoutedEventArgs e)
        {
            chkProcessAll.IsChecked = true;
            cboMachine.IsEnabled = false;
        }

        private void chkProcessAll_Unchecked(object sender, RoutedEventArgs e)
        {
            chkProcessAll.IsChecked = false;
            cboMachine.IsEnabled = true;
        }

        // 검색
        private void SubSearch_Click(object sender, RoutedEventArgs e)
        {
            
        }

        // 이전년도
        private void btnLastYear_Click(object sender, RoutedEventArgs e)
        {
            dtpStartDate.SelectedDate = lib.BringLastYearDatetimeContinue(dtpStartDate.SelectedDate.Value)[0];
            dtpEndDate.SelectedDate = lib.BringLastYearDatetimeContinue(dtpStartDate.SelectedDate.Value)[1];
        }

        //금년 버튼 클릭시
        private void btnThisYear_Click(object sender, RoutedEventArgs e)
        {
            dtpStartDate.SelectedDate = Lib.Instance.BringThisYearDatetimeFormat()[0];
            dtpEndDate.SelectedDate = Lib.Instance.BringThisYearDatetimeFormat()[1];
        }

        private void TxtArticle_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                pf.ReturnCode(txtArticle, 76, txtArticle.Text);

                FillGridSub(WinMachine.YYYY, WinMachine.ProcessID, WinMachine.MachineID);

                
            }
        }        
 
        private void cboMachine_DropDownOpened(object sender, EventArgs e) //2021-08-17 콤보박스 재설정
        {
            if (lblMsg.Visibility == Visibility.Visible)
            {
                if (cboProcess.SelectedValue != null
                    && !cboProcess.SelectedValue.ToString().Equals(string.Empty))
                {
                    strProcess = cboProcess.SelectedValue.ToString();

                    this.cboMachine.ItemsSource = null;

                    ObservableCollection<CodeView> ovcMachine = SetMachine(cboProcess.SelectedValue.ToString());
                    this.cboMachine.ItemsSource = ovcMachine;
                    this.cboMachine.DisplayMemberPath = "code_name";
                    this.cboMachine.SelectedValuePath = "code_id";

                }
            }
        }

        private void cboMachine_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            //if (lblMsg.Visibility == Visibility.Visible)
            //{
            //    if (cboProcess.SelectedValue != null
            //        && !cboProcess.SelectedValue.ToString().Equals(string.Empty))
            //    {
            //        strProcess = cboProcess.SelectedValue.ToString();
            //        this.cboMachine.ItemsSource = null;

            //        ObservableCollection<CodeView> ovcMachine = SetMachine(cboProcess.SelectedValue.ToString());
            //        this.cboMachine.ItemsSource = ovcMachine;
            //        this.cboMachine.DisplayMemberPath = "code_name";
            //        this.cboMachine.SelectedValuePath = "code_id";
            //    }
            //}
        }


    }



    #region CodeView
    class Win_prd_MCRunningGoal_U_CodeView : BaseView
    {
        public int Num { get; set; }

        public string YYYY { get; set; }
        public string ProcessID { get; set; }
        public string Process { get; set; }
        public string MachineID { get; set; }
        public string MachineNo { get; set; }
        public string ProcessMachineID { get; set; }
        public string CT { get; set; }
        public string AutoPassive { get; set; }
        public string AutoPassiveName { get; set; }
    }

    public class Win_prd_MCRunningGoal_U_Sub_CodeView : BaseView
    {
        public int Num { get; set; }

        public string ArticleID { get; set; }
        public string Article { get; set; }
        public string BuyerArticleNo { get; set; }
        public string CT { get; set; }
        public string AutoPassive { get; set; }
        public string AutoPassiveName { get; set; }
        public string ImageFile { get; set; }
        public string ImageFilePath { get; set; }
        public bool imageFlag { get; set; }
        public BitmapImage ImageView { get; set; }
    }
    #endregion
}

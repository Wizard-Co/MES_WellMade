/**
 * 
 * @details 공정별 호기 등록
 * @author 정승학
 * @date 2019-07-29
 * @version 2.0
 * 
 * @section MODIFYINFO 수정정보
 * - 수정일        - 수정자       : 수정내역
 * - 2021-10-01     -정승학        :재작성
 * 
 * 
 * */

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

namespace WizMes_WellMade
{
    /// <summary>
    /// Win_prd_MachineCode_U.xaml에 대한 상호 작용 논리
    /// </summary>
    public partial class Win_prd_MachineCode_U : UserControl
    {
        string stDate = string.Empty;
        string stTime = string.Empty;

        Lib lib = new Lib();
        PlusFinder pf = MainWindow.pf;

        string strFlag = string.Empty;
        int rowNum = 0;
        int rowNumSub = 0;

        public Win_prd_MachineCode_U()
        {
            InitializeComponent();
        }

        private void UserControl_Loaded(object sender, RoutedEventArgs e)
        {
            stDate = DateTime.Now.ToString("yyyyMMdd");
            stTime = DateTime.Now.ToString("HHmm");

            DataStore.Instance.InsertLogByFormS(this.GetType().Name, stDate, stTime, "S");

            lib.UiLoading(sender);
            SetComboBox();

            InputMethod.SetIsInputMethodEnabled(this.TextBoxCode, false);
            InputMethod.SetIsInputMethodEnabled(this.TextBoxSetHitCount, false);
        }


        #region 우측 상단 버튼

        #endregion

        #region 콤보박스
        private void SetComboBox()
        {
            try
            {
                //실적창고
                ObservableCollection<CodeView> ovcLOC = ComboBoxUtil.Instance.Gf_DB_CM_GetComCodeDataset(null, "LOC", "Y", "", "");
                this.ComboBoxProductLoc.ItemsSource = ovcLOC;
                this.ComboBoxProductLoc.DisplayMemberPath = "code_name";
                this.ComboBoxProductLoc.SelectedValuePath = "code_id";

                //CommCollectionYN
                ComboBoxCommCollectionYN.Items.Clear();
                ComboBoxCommCollectionYN.Items.Add("Y");
                ComboBoxCommCollectionYN.Items.Add("N");
            }
            catch(Exception ee)
            {
                MessageBox.Show("예외처리 - " + ee.ToString());
            }
        }
        #endregion

        #region Main 조회
        private void FillGrid()
        {
            if(DataGridMain.Items.Count > 0)
            {
                DataGridMain.Items.Clear();
                DataGridSub.Items.Clear();
                ClearData();

                TextBlockCountMain.Text = string.Empty;
                TextBlockCountSub.Text = string.Empty;
            }

            try
            {
                string process = string.Empty;
                process = CheckBoxProcessSearch.IsChecked == true && TextBoxProcessSearch.Text.Trim().Equals("") != true ? TextBoxProcessSearch.Text : "";

                DataSet ds = null;
                Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                sqlParameter.Clear();
                sqlParameter.Add("Process", process);

                ds = DataStore.Instance.ProcedureToDataSet_LogWrite("xp_Code_sProcess", sqlParameter, true, "R");

                if(ds != null && ds.Tables.Count > 0)
                {
                    DataTable dt = ds.Tables[0];
                    int i = 0;

                    if(dt.Rows.Count == 0)
                    {
                        //MessageBox.Show("조회된 공정이 없습니다.");
                    }
                    else
                    {
                        DataRowCollection drc = dt.Rows;

                        foreach (DataRow dr in drc)
                        {
                            var WPMUC = new Win_prd_MachineCode_U_CodeView()
                            {
                                Num = i + 1,
                                ProcessID = dr["ProcessID"].ToString(),
                                Process = dr["Process"].ToString()
                            };

                            DataGridMain.Items.Add(WPMUC);

                            i++;
                        }
                    }
                    TextBlockCountMain.Text = " ▶ 검색 결과 : " + i + " 건";
                }
            }
            catch(Exception ee)
            {
                MessageBox.Show("예외처리 - " + ee.ToString());
            }
            finally
            {
                DataStore.Instance.CloseConnection();
            }

        }
        #endregion

        #region Sub 조회
        private void FillGridMachine(string strProcessID)
        {
            if (DataGridSub.Items.Count > 0)
            {
                DataGridSub.Items.Clear();
            }

            try
            {
                DataSet ds = null;
                Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                sqlParameter.Clear();
                sqlParameter.Add("sProcessID", strProcessID);

                ds = DataStore.Instance.ProcedureToDataSet("xp_Process_sMachine", sqlParameter, false);

                if (ds != null && ds.Tables.Count > 0)
                {
                    DataTable dt = ds.Tables[0];
                    int i = 0;

                    if (dt.Rows.Count > 0)
                    {
                        DataRowCollection drc = dt.Rows;

                        foreach (DataRow dr in drc)
                        {
                            var WPMUPC = new Win_prd_MachineCode_U_ProcessMachine_CodeView()
                            {
                                Num = i + 1,

                                MachineID = dr["MachineID"].ToString(),
                                MachineNO = dr["MachineNO"].ToString(),
                                Machine = dr["Machine"].ToString(),
                                SetHitCount = Convert.ToDecimal(dr["SetHitCount"]),
                                ProductLocID = dr["ProductLocID"].ToString(),
                                ProductLocName = dr["ProductLocName"].ToString(),
                                ProcessID = dr["ProcessID"].ToString(),
                                Process = dr["Process"].ToString(),
                                CommStationNo = dr["CommStationNo"].ToString(),
                                CommIP = dr["CommIP"].ToString(),
                                CommCollectionYN = dr["CommCollectionYN"].ToString()
                            };

                            DataGridSub.Items.Add(WPMUPC);
                            i++;
                        }
                    }
                    TextBlockCountSub.Text = " ▶ 검색 결과 : " + i + " 건";
                }
            }
            catch (Exception ee)
            {
                MessageBox.Show("예외처리 - " + ee.ToString());
            }
            finally
            {
                DataStore.Instance.CloseConnection();
            }
        }


        #endregion

        #region 저장
        private bool SaveData(string strFlag, string strProcessID)
        {
            bool flag = false;

            try
            {
                if (CheckData())
                {
                    List<Procedure> Prolist = new List<Procedure>();
                    List<Dictionary<string, object>> ListParameter = new List<Dictionary<string, object>>();

                    string sProcessID = string.Empty;
                    string sMachineID = string.Empty;
                    string sMachine = string.Empty;
                    string sMachineNO = string.Empty;
                    decimal setHitCount = 0;
                    string sProdLocID = string.Empty;
                    string commStationNo = string.Empty;
                    string commIP = string.Empty;
                    string commCollectionYN = string.Empty;

                    sProcessID = strProcessID;
                    sMachineID = TextBoxCode.Text == string.Empty ? "" : TextBoxCode.Text;
                    sMachine = TextBoxMachineName.Text == string.Empty ? "" : TextBoxMachineName.Text;
                    sMachineNO = TextBoxMachineNo.Text == string.Empty ? "" : TextBoxMachineNo.Text;
                    setHitCount = TextBoxSetHitCount.Text == string.Empty ? 0 : Convert.ToDecimal(TextBoxSetHitCount.Text);
                    sProdLocID = ComboBoxProductLoc.SelectedValue == null ? "" : ComboBoxProductLoc.SelectedValue.ToString();
                    commStationNo = TextBoxCommStationNo.Text == string.Empty ? "" : TextBoxCommStationNo.Text;
                    commIP = TextBoxCommIP.Text == string.Empty ? "" : TextBoxCommIP.Text;
                    commCollectionYN = ComboBoxCommCollectionYN.SelectedValue == null ? "" : ComboBoxCommCollectionYN.SelectedValue.ToString();

                    Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                    sqlParameter.Clear();
                    sqlParameter.Add("sProcessID", sProcessID);
                    sqlParameter.Add("sMachineID", sMachineID);
                    sqlParameter.Add("sMachine", sMachine);
                    sqlParameter.Add("sMachineNO", sMachineNO);
                    sqlParameter.Add("SetHitCount", setHitCount);
                    sqlParameter.Add("sProdLocID", sProdLocID);

                    sqlParameter.Add("CommStationNo", commStationNo);
                    sqlParameter.Add("CommIP", commIP);
                    sqlParameter.Add("CommCollectionYN", commCollectionYN);


                    #region 추가

                    if (strFlag.Equals("I"))
                    {
                        sqlParameter.Add("sCreateuserID", MainWindow.CurrentUser);

                        Procedure pro1 = new Procedure();
                        pro1.Name = "xp_Process_iMachine";
                        pro1.OutputUseYN = "N";
                        pro1.OutputName = "sMachineID";
                        pro1.OutputLength = "2";

                        Prolist.Add(pro1);
                        ListParameter.Add(sqlParameter);

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

                    #endregion

                    #region 수정

                    else if (strFlag.Equals("U"))
                    {
                        sqlParameter.Add("sUpdateUserID", MainWindow.CurrentUser);

                        Procedure pro2 = new Procedure();
                        pro2.Name = "xp_Process_uMachine";
                        pro2.OutputUseYN = "N";
                        pro2.OutputName = "sMachineID";
                        pro2.OutputLength = "2";

                        Prolist.Add(pro2);
                        ListParameter.Add(sqlParameter);

                        string[] Confirm = new string[2];
                        Confirm = DataStore.Instance.ExecuteAllProcedureOutputNew_NewLog(Prolist, ListParameter,"U");
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

                    #endregion
                }
            }
            catch (Exception ee)
            {
                MessageBox.Show("예외처리 - " + ee.ToString());
            }
            finally
            {
                DataStore.Instance.CloseConnection();
            }

            return flag;
        }
        #endregion

        #region 삭제
        private bool DeleteData(string strProcessID, string strMachineID)
        {
            bool flag = false;

            try
            {
                Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                sqlParameter.Clear();
                sqlParameter.Add("sProcessID", strProcessID);
                sqlParameter.Add("sMachineID", strMachineID);

                string[] result = DataStore.Instance.ExecuteProcedure_NewLog("xp_Process_dMachine", sqlParameter, "D");
                DataStore.Instance.CloseConnection();

                if (result[0].Equals("success"))
                {
                    flag = true;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("예외처리 - " + ex.ToString());
            }
            finally
            {
                DataStore.Instance.CloseConnection();
            }

            return flag;
        }
        #endregion

        #region 유효성 검사
        private bool CheckData()
        {
            bool flag = true;

            // 코드는 공백을 제외한 숫자 2자리를 입력해주세요.
            if (TextBoxCode.Text.Trim().Length < 2)
            {
                MessageBox.Show("공백을 제외한 숫자 2자리를 입력해주세요.");
                flag = false;
                return flag;
            }

            // 코드는 숫자만 입력
            if (lib.CheckConvertInt(TextBoxCode.Text) == false)
            {
                MessageBox.Show("코드는 숫자만 입력 가능합니다.");
                flag = false;
                return flag;
            }

            // 코드 중복 체크
            var processID = DataGridMain.SelectedItem as Win_prd_MachineCode_U_CodeView;
            if (processID != null)
            {
                if (strFlag.Trim().Equals("I") && CheckMachineID(processID.ProcessID, TextBoxCode.Text.Trim()) == false)
                {
                    MessageBox.Show("해당 코드는 이미 존재합니다.");
                    flag = false;
                    return flag;
                }
            }

            //기계명 입력
            if (TextBoxMachineName.Text.Length <= 0 || TextBoxMachineName.Text.Equals(""))
            {
                MessageBox.Show("기계명을 입력해주세요.");
                flag = false;
                return flag;
            }

            // 실적 창고
            if (ComboBoxProductLoc.SelectedValue == null)
            {
                MessageBox.Show("실적 창고를 선택해주세요.");
                flag = false;
                return flag;
            }

            // 설정 타점수는 숫자만 입력
            if (!TextBoxSetHitCount.Text.Trim().Equals("") && lib.CheckConvertInt(TextBoxSetHitCount.Text) == false)
            {
                MessageBox.Show("설정 타점수는 숫자만 입력 가능합니다.");
                flag = false;
                return flag;
            }

            return flag;
        }
        #endregion

        #region 호기 코드 중복 검사
        private bool CheckMachineID(string strProcessID, string strMachineID)
        {
            bool flag = true;

            try
            {
                string processID = string.Empty;
                string machineID = string.Empty;

                processID = strProcessID;
                machineID = strMachineID;

                Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                sqlParameter.Clear();
                sqlParameter.Add("ProcessID", processID);
                sqlParameter.Add("MachineID", machineID);

                DataSet ds = DataStore.Instance.ProcedureToDataSet("xp_Machine_sChkMachineID", sqlParameter, false);

                if (ds != null && ds.Tables.Count > 0)
                {
                    DataTable dt = ds.Tables[0];

                    if (dt.Rows.Count != 0)
                    {
                        DataRowCollection drc = dt.Rows;
                        DataRow dr = drc[0];
                        int Cnt = lib.ConvertInt(dr["Cnt"].ToString());

                        if (Cnt > 0)
                        {
                            return false;
                        }
                    }
                }
            }
            catch (Exception ee)
            {
                MessageBox.Show("예외처리 - " + ee.ToString());
            }
            finally
            {
                DataStore.Instance.CloseConnection();
            }

            return flag;
        }

        #endregion

        #region 활성화 & 비활성화
        /// <summary>
        /// 추가, 수정모드(추가 및 수정 버튼 클릭했을 때)
        /// 저장, 취소 버튼만 활성화 나머지 비활성화
        /// </summary>
        private void AddUpdateMode()
        {
            lib.UiButtonEnableChange_SCControl(this);

            GridInputArea.IsHitTestVisible = true;

            DataGridMain.IsHitTestVisible = false;
            DataGridSub.IsHitTestVisible = false;

        }

        /// <summary>
        /// 조회모드(저장 및 취소 버튼 클릭했을 때)
        /// 저장, 취소 버튼만 비활성화 나머지 활성화
        /// </summary>
        private void SaveCancelMode()
        {
            lib.UiButtonEnableChange_IUControl(this);

            GridInputArea.IsHitTestVisible = false;

            DataGridMain.IsHitTestVisible = true;
            DataGridSub.IsHitTestVisible = true;

        }
        #endregion

        #region 입력창 초기화
        private void ClearData()
        {
            this.DataContext = null;

            TextBoxCode.Text = string.Empty;
            TextBoxMachineName.Text = string.Empty;
            TextBoxMachineNo.Text = string.Empty;
            TextBoxSetHitCount.Text = string.Empty;
            ComboBoxProductLoc.SelectedIndex = -1;
            TextBoxCommStationNo.Text = string.Empty;
            TextBoxCommIP.Text = string.Empty;
            ComboBoxCommCollectionYN.SelectedIndex = -1;

        }
        #endregion

        #region ReSearch
        private void Re_Search()
        {
            FillGrid();

            if(DataGridMain.Items.Count <= 0)
            {
                MessageBox.Show("조회된 내용이 없습니다.");
            }
            else
            {
                DataGridMain.SelectedIndex = rowNum;
            }
        }

        private void BeSave()
        {
            var processInfo = DataGridMain.SelectedItem as Win_prd_MachineCode_U_CodeView;

            if (SaveData(strFlag, processInfo.ProcessID))
            {
                SaveCancelMode();
                Re_Search();
                strFlag = "";
            }
        }

        private void BeDelete()
        {
            //var processInfo = new Win_prd_MachineCode_U_CodeView();
            //var machineInfo = new Win_prd_MachineCode_U_ProcessMachine_CodeView();
            //var processInfoSub = new Win_prd_Process_U_Sub_CodeView();
            var processInfo = DataGridMain.SelectedItem as Win_prd_MachineCode_U_CodeView;
            var machineInfo = DataGridSub.SelectedItem as Win_prd_MachineCode_U_ProcessMachine_CodeView;
            if (machineInfo != null)
            {
                if (DeleteData(processInfo.ProcessID, machineInfo.MachineID))
                {
                    Re_Search();
                }
            }
        }

        private void BeCancel()
        {
            SaveCancelMode();
            strFlag = "";

            Re_Search();
        }
        #endregion

        #region 상단 검색조건 입력창

        #endregion

        #region 데이터그리드 선택 데이터 변경

        #endregion

        #region 입력창 이동 이벤트

        private void TextBoxMachineName_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                TextBoxMachineNo.Focus();
            }
        }

        private void TextBoxMachineNo_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                TextBoxSetHitCount.Focus();
            }
        }

        private void TextBoxSetHitCount_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                ComboBoxProductLoc.Focus();
            }
        }

        private void ComboBoxProductLoc_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                TextBoxCommStationNo.Focus();
            }
        }

        private void TextBoxCommStationNo_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                TextBoxCommIP.Focus();
            }
        }

        private void TextBoxCommIP_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                ComboBoxCommCollectionYN.Focus();
            }
        }

       
        #endregion

        #region etc
        //숫자만 입력
        private void TextBoxCheckIsNumeric_PreviewTextInput(object sender, TextCompositionEventArgs e)
        {
            try
            {
                lib.CheckIsNumeric((TextBox)sender, e);
            }
            catch (Exception ee)
            {
                MessageBox.Show("예외처리 - " + ee.ToString());
            }
        }
        #endregion


        //공정명 라벨 클릭
        private void LabelProcessSearch_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            if(CheckBoxProcessSearch.IsChecked == true)
            {
                CheckBoxProcessSearch.IsChecked = false;
            }
            else
            {
                CheckBoxProcessSearch.IsChecked = true;
            }
        }

        //공정명 체크박스
        private void CheckBoxProcessSearch_Checked(object sender, RoutedEventArgs e)
        {
            CheckBoxProcessSearch.IsChecked = true;
            TextBoxProcessSearch.IsEnabled = true;
        }

        //공정명 체크박스
        private void CheckBoxProcessSearch_Unchecked(object sender, RoutedEventArgs e)
        {
            CheckBoxProcessSearch.IsChecked = false;
            TextBoxProcessSearch.IsEnabled = false;
        }

        //공정명 텍스트박스 키다운 이벤트
        private void TextBoxProcessSearch_KeyDown(object sender, KeyEventArgs e)
        {
            if(e.Key == Key.Enter)
            {
                btnSearch_Click(null, null);
            }
        }

        //추가 버튼
        private void btnAdd_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var process = DataGridMain.SelectedItem as Win_prd_MachineCode_U_CodeView;

                if (process != null)
                {
                    strFlag = "I";
                    AddUpdateMode();
                    rowNum = DataGridMain.Items.Count == 0 ? 0 : DataGridMain.SelectedIndex;
                    rowNumSub = DataGridSub.Items.Count == 0 ? 0 : DataGridSub.SelectedIndex;

                    ClearData();

                }
                else
                {
                    MessageBox.Show("공정을 먼저 선택해 주세요.");
                    return;
                }

                TextBoxMachineName.Focus();

            }
            catch (Exception ee)
            {
                MessageBox.Show("예외처리 - " + ee.ToString());
            }
        }

        //수정 버튼
        private void btnUpdate_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var processInfo = DataGridMain.SelectedItem as Win_prd_MachineCode_U_CodeView;
                if (processInfo != null)
                {
                    var machineInfo = DataGridSub.SelectedItem as Win_prd_MachineCode_U_ProcessMachine_CodeView;
                    if (machineInfo != null)
                    {
                        rowNum = DataGridMain.SelectedIndex;
                        rowNumSub = DataGridSub.SelectedIndex;

                        AddUpdateMode();
                        strFlag = "U";
                    }
                    else
                    {
                        MessageBox.Show("수정할 Machine(설비)정보를 선택해주세요");
                    }
                }
                else
                {
                    MessageBox.Show("먼저 공정을 선택한 후 수정할 Machine(설비)정보를 선택해주세요");
                    return;
                }
            }
            catch(Exception ee)
            {
                MessageBox.Show("예외처리 - " + ee.ToString());
            }
        }

        //삭제 버튼
        private void btnDelete_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var processInfo = DataGridMain.SelectedItem as Win_prd_MachineCode_U_CodeView;
                if(processInfo != null)
                {
                    var machineInfo = DataGridSub.SelectedItem as Win_prd_MachineCode_U_ProcessMachine_CodeView;

                    if (machineInfo == null)
                    {
                        MessageBox.Show("삭제할 데이터가 지정되지 않았습니다. 데이터를 지정하고 눌러주세요");
                    }
                    else
                    {
                        if(WorkCheck_MachineID(processInfo.ProcessID, machineInfo.MachineID) == true)
                        {
                            MessageBox.Show("해당 공정의 호기가 사용중입니다.");
                        }
                        else
                        {
                            rowNum = DataGridMain.SelectedIndex;
                            if (DataGridMain.Items.Count > 0 && DataGridSub.SelectedItem != null)
                            {
                                rowNumSub = DataGridSub.SelectedIndex;
                            }

                            if (MessageBox.Show("선택하신 항목을 삭제하시겠습니까? \r\n " + machineInfo.MachineID + " / " + machineInfo.Machine, "삭제 전 확인", MessageBoxButton.YesNo) == MessageBoxResult.Yes)
                            {
                                using (Loading lw = new Loading(BeDelete))
                                {
                                    lw.ShowDialog();
                                }
                            }
                        }
                    }
                }
                else
                {
                    MessageBox.Show("먼저 공정을 선택해주세요");
                    return;
                }
            }
            catch (Exception ee)
            {
                MessageBox.Show("예외처리 - " + ee.ToString());
            }
        }

        //닫기 버튼
        private void btnClose_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                DataStore.Instance.InsertLogByFormS(this.GetType().Name, stDate, stTime, "E");

                this.DataContext = null;
                DataGridMain = null;
                DataGridSub = null;

                lib.ChildMenuClose(this.ToString());
            }
            catch(Exception ee)
            {
                MessageBox.Show("예외처리 - " + ee.ToString());
            }
        }

        //조회 버튼
        private void btnSearch_Click(object sender, RoutedEventArgs e)
        {
            // 검색버튼 비활성화
            btnSearch.IsEnabled = false;

            Dispatcher.BeginInvoke(new Action(() =>

            {
                //로직
                if (DataGridMain.SelectedItem == null)
                {
                    rowNum = 0;
                }
                else
                {
                    rowNum = DataGridMain.SelectedIndex;
                    rowNumSub = DataGridSub.SelectedIndex;
                }

                rowNum = 0;
                using (Loading lw = new Loading(Re_Search))
                {
                    lw.ShowDialog();
                }

            }), System.Windows.Threading.DispatcherPriority.Background);



            Dispatcher.BeginInvoke(new Action(() =>

            {
                btnSearch.IsEnabled = true;

            }), System.Windows.Threading.DispatcherPriority.Background);

        }

        //저장 버튼
        private void btnSave_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                using (Loading lw = new Loading(BeSave))
                {
                    lw.ShowDialog();
                }
            }
            catch(Exception ee)
            {
                MessageBox.Show("예외처리 -  " + ee.ToString());
            }
        }

        //취소 버튼
        private void btnCancel_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                using (Loading lw = new Loading(BeCancel))
                {
                    lw.ShowDialog();
                }
            }
            catch (Exception ee)
            {
                MessageBox.Show("예외처리 - " + ee.ToString());
            }
        }

        //엑셀 버튼
        private void btnExcel_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                DataTable dt = null;
                string Name = string.Empty;

                string[] dgdStr = new string[4];
                dgdStr[0] = "공정";
                dgdStr[1] = "공정별 Machine";
                dgdStr[2] = DataGridMain.Name;
                dgdStr[3] = DataGridSub.Name;

                ExportExcelxaml ExpExc = new ExportExcelxaml(dgdStr);
                ExpExc.ShowDialog();

                if (ExpExc.DialogResult.HasValue)
                {
                    if (ExpExc.choice.Equals(DataGridMain.Name))
                    {
                        DataStore.Instance.InsertLogByForm(this.GetType().Name, "E");
                        if (ExpExc.Check.Equals("Y"))
                            dt = Lib.Instance.DataGridToDTinHidden(DataGridMain);
                        else
                            dt = Lib.Instance.DataGirdToDataTable(DataGridMain);

                        Name = DataGridMain.Name;
                        if (Lib.Instance.GenerateExcel(dt, Name))
                            Lib.Instance.excel.Visible = true;
                        else
                            return;
                    }
                    else if (ExpExc.choice.Equals(DataGridSub.Name))
                    {
                        if (ExpExc.Check.Equals("Y"))
                            dt = Lib.Instance.DataGridToDTinHidden(DataGridSub);
                        else
                            dt = Lib.Instance.DataGirdToDataTable(DataGridSub);

                        Name = DataGridSub.Name;
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
            catch (Exception ee)
            {
                MessageBox.Show("예외처리 -  " + ee.ToString());
            }
            finally
            {
                lib.ReleaseExcelObject(lib.workSheet);
                lib.ReleaseExcelObject(lib.workBook);
                lib.ReleaseExcelObject(lib.excel);
            }
        }

        //메인 데이터그리드 선택 데이터 변경
        private void DataGridMain_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            try
            {
                var process = DataGridMain.SelectedItem as Win_prd_MachineCode_U_CodeView;

                if(process != null)
                {
                    FillGridMachine(process.ProcessID);

                    if(DataGridSub.Items.Count > 0)
                    {
                        DataGridSub.SelectedIndex = rowNumSub;
                    }
                    else
                    {
                        this.DataContext = null;
                        ClearData();
                        
                    }
                }
            }
            catch(Exception ee)
            {
                MessageBox.Show("예외처리 - " + ee.ToString());
            }
        }

        //서브 데이터그리드 선택 데이터 변경
        private void DataGridSub_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            try
            {
                var processMachine = DataGridSub.SelectedItem as Win_prd_MachineCode_U_ProcessMachine_CodeView;

                if(processMachine != null)
                {
                    this.DataContext = processMachine;
                }

            }
            catch (Exception ee)
            {
                MessageBox.Show("예외처리 - " + ee.ToString());
            }
        }

        private void ComboBoxProductLoc_DropDownClosed(object sender, EventArgs e)
        {
            try
            {
                TextBoxCommStationNo.Focus();
            }
            catch (Exception ee)
            {
                MessageBox.Show("예외처리 - " + ee.ToString());
            }
        }
        
        private void ComboBoxCommCollectionYN_DropDownClosed(object sender, EventArgs e)
        {
            try
            {

            }
            catch (Exception ee)
            {
                MessageBox.Show("예외처리 - " + ee.ToString());
            }
        }

        #region 현재 사용중인 설비 체크
        private bool WorkCheck_MachineID(string processID, string machineID)
        {
            bool result = false;

            try
            {
                Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                sqlParameter.Clear();
                sqlParameter.Add("ProcessID", processID);
                sqlParameter.Add("MachineID", machineID);

                DataSet ds = DataStore.Instance.ProcedureToDataSet("xp_prd_Machine_WorkCheck", sqlParameter, false);

                if(ds != null && ds.Tables.Count > 0)
                {
                    DataTable dt = ds.Tables[0];

                    if(dt.Rows.Count > 0)
                    {
                        result = true;
                    }
                    else
                    {
                        result = false;
                    }
                }

            }
            catch(Exception ee)
            {
                MessageBox.Show("예외처리 - " + ee.ToString());
            }
            finally
            {
                DataStore.Instance.CloseConnection();
            }

            return result;
        }

        #endregion

    }

    #region CodeView
    class Win_prd_MachineCode_U_CodeView : BaseView
    {
        public override string ToString()
        {
            return (this.ReportAllProperties());
        }

        public int Num { get; set; }
        public string ProcessID { get; set; }
        public string Process { get; set; }
        public string DetailProcessYN { get; set; }

    }

    class Win_prd_MachineCode_U_ProcessMachine_CodeView : BaseView
    {
        public override string ToString()
        {
            return (this.ReportAllProperties());
        }

        public int Num { get; set; }
        public string MachineID { get; set; }
        public string Machine { get; set; }
        public string MachineNO { get; set; }
        public decimal SetHitCount { get; set; }
        public string ProductLocID { get; set; }
        public string ProductLocName { get; set; }
        public string ProcessID { get; set; }
        public string Process { get; set; }
        public string CommStationNo { get; set; }
        public string CommIP { get; set; }
        public string CommCollectionYN { get; set; }
    }

    #endregion

}

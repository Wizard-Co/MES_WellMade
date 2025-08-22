using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Text;
using System.Threading;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using WizMes_WellMade.PopUP;

namespace WizMes_WellMade
{
    /**************************************************************************************************
    '** System 명 : WizMes_GLS
    '** Author    : Wizard
    '** 작성자    : 최준호
    '** 내용      : 계측기등록
    '** 생성일자  : 2019.04.15
    '** 변경일자  : 
    '**------------------------------------------------------------------------------------------------
    ''*************************************************************************************************
    ' 변경일자  , 변경자, 요청자    , 요구사항ID  , 요청 및 작업내용
    '**************************************************************************************************
    ' ex) 2015.11.09, 박진성, 오영      ,S_201510_AFT_03 , 월별집계(가로) 순서 변경 : 합계/10월/9월/8월 순으로
    '**************************************************************************************************/

    /// <summary>
    /// Win_Qul_MeasureMachine_U.xaml에 대한 상호 작용 논리
    /// </summary>
    public partial class Win_Qul_MeasureMachine_U : UserControl
    {
        string strFlag = string.Empty;
        int selectedIndex = 0;
        List<string[]> listFtpFile = new List<string[]>();
        Lib lib = new Lib();

        /// <summary>
        /// Main 그리드용
        /// </summary>
        ObservableCollection<Win_Qul_MeasureMachine_U_CodeView> ovcMeasureMachine
            = new ObservableCollection<Win_Qul_MeasureMachine_U_CodeView>();

        WizMes_WellMade.PopUp.NoticeMessage msg = new PopUp.NoticeMessage();

        // FTP 활용모음.
        string strImagePath = string.Empty;
        string strFullPath = string.Empty;

        string strAttPath1 = string.Empty;
        string strAttPath2 = string.Empty;
        string strAttPath3 = string.Empty;

        string FullPath1 = string.Empty;
        string FullPath2 = string.Empty;
        string FullPath3 = string.Empty;

        private FTP_EX _ftp = null;
        //string FTP_ADDRESS = "ftp://wizis.iptime.org/ImageData/Measure";
        string FTP_ADDRESS = "ftp://" + LoadINI.FileSvr + ":" + LoadINI.FTPPort + LoadINI.FtpImagePath + "/Measure";
        //string FTP_ADDRESS = "ftp://222.104.222.145:25000/ImageData/Measure";
        //string FTP_ADDRESS = "ftp://192.168.0.95/ImageData/Measure";
        private const string FTP_ID = "wizuser";
        private const string FTP_PASS = "wiz9999";
        private const string LOCAL_DOWN_PATH = "C:\\Temp";

        public Win_Qul_MeasureMachine_U()
        {
            InitializeComponent();
        }

        private void UserControl_Loaded(object sender, RoutedEventArgs e)
        {
            Lib.Instance.UiLoading(this);
            SetComboBox();
        }

        void SetComboBox()
        {
            int Year = DateTime.Today.Year + 1;
            for (int i = 0; i < 10; i++)
            {
                Year = Year - 1;
                cboYear.Items.Add(Year);
            }
            cboYear.SelectedIndex = 0;
        }

        //계측기명
        private void LblMsrMachineNameSrh_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            if (chkMsrMachineNameSrh.IsChecked == true) { chkMsrMachineNameSrh.IsChecked = false; }
            else { chkMsrMachineNameSrh.IsChecked = true; }
        }

        //계측기명
        private void ChkMsrMachineNameSrh_Checked(object sender, RoutedEventArgs e)
        {
            txtMsrMachineNameSrh.IsEnabled = true;
            txtMsrMachineNameSrh.Focus();
        }

        //계측기명
        private void ChkMsrMachineNameSrh_Unchecked(object sender, RoutedEventArgs e)
        {
            txtMsrMachineNameSrh.IsEnabled = false;
        }


        /// <summary>
        /// 추가,수정 시 동작 모음
        /// </summary>
        private void ControlVisibleAndEnable_AU()
        {
            Lib.Instance.UiButtonEnableChange_SCControl(this);
            dgdMain.IsHitTestVisible = false;
            grdInput.IsHitTestVisible = true;

            txtMsrMachineID.IsHitTestVisible = false;
        }

        /// <summary>
        /// 저장,취소 시 동작 모음
        /// </summary>
        private void ControlVisibleAndEnable_SC()
        {
            Lib.Instance.UiButtonEnableChange_IUControl(this);
            dgdMain.IsHitTestVisible = true;
            grdInput.IsHitTestVisible = false;
            listFtpFile.Clear();
        }

        //추가
        private void BtnAdd_Click(object sender, RoutedEventArgs e)
        {
            if (dgdMain.SelectedItem != null)
            {
                selectedIndex = dgdMain.SelectedIndex;
            }
            this.DataContext = null;
            strFlag = "I";
            ControlVisibleAndEnable_AU();
            tbkMsg.Text = "자료 추가 중";

            txtMsrMachineMgrNo.Focus();
            dtpMsrMachineBuyDate.SelectedDate = DateTime.Now;
            dtpMsrMachineSetDate.SelectedDate = DateTime.Now;
        }

        //수정
        private void BtnUpdate_Click(object sender, RoutedEventArgs e)
        {
            var uMeasureMachine = dgdMain.SelectedItem as Win_Qul_MeasureMachine_U_CodeView;
            if (uMeasureMachine == null)
            {
                MessageBox.Show("수정할 데이터가 지정되지 않았습니다. 수정데이터를 지정하고 눌러주세요");
                return;
            }
            else
            {
                selectedIndex = dgdMain.SelectedIndex;
                strFlag = "U";
                ControlVisibleAndEnable_AU();
                tbkMsg.Text = "자료 수정 중";
            }
        }

        //삭제
        private void BtnDelete_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var dMeasureMachine = dgdMain.SelectedItem as Win_Qul_MeasureMachine_U_CodeView;
                if (dMeasureMachine == null)
                {
                    MessageBox.Show("삭제할 데이터가 지정되지 않았습니다. 삭제데이터를 지정하고 눌러주세요");
                    return;
                }
                else
                {
                    if (MessageBox.Show("선택하신 항목을 삭제하시겠습니까?", "삭제 전 확인", MessageBoxButton.YesNo) == MessageBoxResult.Yes)
                    {
                        if (dgdMain.Items.Count > 0 && dgdMain.SelectedItem != null)
                        {
                            selectedIndex = dgdMain.SelectedIndex;
                        }

                        if (Procedure.Instance.DeleteData(dMeasureMachine.MsrMachineID, "sMsrMachineID"
                            , "xp_MeasureMachine_dMeasureMachine"))
                        {
                            // 접속 경로
                            _ftp = new FTP_EX(FTP_ADDRESS, FTP_ID, FTP_PASS);
                            string[] fileListSimple;
                            fileListSimple = _ftp.directoryListSimple("", Encoding.Default);

                            bool delFtp = FolderInfoAndFlag(fileListSimple, dMeasureMachine.MsrMachineID);
                            if (delFtp)
                                _ftp.removeDir(dMeasureMachine.MsrMachineID);

                            selectedIndex -= 1;
                            FillGrid();
                            if (dgdMain.Items.Count > 0)
                            {
                                dgdMain.SelectedIndex = selectedIndex;
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

        //닫기
        private void BtnClose_Click(object sender, RoutedEventArgs e)
        {
            Lib.Instance.ChildMenuClose(this.ToString());
        }

        //조회
        private void BtnSearch_Click(object sender, RoutedEventArgs e)
        {
            //검색버튼 비활성화
            btnSearch.IsEnabled = false;

            Dispatcher.BeginInvoke(new Action(() =>

            {
                Thread.Sleep(2000);

                //로직
                FillGrid();

                if (dgdMain.Items.Count > 0)
                {
                    dgdMain.SelectedIndex = selectedIndex;
                }

            }), System.Windows.Threading.DispatcherPriority.Background);



            Dispatcher.BeginInvoke(new Action(() =>

            {
                btnSearch.IsEnabled = true;

            }), System.Windows.Threading.DispatcherPriority.Background);
        }

        //저장
        private void BtnSave_Click(object sender, RoutedEventArgs e)
        {
            if (SaveData())
            {
                ControlVisibleAndEnable_SC();
                //저장,취소하면 삭제를 위해 담아둔 뷰를 모두 삭제
                ovcMeasureMachine.Clear();

                FillGrid();
                if (dgdMain.Items.Count > 0)
                {
                    if (strFlag == "I")
                    {
                        dgdMain.SelectedIndex = dgdMain.Items.Count - 1;
                        dgdMain.Focus();
                    }
                    else
                    {
                        dgdMain.SelectedIndex = selectedIndex;
                        Focus();
                    }
                }
                strFlag = string.Empty;
            }
        }

        //취소
        private void BtnCancel_Click(object sender, RoutedEventArgs e)
        {
            ControlVisibleAndEnable_SC();
            strFlag = string.Empty;

            //저장,취소하면 삭제를 위해 담아둔 뷰를 모두 삭제
            ovcMeasureMachine.Clear();

            FillGrid();
            if (dgdMain.Items.Count > 0)
            {
                dgdMain.SelectedIndex = selectedIndex;
                dgdMain.Focus();
            }
        }

        //엑셀
        private void BtnExcel_Click(object sender, RoutedEventArgs e)
        {
            DataTable dt = null;
            string Name = string.Empty;
            Lib lib = new Lib();

            string[] dgdStr = new string[2];
            dgdStr[0] = "계측기 등록 목록";
            dgdStr[1] = dgdMain.Name;

            ExportExcelxaml ExpExc = new ExportExcelxaml(dgdStr);
            ExpExc.ShowDialog();

            if (ExpExc.DialogResult.HasValue)
            {
                if (ExpExc.choice.Equals(dgdMain.Name))
                {
                    if (ExpExc.Check.Equals("Y"))
                        dt = lib.DataGridToDTinHidden(dgdMain);
                    else
                        dt = lib.DataGirdToDataTable(dgdMain);

                    Name = dgdMain.Name;
                    if (lib.GenerateExcel(dt, Name))
                    {
                        lib.excel.Visible = true;
                        lib.ReleaseExcelObject(lib.excel);
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
            }
            lib = null;
        }

        // 조회
        private void FillGrid()
        {
            try
            {
                if (cboYear.SelectedValue == null)
                {
                    MessageBox.Show("검색조건에 연도는 필수입니다.");
                    return;
                }

                ovcMeasureMachine.Clear();

                Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                sqlParameter.Add("chkMachineName", chkMsrMachineNameSrh.IsChecked == true ? 1 : 0);
                sqlParameter.Add("MsrMachineName", chkMsrMachineNameSrh.IsChecked == true ?
                    txtMsrMachineNameSrh.Text : "");
                sqlParameter.Add("YYYY", cboYear.SelectedValue.ToString());
                DataSet ds = DataStore.Instance.ProcedureToDataSet("xp_MeasureMachine_sMeasureMachine", sqlParameter, false);

                if (ds != null && ds.Tables.Count > 0)
                {
                    DataTable dt = ds.Tables[0];

                    if (dt.Rows.Count > 0)
                    {
                        DataRowCollection drc = dt.Rows;
                        int i = 0;

                        foreach (DataRow dr in drc)
                        {
                            i++;
                            var varMeasureMachine = new Win_Qul_MeasureMachine_U_CodeView
                            {
                                Num = i,

                                MsrMachineID = dr["MsrMachineID"].ToString(),
                                MsrMachineMgrNo = dr["MsrMachineMgrNo"].ToString(),
                                MsrMachineName = dr["MsrMachineName"].ToString(),
                                MsrMachineNo = dr["MsrMachineNo"].ToString(),
                                MsrMachineSpec = dr["MsrMachineSpec"].ToString(),

                                MsrMachineMsrBuyCustom = dr["MsrMachineMsrBuyCustom"].ToString(),
                                MsrMachineBuyDate = dr["MsrMachineBuyDate"].ToString().Trim(),
                                MsrMachineMsrCustom = dr["MsrMachineMsrCustom"].ToString(),
                                Comments = dr["Comments"].ToString(),
                                ProofCycle = dr["ProofCycle"].ToString(),

                                LastProofDate = dr["LastProofDate"].ToString().Trim(),
                                NextProofDate = dr["NextProofDate"].ToString().Trim(),
                                AttPath1 = dr["AttPath1"].ToString(),
                                AttFile1 = dr["AttFile1"].ToString(),
                                AttPath2 = dr["AttPath2"].ToString(),

                                AttFile2 = dr["AttFile2"].ToString(),
                                AttPath3 = dr["AttPath3"].ToString(),
                                AttFile3 = dr["AttFile3"].ToString(),
                                MsrMachineRange = dr["MsrMachineRange"].ToString(),
                                MsrMachinePrice = stringFormatN0(dr["MsrMachinePrice"]),

                                MsrMachineSetDate = dr["MsrMachineSetDate"].ToString().Trim(),
                                MsrMachineUseTeam = dr["MsrMachineUseTeam"].ToString(),
                                MsrmachinePerson = dr["MsrmachinePerson"].ToString(),
                                PR1 = dr["PR1"].ToString(),
                                PR2 = dr["PR2"].ToString(),

                                PR3 = dr["PR3"].ToString(),
                                PR4 = dr["PR4"].ToString(),
                                PR5 = dr["PR5"].ToString(),
                                PR6 = dr["PR6"].ToString(),
                                PR7 = dr["PR7"].ToString(),

                                PR8 = dr["PR8"].ToString(),
                                PR9 = dr["PR9"].ToString(),
                                PR10 = dr["PR10"].ToString(),
                                PR11 = dr["PR11"].ToString(),
                                PR12 = dr["PR12"].ToString(),

                                R1 = dr["R1"].ToString(),
                                R2 = dr["R2"].ToString(),
                                R3 = dr["R3"].ToString(),
                                R4 = dr["R4"].ToString(),
                                R5 = dr["R5"].ToString(),

                                R6 = dr["R6"].ToString(),
                                R7 = dr["R7"].ToString(),
                                R8 = dr["R8"].ToString(),
                                R9 = dr["R9"].ToString(),
                                R10 = dr["R10"].ToString(),

                                R11 = dr["R11"].ToString(),
                                R12 = dr["R12"].ToString()
                            };

                            if (!Lib.Instance.CheckNull(varMeasureMachine.LastProofDate).Equals(string.Empty))
                            {
                                varMeasureMachine.LastProofDate_CV =
                                    Lib.Instance.StrDateTimeBar(varMeasureMachine.LastProofDate);
                            }
                            if (!Lib.Instance.CheckNull(varMeasureMachine.NextProofDate).Equals(string.Empty))
                            {
                                varMeasureMachine.NextProofDate_CV =
                                    Lib.Instance.Left(varMeasureMachine.NextProofDate, 10);
                            }
                            if (!Lib.Instance.CheckNull(varMeasureMachine.MsrMachineBuyDate).Equals(string.Empty))
                            {
                                varMeasureMachine.MsrMachineBuyDate_CV =
                                    Lib.Instance.StrDateTimeBar(varMeasureMachine.MsrMachineBuyDate);
                            }
                            if (!Lib.Instance.CheckNull(varMeasureMachine.MsrMachineSetDate).Equals(string.Empty))
                            {
                                varMeasureMachine.MsrMachineSetDate_CV =
                                    Lib.Instance.StrDateTimeBar(varMeasureMachine.MsrMachineSetDate);
                            }

                            ovcMeasureMachine.Add(varMeasureMachine);
                        }
                        tbkIndexCount.Text = "▶검색결과 : " + i + " 건";
                        dgdMain.ItemsSource = ovcMeasureMachine;
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

        //표의 행 선택 변경
        private void DgdMain_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            var MeasureMachine = dgdMain.SelectedItem as Win_Qul_MeasureMachine_U_CodeView;

            if (MeasureMachine != null)
            {
                this.DataContext = MeasureMachine;

                if (MeasureMachine.MsrMachineBuyDate_CV == null)
                {
                    dtpMsrMachineBuyDate.SelectedDate = DateTime.Today;
                }
            }
        }

        //표 더블클릭시 수정
        private void dgdMain_LeftDoubleDown(object sender, MouseButtonEventArgs e)
        {
            if (e.ClickCount == 2)
            {
                BtnUpdate_Click(btnUpdate, null);
            }
        }

        //저장
        private bool SaveData()
        {
            bool flag = false;
            List<Procedure> Prolist = new List<Procedure>();
            List<Dictionary<string, object>> ListParameter = new List<Dictionary<string, object>>();

            try
            {
                if (CheckData())
                {
                    double D_MsrMachinePrice = 0;
                    double.TryParse(txtMsrMachinePrice.Text, out D_MsrMachinePrice);

                    Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                    sqlParameter.Clear();
                    sqlParameter.Add("sMsrMachineID", txtMsrMachineID.Text);
                    sqlParameter.Add("MsrMachineMgrNo", txtMsrMachineMgrNo.Text);
                    sqlParameter.Add("MsrMachineName", txtMsrMachineName.Text);
                    sqlParameter.Add("MsrMachineNo", txtMsrMachineNo.Text);
                    sqlParameter.Add("MsrMachineSpec", txtMsrMachineSpec.Text);

                    sqlParameter.Add("MsrMachineMsrBuyCustom", txtMsrMachineMsrBuyCustom.Text);
                    sqlParameter.Add("MsrMachineBuyDate", chkBuyDate.IsChecked == true ?
                        dtpMsrMachineBuyDate.SelectedDate.Value.ToString("yyyyMMdd") : "");
                    sqlParameter.Add("MsrMachineMsrCustom", txtMsrMachineMsrCustom.Text);
                    sqlParameter.Add("Comments", txtComments.Text);
                    sqlParameter.Add("ProofCycle", txtCycle.Text);

                    sqlParameter.Add("ProofCycleUnit", txtCycle.Text);
                    sqlParameter.Add("AttPath1", "");
                    sqlParameter.Add("AttFile1", "");
                    sqlParameter.Add("AttPath2", "");
                    sqlParameter.Add("AttFile2", "");

                    sqlParameter.Add("AttPath3", "");
                    sqlParameter.Add("AttFile3", "");

                    if (strFlag.Equals("I"))
                    {
                        sqlParameter.Add("MsrMachineRange", txtMsrMachineRange.Text);
                        sqlParameter.Add("MsrMachinePrice", D_MsrMachinePrice);
                        sqlParameter.Add("MsrMachineUseTeam", txtMsrMachineUseTeam.Text);
                        sqlParameter.Add("MsrMachineSetDate", dtpMsrMachineSetDate.SelectedDate != null ?
                            dtpMsrMachineSetDate.SelectedDate.Value.ToString("yyyyMMdd") : "");
                        sqlParameter.Add("MsrMachinePerson", txtMsrMachinePerson.Text);

                        sqlParameter.Add("PR1", chk01.IsChecked == true ? "1" : "0");
                        sqlParameter.Add("PR2", chk02.IsChecked == true ? "1" : "0");
                        sqlParameter.Add("PR3", chk03.IsChecked == true ? "1" : "0");
                        sqlParameter.Add("PR4", chk04.IsChecked == true ? "1" : "0");
                        sqlParameter.Add("PR5", chk05.IsChecked == true ? "1" : "0");

                        sqlParameter.Add("PR6", chk06.IsChecked == true ? "1" : "0");
                        sqlParameter.Add("PR7", chk07.IsChecked == true ? "1" : "0");
                        sqlParameter.Add("PR8", chk08.IsChecked == true ? "1" : "0");
                        sqlParameter.Add("PR9", chk09.IsChecked == true ? "1" : "0");
                        sqlParameter.Add("PR10", chk10.IsChecked == true ? "1" : "0");

                        sqlParameter.Add("PR11", chk11.IsChecked == true ? "1" : "0");
                        sqlParameter.Add("PR12", chk12.IsChecked == true ? "1" : "0");
                        sqlParameter.Add("CreateUserID", MainWindow.CurrentUser);

                        Procedure pro1 = new Procedure();
                        pro1.Name = "xp_MeasureMachine_iMeasureMachine";
                        pro1.OutputUseYN = "Y";
                        pro1.OutputName = "sMsrMachineID";
                        pro1.OutputLength = "5";

                        Prolist.Add(pro1);
                        ListParameter.Add(sqlParameter);

                        List<KeyValue> list_Result = new List<KeyValue>();
                        list_Result = DataStore.Instance.ExecuteAllProcedureOutputGetCS(Prolist, ListParameter);
                        string sGetID = string.Empty;

                        if (list_Result[0].key.ToLower() == "success")
                        {
                            list_Result.RemoveAt(0);
                            for (int i = 0; i < list_Result.Count; i++)
                            {
                                KeyValue kv = list_Result[i];
                                if (kv.key == "sMsrMachineID")
                                {
                                    sGetID = kv.value;
                                    flag = true;
                                }
                            }
                            if (flag)
                            {
                                bool AttachYesNo = false;
                                if (txtAttFile1.Text != string.Empty || txtAttFile2.Text != string.Empty || txtAttFile3.Text != string.Empty)       //첨부파일 1
                                {
                                    if (FTP_Save_File(listFtpFile, sGetID))
                                    {
                                        if (!txtAttFile1.Text.Equals(string.Empty)) { txtAttFile1.Tag = "/ImageData/Measure/" + sGetID; }
                                        if (!txtAttFile2.Text.Equals(string.Empty)) { txtAttFile2.Tag = "/ImageData/Measure/" + sGetID; }
                                        if (!txtAttFile3.Text.Equals(string.Empty)) { txtAttFile3.Tag = "/ImageData/Measure/" + sGetID; }

                                        AttachYesNo = true;
                                    }
                                    else
                                    { MessageBox.Show("데이터 저장이 완료되었지만, 첨부문서 등록에 실패하였습니다."); }

                                    if (AttachYesNo == true) { UpdateFtpPathAndNameData(sGetID); }      //첨부문서 정보 DB 업데이트.
                                }
                            }
                        }
                        else
                        {
                            MessageBox.Show("[저장실패]\r\n" + list_Result[0].value.ToString());
                            flag = false;
                        }
                    }
                    else
                    {
                        sqlParameter.Add("MsrMachinePrice", D_MsrMachinePrice);
                        sqlParameter.Add("UpdateUserID", MainWindow.CurrentUser);
                        //2021-04-22 update할때 파라미터 값이 있어야 되서 추가
                        sqlParameter.Add("MsrMachineUseTeam", txtMsrMachineUseTeam.Text);
                        sqlParameter.Add("MsrMachineSetDate", dtpMsrMachineSetDate.SelectedDate != null ?
                            dtpMsrMachineSetDate.SelectedDate.Value.ToString("yyyyMMdd") : "");
                        sqlParameter.Add("MsrMachinePerson", txtMsrMachinePerson.Text);

                        sqlParameter.Add("PR1", chk01.IsChecked == true ? "1" : "0");
                        sqlParameter.Add("PR2", chk02.IsChecked == true ? "1" : "0");
                        sqlParameter.Add("PR3", chk03.IsChecked == true ? "1" : "0");
                        sqlParameter.Add("PR4", chk04.IsChecked == true ? "1" : "0");
                        sqlParameter.Add("PR5", chk05.IsChecked == true ? "1" : "0");

                        sqlParameter.Add("PR6", chk06.IsChecked == true ? "1" : "0");
                        sqlParameter.Add("PR7", chk07.IsChecked == true ? "1" : "0");
                        sqlParameter.Add("PR8", chk08.IsChecked == true ? "1" : "0");
                        sqlParameter.Add("PR9", chk09.IsChecked == true ? "1" : "0");
                        sqlParameter.Add("PR10", chk10.IsChecked == true ? "1" : "0");

                        sqlParameter.Add("PR11", chk11.IsChecked == true ? "1" : "0");
                        sqlParameter.Add("PR12", chk12.IsChecked == true ? "1" : "0");

                        Procedure pro1 = new Procedure();
                        pro1.Name = "xp_MeasureMachine_uMeasureMachine";
                        pro1.OutputUseYN = "N";
                        pro1.OutputName = "sMsrMachineID";
                        pro1.OutputLength = "5";

                        Prolist.Add(pro1);
                        ListParameter.Add(sqlParameter);

                        string[] Confirm = new string[2];
                        Confirm = DataStore.Instance.ExecuteAllProcedureOutputNew(Prolist, ListParameter);
                        if (Confirm[0] != "success")
                        {
                            MessageBox.Show("[저장실패]\r\n" + Confirm[1].ToString());
                            flag = false;
                        }
                        else
                        {
                            flag = true;
                        }
                        if (flag)
                        {
                            bool AttachYesNo = false;
                            if (txtAttFile1.Text != string.Empty || txtAttFile2.Text != string.Empty || txtAttFile3.Text != string.Empty)       //첨부파일 1
                            {
                                if (FTP_Save_File(listFtpFile, txtMsrMachineID.Text))
                                {
                                    if (!txtAttFile1.Text.Equals(string.Empty)) { txtAttFile1.Tag = "/ImageData/Measure/" + txtMsrMachineID.Text; }
                                    if (!txtAttFile2.Text.Equals(string.Empty)) { txtAttFile2.Tag = "/ImageData/Measure/" + txtMsrMachineID.Text; }
                                    if (!txtAttFile3.Text.Equals(string.Empty)) { txtAttFile3.Tag = "/ImageData/Measure/" + txtMsrMachineID.Text; }

                                    AttachYesNo = true;
                                }
                                else
                                { MessageBox.Show("데이터 수정이 완료되었지만, 첨부문서 등록에 실패하였습니다."); }

                                if (AttachYesNo == true) { UpdateFtpPathAndNameData(txtMsrMachineID.Text); }      //첨부문서 정보 DB 업데이트.
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

            return flag;
        }

        //데이터확인
        private bool CheckData()
        {
            bool flag = true;

            if (chkBuyDate.IsChecked == true)
            {
                if (dtpMsrMachineBuyDate.SelectedDate == null)
                {
                    MessageBox.Show("구입일 체크를 하면 날짜를 선택해주셔야 합니다.");
                    flag = false;
                    return flag;
                }
            }

            if (txtCycle.Text.Length > 2)
            {
                MessageBox.Show("검교정주기(월)은 2자리 내로 입력해주셔야 합니다. ex: 12");
                flag = false;
                return flag;
            }






            return flag;
        }

        private void FileChoice(object sender, RoutedEventArgs e)
        {
            OpenFileAndSetting(sender, e);
        }

        private void DeleteFile(object sender, RoutedEventArgs e)
        {
            DeleteFileAndSetting(sender, e);
        }

        private void FileDownLoad(object sender, RoutedEventArgs e)
        {
            DownloadFileAndSetting(sender, e, txtMsrMachineID.Text);
        }

        /// <summary>
        /// 파일첨부
        /// </summary>
        void OpenFileAndSetting(object sender, RoutedEventArgs e)
        {
            // (버튼)sender 마다 tag를 달자.
            int ClickPoint = Convert.ToInt32((sender as Button).Tag);
            Microsoft.Win32.OpenFileDialog OFdlg = new Microsoft.Win32.OpenFileDialog();
            string[] strTemp = null;
            OFdlg.DefaultExt = ".jpg";
            OFdlg.Filter = "Image files (*.jpg, *.jpeg, *.jpe, *.jfif, *.png) | *.jpg; *.jpeg; *.jpe; *.jfif; *.png | All Files|*.*";

            Nullable<bool> result = OFdlg.ShowDialog();
            if (result == true)
            {
                if (ClickPoint == 1) { FullPath1 = OFdlg.FileName; }  //긴 경로(FULL 사이즈)
                if (ClickPoint == 2) { FullPath2 = OFdlg.FileName; }
                if (ClickPoint == 3) { FullPath3 = OFdlg.FileName; }

                string AttachFileName = OFdlg.SafeFileName;  //명.
                string AttachFilePath = string.Empty;       // 경로

                if (ClickPoint == 1) { AttachFilePath = FullPath1.Replace(AttachFileName, ""); }
                if (ClickPoint == 2) { AttachFilePath = FullPath2.Replace(AttachFileName, ""); }
                if (ClickPoint == 3) { AttachFilePath = FullPath3.Replace(AttachFileName, ""); }

                StreamReader sr = new StreamReader(OFdlg.FileName);
                long File_size = sr.BaseStream.Length;
                if (sr.BaseStream.Length > (2048 * 1000))
                {
                    // 업로드 파일 사이즈범위 초과
                    MessageBox.Show("이미지의 파일사이즈가 2M byte를 초과하였습니다.");
                    sr.Close();
                    return;
                }
                if (ClickPoint == 1)
                {
                    txtAttFile1.Text = AttachFileName;
                    txtAttFile1.Tag = AttachFilePath.ToString();
                }
                else if (ClickPoint == 2)
                {
                    txtAttFile2.Text = AttachFileName;
                    txtAttFile2.Tag = AttachFilePath.ToString();
                }
                else if (ClickPoint == 3)
                {
                    txtAttFile3.Text = AttachFileName;
                    txtAttFile3.Tag = AttachFilePath.ToString();
                }

                strTemp = new string[] { AttachFileName, AttachFilePath.ToString() };
                listFtpFile.Add(strTemp);
            }
        }

        /// <summary>
        /// 파일삭제
        /// </summary>
        private void DeleteFileAndSetting(object sender, RoutedEventArgs e)
        {
            MessageBoxResult msgresult = MessageBox.Show("파일을 삭제 하시겠습니까?", "삭제 확인", MessageBoxButton.YesNo);
            if (msgresult == MessageBoxResult.Yes)
            {
                int ClickPoint = Convert.ToInt32((sender as Button).Tag);

                if ((ClickPoint == 1) && (txtAttFile1.Tag.ToString() != string.Empty))
                {
                    txtAttFile1.Text = string.Empty;
                    txtAttFile1.Tag = string.Empty;
                }
                if ((ClickPoint == 2) && (txtAttFile2.Tag.ToString() != string.Empty))
                {
                    txtAttFile2.Text = string.Empty;
                    txtAttFile2.Tag = string.Empty;
                }
                if ((ClickPoint == 3) && (txtAttFile3.Tag.ToString() != string.Empty))
                {
                    txtAttFile3.Text = string.Empty;
                    txtAttFile3.Tag = string.Empty;
                }
            }
        }

        /// <summary>
        /// 파일 다운로드
        /// </summary>
        private void DownloadFileAndSetting(object sender, RoutedEventArgs e, string FolderName)
        {
            try
            {
                MessageBoxResult msgresult = MessageBox.Show("파일을 보시겠습니까?", "보기 확인", MessageBoxButton.YesNo);
                if (msgresult == MessageBoxResult.Yes)
                {
                    //버튼 태그값.
                    int ClickPoint = Convert.ToInt32((sender as Button).Tag);
                    if ((ClickPoint == 1) && (txtAttFile1.Tag.ToString() == string.Empty))
                    {
                        MessageBox.Show("파일이 없습니다.");
                        return;
                    }
                    if ((ClickPoint == 2) && (txtAttFile2.Tag.ToString() == string.Empty))
                    {
                        MessageBox.Show("파일이 없습니다.");
                        return;
                    }
                    if ((ClickPoint == 3) && (txtAttFile3.Tag.ToString() == string.Empty))
                    {
                        MessageBox.Show("파일이 없습니다.");
                        return;
                    }

                    var ViewReceiver = dgdMain.SelectedItem as Win_Qul_MeasureMachine_U_CodeView;
                    if (ViewReceiver != null)
                    {
                        if (ClickPoint == 1)
                        {
                            FTP_DownLoadFile(ViewReceiver.AttPath1, ViewReceiver.MsrMachineID, ViewReceiver.AttFile1);
                        }
                        else if (ClickPoint == 2)
                        {
                            FTP_DownLoadFile(ViewReceiver.AttPath2, ViewReceiver.MsrMachineID, ViewReceiver.AttFile2);
                        }
                        else if (ClickPoint == 3)
                        {
                            FTP_DownLoadFile(ViewReceiver.AttPath3, ViewReceiver.MsrMachineID, ViewReceiver.AttFile3);
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

        //다운로드
        private void FTP_DownLoadFile(string Path, string FolderName, string ImageName)
        {
            try
            {
                _ftp = new FTP_EX(FTP_ADDRESS, FTP_ID, FTP_PASS);

                string[] fileListSimple;
                string[] fileListDetail;

                fileListSimple = _ftp.directoryListSimple("", Encoding.UTF8);

                bool ExistFile = false;

                ExistFile = FolderInfoAndFlag(fileListSimple, FolderName);

                if (ExistFile)
                {
                    ExistFile = false;
                    fileListDetail = _ftp.directoryListSimple(FolderName, Encoding.UTF8);

                    ExistFile = FileInfoAndFlag(fileListDetail, ImageName);

                    if (ExistFile)
                    {
                        string str_remotepath = string.Empty;
                        string str_localpath = string.Empty;

                        str_remotepath = FTP_ADDRESS + '/' + FolderName + '/' + ImageName;
                        str_localpath = LOCAL_DOWN_PATH + "\\" + ImageName;

                        DirectoryInfo DI = new DirectoryInfo(LOCAL_DOWN_PATH);
                        if (DI.Exists)
                        {
                            DI.Create();
                        }

                        FileInfo file = new FileInfo(str_localpath);
                        if (file.Exists)
                        {
                            file.Delete();
                        }

                        _ftp.download(str_remotepath.Substring(str_remotepath.Substring
                            (0, str_remotepath.LastIndexOf("/")).LastIndexOf("/")), str_localpath);

                        ProcessStartInfo proc = new ProcessStartInfo(str_localpath);
                        proc.UseShellExecute = true;
                        Process.Start(proc);
                    }
                    else
                    {
                        MessageBox.Show("파일을 찾을 수 없습니다.");
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

        /// <summary>
        /// 해당영역에 파일 있는지 확인
        /// </summary>
        bool FileInfoAndFlag(string[] strFileList, string FileName)
        {
            bool flag = false;
            foreach (string FileList in strFileList)
            {
                if (FileList == FileName)
                {
                    flag = true;
                    break;
                }
            }
            return flag;
        }


        // 파일 저장하기.
        private bool FTP_Save_File(List<string[]> listStrArrayFileInfo, string MakeFolderName)
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
                //MIL 폴더에 InspectionID로 저장
                if (_ftp.createDirectory(MakeFolderName) == false)
                {
                    MessageBox.Show("업로드를 위한 폴더를 생성할 수 없습니다.");
                    return false;
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
                return false;
            }
            return true;
        }

        private bool UpdateFtpPathAndNameData(string ID)
        {
            bool flag = false;
            try
            {
                Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                sqlParameter.Clear();
                sqlParameter.Add("sMsrMachineID", ID);
                sqlParameter.Add("AttPath1", txtAttFile1.Tag != null ? txtAttFile1.Tag.ToString() : ""); ;
                sqlParameter.Add("AttFile1", txtAttFile1.Text);
                sqlParameter.Add("AttPath2", txtAttFile2.Tag != null ? txtAttFile2.Tag.ToString() : "");
                sqlParameter.Add("AttFile2", txtAttFile2.Text);
                sqlParameter.Add("AttPath3", txtAttFile3.Tag != null ? txtAttFile3.Tag.ToString() : "");
                sqlParameter.Add("AttFile3", txtAttFile3.Text);

                string[] result = DataStore.Instance.ExecuteProcedure("xp_MeasureMachine_uMeasureMachine_FTP", sqlParameter, false);

                if (result[0].Equals("success"))
                {
                    //MessageBox.Show("성공 *^^*");
                    flag = true;
                }
                else
                {
                    MessageBox.Show("삭제 실패 , 내용 : " + result[1]);
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

        #region 프린트 관련

        //검교정 계획
        private void BtnPrint_Plan_Click(object sender, RoutedEventArgs e)
        {
            ContextMenu menu = btnPrint_Plan.ContextMenu;
            menu.StaysOpen = true;
            menu.IsOpen = true;
        }

        //검교정 등록
        private void BtnPrint_Click(object sender, RoutedEventArgs e)
        {
            ContextMenu menu = btnPrint.ContextMenu;
            menu.StaysOpen = true;
            menu.IsOpen = true;
        }

        private void menuSeeAhead_Click(object sender, RoutedEventArgs e)
        {
            int NumFlag = Convert.ToInt32((sender as MenuItem).Tag);

            if (dgdMain.Items.Count < 1)
            {
                MessageBox.Show("먼저 검색해 주세요.");
                return;
            }

            PrintWorkAll printWorkAll = new PrintWorkAll();

            //1이면 계획 2면 등록대장
            if (NumFlag == 1)
            {
                printWorkAll.PrintWorkMeasureMachinePlan(true, ovcMeasureMachine);
            }
            else
            {
                printWorkAll.PrintWorkMeasureMachineRecordDocument(true, ovcMeasureMachine);
            }
        }

        private void menuRightPrint_Click(object sender, RoutedEventArgs e)
        {
            int NumFlag = Convert.ToInt32((sender as MenuItem).Tag);

            if (dgdMain.Items.Count < 1)
            {
                MessageBox.Show("먼저 검색해 주세요.");
                return;
            }

            PrintWorkAll printWorkAll = new PrintWorkAll();

            //1이면 계획 2면 등록대장
            if (NumFlag == 1)
            {
                printWorkAll.PrintWorkMeasureMachinePlan(false, ovcMeasureMachine);
            }
            else
            {
                printWorkAll.PrintWorkMeasureMachineRecordDocument(false, ovcMeasureMachine);
            }
        }

        private void menuClose_Click(object sender, RoutedEventArgs e)
        {
            ContextMenu menu = btnPrint.ContextMenu;
            menu.StaysOpen = false;
            menu.IsOpen = false;
        }




        #endregion

        #region 키보드 포커스 변환관련 
        // 관리번호에서 계측기 명으로 이동.
        private void txtMsrMachineMgrNo_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                txtMsrMachineName.Focus();
            }
        }
        // 계측기명에서 계측기번호로 이동.
        private void txtMsrMachineName_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                txtMsrMachineNo.Focus();
            }
        }
        // 계측기번호에서 계측기규격으로 이동
        private void txtMsrMachineNo_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                txtMsrMachineSpec.Focus();
            }
        }
        // 계측기규격에서 계측기범위로 이동
        private void txtMsrMachineSpec_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                txtMsrMachineRange.Focus();
            }
        }
        // 계측기범위에서 계측기업체로 이동
        private void txtMsrMachineRange_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                txtMsrMachineMsrBuyCustom.Focus();
            }
        }
        // 계측기업체에서 검교정업체로 이동. (구입일은 오늘날짜 자동세팅 패스)
        private void txtMsrMachineMsrBuyCustom_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                txtMsrMachineMsrCustom.Focus();
            }
        }
        // 검교정업체에서 구입가격으로 이동
        private void txtMsrMachineMsrCustom_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                txtMsrMachinePrice.Focus();
            }
        }
        // 구입가격은 숫자만 입력받을 수 있도록.
        private void txtMsrMachinePrice_PreviewTextInput(object sender, TextCompositionEventArgs e)
        {
            Lib.Instance.CheckIsNumericOnly((TextBox)sender, e);
        }
        // 구입가격에서 사용팀으로 이동
        private void txtMsrMachinePrice_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                txtMsrMachineUseTeam.Focus();
            }
        }
        // 사용팀에서 검교정주기(월) 로 이동
        private void txtMsrMachineUseTeam_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                txtCycle.Focus();
            }
        }
        // 검교정주기에서 담당자로 이동
        private void txtCycle_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                txtMsrMachinePerson.Focus();
            }
        }
        // 담당자에서 비고로 이동
        private void txtMsrMachinePerson_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                txtComments.Focus();

            }
        }
        // 비고에서 멀티라인 사용 
        private void txtComments_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                txtComments.AcceptsReturn = true;

            }
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
}

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
using System.Windows.Media.Imaging;
using WizMes_WellMade.PopUP;

namespace WizMes_WellMade
{
    /// <summary>
    /// Win_Qul_Measure_U.xaml에 대한 상호 작용 논리
    /// </summary>
    public partial class Win_Qul_Measure_U : UserControl
    {
        Lib lib = new Lib();
        PrintWorkAll PrintWorkAll = new PrintWorkAll();
        string strFlag = string.Empty;
        int selectedIndex = 0;
        ObservableCollection<Win_Qul_Measure_U_CodeView> ovcMeasure
            = new ObservableCollection<Win_Qul_Measure_U_CodeView>();

        // FTP 활용모음.
        List<string[]> listFtpFile = new List<string[]>();
        string strImagePath = string.Empty;
        string strFullPath = string.Empty;

        string strAttPath1 = string.Empty;
        string FullPath1 = string.Empty;

        private FTP_EX _ftp = null;
        //string FTP_ADDRESS = "ftp://wizis.iptime.org/ImageData/Correct";
        string FTP_ADDRESS = "ftp://" + LoadINI.FileSvr + ":" + LoadINI.FTPPort + LoadINI.FtpImagePath + "/Measure";
        //string FTP_ADDRESS = "ftp://222.104.222.145:25000/ImageData/Correct";
        //string FTP_ADDRESS = "ftp://192.168.0.95/ImageData/Correct";
        private const string FTP_ID = "wizuser";
        private const string FTP_PASS = "wiz9999";
        private const string LOCAL_DOWN_PATH = "C:\\Temp";

        public Win_Qul_Measure_U()
        {
            InitializeComponent();
        }

        private void UserControl_Loaded(object sender, RoutedEventArgs e)
        {
            Lib.Instance.UiLoading(this);
            chkDate.IsChecked = true;
            btnToday_Click(null, null);
            ControlVisibleAndEnable_SC();
        }

        //일자
        private void LblDate_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            if (chkDate.IsChecked == true) { chkDate.IsChecked = false; }
            else { chkDate.IsChecked = true; }
        }

        //일자 체크시
        private void ChkDate_Checked(object sender, RoutedEventArgs e)
        {
            dtpSDate.IsEnabled = true;
            dtpEDate.IsEnabled = true;
        }

        //일자 체크해제시
        private void ChkDate_Unchecked(object sender, RoutedEventArgs e)
        {
            dtpSDate.IsEnabled = false;
            dtpEDate.IsEnabled = false;
        }

        //전일
        private void btnYesterDay_Click(object sender, RoutedEventArgs e)
        {
            DateTime[] SearchDate = Lib.Instance.BringLastDayDateTimeContinue(dtpEDate.SelectedDate.Value);

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
            btnPfMsrMachineNameSrh.IsEnabled = true;
            chkMsrMachineNameSrh.Focus();
        }

        //계측기명
        private void ChkMsrMachineNameSrh_Unchecked(object sender, RoutedEventArgs e)
        {
            txtMsrMachineNameSrh.IsEnabled = false;
            btnPfMsrMachineNameSrh.IsEnabled = false;
        }

        //계측기명
        private void TxtMsrMachineNameSrh_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                MainWindow.pf.ReturnCode(txtMsrMachineNameSrh, (int)Defind_CodeFind.DCF_QULMSRMACHINE, "");
            }
        }

        //계측기명
        private void BtnPfMsrMachineNameSrh_Click(object sender, RoutedEventArgs e)
        {
            MainWindow.pf.ReturnCode(txtMsrMachineNameSrh, (int)Defind_CodeFind.DCF_QULMSRMACHINE, "");
        }

        //계측기업체
        private void LblMsrMachineMsrBuyCustomSrh_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            if (chkMsrMachineMsrBuyCustomSrh.IsChecked == true) { chkMsrMachineMsrBuyCustomSrh.IsChecked = false; }
            else { chkMsrMachineMsrBuyCustomSrh.IsChecked = true; }
        }

        //계측기업체
        private void ChkMsrMachineMsrBuyCustomSrh_Checked(object sender, RoutedEventArgs e)
        {
            txtMsrMachineMsrBuyCustomSrh.IsEnabled = true;
            txtMsrMachineMsrBuyCustomSrh.Focus();
        }

        //계측기업체
        private void ChkMsrMachineMsrBuyCustomSrh_Unchecked(object sender, RoutedEventArgs e)
        {
            txtMsrMachineMsrBuyCustomSrh.IsEnabled = false;
        }

        //검교정업체
        private void LblMsrMachineMsrCustomSrh_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            if (chkMsrMachineMsrCustomSrh.IsChecked == true) { chkMsrMachineMsrCustomSrh.IsChecked = false; }
            else { chkMsrMachineMsrCustomSrh.IsChecked = true; }
        }

        //검교정업체
        private void ChkMsrMachineMsrCustomSrh_Checked(object sender, RoutedEventArgs e)
        {
            txtMsrMachineMsrCustomSrh.IsEnabled = true;
            txtMsrMachineMsrCustomSrh.Focus();
        }

        //검교정업체
        private void ChkMsrMachineMsrCustomSrh_Unchecked(object sender, RoutedEventArgs e)
        {
            txtMsrMachineMsrCustomSrh.IsEnabled = false;
        }

        //확인버튼
        private void BtnPrint_Plan_Click(object sender, RoutedEventArgs e)
        {
            ContextMenu menu = btnPrint_Plan.ContextMenu;
            menu.StaysOpen = true;
            menu.IsOpen = true;
        }

        //확인버튼 내부 미리보기
        private void menuSeeAhead_Click(object sender, RoutedEventArgs e)
        {
            DataRow dataRow = null;

            if (txtMsrMachineName.Tag != null && !txtMsrMachineName.Tag.ToString().Equals(string.Empty))
            {
                dataRow = GetMeasureMachineOne(txtMsrMachineName.Tag.ToString());
            }

            if (dataRow != null)
            {
                PrintWorkAll.PrintWorkMeasureDocument(true, ovcMeasure, dataRow);
            }
        }

        //확인버튼 내부 바로인쇄
        private void menuRightPrint_Click(object sender, RoutedEventArgs e)
        {
            DataRow dataRow = null;

            if (txtMsrMachineName.Tag != null && !txtMsrMachineName.Tag.ToString().Equals(string.Empty))
            {
                dataRow = GetMeasureMachineOne(txtMsrMachineName.Tag.ToString());
            }

            if (dataRow != null)
            {
                PrintWorkAll.PrintWorkMeasureDocument(false, ovcMeasure, dataRow);
            }
        }

        //확인버튼 내부 닫기
        private void menuClose_Click(object sender, RoutedEventArgs e)
        {
            ContextMenu menu = btnPrint_Plan.ContextMenu;
            menu.StaysOpen = false;
            menu.IsOpen = false;
        }

        /// <summary>
        /// 추가,수정 시 동작 모음
        /// </summary>
        private void ControlVisibleAndEnable_AU()
        {
            Lib.Instance.UiButtonEnableChange_SCControl(this);
            dgdMain.IsHitTestVisible = false;
            grdInput.IsHitTestVisible = true;

            txtProofID.IsHitTestVisible = false;
            txtMsrMachineName.IsHitTestVisible = true;
            txtMsrMachineNo.IsHitTestVisible = true;
            txtMsrMachineMsrBuyCustom.IsHitTestVisible = true;
            txtCycle.IsHitTestVisible = true;
            txtMsrMachineMsrCustom.IsHitTestVisible = true;
            dtpProofDate.IsHitTestVisible = true;
            dtpNextProofDate.IsHitTestVisible = true;
            txtChangePoint.IsHitTestVisible = true;
            txtComments.IsHitTestVisible = true;
            txtPic.IsHitTestVisible = false;

            btnUpFile.IsHitTestVisible = true;
            txtUpFile.IsHitTestVisible = true;
            btnDelPic.IsHitTestVisible = true;

            txtAttFile1.IsHitTestVisible = false;
            txtAttFile2.IsHitTestVisible = false;
            txtAttFile3.IsHitTestVisible = false;

            btnAttFileSee1.IsHitTestVisible = true;
            btnAttFileSee2.IsHitTestVisible = true;
            btnAttFileSee3.IsHitTestVisible = true;




        }

        /// <summary>
        /// 저장,취소 시 동작 모음
        /// </summary>
        private void ControlVisibleAndEnable_SC()
        {
            Lib.Instance.UiButtonEnableChange_IUControl(this);
            dgdMain.IsHitTestVisible = true;

            txtProofID.IsHitTestVisible = false;
            txtMsrMachineName.IsHitTestVisible = false;
            txtMsrMachineNo.IsHitTestVisible = false;
            txtMsrMachineMsrBuyCustom.IsHitTestVisible = false;
            txtCycle.IsHitTestVisible = false;
            txtMsrMachineMsrCustom.IsHitTestVisible = false;
            dtpProofDate.IsHitTestVisible = false;
            dtpNextProofDate.IsHitTestVisible = false;
            txtChangePoint.IsHitTestVisible = false;
            txtComments.IsHitTestVisible = false;
            txtPic.IsHitTestVisible = false;

            btnUpFile.IsHitTestVisible = false;
            txtUpFile.IsHitTestVisible = false;
            btnDelPic.IsHitTestVisible = false;

            txtAttFile1.IsHitTestVisible = false;
            txtAttFile2.IsHitTestVisible = false;
            txtAttFile3.IsHitTestVisible = false;

            btnAttFileSee1.IsHitTestVisible = true;
            btnAttFileSee2.IsHitTestVisible = true;
            btnAttFileSee3.IsHitTestVisible = true;

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

            txtProofID.IsHitTestVisible = false;
            dtpProofDate.SelectedDate = DateTime.Now;
            txtMsrMachineName.Focus();

            //FTP이미지 리스트 비워주고 시작
            listFtpFile.Clear();
            //띄워진 이미지 비워주기
            imgSajin.Source = null;


        }

        //수정
        private void BtnUpdate_Click(object sender, RoutedEventArgs e)
        {
            var uMeasureMachine = dgdMain.SelectedItem as Win_Qul_Measure_U_CodeView;
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
                var dMeasureMachine = dgdMain.SelectedItem as Win_Qul_Measure_U_CodeView;
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

                        if (Procedure.Instance.DeleteData(dMeasureMachine.ProofID, "ProofID"
                            , "xp_Qul_dMeasureMachineProof"))
                        {
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

        //검색
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
                selectedIndex = 0;

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

                FillGrid();
                if (dgdMain.Items.Count > 0)
                {
                    if (strFlag == "I")
                    {
                        dgdMain.SelectedIndex = dgdMain.Items.Count - 1;
                    }
                    else
                    {
                        dgdMain.SelectedIndex = selectedIndex;
                    }
                    dgdMain.Focus();
                }
                strFlag = string.Empty;
            }
        }

        //취소
        private void BtnCancel_Click(object sender, RoutedEventArgs e)
        {
            ControlVisibleAndEnable_SC();
            strFlag = string.Empty;
            FillGrid();
            dgdMain.SelectedIndex = selectedIndex;
            dgdMain.Focus();
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

            try
            {
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
            }
            catch (Exception ex)
            {
                MessageBox.Show("오류 발생, 오류 내용 : " + ex.ToString());
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

        // 조회
        private void FillGrid()
        {
            try
            {
                ovcMeasure.Clear();

                //Tag 가 있는것만 비교 후 대입
                string strMachineID = string.Empty;

                if (chkMsrMachineNameSrh.IsChecked == true && txtMsrMachineNameSrh.Tag != null)
                {
                    strMachineID = txtMsrMachineNameSrh.Tag.ToString();
                }

                Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                sqlParameter.Add("nchkDate", chkDate.IsChecked == true ? 1 : 0);
                sqlParameter.Add("FromDate", chkDate.IsChecked == true ? dtpSDate.SelectedDate.Value.ToString("yyyyMMdd") : "");
                sqlParameter.Add("ToDate", chkDate.IsChecked == true ? dtpEDate.SelectedDate.Value.ToString("yyyyMMdd") : "");
                sqlParameter.Add("nchkMachineID", chkMsrMachineNameSrh.IsChecked == true ? 1 : 0);
                sqlParameter.Add("MachineID", strMachineID);
                sqlParameter.Add("nchkMsrCustom", chkMsrMachineMsrBuyCustomSrh.IsChecked == true ? 1 : 0);
                sqlParameter.Add("MsrMachineMsrCustom", chkMsrMachineMsrBuyCustomSrh.IsChecked == true ?
                    txtMsrMachineMsrBuyCustomSrh.Text : "");
                sqlParameter.Add("nchkProofCustom", chkMsrMachineMsrCustomSrh.IsChecked == true ? 1 : 0);
                sqlParameter.Add("ProofCustom", chkMsrMachineMsrCustomSrh.IsChecked == true ?
                    txtMsrMachineMsrCustomSrh.Text : "");
                DataSet ds = DataStore.Instance.ProcedureToDataSet("xp_Qul_sMeasureMachineProof", sqlParameter, false);

                if (ds != null && ds.Tables.Count > 0)
                {
                    DataTable dt = ds.Tables[0];
                    if (dt.Rows.Count == 0)
                    {
                        MessageBox.Show("조회된 데이터가 없습니다.");
                        this.DataContext = null;
                    }
                    else
                    {
                        DataRowCollection drc = dt.Rows;
                        int i = 0;

                        foreach (DataRow dr in drc)
                        {
                            i++;
                            var varMeasure = new Win_Qul_Measure_U_CodeView
                            {
                                Num = i,

                                ProofID = dr["ProofID"].ToString(),
                                MsrMachineID = dr["MsrMachineID"].ToString(),
                                MsrMachineMgrNo = dr["MsrMachineMgrNo"].ToString(),
                                MsrMachineMsrBuyCustom = dr["MsrMachineMsrBuyCustom"].ToString(),
                                MsrMachineMsrCustom = dr["MsrMachineMsrCustom"].ToString(),

                                ProofCustom = dr["ProofCustom"].ToString(),
                                ProofCycle = dr["ProofCycle"].ToString(),
                                ProofCycleUnit = dr["ProofCycleUnit"].ToString(),
                                ProofCycleUnitName = dr["ProofCycleUnitName"].ToString(),
                                ProofDate = dr["ProofDate"].ToString(),

                                NextProofDate = dr["NextProofDate"].ToString(),
                                Comments = dr["Comments"].ToString(),
                                MsrMachineName = dr["MsrMachineName"].ToString(),
                                AttPath1 = dr["AttPath1"].ToString(),
                                AttFile1 = dr["AttFile1"].ToString(),

                                AttPath2 = dr["AttPath2"].ToString(),
                                AttFile2 = dr["AttFile2"].ToString(),
                                AttPath3 = dr["AttPath3"].ToString(),
                                AttFile3 = dr["AttFile3"].ToString(),
                                ChangePoint = dr["ChangePoint"].ToString(),

                                MsrMachineUseTeam = dr["MsrMachineUseTeam"].ToString(),
                                MsrMachinePrice = dr["MsrMachinePrice"].ToString(),
                                MsrMachineSpec = dr["MsrMachineSpec"].ToString(), //2021-07-06
                                MsrMachineRange = dr["MsrMachineRange"].ToString(),
                                Upfile = dr["Upfile"].ToString(),

                                UpfilePath = dr["UpfilePath"].ToString()
                            };

                            if (!Lib.Instance.CheckNull(varMeasure.ProofDate).Equals(string.Empty))
                            {
                                varMeasure.ProofDate_CV =
                                    Lib.Instance.StrDateTimeBar(varMeasure.ProofDate);
                            }
                            if (!Lib.Instance.CheckNull(varMeasure.NextProofDate).Equals(string.Empty))
                            {
                                varMeasure.NextProofDate_CV =
                                    Lib.Instance.Left(varMeasure.NextProofDate, 10);
                            }

                            ovcMeasure.Add(varMeasure);
                        }
                        tbkIndexCount.Text = "▶검색결과 : " + i + " 건";
                        dgdMain.ItemsSource = ovcMeasure;
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

        //
        private void DgdMain_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            var Measure = dgdMain.SelectedItem as Win_Qul_Measure_U_CodeView;

            if (Measure != null)
            {
                this.DataContext = Measure;

                txtMsrMachineName.Tag = Measure.MsrMachineID;
            }

            imgSajin.Source = null;
        }

        //계측기명
        private void TxtMsrMachineName_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                MainWindow.pf.ReturnCode(txtMsrMachineName, (int)Defind_CodeFind.DCF_QULMSRMACHINE, "");

                if (txtMsrMachineName.Tag != null && !txtMsrMachineName.Tag.ToString().Equals(string.Empty)
                    && txtMsrMachineName.Text.Length > 0)
                {
                    GetMeasureMachineOne(txtMsrMachineName.Tag.ToString());
                    txtMsrMachineMsrCustom.Focus();
                }
            }
        }

        //계측기명
        private void BtnPfMsrMachineName_Click(object sender, RoutedEventArgs e)
        {
            MainWindow.pf.ReturnCode(txtMsrMachineName, (int)Defind_CodeFind.DCF_QULMSRMACHINE, "");

            if (txtMsrMachineName.Tag != null && !txtMsrMachineName.Tag.ToString().Equals(string.Empty)
                    && txtMsrMachineName.Text.Length > 0)
            {
                GetMeasureMachineOne(txtMsrMachineName.Tag.ToString());
                txtMsrMachineMsrCustom.Focus();
            }
        }

        //
        private DataRow GetMeasureMachineOne(string strMsrMachineNameID)
        {
            DataRow dataRow = null;

            try
            {
                Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                sqlParameter.Clear();
                sqlParameter.Add("MsrMachineID", strMsrMachineNameID);
                DataSet ds = DataStore.Instance.ProcedureToDataSet("xp_MeasureMachine_sMeasureMachineOne", sqlParameter, false);

                if (ds != null && ds.Tables.Count > 0)
                {
                    DataTable dt = ds.Tables[0];

                    if (dt.Rows.Count > 0)
                    {
                        dataRow = dt.Rows[0];
                        txtMsrMachineNo.Text = Lib.Instance.CheckNull(dataRow["MsrMachineMgrNo"]);
                        txtMsrMachineMsrBuyCustom.Text = Lib.Instance.CheckNull(dataRow["MsrMachineMsrBuyCustom"]);
                        txtCycle.Text = Lib.Instance.CheckNull(dataRow["ProofCycle"]);
                        txtMsrMachineMsrCustom.Text = Lib.Instance.CheckNull(dataRow["MsrMachineMsrCustom"]);

                        //계측기등록시 업로드한 이미지도 같이 가져온다.

                        txtAttFile1.Tag = Lib.Instance.CheckNull(dataRow["AttPath1"]);
                        txtAttFile1.Text = Lib.Instance.CheckNull(dataRow["AttFile1"]);
                        txtAttFile2.Text = Lib.Instance.CheckNull(dataRow["AttFile2"]);
                        txtAttFile2.Tag = Lib.Instance.CheckNull(dataRow["AttPath2"]);
                        txtAttFile3.Text = Lib.Instance.CheckNull(dataRow["AttFile3"]);
                        txtAttFile3.Tag = Lib.Instance.CheckNull(dataRow["AttPath3"]);
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

            return dataRow;
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
                    Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                    sqlParameter.Clear();
                    sqlParameter.Add("ProofID", txtProofID.Text);
                    sqlParameter.Add("MsrMachineID", txtMsrMachineName.Tag.ToString());
                    sqlParameter.Add("ProofCustom", txtMsrMachineMsrCustom.Text);
                    sqlParameter.Add("ProofDate", dtpProofDate.SelectedDate.Value.ToString("yyyyMMdd"));
                    sqlParameter.Add("ChangePoint", txtChangePoint.Text);

                    sqlParameter.Add("Comments", txtComments.Text);
                    sqlParameter.Add("UpFile", txtUpFile.Text);
                    sqlParameter.Add("UpFilePath", "");

                    if (strFlag.Equals("I"))
                    {
                        sqlParameter.Add("CreateUserID", MainWindow.CurrentUser);

                        Procedure pro1 = new Procedure();
                        pro1.Name = "xp_Qul_iMeasureMachineProof";
                        pro1.OutputUseYN = "Y";
                        pro1.OutputName = "ProofID";
                        pro1.OutputLength = "10";

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
                                if (kv.key == "ProofID")
                                {
                                    sGetID = kv.value;
                                    flag = true;
                                    break;
                                }
                            }

                            if (flag)
                            {
                                bool AttachYesNo = false;
                                if (txtUpFile.Text != string.Empty)       //첨부파일 1
                                {
                                    if (FTP_Save_File(listFtpFile, sGetID))
                                    {
                                        if (!txtUpFile.Text.Equals(string.Empty)) { txtUpFile.Tag = "/ImageData/Measure/" + sGetID; }

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
                        sqlParameter.Add("UserID", MainWindow.CurrentUser);

                        Procedure pro1 = new Procedure();
                        pro1.Name = "xp_Qul_uMeasureMachineProof";
                        pro1.OutputUseYN = "N";
                        pro1.OutputName = "ProofID";
                        pro1.OutputLength = "10";

                        Prolist.Add(pro1);
                        ListParameter.Add(sqlParameter);

                        string[] Confirm = new string[2];
                        Confirm = DataStore.Instance.ExecuteAllProcedureOutputNew(Prolist, ListParameter);
                        if (Confirm[0] != "success")
                        {
                            MessageBox.Show("[저장실패]\r\n" + Confirm[1].ToString());
                            flag = false;

                            if (flag)
                            {
                                bool AttachYesNo = false;
                                if (txtUpFile.Text != string.Empty)       //첨부파일 1
                                {
                                    if (FTP_Save_File(listFtpFile, txtProofID.Text))
                                    {
                                        if (!txtUpFile.Text.Equals(string.Empty)) { txtUpFile.Tag = "/ImageData/Measure/" + txtProofID.Text; }

                                        AttachYesNo = true;
                                    }
                                    else
                                    { MessageBox.Show("데이터 저장이 완료되었지만, 첨부문서 등록에 실패하였습니다."); }

                                    if (AttachYesNo == true) { UpdateFtpPathAndNameData(txtProofID.Text); }      //첨부문서 정보 DB 업데이트.
                                }
                            }
                        }
                        else
                        {
                            flag = true;
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

            if (txtMsrMachineName.Text == null || txtMsrMachineName.Text.ToString().Equals(""))
            {
                MessageBox.Show("기계명이 선택되지 않았습니다. 선택해주세요");
                flag = false;
                return flag;
            }

            return flag;
        }

        #region FTP

        /// <summary>
        /// 파일첨부
        /// </summary>
        void OpenFileAndSetting(object sender, RoutedEventArgs e)
        {
            // (버튼)sender 마다 tag를 달자.
            Microsoft.Win32.OpenFileDialog OFdlg = new Microsoft.Win32.OpenFileDialog();
            string[] strTemp = null;
            OFdlg.DefaultExt = ".jpg";
            //OFdlg.Filter = "Image files (*.jpg, *.jpeg, *.jpe, *.jfif, *.png) | *.jpg; *.jpeg; *.jpe; *.jfif; *.png | All Files|*.*";

            OFdlg.Filter = MainWindow.OFdlg_Filter;

            Nullable<bool> result = OFdlg.ShowDialog();
            if (result == true)
            {
                FullPath1 = OFdlg.FileName;

                string AttachFileName = OFdlg.SafeFileName;  //명.
                string AttachFilePath = string.Empty;       // 경로

                AttachFilePath = FullPath1.Replace(AttachFileName, "");

                StreamReader sr = new StreamReader(OFdlg.FileName);
                long File_size = sr.BaseStream.Length;
                if (sr.BaseStream.Length > (2048 * 1000))
                {
                    // 업로드 파일 사이즈범위 초과
                    MessageBox.Show("이미지의 파일사이즈가 2M byte를 초과하였습니다.");
                    sr.Close();
                    return;
                }

                txtUpFile.Text = AttachFileName;
                txtUpFile.Tag = AttachFilePath.ToString();

                strTemp = new string[] { AttachFileName, AttachFilePath.ToString() };
                listFtpFile.Add(strTemp);
            }
        }

        //FTP
        private void BtnUpFile_Click(object sender, RoutedEventArgs e)
        {
            OpenFileAndSetting(sender, e);
        }

        //FTP
        private void BtnDelPic_Click(object sender, RoutedEventArgs e)
        {
            MessageBoxResult msgresult = MessageBox.Show("파일을 삭제 하시겠습니까?", "삭제 확인", MessageBoxButton.YesNo);
            if (msgresult == MessageBoxResult.Yes)
            {
                txtUpFile.Text = string.Empty;
                txtUpFile.Tag = string.Empty;
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

        //파일 저장시
        private void ChoiceFileUpload(List<string[]> listStrArrayFileInfo, string MakeFolderName)
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

            if (!txtUpFile.Text.Equals(string.Empty)) { txtUpFile.Tag = "/ImageData/Measure/" + MakeFolderName + "/"; }
        }

        private BitmapImage SetImage(string strAttachPath)
        {
            _ftp = new FTP_EX(FTP_ADDRESS, FTP_ID, FTP_PASS);
            BitmapImage bit = _ftp.DrawingImageByByte(FTP_ADDRESS + strAttachPath + "");
            //image.Source = bit;
            return bit;
        }

        //FTP
        private void BtnAttFileSee1_Click(object sender, RoutedEventArgs e)
        {
            if (txtAttFile1.Tag != null && !txtAttFile1.Tag.ToString().Equals(string.Empty))
            {
                string strImage = "/" + txtAttFile1.Tag.ToString().Substring(txtAttFile1.Tag.ToString().Length - 5, 5) + "/" + txtAttFile1.Text;
                imgSajin.Source = SetImage(strImage);
            }
        }

        //FTP
        private void BtnAttFileSee2_Click(object sender, RoutedEventArgs e)
        {
            if (txtAttFile2.Tag != null && !txtAttFile2.Tag.ToString().Equals(string.Empty))
            {
                string strImage = "/" + txtAttFile2.Tag.ToString().Substring(txtAttFile2.Tag.ToString().Length - 5, 5) + "/" + txtAttFile2.Text;
                imgSajin.Source = SetImage(strImage);
            }
        }

        //FTP
        private void BtnAttFileSee3_Click(object sender, RoutedEventArgs e)
        {
            if (txtAttFile3.Tag != null && !txtAttFile3.Tag.ToString().Equals(string.Empty))
            {
                string strImage = "/" + txtAttFile3.Tag.ToString().Substring(txtAttFile3.Tag.ToString().Length - 5, 5) + "/" + txtAttFile3.Text;
                imgSajin.Source = SetImage(strImage);
            }
        }

        #endregion


        // 검교정업체에서 변경사항으로 이동
        private void txtMsrMachineMsrCustom_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                txtChangePoint.Focus();
            }
        }
        // 변경사항에서 비고로 이동
        private void txtChangePoint_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                txtComments.Focus();
            }
        }
        // 비고에서 교정성적서로 이동
        private void txtComments_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                txtUpFile.Focus();
            }
        }
        // 교정성적서. >> 첨부파일.
        private void txtUpFile_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                BtnUpFile_Click(null, null);
            }
        }


        //교정성적서 다운로드 클릭 이벤트
        private void BtnDown_Click(object sender, RoutedEventArgs e)
        {
            DownloadFileAndSetting(sender, e, txtProofID.Text);
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
                    if ((txtUpFile.Tag.ToString() == string.Empty))
                    {
                        MessageBox.Show("파일이 없습니다.");
                        return;
                    }

                    var ViewReceiver = dgdMain.SelectedItem as Win_Qul_Measure_U_CodeView;
                    if (ViewReceiver != null)
                    {
                        FTP_DownLoadFile(ViewReceiver.UpfilePath, ViewReceiver.ProofID, ViewReceiver.Upfile);
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
                sqlParameter.Add("ProofID", ID);
                sqlParameter.Add("UpFilePath", txtUpFile.Tag != null ? txtUpFile.Tag.ToString() : ""); ;
                sqlParameter.Add("UpFile", txtUpFile.Text);

                string[] result = DataStore.Instance.ExecuteProcedure("xp_Qul_uMeasureMachineProof_FTP", sqlParameter, false);

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

    }
}

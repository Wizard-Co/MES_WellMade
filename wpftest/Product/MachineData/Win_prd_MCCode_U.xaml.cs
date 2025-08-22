using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Data;
using System.Diagnostics;
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
using WizMes_WellMade.PopUP;
using WizMes_WellMade.PopUp;
using WPF.MDI;

namespace WizMes_WellMade
{
    /// <summary>
    /// Win_prd_MCCode_U.xaml에 대한 상호 작용 논리
    /// </summary>
    public partial class Win_prd_MCCode_U : UserControl
    {
        string stDate = string.Empty;
        string stTime = string.Empty;

        int rowNum = 0;
        string strFlag = string.Empty;
        Win_prd_MCCode_U_CodeView WinMcCode = new Win_prd_MCCode_U_CodeView();
        Win_prd_MCCode_U_MachineMapping_CodeView WinMcCodeMachineMapping = new Win_prd_MCCode_U_MachineMapping_CodeView();
        Win_prd_MCCode_U_Sub_CodeView WinMcCodeSub = new Win_prd_MCCode_U_Sub_CodeView();
        Lib lib = new Lib();

        //파일 수정 진행 위한 flag 3가지
        bool existFtp = false;
        bool AddFtp = false;
        bool delFtp = false;

        #region FTP
        // FTP 활용모음.
        string strImagePath = string.Empty;
        string strFullPath = string.Empty;
        string strDelFileName = string.Empty;

        List<string[]> listFtpFile = new List<string[]>();
        List<string[]> deleteListFtpFile = new List<string[]>(); // 삭제할 파일 리스트

        private FTP_EX _ftp = null;
        private List<UploadFileInfo> _listFileInfo = new List<UploadFileInfo>();

        internal struct UploadFileInfo          //FTP.
        {
            public string Filename { get; set; }
            public FtpFileType Type { get; set; }
            public DateTime LastModifiedTime { get; set; }
            public long Size { get; set; }
            public string Filepath { get; set; }
        }
        internal enum FtpFileType
        {
            None,
            DIR,
            File
        }

        //string FTP_ADDRESS = "ftp://wizis.iptime.org/ImageData/McCode";
        //string FTP_ADDRESS = "ftp://192.168.0.120";
        //string FTP_ADDRESS = "ftp://wizis.iptime.org/ImageData/Draw";
        //string FTP_ADDRESS = "ftp://222.104.222.145:25000/ImageData/McCode";
        //string FTP_ADDRESS = "ftp://192.168.0.95/ImageData/McCode";
        string FTP_ADDRESS = "ftp://" + LoadINI.FileSvr + ":" + LoadINI.FTPPort + LoadINI.FtpImagePath + "/McCode";
        private const string FTP_ID = "wizuser";
        private const string FTP_PASS = "wiz9999";
        private const string LOCAL_DOWN_PATH = "C:\\Temp";

        #endregion

        public Win_prd_MCCode_U()
        {
            InitializeComponent();
        }

        private void UserControl_Loaded(object sender, RoutedEventArgs e)
        {
            stDate = DateTime.Now.ToString("yyyyMMdd");
            stTime = DateTime.Now.ToString("HHmm");

            DataStore.Instance.InsertLogByFormS(this.GetType().Name, stDate, stTime, "S");

            Lib.Instance.UiLoading(sender);
            SetComboBox();
            _ftp = new FTP_EX(FTP_ADDRESS, FTP_ID, FTP_PASS);
        }

        #region 콤보박스
        //콤보박스 만들기
        private void SetComboBox()
        {
            //ObservableCollection<CodeView> ovcProcessAutoMC = ComboBoxUtil.Instance.GetProcessByAutoMC();
            //this.cboProcess.ItemsSource = ovcProcessAutoMC;
            //this.cboProcess.DisplayMemberPath = "code_name";
            //this.cboProcess.SelectedValuePath = "code_id";

            //if (cboProcess.ItemsSource != null)
            //{
            //    cboProcess.SelectedIndex = 1;
            //}

            //ObservableCollection<CodeView> ovcMachineAutoMC = ComboBoxUtil.Instance.GetMachineCodeByAutoMC(cboProcess.SelectedValue.ToString());
            //this.cboMachine.ItemsSource = ovcMachineAutoMC;
            //this.cboMachine.DisplayMemberPath = "code_name";
            //this.cboMachine.SelectedValuePath = "code_id";

            ObservableCollection<CodeView> ovcLicense = ComboBoxUtil.Instance.Gf_DB_CM_GetComCodeDataset(null, "HRLICE", "Y", "");
            this.cboLicense.ItemsSource = ovcLicense;
            this.cboLicense.DisplayMemberPath = "code_name";
            this.cboLicense.SelectedValuePath = "code_id";


        }
        #endregion

        #region Header 검색조건

        private void lblMcCodeSrh_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            if (chkMcCodeSrh.IsChecked == true)
            {
                chkMcCodeSrh.IsChecked = false;
            }
            else
            {
                chkMcCodeSrh.IsChecked = true;
            }
        }

        private void chkMcCodeSrh_Checked(object sender, RoutedEventArgs e)
        {
            chkMcCodeSrh.IsChecked = true;
            txtMcCodeSrh.IsEnabled = true;
        }

        private void chkMcCodeSrh_Unchecked(object sender, RoutedEventArgs e)
        {
            chkMcCodeSrh.IsChecked = false;
            txtMcCodeSrh.IsEnabled = false;
        }

        #endregion

        #region 우측 상단 버튼 모음 
        /// <summary>
        /// 수정,추가 저장 후
        /// </summary>
        private void CompleteCancelMode()
        {
            Lib.Instance.UiButtonEnableChange_IUControl(this);
            grdInputOne.IsEnabled = false;
            grdInputThree.IsEnabled = false;
            grdInputFive.IsEnabled = false;

            btnImgDel.IsEnabled = false; //대표이미지 삭제 비활성화
            grdMachineMapping.IsEnabled = false;

            btnImgDownload.IsEnabled = true;     // 대표이미지 다운로드 활성화

            if (!txtImage.Text.Trim().Equals(""))
            {
                btnImgDownload.IsEnabled = true;        //활성화
            }
            else
            {
                btnImgDownload.IsEnabled = false;       //비활성화
            }

        }

        /// <summary>
        /// 수정,추가 진행 중
        /// </summary>
        private void SaveUpdateMode()
        {
            //grdInput.IsEnabled = true;
            Lib.Instance.UiButtonEnableChange_SCControl(this);
            grdInputOne.IsEnabled = true;
            //grdInputTwo.IsEnabled = true;
            grdInputThree.IsEnabled = true;
            //grdInputFour.IsEnabled = true;
            grdInputFive.IsEnabled = true;

            grdMachineMapping.IsEnabled = true;
            btnImgDel.IsEnabled = true; //대표이미지 삭제 활성화

            btnImgDownload.IsEnabled = false; // 대표이미지 다운로드 
        }

        //추가
        private void btnAdd_Click(object sender, RoutedEventArgs e)
        {
            SaveUpdateMode();
            strFlag = "I";

            lblMsg.Visibility = Visibility.Visible;
            tbkMsg.Text = "자료 입력 중";
            rowNum = dgdMain.SelectedIndex;

            dgdMain.IsHitTestVisible = false;
            dgdUseMcPart.IsHitTestVisible = true;
            this.DataContext = null;
            dgdMachineMapping.Items.Clear();
            dgdUseMcPart.Items.Clear();
            imgSetting.Refresh();

            dtpBuyDate.SelectedDate = DateTime.Today;
            txtPerson.Text = MainWindow.CurrentPerson;
            txtPerson.Tag = MainWindow.CurrentPersonID;

            chkNotUse.IsChecked = false; //2021-11-15
            // 이미지 초기화
            imgSetting.Source = null;


        }

        //수정
        private void btnUpdate_Click(object sender, RoutedEventArgs e)
        {
            WinMcCode = dgdMain.SelectedItem as Win_prd_MCCode_U_CodeView;

            if (WinMcCode != null)
            {
                rowNum = dgdMain.SelectedIndex;
                //dgdMain.IsEnabled = false;
                //dgdUseMcPart.IsEnabled = true;
                dgdMain.IsHitTestVisible = false;
                dgdUseMcPart.IsHitTestVisible = true;
                tbkMsg.Text = "자료 수정 중";
                lblMsg.Visibility = Visibility.Visible;
                SaveUpdateMode();
                strFlag = "U";

                if (WinMcCode.ImageFile.Replace(" ", "").Equals(""))
                {
                    existFtp = false;
                }
                else
                {
                    existFtp = true;
                }
            }
        }

        //삭제
        private void btnDelete_Click(object sender, RoutedEventArgs e)
        {
            WinMcCode = dgdMain.SelectedItem as Win_prd_MCCode_U_CodeView;

            if (WinMcCode == null)
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

                    if (DeleteData(WinMcCode.mcid))
                    {
                        MappingDeleteData(WinMcCode.mcid);

                        if (txtImage.Tag != null && !txtImage.Text.Equals(string.Empty))
                        {
                            FTP_RemoveDir(WinMcCode.mcid);
                        }

                        imgSetting.Source = null;
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
                try
                {
                    rowNum = 0;
                    using (Loading lw = new Loading(FillGrid))
                    {
                        lw.ShowDialog();
                        if (dgdMain.Items.Count <= 0)
                        {
                            MessageBox.Show("조회된 내용이 없습니다.");
                        }
                        else
                        {
                            dgdMain.SelectedIndex = rowNum;
                        }

                        btnSearch.IsEnabled = true;
                    }
                }
                catch (Exception ee)
                {
                    MessageBox.Show("예외처리 - " + ee.ToString());
                }

            }), System.Windows.Threading.DispatcherPriority.Background);



        }

        //저장
        private void btnSave_Click(object sender, RoutedEventArgs e)
        {
            if (SaveData(strFlag, txtMCID.Text))
            {
                CompleteCancelMode();
                lblMsg.Visibility = Visibility.Hidden;
                //rowNum = 0;
                //dgdMain.IsEnabled = true;
                dgdMain.IsHitTestVisible = true;
                existFtp = false;
                delFtp = false;
                AddFtp = false;
                strFlag = string.Empty;
                strImagePath = string.Empty;
                strDelFileName = string.Empty;
                re_Search(rowNum);

            }
        }

        //취소
        private void btnCancel_Click(object sender, RoutedEventArgs e)
        {
            CompleteCancelMode();

            if (!strFlag.Equals(string.Empty))
            {
                re_Search(rowNum);
            }

            strFlag = string.Empty;
            strImagePath = string.Empty;
            strDelFileName = string.Empty;
            //dgdMain.IsEnabled = true;
            dgdMain.IsHitTestVisible = true;
            existFtp = false;
            txtImage.Text = "";
        }

        //엑셀
        private void btnExcel_Click(object sender, RoutedEventArgs e)
        {
            DataTable dt = null;
            string Name = string.Empty;

            string[] lst = new string[4];
            lst[0] = "기계설비";
            lst[1] = "기계설비 사용부품";
            lst[2] = dgdMain.Name;
            lst[3] = dgdUseMcPart.Name;

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
                else if (ExpExc.choice.Equals(dgdUseMcPart.Name))
                {
                    if (ExpExc.Check.Equals("Y"))
                        dt = Lib.Instance.DataGridToDTinHidden(dgdUseMcPart);
                    else
                        dt = Lib.Instance.DataGirdToDataTable(dgdUseMcPart);

                    Name = dgdUseMcPart.Name;
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
                dgdMain.SelectedIndex = selectedIndex;
            }
            else
            {
                MessageBox.Show("조회된 데이터가 없습니다.");
                return;
            }
        }
        #endregion

        #region 조회 FillGrid
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
                sqlParameter.Add("sMcID", "");
                sqlParameter.Add("nchkCheckNeedMC", 0);
                sqlParameter.Add("chkMcCodeSrh", chkMcCodeSrh.IsChecked == true ? 1 : 0); //2021-11-15 설비검색 조건 체크 유무
                sqlParameter.Add("sMcName", chkMcCodeSrh.IsChecked == true ? txtMcCodeSrh.Text : "");
                sqlParameter.Add("iIncNotUse", chkNoUse.IsChecked == true ? 1 : 0); //2021-11-15 사용안함 추가

                ds = DataStore.Instance.ProcedureToDataSet_LogWrite("xp_McCode_sMcCode", sqlParameter, true, "R");

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

                            var WinMCCodeU = new Win_prd_MCCode_U_CodeView()
                            {
                                Num = i,
                                mcid = dr["mcid"].ToString(),
                                mcname = dr["mcname"].ToString(),
                                managerid = dr["managerid"].ToString(),
                                customid = dr["customid"].ToString(),
                                customname = dr["customname"].ToString(),
                                buycustomid = dr["buycustomid"].ToString(),
                                buycustomname = dr["buycustomname"].ToString(),
                                personid = dr["personid"].ToString(),
                                personname = dr["personname"].ToString(),
                                buydate = dr["buydate"].ToString(),
                                buydate_CV = DatePickerFormat(dr["buydate"].ToString()),
                                useyear = dr["useyear"].ToString(),
                                SetHitQty = stringFormatN0(dr["SetHitQty"]),
                                AfterRepairHitcount = stringFormatN0(dr["AfterRepairHitcount"]),
                                HitCount = stringFormatN0(dr["HitCount"]),
                                ProcessID = dr["ProcessID"].ToString(),
                                MachineID = dr["MachineID"].ToString(),
                                LastChangeDate = dr["LastChangeDate"].ToString(),
                                LastChangeDate_CV = DatePickerFormat(dr["LastChangeDate"].ToString()),
                                ImageFile = dr["ImageFile"].ToString(),
                                ImagePath = dr["ImagePath"].ToString(),
                                HrLicenceID = dr["HrLicenceID"].ToString(),
                                HrLicenceName = dr["HrLicenceName"].ToString(),
                                Spec = dr["Spec"].ToString(),
                                ModelName = dr["ModelName"].ToString(),
                                ProductionDate = dr["ProductionDate"].ToString(),
                                ProductionDate_CV = DatePickerFormat(dr["ProductionDate"].ToString()),
                                UseClss = dr["UseClss"].ToString(), //2021-11-15
                            };

                            dgdMain.Items.Add(WinMCCodeU);

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

        #region dgdMain_SelectionChanged
        //
        private void dgdMain_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            imgSetting.Source = null;
            WinMcCode = dgdMain.SelectedItem as Win_prd_MCCode_U_CodeView;

            if (WinMcCode != null)
            {
                this.DataContext = WinMcCode;
                //2021-11-15 사용 안함 체크 일 경우 체크표시 하기
                if (WinMcCode.UseClss.ToString() == "*")
                {
                    chkNotUse.IsChecked = true;
                }
                else
                {
                    chkNotUse.IsChecked = false;
                }

                FillGridMachineMapping(WinMcCode.mcid);
                FillGridSub(WinMcCode.mcid);

                bool MakeFolder = false;
                if (!txtImage.Text.Replace(" ", "").Equals(""))
                {

                    #region 이건 모름

                    string[] fileListSimple;
                    string[] fileListDetail;

                    fileListSimple = _ftp.directoryListSimple("", Encoding.Default);
                    fileListDetail = _ftp.directoryListDetailed("", Encoding.Default);

                    // 기존 폴더 확인작업.                    
                    for (int i = 0; i < fileListSimple.Length; i++)
                    {
                        if (fileListSimple[i] == WinMcCode.mcid)
                        {
                            MakeFolder = true;
                            break;
                        }
                    }

                    if (MakeFolder)
                    {
                        if (CheckImage(WinMcCode.ImageFile.Trim()))
                        {
                            imgSetting.Source = SetImage("/" + WinMcCode.mcid + "/" + txtImage.Text);
                        }
                        else
                        {
                            MessageBox.Show(WinMcCode.ImageFile + "는 이미지 변환이 불가능합니다.");
                        }
                    }

                    #endregion

                    string imageName = txtImage.Text;

                    var Machine = dgdMain.SelectedItem as Win_prd_MCCode_U;
                    if (Machine != null)
                    {
                        //imgSetting.Source = SetImage(imageName, WinMcCode.mcid);
                    }
                }
            }

            if (!txtImage.Text.Trim().Equals(""))
            {
                btnImgDownload.IsEnabled = true;
            }
            else
            {

                btnImgDownload.IsEnabled = false;
            }

        }

        private BitmapImage SetImage(string ImageName, string FolderName)
        {
            BitmapImage bit = null;
            _ftp = new FTP_EX(FTP_ADDRESS, FTP_ID, FTP_PASS);
            if (_ftp == null) { return null; }

            bit = _ftp.DrawingImageByByte(FTP_ADDRESS + '/' + FolderName + '/' + ImageName + "");

            return bit;
        }

        #endregion

        #region 조회 FillGridMachineMapping
        private void FillGridMachineMapping(string strSubID)
        {
            if (dgdMachineMapping.Items.Count > 0)
            {
                dgdMachineMapping.Items.Clear();
            }

            try
            {
                DataSet ds = null;
                Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                sqlParameter.Clear();
                sqlParameter.Add("sMCID", strSubID);
                sqlParameter.Add("chkSEQ", "");
                sqlParameter.Add("SEQ", "");

                ds = DataStore.Instance.ProcedureToDataSet("xp_McCode_sMappingMachineCode", sqlParameter, false);

                if (ds != null && ds.Tables.Count > 0)
                {
                    DataTable dt = ds.Tables[0];
                    int i = 0;

                    if (dt.Rows.Count == 0)
                    {
                        //MessageBox.Show("조회된 데이터가 없습니다.");
                    }
                    else
                    {
                        DataRowCollection drc = dt.Rows;

                        foreach (DataRow dr in drc)
                        {
                            var WinMcCodeMachineMapping = new Win_prd_MCCode_U_MachineMapping_CodeView()
                            {
                                Num = i + 1,
                                MCID = dr["MCID"].ToString(),
                                ProcessID = dr["ProcessID"].ToString(),
                                MachineID = dr["MachineID"].ToString(),
                                MachineNo = dr["MachineNo"].ToString(),
                                Process = dr["Process"].ToString()
                            };

                            dgdMachineMapping.Items.Add(WinMcCodeMachineMapping);
                            i++;
                        }
                    }
                }

            }
            catch (Exception ee)
            {
                MessageBox.Show("오류 발생, 오류 내용 : " + ee.ToString());
            }
            finally
            {
                DataStore.Instance.CloseConnection();
            }
        }
        #endregion

        #region 조회 FillGridSub
        /// <summary>
        /// 설비제품별 생산조건
        /// </summary>
        /// <param name="strSubID"></param>
        private void FillGridSub(string strSubID)
        {
            if (dgdUseMcPart.Items.Count > 0)
            {
                dgdUseMcPart.Items.Clear();
            }

            try
            {
                DataSet ds = null;
                Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                sqlParameter.Clear();
                sqlParameter.Add("McID", strSubID);
                sqlParameter.Add("McPartID", "");
                sqlParameter.Add("ChangeCheckGbn", "");
                ds = DataStore.Instance.ProcedureToDataSet("xp_McCode_sPartChangeProd", sqlParameter, false);

                if (ds != null && ds.Tables.Count > 0)
                {
                    DataTable dt = ds.Tables[0];
                    int i = 0;

                    if (dt.Rows.Count == 0)
                    {
                        //MessageBox.Show("조회된 데이터가 없습니다.");
                    }
                    else
                    {
                        DataRowCollection drc = dt.Rows;

                        foreach (DataRow dr in drc)
                        {
                            var WinMCCodeUSub = new Win_prd_MCCode_U_Sub_CodeView()
                            {
                                Num = i + 1,
                                McID = dr["McID"].ToString(),
                                McPartID = dr["McPartID"].ToString(),
                                ChangeCheckGbn = dr["ChangeCheckGbn"].ToString(),
                                CycleProdQty = dr["CycleProdQty"].ToString(),
                                StartSetProdQty = dr["StartSetProdQty"].ToString(),
                                StartSetDate = dr["StartSetDate"].ToString(),
                                MCPartName = dr["MCPartName"].ToString()
                            };

                            dgdUseMcPart.Items.Add(WinMCCodeUSub);
                            i++;
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

        #region 삭제
        /// <summary>
        /// 실삭제
        /// </summary>
        /// <param name="strID"></param>
        /// <returns></returns>
        private bool DeleteData(string strID)
        {
            bool flag = false;

            try
            {
                Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                sqlParameter.Clear();
                sqlParameter.Add("sMcID", strID);

                string[] result = DataStore.Instance.ExecuteProcedure_NewLog("xp_McCode_dMcCode", sqlParameter, "D");

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
        #endregion

        #region 공정설비 매핑 삭제 

        private bool MappingDeleteData(string strID)
        {
            bool flag = false;

            try
            {
                Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                sqlParameter.Clear();
                sqlParameter.Add("sMcID", strID);

                string[] result = DataStore.Instance.ExecuteProcedure("xp_McCode_dMappingMachineCode", sqlParameter, false);

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
        #endregion

        #region 저장
        /// <summary>
        /// 저장
        /// </summary>
        /// <param name="strFlag"></param>
        /// <param name="strID"></param>
        /// <returns></returns>
        private bool SaveData(string strFlag, string strID)
        {
            bool flag = false;
            List<Procedure> Prolist = new List<Procedure>();
            List<Dictionary<string, object>> ListParameter = new List<Dictionary<string, object>>();

            string GetKey = "";

            try
            {
                if (CheckData())
                {

                    Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                    sqlParameter.Clear();

                    //sqlParameter.Add("sNewMCID", "");
                    sqlParameter.Add("sMcName", txtMCName.Text);
                    sqlParameter.Add("sManagerID", txtManageID.Text);
                    sqlParameter.Add("sCustomID", txtCustom.Tag != null ? txtCustom.Tag.ToString() : "");
                    sqlParameter.Add("sCustomName", txtCustom.Text);

                    sqlParameter.Add("sBuyCustomID", txtBuyCustom.Tag != null ? txtBuyCustom.Tag.ToString() : "");
                    sqlParameter.Add("sBuyCustomName", txtBuyCustom.Text);
                    sqlParameter.Add("sPersonID", txtPerson.Tag != null ? txtPerson.Tag.ToString() : "");
                    sqlParameter.Add("sPersonName", txtPerson.Text);
                    sqlParameter.Add("sBuyDate", dtpBuyDate.SelectedDate != null ? dtpBuyDate.SelectedDate.Value.ToString("yyyyMMdd") : "");

                    sqlParameter.Add("sUseYear", ConvertInt(txtUseYear.Text.Replace(",", "")));
                    sqlParameter.Add("sSetHitQty", txtSetHitQty.Text.Replace(",", "").Length > 0 ? Convert.ToDouble(txtSetHitQty.Text.Replace(",", "")) : 0);
                    sqlParameter.Add("sProcessID", WinMcCodeMachineMapping.ProcessID != null ? WinMcCodeMachineMapping.ProcessID : "");
                    sqlParameter.Add("sMachineID", WinMcCodeMachineMapping.MachineID != null ? WinMcCodeMachineMapping.MachineID : "");
                    sqlParameter.Add("sLastChangeDate", DateTime.Today.ToString("yyyyMMdd"));

                    sqlParameter.Add("sImageFile", "");
                    sqlParameter.Add("sImagePath", "");
                    sqlParameter.Add("sHrLicenceID", cboLicense.SelectedValue != null ? cboLicense.SelectedValue.ToString() : "");
                    sqlParameter.Add("sSpec", txtSpec.Text);
                    sqlParameter.Add("sProductionDate", dtpProductionDate.SelectedDate != null ? dtpProductionDate.SelectedDate.Value.ToString("yyyyMMdd") : "");
                    sqlParameter.Add("sModelName", txtModelName.Text);

                    sqlParameter.Add("sUseClss", chkNotUse.IsChecked == true ? "*" : ""); //2021-11-15 사용안함 추가


                    #region 추가

                    if (strFlag.Equals("I"))
                    {
                        sqlParameter.Add("sNewMCID", "");
                        sqlParameter.Add("CreateUserID", MainWindow.CurrentUser);

                        Procedure pro1 = new Procedure();
                        pro1.Name = "xp_McCode_iMcCode";
                        pro1.OutputUseYN = "Y";
                        pro1.OutputName = "sNewMCID";
                        pro1.OutputLength = "10";

                        Prolist.Add(pro1);
                        ListParameter.Add(sqlParameter);


                        //동운씨가 만든 아웃풋 값 찾는 방법
                        List<KeyValue> list_Result = new List<KeyValue>();
                        list_Result = DataStore.Instance.ExecuteAllProcedureOutputGetCS_NewLog(Prolist, ListParameter, "C");

                        Prolist.RemoveAt(0);
                        ListParameter.RemoveAt(0);

                        string sGetID = string.Empty;

                        if (list_Result[0].key.ToLower() == "success")
                        {
                            list_Result.RemoveAt(0);
                            for (int i = 0; i < list_Result.Count; i++)
                            {
                                KeyValue kv = list_Result[i];
                                if (kv.key == "sNewMCID")
                                {
                                    sGetID = kv.value;
                                    WinMcCodeMachineMapping.MCID = sGetID;
                                    GetKey = sGetID;
                                    flag = true;
                                }
                            }
                        }
                        else
                        {
                            MessageBox.Show("[저장실패]\r\n" + list_Result[0].value.ToString());
                            flag = false;
                            //return false;
                        }


                        for (int i = 0; i < dgdUseMcPart.Items.Count; i++)
                        {
                            WinMcCodeSub = dgdUseMcPart.Items[i] as Win_prd_MCCode_U_Sub_CodeView;
                            sqlParameter = new Dictionary<string, object>();
                            sqlParameter.Clear();
                            sqlParameter.Add("McID", sGetID);
                            sqlParameter.Add("sNewMCID", "");
                            sqlParameter.Add("McPartID", WinMcCodeSub.McPartID);
                            sqlParameter.Add("ChangeCheckGbn", WinMcCodeSub.ChangeCheckGbn);
                            sqlParameter.Add("ProdCycle", WinMcCodeSub.CycleProdQty != null ? ConvertDouble(WinMcCodeSub.CycleProdQty) : 0);
                            sqlParameter.Add("StartSetProdQty", WinMcCodeSub.StartSetProdQty != null ? ConvertDouble(WinMcCodeSub.StartSetProdQty) : 0);
                            sqlParameter.Add("StartSetDate", WinMcCodeSub.StartSetDate);
                            sqlParameter.Add("CreateUserID", MainWindow.CurrentUser);

                            Procedure pro2 = new Procedure();
                            pro2.Name = "xp_McCode_iPartChangeProd";
                            pro2.OutputUseYN = "N";
                            pro2.OutputName = "sNewMCID";
                            pro2.OutputLength = "10";

                            Prolist.Add(pro2);
                            ListParameter.Add(sqlParameter);

                        }

                        for (int i = 0; i < dgdMachineMapping.Items.Count; i++)
                        {
                            WinMcCodeMachineMapping = dgdMachineMapping.Items[i] as Win_prd_MCCode_U_MachineMapping_CodeView;
                            sqlParameter = new Dictionary<string, object>();
                            sqlParameter.Clear();
                            sqlParameter.Add("sMCID", sGetID); // MCID 값을 못가져와서 수정 WinMcCodeMachineMapping.MCID -> sGetID 2021-10-07
                            sqlParameter.Add("sProcessID", WinMcCodeMachineMapping.ProcessID);
                            sqlParameter.Add("sMachineID", WinMcCodeMachineMapping.MachineID);
                            sqlParameter.Add("CreateUserID", MainWindow.CurrentUser); //로그인한 사람

                            Procedure pro3 = new Procedure();
                            pro3.Name = "xp_McCode_iMappingMachineCode";
                            pro3.OutputUseYN = "N";
                            pro3.OutputName = "sNewMCID";
                            pro3.OutputLength = "10";

                            Prolist.Add(pro3);
                            ListParameter.Add(sqlParameter);

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

                    #endregion

                    #region 수정

                    else if (strFlag.Equals("U"))
                    {
                        sqlParameter.Add("sMCID", strID);
                        sqlParameter.Add("UpdateUserID", MainWindow.CurrentUser);

                        Procedure pro1 = new Procedure();
                        pro1.Name = "xp_McCode_uMcCode";
                        pro1.OutputUseYN = "N";
                        pro1.OutputName = "sMCID";
                        pro1.OutputLength = "10";

                        Prolist.Add(pro1);
                        ListParameter.Add(sqlParameter);

                        for (int i = 0; i < dgdUseMcPart.Items.Count; i++)
                        {
                            WinMcCodeSub = dgdUseMcPart.Items[i] as Win_prd_MCCode_U_Sub_CodeView;
                            sqlParameter = new Dictionary<string, object>();
                            sqlParameter.Clear();
                            sqlParameter.Add("McID", WinMcCodeSub.McID);
                            sqlParameter.Add("sNewMCID", WinMcCode.mcid);
                            sqlParameter.Add("McPartID", WinMcCodeSub.McPartID);
                            sqlParameter.Add("ChangeCheckGbn", WinMcCodeSub.ChangeCheckGbn);
                            sqlParameter.Add("ProdCycle", 0);
                            sqlParameter.Add("StartSetProdQty", 0);
                            sqlParameter.Add("StartSetDate", /*WinMcCodeSub.StartSetDate*/ DateTime.Today.ToString("yyyyMMdd"));
                            sqlParameter.Add("CreateUserID", MainWindow.CurrentUser);

                            Procedure pro2 = new Procedure();
                            pro2.Name = "xp_McCode_iPartChangeProd";
                            pro2.OutputUseYN = "N";
                            pro2.OutputName = "sMCID";
                            pro2.OutputLength = "10";

                            Prolist.Add(pro2);
                            ListParameter.Add(sqlParameter);
                        }

                        for (int i = 0; i < dgdMachineMapping.Items.Count; i++)
                        {
                            WinMcCodeMachineMapping = dgdMachineMapping.Items[i] as Win_prd_MCCode_U_MachineMapping_CodeView;
                            sqlParameter = new Dictionary<string, object>();
                            sqlParameter.Clear();
                            sqlParameter.Add("sMCID", WinMcCode.mcid);
                            sqlParameter.Add("sProcessID", WinMcCodeMachineMapping.ProcessID);
                            sqlParameter.Add("sMachineID", WinMcCodeMachineMapping.MachineID);
                            sqlParameter.Add("CreateUserID", MainWindow.CurrentUser); //로그인한 사람

                            Procedure pro3 = new Procedure();
                            pro3.Name = "xp_McCode_iMappingMachineCode";
                            pro3.OutputUseYN = "N";
                            pro3.OutputName = "sNewMCID";
                            pro3.OutputLength = "10";

                            Prolist.Add(pro3);
                            ListParameter.Add(sqlParameter);

                        }

                        string[] Confirm = new string[2];
                        Confirm = DataStore.Instance.ExecuteAllProcedureOutputNew_NewLog(Prolist, ListParameter, "U"); // 저장되는 소스
                        if (Confirm[0] != "success")
                        {
                            MessageBox.Show("[저장실패]\r\n" + Confirm[1].ToString());
                            flag = false;
                            //return false;
                        }
                        else
                        {
                            flag = true;
                            GetKey = strID;
                        }
                    }

                    // 파일을 올리자 : GetKey != "" 라면 파일을 올려보자
                    if (!GetKey.Trim().Equals(""))
                    {
                        if (deleteListFtpFile.Count > 0)
                        {
                            foreach (string[] str in deleteListFtpFile)
                            {
                                FTP_RemoveFile(GetKey + "/" + str[0]);
                            }
                        }

                        if (listFtpFile.Count > 0) // /ImageData/McRIB/2020010001
                        {
                            if (FTP_Save_File(listFtpFile, GetKey))
                            {
                                UpdateDBFtp(GetKey, txtImage.Text, @"/ImageData/McRIB/" + GetKey);
                            }
                        }

                    }

                    // 파일 List 비워주기
                    listFtpFile.Clear();
                    deleteListFtpFile.Clear();
                    #endregion
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

        #region SaveData → UpdateDBFtp : 파일이름, 경로 업데이트

        private bool UpdateDBFtp(string strID, string ImageName, string ImagePath)
        {
            bool flag = false;

            List<Procedure> Prolist = new List<Procedure>();
            List<Dictionary<string, object>> ListParameter = new List<Dictionary<string, object>>();

            try
            {
                Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                sqlParameter.Clear();

                sqlParameter.Add("MCID", strID);
                sqlParameter.Add("ImageFile", ImageName);
                sqlParameter.Add("ImagePath", ImagePath);

                Procedure pro1 = new Procedure();
                pro1.Name = "xp_prd_uMcCode_FTP";
                pro1.OutputUseYN = "N";
                pro1.OutputName = "sMCID";
                pro1.OutputLength = "10";

                Prolist.Add(pro1);
                ListParameter.Add(sqlParameter);

                string[] Confirm = new string[2];
                Confirm = DataStore.Instance.ExecuteAllProcedureOutputNew(Prolist, ListParameter); // 저장되는 소스
                if (Confirm[0] != "success")
                {
                    MessageBox.Show("[저장실패]\r\n" + Confirm[1].ToString());
                    flag = false;
                    //return false;
                }
                else
                {
                    flag = true;
                    //GetKey = strID;
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

        #region 입력체크
        /// <summary>
        /// 입력사항 체크
        /// </summary>
        /// <returns></returns>
        private bool CheckData()
        {
            bool flag = true;

            if (String.IsNullOrEmpty(txtMCName.Text))
            {
                MessageBox.Show("설비명이 입력되지 않았습니다.");
                flag = false;
                return flag;
            }

            if (String.IsNullOrEmpty(txtCustom.Text))
            {
                MessageBox.Show("제작사가 입력되지 않았습니다.");
                flag = false;
                return flag;
            }

            if (String.IsNullOrEmpty(txtBuyCustom.Text))
            {
                MessageBox.Show("구매처가 입력되지 않았습니다.");
                flag = false;
                return flag;
            }

            if (String.IsNullOrEmpty(txtPerson.Text))
            {
                MessageBox.Show("관리자가 입력되지 않았습니다.");
                flag = false;
                return flag;
            }

            //if (cboProcess.SelectedValue == null)
            //{
            //    MessageBox.Show("공정이 선택되지 않았습니다.");
            //    flag = false;
            //    return flag;
            //}

            // 공정설비매핑 행을 만들고
            // 1. 공정을 입력하지 않았을 시
            // 2. 호기를 입력하지 않았을 시
            for (int i = 0; i < dgdMachineMapping.Items.Count; i++)
            {
                var Mapping = dgdMachineMapping.Items[i] as Win_prd_MCCode_U_MachineMapping_CodeView;
                if (Mapping != null)
                {
                    if (Mapping.ProcessID == null
                        || Mapping.ProcessID.Trim().Equals(""))
                    {
                        MessageBox.Show("공정설비 매핑에 공정 및 호기를 입력해주세요.");
                        flag = false;
                        return flag;
                    }

                    if (Mapping.MachineID == null
                       || Mapping.MachineID.Trim().Equals(""))
                    {
                        MessageBox.Show("공정설비 매핑에 호기를 입력해주세요.");
                        flag = false;
                        return flag;
                    }
                }
            }

            return flag;
        }
        #endregion

        #region Content - 플러스 파인더 모음

        //설비명
        private void txtMCName_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                MainWindow.pf.ReturnCode(txtMCName, 12, "");

                if (txtMCName.Tag != null)
                {
                    txtCustom.Focus();
                }
            }
        }


        //설비명
        private void btnMCName_Click(object sender, RoutedEventArgs e)
        {
            MainWindow.pf.ReturnCode(txtMCName, 12, "");

            //개정일자에 포커스 이동
            txtCustom.Focus();
        }

        //제작사
        private void txtCustom_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                MainWindow.pf.ReturnCode(txtCustom, (int)Defind_CodeFind.DCF_CUSTOM, "");

                if (txtCustom.Tag != null)
                {
                    txtBuyCustom.Focus();
                }
            }
        }

        //제작사
        private void btnPfCustom_Click(object sender, RoutedEventArgs e)
        {
            MainWindow.pf.ReturnCode(txtCustom, (int)Defind_CodeFind.DCF_CUSTOM, "");

            if (txtCustom.Tag != null)
            {
                txtBuyCustom.Focus();
            }
        }

        //구매처
        private void txtBuyCustom_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                MainWindow.pf.ReturnCode(txtBuyCustom, (int)Defind_CodeFind.DCF_CUSTOM, "");

                if (txtBuyCustom.Tag != null)
                {
                    txtPerson.Focus();
                }
            }

        }

        //구매처
        private void btnPfBuyCustom_Click(object sender, RoutedEventArgs e)
        {
            MainWindow.pf.ReturnCode(txtBuyCustom, (int)Defind_CodeFind.DCF_CUSTOM, "");

            if (txtBuyCustom.Tag != null)
            {
                txtPerson.Focus();
            }
        }

        //관리자
        private void txtPerson_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                MainWindow.pf.ReturnCode(txtPerson, (int)Defind_CodeFind.DCF_PERSON, "");

                if (txtPerson.Tag != null)
                {
                    cboLicense.Focus();
                }
            }
        }

        //관리자
        private void btnPfPerson_Click(object sender, RoutedEventArgs e)
        {
            MainWindow.pf.ReturnCode(txtPerson, (int)Defind_CodeFind.DCF_PERSON, "");

            if (cboLicense.Tag != null)
            {
                cboLicense.Focus();
            }
        }
        #endregion

        #region 공정설비 매핑 KeyDown Event

        //(공정)
        private void DataGridCell_MachineMapping_KeyDown(object sender, KeyEventArgs e)
        {

        }

        private void TextBoxFocusInDataGrid_MachineMapping(object sender, KeyEventArgs e)
        {
            Lib.Instance.DataGridINControlFocus(sender, e);
        }

        private void TextBoxFocusInDataGrid_MachineMapping_MouseUp(object sender, MouseButtonEventArgs e)
        {
            Lib.Instance.DataGridINBothByMouseUP(sender, e);
        }

        private void DataGridCell_MachineMapping_GotFocus(object sender, RoutedEventArgs e)
        {
            if (lblMsg.Visibility == Visibility.Visible)
            {
                DataGridCell cell = sender as DataGridCell;
                cell.IsEditing = true;
            }
        }

        //(호기)
        private void DataGridCell_MachineMappingHo_KeyDown(object sender, KeyEventArgs e)
        {

        }

        private void TextBoxFocusInDataGrid_MachineMappingHo(object sender, KeyEventArgs e)
        {
            Lib.Instance.DataGridINControlFocus(sender, e);
        }

        private void TextBoxFocusInDataGrid_MachineMappingHo_MouseUp(object sender, MouseButtonEventArgs e)
        {
            Lib.Instance.DataGridINBothByMouseUP(sender, e);
        }

        private void DataGridCell_MachineMappingHo_GotFocus(object sender, RoutedEventArgs e)
        {
            if (lblMsg.Visibility == Visibility.Visible)
            {
                DataGridCell cell = sender as DataGridCell;
                cell.IsEditing = true;
            }
        }

        #endregion

        #region dgdMachineMapping 행 추가, 삭제
        //공정설비매핑 추가 버튼
        private void btnAddMachineMapping_Click(object sender, RoutedEventArgs e)
        {
            var WinMcCodeMachineMapping = new Win_prd_MCCode_U_MachineMapping_CodeView
            {
                MCID = "",
                ProcessID = "",
                MachineID = "",
                Process = "",
                Machine = "",
                Num = dgdMachineMapping.Items.Count + 1
            };

            dgdMachineMapping.Items.Add(WinMcCodeMachineMapping);
        }

        //공정설비매핑 삭제 버튼
        private void btnDelMachineMapping_Click(object sender, RoutedEventArgs e)
        {
            if (dgdMachineMapping.Items.Count > 0)
            {
                if (dgdMachineMapping.CurrentItem != null)
                {
                    dgdMachineMapping.Items.Remove(dgdMachineMapping.CurrentItem as Win_prd_MCCode_U_MachineMapping_CodeView);
                }
                else
                {
                    dgdMachineMapping.Items.Remove((dgdMachineMapping.Items[dgdMachineMapping.Items.Count - 1]) as Win_prd_MCCode_U_MachineMapping_CodeView);
                }

                dgdMachineMapping.Refresh();
            }
        }

        //공정설비매핑 KeyDown Event
        private void DatagridTextBoxMachineMapping_KeyDown(object sender, KeyEventArgs e)
        {
            if (lblMsg.Visibility == Visibility.Visible)
            {
                if (e.Key == Key.Enter)
                {
                    WinMcCodeMachineMapping = dgdMachineMapping.CurrentItem as Win_prd_MCCode_U_MachineMapping_CodeView;

                    if (WinMcCodeMachineMapping != null)
                    {
                        TextBox tb1 = sender as TextBox;

                        MainWindow.pf.ReturnCode(tb1, (int)Defind_CodeFind.DCF_PROCESS, "");

                        if (tb1.Tag != null)
                        {
                            WinMcCodeMachineMapping.Process = tb1.Text;
                            WinMcCodeMachineMapping.ProcessID = tb1.Tag.ToString();
                        }
                    }
                }
            }
        }

        private void DatagridTextBoxMachineMappingHo_KeyDown(object sender, KeyEventArgs e)
        {
            if (lblMsg.Visibility == Visibility.Visible)
            {
                if (e.Key == Key.Enter)
                {
                    var Machine = dgdMachineMapping.CurrentItem as Win_prd_MCCode_U_MachineMapping_CodeView;

                    if (Machine != null)
                    {
                        if (Machine.ProcessID != null
                            && !Machine.ProcessID.Trim().Equals(""))
                        {
                            TextBox tb1 = sender as TextBox;

                            MainWindow.pf.ReturnCode(tb1, 79, Machine.ProcessID);

                            if (tb1.Tag != null)
                            {
                                Machine.MachineNo = tb1.Text;
                                Machine.MachineID = tb1.Tag.ToString();
                            }
                        }
                        else
                        {
                            MessageBox.Show("공정을 먼저 선택해주세요.");
                            return;
                        }

                    }
                }
            }
        }
        #endregion

        #region 사용부품 KeyDown Event
        //
        private void DataGridCell_KeyDown(object sender, KeyEventArgs e)
        {

        }

        //TextBox or ComboBox Cell Focus(keydown)
        private void TextBoxFocusInDataGrid(object sender, KeyEventArgs e)
        {
            Lib.Instance.DataGridINControlFocus(sender, e);
        }

        //TextBox or ComboBox Cell Focus(MouseClick)
        private void TextBoxFocusInDataGrid_MouseUp(object sender, MouseButtonEventArgs e)
        {
            Lib.Instance.DataGridINBothByMouseUP(sender, e);
        }

        //cellEditingMode 진입
        private void DataGridCell_GotFocus(object sender, RoutedEventArgs e)
        {
            if (lblMsg.Visibility == Visibility.Visible)
            {
                DataGridCell cell = sender as DataGridCell;
                cell.IsEditing = true;
            }
        }

        //서브그리드 내부 입력 이벤트
        private void dgdtpetxtMCPartName_KeyDown(object sender, KeyEventArgs e)
        {
            if (lblMsg.Visibility == Visibility.Visible)
            {
                if (e.Key == Key.Enter)
                {
                    WinMcCodeSub = dgdUseMcPart.CurrentItem as Win_prd_MCCode_U_Sub_CodeView;

                    if (WinMcCodeSub != null)
                    {
                        TextBox tb1 = sender as TextBox;
                        MainWindow.pf.ReturnCode(tb1, (int)Defind_CodeFind.DCF_PART, "");

                        if (tb1.Tag != null)
                        {
                            WinMcCodeSub.MCPartName = tb1.Text;
                            WinMcCodeSub.McPartID = tb1.Tag.ToString();
                        }
                    }
                }
            }
        }
        #endregion

        #region dgdUseMcPart 행 추가, 삭제

        //서브 그리드 행 추가
        private void btnSubAdd_Click(object sender, RoutedEventArgs e)
        {
            var WinMcCodeSub = new Win_prd_MCCode_U_Sub_CodeView
            {
                McID = "",
                ChangeCheckGbn = "",
                CycleProdQty = "",
                McPartID = "",
                MCPartName = "",
                StartSetDate = "",
                StartSetProdQty = "",
                Num = dgdUseMcPart.Items.Count + 1
            };

            dgdUseMcPart.Items.Add(WinMcCodeSub);
        }

        //서브 그리드 행 삭제
        private void btnSubDel_Click(object sender, RoutedEventArgs e)
        {
            if (dgdUseMcPart.Items.Count > 0)
            {
                if (dgdUseMcPart.CurrentItem != null)
                {
                    dgdUseMcPart.Items.Remove((dgdUseMcPart.CurrentItem as Win_prd_MCCode_U_Sub_CodeView));
                }
                else
                {
                    dgdUseMcPart.Items.Remove((dgdUseMcPart.Items[dgdUseMcPart.Items.Count - 1]) as Win_prd_MCCode_U_Sub_CodeView);
                }
                dgdUseMcPart.Refresh();
            }
        }
        #endregion

        #region 이미지 업로드, 삭제
        //업로드할 파일을 선택해준다.
        private void btnImage_Click(object sender, RoutedEventArgs e)
        {
            if (!txtImage.Text.Equals(string.Empty) && strFlag.Equals("U"))
            {
                MessageBox.Show("먼저 해당파일의 삭제를 진행 후 진행해주세요.");
                return;
            }
            else
            {
                FTP_Upload_TextBox(txtImage);
            }
        }

        private void FTP_Upload_TextBox(TextBox textBox)
        {
            if (!textBox.Text.Equals(string.Empty) && strFlag.Equals("U"))
            {
                MessageBox.Show("먼저 해당파일의 삭제를 진행 후 진행해주세요.");
                return;
            }
            else
            {
                Microsoft.Win32.OpenFileDialog OFdlg = new Microsoft.Win32.OpenFileDialog();
                OFdlg.Filter = MainWindow.OFdlg_Filter;

                Nullable<bool> result = OFdlg.ShowDialog();
                if (result == true)
                {
                    strFullPath = OFdlg.FileName;

                    string ImageFileName = OFdlg.SafeFileName;  //명.
                    string ImageFilePath = string.Empty;       // 경로

                    ImageFilePath = strFullPath.Replace(ImageFileName, "");

                    StreamReader sr = new StreamReader(OFdlg.FileName);
                    long FileSize = sr.BaseStream.Length;
                    if (sr.BaseStream.Length > (2048 * 1000))
                    {
                        //업로드 파일 사이즈범위 초과
                        MessageBox.Show("이미지의 파일사이즈가 2M byte를 초과하였습니다.");
                        sr.Close();
                        return;
                    }
                    else
                    {
                        textBox.Text = ImageFileName;
                        textBox.Tag = ImageFilePath;

                        try
                        {
                            Bitmap image = new Bitmap(ImageFilePath + ImageFileName);
                            imgSetting.Source = BitmapToImageSource(image);
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show("해당 파일은 이미지로 변환이 불가능합니다.");
                        }

                        string[] strTemp = new string[] { ImageFileName, ImageFilePath.ToString() };
                        listFtpFile.Add(strTemp);
                    }
                }
            }
        }

        //이미지 삭제(폴더까지 삭제한다)
        private void btnImgDel_Click(object sender, RoutedEventArgs e)
        {
            MessageBoxResult msgresult = MessageBox.Show("파일을 삭제 하시겠습니까?", "삭제 확인", MessageBoxButton.YesNo);
            if (msgresult == MessageBoxResult.Yes)
            {

                #region 이건 모름

                //if (strFlag.Equals("U") && existFtp == true)
                //{
                //    if (FTP_RemoveDir(txtMCID.Text))
                //    {
                //        delFtp = true;
                //    }
                //}

                //strDelFileName = txtImage.Text;
                //txtImage.Text = "";
                //txtImage.Tag = null;

                //imgSetting.Source = null;

                #endregion


                FileDeleteAndTextBoxEmpty(txtImage);
                imgSetting.Source = null;
            }
        }

        private void FileDeleteAndTextBoxEmpty(TextBox txt)
        {
            if (strFlag.Equals("U"))
            {
                // 파일이름, 파일경로
                string[] strFtp = { txt.Text, txt.Tag != null ? txt.Tag.ToString() : "" };

                deleteListFtpFile.Add(strFtp);
            }

            txt.Text = "";
            txt.Tag = "";
        }

        //image 만 Bit로 세팅( imageSource랑 바인딩 )
        private BitmapImage SetImage(string strAttachPath)
        {
            BitmapImage bit = _ftp.DrawingImageByByte(FTP_ADDRESS + strAttachPath + "");
            //image.Source = bit;
            return bit;
        }

        //파일 삭제
        private bool FTP_RemoveFile(string strSaveName)
        {
            _ftp = new FTP_EX(FTP_ADDRESS, FTP_ID, FTP_PASS);
            if (_ftp.delete(strSaveName) == true)
            {
                return true;
            }
            else
            {
                return false;
            }
        }
        //폴더 삭제(내부 파일 자동 삭제)
        private bool FTP_RemoveDir(string strSaveName)
        {
            _ftp = new FTP_EX(FTP_ADDRESS, FTP_ID, FTP_PASS);
            if (_ftp.removeDir(strSaveName) == true)
            {
                return true;
            }
            else
            {
                return false;
            }
        }


        //FTP 업로드시 파일체크 및 경로,파일이름 표시
        private TextBox Ftp_Upload_TextBox()
        {
            TextBox tb = new TextBox();

            Microsoft.Win32.OpenFileDialog OFdlg = new Microsoft.Win32.OpenFileDialog();
            OFdlg.Filter = MainWindow.OFdlg_Filter;

            Nullable<bool> result = OFdlg.ShowDialog();
            if (result == true)
            {
                strFullPath = OFdlg.FileName;

                string ImageFileName = OFdlg.SafeFileName;  //명.
                string ImageFilePath = string.Empty;       // 경로

                ImageFilePath = strFullPath.Replace(ImageFileName, "");

                StreamReader sr = new StreamReader(OFdlg.FileName);
                long FileSize = sr.BaseStream.Length;
                if (sr.BaseStream.Length > (2048 * 1000))
                {
                    //업로드 파일 사이즈범위 초과
                    //MessageBox.Show("이미지의 파일사이즈가 2M byte를 초과하였습니다.");
                    sr.Close();
                    tb.Text = "파일사이즈초과";
                    //return;
                }
                else
                {
                    tb.Text = ImageFileName;
                    tb.Tag = ImageFilePath;

                    Bitmap image = new Bitmap(ImageFilePath + ImageFileName);
                    imgSetting.Source = BitmapToImageSource(image);

                }
            }

            return tb;
        }


        // 비트맵을 비트맵 이미지로 형태변환시키기.<0823 허윤구> 
        BitmapImage BitmapToImageSource(Bitmap bitmap)
        {
            using (MemoryStream memory = new MemoryStream())
            {
                bitmap.Save(memory, System.Drawing.Imaging.ImageFormat.Bmp);
                memory.Position = 0;
                BitmapImage bitmapimage = new BitmapImage();
                bitmapimage.BeginInit();
                bitmapimage.StreamSource = memory;
                bitmapimage.CacheOption = BitmapCacheOption.OnLoad;
                bitmapimage.EndInit();

                return bitmapimage;
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

        // 파일 저장하기.
        private bool FTP_Save_File(List<string[]> listStrArrayFileInfo, string MakeFolderName)
        {
            bool result = false;

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
                result = false;
            }
            else
            {
                result = true;
            }

            return result;
        }

        //업로드 폴더 존재유무 확인
        private bool MakeFileInfoList(string[] simple, string[] detail, string str_InspectID)
        {
            bool tf_return = false;
            foreach (string filename in simple)
            {
                foreach (string info in detail)
                {
                    if (info.Contains(filename) == true)
                    {
                        if (MakeFileInfoList(filename, info, str_InspectID) == true)
                        {
                            tf_return = true;
                        }
                    }
                }
            }
            return tf_return;
        }

        //업로드 폴더 존재유무 확인
        private bool MakeFileInfoList(string simple, string detail, string strCompare)
        {
            UploadFileInfo info = new UploadFileInfo();
            info.Filename = simple;
            info.Filepath = detail;

            if (simple.Length > 0)
            {
                string[] tokens = detail.Split(new[] { ' ' }, 9, StringSplitOptions.RemoveEmptyEntries);
                string name = tokens[3].ToString();
                string permissions = tokens[2].ToString();

                if (permissions.Contains("D") == true)
                {
                    info.Type = FtpFileType.DIR;
                }
                else
                {
                    info.Type = FtpFileType.File;
                }

                if (info.Type == FtpFileType.File)
                {
                    info.Size = Convert.ToInt64(detail.Substring(17, detail.LastIndexOf(simple) - 17).Trim());
                }

                _listFileInfo.Add(info);

                if (string.Compare(simple, strCompare, false) == 0)
                    return true;
            }

            return false;
        }

        ////폴더 삭제(내부 파일 자동 삭제)
        //private bool FTP_RemoveDir(string strSaveName)
        //{
        //    string[] fileListSimple;
        //    string[] fileListDetail;

        //    fileListSimple = _ftp.directoryListSimple("", Encoding.Default);
        //    fileListDetail = _ftp.directoryListDetailed("", Encoding.Default);

        //    bool tf_ExistInspectionID = MakeFileInfoList(fileListSimple, fileListDetail, strSaveName);

        //    if (tf_ExistInspectionID == true)
        //    {
        //        if (_ftp.removeDir(strSaveName) == true)
        //        {
        //            return true;
        //        }
        //        else
        //        {
        //            return false;
        //        }
        //    }
        //    return true;
        //}

        //FTP의 파일을 다운로드
        private void FTP_DownLoadFile(string strFilePath)
        {
            string[] fileListSimple;
            string[] fileListDetail;

            fileListSimple = _ftp.directoryListSimple("", Encoding.Default);
            fileListDetail = _ftp.directoryListDetailed("", Encoding.Default);

            bool ExistFile = false;
            ExistFile = MakeFileInfoList(fileListSimple, fileListDetail, strFilePath.Split('/')[3].Trim());

            int fileLength = _listFileInfo.Count;

            if (ExistFile)
            {
                string str_remotepath = string.Empty;
                string str_localpath = string.Empty;

                str_remotepath = strFilePath.ToString();
                str_localpath = LOCAL_DOWN_PATH + "\\" + strFilePath.Substring(strFilePath.LastIndexOf("/")).ToString();

                DirectoryInfo DI = new DirectoryInfo(LOCAL_DOWN_PATH);
                if (DI.Exists)
                {
                    DI.Create();
                }

                FileInfo file = new FileInfo(str_localpath);
                if (file.Exists)
                {
                    //if (MessageBox.Show("같은 이름의 파일이 존재하여" +
                    //    "진행합니다. 계속 하시겠습니까?", "", MessageBoxButton.YesNo) == MessageBoxResult.Yes)
                    //{
                    //    file.Delete();
                    //}
                    //else
                    //{
                    //    MessageBox.Show("C:Temp 폴더를 확인하세요.");
                    //    return;
                    //}
                    file.Delete();
                }

                _ftp.download(str_remotepath.Substring(str_remotepath.Substring
                    (0, str_remotepath.LastIndexOf("/")).LastIndexOf("/")), str_localpath);

                ProcessStartInfo proc = new ProcessStartInfo(str_localpath);
                proc.UseShellExecute = true;
                Process.Start(proc);
            }
        }






        #endregion

        //이미지 다운로드
        private void btnImgDownload_Click(object sender, RoutedEventArgs e)
        {
            MessageBoxResult msgresult = MessageBox.Show("파일을 다운로드 하시겠습니까?", "보기 확인", MessageBoxButton.YesNo);
            if (msgresult == MessageBoxResult.Yes)
            {

                if (txtImage.Text.Trim().Equals(""))
                {
                    MessageBox.Show("파일이 존재하지 않습니다.");
                    return;
                }

                try
                {
                    var Machine = dgdMain.SelectedItem as Win_prd_MCCode_U_CodeView;

                    if (Machine != null)
                    {
                        // 접속 경로
                        _ftp = new FTP_EX(FTP_ADDRESS, FTP_ID, FTP_PASS);

                        string str_path = string.Empty;
                        str_path = FTP_ADDRESS + '/' + Machine.mcid;
                        _ftp = new FTP_EX(str_path, FTP_ID, FTP_PASS);

                        string str_remotepath = string.Empty;
                        string str_localpath = string.Empty;

                        str_remotepath = txtImage.Text;
                        str_localpath = LOCAL_DOWN_PATH + "\\" + txtImage.Text;

                        DirectoryInfo DI = new DirectoryInfo(LOCAL_DOWN_PATH);      // Temp 폴더가 없는 컴터라면, 만들어 줘야지.
                        if (DI.Exists == false)
                        {
                            DI.Create();
                        }

                        FileInfo file = new FileInfo(str_localpath);
                        if (file.Exists)
                        {
                            file.Delete();
                        }

                        _ftp.download(str_remotepath, str_localpath);

                        ProcessStartInfo proc = new ProcessStartInfo(str_localpath.Trim());
                        proc.UseShellExecute = true;
                        Process.Start(proc);
                    }
                }
                catch (Exception ex) // 뭐든 간에 파일 없다고 하자
                {
                    MessageBox.Show("파일이 존재하지 않습니다.\r관리자에게 문의해주세요.");
                    return;
                }
            }
        }

        #region 테스트용 - 공정 매핑 세팅
        private void btnMapping_Click(object sender, RoutedEventArgs e)
        {
            List<Procedure> Prolist = new List<Procedure>();
            List<Dictionary<string, object>> ListParameter = new List<Dictionary<string, object>>();

            Dictionary<string, object> sqlParameter = new Dictionary<string, object>();

            try
            {
                //for (int i = 0; i < dgdMain.Items.Count; i++)
                //{
                //var Main = dgdMain.Items[i] as Win_prd_MCCode_U_CodeView;
                //if (Main != null)
                //{
                sqlParameter = new Dictionary<string, object>();
                sqlParameter.Clear();
                //sqlParameter.Add("sMCID", Main.mcid);
                //sqlParameter.Add("sProcessID", Main.ProcessID);
                //sqlParameter.Add("sMachineID", Main.MachineID);
                sqlParameter.Add("CreateUserID", MainWindow.CurrentUser);

                Procedure pro3 = new Procedure();
                pro3.Name = "xp_McCode_iuMappingMachineCode_WPF";
                pro3.OutputUseYN = "N";
                pro3.OutputName = "sNewMCID";
                pro3.OutputLength = "10";

                Prolist.Add(pro3);
                ListParameter.Add(sqlParameter);
                //}
                //}

                string[] confirm = new string[2];
                confirm = DataStore.Instance.ExecuteAllProcedureOutputNew(Prolist, ListParameter);

                if (confirm[0] == "success")
                {
                    //MessageBox.Show("성공");
                    //flag = true;
                }
                else
                {
                    MessageBox.Show("실패 : " + confirm[1]);
                    //flag = false;
                }
            }
            catch (Exception ex) // 뭐든 간에 파일 없다고 하자
            {
                MessageBox.Show("저장실패 : " + ex.Message);
                return;
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

        // 확장자 이미지 확인하기, 메인윈도우에 확장자 리스트 세팅
        private bool CheckImage(string ImageName)
        {
            string[] extensions = MainWindow.Extensions;

            bool flag = false;

            ImageName = ImageName.Trim().ToLower();
            foreach (string ext in extensions)
            {
                if (ImageName.EndsWith(ext))
                {
                    flag = true;
                }
            }

            return flag;
        }


        #endregion



        private void lblchkNotUse_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            if (chkNotUse.IsChecked == true)
            {
                chkNotUse.IsChecked = false;
            }
            else
            {
                chkNotUse.IsChecked = true;
            }
        }

        private void lblNoUse_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            if (chkNoUse.IsChecked == true)
            {
                chkNoUse.IsChecked = false;
            }
            else
            {
                chkNoUse.IsChecked = true;
            }
        }

        private void chkNoUse_Checked(object sender, RoutedEventArgs e)
        {
            chkNoUse.IsChecked = true;
        }

        private void chkNoUse_UnChecked(object sender, RoutedEventArgs e)
        {
            chkNoUse.IsChecked = false;
        }

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
    }

    #region CoewView
    class Win_prd_MCCode_U_CodeView : BaseView
    {
        public int Num { get; set; }
        public string mcid { get; set; }
        public string mcname { get; set; }
        public string managerid { get; set; }
        public string customid { get; set; }

        public string customname { get; set; }
        public string buycustomid { get; set; }
        public string buycustomname { get; set; }
        public string personid { get; set; }
        public string personname { get; set; }

        public string buydate { get; set; }
        public string useyear { get; set; }
        public string SetHitQty { get; set; }
        public string AfterRepairHitcount { get; set; }
        public string HitCount { get; set; }

        public string ProcessID { get; set; }
        public string MachineID { get; set; }

        public string LastChangeDate { get; set; }
        public string LastChangeDate_CV { get; set; }

        public string ImageFile { get; set; }
        public string ImagePath { get; set; }

        public string Spec { get; set; }
        public string ProductionDate { get; set; }
        public string ModelName { get; set; }
        public string ProductionDate_CV { get; set; }

        public string HrLicenceID { get; set; }
        public string HrLicenceName { get; set; }
        public string buydate_CV { get; set; }
        public BitmapImage ImageView { get; set; }
        public string UseClss { get; set; } //2021-11-15


    }

    class Win_prd_MCCode_U_MachineMapping_CodeView
    {
        public int Num { get; set; }
        public string MCID { get; set; }
        public string ProcessID { get; set; } //숫자
        public string MachineID { get; set; } //숫자
        public string Process { get; set; }   //한글
        public string Machine { get; set; }   //한글
        public string MachineNo { get; set; }   //한글
    }

    class Win_prd_MCCode_U_Sub_CodeView : BaseView
    {
        public int Num { get; set; }
        public string McID { get; set; }
        public string McPartID { get; set; }
        public string ChangeCheckGbn { get; set; }
        public string CycleProdQty { get; set; }

        public string StartSetProdQty { get; set; }
        public string StartSetDate { get; set; }
        public string MCPartName { get; set; }
    }

    #endregion
}

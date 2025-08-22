using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Net;
using System.Text;
using System.Threading;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media.Imaging;
using WizMes_WellMade.PopUp;
using WizMes_WellMade.PopUP;
using WPF.MDI;

namespace WizMes_WellMade
{
    /// <summary>
    /// Win_prd_RegularInspect_U.xaml에 대한 상호 작용 논리
    /// </summary>
    public partial class Win_prd_RegularInspect_U : UserControl
    {
        string stDate = string.Empty;
        string stTime = string.Empty;

        string strFlag = string.Empty;
        List<string> lstStringBasis = new List<string>();
        int rowNum = 0;
        string strBasisID = string.Empty;
        int dgdInComBoNum = 0;
        Win_prd_RegularInspect_U_CodeView WinMcRegulInspect = new Win_prd_RegularInspect_U_CodeView();
        Win_prd_RegularInspect_U_Sub_CodeView WinMcRegulInspectSub = new Win_prd_RegularInspect_U_Sub_CodeView();
        Lib lib = new Lib();

        // FTP 활용모음.
        string strImagePath = string.Empty;
        string strFullPath = string.Empty;

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

        //string FTP_ADDRESS = "ftp://wizis.iptime.org/ImageData/McReqularInspect";
        //string FTP_ADDRESS = "ftp://wizis.iptime.org/ImageData/Draw";
        //string FTP_ADDRESS = "ftp://222.104.222.145:25000/ImageData/McReqularInspect";
        //string FTP_ADDRESS = "ftp://192.168.0.95/ImageData/McReqularInspect";
        string FTP_ADDRESS = "ftp://" + LoadINI.FileSvr + ":"
            + LoadINI.FTPPort + LoadINI.FtpImagePath + "/McRIB";
        private const string FTP_ID = "wizuser";
        private const string FTP_PASS = "wiz9999";
        private const string LOCAL_DOWN_PATH = "C:\\Temp";

        private List<string> lstMsg = new List<string>();
        private string message = "";

        public Win_prd_RegularInspect_U()
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

            chkInspectDateSrh.IsChecked = true;
            btnThisMonth_Click(null, null);
        }

        private void SetComboBox()
        {
            ObservableCollection<CodeView> ovcMcInsCycleGbnSrh = ComboBoxUtil.Instance.Gf_DB_CM_GetComCodeDataset(null, "MCCYCLEGBN", "Y", "", "");
            this.cboMcInsCycleGbnSrh.ItemsSource = ovcMcInsCycleGbnSrh;
            this.cboMcInsCycleGbnSrh.DisplayMemberPath = "code_name";
            this.cboMcInsCycleGbnSrh.SelectedValuePath = "code_id";

            this.cboMcInsCycleGbn.ItemsSource = ovcMcInsCycleGbnSrh;
            this.cboMcInsCycleGbn.DisplayMemberPath = "code_name";
            this.cboMcInsCycleGbn.SelectedValuePath = "code_id";
        }

        #region 체크박스 in 라벨 & PlusFinder

        //검사일자 라벨 클릭시
        private void lblInspectDateSrh_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            if (chkInspectDateSrh.IsChecked == true) { chkInspectDateSrh.IsChecked = false; }
            else { chkInspectDateSrh.IsChecked = true; }
        }

        //검사일자 라벨 in 체크박스 체크시
        private void chkInspectDateSrh_Checked(object sender, RoutedEventArgs e)
        {
            dtpSDate.IsEnabled = true;
            dtpEDate.IsEnabled = true;
            btnThisMonth.IsEnabled = true;
            btnToday.IsEnabled = true;
        }

        //검사일자 라벨 in 체크박스 언체크시
        private void chkInspectDateSrh_Unchecked(object sender, RoutedEventArgs e)
        {
            dtpSDate.IsEnabled = false;
            dtpEDate.IsEnabled = false;
            btnThisMonth.IsEnabled = false;
            btnToday.IsEnabled = false;
        }

        //금월 버튼 클릭시
        private void btnThisMonth_Click(object sender, RoutedEventArgs e)
        {
            dtpSDate.SelectedDate = Lib.Instance.BringThisMonthDatetimeList()[0];
            dtpEDate.SelectedDate = Lib.Instance.BringThisMonthDatetimeList()[1];
        }

        //금일 버튼 클릭시
        private void btnToday_Click(object sender, RoutedEventArgs e)
        {
            dtpSDate.SelectedDate = DateTime.Today;
            dtpEDate.SelectedDate = DateTime.Today;
        }

        //설비명 라벨 클릭시
        private void lblMcSrh_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            if (chkMcSrh.IsChecked == true) { chkMcSrh.IsChecked = false; }
            else { chkMcSrh.IsChecked = true; }
        }

        //설비명 라벨 in 체크박스 체크시
        private void chkMcSrh_Checked(object sender, RoutedEventArgs e)
        {
            txtMcSrh.IsEnabled = true;
            btnPfMcSrh.IsEnabled = true;
        }

        //설비명 라벨 in 체크박스 언체크시
        private void chkMcSrh_Unchecked(object sender, RoutedEventArgs e)
        {
            txtMcSrh.IsEnabled = false;
            btnPfMcSrh.IsEnabled = false;
        }

        //설비명 엔터키 이벤트용(상단)
        private void txtMcSrh_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                MainWindow.pf.ReturnCode(txtMcSrh, 12, "");
            }
        }

        //설비명 버튼 클릭 이벤트용(상단)
        private void btnPfMcSrh_Click(object sender, RoutedEventArgs e)
        {
            MainWindow.pf.ReturnCode(txtMcSrh, 12, "");
        }

        //정기검사구분 라벨 클릭시
        private void lblMcInsCycleGbnSrh_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            if (chkMcInsCycleGbnSrh.IsChecked == true) { chkMcInsCycleGbnSrh.IsChecked = false; }
            else { chkMcInsCycleGbnSrh.IsChecked = true; }
        }

        //정기검사구분 라벨 in 체크박스 체크시
        private void chkMcInsCycleGbnSrh_Checked(object sender, RoutedEventArgs e)
        {
            cboMcInsCycleGbnSrh.IsEnabled = true;
        }

        //정기검사구분 라벨 in 체크박스 언체크시
        private void chkMcInsCycleGbnSrh_Unchecked(object sender, RoutedEventArgs e)
        {
            cboMcInsCycleGbnSrh.IsEnabled = false;
        }

        #endregion

        /// <summary>
        /// 수정,추가 저장 후
        /// </summary>
        private void CanBtnControl()
        {
            btnAdd.IsEnabled = true;
            btnUpdate.IsEnabled = true;
            btnDelete.IsEnabled = true;
            btnSearch.IsEnabled = true;
            btnSave.Visibility = Visibility.Hidden;
            btnCancel.Visibility = Visibility.Hidden;
            btnExcel.Visibility = Visibility.Visible;
            gbxMcPart.IsEnabled = false;
            lblMsg.Visibility = Visibility.Hidden;
            //dgdInspect.IsEnabled = true;
            dgdInspect.IsHitTestVisible = true;
        }

        /// <summary>
        /// 수정,추가 진행 중
        /// </summary>
        private void CantBtnControl()
        {
            btnAdd.IsEnabled = false;
            btnUpdate.IsEnabled = false;
            btnDelete.IsEnabled = false;
            btnSearch.IsEnabled = false;
            btnSave.Visibility = Visibility.Visible;
            btnCancel.Visibility = Visibility.Visible;
            btnExcel.Visibility = Visibility.Hidden;
            gbxMcPart.IsEnabled = true;
            lblMsg.Visibility = Visibility.Visible;
            //dgdInspect.IsEnabled = false;
            dgdInspect.IsHitTestVisible = false;

        }

        #region 오른 상단 버튼 동작

        //추가
        private void btnAdd_Click(object sender, RoutedEventArgs e)
        {
            CantBtnControl();
            strFlag = "I";

            if (dgdInspectSub1.Items.Count > 0)
            {
                dgdInspectSub1.Items.Clear();
                dgdInspectSub1.Refresh();
            }
            if (dgdInspectSub2.Items.Count > 0)
            {
                dgdInspectSub2.Items.Clear();
                dgdInspectSub2.Refresh();
            }

            lblMsg.Visibility = Visibility.Visible;
            tbkMsg.Text = "자료 입력 중";
            rowNum = dgdInspect.SelectedIndex;
            this.DataContext = null;
            //검사일자 오늘날짜 들어가도록
            dtpInspectDate.SelectedDate = DateTime.Now;
            //검사자는 로그인한 사람이 자동으로 들어가도록
            txtMcRInspectPersonID.Text = MainWindow.CurrentPerson;
            txtMcRInspectPersonID.Tag = MainWindow.CurrentPersonID;

            txtMc.IsEnabled = true;
            btnPfMc.IsEnabled = true;

            //추가버튼 누르면 설비명에 포커스 이동
            txtMc.Focus();
        }

        //수정
        private void btnUpdate_Click(object sender, RoutedEventArgs e)
        {
            WinMcRegulInspect = dgdInspect.SelectedItem as Win_prd_RegularInspect_U_CodeView;

            if (WinMcRegulInspect != null)
            {
                rowNum = dgdInspect.SelectedIndex;
                tbkMsg.Text = "자료 수정 중";
                lblMsg.Visibility = Visibility.Visible;
                CantBtnControl();

                txtMc.IsEnabled = false;
                btnPfMc.IsEnabled = false;

                strFlag = "U";
            }


            

        }

        //삭제
        private void btnDelete_Click(object sender, RoutedEventArgs e)
        {
            WinMcRegulInspect = dgdInspect.SelectedItem as Win_prd_RegularInspect_U_CodeView;

            if (WinMcRegulInspect == null)
            {
                MessageBox.Show("삭제할 데이터가 지정되지 않았습니다. 삭제데이터를 지정하고 눌러주세요");
            }
            else
            {
                if (MessageBox.Show("선택하신 항목을 삭제하시겠습니까?", "삭제 전 확인", MessageBoxButton.YesNo) == MessageBoxResult.Yes)
                {
                    if (dgdInspect.Items.Count > 0 && dgdInspect.SelectedItem != null)
                    {
                        rowNum = dgdInspect.SelectedIndex;
                    }

                    if (DeleteData(WinMcRegulInspect.McRInspectID))
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
                try
                {
                    int rowNum = 0;
                    using (Loading lw = new Loading(FillGrid))
                    {
                        lw.ShowDialog();
                        if (dgdInspect.Items.Count <= 0)
                        {
                            MessageBox.Show("조회된 내용이 없습니다.");
                        }
                        else
                        {
                            dgdInspect.SelectedIndex = rowNum;
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
            if (SaveData(strFlag, txtMcRInspectID.Text))
            {
                CanBtnControl();

                lblMsg.Visibility = Visibility.Hidden;
                //rowNum = 0;

                if (strFlag.Equals("I"))
                {
                   rowNum = dgdInspect.Items.Count;
                   re_Search(rowNum);
                }
                else
                {
                    re_Search(rowNum);
                }
                               
                strBasisID = string.Empty;
            }
        }

        //취소
        private void btnCancel_Click(object sender, RoutedEventArgs e)
        {
            CanBtnControl();
            strBasisID = string.Empty;
            if (!strFlag.Equals(string.Empty))
            {
                re_Search(rowNum);
            }

            strFlag = string.Empty;
        }

        //엑셀
        private void btnExcel_Click(object sender, RoutedEventArgs e)
        {
            DataTable dt = null;
            string Name = string.Empty;

            string[] lst = new string[6];
            lst[0] = "설비점검";
            lst[1] = "설비점검 검사 범례";
            lst[2] = "설비점검 검사 수치";
            lst[3] = dgdInspect.Name;
            lst[4] = dgdInspectSub1.Name;
            lst[5] = dgdInspectSub2.Name;

            ExportExcelxaml ExpExc = new ExportExcelxaml(lst);
            ExpExc.ShowDialog();

            if (ExpExc.DialogResult.HasValue)
            {
                DataStore.Instance.InsertLogByForm(this.GetType().Name, "E");
                if (ExpExc.choice.Equals(dgdInspect.Name))
                {
                    if (ExpExc.Check.Equals("Y"))
                        dt = Lib.Instance.DataGridToDTinHidden(dgdInspect);
                    else
                        dt = Lib.Instance.DataGirdToDataTable(dgdInspect);

                    Name = dgdInspect.Name;
                    Lib.Instance.GenerateExcel(dt, Name);
                    Lib.Instance.excel.Visible = true;
                }
                else if (ExpExc.choice.Equals(dgdInspectSub1.Name))
                {
                    if (ExpExc.Check.Equals("Y"))
                        dt = Lib.Instance.DataGridToDTinHidden(dgdInspectSub1);
                    else
                        dt = Lib.Instance.DataGirdToDataTable(dgdInspectSub1);

                    Name = dgdInspectSub1.Name;
                    Lib.Instance.GenerateExcel(dt, Name);
                    Lib.Instance.excel.Visible = true;
                }
                else if (ExpExc.choice.Equals(dgdInspectSub2.Name))
                {
                    if (ExpExc.Check.Equals("Y"))
                        dt = Lib.Instance.DataGridToDTinHidden(dgdInspectSub2);
                    else
                        dt = Lib.Instance.DataGirdToDataTable(dgdInspectSub2);

                    Name = dgdInspectSub2.Name;
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

        /// <summary>
        /// 재검색(수정,삭제,추가 저장후에 자동 재검색)
        /// </summary>
        /// <param name="selectedIndex"></param>
        private void re_Search(int selectedIndex)
        {
            if (dgdInspect.Items.Count > 0)
            {
                dgdInspect.Items.Clear();
            }

            FillGrid();

            if (dgdInspect.Items.Count > 0)
            {
                dgdInspect.SelectedIndex = selectedIndex;
            }
            else
            {
                this.DataContext = null;
                dgdInspectSub1.Items.Clear();
                dgdInspectSub2.Items.Clear();

                MessageBox.Show("조회된 데이터가 없습니다.");
                return;
            }
        }

        /// <summary>
        /// 실조회
        /// </summary>
        private void FillGrid()
        {
            try
            {
                if(dgdInspect.Items.Count > 0)
                {
                    dgdInspect.Items.Clear();
                }

                DataSet ds = null;
                Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                sqlParameter.Clear();
                sqlParameter.Add("ChkMcRInspectDate", chkInspectDateSrh.IsChecked == true ? 1 : 0);
                sqlParameter.Add("sFromDate", chkInspectDateSrh.IsChecked == true ?
                    dtpSDate.SelectedDate.Value.ToString("yyyyMMdd") : "");
                sqlParameter.Add("sToDate", chkInspectDateSrh.IsChecked == true ?
                    dtpEDate.SelectedDate.Value.ToString("yyyyMMdd") : "");
                sqlParameter.Add("ChkMcID", chkMcSrh.IsChecked == true ? 1 : 0);
                sqlParameter.Add("MCID", chkMcSrh.IsChecked == true ? txtMcSrh.Tag.ToString() : "");
                sqlParameter.Add("ChkInsCycleGbn", chkMcInsCycleGbnSrh.IsChecked == true ? 1 : 0);
                sqlParameter.Add("InsCycleGbn", chkMcInsCycleGbnSrh.IsChecked == true ?
                    cboMcInsCycleGbnSrh.SelectedValue.ToString() : "");
                ds = DataStore.Instance.ProcedureToDataSet_LogWrite("xp_McReqularInspect_sMcReqularInspect", sqlParameter, true, "R");

                if (ds != null && ds.Tables.Count > 0)
                {
                    DataTable dt = ds.Tables[0];
                    int i = 0;

                    if (dt.Rows.Count == 0)
                    {
                        MessageBox.Show("조회된 데이터가 없습니다.");
                    }
                    else
                    {
                        DataRowCollection drc = dt.Rows;

                        foreach (DataRow dr in drc)
                        {
                            var WinMCRegulInspect = new Win_prd_RegularInspect_U_CodeView()
                            {
                                Num = i + 1,
                                McInspectBasisID = dr["McInspectBasisID"].ToString(),
                                MCID = dr["MCID"].ToString(),
                                MCNAME = dr["MCNAME"].ToString(),
                                managerid = dr["managerid"].ToString(),
                                McInsBasisDate = dr["McInsBasisDate"].ToString(),
                                McInsContent = dr["McInsContent"].ToString(),
                                BasisComments = dr["BasisComments"].ToString(),
                                McRInspectID = dr["McRInspectID"].ToString(),
                                McRInspectDate = dr["McRInspectDate"].ToString(),
                                McInsCycleGbn = dr["McInsCycleGbn"].ToString(),
                                McInsCycle = dr["McInsCycle"].ToString(),
                                Name = dr["Name"].ToString(),
                                //McRInspectUserID = dr["McRInspectUserID"].ToString(),
                                DefectContents = dr["DefectContents"].ToString(),
                                DefectReason = dr["DefectReason"].ToString(),
                                DefectRespectContents = dr["DefectRespectContents"].ToString(),
                                Comments = dr["Comments"].ToString()
                            };

                            if (WinMCRegulInspect.McInsBasisDate != null &&
                                !WinMCRegulInspect.McInsBasisDate.Replace(" ", "").Equals(""))
                            {
                                WinMCRegulInspect.McInsBasisDate_Convert =
                                //Lib.Instance.strConvertDate(WinMCRegulInspect.McInsBasisDate);
                                Lib.Instance.StrDateTimeBar(WinMCRegulInspect.McInsBasisDate);
                            }
                            if (WinMCRegulInspect.McRInspectDate != null &&
                                !WinMCRegulInspect.McRInspectDate.Replace(" ", "").Equals(""))
                            {
                                WinMCRegulInspect.McRInspectDate_Convert =
                                //Lib.Instance.strConvertDate(WinMCRegulInspect.McRInspectDate);
                                Lib.Instance.StrDateTimeBar(WinMCRegulInspect.McRInspectDate);
                            }

                            dgdInspect.Items.Add(WinMCRegulInspect);
                            i++;
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

        //설비점검 메인그리드의 행 선택시
        private void dgdInspect_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            WinMcRegulInspect = dgdInspect.SelectedItem as Win_prd_RegularInspect_U_CodeView;

            if (WinMcRegulInspect != null)
            {
                this.DataContext = WinMcRegulInspect;
                FillGridSub(WinMcRegulInspect.McRInspectID, WinMcRegulInspect.McInspectBasisID
                    , WinMcRegulInspect.McInsCycleGbn);
            }
        }

        private BitmapImage SetImage(string strAttachPath, string ImgName)
        {
            _ftp = new FTP_EX(FTP_ADDRESS, FTP_ID, FTP_PASS);
            BitmapImage bit = DrawingImageByByte(FTP_ADDRESS + strAttachPath + "", ImgName);
            //image.Source = bit;
            return bit;
        }

        #region 이미지 다운로드 및 이미지 반환
        /// <summary>
        /// ftp경로를 가지고 Bitmap 정보 리턴한다
        /// </summary>
        /// <param name="ftpFilePath"></param>
        /// <returns></returns>
        private BitmapImage DrawingImageByByte(string ftpFilePath, string ImgName)
        {
            BitmapImage image = new BitmapImage();

            try
            {
                WebClient ftpClient = new WebClient();
                ftpClient.Credentials = new NetworkCredential(FTP_ID, FTP_PASS);
                byte[] imageByte = ftpClient.DownloadData(ftpFilePath);

                //MemoryStream mStream = new MemoryStream();
                //mStream.Write(imageByte, 0, Convert.ToInt32(imageByte.Length));

                using (MemoryStream stream = new MemoryStream(imageByte))
                {
                    image.BeginInit();
                    image.StreamSource = stream;
                    image.CacheOption = BitmapCacheOption.OnLoad;
                    image.EndInit();
                    image.Freeze();
                }

            }
            catch (Exception ex)
            {
                //System.Windows.MessageBox.Show("1" + ex.Message + " / " + ex.Source);
                //throw ex;

                if (message.Trim().Equals(""))
                {
                    message += ImgName;
                    lstMsg.Add(ImgName);
                }
                else
                {
                    if (!lstMsg.Contains(ImgName))
                    {
                        message += ", " + ImgName;
                        lstMsg.Add(ImgName);
                    }
                }
            }

            return image;
        }

        #endregion // 이미지 다운로드 및 이미지 반환

        /// <summary>
        /// 서브 조회
        /// </summary>
        /// <param name="strID"></param>
        private void FillGridSub(string strInsID, string strBasisID, string strCycleGbn)
        {
            message = "";
            lstMsg.Clear();

            if (dgdInspectSub1.Items.Count > 0)
            {
                dgdInspectSub1.Items.Clear();
                dgdInspectSub1.Refresh();
            }
            if (dgdInspectSub2.Items.Count > 0)
            {
                dgdInspectSub2.Items.Clear();
                dgdInspectSub2.Refresh();
            }

            try
            {
                DataSet ds = null;
                Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                sqlParameter.Clear();
                sqlParameter.Add("McRInspectID", strInsID);
                sqlParameter.Add("McInspectBasisID", strBasisID);
                sqlParameter.Add("McInsCycleGbn", strCycleGbn);
                ds = DataStore.Instance.ProcedureToDataSet("xp_McRegularInspect_sMcRegularInspectSub", sqlParameter, false);

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
                            var WinMCRegulInsSub = new Win_prd_RegularInspect_U_Sub_CodeView()
                            {
                                Num = i + 1,
                                McRInspectID = dr["McRInspectID"].ToString(),
                                McRSeq = dr["McRSeq"].ToString(),
                                McRInspectLegend = dr["McRInspectLegend"].ToString(),
                                McRInspectValue = stringFormatN0(dr["McRInspectValue"]),
                                McInspectBasisID = dr["McInspectBasisID"].ToString(),
                                McSeq = dr["McSeq"].ToString(),
                                McInsCheck = dr["McInsCheck"].ToString(),
                                McInsCycle = dr["McInsCycle"].ToString(),
                                McInsRecord = dr["McInsRecord"].ToString(),
                                McInsRecordGbn = dr["McInsRecordGbn"].ToString(),
                                McInsItemName = dr["McInsItemName"].ToString(),
                                McInsContent = dr["McInsContent"].ToString(),
                                McInsCycleGbn = dr["McInsCycleGbn"].ToString(),
                                Legend = dr["Legend"].ToString(),
                                McImagePath = dr["McImagePath"].ToString(),
                                McImageFile = dr["McImageFile"].ToString()
                            };

                            if (WinMCRegulInsSub.McImageFile != null && !WinMCRegulInsSub.McImageFile.Replace(" ", "").Equals(""))
                            {
                                WinMCRegulInsSub.imageFlag = true;

                                if (CheckImage(WinMCRegulInsSub.McImageFile.Trim()))
                                {
                                    string strImage = "/" + WinMCRegulInsSub.McInspectBasisID + "/" + WinMCRegulInsSub.McImageFile;
                                    WinMCRegulInsSub.ImageView = SetImage(strImage, WinMCRegulInsSub.McImageFile);
                                }
                                else
                                {
                                    MessageBox.Show(WinMCRegulInsSub.McImageFile + "는 이미지 변환이 불가능합니다.");
                                }
                            }
                            else
                            {
                                WinMCRegulInsSub.imageFlag = false;
                            }

                            ObservableCollection<CodeView> ovcMcRInspectLegend =
                                ComboBoxUtil.Instance.Gf_DB_CM_GetComCodeDataset(null, "MCLEGEND", "Y", "", "");

                            foreach (CodeView cv in ovcMcRInspectLegend)
                            {
                                if (cv.code_id == WinMCRegulInsSub.McRInspectLegend)
                                {
                                    WinMCRegulInsSub.LegendShape = cv.code_name;
                                    break;
                                }
                            }

                            if (WinMCRegulInsSub != null)
                            {
                                if (WinMCRegulInsSub.McInsRecordGbn.Equals("1"))
                                {
                                    dgdInspectSub1.Items.Add(WinMCRegulInsSub);
                                }
                                else if (WinMCRegulInsSub.McInsRecordGbn.Equals("2"))
                                {
                                    dgdInspectSub2.Items.Add(WinMCRegulInsSub);
                                }
                            }

                            i++;
                        }
                    }

                    if (!message.Trim().Equals(""))
                    {
                        MessageBox.Show(message + " 를 불러올 수 없습니다.");
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

        /// <summary>
        /// 설비기준ID 와 정기검사구분으로 검사할 항목 검색
        /// </summary>
        /// <param name="strBasisID"></param>
        /// <param name="strCycleGbn"></param>
        private void FillGridSubNoResult(string strBasisID, string strCycleGbn)
        {
            message = "";
            lstMsg.Clear();

            if (dgdInspectSub1.Items.Count > 0)
            {
                dgdInspectSub1.Items.Clear();
            }
            if (dgdInspectSub2.Items.Count > 0)
            {
                dgdInspectSub2.Items.Clear();
            }

            try
            {
                Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                sqlParameter.Add("McInspectBasisID", strBasisID);
                sqlParameter.Add("McSeq", 0);
                sqlParameter.Add("McInsCycleGbn", strCycleGbn);

                DataSet ds = DataStore.Instance.ProcedureToDataSet("xp_McReqularInspectBasis_sMcReqularInspectBasisSub", sqlParameter, false);

                if (ds != null && ds.Tables.Count > 0)
                {
                    DataTable dt = ds.Tables[0];

                    if (dt.Rows.Count > 0)
                    {
                        DataRowCollection drc = dt.Rows;

                        foreach (DataRow dr in drc)
                        {
                            var WinSub = new Win_prd_RegularInspect_U_Sub_CodeView()
                            {
                                McInsItemName = dr["McInsItemName"].ToString(),
                                McInsContent = dr["McInsContent"].ToString(),
                                McInsCheck = dr["McInsCheck"].ToString(),
                                McInsCycle = dr["McInsCycle"].ToString(),
                                McInsRecord = dr["McInsRecord"].ToString(),
                                McInspectBasisID = dr["McInspectBasisID"].ToString(),
                                McSeq = dr["McSeq"].ToString(),
                                McInsCycleGbn = dr["McInsCycleGbn"].ToString(),
                                McInsRecordGbn = dr["McInsRecordGbn"].ToString(),
                                McImagePath = dr["McImagePath"].ToString(),
                                McImageFile = dr["McImageFile"].ToString(),
                                flagBool = false
                            };

                            if (WinSub.McImageFile != null && !WinSub.McImageFile.Replace(" ", "").Equals(""))
                            {
                                WinSub.imageFlag = true;
                                
                                if (CheckImage(WinSub.McImageFile.Trim()))
                                {
                                    string strImage = "/" + WinSub.McInspectBasisID + "/" + WinSub.McImageFile;
                                    WinSub.ImageView = SetImage(strImage, WinSub.McImageFile);
                                }
                                else
                                {
                                    MessageBox.Show(WinSub.McImageFile + "는 이미지 변환이 불가능합니다.");
                                }
                            }
                            else
                            {
                                WinSub.imageFlag = false;
                            }

                            if (WinSub.McInsRecordGbn.Equals("1"))
                            {
                                dgdInspectSub1.Items.Add(WinSub);
                            }

                            if (WinSub.McInsRecordGbn.Equals("2"))
                            {
                                dgdInspectSub2.Items.Add(WinSub);
                            }
                        }
                    }

                    if (!message.Trim().Equals(""))
                    {
                        MessageBox.Show(message + " 를 불러올 수 없습니다.");
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
                sqlParameter.Add("McRInspectID", strID);

                string[] result = DataStore.Instance.ExecuteProcedure_NewLog("xp_McRegularInspect_dMcRegularInspect", sqlParameter, "D");

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

            try
            {
                if (CheckData())
                {
                    Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                    sqlParameter.Clear();
                    sqlParameter.Add("McRInspectID", strID);
                    sqlParameter.Add("McRInspectDate", dtpInspectDate.SelectedDate.Value.ToString("yyyyMMdd"));
                    sqlParameter.Add("McRInspectUserID", txtMcRInspectPersonID.Tag.ToString());
                    sqlParameter.Add("McInspectBasisID", strFlag.Equals("I") == true ? strBasisID : WinMcRegulInspect.McInspectBasisID);
                    sqlParameter.Add("McInsBasisDate", dtpMcInsBasisDate.SelectedDate.Value.ToString("yyyyMMdd"));

                    sqlParameter.Add("McInsCycleGbn", cboMcInsCycleGbn.SelectedValue.ToString());
                    sqlParameter.Add("DefectContents", txtDefectContents.Text);
                    sqlParameter.Add("DefectReason", txtDefectReason.Text);
                    sqlParameter.Add("DefectRespectContents", txtDefectRespectContents.Text);
                    sqlParameter.Add("Comments", txtComments.Text);


                    #region 추가

                    if (strFlag.Equals("I"))
                    {
                        sqlParameter.Add("CreateUserID", MainWindow.CurrentUser);

                        Procedure pro1 = new Procedure();
                        pro1.Name = "xp_McRegularInspect_iMcRegularInspect";
                        pro1.OutputUseYN = "Y";
                        pro1.OutputName = "McRInspectID";
                        pro1.OutputLength = "10";

                        Prolist.Add(pro1);
                        ListParameter.Add(sqlParameter);

                        for (int i = 0; i < dgdInspectSub1.Items.Count; i++)
                        {
                            WinMcRegulInspectSub = dgdInspectSub1.Items[i] as Win_prd_RegularInspect_U_Sub_CodeView;
                            sqlParameter = new Dictionary<string, object>();
                            sqlParameter.Clear();
                            sqlParameter.Add("McRInspectID", strID);
                            sqlParameter.Add("McInspectBasisID", WinMcRegulInspectSub.McInspectBasisID);
                            sqlParameter.Add("McInspectBasisSeq", WinMcRegulInspectSub.McSeq);
                            sqlParameter.Add("McRInspectValue", WinMcRegulInspectSub.McRInspectValue != null ?
                                double.Parse(WinMcRegulInspectSub.McRInspectValue.Replace(",", "")) : 0.0);
                            sqlParameter.Add("McRInspectLegend", WinMcRegulInspectSub.McRInspectLegend != null ?
                                WinMcRegulInspectSub.McRInspectLegend : "");
                            sqlParameter.Add("CreateUserID", MainWindow.CurrentUser);

                            Procedure pro2 = new Procedure();
                            pro2.Name = "xp_McRegularInspect_iMcRegularInspectSub";
                            pro2.OutputUseYN = "N";
                            pro2.OutputName = "McRInspectID";
                            pro2.OutputLength = "10";

                            Prolist.Add(pro2);
                            ListParameter.Add(sqlParameter);
                        }

                        for (int i = 0; i < dgdInspectSub2.Items.Count; i++)
                        {
                            WinMcRegulInspectSub = dgdInspectSub2.Items[i] as Win_prd_RegularInspect_U_Sub_CodeView;
                            sqlParameter = new Dictionary<string, object>();
                            sqlParameter.Clear();
                            sqlParameter.Add("McRInspectID", strID);
                            sqlParameter.Add("McInspectBasisID", WinMcRegulInspectSub.McInspectBasisID);
                            sqlParameter.Add("McInspectBasisSeq", WinMcRegulInspectSub.McSeq);
                            sqlParameter.Add("McRInspectValue", WinMcRegulInspectSub.McRInspectValue != null ?
                                double.Parse(WinMcRegulInspectSub.McRInspectValue.Replace(",", "")) : 0.0);
                            sqlParameter.Add("McRInspectLegend", WinMcRegulInspectSub.McRInspectLegend != null ?
                                WinMcRegulInspectSub.McRInspectLegend : "");
                            sqlParameter.Add("CreateUserID", MainWindow.CurrentUser);

                            Procedure pro3 = new Procedure();
                            pro3.Name = "xp_McRegularInspect_iMcRegularInspectSub";
                            pro3.OutputUseYN = "N";
                            pro3.OutputName = "McRInspectID";
                            pro3.OutputLength = "10";

                            Prolist.Add(pro3);
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
                                if (kv.key == "McRInspectID")
                                {
                                    sGetID = kv.value;
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
                    }

                    #endregion

                    #region 수정

                    else if (strFlag.Equals("U"))
                    {
                        sqlParameter.Add("UpdateUserID", MainWindow.CurrentUser);

                        Procedure pro1 = new Procedure();
                        pro1.Name = "xp_McRegularInspect_uMcRegularInspect";
                        pro1.OutputUseYN = "N";
                        pro1.OutputName = "McRInspectID";
                        pro1.OutputLength = "10";

                        Prolist.Add(pro1);
                        ListParameter.Add(sqlParameter);

                        for (int i = 0; i < dgdInspectSub1.Items.Count; i++)
                        {
                            WinMcRegulInspectSub = dgdInspectSub1.Items[i] as Win_prd_RegularInspect_U_Sub_CodeView;
                            sqlParameter = new Dictionary<string, object>();
                            sqlParameter.Clear();
                            sqlParameter.Add("McRInspectID", strID);
                            sqlParameter.Add("McInspectBasisID", WinMcRegulInspectSub.McInspectBasisID);
                            sqlParameter.Add("McInspectBasisSeq", WinMcRegulInspectSub.McSeq);
                            sqlParameter.Add("McRInspectValue", WinMcRegulInspectSub.McRInspectValue != null ?
                                double.Parse(WinMcRegulInspectSub.McRInspectValue.Replace(",", "")) : 0.0);
                            sqlParameter.Add("McRInspectLegend", WinMcRegulInspectSub.McRInspectLegend != null ?
                                WinMcRegulInspectSub.McRInspectLegend : "");
                            sqlParameter.Add("CreateUserID", MainWindow.CurrentUser);

                            Procedure pro2 = new Procedure();
                            pro2.Name = "xp_McRegularInspect_iMcRegularInspectSub";
                            pro2.OutputUseYN = "N";
                            pro2.OutputName = "McRInspectID";
                            pro2.OutputLength = "10";

                            Prolist.Add(pro2);
                            ListParameter.Add(sqlParameter);
                        }

                        for (int i = 0; i < dgdInspectSub2.Items.Count; i++)
                        {
                            WinMcRegulInspectSub = dgdInspectSub2.Items[i] as Win_prd_RegularInspect_U_Sub_CodeView;
                            sqlParameter = new Dictionary<string, object>();
                            sqlParameter.Clear();
                            sqlParameter.Add("McRInspectID", strID);
                            sqlParameter.Add("McInspectBasisID", WinMcRegulInspectSub.McInspectBasisID);
                            sqlParameter.Add("McInspectBasisSeq", WinMcRegulInspectSub.McSeq);
                            sqlParameter.Add("McRInspectValue", WinMcRegulInspectSub.McRInspectValue != null ?
                                double.Parse(WinMcRegulInspectSub.McRInspectValue.Replace(",", "")) : 0.0);
                            sqlParameter.Add("McRInspectLegend", WinMcRegulInspectSub.McRInspectLegend != null ?
                                WinMcRegulInspectSub.McRInspectLegend : "");
                            sqlParameter.Add("CreateUserID", MainWindow.CurrentUser);

                            Procedure pro3 = new Procedure();
                            pro3.Name = "xp_McRegularInspect_iMcRegularInspectSub";
                            pro3.OutputUseYN = "N";
                            pro3.OutputName = "McRInspectID";
                            pro3.OutputLength = "10";

                            Prolist.Add(pro3);
                            ListParameter.Add(sqlParameter);
                        }

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

        /// <summary>
        /// 입력사항 체크
        /// </summary>
        /// <returns></returns>
        private bool CheckData()
        {
            bool flag = true;

            if (txtMc.Text.Length <= 0 || txtMc.Tag.ToString().Trim().Equals("") || txtMc.Tag == null)
            {
                MessageBox.Show("설비명이 입력되지 않았습니다.");
                flag = false;
                return flag;
            }

            if (dtpInspectDate.SelectedDate == null)
            {
                MessageBox.Show("검사일자가 선택되지 않았습니다.");
                flag = false;
                return flag;
            }

            if (cboMcInsCycleGbn.SelectedValue == null)
            {
                MessageBox.Show("정기검사구분이 선택되지 않았습니다.");
                flag = false;
                return flag;
            }

            if (string.IsNullOrWhiteSpace(txtMcRInspectPersonID.Tag?.ToString() ?? ""))
            {
                MessageBox.Show("검사자가 입력되지 않았습니다.");
                flag = false;
                return flag;
            }

            return flag;
        }

        //설비명 엔터키 이벤트용(입력)
        private void txtMc_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                MainWindow.pf.ReturnCode(txtMc, 94, "");

                try
                {
                    if (!txtMc.Text.Equals("") && txtMc.Tag.ToString() != null)
                    {
                        lstStringBasis.Clear();
                        lstStringBasis = AddSub(txtMc.Tag.ToString());
                        strBasisID = lstStringBasis[0];

                        if (!lstStringBasis[1].Replace(" ", "").Equals(""))
                        {
                            //개정일자
                            dtpMcInsBasisDate.SelectedDate = Lib.Instance.strConvertDate(lstStringBasis[1]);
                        }

                        if (cboMcInsCycleGbn.SelectedValue != null)
                        {
                            //정기검사구분
                            FillGridSubNoResult(strBasisID, cboMcInsCycleGbn.SelectedValue.ToString());
                        }
                    }
                }
                catch(Exception ex)
                {
                    MessageBox.Show(ex.ToString());
                }
                

                //검사일자 포커스 이동
                dtpInspectDate.Focus();

            }
        }

        //설비명 버튼 클릭 이벤트용(입력)
        private void btnPfMc_Click(object sender, RoutedEventArgs e)
        {
            MainWindow.pf.ReturnCode(txtMc, 94, "");

            try
            {
                if (!txtMc.Text.Equals("") && txtMc.Tag != null)
                {
                    lstStringBasis.Clear();
                    lstStringBasis = AddSub(txtMc.Tag.ToString());
                    strBasisID = lstStringBasis[0];

                    if (!lstStringBasis[1].Replace(" ", "").Equals(""))
                    {
                        dtpMcInsBasisDate.SelectedDate = Lib.Instance.strConvertDate(lstStringBasis[1]);
                    }

                    if (cboMcInsCycleGbn.SelectedValue != null)
                    {
                        FillGridSubNoResult(strBasisID, cboMcInsCycleGbn.SelectedValue.ToString());
                    }
                }

            }
            catch(Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
            
        }

        //정기검사구분 선택시
        private void cboMcInsCycleGbn_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (cboMcInsCycleGbn.SelectedValue != null)
            {
                if (!strBasisID.Equals(string.Empty))
                {
                    FillGridSubNoResult(strBasisID, cboMcInsCycleGbn.SelectedValue.ToString());
                }
            }
        }

        //검사자 엔터키 이벤트용(입력)
        private void txtMcRInspectPersonID_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                MainWindow.pf.ReturnCode(txtMcRInspectPersonID, 2, "");
            }
            
            //문제내역 포커스 이동
            txtDefectContents.Focus();
        }

        //검사자 버튼 클릭 이벤트용(입력)
        private void btnPfMcRInspectPersonID_Click(object sender, RoutedEventArgs e)
        {
            MainWindow.pf.ReturnCode(txtMcRInspectPersonID, 2, "");

            //문제내역 포커스 이동
            txtDefectContents.Focus();
        }

        /// <summary>
        /// 설비ID 로 설비기준ID와 설비기준생성일자 검색
        /// </summary>
        /// <param name="strMCID"></param>
        /// <returns></returns>
        private List<string> AddSub(string strMCID)
        {
            List<string> lstBasis = new List<string>();
            //string BasisID = string.Empty;
            //string BasisDate = string.Empty;
            Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
            sqlParameter.Add("MCID", strMCID);

            DataSet ds = DataStore.Instance.ProcedureToDataSet("xp_McRegularInspectBasis_sMcRegularInspectBasisByMcID", sqlParameter, false);


            if (ds != null && ds.Tables.Count > 0)
            {
                DataTable dt = ds.Tables[0];

                if (dt.Rows.Count > 0)
                {
                    DataRow dr = dt.Rows[0];

                    if (dt.Rows[0]["McInspectBasisID"] != null && !dt.Rows[0]["McInspectBasisID"].ToString().Trim().Equals(string.Empty))
                    {
                        //if (dt.Rows[0]["Article"] != null && !dt.Rows[0]["Article"].ToString().Trim().Equals(string.Empty))
                            //BasisID = dt.Rows[0]["McInspectBasisID"].ToString();
                            //BasisDate = dt.Rows[0]["McInsBasisDate"].ToString();
                        lstBasis.Add(dt.Rows[0]["McInspectBasisID"].ToString());
                        lstBasis.Add(dt.Rows[0]["McInsBasisDate"].ToString());
                    }
                }
            }

            return lstBasis;
        }

        private void DataGridCell_Sub1_KeyDown(object sender, KeyEventArgs e)
        {
            if (lblMsg.Visibility == Visibility.Visible)
            {
                WinMcRegulInspectSub = dgdInspectSub1.CurrentItem as Win_prd_RegularInspect_U_Sub_CodeView;
                int rowCount = dgdInspectSub1.Items.IndexOf(dgdInspectSub1.CurrentItem);
                int colCount = dgdInspectSub1.Columns.IndexOf(dgdtpeMcRInspectLegend);
                dgdInComBoNum = rowCount;

                if (e.Key == Key.Enter)
                {
                    e.Handled = true;
                    (sender as DataGridCell).IsEditing = false;

                    if (dgdInspectSub1.Items.Count > rowCount + 1)
                    {
                        //dgdInspectSub1.SelectedIndex = rowCount+1;
                        dgdInspectSub1.CurrentCell = new DataGridCellInfo
                            (dgdInspectSub1.Items[rowCount + 1], dgdInspectSub1.Columns[colCount]);
                    }
                    else
                    {
                        if (dgdInspectSub2.Items.Count > 0)
                        {
                            dgdInspectSub2.Focus();
                            //dgdInspectSub2.SelectedIndex = 0;
                            dgdInspectSub2.CurrentCell = new DataGridCellInfo(dgdInspectSub2.Items[0],
                                dgdInspectSub2.Columns[dgdInspectSub2.Columns.IndexOf(dgdtpeMcRInspectValue)]);
                        }
                        else
                        {
                            btnSave.Focus();
                        }
                    }

                    //if (dgdInspectSub1.Columns.Count == colCount  && dgdInspectSub1.Items.Count > rowCount)
                    //{
                    //    dgdInspectSub1.CurrentCell = new DataGridCellInfo(
                    //        dgdInspectSub1.Items[rowCount + 1], dgdInspectSub1.Columns[colCount - 1]);
                    //}
                    //else if (dgdInspectSub1.Columns.Count == colCount  && dgdInspectSub1.Items.Count == rowCount)
                    //{
                    //    btnSave.Focus();
                    //}
                }
            }
        }

        private void TextBoxFocusInDataGrid(object sender, KeyEventArgs e)
        {
            Lib.Instance.DataGridINControlFocus(sender, e);
        }

        private void TextBoxFocusInDataGrid_MouseUp(object sender, MouseButtonEventArgs e)
        {
            Lib.Instance.DataGridINBothByMouseUP(sender, e);
        }

        private void DataGridCell_GotFocus(object sender, RoutedEventArgs e)
        {
            if (lblMsg.Visibility == Visibility.Visible)
            {
                DataGridCell cell = sender as DataGridCell;
                cell.IsEditing = true;
            }
        }

        private void ComboMcRInspectLegend_Loaded(object sender, RoutedEventArgs e)
        {
            ComboBox cboMcRInspectLegend = (ComboBox)sender;
            //cboMcRInspectLegend.ItemsSource = null;

            ObservableCollection<CodeView> ovcMcRInspectLegend = ComboBoxUtil.Instance.
                Gf_DB_CM_GetComCodeDataset(null, "MCLEGEND", "Y", "", "");
            cboMcRInspectLegend.ItemsSource = ovcMcRInspectLegend;
            cboMcRInspectLegend.DisplayMemberPath = "code_name";
            cboMcRInspectLegend.SelectedValuePath = "code_id";
        }

        private void ComboMcRInspectLegend_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            WinMcRegulInspectSub = dgdInspectSub1.CurrentItem as Win_prd_RegularInspect_U_Sub_CodeView;

            ComboBox cboMcRInspectLegend = (ComboBox)sender;

            if (WinMcRegulInspectSub == null)
            {
                WinMcRegulInspectSub = this.dgdInspectSub1.Items[dgdInComBoNum] as Win_prd_RegularInspect_U_Sub_CodeView;
            }

            if (cboMcRInspectLegend.SelectedValue != null && !cboMcRInspectLegend.SelectedValue.ToString().Equals(""))
            {
                var theView = cboMcRInspectLegend.SelectedItem as CodeView;
                if (theView != null)
                {
                    WinMcRegulInspectSub.McRInspectLegend = theView.code_id;
                    WinMcRegulInspectSub.LegendShape = theView.code_name;
                }
                WinMcRegulInspectSub.McRInspectLegend = cboMcRInspectLegend.SelectedValue.ToString();
                sender = cboMcRInspectLegend;
            }
        }

        //dgdInspectSub1_dropdown 방지..
        private void ComboMcRInspectLegend_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                DataGridCell cell = Lib.Instance.GetParent<DataGridCell>(sender as ComboBox);
                cell.Focus();
            }
        }

        private void DataGridCell_Sub2_KeyDown(object sender, KeyEventArgs e)
        {
            if (lblMsg.Visibility == Visibility.Visible)
            {
                WinMcRegulInspectSub = dgdInspectSub1.CurrentItem as Win_prd_RegularInspect_U_Sub_CodeView;
                int rowCount = dgdInspectSub2.Items.IndexOf(dgdInspectSub2.CurrentItem);
                int colCount = dgdInspectSub2.Columns.IndexOf(dgdInspectSub2.CurrentCell.Column);

                if (e.Key == Key.Enter)
                {
                    e.Handled = true;
                    (sender as DataGridCell).IsEditing = false;

                    if (dgdInspectSub2.Items.Count > rowCount + 1)
                    {
                        //dgdInspectSub2.SelectedIndex = rowCount + 1;
                        dgdInspectSub2.CurrentCell = new DataGridCellInfo
                            (dgdInspectSub2.Items[rowCount + 1], dgdInspectSub2.Columns[colCount]);
                    }
                    else
                    {
                        btnSave.Focus();
                    }

                    //if (dgdInspectSub2.Columns.Count == colCount - 1 && dgdInspectSub2.Items.Count > rowCount)
                    //{
                    //    dgdInspectSub2.CurrentCell = new DataGridCellInfo(
                    //        dgdInspectSub2.Items[rowCount + 1], dgdInspectSub2.Columns[colCount - 1]);
                    //}
                    //else if (dgdInspectSub2.Columns.Count == colCount - 1 && dgdInspectSub2.Items.Count == rowCount)
                    //{
                    //    btnSave.Focus();
                    //}
                }
            }
        }

        private void dgdtpetxtMcRInspectValue_TextChanged(object sender, TextChangedEventArgs e)
        {
            if (lblMsg.Visibility == Visibility.Visible)
            {
                WinMcRegulInspectSub = dgdInspectSub2.CurrentItem as Win_prd_RegularInspect_U_Sub_CodeView;

                if (WinMcRegulInspectSub != null)
                {
                    TextBox tb1 = sender as TextBox;

                    if (Lib.Instance.IsNumOrAnother(tb1.Text))
                    {
                        WinMcRegulInspectSub.McRInspectValue = Lib.Instance.returnNumStringTwoExceptDot(tb1.Text);
                        tb1.SelectionStart = tb1.Text.Length;
                        sender = tb1;
                    }
                }
            }
        }

        private void ImagePop_Click(object sender, MouseButtonEventArgs e)
        {
            LargeImagePopUp lgp = new LargeImagePopUp(WinMcRegulInspectSub.ImageView);
            lgp.Show();
        }

        private void btnSeeImage_Click(object sender, RoutedEventArgs e)
        {
            DataGridCellInfo dgdinfoOne = dgdInspectSub1.CurrentCell;
            DataGridCellInfo dgdinfoTwo = dgdInspectSub2.CurrentCell;

            if (dgdinfoOne.Column != null)
            {
                WinMcRegulInspectSub = dgdInspectSub1.CurrentItem as Win_prd_RegularInspect_U_Sub_CodeView;
            }
            if (dgdinfoTwo.Column != null)
            {
                WinMcRegulInspectSub = dgdInspectSub2.CurrentItem as Win_prd_RegularInspect_U_Sub_CodeView;
            }

            if (WinMcRegulInspectSub != null && !WinMcRegulInspectSub.McImageFile.Equals(""))
            {
                FTP_DownLoadFile(WinMcRegulInspectSub.McImagePath + "/" + WinMcRegulInspectSub.McInspectBasisID + "/" + WinMcRegulInspectSub.McImageFile);
            }
        }

        //
        private void FTP_DownLoadFile(string strFilePath)
        {
            _ftp = new FTP_EX(FTP_ADDRESS, FTP_ID, FTP_PASS);

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
        
        private void btnGoBasis_Click(object sender, RoutedEventArgs e)
        {
            int k = 0;
            foreach (MenuViewModel mvm in MainWindow.mMenulist)
            {
                if (mvm.Menu.Equals("설비 점검 기준 등록"))
                {
                    break;
                }
                k++;
            }

            if (MainWindow.MainMdiContainer.Children.Contains(MainWindow.mMenulist[k].subProgramID as MdiChild))
            {
                (MainWindow.mMenulist[k].subProgramID as MdiChild).Focus();
            }
            else
            {
                Type type = Type.GetType("WizMes_WellMade." + MainWindow.mMenulist[k].ProgramID.Trim(), true);
                object uie = Activator.CreateInstance(type);

                MainWindow.mMenulist[k].subProgramID = new MdiChild()
                {
                    Title = "(주)WellMade [" + MainWindow.mMenulist[k].MenuID.Trim() + "] " + MainWindow.mMenulist[k].Menu.Trim() +
                            " (→" + MainWindow.mMenulist[k].ProgramID + ")",
                    Height = SystemParameters.PrimaryScreenHeight * 0.8,
                    MaxHeight = SystemParameters.PrimaryScreenHeight * 0.85,
                    Width = SystemParameters.WorkArea.Width * 0.85,
                    MaxWidth = SystemParameters.WorkArea.Width,
                    Content = uie as UIElement,
                    Tag = MainWindow.mMenulist[k]
                };
                Lib.Instance.AllMenuLogInsert(MainWindow.mMenulist[k].MenuID, MainWindow.mMenulist[k].Menu, MainWindow.mMenulist[k].subProgramID);
                MainWindow.MainMdiContainer.Children.Add(MainWindow.mMenulist[k].subProgramID as MdiChild);
            }
        }


        //검사일자 캘린더 열기
        private void DtpInspectDate_PreviewKeyDown(object sender, KeyEventArgs e)
        {
            if(e.Key == Key.Enter)
            {
                dtpInspectDate.IsDropDownOpen = true;
            }
        }

        //검사일자 -> 개정일자
        private void DtpInspectDate_CalendarClosed(object sender, RoutedEventArgs e)
        {
            dtpMcInsBasisDate.Focus();
        }

        //개정일자 캘린더 열기
        private void DtpMcInsBasisDate_PreviewKeyDown(object sender, KeyEventArgs e)
        {
            if(e.Key == Key.Enter)
            {
                dtpMcInsBasisDate.IsDropDownOpen = true;
            }
        }

        //개정일자 -> 정기검사구분
        private void DtpMcInsBasisDate_CalendarClosed(object sender, RoutedEventArgs e)
        {
            cboMcInsCycleGbn.IsDropDownOpen = true;

            cboMcInsCycleGbn.Focus();
        }
        
        //캘린더가 닫힐 때 검사자 텍스트 박스로 포커스 이동
        private void CboMcInsCycleGbn_DropDownClosed(object sender, EventArgs e)
        {
            txtMcRInspectPersonID.Focus();
        }

        //문제내역 -> 문제원인
        private void TxtDefectContents_KeyDown(object sender, KeyEventArgs e)
        {
            if(e.Key == Key.Enter)
            {
                txtDefectReason.Focus();
            }
        }

        //문제원인 -> 대책 및 조치
        private void TxtDefectReason_KeyDown(object sender, KeyEventArgs e)
        {
            if(e.Key == Key.Enter)
            {
                txtDefectRespectContents.Focus();

            }
        }

        //대책 및 조치 -> 비고
        private void TxtDefectRespectContents_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                txtComments.Focus();

            }
        }

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
    }

    class Win_prd_RegularInspect_U_CodeView : BaseView
    {
        public override string ToString()
        {
            return (this.ReportAllProperties());
        }

        public int Num { get; set; }
        public string McInspectBasisID { get; set; }
        public string MCID { get; set; }
        public string MCNAME { get; set; }
        public string managerid { get; set; }
        public string McInsBasisDate { get; set; }
        public string McInsContent { get; set; }
        public string BasisComments { get; set; }
        public string McRInspectID { get; set; }
        public string McRInspectDate { get; set; }
        public string McInsCycleGbn { get; set; }
        public string McInsCycle { get; set; }
        public string Name { get; set; }
        public string McRInspectUserID { get; set; }
        public string DefectContents { get; set; }
        public string DefectReason { get; set; }
        public string DefectRespectContents { get; set; }
        public string Comments { get; set; }

        //public DateTime McInsBasisDate_Convert { get; set; }
        //public DateTime McRInspectDate_Convert { get; set; }
        public string McInsBasisDate_Convert { get; set; }
        public string McRInspectDate_Convert { get; set; }
    }

    class Win_prd_RegularInspect_U_Sub_CodeView : BaseView
    {
        public override string ToString()
        {
            return (this.ReportAllProperties());
        }

        public int Num { get; set; }
        public string McRInspectID { get; set; }
        public string McRSeq { get; set; }
        public string McRInspectLegend { get; set; }
        public string McRInspectValue { get; set; }
        public string McInspectBasisID { get; set; }
        public string McSeq { get; set; }
        public string McInsCheck { get; set; }
        public string McInsCycle { get; set; }
        public string McInsRecord { get; set; }
        public string McInsRecordGbn { get; set; }
        public string McInsItemName { get; set; }
        public string McInsContent { get; set; }
        public string McInsCycleGbn { get; set; }
        public string Legend { get; set; }
        public string McImagePath { get; set; }
        public string McImageFile { get; set; }

        public string LegendShape { get; set; }
        public bool flagBool { get; set; }
        public BitmapImage ImageView { get; set; }
        public bool imageFlag { get; set; }
    }
}

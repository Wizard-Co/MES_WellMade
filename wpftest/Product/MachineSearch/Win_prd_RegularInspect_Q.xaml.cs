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
using WizMes_WellMade.PopUP;
using WizMes_WellMade.PopUp;
using WPF.MDI;

namespace WizMes_WellMade
{
    /// <summary>
    /// Win_prd_RegularInspect_Q.xaml에 대한 상호 작용 논리
    /// </summary>
    public partial class Win_prd_RegularInspect_Q : UserControl
    {
        string stDate = string.Empty;
        string stTime = string.Empty;

        string strFlag = string.Empty;
        int rowNum = 0;
        string strPoint = string.Empty;
        Lib lib = new Lib();
        Win_prd_RegularInspect_Q_CodeView MCRegulnsQ = new Win_prd_RegularInspect_Q_CodeView();
        Win_prd_RegularInspect_Q_Sub1_CodeView MCRegulnsSub1Q = new Win_prd_RegularInspect_Q_Sub1_CodeView();
        Win_prd_RegularInspect_Q_Sub2_CodeView MCRegulnsSub2Q = new Win_prd_RegularInspect_Q_Sub2_CodeView();

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

        //string FTP_ADDRESS = "ftp://wizis.iptime.org/ImageData/McRIB";
        //string FTP_ADDRESS = "ftp://wizis.iptime.org/ImageData/Draw";
        //string FTP_ADDRESS = "ftp://222.104.222.145:25000/ImageData//*M*/cReqularInspect";
        //string FTP_ADDRESS = "ftp://192.168.0.95/ImageData/McReqularInspect";
        string FTP_ADDRESS = "ftp://" + LoadINI.FileSvr + ":"
            + LoadINI.FTPPort + LoadINI.FtpImagePath + "/McRIB";
        private const string FTP_ID = "wizuser";
        private const string FTP_PASS = "wiz9999";
        private const string LOCAL_DOWN_PATH = "C:\\Temp";

        private List<string> lstMsg = new List<string>();
        private string message = "";

        public Win_prd_RegularInspect_Q()
        {
            InitializeComponent();
        }

        private void UserControl_Loaded(object sender, RoutedEventArgs e)
        {
            stDate = DateTime.Now.ToString("yyyyMMdd");
            stTime = DateTime.Now.ToString("HHmm");

            DataStore.Instance.InsertLogByFormS(this.GetType().Name, stDate, stTime, "S");

            Lib.Instance.UiLoading(sender);
            //chkInspectDate.IsChecked = true;
            dtpSDate.SelectedDate = DateTime.Today;
            dtpEDate.SelectedDate = DateTime.Today;
            SetComboBox();
        }

        #region 콤보박스
        private void SetComboBox()
        {
            ObservableCollection<CodeView> ovcMcInsCycleGbnSrh = ComboBoxUtil.Instance.Gf_DB_CM_GetComCodeDataset(null, "MCCYCLEGBN", "Y", "", "");
            this.cboRegularInspect.ItemsSource = ovcMcInsCycleGbnSrh;
            this.cboRegularInspect.DisplayMemberPath = "code_name";
            this.cboRegularInspect.SelectedValuePath = "code_id";
        }
        #endregion

        //image 만 Bit로 세팅( imageSource랑 바인딩 )
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

        //검사일자
        private void lblInspectDate_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            //if (chkInspectDate.IsChecked == true) { chkInspectDate.IsChecked = false; }
            //else { chkInspectDate.IsChecked = true; }
        }

        //검사일자
        private void chkInspectDate_Checked(object sender, RoutedEventArgs e)
        {
            dtpSDate.IsEnabled = true;
            dtpEDate.IsEnabled = true;
        }

        //검사일자
        private void chkInspectDate_Unchecked(object sender, RoutedEventArgs e)
        {
            dtpSDate.IsEnabled = false;
            dtpEDate.IsEnabled = false;
        }

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


        //정기검사구분 라벨 클릭시
        private void lblRegularInspect_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            if (chkRegularInspect.IsChecked == true) { chkRegularInspect.IsChecked = false; }
            else { chkRegularInspect.IsChecked = true; }
        }

        //정기검사구분 라벨 in 체크박스 체크시
        private void chkRegularInspect_Checked(object sender, RoutedEventArgs e)
        {
            cboRegularInspect.IsEnabled = true;
        }

        //정기검사구분 라벨 in 체크박스 언체크시
        private void chkRegularInspect_Unchecked(object sender, RoutedEventArgs e)
        {
            cboRegularInspect.IsEnabled = false;
            cboRegularInspect.SelectedItem = null;


        }

        //설비명 라벨 클릭시
        private void lblMcPartName_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            if (chkMcPartName.IsChecked == true) { chkMcPartName.IsChecked = false; }
            else { chkMcPartName.IsChecked = true; }
        }

        //설비명 라벨 in 체크박스 체크시
        private void chkMcPartName_Checked(object sender, RoutedEventArgs e)
        {
            txtMcPartName.IsEnabled = true;
            btnPfMcPartName.IsEnabled = true;
        }

        //설비명 라벨 in 체크박스 언체크시
        private void chkMcPartName_Unchecked(object sender, RoutedEventArgs e)
        {
            txtMcPartName.IsEnabled = false;
            btnPfMcPartName.IsEnabled = false;
        }

        //설비명 텍스트 엔터키 이벤트용(상단)
        private void txtMcPartName_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                MainWindow.pf.ReturnCode(txtMcPartName, (int)Defind_CodeFind.DCF_MC, "");
            }
        }

        //설비명 버튼 클릭 이벤트용(상단)
        private void btnPfMcPartName_Click(object sender, RoutedEventArgs e)
        {
            MainWindow.pf.ReturnCode(txtMcPartName, (int)Defind_CodeFind.DCF_MC, "");
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
                        if (dgdMCPartRInsOne.Items.Count <= 0)
                        {
                            MessageBox.Show("조회된 내용이 없습니다.");
                        }
                        else
                        {
                            dgdMCPartRInsOne.SelectedIndex = rowNum;
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

        //년간 점검 Sheet 
        private void btnYearSheet_Click(object sender, RoutedEventArgs e)
        {
            //일단 비워둠
        }

        //일상 점건 Sheet
        private void btnDaySheet_Click(object sender, RoutedEventArgs e)
        {
            //일단 비워둠
        }

        //엑셀
        private void btnExcel_Click(object sender, RoutedEventArgs e)
        {
            DataTable dt = null;
            string Name = string.Empty;

            string[] lst = new string[6];
            lst[0] = "설비점검 조회결과";
            lst[1] = "설비점검 일자별 조회 결과";
            lst[2] = "설비점검 항목별 조회 결과";
            lst[3] = dgdMCPartRInsOne.Name;
            lst[4] = dgdMCPartRInsTwo.Name;
            lst[5] = dgdMCPartRInsThree.Name;

            ExportExcelxaml ExpExc = new ExportExcelxaml(lst);
            ExpExc.ShowDialog();

            if (ExpExc.DialogResult.HasValue)
            {
                DataStore.Instance.InsertLogByForm(this.GetType().Name, "E");
                if (ExpExc.choice.Equals(dgdMCPartRInsOne.Name))
                {
                    if (ExpExc.Check.Equals("Y"))
                        dt = Lib.Instance.DataGridToDTinHidden(dgdMCPartRInsOne);
                    else
                        dt = Lib.Instance.DataGirdToDataTable(dgdMCPartRInsOne);

                    Name = dgdMCPartRInsOne.Name;
                    Lib.Instance.GenerateExcel(dt, Name);
                    Lib.Instance.excel.Visible = true;
                }
                else if (ExpExc.choice.Equals(dgdMCPartRInsTwo.Name))
                {
                    if (ExpExc.Check.Equals("Y"))
                        dt = Lib.Instance.DataGridToDTinHidden(dgdMCPartRInsTwo);
                    else
                        dt = Lib.Instance.DataGirdToDataTable(dgdMCPartRInsTwo);

                    Name = dgdMCPartRInsTwo.Name;
                    Lib.Instance.GenerateExcel(dt, Name);
                    Lib.Instance.excel.Visible = true;
                }
                else if (ExpExc.choice.Equals(dgdMCPartRInsThree.Name))
                {
                    if (ExpExc.Check.Equals("Y"))
                        dt = Lib.Instance.DataGridToDTinHidden(dgdMCPartRInsThree);
                    else
                        dt = Lib.Instance.DataGirdToDataTable(dgdMCPartRInsThree);

                    Name = dgdMCPartRInsThree.Name;
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

        /// <summary>
        /// 재검색(수정,삭제,추가 저장후에 자동 재검색)
        /// </summary>
        /// <param name="selectedIndex"></param>
        private void re_Search(int selectedIndex)
        {
            FillGrid();

            if (dgdMCPartRInsOne.Items.Count > 0)
            {
                dgdMCPartRInsOne.SelectedIndex = selectedIndex;
            }
            else
            {
                MessageBox.Show("조회된 데이터가 없습니다.");
                return;
            }
        }

        /// <summary>
        /// 실조회
        /// </summary>
        private void FillGrid()
        {
            if (dgdMCPartRInsOne.Items.Count > 0)
            {
                dgdMCPartRInsOne.Items.Clear();
            }

            try
            {
                DataSet ds = null;
                Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                sqlParameter.Clear();
                sqlParameter.Add("ChkMcRInspectDate", 1 ); //chkInspectDate.IsChecked == true ? 1 : 0
                sqlParameter.Add("FromDate", dtpSDate.SelectedDate.Value.ToString("yyyyMMdd")); //chkInspectDate.IsChecked == true ? dtpSDate.SelectedDate.Value.ToString("yyyyMM") : ""
                sqlParameter.Add("ToDate", dtpEDate.SelectedDate.Value.ToString("yyyyMMdd")); //chkInspectDate.IsChecked == true ? dtpSDate.SelectedDate.Value.ToString("yyyyMM") : ""

                sqlParameter.Add("chkMCID", chkMcPartName.IsChecked == true ? 1 : 0);
                sqlParameter.Add("MCID", chkMcPartName.IsChecked == true ? (txtMcPartName.Tag != null ?
                    txtMcPartName.Tag.ToString() : "") : "");
                sqlParameter.Add("ChkInsCycleGbn", chkRegularInspect.IsChecked == true ? 1 : 0);
                sqlParameter.Add("McInsCycleGbn", chkRegularInspect.IsChecked == true ?
                    (cboRegularInspect.SelectedValue != null ? cboRegularInspect.SelectedValue.ToString() : "") : "");
                ds = DataStore.Instance.ProcedureToDataSet_LogWrite("xp_McReqularInspect_sMcReqularInspectMax", sqlParameter, true, "R");

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
                            var WinMCRegulInspect = new Win_prd_RegularInspect_Q_CodeView()
                            {
                                Num = i + 1,
                                McInspectBasisID = dr["McInspectBasisID"].ToString(),
                                MCID = dr["MCID"].ToString(),
                                McInsCycleGbn = dr["McInsCycleGbn"].ToString(),
                                McInsCycle = dr["McInsCycle"].ToString(),
                                McInsBasisDate = dr["McInsBasisDate"].ToString(),
                                McInsContent = dr["McInsContent"].ToString(),
                                Comments = dr["Comments"].ToString(),
                                McRInspectID = dr["McRInspectID"].ToString(),
                                McRInspectDate = dr["McRInspectDate"].ToString(),
                                Name = dr["Name"].ToString(),
                                McRInspectUserID = dr["McRInspectUserID"].ToString(),
                                DefectContents = dr["DefectContents"].ToString(),
                                DefectReason = dr["DefectReason"].ToString(),
                                DefectRespectContents = dr["DefectRespectContents"].ToString(),
                                ManagerID = dr["ManagerID"].ToString(),
                                McName = dr["McName"].ToString(),
                                BuyDate = dr["BuyDate"].ToString(),
                                ImageFile = dr["ImageFile"].ToString(),
                                ImagePath = dr["ImagePath"].ToString()
                            };

                            dgdMCPartRInsOne.Items.Add(WinMCRegulInspect);
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

        private void dgdMCPartRInsOne_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            MCRegulnsQ = dgdMCPartRInsOne.SelectedItem as Win_prd_RegularInspect_Q_CodeView;

            if (MCRegulnsQ != null)
            {
                this.DataContext = MCRegulnsQ;
                FillGridSubOne(MCRegulnsQ.McInspectBasisID, MCRegulnsQ.McInsCycleGbn, MCRegulnsQ.McRInspectDate, MCRegulnsQ.McRInspectUserID);
                FillGridSubTwo(MCRegulnsQ.McInspectBasisID, MCRegulnsQ.McInsCycleGbn, MCRegulnsQ.McRInspectDate);
            }
        }

        //Two의 정보
        private void FillGridSubOne(string strBasisID, string strCycleGbn, string strDate, string strUserID)
        {
            if (dgdMCPartRInsTwo.Items.Count > 0)
            {
                dgdMCPartRInsTwo.Items.Clear();
            }

            try
            {
                DataSet ds = null;
                Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                sqlParameter.Clear();
                sqlParameter.Add("McInspectBasisID", strBasisID);
                sqlParameter.Add("McInsCycleGbn", strCycleGbn);
                sqlParameter.Add("sMonth", Lib.Instance.Left(strDate, 6));
                sqlParameter.Add("McInspectUserID", strUserID);
                ds = DataStore.Instance.ProcedureToDataSet("xp_McRegularInspect_sMcRegularInspectContent", sqlParameter, false);

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
                            var WinMCRegulInspect = new Win_prd_RegularInspect_Q_Sub1_CodeView()
                            {
                                Num = i + 1,
                                step = dr["step"].ToString(),
                                McRInspectID = dr["McRInspectID"].ToString(),
                                McRInspectValue = dr["McRInspectValue"].ToString(),
                                McRInspectDate = dr["McRInspectDate"].ToString(),
                                McInsCycleGbn = dr["McInsCycleGbn"].ToString(),
                                McInsRecordGbn = dr["McInsRecordGbn"].ToString(),
                                McSeq = dr["McSeq"].ToString(),
                                McInsCheck = dr["McInsCheck"].ToString(),
                                McInsCycle = dr["McInsCycle"].ToString(),
                                McInsReCord = dr["McInsReCord"].ToString(),
                                McInsItemName = dr["McInsItemName"].ToString(),
                                McInsContent = dr["McInsContent"].ToString(),
                                McInsCycleDate = dr["McInsCycleDate"].ToString(),
                                DefectContents = dr["DefectContents"].ToString(),
                                DefectReason = dr["DefectReason"].ToString(),
                                DefectRespectContents = dr["DefectRespectContents"].ToString(),
                                Comments = dr["Comments"].ToString()
                            };

                            dgdMCPartRInsTwo.Items.Add(WinMCRegulInspect);
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

        //Three의 정보
        private void FillGridSubTwo(string strBasisID, string strCycleGbn, string strDate)
        {
            message = "";
            lstMsg.Clear();

            if (dgdMCPartRInsThree.Items.Count > 0)
            {
                dgdMCPartRInsThree.Items.Clear();
            }

            try
            {
                DataSet ds = null;
                Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                sqlParameter.Clear();
                sqlParameter.Add("McInspectBasisID", strBasisID);
                sqlParameter.Add("McInsCycleGbn", strCycleGbn);
                sqlParameter.Add("sMonth", Lib.Instance.Left(strDate, 6));
                ds = DataStore.Instance.ProcedureToDataSet("xp_McRegularInspect_sMcRegularInspectDaily", sqlParameter, false);

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

                        if (strCycleGbn.Equals("1"))
                        {
                            foreach (DataRow dr in drc)
                            {
                                var WinMCRegulInspect = new Win_prd_RegularInspect_Q_Sub2_CodeView()
                                {
                                    Num = i + 1,
                                    DD1 = dr["DD1"].ToString(),
                                    DD2 = dr["DD2"].ToString(),
                                    DD3 = dr["DD3"].ToString(),
                                    DD4 = dr["DD4"].ToString(),
                                    DD5 = dr["DD5"].ToString(),
                                    DD6 = dr["DD6"].ToString(),
                                    DD7 = dr["DD7"].ToString(),
                                    DD8 = dr["DD8"].ToString(),
                                    DD9 = dr["DD9"].ToString(),
                                    DD10 = dr["DD10"].ToString(),
                                    DD11 = dr["DD11"].ToString(),
                                    DD12 = dr["DD12"].ToString(),
                                    DD13 = dr["DD13"].ToString(),
                                    DD14 = dr["DD14"].ToString(),
                                    DD15 = dr["DD15"].ToString(),
                                    DD16 = dr["DD16"].ToString(),
                                    DD17 = dr["DD17"].ToString(),
                                    DD18 = dr["DD18"].ToString(),
                                    DD19 = dr["DD19"].ToString(),
                                    DD20 = dr["DD20"].ToString(),
                                    DD21 = dr["DD21"].ToString(),
                                    DD22 = dr["DD22"].ToString(),
                                    DD23 = dr["DD23"].ToString(),
                                    DD24 = dr["DD24"].ToString(),
                                    DD25 = dr["DD25"].ToString(),
                                    DD26 = dr["DD26"].ToString(),
                                    DD27 = dr["DD27"].ToString(),
                                    DD28 = dr["DD28"].ToString(),
                                    DD29 = dr["DD29"].ToString(),
                                    DD30 = dr["DD30"].ToString(),
                                    DD31 = dr["DD31"].ToString(),
                                    McRInspectValue1 = dr["McRInspectValue1"].ToString(),
                                    McRInspectValue2 = dr["McRInspectValue2"].ToString(),
                                    McRInspectValue3 = dr["McRInspectValue3"].ToString(),
                                    McRInspectValue4 = dr["McRInspectValue4"].ToString(),
                                    McRInspectValue5 = dr["McRInspectValue5"].ToString(),
                                    McRInspectValue6 = dr["McRInspectValue6"].ToString(),
                                    McRInspectValue7 = dr["McRInspectValue7"].ToString(),
                                    McRInspectValue8 = dr["McRInspectValue8"].ToString(),
                                    McRInspectValue9 = dr["McRInspectValue9"].ToString(),
                                    McRInspectValue10 = dr["McRInspectValue10"].ToString(),
                                    McRInspectValue11 = dr["McRInspectValue11"].ToString(),
                                    McRInspectValue12 = dr["McRInspectValue12"].ToString(),
                                    McRInspectValue13 = dr["McRInspectValue13"].ToString(),
                                    McRInspectValue14 = dr["McRInspectValue14"].ToString(),
                                    McRInspectValue15 = dr["McRInspectValue15"].ToString(),
                                    McRInspectValue16 = dr["McRInspectValue16"].ToString(),
                                    McRInspectValue17 = dr["McRInspectValue17"].ToString(),
                                    McRInspectValue18 = dr["McRInspectValue18"].ToString(),
                                    McRInspectValue19 = dr["McRInspectValue19"].ToString(),
                                    McRInspectValue20 = dr["McRInspectValue20"].ToString(),
                                    McRInspectValue21 = dr["McRInspectValue21"].ToString(),
                                    McRInspectValue22 = dr["McRInspectValue22"].ToString(),
                                    McRInspectValue23 = dr["McRInspectValue23"].ToString(),
                                    McRInspectValue24 = dr["McRInspectValue24"].ToString(),
                                    McRInspectValue25 = dr["McRInspectValue25"].ToString(),
                                    McRInspectValue26 = dr["McRInspectValue26"].ToString(),
                                    McRInspectValue27 = dr["McRInspectValue27"].ToString(),
                                    McRInspectValue28 = dr["McRInspectValue28"].ToString(),
                                    McRInspectValue29 = dr["McRInspectValue29"].ToString(),
                                    McRInspectValue30 = dr["McRInspectValue30"].ToString(),
                                    McRInspectValue31 = dr["McRInspectValue31"].ToString(),
                                    McInsCheck = dr["McInsCheck"].ToString(),
                                    McInsCycle = dr["McInsCycle"].ToString(),
                                    McInsReCord = dr["McInsReCord"].ToString(),
                                    McInsContent = dr["McInsContent"].ToString(),
                                    McInsItemName = dr["McInsItemName"].ToString(),
                                    McInsRecordGbn = dr["McInsRecordGbn"].ToString(),
                                    McInsCycleGbn = dr["McInsCycleGbn"].ToString(),
                                    McInsCycleDate = dr["McInsCycleDate"].ToString(),
                                    EndDay = dr["EndDay"].ToString(),
                                    McSeq = dr["McSeq"].ToString(),
                                    McImageFile = dr["McImageFile"].ToString(),
                                    McImagePath = dr["McImagePath"].ToString()
                                };

                                if (WinMCRegulInspect.McImageFile != null && !WinMCRegulInspect.McImageFile.Replace(" ", "").Equals(""))
                                {
                                    WinMCRegulInspect.imageFlag = true;

                                    if (CheckImage(WinMCRegulInspect.McImageFile.Trim()))
                                    {
                                        string strImage = "/" + MCRegulnsQ.McInspectBasisID + "/" + WinMCRegulInspect.McImageFile;
                                        WinMCRegulInspect.ImageView = SetImage(strImage, WinMCRegulInspect.McImageFile);
                                    }
                                    else
                                    {
                                        MessageBox.Show(WinMCRegulInspect.McImageFile + "는 이미지 변환이 불가능합니다.");
                                    }
                                }
                                else
                                {
                                    WinMCRegulInspect.imageFlag = false;
                                }

                                WinMCRegulInspect.McRInspectValue1 = Lib.Instance.returnNumStringOne(WinMCRegulInspect.McRInspectValue1);
                                WinMCRegulInspect.McRInspectValue2 = Lib.Instance.returnNumStringOne(WinMCRegulInspect.McRInspectValue2);
                                WinMCRegulInspect.McRInspectValue3 = Lib.Instance.returnNumStringOne(WinMCRegulInspect.McRInspectValue3);
                                WinMCRegulInspect.McRInspectValue4 = Lib.Instance.returnNumStringOne(WinMCRegulInspect.McRInspectValue4);
                                WinMCRegulInspect.McRInspectValue5 = Lib.Instance.returnNumStringOne(WinMCRegulInspect.McRInspectValue5);
                                WinMCRegulInspect.McRInspectValue6 = Lib.Instance.returnNumStringOne(WinMCRegulInspect.McRInspectValue6);
                                WinMCRegulInspect.McRInspectValue7 = Lib.Instance.returnNumStringOne(WinMCRegulInspect.McRInspectValue7);
                                WinMCRegulInspect.McRInspectValue8 = Lib.Instance.returnNumStringOne(WinMCRegulInspect.McRInspectValue8);
                                WinMCRegulInspect.McRInspectValue9 = Lib.Instance.returnNumStringOne(WinMCRegulInspect.McRInspectValue9);
                                WinMCRegulInspect.McRInspectValue10 = Lib.Instance.returnNumStringOne(WinMCRegulInspect.McRInspectValue10);
                                WinMCRegulInspect.McRInspectValue11 = Lib.Instance.returnNumStringOne(WinMCRegulInspect.McRInspectValue11);
                                WinMCRegulInspect.McRInspectValue12 = Lib.Instance.returnNumStringOne(WinMCRegulInspect.McRInspectValue12);
                                WinMCRegulInspect.McRInspectValue13 = Lib.Instance.returnNumStringOne(WinMCRegulInspect.McRInspectValue13);
                                WinMCRegulInspect.McRInspectValue14 = Lib.Instance.returnNumStringOne(WinMCRegulInspect.McRInspectValue14);
                                WinMCRegulInspect.McRInspectValue15 = Lib.Instance.returnNumStringOne(WinMCRegulInspect.McRInspectValue15);
                                WinMCRegulInspect.McRInspectValue16 = Lib.Instance.returnNumStringOne(WinMCRegulInspect.McRInspectValue16);
                                WinMCRegulInspect.McRInspectValue17 = Lib.Instance.returnNumStringOne(WinMCRegulInspect.McRInspectValue17);
                                WinMCRegulInspect.McRInspectValue18 = Lib.Instance.returnNumStringOne(WinMCRegulInspect.McRInspectValue18);
                                WinMCRegulInspect.McRInspectValue19 = Lib.Instance.returnNumStringOne(WinMCRegulInspect.McRInspectValue19);
                                WinMCRegulInspect.McRInspectValue20 = Lib.Instance.returnNumStringOne(WinMCRegulInspect.McRInspectValue20);
                                WinMCRegulInspect.McRInspectValue21 = Lib.Instance.returnNumStringOne(WinMCRegulInspect.McRInspectValue21);
                                WinMCRegulInspect.McRInspectValue22 = Lib.Instance.returnNumStringOne(WinMCRegulInspect.McRInspectValue22);
                                WinMCRegulInspect.McRInspectValue23 = Lib.Instance.returnNumStringOne(WinMCRegulInspect.McRInspectValue23);
                                WinMCRegulInspect.McRInspectValue24 = Lib.Instance.returnNumStringOne(WinMCRegulInspect.McRInspectValue24);
                                WinMCRegulInspect.McRInspectValue25 = Lib.Instance.returnNumStringOne(WinMCRegulInspect.McRInspectValue25);
                                WinMCRegulInspect.McRInspectValue26 = Lib.Instance.returnNumStringOne(WinMCRegulInspect.McRInspectValue26);
                                WinMCRegulInspect.McRInspectValue27 = Lib.Instance.returnNumStringOne(WinMCRegulInspect.McRInspectValue27);
                                WinMCRegulInspect.McRInspectValue28 = Lib.Instance.returnNumStringOne(WinMCRegulInspect.McRInspectValue28);
                                WinMCRegulInspect.McRInspectValue29 = Lib.Instance.returnNumStringOne(WinMCRegulInspect.McRInspectValue29);
                                WinMCRegulInspect.McRInspectValue30 = Lib.Instance.returnNumStringOne(WinMCRegulInspect.McRInspectValue30);
                                WinMCRegulInspect.McRInspectValue31 = Lib.Instance.returnNumStringOne(WinMCRegulInspect.McRInspectValue31);

                                dgdMCPartRInsThree.Items.Add(WinMCRegulInspect);
                                i++;
                            }

                            dgdtxtDD13.Visibility = Visibility.Visible;
                            dgdtxtValue13.Visibility = Visibility.Visible;
                            dgdtxtDD14.Visibility = Visibility.Visible;
                            dgdtxtValue14.Visibility = Visibility.Visible;
                            dgdtxtDD15.Visibility = Visibility.Visible;
                            dgdtxtValue15.Visibility = Visibility.Visible;
                            dgdtxtDD16.Visibility = Visibility.Visible;
                            dgdtxtValue16.Visibility = Visibility.Visible;
                            dgdtxtDD17.Visibility = Visibility.Visible;
                            dgdtxtValue17.Visibility = Visibility.Visible;
                            dgdtxtDD18.Visibility = Visibility.Visible;
                            dgdtxtValue18.Visibility = Visibility.Visible;
                            dgdtxtDD19.Visibility = Visibility.Visible;
                            dgdtxtValue19.Visibility = Visibility.Visible;
                            dgdtxtDD20.Visibility = Visibility.Visible;
                            dgdtxtValue20.Visibility = Visibility.Visible;
                            dgdtxtDD21.Visibility = Visibility.Visible;
                            dgdtxtValue21.Visibility = Visibility.Visible;
                            dgdtxtDD22.Visibility = Visibility.Visible;
                            dgdtxtValue22.Visibility = Visibility.Visible;
                            dgdtxtDD23.Visibility = Visibility.Visible;
                            dgdtxtValue23.Visibility = Visibility.Visible;
                            dgdtxtDD24.Visibility = Visibility.Visible;
                            dgdtxtValue24.Visibility = Visibility.Visible;
                            dgdtxtDD25.Visibility = Visibility.Visible;
                            dgdtxtValue25.Visibility = Visibility.Visible;
                            dgdtxtDD26.Visibility = Visibility.Visible;
                            dgdtxtValue26.Visibility = Visibility.Visible;
                            dgdtxtDD27.Visibility = Visibility.Visible;
                            dgdtxtValue27.Visibility = Visibility.Visible;
                            dgdtxtDD28.Visibility = Visibility.Visible;
                            dgdtxtValue28.Visibility = Visibility.Visible;
                            dgdtxtDD29.Visibility = Visibility.Visible;
                            dgdtxtValue29.Visibility = Visibility.Visible;
                            dgdtxtDD30.Visibility = Visibility.Visible;
                            dgdtxtValue30.Visibility = Visibility.Visible;
                            dgdtxtDD31.Visibility = Visibility.Visible;
                            dgdtxtValue31.Visibility = Visibility.Visible;
                        }
                        else
                        {
                            foreach (DataRow dr in drc)
                            {
                                var WinMCRegulInspect = new Win_prd_RegularInspect_Q_Sub2_CodeView()
                                {
                                    Num = i + 1,
                                    DD1 = dr["DD1"].ToString(),
                                    DD2 = dr["DD2"].ToString(),
                                    DD3 = dr["DD3"].ToString(),
                                    DD4 = dr["DD4"].ToString(),
                                    DD5 = dr["DD5"].ToString(),
                                    DD6 = dr["DD6"].ToString(),
                                    DD7 = dr["DD7"].ToString(),
                                    DD8 = dr["DD8"].ToString(),
                                    DD9 = dr["DD9"].ToString(),
                                    DD10 = dr["DD10"].ToString(),
                                    DD11 = dr["DD11"].ToString(),
                                    DD12 = dr["DD12"].ToString(),
                                    McRInspectValue1 = dr["McRInspectValue1"].ToString(),
                                    McRInspectValue2 = dr["McRInspectValue2"].ToString(),
                                    McRInspectValue3 = dr["McRInspectValue3"].ToString(),
                                    McRInspectValue4 = dr["McRInspectValue4"].ToString(),
                                    McRInspectValue5 = dr["McRInspectValue5"].ToString(),
                                    McRInspectValue6 = dr["McRInspectValue6"].ToString(),
                                    McRInspectValue7 = dr["McRInspectValue7"].ToString(),
                                    McRInspectValue8 = dr["McRInspectValue8"].ToString(),
                                    McRInspectValue9 = dr["McRInspectValue9"].ToString(),
                                    McRInspectValue10 = dr["McRInspectValue10"].ToString(),
                                    McRInspectValue11 = dr["McRInspectValue11"].ToString(),
                                    McRInspectValue12 = dr["McRInspectValue12"].ToString(),
                                    McInsCheck = dr["McInsCheck"].ToString(),
                                    McInsCycle = dr["McInsCycle"].ToString(),
                                    McInsReCord = dr["McInsReCord"].ToString(),
                                    McInsContent = dr["McInsContent"].ToString(),
                                    McInsItemName = dr["McInsItemName"].ToString(),
                                    McInsRecordGbn = dr["McInsRecordGbn"].ToString(),
                                    McInsCycleGbn = dr["McInsCycleGbn"].ToString(),
                                    McInsCycleDate = dr["McInsCycleDate"].ToString(),
                                    EndDay = dr["EndDay"].ToString(),
                                    McSeq = dr["McSeq"].ToString(),
                                    McImageFile = dr["McImageFile"].ToString(),
                                    McImagePath = dr["McImagePath"].ToString()
                                };

                                if (WinMCRegulInspect.McImageFile != null && !WinMCRegulInspect.McImageFile.Replace(" ", "").Equals(""))
                                {
                                    WinMCRegulInspect.imageFlag = true;
                                  
                                    if (CheckImage(WinMCRegulInspect.McImageFile.Trim()))
                                    {
                                        string strImage = "/" + MCRegulnsQ.McInspectBasisID + "/" + WinMCRegulInspect.McImageFile;
                                        WinMCRegulInspect.ImageView = SetImage(strImage, WinMCRegulInspect.McImageFile);
                                    }
                                    else
                                    {
                                        MessageBox.Show(WinMCRegulInspect.McImageFile + "는 이미지 변환이 불가능합니다.");
                                    }
                                }
                                else
                                {
                                    WinMCRegulInspect.imageFlag = false;
                                }

                                WinMCRegulInspect.McRInspectValue1 = Lib.Instance.returnNumStringOne(WinMCRegulInspect.McRInspectValue1);
                                WinMCRegulInspect.McRInspectValue2 = Lib.Instance.returnNumStringOne(WinMCRegulInspect.McRInspectValue2);
                                WinMCRegulInspect.McRInspectValue3 = Lib.Instance.returnNumStringOne(WinMCRegulInspect.McRInspectValue3);
                                WinMCRegulInspect.McRInspectValue4 = Lib.Instance.returnNumStringOne(WinMCRegulInspect.McRInspectValue4);
                                WinMCRegulInspect.McRInspectValue5 = Lib.Instance.returnNumStringOne(WinMCRegulInspect.McRInspectValue5);
                                WinMCRegulInspect.McRInspectValue6 = Lib.Instance.returnNumStringOne(WinMCRegulInspect.McRInspectValue6);
                                WinMCRegulInspect.McRInspectValue7 = Lib.Instance.returnNumStringOne(WinMCRegulInspect.McRInspectValue7);
                                WinMCRegulInspect.McRInspectValue8 = Lib.Instance.returnNumStringOne(WinMCRegulInspect.McRInspectValue8);
                                WinMCRegulInspect.McRInspectValue9 = Lib.Instance.returnNumStringOne(WinMCRegulInspect.McRInspectValue9);
                                WinMCRegulInspect.McRInspectValue10 = Lib.Instance.returnNumStringOne(WinMCRegulInspect.McRInspectValue10);
                                WinMCRegulInspect.McRInspectValue11 = Lib.Instance.returnNumStringOne(WinMCRegulInspect.McRInspectValue11);
                                WinMCRegulInspect.McRInspectValue12 = Lib.Instance.returnNumStringOne(WinMCRegulInspect.McRInspectValue12);

                                dgdMCPartRInsThree.Items.Add(WinMCRegulInspect);
                                i++;
                            }

                            dgdtxtDD13.Visibility = Visibility.Hidden;
                            dgdtxtValue13.Visibility = Visibility.Hidden;
                            dgdtxtDD14.Visibility = Visibility.Hidden;
                            dgdtxtValue14.Visibility = Visibility.Hidden;
                            dgdtxtDD15.Visibility = Visibility.Hidden;
                            dgdtxtValue15.Visibility = Visibility.Hidden;
                            dgdtxtDD16.Visibility = Visibility.Hidden;
                            dgdtxtValue16.Visibility = Visibility.Hidden;
                            dgdtxtDD17.Visibility = Visibility.Hidden;
                            dgdtxtValue17.Visibility = Visibility.Hidden;
                            dgdtxtDD18.Visibility = Visibility.Hidden;
                            dgdtxtValue18.Visibility = Visibility.Hidden;
                            dgdtxtDD19.Visibility = Visibility.Hidden;
                            dgdtxtValue19.Visibility = Visibility.Hidden;
                            dgdtxtDD20.Visibility = Visibility.Hidden;
                            dgdtxtValue20.Visibility = Visibility.Hidden;
                            dgdtxtDD21.Visibility = Visibility.Hidden;
                            dgdtxtValue21.Visibility = Visibility.Hidden;
                            dgdtxtDD22.Visibility = Visibility.Hidden;
                            dgdtxtValue22.Visibility = Visibility.Hidden;
                            dgdtxtDD23.Visibility = Visibility.Hidden;
                            dgdtxtValue23.Visibility = Visibility.Hidden;
                            dgdtxtDD24.Visibility = Visibility.Hidden;
                            dgdtxtValue24.Visibility = Visibility.Hidden;
                            dgdtxtDD25.Visibility = Visibility.Hidden;
                            dgdtxtValue25.Visibility = Visibility.Hidden;
                            dgdtxtDD26.Visibility = Visibility.Hidden;
                            dgdtxtValue26.Visibility = Visibility.Hidden;
                            dgdtxtDD27.Visibility = Visibility.Hidden;
                            dgdtxtValue27.Visibility = Visibility.Hidden;
                            dgdtxtDD28.Visibility = Visibility.Hidden;
                            dgdtxtValue28.Visibility = Visibility.Hidden;
                            dgdtxtDD29.Visibility = Visibility.Hidden;
                            dgdtxtValue29.Visibility = Visibility.Hidden;
                            dgdtxtDD30.Visibility = Visibility.Hidden;
                            dgdtxtValue30.Visibility = Visibility.Hidden;
                            dgdtxtDD31.Visibility = Visibility.Hidden;
                            dgdtxtValue31.Visibility = Visibility.Hidden;
                        }
                    }

                    if (!message.Trim().Equals(""))
                    {
                        MessageBox.Show(message + " 를 불러올 수 없습니다.");
                    }

                    dgdMCPartRInsThree.Refresh();
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

        private void btnSeeImage_Click(object sender, RoutedEventArgs e)
        {
            DataGridCellInfo dgdInfo = dgdMCPartRInsThree.CurrentCell;

            if (dgdInfo.Column != null)
            {
                MCRegulnsSub2Q = dgdMCPartRInsThree.CurrentItem as Win_prd_RegularInspect_Q_Sub2_CodeView;
            }

            if (MCRegulnsSub2Q != null && !MCRegulnsSub2Q.McImageFile.Equals(""))
            {
                FTP_DownLoadFile(MCRegulnsSub2Q.McImagePath + "/"
                    + MCRegulnsQ.McInspectBasisID + "/" + MCRegulnsSub2Q.McImageFile);
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
    }

    class Win_prd_RegularInspect_Q_CodeView : BaseView
    {
        public override string ToString()
        {
            return (this.ReportAllProperties());
        }

        public int Num { get; set; }
        public string McInspectBasisID { get; set; }
        public string MCID { get; set; }
        public string McInsCycleGbn { get; set; }
        public string McInsCycle { get; set; }
        public string McInsBasisDate { get; set; }
        public string McInsContent { get; set; }
        public string Comments { get; set; }
        public string McRInspectID { get; set; }
        public string McRInspectDate { get; set; }
        public string Name { get; set; }
        public string McRInspectUserID { get; set; }
        public string DefectContents { get; set; }
        public string DefectReason { get; set; }
        public string DefectRespectContents { get; set; }
        public string ManagerID { get; set; }
        public string McName { get; set; }
        public string BuyDate { get; set; }
        public string ImageFile { get; set; }
        public string ImagePath { get; set; }

        
    }

    class Win_prd_RegularInspect_Q_Sub1_CodeView : BaseView
    {
        public override string ToString()
        {
            return (this.ReportAllProperties());
        }

        public int Num { get; set; }
        public string step { get; set; }
        public string McRInspectID { get; set; }
        public string McRInspectValue { get; set; }
        public string McRInspectDate { get; set; }
        public string McInsCycleGbn { get; set; }
        public string McInsRecordGbn { get; set; }
        public string McSeq { get; set; }
        public string McInsCheck { get; set; }
        public string McInsCycle { get; set; }
        public string McInsReCord { get; set; }
        public string McInsItemName { get; set; }
        public string McInsContent { get; set; }
        public string McInsCycleDate { get; set; }
        public string DefectContents { get; set; }
        public string DefectReason { get; set; }
        public string DefectRespectContents { get; set; }
        public string Comments { get; set; }

        public string McImageFile { get; set; }
        public string McImagePath { get; set; }

        public BitmapImage ImageView { get; set; }
        public bool imageFlag { get; set; }
    }

    class Win_prd_RegularInspect_Q_Sub2_CodeView : BaseView
    {
        public override string ToString()
        {
            return (this.ReportAllProperties());
        }

        public string step { get; set; }
        public int Num { get; set; }
        public string DD1 { get; set; }
        public string DD2 { get; set; }
        public string DD3 { get; set; }
        public string DD4 { get; set; }
        public string DD5 { get; set; }
        public string DD6 { get; set; }
        public string DD7 { get; set; }
        public string DD8 { get; set; }
        public string DD9 { get; set; }
        public string DD10 { get; set; }
        public string DD11 { get; set; }
        public string DD12 { get; set; }
        public string DD13 { get; set; }
        public string DD14 { get; set; }
        public string DD15 { get; set; }
        public string DD16 { get; set; }
        public string DD17 { get; set; }
        public string DD18 { get; set; }
        public string DD19 { get; set; }
        public string DD20 { get; set; }
        public string DD21 { get; set; }
        public string DD22 { get; set; }
        public string DD23 { get; set; }
        public string DD24 { get; set; }
        public string DD25 { get; set; }
        public string DD26 { get; set; }
        public string DD27 { get; set; }
        public string DD28 { get; set; }
        public string DD29 { get; set; }
        public string DD30 { get; set; }
        public string DD31 { get; set; }
        public string McRInspectValue1 { get; set; }
        public string McRInspectValue2 { get; set; }
        public string McRInspectValue3 { get; set; }
        public string McRInspectValue4 { get; set; }
        public string McRInspectValue5 { get; set; }
        public string McRInspectValue6 { get; set; }
        public string McRInspectValue7 { get; set; }
        public string McRInspectValue8 { get; set; }
        public string McRInspectValue9 { get; set; }
        public string McRInspectValue10 { get; set; }
        public string McRInspectValue11 { get; set; }
        public string McRInspectValue12 { get; set; }
        public string McRInspectValue13 { get; set; }
        public string McRInspectValue14 { get; set; }
        public string McRInspectValue15 { get; set; }
        public string McRInspectValue16 { get; set; }
        public string McRInspectValue17 { get; set; }
        public string McRInspectValue18 { get; set; }
        public string McRInspectValue19 { get; set; }
        public string McRInspectValue20 { get; set; }
        public string McRInspectValue21 { get; set; }
        public string McRInspectValue22 { get; set; }
        public string McRInspectValue23 { get; set; }
        public string McRInspectValue24 { get; set; }
        public string McRInspectValue25 { get; set; }
        public string McRInspectValue26 { get; set; }
        public string McRInspectValue27 { get; set; }
        public string McRInspectValue28 { get; set; }
        public string McRInspectValue29 { get; set; }
        public string McRInspectValue30 { get; set; }
        public string McRInspectValue31 { get; set; }

        public string McInsCheck { get; set; }
        public string McInsCycle { get; set; }
        public string McInsReCord { get; set; }
        public string McInsContent { get; set; }
        public string McInsItemName { get; set; }
        public string McInsRecordGbn { get; set; }
        public string McInsCycleGbn { get; set; }
        public string McInsCycleDate { get; set; }
        public string EndDay { get; set; }
        public string McSeq { get; set; }
        public string McImageFile { get; set; }
        public string McImagePath { get; set; }

        public BitmapImage ImageView { get; set; }
        public bool imageFlag { get; set; }
    }
}

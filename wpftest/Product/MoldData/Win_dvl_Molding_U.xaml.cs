using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Drawing;
using System.Linq;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using WizMes_WellMade.PopUP;
using WizMes_WellMade;
using WPF.MDI;
using System.Net;
using System.Windows.Forms.VisualStyles;
using static System.Windows.Forms.AxHost;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.TaskbarClock;
using System.Diagnostics.Eventing.Reader;

namespace WizMes_WellMade
{

    /// <summary>
    /// Win_dvl_Molding_U.xaml에 대한 상호 작용 논리
    /// </summary>
    public partial class Win_dvl_Molding_U : UserControl
    {
        string strFlag = string.Empty;
        int rowNum = 0;
        bool MultiArticle = false;

        Lib lib = new Lib();

        Win_dvl_Molding_U_CodeView WinMold = new Win_dvl_Molding_U_CodeView();
        Win_dvl_Molding_U_Parts_CodeView WinMoldParts = new Win_dvl_Molding_U_Parts_CodeView();
        MoldArticle_CodeView MoldArticleList = new MoldArticle_CodeView();

        // FTP 활용모음.
        string FullPath1 = string.Empty;
        string FullPath2 = string.Empty;
        string FullPath3 = string.Empty;

        private FTP_EX _ftp = null;
        private List<UploadFileInfo> _listFileInfo = new List<UploadFileInfo>();

        string stDate = string.Empty;
        string stTime = string.Empty;

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


#if DEBUG
        string FTP_ADDRESS = "ftp://wizis.iptime.org/ImageData/Mold";
#else
        string FTP_ADDRESS = "ftp://" + LoadINI.FileSvr + ":"
            + LoadINI.FTPPort + LoadINI.FtpImagePath + "/Mold";
#endif
        private const string FTP_ID = "wizuser";
        private const string FTP_PASS = "wiz9999";
        private const string LOCAL_DOWN_PATH = "C:\\Temp";

        public Win_dvl_Molding_U()
        {
            InitializeComponent();
        }


        private void Usercontrol_Loaded(object sender, RoutedEventArgs e)
        {
            stDate = DateTime.Now.ToString("yyyyMMdd");
            stTime = DateTime.Now.ToString("HHmm");
            chkDate.IsChecked = true;

            Lib.Instance.UiLoading(this);
            SetComboBox();
            btnToday_Click(null, null);
        }

        private void SetComboBox()
        {

            List<string[]> lstDvlYN = new List<string[]>();
            string[] strDvl_1 = { "Y", "Y" };
            string[] strDvl_2 = { "N", "N" };
            lstDvlYN.Add(strDvl_1);
            lstDvlYN.Add(strDvl_2);

            ObservableCollection<CodeView> ovcDvlYN = ComboBoxUtil.Instance.Direct_SetComboBox(lstDvlYN);
            this.cboMainUseYN.ItemsSource = ovcDvlYN;
            this.cboMainUseYN.DisplayMemberPath = "code_name";
            this.cboMainUseYN.SelectedValuePath = "code_id";

            List<string[]> lstDsiYN = new List<string[]>();
            string[] strDis_1 = { "N", "사용" };
            string[] strDis_2 = { "Y", "불용" };
            string[] strDis_3 = { "S", "스페어" };
            lstDsiYN.Add(strDis_1);
            lstDsiYN.Add(strDis_2);
            lstDsiYN.Add(strDis_3);

            ObservableCollection<CodeView> ovcForUseSrh = ComboBoxUtil.Instance.Direct_SetComboBox(lstDsiYN);
            this.cboDisCard.ItemsSource = ovcForUseSrh;
            this.cboDisCard.DisplayMemberPath = "code_name";
            this.cboDisCard.SelectedValuePath = "code_id";

            List<string[]> lstColor = new List<string[]>();
            string[] strColor_1 = { "N", "노랑" };
            string[] strColor_2 = { "Y", "빨강" };
            string[] strColor_3 = { "S", "초록" };
            string[] strColor_4 = { "S", "흰색" };
            lstColor.Add(strColor_1);
            lstColor.Add(strColor_2);
            lstColor.Add(strColor_3);
            lstColor.Add(strColor_4);

            ObservableCollection<CodeView> ovcColor = ComboBoxUtil.Instance.Direct_SetComboBox(lstColor);
            this.cboColor.ItemsSource = ovcColor;
            this.cboColor.DisplayMemberPath = "code_name";
            this.cboColor.SelectedValuePath = "code_id";


            ObservableCollection<CodeView> ovMoldPlace = ComboBoxUtil.Instance.Gf_DB_CM_GetComCodeDataset(null, "MOLDPLACE", "Y", "");
            this.cboStorgeLocation.ItemsSource = ovMoldPlace;
            this.cboStorgeLocation.DisplayMemberPath = "code_name";
            this.cboStorgeLocation.SelectedValuePath = "code_id";

            ObservableCollection<CodeView> ovMoldPay = ComboBoxUtil.Instance.Gf_DB_CM_GetComCodeDataset(null, "MoldPay", "Y", "");
            this.cboBoxOwnerOneTimePayYn.ItemsSource = ovMoldPay;
            this.cboBoxOwnerOneTimePayYn.DisplayMemberPath = "code_name";
            this.cboBoxOwnerOneTimePayYn.SelectedValuePath = "code_id";

        }

        #region 라벨 클릭 및 체크박스 이벤트

        //금형발주일
        private void lblDate_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            if (chkDate.IsChecked == true) { chkDate.IsChecked = false; }
            else { chkDate.IsChecked = true; }
        }

        //금형발주일
        private void chkDate_Checked(object sender, RoutedEventArgs e)
        {
            dtpSDate.IsEnabled = true;
            dtpEDate.IsEnabled = true;
        }

        //금형발주일
        private void chkDate_Unchecked(object sender, RoutedEventArgs e)
        {
            dtpSDate.IsEnabled = false;
            dtpEDate.IsEnabled = false;
        }

        //금월
        private void btnThisMonth_Click(object sender, RoutedEventArgs e)
        {
            dtpSDate.SelectedDate = Lib.Instance.BringThisMonthDatetimeList()[0];
            dtpEDate.SelectedDate = Lib.Instance.BringThisMonthDatetimeList()[1];
        }

        //금일
        private void btnToday_Click(object sender, RoutedEventArgs e)
        {
            dtpSDate.SelectedDate = DateTime.Today;
            dtpEDate.SelectedDate = DateTime.Today;
        }
        //금형 점검필요 
        private void lblNeedInspectSrh_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            if (chkNeedInspectSrh.IsChecked == true)
            {
                chkNeedInspectSrh.IsChecked = false;
            }
            else
            {
                chkNeedInspectSrh.IsChecked = true;
            }
        }
        //사용기한 경과
        private void lblExpiredSrh_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            if (chkExpiredSrh.IsChecked == true)
            {             
                chkExpiredSrh.IsChecked = false;
            }else {       
                chkExpiredSrh.IsChecked = true; }
        }
        //폐기건 포함
        private void lblDisCardSrh_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            if (chkDisCardSrh.IsChecked == true) { chkDisCardSrh.IsChecked = false; }
            else { chkDisCardSrh.IsChecked = true; }
        }

        //세척점검필요
        private void lblNeedWashing_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            if (chkNeedWashing.IsChecked == true) { chkNeedWashing.IsChecked = false; }
            else { chkNeedWashing.IsChecked = true; }
        }

        //금형LotNo(%)
        private void lblMoldNoSrh_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            if (chkMoldNoSrh.IsChecked == true) { chkMoldNoSrh.IsChecked = false; }
            else { chkMoldNoSrh.IsChecked = true; }
        }

        //금형LotNo(%)
        private void chkMoldNoSrh_Checked(object sender, RoutedEventArgs e)
        {
            txtMoldNoSrh.IsEnabled = true;
        }

        //금형LotNo(%)
        private void chkMoldNoSrh_Unchecked(object sender, RoutedEventArgs e)
        {
            txtMoldNoSrh.IsEnabled = false;
        }

        //품명
        private void lblArticleSrh_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            if (chkArticleSrh.IsChecked == true) { chkArticleSrh.IsChecked = false; }
            else { chkArticleSrh.IsChecked = true; }
        }

        //품명
        private void chkArticleSrh_Checked(object sender, RoutedEventArgs e)
        {
            txtArticleSrh.IsEnabled = true;
            btnPfArticleSrh.IsEnabled = true;
        }

        //품명
        private void chkArticleSrh_Unchecked(object sender, RoutedEventArgs e)
        {
            txtArticleSrh.IsEnabled = false;
            btnPfArticleSrh.IsEnabled = false;
        }

        //품명
        private void txtArticleSrh_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                MainWindow.pf.ReturnCode(txtArticleSrh, 76, "");
            }
        }

        //품명
        private void btnPfArticleSrh_Click(object sender, RoutedEventArgs e)
        {
            MainWindow.pf.ReturnCode(txtArticleSrh, 76, "");
        }



        #endregion

        /// <summary>
        /// 수정,추가 저장 후
        /// </summary>
        private void CanBtnControl()
        {
            Lib.Instance.UiButtonEnableChange_IUControl(this);
            grdInput1.IsEnabled = false;
            //gbxInput.IsEnabled = false;
            grxInput.IsEnabled = false;
            //dgdMain.IsEnabled = true;
            dgdMain.IsHitTestVisible = true;
        }

        /// <summary>
        /// 수정,추가 진행 중
        /// </summary>
        private void CantBtnControl()
        {
            Lib.Instance.UiButtonEnableChange_SCControl(this);
            grdInput1.IsEnabled = true;
            grxInput.IsEnabled = true;
            dgdMain.IsHitTestVisible = false;
            dgdMoldArticle.IsHitTestVisible = true;
            btnMoldArticleAdd.IsEnabled = true;
            btnMoldArticleDelete.IsEnabled = true;

        }

        //추가
        private void btnAdd_Click(object sender, RoutedEventArgs e)
        {
            CantBtnControl();
            strFlag = "I";

            tbkMsg.Text = "자료 입력 중";

            //유지추가 버튼 false
            if (chkMainTain.IsChecked == false)
            {
                if (dgdPartsCode.Items.Count > 0) dgdPartsCode.Items.Clear();
                if (dgdMoldArticle.Items.Count > 0) dgdMoldArticle.Items.Clear();
                this.DataContext = null;
                cboDisCard.SelectedIndex = 0;
                cboStorgeLocation.SelectedIndex = 0;
                cboMainUseYN.SelectedIndex = 0;
                cboBoxOwnerOneTimePayYn.SelectedIndex = 0;
            }

            rowNum = dgdMain.SelectedIndex;

            chkSetDate.IsChecked = true;
            dtpSetDate.SelectedDate = DateTime.Today;

            chkProdCompDate.IsChecked = true;
            dtpProdCompDate.SelectedDate = DateTime.Today;

            chkProdOrderDate.IsChecked = true;
            dtpProdOrderDate.SelectedDate = DateTime.Today;

            chkProdDueDate.IsChecked = true;
            dtpProdDueDate.SelectedDate = DateTime.Today;

            chkProdCompDate.IsChecked = true;
            dtpProdCompDate.SelectedDate = DateTime.Today;

            chkSetInitHitCountDate.IsChecked = true;
            dtpSetInitHitCountDate.SelectedDate = DateTime.Today;


            txtMoldID.Text = string.Empty;
        }

        //수정
        private void btnUpdate_Click(object sender, RoutedEventArgs e)
        {
            WinMold = dgdMain.SelectedItem as Win_dvl_Molding_U_CodeView;

            if (WinMold != null)
            {
                rowNum = dgdMain.SelectedIndex;
                dgdMain.IsHitTestVisible = false;
                tbkMsg.Text = "자료 수정 중";
                lblMsg.Visibility = Visibility.Visible;
                CantBtnControl();
                strFlag = "U";
            }
        }

        //삭제
        private void btnDelete_Click(object sender, RoutedEventArgs e)
        {
            WinMold = dgdMain.SelectedItem as Win_dvl_Molding_U_CodeView;

            if (WinMold == null)
            {
                MessageBox.Show("삭제할 데이터가 지정되지 않았습니다. 삭제데이터를 지정하고 눌러주세요");
                return;
            }
            else
            {
                if (dgdMain.SelectedIndex == 0)
                    rowNum = 0;
                else
                    rowNum = dgdMain.SelectedIndex - 1;

                if (MessageBox.Show("선택하신 항목을 삭제하시겠습니까?", "삭제 전 확인", MessageBoxButton.YesNo) == MessageBoxResult.Yes)
                {
                    if (DeleteData(WinMold.MoldID))
                    {
                        re_Search(rowNum);
                    }
                }
            }
        }

        //닫기
        private void btnClose_Click(object sender, RoutedEventArgs e)
        {
            DataStore.Instance.InsertLogByFormS(this.GetType().Name, stDate, stTime, "E");
            Lib.Instance.ChildMenuClose(this.ToString());
        }

        //검색
        private void btnSearch_Click(object sender, RoutedEventArgs e)
        {
            rowNum = 0;
            re_Search(rowNum);
        }

        //저장
        private void btnSave_Click(object sender, RoutedEventArgs e)
        {
            if (SaveData(strFlag, txtMoldID.Text))
            {
                CanBtnControl();
                if (dgdPartsCode.Items.Count > 0)
                {
                    dgdPartsCode.Items.Clear();
                }

                re_Search(rowNum);
                strFlag = string.Empty;
                dgdMain.IsHitTestVisible = true;
            }
            else
            {
                MessageBox.Show("저장실패");
            }
        }

        //취소
        private void btnCancel_Click(object sender, RoutedEventArgs e)
        {
            CanBtnControl();
            if (dgdPartsCode.Items.Count > 0)
            {
                dgdPartsCode.Items.Clear();
            }

            if (!strFlag.Equals(string.Empty))
            {
                re_Search(rowNum);
            }

            strFlag = string.Empty;
            dgdMain.IsHitTestVisible = true;
        }

        //엑셀
        private void btnExcel_Click(object sender, RoutedEventArgs e)
        {
            DataTable dt = null;
            string Name = string.Empty;
            Lib lib = new Lib();

            string[] lst = new string[2];
            lst[0] = "금형 현황";
            lst[1] = dgdMain.Name;

            ExportExcelxaml ExpExc = new ExportExcelxaml(lst);

            ExpExc.ShowDialog();

            if (ExpExc.DialogResult.HasValue)
            {
                if (ExpExc.choice.Equals(dgdMain.Name))
                {
                    DataStore.Instance.InsertLogByForm(this.GetType().Name, "E");
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
        }

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
                string sql = string.Empty;
                DataSet ds = null;

                Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                sqlParameter.Clear();
                sqlParameter.Add("chkDate", chkDate.IsChecked == true ? 1 : 0);
                sqlParameter.Add("FromDate", chkDate.IsChecked == true ? dtpSDate.SelectedDate.Value.ToString("yyyyMMdd") : "");
                sqlParameter.Add("ToDate", chkDate.IsChecked == true ? dtpEDate.SelectedDate.Value.ToString("yyyyMMdd") : "");
                sqlParameter.Add("nchkMold", chkMoldNoSrh.IsChecked == true ? 1 : 0);            //금형번호
                sqlParameter.Add("MoldNo", chkMoldNoSrh.IsChecked == true ? txtMoldNoSrh.Text : "");

                sqlParameter.Add("nchkBuyerArticle", chkArticleSrh.IsChecked == true ? 1 : 0);   //품번
                sqlParameter.Add("BuyerArticle", chkArticleSrh.IsChecked == true ? (txtArticleSrh.Tag != null ? txtArticleSrh.Tag.ToString() : "") : "");
                sqlParameter.Add("chkArticle",  0);   
                sqlParameter.Add("ArticleID", "");

                sqlParameter.Add("nNeedInspect", chkNeedInspectSrh.IsChecked == true ? 1 : 0); // 금형점검필요
                sqlParameter.Add("nCheckExpired", chkExpiredSrh.IsChecked == true ? 1 : 0);  //사용기한 경과 
                sqlParameter.Add("nCheckWashingMold", chkNeedWashing.IsChecked == true ? 1 : 0);   //세척필요
                sqlParameter.Add("ChkIncDisCardYN", chkDisCardSrh.IsChecked == true ? "Y" : "N"); //폐기건 

                ds = DataStore.Instance.ProcedureToDataSet("xp_dvlMold_sMold", sqlParameter, false);

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
                                var WinMolding = new Win_dvl_Molding_U_CodeView()
                                {
                                    Num = i + 1,
                                    MoldID = dr["MoldID"].ToString(), //금형번호

                                    MoldNo = dr["MoldNo"].ToString(), //금형명 
                                    MoldTypeID = dr["MoldTypeID"].ToString(), //금형종류
                                    MoldType = dr["MoldType"].ToString(), //금형종류
                                    MoldKind = dr["MoldKind"].ToString(), //금형종류

                                    BuyerModelID = dr["BuyerModelID"].ToString(),  //차종ID
                                    BuyerModel = dr["BuyerModel"].ToString(),  //차종
                                    CustomID = dr["CustomID"].ToString(), //고객사ID 
                                    KCustom = dr["KCustom"].ToString(), //고객사명 
                                    BuyerArticleNo = dr["BuyerArticleNo"].ToString(), 
                                    Article = dr["Article"].ToString(), 

                                    MoldSizeX = Convert.ToDouble(dr["MoldSizeX"]), //가로
                                    MoldSizeY = Convert.ToDouble(dr["MoldSizeY"]), //세로
                                    MoldSizeH = Convert.ToDouble(dr["MoldSizeH"]), //높이
                                    MoldQuality = dr["MoldQuality"].ToString(), //재질
                                    Weight = Convert.ToDouble(dr["Weight"]), //중량

                                    DisCardYN = dr["DisCardYN"].ToString(), //사용여부(폐기건)
                                    Cavity = dr["Cavity"].ToString(),
                                    RealCavity = dr["RealCavity"].ToString(),
                                    Storage = dr["Storage"].ToString(),
                                    StorageName = dr["StorageName"].ToString(),

                                    ProdCustomName = dr["ProdCustomName"].ToString(), //금형제작업체
                                    OwnerCustomName = dr["OwnerCustomName"].ToString(), //금형소유업체
                                    OwnerOneTimePayYn = dr["OwnerOneTimePayYn"].ToString(), //일시불 or 상각
                                    OwnerOneTimePayYnName = dr["OwnerOneTimePayYnName"].ToString(), //일시불 or 상각

                                    SetDate = DatePickerFormat(dr["SetDate"].ToString()), //입고일
                                    ProdOrderDate = DatePickerFormat(dr["ProdOrderDate"].ToString()),  //발주일 
                                    ProdDueDate = DatePickerFormat(dr["ProdDueDate"].ToString()), //완료 예정일
                                    ProdCompDate = DatePickerFormat(dr["ProdCompDate"].ToString()), // 완료일 

                                    MainUseYN = dr["MainUseYN"].ToString(), //주(main) 금형 여부 
                                    Comments = dr["Comments"].ToString(), //비고
                                    MoldPerson = dr["MoldPerson"].ToString(), //관리담당

                                    SetCheckProdQty = Convert.ToDouble(dr["SetCheckProdQty"]), // 점검주기 타발수
                                    AfterRepairHitcount = Convert.ToDouble(dr["AfterRepairHitcount"]), //점검 후 타발수
                                    SetWashingProdQty = Convert.ToDouble(dr["SetWashingProdQty"]), // 세척주기 타발수
                                    AfterWashHitcount = Convert.ToDouble(dr["AfterWashHitcount"]), //세척 후 타발수

                                    SetProdQty = Convert.ToDouble(dr["SetProdQty"]), // 수명 타발수 
                                    HitCount = Convert.ToDouble(dr["HitCount"]), //현재 타발수

                                    SetHitCount = Convert.ToDouble(dr["SetHitCount"]), // 초기설정 타발수
                                    SetHitCountDate = DatePickerFormat(dr["SetHitCountDate"].ToString()), // 초기설정일 
                               
                                    EvalGrade = dr["EvalGrade"].ToString(), //등급
                                    EvalScore = Convert.ToDouble(dr["EvalScore"]), //점수

                                    AttFile1 = dr["AttFile1"].ToString(),
                                    AttPath1 = dr["AttPath1"].ToString(),
                                    AttFile2 = dr["AttFile2"].ToString(),
                                    AttPath2 = dr["AttPath2"].ToString(),
                                    AttFile3 = dr["AttFile3"].ToString(),
                                    AttPath3 = dr["AttPath3"].ToString(),
                                };

                            i++;
                            dgdMain.Items.Add(WinMolding);


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

        //셀렉션item, selectedItem 시 이벤트
        private void dgdMain_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            WinMold = dgdMain.SelectedItem as Win_dvl_Molding_U_CodeView;

            if (WinMold != null)
            {
                this.DataContext = WinMold;
                FillGridPasts(WinMold.MoldID);
                FillGridArticles(WinMold.MoldID);

                chkSetDate.IsChecked = !string.IsNullOrWhiteSpace(WinMold.SetDate);
                chkProdOrderDate.IsChecked = !string.IsNullOrWhiteSpace(WinMold.ProdOrderDate);
                chkProdDueDate.IsChecked = !string.IsNullOrWhiteSpace(WinMold.ProdDueDate);
                chkProdCompDate.IsChecked = !string.IsNullOrWhiteSpace(WinMold.ProdCompDate);
                chkSetInitHitCountDate.IsChecked = !string.IsNullOrWhiteSpace(WinMold.SetHitCountDate);
            }
        }

        private void FillGridPasts(string strMoldID)
        {
            if (dgdPartsCode.Items.Count > 0)
            {
                dgdPartsCode.Items.Clear();
            }

            try
            {
                DataSet ds = null;
                Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                sqlParameter.Clear();
                sqlParameter.Add("MoldID", strMoldID);
                sqlParameter.Add("McPartID", "");
                sqlParameter.Add("ChangeCheckGbn", "");
                ds = DataStore.Instance.ProcedureToDataSet("xp_dvlMold_sMoldChangeProd", sqlParameter, false);

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
                            i++;
                            var WinMoldParts = new Win_dvl_Molding_U_Parts_CodeView()
                            {
                                MoldID = dr["MoldID"].ToString(),
                                Num = i,
                                McPartID = dr["McPartID"].ToString(),
                                MCPartName = dr["MCPartName"].ToString(),
                            };

                            dgdPartsCode.Items.Add(WinMoldParts);
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
        private void FillGridArticles(string strMoldID)
        {
            if (dgdMoldArticle.Items.Count > 0)
            {
                dgdMoldArticle.Items.Clear();
            }

            try
            {
                DataSet ds = null;
                Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                sqlParameter.Clear();
                sqlParameter.Add("MoldID", strMoldID);
                ds = DataStore.Instance.ProcedureToDataSet("xp_dvlMold_sMoldArticle", sqlParameter, false);

                if (ds != null && ds.Tables.Count > 0)
                {
                    DataTable dt = ds.Tables[0];
                    int i = 0;

                    if (dt.Rows.Count != 0)
                    {
                        DataRowCollection drc = dt.Rows;

                        foreach (DataRow dr in drc)
                        {
                            i++;
                            var WinMoldArticle = new MoldArticle_CodeView()
                            {
                                Num = i,
                                ArticleID = dr["ArticleID"].ToString(),
                                Article = dr["Article"].ToString(),
                                BuyerArticleNo = dr["BuyerArticleNo"].ToString(),
                            };

                            dgdMoldArticle.Items.Add(WinMoldArticle);
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

        /// <summary>
        /// 저장
        /// </summary>
        /// <param name="strFlag"></param>
        /// <param name="strYYYY"></param>
        /// <returns></returns>
        private bool SaveData(string strFlag, string strMoldID)
        {
            bool flag = false;
            List<Procedure> Prolist = new List<Procedure>();
            List<Dictionary<string, object>> ListParameter = new List<Dictionary<string, object>>();

            try
            {
                if (CheckData())
                {
                    Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                    sqlParameter.Add("MoldID", strMoldID);
                    sqlParameter.Add("CompanyID", "0001");
                    sqlParameter.Add("MoldNo", txtMoldNo.Text);
                    sqlParameter.Add("CustomID", txtKCustom.Tag?.ToString() ?? "");
                    sqlParameter.Add("BuyerModelID", txtBuyerModel.Tag?.ToString() ?? "");

                    sqlParameter.Add("MoldKind", txtMoldKind.Text ?? "");
                    sqlParameter.Add("MoldTypeID", "");
                    sqlParameter.Add("MoldQuality", txtMoldQuality.Text ?? "");
                    sqlParameter.Add("MoldSizeX", double.TryParse(txtMoldSizeX.Text, out double x) ? x : 0);
                    sqlParameter.Add("MoldSizeY", double.TryParse(txtMoldSizeY.Text, out double y) ? y : 0);

                    sqlParameter.Add("MoldSizeH", double.TryParse(txtMoldSizeH.Text, out double h) ? h : 0);
                    sqlParameter.Add("Weight", double.TryParse(txtWeight.Text, out double weight) ? weight : 0);
                    sqlParameter.Add("DisCardYN", cboDisCard.SelectedValue?.ToString() ?? "");
                    sqlParameter.Add("Cavity", double.TryParse(txtCavity.Text, out double cavity) ? cavity : 0);
                    sqlParameter.Add("RealCavity", double.TryParse(txtRealCavity.Text, out double realCavity) ? realCavity : 0);

                    sqlParameter.Add("Storage", cboStorgeLocation.SelectedValue?.ToString() ?? "");
                    sqlParameter.Add("ProdCustomName", txtProdCustomName.Text ?? "");
                    sqlParameter.Add("OwnerCustomName", TextBoxOwnerCustomName.Text ?? "");
                    sqlParameter.Add("OwnerOneTimePayYn", cboBoxOwnerOneTimePayYn.SelectedValue?.ToString() ?? "");
                    sqlParameter.Add("SetDate", chkSetDate.IsChecked == true && dtpSetDate.SelectedDate != null? dtpSetDate.SelectedDate.Value.ToString("yyyyMMdd") : "");
                    
                    sqlParameter.Add("ProdOrderDate", chkProdOrderDate.IsChecked == true && dtpProdOrderDate.SelectedDate.Value != null ? dtpProdOrderDate.SelectedDate.Value.ToString("yyyyMMdd") : "");
                    sqlParameter.Add("ProdDueDate", chkProdDueDate.IsChecked == true && dtpProdDueDate.SelectedDate.Value != null ? dtpProdDueDate.SelectedDate.Value.ToString("yyyyMMdd") : "");
                    sqlParameter.Add("ProdCompDate", chkProdCompDate.IsChecked == true && dtpProdCompDate.SelectedDate.Value != null ? dtpProdCompDate.SelectedDate.Value.ToString("yyyyMMdd") : "");
                    sqlParameter.Add("MainUseYN", cboMainUseYN.SelectedValue?.ToString() ?? "");
                    sqlParameter.Add("Comments", txtComments.Text);

                    sqlParameter.Add("MoldPerson", txtMoldPerson.Text);
                    sqlParameter.Add("SetCheckProdQty", double.TryParse(txtSetCheckProdQty.Text, out double setCheckQty) ? setCheckQty : 0);
                    sqlParameter.Add("SetWashingProdQty", double.TryParse(txtSetWashingProdQty.Text, out double setWashingQty) ? setWashingQty : 0);
                    sqlParameter.Add("SetProdQty", double.TryParse(txtSetProdQty.Text, out double setProdQty) ? setProdQty : 0);
                    sqlParameter.Add("SetHitCount", double.TryParse(txtSetinitHitCount.Text, out double setHitCount) ? setHitCount : 0);
                   
                    sqlParameter.Add("SetHitCountDate", chkSetInitHitCountDate.IsChecked == true && dtpSetInitHitCountDate.SelectedDate.Value != null ? dtpSetInitHitCountDate.SelectedDate.Value.ToString("yyyyMMdd") : "");
                    sqlParameter.Add("EvalGrade", txtEvalGrade.Text ?? "");
                    sqlParameter.Add("EvalScore", double.TryParse(txtEvalScore.Text, out double score) ? score : 0);
                    sqlParameter.Add("UserID", MainWindow.CurrentUser);

                    #region 추가

                    Procedure pro1 = new Procedure();

                    if (strFlag.Equals("I")){
                                
                        pro1.Name = "xp_dvlMold_iMold";
                        pro1.OutputUseYN = "Y";
                        pro1.OutputName = "MoldID";
                        pro1.OutputLength = "5";
                    } else {
                        pro1.Name = "xp_dvlMold_uMold";
                        pro1.OutputUseYN = "N";
                        pro1.OutputName = "MoldID";
                        pro1.OutputLength = "5";
                    }

                        Prolist.Add(pro1);
                        ListParameter.Add(sqlParameter);

                        for (int i = 0; i < dgdPartsCode.Items.Count; i++)
                        {
                            WinMoldParts = dgdPartsCode.Items[i] as Win_dvl_Molding_U_Parts_CodeView;

                        if(string.IsNullOrWhiteSpace(WinMoldParts.McPartID))
                        {
                            MessageBox.Show("부품이 입력되지 않았습니다");
                            return false;
                        }

                            sqlParameter = new Dictionary<string, object>();
                            sqlParameter.Clear();
                            sqlParameter.Add("MoldID", strMoldID);
                            sqlParameter.Add("McPartID", WinMoldParts.McPartID);
                            sqlParameter.Add("ChangeCheckGbn", 1);
                            sqlParameter.Add("CycleProdQty", 0);
                            sqlParameter.Add("StartSetProdQty", 0);
                            sqlParameter.Add("StartSetDate", DateTime.Today.ToString("yyyyMMdd"));
                            sqlParameter.Add("CreateUserID", MainWindow.CurrentUser);

                            Procedure pro2 = new Procedure();
                            pro2.Name = "xp_dvlMold_iMoldChangeProd";
                            pro2.OutputUseYN = "N";
                            pro2.OutputName = "MoldID";
                            pro2.OutputLength = "5";

                            Prolist.Add(pro2);
                            ListParameter.Add(sqlParameter);
                        }

                        for (int i = 0; i < dgdMoldArticle.Items.Count; i++)
                        {
                            MoldArticleList = dgdMoldArticle.Items[i] as MoldArticle_CodeView;

                        if (string.IsNullOrWhiteSpace(MoldArticleList.ArticleID))
                        {
                            MessageBox.Show("품번이 입력되지 않았습니다");
                            return false;
                        }

                        sqlParameter = new Dictionary<string, object>();
                            sqlParameter.Clear();
                            sqlParameter.Add("MoldID", strMoldID);
                            sqlParameter.Add("ArticleID", MoldArticleList.ArticleID);
                            sqlParameter.Add("CreateUserID", MainWindow.CurrentUser);

                            Procedure pro3 = new Procedure();
                            pro3.Name = "xp_Mold_iMoldArticleData";
                            pro3.OutputUseYN = "N";
                            pro3.OutputName = "MoldID";
                            pro3.OutputLength = "5";

                            Prolist.Add(pro3);
                            ListParameter.Add(sqlParameter);
                        }

                        List<KeyValue> list_Result = new List<KeyValue>();
                        list_Result = DataStore.Instance.ExecuteAllProcedureOutputGetCS(Prolist, ListParameter);
                        string sGetID = string.Empty;

                        if (list_Result[0].key.ToLower() == "success")
                        {
                            list_Result.RemoveAt(0);
                            for (int i = 0; i < list_Result.Count; i++)
                            {
                                KeyValue kv = list_Result[i];
                                if (kv.key == "MoldID")
                                {
                                    sGetID = kv.value;
                                }
                            }

                        flag = true;

                            if (flag)
                            {
                                bool AttachYesNo = false;
                                if (txtAttFile1.Text != string.Empty)       //첨부파일 1
                                {
                                    AttachYesNo = true;
                                    FTP_Save_File(sGetID, txtAttFile1.Text, FullPath1);
                                }
                                if (txtAttFile2.Text != string.Empty)       //첨부파일 2
                                {
                                    AttachYesNo = true;
                                    FTP_Save_File(sGetID, txtAttFile2.Text, FullPath2);
                                }
                                if (txtAttFile3.Text != string.Empty)       //첨부파일 3
                                {
                                    AttachYesNo = true;
                                    FTP_Save_File(sGetID, txtAttFile3.Text, FullPath3);
                                }
                                if (AttachYesNo == true) { AttachFileUpdate(sGetID); }
                            }
                        }
                        else
                        {
                            MessageBox.Show("[저장실패]\r\n" + list_Result[0].value.ToString());
                            flag = false;
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
                sqlParameter.Add("MoldID", strID);

                string[] result = DataStore.Instance.ExecuteProcedure("xp_dvlMold_dMold", sqlParameter, false);

                if (result[0].Equals("success"))
                {
                    //MessageBox.Show("성공 *^^*");
                    flag = true;
                }
                else
                {
                    MessageBox.Show("삭제 실패, 실패 이유 : " + result[1]);
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
        /// 입력사항 체크
        /// 금형LotNo, 차종, 품번, 품명, 고객사명, 보관장소 필수
        /// </summary>
        /// <returns></returns>
        private bool CheckData()
        {
            bool flag = true;

            if (string.IsNullOrWhiteSpace(txtMoldNo.Text))
            {
                flag = false;
                MessageBox.Show("금형명를 입력해주세요.", "필수입력 오류");
                return flag;
            }

            if (string.IsNullOrWhiteSpace(txtMoldKind.Text))
            {
                flag = false;
                MessageBox.Show("금형종류를 입력해주세요.", "필수입력 오류");
                return flag;
            }

            if (dgdMoldArticle.Items.Count == 0)
            {
                flag = false;
                MessageBox.Show("품번/품명을 등록해주세요.");
                btnMoldArticleAdd_Click(null, null);
                return flag;
            }

            //고객사명 txtKCustom
            if (string.IsNullOrWhiteSpace(txtKCustom.Text))
            {
                flag = false;
                MessageBox.Show("고객사명을 입력해주세요.", "필수입력 오류");
                return flag;
            }

         
            return flag;
        }

        private void btnSubAdd_Click(object sender, RoutedEventArgs e)
        {
            Win_dvl_Molding_U_Parts_CodeView PartsMold = new Win_dvl_Molding_U_Parts_CodeView()
            {
                Num = dgdPartsCode.Items.Count + 1,
                McPartID = "",
                MCPartName = "",
                MoldID = ""
            };

            dgdPartsCode.Items.Add(PartsMold);
        }

        private void btnSubDel_Click(object sender, RoutedEventArgs e)
        {
            WinMoldParts = dgdPartsCode.CurrentItem as Win_dvl_Molding_U_Parts_CodeView;

            if (WinMoldParts != null)
            {
                dgdPartsCode.Items.Remove(WinMoldParts);
            }
            else
            {
                if (dgdPartsCode.Items.Count > 0)
                {
                    dgdPartsCode.Items.RemoveAt(dgdPartsCode.Items.Count - 1);
                }
            }
        }

        //
        private void DataGridCell_KeyDown(object sender, KeyEventArgs e)
        {
            WinMoldParts = dgdPartsCode.CurrentItem as Win_dvl_Molding_U_Parts_CodeView;
            int rowCount = dgdPartsCode.Items.IndexOf(dgdPartsCode.CurrentItem);
            int colCountOne = dgdPartsCode.Columns.IndexOf(dgdtpePartsName);
            int colCountTwo = dgdPartsCode.Columns.IndexOf(dgdtpePartsCode);
            int colCount = dgdPartsCode.Columns.IndexOf(dgdPartsCode.CurrentCell.Column);

            if (e.Key == Key.Enter)
            {
                e.Handled = true;
                (sender as DataGridCell).IsEditing = false;

                if (dgdPartsCode.Items.Count - 1 > rowCount && colCount == colCountTwo)
                {
                    dgdPartsCode.SelectedIndex = rowCount + 1;
                    dgdPartsCode.CurrentCell =
                        new DataGridCellInfo(dgdPartsCode.Items[rowCount + 1], dgdPartsCode.Columns[colCountOne]);
                }
                else if (dgdPartsCode.Items.Count - 1 >= rowCount && colCount == colCountOne)
                {
                    dgdPartsCode.CurrentCell =
                        new DataGridCellInfo(dgdPartsCode.Items[rowCount], dgdPartsCode.Columns[colCountTwo]);
                }
                else if (dgdPartsCode.Items.Count - 1 == rowCount && colCount == colCountTwo)
                {
                    if (MessageBox.Show("부품을 추가하시겠습니까?", "추가 확인", MessageBoxButton.YesNo) == MessageBoxResult.Yes)
                    {
                        btnSubAdd_Click(null, null);
                        dgdPartsCode.SelectedIndex = rowCount + 1;
                        dgdPartsCode.CurrentCell =
                            new DataGridCellInfo(dgdPartsCode.Items[rowCount + 1], dgdPartsCode.Columns[colCountOne]);
                    }
                    else
                    {
                        btnSave.Focus();
                    }
                }
                else
                {
                    MessageBox.Show("있으면 찾아보자...");
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
                cell.IsEditing = true;
            }
        }

       

        private void dgdtxtMCPartName_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                WinMoldParts = dgdPartsCode.CurrentItem as Win_dvl_Molding_U_Parts_CodeView;

                if (WinMoldParts != null)
                {
                    TextBox tb1 = sender as TextBox;
                    MainWindow.pf.ReturnCode(tb1, (int)Defind_CodeFind.DCF_PART, "");

                    if (tb1.Tag != null)
                    {
                        WinMoldParts.McPartID = tb1.Tag.ToString();
                        WinMoldParts.MCPartName = tb1.Text;
                    }

                    sender = tb1;
                }
            }
        }

       
        private void dgdtxtMCPartID_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                WinMoldParts = dgdPartsCode.CurrentItem as Win_dvl_Molding_U_Parts_CodeView;

                if (WinMoldParts != null)
                {
                    TextBox tb1 = sender as TextBox;
                    MainWindow.pf.ReturnCode(tb1, (int)Defind_CodeFind.DCF_PART, "");

                    if (tb1.Tag != null)
                    {
                        WinMoldParts.McPartID = tb1.Tag.ToString();
                        WinMoldParts.MCPartName = tb1.Text;
                    }

                    sender = tb1;
                }
            }
        }

        private void chkProdOrderDate_Checked(object sender, RoutedEventArgs e)
        {
            dtpProdOrderDate.IsEnabled = true;
            if (dtpProdOrderDate.SelectedDate == null) dtpProdOrderDate.SelectedDate = DateTime.Today;
        }

        private void chkProdOrderDate_Unchecked(object sender, RoutedEventArgs e)
        {
            dtpProdOrderDate.IsEnabled = false;
        }

        private void chkProdDueDate_Checked(object sender, RoutedEventArgs e)
        {
            dtpProdDueDate.IsEnabled = true;
            if (dtpProdDueDate.SelectedDate == null) dtpProdDueDate.SelectedDate = DateTime.Today;
        }

        private void chkProdDueDate_Unchecked(object sender, RoutedEventArgs e)
        {
            dtpProdDueDate.IsEnabled = false;
        }
        private void chkSetDate_Checked(object sender, RoutedEventArgs e)
        {
            dtpSetDate.IsEnabled = true;
            if (dtpSetDate.SelectedDate == null) dtpSetDate.SelectedDate = DateTime.Today;
        }

        private void chkSetDate_Unchecked(object sender, RoutedEventArgs e)
        {
            dtpSetDate.IsEnabled = false;
        }

        // 파일 저장하기.
        private void FTP_Save_File(string Defect_ID, string FileName, string FullPath)
        {
            UploadFileInfo fileInfo_up = new UploadFileInfo();
            fileInfo_up.Filename = FileName;
            fileInfo_up.Type = FtpFileType.File;

            _ftp = new FTP_EX(FTP_ADDRESS, FTP_ID, FTP_PASS);

            string[] fileListSimple;
            string[] fileListDetail;

            fileListSimple = _ftp.directoryListSimple("", Encoding.Default);
            fileListDetail = _ftp.directoryListDetailed("", Encoding.Default);

            // 기존 폴더 확인작업.
            bool MakeFolder = false;
            for (int i = 0; i < fileListSimple.Length; i++)
            {
                if (fileListSimple[i] == Defect_ID)
                {
                    MakeFolder = true;
                    break;
                }
            }
            if (MakeFolder == false)        // 같은 아이를 찾지 못한경우,
            {
                //MIL 폴더에 InspectionID로 저장
                if (_ftp.createDirectory(Defect_ID) == false)
                {
                    MessageBox.Show("업로드를 위한 폴더를 생성할 수 없습니다.");
                    return;
                }
            }
            // 폴더 생성 후 생성한 폴더에 파일을 업로드
            string str_remotepath = Defect_ID + "/";
            fileInfo_up.Filepath = str_remotepath;
            str_remotepath += FileName;
            if (_ftp.upload(str_remotepath, FullPath) == false)
            {
                MessageBox.Show("파일업로드에 실패하였습니다.");
                return;
            }

            if (FullPath == FullPath1) { txtAttFile1.Tag = "/ImageData/Mold/" + fileInfo_up.Filepath; }
            if (FullPath == FullPath2) { txtAttFile2.Tag = "/ImageData/Mold/" + fileInfo_up.Filepath; }
            if (FullPath == FullPath3) { txtAttFile3.Tag = "/ImageData/Mold/" + fileInfo_up.Filepath; }
        }

        //파일 삭제(FTP상에서)_폴더 삭제는 X
        private void FTP_UploadFile_File_Delete(string strSaveName, string FileName)
        {
            if (!_ftp.delete(strSaveName + "/" + FileName))
            {
                MessageBox.Show("파일이 삭제되지 않았습니다.");
            }
            //if (_ftp.DeleteFileOnFtpServer(new Uri(FTP_ADDRESS + "/" + strSaveName + "/" + FileName)) == true)
            //{
            //}
            else
            {
                MessageBox.Show("파일이 삭제되지 않았습니다.");
            }
        }

        /// <summary>
        /// FTP 업로드 폴더 삭제(안의 파일을 삭제해야 삭제가 된다.)
        /// </summary>
        /// <param name="strSaveName"></param>
        /// <param name="FileName"></param>
        /// <returns></returns>
        private bool FTP_UploadFile_Path_Delete(string strSaveName)
        {
            _ftp = new FTP_EX(FTP_ADDRESS, FTP_ID, FTP_PASS);

            string[] fileListSimple;
            string[] fileListDetail;

            fileListSimple = _ftp.directoryListSimple("", Encoding.Default);
            fileListDetail = _ftp.directoryListDetailed("", Encoding.Default);

            bool tf_ExistInspectionID = MakeFileInfoList(fileListSimple, fileListDetail, strSaveName);

            if (tf_ExistInspectionID == true)
            {
                if (_ftp.removeDir(strSaveName) == true)
                {
                    return true;
                }
                else
                {
                    return false;
                }
                //if (_ftp.DeleteFileOnFtpServer(new Uri(strSaveName)) == true)
                //{
                //    return true;
                //}
                //else
                //{
                //    return false;
                //}
            }
            return true;
        }

        private void btnInsPic_Click(object sender, RoutedEventArgs e)
        {
            // (버튼)sender 마다 tag를 달자.
            string ClickPoint = ((Button)sender).Tag.ToString();

            Microsoft.Win32.OpenFileDialog OFdlg = new Microsoft.Win32.OpenFileDialog();

            OFdlg.DefaultExt = "*.jpg, *.jpeg, *.jpe, *.jfif, *.png";
            //OFdlg.Filter = "Image files (*.jpg, *.jpeg, *.jpe, *.jfif, *.png) | *.jpg; *.jpeg; *.jpe; *.jfif; *.png | All Files|*.*";
            OFdlg.Filter = "All Files|*.*";

            Nullable<bool> result = OFdlg.ShowDialog();
            if (result == true)
            {
                if (ClickPoint == "1") { FullPath1 = OFdlg.FileName; }  //긴 경로(FULL 사이즈)
                if (ClickPoint == "2") { FullPath2 = OFdlg.FileName; }
                if (ClickPoint == "3") { FullPath3 = OFdlg.FileName; }

                string AttachFileName = OFdlg.SafeFileName;  //명.
                string AttachFilePath = string.Empty;       // 경로

                if (ClickPoint == "1") { AttachFilePath = FullPath1.Replace(AttachFileName, ""); }
                if (ClickPoint == "2") { AttachFilePath = FullPath2.Replace(AttachFileName, ""); }
                if (ClickPoint == "3") { AttachFilePath = FullPath3.Replace(AttachFileName, ""); }

                StreamReader sr = new StreamReader(OFdlg.FileName);
                long File_size = sr.BaseStream.Length;
                if (sr.BaseStream.Length > (2048 * 1000))
                {
                    // 업로드 파일 사이즈범위 초과
                    MessageBox.Show("이미지의 파일사이즈가 2M byte를 초과하였습니다.");
                    sr.Close();
                    return;
                }
                if (ClickPoint == "1")
                {
                    txtAttFile1.Text = AttachFileName;
                    txtAttFile1.Tag = AttachFilePath.ToString();
                }
                else if (ClickPoint == "2")
                {
                    txtAttFile2.Text = AttachFileName;
                    txtAttFile2.Tag = AttachFilePath.ToString();
                }
                else if (ClickPoint == "3")
                {
                    txtAttFile3.Text = AttachFileName;
                    txtAttFile3.Tag = AttachFilePath.ToString();
                }
            }
        }

        private void btnDelPic_Click(object sender, RoutedEventArgs e)
        {
            MessageBoxResult msgresult = MessageBox.Show("파일을 삭제 하시겠습니까?", "삭제 확인", MessageBoxButton.YesNo);
            if (msgresult == MessageBoxResult.Yes)
            {
                string ClickPoint = ((Button)sender).Tag.ToString();

                if ((ClickPoint == "1") && (txtAttFile1.Tag.ToString() != string.Empty))
                {
                    if (strFlag.Equals("U"))
                    {
                        if (DetectFtpFile(txtMoldID.Text) && !txtAttFile1.Text.Equals(string.Empty))
                        {
                            FTP_UploadFile_File_Delete(txtMoldID.Text, txtAttFile1.Text);
                        }
                    }

                    txtAttFile1.Text = string.Empty;
                    txtAttFile1.Tag = string.Empty;

                }
                if ((ClickPoint == "2") && (txtAttFile2.Tag.ToString() != string.Empty))
                {
                    if (strFlag.Equals("U"))
                    {
                        if (DetectFtpFile(txtMoldID.Text) && !txtAttFile2.Text.Equals(string.Empty))
                        {
                            FTP_UploadFile_File_Delete(txtMoldID.Text, txtAttFile2.Text);
                        }
                    }

                    txtAttFile2.Text = string.Empty;
                    txtAttFile2.Tag = string.Empty;
                }
                if ((ClickPoint == "3") && (txtAttFile3.Tag.ToString() != string.Empty))
                {
                    if (strFlag.Equals("U"))
                    {
                        if (DetectFtpFile(txtMoldID.Text) && !txtAttFile3.Text.Equals(string.Empty))
                        {
                            FTP_UploadFile_File_Delete(txtMoldID.Text, txtAttFile3.Text);
                        }
                    }

                    txtAttFile3.Text = string.Empty;
                    txtAttFile3.Tag = string.Empty;
                }
            }
        }

        private void btnPreView_Click(object sender, RoutedEventArgs e)
        {
            MessageBoxResult msgresult = MessageBox.Show("파일을 보시겠습니까?", "보기 확인", MessageBoxButton.YesNo);
            if (msgresult == MessageBoxResult.Yes)
            {
                //버튼 태그값.
                string ClickPoint = ((Button)sender).Tag.ToString();

                if ((ClickPoint == "1") && (txtAttFile1.Tag.ToString() == string.Empty))
                {
                    MessageBox.Show("파일이 없습니다.");
                    return;
                }
                if ((ClickPoint == "2") && (txtAttFile2.Tag.ToString() == string.Empty))
                {
                    MessageBox.Show("파일이 없습니다.");
                    return;
                }
                if ((ClickPoint == "3") && (txtAttFile3.Tag.ToString() == string.Empty))
                {
                    MessageBox.Show("파일이 없습니다.");
                    return;
                }

                // 접속 경로
                _ftp = new FTP_EX(FTP_ADDRESS, FTP_ID, FTP_PASS);

                string[] fileListSimple;
                string[] fileListDetail;

                fileListSimple = _ftp.directoryListSimple("", Encoding.Default);
                fileListDetail = _ftp.directoryListDetailed("", Encoding.Default);

                bool ExistFile = false;

                if (ClickPoint == "1")
                {
                    // 경로에 '\\'가 잘못 들어간 경우 오류가 나 멈춤, 이를 방지하기 위한 조건 추가
                    if (txtAttFile1.Tag.ToString().Contains("\\"))
                    {
                        ExistFile = MakeFileInfoList(fileListSimple, fileListDetail, txtAttFile1.Tag.ToString().Split('\\')[3].Trim());
                    }
                    else
                    {
                        ExistFile = MakeFileInfoList(fileListSimple, fileListDetail, txtAttFile1.Tag.ToString().Split('/')[3].Trim());
                    }

                }  //(폴더경로 찾기.)
                if (ClickPoint == "2")
                {
                    if (txtAttFile2.Tag.ToString().Contains("\\"))
                    {
                        ExistFile = MakeFileInfoList(fileListSimple, fileListDetail, txtAttFile2.Tag.ToString().Split('\\')[3].Trim());
                    }
                    else
                    {
                        ExistFile = MakeFileInfoList(fileListSimple, fileListDetail, txtAttFile2.Tag.ToString().Split('/')[3].Trim());
                    }
                }
                if (ClickPoint == "3")
                {
                    if (txtAttFile3.Tag.ToString().Contains("\\"))
                    {
                        ExistFile = MakeFileInfoList(fileListSimple, fileListDetail, txtAttFile3.Tag.ToString().Split('\\')[3].Trim());
                    }
                    else
                    {
                        ExistFile = MakeFileInfoList(fileListSimple, fileListDetail, txtAttFile3.Tag.ToString().Split('/')[3].Trim());
                    }
                }

                int totalCount = _listFileInfo.Count;

                if (ExistFile == true)
                {
                    ExistFile = false;
                    // 접속 경로
                    string str_path = string.Empty;
                    if (ClickPoint == "1") { str_path = FTP_ADDRESS + '/' + txtAttFile1.Tag.ToString().Split('/')[3].Trim(); }
                    if (ClickPoint == "2") { str_path = FTP_ADDRESS + '/' + txtAttFile2.Tag.ToString().Split('/')[3].Trim(); }
                    if (ClickPoint == "3") { str_path = FTP_ADDRESS + '/' + txtAttFile3.Tag.ToString().Split('/')[3].Trim(); }

                    _ftp = new FTP_EX(str_path, FTP_ID, FTP_PASS);

                    if (ClickPoint == "1") { ExistFile = MakeFileInfoList(fileListSimple, fileListDetail, txtAttFile1.Tag.ToString().Split('/')[3].Trim()); }
                    if (ClickPoint == "2") { ExistFile = MakeFileInfoList(fileListSimple, fileListDetail, txtAttFile2.Tag.ToString().Split('/')[3].Trim()); }
                    if (ClickPoint == "3") { ExistFile = MakeFileInfoList(fileListSimple, fileListDetail, txtAttFile3.Tag.ToString().Split('/')[3].Trim()); }

                    totalCount = _listFileInfo.Count;

                    if (ExistFile == true)
                    {
                        string str_remotepath = string.Empty;
                        string str_localpath = string.Empty;

                        if (ClickPoint == "1") { str_remotepath = txtAttFile1.Text.ToString(); }
                        if (ClickPoint == "1") { str_localpath = LOCAL_DOWN_PATH + "\\" + txtAttFile1.Text.ToString(); }
                        if (ClickPoint == "2") { str_remotepath = txtAttFile2.Text.ToString(); }
                        if (ClickPoint == "2") { str_localpath = LOCAL_DOWN_PATH + "\\" + txtAttFile2.Text.ToString(); }
                        if (ClickPoint == "3") { str_remotepath = txtAttFile3.Text.ToString(); }
                        if (ClickPoint == "3") { str_localpath = LOCAL_DOWN_PATH + "\\" + txtAttFile3.Text.ToString(); }

                        DirectoryInfo DI = new DirectoryInfo(LOCAL_DOWN_PATH);      // Temp 폴더가 없는 컴터라면, 만들어 줘야지.
                        if (DI.Exists == false)
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
                            MessageBox.Show("C:Temp 폴더를 확인하세요.");
                            return;
                        }

                        _ftp.download(str_remotepath, str_localpath);
                        //MessageBox.Show("C:Temp 폴더를 확인하세요.");

                        ProcessStartInfo proc = new ProcessStartInfo(str_localpath);
                        proc.UseShellExecute = true;
                        Process.Start(proc);
                    }
                }
                else
                {
                    MessageBox.Show("파일을 찾을 수 없습니다.");
                }
            }
        }

        private bool MakeFileInfoList(string[] simple, string[] detail, string str_ID)
        {
            bool tf_return = false;
            foreach (string filename in simple)
            {
                foreach (string info in detail)
                {
                    if (info.Contains(filename) == true)
                    {

                        if (MakeFileInfoList(filename, info, str_ID) == true)
                        {
                            tf_return = true;
                        }
                    }
                }
            }
            return tf_return;
        }

        private bool MakeFileInfoList(string simple, string detail, string strCompare)
        {
            UploadFileInfo info = new UploadFileInfo();
            info.Filename = simple;
            info.Filepath = detail;

            if (simple.Length > 0)
            {
                string[] tokens = detail.Split(new[] { ' ' }, 9, StringSplitOptions.RemoveEmptyEntries);
                string name = tokens[3].ToString();         // 2017.03.16  허윤구.  토근 배열이 8자리로 되어 있었는데 에러가 나길래 확인해 보니 4자리 배열로 나오길래 바꾸었습니다.
                string permissions = tokens[2].ToString();      // premission도 배열 0번이 아니라 배열 2번인데...;;


                if (permissions.Contains("D") == true)          // 대문자 D로 표시해야 합니다.
                {
                    info.Type = FtpFileType.DIR;
                }
                else
                {
                    info.Type = FtpFileType.File;
                }

                if (info.Type == FtpFileType.File)
                {
                    info.Size = Convert.ToInt64(detail.Substring(17, detail.LastIndexOf(simple) - 17).Trim());      // 사이즈가 중요한가?
                }

                _listFileInfo.Add(info);

                if (string.Compare(simple, strCompare, false) == 0)
                    return true;
            }

            return false;
        }

        /// <summary>
        /// 삭제할 파일이 존재하는지 확인, strSaveName = FullPath(파일이름 포함)
        /// </summary>
        /// <param name="strSaveName"></param>
        /// <returns></returns>
        private bool DetectFtpFile(string strSaveName)
        {
            _ftp = new FTP_EX(FTP_ADDRESS, FTP_ID, FTP_PASS);

            string[] fileListSimple;
            string[] fileListDetail;

            fileListSimple = _ftp.directoryListSimple("", Encoding.Default);
            fileListDetail = _ftp.directoryListDetailed("", Encoding.Default);

            bool tf_ExistInspectionID = MakeFileInfoList(fileListSimple, fileListDetail, strSaveName);

            return tf_ExistInspectionID;
        }

        // 1) 첨부문서가 있을경우, 2) FTP에 정상적으로 업로드가 완료된 경우.  >> DB에 정보 업데이트 
        private void AttachFileUpdate(string ID)
        {
            try
            {
                Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                sqlParameter.Clear();
                sqlParameter.Add("MoldID", ID);

                sqlParameter.Add("AttPath1", txtAttFile1.Text.Equals(string.Empty) ? "" : txtAttFile1.Tag.ToString());
                sqlParameter.Add("AttFile1", txtAttFile1.Text);
                sqlParameter.Add("AttPath2", txtAttFile2.Text.Equals(string.Empty) ? "" : txtAttFile2.Tag.ToString());
                sqlParameter.Add("AttFile2", txtAttFile2.Text);
                sqlParameter.Add("AttPath3", txtAttFile3.Text.Equals(string.Empty) ? "" : txtAttFile3.Tag.ToString());
                sqlParameter.Add("AttFile3", txtAttFile3.Text);

                string[] result = DataStore.Instance.ExecuteProcedure("xp_dvlMold_uMolde_Ftp", sqlParameter, true);
                if (!result[0].Equals("success"))
                {
                    MessageBox.Show("이상발생, 관리자에게 문의하세요");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }

        //private void txtMoldKind_KeyDown(object sender, KeyEventArgs e)
        //{
        //    if (e.Key == Key.Enter)
        //    {
        //        MainWindow.pf.ReturnCode(txtMoldKind, (int)Defind_CodeFind.LG_MOLDN, "");
        //    }
        //}

        //private void btnPfMoldKind_Click(object sender, RoutedEventArgs e)
        //{
        //    MainWindow.pf.ReturnCode(txtMoldKind, (int)Defind_CodeFind.LG_MOLDN, "");
        //}

        private void txtArticle_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                //MainWindow.pf.ReturnCode(txtArticle, 1, "");
                //SetBuyerArticleNo(txtArticle.Tag);
            }
        }

        private void btnPfArticle_Click(object sender, RoutedEventArgs e)
        {
            //MainWindow.pf.ReturnCode(txtArticle, 1, "");
            //SetBuyerArticleNo(txtArticle.Tag);
        }

        //buyerArticleNo 세팅..
        private void SetBuyerArticleNo(object obj)
        {
            try
            {
                string strArticleID = string.Empty;

                if (obj != null)
                {
                    string sql = "select ma.BuyerArticleNo, ma.Article, ma.ArticleID from mt_Article as ma ";
                    sql += "where ma.ArticleID = '" + obj.ToString() + "'   ";

                    DataSet ds = DataStore.Instance.QueryToDataSet(sql);
                    if (ds != null && ds.Tables.Count > 0)
                    {
                        DataTable dt = ds.Tables[0];
                        if (dt.Rows.Count > 0)
                        {
                            //txtBuyerArticleNo.Text = dt.Rows[0].ItemArray[0].ToString();
                            //txtArticle.Text = dt.Rows[0].ItemArray[1].ToString();
                            //txtArticle.Tag = dt.Rows[0].ItemArray[2].ToString();


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

        private void txtKCustom_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                MainWindow.pf.ReturnCode(txtKCustom, 0, "");
            }
        }
        private void txtBuyerModel_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                MainWindow.pf.ReturnCode(txtBuyerModel, (int)Defind_CodeFind.DCF_BUYERMODEL, "");
            }
        }

        private void btnPfBuyerModel_Click(object sender, RoutedEventArgs e)
        {
            MainWindow.pf.ReturnCode(txtBuyerModel, (int)Defind_CodeFind.DCF_BUYERMODEL, "");
        }

        private void lblSetInitHitCountDate_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            if (chkSetInitHitCountDate.IsChecked == true) { chkSetInitHitCountDate.IsChecked = false; }
            else { chkSetInitHitCountDate.IsChecked = true; }
        }

        private void lblProdOrderDate_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            if (chkProdOrderDate.IsChecked == true) { chkProdOrderDate.IsChecked = false; dtpProdOrderDate.IsEnabled = false; }
            else { chkProdOrderDate.IsChecked = true; dtpProdOrderDate.IsEnabled = true; }
        }

        private void lblProdDueDate_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            if (chkProdDueDate.IsChecked == true) { chkProdDueDate.IsChecked = false; dtpProdDueDate.IsEnabled = false; }
            else { chkProdDueDate.IsChecked = true; dtpProdDueDate.IsEnabled = true; }
        }
        private void lblProdCompDate_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            if (chkProdCompDate.IsChecked == true) { chkProdCompDate.IsChecked = false; dtpProdCompDate.IsEnabled = false; }
            else { chkProdCompDate.IsChecked = true; dtpProdCompDate.IsEnabled = true; }
        }

        private void lblSetDate_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            if (chkSetDate.IsChecked == true) { chkSetDate.IsChecked = false; dtpSetDate.IsEnabled = false; }
            else { chkSetDate.IsChecked = true; dtpSetDate.IsEnabled = true; }
        }

        private void chkSetInitHitCountDate_Checked(object sender, RoutedEventArgs e)
        {
            dtpSetInitHitCountDate.IsEnabled = true;
            if(dtpSetInitHitCountDate.SelectedDate == null) dtpSetInitHitCountDate.SelectedDate = DateTime.Today;
        }

        private void chkSetInitHitCountDate_Unchecked(object sender, RoutedEventArgs e)
        {
            dtpSetInitHitCountDate.IsEnabled = false;
        }

        private void chkProdCompDate_Checked(object sender, RoutedEventArgs e)
        {
            dtpProdCompDate.IsEnabled = true;
            if (dtpProdCompDate.SelectedDate == null) dtpProdCompDate.SelectedDate = DateTime.Today;
        }

        private void chkProdCompDate_Unchecked(object sender, RoutedEventArgs e)
        {
            dtpProdCompDate.IsEnabled = false;
        }

        private void txtBuyerArticleNo_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                //MainWindow.pf.ReturnCode(txtBuyerArticleNo, 76, "");

                var MoldArticle = dgdMoldArticle.CurrentItem as MoldArticle_CodeView;

                if (MoldArticle != null)
                {
                    TextBox tb = new TextBox();

                    MainWindow.pf.ReturnCode(tb, 76, MoldArticle.BuyerArticleNo == null ? "" : MoldArticle.BuyerArticleNo);

                    e.Handled = true;

                    if (!tb.Text.Equals("") && tb.Tag != null && !tb.Tag.ToString().Equals(""))
                    {
                        ArticleInfo ai = getArticleInfo(tb.Tag.ToString());

                        if (ai != null)
                        {
                            MoldArticle.BuyerArticleNo = ai.BuyerArticleNo;
                            MoldArticle.ArticleID = ai.ArticleID;
                            MoldArticle.Article = ai.Article;
                        }
                    }
                }

                for (int i = 0; i < 3; i++)
                {
                    int currRow = dgdMoldArticle.Items.IndexOf(dgdMoldArticle.CurrentItem);

                    dgdMoldArticle.SelectedIndex = currRow;
                    dgdMoldArticle.CurrentCell = new DataGridCellInfo(dgdMoldArticle.Items[currRow], dgdMoldArticle.Columns[i]);
                }
            }
        }

        private ArticleInfo getArticleInfo(string ArticleID)
        {
            var getArticleInfo = new ArticleInfo();

            try
            {
                Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                sqlParameter.Add("ArticleID", ArticleID);

                DataSet ds = DataStore.Instance.ProcedureToDataSet("xp_Mold_sArticleData", sqlParameter, false);
                if (ds != null && ds.Tables.Count > 0)
                {
                    DataTable dt = ds.Tables[0];
                    if (dt.Rows.Count > 0)
                    {
                        DataRow dr = dt.Rows[0];

                        getArticleInfo = new ArticleInfo
                        {
                            Article = dr["Article"].ToString(),
                            ArticleID = dr["ArticleID"].ToString(),
                            BuyerArticleNo = dr["BuyerArticleNo"].ToString(),
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
            finally
            {
                DataStore.Instance.CloseConnection();
            }
        }

        private void btnPfBuyerArticleNo_Click(object sender, RoutedEventArgs e)
        {
            //MainWindow.pf.ReturnCode(txtBuyerArticleNo, 76, "");
            //SetBuyerArticleNo(txtBuyerArticleNo.Tag);
        }

        private void btnPfKCustom_Click(object sender, RoutedEventArgs e)
        {
            MainWindow.pf.ReturnCode(txtKCustom, 0, "");
            SetBuyerArticleNo(txtKCustom.Tag);
        }



        #region 품명 그리드 이벤트
        private void btnMoldArticleAdd_Click(object sender, RoutedEventArgs e)
        {
            int i = 1;

            if (dgdMoldArticle.Items.Count > 0) { i = dgdMoldArticle.Items.Count + 1; }

            var MoldArticle = new MoldArticle_CodeView()
            {
                Num = i,
                ArticleID = "",
                BuyerArticleNo = "",
                Article = "",
            };

            dgdMoldArticle.Items.Add(MoldArticle);
        }

        private void btnMoldArticleDelete_Click(object sender, RoutedEventArgs e)
        {
            var MoldArticle = dgdMoldArticle.SelectedItem as MoldArticle_CodeView;

            if (MoldArticle != null)
            {
                dgdMoldArticle.Items.Remove(MoldArticle);
            }
            else
            {
                if (dgdMoldArticle.Items.Count > 0)
                {
                    dgdMoldArticle.Items.Remove(dgdMoldArticle.Items[dgdMoldArticle.Items.Count - 1]);
                }
            }
        }
        private void DataGird_KeyDown(object sender, KeyEventArgs e)
        {
            int currRow = dgdMoldArticle.Items.IndexOf(dgdMoldArticle.CurrentItem);
            int currCol = dgdMoldArticle.Columns.IndexOf(dgdMoldArticle.CurrentCell.Column);
            int startCol = 1;
            int endCol = 2;

            if (e.Key == Key.Enter)
            {
                e.Handled = true;
                (sender as DataGridCell).IsEditing = false;

                // 마지막 열, 마지막 행 아님
                if (endCol == currCol && dgdMoldArticle.Items.Count - 1 > currRow)
                {
                    dgdMoldArticle.SelectedIndex = currRow + 1; // 이건 한줄 파란색으로 활성화 된 걸 조정하는 것입니다.
                    dgdMoldArticle.CurrentCell = new DataGridCellInfo(dgdMoldArticle.Items[currRow + 1], dgdMoldArticle.Columns[startCol]);

                } // 마지막 열 아님
                else if (endCol > currCol && dgdMoldArticle.Items.Count - 1 >= currRow)
                {
                    dgdMoldArticle.CurrentCell = new DataGridCellInfo(dgdMoldArticle.Items[currRow], dgdMoldArticle.Columns[currCol + 1]);
                } // 마지막 열, 마지막 행

            }
            else if (e.Key == Key.Down)
            {
                e.Handled = true;
                (sender as DataGridCell).IsEditing = false;

                // 마지막 행 아님
                if (dgdMoldArticle.Items.Count - 1 > currRow)
                {
                    dgdMoldArticle.SelectedIndex = currRow + 1;
                    dgdMoldArticle.CurrentCell = new DataGridCellInfo(dgdMoldArticle.Items[currRow + 1], dgdMoldArticle.Columns[currCol]);
                } // 마지막 행일때
                else if (dgdMoldArticle.Items.Count - 1 == currRow)
                {
                    if (endCol > currCol) // 마지막 열이 아닌 경우, 열을 오른쪽으로 이동
                    {
                        //dgdMoldArticle.SelectedIndex = 0;
                        dgdMoldArticle.CurrentCell = new DataGridCellInfo(dgdMoldArticle.Items[currRow], dgdMoldArticle.Columns[currCol + 1]);
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
                    dgdMoldArticle.SelectedIndex = currRow - 1;
                    dgdMoldArticle.CurrentCell = new DataGridCellInfo(dgdMoldArticle.Items[currRow - 1], dgdMoldArticle.Columns[currCol]);
                } // 첫 행
                else if (dgdMoldArticle.Items.Count - 1 == currRow)
                {
                    if (0 < currCol) // 첫 열이 아닌 경우, 열을 왼쪽으로 이동
                    {
                        //dgdMoldArticle.SelectedIndex = 0;
                        dgdMoldArticle.CurrentCell = new DataGridCellInfo(dgdMoldArticle.Items[currRow], dgdMoldArticle.Columns[currCol - 1]);
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
                    dgdMoldArticle.CurrentCell = new DataGridCellInfo(dgdMoldArticle.Items[currRow], dgdMoldArticle.Columns[currCol - 1]);
                }
                else if (startCol == currCol)
                {
                    if (0 < currRow)
                    {
                        dgdMoldArticle.SelectedIndex = currRow - 1;
                        dgdMoldArticle.CurrentCell = new DataGridCellInfo(dgdMoldArticle.Items[currRow - 1], dgdMoldArticle.Columns[endCol]);
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

                    dgdMoldArticle.CurrentCell = new DataGridCellInfo(dgdMoldArticle.Items[currRow], dgdMoldArticle.Columns[currCol + 1]);
                }
                else if (endCol == currCol)
                {
                    if (dgdMoldArticle.Items.Count - 1 > currRow)
                    {
                        dgdMoldArticle.SelectedIndex = currRow + 1;
                        dgdMoldArticle.CurrentCell = new DataGridCellInfo(dgdMoldArticle.Items[currRow + 1], dgdMoldArticle.Columns[startCol]);
                    }
                    else
                    {
                        //btnSave.Focus();
                    }
                }
            }
        }

        private void DatagridIn_TextFocus(object sender, KeyEventArgs e)
        {
            // 엔터 → 포커스 = true → cell != null → 해당 텍스트박스가 null이 아니라면 
            // → 해당 텍스트박스가 포커스가 안되있음 SelectAll() or 포커스
            Lib.Instance.DataGridINTextBoxFocus(sender, e);
        }

        private void DataGridCell_MouseUp(object sender, MouseButtonEventArgs e)
        {
            DataGridCell cell = sender as DataGridCell;

            Lib.Instance.DataGridINTextBoxFocusByMouseUP(sender, e);
        }

        #endregion

        #region 기타 메서드

        private void onlyNemeric(object sender, TextCompositionEventArgs e)
        {
            lib.CheckIsNumeric((TextBox)sender, e);
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

        #endregion

    }

    class Win_dvl_Molding_U_CodeView : BaseView
    {
        public int Num { get; set; }
        public string MoldID { get; set; }
        public string MoldNo { get; set; }
        public string MoldTypeID { get; set; }
        public string MoldType { get; set; }
        public string MoldKind { get; set; }

        public string BuyerModelID { get; set; }
        public string BuyerModel { get; set; }
        public string CustomID { get; set; }
        public string KCustom { get; set; }

        public double MoldSizeX { get; set; }
        public double MoldSizeY { get; set; }
        public double MoldSizeH { get; set; }
        public string MoldQuality { get; set; }
        public double Weight { get; set; }

        public string DisCardYN { get; set; }
        public string Cavity { get; set; }
        public string RealCavity { get; set; }
        public string Storage { get; set; }
        public string StorageName { get; set; }

        public string ProdCustomName { get; set; }
        public string OwnerCustomName { get; set; }
        public string OwnerOneTimePayYn { get; set; }
        public string OwnerOneTimePayYnName { get; set; }

        public string SetDate { get; set; }
        public string ProdOrderDate { get; set; }
        public string ProdDueDate { get; set; }
        public string ProdCompDate { get; set; }

        public string MainUseYN { get; set; }
        public string Comments { get; set; }
        public string MoldPerson { get; set; }

        public double SetCheckProdQty { get; set; }
        public double AfterRepairHitcount { get; set; }
        public double SetWashingProdQty { get; set; }
        public double AfterWashHitcount { get; set; }

        public double SetProdQty { get; set; }
        public double HitCount { get; set; }

        public double SetHitCount { get; set; }
        public string SetHitCountDate { get; set; }

        public string EvalGrade { get; set; }
        public double EvalScore { get; set; }

        public string AttFile1 { get; set; }
        public string AttPath1 { get; set; }
        public string AttFile2 { get; set; }
        public string AttPath2 { get; set; }
        public string AttFile3 { get; set; }
        public string AttPath3 { get; set; }

        public string Article { get; set; }
        public string BuyerArticleNo { get; set; }

    }

    
    class Win_dvl_Molding_U_Parts_CodeView : BaseView
    {
        public int Num { get; set; }

        public string MoldID { get; set; }
        public string McPartID { get; set; }
        public string MCPartName { get; set; }
        public string ChangeCheckGbn { get; set; }
        public string CycleProdQty { get; set; }

        public string StartSetProdQty { get; set; }
        public string StartSetDate { get; set; }
    }

    class MoldArticle_CodeView
    {
        public int Num { get; set; }
        public string ArticleID { get; set; }
        public string BuyerArticleNo { get; set; }
        public string BuyerArticleNo2 { get; set; }
        public string Article { get; set; }
        public string Article2 { get; set; }
    }
}
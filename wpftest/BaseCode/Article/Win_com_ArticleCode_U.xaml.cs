using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media.Imaging;
using WizMes_WellMade.PopUp;
using WizMes_WellMade.PopUP;

namespace WizMes_WellMade
{
    /// <summary>
    /// Win_com_Article_U.xaml에 대한 상호 작용 논리
    /// </summary>
    public partial class Win_com_ArticleCode_U : UserControl
    {
        string stDate = string.Empty;
        string stTime = string.Empty;
        //테스트
        Lib lib = new Lib();
        PlusFinder pf = new PlusFinder();

        string strFlag = string.Empty;
        int rowNum = 0;
        Win_com_Article_U_CodeView winArticleCode = new Win_com_Article_U_CodeView();
        Process_CodeView ProcessCodeView = new Process_CodeView();


        ObservableCollection<Win_com_Article_U_CodeView> ovcArticleCode = new ObservableCollection<Win_com_Article_U_CodeView>();
        ObservableCollection<Process_CodeView> ovcArticleProcess = new ObservableCollection<Process_CodeView>();
        ObservableCollection<CustomArticle_CodeView> ovcCustomArticle = new ObservableCollection<CustomArticle_CodeView>();

        /// <summary>
        /// 서브 그리드 삭제용
        /// </summary>
        ObservableCollection<Win_com_Article_U_CodeView> ovcInspectAutoBasisSub_Delete
            = new ObservableCollection<Win_com_Article_U_CodeView>();



        // FTP 활용모음.
        string strImagePath = string.Empty;
        string strFullPath = string.Empty;

        List<string[]> listFtpFile = new List<string[]>();
        List<string[]> deleteListFtpFile = new List<string[]>(); // 삭제할 파일 리스트
        private FTP_EX _ftp = null;

        //string FTP_ADDRESS = "ftp://wizis.iptime.org/ImageData";
        //string FTP_ADDRESS = "ftp://wizis.iptime.org/ImageData/Draw";
        string FTP_ADDRESS = "ftp://" + LoadINI.FileSvr + ":" + LoadINI.FTPPort + LoadINI.FtpImagePath + "/Article";
        private const string FTP_ID = "wizuser";
        private const string FTP_PASS = "wiz9999";
        private const string LOCAL_DOWN_PATH = "C:\\Temp";


        //string FTP_ADDRESS = "ftp://192.168.0.4/Article";
        //string FTP_ADDRESS = "ftp://192.168.0.120";

        public Win_com_ArticleCode_U()
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

            GetProcess();
            ImageOnlySeeMode();

        }

        #region SetComboBox - 콤보박스 세팅

        //
        private void SetComboBox()
        {
            //품명그룹(조회, 입력)
            ObservableCollection<CodeView> ovcArticleGrp = ComboBoxUtil.Instance.Gf_DB_MT_sArticleGrp();
            this.cboArticleGrp.ItemsSource = ovcArticleGrp;
            this.cboArticleGrp.DisplayMemberPath = "code_name";
            this.cboArticleGrp.SelectedValuePath = "code_id";

            this.cboArticleGrpSearch.ItemsSource = ovcArticleGrp;
            this.cboArticleGrpSearch.DisplayMemberPath = "code_name";
            this.cboArticleGrpSearch.SelectedValuePath = "code_id";
            this.cboArticleGrpSearch.SelectedIndex = 0;


            //자재단위(입력)
            ObservableCollection<CodeView> ovcUnitClss = ComboBoxUtil.Instance.Gf_DB_CM_GetComCodeDataset(null, "MTRUNIT", "Y", "");
            this.cboUnitClss.ItemsSource = ovcUnitClss;
            this.cboUnitClss.DisplayMemberPath = "code_name";
            this.cboUnitClss.SelectedValuePath = "code_id";


            //날끝타입(입력)
            //ObservableCollection<CodeView> ovcBladeEndClss = ComboBoxUtil.Instance.Gf_DB_CM_GetComCodeDataset(null, "BLADETYPE", "Y", "");
            //this.cboBladeEndType.ItemsSource = ovcBladeEndClss;
            //this.cboBladeEndType.DisplayMemberPath = "code_name";
            //this.cboBladeEndType.SelectedValuePath = "code_id";

            //공급유형(조회, 입력)
            ObservableCollection<CodeView> ovcSupplyType = ComboBoxUtil.Instance.Gf_DB_CM_GetComCodeDataset(null, "CMMASPLTYPE", "Y", "");
            this.cboSupplyType.ItemsSource = ovcSupplyType;
            this.cboSupplyType.DisplayMemberPath = "code_name";
            this.cboSupplyType.SelectedValuePath = "code_id";

            ObservableCollection<CodeView> ovcSupplyTypeAll = ComboBoxUtil.Instance.Gf_DB_CM_GetComCodeDatasetPlusAll(null, "CMMASPLTYPE", "Y", "");
            this.cboSrhSupplyType.ItemsSource = ovcSupplyTypeAll;
            this.cboSrhSupplyType.DisplayMemberPath = "code_name";
            this.cboSrhSupplyType.SelectedValuePath = "code_id";
            this.cboSrhSupplyType.SelectedIndex = 0;

            //라벨발행품 여부(입력)
            List<string[]> listYN = new List<string[]>();
            string[] YN01 = new string[] { "Y", "Y" };
            string[] YN02 = new string[] { "N", "N" };
            listYN.Add(YN01);
            listYN.Add(YN02);

            ObservableCollection<CodeView> ovcYN = ComboBoxUtil.Instance.Direct_SetComboBox(listYN);
            this.cboLabelPrintYN.ItemsSource = ovcYN;
            this.cboLabelPrintYN.DisplayMemberPath = "code_name";
            this.cboLabelPrintYN.SelectedValuePath = "code_id";
            //cboBigMiSmal
            // 용도 구하기
            //ObservableCollection<CodeView> ovcUseingType = ComboBoxUtil.Instance.Gf_DB_CM_GetComCodeDataset(null, "CMMAUSETYPE", "Y", "");
            //this.cboUseingType.ItemsSource = ovcUseingType;
            //this.cboUseingType.DisplayMemberPath = "code_name";
            //this.cboUseingType.SelectedValuePath = "code_id";

            //FTA 중점관리품 여부(입력)
            this.cboFTAMgrYN.ItemsSource = ovcYN;
            this.cboFTAMgrYN.DisplayMemberPath = "code_name";
            this.cboFTAMgrYN.SelectedValuePath = "code_id";

            this.cboInsPectYN.ItemsSource = ovcYN;
            this.cboInsPectYN.DisplayMemberPath = "code_name";
            this.cboInsPectYN.SelectedValuePath = "code_id";


            //원재료 속성(입력)
            ObservableCollection<CodeView> ovcPART_ATTR = ComboBoxUtil.Instance.Gf_DB_CM_GetComCodeDataset(null, "PART_ATTR", "Y", "");
            this.cboPART_ATTR.ItemsSource = ovcPART_ATTR;
            this.cboPART_ATTR.DisplayMemberPath = "code_name";
            this.cboPART_ATTR.SelectedValuePath = "code_id";

            //매입,매출 화폐단위(입력)
            List<string[]> listPrice = new List<string[]>();
            string[] Price01 = new string[] { "0", "₩" };
            string[] Price02 = new string[] { "1", "$" };
            listPrice.Add(Price01);
            listPrice.Add(Price02);

            ObservableCollection<CodeView> ovcPrice = ComboBoxUtil.Instance.Direct_SetComboBox(listPrice);
            this.cboPriceClss.ItemsSource = ovcPrice;
            this.cboPriceClss.DisplayMemberPath = "code_name";
            this.cboPriceClss.SelectedValuePath = "code_id";

            this.cboPriceClss2.ItemsSource = ovcPrice;
            this.cboPriceClss2.DisplayMemberPath = "code_name";
            this.cboPriceClss2.SelectedValuePath = "code_id";

            ObservableCollection<CodeView> ovcPattern = ComboBoxUtil.Instance.GetProcessPattern("");
            this.cboPattern.ItemsSource = ovcPattern;
            this.cboPattern.DisplayMemberPath = "code_name";
            this.cboPattern.SelectedValuePath = "code_id";

            // 제품군
            ObservableCollection<CodeView> ovcProductGrp = ComboBoxUtil.Instance.Gf_DB_CM_GetComCodeDataset(null, "CMPRDGRPID", "Y", "");
            this.cboProductGrpID.ItemsSource = ovcProductGrp;
            this.cboProductGrpID.DisplayMemberPath = "code_name";
            this.cboProductGrpID.SelectedValuePath = "code_id";

            // 제품군
            ObservableCollection<CodeView> ovcBigMiSmal = ComboBoxUtil.Instance.Gf_DB_CM_GetComCodeDataset(null, "BIGMISMAL", "Y", "");
            this.cboBigMiSmal.ItemsSource = ovcBigMiSmal;
            this.cboBigMiSmal.DisplayMemberPath = "code_name";
            this.cboBigMiSmal.SelectedValuePath = "code_id";
            // 부품분류
            //List<string[]> lstPart = new List<string[]>();
            //string[] Part01 = new string[] { "0", "완제품" };
            //string[] Part02 = new string[] { "1", "재단품" };
            //lstPart.Add(Part01);
            //lstPart.Add(Part02);

            //ObservableCollection<CodeView> ovcPartGBN = ComboBoxUtil.Instance.Direct_SetComboBox(lstPart);
            //this.cboPartGBNID.ItemsSource = ovcPartGBN;
            //this.cboPartGBNID.DisplayMemberPath = "code_name";
            //this.cboPartGBNID.SelectedValuePath = "code_id";

            ObservableCollection<CodeView> ovcPartGBN = ComboBoxUtil.Instance.Gf_DB_CM_GetComCodeDataset(null, "PARTGBNID", "Y", "");
            this.cboPartGBNID.ItemsSource = ovcPartGBN;
            this.cboPartGBNID.DisplayMemberPath = "code_name";
            this.cboPartGBNID.SelectedValuePath = "code_id";
        }

        #endregion // SetComboBox - 콤보박스 세팅

        #region Header 부분 - 검색조건

        // 품명
        private void lblBuyerArticleNoSrh_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            if (chkBuyerArticleNoSrh.IsChecked == true)
            {
                chkBuyerArticleNoSrh.IsChecked = false;
            }
            else
            {
                chkBuyerArticleNoSrh.IsChecked = true;
            }
        }
        private void chkBuyerArticleNoSrh_Checked(object sender, RoutedEventArgs e)
        {
            chkBuyerArticleNoSrh.IsChecked = true;
            txtBuyerArticleNoSrh.IsEnabled = true;
        }

        private void chkBuyerArticleNoSrh_Unchecked(object sender, RoutedEventArgs e)
        {
            chkBuyerArticleNoSrh.IsChecked = false;
            txtBuyerArticleNoSrh.IsEnabled = false;
        }

        // 공급유형
        private void lblSupplyTypeSrh_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            if (chkSupplyTypeSrh.IsChecked == true)
            {
                chkSupplyTypeSrh.IsChecked = false;
            }
            else
            {
                chkSupplyTypeSrh.IsChecked = true;
            }
        }
        private void chkSupplyTypeSrh_Checked(object sender, RoutedEventArgs e)
        {
            chkSupplyTypeSrh.IsChecked = true;
            cboSrhSupplyType.IsEnabled = true;
        }
        private void chkSupplyTypeSrh_Unchecked(object sender, RoutedEventArgs e)
        {
            chkSupplyTypeSrh.IsChecked = false;
            cboSrhSupplyType.IsEnabled = false;
        }

        //품명그룹
        private void lblArticleGrpSrh_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            if (chkArticleGrpSrh.IsChecked == true) { chkArticleGrpSrh.IsChecked = false; }
            else { chkArticleGrpSrh.IsChecked = true; }
        }

        //품명그룹
        private void chkArticleGrpSrh_Checked(object sender, RoutedEventArgs e)
        {
            chkArticleGrpSrh.IsChecked = true;
            cboArticleGrpSearch.IsEnabled = true;
        }

        //품명그룹
        private void chkArticleGrpSrh_Unchecked(object sender, RoutedEventArgs e)
        {
            chkArticleGrpSrh.IsChecked = false;
            cboArticleGrpSearch.IsEnabled = false;
        }

        // 사용안함 포함 라벨
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
        // 사용안함 포함 체크박스
        private void chkNoUse_Checked(object sender, RoutedEventArgs e)
        {
            chkNoUse.IsChecked = true;
        }
        // 사용안함 포함 체크박스
        private void chkNoUse_UnChecked(object sender, RoutedEventArgs e)
        {
            chkNoUse.IsChecked = false;
        }

        #endregion // Header 부분 - 검색조건

        #region 오른쪽 서브 그리드 - 공정 프로시저 : 화면 활성화시 실행됨

        /// <summary>
        /// 공정 DataGrid채우기
        /// </summary>
        private void GetProcess()
        {
            if (dgdProcess.ItemsSource != null)
            {
                dgdProcess.ItemsSource = null;
                ovcArticleProcess.Clear();
            }

            try
            {
                Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                DataSet ds = DataStore.Instance.ProcedureToDataSet("xp_Code_sProcess", sqlParameter, false);

                if (ds != null && ds.Tables.Count > 0)
                {
                    DataTable dt = ds.Tables[0];

                    if (dt.Rows.Count > 0)
                    {
                        DataRowCollection drc = dt.Rows;
                        int i = 0;

                        foreach (DataRow item in drc)
                        {
                            var varProcess = new Process_CodeView()
                            {
                                Num = i + 1,
                                ProcessID = item["ProcessID"].ToString(),
                                Process = item["Process"].ToString(),
                                CheckFlag = false
                            };
                            ovcArticleProcess.Add(varProcess);
                            i++;
                        }
                        dgdProcess.ItemsSource = ovcArticleProcess;
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

        #endregion // 오른쪽 서브 그리드 - 공정 프로시저 : 화면 활성화시 실행됨

        /// <summary>
        /// 수정,추가 저장 후
        /// </summary>
        private void CanBtnControl()
        {
            Lib.Instance.UiButtonEnableChange_IUControl(this);
            dgdArticleCode.IsHitTestVisible = true;
            grdInput1.IsHitTestVisible = false;
            grdInput2.IsHitTestVisible = false;
            //grdMtrInfo.IsEnabled = false;
            //grdFTP.IsEnabled = false;
            cboBigMiSmal.IsHitTestVisible = false;
            ImageOnlySeeMode();


            btnImgSeeCheckAndSetting();
        }

        /// <summary>
        /// 수정,추가 진행 중
        /// </summary>
        private void CantBtnControl()
        {
            Lib.Instance.UiButtonEnableChange_SCControl(this);
            dgdArticleCode.IsHitTestVisible = false;
            grdInput1.IsHitTestVisible = true;
            grdInput2.IsHitTestVisible = true;
            //grdMtrInfo.IsEnabled = true;
            //grdFTP.IsEnabled = true;azzz

            chkBuySaleMainYN.IsChecked = false;
            cboBigMiSmal.IsHitTestVisible = true;
            ImageSaveUpdateMode();

            btnImgSeeCheckAndSetting();
        }


        /// <summary>
        /// 추가,수정 시 동작 모음
        /// </summary>
        //private void ControlVisibleAndEnable_AU()
        //{
        //    Lib.Instance.UiButtonEnableChange_SCControl(this);
        //    dgdArticleCode.IsHitTestVisible = false;
        //    grdInput1.IsHitTestVisible = true;
        //}z

        /// <summary>
        /// 저장 취소시 동작모음
        /// </summary>
        //private void ControlVisibleAndEnable_SC()
        //{
        //    Lib.Instance.UiButtonEnableChange_IUControl(this);
        //    dgdArticleCode.IsHitTestVisible = true;
        //    grdInput1.IsHitTestVisible = false;
        //}

        #region 사진 (첨부파일등록) 보기버튼만 활성화 → 추가, 수정시 나머지 버튼 활성화

        private void ImageOnlySeeMode()
        {
            btnFileUpload1.IsEnabled = false;
            txtSketch1.IsEnabled = false;
            btnFileDelete1.IsEnabled = false;
            //btnFileSee1.IsEnabled = true;

            btnFileUpload2.IsEnabled = false;
            txtSketch2.IsEnabled = false;
            btnFileDelete2.IsEnabled = false;
            //btnFileSee2.IsEnabled = true;

            btnFileUpload3.IsEnabled = false;
            txtSketch3.IsEnabled = false;
            btnFileDelete3.IsEnabled = false;
            //btnFileSee3.IsEnabled = true;

            btnFileUpload4.IsEnabled = false;
            txtSketch4.IsEnabled = false;
            btnFileDelete4.IsEnabled = false;
            //btnFileSee4.IsEnabled = true;

            btnFileUpload5.IsEnabled = false;
            txtSketch5.IsEnabled = false;
            btnFileDelete5.IsEnabled = false;
            //btnFileSee5.IsEnabled = true;

            btnFileUpload6.IsEnabled = false;
            txtSketch6.IsEnabled = false;
            btnFileDelete6.IsEnabled = false;
            //btnFileSee6.IsEnabled = true;


            //btnFileUpload7.IsEnabled = false;
            //txtSketch7.IsEnabled = false;
            //btnFileDelete7.IsEnabled = false;

            // 보기 버튼체크
            btnImgSeeCheckAndSetting();
        }

        private void ImageSaveUpdateMode()
        {
            btnFileUpload1.IsEnabled = true;
            txtSketch1.IsEnabled = true;
            btnFileDelete1.IsEnabled = true;
            //btnFileSee1.IsEnabled = true;

            btnFileUpload2.IsEnabled = true;
            txtSketch2.IsEnabled = true;
            btnFileDelete2.IsEnabled = true;
            //btnFileSee2.IsEnabled = true;

            btnFileUpload3.IsEnabled = true;
            txtSketch3.IsEnabled = true;
            btnFileDelete3.IsEnabled = true;
            //btnFileSee3.IsEnabled = true;

            btnFileUpload4.IsEnabled = true;
            txtSketch4.IsEnabled = true;
            btnFileDelete4.IsEnabled = true;
            //btnFileSee4.IsEnabled = true;

            btnFileUpload5.IsEnabled = true;
            txtSketch5.IsEnabled = true;
            btnFileDelete5.IsEnabled = true;
            //btnFileSee5.IsEnabled = true;

            btnFileUpload6.IsEnabled = true;
            txtSketch6.IsEnabled = true;
            btnFileDelete6.IsEnabled = true;
            //btnFileSee6.IsEnabled = true;


            //btnFileUpload7.IsEnabled = true;
            //txtSketch7.IsEnabled = true;
            //btnFileDelete7.IsEnabled = true;

            // 보기 버튼체크
            btnImgSeeCheckAndSetting();
        }

        #endregion // 사진 (첨부파일등록) 보기버튼만 활성화 → 추가, 수정시 나머지 버튼 활성화

        #region Header 부분 - 상단 오른쪽 버튼

        //추가
        private void btnAdd_Click(object sender, RoutedEventArgs e)
        {

            rowNum = 0;

            CantBtnControl();
            tbkMsg.Text = "자료 입력 중";
            strFlag = "I";
            this.DataContext = null;

            GetProcess();

            // 기본 세팅

            cboArticleGrp.SelectedIndex = 0; //품명그룹
            cboSupplyType.SelectedIndex = 0; //공급유형
            cboUnitClss.SelectedIndex = 0; //단위
            cboLabelPrintYN.SelectedIndex = 0; //라벨관리
            cboInsPectYN.SelectedIndex = 0; //무검사품YN 
            cboProductGrpID.SelectedIndex = 0; //제품군
            cboPartGBNID.SelectedIndex = 0; //부품분류
            cboFTAMgrYN.SelectedIndex = 1; //FTA중점
            cboBigMiSmal.SelectedIndex = 3; //대중소 구분

            txtCode.IsReadOnly = false;
            txtBuyerArticleNo.Focus();
        }

        //수정
        private void btnUpdate_Click(object sender, RoutedEventArgs e)
        {
            winArticleCode = dgdArticleCode.SelectedItem as Win_com_Article_U_CodeView;

            if (winArticleCode != null)
            {
                rowNum = dgdArticleCode.SelectedIndex;
                CantBtnControl();
                tbkMsg.Text = "자료 수정 중";
                strFlag = "U";

                //GetProcess();
                txtCode.IsReadOnly = true;
                txtName.Focus();
            }
            else
            {
                MessageBox.Show("수정할 자료를 선택하고 눌러주십시오.");
            }


        }

        //삭제
        private void btnDelete_Click(object sender, RoutedEventArgs e)
        {
            winArticleCode = dgdArticleCode.SelectedItem as Win_com_Article_U_CodeView;

            if (winArticleCode == null)
            {
                MessageBox.Show("삭제할 데이터가 지정되지 않았습니다. 삭제데이터를 지정하고 눌러주세요");
            }
            else
            {
                if (MessageBox.Show("선택하신 항목을 삭제하시겠습니까?", "삭제 전 확인", MessageBoxButton.YesNo) == MessageBoxResult.Yes)
                {
                    DataStore.Instance.InsertLogByForm(this.GetType().Name, "D");

                    //삭제 전 체크
                    if (!DeleteDataCheck(winArticleCode.ArticleID))
                        return;

                    if (dgdArticleCode.Items.Count > 0 && dgdArticleCode.SelectedItem != null)
                    {
                        //rowNum = dgdArticleCode.SelectedIndex;
                        rowNum = 0;
                    }


                    // 품명 삭제 전 체크하기



                    FTP_RemoveDir(winArticleCode.ArticleID);

                    if (Procedure.Instance.DeleteData(winArticleCode.ArticleID, MainWindow.CurrentUser,
                        "sArticleID", "sUpdateUserID", "xp_Article_dArticle"))
                    {
                        this.DataContext = null;
                        //rowNum -= 1;
                        re_Search(rowNum);
                    }




                }
            }
        }


        //private void beDelete()
        //{
        //    var ArticleCD = dgdArticleCode.SelectedItem as Win_com_Article_U_CodeView;

        //    if (ArticleCD != null)
        //    {
        //        if (DeleteData(ArticleCD.ArticleID))
        //        {
        //            rowNum = 0;
        //            re_Search();
        //        }
        //    }

        //}
        //닫기
        private void btnClose_Click(object sender, RoutedEventArgs e)
        {
            DataStore.Instance.InsertLogByFormS(this.GetType().Name, stDate, stTime, "E");
            Lib.Instance.ChildMenuClose(this.ToString());
        }

        //조회
        private void btnSearch_Click(object sender, RoutedEventArgs e)
        {
            using (Loading ld = new Loading(beSearch))
            {
                ld.ShowDialog();
            }
        }

        private void beSearch()
        {
            //검색버튼 비활성화
            btnSearch.IsEnabled = false;

            Dispatcher.BeginInvoke(new Action(() =>
            {
                //로직
                rowNum = 0;
                re_Search(rowNum);

            }), System.Windows.Threading.DispatcherPriority.Background);

            btnSearch.IsEnabled = true;
        }

        //저장
        private void btnSave_Click(object sender, RoutedEventArgs e)
        {
            winArticleCode = dgdArticleCode.SelectedItem as Win_com_Article_U_CodeView;


            if (SaveData(txtCode.Text, strFlag))
            {
                CanBtnControl();

                strFlag = string.Empty;
                re_Search(rowNum);
            }
        }

        //취소
        private void btnCancel_Click(object sender, RoutedEventArgs e)
        {
            //ControlVisibleAndEnable_SC();

            ovcInspectAutoBasisSub_Delete.Clear();

            CanBtnControl();
            strFlag = string.Empty;
            re_Search(rowNum);


        }

        //엑셀
        private void btnExcel_Click(object sender, RoutedEventArgs e)
        {
            DataTable dt = null;
            string Name = string.Empty;

            string[] dgdStr = new string[2];
            dgdStr[0] = "품명코드";
            dgdStr[1] = dgdArticleCode.Name;

            ExportExcelxaml ExpExc = new ExportExcelxaml(dgdStr);
            ExpExc.ShowDialog();

            if (ExpExc.DialogResult.HasValue)
            {
                if (ExpExc.choice.Equals(dgdArticleCode.Name))
                {
                    DataStore.Instance.InsertLogByForm(this.GetType().Name, "E");
                    if (ExpExc.Check.Equals("Y"))
                        dt = Lib.Instance.DataGridToDTinHidden(dgdArticleCode);
                    else
                        dt = Lib.Instance.DataGirdToDataTable(dgdArticleCode);

                    Name = dgdArticleCode.Name;
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

        #endregion // Header 부분 - 상단 오른쪽 버튼


        #region Content 부분

        // 메인 그리드 더블클릭 → 수정 버튼 클릭 이벤트
        private void dgdArticleCode_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            btnUpdate_Click(null, null);
        }


        // 메인 그리드 선택 이벤트
        private void dgdArticleCode_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            // 일단 이미지 비우자
            imgSetting.Source = null;

            if (dgdArticleCode.SelectedItem != null)
            {
                winArticleCode = dgdArticleCode.SelectedItem as Win_com_Article_U_CodeView;

                if (winArticleCode != null)
                {
                    this.DataContext = winArticleCode;
                    FillGridProcess(winArticleCode.ArticleID);
                    FillGridCustomArticle(winArticleCode.ArticleID);
                    // 사용안함
                    if (winArticleCode.UseClss != null && winArticleCode.UseClss.Trim().Equals("*"))
                    {
                        chkNotUse.IsChecked = true;
                    }
                    else
                    {
                        chkNotUse.IsChecked = false;
                    }

                    // 화폐단위
                    if (winArticleCode.UnitPriceClss != null)
                    {
                        cboPriceClss.SelectedValue = winArticleCode.UnitPriceClss;
                        cboPriceClss2.SelectedValue = winArticleCode.UnitPriceClss;
                    }

                    // 주요 관심품목
                    if (winArticleCode.BuySaleMainYN != null && winArticleCode.BuySaleMainYN.Trim().Equals("Y"))
                    {
                        chkBuySaleMainYN.IsChecked = true;
                    }
                    else
                    {
                        chkBuySaleMainYN.IsChecked = false;
                    }
                }

                // 보기 버튼체크
                btnImgSeeCheckAndSetting();





                // 일단 FTP 막아놓음
                //if (!winArticleCode.Sketch1File.Equals(string.Empty))
                //{
                //    imgSetting.Source = SetImage("/" + winArticleCode.ArticleID + "/" + winArticleCode.Sketch1File);
                //}
            }
        }

        // Content 주요 관심품목 체크박스 이벤트
        private void lblBuySaleMainYN_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            if (chkBuySaleMainYN.IsChecked == true)
            {
                chkBuySaleMainYN.IsChecked = false;
            }
            else
            {
                chkBuySaleMainYN.IsChecked = true;
            }
        }
        // Content 사용안함 체크박스 이벤트
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

        // 서브 그리드 (공정) 선택 이벤트
        private void dgdtpechkChoice_Click(object sender, RoutedEventArgs e)
        {
            if (lblMsg.Visibility == Visibility.Visible)
            {
                ProcessCodeView = dgdProcess.SelectedItem as Process_CodeView;

                if (ProcessCodeView != null)
                {
                    CheckBox chk = sender as CheckBox;

                    if (chk != null)
                    {
                        if (chk.IsChecked == false)
                        {
                            ProcessCodeView.CheckFlag = false;
                        }
                        else
                        {
                            ProcessCodeView.CheckFlag = true;
                        }
                    }
                }
            }
        }

        // 적정 재고량 계산하기
        private void btnCalNeedStockQty_Click(object sender, RoutedEventArgs e)
        {
            Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
            sqlParameter.Clear();

            sqlParameter.Add("sArticleID", txtCode.Text);
            sqlParameter.Add("NeedStockQty", 0);

            Dictionary<string, int> outputParam = new Dictionary<string, int>();
            outputParam.Add("NeedStockQty", 15);
            Dictionary<string, string> dicResult = DataStore.Instance.ExecuteProcedureOutputNoTran("xp_Article_calArticleNeedStockQty", sqlParameter, outputParam, true);

            string result = stringFormatN0(dicResult["NeedStockQty"]);

            if ((result != string.Empty) && (result != "9999"))
            {
                //MessageBox.Show("성공");
                //flag = true;

                txtNeedStockQty.Text = result;
            }
            else
            {
                MessageBox.Show("[프로시저 오류]\r\n" + result);
                //flag = false;
                //return false;
            }
        }

        // 품명 그룹 콤보박스 선택 이벤트
        private void cboArticleGrp_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (lblMsg.Visibility == Visibility.Visible)
            {
                if (cboArticleGrp.SelectedValue == null)
                {
                    return;
                }
                //원자재 선택시 KG, 품명그룹 01 =원자재
                if (cboArticleGrp.SelectedValue.Equals("01"))
                {
                    cboUnitClss.SelectedValue = "2";
                }
                else    //나머지는 EA 
                {
                    cboUnitClss.SelectedValue = "0";
                }

                //cboProductGrpID.IsDropDownOpen = true;
            }
        }

        // 화폐단위 동기화
        private void cboPriceClss_DropDownClosed(object sender, EventArgs e)
        {
            if (cboPriceClss.SelectedValue != null)
            {
                cboPriceClss2.SelectedValue = cboPriceClss.SelectedValue;
            }

            //cboUnitClss.IsDropDownOpen = true;
        }
        private void cboPriceClss2_DropDownClosed(object sender, EventArgs e)
        {
            if (cboPriceClss2.SelectedValue != null)
            {
                cboPriceClss.SelectedValue = cboPriceClss2.SelectedValue;
            }
        }

        // HSCode 플러스파인더 이벤트
        private void txtHSCode_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                MainWindow.pf.ReturnCode(txtHSCode, (int)Defind_CodeFind.DCF_HSCODE, "");

                txtBuyerArticleNo.Focus();
            }
        }
        private void btnHSCodePf_Click(object sender, RoutedEventArgs e)
        {
            MainWindow.pf.ReturnCode(txtHSCode, (int)Defind_CodeFind.DCF_HSCODE, "");

            txtBuyerArticleNo.Focus();
        }

        #endregion // Content 부분

        #region 주요 메서드

        //재검색
        private void re_Search(int selectedIndex)
        {
            FillGrid();

            if (dgdArticleCode.Items.Count > 0)
            {
                dgdArticleCode.SelectedIndex = selectedIndex;
            }
        }


        //특수문자 포함 검색
        private string Escape(string str)
        {
            string result = "";

            for (int i = 0; i < str.Length; i++)
            {
                string txt = str.Substring(i, 1);

                bool isSpecial = Regex.IsMatch(txt, @"[^a-zA-Z0-9가-힣]");

                if (isSpecial == true)
                {
                    result += (@"/" + txt);
                }
                else
                {
                    result += txt;
                }
            }
            return result;
        }

        #region 조회

        private void FillGrid()
        {

            if (dgdArticleCode.Items.Count > 0)
            {
                dgdArticleCode.Items.Clear();
            }

            try
            {
                Dictionary<string, object> sqlParameter = new Dictionary<string, object>();


                sqlParameter.Add("nArticle", chkArticleSrh.IsChecked == true ? 1 : 0);
                sqlParameter.Add("sArticle", chkArticleSrh.IsChecked == true && !txtArticleSrh.Text.Trim().Equals("") ? txtArticleSrh.Text : "");  //@escape함수제거


                sqlParameter.Add("iIncNotUse", chkNoUse.IsChecked == true ? 1 : 0);
                sqlParameter.Add("sArticleGrpID", chkArticleGrpSrh.IsChecked == true && cboArticleGrpSearch.SelectedValue != null ? cboArticleGrpSearch.SelectedValue.ToString() : "");
                sqlParameter.Add("sSupplyType", chkSupplyTypeSrh.IsChecked == true && cboSrhSupplyType.SelectedValue != null ? cboSrhSupplyType.SelectedValue.ToString() : "");
                sqlParameter.Add("nBuyerArticleNo", chkBuyerArticleNoSrh.IsChecked == true ? 1 : 0);
                sqlParameter.Add("BuyerArticleNo", chkBuyerArticleNoSrh.IsChecked == true && !txtBuyerArticleNoSrh.Text.Trim().Equals("") ? txtBuyerArticleNoSrh.Text : "");  //@escape함수제거

                DataSet ds = DataStore.Instance.ProcedureToDataSet_LogWrite("xp_Article_sArticle", sqlParameter, true, "R");

                if (ds != null && ds.Tables.Count > 0)
                {
                    DataTable dt = ds.Tables[0];

                    if (dt.Rows.Count == 0)
                    {
                        MessageBox.Show("조회된 데이터가 없습니다.");
                    }
                    else
                    {
                        int i = 0;
                        DataRowCollection drc = dt.Rows;

                        foreach (DataRow dr in drc)
                        {
                            i++;

                            var dgdModelInfo = new Win_com_Article_U_CodeView()
                            {
                                Num = i + "",

                                ArticleID = dr["ArticleID"].ToString(),
                                CompanyID = dr["CompanyID"].ToString(),
                                Article = dr["Article"].ToString(),
                                BuyerArticleNo = dr["BuyerArticleNo"].ToString(),
                                ArticleGrpID = dr["ArticleGrpID"].ToString(),

                                SupplyType = dr["SupplyType"].ToString(),
                                SupplyTypeName = dr["SupplyTypeName"].ToString(),
                                ProductGrpID = dr["ProductGrpID"].ToString(),
                                ProductGrpName = dr["ProductGrpName"].ToString(),
                                PartGBNID = dr["PartGBNID"].ToString(),

                                PartGBNName = dr["PartGBNName"].ToString(),
                                UseingType = dr["UseingType"].ToString(),
                                UseingTypeName = dr["UseingTypeName"].ToString(),
                                Weight = stringFormatN2(dr["Weight"]), // 주석
                                UseClss = dr["UseClss"].ToString(),

                                Spec = dr["Spec"].ToString(),
                                NeedStockQty = stringFormatN0(dr["NeedStockQty"]),
                                UnitClss = dr["UnitClss"].ToString(),
                                UnitClssName = dr["UnitClssName"].ToString(),
                                ProdQtyPerBox = stringFormatN0(dr["ProdQtyPerBox"]),

                                OutQtyPerBox = stringFormatN0(dr["OutQtyPerBox"]),
                                ImageName = dr["ImageFileName"].ToString(),

                                Defect1Path = dr["Defect1Path"].ToString(),
                                Defect1File = dr["Defect1File"].ToString(),
                                Defect2Path = dr["Defect2Path"].ToString(),
                                Defect2File = dr["Defect2File"].ToString(),
                                Defect3Path = dr["Defect3Path"].ToString(),
                                Defect3File = dr["Defect3File"].ToString(),

                                Sketch1Path = dr["Sketch1Path"].ToString(),
                                Sketch1File = dr["Sketch1File"].ToString(),
                                Sketch2Path = dr["Sketch2Path"].ToString(),
                                Sketch2File = dr["Sketch2File"].ToString(),
                                Sketch3Path = dr["Sketch3Path"].ToString(),
                                Sketch3File = dr["Sketch3File"].ToString(),
                                Sketch4Path = dr["Sketch4Path"].ToString(),
                                Sketch4File = dr["Sketch4File"].ToString(),
                                Sketch5Path = dr["Sketch5Path"].ToString(),
                                Sketch5File = dr["Sketch5File"].ToString(),
                                Sketch6Path = dr["Sketch6Path"].ToString(),
                                Sketch6File = dr["Sketch6File"].ToString(),
                                //Sketch7Path = dr["Sketch7Path"].ToString(),
                                //Sketch7File = dr["Sketch7File"].ToString(),


                                LabelPrintYN = dr["LabelPrintYN"].ToString(),
                                UnitPrice = stringFormatN2(dr["UnitPrice"]),
                                OutUnitPrice = stringFormatN2(dr["OutUnitPrice"]),
                                FTAMgrYN = dr["FTAMgrYN"].ToString(),
                                HSCODE = dr["HSCODE"].ToString(),

                                CoatingSpec = dr["CoatingSpec"].ToString(),
                                BuySaleMainYN = dr["BuySaleMainYN"].ToString(),
                                Comments = dr["Comments"].ToString(),
                                PART_ATTR = dr["PART_ATTR"].ToString(),
                                PatternID = dr["PatternID"].ToString(),



                                UnitPriceClss = dr["UnitPriceClss"].ToString(),

                                //FreeStuffinYN = dr["FreeStuffinYN"].ToString(), // 무검사입고품여부

                                BigMiSmalGbn = dr["BigMiSmalGbn"].ToString(), // 

                                Exdiameter = stringFormatN2(dr["Exdiameter"]),
                                InDiameter = stringFormatN2(dr["InDiameter"]),
                                Length = stringFormatN2(dr["Length"]),


                                ProdDiffiLevel = stringFormatN2(dr["ProdDiffiLevel"]),
                                BuyerModelID = dr["BuyerModelID"].ToString(),
                                BuyerModel = dr["BuyerModel"].ToString(),
                                InspectYN = dr["InspectYN"].ToString(),
                                ArticleGrpName = dr["ArticleGrpName"].ToString(),

                            };

                            // 사용 여부 UseClssName 으로 바인딩함
                            if (dgdModelInfo.UseClss.Trim().Equals(""))
                            {
                                dgdModelInfo.UseClssName = "Y";
                            }
                            else
                            {
                                dgdModelInfo.UseClssName = "N";
                            }

                            dgdModelInfo.ProdDiffiLevel = lib.returnNumStringOne(dgdModelInfo.ProdDiffiLevel);
                            //dgdModelInfo.Weight = lib.returnNumStringOne(dgdModelInfo.Weight);
                            //dgdModelInfo.Exdiameter = lib.returnNumStringOne(dgdModelInfo.Exdiameter);

                            dgdArticleCode.Items.Add(dgdModelInfo);

                        }

                        tbkIndexCount.Text = "▶검색결과 : " + i + " 건";
                    }
                }
            }
            catch (Exception ee)
            {
                MessageBox.Show(ee.Message);
            }
            finally
            {
                DataStore.Instance.CloseConnection();
            }
        }

        #endregion // 조회

        #region 메인 그리드 선택시 - 해당 데이터의 서브그리드 검색 이벤트

        //조회 시
        private void FillGridProcess(string strArticleID)
        {
            if (dgdProcess.ItemsSource != null)
            {
                dgdProcess.ItemsSource = null;
                ovcArticleProcess.Clear();
            }

            try
            {
                Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                sqlParameter.Add("chkArticleID", 1);
                sqlParameter.Add("sArticleID", strArticleID);

                DataSet ds = DataStore.Instance.ProcedureToDataSet("xp_Article_sArticleProcess", sqlParameter, false);

                if (ds != null && ds.Tables.Count > 0)
                {
                    DataTable dt = ds.Tables[0];

                    if (dt.Rows.Count > 0)
                    {
                        int i = 0;
                        DataRowCollection drc = dt.Rows;

                        foreach (DataRow dr in drc)
                        {
                            i++;
                            var dgdProcessInfo = new Process_CodeView()
                            {
                                Num = i,
                                ProcessID = dr["ProcessID"].ToString(),
                                Process = dr["Process"].ToString(),
                                ArticleID = dr["ArticleID"].ToString()
                            };

                            if (dgdProcessInfo.ArticleID != null && !dgdProcessInfo.ArticleID.Equals(""))
                            {
                                if (dgdProcessInfo.ArticleID.Equals(strArticleID))
                                {
                                    dgdProcessInfo.CheckFlag = true;
                                }
                            }

                            ovcArticleProcess.Add(dgdProcessInfo);
                            //dgdProcess.Items.Add(dgdProcessInfo);

                        }

                        dgdProcess.ItemsSource = ovcArticleProcess;
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



        //조회 시
        private void FillGridCustomArticle(string strArticleID)
        {
            if (dgdCustomArticle.ItemsSource != null)
            {
                dgdCustomArticle.ItemsSource = null;
                ovcCustomArticle.Clear();
            }

            try
            {
                Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                sqlParameter.Add("sArticleID", strArticleID);

                DataSet ds = DataStore.Instance.ProcedureToDataSet("xp_Article_sArticleSelectCustom", sqlParameter, false);

                if (ds != null && ds.Tables.Count > 0)
                {
                    DataTable dt = ds.Tables[0];

                    if (dt.Rows.Count > 0)
                    {
                        int i = 0;
                        DataRowCollection drc = dt.Rows;

                        foreach (DataRow dr in drc)
                        {
                            i++;
                            var dgdCustomArticleInfo = new CustomArticle_CodeView()
                            {
                                Num = i,

                                ArticleID = dr["ArticleID"].ToString(),
                                CustomBuyArticle = dr["CustomBuyArticle"].ToString(),
                            };

                            if (dgdCustomArticleInfo.ArticleID != null && !dgdCustomArticleInfo.ArticleID.Equals(""))
                            {
                                if (dgdCustomArticleInfo.ArticleID.Equals(strArticleID))
                                {
                                    dgdCustomArticleInfo.CheckFlag = true;
                                }
                            }

                            ovcCustomArticle.Add(dgdCustomArticleInfo);
                            //dgdProcess.Items.Add(dgdProcessInfo);

                        }

                        dgdCustomArticle.ItemsSource = ovcCustomArticle;
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

        #endregion // 메인 그리드 선택시 - 해당 데이터의 서브그리드 검색 이벤트

        #region 저장

        //저장
        private bool SaveData(string strID, string strFlag)
        {
            bool flag = false;
            List<Procedure> Prolist = new List<Procedure>();
            List<Dictionary<string, object>> ListParameter = new List<Dictionary<string, object>>();


            //List<string[]> UpdateFilesInfo = new List<string[]>();


            string GetKey = "";

            if (CheckData(txtCode.Text, strFlag)) //2021-07-21 품번이 같은 경우를 확인하기 위해 변경 CheckData(txtName.Text, strFlag)
            {
                try
                {
                    Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                    sqlParameter.Clear();


                    sqlParameter.Add("sArticleID", txtCode.Text != null && !txtCode.Text.Trim().Equals("") ? txtCode.Text : "");
                    sqlParameter.Add("sNewArticleID", "");
                    sqlParameter.Add("CompanyID", MainWindow.CompanyID);
                    sqlParameter.Add("sArticle", txtName.Text.Trim().Equals("") ? txtBuyerArticleNo.Text : txtName.Text);
                    sqlParameter.Add("BuyerArticleNo", txtBuyerArticleNo.Text);

                    sqlParameter.Add("sArticleGrpID", cboArticleGrp.SelectedValue != null ? cboArticleGrp.SelectedValue.ToString() : "");
                    sqlParameter.Add("sProductGrpID", cboProductGrpID.SelectedValue != null ? cboProductGrpID.SelectedValue.ToString() : "");
                    sqlParameter.Add("sSupplyType", cboSupplyType.SelectedValue != null ? cboSupplyType.SelectedValue.ToString() : "");
                    sqlParameter.Add("sPartGBNID", cboPartGBNID.SelectedValue != null ? cboPartGBNID.SelectedValue.ToString() : "");
                    sqlParameter.Add("sUseingType", "");

                    //sqlParameter.Add("Weight", ConvertDouble(txtWeight.Text)); // 입력란 없음 → 0으로 세팅

                    sqlParameter.Add("sUseClss", chkNotUse.IsChecked == true ? "*" : "");
                    //sqlParameter.Add("sSpec", txtSpec.T-ext);
                    sqlParameter.Add("sSpec", txtSpec.Text != null ? txtSpec.Text.Trim() : "");


                    sqlParameter.Add("NeedStockQty", ConvertDouble(txtNeedStockQty.Text)); // 적정재고량 소수로 변환            
                    sqlParameter.Add("sUnitClss", cboUnitClss.SelectedValue != null ? cboUnitClss.SelectedValue.ToString() : "");


                    sqlParameter.Add("ProdQtyPerBox", ConvertDouble(TextBoxProdQtyPerBox.Text));
                    sqlParameter.Add("OutQtyPerBox", ConvertDouble(TextBoxOutQtyPerBox.Text));
                    sqlParameter.Add("sImageFIleName", "");
                    sqlParameter.Add("sImageFilePath", "");

                    sqlParameter.Add("sLabelPrintYN", cboLabelPrintYN.SelectedValue != null ? cboLabelPrintYN.SelectedValue.ToString() : "");
                    sqlParameter.Add("nUnitPrice", ConvertDouble(txtUnitPrice.Text.Replace(",", "")));
                    sqlParameter.Add("UnitPriceClss", cboPriceClss.SelectedValue != null ? cboPriceClss.SelectedValue.ToString() : "");

                    sqlParameter.Add("OutUnitPrice", ConvertDouble(txtOutUnitPrice.Text.Replace(",", "")));
                    sqlParameter.Add("sFTAMgrYN", cboFTAMgrYN.SelectedValue != null ? cboFTAMgrYN.SelectedValue.ToString() : "");
                    sqlParameter.Add("sHSCode", txtHSCode.Text);

                    sqlParameter.Add("sBuySaleMainYN", chkBuySaleMainYN.IsChecked == true ? "Y" : "N");

                    sqlParameter.Add("sComments", txtComments.Text);
                    sqlParameter.Add("sPART_ATTR", cboPART_ATTR.SelectedIndex == -1 || cboPART_ATTR.SelectedValue == null ? "" : cboPART_ATTR.SelectedValue.ToString());
                    sqlParameter.Add("sPatternID", cboPattern.SelectedIndex == -1 || cboPattern.SelectedValue == null ? "" : cboPattern.SelectedValue.ToString());
                    sqlParameter.Add("sCoatingSpec", txtCoatingSpec.Text); //도면번호


                    sqlParameter.Add("sExDiameter", txtExdiameter.Text != null && !txtExdiameter.Text.Trim().Equals("") ? ConvertDouble(txtExdiameter.Text) : 0);
                    sqlParameter.Add("sInDiameter", txtInDiameter.Text != null && !txtInDiameter.Text.Trim().Equals("") ? ConvertDouble(txtInDiameter.Text) : 0);
                    sqlParameter.Add("sWeight", txtWeight.Text != null && !txtWeight.Text.Trim().Equals("") ? ConvertDouble(txtWeight.Text) : 0);
                    sqlParameter.Add("sLength", txtLength.Text != null && !txtLength.Text.Trim().Equals("") ? ConvertDouble(txtLength.Text) : 0);
                    sqlParameter.Add("BuyerModelID", txtModel.Tag != null ? txtModel.Tag.ToString() : "");
                    sqlParameter.Add("sBigMiSmalGbn", cboBigMiSmal.SelectedValue != null ? cboBigMiSmal.SelectedValue.ToString() : ""); //공정구분으로 쓸꺼
                    sqlParameter.Add("sInspectYN", cboInsPectYN.SelectedValue != null ? cboInsPectYN.SelectedValue.ToString() : ""); //무검사품 여부




                    #region 추가
                    if (strFlag.Equals("I"))
                    {
                        sqlParameter.Add("CreateUserID", MainWindow.CurrentUser);

                        Procedure pro1 = new Procedure();
                        pro1.Name = "xp_Article_iArticle";
                        pro1.OutputUseYN = "Y";
                        pro1.OutputName = "sNewArticleID";

                        pro1.OutputLength = "10";

                        Prolist.Add(pro1);
                        ListParameter.Add(sqlParameter);

                        List<KeyValue> list_Result = new List<KeyValue>();
                        list_Result = DataStore.Instance.ExecuteAllProcedureOutputGetCS_NewLog(Prolist, ListParameter, "C");
                        //string sGetID = string.Empty;
                        if (list_Result[0].key.ToLower() == "success")
                        {
                            list_Result.RemoveAt(0);
                            for (int i = 0; i < list_Result.Count; i++)
                            {
                                KeyValue kv = list_Result[i];
                                if (kv.key == "sNewArticleID")
                                {
                                    GetKey = kv.value;
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

                        Prolist.Clear();
                        ListParameter.Clear();


                        // 공정 선택한거 넣기
                        for (int i = 0; i < dgdProcess.Items.Count; i++)
                        {
                            var WinProcess = dgdProcess.Items[i] as Process_CodeView;
                            if (WinProcess != null && WinProcess.CheckFlag == true)
                            {
                                sqlParameter = new Dictionary<string, object>();
                                sqlParameter.Clear();

                                //sqlParameter.Add("sArticleID", txtCode.Text != null && !txtCode.Text.Trim().Equals("") ? txtCode.Text : "");
                                sqlParameter.Add("sArticleID", GetKey);
                                sqlParameter.Add("sProcessID", WinProcess.ProcessID);
                                sqlParameter.Add("UseYN", "Y");
                                sqlParameter.Add("UserID", MainWindow.CurrentUser);

                                Procedure pro2 = new Procedure();
                                pro2.Name = "xp_Article_iArticleProcess";
                                pro2.OutputUseYN = "N";
                                pro2.OutputName = "sArticleID";
                                pro2.OutputLength = "10";

                                Prolist.Add(pro2);
                                ListParameter.Add(sqlParameter);
                            }
                        }

                        string[] Confirm = new string[2];
                        Confirm = DataStore.Instance.ExecuteAllProcedureOutputNew_NewLog(Prolist, ListParameter, "C");
                        if (Confirm[0] != "success")
                        {
                            MessageBox.Show("[저장실패]\r\n" + Confirm[1].ToString());
                            flag = false;
                            return false;
                        }
                        else
                        {
                            flag = true;
                        }
                    }
                    #endregion // 추가

                    #region 수정
                    else // 수정
                    {
                        sqlParameter.Add("sUpdateUserID", MainWindow.CurrentUser);

                        Procedure pro1 = new Procedure();
                        pro1.Name = "xp_Article_uArticle";
                        pro1.OutputUseYN = "N";
                        pro1.OutputName = "sArticleID";
                        pro1.OutputLength = "10";

                        Prolist.Add(pro1);
                        ListParameter.Add(sqlParameter);

                        for (int i = 0; i < dgdProcess.Items.Count; i++)
                        {
                            var WinProcess = dgdProcess.Items[i] as Process_CodeView;
                            if (WinProcess != null && WinProcess.CheckFlag == true)
                            {
                                sqlParameter = new Dictionary<string, object>();
                                sqlParameter.Clear();

                                sqlParameter.Add("sArticleID", strID);
                                sqlParameter.Add("sProcessID", WinProcess.ProcessID);
                                sqlParameter.Add("UseYN", "Y");
                                sqlParameter.Add("UserID", MainWindow.CurrentUser);

                                Procedure pro2 = new Procedure();
                                pro2.Name = "xp_Article_iArticleProcess";
                                pro2.OutputUseYN = "N";
                                pro2.OutputName = "sArticleID";
                                pro2.OutputLength = "10";

                                Prolist.Add(pro2);
                                ListParameter.Add(sqlParameter);
                            }
                        }

                        string[] Confirm = new string[2];
                        Confirm = DataStore.Instance.ExecuteAllProcedureOutputNew_NewLog(Prolist, ListParameter, "U");
                        if (Confirm[0] != "success")
                        {
                            MessageBox.Show("[저장실패]\r\n" + Confirm[1].ToString());
                            flag = false;
                            return false;
                        }
                        else
                        {
                            GetKey = strID;
                            flag = true;
                        }
                    }
                    #endregion


                    //if (!_ftp.UploadTempFilesToFTP(UpdateFilesInfo))
                    //{
                    //    MessageBox.Show("파일업로드에 실패하였습니다.");
                    //    flag = false;
                    //    return flag;
                    //}

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

                        //if (listFtpFile.Count == 0)
                        //{
                        //    MessageBox.Show("이미지어디감?");
                        //    flag = false;
                        //    return flag;

                        //}

                        if (listFtpFile.Count > 0)
                        {
                            FTP_Save_File(listFtpFile, GetKey);
                        }


                        UpdateDBFtp(GetKey); // 리스트 갯수가 0개 이상일때 해버리면, 수정시에 저장이 안됨
                    }

                    // 파일 List 비워주기
                    listFtpFile.Clear();
                    deleteListFtpFile.Clear();
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

        #endregion // 저장

        #region 유효성 검사
        //삭제체크
        private bool DeleteDataCheck(string strArticleID)
        {
            bool Flag = false;

            try
            {
                Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                sqlParameter.Clear();
                sqlParameter.Add("sArticleID", strArticleID);

                string[] result = DataStore.Instance.ExecuteProcedure("xp_Article_dArticle_Check", sqlParameter, false);

                if (result[0].Equals("success") && result[1].Equals(""))
                {
                    //MessageBox.Show("성공 *^^*");
                    Flag = true;
                }
                else
                {
                    MessageBox.Show(result[1]);
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

            return Flag;
        }
        private bool CheckData(string strName, string strFlag)
        {
            bool flag = true;

            if (strFlag.Equals("I"))
            {
                //if (txtCode.Text.Length < 5)
                //{
                //    MessageBox.Show("5자리의 코드를 입력해주세요");
                //    return false;
                //}

                // 2020.02.20 일단 품명 중복검사 제외
                // 2021.07.21 일단 품명 중복검사 제외
                if (!GetArticleByName(strName))
                {
                    MessageBox.Show("이미 같은 이름의 코드가 존재합니다.(사용안함 포함)"); //2021-07-21 품번이 같을 경우 MessageBox.Show("이미 같은 이름의 품명이 존재합니다.");
                    return false;
                }


                // 2020.02.20 품번이 필수 입력이 되어야함!!!
                if (txtBuyerArticleNo.Text.Trim().Equals(""))
                {
                    MessageBox.Show("품번이 입력되지 않았습니다.");
                    return flag = false;
                }
            }

            //if(strFlag.Equals("U"))
            //{
            //    if(!GetArticleByName(strName))
            //    {

            //    }
            //}

            if (cboArticleGrp.SelectedIndex == -1 || cboArticleGrp.SelectedValue == null)
            {
                MessageBox.Show("품명그룹을 선택해주세요");
                flag = false;
                return flag;
            }

            if (cboUnitClss.SelectedIndex == -1 || cboUnitClss.SelectedValue == null)
            {
                MessageBox.Show("단위를 선택해주세요");
                flag = false;
                return flag;
            }

            // 제품군 체크
            if (cboProductGrpID.SelectedIndex == -1 || cboProductGrpID.SelectedValue == null)
            {
                MessageBox.Show("제품군을 선택해주세요");
                flag = false;
                return flag;
            }
            // 부품분류
            if (cboPartGBNID.SelectedIndex == -1 || cboPartGBNID.SelectedValue == null)
            {
                MessageBox.Show("부품분류를 선택해주세요");
                flag = false;
                return flag;
            }

            if (cboLabelPrintYN.SelectedIndex == -1 || cboLabelPrintYN.SelectedValue == null)
            {
                MessageBox.Show("라벨관리를 선택해주세요");
                flag = false;
                return flag;
            }

            // 적정재고량(필수입력) 빈칸체크, 수량 소수 변환 가능한지 체크
            if (txtNeedStockQty.Text.Length <= 0 || txtNeedStockQty.Text.Replace(" ", "").Equals(""))
            {
                MessageBox.Show("적정재고량을 입력해주세요.");
                flag = false;
                return flag;
            }

            if (cboSupplyType.SelectedIndex == -1 || cboSupplyType.SelectedValue == null)
            {
                MessageBox.Show("공급유형을 선택해주세요");
                flag = false;
                return flag;
            }


            if (cboFTAMgrYN.SelectedIndex == -1 || cboFTAMgrYN.SelectedValue == null)
            {
                MessageBox.Show("FTP중점관리품 여부를 선택해주세요");
                flag = false;
                return flag;
            }

            //if ( == null)
            //{
            //    MessageBox.Show("넣어라");
            //    flag = false;
            //    return flag;
            //}

            //bool proFlag = false;
            //for (int i = 0; i < dgdProcess.Items.Count; i++)
            //{
            //    var Pro = dgdProcess.Items[i] as Process_CodeView;

            //    if (Pro != null
            //        && Pro.CheckFlag == true)
            //    {
            //        proFlag = true;
            //    }
            //}

            //if (proFlag == false)
            //{
            //    MessageBox.Show("생산공정선택에서 공정을 선택해주세요");
            //    flag = false;
            //    return flag;
            //}

            return flag;
        }

        #endregion 유효성 검사

        //품목이 있는지 체크(추가시에만)
        private bool GetArticleByName(string strArticleName)
        {
            bool flag = true;

            try
            {
                Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                sqlParameter.Add("sArticle", strArticleName);

                // DataSet ds = DataStore.Instance.ProcedureToDataSet("xp_Article_sArticleByName", sqlParameter, false); 2021-07-21
                DataSet ds = DataStore.Instance.ProcedureToDataSet("xp_Article_sArticleByBuyerArticleNo", sqlParameter, false);
                if (ds != null && ds.Tables.Count > 0)
                {
                    DataTable dt = ds.Tables[0];

                    if (dt.Rows.Count > 0)
                    {
                        //string strName = dt.Rows[0]["Article"].ToString(); 2021-07-21
                        string strName = dt.Rows[0]["ArticleID"].ToString();
                        //string strName = dt.Rows[0]["BuyerArticleNo"].ToString();
                        if (!strName.Replace(" ", "").Equals(""))
                        {
                            flag = false;
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


        #region DB 파일명 수정 프로시저 



        private bool UpdateDBFtp(string strArticleID)
        {
            bool flag = false;

            string str_localpath = string.Empty;

            List<string[]> UpdateFilesInfo = new List<string[]>();


            //if (CheckDataFTP(txtName.Text, strFlag))
            //{
            try
            {
                Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                sqlParameter.Clear();
                sqlParameter.Add("sArticleID", strArticleID);
                sqlParameter.Add("sSketch1Path", !txtSketch1.Text.Trim().Equals("") ? "/ImageData/Article/" + strArticleID : "");
                sqlParameter.Add("sSketch1File", txtSketch1.Text);
                sqlParameter.Add("sSketch2Path", !txtSketch2.Text.Trim().Equals("") ? "/ImageData/Article/" + strArticleID : "");
                sqlParameter.Add("sSketch2File", txtSketch2.Text);

                sqlParameter.Add("sSketch3Path", !txtSketch3.Text.Trim().Equals("") ? "/ImageData/Article/" + strArticleID : "");
                sqlParameter.Add("sSketch3File", txtSketch3.Text);
                sqlParameter.Add("sSketch4Path", !txtSketch4.Text.Trim().Equals("") ? "/ImageData/Article/" + strArticleID : "");
                sqlParameter.Add("sSketch4File", txtSketch4.Text);
                sqlParameter.Add("sSketch5Path", !txtSketch5.Text.Trim().Equals("") ? "/ImageData/Article/" + strArticleID : "");

                sqlParameter.Add("sSketch5File", txtSketch5.Text);
                sqlParameter.Add("sSketch6Path", !txtSketch6.Text.Trim().Equals("") ? "/ImageData/Article/" + strArticleID : "");
                sqlParameter.Add("sSketch6File", txtSketch6.Text);

                //sqlParameter.Add("sSketch7Path", !txtSketch7.Text.Trim().Equals("") ? "/ImageData/Article/" + strArticleID : "");
                //sqlParameter.Add("sSketch7File", txtSketch7.Text);

                sqlParameter.Add("sUpdateUserID", MainWindow.CurrentUser);

                string[] result = DataStore.Instance.ExecuteProcedure("xp_Article_uArticle_FTP", sqlParameter, true);


                #region 연습
                //System.IO.FileNotFoundException

                //if (_ftp.UploadTempFilesToFTP(UpdateFilesInfo)
                //   && sqlParameter != null)
                //{
                //    MessageBox.Show("파일 확인 좀?");
                //    flag = false;
                //    return flag;

                //str_localpath = txtSketch1.Text;

                //FileInfo fi = new FileInfo(str_localpath);
                //if (!fi.Exists)
                //{
                //    MessageBox.Show("파일이 없다");
                //    flag = false;
                //    return flag;
                //}
                #endregion

                //MoveTo("/ImageData/Article/" + strArticleID);


                if (result[0].Equals("success"))
                {
                    //MessageBox.Show("성공 *^^*");
                    flag = true;
                }
                else
                {
                    MessageBox.Show("수정 실패 , 내용 : " + result[1]);
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
            //}


            return flag;
        }

        #endregion // DB 파일명 수정 프로시저 

        //ftp 체크
        //private bool CheckDataFTP(string strName, string strFlag)
        //{
        //    _ftp = new FTP_EX(FTP_ADDRESS, FTP_ID, FTP_PASS);

        //    bool flag = true;

        //    if (_ftp.upload(strName, strFlag) == false)
        //    {
        //        MessageBox.Show("이미지가 정상적으로 등록되지 않았습니다.");
        //        flag = false;
        //        return flag;
        //    }

        //    return flag;
        //}

        #endregion // 주요 메서드

        #region 기타 메서드 모음

        // 천마리 콤마, 소수점 버리기
        private string stringFormatN0(object obj)
        {
            return string.Format("{0:N0}", obj);
        }

        // 천마리 콤마, 소수점 한자리
        private string stringFormatN1(object obj)
        {
            return string.Format("{0:N1}", obj);
        }

        // 천마리 콤마, 소수점 두자리
        private string stringFormatN2(object obj)
        {
            return string.Format("{0:N2}", obj);
        }

        // 천마리 콤마, 소수점 두자리
        private string stringFormatN4(object obj)
        {
            return string.Format("{0:N4}", obj);
        }

        // 자릿수 설정
        private string stringFormatN_Number(object obj, int num)
        {
            return string.Format("{0:N" + num + "}", obj);
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
                str = str.Trim().Replace(",", "");

                if (Double.TryParse(str, out chkDouble) == true)
                {
                    result = Double.Parse(str);
                }
            }

            return result;
        }



        #endregion



        #region FTP 

        // image 만 Bit로 세팅( imageSource랑 바인딩 )
        private BitmapImage SetImage(string strAttachPath)
        {
            _ftp = new FTP_EX(FTP_ADDRESS, FTP_ID, FTP_PASS);
            BitmapImage bit = _ftp.DrawingImageByByte(FTP_ADDRESS + strAttachPath + "");
            return bit;





        }

        private BitmapImage SetImage(string ImageName, string FolderName)
        {
            //bool ExistFile = false;
            BitmapImage bit = null;
            _ftp = new FTP_EX(FTP_ADDRESS, FTP_ID, FTP_PASS);
            if (_ftp == null) { return null; }

            //string[] fileListDetail;
            //fileListDetail = _ftp.directoryListSimple(FolderName, Encoding.Default);

            //ExistFile = FileInfoAndFlag(fileListDetail, ImageName);
            //if (ExistFile)
            //{
            bit = _ftp.DrawingImageByByte(FTP_ADDRESS + '/' + FolderName + '/' + ImageName + "");
            //}

            return bit;
        }

        //
        private void FTP_DownLoadFile(string strFilePath)
        {
            _ftp = new FTP_EX(FTP_ADDRESS, FTP_ID, FTP_PASS);

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
                file.Delete();
            }

            _ftp.download(str_remotepath.Substring(str_remotepath.Substring(0, str_remotepath.LastIndexOf("/")).LastIndexOf("/")), str_localpath);

            ProcessStartInfo proc = new ProcessStartInfo(str_localpath);
            proc.UseShellExecute = true;
            Process.Start(proc);
        }

        #region 이미지 파일이 있으면 보기버튼 활성화, 아니면 비활성화


        // 파일체크
        public void MoveTo(string destFileName)
        {
            //string FileName = "ftp://" + LoadINI.FileSvr + ":" + LoadINI.FTPPort + LoadINI.FtpImagePath + "/Article";
            //FileInfo fInfo = new FileInfo(FileName);
            //fInfo.MoveTo("ftp://" + LoadINI.FileSvr + ":" + LoadINI.FTPPort + LoadINI.FtpImagePath + "/Article");

            //if (FileName != LoadINI.FtpImagePath + "/Article")
            //{
            //    MessageBox.Show("ㅠㅠ");
            //}
        }


        private void btnImgSeeCheckAndSetting()
        {
            if (!txtSketch1.Text.Trim().Equals(""))
            {
                btnFileSee1.IsEnabled = true;
            }
            else
            {
                btnFileSee1.IsEnabled = false;
            }

            if (!txtSketch2.Text.Trim().Equals(""))
            {
                btnFileSee2.IsEnabled = true;
            }
            else
            {
                btnFileSee2.IsEnabled = false;
            }

            if (!txtSketch3.Text.Trim().Equals(""))
            {
                btnFileSee3.IsEnabled = true;
            }
            else
            {
                btnFileSee3.IsEnabled = false;
            }

            if (!txtSketch4.Text.Trim().Equals(""))
            {
                btnFileSee4.IsEnabled = true;
            }
            else
            {
                btnFileSee4.IsEnabled = false;
            }

            if (!txtSketch5.Text.Trim().Equals(""))
            {
                btnFileSee5.IsEnabled = true;
            }
            else
            {
                btnFileSee5.IsEnabled = false;
            }

            if (!txtSketch6.Text.Trim().Equals(""))
            {
                btnFileSee6.IsEnabled = true;
            }
            else
            {
                btnFileSee6.IsEnabled = false;
            }

            //if (!txtSketch7.Text.Trim().Equals(""))
            //{
            //    btnFileSee7.IsEnabled = true;
            //}
            //else
            //{
            //    btnFileSee7.IsEnabled = false;
            //}

        }

        #endregion // 이미지 파일이 있으면 보기버튼 활성화, 아니면 비활성화

        // 보기 버튼 클릭
        private void btnFileSee_Click(object sender, RoutedEventArgs e)
        {
            MessageBoxResult msgresult = MessageBox.Show("파일을 보시겠습니까?", "보기 확인", MessageBoxButton.YesNo);
            if (msgresult == MessageBoxResult.Yes)
            {
                //버튼 태그값.
                string ClickPoint = ((Button)sender).Tag.ToString();

                if ((ClickPoint == "01") && (txtSketch1.Text == string.Empty)
                    || (ClickPoint == "02") && (txtSketch2.Text == string.Empty)
                    || (ClickPoint == "03") && (txtSketch3.Text == string.Empty)
                    || (ClickPoint == "04") && (txtSketch4.Text == string.Empty)
                    || (ClickPoint == "05") && (txtSketch5.Text == string.Empty)
                    || (ClickPoint == "06") && (txtSketch6.Text == string.Empty))
                //|| (ClickPoint == "07") && (txtSketch7.Text == string.Empty))
                {
                    MessageBox.Show("파일이 없습니다.");
                    return;
                }

                try
                {
                    // 접속 경로
                    _ftp = new FTP_EX(FTP_ADDRESS, FTP_ID, FTP_PASS);

                    string str_path = string.Empty;
                    str_path = FTP_ADDRESS + '/' + txtCode.Text;
                    _ftp = new FTP_EX(str_path, FTP_ID, FTP_PASS);

                    string str_remotepath = string.Empty;
                    string str_localpath = string.Empty;

                    if (ClickPoint == "01") { str_remotepath = txtSketch1.Text; }
                    else if (ClickPoint == "02") { str_remotepath = txtSketch2.Text; }
                    else if (ClickPoint == "03") { str_remotepath = txtSketch3.Text; }
                    else if (ClickPoint == "04") { str_remotepath = txtSketch4.Text; }
                    else if (ClickPoint == "05") { str_remotepath = txtSketch5.Text; }
                    else if (ClickPoint == "06") { str_remotepath = txtSketch6.Text; }
                    // else if (ClickPoint == "07") { str_remotepath = txtSketch7.Text; }


                    #region temp 폴더에 저장 후에 여는건데 이걸 왜씀?

                    if (ClickPoint == "01") { str_localpath = LOCAL_DOWN_PATH + "\\" + txtSketch1.Text; }
                    else if (ClickPoint == "02") { str_localpath = LOCAL_DOWN_PATH + "\\" + txtSketch2.Text; }
                    else if (ClickPoint == "03") { str_localpath = LOCAL_DOWN_PATH + "\\" + txtSketch3.Text; }
                    else if (ClickPoint == "04") { str_localpath = LOCAL_DOWN_PATH + "\\" + txtSketch4.Text; }
                    else if (ClickPoint == "05") { str_localpath = LOCAL_DOWN_PATH + "\\" + txtSketch5.Text; }
                    else if (ClickPoint == "06") { str_localpath = LOCAL_DOWN_PATH + "\\" + txtSketch6.Text; }
                    // else if (ClickPoint == "07") { str_localpath = LOCAL_DOWN_PATH + "\\" + txtSketch7.Text; }

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

                    //ProcessStartInfo proc = new ProcessStartInfo(str_localpath);
                    //proc.UseShellExecute = true;
                    //Process.Start(proc);

                    #endregion // temp 폴더에 저장 후에 여는건데 이걸 왜씀?

                    // 위에건 뭐여
                    var Article = dgdArticleCode.SelectedItem as Win_com_Article_U_CodeView;
                    if (Article != null)
                    {

                        if (CheckImage(str_remotepath.Trim()))
                        {
                            imgSetting.Source = SetImage(str_remotepath, Article.ArticleID);
                        }
                        else
                        {
                            MessageBox.Show(winArticleCode.ImageName + "는 이미지 변환이 불가능합니다.");
                        }
                    }


                }
                catch (Exception ex) // 뭐든 간에 파일 없다고 하자
                {
                    MessageBox.Show("파일이 존재하지 않습니다.\r관리자에게 문의해주세요.");
                    return;
                }
            }
        }

        // 이미지 더블클릭 → 이미지 크게 보기 
        private void imgSetting_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            if (e.ClickCount == 2)
            {
                Image image = sender as Image;
                BitmapImage bitmapImage = (BitmapImage)image.Source; // bitmapImage
                LargeImagePopUp largeImagePopUp = new LargeImagePopUp(bitmapImage);
                largeImagePopUp.ShowDialog();
            }
        }

        #endregion


        #region Content - FTP : 사진올리기, 삭제, 보기

        // 파일 업로드
        private void btnFileUpload_Click(object sender, RoutedEventArgs e)
        {
            // (버튼)sender 마다 tag를 달자.
            string ClickPoint = ((Button)sender).Tag.ToString();
            if (ClickPoint.Contains("01")) { FTP_Upload_TextBox(txtSketch1); }  //긴 경로(FULL 사이즈)
            else if (ClickPoint.Contains("02")) { FTP_Upload_TextBox(txtSketch2); }
            else if (ClickPoint.Contains("03")) { FTP_Upload_TextBox(txtSketch3); }
            else if (ClickPoint.Contains("04")) { FTP_Upload_TextBox(txtSketch4); }
            else if (ClickPoint.Contains("05")) { FTP_Upload_TextBox(txtSketch5); }
            else if (ClickPoint.Contains("06")) { FTP_Upload_TextBox(txtSketch6); }
            //else if (ClickPoint.Contains("07")) { FTP_Upload_TextBox(txtSketch7); }

            // 보기 버튼체크 → 업로드 한 상태(저장하기 전)에서는 아직 FTP에 파일이 올라가지 않았기 때문에 활성화 시키면 안되지
            //btnImgSeeCheckAndSetting();
        }

        #region FTP_Upload_TextBox - 파일 경로, 이름 텍스트박스에 올림 + 리스트에 ADD

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
                //OFdlg.Filter =
                //    "Image files (*.jpg, *.jpeg, *.jpe, *.jfif, *.png, *.pcx, *.pdf) | *.jpg; *.jpeg; *.jpe; *.jfif; *.png; *.pcx; *.pdf | All Files|*.*";

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

                        string[] strTemp = new string[] { ImageFileName, ImageFilePath.ToString() };
                        listFtpFile.Add(strTemp);
                    }
                }
            }
        }

        #endregion // FTP_Upload_TextBox - 파일 경로, 이름 텍스트박스에 올림 + 리스트에 ADD

        #region FTP_Save_File - 파일 저장, 폴더 생성


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

            bool Makefind = false;
            Makefind = FileInfoAndFlag(fileListSimple, MakeFolderName);


            if (Makefind == false)
            {
                //bool flag = false;

                FileInfo fi = new FileInfo(fileListSimple.ToString());
                if (!fi.Exists)
                {
                    MessageBox.Show("이미지가 등록되지 않았습니다. 다시 시도하십시오.");

                    strFlag = string.Empty;
                    re_Search(rowNum);
                    //flag = false;
                    //return;
                    // 파일 없을 시
                }
                else
                {
                    //flag = true;
                    return;
                    // 파일 있을 시
                }

            }

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

        }

        #endregion // FTP_Save_File - 파일 저장, 폴더 생성

        #region FTP 파일 삭제

        //파일만 삭제 - 버튼에 Tag로 구분
        private void btnFileDelete_Click(object sender, RoutedEventArgs e)
        {
            MessageBoxResult msgresult = MessageBox.Show("파일을 삭제 하시겠습니까?", "삭제 확인", MessageBoxButton.YesNo);
            if (msgresult == MessageBoxResult.Yes)
            {
                string ClickPoint = ((Button)sender).Tag.ToString();

                if ((ClickPoint == "01") && (txtSketch1.Text != string.Empty)) { FileDeleteAndTextBoxEmpty(txtSketch1); }
                else if ((ClickPoint == "02") && (txtSketch2.Text != string.Empty)) { FileDeleteAndTextBoxEmpty(txtSketch2); }
                else if ((ClickPoint == "03") && (txtSketch3.Text != string.Empty)) { FileDeleteAndTextBoxEmpty(txtSketch3); }
                else if ((ClickPoint == "04") && (txtSketch4.Text != string.Empty)) { FileDeleteAndTextBoxEmpty(txtSketch4); }
                else if ((ClickPoint == "05") && (txtSketch5.Text != string.Empty)) { FileDeleteAndTextBoxEmpty(txtSketch5); }
                else if ((ClickPoint == "06") && (txtSketch6.Text != string.Empty)) { FileDeleteAndTextBoxEmpty(txtSketch6); }
                //  else if ((ClickPoint == "07") && (txtSketch7.Text != string.Empty)) { FileDeleteAndTextBoxEmpty(txtSketch7); }
            }

            // 보기 버튼체크
            btnImgSeeCheckAndSetting();
        }
        private void FileDeleteAndTextBoxEmpty(TextBox txt)
        {
            if (strFlag.Equals("U"))
            {
                var Article = dgdArticleCode.SelectedItem as Win_com_Article_U_CodeView;

                if (Article != null)
                {
                    //FTP_RemoveFile(Article.ArticleID + "/" + txt.Text);

                    // 파일이름, 파일경로
                    string[] strFtp = { txt.Text, txt.Tag != null ? txt.Tag.ToString() : "" };

                    deleteListFtpFile.Add(strFtp);
                }
            }

            txt.Text = "";
            txt.Tag = "";
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






        #endregion // FTP 파일 삭제

        #endregion // Content - FTP : 사진올리기, 삭제, 보기

        #region 텍스트 박스 엔터 → 다음 텍스트 박스

        // 1. 품명
        private void txtName_KeyDown(object sender, KeyEventArgs e)
        {
            //if (e.Key == Key.Enter)
            //{
            //    txtSpec.Focus();
            //}
        }

        // 2. 세부내역
        private void txtSpec_KeyDown(object sender, KeyEventArgs e)
        {
            //if (e.Key == Key.Enter)
            //{
            //    cboArticleGrp.IsDropDownOpen = true;
            //}
        }

        // 3. 품명 그룹
        private void cboArticleGrp_DropDownClosed(object sender, EventArgs e)
        {
            //if (lblMsg.Visibility == Visibility.Visible)
            //{
            //    cboLabelPrintYN.IsDropDownOpen = true;
            //}

        }

        // 5. 제품군
        private void cboProductGrpID_DropDownClosed(object sender, EventArgs e)
        {
            //cboSupplyType.IsDropDownOpen = true;
        }

        // 6. 공급유형
        private void cboSupplyType_DropDownClosed(object sender, EventArgs e)
        {
            //cboPartGBNID.IsDropDownOpen = true;
        }

        // 7. 부품분류
        private void cboPartGBNID_DropDownClosed(object sender, EventArgs e)
        {
            //txtNeedStockQty.Focus();
        }

        // 8. 적정재고량
        private void txtNeedStockQty_KeyDown(object sender, KeyEventArgs e)
        {
            //if (e.Key == Key.Enter)
            //{
            //    cboLabelPrintYN.IsDropDownOpen = true;
            //}
        }


        // 9. 라벨 관리
        private void cboLabelPrintYN_DropDownClosed(object sender, EventArgs e)
        {
            //txtOutUnitPrice.Focus();
        }

        // 10. 출고 단가
        private void txtOutUnitPrice_KeyDown(object sender, KeyEventArgs e)
        {
            //if (e.Key == Key.Enter)
            //{
            //    cboPriceClss.IsDropDownOpen = true;
            //}
        }

        // 11. 화폐단위

        // 12. 단위
        private void cboUnitClss_DropDownClosed(object sender, EventArgs e)
        {
            //TextBoxOutQtyPerBox.Focus();
        }

        // 13. 출하용박스 수량
        private void TextBoxOutQtyPerBox_KeyDown(object sender, KeyEventArgs e)
        {
            //if (e.Key == Key.Enter)
            //{
            //    cboFTAMgrYN.IsDropDownOpen = true;
            //}
        }

        // 14. FTA중점관리품
        private void cboFTAMgrYN_DropDownClosed(object sender, EventArgs e)
        {
            //txtUnitPrice.Focus();
        }

        // 15. 단가
        private void txtUnitPrice_KeyDown(object sender, KeyEventArgs e)
        {
            //if (e.Key == Key.Enter)
            //{
            //    cboPART_ATTR.Focus();
            //}
        }

        // 16. 원재료 속성
        private void cboPART_ATTR_DropDownClosed(object sender, EventArgs e)
        {
            // TextBoxProdQtyPerBox.Focus();
        }

        // 17. 생산박스당 수량
        private void TextBoxProdQtyPerBox_KeyDown(object sender, KeyEventArgs e)
        {
            //if (e.Key == Key.Enter)
            //{
            //    txtHSCode.Focus();
            //}
        }


        // 18. H.S CODE

        // 19. 품번
        private void txtBuyerArticleNo_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                txtName.Text = txtBuyerArticleNo.Text;
                txtName.Focus();
                txtName.CaretIndex = txtName.Text.Length;
                //txtName.SelectAll();
            }
        }

        // 20. 보관 위치
        //private void txtStockLocName_KeyDown(object sender, KeyEventArgs e)
        //{
        //    if (e.Key == Key.Enter)
        //    {
        //        txtCoatingSpec.Focus();
        //    }
        //}

        // 21. 도금사양
        private void txtCoatingSpec_KeyDown(object sender, KeyEventArgs e)
        {
            //if (e.Key == Key.Enter)
            //{
            //    txtComments.Focus();
            //}
        }

        // 22. 비고
        private void txtComments_KeyDown(object sender, KeyEventArgs e)
        {
            //if (e.Key == Key.Enter)
            //{
            //    cboPattern.IsDropDownOpen = true;
            //}
        }

        private void cboPattern_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            //btnSave.Focus();
        }




        #endregion // 텍스트 박스 엔터 → 다음 텍스트 박스

        private void TextProdDiffiLevel_KeyDown(object sender, KeyEventArgs e)
        {

        }

        private void TxtBuyerArticleNoSrh_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                rowNum = 0;
                re_Search(rowNum);
            }
        }

        //난이도 텍스트박스에 숫자만 입력
        private void TextProdDiffiLevel_PreviewTextInput(object sender, TextCompositionEventArgs e)
        {
            lib.CheckIsNumeric((TextBox)sender, e);
        }

        //적정재고량 텍스트박스에 숫자만 입력
        private void txtNeedStockQty_PreviewTextInput(object sender, TextCompositionEventArgs e)
        {
            lib.CheckIsNumeric((TextBox)sender, e);
        }

        //출하용박스 수량   숫자만 입력
        private void TextBoxOutQtyPerBox_PreviewTextInput(object sender, TextCompositionEventArgs e)
        {
            lib.CheckIsNumeric((TextBox)sender, e);
        }

        //생산박스당 수량 숫자만 입력
        private void TextBoxProdQtyPerBox_PreviewTextInput(object sender, TextCompositionEventArgs e)
        {
            lib.CheckIsNumeric((TextBox)sender, e);
        }

        //매입단가 숫자만 입력
        private void txtUnitPrice_PreviewTextInput(object sender, TextCompositionEventArgs e)
        {
            lib.CheckIsNumeric((TextBox)sender, e);
        }

        //출고,가공단가 숫자만 입력
        private void txtOutUnitPrice_PreviewTextInput(object sender, TextCompositionEventArgs e)
        {
            lib.CheckIsNumeric((TextBox)sender, e);
        }

        //중량 숫자만 입력
        private void txtWeight_PreviewTextInput(object sender, TextCompositionEventArgs e)
        {
            lib.CheckIsNumeric((TextBox)sender, e);
        }

        //폭 숫자만 입력
        private void txtExdiameter_PreviewTextInput(object sender, TextCompositionEventArgs e)
        {
            lib.CheckIsNumeric((TextBox)sender, e);
        }

        //길이 숫자만 입력   
        private void txtLength_PreviewTextInput(object sender, TextCompositionEventArgs e)
        {
            lib.CheckIsNumeric((TextBox)sender, e);
        }



        private void cboFreeStuffinYN_DropDownClosed(object sender, EventArgs e)
        {

        }


        //라벨 클릭 품명
        private void lblArticleSrh_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            if (chkArticleSrh.IsChecked == true)
            {
                chkArticleSrh.IsChecked = false;
            }
            else
            {
                chkArticleSrh.IsChecked = true;
            }
        }

        //품명체크
        private void chkArticleSrh_Checked(object sender, RoutedEventArgs e)
        {
            chkArticleSrh.IsChecked = true;
            txtArticleSrh.IsEnabled = true;
        }
        //품명 노체크
        private void chkArticleSrh_Unchecked(object sender, RoutedEventArgs e)
        {
            chkArticleSrh.IsChecked = false;
            txtArticleSrh.IsEnabled = false;
        }

        //품명 텍박
        private void TxtArticleSrh_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                rowNum = 0;
                re_Search(rowNum);
            }
        }
        //대중소구분 깡통
        private void CboBigMiSmal_DropDownClosed(object sender, EventArgs e)
        {

        }

        private void TxtOutUnitPrice_PreviewKeyDown(object sender, TextCompositionEventArgs e)
        {
            Lib.Instance.CheckIsNumeric((TextBox)sender, e);
        }

        private void TxtUnitPrice_PreviewTextInput_1(object sender, TextCompositionEventArgs e)
        {
            Lib.Instance.CheckIsNumeric((TextBox)sender, e);
        }

        //차종 키다운 
        private void txtModel_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                MainWindow.pf.ReturnCode(txtModel, (int)Defind_CodeFind.DCF_BUYERMODEL, "");

            }
        }

        //차종
        private void btnPfModel_Click(object sender, RoutedEventArgs e)
        {
            MainWindow.pf.ReturnCode(txtModel, (int)Defind_CodeFind.DCF_BUYERMODEL, "");

        }



        ////숫자 외에 다른 문자열 ㅁ
        //public bool IsNumeric(string source)
        //{
        //    Regex regex = new Regex("[^0-9.-]+");
        //    return !regex.IsMatch(source);
        //}

    }

    #region CodeView

    class Win_com_Article_U_CodeView : BaseView
    {
        public override string ToString()
        {
            return (this.ReportAllProperties());
        }

        public string Num { get; set; }

        public string ArticleID { get; set; }
        public string Article { get; set; }
        public string ThreadID { get; set; }
        public string Thread { get; set; }
        public string StuffWidth { get; set; }

        public string DyeingID { get; set; }
        public string Weight { get; set; }
        public string UseClss { get; set; }
        public string Spec { get; set; }
        public string ImageName { get; set; }

        public string FirstArticlePath { get; set; }
        public string FirstArticleFile { get; set; }
        public string MediumArticlePath { get; set; }
        public string MediumArticleFIle { get; set; }
        public string CompleteArticlePath { get; set; }

        public string CompleteArticleFile { get; set; }
        public string Defect1Path { get; set; }
        public string Defect1File { get; set; }
        public string Defect2Path { get; set; }
        public string Defect2File { get; set; }

        public string Defect3Path { get; set; }
        public string Defect3File { get; set; }

        public string Sketch1Path { get; set; }
        public string Sketch1File { get; set; }
        public string Sketch2Path { get; set; }
        public string Sketch2File { get; set; }
        public string Sketch3Path { get; set; }

        public string Sketch3File { get; set; }
        public string Sketch4Path { get; set; }
        public string Sketch4File { get; set; }
        public string Sketch5Path { get; set; }
        public string Sketch5File { get; set; }

        public string Sketch6Path { get; set; }
        public string Sketch6File { get; set; }


        public string Sketch7Path { get; set; }
        public string Sketch7File { get; set; }

        public string ProcessID { get; set; }
        public string Process { get; set; }

        public string SupplyType { get; set; }
        public string SupplyTypeName { get; set; }

        public string UseingType { get; set; }
        public string UseingTypeName { get; set; }

        public string UnitClss { get; set; }
        public string UnitClssName { get; set; }

        public string LabelPrintYN { get; set; }
        public string Qty { get; set; }
        public string NeedStockQty { get; set; }
        public string ArticleGrpID { get; set; }
        public string ArticleGrp { get; set; }

        public string QtyPerBox { get; set; }
        public string BuyerArticleNo { get; set; }
        public string Unitprice { get; set; }
        public string UnitPriceClss { get; set; }
        public string HSCODE { get; set; }

        public string FTAMgrYN { get; set; }
        public string CoatingSpec { get; set; }
        public string BuySaleMainYN { get; set; }
        public string FreeStuffinYN { get; set; }
        public string Comments { get; set; }
        public string PART_ATTR { get; set; }

        public string PatternID { get; set; }
        public string OutUnitPrice { get; set; }
        public string StockLocName { get; set; }
        public string UnitPrice { get; set; }
        public string UseClssName { get; set; }

        public string PartGBNID { get; set; }
        public string PartGBNName { get; set; }
        public string ProductGrpID { get; set; }
        public string ProductGrpName { get; set; }
        public string CompanyID { get; set; }

        public string ProdQtyPerBox { get; set; }
        public string OutQtyPerBox { get; set; }

        public string Exdiameter { get; set; } // 외경
        public string InDiameter { get; set; } // 내경
        public string Length { get; set; } // 길이

        public string DrawingLength { get; set; } // 내경
        public string CuttingLength { get; set; } // 내경
        public string DregsFront { get; set; } // 내경
        public string DregsBack { get; set; } // 내경
        public string SawBladeLoss { get; set; } // 내경
        public string BonCuttingQty { get; set; } // 내경

        public string BladeQty { get; set; } // 내경
        public string BladeEndType { get; set; } // 내경
        public string BladeType { get; set; } // 내경
        public string CotingType { get; set; } // 내경

        public string ProdDiffiLevel { get; set; } // 난이도 
        public string BigMiSmalGbn { get; set; } // 대중소구분 
        public string BuyerModelID { get; set; }
        public string BuyerModel { get; set; }
        public string InspectYN { get; set; } // 무검사품여부 
        public string ArticleGrpName { get; set; }

        public BitmapImage ImageView { get; set; }

    }

    class Process_CodeView : BaseView
    {
        public override string ToString()
        {
            return (this.ReportAllProperties());
        }

        public int Num { get; set; }
        public string Process { get; set; }
        public string ProcessID { get; set; }
        public string ArticleID { get; set; }
        public bool CheckFlag { get; set; }
    }

    class CustomArticle_CodeView : BaseView
    {
        public override string ToString()
        {
            return (this.ReportAllProperties());
        }

        public int Num { get; set; }
        public string ArticleID { get; set; }
        public string CustomBuyArticle { get; set; }
        public bool CheckFlag { get; set; }
    }

    class Mold_CodeView : BaseView
    {
        public override string ToString()
        {
            return (this.ReportAllProperties());
        }

        public int Num { get; set; }
        public string MoldID { get; set; }
        public string MoldName { get; set; }
        public string MoldNo { get; set; }
        public string ArticleID { get; set; }
        public string MoldKindName { get; set; }
        public string Code_ID { get; set; }
        public bool CheckFlag { get; set; }
    }

    #endregion
}

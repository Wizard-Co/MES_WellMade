using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Data;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using System.Windows.Data;
using WizMes_WellMade.PopUP;
using System.Windows.Input;
using System.Windows.Media.Imaging;
using System.IO;
using System.Drawing;
using System.Text;
using System.Diagnostics;
using System.Net;
using System.Threading;

namespace WizMes_WellMade
{
    /// <summary>
    /// Win_com_BuseoJikChaek_U.xaml에 대한 상호 작용 논리
    /// </summary>
    public partial class Win_dvl_MoldRegularInspectBasis_U : UserControl
    {
        string strFlag = "";
        int rowNum = 0;

#if DEBUG
        string FTP_ADDRESS = "ftp://121.254.224.196/ImageData/MoldBasis";
#else
 string FTP_ADDRESS = "ftp://" + LoadINI.FileSvr + ":"
            + LoadINI.FTPPort + LoadINI.FtpImagePath + "/MoldBasis";
#endif

        //FTP 활용모음
        private const string FTP_ID = "wizuser";
        private const string FTP_PASS = "wiz9999";
        private const string LOCAL_DOWN_PATH = "C:\\Temp";

        private FTP_EX _ftp = null;
        List<string[]> listFtpFile = new List<string[]>();
        List<string[]> deleteListFtpFile = new List<string[]>();

        bool FTP_Trigger = false;

        // 복사 추가용 변수
        // 이미지 이름 : 폴더이름
        Dictionary<string, string> lstFtpFilePath = new Dictionary<string, string>();

        bool editing = false;

        public Win_dvl_MoldRegularInspectBasis_U()
        {
            InitializeComponent();
        }

        private void UserControl_Loaded(object sender, RoutedEventArgs e)
        {
            Lib.Instance.UiLoading(sender);
            SetComboBox();
        }

        private void SetComboBox()
        {
            ObservableCollection<CodeView> ovcMcInsCycleGbnSrh = ComboBoxUtil.Instance.Gf_DB_CM_GetComCodeDataset(null, "MCCYCLEGBN", "Y", "", "");
            this.cboInspectCycle.ItemsSource = ovcMcInsCycleGbnSrh;
            this.cboInspectCycle.DisplayMemberPath = "code_name";
            this.cboInspectCycle.SelectedValuePath = "code_id";

        }

        #region 추가 수정 모드 / 저장완료 취소 모드

        // 추가, 수정 모드
        private void SaveUpdateMode()
        {
            // 메시지
            if (strFlag.Trim().Equals("I"))
            {
                tbkMsg.Text = "자료 추가중";
            }
            else
            {
                tbkMsg.Text = "자료 수정중";
            }
            lblMsg.Visibility = Visibility.Visible;

            // 검색조건
            grdSrh1.IsEnabled = false;
            grdSrh2.IsEnabled = false;

            // 버튼
            btnCopy.IsEnabled = false;
            btnAdd.IsEnabled = false;
            btnUpdate.IsEnabled = false;
            btnDelete.IsEnabled = false;
            btnSearch.IsEnabled = false;
            btnExcel.IsEnabled = false;

            btnSave.Visibility = Visibility.Visible;
            btnCancel.Visibility = Visibility.Visible;
            btnSave.IsEnabled = true;
            btnCancel.IsEnabled = true;

            // 메인 그리드
            dgdMain.IsHitTestVisible = false;

            // 입력사항
            grdInput.IsHitTestVisible = true;

            // 서브그리드
            btnAddSub.IsEnabled = true;
            btnDeleteSub.IsEnabled = true;
            //dgdSub.IsReadOnly = true;
        }

        // 저장완료, 취소 모드
        private void CompleteCancelMode()
        {

            lblMsg.Visibility = Visibility.Hidden;

            // 검색조건
            grdSrh1.IsEnabled = true;
            grdSrh2.IsEnabled = true;

            // 버튼
            btnCopy.IsEnabled = true;
            btnAdd.IsEnabled = true;
            btnUpdate.IsEnabled = true;
            btnDelete.IsEnabled = true;
            btnSearch.IsEnabled = true;
            btnExcel.IsEnabled = true;

            btnSave.Visibility = Visibility.Hidden;
            btnCancel.Visibility = Visibility.Hidden;

            // 메인 그리드
            dgdMain.IsHitTestVisible = true;

            // 입력사항
            grdInput.IsHitTestVisible = false;

            // 서브그리드
            btnAddSub.IsEnabled = false;
            btnDeleteSub.IsEnabled = false;
            //dgdSub.IsReadOnly = false;
        }

        #endregion // 추가 수정 모드 / 저장완료 취소 모드

        #region Header 부분 - 검색 조건

        // 철형명 검색
        private void lblMoldSrh_MouseLeftButtonUp(object sender, System.Windows.Input.MouseButtonEventArgs e)
        {
            if (chkMoldNoSrh.IsChecked == true)
            {
                chkMoldNoSrh.IsChecked = false;
            }
            else
            {
                chkMoldNoSrh.IsChecked = true;
            }
        }
        private void chkMoldSrh_Checked(object sender, RoutedEventArgs e)
        {
            chkMoldNoSrh.IsChecked = true;
            txtMoldNoSrh.IsEnabled = true;
        }
        private void chkMoldSrh_Unchecked(object sender, RoutedEventArgs e)
        {
            chkMoldNoSrh.IsChecked = false;
            txtMoldNoSrh.IsEnabled = false;
        }
        // 품명
        private void lblArticleSrh_MouseLeftButtonUp(object sender, System.Windows.Input.MouseButtonEventArgs e)
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
        private void chkArticleSrh_Checked(object sender, RoutedEventArgs e)
        {
            chkArticleSrh.IsChecked = true;
            txtArticleSrh.IsEnabled = true;
        }
        private void chkArticleSrh_Unchecked(object sender, RoutedEventArgs e)
        {
            chkArticleSrh.IsChecked = false;
            txtArticleSrh.IsEnabled = false;
        }

        #endregion // Header 부분 - 검색 조건

        #region Header 부분 - 오른쪽 상단 버튼

        // 복사 추가 버튼
        private void btnCopy_Click(object sender, RoutedEventArgs e)
        {
            var Mold = dgdMain.SelectedItem as Win_dvl_MoldRegularInspectBasis_U_CodeView;

            if (Mold != null)
            {
                rowNum = 0;

                strFlag = "I";
                SaveUpdateMode();

                // 철형 ID 만 초기화
                txtMoldInspectBasisID.Text = "";

                // 사진 보관
                string Key = Mold.MoldInspectBasisID;

                for (int i = 0; i < dgdSub.Items.Count; i++)
                {
                    var MoldSub = dgdSub.Items[i] as Win_dvl_MoldRegularInspectBasis_U_CodeViewSub;

                    if (MoldSub != null 
                        && MoldSub.MoldInspectImageFile != null
                        && !MoldSub.MoldInspectImageFile.Trim().Equals(""))
                    {
                        lstFtpFilePath.Add(MoldSub.MoldInspectImageFile, Key);
                    }
                }
            }
            else
            {
                MessageBox.Show("복사 추가할 대상을 선택해주세요.");
                return;
            }
        }

        // 추가 버튼
        private void btnAdd_Click(object sender, RoutedEventArgs e)
        {
            rowNum = 0;

            strFlag = "I";
            SaveUpdateMode();

            // 전부 초기화
            this.DataContext = null;

            // 날짜 오늘날짜로
            dtpMoldInspectBasisDate.SelectedDate = DateTime.Today;

            if (dgdSub.Items.Count > 0)
            {
                dgdSub.Items.Clear();
            }
        }

        // 수정 버튼
        private void btnUpdate_Click(object sender, RoutedEventArgs e)
        {
            var Mold = dgdMain.SelectedItem as Win_dvl_MoldRegularInspectBasis_U_CodeView;

            if (Mold != null)
            {
                rowNum = dgdMain.SelectedIndex;

                strFlag = "U";
                SaveUpdateMode();
            }
            else
            {
                MessageBox.Show("수정할 대상을 선택해주세요.");
                return;
            }            
        }

        // 삭제 버튼
        private void btnDelete_Click(object sender, RoutedEventArgs e)
        {
            var WinMold = dgdMain.SelectedItem as Win_dvl_MoldRegularInspectBasis_U_CodeView;

            if (WinMold == null)
            {
                MessageBox.Show("삭제할 데이터를 선택해주세요.");
            }
            else
            {
                if (MessageBox.Show("선택하신 항목을 삭제하시겠습니까?", "삭제 전 확인", MessageBoxButton.YesNo) == MessageBoxResult.Yes)
                {
                    if (Delete(WinMold.MoldInspectBasisID))
                    {
                        rowNum = 0;
                        re_Search(rowNum);
                    }
                }
            }
        }

        // 닫기 버튼
        private void btnClose_Click(object sender, RoutedEventArgs e)
        {
            Lib.Instance.ChildMenuClose(this.ToString());
        }

        // 검색 버튼
        private void btnSearch_Click(object sender, RoutedEventArgs e)
        {
            rowNum = 0;
            re_Search(rowNum);
        }

        // 저장 버튼
        private void btnSave_Click(object sender, RoutedEventArgs e)
        {
            if (SaveData(strFlag))
            {
                CompleteCancelMode();
                strFlag = "";
                re_Search(rowNum);
            }
        }

        // 취소 버튼
        private void btnCancel_Click(object sender, RoutedEventArgs e)
        {   
            CompleteCancelMode();
            strFlag = "";
            re_Search(rowNum);
        }

        // 엑셀 버튼
        private void btnExcel_Click(object sender, RoutedEventArgs e)
        {
            DataTable dt = null;
            string Name = string.Empty;

            string[] dgdStr = new string[2];
            dgdStr[0] = "철형현황 등록";
            dgdStr[1] = dgdMain.Name;

            ExportExcelxaml ExpExc = new ExportExcelxaml(dgdStr);
            ExpExc.ShowDialog();

            if (ExpExc.DialogResult.HasValue)
            {
                if (ExpExc.choice.Equals(dgdMain.Name))
                {
                    if (ExpExc.Check.Equals("Y"))
                        dt = Lib.Instance.DataGridToDTinHidden(dgdMain);
                    else
                        dt = Lib.Instance.DataGirdToDataTable(dgdMain);

                    Name = dgdMain.Name;
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

        #endregion // Header 부분 - 오른쪽 상단 버튼

        // 메인 그리드 선택 이벤트
        private void dgdMain_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            var Mold = dgdMain.SelectedItem as Win_dvl_MoldRegularInspectBasis_U_CodeView;

            if (Mold != null)
            {
                this.DataContext = Mold;

                FillGridSub(Mold.MoldInspectBasisID);
            }
        }

        #region Content - 작성 부분

        // 철형번호 엔터 → 플러스파인더
        private void txtMoldID_KeyDown(object sender, System.Windows.Input.KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                MainWindow.pf.ReturnCode(txtMoldID, 5555, "");
                if (txtMoldID.Tag != null)
                {
                    SetArticle(txtMoldID.Tag.ToString());
                }
            }
        }
        // 철형번호 플러스파인더
        private void btnPfMoldID_Click(object sender, RoutedEventArgs e)
        {
            MainWindow.pf.ReturnCode(txtMoldID, 5555, "");
            if (txtMoldID.Tag != null)
            {
                SetArticle(txtMoldID.Tag.ToString());
            }
        }

        private void SetArticle(string moldID)
        {
            try
            {
                Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                sqlParameter.Add("MoldID", moldID);

                DataSet ds = DataStore.Instance.ProcedureToDataSet("xp_dvlMold_getMoldArticle", sqlParameter, false);
                if (ds != null && ds.Tables.Count > 0)
                {
                    DataTable dt = ds.Tables[0];
                    if (dt.Rows.Count > 0)
                    {
                        DataRow dr = dt.Rows[0];

                        txtArticle.Text = dr["Article"].ToString();
                        txtBuyerArticleNo.Text = dr["BuyerArticleNo"].ToString();
                    }
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
          }
            finally
            {
                DataStore.Instance.CloseConnection();
            }
            
        }

        // 개정일자
        private void lblMoldInspectBasisDate_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            if (chkMoldInspectBasisDate.IsChecked == true)
            {
                chkMoldInspectBasisDate.IsChecked = false;
            }
            else
            {
                chkMoldInspectBasisDate.IsChecked = true;
            }
        }
        private void chkMoldInspectBasisDate_Checked(object sender, RoutedEventArgs e)
        {
            chkMoldInspectBasisDate.IsChecked = true;
            dtpMoldInspectBasisDate.IsEnabled = true;
        }
        private void chkMoldInspectBasisDate_UnChecked(object sender, RoutedEventArgs e)
        {
            chkMoldInspectBasisDate.IsChecked = false;
            dtpMoldInspectBasisDate.IsEnabled = false;
        }

        #endregion // Content - 작성 부분

        #region 서브그리드 버튼 - 추가 삭제

        // 추가
        private void btnAddSub_Click(object sender, RoutedEventArgs e)
        {
            // 인덱스 구하기
            int index = 1;

            if (dgdSub.Items.Count > 0)
            {
                var MoldIndex = dgdSub.Items[dgdSub.Items.Count - 1] as Win_dvl_MoldRegularInspectBasis_U_CodeViewSub;

                if (MoldIndex != null)
                {
                    index = MoldIndex.Num + 1;
                }
            }
            
            // 추가하기
            var Mold = new Win_dvl_MoldRegularInspectBasis_U_CodeViewSub()
            {
                Num = index,

                MoldInspectBasisID = "",
                MoldSeq = 0,
                MoldInspectItemName = "",
                MoldInspectContent = "",
                MoldInspectCheckGbn = "",

                MoldInspectCheckGbn_Name = "",
                MoldInspectCycleGbn = "",
                MoldInspectCycleGbn_Name = "",
                MoldInspectCycleDate = "",
                MoldInspectRecordGbn = "",

                MoldInspectRecordGbn_Name = "",
                MoldInspectComments = "",
                MoldInspectImageFile = "",

                btnName = "업로드",
            };

            dgdSub.Items.Add(Mold);

            int endRow = dgdSub.Items.Count - 1;
            dgdSub.SelectedIndex = endRow;

            editing = true;
            dgdSub.Focus();
            dgdSub.CurrentCell = new DataGridCellInfo(dgdSub.Items[endRow], dgdSub.Columns[1]);
        }

        // 삭제
        private void btnDeleteSub_Click(object sender, RoutedEventArgs e)
        {
            var MoldSub = dgdSub.SelectedItem as Win_dvl_MoldRegularInspectBasis_U_CodeViewSub;

            if (MoldSub != null)
            {
                // 이미지가 있으면 삭제
                if (MoldSub.MoldInspectImageFile != null
                    && !MoldSub.MoldInspectImageFile.Trim().Equals(""))
                {
                    string[] strFtp = { MoldSub.MoldInspectImageFile, MoldSub.MoldInspectBasisID };

                    deleteListFtpFile.Add(strFtp);
                }

                int index = dgdSub.SelectedIndex;

                dgdSub.Items.Remove(MoldSub);

                if (dgdSub.Items.Count > 0)
                {
                    if (index == 0)
                    {
                        dgdSub.SelectedIndex = 0;
                    }
                    else
                    {
                        dgdSub.SelectedIndex = index - 1;
                    }

                }
            }
        }



        #endregion // 서브그리드 버튼 - 추가 삭제

        #region 서브그리드 입력

        // 확인방법
        private void cboMoldInspectCheckGbn_Loaded(object sender, RoutedEventArgs e)
        {
            ComboBox cboMoldInspectCheckGbn = sender as ComboBox;
            if (cboMoldInspectCheckGbn.ItemsSource == null)
            {
                ObservableCollection<CodeView> ovcMoldInspectCheckGbn = ComboBoxUtil.Instance.Gf_DB_CM_GetComCodeDataset(null, "MLDCHECKGBN", "Y", "", "");
                cboMoldInspectCheckGbn.ItemsSource = ovcMoldInspectCheckGbn;
                cboMoldInspectCheckGbn.DisplayMemberPath = "code_name";
                cboMoldInspectCheckGbn.SelectedValuePath = "code_id";
            }

            if (editing == true)
            {
                (sender as ComboBox).IsDropDownOpen = true;

                editing = false;
            }
        }
        private void cboMoldInspectCheckGbn_DropDownClosed(object sender, EventArgs e)
        {
            ComboBox cboMoldInspectCheckGbn = sender as ComboBox;
            var MoldSub = cboMoldInspectCheckGbn.DataContext as Win_dvl_MoldRegularInspectBasis_U_CodeViewSub;

            if (MoldSub != null)
            {
                MoldSub.MoldInspectCheckGbn = cboMoldInspectCheckGbn.SelectedValue != null ? cboMoldInspectCheckGbn.SelectedValue.ToString() : "";
                MoldSub.MoldInspectCheckGbn_Name = cboMoldInspectCheckGbn.Text;
            }

            editing = true;
            int currRow = dgdSub.Items.IndexOf(dgdSub.CurrentItem);
            dgdSub.CurrentCell = new DataGridCellInfo(dgdSub.Items[currRow], dgdSub.Columns[4]);
        }

        // 주기
        private void cboMoldInspectCycleGbn_Loaded(object sender, RoutedEventArgs e)
        {
            ComboBox cboMoldInspectCycleGbn = sender as ComboBox;
            if (cboMoldInspectCycleGbn.ItemsSource == null)
            {
                ObservableCollection<CodeView> ovcMoldInspectCycleGbn = ComboBoxUtil.Instance.Gf_DB_CM_GetComCodeDataset(null, "MLDCYCLEGBN", "Y", "", "");
                cboMoldInspectCycleGbn.ItemsSource = ovcMoldInspectCycleGbn;
                cboMoldInspectCycleGbn.DisplayMemberPath = "code_name";
                cboMoldInspectCycleGbn.SelectedValuePath = "code_id";
            }

            if (editing == true)
            {
                (sender as ComboBox).IsDropDownOpen = true;

                editing = false;
            }
        }
        private void cboMoldInspectCycleGbn_DropDownClosed(object sender, EventArgs e)
        {
            ComboBox cboMoldInspectCycleGbn = sender as ComboBox;
            var MoldSub = cboMoldInspectCycleGbn.DataContext as Win_dvl_MoldRegularInspectBasis_U_CodeViewSub;

            if (MoldSub != null)
            {
                MoldSub.MoldInspectCycleGbn = cboMoldInspectCycleGbn.SelectedValue != null ? cboMoldInspectCycleGbn.SelectedValue.ToString() : "";
                MoldSub.MoldInspectCycleGbn_Name = cboMoldInspectCycleGbn.Text;
            }

            editing = true;
            int currRow = dgdSub.Items.IndexOf(dgdSub.CurrentItem);
            dgdSub.CurrentCell = new DataGridCellInfo(dgdSub.Items[currRow], dgdSub.Columns[5]);
        }

        // 기록구분
        private void cboMoldInspectRecordGbn_Loaded(object sender, RoutedEventArgs e)
        {
            ComboBox cboMoldInspectRecordGbn = sender as ComboBox;
            if (cboMoldInspectRecordGbn.ItemsSource == null)
            {
                ObservableCollection<CodeView> ovcMoldInspectRecordGbn = ComboBoxUtil.Instance.Gf_DB_CM_GetComCodeDataset(null, "MLDRECORDGBN", "Y", "", "");
                cboMoldInspectRecordGbn.ItemsSource = ovcMoldInspectRecordGbn;
                cboMoldInspectRecordGbn.DisplayMemberPath = "code_name";
                cboMoldInspectRecordGbn.SelectedValuePath = "code_id";
            }

            if (editing == true)
            {
                (sender as ComboBox).IsDropDownOpen = true;

                editing = false;
            }
        }
        private void cboMoldInspectRecordGbn_DropDownClosed(object sender, EventArgs e)
        {
            ComboBox cboMoldInspectRecordGbn = sender as ComboBox;
            var MoldSub = cboMoldInspectRecordGbn.DataContext as Win_dvl_MoldRegularInspectBasis_U_CodeViewSub;

            if (MoldSub != null)
            {
                MoldSub.MoldInspectRecordGbn = cboMoldInspectRecordGbn.SelectedValue != null ? cboMoldInspectRecordGbn.SelectedValue.ToString() : "";
                MoldSub.MoldInspectRecordGbn_Name = cboMoldInspectRecordGbn.Text;
            }

            //editing = true;
            //int currRow = dgdSub.Items.IndexOf(dgdSub.CurrentItem);
            //dgdSub.CurrentCell = new DataGridCellInfo(dgdSub.Items[currRow], dgdSub.Columns[8]);
        }

        // 이미지 업로드 / 삭제 이벤트
        private void btnUploadAndDelete_Click(object sender, RoutedEventArgs e)
        {
            if (lblMsg.Visibility == Visibility.Visible)
            {
                Button senderBtn = sender as Button;

                var MoldSub = senderBtn.DataContext as Win_dvl_MoldRegularInspectBasis_U_CodeViewSub;

                if (MoldSub != null
                    && MoldSub.btnName != null)
                {
                    if (MoldSub.btnName.Trim().Equals("업로드"))
                    {
                        if (FTP_Upload_SetImage(MoldSub))
                        {
                            senderBtn.Content = "삭제";

                            int currRow = dgdSub.Items.IndexOf(dgdSub.CurrentItem);
                            dgdSub.CurrentCell = new DataGridCellInfo(dgdSub.Items[currRow], dgdSub.Columns[7]);

                            editing = false;
                        }
                    }
                    else // 삭제
                    {
                        // 파일 삭제
                        if (FileDeleteAndTextBoxEmpty(MoldSub))
                        {
                            //MoldSub.btnName = "업로드";
                            senderBtn.Content = "업로드";

                            int currRow = dgdSub.Items.IndexOf(dgdSub.CurrentItem);
                            dgdSub.CurrentCell = new DataGridCellInfo(dgdSub.Items[currRow], dgdSub.Columns[7]);
                        }

                    }
                }

            }
        }

        #region FTP 모음

        #region FTP 업로드 + 이미지 세팅

        private bool FTP_Upload_SetImage(Win_dvl_MoldRegularInspectBasis_U_CodeViewSub MoldSub)
        {
            bool flag = false;

            if (lblMsg.Visibility == Visibility.Visible)
            {
                try
                {
                    Microsoft.Win32.OpenFileDialog OFdlg = new Microsoft.Win32.OpenFileDialog();
                    OFdlg.Filter =
                        "Image files (*.jpg, *.jpeg, *.jpe, *.jfif, *.png, *.pcx, *.pdf) | *.jpg; *.jpeg; *.jpe; *.jfif; *.png; *.pcx; *.pdf | All Files|*.*";

                    Nullable<bool> result = OFdlg.ShowDialog();
                    if (result == true)
                    {
                        string strFullPath = OFdlg.FileName;

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
                            flag = false;
                            return flag;
                        }
                        else
                        {

                            MoldSub.MoldInspectImageFile = ImageFileName;
                            //MoldSub.ImagePath = ImageFilePath;

                            Bitmap image = new Bitmap(ImageFilePath + ImageFileName);

                            MoldSub.ImageByte = BitmapToImageSource(image);

                            string[] strTemp = new string[] { ImageFileName, ImageFilePath.ToString() };
                            listFtpFile.Add(strTemp);

                            flag = true;
                        }
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                    flag = false;
                    return flag;
                }
            }

            return flag;
        }

        #endregion // FTP 업로드 + 이미지 세팅

        #region FTP 파일 삭제

        private bool FileDeleteAndTextBoxEmpty(Win_dvl_MoldRegularInspectBasis_U_CodeViewSub MoldSub)
        {
            bool flag = false;

            if (lblMsg.Visibility == Visibility.Visible)
            {
                MessageBoxResult msgresult = MessageBox.Show("파일을 삭제 하시겠습니까?", "삭제 확인", MessageBoxButton.YesNo);
                if (msgresult == MessageBoxResult.Yes)
                {
                    //FTP_RemoveFile(Article.ArticleID + "/" + txt.Text);

                    // 파일이름, 파일경로
                    string[] strFtp = { MoldSub.MoldInspectImageFile, MoldSub.MoldInspectBasisID };

                    deleteListFtpFile.Add(strFtp);

                    MoldSub.ImageByte = null;
                    MoldSub.MoldInspectImageFile = "";

                    flag = true;
                }
            }

            return flag;
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

        #endregion FTP 파일 삭제

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

                // 이미지 사이즈 제한 
                bitmapimage.DecodePixelWidth = 200;
                bitmapimage.DecodePixelHeight = 200;

                bitmapimage.EndInit();
                bitmapimage.Freeze();

                return bitmapimage;
            }
        }

        private BitmapImage SetImage(string ImageName, string FolderName)
        {
            BitmapImage bit = null;
            _ftp = new FTP_EX(FTP_ADDRESS, FTP_ID, FTP_PASS);
            if (_ftp == null) { return null; }

            bit = DrawingImageByByte2(FTP_ADDRESS + '/' + FolderName + '/' + ImageName + "");


            return bit;
        }

        private BitmapImage DrawingImageByByte2(string ftpFilePath)
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

                    // 이미지 사이즈 제한 
                    image.DecodePixelWidth = 200;
                    image.DecodePixelHeight = 200;

                    image.EndInit();
                    image.Freeze();
                }

            }
            catch (Exception ex)
            {
                System.Windows.MessageBox.Show("1" + ex.Message + " / " + ex.Source);
                FTP_Trigger = false;
                //throw ex;
            }

            return image;
        }

        #region 복사 추가 메서드

        // FTP Byte 로 저장하기
        private void FTP_Save_FileByFtpServerFilePath(Dictionary<string, string> lstFtpFilePath, string Key)
        {
            try
            {
                // 폴더 경로 포함해서 다시 생성 후 스따뜨
                _ftp = new FTP_EX(FTP_ADDRESS, FTP_ID, FTP_PASS);

                List<string> imgs = new List<string>();

                string[] fileListSimple;
                string[] fileListDetail = null;
                fileListSimple = _ftp.directoryListSimple("", Encoding.Default);

                string MakeFolderName = Key;

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

                _ftp.UploadUsingFtpServerFilePath(lstFtpFilePath, Key);
            }
            catch (Exception ep1)
            {
                MessageBox.Show(ep1.Message);
            }
        }

        #endregion // 복사 추가 메서드

        #region FTP_Save_File - 파일 저장, 폴더 생성

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

        #endregion // FTP 모음

        #endregion // 서브그리드 입력

        #region 서브 데이터그리드 키 입력 이벤트

        // 2019.08.27 PreviewKeyDown 는 key 다운과 같은것 같음
        private void DataGird_PreviewKeyDown(object sender, KeyEventArgs e)
        {
            try
            {
                if (e.Key == Key.Down || e.Key == Key.Up || e.Key == Key.Left || e.Key == Key.Right)
                {
                    DataGird_KeyDown(sender, e);
                }
            }
            catch (Exception ex)
            {

            }
        }
        // KeyDown 이벤트
        private void DataGird_KeyDown(object sender, KeyEventArgs e)
        {
            int currRow = dgdSub.Items.IndexOf(dgdSub.CurrentItem);
            int currCol = dgdSub.Columns.IndexOf(dgdSub.CurrentCell.Column);
            int startCol = 1;
            int endCol = 8;

            if (e.Key == Key.Enter)
            {
                e.Handled = true;
                (sender as DataGridCell).IsEditing = false;

                editing = true;

                // 마지막 열, 마지막 행 아님
                if (endCol == currCol && dgdSub.Items.Count - 1 > currRow)
                {
                    dgdSub.SelectedIndex = currRow + 1; // 이건 한줄 파란색으로 활성화 된 걸 조정하는 것입니다.
                    dgdSub.CurrentCell = new DataGridCellInfo(dgdSub.Items[currRow + 1], dgdSub.Columns[startCol]);

                } // 마지막 열 아님
                else if (endCol > currCol && dgdSub.Items.Count - 1 >= currRow)
                {
                    dgdSub.CurrentCell = new DataGridCellInfo(dgdSub.Items[currRow], dgdSub.Columns[currCol + 1]);
                } // 마지막 열, 마지막 행
                else if (endCol == currCol && dgdSub.Items.Count - 1 == currRow)
                {
                    //btnSave.Focus();
                }
                else
                {
                    MessageBox.Show("나머지가 있나..");
                }
            }
            else if (e.Key == Key.Down)
            {
                e.Handled = true;
                (sender as DataGridCell).IsEditing = false;

                // 마지막 행 아님
                if (dgdSub.Items.Count - 1 > currRow)
                {
                    dgdSub.SelectedIndex = currRow + 1;
                    dgdSub.CurrentCell = new DataGridCellInfo(dgdSub.Items[currRow + 1], dgdSub.Columns[currCol]);
                } // 마지막 행일때
                else if (dgdSub.Items.Count - 1 == currRow)
                {
                    if (endCol > currCol) // 마지막 열이 아닌 경우, 열을 오른쪽으로 이동
                    {
                        //dgdSub.SelectedIndex = 0;
                        dgdSub.CurrentCell = new DataGridCellInfo(dgdSub.Items[currRow], dgdSub.Columns[currCol + 1]);
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
                    dgdSub.SelectedIndex = currRow - 1;
                    dgdSub.CurrentCell = new DataGridCellInfo(dgdSub.Items[currRow - 1], dgdSub.Columns[currCol]);
                } // 첫 행
                else if (dgdSub.Items.Count - 1 == currRow)
                {
                    if (0 < currCol) // 첫 열이 아닌 경우, 열을 왼쪽으로 이동
                    {
                        //dgdSub.SelectedIndex = 0;
                        dgdSub.CurrentCell = new DataGridCellInfo(dgdSub.Items[currRow], dgdSub.Columns[currCol - 1]);
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
                    dgdSub.CurrentCell = new DataGridCellInfo(dgdSub.Items[currRow], dgdSub.Columns[currCol - 1]);
                }
                else if (startCol == currCol)
                {
                    if (0 < currRow)
                    {
                        dgdSub.SelectedIndex = currRow - 1;
                        dgdSub.CurrentCell = new DataGridCellInfo(dgdSub.Items[currRow - 1], dgdSub.Columns[endCol]);
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

                    dgdSub.CurrentCell = new DataGridCellInfo(dgdSub.Items[currRow], dgdSub.Columns[currCol + 1]);
                }
                else if (endCol == currCol)
                {
                    if (dgdSub.Items.Count - 1 > currRow)
                    {
                        dgdSub.SelectedIndex = currRow + 1;
                        dgdSub.CurrentCell = new DataGridCellInfo(dgdSub.Items[currRow + 1], dgdSub.Columns[startCol]);
                    }
                    else
                    {
                        //btnSave.Focus();
                    }
                }
            }
        }
        // KeyUp 이벤트
        private void DatagridIn_TextFocus(object sender, KeyEventArgs e)
        {
            // 엔터 → 포커스 = true → cell != null → 해당 텍스트박스가 null이 아니라면 
            // → 해당 텍스트박스가 포커스가 안되있음 SelectAll() or 포커스
            Lib.Instance.DataGridINTextBoxFocus(sender, e);
        }
        // GotFocus 이벤트
        private void DataGridCell_GotFocus(object sender, RoutedEventArgs e)
        {
            if (lblMsg.Visibility == Visibility.Visible)
            {
                DataGridCell cell = sender as DataGridCell;
                cell.IsEditing = true;
            }
        }
        // 2019.08.27 MouseUp 이벤트
        private void DataGridCell_MouseUp(object sender, MouseButtonEventArgs e)
        {
            Lib.Instance.DataGridINTextBoxFocusByMouseUP(sender, e);
        }

        #endregion


        #region 주요 메서드

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

        #region 조회

        private void FillGrid()
        {

            if (dgdMain.Items.Count > 0)
            {
                dgdMain.Items.Clear();
                dgdSub.Items.Clear();
                this.DataContext = null;
            }

            try
            {
                Dictionary<string, object> sqlParameter = new Dictionary<string, object>();

                sqlParameter.Add("MoldID", "");
                sqlParameter.Add("MoldNo", chkMoldNoSrh.IsChecked == true ? txtMoldNoSrh.Text : "");
                sqlParameter.Add("Article", chkArticleSrh.IsChecked == true ? txtArticleSrh.Text : "");

                DataSet ds = DataStore.Instance.ProcedureToDataSet("xp_DvlMold_sInspectBasis", sqlParameter, false);

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

                            var Mold = new Win_dvl_MoldRegularInspectBasis_U_CodeView()
                            {
                                Num = i,

                                MoldInspectBasisID = dr["MoldInspectBasisID"].ToString(),
                                MoldID = dr["MoldID"].ToString(),
                                MoldNo = dr["MoldNo"].ToString(),
                                Article = dr["Article"].ToString(),
                                BuyerArticleNo = dr["BuyerArticleNo"].ToString(),
                                MoldInspectBasisDate = DatePickerFormat(dr["MoldInspectBasisDate"].ToString()),
                                MoldInspectContent = dr["MoldInspectContent"].ToString(),
                                Comments = dr["Comments"].ToString(),
                                InspectCycle = dr["InspectCycle"].ToString(),
                            };
                            dgdMain.Items.Add(Mold);

                        }


                        tbkIndexCount.Text = "검색건수 : " + i + " 건";
                    }
                    else
                    {
                        MessageBox.Show("조회된 결과가 없습니다.");
                        return;
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

        #region 서브그리드 조회

        private void FillGridSub(string strID)
        {
            FTP_Trigger = true;

            if (dgdSub.Items.Count > 0)
            {
                dgdSub.Items.Clear();
            }

            try
            {
                Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                sqlParameter.Clear();

                sqlParameter.Add("MoldInspectBasisID", strID);

                DataSet ds = DataStore.Instance.ProcedureToDataSet("xp_dvlMold_sMoldRegularInspectBasisSub_New", sqlParameter, false);

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

                            var MoldSub = new Win_dvl_MoldRegularInspectBasis_U_CodeViewSub()
                            {
                                Num = i,

                                MoldInspectBasisID = dr["MoldInspectBasisID"].ToString(),
                                MoldSeq = ConvertInt(dr["MoldSeq"].ToString()),
                                MoldInspectItemName = dr["MoldInspectItemName"].ToString(),
                                MoldInspectContent = dr["MoldInspectContent"].ToString(),
                                MoldInspectCheckGbn = dr["MoldInspectCheckGbn"].ToString(),

                                MoldInspectCheckGbn_Name = dr["MoldInspectCheckGbn_Name"].ToString(),
                                MoldInspectCycleGbn = dr["MoldInspectCycleGbn"].ToString(),
                                MoldInspectCycleGbn_Name = dr["MoldInspectCycleGbn_Name"].ToString(),
                                MoldInspectCycleDate = dr["MoldInspectCycleDate"].ToString(),
                                MoldInspectRecordGbn = dr["MoldInspectRecordGbn"].ToString(),

                                MoldInspectRecordGbn_Name = dr["MoldInspectRecordGbn_Name"].ToString(),
                                MoldInspectImageFile = dr["MoldInspectImageFile"].ToString(),
                                MoldInspectComments = dr["MoldInspectComments"].ToString(),
                            };

                            // 버튼 이름
                            if (!MoldSub.MoldInspectImageFile.Trim().Equals(""))
                            {
                                if (FTP_Trigger == true)
                                {
                                    MoldSub.ImageByte = SetImage(MoldSub.MoldInspectImageFile, MoldSub.MoldInspectBasisID);
                                }
                                MoldSub.btnName = "삭제";
                            }
                            else
                            {
                                MoldSub.btnName = "업로드";
                            }

                            dgdSub.Items.Add(MoldSub);

                        }
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

        #endregion // 서브그리드 조회

        #region 저장

        //저장
        private bool SaveData(string strFlag)
        {
            bool flag = false;
            List<Procedure> Prolist = new List<Procedure>();
            List<Dictionary<string, object>> ListParameter = new List<Dictionary<string, object>>();

            if (CheckData())
            {
                try
                {
                    var mold = dgdMain.SelectedItem as Win_dvl_MoldRegularInspectBasis_U_CodeView;

                    string sGetID = txtMoldInspectBasisID.Text.Trim().ToString();
                    Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                    sqlParameter.Clear();

                    sqlParameter.Add("MoldInspectBasisID", sGetID);
                    sqlParameter.Add("MoldID", txtMoldID.Tag != null ? txtMoldID.Tag.ToString() : "");
                    sqlParameter.Add("MoldInspectBasisDate", dtpMoldInspectBasisDate.SelectedDate != null ? dtpMoldInspectBasisDate.SelectedDate.Value.ToString("yyyyMMdd") : "");
                    sqlParameter.Add("MoldInspectContent", txtMoldInspectContent.Text);
                    sqlParameter.Add("Comments", txtComments.Text);
                    sqlParameter.Add("UserID", MainWindow.CurrentUser);
                    sqlParameter.Add("InspectCycle", cboInspectCycle.SelectedValue ?? "");

                    #region 추가
                    if (strFlag.Equals("I"))
                    {
                        Procedure pro1 = new Procedure();
                        pro1.Name = "xp_dvlMold_iuMoldRegularInspectBasis";
                        pro1.OutputUseYN = "Y";
                        pro1.OutputName = "MoldInspectBasisID";
                        pro1.OutputLength = "10";

                        Prolist.Add(pro1);
                        ListParameter.Add(sqlParameter);

                        for (int i =0; i < dgdSub.Items.Count; i++)
                        {
                            sqlParameter = new Dictionary<string, object>();
                            var inspectSub = dgdSub.Items[i] as Win_dvl_MoldRegularInspectBasis_U_CodeViewSub;

                            inspectSub.MoldSeq = i;

                            sqlParameter.Clear();
                            sqlParameter.Add("MoldInspectBasisID", inspectSub.MoldInspectBasisID);
                            sqlParameter.Add("MoldSeq", inspectSub.MoldSeq);
                            sqlParameter.Add("MoldInspectItemName", inspectSub.MoldInspectItemName);
                            sqlParameter.Add("MoldInspectContent", inspectSub.MoldInspectContent);
                            sqlParameter.Add("MoldInspectCheckGbn", inspectSub.MoldInspectCheckGbn);
                            sqlParameter.Add("MoldInspectCycleGbn", inspectSub.MoldInspectCycleGbn);
                            sqlParameter.Add("MoldInspectCycleDate", inspectSub.MoldInspectCycleDate);
                            sqlParameter.Add("MoldInspectRecordGbn", inspectSub.MoldInspectRecordGbn);
                            sqlParameter.Add("MoldInspectImageFile", inspectSub.MoldInspectImageFile);

                            Procedure pro2 = new Procedure();
                            pro2.Name = "xp_dvlMold_iuMoldRegularInspectBasisSub";
                            pro2.OutputUseYN = "Y";
                            pro2.OutputName = "MoldSeq";
                            pro2.OutputLength = "5";

                            Prolist.Add(pro2);
                            ListParameter.Add(sqlParameter);

                        }

                        List<KeyValue> list_Result = new List<KeyValue>();
                        list_Result = DataStore.Instance.ExecuteAllProcedureOutputGetCS(Prolist, ListParameter);
                        

                        if (list_Result[0].key.ToLower() == "success")
                        {
                            list_Result.RemoveAt(0);
                            for (int i = 0; i < list_Result.Count; i++)
                            {
                                KeyValue kv = list_Result[i];
                                if (kv.key == "MoldInspectBasisID")
                                {
                                    sGetID = kv.value;
                                    flag = true;
                                }
                            }

                        }
                        else
                        {
                            flag = false;
                        }
                    }

                    #endregion


                    #region 수정
                    
                    else if (strFlag.Equals("U"))
                    {
                        Procedure pro1 = new Procedure();
                        pro1.Name = "xp_DvlMold_uMoldRegularInspectBasis";
                        pro1.OutputUseYN = "N";
                        pro1.OutputName = "MoldInspectBasisID";
                        pro1.OutputLength = "10";

                        Prolist.Add(pro1);
                        ListParameter.Add(sqlParameter);

                        for (int i = 0; i < dgdSub.Items.Count; i++)
                        {
                            sqlParameter = new Dictionary<string, object>();
                            var inspectSub = dgdSub.Items[i] as Win_dvl_MoldRegularInspectBasis_U_CodeViewSub;

                            inspectSub.MoldSeq = i;

                            sqlParameter.Clear();
                            sqlParameter.Add("MoldInspectBasisID", txtMoldInspectBasisID.Text);
                            sqlParameter.Add("MoldSeq", inspectSub.MoldSeq);
                            sqlParameter.Add("MoldInspectItemName", inspectSub.MoldInspectItemName);
                            sqlParameter.Add("MoldInspectContent", inspectSub.MoldInspectContent);
                            sqlParameter.Add("MoldInspectCheckGbn", inspectSub.MoldInspectCheckGbn);
                            sqlParameter.Add("MoldInspectCycleGbn", inspectSub.MoldInspectCycleGbn);
                            sqlParameter.Add("MoldInspectCycleDate", inspectSub.MoldInspectCycleDate);
                            sqlParameter.Add("MoldInspectRecordGbn", inspectSub.MoldInspectRecordGbn);
                            sqlParameter.Add("MoldInspectImageFile", inspectSub.MoldInspectImageFile);

                            Procedure pro2 = new Procedure();
                            pro2.Name = "xp_dvlMold_iuMoldRegularInspectBasisSub";
                            pro2.OutputUseYN = "Y";
                            pro2.OutputName = "MoldSeq";
                            pro2.OutputLength = "5";

                            Prolist.Add(pro2);
                            ListParameter.Add(sqlParameter);

                        }

                        string[] Confirm = new string[2];
                        Confirm = DataStore.Instance.ExecuteAllProcedureOutputNew(Prolist, ListParameter);
                        if (Confirm[0] != "success")
                        {
                            flag = false;
                        }
                        else
                        {
                            flag = true;
                        }

                    }

                    #endregion

                    // 파일을 올리자 : GetKey != "" 라면 파일을 올려보자
                    if (!sGetID.Trim().Equals(""))
                    {
                        //삭제할 사진이 있을 경우
                        if (deleteListFtpFile.Count > 0)
                        {
                            foreach (string[] str in deleteListFtpFile)
                            {
                                FTP_RemoveFile(sGetID + "/" + str[0]);
                            }
                        }

                        //추가한 사진이 있을 때
                        if (listFtpFile.Count > 0)
                        {
                            FTP_Save_File(listFtpFile, sGetID);
                            //AttachFileUpdate(GetKey);
                        }

                        //복사추가 했을 떄 
                        if (lstFtpFilePath.Count > 0)
                        {
                            FTP_Save_FileByFtpServerFilePath(lstFtpFilePath, sGetID);
                        }
                    }

                    // 파일 List 비워주기
                    listFtpFile.Clear();
                    deleteListFtpFile.Clear();
                    lstFtpFilePath.Clear();
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

        // 유효성검사
        private bool CheckData()
        {
            bool flag = true;

            if(cboInspectCycle.SelectedValue == null)
            {
                MessageBox.Show("점검주기구분을 선택해주세요");
            }

            if (txtMoldID.Tag == null 
                || txtMoldID.Text.Trim().Equals(""))
            {
                MessageBox.Show("금형 Lot 번호를 선택해주세요.");
                flag = false;
                return flag;
            }

            if (txtMoldInspectContent.Text.Trim().Equals(""))
            {
                MessageBox.Show("개정내용을 입력해주세요.");
                flag = false;
                return flag;
            }

            return flag;
        }

        #endregion // 유효성 검사

        #region 삭제

        private bool Delete(string strID)
        {
            bool flag = false;
            List<Procedure> Prolist = new List<Procedure>();
            List<Dictionary<string, object>> ListParameter = new List<Dictionary<string, object>>();

            try
            {
                Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                sqlParameter.Clear();

                sqlParameter.Add("MoldInspectBasisID", strID);

                Procedure pro2 = new Procedure();
                pro2.Name = "xp_DvlMold_dMolRegularInspectBasis";
                pro2.OutputUseYN = "N";
                pro2.OutputName = "sArticleID";
                pro2.OutputLength = "10";

                Prolist.Add(pro2);
                ListParameter.Add(sqlParameter);

                string[] Confirm = new string[2];
                Confirm = DataStore.Instance.ExecuteAllProcedureOutputNew(Prolist, ListParameter);
                if (Confirm[0] != "success")
                {
                    MessageBox.Show("[저장실패]\r\n" + Confirm[1].ToString());
                    flag = false;
                    return flag;
                }
                else
                {
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


        #endregion // 삭제

        #endregion // 주요 메서드

        #region 기타 메서드

        // 천마리 콤마, 소수점 버리기
        private string stringFormatN0(object obj)
        {
            return string.Format("{0:N0}", obj);
        }

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

        // Int로 변환 가능한지 체크 이벤트
        private bool CheckConvertInt(string str)
        {
            bool flag = false;
            double chkInt = 0;

            if (!str.Trim().Equals(""))
            {
                str = str.Replace(",", "");
                str = str.Replace(".", "");

                if (Double.TryParse(str, out chkInt) == true)
                {
                    flag = true;
                }
            }

            return flag;
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
                str = str.Replace(",", "");

                if (Double.TryParse(str, out chkDouble) == true)
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

        // 월만 가져오기 > 앞에 0 없애기
        private string getDateMonth(string str)
        {
            string month = "";

            if (!str.Trim().Equals(""))
            {
                str = str.Trim().Replace("-", "").Replace(".", "");

                if (str.Length == 8)
                {
                    month = str.Substring(4, 2);

                    if (month.Substring(0, 1).Equals("0"))
                    {
                        month = month.Substring(1, 1);
                    }
                }
            }

            return month;
        }

        // 일만 가져오기 > 앞에 0 없애기
        private string getDateDay(string str)
        {
            string day = "";

            if (!str.Trim().Equals(""))
            {
                str = str.Trim().Replace("-", "").Replace(".", "");

                if (str.Length == 8)
                {
                    day = str.Substring(6, 2);

                    if (day.Substring(0, 1).Equals("0"))
                    {
                        day = day.Substring(1, 1);
                    }
                }
            }

            return day;
        }



        #endregion // 기타 메서드

        // 엔터 → 다음 셀을 위한 텍스트박스 포커스 이벤트
        private void txtBox_Loaded(object sender, RoutedEventArgs e)
        {
            if (editing == true)
            {
                (sender as TextBox).Focus();

                editing = false;
            }
        }

        private void img_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            try
            {
                if (e.ClickCount == 2)
                {

                    System.Windows.Controls.Image senderImg = sender as System.Windows.Controls.Image;
                    var MoldSub = senderImg.DataContext as Win_dvl_MoldRegularInspectBasis_U_CodeViewSub;

                    if (MoldSub != null
                        && MoldSub.MoldInspectImageFile != null
                        && !MoldSub.MoldInspectImageFile.Trim().Equals(""))
                    {

                        string str_path = string.Empty;
                        str_path = FTP_ADDRESS + '/' + MoldSub.MoldInspectBasisID;
                        _ftp = new FTP_EX(str_path, FTP_ID, FTP_PASS);

                        string str_remotepath = string.Empty;
                        string str_localpath = string.Empty;

                        str_remotepath = "/" + MoldSub.MoldInspectImageFile;
                        str_localpath = LOCAL_DOWN_PATH + "\\" + MoldSub.MoldInspectImageFile;

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

                        ProcessStartInfo proc = new ProcessStartInfo(str_localpath);
                        proc.UseShellExecute = true;
                        Process.Start(proc);
                    }
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                return;
            }


        }

        private void cboMoldInspectCheckGbn_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            rowNum = dgdSub.SelectedIndex;
            var winMoldInspectSub = dgdSub.CurrentItem as Win_dvl_MoldRegularInspectBasis_U_CodeViewSub;

            ComboBox cboMoldInspectCheckGbn = (ComboBox)sender;

            if (winMoldInspectSub != null)
            {
                winMoldInspectSub = dgdSub.Items[rowNum] as Win_dvl_MoldRegularInspectBasis_U_CodeViewSub;
            }
            
            if (cboMoldInspectCheckGbn.SelectedValue != null && !cboMoldInspectCheckGbn.SelectedValue.ToString().Equals(""))
            {
                var theView = cboMoldInspectCheckGbn.SelectedItem as CodeView;
                if (theView != null)
                {
                    winMoldInspectSub.MoldInspectCheckGbn = theView.code_id;
                    winMoldInspectSub.MoldInspectCheckGbn_Name = theView.code_name;
                }

                sender = cboMoldInspectCheckGbn;
            }
        }

        private void cboMoldInspectCycleGbn_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            rowNum = dgdSub.SelectedIndex;
            var winMoldInspectSub = dgdSub.CurrentItem as Win_dvl_MoldRegularInspectBasis_U_CodeViewSub;

            ComboBox cboMoldInspectCycleGbn = (ComboBox)sender;

            if (winMoldInspectSub != null)
            {
                winMoldInspectSub = dgdSub.Items[rowNum] as Win_dvl_MoldRegularInspectBasis_U_CodeViewSub;
            }
            
            if (cboMoldInspectCycleGbn.SelectedValue != null && !cboMoldInspectCycleGbn.SelectedValue.ToString().Equals(""))
            {
                var theView = cboMoldInspectCycleGbn.SelectedItem as CodeView;
                if (theView != null)
                {
                    winMoldInspectSub.MoldInspectCycleGbn = theView.code_id;
                    winMoldInspectSub.MoldInspectCycleGbn_Name = theView.code_name;
                }

                sender = cboMoldInspectCycleGbn;
            }
        }

        private void cboMoldInspectRecordGbn_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            rowNum = dgdSub.SelectedIndex;
            var winMoldInspectSub = dgdSub.CurrentItem as Win_dvl_MoldRegularInspectBasis_U_CodeViewSub;

            ComboBox cboMoldInspectRecordGbn = (ComboBox)sender;

            if (winMoldInspectSub != null)
            {
                winMoldInspectSub = dgdSub.Items[rowNum] as Win_dvl_MoldRegularInspectBasis_U_CodeViewSub;
            }
            
            if (cboMoldInspectRecordGbn.SelectedValue != null && !cboMoldInspectRecordGbn.SelectedValue.ToString().Equals(""))
            {
                var theView = cboMoldInspectRecordGbn.SelectedItem as CodeView;
                if (theView != null)
                {
                    winMoldInspectSub.MoldInspectRecordGbn = theView.code_id;
                    winMoldInspectSub.MoldInspectRecordGbn_Name = theView.code_name;
                }

                sender = cboMoldInspectRecordGbn;
            }
        }

        private void txt_PreviewTextInput(object sender, TextCompositionEventArgs e)
        {
            Lib.Instance.CheckIsNumericOnly((TextBox)sender, e);
        }
    }

    class Win_dvl_MoldRegularInspectBasis_U_CodeView : BaseView
    {
        public int Num { get; set; }

        public string MoldInspectBasisID { get; set; } // 기준번호
        public string MoldID { get; set; } // 철형번호
        public string MoldName { get; set; } // 철형명
        public string MoldNo { get; set; } // 철형_관리번호
        public string Article { get; set; } // 철형_관리번호
        public string BuyerArticleNo { get; set; } // 철형_관리번호
        public string Mold_Comments { get; set; } // 철형 비고

        public string MoldInspectBasisDate { get; set; } // 개정일자
        public string MoldInspectContent { get; set; } // 개정내용
        public string Comments { get; set; } // 비고
        public string Compliance { get; set; } // 점검시 준수사항
        public string InspectCycle { get; set; } //
    }

    class Win_dvl_MoldRegularInspectBasis_U_CodeViewSub
    {
        public int Num { get; set; }

        public string MoldInspectBasisID { get; set; } // 기준번호
        public int MoldSeq { get; set; } // 시퀀스
        public string MoldInspectItemName { get; set; } // 점검항목
        public string MoldInspectContent { get; set; } // 점검내용

        public string MoldInspectCheckGbn { get; set; } // 확인방법
        public string MoldInspectCheckGbn_Name { get; set; } // 확인방법

        public string MoldInspectCycleGbn { get; set; } // 주기
        public string MoldInspectCycleGbn_Name { get; set; } // 주기   

        public string MoldInspectCycleDate { get; set; } // 특정 월 일

        public string MoldInspectRecordGbn { get; set; } // 기록구분
        public string MoldInspectRecordGbn_Name { get; set; } // 기록구분

        public string MoldInspectComments { get; set; } // 비고
        public string MoldInspectImageFile { get; set; } // 이미지

        public BitmapImage ImageByte { get; set; }
        public string ImagePath { get; set; }
        public bool ImageFlag { get; set; }
        public string btnName { get; set; } // 비고
    }
}

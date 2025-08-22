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
using WizMes_WellMade.PopUP;

namespace WizMes_WellMade
{
    /// <summary>
    /// Win_Qul_Drawing_U.xaml에 대한 상호 작용 논리
    /// </summary>
    public partial class Win_Qul_Drawing_U : UserControl
    {
        #region 변수 선언 및 로드

        Lib lib = new Lib();
        PlusFinder pf = new PlusFinder();
        int numRowCount = 0;
        string strFlag = string.Empty;

        string strAttPath1 = string.Empty;
        string strAttPath2 = string.Empty;
        string strAttPath3 = string.Empty;

        // FTP 활용모음.
        private FTP_EX _ftp = null;
        List<string[]> listFtpFile = new List<string[]>();


        string FullPath1 = string.Empty;
        string FullPath2 = string.Empty;
        string FullPath3 = string.Empty;
        //string FTP_ADDRESS = "ftp://wizis.iptime.org/ImageData/Draw";
        string FTP_ADDRESS = "ftp://" + LoadINI.FileSvr + ":" + LoadINI.FTPPort + LoadINI.FtpImagePath + "/Draw";
        //string FTP_ADDRESS = "ftp://211.228.238.227:25000/ImageData/Draw";
        //string FTP_ADDRESS = "ftp://HKserver:210/ImageData/Draw";
        private const string FTP_ID = "wizuser";
        private const string FTP_PASS = "wiz9999";
        private const string LOCAL_DOWN_PATH = "C:\\Temp";

        List<string> lstCompareValue = new List<string>();
        Dictionary<string, object> dicCompare = new Dictionary<string, object>();

        public Win_Qul_Drawing_U()
        {
            InitializeComponent();
        }

        private void UserControl_Loaded(object sender, RoutedEventArgs e)
        {
            chkDrawCreateDateSrh.IsChecked = true;
            dtpSDate.SelectedDate = DateTime.Today;
            dtpEDate.SelectedDate = DateTime.Today;
            dtpDrawAcptDate.SelectedDate = DateTime.Today;
            dtpDrawCreateDate.SelectedDate = DateTime.Today;

            SetCombo();

            lib.UiLoading(sender);

            ControlVisibleAndEnable_SC();
        }

        private void SetCombo()
        {
            ObservableCollection<CodeView> ovcDrawAuthor = ComboBoxUtil.Instance.Gf_DB_CM_GetComCodeDataset(null, "DAWAUTH", "Y", "");
            this.cboDrawAuthor.ItemsSource = ovcDrawAuthor;
            this.cboDrawAuthor.DisplayMemberPath = "code_name";
            this.cboDrawAuthor.SelectedValuePath = "code_id";
        }

        #endregion

        #region 상단 체크 이벤트 및 플러스파인더

        //도면생성일(상단) 라벨
        private void lblDrawCreateDateSrh_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            if (chkDrawCreateDateSrh.IsChecked == true) { chkDrawCreateDateSrh.IsChecked = false; }
            else { chkDrawCreateDateSrh.IsChecked = true; }
        }

        //도면생성일(상단)
        private void chkDrawCreateDateSrh_Checked(object sender, RoutedEventArgs e)
        {
            dtpSDate.IsEnabled = true;
            dtpEDate.IsEnabled = true;
        }

        //도면생성일(상단)
        private void chkDrawCreateDateSrh_Unchecked(object sender, RoutedEventArgs e)
        {
            dtpSDate.IsEnabled = false;
            dtpEDate.IsEnabled = false;
        }

        //품명(상단) 라벨
        private void lblArticelSrh_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            if (chkArticleSrh.IsChecked == true) { chkArticleSrh.IsChecked = false; }
            else { chkArticleSrh.IsChecked = true; }
        }

        //품명(상단) 체크박스
        private void chkArticleSrh_Checked(object sender, RoutedEventArgs e)
        {
            txtArticelSrh.IsEnabled = true;
            btnPfArticelSrh.IsEnabled = true;
            txtArticelSrh.Focus();
        }

        //품명(상단) 체크박스
        private void chkArticleSrh_Unchecked(object sender, RoutedEventArgs e)
        {
            txtArticelSrh.IsEnabled = false;
            btnPfArticelSrh.IsEnabled = false;
        }

        //품명(상단) textbox event
        private void txtArticelSrh_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                pf.ReturnCode(txtArticelSrh, 81, txtArticelSrh.Text);
            }
        }

        //품명(상단) plusfinder event
        private void btnPfArticelSrh_Click(object sender, RoutedEventArgs e)
        {
            pf.ReturnCode(txtArticelSrh, 81, txtArticelSrh.Text);
        }

        //도면번호(상단) 라벨
        private void lblDrawNoSrh_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            if (chkDrawNoSrh.IsChecked == true) { chkDrawNoSrh.IsChecked = false; }
            else { chkDrawNoSrh.IsChecked = true; }
        }

        //도면번호(상단) 체크박스
        private void chkDrawNoSrh_Checked(object sender, RoutedEventArgs e)
        {
            txtDrawNoSrh.IsEnabled = true;
            txtDrawNoSrh.Focus();
        }

        //도면번호(상단) 체크박스
        private void chkDrawNoSrh_Unchecked(object sender, RoutedEventArgs e)
        {
            txtDrawNoSrh.IsEnabled = false;
        }

        #endregion

        #region 상단 우측 버튼 이벤트

        //추가,수정 시 동작 모음
        private void ControlVisibleAndEnable_AU()
        {
            lib.UiButtonEnableChange_SCControl(this);
            dgdDraw.IsHitTestVisible = false;
            grdOne.IsHitTestVisible = true;

            btnAddAttFile1.IsHitTestVisible = true;
            btnAddAttFile2.IsHitTestVisible = true;
            btnAddAttFile3.IsHitTestVisible = true;

            txtAttFile1.IsHitTestVisible = true;
            txtAttFile2.IsHitTestVisible = true;
            txtAttFile3.IsHitTestVisible = true;

            btnDelAttFile1.IsHitTestVisible = true;
            btnDelAttFile2.IsHitTestVisible = true;
            btnDelAttFile3.IsHitTestVisible = true;

            btnDownAttFile1.IsEnabled = false;
            btnDownAttFile1.IsHitTestVisible = false;
            btnDownAttFile2.IsEnabled = false;
            btnDownAttFile2.IsHitTestVisible = false;
            btnDownAttFile3.IsEnabled = false;
            btnDownAttFile3.IsHitTestVisible = false;
        }

        //저장,취소 시 동작 모음
        private void ControlVisibleAndEnable_SC()
        {
            lib.UiButtonEnableChange_IUControl(this);

            dgdDraw.IsHitTestVisible = true;
            grdOne.IsHitTestVisible = false;

            btnAddAttFile1.IsHitTestVisible = false;
            btnAddAttFile2.IsHitTestVisible = false;
            btnAddAttFile3.IsHitTestVisible = false;

            txtAttFile1.IsHitTestVisible = false;
            txtAttFile2.IsHitTestVisible = false;
            txtAttFile3.IsHitTestVisible = false;

            btnDelAttFile1.IsHitTestVisible = false;
            btnDelAttFile2.IsHitTestVisible = false;
            btnDelAttFile3.IsHitTestVisible = false;

            btnDownAttFile1.IsEnabled = true;
            btnDownAttFile1.IsHitTestVisible = true;
            btnDownAttFile2.IsEnabled = true;
            btnDownAttFile2.IsHitTestVisible = true;
            btnDownAttFile3.IsEnabled = true;
            btnDownAttFile3.IsHitTestVisible = true;

        }

        //추가
        private void btnAdd_Click(object sender, RoutedEventArgs e)
        {
            if (dgdDraw.Items.Count > 0 && dgdDraw.SelectedIndex > -1)
            {
                numRowCount = dgdDraw.SelectedIndex;    //취소 시 대비
            }

            ControlVisibleAndEnable_AU();
            tbkMsg.Text = "자료 입력(추가) 중";
            strFlag = "I";
            this.DataContext = null;

            dtpDrawAcptDate.SelectedDate = DateTime.Today;
            dtpDrawCreateDate.SelectedDate = DateTime.Today;

            txtArticle.Focus();
            //txtBuyerModel.Focus();
        }

        //수정
        private void btnUpdate_Click(object sender, RoutedEventArgs e)
        {
            var WinDraw = dgdDraw.SelectedItem as Win_Qul_Drawing_U_CodeView;

            if (WinDraw == null)
            {
                MessageBox.Show("수정할 자료를 선택하고 눌러주십시오.");
            }
            else
            {
                numRowCount = dgdDraw.SelectedIndex;
                ControlVisibleAndEnable_AU();
                tbkMsg.Text = "자료 입력(수정) 중";
                strFlag = "U";

                txtArticle.Focus();
                //txtBuyerModel.Focus();
            }
        }

        //삭제
        private void btnDelete_Click(object sender, RoutedEventArgs e)
        {
            var winDraw = dgdDraw.SelectedItem as Win_Qul_Drawing_U_CodeView;

            if (winDraw == null)
            {
                MessageBox.Show("삭제할 데이터가 지정되지 않았습니다. 삭제데이터를 지정하고 눌러주세요");
                return;
            }
            else
            {
                if (dgdDraw.SelectedIndex == dgdDraw.Items.Count - 1)
                {
                    numRowCount = dgdDraw.SelectedIndex - 1;
                }
                else
                {
                    numRowCount = dgdDraw.SelectedIndex;
                }

                if (MessageBox.Show("선택하신 항목을 삭제하시겠습니까?", "삭제 전 확인", MessageBoxButton.YesNo) == MessageBoxResult.Yes)
                {
                    //if (!txtAttFile1.Text.Equals(""))
                    //{
                    //    if (DetectFtpFile(winDraw.DrawID))
                    //    {
                    //        FTP_UploadFile_File_Delete(winDraw.DrawID, txtAttFile1.Text);
                    //    }
                    //}
                    //if (!txtAttFile2.Text.Equals(""))
                    //{
                    //    if (DetectFtpFile(winDraw.DrawID))
                    //    {
                    //        FTP_UploadFile_File_Delete(winDraw.DrawID, txtAttFile2.Text);
                    //    }
                    //}
                    //if (!txtAttFile3.Text.Equals(""))
                    //{
                    //    if (DetectFtpFile(winDraw.DrawID))
                    //    {
                    //        FTP_UploadFile_File_Delete(winDraw.DrawID, txtAttFile3.Text);
                    //    }
                    //}

                    //FTP_UploadFile_Path_Delete(winDraw.DrawID);

                    if (DeleteData(winDraw.DrawID))
                    {
                        re_Search(numRowCount);
                    }
                }
            }
        }

        //닫기
        private void btnClose_Click(object sender, RoutedEventArgs e)
        {
            Lib.Instance.ChildMenuClose(this.ToString());
        }

        //조회
        private void btnSearch_Click(object sender, RoutedEventArgs e)
        {
            //검색버튼 비활성화
            btnSearch.IsEnabled = false;

            Dispatcher.BeginInvoke(new Action(() =>

            {
                Thread.Sleep(2000);

                //로직
                numRowCount = 0;
                re_Search(numRowCount);

                ControlVisibleAndEnable_SC();

                dgdDraw.Focus();

            }), System.Windows.Threading.DispatcherPriority.Background);



            Dispatcher.BeginInvoke(new Action(() =>

            {
                btnSearch.IsEnabled = true;

            }), System.Windows.Threading.DispatcherPriority.Background);

        }

        //저장
        private void btnSave_Click(object sender, RoutedEventArgs e)
        {
            if (strFlag.Equals("I"))
            {
                if (SaveData("", strFlag))
                {
                    ControlVisibleAndEnable_SC();
                    re_Search(dgdDraw.Items.Count - 1);
                    dgdDraw.Focus();

                    strFlag = string.Empty;
                    listFtpFile.Clear();
                }
            }
            else
            {
                if (SaveData(txtDrawID.Text, strFlag))
                {
                    ControlVisibleAndEnable_SC();
                    re_Search(numRowCount);
                    dgdDraw.Focus();


                    strFlag = string.Empty;
                    listFtpFile.Clear();

                }
            }
        }

        //취소
        private void btnCancel_Click(object sender, RoutedEventArgs e)
        {
            InputClear();
            ControlVisibleAndEnable_SC();
            re_Search(numRowCount);
            dgdDraw.Focus();

            listFtpFile.Clear();
        }

        //입력 데이터 클리어
        private void InputClear()
        {
            //foreach (Control child in this.grdOne.Children)
            //{
            //    if (child.GetType() == typeof(TextBox))
            //        ((TextBox)child).Clear();
            //    else if (child.GetType() == typeof(ComboBox))
            //        ((ComboBox)child).SelectedIndex = -1;
            //}

            this.txtDrawID.Clear();
            this.txtDrawNo.Clear();
            this.txtArticleID.Clear();
            this.txtArticle.Clear();
            this.txtBuyerArticle.Clear();
            this.txtBuyerModel.Clear();
            this.chkDrawCreateDate.IsChecked = new bool?(false);
            this.dtpDrawCreateDate.SelectedDate = null;
            this.txtDrawCreateMan.Clear();
            this.txtECONo.Clear();
            this.txtDrawDvlyPlace.Clear();
            this.cboDrawAuthor.SelectedIndex = -1;
            this.chkDrawAcptDate.IsChecked = new bool?(false);
            this.dtpDrawAcptDate.SelectedDate = null;
            this.txtDrawAcptMan.Clear();
            this.txtComments.Clear();

            this.chkDrawCreateDate.IsChecked = new bool?(false);
            this.chkDrawAcptDate.IsChecked = new bool?(false);
            this.txtDrawAcptMan.Clear();
            this.txtComments.Clear();
            this.txtAttFile1.Clear();
            this.txtAttFile2.Clear();
            this.txtAttFile3.Clear();
        }

        //엑셀
        private void btnExcel_Click(object sender, RoutedEventArgs e)
        {
            DataTable dt = null;
            string Name = string.Empty;
            Lib lib2 = new Lib();

            string[] dgdStr = new string[2];
            dgdStr[0] = "도면등록";
            dgdStr[1] = dgdDraw.Name;

            ExportExcelxaml ExpExc = new ExportExcelxaml(dgdStr);
            ExpExc.ShowDialog();

            if (ExpExc.DialogResult.HasValue)
            {
                if (ExpExc.choice.Equals(dgdDraw.Name))
                {
                    if (ExpExc.Check.Equals("Y"))
                        dt = lib2.DataGridToDTinHidden(dgdDraw);
                    else
                        dt = lib2.DataGirdToDataTable(dgdDraw);

                    Name = dgdDraw.Name;
                    if (lib2.GenerateExcel(dt, Name))
                    {
                        lib2.excel.Visible = true;
                        lib2.ReleaseExcelObject(lib2.excel);
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
            lib2 = null;
        }

        private void re_Search(int selectIndex)
        {
            if (dgdDraw.Items.Count > 0)
            {
                dgdDraw.Items.Clear();
            }

            FillGrid();

            if (dgdDraw.Items.Count > 0)
            {
                if (lstCompareValue.Count > 0)
                {
                    dgdDraw.SelectedIndex = lib.reTrunIndex(dgdDraw, lstCompareValue[0]);
                }
                else
                {
                    dgdDraw.SelectedIndex = selectIndex; ;
                }
            }
            else
            {
                InputClear();
            }

            dicCompare.Clear();
            lstCompareValue.Clear();
        }

        #endregion

        #region 우측중단 이벤트 및 enter focus move

        //품명(중단) textbox
        private void txtArticle_KeyDown(object sender, KeyEventArgs e)
        {
            try
            {
                if (e.Key == Key.Enter)
                {
                    pf.ReturnCode(txtArticle, 81, txtArticle.Text);

                    if (txtArticle.Tag != null && !txtArticle.Tag.ToString().Equals(""))
                    {
                        txtArticleID.Text = txtArticle.Tag.ToString();  //앞의 창에 코드 입력

                        //if (cboBuyerArticle.ItemsSource != null)
                        //{
                        //    cboBuyerArticle.ItemsSource = null;
                        //}

                        SetBuyerArticelNo(txtArticle.Tag.ToString());

                        //도면번호 찾아오기는... 도면 등록하는 화면에서 필요가 없어보이는데..?
                        //SetDrawNoByArticleID(txtArticle.Tag.ToString());
                    }

                    //cboBuyerArticle.Focus();
                    //cboBuyerArticle.IsDropDownOpen = true;
                }
            }
            catch (Exception ex)
            {
                //MessageBox.Show("오류 발생, 오류 내용 : " + ex.ToString());
            }
            finally
            {
                DataStore.Instance.CloseConnection();
            }
        }

        //품명(중단) plusfinder
        private void btnPfArticle_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                pf.ReturnCode(txtArticle, 81, txtArticle.Text);

                if (txtArticle.Tag != null && !txtArticle.Tag.ToString().Equals(""))
                {
                    txtArticleID.Text = txtArticle.Tag.ToString();  //앞의 창에 코드 입력

                    //if (cboBuyerArticle.ItemsSource != null)
                    //{
                    //    cboBuyerArticle.ItemsSource = null;
                    //}

                    SetBuyerArticelNo(txtArticle.Tag.ToString());

                    //도면번호 찾아오기는... 도면 등록하는 화면에서 필요가 없어보이는데..?
                    //SetDrawNoByArticleID(txtArticle.Tag.ToString());
                }

                //cboBuyerArticle.Focus();
                //cboBuyerArticle.IsDropDownOpen = true;
            }
            catch (Exception ex)
            {
                //MessageBox.Show("오류 발생, 오류 내용 : " + ex.ToString());
            }
            finally
            {
                DataStore.Instance.CloseConnection();
            }
        }

        //품명 넘버 가져오기
        private void SetBuyerArticelNo(string strArticleID) //품명을 뿌려야하니까 수정 2020.03.19, 장가빈
        {
            DataTable dt = Procedure.Instance.GetArticleData(strArticleID);

            if (dt.Rows.Count > 0)
            {
                if (dt.Rows[0]["Article"] != null &&
                    !dt.Rows[0]["Article"].ToString().Trim().Equals(string.Empty))
                {
                    txtBuyerArticle.Text = dt.Rows[0]["Article"].ToString();
                }
            }

            //ObservableCollection<CodeView> ovcBuyerArticleNo = ComboBoxUtil.Instance.GetBuyerArticleNo_SetComboBox(strArticleID);
            //this.cboBuyerArticle.ItemsSource = ovcBuyerArticleNo;
            //this.cboBuyerArticle.DisplayMemberPath = "code_name";
            //this.cboBuyerArticle.SelectedValuePath = "code_id";

            //ObservableCollection<CodeView> ovcModel = ComboBoxUtil.Instance.GetModelID_SetComboBox(strArticleID);
            //this.cboBuyerModel.ItemsSource = ovcModel;
            //this.cboBuyerModel.DisplayMemberPath = "code_name";
            //this.cboBuyerModel.SelectedValuePath = "code_id";
        }

        private void SetDrawNoByArticleID(string strArticleID)
        {
            txtDrawNo.Text = ComboBoxUtil.Instance.Get_DrawNo(strArticleID);
        }

        //품번
        private void cboBuyerArticle_DropDownClosed(object sender, EventArgs e)
        {
            txtBuyerModel.Focus();
        }

        //차종
        private void txtBuyerModel_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                e.Handled = true;
                btnPfModel_Click(null, null);
                chkDrawCreateDate.Focus();
            }
        }

        //차종
        private void btnPfModel_Click(object sender, RoutedEventArgs e)
        {
            MainWindow.pf.ReturnCode(txtBuyerModel, (int)Defind_CodeFind.DCF_BUYERMODEL, "");
        }

        //도면생성일(중단) 라벨
        private void lblDrawCreateDate_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            if (chkDrawCreateDate.IsChecked == true) { chkDrawCreateDate.IsChecked = false; }
            else { chkDrawCreateDate.IsChecked = true; }
        }

        //도면생성일(중단)
        private void chkDrawCreateDate_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                e.Handled = true;
                if (MessageBox.Show("도면생성일을 수정하시겠습니까?", "", MessageBoxButton.YesNo) == MessageBoxResult.Yes)
                {
                    chkDrawCreateDate.IsChecked = true;
                    dtpDrawCreateDate.Focus();
                }
                else
                {
                    chkDrawCreateDate.IsChecked = false;
                    txtDrawCreateMan.Focus();
                }
            }
        }

        //도면생성일(중단) 체크박스
        private void chkDrawCreateDate_Checked(object sender, RoutedEventArgs e)
        {
            dtpDrawCreateDate.IsEnabled = true;
            dtpDrawCreateDate.Focus();
        }

        //도면생성일(중단) 체크박스
        private void chkDrawCreateDate_Unchecked(object sender, RoutedEventArgs e)
        {
            dtpDrawCreateDate.IsEnabled = false;
            txtDrawCreateMan.Focus();
        }

        //도면생성일(중단)
        private void dtpDrawCreateDate_PreviewKeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                dtpDrawCreateDate.IsDropDownOpen = true;
            }
        }

        //도면생성일(중단)
        private void dtpDrawCreateDate_CalendarClosed(object sender, RoutedEventArgs e)
        {
            txtDrawCreateMan.Focus();
        }

        //도면작성자
        private void txtDrawCreateMan_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                txtECONo.Focus();
            }
        }

        //EO 번호
        private void txtECONo_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                txtDrawDvlyPlace.Focus();
            }
        }

        //도면발송처
        private void txtDrawDvlyPlace_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                cboDrawAuthor.Focus();
                cboDrawAuthor.IsDropDownOpen = true;
            }
        }

        //도면권한
        private void cboDrawAuthor_DropDownClosed(object sender, EventArgs e)
        {
            chkDrawAcptDate.Focus();
        }

        //접수일(중단) 라벨
        private void lblDrawAcptDate_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            if (chkDrawAcptDate.IsChecked == true) { chkDrawAcptDate.IsChecked = false; }
            else { chkDrawAcptDate.IsChecked = true; }
        }

        //접수일(중단)
        private void chkDrawAcptDate_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                e.Handled = true;
                if (MessageBox.Show("접수일을 수정하시겠습니까?", "", MessageBoxButton.YesNo) == MessageBoxResult.Yes)
                {
                    chkDrawAcptDate.IsChecked = true;
                    dtpDrawAcptDate.Focus();
                }
                else
                {
                    chkDrawAcptDate.IsChecked = false;
                    txtDrawAcptMan.Focus();
                }
            }
        }

        //접수일(중단) 체크박스
        private void chkDrawAcptDate_Checked(object sender, RoutedEventArgs e)
        {
            if (dtpDrawAcptDate == null)
                return;

            dtpDrawAcptDate.IsEnabled = true;
            dtpDrawAcptDate.Focus();
        }

        //접수일(중단) 체크박스
        private void chkDrawAcptDate_Unchecked(object sender, RoutedEventArgs e)
        {
            if (dtpDrawAcptDate == null)
                return;

            dtpDrawAcptDate.IsEnabled = false;
            txtDrawAcptMan.Focus();
        }

        //접수일(중단)
        private void dtpDrawAcptDate_PreviewKeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                dtpDrawAcptDate.IsDropDownOpen = true;
            }
        }

        //접수일(중단)
        private void dtpDrawAcptDate_CalendarClosed(object sender, RoutedEventArgs e)
        {
            txtDrawAcptMan.Focus();
        }

        //도면접수자
        private void txtDrawAcptMan_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                txtComments.Focus();
            }
        }

        #endregion

        #region 우측하단 그룹박스 버튼 이벤트

        //파일 첨부
        private void btnFileEnroll_Click(object sender, RoutedEventArgs e)
        {
            // (버튼)sender 마다 tag를 달자.
            string ClickPoint = ((Button)sender).Tag.ToString();
            string[] strTemp = null;
            Microsoft.Win32.OpenFileDialog OFdlg = new Microsoft.Win32.OpenFileDialog();

            OFdlg.DefaultExt = ".jpg";
            OFdlg.Filter = "Image files (*.jpg, *.jpeg, *.jpe, *.jfif, *.png) | *.jpg; *.jpeg; *.jpe; *.jfif; *.png | All Files|*.*";

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
                    txtAttPath1.Text = AttachFilePath.ToString();
                }
                else if (ClickPoint == "2")
                {
                    txtAttFile2.Text = AttachFileName;
                    txtAttPath2.Text = AttachFilePath.ToString();
                }
                else if (ClickPoint == "3")
                {
                    txtAttFile3.Text = AttachFileName;
                    txtAttPath3.Text = AttachFilePath.ToString();
                }

                strTemp = new string[] { AttachFileName, AttachFilePath.ToString() };
                listFtpFile.Add(strTemp);
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
            MakeFolder = FolderInfoAndFlag(fileListSimple, MakeFolderName.Trim());

            if (MakeFolder == false)        // 같은 아이를 찾지 못한경우,
            {
                //MIL 폴더에 InspectionID로 저장
                if (_ftp.createDirectory(MakeFolderName.Trim()) == false)
                {
                    MessageBox.Show("업로드를 위한 폴더를 생성할 수 없습니다.");
                    return false;
                }
            }
            else
            {
                fileListDetail = _ftp.directoryListSimple(MakeFolderName.Trim(), Encoding.Default);
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
                    listStrArrayFileInfo[i][0] = MakeFolderName.Trim() + "/" + listStrArrayFileInfo[i][0];
                    UpdateFilesInfo.Add(listStrArrayFileInfo[i]);
                }
            }
            if (UpdateFilesInfo.Count > 0)
            {
                if (!_ftp.UploadTempFilesToFTP(UpdateFilesInfo))
                {
                    MessageBox.Show("파일업로드에 실패하였습니다.");
                    return false;
                }
            }
            return true;
        }



        //파일 삭제(FTP상에서)_폴더 삭제는 X
        private void FTP_UploadFile_File_Delete(string strSaveName, string FileName)
        {
            if (!_ftp.delete(strSaveName + "/" + FileName))
            {
                MessageBox.Show("파일이 삭제되지 않았습니다.");
            }
        }



        // 파일 삭제하기.(이건 텍스트만 삭제..뭐냐 이거..)
        private void btnFileDel_Click(object sender, RoutedEventArgs e)
        {
            MessageBoxResult msgresult = MessageBox.Show("파일을 삭제 하시겠습니까?", "삭제 확인", MessageBoxButton.YesNo);
            if (msgresult == MessageBoxResult.Yes)
            {
                string ClickPoint = ((Button)sender).Tag.ToString();

                if ((ClickPoint == "1") && (txtAttPath1.Text != string.Empty))
                {
                    //if (DetectFtpFile(txtDrawID.Text))
                    //{
                    //    FTP_UploadFile_File_Delete(txtDrawID.Text, txtAttFile1.Text);
                    //}

                    txtAttFile1.Text = string.Empty;
                    txtAttPath1.Text = string.Empty;
                }
                if ((ClickPoint == "2") && (txtAttPath2.Text != string.Empty))
                {
                    //if (DetectFtpFile(txtDrawID.Text))
                    //{
                    //    FTP_UploadFile_File_Delete(txtDrawID.Text, txtAttFile2.Text);
                    //}

                    txtAttFile2.Text = string.Empty;
                    txtAttPath2.Text = string.Empty;
                }
                if ((ClickPoint == "3") && (txtAttPath3.Text != string.Empty))
                {
                    //if (DetectFtpFile(txtDrawID.Text))
                    //{
                    //    FTP_UploadFile_File_Delete(txtDrawID.Text, txtAttFile3.Text);
                    //}

                    txtAttFile3.Text = string.Empty;
                    txtAttPath3.Text = string.Empty;
                }
            }
        }

        // 파일 내려받기.
        private void btnFileDown_Click(object sender, RoutedEventArgs e)
        {
            MessageBoxResult msgresult = MessageBox.Show("파일을 다운로드 하시겠습니까?", "다운로드 확인", MessageBoxButton.YesNo);
            if (msgresult == MessageBoxResult.Yes)
            {
                //버튼 태그값.
                string ClickPoint = ((Button)sender).Tag.ToString();

                if ((ClickPoint == "1") && (txtAttPath1.Text == string.Empty))
                {
                    MessageBox.Show("파일이 없습니다.");
                    return;
                }
                if ((ClickPoint == "2") && (txtAttPath2.Text == string.Empty))
                {
                    MessageBox.Show("파일이 없습니다.");
                    return;
                }
                if ((ClickPoint == "3") && (txtAttPath3.Text == string.Empty))
                {
                    MessageBox.Show("파일이 없습니다.");
                    return;
                }

                var ViewReceiver = dgdDraw.SelectedItem as Win_Qul_Drawing_U_CodeView;
                if (ViewReceiver != null)
                {
                    if (ClickPoint == "1")
                    {
                        FTP_DownLoadFile(ViewReceiver.AttPath1, ViewReceiver.DrawID, ViewReceiver.AttFile1);
                    }
                    else if (ClickPoint == "2")
                    {
                        FTP_DownLoadFile(ViewReceiver.AttPath2, ViewReceiver.DrawID, ViewReceiver.AttFile2);
                    }
                    else if (ClickPoint == "3")
                    {
                        FTP_DownLoadFile(ViewReceiver.AttPath3, ViewReceiver.DrawID, ViewReceiver.AttFile3);
                    }
                }
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







        #endregion

        #region CRUD

        private void FillGrid()
        {
            try
            {
                Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                int i = 0;
                sqlParameter.Add("chkDate", chkDrawCreateDateSrh.IsChecked == true ? (rbnCreateDate.IsChecked == true ? 1 : 2) : 0);
                sqlParameter.Add("FromDate", chkDrawCreateDateSrh.IsChecked == true ? dtpSDate.SelectedDate.Value.ToString("yyyyMMdd") : "");
                sqlParameter.Add("ToDate", chkDrawCreateDateSrh.IsChecked == true ? dtpEDate.SelectedDate.Value.ToString("yyyyMMdd") : "");
                sqlParameter.Add("nArticleID", chkArticleSrh.IsChecked == true ? (txtArticelSrh.Tag != null ? 1 : 2) : 0);
                sqlParameter.Add("ArticleID", chkArticleSrh.IsChecked == true ? (txtArticelSrh.Tag != null ? @Escape(txtArticelSrh.Tag.ToString()) : @Escape(txtArticelSrh.Text)) : "");
                sqlParameter.Add("nDrawNo", chkDrawNoSrh.IsChecked == true ? 1 : 0);
                sqlParameter.Add("DrawNo", chkDrawNoSrh.IsChecked == true ? txtDrawNoSrh.Text : "");

                DataSet ds = DataStore.Instance.ProcedureToDataSet("xp_dvlDraw_sDraw", sqlParameter, false);

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

                        foreach (DataRow dr in drc)
                        {
                            var WinDrawing = new Win_Qul_Drawing_U_CodeView()
                            {
                                DrawID = dr["DrawID"].ToString(),
                                DrawNo = dr["DrawNo"].ToString(),
                                ArticleID = dr["ArticleID"].ToString(),
                                Article = dr["Article"].ToString(),
                                BuyerArticleNo = dr["BuyerArticleNo"].ToString(),
                                BuyerModelID = dr["BuyerModelID"].ToString(),
                                BuyerModelName = dr["BuyerModelName"].ToString(),
                                DrawCreateDate = dr["DrawCreateDate"].ToString(),
                                DrawCreateMan = dr["DrawCreateMan"].ToString(),
                                EcoNO = dr["EcoNO"].ToString(),
                                DrawDvlyPlace = dr["DrawDvlyPlace"].ToString(),
                                DrawAuthor = dr["DrawAuthor"].ToString(),
                                DrawAuthorname = dr["DrawAuthorname"].ToString(),
                                DrawAcptDate = dr["DrawAcptDate"].ToString(),
                                DrawAcptMan = dr["DrawAcptMan"].ToString(),
                                Comments = dr["Comments"].ToString(),
                                AttPath1 = dr["AttPath1"].ToString(),
                                AttFile1 = dr["AttFile1"].ToString(),
                                AttPath2 = dr["AttPath2"].ToString(),
                                AttFile2 = dr["AttFile2"].ToString(),
                                AttPath3 = dr["AttPath3"].ToString(),
                                AttFile3 = dr["AttFile3"].ToString()
                            };

                            if (WinDrawing.DrawCreateDate != null && !WinDrawing.DrawCreateDate.Equals(""))
                            {
                                WinDrawing.DrawCreateDate_CV = lib.StrDateTimeBar(WinDrawing.DrawCreateDate);
                                chkDrawCreateDate.IsChecked = true;
                            }
                            if (WinDrawing.DrawAcptDate != null && !WinDrawing.DrawAcptDate.Equals(""))
                            {
                                WinDrawing.DrawAcptDate_CV = lib.StrDateTimeBar(WinDrawing.DrawAcptDate);
                                chkDrawAcptDate.IsChecked = true;
                            }

                            if (dicCompare.Count > 0)
                            {
                                if (WinDrawing.DrawID.Equals(dicCompare["DrawID"].ToString()))
                                {
                                    lstCompareValue.Add(WinDrawing.ToString());
                                }
                            }
                            i++;
                            dgdDraw.Items.Add(WinDrawing);
                        }
                        tbkIndexCount.Text = "▶검색결과 : " + i + "건";
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

        private void dgdDraw_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            Win_Qul_Drawing_U_CodeView winDraw = dgdDraw.SelectedItem as Win_Qul_Drawing_U_CodeView;

            if (winDraw != null)
            {
                SetBuyerArticelNo(winDraw.ArticleID);
                this.DataContext = winDraw;
                //cboBuyerArticle.SelectedValue = winDraw.BuyerArticleNo;
                txtBuyerArticle.Text = winDraw.BuyerArticleNo;

                if (winDraw.DrawCreateDate.Trim() == "")
                {
                    chkDrawCreateDate.IsChecked = false;
                }
                else
                {
                    chkDrawCreateDate.IsChecked = true;
                }

                if (winDraw.DrawAcptDate.Trim() == "")
                {
                    chkDrawAcptDate.IsChecked = false;
                }
                else
                {
                    chkDrawAcptDate.IsChecked = true;
                }

                strAttPath1 = winDraw.AttPath1;
                strAttPath2 = winDraw.AttPath2;
                strAttPath3 = winDraw.AttPath3;
            }
        }

        private bool SaveData(string strDrawID, string strflag)
        {
            bool Flag = false;
            List<Procedure> Prolist = new List<Procedure>();
            List<Dictionary<string, object>> ListParameter = new List<Dictionary<string, object>>();


            if (CheckData())
            {
                try
                {
                    Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                    sqlParameter.Add("DrawID", strDrawID);
                    sqlParameter.Add("DrawNo", txtDrawNo.Text);
                    sqlParameter.Add("ArticleID", txtArticle.Tag.ToString());
                    sqlParameter.Add("BuyerModelID", txtBuyerModel.Tag.ToString());
                    sqlParameter.Add("DrawCreateDate", (chkDrawCreateDate.IsChecked == true ? dtpDrawCreateDate.SelectedDate.Value.ToString("yyyyMMdd") : ""));
                    sqlParameter.Add("DrawCreateMan", txtDrawCreateMan.Text);
                    sqlParameter.Add("EcoNO", txtECONo.Text);
                    sqlParameter.Add("DrawDvlyPlace", txtDrawDvlyPlace.Text);
                    sqlParameter.Add("DrawAuthor", cboDrawAuthor.SelectedValue.ToString());
                    sqlParameter.Add("DrawAcptDate", (chkDrawAcptDate.IsChecked == true ? dtpDrawAcptDate.SelectedDate.Value.ToString("yyyyMMdd") : ""));
                    sqlParameter.Add("DrawAcptMan", txtDrawAcptMan.Text);
                    sqlParameter.Add("Comments", txtComments.Text);
                    sqlParameter.Add("AttFile1", "");
                    sqlParameter.Add("AttPath1", "");
                    sqlParameter.Add("AttFile2", "");
                    sqlParameter.Add("AttPath2", "");
                    sqlParameter.Add("AttFile3", "");
                    sqlParameter.Add("AttPath3", "");
                    sqlParameter.Add("UserID", MainWindow.CurrentUser);

                    //lstCompareValue.Add(txtDrawNo.Text);
                    //lstCompareValue.Add(txtArticle.Tag.ToString());
                    //lstCompareValue.Add(cboBuyerArticle.SelectedValue.ToString());
                    //lstCompareValue.Add((chkDrawCreateDate.IsChecked == true ? dtpDrawCreateDate.SelectedDate.Value.ToString("yyyyMMdd") : ""));

                    if (strFlag.Equals("I"))
                    {
                        Procedure pro1 = new Procedure();
                        pro1.Name = "xp_dvlDraw_iDraw";
                        pro1.OutputUseYN = "Y";
                        pro1.OutputName = "DrawID";
                        pro1.OutputLength = "10";

                        Prolist.Add(pro1);
                        ListParameter.Add(sqlParameter);

                        List<KeyValue> list_Result = new List<KeyValue>();
                        list_Result = DataStore.Instance.ExecuteAllProcedureOutputGetCS(Prolist, ListParameter);
                        string sGetDrawID = string.Empty;

                        if (list_Result[0].key.ToLower() == "success")
                        {
                            list_Result.RemoveAt(0);
                            for (int i = 0; i < list_Result.Count; i++)
                            {
                                KeyValue kv = list_Result[i];
                                if (kv.key == "DrawID")
                                {
                                    sGetDrawID = kv.value;
                                    //lstCompareValue.Add(sGetDrawID);
                                    dicCompare.Add("DrawID", sGetDrawID);
                                }
                            }

                            bool AttachYesNo = false;

                            if (txtAttFile1.Text != string.Empty || txtAttFile2.Text != string.Empty || txtAttFile3.Text != string.Empty)       //첨부파일 1
                            {
                                if (FTP_Save_File(listFtpFile, sGetDrawID))
                                {
                                    if (!txtAttFile1.Text.Equals(string.Empty)) { txtAttPath1.Text = "/ImageData/Draw/" + sGetDrawID; }
                                    if (!txtAttFile2.Text.Equals(string.Empty)) { txtAttPath2.Text = "/ImageData/Draw/" + sGetDrawID; }
                                    if (!txtAttFile3.Text.Equals(string.Empty)) { txtAttPath3.Text = "/ImageData/Draw/" + sGetDrawID; }

                                    AttachYesNo = true;
                                }
                                else
                                { MessageBox.Show("데이터 저장이 완료되었지만, 첨부문서 등록에 실패하였습니다."); }

                                if (AttachYesNo == true) { AttachFileUpdate(sGetDrawID); }      //첨부문서 정보 DB 업데이트.
                            }
                            Flag = true;
                        }
                        else
                        {
                            MessageBox.Show("[저장실패]\r\n" + list_Result[0].value.ToString());
                            //return false;
                        }
                    }
                    else    //flag="U";
                    {
                        dicCompare.Add("DrawID", strDrawID);

                        Procedure pro1 = new Procedure();
                        pro1.Name = "xp_dvlDraw_uDraw";
                        pro1.OutputUseYN = "N";
                        pro1.OutputName = "DrawID";
                        pro1.OutputLength = "10";

                        Prolist.Add(pro1);
                        ListParameter.Add(sqlParameter);

                        string[] Confirm = new string[2];
                        Confirm = DataStore.Instance.ExecuteAllProcedureOutputNew(Prolist, ListParameter);

                        if (Confirm[0] == "success")
                        {
                            bool AttachYesNo = false;

                            if (txtAttFile1.Text != string.Empty || txtAttFile2.Text != string.Empty || txtAttFile3.Text != string.Empty)       //첨부파일 1
                            {
                                if (FTP_Save_File(listFtpFile, txtDrawID.Text))
                                {
                                    if (!txtAttFile1.Text.Equals(string.Empty)) { txtAttPath1.Text = "/ImageData/Draw/" + txtDrawID.Text; }
                                    if (!txtAttFile2.Text.Equals(string.Empty)) { txtAttPath2.Text = "/ImageData/Draw/" + txtDrawID.Text; }
                                    if (!txtAttFile3.Text.Equals(string.Empty)) { txtAttPath3.Text = "/ImageData/Draw/" + txtDrawID.Text; }

                                    AttachYesNo = true;
                                }
                                else
                                { MessageBox.Show("데이터 수정이 완료되었지만, 첨부문서 등록에 실패하였습니다."); }

                                if (AttachYesNo == true) { AttachFileUpdate(txtDrawID.Text); }      //첨부문서 정보 DB 업데이트.
                            }
                            Flag = true;
                        }
                        else
                        {
                            MessageBox.Show("[저장실패]\r\n" + Confirm[1].ToString());
                            //return false;
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

            return Flag;
        }

        private bool DeleteData(string strDrawID)
        {
            bool Flag = false;

            try
            {
                Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                sqlParameter.Clear();
                sqlParameter.Add("DrawID", strDrawID);

                string[] result = DataStore.Instance.ExecuteProcedure("xp_dvlDraw_dDraw", sqlParameter, true);

                if (result[0].Equals("success"))
                {
                    //MessageBox.Show("성공 *^^*");
                    Flag = true;
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

            return Flag;
        }

        #endregion

        #region FTP 

        // 1) 첨부문서가 있을경우, 2) FTP에 정상적으로 업로드가 완료된 경우.  >> DB에 정보 업데이트 
        private void AttachFileUpdate(string ID)
        {
            try
            {
                Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                sqlParameter.Clear();
                sqlParameter.Add("DrawID", ID);

                sqlParameter.Add("AttPath1", txtAttPath1.Text);
                sqlParameter.Add("AttFile1", txtAttFile1.Text);
                sqlParameter.Add("AttPath2", txtAttPath2.Text);
                sqlParameter.Add("AttFile2", txtAttFile2.Text);
                sqlParameter.Add("AttPath3", txtAttPath3.Text);
                sqlParameter.Add("AttFile3", txtAttFile3.Text);
                sqlParameter.Add("UpdateUserID", MainWindow.CurrentUser);

                string[] result = DataStore.Instance.ExecuteProcedure("xp_dvlDraw_uDraw_Ftp", sqlParameter, true);
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

        private bool CheckData()
        {
            bool flag = true;

            if (txtArticle.Tag == null) //품명 텍스트박스를 품번으로 쓰고 있기 때문에
            {
                MessageBox.Show("품번 선택이 잘못되었습니다. enter키 또는 품번 옆의 버튼을 이용하여 다시 입력해주세요");
                flag = false;
                return flag;
            }

            if (txtBuyerModel.Text == null || txtBuyerModel.Text.ToString().Equals(""))
            {
                MessageBox.Show("차종이 선택되지 않았습니다. 선택해주세요");
                flag = false;
                return flag;
            }

            if (cboDrawAuthor.SelectedValue == null)
            {
                MessageBox.Show("도면권한이 선택되지 않았습니다. 선택해주세요");
                flag = false;
                return flag;
            }

            if (chkDrawCreateDate.IsChecked == true && dtpDrawCreateDate.SelectedDate == null)
            {
                MessageBox.Show("도면생성일 날짜가 입력되지 않았습니다. 입력해주세요");
                flag = false;
                return flag;
            }

            if (chkDrawAcptDate.IsChecked == false || (chkDrawAcptDate.IsChecked == true && dtpDrawAcptDate.SelectedDate == null))
            {
                MessageBox.Show("접수일 날짜가 입력되지 않았습니다. 입력해주세요");
                flag = false;
                return flag;
            }

            return flag;
        }

        #endregion

        private void cboBuyerArticle_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                e.Handled = true;

                //if (cboBuyerArticle.ItemsSource != null && cboBuyerArticle.SelectedIndex != -1)
                //{

                //}
            }
        }
        // 특수문자 참조시 사용
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
    }

    class Win_Qul_Drawing_U_CodeView : BaseView
    {
        public override string ToString()
        {
            return (this.ReportAllProperties());
        }

        public string DrawID { get; set; }
        public string DrawNo { get; set; }
        public string ArticleID { get; set; }
        public string Article { get; set; }
        public string Article_Sabun { get; set; }
        public string BuyerArticleNo { get; set; }
        public string BuyerModelID { get; set; }
        public string BuyerModelName { get; set; }
        public string DrawCreateDate { get; set; }
        public string DrawCreateMan { get; set; }
        public string EcoNO { get; set; }
        public string DrawDvlyPlace { get; set; }
        public string DrawAuthor { get; set; }
        public string DrawAuthorname { get; set; }
        public string DrawAcptDate { get; set; }
        public string DrawAcptMan { get; set; }
        public string Comments { get; set; }
        public string AttPath1 { get; set; }
        public string AttFile1 { get; set; }
        public string AttPath2 { get; set; }
        public string AttFile2 { get; set; }
        public string AttPath3 { get; set; }
        public string AttFile3 { get; set; }
        public string DrawCreateDate_CV { get; set; }
        public string DrawAcptDate_CV { get; set; }
    }

}

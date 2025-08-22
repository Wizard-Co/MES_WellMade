using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
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
using WizMes_WellMade;
using WPF.MDI;
using System.Net;
using System.Threading;

namespace WizMes_WellMade
{
    /// <summary>
    /// Win_prd_RegularInspectBasis_U.xaml에 대한 상호 작용 논리
    /// </summary>
    public partial class Win_prd_RegularInspectBasis_U : UserControl
    {
        public Win_prd_RegularInspectBasis_U()
        {
            InitializeComponent();
        }

        string stDate = string.Empty;
        string stTime = string.Empty;

        string strFlag = string.Empty;
        bool strCopy = false;
        int rowMainNum = 0;  //메인데이터그리드 rowNum
        int rowSubNum = 0;   //서브데이터그리드 rowNum

        string PRD_MCNAME = "McName";
        Lib lib = new Lib();

        Win_prd_RegularInspectBasis_U_CodeView WinMcRegular = new Win_prd_RegularInspectBasis_U_CodeView();
        Win_prd_RegularInspectBasis_U_CodeView WinMcRegularSub = new Win_prd_RegularInspectBasis_U_CodeView();

        //파일 수정 진행 위한 flag 3가지
        bool existFtp = false;
        bool AddFtp = false;
        bool delFtp = false;

        //FTP 활용모음
        string strImagePath = string.Empty;
        string strFullPath = string.Empty;
        string strDelFileName = string.Empty;

        List<string> deleteListFtpFile = new List<string>(); // 삭제할 파일 리스트
        List<string> lstExistFtpFile = new List<string>();

        // 촤! FTP Server 에 있는 폴더 + 파일 경로를 저장해놓고 그걸로 다운 및 업로드하자 마!
        // 이미지 이름 : 폴더이름
        Dictionary<string, string> lstFtpFilePath = new Dictionary<string, string>();

        private FTP_EX _ftp = null;

        List<string[]> listFtpFile = new List<string[]>();
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

        //삭제할 데이터를 모으는 변수
        List<Win_prd_RegularInspectBasis_U_CodeView> DelItems = new List<Win_prd_RegularInspectBasis_U_CodeView>();


        //string FTP_ADDRESS = "ftp://wizis.iptime.org/ImageData/McRegularInspect";

        //알 FTP test
        //string FTP_ADDRESS = "ftp://192.168.0.120";

        //string FTP_ADDRESS = "ftp://wizis.iptime.org/ImageData/Draw";

        string FTP_ADDRESS = "ftp://" + LoadINI.FileSvr + ":" + LoadINI.FTPPort + LoadINI.FtpImagePath + "/McRIB";
        string ForderName = "McRIB";
        private const string FTP_ID = "wizuser";
        private const string FTP_PASS = "wiz9999";

        private const string LOCAL_DOWN_PATH = "C:\\Temp";

        // FTP 문제 일 때 메시지 모으기
        private string Message = "";
        private List<string> lstMsg = new List<string>();



        private void UserControl_Loaded(object sender, RoutedEventArgs e)
        {
            stDate = DateTime.Now.ToString("yyyyMMdd");
            stTime = DateTime.Now.ToString("HHmm");

            DataStore.Instance.InsertLogByFormS(this.GetType().Name, stDate, stTime, "S");
            Lib.Instance.UiLoading(sender);
        }

        #region 상단 조건부

        //기준번호 클릭
        private void lblStandardNumberSrh_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            if (chkStandardNumberSrh.IsChecked == true) { chkStandardNumberSrh.IsChecked = false; }
            else { chkStandardNumberSrh.IsChecked = true; }
        }

        //기준번호 체크
        private void chkStandardNumberSrh_Checked(object sender, RoutedEventArgs e)
        {
            txtStandardNumberSrh.IsEnabled = true;
            btnStandardNumberSrh.IsEnabled = true;
        }

        //기준번호 체크해제
        private void chkStandardNumberSrh_Unchecked(object sender, RoutedEventArgs e)
        {
            txtStandardNumberSrh.IsEnabled = false;
            btnStandardNumberSrh.IsEnabled = false;
        }

        //기준번호 엔터키 이벤트용(상단)
        private void txtStandardNumberSrh_KeyDown(object sender, KeyEventArgs e)
        {
            //이건 왜 없는 거야
        }

        //기준번호 버튼 클릭 이벤트용(상단)
        private void btnStandardNumberSrh_Click(object sender, RoutedEventArgs e)
        {
            //이건 왜 없는 거냐고
        }

        //상단 설비명 조건 검색 Label
        private void lblMcPartNameSrh_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            if (chkMcPartNameSrh.IsChecked == true)
            {
                chkMcPartNameSrh.IsChecked = false;
            }
            else
            {
                chkMcPartNameSrh.IsChecked = true;
            }
        }

        //상단 설비명 조건 검색 CheckBox
        private void chkMcPartNameSrh_Checked(object sender, RoutedEventArgs e)
        {
            txtMcPartNameSrh.IsEnabled = true;
            btnMcPartNameSrh.IsEnabled = true;
        }

        //상단 설비명 조건 검색 CheckBox UnChecked
        private void chkMcPartNameSrh_Unchecked(object sender, RoutedEventArgs e)
        {
            txtMcPartNameSrh.IsEnabled = false;
            btnMcPartNameSrh.IsEnabled = false;
        }

        //상단 설비명 조건 검색 텍스트박스 keyDown
        private void txtMcPartNameSrh_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                //MainWindow.pf.ReturnCode(txtMcPart, 12, "");  // 설비명 선택해도 텍스트박스에 반영이 안돼 수정함
                MainWindow.pf.ReturnCode(txtMcPartNameSrh, 12, "");
            }
        }

        //상단 설비명 조건 검색 플러스 파인더
        private void btnMcPartNameSrh_Click(object sender, RoutedEventArgs e)
        {
            MainWindow.pf.ReturnCode(txtMcPartNameSrh, 12, "");
        }

        #endregion 상단 조건부

        #region CRUD 

        //복사추가 
        private void btnCopyAdd_Click(object sender, RoutedEventArgs e)
        {
            WinMcRegular = dgdMain.SelectedItem as Win_prd_RegularInspectBasis_U_CodeView;

            if (WinMcRegular != null)
            {
                rowMainNum = dgdMain.SelectedIndex;
                dgdMain.IsHitTestVisible = false;
                tbkMsg.Text = "자료 추가 중";
                lblMsg.Visibility = Visibility.Visible;
                //개정일자가 체크되어 있도록
                chkRevision.IsChecked = true;
                //추가 버튼 눌렀을 때 개정일자는 오늘 날짜로 자동 설정되도록.
                //dtpRevision.SelectedDate = DateTime.Now;
                CantBtnControl();
                txtStandardNumber.Clear();

                //strCopy = true;
                //FillGridSub(WinMcRegular.McInspectBasisID);

                lstFtpFilePath.Clear();

                // 폴더이름 → 대개 객체 PK 값
                string Key = WinMcRegular.McInspectBasisID;
                                                          

                for (int i = 0; i < dgdSub.Items.Count; i++)
                {
                    var sub = dgdSub.Items[i] as Win_prd_RegularInspectBasis_U_CodeView;


                    //이미지 파일이 있을 경우에만 lstFtpFilePath에 정보를 넣도록 하고싶다 ㅠㅠ
                    if(sub.McImageFile != null)
                    {
                        if (!sub.McImageFile.ToString().Trim().Equals(""))
                        {
                            if (!lstFtpFilePath.ContainsKey(sub.McImageFile))
                            {
                                lstFtpFilePath.Add(sub.McImageFile, Key);
                            }
                        }
                    }      
                        
                }
                
                strFlag = "I";
            }
            else
            {
                MessageBox.Show("복사할 대상이 선택되지 않았습니다.");
            }
        }

        //추가
        private void btnAdd_Click(object sender, RoutedEventArgs e)
        {
            CantBtnControl();
            strCopy = false;
            strFlag = "I";

            if (dgdSub.Items.Count > 0)
            {
                dgdSub.Items.Clear();
                dgdSub.Refresh();
            }

            dgdMain.IsHitTestVisible = false;
            lblMsg.Visibility = Visibility.Visible;
            chkRevision.IsChecked = true;
            tbkMsg.Text = "자료 입력 중";
            rowMainNum = dgdMain.SelectedIndex;
            this.DataContext = null;

            //추가 버튼 눌렀을 때 개정일자는 오늘 날짜로 자동 설정되도록.
            dtpRevision.SelectedDate = DateTime.Now;

            //추가를 누르면 설비명에 포커스 이동
            txtMcPart.Focus();
        }

        //수정
        private void btnUpdate_Click(object sender, RoutedEventArgs e)
        {
            WinMcRegular = dgdMain.SelectedItem as Win_prd_RegularInspectBasis_U_CodeView;

            if (WinMcRegular != null)
            {
                // 삭제할 데이터를 임시적으로 저장하는 변수를 초기화
                DelItems.Clear();

                rowMainNum = dgdMain.SelectedIndex;
                dgdMain.IsHitTestVisible = false;
                dgdSub.IsHitTestVisible = true;
                tbkMsg.Text = "자료 수정 중";
                lblMsg.Visibility = Visibility.Visible;
                chkRevision.IsChecked = true;
                CantBtnControl();
                strCopy = false;
                strFlag = "U";
            }
        }

        //삭제
        private void btnDelete_Click(object sender, RoutedEventArgs e)
        {
            WinMcRegular = dgdMain.SelectedItem as Win_prd_RegularInspectBasis_U_CodeView;
            List<string> lstArrayFileName = new List<string>();

            if (WinMcRegular == null)
            {
                MessageBox.Show("삭제할 데이터가 지정되지 않았습니다. 삭제데이터를 지정하고 눌러주세요");
            }
            else
            {
                if (MessageBox.Show("선택하신 항목을 삭제하시겠습니까?", "삭제 전 확인", MessageBoxButton.YesNo) == MessageBoxResult.Yes)
                {
                    if (dgdMain.Items.Count > 0 && dgdMain.SelectedItem != null)
                    {
                        rowMainNum = dgdMain.SelectedIndex;
                    }

                    if (DeleteData(WinMcRegular.McInspectBasisID))
                    {
                        for (int i = 0; i < dgdSub.Items.Count; i++)
                        {
                            var Sub = dgdSub.Items[i] as Win_prd_RegularInspectBasis_U_CodeView;
                            if (Sub != null
                                && Sub.McImageFile != null
                                && Sub.McImageFile.ToString().Trim().Equals("") == false)
                            {
                                deleteListFtpFile.Add(Sub.McInspectBasisID + "/" + Sub.McImageFile);
                            }
                        }

                        if (deleteListFtpFile.Count > 0)
                        {
                            FTP_RemoveFileList(deleteListFtpFile);
                        }

                        rowMainNum -= 1;
                        re_Search(rowMainNum);
                    }
                    else
                    {
                        MessageBox.Show("사용중인 설비이므로 삭제할 수 없습니다.");
                        return;
                    }
                }
            }
        }
        //파일 삭제
        private bool FTP_RemoveFileList(List<string> delete)
        {
            bool flag = false;

            try
            {
                _ftp = new FTP_EX(FTP_ADDRESS, FTP_ID, FTP_PASS);

                for (int i = 0; i < delete.Count; i++)
                {
                    _ftp.delete(delete[i]);
                }
            }
            catch (Exception ex)
            {

            }

            return flag;
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

        //검색(조회)
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
            if (SaveData(strFlag, txtStandardNumber.Text))
            {
                CanBtnControl();
                lblMsg.Visibility = Visibility.Hidden;

                if (strFlag.Equals("I"))
                {

                    rowMainNum = 0;
                    re_Search(rowMainNum);

                    rowMainNum = dgdMain.Items.Count - 1;
                    re_Search(rowMainNum);
                }
                else
                {
                    rowMainNum = dgdMain.SelectedIndex;
                }


                //if (!strFlag.Trim().Equals("U"))
                //{
                //    rowMainNum = 0;
                //}
                dgdMain.IsHitTestVisible = true;
                existFtp = false;
                delFtp = false;
                AddFtp = false;
                strFlag = string.Empty;
                strImagePath = string.Empty;
                strDelFileName = string.Empty;

                re_Search(rowMainNum);

            }
        }

        //취소
        private void btnCancel_Click(object sender, RoutedEventArgs e)
        {
            CanBtnControl();



            dgdMain.IsHitTestVisible = true;
            strCopy = false;


            // 삭제 임시 변수 초기화
            DelItems.Clear();

            if (!strFlag.Equals(string.Empty))
            {
                re_Search(rowMainNum);
            }
            strFlag = string.Empty;
        }

        //엑셀
        private void btnExcel_Click(object sender, RoutedEventArgs e)
        {
            DataTable dt = null;
            string Name = string.Empty;

            string[] lst = new string[4];
            lst[0] = "설비점검 기준번호";
            lst[1] = "설비점검 기준항목 및 내용";
            lst[2] = dgdMain.Name;
            lst[3] = dgdSub.Name;

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
                else if (ExpExc.choice.Equals(dgdSub.Name))
                {
                    if (ExpExc.Check.Equals("Y"))
                        dt = Lib.Instance.DataGridToDTinHidden(dgdSub);
                    else
                        dt = Lib.Instance.DataGirdToDataTable(dgdSub);

                    Name = dgdSub.Name;
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

        #endregion CRUD


        //메인 데이터그리드 SelectionChanged
        private void dgdMain_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {

            WinMcRegular = dgdMain.SelectedItem as Win_prd_RegularInspectBasis_U_CodeView;

            if (WinMcRegular != null)
            {
                this.DataContext = WinMcRegular;

                //설비코드를 다시 넣어보자. 2020.10.29.
                txtMcPart.Tag = WinMcRegular.MCID;

                FillGridSub(WinMcRegular.McInspectBasisID);
            }
        }

        #region DataContext, 텍스트박스들~ 

        //텍스트박스 설비명 keydown
        private void txtMcPart_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                MainWindow.pf.ReturnCode(txtMcPart, 12, "");

                //개정일자에 포커스 이동
                dtpRevision.Focus();
            }
        }

        //플러스 파인더 설비명
        private void btnMcPart_Click(object sender, RoutedEventArgs e)
        {
            MainWindow.pf.ReturnCode(txtMcPart, 12, "");

            //개정일자에 포커스 이동
            dtpRevision.Focus();
        }

        //체크박스 개정일자 Check
        private void chkRevision_Checked(object sender, RoutedEventArgs e)
        {
            dtpRevision.IsEnabled = true;
        }

        //체크박스 개정일자 UnCheck
        private void chkRevision_Unchecked(object sender, RoutedEventArgs e)
        {
            dtpRevision.IsEnabled = false;
        }

        #endregion DataContext, 텍스트박스들~ 

        #region Sub DataGrid 이벤트

        //Sub 데이터 그리드 KeyDown
        private void DataGridCell_KeyDown(object sender, KeyEventArgs e)
        {
            WinMcRegularSub = dgdSub.CurrentItem as Win_prd_RegularInspectBasis_U_CodeView;
            int rowCount = dgdSub.Items.IndexOf(dgdSub.CurrentItem);
            int colCount = dgdSub.Columns.IndexOf(dgdSub.CurrentCell.Column) + 1;
            rowSubNum = rowCount;

            if (e.Key == Key.Enter)
            {
                e.Handled = true;
                (sender as DataGridCell).IsEditing = false;

                if (dgdSub.Items.Count - 1 > rowCount && dgdSub.Columns.Count - 1 > colCount)
                {
                    dgdSub.CurrentCell = new DataGridCellInfo(dgdSub.Items[rowCount], dgdSub.Columns[colCount]);
                }
                else if (dgdSub.Items.Count - 1 > rowCount && dgdSub.Columns.Count - 1 == colCount)
                {
                    dgdSub.CurrentCell = new DataGridCellInfo(dgdSub.Items[rowCount + 1], dgdSub.Columns[1]);
                }
                else if (dgdSub.Items.Count - 1 == rowCount && dgdSub.Columns.Count - 1 > colCount)
                {
                    dgdSub.CurrentCell = new DataGridCellInfo(dgdSub.Items[rowCount], dgdSub.Columns[colCount]);
                }
                else if (dgdSub.Items.Count - 1 == rowCount && dgdSub.Columns.Count - 1 == colCount)
                {
                    btnSave.Focus();
                }
                else
                {
                    MessageBox.Show("이런 경우가 있나요??");
                }
            }
        }

        //Sub 데이터 그리드 Focus 이벤트 
        private void TextBoxFocusInDataGrid(object sender, KeyEventArgs e)
        {
            Lib.Instance.DataGridINControlFocus(sender, e);
        }

        //Sub 데이터 그리드 MouseUp 이벤트
        private void TextBoxFocusInDataGrid_MouseUp(object sender, MouseButtonEventArgs e)
        {
            Lib.Instance.DataGridINBothByMouseUP(sender, e);
        }

        //Sub 데이터 그리드 GotFocus 이벤트 ( 그냥 Focus랑 뭐가 다를까 )
        private void DataGridCell_GotFocus(object sender, RoutedEventArgs e)
        {
            if (lblMsg.Visibility == Visibility.Visible)
            {
                DataGridCell cell = sender as DataGridCell;
                cell.IsEditing = true;
            }
        }

        //점검항목 - 그리드 내 텍스트 Change 이벤트
        private void dgdtpetxtMcInsItemName_TextChanged(object sender, TextChangedEventArgs e)
        {
            if (lblMsg.Visibility == Visibility.Visible)
            {
                WinMcRegularSub = dgdSub.CurrentItem as Win_prd_RegularInspectBasis_U_CodeView;

                if (WinMcRegularSub != null)
                {
                    TextBox tb1 = sender as TextBox;

                    if (tb1 != null)
                    {
                        WinMcRegularSub.McInsItemName = tb1.Text;
                    }

                    sender = tb1;
                }
            }
        }
        //점검내용 - 그리드 내 텍스트 Change 이벤트
        private void dgdtpetxtMcInsContent_TextChanged(object sender, TextChangedEventArgs e)
        {
            if (lblMsg.Visibility == Visibility.Visible)
            {
                WinMcRegularSub = dgdSub.CurrentItem as Win_prd_RegularInspectBasis_U_CodeView;

                if (WinMcRegularSub != null)
                {
                    TextBox tb1 = sender as TextBox;

                    if (tb1 != null)
                    {
                        WinMcRegularSub.McInsContent = tb1.Text;
                    }

                    sender = tb1;
                }
            }
        }

        //Sub 그리드 확인방법 - 콤보박스 생성
        private void dgdtpecboMcInsCheckGbn_Loaded(object sender, RoutedEventArgs e)
        {
            ComboBox cboMcIUnsCheckGbn = (ComboBox)sender;

            ObservableCollection<CodeView> ovcMcIUnsCheckGbn =
            ComboBoxUtil.Instance.GetCMCode_SetComboBox("MCCHECKGBN", "");
            cboMcIUnsCheckGbn.ItemsSource = ovcMcIUnsCheckGbn;
            cboMcIUnsCheckGbn.DisplayMemberPath = "code_name";
            cboMcIUnsCheckGbn.SelectedValuePath = "code_id";
        }

        //Sub 그리드 확인방법 - 콤보박스 중 하나 선택
        private void dgdtpecboMcInsCheckGbn_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            WinMcRegularSub = dgdSub.CurrentItem as Win_prd_RegularInspectBasis_U_CodeView;

            ComboBox cboMcInsCheckGbn = (ComboBox)sender;

            if (WinMcRegularSub == null)
            {
                WinMcRegularSub = dgdSub.Items[rowSubNum] as Win_prd_RegularInspectBasis_U_CodeView;
            }

            if (cboMcInsCheckGbn.SelectedValue != null && !cboMcInsCheckGbn.SelectedValue.ToString().Equals(""))
            {
                var theView = cboMcInsCheckGbn.SelectedItem as CodeView;
                if (theView != null)
                {
                    WinMcRegularSub.McInsCheckGbn = theView.code_id;
                    WinMcRegularSub.McInsCheck = theView.code_name;
                }

                sender = cboMcInsCheckGbn;
            }
        }

        //Sub 그리드 확인방법 - 무한 DropDown 방지
        private void dgdtpecboMcInsCheckGbn_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                DataGridCell cell = Lib.Instance.GetParent<DataGridCell>(sender as ComboBox);
                cell.Focus();
            }
        }

        //Sub 그리드 주기 - 콤보박스 생성
        private void dgdtpecboMcInsCycleGbn_Loaded(object sender, RoutedEventArgs e)
        {
            ComboBox cboMcInsCycleGbn = (ComboBox)sender;

            ObservableCollection<CodeView> ovcMcInsCycleGbn =
                ComboBoxUtil.Instance.GetCMCode_SetComboBox("MCCYCLEGBN", "");
            cboMcInsCycleGbn.ItemsSource = ovcMcInsCycleGbn;
            cboMcInsCycleGbn.DisplayMemberPath = "code_name";
            cboMcInsCycleGbn.SelectedValuePath = "code_id";
        }
        //Sub 그리드 주기 - 콤보박스 중 하나 선택
        private void dgdtpecboMcInsCycleGbn_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            WinMcRegularSub = dgdSub.CurrentItem as Win_prd_RegularInspectBasis_U_CodeView;

            ComboBox cboMcInsCycleGbn = (ComboBox)sender;

            if (WinMcRegularSub == null)
            {
                WinMcRegularSub = dgdSub.Items[rowSubNum] as Win_prd_RegularInspectBasis_U_CodeView;
            }

            if (cboMcInsCycleGbn.SelectedValue != null && !cboMcInsCycleGbn.SelectedValue.ToString().Equals(""))
            {
                var theView = cboMcInsCycleGbn.SelectedItem as CodeView;
                if (theView != null)
                {
                    WinMcRegularSub.McInsCycleGbn = theView.code_id;
                    WinMcRegularSub.McInsCycle = theView.code_name;
                }

                sender = cboMcInsCycleGbn;
            }
        }

        //Sub 그리드 주기 - 무한 DropDown 방지
        private void dgdtpecboMcInsCycleGbn_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                DataGridCell cell = Lib.Instance.GetParent<DataGridCell>(sender as ComboBox);
                cell.Focus();
            }
        }

        //Sub 그리드 특정월일 
        private void dgdtpetxtMcInsCycleDate_TextChanged(object sender, TextChangedEventArgs e)
        {
            if (lblMsg.Visibility == Visibility.Visible)
            {
                WinMcRegularSub = dgdSub.CurrentItem as Win_prd_RegularInspectBasis_U_CodeView;

                if (WinMcRegularSub != null)
                {
                    TextBox tb1 = sender as TextBox;

                    if (tb1 != null)
                    {
                        WinMcRegularSub.McInsCycleDate = tb1.Text;
                    }

                    sender = tb1;
                }
            }
        }

        //Sub 그리드 기록구분 - 콤보박스 생성
        private void dgdtpecboMcInsRecordGbn_Loaded(object sender, RoutedEventArgs e)
        {
            ComboBox cboMcInsRecordGbn = (ComboBox)sender;

            ObservableCollection<CodeView> ovcMcInsRecordGbn =
                ComboBoxUtil.Instance.GetCMCode_SetComboBox("MCRECORDGBN", "");
            cboMcInsRecordGbn.ItemsSource = ovcMcInsRecordGbn;
            cboMcInsRecordGbn.DisplayMemberPath = "code_name";
            cboMcInsRecordGbn.SelectedValuePath = "code_id";
        }

        //Sub 그리드 기록구분 - 콤보박스 중 하나 선택
        private void dgdtpecboMcInsRecordGbn_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            WinMcRegularSub = dgdSub.CurrentItem as Win_prd_RegularInspectBasis_U_CodeView;

            ComboBox cboMcInsRecordGbn = (ComboBox)sender;

            if (WinMcRegularSub == null)
            {
                WinMcRegularSub = dgdSub.Items[rowSubNum] as Win_prd_RegularInspectBasis_U_CodeView;
            }

            if (cboMcInsRecordGbn.SelectedValue != null && !cboMcInsRecordGbn.SelectedValue.ToString().Equals(""))
            {
                var theView = cboMcInsRecordGbn.SelectedItem as CodeView;
                if (theView != null)
                {
                    WinMcRegularSub.McInsRecordGbn = theView.code_id;
                    WinMcRegularSub.McInsRecord = theView.code_name;
                }

                sender = cboMcInsRecordGbn;
            }
        }

        //Sub 그리드 기록구분 - 무한 DropDown 방지
        private void dgdtpecboMcInsRecordGbn_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                DataGridCell cell = Lib.Instance.GetParent<DataGridCell>(sender as ComboBox);
                cell.Focus();
            }
        }

        //Sub 그리드 - 이미지 keydown
        private void dgdtpetxtImage_KeyDown(object sender, KeyEventArgs e)
        {
            if (lblMsg.Visibility == Visibility.Visible)
            {

                if (e.Key == Key.Enter)
                {
                    WinMcRegularSub = dgdSub.CurrentItem as Win_prd_RegularInspectBasis_U_CodeView;

                    if (WinMcRegularSub != null)
                    {
                        if (!WinMcRegularSub.McImageFile.Trim().Equals(string.Empty) && strFlag.Equals("U"))
                        {
                            MessageBox.Show("먼저 해당파일의 삭제를 진행 후 진행해주세요.");
                            return;
                        }
                        else
                        {
                            FTP_Upload_TextBox(sender as TextBox);
                        }
                    }
                }
            }
        }

        //Sub 그리드 - 이미지 보기 클릭 이벤트
        private void btnSeeImage_Click(object sender, RoutedEventArgs e)
        {
            var WinMcRegularSub = dgdSub.CurrentItem as Win_prd_RegularInspectBasis_U_CodeView;

            if (WinMcRegularSub != null && !WinMcRegularSub.McImageFile.Equals(""))
            {
                //FTP_DownLoadFile(WinMcRegularSub.McImagePath + "/" + WinMcRegularSub.McInspectBasisID + "/" + WinMcRegularSub.McImageFile);
                FTP_DownLoadFile2(WinMcRegularSub.McInspectBasisID, WinMcRegularSub.McImageFile);
            }
        }

        //Sub 그리드 - 추가
        private void btnSubAdd_Click(object sender, RoutedEventArgs e)
        {
            SubPlus();
            int colCount = dgdSub.Columns.IndexOf(dgdtpeMcInsItemName);
            dgdSub.Focus();
            //dgdMCRepair_Sub.SelectedIndex = dgdMCRepair_Sub.Items.Count - 1;
            dgdSub.CurrentCell = new DataGridCellInfo(dgdSub.Items[dgdSub.Items.Count - 1], dgdSub.Columns[colCount]);
        }

        //Sub 그리드 - 삭제
        private void btnSubDel_Click(object sender, RoutedEventArgs e)
        {
            SubRemove();
        }

        #endregion Sub DataGrid 이벤트

        #region 주요 매서드들

        #region 버튼 조정

        /// <summary>
        /// 수정,추가 저장 후
        /// </summary>
        private void CanBtnControl()
        {
            btnAdd.IsEnabled = true;
            btnUpdate.IsEnabled = true;
            btnDelete.IsEnabled = true;
            btnSearch.IsEnabled = true;
            btnCopyAdd.IsEnabled = true;
            btnSave.Visibility = Visibility.Hidden;
            btnCancel.Visibility = Visibility.Hidden;
            btnExcel.Visibility = Visibility.Visible;
            gbxInput.IsEnabled = false;
            lblMsg.Visibility = Visibility.Hidden;

            btnSubAdd.IsEnabled = false;
            btnSubDel.IsEnabled = false;
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
            btnCopyAdd.IsEnabled = false;

            btnSave.Visibility = Visibility.Visible;
            btnCancel.Visibility = Visibility.Visible;
            btnExcel.Visibility = Visibility.Hidden;
            gbxInput.IsEnabled = true;
            lblMsg.Visibility = Visibility.Visible;

            btnSubAdd.IsEnabled = true;
            btnSubDel.IsEnabled = true;
            
        }

        #endregion 버튼 조정

        #region 저장(SaveData)

        /// <summary>
        /// 저장
        /// </summary>
        /// <param name="strFlag"></param>
        /// <param name="strYYYY"></param>
        /// <returns></returns>
        private bool SaveData(string strFlag, string strID)
        {
            bool flag = false;
            List<Procedure> Prolist = new List<Procedure>();
            List<Dictionary<string, object>> ListParameter = new List<Dictionary<string, object>>();
            List<string[]> FtpFileList = new List<string[]>();

            string GetKey = "";

            try
            {
                if (CheckData())
                {
                    Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                    sqlParameter.Clear();
                    sqlParameter.Add("McInspectBasisID", strID);
                    sqlParameter.Add("MCID", txtMcPart.Tag.ToString());
                    //sqlParameter.Add("McInsBasisDate", chkRevision.IsChecked==true ?
                    //    dtpRevision.SelectedDate.Value.ToString("yyyyMMdd") : "");
                    sqlParameter.Add("McInsBasisDate", dtpRevision.SelectedDate.Value.ToString("yyyyMMdd"));
                    sqlParameter.Add("McInsContent", txtRevisionContents.Text);
                    sqlParameter.Add("Comments", txtContents.Text);

                    #region 추가

                    if (strFlag.Equals("I"))
                    {
                        sqlParameter.Add("CreateUserID", MainWindow.CurrentUser);

                        Procedure pro1 = new Procedure();
                        pro1.Name = "xp_McRegularInspectBasis_iMcRegularInspectBasis";
                        pro1.OutputUseYN = "Y";
                        pro1.OutputName = "McInspectBasisID";
                        pro1.OutputLength = "10";

                        Prolist.Add(pro1);
                        ListParameter.Add(sqlParameter);

                        List<KeyValue> list_Result = new List<KeyValue>();
                        list_Result = DataStore.Instance.ExecuteAllProcedureOutputGetCS_NewLog(Prolist, ListParameter,"C");
                        string sGetID = string.Empty;

                        if (list_Result[0].key.ToLower() == "success")
                        {
                            list_Result.RemoveAt(0);
                            for (int i = 0; i < list_Result.Count; i++)
                            {
                                KeyValue kv = list_Result[i];
                                if (kv.key == "McInspectBasisID")
                                {
                                    sGetID = kv.value;
                                    flag = true;

                                    GetKey = kv.value;

                                    Prolist.RemoveAt(0);
                                    ListParameter.Clear();
                                }
                            }
                        }
                        else
                        {
                            MessageBox.Show("[저장실패]\r\n" + list_Result[0].value.ToString());
                            flag = false;
                            //return false;
                        }

                        for (int i = 0; i < dgdSub.Items.Count; i++)
                        {
                            WinMcRegularSub = dgdSub.Items[i] as Win_prd_RegularInspectBasis_U_CodeView;

                            sqlParameter = new Dictionary<string, object>();
                            sqlParameter.Clear();
                            sqlParameter.Add("McInspectBasisID", GetKey); //2020.01.13 strID는 null값이라
                            sqlParameter.Add("McSeq", i + 1);
                            sqlParameter.Add("McInsGbn", WinMcRegularSub.McInsGbn);
                            sqlParameter.Add("McInsItemName", WinMcRegularSub.McInsItemName);
                            sqlParameter.Add("McInsContent", WinMcRegularSub.McInsContent);
                            sqlParameter.Add("McInsCheckGbn", WinMcRegularSub.McInsCheckGbn);
                            sqlParameter.Add("McInsCycleGbn", WinMcRegularSub.McInsCycleGbn);
                            sqlParameter.Add("McInsCycleDate", ConvertInt(WinMcRegularSub.McInsCycleDate));
                            sqlParameter.Add("McInsRecordGbn", WinMcRegularSub.McInsRecordGbn);
                            sqlParameter.Add("McImageFile", WinMcRegularSub.McImageFile);
                            sqlParameter.Add("McImagePath", !WinMcRegularSub.McImageFile.Equals("") ? "/ImageData/" + ForderName + "/" + GetKey + "/" : "");
                            sqlParameter.Add("McComments", WinMcRegularSub.McComments);
                            sqlParameter.Add("CreateUserID", MainWindow.CurrentUser);

                            Procedure pro2 = new Procedure();
                            pro2.Name = "xp_McRegularInspectBasis_iMcRegularInspectBasisSub";
                            pro2.OutputUseYN = "N";
                            pro2.OutputName = "McInspectBasisID";
                            pro2.OutputLength = "10";

                            Prolist.Add(pro2);
                            ListParameter.Add(sqlParameter);

                            if (!WinMcRegularSub.McImageFile.Replace(" ", "").Equals(""))
                            {
                                string[] FtpFilePathAndName = new string[2];
                                FtpFilePathAndName[0] = WinMcRegularSub.McImageFile;
                                FtpFilePathAndName[1] = WinMcRegularSub.LocalImagePath;
                                FtpFileList.Add(FtpFilePathAndName);
                            }
                        }



                        string[] confirm = new string[2];
                        confirm = DataStore.Instance.ExecuteAllProcedureOutputNew_NewLog(Prolist, ListParameter,"U");

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
                        sqlParameter.Add("UpdateUserID", MainWindow.CurrentUser);

                        Procedure pro1 = new Procedure();
                        pro1.Name = "xp_McRegularInspectBasis_uMcRegularInspectBasis";
                        pro1.OutputUseYN = "N";
                        pro1.OutputName = "McInspectBasisID";
                        pro1.OutputLength = "10";

                        Prolist.Add(pro1);
                        ListParameter.Add(sqlParameter);

                        for (int i = 0; i < dgdSub.Items.Count; i++)
                        {
                            WinMcRegularSub = dgdSub.Items[i] as Win_prd_RegularInspectBasis_U_CodeView;

                            //Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                            sqlParameter = new Dictionary<string, object>();
                            sqlParameter.Clear();
                            sqlParameter.Add("McInspectBasisID", strID);
                            sqlParameter.Add("McSeq", i + 1);
                            sqlParameter.Add("McInsGbn", WinMcRegularSub.McInsGbn);
                            sqlParameter.Add("McInsItemName", WinMcRegularSub.McInsItemName);
                            sqlParameter.Add("McInsContent", WinMcRegularSub.McInsContent);
                            sqlParameter.Add("McInsCheckGbn", WinMcRegularSub.McInsCheckGbn);
                            sqlParameter.Add("McInsCycleGbn", WinMcRegularSub.McInsCycleGbn);
                            sqlParameter.Add("McInsCycleDate", ConvertInt(WinMcRegularSub.McInsCycleDate));
                            sqlParameter.Add("McInsRecordGbn", WinMcRegularSub.McInsRecordGbn);
                            sqlParameter.Add("McImageFile", WinMcRegularSub.McImageFile);
                            //sqlParameter.Add("McImagePath", "/ImageData/" + ForderName + "/" + GetKey);
                            sqlParameter.Add("McImagePath", !WinMcRegularSub.McImageFile.Equals("") ? "/ImageData/" + ForderName + "/" + strID + "/" : "");
                            sqlParameter.Add("McComments", WinMcRegularSub.McComments);
                            sqlParameter.Add("UpdateUserID", MainWindow.CurrentUser);

                            //System.Diagnostics.Debug.WriteLine("===============" + WinMcRegularSub.McImageFile.ToString());



                            if (!WinMcRegularSub.McImageFile.Replace(" ", "").Equals(""))
                            {
                                string[] FtpFilePathAndName = new string[2];
                                FtpFilePathAndName[0] = WinMcRegularSub.McImageFile;
                                FtpFilePathAndName[1] = WinMcRegularSub.LocalImagePath;
                                FtpFileList.Add(FtpFilePathAndName);
                            }

                            Procedure pro2 = new Procedure();
                            pro2.Name = "xp_McRegularInspectBasis_uMcRegularInspectBasisSub";
                            pro2.OutputUseYN = "N";
                            pro2.OutputName = "McInspectBasisID";
                            pro2.OutputLength = "10";

                            Prolist.Add(pro2);
                            ListParameter.Add(sqlParameter);
                        }

                        // 만약에 삭제할 데이터가 있다면
                        if (DelItems.Count > 0)
                        {

                            for (int i = 0; i < DelItems.Count; i++)
                            {
                                var DeleteData = DelItems[i];

                                sqlParameter = new Dictionary<string, object>();
                                sqlParameter.Clear();
                                sqlParameter.Add("McInspectBasisID", DeleteData.McInspectBasisID);
                                sqlParameter.Add("McSeq", DeleteData.McSeq);

                                Procedure pro3 = new Procedure();
                                pro3.Name = "xp_McRegularInspectBasis_dMcRegularInspectBasisSub";
                                pro3.OutputUseYN = "N";
                                pro3.OutputName = "McInspectBasisID";
                                pro3.OutputLength = "10";

                                Prolist.Add(pro3);
                                ListParameter.Add(sqlParameter);
                            }

                            DelItems.Clear();
                        }

                        string[] Confirm = new string[2];
                        Confirm = DataStore.Instance.ExecuteAllProcedureOutputNew(Prolist, ListParameter);
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
                    #endregion

                    // 파일을 올리자 : GetKey != "" 라면 파일을 올려보자
                    if (!GetKey.Trim().Equals(""))
                    {

                        //추가한 사진이 있을 때
                        if (listFtpFile.Count > 0)
                        {
                            FTP_Save_File(listFtpFile, GetKey);
                        }

                        //복사추가 했을 떄 
                        if (lstFtpFilePath.Count > 0)
                        {
                            FTP_Save_FileByFtpServerFilePath(lstFtpFilePath, GetKey);
                        }
                    }

                    // 파일 List 비워주기
                    listFtpFile.Clear();
                    //deleteListFtpFile.Clear();
                    lstFtpFilePath.Clear();
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

        #endregion 저장(SaveData)

        #region 체크데이터(CheckData)

        /// <summary>
        /// 입력사항 체크
        /// </summary>
        /// <returns></returns>
        private bool CheckData()
        {
            bool flag = true;

            if (txtMcPart.Text.Length <= 0 || txtMcPart.Tag.ToString().Trim().Equals("") || txtMcPart.Tag == null)
            {
                MessageBox.Show("설비가 입력되지 않았습니다.");
                flag = false;
                return flag;
            }

            if (dtpRevision.SelectedDate == null)
            {
                MessageBox.Show("개정일자가 선택되지 않았습니다.");
                flag = false;
                return flag;
            }

            // 개정일자 + 설비(MCID) 로 해당 데이터가 있는지 체크하기



            //// 설비명 중복 체크 
            //if (strFlag.Equals("I") && ChkMcName(PRD_MCNAME, txtDepartID.Text) == false)
            //{
            //    MessageBox.Show("입력하신 설비명이 이미 존재합니다.");
            //    flag = false;
            //    return flag;
            //}


            // 설비명 중복체크
            if (txtMcPart.Tag != null)
            {
                string MCID = txtMcPart.Tag.ToString();
                if (strFlag.Trim().Equals("I")
                    && !ChkMcName(MCID))
                {
                    MessageBox.Show(txtMcPart.Text + " 설비의 설비점검 기준이 이미 존재 합니다.\r\n(실비명을 변경하신 경우, 엔터를 누르시거나, 오른쪽 플러스파인더 버튼을 통해서 설비를 변경해주세요.)", "설비 중복 오류");
                    flag = false;
                    return flag;
                }
            }
           


            return flag;

         

        }


        #region 설비명 찾기
        private bool ChkMcName(string MCID)
        {
            bool flag = true;

            try
            {
                Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                sqlParameter.Clear();

                sqlParameter.Add("MCID", MCID);

                DataSet ds = DataStore.Instance.ProcedureToDataSet("xp_prd_sChkMcname", sqlParameter, false);

                if (ds != null && ds.Tables.Count > 0)
                {
                    DataTable dt = ds.Tables[0];
                    int i = 0;

                    if (dt.Rows.Count != 0)
                    {
                        DataRowCollection drc = dt.Rows;
                        DataRow dr = drc[0];
                        int Cnt = ConvertInt(dr["Cnt"].ToString());

                        if (Cnt > 0)
                        {
                            return false;
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

        #endregion

        #region 개정일자 + 설비(MCID) 로 해당 데이터가 있는지 체크하기

        private bool CheckIsDateAndMCID()
        {
            //bool flag = false;

            //DataSet ds = null;
            //Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
            //sqlParameter.Clear();
            //sqlParameter.Add("chkStandardNumberSrh", chkStandardNumberSrh.IsChecked == true ? 1 : 0);
            //sqlParameter.Add("McInspectBasisID", chkStandardNumberSrh.IsChecked == true ? txtStandardNumberSrh.Text : "");
            //sqlParameter.Add("MCID", chkMcPartNameSrh.IsChecked == true && !txtMcPartNameSrh.Text.ToString().Trim().Equals("") ? txtMcPartNameSrh.Tag.ToString() : "");
            //ds = DataStore.Instance.ProcedureToDataSet("xp_McReqularInspectBasis_sMcReqularInspectBasis", sqlParameter, false);

            //if (ds != null && ds.Tables.Count > 0)
            //{
            //    DataTable dt = ds.Tables[0];
            //    int i = 0;

            //    if (dt.Rows.Count > 0)
            //    { 
            //        DataRowCollection drc = dt.Rows;

            //        foreach (DataRow dr in drc)
            //        {
            //            var WinMCRegul = new Win_prd_RegularInspectBasis_U_CodeView()
            //            {
            //                Num = i + 1,
            //                McInspectBasisID = dr["McInspectBasisID"].ToString(),
            //                ManagerID = dr["ManagerID"].ToString(),
            //                McName = dr["McName"].ToString(),
            //                McInsBasisDate = dr["McInsBasisDate"].ToString(),
            //                McInsContent = dr["McInsContent"].ToString(),
            //                Comments = dr["Comments"].ToString(),
            //                MCID = dr["MCID"].ToString()
            //            };

            //            if (WinMCRegul.McInsBasisDate != null && !WinMCRegul.McInsBasisDate.Replace(" ", "").Equals(""))
            //            {
            //                WinMCRegul.McInsBasisDate = Lib.Instance.StrDateTimeBar(WinMCRegul.McInsBasisDate);
            //            }

            //            dgdMain.Items.Add(WinMCRegul);
            //            i++;
            //        }
            //    }
            //}

            return true;
        }

        #endregion


        #endregion 체크데이터(CheckData)

        #region 하단 그리드 실추가/실삭제 이게 뭘까

        //하단 그리드 실추가
        private void SubPlus()
        {
            int index = dgdSub.Items.Count;

            var WinMCRgl = new Win_prd_RegularInspectBasis_U_CodeView()
            {
                Num = index + 1,
                McInspectBasisID = "",
                McSeq = index + 1,
                McComments = "",
                McInsContent = "",
                McInsCycleDate = "",
                McInsCycleGbn = "",
                McInsRecordGbn = "",
                McInsGbn = "",
                McInsItemName = "",
                McImageFile = "",
                McImagePath = "",
                McInsCheckGbn = "",
                McInsCheck = "",
                McInsCycle = "",
                McInsRecord = ""
            };
            dgdSub.Items.Add(WinMCRgl);
        }

        //하단 그리드 실삭제
        private void SubRemove()
        {
            if (dgdSub.Items.Count > 0)
            {
                if (dgdSub.CurrentItem != null)
                {
                    dgdSub.Items.Remove((dgdSub.Items[dgdSub.Items.Count - 1]) as Win_prd_RegularInspectBasis_U_CodeView);
                }
                else
                {
                    DelItems.Add(dgdSub.SelectedItem as Win_prd_RegularInspectBasis_U_CodeView);
                    dgdSub.Items.Remove((dgdSub.SelectedItem) as Win_prd_RegularInspectBasis_U_CodeView);
                }

                dgdSub.Refresh();
            }
        }

        #endregion 


        #region 삭제(DeleteData)

        /// <summary>
        /// 실삭제
        /// </summary>
        /// <param name="strID"></param>
        /// <returns></returns>
        private bool DeleteData(string strID)
        {
            bool flag = false;

            Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
            sqlParameter.Clear();
            sqlParameter.Add("McInspectBasisID", strID);

            string[] result = DataStore.Instance.ExecuteProcedure_NewLog("xp_McReqularInspectBasis_dMcReqularInspectBasis", sqlParameter, "D");
            DataStore.Instance.CloseConnection();

            if (result[0].Equals("success"))
            {
                //MessageBox.Show("성공 *^^*");
                flag = true;
            }

            return flag;
        }

        #endregion 삭제(DeleteData)

        #region 재조회(re_Search)

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
                this.DataContext = null;

                dgdSub.Items.Clear();

                MessageBox.Show("조호된 데이터가 없습니다.");
                return;
            }
        }

        #endregion 재조회(re_Search)

        #region 조회(FillGrid)

        /// <summary>
        /// 실조회
        /// </summary>
        private void FillGrid()
        {
            if (dgdMain.Items.Count > 0)
                dgdMain.Items.Clear();

            try
            {
                DataSet ds = null;
                Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                sqlParameter.Clear();
                sqlParameter.Add("chkStandardNumberSrh", chkStandardNumberSrh.IsChecked == true ? 1 : 0);
                sqlParameter.Add("McInspectBasisID", chkStandardNumberSrh.IsChecked == true ? txtStandardNumberSrh.Text : "");
                sqlParameter.Add("MCID", chkMcPartNameSrh.IsChecked == true && !txtMcPartNameSrh.Text.ToString().Trim().Equals("") ? txtMcPartNameSrh.Tag.ToString() : "");
                ds = DataStore.Instance.ProcedureToDataSet_LogWrite("xp_McReqularInspectBasis_sMcReqularInspectBasis", sqlParameter, true, "R");

                if (ds != null && ds.Tables.Count > 0)
                {
                    DataTable dt = ds.Tables[0];
                    int i = 0;

                    if (dt.Rows.Count > 0)
                    {
                        DataRowCollection drc = dt.Rows;

                        foreach (DataRow dr in drc)
                        {
                            var WinMCRegul = new Win_prd_RegularInspectBasis_U_CodeView()
                            {
                                Num = i + 1,
                                McInspectBasisID = dr["McInspectBasisID"].ToString(),
                                ManagerID = dr["ManagerID"].ToString(),
                                McName = dr["McName"].ToString(),
                                McInsBasisDate = dr["McInsBasisDate"].ToString(),
                                McInsContent = dr["McInsContent"].ToString(),
                                Comments = dr["Comments"].ToString(),
                                MCID = dr["MCID"].ToString()
                            };

                            if (WinMCRegul.McInsBasisDate != null && !WinMCRegul.McInsBasisDate.Replace(" ", "").Equals(""))
                            {
                                WinMCRegul.McInsBasisDate = Lib.Instance.StrDateTimeBar(WinMCRegul.McInsBasisDate);
                            }

                            dgdMain.Items.Add(WinMCRegul);
                            i++;
                        }
                        tbkCount.Text = "▶ 검색결과 : " + i.ToString() + " 건";
                    }
                }
            }
            catch(Exception ex)
            {
                MessageBox.Show("FillGrid 오류 : " + ex.Message);
            }
            finally
            {
                DataStore.Instance.CloseConnection();
            }
        }


        #endregion 조회(FillGrid)

        #region Sub조회(FillGridSub)

        /// <summary>
        /// 서브 조회
        /// </summary>
        /// <param name="strID"></param>
        private void FillGridSub(string strID)
        {
            //FtpFirstFlag = false;
            Message = "";
            lstMsg.Clear();

            if (dgdSub.Items.Count > 0)
                dgdSub.Items.Clear();

            if (lstExistFtpFile.Count > 0)
            {
                lstExistFtpFile.Clear();
            }

            try
            {
                DataSet ds = null;
                Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                sqlParameter.Clear();
                sqlParameter.Add("McInspectBasisID", strID);
                sqlParameter.Add("McSeq", 0);
                ds = DataStore.Instance.ProcedureToDataSet
                    ("xp_McReqularInspectBasis_sMcReqularInspectBasisSub", sqlParameter, false);

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
                            var WinMCRegulSub = new Win_prd_RegularInspectBasis_U_CodeView()
                            {
                                Num = i + 1,
                                McInspectBasisID = dr["McInspectBasisID"].ToString(),
                                McSeq = ConvertInt(dr["McSeq"].ToString()),
                                McComments = dr["McComments"].ToString(),
                                McInsContent = dr["McInsContent"].ToString(),
                                McInsCycleDate = dr["McInsCycleDate"].ToString(),
                                McInsCycleGbn = dr["McInsCycleGbn"].ToString(),
                                McInsRecordGbn = dr["McInsRecordGbn"].ToString(),
                                McInsGbn = dr["McInsGbn"].ToString(),
                                McInsItemName = dr["McInsItemName"].ToString(),
                                McImageFile = dr["McImageFile"].ToString(),
                                McImagePath = dr["McImagePath"].ToString(),
                                McInsCheckGbn = dr["McInsCheckGbn"].ToString(),
                                McInsCheck = dr["McInsCheck"].ToString(),
                                McInsCycle = dr["McInsCycle"].ToString(),
                                McInsRecord = dr["McInsRecord"].ToString()
                            };

                            if (strCopy)
                            {
                                WinMcRegularSub.McImagePath = "";
                                WinMcRegularSub.McImageFile = "";
                            }
                            else
                            {
                                if (!WinMCRegulSub.McImageFile.Replace(" ", "").Equals(""))
                                {
                                    if (Lib.Instance.Right(WinMCRegulSub.McImageFile, 3).Equals("pdf"))
                                    {
                                        WinMCRegulSub.imageFlag = true;
                                    }
                                    else
                                    {
                                        lstExistFtpFile.Add(WinMCRegulSub.McImageFile);
                                        WinMCRegulSub.imageFlag = true;

                                        if (CheckImage(WinMCRegulSub.McImageFile.Trim()))
                                        {
                                            string strImage = "/" + WinMCRegulSub.McImageFile;

                                            WinMCRegulSub.ImageView = SetImage(strImage, WinMCRegulSub.McInspectBasisID);
                                        }
                                        else
                                        {
                                            MessageBox.Show(WinMCRegulSub.McImageFile + "는 이미지 변환이 불가능합니다.");
                                        }
                                    }
                                }
                                else
                                {
                                    WinMCRegulSub.imageFlag = false;
                                }
                            }

                            dgdSub.Items.Add(WinMCRegulSub);
                            i++;
                        }

                        if (!Message.Trim().Equals(""))
                        {
                            MessageBox.Show(Message + " 의 이미지를 불러올 수 없습니다.");
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

        #endregion Sub조회(FillGridSub)

        #region FTP_이미지

        private BitmapImage SetImage(string ImageName, string FolderName)
        {
            BitmapImage bit = null;
            _ftp = new FTP_EX(FTP_ADDRESS, FTP_ID, FTP_PASS);
            if (_ftp == null) { return null; }

            bit = DrawingImageByByte2(FTP_ADDRESS + '/' + FolderName + '/' + ImageName + "", ImageName);

            return bit;
        }

        /// <summary>
        /// ftp경로를 가지고 Bitmap 정보 리턴한다
        /// </summary>
        /// <param name="ftpFilePath"></param>
        /// <returns></returns>
        private BitmapImage DrawingImageByByte2(string ftpFilePath, string ImageName)
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
                //System.Windows.MessageBox.Show("1" + ex.Message);

                //if (FtpFirstFlag == false)
                //{

                //    //throw ex;
                //    FtpFirstFlag = true;
                //}

                if (Message.Trim().Equals(""))
                {
                    Message += ImageName.Substring(1, ImageName.Length - 1);
                    lstMsg.Add(ImageName);
                }
                else
                {
                    if (!lstMsg.Contains(ImageName))
                    {
                        Message += ", " + ImageName.Substring(1, ImageName.Length - 1);
                        lstMsg.Add(ImageName);
                    }
                }
            }

            return image;
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
                //OFdlg.Filter = "Image files (*.jpg, *.jpeg, *.jpe, *.jfif, *.png, *.pcx, *.pdf) | *.jpg; *.jpeg; *.jpe; *.jfif; *.png; *.pcx; *.pdf | All Files|*.*";
                //OFdlg.Filter = "Image files (*.jpg, *.jpeg, *.jpe, *.jfif, *.png) | *.jpg; *.jpeg; *.jpe; *.jfif; *.png;";
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

                        Bitmap image = new Bitmap(ImageFilePath + ImageFileName);

                        var Hoit = textBox.DataContext as Win_prd_RegularInspectBasis_U_CodeView;
                        Hoit.ImageView = BitmapToImageSource(image);
                        Hoit.imageFlag = true;
                        //MessageBox.Show(Hoit.McInspectBasisID);

                        //imgSetting.Source = BitmapToImageSource(image);

                        string[] strTemp = new string[] { ImageFileName, ImageFilePath.ToString() };
                        listFtpFile.Add(strTemp);
                    }
                }
            }
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
                    file.Delete();
                }

                _ftp.download(str_remotepath.Substring(str_remotepath.Substring
                    (0, str_remotepath.LastIndexOf("/")).LastIndexOf("/")), str_localpath);

                ProcessStartInfo proc = new ProcessStartInfo(str_localpath);
                proc.UseShellExecute = true;
                Process.Start(proc);
            }
        }

        private void FTP_DownLoadFile2(string FolderName, string ImageName)
        {
            // 접속 경로
            _ftp = new FTP_EX(FTP_ADDRESS, FTP_ID, FTP_PASS);

            string str_path = string.Empty;
            str_path = FTP_ADDRESS + '/' + FolderName;
            _ftp = new FTP_EX(str_path, FTP_ID, FTP_PASS);

            string str_remotepath = ImageName;
            string str_localpath = LOCAL_DOWN_PATH + "\\" + ImageName;

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

            if (!_ftp.UploadTempFilesToFTP(UpdateFilesInfo))
            {
                MessageBox.Show("파일업로드에 실패하였습니다.");
                return;
            }
        }

        #endregion // FTP_Save_File - 파일 저장, 폴더 생성

        #endregion FTP_이미지


        #endregion 주요 매서드들

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


        #region 생성자들

        class Win_prd_RegularInspectBasis_U_CodeView : BaseView
        {
            public int Num { get; set; }
            public string McInspectBasisID { get; set; }
            public string ManagerID { get; set; }
            public string McName { get; set; }
            public string MCID { get; set; }
            public string McInsBasisDate { get; set; }
            public string McInsContent { get; set; }
            public string Comments { get; set; }
            public string CreateDate { get; set; }
            public string CreateUserID { get; set; }
            public string LastUpdateDate { get; set; }
            public string LastUpdateUserID { get; set; }

            public int McSeq { get; set; }
            public string McComments { get; set; }
            public string McInsCycleDate { get; set; }
            public string McInsCycleGbn { get; set; }
            public string McInsRecordGbn { get; set; }
            public string McInsGbn { get; set; }
            public string McInsItemName { get; set; }
            public string McImageFile { get; set; }
            public string McImagePath { get; set; }
            public string McInsCheckGbn { get; set; }
            public string McInsCheck { get; set; }
            public string McInsCycle { get; set; }
            public string McInsRecord { get; set; }

            public string LocalImagePath { get; set; } //이건 뭐하는 걸까?

            public BitmapImage ImageView { get; set; }
            public bool imageFlag { get; set; }
        }

        //개정일자 캘린더 열기
        private void DtpRevision_PreviewKeyDown(object sender, KeyEventArgs e)
        {
            if(e.Key == Key.Enter)
            {
                dtpRevision.IsDropDownOpen = true;
            }
        }

        //캘린더 닫힐 때 
        private void DtpRevision_CalendarClosed(object sender, RoutedEventArgs e)
        {
            txtRevisionContents.Focus();
        }

        //개정내용 -> 비고
        private void TxtRevisionContents_KeyDown(object sender, KeyEventArgs e)
        {
            if(e.Key == Key.Enter)
            {
                txtContents.Focus();

            }
        }


        #endregion 생성자들

        //언젠가는 이미지 삭제를 해야겠지...........
        //private void dgdtpetxtImage_KeyUp(object sender, KeyEventArgs e)
        //{
        //    if ((sender as TextBox).Text.Trim().Equals(""))
        //    {
        //        MessageBox.Show("너 삭제임");
        //    }
        //}
    }
}

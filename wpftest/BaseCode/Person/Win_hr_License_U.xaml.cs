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
using WizMes_WellMade.PopUp;

namespace WizMes_WellMade
{
    /// <summary>
    /// Win_hr_License_U.xaml에 대한 상호 작용 논리
    /// </summary>
    public partial class Win_hr_License_U : UserControl
    {
        string strFlag = string.Empty;
        int rowNum = 0;
        Win_hr_License_U_CodeView WinLicense = new Win_hr_License_U_CodeView();
        Lib lib = new Lib();

        // FTP 활용모음.
        List<string[]> listFtpFile = new List<string[]>();
        private FTP_EX _ftp = null;
        string FullPath1 = string.Empty;

        //string FTP_ADDRESS = "ftp://wizis.iptime.org/ImageData/License";
        //string FTP_ADDRESS = "ftp://wizis.iptime.org/ImageData/Draw";
        //string FTP_ADDRESS = "ftp://222.104.222.145:25000/ImageData/License";
        //string FTP_ADDRESS = "ftp://192.168.0.95/ImageData/License";

        string FTP_ADDRESS = "ftp://" + LoadINI.FileSvr + ":"
            + LoadINI.FTPPort + LoadINI.FtpImagePath + "/License";

        string StringImagePath = LoadINI.FtpImagePath + "/License/";


        private const string FTP_ID = "wizuser";
        private const string FTP_PASS = "wiz9999";
        private const string LOCAL_DOWN_PATH = "C:\\Temp";

        public Win_hr_License_U()
        {
            InitializeComponent();
        }

        private void UserControl_Loaded(object sender, RoutedEventArgs e)
        {
            Lib.Instance.UiLoading(sender);
            SetComboBox();
            cboLicenseSrh.SelectedIndex = 0;
            btnPfEmployeeSrh.IsEnabled = false;
            CanBtnControl();
        }

        private void SetComboBox()
        {
            //품목 용도( 입력)
            ObservableCollection<CodeView> ovcLicense = ComboBoxUtil.Instance.Gf_DB_CM_GetComCodeDataset(null, "HRLICE", "Y", "");
            this.cboLicenseSrh.ItemsSource = ovcLicense;
            this.cboLicenseSrh.DisplayMemberPath = "code_name";
            this.cboLicenseSrh.SelectedValuePath = "code_id";

            this.cboLicense.ItemsSource = ovcLicense;
            this.cboLicense.DisplayMemberPath = "code_name";
            this.cboLicense.SelectedValuePath = "code_id";
        }

        //사원명
        private void lblEmployee_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            if (chkEmployee.IsChecked == true) { chkEmployee.IsChecked = false; }
            else { chkEmployee.IsChecked = true; }
        }

        //사원명
        private void chkEmployee_Checked(object sender, RoutedEventArgs e)
        {
            txtEmployeeSrh.IsEnabled = true;
            btnPfEmployeeSrh.IsEnabled = true;
            txtEmployeeSrh.Focus();
        }

        //사원명
        private void chkEmployee_Unchecked(object sender, RoutedEventArgs e)
        {
            txtEmployeeSrh.IsEnabled = false;
            btnPfEmployeeSrh.IsEnabled = false;
        }

        //사원명
        private void txtEmployeeSrh_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                MainWindow.pf.ReturnCode(txtEmployeeSrh, (int)Defind_CodeFind.DCF_PERSON, "");
            }
        }

        //사원명
        private void btnPfEmployeeSrh_Click(object sender, RoutedEventArgs e)
        {
            MainWindow.pf.ReturnCode(txtEmployeeSrh, (int)Defind_CodeFind.DCF_PERSON, "");
        }

        //자격증명
        private void lblLicense_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            if (chkLicense.IsChecked == true) { chkLicense.IsChecked = false; }
            else { chkLicense.IsChecked = true; }
        }

        //자격증명
        private void chkLicense_Checked(object sender, RoutedEventArgs e)
        {
            cboLicenseSrh.IsEnabled = true;
            cboLicenseSrh.Focus();
        }

        //자격증명
        private void chkLicense_Unchecked(object sender, RoutedEventArgs e)
        {
            cboLicenseSrh.IsEnabled = false;
        }

        //퇴사자 포함
        private void lblResigner_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            if (chkResigner.IsChecked == true) { chkResigner.IsChecked = false; }
            else { chkResigner.IsChecked = true; }
        }

        private void chkResigner_Checked(object sender, RoutedEventArgs e)
        {

        }

        private void chkResigner_Unchecked(object sender, RoutedEventArgs e)
        {

        }

        /// <summary>
        /// 수정,추가 저장 후
        /// </summary>
        private void CanBtnControl()
        {
            Lib.Instance.UiButtonEnableChange_IUControl(this);
            //grdInput.IsEnabled = false;
            //grdInput.IsHitTestVisible = false;

            // 수정, 추가 과정 없이도 FTP 보기는 가능해야 한다.
            btnFtpSeeLicense.IsHitTestVisible = true;

            txtEmployee.IsHitTestVisible = false;
            btnPfEmployee.IsHitTestVisible = false;
            cboLicense.IsHitTestVisible = false;
            dtpReceiveDate.IsHitTestVisible = false;
            btnFtpLicenseUpload.IsHitTestVisible = false;
            txtLicense.IsHitTestVisible = false;
            txtComments.IsHitTestVisible = false;
        }

        /// <summary>
        /// 수정,추가 진행 중
        /// </summary>
        private void CantBtnControl()
        {
            Lib.Instance.UiButtonEnableChange_SCControl(this);
            //grdInput.IsEnabled = true;
            //grdInput.IsHitTestVisible = true;

            btnFtpSeeLicense.IsHitTestVisible = true;

            txtEmployee.IsHitTestVisible = true;
            btnPfEmployee.IsHitTestVisible = true;
            cboLicense.IsHitTestVisible = true;
            dtpReceiveDate.IsHitTestVisible = true;
            btnFtpLicenseUpload.IsHitTestVisible = true;
            txtLicense.IsHitTestVisible = true;
            txtComments.IsHitTestVisible = true;

        }

        //추가
        private void btnAdd_Click(object sender, RoutedEventArgs e)
        {
            CantBtnControl();
            tbkMsg.Text = "자료 입력 중";
            strFlag = "I";
            rowNum = dgdMain.SelectedIndex;
            this.DataContext = null;

            dtpReceiveDate.SelectedDate = DateTime.Today;
            cboLicense.SelectedIndex = 0;

            // 사원으로 포커스 두자.
            txtEmployee.Focus();
        }

        //수정
        private void btnUpdate_Click(object sender, RoutedEventArgs e)
        {
            WinLicense = dgdMain.SelectedItem as Win_hr_License_U_CodeView;

            if (WinLicense != null)
            {
                rowNum = dgdMain.SelectedIndex;
                //dgdMain.IsEnabled = false;
                dgdMain.IsHitTestVisible = false;
                tbkMsg.Text = "자료 수정 중";
                CantBtnControl();
                strFlag = "U";
            }
        }

        //삭제
        private void btnDelete_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                WinLicense = dgdMain.SelectedItem as Win_hr_License_U_CodeView;

                if (WinLicense == null)
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

                        if (DeleteData(WinLicense.PersonID, WinLicense.LicenseSeq))
                        {
                            rowNum -= 1;
                            re_Search(rowNum);
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
        private void btnClose_Click(object sender, RoutedEventArgs e)
        {
            Lib.Instance.ChildMenuClose(this.ToString());
        }

        //검색
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
            if (SaveData(strFlag))
            {
                listFtpFile.Clear();
                CanBtnControl();

                if (strFlag == "I")
                {
                    re_Search(dgdMain.Items.Count - 1);
                }
                else
                {
                    re_Search(rowNum);
                }
                dgdMain.IsHitTestVisible = true;
                strFlag = string.Empty;
            }
        }

        //취소
        private void btnCancel_Click(object sender, RoutedEventArgs e)
        {
            listFtpFile.Clear();
            CanBtnControl();

            if (!strFlag.Equals(string.Empty))
            {
                re_Search(rowNum);
            }

            strFlag = string.Empty;
            //dgdMain.IsEnabled = true;
            dgdMain.IsHitTestVisible = true;
        }

        //엑셀
        private void btnExcel_Click(object sender, RoutedEventArgs e)
        {
            DataTable dt = null;
            string Name = string.Empty;
            Lib lib = new Lib();

            string[] lst = new string[2];
            lst[0] = "작업자별 자격증";
            lst[1] = dgdMain.Name;

            ExportExcelxaml ExpExc = new ExportExcelxaml(lst);
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
                if (chkEmployee.IsChecked == true && txtEmployeeSrh.Tag == null)
                {
                    MessageBox.Show("사원을 선택해주세요.");
                    return;
                }

                DataSet ds = null;
                Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                sqlParameter.Clear();
                sqlParameter.Add("chkPersonID", chkEmployee.IsChecked == true ? 1 : 0);
                sqlParameter.Add("sPersonID", chkEmployee.IsChecked == true ? txtEmployeeSrh.Tag.ToString() : "");
                sqlParameter.Add("chkLicenseSeq", 0);
                sqlParameter.Add("nLicenseSeq", 0);
                sqlParameter.Add("chkLicenseID", chkLicense.IsChecked == true ? 1 : 0);
                sqlParameter.Add("sLicenseID", chkLicense.IsChecked == true ?
                    cboLicenseSrh.SelectedValue.ToString() : "");
                sqlParameter.Add("chkIncRetired", chkResigner.IsChecked == true ? 1 : 0);

                ds = DataStore.Instance.ProcedureToDataSet("xp_HRLice_sLicense", sqlParameter, false);

                if (ds != null && ds.Tables.Count > 0)
                {
                    DataTable dt = ds.Tables[0];
                    int i = 0;

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
                            var WinhrLicense = new Win_hr_License_U_CodeView()
                            {
                                Num = i + 1,
                                Comments = dr["Comments"].ToString(),
                                EndDate = dr["EndDate"].ToString(),
                                LicenseFile = dr["LicenseFile"].ToString(),
                                LicenseFilePath = dr["LicenseFilePath"].ToString(),
                                LicenseID = dr["LicenseID"].ToString(),
                                LicenseName = dr["LicenseName"].ToString(),
                                LicenseSeq = dr["LicenseSeq"].ToString(),
                                Name = dr["Name"].ToString(),
                                PersonID = dr["PersonID"].ToString(),
                                ReceiveDate = dr["ReceiveDate"].ToString()
                            };

                            if (WinhrLicense.ReceiveDate.Length > 0)
                            {
                                WinhrLicense.ReceiveDate_CV = Lib.Instance.StrDateTimeBar(WinhrLicense.ReceiveDate);
                            }

                            if (WinhrLicense.EndDate.Replace(" ", "").ToString().Equals(""))
                            {
                                WinhrLicense.RetireYN = "N";
                            }
                            else
                            {
                                WinhrLicense.RetireYN = "Y";
                            }
                            if ((WinhrLicense.ReceiveDate.Trim() != "" && WinhrLicense.ReceiveDate != null))
                            {
                                WinhrLicense.ReceiveDate = WinhrLicense.ReceiveDate.ToString().Substring(0, 4) + "-"
                              + WinhrLicense.ReceiveDate.ToString().Substring(4, 2) + "-"
                              + WinhrLicense.ReceiveDate.ToString().Substring(6, 2);
                            }
                            dgdMain.Items.Add(WinhrLicense);
                            i++;
                        }

                        tbkIndexCount.Text = "▶ 검색결과 : " + i + "건";
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
        private void dgdMain_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            WinLicense = dgdMain.SelectedItem as Win_hr_License_U_CodeView;

            if (WinLicense != null)
            {
                this.DataContext = WinLicense;
            }
        }

        /// <summary>
        /// 실삭제
        /// </summary>
        /// <param name="strID"></param>
        /// <param name="strLicenseSeq"></param>
        /// <returns></returns>
        private bool DeleteData(string strID, string strLicenseSeq)
        {
            bool flag = false;

            try
            {
                Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                sqlParameter.Clear();
                sqlParameter.Add("PersonID", strID);
                sqlParameter.Add("LicenseSeq", strLicenseSeq);

                string[] result = DataStore.Instance.ExecuteProcedure("xp_HRLice_dLicense", sqlParameter, false);

                if (result[0].Equals("success"))
                {
                    //MessageBox.Show("성공 *^^*");
                    flag = true;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("오류 발생,오류 내용 : " + ex.ToString());
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
        /// <param name="strYYYY"></param>
        /// <returns></returns>
        private bool SaveData(string strFlag)
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
                    sqlParameter.Add("PersonID", txtEmployee.Tag.ToString());
                    sqlParameter.Add("LicenseID", cboLicense.SelectedValue.ToString());
                    sqlParameter.Add("ReceiveDate", dtpReceiveDate.SelectedDate.Value.ToString("yyyyMMdd"));
                    sqlParameter.Add("LicenseFilePath", StringImagePath.ToString() + txtEmployee.Tag.ToString());  //txtLicense.Tag ==null ? "": strImagePath
                    sqlParameter.Add("LicenseFile", "");
                    sqlParameter.Add("Comments", txtComments.Text);

                    #region 추가

                    if (strFlag.Equals("I"))
                    {
                        sqlParameter.Add("LicenseSeq", 0);
                        sqlParameter.Add("CreateUserID", MainWindow.CurrentUser);

                        Procedure pro1 = new Procedure();
                        pro1.Name = "xp_HRLice_iLicense";
                        pro1.OutputUseYN = "N";
                        pro1.OutputName = "PersonID";
                        pro1.OutputLength = "10";

                        Prolist.Add(pro1);
                        ListParameter.Add(sqlParameter);

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
                        }

                        if (flag)
                        {
                            bool AttachYesNo = false;
                            if (txtLicense.Text != string.Empty)       //첨부파일 1
                            {
                                if (FTP_Save_File(listFtpFile, txtEmployee.Tag.ToString()))
                                {
                                    //txtLicense.Tag = "/ImageData/License/" + txtEmployee.Tag.ToString();
                                    AttachYesNo = true;
                                }
                                else
                                { MessageBox.Show("데이터 저장이 완료되었지만, 첨부문서 등록에 실패하였습니다."); }
                            }
                            if (AttachYesNo == true) { AttachFileUpdate(txtEmployee.Tag.ToString()); }      //첨부문서 정보 DB 업데이트.
                        }
                    }

                    #endregion

                    #region 수정

                    else if (strFlag.Equals("U"))
                    {
                        sqlParameter.Add("LicenseSeq", WinLicense.LicenseSeq);
                        sqlParameter.Add("UserID", MainWindow.CurrentUser);

                        Procedure pro1 = new Procedure();
                        pro1.Name = "xp_HRLice_uLicense";
                        pro1.OutputUseYN = "N";
                        pro1.OutputName = "PersonID";
                        pro1.OutputLength = "10";

                        Prolist.Add(pro1);
                        ListParameter.Add(sqlParameter);

                        string[] Confirm = new string[2];
                        Confirm = DataStore.Instance.ExecuteAllProcedureOutputNew(Prolist, ListParameter);
                        if (Confirm == null || Confirm[0] != "success")
                        {
                            MessageBox.Show("[저장실패]\r\n" + Confirm[1].ToString());
                            flag = false;
                            //return false;
                        }
                        else
                        {
                            flag = true;
                        }
                        if (flag)
                        {
                            var ViewReceiver = dgdMain.SelectedItem as Win_hr_License_U_CodeView;

                            bool AttachYesNo = false;
                            if (txtLicense.Text != string.Empty)       //첨부파일1 > DB 업로드 조건은 통과
                            {
                                if (txtLicense.Tag.ToString() != ViewReceiver.LicenseFilePath)   // 기존 저장된 경로랑 새로 들어온 경로랑 같지 않을때만,
                                {
                                    if (FTP_Save_File(listFtpFile, txtEmployee.Tag.ToString()))
                                    {
                                        //txtEmployee.Tag = "/ImageData/License/" + txtEmployee.Tag.ToString();
                                        AttachYesNo = true;
                                    }
                                    else
                                    { MessageBox.Show("데이터 수정이 완료되었지만, 첨부문서 등록에 실패하였습니다."); }
                                }
                            }
                            if (AttachYesNo == true) { AttachFileUpdate(txtEmployee.Tag.ToString()); }      //첨부문서 정보 DB 업데이트.
                        }
                    }
                    #endregion
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("오류 발생,오류 내용 : " + ex.ToString());
            }
            finally
            {
                DataStore.Instance.CloseConnection();
            }

            return flag;
        }

        // 1) 첨부문서가 있을경우, 2) FTP에 정상적으로 업로드가 완료된 경우.  >> DB에 정보 업데이트 
        private void AttachFileUpdate(string ID)
        {
            try
            {
                Dictionary<string, object> sqlParameter = new Dictionary<string, object>();

                sqlParameter.Clear();
                sqlParameter.Add("PersonID", ID);

                sqlParameter.Add("LicenseFilePath", StringImagePath.ToString() + txtEmployee.Tag.ToString());
                sqlParameter.Add("LicenseFile", txtLicense.Text);

                string[] result = DataStore.Instance.ExecuteProcedure("xp_HRLice_uLicense_Ftp", sqlParameter, false);
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




        /// <summary>
        /// 입력사항 체크
        /// </summary>
        /// <returns></returns>
        private bool CheckData()
        {
            bool flag = true;

            if (cboLicense.SelectedValue == null)
            {
                MessageBox.Show("자격증이 선택되지 않았습니다.");
                flag = false;
                return flag;
            }

            if (txtEmployee.Text.Length <= 0 || txtEmployee.Text.Equals(""))
            {
                MessageBox.Show("사원이 입력되지 않았습니다.");
                flag = false;
                return flag;
            }

            return flag;
        }

        //사원명
        private void txtEmployee_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                MainWindow.pf.ReturnCode(txtEmployee, (int)Defind_CodeFind.DCF_PERSON, "");
                cboLicense.Focus();
                cboLicense.IsDropDownOpen = true;
            }
        }

        //사원명
        private void btnPfEmployee_Click(object sender, RoutedEventArgs e)
        {
            MainWindow.pf.ReturnCode(txtEmployee, (int)Defind_CodeFind.DCF_PERSON, "");
            cboLicense.Focus();
            cboLicense.IsDropDownOpen = true;
        }




        // 파일 첨부등록
        private void btnFtpLicenseUpload_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                Microsoft.Win32.OpenFileDialog OFdlg = new Microsoft.Win32.OpenFileDialog();

                OFdlg.DefaultExt = ".jpg";
                OFdlg.Filter = "Image files (*.jpg, *.jpeg, *.jpe, *.jfif, *.png) | *.jpg; *.jpeg; *.jpe; *.jfif; *.png | All Files|*.*";

                Nullable<bool> result = OFdlg.ShowDialog();
                if (result == true)
                {
                    FullPath1 = OFdlg.FileName;  //긴 경로(FULL 사이즈)

                    string AttachFileName = OFdlg.SafeFileName;  //명.
                    string AttachFilePath = string.Empty;       // 경로

                    AttachFilePath = FullPath1.Replace(AttachFileName, "");

                    StreamReader sr = new StreamReader(OFdlg.FileName);
                    long File_size = sr.BaseStream.Length;
                    if (sr.BaseStream.Length > (2048 * 1000))
                    {
                        // 업로드 파일 사이즈범위 초과
                        MessageBox.Show("밀시트의 파일사이즈가 2M byte를 초과하였습니다.");
                        sr.Close();
                        return;
                    }

                    txtLicense.Text = AttachFileName;
                    txtLicense.Tag = AttachFilePath.ToString();

                    listFtpFile.Add(new string[] { AttachFileName, AttachFilePath.ToString() });
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
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
            MakeFolder = FolderInfoAndFlag(fileListSimple, MakeFolderName.ToString().Trim());

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
                            MessageBox.Show("동일명칭의 FTP가 이미 등록되어 있습니다.");
                            return true;
                        }
                    }
                }

                if (flag)
                {
                    listStrArrayFileInfo[i][0] = MakeFolderName.Trim() + "/" + listStrArrayFileInfo[i][0];
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


        // 파일 보기 클릭.
        private void btnFtpSeeLicense_Click(object sender, RoutedEventArgs e)
        {
            MessageBoxResult msgresult = MessageBox.Show("파일을 다운로드 하시겠습니까?", "다운로드 확인", MessageBoxButton.YesNo);
            if (msgresult == MessageBoxResult.Yes)
            {
                var ViewReceiver = dgdMain.SelectedItem as Win_hr_License_U_CodeView;

                if (ViewReceiver != null && !ViewReceiver.LicenseFilePath.Equals(""))
                {
                    FTP_DownLoadFile(ViewReceiver.LicenseFilePath.ToString().Trim(), ViewReceiver.PersonID.ToString().Trim(), ViewReceiver.LicenseFile);
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

                ExistFile = FolderInfoAndFlag(fileListSimple, FolderName.ToString().Trim());

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



        private void DgdMain_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            if (e.ClickCount == 2)
            {
                btnUpdate_Click(btnUpdate, null);
            }
        }


        // 자격증명에서 자격증으로 이동(취득일은 자동금일세팅 패스)
        private void cboLicense_DropDownClosed(object sender, EventArgs e)
        {
            txtLicense.Focus();
        }

        private void txtLicense_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                btnFtpLicenseUpload_Click(null, null);
                txtComments.Focus();
            }
        }
        // 비고에서 다시 사원으로 이동(반복)
        private void txtComments_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                txtEmployee.Focus();
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
    }


    class Win_hr_License_U_CodeView : BaseView
    {
        public int Num { get; set; }

        public string PersonID { get; set; }
        public string Name { get; set; }
        public string LicenseSeq { get; set; }
        public string LicenseID { get; set; }
        public string LicenseName { get; set; }

        public string ReceiveDate { get; set; }
        public string LicenseFilePath { get; set; }
        public string LicenseFile { get; set; }
        public string Comments { get; set; }
        public string EndDate { get; set; }

        public string RetireYN { get; set; }
        public string ReceiveDate_CV { get; set; }
    }
}

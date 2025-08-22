using System;
using System.Collections.Generic;
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
    /// Win_hr_Education_U.xaml에 대한 상호 작용 논리
    /// </summary>
    public partial class Win_hr_Education_U : UserControl
    {
        string strFlag = string.Empty;
        int rowNum = 0;
        Win_hr_Education_U_CodeView WinEduMain = new Win_hr_Education_U_CodeView();
        Win_hr_Education_U_Sub_CodeView WinEduSub = new Win_hr_Education_U_Sub_CodeView();
        Lib lib = new Lib();
        // FTP 활용모음.
        string FullPath1 = string.Empty;
        List<string[]> listFtpFile = new List<string[]>();
        private FTP_EX _ftp = null;

        //string FTP_ADDRESS = "ftp://wizis.iptime.org/ImageData/McCode";
        //string FTP_ADDRESS = "ftp://wizis.iptime.org/ImageData/Draw";
        //string FTP_ADDRESS = "ftp://222.104.222.145:25000/ImageData/Education";
        //string FTP_ADDRESS = "ftp://192.168.0.95/ImageData/Education";
        string FTP_ADDRESS = "ftp://" + LoadINI.FileSvr + ":"
            + LoadINI.FTPPort + LoadINI.FtpImagePath + "/Education";
        private const string FTP_ID = "wizuser";
        private const string FTP_PASS = "wiz9999";
        private const string LOCAL_DOWN_PATH = "C:\\Temp";


        public Win_hr_Education_U()
        {
            InitializeComponent();
        }

        private void UserControl_Loaded(object sender, RoutedEventArgs e)
        {
            Lib.Instance.UiLoading(sender);
            chkDate.IsChecked = true;
            dtpSDate.SelectedDate = DateTime.Today;
            dtpEDate.SelectedDate = DateTime.Today;

            CanBtnControl();
        }

        //교육일
        private void lblDate_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            if (chkDate.IsChecked == true) { chkDate.IsChecked = false; }
            else { chkDate.IsChecked = true; }
        }

        //교육일
        private void chkDate_Checked(object sender, RoutedEventArgs e)
        {
            dtpSDate.IsEnabled = true;
            dtpEDate.IsEnabled = true;
        }

        //교육일
        private void chkDate_Unchecked(object sender, RoutedEventArgs e)
        {
            dtpSDate.IsEnabled = false;
            dtpEDate.IsEnabled = false;
        }

        //교육명
        private void lblEduNameSrh_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            if (chkEduNameSrh.IsChecked == true) { chkEduNameSrh.IsChecked = false; }
            else { chkEduNameSrh.IsChecked = true; }
        }

        //교육명
        private void chkEduNameSrh_Checked(object sender, RoutedEventArgs e)
        {
            txtEduNameSrh.IsEnabled = true;
            txtEduNameSrh.Focus();
        }

        //교육명
        private void chkEduNameSrh_Unchecked(object sender, RoutedEventArgs e)
        {
            txtEduNameSrh.IsEnabled = false;
        }

        //사원명
        private void lblPersonSrh_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            if (chkPersonSrh.IsChecked == true) { chkPersonSrh.IsChecked = false; }
            else { chkPersonSrh.IsChecked = true; }
        }

        //사원명
        private void chkPersonSrh_Checked(object sender, RoutedEventArgs e)
        {
            txtPersonSrh.IsEnabled = true;
            txtPersonSrh.Focus();
        }

        //사원명
        private void chkPersonSrh_Unchecked(object sender, RoutedEventArgs e)
        {
            txtPersonSrh.IsEnabled = false;
        }

        /// <summary>
        /// 수정,추가 저장 후
        /// </summary>
        private void CanBtnControl()
        {
            Lib.Instance.UiButtonEnableChange_IUControl(this);
            dgdMain.IsHitTestVisible = true;
            btnSubAdd.IsEnabled = false;
            btnSubDel.IsEnabled = false;

            btnSeeEducation.IsEnabled = true;
            btnSeeEducation.IsHitTestVisible = true;
            btnEducation.IsHitTestVisible = false;
            dtpStartDate.IsHitTestVisible = false;
            dtpEndDate.IsHitTestVisible = false;
            txtEducationName.IsHitTestVisible = false;
            txtEducationContext.IsHitTestVisible = false;
            txtEducationFile.IsHitTestVisible = false;
            txtComments.IsHitTestVisible = false;

        }

        /// <summary>
        /// 수정,추가 진행 중
        /// </summary>
        private void CantBtnControl()
        {
            Lib.Instance.UiButtonEnableChange_SCControl(this);
            dgdMain.IsHitTestVisible = false;
            btnSubAdd.IsEnabled = true;
            btnSubDel.IsEnabled = true;

            btnSeeEducation.IsEnabled = false;
            btnSeeEducation.IsHitTestVisible = false;
            btnEducation.IsHitTestVisible = true;
            dtpStartDate.IsHitTestVisible = true;
            dtpEndDate.IsHitTestVisible = true;
            txtEducationName.IsHitTestVisible = true;
            txtEducationContext.IsHitTestVisible = true;
            txtEducationFile.IsHitTestVisible = true;
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

            //FTP리스트로 깨끗하게 클리어 해주고.
            listFtpFile.Clear();

            if (dgdSub.Items.Count > 0)
            {
                dgdSub.Items.Clear();
            }

            dtpStartDate.SelectedDate = DateTime.Now;
            dtpEndDate.SelectedDate = DateTime.Now;

            // 교육명으로 포커스 시작.
            txtEducationName.Focus();
        }

        //수정
        private void btnUpdate_Click(object sender, RoutedEventArgs e)
        {
            WinEduMain = dgdMain.SelectedItem as Win_hr_Education_U_CodeView;

            if (WinEduMain != null)
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
                WinEduMain = dgdMain.SelectedItem as Win_hr_Education_U_CodeView;

                if (WinEduMain == null)
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

                        if (DeleteData(WinEduMain.EducationID))
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
            int stdate = Convert.ToInt32(dtpStartDate.SelectedDate.Value.ToString("yyyyMMdd"));
            int eddate = Convert.ToInt32(dtpEndDate.SelectedDate.Value.ToString("yyyyMMdd"));

            if (eddate < stdate)
            {
                strFlag = string.Empty;
                MessageBox.Show("교육시작일자가 종료일자보다 클 수 없습니다.");
                return;
            }

            if (SaveData(strFlag))
            {
                CanBtnControl();
                //dgdMain.IsEnabled = true;
                dgdMain.IsHitTestVisible = true;

                if (strFlag == "I")
                {
                    re_Search(dgdMain.Items.Count - 1);
                    dgdMain.Focus();
                }
                else
                {
                    re_Search(rowNum);
                    dgdMain.Focus();
                }
                strFlag = string.Empty;
            }
        }

        //취소
        private void btnCancel_Click(object sender, RoutedEventArgs e)
        {
            CanBtnControl();

            if (!strFlag.Equals(string.Empty))
            {
                re_Search(rowNum);
                dgdMain.Focus();
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

            string[] lst = new string[4];
            lst[0] = "작업자별 교육현황";
            lst[1] = "교육수료 작업자";
            lst[2] = dgdMain.Name;
            lst[3] = dgdSub.Name;

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
                else if (ExpExc.choice.Equals(dgdSub.Name))
                {
                    if (ExpExc.Check.Equals("Y"))
                        dt = lib.DataGridToDTinHidden(dgdSub);
                    else
                        dt = lib.DataGirdToDataTable(dgdSub);

                    Name = dgdSub.Name;

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
            if (dgdSub.Items.Count > 0)
            {
                dgdSub.Items.Clear();
            }

            try
            {
                DataSet ds = null;
                Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                sqlParameter.Clear();
                sqlParameter.Add("nchkEductionID", 0);
                sqlParameter.Add("EductionID", "");
                sqlParameter.Add("nchkDate", chkDate.IsChecked == true ? 1 : 0);
                sqlParameter.Add("StartDate", chkDate.IsChecked == true ?
                    dtpSDate.SelectedDate.Value.ToString("yyyyMMdd") : "");
                sqlParameter.Add("EndDate", chkDate.IsChecked == true ?
                    dtpEDate.SelectedDate.Value.ToString("yyyyMMdd") : "");
                sqlParameter.Add("nchkEducation", chkEduNameSrh.IsChecked == true ? 1 : 0);
                sqlParameter.Add("sEducation", chkEduNameSrh.IsChecked == true ? txtEduNameSrh.Text : "");
                sqlParameter.Add("nchkPersonName", chkPersonSrh.IsChecked == true ? 1 : 0);
                sqlParameter.Add("sPersonName", chkPersonSrh.IsChecked == true ?
                    txtPersonSrh.Text : "");

                ds = DataStore.Instance.ProcedureToDataSet("xp_HREdu_sEducation", sqlParameter, false);

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
                            var WinhrLicense = new Win_hr_Education_U_CodeView()
                            {
                                Num = i + 1,
                                Comments = dr["Comments"].ToString(),
                                EducationContext = dr["EducationContext"].ToString(),
                                EducationEndDate = dr["EducationEndDate"].ToString(),
                                EducationFile = dr["EducationFile"].ToString(),
                                EducationFilePath = dr["EducationFilePath"].ToString(),
                                EducationID = dr["EducationID"].ToString(),
                                EducationName = dr["EducationName"].ToString(),
                                EducationStartDate = dr["EducationStartDate"].ToString()
                            };

                            if (!WinhrLicense.EducationStartDate.Replace(" ", "").Equals(""))
                            {
                                WinhrLicense.EducationStartDate_CV = Lib.Instance.StrDateTimeBar(WinhrLicense.EducationStartDate);
                            }

                            if (!WinhrLicense.EducationEndDate.Replace(" ", "").Equals(""))
                            {
                                WinhrLicense.EducationEndDate_CV = Lib.Instance.StrDateTimeBar(WinhrLicense.EducationEndDate);
                            }

                            WinhrLicense.EduDate = WinhrLicense.EducationStartDate_CV + " ~ " + "\r\n" +
                                WinhrLicense.EducationEndDate_CV;

                            WinhrLicense.PersonCount = FillCount(WinhrLicense.EducationID);


                            dgdMain.Items.Add(WinhrLicense);
                            i++;
                        }
                        tbkIndexCount.Text = "▶검색결과 : " + i + "건";
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



        private void dgdMain_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            WinEduMain = dgdMain.SelectedItem as Win_hr_Education_U_CodeView;

            if (WinEduMain != null)
            {
                FillGridSub(WinEduMain.EducationID);
                this.DataContext = WinEduMain;
            }
        }

        private void FillGridSub(string strID)
        {
            if (dgdSub.Items.Count > 0)
            {
                dgdSub.Items.Clear();
            }

            try
            {
                Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                sqlParameter.Add("nchkEductionID", 1);
                sqlParameter.Add("EductionID", strID);

                DataSet ds = DataStore.Instance.ProcedureToDataSet("xp_HREdu_sEducationSub", sqlParameter, false);

                if (ds != null && ds.Tables.Count > 0)
                {
                    int i = 0;
                    DataTable dt = ds.Tables[0];

                    if (dt.Rows.Count > 0)
                    {
                        DataRowCollection drc = dt.Rows;

                        foreach (DataRow dr in drc)
                        {
                            var WinSub = new Win_hr_Education_U_Sub_CodeView()
                            {
                                Num = (i + 1),
                                Comments = dr["Comments"].ToString(),
                                EducationID = dr["EducationID"].ToString(),
                                EducationSeq = dr["EducationSeq"].ToString(),
                                Name = dr["Name"].ToString(),
                                PersonID = dr["PersonID"].ToString()
                            };

                            dgdSub.Items.Add(WinSub);
                            i++;
                        }
                    }
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
        }

        private int FillCount(string strID)
        {
            int count = 0;

            try
            {
                Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                sqlParameter.Add("nchkEductionID", 1);
                sqlParameter.Add("EductionID", strID);

                DataSet ds = DataStore.Instance.ProcedureToDataSet("xp_HREdu_sEducationSub", sqlParameter, false);

                if (ds != null && ds.Tables.Count > 0)
                {
                    DataTable dt = ds.Tables[0];

                    if (dt.Rows.Count > 0)
                    {
                        count = dt.Rows.Count;
                    }
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

            return count;
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
                sqlParameter.Add("EducationID", strID);

                string[] result = DataStore.Instance.ExecuteProcedure("xp_HREdu_dEducation", sqlParameter, false);

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

                    if (WinEduMain != null)
                    {
                        sqlParameter.Add("EducationID", WinEduMain.EducationID);
                    }
                    else
                    {
                        sqlParameter.Add("EducationID", "");
                    }

                    sqlParameter.Add("EducationStartDate", dtpStartDate.SelectedDate.Value.ToString("yyyyMMdd"));
                    sqlParameter.Add("EducationEndDate", dtpEndDate.SelectedDate.Value.ToString("yyyyMMdd"));
                    sqlParameter.Add("EducationContext", txtEducationContext.Text);
                    sqlParameter.Add("EducationFilePath", "");  // txtEducationFile.Tag != null ? strImagePath:
                    sqlParameter.Add("EducationFile", "");
                    sqlParameter.Add("Comments", txtComments.Text);
                    sqlParameter.Add("EducationName", txtEducationName.Text);

                    #region 추가

                    if (strFlag.Equals("I"))
                    {
                        sqlParameter.Add("CreateUserID", txtComments.Text);

                        Procedure pro1 = new Procedure();
                        pro1.Name = "xp_HREdu_iEducation";
                        pro1.OutputUseYN = "Y";
                        pro1.OutputName = "EducationID";
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
                                if (kv.key == "EducationID")
                                {
                                    sGetID = kv.value;
                                    flag = true;

                                    //메인 그리드 데이터가 중복 저장 되는
                                    Prolist.RemoveAt(0);
                                    ListParameter.Clear();
                                }
                            }

                            if (flag)
                            {
                                bool AttachYesNo = false;
                                if (txtEducationFile.Text != string.Empty)       //첨부파일 1
                                {
                                    if (FTP_Save_File(listFtpFile, sGetID))
                                    {
                                        txtEducationFile.Tag = "/ImageData/Education/" + sGetID;
                                        AttachYesNo = true;
                                    }
                                    else
                                    { MessageBox.Show("데이터 저장이 완료되었지만, 첨부문서 등록에 실패하였습니다."); }
                                }
                                if (AttachYesNo == true) { AttachFileUpdate(sGetID); }      //첨부문서 정보 DB 업데이트.                                
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
                            WinEduSub = dgdSub.Items[i] as Win_hr_Education_U_Sub_CodeView;
                            sqlParameter = new Dictionary<string, object>();
                            sqlParameter.Clear();
                            sqlParameter.Add("EducationID", sGetID);
                            sqlParameter.Add("EducationSeq", i + 1);
                            sqlParameter.Add("PersonID", WinEduSub.PersonID);
                            sqlParameter.Add("Comments", WinEduSub.Comments);
                            sqlParameter.Add("CreateUserID", MainWindow.CurrentUser);

                            Procedure pro2 = new Procedure();
                            pro2.Name = "xp_HREdu_iEducationSub";
                            pro2.OutputUseYN = "N";
                            pro2.OutputName = "EducationID";
                            pro2.OutputLength = "10";

                            Prolist.Add(pro2);
                            ListParameter.Add(sqlParameter);
                        }

                        string[] Confirm = new string[2];
                        Confirm = DataStore.Instance.ExecuteAllProcedureOutputNew(Prolist, ListParameter);
                        if (Confirm[0] != "success")
                        {
                            MessageBox.Show("[저장실패]\r\n" + Confirm[1].ToString());
                            flag = false;
                        }
                        else
                        {
                            flag = true;
                        }


                    }

                    #endregion

                    #region 수정

                    else if (strFlag.Equals("U"))
                    {
                        string EducationFile_upgrade_yn = string.Empty;
                        if (txtEducationFile.Text == string.Empty)
                        {
                            // 첨부파일 깨끗하다면, ""으로 업그레이드.
                            // 경로가 깨끗하지 않다면, ViewReceiver와의 체크가 이루어질때까지 대기. 즉 업그레이드 파라미터 항목에서 제외.
                            EducationFile_upgrade_yn = "Y";
                        }
                        else
                        { EducationFile_upgrade_yn = "N"; }

                        sqlParameter.Add("EducationFile_upgrade_yn", EducationFile_upgrade_yn);
                        sqlParameter.Add("UserID", MainWindow.CurrentUser);

                        Procedure pro1 = new Procedure();
                        pro1.Name = "xp_HREdu_uEducation";
                        pro1.OutputUseYN = "N";
                        pro1.OutputName = "EducationID";
                        pro1.OutputLength = "10";

                        Prolist.Add(pro1);
                        ListParameter.Add(sqlParameter);

                        for (int i = 0; i < dgdSub.Items.Count; i++)
                        {
                            WinEduSub = dgdSub.Items[i] as Win_hr_Education_U_Sub_CodeView;
                            sqlParameter = new Dictionary<string, object>();
                            sqlParameter.Clear();
                            sqlParameter.Add("EducationID", WinEduMain.EducationID);
                            sqlParameter.Add("EducationSeq", i + 1);
                            sqlParameter.Add("PersonID", WinEduSub.PersonID);
                            sqlParameter.Add("Comments", WinEduSub.Comments);
                            sqlParameter.Add("CreateUserID", MainWindow.CurrentUser);

                            Procedure pro2 = new Procedure();
                            pro2.Name = "xp_HREdu_iEducationSub";
                            pro2.OutputUseYN = "N";
                            pro2.OutputName = "EducationID";
                            pro2.OutputLength = "10";

                            Prolist.Add(pro2);
                            ListParameter.Add(sqlParameter);
                        }

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
                            var ViewReceiver = dgdMain.SelectedItem as Win_hr_Education_U_CodeView;

                            bool AttachYesNo = false;
                            if (txtEducationFile.Text != string.Empty)       //첨부파일1 > DB 업로드 조건은 통과
                            {
                                if (txtEducationFile.Tag.ToString() != ViewReceiver.EducationFilePath)   // 기존 저장된 경로랑 새로 들어온 경로랑 같지 않을때만,
                                {
                                    if (FTP_Save_File(listFtpFile, ViewReceiver.EducationID))
                                    {
                                        txtEducationFile.Tag = "/ImageData/Education/" + ViewReceiver.EducationID;
                                        AttachYesNo = true;
                                    }
                                    else
                                    { MessageBox.Show("데이터 수정이 완료되었지만, 첨부문서 등록에 실패하였습니다."); }
                                }
                            }
                            if (AttachYesNo == true) { AttachFileUpdate(ViewReceiver.EducationID); }      //첨부문서 정보 DB 업데이트.
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

        /// <summary>
        /// 입력사항 체크
        /// </summary>
        /// <returns></returns>
        private bool CheckData()
        {
            bool flag = true;

            return flag;
        }

        // 1) 첨부문서가 있을경우, 2) FTP에 정상적으로 업로드가 완료된 경우.  >> DB에 정보 업데이트 
        private void AttachFileUpdate(string ID)
        {
            try
            {
                Dictionary<string, object> sqlParameter = new Dictionary<string, object>();

                sqlParameter.Clear();
                sqlParameter.Add("EducationID", ID);

                sqlParameter.Add("EducationFilePath", txtEducationFile.Tag.ToString());
                sqlParameter.Add("EducationFile", txtEducationFile.Text);

                string[] result = DataStore.Instance.ExecuteProcedure("xp_HREdu_uEducation_Ftp", sqlParameter, false);
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





        private void DataGridCell_KeyDown(object sender, KeyEventArgs e)
        {
            WinEduSub = dgdSub.CurrentItem as Win_hr_Education_U_Sub_CodeView;
            int startColCount = dgdSub.Columns.IndexOf(dgdtpeName);
            int colCount = dgdSub.Columns.IndexOf(dgdSub.CurrentCell.Column);
            int rowCount = dgdSub.Items.IndexOf(dgdSub.CurrentItem);

            if (e.Key == Key.Enter)
            {
                e.Handled = true;
                (sender as DataGridCell).IsEditing = false;

                if (dgdSub.Columns.Count - 1 == colCount && dgdSub.Items.Count - 1 > rowCount)
                {
                    dgdSub.CurrentCell = new DataGridCellInfo(dgdSub.Items[rowCount + 1], dgdSub.Columns[startColCount]);
                }
                else if (dgdSub.Columns.Count - 1 > colCount && dgdSub.Items.Count - 1 > rowCount)
                {
                    dgdSub.CurrentCell = new DataGridCellInfo(dgdSub.Items[rowCount], dgdSub.Columns[colCount + 1]);
                }
                else if (dgdSub.Columns.Count - 1 == colCount && dgdSub.Items.Count - 1 == rowCount)
                {
                    //btnSave.Focus();
                    btnSubAdd_Click(sender, e);
                }
                else if (dgdSub.Columns.Count - 1 > colCount && dgdSub.Items.Count - 1 == rowCount)
                {
                    dgdSub.CurrentCell = new DataGridCellInfo(dgdSub.Items[rowCount], dgdSub.Columns[colCount + 1]);
                }
                else
                {
                    MessageBox.Show("있으면 찾아보자...");
                }
            }
            else if (e.Key == Key.Delete)
            {
                SubRowDel();
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

        private void btnSubAdd_Click(object sender, RoutedEventArgs e)
        {
            SubRowAdd();
            int colCount = dgdSub.Columns.IndexOf(dgdtpeName);
            dgdSub.Focus();
            dgdSub.CurrentCell = new DataGridCellInfo(dgdSub.Items[dgdSub.Items.Count - 1], dgdSub.Columns[colCount]);
        }

        private void SubRowAdd()
        {
            int i = dgdSub.Items.Count;

            var WinhrEduSub = new Win_hr_Education_U_Sub_CodeView()
            {
                Comments = "",
                EducationID = "",
                EducationSeq = "",
                Name = "",
                Num = i + 1,
                PersonID = ""
            };
            dgdSub.Items.Add(WinhrEduSub);
        }

        private void btnSubDel_Click(object sender, RoutedEventArgs e)
        {
            SubRowDel();
        }

        private void SubRowDel()
        {
            if (dgdSub.Items.Count > 0)
            {
                if (dgdSub.CurrentItem != null)
                {
                    dgdSub.Items.Remove((dgdSub.CurrentItem as Win_hr_Education_U_Sub_CodeView));
                }
                else
                {
                    dgdSub.Items.Remove((dgdSub.Items[dgdSub.Items.Count - 1]) as Win_hr_Education_U_Sub_CodeView);
                }
                dgdSub.Refresh();
            }
        }

        //
        private void dgdtpetxtName_TextChanged(object sender, TextChangedEventArgs e)
        {
            if (lblMsg.Visibility == Visibility.Visible)
            {
                WinEduSub = dgdSub.CurrentItem as Win_hr_Education_U_Sub_CodeView;

                if (WinEduSub != null)
                {
                    TextBox tb1 = sender as TextBox;

                    if (tb1 != null && tb1.Tag != null)
                    {
                        WinEduSub.Name = tb1.Text;
                        WinEduSub.PersonID = tb1.Tag.ToString();
                    }

                    sender = tb1;
                }
            }
        }

        //
        private void dgdtpetxtName_KeyDown(object sender, KeyEventArgs e)
        {
            if (lblMsg.Visibility == Visibility.Visible)
            {
                WinEduSub = dgdSub.CurrentItem as Win_hr_Education_U_Sub_CodeView;

                if (e.Key == Key.Enter)
                {
                    TextBox tb1 = sender as TextBox;
                    MainWindow.pf.ReturnCode(tb1, (int)Defind_CodeFind.DCF_PERSON, "");

                    if (tb1.Tag != null)
                    {
                        WinEduSub.Name = tb1.Text;
                        WinEduSub.PersonID = tb1.Tag.ToString();
                    }

                    sender = tb1;
                }
            }
        }

        //
        private void dgdtpetxtComments_TextChanged(object sender, TextChangedEventArgs e)
        {
            if (lblMsg.Visibility == Visibility.Visible)
            {
                WinEduSub = dgdSub.CurrentItem as Win_hr_Education_U_Sub_CodeView;

                if (WinEduSub != null)
                {
                    TextBox tb1 = sender as TextBox;

                    if (tb1 != null)
                    {
                        WinEduSub.Comments = tb1.Text;
                    }

                    sender = tb1;
                }
            }
        }

        private void btnEducation_Click(object sender, RoutedEventArgs e)
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

                    txtEducationFile.Text = AttachFileName;
                    txtEducationFile.Tag = AttachFilePath.ToString();

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




        // FTP 보기버튼 클릭.
        private void btnSeeEducation_Click(object sender, RoutedEventArgs e)
        {
            MessageBoxResult msgresult = MessageBox.Show("파일을 다운로드 하시겠습니까?", "다운로드 확인", MessageBoxButton.YesNo);
            if (msgresult == MessageBoxResult.Yes)
            {
                var ViewReceiver = dgdMain.SelectedItem as Win_hr_Education_U_CodeView;

                if (ViewReceiver != null && !ViewReceiver.EducationFilePath.Equals(""))
                {
                    FTP_DownLoadFile(ViewReceiver.EducationFilePath, ViewReceiver.EducationID, ViewReceiver.EducationFile);
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






        private void DgdMain_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            if (e.ClickCount == 2)
            {
                btnUpdate_Click(btnUpdate, null);
            }
        }
        // 교육명에서 교육내용으로 이동.
        private void txtEducationName_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                txtEducationContext.Focus();
            }
        }
        // 교육자료 엔터쳐서 이미지 삽입 + 비고로 이동.
        private void txtEducationFile_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                btnEducation_Click(null, null);
                txtComments.Focus();
            }
        }
        // 비고에서 교육명으로 이동(반복)
        private void txtComments_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                txtEducationName.Focus();
            }
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
            DateTime[] SearchDate = Lib.Instance.BringLastMonthContinue(dtpEDate.SelectedDate.Value);

            dtpSDate.SelectedDate = SearchDate[0];
            dtpEDate.SelectedDate = SearchDate[1];
        }

        // 금월 버튼 클릭 이벤트
        private void btnThisMonth_Click(object sender, RoutedEventArgs e)
        {
            dtpSDate.SelectedDate = Lib.Instance.BringThisMonthDatetimeList()[0];
            dtpEDate.SelectedDate = Lib.Instance.BringThisMonthDatetimeList()[1];
        }
    }

    class Win_hr_Education_U_CodeView : BaseView
    {
        public int Num { get; set; }
        public string EducationID { get; set; }
        public string EducationName { get; set; }
        public string EducationStartDate { get; set; }
        public string EducationEndDate { get; set; }

        public string EducationContext { get; set; }
        public string EducationFilePath { get; set; }
        public string EducationFile { get; set; }
        public string Comments { get; set; }

        public string EducationStartDate_CV { get; set; }
        public string EducationEndDate_CV { get; set; }
        public string EduDate { get; set; }
        public int PersonCount { get; set; }
    }

    class Win_hr_Education_U_Sub_CodeView : BaseView
    {
        public int Num { get; set; }
        public string EducationID { get; set; }
        public string EducationSeq { get; set; }
        public string PersonID { get; set; }
        public string Name { get; set; }
        public string Comments { get; set; }
    }
}

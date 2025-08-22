using WizMes_WellMade.PopUP;
using System;
using System.Collections.Generic;
using System.Data;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using WPF.MDI;
using System.Collections.ObjectModel;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.StartPanel;
using System.Data.SqlClient;
using System.Drawing;
using System.IO;
using System.Windows.Media.Imaging;
using static System.Net.WebRequestMethods;
using System.Text;
using static System.Net.Mime.MediaTypeNames;
using WizMes_WellMade.PopUp;
using System.Windows.Media;

namespace WizMes_WellMade
{
    /// <summary>
    /// Win_dvl_MoldRegularInspect_U.xaml에 대한 상호 작용 논리
    /// </summary>
    public partial class Win_dvl_MoldRegularInspect_U : UserControl
    {
        #region 변수선언 및 로드

        Lib lib = new Lib();
        string strFlag = string.Empty;
        private int rowNum;

        // FTP 활용모음.
        private FTP_EX _ftp = null;

        List<string[]> listFtpFile = new List<string[]>();
        List<string[]> deleteListFtpFile = new List<string[]>(); // 삭제할 파일 리스트

        private string FTP_ADDRESS = string.Empty;
        private const string FTP_ID = "wizuser";
        private const string FTP_PASS = "wiz9999";
        private string folderPath = "/ImageData/MoldInspect";
        private const string LOCAL_DOWN_PATH = "C:\\Temp";

        public Win_dvl_MoldRegularInspect_U()
        {
            InitializeComponent();

#if DEBUG
         FTP_ADDRESS = "ftp://121.254.224.196";
#else
         FTP_ADDRESS = "ftp://" + LoadINI.FileSvr + ":"
                    + LoadINI.FTPPort + LoadINI.FtpImagePath + "/MoldInspect";
#endif
        }

        private void UserControl_Loaded(object sender, RoutedEventArgs e)
        {
            dtpSDate.SelectedDate = DateTime.Today;
            dtpEDate.SelectedDate = DateTime.Today;

            lib.UiLoading(sender);
            chkInspectDaySrh.IsChecked = true;
            SetComboBox();

            string FTP_FULL_PATH = FTP_ADDRESS + folderPath;
            _ftp = new FTP_EX(FTP_FULL_PATH, FTP_ID, FTP_PASS);
        }

        private void SetComboBox()
        {
            ObservableCollection<CodeView> ovcMcInsCycleGbnSrh = ComboBoxUtil.Instance.Gf_DB_CM_GetComCodeDataset(null, "MCCYCLEGBN", "Y", "", "");

            this.cboInspectCycle.ItemsSource = ovcMcInsCycleGbnSrh;
            this.cboInspectCycle.DisplayMemberPath = "code_name";
            this.cboInspectCycle.SelectedValuePath = "code_id";
        }

        #endregion

        #region 검색조건

        //점검기간 라벨
        private void lblInspectDaySrh_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            if (chkInspectDaySrh.IsChecked == true) { chkInspectDaySrh.IsChecked = false; }
            else { chkInspectDaySrh.IsChecked = true; }
        }

        //점검기간 체크박스
        private void chkInspectDaySrh_Checked(object sender, RoutedEventArgs e)
        {
            dtpSDate.IsEnabled = true;
            dtpEDate.IsEnabled = true;
        }

        //점검기간 체크박스
        private void chkInspectDaySrh_Unchecked(object sender, RoutedEventArgs e)
        {
            dtpSDate.IsEnabled = false;
            dtpEDate.IsEnabled = false;
        }

        //금월
        private void btnThisMonth_Click(object sender, RoutedEventArgs e)
        {
            dtpSDate.SelectedDate = lib.BringThisMonthDatetimeList()[0];
            dtpEDate.SelectedDate = lib.BringThisMonthDatetimeList()[1];
        }

        //금일
        private void btnToday_Click(object sender, RoutedEventArgs e)
        {
            dtpSDate.SelectedDate = DateTime.Today;
            dtpEDate.SelectedDate = DateTime.Today;
        }

        //금형 라벨
        private void lblMoldSrh_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            if (chkMoldSrh.IsChecked == true) { chkMoldSrh.IsChecked = false; }
            else { chkMoldSrh.IsChecked = true; }
        }

        //금형 체크박스
        private void chkMoldSrh_Checked(object sender, RoutedEventArgs e)
        {
            txtMoldSrh.IsEnabled = true;
            btnPfMoldSrh.IsEnabled = true;
            txtMoldSrh.Focus();
        }

        //금형 체크박스
        private void chkMoldSrh_Unchecked(object sender, RoutedEventArgs e)
        {
            txtMoldSrh.IsEnabled = false;
            btnPfMoldSrh.IsEnabled = false;
        }

        //금형 플러스파인더 이벤트(텍스트박스)
        private void txtMoldSrh_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                MainWindow.pf.ReturnCode(txtMoldSrh, 51, "");
            }
        }

        //금형 플러스파인더 이벤트(버튼)
        private void btnPfMoldSrh_Click(object sender, RoutedEventArgs e)
        {
            MainWindow.pf.ReturnCode(txtMoldSrh, 51, "");
        }
        #endregion

        #region 버튼

        //추가,수정 시 동작 모음
        private void ControlVisibleAndEnable_AU()
        {
            lib.UiButtonEnableChange_SCControl(this);
            dgdMoldInspect.IsEnabled = false;
            grbMold.IsEnabled = true;
        }

        //저장,취소 시 동작 모음
        private void ControlVisibleAndEnable_SC()
        {
            lib.UiButtonEnableChange_IUControl(this);
            dgdMoldInspect.IsEnabled = true;
            grbMold.IsEnabled = false;
        }

        //추가
        private void btnAdd_Click(object sender, RoutedEventArgs e)
        {
            if (dgdMoldInspect.Items.Count > 0 && dgdMoldInspect.SelectedItem != null)
            {
                rowNum = dgdMoldInspect.SelectedIndex; 
            }

            ControlVisibleAndEnable_AU();            
            strFlag = "I";
            tbkMsg.Text = "자료 입력(추가) 중";

            this.DataContext = null;
            dtpMoldInspectDate.SelectedDate = DateTime.Today;

            clearGrid();
        }

        //수정
        private void btnUpdate_Click(object sender, RoutedEventArgs e)
        {
            if (dgdMoldInspect.SelectedItem == null)
            {
                MessageBox.Show("수정할 자료를 선택하고 눌러주십시오.");
            }
            else
            {
                rowNum = dgdMoldInspect.SelectedIndex;
                ControlVisibleAndEnable_AU();
                tbkMsg.Text = "자료 입력(수정) 중";
                strFlag = "U";
                txtMoldID.Focus();
            }
        }

        //삭제
        private void btnDelete_Click(object sender, RoutedEventArgs e)
        {
            var winMoldInspect = dgdMoldInspect.SelectedItem as Win_dvl_MoldRegularInspect_U_CodeView;

            if (winMoldInspect == null)
            {
                MessageBox.Show("삭제할 데이터가 지정되지 않았습니다. 삭제데이터를 지정하고 눌러주세요");
                return;
            }
            else
            {
                if (dgdMoldInspect.SelectedIndex == dgdMoldInspect.Items.Count - 1)
                {
                    rowNum = dgdMoldInspect.SelectedIndex - 1;
                }
                else
                {
                    rowNum = dgdMoldInspect.SelectedIndex;
                }

                if (MessageBox.Show("선택하신 항목을 삭제하시겠습니까?", "삭제 전 확인", MessageBoxButton.YesNo) == MessageBoxResult.Yes)
                {
                    if (DeleteData(winMoldInspect.MoldInspectID))
                    {
                        re_Search(rowNum);
                    }
                }
            }
        }

        //닫기
        private void btnClose_Click(object sender, RoutedEventArgs e)
        {
            string stDate = DateTime.Now.ToString("yyyyMMdd");
            string stTime = DateTime.Now.ToString("HHmm");
            DataStore.Instance.InsertLogByFormS(this.GetType().Name, stDate, stTime, "E");
            Lib.Instance.ChildMenuClose(this.ToString());
        }

        //조회
        private void btnSearch_Click(object sender, RoutedEventArgs e)
        {
            rowNum = 0;
            re_Search(rowNum);
        }

        //저장
        private void btnSave_Click(object sender, RoutedEventArgs e)
        {
            if (strFlag.Equals("I"))
            {
                if (SaveData("", strFlag))
                {
                    ControlVisibleAndEnable_SC();
                    rowNum = 0;
                    re_Search(rowNum);
                }
            }
            else
            {
                if (SaveData(txtMoldInspectID.Text, strFlag))
                {
                    ControlVisibleAndEnable_SC();
                    re_Search(rowNum);
                }
            }
        }

        //취소
        private void btnCancel_Click(object sender, RoutedEventArgs e)
        {
            InputClear();
            ControlVisibleAndEnable_SC();
            re_Search(rowNum);
        }

        //입력 데이터 클리어
        private void InputClear()
        {
            this.DataContext = null;
            foreach (Control child in this.grdInput.Children)
            {
                if (child.GetType() == typeof(TextBox))
                    ((TextBox)child).Clear();
            }
            clearGrid();
        }

        //엑셀
        private void btnExcel_Click(object sender, RoutedEventArgs e)
        {
            DataTable dt = null;
            string Name = string.Empty;

            string[] lst = new string[6];
            lst[0] = "금형 일상점검 메인";
            lst[1] = "금형 일상점검_범례";
            lst[2] = "금형 일상점검_수치";
            lst[3] = dgdMoldInspect.Name;
            lst[4] = dgdMold_InspectSub1.Name;
            lst[5] = dgdMold_InspectSub2.Name;

            ExportExcelxaml ExpExc = new ExportExcelxaml(lst);

            ExpExc.ShowDialog();

            if (ExpExc.DialogResult.HasValue)
            {
                if (ExpExc.choice.Equals(dgdMoldInspect.Name))
                {
                    if (ExpExc.Check.Equals("Y"))
                        dt = lib.DataGridToDTinHidden(dgdMoldInspect);
                    else
                        dt = lib.DataGirdToDataTable(dgdMoldInspect);

                    Name = dgdMoldInspect.Name;

                    if (lib.GenerateExcel(dt, Name))
                        lib.excel.Visible = true;
                    else
                        return;
                }
                else if (ExpExc.choice.Equals(dgdMold_InspectSub1.Name))
                {
                    if (ExpExc.Check.Equals("Y"))
                        dt = lib.DataGridToDTinHidden(dgdMold_InspectSub1);
                    else
                        dt = lib.DataGirdToDataTable(dgdMold_InspectSub1);

                    Name = dgdMold_InspectSub1.Name;

                    if (lib.GenerateExcel(dt, Name))
                        lib.excel.Visible = true;
                    else
                        return;
                }
                else if (ExpExc.choice.Equals(dgdMold_InspectSub2.Name))
                {
                    if (ExpExc.Check.Equals("Y"))
                        dt = lib.DataGridToDTinHidden(dgdMold_InspectSub2);
                    else
                        dt = lib.DataGirdToDataTable(dgdMold_InspectSub2);

                    Name = dgdMold_InspectSub2.Name;

                    if (lib.GenerateExcel(dt, Name))
                        lib.excel.Visible = true;
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


        #endregion

        #region CRUD
        //수정,추가,삭제 후 조회 등
        private void re_Search(int index)
        {
            if (dgdMoldInspect.Items.Count > 0)
            {
                dgdMoldInspect.Items.Clear();
            }

            FillGrid();

        }

        private void FillGrid() 
        {
            try
            {
                Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                sqlParameter.Add("nChkDate", chkInspectDaySrh.IsChecked == true ? 1 : 0);
                sqlParameter.Add("SDate", chkInspectDaySrh.IsChecked == true ? dtpSDate.SelectedDate.Value.ToString("yyyyMMdd") : "");
                sqlParameter.Add("EDate", chkInspectDaySrh.IsChecked == true ? dtpEDate.SelectedDate.Value.ToString("yyyyMMdd") : "");
                sqlParameter.Add("nChkMold", chkMoldSrh.IsChecked == true ? 1: 0);
                sqlParameter.Add("MoldID", chkMoldSrh.IsChecked == true  && txtMoldSrh.Tag != null ? txtMoldSrh.Tag.ToString()  : "");

                DataSet ds = DataStore.Instance.ProcedureToDataSet("xp_Mold_sInspect", sqlParameter, false);

                if (ds != null && ds.Tables.Count > 0)
                {
                    DataTable dt = ds.Tables[0];

                    if (dt.Rows.Count > 0)
                    {
                        DataRowCollection drc = dt.Rows;
                        int i = 0;
                        foreach (DataRow dr in drc)
                        {
                            i++;
                            var WinMoldInspect = new Win_dvl_MoldRegularInspect_U_CodeView()
                            {
                                Num = i,
                                MoldInspectID = dr["MoldInspectID"].ToString(),
                                MoldInspectBasisID = dr["MoldInspectBasisID"].ToString(),
                                MoldInspectPersonID = dr["MoldInspectPersonID"].ToString(),
                                MoldInspectPerson = dr["MoldInspectPerson"].ToString(),
                                MoldID = dr["MoldID"].ToString(),
                                MoldInspectDate = DatePickerFormat(dr["MoldInspectDate"].ToString()),
                                Comments = dr["Comments"].ToString(),
                                InspectCycle = dr["InspectCycle"].ToString(),
                                FileName = dr["FileName"].ToString(),
                                FilePath = dr["FilePath"].ToString(),
                                Article = dr["Article"].ToString(),
                            };

                            dgdMoldInspect.Items.Add(WinMoldInspect);
                        }
                    } else
                    {
                        MessageBox.Show("조회된 데이터가 없습니다");
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

        private void getBasisSub(string strMoldID, string cycle)
        {
            clearGrid();

            try
            {
                Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                sqlParameter.Add("MoldID", strMoldID);
                sqlParameter.Add("Cycle", cycle);

                DataSet ds = DataStore.Instance.ProcedureToDataSet("xp_Mold_sBasisSub", sqlParameter, false);

                if (ds != null && ds.Tables.Count > 0)
                {
                    DataTable dt = ds.Tables[0];

                    if (dt.Rows.Count > 0)
                    {
                        DataRowCollection drc = dt.Rows;

                        foreach (DataRow dr in drc)
                        {
                            var WinMoldSub = new Win_dvl_MoldRegularInspect_U_Sub_CodeView()
                            {
                                MoldInspectBasisID = dr["MoldInspectBasisID"].ToString(),
                                MoldID = dr["MoldID"].ToString(),
                                MoldInspectSeq = Convert.ToInt32(dr["MoldInspectSeq"]),
                                MoldInspectItemName = dr["MoldInspectItemName"].ToString(),
                                MoldInspectContent = dr["MoldInspectContent"].ToString(),
                                MoldInspectCheckGbn = dr["MoldInspectCheckGbn"].ToString(),
                                MoldInspectCheckName = dr["MoldInspectCheckName"].ToString(),
                                MoldInspectCycleGbn = dr["MoldInspectCycleGbn"].ToString(),
                                MoldInspectCycleName = dr["MoldInspectCycleName"].ToString(),
                                MoldInspectCycleDate = dr["MoldInspectCycleDate"].ToString(),
                                MoldInspectRecordGbn = dr["MoldInspectRecordGbn"].ToString(),
                                MoldInspectRecordName = dr["MoldInspectRecordName"].ToString(),
                                Comments = ""
                            };

                            if ("01".Equals(WinMoldSub.MoldInspectRecordGbn))
                            {
                                WinMoldSub.Num = dgdMold_InspectSub1.Items.Count + 1;
                                dgdMold_InspectSub1.Items.Add(WinMoldSub);
                            }
                            else if ("02".Equals(WinMoldSub.MoldInspectRecordGbn))
                            {
                                WinMoldSub.Num = dgdMold_InspectSub2.Items.Count + 1;
                                dgdMold_InspectSub2.Items.Add(WinMoldSub);
                            }

                            txtMoldBasisID.Text = WinMoldSub.MoldInspectBasisID;

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

        private void FillGridSub(string strMoldInspectID)
        {
            clearGrid();

            try
            {
                Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                sqlParameter.Add("MoldInspectID", strMoldInspectID);
             
                DataSet ds = DataStore.Instance.ProcedureToDataSet("xp_Mold_sInspectSub", sqlParameter, false);

                if (ds != null && ds.Tables.Count > 0)
                {
                    DataTable dt = ds.Tables[0];

                    if (dt.Rows.Count > 0)
                    {
                        DataRowCollection drc = dt.Rows;

                        foreach (DataRow dr in drc)
                        {
                            var WinMoldSub = new Win_dvl_MoldRegularInspect_U_Sub_CodeView()
                            {
                                MoldInspectID = dr["MoldInspectID"].ToString(),
                                MoldInspectBasisID = dr["MoldInspectBasisID"].ToString(),
                                MoldID = dr["MoldID"].ToString(),
                                MoldInspectSeq = Convert.ToInt32(dr["MoldInspectSeq"]),
                                MoldInspectItemName = dr["MoldInspectItemName"].ToString(),
                                MoldInspectContent = dr["MoldInspectContent"].ToString(),
                                MoldInspectCheckGbn = dr["MoldInspectCheckGbn"].ToString(),
                                MoldInspectCheckName = dr["MoldInspectCheckName"].ToString(),
                                MoldInspectCycleGbn = dr["MoldInspectCycleGbn"].ToString(),
                                MoldInspectCycleName = dr["MoldInspectCycleName"].ToString(),
                                MoldInspectCycleDate = dr["MoldInspectCycleDate"].ToString(),
                                MoldInspectRecordGbn = dr["MoldInspectRecordGbn"].ToString(),
                                MoldInspectRecordName = dr["MoldInspectRecordName"].ToString(),
                                MldInspectLegend = dr["MldInspectLegend"].ToString(),
                                MldValue = Convert.ToDouble(dr["MldValue"]),
                                Comments = dr["Comments"].ToString()
                            };

                            if ("01".Equals(WinMoldSub.MoldInspectRecordGbn)) 
                            {
                                WinMoldSub.Num = dgdMold_InspectSub1.Items.Count + 1;
                                dgdMold_InspectSub1.Items.Add(WinMoldSub);
                            }
                            else if ("02".Equals(WinMoldSub.MoldInspectRecordGbn))
                            {
                                WinMoldSub.Num = dgdMold_InspectSub2.Items.Count + 1;
                                dgdMold_InspectSub2.Items.Add(WinMoldSub);
                            }
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

        private void dgdMoldInspect_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            clearGrid();

            var WinMold = dgdMoldInspect.SelectedItem as Win_dvl_MoldRegularInspect_U_CodeView;

            if (WinMold != null)
            {
                this.DataContext = WinMold;
                FillGridSub(WinMold.MoldInspectID);
            }
        }

        //삭제
        private bool DeleteData(string strMoldInspectID)
        {
            bool flag = true;

            try
            {
                Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                sqlParameter.Add("MoldInspectID", strMoldInspectID);

                string[] result = DataStore.Instance.ExecuteProcedure("xp_dvlMoldIns_dRegularInspect", sqlParameter, true);

                if (!result[0].Equals("success"))
                {
                    //MessageBox.Show("실패 ㅠㅠ");
                }
                else
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

        //추가, 수정
        private bool SaveData(string strMoldInspectID, string strFlag)
        {
            bool flag = true;
            string inspectID = string.Empty;
            List<Procedure> Prolist = new List<Procedure>();
            List<Dictionary<string, object>> ListParameter = new List<Dictionary<string, object>>();

            if (CheckData())
            {
                try
                {
                    Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                    sqlParameter.Add("MoldInspectID", strMoldInspectID);
                    sqlParameter.Add("MoldInspectBasisID", txtMoldBasisID.Text);
                    sqlParameter.Add("MoldInspectDate", dtpMoldInspectDate.SelectedDate.Value.ToString("yyyyMMdd"));
                    sqlParameter.Add("MoldInspectPersonID", txtPerson.Tag.ToString());
                    sqlParameter.Add("Comments", txtComments.Text);
                    sqlParameter.Add("CreateUserID", MainWindow.CurrentUser);

                    if (strFlag.Equals("I"))
                    {
                        Procedure pro1 = new Procedure();
                        pro1.Name = "xp_dvlMoldIns_iRegularInspect";
                        pro1.OutputUseYN = "Y";
                        pro1.OutputName = "MoldInspectID";
                        pro1.OutputLength = "10";

                        Prolist.Add(pro1);
                        ListParameter.Add(sqlParameter);


                        List<KeyValue> list_Result = new List<KeyValue>();
                        list_Result = DataStore.Instance.ExecuteAllProcedureOutputGetCS(Prolist, ListParameter);
                        

                        if (list_Result[0].key.ToLower() == "success")
                        {
                            list_Result.RemoveAt(0);

                            for (int i = 0; i < list_Result.Count; i++)
                            {
                                KeyValue kv = list_Result[i];
                                if (kv.key == "MoldInspectID")
                                {
                                    inspectID = kv.value;
                                    saveSub(inspectID);
                                }
                            }
                        }
                        else
                        {
                            MessageBox.Show("[저장실패]\r\n" + list_Result[0].value.ToString());
                            flag = false;
                        }
                    }
                    else
                    {
                      
                        Procedure pro1 = new Procedure();
                        pro1.Name = "xp_dvlMoldIns_uRegularInspect";
                        pro1.OutputUseYN = "N";
                        pro1.OutputName = "MoldInspectID";
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
                        } else
                        {
                            inspectID = txtMoldInspectID.Text;
                            saveSub(inspectID);
                        }
                    }

                    if (listFtpFile.Count > 0)
                    {
                        if(SaveImage(listFtpFile, inspectID)) UpdateDBFtp(inspectID); ;
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
            else { flag = false; }

            return flag;
        }

        private void saveSub(string inspectID)
        {
            List<Procedure> Prolist = new List<Procedure>();
            List<Dictionary<string, object>> ListParameter = new List<Dictionary<string, object>>();
            Dictionary<string, object> sqlParameter = new Dictionary<string, object>();

            for (int i = 0; i < dgdMold_InspectSub1.Items.Count; i++)
            {
                var sub = dgdMold_InspectSub1.Items[i] as Win_dvl_MoldRegularInspect_U_Sub_CodeView;

                sqlParameter = new Dictionary<string, object>();
                sqlParameter.Clear();
                sqlParameter.Add("MoldInspectID", inspectID);
                sqlParameter.Add("MoldInspectSubSeq", sub.MoldInspectSeq);
                sqlParameter.Add("MoldInsBasisID", sub.MoldInspectBasisID);
                sqlParameter.Add("MldValue", sub.MldValue);
                sqlParameter.Add("MldInspectLegend", sub.MldInspectLegend);
                sqlParameter.Add("Comments", sub.Comments);
                sqlParameter.Add("CreateUserID", MainWindow.CurrentUser);

                Procedure pro2 = new Procedure();
                pro2.Name = "xp_dvlMoldIns_iRegularInspectSub";
                pro2.OutputUseYN = "N";
                pro2.OutputName = "MoldInspectID";
                pro2.OutputLength = "10";

                Prolist.Add(pro2);
                ListParameter.Add(sqlParameter);
            }

            for (int i = 0; i < dgdMold_InspectSub2.Items.Count; i++)
            {
                var sub = dgdMold_InspectSub2.Items[i] as Win_dvl_MoldRegularInspect_U_Sub_CodeView;

                sqlParameter = new Dictionary<string, object>();
                sqlParameter.Clear();
                sqlParameter.Add("MoldInspectID", inspectID);
                sqlParameter.Add("MoldInspectSubSeq", sub.MoldInspectSeq);
                sqlParameter.Add("MoldInsBasisID", sub.MoldInspectBasisID);
                sqlParameter.Add("MldValue", sub.MldValue);
                sqlParameter.Add("MldInspectLegend", sub.MldInspectLegend);
                sqlParameter.Add("Comments", sub.Comments);
                sqlParameter.Add("CreateUserID", MainWindow.CurrentUser);

                Procedure pro3 = new Procedure();
                pro3.Name = "xp_dvlMoldIns_iRegularInspectSub";
                pro3.OutputUseYN = "N";
                pro3.OutputName = "MoldInspectID";
                pro3.OutputLength = "10";

                Prolist.Add(pro3);
                ListParameter.Add(sqlParameter);
            }

            string[] Confirm = new string[2];
            Confirm = DataStore.Instance.ExecuteAllProcedureOutputNew(Prolist, ListParameter);

            if (Confirm[0] != "success")
            {
                MessageBox.Show("[저장실패]\r\n" + Confirm[1].ToString());
            }
        }

        private bool UpdateDBFtp(string MoldInspectID)
        {
            bool flag = false;
            string path = folderPath + "/" +  MoldInspectID + "/";

            try
            {
                Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                sqlParameter.Clear();
                sqlParameter.Add("MoldInspectID", MoldInspectID);
                sqlParameter.Add("FileName", txtResultFile.Text ?? "");
                sqlParameter.Add("FilePath", !string.IsNullOrWhiteSpace(txtResultFile.Text) ? path : "");
                sqlParameter.Add("LastUpdateUserID", MainWindow.CurrentUser);

                string[] result = DataStore.Instance.ExecuteProcedure("xp_dvlMold_uMoldInspect_Ftp", sqlParameter, true);

                if (result[0].Equals("success"))
                {
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

        //추가, 수정 시 필수 입력 체크
        private bool CheckData()
        {
            bool flag = true;

            if (txtMoldID.Tag == null || txtMoldID.Tag.ToString().Equals(""))
            {
                MessageBox.Show("금형 선택이 잘못되었습니다. enter키 또는 품명 옆의 버튼을 이용하여 다시 입력해주세요");
                flag = false;
                return flag;
            }

            if (dtpMoldInspectDate.SelectedDate == null)
            {
                MessageBox.Show("점검일자가 선택되지 않았습니다. 점검일자를 선택해주세요");
                flag = false;
                return flag;
            }

            if (txtPerson.Tag == null || txtPerson.Tag.ToString().Equals(""))
            {
                MessageBox.Show("점검자 선택이 잘못되었습니다. enter키 또는 품명 옆의 버튼을 이용하여 다시 입력해주세요");
                flag = false;
                return flag;
            }

            return flag;
        }

        #endregion

        #region input event 

        //금형번호(textbox)
        private void txtMoldID_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                MainWindow.pf.ReturnCode(txtMoldID, 51, txtMoldID.Text);

                if (txtMoldID.Tag != null)
                {
                    getMoldInfo(txtMoldID.Tag.ToString());
                }

                dtpMoldInspectDate.Focus();
            }
        }

        //금형번호(button)
        private void btnPfMoldID_Click(object sender, RoutedEventArgs e)
        {
            MainWindow.pf.ReturnCode(txtMoldID, 51, txtMoldID.Text);

            if (txtMoldID.Tag != null)
            {
                getMoldInfo(txtMoldID.Tag.ToString());
            }


            dtpMoldInspectDate.Focus();
        }

        //금형번호 선택시, 선택된 금형의 정보를 가져온다.
        private void getMoldInfo(string strMoldID)
        {
            try
            {

                Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                sqlParameter.Add("MoldID", strMoldID);
                DataSet ds = DataStore.Instance.ProcedureToDataSet("xp_Mold_getMoldInfo", sqlParameter, false);

                if (ds != null && ds.Tables.Count > 0)
                {
                    DataTable dt = ds.Tables[0];

                    if (dt.Rows.Count > 0)
                    {
                        DataRowCollection drc = dt.Rows;

                        string date = string.Empty;
                        string personID = string.Empty;
                        string person = string.Empty;

                        foreach (DataRow dr in drc)
                        {

                            txtMoldBasisID.Text = dr["MoldInspectBasisID"].ToString();
                            txtMoldID.Text = dr["MoldID"].ToString();
                            txtMoldInspectID.Text = dr["MoldInspectID"].ToString();
                            txtPerson.Tag = dr["MoldInspectPersonID"].ToString();
                            txtPerson.Text = dr["MoldInspectPerson"].ToString();
                            cboInspectCycle.SelectedValue = dr["InspectCycle"].ToString();
                            txtResultFile.Text = dr["FileName"].ToString();
                            txtResultFile.Tag = dr["FilePath"].ToString();
                            txtComments.Text = dr["Comments"].ToString();
                            txtArticle.Text = dr["Article"].ToString();

                            dtpMoldInspectDate.SelectedDate = DateTime.Now;

                            personID = dr["MoldInspectPersonID"].ToString();
                            txtPerson.Tag = !string.IsNullOrWhiteSpace(personID) ? personID : MainWindow.CurrentUser;

                            person = dr["MoldInspectPerson"].ToString();
                            txtPerson.Text = !string.IsNullOrWhiteSpace(person) ? person : MainWindow.CurrentUser;
                        }
                    } else
                    {
                        MessageBox.Show("해당 금형번호로 등록된 점검기준이 없습니다");
                        InputClear();
                        return;
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

        //점검일자
        private void dtpMoldInspectDate_PreviewKeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                dtpMoldInspectDate.IsDropDownOpen = true;
            }
        }

        //점검일자
        private void dtpMoldInspectDate_CalendarClosed(object sender, RoutedEventArgs e)
        {
            txtPerson.Focus();
        }

        //점검자
        private void txtPerson_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                MainWindow.pf.ReturnCode(txtPerson, 2, "");
                txtComments.Focus();
            }
        }

        //점검자
        private void btnPfPerson_Click(object sender, RoutedEventArgs e)
        {
            MainWindow.pf.ReturnCode(txtPerson, 2, "");
            txtComments.Focus();
        }
        private void cboInspectCycle_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (cboInspectCycle.SelectedValue == null) return;
            if (string.IsNullOrWhiteSpace(txtMoldID.Tag.ToString())) return;

            getBasisSub(txtMoldID.Tag.ToString(), cboInspectCycle.SelectedValue.ToString());

        }
        #endregion

        #region 서브그리드 이벤트
        private void TextBoxOnlyNumber_PreviewTextInput(object sender, TextCompositionEventArgs e)
        {
            lib.CheckIsNumeric((TextBox)sender, e);
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
                DataGrid grid = FindParent<DataGrid>(cell);
                int currCol = grid.Columns.IndexOf(grid.CurrentCell.Column);

                if (currCol >5 && currCol < 8) cell.IsEditing = true;

            }

           
        }
        public static T FindParent<T>(DependencyObject child) where T : DependencyObject
        {
            DependencyObject parent = VisualTreeHelper.GetParent(child);

            while (parent != null && !(parent is T))
            {
                parent = VisualTreeHelper.GetParent(parent);
            }

            return parent as T;
        }


        #endregion


        private void clearGrid()
        {
            if (dgdMold_InspectSub1.Items.Count > 0)
            {
                dgdMold_InspectSub1.Items.Clear();
            }
            if(dgdMold_InspectSub2.Items.Count > 0)
            {
                dgdMold_InspectSub2.Items.Clear();
            }
        }

        // 데이터피커 포맷으로 변경
        private string DatePickerFormat(string str)
        {
            str = str.Trim().Replace("-", "").Replace(".", "");

            if (!str.Equals(""))
            {
                if (str.Length == 8)
                {
                    str = str.Substring(0, 4) + "-" + str.Substring(4, 2) + "-" + str.Substring(6, 2);
                }
            }

            return str;
        }

        private bool SaveImage(List<string[]> listStrArrayFileInfo, string inspectID)
        {
            _ftp = new FTP_EX(FTP_ADDRESS, FTP_ID, FTP_PASS);

            List<string[]> UpdateFilesInfo = new List<string[]>();
            string[] fileListSimple;
            string[] fileListDetail = null;
            fileListSimple = _ftp.directoryListSimple("", Encoding.Default);

            // 기존 폴더 확인작업.
            bool MakeFolder = false;
            MakeFolder = FolderInfoAndFlag(fileListSimple, inspectID);

            if (MakeFolder == false)        // 같은 아이를 찾지 못한경우,
            {
                //MIL 폴더에 InspectionID로 저장

                if (_ftp.createDirectory(inspectID) == false)
                {
                    MessageBox.Show("업로드를 위한 폴더를 생성할 수 없습니다.");
                    return false;
                }
            }
            else
            {
                fileListDetail = _ftp.directoryListSimple(inspectID, Encoding.Default);
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
                    listStrArrayFileInfo[i][0] = inspectID + "/" + listStrArrayFileInfo[i][0];
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

        private void txtResultFile_Click(object sender, MouseButtonEventArgs e)
        {
            TextBox txtbox = (TextBox)sender;
            FTP_Upload_TextBox(txtbox);
        }

        private void txtResultFile_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                TextBox txtbox = (TextBox)sender;
                FTP_Upload_TextBox(txtbox);
            }
        }

        private void FTP_Upload_TextBox(TextBox textBox)
        {
           
                Microsoft.Win32.OpenFileDialog OFdlg = new Microsoft.Win32.OpenFileDialog();
                //OFdlg.Filter =
                //    "Image files (*.jpg, *.jpeg, *.jpe, *.jfif, *.png, *.pcx, *.pdf) | *.jpg; *.jpeg; *.jpe; *.jfif; *.png; *.pcx; *.pdf | All Files|*.*";

                OFdlg.Filter = MainWindow.OFdlg_Filter;

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
                        return;
                    }
                    else
                    {
                        textBox.Text = ImageFileName;
                        textBox.Tag = ImageFilePath;

                        try
                        {
                            Bitmap image = new Bitmap(ImageFilePath + ImageFileName);
                            //imgSetting.Source = BitmapToImageSource(image);
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

        private BitmapImage SetImage(string path)
        {
            BitmapImage bit = _ftp.DrawingImageByByte(path);
            //image.Source = bit;
            return bit;
        }

        private void btnImage_Click(object sender, RoutedEventArgs e)
        {
            if (!string.IsNullOrWhiteSpace(txtResultFile.Text))
            {
                string fullPath = FTP_ADDRESS  + txtResultFile.Tag.ToString() +  txtResultFile.Text;
                BitmapImage img = SetImage(fullPath);
                LargeImagePopUp largeImagePopUp = new LargeImagePopUp(img);
            }
            
        }
    }
}

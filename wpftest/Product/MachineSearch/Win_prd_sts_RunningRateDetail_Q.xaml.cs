using LiveCharts;
using LiveCharts.Wpf;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Data;
using System.Windows;
using System.Windows.Controls;
using WizMes_WellMade.PopUp;
using WizMes_WellMade.PopUP;
using System.Windows.Input;
using System.Threading;

namespace WizMes_WellMade
{
    /**************************************************************************************************
    '** System 명 : WizMES
    '** Author    : Wizard
    '** 작성자    : 김수정
    '** 내용      : 설비가동률 상세조회
    '** 생성일자  : 2022.09
    '** 변경일자  : 
    '**------------------------------------------------------------------------------------------------
    ''*************************************************************************************************
    ' 변경일자  , 변경자, 요청자    , 요구사항ID  , 요청 및 작업내용
    '**************************************************************************************************
    '**************************************************************************************************/

    /// <summary>
    /// Win_prd_sts_RunningRateDetail_Q.xaml에 대한 상호 작용 논리
    /// </summary>
    public partial class Win_prd_sts_RunningRateDetail_Q : UserControl
    {
        string stDate = string.Empty;
        string stTime = string.Empty;
        string ProcessID = string.Empty;

        Lib lib = new Lib();
        public DataGrid FilterGrid { get; set; }
        public DataTable FilterTable { get; set; }

        ObservableCollection<Work_CodeView> WorkCollection = new ObservableCollection<Work_CodeView>();
        ObservableCollection<MC_CodeView> MCCollection = new ObservableCollection<MC_CodeView>();
        ObservableCollection<NoWork_CodeView> NoWorkCollection = new ObservableCollection<NoWork_CodeView>();

        public Win_prd_sts_RunningRateDetail_Q()
        {
            InitializeComponent();
        }

        private void UserControl_Loaded(object sender, RoutedEventArgs e)
        {
            SetComboBox();

            stDate = DateTime.Now.ToString("yyyyMMdd");
            stTime = DateTime.Now.ToString("HHmm");

            DataStore.Instance.InsertLogByFormS(this.GetType().Name, stDate, stTime, "S");

            string ProcessID = string.Empty;
            string MachineID = string.Empty;
            string MCSDate = string.Empty;
            string MCEDate = String.Empty;


            if (MainWindow.MCtemp != null
               && MainWindow.MCtemp.Count > 0)
            {
                ProcessID = MainWindow.MCtemp[0]; //processID
                MachineID = MainWindow.MCtemp[1]; //MachineID
                MCSDate = MainWindow.MCtemp[2]; //SDate
                MCEDate = MainWindow.MCtemp[3]; //EDate

                dtpSDate.SelectedDate = DateTime.Parse(DatePickerFormat(MCSDate));
                dtpEDate.SelectedDate = DateTime.Parse(DatePickerFormat(MCEDate));

            }
            else
            {
                dtpSDate.SelectedDate = DateTime.Today;
                dtpEDate.SelectedDate = DateTime.Today;
            }

            cboProcess.SelectedValue = ProcessID;
            cboMachine.SelectedValue = MachineID;

            Lib.Instance.UiLoading(sender);

           

        }

        #region Header - 검색조건

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

        //호기
        private void lblMachineSrh_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            if (chkMachineSrh.IsChecked == true) { chkMachineSrh.IsChecked = false; }
            else { chkMachineSrh.IsChecked = true; }
        }

        //호기
        private void chkMachineSrh_Checked(object sender, RoutedEventArgs e)
        {
            cboMachine.IsEnabled = true;
        }

        //호기
        private void chkMachineSrh_Unchecked(object sender, RoutedEventArgs e)
        {
            cboMachine.IsEnabled = false;
        }

        //공정
        private void lblProcessSrh_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            if (chkProcessSrh.IsChecked == true) { chkProcessSrh.IsChecked = false; }
            else { chkProcessSrh.IsChecked = true; }
        }

        //공정
        private void chkProcessSrh_Checked(object sender, RoutedEventArgs e)
        {
            cboProcess.IsEnabled = true;
        }

        //공정
        private void chkProcessSrh_Unchecked(object sender, RoutedEventArgs e)
        {
            cboProcess.IsEnabled = false;
        }

        #endregion

        #region Header - 오른쪽 상단 버튼

        //검색
        private void btnSearch_Click(object sender, RoutedEventArgs e)
        {
            //검색버튼 비활성화
            btnSearch.IsEnabled = false;

            Dispatcher.BeginInvoke(new Action(() =>

            {
                Thread.Sleep(2000);

                //로직
                using (Loading lw = new Loading(beSearch))
                {
                    lw.ShowDialog();
                }

            }), System.Windows.Threading.DispatcherPriority.Background);



            Dispatcher.BeginInvoke(new Action(() =>

            {
                btnSearch.IsEnabled = true;

            }), System.Windows.Threading.DispatcherPriority.Background);
            
        }

        private void beSearch()
        {
            if (dtpSDate.SelectedDate == null
                || dtpSDate.SelectedDate.ToString() == ""
                || dtpEDate.SelectedDate == null
                || dtpEDate.SelectedDate.ToString() == null)
            {
                MessageBox.Show("날짜를 정확히 입력해주세요.");
                //검색 다 되면 활성화
                btnSearch.IsEnabled = true;
                return;
            }

            FillGrid_Work();
            FillGrid_Machine();
            FillGrid_NoWork();

            if (dgdWork.Items.Count > 0 && dgdMachine.Items.Count > 0 && dgdNoWork.Items.Count > 0)
            {
                dgdWork.SelectedIndex = 0;
                dgdMachine.SelectedIndex = 0;
                dgdNoWork.SelectedIndex = 0;

            } else
            {
                MessageBox.Show("조회된 내용이 없습니다.");
            }
           

           
        }

        //닫기
        private void btnClose_Click(object sender, RoutedEventArgs e)
        {
            DataStore.Instance.InsertLogByFormS(this.GetType().Name, stDate, stTime, "E");
            Lib.Instance.ChildMenuClose(this.ToString());
        }

        //엑셀
        private void btnExcel_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                DataTable dt = null;
                string Name = string.Empty;

                string[] lst = new string[6];
                lst[0] = "작업 조회";
                lst[1] = "설비 조회";
                lst[2] = "비가동 조회";
                lst[3] = dgdWork.Name;
                lst[4] = dgdMachine.Name;
                lst[5] = dgdNoWork.Name;

                ExportExcelxaml ExpExc = new ExportExcelxaml(lst);

                ExpExc.ShowDialog();

                if (ExpExc.DialogResult.HasValue)
                {
                    if (ExpExc.choice.Equals(dgdWork.Name))
                    {
                        DataStore.Instance.InsertLogByForm(this.GetType().Name, "E");
                        if (ExpExc.Check.Equals("Y"))
                            dt = Lib.Instance.DataGridToDTinHidden(dgdWork);
                        else
                            dt = Lib.Instance.DataGirdToDataTable(dgdWork);

                        Name = dgdWork.Name;

                        if (Lib.Instance.GenerateExcel(dt, Name))
                            Lib.Instance.excel.Visible = true;
                        else
                            return;

                    } else if (ExpExc.choice.Equals(dgdMachine.Name))
                    {
                        DataStore.Instance.InsertLogByForm(this.GetType().Name, "E");
                        if (ExpExc.Check.Equals("Y"))
                            dt = Lib.Instance.DataGridToDTinHidden(dgdMachine);
                        else
                            dt = Lib.Instance.DataGirdToDataTable(dgdMachine);

                        Name = dgdMachine.Name;

                        if (Lib.Instance.GenerateExcel(dt, Name))
                            Lib.Instance.excel.Visible = true;
                        else
                            return;

                    } else if (ExpExc.choice.Equals(dgdNoWork.Name))
                        {
                            DataStore.Instance.InsertLogByForm(this.GetType().Name, "E");
                            if (ExpExc.Check.Equals("Y"))
                                dt = Lib.Instance.DataGridToDTinHidden(dgdNoWork);
                            else
                                dt = Lib.Instance.DataGirdToDataTable(dgdNoWork);

                            Name = dgdNoWork.Name;

                            if (Lib.Instance.GenerateExcel(dt, Name))
                                Lib.Instance.excel.Visible = true;
                            else
                                return;
                        }
                    {
                        if (dt != null)
                        {
                            dt.Clear();
                        }
                    }
                }
            }
            catch(Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        #endregion

        #region 콤보박스 
        private void SetComboBox()
        {

            ObservableCollection<CodeView> ovcProcess = ComboBoxUtil.Instance.GetWorkProcess(0, "");
            this.cboProcess.ItemsSource = ovcProcess;
            this.cboProcess.DisplayMemberPath = "code_name";
            this.cboProcess.SelectedValuePath = "code_id";

            if (cboProcess.SelectedValue != null)
            {
                ProcessID = cboProcess.SelectedValue.ToString();
            }

            ObservableCollection<CodeView> ovcMachine = ComboBoxUtil.Instance.GetMachine(ProcessID);
            this.cboMachine.ItemsSource = ovcMachine;
            this.cboMachine.DisplayMemberPath = "code_name";
            this.cboMachine.SelectedValuePath = "code_id";


        }
        #endregion

        #region 주요 메서드 - FillGrid

        //작업조회
        private void FillGrid_Work()
        {
            if (dgdWork.Items.Count > 0)
            {
                dgdWork.Items.Clear();
            }

            try
            {
                DataSet ds = null;
                Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                sqlParameter.Clear();
                sqlParameter.Add("chkDate", chkMcInOutDate.IsChecked == true ? 1 : 0);
                sqlParameter.Add("sFromDate", chkMcInOutDate.IsChecked == true && dtpSDate.SelectedDate != null ? dtpSDate.SelectedDate.Value.ToString("yyyyMMdd") : "");
                sqlParameter.Add("sToDate", chkMcInOutDate.IsChecked == true && dtpEDate.SelectedDate != null ? dtpEDate.SelectedDate.Value.ToString("yyyyMMdd") : "");
                sqlParameter.Add("chkProcess", chkProcessSrh.IsChecked == true ? 1 : 0);
                sqlParameter.Add("ProcessID", chkProcessSrh.IsChecked == true && cboProcess.SelectedValue != null ? cboProcess.SelectedValue.ToString() : "");
                sqlParameter.Add("chkMachineID", chkMachineSrh.IsChecked == true ? 1 : 0);
                sqlParameter.Add("MachineID", chkMachineSrh.IsChecked == true && cboMachine.SelectedValue != null ? cboMachine.SelectedValue.ToString() : "");
   
                ds = DataStore.Instance.ProcedureToDataSet_LogWrite("xp_prd_sMCRunningRate_Detail_Work", sqlParameter, true, "R");

                if (ds != null && ds.Tables.Count > 0)
                {
                    DataTable dt = ds.Tables[0];
                    int i = 0;

                    if (dt.Rows.Count > 0)
                    {
                        DataRowCollection drc = dt.Rows;

                        foreach (DataRow dr in drc)
                        {
                            i++;

                            var Work = new Work_CodeView()
                            {
                                Num = i,
                                WorkDate = dr["WorkDate"].ToString(),
                                ArticleID = dr["ArticleID"].ToString(),
                                Article = dr["Article"].ToString(),
                                WorkPersonID = dr["WorkPersonID"].ToString(),
                                Name = dr["Name"].ToString(),
                                EduCount = dr["EduCount"].ToString(),
                                License = dr["License"].ToString(),
                                LicenseCnt = dr["LicenseCnt"].ToString(),
                                StartDate = dr["StartDate"].ToString(),

                            };

                            dgdWork.Items.Add(Work);

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

        //호기조회
        private void FillGrid_Machine()
        {
            if (dgdMachine.Items.Count > 0)
            {
                dgdMachine.Items.Clear();
            }

            try
            {
                DataSet ds = null;
                Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                sqlParameter.Clear();
                sqlParameter.Add("chkDate", chkMcInOutDate.IsChecked == true ? 1 : 0);
                sqlParameter.Add("sFromDate", chkMcInOutDate.IsChecked == true && dtpSDate.SelectedDate != null ? dtpSDate.SelectedDate.Value.ToString("yyyyMMdd") : "");
                sqlParameter.Add("sToDate", chkMcInOutDate.IsChecked == true && dtpEDate.SelectedDate != null ? dtpEDate.SelectedDate.Value.ToString("yyyyMMdd") : "");
                sqlParameter.Add("chkProcess", chkProcessSrh.IsChecked == true ? 1 : 0);
                sqlParameter.Add("ProcessID", chkProcessSrh.IsChecked == true && cboProcess.SelectedValue != null ? cboProcess.SelectedValue.ToString() : "");
                sqlParameter.Add("chkMachineID", chkMachineSrh.IsChecked == true ? 1 : 0);
                sqlParameter.Add("MachineID", chkMachineSrh.IsChecked == true && cboMachine.SelectedValue != null ? cboMachine.SelectedValue.ToString() : "");

                ds = DataStore.Instance.ProcedureToDataSet_LogWrite("xp_prd_sMCRunningRate_Detail_MC", sqlParameter, true, "R");

                if (ds != null && ds.Tables.Count > 0)
                {
                    DataTable dt = ds.Tables[0];
                    int i = 0;

                    if (dt.Rows.Count > 0)
                    {
                        DataRowCollection drc = dt.Rows;

                        foreach (DataRow dr in drc)
                        {
                            i++;

                            var Machine = new MC_CodeView()
                            {
                                MCID = dr["MCID"].ToString(),
                                MCNAME = dr["MCNAME"].ToString(),
                                //MachineID = dr["MachineID"].ToString(),
                                MachineNo = dr["MachineNo"].ToString(),
                                LastInspectDate = dr["LastInspectDate"].ToString(),
                                InspectCount = dr["InspectCount"].ToString(),
                                DefectContents = dr["DefectContents"].ToString(),

                            };

                            dgdMachine.Items.Add(Machine);

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

        //비가동조회
        private void FillGrid_NoWork()
        {
            if (dgdNoWork.Items.Count > 0)
            {
                dgdNoWork.Items.Clear();
            }

            try
            {
                DataSet ds = null;
                Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                sqlParameter.Clear();
                sqlParameter.Add("chkDate", chkMcInOutDate.IsChecked == true ? 1 : 0);
                sqlParameter.Add("sFromDate", chkMcInOutDate.IsChecked == true && dtpSDate.SelectedDate != null ? dtpSDate.SelectedDate.Value.ToString("yyyyMMdd") : "");
                sqlParameter.Add("sToDate", chkMcInOutDate.IsChecked == true && dtpEDate.SelectedDate != null ? dtpEDate.SelectedDate.Value.ToString("yyyyMMdd") : "");
                sqlParameter.Add("chkProcess", chkProcessSrh.IsChecked == true ? 1 : 0);
                sqlParameter.Add("ProcessID", chkProcessSrh.IsChecked == true && cboProcess.SelectedValue != null ? cboProcess.SelectedValue.ToString() : "");
                sqlParameter.Add("chkMachineID", chkMachineSrh.IsChecked == true ? 1 : 0);
                sqlParameter.Add("MachineID", chkMachineSrh.IsChecked == true && cboMachine.SelectedValue != null ? cboMachine.SelectedValue.ToString() : "");

                ds = DataStore.Instance.ProcedureToDataSet_LogWrite("xp_prd_sMCRunningRate_Detail_NoWork", sqlParameter, true, "R");

                if (ds != null && ds.Tables.Count > 0)
                {
                    DataTable dt = ds.Tables[0];
                    int i = 0;

                    if (dt.Rows.Count > 0)
                    {
                        DataRowCollection drc = dt.Rows;

                        foreach (DataRow dr in drc)
                        {
                            i++;

                            var NoWork = new NoWork_CodeView()
                            {
                                MachineID = dr["MachineID"].ToString(),
                                WorkDate = dr["WorkDate"].ToString(),
                                NoReworkCode = dr["NoReworkCode"].ToString(),
                                NoReworkName = dr["NoReworkName"].ToString(),
                                NoReworkReason = dr["NoReworkReason"].ToString(),

                            };

                            dgdNoWork.Items.Add(NoWork);

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

        #endregion


        private void cboProcess_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if(cboProcess.SelectedValue != null) {
                ProcessID = cboProcess.SelectedValue.ToString();

                ObservableCollection<CodeView> ovcMachine = ComboBoxUtil.Instance.GetMachine(ProcessID);
                this.cboMachine.ItemsSource = ovcMachine;
                this.cboMachine.DisplayMemberPath = "code_name";
                this.cboMachine.SelectedValuePath = "code_id";
            }
           
        }


        private void dgdMachine_MouseLeftButtonDown(object sender, System.Windows.Input.MouseButtonEventArgs e)
        {
            if (e.ClickCount == 2)
            {
                var NoWorkingCode = dgdMachine.SelectedItem as MC_CodeView;

                string sDate = dtpSDate.SelectedDate.Value.ToString("yyyyMMdd");
                string eDate = dtpEDate.SelectedDate.Value.ToString("yyyyMMdd");

                NoWorkInfo NoWorking = null;

                if (NoWorkingCode != null)
                {
                    if (NoWorkingCode.NoWorkDate == null
                        || ConvertDouble(NoWorkingCode.NoWorkDate) == 0)
                        MessageBox.Show("선택된 자료의 비가동 시간을 확인해보세요.");
                    else
                        NoWorking = new NoWorkInfo(sDate, eDate, NoWorkingCode.MCID);
                }
                else
                {
                    NoWorking = new NoWorkInfo(sDate, eDate, "");
                }

                if (NoWorking != null)
                    NoWorking.ShowDialog();
            }
        }

        #region 기타 메서드 모음

        // 천마리 콤마, 소수점 버리기
        private string stringFormatN0(object obj)
        {
            return string.Format("{0:N0}", obj);
        }

        private string stringFormatN1(object obj)
        {
            return string.Format("{0:N1}", obj);
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

    public class Work_CodeView : BaseView
    {
 
        public int Num { get; set; }
        public string WorkDate { get; set; }
        public string ArticleID { get; set; }
        public string Article { get; set; }
        public string WorkPersonID { get; set; }
        public string Name { get; set; }
        public string EduCount { get; set; }
        public string MachineID { get; set; }
        public string StartDate { get; set; }
        public string License { get; set; }
        public string LicenseCnt { get; set; }


    }

    public class MC_CodeView : BaseView
    {
        public string MCID { get; set; }
        public string MCNAME { get; set; }
        public string MachineID { get; set; }
        public string MachineNo { get; set; }
        public string LastInspectDate { get; set; }
        public string InspectCount { get; set; }
        public string DefectContents { get; set; }
        public string NoWorkDate { get; set; }
    }

    public class NoWork_CodeView : BaseView
    {
        public string WorkDate { get; set; }
        public string NoReworkCode { get; set; }
        public string NoReworkName { get; set; }
        public string NoReworkReason { get; set; }
        public string MachineID { get; set; }
    }
}

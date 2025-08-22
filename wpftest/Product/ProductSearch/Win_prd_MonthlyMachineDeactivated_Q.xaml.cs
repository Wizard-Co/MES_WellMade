using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using WizMes_WellMade.PopUp;
using WizMes_WellMade.PopUP;

namespace WizMes_WellMade
{
    /// <summary>
    /// Win_mtr_Subul_Q_New.xaml에 대한 상호 작용 논리
    /// </summary>
    public partial class Win_prd_MonthlyMachineDeactivated_Q : UserControl  
    {
        int rowNum = 0;
        Lib lib = new Lib();

        public Win_prd_MonthlyMachineDeactivated_Q()
        {
            InitializeComponent();
        }

        private void UserControl_Loaded(object sender, RoutedEventArgs e)
        {
            SetComboBox();

            chkDateSrh.IsChecked = true;

            dtpEDate.SelectedDate = DateTime.Today;
            dtpFDate.SelectedDate = DateTime.Today;
            cboProcess.SelectedIndex = 0;
            cboMachine.SelectedIndex = 0;
        }

        #region 콤보박스 세팅 SetComboBox()

        private void SetComboBox()
        {
            ObservableCollection<CodeView> cbProcess = ComboBoxUtil.Instance.GetWorkProcess(0, "");

            this.cboProcess.ItemsSource = cbProcess;
            this.cboProcess.DisplayMemberPath = "code_name";
            this.cboProcess.SelectedValuePath = "code_id";

            ObservableCollection<CodeView> ovcMachine = GetMachineByProcessID("");
            this.cboMachine.ItemsSource = ovcMachine;
            this.cboMachine.DisplayMemberPath = "code_name";
            this.cboMachine.SelectedValuePath = "code_id";
        }

        #endregion

        #region 프로세스 콤보박스 변경시 → 설비 콤보박스 세팅!!!!!!!!!!

        private void cboProcess_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (cboProcess.SelectedValue != null 
                && !cboProcess.SelectedValue.ToString().Equals(string.Empty)
                && cboProcess.SelectedIndex != 0)
            {
                string strProcess = cboProcess.SelectedValue.ToString();
                this.cboMachine.ItemsSource = null;

                ObservableCollection<CodeView> ovcMachine = GetMachineByProcessID(cboProcess.SelectedValue.ToString());
                this.cboMachine.ItemsSource = ovcMachine;
                this.cboMachine.DisplayMemberPath = "code_name";
                this.cboMachine.SelectedValuePath = "code_id";
            }
            else if (cboProcess.SelectedValue != null
                && cboProcess.SelectedIndex == 0)
            {
                ObservableCollection<CodeView> ovcMachine = GetMachineByProcessID("");
                this.cboMachine.ItemsSource = ovcMachine;
                this.cboMachine.DisplayMemberPath = "code_name";
                this.cboMachine.SelectedValuePath = "code_id";
            }
        }

        #region mt_Machine - 호기 세팅

        /// <summary>
        /// 호기ID 가져오기
        /// </summary>
        /// <param name="value"></param>
        /// <returns></returns>
        public ObservableCollection<CodeView> GetMachineByProcessID(string value)
        {
            //2021-10-25 공정 콤보박스에 전체가 선택되면 호기 공정 콤보박스 안되게 막기
            if (value.Equals(""))
            {
                cboMachine.IsEnabled = false;
            }
            else
            {
                cboMachine.IsEnabled = true;
            }

            ObservableCollection<CodeView> ovcMachine = new ObservableCollection<CodeView>();

            Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
            sqlParameter.Add("sProcessID", value);

            DataSet ds = DataStore.Instance.ProcedureToDataSet("xp_Process_sMachineForComboBoxAndUsing", sqlParameter, false);

            if (ds != null && ds.Tables.Count > 0)
            {
                DataTable dt = ds.Tables[0];
                if (dt.Rows.Count > 0)
                {
                    CodeView CV = new CodeView();
                    CV.code_id = "";
                    CV.code_name = "전체";
                    ovcMachine.Add(CV);

                    DataRowCollection drc = dt.Rows;

                    foreach (DataRow dr in drc)
                    {
                        CodeView mCodeView = new CodeView()
                        {
                            code_id = dr["Code"].ToString().Trim(),
                            code_name = dr["Name"].ToString().Trim()
                        };

                        ovcMachine.Add(mCodeView);
                    }
                }
            }

            return ovcMachine;
        }

        #endregion // mt_Machine - 호기 세팅

        #endregion

        #region Header 부분 - 검색조건

        // 일자
        private void lblDateSrh_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            if (chkDateSrh.IsEnabled == true)
            {
                if (chkDateSrh.IsChecked == true)
                {
                    chkDateSrh.IsChecked = false;
                }
                else
                {
                    chkDateSrh.IsChecked = true;
                }
            }
        }
        private void chkDateSrh_Checked(object sender, RoutedEventArgs e)
        {
            if (chkDateSrh.IsEnabled == true)
            {
                chkDateSrh.IsChecked = true;
                dtpEDate.IsEnabled = true;
                dtpFDate.IsEnabled = true;
            }
        }
        private void chkDateSrh_Unchecked(object sender, RoutedEventArgs e)
        {
            if (chkDateSrh.IsEnabled == true)
            {
                chkDateSrh.IsChecked = false;
                dtpEDate.IsEnabled = false;
                dtpFDate.IsEnabled = false;
            }
        }


        //공정 검색
        private void lblProcess_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            if (chkProcess.IsChecked == true)
            {
                chkProcess.IsChecked = false;
            }
            else
            {
                chkProcess.IsChecked = true;
            }
        }
        private void chkProcess_Checked(object sender, RoutedEventArgs e)
        {
            chkProcess.IsChecked = true;
            cboProcess.IsEnabled = true;
        }
        private void chkProcess_Unchecked(object sender, RoutedEventArgs e)
        {
            chkProcess.IsChecked = false;
            cboProcess.IsEnabled = false;
        }

        // 작업자 검색
        private void lblMachine_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            if (chkMachine.IsChecked == true)
            {
                chkMachine.IsChecked = false;
            }
            else
            {
                chkMachine.IsChecked = true;
            }
        }
        private void chkMachine_Checked(object sender, RoutedEventArgs e)
        {
            chkMachine.IsChecked = true;
            cboMachine.IsEnabled = true;
        }
        private void chkMachine_Unchecked(object sender, RoutedEventArgs e)
        {
            chkMachine.IsChecked = false;
            cboMachine.IsEnabled = false;
        }

        #endregion // Header 부분 - 검색조건

        #region Header 부분 - 오른쪽 버튼 모음 (검색, 닫기, 엑셀)

        // 검색버튼
        private void btnSearch_Click(object sender, RoutedEventArgs e)
        {
            //검색버튼 비활성화
            btnSearch.IsEnabled = false;

            Dispatcher.BeginInvoke(new Action(() =>

            {
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
            re_search(rowNum);
        }

        // 닫기버튼
        private void btnClose_Click(object sender, RoutedEventArgs e)
        {
            Lib.Instance.ChildMenuClose(this.ToString());
        }

        // 엑셀버튼
        private void btnExcel_Click(object sender, RoutedEventArgs e)
        {
            DataTable dt = null;
            string Name = string.Empty;

            string[] lst = new string[2];
            lst[0] = "월별 설비별 비가동 집계";
            lst[1] = dgdMain.Name;

            ExportExcelxaml ExpExc = new ExportExcelxaml(lst);

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

        #endregion // Header 부분 - 오른쪽 버튼 모음 (검색, 닫기, 엑셀)

        private void re_search(int selectedIndex)
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

        #region 조회 메서드

        private void FillGrid()
        {
            dgdTotal.Items.Clear();
            
            if (dgdMain.Items.Count > 0)
            {
                dgdMain.Items.Clear();
            }

            try
            {
                // 공정 호기 세팅
                string ProcessID = "";
                string MachineID = "";

                // 공정을 전체나 선택하지 않았을시 → 호기는 공정 + 호기로 출력 → 공정과 호기를 검색하기 위해서
                if (chkMachine.IsChecked == true
                    && cboMachine.SelectedValue != null
                    && cboMachine.SelectedValue.ToString().Trim().Length == 6)
                {
                    ProcessID = cboMachine.SelectedValue.ToString().Trim().Substring(0, 4);
                    MachineID = cboMachine.SelectedValue.ToString().Trim().Substring(4, 2);
                }
                else
                {
                    ProcessID = chkProcess.IsChecked == true && cboProcess.SelectedValue != null ? cboProcess.SelectedValue.ToString() : "";
                    MachineID = chkMachine.IsChecked == true && cboMachine.SelectedValue != null ? cboMachine.SelectedValue.ToString() : "";
                }


                Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                sqlParameter.Clear();

                sqlParameter.Add("BasisMonth", dtpEDate.SelectedDate != null ? dtpEDate.SelectedDate.Value.ToString("yyyyMMdd") : "");
                sqlParameter.Add("BasisMonthday", dtpFDate.SelectedDate != null ? dtpFDate.SelectedDate.Value.ToString("yyyyMMdd") : "");
                sqlParameter.Add("chkProcessID", chkProcess.IsChecked == true ? 1 : 0);
                sqlParameter.Add("ProcessID", ProcessID);

                sqlParameter.Add("chkMachineID", chkMachine.IsChecked == true ? 1 : 0);
                sqlParameter.Add("MachineID", MachineID);


                DataSet ds = DataStore.Instance.ProcedureToDataSet("xp_prd_sProdNoWorkSumMachine", sqlParameter, false); 
                if (ds != null && ds.Tables.Count > 0)
                {
                    DataTable dt = ds.Tables[0];
                    int i = 0;

                    if (dt.Rows.Count > 1)
                    {
                        DataRowCollection drc = dt.Rows;

                        foreach (DataRow dr in drc)
                        {
                            //if (dr["ProcessID"].ToString().Trim().Equals(""))
                            //{
                            //    continue;
                            //}

                            i++;

                            var WinR = new Win_prd_MonthlyMachineDeactivated_Q_CodeView()
                            {
                                Num = i.ToString(),

                                cls = dr["cls"].ToString(),
                                Workdate = dr["Workdate"].ToString(),
                                ProcessID = dr["ProcessID"].ToString(),
                                Process = dr["Process"].ToString(),
                                AutoMcYNName = dr["AutoMcYNName"].ToString(),
                                Machine = dr["Machine"].ToString(),
                                Machineno = dr["Machineno"].ToString(),

                                NoReworkName1 = dr["NoReworkName1"].ToString(),
                                NoWorkHour1 = stringFormatN2(dr["NoWorkHour1"]),

                                NoReworkName2 = dr["NoReworkName2"].ToString(),
                                NoWorkHour2 = stringFormatN2(dr["NoWorkHour2"]),

                                NoReworkName3 = dr["NoReworkName3"].ToString(),
                                NoWorkHour3 = stringFormatN2(dr["NoWorkHour3"]),

                                NoReworkName4 = dr["NoReworkName4"].ToString(),
                                NoWorkHour4 = stringFormatN2(dr["NoWorkHour4"]),

                                NoReworkName5 = dr["NoReworkName5"].ToString(),
                                NoWorkHour5 = stringFormatN2(dr["NoWorkHour5"]),

                                NoReworkName6 = dr["NoReworkName6"].ToString(),
                                NoWorkHour6 = stringFormatN2(dr["NoWorkHour6"]),

                                NoReworkName7 = dr["NoReworkName7"].ToString(),
                                NoWorkHour7 = stringFormatN2(dr["NoWorkHour7"]),

                                NoReworkName8 = dr["NoReworkName8"].ToString(),
                                NoWorkHour8 = stringFormatN2(dr["NoWorkHour8"]),

                                NoReworkName9 = dr["NoReworkName9"].ToString(),
                                NoWorkHour9 = stringFormatN2(dr["NoWorkHour9"]),

                                NoReworkName10 = dr["NoReworkName10"].ToString(),
                                NoWorkHour10 = stringFormatN2(dr["NoWorkHour10"]),

                                TotNoWorkHour = stringFormatN2(dr["TotNoWorkHour"]),
                                TotWorkHour = stringFormatN2(dr["TotWorkHour"]),
                                TotNoWorkRate = stringFormatN2(dr["TotNoWorkRate"]),
                            };

                            WinR.lstColumnName = new List<string>();

                            for (int k = 0; k < 10; k++)
                            {
                                WinR.lstColumnName.Add(dr["NoReworkName" + (k + 1)].ToString());
                            }

                            setColumnHeader(WinR);

                            if (WinR.cls.Trim().Equals("9"))
                            {
                                
                                WinR.Process = "총계";
                                WinR.Total_Color = true;
                                dgdTotal.Items.Add(WinR);
                            }
                            else
                            {
                                dgdMain.Items.Add(WinR);
                            }


                        }

                        setColumnVisible();
                    }

                    //tbkCount.Text = " ▶ 검색 결과 : " + i + " 건";
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

        #region 컬럼 세팅
        private void setColumnHeader(Win_prd_MonthlyMachineDeactivated_Q_CodeView Win)
        {
            for (int i = 0; i < Win.lstColumnName.Count; i++)
            {
                DataGridColumn dgc = dgdMain.Columns[i + 5];

                if (!Win.lstColumnName[i].Trim().Equals(""))
                {
                    dgc.Header = Win.lstColumnName[i].ToString();
                }
            }

            //if (!Win.NoReworkName1.Trim().Equals(""))
            //{
            //    dtcNoWork1.Header = Win.NoReworkName1;
            //}
            //if (!Win.NoReworkName2.Trim().Equals(""))
            //{
            //    dtcNoWork2.Header = Win.NoReworkName2;
            //}
            //if (!Win.NoReworkName3.Trim().Equals(""))
            //{
            //    dtcNoWork3.Header = Win.NoReworkName3;
            //}
            //if (!Win.NoReworkName4.Trim().Equals(""))
            //{
            //    dtcNoWork4.Header = Win.NoReworkName4;
            //}
            //if (!Win.NoReworkName5.Trim().Equals(""))
            //{
            //    dtcNoWork5.Header = Win.NoReworkName5;
            //}
            //if (!Win.NoReworkName6.Trim().Equals(""))
            //{
            //    dtcNoWork6.Header = Win.NoReworkName6;
            //}
            //if (!Win.NoReworkName7.Trim().Equals(""))
            //{
            //    dtcNoWork7.Header = Win.NoReworkName7;
            //}
            //if (!Win.NoReworkName8.Trim().Equals(""))
            //{
            //    dtcNoWork8.Header = Win.NoReworkName8;
            //}
            //if (!Win.NoReworkName9.Trim().Equals(""))
            //{
            //    dtcNoWork9.Header = Win.NoReworkName9;
            //}
            //if (!Win.NoReworkName10.Trim().Equals(""))
            //{
            //    dtcNoWork10.Header = Win.NoReworkName10;
            //}
        }

        #endregion

        #region 컬럼 헤더 없는것 = 해당 데이터가 없음. → 컬럼 숨기기

        private void setColumnVisible()
        {
            foreach( DataGridColumn dgc in dgdMain.Columns)
            {
                if (dgc.Header.ToString().Trim().Equals(""))
                {
                    dgc.Visibility = Visibility.Hidden;
                }
                else
                {
                    dgc.Visibility = Visibility.Visible;
                }
            }

            //if (dtcNoWork1.Header.ToString().Trim().Equals(""))
            //{
            //    dtcNoWork1.Visibility = Visibility.Hidden;
            //}
            //if(dtcNoWork2.Header.ToString().Trim().Equals(""))
            //{
            //    dtcNoWork2.Visibility = Visibility.Hidden;
            //}
            //if (dtcNoWork3.Header.ToString().Trim().Equals(""))
            //{
            //    dtcNoWork3.Visibility = Visibility.Hidden;
            //}
            //if (dtcNoWork4.Header.ToString().Trim().Equals(""))
            //{
            //    dtcNoWork4.Visibility = Visibility.Hidden;
            //}
            //if (dtcNoWork5.Header.ToString().Trim().Equals(""))
            //{
            //    dtcNoWork5.Visibility = Visibility.Hidden;
            //}
            //if (dtcNoWork6.Header.ToString().Trim().Equals(""))
            //{
            //    dtcNoWork6.Visibility = Visibility.Hidden;
            //}
            //if (dtcNoWork7.Header.ToString().Trim().Equals(""))
            //{
            //    dtcNoWork7.Visibility = Visibility.Hidden;
            //}
            //if (dtcNoWork8.Header.ToString().Trim().Equals(""))
            //{
            //    dtcNoWork8.Visibility = Visibility.Hidden;
            //}
            //if (dtcNoWork9.Header.ToString().Trim().Equals(""))
            //{
            //    dtcNoWork9.Visibility = Visibility.Hidden;
            //}
            //if (dtcNoWork10.Header.ToString().Trim().Equals(""))
            //{
            //    dtcNoWork10.Visibility = Visibility.Hidden;
            //}
        }

        #endregion

        #endregion // 조회 메서드

        #region 기타 메서드 모음

        // 천마리 콤마, 소수점 버리기
        private string stringFormatN0(object obj)
        {
            return string.Format("{0:N0}", obj);
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

        // 시간 형식 6글자라면! 11:11:11
        private string DateTimeFormat(string str)
        {
            str = str.Replace(":", "").Trim();

            if (str.Length == 6)
            {
                string Hour = str.Substring(0, 2);
                string Min = str.Substring(2, 2);
                string Sec = str.Substring(4, 2);

                str = Hour + ":" + Min + ":" + Sec;
            }

            return str;
        }

        // 시간 분 → 11:12 형식으로 변환
        private string DateTimeHourToTimeFormat(string str)
        {
            str = str.Replace(":", "").Trim();

            double num = 0;
            if (double.TryParse(str, out num) == true)
            {
                int Min = (int)(num * 60);

                string hour = (Min / 60).ToString();

                if (num < 0) { hour = "00"; }

                string min = (Min % 60).ToString();

                if (min.Length == 1)
                {
                    min = "0" + min;
                }

                str = hour + ":" + min;
            }

            return str;
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

        private void dgdMain_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {

        }
    }

    class Win_prd_MonthlyMachineDeactivated_Q_CodeView
    {
        public string Num { get; set; }
        public string cls { get; set; }
        public string Workdate { get; set; }
        public string ProcessID { get; set; }
        public string Process { get; set; }
        public string AutoMcYNName { get; set; }
        public string Machine { get; set; }
        public string Machineno { get; set; }

        public string NoReworkName1 { get; set; }
        public string NoWorkHour1 { get; set; }

        public string NoReworkName2 { get; set; }
        public string NoWorkHour2 { get; set; }

        public string NoReworkName3 { get; set; }
        public string NoWorkHour3 { get; set; }

        public string NoReworkName4 { get; set; }
        public string NoWorkHour4 { get; set; }

        public string NoReworkName5 { get; set; }
        public string NoWorkHour5 { get; set; }

        public string NoReworkName6 { get; set; }
        public string NoWorkHour6 { get; set; }

        public string NoReworkName7 { get; set; }
        public string NoWorkHour7 { get; set; }

        public string NoReworkName8 { get; set; }
        public string NoWorkHour8 { get; set; }

        public string NoReworkName9 { get; set; }
        public string NoWorkHour9 { get; set; }

        public string NoReworkName10 { get; set; }
        public string NoWorkHour10 { get; set; }

        public string TotNoWorkHour { get; set; }
        public string TotWorkHour { get; set; }
        public string TotNoWorkRate { get; set; }

        public List<string> lstColumnName { get; set; }

        public bool Total_Color { get; set; }
    }
}

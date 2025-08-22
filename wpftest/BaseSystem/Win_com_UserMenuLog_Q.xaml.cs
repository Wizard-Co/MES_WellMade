using System;
using System.Collections.Generic;
using System.Data;
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
using WPF.MDI;

namespace WizMes_WellMade
{
    /// <summary>
    /// Win_sys_UserMenuLog_Q.xaml에 대한 상호 작용 논리
    /// </summary>
    public partial class Win_com_UserMenuLog_Q : UserControl
    {
        Lib lib = new Lib();

        string stDate = string.Empty;
        string stTime = string.Empty;

        public Win_com_UserMenuLog_Q()
        {
            InitializeComponent();
        }

        private void UserControl_Loaded(object sender, RoutedEventArgs e)
        {
            chkDate.IsChecked = true;
            btnToday_Click(null, null);

            stDate = DateTime.Now.ToString("yyyyMMdd");
            stTime = DateTime.Now.ToString("HHmm");

            DataStore.Instance.InsertLogByFormS(this.GetType().Name, stDate, stTime, "S");
        }

        #region 상단 검색조건 모음

        //사용일자
        private void lblDate_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            if (chkDate.IsChecked == true) 
            {
                chkDate.IsChecked = false; 
            }
            else 
            {
                chkDate.IsChecked = true; 
            }
        }

        //사용일자 체크박스 체크
        private void chkDate_Checked(object sender, RoutedEventArgs e)
        {
            dtpSDate.IsEnabled = true;
            dtpEDate.IsEnabled = true;
        }

        //사용일자 체크박스 체크해제
        private void chkDate_Unchecked(object sender, RoutedEventArgs e)
        {
            dtpSDate.IsEnabled = false;
            dtpEDate.IsEnabled = false;
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

        //전월
        private void btnLastMonth_Click(object sender, RoutedEventArgs e)
        {
            DateTime[] SearchDate = lib.BringLastMonthContinue(dtpSDate.SelectedDate.Value);

            dtpSDate.SelectedDate = SearchDate[0];
            dtpEDate.SelectedDate = SearchDate[1];
        }

        //금월
        private void btnThisMonth_Click(object sender, RoutedEventArgs e)
        {
            dtpSDate.SelectedDate = Lib.Instance.BringThisMonthDatetimeList()[0];
            dtpEDate.SelectedDate = Lib.Instance.BringThisMonthDatetimeList()[1];
        }

        //사원명
        private void lblPersonName_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            if (chkPersonName.IsChecked == true) 
            {
                chkPersonName.IsChecked = false; 
            }
            else 
            {
                chkPersonName.IsChecked = true; 
            }
        }

        //사원명 체크박스 체크
        private void chkPersonName_Checked(object sender, RoutedEventArgs e)
        {
            txtPersonName.IsEnabled = true;
            //btnPfPersonName.IsEnabled = true;
            txtPersonName.Focus();
        }

        //사원명 체크박스 체크해제
        private void chkPersonName_Unchecked(object sender, RoutedEventArgs e)
        {
            txtPersonName.IsEnabled = false;
            //btnPfPersonName.IsEnabled = false;
        }

        //사원명 텍스트박스 키다운 이벤트
        private void txtPersonName_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                MainWindow.pf.ReturnCode(txtPersonName, (int)Defind_CodeFind.DCF_PERSON, "");
            }
        }

        //사원명 플러스파인더 버튼 클릭
        private void btnPfPersonName_Click(object sender, RoutedEventArgs e)
        {
            MainWindow.pf.ReturnCode(txtPersonName, (int)Defind_CodeFind.DCF_PERSON, "");
        }

        #endregion

        //검색
        private void btnSearch_Click(object sender, RoutedEventArgs e)
        {
            if (dgdMain.Items.Count > 0)
            {
                dgdMain.Items.Clear();
            }

            FillGrid();

            if (dgdMain.Items.Count > 0)
            {
                dgdMain.SelectedIndex = 0;
            }
        }

        //닫기
        private void btnClose_Click(object sender, RoutedEventArgs e)
        {
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

            //화면종료 로그 남기기 2022-06-23 추가
            DataStore.Instance.InsertLogByFormS(this.GetType().Name, stDate, stTime, "E");
        }

        //엑셀
        private void btnExcel_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                DataTable dt = null;
                string Name = string.Empty;

                string[] dgdStr = new string[2];
                dgdStr[0] = "사용자 Log조회";
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
            catch (Exception ee)
            {
                MessageBox.Show("오류지점 - 엑셀버튼 클릭 : " + ee.ToString());
            }
        }
        private void DataGrid_SizeChange(object sender, SizeChangedEventArgs e)
        {
            DataGrid dgs = sender as DataGrid;
            if(dgs.ColumnHeaderHeight == 0)
            {
                dgs.ColumnHeaderHeight = 1;
            }
            double a = e.NewSize.Height / 100;
            double b = e.PreviousSize.Height / 100;
            double c = a / b;

            if(c !=double.PositiveInfinity && c != 0 && double.IsNaN(c) == false)
            {
                dgs.ColumnHeaderHeight = dgs.ColumnHeaderHeight * c;
                dgs.FontSize = dgs.FontSize * c;
            }
        }
        

        //실조회
        private void FillGrid()
        {
            try
            {
                Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                sqlParameter.Add("sCompanyID", MainWindow.CompanyID);
                sqlParameter.Add("ChkDate", chkDate.IsChecked == true ? 1 : 0);
                sqlParameter.Add("sFromDate", chkDate.IsChecked == true ? dtpSDate.SelectedDate.Value.ToString("yyyyMMdd") : "");
                sqlParameter.Add("sToDate", chkDate.IsChecked == true ? dtpEDate.SelectedDate.Value.ToString("yyyyMMdd") : "");
                sqlParameter.Add("ChkPerson", chkPersonName.IsChecked == true ? 1:0);
                sqlParameter.Add("sPerson", chkPersonName.IsChecked == false || txtPersonName.Text == null || txtPersonName.Text.Trim().Equals("") ? "" : txtPersonName.Text);

                DataSet ds = DataStore.Instance.ProcedureToDataSet_LogWrite("xp_Common_sLogData_New", sqlParameter, true, "R");

                if (ds != null && ds.Tables.Count > 0)
                {
                    DataTable dt = ds.Tables[0];

                    if (dt.Rows.Count == 0)
                    {
                        MessageBox.Show("조회된 데이터가 없습니다.");
                    }
                    else
                    {
                        DataRowCollection drc = dt.Rows;
                        string pickDate = "";
                        string pickMenu = "";

                        pickDate = drc[0]["WorkDate"].ToString();
                        pickMenu = drc[0]["Menu"].ToString();

                        for (int i = 0; i < drc.Count; i++)
                        {
                            DataRow dr = drc[i];

                            var WinUserMenuLog = new Win_sys_UserMenuLog_Q_CodeView()
                            {
                                Num = (i + 1),
                                WorkDate = dr["WorkDate"].ToString(),
                                //WorkTime = dr["WorkTime"].ToString(),
                                //PersonID = dr["PersonID"].ToString(),
                                UserID = dr["UserID"].ToString(),
                                Name = dr["Name"].ToString(),
                                //MenuID = dr["MenuID"].ToString(),
                                Menu = dr["Menu"].ToString(),
                                StartDate = dr["StartDate"].ToString(),
                                StartTime = dr["StartTime"].ToString(),
                                EndDate = dr["EndDate"].ToString(),
                                EndTime = dr["EndTime"].ToString(),
                                Duration = stringFormatN0(dr["Duration"]),
                                C = stringFormatN0(dr["C"]),
                                R = stringFormatN0(dr["R"]),
                                U = stringFormatN0(dr["U"]),
                                D = stringFormatN0(dr["D"])
                            };

                            //처음에는 무조건 CRUD 넣는다. 
                            if (i > 0)
                            {
                                if (pickDate.Equals(WinUserMenuLog.WorkDate) && pickMenu.Equals(WinUserMenuLog.Menu))
                                {
                                    WinUserMenuLog.C = "";
                                    WinUserMenuLog.R = "";
                                    WinUserMenuLog.U = "";
                                    WinUserMenuLog.D = "";
                                }
                                else
                                {
                                    pickDate = drc[i]["WorkDate"].ToString();
                                    pickMenu = drc[i]["Menu"].ToString();
                                }
                            }

                            WinUserMenuLog.WorkDate_CV = WinUserMenuLog.WorkDate != null ? Lib.Instance.StrDateTimeBar(WinUserMenuLog.WorkDate) : ""; 

                            WinUserMenuLog.StartDate_CV = WinUserMenuLog.StartDate != null ? DatePickerFormat(dr["StartDate"].ToString()) : ""; 
                            WinUserMenuLog.StartTime_CV = WinUserMenuLog.StartTime != null ? DateTimeFormat(WinUserMenuLog.StartTime) : "";

                            WinUserMenuLog.EndDate_CV = WinUserMenuLog.EndDate != null ? DatePickerFormat(dr["EndDate"].ToString()) : "";
                            WinUserMenuLog.EndTime_CV = WinUserMenuLog.EndTime != null ? DateTimeFormat(WinUserMenuLog.EndTime) : "";

                            


                            dgdMain.Items.Add(WinUserMenuLog);

                        }

                        tbkCount.Text = "▶ 검색 결과 : " + dt.Rows.Count.ToString() + " 건";
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("오류지점 - 조회 프로시저 : " + ex.ToString());
            }
            finally
            {
                DataStore.Instance.CloseConnection();
            }
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

        // 시간 형식
        private string DateTimeFormat(string str)
        {
            str = str.Replace(":", "").Trim();

            if (str.Length == 4)
            {
                string Hour = str.Substring(0, 2);
                string Min = str.Substring(2, 2);

                str = Hour + ":" + Min ;
            }

            return str;
        }

        
        private string stringFormatN0(object obj)
        {
            return string.Format("{0:N0}", obj);
        }
    }

    class Win_sys_UserMenuLog_Q_CodeView : BaseView
    {
        public int Num { get; set; }
        public string WorkDate { get; set; }
        public string WorkDate_CV { get; set; }
        public string WorkTime { get; set; }
        public string PersonID { get; set; }
        public string UserID { get; set; }
        public string Name { get; set; }
        public string MenuID { get; set; }
        public string Menu { get; set; }
        public string StartDate { get; set; }
        public string StartDate_CV { get; set; }
        public string StartTime { get; set; }
        public string StartTime_CV { get; set; }
        public string EndDate { get; set; }
        public string EndDate_CV { get; set; }
        public string EndTime { get; set; }
        public string EndTime_CV { get; set; }
        public string Duration { get; set; }
        public string C { get; set; }
        public string R { get; set; }
        public string U { get; set; }
        public string D { get; set; }
    }
}

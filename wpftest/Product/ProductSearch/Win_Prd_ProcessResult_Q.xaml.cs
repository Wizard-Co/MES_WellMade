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
using WPF.MDI;

namespace WizMes_WellMade
{
    /// <summary>
    /// Win_Prd_ProcessResult_Q.xaml에 대한 상호 작용 논리
    /// </summary>
    public partial class Win_Prd_ProcessResult_Q : UserControl
    {
        string stDate = string.Empty;
        string stTime = string.Empty;

        Lib lib = new Lib();
        List<string> lstScanDate = new List<string>();
        List<string> lstProcessID = new List<string>();
        //Dictionary<string, Win_Prd_ProcessResult_Q_Process> dicProcessResult = new Dictionary<string, Win_Prd_ProcessResult_Q_Process>();
        List<Win_Prd_ProcessResult_Q_CodeView> lstProcessResult_Q = new List<Win_Prd_ProcessResult_Q_CodeView>();

        public Win_Prd_ProcessResult_Q()
        {
            InitializeComponent();
        }

        private void UserControl_Loaded(object sender, RoutedEventArgs e)
        {
            stDate = DateTime.Now.ToString("yyyyMMdd");
            stTime = DateTime.Now.ToString("HHmm");

            DataStore.Instance.InsertLogByFormS(this.GetType().Name, stDate, stTime, "S");

            Lib.Instance.UiLoading(sender);
            dtpSDate.SelectedDate = DateTime.Today;
            dtpEDate.SelectedDate = DateTime.Today;

            setComboBox();

            cboProcess.SelectedIndex = 0;
            cboMachine.SelectedIndex = 0;
        }

        private void setComboBox()
        {
            ObservableCollection<CodeView> ovcProcess = ComboBoxUtil.Instance.GetWorkProcess(0, "");
            this.cboProcess.ItemsSource = ovcProcess;
            this.cboProcess.DisplayMemberPath = "code_name";
            this.cboProcess.SelectedValuePath = "code_id";

            ObservableCollection<CodeView> ovcMachine = GetMachineByProcessID("");
            this.cboMachine.ItemsSource = ovcMachine;
            this.cboMachine.DisplayMemberPath = "code_name";
            this.cboMachine.SelectedValuePath = "code_id";

            List<string> strCombo1 = new List<string>();
            strCombo1.Add("전체");
            strCombo1.Add("정상");
            strCombo1.Add("무작업");
            strCombo1.Add("재작업");

            ObservableCollection<CodeView> ovcGugunSearch = ComboBoxUtil.Instance.Direct_SetComboBox(strCombo1);
            this.cboGubun.ItemsSource = ovcGugunSearch;
            this.cboGubun.DisplayMemberPath = "code_name";
            this.cboGubun.SelectedValuePath = "code_id";
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
            if (value == "")
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

            DataSet ds = DataStore.Instance.ProcedureToDataSet("xp_prd_sMachineForComboBoxAndUsing", sqlParameter, false);

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

        #region 날짜버튼 클릭 이벤트

        // 전일 금일 전월 금월 버튼
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

        #endregion

        #region 체크박스&&라디오 버튼 이벤트
       
        //공정명
        private void lblProcess_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            //if (chkProcess.IsChecked == true) { chkProcess.IsChecked = false; }
            //else { chkProcess.IsChecked = true; }

            //chkProcessClick();
        }

        //공정명 클릭시
        private void chkProcess_Click(object sender, RoutedEventArgs e)
        {
            //chkProcessClick();
        }

        //공정명 이벤트
        private void chkProcessClick()
        {
            if (chkProcess.IsChecked == true)
            {
                cboProcess.IsEnabled = true;
                cboProcess.Focus();
            }
            else
            {
                cboProcess.IsEnabled = false;
            }
        }

        private void cboProcess_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (cboProcess.SelectedValue != null)
            {
                ObservableCollection<CodeView> ovcMachine = GetMachineByProcessID(cboProcess.SelectedValue.ToString());
                this.cboMachine.ItemsSource = ovcMachine;
                this.cboMachine.DisplayMemberPath = "code_name";
                this.cboMachine.SelectedValuePath = "code_id";

                cboMachine.SelectedIndex = 0;
            }
        }

        //호기
        private void lblMachine_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            //if (chkMachine.IsChecked == true) { chkMachine.IsChecked = false; }
            //else { chkMachine.IsChecked = true; }

            //chkMachineClick();
        }

        //호기 클릭시
        private void chkMachine_Click(object sender, RoutedEventArgs e)
        {
            //chkMachineClick();
        }

        //호기 이벤트
        private void chkMachineClick()
        {
            if (chkMachine.IsChecked == true)
            {
                cboMachine.IsEnabled = true;
                cboMachine.Focus();
            }
            else
            {
                cboMachine.IsEnabled = false;
            }
        }

        //작업구분
        private void lblGubun_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            if (chkGubun.IsChecked == true) { chkGubun.IsChecked = false; }
            else { chkGubun.IsChecked = true; }

            GubunClick();
        }

        //작업구분 클릭시
        private void chkGubun_Click(object sender, RoutedEventArgs e)
        {
            GubunClick();
        }

        //작업구분 이벤트
        private void GubunClick()
        {
            if (chkGubun.IsChecked == true)
            {
                cboGubun.IsEnabled = true;
                cboGubun.Focus();
            }
            else
            {
                cboGubun.IsEnabled = false;
            }
        }

        //OrderNo
        private void lblOnlyDefect_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            if (chkDefect.IsChecked == true) { chkDefect.IsChecked = false; }
            else { chkDefect.IsChecked = true; }
        }

        //OrderNo
        private void lblOrder_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            if (chkOrder.IsChecked == true) { chkOrder.IsChecked = false; }
            else { chkOrder.IsChecked = true; }
        }

        //OrderNo
        private void chkOrder_Checked(object sender, RoutedEventArgs e)
        {
            txtOrder.IsEnabled = true;
            btnPfOrder.IsEnabled = true;
            txtOrder.Focus();
        }

        //OrderNo
        private void chkOrder_Unchecked(object sender, RoutedEventArgs e)
        {
            txtOrder.IsEnabled = false;
            btnPfOrder.IsEnabled = false;
        }

        //OrderNo
        private void txtOrder_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                MainWindow.pf.ReturnCode(txtOrder, (int)Defind_CodeFind.DCF_ORDER, "");
            }
        }

        //OrderNo
        private void btnPfOrder_Click(object sender, RoutedEventArgs e)
        {
            MainWindow.pf.ReturnCode(txtOrder, (int)Defind_CodeFind.DCF_ORDER, "");
        }

        //거래처
        private void lblCustom_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            if (chkCustom.IsChecked == true) { chkCustom.IsChecked = false; }
            else { chkCustom.IsChecked = true; }
        }

        //거래처
        private void chkCustom_Checked(object sender, RoutedEventArgs e)
        {
            txtCustom.IsEnabled = true;
            btnPfCustom.IsEnabled = true;
            txtCustom.Focus();
        }

        //거래처
        private void chkCustom_Unchecked(object sender, RoutedEventArgs e)
        {
            txtCustom.IsEnabled = false;
            btnPfCustom.IsEnabled = false;
        }

        //거래처
        private void txtCustom_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                //pf.ReturnCode(txtCustom, 0, "");
                MainWindow.pf.ReturnCode(txtCustom, (int)Defind_CodeFind.DCF_CUSTOM, "");
            }
        }

        //거래처
        private void btnPfCustom_Click(object sender, RoutedEventArgs e)
        {
            //pf.ReturnCode(txtCustom, 0, "");
            MainWindow.pf.ReturnCode(txtCustom, (int)Defind_CodeFind.DCF_CUSTOM, "");
        }

        //품명
        private void lblArticle_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            if (chkArticle.IsChecked == true) { chkArticle.IsChecked = false; }
            else { chkArticle.IsChecked = true; }
        }

        //품명
        private void chkArticle_Checked(object sender, RoutedEventArgs e)
        {
            txtArticle.IsEnabled = true;
            btnPfArticle.IsEnabled = true;
            txtArticle.Focus();
        }

        //품명
        private void chkArticle_Unchecked(object sender, RoutedEventArgs e)
        {
            txtArticle.IsEnabled = false;
            btnPfArticle.IsEnabled = false;
        }

        //품명
        private void txtArticle_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                MainWindow.pf.ReturnCode(txtArticle, 77, "");
            }
        }
        //품명
        private void btnPfArticle_Click(object sender, RoutedEventArgs e)
        {
            MainWindow.pf.ReturnCode(txtArticle, 77, "");
        }


        //품번
        private void LabelBuyerArticleNoSearch_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            if (chkBuyerArticleNoSearch.IsChecked == true)
            {
                chkBuyerArticleNoSearch.IsChecked = false;
            }
            else
            {
                chkBuyerArticleNoSearch.IsChecked = true;
            }
        }

        //품번
        private void chkBuyerArticleNoSearch_Checked(object sender, RoutedEventArgs e)
        {
            txtBuyerArticleNoSearch.IsEnabled = true;
            btnpfBuyerArticleNoSearch.IsEnabled = true;
            txtBuyerArticleNoSearch.Focus();
        }

        //품번
        private void chkBuyerArticleNoSearch_Unchecked(object sender, RoutedEventArgs e)
        {
            txtBuyerArticleNoSearch.IsEnabled = false;
            btnpfBuyerArticleNoSearch.IsEnabled = false;
        }

        //품번
        private void txtBuyerArticleNoSearch_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                MainWindow.pf.ReturnCode(txtBuyerArticleNoSearch, 76, txtBuyerArticleNoSearch.Text);
            }
        }

        //품번
        private void btnpfBuyerArticleNoSearch_Click(object sender, RoutedEventArgs e)
        {
            MainWindow.pf.ReturnCode(txtBuyerArticleNoSearch, 76, txtBuyerArticleNoSearch.Text);
        }


        // 작업자
        private void lblWorker_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            if (chkWorker.IsChecked == true) { chkWorker.IsChecked = false; }
            else { chkWorker.IsChecked = true; }
        }
        private void chkWorker_Checked(object sender, RoutedEventArgs e)
        {
            txtWorker.IsEnabled = true;
        }
        private void chkWorker_Unchecked(object sender, RoutedEventArgs e)
        {
            txtWorker.IsEnabled = false;
        }

        private void rbnOrderNo_Click(object sender, RoutedEventArgs e)
        {
            RbnOderCheck();
        }

        private void rbnOrderID_Click(object sender, RoutedEventArgs e)
        {
            RbnOderCheck();
        }

        private void RbnOderCheck()
        {
            if (rbnOrderID.IsChecked == true)
            {
                tbkOrder.Text = "관리 번호";
                dgdtpeOrderID.Visibility = Visibility.Visible;
                dgdtpeOrderNo.Visibility = Visibility.Hidden;
            }
            else if (rbnOrderNo.IsChecked == true)
            {
                tbkOrder.Text = "Order No.";
                dgdtpeOrderID.Visibility = Visibility.Hidden;
                dgdtpeOrderNo.Visibility = Visibility.Visible;
            }
        }

        #endregion

        #region 버튼 클릭 이벤트

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
            FillGrid();

            if (dgdResult.Items.Count == 1)
            {
                MessageBox.Show("조회된 데이터가 없습니다.");
                return;
            }
        }

        //닫기
        private void btnClose_Click(object sender, RoutedEventArgs e)
        {
            DataStore.Instance.InsertLogByFormS(this.GetType().Name, stDate, stTime, "E");
            Lib.Instance.ChildMenuClose(this.ToString());
        }

        private void btnExcel_Click(object sender, RoutedEventArgs e)
        {
            DataTable dt = null;
            string Name = string.Empty;

            string[] dgdStr = new string[2];
            dgdStr[0] = "일생산현황";
            dgdStr[1] = dgdResult.Name;

            ExportExcelxaml ExpExc = new ExportExcelxaml(dgdStr);
            ExpExc.ShowDialog();

            if (ExpExc.DialogResult.HasValue)
            {
                if (ExpExc.choice.Equals(dgdResult.Name))
                {
                    DataStore.Instance.InsertLogByForm(this.GetType().Name, "E");
                    if (ExpExc.Check.Equals("Y"))
                        dt = Lib.Instance.DataGridToDTinHidden(dgdResult);
                    else
                        dt = Lib.Instance.DataGirdToDataTable(dgdResult);

                    Name = dgdResult.Name;
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

        #endregion

        

        #region 주요 메서드 - 조회 FillGrid

        private void FillGrid()
        {
            dgdTotal.Items.Clear();

            if (dgdResult.Items.Count > 0)
            {
                dgdResult.Items.Clear();
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
                sqlParameter.Add("sFromDate", dtpSDate.SelectedDate.Value.ToString("yyyyMMdd"));
                sqlParameter.Add("sToDate", dtpEDate.SelectedDate.Value.ToString("yyyyMMdd"));
                sqlParameter.Add("sProcessID", ProcessID);
                sqlParameter.Add("sMachineID", MachineID);
                sqlParameter.Add("ArticleID", chkArticle.IsChecked == true && txtArticle.Tag != null ? txtArticle.Tag.ToString() : "");

                sqlParameter.Add("CustomID", ((chkCustom.IsChecked == true) ? 
                    (txtCustom.Tag == null ? "" : txtCustom.Tag.ToString()) : ""));
                sqlParameter.Add("nOrderID", ((chkOrder.IsChecked == true) ? (tbkOrder.Text.Equals("관리번호") ? 1 : 2) : 0));
                sqlParameter.Add("sOrderID", ((chkOrder.IsChecked == true) ? txtOrder.Text : ""));
                sqlParameter.Add("nJobGbn", chkGubun.IsChecked == true ? 1: 0);
                sqlParameter.Add("sJobGubun", chkGubun.IsChecked == true && cboGubun.SelectedValue != null && !cboGubun.SelectedValue.ToString().Trim().Equals("0") ? cboGubun.SelectedValue.ToString() : "");

                sqlParameter.Add("nBuyerModel", 0);
                sqlParameter.Add("sBuyerModel", "");
                sqlParameter.Add("nBuyerArticleNo", chkBuyerArticleNoSearch.IsChecked == true ? 1 : 0);
                sqlParameter.Add("sBuyerArticleNo", chkBuyerArticleNoSearch.IsChecked == true ? (txtBuyerArticleNoSearch.Tag != null ? txtBuyerArticleNoSearch.Tag.ToString() : "") : "");

                sqlParameter.Add("nWorkerName", chkWorker.IsChecked == true ? 1 : 0);
                sqlParameter.Add("sWorkerName", chkWorker.IsChecked == true && txtWorker.Text.Trim().Equals("") == false ? txtWorker.Text : "");
                sqlParameter.Add("ndefect", chkDefect.IsChecked == true ? 1 : 0);


                DataSet ds = DataStore.Instance.ProcedureToDataSet_LogWrite("xp_prd_sWKResult_WPF_Q", sqlParameter, true, "R");

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
                            var WinR = new Win_Prd_ProcessResult_Q_CodeView()
                            {
                                Num = i,
                                cls = dr["cls"].ToString().Trim(),

                                WorkDate = dr["WorkDate"].ToString(),
                                WorkDate_CV = DatePickerFormat(dr["WorkDate"].ToString()),
                                ProcessID = dr["ProcessID"].ToString(),
                                Process = dr["Process"].ToString(),
                                OrderID = dr["OrderID"].ToString(),

                                OrderNo = dr["OrderNo"].ToString(),
                                AcptDate = dr["AcptDate"].ToString(),
                                AcptDate_CV = DatePickerFormat(dr["AcptDate"].ToString()),
                                OrderQty = stringFormatN0(dr["OrderQty"]),
                                MachineID = dr["MachineID"].ToString(),

                                InstDate = dr["InstDate"].ToString(),
                                InstDate_CV = DatePickerFormat(dr["InstDate"].ToString()),
                                InstQty = stringFormatN0(dr["InstQty"]),
                                WorkQty = stringFormatN0(dr["WorkQty"]),
                                WorkPersonID = dr["WorkPersonID"].ToString(),

                                BuyerModelID = dr["BuyerModelID"].ToString(),
                                BuyerModel = dr["BuyerModel"].ToString(),
                                CustomID = dr["CustomID"].ToString(),

                                KCustom = dr["KCustom"].ToString(),
                                Worker = dr["Worker"].ToString(),
                                Article = dr["Article"].ToString(),
                                LabelID = dr["LabelID"].ToString(),
                                JobGbn = dr["JobGbn"].ToString(),
                                JobGbnname = dr["JobGbnname"].ToString(),

                                WorkStartDate = dr["WorkStartDate"].ToString(),
                                WorkStartDate_CV = DatePickerFormat(dr["WorkStartDate"].ToString()),
                                WorkStartTime = dr["WorkStartTime"].ToString(),
                                WorkStartTime_CV = ConvertTimeFormat(dr["WorkStartTime"].ToString()),
                                WorkEndDate = dr["WorkEndDate"].ToString(),

                                WorkEndDate_CV = DatePickerFormat(dr["WorkEndDate"].ToString()),
                                WorkEndTime = dr["WorkEndTime"].ToString(),
                                WorkEndTime_CV = ConvertTimeFormat(dr["WorkEndTime"].ToString()),
                                WorkHour = dr["WorkHour"].ToString(),
                                WorkMinute = dr["WorkMinute"].ToString(),

                                JobID = dr["JobID"].ToString(),
                                Articleid = dr["Articleid"].ToString(),
                                WorkCnt = dr["WorkCnt"].ToString(),
                                NoReworkCode = dr["NoReworkCode"].ToString(),
                                NoReworkName = dr["NoReworkName"].ToString(),

                                FourMID = dr["4MID"].ToString(),
                                FourMSubject = dr["4MSubject"].ToString(),
                                BuyerArticleNo = dr["BuyerArticleNo"].ToString(),

                                ScanDate = dr["ScanDate"].ToString(),
                                ScanDate_CV = ConvertTimeFormat(dr["ScanDate"].ToString()),

                                MachineNo = dr["MachineNo"].ToString(),
                            };

                            if (WinR.cls.Equals("1"))
                            {
                                WinR.Time = StartTimeAndEndTime(WinR.WorkStartDate, WinR.WorkStartTime,
                                WinR.WorkEndDate, WinR.WorkEndTime);
                            }
                            else if (WinR.cls.Equals("2")) // 공정계
                            {
                                WinR.WorkDate_CV = "";
                                WinR.Process = "공정계";

                                WinR.OrderQty = "";
                            }
                            else if (WinR.cls.Equals("3")) // 일계
                            {
                                WinR.WorkDate_CV = "일계";

                                WinR.OrderQty = "";
                            }
                            else if (WinR.cls.Equals("4")) // 작업구분계
                            {
                                WinR.WorkDate_CV = "작업구분계";
                                WinR.Process = "";

                                WinR.OrderQty = "";
                            }
                            else if (WinR.cls.Equals("9")) // 총계
                            {
                                WinR.WorkDate_CV = "총계";
                                WinR.Process = "";

                                WinR.OrderQty = "";

                                dgdTotal.Items.Add(WinR);
                            }

                            dgdResult.Items.Add(WinR);
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

        #endregion // 주요 메서드 - 조회 FillGrid

        private string StartTimeAndEndTime(string SDate, string STime, string EDate, string ETime)
        {
            string STandET = string.Empty;
            
            STandET += STime.Substring(0, 2) + ":" + STime.Substring(2, 2) + " ~ ";
            STandET += ETime.Substring(0, 2) + ":" + ETime.Substring(2, 2);

            return STandET;
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

        private string stringFormatN2(object obj)
        {
            return string.Format("{0:N2}", obj);
        }

        private string stringFormatNN(object obj, int length)
        {
            return string.Format("{0:N" + length + "}", obj);
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

        // 시간 : 분 으로 변환
        private string ConvertTimeFormat(string str)
        {
            string result = "";

            str = str.Trim().Replace(":", "");
            if (str.Length > 3 && str.Length < 7)
            {
                string hour = str.Substring(0, 2);
                string min = str.Substring(2, 2);

                result = hour + ":" + min;
            }

            return result;
        }

        #endregion

        // 검색조건 - 텍스트 박스 엔터 → 조회
        private void txtBox_EnterAndSearch(object sender, System.Windows.Input.KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                //rowNum = 0;
                using (Loading lw = new Loading(FillGrid))
                {
                    lw.ShowDialog();
                }
            }
        }

        private void dgdResult_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            try
            {
                var ProcessResultInfo = dgdResult.SelectedItem as Win_Prd_ProcessResult_Q_CodeView;

                if (ProcessResultInfo != null)
                {
                    FillGrid_Defect(ProcessResultInfo.JobID);
                }

            }
            catch (Exception ee)
            {
                MessageBox.Show("오류지점 - " + ee.ToString());
            }
        }

        #region 조회 Defect
        private void FillGrid_Defect(string JobID)
        {
            DataGridDefect.Items.Clear();

            try
            {
                Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                sqlParameter.Clear();
                sqlParameter.Add("JobID", JobID);

                DataSet ds = DataStore.Instance.ProcedureToDataSet("xp_prd_sWKResultDaily_Defect", sqlParameter, false);

                if (ds != null && ds.Tables.Count > 0)
                {
                    DataTable dt = ds.Tables[0];

                    if (dt.Rows.Count == 0)
                    {
                        //MessageBox.Show("조회된 데이터가 없습니다.");
                    }
                    else
                    {
                        DataRowCollection drc = dt.Rows;
                        int i = 0;

                        foreach (DataRow dr in drc)
                        {
                            i++;
                            var WPPRQDC = new Win_Prd_ProcessResult_Q_Defect_CodeView()
                            {
                                Num = i,
                                DefectID = dr["DefectID"].ToString(),
                                KDefect = dr["KDefect"].ToString(),
                                DefectQty = dr["DefectQty"].ToString(),
                            };

                            DataGridDefect.Items.Add(WPPRQDC);

                        }
                    }
                }

            }
            catch (Exception ee)
            {
                MessageBox.Show("오류지점 - " + ee.ToString());
            }
            finally
            {
                DataStore.Instance.CloseConnection();
            }
        }






        #endregion


    }

    class Win_Prd_ProcessResult_Q_CodeView : BaseView
    {
        public override string ToString()
        {
            return (this.ReportAllProperties());
        }

        public string cls { get; set; }
        public string WorkDate { get; set; }
        public string WorkDate_CV { get; set; }
        public string ProcessID { get; set; }
        public string Process { get; set; }

        public string OrderID { get; set; }
        public string OrderNo { get; set; }
        public string AcptDate { get; set; }
        public string AcptDate_CV { get; set; }
        public string OrderQty { get; set; }

        public string MachineID { get; set; }
        public string InstDate { get; set; }
        public string InstDate_CV { get; set; }
        public string InstQty { get; set; }
        public string WorkQty { get; set; }

        public string WorkPersonID { get; set; }
        public string ScanTime { get; set; }
        public string ScanTime_CV { get; set; }
        public string BuyerModelID { get; set; }
        public string BuyerModel { get; set; }

        public string CustomID { get; set; }
        public string KCustom { get; set; }
        public string Worker { get; set; }
        public string Article { get; set; }       
        public string LabelID { get; set; }

        public string JobGbn { get; set; }
        public string JobGbnname { get; set; }
        public string WorkStartDate { get; set; }
        public string WorkStartDate_CV { get; set; }
        public string WorkStartTime { get; set; }

        public string WorkStartTime_CV { get; set; }
        public string WorkEndDate { get; set; }
        public string WorkEndDate_CV { get; set; }
        public string WorkEndTime { get; set; }
        public string WorkEndTime_CV { get; set; }

        public string WorkHour { get; set; }
        public string WorkMinute { get; set; }
        public string JobID { get; set; }
        public string Articleid { get; set; }
        public string WorkCnt { get; set; }

        public string NoReworkCode { get; set; }
        public string NoReworkName { get; set; }        
        public string FourMID { get; set; }
        public string FourMSubject { get; set; }
        public int Num { get; set; }

        public string Time { get; set; }
        public string BuyerArticleNo { get; set; }

        public string ScanDate { get; set; }
        public string ScanDate_CV { get; set; }
        public string MachineNo { get; set; }
    }

    class ProcessList
    {
        public string Process { get; set; }
        public string ProcessID { get; set; }
        public string ScanDate { get; set; }
    }

    class Win_Prd_ProcessResult_Q_Defect_CodeView : BaseView
    {
        public override string ToString()
        {
            return (this.ReportAllProperties());
        }

        public int Num { get; set; }

        public string DefectID { get; set; }
        public string KDefect { get; set; }
        public string DefectQty { get; set; }
    }

    //class Win_Prd_ProcessResult_Q_Process : BaseView
    //{

    //}

    //class Win_Prd_ProcessResult_Q_Child : BaseView
    //{

    //}
}

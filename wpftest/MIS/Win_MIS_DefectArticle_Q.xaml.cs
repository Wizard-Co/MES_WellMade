using LiveCharts;
using LiveCharts.Wpf;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Data;
using System.Threading;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using WizMes_WellMade.PopUP;
using WPF.MDI;

namespace WizMes_WellMade
{
    /// <summary>
    /// Win_MIS_DefectArticle_Q.xaml에 대한 상호 작용 논리
    /// </summary>
    public partial class Win_MIS_DefectArticle_Q : UserControl
    {
        #region 변수 선언 및 로드

        WizMes_WellMade.PopUp.NoticeMessage msg = new PopUp.NoticeMessage();
        Lib lib = new Lib();

        public Win_MIS_DefectArticle_Q()
        {
            InitializeComponent();
        }

        private void UserControl_Loaded(object sender, RoutedEventArgs e)
        {
            lib.UiLoading(sender);
            dtpSDate.SelectedDate = DateTime.Today;
            dtpEDate.SelectedDate = DateTime.Today;

            SetComboBox();

            cboProcess.SelectedIndex = 0;
            cboMachine.SelectedIndex = 0;
            cboGubun.SelectedIndex = 0;
        }

        #endregion

        #region 날짜 관련 이벤트

        //검색기간
        private void lblDate_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            if (chkDate.IsChecked == true) { chkDate.IsChecked = false; }
            else { chkDate.IsChecked = true; }
        }

        //검색기간
        private void chkDate_Checked(object sender, RoutedEventArgs e)
        {
            if (dtpSDate != null && dtpEDate != null)
            {
                dtpSDate.IsEnabled = true;
                dtpEDate.IsEnabled = true;
            }
        }

        //검색기간
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

        #region 우측 상단 버튼 클릭

        //조회
        private void btnSearch_Click(object sender, RoutedEventArgs e)
        {
            //검색버튼 비활성화
            btnSearch.IsEnabled = false;

            Dispatcher.BeginInvoke(new Action(() =>

            {
                Thread.Sleep(2000);

                //로직
                FillGrid();

                if (dgdmain.Items.Count > 0)
                {
                    dgdmain.SelectedIndex = 0;
                    this.DataContext = dgdmain.SelectedItem as Win_MIS_DefectArticle_Q_CodeView;
                }
                else
                {
                    MessageBox.Show("조회된 데이터가 없습니다.");
                }

            }), System.Windows.Threading.DispatcherPriority.Background);



            Dispatcher.BeginInvoke(new Action(() =>

            {
                btnSearch.IsEnabled = true;

            }), System.Windows.Threading.DispatcherPriority.Background);


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
        }

        //인쇄
        private void btnPrint_Click(object sender, RoutedEventArgs e)
        {
            ContextMenu menu = btnPrint.ContextMenu;
            menu.StaysOpen = true;
            menu.IsOpen = true;
        }

        //엑셀
        private void btnExcel_Click(object sender, RoutedEventArgs e)
        {
            DataTable dt = null;
            string Name = string.Empty;

            string[] lst = new string[2];
            lst[0] = "공정 품질불량 조회";
            lst[1] = dgdmain.Name;

            ExportExcelxaml ExpExc = new ExportExcelxaml(lst);

            ExpExc.ShowDialog();

            if (ExpExc.DialogResult.HasValue)
            {
                if (ExpExc.choice.Equals(dgdmain.Name))
                {
                    if (ExpExc.Check.Equals("Y"))
                        dt = lib.DataGridToDTinHidden(dgdmain);
                    else
                        dt = lib.DataGirdToDataTable(dgdmain);

                    Name = dgdmain.Name;

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

        #region 콤보박스 세팅

        private void SetComboBox()
        {
            List<string> strCombo1 = new List<string>();
            strCombo1.Add("전체");
            strCombo1.Add("입고");
            strCombo1.Add("생산");
            strCombo1.Add("자주검사");
            strCombo1.Add("출하");
            ObservableCollection<CodeView> ovcGugunSearch = ComboBoxUtil.Instance.Direct_SetComboBox(strCombo1);
            this.cboGubun.ItemsSource = ovcGugunSearch;
            this.cboGubun.DisplayMemberPath = "code_name";
            this.cboGubun.SelectedValuePath = "code_id";

            this.cboGubun.ItemsSource = ovcGugunSearch;
            this.cboGubun.DisplayMemberPath = "code_name";
            this.cboGubun.SelectedValuePath = "code_id";

            // 공정
            ObservableCollection<CodeView> ovcProcess = ComboBoxUtil.Instance.GetWorkProcess(0, "");
            this.cboProcess.ItemsSource = ovcProcess;
            this.cboProcess.DisplayMemberPath = "code_name";
            this.cboProcess.SelectedValuePath = "code_id";

            this.cboProcess.ItemsSource = ovcProcess;
            this.cboProcess.DisplayMemberPath = "code_name";
            this.cboProcess.SelectedValuePath = "code_id";
        }
        #endregion

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

        #region 검색조건

        //품번 플러스 파인더 
        private void BtnPfArticleNo_Click(object sender, RoutedEventArgs e)
        {
            MainWindow.pf.ReturnCode(txtBuyerArticleNo, 84, txtBuyerArticleNo.Text);
        }

        //품번 키다운
        private void TxtBuyerArticleNo_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                MainWindow.pf.ReturnCode(txtBuyerArticleNo, 84, txtBuyerArticleNo.Text);
            }
        }

        //품번
        private void lblBuyerArticleNo_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            if (chkBuyerArticleNo.IsChecked == true) { chkBuyerArticleNo.IsChecked = false; }
            else { chkBuyerArticleNo.IsChecked = true; }
        }

        //품번
        private void chkBuyerArticleNo_Checked(object sender, RoutedEventArgs e)
        {
            txtBuyerArticleNo.IsEnabled = true;
            btnPfArticleNo.IsEnabled = true;
        }

        //품번
        private void chkBuyerArticleNo_Unchecked(object sender, RoutedEventArgs e)
        {
            txtBuyerArticleNo.IsEnabled = false;
            btnPfArticleNo.IsEnabled = false;
        }

        //공정 콤보박스 클릭
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
        #endregion

        //실 조회
        private void FillGrid()
        {
            try
            {
                if (dgdmain.Items.Count > 0)
                {
                    dgdmain.Items.Clear();
                }

                if (lvcChart.Series != null && lvcChart.Series.Count > 0)
                {
                    lvcChart.Series.Clear();
                }

                dgdmain2.Items.Clear();

                bool chkProcess_I = false;
                bool chkMachine_I = false;
                bool chkGbn_I = false;

                if (!cboProcess.SelectedValue.Equals(""))
                {
                    chkProcess_I = true;
                }
                else { chkProcess_I = false; }

                if (!cboMachine.SelectedValue.Equals(""))
                    chkMachine_I = true;
                else chkMachine_I = false;

                if (!cboGubun.SelectedValue.Equals(""))
                    chkGbn_I = true;
                else chkGbn_I = false;


                var DDD = new Win_MIS_DDD();


                Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                sqlParameter.Add("FromDate", chkDate.IsChecked == true ? dtpSDate.SelectedDate.Value.ToString("yyyyMMdd") : "");
                sqlParameter.Add("ToDate", chkDate.IsChecked == true ? dtpEDate.SelectedDate.Value.ToString("yyyyMMdd") : "");
                //sqlParameter.Add("EndDate", chkDate.IsChecked == true ? dtpEDate.SelectedDate.Value.ToString("yyyyMMdd") : "");
                sqlParameter.Add("nchkProcessID", chkProcess_I == true ? 1 : 0);
                sqlParameter.Add("ProcessID", chkProcess_I == true ? (cboProcess.SelectedValue != null ? cboProcess.SelectedValue.ToString() : "") : "");
                sqlParameter.Add("nchkMachineID", chkMachine_I == true ? 1 : 0);
                sqlParameter.Add("MachineID", chkMachine_I == true ? (cboMachine.SelectedValue != null ? cboMachine.SelectedValue.ToString() : "") : "");
                sqlParameter.Add("nchkDefectStep", chkGbn_I == true ? 1 : 0);
                sqlParameter.Add("DefectStep", chkGbn_I == true ? (cboGubun.SelectedValue != null ? cboGubun.SelectedValue.ToString() : "") : "");
                sqlParameter.Add("nchkBuyerArticleNo", chkBuyerArticleNo.IsChecked == true ? 1 : 0);
                sqlParameter.Add("BuyerArticleNo", chkBuyerArticleNo.IsChecked == true && !txtBuyerArticleNo.Text.Trim().Equals("") ? txtBuyerArticleNo.Text : "");


                DataSet ds = DataStore.Instance.ProcedureToDataSet("xp_Mis_sQualSymptom", sqlParameter, false);

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
                            var DD = new Win_MIS_DefectArticle_Q_CodeView()
                            {
                                Rownum = dr["Rownum"].ToString(),
                                InspectDate = dr["InspectDate"].ToString(),
                                StepName = dr["StepName"].ToString(),
                                Process = dr["Process"].ToString(),
                                //ProcessID = dr["ProcessID"].ToString(),
                                MachineNo = dr["MachineNo"].ToString(),
                                //MachineID = dr["MachineID"].ToString(),
                                //BuyerArticleNo = Convert.ToDouble(dr["BuyerArticleNo"]),
                                BuyerArticleNo = dr["BuyerArticleNo"].ToString(),
                                Defect1 = stringFormatN0AndZeroEmpty(dr["Defect1"].ToString()),
                                Defect2 = stringFormatN0AndZeroEmpty(dr["Defect2"].ToString()),
                                Defect3 = stringFormatN0AndZeroEmpty(dr["Defect3"].ToString()),
                                Defect4 = stringFormatN0AndZeroEmpty(dr["Defect4"].ToString()),
                                Defect5 = stringFormatN0AndZeroEmpty(dr["Defect5"].ToString()),
                                Defect6 = stringFormatN0AndZeroEmpty(dr["Defect6"].ToString()),
                                Defect7 = stringFormatN0AndZeroEmpty(dr["Defect7"].ToString()),
                                Defect8 = stringFormatN0AndZeroEmpty(dr["Defect8"].ToString()),
                                Defect9 = stringFormatN0AndZeroEmpty(dr["Defect9"].ToString()),
                                Defect10 = stringFormatN0AndZeroEmpty(dr["Defect10"].ToString()),
                                DfectHap = stringFormatN0AndZeroEmpty(dr["DfectHap"].ToString()),

                            };
                            if (i == 1)
                            {
                                dgdmain.Columns[0].Header = Convert.ToString(dr["InspectDate"]);
                                dgdmain.Columns[1].Header = Convert.ToString(dr["StepName"]);
                                dgdmain.Columns[2].Header = Convert.ToString(dr["Process"]);
                                dgdmain.Columns[3].Header = Convert.ToString(dr["Machineno"]);
                                dgdmain.Columns[4].Header = Convert.ToString(dr["BuyerArticleNo"]);
                                dgdmain.Columns[5].Header = Convert.ToString(dr["Defect1"]);
                                dgdmain.Columns[6].Header = Convert.ToString(dr["Defect2"]);
                                dgdmain.Columns[7].Header = Convert.ToString(dr["Defect3"]);
                                dgdmain.Columns[8].Header = Convert.ToString(dr["Defect4"]);
                                dgdmain.Columns[9].Header = Convert.ToString(dr["Defect5"]);
                                dgdmain.Columns[10].Header = Convert.ToString(dr["Defect6"]);
                                dgdmain.Columns[11].Header = Convert.ToString(dr["Defect7"]);
                                dgdmain.Columns[12].Header = Convert.ToString(dr["Defect8"]);
                                dgdmain.Columns[13].Header = Convert.ToString(dr["Defect9"]);
                                dgdmain.Columns[14].Header = Convert.ToString(dr["Defect10"]);
                                dgdmain.Columns[15].Header = Convert.ToString(dr["DfectHap"]);

                                dgdmain2.Columns[1].Header = Convert.ToString(dr["Defect1"]);
                                dgdmain2.Columns[2].Header = Convert.ToString(dr["Defect2"]);
                                dgdmain2.Columns[3].Header = Convert.ToString(dr["Defect3"]);
                                dgdmain2.Columns[4].Header = Convert.ToString(dr["Defect4"]);
                                dgdmain2.Columns[5].Header = Convert.ToString(dr["Defect5"]);
                                dgdmain2.Columns[6].Header = Convert.ToString(dr["Defect6"]);
                                dgdmain2.Columns[7].Header = Convert.ToString(dr["Defect7"]);
                                dgdmain2.Columns[8].Header = Convert.ToString(dr["Defect8"]);
                                dgdmain2.Columns[9].Header = Convert.ToString(dr["Defect9"]);
                                dgdmain2.Columns[10].Header = Convert.ToString(dr["Defect10"]);
                            }
                            else
                            {
                                if (dr["Rownum"].ToString().Trim().Equals("99999"))
                                {
                                    DDD.text = "총계";
                                    DDD.Defect1 = Convert.ToDouble(dr["Defect1"]);
                                    DDD.Defect2 = Convert.ToDouble(dr["Defect2"]);
                                    DDD.Defect3 = Convert.ToDouble(dr["Defect3"]);
                                    DDD.Defect4 = Convert.ToDouble(dr["Defect4"]);
                                    DDD.Defect5 = Convert.ToDouble(dr["Defect5"]);
                                    DDD.Defect6 = Convert.ToDouble(dr["Defect6"]);
                                    DDD.Defect7 = Convert.ToDouble(dr["Defect7"]);
                                    DDD.Defect8 = Convert.ToDouble(dr["Defect8"]);
                                    DDD.Defect9 = Convert.ToDouble(dr["Defect9"]);
                                    DDD.Defect10 = Convert.ToDouble(dr["Defect10"]);
                                    DDD.DfectHap = Convert.ToDouble(dr["DfectHap"]);
                                }
                                //else
                                //{
                                //    //MessageBox.Show("검색 자료가 없습니다.");
                                //    return;
                                //}
                                double a1 = ConvertDouble(dr["Defect1"].ToString());
                                double a2 = ConvertDouble(dr["Defect2"].ToString());
                                double a3 = ConvertDouble(dr["Defect3"].ToString());
                                double a4 = ConvertDouble(dr["Defect4"].ToString());
                                double a5 = ConvertDouble(dr["Defect5"].ToString());
                                double a6 = ConvertDouble(dr["Defect6"].ToString());
                                double a7 = ConvertDouble(dr["Defect7"].ToString());
                                double a8 = ConvertDouble(dr["Defect8"].ToString());
                                double a9 = ConvertDouble(dr["Defect9"].ToString());
                                double a10 = ConvertDouble(dr["Defect10"].ToString());

                                SeriesCollection SeriesCollection = new SeriesCollection
                            {
                                new ColumnSeries
                                {
                                    Title = "유형별 불량",
                                    Values = new ChartValues<double>
                                    {
                                        a1,a2,a3,a4,a5,
                                        a6,a7,a8,a9,a10
                                    }
                                }
                            };


                                lvcChart.Series = SeriesCollection;

                                dgdmain.Items.Add(DD);

                                if (DD.Rownum.Equals("99999"))
                                    dgdmain.Items.Remove(DD);
                            }

                        }
                        dgdmain2.Items.Add(DDD);


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


        #region 히든

        //미리보기(인쇄 하위버튼)
        private void menuSeeAhead_Click(object sender, RoutedEventArgs e)
        {
            if (dgdmain.Items.Count < 1)
            {
                MessageBox.Show("먼저 검색해 주세요.");
                return;
            }

            msg.Show();
            msg.Topmost = true;
            msg.Refresh();

            PrintWork(true);
        }

        //바로인쇄(인쇄 하위버튼)
        private void menuRightPrint_Click(object sender, RoutedEventArgs e)
        {
            if (dgdmain.Items.Count < 1)
            {
                MessageBox.Show("먼저 검색해 주세요.");
                return;
            }

            msg.Show();
            msg.Topmost = true;
            msg.Refresh();

            PrintWork(false);
        }

        //닫   기(인쇄 하위버튼)
        private void menuClose_Click(object sender, RoutedEventArgs e)
        {
            ContextMenu menu = btnPrint.ContextMenu;
            menu.StaysOpen = false;
            menu.IsOpen = false;
        }

        private void PrintWork(bool preview_click)
        {

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

        // 0이면 빈칸, 아니면, 천자리 콤마, 소수점 버리기
        private string stringFormatN0AndZeroEmpty(object obj)
        {
            string result = string.Format("{0:N0}", obj);

            if (result.Trim().Equals("0"))
            {
                result = "";
            }

            return result;
        }

        // 천마리 콤마, 소수점 버리기
        private string stringFormatN0(object obj)
        {
            return string.Format("{0:N0}", obj);
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

    class Win_MIS_DefectArticle_Q_CodeView : BaseView
    {
        public override string ToString()
        {
            return (this.ReportAllProperties());
        }

        public string Rownum { get; set; }
        public string InspectDate { get; set; }
        public string StepName { get; set; }
        public string ProcessID { get; set; }
        public string Process { get; set; }
        public string MachineID { get; set; }
        public string MachineNo { get; set; }
        public string BuyerArticleNo { get; set; }
        public string DefectSymtom { get; set; }
        public string defectQty { get; set; }
        public string Defect1 { get; set; }
        public string Defect2 { get; set; }
        public string Defect3 { get; set; }
        public string Defect4 { get; set; }
        public string Defect5 { get; set; }
        public string Defect6 { get; set; }
        public string Defect7 { get; set; }
        public string Defect8 { get; set; }
        public string Defect9 { get; set; }
        public string Defect10 { get; set; }
        public string DfectHap { get; set; }
    }

    class Win_MIS_DDD : BaseView
    {
        public string InspectDate { get; set; }
        public string StepName { get; set; }
        public string ProcessID { get; set; }
        public string Process { get; set; }
        public string MachineID { get; set; }
        public string MachineNo { get; set; }
        public string BuyerArticleNo { get; set; }
        public string DefectSymtom { get; set; }
        public string defectQty { get; set; }
        public double Defect1 { get; set; }
        public double Defect2 { get; set; }
        public double Defect3 { get; set; }
        public double Defect4 { get; set; }
        public double Defect5 { get; set; }
        public double Defect6 { get; set; }
        public double Defect7 { get; set; }
        public double Defect8 { get; set; }
        public double Defect9 { get; set; }
        public double Defect10 { get; set; }
        public double DfectHap { get; set; }
        public string text { get; set; }
    }
}

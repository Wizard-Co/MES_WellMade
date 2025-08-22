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
    /// Win_Prd_ProcessResultSum_Q.xaml에 대한 상호 작용 논리
    /// </summary>
    public partial class Win_Prd_ProcessResultSum_Q : UserControl
    {
        string stDate = string.Empty;
        string stTime = string.Empty;

        Lib lib = new Lib();
        public Win_Prd_ProcessResultSum_Q()
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

            SetComboBox();

            cboProcess.SelectedIndex = 0;
            cboMachine.SelectedIndex = 0;

            rbnOrderID.IsChecked = true;
        }

        private void SetComboBox()
        {
            ObservableCollection<CodeView> cbWork = ComboBoxUtil.Instance.GetWorkProcess(0, "");

            this.cboProcess.ItemsSource = cbWork;
            this.cboProcess.DisplayMemberPath = "code_name";
            this.cboProcess.SelectedValuePath = "code_id";

            ObservableCollection<CodeView> ovcMachine = GetMachineByProcessID("");
            this.cboMachine.ItemsSource = ovcMachine;
            this.cboMachine.DisplayMemberPath = "code_name";
            this.cboMachine.SelectedValuePath = "code_id";
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

            DataSet ds = DataStore.Instance.ProcedureToDataSet("xp_prd_sMachineForComboBox", sqlParameter, false);

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

        #region 날짜 관련 이벤트

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

        #region 체크 등 이벤트

        //작업자
        private void lblPerson_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            if (chkPerson.IsChecked == true) { chkPerson.IsChecked = false; }
            else { chkPerson.IsChecked = true; }
        }

        //작업자
        private void chkPerson_Checked(object sender, RoutedEventArgs e)
        {
            txtPerson.IsEnabled = true;
            txtPerson.Focus();
        }

        //작업자
        private void chkPerson_Unchecked(object sender, RoutedEventArgs e)
        {
            txtPerson.IsEnabled = false;
        }

        //관리번호
        private void rbnOrderNo_Click(object sender, RoutedEventArgs e)
        {
            if (rbnOrderNo.IsChecked == true)
            {
                tbkOrder.Text = " Order No.";
            }
        }

        //관리번호
        private void rbnOrderID_Click(object sender, RoutedEventArgs e)
        {
            if (rbnOrderID.IsChecked == true)
            {
                tbkOrder.Text = " 관리번호";
            }
        }

        //관리번호
        private void lblOrder_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            if (chkOrder.IsChecked == true) { chkOrder.IsChecked = false; }
            else { chkOrder.IsChecked = true; }
        }

        //관리번호
        private void chkOrder_Checked(object sender, RoutedEventArgs e)
        {
            txtOrder.IsEnabled = true;
            btnPfOrder.IsEnabled = true;
            txtOrder.Focus();
        }

        //관리번호
        private void chkOrder_Unchecked(object sender, RoutedEventArgs e)
        {
            txtOrder.IsEnabled = false;
            btnPfOrder.IsEnabled = false;
        }

        //관리번호
        private void txtOrder_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                MainWindow.pf.ReturnCode(txtOrder, (int)Defind_CodeFind.DCF_ORDER, "");
            }
        }

        //관리번호
        private void btnPfOrder_Click(object sender, RoutedEventArgs e)
        {
            MainWindow.pf.ReturnCode(txtOrder, (int)Defind_CodeFind.DCF_ORDER, "");
        }

        //거래처
        private void lblCustom_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
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
        private void lblArticle_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
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

        #endregion

        //조회
        private void btnSearch_Click(object sender, RoutedEventArgs e)
        {
            DataStore.Instance.InsertLogByForm(this.GetType().Name, "R");
            //검색버튼 비활성화
            btnSearch.IsEnabled = false;

            Dispatcher.BeginInvoke(new Action(() =>

            {
                //로직
                using (Loading lw = new Loading(re_Search))
                {
                    lw.ShowDialog();
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
            DataStore.Instance.InsertLogByFormS(this.GetType().Name, stDate, stTime, "E");
            Lib.Instance.ChildMenuClose(this.ToString());
        }

        #region 엑셀 버튼 이벤트
        private void btnExcel_Click(object sender, RoutedEventArgs e)
        {
            DataTable dt = null;
            string Name = string.Empty;

            TabItem nowTab = tabconGrid.SelectedItem as TabItem;

            if (nowTab.Header.ToString().Equals("공정별 호기별 집계"))
            {
                string[] lst = new string[2];
                lst[0] = "공정별 호기별 집계";
                lst[1] = dgdByProcess.Name;

                ExportExcelxaml ExpExc = new ExportExcelxaml(lst);

                ExpExc.ShowDialog();

                if (ExpExc.DialogResult.HasValue)
                {
                    if (ExpExc.choice.Equals(dgdByProcess.Name))
                    {
                        DataStore.Instance.InsertLogByForm(this.GetType().Name, "E");
                        if (ExpExc.Check.Equals("Y"))
                            dt = Lib.Instance.DataGridToDTinHidden(dgdByProcess);
                        else
                            dt = Lib.Instance.DataGirdToDataTable(dgdByProcess);

                        Name = dgdByProcess.Name;

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
            else if (nowTab.Header.ToString().Equals("품번별 집계"))
            {
                string[] lst = new string[2];
                lst[0] = "품번별 집계";
                lst[1] = dgdByArticle.Name;

                ExportExcelxaml ExpExc = new ExportExcelxaml(lst);

                ExpExc.ShowDialog();

                if (ExpExc.DialogResult.HasValue)
                {
                    if (ExpExc.choice.Equals(dgdByArticle.Name))
                    {
                        if (ExpExc.Check.Equals("Y"))
                            dt = Lib.Instance.DataGridToDTinHidden(dgdByArticle);
                        else
                            dt = Lib.Instance.DataGirdToDataTable(dgdByArticle);

                        Name = dgdByArticle.Name;

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
            else if (nowTab.Header.ToString().Equals("작업자별 집계"))
            {
                string[] lst = new string[2];
                lst[0] = "작업자별 집계";
                lst[1] = dgdByWorker.Name;

                ExportExcelxaml ExpExc = new ExportExcelxaml(lst);

                ExpExc.ShowDialog();

                if (ExpExc.DialogResult.HasValue)
                {
                    if (ExpExc.choice.Equals(dgdByWorker.Name))
                    {
                        if (ExpExc.Check.Equals("Y"))
                            dt = Lib.Instance.DataGridToDTinHidden(dgdByWorker);
                        else
                            dt = Lib.Instance.DataGirdToDataTable(dgdByWorker);

                        Name = dgdByWorker.Name;

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
            else if (nowTab.Header.ToString().Equals("일별 집계"))
            {
                string[] lst = new string[2];
                lst[0] = "일별 집계";
                lst[1] = DataGridThisMonth.Name;

                ExportExcelxaml ExpExc = new ExportExcelxaml(lst);

                ExpExc.ShowDialog();

                if (ExpExc.DialogResult.HasValue)
                {
                    if (ExpExc.choice.Equals(DataGridThisMonth.Name))
                    {
                        if (ExpExc.Check.Equals("Y"))
                            dt = Lib.Instance.DataGridToDTinHidden(DataGridThisMonth);
                        else
                            dt = Lib.Instance.DataGirdToDataTable(DataGridThisMonth);

                        Name = DataGridThisMonth.Name;

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
        }
        #endregion // 엑셀 버튼 이벤트

        private bool CheckData()
        {
            bool flag = true;

            if (cboProcess.SelectedValue == null)
            {
                MessageBox.Show("공정이 선택되지 않았습니다. 선택해주세요");
                flag = false;
                return flag;
            }

            if (cboMachine.SelectedValue == null)
            {
                MessageBox.Show("호기가 선택되지 않았습니다. 선택해주세요");
                flag = false;
                return flag;
            }

            return flag;
        }

        private void re_Search()
        {
            if (CheckData())
            {
                TabItem nowTab = tabconGrid.SelectedItem as TabItem;

                if (nowTab.Header.ToString().Equals("공정별 호기별 집계"))
                {
                    FillGridProcessMachine();

                    if (dgdByProcess.Items.Count <= 0)
                    {
                        MessageBox.Show("조회된 데이터가 없습니다.");
                        return;
                    }
                }
                else if (nowTab.Header.ToString().Equals("품번별 집계")) //2021-06-10 GLS는 품명을 품번으로 변경하여 사용
                {
                    FillGridArticle();

                    if (dgdByArticle.Items.Count <= 0)
                    {
                        MessageBox.Show("조회된 데이터가 없습니다.");
                        return;
                    }
                }
                else if (nowTab.Header.ToString().Equals("작업자별 집계"))
                {
                    FillGridWorker();

                    if (dgdByWorker.Items.Count <= 0)
                    {
                        MessageBox.Show("조회된 데이터가 없습니다.");
                        return;
                    }
                }
                else if(nowTab.Header.ToString().Equals("일별 집계"))
                {
                    FillGrid_ThisMonth();

                    if (DataGridThisMonth.Items.Count <= 0)
                    {
                        MessageBox.Show("조회된 데이터가 없습니다.");
                        return;
                    }
                }
            }
        }

        #region 주요 메서드 - 공정별 호기별 집계 조회 FillGridProcessMachine
        private void FillGridProcessMachine()
        {
            dgdByProcessTotal.Items.Clear();

            if (dgdByProcess.Items.Count > 0)
            {
                dgdByProcess.Items.Clear();
            }

            try
            {
                Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                sqlParameter.Clear();

                sqlParameter.Add("sFromDate", dtpSDate.SelectedDate != null ? dtpSDate.SelectedDate.Value.ToString("yyyyMMdd") : "");
                sqlParameter.Add("sToDate", dtpEDate.SelectedDate != null ? dtpEDate.SelectedDate.Value.ToString("yyyyMMdd") : "");
                sqlParameter.Add("sProcessIDS", cboProcess.SelectedValue != null ? cboProcess.SelectedValue.ToString() : "");
                sqlParameter.Add("sMachineIDS", cboMachine.SelectedValue != null ? cboMachine.SelectedValue.ToString() : "");
                sqlParameter.Add("ArticleID", chkArticle.IsChecked == true && txtArticle.Tag != null ? txtArticle.Tag.ToString() : "");
                sqlParameter.Add("CustomID", chkCustom.IsChecked == true && txtCustom.Tag != null ? txtCustom.Tag.ToString() : "");
                sqlParameter.Add("nOrderID", chkOrder.IsChecked == true ? (rbnOrderID.IsChecked == true ? 1 : 2) : 0);
                sqlParameter.Add("sOrderID", chkOrder.IsChecked == true ? (rbnOrderID.IsChecked == true ? txtOrder.Tag.ToString() : txtOrder.Text) : "");
                sqlParameter.Add("sWorker", chkPerson.IsChecked == true ? txtPerson.Text : "");
                sqlParameter.Add("nBuySaleMainYN", chkMainItem.IsChecked == true ? 1 : 0);
                sqlParameter.Add("BuyerArticleNoID", CheckBoxBuyerArticleNoSearch.IsChecked == true && TextBoxBuyerArticleNoSearch.Tag != null ? TextBoxBuyerArticleNoSearch.Tag.ToString() : "");

                DataSet ds = DataStore.Instance.ProcedureToDataSet("xp_prd_sWKResultByProcessMachine", sqlParameter, false);

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
                            var WinM = new Win_Prd_ProcessResultSum_Q_ByProcessMachine()
                            {
                                cls = dr["cls"].ToString(),
                                ProcessID = dr["ProcessID"].ToString(),
                                Process = dr["Process"].ToString(),
                                BuyerModel = dr["BuyerModel"].ToString(),
                                MachineID = dr["MachineID"].ToString(),
                                MachineNo = dr["MachineNo"].ToString(),
                                Article = dr["Article"].ToString(),
                                BuyerArticleNo = dr["BuyerArticleNo"].ToString(),
                                KCustom = dr["KCustom"].ToString(),
                                QtyPerBox = stringFormatN0(dr["QtyPerBox"]),
                                WorkQty = Convert.ToDouble(dr["WorkQty"]),
                                UnitPrice = stringFormatN0(dr["UnitPrice"]),
                                Amount = stringFormatN0(dr["Amount"]),
                                WorkTime = stringFormatN1(dr["WorkTime"]),
                                Num = i
                            };

                            if (WinM.cls.Equals("2")) // 호기계
                            {
                                WinM.BuyerModel = "호기계";

                                WinM.QtyPerBox = "";
                                dgdByProcess.Items.Add(WinM);
                            }
                            else if (WinM.cls.Equals("3")) // 공정계
                            {
                                WinM.MachineNo = "공정계";

                                WinM.QtyPerBox = "";
                                dgdByProcess.Items.Add(WinM);
                            }
                            else if (WinM.cls.Equals("9")) // 총계
                            {
                                WinM.Process = "총계";

                                WinM.QtyPerBox = "";

                                dgdByProcessTotal.Items.Add(WinM);
                            }
                            else
                            {
                                dgdByProcess.Items.Add(WinM);
                            }

                            
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

        #endregion // 공정별 호기별 집계

        #region 주요 메서드 - 품명별 집계 조회 FillGridArticle
        private void FillGridArticle()  //2021-06-10 GLS는 품번별로 변경
        {
            dgdByArticleTotal.Items.Clear();

            if (dgdByArticle.Items.Count > 0)
            {
                dgdByArticle.Items.Clear();
            }

            try
            {
                Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                sqlParameter.Clear();

                sqlParameter.Add("sFromDate", dtpSDate.SelectedDate != null ? dtpSDate.SelectedDate.Value.ToString("yyyyMMdd") : "");
                sqlParameter.Add("sToDate", dtpEDate.SelectedDate != null ? dtpEDate.SelectedDate.Value.ToString("yyyyMMdd") : "");
                sqlParameter.Add("sProcessIDS", cboProcess.SelectedValue != null ? cboProcess.SelectedValue.ToString() : "");
                sqlParameter.Add("sMachineIDS", cboMachine.SelectedValue != null ? cboMachine.SelectedValue.ToString() : "");
                sqlParameter.Add("ArticleID", chkArticle.IsChecked == true && txtArticle.Tag != null ? txtArticle.Tag.ToString() : "");
                sqlParameter.Add("CustomID", chkCustom.IsChecked == true && txtCustom.Tag != null ? txtCustom.Tag.ToString() : "");
                sqlParameter.Add("nOrderID", chkOrder.IsChecked == true ? (rbnOrderID.IsChecked == true ? 1 : 2) : 0);
                sqlParameter.Add("sOrderID", chkOrder.IsChecked == true ? (rbnOrderID.IsChecked == true ? txtOrder.Tag.ToString() : txtOrder.Text) : "");
                sqlParameter.Add("sWorker", chkPerson.IsChecked == true ? txtPerson.Text : "");
                sqlParameter.Add("nBuySaleMainYN", chkMainItem.IsChecked == true ? 1 : 0);
                sqlParameter.Add("BuyerArticleNoID", CheckBoxBuyerArticleNoSearch.IsChecked == true && TextBoxBuyerArticleNoSearch.Tag != null ? TextBoxBuyerArticleNoSearch.Tag.ToString() : "");

                DataSet ds = DataStore.Instance.ProcedureToDataSet("xp_prd_sWKResultByArticle", sqlParameter, false);

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
                            var WinA = new Win_Prd_ProcessResultSum_Q_ByArticle()
                            {
                                Num = i,
                                cls = dr["cls"].ToString(),
                                CustomID = dr["CustomID"].ToString(),
                                KCustom = dr["KCustom"].ToString(),
                                ArticleID = dr["ArticleID"].ToString(),
                                Article = dr["Article"].ToString(),
                                BuyerArticleNo = dr["BuyerArticleNo"].ToString(),
                                ProcessID = dr["ProcessID"].ToString(),
                                Process = dr["Process"].ToString(),
                                WorkQty = Convert.ToDouble(dr["WorkQty"]),
                                ProdQtyPerBox = stringFormatN0(dr["ProdQtyPerBox"]),
                                
                            };

                            if (WinA.cls.Trim().Equals("3")) // 품명계
                            {
                                WinA.BuyerArticleNo = "품명계";
                                dgdByArticle.Items.Add(WinA);
                            }
                            else if (WinA.cls.Trim().Equals("9")) // 총계
                            {
                                WinA.Process = "총계";
                                WinA.BuyerArticleNo  = "";
                                dgdByArticleTotal.Items.Add(WinA);
                            }
                            else
                            {
                                dgdByArticle.Items.Add(WinA);
                            }

                            
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
        #endregion // 품명별 집계

        #region 주요 메서드 - 작업자별 집계 조회 FillGridWorker

        private void FillGridWorker()
        {
            dgdByWorkerTotal.Items.Clear();

            if (dgdByWorker.Items.Count > 0)
            {
                dgdByWorker.Items.Clear();
            }

            try
            {
                Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                sqlParameter.Clear();

                sqlParameter.Add("sFromDate", dtpSDate.SelectedDate != null ? dtpSDate.SelectedDate.Value.ToString("yyyyMMdd") : "");
                sqlParameter.Add("sToDate", dtpEDate.SelectedDate != null ? dtpEDate.SelectedDate.Value.ToString("yyyyMMdd") : "");
                sqlParameter.Add("sProcessIDS", cboProcess.SelectedValue != null ? cboProcess.SelectedValue.ToString() : "");
                sqlParameter.Add("sMachineIDS", cboMachine.SelectedValue != null ? cboMachine.SelectedValue.ToString() : "");
                sqlParameter.Add("ArticleID", chkArticle.IsChecked == true && txtArticle.Tag != null ? txtArticle.Tag.ToString() : "");
                sqlParameter.Add("CustomID", chkCustom.IsChecked == true && txtCustom.Tag != null ? txtCustom.Tag.ToString() : "");
                sqlParameter.Add("nOrderID", chkOrder.IsChecked == true ? (rbnOrderID.IsChecked == true ? 1 : 2) : 0);
                sqlParameter.Add("sOrderID", chkOrder.IsChecked == true ? (rbnOrderID.IsChecked == true ? txtOrder.Tag.ToString() : txtOrder.Text) : "");
                sqlParameter.Add("sWorker", chkPerson.IsChecked == true ? txtPerson.Text : "");
                sqlParameter.Add("nBuySaleMainYN", chkMainItem.IsChecked == true ? 1 : 0);
                sqlParameter.Add("BuyerArticleNoID", CheckBoxBuyerArticleNoSearch.IsChecked == true && TextBoxBuyerArticleNoSearch.Tag != null ? TextBoxBuyerArticleNoSearch.Tag.ToString() : "");


                DataSet ds = DataStore.Instance.ProcedureToDataSet("xp_prd_sWKResultByWorker", sqlParameter, false);

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
                            var WinW = new Win_Prd_ProcessResultSum_Q_ByWorker()
                            {
                                Num = i,
                                cls = dr["cls"].ToString().Trim(),

                                WorkPersonID = dr["WorkPersonID"].ToString(),
                                Name = dr["Name"].ToString(),
                                ProcessID = dr["ProcessID"].ToString(),
                                Process = dr["Process"].ToString(),
                                MachineID = dr["MachineID"].ToString(),

                                Machine = dr["Machine"].ToString(),
                                MachineNo = dr["MachineNo"].ToString(),
                                BuyerModelID = dr["BuyerModelID"].ToString(),
                                Model = dr["Model"].ToString(),
                                ArticleID = dr["ArticleID"].ToString(),

                                Article = dr["Article"].ToString(),
                                BuyerArticleNo = dr["BuyerArticleNo"].ToString(),
                                CustomID = dr["CustomID"].ToString(),
                                KCustom = dr["KCustom"].ToString(),
                                WorkQty = Convert.ToDouble(dr["WorkQty"]),

                                ProdQtyPerBox = stringFormatN0(dr["ProdQtyPerBox"]),
                                
                            };

                            if (WinW.cls.Trim().Equals("3")) // 작업자계
                            {
                                WinW.Process = "작업자계";
                                dgdByWorker.Items.Add(WinW);
                            }
                            else if (WinW.cls.Trim().Equals("9")) // 총계
                            {
                                WinW.Process = "총계";
                                WinW.Name = "";

                                dgdByWorkerTotal.Items.Add(WinW);
                            }
                            else
                            {
                                dgdByWorker.Items.Add(WinW);
                            }
                            
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

        #endregion // 주요 메서드 - 작업자별 집계 조회 FillGridWorker

        #region 조회 일자별 집계
        private void FillGrid_ThisMonth()
        {
            if (DataGridThisMonth.Items.Count > 0)
            {
                DataGridThisMonth.Items.Clear();
            }

            try
            {
                Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                sqlParameter.Clear();
                sqlParameter.Add("ChkDate", chkDateSrh.IsChecked == true ? 1 : 0);
                sqlParameter.Add("sFromDate", chkDateSrh.IsChecked == true ? (dtpSDate.SelectedDate == null ? "" : dtpSDate.SelectedDate.Value.ToString("yyyyMMdd")) : "");
                sqlParameter.Add("sToDate", chkDateSrh.IsChecked == true ? (dtpEDate.SelectedDate == null ? "" : dtpEDate.SelectedDate.Value.ToString("yyyyMMdd")) : "");
                sqlParameter.Add("ChkProcessID", CheckBoxProcessSearch.IsChecked == true && cboProcess.SelectedValue.ToString() != "" ? 1 : 0);
                sqlParameter.Add("sProcessID", CheckBoxProcessSearch.IsChecked == true ? (cboProcess.SelectedValue == null ? "" : cboProcess.SelectedValue.ToString()) : "");
                sqlParameter.Add("ChkMachineID", CheckBoxMachineSearch.IsChecked == true && cboMachine.SelectedValue.ToString() != "" ? 1 : 0);
                sqlParameter.Add("sMachineID", CheckBoxMachineSearch.IsChecked == true ? (cboMachine.SelectedValue == null ? "" : cboMachine.SelectedValue.ToString()) : "");
                sqlParameter.Add("ChkWorker", chkPerson.IsChecked == true ? 1 : 0);
                sqlParameter.Add("sWorker", chkPerson.IsChecked == true ? (txtPerson.Text == string.Empty ? "" : txtPerson.Text) : "");
                sqlParameter.Add("ChkOrderID", chkOrder.IsChecked == true ? 1 : 0);
                sqlParameter.Add("sOrderID", chkOrder.IsChecked == true ? (txtOrder.Tag == null ? "" : txtOrder.Tag.ToString()) : "");
                sqlParameter.Add("ChkCustomID", chkCustom.IsChecked == true ? 1 : 0);
                sqlParameter.Add("sCustomID", chkCustom.IsChecked == true ? (txtCustom.Tag == null ? "" : txtCustom.Tag.ToString()) : "");
                sqlParameter.Add("ChkArticleID", chkArticle.IsChecked == true ? 1 : 0);
                sqlParameter.Add("sArticleID", chkArticle.IsChecked == true && txtArticle.Tag != null ? txtArticle.Tag.ToString() : "");
                sqlParameter.Add("ChkBuySaleMainYN", chkMainItem.IsChecked == true ? 1 : 0);
                sqlParameter.Add("ChkBuyerArticleNo", CheckBoxBuyerArticleNoSearch.IsChecked == true ? 1 : 0);
                sqlParameter.Add("BuyerArticleNo", (CheckBoxBuyerArticleNoSearch.IsChecked == true && TextBoxBuyerArticleNoSearch.Tag != null) ? TextBoxBuyerArticleNoSearch.Tag.ToString() : "");


                DataSet ds = DataStore.Instance.ProcedureToDataSet("xp_prd_sWKResult_Article_ThisMonth", sqlParameter, false);

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
                            var WPPQCT = new Win_Prd_ProcessResultSum_Q_CodeView_ThisMonth()
                            {
                                Num = i,

                                Article = dr["Article"].ToString(),
                                ArticleID = dr["ArticleID"].ToString(),
                                BuyerArticleNo = dr["BuyerArticleNo"].ToString(),

                                SDay01 = Convert.ToDouble(dr["SDay01"]),
                                SDay02 = Convert.ToDouble(dr["SDay02"]),
                                SDay03 = Convert.ToDouble(dr["SDay03"]) ,
                                SDay04 = Convert.ToDouble(dr["SDay04"]),
                                SDay05 = Convert.ToDouble(dr["SDay05"]),
                                SDay06 = Convert.ToDouble(dr["SDay06"]),
                                SDay07 = Convert.ToDouble(dr["SDay07"]),
                                SDay08 = Convert.ToDouble(dr["SDay08"]),
                                SDay09 = Convert.ToDouble(dr["SDay09"]),
                                SDay10 = Convert.ToDouble(dr["SDay10"]),
                                SDay11 = Convert.ToDouble(dr["SDay11"]),
                                SDay12 = Convert.ToDouble(dr["SDay12"]),
                                SDay13 = Convert.ToDouble(dr["SDay13"]),
                                SDay14 = Convert.ToDouble(dr["SDay14"]),
                                SDay15 = Convert.ToDouble(dr["SDay15"]),
                                SDay16 = Convert.ToDouble(dr["SDay16"]),
                                SDay17 = Convert.ToDouble(dr["SDay17"]),
                                SDay18 = Convert.ToDouble(dr["SDay18"]),
                                SDay19 = Convert.ToDouble(dr["SDay19"]),
                                SDay20 = Convert.ToDouble(dr["SDay20"]),
                                SDay21 = Convert.ToDouble(dr["SDay21"]),
                                SDay22 = Convert.ToDouble(dr["SDay22"]),
                                SDay23 = Convert.ToDouble(dr["SDay23"]),
                                SDay24 = Convert.ToDouble(dr["SDay24"]),
                                SDay25 = Convert.ToDouble(dr["SDay25"]),
                                SDay26 = Convert.ToDouble(dr["SDay26"]),
                                SDay27 = Convert.ToDouble(dr["SDay27"]),
                                SDay28 = Convert.ToDouble(dr["SDay28"]),
                                SDay29 = Convert.ToDouble(dr["SDay29"]),
                                SDay30 = Convert.ToDouble(dr["SDay30"]),
                                SDay31 = Convert.ToDouble(dr["SDay31"]),
                            };

                            double sum = WPPQCT.SDay01 + WPPQCT.SDay02 + WPPQCT.SDay03 + WPPQCT.SDay04 + WPPQCT.SDay05
                                + WPPQCT.SDay06 + WPPQCT.SDay07 + WPPQCT.SDay08 + WPPQCT.SDay09 + WPPQCT.SDay10
                                + WPPQCT.SDay11 + WPPQCT.SDay12 + WPPQCT.SDay13 + WPPQCT.SDay14 + WPPQCT.SDay15
                                + WPPQCT.SDay16 + WPPQCT.SDay17 + WPPQCT.SDay18 + WPPQCT.SDay19 + WPPQCT.SDay20
                                + WPPQCT.SDay21 + WPPQCT.SDay22 + WPPQCT.SDay23 + WPPQCT.SDay24 + WPPQCT.SDay25
                                + WPPQCT.SDay26 + WPPQCT.SDay27 + WPPQCT.SDay28 + WPPQCT.SDay29 + WPPQCT.SDay30
                                + WPPQCT.SDay31;

                            WPPQCT.TotalQty = sum;


                            //int sum = Convert.ToInt32(WPPQCT.SDay01) + Convert.ToInt32(WPPQCT.SDay02) + Convert.ToInt32(WPPQCT.SDay03)
                            //        + Convert.ToInt32(WPPQCT.SDay04) + Convert.ToInt32(WPPQCT.SDay05) + Convert.ToInt32(WPPQCT.SDay06)
                            //        + Convert.ToInt32(WPPQCT.SDay07) + Convert.ToInt32(WPPQCT.SDay08) + Convert.ToInt32(WPPQCT.SDay09)
                            //        + Convert.ToInt32(WPPQCT.SDay10) + Convert.ToInt32(WPPQCT.SDay11) + Convert.ToInt32(WPPQCT.SDay12)
                            //        + Convert.ToInt32(WPPQCT.SDay13) + Convert.ToInt32(WPPQCT.SDay14) + Convert.ToInt32(WPPQCT.SDay15)
                            //        + Convert.ToInt32(WPPQCT.SDay16) + Convert.ToInt32(WPPQCT.SDay17) + Convert.ToInt32(WPPQCT.SDay18)
                            //        + Convert.ToInt32(WPPQCT.SDay19) + Convert.ToInt32(WPPQCT.SDay20) + Convert.ToInt32(WPPQCT.SDay21)
                            //        + Convert.ToInt32(WPPQCT.SDay22) + Convert.ToInt32(WPPQCT.SDay23) + Convert.ToInt32(WPPQCT.SDay24)
                            //        + Convert.ToInt32(WPPQCT.SDay25) + Convert.ToInt32(WPPQCT.SDay26) + Convert.ToInt32(WPPQCT.SDay27)
                            //        + Convert.ToInt32(WPPQCT.SDay28) + Convert.ToInt32(WPPQCT.SDay29) + Convert.ToInt32(WPPQCT.SDay30)
                            //        + Convert.ToInt32(WPPQCT.SDay31);
                            //WPPQCT.TotalQty = lib.returnNumStringZero(Convert.ToString(sum));


                            //WPPQCT.SDay01 = lib.returnNumStringZero(WPPQCT.SDay01);
                            //WPPQCT.SDay02 = lib.returnNumStringZero(WPPQCT.SDay02);
                            //WPPQCT.SDay03 = lib.returnNumStringZero(WPPQCT.SDay03);
                            //WPPQCT.SDay04 = lib.returnNumStringZero(WPPQCT.SDay04);
                            //WPPQCT.SDay05 = lib.returnNumStringZero(WPPQCT.SDay05);
                            //WPPQCT.SDay06 = lib.returnNumStringZero(WPPQCT.SDay06);
                            //WPPQCT.SDay07 = lib.returnNumStringZero(WPPQCT.SDay07);
                            //WPPQCT.SDay08 = lib.returnNumStringZero(WPPQCT.SDay08);
                            //WPPQCT.SDay09 = lib.returnNumStringZero(WPPQCT.SDay09);
                            //WPPQCT.SDay10 = lib.returnNumStringZero(WPPQCT.SDay10);
                            //WPPQCT.SDay11 = lib.returnNumStringZero(WPPQCT.SDay11);
                            //WPPQCT.SDay12 = lib.returnNumStringZero(WPPQCT.SDay12);
                            //WPPQCT.SDay13 = lib.returnNumStringZero(WPPQCT.SDay13);
                            //WPPQCT.SDay14 = lib.returnNumStringZero(WPPQCT.SDay14);
                            //WPPQCT.SDay15 = lib.returnNumStringZero(WPPQCT.SDay15);
                            //WPPQCT.SDay16 = lib.returnNumStringZero(WPPQCT.SDay16);
                            //WPPQCT.SDay17 = lib.returnNumStringZero(WPPQCT.SDay17);
                            //WPPQCT.SDay18 = lib.returnNumStringZero(WPPQCT.SDay18);
                            //WPPQCT.SDay19 = lib.returnNumStringZero(WPPQCT.SDay19);
                            //WPPQCT.SDay20 = lib.returnNumStringZero(WPPQCT.SDay20);
                            //WPPQCT.SDay21 = lib.returnNumStringZero(WPPQCT.SDay21);
                            //WPPQCT.SDay22 = lib.returnNumStringZero(WPPQCT.SDay22);
                            //WPPQCT.SDay23 = lib.returnNumStringZero(WPPQCT.SDay23);
                            //WPPQCT.SDay24 = lib.returnNumStringZero(WPPQCT.SDay24);
                            //WPPQCT.SDay25 = lib.returnNumStringZero(WPPQCT.SDay25);
                            //WPPQCT.SDay26 = lib.returnNumStringZero(WPPQCT.SDay26);
                            //WPPQCT.SDay27 = lib.returnNumStringZero(WPPQCT.SDay27);
                            //WPPQCT.SDay28 = lib.returnNumStringZero(WPPQCT.SDay28);
                            //WPPQCT.SDay29 = lib.returnNumStringZero(WPPQCT.SDay29);
                            //WPPQCT.SDay30 = lib.returnNumStringZero(WPPQCT.SDay30);
                            //WPPQCT.SDay31 = lib.returnNumStringZero(WPPQCT.SDay31);

                            DataGridThisMonth.Items.Add(WPPQCT);
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

        #region 텍스트박스 공용 키다운 이벤트
        private void txtBox_KeyDown_Search(object sender, KeyEventArgs e)
        {
            using (Loading lw = new Loading(re_Search))
            {
                lw.ShowDialog();
            }
        }
        #endregion

        private void LabelBuyerArticleNoSearch_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            if(CheckBoxBuyerArticleNoSearch.IsChecked == true)
            {
                CheckBoxBuyerArticleNoSearch.IsChecked = false;
            }
            else
            {
                CheckBoxBuyerArticleNoSearch.IsChecked = true;
            }
        }

        private void CheckBoxBuyerArticleNoSearch_Checked(object sender, RoutedEventArgs e)
        {
            TextBoxBuyerArticleNoSearch.IsEnabled = true;
            ButtonBuyerArticleNoSearch.IsEnabled = true;
            TextBoxBuyerArticleNoSearch.Focus();
        }

        private void CheckBoxBuyerArticleNoSearch_Unchecked(object sender, RoutedEventArgs e)
        {
            TextBoxBuyerArticleNoSearch.IsEnabled = false;
            ButtonBuyerArticleNoSearch.IsEnabled = false;
        }

        private void TextBoxBuyerArticleNoSearch_KeyDown(object sender, KeyEventArgs e)
        {
            if(e.Key == Key.Enter)
            {
                MainWindow.pf.ReturnCode(TextBoxBuyerArticleNoSearch, 76, "");
            }
        }

        private void ButtonBuyerArticleNoSearch_Click(object sender, RoutedEventArgs e)
        {
            MainWindow.pf.ReturnCode(TextBoxBuyerArticleNoSearch, 76, "");
        }
    }

    class Win_Prd_ProcessResultSum_Q_ByProcessMachine : BaseView
    {
        public override string ToString()
        {
            return (this.ReportAllProperties());
        }

        public int Num { get; set; }
        public string cls { get; set; }
        public string ProcessID { get; set; }
        public string Process { get; set; }
        public string BuyerModel { get; set; }
        public string MachineID { get; set; }
        public string MachineNo { get; set; }
        public string Article { get; set; }
        //public string ArticleID { get; set; }
        public string KCustom { get; set; }
        //public string ProdQtyPerBox { get; set; }
        public double WorkQty { get; set; }
        public string UnitPrice { get; set; }
        public string Amount { get; set; }
        public string WorkTime { get; set; }
        //public string OutQtyPerBox { get; set; }
        public string QtyPerBox { get; set; }
        public string BuyerArticleNo { get; set; }
    }

    class Win_Prd_ProcessResultSum_Q_ByArticle : BaseView
    {
        public override string ToString()
        {
            return (this.ReportAllProperties());
        }

        public string cls { get; set; }

        public string CustomID { get; set; }
        public string KCustom { get; set; }
        public string ArticleID { get; set; }
        public string Article { get; set; }

        public string BuyerArticleNo { get; set; }
        public string ProcessID { get; set; }
        public string Process { get; set; }
        public string BuyerModelID { get; set; }
        public string Model { get; set; }

        public double WorkQty { get; set; }
        public string WorkTime { get; set; }
        public string ProdQtyPerBox { get; set; }

        public int Num { get; set; }
    }

    class Win_Prd_ProcessResultSum_Q_ByWorker : BaseView
    {
        public override string ToString()
        {
            return (this.ReportAllProperties());
        }

        public string cls { get; set; }
        public string WorkPersonID { get; set; }
        public string Name { get; set; }
        public string ProcessID { get; set; }
        public string Process { get; set; }

        public string MachineID { get; set; }
        public string Machine { get; set; }
        public string MachineNo { get; set; }
        public string BuyerModelID { get; set; }        
        public string Model { get; set; }

        public string ArticleID { get; set; }
        public string Article { get; set; }
        public string BuyerArticleNo { get; set; }
        public string CustomID { get; set; }
        public string KCustom { get; set; }

        public double WorkQty { get; set; }
        public string ProdQtyPerBox { get; set; }
        public int Num { get; set; }

        public string BuyerModel { get; set; }
    }

    class Win_Prd_ProcessResultSum_Q_CodeView_ThisMonth : BaseView
    {
        public override string ToString()
        {
            return (this.ReportAllProperties());
        }

        public int Num { get; set; }

        public string Article { get; set; }
        public string ArticleID { get; set; }
        public string BuyerArticleNo { get; set; }
        public double TotalQty { get; set; }
        public double SDay01 { get; set; }
        public double SDay02 { get; set; }
        public double SDay03 { get; set; }
        public double SDay04 { get; set; }
        public double SDay05 { get; set; }
        public double SDay06 { get; set; }
        public double SDay07 { get; set; }
        public double SDay08 { get; set; }
        public double SDay09 { get; set; }
        public double SDay10 { get; set; }
        public double SDay11 { get; set; }
        public double SDay12 { get; set; }
        public double SDay13 { get; set; }
        public double SDay14 { get; set; }
        public double SDay15 { get; set; }
        public double SDay16 { get; set; }
        public double SDay17 { get; set; }
        public double SDay18 { get; set; }
        public double SDay19 { get; set; }
        public double SDay20 { get; set; }
        public double SDay21 { get; set; }
        public double SDay22 { get; set; }
        public double SDay23 { get; set; }
        public double SDay24 { get; set; }
        public double SDay25 { get; set; }
        public double SDay26 { get; set; }
        public double SDay27 { get; set; }
        public double SDay28 { get; set; }
        public double SDay29 { get; set; }
        public double SDay30 { get; set; }
        public double SDay31 { get; set; }
    }
}

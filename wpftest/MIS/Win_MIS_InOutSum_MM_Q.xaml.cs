using System;
using System.Collections.Generic;
using System.Data;
using System.Threading;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using WizMes_HanYoung.PopUP;
using WPF.MDI;

namespace WizMes_HanYoung
{
    /// <summary>
    /// Win_MIS_InOutSum_MM_Q.xaml에 대한 상호 작용 논리
    /// </summary>
    public partial class Win_MIS_InOutSum_MM_Q : UserControl
    {
        #region 변수 선언 및 로드

        WizMes_HanYoung.PopUp.NoticeMessage msg = new PopUp.NoticeMessage();
        Lib lib = new Lib();
        public Win_MIS_InOutSum_MM_Q()
        {
            InitializeComponent();
        }

        private void UserControl_Loaded(object sender, RoutedEventArgs e)
        {
            lib.UiLoading(sender);
            dtpSDate.SelectedDate = DateTime.Today;
            dtpEDate.SelectedDate = DateTime.Today;
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
                FillGrid();     //왼쪽 그리드 조회
                FillGrid2();    //오른쪽 그리드 조회

                if (dgdmain2.Items.Count == 0 || (dgdmain.Items.Count == 0))
                {
                    MessageBox.Show("조회결과가 없습니다.");
                }
                else if (dgdmain.Items.Count > 0)
                {
                    dgdmain.SelectedIndex = 0;
                    this.DataContext = dgdmain.SelectedItem as Win_MIS_In_CodeView;
                }
                else if (dgdmain2.Items.Count > 0)
                {
                    dgdmain2.SelectedIndex = 0;
                    this.DataContext = dgdmain2.SelectedItem as Win_MIS_Out_CodeView;
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
            if (dgdmain.Items.Count < 1)
            {
                MessageBox.Show("먼저 검색해 주세요.");
                return;
            }

            if (dgdmain2.Items.Count < 1)
            {
                MessageBox.Show("먼저 검색해 주세요.");
                return;
            }

            Lib lib = new Lib();
            DataTable dt = null;
            string Name = string.Empty;

            string[] lst = new string[4];
            lst[0] = "입고 조회";
            lst[1] = "출고 조회";
            lst[2] = dgdmain.Name;
            lst[3] = dgdmain2.Name;

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
                    lib.GenerateExcel(dt, Name);
                    lib.excel.Visible = true;

                    if (lib.GenerateExcel(dt, Name))
                        lib.excel.Visible = true;

                    else if (ExpExc.choice.Equals(dgdmain2.Name))
                    {
                        if (ExpExc.Check.Equals("Y"))
                            dt = lib.DataGridToDTinHidden(dgdmain2);
                        else
                            dt = lib.DataGirdToDataTable(dgdmain2);
                        Name = dgdmain2.Name;
                        lib.GenerateExcel(dt, Name);
                        lib.excel.Visible = true;
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
        }

        #endregion

        #region 검색조건
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
                MainWindow.pf.ReturnCode(txtCustom, (int)Defind_CodeFind.DCF_CUSTOM, "");
            }
        }

        //거래처
        private void btnPfCustom_Click(object sender, RoutedEventArgs e)
        {
            MainWindow.pf.ReturnCode(txtCustom, (int)Defind_CodeFind.DCF_CUSTOM, "");
        }

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

        #endregion

        //왼쪽 입고 그리드 조회
        private void FillGrid()
        {
            if (dgdmain.Items.Count > 0)
            {
                dgdmain.Items.Clear();
            }

            dgdsum.Items.Clear();

            var insum = new Win_MIS_INSUM();

            try
            {
                Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                sqlParameter.Add("ChkDate", chkDate.IsChecked == true ? 1 : 0);
                sqlParameter.Add("SDate", chkDate.IsChecked == true ? dtpSDate.SelectedDate.Value.ToString("yyyyMMdd") : "");
                sqlParameter.Add("EDate", chkDate.IsChecked == true ? dtpEDate.SelectedDate.Value.ToString("yyyyMMdd") : "");
                sqlParameter.Add("ChkCustomID", chkCustom.IsChecked == true ? 1 : 0);
                sqlParameter.Add("CustomID", chkCustom.IsChecked == true && !txtCustom.Text.Trim().Equals("") ? txtCustom.Text : "");
                sqlParameter.Add("chkBuyerArticleNo", chkBuyerArticleNo.IsChecked == true ? 1 : 0);
                sqlParameter.Add("BuyerArticleNo", chkBuyerArticleNo.IsChecked == true && !txtBuyerArticleNo.Text.Trim().Equals("") ? txtBuyerArticleNo.Text : "");

                DataSet ds = DataStore.Instance.ProcedureToDataSet("xp_MIS_sInwareSum", sqlParameter, false);

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
                            var inware = new Win_MIS_In_CodeView()
                            {
                                Gbn = dr["Gbn"].ToString(),
                                CustomName = dr["CustomName"].ToString(),
                                IODate = Lib.Instance.StrDateTimeToSlash(dr["IODate"].ToString()),
                                BuyerArticleNo = dr["BuyerArticleNo"].ToString(),
                                Qty = dr["Qty"].ToString(),
                                UnitPrice = dr["UnitPrice"].ToString(),
                                Amount = dr["Amount"].ToString(),
                            };

                            if (dr["Gbn"].ToString().Trim().Equals("5")) //거래처계
                            {
                                inware.IODate = "";
                                inware.CustomName = "거래처별 계";
                                inware.BuyerArticleNo = "";
                                inware.TotalColor = "true";
                            }

                            if (dr["Gbn"].ToString().Trim().Equals("9")) //총계
                            {
                                insum.text = "";
                                insum.Qty = dr["Qty"].ToString();
                                //insum.Qty = Convert.ToDouble(dr["Qty"]);
                                //insum.UnitPrice = dr["UnitPrice"].ToString();
                                insum.Amount = dr["Amount"].ToString();
                                //insum.Amount = Convert.ToDouble(dr["Amount"]);
                            }

                            inware.Qty = lib.returnNumStringZero(inware.Qty);
                            inware.UnitPrice = lib.returnNumStringOne(inware.UnitPrice);
                            inware.Amount = lib.returnNumStringZero(inware.Amount);

                            insum.Qty = lib.returnNumStringZero(insum.Qty);
                            insum.Amount = lib.returnNumStringZero(insum.Amount);

                            dgdmain.Items.Add(inware);

                            if (inware.Gbn.Equals("9"))
                                dgdmain.Items.Remove(inware);

                        }
                        dgdsum.Items.Add(insum);
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

        //오른쪽 출고 그리드 조회
        private void FillGrid2()
        {
            if (dgdmain2.Items.Count > 0)
            {
                dgdmain2.Items.Clear();
            }

            dgdsum2.Items.Clear();

            var outsum = new Win_MIS_OutSum();

            try
            {
                Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                sqlParameter.Add("ChkDate", chkDate.IsChecked == true ? 1 : 0);
                sqlParameter.Add("SDate", chkDate.IsChecked == true ? dtpSDate.SelectedDate.Value.ToString("yyyyMMdd") : "");
                sqlParameter.Add("EDate", chkDate.IsChecked == true ? dtpEDate.SelectedDate.Value.ToString("yyyyMMdd") : "");
                sqlParameter.Add("ChkCustomID", chkCustom.IsChecked == true ? 1 : 0);
                sqlParameter.Add("CustomID", chkCustom.IsChecked == true && !txtCustom.Text.Trim().Equals("") ? txtCustom.Text : "");
                sqlParameter.Add("chkBuyerArticleNo", chkBuyerArticleNo.IsChecked == true ? 1 : 0);
                sqlParameter.Add("BuyerArticleNo", chkBuyerArticleNo.IsChecked == true && !txtBuyerArticleNo.Text.Trim().Equals("") ? txtBuyerArticleNo.Text : "");

                DataSet ds = DataStore.Instance.ProcedureToDataSet("xp_MIS_sOutwareSum", sqlParameter, false);

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
                            var outware = new Win_MIS_Out_CodeView()
                            {
                                Gbn = dr["Gbn"].ToString(),
                                CustomName = dr["CustomName"].ToString(),
                                OutDate = Lib.Instance.StrDateTimeToSlash(dr["OutDate"].ToString()),
                                BuyerArticleNo = dr["BuyerArticleNo"].ToString(),
                                OutQty = dr["OutQty"].ToString(),
                                UnitPrice = dr["UnitPrice"].ToString(),
                                Amount = dr["Amount"].ToString(),
                            };

                            if (dr["Gbn"].ToString().Trim().Equals("5")) //거래처계
                            {
                                outware.OutDate = "";
                                outware.CustomName = "거래처별 계";
                                outware.BuyerArticleNo = "";
                                outware.TotalColor = "true";
                            }

                            if (dr["Gbn"].ToString().Trim().Equals("9")) //총계
                            {
                                outsum.text = "";
                                outsum.OutQty = stringFormatN0(dr["OutQty"]);
                                //outsum.OutQty = Convert.ToDouble(dr["OutQty"]);
                                //outsum.UnitPrice = dr["UnitPrice"].ToString();
                                outsum.Amount = stringFormatN0(dr["Amount"]);
                                //outsum.Amount = Convert.ToDouble(dr["Amount"]);
                            }

                            outware.OutQty = lib.returnNumStringZero(outware.OutQty);
                            outware.UnitPrice = lib.returnNumStringOne(outware.UnitPrice);
                            outware.Amount = lib.returnNumStringZero(outware.Amount);


                            dgdmain2.Items.Add(outware);

                            if (outware.Gbn.Equals("9"))
                                dgdmain2.Items.Remove(outware);
                        }
                        dgdsum2.Items.Add(outsum);
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

        // 천마리 콤마, 소수점 버리기
        private string stringFormatN0(object obj)
        {
            return string.Format("{0:N0}", obj);
        }

        private void PrintWork(bool preview_click)
        {

        }

        #endregion
        /// <summary>
        /// 칼럼헤더 정렬막기
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void columnHeader_Click(object sender, RoutedEventArgs e)
        {
            foreach (DataGridColumn column in dgdmain.Columns)
            {
                column.CanUserSort = false;
            }

            foreach (DataGridColumn column in dgdmain2.Columns)
            {
                column.CanUserSort = false;
            }
        }
    }

    class Win_MIS_In_CodeView : BaseView
    {
        public override string ToString()
        {
            return (this.ReportAllProperties());
        }

        public string Gbn { get; set; }
        public string CustomName { get; set; }
        public string IODate { get; set; }
        public string BuyerArticleNo { get; set; }
        public string Qty { get; set; }
        public string UnitPrice { get; set; }
        public string Amount { get; set; }
        public string TotalColor { get; set; }
    }

    class Win_MIS_INSUM : BaseView
    {
        public string Gbn { get; set; }
        public string text { get; set; }
        public string CustomName { get; set; }
        public string IODate { get; set; }
        public string BuyerArticleNo { get; set; }
        public string Qty { get; set; }
        public string UnitPrice { get; set; }
        public string Amount { get; set; }
    }

    class Win_MIS_Out_CodeView : BaseView
    {
        public string Gbn { get; set; }
        public string CustomName { get; set; }
        public string OutDate { get; set; }
        public string BuyerArticleNo { get; set; }
        public string OutQty { get; set; }
        public string UnitPrice { get; set; }
        public string Amount { get; set; }
        public string TotalColor { get; set; }
    }

    class Win_MIS_OutSum : BaseView
    {
        public string Gbn { get; set; }
        public string text { get; set; }
        public string CustomName { get; set; }
        public string OutDate { get; set; }
        public string BuyerArticleNo { get; set; }
        public string OutQty { get; set; }
        public string UnitPrice { get; set; }
        public string Amount { get; set; }
    }
}

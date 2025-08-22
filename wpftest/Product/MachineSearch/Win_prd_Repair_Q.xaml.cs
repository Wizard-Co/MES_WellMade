using System;
using System.Collections.Generic;
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
using WizMes_WellMade.PopUP;
using WPF.MDI;

namespace WizMes_WellMade
{
    /// <summary>
    /// Win_prd_Repair_Q.xaml에 대한 상호 작용 논리
    /// </summary>
    public partial class Win_prd_Repair_Q : UserControl
    {
        #region 변수 선언 및 로드
        Lib lib = new Lib();
        WizMes_WellMade.PopUp.NoticeMessage msg = new PopUp.NoticeMessage();

        public Win_prd_Repair_Q()
        {
            InitializeComponent();
        }

        private void UserControl_Loaded(object sender, RoutedEventArgs e)
        {
            Lib.Instance.UiLoading(sender);
            dtpSDate.SelectedDate = DateTime.Today;
            dtpEDate.SelectedDate = DateTime.Today;
            chkDate.IsChecked = true;
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
            dtpSDate.IsEnabled = true;
            dtpEDate.IsEnabled = true;
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

        #region 기계명 관련 이벤트

        //기계명
        private void lblMCID_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            if (chkMCID.IsChecked == true) { chkMCID.IsChecked = false; }
            else { chkMCID.IsChecked = true; }
        }

        //기계명
        private void chkMCID_Checked(object sender, RoutedEventArgs e)
        {
            txtMCID.IsEnabled = true;
            btnPfMCID.IsEnabled = true;
            txtMCID.Focus();
        }

        //기계명
        private void chkMCID_Unchecked(object sender, RoutedEventArgs e)
        {
            txtMCID.IsEnabled = false;
            btnPfMCID.IsEnabled = false;
        }

        //기계명
        private void txtMCID_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                MainWindow.pf.ReturnCode(txtMCID, 12, "");
            }
        }

        //기계명
        private void btnPfMCID_Click(object sender, RoutedEventArgs e)
        {
            MainWindow.pf.ReturnCode(txtMCID, 12, "");
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
                //로직
                if (dgdRepairQ.Items.Count > 0)
                {
                    dgdRepairQ.Items.Clear();
                }

                FillGrid();

                if (dgdRepairQ.Items.Count > 0)
                {
                    dgdRepairQ.SelectedIndex = 0;
                    this.DataContext = dgdRepairQ.SelectedItem as Win_prd_Repair_Q_CodeView;
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
            lst[0] = "설비 수리 조회";
            lst[1] = dgdRepairQ.Name;

            ExportExcelxaml ExpExc = new ExportExcelxaml(lst);

            ExpExc.ShowDialog();

            if (ExpExc.DialogResult.HasValue)
            {
                if (ExpExc.choice.Equals(dgdRepairQ.Name))
                {
                    if (ExpExc.Check.Equals("Y"))
                        dt = Lib.Instance.DataGridToDTinHidden(dgdRepairQ);
                    else
                        dt = Lib.Instance.DataGirdToDataTable(dgdRepairQ);

                    Name = dgdRepairQ.Name;

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

        #region 부품명 검색 조건 이벤트
        //상단 부품명 check
        private void chkArticleSearch_Checked(object sender, RoutedEventArgs e)
        {
            txtArticleSearch.IsEnabled = true;
            btnPfArticleSearch.IsEnabled = true;
            txtArticleSearch.Focus();
        }

        //상단 부품명 Uncheck
        private void chkArticleSearch_Unchecked(object sender, RoutedEventArgs e)
        {
            txtArticleSearch.Text = "";
            txtArticleSearch.IsEnabled = false;
            btnPfArticleSearch.IsEnabled = false;
        }

        //부품명
        private void lblArticleSearch_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            if (chkArticleSearch.IsChecked == true) { chkArticleSearch.IsChecked = false; }
            else { chkArticleSearch.IsChecked = true; }
        }

        private void txtArticleSearch_KeyDown(object sender, KeyEventArgs e)
        {
            try
            {
                if (e.Key == Key.Enter)
                {
                    MainWindow.pf.ReturnCode(txtArticleSearch, (int)Defind_CodeFind.DCF_PART, "");
                }
            }
            catch (Exception ee)
            {
                MessageBox.Show("오류지점 - txtArticleSearch_KeyDown : " + ee.ToString());
            }
        }

        //부품명 플러스 파인더
        private void btnPfArticleSearch_Click(object sender, RoutedEventArgs e)
        {
            MainWindow.pf.ReturnCode(txtArticleSearch, (int)Defind_CodeFind.DCF_PART, "");
        }

        #endregion

        //실질 조회
        private void FillGrid()
        {
            try
            {
                Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                sqlParameter.Add("nChkDate", chkDate.IsChecked == true ? 1 : 0);
                sqlParameter.Add("StartDate", chkDate.IsChecked == true ? dtpSDate.SelectedDate.Value.ToString("yyyyMMdd") : "");
                sqlParameter.Add("EndDate", chkDate.IsChecked == true ? dtpEDate.SelectedDate.Value.ToString("yyyyMMdd") : "");
                sqlParameter.Add("nMCID", chkMCID.IsChecked == true ? (txtMCID.Tag != null ? 1 : 2) : 0);
                sqlParameter.Add("sMCID", chkMCID.IsChecked == true ? (txtMCID.Tag != null ? txtMCID.Tag.ToString() : txtMCID.Text) : "");
                //2021-08-03 부품명 검색조건 추가
                sqlParameter.Add("nArticle", chkArticleSearch.IsChecked == true ? (txtArticleSearch.Tag != null ? 1 : 2) : 0);
                sqlParameter.Add("sArticle", chkArticleSearch.IsChecked == true ? (txtArticleSearch.Tag != null ? txtArticleSearch.Tag.ToString() : txtArticleSearch.Text) : "");



                DataSet ds = DataStore.Instance.ProcedureToDataSet("xp_mcRepairQ_sRepair", sqlParameter, false);
                dgdTotal.Items.Clear();
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
                            var WinRepairQ = new Win_prd_Repair_Q_CodeView()
                            {
                                Num = i.ToString(),
                                RepairGubun = dr["RepairGubun"].ToString(),
                                mcid = dr["mcid"].ToString(),
                                mcname = dr["mcname"].ToString(),
                                repairdate = dr["repairdate"].ToString(),
                                repairremark = dr["repairremark"].ToString(),
                                MCPartID = dr["MCPartID"].ToString(),
                                MCPartName = dr["MCPartName"].ToString(),
                                partcnt = stringFormatN0(dr["partcnt"]),
                                partremark = dr["partremark"].ToString(),
                                RepairID = dr["RepairID"].ToString(),
                                repairsubseq = dr["repairsubseq"].ToString(),
                                MCCustom = dr["MCCustom"].ToString(),
                                partprice = stringFormatN0(dr["partprice"]),
                                reason = dr["reason"].ToString(),
       
                            };

                            if (WinRepairQ.RepairGubun.Equals("1"))
                            {
                                WinRepairQ.RepairGubun_CV = "수리";
                            }
                            else if (WinRepairQ.RepairGubun.Equals("2"))
                            {
                                WinRepairQ.RepairGubun_CV = "교체";
                            }

                            WinRepairQ.partcnt = Lib.Instance.returnNumStringZero(WinRepairQ.partcnt);

                            if (WinRepairQ.repairdate != null && WinRepairQ.repairdate.Length == 8)
                            {
                                WinRepairQ.repairdate_CV = Lib.Instance.StrDateTimeBar(WinRepairQ.repairdate);
                            }

                            if (WinRepairQ.mcid.Equals("총계"))
                            {
                                
                                dgdTotal.Items.Add(WinRepairQ);
                            }
                            else
                            {
                                dgdRepairQ.Items.Add(WinRepairQ);
                            }
                        }

                        //tbkCount.Text = "▶ 검색결과 : " + i.ToString() + " 건";
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
        // 천단위 콤마, 소수점 버리기
        private string stringFormatN0(object obj)
        {
            return string.Format("{0:N0}", obj);
        }
        //미리보기(인쇄 하위버튼)
        private void menuSeeAhead_Click(object sender, RoutedEventArgs e)
        {
            if (dgdRepairQ.Items.Count < 1)
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
            if (dgdRepairQ.Items.Count < 1)
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

        #endregion
    }

    class Win_prd_Repair_Q_CodeView : BaseView
    {
        public override string ToString()
        {
            return (this.ReportAllProperties());
        }

        public string Num { get; set; }
        public string RepairGubun { get; set; }
        public string RepairGubun_CV { get; set; }
        public string mcid { get; set; }
        public string mcname { get; set; }
        public string repairdate { get; set; }
        public string repairdate_CV { get; set; }
        public string repairremark { get; set; }
        public string MCPartID { get; set; }
        public string MCPartName { get; set; }
        public string partcnt { get; set; }
        public string partremark { get; set; }
        public string RepairID { get; set; }
        public string repairsubseq { get; set; }
        public string MCCustom { get; set; }
        public string partprice { get; set; }
        public string reason { get; set; }
    }
}

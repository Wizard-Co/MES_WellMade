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
    /// Win_prd_PlanInputView_U_U.xaml에 대한 상호 작용 논리
    /// </summary>
    public partial class Win_prd_PlanInputView_U : UserControl
    {
        private Microsoft.Office.Interop.Excel.Application excelapp;
        private Microsoft.Office.Interop.Excel.Workbook workbook;
        private Microsoft.Office.Interop.Excel.Worksheet worksheet;
        private Microsoft.Office.Interop.Excel.Range workrange;
        private Microsoft.Office.Interop.Excel.Worksheet copysheet;
        private Microsoft.Office.Interop.Excel.Worksheet pastesheet;
        // 엑셀 활용 용도 (프린트)

        WizMes_WellMade.PopUp.NoticeMessage msg = new PopUp.NoticeMessage();
        //(기다림 알림 메시지창)
        System.Data.DataTable DT;
        Lib lib = new Lib();

        ObservableCollection<Win_prd_PlanInputView_U_CodeView> ovcPlanView = new ObservableCollection<Win_prd_PlanInputView_U_CodeView>();
        Win_prd_PlanInputView_U_CodeView WinPlanView = new Win_prd_PlanInputView_U_CodeView();
        Win_prd_PlanInputView_U_Sub_CodeView WinPlanSub = new Win_prd_PlanInputView_U_Sub_CodeView();

        int rowNum = 0;
        int count = 0;

        // 인쇄 미리보기 인지 아닌지
        private bool preview_click = false;

        public Win_prd_PlanInputView_U()
        {
            InitializeComponent();
        }

        private void UserControl_Loaded(object sender, RoutedEventArgs e)
        {
            Lib.Instance.UiLoading(sender);
            chkOrderDay.IsChecked = true;

            if (MainWindow.plInputFlag_SavePrint == true)
            {
                string InstID = "";

                if (MainWindow.plInput.Count > 0)
                {
                    foreach (string Key in MainWindow.plInput.Keys)
                    {
                        if (Key.ToUpper().Trim().Equals("DATE"))
                        {
                            dtpSDate.SelectedDate = (DateTime)MainWindow.plInput["Date"];
                            dtpEDate.SelectedDate = (DateTime)MainWindow.plInput["Date"];
                        }
                        else if (Key.ToUpper().Trim().Equals("INSTID"))
                        {
                            InstID = (string)MainWindow.plInput["InstID"];
                        }
                    }
                }

                FillGrid();

                for (int i = 0; i < dgdMain.Items.Count; i++)
                {
                    var Main = dgdMain.Items[i] as Win_prd_PlanInputView_U_CodeView;
                    if (Main != null)
                    {
                        if (Main.InstID.Trim().Equals(InstID))
                        {
                            dgdMain.SelectedIndex = i;
                            Main.IsCheck = true;
                            break;
                        }
                    }
                }

                MainWindow.plInputFlag_SavePrint = false;
                MainWindow.plInput.Clear();
            }
            else
            {
                btnThisMonth_Click(null, null);
            }
        }

        #region 라벨 체크박스 관련 이벤트

        //지시일자
        private void lblOrderDay_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            if (chkOrderDay.IsChecked == true) { chkOrderDay.IsChecked = false; }
            else { chkOrderDay.IsChecked = true; }
        }

        //지시일자
        private void chkOrderDay_Checked(object sender, RoutedEventArgs e)
        {
            if (dtpSDate != null && dtpEDate != null)
            {
                dtpSDate.IsEnabled = true;
                dtpEDate.IsEnabled = true;
            }
        }

        //지시일자
        private void chkOrderDay_Unchecked(object sender, RoutedEventArgs e)
        {
            dtpSDate.IsEnabled = false;
            dtpEDate.IsEnabled = false;
        }

        //금일
        private void btnToday_Click(object sender, RoutedEventArgs e)
        {
            dtpSDate.SelectedDate = DateTime.Today;
            dtpEDate.SelectedDate = DateTime.Today;
        }

        //금월
        private void btnThisMonth_Click(object sender, RoutedEventArgs e)
        {
            dtpSDate.SelectedDate = Lib.Instance.BringThisMonthDatetimeList()[0];
            dtpEDate.SelectedDate = Lib.Instance.BringThisMonthDatetimeList()[1];
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

        //관리번호
        private void lblOrderID_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            if (chkOrderID.IsChecked == true) { chkOrderID.IsChecked = false; }
            else { chkOrderID.IsChecked = true; }
        }

        //관리번호
        private void chkOrderID_Checked(object sender, RoutedEventArgs e)
        {
            txtOrderID.IsEnabled = true;
            txtOrderID.Focus();
        }

        //관리번호
        private void chkOrderID_Unchecked(object sender, RoutedEventArgs e)
        {
            txtOrderID.IsEnabled = false;
        }

        //OrderNo
        private void txtOrderID_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                if (e.Key == Key.Enter)
                {
                    MainWindow.pf.ReturnCode(txtOrderID, (int)Defind_CodeFind.DCF_ORDER, "");
                }
            }
        }

        //지시완료분 포함
        private void lblCompleted_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            if (chkCompleted.IsChecked == true) { chkCompleted.IsChecked = false; }
            else { chkCompleted.IsChecked = true; }
        }
        //지시마감분 포함
        private void lblClosed_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            if (chkClosed.IsChecked == true) { chkClosed.IsChecked = false; }
            else { chkClosed.IsChecked = true; }
        }

        private void rbnOrderNo_Click(object sender, RoutedEventArgs e)
        {
            RadioButton();
        }

        private void rbnOrderID_Click(object sender, RoutedEventArgs e)
        {
            RadioButton();
        }

        private void RadioButton()
        {
            if (rbnOrderNo.IsChecked == true)
            {
                tbkOrderID.Text = "OrderNo";
                dgdtpeOrderNo.Visibility = Visibility.Visible;
                dgdtpeOrderID.Visibility = Visibility.Hidden;
            }
            else if (rbnOrderID.IsChecked == true)
            {
                tbkOrderID.Text = "관리번호";
                dgdtpeOrderNo.Visibility = Visibility.Hidden;
                dgdtpeOrderID.Visibility = Visibility.Visible;
            }
        }

        #endregion

        //검색
        private void btnSearch_Click(object sender, RoutedEventArgs e)
        {
            //검색버튼 비활성화
            btnSearch.IsEnabled = false;

            Dispatcher.BeginInvoke(new Action(() =>

            {
                //로직
                using (Loading lw = new Loading(FillGrid))
                {
                    lw.ShowDialog();
                }

                if (dgdMain.Items.Count == 0)
                {
                    dgdSub.Items.Clear();

                    MessageBox.Show("조회된 데이터가 없습니다.");
                }
                else
                {
                    dgdMain.SelectedIndex = 0;
                }
            }), System.Windows.Threading.DispatcherPriority.Background);



            Dispatcher.BeginInvoke(new Action(() =>

            {
                btnSearch.IsEnabled = true;
                AllCheck.IsChecked = false;

            }), System.Windows.Threading.DispatcherPriority.Background);

        }

        //삭제
        private void btnDelete_Click(object sender, RoutedEventArgs e)
        {
            var InputView = dgdMain.SelectedItem as Win_prd_PlanInputView_U_CodeView;
            if (InputView != null)
            {
                if (MessageBox.Show("선택한 항목을 삭제하시겠습니까?", "삭제 전 확인", MessageBoxButton.YesNo) == MessageBoxResult.Yes)
                {
                    using (Loading ld = new Loading(beDelete))
                    {
                        ld.ShowDialog();
                    }

                    if (dgdMain.Items.Count == 0)
                    {
                        this.DataContext = null;
                        MessageBox.Show("조회된 데이터가 없습니다.");
                        return;
                    }
                }
            }
            else
                MessageBox.Show("삭제할 데이터를 선택해주세요.");
        }

        //닫기
        private void btnClose_Click(object sender, RoutedEventArgs e)
        {
            string strName = this.ToString();

            for (int i = 0; i < MainWindow.MainMdiContainer.Children.Count; i++)
            {
                if (strName.Equals((MainWindow.MainMdiContainer.Children[i] as MdiChild).Content.ToString()))
                {
                    (MainWindow.MainMdiContainer.Children[i] as MdiChild).Close();
                    break;
                }
            }
        }

        //작업지시목록
        private void btnOrderListPrint_Click(object sender, RoutedEventArgs e)
        {
            //PnlListPrint.IsOpen = true;

            // 체크된것 갯수 세기

            for (int i = 0; i < dgdMain.Items.Count; i++)
            {
                var Main = dgdMain.Items[i] as Win_prd_PlanInputView_U_CodeView;

                if (Main != null)
                {
                    if (Main.IsCheck == true)
                    {
                        count++;
                    }
                }
            }

            if (count == 0)
            {
                MessageBox.Show("인쇄할 대상을 선택해주세요");
                return;
            }

            FillPrintData_OnlyChecked();

            // 인쇄 메서드
            ContextMenu menu = btnOrderListPrint.ContextMenu;
            menu.StaysOpen = true;
            menu.IsOpen = true;
        }

        //
        private void btnUpdate_Click(object sender, RoutedEventArgs e)
        {

            WinPlanView = dgdMain.SelectedItem as Win_prd_PlanInputView_U_CodeView;

            if (WinPlanView != null)
            {
                dgdMain.IsHitTestVisible = false;
                Lib.Instance.UiButtonEnableChange_SCControl(this);
                chkMtrExceptYN.IsEnabled = true;
                chkOutWareExceptYN.IsEnabled = true;
                chkTheEnd.IsEnabled = true;
                rowNum = dgdMain.SelectedIndex;
                

            }
            else
            {
                MessageBox.Show("선택 사항이 없습니다. 선택을 먼저 해주십시오");
            }
        }

        //엑셀
        private void btnExcel_Click(object sender, RoutedEventArgs e)
        {
            DataTable dt = null;
            string Name = string.Empty;

            string[] lst = new string[4];
            lst[0] = "작업지시 종합";
            lst[1] = "작업지시 상세";
            lst[2] = dgdMain.Name;
            lst[3] = dgdSub.Name;

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
                else if (ExpExc.choice.Equals(dgdSub.Name))
                {
                    if (ExpExc.Check.Equals("Y"))
                        dt = Lib.Instance.DataGridToDTinHidden(dgdSub);
                    else
                        dt = Lib.Instance.DataGirdToDataTable(dgdSub);

                    Name = dgdSub.Name;

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

        //
        private void btnOrderPrint_Click(object sender, RoutedEventArgs e)
        {

        }

        #region 실조회 관련

        // 천마리 콤마, 소수점 버리기
        private string stringFormatN0(object obj)
        {
            return string.Format("{0:N0}", obj);
        }

        //조회
        private void FillGrid()
        {
            ovcPlanView.Clear();
            dgdTotal.Items.Clear();

            if (dgdMain.Items.Count > 0)
                dgdMain.Items.Clear();

            try
            {
                Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                sqlParameter.Clear();

                sqlParameter.Add("ChkDate", chkOrderDay.IsChecked == true ? 1 : 0);
                sqlParameter.Add("SDate", chkOrderDay.IsChecked == true && dtpSDate.SelectedDate != null ? dtpSDate.SelectedDate.Value.ToString("yyyyMMdd") : "");
                sqlParameter.Add("EDate", chkOrderDay.IsChecked == true && dtpEDate.SelectedDate != null ? dtpEDate.SelectedDate.Value.ToString("yyyyMMdd") : "");
                sqlParameter.Add("ChkCustomID", chkCustom.IsChecked == true ? 1 : 0);
                sqlParameter.Add("CustomID", chkCustom.IsChecked == true && txtCustom.Tag != null ? txtCustom.Tag.ToString() : "");

                sqlParameter.Add("ChkArticleID", chkArticle.IsChecked == true ? 1 : 0);
                sqlParameter.Add("ArticleID", chkArticle.IsChecked == true && txtArticle.Tag != null ? txtArticle.Tag.ToString() : "");
                sqlParameter.Add("ChkOrder", chkOrderID.IsChecked == true ? (rbnOrderID.IsChecked == true ? 1 : 2) : 0);
                sqlParameter.Add("Order", chkOrderID.IsChecked == true ? txtOrderID.Text : "");
                sqlParameter.Add("ChkCompleted", chkCompleted.IsChecked == true ? 1 : 0);
                sqlParameter.Add("ChkClosed", chkClosed.IsChecked == true ? 1 : 0);

                sqlParameter.Add("ChkBuyerArticleNo", CheckBoxBuyerArticleNoSearch.IsChecked == true ? 1 : 0);
                sqlParameter.Add("BuyerArticleNoID", CheckBoxBuyerArticleNoSearch.IsChecked == true ? (TextBoxBuyerArticleNoSearch.Tag != null ? TextBoxBuyerArticleNoSearch.Tag.ToString() : "") : "");


                DataSet ds = DataStore.Instance.ProcedureToDataSet("xp_prd_sPlanInput_WPF", sqlParameter, false);

                int i = 0;
                DataTable dt = ds.Tables[0];
                if (dt.Rows.Count > 0)
                {
                    DataRowCollection drc = dt.Rows;

                    foreach (DataRow dr in drc)
                    {
                        i++;
                        var WinPlanOrder = new Win_prd_PlanInputView_U_CodeView()
                        {
                            Num = i,
                            cls = dr["cls"].ToString(),
                            KCustom = dr["KCustom"].ToString(),
                            Article = dr["Article"].ToString(),
                            OrderID = dr["OrderID"].ToString(),
                            OrderNo = dr["OrderNo"].ToString(),
                            OrderQty = Convert.ToDouble(dr["OrderQty"]),

                            TotOrderinstqty = Convert.ToDouble(dr["TotOrderinstqty"]),
                            notOrderInstQty = Convert.ToDouble(dr["notOrderInstQty"]),
                            OrderInstQty = Convert.ToDouble(dr["OrderInstQty"]),
                            p1WorkQty = Convert.ToDouble(dr["p1WorkQty"]),
                            p1ProcessID = dr["p1ProcessID"].ToString(),

                            p1ProcessName = dr["p1ProcessName"].ToString(),
                            InspectQty = Convert.ToDouble(dr["InspectQty"]),
                            OutQty = Convert.ToDouble(dr["OutQty"]),
                            PatternID = dr["PatternID"].ToString(),
                            ArticleGrpID = dr["ArticleGrpID"].ToString(),
                            ArticleGrp = dr["ArticleGrp"].ToString(),

                            BuyerModel = dr["BuyerModel"].ToString(),
                            BuyerModelID = dr["BuyerModelID"].ToString(),
                            BuyerArticleNo = dr["BuyerArticleNo"].ToString(),
                            Remark = dr["Remark"].ToString(),
                            Completed = dr["Completed"].ToString(),

                            ArticleID = dr["ArticleID"].ToString(),
                            InstID = dr["InstID"].ToString(),
                            InstDate = DatePickerFormat(dr["InstDate"].ToString()),
                            ProcPattern = dr["ProcPattern"].ToString(),
                            MtrExceptYN = dr["MtrExceptYN"].ToString(),

                            OutwareExceptYN = dr["OutwareExceptYN"].ToString(),
                            LotID = dr["LotID"].ToString(),
                            Closed = dr["Closed"].ToString(),
                            Progress = dr["Progress"].ToString() + "%",               //진척률
                            InstSeq = dr["InstSeq"].ToString()           

                        };

                        //if (WinPlanOrder.MtrExceptYN.Equals("Y"))
                        //{
                        //    chkMtrExceptYN.IsChecked = true;
                        //}
                        //else
                        //{
                        //    chkMtrExceptYN.IsChecked = false;
                        //}

                        //if (WinPlanOrder.OutwareExceptYN.Equals("Y"))
                        //{
                        //    chkOutWareExceptYN.IsChecked = true;
                        //}
                        //else
                        //{
                        //    chkOutWareExceptYN.IsChecked = false;
                        //}

                        if (!WinPlanOrder.cls.Trim().Equals("9"))
                        {
                            dgdMain.Items.Add(WinPlanOrder);
                        }
                        // 총계
                        else
                        {
                            WinPlanOrder.KCustom = "";
                            dgdTotal.Items.Add(WinPlanOrder);
                        }
                    }

                    //txtSumOrderQty.Text = string.Format("{0:N0}", sumOrder);
                    //txtSumWorkQty.Text = string.Format("{0:N0}", sumWork);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("오류 발생, 오류 내용 : " + ex.ToString());
            }
        }

        //선택변경
        private void dgdMain_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            WinPlanView = dgdMain.SelectedItem as Win_prd_PlanInputView_U_CodeView;

            if (WinPlanView != null)
            {

                FillGridSub(WinPlanView.InstID);

                if (WinPlanView.MtrExceptYN.Equals("Y"))
                {
                    chkMtrExceptYN.IsChecked = true;
                }
                else
                {
                    chkMtrExceptYN.IsChecked = false;
                }

                if (WinPlanView.OutwareExceptYN.Equals("Y"))
                {
                    chkOutWareExceptYN.IsChecked = true;
                }
                else
                {
                    chkOutWareExceptYN.IsChecked = false;
                }

                if (WinPlanView.Closed.Trim().Equals("*"))
                {
                    chkTheEnd.IsChecked = true;
                }
                else
                {
                    chkTheEnd.IsChecked = false;
                }
            }
        }

        //하단 그리드
        private void FillGridSub(string strInstID)
        {
            if (dgdSub != null)
            {
                dgdSub.Items.Clear();
            }

            try
            {

                Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                sqlParameter.Clear();
                sqlParameter.Add("nChkInstID", 1);
                sqlParameter.Add("sInstID", strInstID);
                DataSet ds = DataStore.Instance.ProcedureToDataSet("xp_prd_sPlanInputDet_WPF", sqlParameter, false);

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
                            var WinPlanSub = new Win_prd_PlanInputView_U_Sub_CodeView()
                            {
                                ProcessID = dr["ProcessID"].ToString(),
                                Process = dr["Process"].ToString(),
                                OrderNo = dr["OrderNo"].ToString(),
                                InstID = dr["InstID"].ToString(),
                                InstDetSeq = dr["InstDetSeq"].ToString(),

                                InstQty = stringFormatN2(dr["InstQty"]),
                                StartDate = dr["StartDate"].ToString(),
                                StartDate_CV = DatePickerFormat(dr["StartDate"].ToString()),
                                EndDate = dr["EndDate"].ToString(),
                                EndDate_CV = DatePickerFormat(dr["EndDate"].ToString()),
                                InstRemark = dr["InstRemark"].ToString(),
                                ArticleID = dr["ArticleID"].ToString(),
                                BuyerArticleNo = dr["BuyerArticleNo"].ToString(),

                                Article = dr["Article"].ToString(),
                                WorkQty = stringFormatN0(dr["WorkQty"]),
                                lotID = dr["lotID"].ToString(),
                                MachineID = dr["MachineID"].ToString(),
                                MachineNo = dr["MachineNo"].ToString(),

                                Machine = dr["Machine"].ToString(),
                                FirstProcessLotID = dr["FirstProcessLotID"].ToString(),

                                ChildArticleID = dr["ChildArticleID"].ToString(),
                                ChildBuyerArticleNo = dr["ChildBuyerArticleNo"].ToString(),
                                MtrExceptYN = dr["MtrExceptYN"].ToString(),
                                FirstInFirstOutYN = dr["FirstInFirstOutYN"].ToString(),
                            };

                            dgdSub.Items.Add(WinPlanSub);
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



        //저장
        private void btnSave_Click(object sender, RoutedEventArgs e)
        {
            using (Loading lw = new Loading(beSave))
            {
                lw.ShowDialog();
            }
        }

        private void beSave()
        {
            if (SaveData())
            {
                //numCompare = 0;
                chkMtrExceptYN.IsEnabled = false;
                chkOutWareExceptYN.IsEnabled = false;
                chkTheEnd.IsChecked = false;
                chkTheEnd.IsEnabled = false;

                //dgdMain.IsEnabled = true;
                dgdMain.IsHitTestVisible = true;
                Lib.Instance.UiButtonEnableChange_IUControl(this);

                dgdMain.Items.Clear();
                dgdSub.Items.Clear();

                FillGrid();

                if (dgdMain.Items.Count == 0)
                {
                    dgdSub.Items.Clear();

                    MessageBox.Show("조회된 데이터가 없습니다.");
                    return;
                }

                dgdMain.SelectedIndex = rowNum;
            }
        }

        private void beDelete()
        {
            btnDelete.IsEnabled = false;

            Dispatcher.BeginInvoke(new Action(() =>
            {
                foreach (Win_prd_PlanInputView_U_CodeView PlanView in ovcPlanView)
                {
                    if (PlanView != null)
                        DeleteData(PlanView.InstID);
                }

                rowNum = 0;
                FillGrid();

            }), System.Windows.Threading.DispatcherPriority.Background);

            btnDelete.IsEnabled = true;
        }

        //취소
        private void btnCancel_Click(object sender, RoutedEventArgs e)
        {
            chkMtrExceptYN.IsEnabled = false;
            chkOutWareExceptYN.IsEnabled = false;

            dgdMain.IsHitTestVisible = true;
            Lib.Instance.UiButtonEnableChange_IUControl(this);

            using (Loading lw = new Loading(FillGrid))
            {
                lw.ShowDialog();
            }

            if (dgdMain.Items.Count == 0)
            {
                dgdSub.Items.Clear();

                MessageBox.Show("조회된 데이터가 없습니다.");
                return;
            }

            dgdMain.SelectedIndex = rowNum;
        }

        //실저장
        private bool SaveData()
        {
            bool flag = true;
            List<Procedure> Prolist = new List<Procedure>();
            List<Dictionary<string, object>> ListParameter = new List<Dictionary<string, object>>();

            try
            {
                Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                sqlParameter.Clear();
                sqlParameter.Add("InstID", WinPlanView.InstID);
                sqlParameter.Add("MtrExceptYN", chkMtrExceptYN.IsChecked == true ? "Y" : "N");
                sqlParameter.Add("OutwareExceptYN", chkOutWareExceptYN.IsChecked == true ? "Y" : "N");
                sqlParameter.Add("OrderInstQty", WinPlanView.OrderInstQty);
                sqlParameter.Add("UpdateUserID", MainWindow.CurrentUser);
                sqlParameter.Add("Remark", WinPlanView.Remark);

                Procedure pro1 = new Procedure();
                pro1.Name = "xp_PlanInput_uPlanInput";
                pro1.OutputUseYN = "N";
                pro1.OutputName = "InstID";
                pro1.OutputLength = "10";

                Prolist.Add(pro1);
                ListParameter.Add(sqlParameter);

                for (int i = 0; i < dgdSub.Items.Count; i++)
                {
                    WinPlanSub = dgdSub.Items[i] as Win_prd_PlanInputView_U_Sub_CodeView;
                    sqlParameter = new Dictionary<string, object>();
                    sqlParameter.Clear();
                    sqlParameter.Add("InstID", WinPlanView.InstID);
                    sqlParameter.Add("nInstDetSeq", i + 1);
                    sqlParameter.Add("sStartDate", WinPlanSub.StartDate);
                    sqlParameter.Add("sEndDate", WinPlanSub.EndDate);
                    sqlParameter.Add("nInstQty", ConvertDouble(WinPlanSub.InstQty.Replace(",", "")));
                    sqlParameter.Add("sInstSubRemark", WinPlanSub.InstRemark);
                    sqlParameter.Add("MachineID", WinPlanSub.MachineID);
                    sqlParameter.Add("TheEnd", chkTheEnd.IsChecked == true ? 1 : 0);
                    sqlParameter.Add("MtrExceptYN", WinPlanSub.MtrExceptYN == null ? "" : WinPlanSub.MtrExceptYN);
                    sqlParameter.Add("FirstInFirstOut", WinPlanSub.FirstInFirstOutYN == null ? "" : WinPlanSub.FirstInFirstOutYN);
                    sqlParameter.Add("UpdateUserID", MainWindow.CurrentUser);

                    Procedure pro2 = new Procedure();
                    pro2.Name = "xp_PlanInput_uPlanInputSub_WPF";
                    pro2.OutputUseYN = "N";
                    pro2.OutputName = "InstID";
                    pro2.OutputLength = "10";

                    Prolist.Add(pro2);
                    ListParameter.Add(sqlParameter);
                }

                string[] Confirm = new string[2];
                Confirm = DataStore.Instance.ExecuteAllProcedureOutputNew(Prolist, ListParameter);
                if (Confirm[0] == "success")
                {
                    //MessageBox.Show("성공");
                    flag = true;
                    return flag;
                }
                else
                {
                    MessageBox.Show("[저장실패]\r\n" + Confirm[1].ToString());
                    flag = false;
                    return flag;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally
            {
                DataStore.Instance.CloseConnection();
            }

            return flag;
        }

        //삭제 
        private bool DeleteData(string InstID)
        {
            bool flag = true;

            try
            {
                Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                sqlParameter.Clear();
                sqlParameter.Add("InstID", InstID);
                sqlParameter.Add("OutMessage", "");

                //string[] result = DataStore.Instance.ExecuteProcedure_NewLog("xp_PlanInput_dPlanInput", sqlParameter, "D");
                DataSet ds = DataStore.Instance.ProcedureToDataSet("xp_PlanInput_dPlanInput", sqlParameter, false);

                if (ds != null && ds.Tables.Count > 0)
                {
                    DataTable dt = ds.Tables[0];
                    int i = 0;

                    if (dt.Rows.Count > 0)
                    {
                        // 문제가 있을 경우에는, OutMessage 하나만 가져오기 때문에, 
                        // 컬럼이 1개인 경우에 해당 메시지를 출력
                        if (dt.Columns.Count == 1)
                        {
                            MessageBox.Show(dt.Rows[0].ItemArray[0].ToString());
                            flag = false;
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

            return flag;
        }


        #region 하단 입력을 위한 이벤트

        #region Content 부분 - 데이터 그리드 키 이벤트

        // 2019.08.27 PreviewKeyDown 는 key 다운과 같은것 같음
        private void DataGird_PreviewKeyDown(object sender, KeyEventArgs e)
        {
            try
            {
                if (e.Key == Key.Down || e.Key == Key.Up || e.Key == Key.Left || e.Key == Key.Right)
                {
                    DataGird_KeyDown(sender, e);
                }
            }
            catch (Exception ex)
            {

            }
        }

        // KeyDown 이벤트
        private void DataGird_KeyDown(object sender, KeyEventArgs e)
        {
            int currRow = dgdSub.Items.IndexOf(dgdSub.CurrentItem);
            int currCol = dgdSub.Columns.IndexOf(dgdSub.CurrentCell.Column);
            int startCol = 6;
            int endCol = 14;

            if (e.Key == Key.Enter)
            {
                e.Handled = true;
                (sender as DataGridCell).IsEditing = false;

                // 마지막 열, 마지막 행 아님
                if (endCol == currCol && dgdSub.Items.Count - 1 > currRow)
                {
                    dgdSub.SelectedIndex = currRow + 1; // 이건 한줄 파란색으로 활성화 된 걸 조정하는 것입니다.
                    dgdSub.CurrentCell = new DataGridCellInfo(dgdSub.Items[currRow + 1], dgdSub.Columns[startCol]);

                } // 마지막 열 아님
                else if (endCol > currCol && dgdSub.Items.Count - 1 >= currRow)
                {
                    dgdSub.CurrentCell = new DataGridCellInfo(dgdSub.Items[currRow], dgdSub.Columns[currCol + 1]);

                } // 마지막 열, 마지막 행
                else if (endCol == currCol && dgdSub.Items.Count - 1 == currRow)
                {

                }
                else
                {
                    MessageBox.Show("나머지가 있나..");
                }
            }
            else if (e.Key == Key.Down)
            {
                e.Handled = true;
                (sender as DataGridCell).IsEditing = false;

                // 마지막 행 아님
                if (dgdSub.Items.Count - 1 > currRow)
                {
                    dgdSub.SelectedIndex = currRow + 1;
                    dgdSub.CurrentCell = new DataGridCellInfo(dgdSub.Items[currRow + 1], dgdSub.Columns[currCol]);
                } // 마지막 행일때
                else if (dgdSub.Items.Count - 1 == currRow)
                {
                    if (endCol > currCol) // 마지막 열이 아닌 경우, 열을 오른쪽으로 이동
                    {
                        //dgdMain.SelectedIndex = 0;
                        dgdSub.CurrentCell = new DataGridCellInfo(dgdSub.Items[currRow], dgdSub.Columns[currCol + 1]);
                    }
                    else
                    {
                        //btnSave.Focus();
                    }
                }
            }
            else if (e.Key == Key.Up)
            {
                e.Handled = true;
                (sender as DataGridCell).IsEditing = false;

                // 첫행 아님
                if (currRow > 0)
                {
                    dgdSub.SelectedIndex = currRow - 1;
                    dgdSub.CurrentCell = new DataGridCellInfo(dgdSub.Items[currRow - 1], dgdSub.Columns[currCol]);
                } // 첫 행
                else if (dgdSub.Items.Count - 1 == currRow)
                {
                    if (0 < currCol) // 첫 열이 아닌 경우, 열을 왼쪽으로 이동
                    {
                        //dgdMain.SelectedIndex = 0;
                        dgdSub.CurrentCell = new DataGridCellInfo(dgdSub.Items[currRow], dgdSub.Columns[currCol - 1]);
                    }
                    else
                    {
                        //btnSave.Focus();
                    }
                }
            }
            else if (e.Key == Key.Left)
            {
                e.Handled = true;
                (sender as DataGridCell).IsEditing = false;

                if (startCol < currCol)
                {
                    dgdSub.CurrentCell = new DataGridCellInfo(dgdSub.Items[currRow], dgdSub.Columns[currCol - 1]);
                }
                else if (startCol == currCol)
                {
                    if (0 < currRow)
                    {
                        dgdSub.SelectedIndex = currRow - 1;
                        dgdSub.CurrentCell = new DataGridCellInfo(dgdSub.Items[currRow - 1], dgdSub.Columns[endCol]);
                    }
                    else
                    {
                        //btnSave.Focus();
                    }
                }
            }
            else if (e.Key == Key.Right)
            {
                e.Handled = true;
                (sender as DataGridCell).IsEditing = false;

                if (endCol > currCol)
                {

                    dgdSub.CurrentCell = new DataGridCellInfo(dgdSub.Items[currRow], dgdSub.Columns[currCol + 1]);
                }
                else if (endCol == currCol)
                {
                    if (dgdSub.Items.Count - 1 > currRow)
                    {
                        dgdSub.SelectedIndex = currRow + 1;
                        dgdSub.CurrentCell = new DataGridCellInfo(dgdSub.Items[currRow + 1], dgdSub.Columns[startCol]);
                    }
                    else
                    {
                        //btnSave.Focus();
                    }
                }
            }

        }
        // KeyUp 이벤트
        private void DataGridIn_TextFocus(object sender, KeyEventArgs e)
        {
            // 엔터 → 포커스 = true → cell != null → 해당 텍스트박스가 null이 아니라면 
            // → 해당 텍스트박스가 포커스가 안되있음 SelectAll() or 포커스
            Lib.Instance.DataGridINTextBoxFocus(sender, e);
        }
        // GotFocus 이벤트
        private void DataGridCell_GotFocus(object sender, RoutedEventArgs e)
        {
            if (btnSave.Visibility == Visibility.Visible)
            {
                int currCol = dgdSub.Columns.IndexOf(dgdSub.CurrentCell.Column);

                if ((currCol >= 6 && currCol < 14)
                    || currCol == 15)
                {
                    DataGridCell cell = sender as DataGridCell;
                    cell.IsEditing = true;
                }
            }
        }
        // 2019.08.27 MouseUp 이벤트
        private void DataGridCell_MouseUp(object sender, MouseButtonEventArgs e)
        {
            Lib.Instance.DataGridINTextBoxFocusByMouseUP(sender, e);
        }
            #endregion // Content 부분 - 데이터 그리드 키 이벤트

            //enter 시 Control 포커스
            private void TextBoxFocusInDataGrid(object sender, KeyEventArgs e)
        {
            Lib.Instance.DataGridINControlFocus(sender, e);
        }

        //마우스 클릭시 control 포커스
        private void TextBoxFocusInDataGrid_MouseUp(object sender, MouseButtonEventArgs e)
        {
            Lib.Instance.DataGridINBothByMouseUP(sender, e);
        }

        //지시수량
        private void dgdtpetxtInstQty_TextChanged(object sender, TextChangedEventArgs e)
        {
            if (btnSave.Visibility == Visibility.Visible)
            {
                WinPlanSub = dgdSub.CurrentItem as Win_prd_PlanInputView_U_Sub_CodeView;

                if (WinPlanSub != null)
                {
                    TextBox tb1 = sender as TextBox;

                    if (Lib.Instance.IsNumOrAnother(tb1.Text))
                    {
                        WinPlanSub.InstQty = Lib.Instance.returnNumStringZero(tb1.Text);
                        tb1.SelectionStart = tb1.Text.Length;
                        sender = tb1;
                    }
                }
            }
        }

        private void dgdtxtOrderInstQty_PreviewTextInput(object sender, TextCompositionEventArgs e)
        {
            Lib.Instance.CheckIsNumeric((TextBox)sender, e);
        }

        //지시수량
        private void dgdtpetxtInstQty_PreviewTextInput(object sender, TextCompositionEventArgs e)
        {
            Lib.Instance.CheckIsNumeric((TextBox)sender, e);
        }

        //시작일
        private void dgdtpedtpSDate_SelectedDateChanged(object sender, SelectionChangedEventArgs e)
        {
            //Win_prd_PlanInputView_U_Sub_CodeView;
            DatePicker dtpSender = sender as DatePicker;
            var pInput = dtpSender.DataContext as Win_prd_PlanInputView_U_Sub_CodeView;

            if (pInput != null
                && dtpSender.SelectedDate != null)
            {
                pInput.StartDate_CV = dtpSender.SelectedDate.Value.ToString("yyyy-MM-dd");
                pInput.StartDate = pInput.StartDate_CV.Replace("-", "");
            }
        }

        //시작일
        private void dgdtpedtpSDate_CalendarClosed(object sender, RoutedEventArgs e)
        {
            WinPlanSub = dgdSub.CurrentItem as Win_prd_PlanInputView_U_Sub_CodeView;
            int rowCount = dgdSub.Items.IndexOf(dgdSub.CurrentItem);
            int colCount = dgdSub.Columns.IndexOf(dgdSub.CurrentCell.Column);
            DatePicker dtpSDate = (DatePicker)sender;

            if (WinPlanSub == null)
            {
                MessageBox.Show("행 없다!");
                return;
            }

            if (dtpSDate.SelectedDate != null)
            {
                WinPlanSub.StartDate = dtpSDate.SelectedDate != null ? dtpSDate.SelectedDate.Value.ToString("yyyyMMdd") : "";
                sender = dtpSDate;
            }
            DataGridCell cell = Lib.Instance.GetParent<DataGridCell>(sender as DatePicker);
            if (cell != null)
            {
                cell.IsEditing = false;
                dgdSub.CurrentCell = new DataGridCellInfo(dgdSub.Items[rowCount], dgdSub.Columns[colCount + 1]);
            }
        }

        //종료일
        private void dgdtpedtpEDate_SelectedDateChanged(object sender, SelectionChangedEventArgs e)
        {
            //Win_prd_PlanInputView_U_Sub_CodeView;
            DatePicker dtpSender = sender as DatePicker;
            var pInput = dtpSender.DataContext as Win_prd_PlanInputView_U_Sub_CodeView;

            if (pInput != null
                && dtpSender.SelectedDate != null)
            {
                pInput.EndDate_CV = dtpSender.SelectedDate != null ? dtpSender.SelectedDate.Value.ToString("yyyyMMdd") : "";
                pInput.EndDate = pInput.EndDate_CV.Replace("-", "");
            }
        }

        //종료일
        private void dgdtpedtpEDate_CalendarClosed(object sender, RoutedEventArgs e)
        {
            WinPlanSub = dgdSub.CurrentItem as Win_prd_PlanInputView_U_Sub_CodeView;
            int rowCount = dgdSub.Items.IndexOf(dgdSub.CurrentItem);
            int colCount = dgdSub.Columns.IndexOf(dgdSub.CurrentCell.Column);
            DatePicker dtpEDate = (DatePicker)sender;

            if (WinPlanSub == null)
            {
                MessageBox.Show("행 없다!");
                return;
            }

            if (dtpEDate.SelectedDate != null)
            {
                WinPlanSub.EndDate = dtpEDate.SelectedDate != null ? dtpEDate.SelectedDate.Value.ToString("yyyyMMdd") : "";
                sender = dtpEDate;
            }
            DataGridCell cell = Lib.Instance.GetParent<DataGridCell>(sender as DatePicker);
            if (cell != null)
            {
                cell.IsEditing = false;
                dgdSub.CurrentCell = new DataGridCellInfo(dgdSub.Items[rowCount], dgdSub.Columns[colCount + 1]);
            }
        }

        //지시사항
        private void dgdtpetxtInstRemark_TextChanged(object sender, TextChangedEventArgs e)
        {
            if (btnSave.Visibility == Visibility.Visible)
            {
                WinPlanSub = dgdSub.CurrentItem as Win_prd_PlanInputView_U_Sub_CodeView;

                if (WinPlanSub != null)
                {
                    TextBox tb1 = sender as TextBox;

                    if (tb1 != null)
                    {
                        WinPlanSub.InstRemark = tb1.Text;
                    }

                    sender = tb1;
                }
            }
        }

        //생산수량
        private void dgdtpetxtWorkQty_TextChanged(object sender, TextChangedEventArgs e)
        {
            if (btnSave.Visibility == Visibility.Visible)
            {
                WinPlanSub = dgdSub.CurrentItem as Win_prd_PlanInputView_U_Sub_CodeView;

                if (WinPlanView != null)
                {
                    TextBox tb1 = sender as TextBox;

                    if (Lib.Instance.IsNumOrAnother(tb1.Text))
                    {
                        WinPlanSub.WorkQty = Lib.Instance.returnNumStringZero(tb1.Text);
                        tb1.SelectionStart = tb1.Text.Length;
                        sender = tb1;
                    }
                }
            }
        }

        //생산수량
        private void dgdtpetxtWorkQty_PreviewTextInput(object sender, TextCompositionEventArgs e)
        {
            Lib.Instance.CheckIsNumeric((TextBox)sender, e);
        }

        //호기
        private void dgdtpetxtMachine_KeyDown(object sender, KeyEventArgs e)
        {
            if (btnSave.Visibility == Visibility.Visible)
            {
                WinPlanSub = dgdSub.CurrentItem as Win_prd_PlanInputView_U_Sub_CodeView;

                if (e.Key == Key.Enter)
                {
                    if (WinPlanView != null)
                    {
                        TextBox tb1 = sender as TextBox;

                        MainWindow.pf.ReturnCode(tb1, 79, WinPlanSub.ProcessID);

                        if (tb1.Tag != null)
                        {
                            WinPlanSub.Machine = tb1.Text;
                            WinPlanSub.MachineID = tb1.Tag.ToString();
                        }

                        sender = tb1;
                    }
                }
            }
        }

        private void dgdtpetxtMachine_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            if (btnSave.Visibility == Visibility.Visible)
            {
                WinPlanSub = dgdSub.CurrentItem as Win_prd_PlanInputView_U_Sub_CodeView;

                if (WinPlanView != null)
                {
                    TextBox tb1 = sender as TextBox;

                    MainWindow.pf.ReturnCode(tb1, 66, WinPlanSub.ProcessID);

                    if (tb1.Tag != null)
                    {
                        WinPlanSub.Machine = tb1.Text;
                        WinPlanSub.MachineID = tb1.Tag.ToString();
                    }

                    sender = tb1;
                }
            }
        }

        //LOTID
        private void dgdtpetxtFirstProcessLotID_TextChanged(object sender, TextChangedEventArgs e)
        {
            if (btnSave.Visibility == Visibility.Visible)
            {
                WinPlanSub = dgdSub.CurrentItem as Win_prd_PlanInputView_U_Sub_CodeView;

                if (WinPlanView != null)
                {
                    TextBox tb1 = sender as TextBox;
                    WinPlanSub.lotID = tb1.Text;
                    sender = tb1;
                }
            }
        }


        #endregion

        int numchkPrWork = 0;
        int numchkPrMachine = 0;

        int numWorkSelect = 0;
        int numMachineSelecte = 0;

        //금일 버튼
        private void btnPrToday_Click(object sender, RoutedEventArgs e)
        {
            prdtpSDate.SelectedDate = DateTime.Today;
            prdtpEDate.SelectedDate = DateTime.Today;
        }

        //금월 버튼
        private void btnPrThisMonth_Click(object sender, RoutedEventArgs e)
        {
            prdtpSDate.SelectedDate = DateTime.Now.AddDays(-(DateTime.Today.Day - 1));
            prdtpEDate.SelectedDate = DateTime.Today;
        }

        //체크박스 클릭(체크박스가 작아졌을때 쉬운 클릭을 위해
        private void lblPrWork_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            if (chkPrWork.IsChecked == true) { chkPrWork.IsChecked = false; }
            else { chkPrWork.IsChecked = true; }
        }

        //공정 체크박스
        private void chkPrWork_Checked(object sender, RoutedEventArgs e)
        {
            cboPrWork.IsEnabled = true;
            numchkPrWork = 1;
            dgdPrintResearch();
        }

        //공정 체크박스
        private void chkPrWork_Unchecked(object sender, RoutedEventArgs e)
        {
            cboPrWork.IsEnabled = false;
            numchkPrWork = 0;
            dgdPrintResearch();
        }

        //체크박스 클릭(체크박스가 작아졌을때 쉬운 클릭을 위해
        private void lblPrMachine_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            if (chkPrMachine.IsChecked == true) { chkPrMachine.IsChecked = false; }
            else { chkPrMachine.IsChecked = true; }
        }

        //호기 체크박스
        private void chkPrMachine_Checked(object sender, RoutedEventArgs e)
        {
            cboPrMachine.IsEnabled = true;
            numchkPrMachine = 1;
            dgdPrintResearch();
        }

        //호기 체크박스
        private void chkPrMachine_Unchecked(object sender, RoutedEventArgs e)
        {
            cboPrMachine.IsEnabled = false;
            numchkPrMachine = 0;
            dgdPrintResearch();
        }

        private void PnlListPrint_Loaded(object sender, RoutedEventArgs e)
        {
            prdtpSDate.SelectedDate = DateTime.Today;
            prdtpEDate.SelectedDate = DateTime.Today;

            SetWorkComboBox();
            cboPrWork.SelectedIndex = 0;
            //SetMachineComboBox(cboPrWork.SelectedValue.ToString());
            cboPrMachine.SelectedIndex = 0;
            //Lib.Instance.UiLoading(sender);

            chkCompleted.IsChecked = true;
        }

        private void PnlListPrint_Opened(object sender, EventArgs e)
        {
            FillPrintData_OnlyChecked();

            //PrintWork(true);
        }

        //확인버튼
        private void btnPrOK_Click(object sender, RoutedEventArgs e)
        {
            ContextMenu menu = btnPrOK.ContextMenu;
            menu.StaysOpen = true;
            menu.IsOpen = true;
        }

        //확인버튼 내부 미리보기
        private void menuSeeAhead_Click(object sender, RoutedEventArgs e)
        {
            //msg.Show();
            //msg.Topmost = true;
            //msg.Refresh();

            //PrintWork();

            preview_click = true;
            using (Loading lw = new Loading("excel", PrintWork))
            {
                lw.ShowDialog();
            }
        }

        //확인버튼 내부 바로인쇄
        private void menuRightPrint_Click(object sender, RoutedEventArgs e)
        {
            //msg.Show();
            //msg.Topmost = true;
            //msg.Refresh();

            //PrintWork();

            preview_click = false;
            using (Loading lw = new Loading("excel", PrintWork))
            {
                lw.ShowDialog();
            }
        }

        //확인버튼 내부 닫기
        private void menuClose_Click(object sender, RoutedEventArgs e)
        {
            ContextMenu menu = btnPrOK.ContextMenu;
            menu.StaysOpen = false;
            menu.IsOpen = false;
        }

        //true 일때 미리보기, false일때 바로인쇄
        private void PrintWork()
        {
            try
            {
                excelapp = new Microsoft.Office.Interop.Excel.Application();

                //string MyBookPath = "C:/Users/Administrator/Desktop/tmp_작업지시목록.xls";
                string MyBookPath = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetEntryAssembly().Location) + "\\Report\\tmp_작업지시목록.xls";
                //MyBookPath = MyBookPath.Substring(0, MyBookPath.LastIndexOf("\\")) + "\\tmp_작업지시목록.xls";
                workbook = excelapp.Workbooks.Add(MyBookPath);
                worksheet = workbook.Sheets["Form"];

                int Page = 0;
                int DataCount = 0;
                int copyLine = 0;

                copysheet = workbook.Sheets["Form"];
                pastesheet = workbook.Sheets["Print"];

                Lib lib = new Lib();

                string prNum = string.Empty;
                string prProcess = string.Empty;
                string prProcessID = string.Empty;
                string prKCustom = string.Empty;
                string prArticle = string.Empty;
                string prArticleID = string.Empty;
                string prArticleNo = string.Empty;
                string prBuyerModel = string.Empty;
                string prSDate = string.Empty;
                string prEDate = string.Empty;
                string prlotID = string.Empty;
                string prMachineNo = string.Empty;
                string prInstQty = string.Empty;

                if (dgdPrintForChecked.Items.Count == 1)
                {
                    while (dgdPrintForChecked.Items.Count > DataCount)
                    {
                        Page++;
                        if (Page != 1) { DataCount++; }
                        copyLine = (Page - 1) * 45;
                        copysheet.Select();
                        copysheet.UsedRange.Copy();
                        pastesheet.Select();
                        workrange = pastesheet.Cells[copyLine + 1, 1];
                        //workrange.UseStandardHeight = copysheet.StandardHeight;
                        workrange.Select();
                        pastesheet.Paste();

                        int j = 0;
                        for (int i = DataCount; i < dgdPrintForChecked.Items.Count; i++)
                        {
                            if (j == 40) { break; }
                            int insertline = copyLine + 6 + j;

                            var Sub = dgdPrintForChecked.Items[i] as Win_prd_Print_Plan_InputDet_CodeView;

                            prNum = (i + 1).ToString();
                            prSDate = DatePickerFormatSlash(Sub.prStartDate);
                            prEDate = DatePickerFormatSlash(Sub.prEndDate);
                            prProcessID = Sub.prProcessID;
                            prProcess = Sub.prProcess;
                            prMachineNo = Sub.prMachineNo;
                            prArticleID = Sub.prArticleID;
                            prArticle = Sub.prArticle;
                            prArticleNo = Sub.prBuyerArticleNo;
                            prBuyerModel = Sub.prBuyerModel;
                            prlotID = Sub.prlotID;
                            prInstQty = Sub.prInstQty;


                            workrange = pastesheet.get_Range("A" + (insertline - 1), "B" + insertline);    //순번
                            workrange.Value2 = prNum;
                            workrange.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                            workrange.Font.Size = 11;

                            workrange = pastesheet.get_Range("C" + (insertline - 1), "E" + insertline);    //시작일~종료일
                            workrange.Value2 = prSDate + "~" + "\n\r" + prEDate;
                            workrange.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                            workrange.Font.Size = 11;

                            workrange = pastesheet.get_Range("F" + (insertline - 1), "J" + insertline);    //공정 /호기
                                                                                                           //workrange.Value2 = prProcess + "\n\r" + prMachineNo;
                            workrange.Value2 = prProcess;
                            workrange.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                            workrange.Font.Size = 11;

                            workrange = pastesheet.get_Range("K" + (insertline - 1), "Q" + insertline);    //품명코드 /품명
                            workrange.Value2 = prArticleNo;
                            workrange.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                            workrange.Font.Size = 11;

                            workrange = pastesheet.get_Range("R" + (insertline - 1), "S" + insertline);    //지시수량
                            workrange.Value2 = prInstQty;
                            workrange.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                            workrange.Font.Size = 9;

                            workrange = pastesheet.get_Range("T" + (insertline - 1), "AB" + (insertline - 1));    //LotID (font변경)
                            workrange.Value2 = "*" + prlotID + "*";
                            workrange.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                            workrange.Font.Size = 18;
                            workrange.Font.FontStyle = "Code39(2:3)";

                            workrange = pastesheet.get_Range("T" + insertline, "AB" + insertline);    //LotID
                            workrange.Value2 = prlotID;
                            workrange.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                            workrange.Font.Size = 8;

                            DataCount = i;
                            j += 2;
                        }
                        DataCount++;
                    }
                }
                else
                {
                    while (dgdPrintForChecked.Items.Count - 1 > DataCount)
                    {
                        Page++;
                        if (Page != 1) { DataCount++; }
                        copyLine = (Page - 1) * 45;
                        copysheet.Select();
                        //copysheet.UsedRange.Copy();

                        copysheet.UsedRange.EntireRow.Copy();
                        pastesheet.Select();
                        workrange = pastesheet.Cells[copyLine + 1, 1];
                        workrange.Select();
                        pastesheet.Paste();

                        int j = 0;
                        for (int i = DataCount; i < dgdPrintForChecked.Items.Count; i++)
                        {
                            if (j == 40) { break; }
                            int insertline = copyLine + 6 + j;

                            var Sub = dgdPrintForChecked.Items[i] as Win_prd_Print_Plan_InputDet_CodeView;

                            prNum = (i + 1).ToString();
                            prSDate = DatePickerFormatSlash(Sub.prStartDate);
                            prEDate = DatePickerFormatSlash(Sub.prEndDate);
                            prProcessID = Sub.prProcessID;
                            prProcess = Sub.prProcess;
                            prMachineNo = Sub.prMachineNo;
                            prArticleID = Sub.prArticleID;
                            prArticle = Sub.prArticle;
                            prArticleNo = Sub.prBuyerArticleNo;
                            prBuyerModel = Sub.prBuyerModel;
                            prlotID = Sub.prlotID;
                            prInstQty = Sub.prInstQty;

                            workrange = pastesheet.get_Range("A" + (insertline - 1), "B" + insertline);    //순번
                            workrange.Value2 = prNum;
                            workrange.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                            workrange.Font.Size = 11;

                            workrange = pastesheet.get_Range("C" + (insertline - 1), "E" + insertline);    //시작일~종료일
                            workrange.Value2 = prSDate + "~" + "\n\r" + prEDate;
                            workrange.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                            workrange.Font.Size = 11;

                            workrange = pastesheet.get_Range("F" + (insertline - 1), "J" + insertline);    //공정 /호기
                            workrange.Value2 = prProcess + "\n\r" + prMachineNo;
                            workrange.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                            workrange.Font.Size = 11;

                            workrange = pastesheet.get_Range("K" + (insertline - 1), "Q" + insertline);    //품명코드 /품명
                            workrange.Value2 = prArticleNo;
                            workrange.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                            workrange.Font.Size = 11;

                            workrange = pastesheet.get_Range("R" + (insertline - 1), "S" + insertline);    //품번/모델
                            workrange.Value2 = prInstQty;
                            workrange.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                            workrange.Font.Size = 9;

                            workrange = pastesheet.get_Range("T" + (insertline - 1), "AB" + (insertline - 1));    //LotID (font변경)
                            workrange.Value2 = "*" + prlotID + "*";
                            workrange.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                            workrange.Font.Size = 18;
                            workrange.Font.FontStyle = "Code39(2:3)";

                            workrange = pastesheet.get_Range("T" + insertline, "AB" + insertline);    //LotID
                            workrange.Value2 = prlotID;
                            workrange.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                            workrange.Font.Size = 8;

                            DataCount = i;
                            j += 2;
                        }
                    }
                }

                if (preview_click == true)
                {
                    excelapp.Visible = true;
                    pastesheet.PrintPreview();
                }
                else
                {
                    excelapp.Visible = true;
                    pastesheet.PrintOutEx();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        //취소버튼
        private void btnPrCancel_Click(object sender, RoutedEventArgs e)
        {
            PnlListPrint.IsOpen = false;
        }

        private void dgdPrintResearch()
        {
            if (DgdPrint != null) { DgdPrint.Items.Clear(); }
            FillPrintData();
            //FillPrintData_OnlyChecked();
        }

        private void SetWorkComboBox()
        {
            ObservableCollection<CodeView> cbWork = ComboBoxUtil.Instance.GetWorkProcess(0, "");

            this.cboPrWork.ItemsSource = cbWork;
            this.cboPrWork.DisplayMemberPath = "code_name";
            this.cboPrWork.SelectedValuePath = "code_id";
        }

        private void SetMachineComboBox(string processID)
        {
            ObservableCollection<CodeView> cbMachine = ComboBoxUtil.Instance.GetMachine(processID);

            //this.cbOrderFormR.ItemsSource = cboOrderForm;
            //this.cbOrderFormR.DisplayMemberPath = "code_name";
            //this.cbOrderFormR.SelectedValuePath = "code_id";

            this.cboPrMachine.ItemsSource = cbMachine;
            this.cboPrMachine.DisplayMemberPath = "code_name";
            this.cboPrMachine.SelectedValuePath = "code_id";
        }

        private void cboPrWork_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            numWorkSelect = cboPrWork.SelectedIndex;
            SetMachineComboBox(cboPrWork.SelectedValue != null ? cboPrWork.SelectedValue.ToString() : "");
            cboPrMachine.SelectedIndex = 0;
            if (DgdPrint != null) { DgdPrint.Items.Clear(); }
            FillPrintData();
        }

        private void cboPrMachine_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            numMachineSelecte = cboPrMachine.SelectedIndex;
            if (DgdPrint != null) { DgdPrint.Items.Clear(); }
            FillPrintData();
        }

        #region 작업지시서 일괄 인쇄 를 위한 검색 이벤트
        private void FillPrintData()
        {
            string strPrEDate = string.Empty;
            string strPrWorkID = string.Empty;
            string strPrMachineID = string.Empty;

            if (prdtpEDate.SelectedDate != null)
            {
                strPrEDate = prdtpEDate.SelectedDate.Value.ToString("yyyyMMdd");
            }
            else
            {
                prdtpEDate.SelectedDate = DateTime.Today;
                strPrEDate = prdtpEDate.SelectedDate.Value.ToString("yyyyMMdd");
            }

            if (cboPrWork.SelectedValue != null)
            {
                strPrWorkID = cboPrWork.SelectedValue != null ? cboPrWork.SelectedValue.ToString() : "";
            }
            else
            {
                strPrWorkID = string.Empty;
            }

            if (cboPrMachine.SelectedValue != null)
            {
                strPrMachineID = cboPrMachine.SelectedValue != null ? cboPrMachine.SelectedValue.ToString() : "";
            }
            else
            {
                strPrMachineID = string.Empty;
            }

            if (DgdPrint.Items.Count > 0)
            {
                DgdPrint.Items.Clear();
            }

            try
            {
                Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                sqlParameter.Add("SDate", prdtpSDate.SelectedDate.Value.ToString("yyyyMMdd"));
                sqlParameter.Add("EDate", prdtpEDate.SelectedDate.Value.ToString("yyyyMMdd"));
                sqlParameter.Add("ChkProcessID", cboPrWork.SelectedIndex == 0 ? 0 : numchkPrWork);
                sqlParameter.Add("ProcessID", strPrWorkID);
                sqlParameter.Add("nMachineID", numchkPrMachine);
                sqlParameter.Add("sMachineID", strPrMachineID);

                DataSet ds = DataStore.Instance.ProcedureToDataSet("xp_PlanInput_sPlanInputDetPrint", sqlParameter, false);

                if (ds != null && ds.Tables.Count > 0)
                {
                    DataTable dt = ds.Tables[0];

                    if (dt.Rows.Count == 0)
                    {
                    }
                    else
                    {
                        DataRowCollection drc = dt.Rows;

                        foreach (DataRow item in drc)
                        {
                            var Print_DTO = new Win_prd_Print_Plan_InputDet_CodeView()
                            {
                                prProcessID = item["ProcessID"] as string,
                                prProcess = item["Process"] as string,
                                prArticle = item["Article"] as string,
                                prArticleID = item["ArticleID"] as string,
                                prBuyerArticleNo = item["BuyerArticleNo"] as string,
                                prBuyerModel = item["BuyerModel"] as string,
                                prInstQty = item["InstQty"].ToString(),
                                prInstRemark = item["InstRemark"] as string,
                                prKCustom = item["KCustom"] as string,
                                prStartDate = item["StartDate"] as string,
                                prEndDate = item["EndDate"] as string,
                                prlotID = item["lotID"] as string,
                                prMachineNo = item["MachineNo"] as string,
                                prBuyerModelID = item["BuyerModelID"] as string,
                                prInstDetSeq = item["InstDetSeq"] as string,
                                prInstID = item["InstID"] as string,
                                prMachineID = item["MachineID"] as string,
                                prOrderNo = item["OrderNo"] as string,
                                prWorkQty = item["WorkQty"] as string
                            };

                            Print_DTO.prInstQty = Lib.Instance.returnNumStringZero(Print_DTO.prInstQty);

                            DgdPrint.Items.Add(Print_DTO);
                        }
                        drc.Clear();
                    }
                    dt.Clear();
                }
                ds.Clear();
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
        #endregion // 작업지시서 일괄 인쇄 를 위한 검색 이벤트

        #region 체크 한것들만 인쇄하기 위한 작업지시서 목록 검색

        private void FillPrintData_OnlyChecked()
        {
            if (dgdPrintForChecked.Items.Count > 0)
            {
                dgdPrintForChecked.Items.Clear();
            }

            string grp_InstID = "";

            string order_InstID = "";

            // 체크한것들만 InstID 모으기
            for (int i = 0; i < dgdMain.Items.Count; i++)
            {
                var Main = dgdMain.Items[i] as Win_prd_PlanInputView_U_CodeView;

                if (Main != null)
                {
                    if (Main.IsCheck == true)
                    {
                        if (grp_InstID.Length == 0)
                        {
                            grp_InstID += "'" + Main.InstID + "'";
                            order_InstID += Main.InstID;
                        }
                        else
                        {
                            grp_InstID += ",'" + Main.InstID + "'";
                            order_InstID += "," + Main.InstID;
                        }
                    }
                }
            }

            try
            {

                Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                sqlParameter.Clear();

                sqlParameter.Add("grp_InstID", grp_InstID);
                sqlParameter.Add("order_InstID", @order_InstID);

                DataSet ds = DataStore.Instance.ProcedureToDataSet("xp_PlanInput_sPlanInputDetPrint_WPF", sqlParameter, false);

                if (ds != null && ds.Tables.Count > 0)
                {
                    DataTable dt = ds.Tables[0];

                    if (dt.Rows.Count > 0)
                    {
                        DataRowCollection drc = dt.Rows;

                        foreach (DataRow item in drc)
                        {
                            var Print_DTO = new Win_prd_Print_Plan_InputDet_CodeView()
                            {
                                prProcessID = item["ProcessID"] as string,
                                prProcess = item["Process"] as string,
                                prArticle = item["Article"] as string,
                                prArticleID = item["ArticleID"] as string,
                                prBuyerArticleNo = item["BuyerArticleNo"] as string,
                                prBuyerModel = item["BuyerModel"] as string,
                                prInstQty = item["InstQty"].ToString(),
                                prInstRemark = item["InstRemark"] as string,
                                prKCustom = item["KCustom"] as string,
                                prStartDate = item["StartDate"] as string,
                                prEndDate = item["EndDate"] as string,
                                prlotID = item["lotID"] as string,
                                prMachineNo = item["MachineNo"] as string,
                                prBuyerModelID = item["BuyerModelID"] as string,
                                prInstDetSeq = item["InstDetSeq"] as string,
                                prInstID = item["InstID"] as string,
                                prMachineID = item["MachineID"] as string,
                                prOrderNo = item["OrderNo"] as string,
                                prWorkQty = item["WorkQty"] as string
                            };

                            Print_DTO.prInstQty = Lib.Instance.returnNumStringZero(Print_DTO.prInstQty);

                            dgdPrintForChecked.Items.Add(Print_DTO);
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

        #endregion // 체크 한것들만 인쇄하기 위한 작업지시서 목록 검색

        //작지목록 이벤트_날짜 바뀌면 조회데이터를 바꿔준다.
        private void prdtpSDate_SelectedDateChanged(object sender, SelectionChangedEventArgs e)
        {
            if (DgdPrint != null) { DgdPrint.Items.Clear(); }
            FillPrintData();
        }

        //작지목록 이벤트_날짜 바뀌면 조회데이터를 바꿔준다.
        private void prdtpEDate_SelectedDateChanged(object sender, SelectionChangedEventArgs e)
        {
            if (DgdPrint != null) { DgdPrint.Items.Clear(); }
            FillPrintData();
        }


        #region 기타 메서드 모음



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

        // 데이터피커 포맷으로 변경
        private string DatePickerFormatSlash(string str)
        {
            string result = "";

            if (str.Length == 8)
            {
                if (!str.Trim().Equals(""))
                {
                    result = str.Substring(4, 2) + "/" + str.Substring(6, 2);
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

        // 메인 데이터그리드 체크박스 이벤트
        private void chkC_Checked(object sender, RoutedEventArgs e)
        {

            CheckBox chkSender = sender as CheckBox;
            var Main = chkSender.DataContext as Win_prd_PlanInputView_U_CodeView;

            if (Main != null)
            {
                if (chkSender.IsChecked == true)
                {
                    Main.IsCheck = true;

                    if (ovcPlanView.Contains(Main) == false)
                        ovcPlanView.Add(Main);
                }
                else
                {
                    Main.IsCheck = false;

                    if (ovcPlanView.Contains(Main) == true)
                        ovcPlanView.Remove(Main);
                }
            }
        }

        private void chkC_Unchecked(object sender, RoutedEventArgs e)
        {
            CheckBox chkSender = sender as CheckBox;
            var Main = chkSender.DataContext as Win_prd_PlanInputView_U_CodeView;
            if (Main != null)
            {
                Main.IsCheck = false;
            }
        }

        // 전체 선택 체크
        private void AllCheck_Checked(object sender, RoutedEventArgs e)
        {
            for (int i = 0; i < dgdMain.Items.Count; i++)
            {
                var Main = dgdMain.Items[i] as Win_prd_PlanInputView_U_CodeView;
                if (Main != null)
                {
                    Main.IsCheck = true;
                }
            }
        }
        // 전체 선택 체크 해제
        private void AllCheck_Unchecked(object sender, RoutedEventArgs e)
        {
            for (int i = 0; i < dgdMain.Items.Count; i++)
            {
                var Main = dgdMain.Items[i] as Win_prd_PlanInputView_U_CodeView;
                if (Main != null)
                {
                    Main.IsCheck = false;
                }
            }
        }

        private void UserControl_SizeChanged(object sender, SizeChangedEventArgs e)
        {
            if (MainWindow.plInputFlag_SavePrint == true)
            {
                string InstID = "";
                bool FirstFlag = false;
                DateTime InstDate = new DateTime();

                if (MainWindow.plInput.Count > 0)
                {
                    foreach (string Key in MainWindow.plInput.Keys)
                    {
                        if (Key.ToUpper().Trim().Equals("DATE"))
                        {
                            InstDate = (DateTime)MainWindow.plInput["Date"];
                        }
                        else if (Key.ToUpper().Trim().Equals("INSTID"))
                        {
                            InstID = (string)MainWindow.plInput["InstID"];
                        }
                        else if (Key.ToUpper().Trim().Equals("FIRSTFLAG"))
                        {
                            FirstFlag = (bool)MainWindow.plInput["FirstFlag"];
                        }
                    }
                }

                if (FirstFlag == false)
                {
                    dtpSDate.SelectedDate = InstDate;
                    dtpEDate.SelectedDate = InstDate;

                    FillGrid();

                    for (int i = 0; i < dgdMain.Items.Count; i++)
                    {
                        var Main = dgdMain.Items[i] as Win_prd_PlanInputView_U_CodeView;
                        if (Main != null)
                        {
                            if (Main.InstID.Trim().Equals(InstID))
                            {
                                dgdMain.SelectedIndex = i;
                                Main.IsCheck = true;
                                break;
                            }
                        }
                    }

                    //MainWindow.plInputFlag_SavePrint = false;
                    MainWindow.plInput.Clear();
                }

            }
        }

        //전월
        private void btnLastMonth_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (dtpSDate.SelectedDate != null)
                {
                    DateTime ThatMonth1 = dtpSDate.SelectedDate.Value.AddDays(-(dtpSDate.SelectedDate.Value.Day - 1)); // 선택한 일자 달의 1일!

                    DateTime LastMonth1 = ThatMonth1.AddMonths(-1); // 저번달 1일
                    DateTime LastMonth31 = ThatMonth1.AddDays(-1); // 저번달 말일

                    dtpSDate.SelectedDate = LastMonth1;
                    dtpEDate.SelectedDate = LastMonth31;
                }
                else
                {
                    DateTime ThisMonth1 = DateTime.Today.AddDays(-(DateTime.Today.Day - 1)); // 이번달 1일

                    DateTime LastMonth1 = ThisMonth1.AddMonths(-1); // 저번달 1일
                    DateTime LastMonth31 = ThisMonth1.AddDays(-1); // 저번달 말일

                    dtpSDate.SelectedDate = LastMonth1;
                    dtpEDate.SelectedDate = LastMonth31;
                }
            }
            catch (Exception ee)
            {
                MessageBox.Show("오류지점 - btnLastMonth_Click : " + ee.ToString());
            }
        }

        //전일
        private void btnYesterday_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (dtpSDate.SelectedDate != null)
                {
                    dtpSDate.SelectedDate = dtpSDate.SelectedDate.Value.AddDays(-1);
                    dtpEDate.SelectedDate = dtpSDate.SelectedDate;
                }
                else
                {
                    dtpSDate.SelectedDate = DateTime.Today.AddDays(-1);
                    dtpEDate.SelectedDate = DateTime.Today.AddDays(-1);
                }
            }
            catch (Exception ee)
            {
                MessageBox.Show("오류지점 - btnYesterday_Click : " + ee.ToString());
            }
        }

        private void LabelBuyerArticleNoSearch_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            if (CheckBoxBuyerArticleNoSearch.IsChecked == true)
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
            if (e.Key == Key.Enter)
            {
                MainWindow.pf.ReturnCode(TextBoxBuyerArticleNoSearch, 76, TextBoxBuyerArticleNoSearch.Text);
            }
        }

        private void ButtonBuyerArticleNoSearch_Click(object sender, RoutedEventArgs e)
        {
            MainWindow.pf.ReturnCode(TextBoxBuyerArticleNoSearch, 76, TextBoxBuyerArticleNoSearch.Text);
        }

        private void DataGridComboBoxMtrExceptYN_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            var WPPUSC = dgdSub.SelectedItem as Win_prd_PlanInputView_U_Sub_CodeView;
            ComboBox comboBoxMtrExceptYN = sender as ComboBox;

            if (WPPUSC != null && WPPUSC.MtrExceptYN != null)
            {
                WPPUSC.MtrExceptYN = comboBoxMtrExceptYN.SelectedValue.ToString();
            }

        }

        private void DataGridComboBoxMtrExceptYN_Loaded(object sender, RoutedEventArgs e)
        {
            ComboBox comboBoxMtrExceptYN = sender as ComboBox;

            comboBoxMtrExceptYN.Items.Add("N");
            comboBoxMtrExceptYN.Items.Add("Y");
        }

        private void DataGridComboBoxMtrExceptYN_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                DataGridCell cell = lib.GetParent<DataGridCell>(sender as ComboBox);
                cell.Focus();
            }
        }

        private void DataGridComboBoxFirstInFirstOutYN_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            var WPPUSC = dgdSub.SelectedItem as Win_prd_PlanInputView_U_Sub_CodeView;
            ComboBox comboBoxMtrExceptYN = sender as ComboBox;

            if (WPPUSC != null && WPPUSC.FirstInFirstOutYN != null)
            {
                WPPUSC.FirstInFirstOutYN = comboBoxMtrExceptYN.SelectedValue.ToString();
            }
        }

        private void DataGridComboBoxFirstInFirstOutYN_Loaded(object sender, RoutedEventArgs e)
        {
            ComboBox comboBoxFirstInFirstOutYN = sender as ComboBox;

            comboBoxFirstInFirstOutYN.Items.Add("N");
            comboBoxFirstInFirstOutYN.Items.Add("Y");
        }

        private void DataGridComboBoxFirstInFirstOutYN_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                DataGridCell cell = lib.GetParent<DataGridCell>(sender as ComboBox);
                cell.Focus();
            }
        }

       
    }

    class Win_prd_PlanInputView_U_CodeView : BaseView
    {
        public override string ToString()
        {
            return (this.ReportAllProperties());
        }

        public int Num { get; set; }
        public string cls { get; set; }
        public string KCustom { get; set; }
        public string Article { get; set; }
        public string OrderID { get; set; }
        public string OrderNo { get; set; }
        public double OrderQty { get; set; }

        public double TotOrderinstqty { get; set; }
        public double notOrderInstQty { get; set; }
        public double OrderInstQty { get; set; }
        public double p1WorkQty { get; set; }
        public string p1ProcessID { get; set; }

        public string p1ProcessName { get; set; }
        public double InspectQty { get; set; }
        public double OutQty { get; set; }
        public string PatternID { get; set; }
        public string ArticleGrpID { get; set; }

        public string BuyerModel { get; set; }
        public string BuyerModelID { get; set; }
        public string BuyerArticleNo { get; set; }
        public string Remark { get; set; }
        public string Completed { get; set; }

        public string ArticleID { get; set; }
        public string InstID { get; set; }
        public string InstDate { get; set; }
        public string ProcPattern { get; set; }
        public string MtrExceptYN { get; set; }

        public string OutwareExceptYN { get; set; }
        public string LotID { get; set; }
        public string ArticleGrp { get; set; }
        public string Closed { get; set; }
        public string Progress { get; set; }
        public string InstSeq { get; set; }

        // 체크 되었는지 안되었는지
        public bool IsCheck { get; set; }
    }

    class Win_prd_PlanInputView_U_Sub_CodeView
    {
        public override string ToString()
        {
            return (this.ReportAllProperties());
        }

        public string ProcessID { get; set; }
        public string Process { get; set; }
        public string OrderNo { get; set; }
        public string InstID { get; set; }
        public string InstDetSeq { get; set; }

        public string InstQty { get; set; }
        public string StartDate { get; set; }
        public string StartDate_CV { get; set; }
        public string EndDate { get; set; }
        public string EndDate_CV { get; set; }
        public string InstRemark { get; set; }
        public string ArticleID { get; set; }
        public string BuyerArticleNo { get; set; }

        public string Article { get; set; }
        public string WorkQty { get; set; }
        public string lotID { get; set; }
        public string MachineID { get; set; }
        public string MachineNo { get; set; }

        public string Machine { get; set; }
        public string FirstProcessLotID { get; set; }
        public string InstDate { get; set; }
        public string OrderArticleID { get; set; }
        public string OrderArticle { get; set; }

        public int Num { get; set; }
        public string InstDate_CV { get; set; }
        public string ChildArticleID { get; set; }
        public string ChildBuyerArticleNo { get; set; }
        public string MtrExceptYN { get; set; }
        public string FirstInFirstOutYN { get; set; }
    }

    class Win_prd_Print_Plan_InputDet_CodeView
    {
        public override string ToString()
        {
            return (this.ReportAllProperties());
        }

        public string prProcessID { get; set; }
        public string prProcess { get; set; }
        public string prOrderNo { get; set; }
        public string prInstID { get; set; }
        public string prInstDetSeq { get; set; }

        public string prInstQty { get; set; }
        public string prStartDate { get; set; }
        public string prEndDate { get; set; }
        public string prInstRemark { get; set; }
        public string prArticleID { get; set; }

        public string prArticle { get; set; }
        public string prBuyerArticleNo { get; set; }
        public string prBuyerModelID { get; set; }
        public string prBuyerModel { get; set; }
        public string prWorkQty { get; set; }


        public string prlotID { get; set; }
        public string prMachineID { get; set; }
        public string prMachineNo { get; set; }
        public string prKCustom { get; set; }
    }
}

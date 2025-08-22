/**
 * 
 * @details 생산계획 하위품 관리
 * @author 정승학
 * @date 2019-07-30
 * @version 1.0
 * 
 * @see 소스 재작성 필요
 * 
 * @section MODIFYINFO 수정정보
 * - 수정일        - 수정자       : 수정내역
 * - 2000-01-01    - 정승학       : -----
 * 
 * 
 * */

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
using WizMes_WellMade.PopUp;

namespace WizMes_WellMade
{
    /// <summary>
    /// Win_prd_PlanInputArticleView_U.xaml에 대한 상호 작용 논리
    /// </summary>
    public partial class Win_prd_PlanInputArticleView_U : UserControl
    {
        string stDate = string.Empty;
        string stTime = string.Empty;

        int rowNum = 0;
        Win_prd_PlanInputArticleView_CodeView WinPlanArticleView = new Win_prd_PlanInputArticleView_CodeView();
        Win_prd_PlanInputView_Sub_CodeView WinPlanView = new Win_prd_PlanInputView_Sub_CodeView();
        Lib lib = new Lib();

        public Win_prd_PlanInputArticleView_U()
        {
            InitializeComponent();
        }

        private void UserControl_Loaded(object sender, RoutedEventArgs e)
        {
            stDate = DateTime.Now.ToString("yyyyMMdd");
            stTime = DateTime.Now.ToString("HHmm");

            DataStore.Instance.InsertLogByFormS(this.GetType().Name, stDate, stTime, "S");

            Lib.Instance.UiLoading(sender);
            chkDateSrh.IsChecked = true;
            btnToday_Click(null, null);
        }


        #region 일자변경
        private void lblDateSrh_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            if (chkDateSrh.IsChecked == true) { chkDateSrh.IsChecked = false; }
            else { chkDateSrh.IsChecked = true; }
        }

        private void chkDateSrh_Checked(object sender, RoutedEventArgs e)
        {
            dtpSDate.IsEnabled = true;
            dtpEDate.IsEnabled = true;
        }

        private void chkDateSrh_Unchecked(object sender, RoutedEventArgs e)
        {
            dtpSDate.IsEnabled = false;
            dtpEDate.IsEnabled = false;
        }

        private void btnThisMonth_Click(object sender, RoutedEventArgs e)
        {
            dtpSDate.SelectedDate = Lib.Instance.BringThisMonthDatetimeList()[0];
            dtpEDate.SelectedDate = Lib.Instance.BringThisMonthDatetimeList()[1];
        }

        private void btnToday_Click(object sender, RoutedEventArgs e)
        {
            dtpSDate.SelectedDate = DateTime.Today;
            dtpEDate.SelectedDate = DateTime.Today;
        }

        #endregion

        #region 상단 레이아웃 활성화 & 비활성화
        //품명
        private void lblArticleSrh_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            if (chkArticleSrh.IsChecked == true) { chkArticleSrh.IsChecked = false; }
            else { chkArticleSrh.IsChecked = true; }
        }

        private void chkArticleSrh_Checked(object sender, RoutedEventArgs e)
        {
            txtArticleSrh.IsEnabled = true;
            btnPfArticleSrh.IsEnabled = true;
            txtArticleSrh.Focus();
        }

        private void chkArticleSrh_Unchecked(object sender, RoutedEventArgs e)
        {
            txtArticleSrh.IsEnabled = false;
            btnPfArticleSrh.IsEnabled = false;
        }

        private void txtArticleSrh_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                MainWindow.pf.ReturnCode(txtArticleSrh, 77, "");
            }
        }

        private void btnPfArticleSrh_Click(object sender, RoutedEventArgs e)
        {
            MainWindow.pf.ReturnCode(txtArticleSrh, 77, "");
        }

        //거래처
        private void lblCustomSrh_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            if (chkCustomSrh.IsChecked == true) { chkCustomSrh.IsChecked = false; }
            else { chkCustomSrh.IsChecked = true; }
        }

        private void chkCustomSrh_Checked(object sender, RoutedEventArgs e)
        {
            txtCustomSrh.IsEnabled = true;
            btnPfCustomSrh.IsEnabled = true;
            txtCustomSrh.Focus();
        }

        private void chkCustomSrh_Unchecked(object sender, RoutedEventArgs e)
        {
            txtCustomSrh.IsEnabled = false;
            btnPfCustomSrh.IsEnabled = false;
        }

        private void txtCustomSrh_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                MainWindow.pf.ReturnCode(txtCustomSrh, (int)Defind_CodeFind.DCF_CUSTOM, "");
            }
        }

        private void btnPfCustomSrh_Click(object sender, RoutedEventArgs e)
        {
            MainWindow.pf.ReturnCode(txtCustomSrh, (int)Defind_CodeFind.DCF_CUSTOM, "");
        }
        #endregion

        #region 우측 상단 버튼
        /// <summary>
        /// 수정,추가 저장 후
        /// </summary>
        private void CanBtnControl()
        {
            Lib.Instance.UiButtonEnableChange_IUControl(this);
            dgdMain.IsHitTestVisible = true;
            dgdSub.IsHitTestVisible = false;

            grdSrh1.IsEnabled = true;
            grdSrh2.IsEnabled = true;
            grdSrh3.IsEnabled = true;
            dtpEDate.IsEnabled = true;
            dtpSDate.IsEnabled = true;
            chkDateSrh.IsEnabled = true;
            chkDateSrh.IsChecked = true;

            GridInputArea.IsEnabled = false;
        }

        /// <summary>
        /// 수정,추가 진행 중
        /// </summary>
        private void CantBtnControl()
        {
            Lib.Instance.UiButtonEnableChange_SCControl(this);
            dgdMain.IsHitTestVisible = false;
            dgdSub.IsHitTestVisible = true;

            grdSrh1.IsEnabled = false;
            dtpEDate.IsEnabled = false;
            grdSrh2.IsEnabled = false;
            grdSrh3.IsEnabled = false;

            GridInputArea.IsEnabled = true;
        }

        private void btnUpdate_Click(object sender, RoutedEventArgs e)
        {
            WinPlanView = dgdMain.SelectedItem as Win_prd_PlanInputView_Sub_CodeView;

            if (WinPlanView != null)
            {
                rowNum = dgdMain.SelectedIndex;
                tbkMsg.Text = "자료 수정 중";
                lblMsg.Visibility = Visibility.Visible;
                CantBtnControl();
            }
        }

        private void btnClose_Click(object sender, RoutedEventArgs e)
        {
            DataStore.Instance.InsertLogByFormS(this.GetType().Name, stDate, stTime, "E");
            Lib.Instance.ChildMenuClose(this.ToString());
        }

        private void btnSearch_Click(object sender, RoutedEventArgs e)
        {
            //검색버튼 비활성화
            btnSearch.IsEnabled = false;

            Dispatcher.BeginInvoke(new Action(() =>

            {
                try
                {
                    rowNum = 0;
                    using (Loading lw = new Loading(FillGrid))
                    {
                        lw.ShowDialog();
                        dgdMain.SelectedIndex = rowNum;

                        if (dgdMain.Items.Count <= 0)
                        {
                            MessageBox.Show("조회된 내용이 없습니다.");
                        }

                        btnSearch.IsEnabled = true;
                    }
                }
                catch (Exception ee)
                {
                    MessageBox.Show("예외처리 - " + ee.ToString());
                }

            }), System.Windows.Threading.DispatcherPriority.Background);
        }

        private void btnSave_Click(object sender, RoutedEventArgs e)
        {
            int Seq = -1;
            WinPlanView = dgdMain.SelectedItem as Win_prd_PlanInputView_Sub_CodeView;

            if (WinPlanView != null && Lib.Instance.IsIntOrAnother(WinPlanView.InstDetSeq))
            {
                Seq = int.Parse(WinPlanView.InstDetSeq);
            }
            else
            {
                MessageBox.Show("수정할 데이터가 비정상적입니다.");
                return;
            }

            if (SaveData(WinPlanView.InstID, int.Parse(WinPlanView.InstDetSeq)))
            {
                CanBtnControl();
                if (dgdSub.Items.Count > 0)
                {
                    dgdSub.Items.Clear();
                }

                re_Search(rowNum);
            }
            else
            {
                MessageBox.Show("저장실패");
            }
        }

        private void btnCancel_Click(object sender, RoutedEventArgs e)
        {
            CanBtnControl();
            if (dgdSub.Items.Count > 0)
            {
                dgdSub.Items.Clear();
            }

            re_Search(rowNum);
        }

        private void btnExcel_Click(object sender, RoutedEventArgs e)
        {
            DataTable dt = null;
            string Name = string.Empty;

            string[] lst = new string[4];
            lst[0] = "작업지시목록";
            lst[1] = "작업지시 하위품명";
            lst[2] = dgdMain.Name;
            lst[3] = dgdSub.Name;

            ExportExcelxaml ExpExc = new ExportExcelxaml(lst);
            ExpExc.ShowDialog();

            if (ExpExc.DialogResult.HasValue)
            {
                if (ExpExc.choice.Equals(dgdMain.Name))
                {
                    DataStore.Instance.InsertLogByForm(this.GetType().Name, "E");
                    if (ExpExc.Check.Equals("Y"))
                        dt = Lib.Instance.DataGridToDTinHidden(dgdMain);
                    else
                        dt = Lib.Instance.DataGirdToDataTable(dgdMain);

                    Name = dgdMain.Name;
                    Lib.Instance.GenerateExcel(dt, Name);
                    Lib.Instance.excel.Visible = true;
                }
                else if (ExpExc.choice.Equals(dgdSub.Name))
                {
                    if (ExpExc.Check.Equals("Y"))
                        dt = Lib.Instance.DataGridToDTinHidden(dgdSub);
                    else
                        dt = Lib.Instance.DataGirdToDataTable(dgdSub);

                    Name = dgdSub.Name;
                    Lib.Instance.GenerateExcel(dt, Name);
                    Lib.Instance.excel.Visible = true;
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

        #region 재검색
        /// <summary>
        /// 재검색(수정,삭제,추가 저장후에 자동 재검색)
        /// </summary>
        /// <param name="selectedIndex"></param>
        private void re_Search(int selectedIndex)
        {
            FillGrid();

            if (dgdMain.Items.Count > 0)
            {
                dgdMain.SelectedIndex = selectedIndex;
            }
            else
            {
                txtInstID.Text = "";
                txtOrderArticle.Text = "";
                txtOrderArticle.Tag = "";
                txtProcess.Text = "";
                txtProcess.Tag = "";
                txtQty.Text = "";
                txtQty.Tag = "";
                txtArticle.Text = "";
                txtArticle.Tag = "";
                txtBatjaWeight.Tag = "";
                txtComments.Tag = "";
                dtpInstDate.SelectedDate = DateTime.Today;
                dgdSub.Items.Clear();

                MessageBox.Show("조회된 데이터가 없습니다.");
                return;
            }
        }

        #endregion

        #region 조회
        /// <summary>
        /// 실조회
        /// </summary>
        private void FillGrid()
        {
            if (dgdMain.Items.Count > 0)
            {
                dgdMain.Items.Clear();

                TextBlockCountMain.Text = string.Empty;
            }

            try
            {
                Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                sqlParameter.Clear();
                sqlParameter.Add("ChkDate", chkDateSrh.IsChecked == true ? 1 : 0);
                sqlParameter.Add("SDate", chkDateSrh.IsChecked == true && dtpSDate.SelectedDate != null ? dtpSDate.SelectedDate.Value.ToString("yyyyMMdd") : "");
                sqlParameter.Add("EDate", chkDateSrh.IsChecked == true && dtpEDate.SelectedDate != null ? dtpEDate.SelectedDate.Value.ToString("yyyyMMdd") : "");
                sqlParameter.Add("ChkCustomID", chkCustomSrh.IsChecked == true ? 1 : 0);
                sqlParameter.Add("CustomID", chkCustomSrh.IsChecked == true ? txtCustomSrh.Tag.ToString() : "");

                sqlParameter.Add("ChkArticleID", chkArticleSrh.IsChecked == true ? 1 : 0);
                sqlParameter.Add("ArticleID", chkArticleSrh.IsChecked == true ? txtArticleSrh.Tag.ToString() : "");
                sqlParameter.Add("nChkOrder", 0);
                sqlParameter.Add("Order", "");
                sqlParameter.Add("nChkInstID", 0);
                sqlParameter.Add("sInstID", "");
                DataSet ds = DataStore.Instance.ProcedureToDataSet_LogWrite("xp_prd_sPlanInputDet_WPF", sqlParameter, true, "R");

                if (ds != null
                    && ds.Tables.Count > 0)
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
                            var WinPlanView = new Win_prd_PlanInputView_Sub_CodeView()
                            {
                                Num = i,
                                ProcessID = dr["ProcessID"].ToString(),
                                Process = dr["Process"].ToString(),
                                OrderNo = dr["OrderNo"].ToString(),
                                InstID = dr["InstID"].ToString(),
                                InstDetSeq = dr["InstDetSeq"].ToString(),

                                StartDate = dr["StartDate"].ToString(),
                                EndDate = dr["EndDate"].ToString(),
                                InstRemark = dr["InstRemark"].ToString(),
                                ArticleID = dr["ArticleID"].ToString(),

                                Article = dr["Article"].ToString(),
                                lotID = dr["lotID"].ToString(),
                                MachineID = dr["MachineID"].ToString(),
                                MachineNo = dr["MachineNo"].ToString(),

                                Machine = dr["Machine"].ToString(),
                                FirstProcessLotID = dr["FirstProcessLotID"].ToString(),
                                InstDate = dr["InstDate"].ToString(),
                                OrderArticleID = dr["OrderArticleID"].ToString(),
                                //OrderArticle = dr["OrderArticle"].ToString(),

                                BuyerArticleNo = dr["BuyerArticleNo"].ToString(),
                                InstQty = Convert.ToDouble(dr["InstQty"]),
                                WorkQty = Convert.ToDouble(dr["WorkQty"]),
                                InstDate_CV = DatePickerFormat(dr["InstDate"].ToString()),
                            };

                            dgdMain.Items.Add(WinPlanView);
                        }

                        TextBlockCountMain.Text = " ▶ 검색 결과 : " + i + " 건";
                    }
                }


            }
            catch (Exception ex)
            {
                MessageBox.Show("오류 발생, 오류 내용 : " + ex.ToString());
            }
        }

        #endregion

        #region 조회 sub
        private void FillGridSub(string InstID, string InstSeq)
        {
            if (dgdSub.Items.Count > 0)
            {
                dgdSub.Items.Clear();
            }

            try
            {
                int Seq = int.Parse(InstSeq);
                DataTable dt = Procedure.Instance.GetPlanInputDetArticleChild(InstID, Seq);
                int i = 0;

                if (dt.Rows.Count == 0)
                {
                    //MessageBox.Show("조회된 데이터가 없습니다.");
                }
                else
                {
                    DataRowCollection drc = dt.Rows;

                    foreach (DataRow dr in drc)
                    {
                        i++;
                        var WinPlanArticleView = new Win_prd_PlanInputArticleView_CodeView()
                        {
                            Num = i,
                            ChildSeq = dr["ChildSeq"].ToString(),
                            CHildArticleID = dr["CHildArticleID"].ToString(),
                            CHildArticle = dr["CHildArticle"].ToString(),
                            ScanExceptYN = dr["ScanExceptYN"].ToString()
                        };

                        dgdSub.Items.Add(WinPlanArticleView);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("오류 발생, 오류 내용 : " + ex.ToString());
            }
        }
        #endregion

        #region DgdMain_SelectionChanged
        private void DgdMain_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            WinPlanView = dgdMain.SelectedItem as Win_prd_PlanInputView_Sub_CodeView;

            if (WinPlanView != null)
            {
                this.DataContext = WinPlanView;
                FillGridSub(WinPlanView.InstID, WinPlanView.InstDetSeq);
            }
        }
        #endregion

        #region dgdsub 데이터그리드의 체크박스
        private void chkScanExcept_Click(object sender, RoutedEventArgs e)
        {
            if (lblMsg.Visibility == Visibility.Visible)
            {
                var PlanInputArticleView = dgdSub.SelectedItem as Win_prd_PlanInputArticleView_CodeView;

                if (PlanInputArticleView != null)
                {
                    CheckBox chk = sender as CheckBox;

                    if (chk != null)
                    {
                        if (chk.IsChecked == false)
                        {
                            PlanInputArticleView.ScanExceptYN = "N";
                        }
                        else
                        {
                            PlanInputArticleView.ScanExceptYN = "Y";
                        }
                    }
                }
            }
        }
        #endregion

        #region 저장
        private bool SaveData(string strInstID, int InstDetSeq)
        {
            bool flag = false;
            List<Procedure> Prolist = new List<Procedure>();
            List<Dictionary<string, object>> ListParameter = new List<Dictionary<string, object>>();

            try
            {
                Dictionary<string, object> sqlParameter = new Dictionary<string, object>();

                for (int i = 0; i < dgdSub.Items.Count; i++)
                {
                    var PlanInputArticleView_CodeView = dgdSub.Items[i] as Win_prd_PlanInputArticleView_CodeView;
                    sqlParameter = new Dictionary<string, object>();
                    sqlParameter.Clear();
                    sqlParameter.Add("sInstID", strInstID);
                    sqlParameter.Add("sInstSeq", InstDetSeq);
                    sqlParameter.Add("ChildSeq", PlanInputArticleView_CodeView.ChildSeq);
                    sqlParameter.Add("sScanExceptYN", PlanInputArticleView_CodeView.ScanExceptYN);
                    sqlParameter.Add("sUpdateUser", MainWindow.CurrentUser);

                    Procedure pro1 = new Procedure();
                    pro1.Name = "xp_PlanInput_uPlanInputDetArticle";
                    pro1.OutputUseYN = "N";
                    pro1.OutputName = "sInstID";
                    pro1.OutputLength = "20";

                    Prolist.Add(pro1);
                    ListParameter.Add(sqlParameter);
                }

                string[] Confirm = new string[2];
                Confirm = DataStore.Instance.ExecuteAllProcedureOutputNew_NewLog(Prolist, ListParameter, "U");
                if (Confirm[0] != "success")
                {
                    MessageBox.Show("[저장실패]\r\n" + Confirm[1].ToString());
                    flag = false;
                }
                else
                {
                    flag = true;
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
        #endregion

        #region 기타 메서드 모음

        // 텍스트 박스 숫자만 입력 되도록
        public void CheckIsNumericOnly(TextBox sender, TextCompositionEventArgs e)
        {
            decimal result;
            if (!(Decimal.TryParse(e.Text, out result)))
            {
                e.Handled = true;
            }
        }

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

        // 확장자 이미지 확인하기, 메인윈도우에 확장자 리스트 세팅
        private bool CheckImage(string ImageName)
        {
            string[] extensions = MainWindow.Extensions;

            bool flag = false;

            ImageName = ImageName.Trim().ToLower();
            foreach (string ext in extensions)
            {
                if (ImageName.EndsWith(ext))
                {
                    flag = true;
                }
            }

            return flag;
        }



        #endregion

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
    }

    #region CodeView 
    class Win_prd_PlanInputArticleView_CodeView : BaseView
    {
        public override string ToString()
        {
            return (this.ReportAllProperties());
        }

        public int Num { get; set; }
        public string ChildSeq { get; set; }
        public string CHildArticleID { get; set; }
        public string CHildArticle { get; set; }
        public string ScanExceptYN { get; set; }
    }

    public class Win_prd_PlanInputView_Sub_CodeView : BaseView
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

        public double InstQty { get; set; }
        public string StartDate { get; set; }
        public string EndDate { get; set; }
        public string InstRemark { get; set; }
        public string ArticleID { get; set; }

        public string Article { get; set; }
        public double WorkQty { get; set; }
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

        public string BuyerArticleNo { get; set; }

        public string Remark { get; set; }
    }

    #endregion
}

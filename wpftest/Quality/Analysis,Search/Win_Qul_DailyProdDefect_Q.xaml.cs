using System;
using System.Data;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
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
    /// Win_Qul_DailyProdDefect_Q.xaml에 대한 상호 작용 논리
    /// </summary>
    public partial class Win_Qul_DailyProdDefect_Q : UserControl
    {
        string stDate = string.Empty;
        string stTime = string.Empty;

        Lib lib = new Lib();

        public Win_Qul_DailyProdDefect_Q()
        {
            InitializeComponent();
        }

        private void UserControl_Loaded(object sender, RoutedEventArgs e)
        {
            stDate = DateTime.Now.ToString("yyyyMMdd");
            stTime = DateTime.Now.ToString("HHmm");

            DataStore.Instance.InsertLogByFormS(this.GetType().Name, stDate, stTime, "S");

            Lib.Instance.UiLoading(sender);
            btnToday_Click(null, null);
        }

        #region 상단컨트롤
        #region 날짜컨트롤
        // 날짜
        private void lblOrderDay_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            search_CheckBox_Control(chkOrderDay);
        }

        private void chkOrderDay_Checked(object sender, RoutedEventArgs e)
        {
            if (dtpSDate != null && dtpEDate != null)
            {
                dtpSDate.IsEnabled = true;
                dtpEDate.IsEnabled = true;
            }
        }

        private void chkOrderDay_Unchecked(object sender, RoutedEventArgs e)
        {
            dtpSDate.IsEnabled = false;
            dtpEDate.IsEnabled = false;
        }

        // 전월
        private void btnLastMonth_Click(object sender, RoutedEventArgs e)   { search_BtnDate_Control(1);}
        // 금월
        private void btnThisMonth_Click(object sender, RoutedEventArgs e)   { search_BtnDate_Control(2); }
        // 전일
        private void btnYesterDay_Click(object sender, RoutedEventArgs e)   { search_BtnDate_Control(3); }
        // 금일
        private void btnToday_Click(object sender, RoutedEventArgs e)       { search_BtnDate_Control(); }
        #endregion 날짜컨트롤

        #region 검색컨트롤
        // 거래처
        private void lblCustom_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)         { search_CheckBox_Control(chkCustom); }
        private void chkCustom_Checked(object sender, RoutedEventArgs e)                        { search_CheckBox_Checked_Control(true, txtCustom, btnPfCustom); }
        private void chkCustom_Unchecked(object sender, RoutedEventArgs e)                      { search_CheckBox_Checked_Control(false, txtCustom, btnPfCustom); }
        private void txtCustom_KeyDown(object sender, KeyEventArgs e)                           { if (e.Key == Key.Enter) search_PlusFinder_Control(txtCustom, (int)Defind_CodeFind.DCF_CUSTOM, ""); }
        private void btnPfCustom_Click(object sender, RoutedEventArgs e)                        { search_PlusFinder_Control(txtCustom, (int)Defind_CodeFind.DCF_CUSTOM, ""); }

        // 최종고객사
        private void lblInCustom_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)       { search_CheckBox_Control(chkInCustom); }
        private void chkInCustom_Checked(object sender, RoutedEventArgs e)                      { search_CheckBox_Checked_Control(true, txtInCustom, btnPfInCustom); }
        private void chkInCustom_Unchecked(object sender, RoutedEventArgs e)                    { search_CheckBox_Checked_Control(false, txtInCustom, btnPfInCustom); }
        private void txtInCustom_KeyDown(object sender, KeyEventArgs e)                         { if (e.Key == Key.Enter) search_PlusFinder_Control(txtInCustom, (int)Defind_CodeFind.DCF_CUSTOM, ""); }
        private void btnPfInCustom_Click(object sender, RoutedEventArgs e)                      { search_PlusFinder_Control(txtInCustom, (int)Defind_CodeFind.DCF_CUSTOM, ""); }

        // 품번
        private void lblBuyerArticleNo_MouseLeftButtonUp(object sender, MouseButtonEventArgs e) { search_CheckBox_Control(chkBuyerArticleNo); }
        private void chkBuyerArticleNo_Checked(object sender, RoutedEventArgs e)                { search_CheckBox_Checked_Control(true, txtBuyerArticleNo, btnPfBuyerArticleNo); }
        private void chkBuyerArticleNo_Unchecked(object sender, RoutedEventArgs e)              { search_CheckBox_Checked_Control(false, txtBuyerArticleNo, btnPfBuyerArticleNo); }
        private void txtBuyerArticleNo_KeyDown(object sender, KeyEventArgs e)                   { if (e.Key == Key.Enter) search_PlusFinder_Control(txtBuyerArticleNo, 76, ""); }
        private void btnPfBuyerArticleNo_Click(object sender, RoutedEventArgs e)                { search_PlusFinder_Control(txtBuyerArticleNo, 76, ""); }

        // 품명
        private void lblArticle_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)        { search_CheckBox_Control(chkArticle); }
        private void chkArticle_Checked(object sender, RoutedEventArgs e)                       { search_CheckBox_Checked_Control(true, txtArticle, btnPfArticle); }
        private void chkArticle_Unchecked(object sender, RoutedEventArgs e)                     { search_CheckBox_Checked_Control(false, txtArticle, btnPfArticle); }
        private void txtArticle_KeyDown(object sender, KeyEventArgs e)                          { if (e.Key == Key.Enter) search_PlusFinder_Control(txtArticle, 77, ""); }
        private void btnPfArticle_Click(object sender, RoutedEventArgs e)                       { search_PlusFinder_Control(txtArticle, 77, ""); }
        #endregion 검색컨트롤

        #region 버튼컨트롤
        // 검색
        private void btnSearch_Click(object sender, RoutedEventArgs e)
        {
            using (Loading lw = new Loading(beSearch))
            {
                lw.ShowDialog();
            }
        }

        // 닫기
        private void btnClose_Click(object sender, RoutedEventArgs e)
        {
            DataStore.Instance.InsertLogByFormS(this.GetType().Name, stDate, stTime, "E");
            Lib.Instance.ChildMenuClose(this.ToString());
        }

        // 엑셀
        private void btnExcel_Click(object sender, RoutedEventArgs e)
        {
            DataTable dt = null;

            string[] lst = new string[4];
            lst[0] = "품번별 불량유형";
            lst[1] = "불량유형별 품번";
            lst[2] = dgdMainLeft.Name;
            lst[3] = dgdMainRight.Name;

            ExportExcelxaml ExpExc = new ExportExcelxaml(lst);
            ExpExc.WindowStartupLocation = WindowStartupLocation.CenterScreen;
            ExpExc.ShowDialog();

            if (ExpExc.DialogResult.HasValue)
            {
                if (ExpExc.choice.Equals(dgdMainLeft.Name))
                {
                    DataStore.Instance.InsertLogByForm(this.GetType().Name, "E");

                    if (ExpExc.Check.Equals("Y"))
                        dt = lib.DataGridToDTinHidden(dgdMainLeft);
                    else
                        dt = lib.DataGirdToDataTable(dgdMainLeft);

                    if (lib.GenerateExcel(dt, dgdMainLeft.Name))
                    {
                        lib.excel.Visible = true;
                        lib.ReleaseExcelObject(lib.excel);
                    }
                    else
                        return;
                }
                else if (ExpExc.DialogResult.HasValue)
                {
                    if (ExpExc.choice.Equals(dgdMainRight.Name))
                    {
                        DataStore.Instance.InsertLogByForm(this.GetType().Name, "E");

                        if (ExpExc.Check.Equals("Y"))
                            dt = lib.DataGridToDTinHidden(dgdMainRight);
                        else
                            dt = lib.DataGirdToDataTable(dgdMainRight);

                        if (lib.GenerateExcel(dt, dgdMainRight.Name))
                        {
                            lib.excel.Visible = true;
                            lib.ReleaseExcelObject(lib.excel);
                        }
                        else
                            return;
                    }
                }
                else
                {
                    if (dt != null)
                        dt.Clear();
                }
            }
        }
        #endregion 버튼컨트롤

        private void search_BtnDate_Control(byte flag = 0)
        {
            // 1: 전월, 2: 금월, 3: 전일, 그외: 금일

            DateTime[] dateTime = { DateTime.Today, DateTime.Today };
            switch (flag)
            {
                case 1: dateTime = lib.BringLastMonthContinue(dtpSDate.SelectedDate.Value); break;
                case 2: dateTime = lib.BringThisMonthDatetimeList().ToArray(); break;
                case 3: dateTime = lib.BringLastDayDateTimeContinue(dtpEDate.SelectedDate.Value); break;
            }

            dtpSDate.SelectedDate = dateTime[0];
            dtpEDate.SelectedDate = dateTime[1];
        }

        private void search_CheckBox_Control(CheckBox checkBox)
        {
            checkBox.IsChecked = checkBox.IsChecked == true ? false : true;
        }

        private void search_CheckBox_Checked_Control(bool isCheck, TextBox textBox, Button button)
        {
            textBox.IsEnabled = isCheck;
            button.IsEnabled = isCheck;

            if (isCheck)
                textBox.Focus();
        }

        private void search_PlusFinder_Control(TextBox textBox, int large, string sMiddle)
        {
            MainWindow.pf.ReturnCode(textBox, large, sMiddle);
        }
        #endregion 상단컨트롤

        #region 주요메서드 - 조회
        private void beSearch()
        {
            btnSearch.IsEnabled = false;

            Dispatcher.BeginInvoke(new Action(() =>
            {
                if (CheckData())
                    FillGrid();
            }), System.Windows.Threading.DispatcherPriority.Background);

            btnSearch.IsEnabled = true;
        }

        private void FillGrid()
        {
            dgdMainLeft.Items.Clear();
            dgdMainRight.Items.Clear();
            dgdTotal.Items.Clear();

            try
            {
                // 2개의 그리드 호출
                for (int i = 0; i < 2; i++)
                {
                    Dictionary<string, object> sqlParameter = new Dictionary<string, object>();

                    // 날짜
                    sqlParameter.Add("ChkDate", chkOrderDay.IsChecked == true ? 1 : 0);
                    sqlParameter.Add("SDate", chkOrderDay.IsChecked == true ? dtpSDate.SelectedDate.Value.ToString("yyyyMMdd") : "");
                    sqlParameter.Add("EDate", chkOrderDay.IsChecked == true ? dtpEDate.SelectedDate.Value.ToString("yyyyMMdd") : "");

                    // 거래처
                    sqlParameter.Add("ChkCustom", chkCustom.IsChecked == true ? 1 : 0);
                    sqlParameter.Add("CustomID", chkCustom.IsChecked == true ? (txtCustom.Tag != null ? txtCustom.Tag.ToString() : "") : "");
                    // 최종고객사
                    sqlParameter.Add("ChkInCustom", chkInCustom.IsChecked == true ? 1 : 0);
                    sqlParameter.Add("InCustomID", chkInCustom.IsChecked == true ? (txtInCustom.Tag != null ? txtInCustom.Tag.ToString() : "") : "");

                    // 품번
                    sqlParameter.Add("ChkArticleID", chkBuyerArticleNo.IsChecked == true ? 1 : 0);
                    sqlParameter.Add("ArticleID", chkBuyerArticleNo.IsChecked == true ? (txtBuyerArticleNo.Tag == null ? "" : txtBuyerArticleNo.Tag.ToString()) : "");
                    // 품명
                    sqlParameter.Add("ChkArticle", chkArticle.IsChecked == true ? 1 : 0);
                    sqlParameter.Add("Article", chkArticle.IsChecked == true ? (txtArticle.Text == string.Empty ? "" : txtArticle.Text) : "");

                    // 구분
                    int gubun = i + 1;
                    bool defect_tot = gubun == 1;
                    sqlParameter.Add("Gubun", gubun);

                    DataSet ds = DataStore.Instance.ProcedureToDataSet_LogWrite("xp_Inspect_sProcessDefect", sqlParameter, true, "R");
                    if (ds != null && ds.Tables.Count > 0)
                    {
                        DataTable dt = ds.Tables[0];
                        if (dt.Rows.Count > 0)
                        {
                            int idx = 0;
                            DataRowCollection drc = dt.Rows;
                            foreach (DataRow dr in drc)
                            {
                                idx++;
                                var Win = new Win_Qul_DailyProdDefect_Q_CodeView()
                                {
                                    Num = idx,
                                    cls = dr["cls"].ToString().Trim(),
                                    BuyerArticleNo = dr["BuyerArticleNo"].ToString(),
                                    ArticleID = dr["ArticleID"].ToString(),
                                    Article = dr["Article"].ToString(),
                                    WorkQty = stringFormatN0(dr["WorkQty"].ToString()),
                                    DefectID = dr["DefectID"].ToString(),
                                    KDefect = dr["KDefect"].ToString(),
                                    DefectQty = stringFormatN0(dr["DefectQty"].ToString()),
                                    DefectRate = dr["DefectRate"].ToString()
                                };
                                
                                if (Win.cls.Equals("1"))
                                {
                                    Win.cls = "";
                                }
                                else if (Win.cls.Equals("2")) // 불량유형계 및 품명계
                                {
                                    Win.cls = defect_tot ? "불량계" : "품명계";

                                    if (defect_tot) Win.KDefect = "";
                                    else            Win.Article = "";

                                    if (defect_tot)
                                    { Win.KDefect = ""; Win.ColorGreen = "true"; }
                                    else
                                    { Win.BuyerArticleNo = ""; Win.ColorGreen = "true"; }
                                }
                                else if (Win.cls.Equals("9")) // 총계
                                {
                                    Win.cls = "총계";

                                    if (i == 0)
                                        dgdTotal.Items.Add(Win);
                                }

                                if (defect_tot) dgdMainLeft.Items.Add(Win);
                                else            dgdMainRight.Items.Add(Win);
                            }
                        }
                    }
                }

                if (dgdMainLeft.Items.Count == 0 && dgdMainRight.Items.Count == 0)
                    MessageBox.Show("조회된 데이터가 없습니다.");
            }
            catch (Exception ex)
            {
                MessageBox.Show("[오류내용]: " + ex.ToString());
            }
            finally
            {
                DataStore.Instance.CloseConnection();
            }
        }
        #endregion 주요메서드 - 조회

        #region 기타메서드
        private void dataGrid_SizeChanged(object sender, SizeChangedEventArgs e)
        {
            DataGrid dgs = sender as DataGrid;
            if (dgs.ColumnHeaderHeight == 0)
            {
                dgs.ColumnHeaderHeight = 1;
            }
            double a = e.NewSize.Height / 100;
            double b = e.PreviousSize.Height / 100;
            double c = a / b;

            if (c != double.PositiveInfinity && c != 0 && double.IsNaN(c) == false)
            {
                dgs.ColumnHeaderHeight = dgs.ColumnHeaderHeight * c;
                dgs.FontSize = dgs.FontSize * c;
            }
        }

        private bool CheckData()
        {
            string word = "";
            if (chkCustom.IsChecked == true && string.IsNullOrEmpty(txtCustom.Text))                        word = "거래처";
            else if (chkInCustom.IsChecked == true && string.IsNullOrEmpty(txtInCustom.Text))               word = "최종고객사";
            else if (chkArticle.IsChecked == true && string.IsNullOrEmpty(txtArticle.Text))                 word = "품명";
            else if (chkBuyerArticleNo.IsChecked == true && string.IsNullOrEmpty(txtBuyerArticleNo.Text))   word = "품번";

            bool flag = true;
            if (word != "")
            {
                flag = false;
                string msg = word + " 선택이 되지 않았습니다.\n체크를 해제하거나 " + word + "을 선택하고 검색해 주세요.";
                MessageBox.Show(msg);
            }

            return flag;
        }

        private string DatePickerFormat(string str)
        {
            string result = "";

            if (str.Length == 8)
            {
                if (!str.Trim().Equals(""))
                    result = str.Substring(0, 4) + "-" + str.Substring(4, 2) + "-" + str.Substring(6, 2);
            }

            return result;
        }

        // 천자리 콤마, 소수점 버리기
        private string stringFormatN0(object obj)
        {
            return string.Format("{0:N0}", obj);
        }
        #endregion 기타메서드
    }

    #region View 클래스
    class Win_Qul_DailyProdDefect_Q_CodeView : BaseView
    {
        public int Num { get; set; }
        public string cls { get; set; }
        public string BuyerArticleNo { get; set; }
        public string ArticleID { get; set; }
        public string Article { get; set; }
        public string WorkQty { get; set; }
        public string DefectID { get; set; }
        public string KDefect { get; set; }
        public string DefectQty { get; set; }
        public string DefectRate { get; set; }

        public string ColorGreen { get; set; }
        public string ColorRed { get; set; }

    }
    #endregion View 클래스
}

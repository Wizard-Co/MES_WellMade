using System;
using System.Collections.Generic;
using System.Data;
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
using static System.Windows.Forms.AxHost;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.TaskbarClock;
using WPF.MDI;
using WizMes_WellMade.PopUP;
using WizMes_WellMade.PopUp;

namespace WizMes_WellMade
{
    /// <summary>
    /// Win_Qul_Sts_AnalDetail_Q.xaml에 대한 상호 작용 논리
    /// </summary>
    public partial class Win_Qul_Sts_AnalDetail_Q : UserControl
    {
        #region 전역 변수 선언
        Lib lib = new Lib();
        PlusFinder pf = new PlusFinder();

        #endregion

        public Win_Qul_Sts_AnalDetail_Q()
        {
            InitializeComponent();
        }
        public void Win_Qul_Sts_AnalDetail_Q_Loaded(object sender, RoutedEventArgs e)
        {
            chkInspectDay.IsChecked = true;
            chkArticle.IsChecked = false;
         
            dtpFromDate.Text = DateTime.Now.ToString("yyyy-MM-dd");
            dtpToDate.Text = DateTime.Now.ToString("yyyy-MM-dd");

            string itemCode = string.Empty;
            string itemName = string.Empty;
            string chkDate = string.Empty;
            string sDate = string.Empty;
            string eDate = string.Empty;

            if (MainWindow.ChartDeatil != null
                && MainWindow.ChartDeatil.Count > 0)
            {
                itemCode = MainWindow.ChartDeatil[0];
                itemName = MainWindow.ChartDeatil[1];
                chkDate = MainWindow.ChartDeatil[2];
                sDate = MainWindow.ChartDeatil[3];
                eDate = MainWindow.ChartDeatil[4];

                if (chkDate.Equals("True"))
                {
                    dtpFromDate.SelectedDate = DateTime.Parse(DatePickerFormat(sDate));
                    dtpToDate.SelectedDate = DateTime.Parse(DatePickerFormat(eDate));
                }

                chkArticle.IsChecked = true;
                txtArticle.Text = itemName;

                FillGrid();
            }
            else
            {
                dtpFromDate.SelectedDate = DateTime.Today;
                dtpToDate.SelectedDate = DateTime.Today;
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

        #region 검색조건 이벤트

        // 일자 체크 이벤트
        private void chkInspectDay_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            if (chkInspectDay.IsChecked == true) { chkInspectDay.IsChecked = false; }
            else { chkInspectDay.IsChecked = false; }
        }

        // 일자 체크 이벤트
        private void chkInspectDay_Checked(object sender, RoutedEventArgs e)
        {
            dtpFromDate.IsEnabled = true;
            dtpToDate.IsEnabled = true;
            btnLastMonth.IsEnabled = true;
            btnThisMonth.IsEnabled = true;
            btnYesterDay.IsEnabled = true;
            btnToday.IsEnabled = true;
        }

        // 일자 체크 이벤트
        private void chkInspectDay_Unchecked(object sender, RoutedEventArgs e)
        {
            dtpFromDate.IsEnabled = false;
            dtpToDate.IsEnabled = false;
            btnLastMonth.IsEnabled = false;
            btnThisMonth.IsEnabled = false;
            btnYesterDay.IsEnabled = false;
            btnToday.IsEnabled = false;
        }

        // 전월 클릭 이벤트
        private void btnLastMonth_Click(object sender, RoutedEventArgs e)
        {
            DateTime[] SearchDate = lib.BringLastMonthContinue(dtpFromDate.SelectedDate.Value);

            dtpFromDate.SelectedDate = SearchDate[0];
            dtpToDate.SelectedDate = SearchDate[1];
        }

        // 금월 클릭 이벤트
        private void btnThisMonth_Click(object sender, RoutedEventArgs e)
        {
            dtpFromDate.SelectedDate = lib.BringThisMonthDatetimeList()[0];
            dtpToDate.SelectedDate = lib.BringThisMonthDatetimeList()[1];
        }

        // 전일 클릭 이벤트
        private void btnYesterDay_Click(object sender, RoutedEventArgs e)
        {
            DateTime[] SearchDate = lib.BringLastDayDateTimeContinue(dtpToDate.SelectedDate.Value);

            dtpFromDate.SelectedDate = SearchDate[0];
            dtpToDate.SelectedDate = SearchDate[1];
        }

        // 금일 클릭 이벤트
        private void btnToday_Click(object sender, RoutedEventArgs e)
        {
            dtpFromDate.SelectedDate = DateTime.Today;
            dtpToDate.SelectedDate = DateTime.Today;
        }

        // 품번 체크 이벤트
        private void chkArticle_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            if (chkArticle.IsChecked == true) { chkArticle.IsChecked = false;}
            else { chkArticle.IsChecked = true;}
        }

        // 품번 체크 이벤트
        private void chkArticle_Checked(object sender, RoutedEventArgs e)
        {
            txtArticle.IsEnabled = true;
            btnArticle.IsEnabled = true;
        }

        // 품번 체크 이벤트
        private void chkArticle_Unchecked(object sender, RoutedEventArgs e)
        {
            txtArticle.IsEnabled = false; 
            btnArticle.IsEnabled = false;
        }

        // 품번 키다운 이벤트
        private void TxtArticle_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                pf.ReturnCode(txtArticle, 83, txtArticle.Text);
            }
        }

        // 품번 플러스파인터 클릭 이벤트
        private void btnArticle_Click(object sender, RoutedEventArgs e)
        {
            pf.ReturnCode(txtArticle, 83, txtArticle.Text);
        }

        private void btnSearch_Click(object sender, RoutedEventArgs e)
        {
            using (Loading ld = new Loading(beSearch))
            {
                ld.ShowDialog();
            }
        }

        private void beSearch()
        {
            DataStore.Instance.InsertLogByForm(this.GetType().Name, "R");
            FillGrid();
        }

        private void btnClose_Click(object sender, RoutedEventArgs e)
        {
            DataStore.Instance.InsertLogByFormS(this.GetType().Name, dtpFromDate.SelectedDate.ToString(), dtpToDate.SelectedDate.ToString(), "E");
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

        private void btnExcel_Click(object sender, RoutedEventArgs e)
        {
            DataTable dt = null;
            string Name = string.Empty;

            string[] lst = new string[6];
            lst[0] = dgdMain.Name;
            lst[1] = dgdSubs.Name;

            ExportExcelxaml ExpExc = new ExportExcelxaml(lst);
            ExpExc.ShowDialog();

            if (ExpExc.DialogResult.HasValue)
            {
                if (ExpExc.choice.Equals(dgdMain.Name))
                {
                    DataStore.Instance.InsertLogByForm(this.GetType().Name, "E");
                    if (ExpExc.Check.Equals("Y"))
                        dt = lib.DataGridToDTinHidden(dgdMain);
                    else
                        dt = lib.DataGirdToDataTable(dgdMain);

                    Name = dgdMain.Name;

                    if (lib.GenerateExcel(dt, Name))
                    {
                        lib.excel.Visible = true;
                        lib.ReleaseExcelObject(lib.excel);
                    }
                }
                else if (ExpExc.choice.Equals(dgdSubs.Name))
                {
                    if (ExpExc.Check.Equals("Y"))
                        dt = lib.DataGridToDTinHidden(dgdSubs);
                    else
                        dt = lib.DataGirdToDataTable(dgdSubs);

                    Name = dgdSubs.Name;
                    if (lib.GenerateExcel(dt, Name))
                    {
                        lib.excel.Visible = true;
                        lib.ReleaseExcelObject(lib.excel);
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



        private void FillGrid()
        {
            try
            {
                dgdMain.Items.Clear();
                dgdSubs.Items.Clear();

                #region 과거 동일 품번 품질 추이
                Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                sqlParameter.Add("FromDate", dtpFromDate.Text.Replace("-","").ToString());
                sqlParameter.Add("ToDate", dtpToDate.Text.Replace("-", "").ToString());
                sqlParameter.Add("BuyerArticleNo", chkArticle.IsChecked == true ? txtArticle.Text.ToString() : "");

                DataSet ds = DataStore.Instance.ProcedureToDataSet("xp_Inspect_sAnalDetail", sqlParameter, false);
                if (ds != null
                    && ds.Tables.Count > 0)
                {
                    // 유형별 불량 모음
                    DataTable dt = ds.Tables[0];

                    if (dt.Rows.Count == 0)
                    {
                        return;
                    }
                    else
                    {
                        DataRowCollection drc = dt.Rows;

                        foreach (DataRow dr in drc)
                        {
                            var winQulDetail = new Win_Qul_Sts_AnalDetail_Q_CodeView()
                            {
                                step = dr["step"].ToString(),
                                CustomID = dr["CustomID"].ToString(),
                                InspectDate = dr["InspectDate"].ToString(),
                                ArticleID = dr["ArticleID"].ToString(),
                                DefectSymtom = dr["DefectSymtom"].ToString(),
                                DefectQty = dr["DefectQty"].ToString(),
                                GroupingNo = dr["GroupingNo"].ToString(),
                                GroupingName = dr["GroupingName"].ToString(),
                                Name = dr["Name"].ToString(),
                            };

                            dgdMain.Items.Add(winQulDetail);
                        }
                    }
                }
                #endregion

                #region 설비별 품질 추이

                sqlParameter.Clear();
                sqlParameter.Add("FromDate", dtpFromDate.Text.Replace("-", "").ToString());
                sqlParameter.Add("ToDate", dtpToDate.Text.Replace("-", "").ToString());
                sqlParameter.Add("BuyerArticleNo", chkArticle.IsChecked == true ? txtArticle.Text.ToString() : "");

                ds = DataStore.Instance.ProcedureToDataSet("xp_Inspect_sAnalDetail_Machine", sqlParameter, false);
                if (ds != null
                    && ds.Tables.Count > 0)
                {
                    DataTable dt = ds.Tables[0];

                    if (dt.Rows.Count == 0)
                    {
                        return;
                    }
                    else
                    {
                        DataRowCollection drc = dt.Rows;

                        foreach (DataRow dr in drc)
                        {
                            var winQulDetailMachine = new Win_Qul_Sts_AnalDetail_Q_Sub_CodeView()
                            {
                                Machine = dr["Machine"].ToString(),
                                BuyerArticleNo = dr["BuyerArticleNo"].ToString(),
                                Article = dr["Article"].ToString(),
                                InspectDate = dr["InspectDate"].ToString(),
                                WorkQty = dr["WorkQty"].ToString(),
                                DefectQty = dr["DefectQty"].ToString(),
                                Name = dr["Name"].ToString(),
                            };

                            dgdSubs.Items.Add(winQulDetailMachine);
                        }
                    }
                }

                #endregion

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
    }

    public class Win_Qul_Sts_AnalDetail_Q_CodeView : BaseView
    {
        public string step { get; set; }
        public string CustomID { get; set; }
        public string InspectDate { get; set; }
        public string ArticleID { get; set; }
        public string DefectSymtom { get; set; }
        public string DefectQty { get; set; }
        public string GroupingNo { get; set; }
        public string GroupingName { get; set; }
        public string Name { get; set; }
    }

    public class Win_Qul_Sts_AnalDetail_Q_Sub_CodeView : BaseView
    {
        public string Machine { get; set; }
        public string BuyerArticleNo { get; set; }
        public string Article { get; set; }
        public string InspectDate { get; set; }
        public string WorkQty { get; set; }
        public string DefectQty { get; set; }
        public string Name { get; set; }
    }
}

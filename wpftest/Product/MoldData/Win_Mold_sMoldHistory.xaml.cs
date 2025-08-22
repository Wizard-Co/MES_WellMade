using WizMes_WellMade.PopUP;
using System;
using System.Collections.Generic;
using System.Data;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using WPF.MDI;
using System.Collections.ObjectModel;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.StartPanel;
using System.Data.SqlClient;

namespace WizMes_WellMade
{
    /// <summary>
    /// Win_Mold_sMoldHistory.xaml에 대한 상호 작용 논리
    /// </summary>
    public partial class Win_Mold_sMoldHistory : UserControl
    {
        #region 변수선언 및 로드

        Lib lib = new Lib();

        public Win_Mold_sMoldHistory()
        {
            InitializeComponent();
        }

        private void UserControl_Loaded(object sender, RoutedEventArgs e)
        {
            dtpSDate.SelectedDate = DateTime.Today;
            dtpEDate.SelectedDate = DateTime.Today;

            lib.UiLoading(sender);
            chkDateSrh.IsChecked = true;
            grbMold.IsHitTestVisible = false;
        }

        #endregion

        #region 검색조건

        //금월
        private void btnThisMonth_Click(object sender, RoutedEventArgs e)
        {
            dtpSDate.SelectedDate = lib.BringThisMonthDatetimeList()[0];
            dtpEDate.SelectedDate = lib.BringThisMonthDatetimeList()[1];
        }

        //금일
        private void btnToday_Click(object sender, RoutedEventArgs e)
        {
            dtpSDate.SelectedDate = DateTime.Today;
            dtpEDate.SelectedDate = DateTime.Today;
        }
        //검색기간 
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
        //금형 로트번호
        private void lblLotNoSrh_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            if (chkLotNoSrh.IsChecked == true) { chkLotNoSrh.IsChecked = false; }
            else { chkLotNoSrh.IsChecked = true; }
        }

        private void chkLotNoSrh_Checked(object sender, RoutedEventArgs e)
        {
            txtLotNoSrh.IsEnabled = true;
            txtLotNoSrh.Focus();
        }

        private void chkLotNoSrh_Unchecked(object sender, RoutedEventArgs e)
        {
            txtLotNoSrh.IsEnabled = false;
        }
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

        //품번
        private void lblBuyerArticleNoSrh_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            if (chkBuyerArticleNoSrh.IsChecked == true) { chkBuyerArticleNoSrh.IsChecked = false; }
            else { chkBuyerArticleNoSrh.IsChecked = true; }
        }

        private void chkBuyerArticleNoSrh_Checked(object sender, RoutedEventArgs e)
        {
            txtBuyerArticleNoSrh.IsEnabled = true;
            btnPfBuyerArticleNoSrh.IsEnabled = true;
            txtBuyerArticleNoSrh.Focus();
        }

        private void chkBuyerArticleNoSrh_Unchecked(object sender, RoutedEventArgs e)
        {
            txtBuyerArticleNoSrh.IsEnabled = false;
            btnPfBuyerArticleNoSrh.IsEnabled = false;
        }
        private void txtBuyerArticleNoSrh_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                MainWindow.pf.ReturnCode(txtBuyerArticleNoSrh, 76, "");
            }
        }

        private void btnPfBuyerArticleNoSrh_Click(object sender, RoutedEventArgs e)
        {
            MainWindow.pf.ReturnCode(txtBuyerArticleNoSrh, 76, "");
        }
        #endregion

        #region 버튼

     
        //닫기
        private void btnClose_Click(object sender, RoutedEventArgs e)
        {
            string stDate = DateTime.Now.ToString("yyyyMMdd");
            string stTime = DateTime.Now.ToString("HHmm");
            DataStore.Instance.InsertLogByFormS(this.GetType().Name, stDate, stTime, "E");
            Lib.Instance.ChildMenuClose(this.ToString());
        }

        //조회
        private void btnSearch_Click(object sender, RoutedEventArgs e)
        {
            FillGrid();
        }

      
        //엑셀
        private void btnExcel_Click(object sender, RoutedEventArgs e)
        {
            DataTable dt = null;
            string Name = string.Empty;

            string[] lst = new string[4];
            lst[0] = "금형이력";
            lst[1] = "금형이력 상세";
            lst[2] = dgdMain.Name;
            lst[3] = dgdSub.Name;

            ExportExcelxaml ExpExc = new ExportExcelxaml(lst);

            ExpExc.ShowDialog();

            if (ExpExc.DialogResult.HasValue)
            {
                if (ExpExc.choice.Equals(dgdMain.Name))
                {
                    if (ExpExc.Check.Equals("Y"))
                        dt = lib.DataGridToDTinHidden(dgdMain);
                    else
                        dt = lib.DataGirdToDataTable(dgdMain);

                    Name = dgdMain.Name;

                    if (lib.GenerateExcel(dt, Name))
                        lib.excel.Visible = true;
                    else
                        return;
                }
                else if (ExpExc.choice.Equals(dgdSub.Name))
                {
                    if (ExpExc.Check.Equals("Y"))
                        dt = lib.DataGridToDTinHidden(dgdSub);
                    else
                        dt = lib.DataGirdToDataTable(dgdSub);

                    Name = dgdSub.Name;

                    if (lib.GenerateExcel(dt, Name))
                        lib.excel.Visible = true;
                    else
                        return;
                }
                else {
                    if (dt != null)
                    {
                        dt.Clear();
                    }
                }
            }
        }


        #endregion

        #region CRUD

        private void FillGrid()
        {
            if (dgdMain.Items.Count > 0)
            {
                dgdMain.Items.Clear();
            }

            try
            {
                string sql = string.Empty;
                DataSet ds = null;

                Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                sqlParameter.Clear();
                sqlParameter.Add("chkDate", chkDateSrh.IsChecked == true ? 1 : 0);
                sqlParameter.Add("FromDate", chkDateSrh.IsChecked == true ? dtpSDate.SelectedDate.Value.ToString("yyyyMMdd") : "");
                sqlParameter.Add("ToDate", chkDateSrh.IsChecked == true ? dtpEDate.SelectedDate.Value.ToString("yyyyMMdd") : "");
                sqlParameter.Add("nchkMold", chkLotNoSrh.IsChecked == true ? 1 : 0);            //금형번호
                sqlParameter.Add("MoldNo", chkLotNoSrh.IsChecked == true ? txtLotNoSrh.Text : "");

                sqlParameter.Add("nchkBuyerArticle", chkBuyerArticleNoSrh.IsChecked == true ? 1 : 0);   //품번
                sqlParameter.Add("BuyerArticle", chkBuyerArticleNoSrh.IsChecked == true ?(txtBuyerArticleNoSrh.Tag != null ? txtBuyerArticleNoSrh.Tag.ToString() : "") : "");
                sqlParameter.Add("chkArticle", chkArticleSrh.IsChecked == true ? 1 : 0);   //품번
                sqlParameter.Add("ArticleID", chkArticleSrh.IsChecked == true ? (txtArticleSrh.Tag != null ? txtArticleSrh.Tag.ToString() : "") : "");

                sqlParameter.Add("nNeedInspect",  0); 
                sqlParameter.Add("nCheckExpired",  0);  
                sqlParameter.Add("ChkIncDisCardYN", "N"); //폐기건 

                ds = DataStore.Instance.ProcedureToDataSet("xp_dvlMold_sMold", sqlParameter, false);

                if (ds != null && ds.Tables.Count > 0)
                {
                    DataTable dt = ds.Tables[0];
                    int i = 0;

                    if (dt.Rows.Count == 0)
                    {
                        MessageBox.Show("조회된 데이터가 없습니다.");
                    }
                    else
                    {

                        DataRowCollection drc = dt.Rows;

                        foreach (DataRow dr in drc)
                        {
                            var WinMolding = new Win_dvl_Molding_U_CodeView()
                            {
                                Num = i + 1,
                                MoldID = dr["MoldID"].ToString(),

                                MoldNo = dr["MoldNo"].ToString(),
                                Article = dr["Article"].ToString(),
                                BuyerArticleNo = dr["BuyerArticleNo"].ToString(),
                                SetProdQty = Convert.ToDouble(dr["SetProdQty"]),
                                HitCount = Convert.ToDouble(dr["HitCount"]),
                               
                            };
                            i++;
                            dgdMain.Items.Add(WinMolding);

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

        private void FillGridSub(string moldID)
        {
            if (dgdSub.Items.Count > 0)
            {
                dgdSub.Items.Clear();
            }

            try
            {
                string sql = string.Empty;
                DataSet ds = null;

                Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                sqlParameter.Clear();
                sqlParameter.Add("MoldID", moldID);

                ds = DataStore.Instance.ProcedureToDataSet("xp_dvlMold_sMoldHitDetail", sqlParameter, false);

                if (ds != null && ds.Tables.Count > 0)
                {
                    DataTable dt = ds.Tables[0];
                    int i = 0;

                    if (dt.Rows.Count == 0)
                    {
                        MessageBox.Show("조회된 데이터가 없습니다.");
                    }
                    else
                    {

                        DataRowCollection drc = dt.Rows;

                        foreach (DataRow dr in drc)
                        {
                            var sub = new HitDetail()
                            {
                                Num = i + 1,
                                MoldID = dr["MoldID"].ToString(),
                                WorkDate = DatePickerFormat(dr["WorkDate"].ToString()),
                                ProcessID = dr["ProcessID"].ToString(),
                                Process = dr["Process"].ToString(),
                                MachineID = dr["MachineID"].ToString(),
                                Machine = dr["Machine"].ToString(),
                                Hitcount = Convert.ToDouble(dr["HitCount"]),
                            };
                            i++;
                            dgdSub.Items.Add(sub);

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

        private void dgdMain_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            var data = dgdMain.SelectedItem as Win_dvl_Molding_U_CodeView;
            if(data != null)
            {
                this.DataContext = data;
                FillGridSub(data.MoldID);
            }
        }

        #region 기타 메서드

        private void TextBoxOnlyNumber_PreviewTextInput(object sender, TextCompositionEventArgs e)
        {
            lib.CheckIsNumeric((TextBox)sender, e);
        }

        // 데이터피커 포맷으로 변경
        private string DatePickerFormat(string str)
        {
            str = str.Trim().Replace("-", "").Replace(".", "");

            if (!str.Equals(""))
            {
                if (str.Length == 8)
                {
                    str = str.Substring(0, 4) + "-" + str.Substring(4, 2) + "-" + str.Substring(6, 2);
                }
            }

            return str;
        }

        #endregion


        public class HitDetail
        {
           public int Num { get; set; }
           public string MoldID { get; set; }
           public string WorkDate { get; set; }
           public string ProcessID { get; set; }
           public string Process { get; set; }
           public string MachineID { get; set; }
           public string Machine { get; set; }
           public double Hitcount { get; set; }
        }

     
    }
}

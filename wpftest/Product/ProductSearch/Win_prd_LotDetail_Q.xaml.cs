using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
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
    /// Win_prd_LotDetail_Q.xaml에 대한 상호 작용 논리
    /// </summary>
    public partial class Win_prd_LotDetail_Q : UserControl
    {
        string stDate = string.Empty;
        string stTime = string.Empty;

        string strLabelFlag = string.Empty;
        Win_prd_LOTDetail_Q_CodeView WinBoxData = new Win_prd_LOTDetail_Q_CodeView();

        public Win_prd_LotDetail_Q()
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
            CheckRBN();
        }

        #region 일자변경
        //생산월
        private void lblDate_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            if (chkDate.IsChecked == true) { chkDate.IsChecked = false; }
            else { chkDate.IsChecked = true; }
        }

        //생산월
        private void chkDate_Checked(object sender, RoutedEventArgs e)
        {
            if (dtpSDate != null && dtpEDate != null)
            {
                dtpSDate.IsEnabled = true;
                dtpEDate.IsEnabled = true;
            }
        }

        //생산월
        private void chkDate_Unchecked(object sender, RoutedEventArgs e)
        {
            dtpSDate.IsEnabled = false;
            dtpEDate.IsEnabled = false;
        }

        //전월
        private void btnLastMonth_Click(object sender, RoutedEventArgs e)
        {
            //dtpSDate.SelectedDate = Lib.Instance.BringLastMonthDatetimeList()[0];
            //dtpEDate.SelectedDate = Lib.Instance.BringLastMonthDatetimeList()[1];

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
        //전일
        private void btnYesterday_Click(object sender, RoutedEventArgs e)
        {
            //dtpSDate.SelectedDate = DateTime.Today.AddDays(-1);
            //dtpEDate.SelectedDate = DateTime.Today.AddDays(-1);

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
        //금월
        private void btnThisMonth_Click(object sender, RoutedEventArgs e)
        {
            dtpSDate.SelectedDate = Lib.Instance.BringThisMonthDatetimeList()[0];
            dtpEDate.SelectedDate = Lib.Instance.BringThisMonthDatetimeList()[1];
        }

     

        //금일
        private void btnToday_Click(object sender, RoutedEventArgs e)
        {
            dtpSDate.SelectedDate = DateTime.Today;
            dtpEDate.SelectedDate = DateTime.Today;
        }



        #endregion

        #region 상단 레이아웃 활성화 & 비활성화
        //자제라벨
        private void lblLabel_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            if (chkLabel.IsChecked == true) { chkLabel.IsChecked = false; }
            else { chkLabel.IsChecked = true; }
        }

        //자제라벨
        private void chkLabel_Checked(object sender, RoutedEventArgs e)
        {
            txtSLabel.IsEnabled = true;
            txtELabel.IsEnabled = true;
        }

        //자제라벨
        private void chkLabel_Unchecked(object sender, RoutedEventArgs e)
        {
            txtSLabel.IsEnabled = false;
            txtELabel.IsEnabled = false;
        }

        //4M 번호
        private void lbl4Mnumber_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            if (chk4Mnumber.IsChecked == true) { chk4Mnumber.IsChecked = false; }
            else { chk4Mnumber.IsChecked = true; }
        }

        //4M 번호
        private void chk4Mnumber_Checked(object sender, RoutedEventArgs e)
        {
            txt4Mnumber.IsEnabled = true;
        }

        //4M 번호
        private void chk4Mnumber_Unchecked(object sender, RoutedEventArgs e)
        {
            txt4Mnumber.IsEnabled = false;
        }

        //제품 RidioButton
        private void rbnArticle_Click(object sender, RoutedEventArgs e)
        {
            CheckRBN();
        }

        //자재로트,하위제품 RidioButton
        private void rbnArticleChild_Click(object sender, RoutedEventArgs e)
        {
            CheckRBN();
        }

        //고객라벨 RidioButton
        private void rbnCustomLabel_Click(object sender, RoutedEventArgs e)
        {
            CheckRBN();
        }

        //제품, 자재로트, 하위제품, 고객라벨 클릭시
        private void CheckRBN()
        {
            if (rbnArticle.IsChecked == true)
            {
                tbkOrder.Text = "Order No.";
                tbkOrderID.Text = "관리번호";
                tbkplID.Text = "작업지시번호";
                strLabelFlag = "2";
                tbkLabel.Text = " 제품 라벨";
            }
            else if (rbnArticleChild.IsChecked == true)
            {
                tbkOrder.Text = "발주명(비고)";
                tbkOrderID.Text = "발주번호";
                tbkplID.Text = "입고번호";
                strLabelFlag = "1";
                tbkLabel.Text = " 자재로트,하위제품 라벨";
            }
            else if (rbnCustomLabel.IsChecked == true)
            {
                tbkOrder.Text = "Order No.";
                tbkOrderID.Text = "관리번호";
                tbkplID.Text = "작업지시번호";
                strLabelFlag = "3";
                tbkLabel.Text = " 고객 라벨";
            }
        }

        //품번
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


        //품번
        private void CheckBoxBuyerArticleNoSearch_Checked(object sender, RoutedEventArgs e)
        {
            TextBoxBuyerArticleNoSearch.IsEnabled = true;
            ButtonBuyerArticleNoSearch.IsEnabled = true;
            TextBoxBuyerArticleNoSearch.Focus();
        }

        //품번
        private void CheckBoxBuyerArticleNoSearch_Unchecked(object sender, RoutedEventArgs e)
        {
            TextBoxBuyerArticleNoSearch.IsEnabled = false;
            ButtonBuyerArticleNoSearch.IsEnabled = false;
        }

        //품번
        private void TextBoxBuyerArticleNoSearch_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                MainWindow.pf.ReturnCode(TextBoxBuyerArticleNoSearch, 76, TextBoxBuyerArticleNoSearch.Text);
            }
        }

        //품번
        private void ButtonBuyerArticleNoSearch_Click(object sender, RoutedEventArgs e)
        {
            MainWindow.pf.ReturnCode(TextBoxBuyerArticleNoSearch, 76, TextBoxBuyerArticleNoSearch.Text);
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

        #endregion

        #region 우측 상단 버튼
        //검색
        private void btnSearch_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                btnSearch.IsEnabled = false;

                using (Loading lw = new Loading(beSearch))
                {
                    lw.ShowDialog();
                }
            }
            catch(Exception ee)
            {
                MessageBox.Show("예외처리 - " + ee.ToString());
            }
            finally
            {
                btnSearch.IsEnabled = true;
            }

            
        }

        private void beSearch()
        {
            FillGridBox();

            if (dgdBoxID.Items.Count > 0)
            {
                dgdBoxID.SelectedIndex = 0;
            }
            else
            {
                this.DataContext = null;

                dgdWork.Items.Clear();
                dgdChild.Items.Clear();

                MessageBox.Show("조회된 데이터가 없습니다.");

                return;
            }
        }

        //닫기
        private void btnClose_Click(object sender, RoutedEventArgs e)
        {
            DataStore.Instance.InsertLogByFormS(this.GetType().Name, stDate, stTime, "E");

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

        //엑셀
        private void btnExcel_Click(object sender, RoutedEventArgs e)
        {
            DataTable dt = null;
            string Name = string.Empty;

            string[] lst = new string[6];
            lst[0] = "BoxID";
            lst[1] = "작업정보";
            lst[2] = "하위결합정보";
            lst[3] = dgdBoxID.Name;
            lst[4] = dgdWork.Name;
            lst[5] = dgdChild.Name;

            ExportExcelxaml ExpExc = new ExportExcelxaml(lst);

            ExpExc.ShowDialog();

            if (ExpExc.DialogResult.HasValue)
            {
                if (ExpExc.choice.Equals(dgdBoxID.Name))
                {
                    DataStore.Instance.InsertLogByForm(this.GetType().Name, "E");
                    if (ExpExc.Check.Equals("Y"))
                        dt = Lib.Instance.DataGridToDTinHidden(dgdBoxID);
                    else
                        dt = Lib.Instance.DataGirdToDataTable(dgdBoxID);

                    Name = dgdBoxID.Name;

                    if (Lib.Instance.GenerateExcel(dt, Name))
                        Lib.Instance.excel.Visible = true;
                    else
                        return;
                }
                else if (ExpExc.choice.Equals(dgdWork.Name))
                {
                    if (ExpExc.Check.Equals("Y"))
                        dt = Lib.Instance.DataGridToDTinHidden(dgdWork);
                    else
                        dt = Lib.Instance.DataGirdToDataTable(dgdWork);

                    Name = dgdWork.Name;

                    if (Lib.Instance.GenerateExcel(dt, Name))
                        Lib.Instance.excel.Visible = true;
                    else
                        return;
                }
                else if (ExpExc.choice.Equals(dgdChild.Name))
                {
                    if (ExpExc.Check.Equals("Y"))
                        dt = Lib.Instance.DataGridToDTinHidden(dgdChild);
                    else
                        dt = Lib.Instance.DataGirdToDataTable(dgdChild);

                    Name = dgdChild.Name;

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

        #region FillGridBox 조회
        //Box List
        private void FillGridBox()
        {
            if (dgdBoxID.Items.Count > 0)
            {
                dgdBoxID.Items.Clear();
            }

            try
            {
                Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                sqlParameter.Add("sLabelGbn", strLabelFlag);
                sqlParameter.Add("chkDate", chkDate.IsChecked == true && chkLabel.IsChecked == false ? 1 : 0);
                sqlParameter.Add("sFromDate", chkDate.IsChecked == true ? dtpSDate.SelectedDate.Value.ToString("yyyyMMdd") : "");
                sqlParameter.Add("sToDate", chkDate.IsChecked == true ? dtpEDate.SelectedDate.Value.ToString("yyyyMMdd") : "");
                sqlParameter.Add("chkLabelID", chkLabel.IsChecked == true ? 1 : 0);
                sqlParameter.Add("sFromLabelID", chkLabel.IsChecked == true ? txtSLabel.Text : "");
                sqlParameter.Add("sTOLabelID", chkLabel.IsChecked == true ? txtELabel.Text : "");
                sqlParameter.Add("chk4MID", chk4Mnumber.IsChecked == true ? 1 : 0);
                sqlParameter.Add("s4MNo", chk4Mnumber.IsChecked == true ? txt4Mnumber.Text : "");
                sqlParameter.Add("chkArticleID", chkArticle.IsChecked == true ? 1 : 0);
                sqlParameter.Add("ArticleID", chkArticle.IsChecked == true ? txtArticle.Tag.ToString() : "");
                sqlParameter.Add("chkBuyerArticleNo", CheckBoxBuyerArticleNoSearch.IsChecked == true ? 1 : 0);
                sqlParameter.Add("BuyerArticleNo", CheckBoxBuyerArticleNoSearch.IsChecked == true ? TextBoxBuyerArticleNoSearch.Tag.ToString() : "");

                DataSet ds = DataStore.Instance.ProcedureToDataSet_LogWrite("xp_prd_sLabelIDList", sqlParameter, true, "R");

                if (ds != null && ds.Tables.Count > 0)
                {
                    DataTable dt = ds.Tables[0];

                    if (dt.Rows.Count > 0)
                    {
                        int i = 0;
                        DataRowCollection drc = dt.Rows;

                        foreach (DataRow dr in drc)
                        {
                            var WinBoxID = new Win_prd_LOTDetail_Q_CodeView()
                            {
                                Num = i + 1,
                                LabelID = dr["LabelID"].ToString(),
                                RelLabelID = dr["RelLabelID"].ToString(),
                                Article = dr["Article"].ToString(),
                                BuyerArticleNo = dr["BuyerArticleNo"].ToString(),
                                WorkDate = dr["workdate"].ToString()
                            };

                            dgdBoxID.Items.Add(WinBoxID);
                            i++;
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

        #endregion

        #region dgdBoxID_SelectionChanged
        //
        private void dgdBoxID_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            WinBoxData = dgdBoxID.SelectedItem as Win_prd_LOTDetail_Q_CodeView;

            if (WinBoxData != null)
            {
                FillGridWorkData(WinBoxData.LabelID.Replace(" ", ""));
                FillGridChildData(WinBoxData.LabelID.Replace(" ", ""));
                FillText(WinBoxData.LabelID.Replace(" ", ""));
            }
        }

        #endregion

        #region 작업정보 조회
        //작업정보
        private void FillGridWorkData(string strLabelID)
        {
            if (dgdWork.Items.Count > 0)
            {
                dgdWork.Items.Clear();
            }

            try
            {
                Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                sqlParameter.Add("sLabelID", strLabelID);

                DataSet ds = DataStore.Instance.ProcedureToDataSet("xp_prd_sLabelIDOneProcess", sqlParameter, false);

                if (ds != null && ds.Tables.Count > 0)
                {
                    DataTable dt = ds.Tables[0];
                    string sql = string.Empty;

                    if (dt.Rows.Count > 0)
                    {

                        int i = 0;

                        foreach (DataRow dr in dt.Select(sql))
                        {
                            i++;
                            var WinWorker = new Win_prd_LOTDetail_Q_Work_CodeView()
                            {
                                Num = i,
                                WorkDate = dr["WorkDate"].ToString(),
                                WorkTime = dr["WorkTime"].ToString(),
                                WorkQty = dr["WokQty"].ToString(),
                                PersonID = dr["PersonID"].ToString(),
                                WorkManName = dr["WorkManName"].ToString(),
                                ProcessID = dr["ProcessID"].ToString(),
                                MachineID = dr["MachineID"].ToString(),
                                Process = dr["Process"].ToString(),
                                Machine = dr["Machine"].ToString(),
                                HrLicence = dr["HrLicence"].ToString(),
                                Remark = dr["Remark"].ToString(),
                                DefectList = dr["DefectList"].ToString(),
                            };

                            if (WinWorker.WorkTime.Length == 5)
                            {
                                WinWorker.WorkTime = Lib.Instance.SixLengthTime("0" + WinWorker.WorkTime);
                            }
                            else if (WinWorker.WorkTime.Length == 6)
                            {
                                WinWorker.WorkTime = Lib.Instance.SixLengthTime(WinWorker.WorkTime);
                            }

                            WinWorker.WorkDate = Lib.Instance.StrDateTimeBar(WinWorker.WorkDate);
                            WinWorker.WorkQty = Lib.Instance.returnNumStringZero(WinWorker.WorkQty);
                            dgdWork.Items.Add(WinWorker);
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

            if (dgdWork.Items.Count > 0)
            {
                dgdWork.SelectedIndex = 0;
            }
        }
        #endregion

        #region 하위결합정보 조회
        //하위결합정보
        private void FillGridChildData(string strLabelID)
        {
            if (dgdChild.Items.Count > 0)
            {
                dgdChild.Items.Clear();
            }

            try
            {
                Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                sqlParameter.Add("sLabelID", strLabelID);

                DataSet ds = DataStore.Instance.ProcedureToDataSet("xp_prd_sLabelIDChild", sqlParameter, false);

                if (ds != null && ds.Tables.Count > 0)
                {
                    int i = 0;
                    DataTable dt = ds.Tables[0];
                    string sql = string.Empty;

                    if (dt.Rows.Count > 0)
                    {
                        //if (chkDate.IsChecked == true)
                        //{
                        //    sql += " InDate  >= " + dtpSDate.SelectedDate.Value.ToString("yyyyMMdd") +
                        //        " and  InDate <= " + dtpEDate.SelectedDate.Value.ToString("yyyyMMdd");
                        //}

                        //DataRowCollection drc = dt.Rows;

                        foreach (DataRow dr in dt.Select(sql))
                        {
                            i++;
                            var WinChild = new Win_prd_LOTDetail_Q_Child_CodeView()
                            {
                                Num = i,
                                ChildLabelID = dr["ChildLabelID"].ToString(),
                                ChildArticleID = dr["ChildArticleID"].ToString(),
                                Article = dr["Article"].ToString().Trim(),
                                inPersonID = dr["inPersonID"].ToString(),
                                InDate = dr["InDate"].ToString(),
                                InTime = dr["InTime"].ToString(),
                                InPersonName = dr["InPersonName"].ToString(),
                                InspectDate = dr["InspectDate"].ToString(),
                                //Inspecttime = dr["Inspecttime"].ToString(),
                                Gubun = dr["Gubun"].ToString(),
                                InspectPersonName = dr["InspectPersonName"].ToString(),
                                BuyerArticleNo = dr["BuyerArticleNo"].ToString().Trim(),
                                Qty = stringFormatN0(dr["Qty"]),
                                Custom = dr["Custom"].ToString().Trim(),
                                StockQty = dr["StockQty"].ToString().Trim().Equals("0") ? "" : stringFormatN0(dr["StockQty"]),
                            };

                            if (WinChild.InTime.Length == 5)
                            {
                                WinChild.InTime = Lib.Instance.SixLengthTime("0" + WinChild.InTime);
                            }
                            else if (WinChild.InTime.Length == 6)
                            {
                                WinChild.InTime = Lib.Instance.SixLengthTime(WinChild.InTime);
                            }

                            WinChild.InspectDate = Lib.Instance.StrDateTimeBar(WinChild.InspectDate);
                            WinChild.InDate = Lib.Instance.StrDateTimeBar(WinChild.InDate);
                            dgdChild.Items.Add(WinChild);
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

            if (dgdChild.Items.Count > 0)
            {
                dgdChild.SelectedIndex = 0;
            }
        }

        #endregion

        #region 하단 레이아웃 텍스트 박스에 들어갈 데이터 조회
        //ShowData
        private void FillText(string strLabelID)
        {
            if (DataContext != null)
            {
                DataContext = null;
            }

            try
            {
                Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                sqlParameter.Add("sLabelID", strLabelID);

                DataSet ds = DataStore.Instance.ProcedureToDataSet("xp_trc_sLabelIDOneDetail", sqlParameter, false);

                if (ds != null && ds.Tables.Count > 0)
                {
                    DataTable dt = ds.Tables[0];

                    if (dt.Rows.Count > 0)
                    {
                        DataRow dr = dt.Rows[0];

                        var WinShow = new Win_prd_LOTDetail_Q_ShowText_CodeView()
                        {
                            ArticleID = dr["ArticleID"].ToString(),
                            Article = dr["Article"].ToString(),
                            BuyerArticleNo = dr["BuyerArticleNo"].ToString(),
                            ProdQtyPerBox = dr["ProdQtyPerBox"].ToString(),
                            Spec = dr["Spec"].ToString(),
                            OrderID = dr["OrderID"].ToString(),
                            OrderNo = dr["OrderNo"].ToString(),
                            InstID = dr["InstID"].ToString(),
                            KCustom = dr["KCustom"].ToString(),
                            CustomBoxID = dr["CustomBoxID"].ToString(),
                            FourMID = dr["4MID"].ToString()
                        };

                        WinShow.ProdQtyPerBox = Lib.Instance.returnNumStringZero(WinShow.ProdQtyPerBox);
                        WinShow.LabelID = strLabelID;

                        this.DataContext = WinShow;
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
        #endregion

        #region 기타 메서드 모음

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

        #endregion

    }

    #region CodeView
    class Win_prd_LOTDetail_Q_CodeView : BaseView
    {
        public override string ToString()
        {
            return (this.ReportAllProperties());
        }

        public int Num { get; set; }
        public string LabelID { get; set; }
        public string RelLabelID { get; set; }
        public string Article { get; set; }
        public string BuyerArticleNo { get; set; }
        public string WorkDate { get; set; }
    }

    class Win_prd_LOTDetail_Q_Work_CodeView : BaseView
    {
        public override string ToString()
        {
            return (this.ReportAllProperties());
        }

        public int Num { get; set; }
        public string WorkDate { get; set; }
        public string WorkTime { get; set; }
        public string WorkQty { get; set; }
        public string PersonID { get; set; }
        public string WorkManName { get; set; }
        public string ProcessID { get; set; }
        public string MachineID { get; set; }
        public string Process { get; set; }
        public string Machine { get; set; }
        public string HrLicence { get; set; }
        public string Remark { get; set; }
        public string DefectList { get; set; }
    }

    class Win_prd_LOTDetail_Q_Child_CodeView : BaseView
    {
        public override string ToString()
        {
            return (this.ReportAllProperties());
        }

        public int Num { get; set; }
        public string ChildLabelID { get; set; }
        public string ChildArticleID { get; set; }
        public string Article { get; set; }
        public string inPersonID { get; set; }
        public string InDate { get; set; }
        public string InTime { get; set; }
        public string InPersonName { get; set; }
        public string InspectDate { get; set; }
        public string Inspecttime { get; set; }
        public string Gubun { get; set; }
        public string InspectPersonName { get; set; }
        public string BuyerArticleNo { get; set; }
        public string Qty { get; set; }
        public string Custom { get; set; }
        public string StockQty { get; set; }
    }

    class Win_prd_LOTDetail_Q_ShowText_CodeView : BaseView
    {
        public override string ToString()
        {
            return (this.ReportAllProperties());
        }

        public string LabelID { get; set; }
        public string ArticleID { get; set; }
        public string Article { get; set; }
        public string BuyerArticleNo { get; set; }
        public string ProdQtyPerBox { get; set; }
        public string Spec { get; set; }
        public string OrderID { get; set; }
        public string OrderNo { get; set; }
        public string InstID { get; set; }
        public string KCustom { get; set; }
        public string CustomBoxID { get; set; }
        public string FourMID { get; set; }
    }
    #endregion

}

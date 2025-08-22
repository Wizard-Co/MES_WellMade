/**
 * 
 * @details 생산계획편성 작성
 * @author 김수정
 * @date 
 * @version 1.0
 * 
 * @section MODIFYINFO 수정정보
 * - 수정일        - 수정자       : 수정내역
 * 
 * 
 * */

using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Data;
using System.Text.RegularExpressions;
using System.Threading;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using WizMes_BooKyong.PopUp;
using WizMes_BooKyong.PopUP;

namespace WizMes_BooKyong
{
    /// <summary>
    /// Win_prd_PlanInputAuto.xaml에 대한 상호 작용 논리
    /// </summary>
    public partial class Win_prd_PlanInputAuto : UserControl
    {
        Win_prd_PlanInputAuto_CodeView Auto = new Win_prd_PlanInputAuto_CodeView();
        List<Win_prd_PlanInputAuto_CodeView> lstAutoPlan = new List<Win_prd_PlanInputAuto_CodeView>();


        private static Win_prd_PlanInputAuto _instance = null;


        public static Win_prd_PlanInputAuto planAuto
        {
            get
            {
                if (_instance  == null)
                    _instance = new Win_prd_PlanInputAuto();

                return _instance;
            }
        }

        

        Lib lib = new Lib();
        int rowNum = 0;
        string stDate = string.Empty;
        string stTime = string.Empty;

        public Win_prd_PlanInputAuto()
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
    
            _instance = this;

            dgdMain.SelectedIndex = rowNum;
        }


        #region 일자변경

        //편성일자
        private void lblDate_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            if (chkDate.IsChecked == true) { chkDate.IsChecked = false; }
            else { chkDate.IsChecked = true; }
        }

        //편성일자
        private void chkDate_Checked(object sender, RoutedEventArgs e)
        {
            if (dtpSDate != null && dtpEDate != null)
            {
                dtpSDate.IsEnabled = true;
                dtpEDate.IsEnabled = true;
            }
        }

        //편성일자
        private void chkDate_Unchecked(object sender, RoutedEventArgs e)
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

        #endregion

        #region 상단 검색조건

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

        //품명 → 품번으로 변경(GLS)
        private void txtArticle_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                MainWindow.pf.ReturnCode(txtArticle, 76, "");
            }
        }

        //품명 → 품번으로 변경(GLS)
        private void btnPfArticle_Click(object sender, RoutedEventArgs e)
        {
            MainWindow.pf.ReturnCode(txtArticle, 76, "");
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

        //최종거래처
        private void lblEndCustom_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            if (chkCustom.IsChecked == true) { chkCustom.IsChecked = false; }
            else { chkCustom.IsChecked = true; }
        }

        //최종거래처
        private void chkEndCustom_Checked(object sender, RoutedEventArgs e)
        {
            txtCustom.IsEnabled = true;
            btnPfCustom.IsEnabled = true;
            txtCustom.Focus();
        }

        //최종거래처
        private void chkEndCustom_Unchecked(object sender, RoutedEventArgs e)
        {
            txtCustom.IsEnabled = false;
            btnPfCustom.IsEnabled = false;
        }

        //최종거래처
        private void txtEndCustom_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                MainWindow.pf.ReturnCode(txtCustom, (int)Defind_CodeFind.DCF_CUSTOM, "");
            }
        }

        //최종거래처
        private void btnEndPfCustom_Click(object sender, RoutedEventArgs e)
        {
            MainWindow.pf.ReturnCode(txtCustom, (int)Defind_CodeFind.DCF_CUSTOM, "");
        }

        //관리번호
        private void lblOrderIDSrh_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            if (chkOrderIDSrh.IsChecked == true) { chkOrderIDSrh.IsChecked = false; }
            else { chkOrderIDSrh.IsChecked = true; }
        }

        //관리번호
        private void chkOrderIDSrh_Checked(object sender, RoutedEventArgs e)
        {
            txtOrderIDSrh.IsEnabled = true;
        }

        //관리번호
        private void chkOrderIDSrh_Unchecked(object sender, RoutedEventArgs e)
        {
            txtOrderIDSrh.IsEnabled = false;

        }

        //OrderNo
        private void txtOrderID_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                if (e.Key == Key.Enter)
                {
                    MainWindow.pf.ReturnCode(txtOrderIDSrh, (int)Defind_CodeFind.DCF_ORDER, "");
                }
            }
        }

        //OrderNo
        private void rbnOrderNo_Click(object sender, RoutedEventArgs e)
        {
            Check_bdrOrder();
        }

        //관리번호
        private void rbnOrderID_Click(object sender, RoutedEventArgs e)
        {
            Check_bdrOrder();
        }

        //OrderNo, 관리번호 체크 
        private void Check_bdrOrder()
        {
            if (rbnOrderID.IsChecked == true)
            {
                tbkOrderSrh.Text = "관리번호";
            }
            else if (rbnOrderNo.IsChecked == true)
            {
                tbkOrderSrh.Text = "Order No.";
            }
        }

        //데이터피커 편성 일자 체크
        private void CheckDate(object sender, RoutedEventArgs e)
        {
            if (dtpSDate.SelectedDate != null || dtpEDate.SelectedDate != null)
            {
                if (dtpEDate.SelectedDate < dtpSDate.SelectedDate)
                {
                    MessageBox.Show("검색일자의 종료일자가 시작날짜보다 작을 수 없습니다");

                    dtpEDate.SelectedDate = dtpSDate.SelectedDate.Value.AddDays(1);

                }
            }
        }
 

        #endregion

        #region 생산계획편성 
        
        private void btnAutoPlan_Click(object sender, RoutedEventArgs e)
        {
            Win_pop_AutoPlan AutoPlan = new Win_pop_AutoPlan();

            AutoPlan.ShowDialog();

        }
        #endregion

       

        #region 우측 상단 버튼
        //검색
        public void btnSearch_Click(object sender, RoutedEventArgs e)
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



            //Dispatcher.BeginInvoke(new Action(() =>

            //{
            //    btnSearch.IsEnabled = true;
                

            //}), System.Windows.Threading.DispatcherPriority.Background);
        }

        //닫기
        private void btnClose_Click(object sender, RoutedEventArgs e)
        {
            DataStore.Instance.InsertLogByFormS(this.GetType().Name, stDate, stTime, "E");
            Lib.Instance.ChildMenuClose(this.ToString());
        }

        //삭제
        private void btnDelete_Click(object sender, RoutedEventArgs e)
        {
            var Auto = dgdMain.SelectedItem as Win_prd_PlanInputAuto_CodeView;

            if (Auto == null)
            {
                MessageBox.Show("삭제할 데이터가 지정되지 않았습니다. 삭제데이터를 지정하고 눌러주세요");
            }
            else
            {
                if (MessageBox.Show("선택하신 항목을 삭제하시겠습니까?", "삭제 전 확인", MessageBoxButton.YesNo) == MessageBoxResult.Yes)
                {
                    if (dgdMain.Items.Count > 0 && dgdMain.SelectedItem != null)
                    {
                        rowNum = dgdMain.SelectedIndex;
                    }

                    if (DeleteData(Auto.AutoID))
                    {
                        rowNum -= 1;
                        re_Search(rowNum);
                    }
                }
            }
        }

        //엑셀
        private void btnExcel_Click(object sender, RoutedEventArgs e)
        {
            DataTable dt = null;
            string Name = string.Empty;

            string[] lst = new string[2];
            lst[0] = "생산계획 편성";
            lst[1] = dgdMain.Name;

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

        #region 조회
        public void FillGrid()
        {
            if (dgdMain.Items.Count > 0)
            {
                dgdMain.Items.Clear();
            }
            if (dgdSub.Items.Count > 0)
            {
                dgdSub.Items.Clear();
            }

            try
            {
                Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                sqlParameter.Clear();
                sqlParameter.Add("ChkDate", chkDate.IsChecked == true ? 1 : 0);
                sqlParameter.Add("SDate", chkDate.IsChecked == true ? (dtpSDate.SelectedDate == null ? "" : dtpSDate.SelectedDate.Value.ToString("yyyyMMdd")) : "");
                sqlParameter.Add("EDate", chkDate.IsChecked == true ? (dtpEDate.SelectedDate == null ? "" : dtpEDate.SelectedDate.Value.ToString("yyyyMMdd")) : "");
                sqlParameter.Add("ChkEndCustom", chkEndCustom.IsChecked == true ? 1 : 0 );
                sqlParameter.Add("EndCustomID", chkEndCustom.IsChecked == true ? ((txtEndCustom.Text != null && txtEndCustom.Tag != null) ? txtEndCustom.Text.ToString() : "" ) : "");
                sqlParameter.Add("ChkCustomID", chkCustom.IsChecked == true ? 1 : 0 );
                sqlParameter.Add("CustomID", chkCustom.IsChecked == true ? ((txtCustom.Text != null && txtCustom.Tag != null) ? txtCustom.Text.ToString() : "") : "");
                sqlParameter.Add("ChkArticleID", chkArticle.IsChecked == true ? 1 : 0);
                sqlParameter.Add("ArticleID", chkArticle.IsChecked == true ? ((txtArticle.Text != null && txtArticle.Tag != null) ? txtArticle.Text.ToString() : "") : "");
                sqlParameter.Add("ChkOrderID", chkOrderIDSrh.IsChecked == true ? 1 : 0);
                sqlParameter.Add("OrderID",  chkOrderIDSrh.IsChecked == true ? ((txtOrderIDSrh.Text != null && txtOrderIDSrh.Tag != null) ? txtOrderIDSrh.Text.ToString() : "" ) : "");

                DataSet ds = DataStore.Instance.ProcedureToDataSet_LogWrite("xp_AutoPlan_sAutoPlan", sqlParameter, true, "R");

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

                            var Auto = new Win_prd_PlanInputAuto_CodeView()
                            {
                                Num = i,
                                AutoDate = dr["AutoDate"].ToString(),
                                AutoID = dr["AutoID"].ToString(),
                                AcptDate = dr["AcptDate"].ToString(),
                                ArticleID = dr["ArticleID"].ToString(),
                                Article = dr["Article"].ToString(),
                                BuyerArticleNo = dr["BuyerArticleNo"].ToString(),
                                OrderID = dr["OrderID"].ToString(),
                                OrderQty = Convert.ToDouble(dr["OrderQty"]),
                                DvlyDate = dr["DvlyDate"].ToString(),
                                SumInstQty = Convert.ToDouble(dr["SumInstQty"]),
                                OrderInstQty = Convert.ToDouble(dr["OrderInstQty"]),
                                NonePlanQty = Convert.ToDouble(dr["NonePlanQty"]),
                                CustomID = dr["CustomID"].ToString(),
                                

                            };

                            dgdMain.Items.Add(Auto);
                        }
                    //    tbkCount.Text = "▶ 검색결과 : " + i.ToString() + " 건";
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

        #region 서브그리드 조회

        private void FillGridSub(string strAutoID)
        {
            if(dgdSub.Items.Count > 0)
            {
                dgdSub.Items.Clear();
            }
            try
            {
                Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                sqlParameter.Clear();
                sqlParameter.Add("AutoID", strAutoID != null ? strAutoID.ToString() : "");
               
                DataSet ds = DataStore.Instance.ProcedureToDataSet_LogWrite("xp_AutoPlan_sAutoPlanDet", sqlParameter, true, "R");

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

                            var AutoSub = new Win_prd_PlanInputAuto_Sub_CodeView()
                            {
                                OrderID = dr["OrderID"].ToString(),
                                BuyerArticleNo = dr["BuyerArticleNo"].ToString(),
                                ArticleID = dr["ArticleID"].ToString(),
                                Article = dr["Article"].ToString(),
                                ProcessID = dr["ProcessID"].ToString(),
                                Process = dr["Process"].ToString(),
                                OrderQty = Convert.ToDouble(dr["OrderQty"]),
                                InstQty = Convert.ToDouble(dr["InstQty"]),
                                StartDate = dr["StartDate"].ToString(),
                                EndDate = dr["EndDate"].ToString(),

                            };

                            dgdSub.Items.Add(AutoSub);
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


        #region 삭제
        private bool DeleteData(string AutoID)
        {
            bool flag = false;

            Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
            sqlParameter.Clear();
            sqlParameter.Add("AutoID", AutoID);
            sqlParameter.Add("OutMessage", "");


            string[] result = DataStore.Instance.ExecuteProcedure_NewLog("xp_AutoPlan_dAutoPlan", sqlParameter, "D");
            DataStore.Instance.CloseConnection();

            if (result[0].Equals("success"))
            {
                MessageBox.Show("삭제가 완료되었습니다");
                flag = true;
            }

            return flag;
        }
        #endregion

        #region 재검색
        /// <summary>
        /// 재검색(수정,삭제,추가 저장후에 자동 재검색)
        /// </summary>
        /// <param name="selectedIndex"></param>
        private void re_Search(int rowNum)
        {

            FillGrid();

            if (dgdMain.Items.Count > 0)
            {
                dgdMain.SelectedIndex = rowNum;
            }
            else
            {
                MessageBox.Show("조회된 데이터가 없습니다.");
                return;
            }
        }

        #endregion

        #region DgdMain_SelectionChanged
        private void DgdMain_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            Auto = dgdMain.SelectedItem as Win_prd_PlanInputAuto_CodeView;

            if (Auto != null)
            {
                this.DataContext = Auto;
                FillGridSub(Auto.AutoID);

            }
        }


        #endregion

        
    }

    #region CodeView
    public class Win_prd_PlanInputAuto_CodeView
    {
        public int Num { get; set; }
        public string AutoDate { get; set; }            //편성일자
        public string AutoID { get; set; }              //편성번호  
        public string AcptDate { get; set; }            //수주일자
        public string ArticleID { get; set; }           //품목ID
        public string Article { get; set; }             //품명
        public string BuyerArticleNo { get; set; }      //품번
        public string OrderID { get; set; }             //오더ID  
        public double OrderQty { get; set; }            //수주량
        public double OrderInstQty { get; set; }        //계획량
        public string DvlyDate { get; set; }            //납기일
        public double SumInstQty { get; set; }          //누계계획량
        public double NonePlanQty { get; set; }         //미계획량
        public string CustomID { get; set; }


    }

    public class Win_prd_PlanInputAuto_Sub_CodeView
    {
        public string OrderID { get; set; }
        public string BuyerArticleNo { get; set; }
        public string ArticleID { get; set; }
        public string Article { get; set; }
        public string ProcessID { get; set; }
        public string Process { get; set; }
        public double OrderQty { get; set; }
        public double InstQty { get; set; }
        public string StartDate { get; set; }
        public string EndDate { get; set; }
    }

    #endregion
}

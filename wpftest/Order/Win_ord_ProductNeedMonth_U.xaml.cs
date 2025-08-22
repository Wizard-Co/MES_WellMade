using System;
using System.Collections.Generic;
using System.Data;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using WizMes_WellMade.PopUp;
using WizMes_WellMade.PopUP;
using Excel = Microsoft.Office.Interop.Excel;

/**************************************************************************************************
'** 프로그램명 : Win_ord_ProductNeedMonth_U
'** 설명       : 월단위 제품 소요조회
'** 작성일자   : 2023.03.30
'** 작성자     : 장시영
'**------------------------------------------------------------------------------------------------
'**************************************************************************************************
' 변경일자  , 변경자, 요청자    , 요구사항ID      , 요청 및 작업내용
'**************************************************************************************************
' 2023.03.30, 장시영, 신규 화면 생성
'**************************************************************************************************/

namespace WizMes_WellMade
{
    /// <summary>
    /// Win_ord_ProductNeedMonth_U.xaml에 대한 상호 작용 논리
    /// </summary>
    public partial class Win_ord_ProductNeedMonth_U : UserControl
    {
        string stDate = string.Empty;
        string stTime = string.Empty;

        string strFlag = string.Empty;
        int rowNum = 0;

        Lib lib = new Lib();        

        // 메인 그리드 선택시 사용
        List<Win_ord_ProductNeedMonth_U_View> listCheckGrid = new List<Win_ord_ProductNeedMonth_U_View>();

        public Win_ord_ProductNeedMonth_U()
        {
            InitializeComponent();
        }

        private void UserControl_Loaded(object sender, RoutedEventArgs e)
        {
            stDate = DateTime.Now.ToString("yyyyMMdd");
            stTime = DateTime.Now.ToString("HHmm");

            DataStore.Instance.InsertLogByFormS(this.GetType().Name, stDate, stTime, "S");

            Lib.Instance.UiLoading(sender);
            btnThisMonth_Click(null, null);
        }

        #region 상단컨트롤
        #region 날짜컨트롤
        // 전월
        private void btnLastMonth_Click(object sender, RoutedEventArgs e) { search_BtnDate_Control(1); }
        // 금월
        private void btnThisMonth_Click(object sender, RoutedEventArgs e) { search_BtnDate_Control(2); }
        #endregion 날짜컨트롤

        #region 검색컨트롤
        // 거래처
        private void lblCustom_MouseLeftButtonUp(object sender, MouseButtonEventArgs e) { search_CheckBox_Control(chkCustom); }
        private void chkCustom_Checked(object sender, RoutedEventArgs e) { search_CheckBox_Checked_Control(true, txtCustom, btnPfCustom); }
        private void chkCustom_Unchecked(object sender, RoutedEventArgs e) { search_CheckBox_Checked_Control(false, txtCustom, btnPfCustom); }
        private void txtCustom_KeyDown(object sender, KeyEventArgs e) { if (e.Key == Key.Enter) search_PlusFinder_Control(txtCustom, (int)Defind_CodeFind.DCF_CUSTOM, ""); }
        private void btnPfCustom_Click(object sender, RoutedEventArgs e) { search_PlusFinder_Control(txtCustom, (int)Defind_CodeFind.DCF_CUSTOM, ""); }

        // 품번
        private void lblBuyerArticleNo_MouseLeftButtonUp(object sender, MouseButtonEventArgs e) { search_CheckBox_Control(chkBuyerArticleNo); }
        private void chkBuyerArticleNo_Checked(object sender, RoutedEventArgs e) { search_CheckBox_Checked_Control(true, txtBuyerArticleNo, btnPfBuyerArticleNo); }
        private void chkBuyerArticleNo_Unchecked(object sender, RoutedEventArgs e) { search_CheckBox_Checked_Control(false, txtBuyerArticleNo, btnPfBuyerArticleNo); }
        private void txtBuyerArticleNo_KeyDown(object sender, KeyEventArgs e) { if (e.Key == Key.Enter) search_PlusFinder_Control(txtBuyerArticleNo, 76, ""); }
        private void btnPfBuyerArticleNo_Click(object sender, RoutedEventArgs e) { search_PlusFinder_Control(txtBuyerArticleNo, 76, ""); }

        // 품명
        private void lblArticle_MouseLeftButtonUp(object sender, MouseButtonEventArgs e) { search_CheckBox_Control(chkArticle); }
        private void chkArticle_Checked(object sender, RoutedEventArgs e) { search_CheckBox_Checked_Control(true, txtArticle, btnPfArticle); }
        private void chkArticle_Unchecked(object sender, RoutedEventArgs e) { search_CheckBox_Checked_Control(false, txtArticle, btnPfArticle); }
        private void txtArticle_KeyDown(object sender, KeyEventArgs e) { if (e.Key == Key.Enter) search_PlusFinder_Control(txtArticle, 77, ""); }
        private void btnPfArticle_Click(object sender, RoutedEventArgs e) { search_PlusFinder_Control(txtArticle, 77, ""); }
        #endregion 검색컨트롤

        #region 버튼컨트롤
        // 추가
        private void btnAdd_Click(object sender, RoutedEventArgs e)
        {
            strFlag = "I";
            DataContext = null;

            rowNum = dgdMain.SelectedIndex > -1 ? dgdMain.SelectedIndex : 0;
            dgdSub.Items.Clear();
        }

        // 수정
        private void btnUpdate_Click(object sender, RoutedEventArgs e)
        {
            Win_ord_ProductNeedMonth_U_View ReqView = dgdMain.SelectedItem as Win_ord_ProductNeedMonth_U_View;
            if (ReqView != null)
            {
                strFlag = "U";
                rowNum = dgdMain.SelectedIndex;
            }
        }

        // 삭제
        private void btnDelete_Click(object sender, RoutedEventArgs e)
        {
        }

        // 닫기
        private void btnClose_Click(object sender, RoutedEventArgs e)
        {
            DataStore.Instance.InsertLogByFormS(this.GetType().Name, stDate, stTime, "E");
            Lib.Instance.ChildMenuClose(this.ToString());
        }

        // 검색
        private void btnSearch_Click(object sender, RoutedEventArgs e)
        {
            using (Loading lw = new Loading(beSearch))
            {
                lw.ShowDialog();
            }
        }

        // 저장
        private void btnSave_Click(object sender, RoutedEventArgs e)
        {
        }

        // 취소
        private void btnCancel_Click(object sender, RoutedEventArgs e)
        {
            if (strFlag.Equals("U"))
                re_Search();
            else
            {
                rowNum = 0;
                re_Search();
            }

            strFlag = string.Empty;
        }

        string upload_fileName = "";

        private void btnUpload_Click(object sender, RoutedEventArgs e)
        {
            Microsoft.Win32.OpenFileDialog file = new Microsoft.Win32.OpenFileDialog();
            file.Filter = "Excel files (*.xls,*xlsx)|*.xls;*xlsx|All files (*.*)|*.*";
            file.InitialDirectory = "C:\\";

            if (file.ShowDialog() == true)
            {
                upload_fileName = file.FileName;

                btnUpload.IsEnabled = false;

                using (Loading ld = new Loading("excel", beUpload))
                {
                    ld.ShowDialog();
                }

                re_Search();

                btnUpload.IsEnabled = true;
            }
        }

        private void beUpload()
        {
            Lib lib2 = new Lib();

            Excel.Application excelapp = null;
            Excel.Workbook workbook = null;
            Excel.Worksheet worksheet = null;
            Excel.Range workrange = null;

            List<OrderExcel> listExcel = new List<OrderExcel>();

            try
            {
                excelapp = new Excel.Application();
                workbook = excelapp.Workbooks.Add(upload_fileName);
                worksheet = workbook.Sheets["Sheet"];
                workrange = worksheet.UsedRange;

                for (int row = 3; row <= workrange.Rows.Count; row++)
                {
                    OrderExcel excel = new OrderExcel();
                    excel.CustomID = workrange.get_Range("A" + row.ToString()).Value2;
                    excel.Model = workrange.get_Range("B" + row.ToString()).Value2;
                    excel.BuyerArticleNo = workrange.get_Range("C" + row.ToString()).Value2;
                    excel.Article = workrange.get_Range("D" + row.ToString()).Value2;
                    excel.UnitClss = workrange.get_Range("E" + row.ToString()).Value2;

                    object objOrderQty = workrange.get_Range("H" + row.ToString()).Value2;
                    if (objOrderQty != null)
                        excel.OrderQty = objOrderQty.ToString();

                    if (!string.IsNullOrEmpty(excel.CustomID)
                        && !string.IsNullOrEmpty(excel.BuyerArticleNo) && !string.IsNullOrEmpty(excel.Article)
                        && !string.IsNullOrEmpty(excel.UnitClss) && !string.IsNullOrEmpty(excel.OrderQty))
                    {
                        listExcel.Add(excel);
                    }

                    if (string.IsNullOrEmpty(excel.CustomID) && string.IsNullOrEmpty(excel.Model)
                        && string.IsNullOrEmpty(excel.BuyerArticleNo) && string.IsNullOrEmpty(excel.Article)
                        && string.IsNullOrEmpty(excel.UnitClss) && string.IsNullOrEmpty(excel.OrderQty))
                    {
                        break;
                    }
                }

                if (listExcel.Count > 0)
                {
                    List<Procedure> Prolist = new List<Procedure>();
                    List<Dictionary<string, object>> ListParameter = new List<Dictionary<string, object>>();
                    for (int i = 0; i < listExcel.Count; i++)
                    {
                        Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                        sqlParameter.Add("CustomID", string.IsNullOrEmpty(listExcel[i].CustomID) ? "" : listExcel[i].CustomID);
                        sqlParameter.Add("Model", string.IsNullOrEmpty(listExcel[i].Model) ? "" : listExcel[i].Model);
                        sqlParameter.Add("BuyerArticleNo", string.IsNullOrEmpty(listExcel[i].BuyerArticleNo) ? "" : listExcel[i].BuyerArticleNo);
                        sqlParameter.Add("Article", string.IsNullOrEmpty(listExcel[i].Article) ? "" : listExcel[i].Article);
                        sqlParameter.Add("UnitClss", string.IsNullOrEmpty(listExcel[i].UnitClss) ? "" : listExcel[i].UnitClss);
                        sqlParameter.Add("OrderQty", string.IsNullOrEmpty(listExcel[i].OrderQty) ? "" : listExcel[i].OrderQty);

                        Procedure pro2 = new Procedure();
                        pro2.Name = "xp_Order_iOrderExcel";
                        pro2.OutputUseYN = "N";
                        pro2.OutputName = "";
                        pro2.OutputLength = "10";

                        Prolist.Add(pro2);
                        ListParameter.Add(sqlParameter);
                    }

                    string[] Confirm = new string[2];
                    Confirm = DataStore.Instance.ExecuteAllProcedureOutputNew_NewLog(Prolist, ListParameter, "C");
                    if (Confirm[0] != "success")
                        MessageBox.Show("[저장실패]\r\n" + Confirm[1].ToString());
                    else
                        MessageBox.Show("업로드가 완료되었습니다.");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message.ToString());
            }
            finally
            {
                excelapp.Visible = true;
                workbook.Close(true);
                excelapp.Quit();

                lib2.ReleaseExcelObject(workbook);
                lib2.ReleaseExcelObject(worksheet);
                lib2.ReleaseExcelObject(excelapp);
                lib2 = null;

                upload_fileName = "";
                listExcel.Clear();
            }
        }

        // 엑셀
        private void btnExcel_Click(object sender, RoutedEventArgs e)
        {
            DataTable dt = null;

            string[] lst = new string[4];
            lst[0] = "계획 리스트";
            lst[1] = "필요 원, 부자재";
            lst[2] = dgdMain.Name;
            lst[3] = dgdSub.Name;

            ExportExcelxaml ExpExc = new ExportExcelxaml(lst);
            ExpExc.WindowStartupLocation = WindowStartupLocation.CenterScreen;
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

                    if (lib.GenerateExcel(dt, dgdMain.Name))
                    {
                        lib.excel.Visible = true;
                        lib.ReleaseExcelObject(lib.excel);
                    }
                    else
                        return;
                }
                else if (ExpExc.DialogResult.HasValue)
                {
                    if (ExpExc.choice.Equals(dgdSub.Name))
                    {
                        DataStore.Instance.InsertLogByForm(this.GetType().Name, "E");

                        if (ExpExc.Check.Equals("Y"))
                            dt = lib.DataGridToDTinHidden(dgdSub);
                        else
                            dt = lib.DataGirdToDataTable(dgdSub);

                        if (lib.GenerateExcel(dt, dgdSub.Name))
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

        private void btnCalc_Click(object sender, RoutedEventArgs e)
        {
            using (Loading lw = new Loading(beCalc))
            {
                lw.ShowDialog();
            }
        }

        private void beCalc()
        {
            rowNum = 0;
            btnCalc.IsEnabled = false;

            Dispatcher.BeginInvoke(new Action(() =>
            {
                try
                {
                    List<Procedure> Prolist = new List<Procedure>();
                    List<Dictionary<string, object>> ListParameter = new List<Dictionary<string, object>>();

                    Dictionary<string, object> sqlParameter = new Dictionary<string, object>();

                    // 날짜
                    sqlParameter.Add("YYYYMM", chkDay.IsChecked == true ? dtpDate.SelectedDate.Value.ToString("yyyyMM") : "");
                    sqlParameter.Add("CreateUserID", MainWindow.CurrentUser);

                    Procedure pro = new Procedure();
                    pro.Name = "xp_WeekPlan_AutoiMtrReqPlan";
                    pro.OutputUseYN = "N";
                    pro.OutputName = "";
                    pro.OutputLength = "";

                    Prolist.Add(pro);
                    ListParameter.Add(sqlParameter);

                    string[] Confirm = new string[2];
                    Confirm = DataStore.Instance.ExecuteAllProcedureOutputNew_NewLog(Prolist, ListParameter, "U");
                    if (Confirm[0] == "success")
                        re_Search();
                    else
                        MessageBox.Show("[저장실패]\r\n" + Confirm[1].ToString());
                }
                catch (Exception ex)
                {
                    MessageBox.Show("[오류내용]: " + ex.ToString());
                }
                finally
                {
                    DataStore.Instance.CloseConnection();
                }
            }), System.Windows.Threading.DispatcherPriority.Background);

            btnCalc.IsEnabled = true;
        }
        #endregion 버튼컨트롤

        private void search_BtnDate_Control(byte flag = 0)
        {
            // 1: 전월, 2: 금월

            DateTime[] dateTime = { DateTime.Today, DateTime.Today };
            switch (flag)
            {
                case 1: dateTime = lib.BringLastMonthContinue(dtpDate.SelectedDate.Value); break;
                case 2: dateTime = lib.BringThisMonthDatetimeList().ToArray(); break;
            }

            dtpDate.SelectedDate = dateTime[0];
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

        #region 주요메서드
        #region 검색
        private void beSearch()
        {
            rowNum = 0;
            btnSearch.IsEnabled = false;

            Dispatcher.BeginInvoke(new Action(() =>
            {
                if (CheckData())
                    re_Search();
            }), System.Windows.Threading.DispatcherPriority.Background);

            btnSearch.IsEnabled = true;
        }

        private void re_Search()
        {
            listCheckGrid.Clear();

            FillGrid();

            if (dgdMain.Items.Count == 0)
                DataContext = null;
        }

        private void FillGrid()
        {
            if (dgdMain.Items.Count > 0)
                dgdMain.Items.Clear();

            if (dgdSub.Items.Count > 0)
                dgdSub.Items.Clear();

            try
            {
                Dictionary<string, object> sqlParameter = new Dictionary<string, object>();

                // 날짜
                sqlParameter.Add("SDate", chkDay.IsChecked == true ? dtpDate.SelectedDate.Value.ToString("yyyyMMdd") : "");
                // 거래처
                sqlParameter.Add("ChkCustomID", chkCustom.IsChecked == true ? 1 : 0);
                sqlParameter.Add("CustomID", chkCustom.IsChecked == true ? (txtCustom.Tag != null ? txtCustom.Tag.ToString() : "") : "");
                // 품번
                sqlParameter.Add("ChkArticleID", chkBuyerArticleNo.IsChecked == true ? 1 : 0);
                sqlParameter.Add("ArticleID", chkBuyerArticleNo.IsChecked == true ? (txtBuyerArticleNo.Tag == null ? "" : txtBuyerArticleNo.Tag.ToString()) : "");
                // 품명
                sqlParameter.Add("ChkArticle", chkArticle.IsChecked == true ? 1 : 0);
                sqlParameter.Add("Article", chkArticle.IsChecked == true ? (txtArticle.Text == string.Empty ? "" : txtArticle.Text) : "");


                DataSet ds = DataStore.Instance.ProcedureToDataSet_LogWrite("xp_ord_sProdPlanMtr", sqlParameter, true, "R");
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
                            var Win = new Win_ord_ProductNeedMonth_U_View()
                            {
                                Num = idx,
                                Step = dr["Step"].ToString(),
                                CustomID = dr["CustomID"].ToString(),
                                KCustom = dr["KCustom"].ToString(),
                                Model = dr["Model"].ToString(),
                                ArticleID = dr["ArticleID"].ToString(),
                                Article = dr["Article"].ToString(),
                                BuyerArticleNo = dr["BuyerArticleNo"].ToString(),
                                StockQty = stringFormatN0(dr["StockQty"].ToString()),
                                PlanQty = stringFormatN0(dr["PlanQty"].ToString()),
                                NeedQty = stringFormatN0(dr["NeedQty"].ToString()),
                                RemainQty = stringFormatN0(dr["RemainQty"].ToString()),
                            };

                            Win.Color = ConvertDouble(Win.RemainQty) < 0 ? "Red" : "Black";

                            if (Win.Step.Equals("1"))
                                dgdMain.Items.Add(Win);
                            else if (Win.Step.Equals("2"))
                                dgdSub.Items.Add(Win);
                        }
                    }
                }

                if (dgdMain.Items.Count == 0)
                    MessageBox.Show("조회된 데이터가 없습니다.");
                else
                {
                    dgdMain.Focus();
                    dgdMain.SelectedIndex = rowNum;
                    dgdMain.CurrentCell = dgdMain.SelectedCells[0];
                }
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
        #endregion 검색
        #endregion 주요메서드

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
            if (chkCustom.IsChecked == true && string.IsNullOrEmpty(txtCustom.Text)) word = "거래처";
            else if (chkArticle.IsChecked == true && string.IsNullOrEmpty(txtArticle.Text)) word = "품명";
            else if (chkBuyerArticleNo.IsChecked == true && string.IsNullOrEmpty(txtBuyerArticleNo.Text)) word = "품번";

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

        // 천자리 콤마
        private string stringFormatN0(string str)
        {
            return ConvertDouble(str) == 0 ? "0" : string.Format("{0:#,#.#####}", ConvertDouble(str));
        }

        // Int로 변환
        private int ConvertInt(string str)
        {
            int result = 0;

            if (!str.Trim().Equals(""))
            {
                str = str.Replace(",", "");

                if (int.TryParse(str, out int chkInt) == true)
                    result = chkInt;
            }

            return result;
        }

        // 소수로 변환
        private double ConvertDouble(string str)
        {
            double result = 0;

            if (!str.Trim().Equals(""))
            {
                str = str.Replace(",", "");

                if (double.TryParse(str, out double chkDouble) == true)
                    result = chkDouble;
            }

            return result;
        }
        #endregion 기타메서드        
    }

    #region View 클래스
    class Win_ord_ProductNeedMonth_U_View : BaseView
    {
        public int Num { get; set; }
        public string Step { get; set; }
        public string CustomID { get; set; }
        public string KCustom { get; set; }
        public string Model { get; set; }
        public string ArticleID { get; set; }
        public string Article { get; set; }
        public string BuyerArticleNo { get; set; } 
        public string StockQty { get; set; }
        public string PlanQty { get; set; }
        public string NeedQty { get; set; }
        public string RemainQty { get; set; }
        public string Color { get; set; }
    }
    #endregion view 클래스
}

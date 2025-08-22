using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
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
using System.Collections.ObjectModel;

/**************************************************************************************************
'** 프로그램명 : Win_out_OutwareReq_U
'** 설명       : 출고지시
'** 작성일자   : 
'** 작성자     : 장시영
'**------------------------------------------------------------------------------------------------
'**************************************************************************************************
' 변경일자  , 변경자, 요청자    , 요구사항ID      , 요청 및 작업내용
'**************************************************************************************************
' 2023.03.29, 장시영, 삼익SDT에서 가져옴
'**************************************************************************************************/

namespace WizMes_WellMade
{
    /// <summary>
    /// Win_Out_OutwareReq_U.xaml에 대한 상호 작용 논리
    /// </summary>
    public partial class Win_out_OutwareReq_U : UserControl
    {
        string stDate = string.Empty;
        string stTime = string.Empty;

        string strFlag = string.Empty;
        string reqFlag = string.Empty;
        int rowNum = 0;
        

        Lib lib = new Lib();

        Win_ord_OutwareReq_U_View ReqView = new Win_ord_OutwareReq_U_View();

        // 인쇄 활용 용도 (프린트)
        private Microsoft.Office.Interop.Excel.Application excelapp;
        private Microsoft.Office.Interop.Excel.Workbook workbook;
        private Microsoft.Office.Interop.Excel.Worksheet worksheet;
        private Microsoft.Office.Interop.Excel.Range workrange;
        private Microsoft.Office.Interop.Excel.Worksheet copysheet;
        private Microsoft.Office.Interop.Excel.Worksheet pastesheet;
        NoticeMessage msg = new NoticeMessage();

        // 출고증 인쇄시 사용
        List<Win_ord_OutwareReq_U_View> listOutWareReqPrint = new List<Win_ord_OutwareReq_U_View>();
        ObservableCollection<Win_ord_OutwareReq_U_View> ovcListOutWareReq  = new ObservableCollection<Win_ord_OutwareReq_U_View>();


        public Win_out_OutwareReq_U()
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

            btnCloseReq.IsEnabled = false;
           
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
        private void btnLastMonth_Click(object sender, RoutedEventArgs e) { search_BtnDate_Control(1); }
        // 금월
        private void btnThisMonth_Click(object sender, RoutedEventArgs e) { search_BtnDate_Control(2); }
        // 전일
        private void btnYesterDay_Click(object sender, RoutedEventArgs e) { search_BtnDate_Control(3); }
        // 금일
        private void btnToday_Click(object sender, RoutedEventArgs e) { search_BtnDate_Control(); }
        #endregion 날짜컨트롤

        #region 검색컨트롤
        // 거래처
        private void lblCustom_MouseLeftButtonUp(object sender, MouseButtonEventArgs e) { search_CheckBox_Control(chkCustom); }
        private void chkCustom_Checked(object sender, RoutedEventArgs e) { search_CheckBox_Checked_Control(true, txtCustom, btnPfCustom); }
        private void chkCustom_Unchecked(object sender, RoutedEventArgs e) { search_CheckBox_Checked_Control(false, txtCustom, btnPfCustom); }
        private void txtCustom_KeyDown(object sender, KeyEventArgs e) { if (e.Key == Key.Enter) search_PlusFinder_Control(txtCustom, (int)Defind_CodeFind.DCF_CUSTOM, ""); }
        private void btnPfCustom_Click(object sender, RoutedEventArgs e) { search_PlusFinder_Control(txtCustom, (int)Defind_CodeFind.DCF_CUSTOM, ""); }

        // 최종고객사
        private void lblInCustom_MouseLeftButtonUp(object sender, MouseButtonEventArgs e) { search_CheckBox_Control(chkInCustom); }
        private void chkInCustom_Checked(object sender, RoutedEventArgs e) { search_CheckBox_Checked_Control(true, txtInCustom, btnPfInCustom); }
        private void chkInCustom_Unchecked(object sender, RoutedEventArgs e) { search_CheckBox_Checked_Control(false, txtInCustom, btnPfInCustom); }
        private void txtInCustom_KeyDown(object sender, KeyEventArgs e) { if (e.Key == Key.Enter) search_PlusFinder_Control(txtInCustom, (int)Defind_CodeFind.DCF_CUSTOM, ""); }
        private void btnPfInCustom_Click(object sender, RoutedEventArgs e) { search_PlusFinder_Control(txtInCustom, (int)Defind_CodeFind.DCF_CUSTOM, ""); }

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
            lblMsg.Content = "자료 입력 중";
            DataContext = null;

            BtnControl(false);

            txtOutwareReqID.Text = "";
            reqTxtCustom.Text = "";
            reqTxtCustom.Tag = null;

            reqDtpEDate.SelectedDate = DateTime.Today;
            reqDtpSDate.SelectedDate = DateTime.Today;

            rowNum = dgdMain.SelectedIndex > -1 ? dgdMain.SelectedIndex : 0;
            dgdSub.Items.Clear();
        }

        // 수정
        private void btnUpdate_Click(object sender, RoutedEventArgs e)
        {
            ReqView = dgdMain.SelectedItem as Win_ord_OutwareReq_U_View;
            if (ReqView != null)
            {
                strFlag = "U";
                lblMsg.Content = "자료 수정 중";

                BtnControl(false);

                reqDtpEDate.SelectedDate = DateTime.Today;
                reqDtpSDate.SelectedDate = DateTime.Today;

                rowNum = dgdMain.SelectedIndex;
            }
        }

        // 삭제
        private void btnDelete_Click(object sender, RoutedEventArgs e)
        {
            if (listOutWareReqPrint.Count == 0)
                MessageBox.Show("삭제할 데이터가 지정되지 않았습니다. 데이터를 지정하고 눌러주세요.");
            else
            {
                if (MessageBox.Show("선택하신 항목을 삭제하시겠습니까?", "삭제 전 확인", MessageBoxButton.YesNo) == MessageBoxResult.Yes)
                {
                    using (Loading lw = new Loading(beDelete))
                    {
                        lw.ShowDialog();
                    }
                }
            }
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
            using (Loading lw = new Loading(beSave))
            {
                lw.ShowDialog();
            }
        }

        // 취소
        private void btnCancel_Click(object sender, RoutedEventArgs e)
        {
            BtnControl(true);

            txtOutwareReqID.Text = "";
            reqTxtCustom.Text = "";
            reqTxtCustom.Tag = null;

            if (strFlag.Equals("U"))
                re_Search();
            else
            {
                rowNum = 0;
                re_Search();
            }

            strFlag = string.Empty;
        }

        // 엑셀
        private void btnExcel_Click(object sender, RoutedEventArgs e)
        {
            DataTable dt = null;

            string[] lst = new string[4];
            lst[0] = "출고지시";
            lst[1] = "출고지시 상세내역";
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

        // 인쇄 (출고증)
        private void btnPrint_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                ContextMenu menu = btnPrint.ContextMenu;
                menu.StaysOpen = true;
                menu.IsOpen = true;
            }
            catch (Exception ee)
            {
                MessageBox.Show("오류지점 - btnPrint_Click : " + ee.ToString());
            }
        }

        // 인쇄 - 미리보기
        private void menuSeeAhead_Click(object sender, RoutedEventArgs e) { menuPrint_Click(true); }
        // 인쇄 - 바로인쇄
        private void menuRighPrint_Click(object sender, RoutedEventArgs e) { menuPrint_Click(false); }

        private void menuPrint_Click(bool seeAhead)
        {
            try
            {
                if (dgdMain.Items.Count == 0)
                {
                    MessageBox.Show("먼저 검색해 주세요.");
                    return;
                }

                var OBJ = dgdMain.SelectedItem as Win_ord_OutwareReq_U_View;
                if (OBJ == null)
                {
                    MessageBox.Show("출고중 항목이 정확히 선택되지 않았습니다.");
                    return;
                }

                DataStore.Instance.InsertLogByForm(this.GetType().Name, "P");

                msg.Show();
                msg.Topmost = true;
                msg.Refresh();

                PrintWork(seeAhead);
                msg.Visibility = Visibility.Hidden;
            }
            catch (Exception ee)
            {
                MessageBox.Show("오류지점 - menuRighPrint_Click : " + ee.ToString());
            }
        }

        private void PrintWork(bool previewYN)
        {
            Lib lib2 = new Lib();
            try
            {
                if (listOutWareReqPrint.Count == 0)
                {
                    MessageBox.Show("인쇄할 출고증을 선택하세요.");
                    lib2 = null;
                    return;
                }

                excelapp = new Microsoft.Office.Interop.Excel.Application();

                string MyBookPath = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetEntryAssembly().Location) + "\\출고증.xlsx";
                workbook = excelapp.Workbooks.Add(MyBookPath);
                worksheet = workbook.Sheets["Form"];
                pastesheet = workbook.Sheets["Print"];

                int copyLine = 1;
                int copyRow = 46;

                //int numberOfPrintPage = 2;

                int inputPossibleRowCnt = 9;    // 내역 입력 가능한 갯수
                int firstStartRowNum = 9;       // 첫번째 내역 입력 시작점
                //int secondStartRowNum = 32;     // 두번째 내역 입력 시작점

                for (int i = 0; i < listOutWareReqPrint.Count; i++)
                {
                    Win_ord_OutwareReq_U_View outwareReq = listOutWareReqPrint[i];
                    if (outwareReq == null)
                        continue;

                    //int startRowNum = i % numberOfPrintPage == 0 ? firstStartRowNum : secondStartRowNum;
                    int startRowNum = firstStartRowNum;
                    int startInfoRowNum = startRowNum - 2;

                    // 부서명
                    workrange = worksheet.get_Range("A" + startInfoRowNum.ToString(), "D" + startInfoRowNum.ToString());
                    workrange.Value2 = listOutWareReqPrint[i].KCustom;

                    // 출고일자
                    workrange = worksheet.get_Range("E" + startInfoRowNum.ToString(), "I" + startInfoRowNum.ToString());
                    workrange.Value2 = listOutWareReqPrint[i].ReqDate.Replace("-", ".");

                    Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                    sqlParameter.Clear();
                    sqlParameter.Add("OutwareReqID", outwareReq.OutwareReqID);

                    DataSet ds = DataStore.Instance.ProcedureToDataSet("xp_Outware_sOutwareReqSub", sqlParameter, false);
                    if (ds != null && ds.Tables.Count > 0)
                    {
                        DataTable dt = ds.Tables[0];
                        if (dt.Rows.Count > 0)
                        {
                            int Num = 1;

                            DataRowCollection drc = dt.Rows;
                            foreach (DataRow dr in drc)
                            {                                
                                int rowNum = startRowNum + (Num - 1);

                                // No
                                workrange = worksheet.get_Range("A" + rowNum.ToString());
                                workrange.Value2 = (Num + 1).ToString();

                                // 발주번호
                                workrange = worksheet.get_Range("B" + rowNum.ToString(), "C" + rowNum.ToString());
                                workrange.Value2 = dr["OrderNo"].ToString();

                                // 제품명
                                workrange = worksheet.get_Range("D" + rowNum.ToString(), "E" + rowNum.ToString());
                                workrange.Value2 = dr["Article"].ToString();

                                // 업체
                                workrange = worksheet.get_Range("F" + rowNum.ToString(), "G" + rowNum.ToString());
                                workrange.Value2 = dr["KInCustom"].ToString();

                                // 수량
                                workrange = worksheet.get_Range("H" + rowNum.ToString());
                                workrange.Value2 = dr["ReqQty"].ToString();

                                // 비고
                                workrange = worksheet.get_Range("I" + rowNum.ToString());
                                workrange.Value2 = dr["Remark"].ToString();

                                Num++;

                                if (Num > inputPossibleRowCnt)
                                    break;
                            }
                        }
                    }

                    // 출고증 다음 및 종료 조건
                    //if (i == listOutWareReqPrint.Count - 1 || i % numberOfPrintPage == 1)
                    {
                        // 붙여넣기
                        worksheet.Select();
                        worksheet.Range["A1", "I" + copyRow.ToString()].EntireRow.Copy();
                        pastesheet.Select();
                        workrange = pastesheet.Rows[copyLine];
                        workrange.Select();
                        pastesheet.Paste();

                        // 내역 삭제
                        workrange = worksheet.get_Range("A" + firstStartRowNum.ToString(), "I" + (firstStartRowNum + inputPossibleRowCnt - 1).ToString());
                        workrange.ClearContents();

                        //workrange = worksheet.get_Range("A" + secondStartRowNum.ToString(), "I" + (secondStartRowNum + inputPossibleRowCnt - 1).ToString());
                        //workrange.ClearContents();

                        // 부서명 삭제
                        workrange = worksheet.get_Range("A" + (firstStartRowNum - 2).ToString(), "D" + (firstStartRowNum - 2).ToString());
                        workrange.ClearContents();

                        //workrange = worksheet.get_Range("A" + (secondStartRowNum - 2).ToString(), "D" + (secondStartRowNum - 2).ToString());
                        //workrange.ClearContents();

                        // 출고일자 삭제
                        workrange = worksheet.get_Range("E" + (firstStartRowNum - 2).ToString(), "I" + (firstStartRowNum - 2).ToString());
                        workrange.ClearContents();

                        //workrange = worksheet.get_Range("E" + (secondStartRowNum - 2).ToString(), "I" + (secondStartRowNum - 2).ToString());
                        //workrange.ClearContents();

                        copyLine += copyRow;
                    }
                }

                excelapp.Visible = true;
                msg.Hide();

                if (previewYN == true)
                    pastesheet.PrintPreview();
                else
                    pastesheet.PrintOutEx();
            }
            catch (Exception ee)
            {
                MessageBox.Show("오류지점 = PrintWork : " + ee.ToString());
            }

            lib2.ReleaseExcelObject(workbook);
            lib2.ReleaseExcelObject(worksheet);
            lib2.ReleaseExcelObject(pastesheet);
            lib2.ReleaseExcelObject(excelapp);
            lib2 = null;
        }

        // 인쇄 - 닫기
        private void menuClose_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                ContextMenu menu = btnPrint.ContextMenu;
                menu.StaysOpen = false;
                menu.IsOpen = false;
            }
            catch (Exception ee)
            {
                MessageBox.Show("오류지점 - menuClose_Click : " + ee.ToString());
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

        #region 하단컨트롤
        #region 날짜컨트롤
        private void reqLblOrderDay_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            search_CheckBox_Control(reqChkOrderDay);
        }

        private void reqChkOrderDay_Checked(object sender, RoutedEventArgs e)
        {
            if (reqDtpSDate != null && reqDtpEDate != null)
            {
                reqDtpSDate.IsEnabled = true;
                reqDtpEDate.IsEnabled = true;
            }
        }

        private void reqChkOrderDay_Unchecked(object sender, RoutedEventArgs e)
        {
            reqDtpSDate.IsEnabled = false;
            reqDtpEDate.IsEnabled = false;
        }
        #endregion 날짜컨트롤

        #region 검색컨트롤
        // 거래처
        /*private void reqLblCustom_MouseLeftButtonUp(object sender, MouseButtonEventArgs e) { search_CheckBox_Control(reqChkCustom); }
        private void reqChkCustom_Checked(object sender, RoutedEventArgs e) { search_CheckBox_Checked_Control(true, reqTxtCustom, reqBtnPfCustom); }
        private void reqChkCustom_Unchecked(object sender, RoutedEventArgs e) { search_CheckBox_Checked_Control(false, reqTxtCustom, reqBtnPfCustom); }*/
        private void reqTxtCustom_KeyDown(object sender, KeyEventArgs e) { if (e.Key == Key.Enter) search_PlusFinder_Control(reqTxtCustom, (int)Defind_CodeFind.DCF_CUSTOM, reqTxtCustom.Text); }
        //private void reqBtnPfCustom_Click(object sender, RoutedEventArgs e) { search_PlusFinder_Control(reqTxtCustom, 1000, reqTxtCustom.Text); }
        private void reqBtnPfCustom_Click(object sender, RoutedEventArgs e) { search_PlusFinder_Control(reqTxtCustom, (int)Defind_CodeFind.DCF_CUSTOM, reqTxtCustom.Text); }

        // 품번
        private void reqLblBuyerArticleNo_MouseLeftButtonUp(object sender, MouseButtonEventArgs e) { search_CheckBox_Control(reqChkBuyerArticleNo); }
        private void reqChkBuyerArticleNo_Checked(object sender, RoutedEventArgs e) { search_CheckBox_Checked_Control(true, reqTxtBuyerArticleNo, reqBtnPfBuyerArticleNo); }
        private void reqChkBuyerArticleNo_Unchecked(object sender, RoutedEventArgs e) { search_CheckBox_Checked_Control(false, reqTxtBuyerArticleNo, reqBtnPfBuyerArticleNo); }
        private void reqTxtBuyerArticleNo_KeyDown(object sender, KeyEventArgs e) { if (e.Key == Key.Enter) search_PlusFinder_Control(reqTxtBuyerArticleNo, 1002, reqTxtBuyerArticleNo.Text); }
        private void reqBtnPfBuyerArticleNo_Click(object sender, RoutedEventArgs e) { search_PlusFinder_Control(reqTxtBuyerArticleNo, 1002, reqTxtBuyerArticleNo.Text); }

        // 품명
        private void reqLblArticle_MouseLeftButtonUp(object sender, MouseButtonEventArgs e) { search_CheckBox_Control(reqChkArticle); }
        private void reqChkArticle_Checked(object sender, RoutedEventArgs e) { search_CheckBox_Checked_Control(true, reqTxtArticle, reqBtnPfArticle); }
        private void reqChkArticle_Unchecked(object sender, RoutedEventArgs e) { search_CheckBox_Checked_Control(false, reqTxtArticle, reqBtnPfArticle); }
        private void reqTxtArticle_KeyDown(object sender, KeyEventArgs e) { if (e.Key == Key.Enter) search_PlusFinder_Control(reqTxtArticle, 1003, reqTxtArticle.Text.Trim()); }
        private void reqBtnPfArticle_Click(object sender, RoutedEventArgs e) { search_PlusFinder_Control(reqTxtArticle, 1003, reqTxtArticle.Text.Trim()); }
        #endregion 검색컨트롤
        #endregion 하단컨트롤

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
            listOutWareReqPrint.Clear();

            FillGrid();

            if (dgdMain.Items.Count == 0)
                DataContext = null;
        }

        private void FillGrid()
        {
            dgdMain.Items.Clear();
            dgdSub.Items.Clear();
            dgdTotal.Items.Clear();
            ovcListOutWareReq.Clear();
            btnCloseReq.IsEnabled = false;

            try
            {
                Dictionary<string, object> sqlParameter = new Dictionary<string, object>();

                // 날짜
                sqlParameter.Add("ChkDate", chkOrderDay.IsChecked == true ? 1 : 0);
                sqlParameter.Add("SDate", chkOrderDay.IsChecked == true ? dtpSDate.SelectedDate.Value.ToString("yyyyMMdd") : "");
                sqlParameter.Add("EDate", chkOrderDay.IsChecked == true ? dtpEDate.SelectedDate.Value.ToString("yyyyMMdd") : "");

                // 거래처
                sqlParameter.Add("ChkCustomID", chkCustom.IsChecked == true ? 1 : 0);
                sqlParameter.Add("CustomID", chkCustom.IsChecked == true ? (txtCustom.Tag != null ? txtCustom.Tag.ToString() : "") : "");
                // 최종고객사
                sqlParameter.Add("ChkInCustomID", chkInCustom.IsChecked == true ? 1 : 0);
                sqlParameter.Add("InCustomID", chkInCustom.IsChecked == true ? (txtInCustom.Tag != null ? txtInCustom.Tag.ToString() : "") : "");

                // 품번
                sqlParameter.Add("ChkArticleID", chkBuyerArticleNo.IsChecked == true ? 1 : 0);
                sqlParameter.Add("ArticleID", chkBuyerArticleNo.IsChecked == true ? (txtBuyerArticleNo.Tag == null ? "" : txtBuyerArticleNo.Tag.ToString()) : "");
                // 품명
                sqlParameter.Add("ChkArticle", chkArticle.IsChecked == true ? 1 : 0);
                sqlParameter.Add("Article", chkArticle.IsChecked == true ? (txtArticle.Text == string.Empty ? "" : txtArticle.Text) : "");

                //지시마감여부
                sqlParameter.Add("ChkCloseReq", chkCloseReq.IsChecked == true ? 1 : 0);
                sqlParameter.Add("CloseReq", chkCloseReq.IsChecked == true ? "1" : "");


                DataSet ds = DataStore.Instance.ProcedureToDataSet_LogWrite("xp_Outware_sOutwareRequest", sqlParameter, true, "R");
                if (ds != null && ds.Tables.Count > 0)
                {
                    DataTable dt = ds.Tables[0];
                    if (dt.Rows.Count > 0)
                    {
                        int idx = 0;
                        int totReqQty = 0;
                        DataRowCollection drc = dt.Rows;
                        foreach (DataRow dr in drc)
                        {
                            idx++;
                            var Win = new Win_ord_OutwareReq_U_View()
                            {
                                Num = idx,
                                OutwareReqID = dr["OutwareReqID"].ToString(),
                                Article = dr["Article"].ToString(),
                                BuyerArticleNo = dr["BuyerArticleNo"].ToString(),
                                CustomID = dr["CustomID"].ToString(),
                                KCustom = dr["KCustom"].ToString(),
                                OutClss = dr["OutClss"].ToString(),
                                ReqDate = DatePickerFormat(dr["ReqDate"].ToString()),
                                ReqQty = lib.returnNumString(dr["ReqQty"].ToString()),
                                ReqCloseClss = dr["ReqCloseClss"].ToString(),
                                OutWareReqTypeID = dr["OutWareReqTypeID"].ToString(),
                                RegDate = dr["RegDate"].ToString(),
                                RegUserID = dr["RegUserID"].ToString(),
                                Remark = dr["Remark"].ToString(),
                            };

                            dgdMain.Items.Add(Win);

                            totReqQty += ConvertInt(Win.ReqQty);
                        }

                        var total = new Win_ord_OutwareReqTot_U_View();
                        total.ReqCount = lib.returnNumString(idx.ToString()) + " 건";
                        total.ReqAmount = lib.returnNumString(totReqQty.ToString()) + " 개";

                        dgdTotal.Items.Add(total);
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

        private void dgdMain_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            try
            {
                var ReqInfo = dgdMain.SelectedItem as Win_ord_OutwareReq_U_View;

                if (ReqInfo != null)
                {
                    this.DataContext = ReqInfo;

                    string ReqID = ReqInfo.OutwareReqID;

                    txtOutwareReqID.Text = ReqID;
                    reqTxtCustom.Tag = ReqInfo.CustomID;
                    reqTxtCustom.Text = ReqInfo.KCustom;

                    FillGridSub(ReqID);
                }
            }
            catch (Exception ee)
            {
                MessageBox.Show("오류지점 - DataGridMain_SelectionChanged : " + ee.ToString());
            }
        }

        private void FillGridSub(string strReqID)
        {
            if (dgdSub.Items.Count > 0)
                dgdSub.Items.Clear();

            try
            {
                Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                sqlParameter.Clear();
                sqlParameter.Add("OutwareReqID", strReqID);
                DataSet ds = DataStore.Instance.ProcedureToDataSet("xp_Outware_sOutwareReqSub", sqlParameter, false);

                if (ds != null && ds.Tables.Count > 0)
                {
                    DataTable dt = ds.Tables[0];
                    if (dt.Rows.Count > 0)
                    {
                        int i = 0;
                        DataRowCollection drc = dt.Rows;

                        foreach (DataRow dr in drc)
                        {
                            i++;
                            var sub = new Win_ord_OutwareReqSub_U_View
                            {
                                Num = i,

                                OutwareReqID = dr["OutwareReqID"].ToString(),
                                OutwareReqSeq = dr["OutwareReqSeq"].ToString(),
                                AcptDate = dr["AcptDate"].ToString(),
                                DvlyDate = dr["DvlyDate"].ToString(),
                                OrderQty = stringFormatN0(dr["OrderQty"]),
                                OutQty = stringFormatN0(dr["OutQty"]),                            
                                OrderID = dr["OrderID"].ToString(),
                                OrderNo = dr["OrderNo"].ToString(),
                                RestOrderQty = stringFormatN0(dr["RestOrderQty"]),
                                InCustomID = dr["InCustomID"].ToString(),
                                KInCustom = dr["KInCustom"].ToString(),
                                ArticleID = dr["ArticleID"].ToString(),
                                Article = dr["Article"].ToString(),
                                BuyerArticleNo = dr["BuyerArticleNo"].ToString(),
                                StockQty = stringFormatN0(dr["StockQty"]),
                                ReqQty = stringFormatN0(dr["ReqQty"]),
                                RemainQty = stringFormatN0(dr["RemainQty"]),
                                Remark = dr["Remark"].ToString()
                            };

                            dgdSub.Items.Add(sub);
                        }
                    }
                }
            }
            catch (Exception ee)
            {
                MessageBox.Show("오류지점 - FillGridSub : " + ee.ToString());
            }
        }
        #endregion 검색

        #region 저장
        private void beSave()
        {
            btnSave.IsEnabled = false;

            Dispatcher.BeginInvoke(new Action(() =>
            {
                if (SaveData())
                {
                    if (strFlag.Equals("I"))
                        rowNum = 0;

                    BtnControl(true);
                    re_Search();
                }
            }), System.Windows.Threading.DispatcherPriority.Background);

            btnSave.IsEnabled = true;
        }

        private bool CloseReq(string OutwareReqID, string ReqFlag)
        {
            bool flag = false;

            try
            {
                Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                sqlParameter.Clear();
                sqlParameter.Add("ReqFlag", ReqFlag);
                sqlParameter.Add("OutwareReqID", OutwareReqID);
                sqlParameter.Add("CloseReqUserID", MainWindow.CurrentUser);

                string[] result = DataStore.Instance.ExecuteProcedure_NewLog("xp_Outware_uOutwareRequest_Close", sqlParameter, "D");

                if (result[0].Equals("success"))
                    flag = true;
            }
            catch (Exception ex)
            {
                MessageBox.Show("오류지점 - DeleteData : " + ex.ToString());
            }
            finally
            {
                DataStore.Instance.CloseConnection();
            }

            return flag;
        }


        private bool SaveData()
        {
            if (string.IsNullOrEmpty(reqTxtCustom.Text) || reqTxtCustom.Tag == null)
            {
                MessageBox.Show("거래처가 입력되지 않았습니다. 먼저 거래처를 입력해주세요");
                return false;
            }
            else if (dgdSub.Items.Count == 0)
            {
                MessageBox.Show("등록된 출고지시가 없습니다. 먼저 출고지시를 등록해주세요");
                return false;
            }

            bool flag = false;
            List<Procedure> Prolist = new List<Procedure>();
            List<Dictionary<string, object>> ListParameter = new List<Dictionary<string, object>>();

            try
            {
                if (dgdSub.Items.Count > 0)
                {
                    Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                    sqlParameter.Clear();

                    double totReqQty = 0;
                    for (int i = 0; i < dgdSub.Items.Count; i++)
                    {
                        var subItem = dgdSub.Items[i] as Win_ord_OutwareReqSub_U_View;
                        if (subItem != null)
                        {
                            double reqQty = lib.returnDouble(subItem.ReqQty);
                            if (reqQty == 0)
                            {
                                MessageBox.Show("출고지시량이 0인 지시건이 있습니다.");
                                return false;
                            }

                            totReqQty += reqQty;
                        }
                    }

                    sqlParameter.Add("ReqID", txtOutwareReqID.Text);
                    sqlParameter.Add("CustomID", reqTxtCustom.Tag != null ? reqTxtCustom.Tag.ToString() : "");
                    sqlParameter.Add("OutClss", "");
                    sqlParameter.Add("ReqDate", DateTime.Today.ToString("yyyyMMdd"));
                    sqlParameter.Add("ReqQty", totReqQty.ToString());
                    sqlParameter.Add("OutWareReqTypeID", "");
                    sqlParameter.Add("RegDate", "");
                    sqlParameter.Add("RegUserID", MainWindow.CurrentUser);

                    string sGetID = strFlag.Equals("I") ? string.Empty : txtOutwareReqID.Text;

                    #region 추가
                    if (strFlag.Equals("I"))
                    {
                        sqlParameter.Add("CreateUserID", MainWindow.CurrentUser);

                        Procedure pro1 = new Procedure();
                        pro1.Name = "xp_Outware_iOutwareRequest";
                        pro1.OutputUseYN = "Y";
                        pro1.OutputName = "ReqID";
                        pro1.OutputLength = "10";

                        Prolist.Add(pro1);
                        ListParameter.Add(sqlParameter);

                        List<KeyValue> list_Result = new List<KeyValue>();
                        list_Result = DataStore.Instance.ExecuteAllProcedureOutputGetCS_NewLog(Prolist, ListParameter, "C");

                        if (list_Result[0].key.ToLower() == "success")
                        {
                            list_Result.RemoveAt(0);
                            for (int i = 0; i < list_Result.Count; i++)
                            {
                                KeyValue kv = list_Result[i];
                                if (kv.key == "ReqID")
                                {
                                    sGetID = kv.value;
                                    flag = true;
                                }
                            }
                        }
                        else
                        {
                            MessageBox.Show("[저장실패]\r\n" + list_Result[0].value.ToString());
                            return false;
                        }

                        Prolist.Clear();
                        ListParameter.Clear();
                    }
                    #endregion 추가

                    #region 수정
                    else if (strFlag.Equals("U"))
                    {
                        sqlParameter.Add("UpdateUserID", MainWindow.CurrentUser);

                        Procedure pro2 = new Procedure();
                        pro2.Name = "xp_Outware_uOutwareRequest";
                        pro2.OutputUseYN = "N";
                        pro2.OutputName = "ReqID";
                        pro2.OutputLength = "10";

                        Prolist.Add(pro2);
                        ListParameter.Add(sqlParameter);

                        // 모든것을 삭제한 후에, 새롭게 추가
                        sqlParameter = new Dictionary<string, object>();
                        sqlParameter.Clear();
                        sqlParameter.Add("OutwareReqID", txtOutwareReqID.Text);

                        Procedure pro3 = new Procedure();
                        pro3.Name = "xp_Outware_dOutwareReqSubAll";
                        pro3.OutputUseYN = "N";
                        pro3.OutputName = "OutwareReqID";
                        pro3.OutputLength = "10";

                        Prolist.Add(pro3);
                        ListParameter.Add(sqlParameter);
                    }
                    #endregion 수정

                    // 서브그리드 추가
                    for (int i = 0; i < dgdSub.Items.Count; i++)
                    {
                        var reqSub = dgdSub.Items[i] as Win_ord_OutwareReqSub_U_View;

                        sqlParameter = new Dictionary<string, object>();
                        sqlParameter.Clear();

                        sqlParameter.Add("OutWareReqID", sGetID);
                        sqlParameter.Add("OutWareReqSeq", i + 1);
                        sqlParameter.Add("OrderID", reqSub.OrderID);
                        sqlParameter.Add("OrderNo", reqSub.OrderNo);
                        sqlParameter.Add("ArticleID", reqSub.ArticleID);
                        sqlParameter.Add("OrderQty", ConvertDouble(reqSub.OrderQty));
                        sqlParameter.Add("RestOrderQty", ConvertDouble(reqSub.RestOrderQty));
                        sqlParameter.Add("ReqEnd", reqSub.ReqEnd != null ? (reqSub.ReqEnd != "1" ? "1" : "") : ""); //reqEnd가 null이 아니면 1인지? 1이면 1 빈칸이면 빈칸 null이면 초기니까 빈칸 
                        sqlParameter.Add("InCustomID", string.IsNullOrEmpty(reqTxtCustom.Text) ? "" : reqTxtCustom.Tag.ToString());
                        sqlParameter.Add("StockQty", ConvertDouble(reqSub.StockQty));
                        sqlParameter.Add("ReqQty", ConvertDouble(reqSub.ReqQty));
                        sqlParameter.Add("RemainQty", ConvertDouble(reqSub.RemainQty));
                        sqlParameter.Add("Remark", reqSub.Remark);

                        sqlParameter.Add("CreateUserID", MainWindow.CurrentUser);

                        Procedure pro2 = new Procedure();
                        pro2.Name = "xp_Outware_iOutwareReqSub";
                        pro2.OutputUseYN = "N";
                        pro2.OutputName = "OutwareReqID";
                        pro2.OutputLength = "10";

                        Prolist.Add(pro2);
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
                        flag = true;
                }
                else
                    MessageBox.Show("출고지시된 수주가 없습니다");
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
            finally
            {
                DataStore.Instance.CloseConnection();
            }

            return flag;
        }
        #endregion 저장

        #region 삭제
        private void beDelete()
        {
            btnDelete.IsEnabled = false;

            Dispatcher.BeginInvoke(new Action(() =>
            {
                foreach (Win_ord_OutwareReq_U_View remove in listOutWareReqPrint)
                    if (remove.Remark.Contains("완료"))
                    {
                        MessageBoxResult result = MessageBox.Show("선택건 중에 출하 완료 된 건이 있습니다. 완료건은 마감처리 하시겠습니까?\r\n 아니오를 누르면 모두 삭제합니다.", "마감 전 확인", MessageBoxButton.YesNoCancel);


                        if (result == MessageBoxResult.Yes)
                        {
                            if (remove.Remark.Contains("완료"))
                            {
                                reqFlag = "1";
                                CloseReq(remove.OutwareReqID, reqFlag);
                            }
                            else
                            {
                                DeleteData(remove.OutwareReqID);
                            }
                        }
                        else if (result == MessageBoxResult.No)
                        {
                                DeleteData(remove.OutwareReqID);
                        }
                        else
                        {
                            FillGrid();
                            return;
                        } 
                    }
                    else
                    {
                        DeleteData(remove.OutwareReqID);

                    }




                rowNum = 0;
                reqFlag = string.Empty;
                re_Search();
            }), System.Windows.Threading.DispatcherPriority.Background);

            btnDelete.IsEnabled = true;
        }

        private bool DeleteData(string ReqID)
        {
            bool flag = false;

            try
            {
                Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                sqlParameter.Clear();
                sqlParameter.Add("OutwareReqID", ReqID);

                string[] result = DataStore.Instance.ExecuteProcedure_NewLog("xp_Outware_dOutwareRequest", sqlParameter, "D");

                if (result[0].Equals("success"))
                    flag = true;
            }
            catch (Exception ex)
            {
                MessageBox.Show("오류지점 - DeleteData : " + ex.ToString());
            }
            finally
            {
                DataStore.Instance.CloseConnection();
            }

            return flag;
        }
        #endregion 삭제
        #endregion 주요메서드

        #region 서브그리드 메서드
        //서브 데이터 그리드 최종고객사 텍스트박스 키다운 이벤트
        private void DataGridSubTextBoxInCustom_KeyDown(object sender, KeyEventArgs e)
        {
            try
            {
                if (e.Key == Key.Enter)
                {
                    if (strFlag == "I" || strFlag == "U")
                    {
                        var SubItem = dgdSub.CurrentItem as Win_ord_OutwareReqSub_U_View;
                        TextBox sub_tb = sender as TextBox;

                        search_PlusFinder_Control(sub_tb, (int)Defind_CodeFind.DCF_CUSTOM, sub_tb.Text.Trim());

                        if (sub_tb.Tag != null && SubItem != null)
                        {
                            SubItem.InCustomID = sub_tb.Tag.ToString();
                            SubItem.KInCustom = sub_tb.Text;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                //MessageBox.Show("오류 발생, 오류 내용 : " + ex.ToString());
            }
            finally
            {
                DataStore.Instance.CloseConnection();
            }
        }

        #region 방향키 이동 및 셀 포커스
        private void DataGridSub_PreviewKeyDown(object sender, KeyEventArgs e)
        {
            try
            {
                if (e.Key == Key.Down || e.Key == Key.Up || e.Key == Key.Left || e.Key == Key.Right)
                {
                    DataGridSub_KeyDown(sender, e);
                }
            }
            catch (Exception ee)
            {
                MessageBox.Show("오류지점 - DataGridSub_PreviewKeyDown " + ee.ToString());
            }
        }

        private void DataGridSub_KeyDown(object sender, KeyEventArgs e)
        {
            //try
            //{
            //    var SubItem = dgdSub.CurrentItem as Win_ord_OutWare_Scan_Sub_CodeView;
            //    int rowCount = dgdSub.Items.IndexOf(dgdSub.CurrentItem);
            //    int colCount = dgdSub.Columns.IndexOf(dgdSub.CurrentCell.Column);
            //    int StartColumnCount = 1; //DataGridSub.Columns.IndexOf(dgdtpeMCoperationRateScore);
            //    int EndColumnCount = 7; //DataGridSub.Columns.IndexOf(dgdtpeComments);

            //    if (e.Key == Key.Enter)
            //    {
            //        e.Handled = true;
            //        (sender as DataGridCell).IsEditing = false;

            //        if (EndColumnCount == colCount && dgdSub.Items.Count - 1 > rowCount)
            //        {
            //            dgdSub.SelectedIndex = rowCount + 1;
            //            dgdSub.CurrentCell = new DataGridCellInfo(dgdSub.Items[rowCount + 1], dgdSub.Columns[StartColumnCount]);
            //        }
            //        else if (EndColumnCount > colCount && dgdSub.Items.Count - 1 > rowCount)
            //        {
            //            dgdSub.CurrentCell = new DataGridCellInfo(dgdSub.Items[rowCount], dgdSub.Columns[colCount + 1]);
            //        }
            //        else if (EndColumnCount == colCount && dgdSub.Items.Count - 1 == rowCount)
            //        {
            //            btnSave.Focus();
            //        }
            //        else if (EndColumnCount > colCount && dgdSub.Items.Count - 1 == rowCount)
            //        {
            //            dgdSub.CurrentCell = new DataGridCellInfo(dgdSub.Items[rowCount], dgdSub.Columns[colCount + 1]);
            //        }
            //        else
            //        {
            //            MessageBox.Show("있으면 찾아보자...");
            //        }
            //    }
            //    else if (e.Key == Key.Down)
            //    {
            //        e.Handled = true;
            //        (sender as DataGridCell).IsEditing = false;

            //        if (dgdSub.Items.Count - 1 > rowCount)
            //        {
            //            dgdSub.SelectedIndex = rowCount + 1;
            //            dgdSub.CurrentCell = new DataGridCellInfo(dgdSub.Items[rowCount + 1], dgdSub.Columns[colCount]);
            //        }
            //        else if (dgdSub.Items.Count - 1 == rowCount)
            //        {
            //            if (EndColumnCount > colCount)
            //            {
            //                dgdSub.SelectedIndex = 0;
            //                dgdSub.CurrentCell = new DataGridCellInfo(dgdSub.Items[0], dgdSub.Columns[colCount + 1]);
            //            }
            //            else
            //            {
            //                btnSave.Focus();
            //            }
            //        }
            //    }
            //    else if (e.Key == Key.Up)
            //    {
            //        e.Handled = true;
            //        (sender as DataGridCell).IsEditing = false;

            //        if (rowCount > 0)
            //        {
            //            dgdSub.SelectedIndex = rowCount - 1;
            //            dgdSub.CurrentCell = new DataGridCellInfo(dgdSub.Items[rowCount - 1], dgdSub.Columns[colCount]);
            //        }
            //    }
            //    else if (e.Key == Key.Left)
            //    {
            //        e.Handled = true;
            //        (sender as DataGridCell).IsEditing = false;

            //        if (colCount > 0)
            //        {
            //            dgdSub.CurrentCell = new DataGridCellInfo(dgdSub.Items[rowCount], dgdSub.Columns[colCount - 1]);
            //        }
            //    }
            //    else if (e.Key == Key.Right)
            //    {
            //        e.Handled = true;
            //        (sender as DataGridCell).IsEditing = false;

            //        if (EndColumnCount > colCount)
            //        {
            //            dgdSub.CurrentCell = new DataGridCellInfo(dgdSub.Items[rowCount], dgdSub.Columns[colCount + 1]);
            //        }
            //        else if (EndColumnCount == colCount)
            //        {
            //            if (dgdSub.Items.Count - 1 > rowCount)
            //            {
            //                dgdSub.SelectedIndex = rowCount + 1;
            //                dgdSub.CurrentCell = new DataGridCellInfo(dgdSub.Items[rowCount + 1], dgdSub.Columns[StartColumnCount]);
            //            }
            //            else
            //            {
            //                btnSave.Focus();
            //            }
            //        }
            //    }
            //}
            //catch (Exception ee)
            //{
            //    MessageBox.Show("오류지점 - DataGridSub_KeyDown " + ee.ToString());
            //}
        }

        private void DataGridSub_TextFocus(object sender, KeyEventArgs e)
        {
            try
            {
                Lib.Instance.DataGridINControlFocus(sender, e);
            }
            catch (Exception ee)
            {
                MessageBox.Show("오류지점 - DataGridSub_TextFocus " + ee.ToString());
            }
        }

        private void DataGridSub_GotFocus(object sender, RoutedEventArgs e)
        {
            try
            {
                if (strFlag.Equals("I") || strFlag.Equals("U"))
                {
                    DataGridCell cell = sender as DataGridCell;
                    cell.IsEditing = true;
                }
            }
            catch (Exception ee)
            {
                MessageBox.Show("오류지점 - DataGridSub_GotFocus " + ee.ToString());
            }
        }

        private void DataGridSub_MouseUp(object sender, MouseButtonEventArgs e)
        {
            try
            {
                Lib.Instance.DataGridINBothByMouseUP(sender, e);
            }
            catch (Exception ee)
            {
                MessageBox.Show("오류지점 - DataGridSub_MouseUp " + ee.ToString());
            }
        }
        #endregion 방향키 이동 및 셀 포커스

        private void chkReq_Click(object sender, RoutedEventArgs e)
        {
            CheckBox chkSender = sender as CheckBox;
            var view = chkSender.DataContext as Win_ord_OutwareReq_U_View;
            if (view != null)
            {
                if (chkSender.IsChecked == true)
                {
                    view.Chk = true;

                    if (listOutWareReqPrint.Contains(view) == false)
                        listOutWareReqPrint.Add(view);
                        ovcListOutWareReq.Add(view);
                }
                else
                {
                    view.Chk = false;

                    if (listOutWareReqPrint.Contains(view) == false)
                        listOutWareReqPrint.Remove(view);
                        ovcListOutWareReq.Remove(view);
                }

                if(ovcListOutWareReq.Count > 0 )
                {
                    if (ovcListOutWareReq.Count > 0 && chkCloseReq.IsChecked == true)
                    {
                        bool flag = false;
                        foreach (Win_ord_OutwareReq_U_View exception in ovcListOutWareReq)
                        {
                            if (exception.ReqCloseClss.Equals("1"))
                                flag = true;
                        }
                        if (flag == true)
                        {
                            btnCloseReq.IsEnabled = true;
                            tbxCloseReq.Text = "지시마감취소";
                        }
                    }
                    else if (ovcListOutWareReq.Count > 0 && chkCloseReq.IsChecked == false)
                    {

                        bool flag = false;
                        foreach (Win_ord_OutwareReq_U_View exception in ovcListOutWareReq)
                        {
                            if (exception.ReqCloseClss.Equals("1"))
                                flag = true;
                        }
                        if (flag == true)
                        {
                            btnCloseReq.IsEnabled = true;
                            tbxCloseReq.Text = "지시마감취소";
                        }
                        else
                        {
                            btnCloseReq.IsEnabled = false;
                            tbxCloseReq.Text = "지시마감처리";
                        }
                    
                    }

                    btnCloseReq.IsEnabled = true;
                }
                else
                {
                    btnCloseReq.IsEnabled =false;
                 
                }

               
                
            }
        }

        private void dgdOutwareSub_btnDelete_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var OutwareSub = dgdSub.SelectedItem as Win_ord_OutwareReqSub_U_View;
                if (OutwareSub != null)
                    dgdSub.Items.Remove(OutwareSub);
            }
            catch (Exception ee)
            {
                MessageBox.Show("오류지점 - dgdOutwareSub_btnDelete_Click : " + ee.ToString());
            }
        }

        private void DataGridTextBoxReqQty_PreviewTextInput(object sender, TextCompositionEventArgs e)
        {
            try
            {
                Lib.Instance.CheckIsNumeric((TextBox)sender, e);
            }
            catch (Exception ee)
            {
                MessageBox.Show("오류지점 - DataGridTextBoxReqQty_PreviewTextInput : " + ee.ToString());
            }
        }

        private void DataGridTextBoxReqQty_TextChanged(object sender, TextChangedEventArgs e)
        {
            try
            {
              var OutwareSub = dgdSub.SelectedItem as Win_ord_OutwareReqSub_U_View;
                if (OutwareSub != null)
                {
                    TextBox tbReqQty = sender as TextBox;

                    //var stockQty = lib.returnDouble(OutwareSub.StockQty.ToString());
                    var reqQty = lib.returnDouble(tbReqQty.Text.ToString());
                    //var remainQty = stockQty - reqQty;
                    var orderQty = lib.returnDouble(OutwareSub.OrderQty.ToString());
                    var outQty = lib.returnDouble(OutwareSub.OutQty.ToString());
                    var remainOrderQty = orderQty - outQty;
                    var remainOrderQtyLast = remainOrderQty - reqQty;


                    OutwareSub.StockQty = lib.returnNumString(OutwareSub.StockQty);
                    OutwareSub.ReqQty = lib.returnNumString(OutwareSub.ReqQty);
                    //OutwareSub.RemainQty = lib.returnNumString(remainQty.ToString());
                    OutwareSub.OrderQty = lib.returnNumString(OutwareSub.OrderQty);
                    OutwareSub.RestOrderQty = lib.returnNumString(remainOrderQtyLast.ToString());

                    tbReqQty.SelectionStart = tbReqQty.Text.Length;
                    tbReqQty.SelectionLength = 0;
                }
            }
            catch (Exception ee)
            {
                MessageBox.Show("오류지점 - DataGridTextBoxColorQty_TextChanged : " + ee.ToString());
            }
        }
        #endregion 서브그리드 메서드

        #region 기타메서드
        private void BtnControl(bool isCan)
        {
            // 버튼 그룹
            btnAdd.IsEnabled = isCan;
            btnUpdate.IsEnabled = isCan;
            btnDelete.IsEnabled = isCan;
            btnSearch.IsEnabled = isCan;
            btnSave.Visibility = isCan ? Visibility.Hidden : Visibility.Visible;
            btnCancel.Visibility = isCan ? Visibility.Hidden : Visibility.Visible;
            btnExcel.Visibility = isCan ? Visibility.Visible : Visibility.Hidden;
            btnPrint.Visibility = isCan ? Visibility.Visible : Visibility.Hidden;

            dgdMain.IsHitTestVisible = isCan;
            dgdSub.IsHitTestVisible = !isCan;

            grdInput.IsHitTestVisible = !isCan;
            grdInput.IsEnabled = !isCan;
            lblMsg.Visibility = isCan ? Visibility.Hidden : Visibility.Visible;
            btnTarget.IsEnabled = !isCan;

            reqChkOrderDay.IsChecked = true;
        }

        private void btnTarget_Click(object sender, RoutedEventArgs e)
        {
            /*if (string.IsNullOrEmpty(reqTxtCustom.Text) || reqTxtCustom.Tag == null)
                MessageBox.Show("거래처가 입력되지 않았습니다. 먼저 거래처를 입력해주세요");*/

            Win_pop_OutwareReq popup = SetPopupAfterOpen();
            if (popup.DialogResult == false)
                return;

            int count = 0;
            string msg = "", exist = "";

            List<Win_ord_OutwareReqSub_U_View> listReq = popup.GetList();
            for (int i = 0; i < listReq.Count; i++)
            {
                var req = listReq[i];
                if (ExistTargetArticle(req.ArticleID))
                {
                    exist += req.Article + "\r";
                    continue;
                }

                req.Num = dgdSub.Items.Count;
                req.ReqQty = "0";
                req.RemainQty = req.StockQty;

                dgdSub.Items.Add(req);
                count++;
            }

            if (exist.Length > 0)
            {
                msg = exist + "위의 품명은 이미 등록되어 있습니다.";
                if (count > 0)
                    msg += "\r(위의 품명을 제외하고 추가되었습니다)";
            }
            else
                msg = "출고대상이 추가되었습니다.";

            MessageBox.Show(msg);
        }

        private Win_pop_OutwareReq SetPopupAfterOpen()
        {
            Win_pop_OutwareReq popup = new Win_pop_OutwareReq();
            //popup.chkDate = reqChkOrderDay.IsChecked == true ? 1 : 0;
            //popup.startDate = reqChkOrderDay.IsChecked == true ? reqDtpSDate.SelectedDate.Value.ToString("yyyyMMdd") : "";
            //popup.endDate = reqChkOrderDay.IsChecked == true ? reqDtpEDate.SelectedDate.Value.ToString("yyyyMMdd") : "";
            //popup.chkCustomID = (reqTxtCustom.Tag != null && reqTxtCustom.Text.Length > 0) ? 1 : 0;
            //popup.customID = (reqTxtCustom.Tag != null ? reqTxtCustom.Tag.ToString() : "");

            popup.ShowDialog();
            return popup;
        }

        private bool ExistTargetArticle(string articldID)
        {
            bool find = false;
            for (int i = 0; i < dgdSub.Items.Count; i++)
            {
                var sub = dgdSub.Items[i] as Win_ord_OutwareReqSub_U_View;
                if (sub == null)
                    continue;

                if (sub.ArticleID == articldID)
                {
                    find = true;
                    break;
                }
            }

            return find;
        }

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
            else if (chkInCustom.IsChecked == true && string.IsNullOrEmpty(txtInCustom.Text)) word = "최종고객사";
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

        // 천자리 콤마, 소수점 버리기
        private string stringFormatN0(object obj)
        {
            return string.Format("{0:N0}", obj);
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
        #endregion 기타메서드        

        private void btnCloseReq_Click(object sender, RoutedEventArgs e)
        {
         

            if (tbxCloseReq.Text.Equals("지시마감처리"))
            {
                if (MessageBox.Show("선택하신 지시를 마감하시겠습니까?", "마감 전 확인", MessageBoxButton.YesNo) == MessageBoxResult.Yes)
                {
                    reqFlag = "1";
                    using (Loading lw = new Loading(beCloseReq))
                    {
                        lw.ShowDialog();
                    }
                }
            }
            else if (tbxCloseReq.Text.Equals("지시마감취소"))
            {
                if (MessageBox.Show("선택하신 지시를 복원하시겠습니까?", "복원 전 확인", MessageBoxButton.YesNo) == MessageBoxResult.Yes)
                {
                    reqFlag = "2";
                    using (Loading lw = new Loading(beCloseReq))
                    {
                        lw.ShowDialog();
                    }
                }
            }

        }

        private void beCloseReq()
        {
            btnCloseReq.IsEnabled = false;

            Dispatcher.BeginInvoke(new Action(() =>
            {
                foreach (Win_ord_OutwareReq_U_View closeReq in ovcListOutWareReq)
                    CloseReq(closeReq.OutwareReqID, reqFlag);

                rowNum = 0;
                reqFlag = string.Empty;
                re_Search();
            }), System.Windows.Threading.DispatcherPriority.Background);

            btnCloseReq.IsEnabled = true;
        }

        private void lblCloseReq_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            if(chkCloseReq.IsChecked == true)
            {
                chkCloseReq.IsChecked = false;
            }
            else
            {
                chkCloseReq.IsChecked = true;
            }
        }
    }

    #region View 클래스
    class Win_ord_OutwareReq_U_View : BaseView
    {
        public int Num { get; set; }
        public bool Chk { get; set; }
        public string OutwareReqID { get; set; }
        public string ArticleID { get; set; }
        public string Article { get; set; }
        public string BuyerArticleNo { get; set; }
        public string CustomID { get; set; }
        public string KCustom { get; set; }
        public string OutClss { get; set; }
        public string ReqDate { get; set; }
        public string ReqQty { get; set; }
        public string ReqCloseClss { get; set; }
        public string OutWareReqTypeID { get; set; }
        public string RegDate { get; set; }
        public string RegUserID { get; set; }
        public string Remark { get; set; }
    }

    public class Win_ord_OutwareReqSub_U_View : BaseView
    {
        public int Num { get; set; }
        public bool Chk { get; set; }
        public string LocName { get; set; }
        public string CustomID { get; set; }
        public string CustomName { get; set; }
        public string AcptDate { get; set; }
        public string DvlyDate { get; set; }
        public string OrderQty { get; set; }
        public string RestOrderQty { get; set; }

        public string KCustom { get; set; }
        public string InCustomID { get; set; }
        public string KInCustom { get; set; }
        public string OrderID { get; set; }
        public string OrderNo { get; set; }
        public string OutwareReqID { get; set; }
        public string OutwareReqSeq { get; set; }
        public string ArticleID { get; set; }
        public string Article { get; set; }
        public string ArticleGrp { get; set; }
        public string BuyerArticleNo { get; set; }
        public string StuffINQty { get; set; }
        public string OutQty { get; set; }
        public string StockQty { get; set; }   
        public string ReqQty { get; set; }
        public string ReqEnd { get; set; }
        public string RemainQty { get; set; }
        public string UnitPrice { get; set; }
        public string Remark { get; set; }
    }

    class Win_ord_OutwareReqTot_U_View : BaseView
    {
        public string ReqCount { get; set; }
        public string ReqAmount { get; set; }
    }
    #endregion view 클래스
}

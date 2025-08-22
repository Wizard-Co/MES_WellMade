using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Data;
using System.Reflection;
using System.Threading;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Input;
using System.Windows.Media;
using WizMes_WellMade.PopUP;

namespace WizMes_WellMade
{
    /// <summary>
    /// Win_mtr_LotStock_Q.xaml에 대한 상호 작용 논리
    /// </summary>
    public partial class Win_mtr_LotStock_Q : UserControl
    {
        string stDate = string.Empty;
        string stTime = string.Empty;

        Lib lib = new Lib();
        int rowNum = 0;

        string dgdFlag = "";

        ScrollViewer scrollView = null;
        ScrollViewer scrollView2 = null;
        ScrollViewer scrollView3 = null;
        ScrollViewer scrollView4 = null;

        private Microsoft.Office.Interop.Excel.Application excel;
        private Microsoft.Office.Interop.Excel.Workbook workBook;
        private Microsoft.Office.Interop.Excel.Worksheet workSheet;
        private Microsoft.Office.Interop.Excel.Range cellRange;

        // 병합시킬 인덱스 모음 - 안의 값 : 첫번째 값은 컬럼번호, 나머진 행높이 
        List<List<int>> lstMerge = new List<List<int>>();

        int startIndex = 0;


        public Win_mtr_LotStock_Q()
        {
            InitializeComponent();
        }

        public void UserControl_Loaded(object sender, RoutedEventArgs e)
        {
            stDate = DateTime.Now.ToString("yyyyMMdd");
            stTime = DateTime.Now.ToString("HHmm");

            DataStore.Instance.InsertLogByFormS(this.GetType().Name, stDate, stTime, "S");

            Lib.Instance.UiLoading(sender);

            dtpSDate.SelectedDate = DateTime.Today;
            dtpEDate.SelectedDate = DateTime.Today;


            chkDate.IsChecked = true;
            //cboWareHouse.IsEnabled = false; //2021-07-09
            chkIncNotApproval.IsChecked = true;
            chkIncAutoInOutWare.IsChecked = true;
            chknIncZeroQty.IsChecked = false;

            ComboBoxSetting();

            // 스크롤
            scrollView = getScrollbar(dgdNum);
            scrollView2 = getScrollbar(dgdArticle);
            scrollView3 = getScrollbar(dgdLotID);
            scrollView4 = getScrollbar(dgdContent);

            if (null != scrollView)
            {
                scrollView.ScrollChanged += new ScrollChangedEventHandler(scrollView_ScrollChanged);
            }
            if (null != scrollView2)
            {
                scrollView2.ScrollChanged += new ScrollChangedEventHandler(scrollView_ScrollChanged);
            }
            if (null != scrollView3)
            {
                scrollView3.ScrollChanged += new ScrollChangedEventHandler(scrollView_ScrollChanged);
            }
            if (null != scrollView4)
            {
                scrollView4.ScrollChanged += new ScrollChangedEventHandler(scrollView_ScrollChanged);
            }


        }

        #region 일자변경

        //검사일자 라벨 이벤트
        private void lblDate_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            if (chkDate.IsChecked == true)
            {
                chkDate.IsChecked = false;
            }
            else
            {
                chkDate.IsChecked = true;
            }
        }
        //검사일자 체크
        private void chkDate_Checked(object sender, RoutedEventArgs e)
        {
            chkDate.IsChecked = true;

            dtpSDate.IsEnabled = true;
            dtpEDate.IsEnabled = true;

            btnYesterDay.IsEnabled = true;
            btnToday.IsEnabled = true;
            btnLastMonth.IsEnabled = true;
            btnThisMonth.IsEnabled = true;

        }
        //검사일자 체크해제
        private void chkDate_Unchecked(object sender, RoutedEventArgs e)
        {
            chkDate.IsChecked = false;

            dtpSDate.IsEnabled = false;
            dtpEDate.IsEnabled = false;

            btnYesterDay.IsEnabled = false;
            btnToday.IsEnabled = false;
            btnLastMonth.IsEnabled = false;
            btnThisMonth.IsEnabled = false;
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

        #region 상단 레이아웃 활성화 & 비활성화

        // 거래처 검색 라벨 이벤트
        private void lblCustom_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            if (chkCustomSrh.IsChecked == true)
            {
                chkCustomSrh.IsChecked = false;
            }
            else
            {
                chkCustomSrh.IsChecked = true;
            }
        }

        //거래처 체크
        private void chkCustomSrh_Checked(object sender, RoutedEventArgs e)
        {
            chkCustomSrh.IsChecked = true;

            txtCustomSrh.IsEnabled = true;
            btnCustomSrh.IsEnabled = true;
        }

        //거래처 체크해제
        private void chkCustomSrh_Unchecked(object sender, RoutedEventArgs e)
        {
            chkCustomSrh.IsChecked = false;

            txtCustomSrh.IsEnabled = false;
            btnCustomSrh.IsEnabled = false;
        }
        // 거래처 엔터 → 플러스파인더
        private void txtCustomSrh_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                MainWindow.pf.ReturnCode(txtCustomSrh, 0, "");
            }
        }
        // 거래처 플러스파인더
        private void btnCustomSrh_Click(object sender, RoutedEventArgs e)
        {
            MainWindow.pf.ReturnCode(txtCustomSrh, 0, "");
        }

        // 자재품명 라벨 이벤트
        private void lblSrhMtrArticle_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            if (chkMtrArticleSrh.IsChecked == true)
            {
                chkMtrArticleSrh.IsChecked = false;
            }
            else
            {
                chkMtrArticleSrh.IsChecked = true;
            }
        }
        //자재품명 체크
        private void chkMtrArticleSrh_Checked(object sender, RoutedEventArgs e)
        {
            chkMtrArticleSrh.IsChecked = true;

            txtMtrArticleSrh.IsEnabled = true;
            btnMtrArticleSrh.IsEnabled = true;
        }
        //자재품명 체크해제
        private void chkMtrArticleSrh_UnChecked(object sender, RoutedEventArgs e)
        {
            chkMtrArticleSrh.IsChecked = false;

            txtMtrArticleSrh.IsEnabled = false;
            btnMtrArticleSrh.IsEnabled = false;
        }
        // 자재품명 엔터 → 플러스파인더
        private void txtMtrArticleSrh_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                MainWindow.pf.ReturnCode(txtMtrArticleSrh, 76, "");
            }
        }
        // 자재품명 플러스파인더
        private void btnMtrArticleSrh_Click(object sender, RoutedEventArgs e)
        {
            MainWindow.pf.ReturnCode(txtMtrArticleSrh, 76, "");
        }

        //자재품명 초기화
        private void btnMtrArticleSrhClean_Click(object sender, RoutedEventArgs e)
        {
            txtMtrArticleSrh.Clear();
        }

        // 제품품명 라벨 이벤트
        private void lblSrhProdArticle_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            if (chkProdArticleSrh.IsChecked == true)
            {
                chkProdArticleSrh.IsChecked = false;
            }
            else
            {
                chkProdArticleSrh.IsChecked = true;
            }
        }
        //제품품명 체크
        private void chkProdArticleSrh_Checked(object sender, RoutedEventArgs e)
        {
            chkProdArticleSrh.IsChecked = true;

            txtProdArticleSrh.IsEnabled = true;
            btnProdArticleSrh.IsEnabled = true;
        }

        //제품품명 체크해제
        private void chkProdArticleSrh_UnChecked(object sender, RoutedEventArgs e)
        {
            chkProdArticleSrh.IsChecked = false;

            txtProdArticleSrh.IsEnabled = false;
            btnProdArticleSrh.IsEnabled = false;
        }
        // 제품품명 엔터 → 플러스파인더
        private void txtProdArticleSrh_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                MainWindow.pf.ReturnCode(txtProdArticleSrh, 76, "");
            }
        }
        //제품품명 플러스파인더
        private void btnProdArticleSrh_Click(object sender, RoutedEventArgs e)
        {
            MainWindow.pf.ReturnCode(txtProdArticleSrh, 76, "");
        }

        //입고미승인건 라벨 버튼
        private void lblIncNotApproval_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            if (chkIncNotApproval.IsChecked == true)
            {
                chkIncNotApproval.IsChecked = false;
            }
            else
            {
                chkIncNotApproval.IsChecked = true;
            }
        }
        //입고미승인건 체크
        private void chkIncNotApproval_Checked(object sender, RoutedEventArgs e)
        {
            chkIncNotApproval.IsChecked = true;
        }

        //입고미승인건 체크해제
        private void chkIncNotApproval_UnChecked(object sender, RoutedEventArgs e)
        {
            chkIncNotApproval.IsChecked = false;
        }


        // 자동 입출건 포함 라벨 이벤트
        private void lblIncAutoInOutWare_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            if (chkIncAutoInOutWare.IsChecked == true)
            {
                chkIncAutoInOutWare.IsChecked = false;
            }
            else
            {
                chkIncAutoInOutWare.IsChecked = true;
            }
        }

        //자동입출고건 체크
        private void chkIncAutoInOutWare_Checked(object sender, RoutedEventArgs e)
        {
            chkIncAutoInOutWare.IsChecked = true;
        }

        //자동입출고건 체크해제
        private void chkIncAutoInOutWare_UnChecked(object sender, RoutedEventArgs e)
        {
            chkIncAutoInOutWare.IsChecked = false;
        }


        // 주요 관심품목 라벨 이벤트
        private void lblMainItem_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            if (chkMainItem.IsChecked == true)
            {
                chkMainItem.IsChecked = false;
            }
            else
            {
                chkMainItem.IsChecked = true;
            }
        }
        //주요관심품목 체크
        private void chkMainItem_Checked(object sender, RoutedEventArgs e)
        {
            chkMainItem.IsChecked = true;
        }

        //주요관심품목 체크해제
        private void chkMainItem_UnChecked(object sender, RoutedEventArgs e)
        {
            chkMainItem.IsChecked = false;
        }


        // 거래처별 관심 품목 라벨 이벤트
        private void lblCustomItem_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            if (chkCustomItem.IsChecked == true)
            {
                chkCustomItem.IsChecked = false;
            }
            else
            {
                chkCustomItem.IsChecked = true;
            }
        }
        //거래처별등록품목 체크
        private void chkCustomItem_Checked(object sender, RoutedEventArgs e)
        {
            chkCustomItem.IsChecked = true;
        }

        //거래처별등록품목 체크해제
        private void chkCustomItem_UnChecked(object sender, RoutedEventArgs e)
        {
            chkCustomItem.IsChecked = false;
        }

        // 재고 0 포함 라벨 버튼 이벤트
        private void lblnIncZeroQty_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            if (chknIncZeroQty.IsChecked == true)
            {
                chknIncZeroQty.IsChecked = false;
            }
            else
            {
                chknIncZeroQty.IsChecked = true;
            }
        }
        //재고 0 포함 체크
        private void chknIncZeroQty_Checked(object sender, RoutedEventArgs e)
        {
            chknIncZeroQty.IsChecked = true;
        }
        //재고 0 포함 체크해제
        private void chknIncZeroQty_UnChecked(object sender, RoutedEventArgs e)
        {
            chknIncZeroQty.IsChecked = false;
        }
        #endregion

        #region 우측 상단 버튼
        //검색 버튼
        private void btnSearch_Click(object sender, RoutedEventArgs e)
        {
            //검색버튼 비활성화
            btnSearch.IsEnabled = false;

            Dispatcher.BeginInvoke(new Action(() =>

            {
                Thread.Sleep(2000);

                //로직
                rowNum = 0;
                re_Search(rowNum);

            }), System.Windows.Threading.DispatcherPriority.Background);



            Dispatcher.BeginInvoke(new Action(() =>

            {
                btnSearch.IsEnabled = true;

            }), System.Windows.Threading.DispatcherPriority.Background);
        }

        //닫기 버튼
        private void btnClose_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                DataStore.Instance.InsertLogByFormS(this.GetType().Name, stDate, stTime, "E");
                Lib.Instance.ChildMenuClose(this.ToString());
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        //인쇄버튼
        //private void cmdPrint_Click(object sender, RoutedEventArgs e)
        //{
        //    System.Windows.Controls.PrintDialog printdlg = new System.Windows.Controls.PrintDialog();
        //    if ((bool)printdlg.ShowDialog().GetValueOrDefault())
        //    {
        //        Size pageSize = new Size(printdlg.PrintableAreaWidth, printdlg.PrintableAreaHeight);

        //        dgdMain.Measure(pageSize);
        //        dgdMain.Arrange(new Rect(5, 5, pageSize.Width, pageSize.Height));
        //        printdlg.PrintVisual(dgdMain, Name);

        //    }
        //}

        //엑셀버튼
        private void cmdExcel_Click(object sender, RoutedEventArgs e)
        {

            // 데이터 그리드 4개를 합쳐서 DataTable 에 넣어야됨.
            try
            {
                DataGrid dgd = new DataGrid();
                dgd = dgdMain;
                //dgd = dgdNum;

                if (dgd.Items.Count == 0)
                {
                    dgd = dgdNum;
                    DataTable dt = null;
                    string Name = string.Empty;

                    string[] lst = new string[2];
                    lst[0] = "LoT별 수불 조회";
                    lst[1] = dgd.Name;

                    ExportExcelxaml ExpExc = new ExportExcelxaml(lst);

                    ExpExc.ShowDialog();

                    if (ExpExc.DialogResult.HasValue)
                    {
                        if (ExpExc.choice.Equals(dgd.Name))
                        {
                            DataStore.Instance.InsertLogByForm(this.GetType().Name, "E");
                            //if (ExpExc.Check.Equals("Y"))
                            //    dt = Lib.Instance.DataGridToDTinHidden(dgd);
                            //else
                            //    dt = Lib.Instance.DataGirdToDataTable(dgd);


                            dt = getDataTableIn4Dgd(dgdNum, dgdArticle, dgdLotID, dgdContent);
                            dt.Columns.Add("");

                            Name = dgd.Name;

                            //if (Lib.Instance.GenerateExcel(dt, Name))
                            //    Lib.Instance.excel.Visible = true;
                            //else
                            //    return;
                            if (GenerateExcel(dt, Name, lstMerge))
                            {
                                excel.Visible = true;
                            }
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
                else
                {
                    DataTable dt = null;
                    string Name = string.Empty;

                    string[] lst = new string[2];
                    lst[0] = "LoT별 수불 조회";
                    lst[1] = dgd.Name;

                    ExportExcelxaml ExpExc = new ExportExcelxaml(lst);

                    ExpExc.ShowDialog();

                    if (ExpExc.DialogResult.HasValue)
                    {
                        if (ExpExc.choice.Equals(dgd.Name))
                        {
                            if (ExpExc.Check.Equals("Y"))
                                dt = Lib.Instance.DataGridToDTinHidden(dgd);
                            else
                                dt = Lib.Instance.DataGirdToDataTable(dgd);


                            //dt = getDataTableIn4Dgd(dgdNum, dgdArticle, dgdLotID, dgdContent);
                            //dt.Columns.Add("");

                            Name = dgd.Name;

                            if (Lib.Instance.GenerateExcel(dt, Name))
                                Lib.Instance.excel.Visible = true;
                            else
                                return;
                            //if (GenerateExcel(dt, Name, lstMerge))
                            //{
                            //    excel.Visible = true;
                            //}
                            //else
                            //    return;
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
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private bool GenerateExcel(System.Data.DataTable dt, string Name, List<List<int>> lstMerge)
        {
            bool result = true;

            try
            {
                if (dt != null && dt.Rows.Count > 0)
                {
                    excel = new Microsoft.Office.Interop.Excel.Application();
                    excel.DisplayAlerts = false;
                    excel.Visible = false;

                    workBook = excel.Workbooks.Add(Type.Missing);
                    workSheet = (Microsoft.Office.Interop.Excel.Worksheet)workBook.ActiveSheet;
                    workSheet.Name = Name;

                    System.Data.DataTable datatable = dt;
                    workSheet.Cells.Font.Size = 11;

                    int rowCount = 1;
                    for (int i = 1; i < datatable.Columns.Count; i++)
                    {
                        workSheet.Cells[1, i] = datatable.Columns[i - 1].ColumnName;
                    }

                    foreach (DataRow row in datatable.Rows)
                    {
                        rowCount += 1;
                        for (int i = 0; i < datatable.Columns.Count; i++)
                        {
                            workSheet.Cells[rowCount, i + 1] = row[i].ToString();
                        }
                    }

                    // 셀 병합 메서드 추가
                    foreach (List<int> merge in lstMerge)
                    {
                        int column = merge[0] + 1;

                        int startRow = 2; // 헤더부분 건너뛰고
                        int endRow = 0;
                        for (int i = 1; i < merge.Count; i++)
                        {
                            endRow = startRow + merge[i] - 1;

                            cellRange = workSheet.Range[workSheet.Cells[startRow, column], workSheet.Cells[endRow, column]];
                            cellRange.Merge();

                            startRow = endRow + 1;
                        }
                    }

                    cellRange = workSheet.Range[workSheet.Cells[1, 1], workSheet.Cells[rowCount, datatable.Columns.Count]];
                    cellRange.EntireColumn.AutoFit();

                    return result;
                }
                else
                {
                    MessageBox.Show("엑셀로 내보낼 자료가 없습니다.");
                    result = false;
                    return result;
                }
            }
            catch (Exception)
            {
                throw;
            }
        }


        private DataTable getDataTableIn4Dgd(DataGrid dgd1, DataGrid dgd2, DataGrid dgd3, DataGrid dgd4)
        {
            DataTable dt = new DataTable();


            // 첫번째 값은 컬럼번호, 나머진 행높이
            lstMerge.Clear();
            List<int> dgd2Merge = new List<int>();
            List<int> dgd3Merge = new List<int>();

            dgd2Merge.Add(1);
            dgd3Merge.Add(2);

            // 2, 3 의 행 높이가 다름
            for (int i = 0; i < dgd2.Items.Count; i++)
            {
                var ProcessMerge = dgd2.Items[i] as LotStock;
                dgd2Merge.Add(ProcessMerge.Count);
            }
            // 2, 3 의 행 높이가 다름
            for (int i = 0; i < dgd3.Items.Count; i++)
            {
                var ProcessMerge = dgd3.Items[i] as LotStock;
                dgd3Merge.Add(ProcessMerge.Count);
            }

            lstMerge.Add(dgd2Merge);
            lstMerge.Add(dgd3Merge);

            List<DataGridColumn> listVisibleDataGridColumns = new List<DataGridColumn>();

            // 헤더 세팅
            for (int i = 0; i < dgd1.Columns.Count; i++)
            {
                dt.Columns.Add(dgd1.Columns[i].Header.ToString());
                listVisibleDataGridColumns.Add(dgd1.Columns[i]);
            }
            for (int i = 0; i < dgd2.Columns.Count; i++)
            {
                dt.Columns.Add(dgd2.Columns[i].Header.ToString());
                listVisibleDataGridColumns.Add(dgd2.Columns[i]);
            }
            for (int i = 0; i < dgd3.Columns.Count; i++)
            {
                dt.Columns.Add(dgd3.Columns[i].Header.ToString());
                listVisibleDataGridColumns.Add(dgd3.Columns[i]);
            }
            for (int i = 0; i < dgd4.Columns.Count; i++)
            {
                dt.Columns.Add(dgd4.Columns[i].Header.ToString());
                listVisibleDataGridColumns.Add(dgd4.Columns[i]);
            }


            // 첫번째 그리드 변수
            int indexDgd1 = 0;
            int stackDgd1 = 0;
            int EndIndex_Dgd1 = 0;

            // 세번째 그리드 변수
            int indexDgd3 = 0;
            int stackDgd3 = 0;
            int EndIndex_Dgd3 = 0;

            // 그리드 객체들
            var Num = new LotStock_Content();
            var Article = new LotStock();
            var LotID = new LotStock();
            var Content = new LotStock_Content();

            // 실제 데이터가 1행씩 되있어, 반복은 데이터 그리드로!!!!
            for (int i = 0; i < dgd4.Items.Count; i++)
            {
                // 이제 데이터 세팅 → 셀이 병합되 있는 경우에는??
                DataRow dr = dt.NewRow();

                Num = dgd1.Items[i] as LotStock_Content;

                // 첫번째 그리드 : 병합
                if (stackDgd1 == 0)
                {
                    Article = dgd2.Items[indexDgd1] as LotStock;
                    EndIndex_Dgd1 = Article.Count - 1;


                    indexDgd1++;
                }

                // 두번째 그리드 : 병합 없음


                // 세번째 그리드 : 병합
                if (stackDgd3 == 0)
                {
                    LotID = dgd3.Items[indexDgd3] as LotStock;
                    EndIndex_Dgd3 = LotID.Count - 1;

                    indexDgd3++;
                }

                // 네번째 그리드 : 병합 없음
                Content = dgd4.Items[i] as LotStock_Content;


                ////////////////////
                // 데이터 넣어주기
                ////////////////////
                startIndex = 0;
                dr = setDataRow(Num, listVisibleDataGridColumns, dr, dgd1.Columns.Count);
                //dr = setDataRow(Gubun, listVisibleDataGridColumns, dr, getDataRowIndex(dr), dgd2.Columns.Count);
                //dr = setDataRow(CarryOver, listVisibleDataGridColumns, dr, getDataRowIndex(dr), dgd3.Columns.Count);
                //dr = setDataRow(Content, listVisibleDataGridColumns, dr, getDataRowIndex(dr), dgd4.Columns.Count);

                dr = setDataRow(Article, listVisibleDataGridColumns, dr, dgd2.Columns.Count);
                dr = setDataRow(LotID, listVisibleDataGridColumns, dr, dgd3.Columns.Count);
                dr = setDataRow(Content, listVisibleDataGridColumns, dr, dgd4.Columns.Count);

                dt.Rows.Add(dr);


                // 첫번째 그리드
                if (stackDgd1 != EndIndex_Dgd1)
                {
                    stackDgd1++;
                }
                else
                {
                    stackDgd1 = 0;
                }

                // 세번째 그리드
                if (stackDgd3 != EndIndex_Dgd3)
                {
                    stackDgd3++;
                }
                else
                {
                    stackDgd3 = 0;
                }

            }


            return dt;
        }

        private int getDataRowIndex(DataRow dr)
        {
            int index = 0;

            foreach (var value in dr.ItemArray)
            {
                if (value == null || value.ToString().Trim().Equals(""))
                    break;

                index++;
            }

            return index;
        }

        // data : 데이트그리드 객체
        // listVisibleDataGridColumns : 헤더명 모음
        // startIndex : 헤더 시작 인덱스
        // colCnt : 해당 데이터 그리드의 컬럼 수
        private DataRow setDataRow(object data, List<DataGridColumn> listVisibleDataGridColumns, DataRow dr, int colCnt)
        {
            // 함수
            for (int i = 0; i < colCnt; i++)
            {
                DataGridColumn dataGridColumn = listVisibleDataGridColumns[startIndex];

                string strValue = string.Empty;
                Binding objBinding = null;
                DataGridBoundColumn dataGridBoundColumn = dataGridColumn as DataGridBoundColumn;

                if (dataGridBoundColumn != null)
                {
                    objBinding = dataGridBoundColumn.Binding as Binding;
                }

                DataGridTemplateColumn dataGridTemplateColumn = dataGridColumn as DataGridTemplateColumn;

                if (dataGridTemplateColumn != null)
                {
                    DependencyObject dependencyObject = dataGridTemplateColumn.CellTemplate.LoadContent();

                    FrameworkElement frameworkElement = dependencyObject as FrameworkElement;
                    if (frameworkElement != null)
                    {
                        FieldInfo fieldInfo = frameworkElement.GetType().GetField("ContentProperty", BindingFlags.Public | BindingFlags.Static | BindingFlags.FlattenHierarchy);
                        if (fieldInfo == null)
                        {
                            if (frameworkElement is System.Windows.Controls.TextBox || frameworkElement is TextBlock || frameworkElement is ComboBox)
                            {
                                fieldInfo = frameworkElement.GetType().GetField("TextProperty");
                            }
                            else if (frameworkElement is DatePicker)
                            {
                                fieldInfo = frameworkElement.GetType().GetField("SelectedDateProperty");
                            }
                            else if (frameworkElement is Grid)
                            {
                                fieldInfo = (frameworkElement as Grid).Children.GetType().GetField("TextProperty");
                            }
                            else if (frameworkElement is System.Windows.Controls.DockPanel) // DockPanel 일때 - 2019.10.10 아무개 추가
                            {
                                foreach (UIElement Item in (frameworkElement as DockPanel).Children)
                                {

                                    if (Item is DatePicker)
                                    {
                                        BindingExpression bindingExpression = (Item as DatePicker).GetBindingExpression(DatePicker.TextProperty);

                                        if (bindingExpression != null)
                                        {
                                            objBinding = bindingExpression.ParentBinding;
                                        }
                                    }
                                    else if (Item is ComboBox)
                                    {
                                        BindingExpression bindingExpression = (Item as ComboBox).GetBindingExpression(ComboBox.ToolTipProperty);

                                        if (bindingExpression != null)
                                        {
                                            objBinding = bindingExpression.ParentBinding;
                                        }
                                    }


                                }
                            }
                        }

                        // objBinding = 해당 요소의 바인딩 정보(바인딩 이름 등등) 가져오기
                        if (fieldInfo != null)
                        {
                            DependencyProperty dependencyProperty = fieldInfo.GetValue(null) as DependencyProperty;
                            if (dependencyProperty != null)
                            {
                                BindingExpression bindingExpression = frameworkElement.GetBindingExpression(dependencyProperty);
                                if (bindingExpression != null)
                                {
                                    objBinding = bindingExpression.ParentBinding;
                                }
                            }
                        }
                    }
                }

                if (objBinding != null)
                {
                    if (!String.IsNullOrEmpty(objBinding.Path.Path))
                    {
                        PropertyInfo pi = data.GetType().GetProperty(objBinding.Path.Path);

                        if (pi != null)
                        {
                            object propValue = pi.GetValue(data, null);

                            if (propValue != null)
                            {
                                strValue = Convert.ToString(propValue);
                            }

                            else
                            {
                                strValue = string.Empty;
                            }
                        }
                    }

                    if (objBinding.Converter != null)
                    {
                        if (!String.IsNullOrEmpty(strValue))
                        {
                            strValue = objBinding.Converter.Convert(strValue, typeof(string), objBinding.ConverterParameter, objBinding.ConverterCulture).ToString();
                        }
                        else
                        {
                            strValue = objBinding.Converter.Convert(data, typeof(string), objBinding.ConverterParameter, objBinding.ConverterCulture).ToString();
                        }
                    }

                    if (strValue != string.Empty)
                    {
                        dr[startIndex] = strValue;
                    }
                }

                startIndex++;
            }

            return dr;
        }
        #endregion

        #region Content 부분

        // 순번 데이터 그리드
        private void dgdNum_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (dgdFlag.Equals(""))
                dgdFlag = "Num";

            if (dgdFlag.Equals("Num"))
            {
                int sIndex = dgdNum.SelectedIndex;

                var LotStock = dgdNum.SelectedItem as LotStock_Content;

                if (LotStock != null)
                {
                    int bIndex = LotStock.bIndex;
                    int mIndex = LotStock.mIndex;

                    if (dgdNum.SelectedIndex == 0)
                    {
                        bIndex = 0;
                        mIndex = 0;
                    }

                    // 품명 선택
                    if (dgdArticle.SelectedIndex != bIndex)
                        dgdArticle.SelectedIndex = bIndex;

                    // Lot 선택
                    if (dgdLotID.SelectedIndex != mIndex)
                        dgdLotID.SelectedIndex = mIndex;

                    // 데이터 선택
                    if (dgdContent.SelectedIndex != sIndex)
                        dgdContent.SelectedIndex = sIndex;

                    dgdFlag = "";
                }
            }
        }
        // 품명 데이터 그리드
        private void dgdArticle_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (dgdFlag.Equals(""))
                dgdFlag = "Article";

            if (dgdFlag.Equals("Article"))
            {

                var LotStock = dgdArticle.SelectedItem as LotStock;

                if (LotStock != null)
                {
                    int mIndex = LotStock.mIndex;
                    int mEndIndex = LotStock.mEndIndex;
                    int sIndex = LotStock.sIndex;
                    int sEndIndex = LotStock.sEndIndex;


                    if (dgdArticle.SelectedIndex == 0)
                    {
                        mIndex = 0;
                        sIndex = 0;
                    }

                    // Lot 선택
                    if (dgdLotID.SelectedIndex < mIndex || dgdLotID.SelectedIndex > mEndIndex)
                        dgdLotID.SelectedIndex = mIndex;

                    // 순번, 데이터 선택
                    if (dgdNum.SelectedIndex < sIndex || dgdNum.SelectedIndex > sEndIndex)
                    {
                        dgdNum.SelectedIndex = sIndex;
                        dgdContent.SelectedIndex = sIndex;
                    }

                    dgdFlag = "";
                }
            }
        }
        // 자재 LotID 데이터 그리드
        private void dgdLotID_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (dgdFlag.Equals(""))
                dgdFlag = "LotID";

            if (dgdFlag.Equals("LotID"))
            {
                var LotStock = dgdLotID.SelectedItem as LotStock;

                if (LotStock != null)
                {
                    int bIndex = LotStock.bIndex;
                    int sIndex = LotStock.sIndex;

                    int sEndIndex = sIndex + LotStock.Count - 1;

                    if (dgdLotID.SelectedIndex == 0)
                    {
                        bIndex = 0;
                        sIndex = 0;
                    }

                    // 품명 선택
                    if (dgdArticle.SelectedIndex != bIndex)
                        dgdArticle.SelectedIndex = bIndex;

                    // 순번, 데이터 선택
                    if (dgdNum.SelectedIndex < sIndex || dgdNum.SelectedIndex > sEndIndex)
                    {
                        dgdNum.SelectedIndex = sIndex;
                        dgdContent.SelectedIndex = sIndex;
                    }
                }

                dgdFlag = "";
            }
        }
        // 내용 데이터 그리드
        private void dgdContent_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (dgdFlag.Equals(""))
                dgdFlag = "Content";

            if (dgdFlag.Equals("Content"))
            {

                int sIndex = dgdContent.SelectedIndex;

                var LotStock = dgdContent.SelectedItem as LotStock_Content;

                if (LotStock != null)
                {
                    int bIndex = LotStock.bIndex;
                    int mIndex = LotStock.mIndex;

                    if (dgdContent.SelectedIndex == 0)
                    {
                        bIndex = 0;
                        mIndex = 0;
                    }

                    // 품명 선택
                    if (dgdArticle.SelectedIndex != bIndex)
                        dgdArticle.SelectedIndex = bIndex;

                    // Lot 선택
                    if (dgdLotID.SelectedIndex != mIndex)
                        dgdLotID.SelectedIndex = mIndex;

                    // 데이터 선택
                    if (dgdNum.SelectedIndex != sIndex)
                        dgdNum.SelectedIndex = sIndex;
                }

                dgdFlag = "";
            }
        }

        #endregion // Content 부분 

        #region 조회
        private void re_Search(int rowNum)
        {

            if (chkCustomSrh.IsChecked == true && txtCustomSrh.Text == "")
            {
                MessageBox.Show("거래처를 입력한 후 검색을 하거나 거래처 체크를 해제 후 검색 하세요");
                return;
            }
            else if (chkMtrArticleSrh.IsChecked == true && txtMtrArticleSrh.Text == "")
            {
                MessageBox.Show("품번를 입력한 후 검색을 하거나 품번 체크를 해제 후 검색 하세요");
                return;
            }
            else if (chkMtrLOTIDSrh.IsChecked == true && txtMtrLOTIDSrh.Text == "")
            {
                MessageBox.Show("LOTID를 입력한 후 검색을 하거나 LOTID 체크를 해제 후 검색 하세요");
                return;
            }


            FillGrid();


            if (dgdArticle.Items.Count > 0)
            {
                dgdArticle.SelectedIndex = rowNum;

                dgdNum.Focus();
                dgdNum.CurrentCell = dgdNum.SelectedCells[0];

            }
            else
            {
                this.DataContext = null;
            }
        }

        private void FillGrid()
        {
            int rowHeight = 30;

            int stack = 0; // 품명계가 두개가 나올수도 있는 경우를 위해서... 의미 있나..

            var AllLotStock = new LotStock_Collection();
            var Lot = new LotStock_All();

            if (dgdArticle.Items.Count > 0)
            {
                dgdArticle.Items.Clear();
                dgdLotID.Items.Clear();
                dgdNum.Items.Clear();
                dgdContent.Items.Clear();
            }

            dgdMain.Items.Clear();
            dgdTotal.Items.Clear();

            try
            {
                Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                sqlParameter.Clear();

                sqlParameter.Add("nChkDate", chkDate.IsChecked == true ? 1 : 0);
                //sqlParameter.Add("nChkDate", chkDate.IsChecked == true ? (chkMtrLOTIDSrh.IsChecked == true && !txtMtrLOTIDSrh.Text.Trim().Equals("") ? 1 : 0) : 0);
                sqlParameter.Add("sSDate", chkDate.IsChecked == true ? (dtpSDate.SelectedDate != null ? dtpSDate.SelectedDate.Value.ToString("yyyyMMdd") : "") : "");
                sqlParameter.Add("SEDate", chkDate.IsChecked == true ? (dtpEDate.SelectedDate != null ? dtpEDate.SelectedDate.Value.ToString("yyyyMMdd") : "") : "");
                sqlParameter.Add("nChkArticleID", chkMtrArticleSrh.IsChecked == true ? 1 : 0); //자재품명 체크
                sqlParameter.Add("sArticleID", chkMtrArticleSrh.IsChecked == true ? (txtMtrArticleSrh.Tag == null ? "" : txtMtrArticleSrh.Tag.ToString()) : ""); //자재품명
                sqlParameter.Add("nChkParentArticleID", chkProdArticleSrh.IsChecked == true ? 1 : 0); //제품품명 체크
                sqlParameter.Add("sParentArticleID", chkProdArticleSrh.IsChecked == true ? (txtProdArticleSrh.Tag == null ? "" : txtProdArticleSrh.Tag.ToString()) : ""); //제품품명
                sqlParameter.Add("nChkCustom", chkCustomSrh.IsChecked == true ? 1 : 0); //거래처 체크
                sqlParameter.Add("sCustomID", chkCustomSrh.IsChecked == true ? (txtCustomSrh.Tag == null ? "" : txtCustomSrh.Tag.ToString()) : ""); //거래처
                sqlParameter.Add("incNotApprovalYN", chkIncNotApproval.IsChecked == true ? "Y" : ""); //입고미승인건
                sqlParameter.Add("incAutoInOutYN", chkIncAutoInOutWare.IsChecked == true ? "Y" : ""); //자동입출고건
                sqlParameter.Add("nMainItem", chkMainItem.IsChecked == true ? 1 : 0); //주요관심품목
                sqlParameter.Add("nCustomItem", chkCustomItem.IsChecked == true ? 1 : 0); //거래처별 등록 품목
                sqlParameter.Add("nIncZeroQty", chknIncZeroQty.IsChecked == true ? 1 : 0); //재고 0 포함, 기본값은 재고 0은 포함 안함
                sqlParameter.Add("nChkLotID", chkMtrLOTIDSrh.IsChecked == true ? 1 : 0); //Lot 췍크 
                sqlParameter.Add("sLotID", chkMtrLOTIDSrh.IsChecked == true ? (txtMtrLOTIDSrh.Text == null ? "" : txtMtrLOTIDSrh.Text.ToString()) : ""); //lot 입력
                sqlParameter.Add("sFromLocID", chkWareHouse.IsChecked == true && cboWareHouse.SelectedValue != null ? cboWareHouse.SelectedValue.ToString() : ""); //창고 추가 2021-07-09
                sqlParameter.Add("sToLocID", "");
                //sqlParameter.Add("sLotID", chkMtrLOTIDSrh.IsChecked == true && !txtMtrLOTIDSrh.Text.Trim().Equals("") ? txtMtrLOTIDSrh.Text.Trim() : ""); //Lot 췌크 텍스트 
                //2021-09-23
                DataSet ds = DataStore.Instance.ProcedureToDataSet("xp_Subul_sMtrSubul_Lot_One", sqlParameter, true);

                if (ds != null && ds.Tables.Count > 0)
                {
                    DataTable dt = ds.Tables[0];

                    #region 닫겨라
                    //if (dt.Rows.Count > 0)
                    //{
                    //    int i = 0;

                    //    int bIndex = 0;
                    //    int mIndex = 0;
                    //    int sIndex = 0;

                    //    DataRowCollection drc = dt.Rows;

                    //    var LotStock = new LotStock();

                    //    // LotID만 넣을 객체를 만들어야 될거 같소
                    //    var LotID = new LotStock();
                    //    LotID.bIndex = bIndex;
                    //    LotID.sIndex = sIndex;

                    //    var Article = new LotStock();
                    //    Article.mIndex = mIndex;
                    //    Article.sIndex = sIndex;

                    //    foreach (DataRow dr in drc)
                    //    {
                    //        i++;

                    //        // cls - 0 : 이월, 1 : 입고, 2 : 출고, 3 : Lot집계, 4 : 품명 집계
                    //        if (dr["cls"].ToString().Trim().Equals("0"))
                    //        {
                    //            stack++;

                    //            // 랏
                    //            LotID.Add(setLotStock_Content(dr, i));
                    //            LotID.LotID = dr["LotID"].ToString();

                    //            // 품명
                    //            Article.Article = dr["Article"].ToString();
                    //            Article.BuyerArticleNo = dr["BuyerArticleNo"].ToString();
                    //            Article.ArticleID = dr["ArticleID"].ToString();
                    //            Article.Add(setLotStock_Content(dr, i));

                    //            // 데이터 넣기
                    //            var LotStockData = setLotStock_Content(dr, i);
                    //            LotStockData.bIndex = bIndex;
                    //            LotStockData.mIndex = mIndex;

                    //            dgdNum.Items.Add(LotStockData);
                    //            dgdContent.Items.Add(LotStockData);

                    //            sIndex++;
                    //        }
                    //        else if (dr["cls"].ToString().Trim().Equals("1")) // 입고
                    //        {
                    //            stack++;

                    //            // 랏
                    //            LotID.Add(setLotStock_Content(dr, i));
                    //            LotID.LotID = dr["LotID"].ToString();

                    //            // 품명
                    //            Article.Add(setLotStock_Content(dr, i));

                    //            // 데이터 넣기
                    //            var LotStockData = setLotStock_Content(dr, i);
                    //            LotStockData.bIndex = bIndex;
                    //            LotStockData.mIndex = mIndex;

                    //            dgdNum.Items.Add(LotStockData);
                    //            dgdContent.Items.Add(LotStockData);

                    //            sIndex++;
                    //        }
                    //        else if (dr["cls"].ToString().Trim().Equals("2"))
                    //        {
                    //            stack++;

                    //            // 랏
                    //            LotID.Add(setLotStock_Content(dr, i));
                    //            LotID.LotID = dr["LotID"].ToString();

                    //            // 품명
                    //            Article.Article = dr["Article"].ToString();
                    //            Article.BuyerArticleNo = dr["BuyerArticleNo"].ToString();
                    //            Article.ArticleID = dr["ArticleID"].ToString();
                    //            Article.Add(setLotStock_Content(dr, i));

                    //            // 데이터 넣기
                    //            var LotStockData = setLotStock_Content(dr, i);
                    //            LotStockData.bIndex = bIndex;
                    //            LotStockData.mIndex = mIndex;

                    //            dgdNum.Items.Add(LotStockData);
                    //            dgdContent.Items.Add(LotStockData);

                    //            sIndex++;
                    //        }
                    //        else if (dr["cls"].ToString().Trim().Equals("3")) // Lot 집계 : 이때 LotID 를 넣자
                    //        {
                    //            stack++;

                    //            // 랏
                    //            LotID.Add(setLotStock_Content(dr, i));
                    //            LotID.LotID = dr["LotID"].ToString();

                    //            LotID.RowHeight = rowHeight * LotID.Count;
                    //            dgdLotID.Items.Add(LotID);

                    //            // 품명
                    //            Article.Add(setLotStock_Content(dr, i));

                    //            // 데이터 넣기
                    //            var LotStockData = setLotStock_Content(dr, i);
                    //            LotStockData.bIndex = bIndex;
                    //            LotStockData.mIndex = mIndex;

                    //            dgdNum.Items.Add(LotStockData);
                    //            dgdContent.Items.Add(LotStockData);

                    //            sIndex++;

                    //            // Lot 초기화 및 인덱스 넣기
                    //            LotID = new LotStock();
                    //            LotID.bIndex = bIndex;
                    //            LotID.sIndex = sIndex;
                    //            mIndex++;
                    //        }
                    //        else if (dr["cls"].ToString().Trim().Equals("4")) // 품명집계 : 품명이 끝났다는 거겠지... 이때 품명을 넣자
                    //        {
                    //            if (stack == 0)
                    //            {
                    //                var tempArticle = dgdArticle.Items[dgdArticle.Items.Count - 1] as LotStock;
                    //                tempArticle.Add(setLotStock_Content(dr, i));
                    //                tempArticle.RowHeight = tempArticle.Count * rowHeight;
                    //            }
                    //            else
                    //            {
                    //                // 품명계 > 를 넣어줄때 Lot도 빈칸을 넣어줘야 함.
                    //                Article.Article = dr["Article"].ToString();
                    //                Article.BuyerArticleNo = dr["BuyerArticleNo"].ToString();
                    //                Article.ArticleID = dr["ArticleID"].ToString();
                    //                Article.Add(setLotStock_Content(dr, i));
                    //                Article.RowHeight = Article.Count * rowHeight;
                    //                Article.mEndIndex = mIndex;
                    //                Article.sEndIndex = sIndex;
                    //                dgdArticle.Items.Add(Article);
                    //            }

                    //            // 데이터 넣기
                    //            var LotStockData = setLotStock_Content(dr, i);
                    //            LotStockData.bIndex = stack == 0 ? bIndex - 1 : bIndex;
                    //            LotStockData.mIndex = mIndex;

                    //            dgdNum.Items.Add(LotStockData);
                    //            dgdContent.Items.Add(LotStockData);

                    //            // LotID
                    //            var emptyLotID = new LotStock();
                    //            emptyLotID.Add(setLotStock_Content(dr, i));
                    //            emptyLotID.RowHeight = rowHeight * emptyLotID.Count;
                    //            emptyLotID.bIndex = stack == 0 ? bIndex - 1 : bIndex;
                    //            emptyLotID.sIndex = sIndex;
                    //            emptyLotID.LotID = "품명계";
                    //            dgdLotID.Items.Add(emptyLotID);
                    //            mIndex++;

                    //            sIndex++;

                    //            Article = new LotStock();
                    //            Article.mIndex = mIndex;
                    //            Article.sIndex = sIndex;

                    //            if (stack != 0)
                    //                bIndex++;

                    //            // Lot 초기화 및 인덱스 넣기
                    //            LotID = new LotStock();
                    //            LotID.bIndex = bIndex;
                    //            LotID.sIndex = sIndex;

                    //            // stack 초기회
                    //            stack = 0;
                    //        }
                    //        else if (dr["cls"].ToString().Trim().Equals("7")) // 총계 맨 마지막에!
                    //        {
                    //            // 총계
                    //            var total = new LotStock();
                    //            total.BuyerArticleNo = "총계"; //
                    //            total.Add(setLotStock_Content(dr, i));
                    //            total.RowHeight = rowHeight * total.Count;
                    //            total.bIndex = bIndex;
                    //            total.mIndex = mIndex;
                    //            total.sIndex = sIndex;
                    //            dgdArticle.Items.Add(total);
                    //            dgdLotID.Items.Add(total);

                    //            // 데이터 넣기
                    //            var LotStockData = setLotStock_Content(dr, i);
                    //            LotStockData.bIndex = bIndex;
                    //            LotStockData.mIndex = mIndex;

                    //            dgdNum.Items.Add(LotStockData);
                    //            dgdContent.Items.Add(LotStockData);

                    //        }

                    //    }
                    //}
                    #endregion //닫겨라

                    #region 검색 1500 이하
                    //if (dt.Rows.Count > 0 && dt.Rows.Count <= 1500)
                    //{
                    //    dgdNum.Visibility = Visibility.Visible;
                    //    dgdArticle.Visibility = Visibility.Visible;
                    //    dgdLotID.Visibility = Visibility.Visible;
                    //    dgdContent.Visibility = Visibility.Visible;

                    //    dgdMain.Visibility = Visibility.Hidden;

                    //    int i = 0;

                    //    int bIndex = 0;
                    //    int mIndex = 0;
                    //    int sIndex = 0;

                    //    DataRowCollection drc = dt.Rows;

                    //    var LotStock = new LotStock();

                    //    // LotID만 넣을 객체를 만들어야 될거 같소
                    //    var LotID = new LotStock();
                    //    LotID.bIndex = bIndex;
                    //    LotID.sIndex = sIndex;

                    //    var Article = new LotStock();
                    //    Article.mIndex = mIndex;
                    //    Article.sIndex = sIndex;

                    //    foreach (DataRow dr in drc)
                    //    {
                    //        //1번이 총계라서 num +1 해야하므로 실행 후 i증가 2021-10-28

                    //        // cls - 0 : 이월, 1 : 입고, 2 : 출고, 3 : Lot별 일자집계, 4 : Lot별 집계 5 : 품명 집계
                    //        if (dr["cls"].ToString().Trim().Equals("0"))
                    //        {
                    //            stack++;

                    //            // 랏
                    //            LotID.Add(setLotStock_Content(dr, i));
                    //            LotID.LotID = dr["LotID"].ToString();

                    //            // 품명
                    //            Article.BuyerArticleNo = dr["BuyerArticleNo"].ToString();
                    //            Article.Article = dr["Article"].ToString();
                    //            Article.ArticleID = dr["ArticleID"].ToString();
                    //            Article.Add(setLotStock_Content(dr, i));

                    //            // 데이터 넣기
                    //            var LotStockData = setLotStock_Content(dr, i);
                    //            LotStockData.bIndex = bIndex;
                    //            LotStockData.mIndex = mIndex;

                    //            dgdNum.Items.Add(LotStockData);
                    //            dgdContent.Items.Add(LotStockData);

                    //            sIndex++;
                    //        }
                    //        else if (dr["cls"].ToString().Trim().Equals("1")) // 입고
                    //        {
                    //            stack++;

                    //            // 랏
                    //            LotID.Add(setLotStock_Content(dr, i));
                    //            LotID.LotID = dr["LotID"].ToString();

                    //            // 품명
                    //            Article.Add(setLotStock_Content(dr, i));

                    //            // 데이터 넣기
                    //            var LotStockData = setLotStock_Content(dr, i);
                    //            LotStockData.bIndex = bIndex;
                    //            LotStockData.mIndex = mIndex;

                    //            dgdNum.Items.Add(LotStockData);
                    //            dgdContent.Items.Add(LotStockData);

                    //            sIndex++;
                    //        }
                    //        else if (dr["cls"].ToString().Trim().Equals("2")) //출고
                    //        {
                    //            stack++;

                    //            // 랏
                    //            LotID.Add(setLotStock_Content(dr, i));
                    //            LotID.LotID = dr["LotID"].ToString();

                    //            // 품명
                    //            Article.BuyerArticleNo = dr["BuyerArticleNo"].ToString();
                    //            Article.Article = dr["Article"].ToString();
                    //            Article.ArticleID = dr["ArticleID"].ToString();
                    //            Article.Add(setLotStock_Content(dr, i));

                    //            // 데이터 넣기
                    //            var LotStockData = setLotStock_Content(dr, i);
                    //            LotStockData.bIndex = bIndex;
                    //            LotStockData.mIndex = mIndex;

                    //            dgdNum.Items.Add(LotStockData);
                    //            dgdContent.Items.Add(LotStockData);

                    //            sIndex++;
                    //        }
                    //        //else if (dr["cls"].ToString().Trim().Equals("3")) // Lot 집계 일자조회
                    //        //{
                    //        //    stack++;

                    //        //    // 랏
                    //        //    LotID.Add(setLotStock_Content(dr, i));
                    //        //    LotID.LotID = dr["LotID"].ToString();

                    //        //    LotID.RowHeight = rowHeight * LotID.Count;
                    //        //    dgdLotID.Items.Add(LotID);

                    //        //    // 품명
                    //        //    Article.Add(setLotStock_Content(dr, i));

                    //        //    // 데이터 넣기
                    //        //    var LotStockData = setLotStock_Content(dr, i);
                    //        //    LotStockData.bIndex = bIndex;
                    //        //    LotStockData.mIndex = mIndex;

                    //        //    dgdNum.Items.Add(LotStockData);
                    //        //    dgdContent.Items.Add(LotStockData);

                    //        //    sIndex++;

                    //        //    // Lot 초기화 및 인덱스 넣기
                    //        //    LotID = new LotStock();
                    //        //    LotID.bIndex = bIndex;
                    //        //    LotID.sIndex = sIndex;
                    //        //    mIndex++;
                    //        //}

                    //        // cls - 0 : 이월, 1 : 입고, 2 : 출고, 3 : Lot별 일자집계, 4 : Lot별 집계 5 : 품명 집계
                    //        else if (dr["cls"].ToString().Trim().Equals("3")) // Lot 집계 : 이때 LotID 를 넣자
                    //        {
                    //            stack++;

                    //            // 랏
                    //            LotID.Add(setLotStock_Content(dr, i));
                    //            LotID.LotID = dr["LotID"].ToString();

                    //            LotID.RowHeight = rowHeight * LotID.Count;
                    //            dgdLotID.Items.Add(LotID);

                    //            // 품명
                    //            Article.Add(setLotStock_Content(dr, i)); 

                    //            // 데이터 넣기
                    //            var LotStockData = setLotStock_Content(dr, i);
                    //            LotStockData.bIndex = bIndex;
                    //            LotStockData.mIndex = mIndex;

                    //            dgdNum.Items.Add(LotStockData);
                    //            dgdContent.Items.Add(LotStockData);

                    //            sIndex++;

                    //            // Lot 초기화 및 인덱스 넣기
                    //            LotID = new LotStock();
                    //            LotID.bIndex = bIndex;
                    //            LotID.sIndex = sIndex;
                    //            mIndex++;
                    //        }
                    //        else if (dr["cls"].ToString().Trim().Equals("4")) // 품명집계 : 품명이 끝났다는 거겠지... 이때 품명을 넣자
                    //        {
                    //            if (stack == 0)
                    //            {
                    //                var tempArticle = dgdArticle.Items[dgdArticle.Items.Count - 1] as LotStock;
                    //                tempArticle.Add(setLotStock_Content(dr, i));
                    //                tempArticle.RowHeight = tempArticle.Count * rowHeight;
                    //            }
                    //            else
                    //            {
                    //                // 품명계 > 를 넣어줄때 Lot도 빈칸을 넣어줘야 함.
                    //                Article.BuyerArticleNo = dr["BuyerArticleNo"].ToString();
                    //                Article.Article = dr["Article"].ToString();
                    //                Article.ArticleID = dr["ArticleID"].ToString();
                    //                Article.Add(setLotStock_Content(dr, i));
                    //                Article.RowHeight = Article.Count * rowHeight;
                    //                Article.mEndIndex = mIndex;
                    //                Article.sEndIndex = sIndex;
                    //                dgdArticle.Items.Add(Article);
                    //            }

                    //            // 데이터 넣기
                    //            var LotStockData = setLotStock_Content(dr, i);
                    //            LotStockData.bIndex = stack == 0 ? bIndex - 1 : bIndex;
                    //            LotStockData.mIndex = mIndex;

                    //            dgdNum.Items.Add(LotStockData);
                    //            dgdContent.Items.Add(LotStockData);

                    //            // LotID
                    //            var emptyLotID = new LotStock();
                    //            emptyLotID.Add(setLotStock_Content(dr, i));
                    //            emptyLotID.RowHeight = rowHeight * emptyLotID.Count;
                    //            emptyLotID.bIndex = stack == 0 ? bIndex - 1 : bIndex;
                    //            emptyLotID.sIndex = sIndex;
                    //            if (LotStockData.ioDate.Trim().Equals(""))
                    //            {
                    //                emptyLotID.LotID = "품명계";
                    //            }
                    //            else
                    //            {
                    //                emptyLotID.LotID = "일자계";
                    //            }

                    //            dgdLotID.Items.Add(emptyLotID);
                    //            mIndex++;

                    //            sIndex++;

                    //            Article = new LotStock();
                    //            Article.mIndex = mIndex;
                    //            Article.sIndex = sIndex;

                    //            if (stack != 0)
                    //                bIndex++;

                    //            // Lot 초기화 및 인덱스 넣기
                    //            LotID = new LotStock();
                    //            LotID.bIndex = bIndex;
                    //            LotID.sIndex = sIndex;

                    //            // stack 초기회
                    //            stack = 0;
                    //        }
                    //        else if (dr["cls"].ToString().Trim().Equals("7")) // 총계 맨 마지막에!
                    //        {
                    //            // 총계
                    //            var total = new LotStock();
                    //            total.Article = "총계";
                    //            total.Add(setLotStock_Content(dr, i));
                    //            total.RowHeight = rowHeight * total.Count;
                    //            total.bIndex = bIndex;
                    //            total.mIndex = mIndex;
                    //            total.sIndex = sIndex;
                    //            dgdTotal.Items.Add(total);
                    //            //dgdArticle.Items.Add(total);
                    //            //dgdLotID.Items.Add(total);

                    //            // 데이터 넣기
                    //            //var LotStockData = setLotStock_Content(dr, i);
                    //            //LotStockData.bIndex = bIndex;
                    //            //LotStockData.mIndex = mIndex;

                    //            //dgdNum.Items.Add(LotStockData);
                    //            //dgdContent.Items.Add(LotStockData);


                    //        }
                    //        i++;
                    //        //tblCount.Text = "▶검색결과 : " + i + "건";
                    //    }
                    //}
                    #endregion

                    #region 검색 1500 초과
                    if (dt.Rows.Count > 0)
                    {
                        dgdNum.Visibility = Visibility.Hidden;
                        dgdArticle.Visibility = Visibility.Hidden;
                        dgdLotID.Visibility = Visibility.Hidden;
                        dgdContent.Visibility = Visibility.Hidden;

                        dgdMain.Visibility = Visibility.Visible;

                        DataRowCollection drc = dt.Rows;

                        int i = 0;

                        foreach (DataRow dr in drc)
                        {


                            //if (dr["cls"].ToString().Trim().Equals("3")) // Lot 일자별
                            //{
                            //    Lot = new LotStock_All()
                            //    {
                            //        Num = i,

                            //        BuyerArticleNo = dr["BuyerArticleNo"].ToString(),
                            //        Article = dr["Article"].ToString(),
                            //        LotID = dr["LotID"].ToString(),
                            //        ioDate = dr["ioDate"].ToString(),
                            //        Gubun = "Lot별 일 집계",
                            //        ioDate_CV = DatePickerFormat(dr["ioDate"].ToString()),
                            //        //ioDate_CV = "",
                            //        StuffQty = stringFormatN0(dr["StuffQty"]),
                            //        UnitClssName = dr["UnitClssName"].ToString(),

                            //        OutQty = stringFormatN0(dr["OutQty"]),
                            //        StockQty = stringFormatN0(dr["StockQty"]),
                            //        Remark = dr["Remark"].ToString(),

                            //        LotColor = true
                            //    };
                            //}
                            if (dr["cls"].ToString().Trim().Equals("3")) // Lot 총계
                            {
                                Lot = new LotStock_All()
                                {
                                    Num = i,

                                    BuyerArticleNo = dr["BuyerArticleNo"].ToString(),
                                    Article = dr["Article"].ToString(),
                                    LotID = dr["LotID"].ToString(),
                                    ioDate = dr["ioDate"].ToString(),
                                    Gubun = "Lot집계",
                                    ioDate_CV = "",
                                    StuffQty = stringFormatN0(dr["StuffQty"]),
                                    UnitClssName = dr["UnitClssName"].ToString(),

                                    OutQty = stringFormatN0(dr["OutQty"]),
                                    StockQty = stringFormatN0(dr["StockQty"]),
                                    Remark = dr["Remark"].ToString(),

                                    LotColor = true
                                };
                            }

                            else if (dr["cls"].ToString().Trim().Equals("4")) // 품명계
                            {
                                Lot = new LotStock_All()
                                {
                                    Num = i,
                                    BuyerArticleNo = dr["BuyerArticleNo"].ToString(),
                                    Article = dr["Article"].ToString(),
                                    LotID = "품명계",
                                    ioDate = dr["ioDate"].ToString(),
                                    Gubun = "",
                                    ioDate_CV = "",
                                    StuffQty = stringFormatN0(dr["StuffQty"]),
                                    UnitClssName = dr["UnitClssName"].ToString(),

                                    OutQty = stringFormatN0(dr["OutQty"]),
                                    StockQty = stringFormatN0(dr["StockQty"]),
                                    Remark = dr["Remark"].ToString(),

                                    ArticleColor = true
                                };
                            }
                            else if (dr["cls"].ToString().Trim().Equals("7")) // 총계
                            {
                                Lot = new LotStock_All()
                                {
                                    Num = i,
                                    Article = "총계",
                                    LotID = "",
                                    //ioDate = dr["ioDate"].ToString(),
                                    Gubun = "",
                                    ioDate_CV = "",
                                    StuffQty = stringFormatN0(dr["StuffQty"]),
                                    //UnitClssName = dr["UnitClssName"].ToString(),

                                    OutQty = stringFormatN0(dr["OutQty"]),
                                    StockQty = stringFormatN0(dr["StockQty"]),
                                    //Remark = dr["Remark"].ToString(),

                                    TotalColor = true
                                };

                            }
                            else
                            {
                                string gubun = "";

                                if (dr["cls"].ToString().Trim().Equals("0"))
                                {
                                    gubun = "이월";
                                }
                                else if (dr["cls"].ToString().Trim().Equals("1"))
                                {
                                    gubun = "입고";
                                }
                                else if (dr["cls"].ToString().Trim().Equals("2"))
                                {
                                    gubun = "출고";
                                }

                                Lot = new LotStock_All()
                                {
                                    Num = i,

                                    BuyerArticleNo = dr["BuyerArticleNo"].ToString(),

                                    Article = dr["Article"].ToString(),
                                    LotID = dr["LotID"].ToString(),
                                    Gubun = gubun,
                                    ioDate = dr["ioDate"].ToString(),
                                    ioDate_CV = DatePickerFormat(dr["ioDate"].ToString()),
                                    StuffQty = stringFormatN0(dr["StuffQty"]),
                                    UnitClssName = dr["UnitClssName"].ToString(),

                                    OutQty = stringFormatN0(dr["OutQty"]),
                                    StockQty = stringFormatN0(dr["StockQty"]),
                                    Remark = dr["Remark"].ToString()
                                };
                            }
                            i++;
                            //tblCount.Text = "▶검색결과 : " + i + "건";
                            //총계는 밑으로 뺌
                            if (dr["cls"].ToString().Trim().Equals("7"))
                            {
                                dgdTotal.Items.Add(Lot);
                            }
                            else
                            {
                                dgdMain.Items.Add(Lot);
                            }
                        }
                    }
                    #endregion
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

        // 객체 값 세팅
        private LotStock_Content setLotStock_Content(DataRow dr, int num)
        {
            var LotStock = new LotStock_Content()
            {
                Num = num,

                cls = dr["cls"].ToString(),
                ArticleID = dr["ArticleID"].ToString(),
                Article = dr["Article"].ToString(),
                LotID = dr["LotID"].ToString().Trim(),
                ioDate = DatePickerFormat(dr["ioDate"].ToString()),

                StuffRoll = dr["StuffRoll"].ToString(),
                StuffQty = stringFormatN0(dr["StuffQty"]),
                UnitClss = dr["UnitClss"].ToString(),
                UnitClssName = dr["UnitClssName"].ToString(),
                OutRoll = dr["OutRoll"].ToString(),

                OutQty = stringFormatN0(dr["OutQty"]),
                StockQty = stringFormatN0(dr["StockQty"]),
                Remark = dr["Remark"].ToString(),

                BuyerArticleNo = dr["BuyerArticleNo"].ToString(),
            };

            if (LotStock.cls.Trim().Equals("0")) // 이월
            {
                LotStock.Gubun = "이월";
            }
            else if (LotStock.cls.Trim().Equals("1")) // 입고 
            {
                LotStock.Gubun = "입고";
            }
            else if (LotStock.cls.Trim().Equals("2")) // 출고
            {
                LotStock.Gubun = "출고";
            }
            //else if (LotStock.cls.Trim().Equals("3")) // Lot집계 > ioDate → 빈칸
            //{
            //    LotStock.Gubun = "Lot별 일집계";
            //    LotStock.ioDate = "";
            //}

            else if (LotStock.cls.Trim().Equals("3")) // Lot집계 > ioDate → 빈칸
            {
                LotStock.Gubun = "Lot집계";
                LotStock.ioDate = "";
            }
            else if (LotStock.cls.Trim().Equals("4")) // 품명계 > LotID → 품명계, ioDate → 빈칸
            {
                if (LotStock.ioDate.Trim().Equals("9999-99-99")) // 품명계 이면
                {
                    LotStock.Gubun = "";
                    LotStock.LotID = "품명계";
                    LotStock.ioDate = "";
                } // 일자계 이면
                else
                {
                    LotStock.Gubun = "";
                    LotStock.LotID = "일자계";
                }


            }
            else if (LotStock.cls.Trim().Equals("7")) // 총계 > Article → 총계 / Lot, ioDate → 빈칸
            {
                LotStock.Gubun = "";
                LotStock.BuyerArticleNo = "총계";
                LotStock.LotID = "";
                LotStock.ioDate = "";
            }

            return LotStock;
        }

        #endregion



        #region 기타 메서드

        #region Scrollbar 4개의 데이터그리드 동기화

        // 1. 스크롤 전역변수 선언 후에
        //ScrollViewer scrollView = null;
        //ScrollViewer scrollView2 = null;

        // 2. UserLoaded 함수에 넣어주면 끝
        //scrollView = getScrollbar(dgdNum);
        //scrollView2 = getScrollbar(dgdArticle);

        //if (null != scrollView)
        //{
        //    scrollView.ScrollChanged += new ScrollChangedEventHandler(scrollView_ScrollChanged);
        //}
        //if (null != scrollView2)
        //{
        //    scrollView2.ScrollChanged += new ScrollChangedEventHandler(scrollView_ScrollChanged);
        //}

        void scrollView_ScrollChanged(object sender, ScrollChangedEventArgs e)
        {
            var newOffset = e.VerticalOffset;

            if ((null != scrollView) && (null != scrollView2) && (null != scrollView3) && (null != scrollView4))
            {
                scrollView.ScrollToVerticalOffset(newOffset);
                scrollView2.ScrollToVerticalOffset(newOffset);
                scrollView3.ScrollToVerticalOffset(newOffset);
                scrollView4.ScrollToVerticalOffset(newOffset);
            }
        }

        private ScrollViewer getScrollbar(DependencyObject dep)
        {
            for (int i = 0; i < VisualTreeHelper.GetChildrenCount(dep); i++)
            {
                var child = VisualTreeHelper.GetChild(dep, i);
                if ((null != child) && child is ScrollViewer)
                {
                    return (ScrollViewer)child;
                }
                else
                {
                    ScrollViewer sub = getScrollbar(child);
                    if (sub != null)
                    {
                        return sub;
                    }
                }
            }
            return null;
        }

        #endregion // Scrollbar 4개의 데이터그리드 동기화

        // 천마리 콤마, 소수점 버리기
        private string stringFormatN0(object obj)
        {
            return string.Format("{0:N0}", obj);
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




        #endregion // 기타 메서드


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
                MessageBox.Show("오류 발생, 오류 내용 : " + ex.ToString());
            }
        }

        private T GetParent<T>(DependencyObject d) where T : class
        {
            while (d != null && !(d is T))
            {
                d = VisualTreeHelper.GetParent(d);
            }
            return d as T;
        }


        private void DataGird_KeyDown(object sender, KeyEventArgs e)
        {
            try
            {
                DataGrid dgd = GetParent<DataGrid>(e.OriginalSource as DependencyObject);
                //DataGridRow row = GetParent<DataGridRow>(e.OriginalSource as DependencyObject);

                int currRow = dgd.Items.IndexOf(dgd.CurrentItem);
                int currCol = dgd.Columns.IndexOf(dgd.CurrentCell.Column);

                if (e.Key == Key.Down)
                {
                    e.Handled = true;

                    //마지막 행 아님
                    if (dgd.Items.Count - 1 > currRow)
                    {
                        dgd.SelectedIndex = currRow + 1;
                        dgd.CurrentCell = new DataGridCellInfo(dgd.Items[currRow + 1], dgd.Columns[currCol]);
                    }  //마지막 행일때
                }
                else if (e.Key == Key.Up)
                {
                    e.Handled = true;

                    //첫행 아님
                    if (currRow > 0)
                    {
                        dgd.SelectedIndex = currRow - 1;
                        dgd.CurrentCell = new DataGridCellInfo(dgd.Items[currRow - 1], dgd.Columns[currCol]);
                    }
                }
                else if (e.Key == Key.Left)
                {
                    e.Handled = true;

                    if (dgd.Name.Trim().Equals("dgdNum"))
                    {
                        return;
                    }
                    else if (dgd.Name.Trim().Equals("dgdArticle")) // 순번으로 이동
                    {
                        //var LotStock = dgd.SelectedItem as LotStock;

                        //if (LotStock != null)
                        //{
                        //    int sIndex = LotStock.sIndex;

                        dgdNum.Focus();
                        dgdNum.CurrentCell = dgdNum.SelectedCells[0];
                        //= new DataGridCellInfo(dgdNum.Items[sIndex], dgdNum.Columns[0]);
                        //}
                    }
                    else if (dgd.Name.Trim().Equals("dgdLotID")) // 품명으로 이동
                    {
                        //var LotStock = dgd.SelectedItem as LotStock;

                        //if (LotStock != null)
                        //{
                        //    int bIndex = LotStock.bIndex;

                        dgdArticle.Focus();
                        dgdArticle.CurrentCell = dgdArticle.SelectedCells[0];
                        //= new DataGridCellInfo(dgdArticle.Items[bIndex], dgdArticle.Columns[0]);
                        //}
                    }
                    else if (dgd.Name.Trim().Equals("dgdContent")) // 랏으로 이동
                    {
                        int startCol = 0;

                        if (startCol < currCol)
                        {
                            dgd.CurrentCell = new DataGridCellInfo(dgd.Items[currRow], dgd.Columns[currCol - 1]);
                        }
                        else if (currCol == startCol)
                        {
                            //var LotStock = dgd.SelectedItem as LotStock_Content;

                            //if (LotStock != null)
                            //{
                            //    int mIndex = LotStock.mIndex;

                            dgdLotID.Focus();
                            dgdLotID.CurrentCell = dgdLotID.SelectedCells[0];
                            //= new DataGridCellInfo(dgdLotID.Items[mIndex], dgdLotID.Columns[0]);
                            //}
                        }
                    }
                }
                else if (e.Key == Key.Right)
                {
                    e.Handled = true;

                    if (dgd.Name.Trim().Equals("dgdNum")) // 품명으로 이동
                    {
                        //var LotStock = dgd.SelectedItem as LotStock_Content;

                        //if (LotStock != null)
                        //{
                        //    int bIndex = LotStock.bIndex;

                        dgdArticle.Focus();
                        dgdArticle.CurrentCell = dgdArticle.SelectedCells[0];
                        //= new DataGridCellInfo(dgdArticle.Items[bIndex], dgdArticle.Columns[0]);
                        //}
                    }
                    else if (dgd.Name.Trim().Equals("dgdArticle")) // 랏으로 이동
                    {
                        //var LotStock = dgd.SelectedItem as LotStock;

                        //if (LotStock != null)
                        //{
                        //    int mIndex = LotStock.mIndex;

                        dgdLotID.Focus();
                        dgdLotID.CurrentCell = dgdLotID.SelectedCells[0];
                        // = new DataGridCellInfo(dgdLotID.Items[mIndex], dgdLotID.Columns[0]);
                        //}
                    }
                    else if (dgd.Name.Trim().Equals("dgdLotID")) // 데이터로 이동
                    {
                        //var LotStock = dgd.SelectedItem as LotStock;

                        //if (LotStock != null)
                        //{
                        //    int sIndex = LotStock.sIndex;

                        dgdContent.Focus();
                        dgdContent.CurrentCell = dgdContent.SelectedCells[0];
                        // = new DataGridCellInfo(dgdContent.Items[sIndex], dgdContent.Columns[0]);
                        //}
                    }
                    else if (dgd.Name.Trim().Equals("dgdContent"))
                    {
                        int endCol = 6;

                        if (endCol > currCol)
                        {
                            dgd.CurrentCell = new DataGridCellInfo(dgd.Items[currRow], dgd.Columns[currCol + 1]);
                        }
                    }

                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }



        }

        private void txtMtrLOTIDSrh_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                rowNum = 0;
                re_Search(rowNum);

            }
        }

        //lot 췌크
        private void chkMtrLOTIDSrh_Checked(object sender, RoutedEventArgs e)
        {
            chkMtrLOTIDSrh.IsChecked = true;
            txtMtrLOTIDSrh.IsEnabled = true;
        }


        //lot 노췌크
        private void chkMtrLOTIDSrh_UnChecked(object sender, RoutedEventArgs e)
        {
            chkMtrLOTIDSrh.IsChecked = false;
            txtMtrLOTIDSrh.IsEnabled = false;
        }

        private void lblSrhMtrLOTID_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            if (chkMtrLOTIDSrh.IsChecked == true)
            {
                chkMtrLOTIDSrh.IsChecked = false;
            }
            else
            {
                chkMtrLOTIDSrh.IsChecked = true;
            }
        }

        //2021-07-09
        private void chkWareHouse_Click(object sender, RoutedEventArgs e)
        {
            if (chkWareHouse.IsChecked == true)
            {
                cboWareHouse.IsEnabled = true;
                cboWareHouse.Focus();
            }
            else
            {
                cboWareHouse.IsEnabled = false;
            }

        }
        // 창고
        private void chkWareHouse_Click(object sender, MouseButtonEventArgs e)
        {
            if (chkWareHouse.IsChecked == true)
            {
                chkWareHouse.IsChecked = false;
                cboWareHouse.IsEnabled = false;
            }
            else
            {
                chkWareHouse.IsChecked = true;
                cboWareHouse.IsEnabled = true;
                cboWareHouse.Focus();
            }
        }


        private void ComboBoxSetting()
        {
            //cboArticleGroup.Items.Clear();
            cboWareHouse.Items.Clear();
            //cboInGbn.Items.Clear();
            //cboOutGbn.Items.Clear();
            //cboSupplyType.Items.Clear();

            ObservableCollection<CodeView> cbArticleGroup = ComboBoxUtil.Instance.Gf_DB_MT_sArticleGrp();
            ObservableCollection<CodeView> cbWareHouse = ComboBoxUtil.Instance.Gf_DB_CM_GetComCodeDataset(null, "LOC", "Y", "", "");
            ObservableCollection<CodeView> cbInGbn = ComboBoxUtil.Instance.Gf_DB_CM_GetComCodeDataset(null, "ICD", "Y", "", "");
            ObservableCollection<CodeView> cbOutGbn = ComboBoxUtil.Instance.Gf_DB_CM_GetComCodeDataset(null, "OCD", "Y", "", "");
            ObservableCollection<CodeView> cbSupplyType = ComboBoxUtil.Instance.Gf_DB_CM_GetComCodeDataset(null, "CMMASPLTYPE", "Y", "", "");

            //this.cboArticleGroup.ItemsSource = cbArticleGroup;
            //this.cboArticleGroup.DisplayMemberPath = "code_name";
            //this.cboArticleGroup.SelectedValuePath = "code_id";
            //this.cboArticleGroup.SelectedIndex = 0;

            this.cboWareHouse.ItemsSource = cbWareHouse;
            this.cboWareHouse.DisplayMemberPath = "code_name";
            this.cboWareHouse.SelectedValuePath = "code_id";
            this.cboWareHouse.SelectedIndex = 0;

            //this.cboInGbn.ItemsSource = cbInGbn;
            //this.cboInGbn.DisplayMemberPath = "code_id_plus_code_name";
            //this.cboInGbn.SelectedValuePath = "code_id";
            //this.cboInGbn.SelectedIndex = 0;

            //this.cboOutGbn.ItemsSource = cbOutGbn;
            //this.cboOutGbn.DisplayMemberPath = "code_id_plus_code_name";
            //this.cboOutGbn.SelectedValuePath = "code_id";
            //this.cboOutGbn.SelectedIndex = 0;

            //this.cboSupplyType.ItemsSource = cbSupplyType;
            //this.cboSupplyType.DisplayMemberPath = "code_name";
            //this.cboSupplyType.SelectedValuePath = "code_id";
            //this.cboSupplyType.SelectedIndex = 0;
        }
        //// Lot별 일자조회
        //private void lblnIncDay_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        //{
        //    if (chknIncDay.IsChecked == true)
        //    {
        //        chknIncDay.IsChecked = false;
        //    }
        //    else
        //    {
        //        chknIncDay.IsChecked = true;
        //    }
        //}
        //// Lot별 일자조회
        //private void chknIncDay_Checked(object sender, RoutedEventArgs e)
        //{
        //    chknIncDay.IsChecked = true;
        //}
        //// Lot별 일자조회
        //private void chknIncDay_UnChecked(object sender, RoutedEventArgs e)
        //{
        //    chknIncDay.IsChecked = false;
        //}
    }

    // cls - 1 : 이월, 2 : LOT집계, 4 : 품명별집계, 7 : 총집계
    // cls - 0 : 이월, 1 : 입고, 2 : 출고, 3 : Lot별 일자집계, 4 : Lot별 집계 5 : 품명 집계, 7 : 총집계

    class LotStock : List<LotStock_Content>
    {
        public int bIndex { get; set; }
        public int mIndex { get; set; }
        public int mEndIndex { get; set; }

        public int sIndex { get; set; }
        public int sEndIndex { get; set; }

        public string ArticleID { get; set; }
        public string Article { get; set; }

        public string LotID { get; set; }

        public string BuyerArticleNo { get; set; }

        public int RowHeight { get; set; }

    }
    class LotStock_All
    {
        public int Num { get; set; }

        public string Gubun { get; set; }

        public string cls { get; set; }
        public string ArticleID { get; set; }
        public string Article { get; set; }
        public string LotID { get; set; }
        public string ioDate { get; set; }
        public string ioDate_CV { get; set; }

        public string StuffRoll { get; set; }
        public string StuffQty { get; set; }
        public string UnitClss { get; set; }
        public string UnitClssName { get; set; }
        public string OutRoll { get; set; }

        public string OutQty { get; set; }
        public string StockQty { get; set; }
        public string BuyerArticleNo { get; set; }
        public string Remark { get; set; }

        public bool LotColor { get; set; }
        public bool ArticleColor { get; set; }
        public bool TotalColor { get; set; }
    }


    class LotStock_Collection : List<LotStock>
    {

    }

    class LotStock_Content
    {
        public int Num { get; set; }

        public string Gubun { get; set; }

        public int bIndex { get; set; }
        public int mIndex { get; set; }
        public int sIndex { get; set; }

        public string cls { get; set; }
        public string ArticleID { get; set; }
        public string Article { get; set; }
        public string LotID { get; set; }
        public string ioDate { get; set; }

        public string StuffRoll { get; set; }
        public string StuffQty { get; set; }
        public string UnitClss { get; set; }
        public string UnitClssName { get; set; }
        public string OutRoll { get; set; }

        public string OutQty { get; set; }
        public string StockQty { get; set; }
        public string BuyerArticleNo { get; set; }
        public string Remark { get; set; }

        public int RowHeight { get; set; }
    }
}

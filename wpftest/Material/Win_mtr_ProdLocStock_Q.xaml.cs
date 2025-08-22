using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Linq;
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
    /// Win_mtr_ProdLocStock_Q.xaml에 대한 상호 작용 논리
    /// </summary>
    public partial class Win_mtr_ProdLocStock_Q : UserControl
    {
        int rowNum = 0;
        string strFlag = string.Empty;

        ScrollViewer scrollView = null;
        ScrollViewer scrollView2 = null;
        ScrollViewer scrollView3 = null;
        ScrollViewer scrollView4 = null;
        Lib lib = new Lib();
        private Microsoft.Office.Interop.Excel.Application excel;
        private Microsoft.Office.Interop.Excel.Workbook workBook;
        private Microsoft.Office.Interop.Excel.Worksheet workSheet;
        private Microsoft.Office.Interop.Excel.Range cellRange;

        // 병합시킬 인덱스 모음 - 안의 값 : 첫번째 값은 컬럼번호, 나머진 행높이 
        List<List<int>> lstMerge = new List<List<int>>();
        int startIndex = 0;

        public Win_mtr_ProdLocStock_Q()
        {
            InitializeComponent();
        }

        // 폼 로드 됬을 때
        private void UserContol_Loaded(object sender, RoutedEventArgs e)
        {
            Lib.Instance.UiLoading(sender);

            // 이번달 세팅하기
            chkStdMonth.IsChecked = true;
            StdMonthDate.IsEnabled = true;
            StdMonthDate.SelectedDate = DateTime.Today;

            // 스크롤
            scrollView = getScrollbar(dgdProcess);
            scrollView2 = getScrollbar(dgdGubun);
            scrollView3 = getScrollbar(dgdCarryOver);
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

        #region Header 부분 - 검색 조건

        private void chkStdMonth_Checked(object sender, RoutedEventArgs e)
        {
            chkStdMonth.IsChecked = true;
            //StdMonthDate.IsEnabled = true;
        }

        private void chkStdMonth_Unchecked(object sender, RoutedEventArgs e)
        {
            chkStdMonth.IsChecked = false;
            //StdMonthDate.IsEnabled = false;
        }


        // 제품그룹 라벨 이벤트
        private void lblArticleGrp_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            if (chkArticleGrp.IsChecked == true)
            {
                chkArticleGrp.IsChecked = false;
            }
            else
            {
                chkArticleGrp.IsChecked = true;
            }
        }
        private void chkArticleGrp_Checked(object sender, RoutedEventArgs e)
        {
            chkArticleGrp.IsChecked = true;
            cboArticleGrp.IsEnabled = true;
        }

        private void chkArticleGrp_Unchecked(object sender, RoutedEventArgs e)
        {
            chkArticleGrp.IsChecked = false;
            cboArticleGrp.IsEnabled = false;
        }
        #endregion // Header 부분 - 검색 조건

        #region  상단 오른쪽 버튼

        private void btnSearch_Click(object sender, RoutedEventArgs e)
        {
            //검색버튼 비활성화
            btnSearch.IsEnabled = false;

            Dispatcher.BeginInvoke(new Action(() =>

            {
                Thread.Sleep(2000);

                //로직
                rowNum = 0;
                strFlag = "S";
                re_Search(rowNum);

            }), System.Windows.Threading.DispatcherPriority.Background);



            Dispatcher.BeginInvoke(new Action(() =>

            {
                btnSearch.IsEnabled = true;

            }), System.Windows.Threading.DispatcherPriority.Background);


        }

        private void btnClose_Click(object sender, RoutedEventArgs e)
        {
            Lib.Instance.ChildMenuClose(this.ToString());
        }

        private void btnExcel_Click(object sender, RoutedEventArgs e)
        {

            // 데이터 그리드 4개를 합쳐서 DataTable 에 넣어야됨.
            try
            {
                DataGrid dgd = new DataGrid();
                dgd = dgdProcess;

                DataTable dt = null;
                string Name = string.Empty;

                string[] lst = new string[2];
                lst[0] = "자재 입고 목록";
                lst[1] = dgd.Name;

                ExportExcelxaml ExpExc = new ExportExcelxaml(lst);

                ExpExc.ShowDialog();

                if (ExpExc.DialogResult.HasValue)
                {
                    if (ExpExc.choice.Equals(dgd.Name))
                    {
                        //if (ExpExc.Check.Equals("Y"))
                        //    dt = Lib.Instance.DataGridToDTinHidden(dgd);
                        //else
                        //    dt = Lib.Instance.DataGirdToDataTable(dgd);

                        dt = getDataTableIn4Dgd(dgdProcess, dgdGubun, dgdCarryOver, dgdContent);
                        dt.Columns.Add("");

                        Name = dgd.Name;

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
            List<int> dgd1Merge1 = new List<int>();
            List<int> dgd1Merge2 = new List<int>();
            List<int> dgd3Merge = new List<int>();

            dgd1Merge1.Add(0);
            dgd1Merge2.Add(1);
            dgd3Merge.Add(3);

            // 1, 3 의 행 높이가 모두 같음
            for (int i = 0; i < dgd1.Items.Count; i++)
            {
                var ProcessMerge = dgd1.Items[i] as Win_mtr_ProdLocStock_Q_CodeView;
                dgd1Merge1.Add(ProcessMerge.Count);
                dgd1Merge2.Add(ProcessMerge.Count);
                dgd3Merge.Add(ProcessMerge.Count);
            }

            lstMerge.Add(dgd1Merge1);
            lstMerge.Add(dgd1Merge2);
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

            //MessageBox.Show("" + dt.Columns.Count);

            // 첫번째 그리드 변수
            int indexDgd1 = 0;
            int stackDgd1 = 0;
            int EndIndex_Dgd1 = 0;

            // 세번째 그리드 변수
            int indexDgd3 = 0;
            int stackDgd3 = 0;
            int EndIndex_Dgd3 = 0;

            // 그리드 객체들
            var Process = new Win_mtr_ProdLocStock_Q_CodeView();
            var Gubun = new ProdLocStock_Content();
            var CarryOver = new Win_mtr_ProdLocStock_Q_CodeView();
            var Content = new ProdLocStock_Content();

            // 실제 데이터가 1행씩 되있어, 반복은 데이터 그리드로!!!!
            for (int i = 0; i < dgd4.Items.Count; i++)
            {
                // 이제 데이터 세팅 → 셀이 병합되 있는 경우에는??
                DataRow dr = dt.NewRow();

                // 첫번째 그리드 : 병합
                if (stackDgd1 == 0)
                {
                    Process = dgd1.Items[indexDgd1] as Win_mtr_ProdLocStock_Q_CodeView;
                    EndIndex_Dgd1 = Process.Count - 1;

                    indexDgd1++;
                }

                // 두번째 그리드 : 병합 없음
                Gubun = dgd2.Items[i] as ProdLocStock_Content;

                // 세번째 그리드 : 병합
                if (stackDgd3 == 0)
                {
                    CarryOver = dgd3.Items[indexDgd3] as Win_mtr_ProdLocStock_Q_CodeView;
                    EndIndex_Dgd3 = CarryOver.Count - 1;

                    indexDgd3++;
                }

                // 네번째 그리드 : 병합 없음
                Content = dgd4.Items[i] as ProdLocStock_Content;


                ////////////////////
                // 데이터 넣어주기
                ////////////////////
                startIndex = 0;
                dr = setDataRow(Process, listVisibleDataGridColumns, dr, dgd1.Columns.Count);
                dr = setDataRow(Gubun, listVisibleDataGridColumns, dr, dgd2.Columns.Count);
                dr = setDataRow(CarryOver, listVisibleDataGridColumns, dr, dgd3.Columns.Count);
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
        #endregion // 상단 오른쪽 버튼 

        #region 주요 메서드

        //재조회
        private void re_Search(int selectedIndex)
        {
            FillGrid();

            if (dgdProcess.Items.Count > 0)
            {
                dgdProcess.SelectedIndex = selectedIndex;
                dgdProcess.Focus();
                dgdProcess.CurrentCell = dgdProcess.SelectedCells[0];
            }
            else
            {
                this.DataContext = null;
                MessageBox.Show("조회된 데이터가 없습니다.");
                return;
            }
        }

        //조회
        private void FillGrid()
        {
            if (dgdProcess.Items.Count > 0)
            {
                //dgdProcess.ItemsSource = null;
                dgdProcess.Items.Clear();
                dgdGubun.ItemsSource = null;
                //dgdCarryOver.ItemsSource = null;
                dgdCarryOver.Items.Clear();
                dgdContent.ItemsSource = null;
            }


            // 월 입력 해야됨!!
            if (chkStdMonth.IsChecked == true && StdMonthDate.SelectedDate == null)
            {
                MessageBox.Show("기준 월을 선택해주세요.");
                return;
            }

            // 제품군 체크박스가 체크 되있으면 선택할 수 있도록
            if (chkArticleGrp.IsChecked == true
                && (cboArticleGrp.SelectedValue == null || cboArticleGrp.SelectedValue.ToString().Trim().Equals("")))
            {
                MessageBox.Show("제품군을 선택해주세요.");
                return;
            }

            var AllLockStock = new ProdLocStock_Collection();

            try
            {
                Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                sqlParameter.Clear();

                sqlParameter.Add("sYYYYMM", chkStdMonth.IsChecked == true ? (StdMonthDate.SelectedDate != null ? StdMonthDate.SelectedDate.Value.ToString("yyyyMM") : "") : DateTime.Today.ToString("yyyyMM"));
                sqlParameter.Add("nProductGrpID", chkArticleGrp.IsChecked == true ? 1 : 0);
                sqlParameter.Add("sProductGrpID", chkArticleGrp.IsChecked == true ?
                    (cboArticleGrp.SelectedValue == null || cboArticleGrp.SelectedValue.ToString().Trim().Equals("") ? "" : cboArticleGrp.SelectedValue.ToString()) : "");
                DataSet ds = DataStore.Instance.ProcedureToDataSet("xp_Prd_sProcessStock", sqlParameter, false);

                if (ds != null && ds.Tables.Count > 0)
                {
                    DataTable dt = ds.Tables[0];

                    if (dt.Rows.Count > 0)
                    {
                        int i = 0;
                        int sIndex = 0;
                        DataRowCollection drc = dt.Rows;

                        foreach (DataRow dr in drc)
                        {
                            i++;

                            var LockStock = new Win_mtr_ProdLocStock_Q_CodeView(i, dr["StepName"].ToString(), dr["initQty"].ToString(), sIndex);

                            // 이월도 입력해야 됨. 

                            // 입고 ProdQty 1~31
                            LockStock.Add(setProdLocStock_Content(dr, "입고", "ProdQty", i - 1, sIndex++));

                            // 반품 ProdReturnQty 1~31 
                            LockStock.Add(setProdLocStock_Content(dr, "반품", "ProdReturnQty", i - 1, sIndex++));

                            // 샘플 SampleQty 1~31
                            LockStock.Add(setProdLocStock_Content(dr, "샘플", "SampleQty", i - 1, sIndex++));

                            // 재고 StockQty 1~31
                            LockStock.Add(setProdLocStock_Content(dr, "재고", "StockQty", i - 1, sIndex++));

                            int rowHeight = 30;

                            // 구분, 공정명 넣기
                            LockStock.RowHeight = LockStock.Count * rowHeight;
                            dgdProcess.Items.Add(LockStock);

                            // 이월 넣기
                            LockStock.RowHeight = LockStock.Count * rowHeight;
                            dgdCarryOver.Items.Add(LockStock);

                            // 마지막에 Collection 에 넣기
                            AllLockStock.Add(LockStock);
                        }

                        // 순번, 프로세스명
                        //var AllLockStock_Process = AllLockStock.Select(x => new { Num = x.Num, Process = x.Process, RowHeight = rowHeight * x.Count });
                        //dgdProcess.ItemsSource = AllLockStock_Process.ToList();

                        // 구분
                        dgdGubun.ItemsSource = AllLockStock.SelectMany(lockStock => lockStock).ToList();

                        // 이월
                        //var AllLockStock_CarryOver = AllLockStock.Select(x => new { CarryOver = x.CarryOver,  RowHeight = rowHeight * x.Count });
                        //dgdCarryOver.ItemsSource = AllLockStock_CarryOver.ToList();

                        // 내용
                        dgdContent.ItemsSource = AllLockStock.SelectMany(lockStock => lockStock).ToList();
                    }
                }

                strFlag = string.Empty;
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


        // header : 구분 이름
        // name : 컬럼 이름
        private ProdLocStock_Content setProdLocStock_Content(DataRow dr, string header, string name, int bIndex, int sIndex)
        {

            int i = 1;

            var ProdContent = new ProdLocStock_Content()
            {
                Gubun = header,

                BIndex = bIndex,
                SIndex = sIndex,

                Day1 = stringFormatN0AndZeroEmpty(dr[name + i++]),
                Day2 = stringFormatN0AndZeroEmpty(dr[name + i++]),
                Day3 = stringFormatN0AndZeroEmpty(dr[name + i++]),
                Day4 = stringFormatN0AndZeroEmpty(dr[name + i++]),
                Day5 = stringFormatN0AndZeroEmpty(dr[name + i++]),

                Day6 = stringFormatN0AndZeroEmpty(dr[name + i++]),
                Day7 = stringFormatN0AndZeroEmpty(dr[name + i++]),
                Day8 = stringFormatN0AndZeroEmpty(dr[name + i++]),
                Day9 = stringFormatN0AndZeroEmpty(dr[name + i++]),
                Day10 = stringFormatN0AndZeroEmpty(dr[name + i++]),

                Day11 = stringFormatN0AndZeroEmpty(dr[name + i++]),
                Day12 = stringFormatN0AndZeroEmpty(dr[name + i++]),
                Day13 = stringFormatN0AndZeroEmpty(dr[name + i++]),
                Day14 = stringFormatN0AndZeroEmpty(dr[name + i++]),
                Day15 = stringFormatN0AndZeroEmpty(dr[name + i++]),

                Day16 = stringFormatN0AndZeroEmpty(dr[name + i++]),
                Day17 = stringFormatN0AndZeroEmpty(dr[name + i++]),
                Day18 = stringFormatN0AndZeroEmpty(dr[name + i++]),
                Day19 = stringFormatN0AndZeroEmpty(dr[name + i++]),
                Day20 = stringFormatN0AndZeroEmpty(dr[name + i++]),

                Day21 = stringFormatN0AndZeroEmpty(dr[name + i++]),
                Day22 = stringFormatN0AndZeroEmpty(dr[name + i++]),
                Day23 = stringFormatN0AndZeroEmpty(dr[name + i++]),
                Day24 = stringFormatN0AndZeroEmpty(dr[name + i++]),
                Day25 = stringFormatN0AndZeroEmpty(dr[name + i++]),

                Day26 = stringFormatN0AndZeroEmpty(dr[name + i++]),
                Day27 = stringFormatN0AndZeroEmpty(dr[name + i++]),
                Day28 = stringFormatN0AndZeroEmpty(dr[name + i++]),
                Day29 = stringFormatN0AndZeroEmpty(dr[name + i++]),
                Day30 = stringFormatN0AndZeroEmpty(dr[name + i++]),

                Day31 = stringFormatN0AndZeroEmpty(dr[name + i++])
            };

            return ProdContent;
        }

        #endregion // 주요 메서드

        private void btnTest_Click(object sender, RoutedEventArgs e)
        {
            var rows = GetDataGridRows(dgdProcess);

            foreach (DataGridRow r in rows)
            {
                try
                {


                    foreach (DataGridColumn column in dgdProcess.Columns)
                    {
                        // column.GetCellContent(r) = {System.Windows.Controls.ContentPresenter}
                        if (column.GetCellContent(r) is ContentPresenter)
                        {
                            ContentPresenter cellContent = column.GetCellContent(r) as ContentPresenter;
                            Binding objBinding = null;
                            BindingExpression bindingExpression = (cellContent as ContentPresenter).GetBindingExpression(ContentPresenter.ContentProperty);

                            if (bindingExpression != null)
                            {
                                objBinding = bindingExpression.ParentBinding;
                                //MessageBox.Show(objBinding.ElementName);
                            }
                        }

                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }
        }

        public IEnumerable<DataGridRow> GetDataGridRows(DataGrid grid)
        {
            var itemsSource = grid.ItemsSource as IEnumerable;
            if (null == itemsSource) yield return null;
            foreach (var item in itemsSource)
            {
                var row = grid.ItemContainerGenerator.ContainerFromItem(item) as DataGridRow;

                try
                {
                    var hoit = row.DataContext as Win_mtr_ProdLocStock_Q_CodeView;
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }

                if (null != row) yield return row;


            }
        }

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

        // 천자리 콤마, 소수점 버리기
        private string stringFormatN0(object obj)
        {
            return string.Format("{0:N0}", obj);
        }

        // 0이면 빈칸, 아니면, 천자리 콤마, 소수점 버리기
        private string stringFormatN0AndZeroEmpty(object obj)
        {
            string result = string.Format("{0:N0}", obj);

            if (result.Trim().Equals("0"))
            {
                result = "";
            }

            return result;
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

        #region 데이터 그리드 이벤트

        // 대분류 선택 > 소분류 선택하기
        private void dgdProcess_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (!strFlag.Equals("S"))
            {
                //var Bigbrother = dgdProcess.SelectedItem as Win_mtr_ProdLocStock_Q_CodeView; 이거 안됨
                //int bIndex = dgdProcess.SelectedIndex;

                //if (Bigbrother != null)
                //{
                //    dgdGubun.SelectedIndex = Bigbrother.SIndex;
                //    dgdCarryOver.SelectedIndex = bIndex; // 대분류
                //    dgdContent.SelectedIndex = Bigbrother.SIndex;
                //}

                int bIndex = dgdProcess.SelectedIndex;
                int sIndex = 0;
                if (bIndex != 0)
                {
                    sIndex = (bIndex + 1) * 4 - 4;

                    //if (sIndex < 0)
                    //    sIndex = 0;
                }

                dgdGubun.SelectedIndex = sIndex;
                dgdCarryOver.SelectedIndex = bIndex; // 대분류
                dgdContent.SelectedIndex = sIndex;
            }

        }

        // 소분류 선택 > 대분류 선택하기 > 대분류 내에서 선택시에는 [대분류 선택]이 발동 안되야 함 > 이거 해봤는데 안해도 됨...
        private void dgdGubun_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (!strFlag.Equals("S"))
            {
                var Smallbrother = dgdGubun.SelectedItem as ProdLocStock_Content;
                int sIndex = dgdGubun.SelectedIndex;

                int bIndex = Smallbrother.BIndex;

                int endIndex = bIndex * 4 - 1;
                int startIndex = endIndex - 4;

                if (Smallbrother != null)
                {
                    dgdProcess.SelectedIndex = bIndex; // 대분류
                    dgdCarryOver.SelectedIndex = bIndex; // 대분류
                    dgdContent.SelectedIndex = sIndex;
                }
            }

        }

        // 대분류 선택 > 소분류 선택하기
        private void dgdCarryOver_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {

            if (!strFlag.Equals("S"))
            {
                int bIndex = dgdCarryOver.SelectedIndex;
                int sIndex = 0;
                if (bIndex != 0)
                {
                    sIndex = (bIndex + 1) * 4 - 4;
                }

                dgdGubun.SelectedIndex = sIndex;
                dgdProcess.SelectedIndex = bIndex; // 대분류
                dgdContent.SelectedIndex = sIndex;
            }
        }

        // 소분류 선택 > 대분류 선택하기 > 대분류 내에서 선택시에는 [대분류 선택]이 발동 안되야 함 > 이거 해봤는데 안해도 됨...
        private void dgdContent_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (!strFlag.Equals("S"))
            {
                var Smallbrother = dgdContent.SelectedItem as ProdLocStock_Content;
                int sIndex = dgdContent.SelectedIndex;

                int bIndex = Smallbrother.BIndex;

                int endIndex = bIndex * 4 - 1;
                int startIndex = endIndex - 4;

                if (Smallbrother != null)
                {
                    dgdProcess.SelectedIndex = bIndex; // 대분류
                    dgdCarryOver.SelectedIndex = bIndex; // 대분류
                    dgdGubun.SelectedIndex = sIndex;
                }
            }

        }

        #endregion // 데이터 그리드 이벤트

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

                    if (dgd.Name.Trim().Equals("dgdProcess"))
                    {
                        int startCol = 0;

                        if (startCol < currCol)
                        {
                            dgd.CurrentCell = new DataGridCellInfo(dgd.Items[currRow], dgd.Columns[currCol - 1]);
                        }
                    }
                    else if (dgd.Name.Trim().Equals("dgdGubun")) // 공정으로 이동
                    {
                        //var LotStock = dgd.SelectedItem as LotStock;

                        //if (LotStock != null)
                        //{
                        //    int sIndex = LotStock.sIndex;

                        dgdProcess.Focus();
                        dgdProcess.CurrentCell = dgdProcess.SelectedCells[0];
                        //= new DataGridCellInfo(dgdNum.Items[sIndex], dgdNum.Columns[0]);
                        //}
                    }
                    else if (dgd.Name.Trim().Equals("dgdCarryOver")) // 구분으로 이동
                    {
                        //var LotStock = dgd.SelectedItem as LotStock;

                        //if (LotStock != null)
                        //{
                        //    int bIndex = LotStock.bIndex;

                        dgdGubun.Focus();
                        dgdGubun.CurrentCell = dgdGubun.SelectedCells[0];
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

                            dgdCarryOver.Focus();
                            dgdCarryOver.CurrentCell = dgdCarryOver.SelectedCells[0];
                            //= new DataGridCellInfo(dgdLotID.Items[mIndex], dgdLotID.Columns[0]);
                            //}
                        }
                    }
                }
                else if (e.Key == Key.Right)
                {
                    e.Handled = true;

                    if (dgd.Name.Trim().Equals("dgdProcess")) // 구분으로 이동
                    {
                        //var LotStock = dgd.SelectedItem as LotStock_Content;

                        //if (LotStock != null)
                        //{
                        //    int bIndex = LotStock.bIndex;
                        int endCol = 1;

                        if (currCol < endCol)
                        {
                            dgd.CurrentCell = new DataGridCellInfo(dgd.Items[currRow], dgd.Columns[currCol + 1]);
                        }
                        else if (currCol == endCol)
                        {
                            dgdGubun.Focus();
                            dgdGubun.CurrentCell = dgdGubun.SelectedCells[0];
                        }

                        //= new DataGridCellInfo(dgdArticle.Items[bIndex], dgdArticle.Columns[0]);
                        //}
                    }
                    else if (dgd.Name.Trim().Equals("dgdGubun")) // 이월으로 이동
                    {
                        //var LotStock = dgd.SelectedItem as LotStock;

                        //if (LotStock != null)
                        //{
                        //    int mIndex = LotStock.mIndex;

                        dgdCarryOver.Focus();
                        dgdCarryOver.CurrentCell = dgdCarryOver.SelectedCells[0];
                        // = new DataGridCellInfo(dgdLotID.Items[mIndex], dgdLotID.Columns[0]);
                        //}
                    }
                    else if (dgd.Name.Trim().Equals("dgdCarryOver")) // 데이터로 이동
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
                        int endCol = 30;

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

    }

    #region 코드뷰 CodeView

    class Win_mtr_ProdLocStock_Q_CodeView : List<ProdLocStock_Content>
    {
        public Win_mtr_ProdLocStock_Q_CodeView() { }

        public Win_mtr_ProdLocStock_Q_CodeView(int num, string process, string carryOver, int sIndex)
        {
            this.Num = num;
            this.Process = process;
            this.CarryOver = carryOver;
            //this.CarryOver = "이월임";
            this.SIndex = sIndex;
        }

        public int Num { get; set; }
        public string Process { get; set; } // 공정명

        public string CarryOver { get; set; } // 이월

        public int SIndex { get; set; } // 소분류 첫번째 인덱스 넣기

        public int RowHeight { get; set; }
    }

    class ProdLocStock_Collection : List<Win_mtr_ProdLocStock_Q_CodeView>
    {

    }

    class ProdLocStock_Content
    {

        public int BIndex { get; set; } // 대분류 인덱스
        public int SIndex { get; set; }

        public string Gubun { get; set; }

        public string Day1 { get; set; }
        public string Day2 { get; set; }
        public string Day3 { get; set; }
        public string Day4 { get; set; }
        public string Day5 { get; set; }

        public string Day6 { get; set; }
        public string Day7 { get; set; }
        public string Day8 { get; set; }
        public string Day9 { get; set; }
        public string Day10 { get; set; }

        public string Day11 { get; set; }
        public string Day12 { get; set; }
        public string Day13 { get; set; }
        public string Day14 { get; set; }
        public string Day15 { get; set; }

        public string Day16 { get; set; }
        public string Day17 { get; set; }
        public string Day18 { get; set; }
        public string Day19 { get; set; }
        public string Day20 { get; set; }

        public string Day21 { get; set; }
        public string Day22 { get; set; }
        public string Day23 { get; set; }
        public string Day24 { get; set; }
        public string Day25 { get; set; }

        public string Day26 { get; set; }
        public string Day27 { get; set; }
        public string Day28 { get; set; }
        public string Day29 { get; set; }
        public string Day30 { get; set; }

        public string Day31 { get; set; }

    }
    #endregion // 코드뷰 CodeView
}

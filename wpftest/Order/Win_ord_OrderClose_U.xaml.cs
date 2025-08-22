using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Data;
using System.Windows.Data;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using WizMes_WellMade.PopUP;
using WizMes_WellMade.PopUp;
using WPF.MDI;
using System.Linq;
using System.Windows.Controls.Primitives;
using System.Windows.Media;
using System.Text.RegularExpressions;

namespace WizMes_WellMade
{
    /// <summary>
    /// Win_ord_OrderClose_U.xaml에 대한 상호 작용 논리
    /// </summary>
    public partial class Win_ord_OrderClose_U : UserControl
    {
        string stDate = string.Empty;
        string stTime = string.Empty;

        private Microsoft.Office.Interop.Excel.Application excelapp;
        private Microsoft.Office.Interop.Excel.Workbook workbook;
        private Microsoft.Office.Interop.Excel.Worksheet worksheet;
        private Microsoft.Office.Interop.Excel.Range workrange;
        private Microsoft.Office.Interop.Excel.Worksheet copysheet;
        private Microsoft.Office.Interop.Excel.Worksheet pastesheet;

        private ToolTip toolTip = new ToolTip();
        Win_ord_OrderClose_U_CodeView WinOrderClose = new Win_ord_OrderClose_U_CodeView();
        Lib lib = new Lib();
        string rowHeaderNum = string.Empty;
        int rowNum = 0;
        int rbnOrder = 0;

        NoticeMessage msg = new NoticeMessage();
        DataTable DT;
        ////private List<DataGridColumn> _dynamicColumns = new List<DataGridColumn>();

        public Win_ord_OrderClose_U()
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
            SetComboBox();
            Check_bdrOrder();

        }

        //콤보박스 세팅
        private void SetComboBox()
        {
            List<string> strValue = new List<string>();
            strValue.Add("전체");
            strValue.Add("진행건");
            strValue.Add("마감건");

            ObservableCollection<CodeView> cbOrderStatus = ComboBoxUtil.Instance.Direct_SetComboBox(strValue);
            cboOrderStatusSrh.ItemsSource = cbOrderStatus;
            cboOrderStatusSrh.DisplayMemberPath = "code_name";
            cboOrderStatusSrh.SelectedValuePath = "code_id";
            cboOrderStatusSrh.SelectedIndex = 0;

            // 주문 구분
            ObservableCollection<CodeView> ovcOrderClss = ComboBoxUtil.Instance.Gf_DB_CM_GetComCodeDataset(null, "ORDGBN", "Y", "", "");
            cboOrderClssSrh.ItemsSource = ovcOrderClss;
            cboOrderClssSrh.DisplayMemberPath = "code_name";
            cboOrderClssSrh.SelectedValuePath = "code_id";
            cboOrderClssSrh.SelectedIndex = 0;

        }

        #region 라벨 체크박스 이벤트 관련

        //일자
        private void lblOrderDay_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            if (chkOrderDay.IsChecked == true) { chkOrderDay.IsChecked = false; }
            else { chkOrderDay.IsChecked = true; }
        }

        //일자
        private void chkOrderDay_Checked(object sender, RoutedEventArgs e)
        {
            if (dtpSDate != null && dtpEDate != null)
            {
                dtpSDate.IsEnabled = true;
                dtpEDate.IsEnabled = true;
            }
        }

        //일자
        private void chkOrderDay_Unchecked(object sender, RoutedEventArgs e)
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

        //금월
        private void btnThisMonth_Click(object sender, RoutedEventArgs e)
        {
            dtpSDate.SelectedDate = Lib.Instance.BringThisMonthDatetimeList()[0];
            dtpEDate.SelectedDate = Lib.Instance.BringThisMonthDatetimeList()[1];
        }

        ////금년
        //private void btnThisYear_Click(object sender, RoutedEventArgs e)
        //{
        //    dtpSDate.SelectedDate = Lib.Instance.BringThisYearDatetimeFormat()[0];
        //    dtpEDate.SelectedDate = Lib.Instance.BringThisYearDatetimeFormat()[1];
        //}

        //전일
        private void BtnYesterDay_Click(object sender, RoutedEventArgs e)
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

        //수주상태
        private void cboOrderStatusSrh_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
          
                if (cboOrderStatusSrh.SelectedIndex == 0)
                {
                    btnFinal.IsEnabled = false;
                }
                else if (cboOrderStatusSrh.SelectedIndex == 1)
                {
                    btnFinal.IsEnabled = true;
                    btnFinal.Content = "마감처리";
                }
                else
                {
                    btnFinal.IsEnabled = true;
                    btnFinal.Content = "진행처리";
                }
            
          
         
        }

        //수주 진행 건은 마감처리 / 마감 건은 진행처리로 변경하는 버튼
        private void BtnFinal_Click(object sender, RoutedEventArgs e)
        {
            //string OrderID = string.Empty;

            // 다중선택 했을 때 각각 OrderID 들어가도록 설정했으므로 이건 안써도 돼
            //var Order = dgdMain.SelectedItem as Win_ord_OrderClose_U_CodeView;
            //if (Order != null)
            //{
            //    OrderID = Order.OrderID;
            //}

            string CloseFlag = string.Empty;
            string CloseClss = string.Empty;

            if (btnFinal.Content.ToString().Equals("마감처리"))
            {
                CloseFlag = "1";
                CloseClss = "1";

                if (MessageBox.Show("해당 건을 마감처리 하시겠습니까?", "처리 전 확인", MessageBoxButton.YesNo) == MessageBoxResult.Yes)
                {
                    if (dgdMain.Items.Count > 0 && dgdMain.SelectedItem != null)
                    {
                        rowNum = dgdMain.SelectedIndex;
                    }
                }
            }
            else if (btnFinal.Content.ToString().Equals("진행처리"))
            {
                CloseFlag = "2";
                CloseClss = "";

                if (MessageBox.Show("해당 건을 진행처리 하시겠습니까?", "처리 전 확인", MessageBoxButton.YesNo) == MessageBoxResult.Yes)
                {
                    if (dgdMain.Items.Count > 0 && dgdMain.SelectedItem != null)
                    {
                        rowNum = dgdMain.SelectedIndex;
                    }
                }
            }

            List<Procedure> Prolist = new List<Procedure>();
            List<Dictionary<string, object>> ListParameter = new List<Dictionary<string, object>>();
            try
            {
                //일괄처리할 때 쓰는 변수
                int CheckCount = 0;

                //데이터그리드의 체크박스 true된 수 많음 CheckCount 수 늘리기
                foreach (Win_ord_OrderClose_U_CodeView OrderCloseU in dgdMain.Items)
                {
                    if (OrderCloseU.IsCheck == true)
                    {
                        CheckCount++;
                    }
                }

                //체크된 그리드가 하나 이상일 경우(1개라도 체크가 되어 있을 경우)
                if (CheckCount > 0)
                {
                    foreach (Win_ord_OrderClose_U_CodeView OrderCloseU in dgdMain.Items)
                    {
                        if (OrderCloseU != null)
                        {
                            if (OrderCloseU.IsCheck == true)
                            {
                                Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                                sqlParameter.Clear();
                                sqlParameter.Add("CloseFlag", CloseFlag);
                                sqlParameter.Add("OrderID", OrderCloseU.OrderID);
                                sqlParameter.Add("CloseClss", CloseClss);

                                Procedure pro1 = new Procedure();
                                pro1.Name = "xp_OrderClose_uCloseClss";     //마감처리 누르면 CloseClss에 1 저장, 진행처리 누르면 '' 저장 Order테이블에.
                                pro1.OutputUseYN = "N";
                                pro1.OutputName = "OrderID";
                                pro1.OutputLength = "10";

                                Prolist.Add(pro1);
                                ListParameter.Add(sqlParameter);
                            }
                        }
                    }

                    string[] Confirm = new string[2];
                    Confirm = DataStore.Instance.ExecuteAllProcedureOutputNew(Prolist, ListParameter);
                    if (Confirm[0] != "success")
                    {
                        MessageBox.Show("[저장실패]\r\n" + Confirm[1].ToString());
                    }
                }
                else
                {
                    MessageBox.Show("[저장실패]\r\n 처리할 체크항목이 없습니다.");
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

            dgdMain.Items.Clear();
            FillGrid();
        }


        //검색조건 - 거래처 - 키다운
        private void txtCustomIDSrh_KeyDown(object sender, KeyEventArgs e)
        {
            if(e.Key == Key.Enter)
            {
                MainWindow.pf.ReturnCode(txtCustomIDSrh, 0, "");
            }
        }
        //검색조건 - 거래처 - 버튼 클릭

        private void btnCustomIDSrh_Click(object sender, RoutedEventArgs e)
        {
            MainWindow.pf.ReturnCode(txtCustomIDSrh, 0, "");
        }


        //검색조건 - 품명 - 키다운
        private void txtArticleIDSrh_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
                MainWindow.pf.ReturnCode(txtCustomIDSrh, 77, "");
        }

        //검색조건 - 품명 - 버튼 클릭
        private void btnArticleIDSrh_Click(object sender, RoutedEventArgs e)
        {
            MainWindow.pf.ReturnCode(txtCustomIDSrh, 77, "");
        }

        //검색조건 - 품번 - 키다운
        private void txtBuyerArticleNoSrh_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
                MainWindow.pf.ReturnCode(txtBuyerArticleNoSrh, 7071, "");
        }
        //검색조건 - 품번 - 버튼 클릭
        private void btnBuyerArticleNoSrh_KeyDown(object sender, RoutedEventArgs e)
        {
              MainWindow.pf.ReturnCode(txtBuyerArticleNoSrh, 7071, "");
        }




        private void rbnOrderNo_Click(object sender, RoutedEventArgs e)
        {

            Check_bdrOrder();
        }

        private void rbnOrderID_Click(object sender, RoutedEventArgs e)
        {
            Check_bdrOrder();
        }

        private void Check_bdrOrder()
        {
            if (rbnOrderID.IsChecked == true)
            {
                tbkOrder.Text = " 관리번호";
                dgdtxtOrderID.Visibility = Visibility.Visible;
                dgdtxtOrderNo.Visibility = Visibility.Hidden;
            }
            else if (rbnOrderNo.IsChecked == true)
            {
                tbkOrder.Text = "Order No";
                dgdtxtOrderID.Visibility = Visibility.Hidden;
                dgdtxtOrderNo.Visibility = Visibility.Visible;
            }
        }


        private void CommonControl_Click(object sender, EventArgs e)
        {
            CheckBox checkBox = null;
            DependencyObject parentGrid = null;

            if (sender is Label label)
            {
                // 라벨의 부모 그리드 찾기
                parentGrid = FindVisualParent<Grid>(label);
                if (parentGrid != null)
                {
                    // 같은 그리드 내에서 체크박스 찾기
                    checkBox = FindChild<CheckBox>(parentGrid);
                    if (checkBox != null)
                    {
                        // 체크박스 상태 토글
                        checkBox.IsChecked = !checkBox.IsChecked;
                    }
                }
            }
            else if (sender is CheckBox clickedCheckBox)
            {
                // 클릭된 것이 체크박스인 경우
                checkBox = clickedCheckBox;
                parentGrid = FindVisualParent<Grid>(checkBox);
            }

            // 체크박스와 부모 그리드가 있으면 컨트롤 활성화/비활성화 처리
            if (checkBox != null && parentGrid != null)
            {
                List<Control> controlsToToggle = new List<Control>();

                // 그리드 내 모든 Control 찾기 (체크박스 제외)
                FindUiObject(parentGrid, obj => {
                    if (obj is Control control && obj != checkBox && !(obj is Label) && !(obj is CheckBox))
                    {
                        controlsToToggle.Add(control);
                    }
                });

                // 컨트롤 활성화/비활성화
                foreach (var control in controlsToToggle)
                {
                    control.IsEnabled = checkBox.IsChecked == true;
                }
            }
        }

        //UI컨트롤 요소찾기
        private void FindUiObject(DependencyObject parent, Action<DependencyObject> action)
        {
            int childCount = VisualTreeHelper.GetChildrenCount(parent);
            for (int i = 0; i < childCount; i++)
            {
                var child = VisualTreeHelper.GetChild(parent, i);

                action?.Invoke(child);

                FindUiObject(child, action);
            }
        }

        //컨트롤 안 특정 타입의 자식 컨트롤을 찾는 함수 (그리드내에서)
        //var parentContainer = VisualTreeHelper.GetParent(checkbox);
        //var datePicker = FindChild<DatePicker>(parentContainer);
        private T FindChild<T>(DependencyObject parent) where T : DependencyObject
        {
            int childCount = VisualTreeHelper.GetChildrenCount(parent);
            for (int i = 0; i < childCount; i++)
            {
                var child = VisualTreeHelper.GetChild(parent, i);
                if (child is T typedChild)
                {
                    return typedChild;
                }

                // 재귀적으로 자식의 자식들도 검색
                var result = FindChild<T>(child);
                if (result != null)
                    return result;
            }
            return null;
        }


        // 자식요소 안에서 부모요소 찾기
        //DataGridRow row = FindVisualParent<DataGridRow>(checkBox); 데이터그리드안의 행속 체크박스의 부모행 찾기
        //DataGrid parentGrid = FindVisualParent<DataGrid>(row); 데이터그리드 행의 부모 데이터그리드 찾기
        private T FindVisualParent<T>(DependencyObject child) where T : DependencyObject
        {
            DependencyObject parentObject = VisualTreeHelper.GetParent(child);

            if (parentObject == null)
                return null;

            T parent = parentObject as T;
            if (parent != null)
                return parent;
            else
                return FindVisualParent<T>(parentObject);
        }


        //8자리 char형태 날짜 년도-월-일 하이픈 삽입
        //16자리 일경우 8자리 사이에 ~ 삽입
        private string DateTypeHyphen(string DigitsDate)
        {
            string pattern1 = @"(\d{4})(\d{2})(\d{2})";
            string pattern2 = @"(\d{4})(\d{2})(\d{2})(\d{4})(\d{2})(\d{2})";
            string pattern3 = @"(\d{4})(\d{2})(\d{2})(\d{2})(\d{2})";
            string pattern4 = @"(\d{2})(\d{2})";

            if (DigitsDate.Length == 8)
            {
                DigitsDate = Regex.Replace(DigitsDate, pattern1, "$1-$2-$3");
            }      
            else if (DigitsDate.Length == 16)
            {
                DigitsDate = Regex.Replace(DigitsDate, pattern2, "$1-$2-$3 ~ $4-$5-$6");
            }
            else if (DigitsDate.Length == 13)
            {
                DigitsDate= DigitsDate.Replace("/", "");
                DigitsDate = Regex.Replace(DigitsDate, pattern3, "$1-$2-$3 $4:$5");
            }
            else if (DigitsDate.Length == 4)
            {
                DigitsDate = Regex.Replace(DigitsDate, pattern4, "$1:$2");

            }
            else if (DigitsDate.Length == 0)
            {
                DigitsDate = string.Empty;
            }

            if (DigitsDate.Length == 1 && DigitsDate.Equals("/")) DigitsDate = string.Empty;

            return DigitsDate;
        }

        private object RemoveComma(object obj, bool returnAsNumber = false, Type returnType = null)
        {
            //파라미터가 만약 null일때
            if (obj == null)
            {
                //숫자타입이 false면 string으로 내보내기
                if (!returnAsNumber) return "0";

                // 만약 숫자타입을 써야되면 returnType파라미터의 받은 형태로 전달
                // null일 때도 returnType에 따라 적절한 타입의 0 반환
                switch (returnType?.Name)
                {
                    case "Decimal": return (object)0m;  //monetary
                    case "Double": return (object)0d;   //double
                    case "Int64": return (object)0L;    //long
                    default: return (object)0;          //int
                }
            }

            string digits = obj.ToString()
                              .Trim()
                              .Replace(",", "");

            //만약 빈공백(blank)이더라도 0으로 내보내야한다.
            if (string.IsNullOrEmpty(digits))
            {
                if (!returnAsNumber) return "0";

                // returnType을 활용해서 적절한 타입으로 반환
                switch (returnType?.Name)
                {
                    case "Decimal": return (object)0m;
                    case "Double": return (object)0d;
                    case "Int64": return (object)0L;
                    default: return (object)0;
                }
            }


            try
            {
                Type targetType = returnType ?? typeof(int);

                //혹시나 하는 예외처리
                //입력 컨트롤간에 LostFocus나 TextChanged같은 걸로 계산을 할 때
                //처리 가능한 숫자 범위를 초과하면 오류가 발생하므로
                //초과하면 해당 자료형타입이 처리할 수 있는 최대 숫자를 표시해줌
                switch (targetType.Name)
                {
                    case "Int32":
                        if (decimal.TryParse(digits, out decimal intParsed))
                        {
                            if (intParsed > int.MaxValue) return int.MaxValue;
                            if (intParsed < int.MinValue) return int.MinValue;
                            return (int)intParsed;
                        }
                        return int.MaxValue;

                    case "Int64":
                        if (decimal.TryParse(digits, out decimal longParsed))
                        {
                            if (longParsed > long.MaxValue) return long.MaxValue;
                            if (longParsed < long.MinValue) return long.MinValue;
                            return (long)longParsed;
                        }
                        return long.MaxValue;

                    case "Double":
                        if (double.TryParse(digits, out double doubleParsed))
                        {
                            return doubleParsed;
                        }
                        return double.MaxValue;

                    case "Decimal":
                        if (decimal.TryParse(digits, out decimal decimalParsed))
                        {
                            return decimalParsed;
                        }
                        return decimal.MaxValue;

                    default:
                        return int.MaxValue;
                }
            }
            catch
            {

                if (returnType != null)
                {
                    switch (returnType.Name)
                    {
                        case "Int32":
                            return int.MaxValue;
                        case "Int64":
                            return long.MaxValue;
                        case "Double":
                            return double.MaxValue;
                        case "Decimal":
                            return decimal.MaxValue;
                        default:
                            return int.MaxValue;
                    }
                }
                return int.MaxValue;
            }
        }

        private string ConvertDate(DatePicker datePicker)
        {
            if (datePicker.SelectedDate != null)
                return datePicker.SelectedDate.Value.ToString("yyyyMMdd");
            else
                return string.Empty;
        }

        private bool IsDatePickerNull(DatePicker datePicker)
        {
            if (datePicker.SelectedDate == null)
                return true;
            else
                return false;
        }



        #endregion

        private void btnSearch_Click(object sender, RoutedEventArgs e)
        {
            if(lib.DatePickerCheck(dtpSDate, dtpEDate, chkOrderDay))
            {
                using (Loading ld = new Loading(re_Search))
                {
                    ld.ShowDialog();
                }
            }
        }

        //닫기
        private void btnClose_Click(object sender, RoutedEventArgs e)
        {
            DataStore.Instance.InsertLogByFormS(this.GetType().Name, stDate, stTime, "E");
            Lib.Instance.ChildMenuClose(this.ToString());
        }

        //인쇄
        private void btnPrint_Click(object sender, RoutedEventArgs e)
        {
            ContextMenu menu = btnPrint.ContextMenu;
            menu.StaysOpen = true;
            menu.IsOpen = true;
        }

        //인쇄 미리보기
        private void menuSeeAhead_Click(object sender, RoutedEventArgs e)
        {
            if (dgdMain.Items.Count < 1)
            {
                MessageBox.Show("먼저 검색해 주세요.");
                return;
            }

            msg.Show();
            msg.Topmost = true;
            msg.Refresh();

            PrintWork(true);
        }

        //바로 인쇄
        private void menuRightPrint_Click(object sender, RoutedEventArgs e)
        {
            if (dgdMain.Items.Count < 1)
            {
                MessageBox.Show("먼저 검색해 주세요.");
                return;
            }

            DataStore.Instance.InsertLogByForm(this.GetType().Name, "P");
            msg.Show();
            msg.Topmost = true;
            msg.Refresh();

            PrintWork(false);
        }

        private void menuClose_Click(object sender, RoutedEventArgs e)
        {
            ContextMenu menu = btnPrint.ContextMenu;
            menu.StaysOpen = false;
            menu.IsOpen = false;
        }

        //엑셀
        private void btnExcel_Click(object sender, RoutedEventArgs e)
        {
            DataTable dt = null;

            string[] lst = new string[2];
            lst[0] = "수주조회";
            lst[1] = dgdMain.Name;
            Lib lib = new Lib();

            ExportExcelxaml ExpExc = new ExportExcelxaml(lst);

            ExpExc.ShowDialog();

            if (ExpExc.DialogResult.HasValue)
            {
                if (ExpExc.choice.Equals(dgdMain.Name))
                {
                    DataStore.Instance.InsertLogByForm(this.GetType().Name, "E");
                    //MessageBox.Show("대분류");
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
            lib = null;
        }

        //실조회 및 하단 합계
        private void FillGrid()
        {
                dgdMain.Items.Clear();
                dgdSum.Items.Clear();



            try
            {
                Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                sqlParameter.Clear();

                sqlParameter.Add("ChkDate", chkOrderDay.IsChecked == true ? 1 : 0);
                sqlParameter.Add("SDate", chkOrderDay.IsChecked == true ? dtpSDate.SelectedDate.Value.ToString("yyyyMMdd") : "");
                sqlParameter.Add("EDate", chkOrderDay.IsChecked == true ? dtpEDate.SelectedDate.Value.ToString("yyyyMMdd") : "");

                // 거래처
                sqlParameter.Add("ChkCustomID", chkCustomIDSrh.IsChecked == true ? 1 : 0);
                sqlParameter.Add("CustomID", chkCustomIDSrh.IsChecked == true ? txtCustomIDSrh.Tag != null ? txtCustomIDSrh.Tag.ToString() : "" : "");
      
                // 품명
                sqlParameter.Add("ChkArticleID", chkArticleIDSrh.IsChecked == true ? 1 : 0);
                sqlParameter.Add("ArticleID", chkArticleIDSrh.IsChecked == true ? txtArticleIDSrh.Tag != null ? txtArticleIDSrh.Tag.ToString() : "" : "");

                //품번
                sqlParameter.Add("ChkBuyerArticleNo", chkBuyerArticleNoSrh.IsChecked == true ? 1 : 0);
                sqlParameter.Add("BuyerArticleNo", chkBuyerArticleNoSrh.IsChecked == true ? txtBuyerArticleNoSrh.Tag != null ? txtBuyerArticleNoSrh.Tag.ToString() : "" : "");

                //주문구분
                sqlParameter.Add("ChkOrderClss", chkOrderClssSrh.IsChecked == true ? 1 : 0);
                sqlParameter.Add("OrderClss", chkOrderClssSrh.IsChecked == true ? cboOrderClssSrh.SelectedValue != null ? cboOrderClssSrh.SelectedValue.ToString() : "" : "");

                // 관리번호
                sqlParameter.Add("ChkOrderID", chkOrderNoSrh.IsChecked == true ? (rbnOrderID.IsChecked == true ? 1 : 2) : 0);
                sqlParameter.Add("OrderID", txtOrderNoSrh.Text);
               
                // 수주상태
                sqlParameter.Add("ChkClose", int.Parse(cboOrderStatusSrh.SelectedValue != null ? cboOrderStatusSrh.SelectedValue.ToString() : ""));



                DataSet ds = DataStore.Instance.ProcedureToDataSet_LogWrite("xp_Order_sOrderTotal", sqlParameter, true, "R");

                if (ds != null && ds.Tables.Count > 0)
                {
                    DataTable dt = ds.Tables[0];

                    //dataGrid.Items.Clear();
                    if (dt.Rows.Count == 0)
                    {

                        MessageBox.Show("조회된 데이터가 없습니다.");
                    }
                    else
                    {
                        DataRowCollection drc = dt.Rows;

                        int i = 0;
                        int orderSum = 0;
                        int insertSum = 0;
                        int realQtySum = 0;
                        int outQtySum = 0;
                        int outQtyYetSum = 0;

                        foreach(DataRow dr in drc)
                        {
                            i++;
                            var OrderItem = new Win_ord_OrderClose_U_CodeView
                            {
                                Num = i,
                                OrderID = dr["OrderID"].ToString(),
                                OrderNo = dr["OrderNo"].ToString(),
                                KCustom = dr["KCustom"].ToString(),
                                BuyerArticleNo = dr["BuyerArticleNo"].ToString(),
                                Article = dr["Article"].ToString(),
                                BuyerModel = dr["BuyerModel"].ToString(),
                                AcptDate = DateTypeHyphen(dr["AcptDate"].ToString()),
                                DvlyDate = DateTypeHyphen(dr["DvlyDate"].ToString()),
                                //OrderSpec = dr["OrderSpec"].ToString(),
                                OrderQty = stringFormatN0(dr["OrderQty"]),
                                UnitClssName = dr["UnitClssName"].ToString(),
                                InputDate = DateTypeHyphen(dr["InputDate"].ToString()),
                                InputQty = stringFormatN0(dr["InputQty"]),
                                ExamDate = DateTypeHyphen(dr["ExamDate"].ToString()),                                
                                RealQty = stringFormatN0(dr["RealQty"]),
                                OutQty = stringFormatN0(dr["OutQty"]),
                                OutDate = DateTypeHyphen(dr["OutDate"].ToString()),
                                OutQtyYet = stringFormatN0(dr["OutQtyYet"]),
                                Remark = dr["Remark"].ToString(),
                            };

                            orderSum += (int)RemoveComma(dr["OrderQty"].ToString(), true);
                            insertSum += (int)RemoveComma(dr["InputQty"].ToString(), true);
                            realQtySum += (int)RemoveComma(dr["RealQty"].ToString(), true);
                            outQtySum += (int)RemoveComma(dr["OutQty"].ToString(), true);
                            outQtyYetSum += (int)RemoveComma(dr["OutQtyYet"].ToString(), true);


                            dgdMain.Items.Add(OrderItem);
                        }

                        if(dgdMain.Items.Count > 0)
                        {
                            var OrderTotal = new dgOrderSum
                            {
                                Count = i,
                                OrderSum = orderSum,
                                OutQtySum = outQtySum,
                                OutQtyYetSum = outQtyYetSum,
                                ReaQtySum = realQtySum,
                                InsertSum = insertSum,
                            };

                            dgdSum.Items.Add(OrderTotal);
                        }


                    }
         

                }
                if(ds.Tables.Count == 0)
                {
                    MessageBox.Show("조회된 데이터가 없습니다.");   
                    
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

        

        private bool HasNonNullValue(DataRowCollection drc, string propertyName)
        {
            foreach (DataRow row in drc)
            {
                if (row[propertyName] != null && !string.IsNullOrEmpty(row[propertyName].ToString()))
                {
                    return true;
                }
            }
            return false;
        }


        //전체선택
        private void btnAllCheck_Click(object sender, RoutedEventArgs e)
        {
            foreach (Win_ord_OrderClose_U_CodeView woccv in dgdMain.Items)
            {
                woccv.IsCheck = true;
            }
        }

        //선택해제
        private void btnAllNone_Click(object sender, RoutedEventArgs e)
        {
            foreach (Win_ord_OrderClose_U_CodeView woccv in dgdMain.Items)
            {
                woccv.IsCheck = false;
            }
        }

        //인쇄 실질 동작
        private void PrintWork(bool preview_click)
        {
            Lib lib2 = new Lib();

            try
            {
                excelapp = new Microsoft.Office.Interop.Excel.Application();

                string MyBookPath = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetEntryAssembly().Location) + "\\Report\\수주진행현황(영업관리).xls";
                //MyBookPath = MyBookPath.Substring(0, MyBookPath.LastIndexOf("\\")) + "\\order_standard.xls";
                //string MyBookPath = "C:/Users/Administrator/Desktop/order_standard.xls";
                workbook = excelapp.Workbooks.Add(MyBookPath);
                worksheet = workbook.Sheets["Form"];

                //상단의 일자 
                if (chkOrderDay.IsChecked == true)
                {
                    workrange = worksheet.get_Range("E2", "Q2");//셀 범위 지정
                    workrange.Value2 = dtpSDate.Text + "~" + dtpEDate.Text;
                    workrange.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                    //workrange.Font.Size = 10;
                }
                else
                {
                    workrange = worksheet.get_Range("E2", "K2");//셀 범위 지정
                    workrange.Value2 = "전체"; //"" + "~" + "";
                    workrange.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                    //workrange.Font.Size = 10;
                }


                //오더번호 혹은 관리번호 
                if (rbnOrderNo.IsChecked == true)
                {
                    workrange = worksheet.get_Range("C5", "F5");//셀 범위 지정
                    workrange.Value2 = "오더번호";
                    workrange.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                    //workrange.Font.Size = 10;
                }
                else
                {
                    workrange = worksheet.get_Range("C5", "F5");//셀 범위 지정
                    workrange.Value2 = "관리번호";
                    workrange.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                    //workrange.Font.Size = 10;
                }

                //하단의 회사명
                workrange = worksheet.get_Range("AN35", "AU35");//셀 범위 지정
                workrange.Value2 = "부경테크";
                workrange.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                workrange.Font.Size = 11;


                /////////////////////////
                int Page = 0;
                int DataCount = 0;
                int copyLine = 0;

                copysheet = workbook.Sheets["Form"];
                pastesheet = workbook.Sheets["Print"];

                DT = lib2.DataGirdToDataTable(dgdMain);

                string str_Num = string.Empty;
                string str_OrderID = string.Empty;
                string str_OrderID_CV = string.Empty;
                string str_KCustom = string.Empty;
                string str_Article = string.Empty;
                string str_Model = string.Empty;
                string str_ArticleNo = string.Empty;
                string str_DvlyDate = string.Empty;
                string str_Work = string.Empty;
                string str_OrderQty = string.Empty;
                string str_UnitClssName = string.Empty;
                string str_DayAndTime = string.Empty;
                string str_p1WorkQty = string.Empty;
                string str_InspectQty = string.Empty;
                string str_PassQty = string.Empty;
                string str_DefectQty = string.Empty;
                string str_OutQty = string.Empty;

                int TotalCnt = dgdMain.Items.Count;
                int canInsert = 27; //데이터가 입력되는 행 수 27개

                int PageCount = (int)Math.Ceiling(1.0 * TotalCnt / canInsert);

                var Sum = new dgOrderSum();

                //while (dgdMain.Items.Count > DataCount + 1)
                for (int k = 0; k < PageCount; k++)
                {
                    Page++;
                    if (Page != 1) { DataCount++; }  //+1
                    copyLine = (Page - 1) * 38;
                    copysheet.Select();
                    copysheet.UsedRange.Copy();
                    pastesheet.Select();
                    workrange = pastesheet.Cells[copyLine + 1, 1];
                    workrange.Select();
                    pastesheet.Paste();

                    int j = 0;
                    for (int i = DataCount; i < dgdMain.Items.Count; i++)
                    {
                        if (j == 27) { break; }
                        int insertline = copyLine + 7 + j;

                        str_Num = (j + 1).ToString();
                        str_OrderID       = DT.Rows[i][1] != null ? DT.Rows[i][1].ToString() : "";
                        str_OrderID_CV    = DT.Rows[i][2] != null ? DT.Rows[i][2].ToString() : "";
                        str_KCustom       = DT.Rows[i][3] != null ? DT.Rows[i][3].ToString() : "";
                        str_ArticleNo     = DT.Rows[i][4] != null ? DT.Rows[i][4].ToString() : "";
                        str_Article       = DT.Rows[i][5] != null ? DT.Rows[i][4].ToString() : "";
                        str_Model         = string.Empty;                       
                        str_DvlyDate      = DT.Rows[i][7] != null ? DT.Rows[i][7].ToString() : "";
                        str_Work          = string.Empty;
                        str_OrderQty      = DT.Rows[i][8] != null ? DT.Rows[i][8].ToString() : ""; 
                        str_UnitClssName  = DT.Rows[i][9] != null ? DT.Rows[i][9].ToString() : "";
                        str_DayAndTime    = string.Empty;
                        str_p1WorkQty     = string.Empty;
                        str_InspectQty    = DT.Rows[i][13].ToString();
                        str_PassQty       = DT.Rows[i][13].ToString();
                        str_DefectQty     = "0";
                        str_OutQty        = DT.Rows[i][15].ToString();

                        workrange = pastesheet.get_Range("A" + insertline, "B" + insertline);    //순번
                        workrange.Value2 = str_Num;
                        workrange.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                        workrange.Font.Size = 10;
                        workrange.ColumnWidth = 1.3;

                        if (dgdtxtOrderID.ToString().Equals("오더번호"))
                        {
                            workrange = pastesheet.get_Range("C" + insertline, "F" + insertline);    //오더번호
                            workrange.Value2 = str_OrderID;
                            workrange.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                            workrange.Font.Size = 9;
                            workrange.ColumnWidth = 1.8;
                        }
                        else
                        {
                            workrange = pastesheet.get_Range("C" + insertline, "F" + insertline);    //관리번호
                            workrange.Value2 = str_OrderID_CV;
                            workrange.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                            workrange.Font.Size = 9;
                            workrange.ColumnWidth = 1.8;
                        }

                        workrange = pastesheet.get_Range("G" + insertline, "J" + insertline);     //거래처
                        workrange.Value2 = str_KCustom;
                        workrange.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                        workrange.Font.Size = 9;
                        workrange.ColumnWidth = 2.7;

                        workrange = pastesheet.get_Range("K" + insertline, "N" + insertline);    //품명
                        workrange.Value2 = str_Article;
                        workrange.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                        workrange.Font.Size = 10;
                        workrange.ColumnWidth = 2.7;

                        workrange = pastesheet.get_Range("O" + insertline, "R" + insertline);    //차종 -> 재질
                        workrange.Value2 = str_Model;
                        workrange.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                        workrange.Font.Size = 10;
                        workrange.ColumnWidth = 0.9;

                        workrange = pastesheet.get_Range("S" + insertline, "V" + insertline);    //품번
                        workrange.Value2 = str_ArticleNo;
                        workrange.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                        workrange.Font.Size = 10;
                        workrange.ColumnWidth = 2.7;

                        workrange = pastesheet.get_Range("W" + insertline, "Y" + insertline);    //가공구분 -> 수주일
                        workrange.Value2 = str_Work;
                        workrange.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                        workrange.Font.Size = 10;
                        workrange.ColumnWidth = 1.8;

                        workrange = pastesheet.get_Range("Z" + insertline, "AA" + insertline);    //납기일
                        workrange.Value2 = str_DvlyDate;
                        workrange.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                        workrange.Font.Size = 10;
                        workrange.ColumnWidth = 3.8;

                        workrange = pastesheet.get_Range("AB" + insertline, "AC" + insertline);    //투입일

                        if (str_DayAndTime.Length > 5)
                        {
                            workrange.Value2 = str_DayAndTime.Substring(0, 5);
                        }
                        else
                        {
                            workrange.Value2 = str_DayAndTime;
                        }

                        workrange.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                        workrange.Font.Size = 10;
                        workrange.ColumnWidth = 3.8;

                        workrange = pastesheet.get_Range("AD" + insertline, "AF" + insertline);    //수주량
                        workrange.Value2 = str_OrderQty;
                        workrange.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                        workrange.Font.Size = 10;
                        workrange.ColumnWidth = 1.7;

                        workrange = pastesheet.get_Range("AG" + insertline, "AI" + insertline);    //투입량
                        workrange.Value2 = str_p1WorkQty;
                        workrange.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                        workrange.Font.Size = 10;
                        workrange.ColumnWidth = 1.2;

                        workrange = pastesheet.get_Range("AJ" + insertline, "AL" + insertline);    //검사량
                        workrange.Value2 = str_InspectQty;
                        workrange.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                        workrange.Font.Size = 10;
                        workrange.ColumnWidth = 1.2;

                        workrange = pastesheet.get_Range("AM" + insertline, "AO" + insertline);    //합격량
                        workrange.Value2 = str_PassQty;
                        workrange.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                        workrange.Font.Size = 10;
                        workrange.ColumnWidth = 1.2;

                        workrange = pastesheet.get_Range("AP" + insertline, "AR" + insertline);    //불합격량
                        workrange.Value2 = str_DefectQty;
                        workrange.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                        workrange.Font.Size = 10;
                        workrange.ColumnWidth = 1.2;

                        workrange = pastesheet.get_Range("AS" + insertline, "AU" + insertline);    //출고량
                        workrange.Value2 = str_OutQty;
                        workrange.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                        workrange.Font.Size = 10;
                        workrange.ColumnWidth = 1.2;

                        DataCount = i;
                        j++;

                        // 합계 누적
                        Sum.OrderSum += ConvertInt(str_OrderQty);
                        Sum.InsertSum += ConvertInt(str_p1WorkQty);

                        Sum.InspectSum += ConvertDouble(str_InspectQty);
                        Sum.PassSum += ConvertDouble(str_PassQty);
                        Sum.DefectSum += ConvertDouble(str_DefectQty);
                        Sum.OutSum += ConvertDouble(str_OutQty);


                    }

                    // 합계 출력
                    int totalLine = 34 + ((Page - 1) * 38);

                    Sum.Count = DataCount + 1;


                    workrange = pastesheet.get_Range("AB" + totalLine, "AC" + totalLine);    // 건수
                    workrange.Value2 = Sum.Count + " 건";
                    workrange.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                    workrange.Font.Size = 10;

                    workrange = pastesheet.get_Range("AD" + totalLine, "AF" + totalLine);    // 총 수주량
                    workrange.Value2 = Sum.OrderSum;
                    workrange.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                    workrange.Font.Size = 10;

                    workrange = pastesheet.get_Range("AG" + totalLine, "AI" + totalLine);    // 총 투입량
                    workrange.Value2 = Sum.InsertSum;
                    workrange.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                    workrange.Font.Size = 10;

                    workrange = pastesheet.get_Range("AJ" + totalLine, "AL" + totalLine);    // 총 검일시
                    workrange.Value2 = Sum.InspectSum;
                    workrange.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                    workrange.Font.Size = 10;

                    workrange = pastesheet.get_Range("AM" + totalLine, "AO" + totalLine);    // 총 통과량
                    workrange.Value2 = Sum.PassSum;
                    workrange.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                    workrange.Font.Size = 10;

                    workrange = pastesheet.get_Range("AP" + totalLine, "AR" + totalLine);    // 총 불합격량
                    workrange.Value2 = Sum.DefectSum;
                    workrange.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                    workrange.Font.Size = 10;

                    workrange = pastesheet.get_Range("AS" + totalLine, "AU" + totalLine);    // 총 출고량
                    workrange.Value2 = Sum.OutSum;
                    workrange.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                    workrange.Font.Size = 10;

                }

                pastesheet.PageSetup.TopMargin = 0;
                pastesheet.PageSetup.BottomMargin = 0;
                //pastesheet.PageSetup.Zoom = 43;

                msg.Hide();

                if (preview_click == true)
                {
                    excelapp.Visible = true;
                    pastesheet.PrintPreview();
                }
                else
                {
                    excelapp.Visible = true;
                    pastesheet.PrintOutEx();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("오류 발생, 오류 내용 : " + ex.ToString());
            }
            finally
            {
                lib2.ReleaseExcelObject(workbook);
                lib2.ReleaseExcelObject(worksheet);
                lib2.ReleaseExcelObject(pastesheet);
                lib2.ReleaseExcelObject(excelapp);
                lib2 = null;
            }
        }

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

        private Double ConvertDouble(string str)
        {
            Double result = 0;
            Double chkDouble = 0;

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

        private void re_Search()
        {
            //검색버튼 비활성화
            btnSearch.IsEnabled = false;

            Dispatcher.BeginInvoke(new Action(() =>
            {
                //로직
                FillGrid();
            }), System.Windows.Threading.DispatcherPriority.Background);

            btnSearch.IsEnabled = true;
        }


        private void DataGrid_SizeChange(object sender, SizeChangedEventArgs e)
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

        // 천자리 콤마, 소수점 버리기
        private string stringFormatN0(object obj)
        {
            return string.Format("{0:N0}", obj);
        }

     
    }

    class Win_ord_OrderClose_U_CodeView : BaseView
    {
        public override string ToString()
        {
            return (this.ReportAllProperties());
        }

        public bool IsCheck { get; set; }
        public string cls { get; set; }
        public bool RowColor { get; set; }
        public string OrderNo { get; set; }
        public string OrderID { get; set; }
        public string CustomID { get; set; }
        public string KCustom { get; set; }

        public string DvlyDate { get; set; }
        public string CloseClss { get; set; }
        public string ChunkRate { get; set; }
        public string LossRate { get; set; }
        public string Article { get; set; }

        public string WorkName { get; set; }

        //public string ArticleID { get; set; }
        public string WorkWidth { get; set; }
        public string OrderQty { get; set; }
        public string UnitClss { get; set; }
        public string InspectQty { get; set; }
        public string PassQty { get; set; }
        public string DefectQty { get; set; }
        public string OutQty { get; set; }
        public string OutDate { get; set; }
        public string ColorQty { get; set; }
        public string BuyerModel { get; set; }
        public string BuyerModelID { get; set; }
        public string BuyerArticleNo { get; set; }
        public string UnitClssName { get; set; }
        public string OrderClss { get; set; }
        public string OrderSpec { get; set; }
        public string InputQty { get; set; }
        public string InputDate { get; set; }
        public string OutQtyYet { get; set; }
        public string AcptDate { get; set; }
        public string RealQty { get; set; }
        public string ExamDate { get; set; }
        public string Remark { get; set; }

        public string p1StartWorkDate { get; set; }
        public string p1StartWorkDTime { get; set; }
        public string p1WorkQty { get; set; }
        public string p1ProcessID { get; set; }
        public string p1ProcessName { get; set; }
        public string p1DayAndTime { get; set; }


        public string p2StartWorkDate { get; set; }
        public string p2StartWorkDTime { get; set; }
        public string p2WorkQty { get; set; }
        public string p2ProcessID { get; set; }
        public string p2ProcessName { get; set; }
        public string p2DayAndTime { get; set; }


        public string p3StartWorkDate { get; set; }
        public string p3StartWorkDTime { get; set; }
        public string p3WorkQty { get; set; }
        public string p3ProcessID { get; set; }
        public string p3ProcessName { get; set; }
        public string p3DayAndTime { get; set; }


        public string p4StartWorkDate { get; set; }
        public string p4StartWorkDTime { get; set; }
        public string p4WorkQty { get; set; }
        public string p4ProcessID { get; set; }
        public string p4ProcessName { get; set; }
        public string p4DayAndTime { get; set; }


        public string p5StartWorkDate { get; set; }
        public string p5StartWorkDTime { get; set; }
        public string p5WorkQty { get; set; }
        public string p5ProcessID { get; set; }
        public string p5ProcessName { get; set; }
        public string p5DayAndTime { get; set; }

        public string p6StartWorkDate { get; set; }
        public string p6StartWorkDTime { get; set; }
        public string p6WorkQty { get; set; }
        public string p6ProcessID { get; set; }
        public string p6ProcessName { get; set; }
        public string p6DayAndTime { get; set; }

        public string p7StartWorkDate { get; set; }
        public string p7StartWorkDTime { get; set; }
        public string p7WorkQty { get; set; }
        public string p7ProcessID { get; set; }
        public string p7ProcessName { get; set; }
        public string p7DayAndTime { get; set; }

        public string p8StartWorkDate { get; set; }
        public string p8StartWorkDTime { get; set; }
        public string p8WorkQty { get; set; }
        public string p8ProcessID { get; set; }
        public string p8ProcessName { get; set; }
        public string p8DayAndTime { get; set; }

        public string p9StartWorkDate { get; set; }
        public string p9StartWorkDTime { get; set; }
        public string p9WorkQty { get; set; }
        public string p9ProcessID { get; set; }
        public string p9ProcessName { get; set; }
        public string p9DayAndTime { get; set; }

        public string p10StartWorkDate { get; set; }
        public string p10StartWorkDTime { get; set; }
        public string p10WorkQty { get; set; }
        public string p10ProcessID { get; set; }
        public string p10ProcessName { get; set; }
        public string p10DayAndTime { get; set; }

        public string DayAndTime { get; set; }
        public string DvlyDateEdit { get; set; }
        public string ProductGrpID { get; set; }
        public string ProductGrpName { get; set; }
        //public string AcptDate { get; set; }
        public double OverAndShort { get; set; }

        
        public string OrderID_CV { get; set; }
        public int Num { get; set; }
    }

    public class dgOrderSum
    {
        public int Count { get; set; }
        public int OrderSum { get; set; }
        public int InsertSum { get; set; }
        public double InspectSum { get; set; }
        public double PassSum { get; set; }
        public double DefectSum { get; set; }
        public double OutSum { get; set; }
        public double OasSum { get; set; }
        public string TextData { get; set; }
        public int ReaQtySum { get; set; }
        public int OutQtySum { get; set; }
        public int OutQtyYetSum { get; set; }
    }
}


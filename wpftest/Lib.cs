using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Net;
using System.Net.Sockets;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Data;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Media.Media3D;
using System.Windows.Threading;
using WPF.MDI;

namespace WizMes_WellMade
{
    public class Lib
    {
        private static Lib mLib = null;
        private static ToolTip currentToolTip;
        private static System.Windows.Threading.DispatcherTimer currentTimer;

        public static Lib Instance
        {
            get
            {
                if (mLib == null)
                {
                    mLib = new Lib();
                }

                return mLib;
            }
        }

        static void DataGridHeaderInTextChanged()
        {

        }


        #region DatePicker 선택

        // DateTime에서 자주쓰임 >> 전월 / 금월 버튼을 클릭하여 전월과 금월의 첫날, 막날, 오늘 등의 자료를 스트링으로 던져줌.
        /// <summary>
        /// 전월 적용 return string
        /// [0] 전월1일, [1] 전월 말일
        /// </summary>
        /// <returns></returns>
        public string[] BringLastMonthDatetime()
        {
            DateTime dFirstDayOfThisMonth = DateTime.Today.AddDays(-(DateTime.Today.Day - 1));  // 이게 이번달 1일.

            DateTime dFirstDayOfLastMonth = dFirstDayOfThisMonth.AddMonths(-1);     // 이게 지난 달 1일
            DateTime dLastDayOfLastMonth = dFirstDayOfThisMonth.AddDays(-1);        // 이게 지난 달 말일

            string[] LastMonthInfo = new string[2];
            LastMonthInfo[0] = dFirstDayOfLastMonth.ToString("yyyy-MM-dd");
            LastMonthInfo[1] = dLastDayOfLastMonth.ToString("yyyy-MM-dd");

            return LastMonthInfo;
        }

        //메서드 실행전에 파라미터가 null일 수 있어서 추가 작성
        public DateTime[] BringLastDayDateTimeContinue(DatePicker fromDatePicker)
        {
            DateTime[] LastDayInfo = new DateTime[2];

            try
            {

                if (fromDatePicker != null && fromDatePicker.SelectedDate != null)
                {

                    DateTime FromDate = fromDatePicker.SelectedDate.Value;

                    DateTime BeforeDay = FromDate.AddDays(-1); // 선택한 일자 전일!

                    LastDayInfo[0] = BeforeDay;
                    LastDayInfo[1] = BeforeDay;

                    return LastDayInfo;

                }
                else
                {
                    DateTime BeforeDay = DateTime.Today.AddDays(-1); // 어제

                    LastDayInfo[0] = BeforeDay;
                    LastDayInfo[1] = BeforeDay;
                }
            }
            catch
            {

            }

            return LastDayInfo;

        }

        /// <summary>
        /// 전월(계속)
        /// </summary>
        /// <param name="FromDate"></param>
        /// <returns></returns>
        public DateTime[] BringLastMonthContinue(DateTime FromDate)
        {
            DateTime[] LastMonthInfo = new DateTime[2];
            if (FromDate != null)
            {
                DateTime ThatMonth1 = FromDate.AddDays(-(FromDate.Day - 1)); // 선택한 일자 달의 1일!

                DateTime LastMonth1 = ThatMonth1.AddMonths(-1); // 저번달 1일
                DateTime LastMonth31 = ThatMonth1.AddDays(-1); // 저번달 말일

                LastMonthInfo[0] = LastMonth1;
                LastMonthInfo[1] = LastMonth31;
            }
            else
            {
                DateTime ThisMonth1 = DateTime.Today.AddDays(-(DateTime.Today.Day - 1)); // 이번달 1일

                DateTime LastMonth1 = ThisMonth1.AddMonths(-1); // 저번달 1일
                DateTime LastMonth31 = ThisMonth1.AddDays(-1); // 저번달 말일

                LastMonthInfo[0] = LastMonth1;
                LastMonthInfo[1] = LastMonth31;
            }


            return LastMonthInfo;
        }

        public DateTime[] BringLastMonthContinue(DatePicker fromDate)
        {
            DateTime[] LastMonthInfo = new DateTime[2];
            try
            {
                if (fromDate != null && fromDate.SelectedDate != null)
                {
                    DateTime FromDate = fromDate.SelectedDate.Value;
                    DateTime ThatMonth1 = FromDate.AddDays(-(FromDate.Day - 1)); // 선택한 일자 달의 1일!

                    DateTime LastMonth1 = ThatMonth1.AddMonths(-1); // 저번달 1일
                    DateTime LastMonth31 = ThatMonth1.AddDays(-1); // 저번달 말일

                    LastMonthInfo[0] = LastMonth1;
                    LastMonthInfo[1] = LastMonth31;
                }
                else
                {
                    DateTime ThisMonth1 = DateTime.Today.AddDays(-(DateTime.Today.Day - 1)); // 이번달 1일

                    DateTime LastMonth1 = ThisMonth1.AddMonths(-1); // 저번달 1일
                    DateTime LastMonth31 = ThisMonth1.AddDays(-1); // 저번달 말일

                    LastMonthInfo[0] = LastMonth1;
                    LastMonthInfo[1] = LastMonth31;
                }
            }
            catch
            {

            }


            return LastMonthInfo;
        }

        /// <summary>
        /// 금월 string 배열 형식
        /// [0]이달 1일 , [1] 현재일자
        /// </summary>
        /// <returns></returns>
        public string[] BringThisMonthDatetime()
        {
            DateTime dFirstDayOfThisMonth = DateTime.Today.AddDays(-(DateTime.Today.Day - 1));  // 이게 이번달 1일.
            DateTime dToday = DateTime.Today;

            string[] ThisMonthInfo = new string[2];
            ThisMonthInfo[0] = dFirstDayOfThisMonth.ToString("yyyy-MM-dd");
            ThisMonthInfo[1] = dToday.ToString("yyyy-MM-dd");

            return ThisMonthInfo;
        }

        /// <summary>
        /// 전월 Datetime 리스트 형식
        /// list[0] 전월1일, list[1] 전월 말일
        /// </summary>
        /// <returns></returns>
        public List<DateTime> BringLastMonthDatetimeList()
        {
            DateTime dFirstDayOfThisMonth = DateTime.Today.AddDays(-(DateTime.Today.Day - 1));  // 이게 이번달 1일.

            DateTime dFirstDayOfLastMonth = dFirstDayOfThisMonth.AddMonths(-1);     // 이게 지난 달 1일
            DateTime dLastDayOfLastMonth = dFirstDayOfThisMonth.AddDays(-1);        // 이게 지난 달 말일

            List<DateTime> ld = new List<DateTime>();
            ld.Add(dFirstDayOfLastMonth);
            ld.Add(dLastDayOfLastMonth);

            return ld;
        }

        /// <summary>
        /// 금월 Datetime 리스트 형식
        /// list[0] 금월1일, list[1] 현재일자
        /// </summary>
        /// <returns></returns>
        public List<DateTime> BringThisMonthDatetimeList()
        {
            DateTime dFirstDayOfThisMonth = DateTime.Today.AddDays(-(DateTime.Today.Day - 1));  // 이게 이번달 1일.
            DateTime dToday = DateTime.Today;

            List<DateTime> ld = new List<DateTime>();
            ld.Add(dFirstDayOfThisMonth);
            ld.Add(dToday);

            return ld;
        }

        /// <summary>
        /// 작년 Datetime 리스트 형식
        /// list[0] 작년1월1일, list[1] 작년말일
        /// </summary>
        /// <returns></returns>
        public List<DateTime> BringLastYearDatetime()
        {
            DateTime dFirstDayOfThisYear = new DateTime(DateTime.Now.Year - 1, 1, 1);
            DateTime dToday = new DateTime(DateTime.Now.Year, 1, 1).AddDays(-1);

            List<DateTime> ld = new List<DateTime>();
            ld.Add(dFirstDayOfThisYear);
            ld.Add(dToday);

            return ld;
        }
        /// <summary>
        /// 이전 년도 (계속)
        /// list[0] 작년1월1일, list[1] 작년말일
        /// </summary>
        /// <returns></returns>
        public List<DateTime> BringLastYearDatetimeContinue(DateTime pickDate)
        {
            DateTime dFirstDayOfThisYear = new DateTime(pickDate.Year - 1, 1, 1);
            DateTime dToday = new DateTime(pickDate.Year, 1, 1).AddDays(-1);

            List<DateTime> ld = new List<DateTime>();
            ld.Add(dFirstDayOfThisYear);
            ld.Add(dToday);

            return ld;
        }

        /// <summary>
        /// 올해 Datetime 리스트 형식
        /// list[0] 올해1월1일, list[1] 현재일자
        /// </summary>
        /// <returns></returns>
        public List<DateTime> BringThisYearDatetimeFormat()
        {
            DateTime dFirstDayOfThisYear = new DateTime(DateTime.Now.Year, 1, 1);
            DateTime dToday = DateTime.Today;

            List<DateTime> ld = new List<DateTime>();
            ld.Add(dFirstDayOfThisYear);
            ld.Add(dToday);

            return ld;
        }

        /// <summary>
        /// 최근 6개월 구하기
        /// [0] 6개월전 달(Month)의 1일,[1] 이달 말일(?)
        /// </summary>
        /// <returns></returns>
        public List<DateTime> BringLastSixMonthDateTimeList()
        {
            DateTime end = new DateTime(DateTime.Now.AddMonths(1).Year, DateTime.Now.AddMonths(1).Month, 1).AddDays(-1);
            DateTime start = end.AddMonths(-6);

            List<DateTime> ld = new List<DateTime>();
            ld.Add(start);
            ld.Add(end);

            return ld;
        }

        /// <summary>
        /// 전일 string 배열 형식
        /// [0] 어제, [1] 오늘
        /// </summary>
        /// <returns></returns>
        public string[] BringYesterdayDatetime()
        {
            DateTime dToday = DateTime.Today;
            DateTime dYesterday = dToday.AddDays(-1);

            string[] YesterdayInfo = new string[2];
            YesterdayInfo[0] = dYesterday.ToString("yyyy-MM-dd");
            YesterdayInfo[1] = dYesterday.ToString("yyyy-MM-dd");

            return YesterdayInfo;
        }

        public DateTime[] BringLastDayDateTimeContinue(DateTime FromDate)
        {
            DateTime[] LastDayInfo = new DateTime[2];
            if (FromDate != null)
            {
                DateTime BeforeDay = FromDate.AddDays(-1); // 선택한 일자 전일!

                LastDayInfo[0] = BeforeDay;
                LastDayInfo[1] = BeforeDay;
            }
            else
            {
                DateTime BeforeDay = DateTime.Today.AddDays(-1); // 어제

                LastDayInfo[0] = BeforeDay;
                LastDayInfo[1] = BeforeDay;
            }

            return LastDayInfo;
        }
        /// <summary>
        /// 금년 string 배열 형식
        /// [0] 올해 1월 1일 , [1] 현재 일자
        /// </summary>
        /// <returns></returns>
        public string[] BringThisYearDatetime()
        {
            DateTime dFirstDayOfThisYear = new DateTime(DateTime.Now.Year, 1, 1);
            DateTime dToday = DateTime.Today;

            string[] ThisYearInfo = new string[2];
            ThisYearInfo[0] = dFirstDayOfThisYear.ToString("yyyy-MM-dd");
            ThisYearInfo[1] = dToday.ToString("yyyy-MM-dd");

            return ThisYearInfo;
        }

        /// <summary>
        /// 새로 하나 추가합니다. _ 이번달의 말일(마지막 일자) 을 가져오기.
        /// 사실상 string return
        /// </summary>
        /// <returns></returns>
        public string[] BringThisMonthLastDatetime()
        {
            DateTime dFirstDayOfThisMonth = DateTime.Today.AddDays(-(DateTime.Today.Day - 1));  // 이게 이번달 1일.
            DateTime dThisMonthEndDate = dFirstDayOfThisMonth.AddMonths(1).AddDays(-1);     // 이게 이번달 말일.  (오늘이 아니라 / 이번달의 마지막 일자)

            string[] ThisMonthLastDayInfo = new string[1];
            ThisMonthLastDayInfo[0] = dThisMonthEndDate.ToString("yyyy-MM-dd");

            return ThisMonthLastDayInfo;
        }

        /// <summary>
        /// 하나 또 추가.ㅠㅠ (2018.06.01) 최근 마지막 반년 구하기 .. 6개월 기간.
        /// [0] 6개월전 달(Month)의 1일,[1] 이달 말일(?)
        /// </summary>
        /// <returns></returns>
        public string[] BringLastSixMonthDateTime()
        {
            DateTime end = new DateTime(DateTime.Now.AddMonths(1).Year, DateTime.Now.AddMonths(1).Month, 1).AddDays(-1);
            DateTime start = end.AddMonths(-6);

            string[] LastSixMonthInfo = new string[2];
            LastSixMonthInfo[0] = start.ToString("yyyy-MM-dd");
            LastSixMonthInfo[1] = end.ToString("yyyy-MM-dd");

            return LastSixMonthInfo;
        }

        #endregion

        #region 숫자 관련

        /// <summary>
        /// 숫자변환(int형만) 가능한 값인지 체크.
        /// </summary>
        /// <param name="value"></param>
        /// <returns></returns>
        public bool IsIntOrAnother(string value)
        {
            int num;
            if (int.TryParse(value, out num) == false) { return false; }
            else { return true; }
        }

        /// <summary>
        /// 숫자변환(int,double) 가능한 값인지 체크
        /// </summary>
        /// <param name="value"></param>
        /// <returns></returns>
        public bool IsNumOrAnother(string value)
        {
            bool flag = true;

            double doNum;
            int inNum;

            if (value != null && !value.Equals(""))
            {
                if (value.Contains("."))
                {
                    if (value.Contains(","))
                    {
                        if (double.TryParse(value.Replace(",", ""), out doNum) == false) { flag = false; }
                    }
                    else
                    {
                        if (double.TryParse(value, out doNum) == false) { flag = false; }
                    }

                }
                else
                {
                    if (value.Contains(","))
                    {
                        if (int.TryParse(value.Replace(",", ""), out inNum) == false) { flag = false; }
                    }
                    else
                    {
                        if (int.TryParse(value, out inNum) == false) { flag = false; }
                    }
                }
            }
            else
            {
                flag = false;
            }

            return flag;
        }

        /// <summary>
        /// 텍스트 박스 숫자만 입력(소수점 가능)
        /// previewTextInput 이벤트 주로 사용
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        public void CheckIsNumeric(TextBox sender, TextCompositionEventArgs e)
        {
            decimal result;
            bool isNegative = e.Text.Equals("-") && sender.SelectionStart == 0 && !sender.Text.Contains("-");
            bool dot = sender.Text.IndexOf(".") < 0 && e.Text.Equals(".") && sender.Text.Length > 0;
            if (!(Decimal.TryParse(e.Text, out result) || dot || isNegative))
            {
                e.Handled = true;
            }
            //한글 입력방지
            InputMethod.SetIsInputMethodEnabled(sender, false);

        }

        /// <summary>
        /// 텍스트 박스 숫자만 입력(정수만)
        /// previewTextInput 이벤트 주로 사용
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        public void CheckIsNumericOnly(TextBox sender, TextCompositionEventArgs e)
        {
            decimal result;
            if (!(Decimal.TryParse(e.Text, out result)))
            {
                e.Handled = true;
            }
        }

        /// <summary>
        /// 텍스트 박스 숫자만 입력(소수점 3자리 제한)
        /// (주의 : 텍스트 박스 소수점 3자리가 미리 입력되어 있으면 
        /// 내용을 소수점 2자리까지 만들어야 입력이 된다.)
        /// previewTextInput 이벤트 주로 사용
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        public void CheckIsNumericbyThree(TextBox sender, TextCompositionEventArgs e)
        {
            bool flag = sender.Text.IndexOf(".") < 0 && e.Text.Equals(".") && sender.Text.Length > 0;
            TextBox textBox = sender;
            Decimal result;
            if (!(Decimal.TryParse(e.Text, out result) | flag))
                e.Handled = true;
            if (!textBox.Text.Contains(".") || textBox.Text.IndexOf(".") + 3 >= textBox.Text.Length)
                return;
            e.Handled = true;
        }

        /// <summary>
        /// 텍스트 박스 숫자만 입력(소수점 가능)
        /// textchanged 이벤트 주로사용
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        public void TextChangedOnlyNumber(object sender, TextChangedEventArgs e)
        {
            TextBox textBox = sender as TextBox;

            if (!IsNumOrAnother(textBox.Text))
            {
                TextChange textChange = e.Changes.ElementAt<TextChange>(0);
                int iAddedLength = textChange.AddedLength;
                int iOffset = textChange.Offset;

                textBox.Text = textBox.Text.Remove(iOffset, iAddedLength);
            }
        }

        /// <summary>
        /// 소수점 3자리까지만 표현
        /// </summary>
        /// <param name="strNumvalue"></param>
        /// <returns></returns>
        public string TextBoxDisplay(string strNumvalue)
        {
            string strValue = string.Empty;

            if (IsNumOrAnother(strNumvalue))
            {
                if (strNumvalue.Contains("."))
                {
                    double douValue = double.Parse(strNumvalue);
                    strValue = string.Format("{0:N3}", strValue);
                }
                else
                {
                    int intValue = int.Parse(strNumvalue);
                    strValue = string.Format("{0:N3}", strValue);   //통일성 위해 그냥 .000 넣음
                }
            }
            else
            {
                MessageBox.Show("값은 숫자로 변환이 가능한 . 이나 숫자만을 포함해야 합니다.");
            }

            return strValue;
        }

        #endregion

        #region 엑셀

        //엑셀로 보여주기 위한 참조 추가 및 변수 선언
        public Microsoft.Office.Interop.Excel.Application excel;
        public Microsoft.Office.Interop.Excel.Workbook workBook;
        public Microsoft.Office.Interop.Excel.Worksheet workSheet;
        public Microsoft.Office.Interop.Excel.Range cellRange;

        #region 일단 주석 처리한 것들

        //데이터그리드에서 히든된 데이터도 추가해서 테이블로 만듬
        //public System.Data.DataTable DataGridToDTinHidden(DataGrid dg)
        //{
        //    DataTable dt = new DataTable();
        //    string[] Fi;
        //    string[] Top = new string[dg.Columns.Count];
        //    int n = 0;

        //    try
        //    {
        //        if (dg != null && dg.Items.Count > 0)
        //        {
        //            foreach (DataGridTextColumn dgtc in dg.Columns)
        //            {                      
        //                dt.Columns.Add(dgtc.Header.ToString().ToUpper(), typeof(string));
        //                Top[n] = dgtc.Header.ToString();
        //                n++;
        //            }
        //            //한 컬럼이 부족한걸 메우기 위한 컬럼 추가
        //            dt.Columns.Add("");

        //            DataRow dr;

        //            for (int i = 0; i < dg.Items.Count; i++)
        //            {
        //                dr = dt.NewRow();
        //                var temp = dg.Items.GetItemAt(i).ExcelAllProperties();
        //                Fi = temp.Split(new char[] { '/' });

        //                for (int j = 0; j < Fi.Length - 1; j++)
        //                {
        //                    if (Fi[j].Contains(":"))
        //                        dr[Top[j]] = Fi[j].Substring(Fi[j].IndexOf(":") + 1).Trim();
        //                    else
        //                        dr[Top[j]] = Fi[j];
        //                }
        //                dt.Rows.Add(dr);
        //            }
        //        }
        //        else
        //        {
        //            dt.Columns.Add("none");
        //        }
        //    }
        //    catch (Exception e)
        //    {
        //        MessageBox.Show(e.Message);
        //    }

        //    return dt;
        //}

        //DataGrid를 DataTable로 바꿔주기 위해 만듬(ItemSource 로 받아오면 null 일때 사용)
        //public System.Data.DataTable DataGirdToDataTable(DataGrid dg)
        //{
        //    if (dg != null && dg.Items.Count>0)
        //    {
        //        dg.SelectAllCells();
        //        //dg.SelectAll();
        //        dg.ClipboardCopyMode = DataGridClipboardCopyMode.IncludeHeader;
        //        ApplicationCommands.Copy.Execute(null, dg);
        //        dg.UnselectAllCells();
        //        //dg.UnselectAll();

        //        string result = (string)Clipboard.GetData(DataFormats.CommaSeparatedValue);
        //        string[] Lines = result.Split(new string[] { "\r\n", "\n" }, StringSplitOptions.None);
        //        string[] Fields;
        //        Fields = Lines[0].Split(new char[] { ',' });

        //        int Cols = Fields.GetLength(0);
        //        System.Data.DataTable dt = new System.Data.DataTable();

        //        for (int i = 0; i < Cols; i++)
        //        {
        //            dt.Columns.Add(Fields[i].ToUpper(), typeof(string));
        //        }
        //        dt.Columns.Add("");

        //        DataRow dr;

        //        for (int i = 0; i < Lines.GetLength(0) - 1; i++)
        //        {
        //            Fields = Lines[i].Split(new char[] { ',' });
        //            dr = dt.NewRow();
        //            for (int j = 0; j < Cols; j++)
        //            {
        //                dr[j] = Fields[j];
        //            }
        //            dt.Rows.Add(dr);
        //        }

        //        return dt;
        //    }
        //    else
        //    {
        //        System.Data.DataTable dt = new System.Data.DataTable();
        //        dt.Columns.Add("none");
        //        //DataRow dr = dt.NewRow();
        //        //dr["none"] = "none";

        //        return dt;
        //    }
        //}


        #endregion

        /// <summary>
        /// 현재 작업내용 없음
        /// </summary>
        /// <param name="listBox"></param>
        /// <param name="Header"></param>
        /// <returns></returns>
        public DataTable ListToDataTable(ListBox listBox, List<string> Header)
        {
            DataTable dt = new DataTable();

            if (listBox != null)
            {
                foreach (string str in Header)
                {
                    dt.Columns.Add(str);
                }
                dt.Columns.Add("");
            }

            return dt;
        }

        /// <summary>
        /// 현재 작업내용 없음
        /// </summary>
        /// <param name="listbox"></param>
        /// <returns></returns>
        public DataTable ListBoxTotoDT(ListBox listbox)
        {
            DataTable dt = new DataTable();

            if (listbox != null)
            {
                StringBuilder sblistbox = new StringBuilder();

                List<string> listColumns = new List<string>();
                List<ListBoxItem> listBoxItem = new List<ListBoxItem>();
                List<string> listHeader = new List<string>();

                //int rowCount = 0;
                //int colCount = 0;

                //try
                //{

                //}
                //catch (Exception ex)
                //{
                //}
            }

            return dt;
        }

        /// <summary>
        /// DataGrid의 내용을 DataTable로 추출
        /// Hidden(숨겨진) Column의 내용포함
        /// </summary>
        /// <param name="dg"></param>
        /// <returns></returns>
        public System.Data.DataTable DataGridToDTinHidden(DataGrid dg)
        {
            DataTable dt = new DataTable();

            if (dg != null)
            {
                StringBuilder sbGridData = new StringBuilder();

                List<string> listColumns = new List<string>();
                List<DataGridColumn> listAllDataGridColumns = new List<DataGridColumn>();
                List<string> listHeader = new List<string>();

                int rowCount = 0;
                int colCount = 0;

                try
                {
                    if (dg.HeadersVisibility == DataGridHeadersVisibility.None || dg.HeadersVisibility == DataGridHeadersVisibility.Column || dg.HeadersVisibility == DataGridHeadersVisibility.All || dg.HeadersVisibility == DataGridHeadersVisibility.Row)
                    {
                        foreach (DataGridColumn dataGridColumn in dg.Columns)
                        {
                            listAllDataGridColumns.Add(dataGridColumn);
                            if (dataGridColumn.Header != null)
                            {
                                listHeader.Add(dataGridColumn.Header.ToString());
                            }
                            else  //header가 없는 경우 빈값을 줘야 열이 맞다.
                            {
                                listHeader.Add("");
                            }
                            dt.Columns.Add(listHeader[colCount]);
                            colCount++;
                        }

                        //마지막열 헤더를 보이게 하기위해 추가_없으면 이상하게 안나옴 ㅡ,.ㅡ
                        dt.Columns.Add("");

                        if (dg.ItemsSource != null)
                        {
                            foreach (object data in dg.ItemsSource)
                            {
                                listColumns.Clear();
                                colCount = 0;
                                rowCount++;
                                DataRow dr = dt.NewRow();
                                foreach (DataGridColumn dataGridColumn in listAllDataGridColumns)
                                {
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
                                            }

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
                                    }

                                    listColumns.Add(strValue);
                                    dr[colCount] = strValue;
                                    colCount++;
                                }
                                dt.Rows.Add(dr);
                            }
                        }
                        else
                        {
                            foreach (object data in dg.Items)
                            {
                                listColumns.Clear();
                                colCount = 0;
                                rowCount++;
                                DataRow dr = dt.NewRow();
                                foreach (DataGridColumn dataGridColumn in listAllDataGridColumns)
                                {
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
                                            }

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
                                    }

                                    listColumns.Add(strValue);
                                    dr[colCount] = strValue;
                                    colCount++;
                                }
                                dt.Rows.Add(dr);
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                }
            }

            return dt;
        }

        /// <summary>
        /// DataGrid의 내용을 DataTable로 추출
        /// Hidden(숨겨진) Column의 내용제외
        /// </summary>
        /// <param name="dg"></param>
        /// <returns></returns>
        public System.Data.DataTable DataGirdToDataTable(DataGrid dg)
        {
            DataTable dt = new DataTable();

            if (dg != null)
            {
                StringBuilder sbGridData = new StringBuilder();

                List<string> listColumns = new List<string>();
                List<DataGridColumn> listVisibleDataGridColumns = new List<DataGridColumn>();
                List<string> listHeader = new List<string>();

                int rowCount = 0;
                int colCount = 0;

                try
                {
                    if (dg.HeadersVisibility == DataGridHeadersVisibility.None || (dg.HeadersVisibility == DataGridHeadersVisibility.Column || dg.HeadersVisibility == DataGridHeadersVisibility.All || dg.HeadersVisibility == DataGridHeadersVisibility.Row))
                    {
                        foreach (DataGridColumn dataGridColumn in dg.Columns.Where(dataGridColumn => dataGridColumn.Visibility == Visibility.Visible))
                        {
                            listVisibleDataGridColumns.Add(dataGridColumn);
                            if (dataGridColumn.Header != null)
                            {
                                listHeader.Add(dataGridColumn.Header.ToString());
                            }

                            dt.Columns.Add(listHeader[colCount]);
                            colCount++;

                        }

                        //마지막열 헤더를 보이게 하기위해 추가_없으면 이상하게 안나옴 ㅡ,.ㅡ
                        dt.Columns.Add("");

                        foreach (object data in dg.Items)
                        {
                            //MessageBox.Show("" + data.ToString());


                            listColumns.Clear();
                            colCount = 0;
                            rowCount++;
                            DataRow dr = dt.NewRow();
                            foreach (DataGridColumn dataGridColumn in listVisibleDataGridColumns)
                            {
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
                                            else if (frameworkElement is DockPanel)
                                            {
                                                BindingExpression bindingExpression = frameworkElement.GetBindingExpression(DockPanel.ToolTipProperty);
                                                if (bindingExpression != null)
                                                {
                                                    objBinding = bindingExpression.ParentBinding;
                                                }
                                            }
                                        }

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
                                }

                                listColumns.Add(strValue);
                                dr[colCount] = strValue;
                                colCount++;
                            }
                            dt.Rows.Add(dr);
                        }
                    }
                }
                catch (Exception ex)
                {
                    // 테스트
                    MessageBox.Show("에러 내용 : " + ex.Message);
                }
            }

            return dt;
        }

        /// <summary>
        /// 엑셀로 행열을 맞춰 넣어둔다
        /// </summary>
        /// <param name="dt"></param>
        /// <param name="Name"></param>
        /// <returns></returns>
        public bool GenerateExcel(System.Data.DataTable dt, string Name)
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
                    cellRange = workSheet.Range[workSheet.Cells[1, 1], workSheet.Cells[rowCount + 1, datatable.Columns.Count]];
                    cellRange.EntireColumn.AutoFit();

                    ReleaseExcelObject(workSheet);
                    ReleaseExcelObject(workBook);
                    //ReleaseExcelObject(excel);
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

        /// <summary>
        /// 타이틀, 검색조건 추가한 엑셀 형식
        /// </summary>
        /// <param name="dt"></param>
        /// <param name="Name"></param>
        /// <returns></returns>
        public bool TitleAddGenerateExcel(System.Data.DataTable dt, string[] Name)
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
                    workSheet.Name = Name[0];
                    workSheet.Cells.Font.Size = 11;

                    System.Data.DataTable datatable = dt;

                    //타이틀
                    cellRange = workSheet.Range[workSheet.Cells[1, 1], workSheet.Cells[1, datatable.Columns.Count]];
                    cellRange.Merge(true);
                    cellRange.Font.Size = 18;
                    cellRange.Font.Bold = true;
                    cellRange.Font.Underline = true;
                    cellRange.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;

                    workSheet.Cells[1, 1] = Name[0];

                    //조건들
                    int rowCount = 2;

                    for (int i = 1; i < Name.Length; i++)
                    {
                        if (Name[i] != null && !Name[i].Equals(""))
                        {
                            workSheet.Cells[rowCount, 1] = Name[i];
                            cellRange = workSheet.Range[workSheet.Cells[rowCount, 1], workSheet.Cells[1, datatable.Columns.Count]];
                            cellRange.Merge(true);
                            rowCount += 1;
                        }
                    }

                    //칼럼헤더
                    for (int i = 1; i < datatable.Columns.Count; i++)
                    {
                        workSheet.Cells[rowCount, i] = datatable.Columns[i - 1].ColumnName;
                    }

                    //데이터그리드
                    foreach (DataRow row in datatable.Rows)
                    {
                        rowCount += 1;
                        for (int i = 0; i < datatable.Columns.Count; i++)
                        {
                            workSheet.Cells[rowCount, i + 1] = row[i].ToString();
                        }
                    }
                    cellRange = workSheet.Range[workSheet.Cells[2, 1], workSheet.Cells[rowCount + 1, datatable.Columns.Count]];
                    cellRange.EntireColumn.AutoFit();

                    ReleaseExcelObject(workSheet);
                    ReleaseExcelObject(workBook);
                    //ReleaseExcelObject(excel);
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

        /// <summary>
        /// 로우헤더, 칼럼헤더가 있는 엑셀 형식
        /// </summary>
        /// <param name="dt"></param>
        /// <param name="Name"></param>
        /// <param name="excelRowHeaderName"></param>
        /// <returns></returns>
        public bool HeaderAddGenerateExcel(System.Data.DataTable dt, string Name, string[] excelRowHeaderName)
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

                    //Row헤더영역 
                    for (int i = 2; i < datatable.Rows.Count + 2; i++)
                    {
                        workSheet.Cells[i, 1] = excelRowHeaderName[i - 2];
                    }

                    //Colums헤더영역
                    for (int i = 1; i < datatable.Columns.Count; i++)
                    {
                        workSheet.Cells[1, i + 1] = datatable.Columns[i - 1].ColumnName;
                    }

                    //데이터영역
                    foreach (DataRow row in datatable.Rows)
                    {
                        rowCount += 1;
                        for (int i = 1; i < datatable.Columns.Count + 1; i++)
                        {
                            workSheet.Cells[rowCount, i + 1] = row[i - 1].ToString();
                        }
                    }
                    cellRange = workSheet.Range[workSheet.Cells[1, 1], workSheet.Cells[rowCount, datatable.Columns.Count]];
                    cellRange.EntireColumn.AutoFit();

                    ReleaseExcelObject(workSheet);
                    ReleaseExcelObject(workBook);
                    //ReleaseExcelObject(excel);
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

        /// <summary>
        /// 타이틀, 검색조건, 로우헤더, 칼럼헤더가 있는 엑셀 형식
        /// </summary>
        /// <param name="dt"></param>
        /// <param name="Name"></param>
        /// <param name="excelRowHeaderName"></param>
        /// <returns></returns>
        public bool HeaderAddNewGenerateExcel(System.Data.DataTable dt, string[] Name, string[] excelRowHeaderName)
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
                    workSheet.Name = Name[0];

                    System.Data.DataTable datatable = dt;
                    workSheet.Cells.Font.Size = 11;

                    //타이틀
                    cellRange = workSheet.Range[workSheet.Cells[1, 1], workSheet.Cells[1, datatable.Columns.Count]];
                    cellRange.Merge(true);
                    cellRange.Font.Size = 18;
                    cellRange.Font.Bold = true;
                    cellRange.Font.Underline = true;
                    cellRange.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;

                    workSheet.Cells[1, 1] = Name[0];

                    //조건들
                    int rowCount = 2;

                    for (int i = 1; i < Name.Length; i++)
                    {
                        if (Name[i] != null && !Name[i].Equals(""))
                        {
                            workSheet.Cells[rowCount, 1] = Name[i];
                            cellRange = workSheet.Range[workSheet.Cells[rowCount, 1], workSheet.Cells[1, datatable.Columns.Count]];
                            cellRange.Merge(true);
                            rowCount += 1;
                        }
                    }

                    //Row헤더영역 
                    for (int i = rowCount; i < datatable.Rows.Count + 2; i++)
                    {
                        workSheet.Cells[i, 1] = excelRowHeaderName[i - 2];
                    }

                    //Colums헤더영역
                    for (int i = 1; i < datatable.Columns.Count; i++)
                    {
                        workSheet.Cells[rowCount, i + 1] = datatable.Columns[i - 1].ColumnName;
                    }

                    //데이터영역
                    foreach (DataRow row in datatable.Rows)
                    {
                        rowCount += 1;
                        for (int i = 1; i < datatable.Columns.Count + 1; i++)
                        {
                            workSheet.Cells[rowCount, i + 1] = row[i - 1].ToString();
                        }
                    }
                    cellRange = workSheet.Range[workSheet.Cells[1, 1], workSheet.Cells[rowCount, datatable.Columns.Count]];
                    cellRange.EntireColumn.AutoFit();

                    ReleaseExcelObject(workSheet);
                    ReleaseExcelObject(workBook);
                    //ReleaseExcelObject(excel);
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

        #endregion

        #region DateTime

        /// <summary>
        /// 8 length string, yyyy,MM,dd string return
        /// </summary>
        public string StrDateTime(string strDate)
        {
            string date = string.Empty;

            if (strDate.Length == 8)
            {
                string yyyy = strDate.Substring(0, 4);
                string MM = strDate.Substring(4, 2);
                string dd = strDate.Substring(6, 2);

                date = yyyy + "," + MM + "," + dd;
            }

            return date;
        }

        /// <summary>
        /// 8 length string, yyyy-MM-dd string return
        /// </summary>
        public string StrDateTimeBar(string strDate)
        {
            string date = string.Empty;

            if (strDate.Length == 8)
            {
                string yyyy = strDate.Substring(0, 4);
                string MM = strDate.Substring(4, 2);
                string dd = strDate.Substring(6, 2);

                date = yyyy + "-" + MM + "-" + dd;
            }

            return date;
        }

        /// <summary>
        /// yyyy-MM-dd string or yyyy,MM,dd string, yyyyMMdd string return 
        /// </summary>
        public string DateFormat(string strDate)
        {
            string date = string.Empty;

            date = strDate.ToString().Substring(0, 10).Replace("-", "");

            if (date.Length > 8)
            {
                date = date.Substring(0, 10).Replace(",", "");
            }

            return date;
        }

        /// <summary>
        /// yyyyMMdd string convert yy-MM-dd, string yy-MM-dd return
        /// </summary>
        public string SixLengthDate(string strDate)
        {
            string date = string.Empty;

            date = strDate.Substring(2, 2) + "-" + strDate.Substring(4, 2) + "-" + strDate.Substring(6, 2);

            return date;
        }

        /// <summary>
        /// yyyyMMdd stirng, datetime convert return
        /// </summary>
        public DateTime strConvertDate(string strDate)
        {
            DateTime dtDate = DateTime.ParseExact(strDate, "yyyyMMdd", System.Globalization.CultureInfo.InvariantCulture);
            return dtDate;
        }

        /// <summary>
        /// yyyyMMdd or yyyy-MM-dd 형식을 yyyy.MM.dd 로 return
        /// </summary>
        public string StrDateTimeDot(string strDate)
        {
            string returnString = string.Empty;

            if (strDate is null)
            {
                return returnString;
            }
            else
            {
                if (strDate.Contains("-") && strDate.Length == 10)
                {
                    returnString = strDate.Substring(0, 4) + "." + strDate.Substring(5, 2)
                        + "." + strDate.Substring(8, 2);
                }
                else if (strDate.Length == 8)
                {
                    returnString = strDate.Substring(0, 4) + "." + strDate.Substring(4, 2)
                        + "." + strDate.Substring(6, 2);
                }

                return returnString;
            }
        }
        /// <summary>
        /// yyyyMMdd 형식을 yyyy-MM-dd 로 return
        /// </summary>
        public string StrDateTimeToSlash(string strDate)
        {
            string returnString = string.Empty;

            if (strDate is null || strDate.Contains("-") || strDate.Contains(".") || strDate.Contains("/"))
            {
                return returnString;
            }
            else
            {
                if (strDate.Length == 8)
                {
                    returnString = strDate.Substring(0, 4) + "/" + strDate.Substring(4, 2)
                        + "." + strDate.Substring(6, 2);
                }
                else if (strDate.Length == 4)
                {
                    returnString = strDate.Substring(0, 2) + "/" + strDate.Substring(2, 2);
                }

                return returnString;
            }
        }
        /// <summary>
        /// HHmmss string, HH:mm:ss string return
        /// </summary>
        public string SixLengthTime(string Time)
        {
            string sixlengthTime = string.Empty;

            if (Time.Length == 6)
            {
                string HH = Time.Substring(0, 2);
                string mm = Time.Substring(2, 2);
                string ss = Time.Substring(4, 2);

                sixlengthTime = HH + ":" + mm + ":" + ss;
            }

            return sixlengthTime;
        }

        #endregion

        #region 문자열 자르기 모음

        /// <summary>
        /// 문자열 왼쪽편처음부터 지정된 문자열값 리턴(VBScript Left기능)
        /// </summary>
        /// <param name="target">얻을 문자열</param>
        /// <param name="length">얻을 문자열길이</param>
        /// <returns>얻은 문자열 값</returns>
        public string Left(string target, int length)
        {
            if (length <= target.Length)
            {
                return target.Substring(0, length);
            }
            return target;
        }

        /// <summary>
        /// 지정된 위치이후 모든 문자열 리턴 (VBScript Mid기능)
        /// </summary>
        /// <param name="target">얻을 문자열</param>
        /// <param name="start">얻을 시작위치</param>
        /// <returns>지정된 위치 이후 모든 문자열리턴</returns>
        public string Mid(string target, int start)
        {
            if (start <= target.Length)
            {
                return target.Substring(start - 1);
            }
            return string.Empty;
        }

        /// <summary>
        /// 문자열이 지정된 위치에서 지정된 길이만큼까지의 문자열 리턴 (VBScript Mid기능)
        /// </summary>
        /// <param name="target">얻을 문자열</param>
        /// <param name="start">얻을 시작위치</param>
        /// <param name="length">얻을 문자열길이</param>
        /// <returns>지정된 길이만큼의 문자열 리턴</returns>
        public string Mid(string target, int start, int length)
        {
            if (start <= target.Length)
            {
                if (start + length - 1 <= target.Length)
                {
                    return target.Substring(start - 1, length);
                }
                return target.Substring(start - 1);
            }
            return string.Empty;
        }

        /// <summary>
        /// 문자열 오른쪽편처음부터 지정된 문자열값 리턴(VBScript Right기능) 
        /// </summary>
        /// <param name="target">얻을 문자열</param>
        /// <param name="length">얻을 문자열길이</param>
        /// <returns>얻은 문자열 값</returns>
        public string Right(string target, int length)
        {
            if (length <= target.Length)
            {
                return target.Substring(target.Length - length);
            }
            return target;
        }

        /// <summary>
        /// 문자열에 .이 포함되어 있을때 그 이하 인덱스 자른 후 리턴
        /// </summary>
        /// <param name="target"></param>
        /// <returns></returns>
        public string SubString(string target)
        {
            if (target != null && !target.Equals(""))
            {
                if (target.Contains("."))
                {
                    return target.Substring(0, target.IndexOf("."));
                }
            }
            return "0";
        }

        /// <summary>
        /// ****-**-**** 의 형태를 ********** 로 바꿔준다.
        /// </summary>
        /// <param name="OrderID"></param>
        /// <returns></returns>
        public string OrderID(string OrderID)
        {
            StringBuilder sb = new StringBuilder();
            sb.Append(OrderID.Substring(0, 4));
            sb.Append(OrderID.Substring(5, 2));
            sb.Append(OrderID.Substring(8, 4));

            return sb.ToString();
        }

        #endregion

        #region DataGridCell 자료 가져올때 사용

        //public Dictionary<string, string> ListData(Dictionary<string,object> dictionary)
        //{
        //}

        /// <summary>
        /// 셀 내부 Content 접근
        /// </summary>
        /// <param name="row"></param>
        /// <param name="column"></param>
        /// <param name="GridView1"></param>
        /// <returns></returns>
        public DataGridCell GetCell(int row, int column, DataGrid GridView1)
        {
            DataGridRow rowContainer = GetRow(row, GridView1);
            if (rowContainer != null)
            {
                DataGridCellsPresenter presenter = GetVisualChild<DataGridCellsPresenter>(rowContainer);

                if (presenter == null)
                {
                    GridView1.ScrollIntoView(rowContainer, GridView1.Columns[column]);
                    presenter = GetVisualChild<DataGridCellsPresenter>(rowContainer);
                }

                DataGridCell cell = (DataGridCell)presenter.ItemContainerGenerator.ContainerFromIndex(column);
                return cell;
            }
            return null;
        }

        /// <summary>
        /// DataGrid의 row 데이터 접근
        /// </summary>
        /// <param name="index"></param>
        /// <param name="GridView1"></param>
        /// <returns></returns>
        public DataGridRow GetRow(int index, DataGrid GridView1)
        {
            DataGridRow row = null;
            int retryCount = 0;

            while (row == null && retryCount < 3)
            {
                row = Application.Current.Dispatcher.Invoke<DataGridRow>(() =>
                {
                    DataGridRow r = (DataGridRow)GridView1.ItemContainerGenerator.ContainerFromIndex(index);
                    if (r == null)
                    {
                        GridView1.UpdateLayout();
                        GridView1.ScrollIntoView(GridView1.Items[index]);
                        r = (DataGridRow)GridView1.ItemContainerGenerator.ContainerFromIndex(index);
                    }
                    return r;
                }, DispatcherPriority.SystemIdle);

                if (row == null)
                {
                    retryCount++;
                    Thread.Sleep(100); // 쓰레드 100ms 대기 후 다시 시도
                }
            }

            return row;

            #region 원본코드
            ////DataGridRow row = (DataGridRow)GridView1.ItemContainerGenerator.ContainerFromIndex(index);

            ////if (row == null)
            ////{
            ////    GridView1.UpdateLayout();
            ////    GridView1.ScrollIntoView(GridView1.Items[index]);
            ////    row = (DataGridRow)GridView1.ItemContainerGenerator.ContainerFromIndex(index);
            ////}
            ////return row;
            #endregion


        }

        /// <summary>
        /// 의존적 컨트롤의 하위 Control 정보 접근
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="parent"></param>
        /// <returns></returns>
        public T GetVisualChild<T>(Visual parent) where T : Visual
        {
            T child = default(T);
            int numVisuals = VisualTreeHelper.GetChildrenCount(parent);
            for (int i = 0; i < numVisuals; i++)
            {
                Visual v = (Visual)VisualTreeHelper.GetChild(parent, i);
                child = v as T;
                if (child == null)
                {
                    child = GetVisualChild<T>(v);
                }
                if (child != null)
                {
                    break;
                }
            }
            return child;
        }

        /// <summary>
        /// 의존적 컨트롤의 부모 컨트롤 가져오기
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="d"></param>
        /// <returns></returns>
        public T GetParent<T>(DependencyObject d) where T : class
        {
            while (d != null && !(d is T))
            {
                d = VisualTreeHelper.GetParent(d);
            }
            return d as T;
        }

        #endregion

        #region FindVisual 여러개

        /// <summary>
        /// 자신의 컨트롤 한단계 아래부터 사용자정의 이름으로 검색합니다.
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="_Control"></param>
        /// <param name="_FindControlName"></param>
        /// <returns></returns>
        public T FindVisualChildByName<T>(DependencyObject _Control, string _FindControlName) where T : DependencyObject
        {

            for (int i = 0; i < VisualTreeHelper.GetChildrenCount(_Control); i++)
            {
                var child = VisualTreeHelper.GetChild(_Control, i);
                string controlName = child.GetValue(Control.NameProperty) as string;

                if (controlName == _FindControlName)
                {
                    return child as T;
                }
                else
                {
                    T result = FindVisualChildByName<T>(child, _FindControlName);
                    if (result != null)
                    {
                        return result;
                    }
                }

            }
            return null;
        }
        /// <summary>
        /// 자신의 컨트롤 한단계 아래부터 사용자정의 이름으로 검색합니다.
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="_Control"></param>
        /// <param name="_FindControlName"></param>
        /// <returns></returns>
        public List<T> FindVisualChildByContainName<T>(DependencyObject _Control, string _FindControlName) where T : DependencyObject
        {
            List<T> List_Con = new List<T>();

            for (int i = 0; i < VisualTreeHelper.GetChildrenCount(_Control); i++)
            {
                var child = VisualTreeHelper.GetChild(_Control, i);
                string controlName = child.GetValue(Control.NameProperty) as string;

                if (controlName.Contains(_FindControlName))
                {
                    List_Con.Add(child as T);
                }
            }
            return List_Con;
        }

        //일단 추가하고 LIB로 옮길 예정
        public void FindChildGroup<T>(DependencyObject parent, string childName, ref List<T> list) where T : DependencyObject
        {
            // Checks should be made, but preferably one time before calling.
            // And here it is assumed that the programmer has taken into
            // account all of these conditions and checks are not needed.
            //if ((parent == null) || (childName == null) || (<Type T is not inheritable from FrameworkElement>))
            //{
            //    return;
            //}

            int childrenCount = VisualTreeHelper.GetChildrenCount(parent);

            for (int i = 0; i < childrenCount; i++)
            {
                // Get the child
                var child = VisualTreeHelper.GetChild(parent, i);

                // Compare on conformity the type
                T child_Test = child as T;

                // Not compare - go next
                if (child_Test == null)
                {
                    // Go the deep
                    FindChildGroup<T>(child, childName, ref list);
                }
                else
                {
                    // If match, then check the name of the item
                    FrameworkElement child_Element = child_Test as FrameworkElement;

                    if (child_Element.Name == childName)
                    {
                        // Found
                        list.Add(child_Test);
                    }

                    // We are looking for further, perhaps there are
                    // children with the same name
                    FindChildGroup<T>(child, childName, ref list);
                }
            }

            return;
        }

        public T FindVisualChild<T>(DependencyObject obj) where T : DependencyObject
        {
            for (int i = 0; i < VisualTreeHelper.GetChildrenCount(obj); i++)
            {
                DependencyObject child = VisualTreeHelper.GetChild(obj, i);
                if (child != null && child is T)
                    return (T)child;
                else
                {
                    T childOfChild = FindVisualChild<T>(child);
                    if (childOfChild != null)
                        return childOfChild;
                }
            }
            return null;
        }

        //데이터 그리드 셀 인포 를 >> 데이터 그리드 셀로 형 변환해서 가져오기.
        //2018.08.02 허윤구.
        public DataGridCell GetDataGridCell(DataGridCellInfo cellInfo)
        {
            var cellContent = cellInfo.Column.GetCellContent(cellInfo.Item);
            if (cellContent != null)
                return (DataGridCell)cellContent.Parent;

            return null;
        }

        public static IEnumerable<T> FindVisualChildren<T>(DependencyObject depObj) where T : DependencyObject
        {
            if (depObj != null)
            {
                for (int i = 0; i < VisualTreeHelper.GetChildrenCount(depObj); i++)
                {
                    DependencyObject child = VisualTreeHelper.GetChild(depObj, i);
                    if (child != null && child is T)
                    {
                        yield return (T)child;
                    }

                    foreach (T childOfChild in FindVisualChildren<T>(child))
                    {
                        yield return childOfChild;
                    }
                }
            }
        }

        #endregion

        #region 레지스트리

        /// <summary>
        /// LogID 레지스트리 등록
        /// </summary>
        /// <param name="strLogID"></param>
        public void SetLogResitry(string strLogID)
        {
            using (var root = RegistryKey.OpenBaseKey(RegistryHive.LocalMachine, RegistryView.Registry64))
            {
                using (RegistryKey rk = root.OpenSubKey(@"SoftWare\\WPF_LogIDKey", true))
                {
                    if (rk == null)
                    {
                        //rk = Registry.LocalMachine.CreateSubKey(@"SoftWare\\WPF_LogIDKey");
                    }
                    else
                    {
                        rk.SetValue("LogName", strLogID);
                    }
                };
            };

            //string regSubkey = "SoftWare\\WPF_LogIDKey";
            //RegistryKey rk = Registry.LocalMachine.OpenSubKey(regSubkey, true);

            //if (rk == null)
            //{
            //   rk = Registry.LocalMachine.CreateSubKey(regSubkey);
            //}
        }

        /// <summary>
        /// LogID 레지스트리 값 가져오기
        /// </summary>
        /// <returns></returns>
        public string GetLogResitry()
        {
            using (var root = RegistryKey.OpenBaseKey(RegistryHive.LocalMachine, RegistryView.Registry64))
            {
                using (RegistryKey rk = root.OpenSubKey(@"SoftWare\\WPF_LogIDKey", true))
                {
                    if (rk == null)
                    {
                        //rk = Registry.LocalMachine.CreateSubKey(@"SoftWare\\WPF_LogIDKey");
                    }
                    else
                    {
                        string getRegLogID = rk.GetValue("LogName") as string;
                        return getRegLogID;
                    }
                };
            };

            return "";

            //string regSubkey = "SoftWare\\WPF_LogIDKey";
            //RegistryKey rk = Registry.LocalMachine.OpenSubKey(regSubkey, true);
            //if (rk != null)
            //{
            //    string getRegLogID = rk.GetValue("LogName") as string;
            //    return getRegLogID;
            //}
            //return "";
        }

        /// <summary>
        /// Log 레지스트리 값 삭제
        /// </summary>
        public void DelLogResitry()
        {
            string regSubkey = "SoftWare\\WPF_LogIDKey";
            Registry.LocalMachine.DeleteSubKey(regSubkey);
        }

        #endregion

        #region 화면의 버튼 IsEnable Change

        /// <summary>
        /// 화면 로드시 검색, 추가,수정,삭제, 인쇄의 권한별 사용여부를 판별하여 enable이 적용된다.
        /// </summary>
        /// <param name="sender"></param>
        public void UiLoading(object sender)
        {
            MenuViewModel menuView = MainWindow.MainMdiContainer.ActiveMdiChild.Tag as MenuViewModel;
            UserControl userControl = (sender as UserControl);



            object objAdd = null;
            object objDelete = null;
            object objSearch = null;
            object objUpdate = null;
            object objPrint = null;

            if (userControl != null)
            {
                objAdd = userControl.FindName("btnAdd");

                if (objAdd != null)
                {
                    if (menuView.AddNewClss.Equals("*"))
                    {
                        (objAdd as Button).IsEnabled = true;
                    }
                    else
                    {
                        (objAdd as Button).IsEnabled = false;
                    }
                }

                objDelete = userControl.FindName("btnDelete");

                if (objDelete != null)
                {
                    if (menuView.DeleteClss.Equals("*"))
                    {
                        (objDelete as Button).IsEnabled = true;
                    }
                    else
                    {
                        (objDelete as Button).IsEnabled = false;
                    }
                }

                objSearch = userControl.FindName("btnSearch");

                if (objSearch != null)
                {
                    if (menuView.SelectClss.Equals("*"))
                    {
                        (objSearch as Button).IsEnabled = true;
                    }
                    else
                    {
                        (objSearch as Button).IsEnabled = false;
                    }
                }

                objUpdate = userControl.FindName("btnUpdate");

                if (objUpdate != null)
                {
                    if (menuView.UpdateClss.Equals("*"))
                    {
                        (objUpdate as Button).IsEnabled = true;
                    }
                    else
                    {
                        (objUpdate as Button).IsEnabled = false;
                    }
                }

                objPrint = userControl.FindName("btnPrint");

                if (objPrint != null)
                {
                    if (menuView.PrintClss.Equals("*"))
                    {
                        (objPrint as Button).IsEnabled = true;
                    }
                    else
                    {
                        (objPrint as Button).IsEnabled = false;
                    }
                }
            }
        }

        public void UiLoadMakeEvent(object sender)
        {
            MenuViewModel menuView = MainWindow.MainMdiContainer.ActiveMdiChild.Tag as MenuViewModel;
            UserControl userControl = (sender as UserControl);

            var dowdow = Window.GetWindow(userControl);
            dowdow.KeyDown += HandleKeyPress;
        }

        public void HandleKeyPress(object sender, KeyEventArgs e)
        {
            Window dowdow = (sender as Window);
            MdiChild userControl = MainWindow.MainMdiContainer.ActiveMdiChild;
            UserControl useruser = userControl.Content as UserControl;

            object objAdd = null;
            object objDelete = null;
            object objSearch = null;
            object objUpdate = null;
            object objPrint = null;
            object objExcel = null;
            object objSave = null;
            object objCancel = null;

            if (useruser != null)
            {
                e.Handled = true;
                if (Keyboard.Modifiers == ModifierKeys.Control && e.Key == Key.F)
                {
                    objSearch = useruser.FindName("btnSearch");

                    if (objSearch != null)
                    {
                        (objSearch as Button).RaiseEvent(new RoutedEventArgs(ButtonBase.ClickEvent));
                        //ButtonAutomationPeer peer = new ButtonAutomationPeer(objSearch as Button);
                        //IInvokeProvider invokeProv = peer.GetPattern(PatternInterface.Invoke) as IInvokeProvider;
                        //invokeProv.Invoke();
                    }
                }
            }
        }

        /// <summary>
        /// 저장에 성공했거나 취소했을시 적용되는 각 버튼 컨트롤
        /// </summary>
        /// <param name="userControl"></param>
        public void UiButtonEnableChange_IUControl(UserControl userControl)
        {
            MenuViewModel menuView = MainWindow.MainMdiContainer.ActiveMdiChild.Tag as MenuViewModel;

            object objAdd = null;
            object objDelete = null;
            object objSearch = null;
            object objUpdate = null;
            object objPrint = null;

            object objSave = null;
            object objCancel = null;
            object objExcel = null;
            object objMsg = null;

            objAdd = userControl.FindName("btnAdd");

            if (objAdd != null)
            {
                if (menuView.AddNewClss.Equals("*"))
                {
                    (objAdd as Button).IsEnabled = true;
                }
            }

            objDelete = userControl.FindName("btnDelete");

            if (objDelete != null)
            {
                if (menuView.DeleteClss.Equals("*"))
                {
                    (objDelete as Button).IsEnabled = true;
                }
            }

            objSearch = userControl.FindName("btnSearch");

            if (objSearch != null)
            {
                if (menuView.SelectClss.Equals("*"))
                {
                    (objSearch as Button).IsEnabled = true;
                }
                //(objSearch as Button).IsEnabled = false;
            }

            objUpdate = userControl.FindName("btnUpdate");

            if (objUpdate != null)
            {
                if (menuView.UpdateClss.Equals("*"))
                {
                    (objUpdate as Button).IsEnabled = true;
                }
            }

            objPrint = userControl.FindName("btnPrint");

            if (objPrint != null)
            {
                if (menuView.PrintClss.Equals("*"))
                {
                    (objPrint as Button).IsEnabled = true;
                }
            }


            objSave = userControl.FindName("btnSave");

            if (objSave != null)
            {
                (objSave as Button).Visibility = Visibility.Hidden;
            }

            objCancel = userControl.FindName("btnCancel");

            if (objCancel != null)
            {
                (objCancel as Button).Visibility = Visibility.Hidden;
            }

            objExcel = userControl.FindName("btnExcel");

            if (objExcel != null)
            {
                (objExcel as Button).Visibility = Visibility.Visible;
            }

            objMsg = userControl.FindName("lblMsg");

            if (objMsg != null)
            {
                (objMsg as Label).Visibility = Visibility.Hidden;
            }
        }

        /// <summary>
        /// 추가나 수정 버튼을 눌렀을시 적용되는 각 버튼 컨를롤
        /// </summary>
        /// <param name="userControl"></param>
        public void UiButtonEnableChange_SCControl(UserControl userControl)
        {
            MenuViewModel menuView = MainWindow.MainMdiContainer.ActiveMdiChild.Tag as MenuViewModel;

            object objAdd = null;
            object objDelete = null;
            object objSearch = null;
            object objUpdate = null;
            object objPrint = null;

            object objSave = null;
            object objCancel = null;
            object objExcel = null;
            object objMsg = null;

            objAdd = userControl.FindName("btnAdd");

            if (objAdd != null)
            {
                //if (menuView.AddNewClss.Equals("*"))
                //{
                //    (objAdd as Button).IsEnabled = true;
                //}
                (objAdd as Button).IsEnabled = false;
            }

            objDelete = userControl.FindName("btnDelete");

            if (objDelete != null)
            {
                //if (menuView.DeleteClss.Equals("*"))
                //{
                //    (objDelete as Button).IsEnabled = true;
                //}
                (objDelete as Button).IsEnabled = false;
            }

            objSearch = userControl.FindName("btnSearch");

            if (objSearch != null)
            {
                //if (menuView.SelectClss.Equals("*"))
                //{
                //    (objSearch as Button).IsEnabled = false;
                //}
                (objSearch as Button).IsEnabled = false;
            }

            objUpdate = userControl.FindName("btnUpdate");

            if (objUpdate != null)
            {
                //if (menuView.UpdateClss.Equals("*"))
                //{
                //    (objUpdate as Button).IsEnabled = true;
                //}
                (objUpdate as Button).IsEnabled = false;
            }

            objPrint = userControl.FindName("btnPrint");

            if (objPrint != null)
            {
                //if (menuView.PrintClss.Equals("*"))
                //{
                //    (objPrint as Button).IsEnabled = true;
                //}
                (objPrint as Button).IsEnabled = true;
            }


            objSave = userControl.FindName("btnSave");

            if (objSave != null)
            {
                (objSave as Button).Visibility = Visibility.Visible;
            }

            objCancel = userControl.FindName("btnCancel");

            if (objCancel != null)
            {
                (objCancel as Button).Visibility = Visibility.Visible;
            }

            objExcel = userControl.FindName("btnExcel");

            if (objExcel != null)
            {
                (objExcel as Button).Visibility = Visibility.Hidden;
            }

            objMsg = userControl.FindName("lblMsg");

            if (objMsg != null)
            {
                (objMsg as Label).Visibility = Visibility.Visible;
            }
        }



        #endregion

        #region DataGrid 내에서 포커스 주기

        /// <summary>
        /// DataGridInTextFocus using KeyUP
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        public void DataGridINTextBoxFocus(object sender, KeyEventArgs e)
        {
            DataGridCell cell = sender as DataGridCell;

            if (e.Key == Key.Enter)
            {
                if (cell.IsFocused == true)
                {
                    if (cell != null)
                    {
                        TextBox tb = FindVisualChild<TextBox>(cell);
                        if (tb != null)
                        {
                            if ((tb as TextBox).IsFocused == false)
                            {
                                (tb as TextBox).SelectAll();
                            }
                            (tb as TextBox).Focus();
                        }
                    }
                }
            }
        }

        /// <summary>
        /// DataGridInCombobox using KeyUP
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        public void DataGridINComboBoxFocus(object sender, KeyEventArgs e)
        {
            DataGridCell cell = sender as DataGridCell;

            if (e.Key == Key.Enter)
            {
                if (cell.IsFocused == true)
                {
                    if (cell != null)
                    {
                        ComboBox tb = FindVisualChild<ComboBox>(cell);
                        if (tb != null)
                        {
                            //if ((tb as ComboBox).IsFocused == false)
                            //{
                            //    //(tb as ComboBox).SelectAll();
                            //}
                            (tb as ComboBox).Focus();
                        }
                    }
                }
            }
        }

        public void DataGridINControlFocus(object sender, RoutedEventArgs e)
        {
            DataGridCell cell = sender as DataGridCell;

            if (cell.IsFocused == true)
            {
                if (cell != null)
                {
                    TextBox tb = FindVisualChild<TextBox>(cell);
                    ComboBox cb = FindVisualChild<ComboBox>(cell);
                    DatePicker dp = FindVisualChild<DatePicker>(cell);

                    if (dp != null)
                    {
                        //Button bt = FindVisualChild<Button>(cell);
                        //(bt as Button).Focus();
                        //(dp as DatePicker).Focus();
                        Popup popup = FindVisualChild<Popup>(cell);
                        (popup as Popup).IsOpen = true;
                    }
                    else if (tb != null)
                    {
                        if ((tb as TextBox).IsFocused == false)
                        {
                            (tb as TextBox).SelectAll();
                        }
                            (tb as TextBox).Focus();
                    }
                    else if (cb != null)
                    {
                        (cb as ComboBox).Focus();
                    }

                }
            }
        }


        /// <summary>
        /// MouseUP 으로 DataGrid안의 TextBoxFocus
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        public void DataGridINTextBoxFocusByMouseUP(object sender, MouseButtonEventArgs e)
        {
            DataGridCell cell = sender as DataGridCell;

            if (cell.IsFocused == true)
            {
                if (cell != null)
                {
                    TextBox tb = FindVisualChild<TextBox>(cell);
                    if (tb != null)
                    {
                        if ((tb as TextBox).IsFocused == false)
                        {
                            (tb as TextBox).SelectAll();
                        }
                        (tb as TextBox).Focus();
                    }
                }
            }
        }

        public void DataGridINComboBoxFocusByMouseUP(object sender, MouseButtonEventArgs e)
        {
            DataGridCell cell = sender as DataGridCell;

            if (cell.IsFocused == true)
            {
                if (cell != null)
                {
                    ComboBox tb = FindVisualChild<ComboBox>(cell);
                    if (tb != null)
                    {
                        //if ((tb as ComboBox).IsFocused == false)
                        //{
                        //    //(tb as ComboBox).SelectAll();
                        //}
                        (tb as ComboBox).Focus();
                    }
                }
            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        public void DataGridINBothByMouseUP(object sender, MouseButtonEventArgs e)
        {
            DataGridCell cell = sender as DataGridCell;

            if (cell.IsFocused == true)
            {
                if (cell != null)
                {
                    TextBox tb = FindVisualChild<TextBox>(cell);
                    ComboBox cb = FindVisualChild<ComboBox>(cell);
                    DatePicker dp = FindVisualChild<DatePicker>(cell);

                    if (dp != null)
                    {
                        //Button bt = FindVisualChild<Button>(cell);
                        //(bt as Button).Focus();
                        //(dp as DatePicker).Focus();
                        Popup popup = FindVisualChild<Popup>(cell);
                        (popup as Popup).IsOpen = true;
                    }
                    else if (tb != null)
                    {
                        if ((tb as TextBox).IsFocused == false)
                        {
                            (tb as TextBox).SelectAll();
                        }
                            (tb as TextBox).Focus();
                    }
                    else if (cb != null)
                    {
                        (cb as ComboBox).Focus();
                    }
                }
            }
        }

        #endregion


        #region 백그라운드에 남아 있는 엑셀 삭제
        //엑셀 백그라운드 증발 - 달달 2021-09-15
        public void ReleaseExcelObject(object obj)
        {
            try
            {
                if (obj != null)
                {
                    Marshal.ReleaseComObject(obj);
                    obj = null;
                }
            }
            catch (Exception ex)
            {
                obj = null;
                throw ex;
            }
            finally
            {
                //GC.Collect();
            }
        }
        #endregion

        public string ReturnGrade(string strScore)
        {
            string strGrade = string.Empty;

            if (int.Parse(strScore) >= 85)
            {
                strGrade = "A";
            }
            else if (int.Parse(strScore) < 85 && int.Parse(strScore) >= 75)
            {
                strGrade = "B";
            }
            else if (int.Parse(strScore) < 75 && int.Parse(strScore) >= 60)
            {
                strGrade = "C";
            }
            else
            {
                strGrade = "D";
            }

            return strGrade;
        }

        public string LogCompany(string strUserID)
        {
            string strComPanyID = string.Empty;

            string sql = "select mp.CompanyID from mt_Person mp         ";
            sql += "   WHERE 1          = 1                             ";
            sql += "    and     mp.UserID = '" + strUserID + "'             ";

            try
            {
                DataSet ds = DataStore.Instance.QueryToDataSet(sql);
                if (ds != null && ds.Tables.Count > 0)
                {
                    DataTable dt = ds.Tables[0];
                    if (dt.Rows.Count == 0)
                    {

                    }
                    else
                    {
                        DataRowCollection drc = dt.Rows;
                        strComPanyID = drc[0]["CompanyID"].ToString();
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("오류 발생, 오류내용 : " + ex.ToString());
            }
            finally
            {
                DataStore.Instance.CloseConnection();
            }


            return strComPanyID;
        }




        /// <summary>
        /// List<string> 이 모두 포함된 DataGrid 의 RowIndex 를 찾아서 return. 
        /// </summary>
        /// <param name="dgd"></param>
        /// <param name="lstTargetString"></param>
        /// <returns></returns>
        public int ReTrunIndex(DataGrid dgd, List<string> lstTargetString)
        {
            bool flag = true;
            int count = 0;
            ItemCollection item = dgd.Items;

            if (lstTargetString.Count <= 0)
            {
                return count;
            }

            for (int i = 0; i < item.Count; i++)
            {
                flag = true;
                for (int j = 0; j < lstTargetString.Count; j++)
                {
                    if (!item[i].ToString().Contains(lstTargetString[j]))
                    {
                        flag = false;
                        break;
                    }
                }

                if (flag)
                {
                    count = i;
                    break;
                }
            }

            return count;
        }

        /// <summary>
        /// string값이 포함된 DataGrid 의 RowIndex 를 찾아서 return. 
        /// </summary>
        /// <param name="dgd"></param>
        /// <param name="lstTargetString"></param>
        /// <returns></returns>
        public int ReTrunIndex(DataGrid dgd, string lstTargetString)
        {
            int count = 0;
            ItemCollection item = dgd.Items;

            if (lstTargetString == null || lstTargetString.Equals(""))
            {
                return count;
            }

            for (int i = 0; i < item.Count; i++)
            {
                if (item[i].ToString().Equals(lstTargetString))
                {
                    count = i;
                    break;
                }
            }

            return count;
        }

        #region 기타

        public string UserIPAddress
        {
            get
            {
                IPHostEntry IPHost = Dns.GetHostByName(Dns.GetHostName());

                string _UserIPAddress = IPHost.AddressList[0].ToString();

                return _UserIPAddress;
            }
        }

        // 딜레이 먹이기.
        public void Delay(int MS)
        {
            //var timer = new DispatcherTimer { Interval = TimeSpan.FromMilliseconds(MS) };
            //timer.Start();
            //timer.Tick += (sender, args) =>
            //{
            //    timer.Stop();               
            //};

            DateTime ThisMoment = DateTime.Now;
            TimeSpan duration = new TimeSpan(0, 0, 0, 0, MS);
            DateTime AfterWards = ThisMoment.Add(duration);
            while (AfterWards >= ThisMoment)
            {
                Application.Current.Dispatcher.Invoke(DispatcherPriority.Loaded, new Action(delegate { }));
                ThisMoment = DateTime.Now;
            }
        }

        /// <summary>
        ///   Sends the specified key.
        ///   키 입력. sndekey가 wpf에는 없어서 꼼수로 구현한 거.  
        /// </summary>
        /// <param name="key">The key.</param>
        public void SendK(Key key, UserControl nowForm)
        {
            if (Keyboard.PrimaryDevice != null)
            {
                if (Keyboard.PrimaryDevice.ActiveSource != null)
                {
                    Keyboard.Focus(nowForm);
                    PresentationSource source = PresentationSource.FromVisual(nowForm);
                    var e = new KeyEventArgs(Keyboard.PrimaryDevice,
                        source, 0, key)
                    {
                        RoutedEvent = Keyboard.KeyDownEvent
                    };
                    Keyboard.Focus(nowForm);
                    InputManager.Current.ProcessInput(e);
                    Keyboard.Focus(nowForm);

                    // Note: Based on your requirements you may also need to fire events for:
                    // RoutedEvent = Keyboard.PreviewKeyDownEvent
                    // RoutedEvent = Keyboard.KeyUpEvent
                    // RoutedEvent = Keyboard.PreviewKeyUpEvent
                }
            }
        }


        // null + 여백(스페이스 바) 체크 작업.
        public bool IsNullOrWhiteSpace(string value)
        {
            if (value == null) return true;

            for (int i = 0; i < value.Length; i++)
            {
                if (!Char.IsWhiteSpace(value[i])) return false;
            }

            return true;
        }

        /// 문자열 사이사이 마다 스페이스바 넣어주는 로직
        public string SetStringSpace(string str)
        {
            StringBuilder sb = new StringBuilder();
            int Len = str.Trim().Length;
            if (Len > 1)
            {
                for (int i = 0; i < Len; i++)
                {
                    sb.Append(str[i]);
                    if ((Len - 1) != i)//마지막번째 i면 스페이스바 추가 안함
                    {
                        sb.Append(" ");
                    }
                }
            }
            else
            {
                sb.Append(str.Trim());
            }

            return sb.ToString();
        }

        //뭔지 기억이...
        private DataTemplate GenerateTextBlockTemplate(string property)
        {
            FrameworkElementFactory factory = new FrameworkElementFactory(typeof(TextBlock));
            factory.SetBinding(TextBlock.TextProperty, new Binding(property));

            return new DataTemplate { VisualTree = factory };
        }

        // 사용자 컴터 ip 받아오기.
        public static IPAddress GetIPAddress()
        {
            IPAddress[] hostAddresses = Dns.GetHostAddresses("");

            foreach (IPAddress hostAddress in hostAddresses)
            {
                if (hostAddress.AddressFamily == AddressFamily.InterNetwork &&
                    !IPAddress.IsLoopback(hostAddress) &&  // ignore loopback addresses
                    !hostAddress.ToString().StartsWith("169.254."))  // ignore link-local addresses
                    return hostAddress;
            }
            return null; // or IPAddress.None if you prefer
        }

        // 메뉴 시작전, 메뉴 로그정보 저장하기.
        public void AllMenuLogInsert(string MenuID, string MenuName, object objList)
        {
            try
            {
                IPAddress userhost = GetIPAddress();

                Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                sqlParameter.Clear();
                sqlParameter.Add("sCompanyID", MainWindow.CompanyID);
                sqlParameter.Add("sMenuID", MenuID);
                sqlParameter.Add("sWorkFlag", "");
                sqlParameter.Add("sWorkDate", DateTime.Now.ToString("yyyyMMdd"));
                sqlParameter.Add("sWorkTime", DateTime.Now.ToString("HHmm"));

                sqlParameter.Add("sUserID", MainWindow.CurrentUser);
                sqlParameter.Add("sWorkComputer", System.Environment.MachineName);
                sqlParameter.Add("sWorkComputerIP", userhost.ToString());
                sqlParameter.Add("sWorkLog", "WellMade" + " " + "[" + MenuID + "]" + " " + MenuName + " " + "(" + objList.ToString() + ")");


                string[] result = DataStore.Instance.ExecuteProcedure("xp_iWorkLogWPF", sqlParameter, false);
                if (!result[0].Equals("success"))
                {
                    MessageBox.Show("이상발생, 관리자에게 문의하세요 " + result[1]);
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

        public int reTrunIndex(DataGrid dgd, List<string> lstTargetString)
        {
            bool flag = true;
            int count = 0;
            ItemCollection item = dgd.Items;

            if (lstTargetString.Count <= 0)
            {
                return count;
            }

            for (int i = 0; i < item.Count; i++)
            {
                flag = true;
                for (int j = 0; j < lstTargetString.Count; j++)
                {
                    if (!item[i].ToString().Contains(lstTargetString[j]))
                    {
                        flag = false;
                        break;
                    }
                }

                if (flag)
                {
                    count = i;
                    break;
                }
            }

            return count;
        }

        public int reTrunIndex(DataGrid dgd, string lstTargetString)
        {
            int count = 0;
            ItemCollection item = dgd.Items;

            if (lstTargetString == null || lstTargetString.Equals(""))
            {
                return count;
            }

            for (int i = 0; i < item.Count; i++)
            {
                if (item[i].ToString().Equals(lstTargetString))
                {
                    count = i;
                    break;
                }
            }

            return count;
        }

        //
        public string returnNumString(string NumString)
        {
            string returnString = string.Empty;

            if (NumString != null)
            {
                int targetIndex = 0;

                if (NumString.Contains("."))
                {
                    targetIndex = CheckNotZero(NumString);
                }

                returnString = returnNumStringTargetNum(NumString, targetIndex);
            }

            return returnString;
        }

        //정수로 리턴
        public double returnDouble(object Target)
        {
            double returnDoubleNum = 0.00;
            if (Target != null)
            {
                if (!Target.ToString().Trim().Equals(string.Empty) && IsNumOrAnother(Target.ToString()))
                {
                    returnDoubleNum = Convert.ToDouble(Lib.Instance.returnNumStringTargetNum(Target.ToString(), 3).Replace(",", ""));
                }
            }
            return returnDoubleNum;
        }

        //정수로 리턴
        public string returnNumStringZero(string strTarget)
        {
            string returnString = string.Empty;

            if (strTarget.Contains(","))
            {
                strTarget = strTarget.Replace(",", "");
            }

            if (this.IsNumOrAnother(strTarget))
            {
                returnString = string.Format("{0:N0}", double.Parse(strTarget));
            }
            else
            {
                returnString = strTarget;
            }

            return returnString;
        }

        //소수점 한자리 리턴
        public string returnNumStringOne(string strTarget)
        {
            string returnString = string.Empty;

            if (strTarget.Contains(","))
            {
                strTarget = strTarget.Replace(",", "");
            }

            if (this.IsNumOrAnother(strTarget))
            {
                returnString = string.Format("{0:N1}", double.Parse(strTarget));
            }
            else
            {
                returnString = strTarget;
            }

            return returnString;
        }

        //소수점 두자리 리턴
        public string returnNumStringTwo(string strTarget)
        {
            string returnString = string.Empty;

            if (strTarget.Contains(","))
            {
                strTarget = strTarget.Replace(",", "");
            }

            if (this.IsNumOrAnother(strTarget))
            {
                returnString = string.Format("{0:N2}", double.Parse(strTarget));
            }
            else
            {
                returnString = strTarget;
            }

            return returnString;
        }


        //소수점 시자리 리턴
        public string returnNumStringThree(string strTarget)
        {
            string returnString = string.Empty;

            if (strTarget.Contains(","))
            {
                strTarget = strTarget.Replace(",", "");
            }

            if (this.IsNumOrAnother(strTarget))
            {
                returnString = string.Format("{0:N3}", double.Parse(strTarget));
            }
            else
            {
                returnString = strTarget;
            }

            return returnString;
        }

        //소수점은 놔두고 리턴
        public string returnNumStringTwoExceptDot(string strTarget)
        {
            string returnString = string.Empty;
            string strDotBack = string.Empty;

            if (strTarget.Contains("."))
            {
                strDotBack = strTarget.Substring(strTarget.IndexOf("."));
                strTarget = strTarget.Substring(0, strTarget.IndexOf("."));
            }

            if (strTarget.Contains(","))
            {
                strTarget = strTarget.Replace(",", "");
            }

            if (this.IsNumOrAnother(strTarget))
            {
                returnString = string.Format("{0:N0}", double.Parse(strTarget));
            }
            else
            {
                returnString = strTarget;
            }

            returnString += strDotBack;
            return returnString;
        }

        //소수점은 놔두고 리턴
        public string returnNumStringTargetNum(string strTarget, int targetNum)
        {
            string returnString = string.Empty;

            if (strTarget.Contains(","))
            {
                strTarget = strTarget.Replace(",", "");
            }

            if (this.IsNumOrAnother(strTarget))
            {
                returnString = string.Format("{0:N" + targetNum + "}", double.Parse(strTarget));
            }
            else
            {
                returnString = strTarget;
            }

            return returnString;
        }

        /// <summary>
        /// string 값을 받아야 할경우 Null이면 "" 리턴 아니면 원래값 리턴
        /// </summary>
        /// <param name="strNullCheck"></param>
        /// <returns></returns>
        public string CheckNull(string strNullCheck)
        {
            string strReturn = string.Empty;

            if (strNullCheck is null)
            {
                strReturn = "";
            }
            else
            {
                strReturn = strNullCheck;
            }

            return strReturn;
        }

        /// <summary>
        /// object를 받아서 Null이면 "", 아니면 .ToString() 반환
        /// </summary>
        /// <param name="sender"></param>
        /// <returns></returns>
        public string CheckNull(object sender)
        {
            if (sender is null)
            {
                return "";
            }
            else
            {
                return sender.ToString();
            }
        }

        /// <summary>
        /// 숫자 값을 받아야 할경우 Null이면 "0" 리턴 아니면 원래값 리턴
        /// </summary>
        /// <param name="strNullCheck"></param>
        /// <returns></returns>
        public string CheckNullZero(string strNullCheck)
        {
            string strReturn = string.Empty;

            if (strNullCheck is null || strNullCheck.Equals(string.Empty))
            {
                strReturn = "0";
            }
            else
            {
                strReturn = strNullCheck;
            }

            return strReturn;
        }

        public int CheckNotZero(string sender)
        {
            int index = 0;
            int startIndex = sender.IndexOf(".");
            int lastIndex = sender.Length - startIndex;
            string strSender = sender.Substring(startIndex, lastIndex);

            char[] charArray = strSender.ToCharArray();
            for (int i = charArray.Length - 1; i > 0; i--)
            {
                if (charArray[i] != '0')
                {
                    index = i;
                    break;
                }
            }

            return index;
        }

        // Int로 변환
        public int ConvertInt(string str)
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

        // 숫자로 변환 가능한지 체크 이벤트
        public bool CheckConvertInt(string str)
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

        // Double형으로 변환 가능한지 체크 이벤트
        public bool CheckConvertDouble(string str)
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
        /// <summary>
        /// 0 은 이름, 1은 tag
        /// </summary>
        /// <returns></returns>
        public string[] SetPerson()
        {
            string[] strArray = new string[2];

            ; string sql = "SELECT PersonID , Name    from  mt_Person    ";
            sql += "   WHERE 1          = 1                            ";
            sql += "   AND UserID     =   '" + MainWindow.CurrentUser + "'    ";

            try
            {
                DataSet ds = DataStore.Instance.QueryToDataSet(sql);
                if (ds != null && ds.Tables.Count > 0)
                {
                    DataTable dt = ds.Tables[0];
                    if (dt.Rows.Count > 0)
                    {
                        DataRowCollection drc = dt.Rows;

                        strArray[0] = drc[0]["Name"].ToString();
                        strArray[1] = drc[0]["PersonID"].ToString();
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("오류 발생 : " + ex.ToString());
            }
            finally
            {
                DataStore.Instance.CloseConnection();
            }

            return strArray;
        }

        /// <summary>
        /// 화면의 닫기 버튼
        /// </summary>
        /// <param name="strName"></param>
        public void ChildMenuClose(string strName)
        {
            for (int i = 0; i < MainWindow.MainMdiContainer.Children.Count; i++)
            {
                if (strName.Equals((MainWindow.MainMdiContainer.Children[i] as MdiChild).Content.ToString()))
                {
                    (MainWindow.MainMdiContainer.Children[i] as MdiChild).Close();
                    break;
                }
            }

        }
        public void DBReIndex() //2021-11-10 재고조사 후 DBREINDEX 하기 위해 생성
        {
            Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
            sqlParameter.Clear();
            string[] result = DataStore.Instance.ExecuteProcedure("xp_DB_REINDEX", sqlParameter, false);
            if (!result[0].Equals("success"))
            {
                MessageBox.Show("이상발생, 관리자에게 문의하세요.");
                return;
            }
        }

        public string FindMoldNo(string strMoldID)
        {
            string strReturn = string.Empty;

            try
            {
                string sql = "SELECT MoldNo                              ";
                sql += "    FROM dvl_Mold                                ";
                sql += "   WHERE 1          = 1                          ";
                sql += "   and   MoldID     = '" + strMoldID + "'           ";

                DataSet ds = DataStore.Instance.QueryToDataSet(sql);
                if (ds != null && ds.Tables.Count > 0)
                {
                    DataTable dt = ds.Tables[0];
                    if (dt.Rows.Count > 0)
                    {
                        strReturn = dt.Rows[0]["MoldNo"].ToString();
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("오류 발생,오류 내용 : " + ex.ToString());
            }
            finally
            {
                DataStore.Instance.CloseConnection();
            }

            return strReturn;
        }

        #region 라벨 클릭, 체크 박스 클릭 

        //사용하려면 같은 그리드로 묶어주세요
        //<Grid><Label/><CheckBox/></Grid>
        public void CommonControl_Click(object sender, EventArgs e)
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

        //컨트롤 안 특정 타입의 자식 컨트롤을 찾는 함수 (그리드내에서)
        //var parentContainer = VisualTreeHelper.GetParent(checkbox);
        //var datePicker = FindChild<DatePicker>(parentContainer);
        public T FindChild<T>(DependencyObject parent) where T : DependencyObject
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

        // 자식요소 안에서 부모요서 찾기
        //DataGridRow row = FindVisualParent<DataGridRow>(checkBox); 데이터그리드안의 행속 체크박스의 부모행 찾기
        //DataGrid parentGrid = FindVisualParent<DataGrid>(row); 데이터그리드 행의 부모 데이터그리드 찾기
        public T FindVisualParent<T>(DependencyObject child) where T : DependencyObject
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

        //UI컨트롤 요소찾기 
        //StackPanel같이 다중 트리로 되어있는 UI들도 깊게 탐색
        public void FindUiObject(DependencyObject parent, Action<DependencyObject> action)
        {
            if (parent == null) return;

            var visited = new HashSet<DependencyObject>();
            var queue = new Queue<DependencyObject>();

            queue.Enqueue(parent);

            while (queue.Count > 0 && visited.Count < 2000)
            {
                var current = queue.Dequeue();
                if (current == null || visited.Contains(current)) continue;

                visited.Add(current);
                action?.Invoke(current);

                // 1. 비쥬얼트리 자식들 추가
                if (current is Visual || current is Visual3D)
                {
                    try
                    {
                        int childCount = VisualTreeHelper.GetChildrenCount(current);
                        for (int i = 0; i < childCount; i++)
                        {
                            var child = VisualTreeHelper.GetChild(current, i);
                            if (child != null && !visited.Contains(child))
                            {
                                queue.Enqueue(child);
                            }
                        }
                    }
                    catch { }
                }

                // 2. LogicalTree 자식들도 추가
                try
                {
                    foreach (object logicalChild in LogicalTreeHelper.GetChildren(current))
                    {
                        if (logicalChild is DependencyObject child && !visited.Contains(child))
                        {
                            queue.Enqueue(child);
                        }
                    }
                }
                catch { }

                // 3. 그외
                switch (current)
                {
                    case ContentPresenter cp when cp.Content is DependencyObject content:
                        if (!visited.Contains(content))
                            queue.Enqueue(content);
                        break;

                    case ContentControl cc when cc.Content is DependencyObject controlContent:
                        if (!visited.Contains(controlContent))
                            queue.Enqueue(controlContent);
                        break;

                    case Decorator decorator when decorator.Child != null:
                        if (!visited.Contains(decorator.Child))
                            queue.Enqueue(decorator.Child);
                        break;
                }
            }
        }

        //같은 그리드안에 이웃한 요소 한개를 반환(현재 선택값과 비교가능) 
        public T FindSiblingControl<T>(UIElement currentElement) where T : UIElement
        {
            var parentGrid = FindVisualParent<Grid>(currentElement);
            if (parentGrid == null) return null;

            int currentColumn = Grid.GetColumn(currentElement);
            int currentRow = Grid.GetRow(currentElement);

            foreach (UIElement child in parentGrid.Children)
            {
                if (child is T targetControl && child != currentElement)
                {
                    int childColumn = Grid.GetColumn(child);
                    int childRow = Grid.GetRow(child);

                    // 같은 Row이면서 다른 Column인 컨트롤
                    if (childRow == currentRow && childColumn != currentColumn)
                    {
                        return targetControl;
                    }
                }
            }
            return null;
        }

        //같은 그리드안에 이웃한 요소 여러개를  리스트로 반환(현재 선택값과 비교가능) 단, 반복문으로
        public List<T> FindAllSiblingControls<T>(UIElement currentElement) where T : UIElement
        {
            var siblings = new List<T>();
            var parentGrid = FindVisualParent<Grid>(currentElement);
            if (parentGrid == null) return siblings;

            int currentColumn = Grid.GetColumn(currentElement);
            int currentRow = Grid.GetRow(currentElement);

            foreach (UIElement child in parentGrid.Children)
            {
                if (child is T targetControl && child != currentElement)
                {
                    int childColumn = Grid.GetColumn(child);
                    int childRow = Grid.GetRow(child);

                    // 같은 Row이면서 다른 Column인 컨트롤
                    if (childRow == currentRow && childColumn != currentColumn)
                    {
                        siblings.Add(targetControl);
                    }
                }
            }
            return siblings;
        }

        #endregion

        #region 콤마 제거 오버로딩 함수들


        // 문자열 반환 (콤마 제거만)
        public string RemoveComma(object obj)
        {
            if (obj == null) return "0";

            string result = obj.ToString().Trim().Replace(",", "");
            return string.IsNullOrEmpty(result) ? "0" : result;
        }

        // int 반환
        public int RemoveComma(object obj, int defaultValue, bool showErrorMessage = true)
        {
            if (obj == null) return defaultValue;

            string digits = obj.ToString().Trim().Replace(",", "");
            if (string.IsNullOrEmpty(digits)) return defaultValue;

            if (!decimal.TryParse(digits, out decimal parsedValue))
            {
                if (showErrorMessage)
                    MessageBox.Show($"'{obj}'는 올바른 숫자 형식이 아닙니다.", "입력 오류");
                return defaultValue;
            }

            if (parsedValue > int.MaxValue || parsedValue < int.MinValue)
            {
                if (showErrorMessage)
                    ShowRangeError("정수", int.MinValue, int.MaxValue);
                return defaultValue;
            }

            return (int)parsedValue;
        }

        // long 반환
        public long RemoveComma(object obj, long defaultValue, bool showErrorMessage = true)
        {
            if (obj == null) return defaultValue;

            string digits = obj.ToString().Trim().Replace(",", "");
            if (string.IsNullOrEmpty(digits)) return defaultValue;

            if (!decimal.TryParse(digits, out decimal parsedValue))
            {
                if (showErrorMessage)
                    MessageBox.Show($"'{obj}'는 올바른 숫자 형식이 아닙니다.", "입력 오류");
                return defaultValue;
            }

            if (parsedValue > long.MaxValue || parsedValue < long.MinValue)
            {
                if (showErrorMessage)
                    ShowRangeError("정수", long.MinValue, long.MaxValue);
                return defaultValue;
            }

            return (long)parsedValue;
        }

        // decimal 반환
        public decimal RemoveComma(object obj, decimal defaultValue, bool showErrorMessage = true)
        {
            if (obj == null) return defaultValue;

            string digits = obj.ToString().Trim().Replace(",", "");
            if (string.IsNullOrEmpty(digits)) return defaultValue;

            if (!decimal.TryParse(digits, out decimal parsedValue))
            {
                if (showErrorMessage)
                    MessageBox.Show($"'{obj}'는 올바른 숫자 형식이 아닙니다.", "입력 오류");
                return defaultValue;
            }

            return parsedValue;
        }

        // double 반환
        public double RemoveComma(object obj, double defaultValue, bool showErrorMessage = true)
        {
            if (obj == null) return defaultValue;

            string digits = obj.ToString().Trim().Replace(",", "");
            if (string.IsNullOrEmpty(digits)) return defaultValue;

            if (!decimal.TryParse(digits, out decimal parsedValue))
            {
                if (showErrorMessage)
                    MessageBox.Show($"'{obj}'는 올바른 숫자 형식이 아닙니다.", "입력 오류");
                return defaultValue;
            }

            double doubleVal = (double)parsedValue;
            if (double.IsInfinity(doubleVal))
            {
                if (showErrorMessage)
                    MessageBox.Show("입력한 값이 너무 큽니다.", "범위 초과");
                return defaultValue;
            }

            return doubleVal;
        }

        // float 반환
        public float RemoveComma(object obj, float defaultValue, bool showErrorMessage = true)
        {
            if (obj == null) return defaultValue;

            string digits = obj.ToString().Trim().Replace(",", "");
            if (string.IsNullOrEmpty(digits)) return defaultValue;

            if (!decimal.TryParse(digits, out decimal parsedValue))
            {
                if (showErrorMessage)
                    MessageBox.Show($"'{obj}'는 올바른 숫자 형식이 아닙니다.", "입력 오류");
                return defaultValue;
            }

            float floatVal = (float)parsedValue;
            if (float.IsInfinity(floatVal))
            {
                if (showErrorMessage)
                    MessageBox.Show("입력한 값이 너무 큽니다.", "범위 초과");
                return defaultValue;
            }

            return floatVal;
        }

        public void ShowRangeError(string type, object min, object max)
        {
            MessageBox.Show($"입력한 값이 {type} 처리 가능한 범위를 벗어났습니다.\n(범위: {min:N0} ~ {max:N0})",
                            "범위 초과");
        }


        #endregion

        #region 날짜변환

        public bool DatePickerCheck(DatePicker sDatePicker, DatePicker eDatePicker, CheckBox chkDatePicker = null)
        {
            try
            {
                if (chkDatePicker != null)
                {
                    if (chkDatePicker.IsChecked == false)
                        return true;
                }

                if ((sDatePicker == null && eDatePicker == null) || sDatePicker.SelectedDate == null || eDatePicker.SelectedDate == null)
                {
                    MessageBox.Show("검색기간 시작일과 종료일이 지정되었는지 확인하세요", "확인");
                    return false;
                }
                else
                {
                    if (sDatePicker.SelectedDate > eDatePicker.SelectedDate)
                    {
                        MessageBox.Show("검색기간 시작일이 종료일보다 앞설 수 없습니다.", "확인");
                        return false;
                    }

                }

            }
            catch
            {
                return false;
            }

            return true;
        }

        //8자리 char형태 날짜 년도-월-일 하이픈 삽입
        //16자리 일경우 8자리 사이에 ~ 삽입
        public string DateTypeHyphen(string DigitsDate)
        {
            string pattern1 = @"(\d{4})(\d{2})(\d{2})";
            string pattern2 = @"(\d{4})(\d{2})(\d{2})(\d{4})(\d{2})(\d{2})";
            string pattern3 = @"(\d{4})(\d{2})";

            if (DigitsDate.Length == 8)
            {
                DigitsDate = Regex.Replace(DigitsDate, pattern1, "$1-$2-$3");
            }
            else if(DigitsDate.Length == 6)
            {
                DigitsDate = Regex.Replace(DigitsDate, pattern3, "$1-$2");
            }
            else if (DigitsDate.Length == 16)
            {
                DigitsDate = Regex.Replace(DigitsDate, pattern2, "$1-$2-$3 ~ $4-$5-$6");
            }
            else if (DigitsDate.Length == 0)
            {
                DigitsDate = string.Empty;
            }

            return DigitsDate;
        }

        public object RemoveHyphen(object obj)
        {
            if (obj == null)
                return string.Empty;

            if (obj.ToString() == string.Empty)
                return string.Empty;

            if (obj.ToString().Contains("-"))
            {
                return obj.ToString().Replace("-", "");
            }


            return obj;
        }

        public string SetToDate(object obj)
        {
            if (DateTime.TryParse(obj.ToString(), out DateTime date))
            {
                return date.ToString("yyyyMMdd");
            }
            return obj.ToString();
        }


        public string ConvertDate(DatePicker datePicker)
        {
            if (datePicker.SelectedDate != null)
                return datePicker.SelectedDate.Value.ToString("yyyyMMdd");
            else
                return string.Empty;
        }

        public bool IsDatePickerNull(DatePicker datePicker)
        {
            if (datePicker.SelectedDate == null)
                return true;
            else
                return false;
        }


        public string TimeTypeColon(string DigitsTime)
        {
            string pattern1 = @"(\d{2})(\d{2})";

            if (DigitsTime.Length == 4)
            {
                DigitsTime = Regex.Replace(DigitsTime, pattern1, "$1:$2");
            }

            return DigitsTime;
        }
        #endregion

        #region UI컨트롤에 툴팁띄우기
        public void CloseToolTip()
        {
            if (currentToolTip != null && currentToolTip.IsOpen)
            {
                currentToolTip.IsOpen = false;
                if (currentTimer != null)
                {
                    currentTimer.Stop();
                    currentTimer = null;
                }
            }
        }

        public void ShowTooltipMessage(FrameworkElement element, string message, MessageBoxImage iconType = MessageBoxImage.None, PlacementMode placement = PlacementMode.Bottom, double scale = 1.0)
        {
            // 이미 열려있는 툴팁이 있다면 닫기
            if (currentToolTip != null && currentToolTip.IsOpen)
            {
                currentToolTip.IsOpen = false;
                if (currentTimer != null)
                {
                    currentTimer.Stop();
                    currentTimer = null;
                }
            }

            object tooltipContent;

            // 아이콘이 필요 없는 경우
            if (iconType == MessageBoxImage.None)
            {
                tooltipContent = message;
            }
            else
            {
                // StackPanel 생성
                var stackPanel = new StackPanel
                {
                    Orientation = Orientation.Horizontal
                };

                // 시스템 아이콘 설정
                System.Drawing.Icon systemIcon;
                switch (iconType)
                {
                    case MessageBoxImage.Information:
                        systemIcon = System.Drawing.SystemIcons.Information;
                        break;
                    case MessageBoxImage.Warning:
                        systemIcon = System.Drawing.SystemIcons.Warning;
                        break;
                    case MessageBoxImage.Error:
                        systemIcon = System.Drawing.SystemIcons.Error;
                        break;
                    case MessageBoxImage.Question:
                        systemIcon = System.Drawing.SystemIcons.Question;
                        break;
                    default:
                        systemIcon = null;
                        break;
                }

                if (systemIcon != null)
                {
                    // System.Drawing에서 아이콘 가져오기
                    System.Windows.Media.Imaging.BitmapSource iconSource =
                        System.Windows.Interop.Imaging.CreateBitmapSourceFromHIcon(
                            systemIcon.Handle,
                            System.Windows.Int32Rect.Empty,
                            System.Windows.Media.Imaging.BitmapSizeOptions.FromEmptyOptions());

                    // 이미지 생성
                    var image = new Image
                    {
                        Source = iconSource,
                        Width = 16,
                        Height = 16,
                        Margin = new Thickness(0, 0, 5, 0)
                    };

                    // StackPanel에 추가
                    stackPanel.Children.Add(image);
                }

                // 텍스트블록 생성
                var textBlock = new TextBlock
                {
                    Text = message,
                    VerticalAlignment = VerticalAlignment.Center
                };

                stackPanel.Children.Add(textBlock);
                tooltipContent = stackPanel;
            }

            // 새 툴팁 생성
            var tooltip = new ToolTip
            {
                Content = tooltipContent,
                PlacementTarget = element,
                Placement = placement,
                IsOpen = true,
                LayoutTransform = new ScaleTransform(scale, scale)
            };

            // 위치에 따른 설정
            if (placement == PlacementMode.Bottom)
            {
                tooltip.VerticalOffset = 5;
            }
            else if (placement == PlacementMode.Right)
            {

                element.Dispatcher.BeginInvoke(new Action(() =>
                {
                    tooltip.HorizontalOffset = 5;
                }));
            }

            currentToolTip = tooltip;

            // 3초 후 툴팁 자동 닫기
            var timer = new System.Windows.Threading.DispatcherTimer
            {
                Interval = TimeSpan.FromSeconds(3)
            };
            timer.Tick += (s, eventArgs) =>
            {
                tooltip.IsOpen = false;
                timer.Stop();
            };
            timer.Start();
            currentTimer = timer;
        }
        #endregion

        #region 스크린샷

        public void ScreenCapture(FrameworkElement targetElement)
        {
            try
            {
                //파라미터로 전달받기
                //FrameworkElement targetElement = this;
                //너무 커지는 것을 방지
                double renderWidth = Math.Max(targetElement.ActualWidth, 400);
                double renderHeight = Math.Max(targetElement.ActualHeight, 300);

                //보이는 것들을 렌더링
                RenderTargetBitmap renderTarget = new RenderTargetBitmap(
                    (int)renderWidth,
                    (int)renderHeight,
                    96, 96,
                    PixelFormats.Pbgra32
                );

                //지정이 안되어서 투명하게 되는것들(누끼따지는 것처럼 되는 것들)
                //기본 WPF배경색상 채우기
                DrawingVisual backgroundVisual = new DrawingVisual();
                using (DrawingContext context = backgroundVisual.RenderOpen())
                {
                    context.DrawRectangle(
                        new SolidColorBrush(Color.FromRgb(0xF0, 0xF0, 0xF0)),
                        null,
                        new Rect(0, 0, renderWidth, renderHeight)
                    );
                }

                //그리기
                renderTarget.Render(backgroundVisual);
                renderTarget.Render(targetElement);

                // PNG 파일로 저장
                //string fileName = @"c:\temp\" + DateTime.Now.ToString("yyyy-MM-dd_HHmmss") + ".png";
                //PngBitmapEncoder encoder = new PngBitmapEncoder();
                //encoder.Frames.Add(BitmapFrame.Create(renderTarget));

                //using (FileStream stream = new FileStream(fileName, FileMode.Create, FileAccess.Write))
                //{
                //    encoder.Save(stream);
                //}

                //이미지를 팝업창에 전달
                System.Windows.Controls.Image ImgImage = new System.Windows.Controls.Image();
                ImgImage.Source = renderTarget;

                if (ImgImage.Source != null)
                {
                    MainWindow.ScreenCapture.Clear();
                    MainWindow.ScreenCapture.Add(ImgImage);
                }
                PopUp.ScreenShot SCshot = new PopUp.ScreenShot();

                SCshot.ShowDialog();


            }
            catch (Exception ex)
            {
                MessageBox.Show($"스크린샷 캡처 중 오류가 발생했습니다: {ex.Message}");
            }
        }

        #endregion

        #region 출하

        //박스수량구하기, 재고함수 프로시저 사용
        public int GetBoxQty(string BoxID)
        {
            int qty = 0;

            try
            {

                Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                sqlParameter.Clear();
                sqlParameter.Add("BoxID", BoxID);

                DataSet ds = DataStore.Instance.ProcedureToDataSet("xp_Outware_sGetBoxQty", sqlParameter, false);


                if (ds != null && ds.Tables.Count > 0)
                {
                    DataTable dt = ds.Tables[0];

                    if (dt.Rows.Count > 0)
                    {
                        DataRowCollection drc = dt.Rows;

                        foreach (DataRow dr in drc)
                        {
                            qty = RemoveComma(dr["BoxQty"].ToString(), 0);
                        }
                    }
                }
            }
            catch
            {

            }


            return qty;
        }

        #endregion
    }

    public class TextBoxColumnControl : TextBox
    {
        public event EventHandler<EventArgs> txtAction = delegate { };

        public TextBoxColumnControl()
        {
            PreviewKeyUp += TextBoxColumnControl_PreviewKeyUp;
        }

        void TextBoxColumnControl_PreviewKeyUp(object sender, System.Windows.Input.KeyEventArgs e)
        {
            if (e.Key != System.Windows.Input.Key.Enter) { return; }
            txtAction(this, EventArgs.Empty);
        }
    }


}

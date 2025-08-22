using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using WizMes_WellMade.PopUP;
using WPF.MDI;

namespace WizMes_WellMade
{
    /**************************************************************************************************
    '** System 명 : WizMes_WellMade
    '** Author    : Wizard
    '** 작성자    : 최준호
    '** 내용      : 금형/설비 부품 수불조회
    '** 생성일자  : 2018.10월~2019.01월 사이
    '** 변경일자  : 
    '**------------------------------------------------------------------------------------------------
    ''*************************************************************************************************
    ' 변경일자  , 변경자, 요청자    , 요구사항ID  , 요청 및 작업내용
    '**************************************************************************************************
    ' ex) 2015.11.09, 박진성, 오영      ,S_201510_AFT_03 , 월별집계(가로) 순서 변경 : 합계/10월/9월/8월 순으로
    ' 2019.07.08  최준호 , 최규환  설비명-> 부품명으로 , 용도 추가
    ' 2019.07.09  최준호 , 최규한  부품명-> 예비품으로 , 부품명, 입출고처, 비고를 크게 잡고 나머지 좀 작게
                                   부품명을 기준으로 왼쪽을 고정시켜달라., 출고량 -> 사용량(작업 중)
    '**************************************************************************************************/
    /// <summary>
    /// Win_prd_PartSubul_Q.xaml에 대한 상호 작용 논리
    /// </summary>
    public partial class Win_prd_PartSubul_Q : UserControl
    {
        #region 변수 선언 및 로드
        Lib lib = new Lib();
        public Win_prd_PartSubul_Q()
        {
            InitializeComponent();
        }

        private void UserControl_Loaded(object sender, RoutedEventArgs e)
        {
            Lib.Instance.UiLoading(sender);
            dtpSDate.SelectedDate = DateTime.Today;
            dtpEDate.SelectedDate = DateTime.Today;

            SetComboBox();
            chkMcInOutDate.IsChecked = true;
            tgnAll.IsChecked = true;
        }

        private void SetComboBox()
        {
            List<string[]> lstForUse = new List<string[]>();
            string[] strForUse_1 = { "1", "공용" };
            string[] strForUse_2 = { "2", "설비" };
            string[] strForUse_3 = { "3", "TOOL" };
            lstForUse.Add(strForUse_1);
            lstForUse.Add(strForUse_2);
            lstForUse.Add(strForUse_3);

            ObservableCollection<CodeView> ovcDvlYN = ComboBoxUtil.Instance.Direct_SetComboBox(lstForUse);
            this.cboMCPartTypeGubun.ItemsSource = ovcDvlYN;
            this.cboMCPartTypeGubun.DisplayMemberPath = "code_name";
            this.cboMCPartTypeGubun.SelectedValuePath = "code_id";
        }

        #endregion

        #region 날짜 관련 이벤트

        //입출일자
        private void lblMcInOutDate_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            if (chkMcInOutDate.IsChecked == true) { chkMcInOutDate.IsChecked = false; }
            else { chkMcInOutDate.IsChecked = true; }
        }

        //입출일자
        private void chkMcInOutDate_Checked(object sender, RoutedEventArgs e)
        {
            dtpSDate.IsEnabled = true;
            dtpEDate.IsEnabled = true;
        }

        //입출일자
        private void chkMcInOutDate_Unchecked(object sender, RoutedEventArgs e)
        {
            dtpSDate.IsEnabled = false;
            dtpEDate.IsEnabled = false;
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

        #region 토글 버튼 이벤트 모음

        // 공용 버튼
        private void tgnCommon_Checked(object sender, RoutedEventArgs e)
        {
            tgnCommon.IsChecked = true;
            tgnSpare.IsChecked = false;
            tgnTool.IsChecked = false;
            tgnAll.IsChecked = false;

            //dgtc_MCPartName.Header = "설비(부품)명";
            dgtc_MCPartName.Header = "공용";
            TextBlockMCPartSearch.Text = "공용";
            re_Search(0);

        }

        private void tgnCommon_Unchecked(object sender, RoutedEventArgs e)
        {
            e.Handled = true;
        }

        private void tgnCommon_PreviewMouseDown(object sender, MouseButtonEventArgs e)
        {
            if (tgnCommon.IsChecked == true)
            {
                e.Handled = true;
            }
        }

        // 설비예비품
        private void tgnSpare_Checked(object sender, RoutedEventArgs e)
        {
            tgnCommon.IsChecked = false;
            tgnSpare.IsChecked = true;
            tgnTool.IsChecked = false;
            tgnAll.IsChecked = false;

            dgtc_MCPartName.Header = "예비품명";
            TextBlockMCPartSearch.Text = "예비품명";
            
            re_Search(0);
        }

        private void tgnSpare_Unchecked(object sender, RoutedEventArgs e)
        {
            e.Handled = true;
        }

        private void tgnSpare_PreviewMouseDown(object sender, MouseButtonEventArgs e)
        {
            if (tgnSpare.IsChecked == true)
            {
                e.Handled = true;
            }
        }

        // 툴
        private void tgnTool_Checked(object sender, RoutedEventArgs e)
        {
            tgnCommon.IsChecked = false;
            tgnSpare.IsChecked = false;
            tgnTool.IsChecked = true;
            tgnAll.IsChecked = false;

            dgtc_MCPartName.Header = "Tool 명";
            TextBlockMCPartSearch.Text = "Tool 명";

            re_Search(0);
        }

        private void tgnTool_Unchecked(object sender, RoutedEventArgs e)
        {
            e.Handled = true;
        }

        private void tgnTool_PreviewMouseDown(object sender, MouseButtonEventArgs e)
        {
            if (tgnTool.IsChecked == true)
            {
                e.Handled = true;
            }
        }

        // 전체
        private void tgnAll_Checked(object sender, RoutedEventArgs e)
        {
            tgnCommon.IsChecked = false;
            tgnSpare.IsChecked = false;
            tgnTool.IsChecked = false;
            tgnAll.IsChecked = true;

            dgtc_MCPartName.Header = "품명";
            TextBlockMCPartSearch.Text = "품명";

            re_Search(0);
        }

        private void tgnAll_Unchecked(object sender, RoutedEventArgs e)
        {
            e.Handled = true;
        }

        private void tgnAll_PreviewMouseDown(object sender, MouseButtonEventArgs e)
        {
            if (tgnAll.IsChecked == true)
            {
                e.Handled = true;
            }
        }


        #endregion

        #region 체크 박스 및 플러스 파인더 이벤트

        //부품용도
        private void lblMCPartTypeGubun_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            if (chkMCPartTypeGubun.IsChecked == true) { chkMCPartTypeGubun.IsChecked = false; }
            else { chkMCPartTypeGubun.IsChecked = true; }
        }

        //부품용도
        private void chkMCPartTypeGubun_Checked(object sender, RoutedEventArgs e)
        {
            cboMCPartTypeGubun.IsEnabled = true;
            cboMCPartTypeGubun.Focus();
        }

        //부품용도
        private void chkMCPartTypeGubun_Unchecked(object sender, RoutedEventArgs e)
        {
            cboMCPartTypeGubun.IsEnabled = false;
        }

        //거래처
        private void lblCustomSrh_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            if (chkCustomSrh.IsChecked == true) { chkCustomSrh.IsChecked = false; }
            else { chkCustomSrh.IsChecked = true; }
        }

        //거래처
        private void chkCustomSrh_Checked(object sender, RoutedEventArgs e)
        {
            txtCustomSrh.IsEnabled = true;
            btnPfCustomSrh.IsEnabled = true;
            txtCustomSrh.Focus();
        }

        //거래처
        private void chkCustomSrh_Unchecked(object sender, RoutedEventArgs e)
        {
            txtCustomSrh.IsEnabled = false;
            btnPfCustomSrh.IsEnabled = false;
        }

        //거래처
        private void txtCustomSrh_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                MainWindow.pf.ReturnCode(txtCustomSrh, (int)Defind_CodeFind.DCF_CUSTOM, "");
            }
        }

        //거래처
        private void btnPfCustomSrh_Click(object sender, RoutedEventArgs e)
        {
            MainWindow.pf.ReturnCode(txtCustomSrh, (int)Defind_CodeFind.DCF_CUSTOM, "");
        }

        //부품명
        private void lblMCPart_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            if (chkMCPart.IsChecked == true) { chkMCPart.IsChecked = false; }
            else { chkMCPart.IsChecked = true; }
        }

        //부품명
        private void chkMCPart_Checked(object sender, RoutedEventArgs e)
        {
            txtMCPart.IsEnabled = true;
            btnPfMCPart.IsEnabled = true;
            txtMCPart.Focus();
        }

        //부품명
        private void chkMCPart_Unchecked(object sender, RoutedEventArgs e)
        {
            txtMCPart.IsEnabled = false;
            btnPfMCPart.IsEnabled = false;
        }

        //부품명
        private void txtMCPart_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                MainWindow.pf.ReturnCode(txtMCPart, (int)Defind_CodeFind.DCF_PART, "");
            }
        }

        //부품명
        private void btnPfMCPart_Click(object sender, RoutedEventArgs e)
        {
            MainWindow.pf.ReturnCode(txtMCPart, (int)Defind_CodeFind.DCF_PART, "");
        }

        #endregion

        #region 우측 상단 버튼 클릭

        //조회
        private void btnSearch_Click(object sender, RoutedEventArgs e)
        {
            //검색버튼 비활성화
            btnSearch.IsEnabled = false;

            Dispatcher.BeginInvoke(new Action(() =>

            {
                //로직
                re_Search(0);

            }), System.Windows.Threading.DispatcherPriority.Background);



            Dispatcher.BeginInvoke(new Action(() =>

            {
                btnSearch.IsEnabled = true;

            }), System.Windows.Threading.DispatcherPriority.Background);
            

        }

        //닫기
        private void btnClose_Click(object sender, RoutedEventArgs e)
        {
            Lib.Instance.ChildMenuClose(this.ToString());
        }

        //엑셀
        private void btnExcel_Click(object sender, RoutedEventArgs e)
        {
            DataTable dt = null;
            string Name = string.Empty;

            string[] dgdStr = new string[2];
            dgdStr[0] = "설비부품 수불조회";
            dgdStr[1] = dgdPartSubul.Name;

            ExportExcelxaml ExpExc = new ExportExcelxaml(dgdStr);
            ExpExc.ShowDialog();

            if (ExpExc.DialogResult.HasValue)
            {
                if (ExpExc.choice.Equals(dgdPartSubul.Name))
                {
                    if (ExpExc.Check.Equals("Y"))
                        dt = Lib.Instance.DataGridToDTinHidden(dgdPartSubul);
                    else
                        dt = Lib.Instance.DataGirdToDataTable(dgdPartSubul);

                    Name = dgdPartSubul.Name;
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

        private void re_Search(int selectedIndex)
        {
            FillGrid();

            if (dgdPartSubul.Items.Count > 1)
            {
                dgdPartSubul.SelectedIndex = 0;
            }
            else if (dgdPartSubul.Items.Count == 1)
            {
                MessageBox.Show("조회된 데이터가 없습니다.");
                return;
            }
        }

        //실조회
        private void FillGrid()
        {
            try
            {
                if (dgdPartSubul.Items.Count > 0)
                {
                    dgdPartSubul.Items.Clear();
                }
                dgdTotal.Items.Clear();

                string sForUse = "";
                if (tgnCommon.IsChecked == true) { sForUse = "1"; }
                if (tgnSpare.IsChecked == true) { sForUse = "2"; }
                if (tgnTool.IsChecked == true) { sForUse = "3"; }

                Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                sqlParameter.Add("nChkDate", chkMcInOutDate.IsChecked == true ? 1 : 0);
                sqlParameter.Add("sSDate", chkMcInOutDate.IsChecked == true && dtpSDate.SelectedDate != null ? dtpSDate.SelectedDate.Value.ToString("yyyyMMdd") : "");
                sqlParameter.Add("sEDate", chkMcInOutDate.IsChecked == true && dtpSDate.SelectedDate != null ? dtpEDate.SelectedDate.Value.ToString("yyyyMMdd") : "");
                sqlParameter.Add("nChkCustom", chkCustomSrh.IsChecked == true ? 1 : 0);
                sqlParameter.Add("sCustomID", chkCustomSrh.IsChecked == true && txtCustomSrh.Tag != null ? txtCustomSrh.Tag.ToString() : "");
                sqlParameter.Add("nChkArticleID", chkMCPart.IsChecked == true ? 1 : 0);
                sqlParameter.Add("sArticleID", chkMCPart.IsChecked == true && txtMCPart.Tag != null ? txtMCPart.Tag.ToString() : "");
                sqlParameter.Add("sForUse", sForUse);

                DataSet ds = DataStore.Instance.ProcedureToDataSet("xp_mc_sMcPartSubul", sqlParameter, false);

                if (ds != null && ds.Tables.Count > 0)
                {
                    DataTable dt = ds.Tables[0];

                    if (dt.Rows.Count > 0)
                    {
                        DataRowCollection drc = dt.Rows;

                        for (int i = 0; i < drc.Count; i++)
                        {
                            DataRow dr = drc[i];

                            var WinMcSubul = new Win_prd_PartSubul_Q_CodeView()
                            {
                                Num = (i + 1).ToString(),
                                cls = dr["cls"].ToString(),
                                MCPartID = dr["MCPartID"].ToString(),
                                MCPartName = dr["MCPartName"].ToString(),
                                IODate = dr["IODate"].ToString(),
                                IODate_CV = DatePickerFormat(dr["IODate"].ToString()),
                                StuffRoll = Convert.ToDouble(dr["StuffRoll"]),
                                RemainQty = Convert.ToDouble(dr["RemainQty"]),    //20210527 이월 재고 따로 보이기
                                StuffQty = Convert.ToDouble(dr["StuffQty"]),
                                Unitclss = dr["Unitclss"].ToString(),
                                UnitClssName = dr["UnitClssName"].ToString(),
                                CustomID = dr["CustomID"].ToString(),
                                CustomName = dr["CustomName"].ToString(),
                                OutRoll = Convert.ToDouble(dr["OutRoll"]),
                                OutQty = Convert.ToDouble(dr["OutQty"]),
                                Remark = dr["Remark"].ToString(),
                                StockQty = Convert.ToDouble(dr["StockQty"]),
                                ForUse = dr["ForUse"].ToString(),
                                ForUseName = dr["ForUseName"].ToString()
                            };

                            if (WinMcSubul.cls.Equals("6"))
                            {
                                
                                WinMcSubul.MCPartName = "총계";
                                WinMcSubul.IODate_CV = "";
                                dgdTotal.Items.Add(WinMcSubul);
                            }
                            else if (WinMcSubul.cls.Equals("5"))
                            {
                                WinMcSubul.MCPartName = "부품별 집계";
                                WinMcSubul.IODate_CV = "";
                                dgdPartSubul.Items.Add(WinMcSubul);
                            }
                            else if (WinMcSubul.cls.Equals("0"))
                            {
                                WinMcSubul.MCPartName = "이월";
                                WinMcSubul.IODate_CV = "";
                                dgdPartSubul.Items.Add(WinMcSubul);
                            }
                            else
                            {
                                dgdPartSubul.Items.Add(WinMcSubul);
                            }

                            //if (WinMcSubul.cls.Equals("1")|| WinMcSubul.cls.Equals("5") || WinMcSubul.cls.Equals("6"))
                            
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

        #region 기타 메서드 모음

        // 텍스트 박스 숫자만 입력 되도록
        public void CheckIsNumericOnly(TextBox sender, TextCompositionEventArgs e)
        {
            decimal result;
            if (!(Decimal.TryParse(e.Text, out result)))
            {
                e.Handled = true;
            }
        }

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

        // 확장자 이미지 확인하기, 메인윈도우에 확장자 리스트 세팅
        private bool CheckImage(string ImageName)
        {
            string[] extensions = MainWindow.Extensions;

            bool flag = false;

            ImageName = ImageName.Trim().ToLower();
            foreach (string ext in extensions)
            {
                if (ImageName.EndsWith(ext))
                {
                    flag = true;
                }
            }

            return flag;
        }



        #endregion

        private void columnHeader_Click(object sender, RoutedEventArgs e)
        {
            foreach (DataGridColumn column in dgdPartSubul.Columns)
            {
                column.CanUserSort = false;
            }

        }

    }

    class Win_prd_PartSubul_Q_CodeView : BaseView
    {
        public override string ToString()
        {
            return (this.ReportAllProperties());
        }

        public string Num { get; set; }
        public string cls { get; set; }
        public string MCPartID { get; set; }
        public string MCPartName { get; set; }
        public string IODate { get; set; }
        public string IODate_CV { get; set; }
        public double StuffRoll { get; set; }
        public double RemainQty { get; set; }   //2021-05-27 이월 재고 따로 보이게 하기 위해 추가    
        public double StuffQty { get; set; }
        public string Unitclss { get; set; }
        public string UnitClssName { get; set; }
        public string CustomID { get; set; }
        public string CustomName { get; set; }
        public double OutRoll { get; set; }
        public double OutQty { get; set; }
        public string Remark { get; set; }
        public double StockQty { get; set; }
        public string ForUse { get; set; }
        public string ForUseName { get; set; }
    }
}

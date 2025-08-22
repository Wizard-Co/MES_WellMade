using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Data;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using WizMes_WellMade.PopUP;

namespace WizMes_WellMade
{
    /// <summary>
    /// Win_Qul_DateProd_Q.xaml에 대한 상호 작용 논리
    /// </summary>
    public partial class Win_Qul_DateProd_Q : UserControl
    {
        int rowNum = 0;
        Lib lib = new Lib();
        PlusFinder pf = new PlusFinder();

        public Win_Qul_DateProd_Q()
        {
            InitializeComponent();
        }

        private void UserControl_Loaded(object sender, RoutedEventArgs e)
        {
            Lib.Instance.UiLoading(sender);

            //콤보박스 셋팅       
            SetComboBox();

            //입고일자 체크
            chkDate.IsChecked = true;

            //데이트피커 오늘 날짜
            dtpSDate.SelectedDate = DateTime.Today;
            dtpEDate.SelectedDate = DateTime.Today;

            //콤보박스 기본값 '전체'
            cboProcessIDSrh.SelectedIndex = 0;

        }

        //콤보박스 셋팅
        private void SetComboBox()
        {
            //공정
            ObservableCollection<CodeView> cboProcessGroup = ComboBoxUtil.Instance.GetWorkProcess(0, "");
            this.cboProcessIDSrh.ItemsSource = cboProcessGroup;
            this.cboProcessIDSrh.DisplayMemberPath = "code_name";
            this.cboProcessIDSrh.SelectedValuePath = "code_id";
            cboProcessIDSrh.SelectedIndex = 0;


            ObservableCollection<CodeView> cboMachineGroup = GetMachineByProcessID("");
            this.cboMachineIDSrh.ItemsSource = cboMachineGroup;
            this.cboMachineIDSrh.DisplayMemberPath = "code_name";
            this.cboMachineIDSrh.SelectedValuePath = "code_id";
            cboMachineIDSrh.SelectedIndex = 0;


        }

        #region 클릭 이벤트

        //입고일자 라벨 클릭 이벤트
        private void LblchkDay_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            if (chkDate.IsChecked == true)
            {
                chkDate.IsChecked = false;
                dtpSDate.IsEnabled = false;
                dtpEDate.IsEnabled = false;
            }
            else
            {
                chkDate.IsChecked = true;
                dtpSDate.IsEnabled = true;
                dtpEDate.IsEnabled = true;
            }
        }

        //입고일자 체크 이벤트
        private void ChkDate_Checked(object sender, RoutedEventArgs e)
        {
            dtpSDate.IsEnabled = true;
            dtpEDate.IsEnabled = true;
        }

        //입고일자 체크해제 이벤트
        private void ChkDate_Unchecked(object sender, RoutedEventArgs e)
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



   

       
    

        #region 플러스파인더

        // 플러스파인더 _ 품명 찾기
        private void btnArticleIDSrh_Click(object sender, RoutedEventArgs e)
        {
            pf.ReturnCode(txtArticleIDSrh, 77, txtArticleIDSrh.Text);
        }

        // 품명 키다운 _ 품명 찾기
        private void txtArticleIDSrh_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                pf.ReturnCode(txtArticleIDSrh, 77, txtArticleIDSrh.Text);
            }
        }

        // 플러스파인더 _ 품번 찾기
        private void btnBuyerAritlceNoSrh_Click(object sender, RoutedEventArgs e)
        {
            pf.ReturnCode(txtBuyerArticleNoSrh, 76, txtBuyerArticleNoSrh.Text);
        }

        // 품번 키다운 
        private void txtBuyerArticleNoSrh_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                pf.ReturnCode(txtBuyerArticleNoSrh, 76, txtBuyerArticleNoSrh.Text);
            }
        }

      
        #endregion

        #endregion 클릭이벤트, 날짜

        #region CRUD 버튼

        //검색(조회)
        private void btnSearch_Click(object sender, RoutedEventArgs e)
        {
            if(lib.DatePickerCheck(dtpSDate, dtpEDate, chkDate))
            {
                //검색버튼 비활성화
                btnSearch.IsEnabled = false;

                Dispatcher.BeginInvoke(new Action(() =>

                {
                    Thread.Sleep(2000);

                    //로직
                    if (CheckData())
                    {
                        re_Search(rowNum);
                    }

                }), System.Windows.Threading.DispatcherPriority.Background);

                Dispatcher.BeginInvoke(new Action(() =>

                {
                    btnSearch.IsEnabled = true;

                }), System.Windows.Threading.DispatcherPriority.Background);
            }
  
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
            Lib lib = new Lib();

            string[] lst = new string[2];
            lst[0] = "공정불량현황";
            lst[1] = dgdMain.Name;

            ExportExcelxaml ExpExc = new ExportExcelxaml(lst);

            ExpExc.ShowDialog();

            if (ExpExc.DialogResult.HasValue)
            {
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
            lib = null;
        }

        #endregion CRUD 버튼


        #region 데이터그리드 이벤트

        //데이터그리드 셀렉션체인지드
        private void dgdMain_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            //조회만 하는 화면이라 이 친구는 필요가 없지요.
        }

        #endregion 데이터그리드 이벤트

        #region 조회관련(Fillgrid)

        //재조회
        private void re_Search(int selectedIndex)
        {
            FillGrid();

            if (dgdMain.Items.Count > 0)
            {
                dgdMain.SelectedIndex = selectedIndex;
            }
        }

        //조회
        private void FillGrid()
        {
            dgdMain.Items.Clear();
            dgdTotal.Items.Clear();

            try
            {             

                Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                sqlParameter.Clear();

                sqlParameter.Add("chkDate", chkDate.IsChecked == true ? 1 : 0);
                sqlParameter.Add("sDate", chkDate.IsChecked == true ? dtpSDate.SelectedDate.Value.ToString("yyyyMMdd") : "");
                sqlParameter.Add("eDate", chkDate.IsChecked == true ? dtpEDate.SelectedDate.Value.ToString("yyyyMMdd") : "");

                sqlParameter.Add("chkProcessID", chkProcessIDSrh.IsChecked == true ? 1 : 0);
                sqlParameter.Add("ProcessID", chkProcessIDSrh.IsChecked == true ? cboProcessIDSrh.SelectedValue?.ToString() ?? string.Empty : string.Empty);

                sqlParameter.Add("chkArticleID", chkArticleIDSrh.IsChecked == true ? 1:0);
                sqlParameter.Add("ArticleID", chkArticleIDSrh.IsChecked == true? txtArticleIDSrh.Tag?.ToString() ?? string.Empty : string.Empty);

                sqlParameter.Add("chkBuyerArticleNo", chkBuyerArticleNoSrh.IsChecked == true ? 1:0);
                sqlParameter.Add("BuyerArticleNo", chkBuyerArticleNoSrh.IsChecked == true ? txtBuyerArticleNoSrh.Tag?.ToString() ?? string.Empty : string.Empty);

                sqlParameter.Add("chkMachineID", chkMachineIDSrh.IsChecked == true ? 1 : 0);
                sqlParameter.Add("MachineID", chkMachineIDSrh.IsChecked == true ? cboMachineIDSrh.SelectedValue?.ToString() ?? string.Empty : string.Empty);

                sqlParameter.Add("chkExceptInitialDefect", chkExceptInitialDefectSrh.IsChecked == true ? 1 : 0);

                DataSet ds = DataStore.Instance.ProcedureToDataSet("xp_Inspect_sInspectDefect", sqlParameter, false);

                if (ds != null && ds.Tables.Count > 0)
                {
                    DataTable dt = ds.Tables[0];

                    if (dt.Rows.Count == 0)
                    {
                        MessageBox.Show("조회결과가 없습니다.");
                        return;
                    }
                    else
                    {
                        DataRowCollection drc = dt.Rows;

                        int i = 0;
                        int TotalrowCount = 0;
                        int TotalDefectQty = 0;
                        int InitialDefectQty = 0;
                        foreach (DataRow dr in drc)
                        {
                            i++;
                            var DefectInfo = new Win_Qul_DateProd_Q_CodeView()
                            {
                                Num = i,
                                gbn = dr["gbn"].ToString(),
                                ScanDate = lib.DateTypeHyphen(dr["ScanDate"].ToString()),
                                ProcessID = dr["ProcessID"].ToString(),
                                Process = dr["Process"].ToString(),
                                BuyerModelID = dr["BuyerModelID"].ToString(),
                                ArticleID = dr["ArticleID"].ToString(),
                                Article = dr["Article"].ToString(),
                                BuyerArticleNo = dr["BuyerArticleNo"].ToString(),
                                DefectID = dr["DefectID"].ToString(),
                                KDefect = dr["KDefect"].ToString(),
                                DefectQty = stringFormatN0(dr["DefectQty"]),
                                WorkPersonID = dr["WorkPersonID"].ToString(),
                                WorkPersonName = dr["WorkPersonName"].ToString(),
                                WorkQty = stringFormatN0(dr["WorkQty"]),
                                MCNAME = dr["MCNAME"].ToString()
                                //LabelID = dr["LabelID"].ToString(),
                                //ChildLabelID = dr["ChildLabelID"].ToString()
                            };

                            if(DefectInfo.gbn.Equals("1") || DefectInfo.Equals("2"))
                            {
                                TotalrowCount++;
                                TotalDefectQty += lib.RemoveComma(dr["DefectQty"].ToString(), 0);

                                if (DefectInfo.KDefect.Contains("초도"))
                                    InitialDefectQty += lib.RemoveComma(dr["DefectQty"].ToString(), 0);

                                dgdMain.Items.Add(DefectInfo);
                            }
                            else if (DefectInfo.gbn.Equals("3"))
                            {
                                DefectInfo.Color1 = true;
                                dgdMain.Items.Add(DefectInfo);
                            }

                        }

                        if(dgdMain.Items.Count > 0)
                        {
                            var DefectInfo_Total = new Win_Qul_DateProd_Q_CodeView_Total()
                            {
                                TotalCount = TotalrowCount,
                                TotalDefectQty = stringFormatN0(TotalDefectQty),
                                TotalInitialDefectQty = stringFormatN0(InitialDefectQty)
                            };

                            dgdTotal.Items.Add(DefectInfo_Total);
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

        //검색 조건 Check
        private bool CheckData()
        {
            bool flag = true;

     

            return flag;
        }


        #endregion 조회관련(Fillgrid)

        #region 기타 메소드 
        //특수문자 포함 검색
        private string Escape(string str)
        {
            string result = "";

            for (int i = 0; i < str.Length; i++)
            {
                string txt = str.Substring(i, 1);

                bool isSpecial = Regex.IsMatch(txt, @"[^a-zA-Z0-9가-힣]");

                if (isSpecial == true)
                {
                    result += (@"/" + txt);
                }
                else
                {
                    result += txt;
                }
            }
            return result;
        }

        // 천단위 콤마, 소수점 버리기
        private string stringFormatN0(object obj)
        {
            return string.Format("{0:N0}", obj);
        }
        #endregion

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

        private void CommonControl_Click(object sender, MouseButtonEventArgs e)
        {
            lib.CommonControl_Click(sender, e);
        }

        private void CommonControl_Click(object sender, RoutedEventArgs e)
        {
            lib.CommonControl_Click(sender, e);
        }

        private void cboProcessIDSrh_SelectionChagned(object sender, SelectionChangedEventArgs e)
        {
            if (cboProcessIDSrh.SelectedValue != null)
            {
                ObservableCollection<CodeView> ovcMachine = GetMachineByProcessID(cboProcessIDSrh.SelectedValue?.ToString() ?? string.Empty);
                var filteredMachines = ovcMachine.Where(x => !string.IsNullOrEmpty(x.code_id)).ToList();
                this.cboMachineIDSrh.ItemsSource = filteredMachines;
                this.cboMachineIDSrh.DisplayMemberPath = "code_name";
                this.cboMachineIDSrh.SelectedValuePath = "code_id";
                if (ovcMachine.Count > 0)
                {
                    cboMachineIDSrh.SelectedIndex = 0;
                }

            }   
        }

        /// <summary>
        /// 호기ID 가져오기
        /// </summary>
        /// <param name="value"></param>
        /// <returns></returns>
        public ObservableCollection<CodeView> GetMachineByProcessID(string value)
        {
            ObservableCollection<CodeView> ovcMachine = new ObservableCollection<CodeView>();

            Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
            sqlParameter.Add("sProcessID", value);

            DataSet ds = DataStore.Instance.ProcedureToDataSet("xp_prd_sMachineForComboBoxAndUsing", sqlParameter, false);

            if (ds != null && ds.Tables.Count > 0)
            {
                DataTable dt = ds.Tables[0];
                if (dt.Rows.Count > 0)
                {
                    CodeView CV = new CodeView();
                    CV.code_id = "";
                    CV.code_name = "전체";
                    ovcMachine.Add(CV);

                    DataRowCollection drc = dt.Rows;

                    foreach (DataRow dr in drc)
                    {
                        CodeView mCodeView = new CodeView()
                        {
                            code_id = dr["Code"].ToString().Trim(),
                            code_name = dr["Name"].ToString().Trim()
                        };

                        ovcMachine.Add(mCodeView);
                    }
                }
            }

            return ovcMachine;
        }
    }

    #region 생성자들(CodeView)

    class Win_Qul_DateProd_Q_CodeView : BaseView
    {
        public int Num { get; set; }
        public string gbn { get; set; }
        public string ScanDate { get; set; }
        public string ProcessID { get; set; }
        public string Process { get; set; }
        public string BuyerModelID { get; set; }
        public string ArticleID { get; set; }
        public string Article { get; set; }
        public string BuyerArticleNo { get; set; }
        public string DefectID { get; set; }
        public string KDefect { get; set; }
        public string DefectQty { get; set; }
        public string WorkPersonID { get; set; }
        public string WorkPersonName { get; set; }
        public string WorkQty { get; set; }
        public string MCNAME { get; set; }

        //public string LabelID { get; set; }
        //public string ChildLabelID { get; set; }
        //public string ColorLightLightGray { get; set; }
        //public string ColorGold { get; set; }
        public bool Color1 { get; set; } = false;
        public bool Color2 { get; set; } = false;
    }

    public class Win_Qul_DateProd_Q_CodeView_Total : BaseView
    {
        public int TotalCount { get; set; }
        public string TotalDefectQty { get; set; }
        public string TotalInitialDefectQty { get; set; }
    }

    #endregion 생성자들(CodeView)
}
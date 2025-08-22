using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Data;
using System.Threading;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using WizMes_WellMade.PopUP;

namespace WizMes_WellMade
{
    /// <summary>
    /// Win_Qul_4MChange_U.xaml에 대한 상호 작용 논리
    /// </summary>
    public partial class Win_Qul_4MChange_U : UserControl
    {
        #region 변수선언 및 로드

        Win_Qul_4MChange_NU_CodeView win4MC = new Win_Qul_4MChange_NU_CodeView();
        Win_Qul_4MChangeNUSub_CodeView Win4MSub = new Win_Qul_4MChangeNUSub_CodeView();
        Lib lib = new Lib();
        PlusFinder pf = new PlusFinder();
        int numSaveRowCount = 0;
        string strFlag = string.Empty;
        string strArticleID = string.Empty;
        int dgdInComboNum = 0;
        string strDGinCombo1 = string.Empty;
        string strDGinCombo2 = string.Empty;
        Dictionary<string, object> dicCompare = new Dictionary<string, object>();
        List<string> lstCompareValue = new List<string>();
        ObservableCollection<CodeView> ovcBAN = new ObservableCollection<CodeView>();
        ObservableCollection<CodeView> ovcBMI = new ObservableCollection<CodeView>();
        ObservableCollection<CodeView> ovcModel = new ObservableCollection<CodeView>();

        //품명가져오기 변수
        string tempArticle = string.Empty;

        public Win_Qul_4MChange_U()
        {
            InitializeComponent();
        }

        private void UserControl_Loaded(object sender, RoutedEventArgs e)
        {
            chkDaySrh.IsChecked = true;
            dtpSDate.SelectedDate = lib.BringThisMonthDatetimeList()[0];
            dtpEDate.SelectedDate = lib.BringThisMonthDatetimeList()[1];
            SetComboBox();

            if (dgdMain.Items.Count > 0)
            {
                dgdMain.Items.Clear();
            }

            lib.UiLoading(sender);
        }

        private void SetComboBox()
        {
            string[] str1 = { "1", "내부" };
            string[] str2 = { "2", "외부" };
            List<string[]> listStr = new List<string[]>();
            listStr.Add(str1);
            listStr.Add(str2);

            ObservableCollection<CodeView> ovcInOut = ComboBoxUtil.Instance.Direct_SetComboBox(listStr);
            this.cboInOutGbn.ItemsSource = ovcInOut;
            this.cboInOutGbn.DisplayMemberPath = "code_name";
            this.cboInOutGbn.SelectedValuePath = "code_id";

            ObservableCollection<CodeView> ovcChangeGBN = ComboBoxUtil.Instance.Gf_DB_CM_GetComCodeDataset(null, "CHANGEGBN", "Y", "", "");
            cboChangeGBN.ItemsSource = ovcChangeGBN;
            cboChangeGBN.DisplayMemberPath = "code_name";
            cboChangeGBN.SelectedValuePath = "code_id";

            cboChangeGBNSrh.ItemsSource = ovcChangeGBN;
            cboChangeGBNSrh.DisplayMemberPath = "code_name";
            cboChangeGBNSrh.SelectedValuePath = "code_id";

            cboBuyerModel.ItemsSource = ovcModel;
        }

        #endregion

        //Main DataGrid RowHeader
        private void dgdMain_LoadingRow(object sender, DataGridRowEventArgs e)
        {
            e.Row.Header = (e.Row.GetIndex() + 1).ToString();
        }

        #region 상단 날짜관련 버튼 클릭 이벤트

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

        #region 상단 체크박스 이벤트

        //라벨 일자 클릭
        private void lblDaySrh_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            if (chkDaySrh.IsChecked == true) { chkDaySrh.IsChecked = false; }
            else { chkDaySrh.IsChecked = true; }
        }

        //체크 일자 클릭
        private void chkDaySrh_Checked(object sender, RoutedEventArgs e)
        {
            dtpSDate.IsEnabled = true;
            dtpEDate.IsEnabled = true;
        }

        //체크 일자 클릭
        private void chkDaySrh_Unchecked(object sender, RoutedEventArgs e)
        {
            dtpSDate.IsEnabled = false;
            dtpEDate.IsEnabled = false;
        }

        //라벨 고객사 클릭
        private void lblCustomSrh_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            if (chkCustomSrh.IsChecked == true) { chkCustomSrh.IsChecked = false; }
            else { chkCustomSrh.IsChecked = true; }
        }

        //체크 고객사 클릭
        private void chkCustomSrh_Checked(object sender, RoutedEventArgs e)
        {
            txtCustomSrh.IsEnabled = true;
            btnPfCustomSrh.IsEnabled = true;
            txtCustomSrh.Focus();
        }

        //체크 고객사 클릭
        private void chkCustomSrh_Unchecked(object sender, RoutedEventArgs e)
        {
            txtCustomSrh.IsEnabled = false;
            btnPfCustomSrh.IsEnabled = false;
        }

        //조회 거래처 텍스트박스 플러스파인더
        private void txtCustomSrh_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                pf.ReturnCode(txtCustomSrh, 0, "");
            }
        }

        //조회 거래처 플러스파인더
        private void btnPfCustomSrh_Click(object sender, RoutedEventArgs e)
        {
            pf.ReturnCode(txtCustomSrh, 0, "");
        }

        //라벨 품명 클릭
        private void lblArticleSrh_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            if (chkArticleSrh.IsChecked == true) { chkArticleSrh.IsChecked = false; }
            else { chkArticleSrh.IsChecked = true; }
        }

        //체크 품명 클릭
        private void chkArticleSrh_Checked(object sender, RoutedEventArgs e)
        {
            txtArticleSrh.IsEnabled = true;
            btnPfArticleSrh.IsEnabled = true;
            txtArticleSrh.Focus();
        }

        //체크 품명 클릭
        private void chkArticleSrh_Unchecked(object sender, RoutedEventArgs e)
        {
            txtArticleSrh.IsEnabled = false;
            btnPfArticleSrh.IsEnabled = false;
        }

        //조회 품명 텍스트박스 플러스파인더(품번으로 수정요청, 2020.03.19, 장가빈)
        private void txtArticleSrh_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                //pf.ReturnCode(txtArticleSrh, 1, "");
                //pf.ReturnCode(txtArticle, 71, "", "", "", txtSabun, txtBuyerArticleNo);
                pf.ReturnCode(txtArticleSrh, 84, txtArticleSrh.Text);
            }
        }

        //조회 품명 플러스파인더(품번으로 수정요청, 2020.03.19, 장가빈)
        private void btnPfArticleSrh_Click(object sender, RoutedEventArgs e)
        {
            //pf.ReturnCode(txtArticleSrh, 1, "");
            //pf.ReturnCode(txtArticle, 71, "", "", "", txtSabun, txtBuyerArticleNo);
            pf.ReturnCode(txtArticleSrh, 84, txtArticleSrh.Text);
        }

        //라벨 신고제목 클릭
        private void lbl4MSubjectSrh_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            if (chk4MSubjectSrh.IsChecked == true) { chk4MSubjectSrh.IsChecked = false; }
            else { chk4MSubjectSrh.IsChecked = true; }
        }

        //체크 신고제목 클릭
        private void chk4MSubjectSrh_Checked(object sender, RoutedEventArgs e)
        {
            txt4MSubjectSrh.IsEnabled = true;
            txt4MSubjectSrh.Focus();
        }

        //체크 신고제목 클릭
        private void chk4MSubjectSrh_Unchecked(object sender, RoutedEventArgs e)
        {
            txt4MSubjectSrh.IsEnabled = false;
        }

        //변경구분 클릭
        private void LblChangeGBN_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            if (chkChangeGBN.IsChecked == true) { chkChangeGBN.IsChecked = false; }
            else { chkChangeGBN.IsChecked = true; }
        }

        //변경구분 클릭
        private void ChkChangeGBN_Checked(object sender, RoutedEventArgs e)
        {
            cboChangeGBNSrh.IsEnabled = true;
            cboChangeGBNSrh.Focus();
        }

        //변경구분 클릭
        private void ChkChangeGBN_Unchecked(object sender, RoutedEventArgs e)
        {
            cboChangeGBNSrh.IsEnabled = false;
        }

        #endregion

        #region 상단 우측 버튼 클릭 이벤트

        //조회
        private void btnSearch_Click(object sender, RoutedEventArgs e)
        {
            //검색버튼 비활성화
            btnSearch.IsEnabled = false;

            Dispatcher.BeginInvoke(new Action(() =>

            {
                Thread.Sleep(2000);

                //로직
                if (dgdMain.Items.Count > 0)
                {
                    dgdMain.Items.Clear();
                }

                FillGridMain();

                if (dgdMain.Items.Count > 0)
                {
                    if (lstCompareValue.Count > 0)
                    {
                        dgdMain.SelectedIndex = lib.reTrunIndex(dgdMain, lstCompareValue[0]);
                    }
                    else
                    {
                        dgdMain.SelectedIndex = 0;
                    }
                }
                else
                {
                    InputClear();
                }

                dicCompare.Clear();
                lstCompareValue.Clear();

            }), System.Windows.Threading.DispatcherPriority.Background);



            Dispatcher.BeginInvoke(new Action(() =>

            {
                btnSearch.IsEnabled = true;

            }), System.Windows.Threading.DispatcherPriority.Background);



        }

        //추가
        private void btnAdd_Click(object sender, RoutedEventArgs e)
        {
            VisibleSaveCancelETC();

            strFlag = "I";  //저장 시 추가,수정 구분용
            this.DataContext = null;

            if (dgdSub_Article.Items.Count > 0)
            {
                dgdSub_Article.Items.Clear();
            }
            if (dgdValidation.Items.Count > 0)
            {
                dgdValidation.Items.Clear();

                var Valid = new Win_Qul_4MChangeNUValid_CodeView()
                {
                    GubunName = "초기유동",
                    Valid1 = "",
                    Valid2 = "",
                    Valid3 = ""
                };

                dgdValidation.Items.Add(Valid);

                var ValidC = new Win_Qul_4MChangeNUValid_CodeView()
                {
                    GubunName = "확인결과",
                    Valid1 = "",
                    Valid2 = "",
                    Valid3 = ""
                };

                dgdValidation.Items.Add(ValidC);

            }

            numSaveRowCount = dgdMain.SelectedIndex;    //취소시 조회를 위해
            dtpECODate.SelectedDate = DateTime.Today;

            //날짜 들어가는 부분은 추가 버튼 클릭시 다 오늘 날짜 기본 셋팅 해줘야 할 것 같아서..
            dtpCustomApprovementDate.SelectedDate = DateTime.Today;
            dtpInHouseExpectedDate.SelectedDate = DateTime.Today;
            dtpCustomExpectedDate.SelectedDate = DateTime.Today;
            dtpISIRDate.SelectedDate = DateTime.Today;
            dtpInHouseApplicationDate.SelectedDate = DateTime.Today;
            dtpCustomApplicationDate.SelectedDate = DateTime.Today;


            //cboInOutGbn.Focus();
            //cboInOutGbn.IsDropDownOpen = true;
        }

        //수정
        private void btnUpdate_Click(object sender, RoutedEventArgs e)
        {
            win4MC = dgdMain.SelectedItem as Win_Qul_4MChange_NU_CodeView;

            if (win4MC == null)
            {
                MessageBox.Show("수정할 자료를 선택하고 눌러주십시오.");
            }
            else
            {
                VisibleSaveCancelETC();

                strFlag = "U";  //저장 시 추가,수정 구분용
                numSaveRowCount = dgdMain.SelectedIndex;    //취소시 조회를 위해

                //cboInOutGbn.Focus();
                //cboInOutGbn.IsDropDownOpen = true;
            }
        }

        //삭제
        private void btnDelete_Click(object sender, RoutedEventArgs e)
        {
            win4MC = dgdMain.SelectedItem as Win_Qul_4MChange_NU_CodeView;
            numSaveRowCount = dgdMain.SelectedIndex;

            if (win4MC == null)
            {
                MessageBox.Show("삭제할 데이터가 지정되지 않았습니다. 삭제데이터를 지정하고 눌러주세요");
                return;
            }
            else
            {
                if (MessageBox.Show("선택하신 항목을 삭제하시겠습니까?", "삭제 전 확인", MessageBoxButton.YesNo) == MessageBoxResult.Yes)
                {
                    if (DeleteData(win4MC.FourMID)) //삭제 성공
                    {
                        Check_dgdMain_SelectIndex(numSaveRowCount - 1);
                    }
                }
            }
        }

        //저장
        private void btnSave_Click(object sender, RoutedEventArgs e)
        {
            if (strFlag.Equals("I"))    //추가
            {
                if (SaveData(""))
                {
                    HiddenSavaCancelETC();

                    numSaveRowCount = 0;
                    Check_dgdMain_SelectIndex(numSaveRowCount);
                }
            }
            else   //수정
            {
                if (SaveData(txt4MID.Text))
                {
                    HiddenSavaCancelETC();

                    Check_dgdMain_SelectIndex(numSaveRowCount);
                }
            }
        }

        //취소
        private void btnCancel_Click(object sender, RoutedEventArgs e)
        {
            InputClear();
            HiddenSavaCancelETC();

            Check_dgdMain_SelectIndex(numSaveRowCount);
        }

        //입력 데이터 클리어
        private void InputClear()
        {
            //foreach (Control child in this.grdR1.Children)
            //{
            //    if (child.GetType() == typeof(TextBox))
            //        ((TextBox)child).Clear();
            //}
            //foreach (Control child in this.grdChk1.Children)
            //{
            //    if (child.GetType() == typeof(CheckBox))
            //        ((CheckBox)child).IsChecked = new bool?(false);
            //}
            //foreach (Control child in this.grdChk2.Children)
            //{
            //    if (child.GetType() == typeof(CheckBox))
            //        ((CheckBox)child).IsChecked = new bool?(false);
            //}
            //foreach (Control child in this.grdChk3.Children)
            //{
            //    if (child.GetType() == typeof(CheckBox))
            //        ((CheckBox)child).IsChecked = new bool?(false);
            //}
            //if (this.dgdSub_Article.Items.Count > 0)
            //    this.dgdSub_Article.Items.Clear();
            //this.txt4MContents.Clear();
        }

        //저장,취소 시 각 변화
        private void HiddenSavaCancelETC()
        {
            lib.UiButtonEnableChange_IUControl(this);
            dgdMain.IsHitTestVisible = true;
            bdrInput.IsHitTestVisible = false;
        }

        //추가,수정 시 각 변화
        private void VisibleSaveCancelETC()
        {
            lib.UiButtonEnableChange_SCControl(this);
            dgdMain.IsHitTestVisible = false;
            bdrInput.IsHitTestVisible = true;
        }

        //추가,수정,삭제 후 데이터 재조회
        private void Check_dgdMain_SelectIndex(int count)
        {
            if (dgdMain.Items.Count > 0)
            {
                dgdMain.Items.Clear();
            }

            FillGridMain();

            if (dgdMain.Items.Count > 0)
            {
                if (lstCompareValue.Count > 0)
                {
                    dgdMain.SelectedIndex = lib.reTrunIndex(dgdMain, lstCompareValue[0]);
                }
                else
                {
                    dgdMain.SelectedIndex = count; ;
                }
            }

            dicCompare.Clear();
            lstCompareValue.Clear();
        }

        //엑셀
        private void btnExcel_Click(object sender, RoutedEventArgs e)
        {
            DataTable dt = null;
            string Name = string.Empty;
            Lib lib2 = new Lib();

            string[] dgdStr = new string[6];
            dgdStr[0] = "4M변경신고";
            dgdStr[1] = "유효성점검";
            dgdStr[2] = "적용품목";
            dgdStr[3] = dgdMain.Name;
            dgdStr[4] = dgdValidation.Name;
            dgdStr[5] = dgdSub_Article.Name;

            ExportExcelxaml ExpExc = new ExportExcelxaml(dgdStr);
            ExpExc.ShowDialog();

            if (ExpExc.DialogResult.HasValue)
            {
                if (ExpExc.choice.Equals(dgdMain.Name))
                {
                    if (ExpExc.Check.Equals("Y"))
                        dt = lib2.DataGridToDTinHidden(dgdMain);
                    else
                        dt = lib2.DataGirdToDataTable(dgdMain);

                    Name = dgdMain.Name;

                    if (lib2.GenerateExcel(dt, Name))
                    {
                        lib2.excel.Visible = true;
                        lib2.ReleaseExcelObject(lib2.excel);
                    }
                    else
                        return;
                }
                else if (ExpExc.choice.Equals(dgdValidation.Name))
                {
                    if (ExpExc.Check.Equals("Y"))
                        dt = lib2.DataGridToDTinHidden(dgdValidation);
                    else
                        dt = lib2.DataGirdToDataTable(dgdValidation);

                    Name = dgdValidation.Name;
                    if (lib2.GenerateExcel(dt, Name))
                    {
                        lib2.excel.Visible = true;
                        lib2.ReleaseExcelObject(lib2.excel);
                    }
                    else
                        return;
                }
                else if (ExpExc.choice.Equals(dgdSub_Article.Name))
                {
                    if (ExpExc.Check.Equals("Y"))
                        dt = lib2.DataGridToDTinHidden(dgdSub_Article);
                    else
                        dt = lib2.DataGirdToDataTable(dgdSub_Article);

                    Name = dgdSub_Article.Name;
                    if (lib2.GenerateExcel(dt, Name))
                    {
                        lib2.excel.Visible = true;
                        lib2.ReleaseExcelObject(lib2.excel);
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
            lib2 = null;
        }

        //닫기
        private void btnClose_Click(object sender, RoutedEventArgs e)
        {
            lib.ChildMenuClose(this.ToString());
        }


        #endregion

        #region CRUD

        //Main 실질 조회
        private void FillGridMain()
        {
            if (chkCustomSrh.IsChecked == true)
            {
                if (txtCustomSrh.Tag == null || txtCustomSrh.Tag.ToString().Equals(""))
                {
                    MessageBox.Show("조회의 고객사 정보가 잘못되었습니다. 다시 확인해보시기 바랍니다.");
                    return;
                }
            }

            //if (chkArticleSrh.IsChecked == true)
            //{
            //    if (txtArticleSrh.Tag == null || txtArticleSrh.Tag.ToString().Equals(""))
            //    {
            //        MessageBox.Show("조회의 품명 정보가 잘못되었습니다. 다시 확인해보시기 바랍니다.");
            //        return;
            //    }
            //}

            try
            {
                Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                sqlParameter.Add("nchkDate", (chkDaySrh.IsChecked == true ? 1 : 0));
                sqlParameter.Add("FromDate", (chkDaySrh.IsChecked == true ? dtpSDate.SelectedDate.Value.ToString("yyyyMMdd") : ""));
                sqlParameter.Add("ToDate", (chkDaySrh.IsChecked == true ? dtpEDate.SelectedDate.Value.ToString("yyyyMMdd") : ""));
                sqlParameter.Add("nchkCustomID", (chkCustomSrh.IsChecked == true ? (txtCustomSrh.Tag != null ? 1 : 2) : 0));
                sqlParameter.Add("CustomID", (chkCustomSrh.IsChecked == true ? (txtCustomSrh.Tag != null ? txtCustomSrh.Tag.ToString() : txtCustomSrh.Text) : ""));
                sqlParameter.Add("nchkArticleID", (chkArticleSrh.IsChecked == true ? 1 : 0));
                sqlParameter.Add("ArticleID", (chkArticleSrh.IsChecked == true ? txtArticleSrh.Text : ""));
                sqlParameter.Add("nchkSubject", (chk4MSubjectSrh.IsChecked == true ? 1 : 0));
                sqlParameter.Add("Subject", (chk4MSubjectSrh.IsChecked == true ? txt4MSubjectSrh.Text : ""));
                sqlParameter.Add("nchkCHANGEGBN", (chkChangeGBN.IsChecked == true ? 1 : 0));
                sqlParameter.Add("sCHANGEGBN", (chkChangeGBN.IsChecked == true ? cboChangeGBNSrh.SelectedValue.ToString() : ""));

                DataSet ds = DataStore.Instance.ProcedureToDataSet("xp_Qul_s4MChange", sqlParameter, false);

                if (ds != null && ds.Tables.Count > 0)
                {
                    DataTable dt = ds.Tables[0];
                    if (dt.Rows.Count == 0)
                    {
                        MessageBox.Show("조회된 데이터가 없습니다.");
                        this.DataContext = null;
                    }
                    else
                    {
                        int i = 0;
                        DataRowCollection drc = dt.Rows;
                        foreach (DataRow dr in drc)
                        {
                            i++;
                            var Qul_4MChange = new Win_Qul_4MChange_NU_CodeView()
                            {
                                FourMID = dr["4MID"].ToString(),
                                FourMSubject = dr["4MSubject"].ToString(),
                                InOutGbn = dr["InOutGbn"].ToString(),
                                CustomID = dr["CustomID"].ToString(),
                                KCustom = dr["KCustom"].ToString(),

                                ArticleID = dr["ArticleID"].ToString(),
                                Article = dr["Article"].ToString(),
                                BuyerArticleNo = dr["BuyerArticleNo"].ToString(),
                                BuyerModelID = dr["BuyerModelID"].ToString(),
                                Model = dr["Model"].ToString(),
                                CHANGEGBN = dr["CHANGEGBN"].ToString(),

                                CHANGEGB_Name = dr["CHANGEGB_Name"].ToString(),
                                ChangeGBNDetail = dr["ChangeGBNDetail"].ToString(),
                                ChangeReason = dr["ChangeReason"].ToString(),
                                DrawNo = dr["DrawNo"].ToString(),
                                ECONO = dr["ECONO"].ToString(),

                                CustomApplicationDate = dr["CustomApplicationDate"].ToString(),
                                CustomApprovementDate = dr["CustomApprovementDate"].ToString(),
                                CustomExpectedDate = dr["CustomExpectedDate"].ToString(),
                                ECODate = dr["ECODate"].ToString(),
                                InHouseApplicationDate = dr["InHouseApplicationDate"].ToString(),

                                InHouseExpectedDate = dr["InHouseExpectedDate"].ToString(),
                                ISIRDate = dr["ISIRDate"].ToString(),
                                Sabun = dr["Sabun"].ToString(),
                                ValidationCheck = dr["ValidationCheck"].ToString(),
                                Comments = dr["Comments"].ToString(),

                                Validation1 = dr["Validation1"].ToString(),
                                Validation2 = dr["Validation2"].ToString(),
                                Validation3 = dr["Validation3"].ToString(),
                                ValidationCheck1 = dr["ValidationCheck1"].ToString(),
                                ValidationCheck2 = dr["ValidationCheck2"].ToString(),
                                ValidationCheck3 = dr["ValidationCheck3"].ToString(),
                            };

                            if (Qul_4MChange.CustomApplicationDate.Length == 8)
                            {
                                Qul_4MChange.CustomApplicationDate_CV =
                                    lib.StrDateTimeBar(Qul_4MChange.CustomApplicationDate);
                            }

                            if (Qul_4MChange.CustomApprovementDate.Length == 8)
                            {
                                Qul_4MChange.CustomApprovementDate_CV =
                                    lib.StrDateTimeBar(Qul_4MChange.CustomApprovementDate);
                            }

                            if (Qul_4MChange.CustomExpectedDate.Length == 8)
                            {
                                Qul_4MChange.CustomExpectedDate_CV =
                                    lib.StrDateTimeBar(Qul_4MChange.CustomExpectedDate);
                            }

                            if (Qul_4MChange.ECODate.Length == 8)
                            {
                                Qul_4MChange.ECODate_CV =
                                    lib.StrDateTimeBar(Qul_4MChange.ECODate);
                            }

                            if (Qul_4MChange.InHouseApplicationDate.Length == 8)
                            {
                                Qul_4MChange.InHouseApplicationDate_CV =
                                    lib.StrDateTimeBar(Qul_4MChange.InHouseApplicationDate);
                            }

                            if (Qul_4MChange.InHouseExpectedDate.Length == 8)
                            {
                                Qul_4MChange.InHouseExpectedDate_CV =
                                    lib.StrDateTimeBar(Qul_4MChange.InHouseExpectedDate);
                            }

                            if (Qul_4MChange.ISIRDate.Length == 8)
                            {
                                Qul_4MChange.ISIRDate_CV =
                                    lib.StrDateTimeBar(Qul_4MChange.ISIRDate);
                            }

                            if (Qul_4MChange.InOutGbn.Equals("1"))
                            {
                                Qul_4MChange.InCustom = Qul_4MChange.KCustom;
                                Qul_4MChange.OutCustom = "-";
                            }
                            else if (Qul_4MChange.InOutGbn.Equals("2"))
                            {
                                Qul_4MChange.InCustom = "-";
                                Qul_4MChange.OutCustom = Qul_4MChange.KCustom;
                            }

                            //Qul_4MChange.InHouseExpectedDate = DatePickerFormat(Qul_4MChange.InHouseApplicationDate);

                            //Qul_4MChange.ovcModel = ComboBoxUtil.Instance.GetModelID_SetComboBox(Qul_4MChange.ArticleID);
                            tbkIndexCount.Text = "▶검색결과 : " + i + " 건";
                            dgdMain.Items.Add(Qul_4MChange);
                        }
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
        }

        //Main DataGrid cell selection Event
        private void dgdMain_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            win4MC = dgdMain.SelectedItem as Win_Qul_4MChange_NU_CodeView;

            if (win4MC != null)
            {
                this.DataContext = win4MC;
                FillGridValid(win4MC);
                FillGridSub(win4MC.FourMID);
            }
            else
            {
                this.DataContext = null;
                if (dgdValidation.Items.Count > 0)
                {
                    dgdValidation.Items.Clear();
                }
                if (dgdSub_Article.Items.Count > 0)
                {
                    dgdSub_Article.Items.Clear();
                }
            }
        }

        private void FillGridValid(Win_Qul_4MChange_NU_CodeView FourMNU)
        {
            if (dgdValidation.Items.Count > 0)
            {
                dgdValidation.Items.Clear();
            }

            if (FourMNU != null)
            {
                var Valid = new Win_Qul_4MChangeNUValid_CodeView()
                {
                    GubunName = "초기유동",
                    Valid1 = FourMNU.Validation1,
                    Valid2 = FourMNU.Validation2,
                    Valid3 = FourMNU.Validation3
                };

                dgdValidation.Items.Add(Valid);

                var ValidC = new Win_Qul_4MChangeNUValid_CodeView()
                {
                    GubunName = "확인결과",
                    Valid1 = FourMNU.ValidationCheck1,
                    Valid2 = FourMNU.ValidationCheck2,
                    Valid3 = FourMNU.ValidationCheck3
                };

                dgdValidation.Items.Add(ValidC);
            }
        }

        //SubArticle 실질 조회
        private void FillGridSub(string FourMid)
        {
            if (dgdSub_Article.Items.Count > 0)
            {
                dgdSub_Article.Items.Clear();
            }

            try
            {
                Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                sqlParameter.Add("s4MID", FourMid);

                DataSet ds = DataStore.Instance.ProcedureToDataSet("xp_Qul_s4MChangeArticle", sqlParameter, false);

                if (ds != null && ds.Tables.Count > 0)
                {
                    DataTable dt = ds.Tables[0];

                    if (dt.Rows.Count > 0)
                    {
                        DataRowCollection drc = dt.Rows;

                        foreach (DataRow dr in drc)
                        {
                            var dgd_Sub = new Win_Qul_4MChangeNUSub_CodeView()
                            {
                                FourMID = dr["4MID"].ToString(),
                                Article = dr["Article"].ToString(),
                                ArticleID = dr["ArticleID"].ToString(),
                                BuyerArticleNo = dr["BuyerArticleNo"].ToString(),
                                BuyerModelID = dr["BuyerModelID"].ToString(),
                                Model = dr["Model"].ToString(),

                                flagArticle = false,
                                flagArticleNO = false,
                                flagModel = false
                            };

                            strArticleID = dgd_Sub.ArticleID;

                            //dgd_Sub.ovcBuyerArticleNo = ComboBoxUtil.Instance.GetBuyerArticleNo_SetComboBox(dgd_Sub.ArticleID);
                            dgdSub_Article.Items.Add(dgd_Sub);
                        }
                        drc.Clear();
                    }
                    dt.Clear();
                }
                ds.Clear();
            }
            catch (Exception ex)
            {
                MessageBox.Show("오류 발생,오류 내용 : " + ex.ToString());
            }
            finally
            {
                DataStore.Instance.CloseConnection();
            }
        }

        //
        private bool SaveData(string strID)
        {
            bool flag = true;
            List<Procedure> Prolist = new List<Procedure>();
            List<Dictionary<string, object>> ListParameter = new List<Dictionary<string, object>>();

            try
            {
                if (CheckData())
                {
                    Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                    sqlParameter.Add("4MID", strID);
                    sqlParameter.Add("4MSubject", txt4MSubject.Text);
                    sqlParameter.Add("InOutGbn", cboInOutGbn.SelectedValue != null ?
                        cboInOutGbn.SelectedValue.ToString() :
                        (txtCustom.Tag != null && txtCustom.Tag.ToString().Equals("0001")) ? "1" : "2");
                    sqlParameter.Add("CustomID", txtCustom.Tag.ToString());
                    sqlParameter.Add("ECONO", txtECONO.Text);

                    sqlParameter.Add("Sabun", "");
                    sqlParameter.Add("ArticleID", "");  //2020.05.22, 장가빈, 품명은 sub그리드에 모두 넣기로 함.
                    //sqlParameter.Add("DrawNo", txtDrawNo.Text);
                    sqlParameter.Add("DrawNo", "");
                    sqlParameter.Add("BuyerModelID", txtBuyerModel.Tag != null ? txtBuyerModel.Tag.ToString() : "");
                    sqlParameter.Add("ChangeGBN", cboChangeGBN.SelectedValue.ToString());

                    sqlParameter.Add("ChangeGBNDetail", txtChangeGBNDetail.Text);
                    sqlParameter.Add("ChangeReason", txtChangeReason.Text);
                    sqlParameter.Add("ECODate", dtpECODate.SelectedDate != null ? dtpECODate.SelectedDate.Value.ToString("yyyyMMdd") : "");
                    sqlParameter.Add("ISIRDate", dtpISIRDate.SelectedDate != null ? dtpISIRDate.SelectedDate.Value.ToString("yyyyMMdd") : "");
                    sqlParameter.Add("CustomApprovementDate", dtpCustomApprovementDate.SelectedDate != null ?
                        dtpCustomApprovementDate.SelectedDate.Value.ToString("yyyyMMdd") : "");

                    sqlParameter.Add("InHouseExpectedDate", dtpCustomApprovementDate.SelectedDate != null ?
                        dtpInHouseExpectedDate.SelectedDate.Value.ToString("yyyyMMdd") : "");
                    sqlParameter.Add("InHouseApplicationDate", dtpInHouseApplicationDate.SelectedDate != null ?
                        dtpInHouseApplicationDate.SelectedDate.Value.ToString("yyyyMMdd") : "");
                    sqlParameter.Add("CustomExpectedDate", dtpCustomExpectedDate.SelectedDate != null ?
                        dtpCustomExpectedDate.SelectedDate.Value.ToString("yyyyMMdd") : "");
                    sqlParameter.Add("CustomApplicationDate", dtpCustomApplicationDate.SelectedDate != null ?
                        dtpCustomApplicationDate.SelectedDate.Value.ToString("yyyyMMdd") : "");
                    sqlParameter.Add("ValidationCheck", txtValidationCheck.Text);
                    sqlParameter.Add("Comments", txtComments.Text);

                    if (dgdValidation.Items.Count > 0)
                    {
                        var Valid = dgdValidation.Items[0] as Win_Qul_4MChangeNUValid_CodeView;
                        sqlParameter.Add("Validation1", lib.CheckNull(Valid.Valid1));
                        sqlParameter.Add("Validation2", lib.CheckNull(Valid.Valid2));
                        sqlParameter.Add("Validation3", lib.CheckNull(Valid.Valid3));
                    }

                    if (dgdValidation.Items.Count > 1)
                    {
                        var ValidC = dgdValidation.Items[1] as Win_Qul_4MChangeNUValid_CodeView;
                        sqlParameter.Add("ValidationCheck1", lib.CheckNull(ValidC.Valid1));
                        sqlParameter.Add("ValidationCheck2", lib.CheckNull(ValidC.Valid2));
                        sqlParameter.Add("ValidationCheck3", lib.CheckNull(ValidC.Valid3));
                    }

                    if (strFlag.Equals("I"))
                    {
                        sqlParameter.Add("CreateUserID", MainWindow.CurrentUser);

                        Procedure pro1 = new Procedure();
                        pro1.Name = "xp_Qul_i4MChange";
                        pro1.OutputUseYN = "Y";
                        pro1.OutputName = "4MID";
                        pro1.OutputLength = "10";

                        Prolist.Add(pro1);
                        ListParameter.Add(sqlParameter);

                        for (int i = 0; i < dgdSub_Article.Items.Count; i++)
                        {
                            var dgdSub = dgdSub_Article.Items[i] as Win_Qul_4MChangeNUSub_CodeView;

                            sqlParameter = new Dictionary<string, object>();
                            sqlParameter.Clear();
                            sqlParameter.Add("4MID", strID);
                            sqlParameter.Add("ArticleID", dgdSub.ArticleID);
                            sqlParameter.Add("BuyerArticleNo", dgdSub.BuyerArticleNo);
                            sqlParameter.Add("BuyerModelID", dgdSub.BuyerModelID);
                            sqlParameter.Add("CreateUserID", MainWindow.CurrentUser);

                            Procedure pro2 = new Procedure();
                            pro2.Name = "xp_Qul_i4MChangeArticle";
                            pro2.OutputUseYN = "N";
                            pro2.OutputName = "4MID";
                            pro2.OutputLength = "10";

                            Prolist.Add(pro2);
                            ListParameter.Add(sqlParameter);
                        }

                        List<KeyValue> list_Result = new List<KeyValue>();
                        list_Result = DataStore.Instance.ExecuteAllProcedureOutputGetCS(Prolist, ListParameter);
                        string sGet4MID = string.Empty;

                        if (list_Result[0].key.ToLower() == "success")
                        {
                            list_Result.RemoveAt(0);
                            for (int i = 0; i < list_Result.Count; i++)
                            {
                                KeyValue kv = list_Result[i];
                                if (kv.key == "4MID")
                                {
                                    sGet4MID = kv.value;
                                    dicCompare.Add("4MID", sGet4MID);
                                }
                            }
                        }
                        else
                        {
                            MessageBox.Show("[저장실패]\r\n" + list_Result[0].value.ToString());
                            flag = false;
                            //return false;
                        }
                    }
                    else if (strFlag.Equals("U"))
                    {
                        sqlParameter.Add("UserID", MainWindow.CurrentUser);

                        Procedure pro1 = new Procedure();
                        pro1.Name = "xp_Qul_u4MChange";
                        pro1.OutputUseYN = "N";
                        pro1.OutputName = "4MID";
                        pro1.OutputLength = "10";

                        Prolist.Add(pro1);
                        ListParameter.Add(sqlParameter);

                        for (int i = 0; i < dgdSub_Article.Items.Count; i++)
                        {
                            var dgdSub = dgdSub_Article.Items[i] as Win_Qul_4MChangeNUSub_CodeView;

                            sqlParameter = new Dictionary<string, object>();
                            sqlParameter.Clear();
                            sqlParameter.Add("4MID", strID);
                            sqlParameter.Add("ArticleID", dgdSub.ArticleID);
                            sqlParameter.Add("BuyerArticleNo", dgdSub.BuyerArticleNo);
                            sqlParameter.Add("BuyerModelID", dgdSub.BuyerModelID);
                            sqlParameter.Add("CreateUserID", MainWindow.CurrentUser);

                            Procedure pro2 = new Procedure();
                            pro2.Name = "xp_Qul_i4MChangeArticle";
                            pro2.OutputUseYN = "N";
                            pro2.OutputName = "4MID";
                            pro2.OutputLength = "10";

                            Prolist.Add(pro2);
                            ListParameter.Add(sqlParameter);
                        }

                        string[] Confirm = new string[2];
                        Confirm = DataStore.Instance.ExecuteAllProcedureOutputNew(Prolist, ListParameter);
                        if (Confirm[0] == "success")
                        {
                            flag = true;
                        }
                        else
                        {
                            MessageBox.Show("[저장실패]\r\n" + Confirm[1].ToString());
                            //return false;
                        }
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

            return flag;
        }

        //DB Main,Sub Data Delete
        private bool DeleteData(string str4MID)
        {
            bool flag = true;

            if (CheckData())
            {
                try
                {
                    Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                    sqlParameter.Add("4MID", str4MID);

                    string[] result = DataStore.Instance.ExecuteProcedure("xp_Qul_d4MChange", sqlParameter, true);

                    if (!result[0].Equals("success"))
                    {
                        MessageBox.Show("삭제 실패");
                        flag = false;
                    }
                    else
                    {
                        //MessageBox.Show("성공 *^^*");
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
            }
            else
            {
                flag = false;
            }

            return flag;
        }

        //데이터 체크
        private bool CheckData()
        {
            bool flag = true;

            if (cboChangeGBN.SelectedValue == null)
            {
                MessageBox.Show("변경구분이 선택되지 않았습니다.");
                flag = false;
                return flag;
            }

            if (txtCustom.Tag == null)
            {
                MessageBox.Show("업체명이 선택되지 않았습니다.");
                flag = false;
                return flag;
            }

            //if (txtArticle.Tag == null)
            //{
            //    MessageBox.Show("품명이 선택되지 않았습니다.");
            //    flag = false;
            //    return flag;
            //}

            return flag;
        }

        #endregion

        #region 플러스 파인더 및 enter focus move

        private void cboInOutGbn_DropDownClosed(object sender, EventArgs e)
        {
            try
            {
                if (cboInOutGbn.SelectedValue.Equals("1"))
                {
                    txtCustom.Text = "주식회사 지엘에스";
                    txtCustom.Tag = "0001";
                }
                else if (cboInOutGbn.SelectedValue.Equals("2"))
                {
                    txtCustom.Clear();
                }
                txtCustom.Focus();
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

        //업체명
        private void txtCustom_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                if (txtArticle.Tag != null && txtArticle.Text.Length > 0)
                {
                    pf.ReturnCode(txtCustom, (int)Defind_CodeFind.DCF_CUSTOM, txtArticle.Tag.ToString());
                }
                else
                {
                    pf.ReturnCode(txtCustom, (int)Defind_CodeFind.DCF_CUSTOM, "");
                }
                //if (cboInOutGbn.SelectedValue != null && cboInOutGbn.SelectedValue.Equals("1"))
                //{

                //}
                //else
                //{
                //    pf.ReturnCode(txtCustom, 68, "");
                //}

                if (txtCustom.Tag != null && txtCustom.Tag.ToString().Equals("0001"))
                {
                    cboInOutGbn.SelectedValue = "1";
                }
                else
                {
                    cboInOutGbn.SelectedValue = "2";
                }

                txtECONO.Focus();
            }
        }

        //업체명
        private void btnPfCustom_Click(object sender, RoutedEventArgs e)
        {
            if (txtArticle.Tag != null && txtArticle.Text.Length > 0)
            {
                pf.ReturnCode(txtCustom, (int)Defind_CodeFind.DCF_CUSTOM, txtArticle.Tag.ToString());
            }
            else
            {
                pf.ReturnCode(txtCustom, (int)Defind_CodeFind.DCF_CUSTOM, "");
            }

            //if (cboInOutGbn.SelectedValue != null && cboInOutGbn.SelectedValue.Equals("1"))
            //{

            //}
            //else
            //{
            //    pf.ReturnCode(txtCustom, 68, "");
            //}
            if (txtCustom.Tag != null && txtCustom.Tag.ToString().Equals("0001"))
            {
                cboInOutGbn.SelectedValue = "1";
            }
            else
            {
                cboInOutGbn.SelectedValue = "2";
            }

            txtECONO.Focus();
        }

        //ECONO
        private void TxtECONO_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                txtArticle.Focus();
            }
        }





        //품명 -> 품번으로 변경 2020.03.19, 장가빈 
        private void TxtArticle_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                if (txtCustom.Tag != null && txtCustom.Text.Length > 0 && txtCustom.Text.Length > 0)
                {
                    //pf.ReturnCode(txtArticle, 64, txtCustom.Tag.ToString());
                    if (txtCustom.Tag.ToString().Equals("0001"))
                    {
                        pf.ReturnCode(txtArticle, 80, "0001");
                    }
                    else
                    {
                        pf.ReturnCode(txtArticle, 80, txtCustom.Tag.ToString());
                    }
                }
                else
                {
                    pf.ReturnCode(txtArticle, 81, txtArticle.Text);
                }

                if (txtArticle.Tag != null)
                {
                    //cboBuyerModel.ItemsSource = ComboBoxUtil.Instance.GetModelID_SetComboBox(txtArticle.Tag.ToString());
                }
                //txtDrawNo.Focus();
                cboBuyerModel.Focus();
                SetBuyerArticelNo(txtArticle.Tag.ToString());
            }
        }

        //품명
        private void BtnPfArticle_Click(object sender, RoutedEventArgs e)
        {
            if (txtCustom.Tag != null && txtCustom.Text.Length > 0 && txtCustom.Text.Length > 0)
            {
                //pf.ReturnCode(txtArticle, 64, txtCustom.Tag.ToString());
                if (txtCustom.Tag.ToString().Equals("0001"))
                {
                    pf.ReturnCode(txtArticle, 80, "0001");
                }
                else
                {
                    pf.ReturnCode(txtArticle, 80, txtCustom.Tag.ToString());
                }
            }
            else
            {
                pf.ReturnCode(txtArticle, 81, txtArticle.Text);
            }

            if (txtArticle.Tag != null)
            {
                //cboBuyerModel.ItemsSource = ComboBoxUtil.Instance.GetModelID_SetComboBox(txtArticle.Tag.ToString());
            }
            //txtDrawNo.Focus();
            cboBuyerModel.Focus();
            SetBuyerArticelNo(txtArticle.Tag.ToString());
        }

        //품명 넘버 가져오기
        private void SetBuyerArticelNo(string strArticleID) //품명을 뿌려야하니까 수정 2020.03.19, 장가빈
        {
            DataTable dt = Procedure.Instance.GetArticleData(strArticleID);

            if (dt.Rows.Count > 0)
            {
                if (dt.Rows[0]["Article"] != null &&
                    !dt.Rows[0]["Article"].ToString().Trim().Equals(string.Empty))
                {
                    txtBuyerArticleNo.Text = dt.Rows[0]["Article"].ToString();

                    //tempArticle = dt.Rows[0]["Article"].ToString();
                }
            }
        }

        //품명 넘버 가져오기
        private void SetBuyerArticelNoSub(string strArticleID) //품명을 뿌려야하니까 수정 2020.03.19, 장가빈
        {
            DataTable dt = Procedure.Instance.GetArticleData(strArticleID);

            if (dt.Rows.Count > 0)
            {
                if (dt.Rows[0]["Article"] != null &&
                    !dt.Rows[0]["Article"].ToString().Trim().Equals(string.Empty))
                {
                    //txtBuyerArticleNo.Text = dt.Rows[0]["Article"].ToString();

                    tempArticle = dt.Rows[0]["Article"].ToString();
                }
            }
        }

        //도면 품번
        private void TxtDrawNo_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                pf.ReturnCode(txtDrawNo, (int)Defind_CodeFind.DCF_Article, "");
                //pf.ReturnCode(txtDrawNo, 74, "");
                //txtBuyerModel.Focus();
            }
        }

        //도면 품번
        private void BtnPfDrawNo_Click(object sender, RoutedEventArgs e)
        {
            //pf.ReturnCode(txtDrawNo, 74, "");
            //txtBuyerModel.Focus();
        }



        //차종
        private void TxtBuyerModel_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                pf.ReturnCode(txtBuyerModel, 28, "");
                cboChangeGBN.Focus();
                cboChangeGBN.IsDropDownOpen = true;
            }
        }

        //차종
        private void BtnPfBuyerModel_Click(object sender, RoutedEventArgs e)
        {
            pf.ReturnCode(txtBuyerModel, 28, "");
            cboChangeGBN.Focus();
            cboChangeGBN.IsDropDownOpen = true;
        }

        //변경구분
        private void CboChangeGBN_DropDownClosed(object sender, EventArgs e)
        {
            txtChangeGBNDetail.Focus();
        }

        //변경구분상세
        private void TxtChangeGBNDetail_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                txtChangeReason.Focus();
            }
        }

        //변경사유
        private void TxtChangeReason_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                dtpECODate.Focus();
            }
        }

        // ECO/4M 등록일자
        private void DtpECODate_PreviewKeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                dtpECODate.IsDropDownOpen = true;
            }
        }

        // ECO/4M 등록일자
        private void DtpECODate_CalendarClosed(object sender, RoutedEventArgs e)
        {
            dtpISIRDate.Focus();
        }

        // ISIR 등록일자
        private void DtpISIRDate_CalendarClosed(object sender, RoutedEventArgs e)
        {
            dtpCustomApprovementDate.Focus();
        }

        //고객승인일자
        private void DtpCustomApprovementDate_CalendarClosed(object sender, RoutedEventArgs e)
        {
            dtpInHouseExpectedDate.Focus();
        }

        //사내적용예정일
        private void DtpInHouseExpectedDate_CalendarClosed(object sender, RoutedEventArgs e)
        {
            dtpInHouseApplicationDate.Focus();
        }

        //사내적용일
        private void DtpInHouseApplicationDate_CalendarClosed(object sender, RoutedEventArgs e)
        {
            dtpCustomExpectedDate.Focus();
        }

        //고객적용예정일
        private void DtpCustomExpectedDate_CalendarClosed(object sender, RoutedEventArgs e)
        {
            dtpCustomApplicationDate.Focus();
        }

        //고객적용일
        private void DtpCustomApplicationDate_CalendarClosed(object sender, RoutedEventArgs e)
        {
            //txtValidationCheck.Focus();
            txtComments.Focus();
        }

        //유효성점검
        private void TxtValidationCheck_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                txtComments.Focus();
            }
        }

        //비고
        private void TxtComments_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                btnSubPlus.Focus();
            }
        }

        #endregion

        #region 서브 그리드 입력

        //
        private void btnSubPlus_Click(object sender, RoutedEventArgs e)
        {
            SubPlus();

            //dgdSub_Article.Focus();
            ////dgdSub_Article.SelectedIndex = dgdSub_Article.Items.Count - 1;
            //dgdSub_Article.CurrentCell = 
            //    new DataGridCellInfo(dgdSub_Article.Items[dgdSub_Article.Items.Count - 1], dgdSub_Article.Columns[0]);
        }

        private void SubPlus()
        {
            var win4M_Sub = new Win_Qul_4MChangeNUSub_CodeView()
            {
                Article = string.Empty,
                ArticleID = string.Empty,
                BuyerArticleNo = string.Empty,
                Model = string.Empty,
                BuyerModelID = string.Empty,
                ovcBuyerArticleNo = ovcBAN,
                ovcBuyerModel = ovcBMI
            };

            dgdSub_Article.Items.Add(win4M_Sub);
        }

        //
        private void btnSubDel_Click(object sender, RoutedEventArgs e)
        {
            SubRemove();
        }

        //서브 행 삭제
        private void SubRemove()
        {
            if (dgdSub_Article.Items.Count > 0)
            {
                if (dgdSub_Article.CurrentItem != null)
                {
                    dgdSub_Article.Items.Remove((dgdSub_Article.CurrentItem as Win_Qul_4MChangeNUSub_CodeView));
                }
                else
                {
                    dgdSub_Article.Items.Remove((dgdSub_Article.Items[dgdSub_Article.Items.Count - 1]) as Win_Qul_4MChangeNUSub_CodeView);
                }
                dgdSub_Article.Refresh();
            }
        }

        //
        private void DataGridValidCell_KeyDown(object sender, KeyEventArgs e)
        {
            if (lblMsg.Visibility == Visibility.Visible)
            {
                var QulValid = dgdValidation.CurrentItem as Win_Qul_4MChangeNUValid_CodeView;
                int rowCount = dgdValidation.Items.IndexOf(dgdValidation.CurrentItem);
                int colCount = dgdValidation.Columns.IndexOf(dgdValidation.CurrentCell.Column);

                if (e.Key == Key.Enter)
                {
                    e.Handled = true;
                    (sender as DataGridCell).IsEditing = false;

                    if (dgdValidation.Columns.Count - 1 == colCount && rowCount == 0)
                    {
                        dgdValidation.SelectedIndex = rowCount + 1;
                        dgdValidation.CurrentCell = new DataGridCellInfo(dgdValidation.Items[rowCount + 1], dgdValidation.Columns[1]);
                    }
                    else if (dgdValidation.Columns.Count - 1 == colCount && rowCount == 1)
                    {
                        btnSubPlus.Focus();
                    }
                    else if (dgdValidation.Columns.Count - 1 > colCount)
                    {
                        dgdValidation.CurrentCell = new DataGridCellInfo(dgdValidation.Items[rowCount], dgdValidation.Columns[colCount + 1]);
                    }
                    else
                    {

                    }
                }
            }
        }

        //
        private void DataGridSubCell_KeyDown(object sender, KeyEventArgs e)
        {
            if (lblMsg.Visibility == Visibility.Visible)
            {
                var Qul4M = dgdSub_Article.CurrentItem as Win_Qul_4MChangeNUSub_CodeView;
                int rowCount = dgdSub_Article.Items.IndexOf(dgdSub_Article.CurrentItem);
                int colCount = dgdSub_Article.Columns.IndexOf(dgdSub_Article.CurrentCell.Column);
                int lastColcount = dgdSub_Article.Columns.Count - 1;

                //MessageBox.Show(e.Key.ToString());

                if (e.Key == Key.Enter)
                {
                    e.Handled = true;
                    (sender as DataGridCell).IsEditing = false;

                    if (lastColcount == colCount && dgdSub_Article.Items.Count - 1 > rowCount)
                    {
                        dgdSub_Article.SelectedIndex = rowCount + 1;
                        dgdSub_Article.CurrentCell = new DataGridCellInfo(dgdSub_Article.Items[rowCount + 1], dgdSub_Article.Columns[1]);
                    }
                    else if (lastColcount > colCount && dgdSub_Article.Items.Count - 1 > rowCount)
                    {
                        dgdSub_Article.CurrentCell = new DataGridCellInfo(dgdSub_Article.Items[rowCount], dgdSub_Article.Columns[colCount + 1]);
                    }
                    else if (lastColcount == colCount && dgdSub_Article.Items.Count - 1 == rowCount)
                    {
                        if (MessageBox.Show("추가하시겠습니까?", "추가 확인", MessageBoxButton.YesNo) == MessageBoxResult.Yes)
                        {
                            SubPlus();
                            dgdSub_Article.SelectedIndex = rowCount + 1;
                            dgdSub_Article.CurrentCell = new DataGridCellInfo(dgdSub_Article.Items[rowCount + 1], dgdSub_Article.Columns[0]);
                        }
                        else
                        {
                            btnSave.Focus();
                        }
                    }
                    else if (lastColcount > colCount && dgdSub_Article.Items.Count - 1 == rowCount)
                    {
                        dgdSub_Article.CurrentCell = new DataGridCellInfo(dgdSub_Article.Items[rowCount], dgdSub_Article.Columns[colCount + 1]);
                    }
                    else
                    {
                        MessageBox.Show("있으면 찾아보자...");
                    }
                }
                else if (e.Key == Key.Delete)
                {
                    e.Handled = true;

                    SubRemove();

                    dgdSub_Article.Refresh();
                    if (dgdSub_Article.Items.Count > 0)
                    {
                        if (dgdSub_Article.Items.Count - 1 > rowCount)
                        {
                            dgdSub_Article.SelectedIndex = rowCount;
                        }
                        else
                        {
                            dgdSub_Article.SelectedIndex = 0;
                        }
                    }
                }
            }
        }

        //
        private void TextBoxFocusInDataGrid(object sender, KeyEventArgs e)
        {
            lib.DataGridINControlFocus(sender, e);
        }

        //
        private void TextBoxFocusInDataGrid_MouseUp(object sender, MouseButtonEventArgs e)
        {
            lib.DataGridINBothByMouseUP(sender, e);
        }

        //
        private void DataGridCell_GotFocus(object sender, RoutedEventArgs e)
        {
            if (lblMsg.Visibility == Visibility.Visible)
            {
                DataGridCell cell = sender as DataGridCell;
                if (cell != null)
                {
                    cell.IsEditing = true;
                }
            }
        }

        private void dgdtxtArticle_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                var Qul4M = dgdSub_Article.CurrentItem as Win_Qul_4MChangeNUSub_CodeView;

                TextBox textBox = sender as TextBox;
                pf.ReturnCode(textBox, 81, textBox.Text);

                SetBuyerArticelNoSub(textBox.Tag.ToString());

                Qul4M.Article = tempArticle;

                //if (Qul4M !=null && textBox.Tag != null)
                //{
                //    //Qul4M.ovcBuyerArticleNo = ComboBoxUtil.Instance.GetBuyerArticleNo_SetComboBox(Qul4M.ArticleID);
                //    Qul4M.ovcBuyerModel = ComboBoxUtil.Instance.GetModelID_SetComboBox(Qul4M.ArticleID);
                //}
            }
        }

        private void dgdtxtArticle_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            var Qul4M = dgdSub_Article.CurrentItem as Win_Qul_4MChangeNUSub_CodeView;

            TextBox textBox = sender as TextBox;
            pf.ReturnCode(textBox, 81, textBox.Text);

            //if (Qul4M != null && textBox.Tag != null)
            //{
            //    //Qul4M.ovcBuyerArticleNo = ComboBoxUtil.Instance.GetBuyerArticleNo_SetComboBox(Qul4M.ArticleID);
            //    Qul4M.ovcBuyerModel = ComboBoxUtil.Instance.GetModelID_SetComboBox(Qul4M.ArticleID);
            //}
        }

        #endregion

        private void LvwMain_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {

        }

        private void dgdtpetxtBuyerModel_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                var Qul4M = dgdSub_Article.CurrentItem as Win_Qul_4MChangeNUSub_CodeView;

                TextBox textBox = sender as TextBox;
                pf.ReturnCode(textBox, (int)Defind_CodeFind.DCF_BUYERMODEL, "");
            }
        }

        private void dgdtpetxtBuyerModel_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            var Qul4M = dgdSub_Article.CurrentItem as Win_Qul_4MChangeNUSub_CodeView;

            TextBox textBox = sender as TextBox;
            pf.ReturnCode(textBox, (int)Defind_CodeFind.DCF_BUYERMODEL, "");
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

        #region 기타 메서드

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
                str = str.Replace(",", "");

                if (Double.TryParse(str, out chkDouble) == true)
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

        // 월만 가져오기 > 앞에 0 없애기
        private string getDateMonth(string str)
        {
            string month = "";

            if (!str.Trim().Equals(""))
            {
                str = str.Trim().Replace("-", "").Replace(".", "");

                if (str.Length == 8)
                {
                    month = str.Substring(4, 2);

                    if (month.Substring(0, 1).Equals("0"))
                    {
                        month = month.Substring(0, 1);
                    }
                }
            }

            return month;
        }

        // 일만 가져오기 > 앞에 0 없애기
        private string getDateDay(string str)
        {
            string day = "";

            if (!str.Trim().Equals(""))
            {
                str = str.Trim().Replace("-", "").Replace(".", "");

                if (str.Length == 8)
                {
                    day = str.Substring(6, 2);

                    if (day.Substring(0, 1).Equals("0"))
                    {
                        day = day.Substring(1, 1);
                    }
                }
            }

            return day;
        }

        #endregion // 기타 메서드


    }



    class Win_Qul_4MChange_NU_CodeView : BaseView
    {
        public override string ToString()
        {
            return (this.ReportAllProperties());
        }

        public string FourMID { get; set; }
        public string FourMSubject { get; set; }
        public string InOutGbn { get; set; }
        public string CustomID { get; set; }
        public string KCustom { get; set; }

        public string ArticleID { get; set; }
        public string Article { get; set; }
        public string BuyerModelID { get; set; }
        public string Model { get; set; }
        public string CHANGEGBN { get; set; }

        public string CHANGEGB_Name { get; set; }
        public string ChangeGBNDetail { get; set; }
        public string ChangeReason { get; set; }
        public string DrawNo { get; set; }
        public string ECONO { get; set; }

        public string CustomApplicationDate { get; set; }
        public string CustomApprovementDate { get; set; }
        public string CustomExpectedDate { get; set; }
        public string ECODate { get; set; }
        public string InHouseApplicationDate { get; set; }

        public string InHouseExpectedDate { get; set; }
        public string ISIRDate { get; set; }
        public string Sabun { get; set; }
        public string ValidationCheck { get; set; }
        public string Comments { get; set; }

        public string Validation1 { get; set; }
        public string ValidationCheck1 { get; set; }
        public string Validation2 { get; set; }
        public string ValidationCheck2 { get; set; }
        public string Validation3 { get; set; }
        public string ValidationCheck3 { get; set; }

        public string CustomApplicationDate_CV { get; set; }
        public string CustomApprovementDate_CV { get; set; }
        public string CustomExpectedDate_CV { get; set; }
        public string ECODate_CV { get; set; }
        public string InHouseApplicationDate_CV { get; set; }

        public string InHouseExpectedDate_CV { get; set; }
        public string ISIRDate_CV { get; set; }
        public ObservableCollection<CodeView> ovcModel { get; set; }
        public string InCustom { get; set; }
        public string OutCustom { get; set; }

        public string BuyerArticleNo { get; set; }

    }

    class Win_Qul_4MChangeNUValid_CodeView : BaseView
    {
        public override string ToString()
        {
            return (this.ReportAllProperties());
        }

        public string GubunName { get; set; }
        public string Valid1 { get; set; }
        public string Valid2 { get; set; }
        public string Valid3 { get; set; }
    }

    class Win_Qul_4MChangeNUSub_CodeView : BaseView
    {
        public override string ToString()
        {
            return (this.ReportAllProperties());
        }

        public string FourMID { get; set; }
        public string Article { get; set; }
        public string ArticleID { get; set; }
        public string BuyerArticleNo { get; set; }
        public string BuyerModelID { get; set; }
        public string Model { get; set; }

        public ObservableCollection<CodeView> ovcBuyerArticleNo { get; set; }
        public ObservableCollection<CodeView> ovcBuyerModel { get; set; }
        public bool flagArticle { get; set; }
        public bool flagArticleNO { get; set; }
        public bool flagModel { get; set; }
    }
}

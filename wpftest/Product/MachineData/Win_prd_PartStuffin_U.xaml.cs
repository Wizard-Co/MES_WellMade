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

namespace WizMes_WellMade
{
    /// <summary>
    /// Win_prd_PartStuffin_U.xaml에 대한 상호 작용 논리
    /// </summary>
    public partial class Win_prd_PartStuffin_U : UserControl
    {
        Win_prd_PartStuffin_U_CodeView WinPartStuff = new Win_prd_PartStuffin_U_CodeView();
        Lib lib = new Lib();
        string jobFlag = string.Empty;
        int numRowCount = 0;
        Dictionary<string, object> dicCompare = new Dictionary<string, object>();
        List<string> lstCompareValue = new List<string>();



        public Win_prd_PartStuffin_U()
        {
            InitializeComponent();
        }

        private void Usercontrol_Loaded(object sender, RoutedEventArgs e)
        {
            Lib.Instance.UiLoading(sender);
            dtpSDate.SelectedDate = DateTime.Today;
            dtpEDate.SelectedDate = DateTime.Today;
            SetComboBox();
            chkForUseSrh.IsChecked = true;
        }


        #region 콤보박스 

        //콤보박스 세팅
        private void SetComboBox()
        {
            ObservableCollection<CodeView> ovcUnit = ComboBoxUtil.Instance.Gf_DB_CM_GetComCodeDataset(null, "MTRUNIT", "Y", "", "");
            this.cboUnit.ItemsSource = ovcUnit;
            this.cboUnit.DisplayMemberPath = "code_name";
            this.cboUnit.SelectedValuePath = "code_id";

            ObservableCollection<CodeView> ovcPriceClss = ComboBoxUtil.Instance.Gf_DB_CM_GetComCodeDataset(null, "CRU", "Y", "", "");
            this.cboPriceClss.ItemsSource = ovcPriceClss;
            this.cboPriceClss.DisplayMemberPath = "code_name";
            this.cboPriceClss.SelectedValuePath = "code_id";

            //Y 부가세별도, N 부가세포함, 0 부가세=0
            List<string[]> lstVat = new List<string[]>();
            lstVat.Add(new string[] { "Y", "Y" });
            lstVat.Add(new string[] { "N", "N" });
            lstVat.Add(new string[] { "0", "0" });

            ObservableCollection<CodeView> ovcVatInd = ComboBoxUtil.Instance.Direct_SetComboBox(lstVat);
            this.cboVatInd.ItemsSource = ovcVatInd;
            this.cboVatInd.DisplayMemberPath = "code_name";
            this.cboVatInd.SelectedValuePath = "code_id";

            List<string[]> lstForUseSrh = new List<string[]>();
            lstForUseSrh.Add(new string[] { "0", "전체" });
            lstForUseSrh.Add(new string[] { "1", "공용" });
            lstForUseSrh.Add(new string[] { "2", "설비예비품" });
            lstForUseSrh.Add(new string[] { "3", "Tool" });

            ObservableCollection<CodeView> ovcForUseSrh = ComboBoxUtil.Instance.Direct_SetComboBox(lstForUseSrh);
            this.cboForUseSrh.ItemsSource = ovcForUseSrh;
            this.cboForUseSrh.DisplayMemberPath = "code_name";
            this.cboForUseSrh.SelectedValuePath = "code_id";
            this.cboForUseSrh.SelectedIndex = 0;

            List<string[]> lstForUse = new List<string[]>();
            lstForUse.Add(new string[] { "1", "공용" });
            lstForUse.Add(new string[] { "2", "설비예비품" });
            lstForUse.Add(new string[] { "3", "Tool" });

            ObservableCollection<CodeView> ovcForUse = ComboBoxUtil.Instance.Direct_SetComboBox(lstForUse);
            this.cboForUse.ItemsSource = ovcForUse;
            this.cboForUse.DisplayMemberPath = "code_name";
            this.cboForUse.SelectedValuePath = "code_id";

            ObservableCollection<CodeView> ovcComID = ComboBoxUtil.Instance.Get_CompanyID();
            this.cboCompanyS.ItemsSource = ovcComID;
            this.cboCompanyS.DisplayMemberPath = "code_name";
            this.cboCompanyS.SelectedValuePath = "code_id";



        }

        #endregion

        #region 왼쪽 상단 버튼

        //왼쪽 상단 TOOL 버튼 
        private void btnTool_Click(object sender, RoutedEventArgs e)
        {
            btnTool.IsChecked = true;
            btnEquip.IsChecked = false;

            dgdGroup.Columns[7].Header = "TOOL명";

            re_Search(numRowCount);

        }

        //왼쪽 상단 설비등록 버튼
        private void btnEquip_Click(object sender, RoutedEventArgs e)
        {
            btnTool.IsChecked = false;
            btnEquip.IsChecked = true;

            dgdGroup.Columns[7].Header = "설비예비품명";

            re_Search(numRowCount);
        }

        //입고일자
        private void lblInDay_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            if (chkInDay.IsChecked == true) { chkInDay.IsChecked = false; }
            else { chkInDay.IsChecked = true; }
        }

        //부품명 keyDown
        // 2021-05-27
        private void txtArticleSearch_KeyDown(object sender, KeyEventArgs e)
        {
            try
            {
                if (e.Key == Key.Enter)
                {
                    MainWindow.pf.ReturnCode(txtArticleSearch, (int)Defind_CodeFind.DCF_PART, "");
                }
            }
            catch (Exception ee)
            {
                MessageBox.Show("오류지점 - txtArticleSearch_KeyDown : " + ee.ToString());
            }
        }

        //부품명 플러스 파인더
        private void btnPfArticleSearch_Click(object sender, RoutedEventArgs e)
        {
            MainWindow.pf.ReturnCode(txtArticleSearch, (int)Defind_CodeFind.DCF_PART, "");
        }

        //왼쪽 상단 부품용도
        private void lblForUseSrh_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            if (chkForUseSrh.IsChecked == true) { chkForUseSrh.IsChecked = false; }
            else { chkForUseSrh.IsChecked = true; }
        }

        //상단 거래처
        private void btnPfCustom_Click(object sender, RoutedEventArgs e)
        {
            MainWindow.pf.ReturnCode(txtCustom, (int)Defind_CodeFind.DCF_CUSTOM, "");
            txtCustom.Focus();
        }

        //상단 거래처 keyDown
        // 2021-05-27
        private void txtCustomSearch_KeyDown(object sender, KeyEventArgs e)
        {

            try
            {
                if (e.Key == Key.Enter)
                {
                    MainWindow.pf.ReturnCode(txtCustomSearch, (int)Defind_CodeFind.DCF_CUSTOM, "");
                }
            }
            catch (Exception ee)
            {
                MessageBox.Show("오류지점 - txtCustomer_KeyDown : " + ee.ToString());
            }
        }

        //거래처 플러스 파인더
        private void btnPfCustomSearch_Click(object sender, RoutedEventArgs e)
        {
            MainWindow.pf.ReturnCode(txtCustomSearch, (int)Defind_CodeFind.DCF_CUSTOM, "");
        }

        //상단 부품명 check
        private void chkArticleSearch_Checked(object sender, RoutedEventArgs e)
        {
            txtArticleSearch.IsEnabled = true;
            btnPfArticleSearch.IsEnabled = true;
            txtArticleSearch.Focus();
        }

        //상단 부품명 Uncheck
        private void chkArticleSearch_Unchecked(object sender, RoutedEventArgs e)
        {
            txtArticleSearch.Text = "";
            txtArticleSearch.IsEnabled = false;
            btnPfArticleSearch.IsEnabled = false;
        }

        //부품명
        private void lblArticleSearch_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            if (chkArticleSearch.IsChecked == true) { chkArticleSearch.IsChecked = false; }
            else { chkArticleSearch.IsChecked = true; }
        }

        //거래처 check
        private void chkCustomSearch_Checked(object sender, RoutedEventArgs e)
        {
            txtCustomSearch.IsEnabled = true;
            btnPfCustomSearch.IsEnabled = true;
            txtCustomSearch.Focus();
        }

        //거래처 unchekck
        private void chkCustomSearch_Unchecked(object sender, RoutedEventArgs e)
        {
            txtCustomSearch.Text = "";
            txtCustomSearch.IsEnabled = false;
            btnPfCustomSearch.IsEnabled = false;
        }

        //거래처
        private void lblCustomSearch_MouseLeftButtonUP(object sender, MouseButtonEventArgs e)
        {
            if (chkCustomSearch.IsChecked == true) { chkCustomSearch.IsChecked = false; }
            else { chkCustomSearch.IsChecked = true; }
        }

        //상단 입고일자 check
        private void chkInDay_Checked(object sender, RoutedEventArgs e)
        {
            dtpSDate.IsEnabled = true;
            dtpEDate.IsEnabled = true;
        }

        //상단 입고일자 uncheck
        private void chkInDay_Unchecked(object sender, RoutedEventArgs e)
        {
            dtpSDate.IsEnabled = false;
            dtpEDate.IsEnabled = false;
        }

        #endregion

        #region 날짜 버튼

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

        #region 오른쪽 상단 버튼

        //추가
        private void btnAdd_Click(object sender, RoutedEventArgs e)
        {
            if (dgdGroup.Items.Count > 0 && dgdGroup.SelectedIndex > -1)
            {
                numRowCount = dgdGroup.SelectedIndex;   //취소 시 대비
            }

            ControVisibleAndEnable_AU();
            tbkMsg.Text = "자료 입력 (추가)중";
            jobFlag = "I";

            if (chkRemainAddSrh.IsChecked == true) { txtStuffinID.Text = null; }
            else { this.DataContext = null; }


            cboUnit.SelectedValue = "0";
            cboCompanyS.SelectedIndex = 0;
            cboPriceClss.SelectedValue = "0";
            cboVatInd.SelectedValue = "Y";

            cboForUse.SelectedValue = "3"; //2021-06-03 부품용도를 처음에 TOOL로 보이게 변경
            dtpDayUpdate.SelectedDate = DateTime.Today;
            chkDayUpdate.Focus();

            btnTool.IsEnabled = false;
            btnEquip.IsEnabled = false;
        }

        //수정
        private void btnUpdate_Click(object sender, RoutedEventArgs e)
        {
            if (dgdGroup.SelectedItem == null)
            {
                MessageBox.Show("수정할 자료를 먼저 선택해주세요.");
            }
            else
            {
                numRowCount = dgdGroup.SelectedIndex;
                ControVisibleAndEnable_AU();
                tbkMsg.Text = "자료 입력 (수정)중";
                jobFlag = "U";
                chkDayUpdate.Focus();

                btnTool.IsEnabled = false;
                btnEquip.IsEnabled = false;
            }
        }

        //삭제
        private void btnDelete_Click(object sender, RoutedEventArgs e)
        {
            WinPartStuff = dgdGroup.SelectedItem as Win_prd_PartStuffin_U_CodeView;

            if (WinPartStuff == null)
            {
                MessageBox.Show("삭제할 데이터가 지정되지 않았습니다. 삭제데이터를 지정하고 눌러주세요");
                return;
            }
            else if (CheckStockQty(WinPartStuff.StuffInID, WinPartStuff.TotQty, WinPartStuff.MCPartID))
            {
                return;
            }
            else
            {
                if (dgdGroup.SelectedIndex == 0)
                    numRowCount = 0;
                else
                    numRowCount = dgdGroup.SelectedIndex - 1;

                if (MessageBox.Show("선택하신 항목을 삭제하시겠습니까?", "삭제 전 확인", MessageBoxButton.YesNo) == MessageBoxResult.Yes)
                {
                    if (DeleteData(WinPartStuff.StuffInID))
                    {
                        re_Search(numRowCount);
                    }
                }
            }
        }

        //닫기
        private void btnClose_Click(object sender, RoutedEventArgs e)
        {
            Lib.Instance.ChildMenuClose(this.ToString());
        }

        //검색
        private void btnSearch_Click(object sender, RoutedEventArgs e)
        {
            //검색버튼 비활성화
            btnSearch.IsEnabled = false;

            Dispatcher.BeginInvoke(new Action(() =>

            {
                //로직
                btnTool.IsChecked = false;
                btnEquip.IsChecked = false;

                numRowCount = 0;
                re_Search(numRowCount);

            }), System.Windows.Threading.DispatcherPriority.Background);



            Dispatcher.BeginInvoke(new Action(() =>

            {
                btnSearch.IsEnabled = true;

            }), System.Windows.Threading.DispatcherPriority.Background);


        }

        //저장
        private void btnSave_Click(object sender, RoutedEventArgs e)
        {
            if (SaveData())
            {
                ControlVisibalAndEnable_SC();

                if (jobFlag.Equals("I"))
                {
                    numRowCount = 0;
                }

                btnEquip.IsEnabled = true;
                btnTool.IsEnabled = true;
                re_Search(numRowCount);
            }
            else
            {
                MessageBox.Show("저장실패");
            }
        }

        //취소
        private void btnCancel_Click(object sender, RoutedEventArgs e)
        {
            ControlVisibalAndEnable_SC();
            btnEquip.IsEnabled = true;
            btnTool.IsEnabled = true;
            re_Search(numRowCount);
        }

        //엑셀
        private void btnExcel_Click(object sender, RoutedEventArgs e)
        {
            DataTable dt = null;
            string Name = string.Empty;

            string[] lst = new string[2];
            lst[0] = "설비 입고";
            lst[1] = dgdGroup.Name;

            ExportExcelxaml ExpExc = new ExportExcelxaml(lst);

            ExpExc.ShowDialog();

            if (ExpExc.DialogResult.HasValue)
            {
                if (ExpExc.Check.Equals("Y"))
                    dt = Lib.Instance.DataGridToDTinHidden(dgdGroup);
                else
                    dt = Lib.Instance.DataGirdToDataTable(dgdGroup);

                Name = dgdGroup.Name;

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




        #endregion

        #region 메인 데이터그리드

        //메인 데이터그리드 
        private void dgdGroup_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            WinPartStuff = dgdGroup.SelectedItem as Win_prd_PartStuffin_U_CodeView;

            if (WinPartStuff != null)
            {
                string kk = WinPartStuff.StuffDate;

                if (kk.Length < 9)
                {
                    if (int.Parse(kk.Substring(0, 1)) < 5)
                    {
                        WinPartStuff.StuffDate = "20" + kk;
                    }
                    else
                    {
                        WinPartStuff.StuffDate = "19" + kk;
                    }

                }

                if (WinPartStuff.ForUse.Equals("2"))
                {
                    lblName.Content = "설비예비품명";
                }
                else if (WinPartStuff.ForUse.Equals("3"))
                {
                    lblName.Content = "Tool명";
                }
                else
                {
                    lblName.Content = "설비(부품)명";
                }


                this.DataContext = WinPartStuff;
                //txtMCPart.Tag = WinPartStuff.MCPartID; 예비품은 안써서 주석처리 
                //txtMCID.Tag = WinPartStuff.NameID;




            }
        }


        #endregion

        #region 오른쪽 상세메뉴

        //일자수정 check
        private void chkDayUpdate_Checked(object sender, RoutedEventArgs e)
        {
            dtpDayUpdate.IsEnabled = true;
            dtpDayUpdate.Focus();
        }

        //일자수정 uncheck
        private void chkDayUpdate_Unchecked(object sender, RoutedEventArgs e)
        {
            dtpDayUpdate.IsEnabled = false;
        }

        //일자수정 체크박스 keyDown
        private void chkDayUpdate_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                e.Handled = true;

                if (MessageBox.Show("일자를 수정하시겠습니까?", "", MessageBoxButton.YesNo) == MessageBoxResult.Yes)
                {
                    chkDayUpdate.IsChecked = true;
                    dtpDayUpdate.Focus();
                }
                else
                {
                    chkDayUpdate.IsChecked = false;
                    dtpDayUpdate.Focus();
                }
            }
        }

        //일자수정
        private void lblDateUPD_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            if (chkDayUpdate.IsChecked == true) { chkDayUpdate.IsChecked = false; }
            else { chkDayUpdate.IsChecked = true; }
        }

        //캘린더 keyDown
        private void dtpDayUpdate_PreviewKeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                dtpDayUpdate.IsDropDownOpen = true;
            }
        }
        //캘린더 
        private void dtpDayUpdate_CalendarClosed(object sender, RoutedEventArgs e)
        {
            //txtMCID.Focus();
            //cboForUse.Focus();
            //txtMCPart.Focus();
        }

        //설비명
        private void txtMCID_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                if (cboForUse.SelectedValue.Equals("2"))
                {
                    MainWindow.pf.ReturnCode(txtMCID, 13, "2");
                }
                else if (cboForUse.SelectedValue.Equals("3"))
                {
                    MainWindow.pf.ReturnCode(txtMCID, 13, "3");
                }
                else
                {
                    MainWindow.pf.ReturnCode(txtMCID, 13, "1");
                }

                txtCustom.Focus();
            }
        }

        //Tool명/설비예비품명 플러스 파인더 
        private void btnPfMCID_Click(object sender, RoutedEventArgs e)
        {
            DataSet ds = null;
            Dictionary<string, object> sqlParameter = new Dictionary<string, object>();

            if (cboForUse.SelectedValue.Equals("2"))
            {
                MainWindow.pf.ReturnCode(txtMCID, 13, "2");

                sqlParameter.Clear();
                sqlParameter.Add("MCPartName", txtMCID.Text);
                ds = DataStore.Instance.ProcedureToDataSet("xp_mcPartUnitPrice", sqlParameter, false);
                if (ds != null && ds.Tables.Count > 0)
                {
                    DataTable dt = ds.Tables[0];

                    if (dt.Rows.Count > 0)
                    {
                        DataRowCollection drc = dt.Rows;

                        foreach (DataRow dr in drc)
                        {
                            txtUnitPrice.Text = dr["UnitPrice"].ToString();
                        }
                    }
                }

            }
            else if (cboForUse.SelectedValue.Equals("3"))
            {
                MainWindow.pf.ReturnCode(txtMCID, 13, "3");

                sqlParameter.Clear();
                sqlParameter.Add("MCPartName", txtMCID.Text);
                ds = DataStore.Instance.ProcedureToDataSet("xp_mcPartUnitPrice", sqlParameter, false);
                if (ds != null && ds.Tables.Count > 0)
                {
                    DataTable dt = ds.Tables[0];

                    if (dt.Rows.Count > 0)
                    {
                        DataRowCollection drc = dt.Rows;

                        foreach (DataRow dr in drc)
                        {
                            txtUnitPrice.Text = dr["UnitPrice"].ToString();
                        }
                    }
                }
            }
            else
            {
                //MainWindow.pf.ReturnCode(txtMCID, (int)Defind_CodeFind.DCF_MC, "1"); 2021-07-23 수정 
                MainWindow.pf.ReturnCode(txtMCID, 13, "1");

                sqlParameter.Clear();
                sqlParameter.Add("MCPartName", txtMCID.Text);
                ds = DataStore.Instance.ProcedureToDataSet("xp_mcPartUnitPrice", sqlParameter, false);
                if (ds != null && ds.Tables.Count > 0)
                {
                    DataTable dt = ds.Tables[0];

                    if (dt.Rows.Count > 0)
                    {
                        DataRowCollection drc = dt.Rows;

                        foreach (DataRow dr in drc)
                        {
                            txtUnitPrice.Text = dr["UnitPrice"].ToString();
                        }
                    }
                }
            }

            txtCustom.Focus();


        }



        //거래처 keyDown
        private void txtCustom_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                MainWindow.pf.ReturnCode(txtCustom, (int)Defind_CodeFind.DCF_CUSTOM, "");
                txtInQty.Focus();
            }
        }

        //입고수량 keyDown
        // 2021-05-27
        private void txtInQty_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                txtUnitPrice.Focus();
                //cboUnit.Focus();
                //cboUnit.IsDropDownOpen = true;
            }
        }



        //입고 수량 이벤트 이건 뭐지??
        private void txtInQty_PreviewTextInput(object sender, TextCompositionEventArgs e)
        {
            Lib.Instance.CheckIsNumericOnly((TextBox)sender, e);
        }

        //입고단위
        private void cboUnit_DropDownClosed(object sender, EventArgs e)
        {
            txtUnitPrice.Focus();
        }

        //단가 keyDown
        // 2021-05-27
        private void txtUnitPrice_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                txtExchRate.Focus();
                //cboPriceClss.Focus();
                //cboPriceClss.IsDropDownOpen = true;
            }
        }

        //화폐단위
        private void cboPriceClss_DropDownClosed(object sender, EventArgs e)
        {
            txtExchRate.Focus();
        }

        //환율
        //2021-05-27
        private void txtExchRate_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                txtRamark.Focus();
                //cboVatInd.Focus();
                //cboVatInd.IsDropDownOpen = true;
            }
        }

        //부가세 별도
        private void cboVatInd_DropDownClosed(object sender, EventArgs e)
        {
            //cboForUse.Focus();
            txtRamark.Focus();
        }

        //부품용도
        private void cboForUse_DropDownClosed(object sender, EventArgs e)
        {
            //txtRamark.Focus();
            txtMCID.Focus();
        }

        //부품용도
        private void CboForUse_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            //    if (cboForUse.SelectedValue != null)
            //    {
            //        if (cboForUse.SelectedValue.Equals("2"))
            //        {
            //            lblName.Content = "설비예비품명";
            //        }
            //        else if (cboForUse.SelectedValue.Equals("3"))
            //        {
            //            lblName.Content = "Tool명";
            //        }
            //    }
        }

        #endregion

        #region 설비 데이터그리드

        private void btnRight_Click(object sender, RoutedEventArgs e)
        {

        }

        private void btnLeft_Click(object sender, RoutedEventArgs e)
        {

        }

        #endregion

        #region 전체MC 데이터그리드

        #endregion

        #region 기타 나머지들

        //추가,수정 시 공통 동작
        private void ControVisibleAndEnable_AU()
        {
            Lib.Instance.UiButtonEnableChange_SCControl(this);
            //dgdGroup.IsEnabled = false;
            dgdGroup.IsHitTestVisible = false;
            bdrShowData.IsEnabled = true;
        }

        //저장,취소 시 공통 동작
        private void ControlVisibalAndEnable_SC()
        {
            Lib.Instance.UiButtonEnableChange_IUControl(this);
            //dgdGroup.IsEnabled = true;
            dgdGroup.IsHitTestVisible = true;
            bdrShowData.IsEnabled = false;
        }

        //추가와 수정
        private bool SaveData()
        {
            bool flag = true;
            List<Procedure> Prolist = new List<Procedure>();
            List<Dictionary<string, object>> ListParameter = new List<Dictionary<string, object>>();

            WinPartStuff = dgdGroup.SelectedItem as Win_prd_PartStuffin_U_CodeView;
            string stuffID = (WinPartStuff == null ? "" : WinPartStuff.StuffInID);

            if (CheckData())
            {
                try
                {

                    Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                    sqlParameter.Add("nAfftecRows", 0);
                    sqlParameter.Add("JobFlag", jobFlag);
                    sqlParameter.Add("StuffDate", dtpDayUpdate.SelectedDate.Value.ToString("yyyyMMdd"));
                    sqlParameter.Add("StuffInID", jobFlag.Equals("I") ? "" : stuffID);
                    sqlParameter.Add("CustomID", txtCustom.Tag.ToString());

                    sqlParameter.Add("InCustom", txtCustom.Text);
                    sqlParameter.Add("UnitClss", cboUnit.SelectedValue.ToString());
                    sqlParameter.Add("TotRoll", 0);
                    sqlParameter.Add("TotQty", Lib.Instance.IsNumOrAnother(txtInQty.Text) ? double.Parse(txtInQty.Text) : 0);
                    sqlParameter.Add("UnitPrice", Lib.Instance.IsNumOrAnother(txtUnitPrice.Text) ? double.Parse(txtUnitPrice.Text) : 0);

                    sqlParameter.Add("PriceClss", cboPriceClss.SelectedValue.ToString());
                    sqlParameter.Add("ExchRate", Lib.Instance.IsNumOrAnother(txtExchRate.Text) ? double.Parse(txtExchRate.Text) : 0);
                    sqlParameter.Add("Vat_Ind_YN", cboVatInd.SelectedValue.ToString());
                    sqlParameter.Add("Remark", txtRamark.Text);
                    sqlParameter.Add("MCPartID", txtMCID.Tag.ToString());

                    sqlParameter.Add("CompanyID", cboCompanyS.SelectedValue.ToString());
                    sqlParameter.Add("sUserID", MainWindow.CurrentUser);
                    sqlParameter.Add("ArticleSubID", "");
                    sqlParameter.Add("MCID", "");
                    sqlParameter.Add("ToolEndDate", "");
                    sqlParameter.Add("sLOTID", "");

                    if (jobFlag.Equals("I"))
                    {
                        Procedure pro1 = new Procedure();
                        pro1.Name = "xp_mcPartStuffIN_iuStuffIN";
                        pro1.OutputUseYN = "Y";
                        pro1.OutputName = "StuffInID";
                        pro1.OutputLength = "12";

                        Prolist.Add(pro1);
                        ListParameter.Add(sqlParameter);

                        List<KeyValue> list_Result = new List<KeyValue>();
                        list_Result = DataStore.Instance.ExecuteAllProcedureOutputGetCS(Prolist, ListParameter);
                        string sGetStuffInID = string.Empty;

                        if (list_Result[0].key.ToLower() == "success")
                        {
                            list_Result.RemoveAt(0);
                            for (int i = 0; i < list_Result.Count; i++)
                            {
                                KeyValue kv = list_Result[i];
                                if (kv.key == "StuffInID")
                                {
                                    sGetStuffInID = kv.value;
                                    dicCompare.Add("StuffInID", sGetStuffInID);
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
                    else
                    {
                        dicCompare.Add("StuffInID", stuffID);

                        Procedure pro1 = new Procedure();
                        pro1.Name = "xp_mcPartStuffIN_iuStuffIN";
                        pro1.OutputUseYN = "Y";
                        pro1.OutputName = "nAfftecRows";
                        pro1.OutputLength = "10";

                        Prolist.Add(pro1);
                        ListParameter.Add(sqlParameter);

                        List<KeyValue> list_Result = new List<KeyValue>();
                        list_Result = DataStore.Instance.ExecuteAllProcedureOutputGetCS(Prolist, ListParameter);
                        string sGetnumResult = string.Empty;

                        if (list_Result[0].key.ToLower() == "success")
                        {
                            list_Result.RemoveAt(0);
                            for (int i = 0; i < list_Result.Count; i++)
                            {
                                KeyValue kv = list_Result[i];
                                if (kv.key == "nAfftecRows")
                                {
                                    sGetnumResult = kv.value;
                                }
                            }
                        }
                        else
                        {
                            MessageBox.Show("[저장실패]\r\n" + list_Result[0].value.ToString());
                            flag = false;
                            //return false;
                        }

                        if (sGetnumResult.Equals("0") || sGetnumResult.Equals(string.Empty) || sGetnumResult.Equals("9999"))
                        {
                            flag = false;
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
            else
            {
                flag = false;
            }

            return flag;
        }

        //Data Checking
        private bool CheckData()
        {
            bool flag = true;

            if (txtCustom.Text.Equals("") || txtCustom.Tag == null)
            {
                MessageBox.Show("거래처는 필수입력 항목입니다.");
                flag = false;
                return flag;
            }

            //if (txtMCPart.Tag.Equals("") || txtMCPart.Text.Equals(""))
            //{
            //    MessageBox.Show("부품명은 필수입력 항목입니다.");
            //    flag = false;
            //    return flag;
            //}

            if (txtMCID.Tag == null || txtMCID.Tag.ToString().Equals("") || txtMCID.Text.Equals(""))
            {
                MessageBox.Show("설비명은 필수입력 항목입니다.");
                flag = false;
                return flag;
            }

            if (txtInQty.Text.Equals("0") || txtInQty.Text.Equals(""))
            {
                MessageBox.Show("입고 수량이 없습니다. 수량을 입력하십시오");
                flag = false;
                return flag;
            }

            if (cboUnit.SelectedIndex < 0)
            {
                MessageBox.Show("입고 단위를 선택하세요");
                flag = false;
                return flag;
            }

            if (cboForUse.SelectedIndex < 0)
            {
                MessageBox.Show("부품 용도를 선택하세요");
                flag = false;
                return flag;
            }

            if (cboPriceClss.SelectedIndex < 0)
            {
                MessageBox.Show("화폐단위를 선택하세요");
                flag = false;
                return flag;
            }

            if (cboVatInd.SelectedIndex < 0)
            {
                MessageBox.Show("부가세 별도 여부를 선택하세요");
                flag = false;
                return flag;
            }

            return flag;

        }

        //재조회
        private void re_Search(int selectindex)
        {

            FillGrid();

            if (dgdGroup.Items.Count > 0)
            {
                dgdGroup.Focus();
                if (lstCompareValue.Count > 0)
                {
                    dgdGroup.SelectedIndex = Lib.Instance.reTrunIndex(dgdGroup, lstCompareValue[0]);
                }
                else
                {
                    dgdGroup.SelectedIndex = selectindex; ;
                }
            }
            else
            {
                this.DataContext = null;

                MessageBox.Show("조회된 데이터가 없습니다.");
                return;
            }

            dicCompare.Clear();
            lstCompareValue.Clear();
        }

        private void FillGrid()
        {

            if (dgdGroup != null && dgdGroup.Items.Count > 0)
            {
                dgdGroup.Items.Clear();
            }

            try
            {
                int nForUse = 0;
                string sForUse = string.Empty;

                if (!cboForUseSrh.SelectedValue.Equals("0"))
                {
                    nForUse = 1;
                }

                if (cboForUseSrh.SelectedValue != null)
                {
                    sForUse = cboForUseSrh.SelectedValue.ToString();
                }

                DataSet ds = null;
                Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                sqlParameter.Clear();
                sqlParameter.Add("nChkDate", (chkInDay.IsChecked == true) ? 1 : 0);
                sqlParameter.Add("sSDate", (chkInDay.IsChecked == true) ?
                    dtpSDate.SelectedDate.Value.ToString("yyyyMMdd") : "");
                sqlParameter.Add("sEDate", (chkInDay.IsChecked == true) ?
                    dtpEDate.SelectedDate.Value.ToString("yyyyMMdd") : "");
                sqlParameter.Add("nChkCustom", (chkCustomSearch.IsChecked == true) ?
                    (txtCustomSearch.Tag != null ? 1 : 2) : 0);
                sqlParameter.Add("sCustom", (chkCustomSearch.IsChecked == true) ?
                    (txtCustomSearch.Tag != null ? txtCustomSearch.Tag.ToString() : txtCustomSearch.Text) : "");
                sqlParameter.Add("nChkMCPartID", (chkArticleSearch.IsChecked == true) ?
                    (txtArticleSearch.Tag != null ? 1 : 2) : 0);
                sqlParameter.Add("sMCPartID", (chkArticleSearch.IsChecked == true) ?
                    (txtArticleSearch.Tag != null ? txtArticleSearch.Tag.ToString() : txtArticleSearch.Text) : "");
                sqlParameter.Add("nChkForUse", nForUse);
                sqlParameter.Add("sForUse", sForUse);

                ds = DataStore.Instance.ProcedureToDataSet("xp_mcPartStuffIN_sStuffIN", sqlParameter, false);

                //Tool버튼 클릭
                if (btnTool.IsChecked == true)
                {
                    sqlParameter.Clear();
                    sqlParameter.Add("nChkDate", (chkInDay.IsChecked == true) ? 1 : 0);
                    sqlParameter.Add("sSDate", (chkInDay.IsChecked == true) ?
                        dtpSDate.SelectedDate.Value.ToString("yyyyMMdd") : "");
                    sqlParameter.Add("sEDate", (chkInDay.IsChecked == true) ?
                        dtpEDate.SelectedDate.Value.ToString("yyyyMMdd") : "");
                    sqlParameter.Add("nChkCustom", (chkCustomSearch.IsChecked == true) ?
                        (txtCustomSearch.Tag != null ? 1 : 2) : 0);
                    sqlParameter.Add("sCustom", (chkCustomSearch.IsChecked == true) ?
                        (txtCustomSearch.Tag != null ? txtCustomSearch.Tag.ToString() : txtCustomSearch.Text) : "");
                    sqlParameter.Add("nChkMCPartID", (chkArticleSearch.IsChecked == true) ?
                        (txtArticleSearch.Tag != null ? 1 : 2) : 0);
                    sqlParameter.Add("sMCPartID", (chkArticleSearch.IsChecked == true) ?
                        (txtArticleSearch.Tag != null ? txtArticleSearch.Tag.ToString() : txtArticleSearch.Text) : "");
                    sqlParameter.Add("nChkForUse", "1");
                    sqlParameter.Add("sForUse", "3");

                    ds = DataStore.Instance.ProcedureToDataSet("xp_mcPartStuffIN_sStuffIN", sqlParameter, false);
                }

                //설비등록 버튼 클릭
                if (btnEquip.IsChecked == true)
                {
                    sqlParameter.Clear();
                    sqlParameter.Add("nChkDate", (chkInDay.IsChecked == true) ? 1 : 0);
                    sqlParameter.Add("sSDate", (chkInDay.IsChecked == true) ?
                        dtpSDate.SelectedDate.Value.ToString("yyyyMMdd") : "");
                    sqlParameter.Add("sEDate", (chkInDay.IsChecked == true) ?
                        dtpEDate.SelectedDate.Value.ToString("yyyyMMdd") : "");
                    sqlParameter.Add("nChkCustom", (chkCustomSearch.IsChecked == true) ?
                        (txtCustomSearch.Tag != null ? 1 : 2) : 0);
                    sqlParameter.Add("sCustom", (chkCustomSearch.IsChecked == true) ?
                        (txtCustomSearch.Tag != null ? txtCustomSearch.Tag.ToString() : txtCustomSearch.Text) : "");
                    sqlParameter.Add("nChkMCPartID", (chkArticleSearch.IsChecked == true) ?
                        (txtArticleSearch.Tag != null ? 1 : 2) : 0);
                    sqlParameter.Add("sMCPartID", (chkArticleSearch.IsChecked == true) ?
                        (txtArticleSearch.Tag != null ? txtArticleSearch.Tag.ToString() : txtArticleSearch.Text) : "");
                    sqlParameter.Add("nChkForUse", "1");
                    sqlParameter.Add("sForUse", "2");

                    ds = DataStore.Instance.ProcedureToDataSet("xp_mcPartStuffIN_sStuffIN", sqlParameter, false);
                }

                if (ds != null && ds.Tables.Count > 0)
                {
                    DataTable dt = ds.Tables[0];

                    if (dt.Rows.Count > 0)
                    {
                        DataRowCollection drc = dt.Rows;
                        int i = 0;

                        foreach (DataRow dr in drc)
                        {
                            i++;

                            //string strQty = dr["TotQty"].ToString();
                            //StringBuilder sbQty = new StringBuilder();
                            //sbQty.Append(strQty.Substring(0, strQty.IndexOf('.')));
                            //sbQty.Append(strQty.Substring(strQty.IndexOf('.'), 3));

                            string strQtyY = dr["TotQtyY"].ToString();
                            StringBuilder sbQtyY = new StringBuilder();
                            sbQtyY.Append(strQtyY.Substring(0, strQtyY.IndexOf('.')));
                            sbQtyY.Append(strQtyY.Substring(strQtyY.IndexOf('.'), 3));

                            //string strAmount = dr["amount"].ToString();
                            //StringBuilder sbAmount = new StringBuilder();
                            //sbAmount.Append(strAmount.Substring(0, strAmount.IndexOf('.')));
                            //sbAmount.Append(strAmount.Substring(strAmount.IndexOf('.'), 3));

                            //string strUnitPrice = string.Format("{0:n0}", double.Parse(dr["UnitPrice"].ToString()));

                            string strPriceClss = string.Empty;

                            if (!dr["PriceClss"].ToString().Equals("1"))
                            {
                                strPriceClss = dr["PriceClss"].ToString();
                            }

                            var partStuffIn = new Win_prd_PartStuffin_U_CodeView()
                            {
                                Num = i.ToString(),
                                StuffInID = dr["StuffInID"].ToString(),
                                StuffDate = dr["StuffDate"].ToString(),
                                CompanyID = dr["CompanyID"].ToString(),
                                StuffClss = dr["StuffClss"].ToString(),
                                //stuffName = dr["stuffName"].ToString(),
                                CustomID = dr["CustomID"].ToString(),
                                Custom = dr["Custom"].ToString(),
                                MCPartID = dr["MCPartID"].ToString(),
                                UnitClss = dr["UnitClss"].ToString(),
                                TotRoll = dr["TotRoll"].ToString(),
                                TotQty = Convert.ToDouble(dr["TotQty"]), //sbQty.ToString(),
                                TotQtyY = sbQtyY.ToString(),
                                UnitClssName = dr["UnitClssName"].ToString(),
                                PriceClss = strPriceClss,
                                PriceClssName = dr["PriceClssName"].ToString(),
                                UnitPrice = Convert.ToDouble(dr["UnitPrice"]), //strUnitPrice,
                                ExchRate = Convert.ToDouble(dr["ExchRate"]), //dr["ExchRate"].ToString(),
                                Vat_Ind_YN = dr["Vat_Ind_YN"].ToString(),
                                Remark = dr["Remark"].ToString(),
                                StuffClssName = dr["StuffClssName"].ToString(),
                                MCPartName = dr["MCPartName"].ToString(),
                                KCompany = dr["KCompany"].ToString(),
                                CustomName = dr["CustomName"].ToString(),
                                amount = Convert.ToDouble(dr["amount"]),  //sbAmount.ToString(),
                                ArticleSubID = dr["ArticleSubID"].ToString(),
                                ArticleSubName = dr["ArticleSubName"].ToString(),
                                ForUse = dr["ForUse"].ToString(),
                                ForUseName = dr["ForUseName"].ToString(),
                                MCID = dr["MCID"].ToString(),
                                MCNAME = dr["MCNAME"].ToString(),
                                //Name = dr["Name"].ToString(),
                                //NameID = dr["NameID"].ToString()
                            };

                            partStuffIn.StuffDate = Lib.Instance.StrDateTimeBar(partStuffIn.StuffDate);
                            partStuffIn.amount = partStuffIn.UnitPrice * partStuffIn.TotQty;

                            if (dicCompare.Count > 0)
                            {
                                if (partStuffIn.StuffInID.Equals(dicCompare["StuffInID"].ToString()))
                                {
                                    lstCompareValue.Add(partStuffIn.ToString());
                                }
                            }

                            dgdGroup.Items.Add(partStuffIn);
                        }
                        tbkCount.Text = "▶ 검색결과 : " + i.ToString() + " 건";
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

        //삭제
        private bool DeleteData(string StuffID)
        {
            bool flag = true;

            try
            {
                Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                sqlParameter.Add("StuffInID", StuffID);

                string[] result = DataStore.Instance.ExecuteProcedure("xp_mcPartStuffIN_dStuffIN", sqlParameter, true);

                if (!result[0].Equals("success"))
                {
                    flag = false;
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

            return flag;
        }

        //private void InputClear() // 인풋박스 초기화 
        //{
        //    try
        //    {
        //        foreach (Control child in grdInput.Children)
        //        {
        //            if (child.GetType() == typeof(TextBox))
        //                ((TextBox)child).Clear();
        //            else if (child.GetType() == typeof(ComboBox))
        //                ((ComboBox)child).SelectedIndex = -1;
        //            else if (child.GetType() == typeof(CheckBox))
        //                ((CheckBox)child).IsChecked = false;
        //            else if (child.GetType() == typeof(Button))
        //                continue;
        //        }
        //    }
        //    catch (Exception ex)
        //    {
        //        MessageBox.Show("에러메시지 : " + ex.Message);
        //        MessageBox.Show("" + grdInput.Children.GetType());
        //    }
        //    chkDayUpdate.IsChecked = false;
        //}

        #endregion

        private bool CheckStockQty(string strStuffInID, double strQty, string strMCPartID)
        {
            bool result = false;

            try
            {
                Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                sqlParameter.Clear();
                sqlParameter.Add("StuffInID", strStuffInID);
                sqlParameter.Add("Qty", strQty);
                sqlParameter.Add("MCPartID", strMCPartID);

                DataSet ds = DataStore.Instance.ProcedureToDataSet("xp_mcPartStuffIn_CheckStockQty", sqlParameter, false);

                if (ds != null && ds.Tables.Count > 0)
                {
                    DataTable dt = ds.Tables[0];
                    DataRow dr = dt.Rows[0];

                    if (dt.Rows.Count > 0 && !dr["Msg"].ToString().ToUpper().Equals(""))
                    {
                        MessageBox.Show("[저장실패]\r\n" + dr["Msg"].ToString());
                        result = true;
                    }
                    else
                    {
                        result = false;
                    }
                }


            }
            catch (Exception ee)
            {
                MessageBox.Show("예외처리 - " + ee.ToString());
            }
            finally
            {
                DataStore.Instance.CloseConnection();
            }

            return result;
        }


        #region 생성자

        class Win_prd_PartStuffin_U_CodeView : BaseView
        {
            public override string ToString()
            {
                return (this.ReportAllProperties());
            }

            public string Num { get; set; }
            public string StuffInID { get; set; }
            public string StuffDate { get; set; }
            public string CompanyID { get; set; }
            public string StuffClss { get; set; }

            //public string stuffName { get; set; }
            public string CustomID { get; set; }
            public string Custom { get; set; }
            public string MCPartID { get; set; }
            public string UnitClss { get; set; }
            public string TotRoll { get; set; }

            public double TotQty { get; set; }
            public string TotQtyY { get; set; }
            public string UnitClssName { get; set; }
            public string PriceClss { get; set; }
            public string PriceClssName { get; set; }

            public double UnitPrice { get; set; }
            public double ExchRate { get; set; }
            public string Vat_Ind_YN { get; set; }
            public string Remark { get; set; }
            public string StuffClssName { get; set; }

            public string MCPartName { get; set; }
            public string KCompany { get; set; }
            public string CustomName { get; set; }
            public double amount { get; set; }
            public string ForUse { get; set; }

            public string ForUseName { get; set; }
            public string ArticleSubID { get; set; }
            public string ArticleSubName { get; set; }
            public string MCID { get; set; }
            public string MCNAME { get; set; }

            public string NameID { get; set; }
            public string Name { get; set; }
        }

        #endregion

        private void lblRemainAddSrh_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            if (chkRemainAddSrh.IsChecked == true) { chkRemainAddSrh.IsChecked = false; }
            else { chkRemainAddSrh.IsChecked = true; }
        }
    }
}

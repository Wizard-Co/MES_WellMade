using WizMes_WellMade.PopUP;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Data;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using WPF.MDI;

namespace WizMes_WellMade
{
    /// <summary>
    /// Win_dvl_MoldRepair_U1.xaml에 대한 상호 작용 논리
    /// </summary>
    public partial class Win_dvl_MoldRepair_U : UserControl
    {
        #region 변수선언 및 로드

        Win_dvl_MoldRepair_U_CodeView WinRepair = new Win_dvl_MoldRepair_U_CodeView();
        Win_dvl_MoldRepair_U_Sub_CodeView WinMoldRepairSub = new Win_dvl_MoldRepair_U_Sub_CodeView();
        Lib lib = new Lib();

        string strFlag = string.Empty;
        int rowNum = 0;

        Dictionary<string, object> dicCompare = new Dictionary<string, object>();

        public Win_dvl_MoldRepair_U()
        {
            InitializeComponent();
        }

        private void UserControl_Loaded(object sender, RoutedEventArgs e)
        {
            dtpSDate.SelectedDate = DateTime.Today;
            dtpEDate.SelectedDate = DateTime.Today;
            chkRepairDaySrh.IsChecked = true;
            setComboBox();
            lib.UiLoading(sender);
        }

        private void setComboBox()
        {
            List<string[]> lstDvlYN = new List<string[]>();
            string[] strDvl_1 = { "Y", "개발" };
            string[] strDvl_2 = { "N", "양산" };
            lstDvlYN.Add(strDvl_1);
            lstDvlYN.Add(strDvl_2);

            ObservableCollection<CodeView> ovcDvlYN = ComboBoxUtil.Instance.Direct_SetComboBox(lstDvlYN);
            this.cboDvlYNSrh.ItemsSource = ovcDvlYN;
            this.cboDvlYNSrh.DisplayMemberPath = "code_name";
            this.cboDvlYNSrh.SelectedValuePath = "code_id";

            List<string[]> lstRepairGubun = new List<string[]>();
            string[] strRepairGubun_1 = { "1", "수리" };
            string[] strRepairGubun_2 = { "2", "교체" };
            lstRepairGubun.Add(strRepairGubun_1);
            lstRepairGubun.Add(strRepairGubun_2);

            ObservableCollection<CodeView> ovcRepairGubun = ComboBoxUtil.Instance.Direct_SetComboBox(lstRepairGubun);
            this.cboRepairGubun.ItemsSource = ovcRepairGubun;
            this.cboRepairGubun.DisplayMemberPath = "code_name";
            this.cboRepairGubun.SelectedValuePath = "code_id";
        }

        #endregion

        #region 상단 체크박스 관련 이벤트

        //수리일 라벨
        private void lblRepairDaySrh_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            if (chkRepairDaySrh.IsChecked == true) { chkRepairDaySrh.IsChecked = false; }
            else { chkRepairDaySrh.IsChecked = true; }
        }

        //수리일 체크박스
        private void chkRepairDaySrh_Checked(object sender, RoutedEventArgs e)
        {
            dtpSDate.IsEnabled = true;
            dtpEDate.IsEnabled = true;
        }

        //수리일 체크박스
        private void chkRepairDaySrh_Unchecked(object sender, RoutedEventArgs e)
        {
            dtpSDate.IsEnabled = false;
            dtpEDate.IsEnabled = false;
        }

        //금월
        private void btnThisMonth_Click(object sender, RoutedEventArgs e)
        {
            dtpSDate.SelectedDate = lib.BringThisMonthDatetimeList()[0];
            dtpEDate.SelectedDate = lib.BringThisMonthDatetimeList()[1];
        }

        //금일
        private void btnToday_Click(object sender, RoutedEventArgs e)
        {
            dtpSDate.SelectedDate = DateTime.Today;
            dtpEDate.SelectedDate = DateTime.Today;
        }

        //개발/양산 라벨
        private void lblDvlYNSrh_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            if (chkDvlYNSrh.IsChecked == true) { chkDvlYNSrh.IsChecked = false; }
            else { chkDvlYNSrh.IsChecked = true; }
        }

        //개발/양산 체크박스
        private void chkDvlYNSrh_Checked(object sender, RoutedEventArgs e)
        {
            cboDvlYNSrh.IsEnabled = true;
            cboDvlYNSrh.Focus();
        }

        //개발/양산 체크박스
        private void chkDvlYNSrh_Unchecked(object sender, RoutedEventArgs e)
        {
            cboDvlYNSrh.IsEnabled = false;
        }

        private void lblArticleSabunSrh_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            if (chkArticleSabunSrh.IsChecked == true) { chkArticleSabunSrh.IsChecked = false; }
            else { chkArticleSabunSrh.IsChecked = true; }
        }

        private void chkArticleSabunSrh_Checked(object sender, RoutedEventArgs e)
        {
            txtArticleSabunSrh.IsEnabled = true;
            btnPfArticleSabunSrh.IsEnabled = true;
            txtArticleSabunSrh.Focus();
        }

        private void chkArticleSabunSrh_Unchecked(object sender, RoutedEventArgs e)
        {
            txtArticleSabunSrh.IsEnabled = false;
            btnPfArticleSabunSrh.IsEnabled = false;
        }

        private void txtArticleSabunSrh_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                MainWindow.pf.ReturnCode(txtArticleSabunSrh, 79, "");
            }
        }

        private void btnPfArticleSabunSrh_Click(object sender, RoutedEventArgs e)
        {
            MainWindow.pf.ReturnCode(txtArticleSabunSrh, 79, "");
        }

        //금형 라벨(금형임_변수는 처음에 잘못알아서 article)
        private void lblArticelSrh_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            if (chkArticleSrh.IsChecked == true) { chkArticleSrh.IsChecked = false; }
            else { chkArticleSrh.IsChecked = true; }
        }

        //금형 체크박스(금형임_변수는 처음에 잘못알아서 article)
        private void chkArticleSrh_Checked(object sender, RoutedEventArgs e)
        {
            txtArticelSrh.IsEnabled = true;
            btnPfArticelSrh.IsEnabled = true;
            txtArticelSrh.Focus();
        }

        //금형 체크박스(금형임_변수는 처음에 잘못알아서 article)
        private void chkArticleSrh_Unchecked(object sender, RoutedEventArgs e)
        {
            txtArticelSrh.IsEnabled = false;
            btnPfArticelSrh.IsEnabled = false;
        }

        //금형 텍스트박스(금형임_변수는 처음에 잘못알아서 article)
        private void txtArticelSrh_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                MainWindow.pf.ReturnCode(txtArticelSrh,51,"");
            }
        }

        //금형 플러스파인더(금형임_변수는 처음에 잘못알아서 article)
        private void btnPfArticelSrh_Click(object sender, RoutedEventArgs e)
        {
            MainWindow.pf.ReturnCode(txtArticelSrh, 51, "");
        }

        //금형LotNo 라벨
        private void lblMoldNoSrh_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            if (chkMoldNoSrh.IsChecked == true) { chkMoldNoSrh.IsChecked = false; }
            else { chkMoldNoSrh.IsChecked = true; }
        }

        //금형LotNo 체크박스
        private void chkMoldNoSrh_Checked(object sender, RoutedEventArgs e)
        {
            txtMoldNoSrh.IsEnabled = true;
            txtMoldNoSrh.Focus();
        }

        //금형LotNo 체크박스
        private void chkMoldNoSrh_Unchecked(object sender, RoutedEventArgs e)
        {
            txtMoldNoSrh.IsEnabled = false;
        }

        #endregion

        #region 상단 버튼 이벤트 && 관련 동작 

        //추가,수정 시 동작 모음
        private void ControlVisibleAndEnable_AU()
        {
            lib.UiButtonEnableChange_SCControl(this);
            //btnSave.Visibility = Visibility.Visible;
            //btnCancel.Visibility = Visibility.Visible;
            //btnAdd.IsEnabled = false;
            //btnUpdate.IsEnabled = false;
            //btnDelete.IsEnabled = false;

            dgdMoldRepair.IsEnabled = false;
            btnSubAdd.IsEnabled = true;
            btnSubDel.IsEnabled = true;
            //dgdMoldRepairSub.IsEnabled = true;

            bdrLeft.IsEnabled = true;
            //lblMsg.Visibility = Visibility.Visible;
            //btnExcel.Visibility = Visibility.Hidden;
        }

        //저장,취소 시 동작 모음
        private void ControlVisibleAndEnable_SC()
        {
            lib.UiButtonEnableChange_IUControl(this);
            //btnSave.Visibility = Visibility.Hidden;
            //btnCancel.Visibility = Visibility.Hidden;
            //btnAdd.IsEnabled = true;
            //btnUpdate.IsEnabled = true;
            //btnDelete.IsEnabled = true;

            dgdMoldRepair.IsEnabled = true;
            btnSubAdd.IsEnabled = false;
            btnSubDel.IsEnabled = false;
            //dgdMoldRepairSub.IsEnabled = false;

            bdrLeft.IsEnabled = false;
            //lblMsg.Visibility = Visibility.Hidden;
            //btnExcel.Visibility = Visibility.Visible;
        }

        //추가
        private void btnAdd_Click(object sender, RoutedEventArgs e)
        {
            ControlVisibleAndEnable_AU();
            dtpRepairDate.SelectedDate = DateTime.Today;
            strFlag = "I";

            txtMold.Clear();
            txtMoldQuality.Clear();
            txtProdCustomName.Clear();
            txtRepairCustom.Clear();
            txtRepairID.Clear();
            txtSpec.Clear();
            txtWeight.Clear();
            txtRepairremark.Clear();

            if (dgdMoldRepairSub.Items.Count > 0)
            {
                dgdMoldRepairSub.Items.Clear();
            }

            //cboRepairGubun.SelectedIndex = 0;

            tbkMsg.Text = "자료 입력(추가) 중";

            //취소 시 추가 전 자료를 선택하기 위해
            if (dgdMoldRepair.SelectedItem != null)
            {
                rowNum = dgdMoldRepair.SelectedIndex;
            }

            cboRepairGubun.Focus();
            cboRepairGubun.IsDropDownOpen = true;
        }

        //수정
        private void btnUpdate_Click(object sender, RoutedEventArgs e)
        {
            if (dgdMoldRepair.SelectedItem == null)
            {
                MessageBox.Show("수정할 자료가 지정되지 않았습니다.수정할 자료를 선택해주세요");
            }
            else
            {
                ControlVisibleAndEnable_AU();
                strFlag = "U";

                tbkMsg.Text = "자료 입력(수정) 중";
                rowNum = dgdMoldRepair.SelectedIndex;

                cboRepairGubun.Focus();
                cboRepairGubun.IsDropDownOpen = true;
            }
        }

        //삭제
        private void btnDelete_Click(object sender, RoutedEventArgs e)
        {
            if (dgdMoldRepair.SelectedItem == null )
            {
                MessageBox.Show("삭제할 데이터가 지정되지 않았습니다. 삭제데이터를 지정하고 눌러주세요");
            }
            else
            {
                WinRepair = dgdMoldRepair.SelectedItem as Win_dvl_MoldRepair_U_CodeView;

                if (MessageBox.Show("선택하신 항목을 삭제하시겠습니까?", "삭제 전 확인", MessageBoxButton.YesNo) == MessageBoxResult.Yes)
                {
                    if (dgdMoldRepair.Items.Count > 0 && dgdMoldRepair.SelectedItem != null)
                    {
                        rowNum = dgdMoldRepair.SelectedIndex;
                    }

                    if (DeleteData(WinRepair.RepairID))
                    {
                        if (rowNum > 0)
                        {
                            rowNum -= 1;
                        }
                        re_Search(rowNum);
                    }
                }
            }
        }

        //닫기
        private void btnClose_Click(object sender, RoutedEventArgs e)
        {
            string stDate = DateTime.Now.ToString("yyyyMMdd");
            string stTime = DateTime.Now.ToString("HHmm");
            DataStore.Instance.InsertLogByFormS(this.GetType().Name, stDate, stTime, "E");
            Lib.Instance.ChildMenuClose(this.ToString());
        }

        //조회
        private void btnSearch_Click(object sender, RoutedEventArgs e)
        {
            rowNum = 0;

            re_Search(rowNum);
        }

        //저장
        private void btnSave_Click(object sender, RoutedEventArgs e)
        {
            if (strFlag.Equals("I"))
            {
                if (InsertData())
                {
                    rowNum = 0; //추가 성공 시에는 첫 자료부터 조회

                    ControlVisibleAndEnable_SC();
                    re_Search(rowNum);
                }
            }
            else  //U
            {
                if (UpdateData())
                {
                    ControlVisibleAndEnable_SC();
                    re_Search(rowNum);
                }
            }
        }

        //취소
        private void btnCancel_Click(object sender, RoutedEventArgs e)
        {
            InputClear();
            ControlVisibleAndEnable_SC();

            re_Search(rowNum);
        }

        //입력 데이터 클리어
        private void InputClear()
        {
            foreach (Control child in this.grdInput1.Children)
            {
                if (child.GetType() == typeof(ComboBox))
                    ((ComboBox)child).SelectedIndex = -1;
                else if (child.GetType() == typeof(TextBox))
                    ((TextBox)child).Clear();
            }
            foreach (Control child in this.grdInput2.Children)
            {
                if (child.GetType() == typeof(ComboBox))
                    ((ComboBox)child).SelectedIndex = -1;
                else if (child.GetType() == typeof(TextBox))
                    ((TextBox)child).Clear();
            }
            foreach (Control child in this.grdInput3.Children)
            {
                if (child.GetType() == typeof(ComboBox))
                    ((ComboBox)child).SelectedIndex = -1;
                else if (child.GetType() == typeof(TextBox))
                    ((TextBox)child).Clear();
            }
            if (this.dgdMoldRepairSub.Items.Count <= 0)
                return;
            this.dgdMoldRepairSub.Items.Clear();
        }

        //엑셀
        private void btnExcel_Click(object sender, RoutedEventArgs e)
        {
            DataTable dt = null;
            string Name = string.Empty;

            string[] lst = new string[4];
            lst[0] = "금형수리등록_메인";
            lst[1] = "금형수리등록_부품";
            lst[2] = dgdMoldRepair.Name;
            lst[3] = dgdMoldRepairSub.Name;

            ExportExcelxaml ExpExc = new ExportExcelxaml(lst);

            ExpExc.ShowDialog();

            if (ExpExc.DialogResult.HasValue)
            {
                if (ExpExc.choice.Equals(dgdMoldRepair.Name))
                {
                    //MessageBox.Show("대분류");
                    if (ExpExc.Check.Equals("Y"))
                        dt = lib.DataGridToDTinHidden(dgdMoldRepair);
                    else
                        dt = lib.DataGirdToDataTable(dgdMoldRepair);

                    Name = dgdMoldRepair.Name;

                    if (lib.GenerateExcel(dt, Name))
                        lib.excel.Visible = true;
                    else
                        return;
                }
                else if (ExpExc.choice.Equals(dgdMoldRepairSub.Name))
                {
                    //MessageBox.Show("소분류");
                    if (ExpExc.Check.Equals("Y"))
                        dt = lib.DataGridToDTinHidden(dgdMoldRepairSub);
                    else
                        dt = lib.DataGirdToDataTable(dgdMoldRepairSub);
                    Name = dgdMoldRepairSub.Name;

                    if (lib.GenerateExcel(dt, Name))
                        lib.excel.Visible = true;
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

        // 조회   ,   추가,수정,삭제 후 재조회
        private void re_Search(int selectIndex)
        {
            if (dgdMoldRepair.Items.Count > 0)
            {
                dgdMoldRepair.Items.Clear();
            }

            FillGrid();

            if (dgdMoldRepair.Items.Count > 0)
            {
               
                    dgdMoldRepair.SelectedIndex = selectIndex;
                
            }
            else
            {
                InputClear();
            }

            dicCompare.Clear();

        }

        #endregion

        #region DB Data CRUD

        //조회
        private void FillGrid()
        {
            try
            {
                Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                sqlParameter.Add("ChkDate", (chkRepairDaySrh.IsChecked == true ? 1 : 0));
                sqlParameter.Add("FromDate", (chkRepairDaySrh.IsChecked == true ? dtpSDate.SelectedDate.Value.ToString("yyyyMMdd") : ""));
                sqlParameter.Add("ToDate", (chkRepairDaySrh.IsChecked == true ? dtpEDate.SelectedDate.Value.ToString("yyyyMMdd") : ""));
                sqlParameter.Add("ChkMoldNo", (chkMoldNoSrh.IsChecked == true ? 1 : 0));
                sqlParameter.Add("MoldNo", (chkMoldNoSrh.IsChecked == true ? txtMoldNoSrh.Text : ""));
                sqlParameter.Add("ChkDvlYN", (chkDvlYNSrh.IsChecked == true ? 1 : 0));
                sqlParameter.Add("DvlYN", (chkDvlYNSrh.IsChecked == true ? (cboDvlYNSrh.SelectedValue != null ? cboDvlYNSrh.SelectedValue.ToString() : "") : ""));
                sqlParameter.Add("ChkArticle", (chkArticleSrh.IsChecked == true ? (txtArticelSrh.Tag != null ? 1 : 2) : 0));
                sqlParameter.Add("ArticleID", (chkArticleSrh.IsChecked == true ? (txtArticelSrh.Tag != null ? txtArticelSrh.Tag.ToString() : txtArticelSrh.Text) : ""));
                sqlParameter.Add("ChkArticleSabun", (chkArticleSabunSrh.IsChecked == true ? 1 : 0));
                sqlParameter.Add("ArticleSabun", chkArticleSabunSrh.IsChecked == true ? txtArticleSabunSrh.Text : "");

                DataSet ds = DataStore.Instance.ProcedureToDataSet("xp_dvlMold_sMoldRepair", sqlParameter, false);

                if (ds != null && ds.Tables.Count > 0)
                {
                    DataTable dt = ds.Tables[0];

                    if (dt.Rows.Count > 0)
                    {
                        DataRowCollection drc = dt.Rows;

                        foreach (DataRow dr in drc)
                        {
                            var MoldRepair_DTO = new Win_dvl_MoldRepair_U_CodeView()
                            {
                                RepairID = dr["RepairID"].ToString(),
                                repairdate = dr["repairdate"].ToString(),
                                RepairGubun = dr["RepairGubun"].ToString(),
                                RepairGubunname = dr["RepairGubunname"].ToString(),
                                MoldID = dr["MoldID"].ToString(),
                                RepairCustom = dr["RepairCustom"].ToString(),
                                Repairremark = dr["Repairremark"].ToString(),
                                MoldKind = dr["MoldKind"].ToString(),
                                MoldQuality = dr["MoldQuality"].ToString(),
                                Weight = dr["Weight"].ToString(),
                                Spec = dr["Spec"].ToString(),
                                MoldNo = dr["MoldNo"].ToString(),
                                ProdCustomName = dr["ProdCustomName"].ToString(),
                                Article = dr["Article"].ToString(),
                                ArticleID = dr["ArticleID"].ToString(),
                                Article_Sabun = dr["Article_Sabun"].ToString()
                            };

                            MoldRepair_DTO.repairdate = lib.StrDateTimeBar(MoldRepair_DTO.repairdate);

                            dgdMoldRepair.Items.Add(MoldRepair_DTO);
                        }
                    }
                    else
                    {
                        MessageBox.Show("조회된 결과가 없습니다.");
                        return;
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

        //메인 DataGrid에서 행이 바뀔때
        private void dgdMoldRepair_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            Win_dvl_MoldRepair_U_CodeView winMoldRepair = dgdMoldRepair.SelectedItem as Win_dvl_MoldRepair_U_CodeView;
            if (winMoldRepair != null)
            {
                if (dgdMoldRepairSub.Items.Count > 0)
                {
                    dgdMoldRepairSub.Items.Clear();
                }

                GetArticleMoldData(winMoldRepair.MoldID);

                FillGridSub(winMoldRepair.RepairID);

                this.DataContext = winMoldRepair;
            }
        }

        //Sub 조회
        private void FillGridSub(string strRepairID)
        {
            try
            {
                Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                sqlParameter.Add("RepairID", strRepairID);
                sqlParameter.Add("RepairSubSeq", 0);

                DataSet ds = DataStore.Instance.ProcedureToDataSet("xp_dvlMold_sMoldRepairSub", sqlParameter, false);

                if (ds != null && ds.Tables.Count > 0)
                {
                    DataTable dt = ds.Tables[0];

                    if (dt.Rows.Count > 0)
                    {
                        DataRowCollection drc = dt.Rows;

                        foreach (DataRow dr in drc)
                        {
                            var dgdSub_DTO = new Win_dvl_MoldRepair_U_Sub_CodeView()
                            {
                                RepairID = dr["RepairID"].ToString(),
                                RepairSubSeq = dr["RepairSubSeq"].ToString(),
                                MCPartName = dr["MCPartName"].ToString(),
                                InCustomName = dr["InCustomName"].ToString(),
                                partcnt = dr["partcnt"].ToString(),
                                partprice = dr["partprice"].ToString(),
                                reason = dr["reason"].ToString(),
                                partremark = dr["partremark"].ToString(),
                                customid = dr["customid"].ToString(),
                                McPartid = dr["McPartid"].ToString()
                            };

                            dgdSub_DTO.flagCustom = false;
                            dgdSub_DTO.flagMcPart = false;
                            dgdSub_DTO.flagPartcnt = false;
                            dgdSub_DTO.flagPartprice = false;
                            dgdSub_DTO.flagPartremark = false;
                            dgdSub_DTO.flagReason = false;

                            dgdMoldRepairSub.Items.Add(dgdSub_DTO);
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

        //추가
        private bool InsertData()
        {
            bool flag = true;
            List<Procedure> Prolist = new List<Procedure>();
            List<Dictionary<string, object>> ListParameter = new List<Dictionary<string, object>>();

            if (CheckData())
            {
                try
                {
                    Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                    sqlParameter.Add("RepairID", "");
                    sqlParameter.Add("repairdate", dtpRepairDate.SelectedDate.Value.ToString("yyyyMMdd"));
                    sqlParameter.Add("RepairGubun", cboRepairGubun.SelectedValue.ToString());
                    sqlParameter.Add("MoldID", txtMold.Tag.ToString());
                    sqlParameter.Add("RepairCustom", txtRepairCustom.Text);
                    sqlParameter.Add("repairremark", txtRepairremark.Text);
                    sqlParameter.Add("CreateUserID", MainWindow.CurrentUser);

                    Procedure pro1 = new Procedure();
                    pro1.Name = "xp_dvlMold_iMoldRepair";
                    pro1.OutputUseYN = "Y";
                    pro1.OutputName = "RepairID";
                    pro1.OutputLength = "10";

                    Prolist.Add(pro1);
                    ListParameter.Add(sqlParameter);

                    for (int i = 0; i < dgdMoldRepairSub.Items.Count; i++)
                    {
                        DataGridRow dgr = lib.GetRow(i, dgdMoldRepairSub);
                        var winRepairSub = dgr.Item as Win_dvl_MoldRepair_U_Sub_CodeView;

                        sqlParameter = new Dictionary<string, object>();
                        sqlParameter.Clear();
                        sqlParameter.Add("RepairID", "");
                        sqlParameter.Add("RepairSubSeq", i + 1);
                        sqlParameter.Add("McPartid", winRepairSub.McPartid);
                        sqlParameter.Add("customid", winRepairSub.customid);
                        sqlParameter.Add("partcnt", winRepairSub.partcnt);
                        sqlParameter.Add("partprice", winRepairSub.partprice);
                        sqlParameter.Add("reason", winRepairSub.reason);
                        sqlParameter.Add("partremark", winRepairSub.partremark);
                        sqlParameter.Add("CreateUserID", MainWindow.CurrentUser);

                        Procedure pro2 = new Procedure();
                        pro2.Name = "xp_dvlMold_iMoldRepairSub";
                        pro2.OutputUseYN = "N";
                        pro2.OutputName = "RepairID";
                        pro2.OutputLength = "10";

                        Prolist.Add(pro2);
                        ListParameter.Add(sqlParameter);
                    }

                    List<KeyValue> list_Result = new List<KeyValue>();
                    list_Result = DataStore.Instance.ExecuteAllProcedureOutputGetCS(Prolist, ListParameter);
                    string sGetRepairID = string.Empty;

                    if (list_Result[0].key.ToLower() == "success")
                    {
                        list_Result.RemoveAt(0);
                        for (int i = 0; i < list_Result.Count; i++)
                        {
                            KeyValue kv = list_Result[i];
                            if (kv.key == "RepairID")
                            {
                                sGetRepairID = kv.value;
                                dicCompare.Add("RepairID", sGetRepairID);
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
                catch (Exception ex)
                {
                    MessageBox.Show("오류 발생,오류 내용 : " + ex.ToString());
                }
                finally
                {
                    DataStore.Instance.CloseConnection();
                }
            }
            else { flag = false; }

            return flag;
        }

        //지금은 안쓰지만 일단 저장
        private bool InsertSubData(string strRepairID,int seq,Win_dvl_MoldRepair_U_Sub_CodeView winRepairSub)
        {
            bool flag = true;

            //Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
            //sqlParameter.Add("RepairID", strRepairID);
            //sqlParameter.Add("RepairSubSeq", seq);
            //sqlParameter.Add("McPartid", winRepairSub.McPartid);
            //sqlParameter.Add("customid", winRepairSub.customid);
            //sqlParameter.Add("partcnt", winRepairSub.partcnt);
            //sqlParameter.Add("partprice", winRepairSub.partprice);
            //sqlParameter.Add("reason", winRepairSub.reason);
            //sqlParameter.Add("partremark", winRepairSub.partremark);
            //sqlParameter.Add("CreateUserID", "");

            //string[] resultSub = DataStore.Instance.ExecuteProcedure("xp_dvlMold_iMoldRepairSub", sqlParameter, false);

            //if (!resultSub[0].Equals("success"))
            //{
            //    flag = false;
            //    //MessageBox.Show("실패 ㅠㅠ컥");
            //}

            return flag;
        }

        //수정
        private bool UpdateData()
        {
            bool flag = true;
            List<Procedure> Prolist = new List<Procedure>();
            List<Dictionary<string, object>> ListParameter = new List<Dictionary<string, object>>();

            try
            {
                Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                sqlParameter.Add("RepairID", txtRepairID.Text);
                sqlParameter.Add("repairdate", dtpRepairDate.SelectedDate.Value.ToString("yyyyMMdd"));
                sqlParameter.Add("RepairGubun", cboRepairGubun.SelectedValue.ToString());
                sqlParameter.Add("MoldID", txtMold.Tag.ToString());
                sqlParameter.Add("RepairCustom", txtRepairCustom.Text);
                sqlParameter.Add("repairremark", txtRepairremark.Text);
                sqlParameter.Add("UserID", MainWindow.CurrentUser);

                dicCompare.Add("RepairID", txtRepairID.Text);

                Procedure pro1 = new Procedure();
                pro1.Name = "xp_dvlMold_uMoldRepair";
                pro1.OutputUseYN = "N";
                pro1.OutputName = "InspectID";
                pro1.OutputLength = "10";

                Prolist.Add(pro1);
                ListParameter.Add(sqlParameter);

                for (int i = 0; i < dgdMoldRepairSub.Items.Count; i++)
                {
                    DataGridRow dgr = lib.GetRow(i, dgdMoldRepairSub);
                    var winRepairSub = dgr.Item as Win_dvl_MoldRepair_U_Sub_CodeView;

                    sqlParameter = new Dictionary<string, object>();
                    sqlParameter.Clear();
                    sqlParameter.Add("RepairID", txtRepairID.Text);
                    sqlParameter.Add("RepairSubSeq", i + 1);
                    sqlParameter.Add("McPartid", winRepairSub.McPartid);
                    sqlParameter.Add("customid", winRepairSub.customid);
                    sqlParameter.Add("partcnt", winRepairSub.partcnt);
                    sqlParameter.Add("partprice", winRepairSub.partprice);
                    sqlParameter.Add("reason", winRepairSub.reason);
                    sqlParameter.Add("partremark", winRepairSub.partremark);
                    sqlParameter.Add("CreateUserID", MainWindow.CurrentUser);

                    Procedure pro2 = new Procedure();
                    pro2.Name = "xp_dvlMold_iMoldRepairSub";
                    pro2.OutputUseYN = "N";
                    pro2.OutputName = "RepairID";
                    pro2.OutputLength = "10";

                    Prolist.Add(pro2);
                    ListParameter.Add(sqlParameter);
                }

                string[] Confirm = new string[2];
                Confirm = DataStore.Instance.ExecuteAllProcedureOutputNew(Prolist, ListParameter);
                if (Confirm[0] != "success")
                {
                    MessageBox.Show("[저장실패]\r\n" + Confirm[1].ToString());
                    flag = false;
                    //return false;
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

        //삭제
        private bool DeleteData(string strRepairID)
        {
            bool flag = true;

            try
            {
                Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                sqlParameter.Add("RepairID", strRepairID);

                string[] result = DataStore.Instance.ExecuteProcedure("xp_dvlMold_dMoldRepair", sqlParameter, true);                

                if (!result[0].Equals("success"))
                {
                    //MessageBox.Show("실패 ㅠㅠ");
                }
                else
                {
                    //MessageBox.Show("성공 *^^*");
                    flag = true;
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

        //추가, 수정 전 데이터 체크
        private bool CheckData()
        {
            bool flag = true;

            if (cboRepairGubun.SelectedIndex == -1)
            {
                MessageBox.Show("처리구분을 선택하지 않았습니다. 처리구분을 선택해주세요");
                flag = false;
            }

            if (txtMold.Tag == null)
            {
                MessageBox.Show("LotNo를 입력 후 enter_key를 누르거나 옆의 버튼을 눌러주시기 바랍니다.(그냥 누르면 선택화면을 볼 수 있습니다.)");
                flag = false;
            }

            if (dtpRepairDate.SelectedDate == null)
            {
                MessageBox.Show("수리일자가 선택되지 않았습니다. 수리일자를 선택해주세요");
                flag = false;
            }

            return flag;
        }

        #endregion

        #region 플러스 파인더 및 enter focus move 이벤트

        //처리구분
        private void cboRepairGubun_DropDownClosed(object sender, EventArgs e)
        {
            txtMold.Focus();
        }

        //중단에 LotNo 옆의 textBox에서 enter key 눌렀을때
        private void txtMold_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                MainWindow.pf.ReturnCode(txtMold,51,"");

                if (txtMold.Tag != null)
                {
                    GetArticleMoldData(txtMold.Tag.ToString());
                }

                dtpRepairDate.Focus();
            }
        }

        //중단에 LotNo 옆의 버튼 눌렀을때
        private void btnfPfMold_Click(object sender, RoutedEventArgs e)
        {
            MainWindow.pf.ReturnCode(txtMold, 51, "");

            if (txtMold.Tag != null)
            {
                GetArticleMoldData(txtMold.Tag.ToString());
            }

            dtpRepairDate.Focus();
        }

        //중단에서 플러스파인더 화면통해 선택시 다른 textBox 채우기
        private void GetArticleMoldData(string strMoldID)
        {
            try
            {
                Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                sqlParameter.Add("MoldID", strMoldID);

                DataSet ds = DataStore.Instance.ProcedureToDataSet("xp_dvlMold_sMoldOne", sqlParameter, false);

                if (ds != null && ds.Tables.Count > 0)
                {
                    DataTable dt = ds.Tables[0];

                    if (dt.Rows.Count > 0)
                    {
                        DataRowCollection drc = dt.Rows;

                        foreach (DataRow dr in drc)
                        {
                            var getArticleByMoldID = new GetAritcleByMoldID()
                            {
                                dvlYNName = dr["dvlYNName"].ToString(),
                                dvlYN = dr["dvlYN"].ToString(),
                                MoldQuality = dr["MoldQuality"].ToString(),
                                Weight = dr["Weight"].ToString(),
                                Spec = dr["Spec"].ToString(),
                                ProdCustomName = dr["ProdCustomName"].ToString()
                            };

                            txtSpec.Text = getArticleByMoldID.Spec;
                            txtWeight.Text = getArticleByMoldID.Weight;
                            txtMoldQuality.Text = getArticleByMoldID.MoldQuality;
                            txtProdCustomName.Text = getArticleByMoldID.ProdCustomName;
                            txtDvlYN.Text = getArticleByMoldID.dvlYNName;
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

        //수리일자
        private void dtpRepairDate_PreviewKeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                dtpRepairDate.IsDropDownOpen = true;
            }
        }

        //수리일자
        private void dtpRepairDate_CalendarClosed(object sender, RoutedEventArgs e)
        {
            txtRepairCustom.Focus();
        }

        //수리업체
        private void txtRepairCustom_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                txtRepairremark.Focus();
            }
        }

        #endregion

        #region 하단 그리드 Enter Event 및 del event

        //서브에서 추가버튼 클릭
        private void btnSubAdd_Click(object sender, RoutedEventArgs e)
        {
            SubPlus();
            dgdMoldRepairSub.Focus();
            //dgdMoldRepairSub.SelectedIndex = dgdMoldRepairSub.Items.Count - 1;
            dgdMoldRepairSub.CurrentCell = new DataGridCellInfo(dgdMoldRepairSub.Items[dgdMoldRepairSub.Items.Count - 1], dgdMoldRepairSub.Columns[0]);
        }

        private void SubPlus()
        {
            var WinMoldRepairSub = new Win_dvl_MoldRepair_U_Sub_CodeView()
            {
                InCustomName = "",
                MCPartName = "",
                partcnt = "",
                partprice = "",
                partremark = "",
                reason = "",
                customid = "",
                McPartid = "",
                RepairID = "",
                RepairSubSeq = "",
                flagCustom = false,
                flagMcPart = false,
                flagPartcnt = false,
                flagPartprice = false,
                flagPartremark = false,
                flagReason = false
            };
            dgdMoldRepairSub.Items.Add(WinMoldRepairSub);
        }

        //서브에서 삭제버튼 클릭
        private void btnSubDel_Click(object sender, RoutedEventArgs e)
        {
            SubRemove();
        }

        private void SubRemove()
        {
            if (dgdMoldRepairSub.Items.Count > 0)
            {
                if (dgdMoldRepairSub.CurrentItem != null)
                {
                    dgdMoldRepairSub.Items.Remove((dgdMoldRepairSub.CurrentItem as Win_dvl_MoldRepair_U_Sub_CodeView));
                }
                else
                {
                    dgdMoldRepairSub.Items.Remove((dgdMoldRepairSub.Items[dgdMoldRepairSub.Items.Count - 1]) as Win_dvl_MoldRepair_U_Sub_CodeView);
                }
                dgdMoldRepairSub.Refresh();
            }
        }

        //모든 셀
        private void DataGridCell_GotFocus(object sender, RoutedEventArgs e)
        {
            if (lblMsg.Visibility == Visibility.Visible)
            {
                DataGridCell cell = sender as DataGridCell;
                cell.IsEditing = true;
            }
        }

        //모든 셀
        private void DataGridCell_MouseUp(object sender, MouseButtonEventArgs e)
        {
            lib.DataGridINTextBoxFocusByMouseUP(sender, e);
        }

        //하단 DataGrid 부품명
        private void dgdtpeMCPartName_KeyDown(object sender, KeyEventArgs e)
        {
            if (lblMsg.Visibility == Visibility.Visible)
            {
                WinMoldRepairSub = dgdMoldRepairSub.CurrentItem as Win_dvl_MoldRepair_U_Sub_CodeView;
                int rowCount = dgdMoldRepairSub.Items.IndexOf(dgdMoldRepairSub.CurrentItem);

                if (e.Key == Key.Enter)
                {
                    e.Handled = true;
                    (sender as DataGridCell).IsEditing = false;
                    dgdMoldRepairSub.CurrentCell = new DataGridCellInfo(dgdMoldRepairSub.Items[rowCount], dgdMoldRepairSub.Columns[1]);
                }
                else if (e.Key == Key.Delete)
                {
                    SubRemove();

                    if (dgdMoldRepairSub.Items.Count > 0)
                    {
                        dgdMoldRepairSub.Focus();
                        //dgdMoldRepairSub.SelectedIndex = dgdMoldRepairSub.Items.Count - 1;
                        if (rowCount == 0)
                        {
                            dgdMoldRepairSub.CurrentCell = new DataGridCellInfo(dgdMoldRepairSub.Items[rowCount], dgdMoldRepairSub.Columns[0]);
                        }
                        else
                        {
                            dgdMoldRepairSub.CurrentCell = new DataGridCellInfo(dgdMoldRepairSub.Items[rowCount - 1], dgdMoldRepairSub.Columns[0]);
                        }                        
                    }
                }
            }
        }

        //하단 DataGrid 부품명
        private void dgdtxtMCPartName_KeyDown(object sender, KeyEventArgs e)
        {
            WinMoldRepairSub = dgdMoldRepairSub.CurrentItem as Win_dvl_MoldRepair_U_Sub_CodeView;
            int rowCount = dgdMoldRepairSub.Items.IndexOf(dgdMoldRepairSub.CurrentItem);
            if (e.Key == Key.Enter)
            {
                TextBox tb1 = sender as TextBox;
                MainWindow.pf.ReturnCode(tb1, 13, "");

                if (tb1.Tag != null)
                {
                    WinMoldRepairSub.MCPartName = tb1.Text;
                    WinMoldRepairSub.McPartid = tb1.Tag.ToString();
                }

                sender = tb1;
            }
        }

        //하단 DataGrid 부품명
        private void dgdtxtMCPartName_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            var WinRepairSub = dgdMoldRepairSub.CurrentItem as Win_dvl_MoldRepair_U_Sub_CodeView;
            int rowCount = dgdMoldRepairSub.Items.IndexOf(dgdMoldRepairSub.CurrentItem);

            TextBox tb1 = sender as TextBox;
            MainWindow.pf.ReturnCode(tb1, 13, "");

            if (tb1.Tag != null)
            {
                WinRepairSub.MCPartName = tb1.Text;
                WinRepairSub.McPartid = tb1.Tag.ToString();
            }

            sender = tb1;
        }

        //하단 DataGrid 부품명_enter key 없이 값 대입
        private void dgdtxtMCPartName_LostFocus(object sender, RoutedEventArgs e)
        {
            if (lblMsg.Visibility == Visibility.Visible)
            {
                WinMoldRepairSub = dgdMoldRepairSub.CurrentItem as Win_dvl_MoldRepair_U_Sub_CodeView;

                if (WinMoldRepairSub != null)
                {
                    TextBox tb1 = sender as TextBox;
                    MainWindow.pf.ReturnCode(tb1, 13, "");

                    if (tb1.Tag != null)
                    {
                        WinMoldRepairSub.MCPartName = tb1.Text;
                        WinMoldRepairSub.McPartid = tb1.Tag.ToString();
                    }

                    sender = tb1;
                }
            }
        }

        // 거래처(구입처)
        private void dgdtpeInCustomName_KeyDown(object sender, KeyEventArgs e)
        {
            if (lblMsg.Visibility == Visibility.Visible)
            {
                WinMoldRepairSub = dgdMoldRepairSub.CurrentItem as Win_dvl_MoldRepair_U_Sub_CodeView;
                int rowCount = dgdMoldRepairSub.Items.IndexOf(dgdMoldRepairSub.CurrentItem);
                if (e.Key == Key.Enter)
                {
                    e.Handled = true;
                    (sender as DataGridCell).IsEditing = false;
                    dgdMoldRepairSub.CurrentCell = new DataGridCellInfo(dgdMoldRepairSub.Items[rowCount], dgdMoldRepairSub.Columns[2]);
                }
                else if (e.Key == Key.Delete)
                {
                    SubRemove();

                    if (dgdMoldRepairSub.Items.Count > 0)
                    {
                        dgdMoldRepairSub.Focus();
                        //dgdMoldRepairSub.SelectedIndex = dgdMoldRepairSub.Items.Count - 1;
                        if (rowCount == 0)
                        {
                            dgdMoldRepairSub.CurrentCell = new DataGridCellInfo(dgdMoldRepairSub.Items[rowCount], dgdMoldRepairSub.Columns[0]);
                        }
                        else
                        {
                            dgdMoldRepairSub.CurrentCell = new DataGridCellInfo(dgdMoldRepairSub.Items[rowCount - 1], dgdMoldRepairSub.Columns[0]);
                        }
                    }
                }
            }
        }

        // 구입처
        private void dgdtxtInCustomName_KeyDown(object sender, KeyEventArgs e)
        {
            WinMoldRepairSub = dgdMoldRepairSub.CurrentItem as Win_dvl_MoldRepair_U_Sub_CodeView;
            int rowCount = dgdMoldRepairSub.Items.IndexOf(dgdMoldRepairSub.CurrentItem);
            if (e.Key == Key.Enter)
            {
                TextBox tb1 = sender as TextBox;
                MainWindow.pf.ReturnCode(tb1, 0, "");    //일단 0으로 ㅠㅠ

                //tb1.Text = "";
                //tb1.Tag = "";

                if (tb1.Tag != null)
                {
                    WinMoldRepairSub.InCustomName = tb1.Text;
                    WinMoldRepairSub.customid = tb1.Tag.ToString();
                }

                sender = tb1;
            }
        }

        // 거래처(구입처)_enter key 없이 값 대입
        private void dgdtxtInCustomName_LostFocus(object sender, RoutedEventArgs e)
        {
            if (lblMsg.Visibility == Visibility.Visible)
            {
                WinMoldRepairSub = dgdMoldRepairSub.CurrentItem as Win_dvl_MoldRepair_U_Sub_CodeView;

                if (WinMoldRepairSub != null)
                {
                    TextBox tb1 = sender as TextBox;

                    if (tb1.Tag != null)
                    {
                        WinMoldRepairSub.InCustomName = tb1.Text;
                        WinMoldRepairSub.customid = tb1.Tag.ToString();
                    }

                    sender = tb1;
                }
            }
        }

        // 구입처
        private void dgdtxtInCustomName_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            var WinRepairSub = dgdMoldRepairSub.CurrentItem as Win_dvl_MoldRepair_U_Sub_CodeView;
            int rowCount = dgdMoldRepairSub.Items.IndexOf(dgdMoldRepairSub.CurrentItem);

            TextBox tb1 = sender as TextBox;
            MainWindow.pf.ReturnCode(tb1, 0, "");    //일단 0으로 ㅠㅠ

            //tb1.Text = "";
            //tb1.Tag = "";

            if (tb1.Tag != null)
            {
                WinRepairSub.InCustomName = tb1.Text;
                WinRepairSub.customid = tb1.Tag.ToString();
            }

            sender = tb1;
        }

        //수량
        private void dgdtpepartcnt_KeyDown(object sender, KeyEventArgs e)
        {
            if (lblMsg.Visibility == Visibility.Visible)
            {
                WinMoldRepairSub = dgdMoldRepairSub.CurrentItem as Win_dvl_MoldRepair_U_Sub_CodeView;
                int rowCount = dgdMoldRepairSub.Items.IndexOf(dgdMoldRepairSub.CurrentItem);

                if (e.Key == Key.Enter)
                {
                    e.Handled = true;
                    (sender as DataGridCell).IsEditing = true;
                    dgdMoldRepairSub.CurrentCell = new DataGridCellInfo(dgdMoldRepairSub.Items[rowCount], dgdMoldRepairSub.Columns[3]);
                }
                else if (e.Key == Key.Delete)
                {
                    SubRemove();

                    if (dgdMoldRepairSub.Items.Count > 0)
                    {
                        dgdMoldRepairSub.Focus();
                        //dgdMoldRepairSub.SelectedIndex = dgdMoldRepairSub.Items.Count - 1;
                        if (rowCount == 0)
                        {
                            dgdMoldRepairSub.CurrentCell = new DataGridCellInfo(dgdMoldRepairSub.Items[rowCount], dgdMoldRepairSub.Columns[0]);
                        }
                        else
                        {
                            dgdMoldRepairSub.CurrentCell = new DataGridCellInfo(dgdMoldRepairSub.Items[rowCount - 1], dgdMoldRepairSub.Columns[0]);
                        }
                        
                    }
                }
            }
        }

        //수량
        private void dgdtxtpartcnt_LostFocus(object sender, RoutedEventArgs e)
        {
            if (lblMsg.Visibility == Visibility.Visible)
            {
                WinMoldRepairSub = dgdMoldRepairSub.CurrentItem as Win_dvl_MoldRepair_U_Sub_CodeView;

                if (WinMoldRepairSub != null)
                {
                    TextBox tb1 = sender as TextBox;
                    WinMoldRepairSub.partcnt = tb1.Text;
                    sender = tb1;
                }
            }
        }

        //수량(숫자만)
        private void dgdtxtpartcnt_PreviewTextInput(object sender, TextCompositionEventArgs e)
        {
            lib.CheckIsNumericOnly((TextBox)sender, e);
        }

        //수리비용
        private void dgdtpepartprice_KeyDown(object sender, KeyEventArgs e)
        {
            if (lblMsg.Visibility == Visibility.Visible)
            {
                WinMoldRepairSub = dgdMoldRepairSub.CurrentItem as Win_dvl_MoldRepair_U_Sub_CodeView;
                int rowCount = dgdMoldRepairSub.Items.IndexOf(dgdMoldRepairSub.CurrentItem);
                if (e.Key == Key.Enter)
                {
                    e.Handled = true;
                    (sender as DataGridCell).IsEditing = false;
                    dgdMoldRepairSub.CurrentCell = new DataGridCellInfo(dgdMoldRepairSub.Items[rowCount], dgdMoldRepairSub.Columns[4]);
                }
                else if (e.Key == Key.Delete)
                {
                    SubRemove();

                    if (dgdMoldRepairSub.Items.Count > 0)
                    {
                        dgdMoldRepairSub.Focus();
                        //dgdMoldRepairSub.SelectedIndex = dgdMoldRepairSub.Items.Count - 1;
                        if (rowCount == 0)
                        {
                            dgdMoldRepairSub.CurrentCell = new DataGridCellInfo(dgdMoldRepairSub.Items[rowCount], dgdMoldRepairSub.Columns[0]);
                        }
                        else
                        {
                            dgdMoldRepairSub.CurrentCell = new DataGridCellInfo(dgdMoldRepairSub.Items[rowCount - 1], dgdMoldRepairSub.Columns[0]);
                        }
                    }
                }
            }
        }

        //수리비용
        private void dgdtxtpartprice_LostFocus(object sender, RoutedEventArgs e)
        {
            if (lblMsg.Visibility == Visibility.Visible)
            {
                WinMoldRepairSub = dgdMoldRepairSub.CurrentItem as Win_dvl_MoldRepair_U_Sub_CodeView;

                if (WinMoldRepairSub != null)
                {
                    TextBox tb1 = sender as TextBox;
                    WinMoldRepairSub.partprice = tb1.Text;
                    sender = tb1;
                }
            }
        }

        //수리비용
        private void dgdtxtpartprice_PreviewTextInput(object sender, TextCompositionEventArgs e)
        {
            lib.CheckIsNumericOnly((TextBox)sender, e);
        }

        //사유
        private void dgdtpereason_KeyDown(object sender, KeyEventArgs e)
        {
            if (lblMsg.Visibility == Visibility.Visible)
            {
                WinMoldRepairSub = dgdMoldRepairSub.CurrentItem as Win_dvl_MoldRepair_U_Sub_CodeView;
                int rowCount = dgdMoldRepairSub.Items.IndexOf(dgdMoldRepairSub.CurrentItem);
                if (e.Key == Key.Enter)
                {
                    e.Handled = true;
                    (sender as DataGridCell).IsEditing = false;
                    dgdMoldRepairSub.CurrentCell = new DataGridCellInfo(dgdMoldRepairSub.Items[rowCount], dgdMoldRepairSub.Columns[5]);
                }
                else if (e.Key == Key.Delete)
                {
                    SubRemove();

                    if (dgdMoldRepairSub.Items.Count > 0)
                    {
                        dgdMoldRepairSub.Focus();
                        //dgdMoldRepairSub.SelectedIndex = dgdMoldRepairSub.Items.Count - 1;
                        if (rowCount == 0)
                        {
                            dgdMoldRepairSub.CurrentCell = new DataGridCellInfo(dgdMoldRepairSub.Items[rowCount], dgdMoldRepairSub.Columns[0]);
                        }
                        else
                        {
                            dgdMoldRepairSub.CurrentCell = new DataGridCellInfo(dgdMoldRepairSub.Items[rowCount - 1], dgdMoldRepairSub.Columns[0]);
                        }
                    }
                }
            }
        }

        //사유
        private void dgdtxtreason_LostFocus(object sender, RoutedEventArgs e)
        {
            if (lblMsg.Visibility == Visibility.Visible)
            {
                WinMoldRepairSub = dgdMoldRepairSub.CurrentItem as Win_dvl_MoldRepair_U_Sub_CodeView;

                if (WinMoldRepairSub != null)
                {
                    TextBox tb1 = sender as TextBox;
                    WinMoldRepairSub.reason = tb1.Text;
                    sender = tb1;
                }
            }
        }

        //비고사항
        private void dgdtpepartremark_KeyDown(object sender, KeyEventArgs e)
        {
            if (lblMsg.Visibility == Visibility.Visible)
            {
                WinMoldRepairSub = dgdMoldRepairSub.CurrentItem as Win_dvl_MoldRepair_U_Sub_CodeView;
                int rowCount = dgdMoldRepairSub.Items.IndexOf(dgdMoldRepairSub.CurrentItem);
                if (e.Key == Key.Enter)
                {
                    e.Handled = true;
                    (sender as DataGridCell).IsEditing = false;

                    if (dgdMoldRepairSub.Items.Count - 1 == rowCount)
                    {
                        SubPlus();
                    }
                    //dgdMoldRepairSub.SelectedIndex = dgdMoldRepairSub.Items.Count - 1;
                    dgdMoldRepairSub.CurrentCell = new DataGridCellInfo(dgdMoldRepairSub.Items[rowCount + 1], dgdMoldRepairSub.Columns[0]);                    
                }
                else if (e.Key == Key.Delete)
                {
                    SubRemove();

                    if (dgdMoldRepairSub.Items.Count > 0)
                    {
                        dgdMoldRepairSub.Focus();
                        //dgdMoldRepairSub.SelectedIndex = dgdMoldRepairSub.Items.Count - 1;
                        if (rowCount == 0)
                        {
                            dgdMoldRepairSub.CurrentCell = new DataGridCellInfo(dgdMoldRepairSub.Items[rowCount], dgdMoldRepairSub.Columns[0]);
                        }
                        else
                        {
                            dgdMoldRepairSub.CurrentCell = new DataGridCellInfo(dgdMoldRepairSub.Items[rowCount - 1], dgdMoldRepairSub.Columns[0]);
                        }
                    }
                }
            }
        }

        //비고사항
        private void dgdtxtpartremark_LostFocus(object sender, RoutedEventArgs e)
        {
            if (lblMsg.Visibility == Visibility.Visible)
            {
                WinMoldRepairSub = dgdMoldRepairSub.CurrentItem as Win_dvl_MoldRepair_U_Sub_CodeView;

                if (WinMoldRepairSub != null)
                {
                    TextBox tb1 = sender as TextBox;
                    WinMoldRepairSub.partremark = tb1.Text;
                    sender = tb1;
                }
            }
        }

        //서브 그리드 인 포커스
        private void DatagridIn_TextFocus(object sender, KeyEventArgs e)
        {
            lib.DataGridINTextBoxFocus(sender, e);
        }

        //(서브)모든 셀에 삭제 적용하기 위해(TextBox 안해서는 keydown이 안먹힌다.)
        //PreviewKewDown 적용, 셀의 delete 이벤트를 탈수 있게 포커스 바꿔줌
        private void DataGridIn_PreviewKeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Delete)
            {
                DataGridCell cell = lib.GetParent<DataGridCell>(sender as TextBox);
                cell.Focus();
            }
        }


        #endregion

        
    }
}

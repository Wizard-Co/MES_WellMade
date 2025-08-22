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
using WizMes_WellMade.PopUp;
using WPF.MDI;

namespace WizMes_WellMade
{
    /// <summary>
    /// Win_prd_Repair_U.xaml에 대한 상호 작용 논리
    /// </summary>
    public partial class Win_prd_Repair_U : UserControl
    {
        #region 변수 선언 및 로드

        int numRowCount = 0;
        string strFlag = string.Empty;
        Win_prd_Repair_U_Sub_CodeView WinMCRepairSub = new Win_prd_Repair_U_Sub_CodeView();
        Dictionary<string, object> dicCompare = new Dictionary<string, object>();
        List<string> lstCompareValue = new List<string>();
        Lib lib = new Lib();
        public Win_prd_Repair_U()
        {
            InitializeComponent();
        }

        private void UserControl_Loaded(object sender, RoutedEventArgs e)
        {
            Lib.Instance.UiLoading(sender);
            dtpSDate.SelectedDate = DateTime.Today;
            dtpEDate.SelectedDate = DateTime.Today;
            chkRepairDaySrh.IsChecked = true;

            SetCombo();
        }

        private void SetCombo()
        {
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

        #region 상단 중간 이벤트

        //수리일자 라벨
        private void lblRepairDaySrh_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            if (chkRepairDaySrh.IsChecked == true) { chkRepairDaySrh.IsChecked = false; }
            else { chkRepairDaySrh.IsChecked = true; }
        }

        //수리일자 체크박스
        private void chkRepairDaySrh_Checked(object sender, RoutedEventArgs e)
        {
            dtpSDate.IsEnabled = true;
            dtpEDate.IsEnabled = true;
        }

        //수리일자 체크박스
        private void chkRepairDaySrh_Unchecked(object sender, RoutedEventArgs e)
        {
            dtpSDate.IsEnabled = false;
            dtpEDate.IsEnabled = false;
        }

        //금월
        private void btnThisMonth_Click(object sender, RoutedEventArgs e)
        {
            dtpSDate.SelectedDate = Lib.Instance.BringThisMonthDatetimeList()[0];
            dtpEDate.SelectedDate = Lib.Instance.BringThisMonthDatetimeList()[1];
        }

        //금일
        private void btnToday_Click(object sender, RoutedEventArgs e)
        {
            dtpSDate.SelectedDate = DateTime.Today;
            dtpEDate.SelectedDate = DateTime.Today;
        }

        //기계명 라벨
        private void lblMachine_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            if (chkMachine.IsChecked == true) { chkMachine.IsChecked = false; }
            else { chkMachine.IsChecked = true; }
        }

        //기계명 체크박스
        private void chkMachine_Checked(object sender, RoutedEventArgs e)
        {
            txtMachine.IsEnabled = true;
            btnPfMachine.IsEnabled = true;
            txtMachine.Focus();
        }

        //기계명 체크박스
        private void chkMachine_Unchecked(object sender, RoutedEventArgs e)
        {
            txtMachine.IsEnabled = false;
            btnPfMachine.IsEnabled = false;
        }

        //기계명 텍스트박스
        private void txtMachine_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                MainWindow.pf.ReturnCode(txtMachine, (int)Defind_CodeFind.DCF_MC, "");
            }
        }

        //기계명 플러스파인더
        private void btnPfMachine_Click(object sender, RoutedEventArgs e)
        {
            MainWindow.pf.ReturnCode(txtMachine, (int)Defind_CodeFind.DCF_MC, "");
        }

        #endregion

        #region 상단 우측 버튼 이벤트

        //추가,수정 시 동작 모음
        private void ControlVisibleAndEnable_AU()
        {
            Lib.Instance.UiButtonEnableChange_SCControl(this);
            //dgdMCRepair.IsEnabled = false;
            dgdMCRepair.IsHitTestVisible = false;
            bdrRepair.IsEnabled = true;
            btnSubPlus.IsEnabled = true;
            btnSubDel.IsEnabled = true;
        }

        //저장,취소 시 동작 모음
        private void ControlVisibleAndEnable_SC()
        {
            Lib.Instance.UiButtonEnableChange_IUControl(this);
            //dgdMCRepair.IsEnabled = true;
            dgdMCRepair.IsHitTestVisible = true;
            bdrRepair.IsEnabled = false;
            btnSubPlus.IsEnabled = false;
            btnSubDel.IsEnabled = false;
        }

        //추가
        private void btnAdd_Click(object sender, RoutedEventArgs e)
        {
            if (dgdMCRepair.Items.Count > 0 && dgdMCRepair.SelectedItem != null)
            {
                numRowCount = dgdMCRepair.SelectedIndex;
            }

            ControlVisibleAndEnable_AU();
            strFlag = "I";
            tbkMsg.Text = "자료 입력(추가) 중";

            this.DataContext = null;

            if (dgdMCRepair_Sub.Items.Count > 0)
            {
                dgdMCRepair_Sub.Items.Clear();
            }

            dtprepairdate.SelectedDate = DateTime.Today;
            txtRepairTime.Text = DateTime.Now.ToString("HH:mm");

            cboRepairGubun.Focus();
            cboRepairGubun.IsDropDownOpen = true;
        }

        //수정
        private void btnUpdate_Click(object sender, RoutedEventArgs e)
        {
            if (dgdMCRepair.SelectedItem == null)
            {
                MessageBox.Show("수정할 자료를 선택하고 눌러주십시오.");
            }
            else
            {
                numRowCount = dgdMCRepair.SelectedIndex;
                ControlVisibleAndEnable_AU();
                tbkMsg.Text = "자료 입력(수정) 중";
                strFlag = "U";

                cboRepairGubun.Focus();
                cboRepairGubun.IsDropDownOpen = true;
            }
        }

        //삭제
        private void btnDelete_Click(object sender, RoutedEventArgs e)
        {
            var winMCRepair = dgdMCRepair.SelectedItem as Win_prd_Repair_U_CodeView;

            if (winMCRepair == null)
            {
                MessageBox.Show("삭제할 데이터가 지정되지 않았습니다. 삭제데이터를 지정하고 눌러주세요");
                return;
            }
            else
            {
                if (dgdMCRepair.SelectedIndex == dgdMCRepair.Items.Count - 1)
                {
                    numRowCount = dgdMCRepair.SelectedIndex - 1;
                }
                else
                {
                    numRowCount = dgdMCRepair.SelectedIndex;
                }

                if (MessageBox.Show("선택하신 항목을 삭제하시겠습니까?", "삭제 전 확인", MessageBoxButton.YesNo) == MessageBoxResult.Yes)
                {
                    if (DeleteData(winMCRepair.RepairID))
                    {
                        re_Search(numRowCount);
                    }
                }
            }
        }

        //닫기
        private void btnClose_Click(object sender, RoutedEventArgs e)
        {
            int i = 0;
            foreach (MenuViewModel mvm in MainWindow.mMenulist)
            {
                if (mvm.subProgramID.ToString().Contains("MDI"))
                {
                    if (this.ToString().Equals((mvm.subProgramID as MdiChild).Content.ToString()))
                    {
                        (MainWindow.mMenulist[i].subProgramID as MdiChild).Close();
                        break;
                    }
                }
                i++;
            }
        }

        //조회
        private void btnSearch_Click(object sender, RoutedEventArgs e)
        {
            //검색버튼 비활성화
            btnSearch.IsEnabled = false;

            Dispatcher.BeginInvoke(new Action(() =>

            {
                try
                {
                    int rowNum = 0;
                    using (Loading lw = new Loading(FillGrid))
                    {
                        lw.ShowDialog();
                        if (dgdMCRepair.Items.Count <= 0)
                        {
                            MessageBox.Show("조회된 내용이 없습니다.");
                        }
                        else
                        {
                            dgdMCRepair.SelectedIndex = rowNum;
                        }

                        btnSearch.IsEnabled = true;
                    }
                }
                catch (Exception ee)
                {
                    MessageBox.Show("예외처리 - " + ee.ToString());
                }

            }), System.Windows.Threading.DispatcherPriority.Background);

        }

        //저장
        private void btnSave_Click(object sender, RoutedEventArgs e)
        {


            if (strFlag.Equals("I"))
            {
                if (SaveData("", strFlag))
                {
                    ControlVisibleAndEnable_SC();
                    numRowCount = 0;
                    re_Search(numRowCount);
                }
            }
            else
            {
                if (SaveData(txtRepairID.Text, strFlag))
                {
                    ControlVisibleAndEnable_SC();
                    re_Search(numRowCount);
                }
            }
        }

        //취소
        private void btnCancel_Click(object sender, RoutedEventArgs e)
        {
            ControlVisibleAndEnable_SC();
            re_Search(numRowCount);
        }

        //입력 데이터 클리어
        private void InputClear()
        {
            foreach (Control child in grdInput.Children)
            {
                if (child.GetType() == typeof(TextBox))
                    ((TextBox)child).Clear();
                else if (child.GetType() == typeof(ComboBox))
                    ((ComboBox)child).SelectedIndex = -1;
                else if (child.GetType() == typeof(CheckBox))
                    ((CheckBox)child).IsChecked = false;
            }

            if (dgdMCRepair_Sub.Items.Count > 0)
            {
                dgdMCRepair_Sub.Items.Clear();
            }
        }

        //엑셀
        private void btnExcel_Click(object sender, RoutedEventArgs e)
        {
            DataTable dt = null;
            string Name = string.Empty;

            string[] lst = new string[4];
            lst[0] = "설비 수리 메인";
            lst[1] = "설비 수리_부품내역";
            lst[2] = dgdMCRepair.Name;
            lst[3] = dgdMCRepair_Sub.Name;

            ExportExcelxaml ExpExc = new ExportExcelxaml(lst);

            ExpExc.ShowDialog();

            if (ExpExc.DialogResult.HasValue)
            {
                if (ExpExc.choice.Equals(dgdMCRepair.Name))
                {
                    if (ExpExc.Check.Equals("Y"))
                        dt = Lib.Instance.DataGridToDTinHidden(dgdMCRepair);
                    else
                        dt = Lib.Instance.DataGirdToDataTable(dgdMCRepair);

                    Name = dgdMCRepair.Name;

                    if (Lib.Instance.GenerateExcel(dt, Name))
                        Lib.Instance.excel.Visible = true;
                    else
                        return;
                }
                else if (ExpExc.choice.Equals(dgdMCRepair_Sub.Name))
                {
                    if (ExpExc.Check.Equals("Y"))
                        dt = Lib.Instance.DataGridToDTinHidden(dgdMCRepair_Sub);
                    else
                        dt = Lib.Instance.DataGirdToDataTable(dgdMCRepair_Sub);

                    Name = dgdMCRepair_Sub.Name;

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

        //수정,추가,삭제 후 조회 등
        private void re_Search(int index)
        {
            if (dgdMCRepair.Items.Count > 0)
            {
                dgdMCRepair.Items.Clear();
            }

            FillGrid();

            if (dgdMCRepair.Items.Count > 0)
            {
                if (lstCompareValue.Count > 0)
                {
                    dgdMCRepair.SelectedIndex = Lib.Instance.reTrunIndex(dgdMCRepair, lstCompareValue[0]);
                }
                else
                {
                    dgdMCRepair.SelectedIndex = index; ;
                }
            }
            else
            {
                InputClear();

                MessageBox.Show("조회된 데이터가 없습니다.");
                return;
            }

            dicCompare.Clear();
            lstCompareValue.Clear();
        }

        #endregion

        #region CRUD

        //조회
        private void FillGrid()
        {
            if (dgdMCRepair.Items.Count > 0)
            {
                dgdMCRepair.Items.Clear();
            }
            try
            {
                Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                sqlParameter.Add("nRepairDate", chkRepairDaySrh.IsChecked == true ? 1 : 0);
                sqlParameter.Add("RepairStartDate", chkRepairDaySrh.IsChecked == true ? dtpSDate.SelectedDate.Value.ToString("yyyyMMdd") : "");
                sqlParameter.Add("RepairEndDate", chkRepairDaySrh.IsChecked == true ? dtpEDate.SelectedDate.Value.ToString("yyyyMMdd") : "");
                sqlParameter.Add("nMachine", chkMachine.IsChecked == true ? (txtMachine.Tag != null ? 1 : 2) : 0);
                sqlParameter.Add("sMachine", chkMachine.IsChecked == true ? (txtMachine.Tag != null ? txtMachine.Tag.ToString() : txtMachine.Text) : "");
                //2021-08-03 부품명 조건 추가
                sqlParameter.Add("nArticle", chkArticleSearch.IsChecked == true ? (txtArticleSearch.Tag != null ? 1 : 2) : 0);
                sqlParameter.Add("sArticle", chkArticleSearch.IsChecked == true ? (txtArticleSearch.Tag != null ? txtArticleSearch.Tag.ToString() : txtArticleSearch.Text) : "");
                DataSet ds = DataStore.Instance.ProcedureToDataSet("xp_mcRepair_sRepair", sqlParameter, false);

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
                            var WinMCRepair = new Win_prd_Repair_U_CodeView()
                            {
                                Num = i,
                                RepairID = dr["RepairID"].ToString(),
                                repairdate = dr["repairdate"].ToString(),
                                repairTime = dr["repairTime"].ToString().Substring(0, 2) + ":" + dr["repairTime"].ToString().Substring(2, 2),
                                RepairGubun = dr["RepairGubun"].ToString(),

                                mcid = dr["mcid"].ToString(),
                                mcname = dr["mcname"].ToString(),
                                managerid = dr["managerid"].ToString(),
                                customid = dr["customid"].ToString(),
                                customname = dr["customname"].ToString(),
                                buycustomid = dr["buycustomid"].ToString(),
                                BuyCustomName = dr["BuyCustomName"].ToString(),
                                personid = dr["personid"].ToString(),
                                personname = dr["personname"].ToString(),
                                RepairRemark = dr["RepairRemark"].ToString(),
                                //price = Convert.ToDouble(dr["price"])
                            };

                            if (WinMCRepair.repairdate != null && !WinMCRepair.repairdate.Equals(""))
                            {
                                WinMCRepair.repairdate_CV = Lib.Instance.StrDateTimeBar(WinMCRepair.repairdate);
                            }

                            if (WinMCRepair.RepairGubun.Equals("1"))
                            {
                                WinMCRepair.RepairGubun_CV = "수리";
                            }
                            else if (WinMCRepair.RepairGubun.Equals("2"))
                            {
                                WinMCRepair.RepairGubun_CV = "교체";
                            }

                            if (dicCompare.Count > 0)
                            {
                                if (WinMCRepair.RepairID.Equals(dicCompare["RepairID"].ToString()))
                                {
                                    lstCompareValue.Add(WinMCRepair.ToString());
                                }
                            }

                            dgdMCRepair.Items.Add(WinMCRepair);
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

        //메인 그리드 one row select change 이벤트
        private void dgdMCRepair_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            Win_prd_Repair_U_CodeView winMCRepair = dgdMCRepair.SelectedItem as Win_prd_Repair_U_CodeView;

            if (winMCRepair != null)
            {
                if (dgdMCRepair_Sub.Items.Count > 0)
                {
                    dgdMCRepair_Sub.Items.Clear();
                }

                if (!winMCRepair.RepairID.Replace(" ", "").Equals(""))
                {
                    FillGridSub(winMCRepair.RepairID);
                }

                this.DataContext = winMCRepair;
            }
        }

        //서브 그리드 조회
        private void FillGridSub(string strRepairID)
        {
            try
            {
                Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                sqlParameter.Add("sRepairID", strRepairID);

                DataSet ds = DataStore.Instance.ProcedureToDataSet("xp_mcRepairSub_sRepairSub", sqlParameter, false);

                if (ds != null && ds.Tables.Count > 0)
                {
                    DataTable dt = ds.Tables[0];

                    if (dt.Rows.Count > 0)
                    {
                        DataRowCollection drc = dt.Rows;

                        foreach (DataRow dr in drc)
                        {
                            var WinMCRepairSub = new Win_prd_Repair_U_Sub_CodeView()
                            {
                                RepairID = dr["RepairID"].ToString(),
                                repairSubseq = dr["repairSubseq"].ToString(),
                                MCPartID = dr["MCPartID"].ToString(),
                                MCPartName = dr["MCPartName"].ToString(),
                                reason = dr["reason"].ToString(),
                                partcnt = dr["partcnt"].ToString(),
                                partprice = dr["partprice"].ToString(),
                                partremark = dr["partremark"].ToString(),
                                CustomID = dr["CustomID"].ToString(),
                                customname = dr["customname"].ToString(),
                                flagcustomname = false,
                                flagMCPartName = false,
                                flagpartcnt = false,
                                flagpartprice = false,
                                flagpartremark = false,
                                flagreason = false
                            };

                            WinMCRepairSub.partcnt = Lib.Instance.returnNumStringZero(WinMCRepairSub.partcnt);
                            WinMCRepairSub.partprice = Lib.Instance.returnNumStringZero(WinMCRepairSub.partprice);
                            dgdMCRepair_Sub.Items.Add(WinMCRepairSub);
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

        //삭제
        private bool DeleteData(string strRepairID)
        {
            bool flag = true;

            try
            {
                Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                sqlParameter.Add("RepairID", strRepairID);

                string[] result = DataStore.Instance.ExecuteProcedure("xp_mcRepair_dRepair", sqlParameter, true);

                if (!result[0].Equals("success"))
                {
                    MessageBox.Show("데이터 삭제를 실패하였습니다.");
                }
                else
                {
                    //MessageBox.Show("성공 *^^*");
                    flag = true;
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

        //추가,수정
        private bool SaveData(string strRepairID, string strflag)
        {
            bool flag = true;
            List<Procedure> Prolist = new List<Procedure>();
            List<Dictionary<string, object>> ListParameter = new List<Dictionary<string, object>>();

            string repairID = string.Empty;

            if (CheckData())
            {
                try
                {

                    Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                    sqlParameter.Add("RepairID", strRepairID);
                    sqlParameter.Add("repairdate", dtprepairdate.SelectedDate.Value.ToString("yyyyMMdd"));
                    sqlParameter.Add("RepairTime", txtRepairTime.Text.Replace(":", ""));
                    sqlParameter.Add("RepairGubun", cboRepairGubun.SelectedValue != null ? cboRepairGubun.SelectedValue : 0);
                    sqlParameter.Add("mcid", txtmcname.Tag.ToString());

                    sqlParameter.Add("managerid", txtmanagerid.Text);
                    sqlParameter.Add("customid", txtcustomname.Tag.ToString());
                    sqlParameter.Add("buycustomid", txtBuyCustomName.Tag.ToString());
                    sqlParameter.Add("personid", txtpersonname.Tag.ToString().Replace(" ", ""));
                    sqlParameter.Add("personname", txtpersonname.Text);

                    sqlParameter.Add("repairremark", txtRepairRemark.Text);

                    //sqlParameter.Add("price", !txtPrice.Text.Equals("") ? Convert.ToDouble(txtPrice.Text) : 0);

                    if (strflag.Equals("I"))
                    {
                        sqlParameter.Add("CreateUserID", MainWindow.CurrentUser);   //추가,수정에서 유일하게 다른 파라미터

                        Procedure pro1 = new Procedure();
                        pro1.Name = "xp_mcRepair_iRepair";
                        pro1.OutputUseYN = "Y";
                        pro1.OutputName = "RepairID";
                        pro1.OutputLength = "10";

                        Prolist.Add(pro1);
                        ListParameter.Add(sqlParameter);

                        for (int i = 0; i < dgdMCRepair_Sub.Items.Count; i++)
                        {
                            //2021-05-19 설비부품 재고량을 알기 위해 추가
                            DataGridRow dgr = Lib.Instance.GetRow(i, dgdMCRepair_Sub);
                            var WinSub = dgr.Item as Win_prd_Repair_U_Sub_CodeView;

                            //Dictionary<string, object> sqlParameter2 = new Dictionary<string, object>();
                            //sqlParameter2.Add("MCPartID", WinSub.MCPartID);
                            //sqlParameter2.Add("PartCnt", int.Parse(WinSub.partcnt.Replace(",", ""))); 
                            //DataSet ds2 = DataStore.Instance.ProcedureToDataSet("xp_mcRepair_StuffinOut", sqlParameter2, false);

                            //if (ds2 != null && ds2.Tables.Count > 0)
                            //{
                            //    DataTable dt2 = ds2.Tables[0];
                            //    DataRow dt3 = dt2.Rows[0];
                            //    if (dt2.Rows.Count > 0 && !dt3["Msg"].ToString().ToUpper().Equals(""))
                            //    {
                            //        MessageBox.Show("[저장실패]\r\n" + dt3["Msg"].ToString());
                            //        flag = false;
                            //        return false;
                            //    }
                            //}
                            sqlParameter = new Dictionary<string, object>();
                            sqlParameter.Add("RepairID", strRepairID);
                            sqlParameter.Add("repairsubseq", i + 1);
                            sqlParameter.Add("MCPartID", WinSub.MCPartID);
                            sqlParameter.Add("customid", WinSub.CustomID);
                            sqlParameter.Add("partcnt", int.Parse(WinSub.partcnt.Replace(",", "")));
                            sqlParameter.Add("partprice", int.Parse(WinSub.partprice.Replace(",", "")));
                            sqlParameter.Add("reason", WinSub.reason);
                            sqlParameter.Add("partremark", WinSub.partremark);
                            sqlParameter.Add("CreateUserID", MainWindow.CurrentUser);

                            Procedure pro2 = new Procedure();
                            pro2.Name = "xp_mcRepairSub_iRepairSub";
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
                    else    //
                    {
                        sqlParameter.Add("UserID", MainWindow.CurrentUser);     //추가,수정에서 유일하게 다른 파라미터

                        dicCompare.Add("RepairID", strRepairID);

                        Procedure pro1 = new Procedure();
                        pro1.Name = "xp_mcRepair_uRepair";
                        pro1.OutputUseYN = "N";
                        pro1.OutputName = "RepairID";
                        pro1.OutputLength = "10";

                        Prolist.Add(pro1);
                        ListParameter.Add(sqlParameter);

                        for (int i = 0; i < dgdMCRepair_Sub.Items.Count; i++)
                        {
                            DataGridRow dgr = Lib.Instance.GetRow(i, dgdMCRepair_Sub);
                            var WinSub = dgr.Item as Win_prd_Repair_U_Sub_CodeView;

                            sqlParameter = new Dictionary<string, object>();
                            sqlParameter.Add("RepairID", strRepairID);
                            sqlParameter.Add("repairsubseq", i + 1);
                            sqlParameter.Add("MCPartID", WinSub.MCPartID);
                            sqlParameter.Add("customid", WinSub.CustomID);
                            sqlParameter.Add("partcnt", int.Parse(WinSub.partcnt.Replace(",", "")));
                            sqlParameter.Add("partprice", int.Parse(WinSub.partprice.Replace(",", "")));
                            sqlParameter.Add("reason", WinSub.reason);
                            sqlParameter.Add("partremark", WinSub.partremark);
                            sqlParameter.Add("CreateUserID", MainWindow.CurrentUser);

                            Procedure pro2 = new Procedure();
                            pro2.Name = "xp_mcRepairSub_iRepairSub";
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
            else { flag = false; }

            return flag;
        }

        //추가,수정 시 서브 데이터 추가
        private bool AddSub(Win_prd_Repair_U_Sub_CodeView WinSub, int seq)
        {
            bool flag = true;

            try
            {
                Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                sqlParameter.Add("RepairID", WinSub.RepairID);
                sqlParameter.Add("repairsubseq", seq);
                sqlParameter.Add("MCPartID", WinSub.MCPartID);
                sqlParameter.Add("customid", WinSub.CustomID);
                sqlParameter.Add("partcnt", int.Parse(WinSub.partcnt.Replace(",", "")));
                sqlParameter.Add("partprice", int.Parse(WinSub.partprice.Replace(",", "")));
                sqlParameter.Add("reason", WinSub.reason);
                sqlParameter.Add("partremark", WinSub.partremark);
                sqlParameter.Add("CreateUserID", MainWindow.CurrentUser);

                string[] resultSub = DataStore.Instance.ExecuteProcedure("xp_mcRepairSub_iRepairSub", sqlParameter, true);

                if (!resultSub[0].Equals("success"))
                {
                    flag = false;
                    MessageBox.Show("실패");
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

        //자료 추가,수정 시 확인
        private bool CheckData()
        {
            bool flag = true;

            var WinSub = dgdMCRepair_Sub.CurrentItem as Win_prd_Repair_U_Sub_CodeView;

            if (txtmcname.Tag == null || txtmcname.Tag.Equals(""))
            {
                MessageBox.Show("기계명'을 입력하지않았습니다. 확인하여 주십시오.");
                flag = false;
                return flag;
            }

            if (txtcustomname.Tag == null || txtcustomname.Tag.ToString().Equals(""))
            {
                MessageBox.Show("구입처'를 입력하지않았습니다. 확인하여 주십시오.");
                flag = false;
                return flag;
            }

            if (dgdMCRepair_Sub.Items.Count > 0)
            {
                for (int i = 0; i < dgdMCRepair_Sub.Items.Count; i++) //2021-05-13 부품수리 내역에서 입력값이 없을 경우 메세지 띄우기
                {
                    DataGridRow dgr = Lib.Instance.GetRow(i, dgdMCRepair_Sub);
                    var WinSubGrid = dgr.Item as Win_prd_Repair_U_Sub_CodeView;

                    if (WinSubGrid.MCPartID == null || WinSubGrid.MCPartID == "")
                    {
                        MessageBox.Show("예비품을 입력하지 않았습니다.");
                        flag = false;
                        return flag;
                    }
                    if (WinSubGrid.CustomID == null || WinSubGrid.CustomID == "")
                    {
                        MessageBox.Show("구입처를 입력하지 않았습니다.");
                        flag = false;
                        return flag;
                    }
                    if (WinSubGrid.partcnt == null || WinSubGrid.partcnt == "")
                    {
                        MessageBox.Show("수량을 입력하지 않았습니다.");
                        flag = false;
                        return flag;
                    }
                    if (WinSubGrid.partprice == null || WinSubGrid.partprice == "")
                    {
                        MessageBox.Show("수리비용을 입력하지 않았습니다.");
                        flag = false;
                        return flag;
                    }
                    //if (WinSubGrid.reason == null || WinSubGrid.reason == "")
                    //{
                    //    MessageBox.Show("사유를 입력하지 않았습니다.");
                    //    flag = false;
                    //    return flag;
                    //}
                }


            }
            //else
            //{
            //    MessageBox.Show("부품수리 내역을 입력하지 않았습니다. 부품추가하여 부품수리내역을 입력하여 주세요.");
            //    flag = false;
            //    return flag;
            //}

            return flag;
        }

        #endregion

        #region 하단 그리드 관련 이벤트

        //하단 그리드 추가 버튼 클릭
        private void btnSubPlus_Click(object sender, RoutedEventArgs e)
        {
            SubPlus();
            int colCount = dgdMCRepair_Sub.Columns.IndexOf(dgdtpeMCPartName);
            dgdMCRepair_Sub.Focus();
            //dgdMCRepair_Sub.SelectedIndex = dgdMCRepair_Sub.Items.Count - 1;
            dgdMCRepair_Sub.CurrentCell = new DataGridCellInfo(dgdMCRepair_Sub.Items[dgdMCRepair_Sub.Items.Count - 1], dgdMCRepair_Sub.Columns[colCount]);
        }

        //하단 그리드 실추가
        private void SubPlus()
        {
            var WinMCRepairSub = new Win_prd_Repair_U_Sub_CodeView()
            {
                RepairID = "",
                repairSubseq = "",
                MCPartID = "",
                MCPartName = "",
                reason = "",
                partcnt = "",
                partprice = "0",
                partremark = "",
                CustomID = "",
                customname = "",
                flagcustomname = false,
                flagMCPartName = false,
                flagpartcnt = false,
                flagpartprice = false,
                flagpartremark = false,
                flagreason = false
            };
            dgdMCRepair_Sub.Items.Add(WinMCRepairSub);
        }

        //하단 그리드 삭제 버튼 클릭
        private void btnSubDel_Click(object sender, RoutedEventArgs e)
        {
            SubRemove();
        }

        //하단 그리드 실삭제
        private void SubRemove()
        {
            if (dgdMCRepair_Sub.Items.Count > 0)
            {
                if (dgdMCRepair_Sub.CurrentItem != null)
                {
                    dgdMCRepair_Sub.Items.Remove((dgdMCRepair_Sub.CurrentItem as Win_prd_Repair_U_Sub_CodeView));
                }
                else
                {
                    dgdMCRepair_Sub.Items.Remove((dgdMCRepair_Sub.Items[dgdMCRepair_Sub.Items.Count - 1]) as Win_prd_Repair_U_Sub_CodeView);
                }
                dgdMCRepair_Sub.Refresh();
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
            Lib.Instance.DataGridINBothByMouseUP(sender, e);
        }

        //부품
        private void dgdtpeMCPartName_KeyDown(object sender, KeyEventArgs e)
        {
            if (lblMsg.Visibility == Visibility.Visible)
            {
                WinMCRepairSub = dgdMCRepair_Sub.CurrentItem as Win_prd_Repair_U_Sub_CodeView;
                int rowCount = dgdMCRepair_Sub.Items.IndexOf(dgdMCRepair_Sub.CurrentItem);
                int colCount = dgdMCRepair_Sub.Columns.IndexOf(dgdtpecustomname);

                if (e.Key == Key.Enter)
                {
                    e.Handled = true;
                    (sender as DataGridCell).IsEditing = false;
                    dgdMCRepair_Sub.CurrentCell = new DataGridCellInfo(dgdMCRepair_Sub.Items[rowCount], dgdMCRepair_Sub.Columns[colCount]);
                }
                else if (e.Key == Key.Delete)
                {
                    SubRemove();

                    if (dgdMCRepair_Sub.Items.Count > 0)
                    {
                        dgdMCRepair_Sub.Focus();
                        //dgdMCRepair_Sub.SelectedIndex = dgdMCRepair_Sub.Items.Count - 1;
                        if (rowCount == 0)
                        {
                            dgdMCRepair_Sub.CurrentCell = new DataGridCellInfo(dgdMCRepair_Sub.Items[rowCount], dgdMCRepair_Sub.Columns[dgdMCRepair_Sub.Columns.IndexOf(dgdtpeMCPartName)]);
                        }
                        else
                        {
                            dgdMCRepair_Sub.CurrentCell = new DataGridCellInfo(dgdMCRepair_Sub.Items[rowCount - 1], dgdMCRepair_Sub.Columns[dgdMCRepair_Sub.Columns.IndexOf(dgdtpeMCPartName)]);
                        }
                    }
                }
            }
        }

        //부품
        private void EditableMCPartName_KeyDown(object sender, KeyEventArgs e)
        {
            WinMCRepairSub = dgdMCRepair_Sub.CurrentItem as Win_prd_Repair_U_Sub_CodeView;
            int rowCount = dgdMCRepair_Sub.Items.IndexOf(dgdMCRepair_Sub.CurrentItem);
            if (e.Key == Key.Enter)
            {
                TextBox tb1 = sender as TextBox;
                MainWindow.pf.ReturnCode(tb1, (int)Defind_CodeFind.DCF_PART, "");

                if (tb1.Tag != null && !tb1.Tag.ToString().Equals(""))
                {
                    WinMCRepairSub.MCPartName = tb1.Text;
                    WinMCRepairSub.MCPartID = tb1.Tag.ToString();

                    String[] lst = GetMcPart(WinMCRepairSub.MCPartID);
                    if (lst != null)
                    {
                        WinMCRepairSub.CustomID = lst[0];
                        WinMCRepairSub.customname = lst[1];

                    }

                }
                sender = tb1;
            }
        }

        //예비품 입력시 거래처,종류 가져오게 
        private String[] GetMcPart(string McPartID)
        {
            String[] lst = new string[3];
            string sql = "select CustomID, Custom, ForUse from mt_McPart mmp left join mc_PartStuffIN mps on mps.MCPartID = mmp.MCPartID where mmp.MCPartID = " + McPartID;

            try
            {
                DataSet ds = DataStore.Instance.QueryToDataSet(sql);
                if (ds != null && ds.Tables.Count > 0)
                {
                    DataTable dt = ds.Tables[0];
                    if (dt.Rows.Count > 0)
                    {
                        DataRowCollection drc = dt.Rows;
                        DataRow dr = drc[0];

                        lst[0] = dr["CustomID"].ToString();
                        lst[1] = dr["Custom"].ToString();

                    }
                    return lst;
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
            return null;
        }



        //부품
        private void EditableMCPartName_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            WinMCRepairSub = dgdMCRepair_Sub.CurrentItem as Win_prd_Repair_U_Sub_CodeView;
            int rowCount = dgdMCRepair_Sub.Items.IndexOf(dgdMCRepair_Sub.CurrentItem);

            TextBox tb1 = sender as TextBox;
            MainWindow.pf.ReturnCode(tb1, (int)Defind_CodeFind.DCF_PART, "");

            if (tb1.Tag != null)
            {
                WinMCRepairSub.MCPartName = tb1.Text;
                WinMCRepairSub.MCPartID = tb1.Tag.ToString();
            }

            sender = tb1;
        }

        //부품_enter key 없이 값 대입
        private void EditableMCPartName_LostFocus(object sender, RoutedEventArgs e)
        {
            if (lblMsg.Visibility == Visibility.Visible)
            {
                WinMCRepairSub = dgdMCRepair_Sub.CurrentItem as Win_prd_Repair_U_Sub_CodeView;
                if (WinMCRepairSub != null)
                {
                    TextBox tb1 = sender as TextBox;

                    if (tb1.Tag != null)
                    {
                        WinMCRepairSub.MCPartName = tb1.Text;
                        WinMCRepairSub.MCPartID = tb1.Tag.ToString();
                    }
                }
            }
        }

        //구입처
        private void dgdtpecustomname_KeyDown(object sender, KeyEventArgs e)
        {
            if (lblMsg.Visibility == Visibility.Visible)
            {
                WinMCRepairSub = dgdMCRepair_Sub.CurrentItem as Win_prd_Repair_U_Sub_CodeView;
                int rowCount = dgdMCRepair_Sub.Items.IndexOf(dgdMCRepair_Sub.CurrentItem);
                int colCount = dgdMCRepair_Sub.Columns.IndexOf(dgdtpepartcnt);
                if (e.Key == Key.Enter)
                {
                    e.Handled = true;
                    (sender as DataGridCell).IsEditing = false;
                    dgdMCRepair_Sub.CurrentCell = new DataGridCellInfo(dgdMCRepair_Sub.Items[rowCount], dgdMCRepair_Sub.Columns[colCount]);
                }
                else if (e.Key == Key.Delete)
                {
                    SubRemove();

                    if (dgdMCRepair_Sub.Items.Count > 0)
                    {
                        dgdMCRepair_Sub.Focus();
                        //dgdMCRepair_Sub.SelectedIndex = dgdMCRepair_Sub.Items.Count - 1;
                        if (rowCount == 0)
                        {
                            dgdMCRepair_Sub.CurrentCell = new DataGridCellInfo(dgdMCRepair_Sub.Items[rowCount], dgdMCRepair_Sub.Columns[dgdMCRepair_Sub.Columns.IndexOf(dgdtpeMCPartName)]);
                        }
                        else
                        {
                            dgdMCRepair_Sub.CurrentCell = new DataGridCellInfo(dgdMCRepair_Sub.Items[rowCount - 1], dgdMCRepair_Sub.Columns[dgdMCRepair_Sub.Columns.IndexOf(dgdtpeMCPartName)]);
                        }
                    }
                }
            }
        }

        //구입처
        private void Editablecustomname_KeyDown(object sender, KeyEventArgs e)
        {
            WinMCRepairSub = dgdMCRepair_Sub.CurrentItem as Win_prd_Repair_U_Sub_CodeView;
            int rowCount = dgdMCRepair_Sub.Items.IndexOf(dgdMCRepair_Sub.CurrentItem);
            if (e.Key == Key.Enter)
            {
                TextBox tb1 = sender as TextBox;
                MainWindow.pf.ReturnCode(tb1, (int)Defind_CodeFind.DCF_CUSTOM, "");

                if (tb1.Tag != null)
                {
                    WinMCRepairSub.customname = tb1.Text;
                    WinMCRepairSub.CustomID = tb1.Tag.ToString();
                }

                sender = tb1;
            }
        }

        //구입처
        private void Editablecustomname_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            if (lblMsg.Visibility == Visibility.Visible)
            {
                WinMCRepairSub = dgdMCRepair_Sub.CurrentItem as Win_prd_Repair_U_Sub_CodeView;

                TextBox tb1 = sender as TextBox;
                MainWindow.pf.ReturnCode(tb1, (int)Defind_CodeFind.DCF_CUSTOM, "");

                if (tb1.Tag != null)
                {
                    WinMCRepairSub.customname = tb1.Text;
                    WinMCRepairSub.CustomID = tb1.Tag.ToString();
                }

                sender = tb1;
            }
        }

        //구입처
        private void Editablecustomname_LostFocus(object sender, RoutedEventArgs e)
        {
            if (lblMsg.Visibility == Visibility.Visible)
            {
                WinMCRepairSub = dgdMCRepair_Sub.CurrentItem as Win_prd_Repair_U_Sub_CodeView;

                if (WinMCRepairSub != null)
                {
                    TextBox tb1 = sender as TextBox;

                    if (tb1.Tag != null)
                    {
                        WinMCRepairSub.customname = tb1.Text;
                        WinMCRepairSub.CustomID = tb1.Tag.ToString();
                    }

                    sender = tb1;
                }
            }
        }

        //수량
        private void dgdtpepartcnt_KeyDown(object sender, KeyEventArgs e)
        {
            if (lblMsg.Visibility == Visibility.Visible)
            {
                WinMCRepairSub = dgdMCRepair_Sub.CurrentItem as Win_prd_Repair_U_Sub_CodeView;
                int rowCount = dgdMCRepair_Sub.Items.IndexOf(dgdMCRepair_Sub.CurrentItem);
                int colCount = dgdMCRepair_Sub.Columns.IndexOf(dgdtpepartprice);

                if (e.Key == Key.Enter)
                {
                    e.Handled = true;
                    (sender as DataGridCell).IsEditing = false;
                    dgdMCRepair_Sub.CurrentCell = new DataGridCellInfo(dgdMCRepair_Sub.Items[rowCount], dgdMCRepair_Sub.Columns[colCount]);
                }
                else if (e.Key == Key.Delete)
                {
                    SubRemove();

                    if (dgdMCRepair_Sub.Items.Count > 0)
                    {
                        dgdMCRepair_Sub.Focus();
                        //dgdMCRepair_Sub.SelectedIndex = dgdMCRepair_Sub.Items.Count - 1;
                        if (rowCount == 0)
                        {
                            dgdMCRepair_Sub.CurrentCell = new DataGridCellInfo(dgdMCRepair_Sub.Items[rowCount], dgdMCRepair_Sub.Columns[dgdMCRepair_Sub.Columns.IndexOf(dgdtpeMCPartName)]);
                        }
                        else
                        {
                            dgdMCRepair_Sub.CurrentCell = new DataGridCellInfo(dgdMCRepair_Sub.Items[rowCount - 1], dgdMCRepair_Sub.Columns[dgdMCRepair_Sub.Columns.IndexOf(dgdtpeMCPartName)]);
                        }
                    }
                }
            }
        }

        //수량
        private void Editablepartcnt_KeyDown(object sender, KeyEventArgs e)
        {
            var WinRepairSub = dgdMCRepair_Sub.CurrentItem as Win_prd_Repair_U_Sub_CodeView;
            int rowCount = dgdMCRepair_Sub.Items.IndexOf(dgdMCRepair_Sub.CurrentItem);

            if (e.Key == Key.Enter)
            {
                TextBox tb1 = sender as TextBox;

                WinRepairSub.partcnt = tb1.Text;
                sender = tb1;
            }
        }

        //수량_enter 이벤트 없이도 값 대입
        private void Editablepartcnt_LostFocus(object sender, RoutedEventArgs e)
        {
            var WinRepairSub = dgdMCRepair_Sub.CurrentItem as Win_prd_Repair_U_Sub_CodeView;

            if (WinRepairSub != null)
            {
                TextBox tb1 = sender as TextBox;
                WinRepairSub.partcnt = tb1.Text;
                sender = tb1;
            }
        }

        //수리비용
        private void dgdtpepartprice_KeyDown(object sender, KeyEventArgs e)
        {
            if (lblMsg.Visibility == Visibility.Visible)
            {
                WinMCRepairSub = dgdMCRepair_Sub.CurrentItem as Win_prd_Repair_U_Sub_CodeView;
                int rowCount = dgdMCRepair_Sub.Items.IndexOf(dgdMCRepair_Sub.CurrentItem);
                int colCount = dgdMCRepair_Sub.Columns.IndexOf(dgdtpereason);
                if (e.Key == Key.Enter)
                {
                    e.Handled = true;
                    (sender as DataGridCell).IsEditing = false;
                    dgdMCRepair_Sub.CurrentCell = new DataGridCellInfo(dgdMCRepair_Sub.Items[rowCount], dgdMCRepair_Sub.Columns[colCount]);
                }
                else if (e.Key == Key.Delete)
                {
                    SubRemove();

                    if (dgdMCRepair_Sub.Items.Count > 0)
                    {
                        dgdMCRepair_Sub.Focus();
                        //dgdMCRepair_Sub.SelectedIndex = dgdMCRepair_Sub.Items.Count - 1;
                        if (rowCount == 0)
                        {
                            dgdMCRepair_Sub.CurrentCell = new DataGridCellInfo(dgdMCRepair_Sub.Items[rowCount], dgdMCRepair_Sub.Columns[dgdMCRepair_Sub.Columns.IndexOf(dgdtpeMCPartName)]);
                        }
                        else
                        {
                            dgdMCRepair_Sub.CurrentCell = new DataGridCellInfo(dgdMCRepair_Sub.Items[rowCount - 1], dgdMCRepair_Sub.Columns[dgdMCRepair_Sub.Columns.IndexOf(dgdtpeMCPartName)]);
                        }
                    }
                }
            }
        }

        //수리비용
        private void Editablepartprice_KeyDown(object sender, KeyEventArgs e)
        {
            var WinRepairSub = dgdMCRepair_Sub.CurrentItem as Win_prd_Repair_U_Sub_CodeView;
            int rowCount = dgdMCRepair_Sub.Items.IndexOf(dgdMCRepair_Sub.CurrentItem);
            if (e.Key == Key.Enter)
            {
                TextBox tb1 = sender as TextBox;

                WinRepairSub.partprice = tb1.Text;
                sender = tb1;
            }
        }

        //수리비용
        private void Editablepartprice_LostFocus(object sender, RoutedEventArgs e)
        {
            if (lblMsg.Visibility == Visibility.Visible)
            {
                var WinRepairSub = dgdMCRepair_Sub.CurrentItem as Win_prd_Repair_U_Sub_CodeView;

                if (WinRepairSub != null)
                {
                    TextBox tb1 = sender as TextBox;

                    WinRepairSub.partprice = tb1.Text;
                    sender = tb1;
                }
            }
        }

        //수량, 수리비용(숫자만 입력)
        private void TextBox_PreviewTextInput(object sender, TextCompositionEventArgs e)
        {
            Lib.Instance.CheckIsNumericbyThree((TextBox)sender, e);
        }

        //사유
        private void dgdtpereason_KeyDown(object sender, KeyEventArgs e)
        {
            if (lblMsg.Visibility == Visibility.Visible)
            {
                WinMCRepairSub = dgdMCRepair_Sub.CurrentItem as Win_prd_Repair_U_Sub_CodeView;
                int rowCount = dgdMCRepair_Sub.Items.IndexOf(dgdMCRepair_Sub.CurrentItem);
                int colCount = dgdMCRepair_Sub.Columns.IndexOf(dgdtpepartremark);
                if (e.Key == Key.Enter)
                {
                    e.Handled = true;
                    (sender as DataGridCell).IsEditing = false;
                    dgdMCRepair_Sub.CurrentCell = new DataGridCellInfo(dgdMCRepair_Sub.Items[rowCount], dgdMCRepair_Sub.Columns[colCount]);
                }
                else if (e.Key == Key.Delete)
                {
                    SubRemove();

                    if (dgdMCRepair_Sub.Items.Count > 0)
                    {
                        dgdMCRepair_Sub.Focus();
                        //dgdMCRepair_Sub.SelectedIndex = dgdMCRepair_Sub.Items.Count - 1;
                        if (rowCount == 0)
                        {
                            dgdMCRepair_Sub.CurrentCell = new DataGridCellInfo(dgdMCRepair_Sub.Items[rowCount], dgdMCRepair_Sub.Columns[dgdMCRepair_Sub.Columns.IndexOf(dgdtpeMCPartName)]);
                        }
                        else
                        {
                            dgdMCRepair_Sub.CurrentCell = new DataGridCellInfo(dgdMCRepair_Sub.Items[rowCount - 1], dgdMCRepair_Sub.Columns[dgdMCRepair_Sub.Columns.IndexOf(dgdtpeMCPartName)]);
                        }

                    }
                }
            }
        }

        //사유
        private void Editablereason_KeyDown(object sender, KeyEventArgs e)
        {
            var WinRepairSub = dgdMCRepair_Sub.CurrentItem as Win_prd_Repair_U_Sub_CodeView;
            int rowCount = dgdMCRepair_Sub.Items.IndexOf(dgdMCRepair_Sub.CurrentItem);
            if (e.Key == Key.Enter)
            {
                TextBox tb1 = sender as TextBox;
                WinRepairSub.reason = tb1.Text;
                sender = tb1;
            }
        }

        //사유
        private void Editablereason_LostFocus(object sender, RoutedEventArgs e)
        {
            if (lblMsg.Visibility == Visibility.Visible)
            {
                var WinRepairSub = dgdMCRepair_Sub.CurrentItem as Win_prd_Repair_U_Sub_CodeView;

                if (WinRepairSub != null)
                {
                    TextBox tb1 = sender as TextBox;
                    WinRepairSub.reason = tb1.Text;
                    sender = tb1;
                }
            }
        }

        //비고사항
        private void dgdtpepartremark_KeyDown(object sender, KeyEventArgs e)
        {
            if (lblMsg.Visibility == Visibility.Visible)
            {
                var WinMCRepairSub = dgdMCRepair_Sub.CurrentItem as Win_prd_Repair_U_Sub_CodeView;
                int rowCount = dgdMCRepair_Sub.Items.IndexOf(dgdMCRepair_Sub.CurrentItem);
                int colCount = dgdMCRepair_Sub.Columns.IndexOf(dgdtpeMCPartName);
                if (e.Key == Key.Enter)
                {
                    e.Handled = true;
                    (sender as DataGridCell).IsEditing = false;
                    //if (dgdMCRepair_Sub.Items.Count - 1 == rowCount)
                    //{
                    //    SubPlus();
                    //}

                    //dgdMCRepair_Sub.Focus();
                    //dgdMCRepair_Sub.SelectedIndex = rowCount+1;                    
                    //dgdMCRepair_Sub.CurrentCell = new DataGridCellInfo(dgdMCRepair_Sub.Items[rowCount + 1], dgdMCRepair_Sub.Columns[colCount]);
                }
                else if (e.Key == Key.Delete)
                {
                    SubRemove();

                    if (dgdMCRepair_Sub.Items.Count > 0)
                    {
                        //dgdMCRepair_Sub.Focus();
                        //dgdMCRepair_Sub.SelectedIndex = dgdMCRepair_Sub.Items.Count - 1;
                        if (rowCount == 0)
                        {
                            dgdMCRepair_Sub.CurrentCell = new DataGridCellInfo(dgdMCRepair_Sub.Items[rowCount], dgdMCRepair_Sub.Columns[dgdMCRepair_Sub.Columns.IndexOf(dgdtpeMCPartName)]);
                        }
                        else
                        {
                            dgdMCRepair_Sub.CurrentCell = new DataGridCellInfo(dgdMCRepair_Sub.Items[rowCount - 1], dgdMCRepair_Sub.Columns[dgdMCRepair_Sub.Columns.IndexOf(dgdtpeMCPartName)]);
                        }

                    }
                }
            }
        }

        //비고사항
        private void Editablepartremark_KeyDown(object sender, KeyEventArgs e)
        {
            var WinRepairSub = dgdMCRepair_Sub.CurrentItem as Win_prd_Repair_U_Sub_CodeView;
            int rowCount = dgdMCRepair_Sub.Items.IndexOf(dgdMCRepair_Sub.CurrentItem);
            if (e.Key == Key.Enter)
            {
                TextBox tb1 = sender as TextBox;
                WinRepairSub.partremark = tb1.Text;
                sender = tb1;
            }
        }

        //비고사항
        private void Editablepartremark_LostFocus(object sender, RoutedEventArgs e)
        {
            if (lblMsg.Visibility == Visibility.Visible)
            {
                var WinRepairSub = dgdMCRepair_Sub.CurrentItem as Win_prd_Repair_U_Sub_CodeView;

                if (WinRepairSub != null)
                {
                    TextBox tb1 = sender as TextBox;
                    WinRepairSub.partremark = tb1.Text;
                    sender = tb1;
                }
            }
        }

        //TextBox Focus
        private void TextBox_InputFocus(object sender, KeyEventArgs e)
        {
            Lib.Instance.DataGridINTextBoxFocus(sender, e);
        }

        //(서브)모든 셀에 삭제 적용하기 위해(TextBox 안해서는 keydown이 안먹힌다.)
        //PreviewKewDown 적용, 셀의 delete 이벤트를 탈수 있게 포커스 바꿔줌
        private void DataGridInTextBox_PreviewKeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Delete)
            {
                DataGridCell cell = Lib.Instance.GetParent<DataGridCell>(sender as TextBox);
                cell.Focus();
            }
        }

        #endregion

        #region 플러스 파인더 및 enter focus move 

        //처리구분
        private void cboRepairGubun_DropDownClosed(object sender, EventArgs e)
        {
            txtmcname.Focus();
        }

        //기계명 추가,수정 시 입력에서 keydown 플러스파이더 호출
        private void txtmcname_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                MainWindow.pf.ReturnCode(txtmcname, (int)Defind_CodeFind.DCF_MC, "");

                if (txtmcname.Tag != null)
                {
                    GetInfoByMcID(txtmcname.Tag.ToString());
                }

                txtmanagerid.Focus();
            }
        }

        //기계명 추가,수정 시 입력에서 버튼클릭 플러스파이더 호출
        private void btnPfmcname_Click(object sender, RoutedEventArgs e)
        {
            MainWindow.pf.ReturnCode(txtmcname, (int)Defind_CodeFind.DCF_MC, "");

            if (txtmcname.Tag != null)
            {
                GetInfoByMcID(txtmcname.Tag.ToString());
            }

            txtmanagerid.Focus();
        }

        //입력시 기계 정보 GET
        private void GetInfoByMcID(string strMcID)
        {
            string sql = "SELECT mcid, mcname, managerid, customid, customname, buycustomid, buycustomname," +
                " personid = (select PersonID from mt_Person where Name=mm.personid), personname FROM mt_MC mm WHERE mcid =" + strMcID;

            try
            {
                DataSet ds = DataStore.Instance.QueryToDataSet(sql);
                if (ds != null && ds.Tables.Count > 0)
                {
                    DataTable dt = ds.Tables[0];
                    if (dt.Rows.Count > 0)
                    {
                        DataRowCollection drc = dt.Rows;
                        DataRow dr = drc[0];

                        //txtmcname.Text = dr["mcname"].ToString();
                        //txtmcname.Tag = dr["mcid"].ToString();
                        txtmanagerid.Text = dr["managerid"].ToString();
                        txtcustomname.Text = dr["customname"].ToString(); //2021-05-06 제작사도 가져오게 하기 위해 추가
                        txtcustomname.Tag = dr["customid"].ToString();    //2021-05-06 제작사도 가져오게 하기 위해 추가
                        txtBuyCustomName.Text = dr["buycustomname"].ToString();
                        txtBuyCustomName.Tag = dr["buycustomid"].ToString();
                        txtpersonname.Text = dr["personname"].ToString();
                        txtpersonname.Tag = dr["personid"].ToString();
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

        //관리번호 
        private void txtmanagerid_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                txtcustomname.Focus();
            }
        }

        //제작사 추가,수정 시 입력에서 keydown 플러스파이더 호출
        private void txtcustomname_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                MainWindow.pf.ReturnCode(txtcustomname, (int)Defind_CodeFind.DCF_CUSTOM, "");
                txtBuyCustomName.Focus();
            }
        }

        //제작사 추가,수정 시 입력에서 버튼클릭 플러스파이더 호출
        private void btnPfcustomname_Click(object sender, RoutedEventArgs e)
        {
            MainWindow.pf.ReturnCode(txtcustomname, (int)Defind_CodeFind.DCF_CUSTOM, "");
            txtBuyCustomName.Focus();
        }

        //제작사(외주)
        private void txtBuyCustomName_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                dtprepairdate.Focus();
            }
        }

        //수리일자
        private void dtprepairdate_PreviewKeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                dtprepairdate.IsDropDownOpen = true;
            }
        }

        //수리일자
        private void dtprepairdate_CalendarClosed(object sender, RoutedEventArgs e)
        {
            txtpersonname.Focus();
        }

        //관리담당자 추가,수정 시 입력에서 keydown 플러스파이더 호출
        private void txtpersonname_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                MainWindow.pf.ReturnCode(txtpersonname, (int)Defind_CodeFind.DCF_PERSON, "");
                txtRepairTime.Focus();
            }
        }

        //관리담당자 추가,수정 시 입력에서 버튼클릭 플러스파이더 호출
        private void btnpersonname_Click(object sender, RoutedEventArgs e)
        {
            MainWindow.pf.ReturnCode(txtpersonname, (int)Defind_CodeFind.DCF_PERSON, "");
            txtRepairTime.Focus();
        }

        //수리시간
        private void txtRepairTime_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                txtRepairRemark.Focus();
            }
        }

        //부품명
        private void lblArticleSearch_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            if (chkArticleSearch.IsChecked == true) { chkArticleSearch.IsChecked = false; }
            else { chkArticleSearch.IsChecked = true; }
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

        //부품명 플러스 파인더
        private void btnPfArticleSearch_Click(object sender, RoutedEventArgs e)
        {
            MainWindow.pf.ReturnCode(txtArticleSearch, (int)Defind_CodeFind.DCF_PART, "");
        }

        //부품명 keyDown
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
        #endregion

        private void btnLastMonth_Click(object sender, RoutedEventArgs e)
        {
            try
            {
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
            catch (Exception ee)
            {
                MessageBox.Show("오류지점 - btnLastMonth_Click : " + ee.ToString());
            }
        }

        private void btnYesterday_Click(object sender, RoutedEventArgs e)
        {
            try
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
            catch (Exception ee)
            {
                MessageBox.Show("오류지점 - btnYesterday_Click : " + ee.ToString());
            }
        }



    }

    class Win_prd_Repair_U_CodeView : BaseView
    {
        public override string ToString()
        {
            return (this.ReportAllProperties());
        }

        public int Num { get; set; }
        public string RepairID { get; set; }
        public string repairdate { get; set; }
        public string repairTime { get; set; }
        public string RepairGubun { get; set; }
        public string mcid { get; set; }
        public string mcname { get; set; }
        public string managerid { get; set; }
        public string customid { get; set; }
        public string customname { get; set; }
        public string buycustomid { get; set; }
        public string BuyCustomName { get; set; }
        public string personid { get; set; }
        public string personname { get; set; }
        public string RepairRemark { get; set; }
        public string repairdate_CV { get; set; }
        public string RepairGubun_CV { get; set; }
        public double price { get; set; }
    }
    
    class Win_prd_Repair_U_Sub_CodeView : BaseView
    {
        public override string ToString()
        {
            return (this.ReportAllProperties());
        }

        public string RepairID { get; set; }
        public string repairSubseq { get; set; }
        public string MCPartID { get; set; }
        public string MCPartName { get; set; }
        public string reason { get; set; }
        public string partcnt { get; set; }
        public string partprice { get; set; }
        public string partremark { get; set; }
        public string CustomID { get; set; }
        public string customname { get; set; }

        public bool flagMCPartName { get; set; }
        public bool flagcustomname { get; set; }
        public bool flagpartcnt { get; set; }
        public bool flagpartprice { get; set; }
        public bool flagreason { get; set; }
        public bool flagpartremark { get; set; }
    }
}

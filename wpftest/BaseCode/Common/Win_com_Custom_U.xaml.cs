using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Data;
using System.Threading;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using WizMes_WellMade.PopUp;
using WizMes_WellMade.PopUP;

namespace WizMes_WellMade
{
    /// <summary>
    /// Win_com_Custom_U.xaml에 대한 상호 작용 논리
    /// </summary>
    public partial class Win_com_Custom_U : UserControl
    {
        #region 변수 선언 및 로드

        Lib lib = new Lib();
        PlusFinder pf = new PlusFinder();

        string stDate = string.Empty;
        string stTime = string.Empty;

        string strBasisID = string.Empty;
        //int dgdInComBoNum = 0;
        string InspectDate = string.Empty;
        //int InspectNum = 0;

        string InspectName = string.Empty;

        string AASS = string.Empty;

        string strFlag = string.Empty;
        int rowNum = 0;
        Win_com_Custom_U_CodeView WinCustom = new Win_com_Custom_U_CodeView();

        public Win_com_Custom_U()
        {
            InitializeComponent();
        }

        private void UserControl_Loaded(object sender, RoutedEventArgs e)
        {
            stDate = DateTime.Now.ToString("yyyyMMdd");
            stTime = DateTime.Now.ToString("HHmm");

            DataStore.Instance.InsertLogByFormS(this.GetType().Name, stDate, stTime, "S");

            Lib.Instance.UiLoading(sender);

            SetComboBox();
        }

        //콤보박스 세팅
        private void SetComboBox()
        {
            List<string[]> lstDvlYN = new List<string[]>();
            string[] strDvl_1 = { "Y", "Y" };
            string[] strDvl_2 = { "N", "N" };
            lstDvlYN.Add(strDvl_1);
            lstDvlYN.Add(strDvl_2);

            ObservableCollection<CodeView> ovcTrade = ComboBoxUtil.Instance.Gf_DB_CM_GetComCodeDataset(null, "CMMTRAD", "Y", "");
            this.cboTradeSrh.ItemsSource = ovcTrade;
            this.cboTradeSrh.DisplayMemberPath = "code_name";
            this.cboTradeSrh.SelectedValuePath = "code_id";

            this.cboTrade.ItemsSource = ovcTrade;
            this.cboTrade.DisplayMemberPath = "code_name";
            this.cboTrade.SelectedValuePath = "code_id";

            ObservableCollection<CodeView> ovcDvlYN = ComboBoxUtil.Instance.Direct_SetComboBox(lstDvlYN);
            this.cboFTAMgrYN.ItemsSource = ovcDvlYN;
            this.cboFTAMgrYN.DisplayMemberPath = "code_name";
            this.cboFTAMgrYN.SelectedValuePath = "code_id";

            //FTA 중점관리품 여부(입력)
            this.cboFTAMgrYN.ItemsSource = ovcDvlYN;
            this.cboFTAMgrYN.DisplayMemberPath = "code_name";
            this.cboFTAMgrYN.SelectedValuePath = "code_id";
        }

        #endregion

        #region 검색 조건

        // 상호 검색어
        private void lblCustomSrh_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
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
        private void chkCustomSrh_Checked(object sender, RoutedEventArgs e)
        {
            chkCustomSrh.IsChecked = true;
            txtCustomSrh.IsEnabled = true;
        }
        private void chkCustomSrh_Unchecked(object sender, RoutedEventArgs e)
        {
            chkCustomSrh.IsChecked = false;
            txtCustomSrh.IsEnabled = false;
        }

        // 거래 구분
        private void lblTradeSrh_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            if (chkTradeSrh.IsChecked == true)
            {
                chkTradeSrh.IsChecked = false;
            }
            else
            {
                chkTradeSrh.IsChecked = true;
            }
        }
        private void chkTradeSrh_Checked(object sender, RoutedEventArgs e)
        {
            //chkCustomSrh.IsChecked = true;
            cboTradeSrh.IsEnabled = true;
        }
        private void chkTradeSrh_Unchecked(object sender, RoutedEventArgs e)
        {
            //chkCustomSrh.IsChecked = false;
            cboTradeSrh.IsEnabled = false;
        }

        #endregion // 검색 조건

        #region 버튼 클릭

        //취소, 저장 후
        private void CanBtnControl()
        {
            Lib.Instance.UiButtonEnableChange_IUControl(this);
            //btnHTML.Visibility = Visibility.Visible;

            gbxInput.IsHitTestVisible = false;
            dgdCustom.IsEnabled = true;
            
        }

        //추가, 수정 클릭 시
        private void CantBtnControl()
        {
            Lib.Instance.UiButtonEnableChange_SCControl(this);
            //btnHTML.Visibility = Visibility.Hidden;

            gbxInput.IsHitTestVisible = true;
            dgdCustom.IsEnabled = false;
            

        }

        //HTML
        private void btnHTML_Click(object sender, RoutedEventArgs e)
        {
            string url = string.Empty;

            try
            {

                url = txtHomepage.Text.Trim();

                // 주소가 등록되지 않았을시
                if (url == "" || url == null)
                {
                    MessageBox.Show("홈페이지 주소가 등록되지 않았습니다.");
                    return;
                }

                System.Diagnostics.Process.Start(url);
            }
            catch (Exception)
            {
                MessageBox.Show("올바른 주소가 아닙니다. (ex : www.google.com)");
            }
        }

        // 사용안함 포함 체크박스
        private void chkNoUse_Checked(object sender, RoutedEventArgs e)
        {
            chkNoUse.IsChecked = true;
        }

        // 사용안함 포함 체크박스
        private void chkNoUse_UnChecked(object sender, RoutedEventArgs e)
        {
            chkNoUse.IsChecked = false;
        }

        //추가
        private void btnAdd_Click(object sender, RoutedEventArgs e)
        {
            if (dgdCustom.SelectedItem != null)
            {
                // 선택된 행 번호 가져오기
                rowNum = dgdCustom.SelectedIndex;

                ClearData();
            }
            else
            {
                ClearData();
            }

            cboTrade.SelectedIndex = 0;
            cboFTAMgrYN.SelectedIndex = 1;

            CantBtnControl();
            txtCustomID.IsReadOnly = false;

            tbkMsg.Text = "자료 추가 중";
            strFlag = "I";
        }

        //수정
        private void btnUpdate_Click(object sender, RoutedEventArgs e)
        {
            if (dgdCustom.SelectedItem != null)
            {
                rowNum = dgdCustom.SelectedIndex;
                CantBtnControl();
                txtCustomID.IsReadOnly = true;
                tbkMsg.Text = "자료 수정 중";
                strFlag = "U";
            }
            else
            {
                MessageBox.Show("수정할 자료를 선택하고 눌러주십시오.");
            }
        }

        //삭제
        private void btnDelete_Click(object sender, RoutedEventArgs e)
        {
            if (dgdCustom.SelectedItem == null)
            {
                MessageBox.Show("삭제할 데이터가 지정되지 않았습니다. 삭제데이터를 지정하고 눌러주세요");
            }
            else
            {

                var winCustom = dgdCustom.SelectedItem as Win_com_Custom_U_CodeView;

                if (MessageBox.Show("선택하신 항목을 삭제하시겠습니까?", "삭제 전 확인", MessageBoxButton.YesNo) == MessageBoxResult.Yes)
                {
                    //삭제 전 체크
                    if (!DeleteDataCheck(winCustom.CustomID))
                        return;

                    if (dgdCustom.Items.Count > 0 && dgdCustom.SelectedItem != null)
                    {
                        rowNum = dgdCustom.SelectedIndex - 1;
                    }

                    if (DeleteData(winCustom.CustomID))
                    {
                        re_Search(rowNum);
                    }
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

        }

        //검색
        private void btnSearch_Click(object sender, RoutedEventArgs e)
        {
            //검색버튼 비활성화
            //btnSearch.IsEnabled = false;

            rowNum = 0;

            using (Loading lw = new Loading(re_Search))
            {
                lw.ShowDialog();
            }

            if (dgdCustom.Items.Count == 0)
            {
                this.DataContext = null;
                MessageBox.Show("조회된 데이터가 없습니다.");
                return;
            }

            //Dispatcher.BeginInvoke(new Action(() =>

            //{
            //    Thread.Sleep(2000);

            //    //로직
            //    ClearData();
            //    rowNum = 0;
            //    re_Search(rowNum);

            //}), System.Windows.Threading.DispatcherPriority.Background);



            //Dispatcher.BeginInvoke(new Action(() =>

            //{
            //    btnSearch.IsEnabled = true;

            //}), System.Windows.Threading.DispatcherPriority.Background);
        }

        //// 저장
        //private void btnSave_Click(object sender, RoutedEventArgs e)
        //{
        //    var WinCustom = dgdCustom.SelectedItem as Win_com_Custom_U_CodeView;

        //    if (strFlag.Equals("I"))
        //    {
        //        if (SaveData(txtCustomID.Text, strFlag))
        //        {
        //            CanBtnControl();
        //            rowNum = 0;
        //            re_Search(rowNum);
        //        }
        //    }
        //    else
        //    {
        //        if (SaveData(txtCustomID.Text, strFlag))
        //        {
        //            CanBtnControl();
        //            re_Search(rowNum);
        //        }
        //    }
        //}

        //저장
        private void btnSave_Click(object sender, RoutedEventArgs e)
        {

            using (Loading lw = new Loading(beSave))
            {

                lw.ShowDialog();
            }
            //re_Search(rowNum);
        }

        private void beSave()
        {
            if (SaveData(txtCustomID.Text, strFlag))
            {
                CanBtnControl();
                strBasisID = string.Empty;
                lblMsg.Visibility = Visibility.Hidden;

                if (strFlag.Equals("I"))
                {
                    InspectName = txtCustomID.ToString();
                    //InspectName = txtKCustom.ToString();
                    //InspectDate = dtpInspectDate.SelectedDate.ToString().Substring(0, 10);

                    rowNum = 0;
                    re_Search(rowNum);
                    return;

                }
                else
                {
                    rowNum = dgdCustom.SelectedIndex;
                }
            }

        }

        //취소
        private void btnCancel_Click(object sender, RoutedEventArgs e)
        {
            using (Loading lw = new Loading(beCancel))
            {
                lw.ShowDialog();
            }

            if (dgdCustom.Items.Count == 0)
            {
                this.DataContext = null;
                MessageBox.Show("조회된 데이터가 없습니다.");
                return;
            }
        }

        private void beCancel()
        {
            ClearData();

            CanBtnControl();
            strFlag = string.Empty;
            re_Search(rowNum);
        }

        //엑셀
        private void btnExcel_Click(object sender, RoutedEventArgs e)
        {
            DataTable dt = null;
            string Name = string.Empty;

            string[] dgdStr = new string[2];
            dgdStr[0] = "거래처 정보";
            dgdStr[1] = dgdCustom.Name;

            ExportExcelxaml ExpExc = new ExportExcelxaml(dgdStr);
            ExpExc.ShowDialog();

            if (ExpExc.DialogResult.HasValue)
            {
                if (ExpExc.choice.Equals(dgdCustom.Name))
                {
                    DataStore.Instance.InsertLogByForm(this.GetType().Name, "E");
                    if (ExpExc.Check.Equals("Y"))
                        dt = Lib.Instance.DataGridToDTinHidden(dgdCustom);
                    else
                        dt = Lib.Instance.DataGirdToDataTable(dgdCustom);

                    Name = dgdCustom.Name;

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

        //재검색
        private void re_Search(int selectIndex)
        {
            FillGrid();

            if (dgdCustom.Items.Count > 0)
            {
                dgdCustom.SelectedIndex = selectIndex;
            }

            rowNum = 0;
        }

        private void re_Search()
        {
            FillGrid();

            if (dgdCustom.Items.Count > 0)
            {
                dgdCustom.SelectedIndex = rowNum;
            }
            else
            {
                MessageBox.Show("조회된 데이터가 없습니다.");
                return;
            }
        }

        #endregion

        //검색
        private void FillGrid()
        {
            if (dgdCustom.Items.Count > 0)
            {
                dgdCustom.Items.Clear();
            }

            try
            {
                Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                sqlParameter.Add("sCustom", chkCustomSrh.IsChecked == true ? txtCustomSrh.Text : "");
                sqlParameter.Add("sTradeID", chkTradeSrh.IsChecked == true && cboTradeSrh.SelectedItem != null ? cboTradeSrh.SelectedValue.ToString() : "");
                sqlParameter.Add("iIncNotUse", chkNoUse.IsChecked == true ? 1 : 0);
                sqlParameter.Add("Chief", CheckBoxChiefSearch.IsChecked == true ? (TextBoxChiefSearch.Text == "" ? "" : TextBoxChiefSearch.Text) : "");

                DataSet ds = DataStore.Instance.ProcedureToDataSet_LogWrite("xp_Custom_sCustom", sqlParameter, true, "R");

                if (ds != null && ds.Tables.Count > 0)
                {
                    DataTable dt = ds.Tables[0];
                    int i = 0;

                    //dataGrid.Items.Clear();
                    if (dt.Rows.Count == 0)
                    {
                        MessageBox.Show("조회된 데이터가 없습니다.");
                    }
                    else
                    {
                        DataRowCollection drc = dt.Rows;

                        foreach (DataRow dr in drc)
                        {
                            i++;
                            var WinCustom = new Win_com_Custom_U_CodeView()
                            {
                                Num = i.ToString(),
                                CustomID = dr["CustomID"].ToString(),
                                ECustom = dr["ECustom"].ToString(),
                                KCustom = dr["KCustom"].ToString(),
                                ShortCustom = dr["ShortCustom"].ToString(),

                                CustomNo = dr["CustomNo"].ToString(),
                                Chief = dr["Chief"].ToString(),
                                Condition = dr["Condition"].ToString(),
                                Category = dr["Category"].ToString(),
                                ZipCode = dr["ZipCode"].ToString(),

                                OldNNewClss = dr["OldNNewClss"].ToString(),
                                GunMoolMngNo = dr["GunMoolMngNo"].ToString(),
                                Address1 = dr["Address1"].ToString(),
                                Address2 = dr["Address2"].ToString(),
                                AddressJiBun1 = dr["AddressJiBun1"].ToString(),

                                AddressJiBun2 = dr["AddressJiBun2"].ToString(),
                                AddressAssist = dr["AddressAssist"].ToString(),
                                Phone1 = dr["Phone1"].ToString(),
                                Phone2 = dr["Phone2"].ToString(),
                                FaxNo = dr["FaxNo"].ToString(),

                                EMail = dr["EMail"].ToString(),
                                HomePage = dr["HomePage"].ToString(),
                                Name = dr["Name"].ToString(),
                                Phone = dr["Phone"].ToString(),
                                TradeID = dr["TradeID"].ToString(),

                                UserID = dr["UserID"].ToString(),
                                UserPassWord = dr["UserPassWord"].ToString(),
                                LossClss = dr["LossClss"].ToString(),
                                SpendingClss = dr["SpendingClss"].ToString(),
                                WorkingClss = dr["WorkingClss"].ToString(),

                                CalClss = dr["CalClss"].ToString(),
                                PointClss = dr["PointClss"].ToString(),
                                UseClss = dr["UseClss"].ToString(),
                                Comments = dr["Comments"].ToString(),
                                RKCustom = dr["RKCustom"].ToString(),

                                RTradeID = dr["RTradeID"].ToString(),
                                TradgbnName = dr["TradgbnName"].ToString(),
                                CountryCode = dr["CountryCode"].ToString(),
                                Country = dr["Country"].ToString(),
                                FTAMgrYN = dr["FTAMgrYN"].ToString(),

                                ReqRemainQty = dr["ReqRemainQty"].ToString(),
                                DefectCount = dr["DefectCount"].ToString(),
                            };

                            tbkSearchIndex.Text = "▶ 검색결과 : " + i + " 건"; //2021-09-24 정현달 합계 추가

                            dgdCustom.Items.Add(WinCustom);
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

        //그리드 index select
        private void dgdCustom_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            var WinCustom = dgdCustom.SelectedItem as Win_com_Custom_U_CodeView;

            if (WinCustom != null)
            {
                this.DataContext = WinCustom;

                if (WinCustom.UseClss == "*")
                {
                    chkUseClss.IsChecked = true;
                }
                else
                {
                    chkUseClss.IsChecked = false;
                }
            }
        }

        //삭제체크
        private bool DeleteDataCheck(string strCustomID)
        {
            bool Flag = false;

            try
            {
                Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                sqlParameter.Clear();
                sqlParameter.Add("sCustomID", strCustomID);


                //outputResult = DataStore.Instance.ExecuteProcedureOutputNoTran("xp_Custom_dCustom_Check", sqlParameter, outputParam, false);
                string[] result = DataStore.Instance.ExecuteProcedure("xp_Custom_dCustom_Check", sqlParameter, false);

                if (result[0].Equals("success") && result[1].Equals(""))
                {
                    //MessageBox.Show("성공 *^^*");
                    Flag = true;
                }
                else
                {
                    MessageBox.Show(result[1]);
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

            return Flag;
        }
        //삭제
        private bool DeleteData(string strCustomID)
        {
            bool Flag = false;

            try
            {
                Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                sqlParameter.Clear();
                sqlParameter.Add("sCustomID", strCustomID);

                string[] result = DataStore.Instance.ExecuteProcedure_NewLog("xp_Custom_dCustom", sqlParameter, "D");

                if (result[0].Equals("success"))
                {
                    //MessageBox.Show("성공 *^^*");
                    Flag = true;
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

            return Flag;
        }

        private bool SaveData(string strCustomID, string strFlag)
        {
            bool flag = false;
            List<Procedure> Prolist = new List<Procedure>();
            List<Dictionary<string, object>> ListParameter = new List<Dictionary<string, object>>();

            try
            {
                if (CheckData())
                {

                    Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                    //sqlParameter.Add("sNewCustomID", strCustomID);
                    //sqlParameter.Add("sCustomID", strCustomID);
                    sqlParameter.Add("sCustomID", txtCustomID.Text !=null && !txtCustomID.Text.Trim().Equals("") ? txtCustomID.Text : "");
                    sqlParameter.Add("sKCustom", txtKCustom.Text);
                    sqlParameter.Add("sShortCustom", txtShortCustom.Text);
                    sqlParameter.Add("sECustom", txtECustom.Text);

                    sqlParameter.Add("sCustomNO", txtCustomNO.Text);
                    sqlParameter.Add("sChief", txtChief.Text);
                    sqlParameter.Add("sCondition", txtCondition.Text);
                    sqlParameter.Add("sCategory", txtCategory.Text);
                    sqlParameter.Add("sZipCode", txtZipCode.Text);

                    sqlParameter.Add("sOldNNewClss", rbnDoro.IsChecked == true ? "0" : "1");    //0 :도로명, 1: 지번
                    sqlParameter.Add("sGunMoolMngNo", WinCustom.GunMoolMngNo == null ? "" : WinCustom.GunMoolMngNo);
                    sqlParameter.Add("sAddress1", txtAddress1.Text);
                    sqlParameter.Add("sAddress2", txtAddress2.Text);
                    sqlParameter.Add("sAddressAssist", txtAddressAssist.Text);

                    sqlParameter.Add("sAddressJiBun1", txtAddressJiBun1.Text);
                    sqlParameter.Add("sAddressJiBun2", txtAddressJiBun2.Text);
                    sqlParameter.Add("sPhone1", txtPhone1.Text);
                    sqlParameter.Add("sPhone2", txtPhone2.Text);
                    sqlParameter.Add("sFaxNO", txtFaxNO.Text);

                    sqlParameter.Add("sEMail", txtEMail.Text);
                    sqlParameter.Add("sHomePage", txtHomepage.Text);
                    sqlParameter.Add("sName", txtName.Text);
                    sqlParameter.Add("sPhone", txtPhone.Text);
                    sqlParameter.Add("sTradeID", cboTrade.SelectedValue.ToString());

                    sqlParameter.Add("sUserID", txtUserID.Text);
                    sqlParameter.Add("sUserPassWord", txtUserPassword.Text);
                    sqlParameter.Add("sLossClss", WinCustom.LossClss == null ? "1" : WinCustom.LossClss);
                    sqlParameter.Add("sSpendingClss", WinCustom.SpendingClss == null ? "1" : WinCustom.SpendingClss);
                    sqlParameter.Add("sWorkingClss", WinCustom.WorkingClss == null ? "1" : WinCustom.WorkingClss);

                    sqlParameter.Add("sCalcClss", WinCustom.CalClss == null ? "1" : WinCustom.CalClss);
                    sqlParameter.Add("sPointClss", WinCustom.PointClss == null ? "1" : WinCustom.PointClss);
                    sqlParameter.Add("sComments", txtComments.Text);
                    sqlParameter.Add("sCountryCode", txtCountry.Tag != null ? txtCountry.Tag.ToString() : "");
                    sqlParameter.Add("sFTAMgrYN", cboFTAMgrYN.SelectedValue.ToString());

                    sqlParameter.Add("sReqRemainQty", ConvertInt(txtReqRemainQty.Text));
                    sqlParameter.Add("sDefectCount", ConvertInt(txtDefectCount.Text));

                    sqlParameter.Add("UseClss", chkUseClss.IsChecked == true ? "*" : "");

                    if (strFlag.Equals("I"))
                    {

                        sqlParameter.Add("sNewCustomID", strCustomID);
                        sqlParameter.Add("CreateUserID", MainWindow.CurrentUser);

                        Procedure pro1 = new Procedure();
                        pro1.Name = "xp_Custom_iCustom";
                        pro1.OutputUseYN = "Y";
                        pro1.OutputName = "sNewCustomID";
                        pro1.OutputLength = "5";

                        Prolist.Add(pro1);
                        ListParameter.Add(sqlParameter);

                        List<KeyValue> list_Result = new List<KeyValue>();
                        list_Result = DataStore.Instance.ExecuteAllProcedureOutputGetCS_NewLog(Prolist, ListParameter,"C");
                        string sGetID = string.Empty;

                        if (list_Result[0].key.ToLower() == "success")
                        {
                            list_Result.RemoveAt(0);
                            for (int i = 0; i < list_Result.Count; i++)
                            {
                                KeyValue kv = list_Result[i];
                                if (kv.key == "sNewCustomID")
                                {
                                    sGetID = kv.value;

                                    //InspectName = kv.value;
                                    AASS = kv.value;

                                    flag = true;
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
                        sqlParameter.Add("LastUpdateUserID", MainWindow.CurrentUser);

                        Procedure pro1 = new Procedure();
                        pro1.Name = "xp_Custom_uCustom";
                        pro1.OutputUseYN = "N";
                        pro1.OutputName = "sCustomID";
                        pro1.OutputLength = "5";

                        Prolist.Add(pro1);
                        ListParameter.Add(sqlParameter);

                        string[] Confirm = new string[2];
                        Confirm = DataStore.Instance.ExecuteAllProcedureOutputNew_NewLog(Prolist, ListParameter,"U");
                        if (Confirm[0] != "success")
                        {
                            MessageBox.Show("[저장실패]\r\n" + Confirm[1].ToString());
                            flag = false;
                            //return false;
                        }
                        else
                        {
                            flag = true;
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

            return flag;
        }

        //데이터 체크
        private bool CheckData()
        {
            bool flag = true;

            if (txtKCustom.Text.Equals("") || txtKCustom.Text.Length == 0)
            {
                MessageBox.Show("거래처가 입력되지 않았습니다.");
                flag = false;
                return flag;
            }

            //if (!GetCustomByUser(txtCustomID.Text, txtUserID.Text))
            //{
            //    MessageBox.Show("동일 사용자ID가 이미 존재합니다.");
            //    flag = false;
            //    return flag;
            //}

            // 거래구분 선택 안할시
            if (cboTrade.SelectedValue == null)
            {
                MessageBox.Show("거래 구분이 입력되지 않았습니다.");
                flag = false;
                return flag;
            }

            // FTP 중점관리 여부 확인
            if (cboFTAMgrYN.SelectedValue == null)
            {
                MessageBox.Show("FTP 중점관리 여부가 선택되지 않았습니다.");
                flag = false;
                return flag;
            }

            return flag;
        }

        //동일 사용자 ID기입 방지
        private bool GetCustomByUser(string strCustomID, string strUserID)
        {
            bool flag = true;

            try
            {
                Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                sqlParameter.Clear();
                sqlParameter.Add("sCustomID", strCustomID);
                sqlParameter.Add("sUserID", strUserID);

                DataSet ds = DataStore.Instance.ProcedureToDataSet("xp_Custom_sCustombyUserID", sqlParameter, false);

                if (ds != null && ds.Tables.Count > 0)
                {
                    DataTable dt = ds.Tables[0];

                    //dataGrid.Items.Clear();
                    if (dt.Rows.Count > 0)
                    {
                        int Cnt = int.Parse(dt.Rows[0]["Cnt"].ToString());

                        if (Cnt > 0)
                        {
                            flag = false;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                flag = false;
                MessageBox.Show("오류 발생, 오류 내용 : " + ex.ToString());
            }
            finally
            {
                DataStore.Instance.CloseConnection();
            }

            return flag;
        }

        private void txtCountry_KeyDown(object sender, System.Windows.Input.KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                e.Handled = true;
                MainWindow.pf.ReturnCode(txtCountry, (int)Defind_CodeFind.DCF_COUNTRY, "");
            }
        }

        private void btnCountryPf_Click(object sender, RoutedEventArgs e)
        {
            MainWindow.pf.ReturnCode(txtCountry, (int)Defind_CodeFind.DCF_COUNTRY, "");
        }

        private void txtZipCode_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                PopUp.Win_Zip_Address ZipPopUp = new PopUp.Win_Zip_Address();
                ZipPopUp.ShowDialog();

                if (ZipPopUp.DialogResult == true)
                {
                    if (ZipPopUp.strGubun.Equals("0"))
                    {
                        txtAddress1.Text = ZipPopUp.Juso;
                        txtAddress2.Text = ZipPopUp.Detail1;
                        txtAddressAssist.Text = ZipPopUp.Detail2;
                        txtZipCode.Text = ZipPopUp.ZipCode;
                    }
                    else if (ZipPopUp.strGubun.Equals("1"))
                    {
                        txtAddressJiBun1.Text = ZipPopUp.Juso;
                        txtZipCode.Text = ZipPopUp.ZipCode;
                    }
                }
            }
        }

        private void btnAddress_Click(object sender, RoutedEventArgs e)
        {
            PopUp.Win_Zip_Address ZipPopUp = new PopUp.Win_Zip_Address();
            ZipPopUp.ShowDialog();

            if (ZipPopUp.DialogResult == true)
            {
                if (ZipPopUp.strGubun.Equals("0"))
                {
                    txtAddress1.Text = ZipPopUp.Juso;
                    txtAddress2.Text = ZipPopUp.Detail1;
                    txtAddressAssist.Text = ZipPopUp.Detail2;
                    txtZipCode.Text = ZipPopUp.ZipCode;
                }
                else if (ZipPopUp.strGubun.Equals("1"))
                {
                    txtAddressJiBun1.Text = ZipPopUp.Juso;
                    txtZipCode.Text = ZipPopUp.ZipCode;
                }
            }
        }

        #region 텍스트박스 엔터 → 다음 텍스트 박스

        private void txtKCustom_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == System.Windows.Input.Key.Enter)
            {
                txtCountry.Focus();
            }
        }

        private void txtCountry_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            txtChief.Focus();
        }

        private void txtShortCustom_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == System.Windows.Input.Key.Enter)
            {
                cboTrade.Focus();
            }
        }

        private void txtECustom_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == System.Windows.Input.Key.Enter)
            {
                txtCondition.Focus();
            }
        }

        private void txtCondition_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == System.Windows.Input.Key.Enter)
            {
                txtCategory.Focus();
            }
        }

        private void txtCategory_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == System.Windows.Input.Key.Enter)
            {
                txtPhone1.Focus();
            }
        }

        private void txtPhone1_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == System.Windows.Input.Key.Enter)
            {
                txtPhone2.Focus();
            }
        }

        private void txtPhone2_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == System.Windows.Input.Key.Enter)
            {
                txtFaxNO.Focus();
            }
        }

        private void txtFaxNO_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == System.Windows.Input.Key.Enter)
            {
                txtName.Focus();
            }
        }

        private void txtChief_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == System.Windows.Input.Key.Enter)
            {
                txtCustomNO.Focus();
            }
        }

        private void txtCustomNO_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == System.Windows.Input.Key.Enter)
            {
                txtShortCustom.Focus();
            }
        }

        private void txtName_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == System.Windows.Input.Key.Enter)
            {
                txtPhone.Focus();
            }
        }

        private void txtPhone_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == System.Windows.Input.Key.Enter)
            {
                txtUserID.Focus();
            }
        }

        private void txtUserID_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == System.Windows.Input.Key.Enter)
            {
                txtUserPassword.Focus();
            }
        }

        private void txtUserPassword_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == System.Windows.Input.Key.Enter)
            {
                txtZipCode.Focus();
            }
        }

        private void txtAddressAssist_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == System.Windows.Input.Key.Enter)
            {
                txtHomepage.Focus();
            }
        }

        private void txtAddressJiBun2_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == System.Windows.Input.Key.Enter)
            {
                txtHomepage.Focus();
            }
        }

        private void txtHomepage_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == System.Windows.Input.Key.Enter)
            {
                txtEMail.Focus();
            }
        }

        private void txtEMail_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == System.Windows.Input.Key.Enter)
            {
                txtComments.Focus();
            }
        }

        private void txtComments_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == System.Windows.Input.Key.Enter)
            {
                btnSave.Focus();
            }
        }

        private void cboTrade_DropDownClosed(object sender, EventArgs e)
        {
            txtECustom.Focus();
        }




        #endregion // 텍스트박스 엔터 → 다음 텍스트 박스

        private void lblNoUse_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            if (chkNoUse.IsChecked == true)
            {
                chkNoUse.IsChecked = false;
            }
            else
            {
                chkNoUse.IsChecked = true;
            }
        }

        //대표자 라벨 마우스 왼쪽버튼 클릭
        private void LabelChiefSearch_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            if (CheckBoxChiefSearch.IsChecked == true)
            {
                CheckBoxChiefSearch.IsChecked = false;
                TextBoxChiefSearch.IsEnabled = false;
            }
            else
            {
                CheckBoxChiefSearch.IsChecked = true;
                TextBoxChiefSearch.IsEnabled = true;
            }
        }

        //대표자 체크박스 체크
        private void CheckBoxChiefSearch_Checked(object sender, RoutedEventArgs e)
        {
            CheckBoxChiefSearch.IsChecked = true;
            TextBoxChiefSearch.IsEnabled = true;
        }

        //대표자 체크박스 체크해제
        private void CheckBoxChiefSearch_Unchecked(object sender, RoutedEventArgs e)
        {
            CheckBoxChiefSearch.IsChecked = false;
            TextBoxChiefSearch.IsEnabled = false;
        }

        //상호검색 텍스트박스 키다운 이벤트
        private void txtCustomSrh_KeyDown(object sender, KeyEventArgs e)
        {
            try
            {
                if (e.Key == Key.Enter)
                {
                    pf.ReturnCode(txtCustomSrh, 86, txtCustomSrh.Text);
                }
            }
            catch (Exception ee)
            {
                MessageBox.Show("오류지점 - 상호검색 엔터키 : " + ee.ToString());
            }
        }

        //대표자 텍스트박스 키다운 이벤트
        private void TextBoxChiefSearch_KeyDown(object sender, KeyEventArgs e)
        {
            try
            {
                if (e.Key == Key.Enter)
                {
                    pf.ReturnCode(TextBoxChiefSearch, 87, TextBoxChiefSearch.Text);
                }
            }
            catch (Exception ee)
            {
                MessageBox.Show("오류지점 - 대표자 엔터키 : " + ee.ToString());
            }
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

        private void ClearData()
        {
            // 오른쪽 상세 내용 비우기
            txtCustomID.Clear();
            txtCountry.Clear();
            txtKCustom.Clear();
            txtShortCustom.Clear();
            txtECustom.Clear();
            txtCondition.Clear();
            txtCategory.Clear();
            txtPhone1.Clear();
            txtPhone2.Clear();
            txtFaxNO.Clear();
            txtChief.Clear();
            txtCustomNO.Clear();
            cboTrade.Text = ""; // 거래구분 콤보박스
            cboFTAMgrYN.Text = ""; // FTA 중점관리 콤보박스
            txtName.Clear(); // 담당자 정보
            txtPhone.Clear();
            txtUserID.Clear(); // 로그인 정보
            txtUserPassword.Clear();
            txtZipCode.Text = ""; // 주소
            txtAddress1.Text = "";
            txtAddress2.Text = "";
            txtAddressAssist.Text = "";
            txtAddressJiBun1.Text = ""; // 지번
            txtAddressJiBun2.Text = "";
            txtHomepage.Text = ""; // 하단 홈페이지, 이메일, 비고
            txtEMail.Text = "";
            txtComments.Text = "";

            chkUseClss.IsChecked = false; // 사용안함 체크박스

            this.DataContext = null;
        }
    }

    class Win_com_Custom_U_CodeView : BaseView
    {
        public override string ToString()
        {
            return (this.ReportAllProperties());
        }

        public string Num { get; set; }
        public string CustomID { get; set; }
        public string ECustom { get; set; }
        public string KCustom { get; set; }
        public string ShortCustom { get; set; }

        public string CustomNo { get; set; }
        public string Chief { get; set; }
        public string Condition { get; set; }
        public string Category { get; set; }
        public string ZipCode { get; set; }

        public string OldNNewClss { get; set; }
        public string GunMoolMngNo { get; set; }
        public string Address1 { get; set; }
        public string Address2 { get; set; }
        public string AddressJiBun1 { get; set; }

        public string AddressJiBun2 { get; set; }
        public string AddressAssist { get; set; }
        public string Phone1 { get; set; }
        public string Phone2 { get; set; }
        public string FaxNo { get; set; }

        public string EMail { get; set; }
        public string HomePage { get; set; }
        public string Name { get; set; }
        public string Phone { get; set; }
        public string TradeID { get; set; }

        public string UserID { get; set; }
        public string UserPassWord { get; set; }
        public string LossClss { get; set; }
        public string SpendingClss { get; set; }
        public string WorkingClss { get; set; }

        public string CalClss { get; set; }
        public string PointClss { get; set; }
        public string UseClss { get; set; }
        public string Comments { get; set; }
        public string RKCustom { get; set; }

        public string RTradeID { get; set; }
        public string TradgbnName { get; set; }
        public string CountryCode { get; set; }
        public string Country { get; set; }
        public string FTAMgrYN { get; set; }

        public string ReqRemainQty { get; set; }
        public string DefectCount { get; set; }


        public string Text { get; set; }
    }
}

using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Data;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using WizMes_WellMade.PopUP;

namespace WizMes_WellMade
{
    /// <summary>
    /// Win_com_CustomArticle_U.xaml에 대한 상호 작용 논리
    /// </summary>
    public partial class Win_com_CustomArticle_U : UserControl
    {
        string stDate = string.Empty;
        string stTime = string.Empty;

        string strFlag = string.Empty;
        int rowNum = 0;

        ObservableCollection<Win_com_CustomArticle_U_CodeView_SorD> CustomArticle = new ObservableCollection<Win_com_CustomArticle_U_CodeView_SorD>();

        Win_com_CustomArticle_U_CodeView winCustomArticle = new Win_com_CustomArticle_U_CodeView();
        List<Win_com_CustomArticle_U_CodeView_SorD> lstAll = new List<Win_com_CustomArticle_U_CodeView_SorD>();
        List<Win_com_CustomArticle_U_CodeView_SorD> lstSelect = new List<Win_com_CustomArticle_U_CodeView_SorD>();

        PlusFinder pf = new PlusFinder();
        int indexAllItem = 0; // 전체품목 검색을 위한 인덱스
        int indexSelItem = 0; // 선택된 품목 검색을 위한 인덱스


        public Win_com_CustomArticle_U()
        {
            InitializeComponent();
        }

        private void UserControl_Loaded(object sender, RoutedEventArgs e)
        {
            stDate = DateTime.Now.ToString("yyyyMMdd");
            stTime = DateTime.Now.ToString("HHmm");

            DataStore.Instance.InsertLogByFormS(this.GetType().Name, stDate, stTime, "S");

            Lib.Instance.UiLoading(sender);
            //cboInGubunSrh.SelectedIndex = 0;
            //btnSearch_Click(sender, e);
            SetComboBox();
            SetAllItemList();
            FillGrid();


            Article_ID.Visibility = Visibility.Hidden;
            //Article.Visibility = Visibility.Hidden;
            BuyerArticleNo.Visibility = Visibility.Hidden;
        }

        //콤보박스
        private void SetComboBox()
        {
            //List<string[]> listGubun = new List<string[]>();
            //string[] YN01 = new string[] { "1", "매출" };
            //string[] YN02 = new string[] { "2", "매입" };
            //listGubun.Add(YN01);
            //listGubun.Add(YN02);

            //ObservableCollection<CodeView> ovcYN = ComboBoxUtil.Instance.Direct_SetComboBox(listGubun);
            //this.cboCustomGubunSrh.ItemsSource = ovcYN;
            //this.cboCustomGubunSrh.DisplayMemberPath = "code_name";
            //this.cboCustomGubunSrh.SelectedValuePath = "code_id";

            //상단 매입/매출처구분
            ObservableCollection<CodeView> ovcTrade = ComboBoxUtil.Instance.Gf_DB_CM_GetComCodeDataset(null, "CMMTRAD", "Y", "");
            this.cboCustomGubunSrh.ItemsSource = ovcTrade;
            this.cboCustomGubunSrh.DisplayMemberPath = "code_name";
            this.cboCustomGubunSrh.SelectedValuePath = "code_id";
            this.cboCustomGubunSrh.SelectedIndex = 0;

            //입력칸 매입/매출처구분
            this.cboInGubunSrh.ItemsSource = ovcTrade;
            this.cboInGubunSrh.DisplayMemberPath = "code_name";
            this.cboInGubunSrh.SelectedValuePath = "code_id";
            this.cboInGubunSrh.SelectedIndex = 0;

        }
        //취소, 저장 후
        private void CanBtnControl()
        {
            Lib.Instance.UiButtonEnableChange_IUControl(this);
            btnAddSelectItem.IsEnabled = false;
            btnDelSelectItem.IsEnabled = false;
            dgdCustomArticle.IsHitTestVisible = true;
            txtCustom.IsEnabled = false;
            btnPfCustom.IsEnabled = false;
            //txtArticle.IsEnabled = false;
            //btnPfArticle.IsEnabled = false;

            chkCustomSrh.IsEnabled = true;
            if (chkCustomSrh.IsChecked == true)
                txtCustomSrh.IsEnabled = true;
        }

        //추가, 수정 클릭 시
        private void CantBtnControl()
        {
            Lib.Instance.UiButtonEnableChange_SCControl(this);
            btnAddSelectItem.IsEnabled = true;
            btnDelSelectItem.IsEnabled = true;
            dgdCustomArticle.IsHitTestVisible = false;
            txtCustom.IsEnabled = true;
            btnPfCustom.IsEnabled = true;
            //txtArticle.IsEnabled = true;
            //btnPfArticle.IsEnabled = true;


            chkCustomSrh.IsEnabled = false;
            txtCustomSrh.IsEnabled = false;


        }

        // 거래처명 라벨 버튼 클릭 이벤트
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
        // 거래처명 체크박스
        private void chkCustomSrh_Checked(object sender, RoutedEventArgs e)
        {
            chkCustomSrh.IsChecked = true;
            txtCustomSrh.IsEnabled = true;
            cboCustomGubunSrh.IsEnabled = true;
        }
        // 거래처명 체크박스
        private void chkCustomSrh_UnChecked(object sender, RoutedEventArgs e)
        {
            chkCustomSrh.IsChecked = false;
            txtCustomSrh.IsEnabled = false;
            cboCustomGubunSrh.IsEnabled = false;
        }

        // 관리품목 보기 라벨 클릭 이벤트
        private void bdrShowDetail_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            if (chkShowDetail.IsChecked == true)
            {
                //Article_ID.Visibility = Visibility.Visible;
                //Article.Visibility = Visibility.Visible;

                chkShowDetail.IsChecked = false;
            }
            else
            {
                chkShowDetail.IsChecked = true;

                //Article_ID.Visibility = Visibility.Hidden;
                //Article.Visibility = Visibility.Hidden;
            }
        }

        //추가
        private void btnAdd_Click(object sender, RoutedEventArgs e)
        {
            if (dgdCustomArticle.SelectedItem != null)
            {
                rowNum = dgdCustomArticle.SelectedIndex;
            }

            if (lstSelect.Count > 0)
            {
                lstSelect.Clear();
            }

            if (dgdSelectItem.Items.Count > 0)
            {
                dgdSelectItem.Items.Clear();
            }

            FillGridAllItem();

            CantBtnControl();
            tbkMsg.Text = "자료 입력 중";
            strFlag = "I";
            this.DataContext = null;

            txtCustom.Focus();
            cboInGubunSrh.SelectedIndex = 0;

        }

        //수정
        private void btnUpdate_Click(object sender, RoutedEventArgs e)
        {
            winCustomArticle = dgdCustomArticle.SelectedItem as Win_com_CustomArticle_U_CodeView;

            if (winCustomArticle != null)
            {
                rowNum = dgdCustomArticle.SelectedIndex;
                CantBtnControl();
                tbkMsg.Text = "자료 수정 중";
                strFlag = "U";
            }
            else
            {
                MessageBox.Show("수정할 자료를 선택하고 눌러주십시오.");
            }

            txtCustom.Focus();
        }

        //삭제
        private void btnDelete_Click(object sender, RoutedEventArgs e)
        {
            winCustomArticle = dgdCustomArticle.SelectedItem as Win_com_CustomArticle_U_CodeView;

            if (winCustomArticle == null)
            {
                MessageBox.Show("삭제할 데이터가 지정되지 않았습니다. 삭제데이터를 지정하고 눌러주세요");
            }
            else
            {
                if (MessageBox.Show("선택하신 거래처의 선택품목을 모두 삭제하시겠습니까?", "삭제 전 확인", MessageBoxButton.YesNo) == MessageBoxResult.Yes)
                {
                    DataStore.Instance.InsertLogByForm(this.GetType().Name, "D");
                    if (dgdCustomArticle.Items.Count > 0 && dgdCustomArticle.SelectedItem != null)
                    {
                        rowNum = dgdCustomArticle.SelectedIndex;
                    }

                    if (Procedure.Instance.DeleteData(winCustomArticle.CustomID, "sCustomID", "xp_Custom_dCustomArticleAll"))
                    {
                        rowNum -= 1;
                        re_Search(rowNum);
                    }
                    //else
                    //{
                    //    MessageBox.Show("선택하신 거래처의 선택품목삭제에 실패하였습니다.");
                    //}
                }
            }
        }

        //닫기
        private void btnClose_Click(object sender, RoutedEventArgs e)
        {
            DataStore.Instance.InsertLogByFormS(this.GetType().Name, stDate, stTime, "E");
            Lib.Instance.ChildMenuClose(this.ToString());
        }

        //조회
        private void btnSearch_Click(object sender, RoutedEventArgs e)
        {
            txtArticleSrhSelectItem.Text = "";
            txtArticleSrhAllItem.Text = "";

            //if(dgdCustomArticle.Items.Count < 0)
            //{
            //    FillGrid();
            //}
            //else
            //{
            rowNum = 0;
            re_Search(rowNum);
            //}




        }

        //저장
        private void btnSave_Click(object sender, RoutedEventArgs e)
        {
            if (txtCustom.Tag == null)
            {
                MessageBox.Show("거래처 정보가 입력되지 않았습니다.");
                return;
            }

            if (SaveData(txtCustom.Tag.ToString(), strFlag))
            {
                CanBtnControl();

                re_Search(rowNum);
            }
        }

        ////저장
        //private void btnSave_Click(object sender, RoutedEventArgs e)
        //{
        //    if (SaveData(strFlag, txtMcRInspectID.Text))
        //    {
        //        CanBtnControl();
        //        strBasisID = string.Empty;
        //        lblMsg.Visibility = Visibility.Hidden;

        //        if (strFlag.Equals("I"))
        //        {
        //            InspectDate = dtpInspectDate.SelectedDate.ToString().Substring(0, 10);
        //            InspectName = txtMc.Tag.ToString();

        //            rowNum = 0;
        //            re_Search(rowNum);
        //        }
        //        else
        //        {
        //            rowNum = dgdInspect.SelectedIndex;
        //        }
        //    }

        //    int i = 0;

        //    foreach (Win_com_CustomArticle_U_CodeView_SorD WMRIC in dgdInspect.Items)
        //    {

        //        string a = WMRIC.McRInspectDate_Convert.ToString();
        //        string b = WMRIC.MCID.ToString();


        //        System.Diagnostics.Debug.WriteLine("a 컬럼은 무슨 데이터?? === " + a);

        //        if (a == InspectDate && b == InspectName)
        //        {
        //            System.Diagnostics.Debug.WriteLine("데이터 같음");

        //            break;
        //        }
        //        else
        //        {
        //            System.Diagnostics.Debug.WriteLine("다름");
        //        }

        //        i++;
        //    }

        //    rowNum = i;
        //    re_Search(rowNum);

        //}





        //취소
        private void btnCancel_Click(object sender, RoutedEventArgs e)
        {
            CanBtnControl();
            strFlag = string.Empty;
            re_Search(rowNum);
        }

        //엑셀
        private void btnExcel_Click(object sender, RoutedEventArgs e)
        {
            DataTable dt = null;
            string Name = string.Empty;

            string[] lst = new string[6];
            lst[0] = "거래처별 관리품목";
            lst[1] = "전체 품목";
            lst[2] = "선택된 품목";
            lst[3] = dgdCustomArticle.Name;
            lst[4] = dgdAllItem.Name;
            lst[5] = dgdSelectItem.Name;

            ExportExcelxaml ExpExc = new ExportExcelxaml(lst);

            ExpExc.ShowDialog();

            if (ExpExc.DialogResult.HasValue)
            {
                if (ExpExc.choice.Equals(dgdCustomArticle.Name))
                {
                    DataStore.Instance.InsertLogByForm(this.GetType().Name, "E");
                    if (ExpExc.Check.Equals("Y"))
                        dt = Lib.Instance.DataGridToDTinHidden(dgdCustomArticle);
                    else
                        dt = Lib.Instance.DataGirdToDataTable(dgdCustomArticle);

                    Name = dgdCustomArticle.Name;

                    if (Lib.Instance.GenerateExcel(dt, Name))
                        Lib.Instance.excel.Visible = true;
                    else
                        return;
                }
                else if (ExpExc.choice.Equals(dgdAllItem.Name))
                {
                    if (ExpExc.Check.Equals("Y"))
                        dt = Lib.Instance.DataGridToDTinHidden(dgdAllItem);
                    else
                        dt = Lib.Instance.DataGirdToDataTable(dgdAllItem);
                    Name = dgdAllItem.Name;

                    if (Lib.Instance.GenerateExcel(dt, Name))
                        Lib.Instance.excel.Visible = true;
                    else
                        return;
                }
                else if (ExpExc.choice.Equals(dgdSelectItem.Name))
                {
                    if (ExpExc.Check.Equals("Y"))
                        dt = Lib.Instance.DataGridToDTinHidden(dgdSelectItem);
                    else
                        dt = Lib.Instance.DataGirdToDataTable(dgdSelectItem);
                    Name = dgdSelectItem.Name;

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
        private void re_Search(int selectedIndex)
        {
            try
            {
                FillGrid();

                if (dgdCustomArticle.Items.Count > 0)
                {
                    dgdCustomArticle.SelectedIndex = selectedIndex;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("오류 발생, 오류 내용 : " + ex.ToString());
            }
        }

        //실조회
        private void FillGrid()
        {
            if (dgdCustomArticle.Items.Count > 0)
            {
                dgdCustomArticle.Items.Clear();
            }

            try
            {
                int numCustom = chkCustomSrh.IsChecked == true ? 1 : 0;
                string strCustom = chkCustomSrh.IsChecked == true ? txtCustomSrh.Text : "";
                string strCustomGubun = chkCustomSrh.IsChecked == true && cboCustomGubunSrh.SelectedItem != null ? cboCustomGubunSrh.SelectedValue.ToString() : "";
                string strYN = chkShowDetail.IsChecked == true ? "Y" : "N";

                int numBuyCustom = chkBuyCustomSrh.IsChecked == true ? 1 : 0;
                string strBuyCustom = chkBuyCustomSrh.IsChecked == true ? txtBuyCustomSrh.Text : "";

                int numArticleID = chkArticleSrh.IsChecked == true ? 1 : 0;
                string strArticleID = chkArticleSrh.IsChecked == true ? txtArticleSrh.Tag.ToString() : "";

                DataTable dt = Procedure.Instance.GetCustomArticle(numCustom, strCustom, strCustomGubun, numArticleID, strArticleID, strYN, numBuyCustom, strBuyCustom);


                DataStore.Instance.InsertLogByForm(this.GetType().Name, "R");

                if (dt.Rows.Count > 0)
                {
                    int i = 0;
                    DataRowCollection drc = dt.Rows;

                    foreach (DataRow dr in drc)
                    {
                        i++;
                        var dgdCustomArticleInfo = new Win_com_CustomArticle_U_CodeView()
                        {
                            Num = i.ToString(),
                            CustomID = dr["CustomID"].ToString(),
                            KCustom = dr["KCustom"].ToString(),
                            ArticleID = dr["ArticleID"].ToString(),
                            BuyerArticleNo = dr["BuyerArticleNo"].ToString(),
                            TradeID = dr["TradeID"].ToString(),
                            Article = dr["Article"].ToString()
                        };

                        dgdCustomArticle.Items.Add(dgdCustomArticleInfo);
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

        //Row 선택 시
        private void dgdCustomArticle_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            winCustomArticle = dgdCustomArticle.SelectedItem as Win_com_CustomArticle_U_CodeView;

            if (winCustomArticle != null)
            {
                txtArticleSrhSelectItem.Text = "";
                txtArticleSrhAllItem.Text = "";

                this.DataContext = winCustomArticle;
                FillGridSelectItem(winCustomArticle.CustomID);
                setLstAllToAllArticle();
                FillGridAllItem();
            }
        }

        #region 데이터그리드 선택시에 전체품목은 초기화시켜주기 위해서 → lstAll 을 선택된 품목을 제외한 품목으로 채우기
        private void setLstAllToAllArticle()
        {
            lstAll.Clear();

            try
            {
                Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                sqlParameter.Add("sArticle", "");
                sqlParameter.Add("iIncNotUse", 0);
                sqlParameter.Add("sArticleGrpID", "");
                sqlParameter.Add("sSupplyType", "");
                sqlParameter.Add("nBuyerArticleNo", 1);
                sqlParameter.Add("BuyerArticleNo", ""); //승인여부

                #region 매출만
                if (cboInGubunSrh.SelectedValue.ToString().Equals("1")) // 매출
                {
                    DataSet ds = DataStore.Instance.ProcedureToDataSet("xp_Article_sArticle_Custom_buy", sqlParameter, true);

                    if (ds != null && ds.Tables.Count > 0)
                    {
                        DataTable dt = ds.Tables[0];

                        if (dt.Rows.Count > 0)
                        {
                            int i = 0;
                            DataRowCollection drc = dt.Rows;

                            foreach (DataRow dr in drc)
                            {
                                bool good = true;
                                for (int k = 0; k < lstSelect.Count; k++)
                                {
                                    var Compare = lstSelect[k] as Win_com_CustomArticle_U_CodeView_SorD;
                                    if (Compare != null)
                                    {
                                        if (Compare.ArticleID.Trim().Equals(dr["ArticleID"].ToString().Trim()))
                                        {
                                            good = false;
                                            break;
                                        }
                                    }
                                }

                                if (good == true)
                                {
                                    i++;
                                    var Article = new Win_com_CustomArticle_U_CodeView_SorD()
                                    {
                                        AllNum = i.ToString(),
                                        ArticleID = dr["ArticleID"].ToString(),
                                        Article = dr["Article"].ToString(),
                                        BuyerArticleNo = dr["BuyerArticleNo"].ToString(),
                                        chkFlag = false
                                    };

                                    lstAll.Add(Article);
                                }
                            }
                        }
                    }
                }
                #endregion 매출만
                #region 매입만
                else if (cboInGubunSrh.SelectedValue.ToString().Equals("2") //매입
                    || cboInGubunSrh.SelectedValue.ToString().Equals("3")) //매입+매출만
                {
                    DataSet ds = DataStore.Instance.ProcedureToDataSet("xp_Article_sArticle_Custom_sales", sqlParameter, true);

                    if (ds != null && ds.Tables.Count > 0)
                    {
                        DataTable dt = ds.Tables[0];

                        if (dt.Rows.Count > 0)
                        {
                            int i = 0;
                            DataRowCollection drc = dt.Rows;

                            foreach (DataRow dr in drc)
                            {
                                bool good = true;
                                for (int k = 0; k < lstSelect.Count; k++)
                                {
                                    var Compare = lstSelect[k] as Win_com_CustomArticle_U_CodeView_SorD;
                                    if (Compare != null)
                                    {
                                        if (Compare.ArticleID.Trim().Equals(dr["ArticleID"].ToString().Trim()))
                                        {
                                            good = false;
                                            break;
                                        }
                                    }
                                }

                                if (good == true)
                                {
                                    i++;
                                    var Article = new Win_com_CustomArticle_U_CodeView_SorD()
                                    {
                                        AllNum = i.ToString(),
                                        ArticleID = dr["ArticleID"].ToString(),
                                        Article = dr["Article"].ToString(),
                                        BuyerArticleNo = dr["BuyerArticleNo"].ToString(),
                                        chkFlag = false
                                    };

                                    lstAll.Add(Article);
                                }
                            }
                        }
                    }
                }
                #endregion 매입만

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

        private void FillGridSelectItem(string strCustomID)
        {
            if (dgdSelectItem.Items.Count > 0) { dgdSelectItem.Items.Clear(); }
            if (lstSelect.Count > 0) { lstSelect.Clear(); }

            try
            {
                DataTable dt = Procedure.Instance.GetCustomArticleSelection(strCustomID);

                if (dt.Rows.Count > 0)
                {
                    int i = 0;
                    DataRowCollection drc = dt.Rows;

                    foreach (DataRow dr in drc)
                    {
                        i++;
                        var dgdItemInfo = new Win_com_CustomArticle_U_CodeView_SorD()
                        {
                            SelectNum = i.ToString(),
                            ArticleID = dr["ArticleID"].ToString(),
                            Article = dr["Article"].ToString(),
                            BuyerArticleNo = dr["BuyerArticleNo"].ToString(),
                            CustomBuyArticle = dr["CustomBuyArticle"].ToString(),
                            FinalCustom = dr["FinalCustom"].ToString(),
                            UnitPrice = stringFormatN0(dr["UnitPrice"]),
                            chkFlag = false
                        };

                        lstSelect.Add(dgdItemInfo);
                    }

                    for (int j = 0; lstSelect.Count > j; j++)
                    {
                        dgdSelectItem.Items.Add(lstSelect[j]);
                    }

                    tbkSelectCount.Text = "선택품목 : " + dgdSelectItem.Items.Count.ToString() + "개";
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

        //
        private void SetAllItemList()
        {
            try
            {
                DataTable dt = Procedure.Instance.GetArticle("", 0, "", "");

                if (dt.Rows.Count > 0)
                {
                    int i = 0;
                    DataRowCollection drc = dt.Rows;

                    foreach (DataRow dr in drc)
                    {
                        i++;
                        var dgdItemInfo = new Win_com_CustomArticle_U_CodeView_SorD()
                        {
                            AllNum = i.ToString(),
                            ArticleID = dr["ArticleID"].ToString(),
                            Article = dr["Article"].ToString(),
                            BuyerArticleNo = dr["BuyerArticleNo"].ToString(),
                            chkFlag = false
                        };

                        lstAll.Add(dgdItemInfo);
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

        //모든품목 표 
        private void FillGridAllItem()
        {
            if (dgdAllItem.Items.Count > 0) { dgdAllItem.Items.Clear(); }

            try
            {
                for (int j = 0; j < lstAll.Count; j++)
                {
                    bool flag = true;
                    var ItemAll = lstAll[j] as Win_com_CustomArticle_U_CodeView_SorD;

                    for (int k = 0; k < lstSelect.Count; k++)
                    {
                        var ItemSelect = lstSelect[k] as Win_com_CustomArticle_U_CodeView_SorD;
                        if (ItemAll.ArticleID.Equals(ItemSelect.ArticleID))
                        {
                            flag = false;
                            break;
                        }
                    }

                    if (flag)
                    {
                        dgdAllItem.Items.Add(ItemAll);

                    }
                }

                tbkAllCount.Text = "전체 품목 : " + dgdAllItem.Items.Count.ToString() + "개";
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

        //저장

        private bool SaveData(string strID, string strFlag)
        {
            bool flag = false;
            List<Procedure> Prolist = new List<Procedure>();
            List<Dictionary<string, object>> ListParameter = new List<Dictionary<string, object>>();

            try
            {
                if (CheckData())
                {
                    Dictionary<string, object> sqlParameter = null;

                    #region 추가

                    if (strFlag.Equals("I"))
                    {
                        for (int i = 0; i < lstSelect.Count; i++)
                        {
                            var SelectItem = lstSelect[i] as Win_com_CustomArticle_U_CodeView_SorD;
                            sqlParameter = new Dictionary<string, object>();
                            sqlParameter.Clear();
                            sqlParameter.Add("sCustomID", strID);
                            sqlParameter.Add("sArticleID", SelectItem.ArticleID);
                            sqlParameter.Add("sFinalCustomID", SelectItem.FinalCustomID);
                            sqlParameter.Add("sCustomBuyArticle", SelectItem.CustomBuyArticle);
                            sqlParameter.Add("UnitPrice", ConvertDouble(SelectItem.UnitPrice));
                            sqlParameter.Add("CreateUserID", MainWindow.CurrentUser);

                            Procedure pro1 = new Procedure();
                            pro1.Name = "xp_Custom_iCustomArticle";
                            pro1.OutputUseYN = "N";
                            pro1.OutputName = "sCustomID";
                            pro1.OutputLength = "10";

                            Prolist.Add(pro1);
                            ListParameter.Add(sqlParameter);
                        }

                        string[] Confirm = new string[2];
                        Confirm = DataStore.Instance.ExecuteAllProcedureOutputNew_NewLog(Prolist, ListParameter, "C");
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

                    #endregion

                    #region 수정

                    else if (strFlag.Equals("U"))
                    {
                        if (Procedure.Instance.DeleteData(strID, "sCustomID", "xp_Custom_dCustomArticleAll"))
                        {
                            for (int i = 0; i < lstSelect.Count; i++)
                            {
                                var SelectItem = lstSelect[i] as Win_com_CustomArticle_U_CodeView_SorD;
                                sqlParameter = new Dictionary<string, object>();
                                sqlParameter.Clear();
                                sqlParameter.Add("sCustomID", strID);
                                sqlParameter.Add("sArticleID", SelectItem.ArticleID);
                                sqlParameter.Add("sFinalCustomID", SelectItem.FinalCustomID);
                                sqlParameter.Add("sCustomBuyArticle", SelectItem.CustomBuyArticle);
                                sqlParameter.Add("UnitPrice", ConvertDouble(SelectItem.UnitPrice));
                                sqlParameter.Add("CreateUserID", MainWindow.CurrentUser);

                                Procedure pro1 = new Procedure();
                                pro1.Name = "xp_Custom_iCustomArticle";
                                pro1.OutputUseYN = "N";
                                pro1.OutputName = "sCustomID";
                                pro1.OutputLength = "10";

                                Prolist.Add(pro1);
                                ListParameter.Add(sqlParameter);
                            }

                            string[] Confirm = new string[2];
                            Confirm = DataStore.Instance.ExecuteAllProcedureOutputNew_NewLog(Prolist, ListParameter, "U");
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

                    #endregion
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("오류 발생, 오류 내용 : " + ex.ToString());
            }
            finally
            {
                DataStore.Instance.CloseConnection(); //2021-09-13 현달씨 DBClose
            }

            return flag;
        }

        /// <summary>
        /// 입력사항 체크
        /// </summary>
        /// <returns></returns>
        private bool CheckData()
        {
            bool flag = true;

            try
            {
                if (txtCustom.Text.Length <= 0 || txtCustom.Tag.ToString().Equals(""))
                {
                    MessageBox.Show("거래처가 입력되지 않았습니다.");
                    flag = false;
                    return flag;
                }

                if (lstSelect.Count <= 0)
                {
                    MessageBox.Show("해당 거래처의 선택된 품목이 없습니다.");
                    flag = false;
                    return flag;
                }

                // 해당 거래처 코드가 존재하는지 체크
                if (strFlag.Trim().Equals("I"))
                {
                    int count = 0;
                    string chkCnt = "";

                    Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                    sqlParameter.Clear();

                    sqlParameter.Add("sCustom", txtCustom.Tag != null ? txtCustom.Tag.ToString() : "");
                    DataSet ds = DataStore.Instance.ProcedureToDataSet("xp_Custom_checkCustomArticle", sqlParameter, false);

                    if (ds != null && ds.Tables.Count > 0)
                    {
                        DataTable dt = ds.Tables[0];

                        if (dt.Rows.Count > 0)
                        {
                            DataRowCollection drc = dt.Rows;

                            foreach (DataRow dr in drc)
                            {
                                chkCnt = dr["Num"].ToString();

                                if (chkCnt != null && !chkCnt.Trim().Equals(""))
                                    break;
                            }

                            count = Int32.Parse(chkCnt);

                            if (count > 0)
                            {
                                MessageBox.Show("해당 거래처 정보는 이미 존재합니다.");
                                flag = false;
                                return flag;
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally
            {
                DataStore.Instance.CloseConnection(); //2021-09-13 현달씨 DBClose
            }

            return flag;
        }

        //거래처
        private void txtCustom_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                e.Handled = true;

                if (cboInGubunSrh.SelectedValue.ToString().Equals("1")) //매출만
                {
                    MainWindow.pf.ReturnCode(txtCustom, 7072, "");
                    //MainWindow.pf.ReturnCode(txtCustom, (int)Defind_CodeFind.DCF_CUSTOM, "");
                }
                else
                {
                    MainWindow.pf.ReturnCode(txtCustom, 7073, ""); //매입+매입매출만
                }

            }
        }

        private void btnPfCustom_Click(object sender, RoutedEventArgs e)
        {

            if (cboInGubunSrh.SelectedValue.ToString().Equals("1")) //매출만
            {
                MainWindow.pf.ReturnCode(txtCustom, 7072, "");
                //MainWindow.pf.ReturnCode(txtCustom, (int)Defind_CodeFind.DCF_CUSTOM, "");
            }
            else
            {
                MainWindow.pf.ReturnCode(txtCustom, 7073, ""); //매입만
            }
            //MainWindow.pf.ReturnCode(txtCustom, (int)Defind_CodeFind.DCF_CUSTOM, "");
        }

        //품목추가
        private void btnAddSelectItem_Click(object sender, RoutedEventArgs e)
        {
            Win_com_CustomArticle_U_CodeView_SorD allItem = null;

            // 에러메시지용 변수
            int MsgCnt = 0;
            string Msg = "";

            // 딱 넘겨주는것만 dgdSel 에 추가해주기
            List<Win_com_CustomArticle_U_CodeView_SorD> lstTemp = new List<Win_com_CustomArticle_U_CodeView_SorD>();

            for (int i = 0; i < dgdAllItem.Items.Count; i++)
            {
                allItem = null;
                allItem = dgdAllItem.Items[i] as Win_com_CustomArticle_U_CodeView_SorD;
                if (allItem.chkFlag == true)
                {
                    //lstSelect.Add(allItem);
                    // 2020.03.18 수정
                    // lstSelect 에 이미 있는 품명이라면, 메시지 띄우기.
                    if (lstSelect.Count > 0)
                    {
                        bool good = true;
                        for (int k = 0; k < lstSelect.Count; k++)
                        {
                            var Article = lstSelect[k] as Win_com_CustomArticle_U_CodeView_SorD;
                            if (Article != null)
                            {
                                if (Article.ArticleID.Trim().Equals(allItem.ArticleID.Trim()))
                                {
                                    MsgCnt++;
                                    Msg += Article.BuyerArticleNo + "\r";

                                    good = false;
                                    break;
                                }
                            }
                        }
                        if (good == true)
                        {
                            lstTemp.Add(allItem);
                            lstSelect.Add(allItem);
                        }
                    }
                    else
                    {
                        lstTemp.Add(allItem);
                        lstSelect.Add(allItem);
                    }
                }

            }

            if (lstTemp.Count <= 0)
            {
                MessageBox.Show("품번이 선택되지 않았습니다.");
            }


            //if (dgdSelectItem.Items.Count > 0)
            //{
            //    dgdSelectItem.Items.Clear();
            //}

            for (int j = 0; lstTemp.Count > j; j++)
            {
                var selectionItem = lstTemp[j] as Win_com_CustomArticle_U_CodeView_SorD;
                selectionItem.chkFlag = false;
                selectionItem.SelectNum = (j + 1).ToString();
                dgdSelectItem.Items.Add(selectionItem);
            }
            tbkSelectCount.Text = "선택품목 : " + dgdSelectItem.Items.Count.ToString() + "개";

            FillGridAllItem();


            // 에러 메시지 띄우기
            if (MsgCnt > 0)
            {
                MessageBox.Show(Msg + "위의 품목(들)은 이미 등록되어 있습니다.");
            }
        }

        //품목제외
        private void btnDelSelectItem_Click(object sender, RoutedEventArgs e)
        {
            Win_com_CustomArticle_U_CodeView_SorD selectItem = null;

            // 딱 넘겨주는것만 dgdSel 에 삭제해주기
            List<Win_com_CustomArticle_U_CodeView_SorD> lstTemp = new List<Win_com_CustomArticle_U_CodeView_SorD>();

            for (int i = 0; i < dgdSelectItem.Items.Count; i++)
            {
                selectItem = null;
                selectItem = dgdSelectItem.Items[i] as Win_com_CustomArticle_U_CodeView_SorD;
                if (selectItem.chkFlag == true)
                {
                    lstSelect.Remove(selectItem);
                    lstTemp.Add(selectItem);
                }

            }

            if (lstTemp.Count <= 0)
            {
                MessageBox.Show("삭제할 품번이 선택되지 않았습니다.");
            }
            //if (dgdSelectItem.Items.Count > 0)
            //{
            //    dgdSelectItem.Items.Clear();
            //}

            for (int j = 0; lstTemp.Count > j; j++)
            {
                //var selectionItem = lstSelect[j] as Win_com_CustomArticle_U_CodeView_SorD;
                //selectionItem.chkFlag = false;
                //selectionItem.SelectNum = (j + 1).ToString();
                dgdSelectItem.Items.Remove(lstTemp[j]);
            }
            tbkSelectCount.Text = "선택품목 : " + dgdSelectItem.Items.Count.ToString() + "개";

            FillGridAllItem();
        }

        //
        private void chkAllItem_Click(object sender, RoutedEventArgs e)
        {
            CheckBox chkSender = sender as CheckBox;
            if (lblMsg.Visibility == Visibility.Visible)
            {
                var dgdAll = chkSender.DataContext as Win_com_CustomArticle_U_CodeView_SorD;

                if (chkSender.IsChecked == true)
                {
                    dgdAll.chkFlag = true;
                }
                else
                {
                    dgdAll.chkFlag = false;
                }
            }
            else
            {
                if (chkSender.IsChecked == true)
                {
                    chkSender.IsChecked = false;
                }
                else
                {
                    chkSender.IsChecked = true;
                }
                MessageBox.Show("체크박스를 사용하려면 먼저 추가나 수정을 누르고 진행해야 합니다.");
            }
        }

        //
        private void chkSelectItem_Click(object sender, RoutedEventArgs e)
        {
            CheckBox chkSender = sender as CheckBox;
            if (lblMsg.Visibility == Visibility.Visible)
            {
                var dgdSelect = chkSender.DataContext as Win_com_CustomArticle_U_CodeView_SorD;

                if (chkSender.IsChecked == true)
                {
                    dgdSelect.chkFlag = true;
                }
                else
                {
                    dgdSelect.chkFlag = false;
                }
            }
            else
            {
                if (chkSender.IsChecked == true)
                {
                    chkSender.IsChecked = false;
                }
                else
                {
                    chkSender.IsChecked = true;
                }
                MessageBox.Show("체크박스를 사용하려면 먼저 추가나 수정을 누르고 진행해야 합니다.");
            }
        }

        //전체선택 - 전체품목
        private void AllCheck_Checked(object sender, RoutedEventArgs e)
        {
            CustomArticle.Clear();

            for (int i = 0; i < dgdAllItem.Items.Count; i++)
            {
                var CtArticle = dgdAllItem.Items[i] as Win_com_CustomArticle_U_CodeView_SorD;
                CtArticle.chkFlag = true;

                CustomArticle.Add(CtArticle);
            }
        }

        //전체선택 - 전체품목
        private void AllCheck_Unchecked(object sender, RoutedEventArgs e)
        {
            CustomArticle.Clear();

            for (int i = 0; i < dgdAllItem.Items.Count; i++)
            {
                var CtArticle = dgdAllItem.Items[i] as Win_com_CustomArticle_U_CodeView_SorD;
                CtArticle.chkFlag = false;
            }
        }

        //전체선택 - 선택된품목
        private void AllCheckTwo_Checked(object sender, RoutedEventArgs e)
        {
            CustomArticle.Clear();

            for (int i = 0; i < dgdSelectItem.Items.Count; i++)
            {
                var CtArticle = dgdSelectItem.Items[i] as Win_com_CustomArticle_U_CodeView_SorD;
                CtArticle.chkFlag = true;

                CustomArticle.Add(CtArticle);
            }
        }
        //전체선택 - 선택된품목
        private void AllCheckTwo_Unchecked(object sender, RoutedEventArgs e)
        {
            CustomArticle.Clear();

            for (int i = 0; i < dgdSelectItem.Items.Count; i++)
            {
                var CtArticle = dgdSelectItem.Items[i] as Win_com_CustomArticle_U_CodeView_SorD;
                CtArticle.chkFlag = true;

                CustomArticle.Add(CtArticle);
            }
        }

        // 전체 품목 검색시!!!!!!!
        private void btnArticleSrhAllItem_Click(object sender, RoutedEventArgs e)
        {
            #region [안씀] 2020.03.18 이전 - 품목 검색 → 해당 검색품목 선택되고, 스크롤도 이동되도록 
            //if (indexAllItem != 0) { indexAllItem += 1; }
            //try
            //{
            //    // 일단은 for문을 돌면서!!! 해당 하는것 찾기!!!!!
            //    for (int i = indexAllItem; i < dgdAllItem.Items.Count; i++)
            //    {
            //        var WinArticle = dgdAllItem.Items[i] as Win_com_CustomArticle_U_CodeView_SorD;
            //        if (WinArticle != null)
            //        {
            //            if (WinArticle.BuyerArticleNo.Trim().Contains(txtArticleSrhAllItem.Text.Trim()))
            //            {
            //                // first focus the grid
            //                dgdAllItem.Focus();
            //                //then create a new cell info, with the item we wish to edit and the column number of the cell we want in edit mode
            //                DataGridCellInfo cellInfo = new DataGridCellInfo(WinArticle, dgdAllItem.Columns[0]);
            //                //set the cell to be the active one
            //                dgdAllItem.CurrentCell = cellInfo;
            //                //scroll the item into view
            //                dgdAllItem.ScrollIntoView(WinArticle);
            //                //begin the edit
            //                dgdAllItem.BeginEdit();

            //                dgdAllItem.SelectedIndex = i;
            //                dgdAllItem.CurrentCell = new DataGridCellInfo(dgdAllItem.Items[i], dgdAllItem.Columns[1]);
            //                //SelectDataGridRow(i);
            //                indexAllItem = i + 1;

            //                txtArticleSrhAllItem.Focus();
            //                txtArticleSrhAllItem.CaretIndex = txtArticleSrhAllItem.Text.Length;
            //                break;
            //            }
            //        }
            //    }
            //}
            //catch (Exception ex)
            //{
            //    MessageBox.Show(ex.Message);
            //}

            #endregion

            #region 그냥 mt_Article 조회 되도록

            lstAll.Clear();

            if (dgdAllItem.Items.Count > 0)
            {
                dgdAllItem.Items.Clear();
            }

            try
            {
                Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                sqlParameter.Add("sArticle", "");
                sqlParameter.Add("iIncNotUse", 0);
                sqlParameter.Add("sArticleGrpID", "");
                sqlParameter.Add("sSupplyType", "");
                sqlParameter.Add("nBuyerArticleNo", 1);
                sqlParameter.Add("BuyerArticleNo", txtArticleSrhAllItem.Text); //승인여부

                #region 이것도저것도 아닐때
                //if (cboInGubunSrh.SelectedValue.ToString().Equals(""))
                //{
                DataSet ds = DataStore.Instance.ProcedureToDataSet("xp_Article_sArticle_Custom", sqlParameter, true);

                if (ds != null && ds.Tables.Count > 0)
                {
                    DataTable dt = ds.Tables[0];

                    if (dt.Rows.Count > 0)
                    {
                        int i = 0;
                        DataRowCollection drc = dt.Rows;

                        foreach (DataRow dr in drc)
                        {
                            i++;
                            var Article = new Win_com_CustomArticle_U_CodeView_SorD()
                            {
                                AllNum = i.ToString(),
                                ArticleID = dr["ArticleID"].ToString(),
                                Article = dr["Article"].ToString(),
                                BuyerArticleNo = dr["BuyerArticleNo"].ToString(),
                                chkFlag = false
                            };

                            dgdAllItem.Items.Add(Article);
                            lstAll.Add(Article);
                        }

                        tbkAllCount.Text = "전체 품목 : " + dgdAllItem.Items.Count.ToString() + "개";
                    }
                }
                //}
                #endregion 이것도저것도 아닐 때

                //#region 매출만
                ////DataSet ds = DataStore.Instance.ProcedureToDataSet("xp_Article_sArticle_Custom", sqlParameter, true);
                //if (cboInGubunSrh.SelectedValue.ToString().Equals("1")) // 매출
                //{
                //    DataSet ds = DataStore.Instance.ProcedureToDataSet("xp_Article_sArticle_Custom_buy", sqlParameter, true);

                //    if (ds != null && ds.Tables.Count > 0)
                //    {
                //        DataTable dt = ds.Tables[0];

                //        if (dt.Rows.Count > 0)
                //        {
                //            int i = 0;
                //            DataRowCollection drc = dt.Rows;

                //            foreach (DataRow dr in drc)
                //            {
                //                i++;
                //                var Article = new Win_com_CustomArticle_U_CodeView_SorD()
                //                {
                //                    AllNum = i.ToString(),
                //                    ArticleID = dr["ArticleID"].ToString(),
                //                    //Article = dr["Article"].ToString(),
                //                    BuyerArticleNo = dr["BuyerArticleNo"].ToString(),
                //                    chkFlag = false
                //                };

                //                dgdAllItem.Items.Add(Article);
                //                lstAll.Add(Article);
                //            }

                //            tbkAllCount.Text = "전체 품목 : " + dgdAllItem.Items.Count.ToString() + "개";
                //        }
                //    }
                //}
                //#endregion 매출만

                //#region 매입만
                //if (cboInGubunSrh.SelectedValue.ToString().Equals("2")) //매입만
                //{
                //    DataSet ds = DataStore.Instance.ProcedureToDataSet("xp_Article_sArticle_Custom_sales", sqlParameter, true);

                //    if (ds != null && ds.Tables.Count > 0)
                //    {
                //        DataTable dt = ds.Tables[0];

                //        if (dt.Rows.Count > 0)
                //        {
                //            int i = 0;
                //            DataRowCollection drc = dt.Rows;

                //            foreach (DataRow dr in drc)
                //            {
                //                i++;
                //                var Article = new Win_com_CustomArticle_U_CodeView_SorD()
                //                {
                //                    AllNum = i.ToString(),
                //                    ArticleID = dr["ArticleID"].ToString(),
                //                    //Article = dr["Article"].ToString(),
                //                    BuyerArticleNo = dr["BuyerArticleNo"].ToString(),
                //                    chkFlag = false
                //                };

                //                dgdAllItem.Items.Add(Article);
                //                lstAll.Add(Article);
                //            }

                //            tbkAllCount.Text = "전체 품목 : " + dgdAllItem.Items.Count.ToString() + "개";
                //        }
                //    }
                //}
                //#endregion 매입만


            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
            finally
            {
                DataStore.Instance.CloseConnection();
            }

            #endregion
        }



        // 개별품목 검색시
        private void btnArticleSrhSelectItem_Click(object sender, RoutedEventArgs e)
        {
            #region [안씀] 2020.03.18 이전 - 품목 검색 → 해당 검색품목 선택되고, 스크롤도 이동되도록 
            ////if (indexSelItem != 0) { indexSelItem += 1; }

            //try
            //{
            //    // 일단은 for문을 돌면서!!! 해당 하는것 찾기!!!!!
            //    for (int i = indexSelItem; i < dgdSelectItem.Items.Count; i++)
            //    {
            //        var WinArticle = dgdSelectItem.Items[i] as Win_com_CustomArticle_U_CodeView_SorD;
            //        if (WinArticle != null)
            //        {
            //            if (WinArticle.BuyerArticleNo.Trim().Contains(txtArticleSrhSelectItem.Text.Trim()))
            //            {
            //                // first focus the grid
            //                dgdSelectItem.Focus();
            //                //then create a new cell info, with the item we wish to edit and the column number of the cell we want in edit mode
            //                DataGridCellInfo cellInfo = new DataGridCellInfo(WinArticle, dgdSelectItem.Columns[0]);
            //                //set the cell to be the active one
            //                dgdSelectItem.CurrentCell = cellInfo;
            //                //scroll the item into view
            //                dgdSelectItem.ScrollIntoView(WinArticle);
            //                //begin the edit
            //                dgdSelectItem.BeginEdit();

            //                dgdSelectItem.SelectedIndex = i;
            //                dgdSelectItem.CurrentCell = new DataGridCellInfo(dgdSelectItem.Items[i], dgdSelectItem.Columns[1]);
            //                //SelectDataGridRow(i);
            //                indexSelItem = i + 1;

            //                txtArticleSrhSelectItem.Focus();
            //                txtArticleSrhSelectItem.CaretIndex = txtArticleSrhSelectItem.Text.Length;
            //                break;
            //            }
            //        }
            //    }
            //}
            //catch(Exception ex)
            //{
            //    MessageBox.Show(ex.Message);
            //}



            //// 만약에 해당하는 품목이 없다면?
            //if (indexSelItem == 0)
            //{
            //    MessageBox.Show("해당하는 품목이 없습니다.\r(해당하는 품목이 존재하지 않거나, 선택된 품목입니다.)");
            //    return;
            //}
            //else
            //{
            //    MessageBox.Show("해당하는 품목이 더이상 존재하지 않습니다.");
            //    return;
            //}

            #endregion

            if (dgdSelectItem.Items.Count > 0)
            {
                dgdSelectItem.Items.Clear();
            }

            // lstSelect 에서 조회 되도록
            if (lstSelect.Count > 0)
            {

                int index = 0;
                // Win_com_CustomArticle_U_CodeView_SorD
                for (int i = 0; i < lstSelect.Count; i++)
                {
                    var Article = lstSelect[i] as Win_com_CustomArticle_U_CodeView_SorD;

                    if (Article != null)
                    {
                        // 둘다 대문자화 시켜서 비교 되도록
                        if (Article.BuyerArticleNo.ToUpper().Trim().Contains(txtArticleSrhSelectItem.Text.ToUpper().Trim()))
                        {
                            index++;
                            Article.SelectNum = index.ToString();
                            dgdSelectItem.Items.Add(lstSelect[i]);
                        }
                    }
                }
            }

            // 검색된게 없다면?
            if (dgdSelectItem.Items.Count == 0)
            {
                MessageBox.Show("조회된 데이터가 없습니다.");
            }
        }

        private void txtArticleSrhSelectItem_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                btnArticleSrhSelectItem_Click(null, null);
            }
            else
            {
                indexSelItem = 0;
            }
        }

        private void dgdSelectItem_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            indexSelItem = dgdSelectItem.SelectedIndex;
        }

        private void txtArticleSrhAllItem_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                btnArticleSrhAllItem_Click(null, null);
            }
            else
            {
                indexAllItem = 0;
            }
        }

        private void dgdAllItem_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            indexAllItem = dgdSelectItem.SelectedIndex;
        }

        //private void ChkShowDetail_Click(object sender, RoutedEventArgs e)
        //{
        //    if(chkShowDetail.IsChecked == true)
        //    {
        //        Article_ID.Visibility = Visibility.Visible;
        //        Article.Visibility = Visibility.Visible;

        //        //Article_ID.Width = new GridLength(Article_ID.ActualWidth);
        //        //Article.Width = new GridLength(Article.ActualWidth);
        //    }
        //    else
        //    {

        //        Article_ID.Visibility = Visibility.Hidden;
        //        Article.Visibility = Visibility.Hidden;


        //        //Article_ID.Width = new GridLength(0);
        //        //Article.Width = new GridLength(0);

        //    }
        //}


        private void ChkShowDetail_Checked(object sender, RoutedEventArgs e)
        {
            Kcustom.Width = new DataGridLength(1, DataGridLengthUnitType.Auto);
            Article_ID.Visibility = Visibility.Visible;
            //Article.Visibility = Visibility.Visible;

            BuyerArticleNo.Visibility = Visibility.Visible;

            re_Search(rowNum);

        }

        private void ChkShowDetail_Unchecked(object sender, RoutedEventArgs e)
        {
            Kcustom.Width = new DataGridLength(1, DataGridLengthUnitType.Star);
            Article_ID.Visibility = Visibility.Hidden;
            //Article.Visibility = Visibility.Hidden;
            BuyerArticleNo.Visibility = Visibility.Hidden;

            re_Search(rowNum);

        }

        private void txtCustomSrh_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                re_Search(0);
                //if (cboCustomGubunSrh.SelectedValuePath.Equals("1"))
                //{

                //}
                //else
                //{

                //}
            }
        }

        // 구분 바뀔 떄
        private void Gubun_Changed(object sender, SelectionChangedEventArgs e)
        {
            btnArticleSrhAllItem_Click(sender, e);

        }

        #region //셀 삽입
        // 2019.08.27 PreviewKeyDown 는 key 다운과 같은것 같음
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

        // KeyDown 이벤트
        private void DataGird_KeyDown(object sender, KeyEventArgs e)
        {
            int currRow = dgdSelectItem.Items.IndexOf(dgdSelectItem.CurrentItem);
            int currCol = dgdSelectItem.Columns.IndexOf(dgdSelectItem.CurrentCell.Column);
            int startCol = 1;
            int endCol = 5;

            if (e.Key == Key.Enter)
            {
                e.Handled = true;
                (sender as DataGridCell).IsEditing = false;

                // 마지막 열, 마지막 행 아님
                if (endCol == currCol && dgdSelectItem.Items.Count - 1 > currRow)
                {
                    dgdSelectItem.SelectedIndex = currRow + 1; // 이건 한줄 파란색으로 활성화 된 걸 조정하는 것입니다.
                    dgdSelectItem.CurrentCell = new DataGridCellInfo(dgdSelectItem.Items[currRow + 1], dgdSelectItem.Columns[startCol]);

                } // 마지막 열 아님
                else if (endCol > currCol && dgdSelectItem.Items.Count - 1 >= currRow)
                {
                    dgdSelectItem.CurrentCell = new DataGridCellInfo(dgdSelectItem.Items[currRow], dgdSelectItem.Columns[currCol + 1]);
                } // 마지막 열, 마지막 행
                else if (endCol == currCol && dgdSelectItem.Items.Count - 1 == currRow)
                {
                    //btnSave.Focus();
                }
                else
                {
                    MessageBox.Show("나머지가 있나..");
                }
            }
            else if (e.Key == Key.Down)
            {
                e.Handled = true;
                (sender as DataGridCell).IsEditing = false;

                // 마지막 행 아님
                if (dgdSelectItem.Items.Count - 1 > currRow)
                {
                    dgdSelectItem.SelectedIndex = currRow + 1;
                    dgdSelectItem.CurrentCell = new DataGridCellInfo(dgdSelectItem.Items[currRow + 1], dgdSelectItem.Columns[currCol]);
                } // 마지막 행일때
                else if (dgdSelectItem.Items.Count - 1 == currRow)
                {
                    if (endCol > currCol) // 마지막 열이 아닌 경우, 열을 오른쪽으로 이동
                    {
                        //dgdSub.SelectedIndex = 0;
                        dgdSelectItem.CurrentCell = new DataGridCellInfo(dgdSelectItem.Items[currRow], dgdSelectItem.Columns[currCol + 1]);
                    }
                    else
                    {
                        //btnSave.Focus();
                    }
                }
            }
            else if (e.Key == Key.Up)
            {
                e.Handled = true;
                (sender as DataGridCell).IsEditing = false;

                // 첫행 아님
                if (currRow > 0)
                {
                    dgdSelectItem.SelectedIndex = currRow - 1;
                    dgdSelectItem.CurrentCell = new DataGridCellInfo(dgdSelectItem.Items[currRow - 1], dgdSelectItem.Columns[currCol]);
                } // 첫 행
                else if (dgdSelectItem.Items.Count - 1 == currRow)
                {
                    if (0 < currCol) // 첫 열이 아닌 경우, 열을 왼쪽으로 이동
                    {
                        //dgdSub.SelectedIndex = 0;
                        dgdSelectItem.CurrentCell = new DataGridCellInfo(dgdSelectItem.Items[currRow], dgdSelectItem.Columns[currCol - 1]);
                    }
                    else
                    {
                        //btnSave.Focus();
                    }
                }
            }
            else if (e.Key == Key.Left)
            {
                e.Handled = true;
                (sender as DataGridCell).IsEditing = false;

                if (startCol < currCol)
                {
                    dgdSelectItem.CurrentCell = new DataGridCellInfo(dgdSelectItem.Items[currRow], dgdSelectItem.Columns[currCol - 1]);
                }
                else if (startCol == currCol)
                {
                    if (0 < currRow)
                    {
                        dgdSelectItem.SelectedIndex = currRow - 1;
                        dgdSelectItem.CurrentCell = new DataGridCellInfo(dgdSelectItem.Items[currRow - 1], dgdSelectItem.Columns[endCol]);
                    }
                    else
                    {
                        //btnSave.Focus();
                    }
                }
            }
            else if (e.Key == Key.Right)
            {
                e.Handled = true;
                (sender as DataGridCell).IsEditing = false;

                if (endCol > currCol)
                {

                    dgdSelectItem.CurrentCell = new DataGridCellInfo(dgdSelectItem.Items[currRow], dgdSelectItem.Columns[currCol + 1]);
                }
                else if (endCol == currCol)
                {
                    if (dgdSelectItem.Items.Count - 1 > currRow)
                    {
                        dgdSelectItem.SelectedIndex = currRow + 1;
                        dgdSelectItem.CurrentCell = new DataGridCellInfo(dgdSelectItem.Items[currRow + 1], dgdSelectItem.Columns[startCol]);
                    }
                    else
                    {
                        //btnSave.Focus();
                    }
                }
            }
        }

        // KeyUp 이벤트

        private void DatagridIn_TextFocus(object sender, KeyEventArgs e)
        {
            // 엔터 → 포커스 = true → cell != null → 해당 텍스트박스가 null이 아니라면 
            // → 해당 텍스트박스가 포커스가 안되있음 SelectAll() or 포커스
            Lib.Instance.DataGridINTextBoxFocus(sender, e);
        }

        // GotFocus 이벤트
        private void DataGridCell_GotFocus(object sender, RoutedEventArgs e)
        {
            if (lblMsg.Visibility == Visibility.Visible)
            {
                int currCol = dgdSelectItem.Columns.IndexOf(dgdSelectItem.CurrentCell.Column);

                DataGridCell cell = sender as DataGridCell;
                if (currCol == 4
                    || currCol == 5
                    || currCol == 6)
                {
                    cell.IsEditing = true;
                }
            }


        }

        // 2019.08.27 MouseUp 이벤트
        private void DataGridCell_MouseUp(object sender, MouseButtonEventArgs e)
        {

            Lib.Instance.DataGridINTextBoxFocusByMouseUP(sender, e);


        }

        private void DataGridCell_PreviewMouseDown(object sender, MouseButtonEventArgs e)
        {

        }

        #endregion

        //최종고객사 
        private void txtFinalCustom_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                var FCustom = dgdSelectItem.SelectedItem as Win_com_CustomArticle_U_CodeView_SorD;

                if (FCustom != null)
                {
                    TextBox tb = new TextBox();

                    pf.ReturnCode(tb, 0, FCustom.FinalCustom); //최종고객사

                    if (tb.Tag != null)
                    {
                        FCustom.FinalCustomID = tb.Tag.ToString();
                        FCustom.FinalCustom = tb.Text;
                    }
                }
            }

        }


        // 소수로 변환
        private double ConvertDouble(string str)
        {
            if (str == null) { return 0; }

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
        // 천마리 콤마, 소수점 한자리
        private string stringFormatN0(object obj)
        {
            return string.Format("{0:N0}", obj);
        }

        //라벨클릭 고객품번
        private void lblBuyCustomSrh_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            if (chkBuyCustomSrh.IsChecked == true)
            {
                chkBuyCustomSrh.IsChecked = false;
            }
            else
            {
                chkBuyCustomSrh.IsChecked = true;
            }
        }
        //고객품번 체크
        private void chkBuyCustomSrh_Checked(object sender, RoutedEventArgs e)
        {
            chkBuyCustomSrh.IsChecked = true;

            txtBuyCustomSrh.IsEnabled = true;
            btnPfBuyCustomSrh.IsEnabled = true;
        }

        //고객품번 체크해제
        private void chkBuyCustomSrh_UnChecked(object sender, RoutedEventArgs e)
        {
            chkBuyCustomSrh.IsChecked = false;

            txtBuyCustomSrh.IsEnabled = false;
            btnPfBuyCustomSrh.IsEnabled = false;
        }

        //고객품번 키다운 pf
        private void txtBuyCustomSrh_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                e.Handled = true;
                MainWindow.pf.ReturnCode(txtBuyCustomSrh, 7076, "");
            }
        }

        //고객품번 버튼클릭
        private void btnPfBuyCustomSrh_Click(object sender, RoutedEventArgs e)
        {
            MainWindow.pf.ReturnCode(txtBuyCustomSrh, 7076, "");
        }


        //라벨 품명 클릭
        private void lblArticleSrh_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            if (chkArticleSrh.IsChecked == true)
            {
                chkArticleSrh.IsChecked = false;
            }
            else
            {
                chkArticleSrh.IsChecked = true;
            }
        }

        //품명 체크
        private void chkArticleSrh_Checked(object sender, RoutedEventArgs e)
        {
            chkArticleSrh.IsChecked = true;

            txtArticleSrh.IsEnabled = true;
            btnPfArticleSrh.IsEnabled = true;
        }
        //품명체크해제
        private void chkArticleSrh_UnChecked(object sender, RoutedEventArgs e)
        {
            chkArticleSrh.IsChecked = false;

            txtArticleSrh.IsEnabled = false;
            btnPfArticleSrh.IsEnabled = false;
        }

        //품명키다운 pf
        private void txtArticleSrh_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                e.Handled = true;
                MainWindow.pf.ReturnCode(txtArticleSrh, 7077, "");
            }
        }
        //품명 pf 버튼클릭
        private void btnPfArticleSrh_Click(object sender, RoutedEventArgs e)
        {
            MainWindow.pf.ReturnCode(txtArticleSrh, 7077, "");
        }
        //단가 콤마찍기
        private void TxtUnitPrice_PreviewKeyDown(object sender, TextCompositionEventArgs e)
        {
            Lib.Instance.CheckIsNumeric((TextBox)sender, e);
        }
    }

    class Win_com_CustomArticle_U_CodeView : BaseView
    {
        public override string ToString()
        {
            return (this.ReportAllProperties());
        }

        public string Num { get; set; }
        public string CustomID { get; set; }
        public string KCustom { get; set; }
        public string ArticleID { get; set; }
        public string Article { get; set; }
        public string BuyerArticleNo { get; set; }
        public string TradeID { get; set; }

    }

    class Win_com_CustomArticle_U_CodeView_SorD : BaseView
    {
        public override string ToString()
        {
            return (this.ReportAllProperties());
        }

        //public bool Chk { get; set; }
        public string AllNum { get; set; }
        public string SelectNum { get; set; }

        public string CustomID { get; set; }
        public string KCustom { get; set; }
        public string ArticleID { get; set; }
        public string Article { get; set; }

        public string BuyerArticleNo { get; set; }
        public string CustomBuyArticle { get; set; } //고객사품번
        public string FinalCustom { get; set; } //최종고객사
        public string FinalCustomID { get; set; } //최종고객사ID
        public string UnitPrice { get; set; } //최종고객사ID

        public bool chkFlag { get; set; }
    }
}
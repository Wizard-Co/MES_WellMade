using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Data;
//using System.Threading.Tasks;
using System.Threading;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using WizMes_WellMade.PopUP;
using WizMes_WellMade.PopUp;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;


namespace WizMes_WellMade
{
    /// <summary>
    /// Win_com_ArticleBOM_U.xaml에 대한 상호 작용 논리
    /// </summary>
    public partial class Win_com_Article_BOM_U : UserControl
    {
        string stDate = string.Empty;
        string stTime = string.Empty;

        string strFlag = string.Empty;
        string strDirection = string.Empty;

        string PChildArticle = string.Empty; //상하위 구분자 상위 - "P" 하위 - "C"

        int rowNum = 0;
        int ArticleBomCnt = 0;
        Lib lib = new Lib();
        // 인쇄 활용 객체
        private Microsoft.Office.Interop.Excel.Application excelapp;
        private Microsoft.Office.Interop.Excel.Workbook workbook;
        private Microsoft.Office.Interop.Excel.Worksheet worksheet;
        private Microsoft.Office.Interop.Excel.Range workrange;
        private Microsoft.Office.Interop.Excel.Worksheet copysheet;
        private Microsoft.Office.Interop.Excel.Worksheet pastesheet;

        WizMes_WellMade.PopUp.NoticeMessage msg = new WizMes_WellMade.PopUp.NoticeMessage();
        ObservableCollection<Win_com_ArticleBOM_ItemList> ovcArticleBom = new ObservableCollection<Win_com_ArticleBOM_ItemList>();

        //
        Win_com_ArticleBOM_ItemList WinArticleBomList = new Win_com_ArticleBOM_ItemList();
        Win_com_ArticleBOM_Code_U WinArticleBomCode = new Win_com_ArticleBOM_Code_U();
        DataTable dataTableArticle = null;
        DataTable dataTableBOM = null;

        public Win_com_Article_BOM_U()
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
            strDirection = "B";
            dataTableArticle = Procedure.Instance.GetArticle("", 1, "", "");
        }

        private void SetComboBox()
        {
            //품명그룹(조회, 입력)
            ObservableCollection<CodeView> ovcArticleGrp = ComboBoxUtil.Instance.Gf_DB_MT_sArticleGrp();
            cboArticleGrpSrh.ItemsSource = ovcArticleGrp;
            cboArticleGrpSrh.DisplayMemberPath = "code_name";
            cboArticleGrpSrh.SelectedValuePath = "code_id";

            cbosArticleGrpP.ItemsSource = ovcArticleGrp;
            cbosArticleGrpP.DisplayMemberPath = "code_name";
            cbosArticleGrpP.SelectedValuePath = "code_id";

            cbosArticleGrpS.ItemsSource = ovcArticleGrp;
            cbosArticleGrpS.DisplayMemberPath = "code_name";
            cbosArticleGrpS.SelectedValuePath = "code_id";

            //단위
            ObservableCollection<CodeView> ovcUnitClss = ComboBoxUtil.Instance.Gf_DB_CM_GetComCodeDataset(null, "MTRUNIT", "Y", "");
            this.cboUnitClss.ItemsSource = ovcUnitClss;
            this.cboUnitClss.DisplayMemberPath = "code_name";
            this.cboUnitClss.SelectedValuePath = "code_id";
        }

        //품명그룹 검색 조건 사용체크
        private void lblArticleGrpSrh_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            if (chkArticleGrpSrh.IsChecked == true) { chkArticleGrpSrh.IsChecked = false; }
            else { chkArticleGrpSrh.IsChecked = true; }
        }

        //품명그룹 검색 조건 사용체크
        private void chkArticleGrpSrh_Checked(object sender, RoutedEventArgs e)
        {
            chkArticleGrpSrh.IsEnabled = true;
            cboArticleGrpSrh.IsEnabled = true;
        }

        //품명그룹 검색 조건 사용체크
        private void chkArticleGrpSrh_Unchecked(object sender, RoutedEventArgs e)
        {
            chkArticleGrpSrh.IsEnabled = true;
            cboArticleGrpSrh.IsEnabled = false;
        }

        //품명 검색 조건 사용체크
        private void lblArticleSrh_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            if (chkArticleSrh.IsChecked == true) { chkArticleSrh.IsChecked = false; }
            else { chkArticleSrh.IsChecked = true; }
        }

        //품명 검색 조건 사용체크
        private void chkArticleSrh_Checked(object sender, RoutedEventArgs e)
        {
            txtArticleSrh.IsEnabled = true;
            txtArticleSrh.Focus();
        }

        //품명 검색 조건 사용체크
        private void chkArticleSrh_Unchecked(object sender, RoutedEventArgs e)
        {
            txtArticleSrh.IsEnabled = false;
        }

        //
        private void TxtArticleSrh_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                MainWindow.pf.ReturnCodeGLS(txtArticleSrh, 77, "");
                //MainWindow.pf.ReturnCode(txtArticleSrh, (int)Defind_CodeFind.DCF_BuyerArticleNo, "");
            }
        }

        //
        private void BtnPfArticleSrh_Click(object sender, RoutedEventArgs e)
        {
            MainWindow.pf.ReturnCodeGLS(txtArticleSrh, 77, "");
            //MainWindow.pf.ReturnCode(txtArticleSrh, (int)Defind_CodeFind.DCF_BuyerArticleNo, "");
        }

        //검색 조건 라디오 버튼 체크(상향식)
        private void rbnTwoSrh_Click(object sender, RoutedEventArgs e)
        {
            rbnTwoSrh.IsChecked = true;
            strDirection = "B";
            tblArticleSrh.Text = "상위품";

        }
        private void tblRbnTwoSrh(object sender, MouseButtonEventArgs e)
        {
            rbnTwoSrh.IsChecked = true;
            strDirection = "B";
            tblArticleSrh.Text = "상위품";

        }

        //검색 조건 라디오 버튼 체크(하향식)
        private void rbnThreeSrh_Click(object sender, RoutedEventArgs e)
        {
            rbnThreeSrh.IsChecked = true;
            strDirection = "U";
            tblArticleSrh.Text = "하위품";

        }
        private void tblRbnThreeSrh(object sender, MouseButtonEventArgs e)
        {
            rbnThreeSrh.IsChecked = true;
            strDirection = "U";
            tblArticleSrh.Text = "하위품";
        }

        // 사용안함 포함 체크박스 이벤트
        private void lblNotUseSrh_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
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

        /// <summary>
        /// 수정,추가 저장 후
        /// </summary>
        private void CanBtnControl()
        {
            // 수정 후, 수정 취소시 상위품목, 품명은 입력 가능하도록 설정
            txtParentArticle.Background = (SolidColorBrush)new BrushConverter().ConvertFrom("#fff2d2");
            txtParentArticle.IsReadOnly = false;
            btnPfParentArticle.IsEnabled = true;

            txtParentBuyArticle.Background = (SolidColorBrush)new BrushConverter().ConvertFrom("#fff2d2");
            txtParentBuyArticle.IsReadOnly = false;
            btnPfParentBuyArticle.IsEnabled = true;


            txtArticle.Background = (SolidColorBrush)new BrushConverter().ConvertFrom("#fff2d2");
            txtArticle.IsReadOnly = false;
            btnPfArticle.IsEnabled = true;

            txtBuyArticle.Background = (SolidColorBrush)new BrushConverter().ConvertFrom("#fff2d2");
            txtBuyArticle.IsReadOnly = false;
            btnPfBuyArticle.IsEnabled = true;

            Lib.Instance.UiButtonEnableChange_IUControl(this);
            gbxInput.IsHitTestVisible = false;
            gbxDataGrid.IsEnabled = false; // 상위품목, 하위품목 리스트는 추가 상태일때만 활성화
            tlvItemList.IsHitTestVisible = true;
        }

        /// <summary>
        /// 수정,추가 진행 중
        /// </summary>
        private void CantBtnControl()
        {
            Lib.Instance.UiButtonEnableChange_SCControl(this);
            gbxInput.IsHitTestVisible = true;
            gbxDataGrid.IsEnabled = true;
            tlvItemList.IsHitTestVisible = false;
        }

        //추가
        private void btnAdd_Click(object sender, RoutedEventArgs e)
        {


            CantBtnControl();
            tbkMsg.Text = "자료 입력 중";
            strFlag = "I";
            this.DataContext = null;
            cboUnitClss.SelectedIndex = 1;
            dtpFromDate.SelectedDate = DateTime.Today;
            dtpToDate.SelectedDate = DateTime.Today.AddYears(500);


            var WinBomList = tlvItemList.SelectedItem as TreeViewItem;

            if (WinBomList == null)
            {
                txtParentArticle.Tag = null;
                txtArticle.Tag = null;
                txtParentBuyArticle.Tag = null; //오른쪽것도 지웁시다
                txtBuyArticle.Tag = null;
            }
            else if (WinBomList != null)
            {
                WinArticleBomList = WinBomList.Header as Win_com_ArticleBOM_ItemList;

                if (WinArticleBomList != null)
                {
                    this.DataContext = WinArticleBomList;

                    txtArticle.Text = null;
                    txtArticle.Tag = null;

                    txtBuyArticle.Text = null; //오른쪽것도 지웁시다
                    txtBuyArticle.Tag = null;

                }

            }


        }

        //수정
        private void btnUpdate_Click(object sender, RoutedEventArgs e)
        {
            var WinBomList = tlvItemList.SelectedItem as TreeViewItem;

            if (WinBomList != null)
            {
                WinArticleBomList = WinBomList.Header as Win_com_ArticleBOM_ItemList;

                if (WinArticleBomList != null)
                {
                    tbkMsg.Text = "자료 수정 중";
                    CantBtnControl();
                    strFlag = "U";

                    // 상위품목, 품명은 수정 불가능 설정
                    txtParentArticle.Background = (SolidColorBrush)new BrushConverter().ConvertFrom("#c2fdc3");
                    txtParentArticle.IsReadOnly = true;
                    btnPfParentArticle.IsEnabled = false;

                    // 상위품목, 품명은 수정 불가능 설정
                    txtParentBuyArticle.Background = (SolidColorBrush)new BrushConverter().ConvertFrom("#c2fdc3");
                    txtParentBuyArticle.IsReadOnly = true;
                    btnPfParentBuyArticle.IsEnabled = false;


                    txtArticle.Background = (SolidColorBrush)new BrushConverter().ConvertFrom("#c2fdc3");
                    txtArticle.IsReadOnly = true;
                    btnPfArticle.IsEnabled = false;

                    txtBuyArticle.Background = (SolidColorBrush)new BrushConverter().ConvertFrom("#c2fdc3");
                    txtBuyArticle.IsReadOnly = true;
                    btnPfBuyArticle.IsEnabled = false;

                    // 상위품목, 하위품목 리스트는 수정일때는 비활성화
                    gbxDataGrid.IsEnabled = false;

                    this.DataContext = WinArticleBomList;
                }
            }
        }

        //삭제
        private void btnDelete_Click(object sender, RoutedEventArgs e)
        {
            var WinBomList = tlvItemList.SelectedItem as TreeViewItem;

            if (WinBomList != null)
            {
                WinArticleBomList = WinBomList.Header as Win_com_ArticleBOM_ItemList;

                if (WinArticleBomList == null)
                {
                    MessageBox.Show("삭제할 데이터가 지정되지 않았습니다. 삭제데이터를 지정하고 눌러주세요");
                }
                else if (WinBomList.Items.Count == 1)
                {
                    MessageBox.Show("해당 품목의 하위품 먼저 삭제해주세요.");
                }
                else
                {
                    if (MessageBox.Show("선택하신 항목을 삭제하시겠습니까?", "삭제 전 확인", MessageBoxButton.YesNo) == MessageBoxResult.Yes)
                    {
                        DataStore.Instance.InsertLogByForm(this.GetType().Name, "D");

                        //삭제 전 체크
                        if (!DeleteDataCheck(WinArticleBomList.ArticleID, WinArticleBomList.PARENTArticleID))
                            return;

                        if (tlvItemList.Items.Count > 0 && tlvItemList.SelectedItem != null)
                        {
                            rowNum = 0;
                        }

                        if (Procedure.Instance.DeleteData(WinArticleBomList.ArticleID, WinArticleBomList.PARENTArticleID,
                            "sArticleID", "sParentArticleID", "xp_Article_dArticleBOM"))
                        {
                            this.DataContext = null;
                            //rowNum -= 1;
                            re_Search(rowNum);
                        }
                    }
                }
            }
        }

        //삭제 
        private bool DeleteData(string ArticleID, string PARENTArticleID, string SubArticleID)
        {
            bool flag = false;

            try
            {
                Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                sqlParameter.Clear();
                sqlParameter.Add("sArticleID", ArticleID);
                sqlParameter.Add("sParentArticleID", PARENTArticleID);
                sqlParameter.Add("sSubArticleID", SubArticleID);

                string[] result = DataStore.Instance.ExecuteProcedure("xp_Article_dArticleBOM", sqlParameter, false);

                if (result[0].Equals("success"))
                {
                    //MessageBox.Show("성공 *^^*");
                    flag = true;
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

        //닫기
        private void btnClose_Click(object sender, RoutedEventArgs e)
        {
            DataStore.Instance.InsertLogByFormS(this.GetType().Name, stDate, stTime, "E");
            Lib.Instance.ChildMenuClose(this.ToString());


            //Application.Exit();

        }

        //조회
        private void btnSearch_Click(object sender, RoutedEventArgs e)
        {
            using (Loading lw = new Loading(FillGrid2))
            {
                lw.ShowDialog();
            }
            FillGridExcel();
            //검색버튼 비활성화
            //btnSearch.IsEnabled = false;

            //Dispatcher.BeginInvoke(new Action(() =>

            //{
            //    Thread.Sleep(2000);

            //    ArticleBomCnt = 0;
            //    rowNum = 0;

            //    re_Search(rowNum);

            //    if (tlvItemList.Items.Count > 0)
            //    {
            //        tblArticleBomCnt.Text = "▶검색결과 : " + stringFormatN0(ArticleBomCnt) + "건";
            //    }

            //}), System.Windows.Threading.DispatcherPriority.Background);



            //Dispatcher.BeginInvoke(new Action(() =>

            //{

            //    btnSearch.IsEnabled = true;

            //}), System.Windows.Threading.DispatcherPriority.Background);


        }

        //저장
        private void btnSave_Click(object sender, RoutedEventArgs e)
        {
            if (SaveData(strFlag))
            {
                CanBtnControl();

                lblMsg.Visibility = Visibility.Hidden;
                rowNum = 0;
                re_Search(rowNum);
            }
        }

        //취소
        private void btnCancel_Click(object sender, RoutedEventArgs e)
        {
            CanBtnControl();

            txtParentArticle.Text = "";
            txtParentArticle.Tag = null;

            txtParentBuyArticle.Text = ""; //오른쪽 품번도 지웁시다
            txtParentBuyArticle.Tag = null;

            txtArticle.Text = "";
            txtArticle.Tag = null;

            txtBuyArticle.Text = ""; //오른쪽 품번도 지웁시다
            txtBuyArticle.Tag = null;

            txtQty.Text = "";
            txtLossQty.Text = "";


            cbosArticleGrpP.SelectedValue = false;
            cbosArticleGrpS.SelectedValue = false;

            strFlag = string.Empty;

            re_Search(rowNum);


            if (dgdArticleP.Items.Count > 0)
            {
                dgdArticleP.Items.Clear();
            }


            if (dgdArticleC.Items.Count > 0)
            {
                dgdArticleC.Items.Clear();
            }
            //re_Search_C(rowNum);
            //re_Search_P(rowNum);


            //dgdArticleP = null;
            //dgdArticleC = null;

        }

        //엑셀
        private void btnExcel_Click(object sender, RoutedEventArgs e)
        {
            DataTable dt = null;
            string Name = string.Empty;

            string[] lst = new string[2];
            lst[0] = "품명BOM 리스트";
            //lst[1] = "품명 BOM_상위품목";
            //lst[2] = "품명 BOM_하위품목";
            lst[1] = dgdExcel.Name;
            //lst[4] = dgdArticleP.Name;
            //lst[5] = dgdArticleC.Name;

            ExportExcelxaml ExpExc = new ExportExcelxaml(lst);
            ExpExc.ShowDialog();

            if (ExpExc.DialogResult.HasValue)
            {
                DataStore.Instance.InsertLogByForm(this.GetType().Name, "E");
                if (ExpExc.choice.Equals(dgdExcel.Name))
                {
                    if (ExpExc.Check.Equals("Y"))
                        dt = Lib.Instance.DataGridToDTinHidden(dgdExcel);
                    else
                        dt = Lib.Instance.DataGirdToDataTable(dgdExcel);

                    Name = dgdExcel.Name;

                    //Lib.Instance.GenerateExcel(dataTableBOM, Name);
                    if (Lib.Instance.GenerateExcel(dt, Name))
                        Lib.Instance.excel.Visible = true;
                    else
                        return;
                }
                //else if (ExpExc.choice.Equals(dgdArticleP.Name))
                //{
                //    if (ExpExc.Check.Equals("Y"))
                //        dt = Lib.Instance.DataGridToDTinHidden(dgdArticleP);
                //    else
                //        dt = Lib.Instance.DataGirdToDataTable(dgdArticleP);

                //    Name = dgdArticleP.Name;
                //    Lib.Instance.GenerateExcel(dt, Name);
                //    Lib.Instance.excel.Visible = true;
                //}
                //else if (ExpExc.choice.Equals(dgdArticleC.Name))
                //{
                //    if (ExpExc.Check.Equals("Y"))
                //        dt = Lib.Instance.DataGridToDTinHidden(dgdArticleC);
                //    else
                //        dt = Lib.Instance.DataGirdToDataTable(dgdArticleC);

                //    Name = dgdArticleC.Name;
                //    Lib.Instance.GenerateExcel(dt, Name);
                //    Lib.Instance.excel.Visible = true;
                //}
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
            FillGrid2();

            //임시 
            if (tlvItemList.Items.Count > 0
                && strFlag.Equals("I"))
            {
                txtParentArticle.Text = "";
                txtArticle.Text = "";
                txtParentBuyArticle.Text = "";
                txtBuyArticle.Text = "";

                txtQty.Text = "";
                txtLossQty.Text = "";
            }

            FillGridExcel();

        }

        private void re_Search_P(int selectedIndex)
        {
            FillGridP();

            if (dgdArticleP.Items.Count > 0)
            {
                dgdArticleP.Items.Clear();
            }
        }
        private void re_Search_C(int selectedIndex)
        {
            FillGridC();
        }



        #region 2020-02-11 신규 조회 FillGrid2()

        private void FillGrid2()
        {
            TreeViewItem mTreeViewItem = null;

            List<TreeViewItem> lstTree = new List<TreeViewItem>();

            List<string> lstParentArticleID = new List<string>();


            if (tlvItemList.Items.Count > 0)
            {
                tlvItemList.Items.Clear();
            }

            try
            {
                Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                sqlParameter.Add("sArticleGrpID", chkArticleGrpSrh.IsChecked == true && cboArticleGrpSrh.SelectedValue != null ? cboArticleGrpSrh.SelectedValue.ToString() : "");
                sqlParameter.Add("sArticleID", chkArticleSrh.IsChecked == true && txtArticleSrh.Tag != null ? txtArticleSrh.Tag.ToString() : "");
                sqlParameter.Add("sDirection", strDirection);
                sqlParameter.Add("sIncNotuse", chkNoUse.IsChecked == true ? 1 : 0);

                DataSet ds = DataStore.Instance.ProcedureToDataSet("xp_Article_sArticleBOM", sqlParameter, false);

                if (ds != null)
                {
                    DataTable dt = ds.Tables[0];
                    //GLS에서 요청 2021-10-21
                    if (dt.Rows.Count == 0)
                    {
                        MessageBox.Show("조회된 데이터가 없습니다.");
                    }
                    if (dt.Rows.Count > 0)
                    {
                        ArticleBomCnt = dt.Rows.Count;

                        int i = 0;
                        DataRowCollection drc = dt.Rows;

                        int BeforeLVL = 0;

                        foreach (DataRow dr in drc)
                        {
                            i++;
                            var ItemList = new Win_com_ArticleBOM_ItemList
                            {
                                Num = i,
                                Article = dr["Article"].ToString(),
                                ArticleID = dr["ArticleID"].ToString(),
                                LVL = ConvertInt(dr["LVL"].ToString()),
                                ArticleGrpID = dr["ArticleGrpID"].ToString(),
                                ArticleP = dr["ArticleP"].ToString(),
                                ChildBuyerArticleNO = dr["ChildBuyerArticleNO"].ToString(),
                                ChildCnt = dr["ChildCnt"].ToString(),
                                FromDate = dr["FromDate"].ToString(),
                                LossPcntClss = dr["LossPcntClss"].ToString(),
                                LossQty = dr["LossQty"].ToString(),
                                LvlPad = dr["LvlPad"].ToString(),
                                ord = dr["ord"].ToString(),
                                PARENTArticleID = dr["PARENTArticleID"].ToString(),
                                ParentArticleIDS = dr["ParentArticleIDS"].ToString(),
                                ParentBuyerArticleNO = dr["ParentBuyerArticleNO"].ToString(),
                                PcntClss = dr["PcntClss"].ToString(),
                                Qty = dr["Qty"].ToString(),
                                ScraptRate = dr["ScraptRate"].ToString(),
                                ToDate = dr["ToDate"].ToString(),
                                UnitClss = dr["UnitClss"].ToString(),
                                UnitClssName = dr["UnitClssName"].ToString(),
                                Weight = dr["Weight"].ToString(),
                                UseYN = dr["UseYN"].ToString(),
                                BuyerArticleNo = dr["BuyerArticleNo"].ToString(),

                                //SubArticleID = dr["SubArticleID"].ToString(),
                                //SubArticleName = dr["SubArticleName"].ToString(),
                                //chkSub = dr["chkSub"].ToString(),
                                //chkSubMe = dr["chkSubMe"].ToString(), //본인대체품인지
                            };

                            ItemList.Qty = Lib.Instance.returnNumString(ItemList.Qty);
                            ItemList.Weight = Lib.Instance.returnNumString(ItemList.Weight);
                            ItemList.LossQty = Lib.Instance.returnNumString(ItemList.LossQty);
                            ItemList.ScraptRate = Lib.Instance.returnNumString(ItemList.ScraptRate);

                            ItemList.FromDate_CV = Lib.Instance.StrDateTimeBar(ItemList.FromDate);
                            ItemList.ToDate_CV = Lib.Instance.StrDateTimeBar(ItemList.ToDate);
                            ItemList.FirstColumnCV = ItemList.LvlPad + "(" + ItemList.ArticleID + ")" + ItemList.BuyerArticleNo;

                            if (ItemList.LvlPad != "")
                            {
                                ItemList.FirstColumnCV2 = ItemList.LvlPad + " " + ItemList.Article;
                            }
                            else
                            {
                                ItemList.FirstColumnCV2 = ItemList.Article;
                            }

                            //if (ItemList.LVL != 1
                            //    && BeforeLVL + 1 != ItemList.LVL)
                            //{
                            //    continue;
                            //}


                            if (ItemList.LVL == 1)
                            {
                                mTreeViewItem = new TreeViewItem() { Header = ItemList, IsExpanded = true };
                                if (mTreeViewItem != null)
                                {
                                    tlvItemList.Items.Add(mTreeViewItem);
                                }

                                if (lstParentArticleID.Count == 0)
                                {
                                    lstParentArticleID.Add(ItemList.ArticleID);

                                    lstTree.Add(mTreeViewItem);
                                }
                                else
                                {
                                    lstParentArticleID = new List<string>();
                                    lstParentArticleID.Add(ItemList.ArticleID);

                                    lstTree = new List<TreeViewItem>();
                                    lstTree.Add(mTreeViewItem);
                                }
                            }
                            else
                            {
                                string CParentArticleID = ItemList.PARENTArticleID;
                                //string PParentArticleID = ;

                                var mTreeViewItemC = new TreeViewItem() { Header = ItemList, IsExpanded = true };

                                if (lstParentArticleID.Count < ItemList.LVL)
                                {
                                    // 1 다음이 3 ~ 9 라면..?
                                    #region 1안 → 봉인
                                    // → LVL 값을 조정
                                    //while (true)
                                    //{
                                    //    if (lstParentArticleID.Count != ItemList.LVL - 1)
                                    //    {
                                    //        ItemList.LVL -= 1;
                                    //        if (ItemList.LvlPad.Length > 1)
                                    //        {
                                    //            ItemList.LvlPad = ItemList.LvlPad.Substring(0, ItemList.LvlPad.Length - 1);
                                    //        }
                                    //    }
                                    //    else
                                    //    {
                                    //        break;
                                    //    }
                                    //}

                                    //ItemList.FirstColumnCV = ItemList.LvlPad + "(" + ItemList.ArticleID + ")" + ItemList.BuyerArticleNo;
                                    #endregion
                                }

                                // 첫 다음 레벨
                                if (lstParentArticleID.Count == ItemList.LVL - 1)
                                {
                                    if (lstParentArticleID[ItemList.LVL - 2].Equals(ItemList.PARENTArticleID))
                                    {
                                        lstTree[ItemList.LVL - 2].Items.Add(mTreeViewItemC);

                                        lstParentArticleID.Add(ItemList.ArticleID);
                                        lstTree.Add(mTreeViewItemC);
                                    }
                                }
                                else // 중복?
                                {
                                    //레벨에 벗어나고 상위품번이 없는경우
                                    if (lstParentArticleID.Count == 0)
                                    {
                                        mTreeViewItem = new TreeViewItem() { Header = ItemList, IsExpanded = true };
                                        if (mTreeViewItem != null)
                                        {
                                            tlvItemList.Items.Add(mTreeViewItem);
                                        }
                                        lstParentArticleID.Add(ItemList.ArticleID);
                                        lstTree.Add(mTreeViewItem);
                                    }

                                    else if (lstParentArticleID.Count <= ItemList.LVL - 2)
                                    {
                                        continue;
                                    }
                                    //else if (lstParentArticleID[ItemList.LVL - 1].Equals(ItemList.PARENTArticleID))
                                    //{
                                    //    lstTree[ItemList.LVL].Items.Add(mTreeViewItemC);

                                    //    lstParentArticleID[ItemList.LVL] = ItemList.ArticleID;
                                    //    lstTree[ItemList.LVL] = mTreeViewItemC;
                                    //}

                                    else if (lstParentArticleID[ItemList.LVL - 2].Equals(ItemList.PARENTArticleID))
                                    {
                                        lstTree[ItemList.LVL - 2].Items.Add(mTreeViewItemC);

                                        lstParentArticleID[ItemList.LVL - 1] = ItemList.ArticleID;
                                        lstTree[ItemList.LVL - 1] = mTreeViewItemC;
                                    }
                                }
                            }

                            BeforeLVL = ItemList.LVL;

                            tblArticleBomCnt.Text = "▶ 검색결과 : " + i + " 건";


                        } //foreach 문 끝 

                    }
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
        }

        #endregion
        #region
        //검색
        private void FillGridExcel()
        {
            if (dgdExcel.Items.Count > 0)
            {
                dgdExcel.Items.Clear();
            }

            try
            {
                Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                sqlParameter.Add("sArticleGrpID", chkArticleGrpSrh.IsChecked == true && cboArticleGrpSrh.SelectedValue != null ? cboArticleGrpSrh.SelectedValue.ToString() : "");
                sqlParameter.Add("sArticleID", chkArticleSrh.IsChecked == true && txtArticleSrh.Tag != null ? txtArticleSrh.Tag.ToString() : "");
                sqlParameter.Add("sDirection", strDirection);
                sqlParameter.Add("sIncNotuse", chkNoUse.IsChecked == true ? 1 : 0);

                DataSet ds = DataStore.Instance.ProcedureToDataSet_LogWrite("xp_Article_sArticleBOM", sqlParameter, true, "R");

                if (ds != null && ds.Tables.Count > 0)
                {
                    DataTable dt = ds.Tables[0];
                    int i = 0;

                    //dataGrid.Items.Clear();e
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
                            var WinExcel = new Win_com_ArticleBOM_ItemList()
                            {
                                Num = i,
                                Article = dr["Article"].ToString(),
                                ArticleID = dr["ArticleID"].ToString(),
                                LVL = ConvertInt(dr["LVL"].ToString()),
                                ArticleGrpID = dr["ArticleGrpID"].ToString(),
                                ArticleP = dr["ArticleP"].ToString(),
                                ChildBuyerArticleNO = dr["ChildBuyerArticleNO"].ToString(),
                                ChildCnt = dr["ChildCnt"].ToString(),
                                FromDate = dr["FromDate"].ToString(),
                                LossPcntClss = dr["LossPcntClss"].ToString(),
                                LossQty = dr["LossQty"].ToString(),
                                LvlPad = dr["LvlPad"].ToString(),
                                ord = dr["ord"].ToString(),
                                PARENTArticleID = dr["PARENTArticleID"].ToString(),
                                ParentArticleIDS = dr["ParentArticleIDS"].ToString(),
                                ParentBuyerArticleNO = dr["ParentBuyerArticleNO"].ToString(),
                                PcntClss = dr["PcntClss"].ToString(),
                                Qty = dr["Qty"].ToString(),
                                ScraptRate = dr["ScraptRate"].ToString(),
                                ToDate = dr["ToDate"].ToString(),
                                UnitClss = dr["UnitClss"].ToString(),
                                UnitClssName = dr["UnitClssName"].ToString(),
                                Weight = dr["Weight"].ToString(),
                                UseYN = dr["UseYN"].ToString(),
                                BuyerArticleNo = dr["BuyerArticleNo"].ToString()

                            };
                            //if ((item["cls"].ToString() != "3")
                            //if (dr["LVL"].ToString() == "1")
                            //{
                            //    WinExcel.LVL1 = WinExcel.LVL; 

                            //}
                            //if (dr["LVL"].ToString() == "2")
                            //{
                            //    WinExcel.LVL2 = WinExcel.LVL;
                            //}
                            //if (dr["LVL"].ToString() == "3")
                            //{
                            //    WinExcel.LVL3 = WinExcel.LVL;
                            //}
                            //if (dr["LVL"].ToString() == "4")
                            //{
                            //    WinExcel.LVL4 = WinExcel.LVL;
                            //}


                            //WinExcel.LVL2
                            //WinExcel.LVL3
                            //WinExcel.LVL4

                            WinExcel.Weight = Lib.Instance.returnNumString(WinExcel.Weight);
                            WinExcel.LossQty = Lib.Instance.returnNumString(WinExcel.LossQty);
                            WinExcel.ScraptRate = Lib.Instance.returnNumString(WinExcel.ScraptRate);

                            WinExcel.FromDate_CV = Lib.Instance.StrDateTimeBar(WinExcel.FromDate);
                            WinExcel.ToDate_CV = Lib.Instance.StrDateTimeBar(WinExcel.ToDate);


                            dgdExcel.Items.Add(WinExcel);
                            //ovcArticleBom.Add(WinExcel);
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

        #endregion

        private void TlvItemList_SelectedItemChanged(object sender, RoutedPropertyChangedEventArgs<object> e)
        {
            var WinBomList = tlvItemList.SelectedItem as TreeViewItem;

            if (WinBomList != null)
            {
                WinArticleBomList = WinBomList.Header as Win_com_ArticleBOM_ItemList;

                if (WinArticleBomList != null)
                {
                    this.DataContext = WinArticleBomList;
                    txtArticle.Tag = WinArticleBomList.ArticleID;
                    txtParentArticle.Tag = WinArticleBomList.PARENTArticleID;

                    // ↑소요량 수정불가 현상을 위해 추가 2022.05.06 (프로시저 저장시 상위품명의 코드를 찾지 못해서 소요량 수정이 정상적으로 작동하지 않음 )


                    if (WinArticleBomList.UseYN.Equals("Y"))
                    {
                        chkUseClss.IsChecked = false;
                    }
                    else
                    {
                        chkUseClss.IsChecked = true;
                    }


                }
            }
        }

        // 사용안함 ViewBox 클릭 이벤트
        private void vbUseClss_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            if (chkUseClss.IsChecked == true)
            {
                chkUseClss.IsChecked = false;
            }
            else
            {
                chkUseClss.IsChecked = true;
            }
        }
        // 퍼센트
        private void vbPercent_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            if (chkPercent.IsChecked == true)
            {
                chkPercent.IsChecked = false;
            }
            else
            {
                chkPercent.IsChecked = true;
            }
        }
        private void vbLossPercent_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            if (chkLossPercent.IsChecked == true)
            {
                chkLossPercent.IsChecked = false;
            }
            else
            {
                chkLossPercent.IsChecked = true;
            }
        }

        /// <summary>
        /// 저장
        /// </summary>
        /// <param name="strFlag"></param>
        /// <param name="strYYYY"></param>
        /// <returns></returns>
        private bool SaveData(string strFlag)
        {
            bool flag = false;
            List<Procedure> Prolist = new List<Procedure>();
            List<Dictionary<string, object>> ListParameter = new List<Dictionary<string, object>>();

            try
            {
                if (CheckData())
                {
                    // 만약에 상위품 하위품의 품명이 같다면!!!!!!!!!! → 그대로 들어가면 문제 생김
                    // 상위품을 지우고 하위품명으로만 들어가도록
                    if (txtParentArticle.Tag != null
                        && !txtParentArticle.Text.Trim().Equals("")
                        && txtArticle.Tag.ToString().Trim().Equals(txtParentArticle.Tag.ToString().Trim()))
                    {
                        txtParentArticle.Tag = "";


                    }

                    Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                    sqlParameter.Clear();

                    sqlParameter.Add("sArticleID", txtArticle.Tag.ToString());
                    sqlParameter.Add("sParentArticleID", txtParentArticle.Tag != null && !txtParentArticle.Text.Trim().Equals("") ? txtParentArticle.Tag.ToString() : "");
                    sqlParameter.Add("Qty", ConvertDouble(txtQty.Text));
                    sqlParameter.Add("PcntClss", chkPercent.IsChecked == true ? "*" : "");
                    sqlParameter.Add("UnitClss", cboUnitClss.SelectedValue.ToString());
                    sqlParameter.Add("LossQty", ConvertDouble(txtLossQty.Text));
                    sqlParameter.Add("LossPcntClss", chkLossPercent.IsChecked == true ? "*" : "");
                    sqlParameter.Add("FromDate", dtpFromDate.SelectedDate.Value.ToString("yyyyMMdd"));
                    sqlParameter.Add("ToDate", dtpToDate.SelectedDate.Value.ToString("yyyyMMdd"));
                    sqlParameter.Add("UseYN", chkUseClss.IsChecked == true ? "N" : "Y");



                    sqlParameter.Add("UserID", MainWindow.CurrentUser);
                    sqlParameter.Add("OutMsg", "");

                    #region 추가

                    if (strFlag.Equals("I"))
                    {
                        Procedure pro1 = new Procedure();
                        pro1.Name = "xp_Article_iArticleBOM";
                        pro1.OutputUseYN = "N";
                        pro1.OutputName = "sArticleID";
                        pro1.OutputLength = "10";

                        Prolist.Add(pro1);
                        ListParameter.Add(sqlParameter);

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
                        Procedure pro1 = new Procedure();
                        pro1.Name = "xp_Article_uArticleBOM";
                        pro1.OutputUseYN = "N";
                        pro1.OutputName = "sArticleID";
                        pro1.OutputLength = "10";

                        Prolist.Add(pro1);
                        ListParameter.Add(sqlParameter);

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

                    #endregion
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

        /// <summary>
        /// 입력사항 체크
        /// </summary>
        /// <returns></returns>
        private bool CheckData()
        {
            bool flag = true;

            //if (txtParentArticle.Text.Length <= 0 || txtParentArticle.Text.Equals(""))
            //{
            //    MessageBox.Show("상위품명이 입력되지 않았습니다.");
            //    flag = false;
            //    return flag;
            //}

            if (txtArticle.Tag == null || txtArticle.Text.Trim().Equals(""))
            {
                MessageBox.Show("품명이 입력되지 않았습니다.");
                flag = false;
                return flag;
            }

            if (txtQty.Text.Length <= 0 || txtQty.Text.Equals(""))
            {
                MessageBox.Show("소요량이 입력되지 않았습니다.");
                flag = false;
                return flag;
            }

            //if (txtLossQty.Text.Length <= 0 || txtLossQty.Text.Equals(""))
            //{
            //    MessageBox.Show("로스량이 입력되지 않았습니다.");
            //    flag = false;
            //    return flag;
            //}

            // 만약에 상위품 하위품의 품명이 같다면!!!!!!!!!! → 그대로 들어가면 문제 생김
            // 상위품 하위품 같은 품명으로 못들어가도록
            if (txtParentArticle.Tag != null
                && !txtParentArticle.Text.Trim().Equals("")
                && txtArticle.Tag.ToString().Trim().Equals(txtParentArticle.Tag.ToString().Trim()))
            {
                MessageBox.Show("상위품과 하위품이 같은 품명으로 저장이 불가능합니다.\r상위품을 지우거나 하위품을 변경해주세요.");
                flag = false;
                return flag;
            }

            return flag;
        }
        //삭제체크
        private bool DeleteDataCheck(string strArticleID, string strPArticleID)
        {
            bool Flag = false;

            try
            {
                Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                sqlParameter.Clear();
                sqlParameter.Add("sArticleID", strArticleID);
                sqlParameter.Add("sPArticleID", strPArticleID);

                string[] result = DataStore.Instance.ExecuteProcedure("xp_Article_dArticleBOM_Check", sqlParameter, false);
                string[] resultSplit;

                if (result[0].Equals("success") && result[1].Equals(""))
                {
                    Flag = true;
                }
                else
                {
                    resultSplit = result[1].Split('/');

                    if (resultSplit.Length == 2)
                    {
                        if (Convert.ToInt32(resultSplit[0]) <= 10)
                        {
                            if (MessageBox.Show(resultSplit[1] + " 무시하고 계속 진행하시겠습니까?", "삭제 전 확인", MessageBoxButton.YesNo) == MessageBoxResult.Yes)
                            {
                                Flag = true;
                            }
                        }
                        else
                        {
                            MessageBox.Show(resultSplit[1]);
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

            return Flag;
        }


        //상위품명
        private void txtParentArticle_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                MainWindow.pf.ReturnCodeGLS(txtParentArticle, 77, "");
                //MainWindow.pf.ReturnCode(txtParentArticle, (int)Defind_CodeFind.DCF_BuyerArticleNo, "");

                PChildArticle = "P";

                if (txtParentArticle.Tag != null)
                {
                    getArticleInfo(txtParentArticle.Tag.ToString());
                }
                PChildArticle = "";

            }
        }

        //상위품명
        private void btnPfParentArticle_Click(object sender, RoutedEventArgs e)
        {
            PChildArticle = "P";
            MainWindow.pf.ReturnCodeGLS(txtParentArticle, 77, "");
            //MainWindow.pf.ReturnCode(txtParentArticle, (int)Defind_CodeFind.DCF_BuyerArticleNo, "");

            if (txtParentArticle.Tag != null)
            {
                getArticleInfo(txtParentArticle.Tag.ToString());
            }

            PChildArticle = "";

        }

        //하위품명
        private void txtArticle_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                MainWindow.pf.ReturnCodeGLS(txtArticle, 77, "");
                //MainWindow.pf.ReturnCode(txtArticle, (int)Defind_CodeFind.DCF_BuyerArticleNo, "");
                PChildArticle = "C";

                if (txtArticle.Tag != null)
                {
                    getArticleInfo(txtArticle.Tag.ToString());
                }
                PChildArticle = "";
            }
        }

        //하위품명
        private void btnPfArticle_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                MainWindow.pf.ReturnCodeGLS(txtArticle, 77, "");
                //MainWindow.pf.ReturnCode(txtArticle, (int)Defind_CodeFind.DCF_BuyerArticleNo, "");

                PChildArticle = "C";

                if (txtArticle.Tag != null)
                {
                    getArticleInfo(txtArticle.Tag.ToString());
                }

                PChildArticle = "";

            }
            catch (Exception ex)
            {
                MessageBox.Show("오류 발생, 오류 내용 : " + ex.ToString());
            }

        }

        #region 품명 선택시 단위 자동으로 선택

        // 위의 getArticleIdByReq 로 가져온
        // ArticleID 로 Article 정보 가져오기
        private void getArticleInfo(string setArticleID)
        {
            try
            {
                Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                sqlParameter.Add("ArticleID", setArticleID);

                DataSet ds = DataStore.Instance.ProcedureToDataSet("xp_Order_sArticleData", sqlParameter, false);

                if (ds != null && ds.Tables.Count > 0)
                {
                    DataTable dt = ds.Tables[0];

                    if (dt.Rows.Count > 0)
                    {
                        DataRow dr = dt.Rows[0];

                        var getArticleInfo = new BOMArticleInfo
                        {
                            ArticleGrpID = dr["ArticleGrpID"].ToString(),
                            UnitPrice = dr["UnitPrice"].ToString(),
                            UnitPriceClss = dr["UnitPriceClss"].ToString(),
                            UnitClss = dr["UnitClss"].ToString(),
                            PartGBNID = dr["PartGBNID"].ToString(),
                            ProductGrpID = dr["ProductGrpID"].ToString(),
                            Article = dr["Article"].ToString(),
                            BuyerArticleNo = dr["BuyerArticleNo"].ToString(),
                            ArticleID = dr["ArticleID"].ToString()
                        };

                        cboUnitClss.SelectedValue = getArticleInfo.UnitClss;

                        if (PChildArticle.Equals("P"))
                        {
                            txtParentBuyArticle.Text = getArticleInfo.BuyerArticleNo;
                            txtParentArticle.Text = getArticleInfo.Article;
                            txtParentArticle.Tag = getArticleInfo.ArticleID;

                        }
                        else
                        {
                            txtBuyArticle.Text = getArticleInfo.BuyerArticleNo;
                            txtArticle.Text = getArticleInfo.Article;
                            txtArticle.Tag = getArticleInfo.ArticleID;

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

        #endregion // 품명 선택시 단위 자동으로 선택

        //트리뷰 펼쳐보기
        private void btnExpanding_Click(object sender, RoutedEventArgs e)
        {
            for (int i = 0; i < tlvItemList.Items.Count; i++)
            {
                var OneTree = tlvItemList.Items[i] as TreeViewItem;
                OneTree.IsExpanded = true;
            }
        }

        //트리뷰 접어보기
        private void btnFolding_Click(object sender, RoutedEventArgs e)
        {
            for (int i = 0; i < tlvItemList.Items.Count; i++)
            {
                var OneTree = tlvItemList.Items[i] as TreeViewItem;
                OneTree.IsExpanded = false;
            }
        }

        // 상위품목-제품 반제품 리스트 조회
        private void FillGridP()
        {
            string sql = string.Empty;

            if (dgdArticleP.Items.Count > 0)
            {
                dgdArticleP.Items.Clear();
            }

            try
            {
                DataTable dt = dataTableArticle;
                sql += " ArticleGrpID = '" + cbosArticleGrpP.SelectedValue.ToString() + "' ";
                sql += "and Article like '%" + txtSrhArticleP.Text + "%' ";


                //GLS에서 요청 2021-10-21
                if (dt.Rows.Count == 0)
                {
                    MessageBox.Show("조회된 데이터가 없습니다.");
                }

                if (dt.Rows.Count > 0)
                {
                    int i = 0;

                    foreach (DataRow dr in dt.Select(sql))
                    {
                        i++;
                        var ItemList = new Win_com_ArticleBOM_Code_U
                        {
                            Num = i.ToString(),
                            Article = dr["Article"].ToString(),
                            ArticleID = dr["ArticleID"].ToString(),
                            ArticleGrpID = dr["ArticleGrpID"].ToString(),
                            //Qty = dr["Qty"].ToString(),
                            UnitClss = dr["UnitClss"].ToString(),
                            UnitClssName = dr["UnitClssName"].ToString(),
                            Weight = dr["Weight"].ToString(),
                            //ArticleGrp = dr["ArticleGrp"].ToString(),
                            BuyerArticleNo = dr["BuyerArticleNo"].ToString(),
                            BuySaleMainYN = dr["BuySaleMainYN"].ToString(),
                            CoatingSpec = dr["CoatingSpec"].ToString(),
                            //DyeingID = dr["DyeingID"].ToString(),
                            FTAMgrYN = dr["FTAMgrYN"].ToString(),
                            HSCODE = dr["HSCODE"].ToString(),
                            LabelPrintYN = dr["LabelPrintYN"].ToString(),
                            NeedStockQty = dr["NeedStockQty"].ToString(),
                            OutUnitPrice = dr["OutUnitPrice"].ToString(),
                            PART_ATTR = dr["PART_ATTR"].ToString(),
                            PatternID = dr["PatternID"].ToString(),
                            //Process = dr["Process"].ToString(),
                            //ProcessID = dr["ProcessID"].ToString(),
                            //QtyPerBox = dr["QtyPerBox"].ToString(),
                            Spec = dr["Spec"].ToString(),
                            //StuffWidth = dr["StuffWidth"].ToString(),
                            SupplyType = dr["SupplyType"].ToString(),
                            SupplyTypeName = dr["SupplyTypeName"].ToString(),
                            //Thread = dr["Thread"].ToString(),
                            //ThreadID = dr["ThreadID"].ToString(),
                            Unitprice = dr["Unitprice"].ToString(),
                            UnitPriceClss = dr["UnitPriceClss"].ToString(),
                            UseClss = dr["UseClss"].ToString(),
                            UseingType = dr["UseingType"].ToString(),
                            UseingTypeName = dr["UseingTypeName"].ToString()
                        };

                        dgdArticleP.Items.Add(ItemList);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("오류 발생, 오류 내용 : " + ex.ToString());
            }
            finally
            {
                DataStore.Instance.CloseConnection(); //2021-09-13
            }
        }

        // 하위품 리스트 조회
        private void FillGridC()
        {
            string sql = string.Empty;

            if (dgdArticleC.Items.Count > 0)
            {
                dgdArticleC.Items.Clear();
            }

            try
            {
                DataTable dt = dataTableArticle;
                sql += " ArticleGrpID = '" + cbosArticleGrpS.SelectedValue.ToString() + "' ";
                sql += "and Article like '%" + txtSrhArticleC.Text + "%' ";

                //GLS에서 요청 2021-10-21
                if (dt.Rows.Count == 0)
                {
                    MessageBox.Show("조회된 데이터가 없습니다.");
                }

                if (dt.Rows.Count > 0)
                {
                    int i = 0;

                    foreach (DataRow dr in dt.Select(sql))
                    {
                        i++;
                        var ItemList = new Win_com_ArticleBOM_Code_U
                        {
                            Num = i.ToString(),
                            Article = dr["Article"].ToString(),
                            ArticleID = dr["ArticleID"].ToString(),
                            ArticleGrpID = dr["ArticleGrpID"].ToString(),
                            //Qty = dr["Qty"].ToString(),
                            UnitClss = dr["UnitClss"].ToString(),
                            UnitClssName = dr["UnitClssName"].ToString(),
                            Weight = dr["Weight"].ToString(),
                            //ArticleGrp = dr["ArticleGrp"].ToString(),
                            BuyerArticleNo = dr["BuyerArticleNo"].ToString(),
                            BuySaleMainYN = dr["BuySaleMainYN"].ToString(),
                            CoatingSpec = dr["CoatingSpec"].ToString(),
                            Comments = dr["Comments"].ToString(),
                            //DyeingID = dr["DyeingID"].ToString(),
                            FTAMgrYN = dr["FTAMgrYN"].ToString(),
                            HSCODE = dr["HSCODE"].ToString(),
                            //ImageName = dr["ImageName"].ToString(),
                            LabelPrintYN = dr["LabelPrintYN"].ToString(),
                            NeedStockQty = dr["NeedStockQty"].ToString(),
                            OutUnitPrice = dr["OutUnitPrice"].ToString(),
                            PART_ATTR = dr["PART_ATTR"].ToString(),
                            PatternID = dr["PatternID"].ToString(),
                            //Process = dr["Process"].ToString(),
                            //ProcessID = dr["ProcessID"].ToString(),
                            //QtyPerBox = dr["QtyPerBox"].ToString(),
                            Spec = dr["Spec"].ToString(),
                            //StuffWidth = dr["StuffWidth"].ToString(),
                            SupplyType = dr["SupplyType"].ToString(),
                            SupplyTypeName = dr["SupplyTypeName"].ToString(),
                            //Thread = dr["Thread"].ToString(),
                            //ThreadID = dr["ThreadID"].ToString(),
                            Unitprice = dr["Unitprice"].ToString(),
                            UnitPriceClss = dr["UnitPriceClss"].ToString(),
                            UseClss = dr["UseClss"].ToString(),
                            UseingType = dr["UseingType"].ToString(),
                            UseingTypeName = dr["UseingTypeName"].ToString()
                        };

                        dgdArticleC.Items.Add(ItemList);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("오류 발생, 오류 내용 : " + ex.ToString());
            }
            finally
            {
                DataStore.Instance.CloseConnection(); //2021-09-13
            }
        }

        // 상위품목-제품 반제품 리스트 조회
        private void btnSearchP_Click(object sender, RoutedEventArgs e)
        {
            if (cbosArticleGrpP.SelectedValue == null)
            {
                MessageBox.Show("품명그룹을 선택하세요");
                return;
            }

            FillGridP();
        }

        // 하위품 리스트 조회
        private void btnSearchC_Click(object sender, RoutedEventArgs e)
        {
            if (cbosArticleGrpS.SelectedValue == null)
            {
                MessageBox.Show("품명그룹을 선택하세요");
                return;
            }

            FillGridC();
        }

        private void DgdArticleP_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            var ArticleP = dgdArticleP.SelectedItem as Win_com_Article_U_CodeView;

            if (ArticleP != null)
            {
                txtParentArticle.Text = ArticleP.Article;
                txtParentArticle.Tag = ArticleP.ArticleID;
            }
        }

        private void DgdArticleC_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            var ArticleC = dgdArticleC.SelectedItem as Win_com_Article_U_CodeView;

            if (ArticleC != null)
            {
                txtArticle.Text = ArticleC.Article;
                txtArticle.Tag = ArticleC.ArticleID;
            }
        }


        // 상위품목-제품 반제품 리스트 선택버튼 클릭 이벤트
        private void btnArticleP_Click(object sender, MouseButtonEventArgs e)
        {
            WinArticleBomCode = dgdArticleP.SelectedItem as Win_com_ArticleBOM_Code_U;
            txtParentArticle.Tag = WinArticleBomCode.ArticleID.ToString();
            txtParentArticle.Text = WinArticleBomCode.Article.ToString();
        }

        // 하위품목 리스트 선택버튼 클릭 이벤트
        private void btnArticleC_Click(object sender, MouseButtonEventArgs e)
        {
            WinArticleBomCode = dgdArticleC.SelectedItem as Win_com_ArticleBOM_Code_U;
            txtArticle.Tag = WinArticleBomCode.ArticleID.ToString();
            txtArticle.Text = WinArticleBomCode.Article.ToString();
        }

        #region 기타 메서드 모음

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





        #endregion

        private void btnTest_Click(object sender, RoutedEventArgs e)
        {
            FillGrid2();
        }

        //소요량 숫자만 입력
        private void txtQty_PreviewTextInput(object sender, TextCompositionEventArgs e)
        {
            Lib.Instance.CheckIsNumeric((TextBox)sender, e);
        }

        //스크랩양 숫자만 입력
        private void txtLossQty_PreviewTextInput(object sender, TextCompositionEventArgs e)
        {
            Lib.Instance.CheckIsNumeric((TextBox)sender, e);
        }


        private void tlvItemList_PreviewMouseWheel(object sender, MouseWheelEventArgs e)
        {
            if (!e.Handled)
            {
                e.Handled = true;
                var evtArg = new MouseWheelEventArgs(e.MouseDevice, e.Timestamp, e.Delta)
                { RoutedEvent = MouseWheelEvent, Source = sender };

                var parent = ((Control)sender).Parent as UIElement;
                if (parent != null)
                    parent.RaiseEvent(evtArg);
            }
        }

        private void btnExcelUp_Click(object sender, RoutedEventArgs e)
        {
            // 인쇄 메서드
            ContextMenu menu = btnExcel.ContextMenu;
            menu.StaysOpen = true;
            menu.IsOpen = true;
        }

        #region BOM 인쇄 메서드

        // 미리보기
        private void menuSeeAhead_Click(object sender, RoutedEventArgs e)
        {
            if (ovcArticleBom.Count < 1)
            {
                MessageBox.Show("해당 자료가 존재하지 않습니다.");
                return;
            }

            msg.Show();
            msg.Topmost = true;
            msg.Refresh();

            Lib.Instance.Delay(1000);

            PrintWork(true);

            msg.Visibility = Visibility.Hidden;
        }

        // 바로 인쇄
        private void menuRightPrint_Click(object sender, RoutedEventArgs e)
        {
            if (ovcArticleBom.Count < 1)
            {
                MessageBox.Show("해당 자료가 존재하지 않습니다.");
                return;
            }
            DataStore.Instance.InsertLogByForm(this.GetType().Name, "P");
            msg.Show();
            msg.Topmost = true;
            msg.Refresh();

            Lib.Instance.Delay(1000);

            PrintWork(false);

            msg.Visibility = Visibility.Hidden;
        }
        // 닫기
        private void menuClose_Click(object sender, RoutedEventArgs e)
        {
            ContextMenu menu = btnExcel.ContextMenu;
            menu.StaysOpen = false;
            menu.IsOpen = false;
        }

        // 프린터 엑셀 작업
        // 실제 엑셀작업
        private void PrintWork(bool previewYN)
        {
            try
            {

                // 코드
                // 상호
                string ArticleID = "";
                // 사업장 주소
                string BuyerArticleNo = "";
                // 성명
                string Qty = "";
                //자회사명
                string UnitClssName = "";
                //자회사 대표


                var ArticleBom = dgdExcel.SelectedItem as Win_com_ArticleBOM_ItemList;

                if (ArticleBom != null)
                {

                    //CustomName = ArticleBom.CustomName;
                    //CustomChief = ArticleBom.CustomChief;
                    //CustomAddr = ArticleBom.CustomAddr1 + " " + ArticleBom.CustomAddr2 + " " + ArticleBom.CustomAddr3;
                    //kCompany = ArticleBom.kCompany;
                    //Chief = ArticleBom.Chief;
                }

                excelapp = new Microsoft.Office.Interop.Excel.Application();


                string MyBookPath = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetEntryAssembly().Location) + "\\Report\\비오엠엑셀.xls";
                workbook = excelapp.Workbooks.Add(MyBookPath);
                worksheet = workbook.Sheets["Form"];
                pastesheet = workbook.Sheets["Print"];

                //// 거래일자는 어떻게 해야 하는가
                //workrange = worksheet.get_Range("C5");
                //workrange.Value2 = DateTime.Today.ToString("yyyy.MM.dd");

                //// 상호
                //workrange = worksheet.get_Range("G6");
                //workrange.Value2 = CustomName;

                //// 사업장 주소
                //workrange = worksheet.get_Range("G8");
                //workrange.Value2 = CustomAddr;

                //// 성명
                //workrange = worksheet.get_Range("G10");
                //workrange.Value2 = CustomChief;

                //// 회사명
                //workrange = worksheet.get_Range("W8");
                //workrange.Value2 = kCompany;

                //workrange = worksheet.get_Range("AE8");
                //workrange.Value2 = Chief;

                // 페이지 계산 등
                int rowCount = ovcArticleBom.Count;
                int excelStartRow = 2;

                // 총 데이터를 입력할수 있는 갯수
                int totalDataInput = 350;

                //// 카피할 다음페이지 인덱스
                //int nextCopyLine = 380;


                int copyLine = 0;
                int Page = 0;
                int PageAll = (int)Math.Ceiling(1.0 * rowCount / totalDataInput);
                int DataCount = 0;




                // 총 금액 계산하기
                //double SumAmount = 0;

                for (int k = 0; k < PageAll; k++)
                {
                    Page++;
                    //copyLine = ((Page - 1) * (nextCopyLine - 1));
                    //copyLine = ((Page - 1) * 37);

                    int excelNum = 0;

                    // 기존에 있는 데이터 지우기 "A7", "W41"
                    worksheet.Range["A2", "P350"].EntireRow.ClearContents();


                    for (int i = DataCount; i < rowCount; i++)
                    {
                        //11
                        if (i == totalDataInput * Page)
                        {
                            break;
                        }

                        var OcArticleBom = ovcArticleBom[i];

                        int excelRow = excelStartRow + excelNum;

                        int excelRowStairTwo = excelStartRow + excelNum - 1;
                        int excelRowStairThree = excelStartRow + excelNum - 2;
                        int excelRowStairFour = excelStartRow + excelNum - 3;

                        int Temp = ConvertInt(OcArticleBom.ChildCnt);
                        int excelRowTest = excelStartRow + excelNum - Temp;



                        if (OcArticleBom != null)
                        {
                            if (OcArticleBom.LVL == 1)
                            {

                                if (excelNum < 10)
                                {
                                    workrange = worksheet.get_Range("A" + excelRow);
                                    workrange.Value2 = OcArticleBom.ArticleID;
                                    workrange.Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThin;
                                    workrange.Borders[Excel.XlBordersIndex.xlEdgeTop].Weight = Excel.XlBorderWeight.xlThin;

                                    workrange = worksheet.get_Range("B" + excelRow);
                                    workrange.Value2 = OcArticleBom.BuyerArticleNo;
                                    workrange.Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThin;
                                    workrange.Borders[Excel.XlBordersIndex.xlEdgeTop].Weight = Excel.XlBorderWeight.xlThin;
                                    workrange = worksheet.get_Range("C" + excelRow);
                                    workrange.Value2 = OcArticleBom.Qty;
                                    workrange.Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThin;
                                    workrange.Borders[Excel.XlBordersIndex.xlEdgeTop].Weight = Excel.XlBorderWeight.xlThin;
                                    workrange = worksheet.get_Range("D" + excelRow);
                                    workrange.Value2 = OcArticleBom.UnitClssName;
                                    workrange.Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThin;
                                    workrange.Borders[Excel.XlBordersIndex.xlEdgeTop].Weight = Excel.XlBorderWeight.xlThin;
                                    workrange.Borders[Excel.XlBordersIndex.xlEdgeRight].Weight = Excel.XlBorderWeight.xlThin;
                                }
                                else
                                {
                                    workrange = worksheet.get_Range("A" + (excelRow - 1));
                                    workrange.Value2 = OcArticleBom.ArticleID;
                                    workrange.Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThin;
                                    workrange.Borders[Excel.XlBordersIndex.xlEdgeTop].Weight = Excel.XlBorderWeight.xlThin;

                                    workrange = worksheet.get_Range("B" + (excelRow - 1));
                                    workrange.Value2 = OcArticleBom.BuyerArticleNo;
                                    workrange.Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThin;
                                    workrange.Borders[Excel.XlBordersIndex.xlEdgeTop].Weight = Excel.XlBorderWeight.xlThin;
                                    workrange = worksheet.get_Range("C" + (excelRow - 1));
                                    workrange.Value2 = OcArticleBom.Qty;
                                    workrange.Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThin;
                                    workrange.Borders[Excel.XlBordersIndex.xlEdgeTop].Weight = Excel.XlBorderWeight.xlThin;
                                    workrange = worksheet.get_Range("D" + (excelRow - 1));
                                    workrange.Value2 = OcArticleBom.UnitClssName;
                                    workrange.Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThin;
                                    workrange.Borders[Excel.XlBordersIndex.xlEdgeTop].Weight = Excel.XlBorderWeight.xlThin;
                                    workrange.Borders[Excel.XlBordersIndex.xlEdgeRight].Weight = Excel.XlBorderWeight.xlThin;
                                }
                            }
                            else if (OcArticleBom.LVL == 2)
                            {
                                if (excelNum < 10)
                                {

                                    workrange = worksheet.get_Range("E" + excelRowStairTwo);
                                    workrange.Value2 = OcArticleBom.ArticleID;
                                    workrange.Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThin;
                                    workrange.Borders[Excel.XlBordersIndex.xlEdgeTop].Weight = Excel.XlBorderWeight.xlThin;
                                    workrange.Borders[Excel.XlBordersIndex.xlEdgeLeft].Weight = Excel.XlBorderWeight.xlThin;

                                    workrange = worksheet.get_Range("F" + excelRowStairTwo);
                                    workrange.Value2 = OcArticleBom.BuyerArticleNo;
                                    workrange.Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThin;
                                    workrange.Borders[Excel.XlBordersIndex.xlEdgeTop].Weight = Excel.XlBorderWeight.xlThin;
                                    workrange = worksheet.get_Range("G" + excelRowStairTwo);
                                    workrange.Value2 = OcArticleBom.Qty;
                                    workrange.Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThin;
                                    workrange.Borders[Excel.XlBordersIndex.xlEdgeTop].Weight = Excel.XlBorderWeight.xlThin;
                                    workrange = worksheet.get_Range("H" + excelRowStairTwo);
                                    workrange.Value2 = OcArticleBom.UnitClssName;
                                    workrange.Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThin;
                                    workrange.Borders[Excel.XlBordersIndex.xlEdgeTop].Weight = Excel.XlBorderWeight.xlThin;
                                    workrange.Borders[Excel.XlBordersIndex.xlEdgeRight].Weight = Excel.XlBorderWeight.xlThin;
                                }

                                else
                                {
                                    workrange = worksheet.get_Range("E" + (excelRowStairTwo - 1));
                                    workrange.Value2 = OcArticleBom.ArticleID;
                                    workrange.Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThin;
                                    workrange.Borders[Excel.XlBordersIndex.xlEdgeTop].Weight = Excel.XlBorderWeight.xlThin;
                                    workrange.Borders[Excel.XlBordersIndex.xlEdgeLeft].Weight = Excel.XlBorderWeight.xlThin;

                                    workrange = worksheet.get_Range("F" + (excelRowStairTwo - 1));
                                    workrange.Value2 = OcArticleBom.BuyerArticleNo;
                                    workrange.Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThin;
                                    workrange.Borders[Excel.XlBordersIndex.xlEdgeTop].Weight = Excel.XlBorderWeight.xlThin;
                                    workrange = worksheet.get_Range("G" + (excelRowStairTwo - 1));
                                    workrange.Value2 = OcArticleBom.Qty;
                                    workrange.Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThin;
                                    workrange.Borders[Excel.XlBordersIndex.xlEdgeTop].Weight = Excel.XlBorderWeight.xlThin;
                                    workrange = worksheet.get_Range("H" + (excelRowStairTwo - 1));
                                    workrange.Value2 = OcArticleBom.UnitClssName;
                                    workrange.Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThin;
                                    workrange.Borders[Excel.XlBordersIndex.xlEdgeTop].Weight = Excel.XlBorderWeight.xlThin;
                                    workrange.Borders[Excel.XlBordersIndex.xlEdgeRight].Weight = Excel.XlBorderWeight.xlThin;
                                }


                            }
                            else if (OcArticleBom.LVL == 3)
                            {
                                if (excelNum < 10)
                                {

                                    workrange = worksheet.get_Range("I" + excelRowStairThree);
                                    workrange.Value2 = OcArticleBom.ArticleID;
                                    workrange.Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThin;
                                    workrange.Borders[Excel.XlBordersIndex.xlEdgeTop].Weight = Excel.XlBorderWeight.xlThin;
                                    workrange.Borders[Excel.XlBordersIndex.xlEdgeLeft].Weight = Excel.XlBorderWeight.xlThin;

                                    workrange = worksheet.get_Range("J" + excelRowStairThree);
                                    workrange.Value2 = OcArticleBom.BuyerArticleNo;
                                    workrange.Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThin;
                                    workrange.Borders[Excel.XlBordersIndex.xlEdgeTop].Weight = Excel.XlBorderWeight.xlThin;
                                    workrange = worksheet.get_Range("K" + excelRowStairThree);
                                    workrange.Value2 = OcArticleBom.Qty;
                                    workrange.Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThin;
                                    workrange.Borders[Excel.XlBordersIndex.xlEdgeTop].Weight = Excel.XlBorderWeight.xlThin;
                                    workrange = worksheet.get_Range("L" + excelRowStairThree);
                                    workrange.Value2 = OcArticleBom.UnitClssName;
                                    workrange.Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThin;
                                    workrange.Borders[Excel.XlBordersIndex.xlEdgeTop].Weight = Excel.XlBorderWeight.xlThin;
                                    workrange.Borders[Excel.XlBordersIndex.xlEdgeRight].Weight = Excel.XlBorderWeight.xlThin;
                                }

                                else
                                {
                                    workrange = worksheet.get_Range("I" + (excelRowStairThree - 1));
                                    workrange.Value2 = OcArticleBom.ArticleID;
                                    workrange.Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThin;
                                    workrange.Borders[Excel.XlBordersIndex.xlEdgeTop].Weight = Excel.XlBorderWeight.xlThin;
                                    workrange.Borders[Excel.XlBordersIndex.xlEdgeLeft].Weight = Excel.XlBorderWeight.xlThin;

                                    workrange = worksheet.get_Range("J" + (excelRowStairThree - 1));
                                    workrange.Value2 = OcArticleBom.BuyerArticleNo;
                                    workrange.Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThin;
                                    workrange.Borders[Excel.XlBordersIndex.xlEdgeTop].Weight = Excel.XlBorderWeight.xlThin;
                                    workrange = worksheet.get_Range("K" + (excelRowStairThree - 1));
                                    workrange.Value2 = OcArticleBom.Qty;
                                    workrange.Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThin;
                                    workrange.Borders[Excel.XlBordersIndex.xlEdgeTop].Weight = Excel.XlBorderWeight.xlThin;
                                    workrange = worksheet.get_Range("L" + (excelRowStairThree - 1));
                                    workrange.Value2 = OcArticleBom.UnitClssName;
                                    workrange.Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThin;
                                    workrange.Borders[Excel.XlBordersIndex.xlEdgeTop].Weight = Excel.XlBorderWeight.xlThin;
                                    workrange.Borders[Excel.XlBordersIndex.xlEdgeRight].Weight = Excel.XlBorderWeight.xlThin;
                                }
                            }
                            else if (OcArticleBom.LVL == 4)
                            {
                                if (excelNum < 10)
                                {

                                    workrange = worksheet.get_Range("M" + excelRowStairFour);
                                    workrange.Value2 = OcArticleBom.ArticleID;
                                    workrange.Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThin;
                                    workrange.Borders[Excel.XlBordersIndex.xlEdgeTop].Weight = Excel.XlBorderWeight.xlThin;
                                    workrange.Borders[Excel.XlBordersIndex.xlEdgeLeft].Weight = Excel.XlBorderWeight.xlThin;

                                    workrange = worksheet.get_Range("N" + excelRowStairFour);
                                    workrange.Value2 = OcArticleBom.BuyerArticleNo;
                                    workrange.Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThin;
                                    workrange.Borders[Excel.XlBordersIndex.xlEdgeTop].Weight = Excel.XlBorderWeight.xlThin;
                                    workrange = worksheet.get_Range("O" + excelRowStairFour);
                                    workrange.Value2 = OcArticleBom.Qty;
                                    workrange.Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThin;
                                    workrange.Borders[Excel.XlBordersIndex.xlEdgeTop].Weight = Excel.XlBorderWeight.xlThin;
                                    workrange = worksheet.get_Range("P" + excelRowStairFour);
                                    workrange.Value2 = OcArticleBom.UnitClssName;
                                    workrange.Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThin;
                                    workrange.Borders[Excel.XlBordersIndex.xlEdgeTop].Weight = Excel.XlBorderWeight.xlThin;
                                    workrange.Borders[Excel.XlBordersIndex.xlEdgeRight].Weight = Excel.XlBorderWeight.xlThin;

                                }
                                else
                                {
                                    workrange = worksheet.get_Range("M" + (excelRowStairFour - 1));
                                    workrange.Value2 = OcArticleBom.ArticleID;
                                    workrange.Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThin;
                                    workrange.Borders[Excel.XlBordersIndex.xlEdgeTop].Weight = Excel.XlBorderWeight.xlThin;
                                    workrange.Borders[Excel.XlBordersIndex.xlEdgeLeft].Weight = Excel.XlBorderWeight.xlThin;

                                    workrange = worksheet.get_Range("N" + (excelRowStairFour - 1));
                                    workrange.Value2 = OcArticleBom.BuyerArticleNo;
                                    workrange.Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThin;
                                    workrange.Borders[Excel.XlBordersIndex.xlEdgeTop].Weight = Excel.XlBorderWeight.xlThin;
                                    workrange = worksheet.get_Range("O" + (excelRowStairFour - 1));
                                    workrange.Value2 = OcArticleBom.Qty;
                                    workrange.Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThin;
                                    workrange.Borders[Excel.XlBordersIndex.xlEdgeTop].Weight = Excel.XlBorderWeight.xlThin;
                                    workrange = worksheet.get_Range("P" + (excelRowStairFour - 1));
                                    workrange.Value2 = OcArticleBom.UnitClssName;
                                    workrange.Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThin;
                                    workrange.Borders[Excel.XlBordersIndex.xlEdgeTop].Weight = Excel.XlBorderWeight.xlThin;
                                    workrange.Borders[Excel.XlBordersIndex.xlEdgeRight].Weight = Excel.XlBorderWeight.xlThin;
                                }


                            }

                        }


                        //workrange.Borders[Excel.XlBordersIndex.xlEdgeTop].Weight = Excel.XlBorderWeight.xlThin;
                        //workrange.Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThin;

                        excelNum++;
                        DataCount = i;
                    }


                    // 2장 이상 넘어가면 페이지 넘버 입력
                    //if (PageAll > 1)
                    //{
                    //    pastesheet.PageSetup.CenterFooter = "&P / &N";
                    //}

                    //Form 시트 내용 Print 시트에 복사 붙여넣기
                    worksheet.Select();
                    worksheet.UsedRange.EntireRow.Copy();
                    pastesheet.Select();
                    workrange = pastesheet.Cells[copyLine + 1, 1];
                    workrange.Select();
                    pastesheet.Paste();


                    DataCount++;
                }

                //// 총금액 입력하기 : 10, 50, 90
                //for (int i = 0; i < PageAll; i++)
                //{
                //    int sumAmount_Index = 10 + (40 * i);

                //    workrange = pastesheet.get_Range("E" + sumAmount_Index);
                //    workrange.Value2 = SumAmount;
                //}

                pastesheet.UsedRange.EntireRow.Select();

                //
                excelapp.Visible = true;
                msg.Hide();

                if (previewYN == true)
                {
                    pastesheet.PrintPreview();
                }
                else
                {
                    pastesheet.PrintOutEx();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally
            {
                // Clean up 백그라운드에서 엑셀을 지우자 - 달달

                ReleaseExcelObject(workbook);
                ReleaseExcelObject(worksheet);
                ReleaseExcelObject(pastesheet);
                ReleaseExcelObject(workrange);
                ReleaseExcelObject(excelapp);


            }
        }
        //엑셀 백그라운드 증발 - 달달
        private static void ReleaseExcelObject(object obj)
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
                GC.Collect();
            }
        }



        #endregion // 거래명세서 인쇄 메서드

        //오른쪽 상위품번
        private void txtParentBuyArticle_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                MainWindow.pf.ReturnCodeGLS(txtParentBuyArticle, 76, "");
                //MainWindow.pf.ReturnCode(txtParentArticle, (int)Defind_CodeFind.DCF_BuyerArticleNo, "");

                PChildArticle = "P";

                if (txtParentBuyArticle.Tag != null)
                {
                    getArticleInfo(txtParentBuyArticle.Tag.ToString());
                }

                PChildArticle = "";

            }
        }
        //오른쪽 상위품번

        private void btnPfParentBuyArticle_Click(object sender, RoutedEventArgs e)
        {
            PChildArticle = "P";

            MainWindow.pf.ReturnCodeGLS(txtParentBuyArticle, 76, "");
            //MainWindow.pf.ReturnCode(txtParentArticle, (int)Defind_CodeFind.DCF_BuyerArticleNo, "");

            if (txtParentBuyArticle.Tag != null)
            {
                getArticleInfo(txtParentBuyArticle.Tag.ToString());
            }
            PChildArticle = "";

        }


        //우측 하위품번 
        private void txtBuyArticle_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                MainWindow.pf.ReturnCodeGLS(txtBuyArticle, 76, "");
                //MainWindow.pf.ReturnCode(txtArticle, (int)Defind_CodeFind.DCF_BuyerArticleNo, "");
                PChildArticle = "C";

                if (txtBuyArticle.Tag != null)
                {
                    getArticleInfo(txtBuyArticle.Tag.ToString());
                }

                PChildArticle = "";

            }
        }
        //우측 하위품번 버튼
        private void btnPfBuyArticle_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                MainWindow.pf.ReturnCodeGLS(txtBuyArticle, 76, "");
                //MainWindow.pf.ReturnCode(txtArticle, (int)Defind_CodeFind.DCF_BuyerArticleNo, "");
                PChildArticle = "C";

                if (txtBuyArticle.Tag != null)
                {
                    getArticleInfo(txtBuyArticle.Tag.ToString());
                }

                PChildArticle = "";

            }
            catch (Exception ex)
            {
                MessageBox.Show("오류 발생, 오류 내용 : " + ex.ToString());
            }
        }
    }

    class Win_com_ArticleBOM_ItemList : BaseView
    {
        public override string ToString()
        {
            return (this.ReportAllProperties());
        }

        public int Num { get; set; }

        public int LVL { get; set; }
        public string ArticleID { get; set; }
        public string PARENTArticleID { get; set; }
        public string ord { get; set; }
        public string Article { get; set; }

        public string ArticleP { get; set; }
        public string ArticleGrpID { get; set; }
        public string ChildCnt { get; set; }
        public string Qty { get; set; }
        public string FromDate { get; set; }

        public string ToDate { get; set; }
        public string UnitClss { get; set; }
        public string PcntClss { get; set; }
        public string ParentArticleIDS { get; set; }
        public string LvlPad { get; set; }

        public string LossQty { get; set; }
        public string LossPcntClss { get; set; }
        public string ScraptRate { get; set; }
        public string UnitClssName { get; set; }
        public string Weight { get; set; }

        public string ChildBuyerArticleNO { get; set; }
        public string ParentBuyerArticleNO { get; set; }

        public string FromDate_CV { get; set; }
        public string ToDate_CV { get; set; }
        public string FirstColumnCV { get; set; }
        public string UseYN { get; set; }

        public string BuyerArticleNo { get; set; }

        public string FirstColumnCV2 { get; set; }

        public int LVL1 { get; set; }
        public int LVL2 { get; set; }
        public int LVL3 { get; set; }
        public int LVL4 { get; set; }

        //public string SubArticleID { get; set; }
        //public string SubArticleName { get; set; }
        //public string chkSub { get; set; }
        //public string chkSubMe { get; set; }


    }

    class Win_com_ArticleBOM_Code_U
    {
        public override string ToString()
        {
            return (this.ReportAllProperties());
        }

        public string Num { get; set; }
        public string Article { get; set; }
        public string ArticleID { get; set; }
        public string ArticleGrpID { get; set; }
        public string Qty { get; set; }
        public string UnitClss { get; set; }
        public string UnitClssName { get; set; }
        public string Weight { get; set; }
        public string ArticleGrp { get; set; }
        public string BuyerArticleNo { get; set; }
        public string BuySaleMainYN { get; set; }
        public string CoatingSpec { get; set; }
        public string Comments { get; set; }
        public string DyeingID { get; set; }
        public string FTAMgrYN { get; set; }
        public string HSCODE { get; set; }
        public string ImageName { get; set; }
        public string LabelPrintYN { get; set; }
        public string NeedStockQty { get; set; }
        public string OutUnitPrice { get; set; }
        public string PART_ATTR { get; set; }
        public string PatternID { get; set; }
        public string Process { get; set; }
        public string ProcessID { get; set; }
        public string QtyPerBox { get; set; }
        public string Spec { get; set; }
        public string StuffWidth { get; set; }
        public string SupplyType { get; set; }
        public string SupplyTypeName { get; set; }
        public string Thread { get; set; }
        public string ThreadID { get; set; }
        public string Unitprice { get; set; }
        public string UnitPriceClss { get; set; }
        public string UseClss { get; set; }
        public string UseingType { get; set; }
        public string UseingTypeName { get; set; }

    }

    class BOMArticleInfo : BaseView
    {
        public string ArticleGrpID { get; set; }
        public string UnitPrice { get; set; }
        public string UnitPriceClss { get; set; }
        public string UnitClss { get; set; }
        public string PartGBNID { get; set; }
        public string ProductGrpID { get; set; }
        public string Article { get; set; }
        public string BuyerArticleNo { get; set; }
        public string ArticleID { get; set; }
    }
}

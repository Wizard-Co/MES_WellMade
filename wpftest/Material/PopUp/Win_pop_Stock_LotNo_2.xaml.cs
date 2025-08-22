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
using System.Windows.Shapes;

namespace WizMes_WellMade.PopUp
{
    /// <summary>
    /// RheoChoice.xaml에 대한 상호 작용 논리
    /// </summary>
    public partial class Win_pop_Stock_LotNo_2 : Window
    {
        int rowNum = 0;

        public string ArticleID = "";
        public string Article = "";
        public string LotID = "";

        public string BuyerArticleNo = "";
        public string ArticleGrp = "";
        public string UnitClssName = "";
        public string StockQty = "";

        public string date = "";

        Lib lib = new Lib();

        public Win_mtr_LotStockControl_U StockControl = new Win_mtr_LotStockControl_U();

        public List<Win_mtr_LotStockControl_U_CodeView> lstLotClonePop = new List<Win_mtr_LotStockControl_U_CodeView>();
        
        

        public Win_pop_Stock_LotNo_2()
        {
            InitializeComponent();
        }

        public Win_pop_Stock_LotNo_2(List<Win_mtr_LotStockControl_U_CodeView> lstLotClonePop)
        {
            InitializeComponent();

            this.lstLotClonePop = lstLotClonePop;
        }

        public Win_pop_Stock_LotNo_2(string ArticleID, string Article, string LotID, string BuyerArticleNo, string ArticleGrp, string UnitClssName, string StockQty)
        {
            InitializeComponent();

            this.ArticleID = ArticleID;
            this.Article = Article;
            this.LotID = LotID;

            this.BuyerArticleNo = BuyerArticleNo;
            this.ArticleGrp = ArticleGrp;
            this.UnitClssName = UnitClssName;
            this.StockQty = StockQty;
        }

        // 콤보박스셋팅
        private void ComboBoxSetting()
        {
            cboArticleGroup.Items.Clear();
            cboWareHouse.Items.Clear();

            ObservableCollection<CodeView> cbArticleGroup = ComboBoxUtil.Instance.Gf_DB_MT_sArticleGrp();
            ObservableCollection<CodeView> cbWareHouse = ComboBoxUtil.Instance.Gf_DB_CM_GetComCodeDataset(null, "LOC", "Y", "", "");

            this.cboArticleGroup.ItemsSource = cbArticleGroup;
            this.cboArticleGroup.DisplayMemberPath = "code_name";
            this.cboArticleGroup.SelectedValuePath = "code_id";
            this.cboArticleGroup.SelectedIndex = 0;


            

            this.cboWareHouse.ItemsSource = cbWareHouse;
            this.cboWareHouse.DisplayMemberPath = "code_name";
            this.cboWareHouse.SelectedValuePath = "code_id";
            this.cboWareHouse.SelectedIndex = 0;

        }

        private void MoveSub_Loaded(object sender, RoutedEventArgs e)
        {
            //dtpAdjustDate.SelectedDate = DateTime.Today;

            
            ComboBoxSetting();

            if (!LotID.Trim().Equals(""))
            {
                chkLotIDSrh.IsChecked = true;
                txtLotIDSrh.Text = LotID;
            }
            else if (!ArticleID.Trim().Equals(""))
            {
                txtArticleSrh.Text = Article;

                //FillGrid_ArticleID(ArticleID);

                return;
            }
            else if (!Article.Trim().Equals(""))
            {
                chkArticle.IsChecked = true;
                txtArticleSrh.Text = Article;
            }

            FillGrid();

            //dtpAdjustDate.SelectedDate = DateTime.Today;
        }

        #region 주요 버튼 이벤트 - 확인, 닫기, 검색

        public List<Win_mtr_LotStockControl_U_CodeView> lstLotStock = new List<Win_mtr_LotStockControl_U_CodeView>();

        //확인
        private void btnConfirm_Click(object sender, RoutedEventArgs e)
        {
            for(int i = 0 ;  i < dgdMain.Items.Count; i++)
            {
                var main = dgdMain.Items[i] as Win_mtr_LotStockControl_U_CodeView;

                if(main != null && main.Chk == true)
                {
                    lstLotStock.Add(main);

                }

            }

            this.DialogResult = true;
            
        }

        //닫기
        private void btnCancel_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }

        //검색
        private void btnSearch_Click(object sender, RoutedEventArgs e)
        {
            //검색버튼 비활성화
            btnSearch.IsEnabled = false;

            Dispatcher.BeginInvoke(new Action(() =>

            {
                Thread.Sleep(2000);

                //로직
                re_Search(rowNum);

            }), System.Windows.Threading.DispatcherPriority.Background);



            Dispatcher.BeginInvoke(new Action(() =>

            {
                btnSearch.IsEnabled = true;

            }), System.Windows.Threading.DispatcherPriority.Background);
        }

        #endregion // 주요 버튼 이벤트


        #region Header 부분 - 검색조건


        // 품명
        private void lblArticleSrh_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            if (chkArticle.IsChecked == true)
            {
                chkArticle.IsChecked = false;
            }
            else
            {
                chkArticle.IsChecked = true;
            }
        }
        private void chkArticleSrh_Checked(object sender, RoutedEventArgs e)
        {
            chkArticle.IsChecked = true;
            txtArticleSrh.IsEnabled = true;
        }
        private void chkArticleSrh_Unchecked(object sender, RoutedEventArgs e)
        {
            chkArticle.IsChecked = false;
            txtArticleSrh.IsEnabled = false;
        }

        // LotID - 바코드 검색
        private void lblLotIDSrh_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            if (chkLotIDSrh.IsChecked == true)
            {
                chkLotIDSrh.IsChecked = false;
            }
            else
            {
                chkLotIDSrh.IsChecked = true;
            }
        }

        private void chkLotIDSrh_Checked(object sender, RoutedEventArgs e)
        {
            chkLotIDSrh.IsChecked = true;
            txtLotIDSrh.IsEnabled = true;
        }

        private void chkLotIDSrh_Unchecked(object sender, RoutedEventArgs e)
        {
            chkLotIDSrh.IsChecked = false;
            txtLotIDSrh.IsEnabled = false;
        }

        // 재고 0 포함 라벨 버튼 이벤트 2021-06-29
        private void lblnIncZeroQty_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            if (chknIncZeroQty.IsChecked == true)
            {
                chknIncZeroQty.IsChecked = false;
            }
            else
            {
                chknIncZeroQty.IsChecked = true;
            }
        }
        //재고 0 포함 체크
        private void chknIncZeroQty_Checked(object sender, RoutedEventArgs e)
        {
            chknIncZeroQty.IsChecked = true;
        }
        //재고 0 포함 체크해제
        private void chknIncZeroQty_UnChecked(object sender, RoutedEventArgs e)
        {
            chknIncZeroQty.IsChecked = false;
        }

        #endregion // Header 부분 - 검색조건

        #region Header 부분 - 검색조건 : 바코드 검색 → 바코드 비워주기 (다음 바코드를 바로 입력할 수 있도록)

        //Lot ID
        private void txtLotID_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                e.Handled = true;
                FillGrid();

                txtLotIDSrh.Text = "";

            }
        }

        #endregion

        #region 주요 메서드 모음

        private void re_Search(int rowNum)
        {
            FillGrid();

            if (dgdMain.Items.Count > 0)
            {
                dgdMain.SelectedIndex = rowNum;
            }
            else
            {
                this.DataContext = null;
                MessageBox.Show("조회된 데이터가 없습니다.");
                return;
            }
        }

        #region 조회

        private void FillGrid()
        {
            if (dgdMain.Items.Count > 0)
            {
                dgdMain.Items.Clear();
            }

            try
            {
                Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                sqlParameter.Clear();

                //sqlParameter.Add("ChkArticleID", 0);
                //sqlParameter.Add("ArticleID", "");
                //sqlParameter.Add("ChkArticle", chkArticleSrh.IsChecked == true ? 1 : 0);
                //sqlParameter.Add("Article", chkArticleSrh.IsChecked == true && !txtArticleSrh.Text.Trim().Equals("") ? txtArticleSrh.Text : "");
                //sqlParameter.Add("ChkLotID", chkLotIDSrh.IsChecked == true ? 1 : 0);
                //sqlParameter.Add("LotID", chkLotIDSrh.IsChecked == true && !txtLotIDSrh.Text.Trim().Equals("") ? txtLotIDSrh.Text : "");
                //sqlParameter.Add("ArticleGrpID", chkArticleGroup.IsChecked == true && cboArticleGroup.SelectedValue != null ? cboArticleGroup.SelectedValue.ToString() : "");


                sqlParameter.Add("sDate", date);
                sqlParameter.Add("ChkArticle", chkArticle.IsChecked == true ? 1 : 0);
                sqlParameter.Add("ArticleID", chkArticle.IsChecked == true && txtArticleSrh.Text != null ? txtArticleSrh.Text.ToString() : "");
                sqlParameter.Add("ChkLotID", chkLotIDSrh.IsChecked == true ? 1 : 0);
                sqlParameter.Add("LotID", chkLotIDSrh.IsChecked == true && txtLotIDSrh.Text.Trim().Length > 0 ? txtLotIDSrh.Text.Trim() : "");
                sqlParameter.Add("ArticleGrpID", chkArticleGroup.IsChecked == true && cboArticleGroup.SelectedValue != null ? cboArticleGroup.SelectedValue.ToString() : ""); //제품그룹
                sqlParameter.Add("sToLocID", chkToLocSrh.IsChecked == true ? (cboWareHouse.SelectedValue != null ? cboWareHouse.SelectedValue.ToString() : "") : ""); // 창고
                sqlParameter.Add("nIncZeroQty", chknIncZeroQty.IsChecked == true ? 1 : 0); //재고 0 포함, 기본값은 재고 0은 포함 안함 2021-06-29

                //DataSet ds = DataStore.Instance.ProcedureToDataSet("xp_mtr_StockLotID_WPF", sqlParameter, false);
                DataSet ds = DataStore.Instance.ProcedureToDataSet("xp_sbStock_sLotStockControl_LotStock_2", sqlParameter, false);

                #region 봉인
                //if (ds != null && ds.Tables.Count > 0)
                //{
                //    DataTable dt = ds.Tables[0];

                //    if (dt.Rows.Count > 0)
                //    {
                //        DataRowCollection drc = dt.Rows;

                //        int i = 0;

                //        foreach (DataRow dr in drc)
                //        {
                //            i++;

                //            var Main = new Win_mtr_StockControl_U_Stuffin()
                //            {
                //                Num = i.ToString(),
                //              //  StuffDate = dr["StuffDate"].ToString(),
                //               // StuffDate_CV = DatePickerFormat(dr["StuffDate"].ToString()),
                //                BuyerArticleNo = dr["BuyerArticleNo"].ToString(),
                //                Article = dr["Article"].ToString(),
                //                ArticleID = dr["ArticleID"].ToString(),
                //                LotID = dr["LotID"].ToString(),
                //                UnitClss = dr["UnitClss"].ToString(),

                //                UnitClssName = dr["UnitClssName"].ToString(),
                //                ArticleGrpID = dr["ArticleGrpID"].ToString(),
                //                ArticleGrp = dr["ArticleGrp"].ToString(),
                //                TOLocID = dr["TOLocID"].ToString(),
                //                ToLocName = dr["ToLocName"].ToString(),
                //                Qty = stringFormatN0(dr["Qty"]), //현재고는 어떻게 구하니?

                //            };

                //            dgdMain.Items.Add(Main);

                //        }

                //        tblCount.Text = "▶검색개수 : " + i + "건";
                //    }
                //}
                #endregion
                if (ds != null && ds.Tables.Count > 0)
                {
                    DataTable dt = ds.Tables[0];

                    if (dt.Rows.Count == 0)
                    {
                        MessageBox.Show("조회된 데이터가 없습니다.");
                    }
                    else
                    {
                        int index = 0;
                        DataRowCollection drc = dt.Rows;

                        foreach (DataRow dr in drc)
                        {
                            index++;
                            var NowStockData = new Win_mtr_LotStockControl_U_CodeView
                            {
                                Num = index,
                                ArticleID = dr["ArticleID"].ToString(),
                                LotID = dr["LotID"].ToString(),
                                LotName = dr["LotID"].ToString(),
                                UnitClss = dr["UnitClss"].ToString(),
                                StuffinQty = stringFormatN0(dr["StuffinQty"]),
                                OutQty = stringFormatN0(dr["Outqty"]),
                                StockQty = stringFormatN0(dr["StockQty"]),                               
                                Article = dr["Article"].ToString(),
                                UnitClssName = dr["UnitClssName"].ToString(),
                                BuyerArticleNo = dr["BuyerArticleNo"].ToString(),


                                ArticleGrpID = dr["ArticleGrpID"].ToString(),
                                ArticleGrp = dr["ArticleGrp"].ToString(),
                                TOLocID = dr["TOLocID"].ToString(),
                                ToLocName = dr["ToLocName"].ToString(),
                                LastDate = dr["LastDate"].ToString(),
                                UDFlag = true,

                            };

                            if (lstLotClonePop.Count > 0)
                            {
                                for (int i = 0; i < lstLotClonePop.Count; i++)
                                {
                                    if (NowStockData.LotID.Equals(lstLotClonePop[i].LotID.Trim()) && NowStockData.ArticleID.Equals(lstLotClonePop[i].ArticleID.Trim())) //2021-06-26 LOTID는 같고 ArticleID는 다를 경우 위해 수정
                                    {
                                        NowStockData.StockQty = lstLotClonePop[i].StockQty;
                                    }
                                }
                            }

                            NowStockData.ControlQty = "0";

                            dgdMain.Items.Add(NowStockData);
                        }
                        tblCount.Text = "▶ 검색결과 : " + index + "건";

                    }
                }

            }
            catch (Exception ee)
            {


                MessageBox.Show("조회 오류 : " + ee.Message);
            }
            finally
            {
                DataStore.Instance.CloseConnection();
            }
        }

        #endregion

        #region 조회 - ArticleID 로!

        //private void FillGrid_ArticleID(string ArticleID)
        //{
        //    if (dgdMain.Items.Count > 0)
        //    {
        //        dgdMain.Items.Clear();
        //    }

        //    try
        //    {
        //        Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
        //        sqlParameter.Clear();


        //        sqlParameter.Add("ChkArticleID", 1);
        //        sqlParameter.Add("ArticleID", ArticleID);

        //        sqlParameter.Add("ChkArticle", 0);
        //        sqlParameter.Add("Article", "");

        //        sqlParameter.Add("ChkLotID", 0);
        //        sqlParameter.Add("LotID", "");

        //        DataSet ds = DataStore.Instance.ProcedureToDataSet("xp_mtr_StockLotID_WPF", sqlParameter, false);

        //        if (ds != null && ds.Tables.Count > 0)
        //        {
        //            DataTable dt = ds.Tables[0];

        //            if (dt.Rows.Count > 0)
        //            {
        //                DataRowCollection drc = dt.Rows;

        //                int i = 0;

        //                foreach (DataRow dr in drc)
        //                {
        //                    i++;

        //                    var Main = new Win_mtr_StockControl_U_Stuffin()
        //                    {
        //                        Num = i.ToString(),

        //                        BuyerArticleNo = dr["BuyerArticleNo"].ToString(),

        //                        Article = dr["Article"].ToString(),
        //                        ArticleID = dr["ArticleID"].ToString(),
        //                        LotID = dr["LotID"].ToString(),
        //                        Qty = stringFormatN0(dr["Qty"]),

        //                    };

        //                    dgdMain.Items.Add(Main);

        //                }

        //                tblCount.Text = "▶검색개수 : " + i + "건";
        //            }
        //        }
        //    }
        //    catch (Exception ee)
        //    {
        //        MessageBox.Show("조회 오류 : " + ee.Message);
        //    }
        //    finally
        //    {
        //        DataStore.Instance.CloseConnection();
        //    }
        //}

        #endregion

        #endregion

        #region 유효성 검사

        private bool CheckData()
        {
            bool flag = true;

            return flag;
        }

        #endregion

        #region 데이터 그리드 체크박스 이벤트

        // 팝업창 체크박스 이벤트
        private void CHK_Click_Sub(object sender, RoutedEventArgs e)
        {
            //CheckBox chkSender = sender as CheckBox;
            //var MoveSub = chkSender.DataContext as Win_mtr_Move_U_CodeViewSub;

            //if (MoveSub != null)
            //{
            //    if (chkSender.IsChecked == true)
            //    {
            //        MoveSub.Chk = true;
            //        MoveSub.FontColor = true;

            //        if (ovcMoveSub.Contains(MoveSub) == false)
            //        {
            //            ovcMoveSub.Add(MoveSub);
            //        }
            //    }
            //    else
            //    {
            //        MoveSub.Chk = false;
            //        MoveSub.FontColor = false;

            //        if (ovcMoveSub.Contains(MoveSub) == true)
            //        {
            //            ovcMoveSub.Remove(MoveSub);
            //        }
            //    }
            //}
        }

        #endregion // 데이터 그리드 체크박스 이벤트

        #region 전체 선택 체크박스 이벤트

        // 전체 선택 체크박스 체크 이벤트
        private void AllCheck_Checked(object sender, RoutedEventArgs e)
        {
            //ovcMoveSub.Clear();

            //if (dgdMain.Visibility == Visibility.Visible)
            //{
            //    for (int i = 0; i < dgdMain.Items.Count; i++)
            //    {
            //        var MoveSub = dgdMain.Items[i] as Win_mtr_Move_U_CodeViewSub;
            //        MoveSub.Chk = true;
            //        MoveSub.FontColor = true;

            //        ovcMoveSub.Add(MoveSub);
            //    }
            //}
        }

        // 전체 선택 체크박스 언체크 이벤트
        private void AllCheck_Unchecked(object sender, RoutedEventArgs e)
        {
            //ovcMoveSub.Clear();

            //if (dgdMain.Visibility == Visibility.Visible)
            //{
            //    for (int i = 0; i < dgdMain.Items.Count; i++)
            //    {
            //        var MoveSub = dgdMain.Items[i] as Win_mtr_Move_U_CodeViewSub;
            //        MoveSub.Chk = false;
            //        MoveSub.FontColor = false;
            //    }
            //}
        }

        #endregion // 전체 선택 체크박스 이벤트

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






        #endregion // 기타 메서드

        // 메인 그리드 더블클릭시 선택한걸로!!
        private void dgdMain_PreviewMouseDown(object sender, MouseButtonEventArgs e)
        {
            //if (e.ClickCount == 2)
            //{
            //    btnConfirm_Click(null, null);
            //}
        }

        private void chkReq_Click(object sender, RoutedEventArgs e)
        {
            CheckBox chkSender = sender as CheckBox;
            var LotStock = chkSender.DataContext as Win_mtr_LotStockControl_U_CodeView;

            if (LotStock != null)
            {
                if (chkSender.IsChecked == true)
                {
                    LotStock.Chk = true;
                }
                else
                {
                    LotStock.Chk = false;
                }

            }
        }

        //제품그룹
        private void chkArticleGroup_Click(object sender, RoutedEventArgs e)
        {
            if (chkArticleGroup.IsChecked == true)
            {
                cboArticleGroup.IsEnabled = true;
                cboArticleGroup.Focus();
            }
            else
            {
                cboArticleGroup.IsEnabled = false;
            }
        }

        //제품그룹
        private void chkArticleGroup_Click(object sender, MouseButtonEventArgs e)
        {
            if (chkArticleGroup.IsChecked == true)
            {
                chkArticleGroup.IsChecked = false;
                cboArticleGroup.IsEnabled = false;

            }
            else
            {
                chkArticleGroup.IsChecked = true;
                cboArticleGroup.IsEnabled = true;
                cboArticleGroup.Focus();
            }
        }
        //2021-05-29(2021-07-12 해제도 추가)
        private void BtnAllChoice_Click(object sender, RoutedEventArgs e)
        {
            if (dgdMain.Items.Count > 0)
            {
                foreach (Win_mtr_LotStockControl_U_CodeView Silsadata in dgdMain.Items)
                {
                   
                    if (Silsadata != null && Silsadata.Chk == false)
                    {
                        Silsadata.Chk = true;
                    }
                    else
                    {
                        Silsadata.Chk = false;
                    }

                }

                dgdMain.Items.Refresh();
            }
        }

        private void dtpAdjustDate_PreviewKeyUp(object sender, KeyEventArgs e)
        {

        }

        private void dtpAdjustDate_CalendarClosed(object sender, RoutedEventArgs e)
        {

        }

       // 창고체크박스
        private void chkToLocSrh_Click(object sender, RoutedEventArgs e)
        {
            if (chkToLocSrh.IsChecked == true)
            {
                cboWareHouse.IsEnabled = true;
                cboWareHouse.Focus();
            }
            else
            {
                cboWareHouse.IsEnabled = false;
            }
        }

        // 창고
        private void chkToLocSrh_Click(object sender, MouseButtonEventArgs e)
        {
            if (chkToLocSrh.IsChecked == true)
            {
                chkToLocSrh.IsChecked = false;
                cboWareHouse.IsEnabled = false;

            }
            else
            {
                chkToLocSrh.IsChecked = true;
                cboWareHouse.IsEnabled = true;
                cboWareHouse.Focus();
            }
        }

    }


}

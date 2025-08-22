using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Data;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;
using WizMes_WellMade.PopUp;

/**************************************************************************************************
'** 프로그램명 : Win_pop_OutwareReq
'** 설명       : 출고지시 대상조회
'** 작성일자   : 
'** 작성자     : 장시영
'**------------------------------------------------------------------------------------------------
'**************************************************************************************************
' 변경일자  , 변경자, 요청자    , 요구사항ID      , 요청 및 작업내용
'**************************************************************************************************
' 2023.03.29, 장시영, 삼익SDT에서 가져옴
'**************************************************************************************************/

namespace WizMes_WellMade.PopUp
{
    /// <summary>
    /// RheoChoice.xaml에 대한 상호 작용 논리
    /// </summary>
    public partial class Win_pop_OutwareReq : Window
    {
        Lib lib = new Lib();

        bool dounleClick = false;
        int rowNum = 0;

        string stDate = string.Empty;
        string stTime = string.Empty;

        public int chkDate { private get; set; }
        public string startDate { private get; set; }
        public string endDate { private get; set; }
        public int chkCustomID { private get; set; }
        public string customID { private get; set; }

        private List<Win_ord_OutwareReqSub_U_View> listReq = new List<Win_ord_OutwareReqSub_U_View>();

        public Win_pop_OutwareReq()
        {
            InitializeComponent();
        }

        private void OutwareReq_Loaded(object sender, RoutedEventArgs e)
        {
            stDate = DateTime.Now.ToString("yyyyMMdd");
            stTime = DateTime.Now.ToString("HHmm");
            DataStore.Instance.InsertLogByFormS(this.GetType().Name, stDate, stTime, "S");

            dtpSDate.Text = DateTime.Now.ToString("yyyy-MM-dd");
            dtpEDate.Text = DateTime.Now.ToString("yyyy-MM-dd");   // 오늘 날짜 자동 반영

            re_Search();
        }

        #region 주요 버튼 이벤트 - 확인, 닫기, 검색

        //확인        
        private void btnConfirm_Click(object sender, RoutedEventArgs e)
        {
            listReq.Clear();

            if (dounleClick)
            {
                var main = dgdMain.SelectedItem as Win_ord_OutwareReqSub_U_View;
                listReq.Add(main);
            }
            else
            {
                for (int i = 0; i < dgdMain.Items.Count; i++)
                {
                    var main = dgdMain.Items[i] as Win_ord_OutwareReqSub_U_View;
                    if (main != null && main.Chk == true)
                        listReq.Add(main);
                }
            }

            DataStore.Instance.InsertLogByFormS(this.GetType().Name, stDate, stTime, "E");
            dounleClick = false;
            this.DialogResult = true;
        }

        //닫기
        private void btnCancel_Click(object sender, RoutedEventArgs e)
        {
            DataStore.Instance.InsertLogByFormS(this.GetType().Name, stDate, stTime, "E");
            this.Close();
        }

        //검색
        private void btnSearch_Click(object sender, RoutedEventArgs e)
        {
            using (Loading ld = new Loading(beSearch))
            {
                ld.ShowDialog();
            }
            
        }

        private void beSearch()
        {
            btnSearch.IsEnabled = false;

            Dispatcher.BeginInvoke(new Action(() =>
            {
                re_Search();
            }), System.Windows.Threading.DispatcherPriority.Background);

            btnSearch.IsEnabled = true;
        }

        #endregion // 주요 버튼 이벤트

        #region Header 부분 - 검색조건

        // 거래처
        private void lblCustom_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            chkCustom.IsChecked = chkCustom.IsChecked == true ? false : true;
        }
        private void chkCustom_Checked(object sender, RoutedEventArgs e)
        {
            chkCustom.IsChecked = true;
            txtCustomSrh.IsEnabled = true;
            btnCustom.IsEnabled = true;
        }
        private void chkCustom_Unchecked(object sender, RoutedEventArgs e)
        {
            chkCustom.IsChecked = false;
            txtCustomSrh.IsEnabled = false;
            btnCustom.IsEnabled = false;
        }

        // 최종고객사
        private void lblInCustom_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            chkInCustom.IsChecked = chkInCustom.IsChecked == true ? false : true;
        }
        private void chkInCustom_Checked(object sender, RoutedEventArgs e)
        {
            chkInCustom.IsChecked = true;
            txtInCustomSrh.IsEnabled = true;
            btnInCustom.IsEnabled = true;
        }
        private void chkInCustom_Unchecked(object sender, RoutedEventArgs e)
        {
            chkInCustom.IsChecked = false;
            txtInCustomSrh.IsEnabled = false;
            btnInCustom.IsEnabled = false;
        }

        // 품번
        private void lblBuyerArticleNo_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            chkBuyerArticleNo.IsChecked = chkBuyerArticleNo.IsChecked == true ? false : true;
        }
        private void chkBuyerArticleNo_Checked(object sender, RoutedEventArgs e)
        {
            chkBuyerArticleNo.IsChecked = true;
            txtBuyerArticleNoSrh.IsEnabled = true;
            btnBuyerArticleNo.IsEnabled = true;
        }
        private void chkBuyerArticleNo_Unchecked(object sender, RoutedEventArgs e)
        {
            chkBuyerArticleNo.IsChecked = false;
            txtBuyerArticleNoSrh.IsEnabled = false;
            btnBuyerArticleNo.IsEnabled = false;
        }

        // 품명
        private void lblArticleSrh_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            chkArticle.IsChecked = chkArticle.IsChecked == true ? false : true;
        }
        private void chkArticleSrh_Checked(object sender, RoutedEventArgs e)
        {
            chkArticle.IsChecked = true;
            txtArticleSrh.IsEnabled = true;
            btnArticle.IsEnabled = true;
        }
        private void chkArticleSrh_Unchecked(object sender, RoutedEventArgs e)
        {
            chkArticle.IsChecked = false;
            txtArticleSrh.IsEnabled = false;
            btnArticle.IsEnabled = false;
        }


        #endregion // Header 부분 - 검색조건

        #region 주요 메서드 모음

        private void re_Search()
        {
            FillGrid();

            if (dgdMain.Items.Count > 0)
                dgdMain.SelectedIndex = rowNum;
            else
                this.DataContext = null;
        }

        #region 조회

        private void FillGrid()
        {
            if (dgdMain.Items.Count > 0)
                dgdMain.Items.Clear();

            try
            {
                DataStore.Instance.InsertLogByForm(this.GetType().Name, "R");

                Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                sqlParameter.Clear();

                //수주일 조건
                int ChkOrderDay = chkOrderDay.IsChecked == true ? 1 : 0;
                string sDate = string.Empty;
                string eDate = string.Empty;
                if(chkOrderDay.IsChecked == true)
                {
                    sDate = dtpSDate.SelectedDate.Value.ToString("yyyyMMdd");
                    eDate = dtpEDate.SelectedDate.Value.ToString("yyyyMMdd");
                }
                else
                {   sDate = "''";  eDate = "''"; }


                //거래처 조건
                int ChkCustom = chkCustom.IsChecked == true ? 1 : 0;
                string Custom = string.Empty;
                if (chkCustom.IsChecked == true)
                     Custom = txtCustomSrh.Text;
                else Custom = "''";


                //최종거래처
                int ChkInCustom = chkInCustom.IsChecked == true ? 1 : 0;
                string InCustom = string.Empty;
                if (chkInCustom.IsChecked == true)
                     InCustom = txtInCustomSrh.Text;
                else InCustom = "''";

                int checkArticleID = chkBuyerArticleNo.IsChecked == true ? 1 : 0;
                string articleID = checkArticleID == 1 ? (txtBuyerArticleNoSrh.Tag == null ? "''" : txtBuyerArticleNoSrh.Tag.ToString()) : "''";

                if (checkArticleID == 0)
                {
                    checkArticleID = chkArticle.IsChecked == true ? 1 : 0;
                    articleID = checkArticleID == 1 ? (txtArticleSrh.Tag == null ? "''" : txtArticleSrh.Tag.ToString()) : "''";
                }

                //string sql = "select  ArticleID	        = fms.ArticleID                                             "
                //           + "      , LocName           = cc.Code_Name                                              "
                //           + "      , Article                                                                       "
                //           + "      , BuyerArticleNo                                                                "
                //           + "      , ArticleGrp                                                                    "
                //           + "      , UnitPrice                                                                     "
                //           + "      , StuffINQty        = sum(StuffINQty)                                           "
                //           + "      , OutQty            = sum(OutQty)                                               "
                //           + "      , Stockqty          = sum(fms.StuffINQty - fms.OutQty)                          "
                //           + "from  [dbo].[fn_mtr_sLotStockLoc] (CONVERT(CHAR(8), DATEADD(dd, 1, GETDATE()), 112)   "
                //           + "                                      , 0                                             "
                //           + "                                      , ''                                            "
                //           + "                                      , " + checkArticleID.ToString() + "             "
                //           + "                                      , '" + articleID + "'                           "
                //           + "                                      , 'Y'                                           "
                //           + "                                      , 'Y'                                           "
                //           + "                                      , 0                                             "
                //           + "                                      , ''                                            "
                //           + "                                      , 1                                             "
                //           + "                                      )   fms                                         "
                //           + "left outer join CM_Code cc        on cc.Code_ID = fms.LocID                           "
                //           + "left outer join mt_Article ma     on ma.ArticleID = fms.ArticleID                     "
                //           + "left outer join mt_ArticleGrp mag	on mag.ArticleGrpID = ma.ArticleGrpID               "
                //           + "where	    cc.Code_GBN = 'LOC'                                                         "
                //           + "      and BigMiSmalGbn = '04'                                                         "
                //           + "      and ma.ArticleGrpID = '05'                                                      "
                //           + "      and (fms.StuffINQty - fms.OutQty) > 0                                           "
                //           + "group by fms.ArticleID, cc.Code_Name, Article, BuyerArticleNo, ArticleGrp, UnitPrice  "             
                //           ;
                //string sql = "select * from dbo.fn_ord_sLotStock(" + checkArticleID.ToString() + ", '" + articleID + "')";

                string sql      = $" select * from [dbo].[fn_ord_sLotStock_For_OutwareReq] (               " +
                                  $"                                                                       "+
                                  $"                                          {ChkOrderDay}      "                +
                                  $"                                         ,{sDate}"                      +
                                  $"                                         ,{eDate}"                      +
                                  $"                                         ,{ChkCustom}"                  +
                                  $"                                         ,{Custom}"                     +
                                  $"                                         ,{ChkInCustom}"                +
                                  $"                                         ,{InCustom}                     )";
                DataSet ds = DataStore.Instance.QueryToDataSet(sql);

                if (ds != null && ds.Tables.Count > 0)
                {
                    int i = 0;

                    DataTable dt = ds.Tables[0];
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
                            var OrderCodeView = new Win_ord_OutwareReqSub_U_View
                            {
                                Num = i,
                                Chk = false,
                                OrderID = dr["OrderID"].ToString(),
                                OrderNo = dr["OrderNo"].ToString(),
                                OrderQty = dr["OrderQty"].ToString(),
                                RestOrderQty = stringFormatN0(dr["RestOrderQty"]),
                                AcptDate = dr["AcptDate"].ToString(),
                                DvlyDate = dr["DvlyDate"].ToString(),
                                CustomName = dr["CustomName"].ToString(),
                                ArticleID = dr["ArticleID"].ToString(),
                                LocName = dr["LocName"].ToString(),
                                Article = dr["Article"].ToString(),
                                BuyerArticleNo = dr["BuyerArticleNo"].ToString(),
                                ArticleGrp = dr["ArticleGrp"].ToString(),
                                //UnitPrice = dr["UnitPrice"].ToString(),
                                //StuffINQty = lib.returnNumString(dr["StuffINQty"].ToString()),
                                OutQty = lib.returnNumString(dr["OutQty"].ToString()),
                                StockQty = lib.returnNumString(dr["StockQty"].ToString()),
                            };

                            dgdMain.Items.Add(OrderCodeView);
                        }
                           
                        tblCount.Text = "▶검색개수 : " + i + "건";
                    }
                }

                #region 사용안함
                /*sqlParameter.Add("ChkDate", chkDate);
                sqlParameter.Add("SDate", startDate != null ? startDate : "");
                sqlParameter.Add("EDate", endDate != null ? endDate : "");

                // 거래처
                sqlParameter.Add("ChkCustom", chkCustomID);
                sqlParameter.Add("CustomID", customID != null ? customID : "" );
                // 최종고객사
                sqlParameter.Add("ChkInCustom", chkInCustomID);
                sqlParameter.Add("InCustomID", inCustomID != null ? inCustomID : "");


                // 품번
                sqlParameter.Add("ChkArticleID", chkBuyerArticleNo.IsChecked == true ? 1 : 0);
                sqlParameter.Add("ArticleID", chkBuyerArticleNo.IsChecked == true ? (txtBuyerArticleNoSrh.Tag == null ? "" : txtBuyerArticleNoSrh.Tag.ToString()) : "");
                // 품명
                sqlParameter.Add("ChkArticle", chkArticle.IsChecked == true ? 1 : 0);
                sqlParameter.Add("Article", chkArticle.IsChecked == true ? (txtArticleSrh.Text == string.Empty ? "" : txtArticleSrh.Text) : "");


                // 관리번호
                sqlParameter.Add("ChkOrderID", 0);
                sqlParameter.Add("OrderID", "");
                // 완료구분
                sqlParameter.Add("ChkCloseClss", 1);
                sqlParameter.Add("CloseClss", "");


                // 품명대분류
                sqlParameter.Add("ChkArticleGbn", 0);
                sqlParameter.Add("ArticleGbn", "");
                // 수주구분
                sqlParameter.Add("ChkOrderFlag", 0);
                sqlParameter.Add("OrderFlag", "");

                DataSet ds = DataStore.Instance.ProcedureToDataSet_LogWrite("xp_ord_sOrder", sqlParameter, true, "R");

                if (ds != null && ds.Tables.Count > 0)
                {
                    int i = 0;

                    DataTable dt = ds.Tables[0];
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
                            var OrderCodeView = new Win_ord_OutwareReqSub_U_View
                            {
                                Num = i,
                                Chk = false,
                                OrderID = dr["OrderID"].ToString(),
                                OrderNo = dr["OrderNo"].ToString(),
                                OrderSeq = dr["OrderSeq"].ToString(),
                                CustomID = dr["CustomID"].ToString(),
                                KCustom = dr["KCustom"].ToString(),
                                InCustomID = dr["InCustomID"].ToString(),
                                KInCustom = dr["KInCustom"].ToString(),
                                ArticleID = dr["ArticleID"].ToString(),
                                Article = dr["Article"].ToString(),
                                BuyerArticleNo = dr["BuyerArticleNo"].ToString(),
                                ArticleGrpID = dr["ArticleGrpID"].ToString(),
                                ArticleGrp = dr["ArticleGrp"].ToString(),
                                UnitPrice = stringFormatN0(dr["UnitPrice"]),
                                OrderQty = dr["OrderQty"].ToString(),
                                AcptDate = DatePickerFormat(dr["AcptDate"].ToString()),
                                DvlyDate = DatePickerFormat(dr["DvlyDate"].ToString()),
                            };

                            dgdMain.Items.Add(OrderCodeView);
                        }
                    }

                    tblCount.Text = "▶검색개수 : " + i + "건";
                }*/
                #endregion

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

        #endregion

        #region 전체 선택 체크박스 이벤트

        // 전체 선택 체크박스 체크 이벤트
        private void AllCheck_Checked(object sender, RoutedEventArgs e)
        {
            for (int i = 0; i < dgdMain.Items.Count; i++)
            {
                var order = dgdMain.Items[i] as Win_ord_OutwareReqSub_U_View;
                order.Chk = true;
            }
        }

        // 전체 선택 체크박스 언체크 이벤트
        private void AllCheck_Unchecked(object sender, RoutedEventArgs e)
        {
            if (dgdMain.Visibility == Visibility.Visible)
            {
                for (int i = 0; i < dgdMain.Items.Count; i++)
                {
                    var order = dgdMain.Items[i] as Win_ord_OutwareReqSub_U_View;
                    order.Chk = false;
                }
            }
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
                    result = str.Substring(0, 4) + "-" + str.Substring(4, 2) + "-" + str.Substring(6, 2);
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
            if (e.ClickCount == 2)
            {
                dounleClick = true;
                btnConfirm_Click(null, null);
            }
        }

        private void chkReq_Click(object sender, RoutedEventArgs e)
        {
            CheckBox chkSender = sender as CheckBox;
            var order = chkSender.DataContext as Win_ord_OutwareReqSub_U_View;

            if (order != null)
            {
                if (chkSender.IsChecked == true)
                    order.Chk = true;
                else
                    order.Chk = false;
            }
        }

        // 거래처 검색
        private void TxtCustomSrh_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
                MainWindow.pf.ReturnCode(txtCustomSrh, 76, txtCustomSrh.Text);
        }

        private void btnCustom_Click(object sender, RoutedEventArgs e)
        {
            MainWindow.pf.ReturnCode(txtCustomSrh, 76, txtCustomSrh.Text);
        }

        // 최종고객사 검색
        private void TxtInCustomSrh_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
                MainWindow.pf.ReturnCode(txtInCustomSrh, 76, txtInCustomSrh.Text);
        }

        private void btnInCustom_Click(object sender, RoutedEventArgs e)
        {
            MainWindow.pf.ReturnCode(txtInCustomSrh, 76, txtInCustomSrh.Text);
        }

        // 품번 검색
        private void TxtBuyerArticleNoSrh_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
                MainWindow.pf.ReturnCode(txtBuyerArticleNoSrh, 76, txtBuyerArticleNoSrh.Text);
        }

        private void btnBuyerArticleNo_Click(object sender, RoutedEventArgs e)
        {
            MainWindow.pf.ReturnCode(txtBuyerArticleNoSrh, 76, txtBuyerArticleNoSrh.Text);
        }

        // 품명 검색
        private void TxtArticleSrh_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
                MainWindow.pf.ReturnCode(txtArticleSrh, 77, txtArticleSrh.Text);
        }

        private void btnArticleSrh_Click(object sender, RoutedEventArgs e)
        {
            MainWindow.pf.ReturnCode(txtArticleSrh, 77, txtArticleSrh.Text);
        }

        public List<Win_ord_OutwareReqSub_U_View> GetList() { return listReq; }

        private void chkOrderDay_Checked(object sender, RoutedEventArgs e)
        {
            if (dtpSDate != null && dtpEDate != null)
            {
                dtpSDate.IsEnabled = true;
                dtpEDate.IsEnabled = true;
            }
        }

        private void chkOrderDay_Unchecked(object sender, RoutedEventArgs e)
        {
            dtpSDate.IsEnabled = false;
            dtpEDate.IsEnabled = false;
        }

        private void lblOrderDay_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            search_CheckBox_Control(chkOrderDay);
        }

        private void search_CheckBox_Control(CheckBox checkBox)
        {
            checkBox.IsChecked = checkBox.IsChecked == true ? false : true;
        }
    }
}
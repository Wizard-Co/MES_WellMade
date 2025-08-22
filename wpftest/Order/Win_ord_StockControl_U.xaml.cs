using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Data;
using System.Threading;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using WizMes_WellMade.PopUP;
using WPF.MDI;

namespace WizMes_WellMade
{
    /// <summary>
    /// Win_ord_StockControl_U.xaml에 대한 상호 작용 논리
    /// </summary>
    public partial class Win_ord_StockControl_U : UserControl
    {
        Lib lib = new Lib();
        public Win_ord_StockControl_U()
        {
            InitializeComponent();
        }
        PlusFinder pf = new PlusFinder();

        // 추가저장인지 / 수정저장인지 구별하는 용도입니다.
        string ButtonTag = string.Empty;
        // 추가/수정 시 저장취소와 연계되어 value를 자동으로 찾아 select해 줄 value list 값입니다.
        List<string> lstCompareValue = new List<string>();

        // 로드 시.
        private void Win_ord_StockControl_U_Loaded(object sender, RoutedEventArgs e)
        {
            First_Step();
        }


        #region 첫 스텝 / 조회용 체크박스이벤트 / 날짜버튼

        //첫 단계.
        private void First_Step()
        {
            chkAdjustDate.IsChecked = true;
            dtpFromDate.Text = DateTime.Now.ToString("yyyy-MM-dd");
            dtpToDate.Text = DateTime.Now.ToString("yyyy-MM-dd");

            txtArticle.IsEnabled = false;
            btnArticle.IsEnabled = false;

            dtpAdjustDate.Text = DateTime.Now.ToString("yyyy-MM-dd");
            bdrAdjust.IsEnabled = false;

            btnSave.Visibility = Visibility.Hidden;
            btnCancel.Visibility = Visibility.Hidden;
            EventLabel.Visibility = Visibility.Hidden;

        }

        //전일
        private void btnYesterDay_Click(object sender, RoutedEventArgs e)
        {
            DateTime[] SearchDate = lib.BringLastDayDateTimeContinue(dtpToDate.SelectedDate.Value);

            dtpFromDate.SelectedDate = SearchDate[0];
            dtpToDate.SelectedDate = SearchDate[1];
        }

        //금일
        private void btnToday_Click(object sender, RoutedEventArgs e)
        {
            dtpFromDate.SelectedDate = DateTime.Today;
            dtpToDate.SelectedDate = DateTime.Today;
        }

        // 전월 버튼 클릭 이벤트
        private void btnLastMonth_Click(object sender, RoutedEventArgs e)
        {
            DateTime[] SearchDate = lib.BringLastMonthContinue(dtpFromDate.SelectedDate.Value);

            dtpFromDate.SelectedDate = SearchDate[0];
            dtpToDate.SelectedDate = SearchDate[1];
        }

        // 금월 버튼 클릭 이벤트
        private void btnThisMonth_Click(object sender, RoutedEventArgs e)
        {
            dtpFromDate.SelectedDate = lib.BringThisMonthDatetimeList()[0];
            dtpToDate.SelectedDate = lib.BringThisMonthDatetimeList()[1];
        }

        // 조정일자 클릭
        private void chkAdjustDate_Click(object sender, RoutedEventArgs e)
        {
            if (chkAdjustDate.IsChecked == true)
            {
                dtpFromDate.IsEnabled = true;
                dtpToDate.IsEnabled = true;
            }
            else
            {
                dtpFromDate.IsEnabled = false;
                dtpToDate.IsEnabled = false;
            }
        }
        //조정일자 클릭.
        private void chkAdjustDate_Click(object sender, MouseButtonEventArgs e)
        {
            if (chkAdjustDate.IsChecked == true)
            {
                chkAdjustDate.IsChecked = false;
                dtpFromDate.IsEnabled = false;
                dtpToDate.IsEnabled = false;
            }
            else
            {
                chkAdjustDate.IsChecked = true;
                dtpFromDate.IsEnabled = true;
                dtpToDate.IsEnabled = true;
            }
        }

        // 품명 클릭
        private void chkArticle_Click(object sender, RoutedEventArgs e)
        {
            if (chkArticle.IsChecked == true)
            {
                txtArticle.IsEnabled = true;
                txtArticle.Focus();
                btnArticle.IsEnabled = true;
            }
            else
            {
                txtArticle.IsEnabled = false;
                btnArticle.IsEnabled = false;
            }
        }
        // 품명 클릭.
        private void chkArticle_Click(object sender, MouseButtonEventArgs e)
        {
            if (chkArticle.IsChecked == true)
            {
                chkArticle.IsChecked = false;
                txtArticle.IsEnabled = false;
                btnArticle.IsEnabled = false;
            }
            else
            {
                chkArticle.IsChecked = true;
                txtArticle.IsEnabled = true;
                btnArticle.IsEnabled = true;
                txtArticle.Focus();
            }
        }

        #endregion


        #region 공통버튼 이벤트
        //공통 사용가능
        private void PublicEnableTrue()
        {
            btnSave.Visibility = Visibility.Hidden;
            btnCancel.Visibility = Visibility.Hidden;
            EventLabel.Visibility = Visibility.Hidden;

            btnSearch.IsEnabled = true;
            btnAdd.IsEnabled = true;
            btnEdit.IsEnabled = true;
            btnDelete.IsEnabled = true;
            btnExcel.IsEnabled = true;

            bdrAdjust.IsEnabled = false;

            //dgdAdjust.IsEnabled = true; //메인그리드 사용가능.
            dgdAdjust.IsHitTestVisible = true;
            dgdAdjust_sub.SelectionUnit = DataGridSelectionUnit.FullRow;

        }



        // 공통 버튼이벤트.
        private void PublicEnableFalse()
        {
            btnSave.Visibility = Visibility.Visible;
            btnCancel.Visibility = Visibility.Visible;
            EventLabel.Visibility = Visibility.Visible;

            btnSearch.IsEnabled = false;
            btnAdd.IsEnabled = false;
            btnEdit.IsEnabled = false;
            btnDelete.IsEnabled = false;
            btnExcel.IsEnabled = false;

            bdrAdjust.IsEnabled = true;

            //조정번호는 건들지 마.
            txtAdjustNumber.IsReadOnly = true;
            //dgdAdjust.IsEnabled = false;    // 메인그리드 못건드리게.
            dgdAdjust.IsHitTestVisible = false;
            dgdAdjust_sub.SelectionUnit = DataGridSelectionUnit.Cell;
        }
        #endregion


        #region 플러스파인더

        // 플러스파인더_품명클릭(품번)수정 2020.03.18, 장가빈
        private void btnArticle_Click(object sender, RoutedEventArgs e)
        {
            //pf.ReturnCode(txtArticle, 81, txtArticle.Text);
            pf.ReturnCode(txtArticle, 84, txtArticle.Text);
        }

        //키다운_품명클릭(품번)수정 2020.03.18, 장가빈
        private void TxtArticle_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                //pf.ReturnCode(txtArticle, 81, txtArticle.Text);
                pf.ReturnCode(txtArticle, 84, txtArticle.Text);
            }
        }

        private void btnWorker_Click(object sender, RoutedEventArgs e)
        {
            pf.ReturnCode(txtWorker, 2, "");
            Lib.Instance.SendK(Key.Tab, this);
        }
        #endregion


        #region 조회 / 조회용 프로시저
        // 조회.
        private void btnSearch_Click(object sender, RoutedEventArgs e)
        {
            //검색버튼 비활성화
            btnSearch.IsEnabled = false;

            Dispatcher.BeginInvoke(new Action(() =>

            {
                Thread.Sleep(2000);

                //로직
                FillGrid();

            }), System.Windows.Threading.DispatcherPriority.Background);



            Dispatcher.BeginInvoke(new Action(() =>

            {
                btnSearch.IsEnabled = true;

            }), System.Windows.Threading.DispatcherPriority.Background);

        }
        private void FillGrid()
        {

            int ChkDate = 0;
            if (chkAdjustDate.IsChecked == true) { ChkDate = 1; }
            string FromDate = dtpFromDate.ToString().Substring(0, 10).Replace("-", "");
            string ToDate = dtpToDate.ToString().Substring(0, 10).Replace("-", "");         //조정일자

            int ChkArticle = 0;
            if (chkArticle.IsChecked == true) { ChkArticle = 1; }
            else { txtArticle.Tag = ""; }

            if (dgdAdjust.Items.Count > 0)
            {
                dgdAdjust.Items.Clear();
            }

            if (dgdAdjust_sub.Items.Count > 0)
            {
                dgdAdjust_sub.Items.Clear();
            }


            try
            {
                Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                sqlParameter.Add("ChkDate", ChkDate);
                sqlParameter.Add("FromDate", FromDate);
                sqlParameter.Add("ToDate", ToDate);
                sqlParameter.Add("ChkArticle", 0); // ChkArticle);
                sqlParameter.Add("ArticleID", ""); //txtArticle.Tag.ToString());
                sqlParameter.Add("BuyerArticleNo", chkArticle.IsChecked == true ? txtArticle.Text : "");

                DataSet ds = DataStore.Instance.ProcedureToDataSet("xp_sbStock_sStockControl", sqlParameter, false);

                if (ds != null && ds.Tables.Count > 0)
                {
                    DataTable dt = null;
                    dt = ds.Tables[0];

                    if (dt.Rows.Count == 0)
                    {
                        dgdAdjust.Items.Clear(); //1개 밖에 없는 데이터까지 삭제한 경우 계속 남아있어서 클리어함
                        MessageBox.Show("조회결과가 없습니다.");
                        return;
                    }
                    else
                    {
                        dgdAdjust.Items.Clear();
                        DataRowCollection drc = dt.Rows;
                        int i = 1;
                        foreach (DataRow item in drc)
                        {
                            var Win_ord_StockControl_U_Insert = new Win_ord_StockControl_U_View()
                            {
                                ControlID = item["ControlID"].ToString(),
                                ControlDate = item["ControlDate"].ToString(),
                                Full_ControlDate = DateTime.ParseExact(item["ControlDate"].ToString(), "yyyyMMdd", null).ToString("yyyy-MM-dd"),
                                Name = item["Name"].ToString(),
                                PersonID = item["PersonID"].ToString(),
                                Comments = item["Comments"].ToString(),
                                ArticleID = item["ArticleID"].ToString(),
                                Article = item["Article"].ToString()
                            };
                            dgdAdjust.Items.Add(Win_ord_StockControl_U_Insert);
                            //자동 셀 선택되게
                            dgdAdjust.SelectedIndex = 0;
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

        #endregion


        #region 메인그리드 연동 - 서브그리드 showData
        //조정그리드 셀 항목 클릭 시.
        private void dgdAdjust_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            this.DataContext = dgdAdjust.SelectedItem as Win_ord_StockControl_U_View;
            var ViewReceiver = dgdAdjust.SelectedItem as Win_ord_StockControl_U_View;


            if (ViewReceiver != null)
            {
                dtpAdjustDate.Text = ViewReceiver.Full_ControlDate;  // 날짜.

                if (ViewReceiver.ControlID != "")   // 조정번호에 빈값이 들어간게 아니라면,
                {

                    txtWorker.Tag = ViewReceiver.PersonID;

                    Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                    sqlParameter.Add("ControlID", ViewReceiver.ControlID);

                    DataSet ds = DataStore.Instance.ProcedureToDataSet("xp_sbStock_sStockControlSub", sqlParameter, false);

                    if (ds != null && ds.Tables.Count > 0)
                    {
                        DataTable dt = null;
                        dt = ds.Tables[0];

                        if (dt.Rows.Count == 0)
                        {
                            return;
                        }
                        else
                        {
                            dgdAdjust_sub.Items.Clear();
                            DataRowCollection drc = dt.Rows;
                            int i = 1;
                            foreach (DataRow item in drc)
                            {
                                var Win_ord_StockControl_U_Insert = new Win_ord_StockControl_U_View()
                                {
                                    ControlID_s = item["ControlID"].ToString(),
                                    ControlSeq_s = item["ControlSeq"].ToString(),
                                    ArticleID_s = item["ArticleID"].ToString(),
                                    ControlQty_s = stringFormatN0(item["ControlQty"]),
                                    UnitClss_s = item["UnitClss"].ToString(),
                                    UnitClss_string_s = item["Code_Name"].ToString(),
                                    Comments_s = item["Comments"].ToString(),
                                    BuyerArticleNo = item["BuyerArticleNo"].ToString(),
                                    Article_s = item["Article"].ToString(),
                                    StuffINID_s = item["StuffINID"].ToString()
                                };
                                dgdAdjust_sub.Items.Add(Win_ord_StockControl_U_Insert);
                            }
                        }
                    }
                }
            }
        }

        #endregion


        #region (추가 / 수정 / 삭제 / 저장 / 취소 ) 버튼 이벤트 모음

        // 추가버튼 클릭 시.
        private void btnAdd_Click(object sender, RoutedEventArgs e)
        {
            //1. 보더 데이터 클리어.
            bdrAdjustDataClear();
            //2. 서브 그리드 클리어.
            dgdAdjust_sub.Items.Clear();
            //3. 공통 버튼이벤트
            PublicEnableFalse();

            EventLabel.Content = "자료입력(추가)중..";
            ButtonTag = ((Button)sender).Tag.ToString();
            dtpAdjustDate.Focus();
        }

        // 수정 버튼 클릭 시.
        private void btnEdit_Click(object sender, RoutedEventArgs e)
        {
            // 1. 수정할 자격은 있는거야? 조회? 데이터 선택??
            if (dgdAdjust.Items.Count < 1)
            {
                MessageBox.Show("먼저 검색해 주세요.");
                return;
            }
            var OBJ = dgdAdjust.SelectedItem as Win_ord_StockControl_U_View;
            if (OBJ == null)
            {
                MessageBox.Show("수정할 항목이 정확히 선택되지 않았습니다.");
                return;
            }

            // 공통 버튼이벤트
            PublicEnableFalse();

            EventLabel.Content = "자료입력(수정)중..";
            ButtonTag = ((Button)sender).Tag.ToString();
        }

        // 삭제버튼 클릭 시.
        private void btnDelete_Click(object sender, RoutedEventArgs e)
        {
            // 1. 삭제할 자격은 있는거야? 조회? 데이터 선택??
            if (dgdAdjust.Items.Count < 1)
            {
                MessageBox.Show("먼저 검색해 주세요.");
                return;
            }
            var OBJ = dgdAdjust.SelectedItem as Win_ord_StockControl_U_View;
            if (OBJ == null)
            {
                MessageBox.Show("삭제할 항목이 정확히 선택되지 않았습니다.");
                return;
            }

            MessageBoxResult msgresult = MessageBox.Show("선택하신 항목을 삭제하시겠습니까?", "삭제 확인", MessageBoxButton.YesNo);
            if (msgresult == MessageBoxResult.Yes)
            {
                // 2.  삭제용
                DeleteData();
                dgdAdjust.Refresh();

                // 3. 화면정리.
                btnCancel_Click(null, null);
            }

        }

        //저장 버튼 클릭.
        private void btnSave_Click(object sender, RoutedEventArgs e)
        {
            // 1. 데이터 기입체크.(항목을 제대로 모두 똑바로 넣고 저장버튼을 누르는 거야??) 
            if (BDRAdjustDataCheck() == false) { return; }
            // 2. 저장.
            SaveData(ButtonTag);

            btnCancel_Click(null, null);        //화면 정리         
        }


        //취소 버튼 클릭
        private void btnCancel_Click(object sender, RoutedEventArgs e)
        {
            //공통 버튼이벤트
            PublicEnableTrue();
            //보더 데이터 클리어.
            bdrAdjustDataClear();
            //서브 그리드 클리어.
            dgdAdjust_sub.Items.Clear();
            FillGrid();                         //재 조회.

            // 추가/수정 이후 저장.취소시 Target 자동 세팅.
            if (ButtonTag != string.Empty)
            {
                int ReturnCount = Lib.Instance.reTrunIndex(dgdAdjust, lstCompareValue);
                if (dgdAdjust.Items.Count > 0)
                {
                    object item = dgdAdjust.Items[ReturnCount];
                    dgdAdjust.SelectedItem = item;
                    dgdAdjust.ScrollIntoView(item);
                    DataGridRow row = dgdAdjust.ItemContainerGenerator.ContainerFromIndex(ReturnCount) as DataGridRow;
                    if (row != null)
                    {
                        row.MoveFocus(new TraversalRequest(FocusNavigationDirection.Next));
                    }
                }
            }
            ButtonTag = string.Empty;
            this.Focusable = true;
            this.Focus();
        }

        #endregion


        #region 실제 프로시저 모음 CRUD
        // 저장.
        private void SaveData(string TagNUM)
        {

            try
            {
                List<Procedure> Prolist = new List<Procedure>();
                List<Dictionary<string, object>> ListParameter = new List<Dictionary<string, object>>();
                if (TagNUM == "1")      // 신규추가입니다.
                {
                    // 신규추가 저장 insert.

                    //1. 메인정보 저장하기
                    string ControlDate = dtpAdjustDate.Text.Substring(0, 10).Replace("-", "");
                    string PersonID = txtWorker.Tag.ToString();
                    string Comments = txtReason.Text;

                    // 추가저장시 > findTargetValue 설정.
                    lstCompareValue.Add(ControlDate);
                    lstCompareValue.Add(PersonID);
                    lstCompareValue.Add(Comments);


                    Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                    sqlParameter.Clear();
                    sqlParameter.Add("ControlID", "");
                    sqlParameter.Add("ControlDate", ControlDate);
                    sqlParameter.Add("PersonID", PersonID);
                    sqlParameter.Add("Comments", Comments);
                    sqlParameter.Add("CreateUserID", MainWindow.CurrentUser);

                    Procedure pro1 = new Procedure();
                    pro1.Name = "xp_sbStock_iStockControl";
                    pro1.OutputUseYN = "Y";
                    pro1.OutputName = "ControlID";
                    pro1.OutputLength = "12";

                    Prolist.Add(pro1);
                    ListParameter.Add(sqlParameter);


                    //2. 서브정보 저장하기
                    int upgradePoint = dgdAdjust_sub.Items.Count;
                    for (int i = 0; i < upgradePoint; i++)
                    {
                        DataGridRow dgr = Lib.Instance.GetRow(i, dgdAdjust_sub);
                        var ViewReceiver = dgr.Item as Win_ord_StockControl_U_View;

                        string ArticleID = ViewReceiver.ArticleID_s.ToString();
                        string ControlQty = ViewReceiver.ControlQty_s.ToString();
                        string UnitClss = ViewReceiver.UnitClss_s.ToString();
                        string Comments_s = ViewReceiver.Comments_s.ToString();
                        string StuffINID = ViewReceiver.StuffINID_s.ToString();

                        Dictionary<string, object> sqlParameter2 = new Dictionary<string, object>();
                        sqlParameter2.Clear();
                        sqlParameter2.Add("ControlID", "");
                        sqlParameter2.Add("ControlSeq", i + 1);
                        sqlParameter2.Add("ArticleID", ArticleID);
                        sqlParameter2.Add("ControlQty", ControlQty);
                        sqlParameter2.Add("UnitClss", UnitClss);
                        sqlParameter2.Add("StuffINID", StuffINID);
                        sqlParameter2.Add("Comments", Comments_s);
                        sqlParameter2.Add("CreateUserID", MainWindow.CurrentUser);

                        Procedure pro2 = new Procedure();
                        pro2.Name = "xp_sbStock_iStockControlSub";
                        pro2.OutputUseYN = "N";
                        pro2.OutputName = "ControlID";
                        pro2.OutputLength = "12";

                        Prolist.Add(pro2);
                        ListParameter.Add(sqlParameter2);
                    }

                    string[] Confirm = new string[2];
                    Confirm = DataStore.Instance.ExecuteAllProcedureOutputNew(Prolist, ListParameter);
                    if (Confirm[0] == "success")
                    {
                        //return true;
                    }
                    else
                    {
                        MessageBox.Show("[저장실패]\r\n" + Confirm[1].ToString());
                        //return false;
                    }
                }

                else if (TagNUM == "2")         // 수정 저장입니다.
                {
                    // 수정 저장 update.

                    string ControlID = txtAdjustNumber.Text.ToString();
                    string ControlDate = dtpAdjustDate.Text.Substring(0, 10).Replace("-", "");
                    string PersonID = txtWorker.Tag.ToString();
                    string Comments = txtReason.Text;

                    // 수정저장시 > findTargetValue 설정.
                    lstCompareValue.Add(ControlID);
                    lstCompareValue.Add(ControlDate);
                    lstCompareValue.Add(PersonID);
                    lstCompareValue.Add(Comments);

                    Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                    sqlParameter.Clear();
                    sqlParameter.Add("ControlID", ControlID);
                    sqlParameter.Add("ControlDate", ControlDate);
                    sqlParameter.Add("PersonID", PersonID);
                    sqlParameter.Add("Comments", Comments);
                    sqlParameter.Add("CreateUserID", MainWindow.CurrentUser);

                    Procedure pro1 = new Procedure();
                    pro1.Name = "xp_sbStock_uStockControl";
                    pro1.OutputUseYN = "N";
                    pro1.OutputName = "ControlID";
                    pro1.OutputLength = "12";

                    Prolist.Add(pro1);
                    ListParameter.Add(sqlParameter);


                    //2. 서브정보 저장하기
                    int upgradePoint = dgdAdjust_sub.Items.Count;
                    for (int i = 0; i < upgradePoint; i++)
                    {
                        DataGridRow dgr = Lib.Instance.GetRow(i, dgdAdjust_sub);
                        var ViewReceiver = dgr.Item as Win_ord_StockControl_U_View;

                        string ArticleID = ViewReceiver.ArticleID_s.ToString();
                        string ControlQty = ViewReceiver.ControlQty_s.Replace(",", "").ToString();
                        string UnitClss = ViewReceiver.UnitClss_s.ToString();
                        string Comments_s = ViewReceiver.Comments_s.ToString();
                        string StuffINID = ViewReceiver.StuffINID_s.ToString();

                        Dictionary<string, object> sqlParameter2 = new Dictionary<string, object>();
                        sqlParameter2.Clear();
                        sqlParameter2.Add("ControlID", ControlID);
                        sqlParameter2.Add("ControlSeq", i + 1);
                        sqlParameter2.Add("ArticleID", ArticleID);
                        sqlParameter2.Add("ControlQty", ControlQty);
                        sqlParameter2.Add("UnitClss", UnitClss);
                        sqlParameter2.Add("StuffINID", StuffINID);
                        sqlParameter2.Add("Comments", Comments_s);
                        sqlParameter2.Add("CreateUserID", MainWindow.CurrentUser);

                        Procedure pro2 = new Procedure();
                        pro2.Name = "xp_sbStock_iStockControlSub";
                        pro2.OutputUseYN = "N";
                        pro2.OutputName = "ControlID";
                        pro2.OutputLength = "12";

                        Prolist.Add(pro2);
                        ListParameter.Add(sqlParameter2);
                    }

                    string[] Confirm = new string[2];
                    Confirm = DataStore.Instance.ExecuteAllProcedureOutputNew(Prolist, ListParameter);
                    if (Confirm[0] == "success")
                    {
                        //return true;
                    }
                    else
                    {
                        MessageBox.Show("[저장실패]\r\n" + Confirm[1].ToString());
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

        // 실제 삭제 프로시저.
        private void DeleteData()
        {
            try
            {
                string ControlID = txtAdjustNumber.Text.ToString();

                Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                sqlParameter.Clear();
                sqlParameter.Add("ControlID", ControlID);

                string[] result = DataStore.Instance.ExecuteProcedure("xp_sbStock_dStockControl", sqlParameter, false);
                if (!result[0].Equals("success"))
                {
                    MessageBox.Show("이상발생, 관리자에게 문의하세요.");
                    return;
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

        #endregion


        #region Adjust 블럭 체크데이터 / 데이터 클리어

        // 조정블럭 저장전, 데이터 체크!.
        private bool BDRAdjustDataCheck()
        {
            //1. 작업자.
            if ((txtWorker.Tag == null))  // || (lib.IsNullOrWhiteSpace(txtWorker.Tag.ToString()) == true))
            {
                MessageBox.Show("작업자를 선택해야 합니다.");
                return false;
            }
            //2. 사유 (특히 강조)
            if (Lib.Instance.IsNullOrWhiteSpace(txtReason.Text) == true)
            {
                MessageBox.Show("사유는 반드시 기입해야 합니다.");
                return false;
            }

            int upgradePoint = dgdAdjust_sub.Items.Count;
            if (upgradePoint == 0)
            {
                MessageBox.Show("최소 하나 이상의 데이터를 추가해야 합니다.");
                return false;
            }
            for (int i = 0; i < upgradePoint; i++)
            {
                DataGridRow dgr = Lib.Instance.GetRow(i, dgdAdjust_sub);
                var ViewReceiver = dgr.Item as Win_ord_StockControl_U_View;

                DataGridCell cell2 = Lib.Instance.GetCell(i, 3, dgdAdjust_sub);
                TextBox tb2 = Lib.Instance.GetVisualChild<TextBox>(cell2);
                DataGridCell cell4 = Lib.Instance.GetCell(i, 5, dgdAdjust_sub);
                TextBox tb4 = Lib.Instance.GetVisualChild<TextBox>(cell4);

                ViewReceiver.ControlQty_s = tb2.Text;
                ViewReceiver.Comments_s = tb4.Text;


                // 서브그리드 품목
                if (Lib.Instance.IsNullOrWhiteSpace(ViewReceiver.ArticleID_s.ToString()) == true)
                {
                    MessageBox.Show((i + 1) + "번째 품목은 반드시 기입해야 합니다.");
                    return false;
                }
                // 서브그리드 수량
                if (Lib.Instance.IsIntOrAnother(ViewReceiver.ControlQty_s.Replace(",", "").ToString()) == false)
                {
                    MessageBox.Show((i + 1) + "번째 수량은 반드시 숫자로 기입해야 합니다.");
                    return false;
                }
                // 서브그리드 단위
                if (Lib.Instance.IsNullOrWhiteSpace(ViewReceiver.UnitClss_string_s.ToString()) == true)
                {
                    MessageBox.Show((i + 1) + "번째 단위는 반드시 기입해야 합니다.");
                    return false;
                }
            }
            return true;
        }


        // 조정 블럭 데이터 클리어.
        private void bdrAdjustDataClear()
        {
            txtAdjustNumber.Text = string.Empty;
            dtpAdjustDate.Text = DateTime.Now.ToString("yyyy-MM-dd");
            txtWorker.Text = string.Empty;
            txtWorker.Tag = null;
            txtReason.Text = string.Empty;
        }

        #endregion


        #region 보더블럭 서브그리드 내부 추가버튼 / 삭제버튼 기능구현

        // 보더 블럭 내부의 추가버튼 클릭 시.
        private void btnPlus_Click(object sender, RoutedEventArgs e)
        {
            //0. 마우스를 쓰고있을 당신을 위해, + 혹시모르니까. 추가 누를때마다, sub 그리드 view upload.
            int upgradePoint = dgdAdjust_sub.Items.Count;
            for (int i = 0; i < upgradePoint; i++)
            {
                DataGridRow dgr = Lib.Instance.GetRow(i, dgdAdjust_sub);
                var ViewReceiver = dgr.Item as Win_ord_StockControl_U_View;
                if (ViewReceiver != null)
                {
                    DataGridCell cell1 = Lib.Instance.GetCell(i, 3, dgdAdjust_sub);
                    TextBox tb1 = Lib.Instance.GetVisualChild<TextBox>(cell1);
                    DataGridCell cell2 = Lib.Instance.GetCell(i, 5, dgdAdjust_sub);
                    TextBox tb2 = Lib.Instance.GetVisualChild<TextBox>(cell2);

                    ViewReceiver.ControlQty_s = tb1.Text;
                    ViewReceiver.Comments_s = tb2.Text;
                }
            }
            dgdAdjust_sub.Items.Refresh();

            //1.  추가되어 지는 새 항목을 넣는 작업.
            var Win_ord_StockControl_U_Insert = new Win_ord_StockControl_U_View()
            {
                ControlID_s = string.Empty,
                ControlSeq_s = string.Empty,
                ControlQty_s = string.Empty,
                UnitClss_string_s = string.Empty,
                UnitClss_s = string.Empty,
                Comments_s = string.Empty,
                ArticleID_s = string.Empty,
                Article_s = string.Empty,
                StuffINID_s = string.Empty
            };
            dgdAdjust_sub.Items.Add(Win_ord_StockControl_U_Insert);


            if (dgdAdjust_sub.Items.Count == 1)
            {
                dgdAdjust_sub.Focus();
                dgdAdjust_sub.CurrentCell = new DataGridCellInfo(dgdAdjust_sub.Items[0], dgdAdjust_sub.Columns[0]);
            }
        }

        // 보더블록 내부의 삭제버튼 클릭 시.
        private void btnMinus_Click(object sender, RoutedEventArgs e)
        {
            var ViewReceiver = dgdAdjust_sub.CurrentCell.Item as Win_ord_StockControl_U_View;  //선택 줄.
            if (ViewReceiver == null)   // 선택한 줄이 없다면,
            {
                int upgradePoint = dgdAdjust_sub.Items.Count;
                DataGridRow dgr = Lib.Instance.GetRow(upgradePoint - 1, dgdAdjust_sub);
                ViewReceiver = dgr.Item as Win_ord_StockControl_U_View;
            }

            dgdAdjust_sub.Items.Remove(ViewReceiver);
        }

        #endregion


        #region 서브그리드 더블클릭 _플러스 파인더 연동하기

        // 서브그리드 항목 중 (품명 , 단위) 항목을 마우스 더블클릭. _ 플러스파인더 활용
        private void dgdAdjust_sub_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            // 지금이 추가 혹은 수정단계여야 한다 그지?  즉 이벤트라벨이 지금 보이고.
            if (EventLabel.Visibility == Visibility.Visible)
            {
                var ViewReceiver = dgdAdjust_sub.CurrentCell.Item as Win_ord_StockControl_U_View;
                if (ViewReceiver != null)
                {
                    List<TextBox> list = new List<TextBox>();
                    Lib.Instance.FindChildGroup<TextBox>(dgdAdjust_sub, "txtdgdArticle_s", ref list);

                    int target = dgdAdjust_sub.Items.IndexOf(dgdAdjust_sub.CurrentCell.Item);

                    TextBox txtdgdArticle_s = list[target];
                    txtdgdArticle_s.IsReadOnly = false;
                    Dispatcher.BeginInvoke((ThreadStart)delegate
                    {
                        txtdgdArticle_s.Focus();
                    });
                }
            }
        }

        // ArticleID로 UnitClss 가져오기.
        private string[] GetUnitClssData(string param)
        {
            string[] UnitClssData = new string[3];
            Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
            sqlParameter.Add("ArticleID", param);

            DataSet ds = DataStore.Instance.ProcedureToDataSet("xp_Order_sArticleData", sqlParameter, false);

            if (ds != null && ds.Tables.Count > 0)
            {
                DataTable dt = null;
                dt = ds.Tables[0];

                if (dt.Rows.Count == 0)
                {
                }
                else
                {
                    UnitClssData[0] = dt.Rows[0]["Code_Name"].ToString();
                    UnitClssData[1] = dt.Rows[0]["UnitClss"].ToString();
                    UnitClssData[2] = dt.Rows[0]["Article"].ToString();
                }
            }
            return UnitClssData;
        }


        #endregion


        #region 서브그리드 타이핑 항목 클릭 이벤트

        // 서브그리드 항목 중 수량쪽을 더블클릭 시.
        private void ContentRender(object sender, MouseButtonEventArgs e)
        {
            // 지금이 추가 혹은 수정단계여야 한다 그지?  즉 이벤트라벨이 지금 보이고.
            if (EventLabel.Visibility == Visibility.Visible)
            {
                var ViewReceiver = dgdAdjust_sub.CurrentCell.Item as Win_ord_StockControl_U_View;
                if (ViewReceiver != null)
                {
                    List<TextBox> list = new List<TextBox>();
                    Lib.Instance.FindChildGroup<TextBox>(dgdAdjust_sub, "txtdgdsub_Controlqty", ref list);

                    int target = dgdAdjust_sub.Items.IndexOf(dgdAdjust_sub.CurrentCell.Item);

                    TextBox txtdgdsub_Controlqty = list[target];
                    txtdgdsub_Controlqty.IsReadOnly = false;
                    Dispatcher.BeginInvoke((ThreadStart)delegate
                    {
                        txtdgdsub_Controlqty.Focus();
                    });
                }
            }
        }
        // 서브그리드 항목 중 비고쪽을 더블클릭 시.
        private void ContentRender1(object sender, MouseButtonEventArgs e)
        {
            // 지금이 추가 혹은 수정단계여야 한다 그지?  즉 이벤트라벨이 지금 보이고.
            if (EventLabel.Visibility == Visibility.Visible)
            {
                var ViewReceiver = dgdAdjust_sub.CurrentCell.Item as Win_ord_StockControl_U_View;
                if (ViewReceiver != null)
                {
                    List<TextBox> list = new List<TextBox>();
                    Lib.Instance.FindChildGroup<TextBox>(dgdAdjust_sub, "txtdgdsub_Comments", ref list);

                    int target = dgdAdjust_sub.Items.IndexOf(dgdAdjust_sub.CurrentCell.Item);

                    TextBox txtdgdsub_Comments = list[target];
                    txtdgdsub_Comments.IsReadOnly = false;
                    Dispatcher.BeginInvoke((ThreadStart)delegate
                    {
                        txtdgdsub_Comments.Focus();
                    });
                }
            }
        }


        #endregion


        #region 키보드이벤트(유저편의)
        //데이터그리드 키 다운 이벤트.
        private void dgdAdjust_sub_KeyDown(object sender, KeyEventArgs e)
        {
            if (EventLabel.Visibility == Visibility.Visible)
            {
                if (e.Key == Key.Enter)
                {
                    if (((DataGridCell)sender).Column.Header.ToString() == "품번")
                    {
                        int point = dgdAdjust_sub.Items.IndexOf(dgdAdjust_sub.CurrentCell.Item);
                        dgdAdjust_sub_MouseDoubleClick(null, null);
                    }
                    else if (((DataGridCell)sender).Column.Header.ToString() == "수량")
                    {
                        int point = dgdAdjust_sub.Items.IndexOf(dgdAdjust_sub.CurrentCell.Item);
                        ContentRender(null, null);
                    }
                    else if (((DataGridCell)sender).Column.Header.ToString() == "비고")
                    {
                        int point = dgdAdjust_sub.Items.IndexOf(dgdAdjust_sub.CurrentCell.Item);
                        ContentRender1(null, null);
                    }
                }
                else if ((e.Key == Key.Delete) && (dgdAdjust_sub.CurrentColumn == dgdAdjust_sub.Columns[0]))
                {
                    btnMinus_Click(null, null);

                    if (dgdAdjust_sub.Items.Count > 0)
                    {
                        dgdAdjust_sub.Focus();
                        dgdAdjust_sub.CurrentCell = new DataGridCellInfo(dgdAdjust_sub.Items[dgdAdjust_sub.Items.Count - 1], dgdAdjust_sub.Columns[0]);
                    }
                }
            }
        }

        // 서브 그리드 내부 템플컬럼 자체에서 키 이벤트 먹이기.
        private void dgdAdjust_sub_TempleColumnsKeyDown(object sender, KeyEventArgs e)
        {
            if (EventLabel.Visibility == Visibility.Visible)
            {
                if (e.Key == Key.Enter)
                {
                    e.Handled = true;

                    if (((TextBox)sender).Tag.ToString() == "2")
                    {
                        int point = dgdAdjust_sub.Items.IndexOf(dgdAdjust_sub.CurrentCell.Item);
                        var ViewReceiver = dgdAdjust_sub.CurrentCell.Item as Win_ord_StockControl_U_View;  //선택 줄.
                        if (ViewReceiver != null)   // 널이 아니라면,
                        {
                            DataGridCell cell1 = Lib.Instance.GetCell(point, 3, dgdAdjust_sub);
                            TextBox tb1 = Lib.Instance.GetVisualChild<TextBox>(cell1);

                            ViewReceiver.ControlQty_s = tb1.Text;

                            int rowCount = dgdAdjust_sub.Items.IndexOf(dgdAdjust_sub.CurrentItem);
                            dgdAdjust_sub.Focus();
                            ContentRender1(null, null);
                            dgdAdjust_sub.CurrentCell = new DataGridCellInfo(dgdAdjust_sub.Items[rowCount], dgdAdjust_sub.Columns[5]);
                        }
                    }

                    if (((TextBox)sender).Tag.ToString() == "4")
                    {
                        int point = dgdAdjust_sub.Items.IndexOf(dgdAdjust_sub.CurrentCell.Item);
                        var ViewReceiver = dgdAdjust_sub.CurrentCell.Item as Win_ord_StockControl_U_View;  //선택 줄.
                        if (ViewReceiver != null)   // 널이 아니라면,
                        {
                            DataGridCell cell2 = Lib.Instance.GetCell(point, 5, dgdAdjust_sub);
                            TextBox tb2 = Lib.Instance.GetVisualChild<TextBox>(cell2);

                            ViewReceiver.Comments_s = tb2.Text;

                            if (point + 1 == dgdAdjust_sub.Items.Count)  // 마지막 줄 비고라면,
                            {
                                btnPlus_Click(null, null);
                                dgdAdjust_sub.CurrentCell = new DataGridCellInfo(dgdAdjust_sub.Items[point + 1], dgdAdjust_sub.Columns[0]);
                            }
                            else
                            { dgdAdjust_sub.CurrentCell = new DataGridCellInfo(dgdAdjust_sub.Items[point], dgdAdjust_sub.Columns[0]); }
                        }

                    }
                }
            }
        }

        // 품명 텍스트박스 키 다운.(품번이 조회되도록 수정, 2020.03.18, 장가빈)
        private void txtdgdArticle_s_KeyDown(object sender, KeyEventArgs e)
        {
            if (EventLabel.Visibility == Visibility.Visible)
            {
                if (e.Key == Key.Enter)
                {
                    e.Handled = true;

                    var ViewReceiver = dgdAdjust_sub.CurrentCell.Item as Win_ord_StockControl_U_View;

                    if (ViewReceiver != null)
                    {
                        //내가 몇 번째 셀인지 기억해
                        ViewReceiver.nRow = dgdAdjust_sub.Items.IndexOf(dgdAdjust_sub.CurrentItem);

                        TextBox dgdAdjustsub_FindText = new TextBox();
                        dgdAdjustsub_FindText.Text = ((TextBox)sender).Text;
                        pf.ReturnCode(dgdAdjustsub_FindText, 84, "");

                        string FindText = dgdAdjustsub_FindText.Text.ToString();        // Article_s
                        if (dgdAdjustsub_FindText.Tag == null) { return; }
                        string FindTag = dgdAdjustsub_FindText.Tag.ToString();          // ArticleID_s


                        // ID로 UNITCLSS 입혀버리기.
                        string[] UnitClssData = GetUnitClssData(FindTag);


                        ViewReceiver.BuyerArticleNo = FindText;     //품번
                        ViewReceiver.ArticleID_s = FindTag;
                        ViewReceiver.UnitClss_string_s = UnitClssData[0];               //UnitClss_string_s
                        ViewReceiver.UnitClss_s = UnitClssData[1];                      //UnitClss_s
                        ViewReceiver.Article_s = UnitClssData[2];

                        dgdAdjust_sub.Items.Refresh();


                        if (FindTag != null)
                        {
                            //비교할 내 자신
                            string ArticleTag = FindTag;

                            int i = 0;

                            for (i = 0; i < dgdAdjust_sub.Items.Count; i++)
                            {
                                var SubData = dgdAdjust_sub.Items[i] as Win_ord_StockControl_U_View;

                                if (SubData.ArticleID_s == ArticleTag)
                                {
                                    if (i != ViewReceiver.nRow)  //select된 내 라인은 제외해야겠지
                                    {
                                        MessageBox.Show("등록하신 품번은 이미 " + (i + 1) + "번째 등록된 품번입니다.");

                                        //비워주기
                                        ArticleTag = "";
                                        ViewReceiver.BuyerArticleNo = "";
                                        ViewReceiver.ArticleID_s = "";
                                        ViewReceiver.Article_s = "";
                                        ViewReceiver.UnitClss_s = "";
                                        ViewReceiver.UnitClss_string_s = "";

                                        return;
                                    }
                                }
                            }
                        }

                        int rowCount = dgdAdjust_sub.Items.IndexOf(dgdAdjust_sub.CurrentItem);
                        dgdAdjust_sub.CurrentCell = new DataGridCellInfo(dgdAdjust_sub.Items[rowCount], dgdAdjust_sub.Columns[3]);

                        ContentRender(null, null);
                    }
                }
            }
        }

        #endregion


        // 닫기 버튼 클릭.
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


        #region 엑셀
        //엑셀버튼 클릭.
        private void btnExcel_Click(object sender, RoutedEventArgs e)
        {
            if (dgdAdjust.Items.Count < 1)
            {
                MessageBox.Show("먼저 검색해 주세요.");
                return;
            }

            Lib lib2 = new Lib();
            DataTable dt = null;
            string Name = string.Empty;

            string[] lst = new string[4];
            lst[0] = "메인그리드";
            lst[1] = "서브그리드";
            lst[2] = dgdAdjust.Name;
            lst[3] = dgdAdjust_sub.Name;


            ExportExcelxaml ExpExc = new ExportExcelxaml(lst);

            ExpExc.ShowDialog();

            try
            {
                if (ExpExc.DialogResult.HasValue)
                {
                    if (ExpExc.choice.Equals(dgdAdjust.Name))
                    {
                        //MessageBox.Show("대분류");
                        if (ExpExc.Check.Equals("Y"))
                            dt = lib2.DataGridToDTinHidden(dgdAdjust);
                        else
                            dt = lib2.DataGirdToDataTable(dgdAdjust);

                        Name = dgdAdjust.Name;

                        if (lib2.GenerateExcel(dt, Name))
                        {
                            lib2.excel.Visible = true;
                            lib2.ReleaseExcelObject(lib2.excel);
                        }
                    }
                    else if (ExpExc.choice.Equals(dgdAdjust_sub.Name))
                    {
                        //MessageBox.Show("소분류");
                        if (ExpExc.Check.Equals("Y"))
                            dt = lib2.DataGridToDTinHidden(dgdAdjust_sub);
                        else
                            dt = lib2.DataGirdToDataTable(dgdAdjust_sub);

                        Name = dgdAdjust_sub.Name;

                        if (lib2.GenerateExcel(dt, Name))
                        {
                            lib2.excel.Visible = true;
                            lib2.ReleaseExcelObject(lib2.excel);
                        }
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
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally
            {
                lib2 = null;
            }

        }

        #endregion


        #region 텍스트박스 엔터 키 이동

        private void dtpAdjustDate_PreviewKeyUp(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                dtpAdjustDate.IsDropDownOpen = true;
            }
        }
        private void dtpAdjustDate_CalendarClosed(object sender, RoutedEventArgs e)
        {
            Lib.Instance.SendK(Key.Tab, this);
        }
        private void txtWorker_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                btnWorker_Click(null, null);
            }
        }

        // 엔터 키를 통한 탭 인덱스 키 이동.
        private void EnterMove_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                Lib.Instance.SendK(Key.Tab, this);
            }
        }


        #endregion


    }



    class Win_ord_StockControl_U_View : BaseView
    {
        public override string ToString()
        {
            return (this.ReportAllProperties());
        }

        public ObservableCollection<CodeView> cboTrade { get; set; }

        // 조회 값.    
        public string ControlID { get; set; }
        public string ControlDate { get; set; }
        public string Name { get; set; }
        public string PersonID { get; set; }
        public string Comments { get; set; }

        public string ArticleID { get; set; }
        public string Article { get; set; }
        public string Full_ControlDate { get; set; }


        //서브 조회 값.
        public string ControlID_s { get; set; }
        public string ControlSeq_s { get; set; }
        public string ArticleID_s { get; set; }
        public string Sabun_s { get; set; }

        public string ControlQty_s { get; set; }
        public string UnitClss_s { get; set; }

        public string UnitClss_string_s { get; set; }

        public string Comments_s { get; set; }
        public string BuyerArticleNo { get; set; }
        public string Article_s { get; set; }
        public string StuffINID_s { get; set; }

        public int nRow { get; set; }

    }

}

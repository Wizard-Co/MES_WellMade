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
    /// Win_sbl_LotStockSilsa_U.xaml에 대한 상호 작용 논리
    /// </summary>
    public partial class Win_sbl_LotStockSilsa_U : UserControl
    {
        public Win_sbl_LotStockSilsa_U()
        {
            InitializeComponent();
        }

        Lib lib = new Lib();
        PlusFinder pf = new PlusFinder();

        int rowNum = 0;
        string strFlag = string.Empty;

        Win_sbl_LotStockSilsa_U_CodeView SilsaData = new Win_sbl_LotStockSilsa_U_CodeView();

        private void UserControl_Loaded(object sender, RoutedEventArgs e)
        {
            SetCombobox();

            dtpFromDate.SelectedDate = DateTime.Today;
            dtpToDate.SelectedDate = DateTime.Today;
        }

        //콤보박스 셋팅
        private void SetCombobox()
        {
            cboLoc.Items.Clear();
            cboLoc.IsEnabled = false;

            cboLocSave.Items.Clear();
            cboLocSave.IsEnabled = false;

            //창고
            ObservableCollection<CodeView> cbWareHouse = ComboBoxUtil.Instance.Gf_DB_CM_GetComCodeDataset(null, "LOC", "Y", "", "");

            this.cboLoc.ItemsSource = cbWareHouse;
            this.cboLoc.DisplayMemberPath = "code_name";
            this.cboLoc.SelectedValuePath = "code_id";


            //창고
            ObservableCollection<CodeView> cbWareHouseSave = ComboBoxUtil.Instance.Gf_DB_CM_GetComCodeDataset(null, "LOC", "Y", "", "");

            this.cboLocSave.ItemsSource = cbWareHouse;
            this.cboLocSave.DisplayMemberPath = "code_name";
            this.cboLocSave.SelectedValuePath = "code_id";


        }

        #region 상단조건

        //실사일자 라벨
        private void Label_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            if (chkSilsaDate.IsChecked == true)
            {
                chkSilsaDate.IsChecked = false;
                dtpFromDate.IsEnabled = false;
                dtpToDate.IsEnabled = false;
            }
            else
            {
                chkSilsaDate.IsChecked = true;
                dtpFromDate.IsEnabled = true;
                dtpToDate.IsEnabled = true;
            }
        }

        //실사일자 체크
        private void ChkSilsaDate_Checked(object sender, RoutedEventArgs e)
        {
            if (dtpFromDate != null && dtpToDate != null)
            {
                dtpFromDate.IsEnabled = true;
                dtpToDate.IsEnabled = true;
            }
        }

        //실사일자 체크해제
        private void ChkSilsaDate_Unchecked(object sender, RoutedEventArgs e)
        {
            dtpFromDate.IsEnabled = false;
            dtpToDate.IsEnabled = false;
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

        //품명 라벨
        private void Article_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
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

        //품명 체크
        private void ChkArticle_Checked(object sender, RoutedEventArgs e)
        {
            txtArticle.IsEnabled = true;
            btnArticle.IsEnabled = true;
        }

        //품명 체크 해제
        private void ChkArticle_Unchecked(object sender, RoutedEventArgs e)
        {
            txtArticle.IsEnabled = false;
            btnArticle.IsEnabled = false;
        }

        //품명(검색조건) 키다운
        private void TxtArticleSrh_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                MainWindow.pf.ReturnCode(txtArticle, 76, "");
            }
        }

        //품명(검색조건) 플러스 파인더
        private void btnArticleSrh_Click(object sender, RoutedEventArgs e)
        {
            MainWindow.pf.ReturnCode(txtArticle, 76, "");
        }

        #endregion 상단조건

        #region CRUD 

        //추가
        private void btnAdd_Click(object sender, RoutedEventArgs e)
        {
            CantBtnControl();

            strFlag = "I";
            this.DataContext = null;
            dgdTotal.Items.Clear();
            if (dgdSub.Items.Count > 0)
            {
                dgdSub.Items.Clear();

            }
            txtArticleNo.Tag = "";
            txtArticleNo.Text = "";


            //조사일자는 당일날짜 기본 셋팅..
            dtpCheckFromDate.SelectedDate = DateTime.Today;
            dtpCheckToDate.SelectedDate = DateTime.Today;

            //실사처리일자도 당일날짜 기본 셋팅..
            dtpSilsaDate.SelectedDate = DateTime.Today;

            //작업자는 로그인한 아이디로 기본 셋팅..
            txtWorker.Tag = MainWindow.CurrentUser;
            txtWorker.Text = MainWindow.CurrentPerson;

            //창고는 사내창고 기본 셋팅
            cboLoc.SelectedIndex = 0;
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
                Thread.Sleep(2000);

                //로직
                re_Search(rowNum);

            }), System.Windows.Threading.DispatcherPriority.Background);



            Dispatcher.BeginInvoke(new Action(() =>

            {
                btnSearch.IsEnabled = true;

            }), System.Windows.Threading.DispatcherPriority.Background);

        }

        //엑셀
        private void btnExcel_Click(object sender, RoutedEventArgs e)
        {
            DataTable dt = null;
            string Name = string.Empty;

            string[] lst = new string[4];
            lst[0] = "실사처리목록";
            lst[1] = "실사처리상세내역";
            lst[2] = dgdMain.Name;
            lst[3] = dgdSub.Name;

            ExportExcelxaml ExpExc = new ExportExcelxaml(lst);
            ExpExc.ShowDialog();
            try
            {
                if (ExpExc.DialogResult.HasValue)
                {
                    if (ExpExc.choice.Equals(dgdMain.Name))
                    {
                        if (ExpExc.Check.Equals("Y"))
                            dt = Lib.Instance.DataGridToDTinHidden(dgdMain);
                        else
                            dt = Lib.Instance.DataGirdToDataTable(dgdMain);

                        Name = dgdMain.Name;
                        if (Lib.Instance.GenerateExcel(dt, Name))
                            Lib.Instance.excel.Visible = true;
                    }
                    else if (ExpExc.choice.Equals(dgdSub.Name))
                    {
                        if (ExpExc.Check.Equals("Y"))
                            dt = Lib.Instance.DataGridToDTinHidden(dgdSub);
                        else
                            dt = Lib.Instance.DataGirdToDataTable(dgdSub);

                        Name = dgdSub.Name;
                        if (Lib.Instance.GenerateExcel(dt, Name))
                            Lib.Instance.excel.Visible = true;
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
        }

        //저장
        private void btnSave_Click(object sender, RoutedEventArgs e)
        {
            if (SaveData(strFlag, txtControlID.Text))
            {
                //Lib.Instance.DBReIndex(); //2021-11-10 재고 실사 후 DBReIndex
                CanBtnControl();

                //방금 저장한 데이터가 선택되도록..
                rowNum = dgdMain.Items.Count + 1;

                re_Search(rowNum);

                //저장이 성공 했다면, FillGrid 후에 strFlag 비워주기 
                strFlag = string.Empty;
            }
        }

        //취소
        private void btnCancel_Click(object sender, RoutedEventArgs e)
        {
            strFlag = string.Empty;

            //메인에 아무 데이터가 없다면
            if (dgdMain.Items.Count == 0)
            {
                dgdSub.Items.Clear();
            }

            CanBtnControl();
            dtpCheckFromDate.SelectedDate = DateTime.Today;
            dtpCheckToDate.SelectedDate = DateTime.Today;
            re_Search(rowNum);
        }

        #endregion CRUD

        #region 버튼 컨트롤

        //수정, 추가 저장 후
        private void CanBtnControl()
        {
            btnAdd.IsEnabled = true;
            btnSearch.IsEnabled = true;
            btnSave.Visibility = Visibility.Hidden;
            btnCancel.Visibility = Visibility.Hidden;
            btnExcel.Visibility = Visibility.Visible;
            EventLabel.Visibility = Visibility.Hidden;
            dgdMain.IsHitTestVisible = true;
            dgdSub.IsHitTestVisible = true;
            dtpSilsaDate.IsHitTestVisible = false;      //실사처리 날짜
            txtWorker.IsHitTestVisible = false;
            btnWorker.IsHitTestVisible = false;
            txtReason.IsHitTestVisible = false;
            btnMinus.IsEnabled = false;

            btnAllChoice.IsEnabled = false;             //전체선택
            //dtpCheckFromDate.IsEnabled = false;         //조사일자
            //dtpCheckToDate.IsEnabled = false;
            btnSilsaList.IsEnabled = false;             //대상조회
            //txtArticleNo.IsEnabled = false;
            //lblCheckText.Visibility = Visibility.Hidden;    //조사일자 라벨
            //dtpCheckToDate.Visibility = Visibility.Hidden;  //조사일자 데이트피커
            //dtpCheckFromDate.Visibility = Visibility.Hidden; //조사일자 데이트피커
            //txbCheckText.Visibility = Visibility.Hidden;    //조사일자 물결표시
            //lblCheckText2.Visibility = Visibility.Hidden;
            //txtArticleNo.Visibility = Visibility.Hidden;
            btnSilsaList2.Visibility = Visibility.Visible;
            cboLoc.IsEnabled = false;

            dgdSub.Items.Clear();
            dgdTotal.Items.Clear();

        }

        //수정, 추가 진행 중
        private void CantBtnControl()
        {
            btnAdd.IsEnabled = false;
            btnSearch.IsEnabled = false;
            btnSave.Visibility = Visibility.Visible;
            btnCancel.Visibility = Visibility.Visible;
            btnExcel.Visibility = Visibility.Hidden;
            EventLabel.Visibility = Visibility.Visible;
            EventLabel.Content = "자료 입력 중";
            dgdMain.IsHitTestVisible = false;
            dgdSub.IsHitTestVisible = true;

            dtpSilsaDate.IsHitTestVisible = true;
            txtWorker.IsHitTestVisible = true;
            btnWorker.IsHitTestVisible = true;
            txtReason.IsHitTestVisible = true;
            //btnStockCheck.IsEnabled = true;
            btnMinus.IsEnabled = true;
            txtArticleNo.IsEnabled = true;
            btnAllChoice.IsEnabled = true;
            dtpCheckFromDate.IsEnabled = true;      //조사일자
            dtpCheckToDate.IsEnabled = true;
            btnSilsaList.IsEnabled = true;          //대상조회

            lblCheckText2.Visibility = Visibility.Visible;
            txtArticleNo.Visibility = Visibility.Visible;
            lblCheckText.Visibility = Visibility.Visible;    //조사일자 라벨
            dtpCheckToDate.Visibility = Visibility.Visible;  //조사일자 데이트피커
            dtpCheckFromDate.Visibility = Visibility.Visible; //조사일자 데이트피커
            txbCheckText.Visibility = Visibility.Visible;    //조사일자 물결표시
            btnSilsaList2.Visibility = Visibility.Hidden;
            cboLoc.IsEnabled = true;
        }

        #endregion 버튼 컨트롤


        #region 조회

        private void re_Search(int rowNum)
        {
            FillGrid();

            if (dgdMain.Items.Count > 0)
            {
                dgdMain.SelectedIndex = rowNum;
            }


        }

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

                sqlParameter.Add("nChkDate", chkSilsaDate.IsChecked == true ? 1 : 0);
                sqlParameter.Add("sSDate", chkSilsaDate.IsChecked == true && dtpFromDate.SelectedDate != null
                                                                        ? dtpFromDate.SelectedDate.Value.ToString("yyyyMMdd") : "");
                sqlParameter.Add("sEDate", chkSilsaDate.IsChecked == true && dtpToDate.SelectedDate != null
                                                                        ? dtpToDate.SelectedDate.Value.ToString("yyyyMMdd") : "");
                sqlParameter.Add("sArticleID", chkArticle.IsChecked == true && txtArticle.Tag != null ? txtArticle.Tag.ToString() : "");


                DataSet ds = DataStore.Instance.ProcedureToDataSet("xp_sbStockCtl_sStockControlSilsa", sqlParameter, false);

                if (ds != null && ds.Tables.Count > 0)
                {
                    DataTable dt = ds.Tables[0];

                    if (dt.Rows.Count == 0)
                    {
                        MessageBox.Show("조회된 데이터가 없습니다.");
                        if (dgdSub.Items.Count > 0)
                        {
                            dgdSub.Items.Clear();
                        }
                        txtControlID.Text = "";
                        txtWorker.Text = "";
                        txtReason.Text = "";
                        dtpSilsaDate.SelectedDate = null;
                    }
                    else
                    {
                        int i = 0;
                        DataRowCollection drc = dt.Rows;

                        foreach (DataRow dr in drc)
                        {
                            i++;
                            var SilsaData = new Win_sbl_LotStockSilsa_U_CodeView
                            {
                                Num = i,

                                ctlSilsaID = dr["ctlSilsaID"].ToString(),
                                CtlSilsalDate = DatePickerFormat(dr["CtlSilsalDate"].ToString()),
                                SilsaStartDate = DatePickerFormat(dr["SilsaStartDate"].ToString()),
                                SilsaEndDate = DatePickerFormat(dr["SilsaEndDate"].ToString()),
                                PersonID = dr["PersonID"].ToString(),
                                PersonName = dr["PersonName"].ToString(),
                                ArticleQty = dr["ArticleQty"].ToString(),
                                CtlSilsaLocID = dr["CtlSilsaLocID"].ToString(),
                                LOCName = dr["LOCName"].ToString(),
                                Comments = dr["Comments"].ToString(),
                            };

                            //조사기간
                            SilsaData.CheckPeriod = SilsaData.SilsaStartDate + " ~ " + SilsaData.SilsaEndDate;

                            dgdMain.Items.Add(SilsaData);
                        }
                        tbkCount.Text = "▶ 검색 결과 : " + i + "건";
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


        private void FillGridSub(string StrID)
        {
            //실사더하기
            var SilsaSum = new Win_mtr_StockSils_Sum();

            //if (dgdSub.Items.Count > 0) 2021-07-13
            //{
            dgdSub.Items.Clear();
            dgdTotal.Items.Clear();
            //}

            try
            {
                Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                sqlParameter.Clear();

                sqlParameter.Add("CtlSilsaID", StrID);

                DataSet ds = DataStore.Instance.ProcedureToDataSet("xp_sbStockCtl_sStockControlSilsaSub", sqlParameter, false);

                if (ds != null && ds.Tables.Count > 0)
                {
                    DataTable dt = ds.Tables[0];

                    if (dt.Rows.Count == 0)
                    {
                        MessageBox.Show("조회된 데이터가 없습니다.");
                    }
                    else
                    {
                        int i = 0;
                        DataRowCollection drc = dt.Rows;

                        foreach (DataRow dr in drc)
                        {
                            i++;
                            var SilsaDataSub = new Win_sbl_LotStockSilsa_U_CodeView_Sub
                            {
                                Num = i,

                                SilsaID = dr["SilsaID"].ToString(),
                                SilsaSeq = dr["SilsaSeq"].ToString(),
                                SilsaDate = DatePickerFormat(dr["SilsaDate"].ToString()),
                                LocID = dr["LocID"].ToString(),
                                LOCName = dr["LOCName"].ToString(),
                                LOTID = dr["LOTID"].ToString(),

                                ArticleID = dr["ArticleID"].ToString(),
                                Article = dr["Article"].ToString(),
                                BuyerArticleNo = dr["BuyerArticleNo"].ToString(),

                                StockQty = stringFormatN0(dr["StockQty"]),
                                UnitClss = dr["UnitClss"].ToString(),
                                UnitclssName = dr["UnitclssName"].ToString(),
                                SilsaStockQty = stringFormatN0(dr["SilsaStockQty"]),
                                PersonID = dr["PersonID"].ToString(),
                                PersonName = dr["PersonName"].ToString(),

                                Comments = dr["Comments"].ToString(),

                                Chk = false,
                            };

                            dgdSub.Items.Add(SilsaDataSub);

                            SilsaSum.SilsaStockQtyAmount += ConvertDouble(SilsaDataSub.SilsaStockQty);
                            SilsaSum.StockQtyAmount += ConvertDouble(SilsaDataSub.StockQty); //2021-07-13
                        }

                        dgdTotal.Items.Add(SilsaSum);
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


        //대상조회 
        private void FillGrid_SilsaList()
        {
            //실사더하기
            var SilsaSum = new Win_mtr_StockSils_Sum();

            try
            {
                //if (dgdSub.Items.Count > 0)
                //{
                dgdSub.Items.Clear();
                dgdTotal.Items.Clear();
                //}

                Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                sqlParameter.Clear();

                sqlParameter.Add("nChkDate", 1); //날짜체크는 무조건 되어 있으니 1 
                sqlParameter.Add("sSDate", dtpCheckFromDate.SelectedDate != null ? dtpCheckFromDate.SelectedDate.Value.ToString("yyyyMMdd") : "");
                sqlParameter.Add("sEDate", dtpCheckToDate.SelectedDate != null ? dtpCheckToDate.SelectedDate.Value.ToString("yyyyMMdd") : "");
                sqlParameter.Add("sLocID", cboLoc.SelectedValue != null ? cboLoc.SelectedValue.ToString() : "");
                sqlParameter.Add("sArticleNo", txtArticleNo.Text);

                sqlParameter.Add("ChkLotID", txtLotID.Text == "" ? 0 : 1);
                sqlParameter.Add("LotID", txtLotID.Text);
                DataSet ds = DataStore.Instance.ProcedureToDataSet("xp_sbStockCtl_sStockSilsa_20210312", sqlParameter, false);

                if (ds != null && ds.Tables.Count > 0)
                {
                    DataTable dt = ds.Tables[0];

                    if (dt.Rows.Count == 0)
                    {
                        MessageBox.Show("조회된 데이터가 없습니다.");
                    }
                    else
                    {
                        int i = 0;
                        DataRowCollection drc = dt.Rows;

                        foreach (DataRow dr in drc)
                        {
                            i++;
                            var SilsaList = new Win_sbl_LotStockSilsa_U_CodeView_Sub
                            {
                                Num = i,

                                SilsaID = dr["SilsaID"].ToString(),
                                SilsaSeq = dr["SilsaSeq"].ToString(),
                                SilsaDate = DatePickerFormat(dr["SilsaDate"].ToString()),
                                LocID = dr["LocID"].ToString(),
                                LOCName = dr["LOCName"].ToString(),
                                LOTID = dr["LOTID"].ToString(),

                                ArticleID = dr["ArticleID"].ToString(),
                                Article = dr["Article"].ToString(),
                                BuyerArticleNo = dr["BuyerArticleNo"].ToString(),

                                StockQty = stringFormatN0(dr["StockQty"]),
                                UnitClss = dr["UnitClss"].ToString(),
                                UnitclssName = dr["UnitclssName"].ToString(),
                                SilsaStockQty = stringFormatN0(dr["SilsaStockQty"]),
                                PersonID = dr["PersonID"].ToString(),
                                PersonName = dr["PersonName"].ToString(),

                                Comments = dr["Comments"].ToString(),

                                //대상 불러왔을 때는 해제여야 하겠지.
                                Chk = false,
                            };

                            dgdSub.Items.Add(SilsaList);
                            SilsaSum.SilsaStockQtyAmount += ConvertDouble(SilsaList.SilsaStockQty);
                            SilsaSum.StockQtyAmount += ConvertDouble(SilsaList.StockQty); //2021-07-13
                        }

                        dgdTotal.Items.Add(SilsaSum);
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



        private void FillGridSub2(string StrID)
        {
            //실사더하기
            var SilsaSum = new Win_mtr_StockSils_Sum();

            //if (dgdSub.Items.Count > 0)
            //{
            dgdSub.Items.Clear();
            dgdTotal.Items.Clear();
            //}

            try
            {
                Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                sqlParameter.Clear();
                sqlParameter.Add("CtlSilsaID", StrID);


                sqlParameter.Add("chkDate", 1);
                sqlParameter.Add("sFromDate", dtpCheckFromDate.SelectedDate != null ? dtpCheckFromDate.SelectedDate.Value.ToString("yyyyMMdd") : "");
                sqlParameter.Add("sTODate", dtpCheckToDate.SelectedDate != null ? dtpCheckToDate.SelectedDate.Value.ToString("yyyyMMdd") : "");
                sqlParameter.Add("ChkArticle", txtArticleNo.Text == null ? 0 : 1);
                sqlParameter.Add("BuyerArticleNo", txtArticleNo.Tag);
                sqlParameter.Add("ChkLotID", txtLotID.Text == "" ? 0 : 1);
                sqlParameter.Add("LotID", txtLotID.Text);

                DataSet ds = DataStore.Instance.ProcedureToDataSet("xp_sbStockCtl_sStockControlSilsaSub_WPF", sqlParameter, false);

                if (ds != null && ds.Tables.Count > 0)
                {
                    DataTable dt = ds.Tables[0];

                    if (dt.Rows.Count == 0)
                    {
                        MessageBox.Show("조회된 데이터가 없습니다.");
                    }
                    else
                    {
                        int i = 0;
                        DataRowCollection drc = dt.Rows;

                        foreach (DataRow dr in drc)
                        {
                            i++;
                            var SilsaDataSub = new Win_sbl_LotStockSilsa_U_CodeView_Sub
                            {
                                Num = i,

                                SilsaID = dr["SilsaID"].ToString(),
                                SilsaSeq = dr["SilsaSeq"].ToString(),
                                SilsaDate = DatePickerFormat(dr["SilsaDate"].ToString()),
                                LocID = dr["LocID"].ToString(),
                                LOCName = dr["LOCName"].ToString(),
                                LOTID = dr["LOTID"].ToString(),

                                ArticleID = dr["ArticleID"].ToString(),
                                Article = dr["Article"].ToString(),
                                BuyerArticleNo = dr["BuyerArticleNo"].ToString(),

                                StockQty = stringFormatN0(dr["StockQty"]),
                                UnitClss = dr["UnitClss"].ToString(),
                                UnitclssName = dr["UnitclssName"].ToString(),
                                SilsaStockQty = stringFormatN0(dr["SilsaStockQty"]),
                                PersonID = dr["PersonID"].ToString(),
                                PersonName = dr["PersonName"].ToString(),

                                Comments = dr["Comments"].ToString(),

                                Chk = false,
                            };

                            dgdSub.Items.Add(SilsaDataSub);

                            SilsaSum.SilsaStockQtyAmount += ConvertDouble(SilsaDataSub.SilsaStockQty);
                            SilsaSum.StockQtyAmount += ConvertDouble(SilsaDataSub.StockQty); //2021-07-13
                        }

                        dgdTotal.Items.Add(SilsaSum);
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

        #endregion 조회


        #region 유효성 검사 CheckData

        private bool CheckData()
        {
            bool flag = true;

            if (dtpSilsaDate.SelectedDate == null)
            {
                MessageBox.Show("실사일자는 필수입력 항목입니다. 날짜를 선택해주세요.");
                flag = false;
                return flag;
            }

            if (txtWorker.Text.Length <= 0 || txtWorker.Tag == null)
            {
                MessageBox.Show("작업자가 입력되지 않았습니다. 작업자를 입력해주세요.");
                flag = false;
                return flag;
            }

            if (txtReason.Text.Length <= 0)
            {
                MessageBox.Show("실사처리 사유는 필수입력 항목입니다. 사유를 입력해주세요.");
                flag = false;
                return flag;
            }

            if (dgdSub.Items.Count == 0)
            {
                MessageBox.Show("오른쪽 하단부 실사처리할 항목이 조회되지 않았습니다. 날짜 지정 후 대상 조회를 눌러주세요. ");
                flag = false;
                return flag;
            }


            int sum = 0;

            for (int i = 0; i < dgdSub.Items.Count; i++)
            {
                var SilsaStockControl = dgdSub.Items[i] as Win_sbl_LotStockSilsa_U_CodeView_Sub;

                if (SilsaStockControl != null)
                {
                    //체크된 것만 실사처리 적용.
                    if (SilsaStockControl.Chk == true)
                    {
                        sum++;
                    }
                }

            }

            if (sum == 0)
            {
                MessageBox.Show("실사를 적용할 체크된 항목이 없습니다. 적용할 항목을 체크박스 체크해주세요.");
                flag = false;
                return flag;
            }
            return flag;
        }

        #endregion 유효성 검사 CheckData


        #region 저장 SaveData

        private bool SaveData(string strFlag, string strID)
        {
            bool flag = false;
            List<Procedure> Prolist = new List<Procedure>();
            List<Dictionary<string, object>> ListParameter = new List<Dictionary<string, object>>();

            string GetKey = string.Empty;

            try
            {
                if (CheckData())
                {
                    Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                    sqlParameter.Clear();

                    sqlParameter.Add("sControlID", "");
                    sqlParameter.Add("CtlSilsalDate", dtpSilsaDate.SelectedDate != null ? dtpSilsaDate.SelectedDate.Value.ToString("yyyyMMdd") : "");
                    sqlParameter.Add("CtlSilsaPersonID", txtWorker.Tag != null ? txtWorker.Tag.ToString() : "");
                    sqlParameter.Add("CtlSilsaLocID", cboLocSave.SelectedValue != null ? cboLocSave.SelectedValue.ToString() : "");
                    sqlParameter.Add("Comments", txtReason.Text);

                    if (strFlag.Equals("I"))
                    {
                        sqlParameter.Add("CreateUserID", MainWindow.CurrentUser);

                        Procedure pro1 = new Procedure();
                        pro1.Name = "xp_sbStockCtl_iStockSilsa";
                        pro1.OutputUseYN = "Y";
                        pro1.OutputName = "sControlID";
                        pro1.OutputLength = "12";

                        Prolist.Add(pro1);
                        ListParameter.Add(sqlParameter);


                        //동운씨가 만든 아웃풋 값 찾는 방법
                        List<KeyValue> list_Result = new List<KeyValue>();
                        list_Result = DataStore.Instance.ExecuteAllProcedureOutputGetCS(Prolist, ListParameter);

                        Prolist.RemoveAt(0);
                        ListParameter.RemoveAt(0);

                        if (list_Result[0].key.ToLower() == "success")
                        {
                            list_Result.RemoveAt(0);
                            for (int i = 0; i < list_Result.Count; i++)
                            {
                                KeyValue kv = list_Result[i];
                                if (kv.key == "sControlID")
                                {
                                    GetKey = kv.value;
                                    flag = true;
                                }
                            }
                            //flag = true;
                        }
                        else
                        {
                            MessageBox.Show("[저장실패]\r\n" + list_Result[0].value.ToString());
                            flag = false;
                        }

                        //for문 타면서 서브그리드 저장
                        for (int i = 0; i < dgdSub.Items.Count; i++)
                        {
                            var SilsaStockControl = dgdSub.Items[i] as Win_sbl_LotStockSilsa_U_CodeView_Sub;

                            if (SilsaStockControl != null)
                            {
                                //체크된 것만 실사처리 적용.
                                if (SilsaStockControl.Chk == true)
                                {
                                    //클리어만 해서는 소용이 없어, 이건 무조건 한 세트야
                                    sqlParameter = new Dictionary<string, object>();
                                    sqlParameter.Clear();

                                    sqlParameter.Add("CtlSilsaID", GetKey);
                                    sqlParameter.Add("LotID", SilsaStockControl.LOTID);
                                    sqlParameter.Add("ArticleID", SilsaStockControl.ArticleID);
                                    sqlParameter.Add("CtrlSilsaStockQty", SilsaStockControl.SilsaStockQty.Replace(",", ""));     //조정재고 = 실사처리재고
                                    sqlParameter.Add("nStockQty", SilsaStockControl.StockQty.Replace(",", ""));      //현재고

                                    sqlParameter.Add("SilsaID", SilsaStockControl.SilsaID);
                                    sqlParameter.Add("SilsaSeq", SilsaStockControl.SilsaSeq);
                                    //sqlParameter.Add("LocID", SilsaStockControl.LocID);
                                    sqlParameter.Add("Comments", SilsaStockControl.Comments);
                                    sqlParameter.Add("CreateUserID", MainWindow.CurrentUser);


                                    Procedure pro2 = new Procedure();
                                    pro2.Name = "xp_sbStockCtl_iStockSilsaSub";
                                    pro2.OutputUseYN = "N";
                                    pro2.OutputName = "sControlID";
                                    pro2.OutputLength = "12";

                                    Prolist.Add(pro2);
                                    ListParameter.Add(sqlParameter);
                                }
                            }
                        }

                        //클리어만 해서는 소용이 없어, 이건 무조건 한 세트야
                        //서브그리드 저장  실사재고 안된거 0
                        sqlParameter = new Dictionary<string, object>();
                        sqlParameter.Clear();

                        sqlParameter.Add("CtrlSilSaID", GetKey);
                        sqlParameter.Add("CtlSilsaLocID", cboLocSave.SelectedValue != null ? cboLocSave.SelectedValue.ToString() : "");
                        sqlParameter.Add("CreateUserID", MainWindow.CurrentUser);

                        Procedure pro3 = new Procedure();
                        pro3.Name = "xp_sbStockCtl_iStockSilsaLast";
                        pro3.OutputUseYN = "N";
                        pro3.OutputName = "sControlID";
                        pro3.OutputLength = "12";

                        Prolist.Add(pro3);
                        ListParameter.Add(sqlParameter);


                        string[] confirm = new string[2];
                        confirm = DataStore.Instance.ExecuteAllProcedureOutputNew(Prolist, ListParameter);

                        if (confirm[0] == "success")
                        {
                            //MessageBox.Show("성공");
                            flag = true;
                        }
                        else
                        {
                            MessageBox.Show("실패 : " + confirm[1]);
                            flag = false;
                        }

                        //dgdSub.Items.Clear();
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

            return flag;
        }

        #endregion 저장 SaveData



        //메인 데이터 그리드 
        private void DgdMain_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            try
            {
                SilsaData = dgdMain.SelectedItem as Win_sbl_LotStockSilsa_U_CodeView;

                if (SilsaData != null)
                {
                    this.DataContext = SilsaData;

                    txtWorker.Tag = SilsaData.PersonID;

                    FillGridSub(SilsaData.ctlSilsaID);
                    dtpCheckFromDate.SelectedDate = DateTime.Today;
                    dtpCheckToDate.SelectedDate = DateTime.Today;

                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }

        //작업자 키다운
        private void txtWorker_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                pf.ReturnCode(txtWorker, 2, txtWorker.Text);
            }
        }

        //작업자 플러스 파인더
        private void btnWorker_Click(object sender, RoutedEventArgs e)
        {
            pf.ReturnCode(txtWorker, 2, txtWorker.Text);
        }

        //전체선택 버튼
        private void BtnAllChoice_Click(object sender, RoutedEventArgs e)
        {
            if (dgdSub.Items.Count > 0)
            {
                foreach (Win_sbl_LotStockSilsa_U_CodeView_Sub Silsadata in dgdSub.Items)
                {
                    if (Silsadata != null && Silsadata.Chk == false) //2021-07-12 전체선택 해제 추가
                    {
                        Silsadata.Chk = true;
                    }
                    else //2021-07-12 전체선택 해제 추가
                    {
                        Silsadata.Chk = false;
                    }
                }

                dgdSub.Items.Refresh();
            }
        }

        //대상조회 버튼
        private void BtnSilsaList_Click(object sender, RoutedEventArgs e)
        {

            try
            {
                FillGrid_SilsaList();

                // 대상조회하면 '창고 콤보박스'기준으로 save창고가 똑같이 조회됩니다.
                if (cboLoc.SelectedIndex == 0) //사내
                {
                    cboLocSave.SelectedIndex = 0;
                }
                if (cboLoc.SelectedIndex == 1)  //외주
                {
                    cboLocSave.SelectedIndex = 1;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }


        }

        //행 삭제 버튼
        private void btnMinus_Click(object sender, RoutedEventArgs e)
        {
            //2021-11-17 체크 표시 유무로 삭제하기 위해 로직 수정
            for (int i = dgdSub.Items.Count; i != 0; i--)
            {
                var Sub = dgdSub.Items[i - 1] as Win_sbl_LotStockSilsa_U_CodeView_Sub;

                if (Sub != null)
                {
                    if (Sub.Chk == true)
                    {
                        //2021-07-12
                        Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                        sqlParameter.Clear();
                        sqlParameter.Add("SilsaID", Sub.SilsaID.ToString());
                        sqlParameter.Add("CtlSilsaID", Sub.SilsaSeq.ToString());
                        try
                        {

                            string[] result = DataStore.Instance.ExecuteProcedure("xp_sbStockCtl_dStockSilsa", sqlParameter, true);
                            dgdSub.Items.Remove(Sub);

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

                }
            }
            #region 2021-11-17 이전 소스
            //var Sub = dgdSub.SelectedItem as Win_sbl_LotStockSilsa_U_CodeView_Sub;
            //if (Sub != null)
            //{
            //    //2021-07-12
            //    Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
            //    sqlParameter.Clear();
            //    sqlParameter.Add("SilsaID", Sub.SilsaID.ToString());
            //    sqlParameter.Add("CtlSilsaID", Sub.SilsaSeq.ToString()); ;
            //    try
            //    {

            //        string[] result = DataStore.Instance.ExecuteProcedure("xp_sbStockCtl_dStockSilsa", sqlParameter, true);
            //        dgdSub.Items.Remove(Sub);

            //    }
            //    catch (Exception ex)
            //    {
            //        MessageBox.Show(ex.ToString());
            //    }
            //    finally
            //    {
            //        DataStore.Instance.CloseConnection();
            //    }

            //}
            #endregion
        }

        //Sub데이터 그리드 체크박스 체크 이벤트
        private void chkItem_Click(object sender, RoutedEventArgs e)
        {
            CheckBox chkSender = sender as CheckBox;
            if (lblCheckText.Visibility == Visibility.Visible)
            {
                var dgdAll = chkSender.DataContext as Win_sbl_LotStockSilsa_U_CodeView_Sub;


                if (chkSender.IsChecked == true)
                {
                    dgdAll.Chk = true;
                }
                else
                {
                    dgdAll.Chk = false;
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
                MessageBox.Show("체크박스를 사용하려면 먼저 추가 버튼을 누르고 진행해야 합니다.");
            }
        }


        #region 기타 메서드 모음

        // 천 단위 콤마, 소수점 버리기
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

            try
            {
                if (str.Length == 8)
                {
                    if (!str.Trim().Equals(""))
                    {
                        result = str.Substring(0, 4) + "-" + str.Substring(4, 2) + "-" + str.Substring(6, 2);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }


            return result;
        }

        // 시간 : 분 으로 변환
        private string ConvertTimeFormat(string str)
        {
            string result = "";

            str = str.Trim().Replace(":", "");
            if (str.Length > 5)
            {
                string hour = str.Substring(0, 2);
                string min = str.Substring(2, 2);
                string sec = str.Substring(4, 2);

                result = hour + ":" + min + ":" + sec;
            }
            else if (str.Length > 3 && str.Length < 5)
            {
                string hour = str.Substring(0, 2);
                string min = str.Substring(2, 2);

                result = hour + ":" + min;
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

        //삭제버튼
        private void btnDelete_Click(object sender, RoutedEventArgs e)
        {
            var DeleteItem = dgdMain.SelectedItem as Win_sbl_LotStockSilsa_U_CodeView;

            if (DeleteItem != null)
            {
                if (MessageBox.Show("선택한 항목을 삭제하시겠습니까?", "삭제 전 확인", MessageBoxButton.YesNo) == MessageBoxResult.Yes)
                {
                    if (DeleteData(DeleteItem.ctlSilsaID))
                    {
                        re_Search(rowNum);
                    }
                }
            }
            else
            {
                MessageBox.Show("삭제할 데이터를 선택해주세요.");
            }
        }


        #region DeleteData

        private bool DeleteData(string OrderID)
        {
            bool flag = false;

            Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
            sqlParameter.Clear();
            sqlParameter.Add("CtlSilsaID", txtControlID.Text);


            try
            {
                //if (SilsaData.CtlSilsalDate.ToString().Replace("-", "") == DateTime.Now.ToString("yyyyMMdd")) //오늘날짜랑 비교해서 오늘날짜 아니면 삭제 불가능
                //{   
                string[] result = DataStore.Instance.ExecuteProcedure("xp_sbStockCtl_dStockControlSilsa", sqlParameter, true);


                if (!result[0].Equals("success"))
                {
                    MessageBox.Show("삭제 실패");
                    flag = false;
                }
                else
                {
                    //MessageBox.Show("성공적으로 삭제되었습니다.");
                    flag = true;
                }
                //}
                //else
                //{
                //    MessageBox.Show("오늘 날짜만 삭제할 수 있습니다.");

                //    return flag;
                //}



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

        #endregion DeleteData

        private void TxtArticleNoSrh_KeyDown(object sender, KeyEventArgs e)
        {

            if (e.Key == Key.Enter)
            {
                pf.ReturnCode(txtArticleNo, 84, txtArticleNo.Text);
            }
        }

        private void BtnSilsaList2_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (SilsaData.Num == 0)
                {
                    SilsaData.ctlSilsaID = "";
                }

                FillGridSub2(SilsaData.ctlSilsaID);
            }
            catch (Exception ex)
            {
                MessageBox.Show("조회된 데이터가 없습니다.");
            }
        }
    }



    #region 코드뷰  

    public class Win_sbl_LotStockSilsa_U_CodeView : BaseView
    {
        public int Num { get; set; }

        public string ctlSilsaID { get; set; }                //실사처리 코드
        public string CtlSilsalDate { get; set; }             //실사처리 일자
        public string SilsaStartDate { get; set; }            //조사기간 시작 일자
        public string SilsaEndDate { get; set; }              //조사기간 끝 일자

        public string CheckPeriod { get; set; }               //조사기간

        public string PersonID { get; set; }                  //작업자코드
        public string PersonName { get; set; }                //작업자명

        public string ArticleQty { get; set; }                //품명 수 (실사처리한 품명 종류 수 말하는 듯)
        public string CtlSilsaLocID { get; set; }             //창고코드
        public string LOCName { get; set; }                   //창고명
        public string Comments { get; set; }                  //비고



    }


    public class Win_sbl_LotStockSilsa_U_CodeView_Sub : BaseView
    {
        public int Num { get; set; }

        public bool Chk { get; set; }

        public string SilsaID { get; set; }               //실사코드
        public string SilsaSeq { get; set; }              //실사순서
        public string SilsaDate { get; set; }             //실사 일자
        public string LocID { get; set; }                 //창고코드
        public string LOCName { get; set; }               //창고명      
        public string LOTID { get; set; }                 //로트ID 

        public string ArticleID { get; set; }             //품명코드    
        public string Article { get; set; }               //품명
        public string BuyerArticleNo { get; set; }        //품번
        public string StockQty { get; set; }              //현재고

        public string UnitClss { get; set; }              //단위코드
        public string UnitclssName { get; set; }          //단위명     
        public string SilsaStockQty { get; set; }         //실사수량      
        public string PersonID { get; set; }              //조사작업자
        public string PersonName { get; set; }            //조사작업자명

        public string Comments { get; set; }              //비고       
        public string SilsaStockQtyAmount { get; set; }              //설정재고합       

        // 수정, 삭제가 가능한지 체크하는 변수
        public bool UDFlag { get; set; }
        public string LastDate { get; set; } // 이 라벨로 마지막에 조정재고 등록한 날짜가 언제인지

    }

    class Win_mtr_StockSils_Sum : BaseView
    {
        public double SilsaStockQtyAmount { get; set; }
        public double StockQtyAmount { get; set; }

    }

    #endregion 코드뷰
}

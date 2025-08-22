using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Data;
using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using WizMes_WellMade.PopUP;
using WPF.MDI;

namespace WizMes_WellMade
{
    /// <summary>
    /// Win_sbl_Outware_Q.xaml에 대한 상호 작용 논리
    /// </summary>
    public partial class Win_mtr_Outware_Q : UserControl
    {
        string stDate = string.Empty;
        string stTime = string.Empty;

        Win_mtr_Outware_QView WMOV = new Win_mtr_Outware_QView();
        Lib lib = new Lib();
        PlusFinder pf = new PlusFinder();
        int rowNum = 0;

        // 엑셀 활용 용도 (프린트)
        private Microsoft.Office.Interop.Excel.Application excelapp;
        private Microsoft.Office.Interop.Excel.Workbook workbook;
        private Microsoft.Office.Interop.Excel.Worksheet worksheet;
        private Microsoft.Office.Interop.Excel.Range workrange;
        private Microsoft.Office.Interop.Excel.Worksheet copysheet;
        private Microsoft.Office.Interop.Excel.Worksheet pastesheet;
        //(기다림 알림 메시지창)
        WizMes_WellMade.PopUp.NoticeMessage msg = new PopUp.NoticeMessage();

        // 출고 명세서 인쇄를 위해서
        // 순수 데이터 만 넣기(오더계, 일계, 거래처계, 총계 다 뺌)
        List<Win_mtr_Outware_QView> lstOutware = new List<Win_mtr_Outware_QView>();

        System.Data.DataTable DT;

        public Win_mtr_Outware_Q()
        {
            InitializeComponent();
            this.DataContext = WMOV;
        }

        private void Window_SubulOutware_Loaded(object sender, RoutedEventArgs e)
        {
            stDate = DateTime.Now.ToString("yyyyMMdd");
            stTime = DateTime.Now.ToString("HHmm");

            DataStore.Instance.InsertLogByFormS(this.GetType().Name, stDate, stTime, "S");

            First_Step();
            ComboBoxSetting();
        }

        #region 시작 첫단계 // 콤보박스 세팅 // 조회용 각종 체크박스 활성화 // 일자버튼

        // 시작 첫 단계.
        private void First_Step()
        {
            chkOutwareDay.IsChecked = true;
            dtpFromDate.Text = DateTime.Now.ToString("yyyy-MM-dd");
            dtpToDate.Text = DateTime.Now.ToString("yyyy-MM-dd");

            rbnManageNumber.IsChecked = true;

            // no check > no use.
            cboArticleGroup.IsEnabled = false;
            txtArticle.IsEnabled = false;
            // btnArticle.IsEnabled = false;
            txtOrderID.IsEnabled = false;
            cboFromLoc.IsEnabled = false;
            cboToLoc.IsEnabled = false;
            cboOutClss.IsEnabled = false;
        }

        // 어제.
        private void btnYesterday_Click(object sender, RoutedEventArgs e)
        {
            //string[] receiver = lib.BringYesterdayDatetime();

            //dtpFromDate.Text = receiver[0];
            //dtpToDate.Text = receiver[1];
            try
            {
                if (dtpFromDate.SelectedDate != null)
                {
                    dtpFromDate.SelectedDate = dtpFromDate.SelectedDate.Value.AddDays(-1);
                    dtpToDate.SelectedDate = dtpFromDate.SelectedDate;
                }
                else
                {
                    dtpFromDate.SelectedDate = DateTime.Today.AddDays(-1);
                    dtpToDate.SelectedDate = DateTime.Today.AddDays(-1);
                }
            }
            catch (Exception ee)
            {
                MessageBox.Show("오류지점 - btnYesterday_Click : " + ee.ToString());
            }
        }
        // 오늘
        private void btnToday_Click(object sender, RoutedEventArgs e)
        {
            dtpFromDate.Text = DateTime.Now.ToString("yyyy-MM-dd");
            dtpToDate.Text = DateTime.Now.ToString("yyyy-MM-dd");
        }
        // 지난 달
        private void btnLastMonth_Click(object sender, RoutedEventArgs e)
        {
            //string[] receiver = lib.BringLastMonthDatetime();

            //dtpFromDate.Text = receiver[0];
            //dtpToDate.Text = receiver[1];
            try
            {
                if (dtpFromDate.SelectedDate != null)
                {
                    DateTime ThatMonth1 = dtpFromDate.SelectedDate.Value.AddDays(-(dtpFromDate.SelectedDate.Value.Day - 1)); // 선택한 일자 달의 1일!

                    DateTime LastMonth1 = ThatMonth1.AddMonths(-1); // 저번달 1일
                    DateTime LastMonth31 = ThatMonth1.AddDays(-1); // 저번달 말일

                    dtpFromDate.SelectedDate = LastMonth1;
                    dtpToDate.SelectedDate = LastMonth31;
                }
                else
                {
                    DateTime ThisMonth1 = DateTime.Today.AddDays(-(DateTime.Today.Day - 1)); // 이번달 1일

                    DateTime LastMonth1 = ThisMonth1.AddMonths(-1); // 저번달 1일
                    DateTime LastMonth31 = ThisMonth1.AddDays(-1); // 저번달 말일

                    dtpFromDate.SelectedDate = LastMonth1;
                    dtpToDate.SelectedDate = LastMonth31;
                }
            }
            catch (Exception ee)
            {
                MessageBox.Show("오류지점 - btnLastMonth_Click : " + ee.ToString());
            }

        }
        // 이번 달
        private void btnThisMonth_Click(object sender, RoutedEventArgs e)
        {
            string[] receiver = lib.BringThisMonthDatetime();

            dtpFromDate.Text = receiver[0];
            dtpToDate.Text = receiver[1];
        }


        //출고일자(날짜) 체크
        private void chkOutwareDay_Click(object sender, RoutedEventArgs e)
        {
            if (chkOutwareDay.IsChecked == true)
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
        //출고일자(날짜) 체크
        private void chkOutwareDay_Click(object sender, System.Windows.Input.MouseButtonEventArgs e)
        {
            if (chkOutwareDay.IsChecked == true)
            {
                chkOutwareDay.IsChecked = false;
                dtpFromDate.IsEnabled = false;
                dtpToDate.IsEnabled = false;
            }
            else
            {
                chkOutwareDay.IsChecked = true;
                dtpFromDate.IsEnabled = true;
                dtpToDate.IsEnabled = true;
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
            else { cboArticleGroup.IsEnabled = false; }
        }
        //제품그룹
        private void chkArticleGroup_Click(object sender, System.Windows.Input.MouseButtonEventArgs e)
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
        // 품명
        private void chkArticle_Click(object sender, RoutedEventArgs e)
        {
            if (chkArticle.IsChecked == true)
            {
                txtArticle.IsEnabled = true;
                txtArticle.Focus();
                btnPfArticleSrh.IsEnabled = true;
                //btnArticle.IsEnabled = true;
            }
            else
            {
                txtArticle.IsEnabled = false;
                btnPfArticleSrh.IsEnabled = false;
                //btnArticle.IsEnabled = false;
            }
        }
        // 품명
        private void chkArticle_Click(object sender, System.Windows.Input.MouseButtonEventArgs e)
        {
            if (chkArticle.IsChecked == true)
            {
                chkArticle.IsChecked = false;
                txtArticle.IsEnabled = false;
                btnPfArticleSrh.IsEnabled = false;
                //btnArticle.IsEnabled = false;
            }
            else
            {
                chkArticle.IsChecked = true;
                btnPfArticleSrh.IsEnabled = true;
                txtArticle.IsEnabled = true;
                txtArticle.Focus();
                //btnArticle.IsEnabled = true;
            }
        }
        //관리번호 + order NO.
        private void chkOrderID_Click(object sender, RoutedEventArgs e)
        {
            if (chkOrderID.IsChecked == true)
            {
                txtOrderID.IsEnabled = true;
                txtOrderID.Focus();
            }
            else { txtOrderID.IsEnabled = false; }
        }
        //관리번호 + order NO.
        private void chkOrderID_Click(object sender, System.Windows.Input.MouseButtonEventArgs e)
        {
            if (chkOrderID.IsChecked == true)
            {
                chkOrderID.IsChecked = false;
                txtOrderID.IsEnabled = false;
            }
            else
            {
                chkOrderID.IsChecked = true;
                txtOrderID.IsEnabled = true;
                txtOrderID.Focus();
            }
        }
        // 이전창고
        private void chkFromLoc_Click(object sender, RoutedEventArgs e)
        {
            if (chkFromLoc.IsChecked == true)
            {
                cboFromLoc.IsEnabled = true;
                cboFromLoc.Focus();
            }
            else { cboFromLoc.IsEnabled = false; }
        }
        // 이전창고
        private void chkFromLoc_Click(object sender, System.Windows.Input.MouseButtonEventArgs e)
        {
            if (chkFromLoc.IsChecked == true)
            {
                chkFromLoc.IsChecked = false;
                cboFromLoc.IsEnabled = false;
            }
            else
            {
                chkFromLoc.IsChecked = true;
                cboFromLoc.IsEnabled = true;
                cboFromLoc.Focus();
            }
        }
        // 이후창고
        private void chkToLoc_Click(object sender, RoutedEventArgs e)
        {
            if (chkToLoc.IsChecked == true)
            {
                cboToLoc.IsEnabled = true;
                cboToLoc.Focus();
            }
            else { cboToLoc.IsEnabled = false; }
        }
        // 이후창고
        private void chkToLoc_Click(object sender, System.Windows.Input.MouseButtonEventArgs e)
        {
            if (chkToLoc.IsChecked == true)
            {
                chkToLoc.IsChecked = false;
                cboToLoc.IsEnabled = false;
            }
            else
            {
                chkToLoc.IsChecked = true;
                cboToLoc.IsEnabled = true;
                cboToLoc.Focus();
            }
        }
        // 출고구분
        private void chkOutClss_Click(object sender, RoutedEventArgs e)
        {
            if (chkOutClss.IsChecked == true)
            {
                cboOutClss.IsEnabled = true;
                cboOutClss.Focus();
            }
            else { cboOutClss.IsEnabled = false; }
        }
        // 출고구분
        private void chkOutClss_Click(object sender, System.Windows.Input.MouseButtonEventArgs e)
        {
            if (chkOutClss.IsChecked == true)
            {
                chkOutClss.IsChecked = false;
                cboOutClss.IsEnabled = false;
            }
            else
            {
                chkOutClss.IsChecked = true;
                cboOutClss.IsEnabled = true;
                cboOutClss.Focus();
            }
        }

        private void rbnOrderNO_Click(object sender, RoutedEventArgs e)
        {
            txbOrderID.Text = "Order NO";
        }

        private void rbnManageNumber_Click(object sender, RoutedEventArgs e)
        {
            txbOrderID.Text = "관리번호";
        }

        // 콤보박스 목록 불러오기.
        private void ComboBoxSetting()
        {
            cboArticleGroup.Items.Clear();
            cboOutClss.Items.Clear();
            cboFromLoc.Items.Clear();
            cboToLoc.Items.Clear();

            ObservableCollection<CodeView> cbArticleGroup = ComboBoxUtil.Instance.Gf_DB_MT_sArticleGrp();
            ObservableCollection<CodeView> cbOutClss = ComboBoxUtil.Instance.Gf_DB_CM_GetComCodeDataset(null, "OCD", "Y", "", "");
            ObservableCollection<CodeView> cbFromToLoc = ComboBoxUtil.Instance.Gf_DB_CM_GetComCodeDataset(null, "LOC", "Y", "", "");

            this.cboArticleGroup.ItemsSource = cbArticleGroup;
            this.cboArticleGroup.DisplayMemberPath = "code_name";
            this.cboArticleGroup.SelectedValuePath = "code_id";
            this.cboArticleGroup.SelectedIndex = 0;

            this.cboOutClss.ItemsSource = cbOutClss;
            this.cboOutClss.DisplayMemberPath = "code_id_plus_code_name";
            this.cboOutClss.SelectedValuePath = "code_id";
            this.cboOutClss.SelectedIndex = 0;

            this.cboFromLoc.ItemsSource = cbFromToLoc;
            this.cboFromLoc.DisplayMemberPath = "code_name";
            this.cboFromLoc.SelectedValuePath = "code_id";
            this.cboFromLoc.SelectedIndex = 0;

            this.cboToLoc.ItemsSource = cbFromToLoc;
            this.cboToLoc.DisplayMemberPath = "code_name";
            this.cboToLoc.SelectedValuePath = "code_id";
            this.cboToLoc.SelectedIndex = 0;

        }

        #endregion


        #region 플러스 파인더
        //플러스 파인더

        // 품명
        private void btnArticle_Click(object sender, RoutedEventArgs e)
        {
            pf.ReturnCode(txtArticle, 1, "");
        }

        #endregion


        //특수문자 포함 검색
        private string Escape(string str)
        {
            string result = "";

            for (int i = 0; i < str.Length; i++)
            {
                string txt = str.Substring(i, 1);

                bool isSpecial = Regex.IsMatch(txt, @"[^a-zA-Z0-9가-힣]");

                if (isSpecial == true)
                {
                    result += (@"/" + txt);
                }
                else
                {
                    result += txt;
                }
            }
            return result;
        }

        #region 조회 // 조회 프로시저

        // 검색. 조회버튼 클릭.
        private void btnSearch_Click(object sender, RoutedEventArgs e)
        {
            FillGrid();
        }

        private void FillGrid()
        {
            lstOutware.Clear();

            Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
            sqlParameter.Add("ChkDate", chkOutwareDay.IsChecked == true ? 1 : 0);
            sqlParameter.Add("SDate", dtpFromDate.SelectedDate.Value.ToString("yyyyMMdd"));
            sqlParameter.Add("EDate", dtpToDate.SelectedDate.Value.ToString("yyyyMMdd"));
            sqlParameter.Add("ChkCustomID", chkCustomSrh.IsChecked == true ? 1 : 0);
            //sqlParameter.Add("ChkCustomID", 0); //거래처. 본 폼에는 구별할 길 없음..
            sqlParameter.Add("CustomID", "");


            sqlParameter.Add("Custom", chkCustomSrh.IsChecked == true && !txtCustomSrh.Text.Trim().Equals("") ? txtCustomSrh.Text : "");

            //sqlParameter.Add("ChkArticleID", 0);
            sqlParameter.Add("ChkArticleID", chkArticleNoSrh.IsChecked == true ? 1 : 0);
            //sqlParameter.Add("ArticleID", "");
            sqlParameter.Add("ArticleID", txtArticleNo.Tag != null && !txtArticleNo.Text.Trim().Equals("") ? txtArticleNo.Tag.ToString() : "");
            sqlParameter.Add("ChkOrder", chkOrderID.IsChecked == true ? (rbnManageNumber.IsChecked == true ? 1 : 2) : 0);
            sqlParameter.Add("Order", txtOrderID.Text);
            sqlParameter.Add("OrderFlag", 0);       //무쓸모.

            sqlParameter.Add("ArticleGrpID", chkArticleGroup.IsChecked == true && cboArticleGroup.SelectedValue != null ? cboArticleGroup.SelectedValue.ToString() : "");
            sqlParameter.Add("FromLocID", chkFromLoc.IsChecked == true && cboFromLoc.SelectedValue != null ? cboFromLoc.SelectedValue.ToString() : "");
            sqlParameter.Add("ToLocID", chkToLoc.IsChecked == true && cboToLoc.SelectedValue != null ? cboToLoc.SelectedValue.ToString() : "");
            sqlParameter.Add("OutClss", chkOutClss.IsChecked == true && cboOutClss.SelectedValue != null ? cboOutClss.SelectedValue.ToString() : "");
            sqlParameter.Add("sProductYN", "");

            sqlParameter.Add("nBuyerArticleNo", chkArticle.IsChecked == true ? 1 : 0);
            sqlParameter.Add("BuyerArticleNo", chkArticle.IsChecked == true && !txtArticle.Text.Trim().Equals("") ? @Escape(txtArticle.Text) : "");

            DataSet ds = DataStore.Instance.ProcedureToDataSet_LogWrite("xp_Outware_sOutwareDetail", sqlParameter, true, "R");


            if (ds != null && ds.Tables.Count > 0)
            {
                System.Data.DataTable dt = null;
                dt = ds.Tables[0];
                DT = null;
                DT = dt;

                if (dt.Rows.Count == 0)
                {
                    MessageBox.Show("조회결과가 없습니다.");
                    return;
                }
                else
                {
                    dgdOutware.Items.Clear();
                    DataRowCollection drc = dt.Rows;
                    int i = 1;
                    foreach (DataRow item in drc)
                    {
                        if (item["Depth"].ToString() == "0")
                        {
                            var Window_SubulOutwareViewInsert = new Win_mtr_Outware_QView()
                            {
                                NUM = i.ToString(),
                                Depth = item["Depth"].ToString(),

                                OutwareiD = item["OutwareID"].ToString(),
                                OrderID = item["OrderID"].ToString(),
                                OrderNo = item["OrderNo"].ToString(),
                                CustomID = item["CustomID"].ToString(),
                                KCustom = item["KCustom"].ToString(),

                                OutDate = item["OutDate"].ToString().Substring(4, 2) + "/" + item["OutDate"].ToString().Substring(6, 2),
                                Article = item["Article"].ToString(),
                                BuyerArticleNo = item["BuyerArticleNo"].ToString(),

                                //OutCustom = item["OutCustom"].ToString(),
                                OrderQty = item["OrderQty"].ToString(),
                                UnitClss = item["UnitClss"].ToString(),

                                UnitClssName = item["UnitClssName"].ToString(),
                                FromLocName = item["FromLocName"].ToString(),
                                TOLocname = item["TOLocname"].ToString(),
                                OutClssname = item["OutClssname"].ToString(),
                                LabelID = item["LabelID"].ToString(),

                                LabelGubun = item["LabelGubun"].ToString(),
                                OutRoll = item["OutRoll"].ToString(),


                                OutQty = stringFormatN0(ConvertDouble(item["OutQty"].ToString())),

                                //OutQty = item["OutQty"].ToString().Split('.')[0].Trim(),
                                UnitPrice = item["UnitPrice"].ToString(),
                                WorkName = item["WorkName"].ToString(),

                                // N : 정상, S : 샘플, D : 불량
                                NQty = stringFormatN0(item["NQty"]),
                                SQty = stringFormatN0(item["SQty"]),
                                DQty = stringFormatN0(item["DQty"]),

                                VatAmount = item["VatAmount"].ToString(),


                                Amount = stringFormatN0(item["Amount"]), // 금액 → 소수점 버림 + 천 단위

                                TotAmount = item["TotAmount"].ToString(),
                                Remark = item["Remark"].ToString(),

                                ColorGreen = "false",
                                ColorRed = "false"
                            };
                            dgdOutware.Items.Add(Window_SubulOutwareViewInsert);
                            i++;

                            lstOutware.Add(Window_SubulOutwareViewInsert);
                        }
                        if (item["Depth"].ToString() == "2")        // 오더계
                        {
                            var Window_SubulOutwareViewInsert = new Win_mtr_Outware_QView()
                            {
                                NUM = i.ToString(),
                                Depth = item["Depth"].ToString(),

                                OutwareiD = "",
                                OrderID = item["OrderID"].ToString(),
                                OrderNo = "오더 계",
                                CustomID = item["CustomID"].ToString(),
                                KCustom = item["KCustom"].ToString(),

                                OutDate = item["OutDate"].ToString().Substring(4, 2) + "/" + item["OutDate"].ToString().Substring(6, 2),
                                Article = item["Article"].ToString(),
                                BuyerArticleNo = item["BuyerArticleNo"].ToString(),

                                //OutCustom = item["OutCustom"].ToString(),
                                OrderQty = item["OrderQty"].ToString(),
                                UnitClss = item["UnitClss"].ToString(),

                                UnitClssName = item["UnitClssName"].ToString(),
                                FromLocName = "",
                                TOLocname = "",
                                OutClssname = "",
                                LabelID = item["LabelID"].ToString(),

                                LabelGubun = item["LabelGubun"].ToString(),
                                OutRoll = item["OutRoll"].ToString(),
                                OutQty = stringFormatN0(ConvertDouble(item["OutQty"].ToString())),
                                UnitPrice = item["UnitPrice"].ToString(),
                                WorkName = item["WorkName"].ToString(),

                                // N : 정상, S : 샘플, D : 불량
                                NQty = stringFormatN0(item["NQty"]),
                                SQty = stringFormatN0(item["SQty"]),
                                DQty = stringFormatN0(item["DQty"]),

                                VatAmount = item["VatAmount"].ToString(),
                                Amount = stringFormatN0(item["Amount"]),
                                TotAmount = item["TotAmount"].ToString(),
                                Remark = item["Remark"].ToString(),

                                ColorGreen = "false",
                                ColorRed = "false",
                                ColorOrder = "true"
                            };
                            dgdOutware.Items.Add(Window_SubulOutwareViewInsert);
                            i++;
                        }
                        if (item["Depth"].ToString() == "3")            //거래처 계
                        {
                            var Window_SubulOutwareViewInsert = new Win_mtr_Outware_QView()
                            {
                                NUM = i.ToString(),
                                Depth = item["Depth"].ToString(),

                                OutwareiD = "",
                                OrderID = item["OrderID"].ToString(),
                                OrderNo = "",
                                CustomID = item["CustomID"].ToString(),
                                KCustom = "거래처 계",

                                OutDate = item["OutDate"].ToString().Substring(4, 2) + "/" + item["OutDate"].ToString().Substring(6, 2),
                                Article = "",
                                BuyerArticleNo = "",
                                //OutCustom = item["OutCustom"].ToString(),
                                OrderQty = item["OrderQty"].ToString(),
                                UnitClss = item["UnitClss"].ToString(),

                                UnitClssName = item["UnitClssName"].ToString(),
                                FromLocName = "",
                                TOLocname = "",
                                OutClssname = "",
                                LabelID = item["LabelID"].ToString(),

                                LabelGubun = item["LabelGubun"].ToString(),
                                OutRoll = item["OutRoll"].ToString(),
                                OutQty = stringFormatN0(ConvertDouble(item["OutQty"].ToString())),
                                UnitPrice = item["UnitPrice"].ToString(),
                                WorkName = item["WorkName"].ToString(),

                                // N : 정상, S : 샘플, D : 불량
                                NQty = stringFormatN0(item["NQty"]),
                                SQty = stringFormatN0(item["SQty"]),
                                DQty = stringFormatN0(item["DQty"]),

                                VatAmount = item["VatAmount"].ToString(),
                                Amount = stringFormatN0(item["Amount"]),
                                TotAmount = item["TotAmount"].ToString(),
                                Remark = item["Remark"].ToString(),

                                ColorGreen = "true",
                                ColorRed = "false"
                            };
                            dgdOutware.Items.Add(Window_SubulOutwareViewInsert);
                            i++;
                        }
                        if (item["Depth"].ToString() == "4")            //일자 계
                        {
                            var Window_SubulOutwareViewInsert = new Win_mtr_Outware_QView()
                            {
                                NUM = i.ToString(),
                                Depth = item["Depth"].ToString(),

                                OutwareiD = "",
                                OrderID = item["OrderID"].ToString(),
                                OrderNo = "",
                                CustomID = item["CustomID"].ToString(),
                                KCustom = "",

                                OutDate = "일 계",
                                Article = "",
                                BuyerArticleNo = "",

                                //OutCustom = item["OutCustom"].ToString(),
                                OrderQty = item["OrderQty"].ToString(),
                                UnitClss = item["UnitClss"].ToString(),

                                UnitClssName = item["UnitClssName"].ToString(),
                                FromLocName = "",
                                TOLocname = "",
                                OutClssname = "",
                                LabelID = item["LabelID"].ToString(),

                                LabelGubun = item["LabelGubun"].ToString(),
                                OutRoll = item["OutRoll"].ToString(),
                                OutQty = stringFormatN0(ConvertDouble(item["OutQty"].ToString())),
                                UnitPrice = item["UnitPrice"].ToString(),
                                WorkName = item["WorkName"].ToString(),

                                // N : 정상, S : 샘플, D : 불량
                                NQty = stringFormatN0(item["NQty"]),
                                SQty = stringFormatN0(item["SQty"]),
                                DQty = stringFormatN0(item["DQty"]),

                                VatAmount = item["VatAmount"].ToString(),
                                Amount = stringFormatN0(item["Amount"]),
                                TotAmount = item["TotAmount"].ToString(),
                                Remark = item["Remark"].ToString(),

                                ColorGreen = "true",
                                ColorRed = "false"
                            };
                            dgdOutware.Items.Add(Window_SubulOutwareViewInsert);
                            i++;
                        }
                        if (item["Depth"].ToString() == "5")            //월 계
                        {
                            var Window_SubulOutwareViewInsert = new Win_mtr_Outware_QView()
                            {
                                NUM = i.ToString(),
                                Depth = item["Depth"].ToString(),

                                OutwareiD = "",
                                OrderID = item["OrderID"].ToString(),
                                OrderNo = "",
                                CustomID = item["CustomID"].ToString(),
                                KCustom = "",

                                OutDate = item["OutDate"].ToString().Substring(4, 2) + "월 계",
                                Article = "",
                                BuyerArticleNo = "",

                                //OutCustom = item["OutCustom"].ToString(),
                                OrderQty = item["OrderQty"].ToString(),
                                UnitClss = item["UnitClss"].ToString(),

                                UnitClssName = item["UnitClssName"].ToString(),
                                FromLocName = "",
                                TOLocname = "",
                                OutClssname = "",
                                LabelID = item["LabelID"].ToString(),

                                LabelGubun = item["LabelGubun"].ToString(),
                                OutRoll = item["OutRoll"].ToString(),
                                OutQty = stringFormatN0(ConvertDouble(item["OutQty"].ToString())),
                                UnitPrice = item["UnitPrice"].ToString(),
                                WorkName = item["WorkName"].ToString(),

                                // N : 정상, S : 샘플, D : 불량
                                NQty = stringFormatN0(item["NQty"]),
                                SQty = stringFormatN0(item["SQty"]),
                                DQty = stringFormatN0(item["DQty"]),

                                VatAmount = item["VatAmount"].ToString(),
                                Amount = stringFormatN0(item["Amount"]),
                                TotAmount = item["TotAmount"].ToString(),
                                Remark = item["Remark"].ToString(),

                                ColorGreen = "false",
                                ColorRed = "true"
                            };
                            dgdOutware.Items.Add(Window_SubulOutwareViewInsert);
                            i++;
                        }
                        if (item["Depth"].ToString() == "6")            // 총 합 계
                        {
                            var Window_SubulOutwareViewInsert = new Win_mtr_Outware_QView()
                            {
                                NUM = i.ToString(),
                                Depth = item["Depth"].ToString(),

                                OutwareiD = "",
                                OrderID = item["OrderID"].ToString(),
                                OrderNo = "",
                                CustomID = item["CustomID"].ToString(),
                                KCustom = "",

                                OutDate = "총 합계",
                                Article = "",
                                BuyerArticleNo = "",

                                //OutCustom = item["OutCustom"].ToString(),
                                OrderQty = item["OrderQty"].ToString(),
                                UnitClss = item["UnitClss"].ToString(),

                                UnitClssName = item["UnitClssName"].ToString(),
                                FromLocName = "",
                                TOLocname = "",
                                OutClssname = "",
                                LabelID = item["LabelID"].ToString(),

                                LabelGubun = item["LabelGubun"].ToString(),
                                OutRoll = item["OutRoll"].ToString(),
                                OutQty = stringFormatN0(ConvertDouble(item["OutQty"].ToString())),
                                UnitPrice = item["UnitPrice"].ToString(),
                                WorkName = item["WorkName"].ToString(),

                                // N : 정상, S : 샘플, D : 불량
                                NQty = stringFormatN0(item["NQty"]),
                                SQty = stringFormatN0(item["SQty"]),
                                DQty = stringFormatN0(item["DQty"]),

                                VatAmount = item["VatAmount"].ToString(),
                                Amount = stringFormatN0(item["Amount"]),
                                TotAmount = item["TotAmount"].ToString(),
                                Remark = item["Remark"].ToString(),

                                ColorGreen = "false",
                                ColorRed = "true"
                            };
                            dgdOutware.Items.Add(Window_SubulOutwareViewInsert);
                        }
                    }

                }

            }
        }

        #endregion


        #region 엑셀

        // 엑셀버튼 클릭.
        private void btnExcel_Click(object sender, RoutedEventArgs e)
        {
            if (dgdOutware.Items.Count < 1)
            {
                MessageBox.Show("먼저 검색해 주세요.");
                return;
            }

            Lib lib = new Lib();
            System.Data.DataTable dt = null;
            string Name = string.Empty;

            string[] lst = new string[4];
            lst[0] = "메인 그리드";
            lst[2] = dgdOutware.Name;

            ExportExcelxaml ExpExc = new ExportExcelxaml(lst);

            ExpExc.ShowDialog();

            if (ExpExc.DialogResult.HasValue)
            {
                if (ExpExc.choice.Equals(dgdOutware.Name))
                {
                    DataStore.Instance.InsertLogByForm(this.GetType().Name, "E");
                    //MessageBox.Show("대분류");
                    if (ExpExc.Check.Equals("Y"))
                        dt = lib.DataGridToDTinHidden(dgdOutware);
                    else
                        dt = lib.DataGirdToDataTable(dgdOutware);

                    Name = dgdOutware.Name;
                    lib.GenerateExcel(dt, Name);
                    lib.excel.Visible = true;
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

        #endregion


        #region 프린트 인쇄

        // 인쇄 기능.
        private void btnPrint_Click(object sender, RoutedEventArgs e)
        {
            ContextMenu menu = btnPrint.ContextMenu;
            menu.StaysOpen = true;
            menu.IsOpen = true;
        }

        // 인쇄 서브메뉴1. 미리보기
        private void menuSeeAhead_Click(object sender, RoutedEventArgs e)
        {
            if (dgdOutware.Items.Count < 1)
            {
                MessageBox.Show("먼저 검색해 주세요.");
                return;
            }

            msg.Show();
            msg.Topmost = true;
            msg.Refresh();

            PrintWork(true);
        }

        // 인쇄 서브메뉴2. 바로인쇄
        private void menuRighPrint_Click(object sender, RoutedEventArgs e)
        {
            if (dgdOutware.Items.Count < 1)
            {
                MessageBox.Show("먼저 검색해 주세요.");
                return;
            }

            DataStore.Instance.InsertLogByForm(this.GetType().Name, "P");
            msg.Show();
            msg.Topmost = true;
            msg.Refresh();

            PrintWork(false);
        }
        //인쇄 서브메뉴3. 그냥 닫기
        private void menuClose_Click(object sender, RoutedEventArgs e)
        {
            ContextMenu menu = btnPrint.ContextMenu;
            menu.StaysOpen = false;
            menu.IsOpen = false;
        }

        private void PrintWork(bool previewYN)
        {
            excelapp = new Microsoft.Office.Interop.Excel.Application();

            string MyBookPath = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetEntryAssembly().Location) + "\\Report\\자재출고명세서.xls";
            workbook = excelapp.Workbooks.Add(MyBookPath);
            worksheet = workbook.Sheets["Form"];
            pastesheet = workbook.Sheets["Print"];

            // 이전 창고로 검색했다면 이전창고 입력 : D3
            workrange = worksheet.get_Range("D3");
            workrange.Value2 = chkFromLoc.IsChecked == true && cboFromLoc.SelectedValue != null ? cboFromLoc.SelectedValue.ToString() : "";

            // 기간 설정후에 검색했다면 검색일자 2019.10.01 ~ 2019.10.31 입력 : D4
            workrange = worksheet.get_Range("D4");
            workrange.Value2 = chkOutwareDay.IsChecked == true ? dtpFromDate.SelectedDate.Value.ToString("yyyy-MM-dd") + " ~ " + dtpToDate.SelectedDate.Value.ToString("yyyy-MM-dd") : "";

            // 일자 입력(오늘일자) : AE4
            workrange = worksheet.get_Range("AE4");
            workrange.Value2 = DateTime.Today.ToString("yyyy-MM-dd");

            // 페이지 계산 등
            int rowCount = lstOutware.Count;
            int excelStartRow = 7;

            // 총 데이터를 입력할수 있는 갯수
            int totalDataInput = 37;

            // 카피할 다음페이지 인덱스
            int nextCopyLine = 46;

            int copyLine = 0;
            int Page = 0;
            int PageAll = (int)Math.Ceiling(1.0 * rowCount / totalDataInput);
            int DataCount = 0;

            for (int k = 0; k < PageAll; k++)
            {
                Page++;
                copyLine = ((Page - 1) * (nextCopyLine - 1));

                int excelNum = 0;

                // 기존에 있는 데이터 지우기 "A7", "AG43"
                worksheet.Range["A7", "AG43"].EntireRow.ClearContents();

                for (int i = DataCount; i < rowCount; i++)
                {
                    if (i == totalDataInput * Page)
                    {
                        break;
                    }

                    var Outware = lstOutware[i];
                    int excelRow = excelStartRow + excelNum;

                    if (Outware != null)
                    {
                        //일자
                        workrange = worksheet.get_Range("A" + excelRow);
                        workrange.Value2 = Outware.OutDate;

                        // 출고처명
                        workrange = worksheet.get_Range("C" + excelRow);
                        workrange.Value2 = Outware.KCustom;

                        // 품명
                        workrange = worksheet.get_Range("G" + excelRow);
                        workrange.Value2 = Outware.BuyerArticleNo;

                        // 이전창고
                        workrange = worksheet.get_Range("K" + excelRow);
                        workrange.Value2 = Outware.FromLocName;

                        // 이후창고
                        workrange = worksheet.get_Range("O" + excelRow);
                        workrange.Value2 = Outware.TOLocname;

                        // 정상수량
                        workrange = worksheet.get_Range("S" + excelRow);
                        workrange.Value2 = Outware.NQty;

                        // 샘플수량
                        workrange = worksheet.get_Range("W" + excelRow);
                        workrange.Value2 = Outware.SQty;

                        // 불량수량
                        workrange = worksheet.get_Range("Z" + excelRow);
                        workrange.Value2 = Outware.DQty;

                        // 출고구분
                        workrange = worksheet.get_Range("AC" + excelRow);
                        workrange.Value2 = Outware.OutClssname;

                        // 비고
                        workrange = worksheet.get_Range("AG" + excelRow);
                        workrange.Value2 = Outware.Remark;

                        //SumAmount += ConvertDouble(OcReqSub.Amount);

                        excelNum++;
                        DataCount = i;
                    }
                }

                // 2장 이상 넘어가면 페이지 넘버 입력
                if (PageAll > 1)
                {
                    pastesheet.PageSetup.CenterFooter = "&P / &N";
                }

                // Form 시트 내용 Print 시트에 복사 붙여넣기
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

        #endregion



        //닫기 기능.
        private void btnClose_Click(object sender, RoutedEventArgs e)
        {
            DataStore.Instance.InsertLogByFormS(this.GetType().Name, stDate, stTime, "E");
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

        //재검색
        private void re_Search(int selectedIndex)
        {
            FillGrid();

            if (dgdOutware.Items.Count > 0)
            {
                dgdOutware.SelectedIndex = selectedIndex;
            }
        }


        //정렬 버튼이벤트.
        private void btnMultiSort_Click(object sender, RoutedEventArgs e)
        {
            PopUp.MultiLevelSort MLS = new PopUp.MultiLevelSort(dgdOutware);
            MLS.ShowDialog();

            if (MLS.DialogResult.HasValue)
            {
                string targetSortProperty = string.Empty;
                int targetColIndex;
                dgdOutware.Items.SortDescriptions.Clear();

                for (int x = 0; x < MLS.ColName.Count; x++)
                {
                    targetSortProperty = MLS.SortingProperty[x];
                    targetColIndex = MLS.ColIndex[x];
                    var targetCol = dgdOutware.Columns[targetColIndex];

                    if (targetSortProperty == "UP")
                    {
                        dgdOutware.Items.SortDescriptions.Add(new SortDescription(targetCol.SortMemberPath, ListSortDirection.Ascending));
                        targetCol.SortDirection = ListSortDirection.Ascending;
                    }
                    else
                    {
                        dgdOutware.Items.SortDescriptions.Add(new SortDescription(targetCol.SortMemberPath, ListSortDirection.Descending));
                        targetCol.SortDirection = ListSortDirection.Descending;
                    }
                }
                dgdOutware.Refresh();
            }
        }

        private void txtArticle_KeyDown(object sender, System.Windows.Input.KeyEventArgs e)
        {
            //if (e.Key == Key.Enter)
            //{
            //    rowNum = 0;
            //    re_Search(rowNum);
            //}
            if (e.Key == Key.Enter)
            {
                e.Handled = true;
                MainWindow.pf.ReturnCode(txtArticle, (int)Defind_CodeFind.DCF_BuyerArticleNo, "");
            }
        }

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

        private void btnPfArticleSrh_Click(object sender, RoutedEventArgs e)
        {
            MainWindow.pf.ReturnCode(txtArticle, (int)Defind_CodeFind.DCF_BuyerArticleNo, "");
        }

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

        private void txtCustomSrh_KeyDown(object sender, KeyEventArgs e)
        {
            try
            {
                if (e.Key == Key.Enter)
                {
                    re_Search(0);
                }
            }
            catch (Exception ee)
            {
                MessageBox.Show("오류지점 - 거래처 검색 엔터키 : " + ee.ToString());
            }
        }

        //품명 라벨체크
        private void lblArticleNo_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            if (chkArticleNoSrh.IsChecked == true)
            {
                chkArticleNoSrh.IsChecked = false;
            }
            else
            {
                chkArticleNoSrh.IsChecked = true;
            }
        }

        //품명 체크 
        private void chkArticleNo_Checked(object sender, RoutedEventArgs e)
        {
            chkArticleNoSrh.IsChecked = true;

            txtArticleNo.IsEnabled = true;

            btnPfArticleNoSrh.IsEnabled = true;
        }
        //품명 안체크
        private void chkArticleNo_Unchecked(object sender, RoutedEventArgs e)
        {
            chkArticleNoSrh.IsChecked = false;

            txtArticleNo.IsEnabled = false;

            btnPfArticleNoSrh.IsEnabled = false;
        }
        //품명 키다운 
        private void txtArticleNo_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                e.Handled = true;
                MainWindow.pf.ReturnCode(txtArticleNo, 77, "");
            }
        }

        //품명 플러스파인더
        private void btnPfArticleNoSrh_Click(object sender, RoutedEventArgs e)
        {
            MainWindow.pf.ReturnCode(txtArticleNo, 77, "");
        }

    }

    /////////////////////////////////////////////////////////////////////
    /// Win_mtr_Outware_QView
    class Win_mtr_Outware_QView : BaseView
    {
        public override string ToString()
        {
            return (this.ReportAllProperties());
        }

        public ObservableCollection<CodeView> cboTrade { get; set; }

        // 조회 값.    
        public string OutwareiD { get; set; }
        public string OrderID { get; set; }
        public string OrderNo { get; set; }
        public string CustomID { get; set; }
        public string KCustom { get; set; }

        public string OutDate { get; set; }
        public string Article { get; set; }
        public string BuyerArticleNo { get; set; }
        public string ArticleID { get; set; }
        public string Sabun { get; set; }


        public string OutCustom { get; set; }
        public string OrderQty { get; set; }
        public string UnitClss { get; set; }

        public string UnitClssName { get; set; }
        public string FromLocName { get; set; }
        public string TOLocname { get; set; }
        public string OutClssname { get; set; }
        public string LabelID { get; set; }

        public string LabelGubun { get; set; }
        public string OutRoll { get; set; }
        public string OutQty { get; set; }
        public string UnitPrice { get; set; }
        public string WorkName { get; set; }

        // N : 정상, S : 샘플, D : 불량
        public string NQty { get; set; }
        public string SQty { get; set; }
        public string DQty { get; set; }

        public string VatAmount { get; set; }
        public string Amount { get; set; }
        public string TotAmount { get; set; }
        public string Remark { get; set; }
        public string Depth { get; set; }


        //순번 체크용/
        public string NUM { get; set; }

        //컬러 칠하기
        public string ColorGreen { get; set; }
        public string ColorRed { get; set; }
        public string ColorOrder { get; set; }

    }



}

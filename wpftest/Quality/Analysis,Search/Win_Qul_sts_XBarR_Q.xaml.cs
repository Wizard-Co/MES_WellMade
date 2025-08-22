using LiveCharts;
using LiveCharts.Wpf;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Data;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Threading;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Data;
using System.Windows.Input;
using System.Windows.Media.Imaging;
using WizMes_WellMade.PopUP;
namespace WizMes_WellMade
{
    /// <summary>
    /// Win_Qul_sts_XBarR_Q.xaml에 대한 상호 작용 논리
    /// </summary>
    public partial class Win_Qul_sts_XBarR_Q : UserControl
    {
        Lib lib = new Lib();
        PlusFinder pf = new PlusFinder();

        string RASpecMax = string.Empty;
        string RASpecMin = string.Empty;

        // X-BAR 차트용 전역변수.
        double X_chart_UCL = 0;
        double X_chart_CL = 0;
        double X_chart_LCL = 0;
        List<double> X_chart_Val = new List<double>();

        // R 차트용 전역변수.
        double R_chart_UCL = 0;
        double R_chart_CL = 0;
        List<double> R_chart_Val = new List<double>();

        //---------------------------------
        // Chart Header.
        List<string> Chart_Header = new List<string>();

        //Image 변수 선언
        System.Windows.Controls.Image ImageData = new System.Windows.Controls.Image();

        double LValue1 = 0;
        double LValue2 = 0;
        double LValue3 = 0;
        double LValue4 = 0;
        double LValue5 = 0;

        double LValue6 = 0;
        double LValue7 = 0;
        double LValue8 = 0;
        double LValue9 = 0;
        double LValue10 = 0;

        double LValue11 = 0;
        double LValue12 = 0;
        double LValue13 = 0;
        double LValue14 = 0;
        double LValue15 = 0;

        double LValue16 = 0;
        double LValue17 = 0;
        double LValue18 = 0;
        double LValue19 = 0;
        double LValue20 = 0;

        double LValue21 = 0;
        double LValue22 = 0;
        double LValue23 = 0;
        double LValue24 = 0;
        double LValue25 = 0;

        double LValue26 = 0;
        double LValue27 = 0;
        double LValue28 = 0;
        double LValue29 = 0;
        double LValue30 = 0;

        double LValue31 = 0;
        double LValue32 = 0;
        double LValue33 = 0;
        double LValue34 = 0;
        double LValue35 = 0;

        double LValue36 = 0;
        double LValue37 = 0;
        double LValue38 = 0;
        double LValue39 = 0;
        double LValue40 = 0;

        private Microsoft.Office.Interop.Excel.Application excelapp;
        private Microsoft.Office.Interop.Excel.Workbook workbook;
        private Microsoft.Office.Interop.Excel.Worksheet worksheet;
        private Microsoft.Office.Interop.Excel.Range workrange;
        private Microsoft.Office.Interop.Excel.Worksheet copysheet;
        private Microsoft.Office.Interop.Excel.Worksheet pastesheet;
        //---------------------------------



        SeriesCollection SeriesCollection;
        public string[] X_Linelbl { get; set; }


        public Win_Qul_sts_XBarR_Q()
        {
            InitializeComponent();
        }


        // 첫 로드시.
        private void Win_Qul_sts_XBarR_Q_Loaded(object sender, RoutedEventArgs e)
        {
            First_Step();
            ComboxSetting();
            Create_Row_Header();
        }


        #region 첫 단계 / 날짜버튼 세팅 / 조회용 체크박스 세팅 
        // 첫 단계
        private void First_Step()
        {
            chkMonthDate.IsChecked = true;
            dtpFromDate.Text = DateTime.Now.ToString("yyyy-MM-dd");
            dtpToDate.Text = DateTime.Now.ToString("yyyy-MM-dd");

            //dtpToDate.Visibility = Visibility.Hidden;           //날짜는 하나만 있으면 됨.

            tbnJaju.IsChecked = true;

            dtpFromDate.IsEnabled = true;
            dtpToDate.IsEnabled = true;

            txtCustomer.IsEnabled = false;
            btnCustomer.IsEnabled = false;
            txtArticle.IsEnabled = false;
            btnArticle.IsEnabled = false;


        }

        // 검사일자
        private void chkMonthDate_Click(object sender, RoutedEventArgs e)
        {
            if (chkMonthDate.IsChecked == true)
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
        //검사일자
        private void chkMonthDate_Click(object sender, MouseButtonEventArgs e)
        {
            if (chkMonthDate.IsChecked == true)
            {
                chkMonthDate.IsChecked = false;
                dtpFromDate.IsEnabled = false;
                dtpToDate.IsEnabled = false;
            }
            else
            {
                chkMonthDate.IsChecked = true;
                dtpFromDate.IsEnabled = true;
                dtpToDate.IsEnabled = true;
            }
        }
        // 품명
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
        // 품명
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
        // 거래처
        private void chkCustomer_Click(object sender, RoutedEventArgs e)
        {
            if (chkCustomer.IsChecked == true)
            {
                txtCustomer.IsEnabled = true;
                txtCustomer.Focus();
                btnCustomer.IsEnabled = true;
            }
            else
            {
                txtCustomer.IsEnabled = false;
                btnCustomer.IsEnabled = false;
            }
        }
        // 거래처
        private void chkCustomer_Click(object sender, MouseButtonEventArgs e)
        {
            if (chkCustomer.IsChecked == true)
            {
                chkCustomer.IsChecked = false;
                txtCustomer.IsEnabled = false;
                btnCustomer.IsEnabled = false;
            }
            else
            {
                chkCustomer.IsChecked = true;
                txtCustomer.IsEnabled = true;
                txtCustomer.Focus();
                btnCustomer.IsEnabled = true;
            }
        }

        #endregion

        #region Content 통합 검색용 콤보박스
        private void ComboxSetting()
        {
            // 통합 검색 용 콤보박스 셋팅
            List<string[]> TotSrh = new List<string[]>();
            string[] TotSrh1 = new string[] { "1", "사번" };
            string[] TotSrh2 = new string[] { "2", "품명" };
            string[] TotSrh3 = new string[] { "3", "검사항목" };
            //string[] TotSrh4 = new string[] { "4", "재질" };
            //string[] TotSrh4 = new string[] { "4", "고객사" };
            //string[] TotSrh6 = new string[] { "6", "차종" };

            TotSrh.Add(TotSrh1);
            TotSrh.Add(TotSrh2);
            TotSrh.Add(TotSrh3);
            //TotSrh.Add(TotSrh4);
            //TotSrh.Add(TotSrh5);
            //TotSrh.Add(TotSrh6);

            ObservableCollection<CodeView> ovcTotSrh = ComboBoxUtil.Instance.Direct_SetComboBox(TotSrh);
            this.cboTotSearch.ItemsSource = ovcTotSrh;
            this.cboTotSearch.DisplayMemberPath = "code_name";
            this.cboTotSearch.SelectedValuePath = "code_id";
        }
        #endregion


        private void Create_Row_Header()
        {
            string[] RowHeaderName = new string[1];
            RowHeaderName[0] = "측정값";

            //for (int i = 0; i < 1; i++)
            //{
            //    var Win_Qul_sts_XBarR_Q_Insert = new Win_Qul_sts_XBarR_Q_View()
            //    {
            //        Spread_RowHeaderColumns = RowHeaderName[i]
            //    };
            //    dgdXBar_DailySpread.Items.Add(Win_Qul_sts_XBarR_Q_Insert);
            //}
        }

        #region 플러스 파인더
        // 플러스파인더 _ 고객사 찾기.
        private void btnCustomer_Click(object sender, RoutedEventArgs e)
        {
            pf.ReturnCode(txtCustomer, 0, "");
        }
        // 플러스파인더 _ 품명 찾기.
        private void btnArticle_Click(object sender, RoutedEventArgs e)
        {
            pf.ReturnCode(txtArticle, 1, "");
        }

        #endregion

        private void Allclear()
        {
            dgdXBar_std.Items.Clear();
            dgdXBar_DailySpread.Items.Clear();
            txtCP.Clear();
            txtCPK.Clear();
            txtCPL.Clear();
            txtCPU.Clear();
            txtMaxValue.Clear();
            txtMinValue.Clear();
            txtRCL.Clear();
            txtRUCL.Clear();
            txtXbarCL.Clear();
            txtXbarLCL.Clear();
            txtXbarUCL.Clear();
            txtAverage.Clear();
            txtStandardDeviation.Clear();

            if (lvcXBarChart.Series != null && lvcRChart.Series != null)
            {
                lvcXBarChart.Series.Clear();
                lvcRChart.Series.Clear();
            }

        }

        #region 조회 / 조회용 프로시저
        // 검색버튼 클릭.
        private void btnSearch_Click(object sender, RoutedEventArgs e)
        {
            //검색버튼 비활성화
            btnSearch.IsEnabled = false;

            Dispatcher.BeginInvoke(new Action(() =>

            {
                Thread.Sleep(2000);

                //로직
                Allclear();
                FillGrid();

            }), System.Windows.Threading.DispatcherPriority.Background);



            Dispatcher.BeginInvoke(new Action(() =>

            {
                btnSearch.IsEnabled = true;

            }), System.Windows.Threading.DispatcherPriority.Background);


        }

        private void FillGrid()
        {
            int nchkDate = 0;
            if (chkMonthDate.IsChecked == true) { nchkDate = 1; }

            string SDate = string.Empty;
            string EDate = string.Empty;
            SDate = dtpFromDate.SelectedDate.Value.ToString("yyyyMMdd");
            EDate = dtpToDate.SelectedDate.Value.ToString("yyyyMMdd");

            string InspectPoint = string.Empty;
            if (tbnInCome.IsChecked == true) { InspectPoint = "1"; }
            else if (tbnProcessCycle.IsChecked == true) { InspectPoint = "3"; }
            else if (tbnOutCome.IsChecked == true) { InspectPoint = "5"; }
            else if (tbnJaju.IsChecked == true) { InspectPoint = "9"; } //자주


            int ChkCustomID = 0;                //거래처.
            string CustomID = string.Empty;
            string Custom = string.Empty;

            int ChkArticleID = 0;              //품명.
            string ArticleID = string.Empty;
            string Article = string.Empty;


            if (chkCustomer.IsChecked == true)
            {
                if (txtCustomer.Tag == null)
                {
                    txtCustomer.Tag = "";
                    if (txtCustomer.Text.Length > 0)
                    {
                        ChkCustomID = 2;
                        Custom = txtCustomer.Text;
                    }
                }
                else
                {
                    ChkCustomID = 1;
                    CustomID = txtCustomer.Tag.ToString();
                }
            }


            if (chkArticle.IsChecked == true)
            {
                if (txtArticle.Tag == null)
                {
                    txtArticle.Tag = "";
                    if (txtArticle.Text.Length > 0)
                    {
                        ChkArticleID = 2;
                        Article = txtArticle.Text;
                    }
                }
                else
                {
                    ChkArticleID = 1;
                    ArticleID = txtArticle.Tag.ToString();
                }
            }


            Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
            sqlParameter.Clear();
            sqlParameter.Add("nchkDate", nchkDate);               //int         
            sqlParameter.Add("FromDate", SDate);
            sqlParameter.Add("ToDate", EDate);

            sqlParameter.Add("InspectPoint", InspectPoint);

            sqlParameter.Add("nchkCustom", ChkCustomID);             //int
            sqlParameter.Add("CustomID", CustomID);
            sqlParameter.Add("Custom", Custom);

            sqlParameter.Add("nchkArticleID", ChkArticleID);          //int
            sqlParameter.Add("ArticleID", ArticleID);
            sqlParameter.Add("Article", Article);

            // 통합 검색용 파라미터
            sqlParameter.Add("ntotSearch", chkTotSearch.IsChecked == true &&
                                          cboTotSearch.SelectedValue != null &&
                                           txtTotSearch.Text.Trim().Equals("") == false ? 1 : 0);
            sqlParameter.Add("ntotSearchGbn", cboTotSearch.SelectedValue != null ?
                                              cboTotSearch.SelectedValue.ToString() : "");
            sqlParameter.Add("stotSearch", txtTotSearch.Text);

            DataSet ds = DataStore.Instance.ProcedureToDataSet("xp_Qual_sSpc_std_20220124", sqlParameter, false);

            if (ds != null && ds.Tables.Count > 0)
            {
                DataTable dt = null;
                dt = ds.Tables[0];

                if (dt.Rows.Count == 0)
                {
                    MessageBox.Show("SPC 항목 조회결과가 없습니다.");
                    return;
                }
                else
                {
                    dgdXBar_std.Items.Clear();
                    int i = 1;
                    DataRowCollection drc = dt.Rows;
                    foreach (DataRow item in drc)
                    {
                        var Win_Qul_sts_XBarR_Q_Insert = new Win_Qul_sts_XBarR_Q_View()
                        {
                            STD_NUM = i.ToString(),

                            STD_Article = item["Article"].ToString(),
                            STD_ArticleID = item["ArticleID"].ToString(),
                            STD_Sabun = item["Article_Sabun"].ToString(),
                            STD_EcoNo = item["EcoNo"].ToString(),
                            STD_InspectBasisID = item["InspectBasisID"].ToString(),
                            STD_insItemName = item["insItemName"].ToString(),

                            STD_SubSeq = item["SubSeq"].ToString()
                        };
                        dgdXBar_std.Items.Add(Win_Qul_sts_XBarR_Q_Insert);
                        i++;
                    }
                }
            }
        }

        #endregion

        #region 닫기
        // 닫기.
        private void btnClose_Click(object sender, RoutedEventArgs e)
        {
            Lib.Instance.ChildMenuClose(this.ToString());
        }

        #endregion

        #region 엑셀
        // 엑셀
        private void btnExcel_Click(object sender, RoutedEventArgs e)
        {
            if (dgdXBar_std.Items.Count < 1)
            {
                MessageBox.Show("먼저 검색해 주세요.");
                return;
            }

            DataTable dt = null;
            string Name = string.Empty;
            Lib lib = new Lib();

            string[] lst = new string[4];
            lst[0] = "메인그리드";
            lst[1] = "통계치그리드";
            lst[2] = dgdXBar_std.Name;
            lst[3] = dgdXBar_DailySpread.Name;

            ExportExcelxaml ExpExc = new ExportExcelxaml(lst);

            ExpExc.ShowDialog();

            if (ExpExc.DialogResult.HasValue)
            {
                if (ExpExc.choice.Equals(dgdXBar_std.Name))
                {
                    //MessageBox.Show("대분류");
                    if (ExpExc.Check.Equals("Y"))
                        dt = lib.DataGridToDTinHidden(dgdXBar_std);
                    else
                        dt = lib.DataGirdToDataTable(dgdXBar_std);

                    Name = dgdXBar_std.Name;

                    if (lib.GenerateExcel(dt, Name))
                    {
                        lib.excel.Visible = true;
                        lib.ReleaseExcelObject(lib.excel);
                    }
                }
                else if (ExpExc.choice.Equals(dgdXBar_DailySpread.Name))
                {
                    //MessageBox.Show("정성류");
                    if (ExpExc.Check.Equals("Y"))
                        dt = lib.DataGridToDTinHidden(dgdXBar_DailySpread);
                    else
                        dt = lib.DataGirdToDataTable(dgdXBar_DailySpread);
                    Name = dgdXBar_DailySpread.Name;

                    if (lib.GenerateExcel(dt, Name))
                    {
                        lib.excel.Visible = true;
                        lib.ReleaseExcelObject(lib.excel);
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


        #endregion

        #region 캡처버튼 클릭 이벤트
        private void BtnCapture_Click(object sender, RoutedEventArgs e)
        {
            ScreenCapture();

            if (!ImgImage.Source.Equals(null))
            {
                //전역변수 ImageData 소스에 원본 ImgImage 소스를 대입
                ImageData.Source = ImgImage.Source;

                //MainWindow에 imgage리스트에 담아서 ScreenChot페이지로 넘겨준다.
                MainWindow.ScreenCapture.Clear();
                MainWindow.ScreenCapture.Add(ImageData);

            }

            PopUp.ScreenShot SCshot = new PopUp.ScreenShot();

            //보여줘
            SCshot.ShowDialog();
        }

        public void ScreenCapture()
        {
            //화면의 크기 정보 
            int width = (int)SystemParameters.PrimaryScreenWidth + 70;
            int height = (int)SystemParameters.PrimaryScreenHeight;

            //화면의 크기만큼 bitmap생성
            using (Bitmap bmp = new Bitmap(width, height, System.Drawing.Imaging.PixelFormat.Format32bppArgb))
            {
                //bitmap 이미지 변경을 위해 Grapics 객체 생성
                using (Graphics gr = Graphics.FromImage(bmp))
                {
                    // 화면을 그대로 카피해서 Bitmap 메모리에 저장 
                    gr.CopyFromScreen(280, 130, 0, 0, bmp.Size);
                }

                //Bitmap 데이터를 파일로(저장 경로를 지정해서??)
                bmp.Save(@"c:\temp\" + DateTime.Now.ToString("yyyy-MM-dd,HHmmss") + ".png", ImageFormat.Png);

                using (MemoryStream memory = new MemoryStream())
                {
                    bmp.Save(memory, ImageFormat.Bmp);
                    memory.Position = 0;
                    BitmapImage bitmapImage = new BitmapImage();
                    bitmapImage.BeginInit();
                    bitmapImage.StreamSource = memory;
                    bitmapImage.CacheOption = BitmapCacheOption.OnLoad;
                    bitmapImage.EndInit();

                    ImgImage.Source = bitmapImage;

                }
            }
        }
        #endregion

        #region 메인그리드 row enter. Show Data
        // 좌측 std 그리드 클릭시. 로우엔터 이벤트.
        private void dgdXBar_std_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            var ViewReceiver = dgdXBar_std.SelectedItem as Win_Qul_sts_XBarR_Q_View;

            if (ViewReceiver != null)
            {
                int nchkDate = 0;
                if (chkMonthDate.IsChecked == true) { nchkDate = 1; }

                string SDate = string.Empty;
                string EDate = string.Empty;
                SDate = dtpFromDate.SelectedDate.Value.ToString("yyyyMMdd");
                EDate = dtpToDate.SelectedDate.Value.ToString("yyyyMMdd");

                string InspectPoint = string.Empty;
                if (tbnInCome.IsChecked == true) { InspectPoint = "1"; }
                else if (tbnProcessCycle.IsChecked == true) { InspectPoint = "3"; }
                else if (tbnOutCome.IsChecked == true) { InspectPoint = "5"; }
                else if (tbnJaju.IsChecked == true) { InspectPoint = "9"; } //자주


                string InspectBasisID = ViewReceiver.STD_InspectBasisID;
                int Seq = Convert.ToInt32(ViewReceiver.STD_SubSeq);


                // Specification 공간 채울 데이터 구하기.
                FillGrid_Specification(nchkDate, SDate, EDate, InspectBasisID, Seq);
                //(최대값 / 최소값 구하기 용도...)


                //통계치 Summary 공간 채울 데이터 구하기.


                CreateDataGridColumns(InspectPoint, nchkDate, SDate, EDate, InspectBasisID, Seq);
                FillGrid_Summary(InspectPoint, nchkDate, SDate, EDate, InspectBasisID, Seq);


                // 차트 그리기.
                FillChart_Double();


                // 전역 값들 초기화  클리어작업.
                RASpecMax = string.Empty;
                RASpecMin = string.Empty;

                X_chart_UCL = 0;
                X_chart_CL = 0;
                X_chart_LCL = 0;
                X_chart_Val.Clear();

                R_chart_UCL = 0;
                R_chart_CL = 0;
                R_chart_Val.Clear();

                Chart_Header.Clear();

                //---------------------------------

                LValue1 = 0;
                LValue2 = 0;
                LValue3 = 0;
                LValue4 = 0;
                LValue5 = 0;

                LValue6 = 0;
                LValue7 = 0;
                LValue8 = 0;
                LValue9 = 0;
                LValue10 = 0;

                LValue11 = 0;
                LValue12 = 0;
                LValue13 = 0;
                LValue14 = 0;
                LValue15 = 0;

                LValue16 = 0;
                LValue17 = 0;
                LValue18 = 0;
                LValue19 = 0;
                LValue20 = 0;

                LValue21 = 0;
                LValue22 = 0;
                LValue23 = 0;
                LValue24 = 0;
                LValue25 = 0;

                LValue26 = 0;
                LValue27 = 0;
                LValue28 = 0;
                LValue29 = 0;
                LValue30 = 0;

                LValue31 = 0;
                LValue32 = 0;
                LValue33 = 0;
                LValue34 = 0;
                LValue35 = 0;

                LValue36 = 0;
                LValue37 = 0;
                LValue38 = 0;
                LValue39 = 0;
                LValue40 = 0;

                //---------------------------------

            }
        }

        #endregion

        //통계치 Summary 공간 채울 데이터 구하기.
        private void CreateDataGridColumns(string InspectPoint, int nchkDate, string sFromDate, string sToDate, string InspectBasisID,
                                        int InspectSubSeq)
        {
            try
            {
                if (dgdXBar_DailySpread.Columns.Count > 0)
                {
                    dgdXBar_DailySpread.Columns.Clear();
                }

                Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                sqlParameter.Clear();
                sqlParameter.Add("InspectPoint", InspectPoint);

                sqlParameter.Add("nchkDate", nchkDate);                     //int
                sqlParameter.Add("sFromDate", sFromDate);
                sqlParameter.Add("sToDate", sToDate);  //sMonth.Substring(0, 6));

                sqlParameter.Add("InspectBasisID", InspectBasisID);
                sqlParameter.Add("InspectSubSeq", InspectSubSeq);          //int

                sqlParameter.Add("CustomID", "");

                DataSet ds = DataStore.Instance.ProcedureToDataSet("xp_Qual_sSpc_DailySpread_20220124", sqlParameter, false);

                if (ds != null && ds.Tables.Count > 0)
                {
                    DataTable dt = ds.Tables[0];
                    if (dt.Rows.Count == 0)
                    {
                        return;
                    }
                    else
                    {
                        int DefectCnt = dt.Rows.Count;

                        DataGridTemplateColumn dgdTxtCol = new DataGridTemplateColumn();
                        dgdTxtCol.Header = "";
                        dgdTxtCol.MinWidth = 70;
                        FrameworkElementFactory tblt = new FrameworkElementFactory(typeof(TextBlock));
                        tblt.SetValue(TextBlock.TextAlignmentProperty, TextAlignment.Right);
                        tblt.SetValue(TextBlock.PaddingProperty, new Thickness(0, 0, 3, 0));
                        tblt.SetValue(TextBlock.VerticalAlignmentProperty, VerticalAlignment.Center);
                        tblt.SetValue(TextBlock.TextProperty, new Binding("lstInspect[" + 0 + "]"));

                        DataTemplate dataTt = new DataTemplate();
                        dataTt.VisualTree = tblt;

                        dgdTxtCol.CellTemplate = dataTt;

                        dgdXBar_DailySpread.Columns.Add(dgdTxtCol);

                        for (int i = 0; i < dt.Rows.Count; i++)
                        {
                            dgdTxtCol = new DataGridTemplateColumn();
                            dgdTxtCol.Header = " " + dt.Rows[i]["lotid"].ToString().Trim() + " ";
                            dgdTxtCol.MinWidth = 70;
                            FrameworkElementFactory tbl = new FrameworkElementFactory(typeof(TextBlock));
                            tbl.SetValue(TextBlock.TextAlignmentProperty, TextAlignment.Right);
                            tbl.SetValue(TextBlock.PaddingProperty, new Thickness(0, 0, 3, 0));
                            tbl.SetValue(TextBlock.VerticalAlignmentProperty, VerticalAlignment.Center);
                            tbl.SetValue(TextBlock.TextProperty, new Binding("lstInspect[" + (i + 1) + "]"));

                            DataTemplate dataT = new DataTemplate();
                            dataT.VisualTree = tbl;

                            dgdTxtCol.CellTemplate = dataT;

                            dgdXBar_DailySpread.Columns.Add(dgdTxtCol);
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



        private void FillGrid_Specification(int nchkDate, string FromDate, string ToDate,
                                            string InspectBasisID, int SubSeq)
        {
            Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
            sqlParameter.Clear();
            sqlParameter.Add("nchkDate", nchkDate);
            sqlParameter.Add("FromDate", FromDate);
            sqlParameter.Add("ToDate", ToDate);
            sqlParameter.Add("InspectBasisID", InspectBasisID);
            sqlParameter.Add("SubSeq", SubSeq);

            DataSet ds = DataStore.Instance.ProcedureToDataSet("xp_Qual_sSpc_spec_220124", sqlParameter, false);

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
                    DataRow dr = dt.Rows[0];
                    RASpecMax = dr["InsRASpecMax"].ToString();
                    RASpecMin = dr["InsRASpecMin"].ToString();
                }

            }
        }


        #region 통계치 Summary
        //통계치 Summary 공간 채울 데이터 구하기.
        private void FillGrid_Summary(string InspectPoint, int nchkDate, string sFromDate, string sToDate, string InspectBasisID,
                                        int InspectSubSeq)
        {
            Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
            sqlParameter.Clear();
            sqlParameter.Add("InspectPoint", InspectPoint);

            sqlParameter.Add("nchkDate", nchkDate);                     //int
            sqlParameter.Add("sFromDate", sFromDate);
            sqlParameter.Add("sToDate", sToDate);  //sMonth.Substring(0, 6));

            sqlParameter.Add("InspectBasisID", InspectBasisID);
            sqlParameter.Add("InspectSubSeq", InspectSubSeq);          //int

            sqlParameter.Add("CustomID", "");

            DataSet ds = DataStore.Instance.ProcedureToDataSet("xp_Qual_sSpc_DailySpread_20220124", sqlParameter, false);

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

                    dgdXBar_DailySpread.Items.Clear();
                    //dgdXBar_DailySpread.Columns.Clear();

                    int LotColumnCount = 1;
                    double X_VALUE = 0;
                    double R_VALUE = 0;
                    string Chart_H = string.Empty;

                    DataRowCollection drc = dt.Rows;

                    var WinQ = new Win_Qul_sts_XBarR_Q_View();
                    WinQ.lstInspect = new List<string>();

                    //WinQ = new Win_Qul_sts_XBarR_Q_View();
                    //WinQ.lstInspect = new List<string>();



                    //dgdXBar_DailySpread.Items.Add(WinQ1);

                    //for (int i = 1; i < 4; i++)
                    //{
                    //    var WinQ2 = new Win_Qul_sts_XBarR_Q_View()
                    //    {
                    //        columnHeader = i.ToString()
                    //    };
                    //    dgdXBar_DailySpread.Items.Add(WinQ2);
                    //}

                    X_VALUE = 0;                // 그릇 초기화.
                    R_VALUE = 0;                // 그릇 초기화.
                    Chart_H = string.Empty;     // 그릇 초기화.

                    DataGridTextColumn textColumn = new DataGridTextColumn();
                    textColumn.Header = "";
                    textColumn.MinWidth = 70;
                    textColumn.Binding = new Binding("L" + LotColumnCount);
                    //dgdXBar_DailySpread.Columns.Add(textColumn);

                    LotColumnCount++;
                    WinQ.lstInspect.Add("평균값");

                    foreach (DataRow item in drc)
                    {

                        X_VALUE = 0;                // 그릇 초기화.
                        R_VALUE = 0;                // 그릇 초기화.
                        Chart_H = string.Empty;     // 그릇 초기화.

                        DataGridTextColumn textColumn2 = new DataGridTextColumn();
                        textColumn2.Header = " " + item["lotid"].ToString().Trim() + " ";
                        Chart_H = item["lotid"].ToString().Trim();
                        textColumn2.MinWidth = 70;
                        textColumn2.Binding = new Binding("L" + LotColumnCount);
                        //dgdXBar_DailySpread.Columns.Add(textColumn);

                        Double.TryParse(item["AvgInspectValue"].ToString(), out X_VALUE);
                        Double.TryParse(item["R_Value"].ToString(), out R_VALUE);

                        X_chart_Val.Add(X_VALUE);
                        R_chart_Val.Add(R_VALUE);
                        Chart_Header.Add(Chart_H);

                        LotColumnCount++;
                        WinQ.lstInspect.Add(stringFormatN3(item["AvgInspectValue"]));
                    }
                    dgdXBar_DailySpread.Items.Add(WinQ);

                    //for (int i = 1; i < 4; i++)
                    //{
                    //    var WinQ2 = new Win_Qul_sts_XBarR_Q_View()
                    //    {
                    //        columnHeader = i.ToString()
                    //    };
                    //    dgdXBar_DailySpread.Items.Add(WinQ2);
                    //}

                    for (int i = 1; i < 4; i++)
                    {
                        WinQ = new Win_Qul_sts_XBarR_Q_View();

                        WinQ.lstInspect = new List<string>();

                        WinQ.lstInspect.Add((i).ToString());

                        foreach (DataRow item in drc)
                        {
                            WinQ.lstInspect.Add(stringFormatN3(item["InspectValue" + i]));
                        }
                        dgdXBar_DailySpread.Items.Add(WinQ);
                    }


                    //WinQ = new Win_Qul_sts_XBarR_Q_View();

                    //WinQ.lstInspect = new List<string>();

                    //foreach (DataRow item in drc)
                    //{

                    //    WinQ.lstInspect.Add(stringFormatN3(item["InspectValue2"]));

                    //}
                    //dgdXBar_DailySpread.Items.Add(WinQ);

                    //WinQ = new Win_Qul_sts_XBarR_Q_View();

                    //WinQ.lstInspect = new List<string>();

                    //foreach (DataRow item in drc)
                    //{

                    //    WinQ.lstInspect.Add(stringFormatN3(item["InspectValue3"]));

                    //}
                    //dgdXBar_DailySpread.Items.Add(WinQ);

                    int LotValueCount = 1;

                    for (LotValueCount = 1; LotValueCount < dt.Rows.Count + 1; LotValueCount++)
                    {

                        switch (LotValueCount)
                        {
                            case 1:
                                Double.TryParse(dt.Rows[LotValueCount - 1]["AvgInspectValue"].ToString(), out LValue1);
                                txtA2.Text = dt.Rows[LotValueCount - 1]["A2_Value"].ToString();
                                txtXbarUCL.Text = dt.Rows[LotValueCount - 1]["XBarUSL"].ToString();
                                txtXbarCL.Text = dt.Rows[LotValueCount - 1]["XBarCL"].ToString();
                                txtXbarLCL.Text = dt.Rows[LotValueCount - 1]["XBarLSL"].ToString();
                                txtRUCL.Text = dt.Rows[LotValueCount - 1]["RUCL"].ToString();
                                txtRCL.Text = dt.Rows[LotValueCount - 1]["RCL"].ToString();
                                txtMaxValue.Text = dt.Rows[LotValueCount - 1]["MaxValue"].ToString();
                                txtMinValue.Text = dt.Rows[LotValueCount - 1]["MinValue"].ToString();
                                txtAverage.Text = dt.Rows[LotValueCount - 1]["AvgValue"].ToString();
                                txtStandardDeviation.Text = dt.Rows[LotValueCount - 1]["StdevValue"].ToString();
                                break;
                                //case 2:
                                //    Double.TryParse(dt.Rows[LotValueCount - 1]["InspectValue"].ToString(), out LValue2);
                                //    break;
                                //case 3:
                                //    Double.TryParse(dt.Rows[LotValueCount - 1]["InspectValue"].ToString(), out LValue3);
                                //    break;
                                //case 4:
                                //    Double.TryParse(dt.Rows[LotValueCount - 1]["InspectValue"].ToString(), out LValue4);
                                //    break;
                                //case 5:
                                //    Double.TryParse(dt.Rows[LotValueCount - 1]["InspectValue"].ToString(), out LValue5);
                                //    break;
                                //case 6:
                                //    Double.TryParse(dt.Rows[LotValueCount - 1]["InspectValue"].ToString(), out LValue6);
                                //    break;
                                //case 7:
                                //    Double.TryParse(dt.Rows[LotValueCount - 1]["InspectValue"].ToString(), out LValue7);
                                //    break;
                                //case 8:
                                //    Double.TryParse(dt.Rows[LotValueCount - 1]["InspectValue"].ToString(), out LValue8);
                                //    break;
                                //case 9:
                                //    Double.TryParse(dt.Rows[LotValueCount - 1]["InspectValue"].ToString(), out LValue9);
                                //    break;
                                //case 10:
                                //    Double.TryParse(dt.Rows[LotValueCount - 1]["InspectValue"].ToString(), out LValue10);
                                //    break;
                                //case 11:
                                //    Double.TryParse(dt.Rows[LotValueCount - 1]["InspectValue"].ToString(), out LValue11);
                                //    break;
                                //case 12:
                                //    Double.TryParse(dt.Rows[LotValueCount - 1]["InspectValue"].ToString(), out LValue12);
                                //    break;
                                //case 13:
                                //    Double.TryParse(dt.Rows[LotValueCount - 1]["InspectValue"].ToString(), out LValue13);
                                //    break;
                                //case 14:
                                //    Double.TryParse(dt.Rows[LotValueCount - 1]["InspectValue"].ToString(), out LValue14);
                                //    break;
                                //case 15:
                                //    Double.TryParse(dt.Rows[LotValueCount - 1]["InspectValue"].ToString(), out LValue15);
                                //    break;
                                //case 16:
                                //    Double.TryParse(dt.Rows[LotValueCount - 1]["InspectValue"].ToString(), out LValue16);
                                //    break;
                                //case 17:
                                //    Double.TryParse(dt.Rows[LotValueCount - 1]["InspectValue"].ToString(), out LValue17);
                                //    break;
                                //case 18:
                                //    Double.TryParse(dt.Rows[LotValueCount - 1]["InspectValue"].ToString(), out LValue18);
                                //    break;
                                //case 19:
                                //    Double.TryParse(dt.Rows[LotValueCount - 1]["InspectValue"].ToString(), out LValue19);
                                //    break;
                                //case 20:
                                //    Double.TryParse(dt.Rows[LotValueCount - 1]["InspectValue"].ToString(), out LValue20);
                                //    break;
                                //case 21:
                                //    Double.TryParse(dt.Rows[LotValueCount - 1]["InspectValue"].ToString(), out LValue21);
                                //    break;
                                //case 22:
                                //    Double.TryParse(dt.Rows[LotValueCount - 1]["InspectValue"].ToString(), out LValue22);
                                //    break;
                                //case 23:
                                //    Double.TryParse(dt.Rows[LotValueCount - 1]["InspectValue"].ToString(), out LValue23);
                                //    break;
                                //case 24:
                                //    Double.TryParse(dt.Rows[LotValueCount - 1]["InspectValue"].ToString(), out LValue24);
                                //    break;
                                //case 25:
                                //    Double.TryParse(dt.Rows[LotValueCount - 1]["InspectValue"].ToString(), out LValue25);
                                //    break;
                                //case 26:
                                //    Double.TryParse(dt.Rows[LotValueCount - 1]["InspectValue"].ToString(), out LValue26);
                                //    break;
                                //case 27:
                                //    Double.TryParse(dt.Rows[LotValueCount - 1]["InspectValue"].ToString(), out LValue27);
                                //    break;
                                //case 28:
                                //    Double.TryParse(dt.Rows[LotValueCount - 1]["InspectValue"].ToString(), out LValue28);
                                //    break;
                                //case 29:
                                //    Double.TryParse(dt.Rows[LotValueCount - 1]["InspectValue"].ToString(), out LValue29);
                                //    break;
                                //case 30:
                                //    Double.TryParse(dt.Rows[LotValueCount - 1]["InspectValue"].ToString(), out LValue30);
                                //    break;
                                //case 31:
                                //    Double.TryParse(dt.Rows[LotValueCount - 1]["InspectValue"].ToString(), out LValue31);
                                //    break;
                                //case 32:
                                //    Double.TryParse(dt.Rows[LotValueCount - 1]["InspectValue"].ToString(), out LValue32);
                                //    break;
                                //case 33:
                                //    Double.TryParse(dt.Rows[LotValueCount - 1]["InspectValue"].ToString(), out LValue33);
                                //    break;
                                //case 34:
                                //    Double.TryParse(dt.Rows[LotValueCount - 1]["InspectValue"].ToString(), out LValue34);
                                //    break;
                                //case 35:
                                //    Double.TryParse(dt.Rows[LotValueCount - 1]["InspectValue"].ToString(), out LValue35);
                                //    break;
                                //case 36:
                                //    Double.TryParse(dt.Rows[LotValueCount - 1]["InspectValue"].ToString(), out LValue36);
                                //    break;
                                //case 37:
                                //    Double.TryParse(dt.Rows[LotValueCount - 1]["InspectValue"].ToString(), out LValue37);
                                //    break;
                                //case 38:
                                //    Double.TryParse(dt.Rows[LotValueCount - 1]["InspectValue"].ToString(), out LValue38);
                                //    break;
                                //case 39:
                                //    Double.TryParse(dt.Rows[LotValueCount - 1]["InspectValue"].ToString(), out LValue39);
                                //    break;
                                //case 40:
                                //    Double.TryParse(dt.Rows[LotValueCount - 1]["InspectValue"].ToString(), out LValue40);
                                //    break;
                        }
                    }

                    double dSTDev = Convert.ToDouble(txtStandardDeviation.Text);    // 표준편차 값.
                    double dAvr;
                    double dUSL;
                    double dLSL;

                    if (dSTDev == 0)
                    {
                        txtCPU.Text = "0";
                        txtCPL.Text = "0";
                        txtCP.Text = "0";
                    }
                    else
                    {
                        dAvr = Convert.ToDouble(txtAverage.Text);                // 평균.
                        dUSL = Convert.ToDouble(RASpecMax);           // 상한공차 값.
                        dLSL = Convert.ToDouble(RASpecMin);           // 하한공차 값.

                        txtCPU.Text = String.Format("{0:N3}", ((dUSL - dAvr) / (3 * dSTDev)));
                        txtCPL.Text = String.Format("{0:N3}", ((dAvr - dLSL) / (3 * dSTDev)));
                        txtCP.Text = String.Format("{0:N3}", ((dUSL - dLSL) / (6 * dSTDev)));
                    }
                    if (Convert.ToDouble(txtCPU.Text) > Convert.ToDouble(txtCPL.Text))
                    {
                        txtCPK.Text = txtCPL.Text;
                    }
                    else
                    {
                        txtCPK.Text = txtCPU.Text;
                    }


                }
            }
        }

        // 천마리 콤마, 소수점 버리기
        private string stringFormatN3(object obj)
        {
            if (obj == null)
            {
                obj = 0;
            }

            return string.Format("{0:N3}", obj);
        }

        #endregion


        #region 차트
        // 차트 ㅡ그리기.
        private void FillChart_Double()
        {
            if (txtXbarUCL.Text != string.Empty) { X_chart_UCL = Convert.ToDouble(txtXbarUCL.Text); }
            if (txtXbarCL.Text != string.Empty) { X_chart_CL = Convert.ToDouble(txtXbarCL.Text); }
            if (txtXbarLCL.Text != string.Empty) { X_chart_LCL = Convert.ToDouble(txtXbarLCL.Text); }

            if (txtRUCL.Text != string.Empty) { R_chart_UCL = Convert.ToDouble(txtRUCL.Text); }
            if (txtRCL.Text != string.Empty) { R_chart_CL = Convert.ToDouble(txtRCL.Text); }


            ChartValues<double> XUCL_Charts = new ChartValues<double>();
            ChartValues<double> XCL_Charts = new ChartValues<double>();
            ChartValues<double> XLCL_Charts = new ChartValues<double>();
            ChartValues<double> XVal_Charts = new ChartValues<double>();


            int j = 0;
            for (int i = 0; i < X_chart_Val.Count; i++)
            {
                if (X_chart_Val[i] > 0.00)
                {
                    j++;
                    XUCL_Charts.Add(X_chart_UCL);
                    XCL_Charts.Add(X_chart_CL);
                    XLCL_Charts.Add(X_chart_LCL);
                    XVal_Charts.Add(X_chart_Val[i]);
                }
            }

            SeriesCollection SeriesCollection = new SeriesCollection
            {
                new LineSeries
                {
                    Title = "UCL",
                    Values = XUCL_Charts
                },
                new LineSeries
                {
                    Title = "CL",
                    Values = XCL_Charts
                },
                new LineSeries
                {
                    Title = "LCL",
                    Values = XLCL_Charts
                },
                new LineSeries
                {
                    Title = "측정값",
                    Values = XVal_Charts
                }
            };

            DataContext = null;
            X_Linelbl = Chart_Header.ToArray();
            lvcXBarChart.Series = SeriesCollection;



            ChartValues<double> RUCL_Charts = new ChartValues<double>();
            ChartValues<double> RCL_Charts = new ChartValues<double>();
            ChartValues<double> RVal_Charts = new ChartValues<double>();

            for (int i = 0; i < R_chart_Val.Count; i++)
            {
                if (R_chart_Val[i] > 0.00)
                {
                    j++;
                    RUCL_Charts.Add(R_chart_UCL);
                    RCL_Charts.Add(R_chart_CL);
                    RVal_Charts.Add(R_chart_Val[i]);
                }
            }


            SeriesCollection SeriesCollection2 = new SeriesCollection
            {
                new LineSeries
                {
                    Title = "UCL",
                    Values = RUCL_Charts
                },
                new LineSeries
                {
                    Title = "CL",
                    Values = RCL_Charts
                },
                new LineSeries
                {
                    Title = "측정값",
                    Values = RVal_Charts
                }
            };

            X_Linelbl = Chart_Header.ToArray();
            lvcRChart.Series = SeriesCollection2;
            DataContext = this;
        }

        #endregion



        private void DataGrid_SizeChange(object sender, SizeChangedEventArgs e)
        {
            DataGrid dgs = sender as DataGrid;
            if (dgs.ColumnHeaderHeight == 0)
            {
                dgs.ColumnHeaderHeight = 1;
            }
            double a = e.NewSize.Height / 100;
            double b = e.PreviousSize.Height / 100;
            double c = a / b;

            if (c != double.PositiveInfinity && c != 0 && double.IsNaN(c) == false)
            {
                dgs.ColumnHeaderHeight = dgs.ColumnHeaderHeight * c;
                dgs.FontSize = dgs.FontSize * c;
            }
        }

        #region Content 통합 검색
        private void lblTotSearch_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            if (chkTotSearch.IsChecked == false)
            {
                chkTotSearch.IsChecked = true;
                cboTotSearch.IsEnabled = true;
                txtTotSearch.IsEnabled = true;
            }
            else
            {
                chkTotSearch.IsChecked = false;
                cboTotSearch.IsEnabled = false;
                txtTotSearch.IsEnabled = false;
            }
        }

        private void chkTotSearch_Click(object sender, RoutedEventArgs e)
        {
            if (chkTotSearch.IsChecked == false)
            {
                chkTotSearch.IsChecked = true;
                cboTotSearch.IsEnabled = true;
                txtTotSearch.IsEnabled = true;
            }
            else
            {
                chkTotSearch.IsChecked = false;
                cboTotSearch.IsEnabled = false;
                txtTotSearch.IsEnabled = false;
            }
        }

        private void txtTotSearch_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                if (txtTotSearch.Text != "")
                {
                    FillGrid();
                }
                else
                {
                    MessageBox.Show("검색할 내용이 없습니다.", "주의");

                    dgdXBar_std.Items.Clear();
                    dgdXBar_DailySpread.Items.Clear();

                    SpsSummarytextBoxClear_Up();
                    SpsSummarytextBoxClear_Down();

                    if (lvcXBarChart.Series != null) lvcXBarChart.Series.Clear();

                    if (lvcRChart.Series != null) lvcRChart.Series.Clear();

                    return;
                }
            }
        }

        private void SpsSummarytextBoxClear_Up()
        {
            txtCPK.Clear(); txtCPL.Clear();
            txtCPU.Clear(); txtCP.Clear();
            txtMaxValue.Clear(); txtMinValue.Clear();
            txtAverage.Clear(); txtStandardDeviation.Clear();
        }

        private void SpsSummarytextBoxClear_Down()
        {
            txtXbarCL.Clear(); txtXbarUCL.Clear();
            txtXbarLCL.Clear(); txtRUCL.Clear();
            txtRCL.Clear();
        }
        #endregion

        #region 토글버튼 이벤트
        private void TbnInCome_Click(object sender, RoutedEventArgs e)
        {
            ToggleButton toggleButton = sender as ToggleButton;
            if (toggleButton.IsChecked == true)
            {
                tbnOutCome.IsChecked = false;
                tbnProcessCycle.IsChecked = false;
                tbnJaju.IsChecked = false;
            }
            else
            {
                toggleButton.IsChecked = true;
            }
        }

        private void TbnProcessCycle_Click(object sender, RoutedEventArgs e)
        {
            ToggleButton toggleButton = sender as ToggleButton;
            if (toggleButton.IsChecked == true)
            {
                tbnInCome.IsChecked = false;
                tbnOutCome.IsChecked = false;
                tbnJaju.IsChecked = false;
            }
            else
            {
                toggleButton.IsChecked = true;
            }
        }

        private void TbnJaju_Click(object sender, RoutedEventArgs e)
        {
            ToggleButton toggleButton = sender as ToggleButton;
            if (toggleButton.IsChecked == true)
            {
                tbnInCome.IsChecked = false;
                tbnOutCome.IsChecked = false;
                tbnProcessCycle.IsChecked = false;
            }
            else
            {
                toggleButton.IsChecked = true;
            }
        }

        private void TbnOutCome_Click(object sender, RoutedEventArgs e)
        {
            ToggleButton toggleButton = sender as ToggleButton;
            if (toggleButton.IsChecked == true)
            {
                tbnInCome.IsChecked = false;
                tbnProcessCycle.IsChecked = false;
                tbnJaju.IsChecked = false;
            }
            else
            {
                toggleButton.IsChecked = true;
            }
        }

        #endregion

        #region 날짜 선택 버튼 이벤트
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
        #endregion
    }
}

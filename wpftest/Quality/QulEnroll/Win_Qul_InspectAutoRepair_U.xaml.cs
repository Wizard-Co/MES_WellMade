using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using WizMes_WellMade.PopUP;
using WPF.MDI;

namespace WizMes_WellMade
{
    /// <summary>
    /// Win_Qul_InspectAutoRepair_U.xaml에 대한 상호 작용 논리
    /// </summary>
    public partial class Win_Qul_InspectAutoRepair_U : UserControl
    {
        #region 전역 변수선언
        Lib lib = new Lib();
        PlusFinder pf = new PlusFinder();

        int Wh_Ar_SelectedLastIndex = 0;        // 그리드 마지막 선택 줄 임시저장 그릇


        // FTP 활용모음.
        List<string[]> listFtpFile = new List<string[]>();
        private FTP_EX _ftp = null;
        string FullPath1 = string.Empty;

        //string FTP_ADDRESS = "ftp://wizis.iptime.org/ImageData/AutoInspect";
        //string FTP_ADDRESS = "ftp://HKserver:210/ImageData/AutoInspect";
        string FTP_ADDRESS = "ftp://" + LoadINI.FileSvr + ":" + LoadINI.FTPPort + LoadINI.FtpImagePath + "/AutoInspect";
        private const string FTP_ID = "wizuser";
        private const string FTP_PASS = "wiz9999";
        private const string LOCAL_DOWN_PATH = "C:\\Temp";

        #endregion

        public Win_Qul_InspectAutoRepair_U()
        {
            InitializeComponent();
        }

        // 첫 로드시.
        private void Win_Qul_InspectAutoRepair_U_Loaded(object sender, RoutedEventArgs e)
        {
            First_Step();
            ComboBoxSetting();
        }


        #region 첫 스텝 // 날짜용 버튼 // 조회용 체크박스 세팅 
        // 첫 스텝
        private void First_Step()
        {
            chkInspectDay.IsChecked = true;
            dtpFromDate.Text = DateTime.Now.ToString("yyyy-MM-dd");
            dtpToDate.Text = DateTime.Now.ToString("yyyy-MM-dd");

            txtArticle.IsEnabled = false;
            btnArticle.IsEnabled = false;
            rbnChoiceAll.IsChecked = true;
            txtBuyerArticleNo.IsEnabled = false;
            btnBuyerArticleNo.IsEnabled = false;

            btnSave.Visibility = Visibility.Hidden;
            btnCancel.Visibility = Visibility.Hidden;
            EventLabel.Visibility = Visibility.Hidden;

            grbInspectBox.IsEnabled = false;
            grbAttachBox.IsEnabled = false;
            grbRepairBox.IsEnabled = false;
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

        //검사일자
        private void chkInspectDay_Click(object sender, RoutedEventArgs e)
        {
            if (chkInspectDay.IsChecked == true)
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
        private void chkInspectDay_Click(object sender, MouseButtonEventArgs e)
        {
            if (chkInspectDay.IsChecked == true)
            {
                chkInspectDay.IsChecked = false;
                dtpFromDate.IsEnabled = false;
                dtpToDate.IsEnabled = false;
            }
            else
            {
                chkInspectDay.IsChecked = true;
                dtpFromDate.IsEnabled = true;
                dtpToDate.IsEnabled = true;
            }
        }
        //품명
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
        //품명
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


        #region 콤보박스 세팅

        //콤보박스 세팅
        private void ComboBoxSetting()
        {
            cboInspectPoint.Items.Clear();
            cboDefectReason.Items.Clear();
            cboDefectYN.Items.Clear();
            cboPartQualityRisk.Items.Clear();
            cboInspectGubun.Items.Clear();
            cboInspectClss.Items.Clear();
            cboRepairYN.Items.Clear();


            DataTable dt = new DataTable();
            dt.Columns.Add("value");
            dt.Columns.Add("display");

            DataRow row0 = dt.NewRow();
            row0["value"] = "1";
            row0["display"] = "인수검사";
            DataRow row1 = dt.NewRow();
            row1["value"] = "3";
            row1["display"] = "공정순회검사";
            DataRow row2 = dt.NewRow();
            row2["value"] = "5";
            row2["display"] = "출하샘플검사";
            DataRow row3 = dt.NewRow();
            row3["value"] = "9";
            row3["display"] = "자주검사";

            dt.Rows.Add(row0);
            dt.Rows.Add(row1);
            dt.Rows.Add(row2);
            dt.Rows.Add(row3);

            this.cboInspectPoint.ItemsSource = dt.DefaultView;
            this.cboInspectPoint.DisplayMemberPath = "display";
            this.cboInspectPoint.SelectedValuePath = "value";
            this.cboInspectPoint.SelectedIndex = 0;

            ////////////////////////////////////////////////////////////////////////////

            dt = new DataTable();
            dt.Columns.Add("value");
            dt.Columns.Add("display");

            row0 = dt.NewRow();
            row0["value"] = "Y";
            row0["display"] = "불합격";

            row1 = dt.NewRow();
            row1["value"] = "N";
            row1["display"] = "합격";

            dt.Rows.Add(row0);
            dt.Rows.Add(row1);

            this.cboDefectYN.ItemsSource = dt.DefaultView;
            this.cboDefectYN.DisplayMemberPath = "display";
            this.cboDefectYN.SelectedValuePath = "value";
            this.cboInspectPoint.SelectedIndex = 0;


            ////////////////////////////////////////////////////////////////////////////

            dt = new DataTable();
            dt.Columns.Add("value");
            dt.Columns.Add("display");

            row0 = dt.NewRow();
            row0["value"] = "Y";
            row0["display"] = "종결";

            row1 = dt.NewRow();
            row1["value"] = "N";
            row1["display"] = "미 종결";

            dt.Rows.Add(row0);
            dt.Rows.Add(row1);

            this.cboRepairYN.ItemsSource = dt.DefaultView;
            this.cboRepairYN.DisplayMemberPath = "display";
            this.cboRepairYN.SelectedValuePath = "value";
            this.cboRepairYN.SelectedIndex = -1;

            //////////////////////////////////////////////////////////////////////////////

            ObservableCollection<CodeView> cbDefectReason = ComboBoxUtil.Instance.Gf_DB_CM_GetComCodeDataset(null, "QULDFRSN", "Y", "", "");
            ObservableCollection<CodeView> cbPartQualityRisk = ComboBoxUtil.Instance.Gf_DB_CM_GetComCodeDataset(null, "INSDNGRLVL", "Y", "", "");
            ObservableCollection<CodeView> cbInspectGubun = ComboBoxUtil.Instance.Gf_DB_CM_GetComCodeDataset(null, "INSPECTGBN", "Y", "", "");
            ObservableCollection<CodeView> cbInspectClss = ComboBoxUtil.Instance.Gf_DB_CM_GetComCodeDataset(null, "INSPECTCLSS", "Y", "", "");

            this.cboDefectReason.ItemsSource = cbDefectReason;
            this.cboDefectReason.DisplayMemberPath = "code_name";
            this.cboDefectReason.SelectedValuePath = "code_id";
            this.cboDefectReason.SelectedIndex = -1;

            this.cboPartQualityRisk.ItemsSource = cbPartQualityRisk;
            this.cboPartQualityRisk.DisplayMemberPath = "code_name";
            this.cboPartQualityRisk.SelectedValuePath = "code_id";
            this.cboPartQualityRisk.SelectedIndex = -1;

            this.cboInspectGubun.ItemsSource = cbInspectGubun;
            this.cboInspectGubun.DisplayMemberPath = "code_name";
            this.cboInspectGubun.SelectedValuePath = "code_id";
            this.cboInspectGubun.SelectedIndex = -1;

            this.cboInspectClss.ItemsSource = cbInspectClss;
            this.cboInspectClss.DisplayMemberPath = "code_name";
            this.cboInspectClss.SelectedValuePath = "code_id";
            this.cboInspectClss.SelectedIndex = -1;
        }

        #endregion


        // 플러스 파인더 _ 품번찾기.
        private void btnArticle_Click(object sender, RoutedEventArgs e)
        {
            pf.ReturnCode(txtArticle, 77, txtArticle.Text);
        }

        // 키다운 플러스 파인더 _ 품번찾기.
        private void TxtArticle_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                pf.ReturnCode(txtArticle, 77, txtArticle.Text);
            }
        }


        #region 공통버튼 이벤트

        //공통 사용가능
        private void PublicEnableTrue()
        {
            btnSave.Visibility = Visibility.Hidden;
            btnCancel.Visibility = Visibility.Hidden;
            EventLabel.Visibility = Visibility.Hidden;

            btnSearch.IsEnabled = true;
            btnEdit.IsEnabled = true;
            btnDelete.IsEnabled = true;
            btnExcel.IsEnabled = true;

            txtMillFile.IsEnabled = false;
            txtAttachFile.IsEnabled = false;
            btnAttachDown.IsEnabled = false;
            btnMillDown.IsEnabled = false;

            grbAttachBox.IsEnabled = false;
            grbRepairBox.IsEnabled = false;

            //dgdInspect.IsEnabled = true; //메인그리드 사용가능.
            dgdInspect.IsHitTestVisible = true; //메인그리드 사용가능.
        }

        // 공통 버튼이벤트.
        private void PublicEnableFalse()
        {
            btnSave.Visibility = Visibility.Visible;
            btnCancel.Visibility = Visibility.Visible;
            EventLabel.Visibility = Visibility.Visible;

            btnSearch.IsEnabled = false;
            btnEdit.IsEnabled = false;
            btnDelete.IsEnabled = false;
            btnExcel.IsEnabled = false;

            grbRepairBox.IsEnabled = true;

            //dgdInspect.IsEnabled = false;    // 메인그리드 못건드리게.
            dgdInspect.IsHitTestVisible = false;    // 메인그리드 못건드리게.
        }

        #endregion


        #region  조회 / 조회용 프로시저

        // 검색 버튼 클릭.
        private void btnSearch_Click(object sender, RoutedEventArgs e)
        {
            //검색버튼 비활성화
            btnSearch.IsEnabled = false;

            Dispatcher.BeginInvoke(new Action(() =>

            {
                Thread.Sleep(2000);

                //로직
                Wh_Ar_SelectedLastIndex = 0;
                re_Search(Wh_Ar_SelectedLastIndex);

            }), System.Windows.Threading.DispatcherPriority.Background);



            Dispatcher.BeginInvoke(new Action(() =>

            {
                btnSearch.IsEnabled = true;

            }), System.Windows.Threading.DispatcherPriority.Background);

        }

        private void re_Search(int selectedIndex)
        {
            FillGrid();

            if (dgdInspect.Items.Count > 0)
            {
                dgdInspect.SelectedIndex = selectedIndex;
            }
        }

        private void FillGrid()
        {
            if (dgdInspect.Items.Count > 0)
            {
                dgdInspect.Items.Clear();
            }

            // 시작단계에서 생성 후 바로 첫번째 값 부여, 널일수가 없음.
            string InspectPoint = cboInspectPoint.SelectedValue.ToString();

            string FromDate = "";
            string ToDate = "";
            string ArticleID = "";
            int nchkArticleID = 0;
            string Article = "";
            int nTargetGubun = 0;

            if (chkInspectDay.IsChecked == true)
            {
                FromDate = dtpFromDate.ToString().Substring(0, 10).Replace("-", "");
                ToDate = dtpToDate.ToString().Substring(0, 10).Replace("-", "");
            }

            if (chkArticle.IsChecked == true)
            {
                if (txtArticle.Tag == null)
                {
                    txtArticle.Tag = "";
                    if (txtArticle.Text.Length > 0)
                    {
                        nchkArticleID = 2;
                        Article = txtArticle.Text;
                    }
                }
                else
                {
                    nchkArticleID = 1;
                    ArticleID = txtArticle.Tag.ToString();
                }
            }

            if (rbnChoiceAll.IsChecked == true) { nTargetGubun = 0; }
            else if (rbnChoiceING.IsChecked == true) { nTargetGubun = 1; }
            else if (rbnChoiceFinished.IsChecked == true) { nTargetGubun = 2; }

            try
            {

                Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                sqlParameter.Clear();
                sqlParameter.Add("InspectPoint", InspectPoint);
                sqlParameter.Add("FromDate", FromDate);
                sqlParameter.Add("ToDate", ToDate);
                sqlParameter.Add("nchkArticleID", chkArticle.IsChecked == true ? 1:0);    //int.
                sqlParameter.Add("ArticleID", chkArticle.IsChecked == true ? (txtArticle.Tag != null ? txtArticle.Tag.ToString() : ""):""); 
                                                  
                sqlParameter.Add("nTargetGubun", nTargetGubun);       //int.
                sqlParameter.Add("nchkBuyerArticleNo", chkBuyerArticleNo.IsChecked == true ? 1 : 0);
                sqlParameter.Add("BuyerArticleNo",chkBuyerArticleNo.IsChecked == true? (txtBuyerArticleNo.Tag != null ? txtBuyerArticleNo.Tag.ToString() : ""):"");       //int.

                DataSet ds = DataStore.Instance.ProcedureToDataSet("xp_Inspect_sAutoInspectRepair", sqlParameter, false);

                if (ds != null && ds.Tables.Count > 0)
                {
                    DataTable dt = null;
                    dt = ds.Tables[0];
                    int i = 0;
                    if (dt.Rows.Count == 0)
                    {
                        MessageBox.Show("조회결과가 없습니다.");
                        return;
                    }
                    else
                    {
                        //조회결과가 있다면,
                        dgdInspect.Items.Clear();
                        DataRowCollection drc = dt.Rows;
                        foreach (DataRow item in drc)
                        {
                            var Win_Qul_InspectAutoRepair_U_Insert = new Win_Qul_InspectAutoRepair_U_View()
                            {
                                InspectID = item["InspectID"].ToString(),
                                InspectSeq = item["InspectSeq"].ToString(),
                                ArticleID = item["ArticleID"].ToString(),
                                Article = item["Article"].ToString(),
                                InspectGubun = item["InspectGubun"].ToString(),

                                InspectDate = item["InspectDate"].ToString(),
                                Full_InspectDate = DateTime.ParseExact(item["InspectDate"].ToString(), "yyyyMMdd", null).ToString("yyyy-MM-dd"),

                                LotID = item["LotID"].ToString(),
                                InspectQty = item["InspectQty"].ToString(),
                                ECONo = item["ECONo"].ToString(),
                                Comments = item["Comments"].ToString(),

                                InspectLevel = item["InspectLevel"].ToString(),
                                SketchPath = item["SketchPath"].ToString(),
                                SketchFile = item["SketchFile"].ToString(),
                                AttachedPath = item["AttachedPath"].ToString(),
                                AttachedFile = item["AttachedFile"].ToString(),

                                InspectUserID = item["InspectUserID"].ToString(),
                                InspectBasisID = item["InspectBasisID"].ToString(),
                                ProcessID = item["ProcessID"].ToString(),
                                DefectYN = item["DefectYN"].ToString(),
                                Process = item["Process"].ToString(),

                                BuyerArticleNo = item["BuyerArticleNo"].ToString(),
                                InspectPoint = item["InspectPoint"].ToString(),
                                ImportSecYN = item["ImportSecYN"].ToString(),
                                ImportlawYN = item["ImportlawYN"].ToString(),
                                ImportImpYN = item["ImportImpYN"].ToString(),

                                ImportNorYN = item["ImportNorYN"].ToString(),
                                IRELevel = item["IRELevel"].ToString(),
                                IRELevelName = item["IRELevelName"].ToString(),
                                InpCustomID = item["InpCustomID"].ToString(),
                                InpCustomName = item["InpCustomName"].ToString(),

                                InpDate = item["InpDate"].ToString(),
                                OutCustomID = item["OutCustomID"].ToString(),
                                OutCustomName = item["OutCustomName"].ToString(),
                                OutDate = item["OutDate"].ToString(),
                                MachineID = item["MachineID"].ToString(),

                                BuyerModelID = item["BuyerModelID"].ToString(),
                                BuyerModel = item["BuyerModel"].ToString(),
                                FMLGubun = item["FMLGubun"].ToString(),
                                insItemName = item["insItemName"].ToString(),
                                InspectSpec = item["InspectSpec"].ToString(),

                                InspectValue = item["InspectValue"].ToString(),
                                RepairYN = item["RepairYN"].ToString(),
                                DefectReasonCode = item["DefectReasonCode"].ToString(),
                                ReasonImput = item["ReasonImput"].ToString(),
                                CorrContents = item["CorrContents"].ToString(),  //문제발생내역

                                DefectRespectContents = item["DefectRespectContents"].ToString()  // 개선대책

                            };
                            Win_Qul_InspectAutoRepair_U_Insert.InspectQty = lib.returnNumString(Win_Qul_InspectAutoRepair_U_Insert.InspectQty);
                            dgdInspect.Items.Add(Win_Qul_InspectAutoRepair_U_Insert);
                            i++;
                        }

                        tbkIndexCount.Text = "▶검색결과 : " + i + "건";
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


        #region 그리드 셀 체인지 _ Show Data

        // 그리드 셀 체인지 이벤트 _ show data.
        private void dgdInspect_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            this.DataContext = dgdInspect.SelectedItem as Win_Qul_InspectAutoRepair_U_View;
            var ViewReceiver = dgdInspect.SelectedItem as Win_Qul_InspectAutoRepair_U_View;

            if (ViewReceiver != null)
            {
                if ((ViewReceiver.InspectPoint == "3") || (ViewReceiver.InspectPoint == "9"))   // 자주, 공정 검색일때만,
                {
                    string ProcessID = ViewReceiver.ProcessID;
                    string MachineID = ViewReceiver.MachineID;
                    FindMachineName(ProcessID, MachineID);      // 공정정보 기반으로 호기 명 찾기.
                }

                string PersonID = ViewReceiver.InspectUserID;
                //FindInspector(PersonID);    // ID로 검사자 명 찾기._아이디 자체를 명으로 가져오네.. 필요없슴.

                if (ViewReceiver.SketchPath != string.Empty)
                {
                    grbAttachBox.IsEnabled = true;
                    txtMillFile.IsEnabled = false;
                    txtAttachFile.IsEnabled = false;
                    btnAttachDown.IsEnabled = false;

                    btnMillDown.IsEnabled = true;
                }

                if (ViewReceiver.AttachedPath != string.Empty)
                {
                    grbAttachBox.IsEnabled = true;
                    txtMillFile.IsEnabled = false;
                    txtAttachFile.IsEnabled = false;
                    btnMillDown.IsEnabled = false;

                    btnAttachDown.IsEnabled = true;
                }

            }
        }

        #endregion


        #region (수정, 저장, 삭제, 취소) 버튼 이벤트 모음

        //수정(작성) 버튼 클릭 시.
        private void btnEdit_Click(object sender, RoutedEventArgs e)
        {
            // 1. 수정할 자격은 있는거야? 조회? 데이터 선택??
            if (dgdInspect.Items.Count < 1)
            {
                MessageBox.Show("먼저 검색해 주세요.");
                return;
            }
            var OBJ = dgdInspect.SelectedItem as Win_Qul_InspectAutoRepair_U_View;
            if (OBJ == null)
            {
                MessageBox.Show("수정할 항목이 정확히 선택되지 않았습니다.");
                return;
            }

            Wh_Ar_SelectedLastIndex = dgdInspect.SelectedIndex;

            // 공통 버튼이벤트
            PublicEnableFalse();
            EventLabel.Content = "자료입력(작성) 중..";

            cboDefectReason.Focus();
            cboDefectReason.IsDropDownOpen = true;
        }

        //삭제버튼 클릭 시.
        private void btnDelete_Click(object sender, RoutedEventArgs e)
        {
            try
            {

                // 1. 삭제할 자격은 있는거야? 조회? 데이터 선택??
                if (dgdInspect.Items.Count < 1)
                {
                    MessageBox.Show("먼저 검색해 주세요.");
                    return;
                }
                var OBJ = dgdInspect.SelectedItem as Win_Qul_InspectAutoRepair_U_View;
                if (OBJ == null)
                {
                    MessageBox.Show("삭제할 항목이 정확히 선택되지 않았습니다.");
                    return;
                }

                MessageBoxResult msgresult = MessageBox.Show("선택하신 항목을 삭제하시겠습니까?", "삭제 확인", MessageBoxButton.YesNo);
                if (msgresult == MessageBoxResult.Yes)
                {
                    string InspectID = txtInspectID.Text;
                    int InspectSeq = 0;
                    if (txtInspectSeq.Text != string.Empty) { InspectSeq = Convert.ToInt32(txtInspectSeq.Text); }
                    string DefectReasonCode = string.Empty;
                    string ReasonImput = string.Empty;
                    string CorrContents = string.Empty;
                    string DefectRespectContents = string.Empty;
                    string RepairYN = "N";

                    if (dgdInspect.Items.Count > 0 && dgdInspect.SelectedItem != null)
                    {
                        Wh_Ar_SelectedLastIndex = dgdInspect.SelectedIndex;
                    }


                    //말이 좋아 삭제지, 실상은 불량정보만 새로 지워 업데이트 치는 과정
                    SaveData(InspectID, InspectSeq, DefectReasonCode, ReasonImput,
                                CorrContents, DefectRespectContents, RepairYN);

                    dgdInspect.Refresh();

                    Wh_Ar_SelectedLastIndex -= 1;
                    re_Search(Wh_Ar_SelectedLastIndex);

                    // 2. 화면정리.
                    btnCancel_Click(null, null);
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

        // 저장 버튼 클릭 시.
        private void btnSave_Click(object sender, RoutedEventArgs e)
        {
            string InspectID = txtInspectID.Text;
            int InspectSeq = 0;
            if (txtInspectSeq.Text != string.Empty) { InspectSeq = Convert.ToInt32(txtInspectSeq.Text); }
            string DefectReasonCode = string.Empty;
            if (cboDefectReason.SelectedIndex >= 0)
            {
                DefectReasonCode = cboDefectReason.SelectedValue.ToString();
            }
            string ReasonImput = txtReasonImput.Text;
            string CorrContents = txtCorrContents.Text;
            string DefectRespectContents = txtDefectRespectContents.Text;
            string RepairYN = cboRepairYN.SelectedValue.ToString();

            SaveData(InspectID, InspectSeq, DefectReasonCode, ReasonImput,
                        CorrContents, DefectRespectContents, RepairYN);

            //공통 버튼이벤트
            PublicEnableTrue();

            //그룹박스 데이터 클리어
            grbEnrollBoxDataClear();

            re_Search(Wh_Ar_SelectedLastIndex);
            dgdInspect.Focus();

        }

        // 취소 버튼 클릭 시.
        private void btnCancel_Click(object sender, RoutedEventArgs e)
        {
            //공통 버튼이벤트
            PublicEnableTrue();

            //그룹박스 데이터 클리어
            grbEnrollBoxDataClear();

            re_Search(Wh_Ar_SelectedLastIndex);
            dgdInspect.Focus();

        }

        #endregion


        #region CRUD / 각종 프로시저 모음.

        // 저장
        private void SaveData(string InspectID, int InspectSeq, string DefectReasonCode, string ReasonImput,
                                string CorrContents, string DefectRespectContents, string RepairYN)
        {

            try
            {
                Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                sqlParameter.Clear();
                sqlParameter.Add("InspectID", InspectID);
                sqlParameter.Add("InspectSeq", InspectSeq);            //int.
                sqlParameter.Add("DefectReasonCode", DefectReasonCode);
                sqlParameter.Add("ReasonImput", ReasonImput);
                sqlParameter.Add("CorrContents", CorrContents);
                sqlParameter.Add("DefectRespectContents", DefectRespectContents);
                sqlParameter.Add("RepairYN", RepairYN);
                sqlParameter.Add("UserID", MainWindow.CurrentUser);

                string[] result = DataStore.Instance.ExecuteProcedure("xp_Inspect_uAutoInspectRepair", sqlParameter, false);
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




        // 공정정보 기반으로 호기 명 찾기.
        private void FindMachineName(string ProcessID, string MachineID)
        {
            Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
            sqlParameter.Clear();
            sqlParameter.Add("sProcessID", ProcessID);

            DataSet ds = DataStore.Instance.ProcedureToDataSet("xp_Process_sMachine", sqlParameter, false);

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
                    for (int i = 0; i < dt.Rows.Count; i++)
                    {
                        if (MachineID == dt.Rows[i]["MachineID"].ToString())
                        {
                            txtMachine.Text = dt.Rows[i]["Machine"].ToString();
                        }
                    }
                }
            }

            DataStore.Instance.CloseConnection();
        }

        // ID로 검사자 명 찾기.
        private void FindInspector(string PersonID)
        {
            Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
            sqlParameter.Clear();
            sqlParameter.Add("PersonID", PersonID);

            DataSet ds = DataStore.Instance.ProcedureToDataSet("xp_Person_sPersonByPersonID", sqlParameter, false);

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
                    txtInspector.Text = dt.Rows[0]["Name"].ToString();
                }
            }

            DataStore.Instance.CloseConnection();
        }

        #endregion



        // 작성상태로 변하는 grb의 데이터 클리어.
        // (grbRepairBox)
        private void grbEnrollBoxDataClear()
        {
            cboDefectReason.SelectedIndex = -1;
            txtReasonImput.Text = string.Empty;
            txtDefectRespectContents.Text = string.Empty;
            txtCorrContents.Text = string.Empty;
            cboRepairYN.SelectedIndex = -1;
        }


        //닫기.
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
                    }
                }
                i++;
            }
        }


        #region 엑셀

        //엑셀.
        private void btnExcel_Click(object sender, RoutedEventArgs e)
        {
            if (dgdInspect.Items.Count < 1)
            {
                MessageBox.Show("먼저 검색해 주세요.");
                return;
            }

            Lib lib2 = new Lib();

            DataTable dt = null;
            string Name = string.Empty;

            string[] lst = new string[2];
            lst[0] = "메인그리드";
            lst[1] = dgdInspect.Name;

            ExportExcelxaml ExpExc = new ExportExcelxaml(lst);

            ExpExc.ShowDialog();

            if (ExpExc.DialogResult.HasValue)
            {
                if (ExpExc.choice.Equals(dgdInspect.Name))
                {
                    //MessageBox.Show("대분류");
                    if (ExpExc.Check.Equals("Y"))
                        dt = lib2.DataGridToDTinHidden(dgdInspect);
                    else
                        dt = lib2.DataGirdToDataTable(dgdInspect);

                    Name = dgdInspect.Name;

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

            lib2 = null;
        }



        #endregion


        #region 내려받기 _
        // 내려받기 버튼 클릭시. (공통으로 물고 가보자.)
        private void btnFileDown_Click(object sender, RoutedEventArgs e)
        {
            MessageBoxResult msgresult = MessageBox.Show("파일을 다운로드 하시겠습니까?", "다운로드 확인", MessageBoxButton.YesNo);
            if (msgresult == MessageBoxResult.Yes)
            {
                //버튼 태그값.
                string ClickPoint = ((Button)sender).Tag.ToString();

                if ((ClickPoint == "1") && (txtMillFile.Tag == null))
                {
                    MessageBox.Show("파일이 없습니다.");
                    return;
                }
                if ((ClickPoint == "2") && (txtAttachFile.Tag == null))
                {
                    MessageBox.Show("파일이 없습니다.");
                    return;
                }

                var ViewReceiver = dgdInspect.SelectedItem as Win_Qul_InspectAutoRepair_U_View;
                if (ViewReceiver != null)
                {
                    if (ClickPoint == "1")
                    {
                        FTP_DownLoadFile(ViewReceiver.SketchPath, ViewReceiver.InspectID, ViewReceiver.SketchFile);
                    }
                    else if (ClickPoint == "2")
                    {
                        FTP_DownLoadFile(ViewReceiver.AttachedPath, ViewReceiver.InspectID, ViewReceiver.AttachedFile);
                    }
                }
            }
        }

        //다운로드
        private void FTP_DownLoadFile(string Path, string FolderName, string ImageName)
        {
            try
            {
                _ftp = new FTP_EX(FTP_ADDRESS, FTP_ID, FTP_PASS);

                string[] fileListSimple;
                string[] fileListDetail;

                fileListSimple = _ftp.directoryListSimple("", Encoding.UTF8);

                bool ExistFile = false;

                ExistFile = FolderInfoAndFlag(fileListSimple, FolderName);

                if (ExistFile)
                {
                    ExistFile = false;
                    fileListDetail = _ftp.directoryListSimple(FolderName, Encoding.UTF8);

                    ExistFile = FileInfoAndFlag(fileListDetail, ImageName);

                    if (ExistFile)
                    {
                        string str_remotepath = string.Empty;
                        string str_localpath = string.Empty;

                        str_remotepath = FTP_ADDRESS + '/' + FolderName + '/' + ImageName;
                        str_localpath = LOCAL_DOWN_PATH + "\\" + ImageName;

                        DirectoryInfo DI = new DirectoryInfo(LOCAL_DOWN_PATH);
                        if (DI.Exists)
                        {
                            DI.Create();
                        }

                        FileInfo file = new FileInfo(str_localpath);
                        if (file.Exists)
                        {
                            file.Delete();
                        }

                        _ftp.download(str_remotepath.Substring(str_remotepath.Substring
                            (0, str_remotepath.LastIndexOf("/")).LastIndexOf("/")), str_localpath);

                        ProcessStartInfo proc = new ProcessStartInfo(str_localpath);
                        proc.UseShellExecute = true;
                        Process.Start(proc);
                    }
                    else
                    {
                        MessageBox.Show("파일을 찾을 수 없습니다.");
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


        /// <summary>
        /// 해당영역에 폴더가 있는지 확인
        /// </summary>
        bool FolderInfoAndFlag(string[] strFolderList, string FolderName)
        {
            bool flag = false;
            foreach (string FolderList in strFolderList)
            {
                if (FolderList == FolderName)
                {
                    flag = true;
                    break;
                }
            }
            return flag;
        }

        /// <summary>
        /// 해당영역에 파일 있는지 확인
        /// </summary>
        bool FileInfoAndFlag(string[] strFileList, string FileName)
        {
            bool flag = false;
            foreach (string FileList in strFileList)
            {
                if (FileList == FileName)
                {
                    flag = true;
                    break;
                }
            }
            return flag;
        }






        #endregion


        #region 텍스트박스 엔터 키 이벤트

        private void cboDefectReason_DropDownClosed(object sender, EventArgs e)
        {
            lib.SendK(Key.Tab, this);
        }

        // 엔터 키를 통한 탭 인덱스 키 이동.
        private void EnterMove_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                lib.SendK(Key.Tab, this);
            }
        }



        #endregion

        #region 기타 메소드
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

        #endregion

        private void txtCorrContents_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                txtCorrContents.Focus();
            }
        }

        private void txtDefectRespectContents_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                txtDefectRespectContents.Focus();
            }
        }

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

        private void rbnChoiceAll_Click(object sender, RoutedEventArgs e)
        {
            rbnChoiceING.IsChecked = false;
            rbnChoiceFinished.IsChecked = false;
            rbnChoiceAll.IsChecked = true;
        }

        private void rbnChoiceING_Click(object sender, RoutedEventArgs e)
        {
            rbnChoiceING.IsChecked = true;
            rbnChoiceFinished.IsChecked = false;
            rbnChoiceAll.IsChecked = false;
        }

        private void rbnChoiceFinished_Click(object sender, RoutedEventArgs e)
        {
            rbnChoiceING.IsChecked = false;
            rbnChoiceFinished.IsChecked = true;
            rbnChoiceAll.IsChecked = false;
        }

        private void chkBuyerArticleNo_Click(object sender, MouseButtonEventArgs e)
        {
            if(chkBuyerArticleNo.IsChecked == true)
            {
                chkBuyerArticleNo.IsChecked = false;
                txtBuyerArticleNo.IsEnabled = false;
                btnBuyerArticleNo.IsEnabled = false;
            }
            else
            {
                chkBuyerArticleNo.IsChecked = true;
                txtBuyerArticleNo.IsEnabled = true;
                btnBuyerArticleNo.IsEnabled = true;
            }
        }

        private void chkBuyerArticleNo_Click(object sender, RoutedEventArgs e)
        {
            if (chkBuyerArticleNo.IsChecked == true)
            {
                chkBuyerArticleNo.IsChecked = true;
                txtBuyerArticleNo.IsEnabled = true;
                btnBuyerArticleNo.IsEnabled = true;
            }
            else
            {
                chkBuyerArticleNo.IsChecked = false;
                txtBuyerArticleNo.IsEnabled = false;
                btnBuyerArticleNo.IsEnabled = false;
            }
        }

        private void TxtBuyerArticleNo_KeyDown(object sender, KeyEventArgs e)
        {
            if(e.Key == Key.Enter)
            {
                pf.ReturnCode(txtBuyerArticleNo, 76, txtBuyerArticleNo.Text);
            }
        }

        private void btnBuyerArticleNo_Click(object sender, RoutedEventArgs e)
        {
            pf.ReturnCode(txtBuyerArticleNo, 76, txtBuyerArticleNo.Text);

        }
    }

    class Win_Qul_InspectAutoRepair_U_View : BaseView
    {
        public override string ToString()
        {
            return (this.ReportAllProperties());
        }

        public ObservableCollection<CodeView> cboTrade { get; set; }

        // Inspect 조회 값.    
        public string InspectID { get; set; }
        public string InspectSeq { get; set; }
        public string ArticleID { get; set; }
        public string Article { get; set; }
        public string Sabun { get; set; }

        public string InspectGubun { get; set; }

        public string InspectDate { get; set; }
        public string Full_InspectDate { get; set; }

        public string LotID { get; set; }
        public string InspectQty { get; set; }
        public string ECONo { get; set; }
        public string Comments { get; set; }

        public string InspectLevel { get; set; }
        public string SketchPath { get; set; }
        public string SketchFile { get; set; }
        public string AttachedPath { get; set; }
        public string AttachedFile { get; set; }

        public string InspectUserID { get; set; }
        public string InspectBasisID { get; set; }
        public string ProcessID { get; set; }
        public string DefectYN { get; set; }
        public string Process { get; set; }

        public string BuyerArticleNo { get; set; }
        public string InspectPoint { get; set; }
        public string ImportSecYN { get; set; }
        public string ImportlawYN { get; set; }
        public string ImportImpYN { get; set; }

        public string ImportNorYN { get; set; }
        public string IRELevel { get; set; }
        public string IRELevelName { get; set; }
        public string InpCustomID { get; set; }
        public string InpCustomName { get; set; }

        public string InpDate { get; set; }
        public string OutCustomID { get; set; }
        public string OutCustomName { get; set; }
        public string OutDate { get; set; }
        public string MachineID { get; set; }

        public string BuyerModelID { get; set; }
        public string BuyerModel { get; set; }
        public string FMLGubun { get; set; }
        public string insItemName { get; set; }
        public string InspectSpec { get; set; }

        public string InspectValue { get; set; }
        public string RepairYN { get; set; }
        public string DefectReasonCode { get; set; }
        public string ReasonImput { get; set; }
        public string CorrContents { get; set; }

        public string DefectRespectContents { get; set; }

    }


}

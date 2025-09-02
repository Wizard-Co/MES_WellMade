using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Input;
using System.Windows.Media;
using WizMes_WellMade.PopUp;
using WizMes_WellMade.PopUP;
using Excel = Microsoft.Office.Interop.Excel;

/**************************************************************************************************
'** 프로그램명 : Win_ord_Order_U
'** 설명       : 수주등록
'** 작성일자   : 2023.04.03
'** 작성자     : 장시영
'**------------------------------------------------------------------------------------------------
'**************************************************************************************************
' 변경일자  , 변경자, 요청자    , 요구사항ID      , 요청 및 작업내용
'**************************************************************************************************
' 2023.04.03, 장시영, 저장시 xp_Order_dOrderColorAll 내용 삭제 - xp_Order_uOrder 에서 동작하도록 수정
' 2024.08.23, 다중 선택 삭제 지원, 작지, 생산, 출하 있을 시 삭제 불가 조건 추가
'**************************************************************************************************/

namespace WizMes_WellMade
{
    /// <summary>
    /// Win_ord_Order_U.xaml에 대한 상호 작용 논리
    /// </summary>
    public partial class Win_ord_Order_U : UserControl
    {
        string stDate = string.Empty;
        string stTime = string.Empty;

        Lib lib = new Lib();
        PlusFinder pf = new PlusFinder();

        string strFlag = string.Empty;
        int rowNum = 0;

        Win_ord_Order_U_CodeView OrderView = new Win_ord_Order_U_CodeView();
        private List<Win_ord_Order_U_CodeView> OrderView_Del = new List<Win_ord_Order_U_CodeView>();

        ArticleData articleData = new ArticleData();
        string PrimaryKey = string.Empty;

        //FTP 활용모음
        string strImagePath = string.Empty;
        string strFullPath = string.Empty;
        string strDelFileName = string.Empty;

        List<string[]> deleteListFtpFile = new List<string[]>(); // 삭제할 파일 리스트
        List<string[]> lstExistFtpFile = new List<string[]>();

        List<string> OrderID_List = new List<string>();

        // 촤! FTP Server 에 있는 폴더 + 파일 경로를 저장해놓고 그걸로 다운 및 업로드하자 마!
        // 이미지 이름 : 폴더이름
        Dictionary<string, string> lstFtpFilePath = new Dictionary<string, string>();

        private FTP_EX _ftp = null;
        string SketchPath = null;


        List<string[]> listFtpFile = new List<string[]>();
        private List<UploadFileInfo> _listFileInfo = new List<UploadFileInfo>();

        internal struct UploadFileInfo          //FTP.
        {
            public string Filename { get; set; }
            public FtpFileType Type { get; set; }
            public DateTime LastModifiedTime { get; set; }
            public long Size { get; set; }
            public string Filepath { get; set; }
        }

        internal enum FtpFileType
        {
            None,
            DIR,
            File
        }

        //string FTP_ADDRESS = "ftp://wizis.iptime.org/ImageData/InspectAuto";

        string FTP_ADDRESS = "ftp://" + LoadINI.FileSvr + ":" + LoadINI.FTPPort + LoadINI.FtpImagePath + "/Order";
        string ForderName = "Order";
        private const string FTP_ID = "wizuser";
        private const string FTP_PASS = "wiz9999";
        private const string LOCAL_DOWN_PATH = "C:\\Temp";

        ////string FTP_ADDRESS = "ftp://222.104.222.145:25000/ImageData/McRegularInspect";
        //string FTP_ADDRESS = "ftp://192.168.0.95/ImageData/McRegularInspect";


        public Win_ord_Order_U()
        {
            InitializeComponent();
        }

        private void UserControl_Loaded(object sender, RoutedEventArgs e)
        {
            stDate = DateTime.Now.ToString("yyyyMMdd");
            stTime = DateTime.Now.ToString("HHmm");

            DataStore.Instance.InsertLogByFormS(this.GetType().Name, stDate, stTime, "S");

            Lib.Instance.UiLoading(sender);
            btnToday_Click(null, null);
            SetComboBox();

            if (MainWindow.tempContent != null
                && MainWindow.tempContent.Count > 0)
            {
                string OrderId = MainWindow.tempContent[0];
                string sDate = MainWindow.tempContent[1];
                string eDate = MainWindow.tempContent[2];
                string chkYN = MainWindow.tempContent[3];


                if (chkYN.Equals("Y"))
                {
                    ChkDateSrh.IsChecked = true;
                }
                else
                {
                    ChkDateSrh.IsChecked = false;
                }

                dtpSDate.SelectedDate = DateTime.Parse(sDate.Substring(0, 4) + "-" + sDate.Substring(4, 2) + "-" + sDate.Substring(6, 2));
                dtpEDate.SelectedDate = DateTime.Parse(eDate.Substring(0, 4) + "-" + eDate.Substring(4, 2) + "-" + eDate.Substring(6, 2));         

                rowNum = 0;
                re_Search(rowNum);

                MainWindow.tempContent.Clear();
            }
        }

        //콤보박스 만들기
        private void SetComboBox()
        {
            // 가공 구분
            ObservableCollection<CodeView> ovcWork = ComboBoxUtil.Instance.GetCode_SetComboBox("Work", null);
            cboWork.ItemsSource = ovcWork;
            cboWork.DisplayMemberPath = "code_name";
            cboWork.SelectedValuePath = "code_id";

            // 주문 형태 // 1차벤더(1), 2차벤더(2), 3차벤더(4), Direct(5)
            ObservableCollection<CodeView> oveOrderForm = ComboBoxUtil.Instance.Gf_DB_CM_GetComCodeDataset(null, "ORDFRM", "Y", "", ""); 
            cboOrderForm.ItemsSource = oveOrderForm;
            cboOrderForm.DisplayMemberPath = "code_name";
            cboOrderForm.SelectedValuePath = "code_id";

            // 주문 구분 // 초도품(1), REPEAT(2), 시작품(3), 개발품(4)
            ObservableCollection<CodeView> ovcOrderClss = ComboBoxUtil.Instance.Gf_DB_CM_GetComCodeDataset(null, "ORDGBN", "Y", "", "");
            cboOrderClss.ItemsSource = ovcOrderClss;
            cboOrderClss.DisplayMemberPath = "code_name";
            cboOrderClss.SelectedValuePath = "code_id";

            // 주문 구분 (검색)
            cboOrderClassSrh.ItemsSource = ovcOrderClss;
            cboOrderClassSrh.DisplayMemberPath = "code_name";
            cboOrderClassSrh.SelectedValuePath = "code_id";
            cboOrderClassSrh.SelectedIndex = 0;

            // 주문 기준
     
            ObservableCollection<CodeView> ovcWorkUnitClss = ComboBoxUtil.Instance.Gf_DB_CM_GetComCodeDataset(null, "MTRUNIT", "Y", "", "");
            cboUnitClss.ItemsSource = ovcWorkUnitClss;
            cboUnitClss.DisplayMemberPath = "code_name";
            cboUnitClss.SelectedValuePath = "code_id";


            // 품명 종류
            ObservableCollection<CodeView> ovcArticleGrpID = ComboBoxUtil.Instance.GetArticleCode_SetComboBox("", 0);
            cboArticleGroup.ItemsSource = ovcArticleGrpID;
            cboArticleGroup.DisplayMemberPath = "code_name";
            cboArticleGroup.SelectedValuePath = "code_id";

            List<string[]> strAutoProductInspectYN = new List<string[]>();
            string[] strNo = { "N", "N" };
            string[] strYes = { "Y", "Y" };
            strAutoProductInspectYN.Add(strNo);
            strAutoProductInspectYN.Add(strYes);

            ObservableCollection<CodeView> ovcAutProductInspectYN = ComboBoxUtil.Instance.Direct_SetComboBox(strAutoProductInspectYN);
            cboAutoInspect.ItemsSource = ovcAutProductInspectYN;
            this.cboAutoInspect.DisplayMemberPath = "code_name";
            this.cboAutoInspect.SelectedValuePath = "code_id";
            this.cboAutoInspect.SelectedIndex = 0;

            List<string[]> strCloseClss = new List<string[]>();
            string[] strNotClose = { "", "진행중" };
            string[] strClose = { "1", "마감" };
            strCloseClss.Add(strNotClose);
            strCloseClss.Add(strClose);

            ObservableCollection<CodeView> ovcCloseClss = ComboBoxUtil.Instance.Direct_SetComboBox(strCloseClss);
            cboCloseClss.ItemsSource = ovcCloseClss;
            this.cboCloseClss.DisplayMemberPath = "code_name";
            this.cboCloseClss.SelectedValuePath = "code_id";
            this.cboCloseClss.SelectedIndex = 0;

            // 부가세 별도
            /*List<string> strVAT_Value = new List<string>();
            strVAT_Value.Add("Y");
            strVAT_Value.Add("N");
            strVAT_Value.Add("0");

            ObservableCollection<CodeView> cboVAT_YN = ComboBoxUtil.Instance.Direct_SetComboBox(strVAT_Value);
            cboVAT_YN.ItemsSource = cboVAT_YN;
            cboVAT_YN.DisplayMemberPath = "code_name";
            cboVAT_YN.SelectedValuePath = "code_name";*/

            //List<string[]> strArray = new List<string[]>();
            //string[] strOne = { "", "진행" };
            //string[] strTwo = { "1", "완료" };
            //strArray.Add(strOne);
            //strArray.Add(strTwo);

            //// 완료 구분
            //ObservableCollection<CodeView> ovcCloseClssSrh = ComboBoxUtil.Instance.Direct_SetComboBox(strArray);
            //cboCloseClssSrh.ItemsSource = ovcCloseClssSrh;
            //cboCloseClssSrh.DisplayMemberPath = "code_name";
            //cboCloseClssSrh.SelectedValuePath = "code_id";
            //cboCloseClssSrh.SelectedIndex = 0;
            //수주구분 영업(0), 생산(1), 시가공(2), 샘플(3)
            ObservableCollection<CodeView> oveOrderFlag = ComboBoxUtil.Instance.Gf_DB_CM_GetComCodeDataset(null, "ORDFLG", "Y", "", "");
            //영업, 생산오더만 보여주기 위해.            
            oveOrderFlag.RemoveAt(2);
            //카운트 4에서 하나 지우고 나면 카운트 3돼서 또 2번 지움
            oveOrderFlag.RemoveAt(2);

            //cboOrderFlag.ItemsSource = oveOrderFlag;
            //cboOrderFlag.DisplayMemberPath = "code_name";
            //cboOrderFlag.SelectedValuePath = "code_id";
            //cboOrderFlag.SelectedIndex = 1;

            cboOrderNO.ItemsSource = oveOrderFlag;
            cboOrderNO.DisplayMemberPath = "code_name";
            cboOrderNO.SelectedValuePath = "code_id";
            cboOrderNO.SelectedIndex = 1;
        }

        #region 체크박스 연동동작(상단)

        //수주일자
        private void lblDateSrh_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            ChkDateSrh.IsChecked = ChkDateSrh.IsChecked == true ? false : true;
        }

        //수주일자
        private void ChkDateSrh_Checked(object sender, RoutedEventArgs e)
        {
            if (dtpSDate != null && dtpEDate != null)
            {
                dtpSDate.IsEnabled = true;
                dtpEDate.IsEnabled = true;
            }
        }

        //수주일자
        private void ChkDateSrh_Unchecked(object sender, RoutedEventArgs e)
        {
            dtpSDate.IsEnabled = false;
            dtpEDate.IsEnabled = false;
        }

        //금일
        private void btnToday_Click(object sender, RoutedEventArgs e)
        {
            dtpSDate.SelectedDate = DateTime.Today;
            dtpEDate.SelectedDate = DateTime.Today;
        }

        //금월
        private void btnThisMonth_Click(object sender, RoutedEventArgs e)
        {
            dtpSDate.SelectedDate = Lib.Instance.BringThisMonthDatetimeList()[0];
            dtpEDate.SelectedDate = Lib.Instance.BringThisMonthDatetimeList()[1];
        }


        //거래처
        private void txtCustomIDSrh_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
                MainWindow.pf.ReturnCode(txtCustomIDSrh, (int)Defind_CodeFind.DCF_CUSTOM, "");
        }

        //거래처
        private void btnCustomIDSrh_Click(object sender, RoutedEventArgs e)
        {
            MainWindow.pf.ReturnCode(txtCustomIDSrh, (int)Defind_CodeFind.DCF_CUSTOM, "");
        }   

     

        //검색조건 - 품번 텍스트박스 키다운 이벤트
        private void txtBuyerArticleNoSrh_KeyDown(object sender, KeyEventArgs e)
        {
            try
            {
                if (e.Key == Key.Enter)
                {
                    pf.ReturnCode(txtBuyerArticleNoSrh, 7071, txtBuyerArticleNoSrh.Text);
                }
            }
            catch (Exception ee)
            {
                MessageBox.Show("오류지점 - " + ee.ToString());
            }
        }

        //검색조건 - 품번 플러스파인더 버튼
        private void btnBuyerArticleNoSrh_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                pf.ReturnCode(txtBuyerArticleNoSrh, 7071, txtBuyerArticleNoSrh.Text);
            }
            catch (Exception ee)
            {
                MessageBox.Show("오류지점 - " + ee.ToString());
            }
        }


        //품명 텍스트박스 키다운 이벤트
        private void txtArticleIDSrh_KeyDown(object sender, KeyEventArgs e)
        {
            try
            {
                if (e.Key == Key.Enter)
                {
                    pf.ReturnCode(txtArticleIDSrh, 77, txtArticleIDSrh.Text);
                }
            }
            catch (Exception ee)
            {
                MessageBox.Show("오류지점 - " + ee.ToString());
            }
        }

        //품명 플러스파인더 버튼
        private void btnArticleIDSrh_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                pf.ReturnCode(txtArticleIDSrh, 77, txtArticleIDSrh.Text);
            }
            catch (Exception ee)
            {
                MessageBox.Show("오류지점 - " + ee.ToString());
            }
        }

       
        //주문구분
        private void lblOrderClassSrh_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            chkOrderClassSrh.IsChecked = chkOrderClassSrh.IsChecked == true ? false : true;
        }

        //주문구분
        private void chkOrderClassSrh_Checked(object sender, RoutedEventArgs e)
        {
            cboOrderClassSrh.IsEnabled = true;
        }

        //주문구분
        private void chkOrderClassSrh_Unchecked(object sender, RoutedEventArgs e)
        {
            cboOrderClassSrh.IsEnabled = false;
        }

        //모델 - 키다운
        private void txtBuyerModelID_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
                pf.ReturnCode(txtBuyerModelID, 88, "");
        }
        //모델 - 버튼클릭
        private void btnBuyerModelID_Click(object sender, RoutedEventArgs e)
        {
                pf.ReturnCode(txtBuyerModelID, 88, "");
        }

        #endregion

        #region 수주일괄등록복사

        //수주일괄등록복사
        private void btnMassEnrollment_Click(object sender, RoutedEventArgs e)
        {
            popPreviousOrder.IsOpen = true;
        }

        private void popPreviousOrder_Opened(object sender, EventArgs e)
        {
            dtpPreviousMonth.SelectedDate = DateTime.Today.AddMonths(-1);
            dtpThisMonth.SelectedDate = DateTime.Today;
        }

        private void btnPreOrderOK_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                sqlParameter.Add("ChkDate", 1);
                sqlParameter.Add("SDate", dtpThisMonth.SelectedDate.Value.ToString("yyyyMM") + "1");
                sqlParameter.Add("EDate", dtpThisMonth.SelectedDate.Value.ToString("yyyyMM") + "31");


                // 거래처
                sqlParameter.Add("ChkCustomID", 0);
                sqlParameter.Add("CustomID",  "");
                // 품번
                sqlParameter.Add("ChkBuyerArticleNo",  0);
                sqlParameter.Add("BuyerArticleNo",  "");
                // 품명
                sqlParameter.Add("ChkArticleID",  0);
                sqlParameter.Add("ArticleID",  "");

                // 주문구분
                sqlParameter.Add("ChkOrderClss",  0);
                sqlParameter.Add("OrderClss",  "");



                //sqlParameter.Add("ChkCustom", 0);
                //sqlParameter.Add("CustomID", "");
                //sqlParameter.Add("ChkInCustom", 0);
                //sqlParameter.Add("InCustomID", "");

                //sqlParameter.Add("ChkArticleID", 0);
                //sqlParameter.Add("ArticleID", "");
                //sqlParameter.Add("ChkArticle", 0);
                //sqlParameter.Add("Article", 0);

                //sqlParameter.Add("ChkOrderID", 0);
                //sqlParameter.Add("OrderID", "");
                //sqlParameter.Add("ChkCloseClss", "");
                //sqlParameter.Add("CloseClss", "");      

                //sqlParameter.Add("ChkOrderClss", 0);
                //sqlParameter.Add("OrderClss", "");
                //sqlParameter.Add("ChkOrderFlag", 0);
                //sqlParameter.Add("OrderFlag", "");

                DataSet ds = DataStore.Instance.ProcedureToDataSet("xp_ord_sOrder", sqlParameter, false);

                if (ds != null && ds.Tables.Count > 0)
                {
                    DataTable dt = ds.Tables[0];

                    if (dt.Rows.Count > 0)
                    {
                        if (MessageBox.Show(dtpThisMonth.SelectedDate.Value.ToString("yyyyMM") + " 월의 수주가 " + dt.Rows.Count.ToString() + " 건이 있습니다. " +
                            "무시하고 진행하시겠습니까?", "", MessageBoxButton.OKCancel) == MessageBoxResult.OK)
                        {
                            OrderCopy();
                        }
                    }
                    else
                    {
                        if (MessageBox.Show(dtpPreviousMonth.SelectedDate.Value.ToString("yyyyMM") + " 월의 수주가 " + dtpThisMonth.SelectedDate.Value.ToString("yyyyMM") + "월의 수주로 복사됩니다." +
                            "진행하시겠습니까?", "", MessageBoxButton.OKCancel) == MessageBoxResult.OK)
                        {
                            OrderCopy();
                        }
                    }
                }
                else
                {
                    if (MessageBox.Show(dtpPreviousMonth.SelectedDate.Value.ToString("yyyyMM") + " 월의 수주가 " + dtpThisMonth.SelectedDate.Value.ToString("yyyyMM") + "월의 수주로 복사됩니다." +
                        "진행하시겠습니까?", "", MessageBoxButton.OKCancel) == MessageBoxResult.OK)
                    {
                        OrderCopy();
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

            popPreviousOrder.IsOpen = false;
        }

        private void btnPreOrderCC_Click(object sender, RoutedEventArgs e)
        {
            popPreviousOrder.IsOpen = false;
        }

        private void OrderCopy()
        {
            bool Inresult = false;
            List<Procedure> Prolist = new List<Procedure>();
            List<Dictionary<string, object>> ListParameter = new List<Dictionary<string, object>>();
            try
            {
                Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                sqlParameter.Add("FromYYYYMM", dtpPreviousMonth.SelectedDate.Value.ToString("yyyyMM"));
                sqlParameter.Add("ToYYYYMM", dtpThisMonth.SelectedDate.Value.ToString("yyyyMM"));  //후에 Tag.Text 로 바꿔야 한다
                sqlParameter.Add("CreateUserID", MainWindow.CurrentUser);

                Procedure pro1 = new Procedure();
                pro1.Name = "xp_Order_iOrderCopy";
                pro1.OutputUseYN = "N";
                pro1.OutputName = "OrderID";
                pro1.OutputLength = "10";

                Prolist.Add(pro1);
                ListParameter.Add(sqlParameter);

                string[] Confirm = new string[2];
                Confirm = DataStore.Instance.ExecuteAllProcedureOutputNew(Prolist, ListParameter);
                if (Confirm[0] != "success")
                {
                    MessageBox.Show("[저장실패]\r\n" + Confirm[1].ToString());
                }
                else
                {
                    MessageBox.Show("수주 복사가 완료 되었습니다.");
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

        /// <summary>
        /// 수정,추가 저장 후
        /// </summary>
        private void CanBtnControl()
        {
            btnAdd.IsEnabled = true;
            btnUpdate.IsEnabled = true;
            btnDelete.IsEnabled = true;
            btnSearch.IsEnabled = true;
            btnSave.Visibility = Visibility.Hidden;
            btnCancel.Visibility = Visibility.Hidden;
            btnExcel.Visibility = Visibility.Visible;
            btnUpload.IsEnabled = true;

            grdInput.IsHitTestVisible = false;
            lblMsg.Visibility = Visibility.Hidden;
            dgdMain.IsHitTestVisible = true;
        }

        /// <summary>
        /// 수정,추가 진행 중
        /// </summary>
        private void CantBtnControl()
        {
            btnAdd.IsEnabled = false;
            btnUpdate.IsEnabled = false;
            btnDelete.IsEnabled = false;
            btnSearch.IsEnabled = false;
            btnSave.Visibility = Visibility.Visible;
            btnCancel.Visibility = Visibility.Visible;
            btnExcel.Visibility = Visibility.Hidden;
            btnUpload.IsEnabled = false;

            grdInput.IsHitTestVisible = true;
            lblMsg.Visibility = Visibility.Visible;
            dgdMain.IsHitTestVisible = false;
        }

        //추가
        private void btnAdd_Click(object sender, RoutedEventArgs e)
        {
            strFlag = "I";
            //this.DataContext = null;
            OrderView = new Win_ord_Order_U_CodeView();
            this.DataContext = new object();

            txtOrderQty.Text = "0";

            //혹시 모르니까 납기일자의 체크박스가 체크되어 있을 수도 있으니까 해제
            chkDvlyDate.IsChecked = false;

            CantBtnControl();

            cboOrderNO.SelectedIndex = 1;
            cboOrderForm.SelectedIndex = 1;
            cboArticleGroup.SelectedIndex = 0;
            cboOrderClss.SelectedIndex = 0;
            cboUnitClss.SelectedIndex = 0;
            cboWork.SelectedIndex = 0;
            cboCloseClss.SelectedIndex = 0;
            cboAutoInspect.SelectedIndex = 0;

            dtpAcptDate.SelectedDate = DateTime.Today;
            dtpDvlyDate.SelectedDate = DateTime.Today;
            btnNeedStuff.IsEnabled = true;
            tbkMsg.Text = "자료 입력 중";
            rowNum = Math.Max(0, dgdMain.SelectedIndex);
            
            if (dgdNeedStuff.Items.Count > 0)
                dgdNeedStuff.Items.Clear();
        }

        //수정
        private void btnUpdate_Click(object sender, RoutedEventArgs e)
        {
            OrderView = dgdMain.SelectedItem as Win_ord_Order_U_CodeView;
            //MessageBox.Show(txtCustom.Tag.ToString());

            if (OrderView != null)
            {
                //rowNum = dgdMain.SelectedIndex;
                dgdMain.IsHitTestVisible = false;
                btnNeedStuff.IsEnabled = true;
                tbkMsg.Text = "자료 수정 중";
                strFlag = "U";
                CantBtnControl();
                PrimaryKey = OrderView.OrderID;
            }
        }

        //삭제
        private void btnDelete_Click(object sender, RoutedEventArgs e)
        {
            if(OrderID_List.Count > 0)
            {
                MessageBoxResult msgresult = MessageBox.Show("선택하신 수주를 삭제 하시겠습니까?", "삭제 확인", MessageBoxButton.YesNo);
                if (msgresult == MessageBoxResult.Yes)
                {
                    using (Loading ld = new Loading(beDelete))
                    {
                        ld.ShowDialog();
                    }
                }
            }
            else
            {
                MessageBox.Show("삭제하실 데이터를 선택해주세요.", "확인");
            }
         
        }



        //다중 셀렉트 삭제
        private void beDelete()
        {
            btnDelete.IsEnabled = false;

            bool flag = true;
            string msg = string.Empty;

            Dispatcher.BeginInvoke(new Action(() =>
            {
                foreach(string orderID in OrderID_List)
                {
                    flag = CheckFKkey(orderID, out msg);
                    if (msg != string.Empty && flag != true)
                    {
                        msg += " 삭제할 수 없습니다.";
                        MessageBox.Show(msg,"확인");
                        return;
                    }                       
                }

                foreach (string orderID in OrderID_List)
                {
                    flag = DeleteData(orderID);
                }

                if (flag)
                {
                    MessageBox.Show($"선택하신 수주 {OrderID_List.Count}건이 삭제 되었습니다.");
                    re_Search(0);
                }

            }), System.Windows.Threading.DispatcherPriority.Background);

            btnDelete.IsEnabled = true;
        }


        //닫기
        private void btnClose_Click(object sender, RoutedEventArgs e)
        {
            DataStore.Instance.InsertLogByFormS(this.GetType().Name, stDate, stTime, "E");
            Lib.Instance.ChildMenuClose(this.ToString());
        }

        //검색
        private void btnSearch_Click(object sender, RoutedEventArgs e)
        {
            if(lib.DatePickerCheck(dtpSDate, dtpEDate, ChkDateSrh))
            {
                using (Loading ld = new Loading(beSearch))
                {
                    ld.ShowDialog();
                }
            }        
        }

        private void beSearch()
        {
            rowNum = 0;
            btnSearch.IsEnabled = false;

            Dispatcher.BeginInvoke(new Action(() =>
            {
                //로직
                re_Search(rowNum);
            }), System.Windows.Threading.DispatcherPriority.Background);

            btnSearch.IsEnabled = true;
        }

        //저장
        private void btnSave_Click(object sender, RoutedEventArgs e)
        {
            using (Loading ld = new Loading(beSave))
            {
                ld.ShowDialog();
            }
        }

        private void beSave()
        {
            btnSave.IsEnabled = false;

            Dispatcher.BeginInvoke(new Action(() =>
            {
                //로직
                if (SaveData(strFlag))
                {
                    CanBtnControl();
                    lblMsg.Visibility = Visibility.Hidden;
                    dgdMain.IsHitTestVisible = true;
                    btnNeedStuff.IsEnabled = false;
                    re_Search(rowNum);
                    PrimaryKey = string.Empty;
                    rowNum = 0;
                    MessageBox.Show("저장이 완료되었습니다.");
                }
            }), System.Windows.Threading.DispatcherPriority.Background);

            btnSave.IsEnabled = true;
        }

        //취소
        private void btnCancel_Click(object sender, RoutedEventArgs e)
        {
            CanBtnControl();        

            //혹시 모르니까 납기일자의 체크박스가 체크되어 있을 수도 있으니까 해제
            chkDvlyDate.IsChecked = false;

            dgdMain.IsHitTestVisible = true;
            btnNeedStuff.IsEnabled = false;

            if (strFlag.Equals("U"))
            {
                re_Search(rowNum);
            }
            else
            {
         
                rowNum = 0; 
                re_Search(rowNum);
            }

            strFlag = string.Empty;

        }

        //엑셀
        private void btnExcel_Click(object sender, RoutedEventArgs e)
        {
            DataTable dt = null;
            string Name = string.Empty;
            Lib lib = new Lib();

            string[] lst = new string[2];
            lst[0] = "수주 조회 목록";
            lst[1] = dgdMain.Name;

            ExportExcelxaml ExpExc = new ExportExcelxaml(lst);

            ExpExc.ShowDialog();

            if (ExpExc.DialogResult.HasValue)
            {
                if (ExpExc.choice.Equals(dgdMain.Name))
                {
                    DataStore.Instance.InsertLogByForm(this.GetType().Name, "E");
                    if (ExpExc.Check.Equals("Y"))
                        dt = lib.DataGridToDTinHidden(dgdMain);
                    else
                        dt = lib.DataGirdToDataTable(dgdMain);

                    Name = dgdMain.Name;

                    if (lib.GenerateExcel(dt, Name))
                    {
                        lib.excel.Visible = true;
                        lib.ReleaseExcelObject(lib.excel);
                    }
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
            lib = null;
        }

        // 주문일괄 업로드
        string upload_fileName = "";

        private void btnUpload_Click(object sender, RoutedEventArgs e)
        {
            Microsoft.Win32.OpenFileDialog file = new Microsoft.Win32.OpenFileDialog();
            file.Filter = "Excel files (*.xls,*xlsx)|*.xls;*xlsx|All files (*.*)|*.*";
            file.InitialDirectory = "C:\\";

            if (file.ShowDialog() == true)
            {
                upload_fileName = file.FileName;

                btnUpload.IsEnabled = false;

                using (Loading ld = new Loading("excel", beUpload))
                {
                    ld.ShowDialog();
                }

                re_Search(0);

                btnUpload.IsEnabled = true;
            }
        }

        private void beUpload()
        {
            Lib lib2 = new Lib();

            Excel.Application excelapp = null;
            Excel.Workbook workbook = null;
            Excel.Worksheet worksheet = null;
            Excel.Range workrange = null;

            List<OrderExcel> listExcel = new List<OrderExcel>();

            try
            {
                excelapp = new Excel.Application();
                workbook = excelapp.Workbooks.Add(upload_fileName);
                worksheet = workbook.Sheets["Sheet"];
                workrange = worksheet.UsedRange;

                for (int row = 3; row <= workrange.Rows.Count; row++)
                {
                    OrderExcel excel = new OrderExcel();
                    excel.CustomID = workrange.get_Range("A" + row.ToString()).Value2;
                    excel.Model = workrange.get_Range("B" + row.ToString()).Value2;
                    excel.BuyerArticleNo = workrange.get_Range("C" + row.ToString()).Value2;
                    excel.Article = workrange.get_Range("D" + row.ToString()).Value2;
                    excel.UnitClss = workrange.get_Range("E" + row.ToString()).Value2;

                    object objOrderQty = workrange.get_Range("H" + row.ToString()).Value2;
                    if (objOrderQty != null)
                        excel.OrderQty = objOrderQty.ToString();

                    if (!string.IsNullOrEmpty(excel.CustomID)
                        && !string.IsNullOrEmpty(excel.BuyerArticleNo) && !string.IsNullOrEmpty(excel.Article)
                        && !string.IsNullOrEmpty(excel.UnitClss) && !string.IsNullOrEmpty(excel.OrderQty))
                    {
                        listExcel.Add(excel);
                    }

                    if (string.IsNullOrEmpty(excel.CustomID) && string.IsNullOrEmpty(excel.Model)
                        && string.IsNullOrEmpty(excel.BuyerArticleNo) && string.IsNullOrEmpty(excel.Article)
                        && string.IsNullOrEmpty(excel.UnitClss) && string.IsNullOrEmpty(excel.OrderQty))
                    {
                        break;
                    }
                }

                if (listExcel.Count > 0)
                {
                    List<Procedure> Prolist = new List<Procedure>();
                    List<Dictionary<string, object>> ListParameter = new List<Dictionary<string, object>>();
                    for (int i = 0; i < listExcel.Count; i++)
                    {
                        Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                        sqlParameter.Add("CustomID", string.IsNullOrEmpty(listExcel[i].CustomID) ? "" : listExcel[i].CustomID);
                        sqlParameter.Add("Model", string.IsNullOrEmpty(listExcel[i].Model) ? "" : listExcel[i].Model);
                        sqlParameter.Add("BuyerArticleNo", string.IsNullOrEmpty(listExcel[i].BuyerArticleNo) ? "" : listExcel[i].BuyerArticleNo);
                        sqlParameter.Add("Article", string.IsNullOrEmpty(listExcel[i].Article) ? "" : listExcel[i].Article);
                        sqlParameter.Add("UnitClss", string.IsNullOrEmpty(listExcel[i].UnitClss) ? "" : listExcel[i].UnitClss);
                        sqlParameter.Add("OrderQty", string.IsNullOrEmpty(listExcel[i].OrderQty) ? "" : listExcel[i].OrderQty);

                        Procedure pro2 = new Procedure();
                        pro2.Name = "xp_Order_iOrderExcel";
                        pro2.OutputUseYN = "N";
                        pro2.OutputName = "";
                        pro2.OutputLength = "10";

                        Prolist.Add(pro2);
                        ListParameter.Add(sqlParameter);
                    }

                    string[] Confirm = new string[2];
                    Confirm = DataStore.Instance.ExecuteAllProcedureOutputNew_NewLog(Prolist, ListParameter, "C");
                    if (Confirm[0] != "success")
                        MessageBox.Show("[저장실패]\r\n" + Confirm[1].ToString());
                    else
                        MessageBox.Show("업로드가 완료되었습니다.");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message.ToString());
            }
            finally
            {
                excelapp.Visible = true;
                workbook.Close(true);
                excelapp.Quit();

                lib2.ReleaseExcelObject(workbook);
                lib2.ReleaseExcelObject(worksheet);
                lib2.ReleaseExcelObject(excelapp);
                lib2 = null;

                upload_fileName = "";
                listExcel.Clear();
            }
        }

        private int SelectItem(string strPrimary, DataGrid dataGrid)
        {
            int index = 0;

            try
            {
                for (int i = 0; i < dataGrid.Items.Count; i++)
                {
                    var Item = dataGrid.Items[i] as Win_ord_Order_U_CodeView;

                    if (strPrimary.Equals(Item.OrderID))
                    {
                        index = i;
                        break;
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }

            return index;
        }

        private void re_Search(int selectedIndex)
        {
            //ClearGrdInput();

            FillGrid();

            if (dgdMain.Items.Count > 0)
            {
                dgdMain.SelectedIndex = PrimaryKey.Equals(string.Empty) ? 
                    selectedIndex : SelectItem(PrimaryKey, dgdMain);
            }
            else
                this.DataContext = new object();

        }

        //실조회
        private void FillGrid()
        {
            dgdMain.Items.Clear();
            dgdTotal.Items.Clear();

            try
            {
                Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                sqlParameter.Clear();
                sqlParameter.Add("ChkDate", ChkDateSrh.IsChecked == true ? 1 : 0);
                sqlParameter.Add("SDate", ChkDateSrh.IsChecked == true ? dtpSDate.SelectedDate.Value.ToString("yyyyMMdd") : "");
                sqlParameter.Add("EDate", ChkDateSrh.IsChecked == true ? dtpEDate.SelectedDate.Value.ToString("yyyyMMdd") : "");

                // 거래처
                sqlParameter.Add("ChkCustomID", chkCustomIDSrh.IsChecked == true ? 1 : 0);
                sqlParameter.Add("CustomID", chkCustomIDSrh.IsChecked == true ? txtCustomIDSrh.Tag != null ? txtCustomIDSrh.Tag.ToString() : "" : "");
                // 품번
                sqlParameter.Add("ChkBuyerArticleNo", chkBuyerArticleNoSrh.IsChecked == true ? 1 : 0);
                sqlParameter.Add("BuyerArticleNo", chkBuyerArticleNoSrh.IsChecked == true ? txtBuyerArticleNoSrh.Tag != null? txtBuyerArticleNoSrh.Tag.ToString() : "": "");
                // 품명
                sqlParameter.Add("ChkArticleID", chkArticleIDSrh.IsChecked == true ? 1 : 0);
                sqlParameter.Add("ArticleID", chkArticleIDSrh.IsChecked == true ? txtArticleIDSrh.Tag != null? txtArticleIDSrh.Tag.ToString() : "" :""); 

                // 주문구분
                sqlParameter.Add("ChkOrderClss", chkOrderClassSrh.IsChecked == true ? 1 : 0);
                sqlParameter.Add("OrderClss", chkOrderClassSrh.IsChecked == true ? cboOrderClassSrh.SelectedValue != null? cboOrderClassSrh.SelectedValue.ToString() : "": "");  

            

                DataSet ds = DataStore.Instance.ProcedureToDataSet_LogWrite("xp_ord_sOrder", sqlParameter, true, "R");

                if (ds != null && ds.Tables.Count > 0)
                {
                    DataTable dt = ds.Tables[0];

                    if (dt.Rows.Count == 0)
                    {
                        dgdNeedStuff.Items.Clear();
                        MessageBox.Show("조회된 데이터가 없습니다.");                    
                    }
                    else
                    {
                        int i = 0;
                        int OrderAmountSum = 0;
                        DataRowCollection drc = dt.Rows;

                        foreach (DataRow dr in drc)
                        {
                            i++;
                            var OrderCodeView = new Win_ord_Order_U_CodeView
                            {
                                Num = i,

                                OrderID = dr["OrderID"].ToString(),
                                OrderNo = dr["OrderNo"].ToString(),
                                OrderFlag = dr["OrderFlag"].ToString(),
                                AcptDate = DateTypeHyphen(dr["AcptDate"].ToString()),
                                DvlyDate = DateTypeHyphen(dr["DvlyDate"].ToString()),
                                CustomID = dr["CustomID"].ToString(),
                                KCustom = dr["KCustom"].ToString(),
                                InCustomID = dr["InCustomID"].ToString(),
                                InKCustom = dr["InKCustom"].ToString(),
                                DvlyPlace = dr["DvlyPlace"].ToString(),
                                ArticleID = dr["ArticleID"].ToString(),
                                Article = dr["Article"].ToString(),
                                BuyerArticleNo = dr["BuyerArticleNo"].ToString(),
                                BuyerModelID = dr["BuyerModelID"].ToString(),
                                BuyerModel = dr["BuyerModel"].ToString(),
                                OrderQty = stringFormatN0(dr["OrderQty"]),
                                UnitClss = dr["UnitClss"].ToString(),
                                OrderClss = dr["OrderClss"].ToString(),
                                CloseClss = dr["CloseClss"].ToString(),
                                WorkID = dr["WorkID"].ToString(),
                                UnitPrice = stringFormatN1(dr["UnitPrice"]),
                                OrderForm = dr["OrderForm"].ToString(),
                                Remark = dr["Remark"].ToString(),
                                ArticleGrpID = dr["ArticleGrpID"].ToString(),  
                                OrderAmount = stringFormatN0(dr["OrderAmount"])

                            };

                            OrderAmountSum += (int)RemoveComma(dr["OrderAmount"].ToString(), true);

                            dgdMain.Items.Add(OrderCodeView);
                        }

                        if(dgdMain.Items.Count > 0)
                        {
                            var Total = new Win_ord_Order_Total_U_CodeView
                            {
                                count = i,
                                OrderTotalAmount = stringFormatN0(OrderAmountSum)
                            };

                            dgdTotal.Items.Add(Total);
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

        private DataRow FillOneOrderData(string strOrderID)
        {
            DataRow dr = null;

            try
            {
                Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                sqlParameter.Add("OrderID", strOrderID);
                DataSet ds = DataStore.Instance.ProcedureToDataSet("xp_Order_sOrderOne", sqlParameter, false);

                if (ds != null && ds.Tables.Count > 0)
                {
                    DataTable dt = ds.Tables[0];

                    if (dt.Rows.Count > 0)
                    {
                        DataRowCollection drc = dt.Rows;
                        dr = drc[0];
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

            return dr;
        }

       
        /// <summary>
        /// 실삭제
        /// </summary>
        /// <param name="strID"></param>
        /// <returns></returns>
        private bool DeleteData(string strID)
        {
            bool flag = false;

            try
            {
                Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                sqlParameter.Clear();
                sqlParameter.Add("OrderID", strID);

                string[] result = DataStore.Instance.ExecuteProcedure_NewLog("xp_Order_dOrder", sqlParameter, "D");

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

        /// <summary>
        /// 실저장
        /// </summary>
        /// <param name="strID"></param>
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
                    Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                    sqlParameter.Clear();

                    sqlParameter.Add("OrderID", string.IsNullOrEmpty(txtOrderID.Text) ? "" : txtOrderID.Text);
                    sqlParameter.Add("CustomID", txtCustomID.Tag != null ? txtCustomID.Tag.ToString() : "");
                    sqlParameter.Add("OrderNO", string.IsNullOrEmpty(TextBoxOrderNo.Text) ? "" : TextBoxOrderNo.Text);
                    sqlParameter.Add("OrderForm", cboOrderForm.SelectedValue != null ? cboOrderForm.SelectedValue.ToString() : "");
                    sqlParameter.Add("OrderClss", cboOrderClss.SelectedValue != null ? cboOrderClss.SelectedValue.ToString() : "");

                    sqlParameter.Add("AcptDate", !IsDatePickerNull(dtpAcptDate) ? ConvertDate(dtpAcptDate) : "");
                    sqlParameter.Add("DvlyDate", !IsDatePickerNull(dtpDvlyDate) ? chkDvlyDate.IsChecked == true ? ConvertDate(dtpDvlyDate) : "" : "");
                    sqlParameter.Add("ArticleGrpID", cboArticleGroup.SelectedValue != null ? cboArticleGroup.SelectedValue.ToString() : "");
                    sqlParameter.Add("DvlyPlace", txtDvlyPlace.Text);
                    sqlParameter.Add("WorkID", cboWork.SelectedValue != null ? cboWork.SelectedValue.ToString() : "");

                    sqlParameter.Add("ExchRate", 0.00);
                    sqlParameter.Add("Vat_IND_YN", "Y");
                    sqlParameter.Add("OrderQty", RemoveComma(txtOrderQty.Text,true));
                    sqlParameter.Add("UnitClss", cboUnitClss.SelectedValue != null ? cboUnitClss.SelectedValue.ToString() : "");

                    sqlParameter.Add("Remark", txtComments.Text);
                    sqlParameter.Add("OrderFlag", cboOrderNO.SelectedValue != null ? cboOrderNO.SelectedValue.ToString() : ""); 
                    sqlParameter.Add("InCustomID", txtInCustomID.Tag != null ? txtInCustomID.Tag.ToString() : "");
                    sqlParameter.Add("UnitPriceClss", "0");
                    sqlParameter.Add("OrderSpec", "");

                    sqlParameter.Add("BuyerModelID", txtBuyerModelID.Tag != null ? txtBuyerModelID.Tag.ToString() : "");
                    sqlParameter.Add("ProductAutoInspectYN", cboAutoInspect.SelectedValue != null ? cboAutoInspect.SelectedValue.ToString() : "");
                    sqlParameter.Add("UnitPrice", RemoveComma(txtUnitPrice.Text, true, typeof(decimal)));
                    sqlParameter.Add("CloseClss", cboCloseClss.SelectedValue != null ? cboCloseClss.SelectedValue.ToString() : "");

                    string sGetID = strFlag.Equals("I") ? string.Empty : txtOrderID.Text;
                    #region 추가

                    if (strFlag.Equals("I"))
                    {
                        sqlParameter.Add("CreateUserID", MainWindow.CurrentUser);

                        Procedure pro1 = new Procedure();
                        pro1.Name = "xp_Order_iOrder";
                        pro1.OutputUseYN = "Y";
                        pro1.OutputName = "OrderID";
                        pro1.OutputLength = "10";

                        Prolist.Add(pro1);
                        ListParameter.Add(sqlParameter);

                        List<KeyValue> list_Result = new List<KeyValue>();
                        list_Result = DataStore.Instance.ExecuteAllProcedureOutputGetCS_NewLog(Prolist, ListParameter,"C");

                        if (list_Result[0].key.ToLower() == "success")
                        {
                            list_Result.RemoveAt(0);
                            for (int i = 0; i < list_Result.Count; i++)
                            {
                                KeyValue kv = list_Result[i];
                                if (kv.key == "OrderID")
                                {
                                    sGetID = kv.value;
                                    PrimaryKey = sGetID;
                                    flag = true;
                                }
                            }
                        }
                        else
                        {
                            MessageBox.Show("[저장실패]\r\n" + list_Result[1].value.ToString());
                            //flag = false;
                            return false;
                        }

                        Prolist.Clear();
                        ListParameter.Clear();
                    }
                    #endregion

                    #region 수정
                    else if (strFlag.Equals("U"))
                    {
                        sqlParameter.Add("LastUpdateUserID", MainWindow.CurrentUser);

                        Procedure pro3 = new Procedure();
                        pro3.Name = "xp_Order_uOrder";
                        pro3.OutputUseYN = "N";
                        pro3.OutputName = "OrderID";
                        pro3.OutputLength = "10";

                        Prolist.Add(pro3);
                        ListParameter.Add(sqlParameter);
                    }
                    #endregion


                    //OrderColor추가
                    //xp_Order_uOrder 프로시저에서 ordercolor를 삭제후 삽입 
                    sqlParameter = new Dictionary<string, object>();
                    sqlParameter.Clear();

                    sqlParameter.Add("OrderID", sGetID);
                    sqlParameter.Add("OrderSeq", 1);
                    sqlParameter.Add("ArticleID", txtBuyerArticleNo.Tag != null ? txtBuyerArticleNo.Tag.ToString() : "");
                    sqlParameter.Add("ArticleGrpID", cboArticleGroup.SelectedValue != null ? cboArticleGroup.SelectedValue.ToString() : "");
                    sqlParameter.Add("UnitPrice", ConvertDouble(txtUnitPrice.Text));
                    sqlParameter.Add("ColorQty", RemoveComma(txtOrderQty.Text,true));
                    sqlParameter.Add("NewProductYN", "");
                    sqlParameter.Add("UnitPriceClss", "1");
                    sqlParameter.Add("UnitClss", cboUnitClss.SelectedValue.ToString());

                    sqlParameter.Add("CreateUserID", MainWindow.CurrentUser);

                    Procedure pro2 = new Procedure();
                    pro2.Name = "xp_ord_iOrderSub";
                    pro2.OutputUseYN = "N";
                    pro2.OutputName = "OrderID";
                    pro2.OutputLength = "10";

                    Prolist.Add(pro2);
                    ListParameter.Add(sqlParameter);

                    string[] Confirm = new string[2];
                    Confirm = DataStore.Instance.ExecuteAllProcedureOutputNew_NewLog(Prolist, ListParameter,"U");
                    if (Confirm[0] != "success")
                    {
                        MessageBox.Show("[저장실패]\r\n" + Confirm[1].ToString());
                        flag = false;
                    }
                    else
                        flag = true;




                    //if (!PrimaryKey.Trim().Equals("") && !txtOrderID.Text.Trim().Equals(""))
                    //{
                    //    string PKKey = string.Empty;
                    //    if (strFlag.Trim().Equals("I")) PKKey = PrimaryKey;
                    //    else if (strFlag.Trim().Equals("U")) PKKey = txtOrderID.Text;

                    //    if (deleteListFtpFile.Count > 0)
                    //    {
                    //        foreach (string[] str in deleteListFtpFile)
                    //        {
                    //            FTP_RemoveFile(PKKey + "/" + str[0]);
                    //        }
                    //    }

                    //    if (listFtpFile.Count > 0)
                    //    {
                    //        FTP_Save_File(listFtpFile, PKKey);
                    //    }

                    //}

                    //// 파일 List 비워주기
                    //listFtpFile.Clear();
                    //lstFilesName.Clear();
                    //deleteListFtpFile.Clear();


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

        //점2
        private bool CheckData()
        {
            string msg = "";

            bool flag = true;

            if (txtCustomID.Text.Length <= 0 || txtCustomID.Tag == null)
                msg = "거래처가 입력되지 않았습니다. 먼저 거래처를 입력해주세요";
            else if (cboOrderForm.SelectedValue == null)
                msg = "주문형태가 선택되지 않았습니다. 먼저 주문형태를 선택해주세요";
            else if (cboOrderClss.SelectedValue == null)
                msg = "주문구분이 선택되지 않았습니다. 먼저 주문구분을 선택해주세요";
            else if (cboUnitClss.SelectedValue == null)
                msg = "주문기준이 선택되지 않았습니다. 먼저 주문기준을 선택해주세요";
            else if (cboArticleGroup.SelectedValue == null)
                msg = "품명종류가 선택되지 않았습니다. 먼저 품명종류를 선택해주세요";
            else if (string.IsNullOrEmpty(txtBuyerArticleNo.Text) || txtBuyerArticleNo.Tag == null)
                msg = "품번이 선택되지 않았습니다. 먼저 품번을 선택해주세요";
            else if (cboWork.SelectedValue == null)
                msg = "가공구분이 선택되지 않았습니다. 먼저 가공구분을 선택해주세요";
            else if (strFlag == "U" && txtOrderID.Text.Trim() != string.Empty && txtOrderID.Text != null)
            {
                flag = CheckFKkey(txtOrderID.Text, out msg);
                if(msg != string.Empty) msg += " 저장 할 수 없습니다.";
            }
    

            if (!string.IsNullOrEmpty(msg) ) 
            {
                if (!string.IsNullOrEmpty(msg))
                {
                    MessageBox.Show(msg,"확인");
                }         
                flag = false;
            }

            return flag;
        }

        private bool CheckFKkey(string orderID, out string msg)
        {
            bool flag = true;

            //가장 나중에 하는걸 역순으로 검사

            string[] sqlList = { "select orderid from outware where outclss= '01' AND orderid = '{0}' ",
                                 "select orderid from Inspect where orderid = '{0}' ",
                                  "select orderid from outware where outclss= '03' AND orderid = '{0}' ",
                                 "select orderid from pl_Input where orderid = '{0}' ",

            };

            string[] errMsg = {$"관리번호 {orderID}는 출하이력이 있으므로",                   //사무실정상출고 01
                               $"관리번호 {orderID}는 검사/포장 실적이 있으므로",             //검사포장하면 inspect에 들어간다
                               $"관리번호 {orderID}는 생산이력이 있으므로",                   //생산에하위품자동출고 03
                               $"관리번호 {orderID}는 작업지시가 내려져 있으므로",            //작업지시

                               
            };
            int errSeq = 0;
            msg = string.Empty;

            //반복문을 돌다가 걸리면 종료, 경고문 띄우고 false반환
            for (int i = 0; i < sqlList.Length; i++)
            {
                DataSet ds = DataStore.Instance.QueryToDataSet(string.Format(sqlList[i], orderID));
                if (ds != null && ds.Tables.Count > 0)
                {
                    DataTable dt = ds.Tables[0];
                    if (dt.Rows.Count > 0)
                    {
                        flag = false;
                        errSeq = i;
                        break;
                    }
                }
                else
                {
                    continue;
                }
            }

            if (flag == false)
            {
                msg = errMsg[errSeq];
            }

            return flag;
        }


        #region 입력시 Event

        private void TextBox_SetTagNull(object sender, TextChangedEventArgs e)
        {
            TextBox txtbox = sender as TextBox;
            if (txtbox.Text == string.Empty && txtbox.IsKeyboardFocused)
            {
                txtbox.Tag = null;
                if (txtbox.Name.Equals("txtCustomID"))
                {
                    txtDvlyPlace.Text = string.Empty;
                    txtInCustomID.Text = string.Empty;
                    txtInCustomID.Tag = null;
                }
             
            }
        }

        //거래처
        private void txtCustomID_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                MainWindow.pf.ReturnCode(txtCustomID, (int)Defind_CodeFind.DCF_CUSTOM, "");

                if (txtCustomID.Tag != null)
                {
                    //CallCustomData(txtCustom.Tag.ToString());
                    txtDvlyPlace.Text = txtCustomID.Text;
                    txtInCustomID.Text = txtCustomID.Text;
                    txtInCustomID.Tag = txtCustomID.Tag;
                }

                e.Handled = true;
            }
        }

        //거래처
        private void btnCustomID_Click(object sender, RoutedEventArgs e)
        {
            MainWindow.pf.ReturnCode(txtCustomID, (int)Defind_CodeFind.DCF_CUSTOM, "");

            if (txtCustomID.Tag != null)
            {
                //CallCustomData(txtCustom.Tag.ToString());
                txtDvlyPlace.Text = txtCustomID.Text;
                txtInCustomID.Text = txtCustomID.Text;
                txtInCustomID.Tag = txtCustomID.Tag;
            }
        }

        //품명
        private void txtArticle_KeyDown(object sender, KeyEventArgs e)
        {
            try
            {
                if (e.Key == Key.Enter)
                {

                   
                    MainWindow.pf.ReturnCodeGLS(txtArticleID, 77, "");                

                    if (txtArticleID.Tag != null)
                    {
                        CallArticleData(txtArticleID.Tag.ToString());

                        cboArticleGroup.SelectedValue = articleData.ArticleGrpID;
                        txtBuyerArticleNo.Text = articleData.BuyerArticleNo;
                        txtUnitPrice.Text = articleData.OutUnitPrice;


                    }

         
                }
            }
            catch (Exception ex)
            {
                //MessageBox.Show("오류 발생, 오류 내용 : " + ex.ToString());
            }
            finally
            {
                DataStore.Instance.CloseConnection();
            }
        }

        //품명
        private void btnArticleID_Click(object sender, RoutedEventArgs e)
        {
            try
            {
      
                MainWindow.pf.ReturnCodeGLS(txtArticleID, 77, "");            

                if (txtArticleID.Tag != null)
                {
                    CallArticleData(txtArticleID.Tag.ToString());
                    //품명종류 대입(ex.제품 등)
                    //cboArticleGroup.SelectedValue = articleData.ArticleGrpID;

                    //품번 대입
                    //txtBuyerArticleNO.Text = articleData.BuyerArticleNo;
                    //품명 대입
                    //txtBuyerArticleNO.Text = articleData.Article;
                    //단가 대입
                    txtUnitPrice.Text = articleData.OutUnitPrice;
                }


            }
            catch (Exception ex)
            {
                //MessageBox.Show("오류 발생, 오류 내용 : " + ex.ToString());
            }
            finally
            {
                DataStore.Instance.CloseConnection();
            }
        }


        private void CallArticleData(string strArticleID)
        {
            try
            {
                Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                sqlParameter.Add("ArticleID", strArticleID);

                DataSet ds = DataStore.Instance.ProcedureToDataSet("xp_Order_sArticleData", sqlParameter, false);

                if (ds != null && ds.Tables.Count > 0)
                {
                    DataTable dt = ds.Tables[0];

                    if (dt.Rows.Count > 0)
                    {
                        DataRow dr = dt.Rows[0];

                        articleData = new ArticleData
                        {
                            ArticleID = dr["ArticleID"].ToString(),
                            Article = dr["Article"].ToString(),
                            ThreadID = dr["ThreadID"].ToString(),
                            thread = dr["thread"].ToString(),
                            StuffWidth = dr["StuffWidth"].ToString(),
                            DyeingID = dr["DyeingID"].ToString(),
                            Weight = dr["Weight"].ToString(),
                            Spec = dr["Spec"].ToString(),
                            ArticleGrpID = dr["ArticleGrpID"].ToString(),
                            BuyerArticleNo = dr["BuyerArticleNo"].ToString(),
                            UnitPrice = dr["UnitPrice"].ToString(),
                            UnitPriceClss = dr["UnitPriceClss"].ToString(),
                            UnitClss = dr["UnitClss"].ToString(),
                            Code_Name = dr["Code_Name"].ToString(),
                            //ProcessName = dr["ProcessName"].ToString(),
                            //HSCode = dr["HSCode"].ToString(),
                            OutUnitPrice = dr["OutUnitPrice"].ToString(),
                            BuyerModelID = dr["BuyerModelID"].ToString(),
                            BuyerModel = dr["BuyerModel"].ToString(),
                        };
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

        private void chkDvlyDate_Checked(object sender, RoutedEventArgs e)
        {
            dtpDvlyDate.IsEnabled = true;
        }

        private void chkDvlyDate_Unchecked(object sender, RoutedEventArgs e)
        {
            dtpDvlyDate.IsEnabled = false;
        }
        


        private void FillNeedStockQty(string strArticleID, string strQty)
        {
            if (dgdNeedStuff.Items.Count > 0)
                dgdNeedStuff.Items.Clear();

            try
            {
                Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                sqlParameter.Add("ArticleID", strArticleID);
                sqlParameter.Add("OrderQty", RemoveComma(strQty, true, typeof(decimal)));
                DataSet ds = DataStore.Instance.ProcedureToDataSet("xp_Order_sArticleNeedStockQty", sqlParameter, false);

                if (ds != null && ds.Tables.Count > 0)
                {
                    DataTable dt = ds.Tables[0];

                    if (dt.Rows.Count > 0)
                    {
                        DataRowCollection drc = dt.Rows;
                        int i = 0;

                        foreach (DataRow dr in drc.Cast<DataRow>().Skip(1))
                        {
                            i++;
                            var NeedStockQty = new ArticleNeedStockQty()
                            {
                                num = i,
                                BuyerArticleNo = dr["BuyerArticleNo"].ToString(),
                                Article = dr["Article"].ToString(),
                                UnitClssName = dr["UnitClssName"].ToString(),
                                FinalNeedQty = stringFormatN2(dr["FinalNeedQty"]),
                                StuffInQty = stringFormatN0(dr["StuffinQty"]),
                                NeededQty = stringFormatN0(dr["NeededQty"]),
                            };

                            

                            //if (Lib.Instance.IsNumOrAnother(NeedStockQty.FinalNeedQty))
                            //{
                            //    double finalNeedQty;
                            //    if (double.TryParse(NeedStockQty.FinalNeedQty, out finalNeedQty))
                            //    {
                            //        // FinalNeedQty의 소숫점 아래가 전부 0이면 정수형태로 표현, 아니면 소숫점 5자리까지 표현
                            //        string formattedFinalNeedQty = finalNeedQty % 1 == 0 ? finalNeedQty.ToString("N0") : finalNeedQty.ToString("N2");

                            //        // 결과 할당
                            //        NeedStockQty.FinalNeedQty = formattedFinalNeedQty;
                            //    }
                            //}

                            dgdNeedStuff.Items.Add(NeedStockQty);
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

    

        private void dgdMain_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            if (btnUpdate.IsEnabled == true)
            {
                if (e.ClickCount == 2)
                {
                    btnUpdate_Click(null, null);
                }
            }
        }

        #region 기타 메서드 모음

        // 천마리 콤마, 소수점 버리기
        private string stringFormatN0(object obj)
        {
            return string.Format("{0:N0}", obj);
        }

        // 천마리 콤마, 소수점 한자리
        private string stringFormatN1(object obj)
        {
            return string.Format("{0:N1}", obj);
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

                if (int.TryParse(str, out chkInt) == true)
                    result = int.Parse(str);
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
                if (double.TryParse(str, out chkDouble) == true)
                    flag = true;
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

                if (int.TryParse(str, out chkInt) == true)
                    flag = true;
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

                if (double.TryParse(str, out chkDouble) == true)
                    result = double.Parse(str);
            }

            return result;
        }
        #endregion

        #region keyDown 이벤트(커서이동)

        //숫자 외에 다른 문자열 못들어오도록
        public bool IsNumeric(string source)
        {

            Regex regex = new Regex("[^0-9.-]+");
            return !regex.IsMatch(source);
        }

        //총주문량 숫자 외에 못들어가게 
        private void TxtAmount_PreviewTextInput(object sender, TextCompositionEventArgs e)
        {
            e.Handled = !IsNumeric(e.Text);
        }

        //단가 숫자 외에 못들어가게
        private void TxtUnitPrice_TextInput(object sender, TextCompositionEventArgs e)
        {
            e.Handled = !IsNumeric(e.Text);
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

   
        //최종거래처 
        private void txtInCustomID_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
                MainWindow.pf.ReturnCode(txtInCustomID, 0, "");
        }

        //최종거래처
        private void btnInCustomID_Click(object sender, RoutedEventArgs e)
        {
            MainWindow.pf.ReturnCode(txtInCustomID, 0, "");
        }
        #endregion keydown 이벤트

        //자재필요량조회
        private void btnNeedStuff_Click(object sender, RoutedEventArgs e)
        {
            if (txtBuyerArticleNo.Tag == null   )
            {
                MessageBox.Show("먼저 품명을 선택해주세요");
                return;
            }

            if (RemoveComma(txtOrderQty.Text).ToString() == string.Empty) 
            {
                MessageBox.Show("먼저 총 주문량을 입력해주세요");
                return;
            }

            //자재필요량조회에 필요한 파라미터 값을 넘겨주자, 품명이랑 주문량
            FillNeedStockQty(txtBuyerArticleNo.Tag.ToString(), RemoveComma(txtOrderQty.Text).ToString());
        }

        

        //메인 데이터그리드 선택 이벤트
        private void DataGridMain_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            try
            {
                OrderID_List.Clear();

                foreach (var item in dgdMain.SelectedItems)
                {
                    if (item is Win_ord_Order_U_CodeView OrderView)
                    {
                        OrderID_List.Add(OrderView.OrderID);
                    }
                }

                var OrderInfo = dgdMain.SelectedItem as Win_ord_Order_U_CodeView;

                if(OrderInfo != null)
                {
                    this.DataContext = OrderInfo;
        
                    if (string.IsNullOrEmpty(OrderInfo.CloseClss.Trim())) cboCloseClss.SelectedIndex = 0;
                    else if (OrderInfo.CloseClss != null) cboCloseClss.SelectedValue = OrderInfo.CloseClss;

                    if (!string.IsNullOrEmpty(OrderInfo.DvlyDate.Trim()))
                    {
                        chkDvlyDate.IsChecked = true;
                        dtpDvlyDate.IsEnabled = true;
                    }
                    else
                    {
                        chkDvlyDate.IsChecked = false;
                        dtpDvlyDate.IsEnabled = false;
                    }
                   

                    FillNeedStockQty(OrderInfo.ArticleID, RemoveComma(txtOrderQty.Text).ToString());

                }

            }
            catch (Exception ee)
            {
                MessageBox.Show("오류지점 - DataGridMain_SelectionChanged : " + ee.ToString());
            }     

        }

        private void txtBuyerArticleNo_KeyDown(object sender, KeyEventArgs e)
        {
            try
            {
                if (e.Key == Key.Enter)
                {

                    if(txtCustomID.Text != string.Empty || txtCustomID.Tag != null)
                    {
                        MainWindow.pf.ReturnCode(txtBuyerArticleNo, 7070, txtCustomID.Tag.ToString());
                    }
                    else
                    {
                        MainWindow.pf.ReturnCode(txtBuyerArticleNo, 7071, "");
                    }

                    if (txtBuyerArticleNo.Tag != null)
                    {
                        CallArticleData(txtBuyerArticleNo.Tag.ToString());

                        //품명종류 대입(ex.제품 등)
                        cboArticleGroup.SelectedValue = articleData.ArticleGrpID;
                        cboUnitClss.SelectedValue = articleData.UnitClss;
                        //품명 대입
                        txtArticleID.Text = articleData.Article;
                        txtArticleID.Tag = articleData.ArticleID;

                        txtBuyerArticleNo.Tag = articleData.ArticleID;
                        //단가 대입
                        txtUnitPrice.Text = articleData.OutUnitPrice;
                        //차종 대입
                        txtBuyerModelID.Tag = articleData.BuyerModelID;
                        txtBuyerModelID.Text = articleData.BuyerModel;
                    }
    
                }
            }
            catch (Exception ex)
            {
                //MessageBox.Show("오류 발생, 오류 내용 : " + ex.ToString());
            }
            finally
            {
                DataStore.Instance.CloseConnection();
            }
        }

        private void btnBuyerArticleNo_Click(object sender, RoutedEventArgs e)
        {
            try
            {

                if (txtCustomID.Text != string.Empty || txtCustomID.Tag != null)
                {
                    MainWindow.pf.ReturnCode(txtBuyerArticleNo, 7070, txtCustomID.Tag.ToString());
                }
                else
                {
                    MainWindow.pf.ReturnCode(txtBuyerArticleNo, 7071, "");
                }


                if (txtBuyerArticleNo.Tag != null)
                {
                    CallArticleData(txtBuyerArticleNo.Tag.ToString());

                    //품명종류 대입(ex.제품 등)
                    cboArticleGroup.SelectedValue = articleData.ArticleGrpID;
                    cboUnitClss.SelectedValue = articleData.UnitClss;

                    //품명 대입
                    txtArticleID.Text = articleData.Article;
                    //단가 대입
                    txtUnitPrice.Text = articleData.OutUnitPrice;
                    //차종 대입
                    txtBuyerModelID.Tag = articleData.BuyerModelID;
                    txtBuyerModelID.Text = articleData.BuyerModel;
                }

      
            }
            catch (Exception ex)
            {
                //MessageBox.Show("오류 발생, 오류 내용 : " + ex.ToString());
            }
            finally
            {
                DataStore.Instance.CloseConnection();
            }
        }

        private void rbnOrderNo_Click(object sender, RoutedEventArgs e)
        {
            dgdtxtOrderID.Visibility = Visibility.Hidden;
            dgdtxtOrderNo.Visibility = Visibility.Visible;
        }

        private void rbnOrderID_Click(object sender, RoutedEventArgs e)
        {
            dgdtxtOrderID.Visibility = Visibility.Visible;
            dgdtxtOrderNo.Visibility = Visibility.Hidden;
        }
            

 

        #region 기타 메서드 모음 ADD

        //그리드 클리어
        private void ClearGrdInput()
        {
            List<Grid> grids = new List<Grid> { grdInput };

            foreach (Grid grid in grids)
            {
                FindUiObject(grid, child =>
                {
                    if (child is TextBox txtbox)
                    {
                        txtbox.Text = string.Empty;
                        txtbox.Tag = null;
                    }
                    else if (child is DatePicker dtp)
                    {
                        dtp.SelectedDate = null;
                    }
                    else if (child is ComboBox cb)
                    {
                        cb.SelectedValue = null;
                    }
                    //else if (child is DataGrid dgd)
                    //{
                    //    if (dgd.ItemsSource != null)
                    //    {
                    //        var originalCollection = dgd.ItemsSource;
                    //        dgd.ItemsSource = null;

                    //        if (originalCollection is IList list)
                    //        {
                    //            list.Clear();
                    //            dgd.ItemsSource = originalCollection;
                    //        }
                    //        else if (originalCollection is ObservableCollection<object> ovc)
                    //        {
                    //            ovc.Clear();
                    //            dgd.ItemsSource = originalCollection;
                    //        }

                    //    }
                    //    else
                    //    {
                    //        dgd.Items.Clear();
                    //    }
                    //}

                });
            }
        }

        //라벨클릭, 체크박스 토글 코드가 너무 반복 되어서 작성
        private void CommonControl_Click(object sender, EventArgs e)
        {
            CheckBox checkBox = null;
            DependencyObject parentGrid = null;

            if (sender is Label label)
            {
                // 라벨의 부모 그리드 찾기
                parentGrid = FindVisualParent<Grid>(label);
                if (parentGrid != null)
                {
                    // 같은 그리드 내에서 체크박스 찾기
                    checkBox = FindChild<CheckBox>(parentGrid);
                    if (checkBox != null)
                    {
                        // 체크박스 상태 토글
                        checkBox.IsChecked = !checkBox.IsChecked;
                    }
                }
            }
            else if (sender is CheckBox clickedCheckBox)
            {
                // 클릭된 것이 체크박스인 경우
                checkBox = clickedCheckBox;
                parentGrid = FindVisualParent<Grid>(checkBox);
            }

            // 체크박스와 부모 그리드가 있으면 컨트롤 활성화/비활성화 처리
            if (checkBox != null && parentGrid != null)
            {
                List<Control> controlsToToggle = new List<Control>();

                // 그리드 내 모든 Control 찾기 (체크박스 제외)
                FindUiObject(parentGrid, obj => {
                    if (obj is Control control && obj != checkBox && !(obj is Label) && !(obj is CheckBox))
                    {
                        controlsToToggle.Add(control);
                    }
                });

                // 컨트롤 활성화/비활성화
                foreach (var control in controlsToToggle)
                {
                    control.IsEnabled = checkBox.IsChecked == true;
                }
            }
        }

        //UI컨트롤 요소찾기
        private void FindUiObject(DependencyObject parent, Action<DependencyObject> action)
        {
            int childCount = VisualTreeHelper.GetChildrenCount(parent);
            for (int i = 0; i < childCount; i++)
            {
                var child = VisualTreeHelper.GetChild(parent, i);

                action?.Invoke(child);

                FindUiObject(child, action);
            }
        }

        //컨트롤 안 특정 타입의 자식 컨트롤을 찾는 함수 (그리드내에서)
        //var parentContainer = VisualTreeHelper.GetParent(checkbox);
        //var datePicker = FindChild<DatePicker>(parentContainer);
        private T FindChild<T>(DependencyObject parent) where T : DependencyObject
        {
            int childCount = VisualTreeHelper.GetChildrenCount(parent);
            for (int i = 0; i < childCount; i++)
            {
                var child = VisualTreeHelper.GetChild(parent, i);
                if (child is T typedChild)
                {
                    return typedChild;
                }

                // 재귀적으로 자식의 자식들도 검색
                var result = FindChild<T>(child);
                if (result != null)
                    return result;
            }
            return null;
        }


        // 자식요소 안에서 부모요소 찾기
        //DataGridRow row = FindVisualParent<DataGridRow>(checkBox); 데이터그리드안의 행속 체크박스의 부모행 찾기
        //DataGrid parentGrid = FindVisualParent<DataGrid>(row); 데이터그리드 행의 부모 데이터그리드 찾기
        private T FindVisualParent<T>(DependencyObject child) where T : DependencyObject
        {
            DependencyObject parentObject = VisualTreeHelper.GetParent(child);

            if (parentObject == null)
                return null;

            T parent = parentObject as T;
            if (parent != null)
                return parent;
            else
                return FindVisualParent<T>(parentObject);
        }

        //8자리 char형태 날짜 년도-월-일 하이픈 삽입
        //16자리 일경우 8자리 사이에 ~ 삽입
        private string DateTypeHyphen(string DigitsDate)
        {
            string pattern1 = @"(\d{4})(\d{2})(\d{2})";
            string pattern2 = @"(\d{4})(\d{2})(\d{2})(\d{4})(\d{2})(\d{2})";

            if (DigitsDate.Length == 8)
            {
                DigitsDate = Regex.Replace(DigitsDate, pattern1, "$1-$2-$3");
            }
            else if (DigitsDate.Length == 16)
            {
                DigitsDate = Regex.Replace(DigitsDate, pattern2, "$1-$2-$3 ~ $4-$5-$6");
            }
            else if (DigitsDate.Length == 0)
            {
                DigitsDate = string.Empty;
            }

            return DigitsDate;
        }

        private object RemoveComma(object obj, bool returnAsNumber = false, Type returnType = null)
        {
            //파라미터가 만약 null일때
            if (obj == null)
            {
                //숫자타입이 false면 string으로 내보내기
                if (!returnAsNumber) return "0";

                // 만약 숫자타입을 써야되면 returnType파라미터의 받은 형태로 전달
                // null일 때도 returnType에 따라 적절한 타입의 0 반환
                switch (returnType?.Name)
                {
                    case "Decimal": return (object)0m;  //monetary
                    case "Double": return (object)0d;   //double
                    case "Int64": return (object)0L;    //long
                    default: return (object)0;          //int
                }
            }

            string digits = obj.ToString()
                              .Trim()
                              .Replace(",", "");

            //만약 빈공백(blank)이더라도 0으로 내보내야한다.
            if (string.IsNullOrEmpty(digits))
            {
                if (!returnAsNumber) return "0";

                // returnType을 활용해서 적절한 타입으로 반환
                switch (returnType?.Name)
                {
                    case "Decimal": return (object)0m;
                    case "Double": return (object)0d;
                    case "Int64": return (object)0L;
                    default: return (object)0;
                }
            }


            try
            {
                Type targetType = returnType ?? typeof(int);

                //혹시나 하는 예외처리
                //입력 컨트롤간에 LostFocus나 TextChanged같은 걸로 계산을 할 때
                //처리 가능한 숫자 범위를 초과하면 오류가 발생하므로
                //초과하면 해당 자료형타입이 처리할 수 있는 최대 숫자를 표시해줌
                switch (targetType.Name)
                {
                    case "Int32":
                        if (decimal.TryParse(digits, out decimal intParsed))
                        {
                            if (intParsed > int.MaxValue) return int.MaxValue;
                            if (intParsed < int.MinValue) return int.MinValue;
                            return (int)intParsed;
                        }
                        return int.MaxValue;

                    case "Int64":
                        if (decimal.TryParse(digits, out decimal longParsed))
                        {
                            if (longParsed > long.MaxValue) return long.MaxValue;
                            if (longParsed < long.MinValue) return long.MinValue;
                            return (long)longParsed;
                        }
                        return long.MaxValue;

                    case "Double":
                        if (double.TryParse(digits, out double doubleParsed))
                        {
                            return doubleParsed;
                        }
                        return double.MaxValue;

                    case "Decimal":
                        if (decimal.TryParse(digits, out decimal decimalParsed))
                        {
                            return decimalParsed;
                        }
                        return decimal.MaxValue;

                    default:
                        return int.MaxValue;
                }
            }
            catch
            {

                if (returnType != null)
                {
                    switch (returnType.Name)
                    {
                        case "Int32":
                            return int.MaxValue;
                        case "Int64":
                            return long.MaxValue;
                        case "Double":
                            return double.MaxValue;
                        case "Decimal":
                            return decimal.MaxValue;
                        default:
                            return int.MaxValue;
                    }
                }
                return int.MaxValue;
            }
        }

        private string ConvertDate(DatePicker datePicker)
        {
            if (datePicker.SelectedDate != null)
                return datePicker.SelectedDate.Value.ToString("yyyyMMdd");
            else
                return string.Empty;
        }

        private bool IsDatePickerNull(DatePicker datePicker)
        {
            if (datePicker.SelectedDate == null)
                return true;
            else
                return false;
        }



        //텍스트박스 , DatePicker, 콤보박스의 바인딩 값과 넘겨주는 오브젝트 value가 일치하는 곳에
        //자동으로 바인딩
        //사용하려하면 바인딩하려는 UI개체에 updateSourceTrigger를 propertyChange, Tag값도 변경하려면 mode=TwoWay를 작성하세요
        private void AutoBindDataToControls(object dataObject, DependencyObject parent)
        {
            var properties = dataObject.GetType().GetProperties()
                .ToDictionary(p => p.Name.ToLower(), p => p);

            // TextBox 처리
            var textBoxes = FindAllControls<TextBox>(parent);
            foreach (var textBox in textBoxes)
            {
                // Text 바인딩 처리
                var textBinding = BindingOperations.GetBinding(textBox, TextBox.TextProperty);
                if (textBinding != null && !string.IsNullOrEmpty(textBinding.Path.Path))
                {
                    var textPropertyName = textBinding.Path.Path.ToLower();
                    if (properties.TryGetValue(textPropertyName, out var textProperty))
                    {
                        var textValue = textProperty.GetValue(dataObject)?.ToString();
                        if (decimal.TryParse(textValue, out _))
                            textBox.Text = stringFormatN0(textValue);
                        else
                            textBox.Text = textValue;
                    }
                }

                // Tag 바인딩 처리
                var tagBinding = BindingOperations.GetBinding(textBox, TextBox.TagProperty);
                if (tagBinding != null && !string.IsNullOrEmpty(tagBinding.Path.Path))
                {
                    var tagPropertyName = tagBinding.Path.Path.ToLower();
                    if (properties.TryGetValue(tagPropertyName, out var tagProperty))
                    {
                        textBox.Tag = tagProperty.GetValue(dataObject)?.ToString();
                    }
                }
            }

            // DatePicker 처리
            var datePickers = FindAllControls<DatePicker>(parent);
            foreach (var datePicker in datePickers)
            {
                var binding = BindingOperations.GetBinding(datePicker, DatePicker.SelectedDateProperty);
                if (binding != null && !string.IsNullOrEmpty(binding.Path.Path))
                {
                    var propertyName = binding.Path.Path.ToLower();
                    if (properties.TryGetValue(propertyName, out var property))
                    {
                        datePicker.SelectedDate = ConvertToDateTime(property.GetValue(dataObject)?.ToString());
                    }
                }
            }

            // ComboBox 처리
            var comboBoxes = FindAllControls<ComboBox>(parent);
            foreach (var comboBox in comboBoxes)
            {
                var binding = BindingOperations.GetBinding(comboBox, ComboBox.SelectedValueProperty);
                if (binding != null && !string.IsNullOrEmpty(binding.Path.Path))
                {
                    var propertyName = binding.Path.Path.ToLower();
                    if (properties.TryGetValue(propertyName, out var property))
                    {
                        comboBox.SelectedValue = property.GetValue(dataObject)?.ToString();
                    }
                }
            }
        }

        private IEnumerable<T> FindAllControls<T>(DependencyObject parent) where T : DependencyObject
        {
            var count = VisualTreeHelper.GetChildrenCount(parent);
            for (int i = 0; i < count; i++)
            {
                var child = VisualTreeHelper.GetChild(parent, i);

                if (child is T control)
                    yield return control;

                foreach (var descendant in FindAllControls<T>(child))
                    yield return descendant;
            }
        }

        // 단일 컨트롤을 찾는 메서드도 필요할 수 있습니다
        private T FindControl<T>(DependencyObject parent, string name) where T : FrameworkElement
        {
            if (parent == null) return null;

            var count = VisualTreeHelper.GetChildrenCount(parent);
            for (int i = 0; i < count; i++)
            {
                var child = VisualTreeHelper.GetChild(parent, i);

                if (child is T control && control.Name == name)
                    return control;

                var result = FindControl<T>(child, name);
                if (result != null)
                    return result;
            }

            return null;
        }

        private DateTime? ConvertToDateTime(string dateStr)
        {
            if (string.IsNullOrEmpty(dateStr?.Trim()))
                return null;

            // 특수문자 제거 (숫자만 남김)
            string cleanDate = new string(dateStr.Where(char.IsDigit).ToArray());

            // 8자리가 아닌 경우 null 반환
            if (cleanDate.Length != 8)
                return null;

            try
            {
                return DateTime.ParseExact(cleanDate, "yyyyMMdd", null);
            }
            catch
            {
                return null;
            }
        }

        #endregion

        #region FTP

        //Dictionary<string, string> lstFtpFilePath = new Dictionary<string, string>();
        HashSet<string> lstFilesName = new HashSet<string>(); //중복거르기

        private void FileUpload_Click(object sender, RoutedEventArgs e)
        {
            Button btn = sender as Button;
            if (btn.Tag != null)
            {
                Grid grd = FindVisualParent<Grid>(btn);
                if (grd != null)
                {
                    FindUiObject(grd, child =>
                    {
                        if (child is TextBox txtbox)
                        {
                            FTP_Upload_TextBox(txtbox);
                        }
                    });
                }
            }
        }

        private void FTP_Upload_TextBox(TextBox textBox)
        {
            if (!textBox.Text.Equals(string.Empty) && strFlag.Equals("U"))
            {
                MessageBox.Show("먼저 해당파일의 삭제를 진행 후 진행해주세요.");
                return;
            }
            else
            {
                Microsoft.Win32.OpenFileDialog OFdlg = new Microsoft.Win32.OpenFileDialog();

                Nullable<bool> result = OFdlg.ShowDialog();
                if (result == true)
                {
                    // 선택된 파일의 확장자 체크
                    if (MainWindow.OFdlg_Filter_NotAllowed.Contains(Path.GetExtension(OFdlg.FileName).ToLower()))
                    {
                        MessageBox.Show("보안상의 이유로 해당 파일은 업로드할 수 없습니다.");
                        return;
                    }

                    strFullPath = OFdlg.FileName;

                    string ImageFileName = OFdlg.SafeFileName;  //명.
                    string ImageFilePath = string.Empty;       // 경로

                    ImageFilePath = strFullPath.Replace(ImageFileName, "");

                    StreamReader sr = new StreamReader(OFdlg.FileName);
                    long FileSize = sr.BaseStream.Length;
                    if (sr.BaseStream.Length > (2048 * 1000))
                    {
                        //업로드 파일 사이즈범위 초과
                        MessageBox.Show("업로드하려는 파일사이즈가 2M byte를 초과하였습니다.");
                        sr.Close();
                        return;
                    }
                    if (!FTP_Upload_Name_Cheking(ImageFileName))
                    {
                        MessageBox.Show("업로드 하려는 파일 중, 이름이 중복된 항목이 있습니다." +
                                        "\n파일 이름을 변경하고 다시 시도하여 주세요");
                    }
                    else
                    {
                        textBox.Text = ImageFileName;
                        textBox.Tag = ImageFilePath;

                        string[] strTemp = new string[] { ImageFileName, ImageFilePath.ToString() };
                        listFtpFile.Add(strTemp);
                    }
                }
            }
        }


        private bool FTP_Upload_Name_Cheking(string fileName)
        {
            bool flag = true;

            if (!lstFilesName.Add(fileName))
            {
                flag = false;
                return flag;
            }

            return flag;
        }


        // 파일 저장하기.
        private void FTP_Save_File(List<string[]> listStrArrayFileInfo, string MakeFolderName)
        {
            _ftp = new FTP_EX(FTP_ADDRESS, FTP_ID, FTP_PASS);

            List<string[]> UpdateFilesInfo = new List<string[]>();
            string[] fileListSimple;
            string[] fileListDetail = null;
            fileListSimple = _ftp.directoryListSimple("", Encoding.Default);

            // 기존 폴더 확인작업.
            bool MakeFolder = false;
            MakeFolder = FolderInfoAndFlag(fileListSimple, MakeFolderName);

            bool Makefind = false;
            Makefind = FileInfoAndFlag(fileListSimple, MakeFolderName);


            if (MakeFolder == false)
            {


                if (_ftp.createDirectory(MakeFolderName) == false)
                {
                    MessageBox.Show("업로드를 위한 폴더를 생성할 수 없습니다.");
                    return;
                }

            }
            else
            {
                fileListDetail = _ftp.directoryListSimple(MakeFolderName, Encoding.Default);
            }

            for (int i = 0; i < listStrArrayFileInfo.Count; i++)
            {
                bool flag = true;

                if (fileListDetail != null)
                {
                    foreach (string compare in fileListDetail)
                    {
                        if (compare.Equals(listStrArrayFileInfo[i][0]))
                        {
                            flag = false;
                            break;
                        }
                    }
                }

                if (flag)
                {
                    listStrArrayFileInfo[i][0] = MakeFolderName + "/" + listStrArrayFileInfo[i][0];
                    UpdateFilesInfo.Add(listStrArrayFileInfo[i]);
                }
            }

            if (!_ftp.UploadTempFilesToFTP(UpdateFilesInfo))
            {
                MessageBox.Show("파일업로드에 실패하였습니다.");
                return;
            }

        }

        private void btnFileSee_Click(object sender, RoutedEventArgs e)
        {
            if (txtOrderID.Text != null)
            {
                MessageBoxResult msgresult = MessageBox.Show("다운로드 후 파일을 바로 여시겠습니까?", "보기 확인", MessageBoxButton.YesNoCancel);
                if (msgresult == MessageBoxResult.Yes)
                {

                    string str_remotepath = string.Empty;
                    string str_localpath = string.Empty;

                    Button btn = sender as Button;
                    if (btn.Tag != null)
                    {
                        Grid grd = FindVisualParent<Grid>(btn);
                        if (grd != null)
                        {
                            FindUiObject(grd, child =>
                            {
                                if (child is TextBox txtbox)
                                {
                                    if (txtbox.Text.Trim() != string.Empty)
                                    {
                                        str_remotepath = txtbox.Text;
                                        str_localpath = LOCAL_DOWN_PATH + "\\" + txtbox.Text;
                                    }
                                    else
                                    {
                                        MessageBox.Show("등록된 파일이 없습니다.", "확인");
                                        return;
                                    }
                                }

                            });
                        }
                    }

                    if (str_remotepath == string.Empty) return;

                    try
                    {
                        // 접속 경로
                        _ftp = new FTP_EX(FTP_ADDRESS, FTP_ID, FTP_PASS);

                        string str_path = string.Empty;
                        str_path = FTP_ADDRESS + '/' + txtOrderID.Text;
                        _ftp = new FTP_EX(str_path, FTP_ID, FTP_PASS);

                        DirectoryInfo DI = new DirectoryInfo(LOCAL_DOWN_PATH);
                        if (DI.Exists == false)
                        {
                            DI.Create();
                        }

                        FileInfo file = new FileInfo(str_localpath);
                        try
                        {
                            file.Delete();
                        }
                        catch (IOException)
                        {
                            // 파일명과 확장자 분리
                            string directory = Path.GetDirectoryName(str_localpath);
                            string fileName = Path.GetFileNameWithoutExtension(str_localpath);
                            string extension = Path.GetExtension(str_localpath);

                            // 복사본 파일명 생성 (예: test.hwp -> test - 복사본.hwp)
                            int copyNum = 1;
                            string newPath = Path.Combine(directory, $"{fileName} - 복사본{extension}");

                            // 복사본 파일이 이미 존재하면 번호 추가 (예: test - 복사본 (2).hwp)
                            while (File.Exists(newPath))
                            {
                                copyNum++;
                                newPath = Path.Combine(directory, $"{fileName} - 복사본 ({copyNum}){extension}");
                            }

                            str_localpath = newPath; // 새로운 경로로 업데이트
                            MessageBox.Show("파일이 사용 중이어서 복사본으로 다운로드 했습니다.", "알림");
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show("파일 처리 중 오류가 발생했습니다: " + ex.Message);
                            return;
                        }

                        _ftp.download(str_remotepath, str_localpath);

                        //파일 다운로드 후 바로 열기
                        if (File.Exists(str_localpath) && msgresult == MessageBoxResult.Yes)
                        {
                            try
                            {
                                System.Diagnostics.Process.Start(new ProcessStartInfo
                                {
                                    FileName = str_localpath,
                                    UseShellExecute = true
                                });
                            }
                            catch (Exception ex)
                            {
                                MessageBox.Show("파일을 여는 중 오류가 발생했습니다:" +
                                    "\n파일을 열기위한 프로그램이 없거나 기본 실행프로그램이 지정이 안 되었을 수도 있습니다." + ex.Message);
                            }
                        }
                        else if ((File.Exists(str_localpath) && msgresult == MessageBoxResult.No))
                        {
                            MessageBox.Show("파일을 다운로드 하였습니다.", "확인");
                            try
                            {
                                string folderPath = Path.GetDirectoryName(str_localpath);
                                //폴더이름의 타이틀명을 찾
                                var openFolders = Process.GetProcessesByName("explorer")
                                    .Where(p =>
                                    {
                                        try
                                        {
                                            return p.MainWindowTitle.Contains(Path.GetFileName(folderPath));
                                        }
                                        catch
                                        {
                                            return false;
                                        }
                                    });

                                if (!openFolders.Any())
                                {
                                    // 폴더가 열려있지 않을 때만 새로 열기
                                    Process.Start("explorer.exe", $"\"{folderPath}\"");
                                }
                            }
                            catch (Exception ex)
                            {
                                MessageBox.Show("폴더를 여는 중 오류가 발생했습니다:" + ex.Message);
                            }
                        }

                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("파일이 존재하지 않습니다.\r관리자에게 문의해주세요.");
                        return;
                    }
                }
            }
        }


        private void btnFileDelete_Click(object sender, RoutedEventArgs e)
        {
            MessageBoxResult msgresult = MessageBox.Show("파일을 삭제 하시겠습니까?", "삭제 확인", MessageBoxButton.YesNo);
            if (msgresult == MessageBoxResult.Yes)
            {
                Button btn = sender as Button;
                if (btn.Tag != null)
                {
                    Grid grd = FindVisualParent<Grid>(btn);
                    if (grd != null)
                    {
                        FindUiObject(grd, child =>
                        {
                            if (child is TextBox txtbox)
                            {
                                if (txtbox.Text.Trim() != string.Empty)
                                {
                                    string FileName = txtbox.Text;
                                    TextBox tempTextBox = new TextBox();
                                    tempTextBox.Text = FileName;
                                    FileDeleteAndTextBoxEmpty(tempTextBox);
                                    lstFilesName.Remove(FileName);

                                    BindingExpression textBinding = txtbox.GetBindingExpression(TextBox.TextProperty);
                                    string textPropertyName = textBinding?.ParentBinding?.Path.Path; // "sketch1FileName"

                                    BindingExpression tagBinding = txtbox.GetBindingExpression(TextBox.TagProperty);
                                    string tagPropertyName = tagBinding?.ParentBinding?.Path.Path; // "sketch1FilePath"

                                    var OrderView = dgdMain.SelectedItem as Win_ord_Order_U_CodeView;

                                    OrderView?.GetType().GetProperty(textPropertyName)?.SetValue(OrderView, string.Empty);
                                    OrderView?.GetType().GetProperty(tagPropertyName)?.SetValue(OrderView, string.Empty);
                                }
                                else
                                {
                                    MessageBox.Show("등록된 파일이 없습니다.", "확인");
                                    return;
                                }

                            }
                        });
                    }
                }

                //string ClickPoint = ((Button)sender).Tag.ToString();
                //string fileName = string.Empty;

                //string btndgdSubDown = string.Empty;
                //DataGridCellInfo cell;


                //if ((ClickPoint == "dgdSubDelete") && (btndgdSubDown != string.Empty))
                //{
                //    fileName = btndgdSubDown;
                //    // 임시 TextBox를 생성하고 값을 복사
                //    TextBox tempTextBox = new TextBox();
                //    tempTextBox.Text = btndgdSubDown;
                //    FileDeleteAndTextBoxEmpty(tempTextBox);
                //    lstFilesName.Remove(fileName);

                //    if (cell.Item != null)
                //    {
                //        var item = cell.Item as Win_ord_Order_EstimateSub_U_CodeView;
                //        if (item != null)
                //        {
                //            item.sketch1FileName = string.Empty;
                //            item.sketch1FilePath = string.Empty;
                //        }
                //    }
                //}
            }


        }
        private void FileDeleteAndTextBoxEmpty(TextBox txt)
        {
            if (strFlag.Equals("U"))
            {

                string[] strFtp = { txt.Text, txt.Tag != null ? txt.Tag.ToString() : "" };

                deleteListFtpFile.Add(strFtp);

            }

            txt.Text = string.Empty;
            txt.Tag = string.Empty;
        }


        //파일 삭제
        private bool FTP_RemoveFile(string strSaveName)
        {
            _ftp = new FTP_EX(FTP_ADDRESS, FTP_ID, FTP_PASS);
            if (_ftp.delete(strSaveName) == true)
            {
                return true;
            }
            else
            {
                return false;
            }
        }

        //폴더 삭제(내부 파일 자동 삭제)
        private bool FTP_RemoveDir(string strSaveName)
        {
            _ftp = new FTP_EX(FTP_ADDRESS, FTP_ID, FTP_PASS);
            if (_ftp.removeDir(strSaveName) == true)
            {
                return true;
            }
            else
            {
                return false;
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

        private void SubGridFileUpload(object sender, RoutedEventArgs e)
        {
            //if (lblMsg.Visibility == Visibility.Visible)
            //{
            //    var dgdSubView = dgdSub.CurrentItem as Win_ord_Order_EstimateSub_U_CodeView;
            //    if (dgdSubView != null)
            //    {
            //        if (dgdSubView.sketch1FilePath != string.Empty
            //               && strFlag.Equals("U"))
            //        {
            //            MessageBox.Show("먼저 해당파일의 삭제를 진행 후 진행해주세요.");
            //            return;
            //        }
            //        else
            //        {
            //            var button = sender as Button;
            //            var stackPanel = button.Parent as StackPanel;
            //            var parentContainer = stackPanel.Parent as Panel;
            //            var textBox = parentContainer.Children.OfType<TextBox>().FirstOrDefault();

            //            if (textBox != null)
            //            {

            //                FTP_Upload_TextBox(textBox);
            //            }
            //        }
            //    }
            //}
        }
        #endregion

   
    }


    public class Win_ord_Order_U_CodeView : BaseView
    {
        public string OrderID { get; set; }
        public string OrderNO { get; set; }   
        public string CustomID { get; set; }
        public string KCustom { get; set; }
        public string CloseClss { get; set; }

        public string OrderBox { get; set; }
        public string OrderQty { get; set; }
        public string OrderAmount { get; set; }
        public string UnitClss { get; set; }
        public string Article { get; set; }
        public string ChunkRate { get; set; }
        public string PatternID { get; set; }

        public string Amount { get; set; }
        public string BuyerModelID { get; set; }
        public string BuyerModel { get; set; }
        public string BuyerArticleNo { get; set; }
        public string PONO { get; set; }

        public string OrderForm { get; set; }
        public string OrderClss { get; set; }
        public string InCustomID { get; set; }
        public string InKCustom { get; set; }
        public string AcptDate { get; set; }
        public string DvlyDate { get; set; }

        public string ArticleID { get; set; }
        public string DvlyPlace { get; set; }
        public string WorkID { get; set; }
        public string PriceClss { get; set; }
        public string ExchRate { get; set; }

        public string Vat_IND_YN { get; set; }
        public string ColorCnt { get; set; }
        public string StuffWidth { get; set; }
        public string StuffWeight { get; set; }
        public string CutQty { get; set; }

        public string WorkWidth { get; set; }
        public string WorkWeight { get; set; }
        public string WorkDensity { get; set; }
        public string LossRate { get; set; }
        public string ReduceRate { get; set; }

        public string TagClss { get; set; }
        public string LabelID { get; set; }
        public string BandID { get; set; }
        public string EndClss { get; set; }
        public string MadeClss { get; set; }

        public string SurfaceClss { get; set; }
        public string ShipClss { get; set; }
        public string AdvnClss { get; set; }
        public string LotClss { get; set; }
        public string EndMark { get; set; }

        public string TagArticle { get; set; }
        public string TagArticle2 { get; set; }
        public string TagOrderNo { get; set; }
        public string TagRemark { get; set; }
        public string Tag { get; set; }

        public string BasisID { get; set; }
        public string BasisUnit { get; set; }
        public string SpendingClss { get; set; }
        public string DyeingID { get; set; }
        public string WorkingClss { get; set; }

        public string BTID { get; set; }
        public string BTIDSeq { get; set; }
        public string ChemClss { get; set; }
        public string AccountClss { get; set; }
        public string ModifyClss { get; set; }

        public string ModifyRemark { get; set; }
        public string CancelRemark { get; set; }
        public string Remark { get; set; }
        public string ActiveClss { get; set; }
        public string ModifyDate { get; set; }

        public string OrderFlag { get; set; }
        public string TagRemark2 { get; set; }
        public string TagRemark3 { get; set; }
        public string TagRemark4 { get; set; }
        public string UnitPriceClss { get; set; }

        public string WeightPerYard { get; set; }
        public string WorkUnitClss { get; set; }
        public string ArticleGrpID { get; set; }
        public string OrderSpec { get; set; }
        public string UnitPrice { get; set; }

        public string CompleteArticleFile { get; set; }
        public string CompleteArticlePath { get; set; }
        public string FirstArticleFile { get; set; }
        public string FirstArticlePath { get; set; }
        public string MediumArticleFIle { get; set; }

        public string MediumArticlePath { get; set; }
        public string sketch1Path { get; set; }
        public string sketch1file { get; set; }
        public string sketch2Path { get; set; }
        public string sketch2file { get; set; }

        public string sketch3Path { get; set; }
        public string sketch3file { get; set; }
        public string sketch4Path { get; set; }
        public string sketch4file { get; set; }
        public string sketch5Path { get; set; }

        public string sketch5file { get; set; }
        public string sketch6Path { get; set; }
        public string sketch6file { get; set; }
        public string ProductAutoInspectYN { get; set; }
        public string kBuyer { get; set; }

        public string BuyerID { get; set; }
        public int Num { get; set; }
        public string AcptDate_CV { get; set; }
        public string DvlyDate_CV { get; set; }
        public string Amount_CV { get; set; }

        public string SketchFile { get; set; }
        public string SketchPath { get; set; }
        public string ImageName { get; set; }

        public string CompanyID { get; set; }
        public string OrderNo { get; set; }
        public string PoNo { get; set; }
        public string OrderFormName { get; set; }
        public string BrandClss { get; set; }
        public string WorkName { get; set; }
        public string OrderClssName { get; set; }

        public string NewArticleQty { get; set; }
        public string RePolishingQty { get; set; }
    }

    public class Win_ord_Order_Total_U_CodeView : BaseView
    {
        public int count { get; set; }
        public string OrderTotalAmount { get; set; }
    }
    public class ArticleData : BaseView
    {
        public string ArticleID { get; set; }
        public string Article { get; set; }
        public string ThreadID { get; set; }
        public string thread { get; set; }
        public string StuffWidth { get; set; }
        public string DyeingID { get; set; }
        public string Weight { get; set; }
        public string Spec { get; set; }
        public string ArticleGrpID { get; set; }
        public string BuyerArticleNo { get; set; }
        public string UnitPrice { get; set; }
        public string UnitPriceClss { get; set; }
        public string UnitClss { get; set; }
        public string ProcessName { get; set; }
        public string HSCode { get; set; }
        public string OutUnitPrice { get; set; }
        public string Code_Name { get; set; }
        public string BuyerModelID { get; set; }
        public string BuyerModel { get; set; }
    }

    public class ArticleNeedStockQty : BaseView
    {
        public int num { get; set; }
        public string BuyerArticleNo { get; set; }
        public string Article { get; set; }
        public string NeedQty { get; set; }
        public string FinalNeedQty { get; set; }
        public string UnitClss { get; set; }
        public string UnitClssName { get; set; }
        public string StuffInQty { get; set; }
        public string NeededQty { get; set; }
        public bool IsNegativeNeededQty => !string.IsNullOrEmpty(StuffInQty) && StuffInQty.StartsWith("-");
    }

    public class OrderExcel : BaseView
    {
        public string CustomID { get; set; }
        public string Model { get; set; }
        public string BuyerArticleNo { get; set; }
        public string Article { get; set; }
        public string UnitClss { get; set; }
        public string OrderQty { get; set; }
    }
}


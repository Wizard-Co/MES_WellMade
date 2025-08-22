using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Text;
using System.Threading;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media.Imaging;
using WizMes_WellMade.PopUP;
using WizMes_WellMade.PopUp;

namespace WizMes_WellMade
{
    /**************************************************************************************************
   '** System 명 : WizMes_WellMade
   '** Author    : Wizard
   '** 작성자    : 최준호
   '** 내용      : 부품 등록
   '** 생성일자  : 2018.10월~2019.2월 사이
   '** 변경일자  : 
   '**------------------------------------------------------------------------------------------------
   ''*************************************************************************************************
   ' 변경일자  , 변경자, 요청자    , 요구사항ID  , 요청 및 작업내용
   '**************************************************************************************************
   ' ex) 2015.11.09, 박진성, 오영      ,S_201510_AFT_03 , 월별집계(가로) 순서 변경 : 합계/10월/9월/8월 순으로
   ' 2019.05월초  최준호 , 최규환  ,중량 제외 요청=>완료 
   ' 2019.07.17  최준호 , 최규환   ,예비품 사진 등록하고 보이게 해달라(ftp)
   '**************************************************************************************************/

    /// <summary>
    /// Win_prd_MCTool_U.xaml에 대한 상호 작용 논리
    /// </summary>
    public partial class Win_prd_MCTool_U : UserControl
    {
        string stDate = string.Empty;
        string stTime = string.Empty;

        string strFlag = string.Empty;
        int rowNum = 0;
        Win_prd_MCTool_U_CodeView WinPartCode = new Win_prd_MCTool_U_CodeView();

        Lib lib = new Lib();

        // FTP 활용모음.
        string strImagePath = string.Empty;
        string strFullPath = string.Empty;
        string strDelFileName = string.Empty;

        List<string[]> listFtpFile = new List<string[]>();
        List<string[]> deleteListFtpFile = new List<string[]>(); // 삭제할 파일 리스트

        private FTP_EX _ftp = null;
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
        
        //string FTP_ADDRESS = "ftp://192.168.0.120";
        //string FTP_ADDRESS = "ftp://wizis.iptime.org/ImageData/MtMcPart";
          string FTP_ADDRESS = "ftp://" + LoadINI.FileSvr + ":"
            + LoadINI.FTPPort + LoadINI.FtpImagePath + "/MtMcPart";
        //string FTP_ADDRESS = "ftp://aftkr.iptime.org/ImageData/MtMcPart";
        //string FTP_ADDRESS = "ftp://192.168.0.95/ImageData/MtMcPart";
        private const string FTP_ID = "wizuser";
        private const string FTP_PASS = "wiz9999";
        private const string LOCAL_DOWN_PATH = "C:\\Temp";

        public Win_prd_MCTool_U()
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

            tgnAll.IsChecked = true;
        }

        #region 콤보박스 세팅

        private void SetComboBox()
        {
            //List<string[]> strValue = new List<string[]>();
            //string[] strAll = { "", "전체" };
            //string[] strOne = { "1", "공용" };
            //string[] strTwo = { "2", "설비예비품" };
            //string[] strThree = { "3", "Tool" };
            //strValue.Add(strAll);
            //strValue.Add(strOne);
            //strValue.Add(strTwo);
            //strValue.Add(strThree);

            ObservableCollection<CodeView> ovcUnitClss = ComboBoxUtil.Instance.Gf_DB_CM_GetComCodeDataset(null, "MTRUNIT", "Y", "", "");
            this.cboUnitClss.ItemsSource = ovcUnitClss;
            this.cboUnitClss.DisplayMemberPath = "code_name";
            this.cboUnitClss.SelectedValuePath = "code_id";

            List<string[]> strArrayValue = new List<string[]>();
            string[] strArrayOne = { "1", "공용" };
            string[] strArrayTwo = { "2", "설비예비품" };
            string[] strArrayThree = { "3", "Tool" };
            strArrayValue.Add(strArrayOne);
            strArrayValue.Add(strArrayTwo);
            strArrayValue.Add(strArrayThree);

            ObservableCollection<CodeView> ovcForUse = ComboBoxUtil.Instance.Direct_SetComboBox(strArrayValue);
            this.cboForUse.ItemsSource = ovcForUse;
            this.cboForUse.DisplayMemberPath = "code_name";
            this.cboForUse.SelectedValuePath = "code_id";
        }

        #endregion

        #region 수정, 추가 후 / 수정, 추가 진행중

        /// <summary>
        /// 수정,추가 저장 후
        /// </summary>
        private void CanBtnControl()
        {
            Lib.Instance.UiButtonEnableChange_IUControl(this);
            //grdInput.IsEnabled = false;
            //dgdMain.IsEnabled = true;
            dgdMain.IsHitTestVisible = true;
            listFtpFile.Clear();
            gridInput1.IsHitTestVisible = false;
            gridInput3.IsEnabled = false;

            btnImgSelect.IsEnabled = false; //이미지 버튼
            btnImageDelete.IsEnabled = false; //이미지 버튼
        }

        /// <summary>
        /// 수정,추가 진행 중
        /// </summary>
        private void CantBtnControl()
        {
            Lib.Instance.UiButtonEnableChange_SCControl(this);
            gridInput1.IsHitTestVisible = true;
            gridInput3.IsEnabled = true;
            btnImgSelect.IsEnabled = true; //이미지 버튼
            btnImageDelete.IsEnabled = true; //이미지 버튼
            //dgdMain.IsEnabled = false;
            dgdMain.IsHitTestVisible = false;
        }

        #endregion

        #region Header - 검색조건

        // 설비 부품명
        private void lblMcPartSrh_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            if (chkMcPartSrh.IsChecked == true)
            {
                chkMcPartSrh.IsChecked = false;
            }
            else
            {
                chkMcPartSrh.IsChecked = true;
            }
        }
        private void chkMcPartSrh_Checked(object sender, RoutedEventArgs e)
        {
            chkMcPartSrh.IsChecked = true;
            txtMcPartSrh.IsEnabled = true;
        }
        private void chkMcPartSrh_Unchecked(object sender, RoutedEventArgs e)
        {
            chkMcPartSrh.IsChecked = false;
            txtMcPartSrh.IsEnabled = false;
        }

        //사용안함 포함
        private void lblNotUseSrh_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            if (chkNotUseSrh.IsChecked == true) { chkNotUseSrh.IsChecked = false; }
            else { chkNotUseSrh.IsChecked = true; }
        }

        #endregion

        #region Header - 상단 오른쪽 버튼

        // 테스트용
        private void btnTag_Click(object sender, RoutedEventArgs e)
        {

        }

        //추가
        private void btnAdd_Click(object sender, RoutedEventArgs e)
        {
            txtMcPartName.Focus();

            CantBtnControl();
            strFlag = "I";
            tbkMsg.Text = "자료 입력 중";
            rowNum = dgdMain.SelectedIndex;
            this.DataContext = null;
            imgPart.Source = null;

            dgdSelProduct.Items.Clear();
            dgdselProcess.Items.Clear();

            //Tool 추가 2021-05-27
            cboForUse.SelectedValue = 3;
            cboUnitClss.SelectedIndex = 0;
        }

        //수정
        private void btnUpdate_Click(object sender, RoutedEventArgs e)
        {
            WinPartCode = dgdMain.SelectedItem as Win_prd_MCTool_U_CodeView;

            if (WinPartCode != null)
            {
                rowNum = dgdMain.SelectedIndex;
                tbkMsg.Text = "자료 수정 중";
                //txtImage.IsEnabled = true;
                CantBtnControl();
                strFlag = "U";
            }
        }

        //삭제 : 삭제가 아니라 사용안함으로 안보이게만 하는 것!!
        private void btnDelete_Click(object sender, RoutedEventArgs e)
        {
            int delCount = 0;
            for (int i = 0; i < dgdMain.Items.Count; i++)
            {
                WinPartCode = dgdMain.Items[i] as Win_prd_MCTool_U_CodeView;

                if (WinPartCode.Flag)
                {
                    delCount++;
                }
            }

            if (delCount <= 0)
            {
                MessageBox.Show("삭제할 데이터가 지정되지 않았습니다. 삭제데이터를 지정하고 눌러주세요");
            }
            else
            {
                if (MessageBox.Show("선택하신 항목을 모두 삭제하시겠습니까?", "삭제 전 확인", MessageBoxButton.YesNo) == MessageBoxResult.Yes)
                {
                    if (dgdMain.Items.Count > 0 && dgdMain.SelectedItem != null)
                    {
                        rowNum = dgdMain.SelectedIndex;
                    }

                    if (DeleteData(WinPartCode.MCPartID))
                    {
                        FillGrid();
                    }
                }
            }
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
            //검색버튼 비활성화
            btnSearch.IsEnabled = false;

            Dispatcher.BeginInvoke(new Action(() =>

            {
                try
                {
                    rowNum = 0;
                    using (Loading lw = new Loading(FillGrid))
                    {
                        lw.ShowDialog();
                        if (dgdMain.Items.Count <= 0)
                        {
                            this.DataContext = null;
                            MessageBox.Show("조회된 내용이 없습니다.");
                        }
                        else
                        {
                            dgdMain.SelectedIndex = rowNum;
                        }

                        btnSearch.IsEnabled = true;
                    }
                }
                catch (Exception ee)
                {
                    MessageBox.Show("예외처리 - " + ee.ToString());
                }

            }), System.Windows.Threading.DispatcherPriority.Background);
        }

        //저장
        private void btnSave_Click(object sender, RoutedEventArgs e)
        {
            if (SaveData(strFlag, txtMcPartCode.Text))
            {
                CanBtnControl();
                re_Search(rowNum);
                strFlag = string.Empty;
                strImagePath = string.Empty;
                btnImgSelect.IsEnabled = false; // 이미지 버튼 비활성화
            }
        }

        //취소
        private void btnCancel_Click(object sender, RoutedEventArgs e)
        {
            CanBtnControl();
            if (!strFlag.Equals(string.Empty))
            {
                re_Search(rowNum);
            }

            strFlag = string.Empty;
            strImagePath = string.Empty;
            btnImgSelect.IsEnabled = false; // 이미지 버튼 비활성화
        }

        #region 엑셀 버튼 클릭 이벤트

        //엑셀
        private void btnExcel_Click(object sender, RoutedEventArgs e)
        {
            DataTable dt = null;
            string Name = string.Empty;

            string[] lst = new string[10];
            lst[0] = "부품 상세조회";
            lst[1] = "전체 제품조회";
            lst[2] = "선택 제품조회";
            lst[3] = "전체 MC조회";
            lst[4] = "선택 MC조회";
            lst[5] = dgdMain.Name;
            lst[6] = dgdAllProduct.Name;
            lst[7] = dgdSelProduct.Name;
            lst[8] = dgdAllProcess.Name;
            lst[9] = dgdselProcess.Name;

            ExportExcelxaml ExpExc = new ExportExcelxaml(lst);
            ExpExc.ShowDialog();

            if (ExpExc.DialogResult.HasValue)
            {
                DataStore.Instance.InsertLogByForm(this.GetType().Name, "E");
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
                else if (ExpExc.choice.Equals(dgdAllProduct.Name))
                {
                    if (ExpExc.Check.Equals("Y"))
                        dt = Lib.Instance.DataGridToDTinHidden(dgdAllProduct);
                    else
                        dt = Lib.Instance.DataGirdToDataTable(dgdAllProduct);

                    Name = dgdAllProduct.Name;
                    if (Lib.Instance.GenerateExcel(dt, Name))
                        Lib.Instance.excel.Visible = true;
                }
                else if (ExpExc.choice.Equals(dgdSelProduct.Name))
                {
                    if (ExpExc.Check.Equals("Y"))
                        dt = Lib.Instance.DataGridToDTinHidden(dgdSelProduct);
                    else
                        dt = Lib.Instance.DataGirdToDataTable(dgdSelProduct);

                    Name = dgdSelProduct.Name;
                    if (Lib.Instance.GenerateExcel(dt, Name))
                        Lib.Instance.excel.Visible = true;
                }
                else if (ExpExc.choice.Equals(dgdAllProcess.Name))
                {
                    if (ExpExc.Check.Equals("Y"))
                        dt = Lib.Instance.DataGridToDTinHidden(dgdAllProcess);
                    else
                        dt = Lib.Instance.DataGirdToDataTable(dgdAllProcess);

                    Name = dgdAllProcess.Name;
                    if (Lib.Instance.GenerateExcel(dt, Name))
                        Lib.Instance.excel.Visible = true;
                }
                else if (ExpExc.choice.Equals(dgdselProcess.Name))
                {
                    if (ExpExc.Check.Equals("Y"))
                        dt = Lib.Instance.DataGridToDTinHidden(dgdselProcess);
                    else
                        dt = Lib.Instance.DataGirdToDataTable(dgdselProcess);

                    Name = dgdselProcess.Name;
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

        #endregion

        #endregion

        /// <summary>
        /// 재검색(수정,삭제,추가 저장후에 자동 재검색)
        /// </summary>
        /// <param name="selectedIndex"></param>
        private void re_Search(int selectedIndex)
        {
            CanBtnControl();

            FillGrid();

            FillGridAllProduct();
            FillGridAllProcess();

            if (dgdMain.Items.Count > 0)
            {
                dgdMain.SelectedIndex = selectedIndex;
            }
            else
            {
                MessageBox.Show("조회된 데이터가 없습니다.");
                dgdMain.Items.Clear();

                txtMcPartCode.Text = "";
                txtMcPartName.Text = "";
                txtUnitPrice.Text = "";
                txtNeedStockQty.Text = "";
                txtSetProdQtyt.Text = "";
                txtSpec.Text = "";
                txtImage.Text = "";

                cboForUse.SelectedIndex = -1;
                cboUnitClss.SelectedIndex = -1;
            }
        }

        #region 주요 메서드 조회 FillGrid

        /// <summary>
        /// 실조회
        /// </summary>
        private void FillGrid()
        {
            if (dgdMain.Items.Count > 0)
            {
                dgdMain.Items.Clear();
            }

            try
            {
                string sForUse = "";
                if (tgnCommon.IsChecked == true) { sForUse = "1"; }
                if (tgnSpare.IsChecked == true) { sForUse = "2"; }
                if (tgnTool.IsChecked == true) { sForUse = "3"; }

                Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                sqlParameter.Clear();
                sqlParameter.Add("sMCPartID", "");
                sqlParameter.Add("sMCPartName", chkMcPartSrh.IsChecked == true && !txtMcPartSrh.Text.Trim().Equals("") ? txtMcPartSrh.Text : "");
                sqlParameter.Add("sIncNotUse", chkNotUseSrh.IsChecked == true ? 1 : 0);
                sqlParameter.Add("sForUse", sForUse);

                DataSet ds = DataStore.Instance.ProcedureToDataSet_LogWrite("xp_Code_sMcPart", sqlParameter, true, "R");

                if (ds != null && ds.Tables.Count > 0)
                {
                    DataTable dt = ds.Tables[0];
                    int i = 0;

                    if (dt.Rows.Count > 0)
                    {
                        DataRowCollection drc = dt.Rows;

                        foreach (DataRow dr in drc)
                        {
                            var WinMCPart = new Win_prd_MCTool_U_CodeView()
                            {
                                Num = i + 1,
                                ForUse = dr["ForUse"].ToString(),
                                ForUseName = dr["ForUseName"].ToString(),
                                ImageName = dr["ImageName"].ToString(),
                                MCPartID = dr["MCPartID"].ToString(),
                                MCPartName = dr["MCPartName"].ToString(),
                                NeedStockQty = stringFormatN0(dr["NeedStockQty"]),
                                Spec = dr["Spec"].ToString(),
                                UnitClss = dr["UnitClss"].ToString(),
                                UnitClssName = dr["UnitClssName"].ToString(),
                                UseClss = dr["UseClss"].ToString(),
                                Weight = stringFormatN0(dr["Weight"]),
                                SetProdQty = stringFormatN0(dr["SetProdQty"]),
                                CreateDate = dr["CreateDate"].ToString(),
                                CreateDate_CV = DatePickerFormat(dr["CreateDate"].ToString()),
                                ManageNo = dr["ManageNo"].ToString(),
                                UnitPrice = ConvertDouble(dr["UnitPrice"].ToString()) //2021-06-10 단가 추가
                                            
                            };

                            if (WinMCPart.UseClss.Trim().Equals("*"))
                            {
                                WinMCPart.UseClssChar = "X";
                            }
                            else
                            {
                                WinMCPart.UseClssChar = "○";
                            }

                            dgdMain.Items.Add(WinMCPart);
                            i++;
                        }

                        tbkCount.Text = "▶ 검색결과 : " + i.ToString() + "건";
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

        //
        private void dgdMain_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            //만약 가져올 이미지가 없다면 비워주기
            imgPart.Source = null;

            WinPartCode = dgdMain.SelectedItem as Win_prd_MCTool_U_CodeView;

            if (WinPartCode != null)
            {
                this.DataContext = WinPartCode;
                FillGrid_selProduct(txtMcPartCode.Text);
                FillGrid_selProcess(txtMcPartCode.Text);

                _ftp = new FTP_EX(FTP_ADDRESS, FTP_ID, FTP_PASS);

                bool MakeFolder = false;
                if (!txtImage.Text.Replace(" ", "").Equals(""))
                {

                    string imageName = txtImage.Text;

                    var WinPartCode = dgdMain.SelectedItem as Win_prd_MCTool_U_CodeView;
                    if (WinPartCode != null
                        && WinPartCode.ImageName != null
                        && WinPartCode.MCPartID != null)
                    {
                        if (CheckImage(imageName.Trim()))
                        {
                            imgPart.Source = SetImage(imageName, WinPartCode.MCPartID);
                        }
                        else
                        {
                            MessageBox.Show(imageName + "는 이미지 변환이 불가능합니다.");
                        }
                    }
                }
            }
        }

        #region 주요 메서드 - 삭제 DeleteData

        /// <summary>
        /// 실삭제
        /// </summary>
        /// <param name="strID"></param>
        /// <returns></returns>
        private bool DeleteData(string strID)
        {
            bool flag = false;
            List<Procedure> Prolist = new List<Procedure>();
            List<Dictionary<string, object>> ListParameter = new List<Dictionary<string, object>>();

            try
            {
                for (int i = 0; i < dgdMain.Items.Count; i++)
                {
                    WinPartCode = dgdMain.Items[i] as Win_prd_MCTool_U_CodeView;

                    if (WinPartCode.Flag
                        && WinPartCode.MCPartID != null)
                    {
                        Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                        sqlParameter.Clear();
                        sqlParameter.Add("sMCPartID", WinPartCode.MCPartID);

                        Procedure pro1 = new Procedure();
                        pro1.Name = "xp_Code_dMcPart";
                        pro1.OutputUseYN = "N";
                        pro1.OutputName = "MCPartID";
                        pro1.OutputLength = "10";

                        Prolist.Add(pro1);
                        ListParameter.Add(sqlParameter);
                    }
                }

                string[] Confirm = new string[2];
                Confirm = DataStore.Instance.ExecuteAllProcedureOutputNew_NewLog(Prolist, ListParameter,"D");
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
            catch (Exception ex)
            {
                MessageBox.Show("오류 발생,오류 내용 : " + ex.ToString());
            }
            finally
            {
                DataStore.Instance.CloseConnection();
            }

            return flag;
        }

        #endregion

        /// <summary>
        /// 저장
        /// </summary>
        /// <param name="strFlag"></param>
        /// <param name="strID"></param>
        /// <returns></returns>
        private bool SaveData(string strFlag, string strID)
        {
            bool flag = false;
            List<Procedure> Prolist = new List<Procedure>();
            List<Dictionary<string, object>> ListParameter = new List<Dictionary<string, object>>();

            string GetKey = "";

            try
            {
                if (CheckData())
                {
                    Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                    sqlParameter.Clear();
                    sqlParameter.Add("sMCPartID", strID);
                    sqlParameter.Add("sMCPartName", txtMcPartName.Text);
                    sqlParameter.Add("Weight", ConvertDouble(txtWeight.Text));
                    sqlParameter.Add("sSpec", txtSpec.Text);
                    sqlParameter.Add("sUseClss", "");

                    sqlParameter.Add("NeedStockQty", ConvertDouble(txtNeedStockQty.Text)); 
                    sqlParameter.Add("sUnitClss", cboUnitClss.SelectedValue != null ? cboUnitClss.SelectedValue.ToString() : "");
                    sqlParameter.Add("sImageName", txtImage.Text);
                    sqlParameter.Add("sForUse", cboForUse.SelectedValue != null ? cboForUse.SelectedValue.ToString() : "");
                    sqlParameter.Add("nSetProdQty", ConvertDouble(txtSetProdQtyt.Text));

                    sqlParameter.Add("ManageNo", "");
                    sqlParameter.Add("UnitPrice", txtUnitPrice.Text == "" ? "0" : txtUnitPrice.Text.Replace(",", "")); //2021-06-10 설비부품 단가 추가

                    #region 추가

                    if (strFlag.Equals("I"))
                    {
                        sqlParameter.Add("CreateUserID", MainWindow.CurrentUser);

                        Procedure pro1 = new Procedure();
                        pro1.Name = "xp_Code_iMcPart";
                        pro1.OutputUseYN = "Y";
                        pro1.OutputName = "sMCPartID";
                        pro1.OutputLength = "5";

                        Prolist.Add(pro1);
                        ListParameter.Add(sqlParameter);


                        //선택 제품 추가 후 저장하는 프로시저!!
                        for (int i = 0; i < dgdSelProduct.Items.Count; i++)
                        {
                            var winSel = dgdSelProduct.Items[i] as Win_prd_MCTool_U_Product_CodeView;

                            if (winSel != null)
                            {
                                sqlParameter = new Dictionary<string, object>();
                                sqlParameter.Clear();
                                sqlParameter.Add("sMCPartID", strID);
                                sqlParameter.Add("sArticleID", winSel.ArticleID);
                                sqlParameter.Add("UserID", MainWindow.CurrentUser);

                                Procedure pro2 = new Procedure();
                                pro2.Name = "xp_Code_iMcPartArtice";
                                pro2.OutputUseYN = "N";
                                pro2.OutputName = "sMCPartID";
                                pro2.OutputLength = "10";

                                Prolist.Add(pro2);
                                ListParameter.Add(sqlParameter);
                            }
                        }


                        //선택 공정/호기 추가 후 저장하는 프로시저!!
                        for (int i = 0; i < dgdselProcess.Items.Count; i++)
                        {
                            var winSel2 = dgdselProcess.Items[i] as Win_prd_MCTool_U_Process_CodeView;

                            if (winSel2 != null)
                            {
                                sqlParameter = new Dictionary<string, object>();
                                sqlParameter.Clear();
                                sqlParameter.Add("sMCPartID", strID);
                                sqlParameter.Add("sMCID", winSel2.MCID);
                                sqlParameter.Add("UserID", MainWindow.CurrentUser);

                                Procedure pro3 = new Procedure();
                                pro3.Name = "xp_Code_iMcPartMC";
                                pro3.OutputUseYN = "N";
                                pro3.OutputName = "sMCPartID";
                                pro3.OutputLength = "10";

                                Prolist.Add(pro3);
                                ListParameter.Add(sqlParameter);
                            }
                        }

                        List<KeyValue> list_Result = new List<KeyValue>();
                        list_Result = DataStore.Instance.ExecuteAllProcedureOutputGetCS_NewLog(Prolist, ListParameter,"C");
                        string sGetID = string.Empty;

                        if (list_Result[0].key.ToLower() == "success")
                        {
                            list_Result.RemoveAt(0);
                            for (int i = 0; i < list_Result.Count; i++)
                            {
                                KeyValue kv = list_Result[i];
                                if (kv.key == "sMCPartID")
                                {
                                    sGetID = kv.value;
                                    GetKey = sGetID;
                                    flag = true;
                                }
                            }

                            if (flag)
                            {
                                FTP_Save_File(listFtpFile, sGetID);
                            }
                        }
                        else
                        {
                            MessageBox.Show("[저장실패]\r\n" + list_Result[0].value.ToString());
                            flag = false;
                            //return false;
                        }
                    }

                    #endregion

                    #region 수정

                    else if (strFlag.Equals("U"))
                    {
                        sqlParameter.Add("UpdateUserID", MainWindow.CurrentUser);

                        Procedure pro1 = new Procedure();
                        pro1.Name = "xp_Code_uMcPart";
                        pro1.OutputUseYN = "N";
                        pro1.OutputName = "sMCPartID";
                        pro1.OutputLength = "10";

                        Prolist.Add(pro1);
                        ListParameter.Add(sqlParameter);

                        //선택 제품 추가 후 저장하는 프로시저!!를 수정에서 한번 더 
                        for (int i = 0; i < dgdSelProduct.Items.Count; i++)
                        {
                            DataGridRow dgr = Lib.Instance.GetRow(i, dgdSelProduct);
                            var winSelProductUpdate = dgr.Item as Win_prd_MCTool_U_Product_CodeView;

                            sqlParameter = new Dictionary<string, object>();
                            sqlParameter.Clear();
                            sqlParameter.Add("sMCPartID", strID);
                            sqlParameter.Add("sArticleID", winSelProductUpdate.ArticleID);
                            sqlParameter.Add("UserID", MainWindow.CurrentUser);

                            Procedure pro2 = new Procedure();
                            pro2.Name = "xp_Code_iMcPartArtice";
                            pro2.OutputUseYN = "N";
                            pro2.OutputName = "sMCPartID";
                            pro2.OutputLength = "10";

                            Prolist.Add(pro2);
                            ListParameter.Add(sqlParameter);
                        }

                        //선택 공정/호기 추가 후 저장하는 프로시저!!를 수정에서 한번 더 
                        for (int i = 0; i < dgdselProcess.Items.Count; i++)
                        {
                            DataGridRow dgr = Lib.Instance.GetRow(i, dgdselProcess);
                            var winSelProcess = dgr.Item as Win_prd_MCTool_U_Process_CodeView;

                            sqlParameter = new Dictionary<string, object>();
                            sqlParameter.Clear();
                            sqlParameter.Add("sMCPartID", strID);
                            sqlParameter.Add("sMCID", winSelProcess.MCID);
                            sqlParameter.Add("UserID", MainWindow.CurrentUser);

                            Procedure pro3 = new Procedure();
                            pro3.Name = "xp_Code_iMcPartMC";
                            pro3.OutputUseYN = "N";
                            pro3.OutputName = "sMCPartID";
                            pro3.OutputLength = "10";

                            Prolist.Add(pro3);
                            ListParameter.Add(sqlParameter);
                        }

                        ////전체 공정/호기를 못불러와서 다시 한번 프로시저 돌림
                        //for (int i = 0; i < dgdAllProcess.Items.Count; i++)
                        //{
                        //    DataGridRow dgr = Lib.Instance.GetRow(i, dgdAllProcess);
                        //    var winAllProcess = dgr.Item as Win_prd_MCTool_U_Process_CodeView;

                        //    sqlParameter = new Dictionary<string, object>();
                        //    sqlParameter.Clear();
                        //    sqlParameter.Add("sMCPartID", strID);

                        //    Procedure pro4 = new Procedure();
                        //    pro4.Name = "xp_Code_sMcPartAllMC";
                        //    pro4.OutputUseYN = "N";
                        //    pro4.OutputName = "sMCPartID";
                        //    pro4.OutputLength = "10";

                        //    Prolist.Add(pro4);
                        //    ListParameter.Add(sqlParameter);
                        //}

                        ////전체 제품을 못불러와서 다시 한번 프로시저 돌림
                        //for (int i = 0; i < dgdAllProduct.Items.Count; i++)
                        //{
                        //    DataGridRow dgr = Lib.Instance.GetRow(i, dgdAllProduct);
                        //    var winAllProduct = dgr.Item as Win_prd_MCTool_U_Product_CodeView;

                        //    sqlParameter = new Dictionary<string, object>();
                        //    sqlParameter.Clear();
                        //    sqlParameter.Add("sMCPartID", strID);

                        //    Procedure pro5 = new Procedure();
                        //    pro5.Name = "xp_Code_sMcPartAllArtice";
                        //    pro5.OutputUseYN = "N";
                        //    pro5.OutputName = "sMCPartID";
                        //    pro5.OutputLength = "10";

                        //    Prolist.Add(pro5);
                        //    ListParameter.Add(sqlParameter);
                        //}

                        string[] Confirm = new string[2];
                        Confirm = DataStore.Instance.ExecuteAllProcedureOutputNew_NewLog(Prolist, ListParameter,"U");
                        if (Confirm[0] != "success")
                        {
                            MessageBox.Show("[저장실패]\r\n" + Confirm[1].ToString());
                            flag = false;
                            //return false;
                        }
                        else
                        {
                            flag = true;
                            GetKey = strID;
                            //FTP_Save_File(listFtpFile, strID);
                        }
                    }

                    // 파일을 올리자 : GetKey != "" 라면 파일을 올려보자
                    if (!GetKey.Trim().Equals(""))
                    {
                        if (deleteListFtpFile.Count > 0)
                        {
                            foreach (string[] str in deleteListFtpFile)
                            {
                                FTP_RemoveFile(GetKey + "/" + str[0]);
                            }
                        }

                        if (listFtpFile.Count > 0)
                        {
                            FTP_Save_File(listFtpFile, GetKey);
                            //UpdateDBFtp(GetKey);
                        }

                    }

                    // 파일 List 비워주기
                    listFtpFile.Clear();
                    deleteListFtpFile.Clear();


                    #endregion
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

        #region 유효성검사 CheckData

        /// <summary>
        /// 입력사항 체크
        /// </summary>
        /// <returns></returns>
        private bool CheckData()
        {
            bool flag = true;

            if (txtMcPartName.Text.Length <= 0)
            {
                MessageBox.Show(lblMcPartName.Text + "을 입력해주세요.");
                return false;
            }

            if (strFlag.Equals("I"))
            {
                try
                {
                    DataSet ds = null;
                    Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                    sqlParameter.Clear();
                    sqlParameter.Add("sMCPartName", txtMcPartName.Text);
                    sqlParameter.Add("sIncNotUse", 1);

                    ds = DataStore.Instance.ProcedureToDataSet("xp_Code_sMcPartbyName", sqlParameter, false);

                    if (ds != null && ds.Tables.Count > 0)
                    {
                        DataTable dt = ds.Tables[0];

                        if (dt.Rows.Count > 0)
                        {
                            DataRowCollection drc = dt.Rows;
                            if (drc.Count > 0)
                            {
                                MessageBox.Show("동일한 이름의 품명이 " + drc[0]["McPartID"].ToString() + " 에 있습니다.");
                                flag = false;
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

            // 적정재고량
            if (CheckConvertDouble(txtNeedStockQty.Text) == false)
            {
                MessageBox.Show("적정재고량은 숫자만 입력 가능합니다.");
                return false;
            }

            // 한계수명 2021-05-27 한계수명 체크 안함
            //if (CheckConvertDouble(txtSetProdQtyt.Text) == false)
            //{
            //    MessageBox.Show("한계수명은은 숫자만 입력 가능합니다.");
            //    return false;
            //}

            return flag;
        }

        #endregion

        #region FTP 메서드 모음

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

        private BitmapImage SetImage(string ImageName, string FolderName)
        {
            bool ExistFile = false;
            BitmapImage bit = new BitmapImage();

            try
            {
                _ftp = new FTP_EX(FTP_ADDRESS, FTP_ID, FTP_PASS);
                if (_ftp == null) { return null; }

                string[] fileListDetail;
                fileListDetail = _ftp.directoryListSimple(FolderName, Encoding.Default);

                ExistFile = FileInfoAndFlag(fileListDetail, ImageName);
                if (ExistFile)
                {
                    bit = _ftp.DrawingImageByByte(FTP_ADDRESS + '/' + FolderName + '/' + ImageName + "");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("파일을 찾을 수 없습니다.");
            }


            return bit;
        }

        //FTP 업로드시 파일체크 및 경로,파일이름 표시
        private TextBox Ftp_Upload_TextBox()
        {
            TextBox tb = new TextBox();
            string[] strTemp = null;
            Microsoft.Win32.OpenFileDialog OFdlg = new Microsoft.Win32.OpenFileDialog();
            //OFdlg.Filter = "Image files (*.jpg, *.jpeg, *.jpe, *.jfif, *.png, *.pcx, *.pdf) | *.jpg; *.jpeg; *.jpe; *.jfif; *.png; *.pcx; *.pdf | All Files|*.*";
            OFdlg.Filter = MainWindow.OFdlg_Filter;

            Nullable<bool> result = OFdlg.ShowDialog();
            if (result == true)
            {
                strFullPath = OFdlg.FileName;

                string ImageFileName = OFdlg.SafeFileName;  //명.
                string ImageFilePath = string.Empty;       // 경로

                ImageFilePath = strFullPath.Replace(ImageFileName, "");

                StreamReader sr = new StreamReader(OFdlg.FileName);
                long FileSize = sr.BaseStream.Length;
                if (sr.BaseStream.Length > (2048 * 1000))
                {
                    //업로드 파일 사이즈범위 초과
                    //MessageBox.Show("이미지의 파일사이즈가 2M byte를 초과하였습니다.");
                    sr.Close();
                    tb.Text = "파일사이즈초과";
                    //return;
                }
                else
                {
                    tb.Text = ImageFileName;
                    tb.Tag = ImageFilePath;
                }

                strTemp = new string[] { ImageFileName, ImageFilePath.ToString() };
                listFtpFile.Add(strTemp);
            }

            return tb;
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

            if (MakeFolder == false)        // 같은 아이를 찾지 못한경우,
            {
                //MIL 폴더에 InspectionID로 저장
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

        //파일 다운로드
        private void FTP_DownLoadFile(string FilePath, string FileName)
        {
            try
            {
                string str_path = FTP_ADDRESS + '/' + FilePath;     //풀 경로.

                _ftp = new FTP_EX(str_path, FTP_ID, FTP_PASS);

                string str_remotepath = FileName;
                string str_localpath = str_localpath = LOCAL_DOWN_PATH + "\\" + FileName;

                DirectoryInfo DI = new DirectoryInfo(LOCAL_DOWN_PATH);      // Temp 폴더가 없는 컴터라면, 만들어 줘야지.
                if (DI.Exists == false)
                {
                    DI.Create();
                }

                FileInfo file = new FileInfo(str_localpath);
                if (file.Exists)
                {
                    file.Delete();
                }

                _ftp.download(str_remotepath, str_localpath);

                ProcessStartInfo proc = new ProcessStartInfo(str_localpath);
                proc.UseShellExecute = true;
                Process.Start(proc);
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

        //FTP 업로드시 파일체크 및 경로,파일이름 표시
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
                //OFdlg.Filter = "Image files (*.jpg, *.jpeg, *.jpe, *.jfif, *.png, *.pcx, *.pdf) | *.jpg; *.jpeg; *.jpe; *.jfif; *.png; *.pcx; *.pdf | All Files|*.*";
                //OFdlg.Filter = "Image files (*.jpg, *.jpeg, *.jpe, *.jfif, *.png) | *.jpg; *.jpeg; *.jpe; *.jfif; *.png;";
                OFdlg.Filter = MainWindow.OFdlg_Filter;

                Nullable<bool> result = OFdlg.ShowDialog();
                if (result == true)
                {
                    strFullPath = OFdlg.FileName;

                    string ImageFileName = OFdlg.SafeFileName;  //명.
                    string ImageFilePath = string.Empty;       // 경로

                    ImageFilePath = strFullPath.Replace(ImageFileName, "");

                    StreamReader sr = new StreamReader(OFdlg.FileName);
                    long FileSize = sr.BaseStream.Length;
                    if (sr.BaseStream.Length > (2048 * 1000))
                    {
                        //업로드 파일 사이즈범위 초과
                        MessageBox.Show("이미지의 파일사이즈가 2M byte를 초과하였습니다.");
                        sr.Close();
                        return;
                    }
                    else
                    {
                        textBox.Text = ImageFileName;
                        textBox.Tag = ImageFilePath;

                        try
                        {
                            Bitmap image = new Bitmap(ImageFilePath + ImageFileName);
                            imgPart.Source = BitmapToImageSource(image);
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show("해당 파일은 이미지로 변환이 불가능합니다.");
                        }

                        string[] strTemp = new string[] { ImageFileName, ImageFilePath.ToString() };
                        listFtpFile.Add(strTemp);
                    }
                }
            }
        }

        // 비트맵을 비트맵 이미지로 형태변환시키기.<0823 허윤구> 
        BitmapImage BitmapToImageSource(Bitmap bitmap)
        {
            using (MemoryStream memory = new MemoryStream())
            {
                bitmap.Save(memory, System.Drawing.Imaging.ImageFormat.Bmp);
                memory.Position = 0;
                BitmapImage bitmapimage = new BitmapImage();
                bitmapimage.BeginInit();
                bitmapimage.StreamSource = memory;
                bitmapimage.CacheOption = BitmapCacheOption.OnLoad;
                bitmapimage.EndInit();

                return bitmapimage;
            }
        }

        #endregion

        #region AllProduct 전체 제품 데이터그리드

        private void FillGridAllProduct()
        {

            try
            {
                if (dgdAllProduct.Items.Count > 0)
                {
                    dgdAllProduct.Items.Clear();
                }

                DataSet ds = null;
                Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                sqlParameter.Clear();
                sqlParameter.Add("sMCPartID", txtMcPartCode.Text);
                ds = DataStore.Instance.ProcedureToDataSet("xp_Code_sMcPartAllArtice", sqlParameter, false);

                if (ds != null && ds.Tables.Count > 0)
                {
                    DataTable dt = ds.Tables[0];
                    int i = 0;

                    if (dt.Rows.Count == 0)
                    {
                        //MessageBox.Show("조회된 데이터가 없습니다.");
                    }
                    else
                    {
                        DataRowCollection drc = dt.Rows;

                        foreach (DataRow dr in drc)
                        {
                            var WinAllProduct = new Win_prd_MCTool_U_Product_CodeView()
                            {
                                Num = i + 1,
                                ArticleID = dr["ArticleID"].ToString(),
                                Article = dr["Article"].ToString(),
                                BuyerArticleNo = dr["BuyerArticleNo"].ToString(),
                            };

                            dgdAllProduct.Items.Add(WinAllProduct);
                            i++;
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

        #region selProduct 선택 제품 데이터그리드

        private void FillGrid_selProduct(string MCPartID)
        {
            try
            {
                if (dgdSelProduct.Items.Count > 0)
                {
                    dgdSelProduct.Items.Clear();
                }

                DataSet ds = null;
                Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                sqlParameter.Clear();
                sqlParameter.Add("sMCPartID", MCPartID);

                ds = DataStore.Instance.ProcedureToDataSet("xp_Code_sMcPartArtice", sqlParameter, false);

                if (ds != null && ds.Tables.Count > 0)
                {
                    DataTable dt = ds.Tables[0];
                    int i = 0;
                    string strSelProduct = string.Empty;

                    if (dt.Rows.Count == 0)
                    {
                        //MessageBox.Show("조회된 데이터가 없습니다.");
                    }
                    else
                    {
                        DataRowCollection drc = dt.Rows;

                        foreach (DataRow dr in drc)
                        {
                            var WinSelProduct = new Win_prd_MCTool_U_Product_CodeView()
                            {
                                Num = i + 1,
                                ArticleID = dr["ArticleID"].ToString(),
                                Article = dr["Article"].ToString(),
                                BuyerArticleNo = dr["BuyerArticleNo"].ToString(),
                            };

                            dgdSelProduct.Items.Add(WinSelProduct);
                            i++;

                            if (i == drc.Count)
                            {
                                strSelProduct += WinSelProduct.Article;
                            }
                            else
                            {
                                strSelProduct += WinSelProduct.Article + "→";
                            }
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

        #region AllProcess 전체 공정/호기 데이터그리드

        private void FillGridAllProcess()
        {
            try
            {
                if (dgdAllProcess.Items.Count > 0)
                {
                    dgdAllProcess.Items.Clear();
                }

                DataSet ds = null;
                Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                sqlParameter.Clear();
                sqlParameter.Add("sMCPartID", txtMcPartCode.Text);
                ds = DataStore.Instance.ProcedureToDataSet("xp_Code_sMcPartAllMC", sqlParameter, false);

                if (ds != null && ds.Tables.Count > 0)
                {
                    DataTable dt = ds.Tables[0];
                    int i = 0;

                    if (dt.Rows.Count == 0)
                    {
                        //MessageBox.Show("조회된 데이터가 없습니다.");
                    }
                    else
                    {
                        DataRowCollection drc = dt.Rows;

                        foreach (DataRow dr in drc)
                        {
                            var WinAllProcess = new Win_prd_MCTool_U_Process_CodeView()
                            {
                                Num = i + 1,
                                MCID = dr["MCID"].ToString(),
                                MCNAME = dr["MCNAME"].ToString()
                            };

                            dgdAllProcess.Items.Add(WinAllProcess);
                            i++;
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

        #region selProcess 선택 공정/호기 데이터그리드

        private void FillGrid_selProcess(string MCPartID)
        {
            try
            {
                if (dgdselProcess.Items.Count > 0)
                {
                    dgdselProcess.Items.Clear();
                }

                DataSet ds = null;
                Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                sqlParameter.Clear();
                sqlParameter.Add("sMCPartID", MCPartID);

                ds = DataStore.Instance.ProcedureToDataSet("xp_Code_sMcPartMC", sqlParameter, false);

                if (ds != null && ds.Tables.Count > 0)
                {
                    DataTable dt = ds.Tables[0];
                    int i = 0;
                    string strSelProcess = string.Empty;

                    if (dt.Rows.Count == 0)
                    {
                        //MessageBox.Show("조회된 데이터가 없습니다.");
                    }
                    else
                    {
                        DataRowCollection drc = dt.Rows;

                        foreach (DataRow dr in drc)
                        {
                            var WinSelProcess = new Win_prd_MCTool_U_Process_CodeView()
                            {
                                Num = i + 1,
                                MCID = dr["MCID"].ToString(),
                                MCNAME = dr["MCNAME"].ToString(),

                            };

                            dgdselProcess.Items.Add(WinSelProcess);
                            i++;

                            if (i == drc.Count)
                            {
                                strSelProcess += WinSelProcess.MCID;
                            }
                            else
                            {
                                strSelProcess += WinSelProcess.MCNAME + "→";
                            }
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

        #region 오른쪽 하단 제품, 공정/호기 선택 데이터그리드 이벤트

        private void btnRight_Click(object sender, RoutedEventArgs e)
        {
            var winAddAllP = dgdAllProduct.SelectedItem as Win_prd_MCTool_U_Product_CodeView;
            bool flag = true;

            if (winAddAllP != null)
            {
                for (int i = 0; i < dgdSelProduct.Items.Count; i++)
                {
                    var WinSelAdd = dgdSelProduct.Items[i] as Win_prd_MCTool_U_Product_CodeView;

                    if (WinSelAdd != null
                        && WinSelAdd.ArticleID != null
                        && WinSelAdd.ArticleID.Equals(winAddAllP.ArticleID)) 
                    {
                        flag = false;
                        break;
                    }
                }

                if (flag)
                {
                    winAddAllP.Num = dgdSelProduct.Items.Count + 1;

                    // 추가할 데이터 selProduct에 추가
                    dgdSelProduct.Items.Add(winAddAllP);
                    
                }
                else
                {
                    MessageBox.Show("이미 추가되어 있는 제품입니다.");
                }
            }
        }

        private void btnLeft_Click(object sender, RoutedEventArgs e)
        {
            var winsubProduct = dgdSelProduct.SelectedItem as Win_prd_MCTool_U_Product_CodeView;

            if(winsubProduct != null)
            {
                dgdSelProduct.Items.Remove(winsubProduct);

                if (dgdSelProduct.Items.Count > 0)
                {
                    // 제외 후 selProduct Num 재정렬
                    for (int i = 0; i < dgdSelProduct.Items.Count; i++)
                    {
                        var WinSelP = dgdSelProduct.Items[i] as Win_prd_MCTool_U_Product_CodeView;

                        if (WinSelP != null)
                        {
                            WinSelP.Num = i + 1;
                        }
                    }
                }
            }
            else
            {
                MessageBox.Show("제외할 제품이 선택되지 않았습니다.");
            }
        }

        private void btnRight2_Click(object sender, RoutedEventArgs e)
        {
            var winAddAllP = dgdAllProcess.SelectedItem as Win_prd_MCTool_U_Process_CodeView;
            bool flag = true;

            if (winAddAllP != null)
            {
                for (int i = 0; i < dgdselProcess.Items.Count; i++)
                {
                    var WinPAdd = dgdselProcess.Items[i] as Win_prd_MCTool_U_Process_CodeView;

                    if (WinPAdd != null
                        && WinPAdd.MCID != null
                        && WinPAdd.MCID.Equals(winAddAllP.MCID))
                    {
                        flag = false;
                        break;
                    }
                }

                if (flag)
                {

                    winAddAllP.Num = dgdselProcess.Items.Count + 1;

                    dgdselProcess.Items.Add(winAddAllP);
                }
                else
                {
                    MessageBox.Show("이미 추가되어 있는 공정입니다.");
                }
            }
        }
               
        private void btnLeft2_Click(object sender, RoutedEventArgs e)
        {
            var winsubProcess = dgdselProcess.SelectedItem as Win_prd_MCTool_U_Process_CodeView;

            bool flag = true;

            if (winsubProcess != null)
            {
                dgdselProcess.Items.Remove(winsubProcess);

                if (dgdselProcess.Items.Count > 0)
                {
                    // 제외 후 selProcess Num 재정렬
                    for (int i = 0; i < dgdselProcess.Items.Count; i++)
                    {
                        var WinSel = dgdselProcess.Items[i] as Win_prd_MCTool_U_Process_CodeView;

                        if (WinSel != null)
                        {
                            WinSel.Num = i + 1;
                        }
                    }
                }
            }
        }

        private void dgtcAllProduct_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            if (e.ClickCount == 2)
            {
                btnRight_Click(null, null);
            }
        }

        private void dgtcSelProduct_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            if (e.ClickCount == 2)
            {
                btnLeft_Click(null, null);
            }
        }

        private void dgtcAllMC_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            if (e.ClickCount == 2)
            {
                btnRight2_Click(null, null);
            }
        }

        private void dgtcSelMC_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            if (e.ClickCount == 2)
            {
                btnLeft2_Click(null, null);
            }
        }

        #endregion

        #region 왼쪽 데이터그리드 체크박스 이벤트

        //전체선택 체크박스 체크 이벤트
        private void CheckAll_Click(object sender, RoutedEventArgs e)
        {
            if (dgdMain.Items.Count > 0)
            {
                for (int i = 0; i < dgdMain.Items.Count; i++)
                {
                    var Check = dgdMain.Items[i] as Win_prd_MCTool_U_CodeView;
                    if (Check != null)
                    {
                        Check.Flag = true;
                    }
                }
            }
        }

        //전체선택 체크박스 해제 이벤트
        private void UnCheckAll_Click(object sender, RoutedEventArgs e)
        {
            if (dgdMain.Items.Count > 0)
            {
                for (int i = 0; i < dgdMain.Items.Count; i++)
                {
                    var Check = dgdMain.Items[i] as Win_prd_MCTool_U_CodeView;
                    if (Check != null)
                    {
                        Check.Flag = false;
                    }
                }
            }
        }

        //단일 체크박스 체크 이벤트
        private void cbxCheck_click(object sender, RoutedEventArgs e)
        {
            WinPartCode = dgdMain.CurrentItem as Win_prd_MCTool_U_CodeView;

            if (WinPartCode != null)
            {
                if (WinPartCode.Flag)
                    WinPartCode.Flag = false;
                else
                    WinPartCode.Flag = true;
            }
        }




        #endregion

        #region 이미지 버튼 이벤트 모음

        //업로드할 이미지 선택
        private void BtnImgSelect_Click(object sender, RoutedEventArgs e)
        {
            if (!txtImage.Text.Equals(string.Empty) && strFlag.Equals("U"))
            {
                MessageBox.Show("먼저 해당파일의 삭제를 진행 후 진행해주세요.");
                return;
            }
            else
            {
                FTP_Upload_TextBox(txtImage);
            }
        }

        // 이미지 보기 버튼 클릭 → 다운로드
        private void BtnImageView_Click(object sender, RoutedEventArgs e)
        {
            WinPartCode = dgdMain.SelectedItem as Win_prd_MCTool_U_CodeView;

            if (WinPartCode != null && !WinPartCode.ImageName.Equals(""))
            {
                FTP_DownLoadFile(WinPartCode.MCPartID, WinPartCode.ImageName);
            }
        }

        // 이미지 삭제
        private void BtnImageDelete_Click(object sender, RoutedEventArgs e)
        {
            if (MessageBox.Show("해당 이미지를 삭제하시겠습니까?", "삭제 전 확인", MessageBoxButton.YesNo) == MessageBoxResult.Yes)
            {
                txtImage.Text = "";
                txtImage.Tag = "";

                imgPart.Source = null;
            }
        }

        #endregion

        #region 기타 메서드 모음

        //천 단위 구분기호, 소수점 자릿수 0 
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

        // 확장자 이미지 확인하기, 메인윈도우에 확장자 리스트 세팅
        private bool CheckImage(string ImageName)
        {
            string[] extensions = MainWindow.Extensions;

            bool flag = false;

            ImageName = ImageName.Trim().ToLower();
            foreach (string ext in extensions)
            {
                if (ImageName.EndsWith(ext))
                {
                    flag = true;
                }
            }

            return flag;
        }

        // 텍스트 박스 숫자만 입력 되도록
        public void CheckIsNumericOnly(TextBox sender, TextCompositionEventArgs e)
        {
            decimal result;
            if (!(Decimal.TryParse(e.Text, out result)))
            {
                e.Handled = true;
            }
        }

        #endregion

        #region 텍스트박스 공용 메서드 모음 : 설비명 엔터 → 검색 / 숫자만 입력 가능하도록


        private void txtBox_EnterAndSearch(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                rowNum = 0;
                re_Search(rowNum);
            }
        }

        // 텍스트박스 숫자만 입력 가능하도록
        private void txtBox_PreviewTextInput(object sender, TextCompositionEventArgs e)
        {
            CheckIsNumericOnly((TextBox)sender, e);
        }

        #endregion

        #region 토글 버튼 이벤트 모음

        // 공용 버튼
        private void tgnCommon_Checked(object sender, RoutedEventArgs e)
        {
            tgnCommon.IsChecked = true;
            tgnSpare.IsChecked = false;
            tgnTool.IsChecked = false;
            tgnAll.IsChecked = false;

            dgtc_MCPartName.Header = "설비(부품)명";
            lblMcPartName.Text = "설비(부품)명";

            re_Search(0);
        }

        private void tgnCommon_Unchecked(object sender, RoutedEventArgs e)
        {
            e.Handled = true;
        }

        private void tgnCommon_PreviewMouseDown(object sender, MouseButtonEventArgs e)
        {
            if (tgnCommon.IsChecked == true)
            {
                e.Handled = true;
            }
        }

        // 설비예비품
        private void tgnSpare_Checked(object sender, RoutedEventArgs e)
        {
            tgnCommon.IsChecked = false;
            tgnSpare.IsChecked = true;
            tgnTool.IsChecked = false;
            tgnAll.IsChecked = false;

            dgtc_MCPartName.Header = "설비(부품)명";
            lblMcPartName.Text = "설비(부품)명";

            re_Search(0);
        }

        private void tgnSpare_Unchecked(object sender, RoutedEventArgs e)
        {
            e.Handled = true;
        }

        private void tgnSpare_PreviewMouseDown(object sender, MouseButtonEventArgs e)
        {
            if (tgnSpare.IsChecked == true)
            {
                e.Handled = true;
            }
        }

        // 툴
        private void tgnTool_Checked(object sender, RoutedEventArgs e)
        {
            tgnCommon.IsChecked = false;
            tgnSpare.IsChecked = false;
            tgnTool.IsChecked = true;
            tgnAll.IsChecked = false;

            dgtc_MCPartName.Header = "Tool 명";
            lblMcPartName.Text = "Tool 명";

            re_Search(0);
        }

        private void tgnTool_Unchecked(object sender, RoutedEventArgs e)
        {
            e.Handled = true;
        }

        private void tgnTool_PreviewMouseDown(object sender, MouseButtonEventArgs e)
        {
            if (tgnTool.IsChecked == true)
            {
                e.Handled = true;
            }
        }

        // 전체
        private void tgnAll_Checked(object sender, RoutedEventArgs e)
        {
            tgnCommon.IsChecked = false;
            tgnSpare.IsChecked = false;
            tgnTool.IsChecked = false;
            tgnAll.IsChecked = true;

            dgtc_MCPartName.Header = "설비(부품)명";
            lblMcPartName.Text = "설비(부품)명";

            re_Search(0);
        }

        private void tgnAll_Unchecked(object sender, RoutedEventArgs e)
        {
            e.Handled = true;
        }

        private void tgnAll_PreviewMouseDown(object sender, MouseButtonEventArgs e)
        {
            if (tgnAll.IsChecked == true)
            {
                e.Handled = true;
            }
        }



        #endregion

        #region 입력창 이동 이벤트

        private void txtMcPartName_KeyDown(object sender, KeyEventArgs e)
        {
            if(e.Key == Key.Enter)
            {
                cboForUse.Focus();
            }
        }

        private void cboForUse_DropDownClosed(object sender, EventArgs e)
        {
            cboUnitClss.Focus();
        }

        private void cboForUse_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                cboUnitClss.Focus();
            }
        }

        private void cboUnitClss_DropDownClosed(object sender, EventArgs e)
        {
            txtUnitPrice.Focus();
        }

        private void cboUnitClss_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                txtUnitPrice.Focus();
            }
        }

        private void txtUnitPrice_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                txtNeedStockQty.Focus();
            }
        }

        private void txtNeedStockQty_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                txtSetProdQtyt.Focus();
            }
        }

        private void txtSetProdQtyt_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                txtSpec.Focus();
            }
        }

        #endregion


    }

    class Win_prd_MCTool_U_CodeView : BaseView
    {
        public int Num { get; set; }
        public bool Flag { get; set; }

        public string MCPartID { get; set; }
        public string MCPartName { get; set; }
        public string Weight { get; set; }
        public string Spec { get; set; }
        public string ImageName { get; set; }

        public string UseClss { get; set; }
        public string UnitClssName { get; set; }
        public string NeedStockQty { get; set; }
        public string UnitClss { get; set; }
        public string ForUse { get; set; }

        public string SetProdQty { get; set; }
        public string CreateDate { get; set; }
        public string CreateDate_CV { get; set; }
        public string ForUseName { get; set; }
        public string UseClssChar { get; set; }        
        public string ManageNo { get; set; } //삼주 회의시 추가요청
        public Double UnitPrice { get; set; }
    }

    class Win_prd_MCTool_U_Product_CodeView : BaseView
    {
       public int Num { get; set; }
       public string ArticleID { get; set; }
       public string Article { get; set; }
       public string BuyerArticleNo { get; set; }
    }
    
    class Win_prd_MCTool_U_Process_CodeView : BaseView
    {
        public int Num { get; set; }
        public string MCID { get; set; }
        public string MCNAME { get; set; }
    }

}

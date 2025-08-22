using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Data;
using System.Linq;
using System.Threading;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using WizMes_HanYoung.PopUP;

namespace WizMes_HanYoung
{
    /// <summary>
    /// Win_com_ArticleConditioin_U.xaml에 대한 상호 작용 논리
    /// </summary>
    public partial class Win_com_ArticleConditioin_U : UserControl
    {
        string stDate = string.Empty;
        string stTime = string.Empty;

        string getCodeName = string.Empty;

        Lib lib = new Lib();
        PlusFinder pf = new PlusFinder();

        // 인쇄 활용 용도 (프린트)
        private Microsoft.Office.Interop.Excel.Application excelapp;
        private Microsoft.Office.Interop.Excel.Workbook workbook;
        private Microsoft.Office.Interop.Excel.Worksheet worksheet;
        private Microsoft.Office.Interop.Excel.Range workrange;
        private Microsoft.Office.Interop.Excel.Worksheet copysheet;
        private Microsoft.Office.Interop.Excel.Worksheet pastesheet;
        WizMes_HanYoung.PopUp.NoticeMessage msg = new WizMes_HanYoung.PopUp.NoticeMessage();

        List<Win_com_ArticleConditioin_U_CodeView> lstOutwarePrint = new List<Win_com_ArticleConditioin_U_CodeView>();
        Win_com_ArticleConditioin_U_Sub_CodeView ComboboxSub = new Win_com_ArticleConditioin_U_Sub_CodeView();


        // 수정 정보를 보관하기 위한 변수
        //List<Win_com_ArticleConditioin_U_Sub_CodeView> lstBoxID = new List<Win_com_ArticleConditioin_U_Sub_CodeView>();
        List<Win_com_ArticleConditioin_U_Sub_CodeView> ListOutwareSub = new List<Win_com_ArticleConditioin_U_Sub_CodeView>();

        int rowNum = 0;                          // 조회시 데이터 줄 번호 저장용도
        string strFlag = string.Empty;           // 추가, 수정 구분 
        string GetKey = "";


        int rowSubNum = 0;

        List<string> LabelGroupList = new List<string>();         // packing ID 스캔에 따른 LabelID를 모아 담을 리스트 그릇입니다.


        public Win_com_ArticleConditioin_U()
        {
            InitializeComponent();
        }

        private void UserControl_Loaded(object sender, RoutedEventArgs e)
        {
            try
            {
                stDate = DateTime.Now.ToString("yyyyMMdd");
                stTime = DateTime.Now.ToString("HHmm");

                DataStore.Instance.InsertLogByFormS(this.GetType().Name, stDate, stTime, "S");
                

                CantBtnControl();
                SetComboBox();
            }
            catch (Exception ee)
            {
                MessageBox.Show("오류지점 - UserControl_Loaded : " + ee.ToString());
            }
        }

        #region 콤보박스
        private void SetComboBox()
        {

        }
        #endregion 콤보박스

        #region 상단 레이아웃 조건 모음

        // 품명
        // 품명 검색 라벨 왼쪽 클릭 이벤트
        private void lblArticleSrh_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            if (chkArticleSrh.IsChecked == true)
            {
                chkArticleSrh.IsChecked = false;
            }
            else
            {
                chkArticleSrh.IsChecked = true;
            }
        }
        // 품명 검색 체크박스 이벤트
        private void chkArticleSrh_Checked(object sender, RoutedEventArgs e)
        {
            chkArticleSrh.IsChecked = true;

            txtArticleSrh.IsEnabled = true;
            btnPfArticleSrh.IsEnabled = true;
        }
        private void chkArticleSrh_Unchecked(object sender, RoutedEventArgs e)
        {
            chkArticleSrh.IsChecked = false;

            txtArticleSrh.IsEnabled = false;
            btnPfArticleSrh.IsEnabled = false;
        }
        // 품명 검색 엔터 → 플러스 파인더 이벤트
        private void txtArticleSrh_KeyDown(object sender, KeyEventArgs e)
        {

            if (e.Key == Key.Enter)
            {
                e.Handled = true;
                MainWindow.pf.ReturnCode(txtArticleSrh, 76, "");
            }
        }
        // 품명 검색 플러스파인더 이벤트
        private void btnPfArticleSrh_Click(object sender, RoutedEventArgs e)
        {
            MainWindow.pf.ReturnCode(txtArticleSrh, 76, "");
        }



        #endregion

        #region 버튼 모음
        //유지추가
        private void btnReMainAdd_Click(object sender, RoutedEventArgs e)
        {
            strFlag = "U";

            txtArticle.Text = "";
            txtArticle.Tag = null;
            txtBuyerArticle.Text = "";

            CanBtnControl();

        }

        //추가버튼 클릭
        private void btnAdd_Click(object sender, RoutedEventArgs e)
        {

            try
            {
                strFlag = "I";

                this.DataContext = null;
                txtArticle.Tag = null;
                CanBtnControl();                             //버튼 컨트롤


                //클리어
                if (ListOutwareSub.Count > 0)
                {
                    ListOutwareSub.Clear();
                }

                //클리어
                if (dgdOutwareSub.Items.Count > 0)
                {
                    dgdOutwareSub.Items.Clear();
                }


                //CM_CODE에 등록된 놈들 가져와서 뿌리기
                FillgridSaGathItem();


            }
            catch (Exception ee)
            {
                MessageBox.Show("오류지점 - btnAdd_Click : " + ee.ToString());
            }
        }

        //수정버튼 클릭
        private void btnUpdate_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var OutwareItem = dgdOutware.SelectedItem as Win_com_ArticleConditioin_U_CodeView;

                if (OutwareItem != null)
                {
                    strFlag = "U";

                    rowNum = dgdOutware.SelectedIndex;
                    CanBtnControl();
                }
            }
            catch (Exception ee)
            {
                MessageBox.Show("오류지점 - btnUpdate_Click : " + ee.ToString());
            }
        }

        //삭제버튼 클릭
        private void btnDelete_Click(object sender, RoutedEventArgs e)
        {
            if (dgdOutware.SelectedItem == null)
            {
                MessageBox.Show("삭제할 데이터가 지정되지 않았습니다. 삭제데이터를 지정하고 눌러주세요");
            }
            else
            {

                var Condtion = dgdOutware.SelectedItem as Win_com_ArticleConditioin_U_CodeView;

                if (MessageBox.Show("선택하신 항목을 삭제하시겠습니까?", "삭제 전 확인", MessageBoxButton.YesNo) == MessageBoxResult.Yes)
                {

                    if (dgdOutware.Items.Count > 0 && dgdOutware.SelectedItem != null)
                    {
                        rowNum = dgdOutware.SelectedIndex - 1;
                    }

                    if (DeleteData(Condtion.ArticleID))
                    {
                        rowNum = 0;
                        re_Search(rowNum);
                    }
                }
            }


        }

        //닫기버튼 클릭
        private void btnClose_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                DataStore.Instance.InsertLogByFormS(this.GetType().Name, stDate, stTime, "E");
                Lib.Instance.ChildMenuClose(this.ToString());
            }
            catch (Exception ee)
            {
                MessageBox.Show("오류지점 - btnClose_Click : " + ee.ToString());
            }
        }

        //검색버튼 클릭
        private void btnSearch_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                rowNum = 0;
                re_Search(rowNum);
            }
            catch (Exception ee)
            {
                MessageBox.Show("오류지점 - btnSearch_Click : " + ee.ToString());
            }
        }

        //저장버튼 클릭
        private void btnSave_Click(object sender, RoutedEventArgs e)
        {
            if (txtArticle.Tag == null)
            {
                MessageBox.Show("품번 정보가 입력되지 않았습니다.");
                return;
            }

            if (SaveData(txtArticle.Tag.ToString(), strFlag))
            {
                CantBtnControl();

                re_Search(rowNum);
            }

        }

        //취소버튼 클릭
        private void btnCancel_Click(object sender, RoutedEventArgs e)
        {
            rowNum = 0;

            strFlag = string.Empty;
            CantBtnControl();
            re_Search(rowNum);


        }

        //엑셀버튼 클릭
        private void btnExcel_Click(object sender, RoutedEventArgs e)
        {
            Lib lib2 = new Lib();
            try
            {
                if (dgdOutware.Items.Count < 1)
                {
                    MessageBox.Show("먼저 검색해 주세요.");
                    return;
                }
                DataTable dt = null;
                string Name = string.Empty;

                string[] lst = new string[4];
                lst[0] = "메인그리드";
                lst[1] = "서브그리드";
                lst[2] = dgdOutware.Name;
                lst[3] = dgdOutwareSub.Name;

                ExportExcelxaml ExpExc = new ExportExcelxaml(lst);

                ExpExc.ShowDialog();

                if (ExpExc.DialogResult.HasValue)
                {
                    if (ExpExc.choice.Equals(dgdOutware.Name))
                    {
                        DataStore.Instance.InsertLogByForm(this.GetType().Name, "E");
                        //MessageBox.Show("대분류");
                        if (ExpExc.Check.Equals("Y"))
                            dt = lib2.DataGridToDTinHidden(dgdOutware);
                        else
                            dt = lib2.DataGirdToDataTable(dgdOutware);

                        Name = dgdOutware.Name;
                        if (lib2.GenerateExcel(dt, Name))
                        {
                            lib2.excel.Visible = true;
                            lib2.ReleaseExcelObject(lib2.excel);
                        }
                    }
                    else if (ExpExc.choice.Equals(dgdOutwareSub.Name))
                    {
                        //MessageBox.Show("정성류");
                        if (ExpExc.Check.Equals("Y"))
                            dt = lib2.DataGridToDTinHidden(dgdOutwareSub);
                        else
                            dt = lib2.DataGirdToDataTable(dgdOutwareSub);
                        Name = dgdOutwareSub.Name;
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
            catch (Exception ee)
            {
                MessageBox.Show("오류지점 - btnExcel_Click : " + ee.ToString());
            }
            finally
            {
                lib2 = null;
            }
        }



        #endregion

        #region 플러스파인더 및 데이터그리드 선택 변경

        //메인 데이터그리드 선택 변경
        private void dgdOutware_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            var getArticle = dgdOutware.SelectedItem as Win_com_ArticleConditioin_U_CodeView;

            if (getArticle != null)
            {
                this.DataContext = getArticle;
                txtArticle.Tag = getArticle.ArticleID;
                FillGridSub(getArticle.ArticleID);

                if (dgdOutwareSub.Items.Count > 0)
                {
                    dgdOutwareSub.SelectedIndex = 0;
                }

            }

        }




        //서브 데이터 그리드 키다운 이벤트
        private void dgdOutwareSub_KeyDown(object sender, KeyEventArgs e)
        {
            try
            {
                if (e.Key == Key.Delete)
                {
                    //추가 / 수정 이벤트가 진행중인 경우,
                    if ((btnSave.Visibility == Visibility.Visible) && (btnCancel.Visibility == Visibility.Visible))
                    {
                        var OutwareSub = dgdOutwareSub.SelectedItem as Win_com_ArticleConditioin_U_Sub_CodeView;
                        if (OutwareSub != null)
                        {
                            dgdOutwareSub.Items.Remove(OutwareSub);
                        }
                    }
                }
            }
            catch (Exception ee)
            {
                MessageBox.Show("오류지점 - dgdOutwareSub_KeyDown : " + ee.ToString());
            }
        }

        #endregion

        #region Research
        private void re_Search(int rowNum)
        {
            try
            {

                FillGrid();

                if (dgdOutware.Items.Count > 0)
                {
                    dgdOutware.SelectedIndex = rowNum;
                }

            }
            catch (Exception ee)
            {
                MessageBox.Show("오류지점 - re_Search : " + ee.ToString());
            }
        }

        #endregion

        #region 조회
        private void FillGrid()
        {
            if (dgdOutware.Items.Count > 0)
            {
                dgdOutware.Items.Clear();
            }

            try
            {
                Dictionary<string, object> sqlParameter = new Dictionary<string, object>();

                sqlParameter.Add("nChkArticleID", chkArticleSrh.IsChecked == true ? 1 : 0);
                sqlParameter.Add("sArticleID", chkArticleSrh.IsChecked == true && txtArticleSrh.Tag != null ? (txtArticleSrh.Tag.ToString()) : "");

                DataSet ds = DataStore.Instance.ProcedureToDataSet("xp_Condition_sArticle_Main", sqlParameter, false);

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

                            var dgdCondition = new Win_com_ArticleConditioin_U_CodeView()
                            {
                                Num = i + "",

                                ArticleID = dr["ArticleID"].ToString(),

                                Article = dr["Article"].ToString(),
                                BuyerArticleNo = dr["BuyerArticleNo"].ToString(),
                                ConditionDate = dr["ConditionDate"].ToString()


                            };

                            dgdCondition.ConditionDate = DatePickerFormat(dgdCondition.ConditionDate);
                            dgdOutware.Items.Add(dgdCondition);

                            txtArticle.Tag = dgdCondition.ArticleID;
                        }


                    }
                }
            }
            catch (Exception ee)
            {
                MessageBox.Show(ee.Message);
            }
            finally
            {
                DataStore.Instance.CloseConnection();
            }

        }
        #endregion

        #region Sub조회
        private void FillGridSub(string ArticleID)
        {
            if (dgdOutwareSub.Items.Count > 0)
            {
                dgdOutwareSub.Items.Clear();
            }

            try
            {
                Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                sqlParameter.Add("ArticleID", ArticleID);

                DataSet ds = DataStore.Instance.ProcedureToDataSet("xp_Condition_sArticleID", sqlParameter, false);

                if (ds != null && ds.Tables.Count > 0)
                {
                    DataTable dt = ds.Tables[0];

                    if (dt.Rows.Count > 0)
                    {
                        DataRowCollection drc = dt.Rows;
                        int j = 0;

                        for (int i = 0; i < drc.Count; i++)
                        {
                            j = i;
                            DataRow dr = drc[i];

                            var CondtionSub = new Win_com_ArticleConditioin_U_Sub_CodeView()
                            {

                                CodeID = dr["CodeID"].ToString(),
                                CodeName = dr["CodeName"].ToString(),
                                Spec = dr["Spec"].ToString(),
                                SpecMin = stringFormatN0(dr["SpecMin"]),
                                SpecMax = stringFormatN0(dr["SpecMax"]),

                                SpecTextMin = dr["SpecTextMin"].ToString(),
                                SpecTextMax = dr["SpecTextMax"].ToString(),
                                WarningSpecMin = stringFormatN0(dr["WarningSpecMin"]),
                                WarningSpecMax = stringFormatN0(dr["WarningSpecMax"]),
                                WarningSpecPecntYN = dr["WarningSpecPecntYN"].ToString(),

                                Comments = dr["Comments"].ToString(),



                            };

                            dgdOutwareSub.Items.Add(CondtionSub);

                        }
                    } // for문 끝
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

        #endregion Sub조회


        /// <summary>
        /// code_gbn = SaGathItem 
        /// 한번 다 뿌리기
        /// </summary>
        #region  CM_CODE  SaGathItem 조회 
        private void FillgridSaGathItem()
        {
            if (dgdOutwareSub.Items.Count > 0)
            {
                dgdOutwareSub.Items.Clear();
            }

            try
            {
                Dictionary<string, object> sqlParameter = new Dictionary<string, object>();

                DataSet ds = DataStore.Instance.ProcedureToDataSet_LogWrite("xp_Condition_sSaGathItem", sqlParameter, true, "R");

                if (ds != null && ds.Tables.Count > 0)
                {
                    DataTable dt = ds.Tables[0];
                    int i = 0;


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
                            var SaGathItem = new Win_com_ArticleConditioin_U_Sub_CodeView()
                            {

                                CodeID = dr["CodeID"].ToString(),
                                CodeName = dr["CodeName"].ToString(),
                                WarningSpecPecntYN = "N"

                            };

                            dgdOutwareSub.Items.Add(SaGathItem);


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
        #endregion CM_CODE  SaGathItem

        #region 조건명 중복체크
        private void ChoiceSaGathItem()
        {
            var CheckCondtion = dgdOutwareSub.SelectedItem as Win_com_ArticleConditioin_U_Sub_CodeView;

            if (CheckCondtion != null
                && CheckCondtion.CodeName != null)
            {

                if (CheckIsCondi(CheckCondtion.CodeName, true) == false)
                {
                    MessageBox.Show("해당 라벨은 이미 등록되어 있습니다.");
                    return;
                }

            }


        }

        // 중복으로 라벨 등록하는걸 막기 위한 체크 이벤트
        // → 선택된 그 라벨은 제외 하고 검색을 해야 됨
        // ExcptSelLot : true (지금 서브 그리드 선택된 행의 LotID 를 제외 하고)
        // ExcptSelLot : false (지금 서브 그리드 선택된 행의 LotID 를 포함 해서)
        private bool CheckIsCondi(string CodeName, bool ExcptSelLot)
        {
            bool flag = true;

            string SelLotID = "";

            // 지금 활성화된 라벨
            var ConditionSub = dgdOutwareSub.SelectedItem as Win_com_ArticleConditioin_U_Sub_CodeView;
            if (ConditionSub != null)
            {
                SelLotID = ConditionSub.CodeName;
            }

            for (int i = 0; i < dgdOutwareSub.Items.Count; i++)
            {
                var Sub = dgdOutwareSub.Items[i] as Win_com_ArticleConditioin_U_Sub_CodeView;
                if (Sub != null
                    && Sub.CodeName != null
                    && !Sub.CodeName.Trim().Equals(""))
                {
                    if (ExcptSelLot == true
                        && SelLotID.Equals("") == false
                        && Sub.CodeName.Equals(SelLotID))
                    {
                        continue;
                    }

                    if (Sub.CodeName.ToUpper().Trim().Equals(CodeName.ToUpper().Trim()))
                    {
                        flag = false;
                        break;
                    }
                }
            }

            return flag;
        }

        #endregion

        #region 저장
        private bool SaveData(string strID, string strFlag)
        {
            bool flag = false;
            List<Procedure> Prolist = new List<Procedure>();
            List<Dictionary<string, object>> ListParameter = new List<Dictionary<string, object>>();

            //최소생산량 없으면 다 거름
            SelectCondtion();

            try
            {
                if (CheckData())
                {
                    Dictionary<string, object> sqlParameter = null;

                    #region 추가

                    if (strFlag.Equals("I"))
                    {
                        for (int i = 0; i < ListOutwareSub.Count; i++)
                        {
                            var SelectItem = ListOutwareSub[i] as Win_com_ArticleConditioin_U_Sub_CodeView;
                            sqlParameter = new Dictionary<string, object>();
                            sqlParameter.Clear();
                            sqlParameter.Add("sArticleID", strID);
                            sqlParameter.Add("sSetItemCode", SelectItem.CodeID);
                            sqlParameter.Add("sSpec", SelectItem.Spec);
                            sqlParameter.Add("sSpecMin", SelectItem.SpecMin == null ? 0 : ConvertDouble(SelectItem.SpecMin));
                            sqlParameter.Add("sSpecMax", SelectItem.SpecMax == null ? 0 : ConvertDouble(SelectItem.SpecMax));

                            sqlParameter.Add("sSpecTextMin", SelectItem.SpecTextMin);
                            sqlParameter.Add("sSpecTextMax", SelectItem.SpecTextMax);
                            sqlParameter.Add("sWarningSpecMin", SelectItem.WarningSpecMin == null ? 0 : ConvertDouble(SelectItem.WarningSpecMin));
                            sqlParameter.Add("sWarningSpecMax", SelectItem.WarningSpecMax == null ? 0 : ConvertDouble(SelectItem.WarningSpecMax));
                            sqlParameter.Add("sWarningSpecPecntYN", SelectItem.WarningSpecPecntYN);

                            sqlParameter.Add("sComments", SelectItem.Comments);

                            sqlParameter.Add("CreateUserID", MainWindow.CurrentUser);


                            Procedure pro1 = new Procedure();
                            pro1.Name = "xp_Condition_iArticleID";
                            pro1.OutputUseYN = "N";
                            pro1.OutputName = "sArticleID";
                            pro1.OutputLength = "10";

                            Prolist.Add(pro1);
                            ListParameter.Add(sqlParameter);


                        }

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
                        if (Procedure.Instance.DeleteData(strID, "sArticleID", "xp_Condition_dConditionArticleAll"))
                        {
                            for (int i = 0; i < dgdOutwareSub.Items.Count; i++)
                            {
                                //var SelectItem = ListOutwareSub[i] as Win_com_ArticleConditioin_U_Sub_CodeView;
                                var SelectItem = dgdOutwareSub.Items[i] as Win_com_ArticleConditioin_U_Sub_CodeView;

                                sqlParameter = new Dictionary<string, object>();
                                sqlParameter.Clear();

                                sqlParameter.Add("sArticleID", strID);
                                sqlParameter.Add("sSetItemCode", SelectItem.CodeID);
                                sqlParameter.Add("sSpec", SelectItem.Spec);
                                sqlParameter.Add("sSpecMin", ConvertDouble(SelectItem.SpecMin));
                                sqlParameter.Add("sSpecMax", ConvertDouble(SelectItem.SpecMax));

                                sqlParameter.Add("sSpecTextMin", SelectItem.SpecTextMin);
                                sqlParameter.Add("sSpecTextMax", SelectItem.SpecTextMax);
                                sqlParameter.Add("sWarningSpecMin", ConvertDouble(SelectItem.WarningSpecMin));
                                sqlParameter.Add("sWarningSpecMax", ConvertDouble(SelectItem.WarningSpecMax));
                                sqlParameter.Add("sWarningSpecPecntYN", SelectItem.WarningSpecPecntYN);

                                sqlParameter.Add("sComments", SelectItem.Comments);

                                sqlParameter.Add("CreateUserID", MainWindow.CurrentUser);

                                Procedure pro1 = new Procedure();

                                pro1.Name = "xp_Condition_iArticleID";
                                pro1.OutputUseYN = "N";
                                pro1.OutputName = "sArticleID";
                                pro1.OutputLength = "10";

                                Prolist.Add(pro1);
                                ListParameter.Add(sqlParameter);
                            }

                            string[] Confirm = new string[2];
                            //Confirm = DataStore.Instance.ExecuteAllProcedureOutputNew_NewLog(Prolist, ListParameter, "U");
                            Confirm = DataStore.Instance.ExecuteAllProcedureOutputNew(Prolist, ListParameter);

                            if (Confirm[0] != "success")
                            {
                                MessageBox.Show("[저장실패]\r\n" + Confirm[1].ToString());
                                flag = false;

                            }
                            else
                            {
                                flag = true;
                            }
                        }
                    }

                    #endregion
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

            return flag;

        }

        #endregion 저장
        /// <summary>
        /// 데이터 있는 것만 카운트함
        /// </summary>
        //데이터 있는것만 카운트
        private void SelectCondtion()
        {
            Win_com_ArticleConditioin_U_Sub_CodeView AllCondtion = null;
            // 에러메시지용 변수
            int MsgCnt = 0;
            string Msg = "";


            // 딱 넘겨주는것만 dgdOutwareSub 에 추가해주기
            List<Win_com_ArticleConditioin_U_Sub_CodeView> lstTemp = new List<Win_com_ArticleConditioin_U_Sub_CodeView>();


            for (int i = 0; i < dgdOutwareSub.Items.Count; i++)
            {
                AllCondtion = null;
                AllCondtion = dgdOutwareSub.Items[i] as Win_com_ArticleConditioin_U_Sub_CodeView;
                //생산최소 데이터 있는 것만 추가
                //조건 추가하던가?
                if (AllCondtion.SpecMin != ""
                    && AllCondtion.SpecMin != null)
                {

                    // ListOutwareSub 에 이미 있는 코드라면, 메시지 띄우기.
                    if (ListOutwareSub.Count > 0)
                    {
                        bool good = true;
                        for (int k = 0; k < ListOutwareSub.Count; k++)
                        {
                            var Condition = ListOutwareSub[k] as Win_com_ArticleConditioin_U_Sub_CodeView;
                            if (Condition != null)
                            {
                                if (Condition.CodeID.Trim().Equals(AllCondtion.CodeID.Trim()))
                                {
                                    MsgCnt++;
                                    Msg += Condition.CodeName + "\r";

                                    good = false;
                                    break;
                                }
                            }
                        }
                        if (good == true)
                        {
                            lstTemp.Add(AllCondtion);
                            ListOutwareSub.Add(AllCondtion);
                        }
                    }
                    else
                    {
                        lstTemp.Add(AllCondtion);
                        ListOutwareSub.Add(AllCondtion);
                    }
                }
            }


        }

        #region 데이터 체크
       
        // 그룹박스 데이터 기입체크
        private bool CheckData()
        {
            bool flag = true;

            if (txtArticle.Text.Length <= 0 || txtArticle.Tag.ToString().Equals(""))
            {
                MessageBox.Show("품번이 입력되지 않았습니다.");
                flag = false;
                return flag;
            }

            
            return flag;
        }
        #endregion

        #region 삭제
        private bool DeleteData(string strArticleID)
        {
            bool Flag = false;

            try
            {
                Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                sqlParameter.Clear();
                sqlParameter.Add("sArticleID", strArticleID);

                string[] result = DataStore.Instance.ExecuteProcedure_NewLog("xp_Condition_dConditionArticleAll", sqlParameter, "D");

                if (result[0].Equals("success"))
                {

                    Flag = true;
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


        #endregion 삭제

        //서브 데이터 그리드 삭제컬럼 버튼 클릭
        private void dgdOutwareSub_btnDelete_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var OutwareSub = dgdOutwareSub.SelectedItem as Win_com_ArticleConditioin_U_Sub_CodeView;
                if (OutwareSub != null)
                {
                    dgdOutwareSub.Items.Remove(OutwareSub);
                }


            }
            catch (Exception ee)
            {
                MessageBox.Show("오류지점 - dgdOutwareSub_btnDelete_Click : " + ee.ToString());
            }
        }




        #region 추가버튼 누를때 mt_ArticleCondition 정보가지고오기
        private void SelectArticleID(string strArticleID)
        {
            try
            {
                Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                sqlParameter.Add("ArticleID", strArticleID);

                DataSet ds = DataStore.Instance.ProcedureToDataSet("xp_Order_sArticleData_mtr", sqlParameter, false);

                if (ds != null && ds.Tables.Count > 0)
                {
                    DataTable dt = ds.Tables[0];

                    if (dt.Rows.Count > 0)
                    {
                        DataRow dr = dt.Rows[0];

                        var getArticleInfo = new GetConditionArticle
                        {
                            //ArticleGrpID = dr["ArticleGrpID"].ToString(),
                            //UnitPrice = dr["UnitPrice"].ToString(),


                        };

                        //cboArticleGrp.SelectedValue = getArticleInfo.ArticleGrpID;


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

        //추가, 수정일 때 
        private void CanBtnControl()
        {
            btnAdd.IsEnabled = false;               //추가
            btnUpdate.IsEnabled = false;            //수정
            btnDelete.IsEnabled = false;            //삭제
            btnClose.IsEnabled = true;              //닫기
            btnSearch.IsEnabled = false;            //검색
            btnSave.Visibility = Visibility.Visible;             //저장
            btnCancel.Visibility = Visibility.Visible;             //취소
            btnExcel.IsEnabled = false;             //엑셀
            btnReMainAdd.IsEnabled = false;     //유지추가
            EventLabel.Visibility = Visibility.Visible; //자료입력중

            dgdOutware.IsHitTestVisible = false;        //데이터그리드 클릭 안되게
            SubGridTop.IsHitTestVisible = true;         //서브그리드 TOP 클릭되게
            ButtonDataGridSubRowAdd.IsEnabled = true; //서브그리드 추가 버튼
            ButtonDataGridSubRowDel.IsEnabled = true; //서브그리드 삭제 버튼



        }
        //저장, 취소일 때
        private void CantBtnControl()
        {
            btnAdd.IsEnabled = true;               //추가
            btnUpdate.IsEnabled = true;            //수정
            btnDelete.IsEnabled = true;            //삭제
            btnClose.IsEnabled = true;             //닫기
            btnSearch.IsEnabled = true;            //검색
            btnSave.Visibility = Visibility.Hidden;             //저장
            btnCancel.Visibility = Visibility.Hidden;             //취소
            btnExcel.IsEnabled = true;             //엑셀
            btnReMainAdd.IsEnabled = true;     //유지추가



            EventLabel.Visibility = Visibility.Hidden; //자료입력 완료

            dgdOutware.IsHitTestVisible = true;        //데이터그리드 클릭되게
            SubGridTop.IsHitTestVisible = false;         //서브그리드 TOP 클릭 안되게

            ButtonDataGridSubRowAdd.IsEnabled = false; //서브그리드 추가 버튼
            ButtonDataGridSubRowDel.IsEnabled = false; //서브그리드 삭제 버튼




        }

        


        // 천자리 콤마, 소수점 버리기
        private string stringFormatN0(object obj)
        {
            return string.Format("{0:N0}", obj);
        }

        //더블로 형식 변환
        private double ConvertDouble(string str)
        {
            double result = 0;
            double chkDouble = 0;

            try
            {
                if (!str.Trim().Equals(""))
                {
                    str = str.Trim().Replace(",", "");

                    if (double.TryParse(str, out chkDouble) == true)
                    {
                        result = double.Parse(str);
                    }
                }
                return result;
            }
            catch (Exception ee)
            {
                MessageBox.Show("오류지점 - ConvertDouble : " + ee.ToString());
                return result;
            }
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

                return result;
            }
            catch (Exception ee)
            {
                MessageBox.Show("오류지점 - DatePickerFormat : " + ee.ToString());
                return result;
            }
        }


        private void chkReq_Click(object sender, RoutedEventArgs e)
        {
            CheckBox chkSender = sender as CheckBox;
            var Outware = chkSender.DataContext as Win_com_ArticleConditioin_U_CodeView;

            if (Outware != null)
            {
                if (chkSender.IsChecked == true)
                {
                    Outware.Chk = true;

                    if (lstOutwarePrint.Contains(Outware) == false)
                    {
                        lstOutwarePrint.Add(Outware);
                    }
                }
                else
                {
                    Outware.Chk = false;

                    if (lstOutwarePrint.Contains(Outware) == true)
                    {
                        lstOutwarePrint.Remove(Outware);
                    }
                }

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


        private void ButtonDataGridSubRowAdd_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                SubRowAdd();
            }
            catch (Exception ee)
            {
                MessageBox.Show("오류지점 - ButtonDataGridSubRowAdd_Click : " + ee.ToString());
            }
        }

        private void ButtonDataGridSubRowDel_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                SubRowDel();

            }
            catch (Exception ee)
            {
                MessageBox.Show("오류지점 - ButtonDataGridSubRowDel_Click : " + ee.ToString());
            }
        }

        //서브 그리드 추가
        private void SubRowAdd()
        {
            try
            {
                int index = dgdOutwareSub.Items.Count;

                var WOOSSC = new Win_com_ArticleConditioin_U_Sub_CodeView()
                {

                    CodeID = "",
                    CodeName = "",
                    Spec = "",
                    SpecMin = "",
                    SpecMax = "",

                    SpecTextMin = "",
                    SpecTextMax = "",
                    WarningSpecMin = "",
                    WarningSpecMax = "",
                    WarningSpecPecntYN = "N",

                    Comments = "",

                };
                dgdOutwareSub.Items.Add(WOOSSC);




            }
            catch (Exception ee)
            {
                MessageBox.Show("오류지점 - ButtonDataGridSubRowDel_Click : " + ee.ToString());
            }
        }

        //서브 그리드 삭제
        private void SubRowDel()
        {
            try
            {
                if (dgdOutwareSub.Items.Count > 0)
                {
                    if (dgdOutwareSub.SelectedItem != null)
                    {
                        if (dgdOutwareSub.CurrentItem != null)
                        {
                            dgdOutwareSub.Items.Remove(dgdOutwareSub.CurrentItem as Win_com_ArticleConditioin_U_Sub_CodeView);
                        }
                        else
                        {
                            ListOutwareSub.Add(dgdOutwareSub.SelectedItem as Win_com_ArticleConditioin_U_Sub_CodeView);
                            dgdOutwareSub.Items.Remove((dgdOutwareSub.Items[dgdOutwareSub.SelectedIndex]) as Win_com_ArticleConditioin_U_Sub_CodeView);
                        }

                        dgdOutwareSub.Refresh();
                    }
                    else
                    {
                        MessageBox.Show("삭제할 데이터를 먼저 선택하세요.");
                    }
                }
            }
            catch (Exception ee)
            {
                MessageBox.Show("오류지점 - ButtonDataGridSubRowDel_Click : " + ee.ToString());
            }
        }

        #region 서브 데이터그리드 방향키 이동 및 셀 포커스
        private void DataGridSub_PreviewKeyDown(object sender, KeyEventArgs e)
        {
            try
            {
                if (e.Key == Key.Down || e.Key == Key.Up || e.Key == Key.Left || e.Key == Key.Right)
                {
                    DataGridSub_KeyDown(sender, e);
                }
            }
            catch (Exception ee)
            {
                MessageBox.Show("오류지점 - DataGridSub_PreviewKeyDown " + ee.ToString());
            }
        }

        private void DataGridSub_KeyDown(object sender, KeyEventArgs e)
        {
            try
            {
                var SubItem = dgdOutwareSub.CurrentItem as Win_com_ArticleConditioin_U_Sub_CodeView;
                int rowCount = dgdOutwareSub.Items.IndexOf(dgdOutwareSub.CurrentItem);
                int colCount = dgdOutwareSub.Columns.IndexOf(dgdOutwareSub.CurrentCell.Column);
                int StartColumnCount = 1; //DataGridSub.Columns.IndexOf(dgdtpeMCoperationRateScore);
                int EndColumnCount = 10; //DataGridSub.Columns.IndexOf(dgdtpeComments);

                if (e.Key == Key.Enter)
                {
                    e.Handled = true;
                    (sender as DataGridCell).IsEditing = false;

                    if (EndColumnCount == colCount && dgdOutwareSub.Items.Count - 1 > rowCount)
                    {
                        dgdOutwareSub.SelectedIndex = rowCount + 1;
                        dgdOutwareSub.CurrentCell = new DataGridCellInfo(dgdOutwareSub.Items[rowCount + 1], dgdOutwareSub.Columns[StartColumnCount]);
                    }
                    else if (EndColumnCount > colCount && dgdOutwareSub.Items.Count - 1 > rowCount)
                    {
                        dgdOutwareSub.CurrentCell = new DataGridCellInfo(dgdOutwareSub.Items[rowCount], dgdOutwareSub.Columns[colCount + 1]);
                    }
                    else if (EndColumnCount == colCount && dgdOutwareSub.Items.Count - 1 == rowCount)
                    {
                        btnSave.Focus();
                    }
                    else if (EndColumnCount > colCount && dgdOutwareSub.Items.Count - 1 == rowCount)
                    {
                        dgdOutwareSub.CurrentCell = new DataGridCellInfo(dgdOutwareSub.Items[rowCount], dgdOutwareSub.Columns[colCount + 1]);
                    }
                    else
                    {
                        MessageBox.Show("있으면 찾아보자...");
                    }
                }
                else if (e.Key == Key.Down)
                {
                    e.Handled = true;
                    (sender as DataGridCell).IsEditing = false;

                    if (dgdOutwareSub.Items.Count - 1 > rowCount)
                    {
                        dgdOutwareSub.SelectedIndex = rowCount + 1;
                        dgdOutwareSub.CurrentCell = new DataGridCellInfo(dgdOutwareSub.Items[rowCount + 1], dgdOutwareSub.Columns[colCount]);
                    }
                    else if (dgdOutwareSub.Items.Count - 1 == rowCount)
                    {
                        if (EndColumnCount > colCount)
                        {
                            dgdOutwareSub.SelectedIndex = 0;
                            dgdOutwareSub.CurrentCell = new DataGridCellInfo(dgdOutwareSub.Items[0], dgdOutwareSub.Columns[colCount + 1]);
                        }
                        else
                        {
                            btnSave.Focus();
                        }
                    }
                }
                else if (e.Key == Key.Up)
                {
                    e.Handled = true;
                    (sender as DataGridCell).IsEditing = false;

                    if (rowCount > 0)
                    {
                        dgdOutwareSub.SelectedIndex = rowCount - 1;
                        dgdOutwareSub.CurrentCell = new DataGridCellInfo(dgdOutwareSub.Items[rowCount - 1], dgdOutwareSub.Columns[colCount]);
                    }
                }
                else if (e.Key == Key.Left)
                {
                    e.Handled = true;
                    (sender as DataGridCell).IsEditing = false;

                    if (colCount > 0)
                    {
                        dgdOutwareSub.CurrentCell = new DataGridCellInfo(dgdOutwareSub.Items[rowCount], dgdOutwareSub.Columns[colCount - 1]);
                    }
                }
                else if (e.Key == Key.Right)
                {
                    e.Handled = true;
                    (sender as DataGridCell).IsEditing = false;

                    if (EndColumnCount > colCount)
                    {
                        dgdOutwareSub.CurrentCell = new DataGridCellInfo(dgdOutwareSub.Items[rowCount], dgdOutwareSub.Columns[colCount + 1]);
                    }
                    else if (EndColumnCount == colCount)
                    {
                        if (dgdOutwareSub.Items.Count - 1 > rowCount)
                        {
                            dgdOutwareSub.SelectedIndex = rowCount + 1;
                            dgdOutwareSub.CurrentCell = new DataGridCellInfo(dgdOutwareSub.Items[rowCount + 1], dgdOutwareSub.Columns[StartColumnCount]);
                        }
                        else
                        {
                            btnSave.Focus();
                        }
                    }
                }
            }
            catch (Exception ee)
            {
                MessageBox.Show("오류지점 - DataGridSub_KeyDown " + ee.ToString());
            }
        }

        private void DataGridSub_TextFocus(object sender, KeyEventArgs e)
        {
            try
            {
                Lib.Instance.DataGridINControlFocus(sender, e);
            }
            catch (Exception ee)
            {
                MessageBox.Show("오류지점 - DataGridSub_TextFocus " + ee.ToString());
            }
        }

        private void DataGridSub_GotFocus(object sender, RoutedEventArgs e)
        {
            try
            {
                if (EventLabel.Visibility == Visibility.Visible)
                {

                    DataGridCell cell = sender as DataGridCell;

                    cell.IsEditing = true;


                }

            }
            catch (Exception ee)
            {
                MessageBox.Show("오류지점 - DataGridSub_GotFocus " + ee.ToString());
            }
        }

        private void DataGridSub_MouseUp(object sender, MouseButtonEventArgs e)
        {
            try
            {
                Lib.Instance.DataGridINBothByMouseUP(sender, e);
            }
            catch (Exception ee)
            {
                MessageBox.Show("오류지점 - DataGridSub_MouseUp " + ee.ToString());
            }
        }
        #endregion

        //서브 데이터 그리드 수량 숫자만 입력
        private void DataGridTextBoxColorQty_PreviewTextInput(object sender, TextCompositionEventArgs e)
        {
            try
            {
                Lib.Instance.CheckIsNumeric((TextBox)sender, e);
            }
            catch (Exception ee)
            {
                MessageBox.Show("오류지점 - DataGridTextBoxColorQty_PreviewTextInput : " + ee.ToString());
            }
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


        //입력칸 출력
        private void txtArticle_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                e.Handled = true;
                MainWindow.pf.ReturnCodeGLS(txtArticle, 7071, "");
            }

            if (txtArticle.Tag != null)
            {
                getArticleInfo(txtArticle.Tag.ToString());
            }

            //품명있는거 찾기
            if (txtArticle.Tag != null
                && strFlag.Equals("I")
                )
            {
                for (int k = 0; k < dgdOutware.Items.Count; k++)
                {
                    var dgdMain = dgdOutware.Items[k] as Win_com_ArticleConditioin_U_CodeView;

                    if (dgdMain != null)
                    {
                        if (dgdMain.ArticleID != null && dgdMain.ArticleID.Trim().Equals(txtArticle.Tag))
                        {
                            txtArticle.Text = "";
                            txtArticle.Tag = null;
                            txtBuyerArticle.Text = "";
                            MessageBox.Show("이미 입력된 품명입니다.");

                            return;
                        }
                    }
                }
            }


        }

        private void btnPfArticle_Click(object sender, RoutedEventArgs e)
        {
            MainWindow.pf.ReturnCodeGLS(txtArticle, 7071, "");



            if (txtArticle.Tag != null)
            {
                getArticleInfo(txtArticle.Tag.ToString());
            }

            //품명있는거 찾기
            if (txtArticle.Tag != null
                && strFlag.Equals("I")
                )
            {
                for (int k = 0; k < dgdOutware.Items.Count; k++)
                {
                    var dgdMain = dgdOutware.Items[k] as Win_com_ArticleConditioin_U_CodeView;

                    if (dgdMain != null)
                    {
                        if (dgdMain.ArticleID != null && dgdMain.ArticleID.Trim().Equals(txtArticle.Tag))
                        {
                            MessageBox.Show("이미 입력된 품번입니다.");

                            txtArticle.Text = "";
                            txtArticle.Tag = null;
                            txtBuyerArticle.Text = "";

                            return;
                        }
                    }
                }
            }

        }

        #region ArticleID 로 Article 정보 가져오기

        private void getArticleInfo(string setArticleID)
        {
            try
            {
                Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                sqlParameter.Add("ArticleID", setArticleID);

                DataSet ds = DataStore.Instance.ProcedureToDataSet("xp_Order_sArticleData_move", sqlParameter, false);

                if (ds != null && ds.Tables.Count > 0)
                {
                    DataTable dt = ds.Tables[0];

                    if (dt.Rows.Count > 0)
                    {
                        DataRow dr = dt.Rows[0];

                        var getArticleInfo = new ArticleInfo
                        {

                            BuyerArticleNo = dr["BuyerArticleNo"].ToString(),
                            Article = dr["Article"].ToString(),


                        };

                        txtBuyerArticle.Text = getArticleInfo.Article;

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






        private void ComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            ComboboxSub = dgdOutwareSub.CurrentItem as Win_com_ArticleConditioin_U_Sub_CodeView;
            ComboBox cboNewYN = (ComboBox)sender;

            if (ComboboxSub == null)
            {
                ComboboxSub = dgdOutwareSub.Items[rowSubNum] as Win_com_ArticleConditioin_U_Sub_CodeView;
            }

            if (cboNewYN.SelectedValue != null && !cboNewYN.SelectedValue.ToString().Equals(""))
            {

                var theView = cboNewYN.SelectedItem as CodeView;
                if (theView != null)
                {
                    ComboboxSub.WarningSpecPecnt = theView.code_name;
                    ComboboxSub.WarningSpecPecntYN = theView.code_id;
                    //cboPriceClss.SelectedValue = 0;

                }
                sender = cboNewYN;
            }
        }


        //비율처리여부 콤보박스
        private void ComboBox_Loaded(object sender, RoutedEventArgs e)
        {
            ComboBox CboWarningSpecPecntYN = (ComboBox)sender;

            //비율처리여부 콤보박스
            List<string[]> listNewYN = new List<string[]>();
            string[] New01 = new string[] { "Y", "Y" };
            string[] New02 = new string[] { "N", "N" };
            listNewYN.Add(New01);
            listNewYN.Add(New02);

            ObservableCollection<CodeView> ovcNewYN = ComboBoxUtil.Instance.Direct_SetComboBox(listNewYN);
            CboWarningSpecPecntYN.ItemsSource = ovcNewYN;
            CboWarningSpecPecntYN.DisplayMemberPath = "code_name";
            CboWarningSpecPecntYN.SelectedValuePath = "code_id";
            CboWarningSpecPecntYN.SelectedIndex = 1;
        }



        //조건명 팝업창 띄우기
        private void txtCodeName_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {

                var OcReqSub = dgdOutwareSub.CurrentItem as Win_com_ArticleConditioin_U_Sub_CodeView;


                if (OcReqSub != null)
                {
                    TextBox tb = new TextBox();

                   
                    e.Handled = true;


                    MainWindow.pf.ReturnCode(tb, 7076, "");

                    //조건명 있는거 찾기
                    for (int k = 0; k < dgdOutwareSub.Items.Count; k++)
                    {
                        var MoveSub = dgdOutwareSub.Items[k] as Win_com_ArticleConditioin_U_Sub_CodeView;

                        if (MoveSub != null)
                        {

                            if (MoveSub.CodeName != null && MoveSub.CodeName.Trim().Equals(tb.Text))
                            {
                                MessageBox.Show("이미 입력된 조건명입니다.");
                                //MoveSub.CodeName = "";
                                return;
                            }
                        }
                    }

                    //ChoiceSaGathItem();

                    if (tb.Tag != null)
                    {

                        OcReqSub.CodeID = tb.Tag.ToString();
                        OcReqSub.CodeName = tb.Text;

                    }



                }
            }
        }


    }


    class Win_com_ArticleConditioin_U_CodeView : BaseView
    {

        public bool Chk { get; set; }
        public string Num { get; set; }
        public string ArticleID { get; set; }
        public string Article { get; set; }
        public string BuyerArticleNo { get; set; }
        public string ConditionDate { get; set; }


    }

    public class Win_com_ArticleConditioin_U_Sub_CodeView : BaseView
    {
        public override string ToString()
        {
            return (this.ReportAllProperties());
        }


        public int Num { get; set; }

        public bool Chk { get; set; }
        public string ArticleID { get; set; }
        public string CodeID { get; set; }

        public string SetItemCode { get; set; }
        public string CodeName { get; set; }
        public string Spec { get; set; }

        public string SpecMin { get; set; }
        public string SpecMax { get; set; }
        public string SpecTextMin { get; set; }
        public string SpecTextMax { get; set; }
        public string WarningSpecMin { get; set; }
        public string WarningSpecMax { get; set; }
        public string WarningSpecPecntYN { get; set; }
        public string Comments { get; set; }
        public bool UDFlag { get; set; }

        public string WarningSpecPecnt { get; set; }

    }

    class GetConditionArticle : BaseView
    {

        public int Num { get; set; }
        public string ArticleID { get; set; }
        public string CodeID { get; set; }
        public string CodeName { get; set; }
    }




}

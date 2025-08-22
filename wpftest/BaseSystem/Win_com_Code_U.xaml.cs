using System;
using System.Collections.Generic;
using System.Data;
using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Controls;
using WizMes_WellMade.PopUP;

namespace WizMes_WellMade
{
    /// <summary>
    /// Win_sys_CommonCode_U.xaml에 대한 상호 작용 논리
    /// </summary>
    public partial class Win_com_Code_U : UserControl
    {
        string stDate = string.Empty;
        string stTime = string.Empty;

        string strFlag = string.Empty;
        Win_sys_CommonCode_U_CodeView winCode = new Win_sys_CommonCode_U_CodeView();
        Win_sys_CommonCode_U_CodeView_Sub winCodeSub = new Win_sys_CommonCode_U_CodeView_Sub();
        int rowNum = 0;
        int rowNumSub = 0;
        Lib lib = new Lib();

        public Win_com_Code_U()
        {
            InitializeComponent();
            //lbCS.Content = "코드 자리수는 영숫자 기준 [" + lbCode_Size + "] 자리 입니다.";
        }

        //화면 로드시
        private void UserControl_Loaded(object sender, RoutedEventArgs e)
        {
            stDate = DateTime.Now.ToString("yyyyMMdd");
            stTime = DateTime.Now.ToString("HHmm");

            DataStore.Instance.InsertLogByFormS(this.GetType().Name, stDate, stTime, "S");

            lib.UiLoading(sender);
        }

        #region 버튼 클릭시 작업

        //하단 텍스트 사용가능
        private void CanSMGroupControl()
        {
            //lblMsg.Visibility = Visibility.Visible;
            gbxCodeInfo.IsEnabled = true;
            rbnUseY.IsChecked = true;
        }

        //하단 텍스트 사용불가
        private void CantSMGroupControl()
        {
            //lblMsg.Visibility = Visibility.Hidden;
            gbxCodeInfo.IsEnabled = false;
        }

        //조회 클릭시
        private void btnSearch_Click(object sender, RoutedEventArgs e)
        {
            //검색버튼 비활성화
            btnSearch.IsEnabled = false;
            //딜레이주면 표시남. 딜레이 안주면 표가 안남.
            lib.Delay(500);

            rowNum = 0;
            rowNumSub = 0;
            re_Search(rowNum, rowNumSub);

            //검색 다 되면 활성화
            btnSearch.IsEnabled = true;

        }

        //추가 클릭시
        private void btnAdd_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (dgdMcg.SelectedItem != null && dgdScg.SelectedItem != null)
                {
                    rowNum = dgdMcg.SelectedIndex;
                    rowNumSub = dgdScg.SelectedIndex;

                    lib.UiButtonEnableChange_SCControl(this);
                    CanSMGroupControl();
                    strFlag = "I";
                    this.DataContext = null;

                    var Common = dgdMcg.SelectedItem as Win_sys_CommonCode_U_CodeView;

                    if (Common != null)
                    {
                        txtCode.MaxLength = ConvertInt(Common.Code_Size);
                    }

                    dgdMcg.IsEnabled = false;
                    dgdScg.IsEnabled = false;

                    txtSEQ.Text = string.Empty;

                    txtCode.IsEnabled = true; //코드 변경불가 

                    txtCode.Focus();
                }
                else
                {
                    MessageBox.Show("검색을 먼저 해주세요.");
                }
            }
            catch (Exception ee)
            {
                MessageBox.Show("오류지점 - 추가버튼 클릭 : " + ee.ToString());
            }
        }

        //수정 클릭시
        private void btnUpdate_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (dgdMcg.SelectedItem != null && dgdScg.SelectedItem != null)
                {
                    rowNum = dgdMcg.SelectedIndex;
                    rowNumSub = dgdScg.SelectedIndex;


                    var Common = dgdMcg.SelectedItem as Win_sys_CommonCode_U_CodeView;

                    if (Common != null)
                    {
                        txtCode.MaxLength = ConvertInt(Common.Code_Size);
                    }


                    lib.UiButtonEnableChange_SCControl(this);
                    CanSMGroupControl();
                    strFlag = "U";

                    dgdMcg.IsEnabled = false;
                    dgdScg.IsEnabled = false;

                    txtCode.IsEnabled = false; //코드 변경불가 
                    txtCode.Focus();
                }
                else
                {
                    MessageBox.Show("수정할 자료를 선택하고 눌러주십시오.");
                }
            }
            catch (Exception ee)
            {
                MessageBox.Show("오류지점 - 수정버튼 클릭 : " + ee.ToString());
            }
        }

        //삭제 클릭시
        private void btnDelete_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (dgdScg.SelectedItem == null)
                {
                    MessageBox.Show("삭제할 데이터가 지정되지 않았습니다. 삭제데이터를 지정하고 눌러주세요");
                }
                else
                {
                    winCode = dgdMcg.SelectedItem as Win_sys_CommonCode_U_CodeView;
                    var winCodeSub = dgdScg.SelectedItem as Win_sys_CommonCode_U_CodeView_Sub;

                    if (MessageBox.Show("선택하신 항목을 삭제하시겠습니까?", "삭제 전 확인", MessageBoxButton.YesNo) == MessageBoxResult.Yes)
                    {
                        if (dgdScg.Items.Count > 0 && dgdScg.SelectedItem != null)
                        {
                            rowNum = dgdMcg.SelectedIndex;
                            rowNumSub = dgdScg.SelectedIndex;
                        }

                        if (DeleteData(winCode.Code_ID, winCodeSub.Code_ID))
                        {
                            if (rowNumSub > 0)
                            {
                                rowNumSub -= 1;
                            }
                            re_Search(rowNum, rowNumSub);
                        }
                    }
                }
            }
            catch (Exception ee)
            {
                MessageBox.Show("오류지점 - 삭제버튼 클릭 : " + ee.ToString());
            }
        }

        //닫기 클릭시
        private void btnClose_Click(object sender, RoutedEventArgs e)
        {
            DataStore.Instance.InsertLogByFormS(this.GetType().Name, stDate, stTime, "E");
            lib.ChildMenuClose(this.ToString());
        }

        //저장 클릭시
        private void btnSave_Click(object sender, RoutedEventArgs e)
        {
            if (strFlag.Equals("I"))
            {
                if (SaveData())
                {
                    rowNum = 0; //추가 성공 시에는 첫 자료부터 조회
                    rowNumSub = 0;
                    lib.UiButtonEnableChange_IUControl(this);
                    CantSMGroupControl();
                    dgdMcg.IsEnabled = true;
                    dgdScg.IsEnabled = true;
                    strFlag = string.Empty;
                    re_Search(rowNum, rowNumSub);
                }
            }
            else  //U
            {
                if (UpdateData())
                {
                    lib.UiButtonEnableChange_IUControl(this);
                    CantSMGroupControl();
                    dgdMcg.IsEnabled = true;
                    dgdScg.IsEnabled = true;
                    strFlag = string.Empty;
                    re_Search(rowNum, rowNumSub);
                }
            }
        }

        //취소 클릭시
        private void btnCancel_Click(object sender, RoutedEventArgs e)
        {
            lib.UiButtonEnableChange_IUControl(this);
            CantSMGroupControl();
            dgdMcg.IsEnabled = true;
            dgdScg.IsEnabled = true;
            strFlag = string.Empty;
            re_Search(rowNum, rowNumSub);
        }

        //엑셀 클릭시
        private void btnExcel_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                DataTable dt = null;
                string Name = string.Empty;

                string[] lst = new string[4];
                lst[0] = "Code 메인 정보";
                lst[1] = "Code 세부 정보";
                lst[2] = dgdMcg.Name;
                lst[3] = dgdScg.Name;

                ExportExcelxaml ExpExc = new ExportExcelxaml(lst);

                ExpExc.ShowDialog();

                if (ExpExc.DialogResult.HasValue)
                {
                    if (ExpExc.choice.Equals(dgdMcg.Name))
                    {
                        DataStore.Instance.InsertLogByForm(this.GetType().Name, "E");
                        //MessageBox.Show("대분류");
                        if (ExpExc.Check.Equals("Y"))
                            dt = lib.DataGridToDTinHidden(dgdMcg);
                        else
                            dt = lib.DataGirdToDataTable(dgdMcg);

                        Name = dgdMcg.Name;

                        if (lib.GenerateExcel(dt, Name))
                            lib.excel.Visible = true;
                        else
                            return;
                    }
                    else if (ExpExc.choice.Equals(dgdScg.Name))
                    {
                        //MessageBox.Show("소분류");
                        if (ExpExc.Check.Equals("Y"))
                            dt = lib.DataGridToDTinHidden(dgdScg);
                        else
                            dt = lib.DataGirdToDataTable(dgdScg);
                        Name = dgdScg.Name;

                        if (lib.GenerateExcel(dt, Name))
                            lib.excel.Visible = true;
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
            }
            catch (Exception ee)
            {
                MessageBox.Show("오류지점 - 엑셀 클릭 : " + ee.ToString());
            }
        }

        //검색 + selectIndex
        private void re_Search(int selectIndex, int selectSubIndex)
        {
            FillGrid();

            if (dgdMcg.Items.Count > 0)
            {
                dgdMcg.SelectedIndex = selectIndex;

                if (dgdScg.Items.Count > 0)
                {
                    dgdScg.SelectedIndex = selectSubIndex;
                }
                else
                {
                    txtCode.Text = "";
                    txtCode_Name.Text = "";
                    txtCode_Name_Eng.Text = "";
                    txtRelation.Text = "";
                    txtSEQ.Text = "";
                    txtContent.Text = "";
                    rbnUseN.IsChecked = true;
                }
            }
            else
            {
                txtCode.Text = "";
                txtCode_Name.Text = "";
                txtCode_Name_Eng.Text = "";
                txtRelation.Text = "";
                txtSEQ.Text = "";
                txtContent.Text = "";
                rbnUseN.IsChecked = true;
            }
        }

        #endregion

        #region 대분류, 소분류 그리드 조회

        //대분류 그리드
        private void FillGrid()
        {
            if (dgdMcg.Items.Count > 0)
            {
                dgdMcg.Items.Clear();
            }

            try
            {
                Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                sqlParameter.Add("CodeSrh", txtCodeSrh.Text);

                DataSet ds = DataStore.Instance.ProcedureToDataSet_LogWrite("xp_Code_sCmCode_MainCategory", sqlParameter, true, "R");

                if (ds != null && ds.Tables.Count > 0)
                {
                    DataTable dt = ds.Tables[0];

                    //dataGrid.Items.Clear();
                    if (dt.Rows.Count == 0)
                    {
                        MessageBox.Show("조회된 데이터가 없습니다.");
                    }
                    else
                    {
                        DataRowCollection drc = dt.Rows;
                        int i = 0;
                        foreach (DataRow item in drc)
                        {
                            var window_commonCode_DTO = new Win_sys_CommonCode_U_CodeView()
                            {
                                Code_ID = item["Code_ID"] as string,
                                Code_Name = item["Code_Name"] as string,
                                Code_Name_Eng = item["Code_Name_Eng"] as string,
                                Code_Size = item["Code_Size"].ToString()
                            };

                            //if (window_commonCode_DTO.Code_Name.Contains("검사"))
                            //{
                            //    window_commonCode_DTO.IsInspect = true;
                            //}
                            //else
                            //{
                            //    window_commonCode_DTO.IsInspect = false;
                            //}
                            i++;
                            dgdMcg.Items.Add(window_commonCode_DTO);
                        }
                        //tbkCount.Text = "▶ 검색 결과 : " + i +"건";
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

        //대분류에서 index 선택
        private void dgdMcg_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            winCode = dgdMcg.SelectedItem as Win_sys_CommonCode_U_CodeView;

            if (winCode != null)
            {
                tbkCodeMsg.Text = "코드 자리수는 영숫자 기준 [" + winCode.Code_Size.ToString() + "] 자리 입니다.";
                FillGridSub(winCode.Code_ID);

                if (dgdScg.Items.Count > 0)
                {
                    dgdScg.SelectedIndex = 0;
                }
            }

            this.DataContext = dgdScg.SelectedItem as Win_sys_CommonCode_U_CodeView_Sub;
        }

        //소분류 그리드
        private void FillGridSub(string strCodeID)
        {
            if (dgdScg.Items.Count > 0)
            {
                dgdScg.Items.Clear();
            }

            try
            {
                Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                sqlParameter.Add("Code_GBN", strCodeID);
                sqlParameter.Add("CheckTF", (chkNoUse.IsChecked == true ? "" : "T"));

                DataSet ds = DataStore.Instance.ProcedureToDataSet("xp_Code_sCmCode_SmallCategory", sqlParameter, false);

                if (ds != null && ds.Tables.Count > 0)
                {
                    DataTable dt = ds.Tables[0];

                    if (dt.Rows.Count == 0)
                    {
                        MessageBox.Show("조회된 데이터가 없습니다.");
                    }
                    else
                    {
                        DataRowCollection drc = dt.Rows;

                        foreach (DataRow item in drc)
                        {
                            var window_commonCode_DTO2 = new Win_sys_CommonCode_U_CodeView_Sub()
                            {
                                Code_ID = item["Code_ID"] as string,
                                Code_Name = item["Code_Name"] as string,
                                Comments = item["Comments"] as string,
                                SEQ = item["SEQ"].ToString(),
                                Use_YN = item["Use_YN"] as string,
                                Code_Name_Eng = item["Code_Name_Eng"] as string,
                                Relation = item["Relation"] as string,
                                Code_Size = item["Code_Size"].ToString(),
                                Parent_ID = item["Parent_ID"] as string,
                                Level = item["Level"].ToString()
                            };
                            dgdScg.Items.Add(window_commonCode_DTO2);
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

        //소분류 그리드 index 선택
        private void dgdScg_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            winCodeSub = dgdScg.SelectedItem as Win_sys_CommonCode_U_CodeView_Sub;
            this.DataContext = winCodeSub;
        }

        #endregion

        // 소분류 데이터 실삭제
        private bool DeleteData(string strCodeGBN, string strCodeID)
        {
            bool CheckDelete = true;

            try
            {
                if (!strCodeGBN.Equals(string.Empty) && !strCodeID.Equals(string.Empty))
                {
                    Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                    sqlParameter.Add("CodeGBN", strCodeGBN);
                    sqlParameter.Add("Code_ID", strCodeID);

                    string[] result = DataStore.Instance.ExecuteProcedure_NewLog("xp_Code_dCmCode", sqlParameter, "D");

                    if (!result[0].Equals("success"))
                    {
                        CheckDelete = false;
                        //MessageBox.Show("실패 ㅠㅠ");
                    }
                    else
                    {
                        //MessageBox.Show("성공 *^^*");
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

            return CheckDelete;
        }

        //텍스트 박스 필수입력사항 체크
        private bool CheckData()
        {
            bool txtCheck = true;

            //코드입력 체크
            if (txtCode.Text == "" && txtCode.Text.Equals(""))
            {
                MessageBox.Show("코드를 입력해주세요");
                txtCheck = false;
                return txtCheck;
            }
            else
            {
                if (CheckCode() == false)
                {
                    MessageBox.Show("코드가 중복되거나 길이가 잘못되었습니다. 다른 코드를 입력하세요");
                    txtCheck = false;
                    return txtCheck;
                }
            }

            //한글명칭 입력체크
            if (txtCode_Name.Text == "" && txtCode_Name.Text.Equals(""))
            {
                MessageBox.Show("한글명칭을 입력해주세요");
                txtCheck = false;
                return txtCheck;
            }

            //관리순서 입력체크
            if (txtSEQ.Text.Trim() == "" && txtSEQ.Text.Trim().Equals(""))
            {
                MessageBox.Show("관리순서를 입력해주세요");
                txtCheck = false;
                return txtCheck;
            }

            //사용여부 체크여부
            if (rbnUseY.IsChecked == false && rbnUseN.IsChecked == false)
            {
                MessageBox.Show("사용여부를 체크해주세요");
                txtCheck = false;
                return txtCheck;
            }

            return txtCheck;
        }

        //공통코드의 코드ID 가 중복인지 확인(아닐시에 삽입)
        private bool CheckCode()
        {
            bool CheckCodeID = true;

            //int selected_index1 = dgCommonCode.SelectedIndex + 1;
            //Window_commonCodeView1 dgCC = dgCommonCode.SelectedItem as Window_commonCodeView1;

            try
            {
                Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                sqlParameter.Add("Code_GBN", winCode.Code_ID);
                sqlParameter.Add("Code_ID", txtCode.Text);

                DataSet ds = DataStore.Instance.ProcedureToDataSet("xp_Code_CheckCmCode", sqlParameter, false);

                if (ds != null && ds.Tables.Count > 0)
                {
                    DataTable dt = ds.Tables[0];

                    // 검색했을때 데이터가 존재할 경우 중복 체크
                    if (dt.Rows.Count != 0)
                    {
                        DataRowCollection drc = dt.Rows;

                        foreach (DataRow item in drc)
                        {

                            if (item[1].ToString().Equals(txtCode.Text))
                            {
                                CheckCodeID = false;
                                return CheckCodeID;
                            }
                        }
                        drc.Clear();
                    }
                    else // 중복이 되지 않았을때
                    {

                    }
                    dt.Clear();
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

            //if (txtCode.Text.Length != int.Parse(winCode.Code_Size))
            //{
            //    CheckCodeID = false;
            //    return CheckCodeID;
            //}

            return CheckCodeID;
        }

        //소분류 추가
        private bool SaveData()
        {
            bool flag = false;
            string UseYN = string.Empty;
            List<Procedure> Prolist = new List<Procedure>();
            List<Dictionary<string, object>> ListParameter = new List<Dictionary<string, object>>();

            try
            {
                if (CheckData())
                {
                    if (rbnUseY.IsChecked == true) { UseYN = "Y"; }
                    else { UseYN = "N"; }

                    if (winCode != null)
                    {
                        Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                        sqlParameter.Add("Code_GBN", winCode.Code_ID);    // 데이터 그리드1의 Code_ID == 데이터 그리드2의 Code_GBN
                        sqlParameter.Add("Code_ID", txtCode.Text);   //입력된 코드를 추가
                        sqlParameter.Add("Code_Name", txtCode_Name.Text);  //입력된 한글명칭 추가
                        sqlParameter.Add("Code_Name_Eng", txtCode_Name_Eng.Text); //입력된 영어명칭 추가
                        sqlParameter.Add("Comments", txtContent.Text);     //입력된 비고 추가
                        sqlParameter.Add("SEQ", int.Parse(txtSEQ.Text));      //입력된 관리순서 추가
                        sqlParameter.Add("Use_YN", UseYN);
                        sqlParameter.Add("Code_Size", 0);
                        sqlParameter.Add("Parent_ID", winCode.Code_ID);
                        sqlParameter.Add("Level", 0);
                        sqlParameter.Add("Relation", txtRelation.Text);
                        sqlParameter.Add("CreateUserID", MainWindow.CurrentUser);

                        string[] result = DataStore.Instance.ExecuteProcedure_NewLog("xp_Code_iCmCode", sqlParameter, "C");

                        //이건 왜 있는지?
                        //Procedure pro1 = new Procedure();
                        //pro1.Name = "xp_Code_iCmCode";
                        //pro1.OutputUseYN = "N";
                        //pro1.OutputName = "BankID";
                        //pro1.OutputLength = "5";

                        //Prolist.Add(pro1);
                        //ListParameter.Add(sqlParameter);

                        //List<KeyValue> list_Result = new List<KeyValue>();
                        //list_Result = DataStore.Instance.ExecuteAllProcedureOutputGetCS(Prolist, ListParameter);

                        //테스트용 sql문
                        //string strConn = "data source=wizis.iptime.org,20140;Initial catalog=MES_GLS; Integrated Security=True; connection timeout=180";
                        //string sql = "[xp_Code_iCmCode] 'INSBS', '06', 'aa', 'aa', '1', 1, 'y', 0, 'aa', 0, '1', 'aaa'";
                        //SqlConnection conn = new SqlConnection(strConn);
                        //conn.Open();
                        //SqlCommand cmd = new SqlCommand(sql, conn);

                        //cmd.ExecuteNonQuery();
                        //conn.Close();
                        //MessageBox.Show("성공");
                        if (result[0] != "success")
                        {
                            flag = false;
                            MessageBox.Show("실패 ㅠㅠ");
                        }
                        else
                        {
                            flag = true;
                            //MessageBox.Show("성공 *^^*");
                        }
                    }
                }
                else
                {
                    flag = false;
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

        //소분류 수정
        private bool UpdateData()
        {
            bool flag = false;
            string Use_YN = string.Empty;

            if (rbnUseY.IsChecked == true)
                Use_YN = "Y";
            else
                Use_YN = "N";

            try
            {
                Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                sqlParameter.Add("Code_GBN", winCode.Code_ID);  //
                sqlParameter.Add("Code_ID", winCodeSub.Code_ID);  //
                sqlParameter.Add("Code_Name", txtCode_Name.Text);  //입력된 한글명칭 추가
                sqlParameter.Add("Code_Name_Eng", txtCode_Name_Eng.Text); //입력된 영어명칭 추가
                sqlParameter.Add("Comments", txtContent.Text);     //입력된 비고 추가
                sqlParameter.Add("Relation", txtRelation.Text);
                sqlParameter.Add("SEQ", int.Parse(txtSEQ.Text));      //입력된 관리순서 추가
                sqlParameter.Add("Use_YN", Use_YN);
                sqlParameter.Add("UserID", MainWindow.CurrentUser);

                string[] result = DataStore.Instance.ExecuteProcedure_NewLog("xp_Code_uCmCode", sqlParameter, "U");

                if (!result[0].Equals("success"))
                {
                    flag = false;
                    //MessageBox.Show("실패 ㅠㅠ");
                }
                else
                {
                    flag = true;
                    //MessageBox.Show("성공 *^^*");
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

        #region 텍스트박스 엔터 → 다음 텍스트 박스 이동

        private void txtCode_KeyDown(object sender, System.Windows.Input.KeyEventArgs e)
        {
            if (e.Key == System.Windows.Input.Key.Enter)
            {
                txtCode_Name.Focus();
            }
        }

        private void txtCode_Name_KeyDown(object sender, System.Windows.Input.KeyEventArgs e)
        {
            if (e.Key == System.Windows.Input.Key.Enter)
            {
                txtCode_Name_Eng.Focus();
            }
        }

        private void txtCode_Name_Eng_KeyDown(object sender, System.Windows.Input.KeyEventArgs e)
        {
            if (e.Key == System.Windows.Input.Key.Enter)
            {
                txtSEQ.Focus();
            }
        }

        private void txtSEQ_KeyDown(object sender, System.Windows.Input.KeyEventArgs e)
        {
            if (e.Key == System.Windows.Input.Key.Enter)
            {
                txtRelation.Focus();
            }
        }

        private void txtRelation_KeyDown(object sender, System.Windows.Input.KeyEventArgs e)
        {
            if (e.Key == System.Windows.Input.Key.Enter)
            {
                txtContent.Focus();
            }
        }

        private void txtContent_KeyDown(object sender, System.Windows.Input.KeyEventArgs e)
        {
            if (e.Key == System.Windows.Input.Key.Enter)
            {
                btnSave.Focus();
            }
        }

        #endregion // 텍스트박스 엔터 → 다음 텍스트 박스 이동

        private void txtCode_PreviewTextInput(object sender, System.Windows.Input.TextCompositionEventArgs e)
        {
            lib.CheckIsNumeric((TextBox)sender, e);
        }

        private void txtCode_PreviewKeyDown(object sender, System.Windows.Input.KeyEventArgs e)
        {
            var Main = dgdMcg.SelectedItem as Win_sys_CommonCode_U_CodeView;
            if (Main != null)
            {
                //int CodeSize = ConvertInt(Main.Code_Size);
                //if (txtCode.Text.Length == CodeSize)
                //{
                //    e.Handled = true;
                //}
            }
        }

        //관리순서에 숫자만 들어가게
        private void txtSEQ_PreviewTextInput(object sender, System.Windows.Input.TextCompositionEventArgs e)
        {
            if (IsNumeric(e.Text) == false)
            {
                e.Handled = !IsNumeric(e.Text);
                MessageBox.Show("관리순서는 숫자만 입력하세요.");
            }
        }

        //숫자 외에 다른 문자열 못들어오도록
        public bool IsNumeric(string source)
        {
            Regex regex = new Regex("[^0-9.-]+");
            return !regex.IsMatch(source);
        }


    }

    //대분류 그리드
    public class Win_sys_CommonCode_U_CodeView : BaseView
    {
        public override string ToString()
        {
            return (this.ReportAllProperties());
        }

        //CM_Code 의 테이블에서 가져온 컬럼 Name 

        public string Code_ID { get; set; }     //소분류에서 Code_GBN
        public string Code_Name { get; set; }
        public string Code_Name_Eng { get; set; }
        public string Code_Size { get; set; }
        public bool IsInspect { get; set; }
    }

    public class Win_sys_CommonCode_U_CodeView_Sub : BaseView
    {
        public override string ToString()
        {
            return (this.ReportAllProperties());
        }

        public string Code_ID { get; set; }
        public string Code_GBN { get; set; }
        public string Code_Name { get; set; }
        public string Code_Name_Eng { get; set; }
        public string SEQ { get; set; }
        public string Relation { get; set; }
        public string Comments { get; set; }
        public string Use_YN { get; set; }
        public string Parent_ID { get; set; }
        public string Level { get; set; }
        public string Code_Size { get; set; }
        public string CreateDate { get; set; }
        public string CreateUserID { get; set; }
        public string LastUpdateDate { get; set; }
        public string LastUpdateUserID { get; set; }
    }
}

using System;
using System.Collections.Generic;
using System.Data;
using System.Windows;
using System.Windows.Controls;
using WizMes_WellMade.PopUP;

namespace WizMes_WellMade
{
    /// <summary>
    /// Win_com_BuseoJikChaek_U.xaml에 대한 상호 작용 논리
    /// </summary>
    public partial class Win_com_BuseoJikChaek_U : UserControl
    {
        string stDate = string.Empty;
        string stTime = string.Empty;

        string strFlag = string.Empty;
        int rowNumB = 0;
        int rowNumJ = 0;
        string CD_DEPART = "Depart";
        string CD_RESABLY = "Resably";
        Lib lib = new Lib();
        Win_com_Buseo_U_CodeView winBuseo = new Win_com_Buseo_U_CodeView();
        Win_com_JikChaek_U_CodeView winJikchaek = new Win_com_JikChaek_U_CodeView();

        public Win_com_BuseoJikChaek_U()
        {
            InitializeComponent();
        }

        private void UserControl_Loaded(object sender, RoutedEventArgs e)
        {
            stDate = DateTime.Now.ToString("yyyyMMdd");
            stTime = DateTime.Now.ToString("HHmm");

            DataStore.Instance.InsertLogByFormS(this.GetType().Name, stDate, stTime, "S");

            //Lib.Instance.UiLoading(sender);

            // 활성화된 메뉴 이름을 가져와서, 앞 두글자가 부서면 부서 / 직책이면 직책 탭 활성화
            MenuViewModel menuView = MainWindow.MainMdiContainer.ActiveMdiChild.Tag as MenuViewModel;
            string MenuName = menuView.Menu.Substring(0, 2);

            // 직책 메뉴 선택시 직책 매뉴 탭 활성화
            if (!MenuName.Equals("부서"))
            {
                tabItemBuseo.IsSelected = false; // 부서 탭 비활성화
                tabItemJikChaek.IsSelected = true; // 직책 탭 활성화

                // 직책 폼으로 활성화
                tblSrh.Text = "직책 검색";
                tbkBuseoCount.Visibility = Visibility.Collapsed;
                tbANTikChaekCount.Visibility = Visibility.Visible;
            }
        }

        //수정,추가 저장 후
        private void CanBtnControl()
        {
            Lib.Instance.UiButtonEnableChange_IUControl(this);
            lblMsgJ.Visibility = Visibility.Hidden;
            lblMsg.Visibility = Visibility.Hidden;
            gbxBoseo.IsHitTestVisible = false;
            gbxJikChaek.IsHitTestVisible = false;

            // 코드 관련 설명 지우기
            infoDepartID.Text = "";
            infoResablyID.Text = "";

            // 코드 텍스트박스 색상 변경
            //txtResablyID.Background = (SolidColorBrush)new BrushConverter().ConvertFrom("#c2fdc3");
            //txtDepartID.Background = (SolidColorBrush)new BrushConverter().ConvertFrom("#c2fdc3");

        }

        //수정,추가 진행 중
        private void CantBtnControl()
        {
            Lib.Instance.UiButtonEnableChange_SCControl(this);
            lblMsgJ.Visibility = Visibility.Visible;
            lblMsg.Visibility = Visibility.Visible;
            gbxBoseo.IsHitTestVisible = true;
            gbxJikChaek.IsHitTestVisible = true;
        }

        // 검색조건
        private void lblSrh_MouseLeftButtonUp(object sender, System.Windows.Input.MouseButtonEventArgs e)
        {
            if (chkSrh.IsChecked == true)
            {
                chkSrh.IsChecked = false;
            }
            else
            {
                chkSrh.IsChecked = true;
            }
        }
        private void chkSrh_Checked(object sender, RoutedEventArgs e)
        {
            chkSrh.IsChecked = true;
            txtBuseoSrh.IsEnabled = true;
            txtJikChaekSrh.IsEnabled = true;
        }
        private void chkSrh_Unchecked(object sender, RoutedEventArgs e)
        {
            chkSrh.IsChecked = false;
            txtBuseoSrh.IsEnabled = false;
            txtJikChaekSrh.IsEnabled = false;
        }

        //추가
        private void btnAdd_Click(object sender, RoutedEventArgs e)
        {
            CantBtnControl();
            strFlag = "I";

            if (tabItemBuseo.IsSelected == true)
            {
                lblMsg.Visibility = Visibility.Visible;
                tbkMsg.Text = "자료 입력 중";
                txtDepartID.Text = "";
                txtDepart.Text = "";

                // 추가시
                //txtDepartID.Background = (SolidColorBrush)new BrushConverter().ConvertFrom("#fff2d2");
                //txtDepartID.IsReadOnly = false;

                // 부서코드 등록 설명 추가
                //infoDepartID.Text = "* 부서코드는 2자리 이하의 숫자를 입력해주세요.";

                if (dgdBuseo.SelectedItem != null)
                {
                    rowNumB = dgdBuseo.SelectedIndex;
                }

                dgdBuseo.IsHitTestVisible = false;

                txtDepart.Focus();
            }
            else if (tabItemJikChaek.IsSelected == true)
            {
                lblMsgJ.Visibility = Visibility.Visible;
                tbkMsgJ.Text = "자료 입력 중";
                txtResablyID.Text = "";
                txtResably.Text = "";

                // 추가시
                //txtResablyID.Background = (SolidColorBrush)new BrushConverter().ConvertFrom("#fff2d2");
                //txtResablyID.IsReadOnly = false;

                // 직책코드 등록 설명 추가
                //infoResablyID.Text = "* 직책코드는 5자리 이하의 숫자를 입력해주세요.";

                if (dgdJikChaek.SelectedItem != null)
                {
                    rowNumJ = dgdJikChaek.SelectedIndex;
                }

                dgdJikChaek.IsHitTestVisible = false;

                txtResably.Focus();
            }
        }

        //수정
        private void btnUpdate_Click(object sender, RoutedEventArgs e)
        {
            if (tabItemBuseo.IsSelected == true)
            {
                winBuseo = dgdBuseo.SelectedItem as Win_com_Buseo_U_CodeView;
            }
            else if (tabItemJikChaek.IsSelected == true)
            {
                winJikchaek = dgdJikChaek.SelectedItem as Win_com_JikChaek_U_CodeView;
            }

            if ((tabItemBuseo.IsSelected == true && winBuseo != null)
                || (tabItemJikChaek.IsSelected == true && winJikchaek != null))
            {
                if (tabItemBuseo.IsSelected == true)
                {
                    rowNumB = dgdBuseo.SelectedIndex;
                    dgdBuseo.IsHitTestVisible = false;
                    tbkMsg.Text = "자료 수정 중";
                    lblMsg.Visibility = Visibility.Visible;

                    // 수정시 부서ID 변경 불가능
                    //txtDepartID.Background = (SolidColorBrush)new BrushConverter().ConvertFrom("#c2fdc3");
                    //txtDepartID.IsReadOnly = true;

                    // 부서코드 수정 설명 추가
                    //infoDepartID.Text = "* 부서코드는 수정이 불가능합니다.";
                }
                else if (tabItemJikChaek.IsSelected == true)
                {
                    rowNumJ = dgdJikChaek.SelectedIndex;
                    dgdJikChaek.IsHitTestVisible = false;
                    tbkMsgJ.Text = "자료 수정 중";
                    lblMsgJ.Visibility = Visibility.Visible;

                    // 수정시 직책ID 변경 불가능
                    //txtResablyID.Background = (SolidColorBrush)new BrushConverter().ConvertFrom("#c2fdc3");
                    //txtResablyID.IsReadOnly = true;

                    // 직책코드 수정 설명 추가
                    //infoResablyID.Text = "* 직책코드는 수정이 불가능합니다.";
                }

                CantBtnControl();
                strFlag = "U";
            }
            else
            {
                MessageBox.Show("수정할 자료를 선택하고 눌러주십시오.");
            }
        }

        //삭제
        private void btnDelete_Click(object sender, RoutedEventArgs e)
        {
            if (tabItemBuseo.IsSelected == true)
            {
                winBuseo = dgdBuseo.SelectedItem as Win_com_Buseo_U_CodeView;
            }
            else if (tabItemJikChaek.IsSelected == true)
            {
                winJikchaek = dgdJikChaek.SelectedItem as Win_com_JikChaek_U_CodeView;
            }

            if ((tabItemBuseo.IsSelected == true && winBuseo == null)
                || (tabItemJikChaek.IsSelected == true && winJikchaek == null))
            {
                MessageBox.Show("삭제할 데이터가 지정되지 않았습니다. 삭제데이터를 지정하고 눌러주세요");
            }
            else
            {
                if (MessageBox.Show("선택하신 항목을 삭제하시겠습니까?", "삭제 전 확인", MessageBoxButton.YesNo) == MessageBoxResult.Yes)
                {
                    if (tabItemBuseo.IsSelected == true)
                    {
                        if (dgdBuseo.Items.Count > 0 && dgdBuseo.SelectedItem != null)
                        {
                            rowNumB = dgdBuseo.SelectedIndex;
                        }

                        if (DeleteData(winBuseo.DepartID, CD_DEPART))
                        {
                            rowNumB -= 1;
                            re_Search(rowNumB, rowNumJ);
                        }
                    }
                    else if (tabItemJikChaek.IsSelected == true)
                    {
                        if (dgdJikChaek.Items.Count > 0 && dgdJikChaek.SelectedItem != null)
                        {
                            rowNumJ = dgdJikChaek.SelectedIndex;
                        }

                        if (DeleteData(winJikchaek.ResablyID, CD_RESABLY))
                        {
                            rowNumJ -= 1;
                            re_Search(rowNumB, rowNumJ);
                        }
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

        //조회
        private void btnSearch_Click(object sender, RoutedEventArgs e)
        {
            //검색버튼 비활성화
            btnSearch.IsEnabled = false;
            //딜레이주면 표시남. 딜레이 안주면 표가 안남.
            lib.Delay(500);

            re_Search(rowNumB, rowNumJ);

            //검색 다 되면 활성화
            btnSearch.IsEnabled = true;
        }

        //저장
        private void btnSave_Click(object sender, RoutedEventArgs e)
        {
            if (tabItemBuseo.IsSelected == true)    //부서
            {
                winBuseo = dgdBuseo.SelectedItem as Win_com_Buseo_U_CodeView;

                if (strFlag.Equals("I"))    //Insert
                {
                    if (SaveData(strFlag))
                    {
                        CanBtnControl();
                        lblMsg.Visibility = Visibility.Hidden;
                        rowNumB = 0;
                        dgdBuseo.IsHitTestVisible = true;
                        re_Search(rowNumB, rowNumJ);
                    }
                }
                else    //Update
                {
                    if (SaveData(strFlag))
                    {
                        CanBtnControl();
                        lblMsg.Visibility = Visibility.Hidden;
                        dgdBuseo.IsHitTestVisible = true;
                        re_Search(rowNumB, rowNumJ);
                    }
                }
            }
            else if (tabItemJikChaek.IsSelected == true)    //직책
            {
                winJikchaek = dgdJikChaek.SelectedItem as Win_com_JikChaek_U_CodeView;

                if (strFlag.Equals("I"))    //Insert
                {

                    if (SaveData(strFlag))
                    {
                        CanBtnControl();
                        lblMsgJ.Visibility = Visibility.Hidden;
                        rowNumJ = 0;
                        dgdJikChaek.IsHitTestVisible = true;
                        re_Search(rowNumB, rowNumJ);
                    }
                }
                else    //Update
                {
                    if (SaveData(strFlag))
                    {
                        CanBtnControl();
                        lblMsgJ.Visibility = Visibility.Hidden;
                        dgdJikChaek.IsHitTestVisible = true;
                        re_Search(rowNumB, rowNumJ);
                    }
                }
            }
        }

        //취소
        private void btnCancel_Click(object sender, RoutedEventArgs e)
        {
            CanBtnControl();
            strFlag = string.Empty;

            if (tabItemBuseo.IsSelected == true) { dgdBuseo.IsHitTestVisible = true; }
            else if (tabItemJikChaek.IsSelected == true) { dgdJikChaek.IsHitTestVisible = true; }

            re_Search(rowNumB, rowNumJ);
        }

        //엑셀
        private void btnExcel_Click(object sender, RoutedEventArgs e)
        {
            if (tabItemBuseo.Focus())
            {
                DataTable dt = null;
                string Name = string.Empty;

                string[] dgdStr = new string[2];
                dgdStr[0] = "부서 정보";
                dgdStr[1] = dgdBuseo.Name;

                ExportExcelxaml ExpExc = new ExportExcelxaml(dgdStr);
                ExpExc.ShowDialog();

                if (ExpExc.DialogResult.HasValue)
                {
                    if (ExpExc.choice.Equals(dgdBuseo.Name))
                    {
                        DataStore.Instance.InsertLogByForm(this.GetType().Name, "E");
                        if (ExpExc.Check.Equals("Y"))
                            if (ExpExc.Check.Equals("Y"))
                                dt = Lib.Instance.DataGridToDTinHidden(dgdBuseo);
                            else
                                dt = Lib.Instance.DataGirdToDataTable(dgdBuseo);

                        Name = dgdBuseo.Name;
                        if (Lib.Instance.GenerateExcel(dt, Name))
                            Lib.Instance.excel.Visible = true;
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
            else if (tabItemJikChaek.Focus())
            {
                DataTable dt = null;
                string Name = string.Empty;

                string[] dgdStr = new string[2];
                dgdStr[0] = "직책 정보";
                dgdStr[1] = dgdJikChaek.Name;

                ExportExcelxaml ExpExc = new ExportExcelxaml(dgdStr);
                ExpExc.ShowDialog();

                if (ExpExc.DialogResult.HasValue)
                {
                    if (ExpExc.choice.Equals(dgdJikChaek.Name))
                    {
                        if (ExpExc.Check.Equals("Y"))
                            dt = Lib.Instance.DataGridToDTinHidden(dgdJikChaek);
                        else
                            dt = Lib.Instance.DataGirdToDataTable(dgdJikChaek);

                        Name = dgdJikChaek.Name;
                        if (Lib.Instance.GenerateExcel(dt, Name))
                            Lib.Instance.excel.Visible = true;
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
        }

        //재검색
        private void re_Search(int selectedIndexB, int selectedIndexJ)
        {
            FillGrid();

            if (dgdBuseo.Items.Count > 0)
            {
                dgdBuseo.SelectedIndex = selectedIndexB;
            }

            if (dgdJikChaek.Items.Count > 0)
            {
                dgdJikChaek.SelectedIndex = selectedIndexJ;
            }
        }

        //조회
        private void FillGrid()
        {

            if (dgdBuseo.Items.Count > 0)
            {
                dgdBuseo.Items.Clear();
            }
            if (dgdJikChaek.Items.Count > 0)
            {
                dgdJikChaek.Items.Clear();
            }

            try
            {
                Dictionary<string, object> sqlParameter = new Dictionary<string, object>();

                // 부서
                sqlParameter.Clear();
                sqlParameter.Add("sDepart", chkSrh.IsChecked == true && !txtBuseoSrh.Text.Trim().Equals("") ? txtBuseoSrh.Text : "");
                sqlParameter.Add("sUseClss", chkUseClss.IsChecked == true ? 1 : 0);

                DataSet ds = DataStore.Instance.ProcedureToDataSet_LogWrite("xp_Code_sDepart", sqlParameter, true, "R");
                if (ds != null && ds.Tables.Count > 0)
                {
                    DataTable dt = ds.Tables[0];

                    if (dt.Rows.Count > 0)
                    {
                        int i = 0;
                        DataRowCollection drc = dt.Rows;

                        foreach (DataRow dr in drc)
                        {
                            i++;
                            var BuseoInfo = new Win_com_Buseo_U_CodeView()
                            {
                                Num = i + "",
                                DepartID = dr["DepartID"].ToString(),
                                Depart = dr["Depart"].ToString(),
                                UseClss = dr["UseClss"].ToString(),
                                Seq = dr["Seq"].ToString(),
                            };

                            dgdBuseo.Items.Add(BuseoInfo);
                        }

                        //tbkIndexCount.Text = "▶ 검색결과 : " + i.ToString() + " 건";
                    }
                }

                //직책
                sqlParameter.Clear();
                ds.Clear();


                sqlParameter.Add("sResably", chkSrh.IsChecked == true && !txtJikChaekSrh.Text.Trim().Equals("") ? txtJikChaekSrh.Text : "");
                sqlParameter.Add("sUseClss", chkUseClss.IsChecked == true ? 1 : 0);

                ds = DataStore.Instance.ProcedureToDataSet("xp_Code_sResably", sqlParameter, false);
                if (ds != null && ds.Tables.Count > 0)
                {
                    DataTable dt = ds.Tables[0];

                    if (dt.Rows.Count > 0)
                    {
                        int i = 0;
                        DataRowCollection drc = dt.Rows;

                        foreach (DataRow dr in drc)
                        {
                            i++;
                            var JikChaekInfo = new Win_com_JikChaek_U_CodeView()
                            {
                                Num = i + "",
                                ResablyID = dr["ResablyID"].ToString(),
                                Resably = dr["Resably"].ToString(),
                                UseClss = dr["UseClss"].ToString(),
                                Seq = dr["Seq"].ToString(),
                            };

                            dgdJikChaek.Items.Add(JikChaekInfo);
                        }

                        // 2019.08.28 검색결과에 갯수 추가
                        tbANTikChaekCount.Text = "▶검색 결과 : " + i + "건";
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

        //탭 체인지
        private void TabControl_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (tabItemBuseo.IsSelected == true)
            {
                tblSrh.Text = "부서 검색";
                txtBuseoSrh.Visibility = Visibility.Visible;
                txtJikChaekSrh.Visibility = Visibility.Hidden;

                tbkBuseoCount.Visibility = Visibility.Visible;
                tbANTikChaekCount.Visibility = Visibility.Collapsed;
            }
            else if (tabItemJikChaek.IsSelected == true)
            {
                tblSrh.Text = "직책 검색";
                txtBuseoSrh.Visibility = Visibility.Hidden;
                txtJikChaekSrh.Visibility = Visibility.Visible;

                tbkBuseoCount.Visibility = Visibility.Collapsed;
                tbANTikChaekCount.Visibility = Visibility.Visible;
            }
        }

        private void dgdBuseo_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            winBuseo = dgdBuseo.SelectedItem as Win_com_Buseo_U_CodeView;

            if (winBuseo != null)
            {
                //winBuseoJikchaek.DepartID = winBuseo.DepartID;
                //winBuseoJikchaek.Depart = winBuseo.Depart;
                tabItemBuseo.DataContext = winBuseo;
            }
        }

        private void dgdJikChaek_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            winJikchaek = dgdJikChaek.SelectedItem as Win_com_JikChaek_U_CodeView;

            if (winJikchaek != null)
            {
                //winBuseoJikchaek.ResablyID = winJikChaek.ResablyID;
                //winBuseoJikchaek.Resably = winJikChaek.Resably;
                tabItemJikChaek.DataContext = winJikchaek;
            }
        }

        //실삭제
        private bool DeleteData(string strID, string strTableName)
        {
            bool flag = false;

            Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
            sqlParameter.Clear();
            sqlParameter.Add("sTableName", strTableName);
            sqlParameter.Add("sCodeID", strTableName + "ID");
            sqlParameter.Add("sID", strID);

            string[] result = DataStore.Instance.ExecuteProcedure_NewLog("xp_Code_dCode_NotUse", sqlParameter, "D");
            DataStore.Instance.CloseConnection();

            if (result[0].Equals("success"))
            {
                //MessageBox.Show("성공 *^^*");
                flag = true;
            }

            return flag;
        }

        //실서장(추가,수정 후)
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

                    // CheckData() 로 유효성 검사 후에 부서 탭, 직책 탭 여부에 따라 파라미터 값 지정
                    if (tabItemBuseo.IsSelected == true)
                    {
                        sqlParameter.Add("sID", txtDepartID.Text);
                        sqlParameter.Add("sData", txtDepart.Text);
                        sqlParameter.Add("sSeq", txtDepartSeq.Text);

                        sqlParameter.Add("UseClss", chkBuseoUseClss.IsChecked == true ? "*" : "");

                        if (strFlag.Equals("I"))
                        {
                            sqlParameter.Add("CreateUserID", MainWindow.CurrentUser);

                            Procedure pro1 = new Procedure();
                            pro1.Name = "xp_Code_iDepart";
                            pro1.OutputUseYN = "N";
                            pro1.OutputName = "sCodeID";
                            pro1.OutputLength = "30";

                            Prolist.Add(pro1);
                            ListParameter.Add(sqlParameter);

                        }
                        else if (strFlag.Equals("U"))
                        {
                            sqlParameter.Add("LastUpdateUserID", MainWindow.CurrentUser);

                            Procedure pro1 = new Procedure();
                            pro1.Name = "xp_Code_uDepart";
                            pro1.OutputUseYN = "N";
                            pro1.OutputName = "sCodeID";
                            pro1.OutputLength = "30";

                            Prolist.Add(pro1);
                            ListParameter.Add(sqlParameter);

                        }
                    }
                    else if (tabItemJikChaek.IsSelected == true)
                    {
                        sqlParameter.Add("sID", txtResablyID.Text);
                        sqlParameter.Add("sData", txtResably.Text);
                        sqlParameter.Add("sSeq", txtResablySeq.Text);

                        sqlParameter.Add("UseClss", chANTikChaekUseClss.IsChecked == true ? "*" : "");

                        if (strFlag.Equals("I"))
                        {
                            sqlParameter.Add("CreateUserID", MainWindow.CurrentUser);

                            Procedure pro1 = new Procedure();
                            pro1.Name = "xp_Code_iResably";
                            pro1.OutputUseYN = "N";
                            pro1.OutputName = "sCodeID";
                            pro1.OutputLength = "30";

                            Prolist.Add(pro1);
                            ListParameter.Add(sqlParameter);

                        }
                        else if (strFlag.Equals("U"))
                        {
                            sqlParameter.Add("LastUpdateUserID", MainWindow.CurrentUser);

                            Procedure pro1 = new Procedure();
                            pro1.Name = "xp_Code_uResably";
                            pro1.OutputUseYN = "N";
                            pro1.OutputName = "sCodeID";
                            pro1.OutputLength = "30";

                            Prolist.Add(pro1);
                            ListParameter.Add(sqlParameter);

                        }
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
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }

            return flag;
        }

        //데이터체크
        private bool CheckData()
        {
            bool flag = true;

            if (tabItemBuseo.IsSelected == true)
            {

                // → 코드 자동 생성으로 변경 : 이건 사용안함
                //// 부서코드 숫자만 입력되도록 유효성 검사
                //int chkNum = 0;
                //if (Int32.TryParse(txtDepartID.Text, out chkNum) == false)
                //{
                //    MessageBox.Show("부서코드는 숫자만 입력 가능합니다.");
                //    flag = false;
                //    return flag;
                //}

                //// 부서코드 체크 (2자리의 자릿수로 입력되도록)
                //if (txtDepartID.Text.Length > 2)
                //{
                //    MessageBox.Show("부서코드 자릿수를 초과하셨습니다.");
                //    flag = false;
                //    return flag;
                //}
                //else
                //{
                //    // 입력되지 않았을 시
                //    if (txtDepartID.Text.Length <= 0 || txtDepartID.Text.Equals(""))
                //    {
                //        MessageBox.Show("부서코드가 입력되지 않았습니다.");
                //        flag = false;
                //        return flag;
                //    }

                //    // 만약 입력한 숫자가 한자리 라면 앞에 0을 추가
                //    if (txtDepartID.Text.Length < 2)
                //    {
                //        txtDepartID.Text = txtDepartID.Text.Insert(0, "0");
                //    }
                //}

                //// 부서코드 중복체크
                //if (strFlag.Equals("I") && CheckCode(CD_DEPART, txtDepartID.Text) == false)
                //{
                //    MessageBox.Show("입력하신 부서코드가 이미 존재합니다.");
                //    flag = false;
                //    return flag;
                //}

                // 부서코드 숫자만 입력되도록 유효성 검사

                // 부서명 체크
                if (txtDepart.Text.Length <= 0 || txtDepart.Text.Equals(""))
                {
                    MessageBox.Show("부서명이 입력되지 않았습니다.");
                    flag = false;
                    return flag;
                }
            }
            else if (tabItemJikChaek.IsSelected == true)
            {

                // → 코드 자동 생성으로 변경 : 이건 사용안함
                //// 직책코드 숫자만 입력되도록 유효성 검사
                //int chkNum = 0;
                //if (Int32.TryParse(txtResablyID.Text, out chkNum) == false)
                //{
                //    MessageBox.Show("직책코드는 숫자만 입력 가능합니다.");
                //    flag = false;
                //    return flag;
                //}

                //// 직책코드 체크 (5자리의 자릿수로 입력되도록)
                //if (txtResablyID.Text.Length > 5)
                //{
                //    MessageBox.Show("직책코드 자릿수를 초과하셨습니다.");
                //    flag = false;
                //    return flag;
                //}
                //else
                //{
                //    // 입력되지 않았을 시
                //    if (txtResablyID.Text.Length <= 0 || txtResablyID.Text.Equals(""))
                //    {
                //        MessageBox.Show("직책코드가 입력되지 않았습니다.");
                //        flag = false;
                //        return flag;
                //    }

                //    // 만약 입력한 숫자가 5자리 미만이라면 앞에 0을 추가
                //    if (txtResablyID.Text.Length < 5)
                //    {
                //        for(int i = txtResablyID.Text.Length; i < 5; i++ )
                //        {
                //            txtResablyID.Text = txtResablyID.Text.Insert(0, "0");                          
                //        }
                //    }
                //}

                //// 직책코드 중복체크
                //if (strFlag.Equals("I") && CheckCode(CD_RESABLY, txtResablyID.Text) == false)
                //{
                //    MessageBox.Show("입력하신 직책코드가 이미 존재합니다.");
                //    flag = false;
                //    return flag;
                //}

                // 직책명 체크
                if (txtResably.Text.Length <= 0 || txtResably.Text.Equals(""))
                {
                    MessageBox.Show("직책명이 입력되지 않았습니다.");
                    flag = false;
                    return flag;
                }
            }

            return flag;
        }

        // 추가시 부서코드, 직책코드 중복 체크
        public bool CheckCode(string tableName, string code)
        {
            Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
            sqlParameter.Add("TableName", "mt_" + tableName);
            sqlParameter.Add("sCodeID", tableName + "ID");
            sqlParameter.Add("sID", code);

            DataSet ds = DataStore.Instance.ProcedureToDataSet("xp_Code_sCheckDepartID", sqlParameter, false);
            DataTable dt = ds.Tables[0];
            DataRow dr = dt.Rows[0];

            int count = Convert.ToInt32(dr["num"].ToString());

            // 코드 갯수가 0보다 크다면 false 반환
            if (count > 0) { return false; }

            return true;
        }

        private void chkUseClss_Checked(object sender, RoutedEventArgs e)
        {
            chkUseClss.IsChecked = true;
        }

        private void chkUseClss_UnChecked(object sender, RoutedEventArgs e)
        {
            chkUseClss.IsChecked = false;
        }

        // 부서 엔터 → 저장 버튼 포커스
        private void txtDepart_KeyDown(object sender, System.Windows.Input.KeyEventArgs e)
        {
            if (e.Key == System.Windows.Input.Key.Enter)
            {
                btnSave.Focus();
            }

        }
        // 직책 엔터 → 저장 버튼 포커스
        private void txtResably_KeyDown(object sender, System.Windows.Input.KeyEventArgs e)
        {
            if (e.Key == System.Windows.Input.Key.Enter)
            {
                btnSave.Focus();
            }
        }

        private void lblUseClss_MouseLeftButtonUp(object sender, System.Windows.Input.MouseButtonEventArgs e)
        {
            if (chkUseClss.IsChecked == true)
            {
                chkUseClss.IsChecked = false;
            }
            else
            {
                chkUseClss.IsChecked = true;
            }
        }
    }

    class Win_com_Buseo_U_CodeView : BaseView
    {
        public override string ToString()
        {
            return (this.ReportAllProperties());
        }

        public string Num { get; set; }
        public string DepartID { get; set; }
        public string Depart { get; set; }
        public string UseClss { get; set; }
        public string Seq { get; set; }
    }

    class Win_com_JikChaek_U_CodeView : BaseView
    {
        public override string ToString()
        {
            return (this.ReportAllProperties());
        }

        public string Num { get; set; }
        public string ResablyID { get; set; }
        public string Resably { get; set; }
        public string UseClss { get; set; }
        public string Seq { get; set; }

    }
}

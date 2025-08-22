using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Data;
using System.Threading;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Input;
using System.Windows.Threading;
using WizMes_WellMade.PopUP;

//*******************************************************************************
//프로그램명    Win_Qul_InspectCode_U.cs
//메뉴ID        
//설명          Win_com_InspectCode_U 메인소스입니다.
//작성일        2019.07.25
//개발자        허윤구
//*******************************************************************************
// 변경일자     변경자      요청자      요구사항ID          요청 및 작업내용
//*******************************************************************************
//
//
//*******************************************************************************


namespace WizMes_WellMade
{
    /// <summary>
    /// Win_Qul_InspectCode_U.xaml에 대한 상호 작용 논리
    /// </summary>
    public partial class Win_Qul_InspectCode_U : UserControl
    {
        string stDate = string.Empty;
        string stTime = string.Empty;

        string strFlag = string.Empty;

        int Wh_Ar_SelectedLastIndex = 0;        // 그리드 마지막 선택 줄 임시저장 그릇

        ObservableCollection<Win_Qul_InspectCode_U_Defect_CodeView> ovcDefect
            = new ObservableCollection<Win_Qul_InspectCode_U_Defect_CodeView>();

        ObservableCollection<Win_Qul_InspectCode_U_Basic_CodeView> ovcBasis
            = new ObservableCollection<Win_Qul_InspectCode_U_Basic_CodeView>();

        ObservableCollection<Win_Qul_InspectCode_U_Grade_CodeView> ovcGrade
            = new ObservableCollection<Win_Qul_InspectCode_U_Grade_CodeView>();

        // 불량별 공정
        ObservableCollection<Win_Qul_InspectCode_U_DefectProcess_All_CodeView> ovcDefectProcess
            = new ObservableCollection<Win_Qul_InspectCode_U_DefectProcess_All_CodeView>();
        ObservableCollection<Win_Qul_InspectCode_U_DefectProcess_Select_CodeView> ovcDefectProcess_Select
            = new ObservableCollection<Win_Qul_InspectCode_U_DefectProcess_Select_CodeView>();

        Lib lib = new Lib();
        public Win_Qul_InspectCode_U()
        {
            InitializeComponent();
        }

        string SEQCheck = string.Empty; //버튼seq 확인하는 변수


        private void UserControl_Loaded(object sender, RoutedEventArgs e)
        {
            stDate = DateTime.Now.ToString("yyyyMMdd");
            stTime = DateTime.Now.ToString("HHmm");

            DataStore.Instance.InsertLogByFormS(this.GetType().Name, stDate, stTime, "S");
            SetComboBox();
            FillGrid();
            FillProcessGrid();
        }

        #region(콤보박스 설정) SetComboBox
        /// <summary>
        /// 콤보박스 setting
        /// </summary>
        private void SetComboBox()
        {
            ObservableCollection<CodeView> ovcDefect = ComboBoxUtil.Instance.Gf_DB_CM_GetComCodeDataset(null, "DFGRP", "Y", "");
            this.cboDefect.ItemsSource = ovcDefect;
            this.cboDefect.DisplayMemberPath = "code_name";
            this.cboDefect.SelectedValuePath = "code_id";

            ovcDefect = ComboBoxUtil.Instance.Gf_DB_CM_GetComCodeDataset(null, "INSDFTGRP", "Y", "");
            this.cboSebuDefect.ItemsSource = ovcDefect;
            this.cboSebuDefect.DisplayMemberPath = "code_name";
            this.cboSebuDefect.SelectedValuePath = "code_id";
        }

        #endregion


        /// <summary>
        /// 각 Tab아이템 선택시 동작
        /// </summary>
        private void TabControl_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (tabOne.IsSelected == true)
            {
                tbkOne.Visibility = Visibility.Visible;
                tbkTwo.Visibility = Visibility.Collapsed;
                tbkThree.Visibility = Visibility.Collapsed;
            }
            else if (tabTwo.IsSelected == true)
            {
                tbkOne.Visibility = Visibility.Collapsed;
                tbkTwo.Visibility = Visibility.Visible;
                tbkThree.Visibility = Visibility.Collapsed;
            }
            else
            {
                tbkOne.Visibility = Visibility.Collapsed;
                tbkTwo.Visibility = Visibility.Collapsed;
                tbkThree.Visibility = Visibility.Visible;
            }
        }

        /// <summary>
        /// 필터링
        /// </summary>
        private bool UseFilter(object Items)
        {
            if (tabOne.IsSelected == true)
            {
                if (String.IsNullOrEmpty(txtCodeSrh.Text))
                    return true;
                else
                    return ((Items as Win_Qul_InspectCode_U_Defect_CodeView).KDefect.IndexOf
                        (txtCodeSrh.Text, StringComparison.OrdinalIgnoreCase) >= 0);
            }
            else if (tabTwo.IsSelected == true)
            {
                if (String.IsNullOrEmpty(txtCodeSrh.Text))
                    return true;
                else
                    return ((Items as Win_Qul_InspectCode_U_Basic_CodeView).Basis.IndexOf
                        (txtCodeSrh.Text, StringComparison.OrdinalIgnoreCase) >= 0);
            }
            else
            {
                if (String.IsNullOrEmpty(txtCodeSrh.Text))
                    return true;
                else
                    return ((Items as Win_Qul_InspectCode_U_Grade_CodeView).Grade.IndexOf
                        (txtCodeSrh.Text, StringComparison.OrdinalIgnoreCase) >= 0);
            }
        }

        #region(실검색) FillGrid
        /// <summary>
        /// 실검색
        /// </summary>
        private void FillGrid()
        {
            try
            {
                DataStore.Instance.InsertLogByForm(this.GetType().Name, "R");

                ovcDefect.Clear();
                ovcBasis.Clear();
                ovcGrade.Clear();

                DataTable dtOne = Procedure.Instance.GetDefect();
                DataTable dtTwo = Procedure.Instance.GetBasis();
                DataTable dtThree = Procedure.Instance.GetGrade();

                if (dtOne.Rows.Count > 0)
                {
                    DataRowCollection drc = dtOne.Rows;
                    int i = 0;
                    foreach (DataRow dr in drc)
                    {
                        var WinOne = new Win_Qul_InspectCode_U_Defect_CodeView()
                        {
                            Num = i + 1,
                            DefectID = dr["DefectID"].ToString(),
                            Display1 = dr["Display1"].ToString(),
                            Display2 = dr["Display2"].ToString(),
                            Display3 = dr["Display3"].ToString(),
                            KDefect = dr["KDefect"].ToString(),
                            Edefect = dr["Edefect"].ToString(),
                            TagName = dr["TagName"].ToString(),
                            DefectClss = dr["DefectClss"].ToString(),
                            DefectClssSub = dr["DefectClssSub"].ToString(),
                            ButtonSeq = dr["ButtonSeq"].ToString()
                        };

                        ovcDefect.Add(WinOne);
                        i++;
                    }
                    dgdOne.ItemsSource = null;
                    dgdOne.ItemsSource = ovcDefect;
                    dgdOne.Items.Refresh();

                    CollectionView view = (CollectionView)CollectionViewSource.GetDefaultView(ovcDefect);
                    view.Filter = UseFilter;

                    if (dgdOne.Items.Count > 0)
                    {
                        dgdOne.SelectedIndex = 0;
                        var One = dgdOne.Items[0] as Win_Qul_InspectCode_U_Defect_CodeView;
                        tabOne.DataContext = One;
                    }
                }

                //if (dtTwo.Rows.Count > 0)
                //{
                //    DataRowCollection drc = dtTwo.Rows;
                //    int i = 0;
                //    foreach (DataRow dr in drc)
                //    {
                //        var WinTwo = new Win_Qul_InspectCode_U_Basic_CodeView()
                //        {
                //            Num = i + 1,
                //            Basis = dr["Basis"].ToString(),
                //            BasisID = dr["BasisID"].ToString()
                //        };

                //        ovcBasis.Add(WinTwo);
                //        i++;
                //    }

                //    dgdTwo.ItemsSource = ovcBasis;
                //    CollectionView view = (CollectionView)CollectionViewSource.GetDefaultView(ovcBasis);
                //    view.Filter = UseFilter;

                //    if (dgdTwo.Items.Count > 0)
                //    {
                //        dgdTwo.SelectedIndex = 0;
                //        var Two = dgdTwo.Items[0] as Win_Qul_InspectCode_U_Basic_CodeView;
                //        tabTwo.DataContext = Two;
                //    }
                //}

                //if (dtThree.Rows.Count > 0)
                //{
                //    DataRowCollection drc = dtThree.Rows;
                //    int i = 0;
                //    foreach (DataRow dr in drc)
                //    {
                //        var WinThree = new Win_Qul_InspectCode_U_Grade_CodeView()
                //        {
                //            Num = i + 1,
                //            Grade = dr["Grade"].ToString(),
                //            GradeID = dr["GradeID"].ToString()
                //        };

                //        ovcGrade.Add(WinThree);
                //        i++;
                //    }

                //    dgdThree.ItemsSource = ovcGrade;
                //    CollectionView view = (CollectionView)CollectionViewSource.GetDefaultView(ovcGrade);
                //    view.Filter = UseFilter;

                //    if (dgdThree.Items.Count > 0)
                //    {
                //        dgdThree.SelectedIndex = 0;
                //        var Three = dgdThree.Items[0] as Win_Qul_InspectCode_U_Grade_CodeView;
                //        tabThree.DataContext = Three;
                //    }
                //}
            }
            catch (Exception ex)
            {
                MessageBox.Show("오류 발생, 오류 내용 : " + ex.ToString());
            }
        }

        #endregion

        // 공정 그리드 채우기 >> 전체 그리드.
        private void FillProcessGrid()
        {
            ovcDefectProcess.Clear();
            ovcDefectProcess_Select.Clear();
            try
            {
                string sql = "SELECT * FROM MT_PROCESS where SUBSTRING(ProcessID,3,2) != '00' and UseClss != '*'";
                DataSet ds = DataStore.Instance.QueryToDataSet(sql);

                if (ds != null && ds.Tables.Count > 0)
                {
                    DataTable dt = ds.Tables[0];
                    if (dt.Rows.Count == 0)
                    {
                    }
                    else
                    {
                        int AllProcessNum = 1;
                        DataRowCollection drc = dt.Rows;
                        foreach (DataRow item in drc)
                        {
                            Win_Qul_InspectCode_U_DefectProcess_All_CodeView WQICDP = new Win_Qul_InspectCode_U_DefectProcess_All_CodeView()
                            {
                                AllProcessNum = AllProcessNum,
                                chkFlag = false,
                                Process = item["Process"].ToString(),
                                ProcessID = item["ProcessID"].ToString()
                            };
                            dgdAllProcess.Items.Add(WQICDP);
                            ovcDefectProcess.Add(WQICDP);
                            AllProcessNum++;
                        }
                        tbkAllCount.Text = "선택가능품목 : " + dgdAllProcess.Items.Count.ToString() + "개";
                        tbkSelectCount.Text = "선택품목 : " + dgdSelectProcess.Items.Count.ToString() + "개";
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
        /// 상단 텍스트 박스 텍스트 바뀔시
        /// </summary>
        private void TxtCodeSrh_TextChanged(object sender, TextChangedEventArgs e)
        {
            if (tabOne.IsSelected == true)
            {
                CollectionViewSource.GetDefaultView(dgdOne.ItemsSource).Refresh();
            }
            else if (tabTwo.IsSelected == true)
            {
                CollectionViewSource.GetDefaultView(dgdTwo.ItemsSource).Refresh();
            }
            else
            {
                CollectionViewSource.GetDefaultView(dgdThree.ItemsSource).Refresh();
            }
        }

        /// <summary>
        /// 추가,수정 시 동작 모음
        /// </summary>
        private void ControlVisibleAndEnable_AU()
        {
            Lib.Instance.UiButtonEnableChange_SCControl(this);
            tabOne.IsHitTestVisible = false;
            tabTwo.IsHitTestVisible = false;
            tabThree.IsHitTestVisible = false;
            if (tabOne.IsSelected == true)
            {
                tabOne.IsHitTestVisible = true;
                lblMsg1.Visibility = Visibility.Visible;

                if (strFlag.Equals("I"))
                    tbkMsg1.Text = "자료 입력 중";
                else
                    tbkMsg1.Text = "자료 수정 중";

                dgdOne.IsHitTestVisible = false;
                gbxOne.IsHitTestVisible = true;
                btnAddSelectItem.IsHitTestVisible = true;
                btnDelSelectItem.IsHitTestVisible = true;
                btnAllSelect_All.IsHitTestVisible = true;
                btnAllSelect_Select.IsHitTestVisible = true;
            }
            else if (tabTwo.IsSelected == true)
            {
                tabTwo.IsHitTestVisible = true;
                lblMsg2.Visibility = Visibility.Visible;

                if (strFlag.Equals("I"))
                    tbkMsg2.Text = "자료 입력 중";
                else
                    tbkMsg2.Text = "자료 수정 중";

                dgdTwo.IsHitTestVisible = false;
                gbxTwo.IsHitTestVisible = true;
            }
            else
            {
                tabThree.IsHitTestVisible = true;
                lblMsg3.Visibility = Visibility.Visible;

                if (strFlag.Equals("I"))
                    tbkMsg3.Text = "자료 입력 중";
                else
                    tbkMsg3.Text = "자료 수정 중";

                dgdThree.IsHitTestVisible = false;
                gbxThree.IsHitTestVisible = true;
            }
        }

        /// <summary>
        /// 저장,취소 시 동작 모음
        /// </summary>
        private void ControlVisibleAndEnable_SC()
        {
            Lib.Instance.UiButtonEnableChange_IUControl(this);
            tabOne.IsHitTestVisible = true;
            tabTwo.IsHitTestVisible = true;
            tabThree.IsHitTestVisible = true;
            if (tabOne.IsSelected == true)
            {
                lblMsg1.Visibility = Visibility.Hidden;
                dgdOne.IsHitTestVisible = true;
                gbxOne.IsHitTestVisible = false;
                btnAddSelectItem.IsHitTestVisible = false;
                btnDelSelectItem.IsHitTestVisible = false;
                btnAllSelect_All.IsHitTestVisible = false;
                btnAllSelect_Select.IsHitTestVisible = false;

                // 서브 그리드 두개 일단 클리어.
                dgdAllProcess.Items.Clear();
                dgdSelectProcess.Items.Clear();

            }
            else if (tabTwo.IsSelected == true)
            {
                lblMsg2.Visibility = Visibility.Hidden;
                dgdTwo.IsHitTestVisible = true;
                gbxTwo.IsHitTestVisible = false;
            }
            else
            {
                lblMsg3.Visibility = Visibility.Hidden;
                dgdThree.IsHitTestVisible = true;
                gbxThree.IsHitTestVisible = false;
            }
        }

        /// <summary>
        /// 추가 클릭
        /// </summary>
        private void btnAdd_Click(object sender, RoutedEventArgs e)
        {
            strFlag = "I";
            ControlVisibleAndEnable_AU();
            if (tabOne.IsSelected == true)
            {
                if (dgdOne.Items.Count > 0)
                {
                    Wh_Ar_SelectedLastIndex = dgdOne.SelectedIndex;
                }
                else
                {
                    Wh_Ar_SelectedLastIndex = 0;
                }

                tabOne.DataContext = null;

                dgdSelectProcess.Items.Clear();
                ovcDefectProcess_Select.Clear();
                tbkSelectCount.Text = "선택품목 : " + dgdSelectProcess.Items.Count.ToString() + "개";

                cboDefect.SelectedIndex = 0;
                cboSebuDefect.SelectedIndex = 0;
                //추가 시작 포커스. (단말기 Display1)
                txtDisplay1.Focus();
            }
            else if (tabTwo.IsSelected == true)
            {
                //SelectedIndex = dgdTwo.SelectedIndex;
                tabTwo.DataContext = null;
            }
            else
            {
                //SelectedIndex = dgdThree.SelectedIndex;
                tabThree.DataContext = null;
            }
        }

        /// <summary>
        /// 수정 클릭
        /// </summary>
        private void btnUpdate_Click(object sender, RoutedEventArgs e)
        {
            if (tabOne.IsSelected == true)
            {
                var One = dgdOne.SelectedItem as Win_Qul_InspectCode_U_Defect_CodeView;
                if (One == null)
                {
                    MessageBox.Show("수정할 자료를 선택하고 눌러주십시오.");
                    return;
                }
                Wh_Ar_SelectedLastIndex = dgdOne.SelectedIndex;

                Win_Qul_InspectCode_U_DefectProcess_All_CodeView AllProcess = null;
                foreach (Win_Qul_InspectCode_U_DefectProcess_Select_CodeView selectProcess in dgdSelectProcess.Items)
                {
                    string SelectProcessID = selectProcess.ProcessID;

                    for (int i = 0; i < dgdAllProcess.Items.Count; i++)
                    {
                        AllProcess = dgdAllProcess.Items[i] as Win_Qul_InspectCode_U_DefectProcess_All_CodeView;

                        if (AllProcess.ProcessID == SelectProcessID)
                        {
                            ovcDefectProcess.Remove(AllProcess);
                        }
                    }
                }
                dgdsubgrid_refill();

            }
            else if (tabTwo.IsSelected == true)
            {
                var One = dgdTwo.SelectedItem as Win_Qul_InspectCode_U_Basic_CodeView;
                if (One == null)
                {
                    MessageBox.Show("수정할 자료를 선택하고 눌러주십시오.");
                    return;
                }
                //SelectedIndex = dgdTwo.SelectedIndex;
            }
            else
            {
                var One = dgdThree.SelectedItem as Win_Qul_InspectCode_U_Grade_CodeView;
                if (One == null)
                {
                    MessageBox.Show("수정할 자료를 선택하고 눌러주십시오.");
                    return;
                }
                //SelectedIndex = dgdThree.SelectedIndex;
            }

            strFlag = "U";
            ControlVisibleAndEnable_AU();
        }

        /// <summary>
        /// 삭제 클릭
        /// </summary>
        private void btnDelete_Click(object sender, RoutedEventArgs e)
        {
            DataStore.Instance.InsertLogByForm(this.GetType().Name, "D");
            string strID = string.Empty;
            if (tabOne.IsSelected == true)
            {
                var Item = dgdOne.SelectedItem as Win_Qul_InspectCode_U_Defect_CodeView;
                if (Item != null)
                    strID = Item.DefectID;
            }
            else if (tabTwo.IsSelected == true)
            {
                var Item = dgdTwo.SelectedItem as Win_Qul_InspectCode_U_Basic_CodeView;
                if (Item != null)
                    strID = Item.BasisID;
            }
            else
            {
                var Item = dgdThree.SelectedItem as Win_Qul_InspectCode_U_Grade_CodeView;
                if (Item != null)
                    strID = Item.GradeID;
            }

            if (MessageBox.Show("선택하신 항목을 삭제하시겠습니까?", "삭제 전 확인", MessageBoxButton.YesNo) == MessageBoxResult.Yes)
            {

                if (dgdOne.Items.Count > 0 && dgdOne.SelectedItem != null)
                {
                    Wh_Ar_SelectedLastIndex = dgdOne.SelectedIndex;
                }

                if (DeleteData(strID))
                {
                    Wh_Ar_SelectedLastIndex -= 1;
                    btnSearch_Click(null, null);
                    if (dgdOne.Items.Count > 0)
                    {
                        dgdOne.SelectedIndex = Wh_Ar_SelectedLastIndex;
                        dgdOne.Focus();
                    }
                }
            }
        }

        /// <summary>
        /// 닫기 클릭
        /// </summary>
        private void btnClose_Click(object sender, RoutedEventArgs e)
        {
            DataStore.Instance.InsertLogByFormS(this.GetType().Name, stDate, stTime, "E");
            Lib.Instance.ChildMenuClose(this.ToString());
        }

        /// <summary>
        /// 검색 클릭
        /// </summary>
        private void btnSearch_Click(object sender, RoutedEventArgs e)
        {
            //검색버튼 비활성화
            btnSearch.IsEnabled = false;

            Dispatcher.BeginInvoke(new Action(() =>

            {
                Thread.Sleep(2000);

                //로직
                FillGrid();

                dgdOne.ItemsSource = null;
                dgdOne.Items.Refresh();
                dgdOne.ItemsSource = ovcDefect;
                dgdOne.Items.Refresh();

                if (tabOne.IsSelected == true)
                {
                    //dgdOne.SelectedIndex = SelectedIndex;
                }
                else if (tabTwo.IsSelected == true)
                {
                    //dgdTwo.SelectedIndex = SelectedIndex;
                }
                else
                {
                    //dgdThree.SelectedIndex = SelectedIndex;
                }

            }), System.Windows.Threading.DispatcherPriority.Background);



            Dispatcher.BeginInvoke(new Action(() =>

            {
                btnSearch.IsEnabled = true;

            }), System.Windows.Threading.DispatcherPriority.Background);


        }

        /// <summary>
        /// 저장 클릭
        /// </summary>
        private void btnSave_Click(object sender, RoutedEventArgs e)
        {
            string strID = string.Empty;

            if (tabOne.IsSelected == true)
            {
                strID = txtCode.Text;
            }
            else if (tabTwo.IsSelected == true)
            {
                strID = txtBasisID.Text;
            }
            else
            {
                strID = txtGradeID.Text;
            }

            try
            {
                if (CheckData())
                {
                    if (SaveData(strID))
                    {
                        ControlVisibleAndEnable_SC();
                        btnSearch_Click(null, null);
                        FillProcessGrid();

                        if (strFlag == "I")     //1. 추가 > 저장했다면,
                        {
                            if (dgdOne.Items.Count > 0)
                            {
                                dgdOne.SelectedIndex = dgdOne.Items.Count - 1;
                                dgdOne.Focus();
                            }
                        }
                        else        //2. 수정 > 저장했다면,
                        {
                            dgdOne.SelectedIndex = Wh_Ar_SelectedLastIndex;
                            dgdOne.Focus();
                        }

                        strFlag = string.Empty; // 추가했는지, 수정했는지 알려면 맨 마지막에 flag 값을 비어야 한다.
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("오류 발생, 오류 내용 : " + ex.ToString());
            }
        }


        //private bool CheckData()
        //{
        //    bool result = false;

        //    if (Lib.Instance.IsNullOrWhiteSpace(txtDisplay1.Text))
        //    {
        //        MessageBox.Show("단말기 Display1번은 필수입력입니다.");
        //        return result;
        //    }
        //    if (Lib.Instance.IsNullOrWhiteSpace(txtKDefect.Text))
        //    {
        //        MessageBox.Show("불량명 (한글)은 필수입력입니다.");
        //        return result;
        //    }

        //    //저장 전 ButtonSeq가 중복되지 않도록 체크. 

        //    if (strFlag.Equals("I") || strFlag.Equals("U"))
        //    {
        //        List<Procedure> Prolist = new List<Procedure>();
        //        List<Dictionary<string, object>> ListParameter = new List<Dictionary<string, object>>();

        //        Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
        //        sqlParameter.Clear();
        //        sqlParameter.Add("DefectID", txtCode.Text);
        //        sqlParameter.Add("NewButtonSeq", txtButtonSeq.Text.Trim());
        //        sqlParameter.Add("sMessage", "");

        //        Procedure pro1 = new Procedure();
        //        pro1.Name = "xp_Code_chkDefectButtonSeq";
        //        pro1.OutputUseYN = "Y";
        //        pro1.OutputName = "sMessage";
        //        pro1.OutputLength = "1000";

        //        Prolist.Add(pro1);
        //        ListParameter.Add(sqlParameter);

        //        //동운씨가 만든 아웃풋 값 찾는 방법
        //        List<KeyValue> list_Result = new List<KeyValue>();
        //        list_Result = DataStore.Instance.ExecuteAllProcedureOutputGetCS(Prolist, ListParameter);

        //        //Prolist.RemoveAt(0);
        //        //ListParameter.RemoveAt(0);

        //        string sGetID = string.Empty;

        //        if (list_Result[0].key.ToLower() == "success")
        //        {
        //            //list_Result.RemoveAt(0);
        //            for (int i = 0; i < list_Result.Count; i++)
        //            {
        //                KeyValue kv = list_Result[i];
        //                if (kv.key == "sMessage")
        //                {
        //                    sGetID = kv.value;

        //                    if (sGetID.Equals(""))
        //                    {
        //                        continue;
        //                    }

        //                    MessageBox.Show(" 알림 : " + sGetID.ToString());
        //                        result = false;
        //                    //strFlag = string.Empty;
        //                }
        //            }
        //        }
        //        Prolist.Clear();
        //        ListParameter.Clear();
        //        return result;
        //    }

        //    result = true;
        //    return result;
        //}



        //저장 전 ButtonSeq가 중복되지 않도록 체크. 
        private bool CheckData()
        {
            bool flag = true;

            if (Lib.Instance.IsNullOrWhiteSpace(txtDisplay1.Text))
            {
                MessageBox.Show("단말기 Display1번은 필수입력입니다.");
                flag = false;
                return flag;
            }
            if (Lib.Instance.IsNullOrWhiteSpace(txtKDefect.Text))
            {
                MessageBox.Show("불량명 (한글)은 필수입력입니다.");
                flag = false;
                return flag;
            }

            if (strFlag.Equals("I") || strFlag.Equals("U"))
            {
                if (strFlag.Equals("U") && SEQCheck == txtButtonSeq.Text)
                {
                    return flag;
                }

                //2020.05.26, 장가빈 추가. ButtonSeq에는 0이나 빈 값이 들어갈 수 없다.(빈 값일 경우 자동으로 값을 배정해주지 않기 때문에 체크해야 함) 
                if (txtButtonSeq.Text.Equals("0") || (txtButtonSeq.Text.Length <= 0 && txtButtonSeq.Text.Equals("")))
                {
                    MessageBox.Show("ButtonSeq는 0 혹은 빈 값일 수 없습니다. 중복되지 않은 순번을 입력해주세요.");
                    flag = false;
                    return flag;
                }
                else
                {
                    List<Procedure> Prolist = new List<Procedure>();
                    List<Dictionary<string, object>> ListParameter = new List<Dictionary<string, object>>();

                    Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                    sqlParameter.Clear();
                    sqlParameter.Add("DefectID", txtCode.Text);
                    sqlParameter.Add("NewButtonSeq", txtButtonSeq.Text.Trim());
                    sqlParameter.Add("sMessage", "");

                    Procedure pro1 = new Procedure();
                    pro1.Name = "xp_Code_chkDefectButtonSeq";
                    pro1.OutputUseYN = "Y";
                    pro1.OutputName = "sMessage";
                    pro1.OutputLength = "1000";

                    Prolist.Add(pro1);
                    ListParameter.Add(sqlParameter);

                    //동운씨가 만든 아웃풋 값 찾는 방법
                    List<KeyValue> list_Result = new List<KeyValue>();
                    list_Result = DataStore.Instance.ExecuteAllProcedureOutputGetCS(Prolist, ListParameter);

                    //Prolist.RemoveAt(0);
                    //ListParameter.RemoveAt(0);

                    string sGetID = string.Empty;

                    if (list_Result[0].key.ToLower() == "success")
                    {
                        //list_Result.RemoveAt(0);
                        for (int i = 0; i < list_Result.Count; i++)
                        {
                            KeyValue kv = list_Result[i];
                            if (kv.key == "sMessage")
                            {
                                sGetID = kv.value;

                                if (sGetID.Equals(""))
                                {
                                    continue;
                                }

                                MessageBox.Show(" 알림 : " + sGetID.ToString());
                                flag = false;
                                //strFlag = string.Empty;
                            }
                        }
                    }
                    Prolist.Clear();
                    ListParameter.Clear();
                    return flag;
                }
            }

            return flag;
        }




        /// <summary>
        /// 취소 클릭
        /// </summary>
        private void btnCancel_Click(object sender, RoutedEventArgs e)
        {
            ControlVisibleAndEnable_SC();
            btnSearch_Click(null, null);
            FillProcessGrid();

            if (strFlag == "I") // 1. 추가하다가 취소했다면,
            {
                if (dgdOne.Items.Count > 0)
                {
                    dgdOne.SelectedIndex = Wh_Ar_SelectedLastIndex;
                    dgdOne.Focus();
                }
            }
            else        //2. 수정하다가 취소했다면
            {
                dgdOne.SelectedIndex = Wh_Ar_SelectedLastIndex;
                dgdOne.Focus();
            }

            strFlag = string.Empty; // 추가했는지, 수정했는지 알려면 맨 마지막에 flag 값을 비어야 한다.
        }

        /// <summary>
        /// 엑셀 클릭
        /// </summary>
        private void btnExcel_Click(object sender, RoutedEventArgs e)
        {
            DataTable dt = null;
            string Name = string.Empty;
            Lib lib = new Lib();

            string[] dgdStr = new string[2];
            dgdStr[0] = "불량코드";
            dgdStr[1] = dgdOne.Name;


            ExportExcelxaml ExpExc = new ExportExcelxaml(dgdStr);
            ExpExc.ShowDialog();

            if (ExpExc.DialogResult.HasValue)
            {
                if (ExpExc.choice.Equals(dgdOne.Name))
                {
                    DataStore.Instance.InsertLogByForm(this.GetType().Name, "E");
                    if (ExpExc.Check.Equals("Y"))
                        dt = lib.DataGridToDTinHidden(dgdOne);
                    else
                        dt = lib.DataGirdToDataTable(dgdOne);

                    Name = dgdOne.Name;
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

        private void DgdOne_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            var One = dgdOne.SelectedItem as Win_Qul_InspectCode_U_Defect_CodeView;

            if (One != null)
            {
                tabOne.DataContext = One;
                SEQCheck = One.ButtonSeq;
            }

            // one이 잘 보인다면, 공정별 불량유형 (선택대상)을 뿌려줘야지.
            if (tbkOne.Visibility == Visibility.Visible)
            {
                dgdSelectProcess.Items.Clear();
                ovcDefectProcess_Select.Clear();

                var ViewReceiver = dgdOne.SelectedItem as Win_Qul_InspectCode_U_Defect_CodeView;
                if (ViewReceiver != null)
                {
                    string Select_DefectID = ViewReceiver.DefectID;

                    Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                    sqlParameter.Add("sDefectID", Select_DefectID);
                    DataSet ds = DataStore.Instance.ProcedureToDataSet("xp_Code_sDefectProcess", sqlParameter, false);

                    if (ds != null && ds.Tables.Count > 0)
                    {
                        DataTable dt = ds.Tables[0];

                        if (dt.Rows.Count > 0)
                        {
                            DataRowCollection drc = dt.Rows;
                            int i = 0;
                            foreach (DataRow dr in drc)
                            {
                                var Search_Select = new Win_Qul_InspectCode_U_DefectProcess_Select_CodeView()
                                {
                                    SelectProcessNum = i + 1,
                                    chkFlag = false,
                                    ProcessID = dr["ProcessID"].ToString(),
                                    Process = dr["Process"].ToString()
                                };
                                ovcDefectProcess_Select.Add(Search_Select);
                                dgdSelectProcess.Items.Add(Search_Select);
                                i++;
                            }
                        }
                    }

                    DataStore.Instance.CloseConnection();
                }
                tbkSelectCount.Text = "선택품목 : " + dgdSelectProcess.Items.Count.ToString() + "개";
            }
        }

        private void DgdTwo_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            var Two = dgdTwo.SelectedItem as Win_Qul_InspectCode_U_Basic_CodeView;

            if (Two != null)
            {
                tabTwo.DataContext = Two;
            }
        }

        private void DgdThree_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            var Three = dgdThree.SelectedItem as Win_Qul_InspectCode_U_Grade_CodeView;

            if (Three != null)
            {
                tabThree.DataContext = Three;
            }
        }

        /// <summary>
        /// 실삭제
        /// </summary>
        private bool DeleteData(string strID)
        {
            bool flag = false;

            try
            {
                if (tabOne.IsSelected == true)
                {
                    flag = Procedure.Instance.DeleteData(strID, "DefectID", "xp_Defect_dDefect");
                }
                else if (tabTwo.IsSelected == true)
                {
                    Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                    sqlParameter.Clear();
                    sqlParameter.Add("sTableName", "Basis");
                    sqlParameter.Add("sCodeID", "BasisID");
                    sqlParameter.Add("sID", strID);

                    string[] result = DataStore.Instance.ExecuteProcedure("xp_Code_dCode", sqlParameter, true);

                    if (result[0].Equals("success"))
                    {
                        //MessageBox.Show("성공 *^^*");
                        flag = true;
                    }
                }
                else
                {
                    flag = Procedure.Instance.DeleteData(strID, "GradeID", "xp_Grade_dGrade");
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
        /// 저장
        /// </summary>
        private bool SaveData(string strID)
        {
            bool flag = false;
            List<Procedure> Prolist = new List<Procedure>();
            List<Dictionary<string, object>> ListParameter = new List<Dictionary<string, object>>();

            try
            {
                if (tabOne.IsSelected == true)
                {
                    if (CheckData())
                    {
                        Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                        sqlParameter.Clear();
                        sqlParameter.Add("DefectID", strID);
                        sqlParameter.Add("ButtonSeq", txtButtonSeq.Text);
                        sqlParameter.Add("Display1", txtDisplay1.Text);
                        sqlParameter.Add("Display2", txtDisplay2.Text);
                        sqlParameter.Add("Display3", txtDisplay3.Text);
                        sqlParameter.Add("KDefect", txtKDefect.Text);
                        sqlParameter.Add("EDefect", txtEDefect.Text);
                        sqlParameter.Add("TagName", txtTagName.Text);
                        sqlParameter.Add("KindID", cboDefect.SelectedValue != null ? cboDefect.SelectedValue.ToString() : "");
                        sqlParameter.Add("KindIDSub", cboSebuDefect.SelectedValue != null ? cboSebuDefect.SelectedValue.ToString() : "");

                        if (strFlag.Equals("I"))
                        {
                            DataStore.Instance.InsertLogByForm(this.GetType().Name, "C");
                            sqlParameter.Add("BasisID", "0");
                            sqlParameter.Add("CreateUserID", MainWindow.CurrentUser);

                            Procedure pro1 = new Procedure();
                            pro1.Name = "xp_Defect_iDefect";
                            pro1.OutputUseYN = "Y";
                            pro1.OutputName = "DefectID";
                            pro1.OutputLength = "10";

                            Prolist.Add(pro1);
                            ListParameter.Add(sqlParameter);


                            for (int i = 0; i < dgdSelectProcess.Items.Count; i++)
                            {
                                var SelectProcess = dgdSelectProcess.Items[i] as Win_Qul_InspectCode_U_DefectProcess_Select_CodeView;
                                Dictionary<string, object> sqlParameter2 = new Dictionary<string, object>();
                                sqlParameter2.Clear();
                                sqlParameter2.Add("DefectID", strID);
                                sqlParameter2.Add("ProcessID", SelectProcess.ProcessID);
                                sqlParameter2.Add("CreateUserID", MainWindow.CurrentUser);

                                Procedure pro2 = new Procedure();
                                pro2.Name = "xp_Defect_iDefectProcess";
                                pro2.OutputUseYN = "N";
                                pro2.OutputName = "DefectID";
                                pro2.OutputLength = "10";

                                Prolist.Add(pro2);
                                ListParameter.Add(sqlParameter2);
                            }
                        }
                        else
                        {
                            DataStore.Instance.InsertLogByForm(this.GetType().Name, "U");
                            sqlParameter.Add("LastUpdateUserID", MainWindow.CurrentUser);

                            Procedure pro1 = new Procedure();
                            pro1.Name = "xp_Defect_uDefect";
                            pro1.OutputUseYN = "N";
                            pro1.OutputName = "DefectID";
                            pro1.OutputLength = "10";

                            Prolist.Add(pro1);
                            ListParameter.Add(sqlParameter);


                            for (int i = 0; i < dgdSelectProcess.Items.Count; i++)
                            {
                                var SelectProcess = dgdSelectProcess.Items[i] as Win_Qul_InspectCode_U_DefectProcess_Select_CodeView;
                                Dictionary<string, object> sqlParameter2 = new Dictionary<string, object>();
                                sqlParameter2.Clear();
                                sqlParameter2.Add("DefectID", strID);
                                sqlParameter2.Add("ProcessID", SelectProcess.ProcessID);
                                sqlParameter2.Add("CreateUserID", MainWindow.CurrentUser);

                                Procedure pro2 = new Procedure();
                                pro2.Name = "xp_Defect_iDefectProcess";
                                pro2.OutputUseYN = "N";
                                pro2.OutputName = "DefectID";
                                pro2.OutputLength = "10";

                                Prolist.Add(pro2);
                                ListParameter.Add(sqlParameter2);
                            }

                        }

                        List<KeyValue> list_Result = new List<KeyValue>();
                        list_Result = DataStore.Instance.ExecuteAllProcedureOutputGetCS(Prolist, ListParameter);
                        string sGetID = string.Empty;

                        if (list_Result[0].key.ToLower() == "success")
                        {
                            flag = true;
                        }
                        else
                        {
                            MessageBox.Show("[저장실패]\r\n" + list_Result[0].value.ToString());
                            flag = false;
                            //return false;
                        }
                    }
                    //else if (tabTwo.IsSelected == true)
                    //{
                    //    Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                    //    sqlParameter.Clear();
                    //    sqlParameter.Add("sTableName", "Basis");
                    //    sqlParameter.Add("sCodeID", "BasisID");
                    //    sqlParameter.Add("sData", txtBasis.Text);

                    //    if (strFlag.Equals("I"))
                    //    {
                    //        sqlParameter.Add("sID", Procedure.Instance.GetMaxValue("BasisID", "mt_Basis"));
                    //        sqlParameter.Add("CreateUserID", MainWindow.CurrentUser);
                    //        string[] result = DataStore.Instance.ExecuteProcedure("xp_Code_iCode", sqlParameter, false);
                    //        if (!result[0].Equals("success"))
                    //        {
                    //            MessageBox.Show("실패 원인 : " + result[1]);
                    //        }
                    //        else
                    //        {
                    //            flag = true;
                    //        }
                    //    }
                    //    else
                    //    {
                    //        sqlParameter.Add("sID", strID);
                    //        sqlParameter.Add("LastUpdateUserID", MainWindow.CurrentUser);
                    //        string[] result = DataStore.Instance.ExecuteProcedure("xp_Code_uCode", sqlParameter, false);
                    //        if (!result[0].Equals("success"))
                    //        {
                    //            MessageBox.Show("실패 원인 : " + result[1]);
                    //        }
                    //        else
                    //        {
                    //            flag = true;
                    //        }
                    //    }
                    //}
                    //else
                    //{
                    //    Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                    //    sqlParameter.Clear();
                    //    sqlParameter.Add("GradeID", strID);
                    //    sqlParameter.Add("Grade", txtGrade.Text);

                    //    if (strFlag.Equals("I"))
                    //    {
                    //        sqlParameter.Add("CreateUserID", MainWindow.CurrentUser);

                    //        Dictionary<string, int> outputParam = new Dictionary<string, int>();
                    //        outputParam.Add("GradeID", 10);
                    //        Dictionary<string, string> dicResult =
                    //            DataStore.Instance.ExecuteProcedureOutputNoTran("xp_Grade_iGrade", sqlParameter, outputParam, true);
                    //        string result = dicResult["GradeID"];
                    //        if ((result != string.Empty) || (result != "9999"))
                    //        {
                    //            flag = true;
                    //        }
                    //    }
                    //    else
                    //    {
                    //        sqlParameter.Add("LastUpdateUserID", MainWindow.CurrentUser);
                    //        string[] result = DataStore.Instance.ExecuteProcedure("xp_Grade_uGrade", sqlParameter, false);
                    //        if (!result[0].Equals("success"))
                    //        {
                    //            MessageBox.Show("실패 원인 : " + result[1]);
                    //        }
                    //        else
                    //        {
                    //            flag = true;
                    //        }
                    //    }
                    //}
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


        // [서브그리드 전체공정] 체크박스 클릭 이벤트.
        private void chkAllItem_Click(object sender, RoutedEventArgs e)
        {
            CheckBox chkSender = sender as CheckBox;
            if (lblMsg1.Visibility == Visibility.Visible)
            {
                var dgdAll = chkSender.DataContext as Win_Qul_InspectCode_U_DefectProcess_All_CodeView;
                if (chkSender.IsChecked == true)
                {
                    dgdAll.chkFlag = true;
                }
                else
                {
                    dgdAll.chkFlag = false;
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
                MessageBox.Show("체크박스를 사용하려면 먼저 추가나 수정을 누르고 진행해야 합니다.");
            }
        }

        // [서브그리드 선택공정] 체크박스 클릭 이벤트.
        private void chkSelectItem_Click(object sender, RoutedEventArgs e)
        {
            CheckBox chkSender = sender as CheckBox;
            if (lblMsg1.Visibility == Visibility.Visible)
            {
                var dgdSelect = chkSender.DataContext as Win_Qul_InspectCode_U_DefectProcess_Select_CodeView;

                if (chkSender.IsChecked == true)
                {
                    dgdSelect.chkFlag = true;
                }
                else
                {
                    dgdSelect.chkFlag = false;
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
                MessageBox.Show("체크박스를 사용하려면 먼저 추가나 수정을 누르고 진행해야 합니다.");
            }
        }




        // 전체공정 에서 >> SELECT 공정으로 이동버튼 클릭.
        private void btnAddSelectItem_Click(object sender, RoutedEventArgs e)
        {
            OVC_Remake_Select();
            dgdsubgrid_refill();
        }
        // Select공정 에서 >> 전체공정으로 이동버튼 클릭. (공정제외)
        private void btnDelSelectItem_Click(object sender, RoutedEventArgs e)
        {
            OVC_Remake_All();
            dgdsubgrid_refill();
        }

        private void OVC_Remake_Select()
        {
            Win_Qul_InspectCode_U_DefectProcess_All_CodeView AllProcess = null;
            Win_Qul_InspectCode_U_DefectProcess_Select_CodeView selectProcess = null;
            int j = 0;
            for (int i = 0; i < dgdAllProcess.Items.Count; i++)
            {
                AllProcess = dgdAllProcess.Items[i] as Win_Qul_InspectCode_U_DefectProcess_All_CodeView;

                if (AllProcess.chkFlag == true)
                {
                    selectProcess = new Win_Qul_InspectCode_U_DefectProcess_Select_CodeView()
                    {
                        SelectProcessNum = j + 1,
                        chkFlag = false,
                        ProcessID = AllProcess.ProcessID,
                        Process = AllProcess.Process
                    };
                    ovcDefectProcess_Select.Add(selectProcess);
                    ovcDefectProcess.Remove(AllProcess);
                    j++;
                }
            }
        }

        private void OVC_Remake_All()
        {
            Win_Qul_InspectCode_U_DefectProcess_All_CodeView AllProcess = null;
            Win_Qul_InspectCode_U_DefectProcess_Select_CodeView selectProcess = null;
            int j = 0;
            for (int i = 0; i < dgdSelectProcess.Items.Count; i++)
            {
                selectProcess = dgdSelectProcess.Items[i] as Win_Qul_InspectCode_U_DefectProcess_Select_CodeView;
                if (selectProcess.chkFlag == true)
                {
                    AllProcess = new Win_Qul_InspectCode_U_DefectProcess_All_CodeView()
                    {
                        AllProcessNum = j + 1,
                        chkFlag = false,
                        ProcessID = selectProcess.ProcessID,
                        Process = selectProcess.Process
                    };
                    ovcDefectProcess.Add(AllProcess);
                    ovcDefectProcess_Select.Remove(selectProcess);
                    j++;
                }
            }
        }

        private void dgdsubgrid_refill()
        {
            int j = 0;
            int t = 0;

            if (dgdAllProcess.Items.Count > 0)
            {
                dgdAllProcess.Items.Clear();
            }
            if (dgdSelectProcess.Items.Count > 0)
            {
                dgdSelectProcess.Items.Clear();
            }
            for (j = 0; ovcDefectProcess.Count > j; j++)
            {
                var selectionItem = ovcDefectProcess[j] as Win_Qul_InspectCode_U_DefectProcess_All_CodeView;
                selectionItem.chkFlag = false;
                selectionItem.AllProcessNum = (j + 1);
                dgdAllProcess.Items.Add(selectionItem);
            }
            for (t = 0; ovcDefectProcess_Select.Count > t; t++)
            {
                var selectionItem = ovcDefectProcess_Select[t] as Win_Qul_InspectCode_U_DefectProcess_Select_CodeView;
                selectionItem.chkFlag = false;
                selectionItem.SelectProcessNum = (t + 1);
                dgdSelectProcess.Items.Add(selectionItem);
            }
            tbkAllCount.Text = "선택가능품목 : " + dgdAllProcess.Items.Count.ToString() + "개";
            tbkSelectCount.Text = "선택품목 : " + dgdSelectProcess.Items.Count.ToString() + "개";
        }


        // 서브 그리드 내 전체선택 버튼 클릭 시.
        // 모두 선택만 되도록. 어차피 움직이는건 위에 버튼들이 할거야.
        private void btnAllSelect_All_Click(object sender, RoutedEventArgs e)
        {
            foreach (Win_Qul_InspectCode_U_DefectProcess_All_CodeView AllProcess in dgdAllProcess.Items)
            {
                AllProcess.chkFlag = true;
            }
        }

        private void btnAllSelect_Select_Click(object sender, RoutedEventArgs e)
        {
            foreach (Win_Qul_InspectCode_U_DefectProcess_Select_CodeView SelectProcess in dgdSelectProcess.Items)
            {
                SelectProcess.chkFlag = true;
            }
        }



        #region 포커스 이동용 키 다운 이벤트 모음
        // 포커스 이동용 키 다운 이벤트 모음
        private void txtDisplay1_KeyDown(object sender, System.Windows.Input.KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                txtDisplay2.Focus();
            }
        }
        private void txtDisplay2_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                txtDisplay3.Focus();
            }
        }
        private void txtDisplay3_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                txtKDefect.Focus();
            }
        }
        private void txtKDefect_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                txtEDefect.Focus();
            }
        }
        private void txtEDefect_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                txtTagName.Focus();
            }
        }
        private void txtTagName_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                cboDefect.Focus();
                cboDefect.IsDropDownOpen = true;
            }
        }
        private void cboDefect_DropDownClosed(object sender, EventArgs e)
        {
            cboSebuDefect.Focus();
            cboSebuDefect.IsDropDownOpen = true;
        }
        private void cboSebuDefect_DropDownClosed(object sender, EventArgs e)
        {
            //txtDisplay1.Focus();

            txtButtonSeq.Focus();
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
    }
}

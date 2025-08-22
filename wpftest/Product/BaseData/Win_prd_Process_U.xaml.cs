using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Data;
using System.Security.Cryptography;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using WizMes_WellMade.PopUp;
using WizMes_WellMade.PopUP;

namespace WizMes_WellMade
{
    /// <summary>
    /// Win_com_Process_U.xaml에 대한 상호 작용 논리
    /// </summary>
    public partial class Win_prd_Process_U : UserControl
    {
        string stDate = string.Empty;
        string stTime = string.Empty;

        Lib lib = new Lib();
        PlusFinder pf = MainWindow.pf;

        string strFlag = string.Empty;
        string strProcessID = string.Empty;
        int rowNum = 0;

        public Win_prd_Process_U()
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

            tgnSub.IsChecked = true;
            
        }

        #region 콤보박스 세팅
        /// <summary>
        /// 콤보박스 세팅 모두
        /// </summary>
        private void SetComboBox()
        {
            ObservableCollection<CodeView> ovcArticleGroupInput = ComboBoxUtil.Instance.Gf_DB_MT_sArticleGrp();
            this.cboArticleGrp.ItemsSource = ovcArticleGroupInput;
            this.cboArticleGrp.DisplayMemberPath = "code_name";
            this.cboArticleGrp.SelectedValuePath = "code_id";

            ObservableCollection<CodeView> ovcArticleGroup = ComboBoxUtil.Instance.Gf_DB_MT_sArticleGrp();
            this.cboArticleGrpSrh.ItemsSource = ovcArticleGroup;
            this.cboArticleGrpSrh.DisplayMemberPath = "code_name";
            this.cboArticleGrpSrh.SelectedValuePath = "code_id";

            ObservableCollection<CodeView> ovcProcessGroup = ComboBoxUtil.Instance.GetProcessGroup();
            cboProcessGroup.ItemsSource = ovcProcessGroup;
            cboProcessGroup.DisplayMemberPath = "code_name";
            cboProcessGroup.SelectedValuePath = "code_id";

            // 하위품 출고여부 Y N
            // cboChildCheckYN
            //라벨발행품 여부(입력)
            List<string[]> listYN = new List<string[]>();
            string[] YN01 = new string[] { "Y", "Y" };
            string[] YN02 = new string[] { "N", "N" };
            listYN.Add(YN01);
            listYN.Add(YN02);

            ObservableCollection<CodeView> ovcYN = ComboBoxUtil.Instance.Direct_SetComboBox(listYN);
            this.cboChildCheckYN.ItemsSource = ovcYN;
            this.cboChildCheckYN.DisplayMemberPath = "code_name";
            this.cboChildCheckYN.SelectedValuePath = "code_id";
        }
        #endregion

        #region 추가, 수정시 / 저장완료, 취소 시

        /// <summary>
        /// 저장완료, 취소 시
        /// </summary>
        private void CompleteCancelMode()
        {
            //grdInput.IsEnabled = false;
            Lib.Instance.UiButtonEnableChange_IUControl(this);

            txtCode.Visibility = Visibility.Visible;
            txtCode.IsEnabled = true;

            //cboProcessGroup.IsHitTestVisible = true;
            cboArticleGrp.IsHitTestVisible = false;
            txtSortSeq.IsHitTestVisible = false;
            txtProcess.IsHitTestVisible = false;
            txtEProcess.IsHitTestVisible = false;
            txtFProcess.IsHitTestVisible = false;
            cboChildCheckYN.IsHitTestVisible = false;
            chkUseClss.IsHitTestVisible = false;

            grdSrh1.IsHitTestVisible = true;
            grdSrh2.IsHitTestVisible = true;
            grdSrh3.IsHitTestVisible = true;

            dgdMain.IsHitTestVisible = true;
            dgdSub.IsHitTestVisible = true;
        }

        /// <summary>
        /// 추가, 수정 시
        /// </summary>
        private void SaveUpdateMode()
        {
            //grdInput.IsEnabled = true;
            Lib.Instance.UiButtonEnableChange_SCControl(this);

            if (strFlag.Trim().Equals("I"))
            {
                txtCode.Visibility = Visibility.Hidden;
                txtCodeFront.Visibility = Visibility.Visible;
                txtCodeBack.Visibility = Visibility.Visible;
            }
            else
            {
                txtCodeFront.Visibility = Visibility.Hidden;
                txtCodeBack.Visibility = Visibility.Hidden;
                txtCode.IsEnabled = false;
            }

            if (tgnMain.IsChecked == true) // 대분류 버튼 체크 시
            {
                txtCodeFront.IsEnabled = true;
                txtCodeBack.IsEnabled = false;
            }
            else // 소분류 버튼 체크 시
            {
                txtCodeFront.IsEnabled = false;
                txtCodeBack.IsEnabled = true;
            }

            //cboProcessGroup.IsHitTestVisible = true;
            cboArticleGrp.IsHitTestVisible = true;
            txtSortSeq.IsHitTestVisible = true;
            txtProcess.IsHitTestVisible = true;
            txtEProcess.IsHitTestVisible = true;
            txtFProcess.IsHitTestVisible = true;
            cboChildCheckYN.IsHitTestVisible = true;
            chkUseClss.IsHitTestVisible = true;

            grdSrh1.IsHitTestVisible = false;
            grdSrh2.IsHitTestVisible = false;
            grdSrh3.IsHitTestVisible = false;
        }

        #endregion // 추가, 수정시 / 저장완료, 취소 시

        #region 토글버튼 대분류 / 소분류

        /// <summary>
        /// 대분류 클릭 시
        /// </summary>
        private void tgnMain_Checked(object sender, RoutedEventArgs e)
        {           
            tgnMain.IsChecked = true;
            tgnSub.IsChecked = false;

            // SetComboBox() - 이 함수로 초기화 시
            // 바인딩된 ArticleGrpID 가 초기화되어, 데이터 그리드 객체값도 null 값이 되기 때문에
            // 값을 받은 다음 초기화
            //string ArticleGrpID = cboArticleGrp.SelectedValue != null ? cboArticleGrp.SelectedValue.ToString() : "";
            //SetComboBox();
            ObservableCollection<CodeView> ovcProcessGroup = ComboBoxUtil.Instance.GetProcessGroup();
            cboProcessGroup.ItemsSource = ovcProcessGroup;
            cboProcessGroup.DisplayMemberPath = "code_name";
            cboProcessGroup.SelectedValuePath = "code_id";

            // 소분류 클릭시에 공정분류 출력란을 활성화 시키기 위해서
            InputRHF.Height = new GridLength(0);
            InputRHL.Height = new GridLength(22, GridUnitType.Star);

            //dgdMain.SetValue(Grid.ColumnSpanProperty, 3);
            // 소분류 클릭시에 공정분류 출력란을 활성화 시키기 위해서
            grdDgdCol2.Width = new GridLength(0);
            grdDgdCol3.Width = new GridLength(0);

            grdDgd1_1.Height = new GridLength(0);
            grdDgd2_1.Height = new GridLength(0);

            // 대분류 선택시에, 오른쪽 상새내역을 다 비워주고, 선택된걸로 조지기
            this.DataContext = null;
            var Process = dgdMain.SelectedItem as Win_com_Process_U_CodeView;
            if (Process != null)
            {
                //Process.ArticleGrpid = ArticleGrpID;
                this.DataContext = Process;

                if (Process.UseClss.Trim().Equals("*"))
                {
                    chkUseClss.IsChecked = true;
                }
                else
                {
                    chkUseClss.IsChecked = false;
                }
            }
        }
        // 토글버튼을 라디오버튼처럼 쓰기 위해서
        private void tgnMain_PreviewMouseDown(object sender, System.Windows.Input.MouseButtonEventArgs e)
        {
            if (tgnMain.IsChecked == true)
            {
                e.Handled = true;
            }
        }
        private void tgnMain_Unchecked(object sender, RoutedEventArgs e)
        {
            e.Handled = true;
        }


        /// <summary>
        /// 소분류 클릭 시
        /// </summary>
        private void tgnSub_Checked(object sender, RoutedEventArgs e)
        {
            tgnMain.IsChecked = false;
            tgnSub.IsChecked = true;

            // SetComboBox() - 이 함수로 초기화 시
            // 바인딩된 ArticleGrpID 가 초기화되어, 데이터 그리드 객체값도 null 값이 되기 때문에
            // 값을 받은 다음 초기화
            //string ArticleGrpID = cboArticleGrp.SelectedValue != null ? cboArticleGrp.SelectedValue.ToString() : "";
            //SetComboBox();
            ObservableCollection<CodeView> ovcProcessGroup = ComboBoxUtil.Instance.GetProcessGroup();
            cboProcessGroup.ItemsSource = ovcProcessGroup;
            cboProcessGroup.DisplayMemberPath = "code_name";
            cboProcessGroup.SelectedValuePath = "code_id";

            // 소분류 클릭시에 공정분류 출력란을 활성화 시키기 위해서
            InputRHF.Height = new GridLength(22, GridUnitType.Star);
            InputRHL.Height = new GridLength(0);

            //dgdMain.SetValue(Grid.ColumnSpanProperty, 1);
            // 소분류 클릭시에 공정분류 출력란을 활성화 시키기 위해서
            grdDgdCol1.Width = new GridLength(318, GridUnitType.Star);
            grdDgdCol2.Width = new GridLength(1, GridUnitType.Star);
            grdDgdCol3.Width = new GridLength(320, GridUnitType.Star);

            grdDgd1_1.Height = new GridLength(22, GridUnitType.Star);
            grdDgd2_1.Height = new GridLength(22, GridUnitType.Star);

            // 대분류 - ProcessGrp 콤보박스 세팅
            var Process = dgdMain.SelectedItem as Win_com_Process_U_CodeView;
            if (Process != null && Process.ProcessID != null)
            {
                cboProcessGroup.SelectedValue = Process.ProcessID;

                // 조지기 전에 검색이 안되었으면 검색하기
                if (dgdSub.Items.Count < 1)
                {
                    FillGridSub(Process.ProcessID);
                }
            }

            // 소분류 선택시에, 오른쪽 상새내역을 다 비워주고, 선택된걸로 조지기
            this.DataContext = null;
            chkUseClss.IsChecked = false;
            var ProcessSub = dgdSub.SelectedItem as Win_com_Process_U_CodeView;
            if (ProcessSub != null)
            {
                //ProcessSub.ArticleGrpid = ArticleGrpID;
                this.DataContext = ProcessSub;

                if (ProcessSub.UseClss.Trim().Equals("*"))
                {
                    chkUseClss.IsChecked = true;
                }
                else
                {
                    chkUseClss.IsChecked = false;
                }
            }
        }

        private void tgnSub_PreviewMouseDown(object sender, System.Windows.Input.MouseButtonEventArgs e)
        {
            if (tgnSub.IsChecked == true)
            {
                e.Handled = true;
            }
        }

        private void tgnSub_Unchecked(object sender, RoutedEventArgs e)
        {
            e.Handled = true;
        }

        #endregion // 토글버튼 대분류 / 소분류

        #region Header - 검색조건

        // 공정그룹 - dgdMain 검색 : 공정명 Like 검색
        private void lblProcessSrh_MouseLeftButtonDown(object sender, System.Windows.Input.MouseButtonEventArgs e)
        {
            if (chkProcessSrh.IsChecked == true)
            {
                chkProcessSrh.IsChecked = false;
            }
            else
            {
                chkProcessSrh.IsChecked = true;
            }
        }
        private void chkProcessSrh_Checked(object sender, RoutedEventArgs e)
        {
            chkProcessSrh.IsChecked = true;
            txtProcessSrh.IsEnabled = true;
            txtProcessSrh.Focus();
        }
        private void chkProcessSrh_Unchecked(object sender, RoutedEventArgs e)
        {
            chkProcessSrh.IsChecked = false;
            txtProcessSrh.IsEnabled = false;
        }

        /// <summary>
        /// 제품 그룹 검색조건
        /// </summary>
        private void lblArticleGrpSrh_MouseLeftButtonDown(object sender, System.Windows.Input.MouseButtonEventArgs e)
        {
            if (chkArticleGrpSrh.IsChecked == true)
            {
                chkArticleGrpSrh.IsChecked = false;
            }
            else
            {
                chkArticleGrpSrh.IsChecked = true;
            }
        }
        private void chkArticleGrpSrh_Checked(object sender, RoutedEventArgs e)
        {
            chkArticleGrpSrh.IsChecked = true;
            cboArticleGrpSrh.IsEnabled = true;
            cboArticleGrpSrh.IsDropDownOpen = true;
        }
        private void chkArticleGrpSrh_Unchecked(object sender, RoutedEventArgs e)
        {
            chkArticleGrpSrh.IsChecked = false;
            cboArticleGrpSrh.IsEnabled = false;
        }

        // 사용안함 포함
        private void lblUseClssSrh_MouseLeftButtonDown(object sender, System.Windows.Input.MouseButtonEventArgs e)
        {
            if (chkUseClssSrh.IsChecked == true)
            {
                chkUseClssSrh.IsChecked = false;
            }
            else
            {
                chkUseClssSrh.IsChecked = true;
            }
        }

        #endregion // Header - 검색조건

        #region Header 오른쪽 상단 버튼

        //추가
        private void btnAdd_Click(object sender, RoutedEventArgs e)
        {
            if(tgnSub.IsChecked == true)
            {
                if(dgdMain.Items.Count == 0)
                {
                    MessageBox.Show("선택된 대분류 공정이 없습니다. 먼저 검색해주세요.");
                    return;
                }
            }
            
            var Process = dgdMain.SelectedItem as Win_com_Process_U_CodeView;
            string ProcessID = "";

            if (Process != null)
            {
                ProcessID = Process.ProcessID;
                if (Process.UseClss.Trim().Equals("*"))
                {
                    MessageBox.Show("사용 안하는 대분류 공정에 소분류 공정 추가는 불가능합니다.", "[추가 오류]");
                    return;
                }
            }

            strFlag = "I";
            SaveUpdateMode();

            //lblMsg.Visibility = Visibility.Visible;
            //tbkMsg.Text = "자료 입력 중";

            rowNum = dgdMain.SelectedIndex;

            this.DataContext = null;
            txtSortSeq.Text = "1";

            cboChildCheckYN.SelectedIndex = 0;

            if (tgnMain.IsChecked == true)
            {
                txtCodeBack.Text = "00";
            }
            else
            {
                if (ProcessID.Length > 3)
                {
                    txtCodeFront.Text = Process.ProcessID.Substring(0, 2);
                }
            }

            cboArticleGrp.Focus();
            cboArticleGrp.IsDropDownOpen = true;

        }

        //수정
        private void btnUpdate_Click(object sender, RoutedEventArgs e)
        {

            if (tgnMain.IsChecked == true)
            {
                var Process = dgdMain.SelectedItem as Win_com_Process_U_CodeView;

                if (Process == null)
                {
                    MessageBox.Show("수정할 대상을 선택해주세요.");
                    return;
                }
            }
            else
            {
                var ProcessSub = dgdSub.SelectedItem as Win_com_Process_U_CodeView;

                if (ProcessSub == null)
                {
                    MessageBox.Show("수정할 대상을 선택해주세요.");
                    return;
                }             
            }

            rowNum = dgdMain.SelectedIndex;

            strFlag = "U";
            SaveUpdateMode();
            dgdMain.IsHitTestVisible = false;
            dgdSub.IsHitTestVisible = false;
        }

        //삭제
        private void btnDelete_Click(object sender, RoutedEventArgs e)
        {
            var Process = new Win_com_Process_U_CodeView();

            if (tgnMain.IsChecked == true)
            {
                Process = dgdMain.SelectedItem as Win_com_Process_U_CodeView;
                if (Process == null)
                {
                    MessageBox.Show("삭제할 대분류 공정을 선택해주세요.");
                    return;
                }

                if(PlanCheck_ProcessID(0, Process.ProcessID) == true)
                {
                    MessageBox.Show("해당 대분류 공정이 포함된 작업지시서가 존재합니다.");
                    return;
                }
            }
            else
            {
                Process = dgdSub.SelectedItem as Win_com_Process_U_CodeView;
                if (Process == null)
                {
                    MessageBox.Show("삭제할 소분류 공정을 선택해주세요.");
                    return;
                }
                
                if (PlanCheck_ProcessID(1, Process.ProcessID) == true)
                {
                    MessageBox.Show("해당 소분류 공정이 포함된 작업지시서가 존재합니다.");
                    return;
                }
            }

            if (Process != null)
            {
                rowNum = dgdMain.SelectedIndex;

                if (MessageBox.Show("선택하신 항목을 삭제하시겠습니까?", "삭제 전 확인", MessageBoxButton.YesNo) == MessageBoxResult.Yes)
                {
                    using (Loading lw = new Loading(beDelete))
                    {
                        lw.ShowDialog();
                    }
                }
            }
        }

        private void beDelete()
        {
            var Process = new Win_com_Process_U_CodeView();

            if (tgnMain.IsChecked == true)
            {
                Process = dgdMain.SelectedItem as Win_com_Process_U_CodeView;
            }
            else
            {
                Process = dgdSub.SelectedItem as Win_com_Process_U_CodeView;
            }

            if (Process != null)
            {
                if (DeleteData(Process.ProcessID))
                {
                    re_Search();
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
            try
            {
                btnSearch.IsEnabled = false;

                rowNum = 0;
                using (Loading lw = new Loading(re_Search))
                {
                    lw.ShowDialog();
                }
            }
            catch(Exception ee)
            {
                MessageBox.Show("예외처리 - " + ee.ToString());
            }
            finally
            {
                btnSearch.IsEnabled = true;
            }

            
        }

        //저장
        private void btnSave_Click(object sender, RoutedEventArgs e)
        {
            using (Loading lw = new Loading(beSave))
            {
                lw.ShowDialog();
            }
        }

        private void beSave()
        {
            if (SaveData(strFlag, txtCodeFront.Text + txtCodeBack.Text))
            {

                //lblMsg.Visibility = Visibility.Hidden;                
                CompleteCancelMode();

                re_Search();

                strFlag = "";
            }
        }

        //취소
        private void btnCancel_Click(object sender, RoutedEventArgs e)
        {
            using (Loading lw = new Loading(beCancel))
            {
                lw.ShowDialog();
            }
        }

        private void beCancel()
        {
            CompleteCancelMode();
            strFlag = "";

            re_Search();
        }

        #region 엑셀
        private void btnExcel_Click(object sender, RoutedEventArgs e)
        {
            DataTable dt = null;
            string Name = string.Empty;

            string[] dgdStr = new string[2];
            dgdStr[0] = "공정코드";
            dgdStr[1] = dgdMain.Name;

            ExportExcelxaml ExpExc = new ExportExcelxaml(dgdStr);
            ExpExc.ShowDialog();

            if (ExpExc.DialogResult.HasValue)
            {
                if (ExpExc.choice.Equals(dgdMain.Name))
                {
                    DataStore.Instance.InsertLogByForm(this.GetType().Name, "P");
                    if (ExpExc.Check.Equals("Y"))
                        dt = Lib.Instance.DataGridToDTinHidden(dgdMain);
                    else
                        dt = Lib.Instance.DataGirdToDataTable(dgdMain);

                    Name = dgdMain.Name;
                    if (Lib.Instance.GenerateExcel(dt, Name))
                    {
                        Lib.Instance.excel.Visible = true;
                        lib.ReleaseExcelObject(lib.excel);
                    }
                    else
                    {
                        return;
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
        #endregion 엑셀

        #endregion // Header 오른쪽 상단 버튼

        #region Content

        /// <summary>
        /// ShowData
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void dgdMain_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            var Process = dgdMain.SelectedItem as Win_com_Process_U_CodeView;

            if (Process != null)
            {
                // 대분류 일 때는 메인 그리드 선택시에 항목 내용이 나오고, 
                // 소분류 일 때는 메인 그리드 선택시에 서브그리드 조회!!!!!!!!
                if (tgnMain.IsChecked == true)
                {
                    this.DataContext = Process;

                    if (Process.UseClss.Trim().Equals("*"))
                    {
                        chkUseClss.IsChecked = true;
                    }
                    else
                    {
                        chkUseClss.IsChecked = false;
                    }
                }
                else
                {
                    FillGridSub(Process.ProcessID);

                    // 서브그리드에 아무것도 없다면 오른쪽 상세 내용을 비워줘야 되겠지.
                    if (dgdSub.Items.Count == 0) 
                    {
                        this.DataContext = null;
                    }
                    else
                    {
                        dgdSub.SelectedIndex = 0;
                    }

                    cboProcessGroup.SelectedValue = Process.ProcessID;
                }
            }
        }

        // 서브그리드 선택시
        private void dgdSub_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (tgnSub.IsChecked == true)
            {
                var Process = dgdSub.SelectedItem as Win_com_Process_U_CodeView;
                if (Process != null)
                {
                    this.DataContext = Process;

                    if (Process.UseClss.Trim().Equals("*"))
                    {
                        chkUseClss.IsChecked = true;
                    }
                    else
                    {
                        chkUseClss.IsChecked = false;
                    }
                }
            }
        }

        // 공정 그룹 선택시 코드 세팅하기
        private void cboProcessGroup_DropDownClosed(object sender, EventArgs e)
        {
            if (cboProcessGroup.SelectedValue != null)
            {
                txtCodeFront.Text = cboProcessGroup.SelectedValue.ToString().Substring(0, 2);
            }
        }

        // 사용안함 라벨 → 체크
        private void lblUseClss_MouseLeftButtonDown(object sender, System.Windows.Input.MouseButtonEventArgs e)
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

        #endregion // Content

        #region 주요 메서드

        private void re_Search()
        {
            FillGrid();

            if (dgdMain.Items.Count > 0)
            {
                dgdMain.SelectedIndex = strFlag.Trim().Equals("I") && tgnMain.IsChecked == true ? dgdMain.Items.Count - 1 : rowNum;
            }
            else
            {
                MessageBox.Show("조회된 내용이 없습니다.");
            }
        }

        #region 대분류 조회

        private void FillGrid()
        {
            dgdSub.Items.Clear();

            if (dgdMain.Items.Count > 0)
            {
                dgdMain.Items.Clear();
            }

            try
            {
                Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                sqlParameter.Clear();

                sqlParameter.Add("nProcess", chkProcessSrh.IsChecked == true ? 1 : 0);
                sqlParameter.Add("Process", chkProcessSrh.IsChecked == true && !txtProcessSrh.Text.Trim().Equals("") ? txtProcessSrh.Text : "");
                sqlParameter.Add("nArticleGrpID", chkArticleGrpSrh.IsChecked == true ? 1 : 0);
                sqlParameter.Add("ArticleGrpID", chkArticleGrpSrh.IsChecked == true && cboArticleGrpSrh.SelectedValue != null ? cboArticleGrpSrh.SelectedValue.ToString() : "");
                sqlParameter.Add("UseClss", chkUseClssSrh.IsChecked == true ? 1: 0);

                DataSet ds = DataStore.Instance.ProcedureToDataSet_LogWrite("xp_prd_sProcessMain", sqlParameter, true, "R");

                if (ds != null && ds.Tables.Count > 0)
                {
                    DataTable dt = ds.Tables[0];
                    int i = 0;

                    if (dt.Rows.Count > 0)
                    {
                        DataRowCollection drc = dt.Rows;

                        foreach (DataRow dr in drc)
                        {
                            i++;

                            var Process = new Win_com_Process_U_CodeView()
                            {
                                Num = i,
                                ProcessID = dr["ProcessID"].ToString(),
                                Process = dr["Process"].ToString(),
                                ParentProcessID = dr["ParentProcessID"].ToString(),
                                ParentProcessName = dr["ParentProcessName"].ToString(),
                                ArticleGrpid = dr["ArticleGrpid"].ToString().Replace(" ", ""),
                                ArticleGrp = dr["ArticleGrp"].ToString(),
                                UseClss = dr["UseClss"].ToString(),
                                SortSeq = Convert.ToDouble(dr["SortSeq"]),
                                EProcess = dr["EProcess"].ToString(),
                                FProcess = dr["FProcess"].ToString(),
                                ProcessID_Front = dr["ProcessID"].ToString().Trim().Substring(0, 2),
                                ProcessID_Back = dr["ProcessID"].ToString().Trim().Substring(2, 2),
                                ChildCheckYN = dr["ChildCheckYN"].ToString().Trim().Equals("") ? "N" : dr["ChildCheckYN"].ToString(),
                            };

                            // 사용안함이면 글자색을 붉은색으로
                            if (Process.UseClss.Trim().Equals("*"))
                            {
                                Process.FontColor_UseClssN = true;
                            }

                            dgdMain.Items.Add(Process);

                        }
                    }

                    tbkCount.Text = " ▶ 검색 결과 : " + i + " 건";
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

        #endregion // 대분류 조회

        #region 소분류 조회

        private void FillGridSub(string strID)
        {
            if (dgdSub.Items.Count > 0)
            {
                dgdSub.Items.Clear();
            }

            try
            {
                Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                sqlParameter.Clear();

                sqlParameter.Add("ProcessID", strID);
                sqlParameter.Add("nArticleGrpID", 0); // 소분류 검색조건 구현은 어떻게??
                sqlParameter.Add("ArticleGrpID", "");
                sqlParameter.Add("UseClss", chkUseClssSrh.IsChecked == true ? 1 : 0);

                DataSet ds = DataStore.Instance.ProcedureToDataSet("xp_prd_sProcessSub", sqlParameter, false);

                if (ds != null && ds.Tables.Count > 0)
                {
                    DataTable dt = ds.Tables[0];
                    int i = 0;

                    DataRowCollection drc = dt.Rows;

                    foreach (DataRow dr in drc)
                    {

                        i++;

                        var Process = new Win_com_Process_U_CodeView()
                        {
                            Num = i,
                            ProcessID = dr["ProcessID"].ToString(),
                            Process = dr["Process"].ToString(),
                            ParentProcessID = dr["ParentProcessID"].ToString(),
                            ParentProcessName = dr["ParentProcessName"].ToString(),
                            ArticleGrpid = dr["ArticleGrpid"].ToString().Replace(" ", ""),
                            ArticleGrp = dr["ArticleGrp"].ToString(),
                            UseClss = dr["UseClss"].ToString(),
                            SortSeq = Convert.ToDouble(dr["SortSeq"]),
                            EProcess = dr["EProcess"].ToString(),
                            FProcess = dr["FProcess"].ToString(),
                            ProcessID_Front = dr["ProcessID"].ToString().Trim().Substring(0, 2),
                            ProcessID_Back = dr["ProcessID"].ToString().Trim().Substring(2, 2),
                            ChildCheckYN = dr["ChildCheckYN"].ToString().Trim().Equals("") ? "N" : dr["ChildCheckYN"].ToString(),
                        };

                        // 사용안함이면 글자색을 붉은색으로
                        if (Process.UseClss.Trim().Equals("*"))
                        {
                            Process.FontColor_UseClssN = true;
                        }

                        dgdSub.Items.Add(Process);
                    }

                    tbkSubCount.Text = " ▶ 검색 결과 : " + i + " 건";
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

        #endregion // 소분류 조회

        #region 저장

        /// <summary>
        /// 저장
        /// </summary>
        private bool SaveData(string strFlag, string strID)
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
                    sqlParameter.Add("JobFlag", strFlag.Trim());
                    sqlParameter.Add("ProcessID", strID);
                    sqlParameter.Add("Process", txtProcess.Text);
                    sqlParameter.Add("EProcess", txtEProcess.Text);
                    sqlParameter.Add("FProcess", txtFProcess.Text);

                    sqlParameter.Add("ArticleGrpid", cboArticleGrp.SelectedValue != null ? cboArticleGrp.SelectedValue.ToString() : "");
                    sqlParameter.Add("SortSeq", ConvertInt(txtSortSeq.Text));
                    sqlParameter.Add("ProcessGrpID", tgnSub.IsChecked == true && cboProcessGroup.SelectedValue != null ? cboProcessGroup.SelectedValue.ToString() : "");
                    sqlParameter.Add("UseClss", chkUseClss.IsChecked == true ? "*" : "");
                    sqlParameter.Add("ChildCheckYN", cboChildCheckYN.SelectedValue != null ? cboChildCheckYN.SelectedValue.ToString() : "");

                    sqlParameter.Add("sUserID", MainWindow.CurrentUser);

                    Procedure pro1 = new Procedure();
                    pro1.Name = "xp_prd_iuProcess";
                    pro1.OutputUseYN = "N";
                    pro1.OutputName = "ProcessID";
                    pro1.OutputLength = "4";

                    Prolist.Add(pro1);
                    ListParameter.Add(sqlParameter);

                    string[] Confirm = new string[2];
                    Confirm = DataStore.Instance.ExecuteAllProcedureOutputNew_NewLog(Prolist, ListParameter,"C");
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
                MessageBox.Show("오류 발생, 오류 내용 : " + ex.ToString());
            }
            finally
            {
                DataStore.Instance.CloseConnection();
            }

            return flag;
        }

        #endregion // 저장

        #region 유효성 검사

        private bool CheckData()
        {
            bool flag = true;

            // 

            if (tgnMain.IsChecked == true)
            {
                // 입력하지 않았을 시, 
                if (txtCodeFront.Text.Trim().Length != 2)
                {
                    MessageBox.Show("공정코드에 두 자리의 숫자를 입력해주세요.");
                    flag = false;
                    return flag;
                }

                // 숫자를 입력하지 않았을 시
                if (CheckConvertInt(txtCodeFront.Text.Trim()) == false)
                {
                    MessageBox.Show("공정코드는 숫자만 입력 가능합니다. \r(공백을 제외한 두자리의 숫자 입력 필요)");
                    flag = false;
                    return flag;
                }

                string ProcessID = txtCodeFront.Text.Trim() + txtCodeBack.Text.Trim();
                if (strFlag.Trim().Equals("I") 
                    && !ChkProcessID(ProcessID))
                {
                    MessageBox.Show("해당 공정코드는 이미 존재 합니다.");
                    flag = false;
                    return flag;
                }
            }
            else
            {
                // 입력하지 않았을 시, 
                if (txtCodeBack.Text.Trim().Length != 2)
                {
                    MessageBox.Show("공정코드에 두 자리의 숫자를 입력해주세요.");
                    flag = false;
                    return flag;
                }

                // 숫자를 입력하지 않았을 시
                if (CheckConvertInt(txtCodeBack.Text.Trim()) == false)
                {
                    MessageBox.Show("공정코드는 숫자만 입력 가능합니다. \r(공백을 제외한 두자리의 숫자 입력 필요)");
                    flag = false;
                    return flag;
                }

                if (txtCodeBack.Text.Trim().Equals("00"))
                {
                    MessageBox.Show("00 이상의 숫자를 입력해주세요.");
                    flag = false;
                    return flag;
                }

                string ProcessID = txtCodeFront.Text.Trim() + txtCodeBack.Text.Trim();
                if (strFlag.Trim().Equals("I") && !ChkProcessID(ProcessID))
                {
                    MessageBox.Show("해당 공정코드는 이미 존재 합니다.");
                    flag = false;
                    return flag;
                }
            }

            // 정렬순서
            if (CheckConvertDouble(txtSortSeq.Text) == false)
            {
                MessageBox.Show("정렬 순서는 숫자만 입력 가능합니다.");
                flag = false;
                return flag;
            }

            return flag;
        }

        #region 공정ID 
        private bool ChkProcessID(string ProcessID)
        {
            bool flag = true;

            try
            {
                Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                sqlParameter.Clear();

                sqlParameter.Add("ProcessID", ProcessID);

                DataSet ds = DataStore.Instance.ProcedureToDataSet("xp_prd_sChkProcessID", sqlParameter, false);

                if (ds != null && ds.Tables.Count > 0)
                {
                    DataTable dt = ds.Tables[0];
                    int i = 0;

                    if (dt.Rows.Count != 0)
                    {
                        DataRowCollection drc = dt.Rows;
                        DataRow dr = drc[0];
                        int Cnt = ConvertInt(dr["Cnt"].ToString());

                        if (Cnt > 0)
                        {
                            return false;
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

            return flag;
        }

        #endregion

        #endregion // 유효성 검사

        #region 삭제

        private bool DeleteData(string strID)
        {
            bool flag = false;

            try
            {
                Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                sqlParameter.Clear();
                sqlParameter.Add("ProcessID", strID);
                sqlParameter.Add("UpdateUserID", MainWindow.CurrentUser);

                string[] result = DataStore.Instance.ExecuteProcedure_NewLog("xp_prd_dProcess", sqlParameter, "D");
                DataStore.Instance.CloseConnection();

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

        #endregion // 삭제

        #endregion // 주요 메서드

        #region 텍스트 박스 공통 키다운 이벤트

        // 검색조건 - 텍스트 박스 엔터 → 조회
        private void txtBox_EnterAndSearch(object sender, System.Windows.Input.KeyEventArgs e)
        {
            rowNum = 0;
            using (Loading lw = new Loading(re_Search))
            {
                lw.ShowDialog();
            }
        }


        // 텍스트박스 숫자만 입력 가능하도록
        private void txtBox_PreviewTextInput(object sender, TextCompositionEventArgs e)
        {
            CheckIsNumericOnly((TextBox)sender, e);
        }

        #endregion

        #region 기타 메서드 모음

        // 텍스트 박스 숫자만 입력 되도록
        public void CheckIsNumericOnly(TextBox sender, TextCompositionEventArgs e)
        {
            decimal result;
            if (!(Decimal.TryParse(e.Text, out result)))
            {
                e.Handled = true;
            }
        }

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



        #endregion

        #region 테스트용
        // 테스트 키다운
        private void txtTest_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                Console.WriteLine(Test_Go(txtEProcess.Text));
                Console.WriteLine(encryptSHA256(""));
            }
        }

        //
        private string encryptSHA256(string str)
        {
            string result = "";

            try
            {
                Console.WriteLine(Test_Go("7152"));
            }
            catch (Exception ex)
            {

            }

            return result;
        }

        private string Test_Go(string data)
        {
            SHA256 sha = new SHA256Managed();
            byte[] hash = sha.ComputeHash(Encoding.ASCII.GetBytes(data));
            StringBuilder stringBuilder = new StringBuilder();
            foreach (byte b in hash)
            {
                stringBuilder.AppendFormat("{0:x2}", b);
            }
            return stringBuilder.ToString();
        }
        #endregion


        private bool PlanCheck_ProcessID(int category, string processID)
        {
            bool result = false;

            try
            {
                Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                sqlParameter.Clear();
                sqlParameter.Add("Category", category);
                sqlParameter.Add("ProcessID", processID);

                DataSet ds = DataStore.Instance.ProcedureToDataSet("xp_prd_Process_PlanCheck", sqlParameter, false);

                if(ds != null && ds.Tables.Count > 0)
                {
                    DataTable dt = ds.Tables[0];

                    if(dt.Rows.Count > 0)
                    {
                        //System.Diagnostics.Debug.WriteLine("있음" + dt.Rows.Count.ToString());
                        result = true;
                    }
                    else
                    {
                        //System.Diagnostics.Debug.WriteLine("없음" + dt.Rows.Count.ToString());
                        result = false;
                    }
                }
            }
            catch(Exception ee)
            {
                MessageBox.Show("예외처리 - " + ee.ToString());
            }
            finally
            {
                DataStore.Instance.CloseConnection();
            }

            return result;
        }

        private void cboArticleGrp_DropDownClosed(object sender, EventArgs e)
        {
            try
            {
                //대분류
                if(tgnMain.IsChecked == true)
                {
                    txtCodeFront.Focus();
                }
                else if(tgnSub.IsChecked == true)
                {
                    txtCodeBack.Focus();
                }
            }
            catch(Exception ee)
            {
                MessageBox.Show("예외처리 - " + ee.ToString());
            }
        }

        private void txtCodeFrontBack_KeyDown(object sender, KeyEventArgs e)
        {
            try
            {
                if(e.Key == Key.Enter)
                {
                    txtSortSeq.Focus();
                }
            }
            catch(Exception ee)
            {
                MessageBox.Show("예외처리 - " + ee.ToString());
            }
        }

        #region 입력창 이동 이벤트
        //텍스트박스
        private void EnterMoveTextBox_KeyDown(object sender, KeyEventArgs e)
        {
            try
            {
                if (e.Key == Key.Enter)
                {
                    lib.SendK(Key.Tab, this);
                }
            }
            catch (Exception ee)
            {
                MessageBox.Show("예외처리 - " + ee.ToString());
            }
        }

        //콤보박스일때
        private void EnterMoveComboBox_DropDownClosed(object sender, EventArgs e)
        {
            try
            {
                lib.SendK(Key.Tab, this);
            }
            catch (Exception ee)
            {
                MessageBox.Show("예외처리 - " + ee.ToString());
            }
        }

        #endregion

    }

    // 코드뷰
    class Win_com_Process_U_CodeView : BaseView
    {
        public override string ToString()
        {
            return (this.ReportAllProperties());
        }

        public int Num { get; set; }
        public string ProcessID { get; set; }
        public string Process { get; set; }
        public string ParentProcessID { get; set; }
        public string ParentProcessName { get; set; }
        public string ArticleGrpid { get; set; }
        public string ArticleGrp { get; set; }
        public string UseClss { get; set; }
        public double SortSeq { get; set; }
        public string EProcess { get; set; }
        public string FProcess { get; set; }

        public string ProcessID_Front { get; set; }
        public string ProcessID_Back { get; set; }

        public string ChildCheckYN { get; set; }

        public bool FontColor_UseClssN { get; set; }

        public Win_com_Process_U_CodeView Copy()
        {
            return (Win_com_Process_U_CodeView)this.MemberwiseClone();
        }
    }
}

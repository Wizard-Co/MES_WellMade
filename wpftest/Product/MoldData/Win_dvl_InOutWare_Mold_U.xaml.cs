using WizMes_WellMade.PopUP;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Data;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using WPF.MDI;
using Tesseract;
using System.Drawing;
using static System.Windows.Forms.AxHost;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.TaskbarClock;

namespace WizMes_WellMade
{
    /// <summary>
    /// Win_dvl_InOutWare_Mold_U.xaml에 대한 상호 작용 논리
    /// </summary>
    public partial class Win_dvl_InOutWare_Mold_U : UserControl
    {
        #region 변수선언 및 로드

        Lib lib = new Lib();
        PlusFinder pf = new PlusFinder();
        Win_dvl_InOutWare_Mold_U_CodeView winInOut_DTO = new Win_dvl_InOutWare_Mold_U_CodeView();
        int numSaveRowCount = 0;
        string strFlag = string.Empty;
        List<string> lstCompareValue = new List<string>();
        Dictionary<string, object> dicCompare = new Dictionary<string, object>();

        public Win_dvl_InOutWare_Mold_U()
        {
            InitializeComponent();
        }
        
        private void UserControl_Loaded(object sender, RoutedEventArgs e)
        {
            chkTurnInDay.IsChecked = true;
            dtpSDate.SelectedDate = DateTime.Today;
            dtpEDate.SelectedDate = DateTime.Today;

            SetComboBox();

            lib.UiLoading(sender);
        }

        #endregion

        #region 상단 체크박스 이벤트

        //제출일자
        private void lblTurnInDay_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            if (chkTurnInDay.IsChecked == true) { chkTurnInDay.IsChecked = false; }
            else { chkTurnInDay.IsChecked = true; }
        }

        //제출일자
        private void chkTurnInDay_Checked(object sender, RoutedEventArgs e)
        {
            this.dtpSDate.IsEnabled = true;
            this.dtpEDate.IsEnabled = true;
        }

        //제출일자
        private void chkTurnInDay_Unchecked(object sender, RoutedEventArgs e)
        {
            this.dtpSDate.IsEnabled = false;
            this.dtpEDate.IsEnabled = false;
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
            dtpSDate.SelectedDate = lib.BringThisMonthDatetimeList()[0];
            dtpEDate.SelectedDate = lib.BringThisMonthDatetimeList()[1];
        }

        //금형이라는데 뭐지??
        private void lblLotNoSrh_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            if (chkLotNoSrh.IsChecked == true) { chkLotNoSrh.IsChecked = false; }
            else { chkLotNoSrh.IsChecked = true; }
        }

        //금형이라는데 뭐지??
        private void chkLotNoSrh_Checked(object sender, RoutedEventArgs e)
        {
            txtLotNoSrh.IsEnabled = true;
            btnPfLotNoSrh.IsEnabled = true;
            txtLotNoSrh.Focus();
        }

        //금형이라는데 뭐지??
        private void chkLotNoSrh_Unchecked(object sender, RoutedEventArgs e)
        {
            txtLotNoSrh.IsEnabled = false;
            btnPfLotNoSrh.IsEnabled = false;
        }

        //금형이라는데 뭐지??
        private void txtLotNoSrh_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                pf.ReturnCode(txtLotNoSrh, 51, "");
            }
        }

        //금형이라는데 뭐지??
        private void btnPfLotNoSrh_Click(object sender, RoutedEventArgs e)
        {
            pf.ReturnCode(txtLotNoSrh, 51, "");
        }

        //품명
        private void lblArticleSrh_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            if (chkArticleSrh.IsChecked == true) { chkArticleSrh.IsChecked = false; }
            else { chkArticleSrh.IsChecked = true; }
        }

        //품명
        private void chkArticleSrh_Checked(object sender, RoutedEventArgs e)
        {
            txtArticleSrh.IsEnabled = true;
            btnPfArticleSrh.IsEnabled = true;
            txtArticleSrh.Focus();
        }

        //품명
        private void chkArticleSrh_Unchecked(object sender, RoutedEventArgs e)
        {
            txtArticleSrh.IsEnabled = false;
            btnPfArticleSrh.IsEnabled = false;
        }

        //품명
        private void txtArticleSrh_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                pf.ReturnCode(txtArticleSrh, 1, "");
            }
        }

        //품명
        private void btnPfArticleSrh_Click(object sender, RoutedEventArgs e)
        {
            pf.ReturnCode(txtArticleSrh, 1, "");
        }

        //구분
        private void lblGubunSrh_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            if (chkGubunSrh.IsChecked == true) { chkGubunSrh.IsChecked = false; }
            else { chkGubunSrh.IsChecked = true; }
        }

        //구분
        private void chkGubunSrh_Checked(object sender, RoutedEventArgs e)
        {
            cboGubunSrh.IsEnabled = true;
            cboGubunSrh.Focus();
        }

        //구분
        private void chkGubunSrh_Unchecked(object sender, RoutedEventArgs e)
        {
            cboGubunSrh.IsEnabled = false;
        }

        #endregion

        #region 상단 우측 버튼 이벤트

        //조회
        private void btnSearch_Click(object sender, RoutedEventArgs e)
        {
            numSaveRowCount = 0;

            re_Search(numSaveRowCount);
        }

        //추가
        private void btnAdd_Click(object sender, RoutedEventArgs e)
        {
            ControlVisibleAndEnable_AU();
            this.DataContext = null;
            dtpDay.SelectedDate = DateTime.Today;

            //txtBarCode.Clear();
            //txtInOutID.Clear();
            //txtMoldID.Clear();
            //txtMoldNo.Clear();
            cboGubun.SelectedIndex = 0;
            cboInOutPlace.SelectedIndex = 0;
            //txtPerson.Clear();
            //txtInOutQty.Clear();
            //txtComments.Clear();

            tbkMsg.Text = "자료 입력(추가) 중";
            strFlag = "I";
            dtpDay.Focus();
        }

        //수정
        private void btnUpdate_Click(object sender, RoutedEventArgs e)
        {
            var WinInOut = dgdMoldInOut.SelectedItem as Win_dvl_InOutWare_Mold_U_CodeView;

            if (WinInOut == null)
            {
                MessageBox.Show("수정할 자료를 선택하고 눌러주십시오.");
            }
            else
            {
                ControlVisibleAndEnable_AU();

                tbkMsg.Text = "자료 입력(수정) 중";
                strFlag = "U";
                dtpDay.Focus();
            }
        }

        //삭제
        private void btnDelete_Click(object sender, RoutedEventArgs e)
        {
            winInOut_DTO = dgdMoldInOut.SelectedItem as Win_dvl_InOutWare_Mold_U_CodeView;
            numSaveRowCount = dgdMoldInOut.SelectedIndex;

            if (winInOut_DTO == null)
            {
                MessageBox.Show("삭제할 데이터가 지정되지 않았습니다. 삭제데이터를 지정하고 눌러주세요");
                return;
            }
            else
            {
                if (MessageBox.Show("선택하신 항목을 삭제하시겠습니까?", "삭제 전 확인", MessageBoxButton.YesNo) == MessageBoxResult.Yes)
                {
                    if (dgdDelete(winInOut_DTO.InOutID))
                    {
                        re_Search(numSaveRowCount - 1);
                    }
                }
            }
        }

        //저장
        private void btnSave_Click(object sender, RoutedEventArgs e)
        {
            if (strFlag.Equals("I"))
            {
                numSaveRowCount = 0;

                if (dgdInsert())
                {
                    ControlVisibleAndEnable_SC();

                    re_Search(numSaveRowCount);
                }
            }
            else
            {
                numSaveRowCount = dgdMoldInOut.SelectedIndex;

                if (dgdUpdate())
                {
                    ControlVisibleAndEnable_SC();

                    re_Search(numSaveRowCount);
                }
            }
        }

        //취소
        private void btnCancel_Click(object sender, RoutedEventArgs e)
        {
            InputClear();
            ControlVisibleAndEnable_SC();

            re_Search(numSaveRowCount);
        }

        //입력 데이터 클리어
        private void InputClear()
        {
            foreach (Control child in this.grdInput.Children)
            {
                if (child.GetType() == typeof(TextBox))
                    ((TextBox)child).Clear();
                else if (child.GetType() == typeof(ComboBox))
                    ((ComboBox)child).SelectedIndex = -1;
            }

            txtBarCode.Clear();
        }

        //엑셀
        private void btnExcel_Click(object sender, RoutedEventArgs e)
        {
            DataTable dt = null;
            string Name = string.Empty;

            string[] dgdStr = new string[2];
            dgdStr[0] = "금형입출고관리";
            dgdStr[1] = dgdMoldInOut.Name;

            ExportExcelxaml ExpExc = new ExportExcelxaml(dgdStr);
            ExpExc.ShowDialog();

            if (ExpExc.DialogResult.HasValue)
            {
                if (ExpExc.choice.Equals(dgdMoldInOut.Name))
                {
                    if (ExpExc.Check.Equals("Y"))
                        dt = lib.DataGridToDTinHidden(dgdMoldInOut);
                    else
                        dt = lib.DataGirdToDataTable(dgdMoldInOut);

                    Name = dgdMoldInOut.Name;

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

        //닫기
        private void btnClose_Click(object sender, RoutedEventArgs e)
        {
            string stDate = DateTime.Now.ToString("yyyyMMdd");
            string stTime = DateTime.Now.ToString("HHmm");
            DataStore.Instance.InsertLogByFormS(this.GetType().Name, stDate, stTime, "E");
            Lib.Instance.ChildMenuClose(this.ToString());
        }

        //추가,수정 시 동작 모음
        private void ControlVisibleAndEnable_AU()
        {
            lib.UiButtonEnableChange_SCControl(this);
            //btnSave.Visibility = Visibility.Visible;
            //btnCancel.Visibility = Visibility.Visible;
            //btnAdd.IsEnabled = false;
            //btnUpdate.IsEnabled = false;
            //btnDelete.IsEnabled = false;
            dgdMoldInOut.IsEnabled = false;

            bdrLeft.IsEnabled = true;
            //lblMsg.Visibility = Visibility.Visible;
            //btnExcel.Visibility = Visibility.Hidden;
        }

        //저장,취소 시 동작 모음
        private void ControlVisibleAndEnable_SC()
        {
            lib.UiButtonEnableChange_IUControl(this);
            //btnSave.Visibility = Visibility.Hidden;
            //btnCancel.Visibility = Visibility.Hidden;
            //btnAdd.IsEnabled = true;
            //btnUpdate.IsEnabled = true;
            //btnDelete.IsEnabled = true;
            dgdMoldInOut.IsEnabled = true;

            bdrLeft.IsEnabled = false;
            //lblMsg.Visibility = Visibility.Hidden;
            //btnExcel.Visibility = Visibility.Visible;
        }

        #endregion

        #region 중단 플러스파인더 및 enter focus move

        //일자
        private void dtpDay_PreviewKeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                dtpDay.IsDropDownOpen = true;
            }
        }

        //일자
        private void dtpDay_CalendarClosed(object sender, RoutedEventArgs e)
        {
            //txtMoldNo.Focus();
            txtMoldNo.Focus();
        }


        //금형LotNo
        private void txtMoldNo_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                pf.ReturnCode(txtMoldNo, 51, "");

                if (txtMoldNo.Tag != null && !txtMoldNo.Tag.ToString().Equals(""))
                {
                    txtMoldID.Text = txtMoldNo.Tag.ToString();
                }

                cboGubun.Focus();
                cboGubun.IsDropDownOpen = true;
            }
        }

        //금형LotNo
        private void btnMoldLotNo_Click(object sender, RoutedEventArgs e)
        {
            pf.ReturnCode(txtMoldNo, 51, "");

            if (txtMoldNo.Tag != null && !txtMoldNo.Tag.ToString().Equals(""))
            {
                txtMoldID.Text = txtMoldNo.Tag.ToString();
            }

            cboGubun.Focus();
            cboGubun.IsDropDownOpen = true;
        }

        //구분
        private void cboGubun_DropDownClosed(object sender, EventArgs e)
        {
            cboInOutPlace.Focus();
            cboInOutPlace.IsDropDownOpen = true;
        }

        //보관장소
        private void cboInOutPlace_DropDownClosed(object sender, EventArgs e)
        {
            txtPerson.Focus();
        }

        //담당자
        private void txtPerson_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                pf.ReturnCode(txtPerson, 2, "");
                txtInOutQty.Focus();
            }
        }

        //담당자
        private void btnPfPerson_Click(object sender, RoutedEventArgs e)
        {
            pf.ReturnCode(txtPerson, 2, "");
            txtInOutQty.Focus();
        }

        //수량
        private void txtInOutQty_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                txtComments.Focus();
            }
        }

        #endregion

        //콤보박스 세팅
        private void SetComboBox()
        {
            ObservableCollection<CodeView> ovcGubunSrh = ComboBoxUtil.Instance.Gf_DB_CM_GetComCodeDataset(null, "DVLGBN", "Y", "");
            this.cboGubunSrh.ItemsSource = ovcGubunSrh;
            this.cboGubunSrh.DisplayMemberPath = "code_name";
            this.cboGubunSrh.SelectedValuePath = "code_id";

            this.cboGubun.ItemsSource = ovcGubunSrh;
            this.cboGubun.DisplayMemberPath = "code_name";
            this.cboGubun.SelectedValuePath = "code_id";

            ObservableCollection<CodeView> ovcInOutPlace = ComboBoxUtil.Instance.Gf_DB_CM_GetComCodeDataset(null, "MOLDPLACE", "Y", "","");
            this.cboInOutPlace.ItemsSource = ovcInOutPlace;
            this.cboInOutPlace.DisplayMemberPath = "code_name";
            this.cboInOutPlace.SelectedValuePath = "code_id";
        }

        //수정,삭제,조회 selectIndex를 받은 후 조회
        private void re_Search(int selectIndex)
        {
            if (dgdMoldInOut.Items.Count > 0)
            {
                dgdMoldInOut.Items.Clear();
            }

            FillGrid();

            if (dgdMoldInOut.Items.Count > 0)
            {
                if (lstCompareValue.Count > 0)
                {
                    dgdMoldInOut.SelectedIndex = lib.reTrunIndex(dgdMoldInOut, lstCompareValue[0]);
                }
                else
                {
                    dgdMoldInOut.SelectedIndex = selectIndex; ;
                }
            }
            else
            {
                InputClear();
            }

            dicCompare.Clear();
            lstCompareValue.Clear();
        }

        #region DB 통한 데이터 CRUD

        //실질 조회
        private void FillGrid()
        {
            bool flag = true;

            if (chkGubunSrh.IsChecked == true)
            {
                if (cboGubunSrh.SelectedValue == null)
                {
                    MessageBox.Show("구분에 체크된 상태에서 구분자가 선택되지 않았습니다. 선택해주세요");
                    flag = false;
                }
            }

            if (flag)
            {
                try
                {
                    Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                    sqlParameter.Add("chkDate", chkTurnInDay.IsChecked == true ? 1 : 0);
                    sqlParameter.Add("FromDate", chkTurnInDay.IsChecked == true ? dtpSDate.SelectedDate.Value.ToString("yyyyMMdd") : "");
                    sqlParameter.Add("ToDate", chkTurnInDay.IsChecked == true ? dtpEDate.SelectedDate.Value.ToString("yyyyMMdd") : "");
                    sqlParameter.Add("nchkMold", chkLotNoSrh.IsChecked == true ? (txtLotNoSrh.Tag != null ? 1 : 2) : 0);
                    sqlParameter.Add("MoldID", chkLotNoSrh.IsChecked == true ? (txtLotNoSrh.Tag != null ? txtLotNoSrh.Tag.ToString() : txtLotNoSrh.Text) : "");
                    sqlParameter.Add("nchkArticle", chkArticleSrh.IsChecked == true ? (txtArticleSrh.Tag != null ? 1 : 2) : 0);
                    sqlParameter.Add("ArticleID", chkArticleSrh.IsChecked == true ? (txtArticleSrh.Tag != null ? txtArticleSrh.Tag.ToString() : txtArticleSrh.Text) : "");
                    sqlParameter.Add("nInOutGbn", chkGubunSrh.IsChecked == true ? 1 : 0);
                    sqlParameter.Add("InOutGbn", (chkGubunSrh.IsChecked == true ? cboGubunSrh.SelectedValue.ToString() : ""));

                    DataSet ds = DataStore.Instance.ProcedureToDataSet("xp_dvlMold_sMoldInOut", sqlParameter, false);

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
                                var dgdInOut_DTO = new Win_dvl_InOutWare_Mold_U_CodeView()
                                {
                                    Num = i,
                                    InOutID = dr["InOutID"].ToString(),
                                    InOutGbn = dr["InOutGbn"].ToString(),
                                    InOutDate = dr["InOutDate"].ToString(),
                                    MoldID = dr["MoldID"].ToString(),
                                    MoldNo = dr["MoldNo"].ToString(),
                                    InOutPlace = dr["InOutPlace"].ToString().Replace(" ", ""),
                                    InOutQty = dr["InOutQty"].ToString(),
                                    InOutPerson = dr["InOutPerson"].ToString(),
                                    PersonName = dr["PersonName"].ToString(),
                                    Comments = dr["Comments"].ToString(),
                                    MoldName = dr["MoldName"].ToString(),
                                    Place = dr["Place"].ToString(),
                                    InOutName = dr["InOutName"].ToString(),
                                    ArticleID = dr["ArticleID"].ToString()
                                };

                                dgdInOut_DTO.InOutDate = lib.StrDateTimeBar(dgdInOut_DTO.InOutDate);

                                if (dicCompare.Count > 0)
                                {
                                    if (dgdInOut_DTO.InOutID.Equals(dicCompare["InOutID"].ToString()))
                                    {
                                        lstCompareValue.Add(dgdInOut_DTO.ToString());
                                    }
                                }

                                dgdMoldInOut.Items.Add(dgdInOut_DTO);
                            }
                        }
                        else
                        {
                            MessageBox.Show("조회된 결과가 없습니다.");
                            return;
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
        }

        //this.DataContext
        private void dgdMoldInOut_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            Win_dvl_InOutWare_Mold_U_CodeView winInOutWare = dgdMoldInOut.SelectedItem as Win_dvl_InOutWare_Mold_U_CodeView;
            if (winInOutWare != null)
            {
                this.DataContext = winInOutWare;
            }
        }

        //실질 삭제
        private bool dgdDelete(string strInOutID)
        {
            bool flag = false;

            try
            {
                Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                sqlParameter.Add("InOutID", strInOutID);

                string[] result = DataStore.Instance.ExecuteProcedure("xp_dvlMold_dMoldInOut", sqlParameter, true);

                if (result[0].Equals("success"))
                {
                    //MessageBox.Show("성공 *^^*");
                    flag = true;

                }
                //else
                //{
                //    MessageBox.Show("실패 ㅠㅠ");
                //}
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

        //실질 추가
        private bool dgdInsert()
        {
            bool flag = true;
            List<Procedure> Prolist = new List<Procedure>();
            List<Dictionary<string, object>> ListParameter = new List<Dictionary<string, object>>();

            if (CheckData())
            {
                try
                {
                    Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                    sqlParameter.Add("InOutID", "");
                    sqlParameter.Add("MoldID", txtMoldNo.Tag.ToString());
                    sqlParameter.Add("InOutGbn" , cboGubun.SelectedValue == null ? "" : cboGubun.SelectedValue.ToString());
                    sqlParameter.Add("InOutQty", txtInOutQty.Text);
                    sqlParameter.Add("InOutDate", dtpDay.SelectedDate.Value.ToString("yyyyMMdd"));
                    sqlParameter.Add("InOutPlace"   , cboInOutPlace.SelectedValue == null ? "" : cboInOutPlace.SelectedValue.ToString());
                    sqlParameter.Add("Commetns", txtComments.Text);
                    sqlParameter.Add("InOutPerson", txtPerson.Tag.ToString());
                    sqlParameter.Add("CreateUserID", MainWindow.CurrentUser);

                    Procedure pro1 = new Procedure();
                    pro1.Name = "xp_dvlMold_iMoldInOut";
                    pro1.OutputUseYN = "Y";
                    pro1.OutputName = "InOutID";
                    pro1.OutputLength = "10";

                    Prolist.Add(pro1);
                    ListParameter.Add(sqlParameter);

                    List<KeyValue> list_Result = new List<KeyValue>();
                    list_Result = DataStore.Instance.ExecuteAllProcedureOutputGetCS(Prolist, ListParameter);
                    string sGetDefectID = string.Empty;

                    if (list_Result[0].key.ToLower() == "success")
                    {
                        list_Result.RemoveAt(0);
                        for (int i = 0; i < list_Result.Count; i++)
                        {
                            KeyValue kv = list_Result[i];
                            if (kv.key == "InOutID")
                            {
                                sGetDefectID = kv.value;
                                dicCompare.Add("InOutID",sGetDefectID);
                            }
                        }
                    }
                    else
                    {
                        MessageBox.Show("[저장실패]\r\n" + list_Result[0].value.ToString());
                        flag = false;
                        //return false;
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
            }
            else { flag = false; }

            return flag;
        }

        //실질 수정
        private bool dgdUpdate()
        {
            bool flag = true;
            List<Procedure> Prolist = new List<Procedure>();
            List<Dictionary<string, object>> ListParameter = new List<Dictionary<string, object>>();

            try
            {
                Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                sqlParameter.Add("InOutID", txtInOutID.Text);
                sqlParameter.Add("MoldID", txtMoldNo.Tag.ToString());
                sqlParameter.Add("InOutGbn" , cboGubun.SelectedValue == null ? "" : cboGubun.SelectedValue.ToString());
                sqlParameter.Add("InOutQty", txtInOutQty.Text);
                sqlParameter.Add("InOutDate", dtpDay.SelectedDate.Value.ToString("yyyyMMdd"));
                sqlParameter.Add("Commetns", txtComments.Text);
                sqlParameter.Add("InOutPlace" , cboInOutPlace.SelectedValue == null ? "" : cboInOutPlace.SelectedValue.ToString());
                sqlParameter.Add("InOutPerson", txtPerson.Tag.ToString());
                sqlParameter.Add("LastUserID", MainWindow.CurrentUser);

                dicCompare.Add("InOutID", txtInOutID.Text);

                Procedure pro1 = new Procedure();
                pro1.Name = "xp_dvlMold_uMoldInOut";
                pro1.OutputUseYN = "N";
                pro1.OutputName = "InOutID";
                pro1.OutputLength = "10";

                Prolist.Add(pro1);
                ListParameter.Add(sqlParameter);

                string[] Confirm = new string[2];
                Confirm = DataStore.Instance.ExecuteAllProcedureOutputNew(Prolist, ListParameter);

                if (Confirm[0] != "success")
                {
                    MessageBox.Show("[저장실패]\r\n" + Confirm[1].ToString());
                    flag = false;
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

        //추가 하기전 데이터 체크
        private bool CheckData()
        {
            bool flag = true;

            if (txtMoldNo.Tag == null)
            {
                MessageBox.Show("금형LotNo 선택이 잘못되었습니다. enter키 또는 금형LotNo 옆의 버튼을 이용하여 다시 입력해주세요");
                flag = false;
                return flag;
            }

            if (txtPerson.Tag == null)
            {
                MessageBox.Show("담당자 선택이 잘못되었습니다. enter키 또는 담당자 옆의 버튼을 이용하여 다시 입력해주세요");
                flag = false;
                return flag;
            }

            if (txtInOutQty.Text.Equals(""))
            {
                MessageBox.Show("수량이 입력되지 않았습니다. 수량을 입력해주세요");
                flag = false;
            }

            if (cboGubun.SelectedValue == null)
            {
                MessageBox.Show("구분이 선택되지 않았습니다. 선택해주세요");
                flag = false;
                return flag;
            }

            if (cboInOutPlace.SelectedValue == null)
            {
                MessageBox.Show("보관장소가 선택되지 않았습니다. 선택해주세요");
                flag = false;
                return flag;
            }

            return flag;
        }


        #endregion

        
    }
}

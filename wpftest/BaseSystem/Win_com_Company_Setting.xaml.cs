using System;
using System.Collections.Generic;
using System.Data;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using WPF.MDI;

namespace WizMes_WellMade
{
    /// <summary>
    /// Win_sys_Company_Setting_Q.xaml에 대한 상호 작용 논리
    /// </summary>
    public partial class Win_com_Company_Setting : UserControl
    {
        string stDate = string.Empty;
        string stTime = string.Empty;

        public Win_com_Company_Setting()
        {
            InitializeComponent();
        }

        private void UserControl_Loaded(object sender, RoutedEventArgs e)
        {
            stDate = DateTime.Now.ToString("yyyyMMdd");
            stTime = DateTime.Now.ToString("HHmm");

            DataStore.Instance.InsertLogByFormS(this.GetType().Name, stDate, stTime, "S");

            FillGrid();

            if (dgdMain.Items.Count > 0)
            {
                dgdMain.SelectedIndex = 0;
            }

            rbnDoro.IsChecked = true;
        }

        /// <summary>
        /// 취소, 저장 후
        /// </summary>
        private void CanBtnControl()
        {
            btnSave.IsEnabled = false;
            btnCancel.IsEnabled = false;
            dgdMain.IsEnabled = true;
            grdInput.IsHitTestVisible = false;
            lblMsg.Visibility = Visibility.Hidden;
        }

        /// <summary>
        /// 추가, 수정 클릭 시
        /// </summary>
        private void CantBtnControl()
        {
            btnSave.IsEnabled = true;
            btnCancel.IsEnabled = true;
            dgdMain.IsEnabled = false;
            grdInput.IsHitTestVisible = true;
            lblMsg.Visibility = Visibility.Visible;
        }

        //저장
        private void btnSave_Click(object sender, RoutedEventArgs e)
        {
            MessageBoxResult msgresult = MessageBox.Show("변경된 내용을 저장 하시겠습니까?", "수정 확인", MessageBoxButton.YesNo);
            if (msgresult == MessageBoxResult.Yes)
            {
                if (EssentialWriteText() == false)
                {
                    return;
                }

                if (SaveData())
                {
                    CanBtnControl();

                    if (dgdMain.Items.Count > 0)
                    {
                        dgdMain.Items.Clear();
                    }

                    FillGrid();

                    if (dgdMain.Items.Count > 0)
                    {
                        dgdMain.SelectedIndex = 0;
                    }

                    rbnDoro.IsChecked = true;
                }
            }
        }

        //취소
        private void btnCancel_Click(object sender, RoutedEventArgs e)
        {
            CanBtnControl();
        }

        //닫기
        private void btnClose_Click(object sender, RoutedEventArgs e)
        {
            DataStore.Instance.InsertLogByFormS(this.GetType().Name, stDate, stTime, "E");

            try
            {
                int i = 0;
                foreach (MenuViewModel mvm in MainWindow.mMenulist)
                {
                    if (mvm.subProgramID.ToString().Contains("MDI"))
                    {
                        if (this.ToString().Equals((mvm.subProgramID as MdiChild).Content.ToString()))
                        {
                            (MainWindow.mMenulist[i].subProgramID as MdiChild).Close();
                            break;
                        }
                    }
                    i++;
                }
            }
            catch (Exception ee)
            {
                MessageBox.Show("오류지점 - 닫기버튼 : " + ee.ToString());
            }
        }

        //실조회
        private void FillGrid()
        {
            try
            {
                Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                sqlParameter.Add("nChkCompany", 0);
                sqlParameter.Add("sCompanyID", "");
                sqlParameter.Add("sKCompany", "");

                DataSet ds = DataStore.Instance.ProcedureToDataSet_LogWrite("xp_Info_GetCompanyInfo", sqlParameter, true, "R");

                if (ds != null && ds.Tables.Count > 0)
                {
                    DataTable dt = ds.Tables[0];

                    if (dt.Rows.Count > 0)
                    {
                        DataRowCollection drc = dt.Rows;

                        for (int i = 0; i < drc.Count; i++)
                        {
                            DataRow dr = drc[i];

                            var WinCompanySetting = new Win_sys_Company_Setting_Q_CodeView()
                            {
                                Num = (i + 1),
                                Address1 = dr["Address1"].ToString(),
                                Address2 = dr["Address2"].ToString(),
                                AddressAssist = dr["AddressAssist"].ToString(),
                                AddressJiBun1 = dr["AddressJiBun1"].ToString(),
                                AddressJiBun2 = dr["AddressJiBun2"].ToString(),
                                Bank1 = dr["Bank1"].ToString(),
                                Bank2 = dr["Bank2"].ToString(),
                                Bank3 = dr["Bank3"].ToString(),
                                Category = dr["Category"].ToString(),
                                Chief = dr["Chief"].ToString(),
                                CMPN_CD = dr["CMPN_CD"].ToString(),
                                CompanyID = dr["CompanyID"].ToString(),
                                CompanyNo = dr["CompanyNo"].ToString(),
                                Condition = dr["Condition"].ToString(),
                                CreateDate = dr["CreateDate"].ToString(),
                                CreateUserID = dr["CreateUserID"].ToString(),
                                ECompany = dr["ECompany"].ToString(),
                                EMail = dr["EMail"].ToString(),
                                FaxNO = dr["FaxNO"].ToString(),
                                FTPAuthCode1 = dr["FTPAuthCode1"].ToString(),
                                FTPAuthCode2 = dr["FTPAuthCode2"].ToString(),
                                FTPID1 = dr["FTPID1"].ToString(),
                                FTPID2 = dr["FTPID2"].ToString(),
                                FTPPage = dr["FTPPage"].ToString(),
                                FTPPass1 = dr["FTPPass1"].ToString(),
                                FTPPass2 = dr["FTPPass2"].ToString(),
                                FTPPortFrom = dr["FTPPortFrom"].ToString(),
                                FTPPortTo = dr["FTPPortTo"].ToString(),
                                HomePage = dr["HomePage"].ToString(),
                                GunMoolMngNo = dr["GunMoolMngNo"].ToString(),
                                KCompany = dr["KCompany"].ToString(),
                                OldNNewClss = dr["OldNNewClss"].ToString(),
                                Phone1 = dr["Phone1"].ToString(),
                                Phone2 = dr["Phone2"].ToString(),
                                RegistID = dr["RegistID"].ToString(),
                                RpYN = dr["RpYN"].ToString(),
                                ShortCompany = dr["ShortCompany"].ToString(),
                                SMSPortFrom1 = dr["SMSPortFrom1"].ToString(),
                                SMSAuthCode1 = dr["SMSAuthCode1"].ToString(),
                                SMSPortFrom2 = dr["SMSPortFrom2"].ToString(),
                                SMSAuthCode2 = dr["SMSAuthCode2"].ToString(),
                                SMSID1 = dr["SMSID1"].ToString(),
                                SMSID2 = dr["SMSID2"].ToString(),
                                SMSPASS1 = dr["SMSPASS1"].ToString(),
                                SMSPASS2 = dr["SMSPASS2"].ToString(),
                                SMSPortTo1 = dr["SMSPortTo1"].ToString(),
                                SMSPortTo2 = dr["SMSPortTo2"].ToString(),
                                SMSURL1 = dr["SMSURL1"].ToString(),
                                SMSURL2 = dr["SMSURL2"].ToString(),
                                UpdateDate = dr["UpdateDate"].ToString(),
                                UpdateUserID = dr["UpdateUserID"].ToString(),
                                UseYn = dr["UseYn"].ToString(),
                                WebAuthCode1 = dr["WebAuthCode1"].ToString(),
                                WebAuthCode2 = dr["WebAuthCode2"].ToString(),
                                WebID1 = dr["WebID1"].ToString(),
                                WebID2 = dr["WebID2"].ToString(),
                                WebPass1 = dr["WebPass1"].ToString(),
                                WebPass2 = dr["WebPass2"].ToString(),
                                WebPortFrom = dr["WebPortFrom"].ToString(),
                                WebPortTo = dr["WebPortTo"].ToString(),
                                ZipCode = dr["ZipCode"].ToString(),

                            };

                            if (WinCompanySetting.ZipCode.Length > 0)
                            {
                                if (WinCompanySetting.ZipCode.Length == 6)
                                {
                                    WinCompanySetting.ZipCode = WinCompanySetting.ZipCode.Substring(0, 3) + "-"
                                                                 + WinCompanySetting.ZipCode.Substring(3, 3);
                                }
                            }

                            dgdMain.Items.Add(WinCompanySetting);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("오류지점 - 저장 프로시저 : " + ex.ToString());
            }
            finally
            {
                DataStore.Instance.CloseConnection();
            }
        }

        private void dgdMain_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            CantBtnControl();

            txtCompanyKName.Focus();
        }

        //연결되어 있는 버튼이 없음.
        private void btnZipCode_Click(object sender, RoutedEventArgs e)
        {
            PopUp.Win_Zip_Address ZipPopUp = new PopUp.Win_Zip_Address();
            ZipPopUp.ShowDialog();

            if (ZipPopUp.DialogResult == true)
            {
                if (ZipPopUp.strGubun.Equals("0"))
                {
                    txtAddress1.Text = ZipPopUp.Juso;
                    txtAddress2.Text = ZipPopUp.Detail1;
                    txtAddressAssist.Text = ZipPopUp.Detail2;
                    txtZipCode.Text = ZipPopUp.ZipCode;
                    txtGunMoolMngNo.Text = ZipPopUp.GunMoolMngNo;
                }
                else if (ZipPopUp.strGubun.Equals("1"))
                {
                    txtAddressJiBun1.Text = ZipPopUp.Juso;
                    txtZipCode.Text = ZipPopUp.ZipCode;
                }
            }
        }

        private void dgdMain_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            var WinCompany = dgdMain.SelectedItem as Win_sys_Company_Setting_Q_CodeView;

            if (WinCompany != null)
            {
                this.DataContext = WinCompany;
                if (WinCompany.RpYN.Equals("Y"))
                {
                    rbnRpY.IsChecked = true;
                }
                else
                {
                    rbnRpN.IsChecked = true;
                }

                if (WinCompany.UseYn.Equals("Y"))
                {
                    rbnUseY.IsChecked = true;
                }
                else
                {
                    rbnUseN.IsChecked = true;
                }
            }
        }

        private bool SaveData()
        {
            bool flag = false;

            try
            {
                Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                sqlParameter.Add("CompanyID", txtCompanyCode.Text);
                sqlParameter.Add("ECompany", txtECompany.Text);
                sqlParameter.Add("KCompany", txtCompanyKName.Text);
                sqlParameter.Add("ShortCompany", txtShortCompany.Text);
                sqlParameter.Add("CompanyNo", txtCompanyNO.Text);
                sqlParameter.Add("RegistID", txtRegistID.Text);
                sqlParameter.Add("Chief", txtChief.Text);
                sqlParameter.Add("Condition", txtCondition.Text);
                sqlParameter.Add("Category", txtCategory.Text);
                sqlParameter.Add("ZipCode", txtZipCode.Text);
                sqlParameter.Add("OldNNewClss", rbnDoro.IsChecked == true ? 0 : 1);                            //도로명 / 지번 나누는 기준인듯 한데.. plusfinder가....
                sqlParameter.Add("GunMoolMngNo", txtGunMoolMngNo.Text);       // 이것도 plusfinder가 값을 넘겨주는 듯 한디...
                sqlParameter.Add("Address1", txtAddress1.Text);
                sqlParameter.Add("Address2", txtAddress2.Text);
                sqlParameter.Add("AddressAssist", txtAddressAssist.Text);
                sqlParameter.Add("AddressJiBun1", txtAddressJiBun1.Text);
                sqlParameter.Add("AddressJiBun2", txtAddressJiBun2.Text);
                sqlParameter.Add("Phone1", txtPhone1.Text);
                sqlParameter.Add("Phone2", txtPhone2.Text);
                sqlParameter.Add("FaxNo", txtFaxNo.Text);
                sqlParameter.Add("EMail", txtEMail.Text);
                sqlParameter.Add("HomePage", txtHomePage.Text);

                //sqlParameter.Add("WebPortFrom", txtWebPortFrom.Text);
                //sqlParameter.Add("WebPortTo", txtWebPortTo.Text);
                //sqlParameter.Add("WebID1", txtWebID1.Text);
                //sqlParameter.Add("WebPass1", txtWebPass1.Text);
                //sqlParameter.Add("WebAuthCode1", txtWebAuthCode1.Text);
                //sqlParameter.Add("WebID2", txtWebID2.Text);
                //sqlParameter.Add("WebPass2", txtWebPass2.Text);
                //sqlParameter.Add("WebAuthCode2", txtWebAuthCode2.Text);
                //sqlParameter.Add("FTPPage", txtFTPPage.Text);
                //sqlParameter.Add("FTPPortFrom", txtFTPPortFrom.Text);
                //sqlParameter.Add("FTPPortTo", txtFTPPortTo.Text);
                //sqlParameter.Add("FTPID1", txtFTPID1.Text);
                //sqlParameter.Add("FTPPass1", txtFTPPass1.Text);
                //sqlParameter.Add("FTPAuthCode1", txtFTPAuthCode1.Text);
                //sqlParameter.Add("FTPID2", txtFTPID2.Text);
                //sqlParameter.Add("FTPPass2", txtFTPPass2.Text);
                //sqlParameter.Add("FTPAuthCode2", txtFTPAuthCode2.Text);
                //sqlParameter.Add("SMSURL1", txtSMSURL1.Text);
                //sqlParameter.Add("SMSPortFrom1", txtSMSPortFrom1.Text);
                //sqlParameter.Add("SMSPortTo1", txtSMSPortTo1.Text);
                //sqlParameter.Add("SMSID1", txtSMSID1.Text);
                //sqlParameter.Add("SMSPASS1", txtSMSPASS1.Text);
                //sqlParameter.Add("SMSAuthCode1", txtSMSAuthCode1.Text);
                //sqlParameter.Add("SMSURL2", txtSMSURL2.Text);
                //sqlParameter.Add("SMSPortFrom2", txtSMSPortFrom2.Text);
                //sqlParameter.Add("SMSPortTo2", txtSMSPortTo2.Text);
                //sqlParameter.Add("SMSID2", txtSMSID2.Text);
                //sqlParameter.Add("SMSPASS2", txtSMSPASS2.Text);
                //sqlParameter.Add("SMSAuthCode2", txtSMSAuthCode2.Text);

                sqlParameter.Add("WebPortFrom", "");
                sqlParameter.Add("WebPortTo", "");
                sqlParameter.Add("WebID1", "");
                sqlParameter.Add("WebPass1", "");
                sqlParameter.Add("WebAuthCode1", "");
                sqlParameter.Add("WebID2", "");
                sqlParameter.Add("WebPass2", "");
                sqlParameter.Add("WebAuthCode2", "");
                sqlParameter.Add("FTPPage", "");
                sqlParameter.Add("FTPPortFrom", "");
                sqlParameter.Add("FTPPortTo", "");
                sqlParameter.Add("FTPID1", "");
                sqlParameter.Add("FTPPass1", "");
                sqlParameter.Add("FTPAuthCode1", "");
                sqlParameter.Add("FTPID2", "");
                sqlParameter.Add("FTPPass2", "");
                sqlParameter.Add("FTPAuthCode2", "");
                sqlParameter.Add("SMSURL1", "");
                sqlParameter.Add("SMSPortFrom1", "");
                sqlParameter.Add("SMSPortTo1", "");
                sqlParameter.Add("SMSID1", "");
                sqlParameter.Add("SMSPASS1", "");
                sqlParameter.Add("SMSAuthCode1", "");
                sqlParameter.Add("SMSURL2", "");
                sqlParameter.Add("SMSPortFrom2", "");
                sqlParameter.Add("SMSPortTo2", "");
                sqlParameter.Add("SMSID2", "");
                sqlParameter.Add("SMSPASS2", "");
                sqlParameter.Add("SMSAuthCode2", "");

                sqlParameter.Add("Bank1", txtBank1.Text);
                sqlParameter.Add("Bank2", txtBank2.Text);
                sqlParameter.Add("Bank3", txtBank3.Text);

                sqlParameter.Add("CMPN_CD", txtCMPN_CD.Text);
                //sqlParameter.Add("BuyCustomID", ""); //자사 정보 설정에서는 해당 데이터를 수정해서는 안됨
                //sqlParameter.Add("SaleCustomID", ""); //자사 정보 설정에서는 해당 데이터를 수정해서는 안됨       // 이게 먼지 아직 정확하지 않음....ㅠ        

                //sqlParameter.Add("RPYN", rbnUseY.IsChecked == true ? "Y" : "N"); 
                //sqlParameter.Add("UseYN", rbnUseY.IsChecked == true ? "Y" : "N");

                sqlParameter.Add("RPYN", rbnRpY.IsChecked == true ? "Y" : "N");
                sqlParameter.Add("UseYN", rbnUseY.IsChecked == true ? "Y" : "N");

                string[] result = DataStore.Instance.ExecuteProcedure("xp_Info_UpdateCompanyInfo", sqlParameter, false);
                if (!result[0].Equals("success"))
                {
                    MessageBox.Show("오류지점 - 저장 프로시저 ");
                    //return;
                    flag = false;
                }
                else
                {
                    flag = true;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("오류지점 - 저장 프로시저 : " + ex.ToString());
            }
            finally
            {
                DataStore.Instance.CloseConnection();
            }

            return flag;
        }

        // 필수 입력사항들을 넣고 저장버튼을 누르는거야 지금???

        //private bool EssentialWriteText()
        //{
        //    if (txtCompanyCode.Text == string.Empty)
        //    {
        //        MessageBox.Show("기업 코드정보 란을 빈칸으로 둘 수 없습니다.");
        //        return false;
        //    }
        //    else if (txtCompanyKName.Text == string.Empty)
        //    {
        //        MessageBox.Show("한글 상호명칭 란을 빈칸으로 둘 수 없습니다.");
        //        return false;
        //    }
        //    else if (txtCompanyNO.Text == string.Empty)
        //    {
        //        MessageBox.Show("사업자 번호 란을 빈칸으로 둘 수 없습니다.");
        //        return false;
        //    }
        //    else if (txtAddress1.Text == string.Empty)
        //    {
        //        MessageBox.Show("도로명 주소 란을 빈칸으로 둘 수 없습니다.");
        //        return false;
        //    }
        //    else if (txtAddressJiBun1.Text == string.Empty)
        //    {
        //        MessageBox.Show("지번주소 상단 란을 빈칸으로 둘 수 없습니다.");
        //        return false;
        //    }

        //    return true;
        //}
        private bool EssentialWriteText()
        {
            if (txtCompanyCode.Text == string.Empty)
            {
                MessageBox.Show("기업 코드정보 란을 빈칸으로 둘 수 없습니다.");
                return false;
            }
            else if (txtCompanyKName.Text == string.Empty)
            {
                MessageBox.Show("한글 상호명칭 란을 빈칸으로 둘 수 없습니다.");
                return false;
            }
            else if (txtCompanyNO.Text == string.Empty)
            {
                MessageBox.Show("사업자 번호 란을 빈칸으로 둘 수 없습니다.");
                return false;
            }
            else if (txtAddress1.Text == string.Empty)
            {
                MessageBox.Show("도로명 주소 란을 빈칸으로 둘 수 없습니다.");
                return false;
            }
            else if (txtAddressJiBun1.Text == string.Empty)
            {
                MessageBox.Show("지번주소 상단 란을 빈칸으로 둘 수 없습니다.");
                return false;
            }
            else if (txtChief.Text == string.Empty)
            {
                MessageBox.Show("대표자 란을 빈칸으로 둘 수 없습니다.");
            }
            else if (txtFaxNo.Text == string.Empty)
            {
                MessageBox.Show("팩스번호를 빈칸으로 둘 수 없습니다.");
                return false;
            }
            else if (txtCondition.Text == string.Empty)
            {
                MessageBox.Show("업태 란을 빈칸으로 둘 수 없습니다.");
                return false;
            }
            else if (txtCategory.Text == string.Empty)
            {
                MessageBox.Show("종목 란을 빈칸으로 둘 수 없습니다.");
                return false;
            }
            else if (txtZipCode.Text == string.Empty)
            {
                MessageBox.Show("우편번호 란을 빈칸으로 둘 수 없습니다.");
                return false;
            }
            else if (txtAddress1.Text == string.Empty)
            {
                MessageBox.Show("주소(지번) 란을 빈칸으로 둘 수 없습니다.");
                return false;
            }
            else if (txtAddress2.Text == string.Empty)
            {
                MessageBox.Show("상세주소(지번) 란을 빈칸으로 둘 수 없습니다.");
                return false;
            }
            return true;
        }

        // 우편번호 찾기
        private void txtZipCode_KeyDown(object sender, KeyEventArgs e)
        {
            try
            {
                if (e.Key == Key.Enter)
                {
                    PopUp.Win_Zip_Address ZipPopUp = new PopUp.Win_Zip_Address();
                    ZipPopUp.ShowDialog();

                    if (ZipPopUp.DialogResult == true)
                    {
                        if (ZipPopUp.strGubun.Equals("0"))
                        {
                            txtAddress1.Text = ZipPopUp.Juso;
                            txtAddress2.Text = ZipPopUp.Detail1;
                            txtAddressAssist.Text = ZipPopUp.Detail2;
                            txtZipCode.Text = ZipPopUp.ZipCode;
                            txtGunMoolMngNo.Text = ZipPopUp.GunMoolMngNo;
                        }
                        else if (ZipPopUp.strGubun.Equals("1"))
                        {
                            txtAddressJiBun1.Text = ZipPopUp.Juso;
                            txtZipCode.Text = ZipPopUp.ZipCode;
                        }
                    }
                }
            }
            catch (Exception ee)
            {
                MessageBox.Show("오류지점 - 우편번호 찾기 : " + ee.ToString());
            }
        }

        private void btnPfZipCode_Click(object sender, RoutedEventArgs e)
        {
            PopUp.Win_Zip_Address ZipPopUp = new PopUp.Win_Zip_Address();
            ZipPopUp.ShowDialog();

            if (ZipPopUp.DialogResult == true)
            {
                if (ZipPopUp.strGubun.Equals("0"))
                {
                    txtAddress1.Text = ZipPopUp.Juso;
                    txtAddress2.Text = ZipPopUp.Detail1;
                    txtAddressAssist.Text = ZipPopUp.Detail2;
                    txtZipCode.Text = ZipPopUp.ZipCode;
                    txtGunMoolMngNo.Text = ZipPopUp.GunMoolMngNo;
                }
                else if (ZipPopUp.strGubun.Equals("1"))
                {
                    txtAddressJiBun1.Text = ZipPopUp.Juso;
                    txtZipCode.Text = ZipPopUp.ZipCode;
                }
            }
        }


        #region 텍스트박스 엔터 → 다음 텍스트 박스 이동

        private void txtCompanyKName_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                txtCMPN_CD.Focus();
            }
        }

        private void txtCMPN_CD_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                txtECompany.Focus();
            }
        }

        private void txtECompany_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                txtShortCompany.Focus();
            }
        }

        private void txtShortCompany_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                txtChief.Focus();
            }
        }

        private void txtChief_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                txtCompanyNO.Focus();
            }
        }

        private void txtCompanyNO_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                txtRegistID.Focus();
            }
        }

        private void txtRegistID_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                txtCondition.Focus();
            }
        }

        private void txtCondition_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                txtCategory.Focus();
            }
        }

        private void txtCategory_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                txtZipCode.Focus();
            }
        }

        private void txtAddressAssist_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                txtPhone1.Focus();
            }
        }

        private void txtAddressJiBun2_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                txtPhone1.Focus();
            }
        }

        private void txtPhone1_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                txtPhone2.Focus();
            }
        }

        private void txtPhone2_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                txtFaxNo.Focus();
            }
        }

        private void txtFaxNo_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                txtEMail.Focus();
            }
        }

        private void txtEMail_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                txtHomePage.Focus();
            }
        }

        private void txtHomePage_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                txtBank1.Focus();
            }
        }

        private void txtBank1_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                txtBank2.Focus();
            }
        }

        private void txtBank2_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                txtBank3.Focus();
            }
        }

        private void txtBank3_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                btnSave.Focus();
            }
        }

        #endregion // 텍스트박스 엔터 → 다음 텍스트 박스 이동
    }

    class Win_sys_Company_Setting_Q_CodeView : BaseView
    {
        public int Num { get; set; }

        public string CompanyID { get; set; }
        public string KCompany { get; set; }
        public string ECompany { get; set; }
        public string ShortCompany { get; set; }
        public string Chief { get; set; }

        public string CompanyNo { get; set; }
        public string RegistID { get; set; }
        public string Condition { get; set; }
        public string Category { get; set; }
        public string AddressJiBun1 { get; set; }

        public string AddressJiBun2 { get; set; }
        public string OldNNewClss { get; set; }
        public string GunMoolMngNo { get; set; }
        public string Address1 { get; set; }
        public string Address2 { get; set; }

        public string AddressAssist { get; set; }
        public string ZipCode { get; set; }
        public string Phone1 { get; set; }
        public string Phone2 { get; set; }
        public string FaxNO { get; set; }

        public string EMail { get; set; }
        public string HomePage { get; set; }
        public string Bank1 { get; set; }
        public string Bank2 { get; set; }
        public string Bank3 { get; set; }

        public string RpYN { get; set; }
        public string UseYn { get; set; }
        public string CreateDate { get; set; }
        public string CreateUserID { get; set; }
        public string UpdateDate { get; set; }

        public string UpdateUserID { get; set; }
        public string WebPortFrom { get; set; }
        public string WebPortTo { get; set; }
        public string WebID1 { get; set; }
        public string WebPass1 { get; set; }

        public string WebAuthCode1 { get; set; }
        public string WebID2 { get; set; }
        public string WebPass2 { get; set; }
        public string WebAuthCode2 { get; set; }
        public string FTPPage { get; set; }

        public string FTPPortFrom { get; set; }
        public string FTPPortTo { get; set; }
        public string FTPID1 { get; set; }
        public string FTPPass1 { get; set; }
        public string FTPAuthCode1 { get; set; }

        public string FTPID2 { get; set; }
        public string FTPPass2 { get; set; }
        public string FTPAuthCode2 { get; set; }
        public string SMSURL1 { get; set; }
        public string SMSPortFrom1 { get; set; }

        public string SMSPortTo1 { get; set; }
        public string SMSID1 { get; set; }
        public string SMSPASS1 { get; set; }
        public string SMSAuthCode1 { get; set; }
        public string SMSURL2 { get; set; }

        public string SMSPortFrom2 { get; set; }
        public string SMSPortTo2 { get; set; }
        public string SMSID2 { get; set; }
        public string SMSPASS2 { get; set; }
        public string SMSAuthCode2 { get; set; }

        public string CMPN_CD { get; set; }
    }
}

using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Data;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;

namespace WizMes_WellMade.PopUp
{
    /// <summary>
    /// Win_Zip_Address.xaml에 대한 상호 작용 논리
    /// </summary>
    public partial class Win_Zip_Address : Window
    {
        string strSido = string.Empty;
        string SiDTable = string.Empty;
        public string strGubun = "0";
        string strGubun2 = string.Empty;
        string strSiGunGu = string.Empty;

        /// <summary>
        /// 우편번호
        /// </summary>
        public string ZipCode = string.Empty;

        /// <summary>
        /// 주소
        /// </summary>
        public string Juso = string.Empty;

        /// <summary>
        /// 보조번호(도로명 건물번호, 지번 번지)
        /// </summary>
        public string Detail1 = string.Empty;

        /// <summary>
        /// 보조주소
        /// </summary>
        public string Detail2 = string.Empty;

        /// <summary>
        /// 건물번호
        /// </summary>
        public string GunMoolMngNo = string.Empty;

        public Win_Zip_Address()
        {
            InitializeComponent();
        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            SetComboBox();
        }

        //콤보박스 셋팅!
        private void SetComboBox()
        {
            ObservableCollection<CodeView> ovcSiDO = ComboBoxUtil.Instance.GetSido("0");
            this.cboSiDo.ItemsSource = ovcSiDO;
            this.cboSiDo.DisplayMemberPath = "code_name";
            this.cboSiDo.SelectedValuePath = "code_id";
            this.cboSiDo.SelectedIndex = 0;
        }

        //
        private void cboSiDo_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (cboSiDo.SelectedIndex > -1)
            {
                strSido = cboSiDo.SelectedValue.ToString();

                ObservableCollection<CodeView> ovcSiGunGu = ComboBoxUtil.Instance.GetSiGunGu(strSido, "0"); this.cboSiGunGu.ItemsSource = ovcSiGunGu;
                this.cboSiGunGu.DisplayMemberPath = "code_name";
                this.cboSiGunGu.SelectedValuePath = "code_id";
                this.cboSiGunGu.SelectedIndex = 0;
            }
            else
            {
                strSido = "";
            }
        }

        //
        private void cboSiGunGu_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (cboSiGunGu.SelectedIndex > -1)
            {
                strSiGunGu = cboSiGunGu.SelectedValue.ToString();
            }
            else
            {
                strSiGunGu = "";
            }
        }

        //검색 클릭
        private void btnSearch_Click(object sender, RoutedEventArgs e)
        {
            switch (strSido)
            {
                case "11":  //서울
                    SiDTable = "ZipZone_Seoul";
                    break;
                case "26":  //부산
                    SiDTable = "ZipZone_BuSan";
                    break;
                case "27":  //대구
                    SiDTable = "ZipZone_DaeGu";
                    break;
                case "28":  //인천
                    SiDTable = "ZipZone_InCheon";
                    break;
                case "29":  //광주
                    SiDTable = "ZipZone_GwangJu";
                    break;
                case "30":  //대전
                    SiDTable = "ZipZone_DaeJeon";
                    break;
                case "31":  //울산
                    SiDTable = "ZipZone_UlSan";
                    break;
                case "36":  //세종특별자치시
                    SiDTable = "ZipZone_SeJong";
                    break;
                case "41":  //경기도
                    SiDTable = "ZipZone_GyeongGi";
                    break;
                case "42":  //강원도
                    SiDTable = "ZipZone_GangWon";
                    break;
                case "43":  //충청북도
                    SiDTable = "ZipZone_ChungBuk";
                    break;
                case "44":  //충청남도
                    SiDTable = "ZipZone_ChungNam";
                    break;
                case "45":  //전라북도
                    SiDTable = "ZipZone_JeonBuk";
                    break;
                case "46":  //전라남도
                    SiDTable = "ZipZone_JeonNam";
                    break;
                case "47":  //경상북도
                    SiDTable = "ZipZone_GyeongBuk";
                    break;
                case "48":  //경상남도
                    SiDTable = "ZipZone_GyeongNam";
                    break;
                case "50":  //제주특별자치시도
                    SiDTable = "ZipZone_JeJu";
                    break;
            }

            if ((sender as Button).Name.Equals("btnSearch"))
            {
                FillGridDoRo();
            }
            else if ((sender as Button).Name.Equals("btnJBSearch"))
            {
                if (txtVillage.Text.Length > 0)
                {
                    FillGridJiBun();
                }
                else
                {
                    MessageBox.Show("검색할 동(읍/면/리)을 입력해주세요");
                }
            }
        }

        //선택 클릭
        private void btnOK_Click(object sender, RoutedEventArgs e)
        {
            if (strGubun.Equals("0"))
            {
                var Juso = dgdJuso.SelectedItem as ZipAddress;

                if (Juso != null)
                {
                    this.ZipCode = Juso.Zip_Code_CV;
                    this.Juso = Juso.addr;
                    this.Detail1 = Juso.addr2;
                    this.Detail2 = Juso.AssistAddr;
                    this.GunMoolMngNo = Juso.GunMoolMng_No;
                }
            }
            else if (strGubun.Equals("1"))
            {
                var Juso = dgdJiBun.SelectedItem as ZipAddress;

                if (Juso != null)
                {
                    this.ZipCode = Juso.ZipCode;
                    this.Juso = Juso.ContainJuso;
                    this.Detail1 = Juso.Detail1;
                }
            }

            DialogResult = true;
        }

        //탭의 상단 클릭
        private void tbConJuso_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            TabItem NowTI = ((TabControl)sender).SelectedItem as TabItem;

            if (NowTI.Header.ToString().Equals("도로명"))
            {
                strGubun = "0";
            }
            else if (NowTI.Header.ToString().Equals("지번"))
            {
                strGubun = "1";
            }
        }
        /// <summary>
        /// 라디오 버튼 선택 도로명+건물번호
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void rbnJuso1_Checked(object sender, RoutedEventArgs e)
        {
            strGubun2 = "0";
        }

        /// <summary>
        /// 라디오 버튼 선택 동읍리명
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void rbnJuso2_Checked(object sender, RoutedEventArgs e)
        {
            strGubun2 = "1";
        }

        /// <summary>
        /// 라디오 버튼 선택 건물명
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void rbnJuso3_Checked(object sender, RoutedEventArgs e)
        {
            strGubun2 = "2";
        }

        //도로명 검색
        private void FillGridDoRo()
        {
            if (dgdJuso.Items.Count > 0)
            {
                dgdJuso.Items.Clear();
            }

            Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
            sqlParameter.Clear();
            sqlParameter.Add("sGubun", strGubun);
            sqlParameter.Add("sGubun2", strGubun2);
            sqlParameter.Add("sSido", strSido);
            sqlParameter.Add("sSidoTable", SiDTable);
            sqlParameter.Add("sSiGunGu", strSiGunGu);
            sqlParameter.Add("sSearchWord", txtName.Text);

            DataSet ds = DataStore.Zip_Instance.ProcedureToDataSetByZip("xp_ZipCode_sAddress", sqlParameter, false);

            if (ds != null && ds.Tables.Count > 0)
            {
                DataTable dt = ds.Tables[0];

                if (dt.Rows.Count > 0)
                {
                    DataRowCollection drc = dt.Rows;
                    int i = 0;
                    //dgdJuso.ItemsSource = dt.DefaultView;

                    foreach (DataRow dr in drc)
                    {
                        i++;
                        var ZipJuso = new ZipAddress()
                        {
                            Num = i,
                            GunMoolMng_No = dr["GunMoolMng_No"].ToString(),
                            Sido_Name = dr["Sido_Name"].ToString(),
                            SiGunGu_Name = dr["SiGunGu_Name"].ToString(),
                            EupMyunDong_Name = dr["EupMyunDong_Name"].ToString(),
                            BubLi_Name = dr["BubLi_Name"].ToString(),
                            JiBunBon_No = dr["JiBunBon_No"].ToString(),
                            JiBunBoo_No = dr["JiBunBoo_No"].ToString(),
                            Doro_Name = dr["Doro_Name"].ToString(),
                            GunMool_Name = dr["GunMool_Name"].ToString(),
                            GunMoolBon_No = dr["GunMoolBon_No"].ToString(),
                            GunMoolBoo_No = dr["GunMoolBoo_No"].ToString(),
                            addr = dr["addr"].ToString(),
                            addr2 = dr["addr2"].ToString(),
                            AssistAddr = dr["AssistAddr"].ToString(),
                            HaengDong_Name = dr["HaengDong_Name"].ToString(),
                            Zip_Code = dr["Zip_Code"].ToString(),
                            Zip_Code_Seq = dr["Zip_Code_Seq"].ToString()
                        };

                        if (ZipJuso.Zip_Code.Length == 5)
                        {
                            ZipJuso.Zip_Code_CV = ZipJuso.Zip_Code.Substring(0, 3) + "-"
                                + ZipJuso.Zip_Code.Substring(ZipJuso.Zip_Code.Length - 2);
                        }

                        dgdJuso.Items.Add(ZipJuso);
                    }
                }
            }
        }

        //지번 검색
        private void FillGridJiBun()
        {
            if (dgdJiBun.Items.Count > 0)
            {
                dgdJiBun.Items.Clear();
            }

            Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
            sqlParameter.Clear();
            sqlParameter.Add("sVillage", txtVillage.Text);
            DataSet ds = DataStore.Zip_Instance.ProcedureToDataSetByZip("xp_Common_sZipCode", sqlParameter, false);

            if (ds != null && ds.Tables.Count > 0)
            {
                DataTable dt = ds.Tables[0];

                if (dt.Rows.Count > 0)
                {
                    DataRowCollection drc = dt.Rows;
                    int i = 0;
                    //dgdJuso.ItemsSource = dt.DefaultView;

                    foreach (DataRow dr in drc)
                    {
                        i++;
                        var ZipJuso = new ZipAddress()
                        {
                            Num = i,
                            ZipCode = dr["ZipCode"].ToString(),
                            City = dr["City"].ToString(),
                            Section = dr["Section"].ToString(),
                            Village = dr["Village"].ToString(),
                            Detail1 = dr["Detail1"].ToString(),
                            Detail2 = dr["Detail2"].ToString()
                        };

                        ZipJuso.ContainJuso = ZipJuso.City + " " + ZipJuso.Section + " " + ZipJuso.Village;

                        dgdJiBun.Items.Add(ZipJuso);
                    }
                }
            }
        }

        //도로명
        private void dgdJuso_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (strGubun.Equals("0"))
            {
                var Juso = dgdJuso.SelectedItem as ZipAddress;

                if (Juso != null)
                {
                    this.DataContext = Juso;
                }
            }
        }

        //지번
        private void dgdJiBun_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (strGubun.Equals("1"))
            {
                var Juso = dgdJiBun.SelectedItem as ZipAddress;

                if (Juso != null)
                {
                    this.DataContext = Juso;
                }
            }
        }
    }

    class ZipAddress : BaseView
    {
        public int Num { get; set; }
        public string GunMoolMng_No { get; set; }
        public string Sido_Name { get; set; }
        public string SiGunGu_Name { get; set; }
        public string EupMyunDong_Name { get; set; }
        public string BubLi_Name { get; set; }
        public string JiBunBon_No { get; set; }
        public string JiBunBoo_No { get; set; }
        public string Doro_Name { get; set; }
        public string GunMool_Name { get; set; }
        public string GunMoolBon_No { get; set; }
        public string GunMoolBoo_No { get; set; }
        public string addr { get; set; }
        public string addr2 { get; set; }
        public string AssistAddr { get; set; }
        public string HaengDong_Name { get; set; }
        public string Zip_Code { get; set; }
        public string Zip_Code_Seq { get; set; }
        public string Zip_Code_CV { get; set; }

        public string ZipCode { get; set; }
        public string City { get; set; }
        public string Section { get; set; }
        public string Village { get; set; }
        public string Detail1 { get; set; }
        public string Detail2 { get; set; }
        public string ContainJuso { get; set; }
    }
}

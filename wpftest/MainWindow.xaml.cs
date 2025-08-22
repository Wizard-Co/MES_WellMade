using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Input;
using WizMes_WellMade.PopUp;
using WPF.MDI;

namespace WizMes_WellMade
{
    /// <summary>
    /// MainWindow.xaml에 대한 상호 작용 논리
    /// </summary>
    public partial class MainWindow : Window
    {
        public static string mainStDate = string.Empty;
        public static string mainStTime = string.Empty;


        public static List<MenuViewModel> mMenulist = new List<MenuViewModel>();
        public static MdiContainer MainMdiContainer = new MdiContainer();
        public static string CurrentUser = string.Empty;
        public static string CurrentName = string.Empty;
        public static string CompanyID = string.Empty;
        public static PlusFinder pf = new PlusFinder();
        public static PlusFinder pf2 = new PlusFinder(); //2021-11-10 자재입고반품에서 2번 연속으로 띄워야되서 하나 더 추가함
        public static List<object> objList = new List<object>();
        public string mfont { get; set; } //
        Lib lib = new Lib();
        public string[] strFavorites = null;

        public static List<MenuViewModel> listFavorites = new List<MenuViewModel>();
        public MenuViewModel currentMenuViewModel = null;
        public static string CurrentPerson = string.Empty;
        public static string CurrentPersonID = string.Empty;
        public static double StdFontSize = 0;

        public int TheFont { get; set; }
        public double TheHeight { get; set; }
        public double TheWidth { get; set; }
       //aa
        // 넘겨줄 임시 데이터 변수(수주 진행 및 마감 -> 수주등록 화면으로)
        public static List<string> tempContent = new List<string>();
        // 넘겨줄 임시 데이터 변수(품목별 불량현황 -> 스크린샷 popup)
        public static List<Image> ScreenCapture = new List<Image>();
        // 넘겨줄 임시 데이터 변수(설비가동률 -> 설비가동률상세조회)
        public static List<string> MCtemp = new List<string>();
        // 넘겨줄 임시 데이터 변수(XbarChart -> 상세분석)
        public static List<string> ChartDeatil = new List<string>();

        // [생산] - 생산계획작성 - 저장과 동시 작업지시 안내를 위한 변수
        public static bool plInputFlag_SavePrint = false;
        public static Dictionary<string, object> plInput = new Dictionary<string, object>();

        // 확장자 이름 → SetImage 메서드에서 오류 발생하여, 혹여나 이미지가 아닌 다른 형식의 파일이 업로드 됬을때를 대비해서, 그냥 이미지으로 마지막 확장자명만 체크를 위한 변수
        public static string[] Extensions = { ".jpg", ".jpeg", ".jpe", ".jfif", ".png" };
        // 이미지 업로드 시 확장자 필터 공용관리
        public static string OFdlg_Filter = "Image files (*.jpg, *.jpeg, *.jpe, *.jfif, *.png) | *.jpg; *.jpeg; *.jpe; *.jfif; *.png;";


        public static string[] OFdlg_Filter_NotAllowed = new string[]
        {
            // 실행 파일
            ".exe", ".msi", ".scr", ".com", 
    
            // 스크립트 파일
            ".bat", ".cmd", ".vbs", ".js", ".jse", ".wsf", ".wsh", ".ps1", ".psm1", ".psd1",
    
            // 시스템/드라이버 파일
            ".sys", ".dll", ".ocx", ".drv", ".cpl",
    
            // 레지스트리 관련
            ".reg",
    
            // 바로가기
            ".lnk", 
    
            // 기타 특수 실행 파일
            ".hta", ".msc", ".jar", ".gadget",
    
            // 매크로 포함 가능한 문서
            ".xlsm", ".xltm", ".xlam", ".pptm", ".potm", ".ppam", ".ppsm", ".sldm",
            ".docm", ".dotm"
        };

        #region 생성자

        public MainWindow()
        {
            InitializeComponent();

            if (Login())
            {
                Style = (Style)FindResource(typeof(Window));
                menuLoad();

                mainStDate = DateTime.Now.ToString("yyyyMMdd");
                mainStTime = DateTime.Now.ToString("HHmm");

                this.Height = SystemParameters.WorkArea.Height;
                this.Width = SystemParameters.WorkArea.Width;

                mdiPanel.Height = SystemParameters.WorkArea.Height;
                mdiPanel.Children.Add(MainMdiContainer);

                uiScaleSlider.MouseDoubleClick += new MouseButtonEventHandler(RestoreScalingFactor);
                uiScaleSliderChild.MouseDoubleClick += new MouseButtonEventHandler(RestoreScalingFactor);
            }
            else
            {
                Environment.Exit(0);
                System.Diagnostics.Process.GetCurrentProcess().Kill();
            }
        }

        #endregion

        #region Initial : 서버 체크

        private void CheckServer()
        {
            try
            {
                //DataStore.Log_Instance.
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        #endregion

        #region Initial : Login 처리 : Login()

        private bool Login()
        {
            bool loginFlag = false;

            //// 로그인 하기
            LoginPage login = new LoginPage();
            mfont = "ko-kr"; //2021-09-13 현달씨가 준 파일에 있음
            login.ShowDialog();

            // 비밀번호 변경
            Login_PwChange pwChange = new Login_PwChange();

            if (login.DialogResult == true)
            {
                // UserMenu 테이블을 이용해서 정보를 가져오기 위한 매개변수
                //  : 아이디를 세팅
                MainWindow.CurrentUser = login.strLogRegID;
                MainWindow.CurrentName = login.strUserName;
                MainWindow.CurrentPersonID = login.PersonID;
                //MainWindow.CurrentUser = "admin";

                loginFlag = true;

                // 개인정보활용 동의
                if (loginFlag == true && (login.AccessControl.Trim().Equals("")
                    || login.AccessControl.Trim().Equals("N")))
                {
                    Login_PersonInfoUse loginInfoUse = new Login_PersonInfoUse();
                    loginInfoUse.setData(login.UserName);

                    loginInfoUse.ShowDialog();

                    if (loginInfoUse.DialogResult != true)
                    {
                        loginFlag = false;
                    }
                    else if (loginInfoUse.DialogResult == true)
                    {
                        pwChange.setInitChangeMode();
                        pwChange.setPw(login.exPassword);
                        pwChange.ShowDialog();

                        if (pwChange.DialogResult == true)
                        {
                            login.initChange = "Y";
                        }
                        else
                        {
                            loginFlag = false;
                        }
                    }
                }

                // 초기변경 여부체크
                bool initChangFlag = false;
                // 비밀번호 초기 변경이 되지 않았다면 → (setDate null 이면 안된걸로 판단)
                if (loginFlag == true && !login.initChange.Trim().Equals("Y"))
                {
                    pwChange.setInitChangeMode();
                    pwChange.setPw(login.exPassword);
                    pwChange.ShowDialog();

                    if (pwChange.DialogResult != true)
                    {
                        loginFlag = false;
                    }
                    else
                    {
                        initChangFlag = true;
                    }
                }

                // 비밀번호 변경한지 3개월 지났으면 변경화면 띄우기
                if (loginFlag == true && login.initChange.Trim().Equals("Y") && login.dayDiff > 90 && initChangFlag != true)
                {
                    pwChange.setChangePwMode();
                    pwChange.setPw(login.exPassword);
                    pwChange.ShowDialog();

                    // 왜인지 모르게 false 를 반환한다면 종료
                    if (pwChange.DialogResult != true)
                    {
                        loginFlag = false;
                    }
                }
            }
            else // 로그인 실패 시
            {
                loginFlag = false;
            }

            return loginFlag;
        }

        #endregion

        void RestoreScalingFactor(object sender, MouseButtonEventArgs args)
        {
            ((Slider)sender).Value = 1.0;
        }

        #region menuLoad

        private void menuLoad()
        {
            SettingINI setting = new SettingINI();
            setting.GetSettingINI();
            this.FontSize = setting.setFontSize;
            StdFontSize = setting.setFontSize;
            this.FontFamily = setting.setFontFamily;
            this.FontStyle = setting.setFontStyle;
            this.FontWeight = setting.setFontWeight;
            uiScaleSlider.Value = setting.setMainScale;
            uiScaleSliderChild.Value = setting.setChildScale;
            TheHeight = SystemParameters.WorkArea.Height;
            TheWidth = SystemParameters.WorkArea.Width;
            TheFont = (int)this.FontSize;
            setMenuList();
            setTreeMenu();
            setMainmenu();

            string[] Person = lib.SetPerson();
            CurrentPerson = Person[0];
            CurrentPersonID = Person[1];

            MainMenu.FontSize = setting.setFontSize;
            MainMenu.FontFamily = setting.setFontFamily;
            MainMenu.FontStyle = setting.setFontStyle;
            MainMenu.FontWeight = setting.setFontWeight;

            BookMarkINI bookMarkINI = new BookMarkINI();
            bookMarkINI.GetBookMarkINI();
            strFavorites = bookMarkINI.strBookMarkMenu;

            if (strFavorites != null && strFavorites.Length > 0)
            {
                SetBookMarkListBox(strFavorites);
            }

            Win_Com_info();
        }

        #region 로그인 시 공지사항 띄우기 Win_Com_info()

        //공지사항 불러오기
        private void Win_Com_info()
        {
            //Win_com_Info Win_Com_info = new Win_com_Info();

            // 공지사항(공지사항)
            int i = 0;
            foreach (MenuViewModel mvm in MainWindow.mMenulist)
            {
                if (mvm.Menu.Equals("공지사항"))
                {
                    break;
                }
                i++;
            }
            try
            {
                if (MainWindow.MainMdiContainer.Children.Contains(MainWindow.mMenulist[i].subProgramID as MdiChild))
                {
                    (MainWindow.mMenulist[i].subProgramID as MdiChild).Focus();
                }
                else
                {
                    Type type = Type.GetType("WizMes_WellMade." + MainWindow.mMenulist[i].ProgramID.Trim(), true);
                    object uie = Activator.CreateInstance(type);

                    MainWindow.mMenulist[i].subProgramID = new MdiChild()
                    {
                        Title = "WizMes_WellMade [" + MainWindow.mMenulist[i].MenuID.Trim() + "] " + MainWindow.mMenulist[i].Menu.Trim() +
                                " (→" + MainWindow.mMenulist[i].ProgramID.Trim() + ")",
                        Height = SystemParameters.PrimaryScreenHeight * 0.8,
                        MaxHeight = SystemParameters.PrimaryScreenHeight * 0.85,
                        Width = SystemParameters.WorkArea.Width * 0.85,
                        MaxWidth = SystemParameters.WorkArea.Width,
                        Content = uie as UIElement,
                        Tag = MainWindow.mMenulist[i]
                    };
                    Lib.Instance.AllMenuLogInsert(MainWindow.mMenulist[i].MenuID, MainWindow.mMenulist[i].Menu, MainWindow.mMenulist[i].subProgramID);
                    MainWindow.MainMdiContainer.Children.Add(MainWindow.mMenulist[i].subProgramID as MdiChild);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("해당 화면이 존재하지 않습니다.");
            }
        }

        #endregion

        private void SetBookMarkListBox(string[] FavoritesArray)
        {
            foreach (MenuViewModel mvm in mMenulist)
            {
                foreach (string str in FavoritesArray)
                {
                    if (mvm.Menu.Trim().Equals(str))
                    {
                        listFavorites.Add(mvm);
                        ListBoxItem lbi = new ListBoxItem();
                        lbi = SetLBItem(mvm, mvm.Menu.Trim());

                        listBookMark.Items.Add(lbi);
                        break;
                    }
                }
            }
        }

        private void setMainmenu()
        {
            MenuItem mMenuItem0 = null;
            MenuItem mMenuItem1 = null;
            MenuItem mMenuItem3 = null;

            int FirstTrigger = 0;

            foreach (MenuViewModel mvm in mMenulist)
            {
                //상위 메뉴가 이상하여 확인결과  메뉴를 먼저 추가하고 다음위치 지정해야한다.
                //이전은 위치 먼저잡고 메뉴를 추가하고 있었다.
                if (mvm.Level == 0)
                {
                    // 둘리가 추가 - 공지사항, 로그 나부랭이 들을 메인 메뉴에서는 자사 설정으로 묶어서 표시하기 위해서 추가
                    if (mvm.MenuID.Trim().Substring(0, 1).Equals("0"))
                    {
                        if (FirstTrigger == 0)
                        {
                            mMenuItem0 = new MenuItem() { Header = "자사 시스템 설정", Tag = "" };
                            MainMenu.Items.Add(mMenuItem0);
                            FirstTrigger = 1;
                        }

                        mMenuItem1 = new MenuItem() { Header = mvm.Menu, Tag = mvm };
                        if (mMenuItem1 != null) { mMenuItem0.Items.Add(mMenuItem1); };

                        continue;
                    }

                    mMenuItem0 = new MenuItem() { Header = mvm.Menu, Tag = mvm };
                    if (mMenuItem0 != null)
                    {
                        if (!Lib.Instance.Right(mvm.MenuID.Replace(" ", ""), 1).Equals("0"))
                        {
                            mMenuItem0.Header = mvm.MenuID + "." + mvm.Menu;
                            MainMenu.Items.Add(mMenuItem0);
                            mMenuItem0.MouseLeftButtonUp += fmenu_click;
                        }
                        else
                        {
                            MainMenu.Items.Add(mMenuItem0);
                        }
                    };

                    //mMenuItem0 = new MenuItem() { Header = mvm.Menu, Tag = mvm };
                    //if (mMenuItem0 != null) { MainMenu.Items.Add(mMenuItem0); };
                    //mMenuItem0.Click += (s, e) => { fmenu_click(s, null); };
                }
                else if (mvm.Level == 1)
                {
                    mMenuItem1 = new MenuItem() { Header = mvm.Menu, Tag = mvm };
                    if (mMenuItem1 != null) { mMenuItem0.Items.Add(mMenuItem1); };


                }
                else if (mvm.Level == 3)
                {
                    mMenuItem3 = new MenuItem() { Header = mvm.Menu + "(" + mvm.MenuID + ")", Tag = mvm };
                    if (mMenuItem3 != null) { mMenuItem1.Items.Add(mMenuItem3); };
                    mMenuItem3.Click += (s, e) => { fmenu_click(s, null); };
                }
            }

        }

        private void setTreeMenu()
        {
            TreeViewItem mTreeViewItem0 = null;
            TreeViewItem mTreeViewItem1 = null;
            TreeViewItem mTreeViewItem3 = null;

            foreach (MenuViewModel mvm in mMenulist)
            {
                if (mvm.Level == 0)
                {
                    mTreeViewItem0 = new TreeViewItem() { Header = mvm.Menu, Tag = mvm, IsExpanded = true };
                    mTreeViewItem0.Template = (ControlTemplate)FindResource("ImageTreeViewItemEx");
                    if (mTreeViewItem0 != null)
                    {
                        if (!Lib.Instance.Right(mvm.MenuID.Replace(" ", ""), 1).Equals("0"))
                        {
                            mTreeViewItem0.Header = mvm.MenuID + "." + mvm.Menu;
                            mTree.Items.Add(mTreeViewItem0);
                            mTreeViewItem0.MouseLeftButtonUp += fmenu_click;
                        }
                        else
                        {
                            mTree.Items.Add(mTreeViewItem0);
                        }
                    };
                    //mTreeViewItem0 = new TreeViewItem() { Header = mvm.Menu, Tag = mvm };
                    //if (mTreeViewItem0 != null)
                    //{
                    //    mTree.Items.Add(mTreeViewItem0);
                    //};
                }
                else if (mvm.Level == 1)
                {
                    mTreeViewItem1 = new TreeViewItem() { Header = mvm.Menu, Tag = mvm };


                    //mTreeViewItem3 = new TreeViewItem() { Header = mvm.Menu + "(" + mvm.MenuID + ")", Tag = mvm };

                    mTreeViewItem1.Template = (ControlTemplate)FindResource("ImageTreeViewItemEx");
                    if (mTreeViewItem1 != null) { mTreeViewItem0.Items.Add(mTreeViewItem1); };
                }
                else if (mvm.Level == 3)
                {
                    mTreeViewItem3 = new TreeViewItem() { Header = mvm.Menu + "(" + mvm.MenuID + ")", Tag = mvm };
                    mTreeViewItem3.Template = (ControlTemplate)FindResource("ImageTreeViewItem");
                    if (mTreeViewItem3 != null) { mTreeViewItem1.Items.Add(mTreeViewItem3); };
                    mTreeViewItem3.MouseLeftButtonUp += fmenu_click;
                }
            }
        }



        private void fmenu_click(object sender, System.Windows.Input.MouseButtonEventArgs e)
        {
            string window_name = "";
            Type tt1 = null;
            MenuViewModel MenuViewModel = null;
            MdiChild mdiChild = null;
            object uie = null;

            //2021-07-27 최대화 후 다른 메뉴 실행하면 노말로 변경
            //2021-09-13 활성화된 자식폼이 없을때 에러
            if (MainMdiContainer.ActiveMdiChild != null && MainMdiContainer.ActiveMdiChild.WindowState == WindowState.Maximized)
            {
                MainMdiContainer.ActiveMdiChild.WindowState = WindowState.Normal;
            }



            if (sender is MenuItem)
            {
                MenuViewModel = (sender as MenuItem).Tag as MenuViewModel;
                window_name = MenuViewModel.MenuID;
                mdiChild = MenuViewModel.subProgramID as MdiChild;


            }
            if (sender is TreeViewItem)
            {
                MenuViewModel = (sender as TreeViewItem).Tag as MenuViewModel;
                window_name = MenuViewModel.MenuID;
                mdiChild = MenuViewModel.subProgramID as MdiChild;

                this.Title = "WizMes 생산관리시스템    " + MenuViewModel.Menu;
            }

            if (sender is ListBoxItem)
            {
                MenuViewModel = (sender as ListBoxItem).Tag as MenuViewModel;
                window_name = MenuViewModel.MenuID;
                mdiChild = MenuViewModel.subProgramID as MdiChild;


            }

            if (MainMdiContainer.Children.Contains(mdiChild))
            {
                if (mdiChild.WindowState == WindowState.Minimized)
                {
                    mdiChild.WindowState = WindowState.Normal;
                }
                mdiChild.Focus();


            }
            else
            {
                try
                {
                    tt1 = Type.GetType("WizMes_WellMade." + MenuViewModel.ProgramID.Trim(), true);
                    uie = Activator.CreateInstance(tt1);

                    MenuViewModel.subProgramID = new MdiChild()
                    {
                        Title = "WizMes_WellMade [" + MenuViewModel.MenuID + "] " + MenuViewModel.Menu + " (→" + MenuViewModel.ProgramID.Trim() + ")",
                        Height = SystemParameters.PrimaryScreenHeight * 0.8,
                        MaxHeight = SystemParameters.PrimaryScreenHeight * 0.85,
                        MinHeight = 640,
                        Width = SystemParameters.WorkArea.Width * 0.85,
                        MaxWidth = SystemParameters.WorkArea.Width,
                        MinWidth = 800,
                        FontSize = 12,
                        Content = uie as UIElement,
                        Tag = MenuViewModel
                    };

                    Lib.Instance.AllMenuLogInsert(MenuViewModel.MenuID, MenuViewModel.Menu, MenuViewModel.ProgramID);
                    MainMdiContainer.Children.Add(MenuViewModel.subProgramID as MdiChild);
                }
                catch (Exception ex)
                {
                    MessageBox.Show("해당화면이 존재하지 않습니다. (" + (MenuViewModel.ProgramID != null ? MenuViewModel.ProgramID.Trim() : "") + ")");
                    return;
                }
                //this.Title = "WizMes 생산관리시스템    " + MenuViewModel.Menu;

            }
        }

        private void CommonClick(object sender, System.Windows.Input.MouseButtonEventArgs e)
        {
            System.Diagnostics.ProcessStartInfo startInfo = new System.Diagnostics.ProcessStartInfo();

            string strPath = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetEntryAssembly().Location);//+ "\\wizmes.exe";
            strPath = strPath + "\\WizMes_WellMade2.exe";
            startInfo.FileName = strPath;
            startInfo.Arguments = CurrentUser;
            startInfo.RedirectStandardOutput = true;
            startInfo.RedirectStandardError = true;
            startInfo.UseShellExecute = false;
            startInfo.CreateNoWindow = true;

            System.Diagnostics.Process processTemp = new System.Diagnostics.Process();
            processTemp.StartInfo = startInfo;
            //processTemp.EnableRaisingEvents = true;
            try
            {
                processTemp.Start();
                //Environment.Exit(0);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }

        private void setMenuList()
        {
            //2021-09-13 현달씨 DBclose 추가
            try
            {
                string str = string.Empty;
                //string[] arg = Environment.GetCommandLineArgs();
                //CurrentUser = arg[1];
                //CurrentUser = "admin";

                CompanyID = Lib.Instance.LogCompany(CurrentUser);

                Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                //sqlParameter.Add("@sUserID", CurrentUser);
                //sqlParameter.Add("@sPgGubun", "9");
                sqlParameter.Add("@sUserID", CurrentUser);
                sqlParameter.Add("@sPgGubun", "7");

                DataSet ds = DataStore.Instance.ProcedureToDataSet("xp_Menu_sUserMenu", sqlParameter, false);

                if (ds != null && ds.Tables.Count > 0)
                {
                    DataTable dt = ds.Tables[0];

                    if (dt.Rows.Count == 0)
                    {
                        // admin 으로 접속했는데 조회된 데이터가 없다고 했을때?
                        // 자동으로 메뉴 다 긁어와서 UserMenu에 추가
                        if (CurrentPersonID != null
                            && CurrentPersonID.Trim().Equals("admin"))
                        {
                            if (FillAdminUserMenu())
                            {
                                setMenuList();
                            }
                        }
                        else
                        {
                            MessageBox.Show("조회된 데이터가 없습니다.");
                        }
                    }
                    else
                    {
                        DataRowCollection drc = dt.Rows;

                        // 생산 새로운 화면 확인을 위한 → UserID 가 sadmin(생산관리자) 로 접속했을경우에 볼수 있도록 추가
                        if (MainWindow.CurrentUser.Trim().Equals("padmin"))
                        {
                            // 작업시지 뉴
                            var mModel = new MenuViewModel()
                            {
                                MenuID = "5001",
                                Menu = "작지화면 New",
                                Level = 0,
                                ParentID = "0",
                                SelectClss = "*",
                                AddNewClss = "*",
                                UpdateClss = "*",
                                DeleteClss = "*",
                                PrintClss = "*",
                                ProgramID = "Win_prd_PlanInput_U_New",
                                subProgramID = "Win_prd_PlanInput_U_New"
                            };

                            str = mModel.Menu.Trim();
                            mMenulist.Add(mModel);

                            // 작업시지 뉴
                            var mModel1 = new MenuViewModel()
                            {
                                MenuID = "5001",
                                Menu = "Test",
                                Level = 0,
                                ParentID = "0",
                                SelectClss = "*",
                                AddNewClss = "*",
                                UpdateClss = "*",
                                DeleteClss = "*",
                                PrintClss = "*",
                                ProgramID = "Win_prd_Test",
                                subProgramID = "Win_prd_Test"
                            };

                            str = mModel1.Menu.Trim();
                            mMenulist.Add(mModel1);

                            // LoadingTest
                            var mModel2 = new MenuViewModel()
                            {
                                MenuID = "5001",
                                Menu = "Test",
                                Level = 0,
                                ParentID = "0",
                                SelectClss = "*",
                                AddNewClss = "*",
                                UpdateClss = "*",
                                DeleteClss = "*",
                                PrintClss = "*",
                                ProgramID = "Win_prd_MCTool_U_New",
                                subProgramID = "Win_prd_MCTool_U_New"
                            };

                            str = mModel2.Menu.Trim();
                            mMenulist.Add(mModel2);
                        }

                        foreach (DataRow item in drc)
                        {
                            var mMenuviewModel = new MenuViewModel()
                            {
                                MenuID = item["MenuID"] as string,
                                Menu = item["Menu"] as string,
                                Level = Convert.ToInt32(item["Level"]),
                                ParentID = item["ParentID"] as string,
                                SelectClss = item["SelectClss"] as string,
                                AddNewClss = item["AddNewClss"] as string,
                                UpdateClss = item["UpdateClss"] as string,
                                DeleteClss = item["DeleteClss"] as string,
                                PrintClss = item["PrintClss"] as string,
                                ////Remark = "WizMes_WellMade." + item["Remark"].ToString(),
                                ////subRemark = item["Remark"] as object,
                                ProgramID = item["ProgramID"] as string,
                                subProgramID = item["ProgramID"] as object
                            };

                            if ((mMenuviewModel.MenuID.Substring(0, 1)).Equals("0"))
                            {
                                str = mMenuviewModel.Menu.Trim();
                                mMenulist.Add(mMenuviewModel);
                            }
                            else if ((mMenuviewModel.MenuID.Substring(0, 1)).Equals("1"))
                            {
                                str = mMenuviewModel.Menu.Trim();
                                mMenulist.Add(mMenuviewModel);
                            }
                            else if ((mMenuviewModel.MenuID.Substring(0, 1)).Equals("3"))
                            {
                                str = mMenuviewModel.Menu.Trim();
                                mMenulist.Add(mMenuviewModel);
                            }
                            else if ((mMenuviewModel.MenuID.Substring(0, 1)).Equals("4")) //자재 4번으로 수정하여 조건 추가
                            {
                                str = mMenuviewModel.Menu.Trim();
                                mMenulist.Add(mMenuviewModel);
                            }
                            else if ((mMenuviewModel.MenuID.Substring(0, 1)).Equals("5"))
                            {
                                if ((mMenuviewModel.MenuID.Substring(0, 2)).Equals("56")) // 모니터링
                                {
                                    continue;
                                }

                                str = mMenuviewModel.Menu.Trim();
                                mMenulist.Add(mMenuviewModel);
                            }
                            else if ((mMenuviewModel.MenuID.Substring(0, 1)).Equals("6"))
                            {
                                str = mMenuviewModel.Menu.Trim();
                                mMenulist.Add(mMenuviewModel);
                            }
                            else if ((mMenuviewModel.MenuID.Substring(0, 1)).Equals("7"))
                            {
                                str = mMenuviewModel.Menu.Trim();
                                mMenulist.Add(mMenuviewModel);
                            }

                            else if ((mMenuviewModel.MenuID.Substring(0, 1)).Equals("8"))
                            {
                                str = mMenuviewModel.Menu.Trim();
                                mMenulist.Add(mMenuviewModel);
                            }

                            else if ((mMenuviewModel.MenuID.Substring(0, 1)).Equals("9"))
                            {
                                str = mMenuviewModel.Menu.Trim();
                                mMenulist.Add(mMenuviewModel);
                            }

                            //{
                            //if (!(mMenuviewModel.MenuID.Substring(0, 2)).Equals("31") && !(mMenuviewModel.MenuID.Substring(0, 2)).Equals("32"))
                            //{
                            //    //str = mMenuviewModel.Menu.Replace(" ", "");
                            //    str = mMenuviewModel.Menu.Trim();
                            //    mMenulist.Add(mMenuviewModel);
                            //}                            
                            //}
                        }
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

        // 만약에 admin 으로 접속을 했는데, 유저메뉴가 비워져 있다면?? 그럼 알아서 채워줘야겠지?
        // 그걸 위한 함수를 추가
        private bool FillAdminUserMenu()
        {
            bool flag = false;

            List<Procedure> Prolist = new List<Procedure>();
            List<Dictionary<string, object>> ListParameter = new List<Dictionary<string, object>>();

            try
            {
                Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                sqlParameter.Clear();

                sqlParameter.Add("PersonID", CurrentPersonID);

                Procedure pro1 = new Procedure();
                pro1.Name = "xp_Person_iAdminUserMenu";
                pro1.OutputUseYN = "N";
                pro1.OutputName = "StuffInID";
                pro1.OutputLength = "12";

                Prolist.Add(pro1);
                ListParameter.Add(sqlParameter);

                List<KeyValue> list_Result = new List<KeyValue>();
                list_Result = DataStore.Instance.ExecuteAllProcedureOutputGetCS(Prolist, ListParameter);
                string sGetID = string.Empty;

                if (list_Result[0].key.ToLower() == "success")
                {
                    //list_Result.RemoveAt(0);
                    //for (int i = 0; i < list_Result.Count; i++)
                    //{
                    //    KeyValue kv = list_Result[i];
                    //    if (kv.key == "StuffInID")
                    //    {
                    //        sGetID = kv.value;
                    //        flag = true;
                    //    }
                    //}
                    flag = true;
                }
                else
                {
                    MessageBox.Show("조회된 메뉴가 없습니다 + Admin 유저 메뉴 INSERT 실패.");
                    return false;
                }
                DataStore.Instance.CloseConnection();
                return flag;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                DataStore.Instance.CloseConnection();
                return false;
            }
            //finally
            //{
            //    DataStore.Instance.CloseConnection();
            //}           
        }

        #endregion

        private void mmClick(object sender, RoutedEventArgs e)
        {
            if (this.mMenuWidth.Width != new GridLength(0))
            {
                this.mMenuWidth.Width = new GridLength(0);
            }
            else
            {
                this.mMenuWidth.Width = new GridLength(150);
            }
        }

        #region 메인화면 닫을시 종료
        protected override void OnClosing(CancelEventArgs e1)
        {
            try
            {
                if (MessageBox.Show("WizMes_WellMade를 종료하시겠습니까?", "", MessageBoxButton.YesNo) == MessageBoxResult.Yes)
                {
                    DataStore.Instance.InsertLogByFormAllUpdate(mainStDate, mainStTime);
                    Environment.Exit(0);
                    System.Diagnostics.Process.GetCurrentProcess().Kill();
                }
                else
                {
                    e1.Cancel = true;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("오류 발생, 오류 내용 : " + ex.ToString());
            }
        }
        void _HideThisWindow()
        {
            this.Hide();
        }

        #endregion

        private void OnClosing(object sender, RoutedEventArgs e)
        {
            if (MessageBox.Show("WizMes_WellMade를 종료하시겠습니까?", "", MessageBoxButton.YesNo) == MessageBoxResult.Yes)
            {
                DataStore.Instance.InsertLogByFormAllUpdate(mainStDate, mainStTime);
                SaveFontSetting();
                Environment.Exit(0);
                System.Diagnostics.Process.GetCurrentProcess().Kill();
            }
        }

        private void btnBefore_Click(object sender, RoutedEventArgs e)
        {
            int index = -1;
            for (int i = 0; i < MainMdiContainer.Children.Count; i++)
            {
                if (MainMdiContainer.Children[i].Equals(MainMdiContainer.ActiveMdiChild))
                {
                    index = i;
                    break;
                }
            }

            if (MainMdiContainer.Children.Count > 1)
            {
                if (index > 0)
                {
                    MainMdiContainer.Children[index - 1].Focus();
                }
                else if (index == 0)
                {
                    MainMdiContainer.Children[MainMdiContainer.Children.Count - 1].Focus();
                }
            }
        }

        private void btnAfter_Click(object sender, RoutedEventArgs e)
        {
            int index = -1;
            for (int i = 0; i < MainMdiContainer.Children.Count; i++)
            {
                if (MainMdiContainer.Children[i].Equals(MainMdiContainer.ActiveMdiChild))
                {
                    index = i;
                    break;
                }
            }

            if (MainMdiContainer.Children.Count > 0)
            {
                if (index == MainMdiContainer.Children.Count - 1)
                {
                    MainMdiContainer.Children[0].Focus();
                }
                else if (index >= 0)
                {
                    MainMdiContainer.Children[index + 1].Focus();
                }
            }
        }

        //글꼴 설정
        private void SetMySetting(object sender, RoutedEventArgs e)
        {
            TheFont = (int)this.FontSize;
            PopUp.FontPopUP fontPop = new PopUp.FontPopUP(MainMenu);
            fontPop.ShowDialog();

            if (fontPop.DialogResult == true)
            {
                this.FontFamily = fontPop.ResultFontFamily;
                TheFont = (int)fontPop.ResultFontSize;
                StdFontSize = (double)fontPop.ResultFontSize;
                this.FontSize = TheFont;
                this.FontStyle = fontPop.ResultTypeFace.Style;
                this.FontWeight = fontPop.ResultTypeFace.Weight;
                MainMenu.FontFamily = fontPop.ResultFontFamily;
                MainMenu.FontSize = TheFont;
                MainMenu.FontStyle = fontPop.ResultTypeFace.Style;
                MainMenu.FontWeight = fontPop.ResultTypeFace.Weight;

                SaveFontSetting();
            }
        }

        private void SaveFontSetting()
        {
            SettingINI setting = new SettingINI();
            SettingINI.myFontSize.Clear();
            SettingINI.myFontFamily.Clear();
            SettingINI.myFontStyle.Clear();
            SettingINI.myFontWeight.Clear();
            SettingINI.myFontSize.Append(MainMenu.FontSize.ToString());
            SettingINI.myFontFamily.Append(MainMenu.FontFamily.ToString());
            SettingINI.myFontStyle.Append(MainMenu.FontStyle.ToString());
            SettingINI.myFontWeight.Append(MainMenu.FontWeight.ToString());

            SettingINI.myMainScale.Clear();
            SettingINI.myChildScale.Clear();
            SettingINI.myMainScale.Append(uiScaleSlider.Value.ToString());
            SettingINI.myChildScale.Append(uiScaleSliderChild.Value.ToString());
            setting.WriteSettingINI();
        }

        private void SaveBookMark()
        {
            BookMarkINI bookMarkINI = new BookMarkINI();
            BookMarkINI.myBookMarkMenu.Clear();

            for (int i = 0; i < listBookMark.Items.Count; i++)
            {
                var mvm = (listBookMark.Items[i] as ListBoxItem).Tag as MenuViewModel;
                if (i == listBookMark.Items.Count - 1)
                {
                    BookMarkINI.myBookMarkMenu.Append(mvm.Menu.Trim());
                }
                else
                {
                    BookMarkINI.myBookMarkMenu.Append(mvm.Menu.Trim());
                    BookMarkINI.myBookMarkMenu.Append("/");
                }
            }
            bookMarkINI.WriteBookMarkINI();
        }

        //모두닫기
        private void btnAllClose_Click(object sender, RoutedEventArgs e)
        {
            MainMdiContainer.Children.Clear();
        }

        ///// <summary>
        /// 메인 컨테이너 사이즈 조절(휠로)
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Window_MouseWheel(object sender, MouseWheelEventArgs e)
        {
            if (Keyboard.IsKeyDown(Key.LeftCtrl) || Keyboard.IsKeyDown(Key.RightCtrl))
            {
                uiScaleSliderChild.Value += e.Delta * 0.0001;
            }
        }


        private void btnFavoriteAddtion_Click(object sender, RoutedEventArgs e)
        {
            ContextMenu menu = btnFavoriteAddtion.ContextMenu;
            menu.StaysOpen = true;
            menu.IsOpen = true;
        }


        private void ShowMenuAdd_Click(object sender, RoutedEventArgs e)
        {
            listBookMark.SelectedIndex = -1;
            //현재 즐겨찾기 메뉴 추가
            if (listFavorites.Count > 0)
            {
                listFavorites.Clear();
            }

            if (listBookMark.Items.Count > 0)
            {
                foreach (ListBoxItem listItem in listBookMark.Items)
                {
                    var Compare = listItem.Tag as MenuViewModel;
                    listFavorites.Add(Compare);
                }
            }

            PopUp.FavoriterAddtionPopUP BookMarkPopUp = new PopUp.FavoriterAddtionPopUP(mMenulist, listFavorites);
            BookMarkPopUp.ShowDialog();
            if (BookMarkPopUp.DialogResult == true)
            {
                //listFavorites.Clear();
                listFavorites = BookMarkPopUp.listBMMenu;

                listBookMark.Items.Clear();
                foreach (MenuViewModel mvm in listFavorites)
                {
                    ListBoxItem lbi = new ListBoxItem();
                    lbi = SetLBItem(mvm, mvm.Menu.Trim());

                    listBookMark.Items.Add(lbi);
                }
                SaveBookMark();
            }
        }


        /// <summary>
        /// 현재 포커싱된 화면을 즐겨찾기에 추가한다
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void CurrentBookMark_Click(object sender, RoutedEventArgs e)
        {
            bool flag = true;

            for (int i = 0; i < MainMdiContainer.Children.Count; i++)
            {
                //munu번호와 menu 네임을 받는다
                if (MainMdiContainer.Children[i].Focused == true)
                {
                    currentMenuViewModel = (MainMdiContainer.Children[i].Tag as MenuViewModel);
                }
            }

            foreach (ListBoxItem listItem in listBookMark.Items)
            {

                var Compare = listItem.Tag as MenuViewModel;
                if (currentMenuViewModel == Compare)
                {
                    flag = false;
                    break;
                }
            }

            if (flag)
            {
                ListBoxItem lbi = new ListBoxItem();
                lbi = SetLBItem(currentMenuViewModel, currentMenuViewModel.Menu.Trim());

                listBookMark.Items.Add(lbi);
                //listBookMark.Items.Add(str);
                SaveBookMark();
            }
            else
            {
                MessageBox.Show("이미 같은 이름의 메뉴가 추가되어 있습니다.");
            }
        }

        /// <summary>
        /// 메뉴번호를 tag로 저장하고 content를 네임으로 저장한다
        /// 리스트 아이템으로 contextMenu도 추가해준다
        /// </summary>
        /// <param name="strTag"></param>
        /// <param name="strItem"></param>
        /// <returns></returns>
        private ListBoxItem SetLBItem(MenuViewModel mvm, string strItem)
        {
            ListBoxItem listBoxItem = new ListBoxItem();
            ContextMenu contextMenu = new ContextMenu();

            MenuItem menuOne = new MenuItem();
            menuOne.Header = "선택화면으로 이동";
            menuOne.Tag = strItem;
            menuOne.Click += new RoutedEventHandler(btnOneMenuClick);

            MenuItem menuTwo = new MenuItem();
            menuTwo.Header = "선택화면삭제";
            menuTwo.Tag = strItem;
            menuTwo.Click += new RoutedEventHandler(btnTwoMenuClick);

            contextMenu.Items.Add(menuOne);
            contextMenu.Items.Add(menuTwo);
            listBoxItem.ContextMenu = contextMenu;
            listBoxItem.MouseRightButtonUp += new MouseButtonEventHandler(btnFavoritesMenu);
            //lbi.Click += new RoutedEventHandler(btnFavoritesMenuSee);
            listBoxItem.MouseLeftButtonUp += new MouseButtonEventHandler(btnFavoritesMenuSee);
            listBoxItem.Content = strItem;
            listBoxItem.Tag = mvm;

            return listBoxItem;
        }

        //보이는 메뉴 클릭
        private void btnFavoritesMenuSee(object sender, RoutedEventArgs e)
        {
            ListBoxItem lbxSend = listBookMark.SelectedItem as ListBoxItem;
            fmenu_click(lbxSend, null);
        }

        //보이는 메뉴 클릭
        private void btnFavoritesMenuSee(object sender, MouseButtonEventArgs e)
        {
            var strSend = (listBookMark.SelectedItem as ListBoxItem);
            fmenu_click(strSend, null);
        }

        //선택화면의 메뉴 보이기
        private void btnFavoritesMenu(object sender, MouseButtonEventArgs e)
        {
            ContextMenu menu = (sender as ListBoxItem).ContextMenu;
            menu.StaysOpen = true;
            menu.IsOpen = true;
        }

        private void btnOneMenuClick(object sender, RoutedEventArgs e)
        {
            var strSend = (listBookMark.SelectedItem as ListBoxItem);
            fmenu_click(strSend, null);
        }

        private void btnTwoMenuClick(object sender, RoutedEventArgs e)
        {
            if (MessageBox.Show("선택하신 항목을 즐겨찾기에서 삭제하시겠습니까?", "즐겨찾기 목록 편집", MessageBoxButton.YesNo) == MessageBoxResult.Yes)
            {
                listBookMark.Items.Remove(listBookMark.SelectedItem);
            }
        }

        private void ChildbtnSearchEvent(object sender, ExecutedRoutedEventArgs e)
        {
            Object obj = MainMdiContainer.ActiveMdiChild.Content;

            if (obj != null)
            {
                UserControl CurrentUserControl = obj as UserControl;

                if (CurrentUserControl != null)
                {
                    object objSearch = CurrentUserControl.FindName("btnSearch");

                    if (objSearch != null)
                    {
                        if ((objSearch as Button).IsEnabled == true)
                        {
                            (objSearch as Button).RaiseEvent(new RoutedEventArgs(ButtonBase.ClickEvent));
                        }
                    }
                }
            }
        }

        private void ChildbtnCloseEvent(object sender, ExecutedRoutedEventArgs e)
        {
            Object obj = MainMdiContainer.ActiveMdiChild.Content;

            if (obj != null)
            {
                UserControl CurrentUserControl = obj as UserControl;

                if (CurrentUserControl != null)
                {
                    object objClose = CurrentUserControl.FindName("btnClose");

                    if (objClose != null)
                    {
                        if ((objClose as Button).IsEnabled == true)
                        {
                            (objClose as Button).RaiseEvent(new RoutedEventArgs(ButtonBase.ClickEvent));
                        }
                    }
                }
            }
        }

        private void ChildbtnCancelEvent(object sender, ExecutedRoutedEventArgs e)
        {
            Object obj = MainMdiContainer.ActiveMdiChild.Content;

            if (obj != null)
            {
                UserControl CurrentUserControl = obj as UserControl;

                if (CurrentUserControl != null)
                {
                    object objCancel = CurrentUserControl.FindName("btnCancel");

                    if (objCancel != null)
                    {
                        if ((objCancel as Button).Visibility == Visibility.Visible)
                        {
                            (objCancel as Button).RaiseEvent(new RoutedEventArgs(ButtonBase.ClickEvent));
                        }
                    }
                }
            }
        }

        private void ChildbtnAddEvent(object sender, ExecutedRoutedEventArgs e)
        {
            Object obj = MainMdiContainer.ActiveMdiChild.Content;

            if (obj != null)
            {
                UserControl CurrentUserControl = obj as UserControl;

                if (CurrentUserControl != null)
                {
                    object objAdd = CurrentUserControl.FindName("btnAdd");

                    if (objAdd != null)
                    {
                        if ((objAdd as Button).IsEnabled == true)
                        {
                            (objAdd as Button).RaiseEvent(new RoutedEventArgs(ButtonBase.ClickEvent));
                        }
                    }
                }
            }
        }

        private void ChildbtnUpdateEvent(object sender, ExecutedRoutedEventArgs e)
        {
            Object obj = MainMdiContainer.ActiveMdiChild.Content;

            if (obj != null)
            {
                UserControl CurrentUserControl = obj as UserControl;

                if (CurrentUserControl != null)
                {
                    object objUpdate = CurrentUserControl.FindName("btnUpdate");
                    object objEdit = CurrentUserControl.FindName("btnEdit");

                    if (objUpdate != null)
                    {
                        if ((objUpdate as Button).IsEnabled == true)
                        {
                            (objUpdate as Button).RaiseEvent(new RoutedEventArgs(ButtonBase.ClickEvent));
                        }
                    }
                    else if (objEdit != null)
                    {
                        if ((objEdit as Button).IsEnabled == true)
                        {
                            (objEdit as Button).RaiseEvent(new RoutedEventArgs(ButtonBase.ClickEvent));
                        }
                    }
                }
            }
        }

        private void ChildbtnDeleteEvent(object sender, ExecutedRoutedEventArgs e)
        {
            Object obj = MainMdiContainer.ActiveMdiChild.Content;

            if (obj != null)
            {
                UserControl CurrentUserControl = obj as UserControl;

                if (CurrentUserControl != null)
                {
                    object objDelete = CurrentUserControl.FindName("btnDelete");

                    if (objDelete != null)
                    {
                        if ((objDelete as Button).IsEnabled == true)
                        {
                            (objDelete as Button).RaiseEvent(new RoutedEventArgs(ButtonBase.ClickEvent));
                        }
                    }
                }
            }
        }

        private void ChildbtnSaveEvent(object sender, ExecutedRoutedEventArgs e)
        {
            Object obj = MainMdiContainer.ActiveMdiChild.Content;

            if (obj != null)
            {
                UserControl CurrentUserControl = obj as UserControl;

                if (CurrentUserControl != null)
                {
                    object objSave = CurrentUserControl.FindName("btnSave");

                    if (objSave != null)
                    {
                        if ((objSave as Button).Visibility == Visibility.Visible)
                        {
                            (objSave as Button).RaiseEvent(new RoutedEventArgs(ButtonBase.ClickEvent));
                        }
                    }
                }
            }
        }

        private void ChildbtnExcelEvent(object sender, ExecutedRoutedEventArgs e)
        {
            Object obj = MainMdiContainer.ActiveMdiChild.Content;

            if (obj != null)
            {
                UserControl CurrentUserControl = obj as UserControl;

                if (CurrentUserControl != null)
                {
                    object objExcel = CurrentUserControl.FindName("btnExcel");

                    if (objExcel != null)
                    {
                        if ((objExcel as Button).IsEnabled == true)
                        {
                            (objExcel as Button).RaiseEvent(new RoutedEventArgs(ButtonBase.ClickEvent));
                        }
                    }
                }
            }
        }

        private void ChildbtnPrintEvent(object sender, ExecutedRoutedEventArgs e)
        {
            Object obj = MainMdiContainer.ActiveMdiChild.Content;

            if (obj != null)
            {
                UserControl CurrentUserControl = obj as UserControl;

                if (CurrentUserControl != null)
                {
                    object objPrint = CurrentUserControl.FindName("btnPrint");

                    if (objPrint != null)
                    {
                        if ((objPrint as Button).IsEnabled == true)
                        {
                            (objPrint as Button).RaiseEvent(new RoutedEventArgs(ButtonBase.ClickEvent));
                        }
                    }
                }
            }
        }

        private void ChildbtnUploadEvent(object sender, ExecutedRoutedEventArgs e)
        {
            Object obj = MainMdiContainer.ActiveMdiChild.Content;
            if (obj != null)
            {
                UserControl CurrentUserControl = obj as UserControl;
                if (CurrentUserControl != null)
                {
                    object objPrint = CurrentUserControl.FindName("btnUpload");
                    if (objPrint != null)
                    {
                        if ((objPrint as Button).IsEnabled == true)
                        {
                            (objPrint as Button).RaiseEvent(new RoutedEventArgs(ButtonBase.ClickEvent));
                        }
                    }
                }
            }
        }

        private void btnUpAndDown_Click(object sender, RoutedEventArgs e)
        {
            if (bdrFavorite.Width < 20)
            {
                btnUpAndDown.Content = "-";
                bdrFavorite.Width = 250;
                listBookMark.Visibility = Visibility.Visible;
                tbkFavorite.Text = "즐겨찾기 접기"
;
            }
            else if (bdrFavorite.Width > 200)
            {
                btnUpAndDown.Content = "+";
                bdrFavorite.Width = 15;
                listBookMark.Visibility = Visibility.Hidden;
                tbkFavorite.Text = "즐겨찾기 펼치기";
            }
        }
    }

    public class MenuViewModel
    {
        public override string ToString()
        {
            return (this.ReportAllProperties());
        }

        public string MenuID { get; set; }
        public string Menu { get; set; }
        public int Level { get; set; }
        public string ParentID { get; set; }
        public string SelectClss { get; set; }
        public string AddNewClss { get; set; }
        public string UpdateClss { get; set; }
        public string DeleteClss { get; set; }
        public string PrintClss { get; set; }
        public string seq { get; set; }
        public string Remark { get; set; }
        public object subRemark { get; set; }
        public string ProgramID { get; set; }
        public object subProgramID { get; set; }
    }
}

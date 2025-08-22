using System.Collections.Generic;
using System.Windows;
using System.Windows.Controls;

namespace WizMes_WellMade.PopUp
{
    /// <summary>
    /// FavoriterAddtionPopUP.xaml에 대한 상호 작용 논리
    /// </summary>
    public partial class FavoriterAddtionPopUP : Window
    {
        /// <summary>
        /// 전체 메뉴 list
        /// </summary>
        List<string> listMenu = new List<string>();

        /// <summary>
        /// 즐겨찾기 메뉴 list
        /// </summary>
        public List<MenuViewModel> listBMMenu = new List<MenuViewModel>();

        /// <summary>
        /// 첫 listBox 설정시 사용
        /// </summary>
        List<string> listRemoveString = new List<string>();

        List<string> First = new List<string>();
        Dictionary<string, List<string>> SecontDitionary = new Dictionary<string, List<string>>();
        Dictionary<string, List<string>> ThirdDitionary = new Dictionary<string, List<string>>();

        public FavoriterAddtionPopUP()
        {
            InitializeComponent();
        }

        /// <summary>
        /// List MenuViewModel 과 스트링 배열 List를 받는 생성자
        /// </summary>
        /// <param name="listmenuViewModels"></param>
        /// <param name="listFAMen"></param>
        public FavoriterAddtionPopUP(List<MenuViewModel> listmenuViewModels, List<MenuViewModel> listFAMen)
        {
            InitializeComponent();
            SetlistBoxBookMarkMenu(listFAMen);
            ForTreeViewSetting(listmenuViewModels);
            SetFirstMenu();
        }

        /// <summary>
        /// 즐겨찾기 메뉴 가져오기
        /// </summary>
        /// <param name="listStr"></param>
        private void SetlistBoxBookMarkMenu(List<MenuViewModel> listMVM)
        {
            dgdBoxBookMarkMenu.Items.Clear();
            listBMMenu = listMVM;

            foreach (MenuViewModel mvm in listBMMenu)
            {
                var BookMarkValue = new ThirdMenu()
                {
                    SecondValue = mvm,
                    ThirdValue = mvm.Menu.Trim()
                };

                dgdBoxBookMarkMenu.Items.Add(BookMarkValue);
            }
        }

        List<string> TreeFirst = new List<string>();
        List<string[]> TreeSecond = new List<string[]>();
        List<MenuViewModel> TreeThird = new List<MenuViewModel>();
        List<string> ThirdConnectSecond = new List<string>();

        /// <summary>
        /// Tree 구조를 보여주기 위한 과정
        /// </summary>
        /// <param name="viewModels"></param>
        private void ForTreeViewSetting(List<MenuViewModel> viewModels)
        {
            string strFirst = string.Empty;
            string strSecond = string.Empty;

            for (int i = 0; i < viewModels.Count; i++)
            {
                var mvm = viewModels[i] as MenuViewModel;

                if (mvm.Level == 0)
                {
                    TreeFirst.Add(mvm.Menu);
                    strFirst = mvm.Menu;
                }
                else if (mvm.Level == 1)
                {
                    string[] strArray = { strFirst, mvm.Menu };
                    TreeSecond.Add(strArray);
                    strSecond = mvm.Menu.Trim();
                }
                else if (mvm.Level == 3)
                {
                    ThirdConnectSecond.Add(strSecond);
                    TreeThird.Add(mvm);
                }
            }
        }

        string strFirstValue = string.Empty;
        string strSecondValue = string.Empty;

        /// <summary>
        /// 최상위 메뉴 세팅
        /// </summary>
        private void SetFirstMenu()
        {
            if (dgdFirst.Items.Count > 0)
            {
                dgdFirst.Items.Clear();
            }

            for (int i = 0; i < TreeFirst.Count; i++)
            {
                var First = new FirstMenu() { FirstValue = TreeFirst[i] };
                dgdFirst.Items.Add(First);
            }

            dgdFirst.SelectedIndex = 0;
        }

        //상위 메뉴 세팅
        private void dgdFirst_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (dgdFirst.Items.Count > 0)
            {
                var FirstValue = dgdFirst.SelectedItem as FirstMenu;
                strFirstValue = FirstValue.FirstValue;
            }

            if (dgdSecond.Items.Count > 0)
            {
                dgdSecond.Items.Clear();
            }

            for (int i = 0; i < TreeSecond.Count; i++)
            {
                var Second = new SecondMenu()
                {
                    FirstValue = TreeSecond[i][0].ToString(),
                    SecondValue = TreeSecond[i][1].ToString()
                };

                if (Second.FirstValue == strFirstValue)
                {
                    dgdSecond.Items.Add(Second);
                }
            }

            dgdSecond.SelectedIndex = 0;
        }

        /// <summary>
        /// 메뉴 세팅
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void dgdSecond_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            //세팅을 위한 데이터 GET
            if (dgdSecond.Items.Count > 0)
            {
                var SecondValue = dgdSecond.SelectedItem as SecondMenu;
                strSecondValue = SecondValue.SecondValue;
            }

            SetThirdMenu();
        }

        /// <summary>
        /// 세팅된 데이터를 이요하여 메뉴 세팅
        /// </summary>
        private void SetThirdMenu()
        {
            bool flag = true;

            //세팅 전 정리
            if (dgdThird.Items.Count > 0)
            {
                dgdThird.Items.Clear();
            }

            for (int i = 0; i < TreeThird.Count; i++)
            {
                var mvm = TreeThird[i];
                var Third = new ThirdMenu()
                {
                    FirstValue = ThirdConnectSecond[i],
                    SecondValue = mvm,
                    ThirdValue = mvm.Menu.Trim()
                };

                flag = true;

                if (listBMMenu.Count > 0)
                {
                    for (int j = 0; j < listBMMenu.Count; j++)
                    {
                        var menuViewModel = listBMMenu[j] as MenuViewModel;
                        if (Third.ThirdValue.Equals(menuViewModel.Menu.Trim()))
                        {
                            flag = false;
                            break;
                        }
                    }
                }

                if (flag)
                {
                    if (Third.FirstValue.Trim() == strSecondValue)
                    {
                        dgdThird.Items.Add(Third);
                    }
                }
            }
        }

        //즐겨찾기 메뉴로 추가
        private void btnBMAdd_Click(object sender, RoutedEventArgs e)
        {
            if (dgdThird.SelectedItems.Count > 0)
            {
                for (int i = 0; i < dgdThird.SelectedItems.Count; i++)
                {
                    var GetThird = dgdThird.SelectedItems[i] as ThirdMenu;
                    var mvm = GetThird.SecondValue as MenuViewModel;
                    //ListBoxItem lbxItem = new ListBoxItem();
                    //lbxItem.Tag = GetThird.SecondValue;
                    //lbxItem.Content = GetThird.ThirdValue;

                    //string[] strArray = { GetThird.SecondValue, GetThird.ThirdValue };
                    dgdBoxBookMarkMenu.Items.Add(GetThird);
                    listBMMenu.Add(mvm);
                }

                SetThirdMenu();
            }
        }

        //즐겨찾기 메뉴에서 제외
        private void btnBMException_Click(object sender, RoutedEventArgs e)
        {
            bool flag = true;

            if (dgdBoxBookMarkMenu.SelectedIndex != -1)
            {
                for (int i = 0; i < dgdBoxBookMarkMenu.SelectedItems.Count; i++)
                {
                    var lbxItem = dgdBoxBookMarkMenu.SelectedItems[i] as ThirdMenu;

                    //string[] strArray = { lbxItem.SecondValue, lbxItem.ThirdValue };
                    //listBMMenu.Remove(strArray);
                    for (int j = 0; j < listBMMenu.Count; j++)
                    {
                        if (lbxItem.ThirdValue.Equals(listBMMenu[j].Menu.Trim()))
                        {
                            listBMMenu.RemoveAt(j);
                        }
                    }
                }
            }

            dgdBoxBookMarkMenu.Items.Clear();

            for (int i = 0; i < listBMMenu.Count; i++)
            {
                var lbxItem = new ThirdMenu()
                {
                    SecondValue = listBMMenu[i],
                    ThirdValue = listBMMenu[i].Menu.Trim()
                };

                //ListBoxItem lbxItem = new ListBoxItem();
                //lbxItem.Tag = listBMMenu[i][0];
                //lbxItem.Content = listBMMenu[i][1];

                dgdBoxBookMarkMenu.Items.Add(lbxItem);
            }

            SetFirstMenu();
        }

        //
        private void btnOK_Click(object sender, RoutedEventArgs e)
        {
            listBMMenu.Clear();

            foreach (ThirdMenu item in dgdBoxBookMarkMenu.Items)
            {
                //string[] strArray = { item.SecondValue, item.ThirdValue };
                listBMMenu.Add(item.SecondValue);
            }

            DialogResult = true;
        }

        //
        private void btnCancel_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
        }
    }

    class FirstMenu : BaseView
    {
        public string FirstValue { get; set; }
    }

    class SecondMenu : BaseView
    {
        public string FirstValue { get; set; }
        public string SecondValue { get; set; }
    }

    class ThirdMenu : BaseView
    {
        public string FirstValue { get; set; }
        public MenuViewModel SecondValue { get; set; }
        public string ThirdValue { get; set; }
    }
}

using System;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Media;
using System.Windows.Media.Animation;
using System.Windows.Threading;

namespace WizMes_WellMade.PopUp
{
    /// <summary>
    /// LoginPage.xaml에 대한 상호 작용 논리
    /// </summary>
    public partial class Loading : Window, IDisposable
    {
        public Action Function { get; set; }

        // 엑셀때 사용되는지 여부
        bool excel = false;

        // 로딩 애니메이션 속도
        double speed = 2;

        public Loading(Action function)
        {
            InitializeComponent();

            Function = function ?? throw new ArgumentNullException();
        }

        public Loading(string str, Action function)
        {
            InitializeComponent();

            Function = function ?? throw new ArgumentNullException();

            if (str.ToUpper().Trim().Equals("EXCEL"))
            {
                speed = 5;
                excel = true;
            }
        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            AnimationStart(speed);

            //Task.Factory.StartNew(SetAction).ContinueWith(t => { this.Close(); }, TaskScheduler.FromCurrentSynchronizationContext());
            Task.Factory.StartNew(SetAction).ContinueWith(t => { this.DialogResult = true; }, TaskScheduler.FromCurrentSynchronizationContext());

        }

        private void SetAction()
        {
            // 엑셀이라면
            if (excel)
            {
                Function();
                return;
            }

            // 너무 짧으면 안되니까..
            // 로딩화면이 보일 수 있도록 조금 추가하도록 하자.
            Thread.Sleep(300);

            Dispatcher.Invoke(DispatcherPriority.Normal, new Action(delegate
            {
                Function();
            }));
        }

        // 이건 뭐지..
        public void Dispose() { }

        // 이미지 돌리기
        private void AnimationStart(double speed)
        {
            DoubleAnimation dba1 = new DoubleAnimation();
            dba1.From = 0;
            dba1.To = 360;

            dba1.Duration = new Duration(TimeSpan.FromSeconds(speed));
            dba1.RepeatBehavior = RepeatBehavior.Forever;

            RotateTransform rt = new RotateTransform();
            ImgLoading.RenderTransform = rt;
            rt.CenterX += ImgLoading.Width / 2;
            rt.CenterY += ImgLoading.Height / 2;

            rt.BeginAnimation(RotateTransform.AngleProperty, dba1);
        }


        #region 기타 메서드

        // 천마리 콤마, 소수점 버리기
        private string stringFormatN0(object obj)
        {
            return string.Format("{0:N0}", obj);
        }

        // DateTime 으로 변환 가능한지 체크
        private bool CheckConvertDateTime(string str)
        {
            bool flag = false;

            DateTime chkDt;

            if (!str.Trim().Equals(""))
            {
                if (str.Length == 8)
                {
                    str = DatePickerFormat(str);

                    if (DateTime.TryParse(str, out chkDt) == true)
                    {
                        flag = true;
                        return flag;
                    }
                }
                else
                {
                    if (DateTime.TryParse(str, out chkDt) == true)
                    {
                        flag = true;
                        return flag;
                    }
                }
            }

            return flag;
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


        #endregion // 기타 메서드
    }
}

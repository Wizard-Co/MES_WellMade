using System;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;

namespace WizMes_WellMade
{
    public class Attached
    {
        public static readonly DependencyProperty OnKeyDownActionProperty = DependencyProperty.RegisterAttached(
        "OnKeyDownAction", typeof(Action<object>), typeof(Attached), new PropertyMetadata(default(Action<object>)));

        public static void SetOnKeyDownAction(DependencyObject element, Action<object> value)
        {
            element.SetValue(OnKeyDownActionProperty, value);
        }

        public static Action<object> GetOnKeyDownAction(DependencyObject element)
        {
            return (Action<object>)element.GetValue(OnKeyDownActionProperty);
        }

        public static readonly DependencyProperty IsReactsOnKeyDownProperty = DependencyProperty.RegisterAttached(
            "IsReactsOnKeyDown", typeof(bool), typeof(Attached), new PropertyMetadata(default(bool), IsReactsOnKeyDownPropertyChangedCallback));

        private static void IsReactsOnKeyDownPropertyChangedCallback(DependencyObject sender, DependencyPropertyChangedEventArgs args)
        {
            var val = (bool)args.NewValue;
            var cell = sender as DataGridCell;
            if (cell == null)
                return;
            if (val == false)
            {
                cell.KeyDown -= CellOnKeyDown;
            }
            else
            {
                cell.KeyDown += CellOnKeyDown;
            }

        }

        private static void CellOnKeyDown(object sender, KeyEventArgs keyEventArgs)
        {
            var cell = sender as DataGridCell;
            if (cell == null)
                return;
            var action = cell.GetValue(OnKeyDownActionProperty) as Action<object>;
            if (action == null) return;
            action(keyEventArgs);
        }

        public static void SetIsReactsOnKeyDown(DependencyObject element, bool value)
        {
            element.SetValue(IsReactsOnKeyDownProperty, value);
        }

        public static bool GetIsReactsOnKeyDown(DependencyObject element)
        {
            return (bool)element.GetValue(IsReactsOnKeyDownProperty);
        }
    }
}

﻿//    This file is part of QTTabBar, a shell extension for Microsoft
//    Windows Explorer.
//    Copyright (C) 2007-2010  Quizo, Paul Accisano
//
//    QTTabBar is free software: you can redistribute it and/or modify
//    it under the terms of the GNU General Public License as published by
//    the Free Software Foundation, either version 3 of the License, or
//    (at your option) any later version.
//
//    QTTabBar is distributed in the hope that it will be useful,
//    but WITHOUT ANY WARRANTY; without even the implied warranty of
//    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
//    GNU General Public License for more details.
//
//    You should have received a copy of the GNU General Public License
//    along with QTTabBar.  If not, see <http://www.gnu.org/licenses/>.

using System;
using System.ComponentModel;
using System.Globalization;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Markup;

namespace QTTabBarLib {

    // Universal converters.  This allows us to keep the real converters as
    // nested classes, so we can avoid cluttering up the namespace
    internal class Converter : GenericConverter<IValueConverter>, IValueConverter {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture) {
            CreateConverter();
            return converter.Convert(value, targetType, parameter, culture);
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture) {
            CreateConverter();
            return converter.ConvertBack(value, targetType, parameter, culture);
        }
    }

    internal class MultiConverter : GenericConverter<IMultiValueConverter>, IMultiValueConverter {
        public object Convert(object[] values, Type targetType, object parameter, CultureInfo culture) {
            CreateConverter();
            return converter.Convert(values, targetType, parameter, culture);
        }

        public object[] ConvertBack(object value, Type[] targetTypes, object parameter, CultureInfo culture) {
            CreateConverter();
            return converter.ConvertBack(value, targetTypes, parameter, culture);
        }
    }

    internal class GenericConverter<T> where T : class {
        protected T converter;
        private Type type;
        public Type Type {
            get { return type; }
            set {
                if(value == type) return;
                if(value.GetInterface(typeof(T).Name) == null) {
                    throw new ArgumentException(string.Format("Type {0} doesn't implement {1}", value.FullName, typeof(T).Name), "value");
                }
                type = value;
                converter = null;
            }
        }

        protected void CreateConverter() {
            if(converter != null) return;
            if(type == null) throw new InvalidOperationException("Converter type is not defined");
            converter = Activator.CreateInstance(type) as T;
        }
    }

    // Overloaded RadioButton class to work around .NET 3.5's horribly HORRIBLY
    // bugged RadioButton data binding.
    public class RadioButtonEx : RadioButton {

        private bool bIsChanging;

        public RadioButtonEx() {
            Checked += RadioButtonExtended_Checked;
            Unchecked += RadioButtonExtended_Unchecked;
        }

        void RadioButtonExtended_Unchecked(object sender, RoutedEventArgs e) {
            if(!bIsChanging) IsCheckedReal = false;
        }

        void RadioButtonExtended_Checked(object sender, RoutedEventArgs e) {
            if(!bIsChanging) IsCheckedReal = true;
        }

        public bool? IsCheckedReal {
            get { return (bool?)GetValue(IsCheckedRealProperty); }
            set { SetValue(IsCheckedRealProperty, value); }
        }

        // Using a DependencyProperty as the backing store for IsCheckedReal.
        // This enables animation, styling, binding, etc...
        public static readonly DependencyProperty IsCheckedRealProperty =
                DependencyProperty.Register("IsCheckedReal", typeof(bool?), typeof(RadioButtonEx),
                new FrameworkPropertyMetadata(false, FrameworkPropertyMetadataOptions.Journal |
                FrameworkPropertyMetadataOptions.BindsTwoWayByDefault,
                OnIsCheckedRealChanged));

        private static void OnIsCheckedRealChanged(DependencyObject d, DependencyPropertyChangedEventArgs e) {
            RadioButtonEx rbx = ((RadioButtonEx)d);
            rbx.bIsChanging = true;
            rbx.IsChecked = (bool?)e.NewValue;
            rbx.bIsChanging = false;
        }
    }

    // This allows us to use ListViews inside a ScrollViewer, and set the minimum height of the ListView.
    public class RestrictDesiredSize : Decorator {
        Size lastArrangeSize = new Size(double.PositiveInfinity, double.PositiveInfinity);
        protected override Size MeasureOverride(Size constraint) {
            base.MeasureOverride(new Size(Math.Min(lastArrangeSize.Width, constraint.Width),
                                          Math.Min(lastArrangeSize.Height, constraint.Height)));
            return new Size(0, 0);
        }

        protected override Size ArrangeOverride(Size arrangeSize) {
            if(lastArrangeSize != arrangeSize) {
                lastArrangeSize = arrangeSize;
                base.MeasureOverride(arrangeSize);
            }
            return base.ArrangeOverride(arrangeSize);
        }
    }

    public class BindableRun : Run {
        public static readonly DependencyProperty TextProperty =
                DependencyProperty.Register("Text", typeof(string),
                typeof(BindableRun), new PropertyMetadata(OnTextChanged));

        private static void OnTextChanged(DependencyObject d, DependencyPropertyChangedEventArgs e) {
            ((Run)d).Text = (string)e.NewValue;
        }

        public new string Text {
            get { return (string)GetValue(TextProperty); }
            set { SetValue(TextProperty, value); }
        }
    }

    [MarkupExtensionReturnType(typeof(string))]
    class Resx : MarkupExtension {
       
        // Use weak events to avoid memory leaks.
        private class ResxEventManager : WeakEventManager {
            private static EventHandler mHandler = (s, e) => CurrentManager.DeliverEvent(s, e);
            private static ResxEventManager CurrentManager {
                get {
                    var manager = GetCurrentManager(typeof(ResxEventManager)) as ResxEventManager;
                    if(manager == null) {
                        manager = new ResxEventManager();
                        SetCurrentManager(typeof(ResxEventManager), manager);
                    }
                    return manager;
                }
            }
            public static void AddListener(IWeakEventListener listener) {
                CurrentManager.ProtectedAddListener(typeof(Resx), listener);
            }
            protected override void StartListening(object source) {
                OnUpdate += mHandler;
            }
            protected override void StopListening(object source) {
                OnUpdate -= mHandler;
            }           
        }

        private class ResxListener : INotifyPropertyChanged, IWeakEventListener {
            public event PropertyChangedEventHandler PropertyChanged;
            // ReSharper disable MemberCanBePrivate.Local
            // ReSharper disable UnusedAutoPropertyAccessor.Local
            public string Value { get; set; }
            // ReSharper restore UnusedAutoPropertyAccessor.Local
            // ReSharper restore MemberCanBePrivate.Local
            private Resx parent;

            public ResxListener(Resx parent) {
                this.parent = parent;
                Value = parent.GetValue();
            }

            public bool ReceiveWeakEvent(Type managerType, object sender, EventArgs e) {
                Value = parent.GetValue();
                return true;
            }
        }

        private static event EventHandler OnUpdate; 
        private static bool debug;
        public static bool DebugMode {
            get { return debug; }
            set { debug = value; UpdateAll(); }
        }
        public static void UpdateAll() {
            if(OnUpdate != null) {
                OnUpdate(typeof(Resx), new EventArgs());
            }
        }

        public static readonly DependencyProperty ParamProperty = DependencyProperty.RegisterAttached(
                "Param", typeof(string), typeof(Resx), new PropertyMetadata(null));

        public static void SetParam(UIElement element, string value) {
            element.SetValue(ParamProperty, value);
        }
        public static string GetParam(UIElement element) {
            return (string)element.GetValue(ParamProperty);
        }

        public Resx() { }
        public Resx(string key, int index = 0) {
            Key = key;
            Index = index;
        }
        public string Key { get; set; }
        public int Index { get; set; }
        private DependencyObject targetObject;

        public override object ProvideValue(IServiceProvider serviceProvider) {            
            IProvideValueTarget target = serviceProvider.GetService(typeof(IProvideValueTarget)) as IProvideValueTarget;
            if (target != null) {
                targetObject = target.TargetObject as DependencyObject;
                if(targetObject == null && !(target.TargetObject is Setter)) return this;
            }
            ResxListener listener = new ResxListener(this);
            ResxEventManager.AddListener(listener);
            return (new Binding("Value") { Source = listener }).ProvideValue(serviceProvider);
        }

        private string GetValue() {
            if(DebugMode) return Key + "[" + Index + "]";
            string[] res;
            if(!QTUtility.TextResourcesDic.TryGetValue(Key, out res) || Index >= res.Length) return "";
            string ret = res[Index];
            string param = targetObject == null
                    ? null
                    : (string)targetObject.GetValue(ParamProperty);
            if(param != null) ret = string.Format(ret, param);
            return ret.Replace("&", "_");
        }
    }
}

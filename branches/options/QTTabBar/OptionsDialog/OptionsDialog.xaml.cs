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
using System.Collections;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Collections.Specialized;
using System.Globalization;
using System.Linq;
using System.Threading;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Input;
using System.Windows.Interop;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using QTTabBarLib.Interop;
using Font = System.Drawing.Font;
using Bitmap = System.Drawing.Bitmap;
using Keys = System.Windows.Forms.Keys;

namespace QTTabBarLib {
    /// <summary>
    /// Interaction logic for OptionsDialog.xaml
    /// </summary>
    internal partial class OptionsDialog : IDisposable {
        private static OptionsDialog instance;
        private static Thread instanceThread;
        private static Thread launchingThread;
        private Config WorkingConfig;
        
        #region ---------- Static Methods ----------

        public static void Open() {
            // TODO: Primary process only
            lock(typeof(OptionsDialog)) {
                // Prevent reentrant calls that might happen during the Wait call below.
                if(launchingThread == Thread.CurrentThread) return;
                try {
                    launchingThread = Thread.CurrentThread;

                    if(instance == null) {
                        instanceThread = new Thread(ThreadEntry);
                        instanceThread.SetApartmentState(ApartmentState.STA);
                        lock(instanceThread) {
                            instanceThread.Start();
                            // Don't return until we know that instance is set.
                            Monitor.Wait(instanceThread);
                        }
                    }
                    else {
                        instance.Dispatcher.Invoke(new Action(() => {
                            if(instance.WindowState == WindowState.Minimized) {
                                instance.WindowState = WindowState.Normal;
                            }
                            else {
                                instance.Activate();
                            }
                        }));
                    }
                }
                finally {
                    launchingThread = null;
                }
            }
        }

        public static void ForceClose() {
            lock(typeof(OptionsDialog)) {
                if(instance != null) {
                    instance.Dispatcher.Invoke(new Action(() => instance.Close()));
                }
            }
        }

        private static void ThreadEntry() {
            using(instance = new OptionsDialog()) {
                lock(instanceThread) {
                    Monitor.Pulse(instanceThread);
                }
                instance.ShowDialog();
            }
            lock(typeof(OptionsDialog)) {
                instance = null;
            }
        }

        #endregion

        private OptionsDialog() {
            InitializeComponent();

            tabbedPanel.ItemsSource = new OptionsDialogTab[] {
                new Options01_Window        { TabLabel = "Window"       },
                new Options02_Tabs          { TabLabel = "Tabs"         },
                new Options03_Tweaks        { TabLabel = "Tweaks"       },
                new Options04_Tooltips      { TabLabel = "Tooltips"     },
                new Options05_General       { TabLabel = "General"      },
                new Options06_Appearance    { TabLabel = "Appearance"   },
                new Options07_Mouse         { TabLabel = "Mouse"        },
                new Options08_Keys          { TabLabel = "Keys"         },
                new Options09_Groups        { TabLabel = "Groups"       },
                new Options10_Apps          { TabLabel = "Applications" },
                new Options11_ButtonBar     { TabLabel = "Button Bar"   },
                new Options12_Plugins       { TabLabel = "Plugins"      },
                new Options13_Language      { TabLabel = "Language"     },
                new Options14_About         { TabLabel = "About"        },
            };

            WorkingConfig = QTUtility2.DeepClone(ConfigManager.LoadedConfig);
            foreach(OptionsDialogTab tab in tabbedPanel.Items) {
                tab.WorkingConfig = WorkingConfig;
                if(tab is IHotkeyContainer) {
                    ((IHotkeyContainer)tab).NewHotkeyRequested += ProcessNewHotkey;
                }
                tab.InitializeConfig();
            }
        }

        public void Dispose() {
            // TODO
        }

        private void UpdateOptions() {
            // A small caveat in my brilliant plan to separate the options tabs into separate files:
            // The plugin tab has to be committed first, before the others.
            OptionsDialogTab pluginTab = tabbedPanel.Items.OfType<Options12_Plugins>().First();
            pluginTab.CommitConfig();

            foreach(OptionsDialogTab tab in tabbedPanel.Items) {
                if(tab != pluginTab) tab.CommitConfig();
            }

            // todo: redo this crap...
            bool fButtonBarNeedsRefresh = Config.BBar.LargeButtons != WorkingConfig.bbar.LargeButtons;
            ConfigManager.LoadedConfig = QTUtility2.DeepClone(WorkingConfig);
            ConfigManager.WriteConfig();
            QTTabBarClass tabBar = InstanceManager.CurrentTabBar;
            if(tabBar != null) {
                tabBar.Invoke(new Action(tabBar.RefreshOptions));
            }
            QTButtonBar.BroadcastConfigChanged(fButtonBarNeedsRefresh);
        }

        private void CategoryListBoxItem_PreviewMouseRightButtonDown(object sender, MouseButtonEventArgs e) {
            lstCategories.Focus();
            e.Handled = true;
        }

        private void CategoryListBoxItem_PreviewMouseDown(object sender, MouseButtonEventArgs e) {
            ((ListBoxItem)sender).Focus();
            ((ListBoxItem)sender).IsSelected = true;
            e.Handled = true;
        }

        private void btnResetPage_Click(object sender, RoutedEventArgs e) {
            ((OptionsDialogTab)tabbedPanel.SelectedItem).ResetConfig();
        }

        private void btnResetAll_Click(object sender, RoutedEventArgs e) {
            foreach(OptionsDialogTab tab in tabbedPanel.Items) {
                tab.ResetConfig();
            }
        }

        private void btnOK_Click(object sender, RoutedEventArgs e) {
            UpdateOptions();
            Close();
        }

        private void btnCancel_Click(object sender, RoutedEventArgs e) {
            Close();
        }

        private void btnApply_Click(object sender, RoutedEventArgs e) {
            UpdateOptions();
            foreach(OptionsDialogTab tab in tabbedPanel.Items) {
                tab.InitializeConfig();
            }
        }

        private bool ProcessNewHotkey(KeyEventArgs e, Keys current, out Keys finalKey) {
            finalKey = Keys.None;
            Key wpfKey = (e.Key == Key.System ? e.SystemKey : e.Key);
            ModifierKeys wpfModKeys = Keyboard.Modifiers;

            // Ignore modifier keys.
            if(wpfKey == Key.LeftShift || wpfKey == Key.RightShift
                    || wpfKey == Key.LeftCtrl || wpfKey == Key.RightCtrl
                    || wpfKey == Key.LeftAlt || wpfKey == Key.RightAlt
                    || wpfKey == Key.LWin || wpfKey == Key.RWin) {
                return false;
            }

            Keys hotkey = (Keys)KeyInterop.VirtualKeyFromKey(wpfKey);
            if(hotkey == Keys.Escape) {
                // Escape = clear
                return true;
            }

            // Urgh, so many conversions between WPF and WinForms...
            Keys modkey = Keys.None;
            if((wpfModKeys & ModifierKeys.Alt) != 0)        modkey |= Keys.Alt;
            if((wpfModKeys & ModifierKeys.Control) != 0)    modkey |= Keys.Control;
            if((wpfModKeys & ModifierKeys.Shift) != 0)      modkey |= Keys.Shift;

            // don't allow keystrokes without at least one modifier key
            if(modkey == Keys.None) {
                return false;
            }

            modkey |= hotkey;
            if(modkey == current) {
                // trying to assign the same hotkey
                return false;
            }

            // keys not allowed even with any modifier keys 
            switch(hotkey) {
                case Keys.None:
                case Keys.Enter:
                case Keys.ControlKey:
                case Keys.ShiftKey:
                case Keys.Menu:				// Alt itself
                case Keys.NumLock:
                case Keys.ProcessKey:
                case Keys.IMEConvert:
                case Keys.IMENonconvert:
                case Keys.KanaMode:
                case Keys.KanjiMode:
                    return false;
            }

            // keys not allowed as one key
            switch(modkey) {
                case Keys.LWin:
                case Keys.RWin:
                case Keys.Delete:
                case Keys.Apps:
                case Keys.Tab:
                case Keys.Left:
                case Keys.Up:
                case Keys.Right:
                case Keys.Down:
                    return false;
            }

            // invalid key combinations 
            switch(modkey) {
                case Keys.Control | Keys.C:
                case Keys.Control | Keys.A:
                case Keys.Control | Keys.Z:
                case Keys.Control | Keys.V:
                case Keys.Control | Keys.X:
                case Keys.Alt | Keys.Left:
                case Keys.Alt | Keys.Right:
                case Keys.Alt | Keys.F4:
                    System.Media.SystemSounds.Hand.Play();
                    return false;
            }

            // check for key conflicts
            const string Conflict = "This key is already assigned to:\n{0}\n\nReassign?";
            const string MsgTitle = "Keystroke conflict";
            IHotkeyEntry conflictingEntry = tabbedPanel.Items
                    .OfType<IHotkeyContainer>()
                    .SelectMany(hc => hc.GetHotkeyEntries())
                    .FirstOrDefault(entry => entry.ShortcutKey == modkey);
            if(conflictingEntry != null) {
                if(MessageBoxResult.OK != MessageBox.Show(
                        string.Format(Conflict, conflictingEntry.KeyActionText),
                        MsgTitle,
                        MessageBoxButton.OKCancel,
                        MessageBoxImage.Warning)) {
                    return false;
                }
                conflictingEntry.ShortcutKey = Keys.None;              
            }

            finalKey = modkey;
            return true;
        }

        #region ---------- Converters ----------

        // Inverts the value of a boolean
        internal class BoolInverterConverter : IValueConverter {
            public object Convert(object value, Type targetType, object parameter, CultureInfo culture) {
                return value is bool ? !(bool)value : value;
            }

            public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture) {
                return value is bool ? !(bool)value : value;
            }
        }

        // Converts between booleans and one using logical and.
        internal class LogicalAndMultiConverter : IMultiValueConverter {
            public object Convert(object[] values, Type targetType, object parameter, CultureInfo culture) {
                return values.All(b => b is bool && (bool)b);
            }

            public object[] ConvertBack(object value, Type[] targetTypes, object parameter, CultureInfo culture) {
                return new object[] { value };
            }
        }

        // Converts between many booleans and a string by StringJoining them.
        internal class BoolJoinMultiConverter : IMultiValueConverter {
            public object Convert(object[] values, Type targetType, object parameter, CultureInfo culture) {
                return values.StringJoin(",");
            }

            public object[] ConvertBack(object value, Type[] targetTypes, object parameter, CultureInfo culture) {
                return ((string)value).Split(',').Select(s => (object)bool.Parse(s)).ToArray();
            }
        }

        // Converts between a boolean and a string by comparing the string to the 
        // passed parameter.
        [ValueConversion(typeof(string), typeof(bool))]
        internal class StringEqualityConverter : IValueConverter {
            public object Convert(object value, Type targetType, object parameter, CultureInfo culture) {
                return (string)parameter == (string)value;
            }

            public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture) {
                return (bool)value ? parameter : Binding.DoNothing;
            }
        }

        // Converts Bitmaps to ImageSources.
        [ValueConversion(typeof(Bitmap), typeof(ImageSource))]
        internal class BitmapToImageSourceConverter : IValueConverter {
            public object Convert(object value, Type targetType, object parameter, CultureInfo culture) {
                if(value == null || !(value is Bitmap)) return null;
                IntPtr hBitmap = ((Bitmap)value).GetHbitmap();
                try {
                    return Imaging.CreateBitmapSourceFromHBitmap(hBitmap, IntPtr.Zero,
                            Int32Rect.Empty, BitmapSizeOptions.FromEmptyOptions());
                }
                finally {
                    PInvoke.DeleteObject(hBitmap);
                }
            }

            public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture) {
                throw new NotSupportedException();
            }
        }

        // Converts between Colors and Brushes
        [ValueConversion(typeof(System.Drawing.Color), typeof(Brush))]
        internal class ColorToBrushConverter : IValueConverter {
            public object Convert(object value, Type targetType, object parameter, CultureInfo culture) {
                var c = (System.Drawing.Color)(value ?? System.Drawing.Color.Red);
                return new SolidColorBrush(Color.FromArgb(c.A, c.R, c.G, c.B));
            }

            public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture) {
                throw new NotSupportedException();
            }
        }

        // Converts between Fonts and strings.
        [ValueConversion(typeof(Font), typeof(string))]
        internal class FontStringConverter : IValueConverter {
            public object Convert(object value, Type targetType, object parameter, CultureInfo culture) {
                if(value == null) return "";
                Font font = (Font)value;
                return string.Format("{0}, {1} pt", font.Name, Math.Round(font.SizeInPoints));
            }

            public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture) {
                throw new NotSupportedException();
            }
        }

        // Converts an object to its class name.
        // You can create ObjectToTypeConverter instead of this,
        // but VS2010 WPF Designer would refuse an expression like {x:Type SomeClass+NestedClass}
        [ValueConversion(typeof(object), typeof(string))]
        internal class ObjectToClassNameConverter : IValueConverter {
            public object Convert(object value, Type targetType, object parameter, CultureInfo culture) {
                return value == null ? null : value.GetType().Name;
            }

            public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture) {
                throw new NotSupportedException();
            }
        }

        #endregion
    }

    internal interface IHotkeyEntry {
        Keys ShortcutKey { get; set; }
        string KeyActionText { get; }
    }

    internal delegate bool NewHotkeyRequestedHandler(KeyEventArgs keyEvent, Keys currentKey, out Keys newKey);
    internal interface IHotkeyContainer {
        IEnumerable<IHotkeyEntry> GetHotkeyEntries();
        event NewHotkeyRequestedHandler NewHotkeyRequested;
    }

    /// <summary>
    /// The base class for the tab pages of the OptionsDialog.
    /// Contains a few things common to more than one page.
    /// </summary>
    internal abstract class OptionsDialogTab : UserControl {
        public static readonly DependencyProperty WorkingConfigProperty =
                DependencyProperty.Register("WorkingConfig", typeof(Config), typeof(OptionsDialogTab),
                new FrameworkPropertyMetadata(null, FrameworkPropertyMetadataOptions.BindsTwoWayByDefault));

        public Config WorkingConfig {
            get { return (Config)GetValue(WorkingConfigProperty); }
            set { SetValue(WorkingConfigProperty, value); }
        }

        internal PluginManager pluginManager;
        protected OptionsDialogTab() {
            // TODO: Find a way to get rid of this.
            QTTabBarClass tabBar = InstanceManager.CurrentTabBar;
            if(tabBar != null) {
                pluginManager = tabBar.GetPluginManager();
            }
        }

        // This is the string that will be displayed in the category list.
        public string TabLabel { get; set; }

        // Called when the options dialog is first shown, and when the user clicks Apply (after commit)
        public abstract void InitializeConfig();

        // Called when the user clicks the Reset buttons.
        public abstract void ResetConfig();

        // Called when the user clicks Apply or OK.
        public abstract void CommitConfig();
        

        #region ---------- Interfaces / Helper Classes ----------

        // Interface for Binding Classes that have some editable component
        protected interface IEditableEntry {
            bool IsEditing { get; set; }
        }

        // Interface for Binding Classes that belong to a ParentedCollection list.
        protected interface IChildItem {
            IList ParentList { get; set; }
        }

        // Interface for TreeView items, to control selectedness and expandedness.
        protected interface ITreeViewItem : IChildItem {
            bool IsSelected { get; set; }
            bool IsExpanded { get; set; }
        }

        // A subclass of ObservableCollection that allows the parent list to be accessed from its children.
        protected sealed class ParentedCollection<TChild> : ObservableCollection<TChild>
            where TChild : class, IChildItem {
            public ParentedCollection(IEnumerable<TChild> collection = null) {
                if(collection != null) {
                    foreach(TChild child in collection) {
                        child.ParentList = this;
                        Add(child);
                    }
                }
                CollectionChanged += ParentedCollection_CollectionChanged;
            }

            private void ParentedCollection_CollectionChanged(object sender, NotifyCollectionChangedEventArgs e) {
                if(e.NewItems != null) {
                    foreach(TChild newItem in e.NewItems) {
                        newItem.ParentList = this;
                    }
                }
                if(e.OldItems != null) {
                    foreach(TChild oldItem in e.OldItems) {
                        oldItem.ParentList = null;
                    }
                }
            }
        }

        // Simple overloaded ColorDialog to start with the color picker active
        protected sealed class ColorDialogEx : System.Windows.Forms.ColorDialog {
            protected override int Options {
                get {
                    return (base.Options | 2);
                }
            }
        }

        #endregion

        // Common Font Chooser button click handler.
        protected void btnFontChoose_Click(object sender, RoutedEventArgs e) {
            var button = (Button)sender;
            try {
                using(var dialog = new System.Windows.Forms.FontDialog()) {
                    dialog.Font = (Font)button.Tag;
                    dialog.ShowEffects = false;
                    dialog.AllowVerticalFonts = false;
                    if(System.Windows.Forms.DialogResult.OK == dialog.ShowDialog()) {
                        button.Tag = dialog.Font;
                    }
                }
            }
            catch { }
        }

        // Utility method to move nodes up and down in a TreeView.
        protected static void UpDownOnTreeView(TreeView tvw, bool up) {
            ITreeViewItem val = tvw.SelectedItem as ITreeViewItem;
            if(val == null) return;
            IList list = val.ParentList;
            int index = list.IndexOf(val);
            if(index == -1) {
                return;
            }
            if(up) {
                if(index == 0) {
                    return;
                }
            }
            else {
                if(index == list.Count - 1) {
                    return;
                }
            }

            bool expanded = val.IsExpanded;
            list.RemoveAt(index);
            list.Insert(index + (up ? -1 : 1), val);
            val.IsExpanded = expanded;
            val.IsSelected = true;
        }
    }
}
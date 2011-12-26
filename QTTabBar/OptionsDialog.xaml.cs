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
using System.ComponentModel;
using System.Globalization;
using System.Linq;
using System.Reflection;
using System.Threading;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Data;
using System.Windows.Input;
using System.Windows.Interop;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;
using QTPlugin;
using QTTabBarLib.Interop;
using Image = System.Drawing.Image;
using Size = System.Drawing.Size;
using Font = System.Drawing.Font;
using Bitmap = System.Drawing.Bitmap;
using OpenFileDialog = System.Windows.Forms.OpenFileDialog;
using Keys = System.Windows.Forms.Keys;

namespace QTTabBarLib {
    /// <summary>
    /// Interaction logic for OptionsDialog.xaml
    /// </summary>
    public partial class OptionsDialog : IDisposable {

        private static OptionsDialog instance;
        private static Thread instanceThread;
        private static Thread launchingThread = null;
        public Config workingConfig { get; set; }

        // Button bar stuff
        private ImageStrip imageStripLarge;
        private ImageStrip imageStripSmall;
        private string[] ButtonItemsDisplayName;
        private ObservableCollection<ButtonEntry> ButtonPool;
        private ObservableCollection<ButtonEntry> CurrentButtons;

        // Plugin stuff
        private PluginManager pluginManager;
        private ObservableCollection<PluginEntry> CurrentPlugins;

        // Mouse stuff
        private ObservableCollection<MouseEntry> MouseBindings;

        // Tooltips stuff
        private ObservableCollection<FileTypeEntry> TextFileTypes;
        private ObservableCollection<FileTypeEntry> MediaFileTypes;

        // Groups/Apps stuff
        private ParentedCollection<GroupEntry> CurrentGroups;
        private ParentedCollection<AppEntry> CurrentApps;

        private List<HotkeyEntry> HotkeyEntries;

        // todo: localize all of these
        #region ToLocalize

        private readonly static Dictionary<StretchMode, string> StretchModeItems
                = new Dictionary<StretchMode, string> {
            {StretchMode.Full,  "Stretch"},
            {StretchMode.Real,  "True size"},
            {StretchMode.Tile,  "Tile"},
        };
        private readonly static Dictionary<TabPos, string> NewTabPosItems
                = new Dictionary<TabPos, string> {
            {TabPos.Left,       "Left of the current tab"   },
            {TabPos.Right,      "Right of the current tab"  },
            {TabPos.Leftmost,   "In the leftmost position"  },
            {TabPos.Rightmost,  "In the rightmost position" },
        };
        private readonly static Dictionary<TabPos, string> NextAfterCloseItems
                = new Dictionary<TabPos, string> {
            {TabPos.Left,       "Tab to the left"   },
            {TabPos.Right,      "Tab to the right"  },
            {TabPos.Leftmost,   "Leftmost tab"      },
            {TabPos.Rightmost,  "Rightmost tab"     },
            {TabPos.LastActive, "Last activated tab"},
        };
        private readonly static Dictionary<MouseTarget, string> MouseTargetItems
                = new Dictionary<MouseTarget, string> {
            {MouseTarget.Anywhere,          "Anywhere"              },
            {MouseTarget.Tab,               "Tab"                   },
            {MouseTarget.TabBarBackground,  "Tab Bar Background"    },
            {MouseTarget.FolderLink,        "Folder Link"           },
            {MouseTarget.ExplorerItem,      "File or Folder"        },
            {MouseTarget.ExplorerBackground,"Explorer Background"   },
        };
        private readonly static Dictionary<BindAction, string> MouseActionItems
                = new Dictionary<BindAction, string> {
            {BindAction.GoBack,                 "Go back"},
            {BindAction.GoForward,              "Go forward"},
            {BindAction.GoFirst,                "Go to back to start"},
            {BindAction.GoLast,                 "Go forward to end"},
            {BindAction.NextTab,                "Next tab"},
            {BindAction.PreviousTab,            "Previous tab"},
            {BindAction.FirstTab,               "First tab"},
            {BindAction.LastTab,                "Last tab"},
            {BindAction.CloseCurrent,           "Close current tab"},
            {BindAction.CloseAllButCurrent,     "Close all tabs except current"},
            {BindAction.CloseLeft,              "Close tabs to the left of current"},
            {BindAction.CloseRight,             "Close tabs to the right of current"},
            {BindAction.CloseWindow,            "Close window"},
            {BindAction.RestoreLastClosed,      "Restore last closed tab"},
            {BindAction.CloneCurrent,           "Clone current tab"},
            {BindAction.TearOffCurrent,         "Open current tab in new window"},
            {BindAction.LockCurrent,            "Lock current tab"},
            {BindAction.LockAll,                "Lock all tabs"},
            {BindAction.BrowseFolder,           "Browse for folder"},
            {BindAction.CreateNewGroup,         "Create new group"},
            {BindAction.ShowOptions,            "Show options"},
            {BindAction.ShowToolbarMenu,        "Show Toolbar menu"},
            {BindAction.ShowTabMenuCurrent,     "Show current tab context menu"},
            {BindAction.ShowGroupMenu,          "Show Button Bar Group menu"},
            {BindAction.ShowRecentFolderMenu,   "Show Button Bar Recently Closed menu" },
            {BindAction.ShowUserAppsMenu,       "Show Button Bar Apps menu"},
            {BindAction.ToggleMenuBar,          "Toggle Explorer Menu Bar"},
            {BindAction.CopySelectedPaths,      "Copy paths of selected files"},
            {BindAction.CopySelectedNames,      "Copy names of selected files"},
            {BindAction.ChecksumSelected,       "View checksums selected files"},
            {BindAction.ToggleTopMost,          "Toggle always on top"},
            {BindAction.TransparencyPlus,       "Transparency +"},
            {BindAction.TransparencyMinus,      "Transparency -"},
            {BindAction.FocusFileList,          "Focus file list"},
            {BindAction.FocusSearchBarReal,     "Focus Explorer search box"},
            {BindAction.FocusSearchBarBBar,     "Focus Button Bar search box"},
            {BindAction.ShowSDTSelected,        "Show Subdirectory Tip menu for selected folder"},
            {BindAction.SendToTray,             "Send window to tray"},
            {BindAction.FocusTabBar,            "Focus tab bar"},
            {BindAction.NewTab,                 "Open new tab"},
            {BindAction.NewWindow,              "Open new window"},
            {BindAction.NewFolder,              "Create new folder"},
            {BindAction.NewFile,                "Create new empty file"},
            {BindAction.SwitchToLastActivated,  "Switch to last activated tab"},
            {BindAction.MergeWindows,           "Merge open windows"},
            {BindAction.ShowRecentFilesMenu,    "Show Button Bar Recent Files menu"},
            {BindAction.SortTabsByName,         "Sort tabs by name"},
            {BindAction.SortTabsByPath,         "Sort tabs by path"},
            {BindAction.SortTabsByActive,       "Sort tabs by last activated"},
            {BindAction.UpOneLevel,             "Up one level"},
            {BindAction.Refresh,                "Refresh"},
            {BindAction.Paste,                  "Paste"},
            {BindAction.Maximize,               "Maximize"},
            {BindAction.Minimize,               "Minimize"},
            {BindAction.ItemOpenInNewTab,       "Open in new tab"},
            {BindAction.ItemOpenInNewTabNoSel,  "Open in new tab without activating"},
            {BindAction.ItemOpenInNewWindow,    "Open in new window"},
            {BindAction.ItemCut,                "Cut"},
            {BindAction.ItemCopy,               "Copy"},
            {BindAction.ItemDelete,             "Delete"},
            {BindAction.ItemProperties,         "Properties"},
            {BindAction.CopyItemPath,           "Copy path"},
            {BindAction.CopyItemName,           "Copy name"},
            {BindAction.ChecksumItem,           "View checksum"},
            {BindAction.CloseTab,               "Close tab"},
            {BindAction.CloseLeftTab,           "Close tabs to left"},
            {BindAction.CloseRightTab,          "Close tabs to right"},
            {BindAction.UpOneLevelTab,          "Up one level"},
            {BindAction.LockTab,                "Lock"},
            {BindAction.ShowTabMenu,            "Context menu"},
            {BindAction.TearOffTab,             "Open in new window"},
            {BindAction.CloneTab,               "Clone"},
            {BindAction.CopyTabPath,            "Copy path"},
            {BindAction.TabProperties,          "Properties"},
            {BindAction.ShowTabSubfolderMenu,   "Show Subdirectory Tip menu"},
            {BindAction.CloseAllButThis,        "Close all but this"}
        };
        private static readonly Dictionary<MouseChord, string> MouseButtonItems
                = new Dictionary<MouseChord, string> {
            {MouseChord.Left,   "Left Click"},
            {MouseChord.Right,  "Right Click"},
            {MouseChord.Middle, "Middle Click"},
            {MouseChord.Double, "Double Click"},
            {MouseChord.X1,     "X1 Click"},
            {MouseChord.X2,     "X2 Click"},
        };
        private static readonly Dictionary<MouseChord, string> MouseModifierItems
                = new Dictionary<MouseChord, string> {
            {MouseChord.None,   "-"},
            {MouseChord.Shift,  "Shift"},
            {MouseChord.Ctrl,   "Ctrl"},
            {MouseChord.Alt,    "Alt"},
        };
        #endregion

        #region Mouse Action dictionaries

        // Which actions are valid for which targets?
        private readonly static Dictionary<MouseTarget, Dictionary<BindAction, string>> MouseTargetActions
                = new Dictionary<MouseTarget, Dictionary<BindAction, string>> {
            {MouseTarget.Anywhere, new BindAction[] {
                BindAction.GoBack,
                BindAction.GoForward,
                BindAction.NextTab,
                BindAction.PreviousTab
            }.ToDictionary(k => k, k => MouseActionItems[k])},
            {MouseTarget.Tab, new BindAction[] {
                BindAction.CloseTab,
                BindAction.CloseAllButThis,
                BindAction.UpOneLevelTab,
                BindAction.LockTab,
                BindAction.ShowTabMenu,
                BindAction.CloneTab,
                BindAction.TearOffTab,
                BindAction.CopyTabPath,
                BindAction.TabProperties,
                BindAction.ShowTabSubfolderMenu,
            }.ToDictionary(k => k, k => MouseActionItems[k])},
            {MouseTarget.TabBarBackground, new BindAction[] {
                BindAction.NewTab,
                BindAction.NewWindow,
                BindAction.UpOneLevel,
                BindAction.CloseAllButCurrent,
                BindAction.CloseWindow,
                BindAction.RestoreLastClosed,
                BindAction.CloneCurrent,
                BindAction.TearOffCurrent,
                BindAction.LockAll,
                BindAction.BrowseFolder,
                BindAction.ShowOptions,
                BindAction.ShowToolbarMenu,
                BindAction.SortTabsByName,
                BindAction.SortTabsByPath,
                BindAction.SortTabsByActive,
            }.ToDictionary(k => k, k => MouseActionItems[k])},
            {MouseTarget.FolderLink, new BindAction[] {
                BindAction.ItemOpenInNewTab,
                BindAction.ItemOpenInNewTabNoSel,
                BindAction.ItemOpenInNewWindow,
                BindAction.ItemProperties,
                BindAction.CopyItemPath,
                BindAction.CopyItemName,
            }.ToDictionary(k => k, k => MouseActionItems[k])},
            {MouseTarget.ExplorerItem, new BindAction[] {
                BindAction.ItemOpenInNewTab,
                BindAction.ItemOpenInNewTabNoSel,
                BindAction.ItemOpenInNewWindow,
                BindAction.ItemCut,
                BindAction.ItemCopy,        
                BindAction.ItemDelete,
                BindAction.ItemProperties,
                BindAction.CopyItemPath,
                BindAction.CopyItemName,
                BindAction.ChecksumItem,
            }.ToDictionary(k => k, k => MouseActionItems[k])},
            {MouseTarget.ExplorerBackground, new BindAction[] {
                BindAction.BrowseFolder,
                BindAction.NewFolder,
                BindAction.NewFile,
                BindAction.UpOneLevel,
                BindAction.Refresh,
                BindAction.Paste,
                BindAction.Maximize,
                BindAction.Minimize,
            }.ToDictionary(k => k, k => MouseActionItems[k])},
        };

        // Which buttons are valid for which targets?
        private static readonly Dictionary<MouseTarget, Dictionary<MouseChord, string>> MouseTargetButtons
                = new Dictionary<MouseTarget, Dictionary<MouseChord, string>> {
            {MouseTarget.Anywhere, new MouseChord[] {
                MouseChord.X1,
                MouseChord.X2,
            }.ToDictionary(k => k, k => MouseButtonItems[k])},
            {MouseTarget.ExplorerBackground, new MouseChord[] {
                MouseChord.Middle,
                MouseChord.Double,
            }.ToDictionary(k => k, k => MouseButtonItems[k])},
            {MouseTarget.ExplorerItem, new MouseChord[] {
                MouseChord.Middle,
            }.ToDictionary(k => k, k => MouseButtonItems[k])},
            {MouseTarget.FolderLink, new MouseChord[] {
                MouseChord.Left,
                MouseChord.Middle,
            }.ToDictionary(k => k, k => MouseButtonItems[k])},
            {MouseTarget.Tab, new MouseChord[] {
                MouseChord.Left,
                MouseChord.Middle,
                MouseChord.Double,
            }.ToDictionary(k => k, k => MouseButtonItems[k])},
            {MouseTarget.TabBarBackground, new MouseChord[] {
                MouseChord.Left,
                MouseChord.Middle,
                MouseChord.Double,
            }.ToDictionary(k => k, k => MouseButtonItems[k])},
        };

        #endregion

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

        #region ---------- Dialog-wide Stuff ----------

        private OptionsDialog() {
            QTTabBarClass tabBar = InstanceManager.CurrentTabBar;
            if(tabBar != null) {
                // we should probably assert this.
                pluginManager = tabBar.GetPluginManager();
            }
            workingConfig = QTUtility2.DeepClone(ConfigManager.LoadedConfig);
            InitializeComponent();
            cmbNewTabPos.ItemsSource = NewTabPosItems;
            cmbNextAfterClosed.ItemsSource = NextAfterCloseItems;
            cmbRebarStretchMode.ItemsSource = StretchModeItems;

            InitializeTips();
            InitializeMouse();
            InitializeKeys();
            InitializeGroups();
            InitializeApps();
            InitializeButtonBar();

            // Initialize the plugin tab
            CurrentPlugins = new ObservableCollection<PluginEntry>();
            foreach(PluginAssembly assembly in PluginManager.PluginAssemblies) {
                CreatePluginEntry(assembly, false);
            }
            lstPluginView.ItemsSource = CurrentPlugins;

            // Took me forever to figure out that this was necessary.  Why isn't this the default?!!
            // Bindings in context menus won't work without this.
            NameScope.SetNameScope(ctxTabTextColor, NameScope.GetNameScope(this));
            NameScope.SetNameScope(ctxShadowTextColor, NameScope.GetNameScope(this));
        }

        public void Dispose() {
            // TODO
        }

        private void UpdateOptions() {
            List<PluginAssembly> assemblies;
            CommitTips();
            CommitMouse();
            CommitPlugins(out assemblies);
            CommitButtonBar();
            CommitGroups();
            bool fButtonBarNeedsRefresh = Config.BBar.LargeButtons != workingConfig.bbar.LargeButtons;
            ConfigManager.LoadedConfig = QTUtility2.DeepClone(workingConfig);
            ConfigManager.WriteConfig();
            QTTabBarClass tabBar = InstanceManager.CurrentTabBar;
            if(tabBar != null) {
                tabBar.Invoke(new Action(tabBar.RefreshOptions));
                tabBar.Invoke(new Action(() => tabBar.odCallback_ManagePlugin(assemblies)));
            }
            QTButtonBar.BroadcastConfigChanged(fButtonBarNeedsRefresh);
            PluginManager.SavePluginAssemblies();
            PluginManager.SavePluginShortcutKeys();
        }

        private void ListBoxItem_PreviewMouseRightButtonDown(object sender, MouseButtonEventArgs e) {
            lstCategories.Focus();
            e.Handled = true;
        }

        private void ListBoxItem_PreviewMouseDown(object sender, MouseButtonEventArgs e) {
            ((ListBoxItem)sender).Focus();
            ((ListBoxItem)sender).IsSelected = true;
            e.Handled = true;
        }

        private void btnResetPage_Click(object sender, RoutedEventArgs e) {
            // todo: confirm
            TabItem tab = ((TabItem)tabbedPanel.SelectedItem);
            BindingExpression expr = tab.GetBindingExpression(DataContextProperty);
            if(expr != null) {
                PropertyInfo prop = typeof(Config).GetProperty(expr.ParentBinding.Path.Path);
                object c = Activator.CreateInstance(prop.PropertyType);
                prop.SetValue(workingConfig, c, null);
                expr.UpdateTarget();
            }
            else {
                // todo
            }
        }

        private void btnResetAll_Click(object sender, RoutedEventArgs e) {
            // todo: confirm
            workingConfig = new Config();
            tabbedPanel.GetBindingExpression(DataContextProperty).UpdateTarget();
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

            // Refill the button bar pool and key listview, to account for plugins
            InitializeButtonBar();
            InitializeKeys();
        }

        #endregion

        #region ---------- Window ----------

        #endregion

        #region ---------- Tweaks ----------

        private void btnAltRowColor_Click(object sender, RoutedEventArgs e) {
            // Works for both buttons.  Each button's Tag is bound to the corresponding property.
            var button = (Button)sender;
            ColorDialogEx cd = new ColorDialogEx { Color = (System.Drawing.Color)button.Tag };
            if(System.Windows.Forms.DialogResult.OK == cd.ShowDialog()) {
                button.Tag = cd.Color;
            }
        }

        #endregion

        #region ---------- Tooltips ----------

        private void InitializeTips() {
            TextFileTypes = new ObservableCollection<FileTypeEntry>(
                    workingConfig.tips.TextExt.Select(ext => new FileTypeEntry(this, ext)));
            MediaFileTypes = new ObservableCollection<FileTypeEntry>(
                    workingConfig.tips.ImageExt.Select(ext => new FileTypeEntry(this, ext)));

            lstTextFileTypes.ItemsSource = TextFileTypes;
            lstMediaFileTypes.ItemsSource = MediaFileTypes;
        }

        private void CommitTips() {
            workingConfig.tips.TextExt = TextFileTypes.Select(entry => entry.Extension).ToList();
            workingConfig.tips.ImageExt = MediaFileTypes.Select(entry => entry.Extension).ToList();
        }

        private void AddNewFileType(ListBox control) {
            ICollection<FileTypeEntry> source = (ICollection<FileTypeEntry>)control.ItemsSource;
            FileTypeEntry item = new FileTypeEntry(this, "");
            source.Add(item);
            control.SelectedItem = item;
            control.ScrollIntoView(item);
            control.Focus();
            item.IsEditing = true;
        }

        private static void RemoveSelectedFileType(ListBox control) {
            var list = (ObservableCollection<FileTypeEntry>)control.ItemsSource;
            int idx = control.SelectedIndex;
            if(idx == -1) return;
            foreach(FileTypeEntry item in new ArrayList(control.SelectedItems)) {
                list.Remove(item);
            }
            control.Focus();
            control.SelectedIndex = Math.Min(idx, list.Count - 1);
        }

        private void btnAddTextFileTypes_Click(object sender, RoutedEventArgs e) {
            AddNewFileType(lstTextFileTypes);
        }

        private void btnAddMediaFileTypes_Click(object sender, RoutedEventArgs e) {
            AddNewFileType(lstMediaFileTypes);
        }

        private void btnRemoveTextFileTypes_Click(object sender, RoutedEventArgs e) {
            RemoveSelectedFileType(lstTextFileTypes);
        }

        private void btnRemoveMediaFileTypes_Click(object sender, RoutedEventArgs e) {
            RemoveSelectedFileType(lstMediaFileTypes);
        }

        private void btnResetTextFileTypes_Click(object sender, RoutedEventArgs e) {
            lstTextFileTypes.ItemsSource = TextFileTypes = new ObservableCollection<FileTypeEntry>(
                    new Config._Tips().TextExt.Select(ext => new FileTypeEntry(this, ext)));
            lstTextFileTypes.ScrollIntoView(TextFileTypes.First());
        }

        private void btnResetMediaFileTypes_Click(object sender, RoutedEventArgs e) {
            lstMediaFileTypes.ItemsSource = MediaFileTypes = new ObservableCollection<FileTypeEntry>(
                    new Config._Tips().ImageExt.Select(ext => new FileTypeEntry(this, ext)));
            lstMediaFileTypes.ScrollIntoView(MediaFileTypes.First());
        }

        private void lstTextFileTypes_OnKeyDown(object sender, KeyEventArgs e) {
            if(e.Key == Key.Delete) RemoveSelectedFileType(lstTextFileTypes);
        }

        private void lstMediaFileTypes_OnKeyDown(object sender, KeyEventArgs e) {
            if(e.Key == Key.Delete) RemoveSelectedFileType(lstMediaFileTypes);
        }

        private void ListViewItem_Selected(object sender, RoutedEventArgs e) {
            FrameworkElement item = (FrameworkElement)sender;
            FrameworkElement parent = (FrameworkElement)item.Tag;
            if(parent.IsFocused) item.Focus();
        }

        #endregion

        #region ---------- Appearance ----------

        private void btnShadTextColor_OnChecked(object sender, RoutedEventArgs e) {
            var button = ((ToggleButton)sender);
            ContextMenu menu = button.ContextMenu;
            foreach(MenuItem mi in menu.Items) {
                mi.Icon = new System.Windows.Controls.Image { Source = ConvertToBitmapSource((Rectangle)mi.Tag) };
            }
            // Yeah, this is necessary even with the IsChecked <=> IsOpen binding.
            // Not sure why.
            menu.PlacementTarget = button;
            menu.Placement = PlacementMode.Bottom;
            menu.IsOpen = true;
        }

        private void miColorMenuEntry_OnClick(object sender, RoutedEventArgs e) {
            var mi = (MenuItem)sender;
            var rect = (Rectangle)mi.Tag;
            ColorDialogEx cd = new ColorDialogEx { Color = (System.Drawing.Color)rect.Tag };
            if(System.Windows.Forms.DialogResult.OK == cd.ShowDialog()) {
                rect.Tag = cd.Color;
            }
        }

        private void btnRebarBGColorChoose_Click(object sender, RoutedEventArgs e) {
            ColorDialogEx cd = new ColorDialogEx { Color = workingConfig.skin.RebarColor };
            if(System.Windows.Forms.DialogResult.OK == cd.ShowDialog()) {
                workingConfig.skin.RebarColor = cd.Color;
            }
        }

        #endregion

        #region ---------- Mouse ----------

        private void InitializeMouse() {
            MouseBindings = new ObservableCollection<MouseEntry>();
            foreach(var p in workingConfig.mouse.GlobalMouseActions) {
                MouseBindings.Add(new MouseEntry(MouseTarget.Anywhere, p.Key, p.Value));
            }
            foreach(var p in workingConfig.mouse.MarginActions) {
                MouseBindings.Add(new MouseEntry(MouseTarget.ExplorerBackground, p.Key, p.Value));
            }
            foreach(var p in workingConfig.mouse.ItemActions) {
                MouseBindings.Add(new MouseEntry(MouseTarget.ExplorerItem, p.Key, p.Value));
            }
            foreach(var p in workingConfig.mouse.LinkActions) {
                MouseBindings.Add(new MouseEntry(MouseTarget.FolderLink, p.Key, p.Value));
            }
            foreach(var p in workingConfig.mouse.TabActions) {
                MouseBindings.Add(new MouseEntry(MouseTarget.Tab, p.Key, p.Value));
            }
            foreach(var p in workingConfig.mouse.BarActions) {
                MouseBindings.Add(new MouseEntry(MouseTarget.TabBarBackground, p.Key, p.Value));
            }
            ICollectionView view = CollectionViewSource.GetDefaultView(MouseBindings);
            PropertyGroupDescription groupDescription = new PropertyGroupDescription("TargetText");
            view.GroupDescriptions.Add(groupDescription);
            lvwMouseBindings.ItemsSource = view;

            cmbMouseModifiers.ItemsSource = MouseModifierItems;
            cmbMouseTarget.ItemsSource = MouseTargetItems;
            cmbMouseTarget.SelectedIndex = 0;
            cmbMouseModifiers.SelectedIndex = 0;
            cmbMouseButtons.SelectedIndex = 0;
        }

        private void CommitMouse() {
            workingConfig.mouse.GlobalMouseActions = MouseBindings
                    .Where(e => e.Target == MouseTarget.Anywhere)
                    .ToDictionary(e => e.Chord, e => e.Action);
            workingConfig.mouse.MarginActions = MouseBindings
                    .Where(e => e.Target == MouseTarget.ExplorerBackground)
                    .ToDictionary(e => e.Chord, e => e.Action);
            workingConfig.mouse.ItemActions = MouseBindings
                    .Where(e => e.Target == MouseTarget.ExplorerItem)
                    .ToDictionary(e => e.Chord, e => e.Action);
            workingConfig.mouse.LinkActions = MouseBindings
                    .Where(e => e.Target == MouseTarget.FolderLink)
                    .ToDictionary(e => e.Chord, e => e.Action);
            workingConfig.mouse.TabActions = MouseBindings
                    .Where(e => e.Target == MouseTarget.Tab)
                    .ToDictionary(e => e.Chord, e => e.Action);
            workingConfig.mouse.BarActions = MouseBindings
                    .Where(e => e.Target == MouseTarget.TabBarBackground)
                    .ToDictionary(e => e.Chord, e => e.Action);
        }

        private void cmbMouseTarget_SelectionChanged(object sender, SelectionChangedEventArgs e) {
            var v = (KeyValuePair<MouseTarget, string>)e.AddedItems[0];
            cmbMouseAction.ItemsSource = MouseTargetActions[v.Key];
            cmbMouseButtons.ItemsSource = MouseTargetButtons[v.Key];
            cmbMouseButtons.SelectedIndex = 0;
        }

        private void btnMouseActionAdd_Click(object sender, RoutedEventArgs e) {
            MouseTarget target = (MouseTarget)cmbMouseTarget.SelectedValue;
            BindAction action = (BindAction)cmbMouseAction.SelectedValue;
            MouseChord chord = (MouseChord)cmbMouseModifiers.SelectedValue | (MouseChord)cmbMouseButtons.SelectedValue;
            MouseEntry entry = MouseBindings.FirstOrDefault(m => m.Target == target && m.Chord == chord);
            if(entry != null) {
                const string removePlugin = "This mouse chord is already bound to the following action:\n\n{0}\n\nReplace?";
                if(MessageBox.Show(string.Format(removePlugin, entry.ActionText), string.Empty, MessageBoxButton.OKCancel, MessageBoxImage.Question) != MessageBoxResult.OK) {
                    return;
                }
                entry.Action = action;
            }
            else {
                MouseBindings.Add(new MouseEntry(target, chord, action));
            }
        }

        #endregion

        #region ---------- Keyboard ----------

        private void InitializeKeys() {
            string[] arrActions = QTUtility.TextResourcesDic["ShortcutKeys_ActionNames"];
            string[] arrGrps = QTUtility.TextResourcesDic["ShortcutKeys_Groups"];
            int[] keys = workingConfig.keys.Shortcuts;
            HotkeyEntries = new List<HotkeyEntry>();
            // todo: validate arrActions length
            for(int i = 0; i <= (int)QTUtility.LAST_KEYBOARD_ACTION; ++i) {
                HotkeyEntries.Add(new HotkeyEntry(keys, i, arrActions[i], arrGrps[0]));
            }

            foreach(string pluginID in QTUtility.dicPluginShortcutKeys.Keys) {
                Plugin p;
                keys = QTUtility.dicPluginShortcutKeys[pluginID];
                if(!pluginManager.TryGetPlugin(pluginID, out p)) continue;
                PluginInformation pi = p.PluginInformation;
                if(pi.ShortcutKeyActions == null) continue;
                string group = pi.Name + " (" + arrGrps[1] + ")";
                if(keys == null && keys.Length == pi.ShortcutKeyActions.Length) {
                    Array.Resize(ref keys, pi.ShortcutKeyActions.Length);
                    // Hmm, I don't like this...
                    QTUtility.dicPluginShortcutKeys[pluginID] = keys;
                }
                HotkeyEntries.AddRange(pi.ShortcutKeyActions.Select((act, i) => new HotkeyEntry(keys, i, act, group)));
            }
            ICollectionView view = CollectionViewSource.GetDefaultView(HotkeyEntries);
            PropertyGroupDescription groupDescription = new PropertyGroupDescription("Group");
            view.GroupDescriptions.Add(groupDescription);
            lvwHotkeys.ItemsSource = view;
        }

        private bool CheckForKeyConflicts(Keys key) {
            const string Conflict = "This key is already assigned to:\n{0}\n\nReassign?";
            const string GroupPrefix = "Open group \"{0}\"";
            //const string AppPrefix = "Launch application \"{0}\"";
            const string MsgTitle = "Keystroke conflict";

            HotkeyEntry hotkey = HotkeyEntries.FirstOrDefault(e => e.Key == key);
            if(hotkey != null) {
                if(MessageBox.Show(string.Format(Conflict, hotkey.Action), MsgTitle, MessageBoxButton.OKCancel, MessageBoxImage.Warning) == MessageBoxResult.OK) {
                    hotkey.Key = Keys.None;
                    return true;
                }
                else {
                    return false;
                }
            }

            GroupEntry group = CurrentGroups.FirstOrDefault(e => e.ShortcutKey == key);
            if(group != null) {
                if(MessageBox.Show(string.Format(Conflict, string.Format(GroupPrefix, group.Name)), MsgTitle, MessageBoxButton.OKCancel, MessageBoxImage.Warning) == MessageBoxResult.OK) {
                    group.ShortcutKey = Keys.None;
                    return true;
                }
                else {
                    return false;
                }
            }

            // todo: apps
            return true;
        }

        private static bool IsInvalidShortcutKey(Keys key, Keys modKeys) {

            if(modKeys == Keys.None) return true;

            // keys not allowed even with any modifier keys 
            switch(key) {
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
                    return true;
            }

            // keys not allowed as one key
            switch(key | modKeys) {
                case Keys.LWin:
                case Keys.RWin:
                case Keys.Delete:
                case Keys.Apps:
                case Keys.Tab:
                case Keys.Left:
                case Keys.Up:
                case Keys.Right:
                case Keys.Down:
                    return true;
            }

            // invalid key combinations 
            bool c = (modKeys & Keys.Control) == Keys.Control;
            bool s = (modKeys & Keys.Shift) == Keys.Shift;
            bool a = (modKeys & Keys.Alt) == Keys.Alt;

            if(c && !s && !a) {
                switch(key) {
                    case Keys.C:	// Ctrl + C
                    case Keys.A:	// Ctrl + A
                    case Keys.Z:	// Ctrl + Z
                    case Keys.V:	// Ctrl + V
                    case Keys.X:	// Ctrl + X
                        System.Media.SystemSounds.Hand.Play();
                        return true;
                }
            }
            else if(!c && !s && a) {
                switch(key) {
                    case Keys.Left:		// Alt + Left
                    case Keys.Right:	// Alt + Right
                    case Keys.F4:		// Alt + F4
                        System.Media.SystemSounds.Hand.Play();
                        return true;
                }
            }

            return false;
        }

        private void lvwHotkeys_PreviewKeyDown(object sender, KeyEventArgs e) {
            if(lvwHotkeys.SelectedItems.Count != 1) return;
            HotkeyEntry entry = (HotkeyEntry)lvwHotkeys.SelectedItem;
            Keys newKey;
            if(!ProcessNewHotkey(e, entry.Key, out newKey)) return;
            bool wasNotAssigned = !entry.Assigned;
            entry.Key = newKey;
            if(wasNotAssigned && entry.Assigned) entry.Enabled = true;
            e.Handled = true;
        }

        private void lvwHotkeys_TextInput(object sender, TextCompositionEventArgs e) {
            e.Handled = true;
        }

        #endregion

        #region ---------- Groups ----------

        private void InitializeGroups() {
            tvwGroups.ItemsSource = CurrentGroups = new ParentedCollection<GroupEntry>(
                    GroupsManager.Groups.Select(g => new GroupEntry(
                    g.Name, g.ShortcutKey, g.Startup, g.Paths.Select(p => new FolderEntry(p)))));
        }

        private void CommitGroups() {
            GroupsManager.Groups = new List<Group>(
                    CurrentGroups.Select(g => new Group(
                    g.Name, g.ShortcutKey, g.Startup, g.Folders.Select(f => f.Path).ToList())));
        }

        private GroupEntry GetParentGroup(FolderEntry folder) {
            return CurrentGroups.FirstOrDefault(entry => entry.Folders.Contains(folder));
        }

        private static void UpDownOnTreeView(TreeView tvw, bool up) {
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

        private void btnGroupsAddGroup_Click(object sender, RoutedEventArgs e) {
            GroupEntry item = new GroupEntry("New Group");
            tvwGroups.Focus();
            IList col = (IList)tvwGroups.ItemsSource;
            object sel = tvwGroups.SelectedItem;
            int idx = sel == null
                    ? tvwGroups.Items.Count
                    : CurrentGroups.IndexOf(sel as GroupEntry ?? GetParentGroup((FolderEntry)sel)) + 1;
            col.Insert(idx, item);
            item.IsSelected = true;
            item.IsEditing = true;
        }

        private void btnGroupsMoveNodeUp_Click(object sender, RoutedEventArgs e) {
            UpDownOnTreeView(tvwGroups, true);
        }

        private void btnGroupsMoveNodeDown_Click(object sender, RoutedEventArgs e) {
            UpDownOnTreeView(tvwGroups, false);
        }

        private void btnGroupsAddFolder_Click(object sender, RoutedEventArgs e) {
            // TODO: Generates new group if the view is empty.

            object sel = tvwGroups.SelectedItem;
            if(sel == null) return;

            GroupEntry group;
            int index;
            if(sel is FolderEntry) {
                FolderEntry entry = (FolderEntry)sel;
                group = GetParentGroup(entry);
                index = group.Folders.IndexOf(entry) + 1;
            }
            else {
                group = (GroupEntry)sel;
                index = group.Folders.Count;
            }

            FolderBrowserDialogEx dlg = new FolderBrowserDialogEx();
            if(dlg.ShowDialog() != System.Windows.Forms.DialogResult.OK) return;
            FolderEntry folder = new FolderEntry(dlg.SelectedPath);
            group.Folders.Insert(index, folder);
            group.IsExpanded = true;
            folder.IsSelected = true;
        }

        private void btnGroupsRemoveNode_Click(object sender, RoutedEventArgs e) {
            ITreeViewItem sel = tvwGroups.SelectedItem as ITreeViewItem;
            if(sel == null) return;
            IList col = sel.ParentList;
            int index = col.IndexOf(sel);
            col.RemoveAt(index);
            if(col.Count == 0) return;
            if(index == col.Count) --index;
            ((ITreeViewItem)col[index]).IsSelected = true;
        }

        private void tvwGroupsApps_MouseDown(object sender, MouseButtonEventArgs e) {
            IEditableEntry entry = ((TreeView)sender).SelectedItem as IEditableEntry;
            if(entry != null) entry.IsEditing = false;
        }

        private void txtGroupHotkey_OnPreviewKeyDown(object sender, KeyEventArgs e) {
            e.Handled = true;
            GroupEntry entry = tvwGroups.SelectedItem as GroupEntry;
            if(entry == null) return;
            Keys newKey;
            if(ProcessNewHotkey(e, entry.ShortcutKey, out newKey)) {
                entry.ShortcutKey = newKey;
            }
        }

        #endregion

        #region ---------- Applications ----------

        private void InitializeApps() {
            tvwApps.ItemsSource = CurrentApps = new ParentedCollection<AppEntry>();
        }

        private void CommitApps() {
            // todo
        }

        private void btnAddApp_Click(object sender, RoutedEventArgs e) {
            string path;
            using(OpenFileDialog ofd = new OpenFileDialog()) {
                if(System.Windows.Forms.DialogResult.OK != ofd.ShowDialog()) return;
                path = ofd.FileName;
            }
            AppEntry sel = tvwApps.SelectedItem as AppEntry;
            IList list = sel == null ? CurrentApps : sel.IsFolder ? sel.Children : sel.ParentList;
            int idx = sel == null ? 0 : list.IndexOf(sel) + 1;
            AppEntry entry = new AppEntry(false, path);
            list.Insert(idx, entry);
            entry.IsSelected = true;
            if(sel != null && sel.IsFolder) sel.IsExpanded = true;
        }

        private void btnAddAppFolder_Click(object sender, RoutedEventArgs e) {
            AppEntry sel = tvwApps.SelectedItem as AppEntry;
            IList list = sel == null ? CurrentApps : sel.IsFolder ? sel.Children : sel.ParentList;
            int idx = sel == null ? 0 : list.IndexOf(sel) + 1;
            AppEntry entry = new AppEntry(true, "Folder");
            list.Insert(idx, entry);
            entry.IsSelected = true;
            if(sel != null && sel.IsFolder) sel.IsExpanded = true;
        }

        private void btnRemoveApp_Click(object sender, RoutedEventArgs e) {
            AppEntry sel = tvwApps.SelectedItem as AppEntry;
            if(sel == null) return;
            if(sel.IsFolder && sel.Children.Count > 0) {
                // todo: confirm
            }
            IList list = sel.ParentList;
            int index = list.IndexOf(sel);
            list.RemoveAt(index);
            if(list.Count == 0) return;
            if(index == list.Count) --index;
            ((ITreeViewItem)list[index]).IsSelected = true;
        }

        private void btnAppsMoveNodeUp_Click(object sender, RoutedEventArgs e) {
            UpDownOnTreeView(tvwApps, true);
        }

        private void btnAppsMoveNodeDown_Click(object sender, RoutedEventArgs e) {
            UpDownOnTreeView(tvwApps, false);
        }

        private void btnVars_Click(object sender, RoutedEventArgs e) {
            var button = ((Button)sender);
            ContextMenu menu = button.ContextMenu;
            menu.PlacementTarget = button;
            menu.Placement = PlacementMode.Bottom;
            menu.IsOpen = true;
        }

        private void miArgVars_Click(object sender, RoutedEventArgs e) {
            InsertVar(txtAppArgs, (string)((MenuItem)sender).Tag);
        }

        private void miWorkingVars_Click(object sender, RoutedEventArgs e) {
            InsertVar(txtAppDir, (string)((MenuItem)sender).Tag);
        }

        private static void InsertVar(TextBox textbox, string var) {
            int caret = textbox.CaretIndex;
            textbox.Text = textbox.Text.Insert(caret, var);
            textbox.CaretIndex = caret + var.Length;
        }

        #endregion

        #region ---------- Button Bar ----------

        private void InitializeButtonBar() {
            // Initialize the button bar tab.
            // todo: options, localize, etc...
            ButtonItemsDisplayName = QTUtility.TextResourcesDic["ButtonBar_BtnName"];
            imageStripLarge = new ImageStrip(new Size(24, 24));
            using(Bitmap b = Resources_Image.ButtonStrip24) {
                imageStripLarge.AddStrip(b);
            }
            imageStripSmall = new ImageStrip(new Size(16, 16));
            using(Bitmap b = Resources_Image.ButtonStrip16) {
                imageStripSmall.AddStrip(b);
            }
            ButtonPool = new ObservableCollection<ButtonEntry>();
            CurrentButtons = new ObservableCollection<ButtonEntry>();

            // Create a list of all the plugin buttons, and store the list 
            // index of the first button of each plugin in a dictionary keyed
            // on plugin ID.
            int pluginListPos = 0;
            var dicPluginListPos = new Dictionary<string, int>();
            var lstPluginButtons = new List<ButtonEntry>();
            foreach(PluginInformation pi in PluginManager.PluginInformations.OrderBy(pi => pi.Name)) {
                if(pi.PluginType == PluginType.Interactive) {
                    dicPluginListPos[pi.PluginID] = pluginListPos;
                    lstPluginButtons.Add(new ButtonEntry(this, pluginListPos++, pi, 0));
                }
                else if(pi.PluginType == PluginType.BackgroundMultiple) {
                    Plugin plugin;
                    if(pluginManager.TryGetPlugin(pi.PluginID, out plugin)) {
                        IBarMultipleCustomItems bmci = plugin.Instance as IBarMultipleCustomItems;
                        try {
                            if(bmci != null && bmci.Count > 0) {
                                // This is to maintain backwards compatibility.
                                bmci.Initialize(Enumerable.Range(0, bmci.Count).ToArray());
                                dicPluginListPos[pi.PluginID] = pluginListPos;
                                for(int i = 0; i < bmci.Count; i++) {
                                    lstPluginButtons.Add(new ButtonEntry(this, pluginListPos++, pi, i));
                                }
                            }
                        }
                        catch { }
                    }
                }
            }

            // Add the current buttons (right pane)
            int pluginIndex = 0;
            foreach(int i in workingConfig.bbar.ButtonIndexes) {
                if(i == QTButtonBar.BUTTONINDEX_PLUGIN) {
                    if(pluginIndex < PluginManager.ActivatedButtonsOrder.Count) {
                        var pluginButton = PluginManager.ActivatedButtonsOrder[pluginIndex];
                        if(dicPluginListPos.ContainsKey(pluginButton.id)) {
                            CurrentButtons.Add(lstPluginButtons[dicPluginListPos[pluginButton.id] + pluginButton.index]);
                        }
                    }
                    pluginIndex++;
                }
                else {
                    CurrentButtons.Add(new ButtonEntry(this, i));
                }
            }

            // Add the rest of the buttons to the button pool (left pane)
            ButtonPool.Add(new ButtonEntry(this, QTButtonBar.BII_SEPARATOR));
            for(int i = 1; i < QTButtonBar.INTERNAL_BUTTON_COUNT; i++) {
                if(!workingConfig.bbar.ButtonIndexes.Contains(i)) {
                    ButtonPool.Add(new ButtonEntry(this, i));
                }
            }
            foreach(ButtonEntry entry in lstPluginButtons) {
                if(!CurrentButtons.Contains(entry)) {
                    ButtonPool.Add(entry);
                }
            }
            lstButtonBarPool.ItemsSource = ButtonPool;
            lstButtonBarCurrent.ItemsSource = CurrentButtons;
        }

        private void CommitButtonBar() {
            var pluginButtons = new List<PluginManager.PluginButton>();
            for(int i = 0; i < CurrentButtons.Count; i++) {
                ButtonEntry entry = CurrentButtons[i];
                if(entry.Index >= QTButtonBar.BUTTONINDEX_PLUGIN) {
                    if(entry.PluginInfo.Enabled) {
                        pluginButtons.Add(new PluginManager.PluginButton {
                            id = entry.PluginInfo.PluginID,
                            index = entry.PluginButtonIndex
                        });
                    }
                    else {
                        CurrentButtons.RemoveAt(i--);
                    }
                }
            }
            PluginManager.ActivatedButtonsOrder = pluginButtons;
            PluginManager.SaveButtonOrder();
            workingConfig.bbar.ButtonIndexes = CurrentButtons.Select(
                    e => Math.Min(e.Index, QTButtonBar.BUTTONINDEX_PLUGIN)).ToArray();
        }

        private void btnBBarAdd_Click(object sender, RoutedEventArgs e) {
            int sel = lstButtonBarPool.SelectedIndex;
            if(sel == -1) return;
            ButtonEntry entry = ButtonPool[sel];
            if(entry.Index == QTButtonBar.BII_SEPARATOR) {
                entry = new ButtonEntry(this, QTButtonBar.BII_SEPARATOR);
            }
            else {
                ButtonPool.RemoveAt(sel);
                if(sel == ButtonPool.Count) --sel;
                if(sel >= 0) {
                    lstButtonBarPool.SelectedIndex = sel;
                    lstButtonBarPool.ScrollIntoView(lstButtonBarPool.SelectedItem);
                }
            }
            if(lstButtonBarCurrent.SelectedIndex == -1) {
                CurrentButtons.Add(entry);
                lstButtonBarCurrent.SelectedIndex = CurrentButtons.Count - 1;
            }
            else {
                CurrentButtons.Insert(lstButtonBarCurrent.SelectedIndex + 1, entry);
                lstButtonBarCurrent.SelectedIndex++;
            }
            lstButtonBarCurrent.ScrollIntoView(lstButtonBarCurrent.SelectedItem);
        }

        private void btnBBarRemove_Click(object sender, RoutedEventArgs e) {
            int sel = lstButtonBarCurrent.SelectedIndex;
            if(sel == -1) return;
            ButtonEntry entry = CurrentButtons[sel];
            CurrentButtons.RemoveAt(sel);
            if(sel == CurrentButtons.Count) --sel;
            if(sel >= 0) {
                lstButtonBarCurrent.SelectedIndex = sel;
                lstButtonBarCurrent.ScrollIntoView(lstButtonBarCurrent.SelectedItem);
            }
            if(entry.Index != QTButtonBar.BII_SEPARATOR) {
                int i = 0;
                while(i < ButtonPool.Count && ButtonPool[i].Index < entry.Index) ++i;
                ButtonPool.Insert(i, entry);
                lstButtonBarPool.SelectedIndex = i;
            }
            else {
                lstButtonBarPool.SelectedIndex = 0;
            }
            lstButtonBarPool.ScrollIntoView(lstButtonBarPool.SelectedItem);
        }

        private void btnBBarUp_Click(object sender, RoutedEventArgs e) {
            int sel = lstButtonBarCurrent.SelectedIndex;
            if(sel <= 0) return;
            CurrentButtons.Move(sel, sel - 1);
            lstButtonBarCurrent.ScrollIntoView(lstButtonBarCurrent.SelectedItem);
        }

        private void btnBBarDown_Click(object sender, RoutedEventArgs e) {
            int sel = lstButtonBarCurrent.SelectedIndex;
            if(sel == -1 || sel == CurrentButtons.Count - 1) return;
            CurrentButtons.Move(sel, sel + 1);
            lstButtonBarCurrent.ScrollIntoView(lstButtonBarCurrent.SelectedItem);
        }

        #endregion

        #region ---------- Plugins ----------

        private void btnPluginOptions_Click(object sender, RoutedEventArgs e) {
            if(lstPluginView.SelectedIndex == -1) return;
            PluginEntry entry = CurrentPlugins[lstPluginView.SelectedIndex];
            Plugin p;
            if(pluginManager.TryGetPlugin(entry.PluginID, out p) && p.Instance != null) {
                try {
                    p.Instance.OnOption();
                }
                catch(Exception ex) {
                    PluginManager.HandlePluginException(ex, new WindowInteropHelper(this).Handle,
                            entry.Name, "Open plugin option.");
                }
            }
        }

        private void btnPluginDisable_Click(object sender, RoutedEventArgs e) {
            if(lstPluginView.SelectedIndex == -1) return;
            PluginEntry entry = CurrentPlugins[lstPluginView.SelectedIndex];
            if(entry.DisableOnClose) {
                entry.DisableOnClose = false;
            }
            else if(entry.EnableOnClose) {
                entry.EnableOnClose = false;
            }
            else if(entry.Enabled) {
                entry.DisableOnClose = true;
            }
            else {
                entry.EnableOnClose = true;
            }
        }

        private void btnPluginRemove_Click(object sender, RoutedEventArgs e) {
            if(lstPluginView.SelectedIndex == -1) return;
            PluginEntry entry = CurrentPlugins[lstPluginView.SelectedIndex];
            PluginAssembly pluingAssembly = entry.PluginAssembly;
            if(pluingAssembly.PluginInformations.Count > 1) {
                string str = pluingAssembly.PluginInformations.Select(info => info.Name).StringJoin(", ");
                // todo localize
                const string removePlugin = "Uninstalling this plugin will also uninstall the following plugins:\n\n{0}\n\nProceed?";
                if(MessageBox.Show(string.Format(removePlugin, str), string.Empty, MessageBoxButton.OKCancel, MessageBoxImage.Question) != MessageBoxResult.OK) {
                    return;
                }
            }
            for(int i = 0; i < CurrentPlugins.Count; i++) {
                PluginEntry otherEntry = CurrentPlugins[i];
                if(otherEntry.PluginAssembly == entry.PluginAssembly) {
                    if(otherEntry.InstallOnClose) {
                        otherEntry.PluginAssembly.Dispose();
                        CurrentPlugins.RemoveAt(i);
                        --i;
                    }
                    else {
                        otherEntry.UninstallOnClose = true;
                    }
                }
            }
        }

        private void btnBrowsePlugin_Click(object sender, RoutedEventArgs e) {
            using(OpenFileDialog ofd = new OpenFileDialog()) {
                ofd.Filter = "Plugin files (*.dll)|*.dll";
                ofd.RestoreDirectory = true;
                ofd.Multiselect = true;

                if(System.Windows.Forms.DialogResult.OK != ofd.ShowDialog()) return;
                bool fFirst = true;
                foreach(string path in ofd.FileNames) {
                    PluginAssembly pa = new PluginAssembly(path);
                    if(!pa.PluginInfosExist) continue;
                    CreatePluginEntry(pa, true);
                    if(!fFirst) continue;
                    fFirst = false;
                    lstPluginView.SelectedItem = CurrentPlugins[CurrentPlugins.Count - 1];
                    lstPluginView.ScrollIntoView(lstPluginView.SelectedItem);
                }
            }
        }

        private void CreatePluginEntry(PluginAssembly pa, bool fAddedByUser) {
            if(!pa.PluginInfosExist || CurrentPlugins.Any(pe => pe.Path == pa.Path)) {
                return;
            }
            foreach(PluginInformation pi in pa.PluginInformations) {
                PluginEntry entry = new PluginEntry(this, pi, pa) { InstallOnClose = fAddedByUser };
                int i = 0;
                while(i < CurrentPlugins.Count && string.Compare(CurrentPlugins[i].Title, entry.Title, true) <= 0) ++i;
                CurrentPlugins.Insert(i, entry);
            }
        }

        // Call this BEFORE committing the button bar!
        private void CommitPlugins(out List<PluginAssembly> assemblies) {
            assemblies = new List<PluginAssembly>();

            // Don't dispose the assemblies here.  That will be done by the plugin manager
            // when the plugins are unloaded.
            for(int i = 0; i < CurrentPlugins.Count; ++i) {
                if(CurrentPlugins[i].UninstallOnClose) {
                    CurrentPlugins[i].Enabled = false;
                    CurrentPlugins.RemoveAt(i--);
                }
            }

            List<string> enabled = new List<string>();
            foreach(PluginEntry entry in CurrentPlugins) {
                PluginAssembly pa = entry.PluginAssembly;
                if(!assemblies.Contains(pa)) {
                    pa.Enabled = false;
                    assemblies.Add(pa);
                }
                if(entry.DisableOnClose) {
                    entry.Enabled = false;
                }
                else if(entry.EnableOnClose || entry.InstallOnClose) {
                    entry.Enabled = true;
                }
                entry.EnableOnClose = entry.DisableOnClose = entry.InstallOnClose = false;

                if(entry.Enabled) {
                    pa.Enabled = true;
                    enabled.Add(entry.PluginID);
                }
            }
            workingConfig.plugin.Enabled = enabled.ToArray();
        }

        #endregion

        #region ---------- Common Utility Functions ----------

        private bool ProcessNewHotkey(KeyEventArgs e, Keys current, out Keys hotkey) {
            Key wpfKey = (e.Key == Key.System ? e.SystemKey : e.Key);
            ModifierKeys wpfModKeys = Keyboard.Modifiers;

            // Ignore modifier keys.
            if(wpfKey == Key.LeftShift || wpfKey == Key.RightShift
                    || wpfKey == Key.LeftCtrl || wpfKey == Key.RightCtrl
                    || wpfKey == Key.LeftAlt || wpfKey == Key.RightAlt
                    || wpfKey == Key.LWin || wpfKey == Key.RWin) {
                hotkey = Keys.None;
                return false;
            }

            // Urgh, so many conversions between WPF and WinForms...
            hotkey = (Keys)KeyInterop.VirtualKeyFromKey(wpfKey);
            Keys modKeys = Keys.None;
            if((wpfModKeys & ModifierKeys.Alt) != 0) modKeys |= Keys.Alt;
            if((wpfModKeys & ModifierKeys.Control) != 0) modKeys |= Keys.Control;
            if((wpfModKeys & ModifierKeys.Shift) != 0) modKeys |= Keys.Shift;


            if(hotkey == Keys.Escape) {
                hotkey = Keys.None;
                return true;
            }
            else if(current != (hotkey | modKeys) && !IsInvalidShortcutKey(hotkey, modKeys) && CheckForKeyConflicts(hotkey | modKeys)) {
                hotkey |= modKeys;
                return true;
            }
            else {
                hotkey = Keys.None;
                return false;
            }
        }

        // Draws a control to a bitmap
        private static BitmapSource ConvertToBitmapSource(UIElement element) {
            var target = new RenderTargetBitmap((int)(element.RenderSize.Width), (int)(element.RenderSize.Height), 96, 96, PixelFormats.Pbgra32);
            var brush = new VisualBrush(element);
            var visual = new DrawingVisual();
            var drawingContext = visual.RenderOpen();

            drawingContext.DrawRectangle(brush, null, new Rect(new Point(0, 0),
                new Point(element.RenderSize.Width, element.RenderSize.Height)));
            drawingContext.Close();
            target.Render(visual);
            return target;
        }

        #endregion

        #region ---------- Interfaces / Helper Classes ----------

        private interface IEditableEntry {
            bool IsEditing { get; set; }
        }

        private interface IChildItem {
            IList ParentList { get; set; }
        }

        private interface ITreeViewItem : IChildItem {
            bool IsSelected { get; set; }
            bool IsExpanded { get; set; }
        }

        private sealed class ParentedCollection<TChild> : ObservableCollection<TChild> 
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

        #endregion

        #region ---------- Binding Classes ----------
        #pragma warning disable 0067 // "The event 'PropertyChanged' is never used"
        // ReSharper disable MemberCanBePrivate.Local
        // ReSharper disable UnusedMember.Local
        // ReSharper disable UnusedAutoPropertyAccessor.Local

        // Thanks to the damn-near LIFE-SAVING Notify Property Weaver,
        // INotifyPropertyChanged is implemented for us.  All properties
        // will automatically raise the PropertyChanged event when 
        // modified.  Bind away!

        private class ButtonEntry : INotifyPropertyChanged {
            public event PropertyChangedEventHandler PropertyChanged;
            private OptionsDialog parent;

            public PluginInformation PluginInfo { get; private set; }
            public int Index { get; private set; }
            public int PluginButtonIndex { get; private set; }
            public string Text {
                get {
                    if(Index >= QTButtonBar.BUTTONINDEX_PLUGIN) {
                        if(PluginInfo.PluginType == PluginType.BackgroundMultiple && PluginButtonIndex != -1) {
                            Plugin plugin;
                            if(parent.pluginManager.TryGetPlugin(PluginInfo.PluginID, out plugin)) {
                                try {
                                    return ((IBarMultipleCustomItems)plugin.Instance).GetName(PluginButtonIndex);
                                }
                                catch { }
                            }
                        }
                        return PluginInfo.Name;
                    }
                    else if(Index < parent.ButtonItemsDisplayName.Length) {
                        return parent.ButtonItemsDisplayName[Index];
                    }
                    else {
                        return "";
                    }
                }
            }

            public Image LargeImage { get { return getImage(true); } }
            public Image SmallImage { get { return getImage(false); } }
            private Image getImage(bool large) {
                if(Index >= QTButtonBar.BUTTONINDEX_PLUGIN) {
                    if(PluginInfo.PluginType == PluginType.BackgroundMultiple && PluginButtonIndex != -1) {
                        Plugin plugin;
                        if(parent.pluginManager.TryGetPlugin(PluginInfo.PluginID, out plugin)) {
                            try {
                                return ((IBarMultipleCustomItems)plugin.Instance).GetImage(large, PluginButtonIndex);
                            }
                            catch { }
                        }
                    }
                    return large
                            ? PluginInfo.ImageLarge ?? Resources_Image.imgPlugin24
                            : PluginInfo.ImageSmall ?? Resources_Image.imgPlugin16;
                }
                else if(Index == 0 || Index >= QTButtonBar.BII_WINDOWOPACITY) {
                    return null;
                }
                else {
                    return large
                            ? parent.imageStripLarge[Index - 1]
                            : parent.imageStripSmall[Index - 1];
                }
            }
            public ButtonEntry(OptionsDialog parent, int Index) {
                this.parent = parent;
                this.Index = Index;
            }
            public ButtonEntry(OptionsDialog parent, int Index, PluginInformation PluginInfo, int PluginButtonIndex) {
                this.parent = parent;
                this.PluginInfo = PluginInfo;
                this.Index = QTButtonBar.BUTTONINDEX_PLUGIN + Index;
                this.PluginButtonIndex = PluginButtonIndex;
            }
        }

        private class PluginEntry : INotifyPropertyChanged {
            public event PropertyChangedEventHandler PropertyChanged;
            private OptionsDialog parent;
            private PluginInformation PluginInfo;
            public PluginAssembly PluginAssembly { get; private set; }

            public Image Icon { get { return PluginInfo.ImageLarge ?? Resources_Image.imgPlugin24; } }
            public string Name { get { return PluginInfo.Name; } }
            public string Title { get { return Name + "  " + PluginInfo.Version; } }
            public string Author { get { return "by " + PluginInfo.Author; } }
            public string Desc { get { return PluginInfo.Description; } }
            public bool IsSelected { get; set; }
            public double Opacity { get { return Enabled ? 1.0 : 0.5; } }
            public bool DisableOnClose { get; set; }
            public bool EnableOnClose { get; set; }
            public bool InstallOnClose { get; set; }
            public bool UninstallOnClose { get; set; }
            public bool Enabled { get { return PluginInfo.Enabled; } set { PluginInfo.Enabled = value; } }
            public string PluginID { get { return PluginInfo.PluginID; } }
            public string Path { get { return PluginInfo.Path; } }

            private bool cachedHasOptions;
            private bool optionsQueried;

            public bool HasOptions {
                get {
                    if(!Enabled) return false;
                    if(optionsQueried) return cachedHasOptions;
                    Plugin p;
                    if(parent.pluginManager.TryGetPlugin(PluginID, out p)) {
                        try {
                            cachedHasOptions = p.Instance.HasOption;
                            optionsQueried = true;
                            return cachedHasOptions;
                        }
                        catch(Exception ex) {
                            PluginManager.HandlePluginException(ex, new WindowInteropHelper(parent).Handle, Name,
                                    "Checking if the plugin has options.");
                        }
                    }
                    return false;
                }
            }

            public PluginEntry(OptionsDialog parent, PluginInformation pluginInfo, PluginAssembly pluginAssembly) {
                this.parent = parent;
                PluginInfo = pluginInfo;
                PluginAssembly = pluginAssembly;
            }
        }

        private class HotkeyEntry : INotifyPropertyChanged {
            public event PropertyChangedEventHandler PropertyChanged;
            private int[] raws;
            public int RawKey {
                get { return raws[Index]; }
                set { raws[Index] = value; }
            }
            public bool Enabled {
                get { return (RawKey & QTUtility.FLAG_KEYENABLED) != 0 && RawKey != QTUtility.FLAG_KEYENABLED; }
                set { if(value) RawKey |= QTUtility.FLAG_KEYENABLED; else RawKey &= ~QTUtility.FLAG_KEYENABLED; }
            }
            public Keys Key {
                get { return (Keys)(RawKey & ~QTUtility.FLAG_KEYENABLED); }
                set { RawKey = (int)value | (Enabled ? QTUtility.FLAG_KEYENABLED : 0); }
            }
            public bool Assigned {
                get { return Key != Keys.None; }
            }
            public string HotkeyString {
                get { return QTUtility2.MakeKeyString(Key); }
            }

            public string Group { get; set; }
            public string Action { get; set; }
            public int Index { get; set; }
            public HotkeyEntry(int[] raws, int index, string action, string group) {
                this.raws = raws;
                Index = index;
                Action = action;
                Group = group;
            }
        }

        private class MouseEntry : INotifyPropertyChanged {
            public event PropertyChangedEventHandler PropertyChanged;
            public string GestureText {
                get {
                    MouseChord modifier = Chord & (MouseChord.Alt | MouseChord.Ctrl | MouseChord.Shift);
                    MouseChord button = Chord & ~(MouseChord.Alt | MouseChord.Ctrl | MouseChord.Shift);
                    return (modifier != MouseChord.None ? MouseModifierItems[modifier] + " + " : "")
                            + MouseButtonItems[button];
                }
            }
            public string TargetText { get { return MouseTargetItems[Target]; } }
            public string ActionText { get { return MouseActionItems[Action]; } }
            public MouseTarget Target { get; private set; }
            public BindAction Action { get; set; }
            public MouseChord Chord { get; private set; }

            public MouseEntry(MouseTarget target, MouseChord chord, BindAction action) {
                Target = target;
                Action = action;
                Chord = chord;
            }
        }

        private class FileTypeEntry : INotifyPropertyChanged, IEditableEntry {
            public event PropertyChangedEventHandler PropertyChanged;
            private OptionsDialog parent;

            private bool _IsEditing;
            public bool IsEditing {
                get { return _IsEditing; }
                set {
                    _IsEditing = value;
                    if(!_IsEditing && string.IsNullOrEmpty(Extension)) {
                        parent.TextFileTypes.Remove(this);
                        parent.MediaFileTypes.Remove(this);
                    }
                }
            }

            public bool IsSelected { get; set; }
            public string Extension { get; set; }

            public string DotExtension {
                get {
                    return "." + Extension;
                }
                set {
                    if(!value.StartsWith(".")) {
                        throw new ArgumentException();
                    }
                    Extension = value.Substring(1);
                }
            }
            public string FriendlyName {
                get {
                    // PENDING: Instead of something like GetFileType.

                    SHFILEINFO psfi = new SHFILEINFO();
                    int sz = System.Runtime.InteropServices.Marshal.SizeOf(psfi);
                    // SHGFI_TYPENAME | SHGFI_USEFILEATTRIBUTES
                    if(IntPtr.Zero == PInvoke.SHGetFileInfo("*" + DotExtension, 0x80, ref psfi, sz, 0x400 | 0x10)) {
                        return null;
                    }
                    else if(string.IsNullOrEmpty(psfi.szTypeName)) {
                        return null;
                    }
                    return psfi.szTypeName;
                }
            }
            public Image Icon {
                get {
                    return QTUtility.GetIcon(DotExtension, true).ToBitmap();
                }
            }
            public FileTypeEntry(OptionsDialog parent, string extension) {
                this.parent = parent;
                if(!extension.StartsWith(".")) {
                    extension = "." + extension;
                }
                DotExtension = extension;
            }
        }

        private class FolderEntry : INotifyPropertyChanged, IEditableEntry, ITreeViewItem {
            public event PropertyChangedEventHandler PropertyChanged;
            public IList ParentList { get; set; }
            public string Path { get; set; }
            public bool IsEditing { get; set; }
            public bool IsSelected { get; set; }
            public bool IsExpanded { get; set; }

            public string DisplayText {
                get {
                    return QTUtility2.MakePathDisplayText(Path, true);
                }
            }
            public Image Icon {
                get {
                    return QTUtility.GetIcon(Path, false).ToBitmap();
                }
            }
            public bool IsVirtualFolder {
                get {
                    return Path.StartsWith("::");
                }
            }

            public FolderEntry(string path) {
                Path = path;
            }

            public FolderEntry() {
            }
        }

        private class GroupEntry : INotifyPropertyChanged, IEditableEntry, ITreeViewItem {
            public event PropertyChangedEventHandler PropertyChanged;
            public IList ParentList { get; set; }
            public string Name { get; set; }
            public Image Icon { get; private set; }
            public ParentedCollection<FolderEntry> Folders { get; private set; }
            public bool Startup { get; set; }
            public Keys ShortcutKey { get; set; }
            public string HotkeyString {
                get { return QTUtility2.MakeKeyString(ShortcutKey); }
            }
            public bool IsEditing { get; set; }
            public bool IsSelected { get; set; }
            public bool IsExpanded { get; set; }
            
            private void Folders_CollectionChanged(object sender, NotifyCollectionChangedEventArgs e) {
                if(e.OldItems != null) {
                    foreach(FolderEntry child in e.OldItems) {
                        child.PropertyChanged -= FolderEntry_PropertyChanged;
                    }
                }
                if(e.NewItems != null) {
                    foreach(FolderEntry child in e.NewItems) {
                        child.PropertyChanged += FolderEntry_PropertyChanged;
                    }
                }
                RefreshIcon();
            }

            private void FolderEntry_PropertyChanged(object sender, PropertyChangedEventArgs e) {
                if(Folders.Count > 0 && sender == Folders.First()) {
                    RefreshIcon();
                }
            }

            private void RefreshIcon() {
                Icon = Folders.Count == 0 ? QTUtility.ImageListGlobal.Images["folder"] : Folders.First().Icon;
            }

            public GroupEntry(string name, Keys shortcutKey, bool startup, IEnumerable<FolderEntry> folders) {
                Name = name;
                Startup = startup;
                ShortcutKey = shortcutKey;
                Folders = new ParentedCollection<FolderEntry>(folders);
                Folders.CollectionChanged += Folders_CollectionChanged;
                RefreshIcon();
            }

            public GroupEntry(string name) {
                Name = name;
                Folders = new ParentedCollection<FolderEntry>();
                Folders.CollectionChanged += Folders_CollectionChanged;
                RefreshIcon();
            }

            public GroupEntry() {
                Folders.CollectionChanged += Folders_CollectionChanged;
                RefreshIcon();
            }
        }

        private class AppEntry : INotifyPropertyChanged, IEditableEntry, ITreeViewItem {
            public event PropertyChangedEventHandler PropertyChanged;
            public IList ParentList { get; set; }
            public ParentedCollection<AppEntry> Children { get; set; }
            public bool IsFolder { get { return Children != null; } }
            public bool IsEditing { get; set; }
            public bool IsSelected { get; set; }
            public bool IsExpanded { get; set; }
            public string Name { get; set; }
            public string Path { get; set; }
            public string Args { get; set; }
            public string WorkingDir { get; set; }

            public Image Icon {
                get {
                    return IsFolder
                      ? QTUtility.ImageListGlobal.Images["folder"]
                      : QTUtility.GetIcon(Path, false).ToBitmap();
                }
            }

            public AppEntry(bool folder, string nameOrPath) {
                if(folder) {
                    Name = nameOrPath;
                    Children = new ParentedCollection<AppEntry>();    
                }
                else {
                    Path = nameOrPath;
                    Name = System.IO.Path.GetFileNameWithoutExtension(Path);
                }
            }
        }

        // ReSharper restore UnusedAutoPropertyAccessor.Local
        // ReSharper restore UnusedMember.Local
        // ReSharper restore MemberCanBePrivate.Local
        #pragma warning restore 0067
        #endregion

        // Common Font Chooser button click handler.
        private void btnFontChoose_Click(object sender, RoutedEventArgs e) {
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

        private sealed class ColorDialogEx : System.Windows.Forms.ColorDialog {
            protected override int Options {
                get {
                    return (base.Options | 2);
                }
            }
        }

        private void btnRecentFilesClear_Click(object sender, RoutedEventArgs e) {
            // TODO
        }

        private void btnRecentTabsClear_Click(object sender, RoutedEventArgs e) {
            // TODO
        }

        private void btnUpdateNow_Click(object sender, RoutedEventArgs e) {
            // TODO
        }
    }

    #region ---------- Converters ----------

    // Inverts the value of a boolean
    public class BoolInverterConverter : IValueConverter {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture) {
            return value is bool ? !(bool)value : value;
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture) {
            return value is bool ? !(bool)value : value;
        }
    }

    // Converts between booleans and one using logical and.
    public class LogicalAndMultiConverter : IMultiValueConverter {
        public object Convert(object[] values, Type targetType, object parameter, CultureInfo culture) {
            return values.All(b => b is bool && (bool)b);
        }

        public object[] ConvertBack(object value, Type[] targetTypes, object parameter, CultureInfo culture) {
            return new object[] { value };
        }
    }

    // Converts between many booleans and a string by StringJoining them.
    public class BoolJoinMultiConverter : IMultiValueConverter {
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
    public class StringEqualityConverter : IValueConverter {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture) {
            return (string)parameter == (string)value;
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture) {
            return (bool)value ? parameter : Binding.DoNothing;
        }
    }

    // Converts Bitmaps to ImageSources.
    [ValueConversion(typeof(Bitmap), typeof(ImageSource))]
    public class BitmapToImageSourceConverter : IValueConverter {
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

    [ValueConversion(typeof(System.Drawing.Color), typeof(Brush))]
    public class ColorToBrushConverter : IValueConverter {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture) {
            var c = (System.Drawing.Color)(value ?? System.Drawing.Color.Red);
            return new SolidColorBrush(Color.FromArgb(c.A, c.R, c.G, c.B));
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture) {
            throw new NotSupportedException();
        }
    }

    [ValueConversion(typeof(Font), typeof(string))]
    public class FontStringConverter : IValueConverter {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture) {
            if(value == null) return "";
            Font font = (Font)value;
            return string.Format("{0}, {1} pt", font.Name, Math.Round(font.SizeInPoints));
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture) {
            throw new NotSupportedException();
        }
    }

    /// <summary>
    /// Convert Level to left margin
    /// Pass a prarameter if you want a unit length other than 19.0.
    /// </summary>
    public class LevelToIndentConverter : IValueConverter {
        public object Convert(object o, Type type, object parameter,
                              CultureInfo culture) {
            return new Thickness((int)o * c_IndentSize, 0, 0, 0);
        }

        public object ConvertBack(object o, Type type, object parameter,
                                  CultureInfo culture) {
            throw new NotSupportedException();
        }

        private const double c_IndentSize = 19.0;
    }

    // You can create ObjectToTypeConverter instead of this,
    // but VS2010 WPF Designer would refuse an expression like {x:Type SomeClass+NestedClass}
    [ValueConversion(typeof(object), typeof(string))]
    public class ObjectToClassNameConverter : IValueConverter {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture) {
            if(value == null) {
                return null;
            }
            return value.GetType().Name;
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture) {
            throw new NotSupportedException();
        }
    }

    #endregion

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
}
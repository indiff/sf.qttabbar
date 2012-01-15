//    This file is part of QTTabBar, a shell extension for Microsoft
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
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Data;
using System.Windows.Input;
using System.Windows.Threading;

namespace QTTabBarLib {
    internal partial class Options07_Mouse : OptionsDialogTab {

        // todo: localize all of these
        #region ToLocalize
        
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
            {BindAction.Nothing,                "Do nothing"},
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
                BindAction.Nothing,
                BindAction.GoBack,
                BindAction.GoFirst,
                BindAction.GoForward,
                BindAction.GoLast,
                BindAction.NextTab,
                BindAction.PreviousTab
            }.ToDictionary(k => k, k => MouseActionItems[k])},
            {MouseTarget.Tab, new BindAction[] {
                BindAction.Nothing,
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
                BindAction.Nothing,
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
                BindAction.Nothing,
                BindAction.ItemOpenInNewTab,
                BindAction.ItemOpenInNewTabNoSel,
                BindAction.ItemOpenInNewWindow,
                BindAction.ItemProperties,
                BindAction.CopyItemPath,
                BindAction.CopyItemName,
            }.ToDictionary(k => k, k => MouseActionItems[k])},
            {MouseTarget.ExplorerItem, new BindAction[] {
                BindAction.Nothing,
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
                BindAction.Nothing,
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

        private static Dictionary<MouseButton, MouseChord> MouseButtonMappings =
                new Dictionary<MouseButton, MouseChord> {
                    { MouseButton.Left, MouseChord.Left },
                    { MouseButton.Right, MouseChord.Right },
                    { MouseButton.Middle, MouseChord.Middle },
                    { MouseButton.XButton1, MouseChord.X1 },
                    { MouseButton.XButton2, MouseChord.X2 }
        };

        #endregion

        private ObservableCollection<MouseEntry> MouseBindings;
        private DispatcherTimer mouseTimer;

        public Options07_Mouse() {
            InitializeComponent();
        }

        public override void InitializeConfig() {
            MouseBindings = new ObservableCollection<MouseEntry>();
            foreach(var p in WorkingConfig.mouse.GlobalMouseActions) {
                MouseBindings.Add(new MouseEntry(MouseTarget.Anywhere, p.Key, p.Value));
            }
            foreach(var p in WorkingConfig.mouse.MarginActions) {
                MouseBindings.Add(new MouseEntry(MouseTarget.ExplorerBackground, p.Key, p.Value));
            }
            foreach(var p in WorkingConfig.mouse.ItemActions) {
                MouseBindings.Add(new MouseEntry(MouseTarget.ExplorerItem, p.Key, p.Value));
            }
            foreach(var p in WorkingConfig.mouse.LinkActions) {
                MouseBindings.Add(new MouseEntry(MouseTarget.FolderLink, p.Key, p.Value));
            }
            foreach(var p in WorkingConfig.mouse.TabActions) {
                MouseBindings.Add(new MouseEntry(MouseTarget.Tab, p.Key, p.Value));
            }
            foreach(var p in WorkingConfig.mouse.BarActions) {
                MouseBindings.Add(new MouseEntry(MouseTarget.TabBarBackground, p.Key, p.Value));
            }
            ICollectionView view = CollectionViewSource.GetDefaultView(MouseBindings);
            PropertyGroupDescription groupDescription = new PropertyGroupDescription("TargetText");
            foreach(MouseTarget target in Enum.GetValues(typeof(MouseTarget))) {
                groupDescription.GroupNames.Add(MouseTargetItems[target]);
            }
            view.GroupDescriptions.Add(groupDescription);
            lvwMouseBindings.ItemsSource = view;
        }

        public override void ResetConfig() {
            DataContext = WorkingConfig.mouse = new Config._Mouse();
            InitializeConfig();
        }

        public override void CommitConfig() {
            WorkingConfig.mouse.GlobalMouseActions = MouseBindings
                    .Where(e => e.Action != BindAction.Nothing && e.Target == MouseTarget.Anywhere)
                    .ToDictionary(e => e.Chord, e => e.Action);
            WorkingConfig.mouse.MarginActions = MouseBindings
                    .Where(e => e.Action != BindAction.Nothing && e.Target == MouseTarget.ExplorerBackground)
                    .ToDictionary(e => e.Chord, e => e.Action);
            WorkingConfig.mouse.ItemActions = MouseBindings
                    .Where(e => e.Action != BindAction.Nothing && e.Target == MouseTarget.ExplorerItem)
                    .ToDictionary(e => e.Chord, e => e.Action);
            WorkingConfig.mouse.LinkActions = MouseBindings
                    .Where(e => e.Action != BindAction.Nothing && e.Target == MouseTarget.FolderLink)
                    .ToDictionary(e => e.Chord, e => e.Action);
            WorkingConfig.mouse.TabActions = MouseBindings
                    .Where(e => e.Action != BindAction.Nothing && e.Target == MouseTarget.Tab)
                    .ToDictionary(e => e.Chord, e => e.Action);
            WorkingConfig.mouse.BarActions = MouseBindings
                    .Where(e => e.Action != BindAction.Nothing && e.Target == MouseTarget.TabBarBackground)
                    .ToDictionary(e => e.Chord, e => e.Action);
        }


        private void rctAddMouseAction_MouseDown(object sender, MouseButtonEventArgs e) {
            FrameworkElement control = ((FrameworkElement)sender);
            MouseChord chord = MouseChord.None;
            if((Keyboard.Modifiers & ModifierKeys.Shift) == ModifierKeys.Shift) {
                chord |= MouseChord.Shift;
            }
            if((Keyboard.Modifiers & ModifierKeys.Control) == ModifierKeys.Control) {
                chord |= MouseChord.Ctrl;
            }
            if((Keyboard.Modifiers & ModifierKeys.Alt) == ModifierKeys.Alt) {
                chord |= MouseChord.Alt;
            }
            // ugh.  wish there was a better way to do this, but I don't think there is one...
            MouseTarget target = MouseTargetItems.First(kv => kv.Value == (string)control.Tag).Key;

            // watch out for double clicks
            if(e.ChangedButton == MouseButton.Left) {
                if(mouseTimer == null || mouseTimer.Tag != control.Tag) {
                    mouseTimer = new DispatcherTimer {
                        Tag = control.Tag,
                        Interval = TimeSpan.FromMilliseconds(
                                System.Windows.Forms.SystemInformation.DoubleClickTime)
                    };
                    mouseTimer.Tick += (sender2, e2) => {
                        mouseTimer.IsEnabled = false;
                        mouseTimer = null;
                        chord |= MouseChord.Left;
                        AddMouseAction(chord, target);
                    };
                    mouseTimer.IsEnabled = true;
                }
                else {
                    mouseTimer.IsEnabled = false;
                    mouseTimer = null;
                    chord |= MouseChord.Double;
                    AddMouseAction(chord, target);
                }
            }
            else {
                if(mouseTimer != null) {
                    mouseTimer.IsEnabled = false;
                    mouseTimer = null;
                }
                chord |= MouseButtonMappings[e.ChangedButton];
                AddMouseAction(chord, target);
            }
        }

        private void AddMouseAction(MouseChord chord, MouseTarget target) {
            MouseChord button = chord & ~(MouseChord.Alt | MouseChord.Ctrl | MouseChord.Shift);
            if(!MouseTargetButtons[target].ContainsKey(button)) {
                MessageBox.Show(
                        "This mouse button cannot be used for this target.  " +
                        "The valid mouse buttons for this target are:\r\n" +
                        MouseTargetButtons[target].Keys.Select(k => MouseButtonItems[k]).StringJoin("\r\n"),
                        "Invalid Button", MessageBoxButton.OK, MessageBoxImage.Hand);
                return;
            }
            MouseEntry entry = MouseBindings.FirstOrDefault(e => e.Chord == chord && e.Target == target);
            if(entry == null) {
                entry = new MouseEntry(target, chord, BindAction.Nothing);
                MouseBindings.Add(entry);
            }
            entry.IsSelected = true;
            lvwMouseBindings.UpdateLayout();
            lvwMouseBindings.ScrollIntoView(entry);
            // Need to wait for ScrollIntoView to finish, or the dropdown will open in the wrong place.
            Dispatcher.BeginInvoke(DispatcherPriority.Input, new Action(() => { entry.IsEditing = true; }));
        }

        private void lvwMouseBindings_KeyDown(object sender, KeyEventArgs e) {
            MouseEntry entry = lvwMouseBindings.SelectedItem as MouseEntry;
            if(entry == null) return;
            if(e.Key == Key.Delete) {
                MouseBindings.Remove(entry);
            }
            else if(e.Key == Key.Space || e.Key == Key.Enter) {
                entry.IsEditing = true;
            }
        }

        private void cmbInlineMouseAction_Loaded(object sender, RoutedEventArgs e) {
            // For some reason, SelectedValue gets wonky when the config is reinitialized.
            // This seems to fix it.
            ((ComboBox)sender).GetBindingExpression(Selector.SelectedValueProperty).UpdateTarget();
        }

        #region ---------- Binding Classes ----------
        // INotifyPropertyChanged is implemented automatically by Notify Property Weaver!
        #pragma warning disable 0067 // "The event 'PropertyChanged' is never used"
        // ReSharper disable MemberCanBePrivate.Local
        // ReSharper disable UnusedMember.Local
        // ReSharper disable UnusedAutoPropertyAccessor.Local

        private class MouseEntry : INotifyPropertyChanged {
            public event PropertyChangedEventHandler PropertyChanged;
            private bool isSelected;
            public bool IsSelected {
                get {
                    return isSelected;
                }
                set {
                    isSelected = value;
                    if(!isSelected) IsEditing = false;
                }
            }
            private bool isEditing;
            public bool IsEditing {
                get {
                    return isEditing;
                }
                set {
                    isEditing = value;
                    if(isEditing) IsSelected = true;
                }
            }
            public Dictionary<BindAction, string> ComboBoxItems {
                get { return MouseTargetActions[Target]; }
            }
            public string GestureText {
                get {
                    string ret = "";
                    foreach(var mod in new MouseChord[] { MouseChord.Ctrl, MouseChord.Shift, MouseChord.Alt }) {
                        if((Chord & mod) == mod) {
                            ret += MouseModifierItems[mod] + " + ";
                        }
                    }
                    MouseChord button = Chord & ~(MouseChord.Alt | MouseChord.Ctrl | MouseChord.Shift);
                    return ret + MouseButtonItems[button];
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

        #endregion
    }
}

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
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Reflection;
using System.Runtime.Serialization.Json;
using System.Text;
using System.Linq;
using System.Windows.Forms;
using Microsoft.Win32;
using Padding = System.Windows.Forms.Padding;
using Key = System.Windows.Forms.Keys;

namespace QTTabBarLib {

     // Wrapper class to get around  Font serialization stupidity
    [Serializable]
    public class XmlSerializableFont {
        public string FontName { get; set; }
        public float FontSize { get; set; }
        public FontStyle FontStyle { get; set; }

        public static XmlSerializableFont FromFont(Font font) {
            return font == null ? null : new XmlSerializableFont
                    {FontName = font.Name, FontSize = font.Size, FontStyle = font.Style};
        }

        public Font ToFont() {
            return ToFont(this);
        }

        public static Font ToFont(XmlSerializableFont xmlSerializableFont) {
            return new Font(
                    xmlSerializableFont.FontName,
                    xmlSerializableFont.FontSize,
                    xmlSerializableFont.FontStyle);
        }
    }

    public enum TabPos {
        Rightmost,
        Right,
        Left,
        Leftmost,
        LastActive,
    }

    public enum StretchMode {
        Full,
        Real,
        Tile,
    }

    public enum MouseTarget {
        Anywhere,
        Tab,
        TabBarBackground,
        FolderLink,
        ExplorerItem,
        ExplorerBackground
    }

    [Flags]
    public enum MouseChord {
        None    =   0,
        Shift   =   1,
        Ctrl    =   2,
        Alt     =   4,
        Left    =   8,
        Right   =  16,
        Middle  =  32,
        Double  =  64,
        X1      = 128,
        X2      = 256,
    }

    // WARNING
    // reordering these will break existing settings.
    public enum BindAction {
        GoBack = 0,
        GoForward,
        GoFirst,
        GoLast,
        NextTab,
        PreviousTab,
        FirstTab,
        LastTab,
        SwitchToLastActivated,
        NewTab,
        NewWindow,
        MergeWindows,
        CloseCurrent,
        CloseAllButCurrent,
        CloseLeft,
        CloseRight,
        CloseWindow,
        RestoreLastClosed,
        CloneCurrent,
        TearOffCurrent,
        LockCurrent,
        LockAll,
        BrowseFolder,
        CreateNewGroup,
        ShowOptions,
        ShowToolbarMenu,
        ShowTabMenuCurrent,
        ShowGroupMenu,
        ShowUserAppsMenu,
        ShowRecentTabsMenu,
        ShowRecentFilesMenu,
        NewFile,
        NewFolder,
        CopySelectedPaths,
        CopySelectedNames,
        CopyCurrentFolderPath,
        CopyCurrentFolderName,
        ChecksumSelected,
        ToggleTopMost,
        TransparencyPlus,
        TransparencyMinus,
        FocusFileList,
        FocusSearchBarReal,
        FocusSearchBarBBar,
        ShowSDTSelected,
        SendToTray,
        FocusTabBar,
        SortTabsByName,
        SortTabsByPath,
        SortTabsByActive,
        
        KEYBOARD_ACTION_COUNT,
        // Mouse-only from here on down

        Nothing = QTUtility.FIRST_MOUSE_ONLY_ACTION,
        UpOneLevel,
        Refresh,
        Paste,
        Maximize,
        Minimize,

        // Item Actions
        ItemOpenInNewTab,
        ItemOpenInNewTabNoSel,
        ItemOpenInNewWindow,
        ItemCut,
        ItemCopy,        
        ItemDelete,
        ItemProperties,
        CopyItemPath,
        CopyItemName,
        ChecksumItem,

        // Tab Actions
        CloseTab,
        CloseLeftTab,
        CloseRightTab,
        UpOneLevelTab, //hmm
        LockTab,
        ShowTabMenu,
        TearOffTab,
        CloneTab,
        CopyTabPath,
        TabProperties,
        ShowTabSubfolderMenu,
        CloseAllButThis,
    }

    [Serializable]
    public class Config {

        // Shortcuts to the loaded config, for convenience.
        public static _Window Window    { get { return ConfigManager.LoadedConfig.window; } }
        public static _Tabs Tabs        { get { return ConfigManager.LoadedConfig.tabs; } }
        public static _Tweaks Tweaks    { get { return ConfigManager.LoadedConfig.tweaks; } }
        public static _Tips Tips        { get { return ConfigManager.LoadedConfig.tips; } }
        public static _Misc Misc        { get { return ConfigManager.LoadedConfig.misc; } }
        public static _Skin Skin        { get { return ConfigManager.LoadedConfig.skin; } }
        public static _BBar BBar        { get { return ConfigManager.LoadedConfig.bbar; } }
        public static _Mouse Mouse      { get { return ConfigManager.LoadedConfig.mouse; } }
        public static _Keys Keys        { get { return ConfigManager.LoadedConfig.keys; } }
        public static _Plugin Plugin    { get { return ConfigManager.LoadedConfig.plugin; } }
        public static _Lang Lang        { get { return ConfigManager.LoadedConfig.lang; } }

        public _Window window   { get; set; }
        public _Tabs tabs       { get; set; }
        public _Tweaks tweaks   { get; set; }
        public _Tips tips       { get; set; }
        public _Misc misc       { get; set; }
        public _Skin skin       { get; set; }
        public _BBar bbar       { get; set; }
        public _Mouse mouse     { get; set; }
        public _Keys keys       { get; set; }
        public _Plugin plugin   { get; set; }
        public _Lang lang       { get; set; }

        public Config() {
            window = new _Window();
            tabs = new _Tabs();
            tweaks = new _Tweaks();
            tips = new _Tips();
            misc = new _Misc();
            skin = new _Skin();
            bbar = new _BBar();
            mouse = new _Mouse();
            keys = new _Keys();
            plugin = new _Plugin();
            lang = new _Lang();
        }

        [Serializable]
        public class _Window {
            public bool CaptureNewWindows        { get; set; }
            public bool RestoreSession           { get; set; }
            public bool RestoreOnlyLocked        { get; set; }
            public bool CloseBtnClosesUnlocked   { get; set; }
            public bool CloseBtnClosesSingleTab  { get; set; }
            public bool TrayOnClose              { get; set; }
            public bool TrayOnMinimize           { get; set; }
            public byte[] DefaultLocation        { get; set; }

            public _Window() {
                CaptureNewWindows = false;
                RestoreSession = false;
                RestoreOnlyLocked = false;
                CloseBtnClosesSingleTab = true;
                CloseBtnClosesUnlocked = false;
                TrayOnClose = false;
                TrayOnMinimize = false;

                string idl = Environment.OSVersion.Version >= new Version(6, 1)
                        ? "::{031E4825-7B94-4DC3-B131-E946B44C8DD5}"  // Libraries
                        : "::{20D04FE0-3AEA-1069-A2D8-08002B30309D}"; // Computer
                using(IDLWrapper w = new IDLWrapper(idl)) {
                    DefaultLocation = w.IDL;
                }
            }
        }

        [Serializable]
        public class _Tabs {
            public TabPos NewTabPosition         { get; set; }
            public TabPos NextAfterClosed        { get; set; }
            public bool ActivateNewTab           { get; set; }
            public bool NeverOpenSame            { get; set; }
            public bool RenameAmbTabs            { get; set; }
            public bool DragOverTabOpensSDT      { get; set; }
            public bool ShowFolderIcon           { get; set; }
            public bool ShowSubDirTipOnTab       { get; set; }
            public bool ShowDriveLetters         { get; set; }
            public bool ShowCloseButtons         { get; set; }
            public bool CloseBtnsWithAlt         { get; set; }
            public bool CloseBtnsOnHover         { get; set; }
            public bool ShowNavButtons           { get; set; }
            public bool NavButtonsOnRight        { get; set; }
            public bool MultipleTabRows          { get; set; }
            public bool ActiveTabOnBottomRow     { get; set; }

            public _Tabs() {
                NewTabPosition = TabPos.Rightmost;
                NextAfterClosed = TabPos.LastActive;
                ActivateNewTab = true;
                NeverOpenSame = true;
                RenameAmbTabs = false;
                DragOverTabOpensSDT = false;
                ShowFolderIcon = true;
                ShowSubDirTipOnTab = true;
                ShowDriveLetters = false;
                ShowCloseButtons = false;
                CloseBtnsWithAlt = false;
                CloseBtnsOnHover = false;
                ShowNavButtons = false;
                MultipleTabRows = true;
                ActiveTabOnBottomRow = true;
            }
        }

        [Serializable]
        public class _Tweaks {
            public bool AlwaysShowHeaders        { get; set; }
            public bool KillExtWhileRenaming     { get; set; }
            public bool RedirectLibraryFolders   { get; set; }
            public bool F2Selection              { get; set; }
            public bool WrapArrowKeySelection    { get; set; }
            public bool BackspaceUpLevel         { get; set; }
            public bool HorizontalScroll         { get; set; }
            public bool ForceSysListView         { get; set; }
            public bool ToggleFullRowSelect      { get; set; }
            public bool DetailsGridLines         { get; set; }
            public bool AlternateRowColors       { get; set; }
            public Color AltRowBackgroundColor   { get; set; }
            public Color AltRowForegroundColor   { get; set; }

            public _Tweaks() {
                AlwaysShowHeaders = !QTUtility.IsXP && !QTUtility.IsWin7;
                KillExtWhileRenaming = true;
                RedirectLibraryFolders = false;
                F2Selection = true;
                WrapArrowKeySelection = false;
                BackspaceUpLevel = QTUtility.IsXP;
                HorizontalScroll = true;
                ForceSysListView = false;
                ToggleFullRowSelect = false;
                DetailsGridLines = false;
                AlternateRowColors = false;
                AltRowForegroundColor = SystemColors.WindowText;
                AltRowBackgroundColor = QTUtility2.MakeColor(0xfaf5f1);
            }
        }

        [Serializable]
        public class _Tips {
            public bool ShowSubDirTips           { get; set; }
            public bool SubDirTipsPreview        { get; set; }
            public bool SubDirTipsFiles          { get; set; }
            public bool SubDirTipsWithShift      { get; set; }
            public bool ShowTooltipPreviews      { get; set; }
            public bool ShowPreviewsWithShift    { get; set; }
            public bool ShowPreviewInfo          { get; set; }
            public int PreviewMaxWidth           { get; set; }
            public int PreviewMaxHeight          { get; set; }
            public Font PreviewFont              { get; set; }
            public List<string> TextExt          { get; set; }
            public List<string> ImageExt         { get; set; }
            
            public _Tips() {
                ShowSubDirTips = true;
                SubDirTipsPreview = true;
                SubDirTipsFiles = true;
                SubDirTipsWithShift = false;
                ShowTooltipPreviews = true;
                ShowPreviewsWithShift = false;
                ShowPreviewInfo = true;
                PreviewMaxWidth = 512;
                PreviewMaxHeight = 256;
                PreviewFont = Control.DefaultFont;
                TextExt = new List<string> {".txt", ".ini", ".inf" ,".cs", ".log", ".js", ".vbs"};
                ImageExt = ThumbnailTooltipForm.MakeDefaultImgExts();
            }
        }

        [Serializable]
        public class _Misc {
            public bool TaskbarThumbnails        { get; set; }
            public bool KeepHistory              { get; set; }
            public int TabHistoryCount           { get; set; }
            public bool KeepRecentFiles          { get; set; }
            public int FileHistoryCount          { get; set; }
            public int NetworkTimeout            { get; set; }
            public bool AutoUpdate               { get; set; }

            public _Misc() {
                TaskbarThumbnails = false;
                KeepHistory = true;
                TabHistoryCount = 15;
                KeepRecentFiles = true;
                FileHistoryCount = 15;
                NetworkTimeout = 0;
                AutoUpdate = true;
            }
        }

        [Serializable]
        public class _Skin {
            public bool UseTabSkin               { get; set; }
            public string TabImageFile           { get; set; }
            public Padding TabSizeMargin         { get; set; }
            public Padding TabContentMargin      { get; set; }
            public int OverlapPixels             { get; set; }
            public bool HitTestTransparent       { get; set; }
            public int TabHeight                 { get; set; }
            public int TabMinWidth               { get; set; }
            public int TabMaxWidth               { get; set; }
            public bool FixedWidthTabs           { get; set; }
            public Font TabTextFont              { get; set; }
            public Color TabTextActiveColor      { get; set; }
            public Color TabTextInactiveColor    { get; set; }
            public Color TabTextHotColor         { get; set; }
            public Color TabShadActiveColor      { get; set; }
            public Color TabShadInactiveColor    { get; set; }
            public Color TabShadHotColor         { get; set; }
            public bool TabTitleShadows          { get; set; }
            public bool TabTextCentered          { get; set; }
            public bool UseRebarBGColor          { get; set; }
            public Color RebarColor              { get; set; }
            public bool UseRebarImage            { get; set; }
            public StretchMode RebarStretchMode  { get; set; }
            public string RebarImageFile         { get; set; }
            public bool RebarImageSeperateBars   { get; set; }
            public Padding RebarSizeMargin       { get; set; }
            public bool ActiveTabInBold          { get; set; }

            public _Skin() {
                UseTabSkin = false;
                TabImageFile = "";
                TabSizeMargin = Padding.Empty;
                TabContentMargin = Padding.Empty;
                OverlapPixels = 0;
                HitTestTransparent = false;
                TabHeight = 24;
                TabMinWidth = 50;
                TabMaxWidth = 200;
                FixedWidthTabs = false;
                TabTextFont = Control.DefaultFont;
                TabTextActiveColor = Color.Black;
                TabTextInactiveColor = Color.Black;
                TabTextHotColor = Color.Black;
                TabShadActiveColor = Color.Gray;
                TabShadInactiveColor = Color.White;
                TabShadHotColor = Color.White;
                TabTitleShadows = false;
                TabTextCentered = false;
                UseRebarBGColor = false;
                RebarColor = Color.Gray;
                UseRebarImage = false;
                RebarStretchMode = StretchMode.Full;
                RebarImageFile = "";
                RebarImageSeperateBars = false;
                RebarSizeMargin = Padding.Empty;
                ActiveTabInBold = false;
            }
        }

        [Serializable]
        public class _BBar {
            public int[] ButtonIndexes           { get; set; }
            public string[] ActivePluginIDs      { get; set; }
            public bool LargeButtons             { get; set; }
            public bool LockSearchBarWidth       { get; set; }
            public bool LockDropDownButtons      { get; set; }
            public bool ShowButtonLabels         { get; set; }
            public string ImageStripPath         { get; set; }
            
            public _BBar() {
                ButtonIndexes = QTUtility.IsXP 
                        ? new int[] {1, 2, 0, 3, 4, 5, 0, 6, 7, 0, 11, 13, 12, 14, 15, 0, 9, 20} 
                        : new int[] {3, 4, 5, 0, 6, 7, 0, 11, 13, 12, 14, 15, 0, 9, 20};
                ActivePluginIDs = new string[0];
                LockDropDownButtons = false;
                LargeButtons = true;
                LockSearchBarWidth = false;
                ShowButtonLabels = false;
                ImageStripPath = "";
            }
        }

        [Serializable]
        public class _Mouse {
            public bool MouseScrollsHotWnd       { get; set; }
            public Dictionary<MouseChord, BindAction> GlobalMouseActions { get; set; }
            public Dictionary<MouseChord, BindAction> TabActions { get; set; }
            public Dictionary<MouseChord, BindAction> BarActions { get; set; }
            public Dictionary<MouseChord, BindAction> LinkActions { get; set; }
            public Dictionary<MouseChord, BindAction> ItemActions { get; set; }
            public Dictionary<MouseChord, BindAction> MarginActions { get; set; }

            public _Mouse() {
                MouseScrollsHotWnd = false;
                GlobalMouseActions = new Dictionary<MouseChord, BindAction> {
                    {MouseChord.X1, BindAction.GoBack},
                    {MouseChord.X2, BindAction.GoForward},
                    {MouseChord.X1 | MouseChord.Ctrl, BindAction.GoFirst},
                    {MouseChord.X2 | MouseChord.Ctrl, BindAction.GoLast}
                };
                TabActions = new Dictionary<MouseChord, BindAction> { 
                    {MouseChord.Middle, BindAction.CloseTab},
                    {MouseChord.Ctrl | MouseChord.Left, BindAction.LockTab},
                    {MouseChord.Double, BindAction.UpOneLevelTab},
                };
                BarActions = new Dictionary<MouseChord, BindAction> {
                    {MouseChord.Double, BindAction.NewTab},
                    {MouseChord.Middle, BindAction.RestoreLastClosed}
                };
                LinkActions = new Dictionary<MouseChord, BindAction> {
                    {MouseChord.Middle, BindAction.ItemOpenInNewTab},
                    {MouseChord.Ctrl | MouseChord.Middle, BindAction.ItemOpenInNewWindow}
                };
                ItemActions = new Dictionary<MouseChord, BindAction> {
                    {MouseChord.Middle, BindAction.ItemOpenInNewTab},
                    {MouseChord.Ctrl | MouseChord.Middle, BindAction.ItemOpenInNewWindow}                        
                };
                MarginActions = new Dictionary<MouseChord, BindAction> {
                    {MouseChord.Double, BindAction.UpOneLevel}
                };
            }
        }

        [Serializable]
        public class _Keys {
            public int[] Shortcuts               { get; set; }
            public Dictionary<string, int[]> PluginShortcuts { get; set; } 
            public bool UseTabSwitcher           { get; set; }

            public _Keys() {
                var dict = new Dictionary<BindAction, Keys> {
                    {BindAction.GoBack,             Key.Left  | Key.Alt},
                    {BindAction.GoForward,          Key.Right | Key.Alt},
                    {BindAction.GoFirst,            Key.Left  | Key.Control | Key.Alt},
                    {BindAction.GoLast,             Key.Right | Key.Control | Key.Alt},
                    {BindAction.NextTab,            Key.Tab   | Key.Control},
                    {BindAction.PreviousTab,        Key.Tab   | Key.Control | Key.Shift},
                    {BindAction.NewTab,             Key.T     | Key.Control},
                    {BindAction.NewWindow,          Key.T     | Key.Control | Key.Shift},
                    {BindAction.CloseCurrent,       Key.W     | Key.Control},
                    {BindAction.CloseAllButCurrent, Key.W     | Key.Control | Key.Shift},
                    {BindAction.RestoreLastClosed,  Key.Z     | Key.Control | Key.Shift},
                    {BindAction.LockCurrent,        Key.L     | Key.Control},
                    {BindAction.LockAll,            Key.L     | Key.Control | Key.Shift},
                    {BindAction.BrowseFolder,       Key.O     | Key.Control},
                    {BindAction.ShowOptions,        Key.O     | Key.Alt},
                    {BindAction.ShowToolbarMenu,    Key.Oemcomma  | Key.Alt},
                    {BindAction.ShowTabMenuCurrent, Key.OemPeriod | Key.Alt},
                    {BindAction.ShowGroupMenu,      Key.G     | Key.Alt},
                    {BindAction.ShowUserAppsMenu,   Key.H     | Key.Alt},
                    {BindAction.ShowRecentTabsMenu, Key.U     | Key.Alt},
                    {BindAction.ShowRecentFilesMenu,Key.F     | Key.Alt},
                    {BindAction.NewFile,            Key.N     | Key.Control},
                    {BindAction.NewFolder,          Key.N     | Key.Control | Key.Shift},
                };
                Shortcuts = new int[(int)BindAction.KEYBOARD_ACTION_COUNT];
                PluginShortcuts = new Dictionary<string, int[]>();
                foreach(var pair in dict) {
                    Shortcuts[(int)pair.Key] = (int)pair.Value | QTUtility.FLAG_KEYENABLED;
                }
                UseTabSwitcher = true;
            }
        }

        [Serializable]
        public class _Plugin {
            public string[] Enabled              { get; set; }

            public _Plugin() {
                Enabled = new string[0];
            }
        }

        [Serializable]
        public class _Lang {
            public string[] PluginLangFiles      { get; set; }
            public bool UseLangFile              { get; set; }
            public string LangFile               { get; set; }
            public string BuiltInLang            { get; set; }

            public _Lang() {
                UseLangFile = false;
                BuiltInLang = "English";
                LangFile = "";
                PluginLangFiles = new string[0];
            }
        }
    }

    public static class ConfigManager {
        public static Config LoadedConfig;

        public static void Initialize() {
            LoadedConfig = new Config();
            ReadConfig();

        }

        public static void UpdateConfig(bool fBroadcast = true) {
            QTUtility.TextResourcesDic = Config.Lang.UseLangFile && File.Exists(Config.Lang.LangFile)
                    ? QTUtility.ReadLanguageFile(Config.Lang.LangFile)
                    : null;
            QTUtility.ValidateTextResources();
            QTUtility.ClosedTabHistoryList.MaxCapacity = Config.Misc.TabHistoryCount;
            QTUtility.ExecutedPathsList.MaxCapacity = Config.Misc.FileHistoryCount;
            DropDownMenuBase.InitializeMenuRenderer();
            ContextMenuStripEx.InitializeMenuRenderer();
            PluginManager.RefreshPlugins();
            InstanceManager.LocalTabBroadcast(tabbar => tabbar.RefreshOptions());
            if(fBroadcast) {
                // SyncTaskBarMenu(); todo
                InstanceManager.StaticBroadcast(() => {
                    ReadConfig();
                    UpdateConfig(false);
                });
            }
        }

        public static void ReadConfig() {
            const string RegPath = RegConst.Root + RegConst.Config;

            // Properties from all categories
            foreach(PropertyInfo category in typeof(Config).GetProperties().Where(c => c.CanWrite)) {
                Type cls = category.PropertyType;
                object val = category.GetValue(LoadedConfig, null);
                using(RegistryKey key = Registry.CurrentUser.CreateSubKey(RegPath + cls.Name.Substring(1))) {
                    foreach(PropertyInfo prop in cls.GetProperties()) {
                        object obj = key.GetValue(prop.Name);
                        if(obj == null) continue;
                        Type t = prop.PropertyType;
                        try {
                            if(t == typeof(bool)) {
                                prop.SetValue(val, (int)obj != 0, null);
                            }
                            else if(t == typeof(int) || t == typeof(string)) {
                                prop.SetValue(val, obj, null);
                            }
                            else if(t.IsEnum) {
                                prop.SetValue(val, Enum.Parse(t, obj.ToString()), null);
                            }
                            else {
                                using(MemoryStream stream = new MemoryStream(Encoding.UTF8.GetBytes(obj.ToString()))) {
                                    if(t == typeof(Font)) {
                                        DataContractJsonSerializer ser = new DataContractJsonSerializer(typeof(XmlSerializableFont));
                                        XmlSerializableFont xsf = ser.ReadObject(stream) as XmlSerializableFont;
                                        prop.SetValue(val, xsf == null ? null : xsf.ToFont(), null);
                                    }
                                    else {
                                        DataContractJsonSerializer ser = new DataContractJsonSerializer(t);
                                        prop.SetValue(val, ser.ReadObject(stream), null);
                                    }
                                }
                            }
                        }
                        catch {
                        }
                    }
                }
            }

            using(IDLWrapper wrapper = new IDLWrapper(Config.Window.DefaultLocation)) {
                if(!wrapper.Available) {
                    Config.Window.DefaultLocation = new Config._Window().DefaultLocation;
                }
            }
            Config.Tips.PreviewFont = Config.Tips.PreviewFont ?? Control.DefaultFont;
            Config.Tips.PreviewMaxWidth = QTUtility.ValidateMinMax(Config.Tips.PreviewMaxWidth, 128, 1920);
            Config.Tips.PreviewMaxHeight = QTUtility.ValidateMinMax(Config.Tips.PreviewMaxHeight, 96, 1200);
            Config.Misc.TabHistoryCount = QTUtility.ValidateMinMax(Config.Misc.TabHistoryCount, 1, 30);
            Config.Misc.FileHistoryCount = QTUtility.ValidateMinMax(Config.Misc.FileHistoryCount, 1, 30);
            Config.Misc.NetworkTimeout = QTUtility.ValidateMinMax(Config.Misc.NetworkTimeout, 0, 120);
            Config.Skin.TabHeight = QTUtility.ValidateMinMax(Config.Skin.TabHeight, 10, 50);
            Config.Skin.TabMinWidth = QTUtility.ValidateMinMax(Config.Skin.TabMinWidth, 10, 50);
            Config.Skin.TabMaxWidth = QTUtility.ValidateMinMax(Config.Skin.TabMaxWidth, 50, 999);
            Config.Skin.OverlapPixels = QTUtility.ValidateMinMax(Config.Skin.TabHeight, 0, 20);
            Config.Skin.TabTextFont = Config.Skin.TabTextFont ?? Control.DefaultFont;
            Func<Padding, Padding> validatePadding = p => {
                p.Left   = QTUtility.ValidateMinMax(p.Left,   0, 99);
                p.Top    = QTUtility.ValidateMinMax(p.Top,    0, 99);
                p.Right  = QTUtility.ValidateMinMax(p.Right,  0, 99);
                p.Bottom = QTUtility.ValidateMinMax(p.Bottom, 0, 99);
                return p;
            };
            Config.Skin.RebarSizeMargin = validatePadding(Config.Skin.RebarSizeMargin);
            Config.Skin.TabContentMargin = validatePadding(Config.Skin.TabContentMargin);
            Config.Skin.TabSizeMargin = validatePadding(Config.Skin.TabSizeMargin);
            using(IDLWrapper wrapper = new IDLWrapper(Config.Skin.TabImageFile)) {
                if(!wrapper.Available) Config.Skin.TabImageFile = "";
            }
            using(IDLWrapper wrapper = new IDLWrapper(Config.Skin.RebarImageFile)) {
                if(!wrapper.Available) Config.Skin.RebarImageFile = "";
            }
            using(IDLWrapper wrapper = new IDLWrapper(Config.BBar.ImageStripPath)) {
                // todo: check dimensions
                if(!wrapper.Available) Config.BBar.ImageStripPath = "";
            }
            List<int> blist = Config.BBar.ButtonIndexes.ToList();
            blist.RemoveAll(i => (i.HiWord() - 1) >= Config.BBar.ActivePluginIDs.Length);
            Config.BBar.ButtonIndexes = blist.ToArray();
            var keys = Config.Keys.Shortcuts;
            Array.Resize(ref keys, (int)BindAction.KEYBOARD_ACTION_COUNT);
            Config.Keys.Shortcuts = keys;
            foreach(var pair in Config.Keys.PluginShortcuts.Where(p => p.Value == null).ToList()) {
                Config.Keys.PluginShortcuts.Remove(pair.Key);
            }
            if(QTUtility.IsXP) Config.Tweaks.AlwaysShowHeaders = false;
            if(!QTUtility.IsWin7) Config.Tweaks.RedirectLibraryFolders = false;
            if(!QTUtility.IsXP) Config.Tweaks.KillExtWhileRenaming = true;
            if(QTUtility.IsXP) Config.Tweaks.BackspaceUpLevel = true;
            if(!QTUtility.IsWin7) Config.Tweaks.ForceSysListView = true;
        }
        public static void WriteConfig() {
            const string RegPath = RegConst.Root + RegConst.Config;

            // Properties from all categories
            foreach(PropertyInfo category in typeof(Config).GetProperties().Where(c => c.CanWrite)) {
                Type cls = category.PropertyType;
                object val = category.GetValue(LoadedConfig, null);
                using(RegistryKey key = Registry.CurrentUser.CreateSubKey(RegPath + cls.Name.Substring(1))) {
                    foreach(var prop in cls.GetProperties()) {
                        Type t = prop.PropertyType;
                        if(t == typeof(bool)) {
                            key.SetValue(prop.Name, (bool)prop.GetValue(val, null) ? 1 : 0);
                        }
                        else if(t == typeof(int) || t == typeof(string) || t.IsEnum) {
                            key.SetValue(prop.Name, prop.GetValue(val, null));
                        }
                        else {
                            object obj = prop.GetValue(val, null);
                            if(t == typeof(Font)) {
                                obj = XmlSerializableFont.FromFont((Font)obj);
                                t = typeof(XmlSerializableFont);
                            }
                            DataContractJsonSerializer ser = new DataContractJsonSerializer(t);
                            using(MemoryStream stream = new MemoryStream()) {
                                try {
                                    ser.WriteObject(stream, obj);
                                }
                                catch(Exception e) {
                                    QTUtility2.MakeErrorLog(e);
                                }
                                stream.Position = 0;
                                StreamReader reader = new StreamReader(stream);
                                key.SetValue(prop.Name, reader.ReadToEnd());                                
                            }
                        }
                    }
                }
            }
        }
    }
}

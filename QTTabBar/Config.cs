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
using System.Linq;
using System.Text;

namespace QTTabBarLib {
    public static class Config {
        
        // Window
        public static bool NoTabsFromOutside { get; set; }
        public static bool RestoreTabs { get; set; }
        public static bool RestoreLockedTabs { get; set; }
        public static bool NeverCloseWndLocked { get; set; }
        public static bool NeverCloseWindow { get; set; }
        public static bool TrayOnClose { get; set; }
        public static bool TrayOnMinimize { get; set; }

        // Tabs
        // public static int NewTabPosition { get; set; }
        // public static int NextAfterClosed { get; set; }
        public static bool ActivateNewTab { get; set; }
        public static bool DontOpenSame { get; set; }
        public static bool NoRenameAmbTabs { get; set; }
        public static bool DragDropOntoTabs { get; set; }
        public static bool FolderIcon { get; set; }
        public static bool ShowSubDirTipOnTab { get; set; }
        public static bool ShowDriveLetters { get; set; }
        public static bool ShowTabCloseButtons { get; set; }
        public static bool TabCloseBtnsWithAlt { get; set; }
        public static bool TabCloseBtnsOnHover { get; set; }
        public static bool ShowNavButtons { get; set; }
        public static bool NavButtonsOnRight { get; set; }
        public static bool MultipleRow1 { get; set; }
        public static bool MultipleRow2 { get; set; }

        // Tweaks
        public static bool AlwaysShowHeaders { get; set; }
        public static bool ExtWhileRenaming { get; set; }
        public static bool F2Selection { get; set; }
        public static bool CursorLoop { get; set; }
        public static bool BackspaceUpLevel { get; set; }
        public static bool HorizontalScroll { get; set; }
        public static bool ForceSysListView { get; set; }
        public static bool ToggleFullRowSelect { get; set; }
        public static bool DetailsGridLines { get; set; }
        public static bool AlternateRowColors { get; set; }

        // Tooltips
        public static bool NoShowSubDirTips { get; set; }
        public static bool SubDirTipsPreview { get; set; } 
        public static bool SubDirTipsFiles { get; set; }
        public static bool SubDirTipsWithShift { get; set; }
        public static bool ShowTooltipPreviews { get; set; }
        public static bool PreviewsWithShift { get; set; }
        public static bool PreviewInfo { get; set; }
        // public static int PreviewWidth { get; set; }
        // public static int PreviewHeight { get; set; }
        // public static Font PreviewFont { get; set; }
        // Image/Text filetypes???

        // Misc
        // public static bool TaskbarThumbnails { get; set; }
        public static bool NoHistory { get; set; }
        // public static int TabHistoryCount { get; set; }
        public static bool NoRecentFiles { get; set; }
        // public static int FileHistoryCount { get; set; }
        // public static int NetworkTimeout { get; set; }
        public static bool AutoUpdate { get; set; }
        // public static bool UseIniFile { get; set; }

        // Appearence
        public static bool UseTabSkin { get; set; }
        // public static string TabImageFile { get; set; }
        // public static int TabSizeMarginL { get; set; }
        // public static int TabSizeMarginT { get; set; }
        // public static int TabSizeMarginR { get; set; }
        // public static int TabSizeMarginB { get; set; }
        // public static int TabHeight { get; set; }
        // public static int MinWidth { get; set; }
        // public static int MaxWidth { get; set; }
        public static bool FixedWidthTabs { get; set; }
        // public static Font TabTextFont { get; set; }
        // public static Font TabTextFont { get; set; }
        // public static Color TabTextActiveColor { get; set; }
        // public static Color TabTextInactiveColor { get; set; }
        // public static Color TabTextHotColor { get; set; }
        // public static Color TabShadActiveColor { get; set; }
        // public static Color TabShadInactiveColor { get; set; }
        // public static Color TabShadHotColor { get; set; }
        public static bool TabTitleShadows { get; set; }
        // public static int TabTextAlignment { get; set; }
        public static bool ToolbarBGColor { get; set; }
        // public static Color RebarColor { get; set; }
        public static bool UseRebarImage { get; set; }
        // public static int RebarStretchMode { get; set; }
        // public static string RebarImageFile { get; set; }
        public static bool RebarImageActual { get; set; }
        // public static int RebarSizeMarginL { get; set; }
        // public static int RebarSizeMarginT { get; set; }
        // public static int RebarSizeMarginR { get; set; }
        // public static int RebarSizeMarginB { get; set; }

        // Mouse
        // public static bool MouseScrollsHotWnd { get; set; }

        // Button Bar


        //-----------

        // Maybe
        public static bool ActiveTabInBold { get; set; }
        public static bool CtrlWheelChangeView { get; set; }
        public static bool ShowHashResult { get; set; }
        public static bool HashTopMost { get; set; }
        public static bool HashFullPath { get; set; }
        public static bool HashClearOnClose { get; set; }
        public static bool KeepOnSeparate { get; set; }        

        // DEATH ROW
        public static bool DontCaptureNewWnds { get; set; }
        public static bool AllRecentFiles { get; set; } 
        public static bool AlignTabTextCenter { get; set; }
        public static bool ShowTooltips        { get; set; }
        public static bool CloseWhenGroup      { get; set; }
        public static bool NoCaptureMidClick   { get; set; }        
        public static bool LimitedWidthTabs    { get; set; }        
        public static bool CaptureX1X2         { get; set; }
        public static bool NoWindowResizing    { get; set; }
        public static bool NoNewWndFolderTree  { get; set; }
        public static bool NoDblClickUpLevel   { get; set; }
        public static bool MidClickNewWindow   { get; set; }
        public static bool HideMenuBar         { get; set; }
        public static bool SaveTransparency    { get; set; }
        public static bool SubDirTipsHidden    { get; set; }
        public static bool SubDirTipsSystem    { get; set; }
        public static bool RebarImageTile      { get; set; }
        public static bool RebarImageStretch2  { get; set; }
        public static bool NoTabSwitcher       { get; set; }
        public static bool NoMidClickTree      { get; set; }
        public static bool XPStyleMenus        { get; set; }
        public static bool NonDefaultMenu      { get; set; }
        public static bool DisableSound        { get; set; }
        
    }
}

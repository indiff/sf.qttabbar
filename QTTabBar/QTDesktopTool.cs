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
using System.Drawing;
using System.IO;
using System.Text;
using System.Threading;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using System.Windows.Forms.VisualStyles;
using BandObjectLib;
using Microsoft.Win32;

using QTTabBarLib.Interop;


namespace QTTabBarLib {
    [ComVisible(true)]
    [Guid("D2BF470E-ED1C-487F-A555-2BD8835EB6CE")]
    public sealed class QTDesktopTool : BandObject, IDeskBand2 {
        // FUNCTIONS  									  |	THREAD		
        // -----------------------------------------------+-----------
        // taskbar toolbar								  | taskbar
        // Menu											  | taskbar (even if shown by double-click on destop)
        // *subDirTip_TB	(real directory menu in menu) |	taskbar
        //
        // Sub Directory Menu							  |	desktop
        // Thumbnail for desktop item					  |	desktop


        #region ---------- Fields ----------

        // PLATFORM INVOKE & INTEROPS
        private IntPtr hHook_MsgDesktop;
        private IntPtr hHook_MsgShell_TrayWnd;
        private IntPtr hHook_KeyDesktop;
        private HookProc hookProc_Msg_Desktop;
        private HookProc hookProc_Msg_ShellTrayWnd;
        private HookProc hookProc_Keys_Desktop;

        private IntPtr hwndListView, hwndShellView, hwndShellTray, hwndThis;

        private IContextMenu2 iContextMenu2; // for taskbar thread, events handled at WndProc
        private IFolderView folderView;
        private IShellBrowser shellBrowser;

        private NativeWindowListener shellViewListener;


        // CONTROLS
        private System.ComponentModel.IContainer components;
        private DropDownMenuReorderable contextMenu, ddmrGroups, ddmrHistory, ddmrUserapps, ddmrRecentFile;
        private TitleMenuItem tmiLabel_Group, tmiLabel_History, tmiLabel_UserApp, tmiLabel_RecentFile;
        private TitleMenuItem tmiGroup, tmiHistory, tmiUserApp, tmiRecentFile;

        private System.Windows.Forms.Timer timerHooks;
        private int iHookTimeout;

        private List<ToolStripItem> lstGroupItems = new List<ToolStripItem>();
        private List<ToolStripItem> lstUndoClosedItems = new List<ToolStripItem>();
        private List<ToolStripItem> lstUserAppItems = new List<ToolStripItem>();
        private List<ToolStripItem> lstRecentFileItems = new List<ToolStripItem>();

        private ContextMenuStripEx contextMenuForSetting;

        private ToolStripMenuItem tsmiTaskBar,
                                  tsmiDesktop,
                                  tsmiLockItems,
                                  tsmiVSTitle,
                                  tsmiOnGroup,
                                  tsmiOnHistory,
                                  tsmiOnUserApps,
                                  tsmiOnRecentFile,
                                  tsmiOneClick,
                                  tsmiAppKeys;

        private ToolStripMenuItem tsmiExperimental;

        private ThumbnailTooltipForm thumbnailTooltip;
        private System.Windows.Forms.Timer timer_Thumbnail, timer_ThumbnailMouseHover, timer_Thumbnail_NoTooltip;
        private bool fThumbnailPending;

        private SubDirTipForm subDirTip, subDirTip_TB;
        private IContextMenu2 iContextMenu2_Desktop; // for desktop thread, handled at shellViewListener_MessageCaptured
        private int thumbnailIndex = -1;
        private int thumbnailIndex_GETINFOTIP = -1;
        private int itemIndexDROPHILITED = -1;
        private int thumbnailIndex_Inactive = -1;
        private System.Windows.Forms.Timer timer_HoverSubDirTipMenu;

        private VisualStyleRenderer bgRenderer;
        private StringFormat stringFormat;


        // SETTINGS & FLAGS		
        private bool[] ExpandState = {false, false, false, false};
        private UniqueList<int> lstItemOrder = new UniqueList<int> {0, 1, 2, 3};
        private List<bool> lstRefreshRequired = new List<bool> {false, false, false, false};
        private int WidthOfBar = 80;

        private bool fCancelClosing;
        private bool fNowMouseHovering;


        // CONSTANTS
        private const int ITEMTYPE_COUNT = 4;
        private const int ITEMINDEX_GROUP = 0;
        private const int ITEMINDEX_RECENTTAB = 1;
        private const int ITEMINDEX_APPLAUNCHER = 2;
        private const int ITEMINDEX_RECENTFILE = 3;


        private const string TEXT_TOOLBAR = "QTTab Desktop Tool";
        private const string MENUKEY_LABEL_GROUP = "labelG";
        private const string MENUKEY_LABEL_HISTORY = "labelH";
        private const string MENUKEY_LABEL_USERAPP = "labelU";
        private const string MENUKEY_LABEL_RECENT = "labelR";

        private const string MENUKEY_ITEM_GROUP = "groupItem";
        private const string MENUKEY_ITEM_HISTORY = "historyItem";
        private const string MENUKEY_ITEM_USERAPP = "userappItem";
        private const string MENUKEY_ITEM_RECENT = "recentItem";

        private const string MENUKEY_SUBMENUS = "submenu";
        private const string MENUKEY_LABELS = "label";

        private const string TSS_NAME_GRP = "groupSep";
        private const string TSS_NAME_APP = "appSep";


        private const string CLSIDSTR_TRASHBIN = "::{645FF040-5081-101B-9F08-00AA002F954E}";


        #endregion



        public QTDesktopTool() {
            QTUtility.Initialize();
        }


        private void InitializeComponent() {
            // handle creation
            this.hwndThis = this.Handle;

            bool reorderEnabled = !Config.Bool(Scts.Desktop_LockMenu);

            this.components = new System.ComponentModel.Container();
            this.contextMenu = new DropDownMenuReorderable(this.components, true, false);
            this.contextMenuForSetting = new ContextMenuStripEx(this.components, true, false);
            this.tmiLabel_Group = new TitleMenuItem(MenuGenre.Group, true);
            this.tmiLabel_History = new TitleMenuItem(MenuGenre.RecentlyClosedTab, true);
            this.tmiLabel_UserApp = new TitleMenuItem(MenuGenre.Application, true);
            this.tmiLabel_RecentFile = new TitleMenuItem(MenuGenre.RecentFile, true);

            this.contextMenu.SuspendLayout();
            this.contextMenuForSetting.SuspendLayout();
            this.SuspendLayout();
            //
            // contextMenu
            //
            this.contextMenu.ProhibitedKey.Add(MENUKEY_ITEM_HISTORY);
            this.contextMenu.ProhibitedKey.Add(MENUKEY_ITEM_RECENT);
            this.contextMenu.ReorderEnabled = reorderEnabled;
            this.contextMenu.MessageParent = this.Handle;
            this.contextMenu.ImageList = QTUtility.ImageListGlobal;
            this.contextMenu.ItemClicked += new ToolStripItemClickedEventHandler(this.dropDowns_ItemClicked);
            this.contextMenu.Closing += new ToolStripDropDownClosingEventHandler(this.contextMenu_Closing);
            this.contextMenu.ReorderFinished += new MenuReorderedEventHandler(this.contextMenu_ReorderFinished);
            this.contextMenu.ItemRightClicked += new ItemRightClickedEventHandler(this.dropDowns_ItemRightClicked);
            if(QTUtility.IsVista) {
                IntPtr hwnd = this.contextMenu.Handle;
            }
            //
            // ddmrGroups 
            //
            this.ddmrGroups = new DropDownMenuReorderable(this.components, true, false);
            this.ddmrGroups.ReorderEnabled = reorderEnabled;
            this.ddmrGroups.ImageList = QTUtility.ImageListGlobal;
            this.ddmrGroups.ReorderFinished += new MenuReorderedEventHandler(this.dropDowns_ReorderFinished);
            this.ddmrGroups.ItemClicked += new ToolStripItemClickedEventHandler(this.dropDowns_ItemClicked);
            this.ddmrGroups.ItemRightClicked += new ItemRightClickedEventHandler(this.dropDowns_ItemRightClicked);
            //
            // tmiGroup 
            //
            this.tmiGroup = new TitleMenuItem(MenuGenre.Group, false);
            this.tmiGroup.DropDown = this.ddmrGroups;
            //
            // ddmrHistory
            //
            this.ddmrHistory = new DropDownMenuReorderable(this.components, true, false);
            this.ddmrHistory.ReorderEnabled = false;
            this.ddmrHistory.ImageList = QTUtility.ImageListGlobal;
            this.ddmrHistory.MessageParent = this.Handle;
            this.ddmrHistory.ItemClicked += new ToolStripItemClickedEventHandler(this.dropDowns_ItemClicked);
            this.ddmrHistory.ItemRightClicked += new ItemRightClickedEventHandler(this.dropDowns_ItemRightClicked);
            //
            // tmiHistory 
            //
            this.tmiHistory = new TitleMenuItem(MenuGenre.RecentlyClosedTab, false);
            this.tmiHistory.DropDown = this.ddmrHistory;
            //
            // ddmrUserapps 
            //
            this.ddmrUserapps = new DropDownMenuReorderable(this.components);
            this.ddmrUserapps.ReorderEnabled = reorderEnabled;
            this.ddmrUserapps.ImageList = QTUtility.ImageListGlobal;
            this.ddmrUserapps.MessageParent = this.Handle;
            this.ddmrUserapps.ReorderFinished += new MenuReorderedEventHandler(this.dropDowns_ReorderFinished);
            this.ddmrUserapps.ItemClicked += new ToolStripItemClickedEventHandler(this.dropDowns_ItemClicked);
            this.ddmrUserapps.ItemRightClicked += new ItemRightClickedEventHandler(this.dropDowns_ItemRightClicked);
            //
            // tmiUserApp
            //
            this.tmiUserApp = new TitleMenuItem(MenuGenre.Application, false);
            this.tmiUserApp.DropDown = this.ddmrUserapps;
            //
            // ddmrRecentFile
            //
            this.ddmrRecentFile = new DropDownMenuReorderable(this.components, false, false, false);
            this.ddmrRecentFile.ImageList = QTUtility.ImageListGlobal;
            this.ddmrRecentFile.MessageParent = this.Handle;
            this.ddmrRecentFile.ItemClicked += new ToolStripItemClickedEventHandler(this.dropDowns_ItemClicked);
            this.ddmrRecentFile.ItemRightClicked += new ItemRightClickedEventHandler(this.dropDowns_ItemRightClicked);
            //
            // tmiRecentFile
            //
            this.tmiRecentFile = new TitleMenuItem(MenuGenre.RecentFile, false);
            this.tmiRecentFile.DropDown = this.ddmrRecentFile;
            //
            // contextMenuForSetting
            //
            this.tsmiTaskBar = new ToolStripMenuItem();
            this.tsmiDesktop = new ToolStripMenuItem();
            this.tsmiLockItems = new ToolStripMenuItem();
            this.tsmiVSTitle = new ToolStripMenuItem();
            this.tsmiTaskBar.Checked = Config.Bool(Scts.Desktop_TaskBarDoubleClickEnabled);
            this.tsmiDesktop.Checked = Config.Bool(Scts.Desktop_DesktopDoubleClickEnabled);
            this.tsmiLockItems.Checked = Config.Bool(Scts.Desktop_LockMenu);
            this.tsmiVSTitle.Checked = Config.Bool(Scts.Desktop_TitleBackground);

            this.tsmiOnGroup = new ToolStripMenuItem();
            this.tsmiOnHistory = new ToolStripMenuItem();
            this.tsmiOnUserApps = new ToolStripMenuItem();
            this.tsmiOnRecentFile = new ToolStripMenuItem();
            this.tsmiOneClick = new ToolStripMenuItem();
            this.tsmiAppKeys = new ToolStripMenuItem();
            this.tsmiOnGroup.Checked = Config.Bool(Scts.Desktop_IncludeGroup);
            this.tsmiOnHistory.Checked = Config.Bool(Scts.Desktop_IncludeRecentTab);
            this.tsmiOnUserApps.Checked = Config.Bool(Scts.Desktop_IncludeApplication);
            this.tsmiOnRecentFile.Checked = Config.Bool(Scts.Desktop_IncludeRecentFile);
            this.tsmiOneClick.Checked = Config.Bool(Scts.Desktop_1ClickMenu);
            this.tsmiAppKeys.Checked = Config.Bool(Scts.Desktop_EnbleApplicationShortcuts);

            this.tsmiExperimental =
                    new ToolStripMenuItem(System.Globalization.CultureInfo.CurrentCulture.Parent.Name == "ja"
                            ? "ŽÀŒ±“I"
                            : "Experimental");
            this.tsmiExperimental.DropDown.Items.Add(new ToolStripMenuItem("dummy"));
            this.tsmiExperimental.DropDownDirection = ToolStripDropDownDirection.Left;
            this.tsmiExperimental.DropDownItemClicked +=
                    new ToolStripItemClickedEventHandler(tsmiExperimental_DropDownItemClicked);
            this.tsmiExperimental.DropDownOpening += new EventHandler(tsmiExperimental_DropDownOpening);

            this.contextMenuForSetting.Items.AddRange(new ToolStripItem[] {
                    this.tsmiTaskBar, this.tsmiDesktop, new ToolStripSeparator(),
                    this.tsmiOnGroup, this.tsmiOnHistory, this.tsmiOnUserApps, this.tsmiOnRecentFile,
                    new ToolStripSeparator(),
                    this.tsmiLockItems, this.tsmiVSTitle, this.tsmiOneClick, this.tsmiAppKeys, this.tsmiExperimental
            });
            this.contextMenuForSetting.ItemClicked +=
                    new ToolStripItemClickedEventHandler(this.contextMenuForSetting_ItemClicked);
            this.RefreshStringResources();

            //
            // QTCoTaskBar
            //
            this.ContextMenuStrip = this.contextMenuForSetting;
            this.Width = this.WidthOfBar;
            this.MinSize = new Size(8, 22);
            this.Dock = DockStyle.Fill;
            this.MouseClick += new MouseEventHandler(this.desktopTool_MouseClick);
            this.MouseDoubleClick += new MouseEventHandler(this.desktopTool_MouseDoubleClick);

            this.contextMenu.ResumeLayout(false);
            this.contextMenuForSetting.ResumeLayout(false);
            this.ResumeLayout(false);
        }


        #region ---------- Overriding Methods ----------

        public override int SetSite(object pUnkSite) {
            // order of method call
            // ctor -> SetSite -> InitializeComponent -> (touches Handle property, WM_CREATE) -> OnHandleCreated -> (OnVisibleChanged)

            if(base.BandObjectSite != null) {
                Marshal.ReleaseComObject(base.BandObjectSite);
            }
            base.BandObjectSite = (IInputObjectSite)pUnkSite;

            //////////////
            Application.EnableVisualStyles();

            this.ReadSetting();
            this.InitializeComponent();
            this.InstallDesktopHook();

            TitleMenuItem.DrawBackground = this.tsmiVSTitle.Checked;

            return S_OK;
        }

        public override int CloseDW(uint dwReserved) {
            // when user disable Desktop Tool
            // this seems not to be called on log off / shut down...

            if(this.iContextMenu2 != null) {
                Marshal.ReleaseComObject(this.iContextMenu2);
                this.iContextMenu2 = null;
            }

            // dispose controls in the thread they're created.
            DisposeInvoker disposeInvoker = new DisposeInvoker(this.InvokeDispose);
            if(this.thumbnailTooltip != null) {
                this.thumbnailTooltip.Invoke(disposeInvoker, new object[] {this.thumbnailTooltip});
                this.thumbnailTooltip = null;
            }
            if(this.subDirTip != null) {
                this.subDirTip.Invoke(disposeInvoker, new object[] {this.subDirTip});
                this.subDirTip = null;
            }

            // unhook, unsubclass
            if(this.hHook_MsgDesktop != IntPtr.Zero) {
                PInvoke.UnhookWindowsHookEx(this.hHook_MsgDesktop);
                this.hHook_MsgDesktop = IntPtr.Zero;
            }

            if(this.hHook_MsgShell_TrayWnd != IntPtr.Zero) {
                PInvoke.UnhookWindowsHookEx(this.hHook_MsgShell_TrayWnd);
                this.hHook_MsgShell_TrayWnd = IntPtr.Zero;
            }

            if(this.hHook_KeyDesktop != IntPtr.Zero) {
                PInvoke.UnhookWindowsHookEx(this.hHook_KeyDesktop);
                this.hHook_KeyDesktop = IntPtr.Zero;
            }

            if(this.shellViewListener != null) {
                this.shellViewListener.ReleaseHandle();
                this.shellViewListener = null;
            }

            return base.CloseDW(dwReserved);
        }

        public override int GetClassID(out Guid pClassID) {
            pClassID = typeof(QTDesktopTool).GUID;
            return S_OK;
        }

        protected override void WndProc(ref Message m) {
            const int WM_COPYDATA = 0x004A;
            const int WM_INITMENUPOPUP = 0x0117;
            const int WM_DRAWITEM = 0x002B;
            const int WM_MEASUREITEM = 0x002C;
            const int WM_MOUSEACTIVATE = 0x0021;
            const int MA_NOACTIVATEANDEAT = 4;
            const int WM_LBUTTONDOWN = 0x0201;
            const int WM_DWMCOMPOSITIONCHANGED = 0x031E;

            switch(m.Msg) {
                case WM_INITMENUPOPUP:
                case WM_DRAWITEM:
                case WM_MEASUREITEM:

                    // these messages are forwarded to draw sub items in 'Send to" of shell context menu.

                    if(this.iContextMenu2 != null) {
                        this.iContextMenu2.HandleMenuMsg(m.Msg, m.WParam, m.LParam);
                        return;
                    }
                    break;


                case WM_MOUSEACTIVATE:

                    if(Config.Bool(Scts.Desktop_1ClickMenu)) {
                        if(((((int)(long)m.LParam) >> 16) & 0xFFFF) == WM_LBUTTONDOWN) {
                            if(this.contextMenu.Visible) {
                                this.contextMenu.Close(ToolStripDropDownCloseReason.AppClicked);

                                m.Result = (IntPtr)MA_NOACTIVATEANDEAT;
                                return;
                            }
                        }
                    }
                    break;


                case WM_COPYDATA:

                    COPYDATASTRUCT cds = (COPYDATASTRUCT)Marshal.PtrToStructure(m.LParam, typeof(COPYDATASTRUCT));

                    switch((int)m.WParam) {
                        case MC.DTT_SHOW_MENU:

                            // command to open menu on this taskbar thread 
                            // cds.dwData : HIWORD is y, LOWORD is x.

                            int x = QTUtility2.GET_X_LPARAM(cds.dwData);
                            int y = QTUtility2.GET_Y_LPARAM(cds.dwData);

                            PInvoke.SetForegroundWindow(this.hwndShellTray);
                            this.ShowMenu(new Point(x, y));
                            break;


                            //case MC.DTT_SHOW_HASHWINDOW:

                            //    // forward the message to desktop thread.
                            //    QTUtility2.SendCOPYDATASTRUCT( this.hwndListView, (IntPtr)MC.DTT_SHOW_HASHWINDOW, Marshal.PtrToStringUni( cds.lpData ), cds.dwData );
                            //    return;

                    }
                    break;


                case WM_DWMCOMPOSITIONCHANGED:

                    this.Invalidate();
                    break;


                case MC.QTDT_REFRESH_TEXTRES:

                    this.RefreshStringResources();
                    break;
            }

            base.WndProc(ref m);
        }

        protected override void Dispose(bool disposing) {
            if(disposing && (components != null)) {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        protected override void OnPaintBackground(PaintEventArgs e) {
            // background
            if(VisualStyleRenderer.IsSupported) {
                if(this.bgRenderer == null) {
                    this.bgRenderer = new VisualStyleRenderer(VisualStyleElement.Taskbar.BackgroundTop.Normal);
                }

                this.bgRenderer.DrawParentBackground(e.Graphics, e.ClipRectangle, this);
            }
            else {
                base.OnPaintBackground(e);
            }


            // strings

            if(this.fNowMouseHovering) {
                Color clr = VisualStyleRenderer.IsSupported ? SystemColors.Window : SystemColors.WindowText;

                if(this.stringFormat == null) {
                    this.stringFormat = new StringFormat();
                    this.stringFormat.Alignment = StringAlignment.Center;
                    this.stringFormat.LineAlignment = StringAlignment.Center;
                    this.stringFormat.FormatFlags = StringFormatFlags.NoWrap;
                }

                using(SolidBrush sb = new SolidBrush(Color.FromArgb(128, clr))) {
                    e.Graphics.DrawString(TEXT_TOOLBAR, this.Font, sb,
                            new Rectangle(0, 5, e.ClipRectangle.Width - 1, e.ClipRectangle.Height - 6),
                            this.stringFormat);
                }

                using(Pen p = new Pen(Color.FromArgb(128, clr))) {
                    e.Graphics.DrawRectangle(p,
                            new Rectangle(0, 2, e.ClipRectangle.Width - 1, e.ClipRectangle.Height - 3));
                }
            }
        }

        protected override void OnMouseEnter(EventArgs e) {
            this.fNowMouseHovering = true;
            base.OnMouseEnter(e);
            this.Invalidate();
        }

        protected override void OnMouseLeave(EventArgs e) {
            this.fNowMouseHovering = false;
            base.OnMouseLeave(e);
            this.Invalidate();
        }

        protected override void OnHandleCreated(EventArgs e) {
            base.OnHandleCreated(e);

            if(this.IsHandleCreated) {
                using(RegistryKey rkUser = Registry.CurrentUser.CreateSubKey(REGNAME.KEY_USERROOT)) {
                    QTUtility2.WriteRegHandle(REGNAME.VAL_DT_HANDLE, rkUser, this.Handle);
                }
            }
        }

        protected override void OnHandleDestroyed(EventArgs e) {
            base.OnHandleDestroyed(e);

            using(RegistryKey rkUser = Registry.CurrentUser.CreateSubKey(REGNAME.KEY_USERROOT)) {
                QTUtility2.WriteRegHandle(REGNAME.VAL_DT_HANDLE, rkUser, IntPtr.Zero);
            }
        }

        private void InvokeDispose(IDisposable disposable) {
            disposable.Dispose();
        }

        #endregion


        #region ---------- Hooks and subclassings ----------


        private void InstallDesktopHook() {
            const int WH_KEYBOARD = 2;
            const int WH_GETMESSAGE = 3;

            IntPtr hwndDesktop = QTDesktopTool.GetDesktopHwnd();
            if(this.timerHooks == null) {
                if(hwndDesktop == IntPtr.Zero) {
                    // wait til desktop window is created 
                    this.timerHooks = new System.Windows.Forms.Timer();
                    this.timerHooks.Tick += new EventHandler(timerHooks_Tick);
                    this.timerHooks.Interval = 3000;
                    this.timerHooks.Start();
                    return;
                }
            }
            else {
                if(hwndDesktop == IntPtr.Zero) {
                    return;
                }
                else {
                    this.timerHooks.Stop();
                    this.timerHooks.Dispose();
                    this.timerHooks = null;
                }
            }

            // Now we've got desktop window handle

            this.hwndListView = hwndDesktop;
            this.hwndShellTray = WindowUtils.GetShellTrayWnd();

            this.hookProc_Msg_Desktop = new HookProc(CallbackGetMsgProc_Desktop);
            this.hookProc_Msg_ShellTrayWnd = new HookProc(CallbackGetMsgProc_ShellTrayWnd);
            this.hookProc_Keys_Desktop = new HookProc(CallbackKeyProc_Desktop);

            int id1, id2;
            int threadID_Desktop = PInvoke.GetWindowThreadProcessId(this.hwndListView, out id1);
            int threadID_ShellTray = PInvoke.GetWindowThreadProcessId(this.hwndShellTray, out id2);

            this.hHook_MsgDesktop = PInvoke.SetWindowsHookEx(WH_GETMESSAGE, this.hookProc_Msg_Desktop, IntPtr.Zero,
                    threadID_Desktop);
            this.hHook_MsgShell_TrayWnd = PInvoke.SetWindowsHookEx(WH_GETMESSAGE, this.hookProc_Msg_ShellTrayWnd,
                    IntPtr.Zero, threadID_ShellTray);
            this.hHook_KeyDesktop = PInvoke.SetWindowsHookEx(WH_KEYBOARD, this.hookProc_Keys_Desktop, IntPtr.Zero,
                    threadID_Desktop);

            // get IFolderView on the desktop thread...
            PInvoke.PostMessage(this.hwndListView, MC.QTDT_GETFOLDERVIEW, IntPtr.Zero, IntPtr.Zero);

            // subclassing ShellView
            this.hwndShellView = PInvoke.GetWindowLongPtr(this.hwndListView, GWL.HWNDPARENT);
            if(this.hwndShellView != IntPtr.Zero) {
                this.shellViewListener = new NativeWindowListener(this.hwndShellView);
                this.shellViewListener.MessageCaptured +=
                        new NativeWindowListener.MessageEventHandler(this.shellViewListener_MessageCaptured);
            }
        }

        private void timerHooks_Tick(object sender, EventArgs e) {
            if(++this.iHookTimeout > 5) {
                this.timerHooks.Stop();
                MessageBox.Show("Failed to hook Desktop. Please re-enable QT Tab Desktop tool.");
                return;
            }

            this.InstallDesktopHook();
        }

        private void GetFolderView() {
            // desktop thread
            // WM_GETISHELLBROWSER is not supported by Windows7 any more

            const int SWC_DESKTOP = 0x00000008;
            const int SWFO_NEEDDISPATCH = 0x00000001;

            SHDocVw.ShellWindows shellWindows = null;
            try {
                shellWindows = new SHDocVw.ShellWindows();
                object oNull1 = null, oNull2 = null;
                int pHWND;
                object o = shellWindows.FindWindowSW(ref oNull1, ref oNull2, SWC_DESKTOP, out pHWND, SWFO_NEEDDISPATCH);

                _IServiceProvider sp = o as _IServiceProvider;
                if(sp != null) {
                    object oShellBrowser;
                    sp.QueryService(COMGUIDS.SID_SShellBrowser, COMGUIDS.IID_IShellBrowser, out oShellBrowser);

                    this.shellBrowser = oShellBrowser as IShellBrowser;
                    if(this.shellBrowser != null) {
                        IShellView shellView;
                        if(S_OK == this.shellBrowser.QueryActiveShellView(out shellView)) {
                            this.folderView = shellView as IFolderView;
                        }
                    }
                }
            }
            catch {
            }
            finally {
                if(shellWindows != null) {
                    Marshal.ReleaseComObject(shellWindows);
                }
            }

            // old codes

            //const int WM_USER = 0x0400;
            //const int WM_GETISHELLBROWSER = ( WM_USER + 7 );

            //IntPtr hwndProgman = QTDesktopTool.GetProgmanHWnd();

            //IntPtr pShellBrowser = PInvoke.SendMessage( hwndProgman, WM_GETISHELLBROWSER, IntPtr.Zero, IntPtr.Zero );
            //if( pShellBrowser != IntPtr.Zero )
            //{
            //    IShellBrowser shellBrowser = null;
            //    try
            //    {
            //        shellBrowser = Marshal.GetObjectForIUnknown( pShellBrowser ) as IShellBrowser;
            //        if( shellBrowser != null )
            //        {
            //            IShellView shellView;
            //            if( S_OK == shellBrowser.QueryActiveShellView( out shellView ) )
            //            {
            //                this.folderView = shellView as IFolderView;

            //                return this.folderView != null;
            //            }
            //        }
            //    }
            //    catch
            //    {
            //    }
            //    finally
            //    {
            //        if( shellBrowser != null )
            //        {
            //            Marshal.ReleaseComObject( shellBrowser );
            //        }

            //        Marshal.Release( pShellBrowser );
            //    }
            //}
            //return false;
        }


        private bool shellViewListener_MessageCaptured(ref Message msg) {
            const int LVN_FIRST = -100;
            const int LVN_ITEMCHANGED = (LVN_FIRST - 1);
            const int LVN_KEYDOWN = (LVN_FIRST - 55);
            const int LVN_GETINFOTIP = (LVN_FIRST - 58);
            const int LVN_HOTTRACK = (LVN_FIRST - 21);
            const int LVN_DELETEITEM = (LVN_FIRST - 3);
            const int LVN_ITEMACTIVATE = (LVN_FIRST - 14);
            const int LVN_BEGINDRAG = (LVN_FIRST - 9);
            const int LVN_BEGINRDRAG = (LVN_FIRST - 11);


            const int NM_FIRST = 0;
            const int NM_KILLFOCUS = (NM_FIRST - 8);

            const uint LVIF_STATE = 0x00000008;
            const uint LVIS_DROPHILITED = 0x0008;

            const int WM_NOTIFY = 0x004E;
            const int WM_INITMENUPOPUP = 0x0117;
            const int WM_DRAWITEM = 0x002B;
            const int WM_MEASUREITEM = 0x002C;


            switch(msg.Msg) {
                case WM_INITMENUPOPUP:
                case WM_DRAWITEM:
                case WM_MEASUREITEM:

                    // these messages are forwarded to draw sub items in 'Send to" of shell context menu on SubDirTip menu.

                    if(this.iContextMenu2_Desktop != null) {
                        this.iContextMenu2_Desktop.HandleMenuMsg(msg.Msg, msg.WParam, msg.LParam);
                        return true;
                    }
                    break;
            }

            if(msg.Msg == WM_NOTIFY) {
                NMHDR nmhdr = (NMHDR)Marshal.PtrToStructure(msg.LParam, typeof(NMHDR));

                // hwndFrom is not SysListView32 handle, do nothing.
                if(nmhdr.hwndFrom != this.hwndListView)
                    return false;

                switch(nmhdr.code) {
                    case LVN_ITEMCHANGED:

                        if(Config.Get(Scts.SubDirTip) > 0 && Config.Bool(Scts.SubDirMenuForDropHilited)) {
                            NMLISTVIEW nmlv = (NMLISTVIEW)Marshal.PtrToStructure(msg.LParam, typeof(NMLISTVIEW));

                            if(nmlv.uChanged == LVIF_STATE) {
                                if(nmlv.iItem != this.itemIndexDROPHILITED) {
                                    uint uNewState_DropHilited = nmlv.uNewState & LVIS_DROPHILITED;
                                    uint uOldState_DropHilited = nmlv.uOldState & LVIS_DROPHILITED;

                                    if(uOldState_DropHilited != uNewState_DropHilited) {
                                        if(uNewState_DropHilited == 0) {
                                            this.HandleDROPHILITED(-1);
                                        }
                                        else {
                                            this.HandleDROPHILITED(nmlv.iItem);
                                        }
                                    }
                                }
                            }
                        }
                        break;


                    case LVN_GETINFOTIP:

                        // show thumbnail window instead of tooltip for pointed file.
                        // return false to show ordinary file tooltip.
                        if(Config.Get(Scts.Preview) > 0 &&
                                (Config.Get(Scts.Preview) == 1 ^ Control.ModifierKeys == Keys.Shift)) {
                            NMLVGETINFOTIP nmlvgit =
                                    (NMLVGETINFOTIP)Marshal.PtrToStructure(msg.LParam, typeof(NMLVGETINFOTIP));

                            IntPtr pidl = this.GetItemPIDL(nmlvgit.iItem);
                            try {
                                if(pidl != IntPtr.Zero) {
                                    this.thumbnailIndex_GETINFOTIP = nmlvgit.iItem;

                                    if(this.thumbnailTooltip != null && this.thumbnailTooltip.Visible &&
                                            nmlvgit.iItem == this.thumbnailIndex)
                                        return true;

                                    if(this.timer_Thumbnail_NoTooltip != null && this.timer_Thumbnail_NoTooltip.Enabled)
                                        return true;

                                    RECT rct = QTDesktopTool.GetLVITEMRECT(nmhdr.hwndFrom, nmlvgit.iItem, false, 0);
                                    return this.ShowThumbnailTooltip(pidl, nmlvgit.iItem,
                                            !PInvoke.PtInRect(ref rct, Control.MousePosition));
                                }
                            }
                            finally {
                                if(pidl != IntPtr.Zero) {
                                    PInvoke.CoTaskMemFree(pidl);
                                }
                            }
                        }
                        break;


                    case LVN_HOTTRACK:

                        if(Config.Get(Scts.Preview) > 0 || Config.Get(Scts.SubDirTip) > 0) {
                            NMLISTVIEW nmlv = (NMLISTVIEW)Marshal.PtrToStructure(msg.LParam, typeof(NMLISTVIEW));
                            Keys modKeys = Control.ModifierKeys;

                            // show thumbnail window immediately
                            if(Config.Get(Scts.Preview) > 0) {
                                if(this.thumbnailTooltip != null &&
                                        (this.thumbnailTooltip.Visible || this.fThumbnailPending ||
                                                this.thumbnailIndex_GETINFOTIP == nmlv.iItem)) {
                                    if(Config.Get(Scts.Preview) == 1 ^ modKeys == Keys.Shift) {
                                        if(nmlv.iItem != this.thumbnailIndex) {
                                            if(nmlv.iItem > -1) {
                                                IntPtr pidl = this.GetItemPIDL(nmlv.iItem);
                                                try {
                                                    if(pidl != IntPtr.Zero) {
                                                        if(this.ShowThumbnailTooltip(pidl, nmlv.iItem, false))
                                                            return false;
                                                    }
                                                }
                                                finally {
                                                    if(pidl != IntPtr.Zero) {
                                                        PInvoke.CoTaskMemFree(pidl);
                                                    }
                                                }
                                            }

                                            // failed to show thumbnail, hide the window
                                            this.HideThumbnailTooltip();
                                        }
                                    }
                                    else {
                                        this.HideThumbnailTooltip();
                                    }
                                }
                                else {
                                    if(Config.Bool(Scts.PreviewForInactiveWindow) &&
                                            (this.hwndListView != PInvoke.GetFocus())) {
                                        this.thumbnailIndex_Inactive = nmlv.iItem;

                                        if(this.timer_ThumbnailMouseHover == null) {
                                            this.timer_ThumbnailMouseHover =
                                                    new System.Windows.Forms.Timer(this.components);
                                            this.timer_ThumbnailMouseHover.Interval = SystemInformation.MouseHoverTime;
                                            this.timer_ThumbnailMouseHover.Tick +=
                                                    new EventHandler(timer_ThumbnailMouseHover_Tick);
                                        }
                                        this.timer_ThumbnailMouseHover.Enabled = false;
                                        this.timer_ThumbnailMouseHover.Enabled = true;
                                    }
                                }

                                // this timer prevents ordinary tooltip showing while mouse cursor is moving on items that have no thumbnail.
                                if(this.timer_Thumbnail_NoTooltip == null) {
                                    this.timer_Thumbnail_NoTooltip = new System.Windows.Forms.Timer(this.components);
                                    this.timer_Thumbnail_NoTooltip.Interval =
                                            (int)(SystemInformation.MouseHoverTime*0.2);
                                    this.timer_Thumbnail_NoTooltip.Tick +=
                                            new EventHandler(timer_Thumbnail_NoTooltip_Tick);
                                }
                                this.timer_Thumbnail_NoTooltip.Enabled = false;
                                this.timer_Thumbnail_NoTooltip.Enabled = true;
                            }

                            // show subdirtip
                            if(Config.Get(Scts.SubDirTip) > 0) {
                                if(Config.Get(Scts.SubDirTip) == 1 ^ modKeys == Keys.Shift) {
                                    if(nmlv.iItem > -1) {
                                        IntPtr pidl = this.GetItemPIDL(nmlv.iItem);
                                        try {
                                            if(pidl != IntPtr.Zero) {
                                                if(this.ShowSubDirTip(pidl, nmlv.iItem, false))
                                                    return false;
                                            }
                                        }
                                        finally {
                                            if(pidl != IntPtr.Zero) {
                                                PInvoke.CoTaskMemFree(pidl);
                                            }
                                        }
                                    }
                                }
                                this.HideSubDirTip();
                            }
                        }
                        break;


                    case LVN_KEYDOWN:

                        const int VK_SHIFT = 0x10;

                        NMLVKEYDOWN nmlvkd = (NMLVKEYDOWN)Marshal.PtrToStructure(msg.LParam, typeof(NMLVKEYDOWN));
                        int lvnVKey = nmlvkd.wVKey;

                        if(Config.Get(Scts.Preview) > 0) {
                            // hide thumbnail tip
                            if(Config.Get(Scts.Preview) == 1 || lvnVKey != VK_SHIFT) {
                                this.HideThumbnailTooltip();
                            }
                        }
                        break;


                    case NM_KILLFOCUS:

                        this.HideThumbnailTooltip();
                        this.HideSubDirTip_DesktopInactivated();
                        this.thumbnailIndex_GETINFOTIP = -1;
                        break;


                    case LVN_DELETEITEM:

                        if(Config.Get(Scts.SubDirTip) > 0) {
                            this.HideSubDirTip();
                        }
                        this.thumbnailIndex_GETINFOTIP = -1;
                        break;


                    case LVN_ITEMACTIVATE:

                        const int LVKF_ALT = 0x0001;
                        const int LVKF_CONTROL = 0x0002;
                        const int LVKF_SHIFT = 0x0004;

                        NMITEMACTIVATE nmia = (NMITEMACTIVATE)Marshal.PtrToStructure(msg.LParam, typeof(NMITEMACTIVATE));

                        bool fEnqExec = Config.Get(Scts.SaveRecentFile) > 0;

                        Keys modKey = ((nmia.uKeyFlags & LVKF_ALT) == LVKF_ALT ? Keys.Alt : Keys.None) |
                                ((nmia.uKeyFlags & LVKF_CONTROL) == LVKF_CONTROL ? Keys.Control : Keys.None) |
                                        ((nmia.uKeyFlags & LVKF_SHIFT) == LVKF_SHIFT ? Keys.Shift : Keys.None);

                        // Note:
                        //			nmia.iItem is useless, because it does not point to an activated item but to a focused item.

                        return this.HandleTabFolderActions(-1, modKey, fEnqExec);


                    case LVN_BEGINDRAG:
                    case LVN_BEGINRDRAG:

                        // hide thumbnails when dragging starts
                        this.HideThumbnailTooltip();
                        this.thumbnailIndex_GETINFOTIP = -1;
                        break;
                }
            }
            return false;
        }

        private IntPtr CallbackGetMsgProc_Desktop(int nCode, IntPtr wParam, IntPtr lParam) {
            const int WM_LBUTTONDBLCLK = 0x0203;
            const int WM_MBUTTONUP = 0x0208;
            const int WM_MOUSEWHEEL = 0x020A;

            if(nCode >= 0) {
                MSG msg = (MSG)Marshal.PtrToStructure(lParam, typeof(MSG));
                switch(msg.message) {
                    case WM_MOUSEWHEEL:

                        // redirect mouse wheel to menu
                        IntPtr hwnd = PInvoke.WindowFromPoint(QTUtility2.GET_POINT_LPARAM(msg.lParam));
                        if(hwnd != IntPtr.Zero && hwnd != msg.hwnd) {
                            Control ctrl = Control.FromHandle(hwnd);
                            if(ctrl != null) {
                                DropDownMenuReorderable ddmr = ctrl as DropDownMenuReorderable;
                                if(ddmr != null) {
                                    if(ddmr.CanScroll) {
                                        PInvoke.SendMessage(hwnd, WM_MOUSEWHEEL, msg.wParam, msg.lParam);
                                    }
                                    Marshal.StructureToPtr(new MSG(), lParam, false);
                                }
                            }
                        }
                        break;


                    case WM_LBUTTONDBLCLK:

                        if(msg.hwnd == this.hwndListView && Config.Bool(Scts.Desktop_DesktopDoubleClickEnabled)) {
                            int index = PInvoke.ListView_HitTest(this.hwndListView, msg.lParam);
                            if(index == -1) {
                                // do the menu on Taskbar thread.
                                QTUtility2.SendCOPYDATASTRUCT(this.hwndThis, (IntPtr)MC.DTT_SHOW_MENU, "fromdesktop",
                                        msg.lParam);
                            }
                        }
                        break;


                    case WM_MBUTTONUP:

                        if(msg.hwnd == this.hwndListView && Config.Positive(Scts.ViewIconMiddleClicked)) {
                            int iItem = PInvoke.ListView_HitTest(this.hwndListView, msg.lParam);

                            if(iItem != -1) {
                                if(this.HandleTabFolderActions(iItem, Control.ModifierKeys, false)) {
                                    Marshal.StructureToPtr(new MSG(), lParam, false);
                                }
                            }
                        }
                        break;


                    case MC.QTDT_GETFOLDERVIEW:

                        if(msg.hwnd == this.hwndListView) {
                            this.GetFolderView();
                        }
                        break;
                }
            }

            return PInvoke.CallNextHookEx(this.hHook_MsgDesktop, nCode, wParam, lParam);
        }

        private IntPtr CallbackGetMsgProc_ShellTrayWnd(int nCode, IntPtr wParam, IntPtr lParam) {
            const int WM_NCLBUTTONDBLCLK = 0x00A3;
            const int WM_MOUSEWHEEL = 0x020A;

            if(nCode >= 0) {
                MSG msg = (MSG)Marshal.PtrToStructure(lParam, typeof(MSG));

                switch(msg.message) {
                    case WM_NCLBUTTONDBLCLK:

                        if(Config.Bool(Scts.Desktop_TaskBarDoubleClickEnabled) && msg.hwnd == this.hwndShellTray) {
                            this.ShowMenu(Control.MousePosition);
                            Marshal.StructureToPtr(new MSG(), lParam, false);
                        }
                        break;


                    case WM_MOUSEWHEEL:

                        IntPtr hwnd =
                                PInvoke.WindowFromPoint(new Point(QTUtility2.GET_X_LPARAM(msg.lParam),
                                        QTUtility2.GET_Y_LPARAM(msg.lParam)));
                        if(hwnd != IntPtr.Zero && hwnd != msg.hwnd) {
                            Control ctrl = Control.FromHandle(hwnd);
                            if(ctrl != null) {
                                DropDownMenuReorderable ddmr = ctrl as DropDownMenuReorderable;
                                if(ddmr != null && ddmr.CanScroll) {
                                    PInvoke.SendMessage(hwnd, WM_MOUSEWHEEL, msg.wParam, msg.lParam);
                                    Marshal.StructureToPtr(new MSG(), lParam, false);
                                }
                            }
                        }
                        break;
                }
            }

            return PInvoke.CallNextHookEx(this.hHook_MsgShell_TrayWnd, nCode, wParam, lParam);
        }

        private IntPtr CallbackKeyProc_Desktop(int nCode, IntPtr wParam, IntPtr lParam) {
            if(nCode >= 0) {
                if(((ulong)lParam & 0x80000000) == 0) {
                    // transition state == 0, key is pressed

                    if(this.HandleKEYDOWN_Desktop(wParam, (((ulong)lParam & 0x40000000) == 0x40000000)))
                        return new IntPtr(1);
                }
                else {
                    // transition state == 1, key is released

                    //this.HideThumbnailTooltip();

                    if(Config.Get(Scts.SubDirTip) == 2) // subfirtip only with shift key 
                    {
                        if(this.subDirTip != null && !this.subDirTip.MenuIsVisible) {
                            this.HideSubDirTip();
                        }
                    }
                }
            }
            return PInvoke.CallNextHookEx(this.hHook_KeyDesktop, nCode, wParam, lParam);
        }


        private bool HandleKEYDOWN_Desktop(IntPtr wParam, bool fRepeat) {
            Keys rawKey = (Keys)(int)wParam;
            int key = (int)wParam | (int)Control.ModifierKeys;

            const int VK_F2 = 0x71;

            if(rawKey == Keys.ShiftKey) {
                if(!fRepeat) {
                    if(Config.Get(Scts.Preview) == 1) // preview is enabled without shift key
                    {
                        this.HideThumbnailTooltip();
                    }

                    if(Config.Get(Scts.SubDirTip) == 1 && // subdirtip is enabled without shift key
                            this.subDirTip != null &&
                                    this.subDirTip.Visible &&
                                            !this.subDirTip.MenuIsVisible) {
                        this.HideSubDirTip();
                    }
                }
                return false;
            }
            else if(rawKey == Keys.Delete) {
                if(!fRepeat) {
                    if(Config.Get(Scts.Preview) > 0) {
                        this.HideThumbnailTooltip();
                    }
                    if(Config.Get(Scts.SubDirTip) > 0) {
                        if(this.subDirTip != null && this.subDirTip.Visible && !this.subDirTip.MenuIsVisible) {
                            this.HideSubDirTip();
                        }
                    }
                }
                return false;
            }
            else if(key == VK_F2) //F2
            {
                if(Config.Bool(Scts.ViewF2ChangeTextSelection)) {
                    QTTabBarClass.HandleF2(this.hwndListView, true);
                }
                return false;
            }



            key |= QTUtility.FLAG_KEYENABLED;
            if(key == QTUtility.ShortcutKeys[KeyShortcuts.CopyPath] ||
                    key == QTUtility.ShortcutKeys[KeyShortcuts.CopyName] ||
                            key == QTUtility.ShortcutKeys[KeyShortcuts.CopyCurrentPath] ||
                                    key == QTUtility.ShortcutKeys[KeyShortcuts.CopyCurrentName] ||
                                            key == QTUtility.ShortcutKeys[KeyShortcuts.ShowHashWindow] ||
                                                    key == QTUtility.ShortcutKeys[KeyShortcuts.CopyFileHash]) {
                if(!fRepeat) {
                    if(this.subDirTip != null && this.subDirTip.MenuIsVisible)
                        return false;

                    int index = 0;
                    if(key == QTUtility.ShortcutKeys[KeyShortcuts.CopyName])
                        index = 1;
                    else if(key == QTUtility.ShortcutKeys[KeyShortcuts.CopyCurrentPath])
                        index = 2;
                    else if(key == QTUtility.ShortcutKeys[KeyShortcuts.CopyCurrentName])
                        index = 3;
                    else if(key == QTUtility.ShortcutKeys[KeyShortcuts.ShowHashWindow])
                        index = 4;
                    else if(key == QTUtility.ShortcutKeys[KeyShortcuts.CopyFileHash])
                        index = 6;

                    this.DoFileTools(index);
                }
                return true;
            }
            else if(key == QTUtility.ShortcutKeys[KeyShortcuts.ShowSubFolderMenu]) {
                // Show SubDirTip for selected folder.

                if(Config.Get(Scts.SubDirTip) != 0) {
                    if(!fRepeat) {
                        this.DoFileTools(5);
                    }
                    return true;
                }
            }
            else if(key == QTUtility.ShortcutKeys[KeyShortcuts.ShowPreview]) {
                if(Config.Get(Scts.Preview) > 0) {
                    if(!fRepeat) {
                        this.ShowThumbnailTooltipForSelectedItem();
                    }
                    return true;
                }
            }
            else if(key == QTUtility.ShortcutKeys[KeyShortcuts.DeleteFile] ||
                    key == QTUtility.ShortcutKeys[KeyShortcuts.DeleteFileNuke]) {
                if(!fRepeat) {
                    this.DeleteSelection(key == QTUtility.ShortcutKeys[KeyShortcuts.DeleteFileNuke]);
                }
                return true;
            }
            else if(key == QTUtility.ShortcutKeys[KeyShortcuts.InvertSelection]) {
                if(!fRepeat) {
                    WindowUtils.ExecuteMenuCommand(this.hwndShellView, ExplorerMenuCommand.InvertSelection);
                }
                return true;
            }
            else if(key == QTUtility.ShortcutKeys[KeyShortcuts.PasteShortcut]) {
                if(!fRepeat) {
                    using(IDLWrapper idlw = new IDLWrapper(new byte[] {0, 0}, false)) {
                        if(idlw.Available && !idlw.IsReadOnly &&
                                ShellMethods.ClipboardContainsFileDropList(this.hwndListView, false, true)) {
                            WindowUtils.ExecuteMenuCommand(this.hwndShellView, ExplorerMenuCommand.PasteShortcut);
                        }
                        else {
                            System.Media.SystemSounds.Beep.Play();
                        }
                    }
                }
                return true;
            }
            else if(key == QTUtility.ShortcutKeys[KeyShortcuts.CreateNewFolder] ||
                    key == QTUtility.ShortcutKeys[KeyShortcuts.CreateNewTxtFile]) {
                if(!fRepeat) {
                    using(IDLWrapper idlw = new IDLWrapper(new byte[] {0, 0}, false)) {
                        ShellMethods.CreateNewItem(this.shellBrowser, idlw,
                                key == QTUtility.ShortcutKeys[KeyShortcuts.CreateNewFolder]);
                    }
                }
                return true;
            }

            else if(key == QTUtility.ShortcutKeys[KeyShortcuts.CreateShortcut] ||
                    key == QTUtility.ShortcutKeys[KeyShortcuts.CopyToFolder] ||
                            key == QTUtility.ShortcutKeys[KeyShortcuts.MoveToFolder]) {
                if(!fRepeat) {
                    if(this.GetSelectionCount() > 0) {
                        ExplorerMenuCommand command = ExplorerMenuCommand.CreateShortcut;
                        if(key == QTUtility.ShortcutKeys[KeyShortcuts.CopyToFolder]) {
                            command = ExplorerMenuCommand.CopyToFolder;
                        }
                        else if(key == QTUtility.ShortcutKeys[KeyShortcuts.MoveToFolder]) {
                            command = ExplorerMenuCommand.MoveToFolder;
                        }

                        WindowUtils.ExecuteMenuCommand(this.hwndShellView, command);
                    }
                    else {
                        System.Media.SystemSounds.Beep.Play();
                    }
                }
                return true;
            }



            else if(Config.Bool(Scts.Desktop_EnbleApplicationShortcuts) &&
                    QTUtility.dicUserAppShortcutKeys.ContainsKey(key)) {
                // Application shortcut keys.

                if(!fRepeat) {
                    MenuItemArguments mia = QTUtility.dicUserAppShortcutKeys[key];
                    try {
                        int index;
                        List<QTPlugin.Address> lstAddress = new List<QTPlugin.Address>();
                        foreach(IntPtr pidl in this.GetSelectedItemPIDL(out index)) {
                            if(pidl != IntPtr.Zero) {
                                using(IDLWrapper idlw = new IDLWrapper(pidl)) {
                                    if(idlw.Available && idlw.HasPath && idlw.IsFileSystem) {
                                        lstAddress.Add(new QTPlugin.Address(pidl, idlw.Path));
                                    }
                                }
                            }
                        }

                        AppLauncher al = new AppLauncher(lstAddress.ToArray(),
                                Environment.GetFolderPath(Environment.SpecialFolder.Desktop));
                        al.ReplaceTokens_WorkingDir(mia);
                        al.ReplaceTokens_Arguments(mia);

                        AppLauncher.Execute(mia, this.hwndListView); // runs on desktop thread
                        return true;
                    }
                    catch(Exception ex) {
                        DebugUtil.AppendToExceptionLog(ex, null);
                    }
                    finally {
                        mia.RestoreOriginalArgs();
                    }
                }
            }
            else if(!fRepeat && QTUtility.dicGroupShortcutKeys.ContainsKey(key)) {
                // Group shortcut keys

                Thread thread = new Thread(this.OpenGroup);
                thread.SetApartmentState(ApartmentState.STA);
                thread.IsBackground = true;
                thread.Start(new object[]
                {new string[] {QTUtility.dicGroupShortcutKeys[key]}, Keys.None});
                return true;
            }

            return false;
        }

        private bool HandleTabFolderActions(int index, Keys modKey, bool fEnqExec) {
            // Handles item activation in Desktop thread

            // Default		..... New Tab / Navigate to
            // C			..... New Window
            // S			..... New Tab without selecting
            // C + S + A    ..... Open all sub folders in new tabs

            bool fMBUTTONUP = index != -1;

            if(fMBUTTONUP && Config.Get(Scts.ViewIconMiddleClicked) == 1) // new window by wheel click
            {
                if(modKey == Keys.Control) {
                    modKey = Keys.None;
                }
                else if(modKey == Keys.None) {
                    modKey = Keys.Control;
                }
            }
            if(!Config.Bool(Scts.ActivateNewTab)) // do not activate new tab
            {
                if((modKey & Keys.Shift) == Keys.Shift) {
                    modKey &= ~Keys.Shift;
                }
                else {
                    modKey |= Keys.Shift;
                }
            }

            List<IntPtr> lstPIDLs = new List<IntPtr>();
            List<byte[]> lstIDLs = new List<byte[]>();
            List<string> lstFiles = new List<string>();


            if(index != -1) {
                // WM_MBUTTONUP

                IntPtr p;
                if((p = this.GetItemPIDL(index)) != IntPtr.Zero) {
                    lstPIDLs.Add(p);
                }
                else {
                    return false;
                }
            }
            else {
                // Get selections
                int iItem;
                lstPIDLs = this.GetSelectedItemPIDL(out iItem);
            }

            foreach(IntPtr pidl in lstPIDLs) {
                using(IDLWrapper idlw = new IDLWrapper(pidl)) {
                    if(idlw.Available && idlw.IsReadyIfDrive) {
                        if(idlw.IsLink) {
                            // dead link check 
                            IDLWrapper idlwTarget;
                            if(!idlw.TryGetLinkTarget(this.hwndListView, out idlwTarget)) {
                                continue;
                            }
                            using(idlwTarget) {
                                if(idlwTarget.IsFolder && (fMBUTTONUP || !(idlw.IsStream && idlw.IsBrowsable)))
                                        // when wheel clicked, allow navigation to zip folder 
                                {
                                    if(idlwTarget.IsReadyIfDrive) {
                                        lstIDLs.Add(idlw.IsFolder ? idlw.IDL : idlwTarget.IDL);
                                    }
                                }
                                else {
                                    if(fEnqExec) {
                                        if(idlw.HasPath) {
                                            lstFiles.Add(idlw.Path);
                                        }
                                    }
                                }
                            }
                        }
                        else if(idlw.IsFolder && (fMBUTTONUP || !(idlw.IsStream && idlw.IsBrowsable)))
                                // when wheel clicked, allow navigation to zip folder 
                        {
                            lstIDLs.Add(idlw.IDL);
                        }
                        else {
                            // wheel cliced on QTG file
                            if(fMBUTTONUP && idlw.HasPath) {
                                if(String.Equals(Path.GetExtension(idlw.Path), QGroupOpener.EXT_QTG,
                                        StringComparison.OrdinalIgnoreCase)) {
                                    List<string> lstGrps = QGroupOpener.ReadGroupFiles(new string[] {idlw.Path},
                                            false);
                                    if(lstGrps.Count > 0) {
                                        // Open Group asynchronously
                                        Thread thread = new Thread(this.OpenGroup);
                                        thread.SetApartmentState(ApartmentState.STA);
                                        thread.IsBackground = true;
                                        thread.Start(new object[] {lstGrps.ToArray(), Control.ModifierKeys});
                                    }

                                    return true;
                                }
                            }

                            if(fEnqExec) {
                                if(idlw.HasPath) {
                                    lstFiles.Add(idlw.Path);
                                }
                            }
                        }
                    }
                }
            } // end of foreach

            if(lstIDLs.Count == 0) {
                if(fEnqExec && lstFiles.Count > 0) {
                    QTUtility.AddRecentFiles(new string[][] {lstFiles.ToArray()}, this.hwndThis);
                }
                return false;
            }
            else if(lstIDLs.Count == 1) {
                this.OpenTab(new object[] {null, modKey, lstIDLs[0]});
                return true;
            }
            else {
                // multiple item is selected...
                Thread th = new Thread(this.OpenFolders2);
                th.SetApartmentState(ApartmentState.STA);
                th.IsBackground = true;
                th.Start(new object[] {lstIDLs, modKey});

                return true;
            }
        }

        private void HandleDROPHILITED(int iItem) {
            // desktop thread

            if(iItem == -1) {
                if(this.timer_HoverSubDirTipMenu != null) {
                    this.timer_HoverSubDirTipMenu.Enabled = false;
                }
                if(this.subDirTip != null) {
                    this.subDirTip.HideMenu();
                    this.HideSubDirTip();
                }
                this.itemIndexDROPHILITED = -1;
                return;
            }

            if(this.timer_HoverSubDirTipMenu == null) {
                this.timer_HoverSubDirTipMenu = new System.Windows.Forms.Timer(this.components);
                this.timer_HoverSubDirTipMenu.Interval = QTTabBarClass.INTERVAL_SHOWMENU;
                this.timer_HoverSubDirTipMenu.Tick += new EventHandler(timer_HoverSubDirTipMenu_Tick);
            }

            this.itemIndexDROPHILITED = iItem;

            this.timer_HoverSubDirTipMenu.Enabled = false;
            this.timer_HoverSubDirTipMenu.Enabled = true;
        }


        private IntPtr GetItemPIDL(int index) {
            if(this.folderView != null) {
                IntPtr pidl;
                if(S_OK == this.folderView.Item(index, out pidl)) {
                    return pidl;
                }
            }
            return IntPtr.Zero;
        }

        private List<IntPtr> GetSelectedItemPIDL(out int index) {
            const int LVM_FIRST = 0x1000;
            const int LVM_GETNEXTITEM = (LVM_FIRST + 12);
            const int LVNI_SELECTED = 0x0002;

            index = -1;
            List<IntPtr> lst = new List<IntPtr>();
            IEnumIDList enumIDList = null;
            try {
                if(this.folderView != null) {
                    if(S_OK == this.folderView.Items(SVGIO.SELECTION, COMGUIDS.IID_IEnumIDList, out enumIDList)) {
                        IntPtr pIDL;
                        int fetched;
                        while(S_OK == enumIDList.Next(1, out pIDL, out fetched)) {
                            lst.Add(pIDL);
                        }
                    }
                }

                if(lst.Count == 1) {
                    index =
                            (int)
                                    PInvoke.SendMessage(this.hwndListView, LVM_GETNEXTITEM, (IntPtr)(-1),
                                            (IntPtr)LVNI_SELECTED);
                }
            }
            finally {
                if(enumIDList != null) {
                    Marshal.ReleaseComObject(enumIDList);
                }
            }
            return lst;
        }

        private static IntPtr GetDesktopHwnd() {
            IntPtr hwndProgman = PInvoke.FindWindowEx(IntPtr.Zero, IntPtr.Zero, "Progman", null);
            IntPtr hwndSHELLDLL_DefView = PInvoke.FindWindowEx(hwndProgman, IntPtr.Zero, "SHELLDLL_DefView", null);
            if(hwndSHELLDLL_DefView == IntPtr.Zero) {
                // seems to be reparented after desktop window created in Windows7

                IntPtr hwndWorkerW = PInvoke.FindWindowEx(IntPtr.Zero, IntPtr.Zero, "WorkerW", null);
                while(hwndWorkerW != IntPtr.Zero) {
                    hwndSHELLDLL_DefView = PInvoke.FindWindowEx(hwndWorkerW, IntPtr.Zero, "SHELLDLL_DefView", null);
                    if(hwndSHELLDLL_DefView != IntPtr.Zero) {
                        break;
                    }
                    hwndWorkerW = PInvoke.FindWindowEx(IntPtr.Zero, hwndWorkerW, "WorkerW", null);
                }
            }
            return PInvoke.FindWindowEx(hwndSHELLDLL_DefView, IntPtr.Zero, "SysListView32", null);
        }

        private int GetSelectionCount() {
            if(this.folderView != null) {
                int c;
                if(S_OK == folderView.ItemCount(SVGIO.SELECTION, out c)) {
                    return c;
                }
            }
            return 0;
        }


        private void DeleteSelection(bool fNuke) {
            int i;
            List<IntPtr> lstSelections = this.GetSelectedItemPIDL(out i);
            List<string> lstPaths = new List<string>();
            foreach(var pidl in lstSelections) {
                using(IDLWrapper idlw = new IDLWrapper(pidl)) {
                    if(idlw.CanDelete && idlw.HasPath) {
                        lstPaths.Add(idlw.Path);
                    }
                    else {
                        // not deletable object found, cancel
                        System.Media.SystemSounds.Beep.Play();
                        return;
                    }
                }
            }

            if(lstPaths.Count > 0) {
                ShellMethods.DeleteFile(lstPaths, fNuke, this.hwndListView);
            }
            else {
                // no item selected
                System.Media.SystemSounds.Beep.Play();
            }
        }



        #endregion


        #region ---------- Tip Controls ----------

        // 2 tip controls are created in desktop thread

        private bool ShowThumbnailTooltip(IntPtr pIDL, int iItem, bool fKey) {
            string path;
            StringBuilder sb = new StringBuilder(260);
            if(PInvoke.SHGetPathFromIDList(pIDL, sb)) {
                path = sb.ToString();

                if(File.Exists(path)) {
                    if(path.StartsWith(IDLWrapper.INDICATOR_NAMESPACE) || path.StartsWith(IDLWrapper.INDICATOR_NETWORK) ||
                            path.ToLower().StartsWith(@"a:\"))
                        return false;

                    string ext = Path.GetExtension(path).ToLower();
                    if(ext == ".lnk") {
                        path = ShellMethods.GetLinkTargetPath(path);
                        if(path.Length == 0)
                            return false;

                        ext = Path.GetExtension(path).ToLower();
                    }

                    if(ThumbnailTooltipForm.ExtIsSupported(ext)) {
                        if(this.thumbnailTooltip == null) {
                            this.thumbnailTooltip = new ThumbnailTooltipForm();
                            this.thumbnailTooltip.ThumbnailVisibleChanged +=
                                    new EventHandler(this.thumbnailTooltip_ThumbnailVisibleChanged);

                            this.timer_Thumbnail = new System.Windows.Forms.Timer(this.components);
                            this.timer_Thumbnail.Interval = 400;
                            this.timer_Thumbnail.Tick += new EventHandler(timer_Thumbnail_Tick);
                        }

                        if(this.thumbnailTooltip.IsShownByKey && !fKey) {
                            this.thumbnailTooltip.IsShownByKey = false;
                            return true;
                        }

                        this.thumbnailIndex = iItem;
                        this.thumbnailTooltip.IsShownByKey = fKey;

                        RECT rct = QTDesktopTool.GetLVITEMRECT(this.hwndListView, iItem, false, 0);

                        return this.thumbnailTooltip.ShowToolTip(path, new Point(rct.right - 16, rct.bottom - 8));
                    }
                }
            }
            this.HideThumbnailTooltip();
            return false;
        }

        private void ShowThumbnailTooltipForSelectedItem() {
            int iItem;
            List<IntPtr> lst = this.GetSelectedItemPIDL(out iItem);
            try {
                if(lst.Count == 1) {
                    this.ShowThumbnailTooltip(lst[0], iItem, true);
                }
            }
            finally {
                foreach(IntPtr pidl in lst) {
                    if(pidl != IntPtr.Zero) {
                        PInvoke.CoTaskMemFree(pidl);
                    }
                }
            }
        }

        private void HideThumbnailTooltip() {
            if(this.thumbnailTooltip != null && this.thumbnailTooltip.Visible) {
                this.thumbnailTooltip.HideToolTip();
            }
        }

        private void thumbnailTooltip_ThumbnailVisibleChanged(object sender, EventArgs e) {
            this.timer_Thumbnail.Enabled = false;

            if(this.thumbnailTooltip.Visible) {
                this.fThumbnailPending = false;
            }
            else {
                this.fThumbnailPending = true;
                this.timer_Thumbnail.Enabled = true;

                this.thumbnailIndex = -1;
            }
        }

        private void timer_Thumbnail_Tick(object sender, EventArgs e) {
            this.timer_Thumbnail.Enabled = false;
            this.fThumbnailPending = false;
        }

        private void timer_ThumbnailMouseHover_Tick(object sender, EventArgs e) {
            // show preview for hot item when desktop does not have focus

            this.timer_ThumbnailMouseHover.Enabled = false;

            if(this.thumbnailTooltip != null && this.thumbnailTooltip.Visible &&
                    this.thumbnailIndex_Inactive == this.thumbnailIndex)
                return;

            Point pnt = Control.MousePosition;
            PInvoke.MapWindowPoints(IntPtr.Zero, this.hwndListView, ref pnt, 1);
            if(this.thumbnailIndex_Inactive ==
                    PInvoke.ListView_HitTest(this.hwndListView, QTUtility2.Make_LPARAM(pnt.X, pnt.Y))) {
                IntPtr pidl = this.GetItemPIDL(this.thumbnailIndex_Inactive);
                try {
                    if(pidl != IntPtr.Zero) {
                        RECT rct = QTDesktopTool.GetLVITEMRECT(this.hwndListView, this.thumbnailIndex_Inactive, false, 0);
                        this.ShowThumbnailTooltip(pidl, this.thumbnailIndex_Inactive,
                                !PInvoke.PtInRect(ref rct, Control.MousePosition));
                    }
                }
                finally {
                    if(pidl != IntPtr.Zero) {
                        PInvoke.CoTaskMemFree(pidl);
                    }
                }
            }
            this.thumbnailIndex_Inactive = -1;
        }

        private void timer_Thumbnail_NoTooltip_Tick(object sender, EventArgs e) {
            this.timer_Thumbnail_NoTooltip.Enabled = false;
        }



        private bool ShowSubDirTip(IntPtr pIDL, int iItem, bool fSkipFocusCheck) {
            // desktop thread ( desktop hook -> mouse hottrack, desktop hook -> keydown )

            if(fSkipFocusCheck || Config.Bool(Scts.SubDirTipForInactiveWindow) ||
                    this.hwndListView == PInvoke.GetFocus()) {
                try {
                    string path = ShellMethods.GetDisplayName(pIDL, false);
                    byte[] idl = ShellMethods.GetIDLData(pIDL);
                    bool fQTG;

                    if(QTTabBarClass.TryMakeSubDirTipPath(ref path, ref idl, false, out fQTG)) {
                        FOLDERVIEWMODE folderViewMode = FOLDERVIEWMODE.FVM_ICON;
                                // this.folderView.GetCurrentViewMode( ref folderViewMode ); 

                        RECT rct = QTDesktopTool.GetLVITEMRECT(this.hwndListView, iItem, true, folderViewMode);
                        Point pnt = new Point(rct.right - 16, rct.bottom - 16);

                        if(this.subDirTip == null) {
                            //IntPtr hwndMessageParent = this.shellViewListener != null ? this.shellViewListener.Handle : IntPtr.Zero;

                            this.subDirTip = new SubDirTipForm(this.hwndShellView, this.hwndListView, false);
                            this.subDirTip.MenuItemClicked +=
                                    new ToolStripItemClickedEventHandler(subDirTip_MenuItemClicked);
                            this.subDirTip.MultipleMenuItemsClicked +=
                                    new EventHandler(subDirTip_MultipleMenuItemsClicked);
                            this.subDirTip.MenuItemRightClicked +=
                                    new ItemRightClickedEventHandler(subDirTip_MenuItemRightClicked);
                            this.subDirTip.MultipleMenuItemsRightClicked +=
                                    new ItemRightClickedEventHandler(subDirTip_MultipleMenuItemsRightClicked);
                        }

                        this.subDirTip.ShowSubDirTip(idl, pnt, this.hwndListView, fQTG);
                        return true;
                    }
                }
                catch(Exception ex) {
                    DebugUtil.AppendToExceptionLog(ex, null);
                }
            }
            return false;
        }

        private void HideSubDirTip() {
            // desktop thread
            if(this.subDirTip != null && this.subDirTip.Visible) {
                this.subDirTip.HideSubDirTip(false);
            }

            this.itemIndexDROPHILITED = -1;
        }

        private void HideSubDirTip_DesktopInactivated() {
            //  desktop thread
            if(this.subDirTip != null && this.subDirTip.Visible) {
                this.subDirTip.OnExplorerInactivated();
            }
        }

        private void subDirTip_MenuItemClicked(object sender, ToolStripItemClickedEventArgs e) {
            // this can run in both desktop and taskbar thread

            IntPtr hwndDialogParent = sender == this.subDirTip ? this.hwndListView : this.hwndShellTray;
                    // desktop thread or taskbar thread

            QMenuItem qmi = (QMenuItem)e.ClickedItem;

            if(qmi.Genre == MenuGenre.SubDirTip_QTGRootItem) {
                Thread thread = new Thread(this.OpenGroup);
                thread.SetApartmentState(ApartmentState.STA);
                thread.IsBackground = true;
                thread.Start(new object[] {new string[] {qmi.Name}, Control.ModifierKeys});
            }
            else if(qmi.Target == MenuTarget.Folder) {
                using(IDLWrapper idlw = new IDLWrapper(qmi.IDL)) {
                    if(!idlw.IsDeadLink(hwndDialogParent)) {
                        Keys modKey = Control.ModifierKeys;
                        if(!Config.Bool(Scts.ActivateNewTab)) {
                            if(modKey == Keys.Shift) {
                                modKey = Keys.None;
                            }
                            else if(modKey == Keys.None) {
                                modKey = Keys.Shift;
                            }
                        }

                        if(idlw.IsLink) {
                            if(!String.IsNullOrEmpty(qmi.TargetPath) &&
                                    qmi.TargetPath.StartsWith(IDLWrapper.INDICATOR_NETWORK) &&
                                            -1 == qmi.TargetPath.IndexOf(@"\", 2) &&
                                                    !ShellMethods.IsIDLNullOrEmpty(qmi.IDLTarget)) {
                                // link target is network server root ( "\\server" ), prevent opening window
                                this.OpenTab(new object[] {null, modKey, qmi.IDLTarget});
                                return;
                            }
                        }
                        this.OpenTab(new object[] {null, modKey, idlw.IDL});
                    }
                }
            }
            else {
                using(IDLWrapper idlw = new IDLWrapper(qmi.IDL)) {
                    if(!idlw.IsDeadLink(hwndDialogParent)) {
                        string work = String.Empty;

                        SHELLEXECUTEINFO sei = new SHELLEXECUTEINFO();
                        sei.cbSize = Marshal.SizeOf(sei);
                        sei.nShow = SHOWWINDOW.SHOWNORMAL;
                        sei.fMask = SEEMASK.IDLIST;
                        sei.lpIDList = idlw.PIDL;
                        sei.hwnd = hwndDialogParent;

                        if(!String.IsNullOrEmpty(qmi.Path)) {
                            work = QTUtility2.MakeDefaultWorkingDirecotryStr(qmi.Path);
                            if(work.Length > 0) {
                                sei.lpDirectory = Marshal.StringToCoTaskMemUni(work);
                            }
                        }

                        try {
                            if(PInvoke.ShellExecuteEx(ref sei)) {
                                QTUtility.AddRecentFiles(
                                        new string[][] {
                                                work.Length > 0
                                                        ? new string[] {qmi.Path, String.Empty, work}
                                                        : new string[] {qmi.Path}
                                        }, this.hwndThis);
                            }
                        }
                        finally {
                            if(sei.lpDirectory != IntPtr.Zero) {
                                Marshal.FreeCoTaskMem(sei.lpDirectory);
                            }
                        }
                    }
                }
            }
        }

        private void subDirTip_MultipleMenuItemsClicked(object sender, EventArgs e) {
            // this can run in both desktop and taskbar thread

            SubDirTipForm sdtf = (SubDirTipForm)sender;
            IntPtr hwndDialogParent = sdtf == this.subDirTip ? this.hwndListView : this.hwndShellTray;

            // SubDirTip_QTGRootItem
            string[] arrGrps = sdtf.ExecutedGroups;
            if(arrGrps.Length > 0) {
                Thread thread = new Thread(this.OpenGroup);
                thread.SetApartmentState(ApartmentState.STA);
                thread.IsBackground = true;
                thread.Start(new object[] {arrGrps, Control.ModifierKeys});
            }

            // folders
            if(sdtf.ExecutedIDLs.Count > 0) {
                List<byte[]> lstIDLs = new List<byte[]>();

                foreach(byte[] idl in sdtf.ExecutedIDLs) {
                    using(IDLWrapper idlw = new IDLWrapper(idl, false)) {
                        if(idlw.IsLink) {
                            IDLWrapper idlwTarget;
                            if(idlw.TryGetLinkTarget(hwndDialogParent, out idlwTarget)) {
                                if(idlwTarget.IsFolder) {
                                    if(idlw.IsFolder) {
                                        lstIDLs.Add(idlw.IDL);
                                    }
                                    else {
                                        lstIDLs.Add(idlwTarget.IDL);
                                    }
                                }
                            }
                        }
                        else {
                            lstIDLs.Add(idlw.IDL);
                        }
                    }
                }

                Keys modKey = Control.ModifierKeys;
                if(!Config.Bool(Scts.ActivateNewTab)) {
                    if(modKey == Keys.Shift) {
                        modKey = Keys.None;
                    }
                    else if(modKey == Keys.None) {
                        modKey = Keys.Shift;
                    }
                }

                if(lstIDLs.Count == 1) {
                    this.OpenTab(new object[] {null, modKey, lstIDLs[0]});
                }
                else if(lstIDLs.Count > 1) {
                    Thread thread = new Thread(this.OpenFolders2);
                    thread.SetApartmentState(ApartmentState.STA);
                    thread.IsBackground = true;
                    thread.Start(new object[] {lstIDLs, modKey});
                }
            }
        }

        private void subDirTip_MenuItemRightClicked(object sender, ItemRightClickedEventArgs e) {
            // the calling thread can be taskBar or desktop

            using(IDLWrapper idlw = new IDLWrapper(((QMenuItem)e.ClickedItem).IDL, false)) {
                if(sender == this.subDirTip) {
                    e.Result = ShellMethods.PopUpShellContextMenu(idlw, e.IsKey ? e.Point : Control.MousePosition,
                            ref this.iContextMenu2_Desktop, this.subDirTip.Handle, false);
                }
                else // subDirTip_TB
                {
                    e.Result = ShellMethods.PopUpShellContextMenu(idlw, e.IsKey ? e.Point : Control.MousePosition,
                            ref this.iContextMenu2, this.subDirTip_TB.Handle, false);
                }

                if(e.Result == MC.COMMANDID_OPENPARENT) {
                    using(IDLWrapper idlwParent = new IDLWrapper(ShellMethods.GetParentIDL(idlw.PIDL))) {
                        if(idlwParent.Available) {
                            Thread thread = new Thread(this.OpenTab);
                            thread.SetApartmentState(ApartmentState.STA);
                            thread.IsBackground = true;
                            thread.Start(new object[] {null, Control.ModifierKeys, idlwParent.IDL});
                        }
                    }
                }
            }
        }

        private void subDirTip_MultipleMenuItemsRightClicked(object sender, ItemRightClickedEventArgs e) {
            if(sender == this.subDirTip) {
                e.Result = ShellMethods.PopUpShellContextMenu(this.subDirTip.ExecutedIDLs,
                        e.IsKey ? e.Point : Control.MousePosition, ref this.iContextMenu2_Desktop, this.subDirTip.Handle);
            }
            else if(sender == this.subDirTip_TB) {
                e.Result = ShellMethods.PopUpShellContextMenu(this.subDirTip_TB.ExecutedIDLs,
                        e.IsKey ? e.Point : Control.MousePosition, ref this.iContextMenu2, this.subDirTip_TB.Handle);
            }
        }


        private void timer_HoverSubDirTipMenu_Tick(object sender, EventArgs e) {
            // drop hilited and MouseHoverTime elapsed 

            // desktop thread
            this.timer_HoverSubDirTipMenu.Enabled = false;
            int iItem = this.itemIndexDROPHILITED;

            if(Control.MouseButtons != MouseButtons.None) {
                Point pnt = Control.MousePosition;
                PInvoke.MapWindowPoints(IntPtr.Zero, this.hwndListView, ref pnt, 1);
                if(iItem == PInvoke.ListView_HitTest(this.hwndListView, QTUtility2.Make_LPARAM(pnt.X, pnt.Y))) {
                    using(IDLWrapper idlw = new IDLWrapper(this.GetItemPIDL(iItem))) {
                        if(idlw.Available) {
                            if(this.subDirTip != null) {
                                this.subDirTip.HideMenu();
                            }

                            if(!String.Equals(idlw.Path, CLSIDSTR_TRASHBIN, StringComparison.OrdinalIgnoreCase)) {
                                if(this.ShowSubDirTip(idlw.PIDL, iItem, true)) {
                                    this.itemIndexDROPHILITED = iItem;
                                    PInvoke.SetFocus(this.hwndListView);
                                    PInvoke.SetForegroundWindow(this.hwndListView);
                                    this.HideThumbnailTooltip();
                                    this.subDirTip.ShowMenuForDropHilited(GetDesktopIconSize());
                                    return;
                                }
                            }
                        }
                    }
                }

                if(this.subDirTip != null) {
                    if(this.subDirTip.IsMouseOnMenus) {
                        this.itemIndexDROPHILITED = -1;
                        return;
                    }
                }
            }

            this.HideSubDirTip();
        }


        private static RECT GetLVITEMRECT(IntPtr hwndListView, int iItem, bool fSubDirTip, FOLDERVIEWMODE fvm) {
            // get the bounding rectangle of item specified by iItem, in the screen coordinates.
            // fSubDirTip	true to get RECT depending on view style, false to get RECT by LVIR_BOUNDS

            const uint LVM_FIRST = 0x1000;
            const uint LVM_GETVIEW = (LVM_FIRST + 143);
            const uint LVM_GETITEMW = (LVM_FIRST + 75);
            const uint LVM_GETSTRINGWIDTHW = (LVM_FIRST + 87);
            const uint LVM_GETITEMSPACING = (LVM_FIRST + 51);
            const int LVIR_BOUNDS = 0;
            const int LVIR_ICON = 1;
            const int LVIR_LABEL = 2;
            const int LV_VIEW_ICON = 0x0000;
            const int LV_VIEW_DETAILS = 0x0001;
            const int LV_VIEW_LIST = 0x0003;
            const int LV_VIEW_TILE = 0x0004;
            const int LVIF_TEXT = 0x00000001;

            int view = (int)PInvoke.SendMessage(hwndListView, LVM_GETVIEW, IntPtr.Zero, IntPtr.Zero);
            int code = view == LV_VIEW_DETAILS ? LVIR_LABEL : LVIR_BOUNDS;

            bool fIcon = false; // for XP
            bool fList = false; // for XP

            if(fSubDirTip) {
                switch(view) {
                    case LV_VIEW_ICON:
                        fIcon = !QTUtility.IsVista;
                        code = LVIR_ICON;
                        break;

                    case LV_VIEW_DETAILS:
                        code = LVIR_LABEL;
                        break;

                    case LV_VIEW_LIST:
                        if(!QTUtility.IsVista) {
                            fList = true;
                            code = LVIR_ICON;
                        }
                        else {
                            code = LVIR_LABEL;
                        }
                        break;

                    case LV_VIEW_TILE:
                        code = LVIR_ICON;
                        break;

                    default:
                        // Here only in case of Vista LV_VIEW_SMALLICON.
                        code = LVIR_BOUNDS;
                        break;
                }
            }

            // get item rectangle
            RECT rct = PInvoke.ListView_GetItemRect(hwndListView, iItem, 0, code);

            // convert to screen coordinates
            PInvoke.MapWindowPoints(hwndListView, IntPtr.Zero, ref rct, 2);

            // adjust rct
            // these magic numbers have no logical meanings
            if(fIcon) {
                // XP, subdirtip.
                // THUMBNAIL, THUMBSTRIP or ICON.
                if(fvm == FOLDERVIEWMODE.FVM_THUMBNAIL || fvm == FOLDERVIEWMODE.FVM_THUMBSTRIP) {
                    rct.right -= 13;
                }
                else // fvm == FVM_ICON
                {
                    int currentIconSpacing =
                            (int)(long)PInvoke.SendMessage(hwndListView, LVM_GETITEMSPACING, IntPtr.Zero, IntPtr.Zero);
                    Size sz = SystemInformation.IconSize;
                    rct.right = rct.left + (((currentIconSpacing & 0xFFFF) - sz.Width)/2) + sz.Width + 8;
                    rct.bottom = rct.top + sz.Height + 6;
                }
            }
            else if(fList) {
                // XP, subdirtip.
                // calculate item text rectangle
                LVITEM lvitem = new LVITEM();
                lvitem.pszText = Marshal.AllocCoTaskMem(520);
                lvitem.cchTextMax = 260;
                lvitem.iItem = iItem;
                lvitem.mask = LVIF_TEXT;
                IntPtr pLI = Marshal.AllocCoTaskMem(Marshal.SizeOf(lvitem));
                Marshal.StructureToPtr(lvitem, pLI, false);

                PInvoke.SendMessage(hwndListView, LVM_GETITEMW, IntPtr.Zero, pLI);

                int w = (int)PInvoke.SendMessage(hwndListView, LVM_GETSTRINGWIDTHW, IntPtr.Zero, lvitem.pszText);
                w += 20;

                Marshal.FreeCoTaskMem(lvitem.pszText);
                Marshal.FreeCoTaskMem(pLI);

                rct.right += w;
                rct.top += 2;
                rct.bottom += 2;
            }

            return rct;
        }

        private static int GetDesktopIconSize() {
            const string KEYNAME = @"Software\Microsoft\Windows\Shell\Bags\1\Desktop";
            const string VALNAME = "IconSize";

            using(RegistryKey rk = Registry.CurrentUser.OpenSubKey(KEYNAME)) {
                if(rk != null) {
                    return (int)rk.GetValue(VALNAME, 48);
                }
            }
            return 48;
        }

        #endregion


        #region ---------- Settings ----------

        private void ReadSetting() {
            this.lstItemOrder.Clear();
            // 0 < value < ITEMTYPE_COUNT
            this.lstItemOrder.Add(QTUtility.ValidateValueRange(Config.Get(Scts.Desktop_FirstItem), 0, ITEMTYPE_COUNT - 1));
            this.lstItemOrder.Add(QTUtility.ValidateValueRange(Config.Get(Scts.Desktop_SecondItem), 0,
                    ITEMTYPE_COUNT - 1));
            this.lstItemOrder.Add(QTUtility.ValidateValueRange(Config.Get(Scts.Desktop_ThirdItem), 0, ITEMTYPE_COUNT - 1));
            this.lstItemOrder.Add(QTUtility.ValidateValueRange(Config.Get(Scts.Desktop_FourthItem), 0,
                    ITEMTYPE_COUNT - 1));

            for(int i = 0; i < ITEMTYPE_COUNT; i++) {
                if(!this.lstItemOrder.Contains(i)) {
                    this.lstItemOrder.Add(i);
                }
            }

            this.ExpandState[0] = Config.Bool(Scts.Desktop_GroupExpanded);
            this.ExpandState[1] = Config.Bool(Scts.Desktop_RecentTabExpanded);
            this.ExpandState[2] = Config.Bool(Scts.Desktop_ApplicationExpanded);
            this.ExpandState[3] = Config.Bool(Scts.Desktop_RecentFileExpanded);

            using(RegistryKey rkUser = Registry.CurrentUser.OpenSubKey(REGNAME.KEY_USERROOT)) {
                if(rkUser != null) {
                    this.WidthOfBar = (int)rkUser.GetValue(REGNAME.VAL_DT_WIDTH, 80);
                }
            }
        }

        private void SaveSetting() {
            Config.Set(Scts.Desktop_FirstItem, this.lstItemOrder[0]);
            Config.Set(Scts.Desktop_SecondItem, this.lstItemOrder[1]);
            Config.Set(Scts.Desktop_ThirdItem, this.lstItemOrder[2]);
            Config.Set(Scts.Desktop_FourthItem, this.lstItemOrder[3]);
            Config.Set(Scts.Desktop_GroupExpanded, this.ExpandState[0]);
            Config.Set(Scts.Desktop_RecentTabExpanded, this.ExpandState[1]);
            Config.Set(Scts.Desktop_ApplicationExpanded, this.ExpandState[2]);
            Config.Set(Scts.Desktop_RecentFileExpanded, this.ExpandState[3]);
            Config.Set(Scts.Desktop_TaskBarDoubleClickEnabled, this.tsmiTaskBar.Checked);
            Config.Set(Scts.Desktop_DesktopDoubleClickEnabled, this.tsmiDesktop.Checked);
            Config.Set(Scts.Desktop_LockMenu, this.tsmiLockItems.Checked);
            Config.Set(Scts.Desktop_TitleBackground, this.tsmiVSTitle.Checked);
            Config.Set(Scts.Desktop_IncludeGroup, this.tsmiOnGroup.Checked);
            Config.Set(Scts.Desktop_IncludeRecentTab, this.tsmiOnHistory.Checked);
            Config.Set(Scts.Desktop_IncludeApplication, this.tsmiOnUserApps.Checked);
            Config.Set(Scts.Desktop_IncludeRecentFile, this.tsmiOnRecentFile.Checked);
            Config.Set(Scts.Desktop_1ClickMenu, this.tsmiOneClick.Checked);
            Config.Set(Scts.Desktop_EnbleApplicationShortcuts, this.tsmiAppKeys.Checked);

            using(RegistryKey rkUser = Registry.CurrentUser.CreateSubKey(REGNAME.KEY_USERROOT)) {
                if(rkUser != null) {
                    Config.Save(rkUser);
                    rkUser.SetValue(REGNAME.VAL_DT_WIDTH, this.Width);
                }
            }
        }

        private void contextMenuForSetting_ItemClicked(object sender, ToolStripItemClickedEventArgs e) {
            if(e.ClickedItem is ToolStripSeparator)
                return;

            if(e.ClickedItem == this.tsmiTaskBar) {
                this.tsmiTaskBar.Checked = !this.tsmiTaskBar.Checked;
            }
            else if(e.ClickedItem == this.tsmiDesktop) {
                this.tsmiDesktop.Checked = !this.tsmiDesktop.Checked;
            }
            else if(e.ClickedItem == this.tsmiLockItems) {
                this.tsmiLockItems.Checked = !this.tsmiLockItems.Checked;

                this.contextMenu.ReorderEnabled =
                        this.ddmrGroups.ReorderEnabled = !this.tsmiLockItems.Checked;
                this.ddmrUserapps.SetChildrenReorderEnabled(!this.tsmiLockItems.Checked);
            }

            else if(e.ClickedItem == this.tsmiOnGroup) {
                this.tsmiOnGroup.Checked = !this.tsmiOnGroup.Checked;

                this.lstRefreshRequired[ITEMINDEX_GROUP] = true;
            }
            else if(e.ClickedItem == this.tsmiOnHistory) {
                this.tsmiOnHistory.Checked = !this.tsmiOnHistory.Checked;

                this.lstRefreshRequired[ITEMINDEX_RECENTTAB] = true;
            }
            else if(e.ClickedItem == this.tsmiOnUserApps) {
                this.tsmiOnUserApps.Checked = !this.tsmiOnUserApps.Checked;

                this.lstRefreshRequired[ITEMINDEX_APPLAUNCHER] = true;
            }
            else if(e.ClickedItem == this.tsmiOnRecentFile) {
                this.tsmiOnRecentFile.Checked = !this.tsmiOnRecentFile.Checked;

                this.lstRefreshRequired[ITEMINDEX_RECENTFILE] = true;
            }

            else if(e.ClickedItem == this.tsmiVSTitle) {
                this.tsmiVSTitle.Checked = !this.tsmiVSTitle.Checked;

                TitleMenuItem.DrawBackground = this.tsmiVSTitle.Checked;
            }
            else if(e.ClickedItem == this.tsmiOneClick) {
                this.tsmiOneClick.Checked = !this.tsmiOneClick.Checked;
            }
            else if(e.ClickedItem == this.tsmiAppKeys) {
                this.tsmiAppKeys.Checked = !this.tsmiAppKeys.Checked;
            }

            this.SaveSetting();
        }

        private void RefreshStringResources() {
            string[] ResTaskbar = QTUtility.StringResourcesDic["TaskBar_Menu"];

            this.tsmiTaskBar.Text = ResTaskbar[0];
            this.tsmiDesktop.Text = ResTaskbar[1];
            this.tsmiLockItems.Text = ResTaskbar[2];
            this.tsmiVSTitle.Text = ResTaskbar[3];
            this.tsmiOneClick.Text = ResTaskbar[4];
            this.tsmiAppKeys.Text = ResTaskbar[5];

            string[] titles = QTUtility.StringResourcesDic["TaskBar_Titles"];

            this.tsmiOnGroup.Text =
                    this.tmiLabel_Group.Text =
                            this.tmiGroup.Text = titles[0];

            this.tsmiOnHistory.Text =
                    this.tmiHistory.Text =
                            this.tmiLabel_History.Text = titles[1];

            this.tsmiOnUserApps.Text =
                    this.tmiUserApp.Text =
                            this.tmiLabel_UserApp.Text = titles[2];

            this.tsmiOnRecentFile.Text =
                    this.tmiRecentFile.Text =
                            this.tmiLabel_RecentFile.Text = titles[3];
        }

        #endregion


        #region ---------- Event Handlers ----------


        private void desktopTool_MouseClick(object sender, MouseEventArgs e) {
            // single click mode
            if(e.Button == MouseButtons.Left && Config.Bool(Scts.Desktop_1ClickMenu)) {
                this.ShowMenu(Control.MousePosition);
            }
        }

        private void desktopTool_MouseDoubleClick(object sender, MouseEventArgs e) {
            if(e.Button == MouseButtons.Left && !Config.Bool(Scts.Desktop_1ClickMenu)) {
                this.ShowMenu(Control.MousePosition);
            }
        }


        private void contextMenu_Closing(object sender, ToolStripDropDownClosingEventArgs e) {
            if(this.fCancelClosing) {
                e.Cancel = true;
                this.fCancelClosing = false;
            }
            else {
                if(this.fRootReordered) {
                    this.fRootReordered = false;

                    List<int> lst = new List<int>();
                    for(int i = 0; i < this.contextMenu.Items.Count; i++) {
                        if(lst.Count == ITEMTYPE_COUNT)
                            break;

                        ToolStripItem item = this.contextMenu.Items[i];

                        if(item is TitleMenuItem) {
                            if(item == this.tmiGroup || item == this.tmiLabel_Group) {
                                lst.Add(ITEMINDEX_GROUP);
                                continue;
                            }
                            if(item == this.tmiHistory || item == this.tmiLabel_History) {
                                lst.Add(ITEMINDEX_RECENTTAB);
                                continue;
                            }
                            if(item == this.tmiUserApp || item == this.tmiLabel_UserApp) {
                                lst.Add(ITEMINDEX_APPLAUNCHER);
                                continue;
                            }
                            if(item == this.tmiRecentFile || item == this.tmiLabel_RecentFile) {
                                lst.Add(ITEMINDEX_RECENTFILE);
                                continue;
                            }
                        }
                    }
                    if(lst.Count < 4) {
                        for(int i = 0; i < this.lstItemOrder.Count; i++) {
                            if(!lst.Contains(this.lstItemOrder[i])) {
                                lst.Add(this.lstItemOrder[i]);
                            }
                        }
                    }

                    this.lstItemOrder.Clear();
                    for(int i = 0; i < lst.Count; i++) {
                        this.lstItemOrder.Add(lst[i]);
                    }

                    this.SaveSetting();
                }
            }
        }

        private bool fRootReordered;

        private void contextMenu_ReorderFinished(object sender, ToolStripItemClickedEventArgs e) {
            this.fRootReordered = e.ClickedItem is TitleMenuItem;
            QMenuItem qmi = e.ClickedItem as QMenuItem;
            if(qmi == null) {
                return;
            }

            if(qmi.Genre == MenuGenre.Group) {
                this.lstGroupItems.Clear();

                using(
                        RegistryKey rkGroup =
                                Registry.CurrentUser.CreateSubKey(REGNAME.KEY_USERROOT + @"\" + REGNAME.KEY_GROUPS)) {
                    int separatorIndex = 1;

                    foreach(string valName in rkGroup.GetValueNames()) {
                        rkGroup.DeleteValue(valName, false);
                    }

                    foreach(ToolStripItem tsi in this.contextMenu.Items) {
                        QMenuItem qmiSub = tsi as QMenuItem;
                        if(qmiSub != null) {
                            if(qmiSub.Genre == MenuGenre.Group) {
                                rkGroup.SetValue(qmiSub.Text, qmiSub.GroupItemInfo.KeyboardShortcut);

                                this.lstGroupItems.Add(qmiSub);
                            }
                        }
                        else if(tsi is ToolStripSeparator && tsi.Name == TSS_NAME_GRP) {
                            rkGroup.SetValue("Separator" + (separatorIndex++), 0);

                            this.lstGroupItems.Add(tsi);
                        }
                    }
                }


                QTUtility.RebuildGroupsDic();
                QTUtility.flagManager_Group.SetFlags();
                InstanceManager.SyncProcesses(MC.MLW_UPDATE_GROUP, IntPtr.Zero, IntPtr.Zero, this.hwndThis);
            }
            else if(qmi.Genre == MenuGenre.Application) {
                // need to sync the list
                this.lstUserAppItems.Clear();

                using(
                        RegistryKey rkUserApp =
                                Registry.CurrentUser.CreateSubKey(REGNAME.KEY_USERROOT + @"\" + REGNAME.KEY_APPLICATIONS)
                        ) {
                    foreach(string valName in rkUserApp.GetValueNames()) {
                        rkUserApp.DeleteValue(valName, false);
                    }

                    int separatorIndex = 1;
                    string[] separatorVal = new string[] {String.Empty, String.Empty, String.Empty, String.Empty};

                    foreach(ToolStripItem tsi in this.contextMenu.Items) {
                        QMenuItem qmiSub = tsi as QMenuItem;
                        if(qmiSub != null) {
                            if(qmiSub.Genre == MenuGenre.Application) {
                                if(qmiSub.Target == MenuTarget.VirtualFolder) {
                                    rkUserApp.SetValue(tsi.Text, new byte[0]);
                                }
                                else {
                                    rkUserApp.SetValue(tsi.Text, QTUtility.dicUserApps[tsi.Text]);
                                }

                                this.lstUserAppItems.Add(tsi);
                            }
                        }
                        else if(tsi is ToolStripSeparator && tsi.Name == TSS_NAME_APP) {
                            rkUserApp.SetValue("Separator" + (separatorIndex++), separatorVal);

                            this.lstUserAppItems.Add(tsi);
                        }
                    }

                    // refresh dictionary
                    QTUtility.dicUserApps.Clear();
                    foreach(string valueName in rkUserApp.GetValueNames()) {
                        if(valueName.Length > 0) {
                            string[] appVal = rkUserApp.GetValue(valueName) as string[];
                            if(appVal != null) {
                                if(appVal.Length > 3) {
                                    QTUtility.dicUserApps.Add(valueName, appVal);
                                }
                            }
                            else {
                                using(RegistryKey rkUserAppSub = rkUserApp.OpenSubKey(valueName, false)) {
                                    if(rkUserAppSub != null) {
                                        QTUtility.dicUserApps.Add(valueName, null);
                                    }
                                }
                            }
                        }
                    }
                }

                QTUtility.flagManager_AppLauncher.SetFlags();
                InstanceManager.SyncProcesses(MC.MLW_UPDATE_APPLICATION, IntPtr.Zero, IntPtr.Zero, this.hwndThis);
            }
        }

        private void dropDowns_ItemClicked(object sender, ToolStripItemClickedEventArgs e) {
            // sicne menu is created on taskbar thread 
            // this runs on the thread

            TitleMenuItem tmi = e.ClickedItem as TitleMenuItem;
            if(tmi != null) {
                if(tmi.IsOpened) {
                    // Close
                    this.OnLabelTitleClickedToClose(tmi.Genre);
                }
                else {
                    // Open
                    this.OnSubMenuTitleClickedToOpen(tmi.Genre);
                }

                return;
            }

            QMenuItem qmi = e.ClickedItem as QMenuItem;
            if(qmi == null) {
                return;
            }

            if(qmi.Genre == MenuGenre.Group) {
                Keys modKeys = Control.ModifierKeys;
                string groupName = qmi.Text;

                //if( modKeys == ( Keys.Control | Keys.Shift ) )
                //{
                //    if( QTUtility.lstStartUpGroups.Contains( groupName ) )
                //    {
                //        QTUtility.lstStartUpGroups.Remove( groupName );
                //    }
                //    else
                //    {
                //        QTUtility.lstStartUpGroups.Add( groupName );
                //    }

                //    using( RegistryKey rkUser = Registry.CurrentUser.CreateSubKey( REGNAME.KEY_USERROOT ) )
                //    {
                //        if( rkUser != null )
                //        {
                //            string startup = String.Empty;

                //            foreach( string grp in QTUtility.lstStartUpGroups )
                //            {
                //                startup += grp + ";";
                //            }

                //            startup = startup.TrimEnd( QTUtility.SEPARATOR_CHAR );
                //            rkUser.SetValue( "StartUpGroups", startup );
                //        }
                //    }
                //    return;
                //}

                // Open Group asynchronously

                Thread thread = new Thread(this.OpenGroup);
                thread.SetApartmentState(ApartmentState.STA);
                thread.IsBackground = true;
                thread.Start(new object[] {new string[] {groupName}, Control.ModifierKeys});

            }
            else if(qmi.Genre == MenuGenre.RecentlyClosedTab) {
                // Undo closed asynchronously

                Thread thread = new Thread(this.OpenTab);
                thread.SetApartmentState(ApartmentState.STA);
                thread.IsBackground = true;
                thread.Start(new object[] {null, Control.ModifierKeys, qmi.IDL});

            }
            else if(qmi.Genre == MenuGenre.Application && qmi.Target == MenuTarget.File) {
                // User apps
                if(!qmi.MenuItemArguments.TokenReplaced) {
                    AppLauncher.ReplaceAllTokens(qmi.MenuItemArguments,
                            Environment.GetFolderPath(Environment.SpecialFolder.Desktop));
                }

                AppLauncher.Execute(qmi.MenuItemArguments, this.Handle);
            }
            else if(qmi.Genre == MenuGenre.RecentFile) {
                AppLauncher.Execute(qmi.MenuItemArguments, this.Handle);
            }
        }

        private void dropDowns_ItemRightClicked(object sender, ItemRightClickedEventArgs e) {
            QMenuItem qmi = e.ClickedItem as QMenuItem;

            // Is not valid menu item, or Virutal folder, do nothing
            if(qmi == null || qmi.Target == MenuTarget.VirtualFolder) {
                // cancel closing.
                e.Result = MC.COMMANDID_USERCANCEL;
                return;
            }

            Point pnt = e.IsKey ? e.Point : Control.MousePosition;

            if(qmi.Genre == MenuGenre.Group) {
                // Group
                DropDownMenuReorderable ddmr = (DropDownMenuReorderable)sender;
                int commandID;
                byte[] idl = MenuUtility.TrackGroupContextMenu(e.ClickedItem.Text, pnt, ddmr.Handle, false,
                        out commandID);

                if(!ShellMethods.IsIDLNullOrEmpty(idl)) {
                    this.OpenTab(new object[] {null, Control.ModifierKeys, idl});
                }
                else {
                    if(commandID == MC.COMMANDID_CREATEQTG) {
                        ddmr.ForceClose();
                        QGroupOpener.CreateGroupFile(qmi.Text, this);
                        e.Result = MC.COMMANDID_CREATEQTG;
                    }
                            //else if( commandID == MC.COMMANDID_REMOVEGROUPITEM )
                            //{
                            //}
                    else {
                        // cancel closing.
                        e.Result = MC.COMMANDID_USERCANCEL;
                    }
                }
            }
            else {
                // RecentlyClosed, User apps, Recent files.

                //						menu items can be removed	|	menu items have idl data
                // RecentClosedTab					Y				|				Y
                // User Apps						N				|				N
                // RecentFiles						Y				|				N

                bool fCanRemove = qmi.Genre != MenuGenre.Application;

                using(
                        IDLWrapper idlw = qmi.Genre == MenuGenre.RecentlyClosedTab
                                ? new IDLWrapper(qmi.IDL, false)
                                : new IDLWrapper(qmi.Path)) {
                    e.Result = ShellMethods.PopUpShellContextMenu(idlw, pnt, ref this.iContextMenu2,
                            ((DropDownMenuReorderable)sender).Handle, fCanRemove);

                    if(e.Result == MC.COMMANDID_OPENPARENT) {
                        using(IDLWrapper idlwParent = new IDLWrapper(ShellMethods.GetParentIDL(idlw.PIDL))) {
                            if(idlwParent.Available) {
                                Thread thread = new Thread(this.OpenTab);
                                thread.SetApartmentState(ApartmentState.STA);
                                thread.IsBackground = true;
                                thread.Start(new object[] {null, Control.ModifierKeys, idlwParent.IDL});
                            }
                        }
                    }
                    else if(e.Result == MC.COMMANDID_REMOVEITEM) {
                        if(qmi.Genre == MenuGenre.RecentlyClosedTab) {
                            QTUtility.RemoveRecentTabs(new LogData[] {new LogData(qmi.IDL, qmi.Path)}, false,
                                    this.hwndThis);
                            this.lstUndoClosedItems.Remove(qmi);
                        }
                        else if(qmi.Genre == MenuGenre.RecentFile) {
                            QTUtility.RemoveRecentFiles(new string[][] {qmi.MenuItemArguments.arrRecentFileData}, false,
                                    this.hwndThis);
                            this.lstRecentFileItems.Remove(qmi);
                        }

                        qmi.Dispose();
                    }
                }
            }
        }

        private void dropDowns_ReorderFinished(object sender, ToolStripItemClickedEventArgs e) {
            if(sender == this.ddmrGroups) {
                QTUtility.OnReorderFinished_Group(this.ddmrGroups.Items, this.hwndThis);

                // need to sync the list
                this.lstGroupItems.Clear();
                foreach(ToolStripItem tsi in this.ddmrGroups.Items) {
                    this.lstGroupItems.Add(tsi);
                }
            }
            else if(sender == this.ddmrUserapps) {
                QTUtility.OnReorderFinished_AppLauncher(this.ddmrUserapps.Items, this.hwndThis);
                QTUtility.flagManager_AppLauncher.SetFlags();

                // need to sync the list
                this.lstUserAppItems.Clear();
                foreach(ToolStripItem tsi in this.ddmrUserapps.Items) {
                    this.lstUserAppItems.Add(tsi);
                }
            }
        }

        private void directoryMenuItems_DoubleClick(object sender, EventArgs e) {
            // DirectoryMenuItem is clicked.
            // It's guaranteed that sender is DirectoryMenuItem.

            string path = ((DirectoryMenuItem)sender).Path;
            if(Directory.Exists(path)) {
                try {
                    this.OpenTab(new object[] {path, Control.ModifierKeys}); //idlize??
                }
                catch {
                    MessageBox.Show(
                            "Operation failed.\r\nPlease make sure the folder exists or you have permission to access to:\r\n\r\n\t" +
                                    path, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }


        private void tsmiExperimental_DropDownOpening(object sender, EventArgs e) {
            if(this.tsmiExperimental.DropDownItems.Count == 1) {
                this.tsmiExperimental.DropDownItems[0].Dispose();
                this.tsmiExperimental.DropDownOpening -= tsmiExperimental_DropDownOpening;

                this.tsmiExperimental.DropDown.SuspendLayout();
                List<string> lst = new List<string>();
                if(System.Globalization.CultureInfo.CurrentCulture.Parent.Name == "ja") {
                    lst.AddRange(new string[] {"ƒŠƒXƒg", "Ú×", "•À‚×‚Ä•\Ž¦", "–ß‚·"});
                }
                else {
                    lst.AddRange(new string[] {"List", "Details", "Tiles", "Default"});
                }
                for(int i = 0; i < lst.Count; i++) {
                    this.tsmiExperimental.DropDown.Items.Add(lst[i]);
                }
                this.tsmiExperimental.DropDown.ResumeLayout();
            }
        }

        private void tsmiExperimental_DropDownItemClicked(object sender, ToolStripItemClickedEventArgs e) {
            if(this.folderView != null) {
                int index = this.tsmiExperimental.DropDown.Items.IndexOf(e.ClickedItem);
                switch(index) {
                    case 0:
                        this.folderView.SetCurrentViewMode(FOLDERVIEWMODE.FVM_LIST);
                        break;

                    case 1:
                        this.folderView.SetCurrentViewMode(FOLDERVIEWMODE.FVM_DETAILS);
                        break;

                    case 2:
                        this.folderView.SetCurrentViewMode(FOLDERVIEWMODE.FVM_TILE);
                        break;

                    case 3:
                        this.folderView.SetCurrentViewMode(FOLDERVIEWMODE.FVM_ICON);
                        break;
                }
            }
        }



        #endregion


        #region ---------- Menu Creation ----------


        private bool[] BuildMenuItems() {
            List<bool> lst = new List<bool> {false, false, false, false};

            // group
            if(QTUtility.flagManager_Group.GetFlag(this.lstGroupItems) || this.lstRefreshRequired[ITEMINDEX_GROUP]) {
                this.lstRefreshRequired[ITEMINDEX_GROUP] = false;

                // clear items
                foreach(ToolStripItem tsi in this.lstGroupItems) {
                    tsi.Dispose();
                }
                this.lstGroupItems.Clear();

                if(Config.Bool(Scts.Desktop_IncludeGroup)) {
                    lst[ITEMINDEX_GROUP] = true;

                    foreach(string groupName in QTUtility.dicGroups.Keys) {
                        GroupItemInfo gi = QTUtility.dicGroups[groupName];
                        if(gi.ItemCount == 0) {
                            ToolStripSeparator tss = new ToolStripSeparator();
                            tss.Name = TSS_NAME_GRP;
                            this.lstGroupItems.Add(tss);
                            continue;
                        }

                        QMenuItem qmi = new QMenuItem(groupName, MenuGenre.Group);
                        qmi.SetImageReservationKey(gi.EnsurePath(0), null);
                        qmi.GroupItemInfo = gi;

                        if(QTUtility.lstStartUpGroups.Contains(groupName)) {
                            MenuUtility.SetStartUpMenuFont(qmi);
                        }

                        this.lstGroupItems.Add(qmi);
                    }
                }
                else {
                    this.contextMenu.Items.Remove(this.tmiLabel_Group);
                    this.contextMenu.Items.Remove(this.tmiGroup);
                }
            }

            // recent tab
            {
                this.lstRefreshRequired[ITEMINDEX_RECENTTAB] = false;

                // clear items
                foreach(ToolStripItem item in this.lstUndoClosedItems) {
                    item.Dispose();
                }
                this.lstUndoClosedItems.Clear();

                if(Config.Bool(Scts.Desktop_IncludeRecentTab)) {
                    lst[ITEMINDEX_RECENTTAB] = true;

                    this.lstUndoClosedItems = MenuUtility.CreateRecentlyClosedItems(null, this.hwndShellTray);
                }
                else {
                    this.contextMenu.Items.Remove(this.tmiLabel_History);
                    this.contextMenu.Items.Remove(this.tmiHistory);
                }
            }

            // application launcher
            if(QTUtility.flagManager_AppLauncher.GetFlag(this.lstUserAppItems) ||
                    this.lstRefreshRequired[ITEMINDEX_APPLAUNCHER]) {
                this.lstRefreshRequired[ITEMINDEX_APPLAUNCHER] = false;

                // clear items
                foreach(ToolStripItem item in this.lstUserAppItems) {
                    item.Dispose();
                }
                this.lstUserAppItems.Clear();

                if(Config.Bool(Scts.Desktop_IncludeApplication)) {
                    lst[ITEMINDEX_APPLAUNCHER] = true;

                    this.lstUserAppItems = MenuUtility.CreateAppLauncherItems(new EventPack(this.Handle,
                            new ItemRightClickedEventHandler(this.dropDowns_ItemRightClicked),
                            new EventHandler(this.directoryMenuItems_DoubleClick),
                            null,
                            new SubDirTipCreator(this.CreateSubDirTip),
                            true),
                            !Config.Bool(Scts.Desktop_LockMenu));
                }
                else {
                    this.contextMenu.Items.Remove(this.tmiLabel_UserApp);
                    this.contextMenu.Items.Remove(this.tmiUserApp);
                }
            }

            // recent file
            if(QTUtility.flagManager_RecentFile.GetFlag(this.lstRecentFileItems) ||
                    this.lstRefreshRequired[ITEMINDEX_RECENTFILE]) {
                this.lstRefreshRequired[ITEMINDEX_RECENTFILE] = false;

                // clear items
                foreach(ToolStripItem item in this.lstRecentFileItems) {
                    item.Dispose();
                }
                this.lstRecentFileItems.Clear();

                if(Config.Bool(Scts.Desktop_IncludeRecentFile)) {
                    lst[ITEMINDEX_RECENTFILE] = true;

                    this.lstRecentFileItems = MenuUtility.CreateRecentFilesItems(this.hwndShellTray);
                }
                else {
                    this.contextMenu.Items.Remove(this.tmiLabel_RecentFile);
                    this.contextMenu.Items.Remove(this.tmiRecentFile);
                }
            }

            return lst.ToArray();
        }

        private void ShowMenu(Point popUpPoint) {
            // Note:
            //		this method must be executed on Taskbar thread
            //		and set taskbar foreground beforehand.

            this.contextMenu.SuspendLayout();
            this.ddmrGroups.SuspendLayout();
            this.ddmrHistory.SuspendLayout();
            this.ddmrUserapps.SuspendLayout();
            this.ddmrRecentFile.SuspendLayout();

            // sync texts

            bool[] flags = this.BuildMenuItems();

            foreach(int index in this.lstItemOrder) {
                if(flags[index]) {
                    switch(index) {
                        case ITEMINDEX_GROUP:
                            this.AddMenuItems_Group();
                            break;

                        case ITEMINDEX_RECENTTAB:
                            this.AddMenuItems_RecentTab();
                            break;

                        case ITEMINDEX_APPLAUNCHER:
                            this.AddMenuItems_AppLauncher();
                            break;

                        case ITEMINDEX_RECENTFILE:
                            this.AddMenuItems_RecentFile();
                            break;
                    }
                }
            }

            this.ddmrUserapps.ResumeLayout();
            this.ddmrHistory.ResumeLayout();
            this.ddmrGroups.ResumeLayout();
            this.ddmrRecentFile.ResumeLayout();
            this.contextMenu.ResumeLayout();

            if(this.contextMenu.Items.Count > 0) {
                if(QTUtility.IsVista)
                    this.contextMenu.SendToBack();

                this.contextMenu.Show(popUpPoint);
            }
        }

        private void AddMenuItems_Group() {
            if(this.ExpandState[ITEMINDEX_GROUP]) {
                // Opened
                int index = GetInsertionIndex(ITEMINDEX_GROUP);
                this.contextMenu.InsertItem(index, this.tmiLabel_Group, MENUKEY_LABEL_GROUP);
                foreach(ToolStripItem item in this.lstGroupItems) {
                    this.contextMenu.InsertItem(++index, item, MENUKEY_ITEM_GROUP);
                }
            }
            else {
                // Closed
                this.ddmrGroups.AddItemsRange(this.lstGroupItems.ToArray(), MENUKEY_ITEM_GROUP);
                this.contextMenu.InsertItem(GetInsertionIndex(ITEMINDEX_GROUP), this.tmiGroup, MENUKEY_SUBMENUS);
            }
        }

        private void AddMenuItems_RecentTab() {
            if(this.ExpandState[ITEMINDEX_RECENTTAB]) {
                int index = GetInsertionIndex(ITEMINDEX_RECENTTAB);
                this.contextMenu.InsertItem(index, this.tmiLabel_History, MENUKEY_LABEL_HISTORY);
                foreach(ToolStripItem item in this.lstUndoClosedItems) {
                    this.contextMenu.InsertItem(++index, item, MENUKEY_ITEM_HISTORY);
                }
            }
            else {
                this.ddmrHistory.AddItemsRange(this.lstUndoClosedItems.ToArray(), MENUKEY_ITEM_HISTORY);
                this.contextMenu.InsertItem(GetInsertionIndex(ITEMINDEX_RECENTTAB), this.tmiHistory, MENUKEY_SUBMENUS);
            }
        }

        private void AddMenuItems_AppLauncher() {
            if(this.ExpandState[ITEMINDEX_APPLAUNCHER]) {
                int index = GetInsertionIndex(ITEMINDEX_APPLAUNCHER);
                this.contextMenu.InsertItem(index, this.tmiLabel_UserApp, MENUKEY_LABEL_USERAPP);
                foreach(ToolStripItem item in this.lstUserAppItems) {
                    this.contextMenu.InsertItem(++index, item, MENUKEY_ITEM_USERAPP);
                }
            }
            else {
                this.contextMenu.InsertItem(GetInsertionIndex(ITEMINDEX_APPLAUNCHER), this.tmiUserApp, MENUKEY_SUBMENUS);
            }
        }

        private void AddMenuItems_RecentFile() {
            if(this.ExpandState[ITEMINDEX_RECENTFILE]) {
                int index = GetInsertionIndex(ITEMINDEX_RECENTFILE);
                this.contextMenu.InsertItem(index, this.tmiLabel_RecentFile, MENUKEY_LABEL_RECENT);
                foreach(ToolStripItem item in this.lstRecentFileItems) {
                    this.contextMenu.InsertItem(++index, item, MENUKEY_ITEM_RECENT);
                }
            }
            else {
                this.ddmrRecentFile.AddItemsRange(this.lstRecentFileItems.ToArray(), MENUKEY_ITEM_RECENT);
                this.contextMenu.InsertItem(GetInsertionIndex(ITEMINDEX_RECENTFILE), this.tmiRecentFile,
                        MENUKEY_SUBMENUS);
            }
        }

        private int GetInsertionIndex(int ITEMINDEX) {
            int prev = -1;
            for(int i = 0; i < this.lstItemOrder.Count; i++) {
                if(this.lstItemOrder[i] == ITEMINDEX) {
                    if(i != 0) {
                        prev = this.lstItemOrder[i - 1];
                    }
                    break;
                }
            }

            if(prev == -1) {
                return 0;
            }
            else {
                for(int i = 0; i < this.contextMenu.Items.Count; i++) {
                    TitleMenuItem titleItem = this.contextMenu.Items[i] as TitleMenuItem;
                    if(titleItem != null) {
                        if(GenreToInt32(titleItem.Genre) == prev) {
                            // previous item found
                            for(int j = i + 1; j < this.contextMenu.Items.Count; j++) {
                                if(this.contextMenu.Items[j] is TitleMenuItem) {
                                    return j;
                                }
                            }

                            return this.contextMenu.Items.Count;
                        }
                    }
                }

                // previsous items not found...
                return GetInsertionIndex(prev);
            }
        }

        private static int GenreToInt32(MenuGenre genre) {
            switch(genre) {
                default:
                case MenuGenre.Group:
                    return ITEMINDEX_GROUP;

                case MenuGenre.RecentlyClosedTab:
                    return ITEMINDEX_RECENTTAB;

                case MenuGenre.Application:
                    return ITEMINDEX_APPLAUNCHER;

                case MenuGenre.RecentFile:
                    return ITEMINDEX_RECENTFILE;
            }
        }



        private void OnLabelTitleClickedToClose(MenuGenre genre) {
            int labelIndex = GenreToInt32(genre);

            this.ExpandState[labelIndex] = !this.ExpandState[labelIndex];
            this.fCancelClosing = true;

            ToolStripMenuItem labelToRemove;
            List<ToolStripItem> listToRemove;
            ToolStripMenuItem itemToAdd;
            string key;

            if(labelIndex == 0) {
                labelToRemove = this.tmiLabel_Group;
                listToRemove = this.lstGroupItems;
                itemToAdd = this.tmiGroup;
                key = MENUKEY_ITEM_GROUP;
            }
            else if(labelIndex == 1) {
                labelToRemove = this.tmiLabel_History;
                listToRemove = this.lstUndoClosedItems;
                itemToAdd = this.tmiHistory;
                key = MENUKEY_ITEM_HISTORY;
            }
            else if(labelIndex == 2) {
                labelToRemove = this.tmiLabel_UserApp;
                listToRemove = this.lstUserAppItems;
                itemToAdd = this.tmiUserApp;
                key = MENUKEY_ITEM_USERAPP;
            }
            else //if( labelIndex == 3 )
            {
                labelToRemove = this.tmiLabel_RecentFile;
                listToRemove = this.lstRecentFileItems;
                itemToAdd = this.tmiRecentFile;
                key = MENUKEY_ITEM_RECENT;
            }

            int index = this.contextMenu.Items.IndexOf(labelToRemove);

            this.contextMenu.SuspendLayout();

            this.contextMenu.Items.Remove(labelToRemove);
            foreach(ToolStripItem tsi in listToRemove) {
                this.contextMenu.Items.Remove(tsi);
            }

            this.contextMenu.InsertItem(index, itemToAdd, MENUKEY_SUBMENUS);

            ((DropDownMenuReorderable)itemToAdd.DropDown).AddItemsRange(listToRemove.ToArray(), key);


            this.contextMenu.ResumeLayout();

        }

        private void OnSubMenuTitleClickedToOpen(MenuGenre genre) {
            int menuIndex = 0;
            switch(genre) {
                case MenuGenre.Group:
                    menuIndex = 0;
                    break;
                case MenuGenre.RecentlyClosedTab:
                    menuIndex = 1;
                    break;
                case MenuGenre.Application:
                    menuIndex = 2;
                    break;
                case MenuGenre.RecentFile:
                    menuIndex = 3;
                    break;
            }

            this.ExpandState[menuIndex] = !this.ExpandState[menuIndex];
            this.fCancelClosing = true;

            ToolStripMenuItem itemToRemove;
            ToolStripMenuItem labelToAdd;
            List<ToolStripItem> listToAdd;
            string key;

            if(menuIndex == 0) {
                itemToRemove = this.tmiGroup;
                labelToAdd = this.tmiLabel_Group;
                listToAdd = this.lstGroupItems;
                key = MENUKEY_ITEM_GROUP;
            }
            else if(menuIndex == 1) {
                itemToRemove = this.tmiHistory;
                labelToAdd = this.tmiLabel_History;
                listToAdd = this.lstUndoClosedItems;
                key = MENUKEY_ITEM_HISTORY;
            }
            else if(menuIndex == 2) {
                itemToRemove = this.tmiUserApp;
                labelToAdd = this.tmiLabel_UserApp;
                listToAdd = this.lstUserAppItems;
                key = MENUKEY_ITEM_USERAPP;
            }
            else //if( menuIndex == 3 )
            {
                itemToRemove = this.tmiRecentFile;
                labelToAdd = this.tmiLabel_RecentFile;
                listToAdd = this.lstRecentFileItems;
                key = MENUKEY_ITEM_RECENT;
            }

            itemToRemove.DropDown.Hide();

            this.contextMenu.SuspendLayout();

            int index = this.contextMenu.Items.IndexOf(itemToRemove);
            this.contextMenu.Items.Remove(itemToRemove);

            this.contextMenu.InsertItem(index, labelToAdd, MENUKEY_LABELS + menuIndex);

            foreach(ToolStripItem tsi in listToAdd) {
                this.contextMenu.InsertItem(++index, tsi, key);
            }

            this.contextMenu.ResumeLayout();

        }


        private SubDirTipForm CreateSubDirTip() {
            // creates menu drop down for real folder contained application menu drop down
            // runs on TaskBar thread

            if(this.subDirTip_TB == null) {
                this.subDirTip_TB = new SubDirTipForm(this.Handle, this.hwndShellTray, false);
                this.subDirTip_TB.MenuItemClicked += new ToolStripItemClickedEventHandler(subDirTip_MenuItemClicked);
                this.subDirTip_TB.MultipleMenuItemsClicked += new EventHandler(subDirTip_MultipleMenuItemsClicked);
                this.subDirTip_TB.MenuItemRightClicked +=
                        new ItemRightClickedEventHandler(subDirTip_MenuItemRightClicked);
                this.subDirTip_TB.MultipleMenuItemsRightClicked +=
                        new ItemRightClickedEventHandler(subDirTip_MultipleMenuItemsRightClicked);
            }

            return this.subDirTip_TB;
        }



        #endregion


        #region ---------- Menu actions ----------


        private void OpenGroup(object obj) {
            object[] arr = (object[])obj;
            string[] groups = (string[])arr[0];
            Keys modKeys = (Keys)arr[1];

            for(int i = 0; i < groups.Length; i++) {
                GroupItemInfo gi;
                if(QTUtility.dicGroups.TryGetValue(groups[i], out gi) && gi.ItemCount > 0) {
                    using(IDLWrapper idlw = new IDLWrapper(gi.IDLS[0])) {
                        IntPtr hwndTabBar;
                        bool fOpened;

                        if(WindowUtils.GetTargetWindow(idlw, true, out hwndTabBar, out fOpened)) {
                            List<string> lstTmp = new List<string>();
                            for(int j = i; j < groups.Length; j++) {
                                lstTmp.Add(groups[j]);
                            }

                            if(lstTmp.Count > 0) {
                                if(fOpened) {
                                    QTUtility2.SendCOPYDATASTRUCT_IDL(hwndTabBar, (IntPtr)MC.TAB_GROUPSWINDOW,
                                            QTUtility2.ArrayToByte(lstTmp.ToArray()), IntPtr.Zero);
                                }
                                else {
                                    bool fForceNewWindow = (modKeys == Keys.Control);
                                    IntPtr wParam = (IntPtr)(fForceNewWindow ? MC.TAB_GROUPSNEWWINDOW : MC.TAB_GROUPS);

                                    QTUtility2.SendCOPYDATASTRUCT_IDL(hwndTabBar, wParam,
                                            QTUtility2.ArrayToByte(lstTmp.ToArray()), IntPtr.Zero);
                                }
                            }
                            return;
                        }
                    }
                }
            }
        }

        private void OpenTab(object obj) {
            object[] arr = (object[])obj;

            string path = (string)arr[0];
            Keys modifierKey = (Keys)arr[1];
            byte[] idl = null;
            if(arr.Length == 3)
                idl = (byte[])arr[2];

            IntPtr hwndTabBar;
            bool fOpened;

            IDLWrapper idlw;
            if(idl != null) {
                idlw = new IDLWrapper(idl, true);
            }
            else {
                idlw = new IDLWrapper(path); //, true
            }

            using(idlw) {
                if(WindowUtils.GetTargetWindow(idlw, false, out hwndTabBar, out fOpened)) {
                    if(!fOpened) {
                        QTUtility2.SendCOPYDATASTRUCT_IDL(hwndTabBar, (IntPtr)MC.TAB_OPEN_IDL, idlw.IDL,
                                (IntPtr)modifierKey);
                    }
                }
            }
        }

        private void OpenFolders2(object obj) {
            object[] arr = (object[])obj;

            List<byte[]> lstIDLs = (List<byte[]>)arr[0];
            int iKeys = (int)arr[1];

            if(lstIDLs.Count == 0)
                return;

            // on Vista, IShellBrowser.BrowseObject opens window in shell process
            // even if 'open in new process' option is on...

            QTUtility.InitializeOpeningsWith(null, null, null, lstIDLs);

            IntPtr hwndTabBar;
            bool fOpened;
            using(IDLWrapper idlw = new IDLWrapper(lstIDLs[0])) {
                if(WindowUtils.GetTargetWindow(idlw, false, out hwndTabBar, out fOpened)) {
                    if(!fOpened) {
                        // if existing window belongs to other process, 
                        // sychronize QTUtility.openingIDLs of the instance in the process
                        int procId, procIDCurrent = PInvoke.GetCurrentProcessId();
                        PInvoke.GetWindowThreadProcessId(hwndTabBar, out procId);
                        if(procId != procIDCurrent) {
                            // MessageListenerWindow of the Shell process surely exists - instantialized by Desktop Tool
                            //PInvoke.SendMessage( WindowUtils.GetMessageListenerWindow(), MC.MLW_QUERY_OPENINGIDLS, hwndTabBar, IntPtr.Zero );

                            if(QTUtility.openingIDLs.Count > 0) {
                                QTUtility2.SendCOPYDATASTRUCT_BINARY<byte[]>(hwndTabBar, (IntPtr)MC.TAB_SET_OPENINGIDLS,
                                        QTUtility.openingIDLs.ToArray(), IntPtr.Zero);
                                QTUtility.openingIDLs.Clear();
                            }
                        }

                        QTUtility2.SendCOPYDATASTRUCT(hwndTabBar, (IntPtr)MC.TAB_OPEN_IDLS, null, (IntPtr)iKeys); //0x09
                    }
                }
            }
        }

        private void DoFileTools(int index) {
            // desktop thread

            // 0	copy path
            // 1	copy name
            // 2	copy path current
            // 3	copy name current
            // 4	file hash
            // 5	show SubDirTip for selected folder
            // 6	copy file hash

            try {
                if(index == 2 || index == 3) {
                    // Send desktop path/name to Clipboard.
                    string str = String.Empty;
                    if(index == 2) {
                        str = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
                    }
                    else {
                        byte[] idl = new byte[] {0, 0};
                        using(IDLWrapper idlw = new IDLWrapper(idl)) {
                            if(idlw.Available) {
                                str = idlw.DisplayName;
                            }
                        }
                    }

                    if(str.Length > 0) {
                        QTUtility2.SetStringClipboard(str);
                    }
                    return;
                }


                int iItem;
                List<IntPtr> lstPIDL = this.GetSelectedItemPIDL(out iItem);

                // File Hash
                if(index == 4 || index == 6) {
                    List<string> lstPaths = new List<string>();
                    foreach(IntPtr pIDL in lstPIDL) {
                        using(IDLWrapper idlw = new IDLWrapper(pIDL)) {
                            if(idlw.IsLink) {
                                string pathLinkTarget = ShellMethods.GetLinkTargetPath(idlw.Path);
                                if(File.Exists(pathLinkTarget)) {
                                    lstPaths.Add(pathLinkTarget);
                                }
                            }
                            else if(idlw.IsFileSystemFile) {
                                lstPaths.Add(idlw.Path);
                            }
                        }
                    }

                    if(index == 4) {
                        FileHashComputer.ShowForm(lstPaths.ToArray());
                    }
                    else {
                        FileHashComputer.GetForPath(lstPaths, this.hwndListView);
                    }
                    return;
                }


                // Show subdirtip.
                if(index == 5) {
                    if(lstPIDL.Count == 1) {
                        if(this.ShowSubDirTip(lstPIDL[0], iItem, false)) {
                            this.subDirTip.PerformClickByKey();
                        }
                        else {
                            this.HideSubDirTip();
                        }
                    }

                    foreach(IntPtr pIDL in lstPIDL) {
                        PInvoke.CoTaskMemFree(pIDL);
                    }
                    return;
                }

                if(index == 0 || index == 1) {
                    bool fPath = index == 0;
                    string str = String.Empty;

                    foreach(IntPtr pIDL in lstPIDL) {
                        string path = ShellMethods.GetDisplayName(pIDL, !fPath);

                        if(path.Length > 0) {
                            str += (str.Length == 0 ? String.Empty : "\r\n") + path;
                        }

                        PInvoke.CoTaskMemFree(pIDL);
                    }

                    if(str.Length > 0) {
                        QTUtility2.SetStringClipboard(str);
                    }
                }
            }
            catch(Exception ex) {
                DebugUtil.AppendToExceptionLog(ex, null);
            }
        }


        #endregion


        #region ---------- Inner Classes ----------

        private sealed class TitleMenuItem : ToolStripMenuItem {
            private MenuGenre genre;
            private bool fOpened;

            private Bitmap bmpArrow_Cls, bmpArrow_Opn;

            private static bool drawBackground;
            private static StringFormat sf;
            private static Bitmap bmpTitle;


            public TitleMenuItem(MenuGenre genre, bool fOpened) {
                this.genre = genre;
                this.fOpened = fOpened;

                this.bmpArrow_Opn = Resources_Image.menuOpen;
                this.bmpArrow_Cls = Resources_Image.menuClose;

                if(TitleMenuItem.sf == null)
                    TitleMenuItem.Init();
            }

            private static void Init() {
                TitleMenuItem.sf = new StringFormat();
                TitleMenuItem.sf.Alignment = StringAlignment.Near;
                TitleMenuItem.sf.LineAlignment = StringAlignment.Center;

                TitleMenuItem.bmpTitle = Resources_Image.TitleBar;
            }

            protected override void Dispose(bool disposing) {
                if(this.bmpArrow_Opn != null) {
                    this.bmpArrow_Opn.Dispose();
                    this.bmpArrow_Opn = null;
                }
                if(this.bmpArrow_Cls != null) {
                    this.bmpArrow_Cls.Dispose();
                    this.bmpArrow_Cls = null;
                }

                base.Dispose(disposing);
            }

            protected override void OnPaint(PaintEventArgs e) {
                if(TitleMenuItem.drawBackground) {
                    Rectangle rct = new Rectangle(1, 0, this.Bounds.Width, this.Bounds.Height);

                    // draw background	100x24
                    e.Graphics.DrawImage(TitleMenuItem.bmpTitle,
                            new Rectangle(new Point(1, 0), new Size(1, this.Bounds.Height)),
                            new Rectangle(Point.Empty, new Size(1, 24)), GraphicsUnit.Pixel);
                    e.Graphics.DrawImage(TitleMenuItem.bmpTitle,
                            new Rectangle(new Point(2, 0), new Size(this.Bounds.Width - 3, 1)),
                            new Rectangle(new Point(1, 0), new Size(98, 1)), GraphicsUnit.Pixel);
                    e.Graphics.DrawImage(TitleMenuItem.bmpTitle,
                            new Rectangle(new Point(this.Bounds.Width - 1, 0), new Size(1, this.Bounds.Height)),
                            new Rectangle(new Point(99, 0), new Size(1, 24)), GraphicsUnit.Pixel);
                    e.Graphics.DrawImage(TitleMenuItem.bmpTitle,
                            new Rectangle(new Point(2, this.Bounds.Height - 1), new Size(this.Bounds.Width - 3, 1)),
                            new Rectangle(new Point(1, 23), new Size(98, 1)), GraphicsUnit.Pixel);
                    e.Graphics.DrawImage(TitleMenuItem.bmpTitle,
                            new Rectangle(new Point(2, 1), new Size(this.Bounds.Width - 3, this.Bounds.Height - 2)),
                            new Rectangle(new Point(1, 1), new Size(98, 22)), GraphicsUnit.Pixel);

                    // draw overwrite highlight
                    if(this.Selected) {
                        SolidBrush sb = new SolidBrush(Color.FromArgb(96, SystemColors.Highlight));
                        e.Graphics.FillRectangle(sb, rct);
                        sb.Dispose();
                    }

                    // draw arrow
                    if(this.HasDropDownItems) {
                        int y = (rct.Height - 16)/2;
                        if(y < 0)
                            y = 5;
                        else
                            y += 5;

                        using(SolidBrush sb = new SolidBrush(Color.FromArgb(this.Selected ? 255 : 128, Color.White))) {
                            Point p = new Point(rct.Width - 15, y);
                            Point[] ps = {p, new Point(p.X, p.Y + 8), new Point(p.X + 4, p.Y + 4)};
                            e.Graphics.FillPolygon(sb, ps);
                        }
                    }

                    // draw string
                    e.Graphics.DrawString(this.Text, this.Font, Brushes.White,
                            new RectangleF(34, 2, rct.Width - 34, rct.Height - 2), TitleMenuItem.sf);
                }
                else
                    base.OnPaint(e);

                // draw image		//Resources_Image.menuOpen : Resources_Image.menuClose,
                e.Graphics.DrawImage(this.fOpened ? this.bmpArrow_Cls : this.bmpArrow_Opn, new Rectangle(5, 4, 16, 16));
            }

            protected override Point DropDownLocation {
                get {
                    // show dropdown in the screen where parent dropdown is contained.
                    // multi monitor support...

                    Point pnt = base.DropDownLocation;

                    if(pnt != Point.Empty && this.HasDropDownItems) {
                        ToolStrip tsOwner = this.Owner;
                                // this ToolStrip must be ToolStripDropDown. not MenuStrip or something.
                        if(tsOwner != null && !Screen.FromPoint(tsOwner.Bounds.Location).Bounds.Contains(pnt)) {
                            pnt.X = tsOwner.Bounds.X - this.DropDown.Width;
                        }
                    }

                    return pnt;
                }
            }

            public MenuGenre Genre {
                get { return genre; }
            }

            public static bool DrawBackground {
                set { TitleMenuItem.drawBackground = value; }
            }

            public bool IsOpened {
                get { return this.fOpened; }
            }
        }

        #endregion


        #region ---------- IDeskBand2 members ----------

        public int CanRenderComposited(out bool pfCanRenderComposited) {
            pfCanRenderComposited = true;
            return S_OK;
        }

        public int SetCompositionState(bool fCompositionEnabled) {
            return S_OK;
        }

        public int GetCompositionState(out bool pfCompositionEnabled) {
            pfCompositionEnabled = true;
            return S_OK;
        }

        #endregion


        #region ---------- Register / Unregister ----------


        [ComRegisterFunction]
        private static void Register(Type t) {
            string guid = t.GUID.ToString("B");

            bool fJa = System.Globalization.CultureInfo.CurrentCulture.Parent.Name == "ja";
            string strDesktopTool = fJa ? "QT Tab ƒfƒXƒNƒgƒbƒv ƒc[ƒ‹" : "QT Tab Desktop Tool";

            // CLSID
            using(RegistryKey rkClass = Registry.ClassesRoot.CreateSubKey(@"CLSID\" + guid)) {
                rkClass.SetValue(null, strDesktopTool);
                rkClass.SetValue("MenuText", strDesktopTool);
                rkClass.SetValue("HelpText", strDesktopTool);
                rkClass.CreateSubKey(@"Implemented Categories\{00021492-0000-0000-C000-000000000046}");
            }

            // delete old class name
            Registry.ClassesRoot.DeleteSubKeyTree("QTTabBarLib.QTCoTaskBarClass", false);
        }

        [ComUnregisterFunction]
        private static void Unregister(Type t) {
            // {D2BF470E-ED1C-487F-A555-2BD8835EB6CE}
            string guid = t.GUID.ToString("B");

            try {
                using(RegistryKey rkClass = Registry.ClassesRoot.CreateSubKey(@"CLSID")) {
                    rkClass.DeleteSubKeyTree(guid);
                }
            }
            catch {
            }
        }

        #endregion

    }
}
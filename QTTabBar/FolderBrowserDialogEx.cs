using System;
using System.Windows.Forms;
using QTTabBarLib.Interop;

namespace QTTabBarLib
{
    public class FolderBrowserDialogEx : CommonDialog
    {
        public FolderBrowserDialogEx()
        {
            ShowNewFolderButton = true;
        }

        public string Description { get; set; }
        public Environment.SpecialFolder RootFolder { get; set; }
        public string SelectedPath { get; set; }
        public bool ShowNewFolderButton { get; set; }

        public override void Reset()
        {
            throw new NotImplementedException();
        }

        protected override bool RunDialog(IntPtr hwndOwner)
        {
            string root = Environment.GetFolderPath(RootFolder);
            using (IDLWrapper wrapper = new IDLWrapper(root)) {
                BROWSEINFO bi = new BROWSEINFO();
                bi.hwndOwner = hwndOwner;
                bi.pidlRoot = wrapper.PIDL;
                bi.pszDisplayName = SelectedPath;
                bi.lpszTitle = Description;
                bi.ulFlags = BROWSEINFO.BIF_NEWDIALOGSTYLE | BROWSEINFO.BIF_SHAREABLE | BROWSEINFO.BIF_EDITBOX;
                if (!ShowNewFolderButton) {
                    bi.ulFlags |= BROWSEINFO.BIF_NONEWFOLDERBUTTON;
                }

                using (IDLWrapper wrapper2 = new IDLWrapper(PInvoke.SHBrowseForFolder(ref bi))) {
                    SelectedPath = wrapper2.Path;
                    if (String.IsNullOrEmpty(SelectedPath)) {
                        return false;
                    }
                    return true;
                }
            }
        }
    }
}

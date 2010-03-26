// Generated by Reflector from c:\program files\qttabbar\QTPluginLib.dll
namespace QTPlugin.Interop
{
  using System;
  using System.Runtime.InteropServices;
  
  [ComImport, InterfaceType(ComInterfaceType.InterfaceIsIUnknown), Guid("000214E6-0000-0000-C000-000000000046")]
  public interface IShellFolder
  {
    [PreserveSig]
    int ParseDisplayName(IntPtr hwnd, IntPtr pbc, [MarshalAs(UnmanagedType.LPWStr)] string pszDisplayName, ref uint pchEaten, out IntPtr ppidl, ref uint pdwAttributes);
    [PreserveSig]
    int EnumObjects(IntPtr hwnd, int grfFlags, out IEnumIDList ppenumIDList);
    [PreserveSig]
    int BindToObject(IntPtr pidl, IntPtr pbc, [In] ref Guid riid, [MarshalAs(UnmanagedType.Interface)] out object ppv);
    [PreserveSig]
    int BindToStorage(IntPtr pidl, IntPtr pbc, [In] ref Guid riid, [MarshalAs(UnmanagedType.Interface)] out object ppv);
    [PreserveSig]
    int CompareIDs(IntPtr lParam, IntPtr pidl1, IntPtr pidl2);
    [PreserveSig]
    int CreateViewObject(IntPtr hwndOwner, [In] ref Guid riid, [MarshalAs(UnmanagedType.Interface)] out object ppv);
    [PreserveSig]
    int GetAttributesOf(uint cidl, [MarshalAs(UnmanagedType.LPArray, SizeParamIndex=0, SizeConst=0)] IntPtr[] apidl, ref uint rgfInOut);
    [PreserveSig]
    int GetUIObjectOf(IntPtr hwndOwner, uint cidl, [MarshalAs(UnmanagedType.LPArray, SizeParamIndex=1, SizeConst=0)] IntPtr[] apidl, [In] ref Guid riid, ref uint rgfReserved, [MarshalAs(UnmanagedType.Interface)] out object ppv);
    [PreserveSig]
    int GetDisplayNameOf(IntPtr pidl, uint uFlags, out STRRET pName);
    [PreserveSig]
    int SetNameOf(IntPtr hwndOwner, IntPtr pidl, [MarshalAs(UnmanagedType.LPWStr)] string pszName, uint uFlags, out IntPtr ppidlOut);
  }
}

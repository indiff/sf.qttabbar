// Generated by Reflector from c:\program files\qttabbar\QTPluginLib.dll
namespace QTPlugin
{
  using QTPlugin.Interop;
  using System;
  using System.Text.RegularExpressions;
  
  public interface IFilterCore : IPluginClient
  {
    bool IsMatch(IShellFolder shellFolder, IntPtr pIDLChild, Regex reQurey);
  }
}

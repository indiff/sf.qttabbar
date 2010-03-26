// Generated by Reflector from c:\program files\qttabbar\QTPluginLib.dll
namespace QTPlugin
{
  using System;
  
  public interface ITab
  {
    bool Browse(QTPlugin.Address address);
    bool Browse(bool fBack);
    void Clone(int index, bool fSelect);
    bool Close();
    QTPlugin.Address[] GetBraches();
    QTPlugin.Address[] GetHistory(bool fBack);
    bool Insert(int index);
    
    QTPlugin.Address Address { get; }
    
    int Index { get; }
    
    bool Locked { get; set; }
    
    bool Selected { get; set; }
    
    string SubText { get; set; }
    
    string Text { get; set; }
  }
}

// Generated by Reflector from c:\program files\qttabbar\QTTabBar.dll
namespace QTTabBarLib
{
  using System;
  
  internal sealed class EventPack
  {
    public EventHandler DirDoubleClickEventHandler;
    public bool FromTaskBar;
    public ItemRightClickedEventHandler ItemRightClickEventHandler;
    public IntPtr MessageParentHandle;
    
    public EventPack(IntPtr hwnd, ItemRightClickedEventHandler handlerRightClick, EventHandler handlerDirDblClick, bool fFromTaskBar)
    {
      this.MessageParentHandle = hwnd;
      this.ItemRightClickEventHandler = handlerRightClick;
      this.DirDoubleClickEventHandler = handlerDirDblClick;
      this.FromTaskBar = fFromTaskBar;
    }
  }
}

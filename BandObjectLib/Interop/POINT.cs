// Generated by Reflector from c:\program files\qttabbar\QTTabBar.dll
namespace BandObjectLib
{
  using System;
  using System.Drawing;
  using System.Runtime.InteropServices;
  
  [StructLayout(LayoutKind.Sequential)]
  public struct POINT
  {
    public int x;
    public int y;
    public POINT(Point pnt)
    {
      this.x = pnt.X;
      this.y = pnt.Y;
    }
  }
}

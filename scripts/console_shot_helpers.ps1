function Quote-ConsoleLiteral([string]$Value){return "'"+$Value.Replace("'","''")+"'"}

function Resolve-ConsoleShell([string]$Name){
 foreach($candidate in $(if($Name -eq 'Auto'){@('pwsh','powershell')}else{@($Name)})){
  $cmd=Get-Command ($candidate+'.exe') -CommandType Application -ErrorAction SilentlyContinue|Select-Object -First 1
  if($cmd){return $cmd.Source}
  $path=if($candidate -eq 'pwsh'){Join-Path $env:ProgramFiles 'PowerShell\7\pwsh.exe'}else{Join-Path $env:SystemRoot 'System32\WindowsPowerShell\v1.0\powershell.exe'}
  if(Test-Path -LiteralPath $path){return $path}
 }
 throw "Requested shell '$Name' was not found. Install PowerShell 7 or explicitly select -Shell powershell."
}

function Initialize-ConsoleCapture {
 if('WordReportCapture.Native' -as [type]){return}
 Add-Type @"
using System;
using System.Text;
using System.Collections.Generic;
using System.Runtime.InteropServices;
namespace WordReportCapture {
 public static class Native {
  public delegate bool EnumProc(IntPtr h, IntPtr p);
  [DllImport("user32.dll")] public static extern bool EnumWindows(EnumProc cb, IntPtr p);
  [DllImport("user32.dll",CharSet=CharSet.Unicode)] public static extern int GetWindowText(IntPtr h,StringBuilder s,int n);
  [DllImport("user32.dll")] public static extern bool IsWindowVisible(IntPtr h);
  [DllImport("user32.dll")] public static extern IntPtr GetForegroundWindow();
  [DllImport("user32.dll")] public static extern bool SetForegroundWindow(IntPtr h);
  [DllImport("user32.dll")] public static extern bool ShowWindow(IntPtr h,int cmd);
  [DllImport("user32.dll",SetLastError=true)] public static extern bool MoveWindow(IntPtr h,int x,int y,int w,int ht,bool repaint);
  [DllImport("user32.dll")] public static extern bool PostMessage(IntPtr h,uint msg,IntPtr w,IntPtr l);
  [DllImport("user32.dll",SetLastError=true)] public static extern IntPtr SetThreadDpiAwarenessContext(IntPtr ctx);
  [DllImport("user32.dll",SetLastError=true)] static extern bool GetClientRect(IntPtr h,out RECT r);
  [DllImport("user32.dll",SetLastError=true)] static extern bool ClientToScreen(IntPtr h,ref POINT p);
  [DllImport("user32.dll")] public static extern uint GetDpiForWindow(IntPtr h);
  [DllImport("user32.dll")] static extern IntPtr WindowFromPoint(POINT p);
  [DllImport("user32.dll")] static extern IntPtr GetAncestor(IntPtr h,uint flags);
  [DllImport("user32.dll")] static extern int GetSystemMetrics(int index);
  [StructLayout(LayoutKind.Sequential)] struct RECT {public int Left,Top,Right,Bottom;}
  [StructLayout(LayoutKind.Sequential)] struct POINT {public int X,Y;}
  public static bool Owns(IntPtr h,string token){
   if(h==IntPtr.Zero || !IsWindowVisible(h)) return false;
   var text=new StringBuilder(1024);GetWindowText(h,text,text.Capacity);
   return text.ToString().IndexOf(token,StringComparison.Ordinal)>=0;
  }
  public static IntPtr Find(string token){
   IntPtr found=IntPtr.Zero;
   EnumWindows(delegate(IntPtr h,IntPtr p){if(Owns(h,token)){found=h;return false;}return true;},IntPtr.Zero);
   return found;
  }
  public static int[] ClientRect(IntPtr h){
   RECT r;POINT p=new POINT();
   if(!GetClientRect(h,out r) || !ClientToScreen(h,ref p)) throw new InvalidOperationException("Cannot resolve client coordinates.");
   return new int[]{p.X,p.Y,r.Right-r.Left,r.Bottom-r.Top};
  }
  public static bool Unobscured(IntPtr h,int[] r){
   int left=GetSystemMetrics(76),top=GetSystemMetrics(77);
   if(r[2]<=0 || r[3]<=0 || r[0]<left || r[1]<top || r[0]+r[2]>left+GetSystemMetrics(78) || r[1]+r[3]>top+GetSystemMetrics(79)) return false;
   for(int y=8;y<r[3]-8;y+=16) for(int x=8;x<r[2]-8;x+=16){
    POINT p=new POINT();p.X=r[0]+x;p.Y=r[1]+y;
    if(GetAncestor(WindowFromPoint(p),2)!=h) return false;
   }
   return true;
  }
  // Inspect the body, not five arbitrary background points or the tab strip.
  public static bool HasContent(byte[] pixels,int width,int height,int stride,int skipTop){
   int start=Math.Max(0,skipTop),bottom=height-6;
   if(width<20 || start>=bottom || stride<width*4 || pixels.Length<stride*height) return false;
   var counts=new Dictionary<int,int>();int background=0,best=0;
   for(int y=start;y<bottom;y+=4) for(int x=6;x<width-6;x+=4){
    int i=y*stride+x*4;int c=(pixels[i]<<16)|(pixels[i+1]<<8)|pixels[i+2];int n;
    counts.TryGetValue(c,out n);counts[c]=++n;if(n>best){best=n;background=c;}
   }
   int hits=0;
   for(int y=start;y<bottom;y++) for(int x=6;x<width-6;x++){
    int i=y*stride+x*4;
    int delta=Math.Abs(pixels[i]-((background>>16)&255))+Math.Abs(pixels[i+1]-((background>>8)&255))+Math.Abs(pixels[i+2]-(background&255));
    if(delta>60 && ++hits>=16) return true;
   }
   return false;
  }
 }
}
"@
}

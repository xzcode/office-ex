# SPDX-License-Identifier: MIT
# Build-time resource editing of our own staged, unsigned preview executable.
param(
  [Parameter(Mandatory=$true)][string]$Executable,
  [Parameter(Mandatory=$true)][string]$Icon
)
$ErrorActionPreference = 'Stop'
if ($env:OS -ne 'Windows_NT') { throw 'Native Windows build required' }
$exePath = (Resolve-Path -LiteralPath $Executable).Path
$icoPath = (Resolve-Path -LiteralPath $Icon).Path
if ([IO.Path]::GetFileName($exePath) -ne 'Office EX.exe') { throw 'Expected staged Office EX executable' }
Add-Type -TypeDefinition @'
using System;
using System.IO;
using System.Collections.Generic;
using System.Runtime.InteropServices;
public static class OfficeExIcon {
 [UnmanagedFunctionPointer(CallingConvention.Winapi)] delegate bool NameCallback(IntPtr module,IntPtr type,IntPtr name,IntPtr param);
 [UnmanagedFunctionPointer(CallingConvention.Winapi)] delegate bool LangCallback(IntPtr module,IntPtr type,IntPtr name,ushort lang,IntPtr param);
 [DllImport("kernel32.dll",CharSet=CharSet.Unicode,ExactSpelling=true,SetLastError=true)] static extern IntPtr LoadLibraryExW(string path,IntPtr file,uint flags);
 [DllImport("kernel32.dll",SetLastError=true)] static extern bool FreeLibrary(IntPtr module);
 [DllImport("kernel32.dll",CharSet=CharSet.Unicode,ExactSpelling=true,SetLastError=true)] static extern bool EnumResourceNamesW(IntPtr module,IntPtr type,NameCallback callback,IntPtr param);
 [DllImport("kernel32.dll",CharSet=CharSet.Unicode,ExactSpelling=true,SetLastError=true)] static extern bool EnumResourceLanguagesW(IntPtr module,IntPtr type,IntPtr name,LangCallback callback,IntPtr param);
 [DllImport("kernel32.dll",CharSet=CharSet.Unicode,ExactSpelling=true,SetLastError=true)] static extern IntPtr FindResourceExW(IntPtr module,IntPtr type,IntPtr name,ushort language);
 [DllImport("kernel32.dll",SetLastError=true)] static extern uint SizeofResource(IntPtr module,IntPtr resource);
 [DllImport("kernel32.dll",SetLastError=true)] static extern IntPtr LoadResource(IntPtr module,IntPtr resource);
 [DllImport("kernel32.dll",SetLastError=true)] static extern IntPtr LockResource(IntPtr resource);
 [DllImport("kernel32.dll",CharSet=CharSet.Unicode,ExactSpelling=true,SetLastError=true)] static extern IntPtr BeginUpdateResourceW(string file,bool deleteExisting);
 [DllImport("kernel32.dll",CharSet=CharSet.Unicode,ExactSpelling=true,SetLastError=true)] static extern bool UpdateResourceW(IntPtr handle,IntPtr type,IntPtr name,ushort lang,byte[] data,uint length);
 [DllImport("kernel32.dll",CharSet=CharSet.Unicode,ExactSpelling=true,SetLastError=true)] static extern bool EndUpdateResourceW(IntPtr handle,bool discard);
 static readonly IntPtr IconType=(IntPtr)3, GroupType=(IntPtr)14;
 sealed class Group { public int Id;public string Name;public ushort Lang; }
 static bool Numeric(IntPtr p) {return ((ulong)p.ToInt64()>>16)==0;}
 static Exception Error(string operation) {return new System.ComponentModel.Win32Exception(Marshal.GetLastWin32Error(),operation);}
 static IntPtr Open(string file) {var module=LoadLibraryExW(file,IntPtr.Zero,2|32);if(module==IntPtr.Zero)throw Error("Load PE resources");return module;}
 static List<Group> Groups(IntPtr module){
  var groups=new List<Group>();
  NameCallback cb=(m,t,n,p)=>{
   int id=Numeric(n)?n.ToInt32():0;string name=Numeric(n)?null:Marshal.PtrToStringUni(n);
   LangCallback lang=(mm,tt,nn,l,pp)=>{groups.Add(new Group{Id=id,Name=name,Lang=l});return true;};
   return EnumResourceLanguagesW(m,t,n,lang,IntPtr.Zero);
  };
  if(!EnumResourceNamesW(module,GroupType,cb,IntPtr.Zero)||groups.Count==0)throw Error("Enumerate icon groups");
  return groups;
 }
 static byte[] Read(IntPtr module,IntPtr type,IntPtr name,ushort lang){
  var r=FindResourceExW(module,type,name,lang);if(r==IntPtr.Zero)throw Error("Find updated resource");
  var size=SizeofResource(module,r);var data=LockResource(LoadResource(module,r));if(size==0||data==IntPtr.Zero)throw Error("Read updated resource");
  var bytes=new byte[checked((int)size)];Marshal.Copy(data,bytes,0,bytes.Length);return bytes;
 }
 static void Equal(byte[] a,byte[] b) {if(a.Length!=b.Length)throw new Exception("Resource length mismatch");for(int i=0;i<a.Length;i++)if(a[i]!=b[i])throw new Exception("Resource content mismatch");}
 static void Write(IntPtr handle,IntPtr type,IntPtr name,ushort lang,byte[] bytes){if(!UpdateResourceW(handle,type,name,lang,bytes,(uint)bytes.Length))throw Error("Update icon resource");}
 public static void Apply(string file,string icon){
  byte[] bytes=File.ReadAllBytes(icon);if(bytes.Length<6||BitConverter.ToUInt16(bytes,0)!=0||BitConverter.ToUInt16(bytes,2)!=1)throw new Exception("Invalid ICO");
  int count=BitConverter.ToUInt16(bytes,4);if(count<1||count>16||bytes.Length<6+16*count)throw new Exception("Invalid icon count");
  var images=new List<byte[]>();var ids=new List<int>();var occupied=new HashSet<int>();List<Group> groups;
  IntPtr module=Open(file);
  try{
   groups=Groups(module);
   NameCallback names=(m,t,n,p)=>{if(Numeric(n))occupied.Add(n.ToInt32());return true;};
   if(!EnumResourceNamesW(module,IconType,names,IntPtr.Zero))throw Error("Enumerate existing icons");
  }finally{FreeLibrary(module);}
  byte[] group=new byte[6+14*count];Array.Copy(bytes,group,6);int next=60000;
  for(int i=0;i<count;i++){
   int p=6+16*i,at=6+14*i,length=checked((int)BitConverter.ToUInt32(bytes,p+8)),offset=checked((int)BitConverter.ToUInt32(bytes,p+12));
   if(length<24||offset<6+16*count||offset>bytes.Length-length)throw new Exception("Invalid icon image bounds");
   byte[] image=new byte[length];Array.Copy(bytes,offset,image,0,length);images.Add(image);
   while(occupied.Contains(next))next++;if(next>65535)throw new Exception("Icon resource id exhaustion");ids.Add(next);occupied.Add(next);
   Array.Copy(bytes,p,group,at,12);Array.Copy(BitConverter.GetBytes((ushort)next),0,group,at+12,2);
  }
  var handle=BeginUpdateResourceW(file,false);if(handle==IntPtr.Zero)throw Error("Open staged executable for resource edit");
  bool committed=false;
  try{
   var languages=new HashSet<ushort>();
   foreach(var g in groups){
    if(languages.Add(g.Lang))for(int i=0;i<count;i++)Write(handle,IconType,(IntPtr)ids[i],g.Lang,images[i]);
    IntPtr name=g.Name==null?(IntPtr)g.Id:Marshal.StringToHGlobalUni(g.Name);
    try{Write(handle,GroupType,name,g.Lang,group);}finally{if(g.Name!=null)Marshal.FreeHGlobal(name);}
   }
   if(!EndUpdateResourceW(handle,false))throw Error("Commit icon resources");committed=true;
  }finally{if(!committed)EndUpdateResourceW(handle,true);}
  module=Open(file);
  try{
   foreach(var g in groups){
    IntPtr name=g.Name==null?(IntPtr)g.Id:Marshal.StringToHGlobalUni(g.Name);
    try{Equal(Read(module,GroupType,name,g.Lang),group);}finally{if(g.Name!=null)Marshal.FreeHGlobal(name);}
    for(int i=0;i<count;i++)Equal(Read(module,IconType,(IntPtr)ids[i],g.Lang),images[i]);
   }
  }finally{FreeLibrary(module);}
  Console.WriteLine("Verified Office EX PE icon resources: "+groups.Count+" groups, "+count+" resolutions");
 }
}
'@
[OfficeExIcon]::Apply($exePath,$icoPath)

/*----------------------------------------------------------------------------
	%%File: OLE32.CS
	%%Unit: OLE32
	%%Contact: rlittle

----------------------------------------------------------------------------*/

using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;
using System.Runtime.InteropServices;
using Microsoft.Win32;
using System.IO;

namespace UWin32
{
	public enum STGTY
	{
		STORAGE	= 1,
		STREAM	= 2,
		LOCKBYTES	= 3,
		PROPERTY	= 4
	};

	public class ole32
	{
	public const int STGM_DIRECT = 0x00000000;
	public const int STGM_TRANSACTED = 0x00010000;
	public const int STGM_SIMPLE = 0x08000000;
	
	public const int STGM_READ = 0x00000000;
	public const int STGM_WRITE = 0x00000001;
	public const int STGM_READWRITE = 0x00000002;
	
	public const int STGM_SHARE_DENY_NONE = 0x00000040;
	public const int STGM_SHARE_DENY_READ = 0x00000030;
	public const int STGM_SHARE_DENY_WRITE = 0x00000020;
	public const int STGM_SHARE_EXCLUSIVE = 0x00000010;
	
	public const int STGM_PRIORITY = 0x00040000;
	public const int STGM_DELETEONRELEASE = 0x04000000;
	public const int STGM_NOSCRATCH = 0x00100000;
	
	public const int STGM_CREATE = 0x00001000;
	public const int STGM_CONVERT = 0x00020000;
	public const int STGM_FAILIFTHERE = 0x00000000;
	
	[DllImport("shell32.dll", EntryPoint="ExtractIconEx", SetLastError=true, CharSet=CharSet.Unicode)] 
	public static extern int ExtractIconEx(
		[MarshalAs(UnmanagedType.LPWStr)] string szFile,
		int nIconIndex,
		out IntPtr hIconLarge,
		out IntPtr hIconSmall,
		int nIcons);

	[DllImport("user32.dll", EntryPoint="DestroyIcon", SetLastError=true, CharSet=CharSet.Unicode)] 
	public static extern int DestroyIcon(IntPtr hIcon);


	[InterfaceType(ComInterfaceType.InterfaceIsIUnknown),
	Guid("0000000d-0000-0000-C000-000000000046")]
	public interface UCOMIEnumSTATSTG
	{
		[PreserveSig]
		int Next(int celt, out STATSTG sstg, IntPtr celtOut);
	};

//	[InterfaceType(ComInterfaceType.InterfaceIsIUnknown),
//	Guid("
	[InterfaceType(ComInterfaceType.InterfaceIsIUnknown),
	Guid("0000000b-0000-0000-C000-000000000046")]
	public interface UCOMIStorage
	{
		void CreateStream(
			[MarshalAs(UnmanagedType.LPWStr)] string wcsName,
			int grfMode, //Access mode for the new stream
			int reserved1, //Reserved; must be zero
			int reserved2, //Reserved; must be zero
			out UCOMIStream stream //Pointer to output variable
			);

		UCOMIStream OpenStream(
			[MarshalAs(UnmanagedType.LPWStr)] string wcsName,
			IntPtr reserved1, //Reserved; must be NULL
			int grfMode, //Access mode for the new stream
			int reserved2 //Reserved; must be zero
//			out UCOMIStream stream //Pointer to output variable
			);
	
		void CreateStorage(
			[MarshalAs(UnmanagedType.LPWStr)] string wcsName,
			int grfMode,
			int reserved1,
			int reserved2,
			out UCOMIStorage istg);

		void OpenStorage(
			[MarshalAs(UnmanagedType.LPWStr)] string wcsName,
			IntPtr istgPriority,
			int grfMode,
			IntPtr snb,
			int reserved,
			out UCOMIStorage istg);

		void CopyTo(
			int ciidExclude,
			IntPtr rgiidExclude,
			IntPtr snb,
			UCOMIStorage istg);

		void MoveElementTo(
			[MarshalAs(UnmanagedType.LPWStr)] string wcsName,
			UCOMIStorage istg,
			[MarshalAs(UnmanagedType.LPWStr)] string wcsNewName,
			int grfFlags);

		void Commit(
			int grfCommitFlags);

		void Revert();

		UCOMIEnumSTATSTG EnumElements(
			Int32 reserved1, 
			IntPtr reserved2, 
			Int32 reserved3);
//			out UCOMIEnumSTATSTG iEnum);

		void DestroyElement(
			[MarshalAs(UnmanagedType.LPWStr)] string wcsName);

	}

	[DllImport("ole32.dll")]
	public static extern int StgOpenStorage (
		[MarshalAs(UnmanagedType.LPWStr)] string wcsName,
		IntPtr stgPriority,
		int grfMode,
		IntPtr snbExclude,
		int reserved,
		out UCOMIStorage stgOpen
		);

	[DllImport("ole32.dll", EntryPoint="StgOpenStorage")]
	public static extern int StgReOpenStorage (
		IntPtr wcsName,
		UCOMIStorage stgPriority,
		int grfMode,
		IntPtr snbExclude,
		int reserved,
		out UCOMIStorage stgOpen
		);

	[DllImport("ole32.dll", EntryPoint="StgCreateDocfile")]
	public static extern int StgCreateDocfile (
		[MarshalAs(UnmanagedType.LPWStr)] string wcsName,
		int grfMode,
		int reserved,
		out UCOMIStorage stgOpen
		);
	}
}


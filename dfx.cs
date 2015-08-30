using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;
using System.Runtime.InteropServices;
using Microsoft.Win32;
using System.IO;
using UWin32;
using UI;


namespace dfx
{
	/// <summary>
	/// Summary description for Form1.
	/// </summary>
	public class m_frm : System.Windows.Forms.Form
	{
		public struct mpExt
		{
			public string sExt;
			public int iIcon;
		};

		public class DIRE
		{
			public int iIcon;
			public ArrayList pldire;
//			public STATSTG sstg;
			public string sName;
			public STGTY type;
			public ole32.UCOMIStorage istg;
			public DIRE direParent;

			public DIRE(DIRE direParentIn)
			{
				pldire = new ArrayList();
				direParent = direParentIn;
			}
		};

		public class TL
		{
			public TL tlChild;
			public TL tlParent;
			public int data;
		}

		private ArrayList m_plmpExtensions;
		private ArrayList m_pldire;
		private DIRE m_direCurrent;
		private DIRE m_direRoot;

		private System.Windows.Forms.TextBox m_ebFilename;
		private System.Windows.Forms.ListView m_lvExplorer;
		private System.Windows.Forms.ToolBar m_tlbMain;
		private System.Windows.Forms.ToolBarButton m_tbbOpen;
		private System.Windows.Forms.ImageList m_ilButtons;
		private System.Windows.Forms.ToolBarButton m_tbbSaveStream;
		private System.Windows.Forms.ToolBarButton m_tbbUp;
		private System.Windows.Forms.ToolBarButton m_tbbNewStorage;
		private System.Windows.Forms.ToolBarButton m_tbbSep1;
		private System.Windows.Forms.ToolBarButton m_tbbSep2;
		private System.ComponentModel.IContainer components;
		private System.Windows.Forms.ToolBarButton m_tbbDelete;
		private string m_sFile;

		public const int itbbOpen = 0;
		public const int itbbNew = 2;
		public const int itbbSave = 3;
		public const int itbbDelete = 4;
		public const int itbbUp = 6;

		/* M  _ F R M */
		/*----------------------------------------------------------------------------
			%%Function: m_frm
			%%Qualified: dfx.m_frm.m_frm
			%%Contact: rlittle

		----------------------------------------------------------------------------*/
		public m_frm()
		{
			//
			// Required for Windows Form Designer support
			//
			InitializeComponent();

			//
			// TODO: Add any constructor code after InitializeComponent call
			//
			ContextMenu cxm;

			cxm = new ContextMenu();
			cxm.MenuItems.AddRange(new MenuItem[] {
													new MenuItem("List", new System.EventHandler(this.ListView), Shortcut.None),
													new MenuItem("Details", new System.EventHandler(this.DetailsView), Shortcut.None),
													new MenuItem("Icons", new System.EventHandler(this.IconView), Shortcut.None),
													new MenuItem("Add Storage", new System.EventHandler(this.AddStorage), Shortcut.None),
													new MenuItem("Delete", new System.EventHandler(this.DeleteDire), Shortcut.None)									
												  });
			m_lvExplorer.ContextMenu = cxm;

												  
			m_lvExplorer.Columns.Add(new ColumnHeader());
			m_lvExplorer.Columns[0].Text = "Name";
			
			m_lvExplorer.LargeImageList = new ImageList();
			m_lvExplorer.SmallImageList = new ImageList();

			m_lvExplorer.LargeImageList.ImageSize = new Size(32, 32);

			m_lvExplorer.LargeImageList.Images.Add(Bitmap.FromFile("d:\\dev\\dfx\\app.ico"));
			m_lvExplorer.SmallImageList.Images.Add(Bitmap.FromFile("d:\\dev\\dfx\\app.ico"));
			
			m_lvExplorer.LargeImageList.Images.Add(Bitmap.FromFile("d:\\dev\\dfx\\folderclosed.ico"));
			m_lvExplorer.SmallImageList.Images.Add(Bitmap.FromFile("d:\\dev\\dfx\\folderclosed.ico"));

			m_rgb = new byte[lcbMax];
//			IntPtr hIcon1, hIcon2;

//			int i = ole32.ExtractIconEx("c:\\app.ico", 0, out hIcon1, out hIcon2, 1);
//			m_lvExplorer.LargeImageList.Images.Add(Bitmap.FromHicon(hIcon1));
//			m_lvExplorer.SmallImageList.Images.Add(Bitmap.FromHicon(hIcon2));
//			if (hIcon1 != IntPtr.Zero)
//				ole32.DestroyIcon(hIcon1);
//
//			if (hIcon2 != IntPtr.Zero)
//				ole32.DestroyIcon(hIcon2);

//			int i = ole32.ExtractIconEx("C:\\WINDOWS\\Installer\\{90110409-6000-11D3-8CFE-0150048383C9}\\wordicon.exe", 1, out hIcon1, out hIcon2, 1);

			m_plmpExtensions = new ArrayList();
			m_pldire = new ArrayList();
			int i = IEnsureIcon(".doc");
			TL tl = new TL();
			TL tl2 = new TL();

			tl.data = 1;
			tl2.data = 2;
			tl.tlChild = tl2;
			tl2.tlParent = tl;
			tl.data = 3;
			tl2.data = 4;
		
			this.m_tbbOpen.Tag = itbbOpen;
			this.m_tbbSaveStream.Tag = itbbSave;
			this.m_tbbNewStorage.Tag = itbbNew;
			this.m_tbbDelete.Tag = itbbDelete;
			this.m_tbbUp.Tag = itbbUp;
		}

		/* I  E N S U R E  I C O N */
		/*----------------------------------------------------------------------------
			%%Function: IEnsureIcon
			%%Qualified: dfx.m_frm.IEnsureIcon
			%%Contact: rlittle

		----------------------------------------------------------------------------*/
		public int IEnsureIcon(string sExt)
		{
			int iIcon = 0;

			// first, check to see if its in our iconlist
			foreach (mpExt mp in m_plmpExtensions)
				{
				if (String.Compare(sExt, mp.sExt) == 0)
					return mp.iIcon;
				}

			// drat.  not there.  look it up in the registry
			RegistryKey rKey = Registry.ClassesRoot;

			rKey = rKey.OpenSubKey(sExt);

			if (rKey != null)
				{
				// find the name
				string sApp = (string)rKey.GetValue(null);

				if (sApp != null)
					{
					// now go to that guy and see if we get an icon
					rKey = Registry.ClassesRoot.OpenSubKey(sApp);

					if (rKey != null)
						{
						// woohoo...almost there
						rKey = rKey.OpenSubKey("DefaultIcon");
						if (rKey != null)
							{
							string sIcon = (string)rKey.GetValue(null);
							string sFile;
							int iIconInFile;

							// hooya!  now to actually get the image
							int iComma = sIcon.LastIndexOf(',');
							if (iComma >= 0)
								{
								sFile = sIcon.Substring(0, iComma);
								iIconInFile = Int32.Parse(sIcon.Substring(iComma + 1));
								}
							else
								{
								sFile = sIcon;
								iIconInFile = 0;
								}

							IntPtr hIcon1, hIcon2;
							int i = ole32.ExtractIconEx(sFile, iIconInFile, out hIcon1, out hIcon2, 1);
							if (i >= 0)
								{
								m_lvExplorer.LargeImageList.Images.Add(Bitmap.FromHicon(hIcon1));
								m_lvExplorer.SmallImageList.Images.Add(Bitmap.FromHicon(hIcon2));
								iIcon = m_lvExplorer.LargeImageList.Images.Count - 1;
								}

							if (hIcon1 != IntPtr.Zero)
								ole32.DestroyIcon(hIcon1);

							if (hIcon2 != IntPtr.Zero)
								ole32.DestroyIcon(hIcon2);
							}
						}
					}
				}
			mpExt mpNew = new mpExt();

			mpNew.sExt = sExt;
			mpNew.iIcon = iIcon;

			m_plmpExtensions.Add(mpNew);
			return iIcon;
		}

		/// <summary>
		/// Clean up any resources being used.
		/// </summary>
		/* D I S P O S E */
		/*----------------------------------------------------------------------------
			%%Function: Dispose
			%%Qualified: dfx.m_frm.Dispose
			%%Contact: rlittle

		----------------------------------------------------------------------------*/
		protected override void Dispose( bool disposing )
		{
			if( disposing )
			{
				if (components != null) 
				{
					components.Dispose();
				}
			}
			base.Dispose( disposing );
		}

		#region Windows Form Designer generated code
		/// <summary>
		/// Required method for Designer support - do not modify
		/// the contents of this method with the code editor.
		/// </summary>
		private void InitializeComponent()
		{
			this.components = new System.ComponentModel.Container();
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(m_frm));
			this.m_ebFilename = new System.Windows.Forms.TextBox();
			this.m_lvExplorer = new System.Windows.Forms.ListView();
			this.m_tlbMain = new System.Windows.Forms.ToolBar();
			this.m_tbbOpen = new System.Windows.Forms.ToolBarButton();
			this.m_tbbSep1 = new System.Windows.Forms.ToolBarButton();
			this.m_tbbNewStorage = new System.Windows.Forms.ToolBarButton();
			this.m_tbbSaveStream = new System.Windows.Forms.ToolBarButton();
			this.m_tbbDelete = new System.Windows.Forms.ToolBarButton();
			this.m_tbbSep2 = new System.Windows.Forms.ToolBarButton();
			this.m_tbbUp = new System.Windows.Forms.ToolBarButton();
			this.m_ilButtons = new System.Windows.Forms.ImageList(this.components);
			this.SuspendLayout();
			// 
			// m_ebFilename
			// 
			this.m_ebFilename.Location = new System.Drawing.Point(8, 8);
			this.m_ebFilename.Name = "m_ebFilename";
			this.m_ebFilename.Size = new System.Drawing.Size(336, 20);
			this.m_ebFilename.TabIndex = 0;
			this.m_ebFilename.Text = "c:\\footest.doc";
			this.m_ebFilename.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.HandleKeyDown);
			// 
			// m_lvExplorer
			// 
			this.m_lvExplorer.AllowDrop = true;
			this.m_lvExplorer.Location = new System.Drawing.Point(8, 32);
			this.m_lvExplorer.Name = "m_lvExplorer";
			this.m_lvExplorer.Size = new System.Drawing.Size(472, 416);
			this.m_lvExplorer.TabIndex = 2;
			this.m_lvExplorer.MouseDown += new System.Windows.Forms.MouseEventHandler(this.HandleMouseDown);
			this.m_lvExplorer.DoubleClick += new System.EventHandler(this.HandleDClick);
			this.m_lvExplorer.DragDrop += new System.Windows.Forms.DragEventHandler(this.HandleDrop);
			this.m_lvExplorer.DragEnter += new System.Windows.Forms.DragEventHandler(this.HandleDragEnter);
			this.m_lvExplorer.ItemDrag += new System.Windows.Forms.ItemDragEventHandler(this.HandleStartDrag);
			this.m_lvExplorer.DragLeave += new System.EventHandler(this.HandleDragLeave);
			// 
			// m_tlbMain
			// 
			this.m_tlbMain.Buttons.AddRange(new System.Windows.Forms.ToolBarButton[] {
																						 this.m_tbbOpen,
																						 this.m_tbbSep1,
																						 this.m_tbbNewStorage,
																						 this.m_tbbSaveStream,
																						 this.m_tbbDelete,
																						 this.m_tbbSep2,
																						 this.m_tbbUp});
			this.m_tlbMain.Divider = false;
			this.m_tlbMain.Dock = System.Windows.Forms.DockStyle.None;
			this.m_tlbMain.ImageList = this.m_ilButtons;
			this.m_tlbMain.Location = new System.Drawing.Point(350, 4);
			this.m_tlbMain.Name = "m_tlbMain";
			this.m_tlbMain.ShowToolTips = true;
			this.m_tlbMain.Size = new System.Drawing.Size(136, 23);
			this.m_tlbMain.TabIndex = 4;
			this.m_tlbMain.ButtonClick += new System.Windows.Forms.ToolBarButtonClickEventHandler(this.DispatchToolbarClick);
			// 
			// m_tbbOpen
			// 
			this.m_tbbOpen.ImageIndex = 0;
			this.m_tbbOpen.ToolTipText = "Open docfile";
			// 
			// m_tbbSep1
			// 
			this.m_tbbSep1.Style = System.Windows.Forms.ToolBarButtonStyle.Separator;
			// 
			// m_tbbNewStorage
			// 
			this.m_tbbNewStorage.ImageIndex = 3;
			this.m_tbbNewStorage.ToolTipText = "Create new storage";
			// 
			// m_tbbSaveStream
			// 
			this.m_tbbSaveStream.ImageIndex = 1;
			this.m_tbbSaveStream.ToolTipText = "Save selected stream";
			// 
			// m_tbbDelete
			// 
			this.m_tbbDelete.ImageIndex = 4;
			// 
			// m_tbbSep2
			// 
			this.m_tbbSep2.Style = System.Windows.Forms.ToolBarButtonStyle.Separator;
			// 
			// m_tbbUp
			// 
			this.m_tbbUp.ImageIndex = 2;
			this.m_tbbUp.ToolTipText = "Up";
			// 
			// m_ilButtons
			// 
			this.m_ilButtons.ColorDepth = System.Windows.Forms.ColorDepth.Depth8Bit;
			this.m_ilButtons.ImageSize = new System.Drawing.Size(16, 16);
			this.m_ilButtons.ImageStream = ((System.Windows.Forms.ImageListStreamer)(resources.GetObject("m_ilButtons.ImageStream")));
			this.m_ilButtons.TransparentColor = System.Drawing.Color.Magenta;
			// 
			// m_frm
			// 
			this.AllowDrop = true;
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(504, 470);
			this.Controls.AddRange(new System.Windows.Forms.Control[] {
																		  this.m_tlbMain,
																		  this.m_lvExplorer,
																		  this.m_ebFilename});
			this.Name = "m_frm";
			this.Text = "Form1";
			this.ResumeLayout(false);

		}
		#endregion

		/// <summary>
		/// The main entry point for the application.
		/// </summary>
		[STAThread]
		/* M A I N */
		/*----------------------------------------------------------------------------
			%%Function: Main
			%%Qualified: dfx.m_frm.Main
			%%Contact: rlittle

		----------------------------------------------------------------------------*/
		static void Main() 
		{
			Application.Run(new m_frm());
		}

		/* A D D  D I R E */
		/*----------------------------------------------------------------------------
			%%Function: AddDire
			%%Qualified: dfx.m_frm.AddDire
			%%Contact: rlittle

		----------------------------------------------------------------------------*/
		private void AddDire(DIRE dire)
		{
			ListViewItem lvi = new ListViewItem(dire.sName, dire.iIcon);
			lvi.Tag = dire;
			m_lvExplorer.Items.Add(lvi);
		}

		/* A D D  S S T G */
		/*----------------------------------------------------------------------------
			%%Function: AddSstg
			%%Qualified: dfx.m_frm.AddSstg
			%%Contact: rlittle

		----------------------------------------------------------------------------*/
		private void AddSstg(STATSTG sstg, DIRE direParent, ref ArrayList pldire)
		{
			// first, check to see if its already in pldire

			foreach (DIRE direT in pldire)
				{
				if (String.Compare(direT.sName, sstg.pwcsName) == 0)
					return;
				}

			string sExt;
			int iIcon = 0;

			if (sstg.type == (int)STGTY.STORAGE)
				{
				iIcon = 1;
				}
			else
				{
				int iDot = sstg.pwcsName.LastIndexOf('.');
				if (iDot >= 0)
					{
					sExt = sstg.pwcsName.Substring(iDot);
					iIcon = IEnsureIcon(sExt);
					}
				}

			DIRE dire = new DIRE(direParent);
			dire.pldire = null;
			dire.sName = sstg.pwcsName;
			dire.type = (STGTY)sstg.type;
			dire.istg = null;
			dire.iIcon = iIcon;
			pldire.Add(dire);
			AddDire(dire);
		}

		/* V I S I T  S T O R A G E */
		/*----------------------------------------------------------------------------
			%%Function: VisitStorage
			%%Qualified: dfx.m_frm.VisitStorage
			%%Contact: rlittle

		----------------------------------------------------------------------------*/
		private void VisitStorage(ole32.UCOMIStorage istg, DIRE direParent, ref ArrayList pldire)
		{
			ole32.UCOMIEnumSTATSTG iEnum;
			STATSTG sstg;

			iEnum = istg.EnumElements(0, IntPtr.Zero, 0); //, out iEnum);
			while (iEnum.Next(1, out sstg, IntPtr.Zero) == 0)
				{
				AddSstg(sstg, direParent, ref pldire);
				}
			Marshal.ReleaseComObject(iEnum);
			iEnum = null;
		}

		/* V I S I T  D I R E */
		/*----------------------------------------------------------------------------
			%%Function: VisitDire
			%%Qualified: dfx.m_frm.VisitDire
			%%Contact: rlittle

		----------------------------------------------------------------------------*/
		private void VisitDire(DIRE dire)
		{
			if (dire.istg == null || dire.pldire == null)
				{
				dire.direParent.istg.OpenStorage(dire.sName, IntPtr.Zero, ole32.STGM_READWRITE|ole32.STGM_SHARE_EXCLUSIVE, IntPtr.Zero, 0, out dire.istg);
				dire.pldire = new ArrayList();
				VisitStorage(dire.istg, dire, ref dire.pldire);
				}
			else
				{
				foreach (DIRE direT in dire.pldire)
					AddDire(direT);
				}
		}

		private void LoadDocfile(string sFile)
		{
			CloseDire(ref m_direRoot);
			m_direCurrent = null;

			m_lvExplorer.Items.Clear();
			ole32.UCOMIStorage istg;
			
			int hr = ole32.StgOpenStorage(sFile, IntPtr.Zero, ole32.STGM_READWRITE | ole32.STGM_SHARE_EXCLUSIVE, IntPtr.Zero, 0, out istg);
			if (hr != 0)
				throw(new Exception(String.Format("Failed: {0:X}", hr)));

//			istg = null;
			m_direRoot = new DIRE(null);
			m_direRoot.istg = istg;
			m_direRoot.sName = sFile;
			m_direRoot.type = STGTY.STORAGE;
			m_direRoot.pldire = new ArrayList();

			VisitStorage(istg, m_direRoot, ref m_direRoot.pldire);
			m_direCurrent = m_direRoot;
			m_ebFilename.Text = sFile;
		}

		/* L O A D  D O C F I L E */
		/*----------------------------------------------------------------------------
			%%Function: LoadDocfile
			%%Qualified: dfx.m_frm.LoadDocfile
			%%Contact: rlittle

		----------------------------------------------------------------------------*/
		private void LoadDocfile(object sender, System.EventArgs e)
		{
			LoadDocfile(m_sFile);
		}

		/* L I S T  V I E W */
		/*----------------------------------------------------------------------------
			%%Function: ListView
			%%Qualified: dfx.m_frm.ListView
			%%Contact: rlittle

		----------------------------------------------------------------------------*/
		private void ListView(object sender, System.EventArgs e)
		{
			m_lvExplorer.View = View.List;
		}

		/* D E T A I L S  V I E W */
		/*----------------------------------------------------------------------------
			%%Function: DetailsView
			%%Qualified: dfx.m_frm.DetailsView
			%%Contact: rlittle

		----------------------------------------------------------------------------*/
		private void DetailsView(object sender, System.EventArgs e)
		{
			m_lvExplorer.View = View.Details;
		}
		
		/* I C O N  V I E W */
		/*----------------------------------------------------------------------------
			%%Function: IconView
			%%Qualified: dfx.m_frm.IconView
			%%Contact: rlittle

		----------------------------------------------------------------------------*/
		private void IconView(object sender, System.EventArgs e)
		{
			m_lvExplorer.View = View.LargeIcon;
		}

		private void DeleteDire(DIRE dire, ListViewItem lvi)
		{
			if (MessageBox.Show("Delete '"+dire.sName+"'?", "DocFile Explorer", MessageBoxButtons.YesNo) == DialogResult.Yes)
				{
				CloseDire(ref dire);
				dire.direParent.istg.DestroyElement(dire.sName);
				m_lvExplorer.Items.Remove(lvi);
				dire.direParent.pldire.Remove(dire);
				}
		}

		/* D E L E T E  D I R E */
		/*----------------------------------------------------------------------------
			%%Function: DeleteDire
			%%Qualified: dfx.m_frm.DeleteDire
			%%Contact: rlittle

		----------------------------------------------------------------------------*/
		private void DeleteDire(object sender, System.EventArgs e)
		{
			ListViewItem lvi = m_lvExplorer.SelectedItems[0];
			DIRE dire = (DIRE)lvi.Tag;
			
			DeleteDire(dire, lvi);
		}

		private void AddStorage()
		{
			string sStorage;
			 
			if (InputBox.ShowInputBox("New storage name", out sStorage))
				{
				ole32.UCOMIStorage istgNew;
	
				m_direCurrent.istg.CreateStorage(sStorage, ole32.STGM_CREATE|ole32.STGM_READWRITE|ole32.STGM_SHARE_EXCLUSIVE, 0, 0, out istgNew);
	
					{
					STATSTG sstg = new STATSTG();
	
					sstg.pwcsName = sStorage;
					sstg.type = (int)STGTY.STORAGE;
					AddSstg(sstg, m_direCurrent, ref m_direCurrent.pldire);
					istgNew.Commit(0);
					Marshal.ReleaseComObject(istgNew);
					istgNew = null;
					// add it to the entry
	
					}
				}
		}

		/* A D D  S T O R A G E */
		/*----------------------------------------------------------------------------
			%%Function: AddStorage
			%%Qualified: dfx.m_frm.AddStorage
			%%Contact: rlittle

		----------------------------------------------------------------------------*/
		private void AddStorage(object sender, System.EventArgs e)
		{
			AddStorage();
		}

		private void CopyStreamToStream(UCOMIStream istmFrom, UCOMIStream istmTo)
		{
			long lcbRead;
			IntPtr pLcb = Marshal.AllocHGlobal(4);
			while (true)
				{
				istmFrom.Read(m_rgb, lcbMax, pLcb);
				lcbRead = Marshal.ReadInt32(pLcb);

				if (lcbRead <= 0)
					break;

				istmTo.Write(m_rgb, (int)lcbRead, IntPtr.Zero);
				}
		}

		private void CopyStorageToStorage(ole32.UCOMIStorage istgFrom, ole32.UCOMIStorage istgTo)
		{
			STATSTG sstg;

			istgFrom.Stat(out sstg, ole32.STATFLAG_NONAME);

			istgTo.SetClass(ref sstg.clsid);

			ole32.UCOMIEnumSTATSTG iEnum;

			iEnum = istgFrom.EnumElements(0, IntPtr.Zero, 0); //, out iEnum);
			while (iEnum.Next(1, out sstg, IntPtr.Zero) == 0)
				{
				switch ((STGTY)sstg.type)
					{
					case STGTY.STORAGE:
						{
						ole32.UCOMIStorage istgFromChild, istgToChild;

						istgFrom.OpenStorage(sstg.pwcsName, IntPtr.Zero, ole32.STGM_READ | ole32.STGM_SHARE_EXCLUSIVE, IntPtr.Zero, 0, out istgFromChild);
						istgTo.CreateStorage(sstg.pwcsName, ole32.STGM_CREATE|ole32.STGM_READWRITE|ole32.STGM_SHARE_EXCLUSIVE, 0, 0, out istgToChild);

						CopyStorageToStorage(istgFromChild, istgToChild);
						Marshal.ReleaseComObject(istgFromChild);
						istgFromChild = null;
						Marshal.ReleaseComObject(istgToChild);
						istgToChild = null;
						break;
						}
					case STGTY.STREAM:
						{
						UCOMIStream istmFrom, istmTo;

						istmFrom = istgFrom.OpenStream(sstg.pwcsName, IntPtr.Zero, ole32.STGM_READ|ole32.STGM_SHARE_EXCLUSIVE, 0);
						istgTo.CreateStream(sstg.pwcsName, ole32.STGM_CREATE|ole32.STGM_READWRITE|ole32.STGM_SHARE_EXCLUSIVE, 0, 0, out istmTo);

						CopyStreamToStream(istmFrom, istmTo);
						Marshal.ReleaseComObject(istmFrom);
						istmFrom = null;
						Marshal.ReleaseComObject(istmTo);
						istmTo = null;
						break;
						}
					}
				}
			Marshal.ReleaseComObject(iEnum);
			iEnum = null;
		}

		private void CopyStorageTo(DIRE dire, string s)
		{
			ole32.UCOMIStorage istg;

			if (dire.istg != null)
				istg = dire.istg;
			else
                m_direCurrent.istg.OpenStorage(dire.sName, IntPtr.Zero, ole32.STGM_READ | ole32.STGM_SHARE_EXCLUSIVE, IntPtr.Zero, 0, out istg);

			ole32.UCOMIStorage istgNew = null;

			int hr = ole32.StgCreateDocfile(s, ole32.STGM_READWRITE | ole32.STGM_SHARE_EXCLUSIVE, 0, out istgNew);

			if (hr != 0)
				throw(new Exception(String.Format("Failed: {0:X}", hr)));

			CopyStorageToStorage(istg, istgNew);
			Marshal.ReleaseComObject(istgNew);
			istgNew = null;

		}

		private void CopyStreamTo(DIRE dire, string s)
		{
			UCOMIStream istm;

			istm = m_direCurrent.istg.OpenStream(dire.sName, IntPtr.Zero, ole32.STGM_READ|ole32.STGM_SHARE_EXCLUSIVE, 0);
			FileStream bs = new FileStream(s, FileMode.Create, FileAccess.Write, FileShare.None, 8, false);
			long lcbRead;
			IntPtr pLcb = Marshal.AllocHGlobal(4);
			while (true)
				{
				istm.Read(m_rgb, lcbMax, pLcb);
				lcbRead = Marshal.ReadInt32(pLcb);

				if (lcbRead <= 0)
					break;

				bs.Write(m_rgb, 0, (int)lcbRead);
				}

			Marshal.ReleaseComObject(istm);
			istm = null;
			bs.Close();
			bs = null;
		}

		private void HandleSave(DIRE dire)
		{
			string s;

			switch (dire.type)
				{
				case STGTY.STREAM:
					if (InputBox.ShowInputBox("Save stream where?", dire.sName, out s))
						CopyStreamTo(dire, s);
					break;
				case STGTY.STORAGE:
					if (InputBox.ShowInputBox("Save docfile where?", dire.sName, out s))
						CopyStorageTo(dire, s);
					break;
				}

		}

		/* H A N D L E  D  C L I C K */
		/*----------------------------------------------------------------------------
			%%Function: HandleDClick
			%%Qualified: dfx.m_frm.HandleDClick
			%%Contact: rlittle

		----------------------------------------------------------------------------*/
		private void HandleDClick(object sender, System.EventArgs e) 
		{
			ListViewItem lvi = ((ListView)sender).SelectedItems[0];
			DIRE dire = (DIRE)lvi.Tag;

			if (dire.type == STGTY.STREAM)
				{
				HandleSave(dire);
				}
			else
				{
				m_lvExplorer.Items.Clear();

				VisitDire(dire);
				m_direCurrent = dire;
				}
		}

		private const int lcbMax = 16384;

		private byte []m_rgb;

		/* A D D  S T R E A M */
		/*----------------------------------------------------------------------------
			%%Function: AddStream
			%%Qualified: dfx.m_frm.AddStream
			%%Contact: rlittle

		----------------------------------------------------------------------------*/
		private void AddStream(string sStream)
		{
			string sName;
			string sStreamName = Path.GetFileName(sStream);
			if (InputBox.ShowInputBox("Add file as?", sStreamName, out sName))
				{
				UCOMIStream istm;
	
				m_direCurrent.istg.CreateStream(sName, ole32.STGM_CREATE|ole32.STGM_READWRITE|ole32.STGM_SHARE_EXCLUSIVE, 0, 0, out istm);
	
				FileStream bs = new FileStream(sStream, FileMode.Open, FileAccess.Read, FileShare.None, 8, false);
				int lcb;
	
				while ((lcb = bs.Read(m_rgb, 0, lcbMax)) > 0)
					{
					istm.Write(m_rgb, lcb, IntPtr.Zero);
					}
				bs.Close();
				bs = null;
				STATSTG sstg = new STATSTG();
	
				sstg.pwcsName = sName;
				sstg.type = (int)STGTY.STREAM;
				AddSstg(sstg, m_direCurrent, ref m_direCurrent.pldire);
				// add it to the entry
				}

		}

		/* H A N D L E  D R O P */
		/*----------------------------------------------------------------------------
			%%Function: HandleDrop
			%%Qualified: dfx.m_frm.HandleDrop
			%%Contact: rlittle

		----------------------------------------------------------------------------*/
		private void HandleDrop(object sender, System.Windows.Forms.DragEventArgs e)
		{
			string []files = (string[])e.Data.GetData(DataFormats.FileDrop);

			foreach (string sFile in files)
				{
                AddStream(sFile);
				}
		}

		/* H A N D L E  D R A G  E N T E R */
		/*----------------------------------------------------------------------------
			%%Function: HandleDragEnter
			%%Qualified: dfx.m_frm.HandleDragEnter
			%%Contact: rlittle

		----------------------------------------------------------------------------*/
		private void HandleDragEnter(object sender, System.Windows.Forms.DragEventArgs e)
		{
			if (e.Data.GetDataPresent(DataFormats.FileDrop)) 
				{
				e.Effect = DragDropEffects.Copy;
				}
			else
				{
				e.Effect = DragDropEffects.None;
				}
		}

		private void HandleMouseDown(object sender, System.Windows.Forms.MouseEventArgs e)
		{
//			Control ctl = (Control) sender;
			if (e.Button == MouseButtons.Left)
				{
				}

		}

		void GoUp()
		{
			if (m_direCurrent.direParent != null)
				{
				m_lvExplorer.Items.Clear();

				VisitDire(m_direCurrent.direParent);
				m_direCurrent = m_direCurrent.direParent;
				}
		}

		/* G O  U P */
		/*----------------------------------------------------------------------------
			%%Function: GoUp
			%%Qualified: dfx.m_frm.GoUp
			%%Contact: rlittle

		----------------------------------------------------------------------------*/
		private void GoUp(object sender, System.EventArgs e) 
		{
			GoUp();
		}

		/* H A N D L E  D R A G  L E A V E */
		/*----------------------------------------------------------------------------
			%%Function: HandleDragLeave
			%%Qualified: dfx.m_frm.HandleDragLeave
			%%Contact: rlittle

		----------------------------------------------------------------------------*/
		private void HandleDragLeave(object sender, System.EventArgs e)
		{
		
		}

		private void CloseDire(ref DIRE dire)
		{
			if (dire == null)
				return;

			// depth first close of everything we have
			if (dire.pldire != null)
				{
				int iMac;

				iMac = dire.pldire.Count - 1;   

				while (iMac >= 0)
					{
					DIRE direChild = (DIRE)dire.pldire[iMac];

					CloseDire(ref direChild);
					dire.pldire.RemoveAt(iMac);
					iMac--;
					}
				}

			if (dire.istg != null)
				{
				Marshal.ReleaseComObject(dire.istg);
				dire.istg = null;
				}

			dire = null;
		}

		private void HandleStartDrag(object sender, System.Windows.Forms.ItemDragEventArgs e)
		{
//			Point pt = new Point(e.X, e.Y);
//			ListViewItem lvi = ((ListView)sender).GetItemAt(pt.X, pt.Y);
			ListViewItem lvi = (ListViewItem)e.Item;
			if (lvi != null)
				{
				DIRE dire = (DIRE)lvi.Tag;

				string s = Path.GetFileName(dire.sName);
				string sTemp = Environment.GetEnvironmentVariable("TEMP");
				string sDest = Path.Combine(sTemp, s);
				string []rgs = new string[1];

				rgs[0] = sDest;
				if (dire.type == STGTY.STORAGE)
					{
					CopyStorageTo(dire, sDest);
					}
				else
					{
					CopyStreamTo(dire, sDest);
					}

				DataObject ido = new DataObject(DataFormats.FileDrop, rgs);

				m_lvExplorer.DoDragDrop(ido, DragDropEffects.Copy);
				}
		
		}

		private void DispatchToolbarClick(object sender, System.Windows.Forms.ToolBarButtonClickEventArgs e)
		{
			switch ((int)e.Button.Tag)
				{
				case itbbSave:
					{
					ListViewItem lvi = m_lvExplorer.SelectedItems[0];

					if (lvi != null)
						{
						DIRE dire = (DIRE)lvi.Tag;

						if (dire.type == STGTY.STREAM || dire.type == STGTY.STORAGE)
							HandleSave(dire);
						}
					break;
					}
				case itbbDelete:
					{
					ListViewItem lvi = m_lvExplorer.SelectedItems[0];
	
					if (lvi != null)
						{
						DIRE dire = (DIRE)lvi.Tag;
	
						DeleteDire(dire, lvi);
						}
					break;
					}
				case itbbNew:
					AddStorage();
					break;
				case itbbUp:
					GoUp();
					break;
				case itbbOpen:
					{
					OpenFileDialog ofd = new OpenFileDialog();

					if (m_sFile != null)
                        ofd.InitialDirectory = Path.GetDirectoryName(m_sFile);
					ofd.Filter = "Docfiles (*.doc)|*.doc|All Files (*.*)|*.*";

					if (ofd.ShowDialog() == DialogResult.OK)
						m_sFile = ofd.FileName;

					LoadDocfile(m_sFile);
					break;
					}
				}

		}

		private void HandleKeyDown(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (e.KeyChar == '\r')
				{
				LoadDocfile(m_ebFilename.Text);				
				e.Handled = true;
				m_lvExplorer.Focus();
				}
		
		}
	}
}

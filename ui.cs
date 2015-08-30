/*----------------------------------------------------------------------------
	%%File: UI.CS
	%%Unit: UI
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

namespace UI
{
	public class InputBox : System.Windows.Forms.Form
	{
		private System.Windows.Forms.TextBox textBox1;
		private System.ComponentModel.Container components = null;
		private System.Windows.Forms.Button button1, button2;
		private bool m_fCanceled = false;

		/* I N P U T  B O X */
		/*----------------------------------------------------------------------------
			%%Function: InputBox
			%%Qualified: UI.InputBox.InputBox
			%%Contact: rlittle

		----------------------------------------------------------------------------*/
		private InputBox(string sPrompt, string sText)
		{
			InitializeComponent();
			if (sText != null)
				textBox1.Text = sText;
	
			this.Text = sPrompt;
		}
	
		/* D I S P O S E */
		/*----------------------------------------------------------------------------
			%%Function: Dispose
			%%Qualified: UI.InputBox.Dispose
			%%Contact: rlittle

		----------------------------------------------------------------------------*/
		protected override void Dispose( bool disposing )
		{
			if( disposing )
			{
				if(components != null)
				{
					components.Dispose();
				}
			}
			base.Dispose( disposing );
		}
	
		/* I N I T I A L I Z E  C O M P O N E N T */
		/*----------------------------------------------------------------------------
			%%Function: InitializeComponent
			%%Qualified: UI.InputBox.InitializeComponent
			%%Contact: rlittle

		----------------------------------------------------------------------------*/
		private void InitializeComponent()
		{
			this.textBox1 = new System.Windows.Forms.TextBox();
			this.button1 = new Button();
			this.button2 = new Button();
			this.SuspendLayout();
			//
			// textBox1
			//
			this.textBox1.Location = new System.Drawing.Point(16, 16);
			this.textBox1.Name = "textBox1";
			this.textBox1.Size = new System.Drawing.Size(256, 20);
			this.textBox1.TabIndex = 0;
			this.textBox1.Text = "textBox1";
			this.textBox1.KeyDown += new System.Windows.Forms.KeyEventHandler(this.textBox1_KeyDown);
			// 
			// button1
			// 
			this.button1.Location = new System.Drawing.Point(168, 48);
			this.button1.Name = "button1";
			this.button1.Size = new System.Drawing.Size(48, 24);
			this.button1.TabIndex = 1;
			this.button1.Text = "OK";
			this.button1.Click += new System.EventHandler(this.HandleOK);
			// 
			// button2
			// 
			this.button2.Location = new System.Drawing.Point(224, 48);
			this.button2.Name = "button2";
			this.button2.Size = new System.Drawing.Size(48, 24);
			this.button2.TabIndex = 2;
			this.button2.Text = "Cancel";
			this.button2.Click += new System.EventHandler(this.HandleCancel);
			//
			// InputBox
			//
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(292, 93);
			this.ControlBox = false;
			this.Controls.AddRange(new System.Windows.Forms.Control[] {
																		 this.button2,
																		 this.button1,
																		 this.textBox1});
			this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
			this.Name = "InputBox";
			this.Text = "InputBox";
			this.ResumeLayout(false);
	
		}
	
		/* T E X T  B O X  1  _ K E Y  D O W N */
		/*----------------------------------------------------------------------------
			%%Function: textBox1_KeyDown
			%%Qualified: UI.InputBox.textBox1_KeyDown
			%%Contact: rlittle

		----------------------------------------------------------------------------*/
		private void textBox1_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter)
				{
				m_fCanceled = false;
				this.Close();
				}
		}
	
		private void HandleOK(object sender, System.EventArgs e) 
		{
			m_fCanceled = false;
			this.Close();
		}

		private void HandleCancel(object sender, System.EventArgs e) 
		{
			m_fCanceled = true;
			this.Close();
		}


		/* S H O W  I N P U T  B O X */
		/*----------------------------------------------------------------------------
			%%Function: ShowInputBox
			%%Qualified: UI.InputBox.ShowInputBox
			%%Contact: rlittle

		----------------------------------------------------------------------------*/
		public static bool ShowInputBox(string sPrompt, out string sResponse)
		{
			return ShowInputBox(sPrompt, null, out sResponse);
		}
	
		/* S H O W  I N P U T  B O X */
		/*----------------------------------------------------------------------------
			%%Function: ShowInputBox
			%%Qualified: UI.InputBox.ShowInputBox
			%%Contact: rlittle

		----------------------------------------------------------------------------*/
		public static bool ShowInputBox(string sPrompt, string s, out string sResponse)
		{
			InputBox box = new InputBox(sPrompt, s);
			box.m_fCanceled = false;

			box.ShowDialog();
			sResponse = box.textBox1.Text;
			return !box.m_fCanceled;
		}    
	}
}

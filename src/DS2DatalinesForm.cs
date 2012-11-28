using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using SAS.Shared.AddIns;

namespace SASPress.CustomTasks.DS2Datalines
{
	/// <summary>
	/// The simple Windows form that shows a preview of the SAS program.
    /// It uses a control that hosts the SAS enhanced program editor.
	/// </summary>
	public class DS2DatalinesForm : SAS.Tasks.Toolkit.Controls.TaskForm
    {
        #region private form members
        private System.Windows.Forms.Button btnOK;
		private System.Windows.Forms.Button btnCancel;
        private System.Windows.Forms.Label lblOutput;
		private System.Windows.Forms.Label lblProgram;
        private System.Windows.Forms.Button btnBrowse;
		private System.Windows.Forms.TextBox txtOutput;
		private System.Windows.Forms.Button btnCopy;
        private SAS.Tasks.Toolkit.Controls.SASTextEditorCtl sasCodeEditor;
        private CheckBox chkPreserveEncoding;
        #endregion

        #region Constructor, initialization, and cleanup
        /// <summary>
		/// Required designer variable.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public DS2DatalinesForm()
		{
			InitializeComponent();
		}

		protected override void OnLoad(EventArgs e)
		{
			base.OnLoad (e);

			txtOutput.Text = TaskModel.OutputData;
            chkPreserveEncoding.Checked = TaskModel.PreserveEncoding;
            RefreshCodeView();
		}

		/// <summary>
		/// Clean up any resources being used.
		/// </summary>
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

        #endregion

        #region Properties
        // reference to the task model class
		internal DS2Datalines TaskModel {get; set;}
        #endregion

        #region Windows Form Designer generated code
        /// <summary>
		/// Required method for Designer support - do not modify
		/// the contents of this method with the code editor.
		/// </summary>
		private void InitializeComponent()
		{
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(DS2DatalinesForm));
            this.btnOK = new System.Windows.Forms.Button();
            this.btnCancel = new System.Windows.Forms.Button();
            this.lblOutput = new System.Windows.Forms.Label();
            this.txtOutput = new System.Windows.Forms.TextBox();
            this.lblProgram = new System.Windows.Forms.Label();
            this.btnBrowse = new System.Windows.Forms.Button();
            this.btnCopy = new System.Windows.Forms.Button();
            this.sasCodeEditor = new SAS.Tasks.Toolkit.Controls.SASTextEditorCtl();
            this.chkPreserveEncoding = new System.Windows.Forms.CheckBox();
            this.SuspendLayout();
            // 
            // btnOK
            // 
            resources.ApplyResources(this.btnOK, "btnOK");
            this.btnOK.DialogResult = System.Windows.Forms.DialogResult.OK;
            this.btnOK.Name = "btnOK";
            // 
            // btnCancel
            // 
            resources.ApplyResources(this.btnCancel, "btnCancel");
            this.btnCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.btnCancel.Name = "btnCancel";
            // 
            // lblOutput
            // 
            resources.ApplyResources(this.lblOutput, "lblOutput");
            this.lblOutput.Name = "lblOutput";
            // 
            // txtOutput
            // 
            resources.ApplyResources(this.txtOutput, "txtOutput");
            this.txtOutput.Name = "txtOutput";
            this.txtOutput.ReadOnly = true;
            // 
            // lblProgram
            // 
            resources.ApplyResources(this.lblProgram, "lblProgram");
            this.lblProgram.Name = "lblProgram";
            // 
            // btnBrowse
            // 
            resources.ApplyResources(this.btnBrowse, "btnBrowse");
            this.btnBrowse.Name = "btnBrowse";
            this.btnBrowse.Click += new System.EventHandler(this.btnBrowse_Click);
            // 
            // btnCopy
            // 
            resources.ApplyResources(this.btnCopy, "btnCopy");
            this.btnCopy.Name = "btnCopy";
            this.btnCopy.Click += new System.EventHandler(this.btnCopy_Click);
            // 
            // sasCodeEditor
            // 
            resources.ApplyResources(this.sasCodeEditor, "sasCodeEditor");
            this.sasCodeEditor.BackColor = System.Drawing.SystemColors.Window;
            this.sasCodeEditor.ContentType = SAS.Tasks.Toolkit.Controls.SASTextEditorCtl.eContentType.SASProgram;
            this.sasCodeEditor.EditorText = null;
            this.sasCodeEditor.Name = "sasCodeEditor";
            this.sasCodeEditor.ReadOnly = true;
            // 
            // chkPreserveEncoding
            // 
            resources.ApplyResources(this.chkPreserveEncoding, "chkPreserveEncoding");
            this.chkPreserveEncoding.Name = "chkPreserveEncoding";
            this.chkPreserveEncoding.UseVisualStyleBackColor = true;
            // 
            // DS2DatalinesForm
            // 
            this.AcceptButton = this.btnOK;
            resources.ApplyResources(this, "$this");
            this.CancelButton = this.btnCancel;
            this.Controls.Add(this.chkPreserveEncoding);
            this.Controls.Add(this.sasCodeEditor);
            this.Controls.Add(this.btnCopy);
            this.Controls.Add(this.btnBrowse);
            this.Controls.Add(this.lblProgram);
            this.Controls.Add(this.txtOutput);
            this.Controls.Add(this.lblOutput);
            this.Controls.Add(this.btnCancel);
            this.Controls.Add(this.btnOK);
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "DS2DatalinesForm";
            this.ResumeLayout(false);
            this.PerformLayout();

		}
		#endregion

        #region event handlers
        /// <summary>
        /// Refresh the SAS program based on the output data field
        /// </summary>
        private void RefreshCodeView()
        {
            SAS.Tasks.Toolkit.Data.SasData sd = new SAS.Tasks.Toolkit.Data.SasData(Consumer.ActiveData as ISASTaskData2);

            DatasetConverter.AssignLocalLibraryIfNeeded(Consumer,TaskModel);

            DatasetConverter convert = new DatasetConverter(
                sd.Server,
                sd.Libref,
                sd.Member);

            // Put just the DATA step definition in the preview window
            // This is a quick operation.  If the data set is large,
            // it could take several moments to populate all of the data
            // values, so we skip that part for the preview.
            sasCodeEditor.EditorText = convert.GetJustDataStepDefinition(txtOutput.Text);
        }

         /// <summary>
        /// Launch the application's "save as" dialog to select a location
        /// for the output data set.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
		private void btnBrowse_Click(object sender, System.EventArgs e)
		{
            // turn the LIBNAME.MEMBER string into an array of
            // strings for use in the ShowOutputDataSelector call.
			string[] parts = txtOutput.Text.Split('.');
			string cookie="";

            // show the File dialog to let the user select a different
            // output data set.
			ISASTaskDataName name = Consumer.ShowOutputDataSelector(this, 
                ServerAccessMode.OneServer, 
                Consumer.AssignedServer, parts[0], parts[1], ref cookie);

            // if a new value was supplied, update the form with the
            // new value and refresh the preview window (program).
			if (name!=null)
			{
				txtOutput.Text = string.Format("{0}.{1}", name.Library, name.Member);
                RefreshCodeView();
			}
		}

        /// <summary>
        /// Copy the SAS program to the clipboard.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
		private void btnCopy_Click(object sender, System.EventArgs e)
		{
            // This call can occasionally throw an exception when running
            // in terminal server mode (remote desktop or Citrix, for example).
            // It's nothing to worry about, but we don't want to halt the task
            // so we need to catch it here.
			try
			{
				System.Windows.Forms.Clipboard.SetDataObject(sasCodeEditor.EditorText);
			}
			catch (System.Runtime.InteropServices.ExternalException)
            {}
        }

        /// <summary>
        /// Window is closing
        /// </summary>
        /// <param name="e"></param>
        protected override void OnClosing(CancelEventArgs e)
        {
            // store the name of the output data set
            // but only if the window is closing due to
            // the OK button being pressed.
            if (DialogResult.OK == DialogResult)
            {
                TaskModel.OutputData = txtOutput.Text;
                TaskModel.PreserveEncoding = chkPreserveEncoding.Checked;
            }

            base.OnClosing(e);
        }
        #endregion
    }
}

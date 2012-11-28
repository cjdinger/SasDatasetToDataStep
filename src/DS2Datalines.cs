using System;
using System.Collections.Generic;
using System.Text;
using SAS.Tasks.Toolkit;
using SAS.Shared.AddIns;
using System.IO;
using System.Xml;
using System.Data.OleDb;
using System.Windows.Forms;

namespace SASPress.CustomTasks.DS2Datalines
{
    /// <summary>
    /// Class: DS2Datalines
    /// This class implements the task add-in interfaces that allow
    /// the task to plug into the application.
    /// 
    /// This class implements ISASTaskExecution, which tells the application
    /// that it will handle the "run" actions itself.
    /// </summary>
    // The ClassId is the unique identifier for this task.
    [SAS.Tasks.Toolkit.ClassId("32E78280-DBEE-460F-85AB-2727143F7938")]
    // The version attribute is simply informational -- for your own version tracking
    [SAS.Tasks.Toolkit.Version(4.2)]
    // The IconLocation attribute supplies the path to find the task icon within this assembly.
    // The icon must be provided as an embedded resource.
    [SAS.Tasks.Toolkit.IconLocation("SASPress.CustomTasks.DS2Datalines.task.ico")]
    // This task uses a data source (data set) as input
    [SAS.Shared.AddIns.InputRequired(InputResourceType.Data)]
    // ApplicationSupported is set to just SAS Enterprise Guide, because this
    // task won't make sense with the SAS Add-In for Microsoft Office
    [SAS.Shared.AddIns.ApplicationSupported(ApplicationName.EGuide)]
    public class DS2Datalines : SAS.Tasks.Toolkit.SasTask, SAS.Shared.AddIns.ISASTaskExecution
    {
        #region Constructor/Initialization
        public DS2Datalines()
        {
            InitializeComponent();

            OutputData = string.Empty;
            PreserveEncoding = true;
        }

        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(DS2Datalines));
            // 
            // DS2Datalines
            // 
            this.GeneratesReportOutput = false;
            this.GeneratesSasCode = false;
            this.ProcsUsed = "DATA step";
            this.ProductsRequired = "BASE";
            resources.ApplyResources(this, "$this");

        }
        #endregion

        #region Properties
        /// <summary>
        /// By making this a public property, we can allow the
        /// dialog form to modify it directly when the end user
        /// selects another libref.member with the Browse button.
        /// </summary>
        public string OutputData { get; set; }

        /// <summary>
        /// Whether to preserve the character encoding 
        /// (saving as UTF-8 instead of ASCII)
        /// </summary>
        public bool PreserveEncoding { get; set; }
        #endregion

        // ISASTaskExecution includes 4 methods that the application
        // (SAS Enterprise Guide) uses to control how the task runs.
        #region ISASTaskExecution implementation

        /// <summary>
        /// Tells the task to cancel its work.  This is called when the
        /// end user selects Stop in the task status window within
        /// SAS Enterprise Guide. 
        /// Implementing true cancel support is optional.  You
        /// can return false if you don't implement cancel support, or if
        /// your task is not "interruptable".
        /// </summary>
        /// <returns>Whether the task was canceled.</returns>
        public bool Cancel()
        {
            // this task does not implement Cancel, so 
            // return false
            return false;
        }

        /// <summary>
        /// How many result structures does this task return?
        /// Just 1 -- it's the SAS program with the DATA step.
        /// </summary>
        public int ResultCount
        {
            get
            {
                return 1;
            }
        }

        /// <summary>
        /// Allows the task to supply its results to the application.
        /// The results are packaged within a class (ResultInfo)
        /// that can deliver its results via the ITaskStream interface.
        /// This method will be called for each result that you advertise 
        /// via the ResultCount property.
        /// </summary>
        /// <param name="Index">Which result stream to supply.</param>
        /// <returns>An ISASTaskStream interface that points to the results.  You can use the
        /// SAS Task Toolkit (SAS.Tasks.Toolkit.Helpers.StreamAdapter) to provide a simple
        /// wrapper.</returns>
        public ISASTaskStream OpenResultStream(int index)
        {
            // this task returns just one result, the SAS program.
            // So we handle this only when the index is 0.
            if (index == 0)
            {
                return new SAS.Tasks.Toolkit.Helpers.StreamAdapter(
                    new MemoryStream(_resultInfo.Bytes),
                    _resultInfo.MimeType,
                    _resultInfo.OriginalFileName);
            }
            else return null;
        }



        SAS.Tasks.Toolkit.Helpers.ResultInfo _resultInfo;

        /// <summary>
        /// Performs the main work of this task.
        /// </summary>
        /// <param name="LogWriter">A log interface that you can use to report messages.</param>
        /// <returns>a status code: did it succeed or fail with errors?</returns>
        public RunStatus Run(ISASTaskTextWriter LogWriter)
        {
            try
            {
                string code = "";

                SAS.Tasks.Toolkit.Data.SasData sd = new SAS.Tasks.Toolkit.Data.SasData(Consumer.ActiveData as ISASTaskData2);
                DatasetConverter.AssignLocalLibraryIfNeeded(Consumer, this);

                DatasetConverter convert = new DatasetConverter(
                    sd.Server,
                    sd.Libref,
                    sd.Member);

                code = convert.GetCompleteSasProgram(OutputData);

                // this builds up the result structure.  It contains the SAS program
                // that was constructed.  By tagging it with the 
                // "application/x-sas" mime type, that serves as a cue
                // to SAS Enterprise Guide to treat this as the generated SAS
                // program in the project.
                _resultInfo = new SAS.Tasks.Toolkit.Helpers.ResultInfo();
                if (PreserveEncoding)
                    _resultInfo.Bytes = System.Text.Encoding.UTF8.GetBytes(code.ToCharArray());
                else
                    _resultInfo.Bytes = System.Text.Encoding.ASCII.GetBytes(code.ToCharArray());
                _resultInfo.MimeType = "application/x-sas";
                _resultInfo.OriginalFileName = string.Format("{0}.sas", Consumer.ActiveData.Member);

                // this puts a summary message into the Log portion of the
                // task in SAS Enterprise Guide
                SAS.Tasks.Toolkit.Helpers.FormattedLogWriter.WriteNormalLine(LogWriter, Messages.ConvertedData);
                SAS.Tasks.Toolkit.Helpers.FormattedLogWriter.WriteNoteLine(LogWriter,
                    string.Format(Messages.ConvertedDataMetrics, convert.ColumnCount, convert.RowCount));

            }
            catch (Exception ex)
            {
                // catch any loose exceptions and report in the log
                SAS.Tasks.Toolkit.Helpers.FormattedLogWriter.WriteErrorLine(LogWriter,
                    string.Format(Messages.ErrorDuringConversion, ex.ToString()));
                return RunStatus.Error;
            }

            return RunStatus.Success;
        }


        #endregion

        #region Overrides

        /// <summary>
        /// SAS Enterprise Guide calls into this method when the 
        /// task is initialized with data.
        /// </summary>
        /// <param name="consumer"></param>
        /// <returns></returns>
        public override bool Connect(ISASTaskConsumer consumer)
        {
            // let the base class initialize things
            base.Connect(consumer);

            // if the default output name is not yet set,
            // calculate a name using the libref.member input name
            if (string.IsNullOrEmpty(OutputData))
            {
                // initialize the name of the output data
                string membername = SAS.Tasks.Toolkit.Helpers.UtilityFunctions.GetValidSasName(
                    // seed the member name with the libref.member of the input data
                    string.Format("{0}.{1}", Consumer.ActiveData.Library, Consumer.ActiveData.Member),
                    // capped at a max length of 32
                    32);
                OutputData = string.Format("WORK.{0}", membername);
            }
            return true;
        }

        /// <summary>
        /// Show the form for the task.
        /// </summary>
        /// <param name="Owner">The handle to the owning application</param>
        /// <returns>the result -- RunNow if OK, Canceled if Cancel</returns>
        public override SAS.Shared.AddIns.ShowResult Show(System.Windows.Forms.IWin32Window Owner)
        {
            // Show the default form for this custom task
            DS2DatalinesForm dlg = new DS2DatalinesForm();
            dlg.Text = Label;
            dlg.TaskModel = this;
            dlg.Consumer = Consumer;

            SAS.Tasks.Toolkit.Helpers.TaskAddInHelpers.InitializeTaskDisplaySettings(Clsid, dlg);
            if (dlg.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                SAS.Tasks.Toolkit.Helpers.TaskAddInHelpers.StoreTaskDisplaySettings(Clsid, dlg);
                return SAS.Shared.AddIns.ShowResult.RunNow;
            }
            else
                return SAS.Shared.AddIns.ShowResult.Canceled;
        }

        #region Serialization
        // The XML state information for this task is very simple:
        //     <DS2Datalines OutputData="libref.member" PreserveEncoding="True|False"/>
        // The GetXmlState and RestoreXmlState methods read and write
        // this simple XML document.

        /// <summary>
        /// Collects the state information from the task.
        /// This information is stored within the project.
        /// For this task, the only option to remember is
        /// the output data set.  The input data set is
        /// tracked within the project, not here within the
        /// task.
        /// </summary>
        /// <returns>The state of the task in XML.</returns>
        public override string GetXmlState()
        {
            // create a simple XML document that stores just the
            // output data name.
            XmlDocument doc = new XmlDocument();
            XmlElement el = doc.CreateElement("DS2Datalines");
            el.SetAttribute("OutputData", OutputData);
            el.SetAttribute("PreserveEncoding", XmlConvert.ToString(PreserveEncoding));
            doc.AppendChild(el);
            return doc.OuterXml;
        }

        /// <summary>
        /// Re-applies the task settings as remembered from within
        /// the project.
        /// </summary>
        /// <param name="xml">The state of the task in XML.</param>
        public override void RestoreStateFromXml(string xml)
        {
            if (xml != null && xml.Length > 0)
            {
                try
                {
                    // expecting a simple XML document with
                    // just the output data location.
                    XmlDocument doc = new XmlDocument();
                    doc.LoadXml(xml);
                    XmlElement el = doc["DS2Datalines"];
                    OutputData = el.Attributes["OutputData"].Value;
                    PreserveEncoding = XmlConvert.ToBoolean(el.Attributes["PreserveEncoding"].Value);
                }
                catch
                {
                    // if the XML is not valid, there isn't
                    // anything we can really do about that.
                    // We'll simply resort to the default value.
                }
            }
        }
        #endregion // Serialization

        #endregion
    }
}

using System;
using System.Collections.Generic;
using System.Text;
using System.Data.OleDb;
using SAS.Tasks.Toolkit;
using SAS.Shared.AddIns;

namespace SASPress.CustomTasks.DS2Datalines
{
    /// <summary>
    /// Class: DatasetConverter
    /// Converts a SAS data set into a DATA step program.
    /// It does this by reading the attributes of the data set
    /// from the SAS dictionary tables and building a DATA step 
    /// definition.  It then reads the actual values from the 
    /// data set and creates a series of DATALINES entries
    /// for the DATA step to read in.
    /// </summary>
    public class DatasetConverter
    {
        #region private fields
        string _server, _libref, _member;
        OleDbConnection _connection = null;
        #endregion

        #region Properties
        /// <summary>
        /// Number of columns read from the data set
        /// </summary>
        public int ColumnCount { get; private set; }
        /// <summary>
        /// Number of rows read from the data set.
        /// </summary>
        public int RowCount { get; private set; }
        #endregion

        #region Constructor
        /// <summary>
        /// Creates a DatasetConverter.
        /// </summary>
        /// <param name="server">The name of the SAS server that hosts the source data.</param>
        /// <param name="libref">The SAS syntax name of the library where the data resides.</param>
        /// <param name="member">The member name of the data set.</param>
        public DatasetConverter(string server, string libref, string member)
        {
            _server = server;
            _libref = libref;
            _member = member;
            ColumnCount = -1;
            RowCount = -1;
        }
        #endregion

        #region routines to generate SAS program
        /// <summary>
        /// Generate a SAS DATA step program from the active data set.
        /// </summary>
        /// <remarks>
        /// <b>NOTE:</b>The DATALINES content is built up entirely in memory.
        /// If you need this to work with very large data sets, this would need
        /// to be refactored to emit the contents to a file-based stream as you
        /// build it out. 
        /// <para>
        /// Using StringBuilder along the way prevents large strings from being copied around
        /// and duplicated, but eventually there is going to 
        /// be a large block of text represented in the process memory.
        /// </para>
        /// </remarks>
        /// <param name="output">The output location of the SAS data set that the program will create.</param>
        /// <returns>The complete SAS DATA step program, complete with DATALINES values.</returns>
        public string GetCompleteSasProgram(string output)
        {
            System.Text.StringBuilder sb = new StringBuilder();

            SasServer sasServer = new SasServer(_server);

            using (_connection = sasServer.GetOleDbConnection())
            {
                try
                {
                    _connection.Open();
                    char[] types;

                    sb.Append(GetDataStepDefinition(output, out types));

                    string valuesdata = string.Format("{0}.{1}", _libref, _member);

                    OleDbCommand command = new OleDbCommand(string.Format("select * from {0}", valuesdata), _connection);
                    using (OleDbDataReader dataReader = command.ExecuteReader())
                    {
                        int fields = dataReader.FieldCount;
                        ColumnCount = fields;
                        RowCount = 0;
                        while (dataReader.Read())
                        {
                            StringBuilder line = new StringBuilder();
                            for (int i = 0; i < fields; i++)
                            {
                                string val = dataReader[i].ToString();
                                if (types[i] == 'C')
                                {
                                    // quote and escape strange chars for character values
                                    val = string.Format("\"{0}\"", FixUpDatalinesString(val));
                                }
                                line.AppendFormat("{0}{1}", val, i < fields - 1 ? "," : "");
                            }     
                            // increment the row count
                            RowCount++;
                            // add to our main program
                            sb.AppendFormat("{0}\n", line.ToString());
                        }
                    }

                    _connection.Close();
                }
                catch (Exception ex)
                {
                    sb.AppendFormat("/*Error trying to read input data: {0} */ \n", ex.Message);
                }
                finally
                {
                    // close out the datalines and add RUN
                    sb.Append(";;;;\nRUN;\n");
                }
            }
            return sb.ToString();
        }

        /// <summary>
        /// A special version of GetDataStepDefinition that returns just the DATA step
        /// header.  It's used for purposes of previewing the data set structure.
        /// </summary>
        /// <param name="output">The name of the output data set.</param>
        /// <returns>The DATA step definition.</returns>
        public string GetJustDataStepDefinition(string output)
        {
            // open the connection to data
            SasServer sasServer = new SasServer(_server);
            // the "using" statement will ensure that the 
            // connection is closed and disposed of when the 
            // scope of the "using" block ends.
            using (_connection = sasServer.GetOleDbConnection())
            {
                try
                {
                    _connection.Open();
                    char[] types;
                    string code = GetDataStepDefinition(output, out types);
                    // add the "values omitted" message to the end of this preview code
                    code = code + Environment.NewLine + Messages.ValuesOmitted;
                    // add the 4 semicolons to terminate the DATALINES4; statement
                    code = code + Environment.NewLine + ";;;;";

                    _connection.Close();
                    return code;
                }
                catch (Exception ex)
                {
                    return string.Format("/*Error trying to read input data: {0} */ \n", ex.Message);
                }
            }
        }

        /// <summary>
        /// Used to discover the attributes of the data set
        /// to build the DATA step definition.
        /// </summary>
        /// <param name="output">The name of the output data set.</param>
        /// <param name="types">An array of variable types, one per column.</param>
        /// <returns>The DATA step definition.</returns>
        private string GetDataStepDefinition(string output, out char[] types)
        {
            StringBuilder sb = new StringBuilder();
            string dsOptions = GetDatasetOptions();

            sb.AppendFormat(Messages.ProgramHeader, _libref + "." + _member);
            sb.AppendLine();
            // Create the DATA statement with output name and DS options
            sb.AppendFormat("DATA {0}{1};\n", output, dsOptions);

            string input;
            sb.Append(GetColumnInformation(out types, out input));

            // finish the INPUT statement;
            input += "\t;\n";

            sb.Append("INFILE DATALINES DSD;\n");

            // add the INPUT statement
            sb.Append(input);

            // now to append the datalines
            sb.Append("DATALINES4;\n");
            return sb.ToString();
        }

        /// <summary>
        /// Worker routine to discover the column information from the data set.
        /// </summary>
        /// <param name="types">Array of data types (C or N)</param>
        /// <param name="input">The INPUT statement that will be built.</param>
        /// <returns>The collection of ATTRIB statements that define the columns.</returns>
        private string GetColumnInformation(out char[] types, out string input)
        {
            // use the SASHELP.VCOLUMN view to find data set column attributes
            string selectclause = "select name, type, length, format, informat, label, sortedby from sashelp.vcolumn where libname='" + _libref + "' and memname='" + _member + "' ORDER BY varnum ASC";
            OleDbCommand command = new OleDbCommand(selectclause, _connection);
            StringBuilder typesSB = new StringBuilder(256);
            StringBuilder inputSB = new StringBuilder("INPUT \n");
            StringBuilder sb = new StringBuilder();
            using (OleDbDataReader dataReader = command.ExecuteReader())
            {
                types = null;
                while (dataReader.Read())
                {
                    inputSB.AppendFormat("\t{0}\n", dataReader["name"].ToString());

                    sb.AppendFormat("\tattrib {0} \n\t\tlength={1}{2}{3}{4}{5};\n",
                        dataReader["name"].ToString(),
                        dataReader["type"].ToString().ToUpper().StartsWith("C") ? "$" : "", // is it a char or num
                        dataReader["length"].ToString(),
                        // add FORMAT and INFORMAT attributes only if defined
                        dataReader["format"].ToString().Trim().Length == 0 ? "" : string.Format("\n\t\tformat={0} ", dataReader["format"].ToString()),
                        dataReader["informat"].ToString().Trim().Length == 0 ? "" : string.Format("\n\t\tinformat={0} ", dataReader["informat"].ToString()),
                        // same with LABEL and also escape quote chars if needed
                        dataReader["label"].ToString().Trim().Length == 0 ? "" : string.Format("\n\t\tlabel='{0}' ", dataReader["label"].ToString().Replace("'", "''"))
                        );

                    typesSB.Append(dataReader["type"].ToString().ToUpper()[0]); // append a C or N to our "type array"
                }
            }

            // convert types string builder to char array
            types = typesSB.ToString().ToCharArray();
            input = inputSB.ToString();
            return sb.ToString();
        }

        /// <summary>
        /// Retrieve the data set options.
        /// Currently this checks only for the data set label (LABEL= option).
        /// </summary>
        /// <returns>string with options, including parentheses, if needed.</returns>
        private string GetDatasetOptions()
        {
            string dsOptions = "";
            string selectclause = string.Format("select memlabel from sashelp.vtable where libname='{0}' and memname='{1}'", _libref, _member);
            OleDbCommand command = new OleDbCommand(selectclause, _connection);
            using (OleDbDataReader dataReader = command.ExecuteReader())
            {
                if (dataReader.Read())
                {
                    string dsLabel = "";
                    dsLabel = dataReader["memlabel"].ToString();
                    // use SASHELP.VTABLE to get data set options, such as DS label
                    if (dsLabel.Trim().Length > 0)
                        dsOptions = string.Format("(label=\"{0}\")", FixUpDatalinesString(dsLabel));
                }
            }
            return dsOptions;
        }

        #endregion

        #region utility
        /// <summary>
        /// Remove/escape characters that might interfere with the DATALINES records
        /// </summary>
        /// <param name="s">character value</param>
        /// <returns>character value appropriate for datalines.  Might be unchanged.</returns>
        private static string FixUpDatalinesString(string s)
        {
            if (s == null) return string.Empty;
            // escape double quotes
            s = s.Replace("\"", "\"\"");
            // remove carriage returns / line feeds
            s = s.Replace("\n", " ");
            s = s.Replace("\r", " ");
            // and make sure we don't terminate the datalines4 early
            s = s.Replace(";;;;", ";");
            // make sure it's not too long
            if (s.Length > 32767)
                s = s.Substring(32767);
            return s;
        }

        /// <summary>
        /// In the special case where we have a local SAS data set file (sas7bdat),
        /// and a local SAS server, we have to make sure that there is a library
        /// assigned.  The DatasetConverter class can read data only from a 
        /// data source that is accessed via a SAS library (LIBNAME.MEMBER).
        /// </summary>
        /// <param name="sd"></param>
        internal static void AssignLocalLibraryIfNeeded(ISASTaskConsumer3 consumer, SasTask taskModel)
        {
            SAS.Tasks.Toolkit.Data.SasData sd = new SAS.Tasks.Toolkit.Data.SasData(consumer.ActiveData as ISASTaskData2);
            // get a SasServer object so we can see if it's the "Local" server
            SAS.Tasks.Toolkit.SasServer server = new SAS.Tasks.Toolkit.SasServer(sd.Server);
            // local server with local file, so we have to assign library
            if (server.IsLocal)
            {
                // see if the data reference is a file path ("c:\data\myfile.sas7bdat")
                if (!string.IsNullOrEmpty(consumer.ActiveData.File) &&
                    consumer.ActiveData.Source == SourceType.SasDataset &&
                    consumer.ActiveData.File.Contains("\\"))
                {
                    string path = System.IO.Path.GetDirectoryName(consumer.ActiveData.File);
                    taskModel.SubmitSASProgramAndWait(string.Format("libname {0} \"{1}\";\r\n", sd.Libref, path));
                }
            }
        }
        #endregion
    }
}

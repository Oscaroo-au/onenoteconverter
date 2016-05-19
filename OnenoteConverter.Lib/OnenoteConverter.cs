using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using System.Text.RegularExpressions;
using System.IO;

namespace OnenoteConverter.Lib
{
    /// <summary>
    /// This is the object that performs the actual conversion across
    /// different folders, etc. This is in effect, the model.
    /// </summary>
    public class OnenoteConverter
    {
        /// <summary>
        /// The source path where there are many notebooks in it.
        /// </summary>
        public Uri SourcePath { get; set; }

        /// <summary>
        /// An output path where the converted notebooks shall be placed into.
        /// </summary>
        public Uri DestinationPath { get; set; }

        /// <summary>
        /// A regex expression that will be used to select which of the 
        /// notebooks in the SourcePath will be used. If null or empty then
        /// the filter will correspond to *all* notebooks in the folder.
        /// </summary>
        public string FilterExpression { get; set; } = "";
        protected Regex m_FilterRegex = null;
        private string input_nb_uri;


        /// <summary>
        /// Checks that all args are valid and that the regex too.
        /// </summary>
        private void CheckArgs()
        {
            // ensure source and dest dir works.
            CheckUri(SourcePath, "Source path not set");
            CheckUri(DestinationPath, "Dest path not set");

            // deal with regex.
            var pattern = FilterExpression;
            if (string.IsNullOrEmpty(pattern))
                pattern = ".*";
            m_FilterRegex = new Regex(pattern);
        }

        /// <summary>
        /// Checks whether the Uri's are valid.
        /// </summary>
        /// <param name="tocheck"></param>
        /// <param name="onerror_msg"></param>
        private void CheckUri(Uri tocheck, string onerror_msg)
        {
            var dir_path = tocheck?.LocalPath ?? onerror_msg;
            var dir_info = new DirectoryInfo(dir_path);
            if (!dir_info.Exists)
                throw new FileNotFoundException("The path '{0}' is not accessible", dir_path);
        }

        /// <summary>
        /// Returns a list of Uris that correspond to the notebooks that will
        /// be processed according to the filter expression.
        /// </summary>
        /// <returns></returns>
        public async Task<List<Uri>> GetURIsToProcessAsync()
        {
            CheckArgs();

            List<Uri> toret = null;

            // Enumerate the directories
            await Task.Run(() =>
            {
                var all_dirs = Directory.EnumerateDirectories(SourcePath.LocalPath);
                var dirs_to_output = new List<Uri>();

                foreach (var dir in all_dirs)
                {
                    var dir_info = new DirectoryInfo(dir);
                    var dir_name = dir_info.Name;

                    if (m_FilterRegex.IsMatch(dir_name))
                        dirs_to_output.Add(new Uri(dir_info.FullName));
                }

                toret = dirs_to_output;
            });

            return toret;
        }


        /// <summary>
        /// Begins the conversion process.
        /// </summary>
        /// <returns></returns>
        public async Task ConvertAsync(ProgressReporter report)
        {
            var toprocess = await GetURIsToProcessAsync();

            report.IncreaseMaxStep(toprocess.Count * 2);
            foreach (var input_nb_uri in toprocess)
            {
                var input_nb_info = new DirectoryInfo(input_nb_uri.LocalPath);
                report.ReportProgress(string.Format("Starting notebook '{0}'", input_nb_info.Name));

                var output_uri = new Uri(Path.Combine(DestinationPath.LocalPath, input_nb_info.Name));

                var nb = new Notebook(OneNote.Instance, input_nb_uri);
                try
                {
                    report.PushIndent();
                    await nb.Publish2013Async(output_uri, report);
                    report.PopIndent();
                }
                catch (PagesAndSectionsNotReadyException)
                {
                    report.IncreaseMaxStep(1);
                    report.ReportProgress(string.Format("   Error notebook '{0}'", input_nb_info.Name));
                }

                nb.Close();
                report.ReportProgress(string.Format("Finished notebook '{0}'", input_nb_info.Name));
            }

        }
        /// <summary>
        /// Begins the conversion process.
        /// </summary>
        /// <returns></returns>
        public async Task ConvertAsync()
        {
            await ConvertAsync(new ProgressReporter());
        }
    }
}

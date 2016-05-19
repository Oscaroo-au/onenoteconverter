using System;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.OneNote;
using System.Xml;
using System.IO;
using System.Diagnostics;
using System.Xml.Linq;

namespace OnenoteConverter.Lib
{
    public class Notebook
    {
        #region "fields"
        /// <summary>
        /// The instance of onenote with which this notebook object was created.
        /// </summary>
        private Application2 m_onenoteInstance;

        /// <summary>
        /// The URI path of the notebook
        /// </summary>
        public Uri Path { get; protected set; } = null;

        /// <summary>
        /// The Onenote ID of the notebook.
        /// </summary>
        public string ID { get; protected set; } = null;
        /// <summary>
        /// The sections in this notebook.
        /// </summary>
        public OrderedDictionary Sections { get; protected set; } = new OrderedDictionary();
        #endregion

        /// <summary>
        /// Prepares to open a new notebook. It will actually be opened 
        /// when a later function is called.
        /// </summary>
        /// <param name="instance"></param>
        /// <param name="path">A folder that contains a notebook param>
        public Notebook(Application2 instance, Uri path)
        {
            this.m_onenoteInstance = instance;
            this.Path = path;
        }

        /// <summary>
        /// Navigates to the notebook.
        /// </summary>
        public void Show()
        {
            LoadInitialHierarchy();
            m_onenoteInstance.NavigateTo(ID);
        }

        /// <summary>
        /// Publishes it to the given 'destination' with the 2007 format.
        /// </summary>
        /// <param name="destination"></param>
        public async Task Publish2007Async(Uri destination, ProgressReporter progress = null)
        {
            progress = progress ?? new ProgressReporter();

            await PublishAsync(destination, progress, PublishFormat.pfOneNote2007);
        }

        /// <summary>
        /// Publishes it to the given 'destination' with the new format.
        /// </summary>
        /// <param name="destination"></param>
        public async Task Publish2013Async(Uri destination, ProgressReporter progress = null)
        {
            progress = progress ?? new ProgressReporter();

            await PublishAsync(destination, progress, PublishFormat.pfOneNote);
        }

        private async Task PublishAsync(Uri destination, ProgressReporter progress, PublishFormat fmt)
        {
            progress.IncreaseMaxStep(3);
            progress.ReportProgress("Loading notebook");
            LoadInitialHierarchy();
            progress.ReportProgress( "Loading sections and pages");
            await LoadSectionsPagesHierarchyAsync();

            progress.ReportProgress("Starting exporting");
            progress.IncreaseMaxStep(Sections.Count);

            // need to make folder, and output each non-empty section at a time.
            System.IO.Directory.CreateDirectory(destination.LocalPath);
            foreach (Section section_obj in Sections.Values)
            {
                if (section_obj.HasPages)
                {
                    progress.ReportProgress(string.Format("Starting converting section '{0}'", section_obj.Name));
                    var section_dest = System.IO.Path.Combine(destination.LocalPath, section_obj.Name + ".one");
                    m_onenoteInstance.Publish(section_obj.ID, section_dest, fmt);
                }
                else
                {
                    progress.ReportProgress(string.Format("Skipping empty section '{0}'", section_obj.Name));
                }
            }

            //need to open the destination place as a notebook and then see
            //the sections it has. Then need to update their sections to be
            //in the same order as ours.
            //Note. This only works if the 'fmt' is not 2007.
            if (fmt == PublishFormat.pfOneNote)
            {
                progress.IncreaseMaxStep(1);
                progress.ReportProgress("Ordering sections");

                await CreateEmptySectionsAsync(destination);
            }
        }

        /// <summary>
        /// Creates the empty missing sections in the destination notebook.
        /// </summary>
        /// <param name="destination"></param>
        private async Task CreateEmptySectionsAsync(Uri destination)
        {
            Notebook dest = new Notebook(m_onenoteInstance, destination);
            dest.LoadInitialHierarchy();
            await dest.LoadSectionsPagesHierarchyAsync();

            // we clone the sections of the source to make sections for a
            // new-dest. Then, we update the sections of the new-dest to have 
            // the ids that the old-dest did ones. For those new-dest without 
            // an old-dest equivalent, their IDs is removed so they look like 
            // new sections to onenote.
            var new_dest = new OrderedDictionary();
            foreach (Section item in Sections.Values)
            {
                var item_copy = item.Clone() as Section;
                item_copy.ID = "";
                new_dest.Add(item_copy.Name, item_copy);
            }
            foreach (Section item in dest.Sections.Values)
            {
                if (!new_dest.Contains(item.Name))
                    throw new InvalidDataException();

                var section_new_dest = new_dest[item.Name] as Section;
                section_new_dest.ID = item.ID;
            }
            dest.Sections = new_dest;

            dest.UpdateSectionsHierarchy();
            dest.Close();
        }


        /// <summary>
        /// Loads the initial hierarchy.
        /// </summary>
        private void LoadInitialHierarchy()
        {
            if (ID != null) return;

            var tmp_notebook_id = "";

            m_onenoteInstance.OpenHierarchy(Path.LocalPath, "", out tmp_notebook_id, CreateFileType.cftNone);
            ID = tmp_notebook_id;
        }
        /// <summary>
        /// Updates the sections and pages for the notebook from the onenote.
        /// While discarding previous values.
        /// 
        /// May raise 'PagesAndSectionsNotReadyException' if the sections are
        /// not ready.
        /// </summary>
        private void LoadSectionsPagesHierarchy()
        {
            LoadInitialHierarchy();

            string tmp_xml;
            m_onenoteInstance.GetHierarchy(ID, HierarchyScope.hsPages, out tmp_xml, XMLSchema.xs2013);

            var doc = new XmlDocument();
            doc.LoadXml(tmp_xml);

            var section_nodes = doc.SelectNodes("one:Notebook/one:Section", OneNote.NamespaceManager);
            var sections = new OrderedDictionary();
            foreach (XmlNode xml_section in section_nodes)
            {
                var section_obj = new Section(xml_section);
                sections[section_obj.Name] = section_obj;
            }

            this.Sections = sections;

            var isLoading = doc.SelectNodes("one:Notebook/one:Section[@areAllPagesAvailable='false']", OneNote.NamespaceManager);

            //if there are any that are still loading then we have to exit 
            if (isLoading.Count > 0)
            {
                throw new PagesAndSectionsNotReadyException();
            }
            CheckSectionsWithFileSystem();
        }

        /// <summary>
        /// Loads the sections and pages and waits appropriately for it to
        /// all be loaded.
        /// </summary>
        /// <returns></returns>
        public async Task  LoadSectionsPagesHierarchyAsync()
        {
            var loader = new PagesAndSectionsNotReadyRetrier();
            loader.TheFunction = new Action(LoadSectionsPagesHierarchy);
            loader.TimeWait = TimeSpan.FromSeconds(3);
            loader.TimesToWait = 20;

            await loader.ExecuteAsync();
        }

        /// <summary>
        /// Checks if the loaded sections are the same as avail in the file
        /// system. If not, then we haven't loaded them all yet and we have
        /// to retry.
        /// </summary>
        private void CheckSectionsWithFileSystem()
        {
            var section_files = System.IO.Directory.GetFiles(this.Path.LocalPath);
            foreach (var file in section_files)
            {
                var fi = new FileInfo(file);
                var fname = fi.NameNoExt();
                if (fi.Extension != ".one") continue;
                if (!this.Sections.Contains(fname))
                {
                    throw new PagesAndSectionsNotReadyException();
                }

            }
        }


        /// <summary>
        /// Updates the onenote hierarchy to have the sections that the section
        /// field contains, in that order. Generates appropriate XML for the
        /// task and submits it to onenote.
        /// </summary>
        private void UpdateSectionsHierarchy()
        {
            XNamespace ns = OneNote.XMLSchema;
            XElement e_notebook = new XElement(ns + "Notebook",
                new XAttribute(XNamespace.Xmlns + "one", ns),
                new XAttribute("ID", this.ID)
                );

            //the first section is added as the child of the notebook.
            //later sections are added after the previous section. This Action
            //is used to move the add-fns around as necessary to achieve this.
            Action<XElement> add_it = new Action<XElement>((a) => { e_notebook.Add(a); });

            foreach (Section section in Sections.Values)
            {
                var e_section = new XElement(ns + "Section",
                        new XAttribute("name", section.Name)
                        );
                if (!string.IsNullOrEmpty(section.ID))
                    e_section.Add(new XAttribute("ID", section.ID));
                if (!string.IsNullOrEmpty(section.Color))
                    e_section.Add(new XAttribute("color", section.Color));

                add_it(e_section);
                add_it = new Action<XElement>((a) => { e_section.AddAfterSelf(a); });
            }

            var output = e_notebook.ToString();
            m_onenoteInstance.UpdateHierarchy(output, XMLSchema.xs2013);
        }


        /// <summary>
        /// Closes the notebook.
        /// Set force to prevent onenote from syncing the notebook back.
        /// Note: After closing this notebook becomes invalid!
        /// </summary>
        /// <param name="force"></param>
        public void Close(bool force = false)
        {
            m_onenoteInstance.CloseNotebook(ID, force);
            InvalidateInstance();
        }

        /// <summary>
        /// Invalidates this instance so no one can use it again.
        /// </summary>
        private void InvalidateInstance()
        {
            m_onenoteInstance = null;
            ID = null;
        }
    }
}

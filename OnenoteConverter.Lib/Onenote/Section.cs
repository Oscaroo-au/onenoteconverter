using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;

namespace OnenoteConverter.Lib
{
    /// <summary>
    /// Represents the sections in a notebook.
    /// </summary>
    public class Section : ICloneable
    {
        /// <summary>
        /// The Onenote ID of the section
        /// </summary>
        public string ID { get; set; } = null;

        /// <summary>
        /// The path to the section. Ie: the .one file.
        /// </summary>
        public Uri Path { get; protected set; } = null;

        /// <summary>
        /// The name of the section.
        /// </summary>
        public string Name { get; protected set; } = "";

        public string Color { get; protected set; } = "";

        /// <summary>
        /// Whether there are pages in this section.
        /// </summary>
        public bool HasPages { get; protected set; } = false;

        /// <summary>
        /// Constructs a section from the XMLNode
        /// </summary>
        /// <param name="xml"></param>
        public Section(XmlNode xml)
        {
            Name = xml.Attributes["name"]?.Value ?? ""; //may not have name
            ID = xml.Attributes["ID"].Value; // must have id
            Path = new Uri(xml.Attributes["path"]?.Value ?? ""); // may have path.
            Color = xml.Attributes["color"]?.Value ?? ""; // may have color

            HasPages = xml.SelectNodes("one:Page", OneNote.NamespaceManager).Count > 0;
        }

        /// <summary>
        /// Makes a clone of this object.
        /// </summary>
        /// <returns></returns>
        public object Clone()
        {
            return this.MemberwiseClone();
        }
    }
}
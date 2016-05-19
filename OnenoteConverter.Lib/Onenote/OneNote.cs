using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;
using Microsoft.Office.Interop.OneNote;

namespace OnenoteConverter.Lib
{
    /// <summary>
    /// Wraps the onenote functions.
    /// </summary>
    public class OneNote
    {
        /// <summary>
        /// Gets a onenote2013 instance.
        /// </summary>
        public static Application2 Instance { get { return s_Instance.Value; } }

        /// <summary>
        /// Gets a namespace manager instance for the XML namespace of onenote
        /// </summary>
        public static XmlNamespaceManager NamespaceManager {  get { return s_NamespaceManager.Value; } }

        protected static Lazy<XmlNamespaceManager> s_NamespaceManager = new Lazy<XmlNamespaceManager>(() =>
        {
           var nms = new XmlNamespaceManager(new NameTable());
           nms.AddNamespace("one", XMLSchema);

           return nms;
        });
        protected static Lazy<Application2> s_Instance = new Lazy<Application2>(() =>
        {
           return new Application2();
        });

        /// <summary>
        /// XML schema for onenote 2013
        /// </summary>
        public static readonly string XMLSchema = "http://schemas.microsoft.com/office/onenote/2013/onenote";
    }
}

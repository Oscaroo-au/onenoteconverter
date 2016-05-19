using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OnenoteConverter.Lib
{
    public static class Extensions
    {
        /// <summary>
        /// Gets the file name without the extension.
        /// </summary>
        /// <param name="fi"></param>
        /// <returns></returns>
        public static String NameNoExt(this FileInfo fi)
        {
            return fi.Name.Substring(0, fi.Name.Length - fi.Extension.Length);

        }
    }
}

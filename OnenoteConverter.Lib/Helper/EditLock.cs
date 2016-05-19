using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OnenoteConverter.Lib
{
    /// <summary>
    /// Thrown when editting is locked but someone tried to edit it anyway.
    /// </summary>
    public class EditIsLockedException : Exception { }

    /// <summary>
    /// Raises an exception when it is asked to allow an edit, if the edit
    /// has been locked.
    /// </summary>
    public class EditLock
    {
        private bool m_isLock = false;
        private object m_syncLock = new object();

        /// <summary>
        /// Sets the lock to unlocked.
        /// </summary>
        public void Unlock()
        {
            lock (m_syncLock)
            {
                m_isLock = false;
            }
        }

        /// <summary>
        /// Sets the state to locked.
        /// </summary>
        public void Lock()
        {
            lock (m_syncLock)
            {
                m_isLock = true;
            }
        }

        /// <summary>
        /// If edit is locked, an exception is raised.
        /// </summary>
        public void TryEdit()
        {
            lock (m_syncLock)
            {
                if (m_isLock)
                    throw new EditIsLockedException();
            }
        }
    }
}

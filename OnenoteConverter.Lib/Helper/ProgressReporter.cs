using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OnenoteConverter.Lib
{
    /// <summary>
    /// An exception raised when there is too much progress.
    /// </summary>
    public class ProgressOverrunException : Exception { }

    /// <summary>
    /// A progress event has stuff for reporting progress of an event.
    /// </summary>
    public class ProgressEventArgs : EventArgs
    {
        /// <summary>
        /// The step of progress.
        /// This is always increasing.
        /// </summary>
        public readonly int CurrentProgress;
        /// <summary>
        /// Progress is complete when CurrentProgress reaches FinalProgress..
        /// Final progress may change from one tick to another.
        /// </summary>
        public readonly int FinalProgress;
        /// <summary>
        /// A message that can be displayed during this progress.
        /// </summary>
        public readonly string Message;

        public ProgressEventArgs(int current, int max, string msg)
        {
            CurrentProgress = current;
            FinalProgress = max;
            Message = msg;
        }
    }

    /// <summary>
    /// An interface for reporting progress.
    /// NOT THREAD SAFE.
    /// </summary>
    public class ProgressReporter
    {
        int m_CurrentStep = 0;
        int m_MaxStep = 0;
        int m_IndentLevel = 0;

        /// <summary>
        /// An event raised when progress is reported.
        /// </summary>
        public event EventHandler<ProgressEventArgs> Progress;

        /// <summary>
        /// Emitted when finished
        /// </summary>
        public event EventHandler Finished;

        /// <summary>
        /// Called to report progress.
        /// </summary>
        /// <param name="message">A message to report with.</param>
        public void ReportProgress(string message)
        {
            m_CurrentStep += 1;
            if (m_CurrentStep > m_MaxStep)
                throw new ProgressOverrunException();
            if (m_MaxStep == 0)
                throw new InvalidOperationException("Need to call IncreaseMaxStep with non-zero at least once before reporting progress");

            for (int i = 0; i < m_IndentLevel; i++)
            {
                message = "  " + message;
            }
            Progress?.Invoke(this, new ProgressEventArgs(m_CurrentStep, m_MaxStep, message));
        }                                                           

        /// <summary>
        /// Resets the progress reporting
        /// </summary>
        public void Reset()
        {
            m_MaxStep = 0;
            m_CurrentStep = 0;
        }

        public void IncreaseMaxStep(int increase_by)
        {
            if (increase_by < 0)
                throw new ArgumentException("increase_by must be >= 0");

            m_MaxStep += increase_by;
        }

        /// <summary>
        /// Called to finish and report progress is all done.
        /// </summary>
        /// <param name="v"></param>
        internal void Complete(string message)
        {
            if (m_CurrentStep < m_MaxStep)
            {
                Progress?.Invoke(this, new ProgressEventArgs(m_MaxStep, m_MaxStep, message));
                m_CurrentStep = m_MaxStep;
            }

            Finished?.Invoke(this, EventArgs.Empty);
        }

        internal void PushIndent()
        {
            m_IndentLevel += 1;
        }

        internal void PopIndent()
        {
            m_IndentLevel -= 1;
        }
    }
}

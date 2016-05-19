using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace OnenoteConverter.Lib
{
    /// <summary>
    /// Base exception for this application.
    /// </summary>
    public class OnenoteConverterException : Exception
    {
    }

    /// <summary>
    /// Thrown when the pages and the sections cannot be relied upon as onenote
    /// reports they are not fully loaded.
    /// </summary>
    public class PagesAndSectionsNotReadyException : OnenoteConverterException
    {

    }

    /// <summary>
    /// Calls a fn and if the exception is raised retries after a configurable
    /// time.
    /// </summary>
    public class PagesAndSectionsNotReadyRetrier
    {
        /// <summary>
        /// The time to wait between retries. 
        /// </summary>
        public TimeSpan TimeWait { get; set; } = TimeSpan.FromSeconds(1);
        /// <summary>
        /// The number of times to try.
        /// Between 1 and INF
        /// </summary>
        public int TimesToWait { get; set; } = 3;
        /// <summary>
        /// The function to call to wait.
        /// </summary>
        public Action TheFunction { get; set; } = null;


        /// <summary>
        /// Checks whether the fields in this class are valid.
        /// </summary>
        private void CheckState()
        {
            if (TimesToWait <= 0)
                throw new ArgumentOutOfRangeException();
            if (TheFunction == null)
                throw new ArgumentNullException();
            if (TimeWait.TotalSeconds <= 0)
                throw new ArgumentOutOfRangeException();
        }

        /// <summary>
        /// Tries to execute the Action the given number of TimesToWait. Each 
        /// time waiting TimeWait period. If it doesn't work by then, the 
        /// PagesAndSectionsNotReadyException is raised.
        /// </summary>
        public async Task ExecuteAsync()
        {
            CheckState();
            bool success = false;
            for (int i = 0; i < TimesToWait && !success; i++)
            {
                try
                {
                    TheFunction();
                    success = true;
                }
                catch (PagesAndSectionsNotReadyException)
                {
                    await Task.Delay(TimeWait);
                }
            }

            if (!success)
                throw new PagesAndSectionsNotReadyException();
        }


    }
}

using System.Collections.Generic;
using System.ComponentModel;
using System.Runtime.CompilerServices;

namespace OnenoteConverter.Lib
{
    /// <summary>
    /// <summary>
    /// Sourced from: http://stackoverflow.com/questions/1315621/implementing-inotifypropertychanged-does-a-better-way-exist
    /// </summary>
    public class PropertyChangedHelper : INotifyPropertyChanged
    {
        // boiler-plate
        public event PropertyChangedEventHandler PropertyChanged;
        protected virtual void OnPropertyChanged([CallerMemberName] string property_name = null)
        {
            PropertyChangedEventHandler handler = PropertyChanged;
            if (handler != null) handler(this, new PropertyChangedEventArgs(property_name));
        }

        /// <summary>
        /// Set the field of the data n stuff.
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="field"></param>
        /// <param name="value"></param>
        /// <param name="property_name"></param>
        /// <returns></returns>
        protected bool SetField<T>(ref T field, T value, [CallerMemberName] string property_name = null)
        {
            if (EqualityComparer<T>.Default.Equals(field, value)) return false;
            field = value;
            OnPropertyChanged(property_name);
            return true;
        }
    }
}

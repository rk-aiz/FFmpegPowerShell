using System;
using System.ComponentModel;
using System.Globalization;
using System.Runtime.CompilerServices;
using System.Threading.Tasks;
using System.Windows.Data;
using System.Windows.Input;

namespace ViewModelHelper
{
    public abstract class ViewModelBase : INotifyPropertyChanged
    {
        public event PropertyChangedEventHandler PropertyChanged;

        protected virtual bool SetProperty<T>(ref T storage, T value, [CallerMemberName] string propertyName = null)
        {
            if (object.Equals(storage, value)) return false;

            storage = value;
            OnPropertyChanged(propertyName);
            return true;
        }

        protected virtual void OnPropertyChanged([CallerMemberName] string propertyName = null)
        {
            if (PropertyChanged != null)
                PropertyChanged.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }

    public class DelegateCommand : ICommand
    {
        public Action<object> ExecuteHandler { get; set; }
        public Func<object, bool> CanExecuteHandler { get; set; }
        public event EventHandler CanExecuteChanged;

        public bool CanExecute(object parameter)
        {
            if (CanExecuteHandler == null) { return true; }
            return CanExecuteHandler(parameter);
        }

        public void Execute(object parameter)
        {
            if (ExecuteHandler != null)
                ExecuteHandler.Invoke(parameter);
        }

        public void RaiseCanExecuteChanged()
        {
            if (CanExecuteChanged != null)
                CanExecuteChanged.Invoke(this, EventArgs.Empty);
        }

        public DelegateCommand(Action<object> execute, Func<object, bool> canExecute = null)
        {
            if (null == execute)
            {
                throw new ArgumentNullException("execute");
            }
            ExecuteHandler = execute;
            CanExecuteHandler = canExecute;
        }

        public DelegateCommand()
        {
        }

    }

    public class OneWayBinding : Binding
    {
        public OneWayBinding(string path) : base(path)
        {
            Mode = BindingMode.OneWay;
            UpdateSourceTrigger = UpdateSourceTrigger.PropertyChanged;
            ConverterCulture = CultureInfo.CurrentUICulture;
        }
    }

    public class TwoWayBinding : Binding
    {
        public TwoWayBinding(string path) : base(path)
        {
            Mode = BindingMode.TwoWay;
            UpdateSourceTrigger = UpdateSourceTrigger.PropertyChanged;
            ConverterCulture = CultureInfo.CurrentUICulture;
        }
    }

    public class AsyncCommand : ICommand
    {
        private readonly Func<object, Task> _executeAsync;
        private readonly Func<object, bool> _canExecute;
        private bool _isExecuting;

        public event EventHandler CanExecuteChanged;

        public AsyncCommand(Func<object, Task> executeAsync, Func<object, bool> canExecute = null)
        {
            if (null == executeAsync)
            {
                throw new ArgumentNullException("executeAsync");
            }
            _executeAsync = executeAsync;
            _canExecute = canExecute;
        }

        public bool CanExecute(object parameter)
        {
            return !_isExecuting && (_canExecute == null || _canExecute(parameter));
        }

        public async void Execute(object parameter)
        {
            if (!CanExecute(parameter)) return;

            _isExecuting = true;
            RaiseCanExecuteChanged();

            try
            {
                await _executeAsync(parameter);
            }
            catch (Exception ex)
            {
                // 例外をログに記録するなどして、アプリケーションのクラッシュを防ぎます。
                System.Diagnostics.Debug.WriteLine(String.Format("An exception occurred in AsyncCommand: {0}", ex.Message));
            }
            finally
            {
                _isExecuting = false;
                RaiseCanExecuteChanged();
            }
        }

        public void RaiseCanExecuteChanged()
        {
            if (CanExecuteChanged != null)
            {
                CanExecuteChanged.Invoke(this, EventArgs.Empty);
            }
        }
    }
}

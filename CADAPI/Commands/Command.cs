using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Input;

namespace CADAPI.Commands
{
    public class Command : ICommand
    {
        private readonly Action<object> _execute;
        private readonly Predicate<object> _canExecute;

        public Command(Command drawFooting)
        {
            DrawFooting = drawFooting;
        }

        public Command(Action<object> execute, Predicate<object> canExecute = null)
        {
            _execute = execute ?? throw new NullReferenceException(nameof(execute));
            _canExecute = canExecute ?? (_ => true);
        }

        public Command DrawFooting { get; }

        public event EventHandler CanExecuteChanged
        {
            add => CommandManager.RequerySuggested += value;
            remove => CommandManager.RequerySuggested -= value;
        }
        public bool CanExecute(object parameter = null) => _canExecute(parameter);

        public void Execute(object parameter) => _execute(parameter);

    }
}

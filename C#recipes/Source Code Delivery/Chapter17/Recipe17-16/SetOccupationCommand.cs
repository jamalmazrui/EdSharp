using System;
using System.ComponentModel;
using System.Windows.Input;

namespace Apress.VisualCSharpRecipes.Chapter17
{
    public class SetOccupationCommand : ICommand
    {
        private Person person;

        public SetOccupationCommand(Person person)
        {
            this.person = person;

            this.person.PropertyChanged +=
                new PropertyChangedEventHandler(person_PropertyChanged);
        }

        // Handle the PropertyChanged event of the person to raise the
        // CanExecuteChanged event
        private void person_PropertyChanged(
            object sender, PropertyChangedEventArgs e)
        {
            if (CanExecuteChanged != null)
            {
                CanExecuteChanged(this, EventArgs.Empty);
            }
        }

        #region ICommand Members

        /// The command can execute if the person has been added,
        /// which means its Status will be set, and if the occupation
        /// parameter is not null
        public bool CanExecute(object parameter)
        {
            if (!string.IsNullOrEmpty(parameter as string))
                if (!string.IsNullOrEmpty(person.Status))
                    return true;

            return false;
        }

        public event EventHandler CanExecuteChanged;

        /// When the command is executed, set the Occupation
        /// property of the person, and update the Status.
        public void Execute(object parameter)
        {
            // Get the occupation string from the command parameter
            person.Occupation = parameter.ToString();

            person.Status =
                string.Format("Added {0} {1}, {2}",
                              person.FirstName, person.LastName, person.Occupation);
        }
        #endregion
    }
}

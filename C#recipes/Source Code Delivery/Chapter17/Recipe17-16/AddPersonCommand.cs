using System;
using System.ComponentModel;
using System.Windows.Input;

namespace Apress.VisualCSharpRecipes.Chapter17
{
    public class AddPersonCommand : ICommand
    {
        private Person person;

        public AddPersonCommand(Person person)
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

        /// The command can execute if there are valid values
        /// for the person's FirstName, LastName, and Age properties
        /// and if it hasn't already been executed and had its
        /// Status property set.
        public bool CanExecute(object parameter)
        {
            if (!string.IsNullOrEmpty(person.FirstName))
                if (!string.IsNullOrEmpty(person.LastName))
                    if (person.Age > 0)
                        if (string.IsNullOrEmpty(person.Status))
                            return true;

            return false;
        }

        public event EventHandler CanExecuteChanged;

        /// When the command is executed, update the
        /// status property of the person.
        public void Execute(object parameter)
        {
            person.Status =
                string.Format("Added {0} {1}",
                              person.FirstName, person.LastName);
        }

        #endregion
    }
}

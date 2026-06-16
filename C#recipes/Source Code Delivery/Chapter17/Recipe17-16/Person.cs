using System;
using System.ComponentModel;
using System.Windows.Input;

namespace Apress.VisualCSharpRecipes.Chapter17
{
    public class Person : INotifyPropertyChanged
    {
        private string firstName;
        private int age;
        private string lastName;
        private string status;
        private string occupation;

        private AddPersonCommand addPersonCommand;
        private SetOccupationCommand setOccupationCommand;

        public string FirstName
        {
            get
            {
                return firstName;
            }
            set
            {
                if (firstName != value)
                {
                    firstName = value;
                    OnPropertyChanged("FirstName");
                }
            }
        }

        public string LastName
        {
            get
            {
                return lastName;
            }
            set
            {
                if (this.lastName != value)
                {
                    this.lastName = value;
                    OnPropertyChanged("LastName");
                }
            }
        }

        public int Age
        {
            get
            {
                return age;
            }
            set
            {
                if (this.age != value)
                {
                    this.age = value;
                    OnPropertyChanged("Age");
                }
            }
        }

        public string Status
        {
            get
            {
                return status;
            }
            set
            {
                if (this.status != value)
                {
                    this.status = value;
                    OnPropertyChanged("Status");
                }
            }
        }

        public string Occupation
        {
            get
            {
                return occupation;
            }
            set
            {
                if (this.occupation != value)
                {
                    this.occupation = value;
                    OnPropertyChanged("Occupation");
                }
            }
        }

        /// Gets an AddPersonCommand for data binding
        public AddPersonCommand Add
        {
            get
            {
                if (addPersonCommand == null)
                    addPersonCommand = new AddPersonCommand(this);

                return addPersonCommand;
            }
        }

        /// Gets a SetOccupationCommand for data binding
        public SetOccupationCommand SetOccupation
        {
            get
            {
                if (setOccupationCommand == null)
                    setOccupationCommand = new SetOccupationCommand(this);

                return setOccupationCommand;
            }
        }

        #region INotifyPropertyChanged Members

        /// Implement INotifyPropertyChanged to notify the binding
        /// targets when the values of properties change.
        public event PropertyChangedEventHandler PropertyChanged;

        private void OnPropertyChanged(string propertyName)
        {
            if (this.PropertyChanged != null)
            {
                this.PropertyChanged(
                    this, new PropertyChangedEventArgs(propertyName));
            }
        }

        #endregion
    }
}

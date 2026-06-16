using System.ComponentModel;

namespace Apress.VisualCSharpRecipes.Chapter17
{
    public class Person : INotifyPropertyChanged
    {
        private string firstName;
        private string lastName;
        private int age;
        private string occupation;

        // Each property calls the OnPropertyChanged method
        // when its value changed, and each property that 
        // affects the Person's Description, also calls the 
        // OnPropertyChanged method for the Description property.

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
                    OnPropertyChanged("Description");
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
                    OnPropertyChanged("Description");
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
                    OnPropertyChanged("Description");
                }
            }
        }

        public string Occupation
        {
            get { return occupation; }
            set
            {
                if (this.occupation != value)
                {
                    this.occupation = value;
                    OnPropertyChanged("Occupation");
                    OnPropertyChanged("Description");
                }
            }
        }

        // The Description property is read-only,
        // and is composed of the values of the
        // other properties.
        public string Description
        {
            get
            {
                return string.Format("{0} {1}, {2} ({3})",
                                     firstName, lastName, age, occupation);
            }
        }

        // The ToString returns the Description,
        // so that it is displayed by default when 
        // the Person object is a binding source.
        public override string ToString()
        {
            return Description;
        }

        #region INotifyPropertyChanged Members

        /// Implement INotifyPropertyChanged to notify the binding
        /// targets when the values of properties change.
        public event PropertyChangedEventHandler PropertyChanged;

        private void OnPropertyChanged(
            string propertyName)
        {
            if (this.PropertyChanged != null)
            {
                // Raise the PropertyChanged event
                this.PropertyChanged(
                    this,
                    new PropertyChangedEventArgs(
                        propertyName));
            }
        }

        #endregion
    }
}
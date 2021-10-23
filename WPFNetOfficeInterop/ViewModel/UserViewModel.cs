using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Input;
using System.ComponentModel;
using WPFNetOfficeInterop.Model;
using WPFNetOfficeInterop.Helpers;

namespace WPFNetOfficeInterop.ViewModel
{
    class UserViewModel
    {
        private IList<User> _UsersList;
        public UserViewModel()
        {
            _UsersList = new List<User>
            {
                new User{UserId=1,FirstName="Jamie",LastName="Bradley",City="Plymouth",Postcode="PL12PQ",Country="UK"},
                new User{UserId=2,FirstName="Clinton",LastName="Valdez",City="Exeter", Postcode="EX17UY", Country="UK"},
                new User{UserId=3,FirstName="Darryl",LastName="Perez",City="Birmingham", Postcode="BI26TS", Country="UK"},
                new User{UserId=4,FirstName="Leon",LastName="Aguilar",City="London", Postcode="SW13PW", Country="UK"},
                new User{UserId=5,FirstName="Alton",LastName="Adams",City="Sheffield", Postcode="SH17UH", Country="UK"},
                new User{UserId=6,FirstName="Guillermo",LastName="Griffin",City="Leeds", Postcode="LE48YH", Country="UK"},
                new User{UserId=7,FirstName="Blanche",LastName="Washington",City="Glasgow", Postcode="GL38UH", Country="UK"},
                new User{UserId=8,FirstName="Robin",LastName="Warren",City="Newcastle", Postcode="NE15TF", Country="UK"}            
            };
        }

        public IList<User> Users
        {
            get { return _UsersList; }
            set { _UsersList = value; }
        }

        private ICommand mUpdateCommand;
        private ICommand mExcelCommand;
        private ICommand mWordCommand;
        private ICommand mPowerpointCommand;


        public ICommand Update
        {
            get
            {
                if (mUpdateCommand == null)
                {
                    mUpdateCommand = new UpdateCommand();
                }
                return mUpdateCommand;
            }
            set
            {
                mUpdateCommand = value;
            }
        }

        public ICommand OutputExcel
        {
            get
            {
                if (mExcelCommand == null)
                {
                    mExcelCommand = new ExcelCommand(_UsersList);
                }
                return mExcelCommand;
            }
            set
            {
                mExcelCommand = value;
            }
        }


        public ICommand OutputWord
        {
            get
            {
                if (mWordCommand == null)
                {
                    mWordCommand = new WordCommand(_UsersList);
                }
                return mWordCommand;
            }
            set
            {
                mWordCommand = value;
            }
        }

        public ICommand OutputPowerpoint
        {
            get
            {
                if (mPowerpointCommand == null)
                {
                    mPowerpointCommand = new PowerPointCommand(_UsersList);
                }
                return mPowerpointCommand;
            }
            set
            {
                mPowerpointCommand = value;
            }
        }


        private class ExcelCommand : ICommand
        {
            #region ICommand Members
            private IList<User> _users;
            public event EventHandler CanExecuteChanged;

            public ExcelCommand(IList<User> users)
            {
                _users = users;
            }

            public void Execute(object parameter)
            {
                ExcelHelper.Export(_users);
            }

            public bool CanExecute(object parameter)
            {
                return true;
            }
            #endregion
        }


        private class WordCommand : ICommand
        {
            #region ICommand Members
            private IList<User> _users;
            public event EventHandler CanExecuteChanged;

            public WordCommand(IList<User> users)
            {
                _users = users;
            }

            public void Execute(object parameter)
            {
                WordHelper.Export(_users);
            }

            public bool CanExecute(object parameter)
            {
                return true;
            }
            #endregion
        }


        private class PowerPointCommand : ICommand
        {
            #region ICommand Members
            private IList<User> _users;
            public event EventHandler CanExecuteChanged;

            public PowerPointCommand(IList<User> users)
            {
                _users = users;
            }

            public void Execute(object parameter)
            {
                PowerPointHelper.Export(_users);
            }

            public bool CanExecute(object parameter)
            {
                return true;
            }
            #endregion
        }

        private class UpdateCommand : ICommand
        {
            #region ICommand Members
            public event EventHandler CanExecuteChanged;

            public bool CanExecute(object parameter)
            {
                return true;
            }
            
            public void Execute(object parameter)
            {
            }
            #endregion
        }
    }
}

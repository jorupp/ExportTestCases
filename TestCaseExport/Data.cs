using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text;
using System.Threading.Tasks;
using Microsoft.TeamFoundation.Client;
using Microsoft.TeamFoundation.TestManagement.Client;
using TestCaseExport.Annotations;

namespace TestCaseExport
{
    public class Data : INotifyPropertyChanged
    {
        public event PropertyChangedEventHandler PropertyChanged;

        [NotifyPropertyChangedInvocator]
        protected virtual void OnPropertyChanged([CallerMemberName] string propertyName = null)
        {
            var handler = PropertyChanged;
            if (handler != null) handler(this, new PropertyChangedEventArgs(propertyName));
        }

        public event EventHandler<bool> IsBusy;

        protected virtual void OnIsBusy(bool e)
        {
            var handler = IsBusy;
            if (handler != null) handler(this, e);
        }

        private ITestManagementTeamProject _selectedProject;
        private BindingList<ITestPlan> _testPlans = new BindingList<ITestPlan>();
        private ITestPlan _selectedTestPlan;
        private BindingList<SelectableTestSuite> _testSuites = new BindingList<SelectableTestSuite>();
        private SelectableTestSuite _selectedTestSuite;
        private bool _suiteIsSelected;

        public ITestManagementTeamProject SelectedProject
        {
            get { return _selectedProject; }
            set
            {
                if (Equals(value, _selectedProject)) return;
                _selectedProject = value;
                OnPropertyChanged();
                OnPropertyChanged("SelectedProjectName");

                // reload the list of test plans (on background thread)
                OnIsBusy(true);
                TestPlans.Clear();
                if (null != _selectedProject)
                {
                    Task.Run(() => _selectedProject.TestPlans.Query("select * from TestPlan").ToList()).ContinueWith(
                        data =>
                        {
                            data.Result.ForEach(TestPlans.Add);
                            OnIsBusy(false);
                        }, TaskScheduler.FromCurrentSynchronizationContext());
                }
            }
        }

        public string SelectedProjectName
        {
            get { return _selectedProject == null ? "" : _selectedProject.TeamProjectName; }
        }

        public BindingList<ITestPlan> TestPlans
        {
            get { return _testPlans; }
        }

        public ITestPlan SelectedTestPlan
        {
            get { return _selectedTestPlan; }
            set
            {
                if (Equals(value, _selectedTestPlan)) return;
                _selectedTestPlan = value;
                OnPropertyChanged();

                // reload the list of test suites (on background thread)
                OnIsBusy(true);
                TestSuites.Clear();
                if (null != _selectedTestPlan)
                {
                    Task.Run(() => _selectedTestPlan.RootSuite.Entries.Where(i => i.TestSuite != null).Select(i => new SelectableTestSuite(i.TestSuite)).ToList()).ContinueWith(
                        data =>
                        {
                            data.Result.ForEach(TestSuites.Add);
                            OnIsBusy(false);
                        }, TaskScheduler.FromCurrentSynchronizationContext());
                }
            }
        }

        public BindingList<SelectableTestSuite> TestSuites
        {
            get { return _testSuites; }
        }

        public SelectableTestSuite SelectedTestSuite
        {
            get { return _selectedTestSuite; }
            set
            {
                if (Equals(value, _selectedTestSuite)) return;
                _selectedTestSuite = value;
                OnPropertyChanged();
                OnPropertyChanged("SuiteIsSelected");
            }
        }

        public bool SuiteIsSelected
        {
            get { return null != _selectedTestSuite; }
        }

        public class SelectableTestSuite
        {
            public ITestSuiteBase TestSuite { get; private set; }
            public string Title { get; private set; }

            public SelectableTestSuite(ITestSuiteBase testSuite)
            {
                TestSuite = testSuite;
                Title = testSuite.Title;
            }
        }
    }
}

using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Windows.Threading;
using Microsoft.TeamFoundation.Client;
using Microsoft.TeamFoundation.TestManagement.Client;
using TestCaseExport.Annotations;
using TestCaseExport.Properties;

namespace TestCaseExport
{
    public class Data : INotifyPropertyChanged
    {
        private List<Task> _pendingTasks = new List<Task>();

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
        private string _exportFileName;

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
                if (null != _selectedProject)
                {
                    OnIsBusy(true);
                    Task task = null;
                    task = Task.Run(() => _selectedProject.TestPlans.Query("select * from TestPlan").ToList()).ContinueWith(
                        data =>
                        {
                            TestPlans.Clear();
                            data.Result.ForEach(TestPlans.Add);
                            SelectedTestPlan = TestPlans.FirstOrDefault();
                            _pendingTasks.Remove(task);
                            OnIsBusy(false);
                        }, TaskScheduler.FromCurrentSynchronizationContext());
                    _pendingTasks.Add(task);
                }
                else
                {
                    TestPlans.Clear();
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
                if (null != _selectedTestPlan)
                {
                    OnIsBusy(true);
                    Task task = null;
                    task = Task.Run(() => _selectedTestPlan.RootSuite.Entries.Where(i => i.TestSuite != null).Select(i => new SelectableTestSuite(i.TestSuite)).ToList()).ContinueWith(
                        data =>
                        {
                            TestSuites.Clear();
                            data.Result.ForEach(TestSuites.Add);
                            SelectedTestSuite = TestSuites.FirstOrDefault();
                            OnIsBusy(false);
                            _pendingTasks.Remove(task);
                        }, TaskScheduler.FromCurrentSynchronizationContext());
                    _pendingTasks.Add(task);
                }
                else
                {
                    TestSuites.Clear();
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

        public string ExportFileName
        {
            get { return _exportFileName; }
            set
            {
                if (value == _exportFileName) return;
                _exportFileName = value;
                OnPropertyChanged();
                OnPropertyChanged("SuiteIsSelected");
            }
        }

        public bool SuiteIsSelected
        {
            get { return null != _selectedTestSuite && !string.IsNullOrEmpty(ExportFileName); }
        }

        internal void SaveSettings(Settings settings)
        {
            settings.TfsUrl = SelectedProject.WitProject.Store.TeamProjectCollection.Uri.AbsoluteUri;
            settings.ProjectName = SelectedProject.TeamProjectName;
            settings.TestPlan = SelectedTestPlan.Name;
            settings.TestSuite = SelectedTestSuite.Title;
            settings.ExportFilename = ExportFileName;
            settings.Save();
        }

        internal void LoadFromSettings(Settings settings)
        {
            if (string.IsNullOrEmpty(settings.TfsUrl) || string.IsNullOrEmpty(settings.ProjectName))
                return;
            var tfs = new TfsTeamProjectCollection(new Uri(settings.TfsUrl));
            SelectedProject = tfs.GetService<ITestManagementService>().GetTeamProject(settings.ProjectName);

            while(_pendingTasks.Count > 0)
                Application.DoEvents();

            if (string.IsNullOrEmpty(settings.TestPlan))
                return;

            SelectedTestPlan = TestPlans.SingleOrDefault(i => i.Name == settings.TestPlan);

            while (_pendingTasks.Count > 0)
                Application.DoEvents();

            if (string.IsNullOrEmpty(settings.TestSuite))
                return;

            SelectedTestSuite = TestSuites.SingleOrDefault(i => i.Title == settings.TestSuite);
            ExportFileName = settings.ExportFilename;
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

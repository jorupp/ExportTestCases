using System;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using Microsoft.TeamFoundation.Client;
using Microsoft.TeamFoundation.TestManagement.Client;

namespace TestCaseExport
{
    public partial class FrmMain : Form
    {
        private Data _data = new Data();
        
        private delegate void Execute();

        public FrmMain()
        {
            InitializeComponent();

            bsData.DataSource = _data;
            _data.IsBusy += (sender, isBusy) =>
            {
                this.UseWaitCursor = isBusy;
            };
        }

        private void btnTeamProject_Click(object sender, EventArgs e)
        {
            //Displaying the Team Project selection dialog to select the desired team project.
            var tpp = new TeamProjectPicker(TeamProjectPickerMode.SingleProject, false);
            tpp.ShowDialog();

            //Following actions will be executed only if a team project is selected in the the opened dialog.
            if (tpp.SelectedTeamProjectCollection != null)
            {
                var tfs = tpp.SelectedTeamProjectCollection;
                var tstSvc = (ITestManagementService)tfs.GetService(typeof(ITestManagementService));
                var teamProject = tstSvc.GetTeamProject(tpp.SelectedProjects[0].Name);

                _data.SelectedProject = teamProject;
            }
        }

        private void btnFolderBrowse_Click(object sender, EventArgs e)
        {
            if (saveFileDialog.ShowDialog() == DialogResult.OK)
            {
                _data.ExportFileName = saveFileDialog.FileName;
            }
        }

        private void btnExport_Click(object sender, EventArgs e)
        {
            this.Cursor = Cursors.WaitCursor;
            this.Enabled = false;

            try
            {
                var filename = _data.ExportFileName;
                new Exporter().Export(filename, _data.SelectedTestSuite.TestSuite);
                Process.Start(filename);

                this.Cursor = Cursors.Default;
                this.Enabled = true;
                MessageBox.Show("Test Cases exported successfully to specified file.", "Success", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Failed to export:" + Environment.NewLine + ex, "Failed", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                this.Cursor = Cursors.Default;
                this.Enabled = true;
            }
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void btnAbout_Click(object sender, EventArgs e)
        {
            new FrmAbout().ShowDialog();
        }
    }
}

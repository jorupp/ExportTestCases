using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Microsoft.TeamFoundation.Client;
using Microsoft.TeamFoundation.TestManagement.Client;
using Microsoft.TeamFoundation.WorkItemTracking.Client;
using Excel = Microsoft.Office.Interop.Excel;

namespace TestCaseExport
{
    public partial class FrmMain : Form
    {
        private TfsTeamProjectCollection _tfs;
        private ITestManagementTeamProject _teamProject;
        private ITestPlanCollection plans;
        private ITestSuiteEntryCollection testSuites;
        private ITestCaseCollection testCases;
        private ITestSuiteEntry suite;
        
        private delegate void Execute();

        int flag1 = 0, flag2 = 0, flag3 = 0, flag4 = 0, flag5 = 0;

        public FrmMain()
        {
            InitializeComponent();
        }

        private void btnTeamProject_Click(object sender, EventArgs e)
        {
            //Displaying the Team Project selection dialog to select the desired team project.
            TeamProjectPicker tpp = new TeamProjectPicker(TeamProjectPickerMode.SingleProject, false);
            tpp.ShowDialog();

            //Following actions will be executed only if a team project is selected in the the opened dialog.
            if (tpp.SelectedTeamProjectCollection != null)
            {
                this._tfs = tpp.SelectedTeamProjectCollection;
                ITestManagementService test_service = (ITestManagementService)_tfs.GetService(typeof(ITestManagementService));
                this._teamProject = test_service.GetTeamProject(tpp.SelectedProjects[0].Name);

                //Populating the text field Team Project name (txtTeamProject) with the name of the selected team project.
                txtTeamProject.Text = tpp.SelectedProjects[0].Name;
                flag1 = 1;

                //Call to method "Get_TestPlans" to get the test plans in the selected team project
                Get_TestPlans(_teamProject);
            }

        }

        private void Get_TestPlans(ITestManagementTeamProject teamProject)
        {
            //Getting all the test plans in the collection "plans" using the query.
            this.plans = teamProject.TestPlans.Query("Select * From TestPlan");
            comBoxTestPlan.Items.Clear();
            flag2 = 0;
            comBoxTestSuite.Items.Clear();
            flag3 = 0;
            foreach (ITestPlan plan in plans)
            {
                //Populating the plan selection dropdown list with the name of Test Plans in the selected team project.
                comBoxTestPlan.Items.Add(plan.Name);
            }
        }

        private void Get_TestSuites(ITestSuiteEntryCollection Suites)
        {
            
            foreach (ITestSuiteEntry suite_entry in Suites)
            {
                this.suite = suite_entry;
                IStaticTestSuite newSuite = suite_entry.TestSuite as IStaticTestSuite;
                comBoxTestSuite.Items.Add(newSuite.Title);
            }
        }

        private void Get_TestCases(IStaticTestSuite testSuite)
        {
            this.testCases = testSuite.AllTestCases;
        }


        //Following method is invoked whenever a Test Plan is selected in the dropdown list.
        //Acording to the selected Test Plan in the dropdown list the Test Suites present in the selected Test Plan are populated in the Test Suite selection dropdown.
        private void comBoxTestPlan_SelectedIndexChanged(object sender, EventArgs e)
        {
            this.Cursor = Cursors.WaitCursor;
            comBoxTestSuite.Items.Clear();
            int i = -1;
            if (comBoxTestPlan.SelectedIndex >= 0)
            {
                i = comBoxTestPlan.SelectedIndex;
                this.testSuites = plans[i].RootSuite.Entries;
                Get_TestSuites(testSuites);
                this.Cursor = Cursors.Arrow;
                flag2 = 1;
            }
        }

        private void btnFolderBrowse_Click(object sender, EventArgs e)
        {
            folderBrowserDialog.ShowDialog();
            txtSaveFolder.Text = folderBrowserDialog.SelectedPath;
            if (txtSaveFolder.Text != null || txtSaveFolder.Text != "")
            {
                flag4 = 1;
            }
        }

        private void comBoxTestSuite_SelectedIndexChanged(object sender, EventArgs e)
        {
            int j = -1;
            if (comBoxTestPlan.SelectedIndex >= 0)
            {
                j = comBoxTestSuite.SelectedIndex;
                this.suite = testSuites[j].TestSuite.TestSuiteEntry;
                IStaticTestSuite suite1 = suite.TestSuite as IStaticTestSuite;
                Get_TestCases(suite1);
                flag3 = 1;
            }
        }

        private void btnExport_Click(object sender, EventArgs e)
        {
            if (txtFileName.Text != null || txtFileName.Text != "")
            {
                flag5 = 1;
            }
            if (flag1 == 1 && flag2 == 1 && flag3 == 1 && flag4 == 1 && flag5 == 1)
            {
                this.Cursor = Cursors.WaitCursor;
                btnExport.Enabled = false;
                btnCancel.Enabled = false;
                btnTeamProject.Enabled = false;
                btnFolderBrowse.Enabled = false;
                comBoxTestPlan.Enabled = false;
                comBoxTestSuite.Enabled = false;

                Excel.Application xlApp;
                Excel.Workbook xlWorkBook;
                Excel.Worksheet xlWorkSheet;
                object misValue = System.Reflection.Missing.Value;
                Excel.Range chartRange;

                xlApp = new Excel.Application();
                xlWorkBook = xlApp.Workbooks.Add(misValue);
                xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);
                xlWorkSheet.Name = "Test Script";

                xlWorkSheet.Cells[1, 1] = "Test Condition";
                xlWorkSheet.Cells[1, 2] = "Action/Description";
                xlWorkSheet.Cells[1, 3] = "Input Data";
                xlWorkSheet.Cells[1, 4] = "Expected Result";
                xlWorkSheet.Cells[1, 5] = "Actual Result";
                xlWorkSheet.Cells[1, 6] = "Pass/Fail";
                xlWorkSheet.Cells[1, 7] = "Comments";

                (xlWorkSheet.Columns["A", Type.Missing]).ColumnWidth = 15;
                (xlWorkSheet.Columns["B", Type.Missing]).ColumnWidth = 50;
                (xlWorkSheet.Columns["C", Type.Missing]).ColumnWidth = 15;
                (xlWorkSheet.Columns["D", Type.Missing]).ColumnWidth = 50;
                (xlWorkSheet.Columns["E", Type.Missing]).ColumnWidth = 50;
                (xlWorkSheet.Columns["F", Type.Missing]).ColumnWidth = 15;
                (xlWorkSheet.Columns["G", Type.Missing]).ColumnWidth = 20;


                int row = 2;
                int col = 1;
                string upperBound = "a";
                string lowerBound = "a";


                foreach (ITestCase testCase in testCases)
                {
                    upperBound = "a";
                    lowerBound = "a";
                    col = 1;
                    //xlWorkSheet.Cells[row, col] = testCase.Title;

                    upperBound += row;
                    TestActionCollection testActions = testCase.Actions;
                    foreach (ITestStep testStep in testActions)
                    {
                        col = 2;
                        xlWorkSheet.Cells[row, 2] = testStep.Title.ToString();
                        xlWorkSheet.Cells[row, 4] = testStep.ExpectedResult.ToString();
                        row++;
                    }
                    lowerBound += (row - 1);
                    xlWorkSheet.get_Range(upperBound, lowerBound).Merge(false);

                    chartRange = xlWorkSheet.get_Range(upperBound, lowerBound);
                    chartRange.FormulaR1C1 = testCase.Title.ToString();
                    chartRange.HorizontalAlignment = 3;
                    chartRange.VerticalAlignment = 1;

                }
                lowerBound = "g";
                lowerBound += (row - 1);
                chartRange = xlWorkSheet.get_Range("a1", "g1");
                chartRange.Font.Bold = true;
                chartRange.Interior.Color = 18018018;


                chartRange = xlWorkSheet.get_Range("a1", lowerBound);
                chartRange.Cells.WrapText = true;


                chartRange.BorderAround(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlMedium, Excel.XlColorIndex.xlColorIndexAutomatic, Excel.XlColorIndex.xlColorIndexAutomatic);


                try
                {

                    xlWorkBook.SaveAs(txtSaveFolder.Text + "\\" + txtFileName.Text + ".xls", Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue, Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
                    xlWorkBook.Close(true, misValue, misValue);
                    xlApp.Quit();

                    releaseObject(xlApp);
                    releaseObject(xlWorkBook);
                    releaseObject(xlWorkSheet);
                    this.Cursor = Cursors.Arrow;
                    btnExport.Enabled = true;
                    btnCancel.Enabled = true;
                    btnTeamProject.Enabled = true;
                    btnFolderBrowse.Enabled = true;
                    comBoxTestPlan.Enabled = true;
                    comBoxTestSuite.Enabled = true;
                    MessageBox.Show("Test Cases exported successfully to specified file.", "Success", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);

                    txtTeamProject.Text = "";
                    comBoxTestPlan.Items.Clear();
                    comBoxTestSuite.Items.Clear();

                    txtSaveFolder.Text = "";
                    txtFileName.Text = "";
                    flag1 = flag2 = flag3 = flag4 = flag5 = 0;
                }
                catch (Exception ex)
                {
                    if (ex.Message == "Cannot access '" + txtFileName.Text + ".xls'.")
                    {
                        MessageBox.Show("File with same name exists in specified location", "File Exists", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                        txtFileName.Text = "";
                        flag5 = 0;
                    }
                    //else
                    //{
                        //MessageBox.Show("Application has encountered Fatal Errro. \nPlease contact your System Administrator.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        MessageBox.Show(ex.ToString(), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    //}
                }
            }
            else
            {
                MessageBox.Show("All fields are not populated.", "Missing Information", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
        }
        private void releaseObject(object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch (Exception ex)
            {
                obj = null;
                MessageBox.Show("Unable to release the Object " + ex.ToString());
            }
            finally
            {
                GC.Collect();
            }
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void btnAbout_Click(object sender, EventArgs e)
        {
            FrmAbout frmAbout = new FrmAbout();
            frmAbout.ShowDialog();
        }

    }
}

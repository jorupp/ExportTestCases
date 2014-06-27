using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using Microsoft.TeamFoundation.Client;
using Microsoft.TeamFoundation.TestManagement.Client;
using Microsoft.TeamFoundation.WorkItemTracking.Client;
using OfficeOpenXml;
using OfficeOpenXml.Style;

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
                var newSuite = suite_entry.TestSuite;
                comBoxTestSuite.Items.Add(newSuite.Title);
            }
        }

        private void Get_TestCases(ITestSuiteBase testSuite)
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
                var suite1 = suite.TestSuite;
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

                var filename = Path.Combine(txtSaveFolder.Text, txtFileName.Text + ".xlsx");
                Export(filename);
                Process.Start(filename);
                MessageBox.Show("Test Cases exported successfully to specified file.", "Success", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);

                this.Cursor = Cursors.Arrow;
                btnExport.Enabled = true;
                btnCancel.Enabled = true;
                btnTeamProject.Enabled = true;
                btnFolderBrowse.Enabled = true;
                comBoxTestPlan.Enabled = true;
                comBoxTestSuite.Enabled = true;

                txtTeamProject.Text = "";
                comBoxTestPlan.Items.Clear();
                comBoxTestSuite.Items.Clear();

                txtSaveFolder.Text = "";
                txtFileName.Text = "";
                flag1 = flag2 = flag3 = flag4 = flag5 = 0;
            }
            else
            {
                MessageBox.Show("All fields are not populated.", "Missing Information", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
        }

        private void Export(string filename)
        {
            using (var pkg = new ExcelPackage())
            {
                var sheet = pkg.Workbook.Worksheets.Add("Test Script");
                sheet.Cells[1, 1].Value = "Test Condition";
                sheet.Cells[1, 2].Value = "Action/Description";
                sheet.Cells[1, 3].Value = "Input Data";
                sheet.Cells[1, 4].Value = "Expected Result";
                sheet.Cells[1, 5].Value = "Actual Result";
                sheet.Cells[1, 6].Value = "Pass/Fail";
                sheet.Cells[1, 7].Value = "Comments";

                sheet.Column(1).Width = 15;
                sheet.Column(2).Width = 50;
                sheet.Column(3).Width = 15;
                sheet.Column(4).Width = 50;
                sheet.Column(5).Width = 50;
                sheet.Column(6).Width = 15;
                sheet.Column(7).Width = 20;


                int row = 2;
                foreach (var testCase in testCases)
                {
                    var firstRow = row;
                    foreach (var testAction in testCase.Actions)
                    {
                        AddSteps(sheet, testAction, ref row);
                    }
                    var merged = sheet.Cells[firstRow, 1, row - 1, 1];
                    merged.Merge = true;
                    merged.Value = CleanupText(testCase.Title);
                    merged.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    merged.Style.VerticalAlignment = ExcelVerticalAlignment.Top;
                }

                var header = sheet.Cells[1, 1, 1, 7];
                header.Style.Font.Bold = true;
                header.Style.Fill.PatternType = ExcelFillStyle.Solid;
                header.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(255, 226, 238, 18));

                sheet.Cells[1, 1, row, 7].Style.WrapText = true;

                //chartRange.BorderAround(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlMedium,
                //    Excel.XlColorIndex.xlColorIndexAutomatic, Excel.XlColorIndex.xlColorIndexAutomatic);

                pkg.SaveAs(new FileInfo(filename));
            }
            
        }

        private void AddSteps(ExcelWorksheet xlWorkSheet, ITestAction testAction, ref int row)
        {
            var testStep = testAction as ITestStep;
            var group = testAction as ITestActionGroup;
            var sharedRef = testAction as ISharedStepReference;
            if (null != testStep)
            {
                xlWorkSheet.Cells[row, 2].Value = CleanupText(testStep.Title.ToString());
                xlWorkSheet.Cells[row, 4].Value = CleanupText(testStep.ExpectedResult.ToString());
            }
            else if(null != group)
            {
                foreach (var action in group.Actions)
                {
                    AddSteps(xlWorkSheet, action, ref row);
                }
            }
            else if (null != sharedRef)
            {
                var step = sharedRef.FindSharedStep();
                foreach (var action in step.Actions)
                {
                    AddSteps(xlWorkSheet, action, ref row);
                }
            }
            row++;
        }

        private static readonly Regex _tag = new Regex("</?([A-Z][A-Z0-9]*)[^>]*>");
        private string CleanupText(string input)
        {
            input = input.Replace("</P><P>", Environment.NewLine);
            input = _tag.Replace(input, "");
            return input;
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

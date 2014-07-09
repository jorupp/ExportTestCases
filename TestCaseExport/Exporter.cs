using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using Microsoft.TeamFoundation.TestManagement.Client;
using OfficeOpenXml;
using OfficeOpenXml.Style;

namespace TestCaseExport
{
    /// <summary>
    /// Exports a passed set of test cases to the supplied file.
    /// </summary>
    public class Exporter
    {
        public void Export(string filename, ITestSuiteBase testSuite)
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
                foreach (var testCase in testSuite.TestCases.Select(i => i.TestCase))
                {
                    var replacementSets = GetReplacementSets(testCase);
                    foreach (var replacements in replacementSets)
                    {
                        var firstRow = row;
                        foreach (var testAction in testCase.Actions)
                        {
                            AddSteps(sheet, testAction, replacements, ref row);
                        }
                        if (firstRow != row)
                        {
                            var merged = sheet.Cells[firstRow, 1, row - 1, 1];
                            merged.Merge = true;
                            merged.Value = CleanupText(testCase.Title, replacements);
                            merged.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                            merged.Style.VerticalAlignment = ExcelVerticalAlignment.Top;
                        }
                    }
                }

                var header = sheet.Cells[1, 1, 1, 7];
                header.Style.Font.Bold = true;
                header.Style.Fill.PatternType = ExcelFillStyle.Solid;
                header.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(255, 226, 238, 18));

                sheet.Cells[1, 1, row, 7].Style.WrapText = true;

                pkg.SaveAs(new FileInfo(filename));
            }
        }

        private List<Dictionary<string, string>> GetReplacementSets(ITestCase testCase)
        {
            var replacementSets = new List<Dictionary<string, string>>();
            foreach (DataRow r in testCase.DefaultTableReadOnly.Rows)
            {
                var replacement = new Dictionary<string, string>();
                foreach (DataColumn c in testCase.DefaultTableReadOnly.Columns)
                {
                    replacement[c.ColumnName] = r[c] as string;
                }
                replacementSets.Add(replacement);
            }
            return replacementSets.DefaultIfEmpty(new Dictionary<string, string>()).ToList();
        }

        private void AddSteps(ExcelWorksheet xlWorkSheet, ITestAction testAction, Dictionary<string, string> replacements, ref int row)
        {
            var testStep = testAction as ITestStep;
            var group = testAction as ITestActionGroup;
            var sharedRef = testAction as ISharedStepReference;
            if (null != testStep)
            {
                xlWorkSheet.Cells[row, 2].Value = CleanupText(testStep.Title.ToString(), replacements);
                xlWorkSheet.Cells[row, 4].Value = CleanupText(testStep.ExpectedResult.ToString(), replacements);
            }
            else if (null != group)
            {
                foreach (var action in group.Actions)
                {
                    AddSteps(xlWorkSheet, action, replacements, ref row);
                }
            }
            else if (null != sharedRef)
            {
                var step = sharedRef.FindSharedStep();
                foreach (var action in step.Actions)
                {
                    AddSteps(xlWorkSheet, action, replacements, ref row);
                }
            }
            row++;
        }

        private static readonly Regex _tag = new Regex("</?([A-Z][A-Z0-9]*)[^>]*>");
        private string CleanupText(string input, Dictionary<string, string> replacements)
        {
            foreach (var kvp in replacements)
            {
                input = input.Replace("@" + kvp.Key, kvp.Value);
            }

            input = input.Replace("</P><P>", Environment.NewLine);
            input = _tag.Replace(input, "");
            return input;
        }
    }
}

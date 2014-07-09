using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using HtmlAgilityPack;
using OfficeOpenXml;
using OfficeOpenXml.Style;

namespace TestCaseExport
{
    /// <summary>
    /// Really ugly implementation of a basic HTML -> rich text parser
    /// </summary>
    public class HtmlToRichTextHelper
    {
        public void HtmlToRichText(ExcelRangeBase cell, string input)
        {
            var html = new HtmlDocument();
            html.LoadHtml(input);

            cell.IsRichText = true;
            Populate(cell.RichText, html.DocumentNode, new List<Action<ExcelRichText>>());
        }

        private void Populate(ExcelRichTextCollection richText, HtmlNode node, List<Action<ExcelRichText>> todo)
        {
            var toUnapply = new List<Action<ExcelRichText>>();
            switch (node.NodeType)
            {
                case HtmlNodeType.Document:
                case HtmlNodeType.Element:
                    switch (node.Name.ToUpperInvariant())
                    {
                        case "B":
                            {
                                bool oldval = false;
                                todo.Add(rt =>
                                {
                                    oldval = rt.Bold;
                                    rt.Bold = true;
                                });
                                toUnapply.Add(rt => rt.Bold = oldval);
                            }
                            break;
                        case "U":
                            {
                                bool oldval = false;
                                todo.Add(rt =>
                                {
                                    oldval = rt.UnderLine;
                                    rt.UnderLine = true;
                                });
                                toUnapply.Add(rt => rt.UnderLine = oldval);
                            }
                            break;
                        case "I":
                        case "EM":
                            {
                                bool oldval = false;
                                todo.Add(rt =>
                                {
                                    oldval = rt.Italic;
                                    rt.Italic = true;
                                });
                                toUnapply.Add(rt => rt.Italic = oldval);
                            }
                            break;
                    }
                    if (node.HasAttributes && node.Attributes.Contains("style"))
                    {
                        foreach (var part in node.Attributes["style"].Value.Split(';'))
                        {
                            var pieces = part.Split(':');
                            if (pieces.Length == 2)
                            {
                                switch (pieces[0])
                                {
                                    case "font-size":
                                        if (pieces[1].EndsWith("px") && pieces[1].Length > 2)
                                        {
                                            int v;
                                            if (int.TryParse(pieces[1].Substring(0, pieces[1].Length - 2), out v))
                                            {
                                                float oldval = 0;
                                                todo.Add(rt =>
                                                {
                                                    oldval = rt.Size;
                                                    rt.Size = v;
                                                });
                                                toUnapply.Add(rt => rt.Size = oldval);
                                            }
                                        }
                                        break;
                                    case "font-family":
                                        {
                                            string oldval = "";
                                            todo.Add(rt =>
                                            {
                                                oldval = rt.FontName;
                                                rt.FontName = pieces[1];
                                            });
                                            toUnapply.Add(rt => rt.FontName = oldval);
                                        }
                                        break;
                                    case "color":
                                        if (pieces[1].Trim().Length > 2)
                                        {
                                            int v;
                                            if (int.TryParse(pieces[1].Trim().Substring(1), NumberStyles.HexNumber,
                                                CultureInfo.InvariantCulture, out v))
                                            {
                                                Color oldval = Color.Black;
                                                todo.Add(rt =>
                                                {
                                                    oldval = rt.Color;
                                                    rt.Color = Color.FromArgb(v);
                                                });
                                                toUnapply.Add(rt => rt.Color = oldval);
                                            }
                                        }
                                        break;
                                    case "font-style":
                                        if (pieces[1] == "italic")
                                        {
                                            bool oldval = false;
                                            todo.Add(rt =>
                                            {
                                                oldval = rt.Italic;
                                                rt.Italic = true;
                                            });
                                            toUnapply.Add(rt => rt.Italic = oldval);
                                        }
                                        break;
                                    case "font-weight":
                                        if (pieces[1] == "bold")
                                        {
                                            bool oldval = false;
                                            todo.Add(rt =>
                                            {
                                                oldval = rt.Bold;
                                                rt.Bold = true;
                                            });
                                            toUnapply.Add(rt => rt.Bold = oldval);
                                        }
                                        break;
                                    case "text-decoration":
                                        if (pieces[1] == "underline")
                                        {
                                            bool oldval = false;
                                            todo.Add(rt =>
                                            {
                                                oldval = rt.UnderLine;
                                                rt.UnderLine = true;
                                            });
                                            toUnapply.Add(rt => rt.UnderLine = oldval);
                                        }
                                        break;
                                }
                            }
                        }
                    }
                    foreach (var child in node.ChildNodes)
                    {
                        Populate(richText, child, todo);
                    }
                    todo.AddRange(toUnapply);
                    switch (node.Name.ToUpperInvariant())
                    {
                        case "P":
                            if (node.NextSibling != null && node.NextSibling.Name.ToUpperInvariant() == "P")
                            {
                                var rt = richText.Add(Environment.NewLine);
                                foreach (var c in todo)
                                {
                                    c(rt);
                                }
                                todo.Clear();
                            }
                            break;
                    }
                    break;
                case HtmlNodeType.Text:
                    {
                        var rt = richText.Add(node.InnerText);
                        foreach (var c in todo)
                        {
                            c(rt);
                        }
                        todo.Clear();
                    }
                    break;
            }
        }
    }
}

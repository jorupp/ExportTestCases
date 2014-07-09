using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.Linq;
using HtmlAgilityPack;
using OfficeOpenXml;
using OfficeOpenXml.Style;

namespace TestCaseExport
{
    /// <summary>
    /// Less ugly implementation of a HTML -> Excel RichText parser.
    /// 
    /// Biggest issue here comes from the mismatch of API styles.  HTML is a tree.  Rich text is a stream with format changes able to be made at the beginning of each block.
    /// However, you *cannot* change the format without emtting a text block.
    /// Basically, we do a depth-first traversal, queueing changes as we go down, and queuing the reverse of our changes as we work our way back up.
    /// </summary>
    public class HtmlToRichTextHelper
    {
        public void HtmlToRichText(ExcelRangeBase cell, string input)
        {
            var html = new HtmlDocument();
            html.LoadHtml(input);

            cell.IsRichText = true;
            VisitHtmlNode(cell.RichText, html.DocumentNode, new List<Action<ExcelRichText>>());
        }

        /// <summary>
        /// Visits a single HTML node and updates the supplied rich text collection accordingly
        /// </summary>
        private void VisitHtmlNode(ExcelRichTextCollection richText, HtmlNode node, List<Action<ExcelRichText>> pendingActions)
        {
            switch (node.NodeType)
            {
                case HtmlNodeType.Document:
                case HtmlNodeType.Element:
                    var toAdd = GetModifiersForNode(node).ToList();
                    pendingActions.AddRange(toAdd.Select(i => (Action<ExcelRichText>)i.Apply));
                    foreach (var child in node.ChildNodes)
                    {
                        VisitHtmlNode(richText, child, pendingActions);
                    }
                    pendingActions.AddRange(toAdd.Select(i => (Action<ExcelRichText>)i.Undo));
                    switch (node.Name.ToUpperInvariant())
                    {
                        case "P":
                            if (node.NextSibling != null && node.NextSibling.Name.ToUpperInvariant() == "P")
                            {
                                AddRichText(richText, Environment.NewLine, pendingActions);
                            }
                            break;
                    }
                    break;
                case HtmlNodeType.Text:
                    AddRichText(richText, node.InnerText, pendingActions);
                    break;
            }
        }

        /// <summary>
        /// Creates a new RichText node with the supplied text and applies all the pending actions
        /// </summary>
        private static void AddRichText(ExcelRichTextCollection richText, string text, List<Action<ExcelRichText>> pendingActions)
        {
            var rt = richText.Add(text);
            foreach (var action in pendingActions)
            {
                action(rt);
            }
            pendingActions.Clear();
        }

        /// <summary>
        /// Given a HTML node, determine what RichText changes should be made to match the appearance of the passed node.
        /// </summary>
        /// <param name="node"></param>
        /// <returns></returns>
        private IEnumerable<IRichTextModifier> GetModifiersForNode(HtmlNode node)
        {
            switch (node.Name.ToUpperInvariant())
            {
                case "B":
                    yield return CreateSimpleModifier(i => i.Bold, (rt, v) => rt.Bold = v, true);
                    break;
                case "U":
                    yield return CreateSimpleModifier(i => i.UnderLine, (rt, v) => rt.UnderLine = v, true);
                    break;
                case "I":
                case "EM":
                    yield return CreateSimpleModifier(i => i.Italic, (rt, v) => rt.Italic = v, true);
                    break;
            }
            if (!node.HasAttributes || !node.Attributes.Contains("style")) 
                yield break;

            foreach (var part in node.Attributes["style"].Value.Split(new[] { ';' }, StringSplitOptions.RemoveEmptyEntries))
            {
                var pieces = part.Trim().Split(new [] {':'}, 2, StringSplitOptions.RemoveEmptyEntries);
                if (pieces.Length != 2) 
                    continue;

                var prop = pieces[0].ToLowerInvariant();
                var value = pieces[1];
                switch (prop)
                {
                    case "font-size":
                        var px = ParseAsPx(value);
                        if (px.HasValue)
                        {
                            yield return CreateSimpleModifier(i => i.Size, (rt, v) => rt.Size = v, px.Value);
                        }
                        break;
                    case "font-family":
                        yield return CreateSimpleModifier(i => i.FontName, (rt, v) => rt.FontName = v, value);
                        break;
                    case "color":
                        var color = ParseAsColor(value);
                        if (color.HasValue)
                        {
                            yield return CreateSimpleModifier(i => i.Color, (rt, v) => rt.Color = v, color.Value);
                        }
                        break;
                    case "font-style":
                        if (value == "italic")
                        {
                            yield return CreateSimpleModifier(i => i.Italic, (rt, v) => rt.Italic = v, true);
                        }
                        else if(value == "normal")
                        {
                            yield return CreateSimpleModifier(i => i.Italic, (rt, v) => rt.Italic = v, false);
                        }
                        break;
                    case "font-weight":
                        if (value == "bold")
                        {
                            yield return CreateSimpleModifier(i => i.Bold, (rt, v) => rt.Bold = v, true);
                        }
                        else if(value == "normal")
                        {
                            yield return CreateSimpleModifier(i => i.Bold, (rt, v) => rt.Bold = v, false);
                        }
                        break;
                    case "text-decoration":
                        if (value == "underline")
                        {
                            yield return CreateSimpleModifier(i => i.UnderLine, (rt, v) => rt.UnderLine = v, true);
                        }
                        else if(value == "none")
                        {
                            yield return CreateSimpleModifier(i => i.UnderLine, (rt, v) => rt.UnderLine = v, false);
                        }
                        break;
                }
            }
        }

        /// <summary>
        /// Parses the supplied value as a CSS color.  Currently, only simple hex values are supported.
        /// </summary>
        private Color? ParseAsColor(string value)
        {
            if (value.Length > 2 && value.StartsWith("#"))
            {
                int v;
                if (int.TryParse(value.Trim().Substring(1), NumberStyles.HexNumber, CultureInfo.InvariantCulture, out v))
                {
                    return Color.FromArgb(v);
                }
            }
            return null;
        }

        /// <summary>
        /// Parses the supplied value as a CSS size.  Currently, only px (and em translated to px @ 11px == 1em) are supported.
        /// </summary>
        private int? ParseAsPx(string value)
        {
            if (value.Length > 2 && value.EndsWith("px"))
            {
                int v;
                if (int.TryParse(value.Substring(0, value.Length - 2), out v))
                {
                    return v;
                }
            } else if (value.Length > 2 && value.EndsWith("em"))
            {
                int v;
                if (int.TryParse(value.Substring(0, value.Length - 2), out v))
                {
                    return v * 11;
                }
            }
            return null;
        }

        /// <summary>
        /// Delegates to <see cref="SimpleModifier{T}"/>'s constructor, but lets the compiler take care of inferring T for us.
        /// </summary>
        private IRichTextModifier CreateSimpleModifier<T>(Func<ExcelRichText, T> getter, Action<ExcelRichText, T> setter, T newValue)
        {
            return new SimpleModifier<T>(getter, setter, newValue);
        }

        /// <summary>
        /// Implemented by classes that support modifying rich text objets
        /// </summary>
        private interface IRichTextModifier
        {
            void Apply(ExcelRichText rt);
            void Undo(ExcelRichText rt);
        }

        /// <summary>
        /// Simple implementation of IRichTextModifier to set a specific value when Apply is called, and restore the original value when Undo is.
        /// </summary>
        /// <typeparam name="T"></typeparam>
        private class SimpleModifier<T> : IRichTextModifier
        {
            private readonly Func<ExcelRichText, T> _getter;
            private readonly Action<ExcelRichText, T> _setter;
            private readonly T _newValue;
            private T _oldValue;

            public SimpleModifier(Func<ExcelRichText, T> getter, Action<ExcelRichText, T> setter, T newValue)
            {
                _getter = getter;
                _setter = setter;
                _newValue = newValue;
            }

            public void Apply(ExcelRichText rt)
            {
                _oldValue = _getter(rt);
                _setter(rt, _newValue);
            }

            public void Undo(ExcelRichText rt)
            {
                _setter(rt, _oldValue);
            }
        }
    }
}

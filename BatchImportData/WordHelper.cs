using Aspose.Words;
using Aspose.Words.Replacing;
using Aspose.Words.Tables;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml;

namespace BatchImportData
{
    public static class WordExtensions
    {
        public static byte[] ToByte(this Document document)
        {
            var ms = new MemoryStream();
            document.Save(ms, SaveFormat.Docx);
            return ms.ToArray();
        }

        public static XmlNode GetSource(this Document document)
        {
            if (document.CustomXmlParts.Count == 0)
            {
                return null;
            }
            var cxp = document.CustomXmlParts[0];
            if (cxp.Data == null)
            {
                return null;
            }
            var xmlDoc = new XmlDocument();
            xmlDoc.Load(new MemoryStream(cxp.Data));

            var items = xmlDoc.GetElementsByTagName("b:Sources");
            if (items.Count == 0)
            {
                return null;
            }
            return items[0];
        }


        public static Paragraph FindParagraph(this Document doc, Func<Paragraph, bool> predicate)
        {
            NodeCollection paragraphs = doc.GetChildNodes(NodeType.Paragraph, true);
            foreach (Paragraph paragraph in paragraphs)
            {
                if (predicate(paragraph)) return paragraph;
            }
            return null;
        }


        public static Paragraph FindParagraphByTitleName(this Document doc, string titleName)
        {
            NodeCollection paragraphs = doc.GetChildNodes(NodeType.Paragraph, true);

            foreach (Paragraph paragraph in paragraphs)
            {
                if (paragraph.ParagraphFormat.OutlineLevel == OutlineLevel.BodyText && !paragraph.IsListItem)
                {
                    continue;
                }
                if (Regex.IsMatch(paragraph.GetText(), titleName, RegexOptions.IgnoreCase))
                {
                    return paragraph;
                }
            }
            return null;
        }

        public static Paragraph FindReferenceTitle(this Document doc)
        {
            NodeCollection paragraphs = doc.GetChildNodes(NodeType.Paragraph, true);

            foreach (Paragraph paragraph in paragraphs)
            {
                if (Regex.IsMatch(paragraph.GetText(), "References", RegexOptions.IgnoreCase)
                    && paragraph.ParagraphFormat.Style.Font.Bold)
                {
                    return paragraph;
                }
            }
            return null;
        }



        public static Paragraph FindParagraphByTitleName(this Document doc, Regex regex)
        {
            NodeCollection paragraphs = doc.GetChildNodes(NodeType.Paragraph, true);

            foreach (Paragraph paragraph in paragraphs)
            {
                if (paragraph.ParagraphFormat.OutlineLevel == OutlineLevel.BodyText && !paragraph.IsListItem)
                {
                    continue;
                }
                if (regex.IsMatch(paragraph.GetText()))
                {
                    return paragraph;
                }
            }
            return null;
        }
        
        public static Paragraph FindParagraph(this Document doc, Regex regex)
        {
            NodeCollection paragraphs = doc.GetChildNodes(NodeType.Paragraph, true);

            foreach (Paragraph paragraph in paragraphs)
            {
                string text = Utils.GetReplaceMethod(paragraph.GetText());
                if (regex.IsMatch(text))
                {
                    return paragraph;
                }
            }
            return null;
        }
        
        public static Paragraph NextParagraph(this Node node)
        {
            Node pre = null;
            do
            {
                pre = node;
                node = node.NextSibling;
            } while (!(node is Paragraph) && node != null);
            if (node == null)
            {
                var parent = pre.GetAncestor(NodeType.Section);
                var nextSection = (Section)parent.NextSibling;
                if (nextSection == null)
                    return null;

                return nextSection.Body.FirstParagraph;

            }
            return (Paragraph)node;
        }


        public static Table NextTable(this Node node)
        {
            do
            {
                node = node.NextSibling;
            } while (node != null && !(node is Table));
            return (Table)node;
        }

        public static Paragraph CloneParagraph(this Node node, bool isCloneChildren = true)
        {
            return (Paragraph)node.Clone(isCloneChildren);
        }

        public static void InsertAfter(this Node node, IEnumerable<Node> list)
        {
            var refer = node;
            foreach (var item in list)
            {
                node.ParentNode.InsertAfter(item, refer);
                refer = item;
            }
        }

        public static void InsertBefore(this Node node, IEnumerable<Node> list)
        {
            var refer = node;
            foreach (var item in list)
            {
                node.ParentNode.InsertBefore(item, refer);
                refer = item;
            }
        }

        public static void SetText(this Paragraph paragraph, string text)
        {
            Run run = null;
            if (paragraph.Runs.Count > 0)
            {
                run = (Run)paragraph.Runs.First().Clone(true);
                run.Text = text;
                paragraph.Runs.Clear();
            }
            else
            {
                run = new Run(paragraph.Document, text);
            }
            if (ExistsComments(text))
            {
                //paragraph.AppendRunWithComments(text);
                return;
            }
            else
            {
                paragraph.Runs.Add(run);
            }

            run.Font.HighlightColor = Color.Yellow;
        }
        public static void AppendText(this Paragraph paragraph, string text)
        {
            if (paragraph.Runs.Count == 0)
            {
                paragraph.Runs.Add(new Run(paragraph.Document));
            }
            Run run = (Run)paragraph.Runs.Last();
            run.Text += text;
            run.Font.HighlightColor = Color.Yellow;

            //Run run = null;
            //if (paragraph.Runs.Count > 0)
            //{
            //    run = new Run(paragraph.Document, text);
            //}
            //if (ExistsComments(text))
            //{
            //    //paragraph.AppendRunWithComments(text);
            //    return;
            //}
            //else
            //{
            //    paragraph.Runs.Add(run);
            //}

        }
        public static void SetText(this Cell cell, string text)
        {
            if (cell.Paragraphs.Count > 1)
            {
                for (int i = cell.Paragraphs.Count - 1; i > 0; i--)
                {
                    cell.Paragraphs[i].Remove();
                }
            }
            else if (cell.Paragraphs.Count == 0)
            {
                cell.Paragraphs.Add(new Paragraph(cell.Document));
            }
            cell.Paragraphs[0].SetText(text);
        }

        public static void AppendText(this Cell cell, string text)
        {
            if (cell.Paragraphs.Count == 0)
            {
                cell.Paragraphs.Add(new Paragraph(cell.Document));
            }
            Paragraph p = (Paragraph)cell.Paragraphs.Last();
            if (p.Runs.Count == 0)
            {
                p.Runs.Add(new Run(cell.Document));
            }
            ((Run)p.Runs.Last()).Text += text;


        }
        public static bool ExistsComments(this string text)
        {
            if (string.IsNullOrEmpty(text))
            {
                return false;
            }
            return text.Contains(CommentsQualifer.CommentsBegin, StringComparison.OrdinalIgnoreCase) && text.Contains(CommentsQualifer.CommentsEnd, StringComparison.OrdinalIgnoreCase);
        }

        /// <summary>
        /// 替换段落文本
        /// </summary>
        /// <param name="paragraph">段落</param>
        /// <param name="pattern">要替换的文本</param>
        /// <param name="replace">替换的值</param>
        public static Paragraph ReplaceText(this Paragraph paragraph, string pattern, string replace)
        {
            if (ExistsComments(replace))
            {
                // 搜索关键字, 高亮并且添加错误注释
                FindReplaceOptions options = new FindReplaceOptions();
                options.ReplacingCallback = new WrapCommentsReplacingCallback();
                var matches = Regex.Matches(replace, @CommentsQualifer.CommentsBegin + "(.*?)" + CommentsQualifer.CommentsEnd);
                paragraph.Range.Replace(pattern, replace);
                foreach (var match in matches)
                {
                    paragraph.Range.Replace(match.ToString(), match.ToString(), options);
                }
            }
            else
            {
                paragraph.Range.Replace(pattern, replace == null ? "" : replace);
                paragraph.HighLightField(replace);
            }
            return paragraph;
        }

        public static Paragraph ReplaceText(this Paragraph paragraph, Regex pattern, string replace)
        {
            if (replace == null)
            {
                return paragraph;
            }
            if (ExistsComments(replace))
            {
                // 搜索关键字, 高亮并且添加错误注释
                FindReplaceOptions options = new FindReplaceOptions();
                options.ReplacingCallback = new WrapCommentsReplacingCallback();

                paragraph.Range.Replace(pattern, replace, options);
            }
            else
            {
                paragraph.Range.Replace(pattern, replace);
                paragraph.HighLightField(replace);
            }
            return paragraph;
        }

        public static void ReplaceText(this Cell cell, string source, string text)
        {
            foreach (Paragraph item in cell.Paragraphs)
            {
                item.ReplaceText(source, text);
            }
        }

        public static void ReplaceText(this Cell cell, Regex source, string text)
        {
            foreach (Paragraph item in cell.Paragraphs)
            {
                item.ReplaceText(source, text);
            }
        }

        /// <summary>
        /// 高亮段落中的某些文本
        /// </summary>
        /// <param name="p"></param>
        /// <param name="highLightText"></param>
        public static void HighLightField(this Paragraph p, string highLightText = null)
        {
            if (string.IsNullOrEmpty(highLightText))
                return;
            FindReplaceOptions options = new FindReplaceOptions();
            var callback = new HighLightReplaceCallback();
            options.ReplacingCallback = callback;
            p.Range.Replace(highLightText, string.Empty, options);
        }

        public static void HighLightField(this Cell cell, string highLightText = null)
        {
            for (int i = 0; i < cell.Paragraphs.Count; i++)
            {
                if (highLightText == null)
                {
                    cell.Paragraphs[i].HighLightField(cell.Paragraphs[i].GetText());
                }
                else
                {
                    cell.Paragraphs[i].HighLightField(highLightText);
                }
            }
        }

        #region 私有方法,类

        private static Node SplitRun(Run run, int position)
        {
            if (position > run.Text.Length)
            {
                while (run.NextSibling != null && position > run.Text.Length)
                {
                    position = position - run.Text.Length;
                    run = (Run)run.NextSibling;
                }
            }

            Run afterRun = (Run)run.Clone(true);
            afterRun.Text = run.Text.Substring(position);
            run.Text = run.Text.Substring(0, position);
            run.ParentNode.InsertAfter(afterRun, run);
            return afterRun;
        }

        private class HighLightReplaceCallback : IReplacingCallback
        {
            public ReplaceAction Replacing(ReplacingArgs e)
            {
                Node currentNode = e.MatchNode;

                if (e.MatchOffset > 0)
                {
                    currentNode = SplitRun((Run)currentNode, e.MatchOffset);
                }

                ArrayList runs = new ArrayList();
                int remainingLength = e.Match.Value.Length;
                while (
                    (remainingLength > 0) &&
                    (currentNode != null) &&
                    (currentNode.GetText().Length <= remainingLength))
                {
                    runs.Add(currentNode);
                    remainingLength = remainingLength - currentNode.GetText().Length;

                    do
                    {
                        currentNode = currentNode.NextSibling;
                    } while ((currentNode != null) && (currentNode.NodeType != NodeType.Run));
                }

                if ((currentNode != null) && (remainingLength > 0))
                {
                    SplitRun((Run)currentNode, remainingLength);
                    runs.Add(currentNode);
                }

                foreach (Run run in runs)
                {
                    run.Font.HighlightColor = Color.Yellow;
                }
                return ReplaceAction.Skip;
            }
        }

        private class WrapCommentsReplacingCallback : IReplacingCallback
        {
            public ReplaceAction Replacing(ReplacingArgs e)
            {
                Node currentNode = e.MatchNode;
                if (currentNode.NodeType != NodeType.Run || currentNode.Range.Text.Trim() == string.Empty)
                    return ReplaceAction.Skip;
                if (e.MatchOffset > 0)
                {
                    currentNode = SplitRun((Run)currentNode, e.MatchOffset);
                }

                List<Node> runs = new List<Node>();
                int remainingLength = e.Match.Value.Length;
                while (
                    (remainingLength > 0) &&
                    (currentNode != null) &&
                    (currentNode.GetText().Length <= remainingLength))
                {
                    runs.Add(currentNode);
                    remainingLength = remainingLength - currentNode.GetText().Length;

                    do
                    {
                        currentNode = currentNode.NextSibling;
                    } while ((currentNode != null) && (currentNode.NodeType != NodeType.Run));
                }

                if ((currentNode != null) && (remainingLength > 0))
                {
                    SplitRun((Run)currentNode, remainingLength);
                    runs.Add(currentNode);
                }

                foreach (Run run in runs)
                {
                    run.Text = string.Empty;
                    run.Font.HighlightColor = CommentsQualifer.GetHighlightColor();
                }
                ((Run)runs.First()).Text = CommentsQualifer.CommentErrorText;
                Aspose.Words.Comment comment = new Aspose.Words.Comment(runs.First().Document);
                comment.Paragraphs.Add(new Paragraph(comment.Document));
                string text = e.Replacement;
                text = text.Replace(CommentsQualifer.CommentsBegin, string.Empty);
                text = text.Replace(CommentsQualifer.CommentsEnd, string.Empty);
                comment.FirstParagraph.SetText(text);

                runs.First().ParentNode.InsertBefore(new CommentRangeStart(comment.Document, comment.Id), runs.First());
                runs.First().ParentNode.InsertAfter(new CommentRangeEnd(comment.Document, comment.Id), runs.Last());

                runs.Last().ParentNode.AppendChild(comment);

                return ReplaceAction.Skip;
            }
        }
        #endregion

    }
}

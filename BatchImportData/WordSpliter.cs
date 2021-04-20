using Aspose.Words;
using Aspose.Words.Markup;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml;
using System.Xml.Linq;
using System.Linq;
using Aspose.Words.Fields;

namespace BatchImportData
{
    public static class WordUtil
    {
        public static void Split(List<Paragraph> paragraphs, string fileName)
        {
            Document document = new Document();

            foreach (Paragraph p in paragraphs)
            {
                var node = document.ImportNode(p, true);
                document.FirstSection.Body.ChildNodes.Add(node);
            }
            document.Save(fileName);
        }

        public static List<string> ExtractCitation(List<Paragraph> paragraphs, Document doc)
        {
            if (paragraphs == null || paragraphs.Count == 0)
            {
                return null;
            }
            HashSet<string> citationMapping = new HashSet<string>();
            foreach (Paragraph paragraph in paragraphs)
            {
                var citations = paragraph.GetChildNodes(NodeType.StructuredDocumentTag, true);
                if (citations.Count == 0)
                {
                    continue;
                }
                foreach (StructuredDocumentTag item in citations)
                {
                    string id = Regex.Match(item.GetText(), @"(?<=CITATION )\S*").Value.Trim();
                    citationMapping.Add(id);
                }
            }

            XmlNode sources = doc.GetSource();
            List<string> res = new List<string>();
            foreach (var item in citationMapping)
            {
                res.Add(GetCitationXml(sources, item));
            }
            return res;
        }

        public static Document GenerateCitationFile(string text)
        {
            Document doc = new Document("template.docx");
            var sources = doc.GetSource();

            var loadDoc = new XmlDocument();
            loadDoc.LoadXml(text);

            sources.AppendChild(sources.OwnerDocument.ImportNode(loadDoc.DocumentElement, true));
            MemoryStream ms = new MemoryStream();
            sources.OwnerDocument.Save(ms);
            doc.CustomXmlParts[0].Data = ms.ToArray();

            return doc;
        }

        private static string GetCitationXml(XmlNode source, string identification)
        {
            XmlNode matchItem = null;
            foreach (XmlNode item in source.ChildNodes)
            {
                bool notMatch = true;
                foreach (XmlNode node in item.ChildNodes)
                {
                    if (node.Name != "b:Tag")
                    {
                        continue;
                    }
                    if (node.InnerText != identification)
                    {
                        break;
                    }
                    notMatch = false;
                }
                if (notMatch)
                {
                    continue;
                }
                matchItem = item;
                break;
            }
            if (matchItem == null)
            {
                return "没有找到匹配的源";
            }
            return matchItem.ToString(0);
        }

        public static string ToString(this XmlNode node, int indentation)
        {
            string result = XElement.Parse(node.OuterXml).ToString();
            return result;
        }
    }
}

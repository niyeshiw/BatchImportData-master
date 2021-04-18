using System.Drawing;

namespace BatchImportData
{
    public class CommentsQualifer
    {
        public const string CommentsBegin = "<%";
        public const string CommentsEnd = "%>";
        public const string CommentErrorText = "Undefined";

        public static Color GetHighlightColor()
        {
            return Color.Red;
        }
    }
}

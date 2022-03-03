using System;
using System.Windows.Forms;

namespace clipboard_converter_for_teams
{
    // クリップボード操作のラッパー
    public static class ClipboardWrapper
    {
        public static string GetText()
        {
            return Clipboard.GetText(TextDataFormat.Text);
        }

        public static string GetHtml()
        {
            return Clipboard.GetText(TextDataFormat.Html);
        }

        public static string GetHtmlFragmentPart()
        {
            return extractFragmentPart(ClipboardWrapper.GetHtml());
        }

        public static bool ContainsText()
        {
            return Clipboard.ContainsText(TextDataFormat.Text);
        }

        public static bool ContainsHtml()
        {
            return Clipboard.ContainsText(TextDataFormat.Html);
        }

        public static void SetHtmlFragmentPartAndText(string html_fragment_part, string text)
        {
            ClipboardHelper.CopyToClipboard(html_fragment_part, text);
        }

        static string extractFragmentPart(string html)
        {
            string start_keyword = "<!--StartFragment-->";  // 開始キーワード
            string end_keyword = "<!--EndFragment-->";      // 終了キーワード
            int start_index = html.IndexOf(start_keyword);  // 開始キーワードの検出位置
            int end_index = html.IndexOf(end_keyword);      // 終了キーワードの検出位置

            // 正常にキーワードが見つかった場合、各インデックスは0より大きくなるはず
            if (start_index > 0 && end_index > 0)
            {
                int start = start_index + start_keyword.Length;
                int length = end_index - start;
                return html.Substring(start, length);
            }
            // 異常時は空文字列を返す
            else
            {
                return String.Empty;
            }
        }
    }
}

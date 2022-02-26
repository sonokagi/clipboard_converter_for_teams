using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Windows.Forms;
using clipboard_converter_for_teams;

namespace UnitTest
{
    [TestClass]
    public class ClipboardConverterCollectionTest
    {
        // httpから始まるテキスト
        static string text_start_with_http_ = "http.....";
        
        // 変換結果として期待する、上記URLへのリンクのHtmlデータ
        static string expect_html_link_to_url_ = "<a href=\"" + text_start_with_http_ + "\"><b><i>HERE!</i></b></a>";

        // http以外で始まるテキスト
        static string other_text_ = "other.....";


        // 投稿へのリンクのHtmlデータ(Teamsで、投稿へのリンクをコピーした際、クリップボードに格納されるHtmlデータの想定値)
        static string html_link_to_post_ = "<div itemprop=\"teams-copy-link\"><a href=\"URL\" title=\"TITLE\">POSTER: POSTED_CONTENTS</a></div><div itemprop=\"teams-copy-link\">POST_TEAM_CHANNEL_DATE_TIME</div><div>&nbsp;</div>";
        
        // 変換結果として期待する、上記を短縮したHtmlデータ
        static string expect_shortened_html_ =  "<a href=\"URL\" title=\"TITLE\">POSTED_CONTENTS</a>";

        // 投稿へのリンク以外のHtmlデータ(例えば、単純なリンク)
        static string other_html_= "<a href=\"URL\">CONTENTS</a>";

        [TestInitialize]
        public void TestInitialize()
        {
            Clipboard.Clear();
        }

        [TestMethod]
        public void Empty_Then_NoAction()
        {
            // クリップボードが空の場合
            Clipboard.Clear();

            ClipboardConverterCollection.Execute();

            // テキストもHtmlデータも無しのまま
            Helper.CheckHasNoTextData();
            Helper.CheckHasNoHtmlData();
        }

        [TestMethod]
        public void TextStartWithHttp_Then_CreateHtmlFormatLink()
        {
            // httpから始まるテキストがある場合
            Helper.SetText(text_start_with_http_);

            ClipboardConverterCollection.Execute();

            // テキストは変化しない
            Helper.CheckText(text_start_with_http_);

            // Htmlデータにリンクが格納される
            Helper.CheckHtmlFragmentPart(expect_html_link_to_url_);
        }

        [TestMethod]
        public void OtherText_Then_NoAction()
        {
            // http以外で始まるテキストがある場合
            Helper.SetText(other_text_);

            ClipboardConverterCollection.Execute();

            // テキストは変化しない
            Helper.CheckText(other_text_);

            // Html形式データは無しのまま
            Helper.CheckHasNoHtmlData();
        }

        [TestMethod]
        public void HtmlLinkToPost_Then_ShortenHtmlLinkToPost()
        {
            // 投稿へのリンクのHtmlデータと、http以外で始まるテキストがある場合
            Helper.SetHtmlFragmentPartAndSetText(html_link_to_post_, other_text_);

            ClipboardConverterCollection.Execute();

            // テキストは変化しない
            Helper.CheckText(other_text_);

            // Htmlデータが短縮したデータに置き換わる
            Helper.CheckHtmlFragmentPart(expect_shortened_html_);
        }

        [TestMethod]
        public void OtherHtml_Then_NoAction()
        {
            // 投稿へのリンク以外のHtmlデータと、http以外で始まるテキストがある場合
            Helper.SetHtmlFragmentPartAndSetText(other_html_, other_text_);

            ClipboardConverterCollection.Execute();

            // テキストもHtmlデータも変化しない
            Helper.CheckText(other_text_);
            Helper.CheckHtmlFragmentPart(other_html_);
        }

        [TestMethod]
        public void HtmlLinkToPostAndTextStartWithHttp_Then_ShortenHtmlLinkToPost()
        {
            // 投稿へのリンクのHtmlデータと、httpから始まるテキスト の両方がある場合
            Helper.SetHtmlFragmentPartAndSetText(html_link_to_post_, text_start_with_http_);

            ClipboardConverterCollection.Execute();

            // テキストは変化しない
            Helper.CheckText(text_start_with_http_);

            // Htmlデータが短縮したデータに置き換わる
            Helper.CheckHtmlFragmentPart(expect_shortened_html_);
        }
    }

    [TestClass]
    public class ClipboardHelperLearningTest
    {
        // 分かったこと
        // - htmlデータだけをクリップボードに設定することはできない
        //   plainText引数に null を渡しても、空のテキストが設定されてしまう

        [TestInitialize]
        public void TestInitialize()
        {
            Clipboard.Clear();
        }

        [TestMethod]
        [Ignore]    // このテストはときどき例外を発生させるので、テストから除外
        public void SetNullToText_Then_ClipboardHasEmptyText()
        {
            // テキストを null で指定した場合
            ClipboardHelper.CopyToClipboard("html", null);
            // 空のテキストデータが格納される
            Helper.CheckHasTextData();
            Helper.CheckText(String.Empty);
        }

        [TestMethod]
        public void SetEmptyToText_Then_ClipboardHasEmptyText()
        {
            // テキストを空で指定した場合
            ClipboardHelper.CopyToClipboard("html", String.Empty);
            // 空のテキストデータが格納される
            Helper.CheckHasTextData();
            Helper.CheckText(String.Empty);
        }
    }
    public static class Helper
    {
        // 指定のTextデータをクリップボードに設定する
        public static void SetText(string expect)
        {
            Clipboard.SetText(expect, TextDataFormat.Text);
        }

        // Textデータが期待通りかチェックする
        public static void CheckText(string expect)
        {
            string actual = Clipboard.GetText(TextDataFormat.Text);
            Assert.AreEqual(expect, actual);
        }

        // 指定のHtmlデータをフォーマット変換してクリップボードに設定し、かつ、指定のTextデータをクリップボードに設定する
        // - Html形式データをクリップボードに格納する場合、下記URLのような変換が必要
        //   https://docs.microsoft.com/ja-jp/windows/win32/dataxchg/html-clipboard-format
        // - Htmlとテキストを設定する場合、両方を同時に設定する必要があり、処理を分割できない
        public static void SetHtmlFragmentPartAndSetText(string expect_html_fragment_part, string expect_text)
        {
            ClipboardHelper.CopyToClipboard(expect_html_fragment_part, expect_text);
        }

        // Htmlデータの <!--StartFragment--> ～ <!--EndFragment--> の部分が期待通りかチェックする
        public static void CheckHtmlFragmentPart(string expect)
        {
            string raw_data = Clipboard.GetText(TextDataFormat.Html);
            string actual = extractFragmentPart(raw_data);
            Assert.AreEqual(expect, actual);
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
                string part = html.Substring(start, length);
                return part;
            }
            // 異常時はテストを失敗させる
            else
            {
                Assert.Fail("cannot find Fragment keywords in data.");
                return String.Empty;
            }
        }

        public static void CheckHasTextData()
        {
            Assert.AreEqual(true, Clipboard.ContainsData(DataFormats.Text));
        }

        public static void CheckHasNoTextData()
        {
            Assert.AreEqual(false, Clipboard.ContainsData(DataFormats.Text));
        }

        public static void CheckHasNoHtmlData()
        {
            Assert.AreEqual(false, Clipboard.ContainsData(DataFormats.Html));
        }
    }
}

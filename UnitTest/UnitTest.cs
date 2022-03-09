using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using clipboard_converter_for_teams;
using System.IO;

namespace UnitTest
{
    [TestClass]
    public class ClipboardConverterCollectionTest
    {
        // httpから始まるテキスト
        private const string text_start_with_http_ = "http.....";

        // 変換結果として期待する、上記URLへのリンクのHtmlデータ
        private const string expect_html_link_to_url_ = "<a href=\"" + text_start_with_http_ + "\"><b><i>HERE!</i></b></a>";

        // http以外で始まるテキスト
        private const string other_text_ = "other.....";


        // 投稿へのリンクのHtmlデータ(Teamsで、投稿へのリンクをコピーした際、クリップボードに格納されるHtmlデータの想定値)
        private const string html_link_to_post_ = "<div itemprop=\"teams-copy-link\"><a href=\"URL\" title=\"TITLE\">POSTER: POSTED_CONTENTS</a></div><div itemprop=\"teams-copy-link\">POST_TEAM_CHANNEL_DATE_TIME</div><div>&nbsp;</div>";

        // 変換結果として期待する、上記を短縮したHtmlデータ
        private const string expect_shortened_html_ = "<a href=\"URL\" title=\"TITLE\">POSTED_CONTENTS</a>";

        // 投稿へのリンク以外のHtmlデータ(例えば、単純なリンク)
        private const string other_html_ = "<a href=\"URL\">CONTENTS</a>";

        [TestInitialize]
        public void TestInitialize()
        {
            ClipboardWrapper.Clear();
        }

        [TestMethod]
        public void Empty_Then_NoAction()
        {
            // クリップボードが空の場合
            ClipboardWrapper.Clear();

            ClipboardConverterCollection.Execute();

            // テキストもHtmlデータも無しのまま
            Check.HasNoTextData();
            Check.HasNoHtmlData();
        }

        [TestMethod]
        public void TextStartWithHttp_Then_CreateHtmlFormatLink()
        {
            // httpから始まるテキストがある場合
            ClipboardWrapper.SetText(text_start_with_http_);

            ClipboardConverterCollection.Execute();

            // テキストは変化しない、Htmlデータにリンクが格納される
            Check.Text(text_start_with_http_);
            Check.HtmlFragmentPart(expect_html_link_to_url_);
        }

        [TestMethod]
        public void OtherText_Then_NoAction()
        {
            // http以外で始まるテキストがある場合
            ClipboardWrapper.SetText(other_text_);

            ClipboardConverterCollection.Execute();

            // テキストは変化しない、Html形式データは無しのまま
            Check.Text(other_text_);
            Check.HasNoHtmlData();
        }

        [TestMethod]
        public void HtmlLinkToPost_Then_ShortenHtmlLinkToPost()
        {
            // 投稿へのリンクのHtmlデータと、http以外で始まるテキストがある場合
            ClipboardWrapper.SetHtmlFragmentPartAndText(html_link_to_post_, other_text_);

            ClipboardConverterCollection.Execute();

            // テキストは変化しない、Htmlデータが短縮したデータに置き換わる
            Check.Text(other_text_);
            Check.HtmlFragmentPart(expect_shortened_html_);
        }

        [TestMethod]
        public void OtherHtml_Then_NoAction()
        {
            // 投稿へのリンク以外のHtmlデータと、http以外で始まるテキストがある場合
            ClipboardWrapper.SetHtmlFragmentPartAndText(other_html_, other_text_);

            ClipboardConverterCollection.Execute();

            // テキストもHtmlデータも変化しない
            Check.Text(other_text_);
            Check.HtmlFragmentPart(other_html_);
        }

        [TestMethod]
        public void HtmlLinkToPostAndTextStartWithHttp_Then_ShortenHtmlLinkToPost()
        {
            // 投稿へのリンクのHtmlデータと、httpから始まるテキスト の両方がある場合
            ClipboardWrapper.SetHtmlFragmentPartAndText(html_link_to_post_, text_start_with_http_);

            ClipboardConverterCollection.Execute();

            // テキストは変化しない、Htmlデータが短縮したデータに置き換わる
            Check.Text(text_start_with_http_);
            Check.HtmlFragmentPart(expect_shortened_html_);
        }

        [TestMethod]
        public void Empty_Then_OutputNothing()
        {
            using (var output = new StringWriter())
            {
                var stdout = Console.Out;
                Console.SetOut(output);

                // クリップボードが空の場合
                ClipboardWrapper.Clear();

                ClipboardConverterCollection.Execute();

                // コンソールには何も出力されない
                Assert.AreEqual(String.Empty, output.ToString());

                Console.SetOut(stdout);
            }
        }

        [TestMethod]
        public void HtmlLinkToPost_Then_OutputCoversionResult()
        {
            using (var output = new StringWriter())
            {
                var stdout = Console.Out;
                Console.SetOut(output);

                // 投稿へのリンクのHtmlデータと、http以外で始まるテキストがある場合
                ClipboardWrapper.SetHtmlFragmentPartAndText(html_link_to_post_, other_text_);

                ClipboardConverterCollection.Execute();

                // コンソールには変換結果を出力
                string expected =
                      "--- [before] ---\r\n"
                    + html_link_to_post_ + "\r\n"
                    + "--- [after] ---\r\n"
                    + expect_shortened_html_ + "\r\n";
                Assert.AreEqual(expected, output.ToString());

                Console.SetOut(stdout);
            }
        }

        [TestMethod]
        public void TextStartWithHttp_Then_OutputCoversionResult()
        {
            using (var output = new StringWriter())
            {
                var stdout = Console.Out;
                Console.SetOut(output);

                // httpから始まるテキストがある場合
                ClipboardWrapper.SetText(text_start_with_http_);

                ClipboardConverterCollection.Execute();

                // コンソールには変換結果を出力
                string expected =
                      "--- [before] ---\r\n"
                    + text_start_with_http_ + "\r\n"
                    + "--- [after] ---\r\n"
                    + expect_html_link_to_url_ + "\r\n";
                Assert.AreEqual(expected, output.ToString());

                Console.SetOut(stdout);
            }
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
            ClipboardWrapper.Clear();
        }

        [TestMethod]
        [Ignore]    // このテストはときどき例外を発生させるので、テストから除外
        public void SetNullToText_Then_ClipboardHasEmptyText()
        {
            // テキストを null で指定した場合
            ClipboardHelper.CopyToClipboard("html", null);
            // 空のテキストデータが格納される
            Check.HasTextData();
            Check.Text(String.Empty);
        }

        [TestMethod]
        public void SetEmptyToText_Then_ClipboardHasEmptyText()
        {
            // テキストを空で指定した場合
            ClipboardHelper.CopyToClipboard("html", String.Empty);
            // 空のテキストデータが格納される
            Check.HasTextData();
            Check.Text(String.Empty);
        }
    }

    public static class Check
    {
        // Textデータが期待通りかチェックする
        public static void Text(string expect)
        {
            Assert.AreEqual(expect, ClipboardWrapper.GetText());
        }

        // Htmlデータの <!--StartFragment--> ～ <!--EndFragment--> の部分が期待通りかチェックする
        public static void HtmlFragmentPart(string expect)
        {
            Assert.AreEqual(expect, ClipboardWrapper.GetHtmlFragmentPart());
        }

        public static void HasTextData()
        {
            Assert.AreEqual(true, ClipboardWrapper.ContainsText());
        }

        public static void HasNoTextData()
        {
            Assert.AreEqual(false, ClipboardWrapper.ContainsText());
        }

        public static void HasNoHtmlData()
        {
            Assert.AreEqual(false, ClipboardWrapper.ContainsHtml());
        }
    }
}

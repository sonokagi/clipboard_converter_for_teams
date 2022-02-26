﻿using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Windows.Forms;
using clipboard_converter_for_teams;

namespace UnitTest
{
    [TestClass]
    public class ClipboardConverterCollectionTest
    {
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

            // テキストもHtml形式データも無しのまま
            Helper.CheckHasNoTextData();
            Helper.CheckHasNoHtmlData();
        }

        [TestMethod]
        public void TextStartWithHttp_Then_CreateHtmlFormatLink()
        {
            string http_text = "http.....";

            // httpから始まるテキストがある場合
            Helper.SetText(http_text);

            ClipboardConverterCollection.Execute();

            // テキストは変化しない
            Helper.CheckText(http_text);

            // Htmlデータにリンクが格納される
            string html_format_link = "<a href=\"" + http_text + "\"><b><i>HERE!</i></b></a>";
            Helper.CheckHtmlFragmentPart(html_format_link);
        }

        [TestMethod]
        public void OtherText_Then_NoAction()
        {
            string other_text = "other.....";

            // http以外で始まるテキストがある場合
            Helper.SetText(other_text);

            ClipboardConverterCollection.Execute();

            // テキストは変化しない
            Helper.CheckText(other_text);

            // Html形式データは無しのまま
            Helper.CheckHasNoHtmlData();
        }

        [TestMethod]
        public void HtmlLinkToPost_Then_ShortenHtmlLinkToPost()
        {
            // Teamsで、投稿へのリンクをコピーした際、クリップボードに格納されるHtmlデータの期待値
            string html_fragment_of_link_to_post  =
                "<div itemprop=\"teams-copy-link\"><a href=\"URL\" title=\"TITLE\">POSTER: POSTED_CONTENTS</a></div><div itemprop=\"teams-copy-link\">POST_TEAM_CHANNEL_DATE_TIME</div><div>&nbsp;</div>";

            // 変換結果として期待する、短縮されたHtmlデータ
            string shortened_html_fragment =
                "<a href=\"URL\" title=\"TITLE\">POSTED_CONTENTS</a>";

            string any_text = "any text";

            // 投稿へのリンクのHtmlデータと、何らかのテキストがある場合
            Helper.SetHtmlFragmentPartAndSetText(html_fragment_of_link_to_post, any_text);

            ClipboardConverterCollection.Execute();

            // テキストは変化しない
            Helper.CheckText(any_text);

            // Htmlデータが短縮したデータに置き換わる
            Helper.CheckHtmlFragmentPart(shortened_html_fragment);
        }

        [TestMethod]
        public void OtherHtml_Then_NoAction()
        {
            // 投稿へのリンク以外のHtmlデータ(例えば、上記で短縮されたHtmlデータ)
            string other_html_fragment = "<a href=\"URL\" title=\"TITLE\">POSTER: POSTED_CONTENTS</a>";

            string any_text = "any text";

            // 投稿へのリンク以外のHtmlデータと、何らかのテキストがある場合
            Helper.SetHtmlFragmentPartAndSetText(other_html_fragment, any_text);

            ClipboardConverterCollection.Execute();

            // テキストは変化しない
            Helper.CheckText(any_text);

            // Htmlデータは変化しない
            Helper.CheckHtmlFragmentPart(other_html_fragment);
        }

        [TestMethod]
        public void HtmlLinkToPostAndTextStartWithHttp_Then_ShortenHtmlLinkToPost()
        {
            // Teamsで、投稿へのリンクをコピーした際、クリップボードに格納されるHtmlデータの期待値
            string html_fragment_of_link_to_post =
                "<div itemprop=\"teams-copy-link\"><a href=\"URL\" title=\"TITLE\">POSTER: POSTED_CONTENTS</a></div><div itemprop=\"teams-copy-link\">POST_TEAM_CHANNEL_DATE_TIME</div><div>&nbsp;</div>";

            // 変換結果として期待する、短縮されたHtmlデータ
            string shortened_html_fragment =
                "<a href=\"URL\" title=\"TITLE\">POSTED_CONTENTS</a>";

            // httpから始まるテキストがある場合
            string http_text = "http.....";

            // 投稿へのリンクのHtmlデータと、httpから始まるテキスト の両方がある場合
            Helper.SetHtmlFragmentPartAndSetText(html_fragment_of_link_to_post, http_text);

            ClipboardConverterCollection.Execute();

            // テキストは変化しない
            Helper.CheckText(http_text);

            // Htmlデータが短縮したデータに置き換わる
            Helper.CheckHtmlFragmentPart(shortened_html_fragment);
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

        // 指定のHtmlデータをクリップボードに設定する
        public static void SetHtml(string expect)
        {
            Clipboard.SetText(expect, TextDataFormat.Html);
        }

        // Htmlデータが期待通りかチェックする
        public static void CheckHtml(string expect)
        {
            // Html形式データを取得時、最後に終端'\0'がつくようなので削除する
            string actual = Clipboard.GetText(TextDataFormat.Html).TrimEnd('\0');
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
        public static void CheckHasHtmlData()
        {
            Assert.AreEqual(true, Clipboard.ContainsData(DataFormats.Html));
        }

        public static void CheckHasNoHtmlData()
        {
            Assert.AreEqual(false, Clipboard.ContainsData(DataFormats.Html));
        }
    }
}

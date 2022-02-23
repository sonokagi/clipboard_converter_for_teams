using Microsoft.VisualStudio.TestTools.UnitTesting;
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
        public void AllwaysPass()
        {
            ClipboardConverterCollection.Execute();
        }

        [TestMethod]
        public void Empty_Then_HasNoTextAndHtml()
        {
            // クリップボードが空の場合
            Clipboard.Clear();

            ClipboardConverterCollection.Execute();

            // テキストもHtml形式データも持たない
            Helper.CheckHasNoTextData();
            Helper.CheckHasNoHtmlData();
        }

        [TestMethod]
        public void TextStartWithHttp_Then_CreateHtmlFormatLink()
        {
            string text = "http.....";

            // httpから始まるテキストがある場合
            Clipboard.SetText(text, TextDataFormat.Text);

            ClipboardConverterCollection.Execute();

            // html形式のリンクが生成される
            Helper.CheckHasHtmlData();
            // テキストは変化しない
            Helper.CheckText(text);
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
        // Textデータが期待通りかチェックする
        public static void CheckText(string expect)
        {
            string actual = Clipboard.GetText(TextDataFormat.Text);
            Assert.AreEqual(actual, expect);
        }


        // Htmlデータの <!--StartFragment--> ～ <!--EndFragment--> の部分が期待通りかチェックする
        public static void CheckHtmlFragmentPart(string expect)
        {
            string raw_data = Clipboard.GetText(TextDataFormat.Html);
            string actual = extractFragmentPart(raw_data);
            Assert.AreEqual(actual, expect);
        }

        static string extractFragmentPart(string html)
        {
            string start_keyword = "<!--StartFragment-->";  // 開始キーワード
            string end_keyword = "<!--EndFragment-->";    // 終了キーワード
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
            Assert.AreEqual(Clipboard.ContainsData(DataFormats.Text), true);
        }

        public static void CheckHasNoTextData()
        {
            Assert.AreEqual(Clipboard.ContainsData(DataFormats.Text), false);
        }
        public static void CheckHasHtmlData()
        {
            Assert.AreEqual(Clipboard.ContainsData(DataFormats.Html), true);
        }

        public static void CheckHasNoHtmlData()
        {
            Assert.AreEqual(Clipboard.ContainsData(DataFormats.Html), false);
        }
    }
}

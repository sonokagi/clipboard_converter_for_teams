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
        public static void CheckText(string expect)
        {
            string actual = Clipboard.GetText(TextDataFormat.Text);
            Assert.AreEqual(actual, expect);
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

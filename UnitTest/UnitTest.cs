using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Windows.Forms;
using clipboard_converter_for_teams;

namespace UnitTest
{
    [TestClass]
    public class ClipboardConverterCollectionTest
    {
        [TestMethod]
        public void AllwaysPass()
        {
            ClipboardConverterCollection.Execute();
        }
    }

    [TestClass]
    public class ClipboardHelperLearningTest
    {
        //分かったこと
        //- htmlデータだけをクリップボードに設定することはできない
        //  plainText引数に null を渡しても、空のテキストが設定されてしまう

        [TestMethod]
        public void SetNullToText_Then_ClipboardHasEmptyText()
        {
            // テキストを null で指定した場合
            ClipboardHelper.CopyToClipboard("html", null);
            // テキストデータを持つ
            Assert.AreEqual(Clipboard.ContainsData(DataFormats.Text), true);
            // テキストデータは空になる
            string actual = Clipboard.GetText(TextDataFormat.Text);
            Assert.AreEqual(actual, String.Empty);
        }

        [TestMethod]
        public void SetEmptyToText_Then_ClipboardHasEmptyText()
        {
            // テキストを空で指定した場合
            ClipboardHelper.CopyToClipboard("html", String.Empty);
            // テキストデータを持つ
            Assert.AreEqual(Clipboard.ContainsData(DataFormats.Text), true);
            // テキストデータは空になる
            string actual = Clipboard.GetText(TextDataFormat.Text);
            Assert.AreEqual(actual, String.Empty);
        }
    }
}

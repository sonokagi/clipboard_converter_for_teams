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
    }
}

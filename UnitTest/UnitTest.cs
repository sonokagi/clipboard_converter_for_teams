using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using clipboard_converter_for_teams;

namespace UnitTest
{
    [TestClass]
    public class UnitTest
    {
        [TestMethod]
        public void AllwaysFail()
        {
            ClipboardConverterCollection.Execute();
            Assert.Fail("first step");
        }
    }
}

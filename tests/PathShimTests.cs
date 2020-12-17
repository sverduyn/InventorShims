﻿using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using InventorShims;

namespace PathShims_Tests
{
    [TestClass]
    public class UpOneLevel
    {

        [TestMethod]
        public void UpOneLevel_TrailingBackslash()
        {
            var stringA = @"C:\A\Test\String\";
            var stringB = @"C:\A\Test\";
            Assert.AreEqual(PathShims.UpOneLevel(stringA), stringB);
        }

        [TestMethod]
        public void UpOneLevel_NoTrailingBackslash()
        {
            var stringA = @"C:\A\Test\String";
            var stringB = @"C:\A\Test\";
            Assert.AreEqual(PathShims.UpOneLevel(stringA), stringB);
        }

        [TestMethod]
        public void UpOneLevel_NoBackslashes()
        {
            var stringA = @"C:";
            string stringB = "";
            Assert.AreNotEqual(PathShims.UpOneLevel(stringA), stringB);
        }

        [TestMethod]
        public void UpOneLevel_OneLevelWithTrailingBackslash()
        {
            var stringA = @"C:\";
            string stringB = "C:";
            Assert.AreEqual(PathShims.UpOneLevel(stringA), stringB);
        }

        [TestMethod]
        public void UpOneLevel_TwoLevelWithTrailingBackslash()
        {
            var stringA = @"C:\A\";
            var stringB = @"C:\";
            Assert.AreEqual(PathShims.UpOneLevel(stringA), stringB);
        }

    }


    [TestClass]
    public class IsLibraryPath
    {
        [TestMethod]
        public void GoodInput()
        {
            var getInventor = GetInventor.Instance;
            var app = GetInventor.Application;
            var test = app.DesignProjectManager.ActiveDesignProject.LibraryPaths[1].Path;

            Assert.IsTrue(PathShims.IsLibraryPath(test, ref app));
        }

        [TestMethod]
        public void GoodInput_NoTrailingSlash()
        {
            var getInventor = GetInventor.Instance;
            var app = GetInventor.Application;
            var test = app.DesignProjectManager.ActiveDesignProject.LibraryPaths[1].Path;
            test = PathShims.TrimTrailingSlash(test);

            Assert.IsTrue(PathShims.IsLibraryPath(test, ref app));
        }

        public void BadInput_WrongPath()
        {
            var getInventor = GetInventor.Instance;
            var app = GetInventor.Application;

            var test = @"C:\Windows";
            Assert.IsFalse(PathShims.IsLibraryPath(test, ref app));

        }

        [TestMethod]
        public void BadInput_Empty()
        {
            var getInventor = GetInventor.Instance;
            var app = GetInventor.Application;
            var test = "";

            Assert.IsFalse(PathShims.IsLibraryPath(test, ref app));
        }

        [TestMethod]
        public void BadInput_Null()
        {
            var getInventor = GetInventor.Instance;
            var app = GetInventor.Application;
            string test = null;

            Assert.IsFalse(PathShims.IsLibraryPath(test, ref app));
        }

    }

    [TestClass]
    public class IsContentCenterPath
    {
        [TestMethod]
        public void GoodInput()
        {
            var getInventor = GetInventor.Instance;
            var app = GetInventor.Application;
            string test = app.DesignProjectManager.ActiveDesignProject.ContentCenterPath;

            Assert.IsTrue(PathShims.IsContentCenterPath(test, ref app));
        }

        [TestMethod]
        public void GoodInput_NoTrailingSlash()
        {
            var getInventor = GetInventor.Instance;
            var app = GetInventor.Application;
            string test = app.DesignProjectManager.ActiveDesignProject.ContentCenterPath;
            test = PathShims.TrimTrailingSlash(test);

            Assert.IsTrue(PathShims.IsContentCenterPath(test, ref app));
        }


        [TestMethod]
        public void BadInput_Wrong()
        {
            var getInventor = GetInventor.Instance;
            var app = GetInventor.Application;
            string test = @"C:\Windows\";

            Assert.IsFalse(PathShims.IsContentCenterPath(test, ref app));
        }

        [TestMethod]
        public void BadInput_Empty()
        {
            var getInventor = GetInventor.Instance;
            var app = GetInventor.Application;
            var test = string.Empty;

            Assert.IsFalse(PathShims.IsContentCenterPath(test, ref app));
        }

        [TestMethod]
        public void BadInput_Null()
        {
            var getInventor = GetInventor.Instance;
            var app = GetInventor.Application;
            string test = null;

            Assert.IsFalse(PathShims.IsContentCenterPath(test, ref app));
        }

    }

    [TestClass]
    public class TrimTrailingSlash
    {
        [TestMethod]
        public void GoodInput_TrailingSlash()
        {
            string test = @"C:\Windows\SysWow\";
            string test2 = @"C:\Windows\SysWow";
            Assert.AreEqual(PathShims.TrimTrailingSlash(test), test2);
        }

        [TestMethod]
        public void GoodInput_NoTrailingSlash()
        {
            string test = @"C:\Windows\SysWow";
            string test2 = @"C:\Windows\SysWow";
            Assert.AreEqual(PathShims.TrimTrailingSlash(test), test2);
        }



    }
}

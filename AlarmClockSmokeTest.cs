using Microsoft.VisualStudio.TestTools.UnitTesting;
using OpenQA.Selenium;
using OpenQA.Selenium.Appium;
using OpenQA.Selenium.Appium.Windows;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium.Support.UI;
using System;
using System.Diagnostics;

namespace UnitTest1
{
    [TestClass]
    public class AlarmClockUnitTest
    {
        static WindowsDriver<WindowsElement> sessionAlarms;
        private static TestContext objTestContext;

        [ClassInitialize]
        public static void PrepareForTestingAlarms(TestContext testContext)
        {
            Debug.WriteLine("Hello ClassInitialize");

            AppiumOptions capCalc = new AppiumOptions();
            capCalc.AddAdditionalCapability("app", "Microsoft.WindowsAlarms_8wekyb3d8bbwe!App");

            sessionAlarms = new WindowsDriver<WindowsElement>(new Uri("http://127.0.0.1:4723"), capCalc);

            objTestContext = testContext;
        }

        [ClassCleanup]
        public static void CleanupAfterAllAlarmsTests()
        {
            Debug.WriteLine("Hello ClassCleanup");

            if (sessionAlarms != null)
            {
                sessionAlarms.Quit();
            }
        }

        [TestInitialize]
        public void BeforeATest()
        {
            Debug.WriteLine("Before a test, calling TestInitialize");
        }

        [TestCleanup]
        public void AfterATest()
        {
            Debug.WriteLine("After a test, calling TestCleanup");
        }

        [TestMethod]
        public void JustAnotherTest()
        {
            Debug.WriteLine("Hello Test");
        }

        [TestMethod]
        public void TestAlarmClockIsLaunchingSuccessfully()
        {
            Assert.AreEqual("Alarms & Clock", sessionAlarms.Title, false,
                 $"Actual title doesn't match expected title: {sessionAlarms.Title}");
        }

        [TestMethod, DataSource("System.Data.OleDb", @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=V:\Technologies\CS\UIAutomation\MSTestOverview\MyTest.xlsx;Extended Properties=""Excel 12.0 Xml;HDR = YES"";", "Clocks$", DataAccessMethod.Sequential)]
        public void VerifyNewClockCanBeAdded()
        {
            sessionAlarms.FindElementByAccessibilityId("ClockButton").Click();

            sessionAlarms.FindElementByName("Add new clock").Click();

            // System.Threading.Thread.Sleep(1000);

            WebDriverWait waitForMe = new WebDriverWait(sessionAlarms, TimeSpan.FromSeconds(10));

            var txtLocation = sessionAlarms.FindElementByName("Enter a location");

            waitForMe.Until(pred => txtLocation.Displayed);

            string locInput = Convert.ToString(objTestContext.DataRow["Location Input"]);

            Debug.WriteLine($"Input from Excel: {locInput}");

            txtLocation.SendKeys(locInput); //"Hanoi, Vietnam");

            txtLocation.SendKeys(Keys.Enter);

            var clockItems = sessionAlarms.FindElementsByAccessibilityId("WorldClockItemGrid");

            bool isClockTileFound = false;

            WindowsElement tileFound = null;

            string expectedValue = Convert.ToString(objTestContext.DataRow["Location Expected"]);

            Debug.WriteLine($"Expected Value: {expectedValue}");

            foreach (WindowsElement clockTile in clockItems)
            {
                if (clockTile.Text.StartsWith(expectedValue)) //"Hanoi"))
                {
                    isClockTileFound = true;
                    Debug.WriteLine("Clock Found!");
                    tileFound = clockTile;
                    break;
                }
            }

            Assert.IsTrue(isClockTileFound, "No clock tile found!");

            Actions actionForRightClick = new Actions(sessionAlarms);
            actionForRightClick.MoveToElement(tileFound);
            actionForRightClick.Click();
            actionForRightClick.ContextClick();
            actionForRightClick.Perform();

            AppiumOptions capDesktop = new AppiumOptions();
            capDesktop.AddAdditionalCapability("app", "Root");

            WindowsDriver<WindowsElement> sessionDesktop = 
                new WindowsDriver<WindowsElement>(new Uri("http://127.0.0.1:4723"), capDesktop);

            var contextItemDelete = sessionDesktop.FindElementByAccessibilityId("ContextMenuDelete");

            WebDriverWait desktopWaitForMe = new WebDriverWait(sessionAlarms, TimeSpan.FromSeconds(10));

            desktopWaitForMe.Until(pred => contextItemDelete.Displayed);

            contextItemDelete.Click();  
        }

    }
}

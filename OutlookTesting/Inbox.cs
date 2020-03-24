using NUnit.Framework;
using OpenQA.Selenium.Appium;
using OpenQA.Selenium.Appium.Windows;
using System;
using System.Reflection;
using System.Windows.Automation;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OpenQA.Selenium.Support.UI;
using System.Threading;
using System.Diagnostics;
using System.Runtime.InteropServices;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace OutlookTesting
{
    [TestFixture]
    public class Inbox
    {
        WindowsDriver<WindowsElement> _driver;
        Outlook.Application oApp = null;
        [SetUp]
        public void TestInit()
        {
            Process.Start("C://Program Files (x86)//Windows Application Driver//WinAppDriver.exe");
            var capabilities = new AppiumOptions();
            //capabilities.AddAdditionalCapability("app", "Microsoft.Office.OUTLOOK.EXE.15");
            capabilities.AddAdditionalCapability("app", @"C:\Program Files\Microsoft Office\root\Office16\OUTLOOK.EXE");
            capabilities.AddAdditionalCapability("platformName", "Windows");
            capabilities.AddAdditionalCapability("deviceName", "WindowsPC");
            _driver = new WindowsDriver<WindowsElement>(new Uri("http://127.0.0.1:4723"), capabilities);
            _driver.Manage().Timeouts().ImplicitWait=TimeSpan.FromSeconds(5);
        }

        [Test]

        public void PrintEmailDetails()
        {
            oApp = new Outlook.Application();
            // Get the MAPI namespace.
            Outlook.NameSpace oNS = oApp.GetNamespace("mapi");
            //Get the Inbox folder.
            Outlook.MAPIFolder oInbox = oNS.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderInbox);
            //Get the Items collection in the Inbox folder.
            Outlook.Items oItems = oInbox.Items;

            //Output some common properties.
            for (int i = oItems.Count; i > oItems.Count - 5; i--)
            {
                Microsoft.Office.Interop.Outlook.MailItem oMsg = (Microsoft.Office.Interop.Outlook.MailItem)oItems[i];
                Console.WriteLine("Subject: {0}", oMsg.Subject);
                Console.WriteLine("From: {0} <{1}>", oMsg.SenderName, oMsg.SenderEmailAddress);
                Console.WriteLine("To: {0}", oMsg.To);
                Console.WriteLine("ReceivedTime: {0}", oMsg.ReceivedTime);
                Console.WriteLine("Links: {0}", oMsg.Links);
                int AttachCnt = oMsg.Attachments.Count;
                Console.WriteLine("Attachments: " + AttachCnt.ToString());
                if (AttachCnt > 0)
                {
                    for (int j = 1; j <= AttachCnt; j++)
                        Console.WriteLine(j.ToString() + "-" + oMsg.Attachments[j].DisplayName);
                }
                Console.WriteLine("Size: {0} KB", oMsg.Size / 1024);

                Console.WriteLine("---------------------------------");
            }
        }
      

    [TearDown]
        public void TestCleanup()
        {
            if (_driver != null)
            {
                //Array.ForEach(Process.GetProcessesByName("Outlook"), x => x.Kill());
                _driver.Quit();
                _driver = null;
                Array.ForEach(Process.GetProcessesByName("WinAppDriver"), x => x.Kill());
            }
        }
    }
}


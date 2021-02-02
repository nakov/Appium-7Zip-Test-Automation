using System;
using System.IO;
using System.Threading;
using NUnit.Framework;
using OpenQA.Selenium;
using OpenQA.Selenium.Appium;
using OpenQA.Selenium.Appium.Windows;

namespace Appium_7Zip_Test_Automation
{
    public class Tests7Zip
    {
        private const string AppiumServerUrl = "http://[::1]:4723/wd/hub";
        private WindowsDriver<WindowsElement> driver;
        private WindowsDriver<WindowsElement> desktopDriver;
        private string workDir;

        [OneTimeSetUp]
        public void Setup()
        {
            var appiumOptions = new AppiumOptions() { PlatformName = "Windows" };
            appiumOptions.AddAdditionalCapability("app", @"C:\Program Files\7-Zip\7zFM.exe");
            driver = new WindowsDriver<WindowsElement>(new Uri(AppiumServerUrl), appiumOptions);

            var appiumOptionsDesktop = new AppiumOptions() { PlatformName = "Windows" };
            appiumOptionsDesktop.AddAdditionalCapability("app", "Root");
            desktopDriver = new WindowsDriver<WindowsElement>(
                new Uri(AppiumServerUrl), appiumOptionsDesktop);

            workDir = Directory.GetCurrentDirectory() + @"\workdir";
            if (Directory.Exists(workDir))
                Directory.Delete(workDir, true);
            Directory.CreateDirectory(workDir);
        }

        [Test]
        public void Test7Zip()
        {
            var textBoxLocationFolder = driver.FindElementByXPath("/Window/Pane/Pane/ComboBox/Edit");
            textBoxLocationFolder.SendKeys(@"C:\Program Files\7-Zip\");
            textBoxLocationFolder.SendKeys(Keys.Enter);

            //var listBoxFiles = driver.FindElementByXPath("/Window/Pane/List");
            //var listBoxFiles = driver.FindElementByClassName("SysListView32");
            var listBoxFiles = driver.FindElementByAccessibilityId("1001");
            listBoxFiles.SendKeys(Keys.Control + 'a');

            //var buttonAdd = driver.FindElementByXPath("/Window/ToolBar/Button[@Name='Add']");
            var buttonAdd = driver.FindElementByName("Add");
            buttonAdd.Click();

            // Wait until the [Add to Archive] dialog opens
            Thread.Sleep(500);
            var windowAddToArchive = desktopDriver.FindElementByName("Add to Archive");

            // Add all selected files to a new 7zip archive
            var textBoxArchiveName = windowAddToArchive.FindElementByXPath(
                "/Window/ComboBox/Edit[@Name='Archive:']");
            string archiveFileName = workDir + "\\" + DateTime.Now.Ticks + ".7z";
            textBoxArchiveName.SendKeys(archiveFileName);

            var comboArchiveFormat = windowAddToArchive.FindElementByXPath(
                "/Window/ComboBox[@Name='Archive format:']");
            comboArchiveFormat.SendKeys("7z");

            var comboCompressionLevel = windowAddToArchive.FindElementByXPath(
                "/Window/ComboBox[@Name='Compression level:']");
            comboCompressionLevel.SendKeys("Ultra");

            var comboDictionarySize = windowAddToArchive.FindElementByXPath(
                "/Window/ComboBox[@Name='Dictionary size:']");
            comboDictionarySize.SendKeys(Keys.End);

            var comboWordSize = windowAddToArchive.FindElementByXPath(
                "/Window/ComboBox[@Name='Word size:']");
            comboWordSize.SendKeys(Keys.End);

            var buttonAddToArchiveOK = windowAddToArchive.FindElementByXPath(
                "/Window/Button[@Name='OK']");
            buttonAddToArchiveOK.Click();

            // Wait while the 7zip archive is being created
            Thread.Sleep(1000);

            // Open the 7zip archive and extract it into the work directory
            textBoxLocationFolder.SendKeys(archiveFileName + Keys.Enter);

            var buttonExtract = driver.FindElementByName("Extract");
            buttonExtract.Click();

            //var buttonExtractOK = driver.FindElementByXPath("/Window/Window/Button[@Name='OK']");
            var buttonExtractOK = driver.FindElementByName("OK");
            buttonExtractOK.Click();

            // Wait for the "extract files" operation to complete
            Thread.Sleep(1000);

            // Assert the original files are the same as the extracted files
            string executable7ZipOriginal = @"C:\Program Files\7-Zip\7zFM.exe";
            string executable7ZipExtracted = workDir + @"\7zFM.exe";
            FileAssert.AreEqual(executable7ZipOriginal, executable7ZipExtracted);
        }

        [OneTimeTearDown]
        public void Shutdown()
        {
            driver.Quit();
            desktopDriver.Quit();
        }
    }
}
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using NPOI.XSSF.UserModel;
using NPOI.SS.UserModel;
using System.IO;
using PotwierdzeniaPrzesylek.SledzewnieService;
using OpenQA.Selenium;
using OpenQA.Selenium.IE;
using OpenQA.Selenium.Support.UI;

namespace PotwierdzeniaPrzesylek
{
    public partial class Form1 : Form
    {
        IWebDriver driver = new InternetExplorerDriver();
        

        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            
            IWorkbook xssfwb;
            using (FileStream file = new FileStream(@"C:\Paczki\rptKAN_PocztowaKsiazkaNadawcza 26 (1).xlsx", FileMode.Open, FileAccess.Read))
            {
                xssfwb = new XSSFWorkbook(file);
            }

            ISheet sheet = xssfwb.GetSheet("rptKAN_PocztowaKsiazkaNadawcza");
            for (int row = 1; row <= 5; row++) //sheet.LastRowNum; row++)
            {
                if (sheet.GetRow(row) != null) 
                {
                   // MessageBox.Show(string.Format("Row {0} = {1}", row, sheet.GetRow(row).GetCell(6).StringCellValue));

                    IRow wiersz = sheet.GetRow(row);

                    ICell cell = wiersz.GetCell(9) ?? wiersz.CreateCell(9,CellType.String);

                    ZapiszPNG(wiersz.GetCell(6).StringCellValue);

                    cell.SetCellValue("Zapisano w");
                }

            }

            using (FileStream file = new FileStream(@"C:\Paczki\rptKAN_PocztowaKsiazkaNadawcza 26 (1).xlsx", FileMode.Create, FileAccess.Write))
            {
                xssfwb.Write(file);
            }

        }
         
        private void ZapiszPNG(String numer)
        {
            driver.Navigate().GoToUrl("http://emonitoring.poczta-polska.pl/?numer=" + numer);

            IWebElement szukajButton = driver.FindElement(By.CssSelector("input[id=BSzukaj]"));
            if (IsElementPresent(driver, By.Id("BSzukajO")))
            {
                IWebElement szukajButton2 = driver.FindElement(By.CssSelector("input[id=BSzukajO]"));
                szukajButton2.Click();
                new WebDriverWait(driver, TimeSpan.FromSeconds(10)).Until(ExpectedConditions.ElementExists((By.Id("sledzenie_td"))));
            }
           
            Screenshot ss = ((ITakesScreenshot)driver).GetScreenshot();
            ss.SaveAsFile("C:\\Paczki\\Zrzuty\\ss"+numer+".png", OpenQA.Selenium.ScreenshotImageFormat.Png);
        }

        private void button3_Click(object sender, EventArgs e)
        {

        }

        private bool IsElementPresent(IWebDriver driver,By by)
        {
            try
            {
                driver.FindElement(by);
                return true;
            }
            catch (NoSuchElementException)
            {
                return false;
            }
        }

        private string pobierzDanePrzesylki(string numer)
        {
            SledzeniePortTypeClient service = new SledzeniePortTypeClient();
            service.ClientCredentials.UserName.UserName = "sledzeniepp";
            service.ClientCredentials.UserName.Password = "PPSA";

            Przesylka a = service.sprawdzPrzesylke(numer);

            return a.status.ToString();
            ///MessageBox.Show(a.status.ToString());
        }
    }
}

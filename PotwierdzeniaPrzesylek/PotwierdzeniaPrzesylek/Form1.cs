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
using NPOI.XWPF.UserModel;
using NPOI.Util;
using GemBox.Document;

namespace PotwierdzeniaPrzesylek
{
    public partial class Form1 : Form
    {
        IWebDriver driver = new InternetExplorerDriver();
        string sciezka = "C:\\paczki\\";
        string sciezkaplikuJpg = "C:\\paczki\\Zrzuty\\";

        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            IWorkbook xssfwb;
            string numer;
            string nrSprawy;
            using (FileStream file = new FileStream(sciezka + "ListaNumerow.xlsx", FileMode.Open, FileAccess.Read))
            {
                xssfwb = new XSSFWorkbook(file);
            }

            ISheet sheet = xssfwb.GetSheet("Arkusz1");
            for (int row = 1; row <= 5; row++) //sheet.LastRowNum; row++)
            {
                if (sheet.GetRow(row) != null)
                {
                    // MessageBox.Show(string.Format("Row {0} = {1}", row, sheet.GetRow(row).GetCell(6).StringCellValue));

                    IRow wiersz = sheet.GetRow(row);

                    numer = wiersz.GetCell(0).StringCellValue;
                    nrSprawy = wiersz.GetCell(1).StringCellValue.Replace('/', '_');

                    ZapiszPNG(numer);
                    zapiszPlikWord(nrSprawy, numer);

                    NPOI.SS.UserModel.ICell cell = wiersz.GetCell(9) ?? wiersz.CreateCell(9, CellType.String);

                    cell.SetCellValue("Zapisano");
                }

            }

            using (FileStream file = new FileStream(sciezka + "ListaNumerow.xlsx", FileMode.Create, FileAccess.Write))
            {
                xssfwb.Write(file);
            }

        }

        private void ZapiszPNG(String numer)
        {
            driver.Navigate().GoToUrl("http://emonitoring.poczta-polska.pl/?numer=" + numer);

            if (IsElementPresent(driver, By.Id("BSzukajO")))
            {
                IWebElement szukajButton2 = driver.FindElement(By.CssSelector("input[id=BSzukajO]"));
                szukajButton2.Click();
                new WebDriverWait(driver, TimeSpan.FromSeconds(10)).Until(ExpectedConditions.ElementExists((By.Id("wyniki"))));
            }

            var content = driver.FindElement(By.Id("wyniki"));
            TakeScreenshot(numer, driver, content);

            //Screenshot ss = ((ITakesScreenshot)driver).GetScreenshot();
            //ss.SaveAsFile(sciezkaplikuJpg + numer + ".jpeg", OpenQA.Selenium.ScreenshotImageFormat.Jpeg);
        }

        public void TakeScreenshot(string numer, IWebDriver driver, IWebElement element)
        {

            Byte[] byteArray = ((ITakesScreenshot)driver).GetScreenshot().AsByteArray;
            using (System.Drawing.Bitmap screenshot = new System.Drawing.Bitmap(new System.IO.MemoryStream(byteArray)))
            {
                Rectangle croppedImage = new Rectangle(element.Location.X, element.Location.Y, element.Size.Width, element.Size.Height);

                Rectangle all = new Rectangle(0, 0, screenshot.Width, screenshot.Height);

                croppedImage.Intersect(all);

                //croppedImage.Intersect(new Rectangle(0, 0, element.Size.Width, element.Size.Height));

                using (System.Drawing.Bitmap screenshot2 = screenshot.Clone(croppedImage, screenshot.PixelFormat))
                {
                    screenshot2.Save(sciezkaplikuJpg + numer + ".jpeg", System.Drawing.Imaging.ImageFormat.Jpeg);
                }
            }

        }


        private bool IsElementPresent(IWebDriver driver, By by)
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

        private void zapiszPlikWord2(string nazwaPlikuWord, string plikJpg)
        {

            var newFile2 = sciezka + nazwaPlikuWord + ".docx";
            using (var fs = new FileStream(newFile2, FileMode.Create, FileAccess.Write))
            {
                //var wDoc = new XWPFDocument();
                //var bytes = File.ReadAllBytes(sciezkaplikuJpg + plikJpg+".jpeg");
                //wDoc.AddPictureData(bytes, (int)NPOI.SS.UserModel.PictureType.JPEG);

                XWPFDocument doc = new XWPFDocument();
                var p0 = doc.CreateParagraph();
                p0.Alignment = ParagraphAlignment.CENTER;
                XWPFRun r0 = p0.CreateRun();

                var bytes = File.ReadAllBytes(sciezkaplikuJpg + plikJpg + ".jpeg");
                Stream stream = new System.IO.MemoryStream(bytes);
                r0.AddPicture(stream, (int)NPOI.SS.UserModel.PictureType.JPEG, "image1", Units.ToEMU(700), Units.ToEMU(800));

                doc.Write(fs);
            }
        }
        private void zapiszPlikWord(string nazwaPlikuWord, string plikJpg)
        {
            ComponentInfo.SetLicense("FREE-LIMITED-KEY");

            DocumentModel document = new DocumentModel();

            var section = new Section(document);

            PageSetup pageSetup = section.PageSetup;

            pageSetup.Orientation = GemBox.Document.Orientation.Landscape;

            document.Sections.Add(section);

            Paragraph paragraph = new Paragraph(document);
            section.Blocks.Add(paragraph);

            Picture picture1 = new Picture(document, sciezkaplikuJpg + plikJpg + ".jpeg");
            paragraph.Inlines.Add(new Run(document, nazwaPlikuWord));
            paragraph.Inlines.Add(picture1);

            document.Save(sciezka + nazwaPlikuWord + ".docx");
        }

    }
}

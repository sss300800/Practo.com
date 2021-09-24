using NPOI.HSSF.UserModel;
using NPOI.XSSF.UserModel;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Firefox;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium.Support.UI;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.Threading;

namespace Practo_mini
{

    class PractoProject
    {
       
        //For Read Data from Excel
        public static List<String>ReadDataFromExcel(String path)
        {
            XSSFWorkbook wb = new XSSFWorkbook(File.Open(path, FileMode.Open));
            XSSFSheet sh = (XSSFSheet)wb.GetSheetAt(0);
            XSSFRow row = (XSSFRow)sh.GetRow(0);
            XSSFCell cell = null;
            List<String> cell_values = new List<string>();
            int i, j;
            for (i = 1; i <= sh.LastRowNum; i++)
            {
                int cell_count = sh.GetRow(0).LastCellNum;
                for (j = 0; j < cell_count; j++)
                {
                    cell = (XSSFCell)sh.GetRow(i).GetCell(j);
                    String cell_value = cell.StringCellValue;
                    cell_values.Add(cell_value);
                }

            }
            return cell_values;
        }
        static void Main(string[] args)
        {
            Console.WriteLine("Hello");
            //Launch Chrome
            IWebDriver driver = new ChromeDriver("C:\\Users\\icon\\source\\repos");
            //Maximize the browser
            driver.Manage().Window.Maximize();
            //Launch Url(Open Google
            driver.Url = "https://www.practo.com/doctors";

            driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(20);
            String path = @"C:\Users\icon\source\repos\cities.xlsx";
            //WorkBook->Sheet->Row->Cell
            //.xls -> XSSFF


            List<String> cell_values = ReadDataFromExcel(path);
            //ENTETR CELL VALUES FROM EXCEL
            foreach (String cell_value in cell_values)
            {
                Console.WriteLine(cell_value);
            }
            //Thread.Sleep(3000);
            driver.FindElement(By.CssSelector("input.c-omni-searchbox")).Clear();
            Thread.Sleep(2000);
            driver.FindElement(By.XPath("//input[@data-qa-id='omni-searchbox-locality']")).SendKeys("Pune");
            Thread.Sleep(3000);
            driver.FindElement(By.XPath("//div[@data-qa-id='omni-suggestion-main' and text()='Pune']")).Click();
            //Thread.Sleep(2000);
            driver.FindElement(By.XPath("//input[@data-qa-id='omni-searchbox-keyword']")).Clear();
            driver.FindElement(By.XPath("//*[@id='c-omni-container']/div/div[2]/div/input")).SendKeys("hospital");
            Thread.Sleep(3000);

            driver.FindElement(By.XPath("//div[@data-qa-id='omni-suggestion-main' and text()='Hospital']")).Click();
            Thread.Sleep(3000);
            //INSTANTIATE FOR HOSPITAL LIST WITH RATING HIGHER THAN THREE AND  HOSPITAL SHOULD HAVE FULLY ACCREDITED
            ReadOnlyCollection<IWebElement> ListOfHospitals = driver.FindElements(By.XPath("//*[@id='container']/div[3]/div/div[2]/div[1]/div/div[2]/div/div/div[1]/div[2]/div/div[1]/div/div/span[1]"));
            List<String> hospitals = new List<String>();
            for (int i = 2; i < ListOfHospitals.Count; i++)
            {
                IWebElement star_rating = driver.FindElement(By.XPath("//*[@id='container']/div[3]/div/div[2]/div[1]/div/div[2]/div[" + i + "]/div/div[1]/div[2]/div/div[1]/div/div/span[1]"));
                String[] stars = star_rating.Text.Split('.');
                int star_value = Convert.ToInt16(stars[0]);
                Console.WriteLine(star_value);
                if (star_value > 3)
                {
                    String hospital_name = driver.FindElement(By.XPath("//*[@id='container']/div[3]/div/div[2]/div[1]/div/div[2]/div[" + i + "]//h2")).Text;
                    Console.WriteLine("Top 5 Hospital Search result for the: " + stars[0] + "are:");
                    hospitals.Add(hospital_name);
                }

            }
            //PRINTING HOSPITAL LIST
            for (int i = 0; i < 5; i++)
            {
                Console.WriteLine(hospitals[i]);
            }
            Console.WriteLine("Staus of Search Result for city Pune" + hospitals[0] + "Pass");
            driver.Quit();
            Console.ReadLine();
        }
        
    }
    
}
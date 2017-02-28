using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OpenQA.Selenium;
using OpenQA.Selenium.Support.UI;
using System.IO;
using System.Collections.ObjectModel;
using System.Text.RegularExpressions;
using System.Xml;
using Excel = Microsoft.Office.Interop.Excel;

namespace LVDownloader
{
    class Program
    {

        static void Main(string[] args)
        {

            List<string> Vals = reader("D:\\Users\\DSIYANCHEV\\Telegram Desktop\\input_formatted.csv"); // список всех регистрационных номеров, очищенный
            //List<string> ValsError = reader("D:\\Users\\DSIYANCHEV\\Downloads\\Telegram Desktop\\input_full.csv"); // список всех регистрационных номеров, неочищенный 

            List<List<string>> output = new List<List<string>>();

            XmlDocument MainDoc = new XmlDocument();
            XmlTextWriter Writer = new XmlTextWriter("D:\\Users\\DSIYANCHEV\\Telegram Desktop\\output.xml", Encoding.UTF8);
            Writer.WriteStartDocument();
            Writer.WriteStartElement("head");
            Writer.WriteEndElement();
            Writer.Close();

            MainDoc.Load("D:\\Users\\DSIYANCHEV\\Telegram Desktop\\output.xml");


            string urlMain = "http://company.lursoft.lv/ru/"; // основная часть ссылки, к ней будут добавляться регистрационные номера

            Console.SetBufferSize(Console.BufferWidth, 1000);   // а вот это для длины консоли

            List<int> statistics = new List<int>();

            IWebDriver DriverChrome = new OpenQA.Selenium.Chrome.ChromeDriver();

            for (int i = 1; i < Vals.Count; i++)
            {

                string CurrentUrl = urlMain + Vals[i];
                DriverChrome.Navigate().GoToUrl(CurrentUrl);
                int statisticsCount = 0;

                ReadOnlyCollection<IWebElement> ListTD = DriverChrome.FindElements(By.TagName("td"));


                if (ListTD.Count > 0 && ListTD[0].Text.Equals("Network Error (tcp_error)")) // вылетела ошибка TCP(сервер перегружен)
                {
                    statisticsCount++;
                    i--;
                    continue;
                }
                else if (ListTD.Count == 0) // не найдено
                {
                    statistics.Add(statisticsCount);
                    List<string> CurrentList = new List<string>();

                    CurrentList.Add("-");               // RecordID
                    CurrentList.Add("-");               // Name
                    CurrentList.Add("-");               // OrgForm
                    //CurrentList.Add(ValsError[i]);      // RegNum
                    CurrentList.Add("-");               // RegDate
                    CurrentList.Add("-");               // SEPA
                    CurrentList.Add("-");               // NDSNum
                    CurrentList.Add("-");               // IsActual
                    CurrentList.Add("-");               // Address
                    CurrentList.Add("-");               // RegisterNo
                    CurrentList.Add("-");               // RegisterDate
                    CurrentList.Add("-");               // LastUpdate
                    CurrentList.Add("-");               // Website
                    CurrentList.Add("-");               // Email
                    CurrentList.Add("-");               // Phone
                    CurrentList.Add("-");               // Fax
                    CurrentList.Add("false");           // IsFound

                    output.Add(CurrentList);
                }
                else // найдено
                {
                    statistics.Add(statisticsCount);
                    string Name = "", OrgForm = "", RegNum = "", RegDate = "", Sepa = "", NDSNum = "",
                        IsActual = "", Address = "", RegisterNo = "", RegisterDate = "", LastUpdate = "", Website = "", Email = "", Phone = "", Fax = "", IsFound = "true";
                    Regex NDS = new Regex("LV\\d{8,15}");
                    Match match;

                    ReadOnlyCollection<IWebElement> ListImg = DriverChrome.FindElements(By.TagName("img"));
                    bool ImgFlag = false;

                    for (int j = 0; j < ListImg.Count; j++)
                    {
                        if (ListImg[j].GetAttribute("alt").ToString().Equals("Активный"))
                        {
                            IsActual = "true";
                            ImgFlag = true;
                        }
                        else if (ListImg[j].GetAttribute("alt").ToString().Equals(" PVN_ne"))
                        {
                            IsActual = "false";
                            ImgFlag = true;
                        }
                    }

                    if (ImgFlag == false)
                    {
                        IsActual = "-";
                    }
                    List<string> matches = new List<string>();
                    bool regFlag = false;
                    for (int j = 0; j < ListTD.Count; j++)
                    {

                        match = NDS.Match(ListTD[j].Text);
                        matches.Add(match.Value.ToString());


                        switch (ListTD[j].Text)
                        {
                            case "Название":
                                string tmp5 = ListTD[++j].Text;
                                int Position = tmp5.IndexOf('П');
                                if (Position == -1)
                                {
                                    Name = tmp5;
                                }
                                else
                                {
                                    Name = tmp5.Substring(0, Position);
                                }
                                Name = Name.Replace('\n', ' ');

                                break;
                            case "Данные из реестра плательщиков НДС":
                                string tmp4 = ListTD[++j].Text;
                                if (tmp4.Equals("Нет"))
                                {
                                    NDSNum = "-";
                                    IsActual = "-";
                                }
                                break;
                            case "Правовая форма":
                                OrgForm = ListTD[++j].Text;
                                break;
                            case "Регистрационный номер, дата":
                                string[] tmp = ListTD[++j].Text.Split(',');
                                RegNum = tmp[0];
                                RegDate = tmp[1].Substring(1, tmp[1].Length - 1);
                                break;
                            case "Идентификатор SEPA":
                                Sepa = ListTD[++j].Text;
                                break;
                            case "Юридический адрес":
                                string tmp2 = ListTD[++j].Text;
                                int Position2 = tmp2.IndexOf('П');
                                if (Position2 == -1)
                                {
                                    Address = tmp2;
                                }
                                else
                                {
                                    Address = tmp2.Substring(0, Position2);
                                }
                                Address.Replace('\n', ' ');
                                break;
                            case "Регистрационное удостоверение":
                                regFlag = true;
                                string[] tmp3 = ListTD[++j].Text.Split(' ');
                                RegisterNo = tmp3[0] + " " + tmp3[1];
                                RegisterDate = tmp3[2];
                                break;
                            case "Последнее обновление в Регистре Предприятий":
                                LastUpdate = ListTD[++j].Text;
                                break;
                        }
                        if (!regFlag)
                        {
                            RegisterNo = "-";
                            RegisterDate = "-";
                        }

                    }

                    for (int j = 0; j < matches.Count; j++)
                    {
                        if (!matches[j].Equals(""))
                        {
                            NDSNum = matches[j];
                            break;
                        }

                        if (j == matches.Count - 1)
                        {
                            NDSNum = "-";
                        }
                    }

                    DriverChrome.SwitchTo().Frame(0);

                    ReadOnlyCollection<IWebElement> ListPhones = DriverChrome.FindElements(By.ClassName("vizitka_contact_phone"));
                    IWebElement Web = DriverChrome.FindElement(By.ClassName("vizitka_contact_web"));

                    if (ListPhones[0].Text.IndexOf('+') == -1)
                    {
                        Phone = "-";
                    }
                    else
                    {
                        Phone = ListPhones[0].Text.Substring(ListPhones[0].Text.IndexOf('+'));
                    }

                    if (ListPhones[1].Text.IndexOf('+') == -1)
                    {
                        Fax = "-";
                    }
                    else
                    {
                        Fax = ListPhones[1].Text.Substring(ListPhones[1].Text.IndexOf('+'));
                    }

                    string[] Webs = Web.Text.Split('\n');

                    if (Webs[0].Equals("Добавь адрес сайта\r"))
                    {
                        Email = "-";
                    }
                    else
                    {
                        Email = Webs[0];
                    }

                    if (Webs[1].Equals("Добавь адрес эл. почты"))
                    {
                        Website = "-";
                    }
                    else
                    {
                        Website = Webs[1];
                    }

                    List<string> CurrentList = new List<string>();

                    CurrentList.Add(i.ToString());          // RecordID
                    CurrentList.Add(Name);                  // Name
                    CurrentList.Add(OrgForm);               // OrgForm
                    CurrentList.Add(RegNum);                // RegNum
                    CurrentList.Add(RegDate);               // RegDate
                    CurrentList.Add(Sepa);                  // SEPA
                    CurrentList.Add(NDSNum);                // NDSNum
                    CurrentList.Add(IsActual);              // IsActual
                    CurrentList.Add(Address);               // Address
                    CurrentList.Add(RegisterNo);            // RegisterNo
                    CurrentList.Add(RegisterDate);          // RegisterDate
                    CurrentList.Add(LastUpdate);            // LastUpdate
                    CurrentList.Add(Website);               // Website
                    CurrentList.Add(Email);                 // Email
                    CurrentList.Add(Phone);                 // Phone
                    CurrentList.Add(Fax);                   // Fax
                    CurrentList.Add(IsFound);               // IsFound

                    output.Add(CurrentList);
                }
            }

            Excel.Application ExcelApp;
            ExcelApp = new Excel.Application();
            ExcelApp.Visible = true;
            ExcelApp.SheetsInNewWorkbook = 1;
            ExcelApp.Workbooks.Add(Type.Missing);

            Excel.Workbooks WorkBooks = ExcelApp.Workbooks;
            Excel.Workbook Workbook = WorkBooks[1];
            Excel.Sheets ExcelSheets = Workbook.Worksheets;
            Excel.Worksheet CurrentSheet = (Excel.Worksheet)ExcelSheets.get_Item(1);

            Excel.Range cell;
            cell = CurrentSheet.get_Range("A1", Type.Missing);
            cell = cell.get_Offset(1, 0);

            for (int i = 0; i < output.Count; i++)
            {
                for (int j = 0; j < output[i].Count; j++)
                {
                    cell.Value2 = output[i][j];
                    cell = cell.get_Offset(0, 1);
                }
                cell = cell.get_Offset(1, 0);
            }

            for (int i = 0; i < output.Count; i++)
            {
                XmlNode Element = MainDoc.CreateElement("BusinessPartner");
                MainDoc.DocumentElement.AppendChild(Element);

                XmlNode RecordID = MainDoc.CreateElement("RecordID");
                RecordID.InnerText = output[i][0];
                Element.AppendChild(RecordID);

                XmlNode Name = MainDoc.CreateElement("Name");
                Name.InnerText = output[i][1];
                Element.AppendChild(Name);

                XmlNode OrgForm = MainDoc.CreateElement("OrgForm");
                OrgForm.InnerText = output[i][2];
                Element.AppendChild(OrgForm);

                XmlNode RegNum = MainDoc.CreateElement("RegNum");
                RegNum.InnerText = output[i][3];
                Element.AppendChild(RegNum);

                XmlNode RegDate = MainDoc.CreateElement("RegDate");
                RegDate.InnerText = output[i][4];
                Element.AppendChild(RegDate);

                XmlNode Sepa = MainDoc.CreateElement("Sepa");
                Sepa.InnerText = output[i][5];
                Element.AppendChild(Sepa);

                XmlNode NDSNum = MainDoc.CreateElement("NDSNum");
                NDSNum.InnerText = output[i][6];
                Element.AppendChild(NDSNum);

                XmlNode IsActual = MainDoc.CreateElement("IsActual");
                IsActual.InnerText = output[i][7];
                Element.AppendChild(IsActual);

                XmlNode Address = MainDoc.CreateElement("Address");
                Address.InnerText = output[i][8];
                Element.AppendChild(Address);

                XmlNode RegisterNo = MainDoc.CreateElement("RegisterNo");
                RegisterNo.InnerText = output[i][9];
                Element.AppendChild(RegisterNo);

                XmlNode RegisterDate = MainDoc.CreateElement("RegisterDate");
                RegisterDate.InnerText = output[i][10];
                Element.AppendChild(RegisterDate);

                XmlNode LastUpdate = MainDoc.CreateElement("LastUpdate");
                LastUpdate.InnerText = output[i][11];
                Element.AppendChild(LastUpdate);

                XmlNode Website = MainDoc.CreateElement("Website");
                Website.InnerText = output[i][12];
                Element.AppendChild(Website);

                XmlNode Email = MainDoc.CreateElement("Email");
                Email.InnerText = output[i][13];
                Element.AppendChild(Email);

                XmlNode Phone = MainDoc.CreateElement("Phone");
                Phone.InnerText = output[i][14];
                Element.AppendChild(Phone);

                XmlNode Fax = MainDoc.CreateElement("Fax");
                Fax.InnerText = output[i][15];
                Element.AppendChild(Fax);

                XmlNode IsFound = MainDoc.CreateElement("IsFound");
                IsFound.InnerText = output[i][16];
                Element.AppendChild(IsFound);
            }

            MainDoc.Save("D:\\Users\\DSIYANCHEV\\Telegram Desktop\\output.xml");



        }

        private static List<string> reader(string path)
        {
            List<string> output = new List<string>();

            string[] allLines = System.IO.File.ReadAllLines(@path);

            for (int i = 1; i < allLines.Length; i++)
            {
                string[] divided = allLines[i].Split(' ');
                output.Add(divided[1]);
            }

            return output;
        }




    }
}

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

namespace LVDownloader
{
    class Program
    {

        static void Main(string[] args)
        {

            List<string> Vals = reader("D:\\Users\\DSIYANCHEV\\Downloads\\Telegram Desktop\\input_formatted.csv"); // список всех регистрационных номеров, очищенный
            //List<string> ValsError = reader("D:\\Users\\DSIYANCHEV\\Downloads\\Telegram Desktop\\input_full.csv"); // список всех регистрационных номеров, неочищенный 

            List<List<string>> output = new List<List<string>>();





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
                    string Name, OrgForm, RegNum, RegDate, Sepa, NDSNum, IsActual, Address, RegisterNo, RegisterDate, LastUpdate, Website, Email, Phone, Fax;
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
                    for (int j = 0; j < ListTD.Count; j++)
                    {

                        match = NDS.Match(ListTD[j].Text);
                        matches.Add(match.Value.ToString());
                        bool regFlag = false;

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

                    for (int j = 0; j < matches.Count; j ++)
                    {
                        if (!matches[j].Equals(""))
                        {
                            NDSNum = matches[i];
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
                        Website = "-";
                    }
                    else
                    {
                        Website = Webs[0];
                    }

                    if (Webs[1].Equals("Добавь адрес эл. почты"))
                    {
                        Fax = "-";
                    }
                    else
                    {
                        Fax = Webs[1];
                    }

                }




            }

            //IWebDriver Browser = new OpenQA.Selenium.Chrome.ChromeDriver();
            //Browser.Navigate().GoToUrl("http://company.lursoft.lv/ru/50003291221");

            //System.Collections.ObjectModel.ReadOnlyCollection<IWebElement> list = Browser.FindElements(By.TagName("td"));

            //ReadOnlyCollection<IWebElement> ListImg = Browser.FindElements(By.TagName("img"));

            //for (int j = 0; j < ListImg.Count; j++)
            //{
            //    Console.WriteLine(ListImg[j].GetAttribute("alt"));
            //    if (ListImg[j].Text == "Активный")
            //    {
            //        Console.WriteLine("Science, beach!");
            //    }
            //}


            //Browser.Manage().Timeouts().PageLoad = new TimeSpan(0, 1, 0);
            //for (int i = 0; i < list.Count; i++)
            //{
            //    Console.WriteLine(i.ToString() + " " + list[i].Text);
            //}

            //Console.WriteLine("---");

            //Browser.SwitchTo().Frame(0);

            //IWebElement element = Browser.FindElement(By.ClassName("vizitka_contact_phone"));

            //Console.WriteLine(element.Text);
            //Console.ReadKey();
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

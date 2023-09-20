using Microsoft.Extensions.Configuration;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using iTextSharp.text;
using iTextSharp.text.pdf;
using System.Net;
using System.Net.Mail;

namespace SalaryEmplV2
{
    class Program
    {
        #region -- GLOBAL VARIABLES --
        public static List<Employee> employeeList = new List<Employee>();
        public static List<TimeSchedule> employeeTimeList = new List<TimeSchedule>();
        public static List<string> ErrorList = new List<string>();
        public static string toDoPath = "";
        public static string donePath = "";
        public static string errorsPath = "";
        public static string resultsPath = "";
        public static string FileName = "";
        public static string FilePath = "";
        public static string myPDFFile = "";
        #endregion
        static void Main(string[] args)
        {
                ClearBeforeNextIteration();
                var configuration = UploadIniFile();
                getFolders(configuration);
                getEmployees(configuration);
                ReadCSVFile(toDoPath);
                var ListOfResults=CalculateSalary();
                LogErrors(errorsPath);
                WriteDataToPDFFile(ListOfResults);
                //EmailSending();
                System.Threading.Thread.Sleep(5000);
        }
        
        #region -- MAIN LOGIC --
        static void EmailSending()
        {
            string smtpServer = "your.smtp.server";
            int smtpPort = 1;//desired Port
            string smtpUsername = "your@gmail.com";
            string smtpPassword = "yourSMTPpassword";
            string senderEmail = "your@gmail.com";
            string recipientEmail = "desired@gmail.com";
            string subject = "Desired Subject";
            string body = "";

            byte[] fileBytes = File.ReadAllBytes(myPDFFile);

            MailMessage message = new MailMessage(senderEmail, recipientEmail, subject, body);

            MemoryStream pdfStream = new MemoryStream(fileBytes);
            string myPDFName = $"Wyplaty za okres: {FileName.Replace("From", "Od").Replace("to", "do")}.pdf";
            Attachment pdfAttachment = new Attachment(pdfStream,myPDFName);
            message.Attachments.Add(pdfAttachment);

            SmtpClient smtpClient = new SmtpClient(smtpServer, smtpPort);
            smtpClient.Credentials = new NetworkCredential(smtpUsername, smtpPassword);
            smtpClient.EnableSsl = true;
            smtpClient.Send(message);

            pdfAttachment.Dispose();
            pdfStream.Dispose();
        }
        static void ClearBeforeNextIteration()
        {
            employeeList.Clear();
            employeeTimeList.Clear();
            ErrorList.Clear();
            toDoPath = "";
            donePath = "";
            errorsPath = "";
            resultsPath = "";
            FileName="";
            FilePath = "";
        }
        static void WriteDataToPDFFile(List<EmployeeResult> listToWrite)
        {
            Document document = new Document();
            myPDFFile = resultsPath + @"\" + "Wypłata " + FileName.Replace('/', '-')+".pdf";
            PdfWriter.GetInstance(document, new FileStream(myPDFFile, FileMode.Create));
            document.Open();
            Font headerFont = FontFactory.GetFont(FontFactory.HELVETICA_BOLD);
            PdfPCell header4 = new PdfPCell(new Phrase("Pracownik", headerFont));
            PdfPCell header5 = new PdfPCell(new Phrase("Wyplata", headerFont));
            PdfPCell header6 = new PdfPCell(new Phrase("Ilosc godzin", headerFont));

            document.Add(new Paragraph($"Utworzono:{DateTime.Now}"));
            document.Add(new Paragraph($"Wyplaty za okres: {FileName.Replace("From","Od").Replace("to","do")}"));
            Paragraph paragraph = new Paragraph("");
            paragraph.SpacingAfter = 10f;
            document.Add(paragraph);

            PdfPTable table = new PdfPTable(3);
            table.AddCell(header4);
            table.AddCell(header5);
            table.AddCell(header6);
            foreach (var line in listToWrite)
            {
                table.AddCell(line.name);
                table.AddCell(line.salary.ToString()+"zl");
                table.AddCell(line.totalHours.ToString());
            }
            document.Add(table);

            List<string> lines = ErrorList.Distinct().ToList();
            string messageString = string.Join(Environment.NewLine, lines);
            document.Add(new Paragraph("Lista bledow:"));
            Paragraph ErrorParagraph = new Paragraph(messageString);
            document.Add(ErrorParagraph); ;

            document.Close();
        }
        static List<EmployeeResult> CalculateSalary()
        {
            List<EmployeeResult> employeeResultList = new List<EmployeeResult>();
            foreach(var employee in employeeList)
            {
                List <DateTime> dateListOfEmployee= new List<DateTime>();
                List<string> timeTable = employeeTimeList.Where(x => x.Name == employee.Name).Select(x => x.dateInfo).ToList();
                foreach(var date in timeTable)
                    dateListOfEmployee.Add(DateTime.Parse(date));
                TimeSpan TotalTime = CalculationsLogic(dateListOfEmployee,employee.Name);

                double salary=CountSalary(employee, TotalTime);
                EmployeeResult employeeResult = new EmployeeResult(TotalTime, employee.Name,salary);
                employeeResultList.Add(employeeResult);
            }
            return employeeResultList;          
        }
        static double CountSalary(Employee employee,TimeSpan totalHours)
        {
            double salary = Math.Round(totalHours.TotalHours * employee.Salary,2);
            return salary;
        }
        static TimeSpan CalculationsLogic(List<DateTime> listToCalculate,string employeeName)
        {
            TimeSpan TotalDayHour=new TimeSpan();
            var checkList = checkIfTwoACtionsPerDay(listToCalculate);
            if (checkList.Count==0)
            {
                DateTime checkDate = new DateTime();
                foreach (var day in listToCalculate)
                {
                    if (checkDate.Date != day.Date)
                    {
                        var countTime = listToCalculate.Where(x => x.Date == day.Date).Select(x => x).ToList();
                        DateTime startTime = countTime.Min();
                        DateTime stopTime = countTime.Max();
                        TimeSpan calculatedTime= stopTime-startTime;
                        TotalDayHour = TotalDayHour + calculatedTime;
                    }
                    checkDate = day.Date;
                }
                return TotalDayHour;
            }
            else
            {
                foreach (var errorLine in checkList)
                    ErrorList.Add($"{employeeName} - brak rozpoczęcia/zakończenia pracy w dniu {errorLine}");
                TimeSpan Empty = new TimeSpan();
                return Empty;
            }
        }
        static List<DateTime> checkIfTwoACtionsPerDay(List<DateTime> listToValidate)
        {
            List<DateTime> checking = new List<DateTime>();
            foreach(var day in listToValidate)
            {
                var checkQty = listToValidate.Where(x => x.Date == day.Date).Select(x => x.Date).ToList();
                if (checkQty.Count() != 2)
                    checking.Add(checkQty.First());
            }
            return checking;
        }
        static void ReadCSVFile(string path)
        {
            if (Directory.Exists(toDoPath))
            {
                string[] files = Directory.GetFiles(toDoPath);
                foreach(string file in files)
                {
                    FilePath = file;
                    foreach(var line in File.ReadLines(file))
                    {
                        string[] lineDetails =  line.Split(','); 
                        Employee? find = null;
                        try
                        {
                             find= employeeList.Find(x => x.Name == lineDetails[1]);
                        }
                        catch (Exception ex){ Console.WriteLine(ex); }
                            if (find!=null)
                            {
                                TimeSchedule EmployeeDayData = new TimeSchedule(lineDetails[1], lineDetails[2]);
                                employeeTimeList.Add(EmployeeDayData);
                            }
                            if (lineDetails[0].Contains("From"))
                                FileName = lineDetails[0];
                    }
                }  
            }
            else
                LogErrors($"Błąd - ścieżka {toDoPath} nie istnieje");
        }
        #endregion

        #region -- PREPARE INI INFO --
        static IConfigurationRoot UploadIniFile()
        {
            var builder = new ConfigurationBuilder().AddIniFile("config.ini");
            var configuration = builder.Build();
            return configuration;
        }
        static void getFolders(IConfigurationRoot configuration)
        {
            toDoPath = configuration.GetSection("FOLDERS")["TODO"];
            donePath = configuration.GetSection("FOLDERS")["DONE"];
            errorsPath = configuration.GetSection("FOLDERS")["ERROR"];
            resultsPath = configuration.GetSection("FOLDERS")["RESULTS"];
        }
        static void getEmployees(IConfigurationRoot configuration)
        {
            try
            {
                foreach (var field in configuration.GetChildren())
                {
                    Employee newEmployee = new Employee(field["Dane"], Convert.ToDouble(field["Stawka"]));
                    if (newEmployee.Name != null)
                        employeeList.Add(newEmployee);
                }
            }
            catch (Exception ex)
            {
                ErrorList.Add("Błąd - " + ex.Message);
            }
        }
        #endregion
        static void LogErrors(string path)
        {
            DateTime date = DateTime.Now;
            string dateString = date.ToString("dd-MM-yyyy");
            string fileName = $"Errors_{dateString}.txt";
            string filePath = path+@"\"+fileName;
            if (ErrorList.Count != 0)
            {
                List<string> lines = ErrorList.Distinct().ToList();
                string messageString = string.Join(Environment.NewLine, lines);
                File.WriteAllText(filePath,messageString);
            }
            else
            {
                string myDoneFileName = ($"Wypłaty-obliczone {FileName.Replace('/','-')}.csv");
                string myNewParth = donePath + @"\" + myDoneFileName;
            }
        }
    }
   #region -- CLASSES --
   class Employee
    {
        public string Name { get; set; }
        public double Salary { get; set; }
        public Employee(string name, double salary)
        {
            Name = name;
            Salary = salary;
        }
    }
   class TimeSchedule
    {
        public string Name { get; set; }
        public string dateInfo { get; set; }
        public TimeSchedule(string name, string DateInfo)
        {
            Name = name;
            dateInfo = DateInfo;
        }
    }
   class DayParse
    {
       public DateTime date { get; set; }
       public DateTime hour { get; set; }
        public DayParse(DateTime Date)
        {
            this.date = Date.Date;
            this.hour = default(DateTime).Add(Date.TimeOfDay);
        }
    }
   class EmployeeResult
    {
        public double totalHours { get; set; }
        public string name { get; set; }
        public double salary { get; set; }

        public EmployeeResult(TimeSpan totalHours, string name, double Salary)
        {
            this.totalHours = Math.Round(Convert.ToDouble(totalHours.TotalHours),2);
            this.salary = Salary;
            this.name = name;
        }
    }
   #endregion
}

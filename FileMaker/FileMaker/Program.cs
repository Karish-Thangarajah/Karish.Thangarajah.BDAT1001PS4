

using DocumentFormat.OpenXml;
using FileMaker.Models;
using FileMaker.Models.Api;
using FileMaker.Models.Utilities;
using FileMaker.Models.Utilities.DocMakers;
using FTPApp.Models.Utilities;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Threading.Tasks;

namespace FileMaker
{
    class Program
    {
        async static Task Main(string[] args)
        {
            //Acquire an Access Token that comes bundled into a ApiResult Json object
            string authResult = await HTTP.AcquireAccessToken(Constants.Urls.BaseUrlApi + Constants.Urls.PostToken,
                                    new
                                    {
                                        UserName = Constants.Credentials.UserName,
                                        Password = Constants.Credentials.Password
                                    });

            //Convert the authResult string to a ApiResult object using Newtonsoft Json Converter
            ApiResult<JwtResult> apiResult = JsonConvert.DeserializeObject<ApiResult<JwtResult>>(authResult);

            string apigetResult = await HTTP.Get("https://webapibasicsstudenttracker.azurewebsites.net/api/securestudents", apiResult.Data.AccessToken);

            ApiResult<List<Student>> apiStudentListResult = JsonConvert.DeserializeObject<ApiResult<List<Student>>>(apigetResult);

            Student me = apiStudentListResult.Data.Find(x => x.StudentCode == "200471940");
            me.MyRecord = true;

            Student findMe = apiStudentListResult.Data.Find(x => x.MyRecord == true);

            Console.WriteLine(apiStudentListResult.Data.Count);



            //WordMaker.CreateWordprocessingDocument(Constants.Local.localPathDefault + Constants.Filename.WordDoc);

            //{
            //    foreach (Student student in apiStudentListResult.Data)
            //    {
            //        WordMaker.OpenAndAddText(Constants.Local.localPathDefault + Constants.Filename.WordDoc, student);

            //    }
            //}

            //Console.WriteLine(FTP.UploadFile(@"C:\Users\karis\OneDrive\Desktop\BDAT1001\Assignment 4\info.docx", Constants.FTP.MyUrl + "/info.docx"));
            //if (FTP.FileExists(Constants.FTP.MyUrl + "/info.docx"))
            //{
            //    Console.WriteLine("Success!");
            //}

            /*
             Word is done, below is the excel stuff
             */

            //StreamReader reader = new StreamReader(@"C:\Users\karis\OneDrive\Desktop\BDAT1001\Assignment 4\students.json");

            //string jsonStr = reader.ReadToEnd();
            //List<Student> students = JsonConvert.DeserializeObject<List<Student>>(jsonStr);

            ExcelMaker.CreateSpreadsheetWorkbook(Constants.Local.localPathDefault + Constants.Filename.ExcelDoc);
            
            ExcelMaker.InsertTextSheet(@"C:\Users\karis\OneDrive\Desktop\BDAT1001\Assignment 4\info.xlsx", $"Hello, my name is {me.FirstName} {me.LastName}",1 ,"A", 2);


            List<string> columns = new List<string> { "A", "B", "C", "D", "E", "F", "G" };
            List<string> columnNames = new List<string> { "StudentId", "StudentCode", "FirstName", "LastName", "DateOfBirth", "IsMe", "Age" };

            for (int i = 0; i < columns.Count; i++)
            {
                ExcelMaker.InsertTextSheet(@"C:\Users\karis\OneDrive\Desktop\BDAT1001\Assignment 4\info.xlsx", columnNames[i], 2, columns[i], 1);
            }
        }
    }
}

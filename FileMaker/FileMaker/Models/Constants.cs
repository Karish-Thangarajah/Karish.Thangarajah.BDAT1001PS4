using System;
using System.Collections.Generic;
using System.Text;

namespace FileMaker.Models
{
    public class Constants
    {
        public class Credentials
        {
            public const string UserName = "karishg220@gmail.com";
            public const string Password = "A$df7890";
        }

        public class Urls
        {
            public const string BaseUrl = "https://webapibasicsstudenttracker.azurewebsites.net";

            public static readonly string BaseUrlApi = $"{BaseUrl}/api";

            public const string GetStudentsUnsecure = "/students";
            public const string GetStudentsSecure = "/securestudents";

            public const string PostToken = "/Tokens";

        }

        public class HTTP
        {
            public const int Timeout = 60000; //60 seconds, a long time if needed
            public const string ContentType = "application/json";

            public class Security
            {
                public const string AuthHeader = "Authorization";
                public const string AuthMethod = "Bearer";
            }
        }

        public class Local
        {
            public const string localPathDefault = @"C:\Users\karis\OneDrive\Desktop\BDAT1001\Assignment 4";
        }

        public class FTP
        {
            public const string UserName = @"bdat100119f\bdat1001";
            public const string Password = "bdat1001";

            public const string BaseUrl = "ftp://waws-prod-dm1-127.ftp.azurewebsites.windows.net/bdat1001-10983";
            public const string MyUrl = BaseUrl + "/200471940 Karish Thangarajah";
            public const string localUrl = @"C:\Users\karis\OneDrive\Desktop\BDAT1001\midterm\";

            public const int OperationPauseTime = 10000;
        }

        public class Filename
        {
            public const string WordDoc = "/info.docx";
            public const string ExcelDoc = "/info.xlsx";
            public const string PowerPointDoc = "/info.pptx";
        }

        public class Student
        {
            public const string InfoCSVFileName = "info.csv";
            public const string MyImageFileName = "myimage.jpg";
        }
    }
}





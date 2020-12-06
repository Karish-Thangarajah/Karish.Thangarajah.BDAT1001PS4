using System;

namespace FileMaker.Models
{
    public class Student
    {
        public string StudentId { get; set; }
        public string FirstName { get; set; }
        public string LastName { get; set; }
        public string StudentCode { get; set; }
        /// <summary>
        /// ImageData stored as Base64
        /// </summary>
        public string ImageData { get; set; }
        public bool MyRecord { get; set; }
        public DateTime DateOfBirth { get; set; }
        public DateTime CreateDate { get; set; }
        public DateTime EditDate { get; set; }

        public string ToCSV()
        {
            return $"{StudentCode},{FirstName},{LastName},{MyRecord}";
        }
        public override string ToString()
        {
            return $"{StudentCode} - {LastName}, {FirstName}";
        }
    }

}

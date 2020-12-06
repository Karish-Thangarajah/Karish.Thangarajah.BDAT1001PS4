using System;

namespace FileMaker.Models.Api
{

    public class ApiResult<T> where T : class
    {
        public string Message { get; set; }
        public int StatusCode { get; set; }
        public bool Succeeded { get; set; }
        public DateTime DateAccessed { get; set; }
        public T Data { get; set; }
    }

}

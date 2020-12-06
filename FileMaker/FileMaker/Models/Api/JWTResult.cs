using System;

namespace FileMaker.Models.Api
{
    public class JwtResult
    {
        public string AccessToken { get; set; }
        public string ExpiresOn { get; set; }
    }
}

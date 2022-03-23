namespace LoginWindow
{
    public class User
    {
        public string role { get; set; }
        public string success { get; set; }
        public string token { get; set; }

        public Credentials userCredentials { get; set; }
    }

    public class Credentials
    {
        public string email { get; set; }
        public string first_name { get; set; }
        public string last_name { get; set; }
    }
}
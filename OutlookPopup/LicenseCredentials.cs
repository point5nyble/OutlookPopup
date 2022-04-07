namespace OutlookPopup
{
    internal class LicenseCredentials
    {
        public string email { get; set; }
        public string password { get; set; }
    }

    internal class LoginResponse
    {
        public bool success { get; set; }
        public string token { get; set; }
        public string role { get; set; }
        //public string MyProperty { get; set; }
        public LoginResponse() { }
    }
}
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.Json.Serialization;
using System.Threading.Tasks;

namespace OutlookPopup
{
    public class ClientInfo
    {
        [JsonPropertyName("email")]
        public string EmailId { get; set; }
        [JsonPropertyName("client_os_version")]
        public string MachineOS { get; set; }
        [JsonPropertyName("outlook_version")]
        public string OutlookVersion { get; set; }
        //public string MachineName { get; set; }
        [JsonPropertyName("activationId")]
        public string ActivationId { get; set; }
    }
}

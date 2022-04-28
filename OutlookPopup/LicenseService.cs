﻿using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.Json;
using System.Threading.Tasks;
using System.Windows;

namespace OutlookPopup
{
    public class LicenseService
    {

        public  static string ServerAddress
        {
            get
            {
                return Properties.Settings.Default.ServerAddress;
            }
        }
        private static readonly log4net.ILog log =
               log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);
        public static async Task<LicenseStatus> LogLicenseUsage(ClientInfo info, string token)
        {

            var json = JsonSerializer.Serialize(info);
            var data = new StringContent(json, Encoding.UTF8, "application/json");

            var url = ServerAddress + "api/checkLicense";
            using (var client = new HttpClient())
            {
                client.DefaultRequestHeaders.Add("x-access-token", token);
                var response = await client.PostAsync(url, data);
                string result = response.Content.ReadAsStringAsync().Result;
                if (result=="No valid user found")
                {

                    return new LicenseStatus { Message = "No License Assigned, Please contact your administrator.", IsValid = false };
                    
                }
                else if(result.Contains("License is expired"))
                {
                    return new LicenseStatus { Message = "License expired, Please contact your administrator.", IsValid = false };
                    
                }
            }
            return new LicenseStatus { Message = "Valid License", IsValid = true };
            
        }

        internal static Task<bool> HasOfflineLimitReachedAsync()
        {
            log.Info("Checking if Offline Limit has reached");
            int daysLeft = Properties.Settings.Default.OfflineLimit;
            DateTime lastUpdated = Properties.Settings.Default.LastUpdated;
            if (daysLeft == 0)
            {
                log.Info("Offline limit Reached.");
                return Task.FromResult(true);                 
            }
            else
            {
                if ((DateTime.Today - lastUpdated).TotalDays != 0)
                {
                    Properties.Settings.Default.OfflineLimit = daysLeft - 1;
                    Properties.Settings.Default.LastUpdated = DateTime.Today;
                    Properties.Settings.Default.Save();
                    log.Info($"Within Offline limit.{daysLeft} left in offline expiration.");
                    
                }
                return Task.FromResult(false);
            }          
        }

        internal static void clearOfflineLimitInfo()
        {
            string filePath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                "OfflineTracker.dat");
            if (File.Exists(filePath))
            {
                File.Delete(filePath);
            }
        }
        internal static async Task<bool> IsTokenValid(string email, string token)
        {
            log.Info("Checking if token is valid.");
            string json = JsonSerializer.Serialize(new ClientInfo
            {
                EmailId = email

            });
            var data = new StringContent(json, Encoding.UTF8, "application/json");
                
            string url = ServerAddress + "api/checkLicense";
            System.Net.ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls | SecurityProtocolType.Tls11 | SecurityProtocolType.Tls12;
            using (var client = new HttpClient())
            {
                client.DefaultRequestHeaders.Add("x-access-token", token);
                var response = await client.PostAsync(url, data);
                
                string result = response.Content.ReadAsStringAsync().Result;
                if (result == "Invalid Token")
                {
                    log.Info($"Invalid token .Response:{response}");
                    return false;
                }
                else
                {
                    log.Info($"Token is valid.");
                    return true;
                }                    
            }            
        }
        internal static async Task<LoginResponse> Login(string email,string password)
        {

            string url = ServerAddress + $"api/auth?email={email}&password={password}";
            ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls | SecurityProtocolType.Tls11 | SecurityProtocolType.Tls12;
            
            using (HttpClient client = new HttpClient())
            {
                using (var response = await client.PostAsync(url, null))
                {
                    if (response.IsSuccessStatusCode)
                    {
                        var result = response.Content.ReadAsStreamAsync().Result;

                        LoginResponse loginResponse = await JsonSerializer.DeserializeAsync<LoginResponse>(result);
                        return loginResponse;
                    }
                    else
                        return new LoginResponse { success = false, token = string.Empty };
                }
            }

        }

        internal static async Task<bool> IsValidLicenseID(string loggedinUserEmailId, string id, string token )
        {
            log.Info("Checking for valid Organization License");
            string softwareId = Properties.Settings.Default.SoftwareID;

            
            string url = $"{ServerAddress}api/software/{softwareId}/licenses/{id}/activationbyemail/{loggedinUserEmailId}";
            System.Net.ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls | SecurityProtocolType.Tls11 | SecurityProtocolType.Tls12;
            using (var client = new HttpClient())
            {
                client.DefaultRequestHeaders.Add("x-access-token", token);
                var response = await client.GetAsync(url);

                if (response.IsSuccessStatusCode)
                {

                    log.Info($"Organization License assigned, User has a valid license for {loggedinUserEmailId}.");
                    return true;

                }
                else
                {
                    string result = response.Content.ReadAsStringAsync().Result;
                    if (response.StatusCode == HttpStatusCode.Unauthorized)
                    {
                        log.Info($"User Not authorized, Server Response:{result}");
                        return false;
                    }
                    else if (response.StatusCode == HttpStatusCode.NotFound)
                    {
                        log.Info($"Organization/User has no valid license.Server Response:{result}");
                        //activate a new license
                        var activationResponse = await AssignLicense(id, loggedinUserEmailId, token);
                        if (activationResponse)
                        {
                            return true;
                        }
                        else
                        {
                            return false;
                        }
                    }
                    else
                        return false;
                }
               
            }
            
        }

        private static async Task<bool> AssignLicense(string id, string loggedinUserEmailId, string token)
        {
            log.Info($"Activating license for the user {loggedinUserEmailId}");
            string softwareId = Properties.Settings.Default.SoftwareID;
            Guid guid = Guid.NewGuid();
            string url = $"{ServerAddress}api/software/{softwareId}/licenses/{id}/activations";
            var json = JsonSerializer.Serialize(new ClientInfo
            {
                EmailId = loggedinUserEmailId,
                ActivationId = guid.ToString()
            });
            var data = new StringContent(json, Encoding.UTF8, "application/json");

            System.Net.ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls | SecurityProtocolType.Tls11 | SecurityProtocolType.Tls12;
            using (var client = new HttpClient())
            {
                client.DefaultRequestHeaders.Add("x-access-token", token);
                var response = await client.PostAsync(url, data);

                if (response.IsSuccessStatusCode)
                {
                    log.Info($"License activation for the user {loggedinUserEmailId} successful.");
                    //Log License Usage
                    LogUsage(loggedinUserEmailId, guid.ToString(), token);
                    return true;
                }
                else
                {
                    string result = response.Content.ReadAsStringAsync().Result;
                    log.Info($"License activation for the user {loggedinUserEmailId} failed. Server Response {result}");
                    return false;
                }
                    
            }
        }

        private static async void LogUsage(string userEmail,string activationId,string token)
        {
            //Log the license Usage.
            ClientInfo client = new ClientInfo();
            client.EmailId = userEmail;
            client.MachineOS = GetMachineOS();
            client.ActivationId = activationId;
            client.OutlookVersion = GetOutlookVersion();
            await LicenseService.LogLicenseUsage(client, token);
        }

        private static string GetOutlookVersion()
        {
            return Globals.ThisAddIn.Application.Version;
        }

        private static string GetMachineOS()
        {
            OperatingSystem os = Environment.OSVersion;

            if (os.Version.Major > 6)
            {
                return "Win 10";
            }
            else if (os.Version.Minor == 2)
            {
                return "Win 8/8.1";
            }
            else if (os.Version.Minor == 1)
            {
                return "Win 7";
            }
            else
                return "Lower than Win 7";

        }

        [DllImport("wininet.dll")]
        private extern static bool InternetGetConnectedState(out int Description, int ReservedValue);
    }
}

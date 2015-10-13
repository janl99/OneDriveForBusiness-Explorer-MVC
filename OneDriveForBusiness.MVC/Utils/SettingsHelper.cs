using System;
using System.Configuration;

namespace OneDriveForBusiness.MVC.Utils
{
    public class SettingsHelper
    {
        private static string _clientId = ConfigurationManager.AppSettings["ida:ClientId"];
        private static string _clientSecret = ConfigurationManager.AppSettings["ida:ClientSecret"];
        //private static string _appKey = ConfigurationManager.AppSettings["ida:AppKey"] ?? ConfigurationManager.AppSettings["ida:Password"];

        private static string _authorizationUri = "https://login.windows.net";
        private static string _graphResourceId = "https://graph.windows.net";
        private static string _authority = "https://login.windows.net/common/";
        private static string _discoverySvcResourceId = "https://api.office.com/discovery/";
        private static string _discoverySvcEndpointUri = "https://api.office.com/discovery/v1.0/me/";

        private static string _capability = "MyFiles";

        public static string Capability
        {
            get
            {
                return _capability;
            }
        }

        public static string ClientId
        {
            get
            {
                return _clientId;
            }
        }

        public static string ClientSecret
        {
            get
            {
                return _clientSecret;
            }
        }

        //public static string AppKey
        //{
        //    get
        //    {
        //        return _appKey;
        //    }
        //}

        public static string AuthorizationUri
        {
            get
            {
                return _authorizationUri;
            }
        }

        public static string Authority
        {
            get
            {
                return _authority;
            }
        }

        public static string AADGraphResourceId
        {
            get
            {
                return _graphResourceId;
            }
        }

        public static string DiscoveryServiceResourceId
        {
            get
            {
                return _discoverySvcResourceId;
            }
        }

        public static Uri DiscoveryServiceEndpointUri
        {
            get
            {
                return new Uri(_discoverySvcEndpointUri);
            }
        }
    }
}

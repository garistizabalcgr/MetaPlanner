﻿using Microsoft.Extensions.Configuration;
using System;
using System.Globalization;
using System.IO;
using Windows.ApplicationModel;

namespace MetaPlanner
{
    /// <summary>
    /// Description of the configuration of an AzureAD public client application (desktop/mobile application). This should
    /// match the application registration done in the Azure portal
    /// </summary>
    public class Configuration
    {
        /// <summary>
        /// instance of Azure AD, for example public Azure or a Sovereign cloud (Azure China, Germany, US government, etc ...)
        /// </summary>
        public string Instance { get; set; } = "https://login.microsoftonline.com/{0}";

        /// <summary>
        /// Graph API endpoint, could be public Azure (default) or a Sovereign cloud (US government, etc ...)
        /// </summary>
        public string ApiUrl { get; set; } = "https://graph.microsoft.com/";

        public string MSGraphURL { get; set; } = "https://graph.microsoft.com/v1.0/";

        /// <summary>
        /// Scope for API call comma separated
        /// </summary>
        public string Scopes { get; set; }

        /// <summary>
        /// Scope for API call comma separated
        /// </summary>
        public int ChunkSize { get; set; }

        /// <summary>
        /// Array of scopes (splited)
        /// </summary>
        public string[] ScopesArray
        {
            get
            {
                var arr = Scopes.Split(",");
                for( int i =0; i < arr.Length; i++)
                {
                    arr[i] = arr[i].Trim();
                }
                return arr;
            }
        }

        /// <summary>
        /// The Tenant is:
        /// - either the tenant ID of the Azure AD tenant in which this application is registered (a guid)
        /// or a domain name associated with the tenant
        /// - or 'organizations' (for a multi-tenant application)
        /// </summary>
        public string Tenant { get; set; }

        /// <summary>
        /// Guid used by the application to uniquely identify itself to Azure AD
        /// </summary>
        public string ClientId { get; set; }

        /// <summary>
        /// URL of the authority
        /// </summary>
        public string Authority
        {
            get
            {
                return String.Format(CultureInfo.InvariantCulture, Instance, Tenant);
            }
        }

        /// <summary>
        /// Client secret (application password)
        /// </summary>
        /// <remarks>Daemon applications can authenticate with AAD through two mechanisms: ClientSecret
        /// (which is a kind of application password: this property)
        /// or a certificate previously shared with AzureAD during the application registration 
        /// (and identified by the CertificateName property belows)
        /// <remarks> 
        public string ClientSecret { get; set; }

        /// <summary>
        /// Name of a certificate in the user certificate store
        /// </summary>
        /// <remarks>Daemon applications can authenticate with AAD through two mechanisms: ClientSecret
        /// (which is a kind of application password: the property above)
        /// or a certificate previously shared with AzureAD during the application registration 
        /// (and identified by this CertificateName property)
        /// <remarks> 
        public string CertificateName { get; set; }

        public string Site { get; set; }

        public string AltSite { get; set; }

        public string Drive { get; set; }

        public string FolderName { get; set; }

        public string SubFolderName { get; set; }

        public bool IsSharePointListEnabled { get; set; }

        /// <summary>
        /// Reads the configuration from a json file
        /// </summary>
        /// <param name="path">Path to the configuration json file</param>
        /// <returns>AuthenticationConfig read from the json file</returns>
        public static Configuration ReadFromJsonFile()
        {
            IConfigurationRoot Configuration;
            string path = Package.Current.InstalledLocation.Path;
            var builder = new ConfigurationBuilder()
                .SetBasePath(path)
                .AddJsonFile("appsettings.json")
                .AddJsonFile("appsettings.production.json", optional: true);
            Configuration = builder.Build();
            return Configuration.Get<Configuration>();
        }

    }
}

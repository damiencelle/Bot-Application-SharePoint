using AuthBot.Models;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Web;
using System.Web.Http;
using System.Web.Routing;

namespace Bot_Application2
{
    public class WebApiApplication : System.Web.HttpApplication
    {
        protected void Application_Start()
        {
            GlobalConfiguration.Configure(WebApiConfig.Register);
            AuthSettings.Mode = ConfigurationManager.AppSettings["ActiveDirectory.Mode"];
            AuthSettings.EndpointUrl = ConfigurationManager.AppSettings["ActiveDirectory.EndpointUrl"];
            AuthSettings.Tenant = ConfigurationManager.AppSettings["ActiveDirectory.Tenant"];
            AuthSettings.RedirectUrl = ConfigurationManager.AppSettings["ActiveDirectory.RedirectUrl"];
            AuthSettings.ClientId = ConfigurationManager.AppSettings["ActiveDirectory.ClientId"];
            AuthSettings.ClientSecret = ConfigurationManager.AppSettings["ActiveDirectory.ClientSecret"];
        }
    }
}

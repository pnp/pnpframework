using Microsoft.SharePoint.Client;
using PnP.Framework;
using System.Security;

namespace ConsoleApp2
{
    internal class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Hello, World!");

            var username = "gautam@gautamdev.onmicrosoft.com";
            var pass = "wQ5FROvgTNp121mVUNhrdrPOuyUa9OiO";

            var securePassword = new SecureString();
            
            foreach (char c in pass)
                securePassword.AppendChar(c);

            using (var authenticationMgr = new AuthenticationManager(username, securePassword))
            {
                var ctx = authenticationMgr.GetContext("https://gautamdev.sharepoint.com/sites/testtz1");

                ctx.Web.EnsureProperty(w => w.Title);

                var value = ctx.Web;
            }            
        }
    }
}
using System;
using System.Net;
using System.Security.Cryptography.X509Certificates;

namespace Apress.VisualCSharpRecipes.Chapter10
{
    class Recipe10_06
    {
        public static void Main() 
        {
            // Create a WebRequest that authenticates the user with a  
            // username and password combination over Basic authentication.
            WebRequest requestA = WebRequest.Create("http://www.somesite.com");
            requestA.Credentials = new NetworkCredential("userName", "password");
            requestA.PreAuthenticate = true;

            // Create a WebRequest that authenticates the current user  
            // with Windows integrated authentication.
            WebRequest requestB = WebRequest.Create("http://www.somesite.com");
            requestB.Credentials = CredentialCache.DefaultCredentials;
            requestB.PreAuthenticate = true;

            // Create a WebRequest that authenticates the user with a client 
            // certificate loaded from a file.
            HttpWebRequest requestC = 
                (HttpWebRequest)WebRequest.Create("http://www.somesite.com");
            X509Certificate cert1 =
                X509Certificate.CreateFromCertFile(@"..\..\TestCertificate.cer");
            requestC.ClientCertificates.Add(cert1);

            // Create a WebRequest that authenticates the user with a client 
            // certificate loaded from a certificate store. Try to find a
            // certificate with a specific subject, but if it is not found
            // present the user with a dialog so they can select the certificate 
            // to use from their personal store.
            HttpWebRequest requestD = 
                (HttpWebRequest)WebRequest.Create("http://www.somesite.com");
            X509Store store = new X509Store();
            X509Certificate2Collection certs =
                store.Certificates.Find(X509FindType.FindBySubjectName, 
                "Allen Jones", false);
           
            if (certs.Count == 1)
            {
                requestD.ClientCertificates.Add(certs[0]);
            }
            else 
            {
                certs = X509Certificate2UI.SelectFromCollection(
                    store.Certificates,
                    "Select Certificate",
                    "Select the certificate to use for authentication.",
                    X509SelectionFlag.SingleSelection);

                if (certs.Count != 0) 
                {
                    requestD.ClientCertificates.Add(certs[0]);
                }
            }
            
            // Now issue the request and process the responses...
        }
    }
}
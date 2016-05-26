﻿using System;
using System.Globalization;
using System.IO;
using System.Security;
using System.Xml;
using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.UserProfiles;
using OfficeDevPnP.Core.Framework.Provisioning.Connectors;
using OfficeDevPnP.Core.Framework.Provisioning.Model;
using OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers;
using OfficeDevPnP.Core.Framework.Provisioning.Providers.Xml;
using Provisioning.CLI.Console.ClParser;

namespace Provisioning.CLI.Console
{
    /// <summary>
    /// This sample demonstrates how to extract and apply templates using the CSOM model.
    /// </summary>
    class Program
    {

        /// <summary>
        /// Member to store the secure password between uploads
        /// </summary>
        private static SecureString securePwd = null;

        /// <summary>
        /// Main methof
        /// </summary>
        /// <param name="args">The command line arguments</param>
        /// <example>Extract template: -action Extracttemplate -Url https://brk365tests.sharepoint.com/sites/nav -LoginMethod SPO -User=brk365tests@brk365tests.onmicrosoft.com -OUTFILE "C:\Users\brk\Desktop\template.xml" -password Enjoy123.</example>
        /// <example>Apply template: -action Applytemplate -Url https://brk365tests.sharepoint.com/sites/nav -LoginMethod SPO -User=brk365tests@brk365tests.onmicrosoft.com -INFILE "C:\Users\brk\Desktop\template.xml" -password Enjoy123.</example>
        /// <example>Apply mutiple templates with absolute paths in file: -action Applytemplate -LoginMethod SPO -User=brk365tests@brk365tests.onmicrosoft.com -INFILE "C:\Users\brk\Desktop\sitesAbsPath.xml" -password Enjoy123.</example>
        /// <example>Apply mutiple templates with relative paths in file: -action Applytemplate -LoginMethod SPO -User=brk365tests@brk365tests.onmicrosoft.com -INFILE "C:\Users\brk\Desktop\sitesRelPath.xml" -Url https://brk365tests.sharepoint.com/sites/nav -password Enjoy123.</example>
        /// <returns>0 in case of success or the error number</returns>
        static int Main(string[] args)
        {

            try
            {
                //Parse command line
                Parser parser = new Parser(args);
                if (!parser.ClIsOk)
                {
                    parser.Usage();
#if DEBUG
                    System.Console.Read();
#endif
                    return 1;
                }

                //Check action to be done
                switch ((Actions)parser.ClParameters[Params.Action])
                {
                    case Actions.Extracttemplate:

                        //We have to extract a template
                        System.Console.Out.WriteLine("Extracting template");
                        System.Console.Out.WriteLine("-------------------");

                        Uri fromuri = new Uri((string)parser.ClParameters[Params.Url]);
                        System.Console.Out.WriteLine("From url: " + fromuri.ToString());
                        using (ClientContext context = new ClientContext(fromuri))
                        {
                            //Login to web
                            LoginToWeb(parser, context);

                            //Setting user language to web install language
                            SetUserLanguageToWebLanguage(context);

                            //Exporting template
                            FileInfo outFile = new FileInfo((string)parser.ClParameters[Params.Outfile]);
                            System.Console.Out.WriteLine("To file: " + outFile.FullName);
                            if (outFile.Exists)
                            {
                                if (System.IO.File.Exists(outFile.FullName + ".bak"))
                                    System.IO.File.Delete(outFile.FullName + ".bak");
                                System.IO.File.Move(outFile.FullName, outFile.FullName + ".bak");
                            }

                            ProvisioningTemplateCreationInformation cri = new ProvisioningTemplateCreationInformation(context.Web);
                            cri.FileConnector = new FileSystemConnector(outFile.Directory.FullName, "");
                            cri.IncludeAllTermGroups = true;
                            cri.IncludeNativePublishingFiles = true;
                            cri.IncludeSearchConfiguration = true;
                            cri.IncludeSiteCollectionTermGroup = true;
                            cri.IncludeSiteGroups = true;
                            cri.PersistBrandingFiles = true;
                            cri.PersistPublishingFiles = true;
                            cri.PersistMultiLanguageResources = true;
                            cri.ProgressDelegate = delegate (String message, Int32 progress, Int32 total)
                            {
                                System.Console.WriteLine("  {0:00}/{1:00} - {2}", progress, total, message);
                            };
                            ProvisioningTemplate template = context.Web.GetProvisioningTemplate(cri);

                            XMLTemplateProvider provider =
                                        new XMLFileSystemTemplateProvider(outFile.Directory.FullName, "");
                            provider.SaveAs(template, outFile.Name);
                        }

                        System.Console.Out.WriteLine("Done");
                        break;

                    case Actions.Applytemplate:

                        //We have to apply a template
                        System.Console.Out.WriteLine("Applying template");
                        System.Console.Out.WriteLine("-----------------");

                        //Check inputfile
                        FileInfo inFile = new FileInfo((string)parser.ClParameters[Params.Infile]);
                        if (!inFile.Exists)
                        {
                            System.Console.Error.WriteLine("Can't find the specified input file: " + inFile.FullName);
                            return 2;
                        }

                        //Check input file type
                        XmlTextReader reader = new XmlTextReader(inFile.FullName);
                        string mode = null;
                        while (reader.Read())
                        {
                            switch (reader.NodeType)
                            {
                                case XmlNodeType.Element:
                                    mode = reader.LocalName;
                                    break;
                            }
                            if (mode != null) break;
                        }
                        reader.Close();
                        if (mode == "Provisioning")
                        {
                            //A provisioning template has been given
                            System.Console.Out.WriteLine("From file: " + inFile.FullName);
                            if (!parser.ClParameters.ContainsKey(Params.Url))
                            {
                                System.Console.Error.WriteLine("Parameter url is required if you like to apply a template xml!");
                                return 3;
                            }
                            string tourl = (string)parser.ClParameters[Params.Url];
                            ApplyTemplate(parser, inFile, tourl);
                        }
                        else
                        {
                            //We have to do multiple sites
                            XmlDocument doc = new XmlDocument();
                            doc.Load(inFile.FullName);
                            foreach (XmlNode site in doc.DocumentElement.ChildNodes)
                            {
                                string fileName = site.Attributes["file"].Value;
                                string siteUrl = site.Attributes["url"].Value;

                                //Check inputfile
                                FileInfo inFileSite = new FileInfo(fileName);
                                if (!inFileSite.Exists)
                                    inFileSite = new FileInfo(inFile.Directory.FullName + Path.DirectorySeparatorChar + fileName);
                                System.Console.Out.WriteLine("From file: " + inFileSite.FullName);
                                if (!inFileSite.Exists)
                                {
                                    System.Console.Error.WriteLine("Can't find the specified input file: " + inFileSite.ToString());
                                    return 4;
                                }

                                //Check site url
                                if (!siteUrl.ToLower().StartsWith("http"))
                                {
                                    if (!parser.ClParameters.ContainsKey(Params.Url))
                                    {
                                        System.Console.Error.WriteLine("The configured site has a relative url. Please specify the url parameter to build an absolute one!");
                                        return 5;
                                    }
                                    siteUrl = (string)parser.ClParameters[Params.Url] + siteUrl;
                                }

                                ApplyTemplate(parser, inFileSite, siteUrl);
                                System.Console.Out.WriteLine("");
                            }
                        }

                        System.Console.Out.WriteLine("Done");
                        break;
                }

                return 0;

            }
            catch (Exception ex)
            {
                System.Console.Error.WriteLine("A unhandeled exception occured:");
                System.Console.Error.WriteLine(ex.ToString());
#if DEBUG
                System.Console.Read();
#endif
            }
            return 999;
        }

        /// <summary>
        /// Method to apply a template to a web
        /// </summary>
        /// <param name="parser">The command line parser</param>
        /// <param name="inFile">The template file to be uploaded</param>
        /// <param name="tourl">The web where the template has to be uploaded</param>
        private static void ApplyTemplate(Parser parser, FileInfo inFile, string tourl)
        {
            Uri touri = new Uri(tourl);
            System.Console.Out.WriteLine("To url: " + touri.ToString());
            using (ClientContext context = new ClientContext(touri))
            {

                //Login to web
                LoginToWeb(parser, context);

                //Prepare template
                ProvisioningTemplateApplyingInformation ptai = new ProvisioningTemplateApplyingInformation();
                ptai.ProgressDelegate = delegate (String message, Int32 progress, Int32 total)
                {
                    System.Console.WriteLine("  {0:00}/{1:00} - {2}", progress, total, message);
                };

                XMLTemplateProvider provider = new XMLFileSystemTemplateProvider(inFile.Directory.FullName, "");
                ProvisioningTemplate template = provider.GetTemplate(inFile.Name);

                //Setting user language to web install language
                SetUserLanguageToTemplateLanguage(context, template);

                //Applying template
                FileSystemConnector connector = new FileSystemConnector(inFile.Directory.FullName, "");
                template.Connector = connector;

                context.Web.ApplyProvisioningTemplate(template, ptai);
            }
        }

        /// <summary>
        /// Logs in to the web
        /// </summary>
        /// <param name="parser">The command line parser</param>
        /// <param name="context">The actual CSOM context</param>
        private static void LoginToWeb(Parser parser, ClientContext context)
        {
            switch ((LoginMethod)parser.ClParameters[Params.Loginmethod])
            {
                case LoginMethod.Spo:

                    if (securePwd == null) securePwd = GetSecurePassword(parser);
                    context.Credentials =
                        new SharePointOnlineCredentials(
                            (string)parser.ClParameters[Params.User],
                            securePwd);
                    break;
            }
        }

        /// <summary>
        /// Configures the user profile corresponding to the template
        /// </summary>
        /// <param name="context">The actual CSOM context</param>
        /// <param name="template">The template to be uploaded</param>
        private static void SetUserLanguageToTemplateLanguage(ClientContext context, ProvisioningTemplate template)
        {
            Microsoft.SharePoint.Client.User user = context.Web.CurrentUser;
            context.Load(user);
            context.ExecuteQuery();
            PeopleManager peopleManager = new PeopleManager(context);
            int lcid = template.RegionalSettings.LocaleId;
            CultureInfo siteCulture = CultureInfo.GetCultureInfo(lcid);
            var muiLanguages = siteCulture.Name;
            System.Console.Out.WriteLine("Using language: " + muiLanguages);
            var customRegionalSettings = "False";
            var locale = "" + lcid;
            var timeZoneID = "" + template.RegionalSettings.TimeZone;
            peopleManager.SetSingleValueProfileProperty(user.LoginName, "SPS-MUILanguages", muiLanguages);
            peopleManager.SetSingleValueProfileProperty(user.LoginName, "SPS-RegionalSettings-FollowWeb", customRegionalSettings);
            peopleManager.SetSingleValueProfileProperty(user.LoginName, "SPS-Locale", locale);
            peopleManager.SetSingleValueProfileProperty(user.LoginName, "SPS-TimeZone", timeZoneID);
            context.ExecuteQuery();
        }

        /// <summary>
        /// Configures the user profile corresponding to the web the template has to be extracted from
        /// </summary>
        /// <param name="context">The actual CSOM context</param>
        private static void SetUserLanguageToWebLanguage(ClientContext context)
        {
            Microsoft.SharePoint.Client.User user = context.Web.CurrentUser;
            context.Load(user);
            context.Web.Context.Load(context.Web.RegionalSettings);
            context.Web.Context.Load(context.Web.RegionalSettings.TimeZone, tz => tz.Id);
            context.ExecuteQuery();
            PeopleManager peopleManager = new PeopleManager(context);
            int lcid = (int)context.Web.RegionalSettings.LocaleId;
            CultureInfo siteCulture = CultureInfo.GetCultureInfo(lcid);
            var muiLanguages = siteCulture.Name;
            System.Console.Out.WriteLine("Using language: " + muiLanguages);
            var customRegionalSettings = "False";
            var locale = "" + lcid;
            var timeZoneID = "" + context.Web.RegionalSettings.TimeZone.Id;
            peopleManager.SetSingleValueProfileProperty(user.LoginName, "SPS-MUILanguages", muiLanguages);
            peopleManager.SetSingleValueProfileProperty(user.LoginName, "SPS-RegionalSettings-FollowWeb", customRegionalSettings);
            peopleManager.SetSingleValueProfileProperty(user.LoginName, "SPS-Locale", locale);
            peopleManager.SetSingleValueProfileProperty(user.LoginName, "SPS-TimeZone", timeZoneID);
            context.ExecuteQuery();
        }

        /// <summary>
        /// Gets the password
        /// </summary>
        /// <param name="parser">The command line parser</param>
        /// <returns>The password as secure string or null if a password can't be defined</returns>
        private static SecureString GetSecurePassword(Parser parser)
        {
            SecureString secPwd = new SecureString();
            if (!parser.ClParameters.ContainsKey(Params.Password))
            {
                if (!parser.ClParameters.ContainsKey(Params.Passwordfile))
                {
                    //Getting password from console
                    if (parser.ClOptions.Contains(Options.NoInteraction))
                        return null;
                    System.Console.Out.Write("Please enter password: ");
                    while (true)
                    {
                        ConsoleKeyInfo i = System.Console.ReadKey(true);
                        if (i.Key == ConsoleKey.Enter)
                        {
                            break;
                        }
                        else if (i.Key == ConsoleKey.Backspace)
                        {
                            if (secPwd.Length > 0)
                            {
                                secPwd.RemoveAt(secPwd.Length - 1);
                                System.Console.Write("\b \b");
                            }
                        }
                        else
                        {
                            secPwd.AppendChar(i.KeyChar);
                            System.Console.Write("*");
                        }
                    }
                }
                else
                {
                    //Getting password from password file
                    FileInfo inFile = new FileInfo((string)parser.ClParameters[Params.Passwordfile]);
                    if (!inFile.Exists)
                    {
                        System.Console.Error.WriteLine("The given passwordFile does not exist: " + inFile.FullName);
                        parser.ClParameters.Remove(Params.Passwordfile);
                        return GetSecurePassword(parser);
                    }
                    using (StreamReader rdr = inFile.OpenText())
                    {
                        string pwd = rdr.ReadLine();
                        rdr.Close();
                        foreach (char chr in pwd)
                            secPwd.AppendChar(chr);
                    }
                }
            }
            else
            {
                //Getting password from command line
                foreach (char chr in ((string)parser.ClParameters[Params.Password]).ToCharArray())
                    secPwd.AppendChar(chr);
            }

            return securePwd;
        }

    }
}

using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.UserProfiles;
using OfficeDevPnP.Core.Framework.Provisioning.Connectors;
using OfficeDevPnP.Core.Framework.Provisioning.Model;
using OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers;
using OfficeDevPnP.Core.Framework.Provisioning.Providers.Xml;
using Provisioning.CLI.Console.ClParser;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Security;
using System.Text;
using System.Threading.Tasks;

namespace Provisioning.CLI.Console
{
    class Program
    {
        static void Main(string[] args)
        {

            //Parse command line
            Parser parser = new Parser(args);
            if (!parser.ClIsOk)
            {
                parser.Usage();
#if DEBUG
                System.Console.Read();
#endif
                return;
            }

            //Check action to be done
            switch ((Actions)parser.ClParameters[Params.Action])
            {
                case Actions.ExtractTemplate:

                    Uri uri = new Uri((string)parser.ClParameters[Params.Url]);
                    using (ClientContext context = new ClientContext(uri))
                    {

                        //Login to web
                        switch ((LoginMethod)parser.ClParameters[Params.Loginmethod])
                        {
                            case LoginMethod.SPO:

                                SecureString securePwd = GetSecurePassword(parser);
                                context.Credentials =
                                    new SharePointOnlineCredentials(
                                        (string)parser.ClParameters[Params.User],
                                        securePwd);
                                break;
                        }

                        //Setting user language to web install language
                        Microsoft.SharePoint.Client.User user = context.Web.CurrentUser;
                        context.Load(user);
                        context.Web.Context.Load(context.Web.RegionalSettings);
                        context.Web.Context.Load(context.Web.RegionalSettings.TimeZone, tz => tz.Id);
                        context.ExecuteQuery();
                        PeopleManager peopleManager = new PeopleManager(context);
                        int lcid = (int)context.Web.RegionalSettings.LocaleId;
                        CultureInfo siteCulture = CultureInfo.GetCultureInfo(lcid);
                        var muiLanguages = siteCulture.Name;
                        var customRegionalSettings = "False";
                        var locale = "" + lcid;
                        var timeZoneID = "" + context.Web.RegionalSettings.TimeZone.Id;
                        peopleManager.SetSingleValueProfileProperty(user.LoginName, "SPS-MUILanguages", muiLanguages);
                        peopleManager.SetSingleValueProfileProperty(user.LoginName, "SPS-RegionalSettings-FollowWeb", customRegionalSettings);
                        peopleManager.SetSingleValueProfileProperty(user.LoginName, "SPS-Locale", locale);
                        peopleManager.SetSingleValueProfileProperty(user.LoginName, "SPS-TimeZone", timeZoneID);
                        context.ExecuteQuery();

                        //Exporting template
                        FileInfo outFile = new FileInfo((string)parser.ClParameters[Params.Outfile]);
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
                            System.Console.WriteLine("{0:00}/{1:00} - {2}", progress, total, message);
                        };
                        ProvisioningTemplate template = context.Web.GetProvisioningTemplate(cri);

                        XMLTemplateProvider provider =
                                    new XMLFileSystemTemplateProvider(outFile.Directory.FullName, "");
                        provider.SaveAs(template, outFile.Name);
                    }

                    break;
            }


        }

        private static SecureString GetSecurePassword(Parser parser)
        {
            SecureString securePwd = new SecureString();
            if (!parser.ClParameters.ContainsKey(Params.Password))
            {
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
                        if (securePwd.Length > 0)
                        {
                            securePwd.RemoveAt(securePwd.Length - 1);
                            System.Console.Write("\b \b");
                        }
                    }
                    else
                    {
                        securePwd.AppendChar(i.KeyChar);
                        System.Console.Write("*");
                    }
                }
            }
            else
            {
                foreach (char chr in ((string)parser.ClParameters[Params.Password]).ToCharArray())
                    securePwd.AppendChar(chr);
            }

            return securePwd;
        }
    }
}

using System;
using System.Globalization;
using System.IO;
using System.Security;
using System.Xml;
using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.UserProfiles;
using Microsoft.Online.SharePoint.TenantAdministration;
using OfficeDevPnP.Core.Framework.Provisioning.Connectors;
using OfficeDevPnP.Core.Framework.Provisioning.Model;
using OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers;
using OfficeDevPnP.Core.Framework.Provisioning.Providers.Xml;
using Provisioning.CLI.Console.ClParser;
using System.Threading;
using System.Collections.Generic;
using System.Net;
using Microsoft.Online.SharePoint.TenantManagement;
using System.Xml.Schema;
using System.Text;

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
        /// <example>Extract template: -action Extracttemplate -Url https://konidev.sharepoint.com/sites/nav -LoginMethod SPO -User=admin@konidev.onmicrosoft.com -OUTFILE "C:\Data\@Trivadis\Lösungen\Kacheln\pnp\ExtractSearch\template.xml" -password Enjoy123.</example>
        /// <example>Extract template for an entire structure: -action Extracttemplate -EntireStructure -Url https://konidev.sharepoint.com/sites/nav -LoginMethod SPO -User=admin@konidev.onmicrosoft.com -OUTFILE "C:\Users\brk\Desktop\Structure" -password Enjoy123.</example>
        /// C:\Data\@Trivadis\Lösungen\Kacheln\pnp\ExtractSearch
        /// 
        /// <example>Apply template: -action Applytemplate -Url https://konidev.sharepoint.com/sites/nav1 -LoginMethod SPO -User=admin@konidev.onmicrosoft.com -INFILE "C:\Data\@Trivadis\Lösungen\Kacheln\srcSP\template.xml" -password Enjoy123.</example>
        /// <example>Apply mutiple templates with absolute paths in file: -action Applytemplate -LoginMethod SPO -User=admin@konidev.onmicrosoft.com -INFILE "C:\Users\brk\Desktop\sitesAbsPath.xml" -password Enjoy123.</example>
        /// <example>Apply mutiple templates with relative paths in file: -action Applytemplate -LoginMethod SPO -User=admin@konidev.onmicrosoft.com -INFILE "C:\Users\brk\Desktop\sitesRelPath.xml" -Url https://konidev.sharepoint.com/sites/nav -password Enjoy123.</example>
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

                        //Exporting template
                        FileInfo outFile = null;
                        string outFilePath = (string)parser.ClParameters[Params.Outfile];
                        if (System.IO.File.Exists(outFilePath))
                        {
                            if (System.IO.File.Exists(outFilePath + ".bak"))
                                System.IO.File.Delete(outFilePath + ".bak");
                            System.IO.File.Move(outFilePath, outFilePath + ".bak");
                            outFile = new FileInfo(outFilePath);
                        }
                        else if (System.IO.Directory.Exists(outFilePath))
                        {
                            if (parser.ClOptions.Contains(Options.Entirestructure))
                                outFilePath += Path.DirectorySeparatorChar + "site.xml";
                            else
                                outFilePath += Path.DirectorySeparatorChar + "template.xml";
                            outFile = new FileInfo(outFilePath);
                        }
                        else
                        {
                            outFile = new FileInfo(outFilePath);
                        }
                        System.Console.Out.WriteLine("To file: " + outFile.FullName);

                        Uri fromuri = new Uri(((string)parser.ClParameters[Params.Url]).TrimEnd("/".ToCharArray()));
                        if (parser.ClOptions.Contains(Options.Entirestructure))
                        {
                            using (StreamWriter txtw = new StreamWriter(outFile.OpenWrite(), Encoding.UTF8))
                            {
                                txtw.WriteLine("<?xml version=\"1.0\" encoding=\"UTF-8\"?>");
                                txtw.WriteLine("<sites>");
                                ExtractTemplateStructure(outFile, txtw, fromuri, parser);
                                txtw.WriteLine("</sites>");
                                txtw.Close();
                            }
                        }
                        else
                        {
                            System.Console.Out.WriteLine("From url: " + fromuri.ToString());
                            using (ClientContext context = new ClientContext(fromuri))
                            {
                                LoginToWeb(parser, context);
                                ExtractTemplate(context, outFile);
                            }
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

                        XmlDocument doc = new XmlDocument();
                        doc.Load(inFile.FullName);
                        mode = doc.DocumentElement.LocalName;

                        if (mode == "Provisioning")
                        {
                            //prepare namespace
                            XmlNamespaceManager xmlnsManager = new XmlNamespaceManager(doc.NameTable);
                            xmlnsManager.AddNamespace("pnp", "http://schemas.dev.office.com/PnP/2015/12/ProvisioningSchema");


                            //A provisioning template has been given
                            System.Console.Out.WriteLine("From file: " + inFile.FullName);
                            if (!parser.ClParameters.ContainsKey(Params.Url))
                            {
                                System.Console.Error.WriteLine("Parameter url is required if you like to apply a template xml!");
                                return 3;
                            }

                            //Check if it is a sequence
                            string tourl = (string)parser.ClParameters[Params.Url];
                            XmlNode seq = doc.DocumentElement.SelectSingleNode("//pnp:Sequence", xmlnsManager);
                            if (seq != null)
                            {
                                FileInfo tmplFile = new FileInfo(inFile.FullName);
                                ApplyTemplateSequence(parser, doc.DocumentElement, tmplFile.Directory, xmlnsManager, tourl);
                            }
                            else
                            {
                                ApplyTemplate(parser, inFile, tourl);
                            }
                        }
                        else
                        {
                            //We have to do multiple sites
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
        /// Extracts an entire site structure
        /// </summary>
        /// <param name="outFile">The sites.xml file</param>
        /// <param name="outFileWriter">The StreamWriter to write to sites.xml file</param>
        /// <param name="fromuri">The actual uri to be extracted</param>
        /// <param name="parser">The command line parser</param>
        /// <param name="parents">A string containing all parents</param>
        private static void ExtractTemplateStructure(FileInfo outFile, StreamWriter outFileWriter, Uri fromuri, Parser parser)
        {
            System.Console.Out.WriteLine("From url: " + fromuri.ToString());
            List<Uri> subUris = new List<Uri>();
            using (ClientContext context = new ClientContext(fromuri))
            {
                LoginToWeb(parser, context);
                Web oWebsite = context.Web;
                context.Load(oWebsite, website => website.Webs, website => website.ServerRelativeUrl);
                context.ExecuteQuery();

                string dirPath = System.Web.HttpUtility.UrlDecode(oWebsite.ServerRelativeUrl).Replace("/", "_");
                System.IO.Directory.CreateDirectory(outFile.Directory.FullName + Path.DirectorySeparatorChar + dirPath);
                FileInfo outPutFile = new FileInfo(outFile.Directory.FullName + 
                    Path.DirectorySeparatorChar + dirPath + Path.DirectorySeparatorChar + "template.xml");
                outFileWriter.WriteLine("  <site url=\"" + oWebsite.ServerRelativeUrl + "\" file=\""+ dirPath + Path.DirectorySeparatorChar + "template.xml\" />");
                outFileWriter.Flush();
                ExtractTemplate(context, outPutFile);
                foreach (Web subWeb in oWebsite.Webs)
                {
                    Uri subUri = new Uri(fromuri.AbsoluteUri.Substring(0, fromuri.AbsoluteUri.Length - fromuri.AbsolutePath.Length) + subWeb.ServerRelativeUrl);
                    subUris.Add(subUri);
                }
            }
            System.Console.Out.WriteLine("");

            foreach (Uri subUri in subUris)
            {
                ExtractTemplateStructure(outFile, outFileWriter, subUri, parser);
            }
        }

        private static void ExtractTemplate(ClientContext context, FileInfo outFile)
        {

            //Setting user language to web install language
            SetUserLanguageToWebLanguage(context);

            ProvisioningTemplateCreationInformation cri = new ProvisioningTemplateCreationInformation(context.Web);
            cri.FileConnector = new FileSystemConnector(outFile.Directory.FullName, "");

            if (!context.Web.IsSubSite())
            {
                cri.IncludeSiteCollectionTermGroup = true;
            }
            cri.IncludeNativePublishingFiles = true;
            cri.IncludeSearchConfiguration = true; //TODO is this valid per web?
            cri.IncludeSiteGroups = true;
            cri.IncludeAllTermGroups = true;
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

        /// <summary>
        /// Method to apply a template to a web
        /// </summary>
        /// <param name="parser">The command line parser</param>
        /// <param name="inFile">The template file to be uploaded</param>
        /// <param name="tourl">The web where the template has to be uploaded</param>
        private static void ApplyTemplateSequence(Parser parser, XmlElement rootNode, DirectoryInfo resourceDir, XmlNamespaceManager xmlnsManager, string tourl)
        {
            Uri touri = new Uri(tourl);
            System.Console.Out.WriteLine("To url: " + touri.ToString());
            using (ClientContext context = new ClientContext(touri))
            {

                //Login to web
                LoginToWeb(parser, context);

                //Reading sequence
                foreach (XmlNode seqNode in rootNode.SelectNodes("pnp:Sequence", xmlnsManager))
                {
                    //Preparing site collection
                    string seqID = seqNode.Attributes["ID"].Value;
                    System.Console.Out.WriteLine("Importing sequence: " + seqID);
                    XmlNode colNode = seqNode.SelectSingleNode("pnp:SiteCollection", xmlnsManager);
                    if (colNode == null)
                    {
                        System.Console.Error.WriteLine("No SiteCollection found in sequence!");
                        continue;
                    }

                    string colLanguage = colNode.Attributes["Language"].Value;
                    string colPrimarySiteCollectionAdmin = colNode.Attributes["PrimarySiteCollectionAdmin"].Value;
                    string colSecondarySiteCollectionAdmin = colNode.Attributes["SecondarySiteCollectionAdmin"].Value;
                    string colTimeZone = colNode.Attributes["TimeZone"].Value;
                    string colTitle = colNode.Attributes["Title"].Value;
                    string colUrl = colNode.Attributes["Url"].Value;
                    string colMembersCanShare = colNode.Attributes["MembersCanShare"].Value;
                    string colUserCodeMaximumLevel = colNode.Attributes["UserCodeMaximumLevel"].Value;
                    string colTemplate = colNode.Attributes["Template"].Value;

                    CreateSiteCollection(context, tourl, tourl + colUrl, colTitle, int.Parse(colLanguage), int.Parse(colTimeZone), 
                        colPrimarySiteCollectionAdmin, colTemplate, int.Parse(colUserCodeMaximumLevel));
                    UpdateSiteCollection(context, tourl, tourl + colUrl, colTitle, colSecondarySiteCollectionAdmin, colMembersCanShare);
                    System.Console.Out.WriteLine("Waiting one minute to let O365 finishing the site collection provisioning");
                    Thread.Sleep(60000);

                    //Applying template to site collection
                    XmlNode colTemplNode = colNode.SelectSingleNode("pnp:Templates/pnp:ProvisioningTemplateReference", xmlnsManager);
                    string templateID = colTemplNode.Attributes["ID"].Value;
                    XmlDocument templDoc = new XmlDocument();
                    CopyTemplate(rootNode, templDoc, templateID, xmlnsManager);
                    FileInfo inFile = new FileInfo((string)parser.ClParameters[Params.Infile]);
                    XmlWriterSettings settings = new XmlWriterSettings
                    {
                        Encoding = Encoding.UTF8,
                        ConformanceLevel = ConformanceLevel.Document,
                        OmitXmlDeclaration = false,
                        CloseOutput = true,
                        Indent = true,
                        IndentChars = "  ",
                        NewLineHandling = NewLineHandling.Replace
                    };
                    using (StreamWriter sw = System.IO.File.CreateText(inFile + ".tmp"))
                    {
                        using (XmlWriter writer = XmlWriter.Create(sw, settings))
                        {
                            templDoc.WriteContentTo(writer);
                            writer.Close();
                        }
                    }
                    System.Console.Out.WriteLine("Applying Site Collection template");
                    ApplyTemplate(parser, new FileInfo(inFile + ".tmp"), tourl + colUrl);
                    System.IO.File.Delete(inFile + ".tmp");

                    //Preparing webs
                    XmlNode locs = rootNode.SelectSingleNode("pnp:Localizations", xmlnsManager);
                    Dictionary<CultureInfo, XmlDocument> docs = new Dictionary<CultureInfo, XmlDocument>();
                    foreach (XmlNode loc in locs.SelectNodes("pnp:Localization", xmlnsManager))
                    {
                        string lcid = loc.Attributes["LCID"].Value;
                        string resourceFile = loc.Attributes["ResourceFile"].Value;
                        CultureInfo ci = new CultureInfo(int.Parse(lcid));
                        FileInfo resFile = resourceDir.GetFiles(resourceFile)[0];
                        XmlDocument doc = new XmlDocument();
                        doc.Load(resFile.FullName);
                        docs.Add(ci, doc);
                    }

                    using (ClientContext siteContext = new ClientContext(tourl + colUrl))
                    {
                        siteContext.Credentials = context.Credentials;

                        foreach (XmlNode siteNode in seqNode.SelectNodes("pnp:Site", xmlnsManager))
                        {
                            string siteLanguage = siteNode.Attributes["Language"].Value;
                            string siteTimeZone = siteNode.Attributes["TimeZone"].Value;
                            string siteTitle = siteNode.Attributes["Title"].Value;
                            string siteUrl = siteNode.Attributes["Url"].Value;
                            string siteQuickLaunchEnabled = siteNode.Attributes["QuickLaunchEnabled"].Value;
                            string siteTemplate = siteNode.Attributes["Template"].Value;

                            CreateSite(siteContext, tourl + colUrl + "/" + siteUrl, siteUrl, siteTitle, int.Parse(siteLanguage), int.Parse(siteTimeZone), siteTemplate);
                            UpdateSite(siteContext, colUrl + "/" + siteUrl, docs, siteTitle, siteQuickLaunchEnabled);

                            //Applying template to web
                            XmlNode siteTemplNode = siteNode.SelectSingleNode("pnp:Templates/pnp:ProvisioningTemplateReference", xmlnsManager);
                            if (siteTemplNode != null)
                            {
                                templateID = colTemplNode.Attributes["ID"].Value;
                                templDoc = new XmlDocument();
                                CopyTemplate(rootNode, templDoc, templateID, xmlnsManager);
                                templDoc.Save(inFile + ".tmp");
                                System.Console.Out.WriteLine("Applying Web template");
                                FileInfo tmpFile = new FileInfo(inFile + ".tmp");
                                ApplyTemplate(parser, tmpFile, tourl + colUrl + "/" + siteUrl);
                                tmpFile.Delete();
                            }
                        }

                    }
                }
            }
        }

        private static void CopyTemplate(XmlElement rootNode, XmlDocument templDoc, string templateID, XmlNamespaceManager xmlnsManager)
        {
            XmlNode Preferences = rootNode.SelectSingleNode("pnp:Preferences", xmlnsManager);
            XmlNode Localizations = rootNode.SelectSingleNode("pnp:Localizations", xmlnsManager);
            XmlNode ProvisioningTemplate = rootNode.SelectSingleNode("//pnp:ProvisioningTemplate[@ID='"+ templateID + "']", xmlnsManager);

            XmlNode newPreferences = templDoc.ImportNode(Preferences, true);
            XmlNode newLocalizations = templDoc.ImportNode(Localizations, true);
            XmlNode newProvisioningTemplate = templDoc.ImportNode(ProvisioningTemplate, true);

            XmlDeclaration xmlDeclaration = templDoc.CreateXmlDeclaration("1.0", "UTF-8", null);
            templDoc.AppendChild(xmlDeclaration);
            XmlSchema schema = new XmlSchema();
            schema.Namespaces.Add("pnp", "http://schemas.dev.office.com/PnP/2015/12/ProvisioningSchema");
            templDoc.Schemas.Add(schema);

            XmlElement rootElement = templDoc.CreateElement("pnp:Provisioning", "http://schemas.dev.office.com/PnP/2015/12/ProvisioningSchema");
            templDoc.AppendChild(rootElement);
            rootElement.AppendChild(newPreferences);
            rootElement.AppendChild(newLocalizations);
            XmlElement templatesElement = templDoc.CreateElement("pnp:Templates", "http://schemas.dev.office.com/PnP/2015/12/ProvisioningSchema");
            templatesElement.SetAttribute("ID", "CONTAINER-TEMPLATE");
            rootElement.AppendChild(templatesElement);
            templatesElement.AppendChild(newProvisioningTemplate);
        }

        private static void UpdateSite(ClientContext siteContext, string targetUrl, Dictionary<CultureInfo, XmlDocument> docs, string siteTitle, string siteQuickLaunchEnabled)
        {
            System.Console.Out.WriteLine("Updating Web " + targetUrl);
            Web theWeb = siteContext.Site.OpenWeb(targetUrl);
            siteContext.ExecuteQuery();
            string resName = siteTitle.Replace("{resource:", "").TrimEnd("}".ToCharArray());
            foreach (CultureInfo ci in docs.Keys)
            {
                XmlDocument doc = docs[ci];
                XmlNode data = doc.DocumentElement.SelectSingleNode("//data[@name='"+ resName + "']");
                if (data != null)
                    theWeb.TitleResource.SetValueForUICulture(ci.Name, data.InnerText);
            }
            if (!string.IsNullOrEmpty(siteQuickLaunchEnabled))
                theWeb.QuickLaunchEnabled = bool.Parse(siteQuickLaunchEnabled.ToLower());
            theWeb.Update();
            siteContext.ExecuteQuery();
        }

        private static void UpdateSiteCollection(ClientContext context, string rootUrl, string targetUrl, string colTitle, string colSecondarySiteCollectionAdmin, string colMembersCanShare)
        {
            System.Console.Out.WriteLine("Updating Site Collection " + targetUrl);
            using (var ctx = new ClientContext(rootUrl.Replace(".sharepoint.com", "-admin.sharepoint.com")))
            {
                ctx.Credentials = context.Credentials;
                var tenant = new Tenant(ctx);
                SPOSitePropertiesEnumerable spp = tenant.GetSiteProperties(0, true);
                ctx.Load(spp);
                ctx.ExecuteQuery();
                foreach (SiteProperties sp in spp)
                {
                    if (sp.Url == targetUrl)
                    {
                        sp.Title = colTitle;
                        if (!string.IsNullOrEmpty(colMembersCanShare) && bool.Parse(colMembersCanShare.ToLower()))
                            sp.SharingCapability = SharingCapabilities.ExternalUserAndGuestSharing;
                        if (!string.IsNullOrEmpty(colSecondarySiteCollectionAdmin))
                            tenant.SetSiteAdmin(targetUrl, colSecondarySiteCollectionAdmin, false);
                        break;
                    }
                }
                ctx.ExecuteQuery();
            }
        }

        /// <summary>
        /// Create a new site.
        /// </summary>
        /// <param name="targetUrl">rootsite + "/" + managedPath + "/" + sitename: e.g. "https://auto.contoso.com/sites/site1/sub1"</param>
        /// <param name="title">site title: e.g. "Test Site"</param>
        /// <param name="siteTemplate">The site template used to create this new site</param>
        private static void CreateSite(ClientContext siteContext, string fullUrl, string targetUrl, string title, 
            int language, int timeZoneId, String siteTemplate)
        {
            targetUrl = targetUrl.TrimStart("/".ToCharArray());
            if (!siteContext.WebExistsFullUrl(fullUrl))
            {
                WebCreationInformation creation = new WebCreationInformation();
                creation.Url = targetUrl;
                creation.Title = title;
                creation.WebTemplate = siteTemplate;
                creation.Language = language;
                siteContext.Web.Webs.Add(creation);
                System.Console.Out.WriteLine("Creating Web " + targetUrl);
                siteContext.ExecuteQuery();
            }
        }

        /// <summary>
        /// Create a new site.
        /// </summary>
        /// <param name="targetUrl">rootsite + "/" + managedPath + "/" + sitename: e.g. "https://auto.contoso.com/sites/site1"</param>
        /// <param name="title">site title: e.g. "Test Site"</param>
        /// <param name="owner">site owner: e.g. admin@contoso.com</param>
        /// <param name="siteTemplate">The site template used to create this new site</param>
        private static void CreateSiteCollection(ClientContext context, string rootUrl, string targetUrl, 
            string title, int language, int timeZoneId, string owner, 
            string siteTemplate, int userCodeMaximumLevel)
        {
            using (var ctx = new ClientContext(rootUrl.Replace(".sharepoint.com", "-admin.sharepoint.com")))
            {
                ctx.Credentials = context.Credentials;
                var tenant = new Tenant(ctx);
                SPOSitePropertiesEnumerable spp = tenant.GetSiteProperties(0, true);
                ctx.Load(spp);
                ctx.ExecuteQuery();
                bool found = false;
                foreach (SiteProperties sp in spp)
                {
                    if (sp.Url == targetUrl)
                    {
                        found = true;
                        break;
                    }
                }

                if (!found)
                {
                    System.Console.Out.Write("Creating Site Collection ");

                    //Create new site collection
                    var newsite = new SiteCreationProperties()
                    {
                        Url = targetUrl,
                        Lcid = (uint)language,
                        Owner = owner,
                        Template = siteTemplate,
                        Title = title,
                        TimeZoneId = timeZoneId,
                        UserCodeMaximumLevel = userCodeMaximumLevel
                    };

                    var spoOperation = tenant.CreateSite(newsite);

                    ctx.Load(spoOperation);
                    ctx.ExecuteQuery();

                    while (!spoOperation.IsComplete)
                    {
                        Thread.Sleep(2000);
                        ctx.Load(spoOperation);
                        ctx.ExecuteQuery();
                        System.Console.Out.Write(".");
                    }

                    System.Console.Out.WriteLine("");
                }
            }
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
                ptai.MessagesDelegate = delegate (string message, ProvisioningMessageType messageType)
                {
                    System.Console.WriteLine("      {0}: {1}", messageType.ToString(), message);
                };
                ptai.HandlersToProcess = Handlers.All;

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
            System.Console.Out.WriteLine("User: " + (string)parser.ClParameters[Params.User]);
            switch ((LoginMethod)parser.ClParameters[Params.Loginmethod])
            {
                case LoginMethod.Spo:

                    if (securePwd == null) securePwd = GetSecurePassword(parser);
                    context.AuthenticationMode = ClientAuthenticationMode.Default;
                    context.Credentials =
                        new SharePointOnlineCredentials(
                            (string)parser.ClParameters[Params.User],
                            securePwd);

                    break;
                case LoginMethod.Onprem:

                    if (parser.ClParameters.ContainsKey(Params.User))
                    {
                        if (securePwd == null) securePwd = GetSecurePassword(parser);
                        context.Credentials =
                            new NetworkCredential(
                                (string)parser.ClParameters[Params.User],
                                securePwd);
                    }
                    else
                    {
                        context.Credentials = CredentialCache.DefaultNetworkCredentials;
                    }

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
            Thread.CurrentThread.CurrentCulture = siteCulture;
            Thread.CurrentThread.CurrentUICulture = siteCulture;
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
            context.Load(context.Site.RootWeb);
            context.Load(context.Web.RegionalSettings);
            context.Load(context.Web.RegionalSettings.TimeZone, tz => tz.Id);
            context.ExecuteQuery();
            PeopleManager peopleManager = new PeopleManager(context);
            int lcid = (int)context.Site.RootWeb.Language;
            CultureInfo siteCulture = CultureInfo.GetCultureInfo(lcid);
            Thread.CurrentThread.CurrentCulture = siteCulture;
            Thread.CurrentThread.CurrentUICulture = siteCulture;
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
                            System.Console.Out.WriteLine("");
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

            return secPwd;
        }

    }
}

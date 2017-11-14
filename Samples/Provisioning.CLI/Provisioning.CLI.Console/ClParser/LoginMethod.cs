using System.Collections.Generic;

namespace Provisioning.CLI.Console.ClParser
{
    /// <summary>
    /// The supported login methods
    /// </summary>
    public enum LoginMethod
    {
        /// <summary>
        /// SharePoint Online login
        /// </summary>
        Spo,

        Onprem
    }

    /// <summary>
    /// Holds the comments to the enums for usage output
    /// </summary>
    public class LoginMethodComments
    {
        public static string Comment = "The supported login methods";
        public static Dictionary<LoginMethod, string> ValueComments = new Dictionary<LoginMethod, string>()
        {
            { LoginMethod.Spo, "SharePoint Online login"},
            { LoginMethod.Onprem, "SharePoint on-premises login"}
        };
    }
}

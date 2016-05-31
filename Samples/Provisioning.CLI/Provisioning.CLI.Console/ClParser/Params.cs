using System.Collections.Generic;

namespace Provisioning.CLI.Console.ClParser
{
    /// <summary>
    /// Possible command line parameters
    /// </summary>
    public enum Params
    {
        /// <summary>
        /// The action to be done
        /// </summary>
        /// <see cref="Actions"/>
        Action,

        /// <summary>
        /// The url from the SharePoint web
        /// </summary>
        Url,

        /// <summary>
        /// The login method to be used
        /// </summary>
        /// <see cref="LoginMethod"/>
        Loginmethod,

        /// <summary>
        /// The login user
        /// </summary>
        User,

        /// <summary>
        /// The login password. We will prompt for it if not specified
        /// </summary>
        Password,

        /// <summary>
        /// The file containing the login password. We will prompt for it if not specified
        /// </summary>
        Passwordfile,

        /// <summary>
        /// The file or path the extracted template has to be stored into
        /// </summary>
        Outfile,

        /// <summary>
        /// The input file to apply a template. Could be a template itself or a xml containing references to templates
        /// </summary>
        Infile
    }

    /// <summary>
    /// Holds the comments to the enums for usage output
    /// </summary>
    public class ParamsComments
    {
        public static string Comment = "Possible command line parameters";
        public static Dictionary<Params, string> ValueComments = new Dictionary<Params, string>()
        {
            { Params.Action, "The action to be done"},
            { Params.Loginmethod, "The login method to be used"},
            { Params.User, "The login user"},
            { Params.Password, "The login password. We will prompt for it if not specified"},
            { Params.Passwordfile, "The file containing the login password. We will prompt for it if not specified"},
            { Params.Url, "The url from the SharePoint web"},
            { Params.Outfile, "The file or path the extracted template has to be stored into"},
            { Params.Infile, "The input file to apply a template. Could be a template itself or a xml containing references to templates"}
        };
    }
}

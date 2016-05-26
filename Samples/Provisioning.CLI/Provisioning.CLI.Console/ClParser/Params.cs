namespace Provisioning.CLI.Console.ClParser
{
    /// <summary>
    /// Possible command line parameters
    /// </summary>
    public enum Params
    {
        /// <summary>
        /// The action to be done
        /// <see cref="Actions"/>
        /// </summary>
        Action,

        /// <summary>
        /// The url from the SharePoint web
        /// </summary>
        Url,

        /// <summary>
        /// The login method to be used
        /// </summary>
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
        /// The file the extracted template has to be stored into
        /// </summary>
        Outfile,

        /// <summary>
        /// The input file to apply a template. Could be a template itself or a xml containing references to templates
        /// </summary>
        Infile
    }
}

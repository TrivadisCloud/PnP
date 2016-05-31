using System.Collections.Generic;

namespace Provisioning.CLI.Console.ClParser
{
    /// <summary>
    /// The actions this tool can handle
    /// </summary>
    public enum Actions
    {
        /// <summary>
        /// Extracts a template
        /// </summary>
        Extracttemplate,

        /// <summary>
        /// Applies a template
        /// </summary>
        Applytemplate
    }

    /// <summary>
    /// Holds the comments to the enums for usage output
    /// </summary>
    public class ActionsComments
    {
        public static string Comment = "The actions this tool can handle";
        public static Dictionary<Actions, string> ValueComments = new Dictionary<Actions, string>()
        {
            { Actions.Extracttemplate, "Extracts a template"},
            { Actions.Applytemplate, "Applies a template"}
        };
    }
}

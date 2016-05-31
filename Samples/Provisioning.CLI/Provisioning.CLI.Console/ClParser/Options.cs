using System.Collections.Generic;

namespace Provisioning.CLI.Console.ClParser
{
    /// <summary>
    /// Possible command line options
    /// </summary>
    public enum Options
    {
        /// <summary>
        /// Speciefies that the tool should not prompt
        /// </summary>
        NoInteraction,

        /// <summary>
        /// Specifies, that the entire structure has to be extracted
        /// </summary>
        Entirestructure
    }

    /// <summary>
    /// Holds the comments to the enums for usage output
    /// </summary>
    public class OptionsComments
    {
        public static string Comment = "Possible command line options";
        public static Dictionary<Options, string> ValueComments = new Dictionary<Options, string>()
        {
            { Options.NoInteraction, "Speciefies that the tool should not prompt"},
            { Options.Entirestructure, "Specifies, that the entire structure has to be extracted"}
        };
    }
}

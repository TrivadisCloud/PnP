using System;
using System.Collections.Generic;
using System.Threading;

namespace Provisioning.CLI.Console.ClParser
{
    /// <summary>
    /// The parser parses the command line arguments and validates them
    /// </summary>
    public class Parser
    {
        /// <summary>
        /// Member to the provided arguments
        /// </summary>
        string[] args = null;

        /// <summary>
        /// Constructor to pass in the arguments
        /// </summary>
        /// <param name="args">Given command line arguments</param>
        public Parser(string[] args)
        {
            this.args = args;
            ClIsOk = false;
            ClParameters = new Dictionary<Params, object>();
            ClOptions = new List<Options>();
            Parse();
        }

        /// <summary>
        /// After parsing the command line this property specifies, if the parsing was successfull
        /// </summary>
        public bool ClIsOk { get; set; }

        /// <summary>
        /// Property holding all found parameters
        /// </summary>
        public Dictionary<Params, object> ClParameters { get; set; }

        /// <summary>
        /// Property holding all found options
        /// </summary>
        public List<Options> ClOptions { get; set; }

        /// <summary>
        /// Prints out the usage of this tool
        /// </summary>
        public void Usage()
        {
            string usage = "Usage:\n";
            usage += "   " + System.AppDomain.CurrentDomain.FriendlyName + " ";
            string[] parms = Enum.GetNames(typeof(Params));
            foreach (string par in parms)
            {
                usage += "-" + par + " \"value\" ";
            }
            string[] opts = Enum.GetNames(typeof(Options));
            foreach (string opt in opts)
            {
                usage += "-" + opt + " ";
            }
            usage += "\n";
            usage += "     " + ParamsComments.Comment + "\n";
            foreach (Params par in ParamsComments.ValueComments.Keys)
                usage += "               -  " + par.ToString() + ": " + ParamsComments.ValueComments[par] + "\n";
            usage += "     " + OptionsComments.Comment + "\n";
            foreach (Options opt in OptionsComments.ValueComments.Keys)
                usage += "               -  " + opt.ToString() + ": " + OptionsComments.ValueComments[opt] + "\n";
            usage += "     Action: " + ActionsComments.Comment + "\n";
            foreach (Actions act in ActionsComments.ValueComments.Keys)
                usage += "               -  " + act.ToString() + ": " + ActionsComments.ValueComments[act] + "\n";
            usage += "     LoginMethod: " + LoginMethodComments.Comment + "\n";
            foreach (LoginMethod log in LoginMethodComments.ValueComments.Keys)
                usage += "               -  " + log.ToString() + ": " + LoginMethodComments.ValueComments[log] + "\n";
            System.Console.Error.WriteLine(usage);
        }

        /// <summary>
        /// Parses the command line
        /// </summary>
        private void Parse()
        {
            if (args.Length == 0) return;
            bool foundError = false;
            for (int i = 0; i < args.Length; i++)
            {
                if (args[i].StartsWith("/") || args[i].StartsWith("-"))
                {
                    string param = args[i].Substring(1);
                    string value = null;
                    if (param.Contains("="))
                    {
                        value = param.Substring(param.IndexOf("=") + 1);
                        param = param.Substring(0, param.IndexOf("="));
                    }
                    param = Thread.CurrentThread.CurrentCulture.TextInfo.ToTitleCase(param.ToLower());
                    Options option = Options.NoInteraction;
                    Params parameter = Params.Action;
                    if (Enum.TryParse<Options>(param, out option))
                    {
                        //Option found
                        if (!ClOptions.Contains(option))
                            ClOptions.Add(option);
                    }
                    else if (Enum.TryParse<Params>(param, out parameter))
                    {
                        //Parameter found
                        if (value == null)
                        {
                            if ((i + 1) >= args.Length)
                            {
                                foundError = true;
                                System.Console.Error.WriteLine("Parameter missing value: " + param);
                            }
                            else
                            {
                                value = args[i + 1];
                                i++;
                            }
                        }
                        if (!ClParameters.ContainsKey(parameter))
                            ClParameters.Add(parameter, value);
                        switch (parameter)
                        {
                            case Params.Action:

                                Actions action = Actions.Extracttemplate;
                                value = Thread.CurrentThread.CurrentCulture.TextInfo.ToTitleCase(value.ToLower());
                                if (Enum.TryParse<Actions>(value, out action))
                                {
                                    ClParameters[parameter] = action;
                                }
                                else
                                {
                                    //Unknown action
                                    foundError = true;
                                    System.Console.Error.WriteLine("Unknown action found: " + value);
                                }

                                break;
                            case Params.Loginmethod:

                                LoginMethod loginMethod = LoginMethod.Spo;
                                value = Thread.CurrentThread.CurrentCulture.TextInfo.ToTitleCase(value.ToLower());
                                if (Enum.TryParse<LoginMethod>(value, out loginMethod))
                                {
                                    ClParameters[parameter] = loginMethod;
                                }
                                else
                                {
                                    //Unknown loginMethod
                                    foundError = true;
                                    System.Console.Error.WriteLine("Unknown loginMethod found: " + value);
                                }

                                break;
                        }
                    }
                    else
                    {
                        //Unknown parameter
                        foundError = true;
                        System.Console.Error.WriteLine("Unknown parameter found at position " + args[i]);
                    }
                }
                else
                {
                    foundError = true;
                    System.Console.Error.WriteLine("Expected a parameter name with leading - or / at position " + args[i]);
                }
            }

            //Checking required fields
            if (!ClParameters.ContainsKey(Params.Action))
            {
                foundError = true;
                System.Console.Error.WriteLine("Parameter action is required!");
            }
            else
            {
                switch ((Actions)ClParameters[Params.Action])
                {
                    case Actions.Extracttemplate:
                        foundError = CheckOutFile(foundError);
                        foundError = CheckUrl(foundError);
                        foundError = CheckLoginMethod(foundError);
                        break;
                    case Actions.Applytemplate:
                        foundError = CheckInFile(foundError);
                        foundError = CheckLoginMethod(foundError);
                        break;
                    case Actions.Applysequence:
                        foundError = CheckInFile(foundError);
                        foundError = CheckLoginMethod(foundError);
                        break;
                }
            }
            if (!ClParameters.ContainsKey(Params.Loginmethod))
            {
                switch ((LoginMethod)ClParameters[Params.Loginmethod])
                {
                    case LoginMethod.Spo:
                        if (!ClParameters.ContainsKey(Params.User))
                        {
                            foundError = true;
                            System.Console.Error.WriteLine("Parameter user is required for LoginMethod SPO!");
                        }
                        break;
                }
            }

            //Done
            if (foundError) return;
            ClIsOk = true;
        }

        /// <summary>
        /// Checks if the OutFile argument is specified
        /// </summary>
        /// <param name="foundError">The actual state</param>
        /// <returns>foundError: set to true, if there was an error</returns>
        private bool CheckOutFile(bool foundError)
        {
            if (!ClParameters.ContainsKey(Params.Outfile))
            {
                foundError = true;
                System.Console.Error.WriteLine("Parameter outFile is required!");
            }
            return foundError;
        }

        /// <summary>
        /// Checks if the InFile argument is specified
        /// </summary>
        /// <param name="foundError">The actual state</param>
        /// <returns>foundError: set to true, if there was an error</returns>
        private bool CheckInFile(bool foundError)
        {
            if (!ClParameters.ContainsKey(Params.Infile))
            {
                foundError = true;
                System.Console.Error.WriteLine("Parameter inFile is required!");
            }
            return foundError;
        }

        /// <summary>
        /// Checks if the Url argument is specified
        /// </summary>
        /// <param name="foundError">The actual state</param>
        /// <returns>foundError: set to true, if there was an error</returns>
        private bool CheckUrl(bool foundError)
        {
            if (!ClParameters.ContainsKey(Params.Url))
            {
                foundError = true;
                System.Console.Error.WriteLine("Parameter url is required!");
            }
            return foundError;
        }

        /// <summary>
        /// Checks if the LoginMethod argument is specified
        /// </summary>
        /// <param name="foundError">The actual state</param>
        /// <returns>foundError: set to true, if there was an error</returns>
        private bool CheckLoginMethod(bool foundError)
        {
            if (!ClParameters.ContainsKey(Params.Loginmethod))
            {
                foundError = true;
                System.Console.Error.WriteLine("Parameter LoginMethod is required!");
            }
            return foundError;
        }
    }
}

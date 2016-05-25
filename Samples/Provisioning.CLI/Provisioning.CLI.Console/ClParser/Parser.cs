using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace Provisioning.CLI.Console.ClParser
{
    public class Parser
    {
        string[] args = null;

        public Parser(string[] args)
        {
            this.args = args;
            ClIsOk = false;
            ClParameters = new Dictionary<Params, object>();
            ClOptions = new List<Options>();
            RequiredOptions = new List<Options>();
            RequiredParameters = new List<Params>();
            RequiredParameters.Add(Params.Action);
            Parse();
        }

        public bool ClIsOk { get; set; }
        public Dictionary<Params, object> ClParameters { get; set; }
        public List<Options> ClOptions { get; set; }
        private List<Params> RequiredParameters { get; set; }
        private List<Options> RequiredOptions { get; set; }

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
            usage += "     Action: possible values are\n";
            string[] acts = Enum.GetNames(typeof(Actions));
            foreach (string act in acts)
                usage += "               -  " + act + "\n";
            usage += "     LoginMethod: possible values are\n";
            string[] logs = Enum.GetNames(typeof(Actions));
            foreach (string log in logs)
                usage += "               -  " + log + "\n";
            System.Console.Error.WriteLine(usage);
        }

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
                    Options option = Options.Debug;
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

                                Actions action = Actions.ExtractTemplate;
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

                                LoginMethod loginMethod = LoginMethod.SPO;
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
                    case Actions.ExtractTemplate:
                        if (!ClParameters.ContainsKey(Params.Outfile))
                        {
                            foundError = true;
                            System.Console.Error.WriteLine("Parameter outFile is required!");
                        }
                        if (!ClParameters.ContainsKey(Params.Url))
                        {
                            foundError = true;
                            System.Console.Error.WriteLine("Parameter url is required!");
                        }
                        if (!ClParameters.ContainsKey(Params.Loginmethod))
                        {
                            foundError = true;
                            System.Console.Error.WriteLine("Parameter LoginMethod is required!");
                        }
                        else
                        {
                            switch ((LoginMethod)ClParameters[Params.Loginmethod])
                            {
                                case LoginMethod.SPO:
                                    if (!ClParameters.ContainsKey(Params.User))
                                    {
                                        foundError = true;
                                        System.Console.Error.WriteLine("Parameter user is required!");
                                    }
                                    break;
                            }
                        }
                        break;
                }
            }

            if (foundError) return;
            ClIsOk = true;
        }

    }
}

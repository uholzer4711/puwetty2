using System;
using System.Linq;
using SuperPUWEtty2.Data;
using log4net;
using System.Diagnostics;
using System.IO;
using System.Text.RegularExpressions;

namespace SuperPUWEtty2.Utils
{
    public class PuttyStartInfo
    {
        private static readonly ILog Log = LogManager.GetLogger(typeof(PuttyStartInfo));

        private static readonly Regex regExEnvVars = new Regex(@"(%\w+%)");

        public static String GetExecutable(SessionData session)
        {
            switch (session.Proto)
            {
                case ConnectionProtocol.Mintty:
                    return TryParseEnvVars(SuperPUWEtty2.Settings.MinttyExe);

                case ConnectionProtocol.VNC:
                    return TryParseEnvVars(SuperPUWEtty2.Settings.VNCExe);

                default:
                    return TryParseEnvVars(SuperPUWEtty2.Settings.PuttyExe);
            }
        }

        public PuttyStartInfo(SessionData session)
        {
            string argsToLog = null;

            this.Executable = GetExecutable(session);

            if (session.Proto == ConnectionProtocol.Cygterm)
            {
                CygtermStartInfo cyg = new CygtermStartInfo(session);
                this.Args = cyg.Args;
                this.WorkingDir = cyg.StartingDir;
            }
            else if (session.Proto == ConnectionProtocol.Mintty)
            {
                MinttyStartInfo mintty = new MinttyStartInfo(session);
                this.Args = mintty.Args;
                this.WorkingDir = mintty.StartingDir;
            }
            else if (session.Proto == ConnectionProtocol.VNC)
            {
                VNCStartInfo vnc = new VNCStartInfo(session);
                this.Args = vnc.Args;
                this.WorkingDir = vnc.StartingDir;
            }
            else
            {
                this.Args = MakeArgs(session, true);
                argsToLog = MakeArgs(session, false);
            }

            // attempt to parse env vars
            this.Args = this.Args.Contains('%') ? TryParseEnvVars(this.Args) : this.Args;

            Log.InfoFormat("PUWEtty2 Args: '{0}'", argsToLog ?? this.Args);
        }

        static string MakeArgs(SessionData session, bool includePassword)
        {
            if (!String.IsNullOrEmpty(session.Password) && includePassword && !SuperPUWEtty2.Settings.AllowPlainTextPuttyPasswordArg)
                Log.Warn("SuperPUWEtty2 is set to NOT allow the use of the -pw <password> argument, this can be overriden in Tools -> Options -> GUI");

            string args = "-" + session.Proto.ToString().ToLower() + " ";
            args += !String.IsNullOrEmpty(session.Password) && session.Password.Length > 0 && SuperPUWEtty2.Settings.AllowPlainTextPuttyPasswordArg 
                ? "-pw " + (includePassword ? session.Password : "XXXXX") + " " 
                : "";
            args += "-P " + session.Port + " ";
            args += !String.IsNullOrEmpty(session.PuttySession) ? "-load \"" + session.PuttySession + "\" " : "";

            args += !String.IsNullOrEmpty(SuperPUWEtty2.Settings.PuttyDefaultParameters) ? SuperPUWEtty2.Settings.PuttyDefaultParameters + " " : "";

            //If extra args contains the password, delete it (it's in session.password)
            string extraArgs = CommandLineOptions.replacePassword(session.ExtraArgs,"");            
            args += !String.IsNullOrEmpty(extraArgs) ? extraArgs + " " : "";
            args += !String.IsNullOrEmpty(session.Username) && session.Username.Length > 0 ? " -l " + session.Username + " " : "";
            args += session.Host;

            return args;
        }

        static string TryParseEnvVars(string args)
        {
            string result = args;
            try
            {
                result = regExEnvVars.Replace(
                    args,
                    m =>
                    {
                        string name = m.Value.Trim('%');
                        return Environment.GetEnvironmentVariable(name) ?? m.Value;
                    });
            }
            catch(Exception ex)
            {
                Log.Warn("Could not parse env vars in args", ex);
            }
            return result;
        }

        /// <summary>
        /// Start of standalone putty process
        /// </summary>
        public void StartStandalone()
        {
            ProcessStartInfo startInfo = new ProcessStartInfo
            {
                FileName = this.Executable,
                Arguments = this.Args
            };
            if (this.WorkingDir != null && Directory.Exists(this.WorkingDir))
            {
                startInfo.WorkingDirectory = this.WorkingDir;
            }
            Process.Start(startInfo);
        }

        public SessionData Session { get; private set; }

        public string Args { get; private set; }
        public string WorkingDir { get; private set; }
        public string Executable { get; private set; }

    }
}

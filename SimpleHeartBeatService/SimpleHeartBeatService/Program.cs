using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Topshelf;

namespace SimpleHeartBeatService
{
    class Program
    {
        static void Main(string[] args)
        {
            var exitCode = HostFactory.Run(x =>
            {
                x.Service<HeartBeat>(s =>
                {
                    s.ConstructUsing(heartbeat => new HeartBeat());
                    s.WhenStarted(heartbeat => heartbeat.Start());
                    s.WhenStopped(heartbeat => heartbeat.Stop());
                });
                x.RunAsLocalSystem();
                x.SetServiceName("SendMail");
                x.SetDisplayName("SendMail Service");
                x.SetDescription("For send EOD Mail");
            });
            int exitCodeValue = (int)Convert.ChangeType(exitCode, exitCode.GetTypeCode());
            Environment.ExitCode = exitCodeValue;
        }
    }
}

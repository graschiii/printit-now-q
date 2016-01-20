using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.ServiceProcess;

namespace Server
{
    public static class Program
    {
        #region Nested class to support running as a service
        public const string ServiceName = "ReportProcessorService";

        public class Service : ServiceBase
        {
            public Service()
            {
                ServiceName = Program.ServiceName;
            }

            protected override void OnStart(string[] args)
            {
                Program.Start(args);
            }

            protected override void OnStop()
            {
                Program.Stop();
            }
        }
        #endregion

        static void Main(string[] args)
        {
            
            if (!Environment.UserInteractive)
            {
                // running as service
                using (var service = new Service())
                    ServiceBase.Run(service);
            }
            else
            {
                // Running from the console
                Start(args);
                Console.WriteLine("Starting RabbitMQ queue processor");               
                Console.ReadLine();
                Stop();
            }
            
        }
        private static void Start(string[] args)
        {
            // onstart code here
            var queueProcessor = new RabbitConsumer() { Enabled = true };
            queueProcessor.Start();
        }

        private static void Stop()
        {
            // onstop code here
        }
    }
}

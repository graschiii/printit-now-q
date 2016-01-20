using System;
using System.Text;
using RabbitMQ.Client;
using RabbitMQ.Client.Events;

namespace Server
{
    /// <summary>
    /// Class to encapsulate receiving messages from RabbitMQ
    /// </summary>
    public class RabbitConsumer : IDisposable
    {
        private const string HostName = "localhost";
        private const string UserName = "guest";
        private const string Password = "guest";
        private const string QueueName = "Module2.Sample1";
        private const string ExchangeName = "";
        private const bool IsDurable = true;
        //The two below settings are just to illustrate how they can be used but we are not using them in
        //this sample as we will use the defaults
        private const string VirtualHost = "";
        private int Port = 0;

        public delegate void OnReceiveMessage(string message);

        public bool Enabled { get; set; }
    
        private ConnectionFactory _connectionFactory;
        private IConnection _connection;
        private IModel _model;
        

        /// <summary>
        /// Ctor with a key to lookup the configuration
        /// </summary>
        public RabbitConsumer()
        {
            // If running interactively display the settings
            if (Environment.UserInteractive)
            {
            DisplaySettings();
            }
            _connectionFactory = new ConnectionFactory
            {
                HostName = HostName,
                UserName = UserName,
                Password = Password
            };

            if (string.IsNullOrEmpty(VirtualHost) == false)
                _connectionFactory.VirtualHost = VirtualHost;
            if (Port > 0)
                _connectionFactory.Port = Port;

            _connection = _connectionFactory.CreateConnection();
            _model = _connection.CreateModel();
            _model.BasicQos(0, 1, false);
        }
        /// <summary>
        /// Displays the rabbit settings
        /// </summary>
        private void DisplaySettings()
        {
            Console.WriteLine("Host: {0}", HostName);
            Console.WriteLine("Username: {0}", UserName);
            Console.WriteLine("Password: {0}", Password);
            Console.WriteLine("QueueName: {0}", QueueName);
            Console.WriteLine("ExchangeName: {0}", ExchangeName);
            Console.WriteLine("VirtualHost: {0}", VirtualHost);
            Console.WriteLine("Port: {0}", Port);
            Console.WriteLine("Is Durable: {0}", IsDurable);
        }
        /// <summary>
        /// Starts receiving a message from a queue
        /// </summary>
        public void Start()
        {
            var consumer = new QueueingBasicConsumer(_model);
            _model.BasicConsume(QueueName, false, consumer);

            while (Enabled)
            {
                // Get next message
                var deliveryArgs = (BasicDeliverEventArgs)consumer.Queue.Dequeue();

                // Serialize message
                var message = Encoding.Default.GetString(deliveryArgs.Body);
                
                // If running interactively display the message received
                if (Environment.UserInteractive)
                {
                    Console.WriteLine("Message Received - {0}", message);
                }
                // Acknowledge the message
                _model.BasicAck(deliveryArgs.DeliveryTag, false);
                // Create the report request instance
                if (message != "QUIT!")
                    {
                    var reportRequest = new ReportGenerate();
                    // Initialize the parameters for the report
                    reportRequest.initParameterValues(message);
                    // Generate the report
                    reportRequest.GenerateReport();
                    // Put it on the disk
                    reportRequest.SaveReportToDisk();
                    }
                else
                {
                    Enabled = false;
                    Environment.Exit(0);
                }

            }
        }
        /// <summary>
        /// Dispose
        /// </summary>
        public void Dispose()
        {
            if (_model != null)
                _model.Dispose();
            if (_connection != null)
                _connection.Dispose();

            _connectionFactory = null;

            GC.SuppressFinalize(this);
        }
    }
}

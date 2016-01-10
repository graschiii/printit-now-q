using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using RabbitMQ.Client;

namespace Client
{
    class Program
    {        
        private static void Main(string[] args)
        {

            Console.WriteLine("Starting RabbitMQ Message Sender");
            Console.WriteLine();
            Console.WriteLine();            

            var messageTxt = "";
            // Some message are issued by the constructor on the sender
            var sender = new RabbitSender();

            
            while (true)
            {
                Console.WriteLine("Enter the message to be sent");
                messageTxt = Console.ReadLine();
                if (messageTxt == "exit")
                {
                    break;
                }
                sender.Send(messageTxt);
                Console.WriteLine("Message has been sent");
            }
            
            //Console.ReadLine();
        }
    }
}

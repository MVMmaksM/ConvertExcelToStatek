using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ConvertExcelToStatek.Services
{
    public class MessageServices : IMessageServices
    {
        public void SendMessage(string message)
        {
            Console.WriteLine(message);
        }
    }
}

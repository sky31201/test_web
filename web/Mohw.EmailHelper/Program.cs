using MimeKit;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Mohw.EmailHelper
{
    internal class Program
    {
        static void Main(string[] args)
        {
            // 讀取信件
            List<MimeMessage> messages = ReadEmail();
        }
    }
}

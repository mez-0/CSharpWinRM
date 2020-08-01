using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace CSharpWinRM
{
    class Logger
        {
            public enum STATUS
            {
                GOOD = 0,
                ERROR = 1,
                INFO = 2
            };
            public static void Print(STATUS status, String msg)
            {
                if (status == STATUS.GOOD)
                {
                    Console.WriteLine("[+] " + msg);
                }
                else if (status == STATUS.ERROR)
                {
                    Console.WriteLine("[-] " + msg);
                }
                else if (status == STATUS.INFO)
                {
                    Console.WriteLine("[*] " + msg);
                }
            }
    }
}
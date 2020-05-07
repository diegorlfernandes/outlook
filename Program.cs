using EasyConsole;
using Microsoft.Office.Interop.Outlook;
using System;
using System.Collections.Generic;
using System.Linq;

namespace outlook
{
    class Program
    {
        private static string Sender = "";
        private static string Subject = "";
        private static EmailSearch ES = new EmailSearch();
        static void Main(string[] args = null)
        {

            if (args.Length == 0)
            {
                Console.WriteLine("Informe pelo menos um parâmentro.");
                return;
            }
            else if (args[0] == "-it")
            {
                Console.WriteLine("indexar todos");
                ES.Indexar((Enum.TiposProcessamentos.Todos_Emails));
                return;
            }
            else if (args[0] == "-ic")
            {
                Console.WriteLine("indexar caixa entrada");
                ES.Indexar((Enum.TiposProcessamentos.Caixa_Entrada));
                return;
            }
            else if (args[0] == "-h")
            {
                Console.WriteLine("opções");
                Console.WriteLine("-it = indexar todas as pastas de e-mail");
                Console.WriteLine("-ic = indexar caixa de entrada");
                Console.WriteLine("arg[0] = Sender, arg[1] = Subject");
                Console.WriteLine("-h = ajuda");
                return;

            }
            else
            {
                if (args.Count() == 1)
                    ES.Imprimir(args[0]);
                else if (args.Count() == 2)
                    ES.Imprimir(args[0], args[1]);
                else
                    Console.WriteLine("o programa recebe apenas dois argumentos");

                return;
            }
        }

    }
}

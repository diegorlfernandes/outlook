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
        private static  EmailSearch ES = new EmailSearch();
        static void Main(string[] args)
        {
            Console.WriteLine("V2");

            //1 - Todos os Emails
            //2 - Carga da caixa de entrada
            //3 - interface

            string res;

            //res = "1";
            res = args[0];

            Int16 opt = Convert.ToInt16(res);

            if(opt != 3)
                ES.Indexar((Enum.TiposProcessamentos)opt);


            if (opt == 3)
            {

                Filtrar(); 

                ES.Imprimir(Sender, Subject);


                ConsoleKeyInfo cki;

                do
                {
                    Console.WriteLine();
                    Console.WriteLine("Opções:");
                    Console.WriteLine("1-Abrir / 2-Pesquisar / Esc-Sair");

                    cki = Console.ReadKey(false); // show the key as you read it
                    switch (cki.KeyChar.ToString())
                    {
                        case "1":
                            Abrir();
                            break;
                        case "2":
                            Console.WriteLine();
                            Filtrar();
                            ES.Imprimir(Sender, Subject);
                            break;
                            // etc..
                    }
                } while (cki.Key != ConsoleKey.Escape);

            }

        }

        static void  Filtrar()
        {
            Console.Clear();
            //Console.WriteLine("**********FILTROS**************");
            Console.WriteLine("                 FILTROS                   ");
            Console.WriteLine("===========================================");
            Console.SetCursorPosition(0, 10);
            Console.WriteLine("===========================================");
            Console.WriteLine();
            Console.SetCursorPosition(0, 4);
            Console.Write("    Sender: ");

            Sender = Console.ReadLine().ToUpper();

            Console.WriteLine();
            Console.Write("    Subject: ");
            Subject = Console.ReadLine().ToUpper();



        }

        static void Abrir()
        {
            Console.Write("Digite o ID do e-mail para abrir: ");

            //Int16 ID = Convert.ToInt16(Console.ReadLine());
            int ID = Input.ReadInt("Digite o ID do e-mail para abrir: ",1,99999);


            ES.Abrir(ID);

            

        }

    }
}

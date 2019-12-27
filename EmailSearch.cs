using Microsoft.Office.Interop.Outlook;
using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace outlook
{
    public class EmailSearch
    {
        private static Microsoft.Office.Interop.Outlook.Application app;
        private static NameSpace outlookNs;
        private static Dictionary<int, string> controle;
        private static Int16 ordem;
        private static Int16 contador;


        public EmailSearch()
        {
            app = new Microsoft.Office.Interop.Outlook.Application();
            outlookNs = app.GetNamespace("MAPI");

        }

        public List<MailItem> Ler(Enum.TiposProcessamentos tiposProcessamentos)
        {

            contador = 1;

            List<MailItem> mailItems = new List<MailItem>();

            foreach (Store store in outlookNs.Stores)
            {
                MAPIFolder rootFolder = store.GetRootFolder();

                Folders subFolders = rootFolder.Folders;

                foreach (Folder folder in subFolders)
                {
                    if (tiposProcessamentos == Enum.TiposProcessamentos.Todos_Emails)
                    {
                        if (folder.Name.Contains("Entrada") || folder.Name.Contains("Enviado") || folder.Name.Contains("Recebido") || folder.Name.Contains("Arquivo"))
                        {
                            ObterEmails(mailItems, folder);

                        }
                    }
                    else if (tiposProcessamentos == Enum.TiposProcessamentos.Caixa_Entrada)
                    {
                        if (folder.Name.Contains("Entrada"))
                        {
                            ObterEmails(mailItems, folder);
                        }
                    }

                }

            }
            return mailItems;
        }

        private static void ObterEmails(List<MailItem> mailItems, Folder folder)
        {
            Items items = folder.Items;


            foreach (object item in items)
            {
                try
                {
                    if (item is MailItem)
                    {
                        MailItem mailItem = item as MailItem;

                        if (String.IsNullOrEmpty(mailItem.SenderName) || String.IsNullOrEmpty(mailItem.Subject))
                            continue;

                        mailItems.Add(mailItem);
                    }
                    //Console.SetCursorPosition(1,5);
                    //Console.Write("Processando item " + contador + " de " + items.Count);
                    contador++;

                }
                catch
                {
                    continue;
                }
            }

            foreach (Folder subfolder in folder.Folders)
            {
                ObterEmails(mailItems, subfolder);
            }
        }


        public void Imprimir(string _Sender = null, string _Subject = null)
        {
            var ListaEmails = LerTodosEmailsIdenxados(_Sender, _Subject);
            ordem = 1;

            controle = new Dictionary<int, string>();


            Console.Write("ID".PadRight(6, ' '));
            Console.Write("SenderName ".PadRight(50, ' '));
            Console.Write("Subject ".PadRight(150, ' '));
            Console.WriteLine("Date ".PadRight(20, ' '));

            foreach (var mailItem in ListaEmails)
            {
                string EntryID = mailItem["EntryID"].ToString();
                string SenderName = mailItem["SenderName"].ToString();
                string Subject = mailItem["Subject"].ToString();
                string Date = mailItem["Date"].ToString();

                controle.Add(ordem, EntryID);

                Console.Write(ordem.ToString().PadRight(6, ' '));

                if (!String.IsNullOrEmpty(SenderName))
                    Console.Write(SenderName.PadRight(50, ' '));

                if (!String.IsNullOrEmpty(Subject))
                    Console.Write(Subject.PadRight(150, ' '));

                Console.WriteLine(Date.PadRight(20, ' '));

                ordem++;

            }

        }


        public void Indexar(Enum.TiposProcessamentos tiposProcessamentos)
        {
            Console.WriteLine("Atualizando indice...");

            var list = Ler(tiposProcessamentos);
 
            DB db = new DB();


            foreach (var item in list)
                db.Inserir(item.EntryID, item.SenderName, item.Subject, item.ReceivedTime.ToString());

            db = null;

        }

        public List<DataRow> LerTodosEmailsIdenxados(string SenderName = null, string Subject = null)
        {
            DB db = new DB();

            return db.ObterTodos(SenderName, Subject);


        }

        public void Abrir(int ID)
        {
            var item = outlookNs.GetItemFromID(controle[ID]) as MailItem;
            item.Display();
        }

    }
}

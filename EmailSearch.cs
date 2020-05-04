﻿using Microsoft.Office.Interop.Outlook;
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
        private static DB db; 


        public EmailSearch()
        {
            app = new Microsoft.Office.Interop.Outlook.Application();
            outlookNs = app.GetNamespace("MAPI");
            db = new DB();

        }

        public void Indexar(Enum.TiposProcessamentos tiposProcessamentos)
        {
            Console.WriteLine("Atualizando indice...");

            contador = 1;

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
                            db.ApagarTodos();
                            GravarEmailPorPastas(folder, tiposProcessamentos);
                        }
                    }
                    else if (tiposProcessamentos == Enum.TiposProcessamentos.Caixa_Entrada)
                    {
                        if (folder.Name.Contains("Entrada"))
                        {
                            GravarEmailPorPastas(folder, tiposProcessamentos);
                        }
                    }

                }

            }

        }
 
        private static void GravarEmailPorPastas(Folder folder, Enum.TiposProcessamentos tiposProcessamentos)
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

                        if (tiposProcessamentos == Enum.TiposProcessamentos.Todos_Emails)
                            db.Inserir(mailItem.EntryID, mailItem.SenderName, mailItem.Subject, mailItem.ReceivedTime.ToString());
                        else
                            db.PesquisareInserir(mailItem.EntryID, mailItem.SenderName, mailItem.Subject, mailItem.ReceivedTime.ToString());

                    }
                    contador++;

                }
                catch
                {
                    continue;
                }
            }

            foreach (Folder subfolder in folder.Folders)
            {
                GravarEmailPorPastas(subfolder, tiposProcessamentos);
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

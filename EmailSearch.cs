using EasyConsole;
using Microsoft.Office.Interop.Outlook;
using Microsoft.VisualBasic.Logging;
using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading;
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
        private static GravarLog GravarLog;


        public EmailSearch()
        {
            app = GetApplicationObject();
            outlookNs = app.GetNamespace("MAPI");
            db = new DB();
            GravarLog = new GravarLog();

        }

        Outlook.Application GetApplicationObject()
        {

            Outlook.Application application = null;
            if (Process.GetProcessesByName("OUTLOOK").Count() > 0)
                application = Marshal.GetActiveObject("Outlook.Application") as Outlook.Application;
            else
                application = new Outlook.Application();

            return application;
        }

        public void Indexar(Enum.TiposProcessamentos tiposProcessamentos)
        {
            GravarLog.Log("Atualizando indice...");

            contador = 1;
            db.ApagarTodos();

            if (tiposProcessamentos == Enum.TiposProcessamentos.Todos_Emails)
                GravarLog.Log("Processando todos os e-mails...");
            else if (tiposProcessamentos == Enum.TiposProcessamentos.Caixa_Entrada)
                GravarLog.Log("Processando apenas caixa de entrada...");


            foreach (Store store in outlookNs.Stores)
            {
                MAPIFolder rootFolder = store.GetRootFolder();

                Folders subFolders = rootFolder.Folders;

                if (tiposProcessamentos == Enum.TiposProcessamentos.Todos_Emails)
                {
                    foreach (Folder folder in subFolders)
                        if (folder.Name.Contains("Entrada") || folder.Name.Contains("Enviado") || folder.Name.Contains("Recebido"))
                            GravarEmailPorPastas(folder, tiposProcessamentos);
                }
                else if (tiposProcessamentos == Enum.TiposProcessamentos.Caixa_Entrada)
                {
                    foreach (Folder folder in subFolders)
                        if (folder.Name.Contains("Entrada"))
                            GravarEmailPorPastas(folder, tiposProcessamentos);
                }

            }

            GravarLog.Log("Total de e-mails processados:" + Convert.ToString(contador));


        }

        private static void GravarEmailPorPastas(Folder folder, Enum.TiposProcessamentos tiposProcessamentos)
        {
            GravarLog.Log("Atualizando pasta: " + folder.Name + " Qtd itens:" + folder.Items.Count);
            Items items = folder.Items;
            db.CreateTrasaction();

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
                        {
                            //mailItem.SaveAs("d:\\emails\\");
                            db.Inserir(mailItem.EntryID, mailItem.SenderName, mailItem.Subject, mailItem.ReceivedTime.ToString());
                        }
                        else
                        {
                            //mailItem.SaveAs("d:\\emails\\");
                            db.PesquisareInserir(mailItem.EntryID, mailItem.SenderName, mailItem.Subject, mailItem.ReceivedTime.ToString());
                        }

                    }
                    contador++;


                }
                catch
                {
                    continue;
                }
            }
            db.commit();

            foreach (Folder subfolder in folder.Folders)
            {
                GravarEmailPorPastas(subfolder, tiposProcessamentos);
            }
        }


        public void Imprimir(string _Sender = "", string _Subject = "")
        {
            var ListaEmails = LerEmailsComFiltro(_Sender, _Subject);
            ordem = 1;

            controle = new Dictionary<int, string>();

            Console.Write("ID".PadRight(6, ' '));
            Console.Write("SenderName ".PadRight(50, ' '));
            Console.Write("Subject ".PadRight(100, ' '));
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
                    Console.Write(Subject.PadRight(100, ' '));

                Console.WriteLine(Date.PadRight(20, ' '));

                ordem++;

            }
            Abrir();
        }



        public List<DataRow> LerEmailsComFiltro(string SenderName = "", string Subject = "")
        {
            DB db = new DB();
            return db.ObterComFiltro(SenderName, Subject);
        }

        public void Abrir()
        {


            int ID = Input.ReadInt("Abrir: ", 1, 99999);

            Process p = Process.Start("C:\\Program Files\\Microsoft Office\\root\\Office16\\OUTLOOK.EXE");

            Thread.Sleep(600);
            var item = outlookNs.GetItemFromID(controle[ID]) as MailItem;
            item.Display();

            outlookNs = null;
            p.Close();
            p.Dispose();
        }

    }
}

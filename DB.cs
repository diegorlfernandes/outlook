
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using Microsoft.Data.Sqlite;

namespace outlook
{
    public class DB
    {
        private static SqliteConnection connection;
        private static SqliteTransaction transaction;
        public DB()
        {
            String path = System.AppDomain.CurrentDomain.BaseDirectory.ToString();
            connection = new SqliteConnection("" + new SqliteConnectionStringBuilder { DataSource = path+"DB" });
            connection.Open();
            CriarTabelas();
        }

        private void CriarTabelas()
        {
            CreateTrasaction();
            string sql = @"CREATE TABLE IF NOT EXISTS email(
                        EntryID text PRIMARY KEY,
                        SenderName text,
                        Subject text,
                        Date text);";

            SqliteCommand command = connection.CreateCommand();
            command.Transaction = transaction;
            command.CommandText = sql;
            _ = command.ExecuteNonQuery();
            commit();

        }

        public bool PesquisareInserir(string EntryID, string SenderName, string Subject, string Date)
        {
            if (ObterPorID(EntryID).Count == 0)
            {
                SqliteCommand command = connection.CreateCommand();
                command.Transaction = transaction;
                command.CommandText = "insert into email ( EntryID, SenderName, Subject, Date ) values ( $EntryID, $SenderName, $Subject, $Date ) ";
                command.Parameters.AddWithValue("$EntryID", EntryID);
                command.Parameters.AddWithValue("$SenderName", SenderName);
                command.Parameters.AddWithValue("$Subject", Subject);
                command.Parameters.AddWithValue("$Date", Date);
                int ret = command.ExecuteNonQuery();


                if (ret == 0)
                    return true;
                else
                    return false;

            }

            return true;

        }

        public void CreateTrasaction()
        {
            transaction = connection.BeginTransaction();
        }

        public void commit()
        {
            transaction.Commit();
        }


        public bool Inserir(string EntryID, string SenderName, string Subject, string Date)
        {
            SqliteCommand command = connection.CreateCommand();
            command.Transaction = transaction;
            command.CommandText = "insert into email ( EntryID, SenderName, Subject, Date ) values ( $EntryID, $SenderName, $Subject, $Date ) ";
            command.Parameters.AddWithValue("$EntryID", EntryID);
            command.Parameters.AddWithValue("$SenderName", SenderName);
            command.Parameters.AddWithValue("$Subject", Subject);
            command.Parameters.AddWithValue("$Date", Date);
            int ret = command.ExecuteNonQuery();


            if (ret == 0)
                return true;
            else
                return false;
        }


        public List<DataRow> ObterComFiltro(string SenderName = "", string Subject = "")
        {

            SqliteCommand command = connection.CreateCommand();
            command.CommandText = @"SELECT * FROM email 
                                    where 
                                    ($SenderName = '' or upper(SenderName) like $SenderName) AND
                                    ($Subject = '' or upper(Subject) like $Subject)
                                    order by substr(date,7,4)||substr(date,4,2)||substr(date,1,2) desc";

            if (String.IsNullOrEmpty(SenderName) | SenderName == "?")
                command.Parameters.AddWithValue("$SenderName", "");
            else
                command.Parameters.AddWithValue("$SenderName", "%" + SenderName + "%");

            if (String.IsNullOrEmpty(Subject) | Subject == "?")
                command.Parameters.AddWithValue("$Subject", "");
            else
                command.Parameters.AddWithValue("$Subject", "%" + Subject + "%");

            var reader = command.ExecuteReader();

            var dt = new DataTable();
            dt.Load(reader);

            List<DataRow> dr = dt.AsEnumerable().ToList();

            string query = command.CommandText;


            return dr;

        }

        public bool ApagarTodos()
        {
            SqliteTransaction transaction = connection.BeginTransaction();

            SqliteCommand command = connection.CreateCommand();
            command.CommandText = @"delete FROM email";
            int ret = command.ExecuteNonQuery();
            transaction.Commit();


            if (ret == 0)
                return true;
            else
                return false;

        }

        public List<DataRow> ObterPorID(string ID)
        {

            SqliteCommand command = connection.CreateCommand();
            command.CommandText = "SELECT * FROM email where EntryID = $EntryID";
            command.Parameters.AddWithValue("$EntryID", ID);
            var reader = command.ExecuteReader();

            var dt = new DataTable();
            dt.Load(reader);

            List<DataRow> dr = dt.AsEnumerable().ToList();


            return dr;

        }



    }
}

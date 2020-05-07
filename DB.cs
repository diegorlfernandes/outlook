
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
        public DB()
        {
            String path = System.AppDomain.CurrentDomain.BaseDirectory.ToString();
            connection = new SqliteConnection("" + new SqliteConnectionStringBuilder { DataSource = path+"DB" });
            connection.Open();
            CriarTabelas();
        }

        private void CriarTabelas()
        {
            string sql = @"CREATE TABLE IF NOT EXISTS email(
                        EntryID text PRIMARY KEY,
                        SenderName text,
                        Subject text,
                        Date text);";

            SqliteTransaction transaction = connection.BeginTransaction();
            SqliteCommand command = connection.CreateCommand();
            command.Transaction = transaction;
            command.CommandText = sql;
            _ = command.ExecuteNonQuery();
            transaction.Commit();

        }

        public bool PesquisareInserir(string EntryID, string SenderName, string Subject, string Date)
        {
            if (ObterPorID(EntryID).Count == 0)
            {
                SqliteTransaction transaction = connection.BeginTransaction();
                SqliteCommand command = connection.CreateCommand();
                command.Transaction = transaction;
                command.CommandText = "insert into email ( EntryID, SenderName, Subject, Date ) values ( $EntryID, $SenderName, $Subject, $Date ) ";
                command.Parameters.AddWithValue("$EntryID", EntryID);
                command.Parameters.AddWithValue("$SenderName", SenderName);
                command.Parameters.AddWithValue("$Subject", Subject);
                command.Parameters.AddWithValue("$Date", Date);
                int ret = command.ExecuteNonQuery();
                transaction.Commit();


                if (ret == 0)
                    return true;
                else
                    return false;

            }

            return true;

        }
        public bool Inserir(string EntryID, string SenderName, string Subject, string Date)
        {
            SqliteTransaction transaction = connection.BeginTransaction();
            SqliteCommand command = connection.CreateCommand();
            command.Transaction = transaction;
            command.CommandText = "insert into email ( EntryID, SenderName, Subject, Date ) values ( $EntryID, $SenderName, $Subject, $Date ) ";
            command.Parameters.AddWithValue("$EntryID", EntryID);
            command.Parameters.AddWithValue("$SenderName", SenderName);
            command.Parameters.AddWithValue("$Subject", Subject);
            command.Parameters.AddWithValue("$Date", Date);
            int ret = command.ExecuteNonQuery();
            transaction.Commit();


            if (ret == 0)
                return true;
            else
                return false;
        }


        public List<DataRow> ObterTodos(string SenderName = "", string Subject = "")
        {

            SqliteCommand command = connection.CreateCommand();
            command.CommandText = @"SELECT * FROM email 
                                    where 
                                    ($SenderName = '' or upper(SenderName) like $SenderName) AND
                                    ($Subject = '' or upper(Subject) like $Subject)
                                    order by Date desc";

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

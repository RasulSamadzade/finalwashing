using System;
using System.Collections.Generic;
using System.Data.SQLite;
using System.IO;
using System.Windows;

namespace WpfApp1
{
    class DatabaseHandling
    {
        SQLiteConnection sqlConnection;
        string connectionString;
        public DatabaseHandling()
        {
            connectionString = "Data Source=" + Directory.GetCurrentDirectory().Substring(0, Directory.GetCurrentDirectory().Length - 9) + "TechnoProbe.db;Version=3;New=False;Compress=True;";
            sqlConnection = new SQLiteConnection(connectionString);
        }

        public void saveToDatabase(Data parameters)
        {
            try
            {
                string query = "INSERT INTO FinalResults (Type, Name, IdCode, Layer, TopBottom, Defect, Input1, Input2, Decision, Date) values (@Type, @Name, @IdCode, @Layer, @TopBottom, @Defect, @Input1, @Input2, @Decision, @Date)";
                SQLiteCommand sqLiteCommand = new SQLiteCommand(query, sqlConnection);
                sqlConnection.Open();
                sqLiteCommand.Parameters.AddWithValue("@Type", parameters.Type);
                sqLiteCommand.Parameters.AddWithValue("@Name", parameters.Name);
                sqLiteCommand.Parameters.AddWithValue("@IdCode", parameters.ID_Code);
                sqLiteCommand.Parameters.AddWithValue("@Layer", parameters.Layer);
                sqLiteCommand.Parameters.AddWithValue("@TopBottom", parameters.TopBottom);
                sqLiteCommand.Parameters.AddWithValue("@Defect", parameters.Defect);
                sqLiteCommand.Parameters.AddWithValue("@Input1", parameters.Input1);
                sqLiteCommand.Parameters.AddWithValue("@Input2", parameters.Input2);
                sqLiteCommand.Parameters.AddWithValue("@Decision", parameters.Decision);
                sqLiteCommand.Parameters.AddWithValue("@Date", (DateTime.Now.Date + DateTime.Now.TimeOfDay).ToString());
                sqLiteCommand.ExecuteScalar();
            }
            catch (Exception ex)
            {
                MessageBox.Show("saveclick" + ex.ToString());
            }
            finally
            {
                sqlConnection.Close();
            }
        }

        public List<string[]> generateListFromDatabase()
        {
            var cellData = new List<string[]>();
            string[] cell;
            try
            {
                string query = "SELECT Type, Name, IdCode, Layer, TopBottom, Defect, Input1, Input2, Decision, Date FROM FinalResults";
                sqlConnection.Open();
                var cmd = new SQLiteCommand(query, sqlConnection);
                using (SQLiteDataReader rdr = cmd.ExecuteReader())
                {
                    while (rdr.Read())
                    {
                        cell = new string[10];
                        for (int i = 0; i < 10; i++)
                        {
                            string a = "";

                            try { a = rdr.GetString(i); }
                            catch (Exception){}

                            cell[i] = a;
                        }
                        cellData.Add(cell);
                    }
                }
                return cellData;
            }
            catch (Exception ex)
            {
                MessageBox.Show("generatelistfromdatabase" + ex.Message.ToString());
            }
            finally
            {
                sqlConnection.Close();
            }
            return cellData;
        }

    }
}

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
            try
            {
                string query = "SELECT Type, Name, IdCode, Layer, TopBottom, Defect, Input1, Input2, Decision, Date FROM FinalResults";
                sqlConnection.Open();
                String a;
                String b;
                String c;
                String d;
                String e;
                String f;
                String g;
                String h;
                String i;
                String j;
                String k;
                var cmd = new SQLiteCommand(query, sqlConnection);
                using (SQLiteDataReader rdr = cmd.ExecuteReader())
                {
                    while (rdr.Read())
                    {
                        try
                        {
                            a = rdr.GetString(0);
                        }
                        catch (Exception ex)
                        {
                            a = "";
                        }
                        try
                        {
                            b = rdr.GetString(1);
                        }
                        catch (Exception ex)
                        {
                            b = "";
                        }
                        try
                        {
                            c = rdr.GetString(2);
                        }
                        catch (Exception ex)
                        {
                            c = "";
                        }
                        try
                        {
                            d = rdr.GetString(3);
                        }
                        catch (Exception ex)
                        {
                            d = "";
                        }
                        try
                        {
                            e = rdr.GetString(4);
                        }
                        catch (Exception ex)
                        {
                            e = "";
                        }
                        try
                        {
                            f = rdr.GetString(5);
                        }
                        catch (Exception ex)
                        {
                            f = "";
                        }
                        try
                        {
                            g = rdr.GetString(6);
                        }
                        catch (Exception ex)
                        {
                            g = "";
                        }
                        try
                        {
                            h = rdr.GetString(7);
                        }
                        catch (Exception ex)
                        {
                            h = "";
                        }
                        try
                        {
                            i = rdr.GetString(8);
                        }
                        catch (Exception ex)
                        {
                            i = "";
                        }
                        try
                        {
                            j = rdr.GetString(9);
                        }
                        catch (Exception ex)
                        {
                            j = "";
                        }
                        var cell = new string[] { a, b, c, d, e, f, g, h, i, j };
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

﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.OleDb;
using System.IO;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.OleDb;
using Excel = Microsoft.Office.Interop.Excel;
using System.IO;

namespace Zebra
{
    class WorkWithDatabase
    {
        private OleDbConnection connection = new OleDbConnection();

        public WorkWithDatabase()
        {
            connection.ConnectionString = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=../../db.mdb;
                                                Persist Security Info=False;";
        }
        /// <summary>
        /// Checking if is connection
        /// </summary>
        /// <returns>enable connection</returns>
        public bool isConnection()
        {
            bool returnValue;
            try
            {
                connection.Open();
                returnValue = true;
            }
            catch (Exception ex)
            {
                returnValue = false;
            }
            finally
            {
                connection.Close();
            }
            return returnValue;
        }

        public int getIdByValue(string name, OleDbConnection connection)
        {
            int returnValue = -1;
            try
            {
                OleDbCommand command = new OleDbCommand();
                command.Connection = connection;
                string query = (@"select * from goods where [Names]='" + name.ToString() + "'");
                command.CommandText = query;

                OleDbDataReader reader = command.ExecuteReader();
                while (reader.Read())
                {
                    //returnValue = Int32.Parse(reader[0].ToString());
                    returnValue = Int32.Parse(reader["id"].ToString());
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex);
                connection.Close();
            }
            return returnValue;
        }

        public string getValueByID(string language, int id)
        {
            string returnString = "";
            try
            {
                connection.Open();
                OleDbCommand command = new OleDbCommand();
                command.Connection = connection;
                string query = "select * from dict where id=" + id + "";
                command.CommandText = query;

                OleDbDataReader reader = command.ExecuteReader();
                while (reader.Read())
                {
                    returnString = reader[language].ToString();
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex);
            }
            finally
            {
                connection.Close();
            }

            return returnString;
        }

        //not tested function yet
        public List<string> getAllWordsByName(string language)
        {
            List<string> returnList = new List<string>();
            try
            {
                connection.Open();
                OleDbCommand command = new OleDbCommand();
                command.Connection = connection;
                string query = "select " + language + " from dict";
                command.CommandText = query;

                OleDbDataReader reader = command.ExecuteReader();
                while (reader.Read())
                {
                    returnList.Add(reader[0].ToString());
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex);
            }
            finally
            {
                connection.Close();
            }
            returnList.Sort();
            return returnList;
        }

        public bool dataExist(string name, OleDbConnection connection)
        {
            return getIdByValue(name, connection) != -1;
        }

        public void insertData(string name, string size, string manufacturer, string composition, string additional)
        {
            connection.Open();
            OleDbCommand command = new OleDbCommand();
            command.Connection = connection;
            try
            {
                //not shure about value or values and breckets
                if (!dataExist(name, connection))
                {
                    command.CommandText = (@"insert into [goods] ([Names],[Sizes],[Manufacturers],[Compositions],[Additionals]) values " +
                        "('" + name.ToString() + "','" + size.ToString() + "','" + manufacturer.ToString() + "','" + composition.ToString() + 
                        "','" + additional.ToString() + "')");   
                }
                else
                {
                    command.CommandText = (@"update [goods] set Sizes='" + size.ToString() + "', Manufacturers='" + manufacturer.ToString() + "', Compositions='" + composition.ToString()
                    + "', Additionals='" + additional.ToString() + "' where id=" + getIdByValue(name, connection));
                }
                Console.WriteLine(command.CommandText);
                command.ExecuteNonQuery();
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex);
            }
            finally
            {
                connection.Close();
            }
        }

        public DataTable fillDataTable()
        {

            //using (SqlConnection sqlConn = new SqlConnection(conSTR))
            //using (SqlCommand cmd = new SqlCommand(query, sqlConn))
            //{
            //    sqlConn.Open();
            //    DataTable dt = new DataTable();
            //    dt.Load(cmd.ExecuteReader());
            //    return dt;
            //}

            try
            {
                connection.Open();
                OleDbCommand command = new OleDbCommand();
                command.Connection = connection;
                string query = (@"SELECT * FROM [goods]");
                command.CommandText = query;
                DataTable dataTable = new DataTable();
                dataTable.Load(command.ExecuteReader());

                return dataTable;
                //command.ExecuteNonQuery();
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex);
            }
            finally
            {
                connection.Close();
            }

            return null;
        }
    }
}

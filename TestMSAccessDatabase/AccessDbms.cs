using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TestMSAccessDatabase
{
    class AccessDbms
    {
        public string ConnectionString { get; set; }

        public AccessDbms()
        {

        }

        public AccessDbms(string connectionString)
        {
            this.ConnectionString = connectionString;
        }

        public bool IsConnectSuccess()
        { 
            // test connect 
            OleDbConnection connection = new OleDbConnection(this.ConnectionString);
            try
            {
                connection.Open();
                connection.Close();

                return true;
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                if (connection.State == ConnectionState.Open)
                {
                    connection.Close();
                }
            } 
        }

        public bool Execute(string commandString)
        {
            OleDbConnection connection = new OleDbConnection(this.ConnectionString);
            OleDbCommand command = new OleDbCommand(commandString, connection);

            try
            { 
                int rowAffected = 0;

                connection.Open();

                rowAffected = command.ExecuteNonQuery();

                connection.Close();

                return (rowAffected > 0);

            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                if (connection.State == ConnectionState.Open)
                {
                    connection.Close();
                }
            } 
        }

        public bool Execute(string commandString, List<OleDbParameter> parameters)
        { 
            OleDbConnection connection = new OleDbConnection(this.ConnectionString);
            OleDbCommand command = new OleDbCommand(commandString, connection);

            try
            { 
                int rowAffected = 0;

                if (parameters != null)
                {
                    foreach (OleDbParameter item in parameters)
                    {
                        command.Parameters.Add(item);
                    }
                }
                 
                connection.Open();

                rowAffected = command.ExecuteNonQuery();

                connection.Close();

                return (rowAffected > 0);

            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                if (connection.State == ConnectionState.Open)
                {
                    connection.Close();
                }
            } 
        }

        public DataSet GetDataSet(string commandString)
        { 
            OleDbDataAdapter adapter = new OleDbDataAdapter(commandString, this.ConnectionString);
            DataSet ds = new DataSet();

            try
            {
                adapter.Fill(ds);

                return ds;
            }
            catch (Exception ex)
            {
                throw ex;
            } 
        }

        public DataSet GetDataSet(string commandString, List<OleDbParameter> parameters)
        { 
            OleDbDataAdapter adapter = new OleDbDataAdapter(commandString, this.ConnectionString);
            DataSet ds = new DataSet();

            try
            { 
                if (parameters != null)
                {
                    foreach (OleDbParameter item in parameters)
                    {
                        adapter.SelectCommand.Parameters.Add(item);
                    }
                }

                adapter.Fill(ds);

                return ds;
            }
            catch (Exception ex)
            {
                throw ex;
            } 
        }

        public DataTable GetDataTable(string commandString)
        { 
            OleDbDataAdapter adapter = new OleDbDataAdapter(commandString, this.ConnectionString);
            DataTable dt = new DataTable();

            try
            {
                adapter.Fill(dt);

                return dt;
            }
            catch (Exception ex)
            {
                throw ex;
            } 
        }

        public DataTable GetDataTable(string commandString, List<OleDbParameter> parameters)
        {

            OleDbDataAdapter adapter = new OleDbDataAdapter(commandString, this.ConnectionString);
            DataTable dt = new DataTable();

            try
            {

                if (parameters != null)
                {
                    foreach (OleDbParameter item in parameters)
                    {
                        adapter.SelectCommand.Parameters.Add(item);
                    }
                }

                adapter.Fill(dt);

                return dt;
            }
            catch (Exception ex)
            {
                throw ex;
            }

        }
    }
}

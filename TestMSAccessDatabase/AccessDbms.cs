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

        // Normal
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

        // Function
        public bool SetInsert(string table, List<Parameter> parameters)
        {
            try
            {
                StringBuilder setColumn = new StringBuilder();
                StringBuilder setParameter = new StringBuilder();

                List<OleDbParameter> OleParameters = new List<OleDbParameter>();
                foreach (var prmt in parameters)
                {
                    int i = parameters.IndexOf(prmt);

                    setColumn.Append(prmt.name);
                    setParameter.Append($"@{prmt.name}");

                    if (i != parameters.Count - 1)
                    {
                        setColumn.Append(", ");
                        setParameter.Append(", ");
                    }

                    OleParameters.Add(new OleDbParameter(prmt.name, prmt.value));
                }

                string query = "INSERT INTO " + table + " (" + setColumn + ") VALUES (" + setParameter + ")";

                return Execute(query, OleParameters);
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public DataSet SetInsertDS(string table, List<Parameter> parameters)
        {
            try
            {
                StringBuilder columns = new StringBuilder();
                StringBuilder values = new StringBuilder();

                List<OleDbParameter> OleParameters = new List<OleDbParameter>();
                foreach (var prmt in parameters)
                {
                    int i = parameters.IndexOf(prmt);

                    columns.Append(prmt.name);
                    values.Append($"@{prmt.name}");

                    if (i != parameters.Count - 1)
                    {
                        columns.Append(", ");
                        values.Append(", ");
                    }

                    OleParameters.Add(new OleDbParameter(prmt.name, prmt.value));
                }

                string query = "INSERT INTO " + table + " (" + columns + ") VALUES (" + values + ")";

                return GetDataSet(query, OleParameters);
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public DataTable SetInsertDT(string table, List<Parameter> parameters)
        {
            try
            {
                StringBuilder columns = new StringBuilder();
                StringBuilder values = new StringBuilder();

                List<OleDbParameter> OleParameters = new List<OleDbParameter>();
                foreach (var prmt in parameters)
                {
                    int i = parameters.IndexOf(prmt);

                    columns.Append(prmt.name);
                    values.Append($"@{prmt.name}");

                    if (i != parameters.Count - 1)
                    {
                        columns.Append(", ");
                        values.Append(", ");
                    }

                    OleParameters.Add(new OleDbParameter(prmt.name, prmt.value));
                }

                string query = "INSERT INTO " + table + " (" + columns + ") VALUES (" + values + ")";

                return GetDataTable(query, OleParameters);
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public bool SetUpdate(string table, List<Parameter> parameters, Parameter where)
        {
            try
            {
                StringBuilder setParameter = new StringBuilder();

                List<OleDbParameter> OleParameters = new List<OleDbParameter>();
                foreach (var prmt in parameters)
                {
                    int i = parameters.IndexOf(prmt);

                    setParameter.Append($"{prmt.name} = @{prmt.name}");

                    if (i != parameters.Count - 1)
                    {
                        setParameter.Append(", ");
                    }

                    OleParameters.Add(new OleDbParameter(prmt.name, prmt.value));
                }

                string setWhere = $"{where.name} = @{where.name}";

                OleParameters.Add(new OleDbParameter(where.name, where.value));

                string query = "UPDATE " + table + " SET " + setParameter + " WHERE " + setWhere;

                return Execute(query, OleParameters);
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public DataSet SetUpdateDS(string table, List<Parameter> parameters, Parameter where)
        {
            try
            {
                StringBuilder setParameter = new StringBuilder();

                List<OleDbParameter> OleParameters = new List<OleDbParameter>();
                foreach (var prmt in parameters)
                {
                    int i = parameters.IndexOf(prmt);

                    setParameter.Append($"{prmt.name} = @{prmt.name}");

                    if (i != parameters.Count - 1)
                    {
                        setParameter.Append(", ");
                    }

                    OleParameters.Add(new OleDbParameter(prmt.name, prmt.value));
                }

                string setWhere = $"{where.name} = @{where.value}";

                OleParameters.Add(new OleDbParameter(where.name, where.value));

                string query = "UPDATE " + table + " SET " + setParameter + " WHERE " + setWhere;

                return GetDataSet(query, OleParameters);
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public DataTable SetUpdateDT(string table, List<Parameter> parameters, Parameter where)
        {
            try
            {
                StringBuilder setParameter = new StringBuilder();

                List<OleDbParameter> OleParameters = new List<OleDbParameter>();
                foreach (var prmt in parameters)
                {
                    int i = parameters.IndexOf(prmt);

                    setParameter.Append($"{prmt.name} = @{prmt.name}");

                    if (i != parameters.Count - 1)
                    {
                        setParameter.Append(", ");
                    }

                    OleParameters.Add(new OleDbParameter(prmt.name, prmt.value));
                }

                string setWhere = $"{where.name} = @{where.value}";

                OleParameters.Add(new OleDbParameter(where.name, where.value));

                string query = "UPDATE " + table + " SET " + setParameter + " WHERE " + setWhere;

                return GetDataTable(query, OleParameters);
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
    }
}
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using System.Data.OleDb;
using System.Configuration;


namespace UploadBio
{
    class Cls_DAL
    {
        OleDbConnection objConnection;
        public Cls_DAL()
        {
            //
            // TODO: Add constructor logic here
            //
        }

        public OleDbConnection getAccessConn()
        {
            string databaseName = ConfigurationSettings.AppSettings["database"];
            //string pathDB = "att2000.mdb";
            string strConnection = "Provider=Microsoft.Jet.OleDb.4.0;";
            //strConnection += @"Data Source=" + pathDB;
            strConnection += @"Data Source=" + ConfigurationSettings.AppSettings["database"] + ";";

            objConnection = new OleDbConnection(strConnection);
            //objConnection.Open();
            return objConnection;

        }

        public void CloseConnection()
        {
            if ((!((objConnection) == null)))
            {
                if ((objConnection.State == ConnectionState.Open))
                {
                    objConnection.Close();
                }
                objConnection.Dispose();
            }
        }

        public DataSet getDs(string sqlStr)
        {
            OleDbCommand cmd = new OleDbCommand();
            cmd.Connection = getAccessConn();

            cmd.CommandText = sqlStr;
            cmd.CommandType = CommandType.Text;

            DataSet ds = new DataSet();
            OleDbDataAdapter da = new OleDbDataAdapter(cmd);
            da.Fill(ds);

            CloseConnection();

            if (ds != null)
                return ds;
            else
                return null;
        }

        public DataTable getDt(string sqlStr)
        {
            OleDbCommand cmd = new OleDbCommand();
            cmd.Connection = getAccessConn();

            cmd.CommandText = sqlStr;
            cmd.CommandType = CommandType.Text;

            DataSet ds = new DataSet();
            OleDbDataAdapter da = new OleDbDataAdapter(cmd);
            da.Fill(ds);

            CloseConnection();

            if (ds != null)
                return ds.Tables[0];
            else
                return null;
        }

        public int execNonQuery(string sqlStr)
        {
            OleDbCommand cmd = new OleDbCommand();
            cmd.Connection = getAccessConn();
            cmd.Connection.Open();

            cmd.CommandText = sqlStr;
            cmd.CommandType = CommandType.Text;

            int rtnVal = cmd.ExecuteNonQuery();

            return rtnVal;
        }
    }
}

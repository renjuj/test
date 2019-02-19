using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using System.Data.OleDb;

namespace UploadBio
{
    class Cls_Upload
    {
        Cls_DAL oaDAL = new Cls_DAL();

        //Declarations
        int empID;

        //Properties
        public int EMPID
        { get { return empID; } set { empID = value; } }

        public DataTable getEmpLogs()
        {
            DataSet ds = new DataSet();
            string strSQL = "";
            strSQL += "SELECT TOP 10 UserID as ID,CheckTime as CTime,CheckType as CType FROM CHECKINOUT";

            //strSQL += "SELECT BadgeNumber,CheckTime,CheckType FROM USERINFO LEFT OUTER JOIN";
            ////strSQL += " (SELECT * FROM CHECKINOUT WHERE DateValue(CheckTime)>=DateValue('" + lastLog + "'))";
            //strSQL += " (SELECT * FROM CHECKINOUT WHERE DateValue(CheckTime)<DateValue('01/08/2011'))";
            //strSQL += " CHECKINOUT ON USERINFO.USERID = CHECKINOUT.USERID";
            //strSQL += " WHERE CheckTime IS NOT NULL";

            ds = oaDAL.getDs(strSQL);
            if (ds.Tables[0].Rows.Count > 0)
                return ds.Tables[0];
            else
                return null;
        }

        public DateTime? getLastLogTime()
        {
            DataSet ds = new DataSet();
            string strSQL = "";
            strSQL += "SELECT MAX(CheckTime) as MaxCTime FROM CHECKINOUT";

            //strSQL += "SELECT BadgeNumber,CheckTime,CheckType FROM USERINFO LEFT OUTER JOIN";
            ////strSQL += " (SELECT * FROM CHECKINOUT WHERE DateValue(CheckTime)>=DateValue('" + lastLog + "'))";
            //strSQL += " (SELECT * FROM CHECKINOUT WHERE DateValue(CheckTime)<DateValue('01/08/2011'))";
            //strSQL += " CHECKINOUT ON USERINFO.USERID = CHECKINOUT.USERID";
            //strSQL += " WHERE CheckTime IS NOT NULL";

            ds = oaDAL.getDs(strSQL);
            if (ds.Tables[0].Rows.Count > 0)
                return DateTime.Parse(ds.Tables[0].Rows[0][0].ToString());
            else
                return null;
        }
    }
}

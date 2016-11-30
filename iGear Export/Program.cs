﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Diagnostics;
using System.Data;
using System.Data.SqlClient;
using System.IO;

namespace iGear_Export
{
    class Program
    {

        static System.Data.SqlClient.SqlConnection sqlConnection1 = new System.Data.SqlClient.SqlConnection(Properties.Settings.Default.sqlConnection);
        static string sSource = "iGear";
        static string sLog = "Dana";


        static void Main(string[] args)
        {

            //make the event log
            //ensure event log is there and working...
            if (!EventLog.SourceExists(sSource))
                EventLog.CreateEventSource(sSource, sLog);

            //in order - first set the uninteresting to 2
            Boolean actTest = exTract();

            if (actTest == true)
            {
                //ok first part passed - do the next
                EventLog.WriteEntry(sSource, "Done", EventLogEntryType.Information);
            }
            else
            {
                //do nothing the first part failed.
                EventLog.WriteEntry(sSource, "Not done", EventLogEntryType.Warning);
            }

        }

        static Boolean exTract()
        {
            //take the date and time - and do it for the previous hour
            DateTime dteStart = DateTime.Now.AddHours(Properties.Settings.Default.intHours*-1);
            DateTime dteStop = DateTime.Now;

            //read through each record - if it is not transID=1 then update processed to 2
            //open database read to write to
            sqlConnection1.Open();

            try
            {

                   string CommandText = " select a_ProductionDate, b.supplierpart MAT_NUM, '' OLD_MAT_REF, '3602' PLANT, 'SOEM' SLOC,'A' BATCH_NUM, " +
                        "a.serialnumber HU_NUM, c.PartSerialNumber SERIAL, c.scantimestamp , d.Number 'WO #' " +
                        "from pallet a " +
                        "inner join PartDefinition b on a.PartDefinition_ID = b.ID " +
                        "inner join PalletDetail c on a.id = c.Pallet_ID " +
                        "inner join WorkOrder d on a.WorkOrderHistory_ID = d.ID " +
                        "where Status_ID = 1 " +
                        "and c.scantimestamp > '" + dteStart.ToString("yyyy-MM-dd HH:mm:ss") + "' and c.scantimestamp <= '" + dteStop.ToString("yyyy-MM-dd HH:mm:ss") + "' " +
                        "order by c.ScanTimestamp";

                //run the command
                SqlDataReader reader;
                reader = new SqlCommand(CommandText, sqlConnection1).ExecuteReader();

                if(reader.HasRows)
                {
                    string streamFile = dteStart.ToString("yyyyMMddHHmmss");
                    using (StreamWriter sw = new StreamWriter(Properties.Settings.Default.streamPath + streamFile + ".txt"))
                    {
                        //write header
                        sw.WriteLine("Production Date\tMAT_NUM\tOLD_MAT_REF\tPLANT\tSLOC\tBATCH_NUM\tHU_NUM\tSERIAL\tscantimestamp\tWO #");
                        while (reader.Read())
                        {
                            //write to data file
                            //streamwrite the output to a timestamped file
                            sw.WriteLine(reader["a_ProductionDate"] + "\t" + reader["MAT_NUM"] + "\t" + reader["OLD_MAT_REF"] + "\t" + reader["PLANT"] + "\t" + reader["SLOC"] + "\t" +
                            reader["BATCH_NUM"] + "\t" + reader["HU_NUM"] + "\t" + reader["SERIAL"] + "\t" + reader["scantimestamp"] + "\t" + reader["WO #"]);
                        }
                    }

                }
                //if I get o here i have been succesfull - return true
                sqlConnection1.Close();
                return true;
            }
            catch (Exception e)
            {
                //write error log, then return false
                EventLog.WriteEntry(sSource, e.Message, EventLogEntryType.Error, 234);
                sqlConnection1.Close();
                return false;
            }

        }

    }
}

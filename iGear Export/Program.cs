using System;
using System.Diagnostics;
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
            Boolean actTest = ExTract();

            if (actTest == true)
            {
                //ok extract weas succesfull
                EventLog.WriteEntry(sSource, "Done", EventLogEntryType.Information, 101);
            }

        }

        static Boolean ExTract()
        {
            //take the date and time - and do it for the previous hour
            //DateTime dteStart = DateTime.Now.AddHours(Properties.Settings.Default.intHours*-1);
            //set strat time to this time yesterday. R1266554 Ver1.0.0.1
            DateTime dteStart = DateTime.Now.AddDays(-1);
            DateTime dteStop = DateTime.Now;
            string filename = "Day";

            //read through each record - if it is not transID=1 then update processed to 2
            //open database read to write to
            sqlConnection1.Open();

            try
            {
                //R1276932 - MSW - 11th Jan 2017
                //workorderhistory changed from workorder
                   string CommandText = " select a_ProductionDate, b.supplierpart MAT_NUM, '' OLD_MAT_REF, '3602' PLANT, 'SOEM' SLOC,'A' BATCH_NUM, " +
                        "a.serialnumber HU_NUM, c.PartSerialNumber SERIAL, c.scantimestamp , d.Number 'WO #' " +
                        "from pallet a " +
                        "inner join PartDefinition b on a.PartDefinition_ID = b.ID " +
                        "inner join PalletDetail c on a.id = c.Pallet_ID " +
                        "inner join WorkOrderHistory d on a.WorkOrderHistory_ID = d.ID " +
                        "where Status_ID = 1 " +
                        "and c.scantimestamp > '" + dteStart.ToString("yyyy-MM-dd HH:mm:ss") + "' and c.scantimestamp <= '" + dteStop.ToString("yyyy-MM-dd HH:mm:ss") + "' " +
                        "order by c.ScanTimestamp";

                //run the command
                SqlDataReader reader;
                reader = new SqlCommand(CommandText, sqlConnection1).ExecuteReader();

                if(reader.HasRows)
                {
                    string streamFile = dteStart.ToString("yyyyMMdd");
                    using (StreamWriter sw = new StreamWriter(Properties.Settings.Default.streamPath + filename + streamFile + ".txt"))
                    {
                        //write header
                        sw.WriteLine("Production Date\tMAT_NUM\tOLD_MAT_REF\tPLANT\tSLOC\tBATCH_NUM\tHU_NUM\tSERIAL\tscantimestamp\tWO #");
                        while (reader.Read())
                        {
                            //write to data file
                            //streamwrite the output to a timestamped file
                            sw.WriteLine(reader["a_ProductionDate"] + "\t" + reader["MAT_NUM"] + "\t" + reader["OLD_MAT_REF"] + "\t" + reader["PLANT"] + "\t" + reader["SLOC"] + "\t" +
                            reader["BATCH_NUM"] + "\t" + reader["HU_NUM"] + "\t" + reader["SERIAL"].ToString() + "\t" + reader["scantimestamp"] + "\t" + reader["WO #"]);
                            //changed serial field to be string - removes trailing .zeros
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
                EventLog.WriteEntry(sSource, e.Message, EventLogEntryType.Error, 201);
                sqlConnection1.Close();
                return false;
            }

        }

    }
}

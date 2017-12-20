using System;
using System.Diagnostics;
using System.Data.SqlClient;
using System.IO;
using System.IO.Compression;

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
                //R13591319 - MSW - 24th July 2017
                //added link to asset table, include asset description in select ---  \\traukbirmfs01\all\iGear\
                string CommandText = "SELECT a_ProductionDate, b.supplierpart MAT_NUM, '' OLD_MAT_REF, '3602' PLANT, 'SOEM' SLOC,'A' BATCH_NUM, " +
                     "a.serialnumber HU_NUM, c.PartSerialNumber SERIAL, c.scantimestamp, d.Number 'WO #', e.Description 'WRKCTR' " +
                     "FROM Pallet a " +
                     "INNER JOIN PartDefinition b ON a.PartDefinition_ID = b.ID " +
                     "INNER JOIN PalletDetail c ON a.id = c.Pallet_ID " +
                     "INNER JOIN WorkOrderHistory d ON a.WorkOrderHistory_ID = d.ID " +
                     "INNER JOIN Asset e ON a.Asset_ID = e.ID " +
                     "WHERE Status_ID = 1 " +
                     "AND c.scantimestamp > '" + dteStart.ToString("yyyy-MM-dd HH:mm:ss") + "' AND c.scantimestamp <= '" + dteStop.ToString("yyyy-MM-dd HH:mm:ss") + "' " +
                     "ORDER BY c.ScanTimestamp";

                //run the command
                SqlDataReader reader;
                reader = new SqlCommand(CommandText, sqlConnection1).ExecuteReader();

                if (reader.HasRows)
                {
                    string streamFile = dteStart.ToString("yyyyMMddHHmmss");

                    using (StreamWriter sw = new StreamWriter(Properties.Settings.Default.streamPath + filename + streamFile + ".txt"))
                    {
                        //write header
                        sw.WriteLine("Production Date\tMAT_NUM\tOLD_MAT_REF\tPLANT\tSLOC\tBATCH_NUM\tHU_NUM\tSERIAL\tscantimestamp\tWO #\tWork Centre");
                        while (reader.Read())
                        {
                            //write to data file
                            //streamwrite the output to a timestamped file
                            //R1359319 - MSW - 24th July 2017
                            //include asset description in output file
                            sw.WriteLine(reader["a_ProductionDate"] + "\t" + reader["MAT_NUM"] + "\t" + reader["OLD_MAT_REF"] + "\t" + reader["PLANT"] + "\t" + reader["SLOC"] + "\t" +
                            reader["BATCH_NUM"] + "\t" + reader["HU_NUM"] + "\t" + reader["SERIAL"].ToString() + "\t" + reader["scantimestamp"] + "\t" + reader["WO #"] + "\t" + reader["WRKCTR"]);
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

        public static void Compress(FileInfo fi)
        {
            // Get the stream of the source file.
            using (FileStream inFile = fi.OpenRead())
            {
                // Prevent compressing hidden and 
                // already compressed files.
                if ((File.GetAttributes(fi.FullName)
                    & FileAttributes.Hidden)
                    != FileAttributes.Hidden & fi.Extension != ".gz")
                {
                    // Create the compressed file.
                    using (FileStream outFile =
                                File.Create(fi.FullName + ".gz"))
                    {
                        using (GZipStream Compress =
                            new GZipStream(outFile,
                            CompressionMode.Compress))
                        {
                            // Copy the source file into 
                            // the compression stream.
                            inFile.CopyTo(Compress);

                            Console.WriteLine("Compressed {0} from {1} to {2} bytes.",
                                fi.Name, fi.Length.ToString(), outFile.Length.ToString());
                        }
                    }
                }
            }
        }



    }
}

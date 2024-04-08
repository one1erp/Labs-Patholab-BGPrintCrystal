using System.Configuration;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data.SqlClient;
using System.IO;
using System.Diagnostics;
using System.Reflection;
using System.Linq;
using System.Windows.Forms;


using System.Runtime.InteropServices;

using CrystalDecisions.CrystalReports.Engine;
using CrystalDecisions.Shared;

using System.Diagnostics;
using Oracle.ManagedDataAccess.Client;
using System.Threading;
using Timer = System.Timers.Timer;


namespace BGPrintCrystal
{
    public partial class Form1 : Form
    {

        //   SdgInfo _sdgInfo = null;
        OracleConnection oraCon = null;
        OracleCommand cmd = null;
        private ReportDocument cr;
        private ConnectionInfo crConnectionInfo;

        private string SdgID, wnName, COMPUTER_NAME;

        //private string configPath =
        //    @"\\vm-nautilus\nautilus-share\Extensions\BGPrintCrystal\BGPrintCrystal.exe.config";
        OracleCommand command;
        string logPath, server, user, password, Interval;

        public Form1()
        {

            InitializeComponent();
            //if ( Environment.MachineName == "ONE1PC1518" )
            //{
            //    configPath = @"C:\Work\Patholab Proj\BGPrintCrystal\BGPrintCrystal\App.config";
            //}
            Settings();
            WriteLogFile("Start program");
            String thisprocessname = Process.GetCurrentProcess().ProcessName;

            if (Process.GetProcesses().Count(p => p.ProcessName == thisprocessname) > 1)
            {
                WriteLogFile("לא ניתן להפעיל את תוכנית ההדפסות יותר מפעם אחת!");
                WriteLogFile("END PROGRAM");
                this.Close();


            }
            listBox1.Items.Add("....... ");

            var connection = GetConnection();
            command = new OracleCommand();
            command.Connection = connection;

            cr = new CrystalDecisions.CrystalReports.Engine.ReportDocument();
            crConnectionInfo = new ConnectionInfo();



            GO();
            timer1.Interval = int.Parse(Interval);
            timer1.Tick += (e, r) => { GO(); };
            timer1.Start();

        }
        bool IsProxy;
        private void Settings()
        {
            var xx = ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None);
            var appSettings = xx.AppSettings;
            //var dd=bb.Settings["Log"].Value;
            //ExeConfigurationFileMap map = new ExeConfigurationFileMap ( );
            //map.ExeConfigFilename = System.Reflection.Assembly.GetEntryAssembly().Location;
            //var assemblyPath = System.Reflection.Assembly.GetEntryAssembly().Location+".config";

            // Configuration cfg = ConfigurationManager.OpenMappedExeConfiguration ( map, ConfigurationUserLevel.None );



            //    var appSettings = cfg.AppSettings;

             IsProxy = appSettings.Settings["IsProxy"].Value == "True";
            logPath = appSettings.Settings["Log"].Value;
            server = appSettings.Settings["Server"].Value;
            user = appSettings.Settings["User"].Value;
            password = appSettings.Settings["Password"].Value;
            Interval = appSettings.Settings["Interval"].Value;

        }

        private void GO()
        {
            try
            {
                timer1.Stop();

                string sql = string.Format(
                    "select * from lims.BG_CRYSTAL_PRINT where WORKSTATION_NAME='{0}' and is_printed='0'",
                    Environment.MachineName);
                command.CommandText = sql;
                var reader = command.ExecuteReader();
                if (!reader.HasRows)
                {
                    listBox1.Items.Add("No Sdg to print.");

                }

                while (reader.Read())
                {
                    try
                    {


                        var id = reader["ID"].ToString();
                        var RECORDID = reader["SDG_ID"].ToString();
                        var rptPath = reader["PATH"].ToString();
                        var param = reader["PARAM_NAME"].ToString();
                        listBox1.Items.Add(RECORDID + " Is Printing");
                        WriteLogFile(RECORDID + " Is Printing.");
                        WriteLogFile("starting to generate letter for " + RECORDID);



                        cr.Load(rptPath);
                        crConnectionInfo.ServerName = server; // _ntlsCon.GetServerName ( );
                        crConnectionInfo.UserID = user; // _ntlsCon.GetUsername ( );
                        crConnectionInfo.Password = password; // _ntlsCon.GetPassword ( );
                        cr.SetParameterValue(param, RECORDID.ToString());
                        var CrTables = cr.Database.Tables;
                        foreach (CrystalDecisions.CrystalReports.Engine.Table crTable in CrTables)
                        {
                            TableLogOnInfo crTableLoginInfo = crTable.LogOnInfo;
                            crTableLoginInfo.ConnectionInfo = crConnectionInfo;
                            crTable.ApplyLogOnInfo(crTableLoginInfo);
                        }

                        WriteLogFile("Sending to printer for " + RECORDID);

                        cr.PrintToPrinter(1, true, 0, 0);
                        command.CommandText = "Update lims.BG_CRYSTAL_PRINT set is_printed='1' where ID ='" + id + "'";
                        WriteLogFile(command.CommandText);

                        command.ExecuteNonQuery();
                        listBox1.Items.Add(RECORDID + " printed successfully");
                        WriteLogFile(RECORDID + " printed successfully");


                    }

                    catch (Exception e)
                    {
                        label1.Text = e.Message;
                        listBox1.Items.Add("Error on printing " + e.Message);
                        if (e.InnerException != null)
                            WriteLogFile("Error on printing " + e.Message + e.InnerException.ToString());
                        if (e.Message == "Load report failed.")
                        {
                            cr.Close();
                            cr.Dispose();
                            cr = null;
                            cr = new CrystalDecisions.CrystalReports.Engine.ReportDocument();
                            crConnectionInfo = new ConnectionInfo();
                        }

                    }


                }
            }
            catch (Exception ex)
            {
                label1.Text = ex.Message;
                listBox1.Items.Add(ex.Message);
                WriteLogFile(ex.Message);
            }
            finally
            {
                timer1.Start();

            }
        }

        void WriteLogFile(string txt)
        {
            try
            {

                string dir = Path.Combine(logPath, Environment.MachineName);

                string fileName = "BGPrintCrystal" + "-" + DateTime.Now.ToString("dd-MM-yyyy") + ".txt";
                if (!Directory.Exists(dir))
                {
                    Directory.CreateDirectory(dir);
                }
                string fullPath = Path.Combine(dir, fileName);

                using (FileStream file = new FileStream(fullPath, FileMode.Append, FileAccess.Write))
                {
                    var streamWriter = new StreamWriter(file);
                    streamWriter.WriteLine(DateTime.Now.ToString("MM/dd/yyyy hh:mm:ss.fff"));

                    streamWriter.WriteLine(txt);
                    streamWriter.WriteLine();
                    streamWriter.Close();
                }
            }
            catch
            {
            }


        }

        public OracleConnection GetConnection()
        {
            OracleConnection connection = null;

            //initialize variables
            string rolecommand;
            //try catch block
            try
            {

                string connectionString;


                connectionString =
                    string.Format("Data Source={0};User ID={1};Password={2};", server, user, password);
                //create connection
                connection = new OracleConnection(connectionString);

                //open the connection
                connection.Open();


            }
            catch (Exception e)
            {
                //throw the exeption
                MessageBox.Show("Err At GetConnection: " + e.Message);
            }

            return connection;
        }


    }
}

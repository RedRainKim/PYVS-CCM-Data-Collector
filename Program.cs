using System;
using System.Threading;
using System.Data;
using System.Data.SqlClient;
using NLog;

namespace PYVS_CCMDataCollector
{
    partial class Program
    {
        //define members
        private static readonly int _delay = 600000;    //main delay time (milisecond)

        private static string _conHIS = "Data Source = 172.19.71.211; Initial Catalog = L2HIS; Persist Security Info=True;User ID = MS_L2USER; Password=MS_L2USER;MultipleActiveResultSets=True;Pooling=true;Max Pool Size=10;Connection Lifetime = 120; Connect Timeout = 60";
        private static string _conEXT = "Data Source = 172.19.71.211; Initial Catalog = L2EXCH; Persist Security Info=True;User ID = MS_L2USER; Password=MS_L2USER;MultipleActiveResultSets=True;Pooling=true;Max Pool Size=10;Connection Lifetime = 120; Connect Timeout = 60";
        private static bool dbStatus = false; //database status
        private static int lastheat = 0;

        //logging object - NLog package library
        private static Logger _log = LogManager.GetCurrentClassLogger();

        static void Main(string[] args)
        {
            //set console size(width,height)
            Console.SetWindowSize(80, 30);

            _log.Info("==================================================");
            _log.Info("PYVS_CCMDataCollector Process start...");
            _log.Info("Location : {0}", Environment.CurrentDirectory);
            _log.Info("==================================================");

            try
            {
                CCMData ccmData = new CCMData();

                while (true)
                {
                    //checking database connection
                    if (dbStatus == false) InitializeDataBase();

                    //data initialize
                    ccmData.InitData();

                    //check last heat information
                    int reportHeatno = GetLastHeat();

                    //for test - old heat "SELECT TOP 100 REPORT_COUNTER FROM REPORTS WHERE AREA_ID = 1100 ORDER BY REPORT_COUNTER DESC"
                    //reportHeatno = 125649;

                    //check report number
                    if (reportHeatno == 0)
                    {
                        _log.Warn("Fail to get heat report : {0}", reportHeatno);
                    }

                    //compare last sent heat number
                    if (reportHeatno > 0 && lastheat != reportHeatno)
                    {
                        _log.Info("Found new heat report : {0}", reportHeatno);

                        //Collect data
                        if (GetCCMData(reportHeatno, ref ccmData))
                        {
                            //data send to MES
                            if (SendMessage(ref ccmData))
                            {
                                lastheat = reportHeatno;
                                _log.Info("Data send complete... {0}/{1}", reportHeatno, ccmData.heat.ToString());
                                _log.Info("==================================================");
                            }
                        }
                    }
                    Console.Write(".");
                    Thread.Sleep(_delay);
                }

            }
            catch (Exception e)
            {
                _log.Error(e.Message);
            }


        }


        private static void InitializeDataBase()
        {
            try
            {
                _log.Warn("Database initialize.....");

                // Connection check...
                using (SqlConnection conn = new SqlConnection(_conHIS))
                {
                    conn.Open();
                    dbStatus = true;
                }
                using (SqlConnection conn = new SqlConnection(_conEXT))
                {
                    conn.Open();
                }
                _log.Info("Database initialize success....");
            }
            catch (Exception ex)
            {
                dbStatus = false;
                _log.Error(ex.Message);
            }
        }

        private static int GetLastHeat()
        {
            int reportno = 0;

            try
            {
                #region SQL
                //Get last report heat number
                string sql = "SELECT REPORT_COUNTER FROM REPORTS WHERE REPORT_COUNTER = (SELECT TOP 1 REPORT_COUNTER FROM REPORTS WHERE AREA_ID = 1100 ORDER BY REPORT_COUNTER DESC)";
                #endregion

                //
                using (SqlConnection con = new SqlConnection(_conHIS))
                using (SqlCommand cmd = new SqlCommand(sql, con))
                {
                    cmd.CommandType = CommandType.Text;
                    con.Open();

                    using (SqlDataReader reader = cmd.ExecuteReader())
                    {
                        if (reader.HasRows)
                        {
                            reader.Read();
                            reportno = Convert.ToInt32(reader["REPORT_COUNTER"]);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                dbStatus = false;
                _log.Error(ex.Message);
            }

            return reportno;
        }

        private static bool GetCCMData(int reportno, ref CCMData cdata)
        {
            bool flag = false;
            try
            {

                #region SQL
                //Get last report heat number
                string sql = "SELECT HEAT_ID, PO_ID, " +
                    "(SELECT COUNT(HEAT_ID) FROM REPORTS WHERE AREA_ID = 1100 AND HEAT_ID = (SELECT HEAT_ID FROM REPORTS WHERE REPORT_COUNTER = " + reportno + ")) AS CNT, " +
                    "ISNULL((SELECT AVG_VALUE FROM REP_CCM_STRAND_VARS WHERE REPORT_COUNTER = " + reportno + " AND STRAND_NO = 1 AND VARIABLE_ID = 68),0) AS MW_A_STD1, " +
                    "ISNULL((SELECT AVG_VALUE FROM REP_CCM_STRAND_VARS WHERE REPORT_COUNTER = " + reportno + " AND STRAND_NO = 2 AND VARIABLE_ID = 68),0) AS MW_A_STD2, " +
                    "ISNULL((SELECT AVG_VALUE FROM REP_CCM_STRAND_VARS WHERE REPORT_COUNTER = " + reportno + " AND STRAND_NO = 3 AND VARIABLE_ID = 68),0) AS MW_A_STD3, " +
                    "ISNULL((SELECT AVG_VALUE FROM REP_CCM_STRAND_VARS WHERE REPORT_COUNTER = " + reportno + " AND STRAND_NO = 4 AND VARIABLE_ID = 68),0) AS MW_A_STD4, " +
                    "ISNULL((SELECT AVG_VALUE FROM REP_CCM_STRAND_VARS WHERE REPORT_COUNTER = " + reportno + " AND STRAND_NO = 5 AND VARIABLE_ID = 68),0) AS MW_A_STD5, " +
                    "ISNULL((SELECT AVG_VALUE FROM REP_CCM_STRAND_VARS WHERE REPORT_COUNTER = " + reportno + " AND STRAND_NO = 6 AND VARIABLE_ID = 68),0) AS MW_A_STD6, " +
                    "ISNULL((SELECT AVG_VALUE FROM REP_CCM_STRAND_VARS WHERE REPORT_COUNTER = " + reportno + " AND STRAND_NO = 1 AND VARIABLE_ID = 70),0) AS MW_B_STD1, " +
                    "ISNULL((SELECT AVG_VALUE FROM REP_CCM_STRAND_VARS WHERE REPORT_COUNTER = " + reportno + " AND STRAND_NO = 2 AND VARIABLE_ID = 70),0) AS MW_B_STD2, " +
                    "ISNULL((SELECT AVG_VALUE FROM REP_CCM_STRAND_VARS WHERE REPORT_COUNTER = " + reportno + " AND STRAND_NO = 3 AND VARIABLE_ID = 70),0) AS MW_B_STD3, " +
                    "ISNULL((SELECT AVG_VALUE FROM REP_CCM_STRAND_VARS WHERE REPORT_COUNTER = " + reportno + " AND STRAND_NO = 4 AND VARIABLE_ID = 70),0) AS MW_B_STD4, " +
                    "ISNULL((SELECT AVG_VALUE FROM REP_CCM_STRAND_VARS WHERE REPORT_COUNTER = " + reportno + " AND STRAND_NO = 5 AND VARIABLE_ID = 70),0) AS MW_B_STD5, " +
                    "ISNULL((SELECT AVG_VALUE FROM REP_CCM_STRAND_VARS WHERE REPORT_COUNTER = " + reportno + " AND STRAND_NO = 6 AND VARIABLE_ID = 70),0) AS MW_B_STD6 " +
                    "FROM REPORTS WHERE REPORT_COUNTER = " + reportno;
                #endregion

                //_log.Debug(sql);

                using (SqlConnection con = new SqlConnection(_conHIS))
                using (SqlCommand cmd = new SqlCommand(sql, con))
                {
                    cmd.CommandType = CommandType.Text;
                    con.Open();

                    using (SqlDataReader reader = cmd.ExecuteReader())
                    {
                        if (reader.HasRows)
                        {
                            reader.Read();
                            cdata.heat = reader["HEAT_ID"].ToString().Trim();
                            cdata.porder = reader["PO_ID"].ToString().Trim();
                            cdata.seq = Convert.ToInt16(reader["CNT"]);

                            cdata.moldwaterInletStd1A = Convert.ToDouble(reader["MW_A_STD1"]);
                            cdata.moldwaterInletStd1B = Convert.ToDouble(reader["MW_B_STD1"]);
                            cdata.moldwaterInletStd2A = Convert.ToDouble(reader["MW_A_STD2"]);
                            cdata.moldwaterInletStd2B = Convert.ToDouble(reader["MW_B_STD2"]);
                            cdata.moldwaterInletStd3A = Convert.ToDouble(reader["MW_A_STD3"]);
                            cdata.moldwaterInletStd3B = Convert.ToDouble(reader["MW_B_STD3"]);
                            cdata.moldwaterInletStd4A = Convert.ToDouble(reader["MW_A_STD4"]);
                            cdata.moldwaterInletStd4B = Convert.ToDouble(reader["MW_B_STD4"]);
                            cdata.moldwaterInletStd5A = Convert.ToDouble(reader["MW_A_STD5"]);
                            cdata.moldwaterInletStd5B = Convert.ToDouble(reader["MW_B_STD5"]);
                            cdata.moldwaterInletStd6A = Convert.ToDouble(reader["MW_A_STD6"]);
                            cdata.moldwaterInletStd6B = Convert.ToDouble(reader["MW_B_STD6"]);

                            _log.Info("Data collecting finished ...");
                            flag = true;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                dbStatus = false;
                _log.Error(ex.Message);
            }
            return flag;
        }

        private static bool SendMessage(ref CCMData msg)
        {
            string msgHead = string.Empty;
            string msgData = string.Empty;
            string sql = string.Empty;
            bool flag = false;
            int decData = 0;

            try
            {
                // Set message - Header
                msgHead += "M20RL070";                              //TC Code
                msgHead += "2000";                                  //send factory code
                msgHead += "L06";                                   //sending process code
                msgHead += "2000";                                  //receive factory code
                msgHead += "M20";                                   //receive process code
                msgHead += DateTime.Now.ToString("yyyyMMddHHmmss"); //sending date time
                msgHead += "L3SentryCCMMAN";                        //send program ID
                msgHead += "RM20L06_01".PadRight(19);               //EAI IF ID (Queue name)
                msgHead += "000185".PadRight(31);                   //message length & spare 

                // Set message - Data
                msgData += msg.heat.ToString().PadRight(6);         //heat id
                msgData += msg.porder.ToString().PadRight(6);       //Production order ID
                msgData += msg.seq.ToString();                      //Repeating count (default 1)

                decData = (int)(msg.moldwaterInletStd1A * 100);       //mold wather inlect strand 1 A line (decimal *2)
                msgData += decData.ToString("D6");  
                decData = (int)(msg.moldwaterInletStd1B * 100);
                msgData += decData.ToString("D6");
                decData = (int)(msg.moldwaterInletStd2A * 100);
                msgData += decData.ToString("D6");
                decData = (int)(msg.moldwaterInletStd2B * 100);
                msgData += decData.ToString("D6");
                decData = (int)(msg.moldwaterInletStd3A * 100);
                msgData += decData.ToString("D6");
                decData = (int)(msg.moldwaterInletStd3B * 100);
                msgData += decData.ToString("D6");
                decData = (int)(msg.moldwaterInletStd4A * 100);
                msgData += decData.ToString("D6");
                decData = (int)(msg.moldwaterInletStd4B * 100);
                msgData += decData.ToString("D6");
                decData = (int)(msg.moldwaterInletStd5A * 100);
                msgData += decData.ToString("D6");
                decData = (int)(msg.moldwaterInletStd5B * 100);
                msgData += decData.ToString("D6");
                decData = (int)(msg.moldwaterInletStd6A * 100);
                msgData += decData.ToString("D6");
                decData = (int)(msg.moldwaterInletStd6B * 100);
                msgData += decData.ToString("D6");

                // Insert database
                sql = "INSERT INTO TT_L2_L3_NEW (HEADER, DATA, MSG_CODE, INTERFACE_ID) VALUES (" +
                       "'" + msgHead + "'," +
                       "'" + msgData + "'," +
                       "'M20RL070'," +
                       "'RM20L06_01')";

                using (SqlConnection con = new SqlConnection(_conEXT))
                {
                    SqlCommand cmd = new SqlCommand(sql, con);
                    cmd.CommandType = CommandType.Text;
                    con.Open();

                    int rowsAffected = cmd.ExecuteNonQuery();
                    flag = true;
                }
                _log.Debug(" --- MES sent ==> sql[{0}]", sql);
            }
            catch (Exception ex)
            {
                _log.Error(ex.Message);
            }
            return flag;
        }
    }
}

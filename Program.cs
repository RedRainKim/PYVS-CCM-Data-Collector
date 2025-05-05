using System;
using System.Threading;
using System.Data;
using System.Data.SqlClient;
using NLog;
using S7.Net;
using static PYVS_CCMDataCollector.Program;
using System.Xml.Linq;


namespace PYVS_CCMDataCollector
{
    partial class Program
    {
        //define members
        private static readonly int _delay = 60000;    //main delay time - 1min (milisecond)

        private static string _conHIS = "Data Source = 172.19.71.211; Initial Catalog = L2HIS; Persist Security Info=True;User ID = MS_L2USER; Password=MS_L2USER;MultipleActiveResultSets=True;Pooling=true;Max Pool Size=10;Connection Lifetime = 120; Connect Timeout = 60";
        private static string _conEXT = "Data Source = 172.19.71.211; Initial Catalog = L2EXCH; Persist Security Info=True;User ID = MS_L2USER; Password=MS_L2USER;MultipleActiveResultSets=True;Pooling=true;Max Pool Size=10;Connection Lifetime = 120; Connect Timeout = 60";
        private static bool dbStatus = false; //database status
        private static int lastheat = 0;

        //last data - cutting length
        private static string scale1LastHeat;
        private static short scale1LastSequenceNo;
        private static short scale1LastStrandNo;

        private static string scale2LastHeat;
        private static short scale2LastSequenceNo;
        private static short scale2LastStrandNo;

        //PLC
        private static Plc plc;
        private static string ipaddr_GENPLC = "172.16.1.180";    //CCM GEN PLC address   
        private static short rack = 0; //PLC Rack
        private static short slot = 3; //Slot 
        private static bool plcStatus = false;

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
                //PLC connectivity information
                plc = new Plc(CpuType.S7300, ipaddr_GENPLC, rack, slot);

                //data structure
                CCMData ccmData = new CCMData();
                TundishData tundishData = new TundishData();
                CutLengthData cutLengthData = new CutLengthData();

                //member variable init
                scale1LastHeat = "V99999";
                scale1LastStrandNo = 9;
                scale1LastSequenceNo = 99;
                scale2LastHeat = "V99999";
                scale2LastStrandNo = 9;
                scale2LastSequenceNo = 99;

                bool heatChange = false;

                //Main
                while (true)
                {
                    //checking database connection
                    if (dbStatus == false) InitializeDataBase();
                    if (plcStatus == false) InitializePLC();


                    #region Cyclic result - Tundish Data

                    bool collectdata = GetTundishData(ref tundishData);

                    _log.Info("Cyclick : heat number check...{0}", tundishData.heat.ToString());

                    if (collectdata)
                    {
                        //debug
                        _log.Info("Tundish Data : {0} / {1:0.00} / {2:0.00}", tundishData.heat.ToString(), tundishData.tundishWeight, tundishData.tundishHeight);

                        SendMessageCyc(ref tundishData);
                    }
                    #endregion

                    #region Cyclic result - Ladle open/close count data
                    //byte[] heatLadle = plc.ReadBytes(DataType.DataBlock, 205, 6, 6);   //Heat ID
                    //string ladleHeat = S7.Net.Types.String.FromByteArray(heatLadle);
                    //ladleHeat.Trim();

                    //_log.Info("Ladle open/close count data : Heat {0}", ladleHeat.ToString());


                    //double cntCST1AutoOepn = ((uint)plc.Read("DB244.DBD352")).ConvertToFloat();
                    //double cntCST1AutoClose = ((uint)plc.Read("DB244.DBD374")).ConvertToFloat();
                    //double cntCST2AutoOpen = ((uint)plc.Read("DB244.DBD396")).ConvertToFloat();
                    //double cntCST2AutoClose = ((uint)plc.Read("DB244.DBD418")).ConvertToFloat();

                    //double cntCST1ManOpen = ((uint)plc.Read("DB244.DBD440")).ConvertToFloat();
                    //double cntCST1ManClose = ((uint)plc.Read("DB244.DBD462")).ConvertToFloat();
                    //double cntCST2ManOpen = ((uint)plc.Read("DB244.DBD484")).ConvertToFloat();
                    //double cntCST2ManClose = ((uint)plc.Read("DB244.DBD506")).ConvertToFloat();

                    //_log.Info("Data {0}/{1}/{2}/{3}  {4}/{5}/{6}/{7}", (int)cntCST1AutoOepn, (int)cntCST1AutoClose, (int)cntCST2AutoOpen, (int)cntCST2AutoClose, (int)cntCST1ManOpen, (int)cntCST1ManClose, (int)cntCST2ManOpen, (int)cntCST2ManClose);
                    
                    #endregion

                    #region Heat finish result
                    //check last heat information
                    int reportHeatno = GetLastHeat();

                    //for test - old heat "SELECT TOP 100 REPORT_COUNTER FROM REPORTS WHERE AREA_ID = 1100 ORDER BY REPORT_COUNTER DESC"
                    //reportHeatno = 147139;

                    //check report number
                    if (reportHeatno == 0)
                    {
                        _log.Warn("Fail to get heat report : {0}", reportHeatno);
                    }

                    //compare last sent heat number
                    if (reportHeatno > 0 && lastheat != reportHeatno)
                    {
                        _log.Info("Found new heat report : {0}", reportHeatno);

                        //data initialize
                        ccmData.InitData();

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
                    #endregion

                    #region SP cutting length and weight data
                    /////////
                    /// No.1 Scale data
                    ////////                    
                    //data initialize
                    cutLengthData.InitData();
                    cutLengthData.scaleNo = 1;  //#1 scale

                    //Check Heat ID and Strand ID and Sequence No. for scale #1
                    byte[] heat1 = plc.ReadBytes(DataType.DataBlock, 244, 198, 6);   //Heat ID
                    cutLengthData.heat = S7.Net.Types.String.FromByteArray(heat1);
                    cutLengthData.heat.Trim();

                    cutLengthData.strandNo = ((ushort)plc.Read("DB244.DBW216")).ConvertToShort();   //strand no.
                    cutLengthData.sequenceNo = ((ushort)plc.Read("DB244.DBW214")).ConvertToShort(); //sequence no.

                    heatChange = false; //init

                    _log.Info("Weight scale 1 : heat number check...{0}", cutLengthData.heat.ToString());

                    //Check basic information
                    if (cutLengthData.heat.Substring(0,1).Equals("V")) //Check Heat ID
                    {
                        _log.Info("Scale #1 Information : {0} / {1} / {2}", cutLengthData.heat.ToString(), cutLengthData.strandNo, cutLengthData.sequenceNo);

                        if (cutLengthData.strandNo > 0 && cutLengthData.strandNo < 7) //Check strand number
                        {
                            if (cutLengthData.sequenceNo > 0 && cutLengthData.sequenceNo < 99) //Check sequence number
                            {
                                //Compare with before sent information
                                if (scale1LastHeat.ToString() != cutLengthData.heat.ToString())
                                {
                                    heatChange = true;
                                }
                                else
                                {
                                    if (scale1LastStrandNo != cutLengthData.strandNo || scale1LastSequenceNo != cutLengthData.sequenceNo)
                                    {
                                        heatChange = true;
                                    }
                                }
                            }
                        }
                    }

                    if (heatChange)
                    {
                        _log.Info("Scale #1 New SP scaled!!! Start collect data ... ");

                        //Collect data
                        cutLengthData.lengthTarget = ((uint)plc.Read("DB244.DBD406")).ConvertToInt();           //Target Length [DINT]
                        cutLengthData.lengthLastcut = ((uint)plc.Read("DB244.DBD414")).ConvertToFloat();        //Last Cut Length [REAL]
                        cutLengthData.lengthCompensation = ((ushort)plc.Read("DB244.DBW424")).ConvertToShort(); //Compensation Length [INT]
                        cutLengthData.weight = ((uint)plc.Read("DB244.DBD218")).ConvertToFloat();             //Weight [REAL]

                        _log.Debug("Scale #1 - Target Length : {0}", cutLengthData.lengthTarget);
                        _log.Debug("Scale #1 - Last Cut Length : {0}", cutLengthData.lengthLastcut);
                        _log.Debug("Scale #1 - Compensation Length : {0}", cutLengthData.lengthCompensation);
                        _log.Debug("Scale #1 - Weight : {0}", cutLengthData.weight);


                        //data send to MES
                        if (SendMessage(ref cutLengthData))
                        {
                            _log.Info("Data send complete... {0} / {1} / {2}", cutLengthData.heat.ToString(), cutLengthData.strandNo, cutLengthData.sequenceNo);

                            //Update last information
                            scale1LastHeat = cutLengthData.heat;
                            scale1LastStrandNo = cutLengthData.strandNo;
                            scale1LastSequenceNo = cutLengthData.sequenceNo;
                        }

                    }


                    /////////
                    /// No.2 Scale data
                    ////////                    
                    //data initialize
                    cutLengthData.InitData();
                    cutLengthData.scaleNo = 2;  //#2 scale

                    //Check Heat ID and Strand ID and Sequence No. for scale #1
                    byte[] heat2 = plc.ReadBytes(DataType.DataBlock, 244, 310, 6);   //Heat ID
                    cutLengthData.heat = S7.Net.Types.String.FromByteArray(heat2);
                    cutLengthData.heat.Trim();

                    cutLengthData.strandNo = ((ushort)plc.Read("DB244.DBW328")).ConvertToShort();
                    cutLengthData.sequenceNo = ((ushort)plc.Read("DB244.DBW326")).ConvertToShort();

                    heatChange = false; //init

                    _log.Info("Weight scale 2 : heat number check...{0}", cutLengthData.heat.ToString());

                    //Check basic information
                    if (cutLengthData.heat.Substring(0, 1).Equals("V")) //Check Heat ID
                    {
                        _log.Info("Scale #2 Information : {0} / {1} / {2}", cutLengthData.heat.ToString(), cutLengthData.strandNo, cutLengthData.sequenceNo);

                        if (cutLengthData.strandNo > 0 && cutLengthData.strandNo < 7) //Check strand number
                        {
                            if (cutLengthData.sequenceNo > 0 && cutLengthData.sequenceNo < 99) //Check sequence number
                            {
                                //Compare with before sent information
                                if (scale2LastHeat.ToString() != cutLengthData.heat.ToString())
                                {
                                    heatChange = true;
                                }
                                else
                                {
                                    if (scale2LastStrandNo != cutLengthData.strandNo || scale2LastSequenceNo != cutLengthData.sequenceNo)
                                    {
                                        heatChange = true;
                                    }
                                }
                            }
                        }
                    }

                    if (heatChange)
                    {
                        _log.Info("Scale #2 New SP scaled!!! Start collect data ... ");

                        //Collect data
                        cutLengthData.lengthTarget = ((uint)plc.Read("DB244.DBD494")).ConvertToInt();           //Target Length [DINT]
                        cutLengthData.lengthLastcut = ((uint)plc.Read("DB244.DBD502")).ConvertToFloat();        //Last Cut Length [REAL]
                        cutLengthData.lengthCompensation = ((ushort)plc.Read("DB244.DBW512")).ConvertToShort(); //Compensation Length [INT]
                        cutLengthData.weight = ((uint)plc.Read("DB244.DBD330")).ConvertToFloat();             //Weight [REAL???]

                        _log.Debug("Scale #2 - Target Length : {0}", cutLengthData.lengthTarget);
                        _log.Debug("Scale #2 - Last Cut Length : {0}", cutLengthData.lengthLastcut);
                        _log.Debug("Scale #2 - Compensation Length : {0}", cutLengthData.lengthCompensation);
                        _log.Debug("Scale #2 - Weight : {0}", cutLengthData.weight);

                        //data send to MES
                        if (SendMessage(ref cutLengthData))
                        {
                            _log.Info("Data send complete... {0} / {1} / {2}", cutLengthData.heat.ToString(), cutLengthData.strandNo, cutLengthData.sequenceNo);

                            //Update last information
                            scale2LastHeat = cutLengthData.heat;
                            scale2LastStrandNo = cutLengthData.strandNo;
                            scale2LastSequenceNo = cutLengthData.sequenceNo;
                        }

                    }
                    _log.Info("==================================================");
                    #endregion

                   

                    //Console.Write(".");
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
                    "ISNULL((SELECT AVG_VALUE FROM REP_CCM_STRAND_VARS WHERE REPORT_COUNTER = " + reportno + " AND STRAND_NO = 6 AND VARIABLE_ID = 70),0) AS MW_B_STD6, " +
                    "ISNULL((SELECT AVG_VALUE FROM REP_CCM_STRAND_VARS WHERE REPORT_COUNTER = " + reportno + " AND STRAND_NO = 1 AND VARIABLE_ID = 13 AND VALUE_CODE = 1),0) AS MP_A_STD1, " +
                    "ISNULL((SELECT AVG_VALUE FROM REP_CCM_STRAND_VARS WHERE REPORT_COUNTER = " + reportno + " AND STRAND_NO = 2 AND VARIABLE_ID = 13 AND VALUE_CODE = 1),0) AS MP_A_STD2, " +
                    "ISNULL((SELECT AVG_VALUE FROM REP_CCM_STRAND_VARS WHERE REPORT_COUNTER = " + reportno + " AND STRAND_NO = 3 AND VARIABLE_ID = 13 AND VALUE_CODE = 1),0) AS MP_A_STD3, " +
                    "ISNULL((SELECT AVG_VALUE FROM REP_CCM_STRAND_VARS WHERE REPORT_COUNTER = " + reportno + " AND STRAND_NO = 4 AND VARIABLE_ID = 13 AND VALUE_CODE = 1),0) AS MP_A_STD4, " +
                    "ISNULL((SELECT AVG_VALUE FROM REP_CCM_STRAND_VARS WHERE REPORT_COUNTER = " + reportno + " AND STRAND_NO = 5 AND VARIABLE_ID = 13 AND VALUE_CODE = 1),0) AS MP_A_STD5, " +
                    "ISNULL((SELECT AVG_VALUE FROM REP_CCM_STRAND_VARS WHERE REPORT_COUNTER = " + reportno + " AND STRAND_NO = 6 AND VARIABLE_ID = 13 AND VALUE_CODE = 1),0) AS MP_A_STD6, " +
                    "ISNULL((SELECT AVG_VALUE FROM REP_CCM_STRAND_VARS WHERE REPORT_COUNTER = " + reportno + " AND STRAND_NO = 1 AND VARIABLE_ID = 16 AND VALUE_CODE = 1),0) AS MP_B_STD1, " +
                    "ISNULL((SELECT AVG_VALUE FROM REP_CCM_STRAND_VARS WHERE REPORT_COUNTER = " + reportno + " AND STRAND_NO = 2 AND VARIABLE_ID = 16 AND VALUE_CODE = 1),0) AS MP_B_STD2, " +
                    "ISNULL((SELECT AVG_VALUE FROM REP_CCM_STRAND_VARS WHERE REPORT_COUNTER = " + reportno + " AND STRAND_NO = 3 AND VARIABLE_ID = 16 AND VALUE_CODE = 1),0) AS MP_B_STD3, " +
                    "ISNULL((SELECT AVG_VALUE FROM REP_CCM_STRAND_VARS WHERE REPORT_COUNTER = " + reportno + " AND STRAND_NO = 4 AND VARIABLE_ID = 16 AND VALUE_CODE = 1),0) AS MP_B_STD4, " +
                    "ISNULL((SELECT AVG_VALUE FROM REP_CCM_STRAND_VARS WHERE REPORT_COUNTER = " + reportno + " AND STRAND_NO = 5 AND VARIABLE_ID = 16 AND VALUE_CODE = 1),0) AS MP_B_STD5, " +
                    "ISNULL((SELECT AVG_VALUE FROM REP_CCM_STRAND_VARS WHERE REPORT_COUNTER = " + reportno + " AND STRAND_NO = 6 AND VARIABLE_ID = 16 AND VALUE_CODE = 1),0) AS MP_B_STD6, " +
                    "ISNULL((SELECT MIN_VALUE FROM REP_CCM_STRAND_VARS WHERE REPORT_COUNTER = " + reportno + " AND VALUE_CODE = 1 AND STRAND_NO = 1 AND VARIABLE_ID = 5),0) AS STD1_CAST_SPD_MIN, " +
                    "ISNULL((SELECT AVG_VALUE FROM REP_CCM_STRAND_VARS WHERE REPORT_COUNTER = " + reportno + " AND VALUE_CODE = 1 AND STRAND_NO = 1 AND VARIABLE_ID = 5),0) AS STD1_CAST_SPD_AVG, " +
                    "ISNULL((SELECT MAX_VALUE FROM REP_CCM_STRAND_VARS WHERE REPORT_COUNTER = " + reportno + " AND VALUE_CODE = 1 AND STRAND_NO = 1 AND VARIABLE_ID = 5),0) AS STD1_CAST_SPD_MAX, " +
                    "ISNULL((SELECT MIN_VALUE FROM REP_CCM_STRAND_VARS WHERE REPORT_COUNTER = " + reportno + " AND VALUE_CODE = 1 AND STRAND_NO = 2 AND VARIABLE_ID = 5),0) AS STD2_CAST_SPD_MIN, " +
                    "ISNULL((SELECT AVG_VALUE FROM REP_CCM_STRAND_VARS WHERE REPORT_COUNTER = " + reportno + " AND VALUE_CODE = 1 AND STRAND_NO = 2 AND VARIABLE_ID = 5),0) AS STD2_CAST_SPD_AVG, " +
                    "ISNULL((SELECT MAX_VALUE FROM REP_CCM_STRAND_VARS WHERE REPORT_COUNTER = " + reportno + " AND VALUE_CODE = 1 AND STRAND_NO = 2 AND VARIABLE_ID = 5),0) AS STD2_CAST_SPD_MAX, " +
                    "ISNULL((SELECT MIN_VALUE FROM REP_CCM_STRAND_VARS WHERE REPORT_COUNTER = " + reportno + " AND VALUE_CODE = 1 AND STRAND_NO = 3 AND VARIABLE_ID = 5),0) AS STD3_CAST_SPD_MIN, " +
                    "ISNULL((SELECT AVG_VALUE FROM REP_CCM_STRAND_VARS WHERE REPORT_COUNTER = " + reportno + " AND VALUE_CODE = 1 AND STRAND_NO = 3 AND VARIABLE_ID = 5),0) AS STD3_CAST_SPD_AVG, " +
                    "ISNULL((SELECT MAX_VALUE FROM REP_CCM_STRAND_VARS WHERE REPORT_COUNTER = " + reportno + " AND VALUE_CODE = 1 AND STRAND_NO = 3 AND VARIABLE_ID = 5),0) AS STD3_CAST_SPD_MAX, " +
                    "ISNULL((SELECT MIN_VALUE FROM REP_CCM_STRAND_VARS WHERE REPORT_COUNTER = " + reportno + " AND VALUE_CODE = 1 AND STRAND_NO = 4 AND VARIABLE_ID = 5),0) AS STD4_CAST_SPD_MIN, " +
                    "ISNULL((SELECT AVG_VALUE FROM REP_CCM_STRAND_VARS WHERE REPORT_COUNTER = " + reportno + " AND VALUE_CODE = 1 AND STRAND_NO = 4 AND VARIABLE_ID = 5),0) AS STD4_CAST_SPD_AVG, " +
                    "ISNULL((SELECT MAX_VALUE FROM REP_CCM_STRAND_VARS WHERE REPORT_COUNTER = " + reportno + " AND VALUE_CODE = 1 AND STRAND_NO = 4 AND VARIABLE_ID = 5),0) AS STD4_CAST_SPD_MAX, " +
                    "ISNULL((SELECT MIN_VALUE FROM REP_CCM_STRAND_VARS WHERE REPORT_COUNTER = " + reportno + " AND VALUE_CODE = 1 AND STRAND_NO = 5 AND VARIABLE_ID = 5),0) AS STD5_CAST_SPD_MIN, " +
                    "ISNULL((SELECT AVG_VALUE FROM REP_CCM_STRAND_VARS WHERE REPORT_COUNTER = " + reportno + " AND VALUE_CODE = 1 AND STRAND_NO = 5 AND VARIABLE_ID = 5),0) AS STD5_CAST_SPD_AVG, " +
                    "ISNULL((SELECT MAX_VALUE FROM REP_CCM_STRAND_VARS WHERE REPORT_COUNTER = " + reportno + " AND VALUE_CODE = 1 AND STRAND_NO = 5 AND VARIABLE_ID = 5),0) AS STD5_CAST_SPD_MAX, " +
                    "ISNULL((SELECT MIN_VALUE FROM REP_CCM_STRAND_VARS WHERE REPORT_COUNTER = " + reportno + " AND VALUE_CODE = 1 AND STRAND_NO = 6 AND VARIABLE_ID = 5),0) AS STD6_CAST_SPD_MIN, " +
                    "ISNULL((SELECT AVG_VALUE FROM REP_CCM_STRAND_VARS WHERE REPORT_COUNTER = " + reportno + " AND VALUE_CODE = 1 AND STRAND_NO = 6 AND VARIABLE_ID = 5),0) AS STD6_CAST_SPD_AVG, " +
                    "ISNULL((SELECT MAX_VALUE FROM REP_CCM_STRAND_VARS WHERE REPORT_COUNTER = " + reportno + " AND VALUE_CODE = 1 AND STRAND_NO = 6 AND VARIABLE_ID = 5),0) AS STD6_CAST_SPD_MAX, " +
                    "ISNULL((SELECT MIN_VALUE FROM REP_CCM_STRAND_VARS WHERE REPORT_COUNTER = " + reportno + " AND VALUE_CODE = 1 AND STRAND_NO = 1 AND VARIABLE_ID = 2),0) AS STD1_MD_LV_MIN, " +
                    "ISNULL((SELECT AVG_VALUE FROM REP_CCM_STRAND_VARS WHERE REPORT_COUNTER = " + reportno + " AND VALUE_CODE = 1 AND STRAND_NO = 1 AND VARIABLE_ID = 2),0) AS STD1_MD_LV_AVG, " +
                    "ISNULL((SELECT MAX_VALUE FROM REP_CCM_STRAND_VARS WHERE REPORT_COUNTER = " + reportno + " AND VALUE_CODE = 1 AND STRAND_NO = 1 AND VARIABLE_ID = 2),0) AS STD1_MD_LV_MAX, " +
                    "ISNULL((SELECT MIN_VALUE FROM REP_CCM_STRAND_VARS WHERE REPORT_COUNTER = " + reportno + " AND VALUE_CODE = 1 AND STRAND_NO = 2 AND VARIABLE_ID = 2),0) AS STD2_MD_LV_MIN, " +
                    "ISNULL((SELECT AVG_VALUE FROM REP_CCM_STRAND_VARS WHERE REPORT_COUNTER = " + reportno + " AND VALUE_CODE = 1 AND STRAND_NO = 2 AND VARIABLE_ID = 2),0) AS STD2_MD_LV_AVG, " +
                    "ISNULL((SELECT MAX_VALUE FROM REP_CCM_STRAND_VARS WHERE REPORT_COUNTER = " + reportno + " AND VALUE_CODE = 1 AND STRAND_NO = 2 AND VARIABLE_ID = 2),0) AS STD2_MD_LV_MAX, " +
                    "ISNULL((SELECT MIN_VALUE FROM REP_CCM_STRAND_VARS WHERE REPORT_COUNTER = " + reportno + " AND VALUE_CODE = 1 AND STRAND_NO = 3 AND VARIABLE_ID = 2),0) AS STD3_MD_LV_MIN, " +
                    "ISNULL((SELECT AVG_VALUE FROM REP_CCM_STRAND_VARS WHERE REPORT_COUNTER = " + reportno + " AND VALUE_CODE = 1 AND STRAND_NO = 3 AND VARIABLE_ID = 2),0) AS STD3_MD_LV_AVG, " +
                    "ISNULL((SELECT MAX_VALUE FROM REP_CCM_STRAND_VARS WHERE REPORT_COUNTER = " + reportno + " AND VALUE_CODE = 1 AND STRAND_NO = 3 AND VARIABLE_ID = 2),0) AS STD3_MD_LV_MAX, " +
                    "ISNULL((SELECT MIN_VALUE FROM REP_CCM_STRAND_VARS WHERE REPORT_COUNTER = " + reportno + " AND VALUE_CODE = 1 AND STRAND_NO = 4 AND VARIABLE_ID = 2),0) AS STD4_MD_LV_MIN, " +
                    "ISNULL((SELECT AVG_VALUE FROM REP_CCM_STRAND_VARS WHERE REPORT_COUNTER = " + reportno + " AND VALUE_CODE = 1 AND STRAND_NO = 4 AND VARIABLE_ID = 2),0) AS STD4_MD_LV_AVG, " +
                    "ISNULL((SELECT MAX_VALUE FROM REP_CCM_STRAND_VARS WHERE REPORT_COUNTER = " + reportno + " AND VALUE_CODE = 1 AND STRAND_NO = 4 AND VARIABLE_ID = 2),0) AS STD4_MD_LV_MAX, " +
                    "ISNULL((SELECT MIN_VALUE FROM REP_CCM_STRAND_VARS WHERE REPORT_COUNTER = " + reportno + " AND VALUE_CODE = 1 AND STRAND_NO = 5 AND VARIABLE_ID = 2),0) AS STD5_MD_LV_MIN, " +
                    "ISNULL((SELECT AVG_VALUE FROM REP_CCM_STRAND_VARS WHERE REPORT_COUNTER = " + reportno + " AND VALUE_CODE = 1 AND STRAND_NO = 5 AND VARIABLE_ID = 2),0) AS STD5_MD_LV_AVG, " +
                    "ISNULL((SELECT MAX_VALUE FROM REP_CCM_STRAND_VARS WHERE REPORT_COUNTER = " + reportno + " AND VALUE_CODE = 1 AND STRAND_NO = 5 AND VARIABLE_ID = 2),0) AS STD5_MD_LV_MAX, " +
                    "ISNULL((SELECT MIN_VALUE FROM REP_CCM_STRAND_VARS WHERE REPORT_COUNTER = " + reportno + " AND VALUE_CODE = 1 AND STRAND_NO = 6 AND VARIABLE_ID = 2),0) AS STD6_MD_LV_MIN, " +
                    "ISNULL((SELECT AVG_VALUE FROM REP_CCM_STRAND_VARS WHERE REPORT_COUNTER = " + reportno + " AND VALUE_CODE = 1 AND STRAND_NO = 6 AND VARIABLE_ID = 2),0) AS STD6_MD_LV_AVG, " +
                    "ISNULL((SELECT MAX_VALUE FROM REP_CCM_STRAND_VARS WHERE REPORT_COUNTER = " + reportno + " AND VALUE_CODE = 1 AND STRAND_NO = 6 AND VARIABLE_ID = 2),0) AS STD6_MD_LV_MAX " +
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
                            if (cdata.moldwaterInletStd1A < 0) cdata.moldwaterInletStd1A = 0;
                            cdata.moldwaterInletStd1B = Convert.ToDouble(reader["MW_B_STD1"]);
                            if (cdata.moldwaterInletStd1B < 0) cdata.moldwaterInletStd1B = 0;
                            cdata.moldwaterInletStd2A = Convert.ToDouble(reader["MW_A_STD2"]);
                            if (cdata.moldwaterInletStd2A < 0) cdata.moldwaterInletStd2A = 0;
                            cdata.moldwaterInletStd2B = Convert.ToDouble(reader["MW_B_STD2"]);
                            if (cdata.moldwaterInletStd2B < 0) cdata.moldwaterInletStd2B = 0;
                            cdata.moldwaterInletStd3A = Convert.ToDouble(reader["MW_A_STD3"]);
                            if (cdata.moldwaterInletStd3A < 0) cdata.moldwaterInletStd3A = 0;
                            cdata.moldwaterInletStd3B = Convert.ToDouble(reader["MW_B_STD3"]);
                            if (cdata.moldwaterInletStd3B < 0) cdata.moldwaterInletStd3B = 0;
                            cdata.moldwaterInletStd4A = Convert.ToDouble(reader["MW_A_STD4"]);
                            if (cdata.moldwaterInletStd4A < 0) cdata.moldwaterInletStd4A = 0;
                            cdata.moldwaterInletStd4B = Convert.ToDouble(reader["MW_B_STD4"]);
                            if (cdata.moldwaterInletStd4B < 0) cdata.moldwaterInletStd4B = 0;
                            cdata.moldwaterInletStd5A = Convert.ToDouble(reader["MW_A_STD5"]);
                            if (cdata.moldwaterInletStd5A < 0) cdata.moldwaterInletStd5A = 0;
                            cdata.moldwaterInletStd5B = Convert.ToDouble(reader["MW_B_STD5"]);
                            if (cdata.moldwaterInletStd5B < 0) cdata.moldwaterInletStd5B = 0;
                            cdata.moldwaterInletStd6A = Convert.ToDouble(reader["MW_A_STD6"]);
                            if (cdata.moldwaterInletStd6A < 0) cdata.moldwaterInletStd6A = 0;
                            cdata.moldwaterInletStd6B = Convert.ToDouble(reader["MW_B_STD6"]);
                            if (cdata.moldwaterInletStd6B < 0) cdata.moldwaterInletStd6B = 0;

                            cdata.moldwaterOutletPressureStd1A = Convert.ToDouble(reader["MP_A_STD1"]);
                            if (cdata.moldwaterOutletPressureStd1A < 0) cdata.moldwaterOutletPressureStd1A = 0;
                            cdata.moldwaterOutletPressureStd1B = Convert.ToDouble(reader["MP_B_STD1"]);
                            if (cdata.moldwaterOutletPressureStd1B < 0) cdata.moldwaterOutletPressureStd1B = 0;
                            cdata.moldwaterOutletPressureStd2A = Convert.ToDouble(reader["MP_A_STD2"]);
                            if (cdata.moldwaterOutletPressureStd2A < 0) cdata.moldwaterOutletPressureStd2A = 0;
                            cdata.moldwaterOutletPressureStd2B = Convert.ToDouble(reader["MP_B_STD2"]);
                            if (cdata.moldwaterOutletPressureStd2B < 0) cdata.moldwaterOutletPressureStd2B = 0;
                            cdata.moldwaterOutletPressureStd3A = Convert.ToDouble(reader["MP_A_STD3"]);
                            if (cdata.moldwaterOutletPressureStd3A < 0) cdata.moldwaterOutletPressureStd3A = 0;
                            cdata.moldwaterOutletPressureStd3B = Convert.ToDouble(reader["MP_B_STD3"]);
                            if (cdata.moldwaterOutletPressureStd3B < 0) cdata.moldwaterOutletPressureStd3B = 0;
                            cdata.moldwaterOutletPressureStd4A = Convert.ToDouble(reader["MP_A_STD4"]);
                            if (cdata.moldwaterOutletPressureStd4A < 0) cdata.moldwaterOutletPressureStd4A = 0;
                            cdata.moldwaterOutletPressureStd4B = Convert.ToDouble(reader["MP_B_STD4"]);
                            if (cdata.moldwaterOutletPressureStd4B < 0) cdata.moldwaterOutletPressureStd4B = 0;
                            cdata.moldwaterOutletPressureStd5A = Convert.ToDouble(reader["MP_A_STD5"]);
                            if (cdata.moldwaterOutletPressureStd5A < 0) cdata.moldwaterOutletPressureStd5A = 0;
                            cdata.moldwaterOutletPressureStd5B = Convert.ToDouble(reader["MP_B_STD5"]);
                            if (cdata.moldwaterOutletPressureStd5B < 0) cdata.moldwaterOutletPressureStd5B = 0;
                            cdata.moldwaterOutletPressureStd6A = Convert.ToDouble(reader["MP_A_STD6"]);
                            if (cdata.moldwaterOutletPressureStd6A < 0) cdata.moldwaterOutletPressureStd6A = 0;
                            cdata.moldwaterOutletPressureStd6B = Convert.ToDouble(reader["MP_B_STD6"]);
                            if (cdata.moldwaterOutletPressureStd6B < 0) cdata.moldwaterOutletPressureStd6B = 0;

                            cdata.castSpeedMinStrand1 = Convert.ToDouble(reader["STD1_CAST_SPD_MIN"]);
                            if (cdata.castSpeedMinStrand1 < 0) cdata.castSpeedMinStrand1 = 0;
                            cdata.castSpeedAvgStrand1 = Convert.ToDouble(reader["STD1_CAST_SPD_AVG"]);
                            if (cdata.castSpeedAvgStrand1 < 0) cdata.castSpeedAvgStrand1 = 0;
                            cdata.castSpeedMaxStrand1 = Convert.ToDouble(reader["STD1_CAST_SPD_MAX"]);
                            if (cdata.castSpeedMaxStrand1 < 0) cdata.castSpeedMaxStrand1 = 0;
                            cdata.castSpeedMinStrand2 = Convert.ToDouble(reader["STD2_CAST_SPD_MIN"]);
                            if (cdata.castSpeedMinStrand2 < 0) cdata.castSpeedMinStrand2 = 0;
                            cdata.castSpeedAvgStrand2 = Convert.ToDouble(reader["STD2_CAST_SPD_AVG"]);
                            if (cdata.castSpeedAvgStrand2 < 0) cdata.castSpeedAvgStrand2 = 0;
                            cdata.castSpeedMaxStrand2 = Convert.ToDouble(reader["STD2_CAST_SPD_MAX"]);
                            if (cdata.castSpeedMaxStrand2 < 0) cdata.castSpeedMaxStrand2 = 0;
                            cdata.castSpeedMinStrand3 = Convert.ToDouble(reader["STD3_CAST_SPD_MIN"]);
                            if (cdata.castSpeedMinStrand3 < 0) cdata.castSpeedMinStrand3 = 0;
                            cdata.castSpeedAvgStrand3 = Convert.ToDouble(reader["STD3_CAST_SPD_AVG"]);
                            if (cdata.castSpeedAvgStrand3 < 0) cdata.castSpeedAvgStrand3 = 0;
                            cdata.castSpeedMaxStrand3 = Convert.ToDouble(reader["STD3_CAST_SPD_MAX"]);
                            if (cdata.castSpeedMaxStrand3 < 0) cdata.castSpeedMaxStrand3 = 0;
                            cdata.castSpeedMinStrand4 = Convert.ToDouble(reader["STD4_CAST_SPD_MIN"]);
                            if (cdata.castSpeedMinStrand4 < 0) cdata.castSpeedMinStrand4 = 0;
                            cdata.castSpeedAvgStrand4 = Convert.ToDouble(reader["STD4_CAST_SPD_AVG"]);
                            if (cdata.castSpeedAvgStrand4 < 0) cdata.castSpeedAvgStrand4 = 0;
                            cdata.castSpeedMaxStrand4 = Convert.ToDouble(reader["STD4_CAST_SPD_MAX"]);
                            if (cdata.castSpeedMaxStrand4 < 0) cdata.castSpeedMaxStrand4 = 0;
                            cdata.castSpeedMinStrand5 = Convert.ToDouble(reader["STD5_CAST_SPD_MIN"]);
                            if (cdata.castSpeedMinStrand5 < 0) cdata.castSpeedMinStrand5 = 0;
                            cdata.castSpeedAvgStrand5 = Convert.ToDouble(reader["STD5_CAST_SPD_AVG"]);
                            if (cdata.castSpeedAvgStrand5 < 0) cdata.castSpeedAvgStrand5 = 0;
                            cdata.castSpeedMaxStrand5 = Convert.ToDouble(reader["STD5_CAST_SPD_MAX"]);
                            if (cdata.castSpeedMaxStrand5 < 0) cdata.castSpeedMaxStrand5 = 0;
                            cdata.castSpeedMinStrand6 = Convert.ToDouble(reader["STD6_CAST_SPD_MIN"]);
                            if (cdata.castSpeedMinStrand6 < 0) cdata.castSpeedMinStrand6 = 0;
                            cdata.castSpeedAvgStrand6 = Convert.ToDouble(reader["STD6_CAST_SPD_AVG"]);
                            if (cdata.castSpeedAvgStrand6 < 0) cdata.castSpeedAvgStrand6 = 0;
                            cdata.castSpeedMaxStrand6 = Convert.ToDouble(reader["STD6_CAST_SPD_MAX"]);
                            if (cdata.castSpeedMaxStrand6 < 0) cdata.castSpeedMaxStrand6 = 0;

                            cdata.mouldLevelMinStrand1 = Convert.ToDouble(reader["STD1_MD_LV_MIN"]);
                            if (cdata.mouldLevelMinStrand1 < 0) cdata.mouldLevelMinStrand1 = 0;
                            cdata.mouldLevelAvgStrand1 = Convert.ToDouble(reader["STD1_MD_LV_AVG"]);
                            if (cdata.mouldLevelAvgStrand1 < 0) cdata.mouldLevelAvgStrand1 = 0;
                            cdata.mouldLevelMaxStrand1 = Convert.ToDouble(reader["STD1_MD_LV_MAX"]);
                            if (cdata.mouldLevelMaxStrand1 < 0) cdata.mouldLevelMaxStrand1 = 0;
                            cdata.mouldLevelMinStrand2 = Convert.ToDouble(reader["STD2_MD_LV_MIN"]);
                            if (cdata.mouldLevelMinStrand2 < 0) cdata.mouldLevelMinStrand2 = 0;
                            cdata.mouldLevelAvgStrand2 = Convert.ToDouble(reader["STD2_MD_LV_AVG"]);
                            if (cdata.mouldLevelAvgStrand2 < 0) cdata.mouldLevelAvgStrand2 = 0;
                            cdata.mouldLevelMaxStrand2 = Convert.ToDouble(reader["STD2_MD_LV_MAX"]);
                            if (cdata.mouldLevelMaxStrand2 < 0) cdata.mouldLevelMaxStrand2 = 0;
                            cdata.mouldLevelMinStrand3 = Convert.ToDouble(reader["STD3_MD_LV_MIN"]);
                            if (cdata.mouldLevelMinStrand3 < 0) cdata.mouldLevelMinStrand3 = 0;
                            cdata.mouldLevelAvgStrand3 = Convert.ToDouble(reader["STD3_MD_LV_AVG"]);
                            if (cdata.mouldLevelAvgStrand3 < 0) cdata.mouldLevelAvgStrand3 = 0;
                            cdata.mouldLevelMaxStrand3 = Convert.ToDouble(reader["STD3_MD_LV_MAX"]);
                            if (cdata.mouldLevelMaxStrand3 < 0) cdata.mouldLevelMaxStrand3 = 0;
                            cdata.mouldLevelMinStrand4 = Convert.ToDouble(reader["STD4_MD_LV_MIN"]);
                            if (cdata.mouldLevelMinStrand4 < 0) cdata.mouldLevelMinStrand4 = 0;
                            cdata.mouldLevelAvgStrand4 = Convert.ToDouble(reader["STD4_MD_LV_AVG"]);
                            if (cdata.mouldLevelAvgStrand4 < 0) cdata.mouldLevelAvgStrand4 = 0;
                            cdata.mouldLevelMaxStrand4 = Convert.ToDouble(reader["STD4_MD_LV_MAX"]);
                            if (cdata.mouldLevelMaxStrand4 < 0) cdata.mouldLevelMaxStrand4 = 0;
                            cdata.mouldLevelMinStrand5 = Convert.ToDouble(reader["STD5_MD_LV_MIN"]);
                            if (cdata.mouldLevelMinStrand5 < 0) cdata.mouldLevelMinStrand5 = 0;
                            cdata.mouldLevelAvgStrand5 = Convert.ToDouble(reader["STD5_MD_LV_AVG"]);
                            if (cdata.mouldLevelAvgStrand5 < 0) cdata.mouldLevelAvgStrand5 = 0;
                            cdata.mouldLevelMaxStrand5 = Convert.ToDouble(reader["STD5_MD_LV_MAX"]);
                            if (cdata.mouldLevelMaxStrand5 < 0) cdata.mouldLevelMaxStrand5 = 0;
                            cdata.mouldLevelMinStrand6 = Convert.ToDouble(reader["STD6_MD_LV_MIN"]);
                            if (cdata.mouldLevelMinStrand6 < 0) cdata.mouldLevelMinStrand6 = 0;
                            cdata.mouldLevelAvgStrand6 = Convert.ToDouble(reader["STD6_MD_LV_AVG"]);
                            if (cdata.mouldLevelAvgStrand6 < 0) cdata.mouldLevelAvgStrand6 = 0;
                            cdata.mouldLevelMaxStrand6 = Convert.ToDouble(reader["STD6_MD_LV_MAX"]);
                            if (cdata.mouldLevelMaxStrand6 < 0) cdata.mouldLevelMaxStrand6 = 0;

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
                msgHead += "000473".PadRight(31);                   //message length & spare 

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

                decData = (int)(msg.moldwaterOutletPressureStd1A * 10000);  //mold water pressure stand 1 A line (decimal * 4)
                msgData += decData.ToString("D6");
                decData = (int)(msg.moldwaterOutletPressureStd1B * 10000);
                msgData += decData.ToString("D6");
                decData = (int)(msg.moldwaterOutletPressureStd2A * 10000);  
                msgData += decData.ToString("D6");
                decData = (int)(msg.moldwaterOutletPressureStd2B * 10000);
                msgData += decData.ToString("D6");
                decData = (int)(msg.moldwaterOutletPressureStd3A * 10000);
                msgData += decData.ToString("D6");
                decData = (int)(msg.moldwaterOutletPressureStd3B * 10000);
                msgData += decData.ToString("D6");
                decData = (int)(msg.moldwaterOutletPressureStd4A * 10000);
                msgData += decData.ToString("D6");
                decData = (int)(msg.moldwaterOutletPressureStd4B * 10000);
                msgData += decData.ToString("D6");
                decData = (int)(msg.moldwaterOutletPressureStd5A * 10000);
                msgData += decData.ToString("D6");
                decData = (int)(msg.moldwaterOutletPressureStd5B * 10000);
                msgData += decData.ToString("D6");
                decData = (int)(msg.moldwaterOutletPressureStd6A * 10000);
                msgData += decData.ToString("D6");
                decData = (int)(msg.moldwaterOutletPressureStd6B * 10000);
                msgData += decData.ToString("D6");

                decData = (int)(msg.castSpeedMinStrand1 * 10000);   //cast speed (decimal * 4)
                msgData += decData.ToString("D6");
                decData = (int)(msg.castSpeedAvgStrand1 * 10000);   
                msgData += decData.ToString("D6");
                decData = (int)(msg.castSpeedMaxStrand1 * 10000);
                msgData += decData.ToString("D6");
                decData = (int)(msg.castSpeedMinStrand2 * 10000);   
                msgData += decData.ToString("D6");
                decData = (int)(msg.castSpeedAvgStrand2 * 10000);
                msgData += decData.ToString("D6");
                decData = (int)(msg.castSpeedMaxStrand2 * 10000);
                msgData += decData.ToString("D6");
                decData = (int)(msg.castSpeedMinStrand3 * 10000);   
                msgData += decData.ToString("D6");
                decData = (int)(msg.castSpeedAvgStrand3 * 10000);
                msgData += decData.ToString("D6");
                decData = (int)(msg.castSpeedMaxStrand3 * 10000);
                msgData += decData.ToString("D6");
                decData = (int)(msg.castSpeedMinStrand4 * 10000);   
                msgData += decData.ToString("D6");
                decData = (int)(msg.castSpeedAvgStrand4 * 10000);
                msgData += decData.ToString("D6");
                decData = (int)(msg.castSpeedMaxStrand4 * 10000);
                msgData += decData.ToString("D6");
                decData = (int)(msg.castSpeedMinStrand5 * 10000);   
                msgData += decData.ToString("D6");
                decData = (int)(msg.castSpeedAvgStrand5 * 10000);
                msgData += decData.ToString("D6");
                decData = (int)(msg.castSpeedMaxStrand5 * 10000);
                msgData += decData.ToString("D6");
                decData = (int)(msg.castSpeedMinStrand6 * 10000);   
                msgData += decData.ToString("D6");
                decData = (int)(msg.castSpeedAvgStrand6 * 10000);
                msgData += decData.ToString("D6");
                decData = (int)(msg.castSpeedMaxStrand6 * 10000);
                msgData += decData.ToString("D6");

                decData = (int)(msg.mouldLevelMinStrand1 * 100);    //mould level (deciaml * 2)
                msgData += decData.ToString("D6");
                decData = (int)(msg.mouldLevelAvgStrand1 * 100);
                msgData += decData.ToString("D6");
                decData = (int)(msg.mouldLevelMaxStrand1 * 100);
                msgData += decData.ToString("D6");
                decData = (int)(msg.mouldLevelMinStrand2 * 100);
                msgData += decData.ToString("D6");
                decData = (int)(msg.mouldLevelAvgStrand2 * 100);
                msgData += decData.ToString("D6");
                decData = (int)(msg.mouldLevelMaxStrand2 * 100);
                msgData += decData.ToString("D6");
                decData = (int)(msg.mouldLevelMinStrand3 * 100);
                msgData += decData.ToString("D6");
                decData = (int)(msg.mouldLevelAvgStrand3 * 100);
                msgData += decData.ToString("D6");
                decData = (int)(msg.mouldLevelMaxStrand3 * 100);
                msgData += decData.ToString("D6");
                decData = (int)(msg.mouldLevelMinStrand4 * 100);
                msgData += decData.ToString("D6");
                decData = (int)(msg.mouldLevelAvgStrand4 * 100);
                msgData += decData.ToString("D6");
                decData = (int)(msg.mouldLevelMaxStrand4 * 100);
                msgData += decData.ToString("D6");
                decData = (int)(msg.mouldLevelMinStrand5 * 100);
                msgData += decData.ToString("D6");
                decData = (int)(msg.mouldLevelAvgStrand5 * 100);
                msgData += decData.ToString("D6");
                decData = (int)(msg.mouldLevelMaxStrand5 * 100);
                msgData += decData.ToString("D6");
                decData = (int)(msg.mouldLevelMinStrand6 * 100);
                msgData += decData.ToString("D6");
                decData = (int)(msg.mouldLevelAvgStrand6 * 100);
                msgData += decData.ToString("D6");
                decData = (int)(msg.mouldLevelMaxStrand6 * 100);
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


        private static void InitializePLC()
        {
            bool bConnPlc = false;

            _log.Warn("PLC Connection initialize start.....");

            try
            {
                // Connect PLC
                if (plc.IsConnected == false)
                {
                    try
                    {
                        plc.Open();
                        bConnPlc = true;
                        _log.Info("PLC communication initialize success...");
                    }
                    catch (Exception)
                    {
                        throw;
                    }
                }
            }
            catch (Exception e)
            {
                _log.Error(e.Message);
            }

            if (bConnPlc)
            {
                plcStatus = true;
            }
        }

        private static bool GetTundishData(ref TundishData cdata)
        {
            bool flag = false;

            try
            {
                //check current casting status                                     
                //bool tcar1 = (bool)plc.Read("DB215.DBX373.3");  //tundish car 1 cast position : DB215.DBX373.3
                //bool tcar2 = (bool)plc.Read("DB215.DBX461.3");  //tundish car 2 cast position : DB215.DBX461.3
                //_log.Info("{0}", tcar1.ToString());
                //_log.Info("{0}", tcar2.ToString());

                //collect - data initialize
                cdata.InitData();

                //tundish tracking heat ID : GEN_PLC DB220.1880 [STRING - 16 char]                    
                byte[] byteArray = plc.ReadBytes(DataType.DataBlock, 220, 1880, 6);
                cdata.heat = S7.Net.Types.String.FromByteArray(byteArray);
                cdata.heat.Trim();
                
                //tundish weight(actual net weight) : GEN_PLC DB220.112 real decimal 2 [0.00 ~ 50.00] ton
                cdata.tundishWeight = ((uint)plc.Read("DB220.DBD112")).ConvertToFloat();                
                //tundish height(vertical position) : GEN_PLC 220.188 real decimal 0 [0 ~ 1000] mm
                cdata.tundishHeight = ((uint)plc.Read("DB220.DBD188")).ConvertToFloat();

                if (cdata.heat.Substring(0,1).Equals("V"))
                {
                    flag = true;    //collect data exist
                }                

            }
            catch (Exception e)
            {
                _log.Error(e.Message);
                //plc disconnect
                if (plc != null)
                {
                    plc.Close();
                    plcStatus = false;
                }
            }

            return flag;
        }

        private static bool SendMessageCyc(ref TundishData msg)
        {
            string msgHead = string.Empty;
            string msgData = string.Empty;
            string sql = string.Empty;
            bool flag = false;
            int decData = 0;

            try
            {
                // Set message - Header
                msgHead += "M20RL072";                              //TC Code
                msgHead += "2000";                                  //send factory code
                msgHead += "L06";                                   //sending process code
                msgHead += "2000";                                  //receive factory code
                msgHead += "M20";                                   //receive process code
                msgHead += DateTime.Now.ToString("yyyyMMddHHmmss"); //sending date time
                msgHead += "L3SentryCCMMAN";                        //send program ID
                msgHead += "RM20L06_01".PadRight(19);               //EAI IF ID (Queue name)
                msgHead += "000130".PadRight(31);                   //message length & spare 

                // Set message - Data
                msgData += DateTime.Now.ToString("yyyyMMddHHmmss"); //datetime
                msgData += msg.heat.ToString().PadRight(6);         //heat id

                decData = (int)(msg.tundishWeight * 100);       //tundish weight ton (decimal *2)
                msgData += decData.ToString("D6");
                decData = (int)(msg.tundishHeight);             //tundish height mm
                msgData += decData.ToString("D4");

                // Insert database
                sql = "INSERT INTO TT_L2_L3_NEW (HEADER, DATA, MSG_CODE, INTERFACE_ID) VALUES (" +
                       "'" + msgHead + "'," +
                       "'" + msgData + "'," +
                       "'M20RL072'," +
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


        private static bool SendMessage(ref CutLengthData msg)
        {
            string msgHead = string.Empty;
            string msgData = string.Empty;
            string sql = string.Empty;
            bool flag = false;
            int decData = 0;

            try
            {
                // Set message - Header
                msgHead += "M20RL073";                              //TC Code
                msgHead += "2000";                                  //send factory code
                msgHead += "L06";                                   //sending process code
                msgHead += "2000";                                  //receive factory code
                msgHead += "M20";                                   //receive process code
                msgHead += DateTime.Now.ToString("yyyyMMddHHmmss"); //sending date time
                msgHead += "L3SentryCCMMAN";                        //send program ID
                msgHead += "RM20L06_01".PadRight(19);               //EAI IF ID (Queue name)
                msgHead += "000130".PadRight(31);                   //message length & spare 

                // Set message - Data
                msgData += msg.heat.ToString().PadRight(6);         //heat id
                msgData += msg.strandNo.ToString("D1");             //strand number
                msgData += msg.sequenceNo.ToString("D2");           //sequence number 

                decData = (int)msg.lengthLastcut;
                msgData += decData.ToString("D5");                  //Order length

                if (msg.lengthCompensation < 0) msgData += msg.lengthCompensation.ToString("D4");   //Compensation length
                else                            msgData += msg.lengthCompensation.ToString("D5");

                decData = (int)msg.lengthTarget;
                msgData += decData.ToString("D5");                  //Target length

                msgData += msg.scaleNo.ToString("D1");              //scale number
                
                decData = (int)msg.weight;
                if (decData < 0) msgData += decData.ToString("D4"); //weight
                else             msgData += decData.ToString("D5");          
                

                // Insert database
                sql = "INSERT INTO TT_L2_L3_NEW (HEADER, DATA, MSG_CODE, INTERFACE_ID) VALUES (" +
                       "'" + msgHead + "'," +
                       "'" + msgData + "'," +
                       "'M20RL073'," +
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

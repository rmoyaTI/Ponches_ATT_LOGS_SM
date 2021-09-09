/**********************************************************
 * Interfase ZKSoftware + WS SAP                       *
 * Autor: Giancarlo Marte                                 *
 * Fecha: 18.04.2018                                      *
***********************************************************/
using System;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.ComponentModel;
using System.Configuration;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Data.OleDb;
using System.IO;
using System.Net.Mail;
using System.Threading;
//using AttLogs.WSPonches;
using System.ServiceModel;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Net;
using System.Threading.Tasks;
using Newtonsoft.Json;
using Newtonsoft.Json.Serialization;
using zkemkeeper;
using System.Linq;
//using System.Web.Script.Serialization;

namespace AttLogs
{
    public partial class AttLogsMain : Form
    {
        static HttpClient client = new HttpClient();
        char isBacked;
        List<EXRelojes> lstExRelojes = new List<EXRelojes>();
        public AttLogsMain()
        {
            InitializeComponent();
        }
        //Create Standalone SDK class dynamicly.        
        public CZKEM axCZKEM1 = new zkemkeeper.CZKEM();

        
        #region Fields
        private DataTable mitabla = new DataTable();
        private OleDbDataReader reader;
        private bool bIsConnected = false;//the boolean value identifies whether the device is connected        
        private int iMachineNumber = 1;//the serial number of the device.After connecting the device ,this value will be changed.        
        #endregion        

        #region OnLoad
        private void AttLogsMain_Load(object sender, EventArgs e)
        {
            Thread.CurrentThread.CurrentCulture = new System.Globalization.CultureInfo("es-DO");
            Connect_Downloads();
            //            MessageBox.Show(System.Threading.Thread.CurrentThread.CurrentCulture.ToString());
        }
        #endregion

        #region TCPIP Communication

        private async void Connect_Downloads()
        {
           
            string sdwEnrollNumber = "";
            int idwVerifyMode = 0;
            int idwInOutMode = 0;
            int idwYear = 0;
            int idwMonth = 0;
            int idwDay = 0;
            int idwHour = 0;
            int idwMinute = 0;
            int idwSecond = 0;
            int idwWorkcode = 0;
            string sdwSerialNumber = "";
            int iGLCount = 0;
            int iIndex = 0;
            List<Machines> lstMachines = getMachines();
            
            int idwErrorCode = 0;
            foreach (Machines mc in lstMachines)
            {
                mc.attList = new List<Atendancelogs>();
                EXRelojes ExReloj = new EXRelojes(mc);
                ExReloj.status = false;
                Cursor = Cursors.WaitCursor;
                bIsConnected = axCZKEM1.Connect_Net(mc.Ip, mc.Port);
                if (bIsConnected == true)
                {
                    ExReloj.status = true;
                    axCZKEM1.SetDeviceTime(iMachineNumber);
                    iMachineNumber = 1;//In fact,when you are using the tcp/ip communication,this parameter will be ignored,that is any integer will all right.Here we use 1.
                    axCZKEM1.RegEvent(iMachineNumber, 65535);//Here you can register the realtime events that you want to be triggered(the parameters 65535 means registering
                   /*try
                    {
                        mc.attList = WaitFor<List<Atendancelogs>>.Run(TimeSpan.FromMinutes(10), () => GetAttLogs(mc.MachineNumber, mc.SerialNumber, ref ExReloj));
                        //mc.attList = GetAttLogs(mc.MachineNumber, mc.SerialNumber, ref ExReloj);
                    }
                    catch (TimeoutException ex)
                    {
                        ExReloj.Errores.Add("Leer data del reloj ha fallado, ErrorCode: " + ex.Message.ToString());
                        axCZKEM1.EnableDevice(iMachineNumber, true);//enable the device
                        ExReloj.status = false;
                        lstExRelojes.Add(ExReloj);
                        continue;
                    }*/

                    if (axCZKEM1.ReadGeneralLogData(iMachineNumber))//read all the attendance records to the memory
                    {
                        while (axCZKEM1.SSR_GetGeneralLogData(iMachineNumber, out sdwEnrollNumber, out idwVerifyMode,
                                    out idwInOutMode, out idwYear, out idwMonth, out idwDay, out idwHour, out idwMinute, out idwSecond, ref idwWorkcode))//get records from the memory
                        {
                         axCZKEM1.GetSerialNumber(iMachineNumber,out sdwSerialNumber);
                            mc.attList.Add(new Atendancelogs
                            {
                                BadgeNumber = sdwEnrollNumber,
                                CheckTime = new DateTime(idwYear, idwMonth, idwDay, idwHour, idwMinute, 0),
                                CheckType = "I",
                                MemoryInfo = "NULL",
                                SensorId = iMachineNumber.ToString(),
                                SerialNumber = sdwSerialNumber,
                                WorkCode = idwWorkcode
                                    
                        });
                            
                        }
                        
                    }
                    ExReloj.attList = mc.attList;
                    if (mc.attList.Count > 0 && mc.attList != null)
                    {
                        Console.WriteLine("Generating time events ");
                        //Console.Read();
                        if (await SendWebserviceAsync(mc, ExReloj))
                        {
                            if (SetFiles(mc, ref ExReloj))
                            {
                                var clear = ConfigurationManager.AppSettings["ClearLogs"];
                                if (clear == "true")
                                {
                                    Console.WriteLine("Clearing Logs ");
                                    //ClearLogs(ref ExReloj);       // clear logs of an connected machine         
                                }
                            }

                            // System.Threading.Thread.Sleep(30000);
                            /*if (SetFiles(mc, ref ExReloj))
                            {
                                
                                
                                //if (InsertAttlogs(mc, ref ExReloj))
                                //{ // insert AttLogs to the MDB DBase of the main app

                                //}
                                //else // in case of error try delete logs of the day in DB, and try to inserdna;lfas;ot again.
                                //{
                                //    if (FixErrors(mc, ref ExReloj))
                                //    {
                                //        if (InsertAttlogs(mc, ref ExReloj))
                                //        {
                                //            //ClearLogs(ref ExReloj);       // clear logs of an connected machine                 
                                //        }
                                //    }
                                //}
                            }*/
                        }
                    }
                }
                else
                {
                    axCZKEM1.GetLastError(ref idwErrorCode);
                    ExReloj.status = false;
                    ExReloj.Errores.Add("Unable to connect the device,ErrorCode=" + idwErrorCode.ToString());
                    lstExRelojes.Add(ExReloj);
                }
                axCZKEM1.Disconnect();
                Cursor = Cursors.Default;
                //lstExRelojes.Add(ExReloj);
            }
            SendResults(lstExRelojes);
            Application.Exit();
        }

        #endregion

        #region AttLogs

        private void ClearLogs(ref EXRelojes exR)
        {
            int idwErrorCode = 0;

            axCZKEM1.EnableDevice(iMachineNumber, false);//disable the device
            if (axCZKEM1.ClearGLog(iMachineNumber))
            {
                axCZKEM1.RefreshData(iMachineNumber);//the data in the device should be refreshed
            }
            else
            {
                axCZKEM1.GetLastError(ref idwErrorCode);
                exR.Errores.Add("ClearLogs failed,ErrorCode=" + idwErrorCode.ToString());
                exR.status = false;
            }
            axCZKEM1.EnableDevice(iMachineNumber, true);//enable the device
        }

        private static bool SetFiles(Machines v_mc, ref EXRelojes exR)
        {
            Console.WriteLine("Generating time event files ");
            string fileName = ConfigurationManager.AppSettings["rutaFILE"] + "\\" + DateTime.Now.ToString("yyyyMMdd") + "_" + DateTime.Now.ToString("HHmmss") + "_" +(string)v_mc.Name + ".TXT";
            try
            {
                List<Atendancelogs> logs = new List<Atendancelogs>();
                logs = v_mc.attList.GroupBy(x => new { x.BadgeNumber, x.CheckTime })
                                .Select(grp => grp.First()).ToList();

                StreamWriter writer = File.AppendText(fileName);
                foreach (Atendancelogs al in logs)
                {
                    writer.WriteLine(al.BadgeNumber.ToString().PadLeft(8,'0') + ":" + "P01:" + al.CheckTime.ToString("yyyyMMdd") + ":" + al.CheckTime.ToString("HHmmss") + ":" + v_mc.Name);
                }
                writer.Close();
            }
            catch (IOException ex)
            {
                exR.Errores.Add(ex.Message);
                exR.status = false;
                return false;
            }
            return true;
        }
        private async Task<bool> SendWebserviceAsync(Machines v_mc, EXRelojes ereloj)
        {
            int i = 0;
            string url = ConfigurationManager.AppSettings["SAPEndPoint"];
            string userName = ConfigurationManager.AppSettings["SAPUser"];
            string password = ConfigurationManager.AppSettings["SAPPassword"];
            string sapClient = ConfigurationManager.AppSettings["SAPClient"];            
            List<SAPEventPairs> events2 = new List<SAPEventPairs>();
            bool insert = false;

            DateTime btime;
            DateTime atime;            

            //SAPEventPairs events = new SAPEventPairs();
            List<Atendancelogs> logs = new List<Atendancelogs>();
            //List<Atendancelogs> orderedLogs = new List<Atendancelogs>();                           

            //System.Diagnostics.Debug.WriteLine("before sorting...");
            //PrintList(v_mc.attList);
            //logs = v_mc.attList.GroupBy(x => new { x.BadgeNumber, x.CheckTime })
            //      .Select(grp => grp.First()).ToList();
            //v_mc.attList.OrderBy(x => x.BadgeNumber);                    
                        
            //System.Diagnostics.Debug.WriteLine("after sorting...");
            //PrintList(v_mc.attList);
            foreach (Atendancelogs al in v_mc.attList)
            {
                insert = false;
                btime = al.CheckTime.AddMinutes(-15); // base para logica de +/- 15 minutos
                atime = al.CheckTime.AddMinutes(15);  // base para logica de +/- 15minutos  
                if (i != 0)
                {
                    insert = logs.Any(x => x.BadgeNumber == al.BadgeNumber && x.CheckTime <= atime && x.CheckTime >= btime);
                }
                //WSPonches.TEVEN t_teven = new TEVEN();                

                if (!insert)
                {
                    events2.Add(new SAPEventPairs()
                    {
                        pernr = al.BadgeNumber.ToString(),
                        ldate = al.CheckTime.ToString("yyyy-MM-dd"),
                        erdat = DateTime.Today.ToString("yyyy-MM-dd"),
                        ltime = al.CheckTime.ToString("HH:mm:ss"),
                        ertim = DateTime.Now.ToString("HH:mm:ss"),
                        satza = "P01",
                        terid = v_mc.Name
                    });                    
                    i++;
                    logs.Add(al);                
                }
                else {
                    System.Diagnostics.Debug.WriteLine("duplicated registry" + al.BadgeNumber + " => " + al.CheckTime);
                    System.Diagnostics.Debug.WriteLine("after sorting...");
                    PrintList(logs);
                }
            }
            JsonSerializer serializaer = new JsonSerializer {
                ContractResolver = new CamelCasePropertyNamesContractResolver()
            };

            var json = JsonConvert.SerializeObject(events2);
            //json = json.Replace("\"", "");

            //ws.INS_TEVEN = teven;

            try
            {
             return await UploadTimeEvents(url, userName, password,json, sapClient, v_mc, ereloj);
                
                //WSPonches.ZWS_CARGAIT2011Response result = client.ZWS_CARGAIT2011(ws);
            }
            catch (Exception ex)
            {
                return false;
            }
            finally
            {
                
            }

        }

        public async Task<bool> UploadTimeEvents(string url, string user, string password, string json, string sapClient, Machines mc, EXRelojes ExReloj)
        {
            //EXRelojes ExReloj = new EXRelojes(mc);
            try
            {
                 //string url = ConfigurationManager.AppSettings["EndPoint"];
                 //var handler = new HttpClientHandler()
                 //{
                 //    Credentials = new NetworkCredential(user, password)
                 //};

                 using (var client = new HttpClient())
                 {
                     client.BaseAddress = new Uri(url);
                     client.DefaultRequestHeaders.Accept.Clear();
                     client.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/x-www-form-urlencoded"));
                     client.DefaultRequestHeaders.ConnectionClose = true;
                     var formContent = new StringContent(json, Encoding.UTF8, "application/json");
                    HttpResponseMessage responseMessage = await client.PostAsync("api/Ponche/CargaPonche/", formContent);
                    string result = await responseMessage.Content.ReadAsStringAsync();

                    Console.WriteLine(result);
                    //Console.Read();
                }
                return true;
            }


            catch (WebException ex)
            {
                Console.WriteLine(ex.Message + ex.StackTrace + ex.InnerException);
                
                return false;
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message + ex.StackTrace + ex.InnerException);
                ExReloj.Errores.Add("Error en conexion con Servicio Web" + " : "+ ex.Message);
                //Console.Read();
                return false;
            }

            finally
            {
                lstExRelojes.Add(ExReloj);
            }

           
        }


        private List<Atendancelogs> GetAttLogs(int v_SensorId, string v_Sn, ref EXRelojes exR)
        {
            List<Atendancelogs> lstAttLogs = new List<Atendancelogs>();

            int idwErrorCode = 0;

            string sdwEnrollNumber = "";
            //            int idwTMachineNumber = 0;
            //            int idwEMachineNumber = 0;
            int idwVerifyMode = 0;
            int idwInOutMode = 0;
            int idwYear = 0;
            int idwMonth = 0;
            int idwDay = 0;
            int idwHour = 0;
            int idwMinute = 0;
            int idwSecond = 0;
            int idwWorkcode = 0;
            int idwEnrollNumber = 0;
            int tmpUserId;
            string sTime = "";

            Cursor = Cursors.WaitCursor;
            axCZKEM1.EnableDevice(iMachineNumber, false);//disable the device
            if (axCZKEM1.ReadGeneralLogData(iMachineNumber))//read the records to the memory
            {
                while (axCZKEM1.GetGeneralLogDataStr(iMachineNumber, ref idwEnrollNumber, ref idwVerifyMode, ref idwInOutMode, ref sTime))//get the records from memory
                {
                    tmpUserId = GetUserId(idwEnrollNumber, ref exR);
                    if (tmpUserId > 0)
                    {
                        lstAttLogs.Add(new Atendancelogs { UserId = tmpUserId, BadgeNumber = idwEnrollNumber.ToString(), SensorId = v_SensorId.ToString(), SerialNumber = v_Sn, MemoryInfo = "NULL", UserExtFmt = 0, WorkCode = 0, CheckType = "I", VerifyCode = idwVerifyMode, CheckTime = Convert.ToDateTime(sTime) });
                    }
                }
                axCZKEM1.EnableDevice(iMachineNumber, true);//enable the device
                Cursor = Cursors.Default;
                if (lstAttLogs.Count == 0)
                {
                    sdwEnrollNumber = "";
                    //            int idwTMachineNumber = 0;
                    //            int idwEMachineNumber = 0;
                    idwVerifyMode = 0;
                    idwInOutMode = 0;
                    idwYear = 0;
                    idwMonth = 0;
                    idwDay = 0;
                    idwHour = 0;
                    idwMinute = 0;
                    idwSecond = 0;
                    idwWorkcode = 0;
                    idwEnrollNumber = 0;
                    tmpUserId = 0;
                    sTime = "";
                    Cursor = Cursors.WaitCursor;
                    axCZKEM1.EnableDevice(iMachineNumber, false);//disable the device
                    if (axCZKEM1.ReadGeneralLogData(iMachineNumber))//read all the attendance records to the memory
                    {
                        while (axCZKEM1.SSR_GetGeneralLogData(iMachineNumber, out sdwEnrollNumber, out idwVerifyMode,
                               out idwInOutMode, out idwYear, out idwMonth, out idwDay, out idwHour, out idwMinute, out idwSecond, ref idwWorkcode))//get records from the memory
                        {
                            tmpUserId = GetUserId(Convert.ToInt32(sdwEnrollNumber), ref exR);
                            if (tmpUserId > 0)
                            {
                                sTime = idwDay.ToString() + "/" + idwMonth.ToString() + "/" + idwYear.ToString() + " " + idwHour.ToString() + ":" + idwMinute.ToString() + ":" + idwSecond.ToString();
                                lstAttLogs.Add(new Atendancelogs { UserId = tmpUserId, BadgeNumber = sdwEnrollNumber, SensorId = v_SensorId.ToString(), WorkCode = idwWorkcode, SerialNumber = v_Sn, MemoryInfo = "NULL", UserExtFmt = 0, CheckType = "I", VerifyCode = idwVerifyMode, CheckTime = Convert.ToDateTime(sTime) });
                            }
                        }
                    }
                }

            }
            else
            {
                Cursor = Cursors.Default;
                axCZKEM1.GetLastError(ref idwErrorCode);

                if (idwErrorCode != 0)
                {
                    exR.Errores.Add("Leer data del reloj ha fallado, ErrorCode: " + idwErrorCode.ToString());
                }
                else
                {
                    exR.Errores.Add("No hay datos en el reloj!");
                }
            }
            axCZKEM1.EnableDevice(iMachineNumber, true);//enable the device
            Cursor = Cursors.Default;
            return lstAttLogs;
        }

        private void PrintList(List<Atendancelogs> dList)
        {
            foreach(Atendancelogs al in dList)
            {
                System.Diagnostics.Debug.WriteLine(al.BadgeNumber + " -> " + al.CheckTime.ToString());
            }
        }
        #endregion

        #region RemoteMDB
        private int GetUserId(int v_userid, ref EXRelojes exR)
        {
            int returnId = 0;
            string connectionString = "Provider=Microsoft.JET.OLEDB.4.0;Jet OLEDB:Database Password=A1BB807030C81;data source=" + ConfigurationManager.AppSettings["rutaDB"];
            OleDbConnection conn = new OleDbConnection(connectionString);
            string sql = "SELECT USERID FROM USERINFO WHERE Badgenumber = '" + v_userid + "';";
            OleDbCommand cmd = new OleDbCommand(sql, conn);

            // set up try si otras personas tienen abierta la DBASE
            conn.Open();
            try
            {
                reader = cmd.ExecuteReader();
            }
            catch (OleDbException ex)
            {
                exR.Errores.Add(ex.Message);
                //                continue; // add to excetions
            }
            DataTable mitabla = new DataTable();
            mitabla.Load(reader);
            foreach (DataRow row in mitabla.Rows)
            {
                returnId = (int)row[0];
            }
            conn.Close();
            return returnId;
        }

        private List<Machines> getMachines()
        {
            List<Machines> lstmac = new List<Machines>();


            string connectionString = "Provider=Microsoft.JET.OLEDB.4.0;Jet OLEDB:Database Password=A1BB807030C81;data source=" + ConfigurationManager.AppSettings["rutaDB"];
            OleDbConnection conn = new OleDbConnection(connectionString);
            //string sql = "SELECT MachineAlias, IP, Port, MachineNumber, SN FROM MACHINES;";
            string sql = "SELECT M.MachineAlias, M.IP, M.Port, M.MachineNumber, M.sn FROM MACHINES as M INNER JOIN RELOJPONCHES as R ON M.id = R.IDMachine;";
            OleDbCommand cmd = new OleDbCommand(sql, conn);
            // set up try si otras personas tienen abierta la DBASE
            try
            {
                conn.Open();
                reader = cmd.ExecuteReader();
                DataTable mitabla = new DataTable();
                mitabla.Load(reader);
                foreach (DataRow row in mitabla.Rows)
                {
                    lstmac.Add(new Machines { Name = (string)row[0], Ip = (string)row[1], Port = (int)row[2], MachineNumber = (int)row[3], SerialNumber = (string)row[4] });
                }
            }
            catch (OleDbException ex)
            {
                ConnectResults(ex.Message);
                this.Dispose();
            }
            conn.Close();

            return lstmac;
        }

        private bool FixErrors(Machines v_mc, ref EXRelojes exR)
        {
            exR.Errores.Clear();
            exR.status = true;
            // Esta función limpia los datos del día para evitar los duplicate index errors.
            bool insertStat = true;
            string connectionString = "Provider=Microsoft.JET.OLEDB.4.0;Jet OLEDB:Database Password=A1BB807030C81;data source=" + ConfigurationManager.AppSettings["rutaDB"];
            OleDbConnection conn = new OleDbConnection(connectionString);
            conn.Open();
            //                string sql = "DELETE * FROM CHECKINOUT WHERE (((CHECKINOUT.CHECKTIME)>=#9/25/2014# And (CHECKINOUT.CHECKTIME)<=#9/25/2014 11:59:59#));";
            string sql = "DELETE * FROM CHECKINOUT WHERE (((CHECKINOUT.CHECKTIME)>=@CHECKTIMEINI And (CHECKINOUT.CHECKTIME)<= @CHECKTIMEFIN) AND ((CHECKINOUT.SENSORID)=@SENSORID));";
            //                MessageBox.Show(sql);
            OleDbCommand cmd = new OleDbCommand(sql, conn);
            DateTime dateIni = new DateTime(DateTime.Today.Year, DateTime.Today.Month, DateTime.Today.Day, 00, 00, 00).AddDays(-400);
            DateTime dateFin = new DateTime(DateTime.Today.Year, DateTime.Today.Month, DateTime.Today.Day, DateTime.Now.Hour, DateTime.Now.Minute, DateTime.Now.Second);
            cmd.Parameters.Add(new OleDbParameter("@CHECKTIMEINI", dateIni));
            cmd.Parameters.Add(new OleDbParameter("@CHECKTIMEFIN", dateFin));
            cmd.Parameters.Add(new OleDbParameter("@SENSORID", v_mc.MachineNumber));
            // set up try si otras personas tienen abierta la DBASE      
            try
            {
                reader = cmd.ExecuteReader();     // try catch                                                    
            }
            catch (OleDbException ex)
            {
                exR.Errores.Add("Error al insertar datos en el BDatos Access: " + ex.Message);
                exR.status = false;
                insertStat = false;
            }
            conn.Close();
            return insertStat;
        }

        private bool InsertAttlogs(Machines v_mc, ref EXRelojes exR)
        {
            bool insertStat = true;
            string connectionString = "Provider=Microsoft.JET.OLEDB.4.0;Jet OLEDB:Database Password=A1BB807030C81;data source=" + ConfigurationManager.AppSettings["rutaDB"];
            OleDbConnection conn = new OleDbConnection(connectionString);
            conn.Open();
            foreach (Atendancelogs atl in v_mc.attList)
            {
                //                string sql = "INSERT INTO CHECKINOUT (USERID, CHECKTIME, CHECKTYPE, VERIFYCODE, SENSORID, Memoinfo, WorkCode, sn, UserExtFmt) VALUES (" + atl.UserId + ",'" + atl.CheckTime + "','" + atl.CheckType + "'," + atl.VerifyCode + ",'" + atl.SensorId + "','" + atl.MemoryInfo + "'," + atl.WorkCode + ",'" + atl.SerialNumber + "'," + atl.UserExtFmt + ");";
                string sql = "INSERT INTO CHECKINOUT (USERID, CHECKTIME, CHECKTYPE, VERIFYCODE, SENSORID, Memoinfo, WorkCode, sn, UserExtFmt) VALUES (@USERID, @CHECKTIME, @CHECKTYPE, @VERIFYCODE, @SENSORID, @Memoinfo, @WorkCode, @sn, @UserExtFmt)";
                //                MessageBox.Show(sql);
                OleDbCommand cmd = new OleDbCommand(sql, conn);
                cmd.Parameters.Add(new OleDbParameter("@USERID", atl.UserId));
                cmd.Parameters.Add(new OleDbParameter("@CHECKTIME", atl.CheckTime));
                cmd.Parameters.Add(new OleDbParameter("@CHECKTYPE", atl.CheckType));
                cmd.Parameters.Add(new OleDbParameter("@VERIFYCODE", atl.VerifyCode));
                cmd.Parameters.Add(new OleDbParameter("@SENSORID", atl.SensorId));
                cmd.Parameters.Add(new OleDbParameter("@Memoinfo", atl.MemoryInfo));
                cmd.Parameters.Add(new OleDbParameter("@WorkCode", atl.WorkCode));
                cmd.Parameters.Add(new OleDbParameter("@sn", atl.SerialNumber));
                cmd.Parameters.Add(new OleDbParameter("@UserExtFmt", atl.UserExtFmt));
                // set up try si otras personas tienen abierta la DBASE      
                try
                {
                    reader = cmd.ExecuteReader();     // try catch                                                    
                }
                catch (OleDbException ex)
                {
                    exR.Errores.Add("Error al insertar datos en el BDatos Access: " + ex.Message);
                    exR.status = false;
                    insertStat = false;
                    continue; // agregar exception a los logs                    
                }
            }
            conn.Close();
            return insertStat;
        }
        #endregion

        #region SendEmail
        private void ConnectResults(string erMessage)
        {
            StringBuilder sbMensaje = new StringBuilder();
            sbMensaje.Append("<!DOCTYPE html>");
            sbMensaje.Append("<html>");
            sbMensaje.Append("<head>");
            sbMensaje.Append("<style>");
            sbMensaje.Append("body,p,h1,h2,h3,h4,table,td,th,ul,ol,textarea,input{font-family:verdana,helvetica,arial,sans-serif;}");
            sbMensaje.Append("h1 {text-decoration:overline;}");
            sbMensaje.Append("h3 {text-decoration:underline;}");
            sbMensaje.Append("h3 {font-weight:bold;}");
            sbMensaje.Append("h4 {text-decoration:blink;}");
            sbMensaje.Append("ul {list-style-type:square; padding: 2px;}");
            sbMensaje.Append("</style>");
            sbMensaje.Append("</head>");
            sbMensaje.Append("<body>");
            sbMensaje.Append("<h1> Transferencia de relojes biometricos </h1>");
            sbMensaje.Append("<h3 style=\"color:red;\">No se ha encontrado la base de datos de relojes " + erMessage + "</h3>");
            sbMensaje.Append("</body>");
            sbMensaje.Append("</html>");

            MailMessage msg = new MailMessage();
            msg.From = new MailAddress(ConfigurationManager.AppSettings["fromNoReply"], ConfigurationManager.AppSettings["displayFrom"]);
            var emails = new List<string>(ConfigurationManager.AppSettings["EmailTo"].Split(new char[] { ',' }));
            foreach (var em in emails)
            {
                msg.To.Add(new MailAddress(em));
            }
            //msg.To.Add(new MailAddress("gmarte@ccn.net.do"));
            //            msg.Bcc.Add(ConfigurationManager.AppSettings["bccMail"]);
            msg.Subject = ConfigurationManager.AppSettings["SubjectEmail"];
            msg.Body = sbMensaje.ToString();

            msg.IsBodyHtml = true;

            System.Net.Mail.SmtpClient smtp = new SmtpClient();

            smtp.Host = ConfigurationManager.AppSettings["EmailServer"];
            smtp.Port = 587;
            smtp.Credentials = new System.Net.NetworkCredential(ConfigurationManager.AppSettings["fromNoReply"], ConfigurationManager.AppSettings["EmailPassword"]);

            smtp.EnableSsl = true;

            //try error handling del email
            try
            {
                smtp.Send(msg);
            }
            catch (Exception ex)
            {
                Console.WriteLine("Exception caught in CreateTestMessage1(): {0}",
                ex.ToString());
            }
        }
        private void SendResults(List<EXRelojes> lstEx)
        {

            StringBuilder sbMensaje = new StringBuilder();
            sbMensaje.Append("<!DOCTYPE html>");
            sbMensaje.Append("<html>");
            sbMensaje.Append("<head>");
            sbMensaje.Append("<style>");
            sbMensaje.Append("body,p,h1,h2,h3,h4,table,td,th,ul,ol,textarea,input{font-family:verdana,helvetica,arial,sans-serif;}");
            sbMensaje.Append("h1 {text-decoration:overline;}");
            sbMensaje.Append("h3 {text-decoration:underline;}");
            sbMensaje.Append("h3 {font-weight:bold;}");
            sbMensaje.Append("h4 {text-decoration:blink;}");
            sbMensaje.Append("ul {list-style-type:square; padding: 2px;}");
            sbMensaje.Append("</style>");
            sbMensaje.Append("</head>");
            sbMensaje.Append("<body>");
            sbMensaje.Append("<h1> Transferencia de relojes biometricos </h1>");
            sbMensaje.Append("<h3>Resumen de relojes</h3>");
            sbMensaje.Append("<ul>");
            foreach (EXRelojes exR in lstEx)
            {
                if (exR.status)
                {
                    sbMensaje.Append("<li> <label style=\"color:blue;\">" + exR.Name + "</label> <label> " + exR.attList.Count.ToString() + " </label> </li>");
                }
                else
                {
                    sbMensaje.Append("<li> <label style=\"font-weight:bold;color:red;\">" + exR.Name + "</label> </li>");
                }
            }
            sbMensaje.Append("</ul>");
            sbMensaje.Append("<h3>Detalle</h3>");
            foreach (EXRelojes exR in lstEx)
            {
                if (exR.Errores.Count == 0)
                {
                    continue;
                }
                sbMensaje.Append("<h4>" + exR.Name + ": Listado de errores</h4>");
                sbMensaje.Append("<ul>");
                foreach (string Message in exR.Errores)
                {
                    sbMensaje.Append("<li>" + Message + "</li>");
                }
                sbMensaje.Append("</ul>");
            }
            sbMensaje.Append("</body>");
            sbMensaje.Append("</html>");

            MailMessage msg = new MailMessage();
            msg.From = new MailAddress(ConfigurationManager.AppSettings["fromNoReply"], ConfigurationManager.AppSettings["displayFrom"]);
            var emails = new List<string>(ConfigurationManager.AppSettings["EmailTo"].Split(new char[] { ',' }));
            foreach (var em in emails)
            {
                msg.To.Add(new MailAddress(em));
            }
            //msg.To.Add(new MailAddress("gmarte@ccn.net.do"));
            //            msg.Bcc.Add(ConfigurationManager.AppSettings["bccMail"]);                        
            //            msg.Bcc.Add(ConfigurationManager.AppSettings["bccMail"]);
            msg.Subject = ConfigurationManager.AppSettings["SubjectEmail"];
            msg.Body = sbMensaje.ToString();

            msg.IsBodyHtml = true;

            System.Net.Mail.SmtpClient smtp = new SmtpClient();

            smtp.Host = ConfigurationManager.AppSettings["EmailServer"];
            smtp.Port = 587;
            smtp.Credentials = new System.Net.NetworkCredential(ConfigurationManager.AppSettings["fromNoReply"], ConfigurationManager.AppSettings["EmailPassword"]);

            smtp.EnableSsl = false;

            //try error handling del email
            try
            {
                smtp.Send(msg);
            }
            catch (Exception ex)
            {
                Console.WriteLine("Exception caught in CreateTestMessage1(): {0}",
                ex.ToString());
            }            
        }
        #endregion       
    }
    public class Machines
    {
        public string Name { get; set; }
        public string Ip { get; set; }
        public int Port { get; set; }
        public int MachineNumber { get; set; }
        public string SerialNumber { get; set; }
        public List<Atendancelogs> attList;
    }
    public class Atendancelogs
    {
        public int UserId { get; set; }
        public string BadgeNumber { get; set; }
        public DateTime CheckTime { get; set; }
        public string CheckType { get; set; } // I
        public int VerifyCode { get; set; } // 1
        public string SensorId { get; set; } // Machine Number
        public string MemoryInfo { get; set; } // null
        public int WorkCode { get; set; } // 0
        public string SerialNumber { get; set; } // # del Machine
        public int UserExtFmt { get; set; } // 0
    }
    public class EXRelojes : Machines
    {
        public bool status { get; set; }
        public List<string> Errores;
        public EXRelojes(Machines mac)
        {
            this.Name = mac.Name;
            this.Ip = mac.Ip;
            this.Port = mac.Port;
            this.Errores = new List<string>();
        }
    }

    public class SAPEventPairs
    {
        public string pernr { get; set; }
        public string ldate { get; set; }
        public string erdat { get; set; }
        public string ltime { get; set; }
        public string ertim { get; set; }
        public string satza { get; set; }
        public string terid { get; set; }

    }

    /// <summary>
    /// Helper class for invoking tasks with timeout. Overhead is 0,005 ms.
    /// </summary>
    /// <typeparam name="TResult">The type of the result.</typeparam>
    /*[Immutable]*/
    public sealed class WaitFor<TResult>
    {
        readonly TimeSpan _timeout;

        /// <summary>
        /// Initializes a new instance of the <see cref="WaitFor{T}"/> class, 
        /// using the specified timeout for all operations.
        /// </summary>
        /// <param name="timeout">The timeout.</param>
        public WaitFor(TimeSpan timeout)
        {
            _timeout = timeout;
        }

        /// <summary>
        /// Executes the spcified function within the current thread, aborting it
        /// if it does not complete within the specified timeout interval. 
        /// </summary>
        /// <param name="function">The function.</param>
        /// <returns>result of the function</returns>
        /// <remarks>
        /// The performance trick is that we do not interrupt the current
        /// running thread. Instead, we just create a watcher that will sleep
        /// until the originating thread terminates or until the timeout is
        /// elapsed.
        /// </remarks>
        /// <exception cref="ArgumentNullException">if function is null</exception>
        /// <exception cref="TimeoutException">if the function does not finish in time </exception>
        public TResult Run(Func<TResult> function)
        {
            if (function == null) throw new ArgumentNullException("function");


            var sync = new object();
            var isCompleted = false;

            WaitCallback watcher = obj =>
            {
                var watchedThread = obj as Thread;

                lock (sync)
                {
                    if (!isCompleted)
                    {
                        Monitor.Wait(sync, _timeout);
                    }
                }
                // CAUTION: the call to Abort() can be blocking in rare situations
                // http://msdn.microsoft.com/en-us/library/ty8d3wta.aspx
                // Hence, it should not be called with the 'lock' as it could deadlock
                // with the 'finally' block below.

                if (!isCompleted)
                {
                    watchedThread.Abort();
                }
            };

            try
            {
                ThreadPool.QueueUserWorkItem(watcher, Thread.CurrentThread);
                return function();
            }
            catch (ThreadAbortException)
            {
                // This is our own exception.
                Thread.ResetAbort();

                throw new TimeoutException(string.Format("The operation has timed out after {0}.", _timeout));
            }
            finally
            {
                lock (sync)
                {
                    isCompleted = true;
                    Monitor.Pulse(sync);
                }
            }
        }

        /// <summary>
        /// Executes the spcified function within the current thread, aborting it
        /// if it does not complete within the specified timeout interval.
        /// </summary>
        /// <param name="timeout">The timeout.</param>
        /// <param name="function">The function.</param>
        /// <returns>result of the function</returns>
        /// <remarks>
        /// The performance trick is that we do not interrupt the current
        /// running thread. Instead, we just create a watcher that will sleep
        /// until the originating thread terminates or until the timeout is
        /// elapsed.
        /// </remarks>
        /// <exception cref="ArgumentNullException">if function is null</exception>
        /// <exception cref="TimeoutException">if the function does not finish in time </exception>
        public static TResult Run(TimeSpan timeout, Func<TResult> function)
        {
            return new WaitFor<TResult>(timeout).Run(function);
        }

        

    }
}
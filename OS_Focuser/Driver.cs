//tabs=4
// --------------------------------------------------------------------------------
// TODO fill in this information for your driver, then remove this line!
//
// ASCOM Focuser driver for OS_Focuser
//
// Description:	Lorem ipsum dolor sit amet, consetetur sadipscing elitr, sed diam 
//				nonumy eirmod tempor invidunt ut labore et dolore magna aliquyam 
//				erat, sed diam voluptua. At vero eos et accusam et justo duo 
//				dolores et ea rebum. Stet clita kasd gubergren, no sea takimata 
//				sanctus est Lorem ipsum dolor sit amet.
//
// Implements:	ASCOM Focuser interface version: <To be completed by driver developer>
// Author:		(XXX) Your N. Here <your@email.here>
//
// Edit Log:
//
// Date			Who	Vers	Description
// -----------	---	-----	-------------------------------------------------------
// dd-mmm-yyyy	XXX	6.0.0	Initial edit, created from ASCOM driver template
// --------------------------------------------------------------------------------
//


// This is used to define code in the template that is specific to one class implementation
// unused code canbe deleted and this definition removed.
#define Focuser

using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Text;
using System.Runtime.InteropServices;

using ASCOM;
using ASCOM.Astrometry;
using ASCOM.Astrometry.AstroUtils;
using ASCOM.Utilities;
using ASCOM.DeviceInterface;
using System.Globalization;
using System.Collections;
using System.Linq;
using System.Threading.Tasks;
using System.IO;
using System.IO.Ports;
using System.Threading;

namespace ASCOM.OS_Focuser
{
    //
    // Your driver's DeviceID is ASCOM.OS_Focuser.Focuser
    //
    // The Guid attribute sets the CLSID for ASCOM.OS_Focuser.Focuser
    // The ClassInterface/None addribute prevents an empty interface called
    // _OS_Focuser from being created and used as the [default] interface
    //
    // TODO Replace the not implemented exceptions with code to implement the function or
    // throw the appropriate ASCOM exception.
    //

    /// <summary>
    /// ASCOM Focuser Driver for OS_Focuser.
    /// </summary>
    [Guid("6d4afa50-a499-41c4-93fd-cf19a78ec98f")]
    [ClassInterface(ClassInterfaceType.None)]
    public class Focuser : IFocuserV2
    {
        /// <summary>
        /// ASCOM DeviceID (COM ProgID) for this driver.
        /// The DeviceID is used by ASCOM applications to load the driver at runtime.
        /// </summary>
        internal static string driverID = "ASCOM.OS_Focuser.Focuser";
        // TODO Change the descriptive string for your driver then remove this line
        /// <summary>
        /// Driver description that displays in the ASCOM Chooser.
        /// </summary>
        private static string driverDescription = "ATC02 Focuser OS";

        internal static string comPortProfileName = "COM Port"; // Constants used for Profile persistence
        internal static string comPortDefault = "COM1";
        internal static string traceStateProfileName = "Trace Level";
        internal static string traceStateDefault = "false";

        internal static string comPort; // Variables to hold the currrent device configuration

        public static ASCOM.Utilities.Serial port;

        public static string OptBFL;
        public static string deltaMax;
        public static string _curBfl;
        public static bool moving = false;
        private double OptimalBFLValue;
        private double MaxBFLDelta;
        private float bfl;
        private double curBfl;
        private double requirement;

        Values lib;

        /// <summary>
        /// Private variable to hold the connected state
        /// </summary>
        private bool connectedState;

        /// <summary>
        /// Private variable to hold an ASCOM Utilities object
        /// </summary>
        private Util utilities;

        /// <summary>
        /// Private variable to hold an ASCOM AstroUtilities object to provide the Range method
        /// </summary>
        private AstroUtils astroUtilities;

        /// <summary>
        /// Variable to hold the trace logger object (creates a diagnostic log file with information that you specify)
        /// </summary>
        internal static TraceLogger tl;

        /// <summary>
        /// Initializes a new instance of the <see cref="OS_Focuser"/> class.
        /// Must be public for COM registration.
        /// </summary>
        public Focuser()
        {
            tl = new TraceLogger("", "OS_Focuser");
            ReadProfile(); // Read device configuration from the ASCOM Profile store

            tl.LogMessage("Focuser", "Starting initialisation");

            connectedState = false; // Initialise connected to false
            utilities = new Util(); //Initialise util object
            astroUtilities = new AstroUtils(); // Initialise astro utilities object
            //TODO: Implement your additional construction here

            tl.LogMessage("Focuser", "Completed initialisation");
        }


        //
        // PUBLIC COM INTERFACE IFocuserV2 IMPLEMENTATION
        //

        #region Common properties and methods.

        /// <summary>
        /// Displays the Setup Dialog form.
        /// If the user clicks the OK button to dismiss the form, then
        /// the new settings are saved, otherwise the old values are reloaded.
        /// THIS IS THE ONLY PLACE WHERE SHOWING USER INTERFACE IS ALLOWED!
        /// </summary>
        public void SetupDialog()
        {
            // consider only showing the setup dialog if not connected
            // or call a different dialog if connected
            if (IsConnected)
                System.Windows.Forms.MessageBox.Show("Already connected, just press OK");

            using (SetupDialogForm F = new SetupDialogForm())
            {
                var result = F.ShowDialog();
                if (result == System.Windows.Forms.DialogResult.OK)
                {
                    WriteProfile(); // Persist device configuration values to the ASCOM Profile store
                }
            }
        }

        public ArrayList SupportedActions
        {
            get
            {
                tl.LogMessage("SupportedActions Get", "Returning empty arraylist");
                return new ArrayList();
            }
        }

        public string Action(string actionName, string actionParameters)
        {
            LogMessage("", "Action {0}, parameters {1} not implemented", actionName, actionParameters);
            throw new ASCOM.ActionNotImplementedException("Action " + actionName + " is not implemented by this driver");
        }

        public void CommandBlind(string command, bool raw)
        {
            CheckConnected("CommandBlind");
            // Call CommandString and return as soon as it finishes
            this.CommandString(command, raw);
            // or
            throw new ASCOM.MethodNotImplementedException("CommandBlind");
            // DO NOT have both these sections!  One or the other
        }

        public bool CommandBool(string command, bool raw)
        {
            CheckConnected("CommandBool");
            string ret = CommandString(command, raw);
            // TODO decode the return string and return true or false
            // or
            throw new ASCOM.MethodNotImplementedException("CommandBool");
            // DO NOT have both these sections!  One or the other
        }

        public string CommandString(string command, bool raw)
        {
            CheckConnected("CommandString");
            // it's a good idea to put all the low level communication with the device here,
            // then all communication calls this function
            // you need something to ensure that only one command is in progress at a time

            throw new ASCOM.MethodNotImplementedException("CommandString");
        }

        public void Dispose()
        {
            // Clean up the tracelogger and util objects
            tl.Enabled = false;
            tl.Dispose();
            tl = null;
            utilities.Dispose();
            utilities = null;
            astroUtilities.Dispose();
            astroUtilities = null;
        }

        public bool Connected
        {
            get
            {
                LogMessage("Connected", "Get {0}", IsConnected);
                return IsConnected;
            }
            set
            {
                tl.LogMessage("Connected", "Set {0}", value);
                if (value == IsConnected)
                    return;

                if (value)
                {
                    port = new ASCOM.Utilities.Serial();
                    connectedState = true;
                    LogMessage("Connected Set", "Connecting to port {0}", comPort);
                    // TODO connect to the device

                    Connect(comPort);

                }
                else
                {
                    connectedState = false;
                    LogMessage("Connected Set", "Disconnecting from port {0}", comPort);
                    // TODO disconnect from the device

                    Disconnect();
                }
            }
        }


        public void Connect(string _comPort)
        {
            if (port.Connected == true)
            {
                port.Connected = false;
            }
            port.PortName = _comPort;
            port.Speed = (SerialSpeed)2400;
            port.StopBits = SerialStopBits.One;
            port.Parity = SerialParity.None;
            port.DataBits = 8;
            port.DTREnable = false;
            port.RTSEnable = false;
            port.Handshake = SerialHandshake.None;
            port.Connected = true;
            int i = 1;
            do
            {
                currentBFL();
                Thread.Sleep(100);
                i++;
            }
            while (i >= 5);
        }

        //new public System.Timers.Timer timer;

        public static void Disconnect()
        {
            port.Connected = false;
        }

        public string Description
        {
            // TODO customise this device description
            get
            {
                tl.LogMessage("Description Get", driverDescription);
                return driverDescription;
            }
        }

        public string DriverInfo
        {
            get
            {
                Version version = System.Reflection.Assembly.GetExecutingAssembly().GetName().Version;
                // TODO customise this driver description
                string driverInfo = "Information about the driver itself. Version: " + String.Format(CultureInfo.InvariantCulture, "{0}.{1}", version.Major, version.Minor);
                tl.LogMessage("DriverInfo Get", driverInfo);
                return driverInfo;
            }
        }

        public string DriverVersion
        {
            get
            {
                Version version = System.Reflection.Assembly.GetExecutingAssembly().GetName().Version;
                string driverVersion = String.Format(CultureInfo.InvariantCulture, "{0}.{1}", version.Major, version.Minor);
                tl.LogMessage("DriverVersion Get", driverVersion);
                return driverVersion;
            }
        }

        public short InterfaceVersion
        {
            // set by the driver wizard
            get
            {
                LogMessage("InterfaceVersion Get", "2");
                return Convert.ToInt16("2");
            }
        }

        public string Name
        {
            get
            {
                string name = "Officina Stellare ATC02 Focuser ASCOM Driver";
                tl.LogMessage("Name Get", name);
                return name;
            }
        }

        #endregion


        #region IFocuser Implementation

        private int focuserPosition = 0; // Class level variable to hold the current focuser position
        private const int focuserSteps = 10000; // Used to be 100000 in the OS program

        public bool Absolute
        {
            get
            {
                tl.LogMessage("Absolute Get", true.ToString());
                return true; // This is an absolute focuser
            }
        }

        public void Halt()
        {
            tl.LogMessage("Halt", "Not implemented");
            throw new ASCOM.MethodNotImplementedException("Halt");
        }

        public bool IsMoving
        {
            get
            {
                tl.LogMessage("IsMoving Get", false.ToString());
                return false; // This focuser always moves instantaneously so no need for IsMoving ever to be True
            }
        }

        public bool Link
        {
            get
            {
                tl.LogMessage("Link Get", this.Connected.ToString());
                return this.Connected; // Direct function to the connected method, the Link method is just here for backwards compatibility
            }
            set
            {
                tl.LogMessage("Link Set", value.ToString());
                this.Connected = value; // Direct function to the connected method, the Link method is just here for backwards compatibility
            }
        }


        private float _maxStep;

        public int MaxStep
        {
            get
            {
                ExecuteCommand("READSETT");
                Thread.Sleep(3000);
                try
                {
                    string receivedRawData = port.Receive();

                    string[] parsedData = receivedRawData.Split(new[] { '\n', '\r' });

                    foreach (string s in parsedData)
                    {
                        if (s.StartsWith("OPBF"))
                        {
                            OptBFL = s.Remove(0, 4).Trim();

                            double _OptBFL = double.Parse(OptBFL);
                            OptimalBFLValue = _OptBFL;
                        }

                        else if (s.StartsWith("DELTAMAX"))
                        {
                            deltaMax = s.Remove(0, 8).Trim();

                            double _deltaMax = double.Parse(deltaMax);
                            MaxBFLDelta = _deltaMax;
                        }
                    }
                    _maxStep = toASCOMPosition(2 * (float)MaxBFLDelta);
                    return (int)_maxStep;
                }
                catch
                {
                    throw new ASCOM.NotConnectedException();
                }
            }
            set
            {
                if((value > 0) && (value <= 10000))
                {
                    _maxStep = value;
                }
            }
        }

        public int MaxIncrement
        {
            get
            {
                tl.LogMessage("MaxIncrement Get", focuserSteps.ToString());
                return (int)_maxStep; // Maximum change in one move
            }
        }

        public void Move(int position)
        {
            if (position >= 0 && position <= 10000)
            {
                focuserPosition = position; // Set the focuser position
                tl.LogMessage("Move", position.ToString());


                float pos = toATCPosition(position) + (float)OptimalBFLValue - (float)MaxBFLDelta;                 // Ensures that M2 stays in its range
                ExecuteCommand("BFL ", pos);
                moving = true;
                bfl = (float)toATCPosition(position);
                Thread.Sleep(100);

                do
                {
                    currentBFL();
                }
                while (curBfl != pos);
            }
            else
            {
                throw new ASCOM.InvalidValueException("Choose a position between 0 and 10000.");
            }
        }

        public int Position
        {
            get
            {
                currentBFL();
                    return toASCOMPosition(curBfl);
            }
        }



        #region Specified Functions


        public static void ExecuteCommand(string command)
        {
            try
            {
                byte[] ende = new byte[1];
                ende = new byte[] { 0x0D };
                
                port.Transmit(command);
                Thread.Sleep(100);
                port.TransmitBinary(ende);
                tl.LogMessage("Sent Command", command);
            }
            catch
            {
                throw new ASCOM.NotConnectedException();
            }
        }

        public void ExecuteCommand(string command, float val)
        {
            try
            {
                string data = command + val.ToString("000.00").Replace(',', '.');       // Microcontroller reads value in [mm] format. But that is not the real scale. 1 step = 1,55 µm * 7
                port.Transmit(data);
                tl.LogMessage("Sent Command", data);
            }
            catch
            {
                throw new ASCOM.NotConnectedException();
            }
        }


        public void currentBFL()
        {
            ExecuteCommand("UPDATEPC");
            Thread.Sleep(3000);
            try
            {

                string receivedRawData = port.Receive();
                string[] parsedData = receivedRawData.Split('\n', '\r');

                foreach (var s in parsedData)
                {
                    if (s.Contains("BFL"))
                    {
                        _curBfl = s.Substring(4, 6);
                        curBfl = double.Parse(_curBfl) - (float)OptimalBFLValue + (float)MaxBFLDelta;

                    }
                }
            }
            catch
            {
                curBfl = 99999;
            }
        }


        /// <summary>
        /// Reads data and sets OpimalBFL and DeltaMax to their values.
        /// </summary>
        //public void readATC()
        //{
        //    ExecuteCommand("READSETT");
        //    Thread.Sleep(3000);
        //    string receivedRawData = port.Receive();

        //    string[] parsedData = receivedRawData.Split('\n', '\r');

        //    foreach (var s in parsedData)
        //    {
        //        if (s.Contains("OPBF"))
        //        {
        //            OptBFL = s.Substring(4, 6);

        //            double _OptBFL = double.Parse(OptBFL);
        //            OptimalBFLValue = _OptBFL;
        //        }

        //        else if (s.Contains("DELTAMAX"))
        //        {
        //            deltaMax = s.Substring(8, 2);

        //            double _deltaMax = double.Parse(deltaMax);
        //            MaxBFLDelta = _deltaMax;
        //        }
        //    }
        //    _maxStep = toASCOMPosition(2 * (float)MaxBFLDelta);
        //}


        /// <summary>
        /// Converts movement from ASCOM [steps] to ATC [mm].
        /// </summary>
        /// <param name="ATCposition"></param>
        /// <returns></returns>
        public int toASCOMPosition(double ATCposition)
        {
            try
            {
                return Convert.ToInt16(ATCposition * 100);
            }
            catch
            {
                throw new Exception("ATCPosition: " + ATCposition.ToString());
            }
        }


        /// <summary>
        /// Converts movement from ATC[mm] to ASCOM [steps]
        /// </summary>
        /// <param name="ASCOMPosition"></param>
        /// <returns></returns>
        public float toATCPosition(int ASCOMPosition)
        {
            return ((float)ASCOMPosition) / 100;
        }

        #endregion



        public double StepSize
        {
            get
            {
                tl.LogMessage("StepSize Get", "10");
                return 10;
                //throw new ASCOM.PropertyNotImplementedException("StepSize", false);
            }
        }
       

        public bool TempComp
        {
            get
            {
                tl.LogMessage("TempComp Get", false.ToString());
                return false;
            }
            set
            {
                tl.LogMessage("TempComp Set", "Not implemented");
                throw new ASCOM.PropertyNotImplementedException("TempComp", false);
            }
        }

        public bool TempCompAvailable
        {
            get
            {
                tl.LogMessage("TempCompAvailable Get", false.ToString());
                return false; // Temperature compensation is not available in this driver
            }
        }

        private double _temp;

        public double Temperature
        {
            get
            {
                ExecuteCommand("UPDATEPC");
                Thread.Sleep(4000);
                try
                {

                    string receivedRawData = port.Receive();
                    string[] parsedData = receivedRawData.Split('\n', '\r');

                    foreach (var s in parsedData)
                    {
                        if (s.Contains("AMBTE"))
                        {
                            string temp = s.Substring(6, 4);
                            _temp = double.Parse(temp);
                            tl.LogMessage("Temperatur Get", temp);

                        }
                    }
                    return _temp;
                }
                catch
                {
                    tl.LogMessage("Temperature Get", "Not implemented");
                    throw new ASCOM.PropertyNotImplementedException("Temperature", false);
                }
            }
        }

        #endregion


        #region Private properties and methods
        // here are some useful properties and methods that can be used as required
        // to help with driver development

        #region ASCOM Registration

        // Register or unregister driver for ASCOM. This is harmless if already
        // registered or unregistered. 
        //
        /// <summary>
        /// Register or unregister the driver with the ASCOM Platform.
        /// This is harmless if the driver is already registered/unregistered.
        /// </summary>
        /// <param name="bRegister">If <c>true</c>, registers the driver, otherwise unregisters it.</param>
        private static void RegUnregASCOM(bool bRegister)
        {
            using (var P = new ASCOM.Utilities.Profile())
            {
                P.DeviceType = "Focuser";
                if (bRegister)
                {
                    P.Register(driverID, driverDescription);
                }
                else
                {
                    P.Unregister(driverID);
                }
            }
        }

        /// <summary>
        /// This function registers the driver with the ASCOM Chooser and
        /// is called automatically whenever this class is registered for COM Interop.
        /// </summary>
        /// <param name="t">Type of the class being registered, not used.</param>
        /// <remarks>
        /// This method typically runs in two distinct situations:
        /// <list type="numbered">
        /// <item>
        /// In Visual Studio, when the project is successfully built.
        /// For this to work correctly, the option <c>Register for COM Interop</c>
        /// must be enabled in the project settings.
        /// </item>
        /// <item>During setup, when the installer registers the assembly for COM Interop.</item>
        /// </list>
        /// This technique should mean that it is never necessary to manually register a driver with ASCOM.
        /// </remarks>
        [ComRegisterFunction]
        public static void RegisterASCOM(Type t)
        {
            RegUnregASCOM(true);
        }

        /// <summary>
        /// This function unregisters the driver from the ASCOM Chooser and
        /// is called automatically whenever this class is unregistered from COM Interop.
        /// </summary>
        /// <param name="t">Type of the class being registered, not used.</param>
        /// <remarks>
        /// This method typically runs in two distinct situations:
        /// <list type="numbered">
        /// <item>
        /// In Visual Studio, when the project is cleaned or prior to rebuilding.
        /// For this to work correctly, the option <c>Register for COM Interop</c>
        /// must be enabled in the project settings.
        /// </item>
        /// <item>During uninstall, when the installer unregisters the assembly from COM Interop.</item>
        /// </list>
        /// This technique should mean that it is never necessary to manually unregister a driver from ASCOM.
        /// </remarks>
        [ComUnregisterFunction]
        public static void UnregisterASCOM(Type t)
        {
            RegUnregASCOM(false);
        }

        #endregion

        /// <summary>
        /// Returns true if there is a valid connection to the driver hardware
        /// </summary>
        private bool IsConnected
        {
            get
            {
                // TODO check that the driver hardware connection exists and is connected to the hardware
                return connectedState;
            }
        }

        /// <summary>
        /// Use this function to throw an exception if we aren't connected to the hardware
        /// </summary>
        /// <param name="message"></param>
        private void CheckConnected(string message)
        {
            if (!IsConnected)
            {
                throw new ASCOM.NotConnectedException(message);
            }
        }

        /// <summary>
        /// Read the device configuration from the ASCOM Profile store
        /// </summary>
        internal void ReadProfile()
        {
            using (Profile driverProfile = new Profile())
            {
                driverProfile.DeviceType = "Focuser";
                tl.Enabled = Convert.ToBoolean(driverProfile.GetValue(driverID, traceStateProfileName, string.Empty, traceStateDefault));
                comPort = driverProfile.GetValue(driverID, comPortProfileName, string.Empty, comPortDefault);
            }
        }

        /// <summary>
        /// Write the device configuration to the  ASCOM  Profile store
        /// </summary>
        internal void WriteProfile()
        {
            using (Profile driverProfile = new Profile())
            {
                driverProfile.DeviceType = "Focuser";
                driverProfile.WriteValue(driverID, traceStateProfileName, tl.Enabled.ToString());
                driverProfile.WriteValue(driverID, comPortProfileName, comPort.ToString());
            }
        }

        /// <summary>
        /// Log helper function that takes formatted strings and arguments
        /// </summary>
        /// <param name="identifier"></param>
        /// <param name="message"></param>
        /// <param name="args"></param>
        internal static void LogMessage(string identifier, string message, params object[] args)
        {
            var msg = string.Format(message, args);
            tl.LogMessage(identifier, msg);
        }
        #endregion
    }// end of class
}// end of namespace

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Reflection;
using System.Diagnostics;
using System.Runtime.InteropServices;

using ASCOM;
using ASCOM.Astrometry;
using ASCOM.Astrometry.AstroUtils;
using ASCOM.Utilities;
using ASCOM.DeviceInterface;
using System.Globalization;
using System.Collections;
using System.Threading.Tasks;
using System.IO;
using System.IO.Ports;
using System.Threading;

namespace ASCOM.OS_Focuser
{
    public class ATC02CommandReply
    {
        public enum ReplyType
        {
            ParametrizedReply,
            ReplyCommand
        }

        public ReplyType replyType { get; set; }
        string lastReceivedToken;

        public ATC02CommandReply(ReplyType type, string lastRecvToken)
        {
            replyType = type;
            lastReceivedToken = lastRecvToken;
        }

        public bool isCommandOver(string receivedString)
        {
            if (receivedString.StartsWith(lastReceivedToken))
                return true;
            return false;
        }
    }


    public class Values
    {
        public float OptimalBFLValue { get; set; }
        public float MaxBFLDelta { get; set; }
    }

    public class ATC02Command
    {
        public string commandString { get; private set; }
        //public bool isCommandComplete { get; private set; }
        public object commandArgs { get; set; }
        public ATC02CommandReply replyType { get; private set; }
        public int commandTimeout { get; private set; }

        #region "Constructors"
        public ATC02Command(String cmdString)
            : this(cmdString, null)
        {

        }

        public ATC02Command(ATC02Command c)
        {
            this.commandArgs = c.commandArgs;
            this.commandString = c.commandString;
            this.replyType = c.replyType;
            this.commandTimeout = c.commandTimeout;
        }

        public ATC02Command(String cmdString, int Timeout)
            : this(cmdString, "", ATC02CommandReply.ReplyType.ReplyCommand, cmdString, Timeout)
        {

        }

        public ATC02Command(String cmdString, Object commandArgs)
            : this(cmdString, commandArgs, ATC02CommandReply.ReplyType.ReplyCommand)
        {
        }

        public ATC02Command(string cmdString, Object commandArgs, ATC02CommandReply.ReplyType reply)
            : this(cmdString, commandArgs, reply, cmdString, 2500)
        {
            if (reply != ATC02CommandReply.ReplyType.ReplyCommand)
                throw (new ATCLibExceptions("ReplyType.ParametrizedCommand need to specify a last token!"));
        }

        public ATC02Command(string cmdString, ATC02CommandReply.ReplyType reply)
            : this(cmdString, null, reply, cmdString, 2500)
        {
            if (reply != ATC02CommandReply.ReplyType.ReplyCommand)
                throw (new ATCLibExceptions("ReplyType.ParametrizedCommand need to specify a last token!"));
        }

        public ATC02Command(string cmdString, ATC02CommandReply.ReplyType reply, string lastToken)
            : this(cmdString, null, reply, lastToken, 2500)
        {
        }

        public ATC02Command(String cmdString, Object commandArgs, ATC02CommandReply.ReplyType reply, string lastReceiveToken, int commandTimeout)
        {
            commandString = cmdString;
            this.commandArgs = commandArgs;
            this.replyType = new ATC02CommandReply(reply, lastReceiveToken);
            this.commandTimeout = commandTimeout;
        }

        #endregion  

    }


    //public class ATCCommands
    //{
    //    public static bool CheckForReply(string receivedData, string expectedReplyString, ref object var)
    //    {

    //    }
    //}
    public class ATCLibExceptions : Exception
    {
        public string message { get; private set; }
        public ATCLibExceptions(string msg)
            : base()
        {
            message = msg;
        }

    }
}

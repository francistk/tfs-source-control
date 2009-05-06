using System;
using System.Collections.Generic;
using System.Text;
using System.Timers;
using DataAccess;
using System.Data;
using System.Text.RegularExpressions;



using XihSolutions.DotMSN;

namespace MSNRobot
{
    class Program
    {
        static Messenger messenger = new Messenger();
        static Timer timer;

        static void Main(string[] args)
        {
            string userName = System.Configuration.ConfigurationManager.AppSettings["MsnAccount"];
            string passWord = System.Configuration.ConfigurationManager.AppSettings["MsnPassword"];
            string welcome = System.Configuration.ConfigurationManager.AppSettings["MsnPassword"];
            //Console.WriteLine(System.Configuration.ConfigurationManager.AppSettings["MsnRobotDbPath"]);
            //Console.Read();
            // by default this example will emulate the official microsoft windows messenger client
            messenger.Credentials.ClientID = "msmsgs@msnmsgr.com";
            messenger.Credentials.ClientCode = "Q1P7W2E4J9R8U3S5";
            messenger.Nameserver.PingAnswer += new PingAnswerEventHandler(Nameserver_PingAnswer);
            messenger.NameserverProcessor.ConnectionEstablished += new EventHandler(NameserverProcessor_ConnectionEstablished);
            messenger.Nameserver.SignedIn += new EventHandler(Nameserver_SignedIn);
            messenger.Nameserver.SignedOff += new SignedOffEventHandler(Nameserver_SignedOff);
            messenger.NameserverProcessor.ConnectionException += new XihSolutions.DotMSN.Core.ProcessorExceptionEventHandler(NameserverProcessor_ConnectionException);
            messenger.Nameserver.ExceptionOccurred += new XihSolutions.DotMSN.Core.HandlerExceptionEventHandler(Nameserver_ExceptionOccurred);
            messenger.Nameserver.AuthenticationError += new XihSolutions.DotMSN.Core.HandlerExceptionEventHandler(Nameserver_AuthenticationError);
            messenger.Nameserver.ServerErrorReceived += new XihSolutions.DotMSN.Core.ErrorReceivedEventHandler(Nameserver_ServerErrorReceived);
            messenger.ConversationCreated += new ConversationCreatedEventHandler(messenger_ConversationCreated);
            messenger.TransferInvitationReceived += new XihSolutions.DotMSN.DataTransfer.MSNSLPInvitationReceivedEventHandler(messenger_TransferInvitationReceived);
            messenger.Nameserver.ReverseAdded += new ContactChangedEventHandler(Nameserver_ReverseAdded);

            messenger.Credentials.Account = userName;
            messenger.Credentials.Password = passWord;

            timer = new Timer(30000);
            timer.Enabled = true;            
            timer.Elapsed += new ElapsedEventHandler(timer_Elapsed);
            timer.Stop();
            WL("Connecting to server");
            try
            {
                messenger.Connect();
            }
            catch (Exception e)
            {

                WL(e.Message);
            }
            RL();

        }
        static void timer_Elapsed(object sender, ElapsedEventArgs e)
        {
            SendPing();
        }
        static private void ReConnect()
        {
            try
            {
                timer.Stop();
                WL("Disconnecting");
                messenger.Disconnect();
                WL("Reconnecting");
                messenger.Connect();
            }
            catch (Exception)
            {
                System.Threading.Thread.Sleep(3000);
                ReConnect();
            }

        }

        static void Nameserver_ReverseAdded(object sender, ContactEventArgs e)
        {
            Console.WriteLine("Nameserver_ReverseAdded");
            messenger.Nameserver.AddContactToList(e.Contact, MSNLists.AllowedList);
            e.Contact.OnAllowedList = true;
            e.Contact.OnForwardList = true;
        }

        static void messenger_TransferInvitationReceived(object sender, XihSolutions.DotMSN.DataTransfer.MSNSLPInvitationEventArgs e)
        {
            WL("messenger_TransferInvitationReceived");
        }

        static void messenger_ConversationCreated(object sender, ConversationCreatedEventArgs e)
        {
            e.Conversation.Switchboard.TextMessageReceived += new TextMessageReceivedEventHandler(Switchboard_TextMessageReceived);
            e.Conversation.Switchboard.ContactJoined += new ContactChangedEventHandler(Switchboard_ContactJoined);
            e.Conversation.Switchboard.ContactLeft += new ContactChangedEventHandler(Switchboard_ContactLeft);
            e.Conversation.Switchboard.SessionClosed += new SBChangedEventHandler(Switchboard_SessionClosed);
            e.Conversation.Switchboard.UserTyping += new UserTypingEventHandler(Switchboard_UserTyping);
        }

        static void Switchboard_UserTyping(object sender, ContactEventArgs e)
        {
            WL(e.Contact.Name + "is typing");
            //发送打字状态
            ((XihSolutions.DotMSN.SBMessageHandler)sender).SendTypingMessage();
        }

        static void Switchboard_SessionClosed(object sender, EventArgs e)
        {
            WL("session closed");
        }

        static void Switchboard_ContactLeft(object sender, ContactEventArgs e)
        {
            WL("{0} left", e.Contact.Name);
        }

        static void Switchboard_ContactJoined(object sender, ContactEventArgs e)
        {
            WL("{0} joined.", e.Contact.Name);

            //jmq add
            XihSolutions.DotMSN.SBMessageHandler handler = (XihSolutions.DotMSN.SBMessageHandler)sender;
            string welcome = System.Configuration.ConfigurationManager.AppSettings["MsnRobotWelcome"];

            handler.SendTextMessage(new TextMessage(welcome));

        }

        static void Switchboard_TextMessageReceived(object sender, TextMessageEventArgs e)
        {
            WL("{0} says : {1}", e.Sender, e.Message);
            XihSolutions.DotMSN.SBMessageHandler handler = (XihSolutions.DotMSN.SBMessageHandler)sender;
            //handler.SendTextMessage(new TextMessage(string.Format("您好，你跟我说：{0}", e.Message.Text)));

            string returnMessage = GetResponse(e.Message.Text);
            handler.SendTextMessage(new TextMessage("->"+returnMessage));
        }

        static void Nameserver_ServerErrorReceived(object sender, MSNErrorEventArgs e)
        {
            WL("Nameserver_ServerErrorReceived");
        }

        static void Nameserver_AuthenticationError(object sender, ExceptionEventArgs e)
        {
            WL("Nameserver_AuthenticationError");
        }

        static void Nameserver_ExceptionOccurred(object sender, ExceptionEventArgs e)
        {
            WL("Nameserver_ExceptionOccurred");
        }

        static void NameserverProcessor_ConnectionException(object sender, ExceptionEventArgs e)
        {
            WL("NameserverProcessor_ConnectionException");
        }

        static void Nameserver_SignedOff(object sender, SignedOffEventArgs e)
        {
            WL("Nameserver_SignedOff");
            ReConnect();
        }

        static void Nameserver_SignedIn(object sender, EventArgs e)
        {
            messenger.Owner.Status = PresenceStatus.Online;
            messenger.Owner.NotifyPrivacy = NotifyPrivacy.AutomaticAdd;

            //messenger.Owner.DisplayImage.Image = System.Drawing.Image.FromFile(@"c:\1235\jmq.jpg"); 

            WL("Nameserver_SignedIn");
            WL("正在加载联系人...");
            foreach (Contact contact in messenger.ContactList.Forward)
            {
                WL(contact.Name);
            }
            foreach (Contact contact in messenger.ContactList.Allowed)
            {
                WL(contact.Name);
            }
            foreach (Contact contact in messenger.ContactList.BlockedList)
            {
                WL(contact.Name);
            }

            WL("联系人加载完成...");
            messenger.Nameserver.SetPrivacyMode(PrivacyMode.AllExceptBlocked);
            timer.Start();
            WL("机器人准备就绪...");
        }
        static void Nameserver_PingAnswer(object sender, PingAnswerEventArgs e)
        {
            WL(e.SecondsToWait.ToString());
            WL("Nameserver_PingAnswer");
        }
        static void SendPing()
        {
            messenger.NameserverProcessor.SendMessage(new PingMessage());
        }
        static void NameserverProcessor_ConnectionEstablished(object sender, EventArgs e)
        {
            WL("NameserverProcessor_ConnectionEstablished");
        }

        /// <summary>
        /// 根据输入值，返回相应的语句
        /// </summary>
        /// <param name="inputTxt"></param>
        /// <returns></returns>
        static string GetResponse(string inputTxt)
        {
            string content = null;
            DataRow dataRow = null;
            //string dbconnection = "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;Data Source=MyMSN.mdb";

            //这里的CHARINDEX换成INSTR
            //string sql = "SELECT CategoryID FROM RobotKeywords WHERE (INSTR( '" + Regex.Replace(inputTxt, "'", "''") + "',KeywordContent) <> 0)";
            string sql = "SELECT top 1 * FROM RobotKeywords WHERE KeywordContent like '%"+inputTxt+"%'";
            DataTable dt = AccessHelper.ExecuteDataTable(sql);

            if (dt.Rows.Count > 0)
            {
                dataRow = dt.Rows[0];
                content = dataRow["CategoryID"].ToString();
            }

            if (content != null && content != "")
            {
                dt = AccessHelper.ExecuteDataTable("SELECT TOP 1 * FROM RobotResponses WHERE (CategoryID = " + content + ") ORDER BY [ResponseID]");
                if (dt.Rows.Count > 0)
                {
                    dataRow = dt.Rows[0];
                    content = dataRow["ResponseContent"].ToString();
                }
            }

            if (content == null || content == "")
            {
                dt = AccessHelper.ExecuteDataTable("SELECT TOP 1 * FROM RobotResponses WHERE (CategoryID = 5) ORDER BY [ResponseID]");
                if (dt.Rows.Count > 0)
                {
                    dataRow = dt.Rows[0];
                    content = dataRow["ResponseContent"].ToString();
                }
            }

            return content;

        }

        #region Helper methods

        private static void WL(object text, params object[] args)
        {
            Console.WriteLine(text.ToString(), args);
        }

        private static void RL()
        {
            Console.ReadLine();
        }

        private static void Break()
        {
            System.Diagnostics.Debugger.Break();
        }

        #endregion
    }
    public class PingMessage : XihSolutions.DotMSN.Core.NSMessage
    {
        public override byte[] GetBytes()
        {

            return System.Text.Encoding.UTF8.GetBytes("PNG\r\n");
        }
    }

   
}

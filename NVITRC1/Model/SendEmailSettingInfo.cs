using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace NVITRC1.Model
{
    public class SendEmailSettingInfo
    {
        private string customerName;

        private string emailAccount;

        private string emailPassword;

        private List<string> emailRecipientList;

        private List<string> ccRecipientList;

        private string smtpHost;

        private int smtpPort;

        private DateTime lastEmailTime;

        public string CustoemrName
        {
            get
            {
                return customerName;
            }
            set
            {
                customerName = value;
            }
        }

        public string EmailAccount
        {
            get
            {
                return emailAccount;
            }
            set
            {
                emailAccount = value;
            }
        }

        public string EmailPassword
        {
            get
            {
                return emailPassword;
            }
            set
            {
                emailPassword = value;
            }
        }

        public List<string> EmailRecipientList
        {
            get
            {
                return emailRecipientList;
            }
            set
            {
                emailRecipientList = value;
            }
        }

        public List<string> CcRecipientList
        {
            get
            {
                return ccRecipientList;
            }
            set
            {
                ccRecipientList = value;
            }
        }

        public string SmtpHost
        {
            get
            {
                if (smtpHost == null || smtpHost.Trim() == "")
                {
                    return "smtp.qiye.163.com";
                }
                else
                {
                    return smtpHost;
                }
            }
            set
            {
                smtpHost = value;
            }
        }

        public int SmtpPort
        {
            get
            {
                if (smtpPort == 0)
                {
                    return 25;
                }
                else
                {
                    return smtpPort;
                }
            }
            set
            {
                smtpPort = value;
            }
        }

        public DateTime LastEmailTime
        {
            get
            {
                return lastEmailTime;
            }
            set
            {
                lastEmailTime = value;
            }
        }

        public SendEmailSettingInfo(string emailAcount, string emailPassword, List<string> emailRecipientList, List<string> ccRecipientList, string smtpHost, int smtpPort)
        {
            this.EmailAccount = emailAcount;
            this.EmailPassword = emailPassword;
            this.EmailRecipientList = emailRecipientList;
            this.CcRecipientList = ccRecipientList;
            this.SmtpHost = smtpHost;
            this.SmtpPort = smtpPort;
        }

        public SendEmailSettingInfo(string emailAcount, string emailPassword, string emailRecipient, string ccRecipient)
        {
            this.EmailAccount = emailAcount;
            this.EmailPassword = emailPassword;
            this.EmailRecipientList = emailRecipientList;
            this.CcRecipientList = ccRecipientList;
        }

        public SendEmailSettingInfo()
        {

        }
    }
}

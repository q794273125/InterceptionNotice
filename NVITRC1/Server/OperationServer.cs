using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Data;
using NVITRC1.Common;
using NVITRC1.Model;

namespace NVITRC1.Server
{
    public class OperationServer
    {
        public string settingPath { get; set; }
        public OperationServer(string _settingPath)
        {
            this.settingPath = _settingPath;
        }


        public DataTable GetDataTable(SqlHelper sqlHelper, string beginTime, string endTime, bool flagWhere)
        {
            DataTable dataTable = new DataTable();
            StringBuilder sqlSB = new StringBuilder();
            sqlSB.Append(string.Format(@"select [deposit_number],[wo_number],[type],[status],[ref_num],[in_param],[err_msg],[updated_on] FROM [order].[dbo].[hkows_web_service_logs] WHERE in_param LIKE '%NA354%' AND status = 'f' AND type = 'I'"));
            if (flagWhere)
            {
                if (!string.IsNullOrEmpty(beginTime) && !string.IsNullOrEmpty(endTime))
                {
                    sqlSB.Append(string.Format("AND updated_on BETWEEN '{0}' AND '{1}'", beginTime, endTime));
                }
            }

            DataSet dataSet = sqlHelper.Query(sqlSB.ToString());
            if (dataSet != null && dataSet.Tables.Count > 0)
            {
                dataTable = dataSet.Tables[0];
            }
            return dataTable;
        }

        public void CreateDirectory(string directoryPath)
        {
            if (!Directory.Exists(directoryPath))
            {
                Directory.CreateDirectory(directoryPath);
            }
        }

        #region  发送邮件
        /// <summary>
        /// 发送邮件
        /// </summary>
        /// <param name="richTextBox"></param>
        /// <param name="as702List"></param>
        /// <param name="factoryName"></param>
        /// <param name="recipientsList"></param>
        public bool SentEmail(List<string> attachmentPathList, SendEmailSettingInfo sendEmailSettingInfo, StringBuilder logSB, DateTime dateTime, StringBuilder tableSB)
        {
            bool flagSentEmail = true;
            try
            {
                StringBuilder emailContentSb = new StringBuilder();
                emailContentSb.Append("您好！<br><br>");
                emailContentSb.Append("**請注意：此郵件是系統自動傳送，請勿直接回覆此郵件<br><br>");
                emailContentSb.Append("以下是 <span style='font-weight:bolder;text-decoration:underline;'>NA354被拦截</span> 的訂單<br><br>");
                emailContentSb.Append(string.Format("訂單來單日期：{0}至{1}<br><br>", dateTime.AddDays(-1).ToString("yyyy年MM月dd日 09:00"), dateTime.ToString("yyyy年MM月dd日 09:00")));
                emailContentSb.AppendLine(tableSB.ToString());
                string titleInfo = "NA354拦截订单邮件提醒";


                emailContentSb.AppendLine("<br><br>");
                emailContentSb.AppendLine("<br>本邮件由系统自动发出，请不要回复，如有问题请联系IT");

                emailContentSb.Append($@"<p class=""stamp"">{DateTime.Now.ToString("yyyy年MM月dd日")}:</p>
        <hr width=""150px;"" align=""left"">
        <p class=""stamp"">HKO IT Team</p>
        <!--<img src=""img/logo.png"">-->
        <img src='data:image/png;base64,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' />
        <p>Disclaimer:</p>
        <p>
            This email and its attachment(s) are intended solely for the use of the individual to whom it is addressed and may
            contain information which is privileged, confidential, proprietary, or exempt from disclosure under applicable law.
            If you are not the intended recipient or the person responsible for delivering the message to the intended
            recipient, you are strictly prohibited from disclosing, distributing, copying, or in any way using this message. If
            you have received this email in error, please notify the sender and destroy / delete any copies you may have
            received.
        </p>");

                EmailHelper.SendMail("HK-IT", sendEmailSettingInfo.EmailAccount, sendEmailSettingInfo.EmailPassword, sendEmailSettingInfo.EmailRecipientList, sendEmailSettingInfo.CcRecipientList, string.Format("{0}", titleInfo), emailContentSb.ToString(), attachmentPathList, sendEmailSettingInfo.SmtpHost, sendEmailSettingInfo.SmtpPort, false);
            }
            catch (Exception ex)
            {
                logSB.AppendLine(ex.Message);
                flagSentEmail = false;
            }
            return flagSentEmail;
        }
        #endregion



        /// <summary>
        /// 装载邮件信息
        /// </summary>
        public SendEmailSettingInfo GetSendEmailSettingInfo()
        {
            string str = "/set/customers/customer";
            string smtpHost = XmlHelper.Read(settingPath, str + "/sendEmailSets/sendEmailSet/smtpHost");
            string smtpPort = XmlHelper.Read(settingPath, str + "/sendEmailSets/sendEmailSet/smtpPort");
            string emailAccount = XmlHelper.Read(settingPath, str + "/sendEmailSets/sendEmailSet/emailAccount");
            string emailPassword = XmlHelper.Read(settingPath, str + "/sendEmailSets/sendEmailSet/emailPassword");
            string emailRecipients = XmlHelper.Read(settingPath, str + "/sendEmailSets/sendEmailSet/emailRecipients");
            string ccRecipients = XmlHelper.Read(settingPath, str + "/sendEmailSets/sendEmailSet/ccRecipients");
            string lastEmailTime = XmlHelper.Read(settingPath, str + "/sendEmailSets/sendEmailSet/lastEmailTime");
            string sendEmailCount = XmlHelper.Read(settingPath, str + "/sendEmailSets/sendEmailSet/sendEmailCount");
            List<string> strEmailRecipientList = new List<string>();
            List<string> emailRecipientList = emailRecipients.Split(';').ToList();
            foreach (string emailRecipient in emailRecipientList)
            {
                if (!string.IsNullOrEmpty(emailRecipient))
                {
                    strEmailRecipientList.Add(emailRecipient);
                }
            }

            List<string> strCcRecipientsList = new List<string>();
            List<string> ccRecipientsList = ccRecipients.Split(';').ToList();
            foreach (string emailRecipient in ccRecipientsList)
            {
                if (!string.IsNullOrEmpty(emailRecipient))
                {
                    strCcRecipientsList.Add(emailRecipient);
                }
            }


            SendEmailSettingInfo sendEmailSettingInfo = new SendEmailSettingInfo()
            {
                SmtpHost = smtpHost,
                SmtpPort = int.Parse(smtpPort),
                EmailAccount = emailAccount,
                EmailPassword = emailPassword,
                EmailRecipientList = strEmailRecipientList,//电子邮件收件人列表
                CcRecipientList = strCcRecipientsList,//抄送收件人列表
                LastEmailTime = Convert.ToDateTime(lastEmailTime),
                //sendEmailCount = 1
            };
            return sendEmailSettingInfo;
        }

        /// <summary>
        /// 装载数据库信息
        /// </summary>
        public DatabaseSetSettingInfo GetDatabaseSetSettingInfo()
        {
            string str = "/set/customers/customer";
            string dataSource = XmlHelper.Read(settingPath, str + "/databaseSet/dataSource");
            string database = XmlHelper.Read(settingPath, str + "/databaseSet/database");
            string user = XmlHelper.Read(settingPath, str + "/databaseSet/user");
            string pwd = XmlHelper.Read(settingPath, str + "/databaseSet/pwd");

            DatabaseSetSettingInfo databaseSetSettingInfo = new DatabaseSetSettingInfo()
            {
                dbHost = dataSource,
                dbName = database,
                dbUser = user,
                dbPwd = pwd
            };
            return databaseSetSettingInfo;
        }
    }
}

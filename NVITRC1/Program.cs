using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Data;  
using NVITRC1.Model;
using NVITRC1.Common;
using NVITRC1.Server;

namespace NVITRC1
{
    class Program
    {
        public static Program program = new Program();
        public static readonly string settingPath = AppDomain.CurrentDomain.BaseDirectory + "\\config\\set.xml";
        public static string startupPath = AppDomain.CurrentDomain.BaseDirectory;
        static void Main(string[] args)
        {
            StringBuilder sbLog = new StringBuilder(); //可变的字符串
            DateTime dateTime = DateTime.Now;           //获取当前的日期

            sbLog.AppendLine(string.Format("数据开始解析，当前时间为：{0}", dateTime.ToString("yyyy-MM-dd HH:mm:ss")));

            string beginTime = dateTime.AddDays(-1).ToString("yyyy-MM-dd 09:00:00");
            string endTime = dateTime.ToString("yyyy-MM-dd 09:00:00");

            OperationServer operationServer = new OperationServer(settingPath);

            sbLog.AppendLine(string.Format("开始创建文价夹"));
            string excelPath = Path.Combine(startupPath, @"Result\Excel\");
            string logPath = Path.Combine(startupPath, @"Result\Log\");
            try
            {
                operationServer.CreateDirectory(excelPath);
                operationServer.CreateDirectory(logPath);
                sbLog.AppendLine(string.Format("文价夹创建成功"));

                SendEmailSettingInfo sendEmailSettingInfo = operationServer.GetSendEmailSettingInfo();
                DatabaseSetSettingInfo databaseSetSettingInfo = operationServer.GetDatabaseSetSettingInfo();

                SqlHelper sqlHelper = new SqlHelper(dbHost: databaseSetSettingInfo.dbHost, dbName: databaseSetSettingInfo.dbName, dbUser: databaseSetSettingInfo.dbUser, dbPwd: databaseSetSettingInfo.dbPwd);

                DataTable dataTable = operationServer.GetDataTable(sqlHelper, beginTime, endTime, true);
                sbLog.AppendLine(string.Format("数据获取成功，共有{0}条", dataTable.Rows.Count));

                string excelFilePath = string.Empty;
                if (dataTable != null && dataTable.Rows.Count > 0)
                {
                    excelFilePath = Path.Combine(excelPath, string.Format("NA354_{0}_{1}_{2}_拦截单.xlsx", dateTime.ToString("yyyy"), dateTime.ToString("MM"), dateTime.ToString("dd")));
                    NPOIHelper npoiHelper = new NPOIHelper();
                    npoiHelper.ExportDataTableToExcel(dataTable, excelFilePath);
                    sbLog.AppendLine(string.Format("Excel生成成功，路径为{0}", excelFilePath));
                }

                if (!string.IsNullOrEmpty(excelFilePath))
                {
                    StringBuilder tableSB = new StringBuilder();
                    tableSB.AppendLine("<table border='1' cellspacing='0' cellpadding='0' style='text-align:center;'>");
                    tableSB.AppendLine("<tr style='background-color:#5b9bd5'>" +
                        "<th width='300px'>deposit_number</th>" +
                        "<th width='350px'>wo_number</th>" +
                        "<th width='200px'>type</th>" +
                        "<th width='200px'>status</th>" +
                        "<th width='300px'>ref_num</th>" +
                        "<th width='200px'>in_param</th>" +
                        "<th width='300px'>err_msg</th>" +
                        "<th width='300px'>updated_on</th>" +
                        "</tr>");
                    if (dataTable.Rows.Count > 0)
                    {
                        foreach (DataRow row in dataTable.Rows)
                        {
                            tableSB.AppendLine(string.Format("<tr><td>{0}</td><td>{1}</td>" +
                                "<td>{2}</td>" +
                                "<td>{3}</td>" +
                                "<td>{4}</td>" +
                                "<td>{5}</td>" +
                                "<td>{6}</td>" +
                                "<td>{7}</td></tr>",
                                row["deposit_number"].ToString(),
                                row["wo_number"].ToString(),
                                row["type"].ToString(),
                                row["status"].ToString(),
                                row["ref_num"].ToString(),
                                row["in_param"].ToString(),
                                row["err_msg"].ToString(),
                                row["updated_on"].ToString()));
                        }
                    }
                    tableSB.AppendLine($"</table");
                    sbLog.AppendLine(string.Format("Table拼接成功"));

                    bool flagSentEmail = operationServer.SentEmail(new List<string>() { excelFilePath }, sendEmailSettingInfo, sbLog, dateTime, tableSB);
                    sbLog.AppendLine(string.Format("发送邮件{0}", flagSentEmail ? "成功" : "失败"));
                }
            }
            catch (Exception ex)
            {
                sbLog.AppendLine(string.Format("发送错误：{0}", ex.Message));
            }

            string logFilePath = Path.Combine(logPath, "Log_" + dateTime.ToString("yyyyMMddHHmmss") + ".TXT");
            TxtHelper.Write(logFilePath, sbLog.ToString());

        }
    }
}

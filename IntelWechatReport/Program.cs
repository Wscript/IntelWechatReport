using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Configuration;
using System.Data;
using System.Data.OleDb;
using System.Data.SqlClient;
using System.Net.Mail;

namespace IntelWechatReport
{
    class Program
    {
        static void Main (string[] args)
        {
            string strReportDate;

            if (args.Length == 0)
            {
                strReportDate = DateTime.Now.Year.ToString () + "-" + DateTime.Now.Month.ToString () + "-" + DateTime.Now.Day.ToString ();
            }else
            {
                strReportDate = args[0].ToString();
            }

            //定义数据内存中缓存,后面填充数据使用  
            DataTable dt = new DataTable ();

            //定义数据库连接语句
            string consqlserver = ConfigurationManager.ConnectionStrings["IDSSConnectionString"].ToString () + ";Password=CSD;";

            //定义SQL查询语句  
            string sql = "EXEC idss_prog_Wechat_Report";

            //定义SQL Server连接对象  
            SqlConnection con = new SqlConnection (consqlserver);

            //数据库命令和数据库连接  
            SqlDataAdapter da = new SqlDataAdapter (sql, con);

            try
            {
                da.Fill (dt);                                    //填充数据
                Log ("数据读取成功！", ConfigurationManager.AppSettings["LogFolder"].ToString (), "IntelWechatReport.log");

                if (dt.Rows.Count > 0)                //判断是否符合条件的数据记录  
                {
                    //判断之前的Report是否存在，如果存在则删除
                    if (System.IO.File.Exists ("WechatReport.xls"))
                    {
                        System.IO.File.Delete ("WechatReport.xls");
                    }

                    //新建一个空白的Report
                    String strExcelConnection = "Provider=Microsoft.Jet.OLEDB.4.0;" +
                                                "Data Source=WechatReport.xls;" +
                                                "Extended Properties=Excel 8.0;";
                    OleDbConnection objExcelConnection = new OleDbConnection (strExcelConnection);
                    OleDbCommand cmdExcel = new OleDbCommand ("Create table Sheet1 ([OrderNo] VarChar, " +
                                                                                    "[LineID] VarChar, " +
                                                                                    "[OrderType] VarChar, " +
                                                                                    "[SKU] VarChar, " +
                                                                                    "[PlanSN] VarChar, " +
                                                                                    "[Waybill] VarChar," +
                                                                                    "[Status] VarChar," +
                                                                                    "[StatusChangeTime] Datetime)", objExcelConnection);
                    objExcelConnection.Open ();
                    cmdExcel.ExecuteNonQuery ();
                    Log ("文件创建成功！", ConfigurationManager.AppSettings["LogFolder"].ToString (), "IntelWechatReport.log");

                    //将从数据库中取出的数据写入到新建的Report中
                    for (int i = 0; i < dt.Rows.Count; i++)
                    {
                        cmdExcel.CommandText = string.Format("INSERT INTO [Sheet1$] VALUES ('{0}','{1}','{2}','{3}','{4}','{5}','{6}','{7}')",
                                                                dt.Rows[i][0].ToString(),
                                                                dt.Rows[i][1].ToString(),
                                                                dt.Rows[i][2].ToString(),
                                                                dt.Rows[i][3].ToString(),
                                                                dt.Rows[i][4].ToString(),
                                                                dt.Rows[i][5].ToString(),
                                                                dt.Rows[i][6].ToString(),
                                                                dt.Rows[i][7]);
                        cmdExcel.ExecuteNonQuery ();
                    }

                    //关闭Excel文件
                    objExcelConnection.Close ();
                    Log ("记录写入成功，共 " + dt.Rows.Count.ToString() + "条.", ConfigurationManager.AppSettings["LogFolder"].ToString (), "IntelWechatReport.log");

                    //将新生成的Report发送出去
                    SendMailUseGmail ();

                    //备份文件，并删除原文件
                    System.IO.File.Copy ("WechatReport.xls",
                                            ConfigurationManager.AppSettings["BackupFolder"].ToString () +
                                            "WechatReport" + " " +
                                            DateTime.Now.Year.ToString () + "-" + DateTime.Now.Month.ToString () + "-" + DateTime.Now.Day.ToString () + " " +
                                            DateTime.Now.Hour.ToString () + "." + DateTime.Now.Minute.ToString () + "." + DateTime.Now.Second.ToString () + " " + 
                                            ".xls",true);
                    //System.IO.File.Delete ("WechatReport.xls");
                    Log ("文件备份成功！", ConfigurationManager.AppSettings["LogFolder"].ToString (), "IntelWechatReport.log");
                }
            }
            catch (Exception msg)
            {
                //将异常写入Log文件
                Log ("处理出现异常 " + msg.Message, ConfigurationManager.AppSettings["LogFolder"].ToString (), "IntelWechatReport.log");
            }
            finally
            {
                //关闭连接，并释放资源
                con.Close ();
                con.Dispose ();
                da.Dispose ();
                dt.Dispose ();
            }  
        }

        static void SendMailUseGmail ()
        {
            System.Net.Mail.MailMessage MailMessage = new System.Net.Mail.MailMessage ();

            if (ConfigurationManager.AppSettings["ToAddress"].ToString () != "")
            {
                MailMessage.To.Add (ConfigurationManager.AppSettings["ToAddress"].ToString ());
            }

            if (ConfigurationManager.AppSettings["CCAddress"].ToString () != "")
            {
                MailMessage.CC.Add (ConfigurationManager.AppSettings["CCAddress"].ToString ());
            }

            if (ConfigurationManager.AppSettings["BCCAddress"].ToString () != "")
            {
                MailMessage.Bcc.Add (ConfigurationManager.AppSettings["BCCAddress"].ToString ());
            }

            //参数分别是发件人地址（可以随便写），发件人姓名，编码
            MailMessage.From = new MailAddress (ConfigurationManager.AppSettings["MailAddress"].ToString (),
                                                ConfigurationManager.AppSettings["SenderName"].ToString (),
                                                System.Text.Encoding.UTF8);

            MailMessage.Subject = ConfigurationManager.AppSettings["MailSubject"].ToString ();
            MailMessage.SubjectEncoding = System.Text.Encoding.UTF8;

            MailMessage.Body = ConfigurationManager.AppSettings["MailBody"].ToString ();
            MailMessage.BodyEncoding = System.Text.Encoding.UTF8;

            MailMessage.IsBodyHtml = false;                             //是否是HTML邮件
            MailMessage.Priority = MailPriority.Normal;                 //邮件优先级

            MailMessage.Attachments.Add (new Attachment ("WechatReport.xls"));

            SmtpClient SmtpClient = new SmtpClient ();
            SmtpClient.Credentials = new System.Net.NetworkCredential (ConfigurationManager.AppSettings["MailAcount"].ToString (),
                                                                       ConfigurationManager.AppSettings["MailPassword"].ToString ());
            //上述写你的GMail邮箱和密码

            SmtpClient.Port = System.Convert.ToInt32(ConfigurationManager.AppSettings["MailPort"].ToString ());
            SmtpClient.Host = ConfigurationManager.AppSettings["MailHost"].ToString ();
            SmtpClient.EnableSsl = true;
            object userState = MailMessage;
            try
            {
                SmtpClient.Send (MailMessage);
                Log ("邮件发送成功！", ConfigurationManager.AppSettings["BackupFolder"].ToString (), "IntelWechatReport.log");
             }
            catch (System.Net.Mail.SmtpException ex)
            {
                Log ("发送邮件出错 " + ex.Message, ConfigurationManager.AppSettings["LogFolder"].ToString (), "IntelWechatReport.log");
            }
        }

        static void Log (string logMessage, string logFolder, string logFilename)
        {
            using (StreamWriter w = File.AppendText (logFolder + logFilename))
            {
                w.Write ("\r\n");
                w.Write ("{0} {1}", DateTime.Now.ToLongDateString (), DateTime.Now.ToLongTimeString ());
                w.Write ("  :{0}", logMessage);
            }
        }
    }
}

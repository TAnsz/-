using System;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Mail;
using System.ServiceProcess;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace SendMailService
{
    public partial class Service1 : ServiceBase
    {
        /// <summary>
        /// 用于监听数据表中数据变化
        /// </summary>
        const string SelectStr = @"SELECT ID,Sender,ToMailAddr,CcMailAddr,BccMailAddr,Subject,BodyPart,Creator,CreateDate,Status,SendTime,SendMsg  FROM  dbo.AutoMail WHERE Status = 0";
        const string SelectAllStr = @"SELECT ID,Sender,ToMailAddr,CcMailAddr,BccMailAddr,Subject,BodyPart,Creator,CreateDate,Status,SendTime,SendMsg  FROM  dbo.AutoMail WHERE Status <> 1";
        string timeout = ConfigurationManager.AppSettings["retryTimeOut"] ?? "10";
        string connStr;
        DataTable _dt = new DataTable();
        SmtpClient smtp;

        public Service1()
        {
            InitializeComponent();
        }

        protected override void OnStart(string[] args)
        {
            Log("服务启动");
            Start();
        }

        protected override void OnStop()
        {
            if (smtp != null)
            {
                try
                {
                    smtp.SendAsyncCancel();
                }
                catch (Exception e)
                {
                    Log(e.Message);
                }
                finally
                {
                    smtp.Dispose();
                }
            }
            SqlDependency.Stop(connStr);
            Log("服务停止");
        }

        protected override void OnShutdown()
        {
            if (smtp != null)
            {
                try
                {
                    smtp.SendAsyncCancel();
                }
                catch (Exception e)
                {
                    Log(e.Message);
                }
                finally
                {
                    smtp.Dispose();
                }
            }
            SqlDependency.Stop(connStr);
            Log("系统关闭");
        }

        void Start()
        {
            try
            {
                ConfigurationManager.RefreshSection("ConnectionStrings");
                connStr = ConfigurationManager.ConnectionStrings["H3CloudConnectionString"].ConnectionString;
                RefreshData(SelectAllStr);
                if (_dt.Rows.Count > 0)
                {
                    Log("开始处理服务停止期间未发送的邮件");
                    SendMail();
                    Log("未发送邮件处理完成");
                }
                SqlDependency.Start(connStr);
                Monitor();
            }
            catch (Exception e)
            {
                Log(e.Message);
                Thread.Sleep(1000 * 60* Convert.ToInt16(timeout));
                Start();
            }
        }

        private void SendMail()
        {
            if (_dt.Rows.Count > 0)
            {
                try
                {
                    var ip = ConfigurationManager.AppSettings["smtpclient"];
                    var port = Convert.ToInt32(ConfigurationManager.AppSettings["port"]);
                    smtp = new SmtpClient(ip, port)
                    {
                        EnableSsl = false,
                        UseDefaultCredentials = false,
                        DeliveryMethod = SmtpDeliveryMethod.Network,
                        Credentials = new NetworkCredential("12", "123"),
                        Timeout = 100000
                    };
                    var mailMsgs = _dt.AsEnumerable().Select(x =>
                    {
                        var mailMsg = new MailMessage();
                        mailMsg.From = new MailAddress(x.Field<string>("Sender") + "@renolit.com.cn", x.Field<string>("Sender"));
                        mailMsg.To.Add(x.Field<string>("ToMailAddr").Replace(';', ','));
                        if (!string.IsNullOrEmpty(x.Field<string>("CcMailAddr")))
                        {
                            mailMsg.CC.Add(x.Field<string>("CcMailAddr").Replace(';', ','));
                        }
                        if (!string.IsNullOrEmpty(x.Field<string>("BccMailAddr")))
                        {
                            mailMsg.Bcc.Add(x.Field<string>("BccMailAddr").Replace(';', ','));
                        }
                        mailMsg.Subject = x.Field<string>("Subject");
                        mailMsg.Body = x.Field<string>("BodyPart");
                        mailMsg.SubjectEncoding = Encoding.UTF8;
                        mailMsg.BodyEncoding = Encoding.UTF8;
                        mailMsg.HeadersEncoding = Encoding.UTF8;
                        mailMsg.IsBodyHtml = true;
                        return new { Id = x.Field<long>("ID"), Sender = x.Field<string>("Sender"), To = x.Field<string>("ToMailAddr"), Subject = x.Field<string>("Subject"), Msg = mailMsg };
                    }).ToList();

                    foreach (var maiMsg in mailMsgs)
                    {
                        var dr = _dt.Rows.Find(maiMsg.Id);
                        dr["SendTime"] = DateTime.Now;
                        try
                        {
                            Log($"ID:{maiMsg.Id} 收件人：{maiMsg.To} 主题：{maiMsg.Subject} 的邮件开始发送！");
                            smtp.Send(maiMsg.Msg);
                            dr["Status"] = "1";
                            Log($"ID:{maiMsg.Id} 收件人：{maiMsg.To} 主题：{maiMsg.Subject} 的邮件发送成功！");
                        }
                        catch (Exception e)
                        {
                            dr["Status"] = "2";
                            dr["SendMsg"] = e.Message;
                            Log($"ID:{maiMsg.Id} 收件人：{maiMsg.To} 主题：{maiMsg.Subject} 的邮件发送失败！异常：" + e.Message);
                        }
                    }
                    Update();
                }
                catch (Exception ex)
                {
                    Log($"的邮件发送出现异常：" + ex.Message);
                }
            }
            if (smtp != null)
            {
                smtp.Dispose();
            }
            Monitor();
        }

        void dependency_OnChange(object sender, SqlNotificationEventArgs e)
        {
            if (e.Info == SqlNotificationInfo.Insert ||
              e.Info == SqlNotificationInfo.Update ||
              e.Info == SqlNotificationInfo.Delete)
            {
                Log("监控到邮件任务");
                RefreshData(SelectStr);
                SendMail();
                Log("邮件处理完成");
            }
            try
            {
                Monitor();
            }
            catch (Exception ex)
            {
                Log(ex.Message);
                Start();
            }
        }

        void RefreshData(string sqlStr)
        {
            try
            {
                connStr = ConfigurationManager.ConnectionStrings["H3CloudConnectionString"].ToString();
                SqlDataAdapter da = new SqlDataAdapter(sqlStr, connStr);
                _dt.Reset();
                da.Fill(_dt);
                _dt.PrimaryKey = new DataColumn[] { _dt.Columns["ID"] };
            }
            catch (Exception e)
            {
                throw e;
            }
        }

        void Monitor()
        {
            try
            {
                //注册数据表监视
                SqlDataAdapter da = new SqlDataAdapter(SelectStr, connStr);
                SqlDependency dependency = new SqlDependency(da.SelectCommand);
                dependency.OnChange += dependency_OnChange;
                da.Fill(_dt);//注册以后一定要执行一次取数
                Log("开始监控邮件数据表");
            }
            catch (Exception e)
            {
                throw e;
            }
        }

        void Update()
        {
            SqlDataAdapter da = new SqlDataAdapter(SelectStr, connStr);
            var scb = new SqlCommandBuilder(da);
            da.Update(_dt.GetChanges());
            _dt.AcceptChanges();
        }
        void Log(string str)
        {
            using (StreamWriter sw = File.AppendText(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, DateTime.Now.ToString("yyyyMMdd") + ".txt")))
            {
                sw.WriteLine("{0}==>{1}", DateTime.Now.ToShortTimeString(), str);
            }
        }
    }
}

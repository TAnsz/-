using MailHelpers;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.ServiceProcess;
using System.Text;
using System.Threading.Tasks;

namespace SendMailService
{
    public partial class Service1 : ServiceBase
    {
        string connStr = ConfigurationManager.ConnectionStrings["H3CloudConnectionString"].ToString();
        public Service1()
        {
            InitializeComponent();
        }

        protected override void OnStart(string[] args)
        {
            SqlDependency.Start(connStr);
            Log("程序启动");
            SendMail();
        }

        protected override void OnStop()
        {
            SqlDependency.Stop(connStr);
        }

        protected override void OnShutdown()
        {
            SqlDependency.Stop(connStr);
        }

        private void SendMail()
        {
            /// <summary>
            /// 用于监听数据表中数据变化
            /// </summary>
            string SelectStr = @"SELECT ID,Sender,ToMailAddr,CcMailAddr,BccMailAddr,Subject,BodyPart,Creator,CreateDate,Status,SendTime,SendMsg  FROM  dbo.AutoMail WHERE Status = 0";
            //处理未发送的邮件
            SqlDataAdapter da = new SqlDataAdapter(SelectStr, connStr);
            var dt = new DataTable();
            da.Fill(dt);
            dt.PrimaryKey = new DataColumn[] { dt.Columns["ID"] };

            if (dt.Rows.Count > 0)
            {
                var mail = new MailHelper();
                mail.SetBatchMailCount(dt.Rows.Count);
                string msg = "";
                foreach (DataRow dr in dt.Rows)
                {
                    msg = $"正在发送ID:{dr["ID"]} 收件人：{dr["ToMailAddr"]} 主题：{dr["Subject"]} 的邮件中...";
                    try
                    {
                        mail.SetMailInfo(dr["Subject"].ToString(), dr["BodyPart"].ToString(), dr["ToMailAddr"].ToString(), dr["CcMailAddr"].ToString(), dr["BccMailAddr"].ToString());
                        mail.AsycUserState = dr;
                        mail.SendMail();
                        dr["Status"] = "1";
                        dr["SendTime"] = DateTime.Now;
                        Log($"ID:{dr["ID"]} 收件人：{dr["ToMailAddr"]} 主题：{dr["Subject"]} 的邮件发送成功");
                    }
                    catch (Exception e)
                    {
                        msg += e.ToString();
                        dr["Status"] = "0";
                        dr["SendTime"] = DateTime.Now;
                        dr["SendMsg"] = e.ToString();
                        Log($"ID:{dr["ID"]} 收件人：{dr["ToMailAddr"]} 主题：{dr["Subject"]} 的邮件发送失败！异常：" + e.ToString());

                    }
                }
                SqlCommandBuilder scb = new SqlCommandBuilder(da);
                da.Update(dt.GetChanges());
                dt.AcceptChanges();
            }
            //注册数据表监视
            SqlDependency dependency = new SqlDependency(da.SelectCommand);
            dependency.OnChange += dependency_OnChange;
            da.Fill(dt);//注册以后一定要执行一次取数
        }


        void dependency_OnChange(object sender, SqlNotificationEventArgs e)
        {
            if (e.Info == SqlNotificationInfo.Insert ||
              e.Info == SqlNotificationInfo.Update ||
              e.Info == SqlNotificationInfo.Delete)
            {
                SendMail();
            }
        }
        void Log(string str)
        {
            using (StreamWriter sw = File.AppendText(DateTime.Now.ToString("yyyyMMdd") + ".txt"))
            {
                sw.WriteLine(str);
            }
        }
    }
}

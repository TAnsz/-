using System;
using System.Data.SqlClient;
using System.Windows.Forms;
using System.Configuration;
using System.Data;
using System.Collections.Generic;
using System.Threading;
using System.Net.Mail;
using System.Text;
using System.Linq;
using System.Net;
using System.Threading.Tasks;

namespace SendMailAuto
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
        string connStr = ConfigurationManager.ConnectionStrings["H3CloudConnectionString"].ToString();
        /// <summary>
        /// 用于显示页面发送邮件的结果数据
        /// </summary>
        string DataStr = @"SELECT TOP 100 ID,Sender,ToMailAddr,CcMailAddr,BccMailAddr,Subject,BodyPart,Creator,CreateDate,Status,SendTime,SendMsg  FROM  dbo.AutoMail Order By ID DESC";
        /// <summary>
        /// 用于监听数据表中数据变化
        /// </summary>
        string SelectStr = @"SELECT ID,Sender,ToMailAddr,CcMailAddr,BccMailAddr,Subject,BodyPart,Creator,CreateDate,Status,SendTime,SendMsg  FROM  dbo.AutoMail WHERE Status = 0";

        DataTable _dt = new DataTable();

        void Form1_Load(object sender, EventArgs e)
        {
            var dic = new Dictionary<string, string> { { "0", "未发送" }, { "1", "已发送" } };
            BindingSource bs = new BindingSource();
            bs.DataSource = dic;

            var col_id = new DataGridViewTextBoxColumn { HeaderText = "ID", DataPropertyName = "ID", AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells };
            var col_send = new DataGridViewTextBoxColumn { HeaderText = "发件人", DataPropertyName = "Sender" };
            var col_to = new DataGridViewTextBoxColumn { HeaderText = "收件", DataPropertyName = "ToMailAddr" };
            var col_cc = new DataGridViewTextBoxColumn { HeaderText = "抄送", DataPropertyName = "CcMailAddr" };
            var col_bcc = new DataGridViewTextBoxColumn { HeaderText = "密送", DataPropertyName = "BccMailAddr" };
            var col_subject = new DataGridViewTextBoxColumn { HeaderText = "标题", DataPropertyName = "Subject" };
            var col_body = new DataGridViewTextBoxColumn { HeaderText = "正文", DataPropertyName = "BodyPart" };
            var col_status = new DataGridViewComboBoxColumn { HeaderText = "状态", DataPropertyName = "Status", DataSource = bs, DisplayMember = "Value", ValueMember = "Key" };
            var col_sendTime = new DataGridViewTextBoxColumn { HeaderText = "发送时间", DataPropertyName = "SendTime" };
            var col_sendMsg = new DataGridViewTextBoxColumn { HeaderText = "发送消息", DataPropertyName = "SendMsg" };
            dataGridView1.Columns.AddRange(new DataGridViewColumn[] { col_id, col_send, col_to, col_cc, col_bcc, col_subject, col_body, col_status, col_sendTime, col_sendMsg });
            dataGridView1.AutoGenerateColumns = false;
            DisplayData();
        }

        private void DisplayData()
        {
            SqlDataAdapter da = new SqlDataAdapter(DataStr, connStr);
            var dt = new DataTable();
            da.Fill(dt);
            //显示界面数据
            dataGridView1.Invoke((MethodInvoker)delegate
            {
                dataGridView1.DataSource = dt;
            });
        }

        private async Task SendMail()
        {
            //处理未发送的邮件
            SqlDataAdapter da = new SqlDataAdapter(SelectStr, connStr);
            _dt = new DataTable();
            da.Fill(_dt);
            _dt.PrimaryKey = new DataColumn[] { _dt.Columns["ID"] };

            if (_dt.Rows.Count > 0)
            {
                var ip = ConfigurationManager.AppSettings["smtpclient"];
                var port = Convert.ToInt32(ConfigurationManager.AppSettings["port"]);
                using (var mail = new SmtpClient(ip, port))
                {
                    mail.EnableSsl = false;
                    mail.UseDefaultCredentials = false;
                    mail.DeliveryMethod = SmtpDeliveryMethod.Network;
                    mail.Credentials = new NetworkCredential("12", "123");
                    mail.Timeout = 100000;
                    var mailMsgs = _dt.AsEnumerable().Select(x =>
                    {
                        var mailMsg = new MailMessage();
                        mailMsg.From = new MailAddress(x.Field<string>("Sender") + "@renolit.com.cn", x.Field<string>("Sender"));
                        mailMsg.To.Add(x.Field<string>("ToMailAddr").Replace(';', ','));
                        //mailMsg.CC.Add(x.Field<string>("CcMailAddr").Replace(';', ','));
                        //mailMsg.Bcc.Add(x.Field<string>("BccMailAddr").Replace(';', ','));
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
                            await mail.SendMailAsync(maiMsg.Msg);
                            dr["Status"] = "1";
                            //Log($"ID:{x.id} 收件人：{x.to} 主题：{x.subject} 的邮件发送成功！");
                            label1.Text = $"ID:{maiMsg.Id} 收件人：{maiMsg.To} 主题：{maiMsg.Subject} 的邮件发送成功！";
                        }
                        catch (Exception e)
                        {
                            dr["Status"] = "0";
                            dr["SendMsg"] = e.Message;
                            //Log($"ID:{x.id} 收件人：{x.to} 主题：{x.subject} 的邮件发送失败！异常：" + e);
                            label1.Text = $"ID:{maiMsg.Id} 收件人：{maiMsg.To} 主题：{maiMsg.Subject} 的邮件发送失败！异常：" + e.Message;
                        }
                    }
                }

                SqlCommandBuilder scb = new SqlCommandBuilder(da);
                da.Update(_dt.GetChanges());
                _dt.AcceptChanges();
                DisplayData();
            }
            //注册数据表监视
            SqlDependency dependency = new SqlDependency(da.SelectCommand);
            dependency.OnChange += dependency_OnChange;
            da.Fill(_dt);//注册以后一定要执行一次取数
        }

        async void dependency_OnChange(object sender, SqlNotificationEventArgs e)
        {
            if (e.Info == SqlNotificationInfo.Insert ||
              e.Info == SqlNotificationInfo.Update ||
              e.Info == SqlNotificationInfo.Delete)
            {
                await SendMail();
                DisplayData();
            }
        }
        void Form1_FormClosed(object sender, FormClosedEventArgs e)
        {
            SqlDependency.Stop(connStr);
        }

        private async void btn_start_Click(object sender, EventArgs e)
        {
            if (btn_start.Text == "开始")
            {
                SqlDependency.Stop(connStr);
                SqlDependency.Start(connStr);
                await SendMail();
                btn_start.Text = "结束";
                MessageBox.Show("开始监视数据库");
            }
            else
            {
                SqlDependency.Stop(connStr);
                btn_start.Text = "开始";
                MessageBox.Show("停止监视数据库");
            }


        }
    }
}

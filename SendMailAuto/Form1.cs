using System;
using System.Data.SqlClient;
using System.Windows.Forms;
using System.Configuration;
using System.Data;
using System.Collections.Generic;
using MailHelpers;
using System.Threading;

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

        private void SendMail()
        {
            //处理未发送的邮件
            SqlDataAdapter da = new SqlDataAdapter(SelectStr, connStr);
            var dt = new DataTable();
            da.Fill(dt);
            dt.PrimaryKey = new DataColumn[] { dt.Columns["ID"] };
            bool asnyc = true;

            if (dt.Rows.Count > 0)
            {
                var mail = new MailHelper(asnyc, "192.168.3.31", 25, "h3bpm", "Bpm3H287");
                mail.SetBatchMailCount(dt.Rows.Count);
                //设置异步回调函数
                if (asnyc)
                {
                    mail.AsyncCallback = (send, arg) =>
                    {
                        var dr = (DataRow)((MailUserState)arg.UserState).UserState;
                        var drf = dt.Rows.Find(dr["ID"]);
                        if (arg.Error == null)
                        {
                            drf["Status"] = "1";
                            drf["SendTime"] = DateTime.Now;
                            toolStrip1.Invoke((MethodInvoker)delegate
                            {
                                label1.Text = $"ID:{dr["ID"]} 收件人：{dr["ToMailAddr"]} 主题：{dr["Subject"]} 的邮件已发送完成";
                            });
                        }
                        else
                        {
                            label1.Text = $"ID:{dr["ID"]} 收件人：{dr["ToMailAddr"]} 主题：{dr["Subject"]} 的邮件发送失败!{Environment.NewLine}异常：" + arg.Error.InnerException == null ? arg.Error.Message : arg.Error.Message + arg.Error.InnerException.Message;
                            // 标识异常已处理，否则若有异常，会抛出异常
                            drf["Status"] = "0";
                            drf["SendTime"] = DateTime.Now;
                            drf["SendMsg"] = arg.Error.InnerException == null ? arg.Error.Message : arg.Error.Message + arg.Error.InnerException.Message;
                            ((MailUserState)arg.UserState).IsErrorHandle = true;
                        }
                        SqlCommandBuilder sc = new SqlCommandBuilder(da);
                        da.Update(dt.GetChanges());
                        dt.AcceptChanges();
                    };
                }
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

                    }
                    catch (Exception e)
                    {
                        msg += e.ToString();
                        //异步需要在回调函数中处理
                        if (!asnyc)
                        {
                            dr["Status"] = "0";
                            dr["SendTime"] = DateTime.Now;
                            dr["SendMsg"] = e.ToString();
                        }
                    }
                    finally
                    {
                        toolStrip1.Invoke((MethodInvoker)delegate
                        {
                            label1.Text = msg;
                        });
                    }
                    //异步需要在回调函数中处理
                    if (!asnyc)
                    {
                        toolStrip1.Invoke((MethodInvoker)delegate
                    {
                        pb1.ProgressBar.Value = (dt.Rows.IndexOf(dr) + 1) * 100 / dt.Rows.Count;
                    });
                    }
                }
                SqlCommandBuilder scb = new SqlCommandBuilder(da);
                da.Update(dt.GetChanges());
                dt.AcceptChanges();
                toolStrip1.Invoke((MethodInvoker)delegate
                {
                    label1.Text = "邮件处理完成";
                });
                while (mail.ExistsSmtpClient())
                {
                    Thread.Sleep(500);
                }
                DisplayData();

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
                DisplayData();
            }
        }
        void Form1_FormClosed(object sender, FormClosedEventArgs e)
        {
            SqlDependency.Stop(connStr);
        }

        private void btn_start_Click(object sender, EventArgs e)
        {
            if (btn_start.Text == "开始")
            {
                SqlDependency.Stop(connStr);
                SqlDependency.Start(connStr);
                SendMail();
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

using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Mail;
using System.Net.Mime;
using System.Text;
using System.Threading;
using System.Web;

namespace MailHelpers
{
    /// <summary>
    /// MailHelper 的摘要说明
    /// </summary>
    public class MailHelper
    {

        #region 构造函数

        /// <summary>
        /// 构建 MailHelper 实例
        /// </summary>
        /// <param name="isAsync">是否启用异步邮件发送，默认为同步发送</param>
        public MailHelper()
        {
        }
        /// <summary>
        /// 构建 MailHelper 实例
        /// </summary>
        /// <param name="smtpclient">smtp服务器地址</param>
        /// <param name="port">端口</param>
        /// <param name="user">用户名</param>
        /// <param name="password">密码</param>
        public MailHelper(string smtpclient = null, int port = 0, string user = null, string password = null)
        {
            SmtpClientAdd = smtpclient;
            Port = port;
            User = user;
            Password = password;
        }

        /// <summary>
        /// 构建 MailHelper 实例
        /// </summary>
        /// <param name="mSmtpClient">SmtpClient实例</param>
        /// <param name="autoReleaseSmtp">是否自动释放SmtpClient实例</param>
        public MailHelper(SmtpClient mSmtpClient, bool autoReleaseSmtp)
        {
            SetSmtpClient(mSmtpClient, autoReleaseSmtp);
        }

        #endregion
        #region 公开属性

        /// <summary>
        /// 设置此电子邮件的发信人地址。
        /// </summary>
        public string From { get; set; }
        /// <summary>
        /// 设置此电子邮件的发信人地址。
        /// </summary>
        public string FromDisplayName { get; set; }

        /// <summary>
        /// 设置此电子邮件的收件地址列表，分号(;)分隔。
        /// </summary>
        public string To { get; set; }

        /// <summary>
        /// 设置此电子邮件的抄送地址列表，分号(;)分隔。
        /// </summary>
        public string Cc { get; set; }

        /// <summary>
        /// 设置此电子邮件的密送地址列表，分号(;)分隔。
        /// </summary>
        public string Bcc { get; set; }


        /// <summary>
        /// 设置此电子邮件的主题。
        /// </summary>
        public string Subject { get; set; }
        /// <summary>
        /// 设置邮件正文。
        /// </summary>
        public string Body { get; set; }

        /// <summary>
        /// 设置邮件正文是否为 Html 格式的值。
        /// </summary>
        public bool IsBodyHtml { get; set; }

        private int priority = 0;
        /// <summary>
        /// 设置此电子邮件的优先级  0-Normal   1-Low   2-High
        /// 默认Normal。
        /// </summary>
        public int Priority
        {
            get { return this.priority; }
            set
            {
                if (value < 0 || value > 2)
                    priority = 0;
                else
                    priority = value;
            }
        }

        #endregion

        #region 内部字段、属性
        /// <summary>
        /// smtp服务器地址
        /// </summary>
        string SmtpClientAdd { get; set; }
        /// <summary>
        /// smtp端口
        /// </summary>
        int Port { get; set; }
        /// <summary>
        /// smtp用户名
        /// </summary>
        string User { get; set; }
        /// <summary>
        /// smtp密码
        /// </summary>
        string Password { get; set; }

        SmtpClient m_SmtpClient { get; set; }

        /// <summary>
        /// 默认为false。设置在 MailHelper 类内部，发送完邮件后是否自动释放 SmtpClient 实例
        /// Smtp不管是在 MailHelper 内部还是在外部都必须进行主动释放，
        /// 因为：SmtpClient 没有提供 Finalize() 终结器，所以GC不会进行回收，只能使用完后主动进行释放，否则会发生内存泄露问题。
        ///
        /// 何时将 autoReleaseSmtp 设置为false，就是SmtpClient需要重复使用的情况，即需要使用“相同MailHelper”向“相同Smtp服务器”发送大批量的邮件时。
        /// </summary>
        bool m_autoDisposeSmtp = false;

        // 附件集合
        Collection<Attachment> Attachments { get; set; }

        // 指定一个电子邮件不同格式显示的副本。
        Collection<AlternateView> AlternateViews { get; set; }

        #endregion


        #region  计划邮件数量 和 已执行完成邮件数量

        // 记录和获取在大批量执行异步短信发送时已经处理了多少条记录
        // 1、根据此值手动或自动释放 SmtpClient .实际上没有需要根据此值进行手动释放，因为完全可以用自动释放替换此逻辑
        // 2、根据此值可以自己设置进度
        long m_CompletedSendCount = 0;
        public long CompletedSendCount
        {
            get { return Interlocked.Read(ref m_CompletedSendCount); }
            private set { Interlocked.Exchange(ref m_CompletedSendCount, value); }
        }

        // 计划邮件数量
        long m_PrepareSendCount = 0;
        public long PrepareSendCount
        {
            get { return Interlocked.Read(ref m_PrepareSendCount); }
            private set { Interlocked.Exchange(ref m_PrepareSendCount, value); }
        }

        #endregion

        #region 异步 发送邮件相关参数

        // 案例：因为异步发送邮件在SmtpClient处必须加锁保证一封一封的发送。
        // 这样阻塞了主线程。所以换用队列的方式以无阻塞的方式进行异步发送大批量邮件

        // 发送任务可能很长，所以使用 Thread 而不是用ThreadPool。（避免长时间暂居线程池线程），并且SmtpClient只支持一次一封邮件发送
        Thread m_SendMailThread = null;

        AutoResetEvent m_AutoResetEvent = null;
        AutoResetEvent AutoResetEvent
        {
            get
            {
                if (m_AutoResetEvent == null)
                    m_AutoResetEvent = new AutoResetEvent(true);
                return m_AutoResetEvent;
            }
        }

        // 待发送队列缓存数量。单独开个计数是为了提高获取此计数的效率
        int m_messageQueueCount = 0;
        // 因为 MessageQueue 可能在 m_SendMailThread 线程中进行出队操作,所以使用并发队列ConcurrentQueue.
        // 队列中的数据只能通过取消异步发送进行清空，或则就会每一元素都执行发送邮件
        private ConcurrentQueue<MailUserState> m_MessageQueue = null;
        private ConcurrentQueue<MailUserState> MessageQueue
        {
            get
            {
                if (m_MessageQueue == null)
                    m_MessageQueue = new ConcurrentQueue<MailUserState>();
                return m_MessageQueue;
            }
        }

        /// <summary>
        /// 在执行异步发送时传递的对象，用于传递给异步发生完成时调用的方法 OnSendCompleted 。
        /// </summary>
        public object AsycUserState { get; set; }
        #endregion



        #region SmtpClient 相关方法

        /// <summary>
        /// 检查此 MailHelper 实例是否已经设置了 SmtpClient
        /// </summary>
        /// <returns>true代表已设置</returns>
        public bool ExistsSmtpClient()
        {
            return m_SmtpClient != null;
        }

        /// <summary>
        /// 设置 SmtpClient 实例 和是否自动释放Smtp的唯一入口
        /// 1、将内部 计划数量 和 已完成数量 清零，重新统计以便自动释放SmtpClient
        /// 2、若要对SmtpClent设置SendCompleted事件，请在调用此方法前进行设置
        /// </summary>
        /// <param name="mSmtpClient"> SmtpClient 实例</param>
        /// <param name="autoReleaseSmtp">设置在 MailHelper 类内部，发送完邮件后是否自动释放 SmtpClient 实例</param>
        public void SetSmtpClient(SmtpClient mSmtpClient, bool autoReleaseSmtp)
        {
            //#if DEBUG
            //            Debug.WriteLine("设置SmtpClient,自动释放为" + (autoReleaseSmtp ? "TRUE" : "FALSE"));
            //#endif
            m_SmtpClient = mSmtpClient;
            m_autoDisposeSmtp = autoReleaseSmtp;

            // 将内部 计划数量 和 已完成数量 清零，重新统计以便自动释放SmtpClient  (MailHelper实例唯一的清零地方)
            //m_PrepareSendCount = 0;
            m_CompletedSendCount = 0;

            // 注册回调事件.释放对象---该事件不进行取消注册，只在释放SmtpClient时，一起释放   （所以SmtpClient与MailHelper绑定后，就不要再单独使用了）
            m_SmtpClient.SendCompleted += SendCompleted4Dispose;
        }

        /// <summary>
        /// 释放 SmtpClient
        /// </summary>
        public void ManualDisposeSmtp()
        {
            this.InnerDisposeSmtp();
        }

        /// <summary>
        /// 释放SmtpClient
        /// </summary>
        void AutoDisposeSmtp()
        {
            if (m_autoDisposeSmtp && m_SmtpClient != null)
            {
                if (PrepareSendCount == 0)
                {
                    // PrepareSendCount=0 说明还未设置计划批量邮件数，所以不自动释放SmtpClient。
                    // 不能因为小于CompletedSendCount就报错，因为可能是先发送再设置计划邮件数量
                }
                else if (PrepareSendCount < CompletedSendCount)
                {
                    throw new Exception(MailValidatorHelper.EMAIL_ASYNC_SEND_PREPARE_ERROR);
                }
                else if (PrepareSendCount == CompletedSendCount)
                {
                    InnerDisposeSmtp();
                }
            }
            else
            {
                // 不清空和Dispose()内部的SmtpClient字段，即用在需要重复使用时不需要再调用 SetSmtpClient() 进行设置。
            }
        }

        /// <summary>
        /// 释放SmtpClient
        /// </summary>
        void InnerDisposeSmtp()
        {
            if (m_SmtpClient != null)
            {
                //#if DEBUG
                //                Debug.WriteLine("释放SMtpClient");
                //#endif
                m_SmtpClient.Dispose();
                m_SmtpClient = null;

                // 在设置 SmtpClient 入口处重新进行设置
                m_autoDisposeSmtp = false;

                PrepareSendCount = 0;
                CompletedSendCount = 0;
            }
        }

        /// <summary>
        /// 根据默认值创建一个链接
        /// </summary>
        public SmtpClient CreatSmtpClient()
        {
            var smtp = new SmtpClient(SmtpClientAdd, Port);
            smtp.EnableSsl = false;
            smtp.UseDefaultCredentials = false;
            smtp.DeliveryMethod = SmtpDeliveryMethod.Network;
            smtp.Credentials = new NetworkCredential(User, Password);
            smtp.Timeout = 100000;
            return smtp;
        }

        #endregion

        #region 发送邮件 相关方法

        /// <summary>
        /// 计划批量发送邮件的个数，配合自动释放SmtpClient。（批量邮件发送不调用此方法就不会自动释放SmtpClient）
        /// 0、此方法可以在发送邮件方法之前或之后调用
        /// 1、只有设置后才会自动根据 m_autoDisposeSmtp 字段进行释放SmtpClient。
        /// 2、若 m_autoDisposeSmtp = false 即由自己手动进行设置的无需调用此方法设置预计邮件数
        /// </summary>
        /// <param name="preCount">计划邮件数量</param>
        public void SetBatchMailCount(long preCount)
        {
            PrepareSendCount = preCount;

            if (preCount < CompletedSendCount)
            {
                throw new ArgumentOutOfRangeException(nameof(preCount), MailValidatorHelper.EMAIL_ASYNC_SEND_PREPARE_ERROR);
            }
            if (preCount == CompletedSendCount)
            {
                if (m_autoDisposeSmtp)
                    InnerDisposeSmtp();
            }
        }

        /// <summary>
        /// 设置邮件基本信息
        /// </summary>
        /// <param name="subject">邮件标题</param>
        /// <param name="body">邮件正文</param>
        /// <param name="to">收件列表(;分隔)</param>
        /// <param name="cc">抄送列表(;分隔)</param>
        /// <param name="bcc">密送列表(;分隔)</param>
        /// <param name="isbodyhtml">邮件内容是否html格式</param>
        /// <param name="priority">邮件优先级 0-Normal   1-Low   2-High</param>
        public void SetMailInfo(string subject, string body, string to, string cc = null, string bcc = null, string from = null, string fromName = null, bool isbodyhtml = true, int priority = 0)
        {
            Subject = subject;
            Body = body;
            To = to;
            Cc = cc ?? "";
            Bcc = bcc ?? "";
            From = from ?? "";
            FromDisplayName = fromName ?? "";
            IsBodyHtml = isbodyhtml;
            Priority = priority;
        }

        /// <summary>
        /// 发送一封邮件
        /// </summary>
        /// <param name="subject">邮件标题</param>
        /// <param name="body">邮件正文</param>
        /// <param name="to">收件列表(;分隔)</param>
        /// <param name="cc">抄送列表(;分隔)</param>
        /// <param name="bcc">密送列表(;分隔)</param>
        public void SendMail(string subject, string body, string to, string cc = null, string bcc = null, string from = null, string fromName = null)
        {
            SetMailInfo(subject, body, to, cc, bcc, from, fromName);
            m_PrepareSendCount = 1;
            SendMail();
        }

        /// <summary>
        /// 发送邮件
        /// </summary>
        /// <param name="smtpClient">smtp链接，如不设置则自动根据默认值创建一个</param>
        /// <param name="autoReleaseSmtp">是否自动释放smtp链接</param>
        public void SendMail(SmtpClient smtpClient = null, bool autoReleaseSmtp = true)
        {
            SetSmtpClient(smtpClient ?? CreatSmtpClient(), autoReleaseSmtp);
            InnerSendMessage();
        }

        /// <summary>
        /// 取消异步邮件发送
        /// </summary>
        public void SendAsyncCancel()
        {
            // 因为此类为非线程安全类，所以 SendAsyncCancel 和发送邮件方法中操作MessageQueue部分的代码肯定是串行化的。
            // 所以不存在一边入队，一边出队导致无法完全取消所有邮件发送

            // 1、清空队列。
            // 2、取消正在异步发送的mail。
            // 3、设置计划数量=完成数量
            // 4、执行 AutoDisposeSmtp()

            // 1、清空队列。
            MailUserState tempMailUserState = null;
            while (MessageQueue.TryDequeue(out tempMailUserState))
            {
                Interlocked.Decrement(ref m_messageQueueCount);
                MailMessage message = tempMailUserState.CurMailMessage;
                InnerDisposeMessage(message);
            }
            tempMailUserState = null;
            // 2、取消正在异步发送的mail。
            m_SmtpClient.SendAsyncCancel();
            // 3、设置计划数量=完成数量
            PrepareSendCount = CompletedSendCount;
            // 4、执行 AutoDisposeSmtp()
            AutoDisposeSmtp();
        }

        /// <summary>
        /// 发送Email
        /// </summary>
        void InnerSendMessage()
        {

            bool hasError = false;

            var msg = CheckSendMail();
            if (msg.Count > 0)
            {
                throw new FormatException(MailInfoHelper.GetMailInfoStr(msg));
            }
            var mMailMessage = CreateMessagex();

            if (!hasError)
            {
                if (m_IsAsync)
                {
                    #region 异步发送邮件

                    if (PrepareSendCount == 1)
                    {
                        // 情况一：不重用 SmtpClient 实例会将PrepareSendCount设置为1
                        // 情况二：计划发送只有一条

                        // PrepareSendCount 是发送单条邮件。
                        MailUserState state = new MailUserState()
                        {
                            AutoReleaseSmtp = m_autoDisposeSmtp,
                            CurMailMessage = mMailMessage,
                            CurSmtpClient = m_SmtpClient,
                            IsSmpleMail = true,
                            UserState = AsycUserState,
                        };
                        if (m_autoDisposeSmtp)
                            // 由发送完成回调函数根据 IsSmpleMail 字段进行释放
                            m_SmtpClient = null;

                        ThreadPool.QueueUserWorkItem((userState) =>
                        {
                            // 无需 catch 发送异常，因为是异步，所以这里 catch 不到。
                            MailUserState curUserState = userState as MailUserState;
                            curUserState.CurSmtpClient.SendAsync(mMailMessage, userState);
                        }, state);

                    }
                    else
                    {
                        // 情况一：重用 SmtpClient 逻辑，即我们可以直接操作全局的 m_SmtpClient
                        // 情况二：批量发送邮件 PrepareSendCount>1
                        // 情况三：PrepareSendCount 还未设置，为0。比如场景在循环中做些判断，再决定发邮件，循环完才调用 SetBatchMailCount 设置计划邮件数量

                        MailUserState state = new MailUserState()
                        {
                            AutoReleaseSmtp = m_autoDisposeSmtp,
                            CurMailMessage = mMailMessage,
                            CurSmtpClient = m_SmtpClient,
                            UserState = AsycUserState,
                        };

                        MessageQueue.Enqueue(state);
                        Interlocked.Increment(ref m_messageQueueCount);

                        if (m_SendMailThread == null)
                        {
                            m_SendMailThread = new Thread(() =>
                            {
                                // noItemCount 次获取不到元素，就抛出线程异常
                                int noItemCount = 0;
                                while (true)
                                {
                                    if (PrepareSendCount != 0 && PrepareSendCount == CompletedSendCount)
                                    {
                                        // 已执行完毕。
                                        this.AutoDisposeSmtp();
                                        break;
                                    }
                                    else
                                    {
                                        MailUserState curUserState = null;

                                        if (!MessageQueue.IsEmpty)
                                        {
                                            //#if DEBUG
                                            //                                            Debug.WriteLine("WaitOne" + Thread.CurrentThread.ManagedThreadId);
                                            //#endif
                                            // 当执行异步取消时，会清空MessageQueue，所以 WaitOne 必须在从MessageQueue中取到元素之前
                                            AutoResetEvent.WaitOne();

                                            if (MessageQueue.TryDequeue(out curUserState))
                                            {
                                                Interlocked.Decrement(ref m_messageQueueCount);
                                                m_SmtpClient.SendAsync(curUserState.CurMailMessage, curUserState);
                                            }
                                        }
                                        else
                                        {
                                            if (noItemCount >= 10)
                                            {
                                                // 没有正确设置 PrepareSendCount 值。导致已没有邮件但此线程出现死循环
                                                this.InnerDisposeSmtp();

                                                throw new Exception(MailValidatorHelper.EMAIL_PREPARESENDCOUNT_NOTSET_ERROR);
                                            }

                                            Thread.Sleep(1000);
                                            noItemCount++;
                                        }
                                    }
                                    // SmtpClient 为null表示异步预计发送邮件数已经发送完，在 OnSendCompleted 进行了 m_SmtpClient 释放
                                    if (m_SmtpClient == null)
                                        break;
                                }

                                m_SendMailThread = null;
                            });
                            m_SendMailThread.Start();
                        }
                    }

                    #endregion
                }
                else
                {
                    #region 同步发送邮件
                    try
                    {
                        m_SmtpClient.Send(mMailMessage);
                        m_CompletedSendCount++;
                    }
                    catch (ObjectDisposedException smtpDisposedEx)
                    {
                        throw smtpDisposedEx;
                    }
                    catch (InvalidOperationException smtpOperationEx)
                    {
                        throw smtpOperationEx;
                    }
                    catch (SmtpFailedRecipientsException smtpFailedRecipientsEx)
                    {
                        throw smtpFailedRecipientsEx;
                    }
                    catch (SmtpException smtpEx)
                    {
                        throw smtpEx;
                    }
                    finally
                    {
                        if (mMailMessage != null)
                        {
                            InnerDisposeMessage(mMailMessage);
                            mMailMessage = null;
                        }
                        AutoDisposeSmtp();
                    }
                    #endregion
                }
            }
        }

        void InnerSendMessageSync()
        {
            var msg = CheckSendMail();
            if (msg.Count > 0)
            {
                throw new FormatException(MailInfoHelper.GetMailInfoStr(msg));
            }
            try
            {
                var mMailMessage = CreateMessagex();
                #region 异步发送邮件

                if (PrepareSendCount == 1)
                {
                    // 情况一：不重用 SmtpClient 实例会将PrepareSendCount设置为1
                    // 情况二：计划发送只有一条
                    m_SmtpClient.SendAsync(mMailMessage, AsycUserState);
                    if (m_autoDisposeSmtp)
                    {
                        m_SmtpClient = null;
                    }
                }
                else
                {
                    // 情况一：重用 SmtpClient 逻辑，即我们可以直接操作全局的 m_SmtpClient
                    // 情况二：批量发送邮件 PrepareSendCount>1
                    // 情况三：PrepareSendCount 还未设置，为0。比如场景在循环中做些判断，再决定发邮件，循环完才调用 SetBatchMailCount 设置计划邮件数量

                    MailUserState state = new MailUserState()
                    {
                        AutoReleaseSmtp = m_autoDisposeSmtp,
                        CurMailMessage = mMailMessage,
                        CurSmtpClient = m_SmtpClient,
                        UserState = AsycUserState,
                    };

                    MessageQueue.Enqueue(state);
                    Interlocked.Increment(ref m_messageQueueCount);

                    if (m_SendMailThread == null)
                    {
                        m_SendMailThread = new Thread(() =>
                        {
                            // noItemCount 次获取不到元素，就抛出线程异常
                            int noItemCount = 0;
                            while (true)
                            {
                                if (PrepareSendCount != 0 && PrepareSendCount == CompletedSendCount)
                                {
                                    // 已执行完毕。
                                    this.AutoDisposeSmtp();
                                    break;
                                }
                                else
                                {
                                    MailUserState curUserState = null;

                                    if (!MessageQueue.IsEmpty)
                                    {
                                        //#if DEBUG
                                        //                                            Debug.WriteLine("WaitOne" + Thread.CurrentThread.ManagedThreadId);
                                        //#endif
                                        // 当执行异步取消时，会清空MessageQueue，所以 WaitOne 必须在从MessageQueue中取到元素之前
                                        AutoResetEvent.WaitOne();

                                        if (MessageQueue.TryDequeue(out curUserState))
                                        {
                                            Interlocked.Decrement(ref m_messageQueueCount);
                                            m_SmtpClient.SendAsync(curUserState.CurMailMessage, curUserState);
                                        }
                                    }
                                    else
                                    {
                                        if (noItemCount >= 10)
                                        {
                                            // 没有正确设置 PrepareSendCount 值。导致已没有邮件但此线程出现死循环
                                            this.InnerDisposeSmtp();

                                            throw new Exception(MailValidatorHelper.EMAIL_PREPARESENDCOUNT_NOTSET_ERROR);
                                        }

                                        Thread.Sleep(1000);
                                        noItemCount++;
                                    }
                                }
                                // SmtpClient 为null表示异步预计发送邮件数已经发送完，在 OnSendCompleted 进行了 m_SmtpClient 释放
                                if (m_SmtpClient == null)
                                    break;
                            }

                            m_SendMailThread = null;
                        });
                        m_SendMailThread.Start();
                    }
                }

                #endregion
            }
            catch (Exception e)
            {
                throw new FormatException(e.Message);
            }
        }



        MailMessage CreateMessagex()
        {
            #region 构建 MailMessage
            bool hasError = false;
            var mMailMessage = new MailMessage();
            try
            {
                mMailMessage.From = new MailAddress(From, FromDisplayName);
                //增加收件人地址等
                foreach (var item in To.Split(';'))//.Where(s => MailValidatorHelper.IsEmail(s)))
                {
                    mMailMessage.To.Add(item);
                }
                if (!string.IsNullOrEmpty(Cc))
                {
                    foreach (var item in Cc.Split(';').Where(s => MailValidatorHelper.IsEmail(s)))
                    {
                        mMailMessage.CC.Add(item);
                    }
                }
                if (!string.IsNullOrEmpty(Bcc))
                {
                    foreach (var item in Bcc.Split(';').Where(s => MailValidatorHelper.IsEmail(s)))
                    {
                        mMailMessage.Bcc.Add(item);
                    }
                }
                mMailMessage.Subject = Subject;
                mMailMessage.Body = Body;

                if (Attachments != null && Attachments.Count > 0)
                {
                    foreach (Attachment attachment in Attachments)
                        mMailMessage.Attachments.Add(attachment);
                }

                mMailMessage.SubjectEncoding = Encoding.UTF8;
                mMailMessage.BodyEncoding = Encoding.UTF8;
                // SmtpClient 的 Headers 中会根据 MailMessage 默认设置些值，所以应该为 UTF8 。
                mMailMessage.HeadersEncoding = Encoding.UTF8;
                mMailMessage.IsBodyHtml = IsBodyHtml;
                if (AlternateViews != null && AlternateViews.Count > 0)
                {
                    foreach (AlternateView alternateView in AlternateViews)
                    {
                        mMailMessage.AlternateViews.Add(alternateView);
                    }
                }
                mMailMessage.Priority = (MailPriority)Priority;
            }
            catch (ArgumentNullException argumentNullEx)
            {
                hasError = true;
                throw argumentNullEx;
            }
            catch (ArgumentException argumentEx)
            {
                hasError = true;
                throw argumentEx;
            }
            catch (FormatException formatEx)
            {
                hasError = true;
                throw formatEx;
            }
            finally
            {
                if (hasError)
                {
                    if (mMailMessage != null)
                    {
                        this.InnerDisposeMessage(mMailMessage);
                        mMailMessage = null;
                    }
                    this.InnerDisposeSmtp();
                }
            }
            return mMailMessage;
            #endregion
        }

        /// <summary>
        /// 释放 MailMessage 对象
        /// </summary>
        void InnerDisposeMessage(MailMessage message)
        {
            if (message != null)
            {
                if (message.AlternateViews.Count > 0)
                {
                    message.AlternateViews.Dispose();
                }

                message.Dispose();
                message = null;
            }
        }

        /// <summary>
        /// 声明在 SmtpClient.SendAsync() 执行完后释放相关对象的回调方法   最后触发的委托
        /// </summary>
        protected void SendCompleted4Dispose(object sender, AsyncCompletedEventArgs e)
        {
            MailUserState state = e.UserState as MailUserState;

            if (state.CurMailMessage != null)
            {
                MailMessage message = state.CurMailMessage;
                this.InnerDisposeMessage(message);
                state.CurMailMessage = null;
            }

            if (state.IsSmpleMail)
            {
                if (state.AutoReleaseSmtp && state.CurSmtpClient != null)
                {
                    //#if DEBUG
                    //                    Debug.WriteLine("释放SmtpClient");
                    //#endif
                    state.CurSmtpClient.Dispose();
                    state.CurSmtpClient = null;
                }
            }
            else
            {
                if (!e.Cancelled)   // 取消的就不计数
                    CompletedSendCount++;

                if (state.AutoReleaseSmtp)
                {
                    AutoDisposeSmtp();
                }

                // 若批量异步发送，需要设置信号
                //#if DEBUG
                //                Debug.WriteLine("Set" + Thread.CurrentThread.ManagedThreadId);
                //#endif
                AutoResetEvent.Set();
            }

            // 先释放资源，处理错误逻辑
            if (e.Error != null && !state.IsErrorHandle)
            {
                throw e.Error;
            }
        }

        #endregion

        #region 异步发送邮件，MessageQueue队列中缓冲的待发邮件数量，使用者可根据此数量来限制邮件数量，以免内存浪费

        /// <summary>
        /// 获取异步发送邮件，MessageQueue队列中缓冲的待发邮件数量
        /// （使用者可根据此数量来限制邮件数量，以免内存浪费）
        /// </summary>
        public int GetAwaitMailCountAsync()
        {
            return Thread.VolatileRead(ref m_messageQueueCount);
        }

        #endregion

        #region 发送邮件前检查 相关方法

        /// <summary>
        /// 发送邮件前检查需要设置的信息是否完整，收集（提示+错误）信息
        /// </summary>
        public Dictionary<MailInfoType, string> CheckSendMail()
        {
            Dictionary<MailInfoType, string> dicMsg = new Dictionary<MailInfoType, string>();

            InnerCheckSendMail4Info(dicMsg);
            InnerCheckSendMail4Error(dicMsg);

            return dicMsg;
        }

        /// <summary>
        /// 发送邮件前检查需要设置的信息是否完整，收集 提示 信息
        /// </summary>
        /// <param name="dicMsg">将检查信息收集到此集合</param>
        void InnerCheckSendMail4Info(Dictionary<MailInfoType, string> dicMsg)
        {

            #region 处理 抄送/密送

            Cc = Cc ?? "";
            Bcc = Bcc ?? "";

            #endregion

            // 邮件主题
            if (Subject.Length == 0)
                dicMsg.Add(MailInfoType.SubjectEmpty, MailInfoHelper.GetMailInfoStr(MailInfoType.SubjectEmpty));

            // 邮件内容
            if (Body.Length == 0 &&
                (Attachments == null || (Attachments != null && Attachments.Count == 0))
                )
            {
                dicMsg.Add(MailInfoType.BodyEmpty, MailInfoHelper.GetMailInfoStr(MailInfoType.BodyEmpty));
            }
        }

        /// <summary>
        /// 发送邮件前检查需要设置的信息是否完整，收集 错误 信息
        /// </summary>
        /// <param name="dicMsg">将检查信息收集到此集合</param>
        void InnerCheckSendMail4Error(Dictionary<MailInfoType, string> dicMsg)
        {
            #region 处理 发件/收件

            if (string.IsNullOrEmpty(From))
            {
                From = "h3bpm@renolit.com.cn";
                FromDisplayName = "h3bpm";
            }
            else
            {
                if (!MailValidatorHelper.IsEmail(From))
                {
                    string strTemp = string.Format(MailInfoHelper.GetMailInfoStr(MailInfoType.FromFormat), FromDisplayName, From);
                    dicMsg.Add(MailInfoType.FromFormat, strTemp);
                }
            }

            if (string.IsNullOrEmpty(To))
            {
                dicMsg.Add(MailInfoType.ToEmpty, MailInfoHelper.GetMailInfoStr(MailInfoType.ToEmpty));
            }
            else
            {
                foreach (var item in To.Split(';'))
                {
                    if (!MailValidatorHelper.IsEmail(item))
                    {
                        string strTemp = string.Format(MailInfoHelper.GetMailInfoStr(MailInfoType.ToFormat), item, item);
                        dicMsg.Add(MailInfoType.ToFormat, strTemp);
                    }
                }
            }
            #endregion

            // SmtpClient 实例未设置
            if (m_SmtpClient == null)
                dicMsg.Add(MailInfoType.SmtpClientEmpty, MailInfoHelper.GetMailInfoStr(MailInfoType.SmtpClientEmpty));
            else
            {
                // SMTP 主服务器设置  （默认端口为25）
                if (m_SmtpClient.Host.Length == 0)
                    dicMsg.Add(MailInfoType.HostEmpty, MailInfoHelper.GetMailInfoStr(MailInfoType.HostEmpty));
                // SMPT 凭证
                if (m_SmtpClient.EnableSsl && m_SmtpClient.ClientCertificates.Count == 0)
                    dicMsg.Add(MailInfoType.CertificateEmpty, MailInfoHelper.GetMailInfoStr(MailInfoType.CertificateEmpty));
            }
        }
        #endregion


    }
}
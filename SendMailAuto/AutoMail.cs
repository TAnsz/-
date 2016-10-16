using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SendMailAuto
{
    public class AutoMail
    {
        #region props

        /// <summary>
        /// ID
        /// </summary>
        public long ID { get; set; }


        /// <summary>
        /// 发件人
        /// </summary>
        public string Sender { get; set; }


        /// <summary>
        /// 收件邮箱地址
        /// </summary>
        public string ToMailAddr { get; set; }


        /// <summary>
        /// 抄送邮箱地址
        /// </summary>
        public string CcMailAddr { get; set; }


        /// <summary>
        /// 密送邮箱地址
        /// </summary>
        public string BccMailAddr { get; set; }


        /// <summary>
        /// 邮件主题
        /// </summary>
        public string Subject { get; set; }


        /// <summary>
        /// 邮件正文
        /// </summary>
        public object BodyPart { get; set; }


        /// <summary>
        /// 创建人
        /// </summary>
        public string Creator { get; set; }


        /// <summary>
        /// 创建时间
        /// </summary>
        public DateTime? CreateDate { get; set; }


        /// <summary>
        /// 发送状态
        /// </summary>
        public string Status { get; set; }


        /// <summary>
        /// 发送时间
        /// </summary>
        public DateTime? SendTime { get; set; }


        /// <summary>
        /// 发送结果
        /// </summary>
        public string SendMsg { get; set; }



        #endregion

    }
}

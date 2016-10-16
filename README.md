# MailHelpers
发送邮件的类，支持同步和异步发送
##快速使用
###使用nuget安装
```
Install-Package MailHelpers
```
###简单发送邮件
```C#
try
{
 var mail = new MailHelper(false,"Ip地址",25, "用户名", "密码"){
                    Subject = "主题",
                    Body = "正文",
                    To = "收件人",
                    From="发件人",
                    FromDisplayName="发件人名称"
 };
 mail.SendMail();
 }
 catch (Exception e)
       {
       //获取错误信息
         }
```
  通过catch获取邮件发送异常信息
###发送多封邮件
 多封邮件如果需要共用smtp连接,需要在使用时设定邮件总数，在发送完成以后会自动释放smtp连接
```C#
try
            {
                var mail = new MailHelper(false, "192.168.3.31", 25, "h3bpm", "Bpm3H287")
                {
                    Subject = "主题",
                    Body = "正文",
                    To = "收件人",
                    From = "发件人",
                    FromDisplayName = "发件人名称"
                };
                //设定传送邮件总数
                mail.SetBatchMailCount(2);
                mail.SendMail();
                mail.SetMailInfo("主题", "正文","收件人");
                mail.SendMail();
            }
            catch (Exception)
            {
                throw;
            }
```
###异步发送邮件
 实例化的时候设定异步标志 `asnyc` 为`true`,发送完成后需要使用回调的方式获取完成状态和发送完成以后执行的方法。
 通过设定`AsyncCallback`和`AsycUserState`来实现完成状态检测。
 ```C#
 try
            {
                var mail = new MailHelper(true, "192.168.3.31", 25, "h3bpm", "Bpm3H287")
                {
                    Subject = "主题",
                    Body = "正文",
                    To = "收件人",
                    From = "发件人",
                    FromDisplayName = "发件人名称"
                };
                mail.AsycUserState = mail.Subject;
                mail.AsyncCallback = (send, arg) =>
                 {
                     var subject = arg.UserState.ToString();
                     if (arg.Error == null)
                     {
                         var msg = $"主题为[{subject}]的邮件发送成功！";
                     }
                     else
                     {
                         var msg = $"主题为[{subject}]的邮件发送失败！异常：{arg.Error.Message}";
                     }
                 };
                mail.SendMail();
            }
            catch (Exception)
            {
                throw;
            }
 ```

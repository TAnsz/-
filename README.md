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
###发送 多封邮件  
 多封邮件如果需要共用smtp连接需要在使用是
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
            }
            catch (Exception)
            {
                throw;
            }
```

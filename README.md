
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
 var mail = new MailHelper(asnyc,"Ip地址",25, "用户名", "密码"){
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
###发送 封邮件  

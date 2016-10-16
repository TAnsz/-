using MahApps.Metro.Controls;
using System;
using System.Windows;
using System.Windows.Controls;

namespace ServiceManager
{
    /// <summary>
    /// MainWindow.xaml 的交互逻辑
    /// </summary>
    public partial class MainWindow : MetroWindow
    {
        string strServiceName = string.Empty;
        public MainWindow()
        {
            InitializeComponent();
            strServiceName = string.IsNullOrEmpty(txtServiceName.Text) ? "SendMailService" : txtServiceName.Text;
            InitControlStatus(strServiceName, btnInstallOrUninstall, btnStartOrEnd, btnGetStatus, rb_Msg);
            //txtServiceName.TextChanged += (o, e) =>
            //{
            //    strServiceName = txtServiceName.Text;
            //    if (!string.IsNullOrEmpty(strServiceName))
            //        InitControlStatus(strServiceName, btnInstallOrUninstall, btnStartOrEnd, btnGetStatus, rb_Msg);
            //};
        }

        /// <summary>
        /// 初始化控件状态
        /// </summary>
        /// <param name="serviceName">服务名称</param>
        /// <param name="btn1">安装按钮</param>
        /// <param name="btn2">启动按钮</param>
        /// <param name="btn3">获取状态按钮</param>
        /// <param name="rb">提示信息</param>
        void InitControlStatus(string serviceName, Button btn1, Button btn2, Button btn3, RichTextBox rb)
        {
            try
            {
                btn1.IsEnabled = true;

                if (ServiceAPI.isServiceIsExisted(serviceName))
                {
                    btn3.IsEnabled = true;
                    btn2.IsEnabled = true;
                    btn1.Content = "卸载服务";
                    rb.AppendText("服务【" + serviceName + "】已安装！");
                    int status = ServiceAPI.GetServiceStatus(serviceName);
                    if (status == 4)
                    {
                        btn2.Content = "停止服务";
                    }
                    else
                    {
                        btn2.Content = "启动服务";
                    }
                    GetServiceStatus(serviceName, txtStaute);
                    //获取服务版本
                    string temp = string.IsNullOrEmpty(ServiceAPI.GetServiceVersion(serviceName)) ? string.Empty : "(" + ServiceAPI.GetServiceVersion(serviceName) + ")\r\n";
                    rb.AppendText(temp);
                }
                else
                {
                    btn1.Content = "安装服务";
                    btn2.IsEnabled = false;
                    btn3.IsEnabled = false;
                    rb.AppendText("服务【" + serviceName + "】未安装！\r\n");
                }
            }
            catch (Exception ex)
            {
                rb.AppendText("error\r\n");
                LogAPI.WriteLog(ex.Message);
            }
        }

        /// <summary>
        /// 安装或卸载服务
        /// </summary>
        /// <param name="serviceName">服务名称</param>
        /// <param name="btnSet">安装、卸载</param>
        /// <param name="btnOn">启动、停止</param>
        /// <param name="rb">提示信息</param>
        void SetServerce(string serviceName, Button btnSet, Button btnOn, Button btnShow, RichTextBox rb)
        {
            try
            {
                string location = System.Reflection.Assembly.GetExecutingAssembly().Location;
                string serviceFileName = location.Substring(0, location.LastIndexOf('\\')) + "\\" + serviceName + ".exe";

                if (btnSet.Content.Equals("安装服务"))
                {
                    ServiceAPI.InstallmyService(null, serviceFileName);
                    if (ServiceAPI.isServiceIsExisted(serviceName))
                    {
                        rb.AppendText("服务【" + serviceName + "】安装成功！");
                        btnOn.IsEnabled = btnShow.IsEnabled = true;
                        string temp = string.IsNullOrEmpty(ServiceAPI.GetServiceVersion(serviceName)) ? string.Empty : "(" + ServiceAPI.GetServiceVersion(serviceName) + ")";
                        rb.AppendText(temp + "\r\n");
                        btnSet.Content = "卸载服务";
                        btnOn.Content = "启动服务";
                    }
                    else
                    {
                        rb.AppendText("服务【" + serviceName + "】安装失败，请检查日志！\r\n");
                    }
                }
                else
                {
                    ServiceAPI.UnInstallmyService(serviceFileName);
                    if (!ServiceAPI.isServiceIsExisted(serviceName))
                    {
                        rb.AppendText("服务【" + serviceName + "】卸载成功！\r\n");
                        btnOn.IsEnabled = btnShow.IsEnabled = false;
                        btnSet.Content = "安装服务";
                    }
                    else
                    {
                        rb.AppendText("服务【" + serviceName + "】卸载失败，请检查日志！\r\n");
                    }
                }
            }
            catch (Exception ex)
            {
                rb.AppendText("error\r\n");
                LogAPI.WriteLog(ex.Message);
            }
        }

        //获取服务状态
        void GetServiceStatus(string serviceName, TextBlock txtStatus)
        {
            try
            {
                if (ServiceAPI.isServiceIsExisted(serviceName))
                {
                    string statusStr = "";
                    int status = ServiceAPI.GetServiceStatus(serviceName);
                    switch (status)
                    {
                        case 1:
                            statusStr = "服务未运行！";
                            break;
                        case 2:
                            statusStr = "服务正在启动！";
                            break;
                        case 3:
                            statusStr = "服务正在停止！";
                            break;
                        case 4:
                            statusStr = "服务正在运行！";
                            break;
                        case 5:
                            statusStr = "服务即将继续！";
                            break;
                        case 6:
                            statusStr = "服务即将暂停！";
                            break;
                        case 7:
                            statusStr = "服务已暂停！";
                            break;
                        default:
                            statusStr = "未知状态！";
                            break;
                    }
                    txtStatus.Text = statusStr;
                }
                else
                {
                    txtStatus.Text = "服务【" + serviceName + "】未安装！";
                }
            }
            catch (Exception ex)
            {
                txtStatus.Text = "error";
                LogAPI.WriteLog(ex.Message);
            }
        }

        //启动服务
        void OnService(string serviceName, Button btn, RichTextBox rb)
        {
            try
            {
                if (btn.Content.Equals("启动服务"))
                {
                    ServiceAPI.RunService(serviceName);

                    int status = ServiceAPI.GetServiceStatus(serviceName);
                    if (status == 2 || status == 4 || status == 5)
                    {
                        rb.AppendText("服务【" + serviceName + "】启动成功！\r\n");
                        btn.Content = "停止服务";
                    }
                    else
                    {
                        rb.AppendText("服务【" + serviceName + "】启动失败！\r\n");
                    }
                }
                else
                {
                    ServiceAPI.StopService(serviceName);

                    int status = ServiceAPI.GetServiceStatus(serviceName);
                    if (status == 1 || status == 3 || status == 6 || status == 7)
                    {
                        rb.AppendText("服务【" + serviceName + "】停止成功！\r\n");
                        btn.Content = "启动服务";
                    }
                    else
                    {
                        rb.AppendText("服务【" + serviceName + "】停止失败！\r\n");
                    }
                }
            }
            catch (Exception ex)
            {
                rb.AppendText("error\r\n");
                LogAPI.WriteLog(ex.Message);
            }
        }

        private void btnInstallOrUninstall_Click(object sender, RoutedEventArgs e)
        {
            btnInstallOrUninstall.IsEnabled = false;
            SetServerce(strServiceName, btnInstallOrUninstall, btnStartOrEnd, btnGetStatus, rb_Msg);
            btnInstallOrUninstall.IsEnabled = true;
            btnInstallOrUninstall.Focus();
        }

        private void btnStartOrEnd_Click(object sender, RoutedEventArgs e)
        {

            btnStartOrEnd.IsEnabled = false;
            OnService(strServiceName, btnStartOrEnd, rb_Msg);
            btnStartOrEnd.IsEnabled = true;
            btnStartOrEnd.Focus();
        }

        private void btnGetStatus_Click(object sender, RoutedEventArgs e)
        {
            btnGetStatus.IsEnabled = false;
            GetServiceStatus(strServiceName, txtStaute);
            btnGetStatus.IsEnabled = true;
            btnGetStatus.Focus();
        }

        private void btn_file_Click(object sender, RoutedEventArgs e)
        {
            var openFileDialog = new Microsoft.Win32.OpenFileDialog()
            {
                Filter = "EXE Files(*.exe)|*.exe"
            };
            var result = openFileDialog.ShowDialog();
            if (result == true)
            {
                strServiceName = openFileDialog.SafeFileName.Substring(0, openFileDialog.SafeFileName.LastIndexOf('.'));
                txtServiceName.Text = strServiceName;
                InitControlStatus(strServiceName, btnInstallOrUninstall, btnStartOrEnd, btnGetStatus, rb_Msg);
            }
        }

    }
}

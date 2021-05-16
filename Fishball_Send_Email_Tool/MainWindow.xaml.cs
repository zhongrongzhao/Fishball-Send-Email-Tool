using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Mail;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace Fishball_Send_Email_Tool
{
    /// <summary>
    /// MainWindow.xaml 的交互逻辑
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        /// <summary>
        /// 发送邮件方法
        /// </summary>
        /// <param name="mailTo">接收人邮件</param>
        /// <param name="mailTitle">发送邮件标题</param>
        /// <param name="mailContent">发送邮件内容</param>
        /// <returns></returns>
        public  void SendEmail(string mailTo, string mailTitle, string mailContent)
        {
            string stmpServer = txt_Smtp.Text;      //smtp服务器地址
            string mailAccount = txt_MyEmailAddress.Text;       //邮箱账号
            IntPtr p = System.Runtime.InteropServices.Marshal.SecureStringToBSTR(this.txt_password.SecurePassword);
            string password = System.Runtime.InteropServices.Marshal.PtrToStringBSTR(p);
            string pwd = password;      //邮箱密码

            //邮件服务设置
            SmtpClient smtpClient = new SmtpClient();
            smtpClient.DeliveryMethod = SmtpDeliveryMethod.Network;     //指定电子邮件发送方式
            smtpClient.Host = stmpServer;   //指定发送方SMTP服务器
            smtpClient.EnableSsl = true;    //使用安全加密连接
            smtpClient.UseDefaultCredentials = true;    //不和请求一起发送
            smtpClient.Credentials = new NetworkCredential(mailAccount, pwd);   //设置发送账号密码

            MailMessage mailMessage = new MailMessage(mailAccount, mailTo);     //实例化邮件信息实体并设置发送方和接收方
            mailMessage.Subject = mailTitle;    //设置发送邮件得标题
            mailMessage.Body = mailContent;     //设置发送邮件内容
            mailMessage.BodyEncoding = Encoding.UTF8;   //设置发送邮件得编码
            mailMessage.IsBodyHtml = false;     //设置标题是否为HTML格式
            mailMessage.Priority = MailPriority.Normal;     //设置邮件发送优先级

            try
            {
                smtpClient.Send(mailMessage);       //发送邮件
                MessageBox.Show("发送成功","提示");
                return;
            }
            catch (SmtpException ex)
            {
                MessageBox.Show(""+ex);
            }
        }

        private void btn_sent_Click(object sender, RoutedEventArgs e)
        {
            SendEmail(txt_TaEmailAddress.Text, txt_label.Text, txt_content.Text);
        }
    }
}

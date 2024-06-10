using ImapX;
using Spire.Email;
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
using System.Windows.Shapes;

namespace WorrdExcelInterface
{
    /// <summary>
    /// Логика взаимодействия для MailInterface.xaml
    /// </summary>
    public partial class MailInterface : Window
    {
        public MailInterface()
        {
            InitializeComponent();
        }

        private void MessageSendButton_Click(object sender, RoutedEventArgs e)
        {
            System.Net.Mail.MailMessage message = new System.Net.Mail.MailMessage(From.Text, To.Text, Subject.Text, "WordFile");
            message.Attachments.Add(new System.Net.Mail.Attachment("WordMessage.docx")); 
            SmtpClient client = new SmtpClient("smtp.mail.ru");
            client.Credentials = new NetworkCredential(From.Text, PasswordTextBox.Text);
            client.EnableSsl = true;
            client.Send(message);
        }
    }
}

using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Net;
using System.Net.Mail;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace bank_calc
{
    public partial class EmailForm : Form
    {
        string fileName = "";
        public EmailForm()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            // отправитель - устанавливаем адрес и отображаемое в письме имя
            MailAddress from = new MailAddress(textBox3.Text);
            // кому отправляем
            MailAddress to = new MailAddress(textBox4.Text);
            // создаем объект сообщения
            MailMessage m = new MailMessage(from, to);
            if(!String.IsNullOrEmpty(fileName)) m.Attachments.Add(new Attachment(fileName));
            // тема письма
            m.Subject = textBox7.Text;
            // текст письма
            m.Body = textBox8.Text;
            // письмо представляет код html
            m.IsBodyHtml = true;
            // адрес smtp-сервера и порт, с которого будем отправлять письмо
            SmtpClient smtp = new SmtpClient(textBox1.Text, int.Parse(textBox2.Text));
            // логин и пароль
            smtp.Credentials = new NetworkCredential(textBox5.Text, textBox6.Text);
            smtp.EnableSsl = true;
            smtp.Send(m);
            
        }

        private void button2_Click(object sender, EventArgs e)
        {
            openFileDialog1.ShowDialog();
            if (openFileDialog1.ShowDialog() == DialogResult.Cancel) return;
            else fileName = openFileDialog1.FileName;
        }
    }
}

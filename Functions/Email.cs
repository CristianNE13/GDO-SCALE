using System.Net.Mail;
using System.Net;
using System;
using System.Drawing;
using System.Threading.Tasks;

namespace Scale_Program.Functions
{
    public class Email
    {
        private string _email { get; set; }
        private string _password { get; set; }
        private string _mensaje { get; set; }
        private string _asunto { get; set; }
        private string _destinatario { get; set; }

        private string _hostGmail = "smtp.gmail.com";
        private int _portGmail = 587;

        public Email(string email, string password, string body, string subject, string to)
        {
            _email = email;
            _password = password;
            _mensaje = body;
            _asunto = subject;
            _destinatario = to;
        } 

        public string SendEmail(string articulo)
        {
            MailMessage mailMessage = new MailMessage();
            mailMessage.From = new MailAddress(_email);
            mailMessage.To.Add(_destinatario);
            mailMessage.Subject = _asunto;
            mailMessage.Body = $"{_mensaje}\n\n" +
                               $"{DateTime.Now}.\nEl ultimo articulo aceptado fue: {articulo}.\n" +
                               $"La contraseña para reanudar el sistema es: 12345"
                ;

            SmtpClient smtpClient = new SmtpClient();
            smtpClient.Host = _hostGmail;
            smtpClient.Port = _portGmail;
            smtpClient.UseDefaultCredentials = false;
            smtpClient.Credentials = new NetworkCredential(_email, _password);
            smtpClient.EnableSsl = true;

            try
            {
                smtpClient.Send(mailMessage);
                return "Correo enviado: falla en la secuencia.";
            }
            catch (Exception ex)
            {
                return $"No se pudo enviar el correo: {ex.Message}";
            }
        }
    }
}
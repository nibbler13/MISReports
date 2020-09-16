using System.Net.Mail;
using System;
using System.Reflection;
using System.IO;
using System.Collections.Generic;
using System.Net.Mime;

namespace MISReports {
	public class SystemMail {
		public static void SendMail (string subject, string body, string receiver, string[] attachmentsPath = null, bool addSignature = true) {
			Logging.ToLog("Отправка сообщения, тема: " + subject + ", текст: " + body);
			Logging.ToLog("Получатели: " + receiver);

			if (string.IsNullOrEmpty(receiver))
				return;

			try {
				string appName = Assembly.GetExecutingAssembly().GetName().Name;

				MailAddress from = new MailAddress(
					Properties.Settings.Default.MailUser + "@" + Properties.Settings.Default.MailDomain, 
					appName);

				List<MailAddress> mailAddressesTo = new List<MailAddress>();

				if (receiver.Contains(";")) {
					string[] receivers = receiver.Split(';');
					foreach (string address in receivers)
						mailAddressesTo.Add(new MailAddress(address));
				} else
					mailAddressesTo.Add(new MailAddress(receiver));
				
				if (addSignature)
					body += Environment.NewLine + Environment.NewLine + 
						"___________________________________________" + Environment.NewLine +
						"Это автоматически сгенерированное сообщение" + Environment.NewLine + 
						"Просьба не отвечать на него" + Environment.NewLine +
 						"Имя системы: " + Environment.MachineName;

				MailMessage message = new MailMessage();

				foreach (MailAddress mailAddress in mailAddressesTo)
					message.To.Add(mailAddress);

				message.IsBodyHtml = body.Contains("<") && body.Contains(">");

				if (message.IsBodyHtml)
					body = body.Replace(Environment.NewLine, "<br>");
				
				if (attachmentsPath != null) 
					foreach (string attachmentPath in attachmentsPath) {
						if (string.IsNullOrEmpty(attachmentPath) || !File.Exists(attachmentPath))
							continue;

						Attachment attachment = new Attachment(attachmentPath, MediaTypeNames.Application.Octet);

						if (message.IsBodyHtml && attachmentPath.EndsWith(".jpg")) {
							attachment.ContentDisposition.Inline = true;

							LinkedResource inline = new LinkedResource(attachmentPath, MediaTypeNames.Image.Jpeg);
							inline.ContentId = Guid.NewGuid().ToString();

							body = body.Replace("Фотография с камеры терминала:", "Фотография с камеры терминала:<br>" +
								string.Format(@"<img src=""cid:{0}"" />", inline.ContentId));

							AlternateView avHtml = AlternateView.CreateAlternateViewFromString(body, null, MediaTypeNames.Text.Html);
							avHtml.LinkedResources.Add(inline);

							message.AlternateViews.Add(avHtml);
						} else
							message.Attachments.Add(attachment);
					}
				
				message.From = from;
				message.Subject = subject;
				message.Body = body;

				string mailCopy = Properties.Settings.Default.MailCopy;
				if (!string.IsNullOrEmpty(mailCopy)) {
					if (mailCopy.Contains(";")) {
						string[] splitted = mailCopy.Split(';');
						foreach (string item in splitted)
							message.CC.Add(item);
					} else
						message.CC.Add(mailCopy);
				}

				SmtpClient client = new SmtpClient(Properties.Settings.Default.MailSmtpServer, 587);
				client.UseDefaultCredentials = false;
				client.DeliveryMethod = SmtpDeliveryMethod.Network;
				client.EnableSsl = false;
				client.Credentials = new System.Net.NetworkCredential(
					Properties.Settings.Default.MailUser,
					Properties.Settings.Default.MailPassword);

				client.Send(message);
				client.Dispose();

				foreach (Attachment attach in message.Attachments)
					attach.Dispose();

				message.Dispose();
				Logging.ToLog("Письмо отправлено успешно");
			} catch (Exception e) {
				Logging.ToLog("SendMail exception: " + e.Message + Environment.NewLine + e.StackTrace);

				if (e.InnerException != null)
					Logging.ToLog("SendMail inner exception: " + e.InnerException.Message + Environment.NewLine + e.InnerException.StackTrace);
			}
		}
	}
}

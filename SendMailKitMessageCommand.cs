using System;
using System.Management.Automation;
using System.Management.Automation.Runspaces;
using MailKit.Net.Smtp;
using MailKit.Security;
using MailKit;
using MimeKit;
using MimeKit.Text;
using System.Text.RegularExpressions;
using System.Text;
using System.Web;
using System.IO;
using System.Runtime.InteropServices;

namespace mailpwshkit
{
    [Cmdlet(VerbsCommunications.Send,"MailKitMessage")]
    [OutputType(typeof(Boolean))]
    public class SendMailKitMessageCommand : PSCmdlet
    {
        [Parameter(
            Mandatory = true,
            Position = 0,
            ValueFromPipeline = true,
            ValueFromPipelineByPropertyName = true)]
        public String[] To { get; set; }


        [Parameter(
            Mandatory = true,
            Position = 1,
            ValueFromPipeline = true,
            ValueFromPipelineByPropertyName = true)]
        public String Subject { get; set; }

        [Parameter(
            Mandatory = true,
            Position = 2,
            ValueFromPipeline = true,
            ValueFromPipelineByPropertyName = true)]
        public String Body { get; set; }

        [Parameter()]
        public String SmtpServer { get; set; }

        [Parameter(
            Mandatory = true,
            Position = 3,
            ValueFromPipeline = true,
            ValueFromPipelineByPropertyName = true)]
        public String From { get; set; }

        [Parameter()]
        public String[] Attachments { get; set; }

        [Parameter()]
        public String[] Bcc { get; set; }

        [Parameter()]
        public SwitchParameter BodyAsHtml { get; set; }

        [Parameter(DontShow = true)]
        public Encoding Encoding { get; set; }

        [Parameter()]
        public ContentEncoding ContentEncoding { get; set; }

        [Parameter()]
        public String[] Cc { get; set; }

        //[Parameter()]
        //public DeliveryNotificationOption { get; set; }

        [Parameter()]
        public PSCredential Credential { get; set; }

        [Parameter()]
        public SwitchParameter UseSsl { get; set; }

        [Parameter()]
        public int Port { get; set; } = 25;

        [Parameter()]
        public MessageImportance Priority { get; set; } = MessageImportance.Normal;

        [Parameter()]
        public SwitchParameter SkipCertificateValidation { get; set; }

        [Parameter()]
        public SecureSocketOptions SecureSocketOptions { get; set; }

        protected override void ProcessRecord()
        {
            var message = new MimeMessage();
            MailboxAddress from_mb = new MailboxAddress("");
            if (MailboxAddress.TryParse(From, out from_mb))
            {
                message.From.Add(from_mb);
            }
            else
            {
                this.WriteError(
                    new ErrorRecord(
                        new Exception(String.Format("Could not parse {0}", From)),
                        "AddressParseFailure",
                        ErrorCategory.ParserError,
                        From));
            }
            foreach (String address in To)
            {
                MailboxAddress to_mb = new MailboxAddress("");
                if (MailboxAddress.TryParse(address, out to_mb)) { 
                    message.To.Add(to_mb);
                }
                else
                {
                    this.WriteError(
                        new ErrorRecord(
                            new Exception(String.Format("Could not parse {0}", address)), 
                            "AddressParseFailure",
                            ErrorCategory.ParserError,
                            address));
                }
            }
            foreach (String address in Cc)
            {
                MailboxAddress cc_mb = new MailboxAddress("");
                if (MailboxAddress.TryParse(address, out cc_mb))
                {
                    message.Cc.Add(cc_mb);
                }
                else
                {
                    this.WriteError(
                        new ErrorRecord(
                            new Exception(String.Format("Could not parse {0}", address)),
                            "AddressParseFailure",
                            ErrorCategory.ParserError,
                            address));
                }
            }

            foreach (String address in Bcc)
            {
                MailboxAddress bcc_mb = new MailboxAddress("");
                if (MailboxAddress.TryParse(address, out bcc_mb))
                {
                    message.Bcc.Add(bcc_mb);
                }
                else
                {
                    this.WriteError(
                        new ErrorRecord(
                            new Exception(String.Format("Could not parse {0}", address)),
                            "AddressParseFailure",
                            ErrorCategory.ParserError,
                            address));
                }
            }

            message.Subject = Subject;
            var body = new MultipartAlternative();
            if (BodyAsHtml.IsPresent)
            {
                Regex reg = new Regex("<[^>]+>", RegexOptions.IgnoreCase);
                var stripped = HttpUtility.HtmlDecode(reg.Replace(Body, ""));
                //add a plain text version for mail clients that can't display the html
                body.Add(new TextPart(TextFormat.Plain) { Text = stripped });
                //HTML goes second as its the more expressive version of the text.
                body.Add(new TextPart(TextFormat.Html) { Text = Body });
            }
            else
            {
                body.Add(new TextPart(TextFormat.Plain) { Text = Body });
            }
            if (Attachments is null) {
                message.Body = body;
            }
            else
            {
                var multipart = new Multipart("mixed");
                multipart.Add(body);
                foreach (var Attachment in Attachments)
                {
                    if (!File.Exists(Attachment)) { continue; }
                    //TODO: maybe allow user to specify mediatype for attachments? could allow image attachments instead.
                    var att = new MimePart("appplication", "octet-stream")
                    {
                        Content = new MimeContent(File.OpenRead(Attachment), ContentEncoding.Default),
                        ContentDisposition = new ContentDisposition(ContentDisposition.Attachment),
                        ContentTransferEncoding = ContentEncoding.Base64,
                        FileName = Path.GetFileName(Attachment)
                    };
                    multipart.Add(att);
                }
                message.Body = multipart;
            }

            using (var client = new SmtpClient())
            {

                if (SkipCertificateValidation.IsPresent)
                {
                    client.ServerCertificateValidationCallback = (s, c, h, e) => true;
                }

                if (!String.IsNullOrEmpty(SmtpServer))
                {
                    client.Connect(SmtpServer, Port, SecureSocketOptions);
                }
                if (!(Credential is null))
                {
                //securestring to string code from 
                // https://blogs.msdn.microsoft.com/fpintos/2009/06/12/how-to-properly-convert-securestring-to-string/
                    IntPtr unmanagedString = IntPtr.Zero;
                    try
                    {
                        unmanagedString = Marshal.SecureStringToGlobalAllocUnicode(Credential.Password);
                        client.Authenticate(Credential.UserName, Marshal.PtrToStringUni(unmanagedString));
                    }
                    finally
                    {
                        Marshal.ZeroFreeGlobalAllocUnicode(unmanagedString);
                    }
                    
                }

                client.Send(message);
                client.Disconnect(true);
            }
        }
        protected override void EndProcessing()
        {
        }
        protected override void BeginProcessing()
        {
        }

    }

}

using MimeKit;
using System;
using System.IO;
using System.Management.Automation;
using System.Reflection;
using System.Web;

namespace Send_MailKitMessage
{
    public class ModuleInitializer : IModuleAssemblyInitializer
    {
        public void OnImport()
        {
            //for some reason running Send-MailKitMessage in Windows PowerShell ALWAYS returned the following exception: "Could not load file or assembly 'System.Buffers, Version=4.0.2.0, Culture=neutral, PublicKeyToken=cc7b13ffcd2ddd51' or one of its dependencies. The system cannot find the file specified."
            //and I COULD NOT nail down what was causing it
            //so per https://devblogs.microsoft.com/powershell/resolving-powershell-module-assembly-dependency-conflicts/ I am using an AssemblyResolve event handler to create a dynamic binding redirect so all calls to System.Buffers use the same assembly
            AppDomain.CurrentDomain.AssemblyResolve += DependencyResolution.ResolveSystemBuffers;
        }
    }

    internal static class DependencyResolution
    {
        private static readonly string CurrentLocation = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);

        public static Assembly ResolveSystemBuffers(object sender, ResolveEventArgs args)
        {
            //parse the assembly name
            var assemblyName = new AssemblyName(args.Name);

            //only handle the dependency we care about
            if (!assemblyName.Name.Equals("System.Buffers"))
            {
                return null;
            }

            return Assembly.LoadFrom(Path.Combine(CurrentLocation, "System.Buffers.dll"));
        }
    }

    internal static class EmailPriority
    {
        public static MessagePriority GetPriority(string priority)
        {
            switch (priority)
            {
                case "0": return MessagePriority.NonUrgent;
                case "1": return MessagePriority.Normal;
                case "2": return MessagePriority.Urgent;
                case "NonUrgent": return MessagePriority.NonUrgent;
                case "Normal": return MessagePriority.Normal;
                case "Urgent": return MessagePriority.Urgent;
                //Low and High used by Send-MailMessage Cmdlet; including as aliases
                case "Low": return MessagePriority.NonUrgent;
                case "High": return MessagePriority.Urgent;
                default: throw new Exception($"Priority '{priority}' not found. Valid priorities include: NonUrgent, Normal, Urgent.");
            }
        }
    }

    [Cmdlet(VerbsCommunications.Send, "MailKitMessage")]    //I think the [CmdletBinding] piece is applicable to true PowerShell functions, not compiled cmdlets https://docs.microsoft.com/en-us/powershell/module/microsoft.powershell.core/about/about_functions_cmdletbindingattribute?view=powershell-7.1#long-description
    [OutputType(typeof(void))]
    public class Send_MailKitMessage : PSCmdlet
    {
        [Parameter(
            Mandatory = false)]
        public SwitchParameter UseSecureConnectionIfAvailable { get; set; } = SwitchParameter.Present;  //default to present if no value is passed

        [Parameter(
            Mandatory = false)]
        public PSCredential Credential { get; set; }

        [Parameter(
            Mandatory = true)]
        public string SMTPServer { get; set; }

        [Parameter(
            Mandatory = false)]
        public int Port { get; set; } = 25; //default to port 25 if no value is passed

        [Parameter(
            Mandatory = false)]
        public string Priority { get; set; }

        [Parameter(
            Mandatory = true)]
        public MailboxAddress From { get; set; }

        [Parameter(
            Mandatory = false)]
        public InternetAddressList ReplyToList { get; set; }

        [Parameter(
            Mandatory = true)]
        public InternetAddressList RecipientList { get; set; }

        [Parameter(
            Mandatory = false)]
        public InternetAddressList CCList { get; set; }

        [Parameter(
            Mandatory = false)]
        public InternetAddressList BCCList { get; set; }

        [Parameter(
            Mandatory = false)]
        public string Subject { get; set; }

        [Parameter(
            Mandatory = false)]
        public string TextBody { get; set; }

        [Parameter(
            Mandatory = false)]
        public string HTMLBody { get; set; }

        [Parameter(
            Mandatory = false)]
        public string[] AttachmentList { get; set; }

        // This method gets called once for each cmdlet in the pipeline when the pipeline starts executing
        protected override void BeginProcessing()
        {
            
        }

        // This method will be called for each input received from the pipeline to this cmdlet; if no input is received, this method is not called
        protected override void ProcessRecord()
        {

            MimeMessage Message = new MimeMessage();
            BodyBuilder Body = new BodyBuilder();
            MailKit.Net.Smtp.SmtpClient Client = new MailKit.Net.Smtp.SmtpClient();

            try
            {
                //priority
                if (!string.IsNullOrWhiteSpace(Priority))
                {
                    Message.Priority = EmailPriority.GetPriority(Priority);
                }

                //from
                Message.From.Add(From);

                //replyto
                if (ReplyToList?.Count > 0)
                {
                    Message.ReplyTo.AddRange(ReplyToList);
                }

                //to
                Message.To.AddRange(RecipientList);

                //cc
                if (CCList?.Count > 0)
                {
                    Message.Cc.AddRange(CCList);
                }

                //bcc
                if (BCCList?.Count > 0)
                {
                    Message.Bcc.AddRange(BCCList);
                }

                //subject
                if (!string.IsNullOrWhiteSpace(Subject))
                {
                    Message.Subject = Subject;
                }

                //text body
                if (!string.IsNullOrWhiteSpace(TextBody))
                {
                    Body.TextBody = TextBody;
                }

                //html body
                if (!string.IsNullOrWhiteSpace(HTMLBody))
                {
                    Body.HtmlBody = HttpUtility.HtmlDecode(HTMLBody);    //decode html in case it was encoded along the way
                }

                //attachment(s)
                if (AttachmentList?.Length > 0)
                {
                    foreach (string Attachment in AttachmentList)
                    {
                        Body.Attachments.Add(Attachment);
                    }
                }

                //add bodybuilder to body
                Message.Body = Body.ToMessageBody();

                //smtp send
                Client.Connect(SMTPServer, Port, (UseSecureConnectionIfAvailable.IsPresent ? MailKit.Security.SecureSocketOptions.Auto : MailKit.Security.SecureSocketOptions.None));
                if (Credential != null)
                {
                    Client.Authenticate(Credential.UserName, (System.Runtime.InteropServices.Marshal.PtrToStringAuto(System.Runtime.InteropServices.Marshal.SecureStringToBSTR(Credential.Password))));
                }
                Client.Send(Message);

            }
            catch (Exception e)
            {
                
                throw e;
            }
            finally
            {
                if (Client.IsConnected)
                {
                    Client.Disconnect(true);
                }
            }
        }

        // This method will be called once at the end of pipeline execution; if no input is received, this method is not called
        protected override void EndProcessing()
        {
            
        }
    }

    public class ModuleCleanup : IModuleAssemblyCleanup
    {
        public void OnRemove(PSModuleInfo psModuleInfo)
        {
            AppDomain.CurrentDomain.AssemblyResolve -= DependencyResolution.ResolveSystemBuffers;
        }
    }
}
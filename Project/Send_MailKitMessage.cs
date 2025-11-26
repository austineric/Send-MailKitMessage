using MimeKit;
using System;
using System.Collections;
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
            Mandatory = true)]
        public int Port { get; set; }

        [Parameter(
            Mandatory = true)]
        public MailboxAddress From { get; set; }

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

        [Parameter(
            Mandatory = false)]
        [ValidateSet(new string[]
        {
            "Normal",
            "High",
            "Low"
        })]
        public string Priority { get; set; } = "Normal";

        [Parameter(
            Mandatory = false)]
        public Hashtable InlineImages { get; set; }

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

                //from
                Message.From.Add(From);

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

                //inline attachment(s)
                if(InlineImages != null && InlineImages.Count > 0)
                {
                    foreach(object obj in InlineImages)
                    {
                        DictionaryEntry dictionaryEntry = (dictionaryEntry)obj;
                        string contentId = dictionaryEntry.Key.ToString();
                        string text = dictionaryEntry.Value.ToString();
                        if(File.Exists(text))
                        {
                            string[] array = MimeTypes.GetMimeType(text).Split(new char[]
                            {
                                '/'
                            });
                            MimePart attachment = new MimePart(new ContentType(array[0], array[1]))
                            {
                                Content = new MimeContent(File.OpenRead(text), ContentEncoding.Default),
                                ContentDisposition = new ContentDisposition("inline"),
                                ContentTransferEncoding = ContentEncoding.Base64,
                                ContentId = contentId,
                                FileName = Path.GetFileName(text)
                            };
                            Body.LinkedResources.Add(attachment);
                        }
                        else
                        {
                            base.WriteWarning("Inline image file not found: " + text);
                        }
                    }
                }

                //add bodybuilder to body
                Message.Body = Body.ToMessageBody();

                //set message priority
                string prio = Priority.ToLower();
                if(!(prio == "high"))
                {
                    if(prio == "low")
                    {
                        Message.Headers.Add("X-Priority", "5 (Lowest)");
                        Message.Headers.Add("X-MSMail-Priority", "Low");
                        Message.Headers.Add("Importance", "Low");
                    }
                }
                else
                {
                    Message.Headers.Add("X-Priority", "1 (Highest)");
                    Message.Headers.Add("X-MSMail-Priority", "High");
                    Message.Headers.Add("Importance", "High");
                }

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

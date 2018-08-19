using OutlookMail;
using System;
using System.Activities;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace Santan
{
    namespace OutlookMail
    {
        [Designer(typeof(SendOutlookMailActivityDesigner))]
        public class SendOutlookMail : CodeActivity
        {
            [DisplayName("BCC"), Category("Receiver"), Description("Email address of the Hidden Recipients")]
            public InArgument<String> BCC { get; set; }
            [DisplayName("CC"), Category("Receiver"), Description("Email address of the Hidden Recipients")]
            public InArgument<String> CC { get; set; }
            [DisplayName("To"), Category("Receiver"), RequiredArgument, Description("Email address of the Recipients")]
            public InArgument<String> To { get; set; }
            [DisplayName("Account"), Category("Input"), Description("The Account Used to send Email")]
            public InArgument<String> Account { get; set; }

            [DisplayName("Body"), Category("Input"), Description("The body of Email")]
            public InArgument<String> Body { get; set; }
            [DisplayName("Subject"), Category("Input"), Description("The Subject of Email")]
            public InArgument<String> Subject { get; set; }
            [DisplayName("isBodyHTML"), Category("Options"), Description("Is Body HTML or Text")]
            public Boolean isBodyHTML { get; set; }

            [DisplayName("Attachments"), Category("Attachment"), Description("Files to be attached in Email")]
            public InArgument<String[]> Attachment { get; set; }



            protected override void Execute(CodeActivityContext context)
            {
                String sTO = To.Get(context);
                String sBCC = BCC.Get(context);
                String sCC = CC.Get(context);
                String sAccount = Account.Get(context);
                String sBody = Body.Get(context);
                String sSubject = Subject.Get(context);
                Boolean bisBodyHTML = isBodyHTML;
                String[] arrAttachment = Attachment.Get(context);
                sendOutlookMail(sTO, sCC, sBCC, sBody, sSubject, sAccount, bisBodyHTML, arrAttachment);
            }


            public Boolean AddReceipents(Outlook.MailItem mailItem, String ToAddress, String CC, String BCC)
            {
                Outlook.Recipients receipents = null;
                Outlook.Recipient receipentsTo = null;
                Outlook.Recipient receipentsCC = null;
                Outlook.Recipient receipentsBCC = null;

                Boolean SuccessFlag = false;

                try
                {
                    //This is for Adding Recepients in Mail.
                    receipents = mailItem.Recipients;

                    //This is for adding To Address
                    if (!string.IsNullOrWhiteSpace(ToAddress))
                    {
                        string[] arrToAddress = ToAddress.Split(new char[] { ',', ';' });
                        foreach (string Addr in arrToAddress)
                        {
                            try
                            {
                                if (!string.IsNullOrWhiteSpace(Addr) && Addr.IndexOf('@') != -1)
                                {
                                    receipentsTo = receipents.Add(Addr.Trim());
                                    receipentsTo.Type = (int)Outlook.OlMailRecipientType.olTo;
                                }
                                else
                                {
                                    throw new Exception("\"To\" Address is not correct : " + Addr);
                                }
                            }
                            catch (Exception e)
                            {
                                throw new System.Exception("Error in \"To\" Address");
                            }
                        }
                    }
                    else
                    {
                        throw new Exception("Value for the required variable \"TO\" was not supplied");
                    }

                    //This is for adding CC Address
                    if (!string.IsNullOrWhiteSpace(CC))
                    {
                        string[] arrCC = CC.Split(new char[] { ',', ';' });
                        foreach (string Addr in arrCC)
                        {
                            try
                            {
                                if (!string.IsNullOrWhiteSpace(Addr) && Addr.IndexOf('@') != -1)
                                {
                                    receipentsCC = receipents.Add(Addr.Trim());
                                    receipentsCC.Type = (int)Outlook.OlMailRecipientType.olCC;
                                }
                                else
                                {
                                    throw new Exception("\"CC\" Address is not correct : " + Addr);
                                }
                            }
                            catch (Exception e)
                            {
                                throw new System.Exception("Error in \"CC\" Address");
                            }
                        }
                    }


                    //This is for adding BCC Address
                    if (!string.IsNullOrWhiteSpace(BCC))
                    {
                        string[] arrBCC = BCC.Split(new char[] { ',', ';' });
                        foreach (string Addr in arrBCC)
                        {
                            if (!string.IsNullOrWhiteSpace(Addr) && Addr.IndexOf('@') != -1)
                            {
                                receipentsBCC = receipents.Add(Addr.Trim());
                                receipentsBCC.Type = (int)Outlook.OlMailRecipientType.olBCC;
                            }
                            else
                            {
                                throw new Exception("\"CC\" Address is not correct " + ToAddress);
                            }
                        }
                    }

                    //Resolving all address
                    SuccessFlag = receipents.ResolveAll();

                }
                catch (Exception e)
                {
                    throw e;
                }
                finally
                {
                    receipents = null;
                    receipentsTo = null;
                    receipentsCC = null;
                    receipentsBCC = null;
                }
                return SuccessFlag;
            }

            public void sendOutlookMail(String ToAddress, String CC, String BCC, String Body, String Subject, String FromAddress, Boolean isHTML, String[] arrAttachment)
            {
                try
                {
                    Outlook.Application OutlookApp = new Outlook.Application();
                    Outlook.MailItem oNewMail = (Outlook.MailItem)OutlookApp.CreateItem(Outlook.OlItemType.olMailItem);

                    //Adding To, Cc, BCC
                    AddReceipents(oNewMail, ToAddress, CC, BCC);

                    //Mail Type 
                    if (isHTML)
                    {
                        oNewMail.HTMLBody = Body;
                    }
                    else
                    {
                        oNewMail.Body = Body;
                    }

                    if (arrAttachment != null)
                    {
                        Outlook.Attachment oAttach;
                        int iPosition = (int)oNewMail.Body.Length + 1;
                        int iAttachType = (int)Outlook.OlAttachmentType.olByValue;

                        foreach (String attach in arrAttachment)
                        {
                            if (!String.IsNullOrWhiteSpace(attach))
                            {
                                oAttach = oNewMail.Attachments.Add(attach, iAttachType, iPosition);
                            }
                        }
                    }

                    oNewMail.Subject = Subject;

                    if (String.IsNullOrWhiteSpace(FromAddress))
                    {

                    }
                    else
                    {
                        Outlook.Accounts AllAccount = OutlookApp.Session.Accounts;
                        Outlook.Account SendAccount = null;

                        foreach (Outlook.Account account in AllAccount)
                        {
                            if (account.SmtpAddress.Equals(FromAddress, StringComparison.CurrentCultureIgnoreCase))
                            {
                                SendAccount = account;
                                break;
                            }
                        }
                        if (SendAccount != null)
                        {
                            oNewMail.SendUsingAccount = SendAccount;
                        }
                        else
                        {
                            throw new Exception("Account does not exist in Outlook : " + FromAddress);
                        }
                    }
                    oNewMail.Send();
                    oNewMail = null;
                    OutlookApp = null;
                }
                catch (Exception ex)
                {
                    throw ex;
                }
            }
        }

        [Designer(typeof(SendOutlookMailWithImageActivityDesigner))]
        public class sendOutlookMailWithImage : CodeActivity
        {
            [DisplayName("BCC"), Category("Receiver"), Description("Email address of the Hidden Recipients")]
            public InArgument<String> BCC { get; set; }
            [DisplayName("CC"), Category("Receiver"), Description("Email address of the Hidden Recipients")]
            public InArgument<String> CC { get; set; }
            [DisplayName("To"), Category("Receiver"), RequiredArgument, Description("Email address of the Recipients")]
            public InArgument<String> To { get; set; }
            [DisplayName("Account"), Category("Input"), Description("The Account Used to send Email")]
            public InArgument<String> Account { get; set; }

            [DisplayName("Body"), Category("Input"), Description("The body of Email")]
            public InArgument<String> Body { get; set; }
            [DisplayName("Subject"), Category("Input"), Description("The Subject of Email")]
            public InArgument<String> Subject { get; set; }

            [DisplayName("Attachments"), Category("Attachment"), Description("Files to be attached in Email")]
            public InArgument<String[]> Attachment { get; set; }
            [DisplayName("Embed Image"), Category("Images"), Description("Dictionary with Key and Value, Key:Image Key, Value:Image Path")]
            public InArgument<Dictionary<String,String>> EmbedImage { get; set; }


            protected override void Execute(CodeActivityContext context)
            {
                String sTO = To.Get(context);
                String sBCC = BCC.Get(context);
                String sCC = CC.Get(context);
                String sAccount = Account.Get(context);
                String sBody = Body.Get(context);
                String sSubject = Subject.Get(context);
                String[] arrAttachment = Attachment.Get(context);
                Dictionary<String, String> dict = EmbedImage.Get(context);

                sendOutlookMailImage(sTO, sCC, sBCC, sBody, sSubject, sAccount, arrAttachment, dict);
            }

            public Boolean AddReceipents(Outlook.MailItem mailItem, String ToAddress, String CC, String BCC)
            {
                Outlook.Recipients receipents = null;
                Outlook.Recipient receipentsTo = null;
                Outlook.Recipient receipentsCC = null;
                Outlook.Recipient receipentsBCC = null;

                Boolean SuccessFlag = false;

                try
                {
                    //This is for Adding Recepients in Mail.
                    receipents = mailItem.Recipients;

                    //This is for adding To Address
                    if (!string.IsNullOrWhiteSpace(ToAddress))
                    {
                        string[] arrToAddress = ToAddress.Split(new char[] { ',', ';' });
                        foreach (string Addr in arrToAddress)
                        {
                            try
                            {
                                if (!string.IsNullOrWhiteSpace(Addr) && Addr.IndexOf('@') != -1)
                                {
                                    receipentsTo = receipents.Add(Addr.Trim());
                                    receipentsTo.Type = (int)Outlook.OlMailRecipientType.olTo;
                                }
                                else
                                {
                                    throw new Exception("\"To\" Address is not correct : " + Addr);
                                }
                            }
                            catch (Exception e)
                            {
                                throw new System.Exception("Error in \"To\" Address");
                            }
                        }
                    }
                    else
                    {
                        throw new Exception("Value for the required variable \"TO\" was not supplied");
                    }

                    //This is for adding CC Address
                    if (!string.IsNullOrWhiteSpace(CC))
                    {
                        string[] arrCC = CC.Split(new char[] { ',', ';' });
                        foreach (string Addr in arrCC)
                        {
                            try
                            {
                                if (!string.IsNullOrWhiteSpace(Addr) && Addr.IndexOf('@') != -1)
                                {
                                    receipentsCC = receipents.Add(Addr.Trim());
                                    receipentsCC.Type = (int)Outlook.OlMailRecipientType.olCC;
                                }
                                else
                                {
                                    throw new Exception("\"CC\" Address is not correct : " + Addr);
                                }
                            }
                            catch (Exception e)
                            {
                                throw new System.Exception("Error in \"CC\" Address");
                            }
                        }
                    }


                    //This is for adding BCC Address
                    if (!string.IsNullOrWhiteSpace(BCC))
                    {
                        string[] arrBCC = BCC.Split(new char[] { ',', ';' });
                        foreach (string Addr in arrBCC)
                        {
                            if (!string.IsNullOrWhiteSpace(Addr) && Addr.IndexOf('@') != -1)
                            {
                                receipentsBCC = receipents.Add(Addr.Trim());
                                receipentsBCC.Type = (int)Outlook.OlMailRecipientType.olBCC;
                            }
                            else
                            {
                                throw new Exception("\"CC\" Address is not correct " + ToAddress);
                            }
                        }
                    }

                    //Resolving all address
                    SuccessFlag = receipents.ResolveAll();

                }
                catch (Exception e)
                {
                    throw e;
                }
                finally
                {
                    receipents = null;
                    receipentsTo = null;
                    receipentsCC = null;
                    receipentsBCC = null;
                }
                return SuccessFlag;
            }

            public void sendOutlookMailImage(String ToAddress, String CC, String BCC, String Body, String Subject, String FromAddress, String[] arrAttachment, Dictionary<String, String> dictimage)
            {
                try
                {
                    Outlook.Application OutlookApp = new Outlook.Application();
                    Outlook.MailItem oNewMail = (Outlook.MailItem)OutlookApp.CreateItem(Outlook.OlItemType.olMailItem);
                    oNewMail.BodyFormat = Outlook.OlBodyFormat.olFormatHTML;

                    AddReceipents(oNewMail, ToAddress, CC, BCC);

                    oNewMail.Subject = Subject;
                    oNewMail.HTMLBody = Body;

                    if (dictimage != null)
                    {
                        foreach (KeyValuePair<String, String> FileName in dictimage)
                        {
                            Outlook.Attachment Att1 = oNewMail.Attachments.Add(FileName.Value, Outlook.OlAttachmentType.olEmbeddeditem, null, "");
                            Att1.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001E", FileName.Key);
                        }
                    }

                    if (arrAttachment != null)
                    {
                        int iPosition = (int)oNewMail.HTMLBody.Length + 1;
                        int iAttachType = (int)Outlook.OlAttachmentType.olByValue;
                        Outlook.Attachment oAttach;

                        foreach (String attach in arrAttachment)
                        {
                            if (!String.IsNullOrWhiteSpace(attach))
                            {
                                oAttach = oNewMail.Attachments.Add(attach, iAttachType, iPosition);
                            }
                        }
                    }

                    if (String.IsNullOrWhiteSpace(FromAddress))
                    {

                    }
                    else
                    {
                        Outlook.Accounts AllAccount = OutlookApp.Session.Accounts;
                        Outlook.Account SendAccount = null;

                        foreach (Outlook.Account account in AllAccount)
                        {
                            if (account.SmtpAddress.Equals(FromAddress, StringComparison.CurrentCultureIgnoreCase))
                            {
                                SendAccount = account;
                                break;
                            }
                        }
                        if (SendAccount != null)
                        {
                            oNewMail.SendUsingAccount = SendAccount;
                        }
                        else
                        {
                            throw new Exception("Account does not exist in Outlook : " + FromAddress);
                        }
                    }
                    oNewMail.Send();
                    oNewMail = null;
                    OutlookApp = null;
                }
                catch (Exception ex)
                {
                    throw ex;
                }
            }
        }

    }
}
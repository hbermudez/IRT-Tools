using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using Office = Microsoft.Office.Core;
using Outlook = Microsoft.Office.Interop.Outlook;
using System.Windows.Forms;
using System.Text.RegularExpressions;
using System.Drawing;
using System.Diagnostics;
using System.Net;
using System.Collections.Specialized;
using System.Collections;
using System.Net.Sockets;
using System.Xml;
using System.Threading;
using System.ComponentModel;

// TODO:  Follow these steps to enable the Ribbon (XML) item:

// 1: Copy the following code block into the ThisAddin, ThisWorkbook, or ThisDocument class.

//  protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject()
//  {
//      return new Ribbon1();
//  }

// 2. Create callback methods in the "Ribbon Callbacks" region of this class to handle user
//    actions, such as clicking a button. Note: if you have exported this Ribbon from the Ribbon designer,
//    move your code from the event handlers to the callback methods and modify the code to work with the
//    Ribbon extensibility (RibbonX) programming model.

// 3. Assign attributes to the control tags in the Ribbon XML file to identify the appropriate callback methods in your code.  

// For more information, see the Ribbon XML documentation in the Visual Studio Tools for Office Help.


namespace OutlookParserAddIn
{
    [ComVisible(true)]
    public class IRTOutlookRibbon : Office.IRibbonExtensibility
    {
        private Office.IRibbonUI ribbon;
        private string selectedPath = null;
        private string summaryPath = null;
        private string emailAlertsReportPath = null;
        private string attachmentPath = null;
        private StreamWriter sw;
        private string delimiter = "------------------------------------";
        private Dictionary<string, string> extractedIPsURLs;       
        private string emailId;        
        private BackgroundWorker bw = new BackgroundWorker();
        private OsintAnalysis osint;
        private Outlook.Selection mailselected;
		
        public IRTOutlookRibbon()
        {
            extractedIPsURLs = new Dictionary<string, string>();            
            bw.DoWork += new DoWorkEventHandler(bw_DoWork);            
            bw.RunWorkerCompleted += new RunWorkerCompletedEventHandler(bw_RunWorkerCompleted);
        }

        public Image GetIcon(Office.IRibbonControl control)
        {
            return OutlookParserAddIn.Properties.Resources.intruder;
        }

        private List<string> LinkExtractor(string html)
        { 
            List<string> list = new List<string>();
            Regex regex = new Regex("(?:href|src)=[\"|']?(.*?)[\"|'|>]+", RegexOptions.Singleline | RegexOptions.CultureInvariant);

            if (regex.IsMatch(html))
            {
                foreach (Match match in regex.Matches(html))
                {
                    list.Add(match.Groups[1].Value);                    
                }
            }

            return list;
        }

        private List<string> IPsExtractor(string header)
        {
            List<string> list = new List<string>();
            Regex regex = new Regex(@"\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3}");

            if (regex.IsMatch(header))
            {
                foreach (Match match in regex.Matches(header))
                {
                    list.Add(match.ToString());

                    if (!extractedIPsURLs.ContainsKey(match.ToString()))
                    {
                        extractedIPsURLs.Add(match.ToString(), "IP");
                    }
                }
            }

            return list;
        }

        private String HexToASCII(String dirtyLink)
        {
            try
            {

                string ascii = string.Empty;

                for (int i = 0; i < dirtyLink.Length; i++)
                {
                    String hs = string.Empty;
                    hs = dirtyLink.Substring(i, 1);

                    if (hs.Equals("-"))
                    {
                        
                        var temp = dirtyLink.Substring(i+1, 2);
                        uint decval = System.Convert.ToUInt32(temp, 16);
                        char character = System.Convert.ToChar(decval);
                            
                        ascii += character;
                        i += 2;
                    }
                    else
                    {
                        ascii += hs;
                    }
                    
                }                
                return ascii;
            }
            catch(Exception ex)
            {
                MessageBox.Show("Error " + ex.Message );
            }
            return string.Empty;

        }
        
        private List<string> LinkAnalysis(List<string> dirtyLinks)
        {
            String proofpointPrepend = "https://urldefense.proofpoint.com/v2/url?u=";
            String proofpointTail = "&amp;d";
            List<string> cleanLinkList = new List<string>();

            foreach (string link in dirtyLinks)
            {
                if (link.Contains(proofpointPrepend))
                {
                    string temp = link.TrimStart(proofpointPrepend.ToCharArray());

                    if (temp.Contains(proofpointTail))
                    {
                        temp = temp.Substring(0, temp.Length - (temp.Length - temp.IndexOf(proofpointTail)));
                    }

                    temp = temp.Replace("-3A__", "hxxp://");                    
                    temp = temp.Replace("_", "/");
                    var cleanLink = HexToASCII(temp);
                    cleanLinkList.Add(cleanLink);
                }
            }

            return cleanLinkList;
        }

        private void WriteToFile(string message, TextWriter w)
        {
            w.WriteLine(message);            
        }

        public void WriteToFile(String filePath, String msg)
        {
            using (sw = File.AppendText(filePath))
            {
                WriteToFile(msg, sw); 
            }  
        }

        private void ProcessMsg(Outlook.MailItem item)
        {
            bool keepRunning = false;        

            //Attachments Analysis
            if (item.Attachments.Count > 0)
            {
                foreach (Outlook.Attachment attachment in item.Attachments)
                {
                    string tempPath = attachmentPath + "\\" + attachment.FileName + ".quarantine";
                    attachment.SaveAsFile(tempPath);
                }
            }

            //Headers
            Outlook.PropertyAccessor oPA = item.PropertyAccessor as Outlook.PropertyAccessor;
            const string PR_MAIL_HEADER_TAG = @"http://schemas.microsoft.com/mapi/proptag/0x007D001E";
            string itemHeader = null;
            try
            {
                itemHeader = (string)oPA.GetProperty(PR_MAIL_HEADER_TAG);
                //TODO: get only the last Received from fields

                using (sw = File.AppendText(summaryPath))
                {
                    WriteToFile(itemHeader, sw);
                    WriteToFile(delimiter, sw);
                }                
                keepRunning = true;
            }
            catch (IOException ex)
            {
                MessageBox.Show("We are sorry but there was an error during execution. See Details below \n" + ex.Message);
                keepRunning = false;

            }
            catch (Exception ex)
            {
                MessageBox.Show("We are sorry but there was an error during execution. See Details below \n" + ex.Message);
                keepRunning = false;

            }
            
            //Extract IPs from HEADER
            var extractedIPs = IPsExtractor(itemHeader);
            if (extractedIPs.Count > 0)
            {
                using (sw = File.AppendText(summaryPath))
                {
                    WriteToFile("IP founds on theMessage", sw);
                    WriteToFile(delimiter, sw);
                }
                
            }
            foreach (string ip in extractedIPs)
            {
                using (sw = File.AppendText(summaryPath))
                {
                    WriteToFile(ip, sw);
                }
            }

            if (keepRunning)
            {
                using (sw = File.AppendText(summaryPath))
                {
                    WriteToFile(delimiter, sw);
                    WriteToFile("Sender Name: " + item.SenderName.ToString(), sw);
                    WriteToFile(delimiter, sw);
                    WriteToFile("Subject: " + item.Subject.ToString(), sw);
                    WriteToFile(delimiter, sw);
                    WriteToFile("Sender: " + item.SenderEmailAddress.ToString(), sw);
                    WriteToFile(delimiter, sw);
                    WriteToFile("Links found: ", sw);
                    WriteToFile(delimiter, sw);
                }         
                
                //Link Extraction
                List<string> dirtylinks = LinkExtractor(item.HTMLBody);

                //Analysis of links
                var cleanLinks = LinkAnalysis(dirtylinks);

                foreach (string cleanLink in cleanLinks)
                {
                    using (sw = File.AppendText(summaryPath))
                    {
                        WriteToFile(cleanLink, sw);
                        if (!extractedIPsURLs.ContainsKey(cleanLink))
                        {
                            extractedIPsURLs.Add(cleanLink, "URL");
                        }
                        
                    }                     
                }
                using (sw = File.AppendText(summaryPath))
                {
                    WriteToFile(delimiter + "END OF FILE" + delimiter, sw);
                } 
            }
            
        }
        
        public void OnIRButtonClick(Office.IRibbonControl control)
        {
            /*Ask user where to save the analysis and attachments*/            
            FolderBrowserDialog folderPath = new FolderBrowserDialog();
            folderPath.Description = "Select the Folder to store the Analysis";
            var result = folderPath.ShowDialog();            

            //Process each selected mail
            if (control.Context is Outlook.Selection)
            {
                if (result == DialogResult.OK)
                {
                    mailselected = control.Context as Outlook.Selection;
                    selectedPath = folderPath.SelectedPath;
                    emailAlertsReportPath = selectedPath + "\\" + "Full_Report.txt";
                    bw.RunWorkerAsync();
                }
            }
            
        }

        private void bw_DoWork(object sender, DoWorkEventArgs e)
        {
            DialogResult firstAction = MessageBox.Show("Would you like to run Links/IPs against Virus Total?" +
                                              "Note that this action could take a while. You can keep working while we take care of this for you.", "OSINT Analysis",
                                              MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            
            try
            {                

                foreach (Outlook.MailItem item in mailselected.OfType<Outlook.MailItem>())
                {                    

                    //Create subfolders per item
                    string tempPath = null;
                    Boolean isAlreadyProcessed = false;

                    if (!Directory.Exists(selectedPath + "\\" + item.EntryID))
                    {
                        emailId = item.EntryID;
                        tempPath = selectedPath + "\\" + item.EntryID;
                        attachmentPath = tempPath + "\\" + "Attachments";
                        Directory.CreateDirectory(tempPath);
                        Directory.CreateDirectory(attachmentPath);
                        summaryPath = tempPath + "\\" + "summary.txt";
                    }
                    else
                    {
                        emailId = item.EntryID;
                        tempPath = selectedPath + "\\" + item.EntryID;
                        attachmentPath = tempPath + "\\" + "Attachments";
                        summaryPath = tempPath + "\\" + "summary.txt";
                    }

                    sw = new StreamWriter(summaryPath);
                    sw.Close();

                    using (sw = File.AppendText(summaryPath))
                    {
                        WriteToFile("Full Report", sw);
                        WriteToFile(delimiter, sw);
                    }

                    //Attachments Analysis
                    if (item.Attachments.Count > 0)
                    {
                        foreach (Outlook.Attachment attachment in item.Attachments)
                        {
                            //check if a .msg was sent as an attachment, in that case process that msg
                            string extension = Path.GetExtension(attachment.FileName);
                            if (extension.Equals(".msg"))
                            {
                                attachment.SaveAsFile(attachmentPath + "\\" + attachment.FileName);
                                Microsoft.Office.Interop.Outlook.Application msg = new Microsoft.Office.Interop.Outlook.Application();
                                var newItem = msg.Session.OpenSharedItem(attachmentPath + "\\" + attachment.FileName) as Outlook.MailItem;
                                ProcessMsg(newItem);
                                isAlreadyProcessed = true;
                            } 
                            else
                            { 
                                //save the attachment - the email was sent to you directly
                            }
                        }

                        if (!isAlreadyProcessed)
                        {
                            ProcessMsg(item);
                        }

                    }
                    else
                    {
                        ProcessMsg(item);
                    }

                    using (sw = File.AppendText(emailAlertsReportPath))
                    {
                        WriteToFile("Email Alerts: " + emailId + "\n", sw);
                        WriteToFile("-------------------------------\n", sw);
                    }

                    if (firstAction == DialogResult.Yes)
                    {
                        osint = new OsintAnalysis(extractedIPsURLs, selectedPath, emailAlertsReportPath, this);
                        //e.Result = osint.OsintRun();
                        e.Result = true;
                        osint.OsintRun();
                        extractedIPsURLs.Clear();
                    }

                } //end foreach
            }
            catch (Exception ex)
            {

                MessageBox.Show("mailItem Foreach loop Error " + ex.ToString());
            }
        }
                
        
        private void bw_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            // First, handle the case where an exception was thrown. 
            if (e.Error != null)
            {
                MessageBox.Show(e.Error.Message);
            }
            else if (e.Cancelled)
            {
                // Next, handle the case where the user canceled  
                // the operation. 
                // Note that due to a race condition in  
                // the DoWork event handler, the Cancelled 
                // flag may not have been set, even though 
                // CancelAsync was called.
                MessageBox.Show("Canceled");
            }
            else
            {
                DialogResult finalAction = MessageBox.Show("Analysis has finished. Do you wanna open the folder's location?", "Post Analysis Action Dialog",
                                    MessageBoxButtons.YesNo, MessageBoxIcon.Question);

                if (finalAction == DialogResult.Yes)
                {
                    Process.Start(selectedPath);
                }             
            }                                   
        }

        #region IRibbonExtensibility Members

        public string GetCustomUI(string ribbonID)
        {
            return GetResourceText("OutlookParserAddIn.Ribbon1.xml");
        }

        #endregion

        #region Ribbon Callbacks
        //Create callback methods here. For more information about adding callback methods, select the Ribbon XML item in Solution Explorer and then press F1

        public void Ribbon_Load(Office.IRibbonUI ribbonUI)
        {
            this.ribbon = ribbonUI;
        }

        #endregion

        #region Helpers

        private static string GetResourceText(string resourceName)
        {
            Assembly asm = Assembly.GetExecutingAssembly();
            string[] resourceNames = asm.GetManifestResourceNames();
            for (int i = 0; i < resourceNames.Length; ++i)
            {
                if (string.Compare(resourceName, resourceNames[i], StringComparison.OrdinalIgnoreCase) == 0)
                {
                    using (StreamReader resourceReader = new StreamReader(asm.GetManifestResourceStream(resourceNames[i])))
                    {
                        if (resourceReader != null)
                        {
                            return resourceReader.ReadToEnd();
                        }
                    }
                }
            }
            return null;
        }

        #endregion
    }
}

using System;
using System.Collections.Generic;
using System.Text;
using System.IO;
using System.Net.Mail;

namespace VanCashFileCopy
{
    class Tester
    {
        public void Run()
        {
            //-- Setup directories and files
            string sourceDirectory = @"O:\Work\Mike\WeeklyCash\";
            string logFile = @sourceDirectory + "vancash.log";
            string targetDirectory = @"L:\ARCSH\";
            string[] dirs = Directory.GetFiles(sourceDirectory, "*.xls");
            int numOfFiles = dirs.Length;

            if (numOfFiles > 0)
            {

                //-- Create/Open log file
                StreamWriter sw = new StreamWriter(logFile);

                //-- Setup the dated directory name
                int mth = DateTime.Now.Month;
                string mnth = "";

                if (mth < 10) { mnth = "0" + mth.ToString(); }
                else { mnth = mth.ToString(); }

                string dateDirName = DateTime.Now.Year.ToString() + mnth + DateTime.Now.Day.ToString() + "v\\";
                string archiveDirPath = @sourceDirectory + "Archive\\" + dateDirName;
                string targetDirPath = @targetDirectory + dateDirName;

                sw.WriteLine("Source Directory: {0}", sourceDirectory);
                sw.WriteLine("Archive Directory: {0}", archiveDirPath);
                sw.WriteLine("Vista Directory: {0}", targetDirPath);

                //-- Check if the directory under Archive exists
                if (Directory.Exists(archiveDirPath))
                {
                    Console.WriteLine("Directory already exists: " + archiveDirPath);
                    sw.WriteLine("");
                    sw.WriteLine("Directory already exists: " + archiveDirPath);
                    sw.Close();
                    return;
                }

                //-- Check if the directory under arcsh exists
                if (Directory.Exists(targetDirPath))
                {
                    Console.WriteLine("Directory already exists: " + targetDirPath);
                    sw.WriteLine("");
                    sw.WriteLine("Directory already exists: " + targetDirPath);
                    sw.Close();
                    return;
                }

                //-- Create needed directories. If problem, write error to log file and abort
                try
                {
                    DirectoryInfo arcDi = Directory.CreateDirectory(archiveDirPath);
                    DirectoryInfo tgtDi = Directory.CreateDirectory(targetDirPath);
                }
                catch
                {
                    Console.WriteLine("Could not create a target directory");
                    sw.WriteLine("");
                    sw.WriteLine("Could not create a target directory");
                    sw.Close();
                    return;
                }

                sw.WriteLine("");

                //-- Cycle thru the files copying to 2 locations then delete
                foreach (string fullFileName in dirs)
                {
                    string justFileName = Path.GetFileName(fullFileName);
                    int justFileNameLength = justFileName.Length;

                    //-- Change filename if it does not contain a CSH prefix
                    string cshFileName = "";

                    if (justFileNameLength == 9)
                    {
                        cshFileName = "CSH" + Path.GetFileName(fullFileName);
                    }
                    else
                    {
                        cshFileName = Path.GetFileName(fullFileName);
                    }

                    //-- Set the copy from/copy to parameters
                    string srcFile = sourceDirectory + Path.GetFileName(fullFileName);
                    string targ1File = archiveDirPath + cshFileName;
                    string targ2File = targetDirPath + cshFileName;

                    //-- Call the copy method
                    CopyFiles(sw, srcFile, targ1File, targ2File);
                }

                string mBody = "Weekly Vantage Files have been copied to \\\\LTC024N\\DATA\\ARCSH and are ready to be loaded to Vista.";
                string mSubj = "Weekly Vantage Cash Files Notification";
                string mFrom = "ITProcessMgr@SavaSC.com";
                string mTo = "Email1@Nowhere.com, Email2@Nowhere.com, Email3@Nowhere.com";
                
                string mServer = "10.10.10.10";     // mail server

                SendEmail(mBody, mSubj, mFrom, mTo, mServer);

                sw.WriteLine("");
                sw.WriteLine("Email sent to: {0}", mTo);
                sw.WriteLine("");
                sw.WriteLine("Process Complete");
                sw.Close();
            }
        }

        void CopyFiles(StreamWriter sw2, string src, string tgt1, string tgt2)
        {
            //-- Copy the source file to the target locations
            File.Copy(src, tgt1);
            File.Copy(src, tgt2);
            sw2.WriteLine("Copied " + src + " to " + tgt1);
            sw2.WriteLine("Copied " + src + " to " + tgt2);

            //-- Delete the source file after making sure it was copied successfully
            if (File.Exists(tgt2))
            {
                File.Delete(src);
                sw2.WriteLine("Deleted " + src);
            }
        }

        static void SendEmail(string msgBody, string msgSubject, string msgFrom,
                   string msgTo, string msgServer)
        {
            using (MailMessage mailMessage = new MailMessage(msgFrom, msgTo))
            {
                //-- Set the mailMessage object properties
                mailMessage.Body = msgBody;
                mailMessage.Subject = msgSubject;

                MailAddress bcc = new MailAddress("bccemail@Nowhere.com");
                mailMessage.Bcc.Add(bcc);

                //-- Initialize smtpClient and send message
                try
                {
                    SmtpClient smtpClient = new SmtpClient(msgServer);
                    smtpClient.Send(mailMessage);
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex.Message);
                }
            }
        }


        static void Main(string[] args)
        {
            //-- Create new instance of the tester class and execute the Run method
            Tester t = new Tester();
            t.Run();
        }
    }
}

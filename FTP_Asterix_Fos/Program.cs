using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Mail;
using System.Text;
using System.Threading.Tasks;
using WinSCP;

namespace FTP_Asterix_Fos
{
    class Program
    {
        static void Main(string[] args)
        {
            //HentFiler(@"\\10.10.11.11\sw\Kamstrup\Rafr120Ind\Ind\", "/MeterChange/SwChangeAndRead/", "*.csv", true);
            //HentFiler(@"\\10.10.11.11\MeterData\Asterix\RRSTemp\", "/asterix/", "*.a56", false);
            //HentFiler(@"\\10.10.11.11\MeterData\FOS\RRSTemp\", "/fos/", "*.txt", false);
            SendFiler(@"\\10.10.11.11\MeterData\OnDemand\Alarmer\", "/OnDemand/Alarmer/", "*.xlsx", false);
            SendFiler(@"\\10.10.11.11\MeterData\OnDemand\Alarmer\", "/OnDemand/Alarmer/", "*.csv", false);
            HentFiler(@"\\10.10.11.11\MeterData\BeofElInterval\", "/ElnetElIntern/", "*.csv", false, true);
            //SendMail("tbm@beof.dk", "tbm@beof.dk", "", "", "Test", "Tester", "c:\\temp\\test.txt");
        }

        public static void HentFiler(string localpath, string remotepath, string filetype, bool MakeBackupFolder, bool removeFile)
        {
            int i = 0;
            //string RemotePath;
            string LocalPath = localpath;
            string LocalPathDate;
            
            try
            {
                //FTP options
                SessionOptions sessionOptions = new SessionOptions
                {
                    Protocol = Protocol.Ftp,
                    HostName = "okfile.hstz2.amr.kamstrup.com",
                    PortNumber = 21,
                    UserName = "okftp",
                    Password = "UH3LFmesvJ8WQBbJ",
                    FtpSecure = FtpSecure.Explicit,
                    TlsHostCertificateFingerprint = "d7:99:c2:9e:cb:4c:96:a5:e2:05:21:05:1c:92:c8:7b:fd:d8:38:69", //Skal skiftes hvis cert fejler. Findes ved at oprette ftp til adressen på port 21 i winscp
                };

                using (Session session = new Session())
                {
                    // Connect
                    session.Open(sessionOptions);

                    //DateTime dt = DateTime.Now;
                    
                        LocalPathDate = LocalPath + DateTime.Now.ToString("yyyy-MM-dd") + "\\";                    

                    string remotePath = remotepath;
                    string fileType = filetype;


                    session.EnumerateRemoteFiles(remotePath, fileType, EnumerationOptions.None);

                    // Hent filliste i dir
                    RemoteDirectoryInfo directoryInfo = session.ListDirectory(remotePath);

                    foreach (RemoteFileInfo fileInfo in directoryInfo.Files)
                    {
                        // Er der nogen filer?
                        if (fileInfo == null)
                        {
                            throw new Exception("Ingen filer fundet");
                        }

                        string d = fileInfo.LastWriteTime.ToString("yyyy-MM-dd");

                        // Hent valgte fil
                        if ((fileInfo.Name != "..") && (d == DateTime.Now.ToString("yyyy-MM-dd")))
                        {

                            session.GetFiles((remotePath + fileInfo.Name), localpath,removeFile).Check();
                            Console.WriteLine("Download af {0} gennemført", fileInfo.Name);
                            i++;
                        }
                    }
                }

                Console.WriteLine("Der blev hentet {0} filer", i.ToString());
                //Console.Read();
                //return 0;
                //string LocalFilePath = localpath;
                string[] files = Directory.GetFiles(LocalPath);
                string fileName;
                string destFile;

                // Kopier filer og overskriv hvis de eksistere i forvejen
                int x = files.Count();
                if ((x > 0) & (MakeBackupFolder))
                {
                    if (!Directory.Exists(LocalPathDate))
                    {
                        Directory.CreateDirectory(LocalPathDate);
                    }
                    foreach (string s in files)
                    {
                        fileName = Path.GetFileName(s);
                        destFile = Path.Combine(LocalPathDate, fileName);

                        File.Copy(s, destFile, true);
                    }
                }


                //Console.Read();
            }
            catch (Exception e)
            {
                Console.WriteLine("Error: {0}", e);
                Console.ReadLine();
                //return 1;
            }
        }

        public static void SendFiler(string localpath, string remotepath, string filetype, bool MakeBackupFolder)
        {
            int i = 0;
            //string RemotePath;
            string LocalPath = localpath;
            string LocalPathDate;
            string myPath = "\\\\10.10.11.11\\MeterData\\OnDemand\\Alarmer\\";


            try
            {

                DirectoryInfo di = new DirectoryInfo(myPath);

                foreach (FileInfo file in di.GetFiles())
                {
                    file.Delete();
                }

                //FTP options
                SessionOptions sessionOptions = new SessionOptions
                {
                    Protocol = Protocol.Ftp,
                    HostName = "okfile.hstz2.amr.kamstrup.com",
                    PortNumber = 21,
                    UserName = "okftp",
                    Password = "UH3LFmesvJ8WQBbJ",
                    FtpSecure = FtpSecure.Explicit,
                    TlsHostCertificateFingerprint = "d7:99:c2:9e:cb:4c:96:a5:e2:05:21:05:1c:92:c8:7b:fd:d8:38:69", //Skal skiftes hvis cert fejler. Findes ved at oprette ftp til adressen på port 21 i winscp
                };

                using (Session session = new Session())
                {
                    // Connect
                    session.Open(sessionOptions);

                    //DateTime dt = DateTime.Now;

                    LocalPathDate = LocalPath + DateTime.Now.ToString("yyyy-MM-dd") + "\\";

                    string remotePath = remotepath;
                    string fileType = filetype;


                    session.EnumerateRemoteFiles(remotePath, fileType, EnumerationOptions.None);

                    // Hent filliste i dir
                    RemoteDirectoryInfo directoryInfo = session.ListDirectory(remotePath);

                    foreach (RemoteFileInfo fileInfo in directoryInfo.Files)
                    {
                        // Er der nogen filer?
                        if (fileInfo == null)
                        {
                            throw new Exception("Ingen filer fundet");
                        }

                        //string d = fileInfo.LastWriteTime.ToString("yyyy-MM-dd");
                        
                        // Hent valgte fil
                        if (fileInfo.Name != "..")
                        {
                            session.GetFiles((remotePath + fileInfo.Name), localpath, true).Check();
                            SendMail("alarm@beof.dk", "alarm@beof.dk", "", "", "Alarmer", "Vedhæftet er alarmer fra Kamstrup", myPath + fileInfo.Name);
                            
                            Console.WriteLine("Der er sendt {0} filer", fileInfo.Name);
                            i++;                            
                        }
                    }
                }

                Console.WriteLine("Der blev sendt {0} filer", i.ToString());
                //Console.Read();
                //return 0;
                //string LocalFilePath = localpath;
                
                //Console.Read();
            }
            catch (Exception e)
            {
                Console.WriteLine("Error: {0}", e);
                //Console.ReadLine();
                //return 1;
            }
        }

        public static void SendMail(string from, string to, string bcc, string cc, string subject, string body, string attName)
        {
            MailMessage mMailMessage = new MailMessage();

            //Attachment Att = new Attachment(attName, "application/zip");
            mMailMessage.From = new MailAddress(from);
            mMailMessage.To.Add(new MailAddress(to));
            if ((bcc != null) && (bcc != string.Empty))
            {
                mMailMessage.Bcc.Add(new MailAddress(bcc));
            }
            if ((cc != null) && (cc != string.Empty))
            {
                mMailMessage.CC.Add(new MailAddress(cc));
            }
            mMailMessage.Subject = subject;
            mMailMessage.Attachments.Add(new Attachment(attName));
            //mMailMessage.Attachments.Add(Att);
            mMailMessage.Body = body;
            mMailMessage.IsBodyHtml = true;
            mMailMessage.Priority = MailPriority.Normal;
            SmtpClient mSmtpClient = new SmtpClient("10.10.10.94");
            mSmtpClient.Send(mMailMessage);


        }

    }
}

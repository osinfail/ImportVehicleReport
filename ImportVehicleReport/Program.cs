using System;
using System.Configuration;
using System.IO;
using System.Linq;
using ImportVehicleReport.Report;
using Microsoft.Office.Interop.Outlook;

namespace ImportVehicleReport
{
    class Program
    {
        static void Main(string[] args)
        {
            Application app = new Application();

            Report.Report report = new Report.Report();
            PlanetVo pvo = new PlanetVo();
            Tec3H tec3h = new Tec3H();

            MailItem item;
            string[] tempData;
            string tempPath;

            string[] emails = Directory.GetFiles(ConfigurationManager.AppSettings["EmailPath"]);

            //PlanetVO XML Zip Report
            tempPath = Array.Find(emails,
                s =>
                    s.Contains(
                        "Import Vehicle PlanetVo FTP files downloaded successfully"));

            item = (MailItem)app.CreateItemFromTemplate(tempPath, Type.Missing);
            tempData = GetEmailData(item.Body);
            pvo.FtpSuccess = Int32.Parse(tempData[1]);
            pvo.FtpFailure = Int32.Parse(tempData[2]);

            //Tec3H XML Zip Report
            tempPath = Array.Find(emails,
                s =>
                    s.Contains(
                        "Import Vehicle Tec3H FTP files downloaded successfully"));
            item = (MailItem)app.CreateItemFromTemplate(tempPath, Type.Missing);
            tempData = GetEmailData(item.Body);
            tec3h.FtpSuccess = Int32.Parse(tempData[1]);
            tec3h.FtpFailure = Int32.Parse(tempData[2]);

            //PlanetVO XML Status
            tempPath = Array.Find(emails,
                s =>
                    s.Contains(
                        "Import Vehicle PlanetVo XML ImportVehicle"));

            item = (MailItem)app.CreateItemFromTemplate(tempPath, Type.Missing);
            tempData = GetEmailData(item.Body);
            pvo.ZipFiles = Int32.Parse(tempData[0]);
            pvo.ImportVehicleRecords = Int32.Parse(tempData[1]);
            pvo.NotWellFormedXmlCount = tempData[2];
            pvo.NotFoundPos = tempData[3];

            //Tec3H XML Status
            tempPath = Array.Find(emails,
                s =>
                    s.Contains(
                        "Import Vehicle Tec3H XML ImportVehicle"));

            item = (MailItem)app.CreateItemFromTemplate(tempPath, Type.Missing);
            tempData = GetEmailData(item.Body);
            tec3h.ZipFiles = Int32.Parse(tempData[0]);
            tec3h.ImportVehicleRecords = Int32.Parse(tempData[1]);
            tec3h.NotWellFormedXmlCount = tempData[2];
            tec3h.NotFoundPos = tempData[3];

            //Import Vehicle Stock Evolution
            tempPath = Array.Find(emails,
                s =>
                    s.Contains(
                        "Import VehicleImport Vehicle Stock Evolution"));
            item = (MailItem)app.CreateItemFromTemplate(tempPath, Type.Missing);
            tempData = GetEmailData(item.Body);
            pvo.StockCount = Int32.Parse(tempData[2]);
            pvo.NewStockCount = Int32.Parse(tempData[4]);
            pvo.DeletedStockCount = Int32.Parse(tempData[6]);
            tec3h.StockCount = Int32.Parse(tempData[3]);
            tec3h.NewStockCount = Int32.Parse(tempData[5]);
            tec3h.DeletedStockCount = Int32.Parse(tempData[7]);

            //Total photo count PlanetVO
            Photo pvoPhoto = new Photo();
            tempPath = Array.Find(emails,
                s =>
                    s.Contains(
                        "PlanetVO Retrieval and processing"));
            item = (MailItem)app.CreateItemFromTemplate(tempPath, Type.Missing);
            pvoPhoto.TotalCount = GetPhotoCount(item.Body);

            //Total photo count Tec3H
            Photo tec3HPhoto = new Photo();
            tempPath = Array.Find(emails,
                s =>
                    s.Contains(
                        "Tec3H Retrieval and processing"));
            item = (MailItem)app.CreateItemFromTemplate(tempPath, Type.Missing);
            tec3HPhoto.TotalCount = GetPhotoCount(item.Body);

            //Total photo summary PlanetVO
            tempPath = Array.Find(emails,
                s =>
                    s.Contains(
                        "PlanetVO Summary of import vehicle photo"));
            item = (MailItem)app.CreateItemFromTemplate(tempPath, Type.Missing);
            tempData = GetEmailData(item.Body);
            pvoPhoto.VehicleToTransfer = Int32.Parse(tempData[1]);
            pvoPhoto.PhotoCount = Int32.Parse(tempData[2]);
            pvoPhoto.NewPhotoCount = Int32.Parse(tempData[3]);
            pvoPhoto.Md5ChangeCount = Int32.Parse(tempData[4]);
            pvoPhoto.FailCount = Int32.Parse(tempData[5]);

            //Total photo summary Tec3H
            tempPath = Array.Find(emails,
                s =>
                    s.Contains(
                        "Tec3H Summary of import vehicle photo"));
            item = (MailItem)app.CreateItemFromTemplate(tempPath, Type.Missing);
            tempData = GetEmailData(item.Body);
            tec3HPhoto.VehicleToTransfer = Int32.Parse(tempData[1]);
            tec3HPhoto.PhotoCount = Int32.Parse(tempData[2]);
            tec3HPhoto.NewPhotoCount = Int32.Parse(tempData[3]);
            tec3HPhoto.Md5ChangeCount = Int32.Parse(tempData[4]);
            tec3HPhoto.FailCount = Int32.Parse(tempData[5]);

            pvo.PhotoStatus = pvoPhoto;
            tec3h.PhotoStatus = tec3HPhoto;

            //HAVAS - JSON uploaded
            tempPath = Array.Find(emails,
                s =>
                    s.Contains(
                        "HAVAS"));
            item = (MailItem)app.CreateItemFromTemplate(tempPath, Type.Missing);

            report.HavasStatus = item.Body.Contains("This mail is to inform you that the JSON file was successfully uploaded. Please do not reply to this mail");

            //PDV Name Change
            tempPath = Array.Find(emails,
                s =>
                    s.Contains(
                        "PDV Name Change"));
            item = (MailItem)app.CreateItemFromTemplate(tempPath, Type.Missing);
            report.PdvNameChange = item.Body.Replace("Batch Execution Successful \r\n\r\n", "");

            //Reset Application pool
            tempPath = Array.Find(emails,
               s =>
                   s.Contains(
                       "Reset Application Pool"));
            item = (MailItem)app.CreateItemFromTemplate(tempPath, Type.Missing);
            report.ResetApplicationPoolStatus = item.Subject.Contains("Successfully");

            //Export Info
            tempPath = Array.Find(emails,
               s =>
                   s.Contains(
                       "Export INFO"));
            item = (MailItem)app.CreateItemFromTemplate(tempPath, Type.Missing);
            report.XmlPdvFile = item.Subject.Contains("successfully");

            //Import Vehicle Status
            tempPath = Array.Find(emails,
               s =>
                   s.Contains(
                       "Import VehicleImport"));
            item = (MailItem)app.CreateItemFromTemplate(tempPath, Type.Missing);
            report.ImportVehicleStatus = item.Subject.Contains("Import VehicleImport");

            //Tec3H Missing Argus Report
            tempPath = Array.Find(emails,
               s =>
                   s.Contains(
                       "XML Tec3H Missing Argus Report"));
            item = (MailItem)app.CreateItemFromTemplate(tempPath, Type.Missing);
            tec3h.VehicleFailedArgus = item.Body.Replace("Détails des véhicules sans Argus\r\n\r\n\r\n", string.Empty).Replace("\n", "<br/>\n");

            //Lucene Generation Status
            tempPath = Array.Find(emails,
               s =>
                   s.Contains(
                       "Lucene Index"));
            item = (MailItem)app.CreateItemFromTemplate(tempPath, Type.Missing);
            report.LuceneStatus = item.Subject.Contains("successfully");

            report.PlanetVoStatus = pvo;
            report.Tec3HStatus = tec3h;

            GenerateHtml(report);
        }

        static void GenerateHtml(Report.Report report)
        {
            string outputHtmlPath = ConfigurationManager.AppSettings["OutputHtmlPath"];
            if(File.Exists(outputHtmlPath))
                File.Delete(outputHtmlPath);

            string htmlTemplate = File.ReadAllText(ConfigurationManager.AppSettings["TemplatePath"]);

            htmlTemplate = htmlTemplate.Replace("[Tec3H01]", report.Tec3HStatus.FtpSuccess.ToString());
            htmlTemplate = htmlTemplate.Replace("[Tec3H02]", report.Tec3HStatus.FtpFailure.ToString());
            htmlTemplate = htmlTemplate.Replace("[Tec3H03]", report.Tec3HStatus.ImportVehicleRecords.ToString());
            htmlTemplate = htmlTemplate.Replace("[Tec3H04]", report.Tec3HStatus.StockCount.ToString());
            htmlTemplate = htmlTemplate.Replace("[Tec3H05]", report.Tec3HStatus.NewStockCount.ToString());
            htmlTemplate = htmlTemplate.Replace("[Tec3H06]", report.Tec3HStatus.DeletedStockCount.ToString());
            htmlTemplate = htmlTemplate.Replace("[Tec3H07]", report.Tec3HStatus.PhotoStatus.TotalCount.ToString());
            htmlTemplate = htmlTemplate.Replace("[Tec3H08]", report.Tec3HStatus.PhotoStatus.VehicleToTransfer.ToString());
            htmlTemplate = htmlTemplate.Replace("[Tec3H09]", report.Tec3HStatus.PhotoStatus.PhotoCount.ToString());
            htmlTemplate = htmlTemplate.Replace("[Tec3H10]", report.Tec3HStatus.PhotoStatus.FailCount.ToString());
            htmlTemplate = htmlTemplate.Replace("[Tec3H11]", report.Tec3HStatus.VehicleFailedArgus);

            htmlTemplate = htmlTemplate.Replace("[PVO01]", report.PlanetVoStatus.FtpSuccess.ToString());
            htmlTemplate = htmlTemplate.Replace("[PVO02]", report.PlanetVoStatus.FtpFailure.ToString());
            htmlTemplate = htmlTemplate.Replace("[PVO03]", report.PlanetVoStatus.ImportVehicleRecords.ToString());
            htmlTemplate = htmlTemplate.Replace("[PVO04]", report.PlanetVoStatus.StockCount.ToString());
            htmlTemplate = htmlTemplate.Replace("[PVO05]", report.PlanetVoStatus.NewStockCount.ToString());
            htmlTemplate = htmlTemplate.Replace("[PVO06]", report.PlanetVoStatus.DeletedStockCount.ToString());
            htmlTemplate = htmlTemplate.Replace("[PVO07]", report.PlanetVoStatus.PhotoStatus.TotalCount.ToString());
            htmlTemplate = htmlTemplate.Replace("[PVO08]", report.PlanetVoStatus.PhotoStatus.VehicleToTransfer.ToString());
            htmlTemplate = htmlTemplate.Replace("[PVO09]", report.PlanetVoStatus.PhotoStatus.PhotoCount.ToString());
            htmlTemplate = htmlTemplate.Replace("[PVO10]", report.PlanetVoStatus.PhotoStatus.FailCount.ToString());

            htmlTemplate = htmlTemplate.Replace("[STATUS06]", report.HavasStatus ? "OK" : "KO");
            htmlTemplate = htmlTemplate.Replace("[STATUS08]", report.LuceneStatus ? "OK" : "KO");
            htmlTemplate = htmlTemplate.Replace("[STATUS09]", report.ResetApplicationPoolStatus ? "OK" : "KO");
            for (int i = 1; i <= 10; i++)
            {
                htmlTemplate = htmlTemplate.Replace(string.Format("[STATUS{0:00}]", i), "&nbsp;");
            }

            File.WriteAllText(outputHtmlPath, htmlTemplate);
        }

        static string[] GetEmailData(string emailBody)
        {
            bool isNewLine = true;
            string temp = string.Empty;

            foreach (char c in emailBody)
            {
                if (isNewLine == false && (c != ' ') && (c != '\r') && (c != '\n'))
                    temp += c;
                if (c == ':')
                {
                    isNewLine = false;
                    temp += ' ';
                }
                else if (c == '\n')
                    isNewLine = true;
            }

            temp += temp + " None";
            return temp.Trim().Split(' ');
        }

        static int GetPhotoCount(string emailBody)
        {
            emailBody = emailBody.Remove(0, 5);

            string result = emailBody.Where(Char.IsDigit).Aggregate(string.Empty, (current, c) => current + c);

            return Int32.Parse(result);
        }
    }
}

using Laserfiche.DocumentServices;
using Laserfiche.RepositoryAccess;
using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Text.Json;
using WorkflowActivity.Scripting;
using Document = Laserfiche.RepositoryAccess.Document;


namespace Laserfiche_Download_Issues
{
    // internal sealed class
    public class DocumentDownload : RAScriptClass110, IDocumentDownload
    {
        private readonly IAzureKeyVaultService _azureKeyVaultService;
        private readonly LaserficheConfig _lfConfig;
        private string _laserficheUsername;
        private string _laserfichePassword;
        private string _laserficheServerName;
        private string _laserficheRepoName;
        private Session laserficheSession;

        private string CurrentDirectory = System.IO.Directory.GetCurrentDirectory();

        // variables from the Laserfiche Workflow Script
        private const string Ecf_Zip_Processing = "ECF Zip Processing";
        private const string Ecf_Zip_Case_Number = "ECF Zip Case Number";
        private const string Ecf_Zip_Date = "ECF Zip Date";
        private const string Ecf_Zip_Entry_Ids = "ECF Zip Entry IDs";
        private const string Ecf_Zip_Name = "ECF Zip Name";
        private const string Path_Separator = "\\";
        private const string True = "True";
        private const string False = "False";

        private string IsZipProcessing = "";
        private List<int> ZipEntryIds = new List<int>();
        private string ZipCaseNumber = null;
        private DateTime ZipDate = DateTime.Today;
        private string ZipParentPath = null;
        private string ZipFolderPath = null;
        private string ZipFilePath = null;
        private string ZipName = null;
        private string ZipError = "Zip Processed Failed";
        // end variables from the Laserfiche Workflow Script

        public DocumentDownload(IAzureKeyVaultService azureKeyVaultService,LaserficheConfig lfConfig)
        {
            this._azureKeyVaultService = azureKeyVaultService;
            this._lfConfig = lfConfig;
            this._laserficheServerName = this._lfConfig.LaserficheServerName;
            this._laserficheRepoName = this._lfConfig.LaserficheRepoName;

            var keyVaultTask = this._azureKeyVaultService.GetKey(this._lfConfig.LaserficheKey);
            keyVaultTask.Wait();
            var keyVaultResut = keyVaultTask.Result;
            var values = System.Text.Json.JsonSerializer.Deserialize<IDictionary<string, string>>(keyVaultResut, new JsonSerializerOptions() { PropertyNameCaseInsensitive = true });
            this._laserficheUsername = values == null ? "Not Found" : values["uid"];
            this._laserfichePassword = values == null ? "Not Found" : values["pwd"];

            this.CurrentDirectory = System.IO.Directory.GetCurrentDirectory();

            Console.WriteLine($"LaserficheService running with username: {this._laserficheUsername}");
        }

        // Not used from inheritance
        protected override void Execute()
        {
            throw new NotImplementedException();
        }

        public IEnumerable<string> DownloadLaserficheDocument()
        {
            List<string> fileNames = new List<string>();

            try
            {
                RepositoryRegistration repository = new RepositoryRegistration(this._laserficheServerName, this._laserficheRepoName);
                this.laserficheSession = new Session();
                this.laserficheSession.LogIn(this._laserficheUsername, this._laserfichePassword, repository);

                using (laserficheSession)
                {
                    Console.WriteLine();
                    Console.WriteLine($"Downloading Zip File of Laserfiche Documents");
                    Console.WriteLine("...");
                    fileNames.Add(ExportZipFile(this.laserficheSession));
                    Console.WriteLine($"Zip File Path: {fileNames[0]}");
                    Console.WriteLine();
                }

            }
            catch (Exception ex)
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine(ex.Message);
            }
            finally
            {
                if (this.laserficheSession != null)
                {
                    this.laserficheSession.Close();
                }
            }

            return fileNames;
        }

        private string ExportZipFile(Session session)
        {
            // Write your code here. The BoundEntryInfo property will access the entry, RASession will get the Repository Access session
            try
            {
                string currentDirectoryPath = System.IO.Directory.GetCurrentDirectory();

                // Get all entry ids to be included in the zip file
                // TODO: User this line in the Workflow SDK Script: using (DocumentInfo zipProcessingDocument = (DocumentInfo)BoundEntryInfo)
                using (DocumentInfo zipProcessingDocument = Document.GetDocumentInfo(195924, session)) // 195924 is the entry id of the zip file processing document
                {
                    this.IsZipProcessing = (string)zipProcessingDocument.GetFieldValue(Ecf_Zip_Processing);
                    if (null != this.IsZipProcessing && !True.ToLower().Equals(this.IsZipProcessing.ToLower()))
                    {
                        // Only process this zip file if the processing flag is not true
                        var zipEntryIds = zipProcessingDocument.GetFieldValue(Ecf_Zip_Entry_Ids);

                        if (zipEntryIds != null)
                        {
                            System.Type zipEntryIdsType = zipEntryIds.GetType();
                            if (typeof(System.Object[]) == zipEntryIdsType)
                            {
                                System.Object[] entryIds = (System.Object[])zipProcessingDocument.GetFieldValue(Ecf_Zip_Entry_Ids);
                                if (entryIds != null && entryIds.Length > 0)
                                {
                                    this.ConvertLaserficheObjectArray(entryIds);
                                }
                                else
                                {
                                    this.ZipEntryIds = new List<int>();
                                }
                            }
                            else if (typeof(System.Int64) == zipEntryIdsType || typeof(System.Int32) == zipEntryIdsType)
                            {
                                this.ZipEntryIds = new List<int>() { System.Convert.ToInt32(zipEntryIds) };
                            }
                            else
                            {
                                this.ZipEntryIds = new List<int>();
                            }
                        }

                        using (FolderInfo parentFolder = zipProcessingDocument.GetParentFolder())
                        {
                            var parentPath = parentFolder.Path;
                            int indexOfPathSeparator = parentPath.IndexOf(Path_Separator);
                            string cleanPath = (indexOfPathSeparator < 0) ? parentPath : parentPath.Remove(indexOfPathSeparator, Path_Separator.Length);
                            this.CurrentDirectory = System.IO.Path.GetFullPath(this.CurrentDirectory);
                            this.ZipParentPath = System.IO.Path.Combine(this.CurrentDirectory, cleanPath);
                        }

                        this.ZipCaseNumber = (string)zipProcessingDocument.GetFieldValue(Ecf_Zip_Case_Number);
                        this.ZipDate = (DateTime)zipProcessingDocument.GetFieldValue(Ecf_Zip_Date);
                        this.ZipName = (string)zipProcessingDocument.GetFieldValue(Ecf_Zip_Name);
                    }
                }

                if (null != this.IsZipProcessing && !True.ToLower().Equals(this.IsZipProcessing.ToLower()))
                {
                    // Set the processing flag so that we block processing if another use clicks the download button at the same time
                    this.UpdateProcessingStatus(session, True);

                    string zipNameWithoutExtension = System.IO.Path.GetFileNameWithoutExtension(this.ZipName);
                    this.ZipFolderPath = System.IO.Path.Combine(this.ZipParentPath, zipNameWithoutExtension);
                    this.ZipFilePath = System.IO.Path.Combine(this.ZipParentPath, this.ZipName);

                    Console.WriteLine("ZipCaseNumber: " + this.ZipCaseNumber);
                    Console.WriteLine("ZipDate: " + this.ZipDate);
                    Console.WriteLine("ZipFolderPath: " + this.ZipFolderPath);
                    Console.WriteLine("Number of documents being processed: " + this.ZipEntryIds.Count);

                    // Always delete the previously created folder
                    this.DeleteZipFile(this.ZipFilePath);
                    this.DeleteCaseNumberFolder(this.ZipFolderPath);
                    string ecfZipFolderPath = this.CreateCaseNumberFolder(this.ZipFolderPath);
                    Console.WriteLine("EcfZipFilePath: " + ecfZipFolderPath);

                    if (!string.IsNullOrWhiteSpace(ecfZipFolderPath) && this.ZipEntryIds.Count > 0)
                    {
                        foreach (int entryId in this.ZipEntryIds)
                        {
                            using (DocumentInfo currentEntry = Document.GetDocumentInfo(entryId, session))
                            {
                                Console.WriteLine(currentEntry.Name);

                                // Call appropriate download logic
                                if (this.IsElectronicDocument(currentEntry)) // check if it is an electronic document
                                {
                                    // export the electronic portion of the document
                                    // this.ExportElectronic(currentEntry, ecfZipFolderPath);
                                    Console.WriteLine("Exported Electronic Document To PDF");
                                }
                                else if (this.HasImagePages(currentEntry)) // check for image pages
                                {
                                    // export images to pdf
                                    // this.ExportImageDocumentToPDF(currentEntry, ecfZipFolderPath);
                                    Console.WriteLine("Exported Image Document To PDF");
                                }
                                else
                                {
                                    if (HasText(currentEntry)) // if it has text
                                    {
                                        // export text file
                                        // this.ExportTextDocument(currentEntry, ecfZipFolderPath);
                                        Console.WriteLine("Exported Text Document To PDF");
                                    }
                                    else
                                    {
                                        // TODO: Probably going to get rid of this code
                                        // this.ExportEmptyDocument(currentEntry, ecfZipFolderPath);
                                        Console.WriteLine("Unable To Export Document To PDF");
                                    }
                                }
                            }
                        }

                        bool zipFolderExists = System.IO.Directory.Exists(ecfZipFolderPath);
                        if (zipFolderExists)
                        {
                            System.IO.Compression.ZipFile.CreateFromDirectory(ecfZipFolderPath, this.ZipFilePath, System.IO.Compression.CompressionLevel.Fastest, false);
                            this.UpdateProcessingStatus(session, False);
                        }
                    }
                }
            }
            catch (Exception e)
            {
                this.ZipError = e.Message;
                // WorkflowApi.TrackError(this.ZipError);
            }
            finally
            {
                // Even if everything fails. Try to update the process flag so that it can be ran again.
                this.UpdateProcessingStatus(session, False);
            }

            return this.ZipFilePath ?? this.ZipError;
        }

        private void UpdateProcessingStatus(Session session, string isProcessing)
        {
            try
            {
                // TODO: Should we set the Processing flag to false here or in the Workflow?
                // TODO: User this line in the Workflow SDK Script: using (DocumentInfo zipProcessingDocument = (DocumentInfo)BoundEntryInfo)
                using (DocumentInfo zipProcessingDocument = Document.GetDocumentInfo(195924, session)) // 195924 is the entry id of the zip file processing document
                {
                    FieldValueCollection existingValues = zipProcessingDocument.GetFieldValues();
                    this.IsZipProcessing = (string)zipProcessingDocument.GetFieldValue(Ecf_Zip_Processing);
                    if (null != this.IsZipProcessing && !isProcessing.ToLower().Equals(this.IsZipProcessing.ToLower()))
                    {
                        existingValues[Ecf_Zip_Processing] = isProcessing;
                        zipProcessingDocument.SetFieldValues(existingValues);
                        zipProcessingDocument.Save();
                    }
                }
            }
            catch (Exception updateException)
            {
                this.ZipError = updateException.Message;
                // WorkflowApi.TrackError(this.ZipError);
            }
        }

        private void ConvertLaserficheObjectArray(object[] entryIds)
        {
            this.ZipEntryIds = this.ZipEntryIds == null ? new List<int>() : this.ZipEntryIds;
            foreach (object entryId in entryIds)
            {
                System.Type entryIdType = entryId.GetType();
                if (typeof(System.Int64) == entryIdType || typeof(System.Int32) == entryIdType)
                {
                    this.ZipEntryIds.Add(System.Convert.ToInt32(entryId));
                }
            }
        }

        private void ExportEmptyDocument(DocumentInfo LFDoc, String sExportFolder)
        {
            try
            {
                System.IO.DirectoryInfo di = new System.IO.DirectoryInfo(sExportFolder);
                if (!di.Exists)
                {
                    di.Create();
                }
                String exportPath = System.IO.Path.Combine(sExportFolder, WindowsFileName(LFDoc.Name));
                System.IO.File.Create(exportPath).Dispose();
            }
            catch (Exception ex)
            {
                this.ZipError = ex.Message;
                // WorkflowApi.TrackError(this.ZipError);
            }
        }

        private void ExportTextDocument(DocumentInfo LFDoc, String sExportFolder)
        {
            try
            {
                String sTxt = null;
                System.IO.DirectoryInfo di = new System.IO.DirectoryInfo(sExportFolder);
                if (!di.Exists)
                {
                    di.Create();
                }
                String exportPath = System.IO.Path.Combine(sExportFolder, WindowsFileName(LFDoc.Name) + ".txt");
                using (PageInfoReader LF_PageInfos = LFDoc.GetPageInfos())
                {
                    foreach (PageInfo PI in LF_PageInfos)
                    {
                        if (PI.HasText)
                        {
                            using (System.IO.StreamReader reader = PI.ReadTextPagePart())
                            {
                                if (String.IsNullOrEmpty(sTxt))
                                {
                                    sTxt = reader.ReadToEnd();
                                }
                                else
                                {
                                    sTxt = sTxt + Environment.NewLine + reader.ReadToEnd();
                                }
                            }
                        }
                    }
                }
                using (System.IO.StreamWriter file = new System.IO.StreamWriter(exportPath))
                {
                    file.Write(sTxt);
                }
            }
            catch (Exception ex)
            {
                this.ZipError = ex.Message;
                // WorkflowApi.TrackError(this.ZipError);
            }
        }

        private void ExportImageDocument(DocumentInfo LFDoc, String sExportFolder)
        {
            try
            {
                System.IO.DirectoryInfo di = new System.IO.DirectoryInfo(sExportFolder);
                if (!di.Exists)
                {
                    di.Create();
                }
                Laserfiche.DocumentServices.DocumentExporter docExporter = new Laserfiche.DocumentServices.DocumentExporter();
                docExporter.PageFormat = Laserfiche.DocumentServices.DocumentPageFormat.Tiff;
                String exportPath = System.IO.Path.Combine(sExportFolder, WindowsFileName(LFDoc.Name) + ".tiff");
                docExporter.ExportPages(LFDoc, GetImagePageSet(LFDoc), exportPath);
            }
            catch (Exception ex)
            {
                this.ZipError = ex.Message;
                // WorkflowApi.TrackError(this.ZipError);
            }
        }

        private void ExportImageDocumentToPDF(DocumentInfo LFDoc, String sExportFolder)
        {
            try
            {
                System.IO.DirectoryInfo di = new System.IO.DirectoryInfo(sExportFolder);
                if (!di.Exists)
                {
                    di.Create();
                }
                Laserfiche.DocumentServices.DocumentExporter docExporter = new Laserfiche.DocumentServices.DocumentExporter();
                String exportPath = System.IO.Path.Combine(sExportFolder, WindowsFileName(LFDoc.Name) + ".pdf");
                docExporter.ExportPdf(LFDoc, GetImagePageSet(LFDoc), Laserfiche.DocumentServices.PdfExportOptions.None, exportPath);
            }
            catch (Exception ex)
            {
                this.ZipError = ex.Message;
                // WorkflowApi.TrackError(this.ZipError);
            }
        }

        private void ExportElectronic(DocumentInfo LFDoc, String sExportFolder)
        {
            try
            {
                System.IO.DirectoryInfo di = new System.IO.DirectoryInfo(sExportFolder);
                if (!di.Exists)
                {
                    di.Create();
                }
                Laserfiche.DocumentServices.DocumentExporter docExporter = new Laserfiche.DocumentServices.DocumentExporter();
                String exportPath = System.IO.Path.Combine(sExportFolder, WindowsFileName(LFDoc.Name) + "." + LFDoc.Extension);
                docExporter.ExportElecDoc(LFDoc, exportPath);
            }
            catch (Exception ex)
            {
                this.ZipError = ex.Message;
                // WorkflowApi.TrackError(this.ZipError);
            }
        }

        private PageSet GetImagePageSet(DocumentInfo LFDoc)
        {
            PageSet psReturn = new PageSet();
            try
            {
                using (PageInfoReader LF_PageInfos = (PageInfoReader)LFDoc.GetPageInfos())
                {
                    foreach (PageInfo PI in LF_PageInfos)
                    {
                        using (LaserficheReadStream lrs = PI.ReadPagePart(new PagePart()))
                        {
                            if (lrs.Length > 0)
                            {
                                psReturn.AddPage(PI.PageNumber);
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                this.ZipError = ex.Message;
                // WorkflowApi.TrackError(this.ZipError);
                psReturn = new PageSet();
            }
            return psReturn;
        }

        private Boolean HasText(DocumentInfo LFDoc)
        {
            Boolean bReturn = false;
            try
            {
                using (PageInfoReader LF_PageInfos = LFDoc.GetPageInfos())
                {
                    foreach (PageInfo PI in LF_PageInfos)
                    {
                        if (PI.HasText)
                        {
                            bReturn = true;
                            break;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                this.ZipError = ex.Message;
                // WorkflowApi.TrackError(this.ZipError);
                bReturn = false;
            }
            return bReturn;
        }

        private Boolean HasImagePages(DocumentInfo LFDoc)
        {
            Boolean bReturn = false;
            try
            {
                using (PageInfoReader LF_PageInfos = LFDoc.GetPageInfos())
                {
                    foreach (PageInfo PI in LF_PageInfos)
                    {
                        using (LaserficheReadStream lrs = PI.ReadPagePart(new PagePart()))
                        {
                            if (lrs.Length > 0)
                            {
                                bReturn = true;
                                break;
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                this.ZipError = ex.Message;
                // WorkflowApi.TrackError(this.ZipError);
                bReturn = false;
            }
            return bReturn;
        }

        private Boolean IsElectronicDocument(DocumentInfo LFDoc)
        {
            Boolean bReturn = false;
            try
            {
                bReturn = LFDoc.IsElectronicDocument;
            }
            catch (Exception ex)
            {
                this.ZipError = ex.Message;
                // WorkflowApi.TrackError(this.ZipError);
                bReturn = false;
            }
            return bReturn;
        }

        private Boolean IsEntryDocument(EntryInfo LFEnt)
        {
            Boolean bReturn = false;
            try
            {
                if (LFEnt.EntryType == EntryType.Document)
                {
                    bReturn = true;
                }
                else
                {
                    bReturn = false;
                }
            }
            catch (Exception ex)
            {
                this.ZipError = ex.Message;
                // WorkflowApi.TrackError(this.ZipError);
                bReturn = false;
            }
            return bReturn;
        }

        private String WindowsFileName(String LFDocName)
        {
            String sReturn = LFDocName.Replace(@"/", "_");
            sReturn = sReturn.Replace(@"\", "_");
            sReturn = sReturn.Replace(@":", "_");
            sReturn = sReturn.Replace(@"*", "_");
            sReturn = sReturn.Replace(@"?", "_");
            sReturn = sReturn.Replace(@"""", "_");
            sReturn = sReturn.Replace(@"<", "_");
            sReturn = sReturn.Replace(@">", "_");
            sReturn = sReturn.Replace(@"|", "_");
            return sReturn;
        }

        private string CreateCaseNumberFolder(string ecfZipFolderPath)
        {
            try
            {
                System.IO.DirectoryInfo di = new System.IO.DirectoryInfo(ecfZipFolderPath);
                if (!di.Exists)
                {
                    di.Create();
                }
                return di.FullName;
            }
            catch (Exception ex)
            {
                this.ZipError = ex.Message;
                // WorkflowApi.TrackError(this.ZipError);
                return null;
            }
        }

        private void DeleteCaseNumberFolder(string ecfZipFolderPath)
        {
            try
            {
                System.IO.DirectoryInfo di = new System.IO.DirectoryInfo(ecfZipFolderPath);
                if (di.Exists)
                {
                    foreach (FileInfo file in di.GetFiles())
                    {
                        file.Delete();
                    }
                    di.Delete();
                }
            }
            catch (Exception ex)
            {
                this.ZipError = ex.Message;
                Console.WriteLine(ex.Message);
                // WorkflowApi.TrackError(this.ZipError);
            }
        }

        private void DeleteZipFile(string ecfZipFilePath)
        {
            try
            {
                if (System.IO.File.Exists(ecfZipFilePath))
                {
                    System.IO.File.Delete(ecfZipFilePath);
                }
            }
            catch (Exception ex)
            {
                this.ZipError = ex.Message;
                Console.WriteLine(ex.Message);
                // WorkflowApi.TrackError(this.ZipError);
            }
        }

    }
}

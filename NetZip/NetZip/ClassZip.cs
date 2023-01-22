using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.IO.Compression;
using System.IO;
using System.Runtime.InteropServices;
using System.Windows;
using System.Globalization;
using System.Security.Cryptography;

namespace NetZip
{
    public class ClassZip
    {
        public ClassZip()
        {
            return;
        }
        // Internal Data
        private string _sourceDirectory = null;
        private string _extractDirectory = null;
        private string[] _excludeFileMask = null;
        private string[] _includeFileMask = null;
        private string[] _noCompressSuffixes = null;
        private string _password = null;
        private string _zipFileName = null;
        private string _lastErrorMsg = null;

        private int _version = 1;
        private int _lastErrorCode = 0;
        private bool _recurseSubDirectories = false;
        private bool _includeBaseDirectory = false;
        private bool _overwrite = false;
        private CompressionLevel _compressionLevel = CompressionLevel.Fastest;
        
        public enum UserErrorCodes
        {
            errVoidArgs = -1, // Void args
            errUnknown = -2
        };


        // Public accsesors
        public int Version { get => _version; protected set => _version = value; }
        public int LastErrorCode { get => _lastErrorCode; protected set => _lastErrorCode = value; }
        public bool RecurseSubDirectories { get => _recurseSubDirectories; set => _recurseSubDirectories = value; }
        public bool IncludeBaseDirectory { get => _includeBaseDirectory; set => _includeBaseDirectory = value; }
        public bool Overwrite { get => _overwrite; set => _overwrite = value; }
        public CompressionLevel CompLevel { get => _compressionLevel; set => _compressionLevel = value; }
        
        public string LastErrorMsg { get => _lastErrorMsg; protected set => _lastErrorMsg = value; }
        public string SourceDirectory { get => _sourceDirectory; set => _sourceDirectory = value; }
        public string ExtractDirectory { get => _extractDirectory; set => _extractDirectory = value; }
        public string[] ExcludeFileMask { get => _excludeFileMask; set => _excludeFileMask = value; }
        public string[] IncludeFileMask { get => _includeFileMask; set => _includeFileMask = value; }
        public string[] NoCompressSuffixes { get => _noCompressSuffixes; set => _noCompressSuffixes = value; }
        public string Password { get => _password; set => _password = value; }
        public string ZipFileName { get => _zipFileName; set => _zipFileName = value; }

        //----------------------------------------------------------------------
        /// <summary>
        /// Adds and compress files in ZIP format to a .zip file
        /// </summary>
        /// <param name="create">TRUE for creating a new ZIP file; false to append in an existing ZIP file</param>
        /// <returns>0 on success; error code otherwise</returns>
        public Int32 Add(bool create)
        {
            LastErrorCode = 0;
            LastErrorMsg = string.Empty;
            
            if (string.IsNullOrEmpty(SourceDirectory) ||
                string.IsNullOrEmpty(ZipFileName))
            {
                LastErrorCode = (int)UserErrorCodes.errVoidArgs;
                LastErrorMsg = string.Empty;
                return LastErrorCode;
            }

            // Try if a simple directory compression is possible
            if (create == true &&
                RecurseSubDirectories == true &&
                ExcludeFileMask.Count() == 0 &&
                NoCompressSuffixes.Count() == 0 &&
                (IncludeFileMask.Count() == 0 || (IncludeFileMask.Count() == 1 && IncludeFileMask[0] == "*")))
            {
                // Compress Directory
                try
                {
                    if (File.Exists(ZipFileName))
                        File.Delete(ZipFileName);
                    ZipFile.CreateFromDirectory(SourceDirectory, ZipFileName, CompLevel, IncludeBaseDirectory);
                    return 0;
                }
                catch (Exception ex)
                {
                    LastErrorMsg = ex.Message;
                    LastErrorCode = ex.HResult;
                    return ex.HResult;
                }
            }

            // General procedure
            // Adds files one to one
            FileMode mode = FileMode.Open;
            if (!File.Exists(ZipFileName) || create)
                mode = FileMode.Create;
            SearchOption serahOpt = RecurseSubDirectories ? SearchOption.AllDirectories : SearchOption.TopDirectoryOnly;
            try
            {
                using (FileStream zipToOpen = new FileStream(ZipFileName, mode))
                {
                    using (ZipArchive archive = new ZipArchive(zipToOpen, ZipArchiveMode.Update))
                    {
                        // 1.- Search files to compress
                        var filenames = Directory.EnumerateFiles(SourceDirectory, "*", serahOpt);
                        foreach (var filename in filenames)
                        {
                            try
                            {
                                // fn: filename (and reñative parh) to be shown in zip file
                                string fn = filename;
                                if (!IncludeBaseDirectory)
                                    // skip name of the base directory
                                    fn = filename.Substring(SourceDirectory.Length + 1);

                                //-----------------------

                                // 2.- Suffixes
                                if (NoCompressSuffixes.Count() > 0)
                                {
                                    bool found = false;
                                    foreach (string strExt in NoCompressSuffixes)
                                    {
                                        if (fn.EndsWith(strExt, StringComparison.OrdinalIgnoreCase))
                                        {
                                            found = true;
                                            break;
                                        }
                                    }
                                    if (found)
                                        continue; // skip this file
                                }

                                // 3.- Exluded files
                                if (ExcludeFileMask.Count() > 0)
                                {
                                    bool found = false;
                                    foreach (string strMask in ExcludeFileMask)
                                    {
                                        if (MatchFullPathMask(strMask, fn))
                                        {
                                            found = true;
                                            break;
                                        }
                                    }
                                    if (found)
                                        continue; // skip file
                                }

                                // 4.- Included files
                                if (IncludeFileMask.Count() > 0)
                                {
                                    bool found = false;
                                    foreach (string strMask in IncludeFileMask)
                                    {
                                        if (MatchFullPathMask(strMask, fn))
                                        {
                                            found = true;
                                            break;
                                        }
                                    }
                                    if (!found)
                                        continue; // skip file
                                }

                                ZipArchiveEntry entry = archive.CreateEntry(fn);
                                using (StreamWriter writer = new StreamWriter(entry.Open()))
                                {
                                    writer.WriteLine("Information about this package.");
                                    writer.WriteLine("========================");
                                }
                            }
                            catch (Exception ex)
                            {
                                LastErrorMsg = ex.Message;
                                LastErrorCode = ex.HResult;
                                return ex.HResult;
                            }
                        } //foreach (var filename in filenames)
                    } //using (ZipArchive archive = new ZipArchive(zipToOpen, ZipArchiveMode.Update))
                } //using (FileStream zipToOpen = new FileStream(ZipFileName, mode))
            }
            catch (Exception ex)
            {
                LastErrorMsg = ex.Message;
                LastErrorCode = ex.HResult;
                return ex.HResult;
            }

            return LastErrorCode;
        } //Add

        //----------------------------------------------------------------------
        /// <summary>
        /// Extract from an existing zip file
        /// </summary>
        /// <returns>0 if success, error code otherwise</returns>
        public Int32 Extract()
        {
            LastErrorCode = 0;
            LastErrorMsg = string.Empty;

            if (string.IsNullOrEmpty(ExtractDirectory) ||
                string.IsNullOrEmpty(ZipFileName))
            {
                LastErrorCode = (int)UserErrorCodes.errVoidArgs;
                LastErrorMsg = string.Empty;
                return LastErrorCode;
            }

            // Try if a simple extract of all files is poccible
            if (RecurseSubDirectories == true &&
                ExcludeFileMask.Count() == 0 &&
                NoCompressSuffixes.Count() == 0 &&
                (IncludeFileMask.Count() == 0 || (IncludeFileMask.Count() == 1 && IncludeFileMask[0] == "*")))
            {
                //+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
                // Extract Full Directory
                //+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
                try
                {
                    ZipFile.ExtractToDirectory(ZipFileName, ExtractDirectory);
                    return 0;
                }
                catch (Exception ex)
                {
                    LastErrorMsg = ex.Message;
                    LastErrorCode = ex.HResult;
                    return ex.HResult;
                }
            }

            //+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
            // Extract with options
            //+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
            try
            {
                using (ZipArchive archive = ZipFile.OpenRead(ZipFileName))
                {
                    foreach (ZipArchiveEntry entry in archive.Entries)
                    {
                        try
                        {
                            // Gets the full path to ensure that relative segments are removed.
                            string destinationPath = Path.GetFullPath(Path.Combine(ExtractDirectory, entry.FullName));
                            // Ensure no redirection paths
                            if (!destinationPath.StartsWith(ExtractDirectory, StringComparison.Ordinal))
                                continue;

                            string destinationDirName = Path.GetDirectoryName(destinationPath);
                            if (!string.IsNullOrEmpty(destinationDirName) && !destinationDirName.EndsWith(Path.DirectorySeparatorChar.ToString(), StringComparison.Ordinal))
                                destinationDirName += Path.DirectorySeparatorChar;

                            // 1.- Subdirectories
                            if (String.Compare(destinationDirName, ExtractDirectory, true) != 0 && RecurseSubDirectories == false)
                                continue; // skip

                            // 2.- Suffixes
                            if (NoCompressSuffixes.Count() > 0)
                            {
                                bool found = false;
                                foreach (string strExt in NoCompressSuffixes)
                                {
                                    if (entry.FullName.EndsWith(strExt, StringComparison.OrdinalIgnoreCase))
                                    {
                                        found = true;
                                        break;
                                    }
                                }
                                if (found)
                                    continue; //skip file
                            }

                            // 3.- Exluded files
                            if (ExcludeFileMask.Count() > 0)
                            {
                                bool found = false;
                                foreach (string strMask in ExcludeFileMask)
                                {
                                    if (MatchFullPathMask(strMask, destinationPath))
                                    {
                                        found = true;
                                        break;
                                    }
                                }
                                if (found)
                                    continue; //skip file
                            }

                            // 4.- Included files
                            if (IncludeFileMask.Count() > 0)
                            {
                                bool found = false;
                                foreach (string strMask in IncludeFileMask)
                                {
                                    if (MatchFullPathMask(strMask, destinationPath))
                                    {
                                        found = true;
                                        break;
                                    }
                                }
                                if (!found)
                                    continue; // skip file
                            }

                            // Create subdirectory if needed
                            if (String.Compare(destinationDirName, ExtractDirectory, true) != 0)
                            {
                                //Its a subdirectory
                                if (!Directory.Exists(destinationDirName))
                                    //Crear
                                    Directory.CreateDirectory(destinationDirName);
                            }

                            entry.ExtractToFile(destinationPath, Overwrite);
                        }
                        catch (Exception ex)
                        {
                            LastErrorMsg = ex.Message;
                            LastErrorCode = ex.HResult;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                LastErrorMsg = ex.Message;
                LastErrorCode = ex.HResult;
                return ex.HResult;
            }
            

            // Not covered: password, hiden files
            return LastErrorCode;
        } //public Int32 Extract()

        /// <summary>
        /// Auxilar and interbal class to perform wildcards ('*' and '?')
        /// </summary>
        private class wildcard
        {
            public int loc = -1;
            public bool isAsterisc = true;
        }

        /// <summary>
        /// Test if a filename matches a file mask, Auxiliar function for MatchFullPathMask
        /// </summary>
        /// <param name="fileMask">mask, with wildcards but without directories</param>
        /// <param name="fileName">filename, wkthout path</param>
        /// <returns>true for match</returns>
        private bool MatchFileMask(string fileMask, string fileName)
        {
            // Obvious cases
            if (String.IsNullOrEmpty(fileMask))
                return true;
            if (String.Compare(fileMask, "*") == 0)
                return true;

            if (fileMask.IndexOf("*", StringComparison.Ordinal) == -1 && fileMask.IndexOf("?", StringComparison.Ordinal) == -1)
            {
                // There are no wildcards
                if (String.Compare(fileMask, fileName, true) == 0)
                    return true;
                else
                    return false;
            }

            // There are wildcards
        
            List<wildcard> wildcards = new List<wildcard>();
            int loc = -1, nast = 0, nques = 0;
            do
            {
                wildcard uwc = new wildcard();
                int loc1 = fileMask.IndexOf("*", loc + 1, StringComparison.Ordinal);
                int loc2 = fileMask.IndexOf("?", loc + 1, StringComparison.Ordinal);
                if (loc1 >= 0 && loc2 >= 0)
                {
                    loc = Math.Min(loc1, loc2);
                    uwc.loc = loc;
                    uwc.isAsterisc = loc1 < loc2;
                    wildcards.Add(uwc);
                    if (loc1 < loc2) nast++;
                    else nques++;
                }
                else if (loc1 >= 0)
                {
                    loc = loc1;
                    uwc.loc = loc;
                    uwc.isAsterisc = true;
                    wildcards.Add(uwc);
                    nast++;
                }
                else if (loc2 >= 0)
                {
                    loc = loc2;
                    uwc.loc = loc;
                    uwc.isAsterisc = false;
                    wildcards.Add(uwc);
                    nques++;
                }
                else
                    loc = -1;
            } while (loc != -1) ;

            if (nast > 0)
            {
                // There are asteriscs (and may be question marks)
                int nwc = wildcards.Count;
                int locd = 0;
                for (int iwc = 0; iwc <= nwc; iwc++)
                {
                    int ini = iwc == 0 ? 0 : wildcards[iwc - 1].loc + 1;
                    int fin = iwc == nwc ? fileMask.Length - 1 : wildcards[iwc].loc - 1;
                    bool isAsterisc = iwc == 0 ? false : wildcards[iwc - 1].isAsterisc;
                    if (ini >= fileMask.Length || fin < ini)
                        continue;
                    String imask = fileMask.Substring(ini, fin - ini + 1);
                    String ifile = fileName.Substring(locd);
                    int locOk = ifile.IndexOf(imask, StringComparison.OrdinalIgnoreCase);
                    if (locOk < 0)
                        return false;
                    if (!isAsterisc && locOk != 0)
                        return false;
                    locd = locOk + fin - ini + 2;
                }
            }
            else
            {
                // Only question marks
                for (int iwc=0; iwc<=nques; iwc++)
                {
                    int ini = iwc == 0 ? 0 : wildcards[iwc - 1].loc + 1;
                    int fin = iwc == nques ? fileMask.Length - 1 : wildcards[iwc].loc - 1;
                    if (ini >= fileMask.Length || fin < ini)
                        continue;
                    String imask = fileMask.Substring(ini, fin - ini + 1);
                    String ifile = fileName.Substring(ini);
                    if (ifile.IndexOf(imask, StringComparison.OrdinalIgnoreCase) != 0)
                        return false;
                }
            }

            return true;
        }

        /// <summary>
        /// Checks a full path name with a mask
        /// </summary>
        /// <param name="filePathMask">Full file path mask with wildcards</param>
        /// <param name="filePathName">Full file path name</param>
        /// <returns>true if matches</returns>
        private bool MatchFullPathMask(string filePathMask, string filePathName)
        {
            string[] splitMask = filePathMask.Split(Path.DirectorySeparatorChar);
            string[] splitName = filePathName.Split(Path.DirectorySeparatorChar);
            int nMask = splitMask.Length;
            int nName = splitName.Length;

            if (nMask > nName || nMask < nName - 1)
                return false;
            for (int iwc = 0; iwc < nMask; iwc++)
                if (!MatchFileMask(splitMask[iwc], splitName[iwc]))
                    return false;

            return true;
        }

        public static string getFileHash256(string fullFileName)
        {
            string hashString = string.Empty;

            // Initialize a SHA256 hash object.
            using (SHA256 mySHA256 = SHA256.Create())
            {
                FileInfo fInfo = new FileInfo(fullFileName);
                if (!fInfo.Exists)
                    return hashString;
                using (FileStream fileStream = fInfo.Open(FileMode.Open))
                {
                    try
                    {
                        // Create a fileStream for the file.
                        // Be sure it's positioned to the beginning of the stream.
                        fileStream.Position = 0;
                        // Compute the hash of the fileStream.
                        byte[] hashValue = mySHA256.ComputeHash(fileStream);
                        // Write the name and hash value of the file to the console.
                        StringBuilder sb = new StringBuilder("", 64);
                        for (int i = 0; i < hashValue.Length; i++)
                        {
                            sb.Append($"{hashValue[i]:X2}");
                        }
                        hashString = sb.ToString();
                    }
                    catch (IOException e)
                    {
                        hashString = ($"I/O Exception: {e.Message}");
                    }
                    catch (UnauthorizedAccessException e)
                    {
                        hashString = ($"Access Exception: {e.Message}");
                    }
                }
            }

            return hashString;
        }
    } //public class ClassZip

    /// <summary>
    /// Assembly GUID and Interface
    /// </summary>
    [Guid("82508DF9-C406-4B57-808B-C23B41B9A633")]
    public interface INetZipInterface
    {
        bool Create();
        void SetRecurseSubDirectories(bool recurseSubDirectories);
        void SetIncludeBaseDirectory(bool includeBaseDirectory);
        void SetOverwrite(bool overwrite);
        void SetCompressionLevel(Int32 compressionLevel);

        void SetSourceDirectory(string sourceDir);
        void SetExtractDirectory(string extractDir);
        void SetExcludeFileMask(string excludeFileMask, char separator);
        void SetIncludeFileMask(string includeFileMask, char separator);
        void SetNoCompressSuffixes(string noCompressSuffixes, char separator);
        void SetPassword(string password);
        void SetZipFileName(string zipFileName);


        int GetVersion();
        bool GetRecurseSubDirectories();
        bool GetIncludeBaseDirectory();
        bool GetOverwrite();
        Int32 GetCompressionLevel();

        string GetSourceDirectory();
        string GetExtractDirectory();
        string GetExcludeFileMask(char separator);
        string GetIncludeFileMask(char separator);
        string GetNoCompressSuffixes(char separator);
        string GetPassword();
        string GetZipFileName();

        //+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
        // Operations
        //+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

        /// <summary>
        /// Compress in ZIP format
        /// </summary>
        /// <param name="create">TRUE for creating a new ZIP file; false to append in an existing ZIP file</param>
        /// <returns></returns>
        Int32 Add(bool create);

        /// <summary>
        /// Extracts from an existing zip file
        /// </summary>
        /// <returns></returns>
        Int32 Extract();

        /// <summary>
        /// Return last error condition, if any
        /// </summary>
        /// <param name="hResult">Error code (0 == S_OK for success</param>
        /// <returns>Error Message</returns>
        string GetLastZipError(ref Int32 hResult);

        //++++++++++++++++++++++
        string GetFileHash256(string fullFileName);
    }
    /// <summary>
    /// Assembly GUID and implementation of the interface
    /// </summary>
    [Guid("5E2A5A32-FA8D-4941-8E95-F4B8F5BB9CCE")]
    public class InterfaceImplementation : INetZipInterface
    {
        private ClassZip zipI = null;

        public bool Create()
        {
            try
            {
                zipI = new ClassZip();
                return zipI != null;
            }
            catch
            {
                throw new NotImplementedException();
            }
        }

        public void SetRecurseSubDirectories(bool recurseSubDirectories) { zipI.RecurseSubDirectories = recurseSubDirectories; }
        public void SetIncludeBaseDirectory(bool includeBaseDirectory) { zipI.IncludeBaseDirectory = includeBaseDirectory; }
        public void SetOverwrite(bool overwrite) { zipI.Overwrite = overwrite; }
        public void SetCompressionLevel(Int32 compressionLevel) {
            if (compressionLevel == 0) zipI.CompLevel = CompressionLevel.NoCompression;
            else if (compressionLevel <= 5) zipI.CompLevel = CompressionLevel.Fastest;
            else zipI.CompLevel = CompressionLevel.Optimal;
        }

        public void SetSourceDirectory(string soureDir)
        {
            // Ensures that the last character on the extraction path is the directory separator char.
            if (!string.IsNullOrEmpty(soureDir) && !soureDir.EndsWith(Path.DirectorySeparatorChar.ToString(), StringComparison.Ordinal))
                soureDir += Path.DirectorySeparatorChar;
            zipI.SourceDirectory = soureDir;
        }
        public void SetExtractDirectory(string extractDir)
        {
            // Ensures that the last character on the extraction path is the directory separator char.
            if (!string.IsNullOrEmpty(extractDir) && !extractDir.EndsWith(Path.DirectorySeparatorChar.ToString(), StringComparison.Ordinal))
                extractDir += Path.DirectorySeparatorChar;
            zipI.ExtractDirectory = extractDir;
        }
        public void SetExcludeFileMask(string excludeFileMask, char separator)
        {
            char[] separators = new char[] { separator };
            zipI.ExcludeFileMask = excludeFileMask.Split(separators, StringSplitOptions.RemoveEmptyEntries);
        }
        public void SetIncludeFileMask(string includeFileMask, char separator)
        {
            char[] separators = new char[] { separator };
            zipI.IncludeFileMask = includeFileMask.Split(separators, StringSplitOptions.RemoveEmptyEntries);
        }
        public void SetNoCompressSuffixes(string noCompressSuffixes, char separator)
        {
            char[] separators = new char[] {separator};
            zipI.NoCompressSuffixes = noCompressSuffixes.Split(separators, StringSplitOptions.RemoveEmptyEntries);
        }
        public void SetPassword(string password) {
            zipI.Password = password;
            if (!string.IsNullOrEmpty(password))
                throw new NotImplementedException();
        }
        public void SetZipFileName(string zipFileName) { zipI.ZipFileName = zipFileName; }

        public string GetSourceDirectory() { return zipI.SourceDirectory; }
        public string GetExtractDirectory() { return zipI.ExtractDirectory; }
        public string GetExcludeFileMask(char separator) {
            StringBuilder sb = new StringBuilder();
            for (int i = 0; i < zipI.NoCompressSuffixes.Count(); i++)
            {
                if (i > 0)
                    sb.Append(separator);
                sb.Append(zipI.ExcludeFileMask[i]);
            }
            return sb.ToString();
        }
        public string GetIncludeFileMask(char separator) {
            StringBuilder sb = new StringBuilder();
            for (int i = 0; i < zipI.NoCompressSuffixes.Count(); i++)
            {
                if (i > 0)
                    sb.Append(separator);
                sb.Append(zipI.IncludeFileMask[i]);
            }
            return sb.ToString();
        }
        public string GetNoCompressSuffixes(char separator)
        {
            StringBuilder sb = new StringBuilder();
            for (int i = 0; i < zipI.NoCompressSuffixes.Count(); i++)
            {
                if(i > 0)
                    sb.Append(separator);
                sb.Append(zipI.NoCompressSuffixes[i]);
            }
            return sb.ToString();
        }
        public string GetPassword() {
            throw new NotImplementedException();
            //return zipI.Password;
        }
        public string GetZipFileName() { return zipI.ZipFileName; }

        public int GetVersion() { return zipI.Version; }
        public bool GetRecurseSubDirectories() { return zipI.RecurseSubDirectories; }
        public bool GetIncludeBaseDirectory() { return zipI.IncludeBaseDirectory; }
        public bool GetOverwrite() { return zipI.Overwrite; }
        public Int32 GetCompressionLevel()
        {
            switch (zipI.CompLevel)
            {
                case CompressionLevel.Optimal: return 9;
                case CompressionLevel.Fastest: return 5;
                case CompressionLevel.NoCompression: return 0;
                default: return 0;
            }
        }
        public Int32 Add(bool create) { return zipI.Add(create); }
        public Int32 Extract() { return zipI.Extract(); }
        public string GetLastZipError(ref Int32 hResult)
        {
            hResult = zipI.LastErrorCode;
            return zipI.LastErrorMsg;
        }

        public string GetFileHash256(string fullFileName)
        {
            return ClassZip.getFileHash256(fullFileName);
        }
    }
}

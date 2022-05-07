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

namespace NetZip
{
    public class ClassZip
    {
        public ClassZip()
        {
            return;
        }
        //Data
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


        //Accesores públicos
        public int m_Version { get => _version; protected set => _version = value; }
        public int m_LastErrorCode { get => _lastErrorCode; protected set => _lastErrorCode = value; }
        public bool m_RecurseSubDirectories { get => _recurseSubDirectories; set => _recurseSubDirectories = value; }
        public bool m_IncludeBaseDirectory { get => _includeBaseDirectory; set => _includeBaseDirectory = value; }
        public bool m_Overwrite { get => _overwrite; set => _overwrite = value; }
        public CompressionLevel m_CompressionLevel { get => _compressionLevel; set => _compressionLevel = value; }
        
        public string m_LastErrorMsg { get => _lastErrorMsg; protected set => _lastErrorMsg = value; }
        public string m_SourceDirectory { get => _sourceDirectory; set => _sourceDirectory = value; }
        public string m_ExtractDirectory { get => _extractDirectory; set => _extractDirectory = value; }
        public string[] m_ExcludeFileMask { get => _excludeFileMask; set => _excludeFileMask = value; }
        public string[] m_IncludeFileMask { get => _includeFileMask; set => _includeFileMask = value; }
        public string[] m_NoCompressSuffixes { get => _noCompressSuffixes; set => _noCompressSuffixes = value; }
        public string m_Password { get => _password; set => _password = value; }
        public string m_ZipFileName { get => _zipFileName; set => _zipFileName = value; }

        //----------------------------------------------------------------------
        /// <summary>
        /// Compress in ZIP format
        /// </summary>
        /// <param name="create">TRUE for creating a new ZIP file; false to append in an existing ZIP file</param>
        /// <returns></returns>
        public Int32 Add(bool create)
        {
            m_LastErrorCode = 0;
            m_LastErrorMsg = string.Empty;
            
            if (string.IsNullOrEmpty(m_SourceDirectory) ||
                string.IsNullOrEmpty(m_ZipFileName))
            {
                m_LastErrorCode = -1;
                m_LastErrorMsg = "Hay parámetros vacíos";
                return m_LastErrorCode;
            }

            if (create == true &&
                m_RecurseSubDirectories == true &&
                m_ExcludeFileMask.Count() == 0 &&
                m_NoCompressSuffixes.Count() == 0 &&
                (m_IncludeFileMask.Count() == 0 || (m_IncludeFileMask.Count() == 1 && m_IncludeFileMask[0] == "*")))
            {
                //Compress Directory
                try
                {
                    ZipFile.CreateFromDirectory(m_SourceDirectory, m_ZipFileName, m_CompressionLevel, m_IncludeBaseDirectory);
                    return 0;
                }
                catch (Exception ex)
                {
                    m_LastErrorMsg = ex.Message;
                    m_LastErrorCode = ex.HResult;
                    return ex.HResult;
                }
            }
            return 0;
        }

        //----------------------------------------------------------------------
        /// <summary>
        /// Extract from an existing zip file
        /// </summary>
        /// <returns></returns>
        public Int32 Extract()
        {
            m_LastErrorCode = 0;
            m_LastErrorMsg = string.Empty;

            if (string.IsNullOrEmpty(m_ExtractDirectory) ||
                string.IsNullOrEmpty(m_ZipFileName))
            {
                m_LastErrorCode = -1;
                m_LastErrorMsg = "Hay parámetros vacíos";
                return m_LastErrorCode;
            }

            if (m_RecurseSubDirectories == true &&
                m_ExcludeFileMask.Count() == 0 &&
                m_NoCompressSuffixes.Count() == 0 &&
                (m_IncludeFileMask.Count() == 0 || (m_IncludeFileMask.Count() == 1 && m_IncludeFileMask[0] == "*")))
            {
                //+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
                // Extract Full Directory
                //+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
                try
                {
                    ZipFile.ExtractToDirectory(m_ZipFileName, m_ExtractDirectory);
                    return 0;
                }
                catch (Exception ex)
                {
                    m_LastErrorMsg = ex.Message;
                    m_LastErrorCode = ex.HResult;
                    return ex.HResult;
                }
            }
            else
            {
                //+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
                // Extract with options
                //+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
                try
                {
                    using (ZipArchive archive = ZipFile.OpenRead(m_ZipFileName))
                    {
                        foreach (ZipArchiveEntry entry in archive.Entries)
                        {
                            try
                            {
                                // Gets the full path to ensure that relative segments are removed.
                                string destinationPath = Path.GetFullPath(Path.Combine(m_ExtractDirectory, entry.FullName));
                                // Ensure no redirection paths
                                if (!destinationPath.StartsWith(m_ExtractDirectory, StringComparison.Ordinal))
                                    continue;

                                string destinationDirName = Path.GetDirectoryName(destinationPath);
                                if (!string.IsNullOrEmpty(destinationDirName) && !destinationDirName.EndsWith(Path.DirectorySeparatorChar.ToString(), StringComparison.Ordinal))
                                    destinationDirName += Path.DirectorySeparatorChar;

                                //1. Subdirectories
                                if (String.Compare(destinationDirName, m_ExtractDirectory, true) != 0 && m_RecurseSubDirectories == false)
                                    continue;

                                //2. Suffixes
                                if (m_NoCompressSuffixes.Count() > 0)
                                {
                                    bool found = false;
                                    foreach (string strExt in m_NoCompressSuffixes)
                                    {
                                        if (entry.FullName.EndsWith(strExt, StringComparison.OrdinalIgnoreCase))
                                        {
                                            found = true;
                                            break;
                                        }
                                    }
                                    if (found)
                                        continue;
                                }

                                //3.Exluded files
                                if (m_ExcludeFileMask.Count() > 0)
                                {
                                    bool found = false;
                                    foreach (string strMask in m_ExcludeFileMask)
                                    {
                                        if (MatchFullPathMask(strMask, destinationPath))
                                        {
                                            found = true;
                                            break;
                                        }
                                    }
                                    if (found)
                                        continue;
                                }

                                //4. Included files
                                if (m_IncludeFileMask.Count() > 0)
                                {
                                    bool found = false;
                                    foreach (string strMask in m_IncludeFileMask)
                                    {
                                        if (MatchFullPathMask(strMask, destinationPath))
                                        {
                                            found = true;
                                            break;
                                        }
                                    }
                                    if (!found)
                                        continue;
                                }

                                //Create subdirectory if needed
                                if (String.Compare(destinationDirName, m_ExtractDirectory, true) != 0)
                                {
                                    //Its a subdirectory
                                    if (!Directory.Exists(destinationDirName))
                                        //Crear
                                        Directory.CreateDirectory(destinationDirName);
                                }

                                entry.ExtractToFile(destinationPath, m_Overwrite);
                            }
                            catch (Exception ex)
                            {
                                m_LastErrorMsg = ex.Message;
                                m_LastErrorCode = ex.HResult;
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    m_LastErrorMsg = ex.Message;
                    m_LastErrorCode = ex.HResult;
                    return ex.HResult;
                }
            }

            // Resto de casos
            return m_LastErrorCode;
        } //public Int32 Extract()

        private class wildcard
        {
            public int loc = -1;
            public bool isAsterisc = true;
        }

        /// <summary>
        /// Test if a filename matches a fila mask, Auxiliar function for MatchFullPathMask
        /// </summary>
        /// <param name="fileMask">mask, with wildcards but without directories</param>
        /// <param name="fileName">filename, wkthout path</param>
        /// <returns></returns>
        private bool MatchFileMask(string fileMask, string fileName)
        {
            if (String.IsNullOrEmpty(fileMask))
                return true;
            if (String.Compare(fileMask, "*") == 0)
                return true;

            if (fileMask.IndexOf("*", StringComparison.Ordinal) == -1 && fileMask.IndexOf("?", StringComparison.Ordinal) == -1)
            {
                //There is no wildcards
                if (String.Compare(fileMask, fileName, true) == 0)
                    return true;
                else
                    return false;
            }

            //There are wildcards
            //Regex emailregex = new Regex("(?<user>[^@]+)@(?<host>.+)");
            //Match m = emailregex.Match(fileName);

        
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
                //Hay asteriscos (y puede que interrogaciones)
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
                //Sólo hay interrogaciones
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
        /// <returns>true fi matches</returns>
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
    } //public class ClassZip

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
    }
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

        public void SetRecurseSubDirectories(bool recurseSubDirectories) { zipI.m_RecurseSubDirectories = recurseSubDirectories; }
        public void SetIncludeBaseDirectory(bool includeBaseDirectory) { zipI.m_IncludeBaseDirectory = includeBaseDirectory; }
        public void SetOverwrite(bool overwrite) { zipI.m_Overwrite = overwrite; }
        public void SetCompressionLevel(Int32 compressionLevel) {
            if (compressionLevel == 0) zipI.m_CompressionLevel = CompressionLevel.NoCompression;
            else if (compressionLevel <= 5) zipI.m_CompressionLevel = CompressionLevel.Fastest;
            else zipI.m_CompressionLevel = CompressionLevel.Optimal;
        }

        public void SetSourceDirectory(string soureDir)
        {
            // Ensures that the last character on the extraction path is the directory separator char.
            if (!string.IsNullOrEmpty(soureDir) && !soureDir.EndsWith(Path.DirectorySeparatorChar.ToString(), StringComparison.Ordinal))
                soureDir += Path.DirectorySeparatorChar;
            zipI.m_SourceDirectory = soureDir;
        }
        public void SetExtractDirectory(string extractDir)
        {
            // Ensures that the last character on the extraction path is the directory separator char.
            if (!string.IsNullOrEmpty(extractDir) && !extractDir.EndsWith(Path.DirectorySeparatorChar.ToString(), StringComparison.Ordinal))
                extractDir += Path.DirectorySeparatorChar;
            zipI.m_ExtractDirectory = extractDir;
        }
        public void SetExcludeFileMask(string excludeFileMask, char separator)
        {
            char[] separators = new char[] { separator };
            zipI.m_ExcludeFileMask = excludeFileMask.Split(separators, StringSplitOptions.RemoveEmptyEntries);
        }
        public void SetIncludeFileMask(string includeFileMask, char separator)
        {
            char[] separators = new char[] { separator };
            zipI.m_IncludeFileMask = includeFileMask.Split(separators, StringSplitOptions.RemoveEmptyEntries);
        }
        public void SetNoCompressSuffixes(string noCompressSuffixes, char separator)
        {
            char[] separators = new char[] {separator};
            zipI.m_NoCompressSuffixes = noCompressSuffixes.Split(separators, StringSplitOptions.RemoveEmptyEntries);
        }
        public void SetPassword(string password) { zipI.m_Password = password; }
        public void SetZipFileName(string zipFileName) { zipI.m_ZipFileName = zipFileName; }

        public string GetSourceDirectory() { return zipI.m_SourceDirectory; }
        public string GetExtractDirectory() { return zipI.m_ExtractDirectory; }
        public string GetExcludeFileMask(char separator) {
            StringBuilder sb = new StringBuilder();
            for (int i = 0; i < zipI.m_NoCompressSuffixes.Count(); i++)
            {
                if (i > 0)
                    sb.Append(separator);
                sb.Append(zipI.m_ExcludeFileMask[i]);
            }
            return sb.ToString();
        }
        public string GetIncludeFileMask(char separator) {
            StringBuilder sb = new StringBuilder();
            for (int i = 0; i < zipI.m_NoCompressSuffixes.Count(); i++)
            {
                if (i > 0)
                    sb.Append(separator);
                sb.Append(zipI.m_IncludeFileMask[i]);
            }
            return sb.ToString();
        }
        public string GetNoCompressSuffixes(char separator)
        {
            StringBuilder sb = new StringBuilder();
            for (int i = 0; i < zipI.m_NoCompressSuffixes.Count(); i++)
            {
                if(i > 0)
                    sb.Append(separator);
                sb.Append(zipI.m_NoCompressSuffixes[i]);
            }
            return sb.ToString();
        }
        public string GetPassword() { return zipI.m_Password; }
        public string GetZipFileName() { return zipI.m_ZipFileName; }

        public int GetVersion() { return zipI.m_Version; }
        public bool GetRecurseSubDirectories() { return zipI.m_RecurseSubDirectories; }
        public bool GetIncludeBaseDirectory() { return zipI.m_IncludeBaseDirectory; }
        public bool GetOverwrite() { return zipI.m_Overwrite; }
        public Int32 GetCompressionLevel()
        {
            switch (zipI.m_CompressionLevel)
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
            hResult = zipI.m_LastErrorCode;
            return zipI.m_LastErrorMsg;
        }
    }
}

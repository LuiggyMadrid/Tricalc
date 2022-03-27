using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO.Compression;
using System.IO;
using System.Runtime.InteropServices;
using System.Windows;


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
        private string _excludeFileMask = null;
        private string _includeFileMask = null;
        private string _noCompressSuffixes = null;
        private string _password = null;
        private string _zipFileName = null;
        private string _lastErrorMsg = null;

        private int _version = 0;
        private int _lastErrorCode = 0;
        private bool _recurseSubDirectories = false;
        private bool _includeBaseDirectory = false;
        private CompressionLevel _compressionLevel = CompressionLevel.Fastest;


        //Accesores públicos
        public int m_Version { get => _version; protected set => _version = value; }
        public int m_LastErrorCode { get => _lastErrorCode; protected set => _lastErrorCode = value; }
        public bool m_RecurseSubDirectories { get => _recurseSubDirectories; set => _recurseSubDirectories = value; }
        public bool m_IncludeBaseDirectory { get => _includeBaseDirectory; set => _includeBaseDirectory = value; }
        public CompressionLevel m_CompressionLevel { get => _compressionLevel; set => _compressionLevel = value; }
        
        public string m_LastErrorMsg { get => _lastErrorMsg; protected set => _lastErrorMsg = value; }
        public string m_SourceDirectory { get => _sourceDirectory; set => _sourceDirectory = value; }
        public string m_ExtractDirectory { get => _extractDirectory; set => _extractDirectory = value; }
        public string m_ExcludeFileMask { get => _excludeFileMask; set => _excludeFileMask = value; }
        public string m_IncludeFileMask { get => _includeFileMask; set => _includeFileMask = value; }
        public string m_NoCompressSuffixes { get => _noCompressSuffixes; set => _noCompressSuffixes = value; }
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
                string.IsNullOrEmpty(m_ExcludeFileMask) &&
                string.IsNullOrEmpty(m_NoCompressSuffixes) &&
                (string.IsNullOrEmpty(m_IncludeFileMask) || m_IncludeFileMask == "*"))
            {
                //Compress Directory
                try
                {
                    ZipFile.CreateFromDirectory(m_SourceDirectory, m_ZipFileName, m_CompressionLevel, false); // includeBaseDirectory);
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
    }

    [Guid("82508DF9-C406-4B57-808B-C23B41B9A633")]
    public interface INetZipInterface
    {
        bool Create();
        void SetRecurseSubDirectories(bool recurseSubDirectories);
        void SetIncludeBaseDirectory(bool includeBaseDirectory);
        void SetCompressionLevel(Int32 compressionLevel);

        void SetSourceDirectory(string sourceDir);
        void SetExtractDirectory(string extractDir);
        void SetExcludeFileMask(string excludeFileMask);
        void SetIncludeFileMask(string includeFileMask);
        void SetNoCompressSuffixes(string noCompressSuffixes);
        void SetPassword(string password);
        void SetZipFileName(string zipFileName);


        int GetVersion();
        bool GetRecurseSubDirectories();
        bool GetIncludeBaseDirectory();
        Int32 GetCompressionLevel();

        string GetSourceDirectory();
        string GetExtractDirectory();
        string GetExcludeFileMask();
        string GetIncludeFileMask();
        string GetNoCompressSuffixes();
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
        public void SetCompressionLevel(Int32 compressionLevel) {
            if (compressionLevel == 0) zipI.m_CompressionLevel = CompressionLevel.NoCompression;
            else if (compressionLevel <= 5) zipI.m_CompressionLevel = CompressionLevel.Fastest;
            else zipI.m_CompressionLevel = CompressionLevel.Optimal;
        }

        public void SetSourceDirectory(string soureDir) { zipI.m_SourceDirectory = soureDir; }
        public void SetExtractDirectory(string extractDir) { zipI.m_ExtractDirectory = extractDir; }
        public void SetExcludeFileMask(string excludeFileMask) { zipI.m_ExcludeFileMask = excludeFileMask; }
        public void SetIncludeFileMask(string includeFileMask) { zipI.m_IncludeFileMask = includeFileMask; }
        public void SetNoCompressSuffixes(string noCompressSuffixes) { zipI.m_NoCompressSuffixes = noCompressSuffixes; }
        public void SetPassword(string password) { zipI.m_Password = password; }
        public void SetZipFileName(string zipFileName) { zipI.m_ZipFileName = zipFileName; }

        public string GetSourceDirectory() { return zipI.m_SourceDirectory; }
        public string GetExtractDirectory() { return zipI.m_ExtractDirectory; }
        public string GetExcludeFileMask() { return zipI.m_ExcludeFileMask; }
        public string GetIncludeFileMask() { return zipI.m_IncludeFileMask; }
        public string GetNoCompressSuffixes() { return zipI.m_NoCompressSuffixes; }
        public string GetPassword() { return zipI.m_Password; }
        public string GetZipFileName() { return zipI.m_ZipFileName; }

        public int GetVersion() { return zipI.m_Version; }
        public bool GetRecurseSubDirectories() { return zipI.m_RecurseSubDirectories; }
        public bool GetIncludeBaseDirectory() { return zipI.m_IncludeBaseDirectory; }
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
        public string GetLastZipError(ref Int32 hResult)
        {
            hResult = zipI.m_LastErrorCode;
            return zipI.m_LastErrorMsg;
        }
    }
}

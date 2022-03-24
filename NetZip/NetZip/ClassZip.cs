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
        private string m_sourceDirectory = null;
        private string m_extractDirectory = null;
        private string m_excludeFileMask = null;
        private string m_includeFileMask = null;
        private string m_noCompressSuffixes = null;
        private string m_password = null;
        private string m_zipFileName = null;

        private int m_version = 0;
        
        
        //Accesores pñublicos
        public int Version { get => m_version; set => m_version = value; }
        public string SourceDirectory { get => m_sourceDirectory; set => m_sourceDirectory = value; }
        public string ExtractDirectory { get => m_extractDirectory; set => m_extractDirectory = value; }
        public string ExcludeFileMask { get => m_excludeFileMask; set => m_excludeFileMask = value; }
        public string IncludeFileMask { get => m_includeFileMask; set => m_includeFileMask = value; }
        public string NoCompressSuffixes { get => m_noCompressSuffixes; set => m_noCompressSuffixes = value; }
        public string Password { get => m_password; set => m_password = value; }
        public string ZipFileName { get => m_zipFileName; set => m_zipFileName = value; }

        //----------------------------------------------------------------------
        /// <summary>
        /// Add files to an existing .zip
        /// </summary>
        /// <returns>0 on succesfull; error code otherwise</returns>
        public long Add()
        {
            return 0;
        }
    }

    [Guid("82508DF9-C406-4B57-808B-C23B41B9A633")]
    public interface INetZipInterface
    {
        bool Create();
        void SetVersion(int version);
        void SetSourceDirectory(string sourceDir);
        void SetExtractDirectory(string extractDir);
        void SetExcludeFileMask(string excludeFileMask);
        void SetIncludeFileMask(string includeFileMask);
        void SetNoCompressSuffixes(string noCompressSuffixes);
        void SetPassword(string password);
        void SetZipFileName(string zipFileName);


        int GetVersion();
        string GetSourceDirectory();
        string GetExtractDirectory();
        string GetExcludeFileMask();
        string GetIncludeFileMask();
        string GetNoCompressSuffixes();
        string GetPassword();
        string GetZipFileName();

        long Add();
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

        public void SetVersion(int version) { zipI.Version = version; }
        public void SetSourceDirectory(string soureDir) { zipI.SourceDirectory = soureDir; }
        public void SetExtractDirectory(string extractDir) { zipI.ExtractDirectory = extractDir; }
        public void SetExcludeFileMask(string excludeFileMask) { zipI.ExcludeFileMask = excludeFileMask; }
        public void SetIncludeFileMask(string includeFileMask) { zipI.IncludeFileMask = includeFileMask; }
        public void SetNoCompressSuffixes(string noCompressSuffixes) { zipI.NoCompressSuffixes = noCompressSuffixes; }
        public void SetPassword(string password) { zipI.Password = password; }
        public void SetZipFileName(string zipFileName) { zipI.ZipFileName = zipFileName; }

        public string GetSourceDirectory() { return zipI.SourceDirectory; }
        public string GetExtractDirectory() { return zipI.ExtractDirectory; }
        public string GetExcludeFileMask() { return zipI.ExcludeFileMask; }
        public string GetIncludeFileMask() { return zipI.IncludeFileMask; }
        public string GetNoCompressSuffixes() { return zipI.NoCompressSuffixes; }
        public string GetPassword() { return zipI.Password; }
        public string GetZipFileName() { return zipI.ZipFileName; }

        public int GetVersion() { return zipI.Version; }

        public long Add() { return zipI.Add(); }
    }
}

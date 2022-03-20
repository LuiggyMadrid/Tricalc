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
        private int m_version = 0;
        
        //Accesores pñublicos
        public int Version { get => m_version; set => m_version = value; }
        public string SourceDirectory { get => m_sourceDirectory; set => m_sourceDirectory = value; }
        public string ExtractDirectory { get => m_extractDirectory; set => m_extractDirectory = value; }
    }

    [Guid("82508DF9-C406-4B57-808B-C23B41B9A633")]
    public interface INetZipInterface
    {
        bool Create();
        void SetVersion(int version);
        void SetSourceDirectory(string sourceDir);
        void SetExtractDirectory(string extractDir);


        int GetVersion();
        string GetSourceDirectory();
        string GetExtractDirectory();

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

        public string GetSourceDirectory() { return zipI.SourceDirectory; }
        public string GetExtractDirectory() { return zipI.ExtractDirectory; }
        public int GetVersion() { return zipI.Version; }
    }
}

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



    }

    [Guid("82508DF9-C406-4B57-808B-C23B41B9A633")]
    public interface INetZipInterface
    {
        bool Create(string lpszWindowName, UInt32 dwStyle,
            //const RECT& rect,
             //CWnd* pParentWnd,
            UInt32 nID,
            //CFile* pPersist = NULL,
            bool bStorage = false);


    }
    [Guid("5E2A5A32-FA8D-4941-8E95-F4B8F5BB9CCE")]
    public class InterfaceImplementation : INetZipInterface
    {
        private ClassZip zipI = null;

    }
}

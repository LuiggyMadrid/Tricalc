#pragma once
#include <afxwin.h>

//#import <mscorlib.tlb> raw_interfaces_only
#ifdef _DEBUG
#import "..\..\NetZip\NetZip\bin\Debug\NetZip.tlb" no_namespace named_guids
#else
#import "..\..\NetZip\NetZip\bin\Release\NetZip.tlb" no_namespace named_guids
#endif

//Clona el interface de la clase CDPolarZip, añadiendo un miembro de la clase calassZip
class CDPolarZipWarper :
    public CWnd
{
protected:
	DECLARE_DYNCREATE(CDPolarZipWarper)
public:
	INetZipInterface* cppIzip;

	CLSID const& GetClsid()
	{
		static CLSID const clsid
			= { 0x9d6efb87, 0x5e95, 0x11d2, { 0xbb, 0xc, 0x44, 0x45, 0x53, 0x54, 0x0, 0x0 } };
		return clsid;
	}
	virtual BOOL Create(LPCTSTR lpszClassName,
		LPCTSTR lpszWindowName, DWORD dwStyle,
		const RECT& rect,
		CWnd* pParentWnd, UINT nID,
		CCreateContext* pContext = NULL)
	{
		return CreateControl(GetClsid(), lpszWindowName, dwStyle, rect, pParentWnd, nID);
	}

	BOOL Create(CWnd* pParentWnd)
	{
		//if (!cppIzip)
		//	cppIzip = new INetZipInterface;
		if (!cppIzip || !cppIzip->Create())
			return FALSE;
		if (!CreateControl(GetClsid(), _T(""), WS_TABSTOP, CRect(0, 0, 0, 0), pParentWnd, 0))
			return FALSE;
		return TRUE;
	}

	// Attributes
public:
	BOOL GetRestoreVolumeLabel();
	void SetRestoreVolumeLabel(BOOL);
	BOOL GetMultiDiskEraseDisks();
	void SetMultiDiskEraseDisks(BOOL);
	BOOL GetConvertSpacesToUnderscores();
	void SetConvertSpacesToUnderscores(BOOL);
	CString GetSkipFilesBeforeDate();
	void SetSkipFilesBeforeDate(LPCTSTR);
	CString GetSkipFilesAfterDate();
	void SetSkipFilesAfterDate(LPCTSTR);
	BOOL GetConvertLFtoCRLF();
	void SetConvertLFtoCRLF(BOOL);
	BOOL GetConvertCRLFtoLF();
	void SetConvertCRLFtoLF(BOOL);
	BOOL GetOverwrite();
	void SetOverwrite(BOOL);
	BOOL GetIncludeDirectoryEntries();
	void SetIncludeDirectoryEntries(BOOL);
	BOOL GetIncludeHiddenFiles();
	void SetIncludeHiddenFiles(BOOL);
	BOOL GetIncludeVolumeLabel();
	void SetIncludeVolumeLabel(BOOL);
	BOOL GetRecurseSubDirectories();
	void SetRecurseSubDirectories(BOOL);
	BOOL GetStorePaths();
	void SetStorePaths(BOOL);
	BOOL GetUsePassword();
	void SetUsePassword(BOOL);
	long GetCompressionLevel();
	void SetCompressionLevel(long);
	long GetAction();
	void SetAction(long);
	CString GetExcludeFileMask();
	void SetExcludeFileMask(LPCTSTR);
	CString GetExtractDirectory();
	void SetExtractDirectory(LPCTSTR);
	CString GetIncludeFileMask();
	void SetIncludeFileMask(LPCTSTR);
	CString GetNoCompressSuffixes();
	void SetNoCompressSuffixes(LPCTSTR);
	CString GetPassword();
	void SetPassword(LPCTSTR);
	CString GetSourceDirectory();
	void SetSourceDirectory(LPCTSTR);
	CString GetTemporaryPath();
	void SetTemporaryPath(LPCTSTR);
	CString GetZipFileName();
	void SetZipFileName(LPCTSTR);
	BOOL GetDelphi();
	void SetDelphi(BOOL);
	BOOL GetMultiDiskCreate();
	void SetMultiDiskCreate(BOOL);
	BOOL GetMultiDiskUseDefaultDialogs();
	void SetMultiDiskUseDefaultDialogs(BOOL);
	CString GetMultiDiskDialogCaption();
	void SetMultiDiskDialogCaption(LPCTSTR);
	long GetMultiDiskMaxVolumeSize();
	void SetMultiDiskMaxVolumeSize(long);
	long GetMultiDiskFreeOnFirstDisk();
	void SetMultiDiskFreeOnFirstDisk(long);
	long GetMultiDiskMinimumFreeOnDisk();
	void SetMultiDiskMinimumFreeOnDisk(long);
	CString GetPolarZipSpanDllDirectory();
	void SetPolarZipSpanDllDirectory(LPCTSTR);

	// Operations
public:
	CString BrowseForFolder(long hWindowHandle, LPCTSTR strTitle);
	long RemoveSFXStubFromZip(LPCTSTR strZipFileName);
	long CreateSFXFromZip(LPCTSTR strZipFileName, LPCTSTR strStubFileName);
	long SetArchiveComment(LPCTSTR strZipFileName, LPCTSTR strComment);
	CString GetArchiveComment(LPCTSTR strZipFileName);
	long CreateZipFromMultiFileArchive(LPCTSTR strInFile, LPCTSTR strOutFile);
	long TestArchive(LPCTSTR strZipFileName);
	long CreateMultiDiskArchiveFromZip(LPCTSTR strInFile, LPCTSTR strOutFile);
	long TimeStampArchive(LPCTSTR strZipFileName);
	long Add();
	CString GetUnzipErrorDescription(long nError);
	CString GetZipErrorDescription(long nError);
	long GetLastZipError();
	long GetLastUnzipError();
	long Extract();
	long FixZipFile(LPCTSTR strZipFileName, BOOL bHarder);
	void AboutBox();
};



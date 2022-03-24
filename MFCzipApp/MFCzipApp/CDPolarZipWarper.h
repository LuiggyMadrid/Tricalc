#pragma once
#include <afxwin.h>

//#import <mscorlib.tlb> raw_interfaces_only
#ifdef _DEBUG
#import "..\..\NetZip\NetZip\bin\Debug\NetZip.tlb" no_namespace named_guids
#else
#import "..\..\NetZip\NetZip\bin\Release\NetZip.tlb" no_namespace named_guids
#endif

//Clona el interface de la clase CDPolarZip, añadiendo un miembro de la clase calassZip
class CDPolarZipWarper
{
public:
	INetZipInterface* cppIzip = nullptr;
	CWnd* pParentWnd = nullptr;

	CDPolarZipWarper() {}
	~CDPolarZipWarper() {
		cppIzip->Release();
	}

	BOOL Create(CWnd* _pParentWnd)
	{
		BOOL ret = TRUE;
		HRESULT hr = S_OK;

		pParentWnd = _pParentWnd;
		hr = CoCreateInstance(CLSID_InterfaceImplementation,
			NULL,
			CLSCTX_INPROC_SERVER,
			IID_INetZipInterface,
			reinterpret_cast<void**>(&cppIzip));

		if (FAILED(hr))
		{
			printf("Couldn't create the instance!... 0x%x\n", hr);
			ret = FALSE;
		}
		if (!cppIzip || !cppIzip->Create())
			ret = FALSE;
		return ret;
	}

	// Attributes
public:
	//BOOL GetRestoreVolumeLabel();
	//void SetRestoreVolumeLabel(BOOL);
	//BOOL GetMultiDiskEraseDisks();
	//void SetMultiDiskEraseDisks(BOOL);
	//BOOL GetConvertSpacesToUnderscores();
	//void SetConvertSpacesToUnderscores(BOOL);
	CString GetSkipFilesBeforeDate() { ASSERT(false); return CString(""); }
	void SetSkipFilesBeforeDate(LPCTSTR) { ASSERT(FALSE); }
	CString GetSkipFilesAfterDate() { ASSERT(false); return CString(""); }
	void SetSkipFilesAfterDate(LPCTSTR) { ASSERT(FALSE); }
	//BOOL GetConvertLFtoCRLF();
	//void SetConvertLFtoCRLF(BOOL);
	//BOOL GetConvertCRLFtoLF();
	//void SetConvertCRLFtoLF(BOOL);
	BOOL GetOverwrite();
	void SetOverwrite(BOOL);
	BOOL GetIncludeDirectoryEntries();
	void SetIncludeDirectoryEntries(BOOL);
	BOOL GetIncludeHiddenFiles();
	void SetIncludeHiddenFiles(BOOL);
	//BOOL GetIncludeVolumeLabel();
	//void SetIncludeVolumeLabel(BOOL);
	BOOL GetRecurseSubDirectories();
	void SetRecurseSubDirectories(BOOL);
	BOOL GetStorePaths();
	void SetStorePaths(BOOL);
	BOOL GetUsePassword();
	void SetUsePassword(BOOL);
	long GetCompressionLevel();
	void SetCompressionLevel(long);
	//long GetAction();
	//void SetAction(long);
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



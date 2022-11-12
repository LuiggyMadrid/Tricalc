#pragma once
#include <afxwin.h>

#ifdef _DEBUG
#import "..\..\NetZip\NetZip\bin\Debug\NetZip.tlb" no_namespace named_guids
#else
#import "..\..\NetZip\NetZip\bin\Release\NetZip.tlb" no_namespace named_guids
#endif


/// <summary>
/// Clones class CDPolarZip public members, addin a classZip member
/// (Clona el interface de la clase CDPolarZip, añadiendo un miembro de la clase classZip)
/// </summary>
class CDPolarZipWarper
{
public:
	INetZipInterface* cppIzip = nullptr;
	CWnd* pParentWnd = nullptr;

	CDPolarZipWarper(CWnd* _pParentWnd = nullptr);
	~CDPolarZipWarper();

	// Attributes
public:
	CString GetSkipFilesBeforeDate() { ASSERT(false); return CString(""); } // Not implemented
	void SetSkipFilesBeforeDate(LPCTSTR) { ASSERT(FALSE); } // Not implemented
	CString GetSkipFilesAfterDate() { ASSERT(false); return CString(""); } // Not implemented
	void SetSkipFilesAfterDate(LPCTSTR) { ASSERT(FALSE); } // Not implemented
	
	BOOL GetOverwrite();
	void SetOverwrite(BOOL);
	BOOL GetIncludeDirectoryEntries();
	void SetIncludeDirectoryEntries(BOOL);
	BOOL GetIncludeBaseDirectory();
	void SetIncludeBaseDirectory(BOOL);
	BOOL GetIncludeHiddenFiles();
	void SetIncludeHiddenFiles(BOOL);
	BOOL GetRecurseSubDirectories();
	void SetRecurseSubDirectories(BOOL);
	BOOL GetStorePaths();
	void SetStorePaths(BOOL);
	BOOL GetUsePassword();
	void SetUsePassword(BOOL);
	long GetCompressionLevel();
	void SetCompressionLevel(long);
	CString GetExcludeFileMask(TCHAR chSeparator = _T(' '));
	void SetExcludeFileMask(LPCTSTR pszExcludeFileMask, TCHAR chSeparator = _T(' '));
	CString GetExtractDirectory();
	void SetExtractDirectory(LPCTSTR);
	CString GetIncludeFileMask(TCHAR chSeparator = _T(' '));
	void SetIncludeFileMask(LPCTSTR, TCHAR chSeparator=_T(' '));
	CString GetNoCompressSuffixes(TCHAR chSeparator = _T(' '));
	void SetNoCompressSuffixes(LPCTSTR pszNoCompressSuffixes, TCHAR chSeparator = _T(' '));
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
	CString BrowseForFolder(long hWindowHandle, LPCTSTR strTitle); // Not implemented
	long RemoveSFXStubFromZip(LPCTSTR strZipFileName); // Not implemented
	long CreateSFXFromZip(LPCTSTR strZipFileName, LPCTSTR strStubFileName); // Not implemented
	long SetArchiveComment(LPCTSTR strZipFileName, LPCTSTR strComment); // Not implemented
	CString GetArchiveComment(LPCTSTR strZipFileName); // Not implemented
	long CreateZipFromMultiFileArchive(LPCTSTR strInFile, LPCTSTR strOutFile); // Not implemented
	long TestArchive(LPCTSTR strZipFileName); // Not implemented
	//long CreateMultiDiskArchiveFromZip(LPCTSTR strInFile, LPCTSTR strOutFile); // Not implemented
	long TimeStampArchive(LPCTSTR strZipFileName); // Not implemented
	
	/// <summary>
	/// Compress in ZIP format
	/// </summary>
	/// <param name="create">TRUE for creating a new ZIP file; false to append in an existing ZIP file</param>
	/// <returns>Error code (S_OK for successful)</returns>
	HRESULT Add(BOOL create = TRUE);
	
	/// <summary>
	/// Get Last error message and code
	/// </summary>
	/// <param name="errorCode">OUT: Error code</param>
	/// <returns>Error message</returns>
	CString GetLastZipError(HRESULT &errorCode);

	/// <summary>
	/// Extract from an existing zip file
	/// </summary>
	/// <param name="create">TRUE for creating a new ZIP file; false to append in an existing ZIP file</param>
	/// <returns>Error code (S_OK for successful)</returns>
	HRESULT Extract();

	long FixZipFile(LPCTSTR strZipFileName, BOOL bHarder); // Not implemented
};



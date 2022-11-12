#include "pch.h"
#include "CDPolarZipWarper.h"

CDPolarZipWarper::CDPolarZipWarper(CWnd* _pParentWnd)
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
	else if (!cppIzip || !cppIzip->Create())
		ret = FALSE;
	if (!ret)
	{
		if (cppIzip)
			cppIzip->Release();
		cppIzip = nullptr;
	}
}

CDPolarZipWarper::~CDPolarZipWarper()
{
	if (cppIzip)
		cppIzip->Release();
}

void CDPolarZipWarper::SetSourceDirectory(LPCTSTR pszSourceDir)
{
	if (this->cppIzip)
		this->cppIzip->SetSourceDirectory(pszSourceDir);
}

CString CDPolarZipWarper::GetSourceDirectory()
{
	if (this->cppIzip)
		return CString(this->cppIzip->GetSourceDirectory().GetBSTR());
	else
		return CString("");
}

void CDPolarZipWarper::SetExtractDirectory(LPCTSTR pszExtractDir)
{
	if (this->cppIzip)
		this->cppIzip->SetExtractDirectory(pszExtractDir);
}

CString CDPolarZipWarper::GetExtractDirectory()
{
	if (this->cppIzip)
		return CString(this->cppIzip->GetExtractDirectory().GetBSTR());
	else
		return CString("");
}

CString CDPolarZipWarper::GetZipFileName()
{
	if (this->cppIzip)
		return CString(this->cppIzip->GetZipFileName().GetBSTR());
	else
		return CString("");
}

void CDPolarZipWarper::SetZipFileName(LPCTSTR pszZipName)
{
	if (this->cppIzip)
		this->cppIzip->SetZipFileName(pszZipName);
}

BOOL CDPolarZipWarper::GetIncludeDirectoryEntries()
{
	//TO DO
	_ASSERTE(false);
	return TRUE;
}

void CDPolarZipWarper::SetIncludeDirectoryEntries(BOOL val)
{
	//TO DO
	//TRUE val is always asumed
	_ASSERTE(false);
}

BOOL CDPolarZipWarper::GetIncludeBaseDirectory()
{
	if (cppIzip)
		return cppIzip->GetIncludeBaseDirectory();
	else
		return FALSE;
}

void CDPolarZipWarper::SetIncludeBaseDirectory(BOOL val)
{
	if (cppIzip)
		cppIzip->SetIncludeBaseDirectory(!!val);

}

BOOL CDPolarZipWarper::GetIncludeHiddenFiles()
{
	//TO DO
	_ASSERTE(false);
	return TRUE;
}

void CDPolarZipWarper::SetIncludeHiddenFiles(BOOL val)
{
	//TO DO
	//TRUE val is akkways asumed
	_ASSERTE(false);
}

BOOL CDPolarZipWarper::GetRecurseSubDirectories()
{
	if (this->cppIzip)
		return this->cppIzip->GetRecurseSubDirectories();
	else
		return FALSE;
}

void CDPolarZipWarper::SetRecurseSubDirectories(BOOL val)
{
	if (this->cppIzip)
		this->cppIzip->SetRecurseSubDirectories(!!val);
}

BOOL CDPolarZipWarper::GetOverwrite()
{
	if (this->cppIzip)
		return this->cppIzip->GetOverwrite();
	else
		return FALSE;
}

void CDPolarZipWarper::SetOverwrite(BOOL val)
{
	if (this->cppIzip)
		this->cppIzip->SetOverwrite(!!val);
}

BOOL CDPolarZipWarper::GetStorePaths()
{
	//TO DO
	_ASSERTE(false);
	return TRUE;
}

void CDPolarZipWarper::SetStorePaths(BOOL val)
{
	//TO DO
	_ASSERTE(false);
}

BOOL CDPolarZipWarper::GetUsePassword()
{
	//TO DO
	_ASSERTE(false);
	return FALSE;
}

void CDPolarZipWarper::SetUsePassword(BOOL val)
{
	//TO DO
	//FALSE val is allways asumed
	_ASSERTE(false);
}

long CDPolarZipWarper::GetCompressionLevel()
{
	if (this->cppIzip)
		return this->cppIzip->GetCompressionLevel();
	else
		return 0;
}

void CDPolarZipWarper::SetCompressionLevel(long val)
{
	if (this->cppIzip)
		this->cppIzip->SetCompressionLevel(val);
}

CString CDPolarZipWarper::GetExcludeFileMask(TCHAR chSeparator)
{
	if (this->cppIzip)
		return CString(this->cppIzip->GetExcludeFileMask(chSeparator).GetBSTR());
	else
		return CString("");
}

void CDPolarZipWarper::SetExcludeFileMask(LPCTSTR pszExcludeFileMask, TCHAR chSeparator)
{
	if (this->cppIzip)
		this->cppIzip->SetExcludeFileMask(pszExcludeFileMask, chSeparator);
}



CString CDPolarZipWarper::GetIncludeFileMask(TCHAR chSeparator)
{
	if (this->cppIzip)
		return CString(this->cppIzip->GetIncludeFileMask(chSeparator).GetBSTR());
	else
		return CString("");
}

void CDPolarZipWarper::SetIncludeFileMask(LPCTSTR pszIncludeFileMask, TCHAR chSeparator)
{
	if (this->cppIzip)
		this->cppIzip->SetIncludeFileMask(pszIncludeFileMask, chSeparator);
}


CString CDPolarZipWarper::GetNoCompressSuffixes(TCHAR chSeparator)
{
	if (this->cppIzip)
		return CString(this->cppIzip->GetNoCompressSuffixes(chSeparator).GetBSTR());
	else
		return CString("");
}

void CDPolarZipWarper::SetNoCompressSuffixes(LPCTSTR pszNoCompressSuffixes, TCHAR chSeparator)
{
	if (this->cppIzip)
		this->cppIzip->SetNoCompressSuffixes(pszNoCompressSuffixes, chSeparator);
}


CString CDPolarZipWarper::GetPassword()
{
	if (this->cppIzip)
		return CString(this->cppIzip->GetPassword().GetBSTR());
	else
		return CString("");
}

void CDPolarZipWarper::SetPassword(LPCTSTR pszPassword)
{
	if (this->cppIzip)
		this->cppIzip->SetPassword(pszPassword);
}

/// <summary>
/// Compress in ZIP format
/// </summary>
/// <param name="create">TRUE for creating a new ZIP file; false to append in an existing ZIP file</param>
/// <returns>Error code (S_OK for successful)</returns>
HRESULT CDPolarZipWarper::Add(BOOL create)
{
	if (this->cppIzip) //Siempre comprobar
	{
		return this->cppIzip->Add(!!create);
	}
	else
		return S_FALSE;
}

/// <summary>
/// Extracts from an existing Zip file
/// </summary>
/// <returns>Error code (S_OK for successful)</returns>
HRESULT CDPolarZipWarper::Extract()
{
	if (this->cppIzip)
	{
		return this->cppIzip->Extract();
	}
	else
		return S_FALSE;
}

/// <summary>
/// Get Last error message and code
/// </summary>
/// <param name="errorCode">OUT: Error code</param>
/// <returns>Error message</returns>
CString CDPolarZipWarper::GetLastZipError(HRESULT& errorCode)
{
	errorCode = S_FALSE;
	if (this->cppIzip)
	{
		return CString(cppIzip->GetLastZipError(& errorCode).GetBSTR());
	}
	else
	{
		errorCode = E_FAIL;
		return CString("Error desconocido");
	}
}

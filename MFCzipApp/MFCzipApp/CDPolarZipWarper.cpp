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
	if (!cppIzip || !cppIzip->Create())
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


BOOL CDPolarZipWarper::GetOverwrite()
{
	//TO DO
	return TRUE;
}

void CDPolarZipWarper::SetOverwrite(BOOL val)
{
	//TO DO
}

BOOL CDPolarZipWarper::GetIncludeDirectoryEntries()
{
	//TO DO
	return TRUE;
}

void CDPolarZipWarper::SetIncludeDirectoryEntries(BOOL val)
{
	//TO DO
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
	return TRUE;
}

void CDPolarZipWarper::SetIncludeHiddenFiles(BOOL val)
{
	//TO DO
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

BOOL CDPolarZipWarper::GetStorePaths()
{
	//TO DO
	return TRUE;
}

void CDPolarZipWarper::SetStorePaths(BOOL val)
{
	//TO DO
}

BOOL CDPolarZipWarper::GetUsePassword()
{
	//TO DO
	return TRUE;
}

void CDPolarZipWarper::SetUsePassword(BOOL val)
{
	//TO DO
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

CString CDPolarZipWarper::GetExcludeFileMask()
{
	if (this->cppIzip)
		return CString(this->cppIzip->GetExcludeFileMask().GetBSTR());
	else
		return CString("");
}

void CDPolarZipWarper::SetExcludeFileMask(LPCTSTR pszExcludeFileMask)
{
	if (this->cppIzip)
		this->cppIzip->SetExcludeFileMask(pszExcludeFileMask);
}



CString CDPolarZipWarper::GetIncludeFileMask()
{
	if (this->cppIzip)
		return CString(this->cppIzip->GetIncludeFileMask().GetBSTR());
	else
		return CString("");
}

void CDPolarZipWarper::SetIncludeFileMask(LPCTSTR pszIncludeFileMask)
{
	if (this->cppIzip)
		this->cppIzip->SetIncludeFileMask(pszIncludeFileMask);
}


CString CDPolarZipWarper::GetNoCompressSuffixes()
{
	if (this->cppIzip)
		return CString(this->cppIzip->GetNoCompressSuffixes().GetBSTR());
	else
		return CString("");
}

void CDPolarZipWarper::SetNoCompressSuffixes(LPCTSTR pszNoCompressSuffixes)
{
	if (this->cppIzip)
		this->cppIzip->SetNoCompressSuffixes(pszNoCompressSuffixes);
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

//Conflicto

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

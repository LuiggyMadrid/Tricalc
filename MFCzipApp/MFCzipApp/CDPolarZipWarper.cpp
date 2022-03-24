#include "pch.h"
#include "CDPolarZipWarper.h"

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
	//TO DO
	return TRUE;
}

void CDPolarZipWarper::SetRecurseSubDirectories(BOOL val)
{
	//TO DO
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
	//TO DO
	return 9;
}

void CDPolarZipWarper::SetCompressionLevel(long val)
{
	//TO DO
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

long CDPolarZipWarper::Add()
{
	if (this->cppIzip)
	{
		CString s1 = this->cppIzip->GetSourceDirectory();
		CString s2 = this->cppIzip->GetZipFileName();
		return this->cppIzip->Add();
	}
	else
		return -1;
}

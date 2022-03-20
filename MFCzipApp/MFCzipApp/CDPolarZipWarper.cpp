#include "pch.h"
#include "CDPolarZipWarper.h"


IMPLEMENT_DYNCREATE(CDPolarZipWarper, CWnd)

void CDPolarZipWarper::SetSourceDirectory(LPCTSTR pszSourceDir)
{
	if (this->cppIzip)
		this->cppIzip->SetSourceDirectory(pszSourceDir);
}

void CDPolarZipWarper::SetExtractDirectory(LPCTSTR pszExtractDir)
{
	if (this->cppIzip)
		this->cppIzip->SetExtractDirectory(pszExtractDir);
}


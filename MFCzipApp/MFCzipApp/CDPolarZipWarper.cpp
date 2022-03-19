#include "pch.h"
#include "CDPolarZipWarper.h"


IMPLEMENT_DYNCREATE(CDPolarZipWarper, CWnd)

void CDPolarZipWarper::SetSourceDirectory(LPCTSTR pszSourceDir)
{
	this->cppIzip->SetSourceDirectory(pszSourceDir);
}



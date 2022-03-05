
// MFCzipApp.h: archivo de encabezado principal para la aplicación MFCzipApp
//
#pragma once

#ifndef __AFXWIN_H__
	#error "incluir 'pch.h' antes de incluir este archivo para PCH"
#endif

#include "resource.h"       // Símbolos principales


// CMFCzipAppApp:
// Consulte MFCzipApp.cpp para obtener información sobre la implementación de esta clase
//

class CMFCzipAppApp : public CWinApp
{
public:
	CMFCzipAppApp() noexcept;


// Reemplazos
public:
	virtual BOOL InitInstance();
	virtual int ExitInstance();

// Implementación

public:
	afx_msg void OnAppAbout();
	DECLARE_MESSAGE_MAP()
};

extern CMFCzipAppApp theApp;

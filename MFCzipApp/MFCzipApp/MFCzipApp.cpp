
// MFCzipApp.cpp: define los comportamientos de las clases para la aplicación.
//

#include "pch.h"
#include "framework.h"
#include "afxwinappex.h"
#include "afxdialogex.h"
#include "MFCzipApp.h"
#include "MainFrm.h"
#include "CDPolarZipWarper.h"


#ifdef _DEBUG
#define new DEBUG_NEW
#endif


// CMFCzipAppApp

BEGIN_MESSAGE_MAP(CMFCzipAppApp, CWinApp)
	ON_COMMAND(ID_APP_ABOUT, &CMFCzipAppApp::OnAppAbout)
	ON_COMMAND(ID_ARCHIVO_TESTZIP,&CMFCzipAppApp::OnAppTestZip)
END_MESSAGE_MAP()


// Construcción de CMFCzipAppApp

CMFCzipAppApp::CMFCzipAppApp() noexcept
{

	// admite Administrador de reinicio
	m_dwRestartManagerSupportFlags = AFX_RESTART_MANAGER_SUPPORT_ALL_ASPECTS;
#ifdef _MANAGED
	// Si la aplicación se compila para ser compatible con Common Language Runtime (/clr):
	//     1) Esta configuración adicional es necesaria para que la compatibilidad con el Administrador de reinicio funcione correctamente.
	//     2) En el proyecto, debe agregar una referencia a System.Windows.Forms para poder compilarlo.
	System::Windows::Forms::Application::SetUnhandledExceptionMode(System::Windows::Forms::UnhandledExceptionMode::ThrowException);
#endif

	// TODO: reemplace la cadena de identificador de aplicación siguiente por una cadena de identificador único; el formato
	// recomendado para la cadena es NombreCompañía.NombreProducto.Subproducto.InformaciónDeVersión
	SetAppID(_T("MFCzipApp.AppID.NoVersion"));

	// TODO: agregar aquí el código de construcción,
	// Colocar toda la inicialización importante en InitInstance
}

// Único objeto CMFCzipAppApp

CMFCzipAppApp theApp;


// Inicialización de CMFCzipAppApp

BOOL CMFCzipAppApp::InitInstance()
{
	// Windows XP requiere InitCommonControlsEx() si un manifiesto de
	// aplicación especifica el uso de ComCtl32.dll versión 6 o posterior para habilitar
	// estilos visuales.  De lo contrario, se generará un error al crear ventanas.
	INITCOMMONCONTROLSEX InitCtrls;
	InitCtrls.dwSize = sizeof(InitCtrls);
	// Establecer para incluir todas las clases de control comunes que desee utilizar
	// en la aplicación.
	InitCtrls.dwICC = ICC_WIN95_CLASSES;
	InitCommonControlsEx(&InitCtrls);

	CWinApp::InitInstance();


	// Inicializar bibliotecas OLE
	if (!AfxOleInit())
	{
		AfxMessageBox(IDP_OLE_INIT_FAILED);
		return FALSE;
	}

	AfxEnableControlContainer();

	EnableTaskbarInteraction(FALSE);

	// Se necesita AfxInitRichEdit2() para usar el control RichEdit
	// AfxInitRichEdit2();

	// Inicialización estándar
	// Si no utiliza estas funcionalidades y desea reducir el tamaño
	// del archivo ejecutable final, debe quitar
	// las rutinas de inicialización específicas que no necesite
	// Cambie la clave del Registro en la que se almacena la configuración
	// TODO: debe modificar esta cadena para que contenga información correcta
	// como el nombre de su compañía u organización
	SetRegistryKey(_T("Aplicaciones generadas con el Asistente para aplicaciones local"));


	// Para crear la ventana principal, este código crea un nuevo objeto de ventana
	// de marco y, a continuación, lo establece como el objeto de ventana principal de la aplicación
	CFrameWnd* pFrame = new CMainFrame;
	if (!pFrame)
		return FALSE;
	m_pMainWnd = pFrame;
	// Crear y cargar el marco con sus recursos
	pFrame->LoadFrame(IDR_MAINFRAME,
		WS_OVERLAPPEDWINDOW | FWS_ADDTOTITLE, nullptr,
		nullptr);





	// Se ha inicializado la única ventana; mostrarla y actualizarla
	pFrame->ShowWindow(SW_SHOW);
	pFrame->UpdateWindow();
	return TRUE;
}

int CMFCzipAppApp::ExitInstance()
{
	//TODO: controlar recursos adicionales que se hayan podido agregar
	AfxOleTerm(FALSE);

	return CWinApp::ExitInstance();
}

// Controladores de mensajes de CMFCzipAppApp


// Cuadro de diálogo CAboutDlg utilizado para el comando Acerca de

class CAboutDlg : public CDialogEx
{
public:
	CAboutDlg() noexcept;

// Datos del cuadro de diálogo
#ifdef AFX_DESIGN_TIME
	enum { IDD = IDD_ABOUTBOX };
#endif

protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // Compatibilidad con DDX/DDV

// Implementación
protected:
	DECLARE_MESSAGE_MAP()
};

CAboutDlg::CAboutDlg() noexcept : CDialogEx(IDD_ABOUTBOX)
{
}

void CAboutDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialogEx::DoDataExchange(pDX);
}

BEGIN_MESSAGE_MAP(CAboutDlg, CDialogEx)
END_MESSAGE_MAP()

// Comando de la aplicación para ejecutar el cuadro de diálogo
void CMFCzipAppApp::OnAppAbout()
{
	CAboutDlg aboutDlg;
	aboutDlg.DoModal();
}

void CMFCzipAppApp::OnAppTestZip()
{
	// Initialize COM and create an instance of the InterfaceImplementation class:
	HRESULT hr = CoInitialize(NULL);
	if (FAILED(hr))
	{
		printf("Couldn't initiate COM!... 0x%x\n", hr);
		return;
	}

	CDPolarZipWarper m_ZIP(this->m_pMainWnd);

	BOOL ret = m_ZIP.cppIzip != nullptr;
	if (!ret)
		return;
	
	//--------------------------------------------------------------------------
	// Ejemplo para comprimir un directorio
	m_ZIP.SetSourceDirectory(_T("C:\\Arktec\\LGT\\ConOcultos"));
	m_ZIP.SetExtractDirectory(_T(""));
	m_ZIP.SetZipFileName(_T("C:\\Arktec\\LGT\\ConOcultos.zip"));
	//Options
	m_ZIP.SetIncludeBaseDirectory(FALSE); //No existía en CDPolarZip
	m_ZIP.SetIncludeHiddenFiles(TRUE);
	m_ZIP.SetIncludeDirectoryEntries(FALSE);
	m_ZIP.SetOverwrite(FALSE);
	m_ZIP.SetRecurseSubDirectories(TRUE);
	m_ZIP.SetStorePaths(TRUE);
	m_ZIP.SetUsePassword(FALSE);
	m_ZIP.SetPassword(_T(""));
	m_ZIP.SetCompressionLevel(9);	// El máximo
	m_ZIP.SetNoCompressSuffixes(_T(""));
	m_ZIP.SetExcludeFileMask(_T(""));//_T("Backup\\*.*"));
	m_ZIP.SetIncludeFileMask(_T("*"));
	//m_ZIP.SetSkipFilesAfterDate(_T(""));  //¿lo necesitamos?
	//m_ZIP.SetSkipFilesBeforeDate(_T(""));
	if (m_ZIP.Add())
	{
		CString sErr = m_ZIP.GetLastZipError(hr);
		this->m_pMainWnd->MessageBox(sErr.GetString(), _T("Error"), MB_OK | MB_ICONSTOP);

	}

	//-----------------------------------------------------------------
	// Ejemplo para extraer todo
	m_ZIP.SetSourceDirectory(_T(""));
	m_ZIP.SetExtractDirectory(_T("C:\\Arktec\\LGT\\ConOcultos extraer"));
	m_ZIP.SetZipFileName(_T("C:\\Arktec\\LGT\\ConOcultos.zip"));
	if (m_ZIP.Extract())
	{
		CString sErr = m_ZIP.GetLastZipError(hr);
		this->m_pMainWnd->MessageBox(sErr.GetString(), _T("Error"), MB_OK | MB_ICONSTOP);

	}

	// Be a good citizen and clean up COM:
	CoUninitialize();
}

// Controladores de mensajes de CMFCzipAppApp





// ChildView.h: interfaz de la clase CChildView
//


#pragma once


// Ventana de CChildView

class CChildView : public CWnd
{
// Construcción
public:
	CChildView();

// Atributos
public:

// Operaciones
public:

// Reemplazos
	protected:
	virtual BOOL PreCreateWindow(CREATESTRUCT& cs);

// Implementación
public:
	virtual ~CChildView();

	// Funciones de asignación de mensajes generadas
protected:
	afx_msg void OnPaint();
	DECLARE_MESSAGE_MAP()
};



// StochasticServer.h : PROJECT_NAME Ӧ�ó������ͷ�ļ�
//

#pragma once

#ifndef __AFXWIN_H__
	#error "�ڰ������ļ�֮ǰ������stdafx.h�������� PCH �ļ�"
#endif

#include "resource.h"		// ������


// CStochasticServerApp: 
// �йش����ʵ�֣������ StochasticServer.cpp
//

class CStochasticServerApp : public CWinApp
{
public:
	CStochasticServerApp();

// ��д
public:
	virtual BOOL InitInstance();

// ʵ��

	DECLARE_MESSAGE_MAP()
};

extern CStochasticServerApp theApp;
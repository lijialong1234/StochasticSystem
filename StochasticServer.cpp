
// StochasticServer.cpp : ����Ӧ�ó��������Ϊ��
//

#include "stdafx.h"
#include "StochasticServer.h"
#include "StochasticServerDlg.h"
#include "SQLThreadConPool.h"
#include <iostream>
#include <string>
#include <fstream>
#include "sciter3\ui.cpp"
#ifdef _DEBUG
#define new DEBUG_NEW
#endif
//EXE�����ļ���·��
TCHAR g_ExePath[MAX_PATH];
//�����ļ�·��
TCHAR g_CfgFilePath[MAX_PATH];
//gsoapʹ��·��:char����
char g_ExePathA[MAX_PATH];
//ָ��MySQL��ConectionPoll
CSQLThreadConPool	*g_pMysqlCP = NULL;
volatile HWND g_hCurrentDlg = NULL;
// CStochasticServerApp
BEGIN_MESSAGE_MAP(CStochasticServerApp, CWinApp)
	ON_COMMAND(ID_HELP, &CWinApp::OnHelp)
END_MESSAGE_MAP()

// CStochasticServerApp ����

CStochasticServerApp::CStochasticServerApp()
{
	// ֧����������������
	m_dwRestartManagerSupportFlags = AFX_RESTART_MANAGER_SUPPORT_RESTART;

	// TODO: �ڴ˴���ӹ�����룬
	// ��������Ҫ�ĳ�ʼ�������� InitInstance ��
}


// Ψһ��һ�� CStochasticServerApp ����

CStochasticServerApp theApp;


// CStochasticServerApp ��ʼ��
void GetExePathA()
{
	USES_CONVERSION;
	sprintf_s(g_ExePathA, MAX_PATH, W2A(g_ExePath));
}
BOOL InitMySQL()
{
	int iRet = RestartMYSQLService(_T("StochasticServer"));
	if (iRet != 0)
	{
		return FALSE;
	}
	USES_CONVERSION;
	TCHAR sDB[MAX_PATH] = { 0 };
	TCHAR sIP[MAX_PATH] = { 0 };
	GetPrivateProfileString(_T("mysql"), _T("db"), _T("stochasticdb"), sDB, MAX_PATH, g_CfgFilePath);
	GetPrivateProfileString(_T("mysql"), _T("ip"), _T("localhost"), sIP, MAX_PATH, g_CfgFilePath);
	g_pMysqlCP = new CSQLThreadConPool(W2A(sDB), W2A(sIP), "root", "mkdsystem");
	try {
		mysqlpp::ScopedConnection cp(*g_pMysqlCP, true);
		if (!cp->thread_aware()) {
			return FALSE;
		}
		cp->disconnect();

	}
	catch (mysqlpp::Exception& e) {
		e.what();
		return FALSE;
	}
	return TRUE;
}
BOOL CStochasticServerApp::InitInstance()
{
	// ���һ�������� Windows XP �ϵ�Ӧ�ó����嵥ָ��Ҫ
	// ʹ�� ComCtl32.dll �汾 6 ����߰汾�����ÿ��ӻ���ʽ��
	//����Ҫ InitCommonControlsEx()��  ���򣬽��޷��������ڡ�
	INITCOMMONCONTROLSEX InitCtrls;
	InitCtrls.dwSize = sizeof(InitCtrls);
	// ��������Ϊ��������Ҫ��Ӧ�ó�����ʹ�õ�
	// �����ؼ��ࡣ
	InitCtrls.dwICC = ICC_WIN95_CLASSES;
	InitCommonControlsEx(&InitCtrls);

	CWinApp::InitInstance();
	TCHAR ExePath[MAX_PATH] = { 0 };
	if (GetModuleFileName(NULL, g_ExePath, MAX_PATH) == 0)
	{

		return FALSE;
	}
	(_tcsrchr(g_ExePath, _T('\\')))[1] = 0; //ɾ���ļ�����ע��·�������/����
	wsprintf(g_CfgFilePath, _T("%sconfig.ini"), g_ExePath);
	if (-1 == (_taccess(g_CfgFilePath, 0)))
	{
		//�����ļ�������
		return FALSE;
	}
	if (!InitMySQL())
	{
		return FALSE;
	}
	AfxEnableControlContainer();

	// ���� shell ���������Է��Ի������
	// �κ� shell ����ͼ�ؼ��� shell �б���ͼ�ؼ���
	CShellManager *pShellManager = new CShellManager;

	// ���Windows Native���Ӿ����������Ա��� MFC �ؼ�����������
	CMFCVisualManager::SetDefaultManager(RUNTIME_CLASS(CMFCVisualManagerWindows));

	// ��׼��ʼ��
	// ���δʹ����Щ���ܲ�ϣ����С
	// ���տ�ִ���ļ��Ĵ�С����Ӧ�Ƴ�����
	// ����Ҫ���ض���ʼ������
	// �������ڴ洢���õ�ע�����
	// TODO: Ӧ�ʵ��޸ĸ��ַ�����
	// �����޸�Ϊ��˾����֯��
	SetRegistryKey(_T("Ӧ�ó��������ɵı���Ӧ�ó���"));
	sciter::archive::instance().open(aux::elements_of(resources));
	SciterSetOption(NULL, SCITER_SET_DEBUG_MODE, TRUE);
	CStochasticServerDlg dlg;
	m_pMainWnd = &dlg;
	INT_PTR nResponse = dlg.DoModal();
	if (nResponse == IDOK)
	{
		// TODO: �ڴ˷��ô����ʱ��
		//  ��ȷ�������رնԻ���Ĵ���
	}
	else if (nResponse == IDCANCEL)
	{
		// TODO: �ڴ˷��ô����ʱ��
		//  ��ȡ�������رնԻ���Ĵ���
	}
	else if (nResponse == -1)
	{
		TRACE(traceAppMsg, 0, "����: �Ի��򴴽�ʧ�ܣ�Ӧ�ó���������ֹ��\n");
		TRACE(traceAppMsg, 0, "����: ������ڶԻ�����ʹ�� MFC �ؼ������޷� #define _AFX_NO_MFC_CONTROLS_IN_DIALOGS��\n");
	}

	// ɾ�����洴���� shell ��������
	if (pShellManager != NULL)
	{
		delete pShellManager;
	}

#ifndef _AFXDLL
	ControlBarCleanUp();
#endif

	// ���ڶԻ����ѹرգ����Խ����� FALSE �Ա��˳�Ӧ�ó���
	//  ����������Ӧ�ó������Ϣ�á�
	return FALSE;
}
/************************************************************************/
/* ��ʼ��MySQLģ��                                                       */
/************************************************************************/


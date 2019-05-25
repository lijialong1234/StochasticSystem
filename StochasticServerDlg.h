
// StochasticServerDlg.h : ͷ�ļ�
//

#pragma once
#include "sciter-x.h"
#include "tiscript-streams.hpp"
#include "sciter-x-api.h"
#include "aux-asset.h"
#include "sciter-x-threads.h"
#include "sciter-x-dom.hpp"
#include "sciter-x-host-callback.h"
#include "sciter-x-behavior.h"
#include "Jason/json.h"
#include "SQLThreadConPool.h"
#include <mysql++.h>
#include <ssqls.h>
#include "sciter-x-threads.h"
#include "sciter-x-host-callback.h"
#include "StochasticServer.h"
// CStochasticServerDlg �Ի���
class CStochasticServerDlg : public CDialogEx,
	public sciter::host<CStochasticServerDlg>,
	public sciter::event_handler
{
// ����
public:
	CStochasticServerDlg(CWnd* pParent = NULL);	// ��׼���캯��
	HWINDOW   get_hwnd();
	HINSTANCE get_resource_instance();
	// Sciter DOM event handlers, sciter::event_handler overridables 
	virtual bool handle_mouse(HELEMENT he, MOUSE_PARAMS& params);
	virtual bool handle_key(HELEMENT he, KEY_PARAMS& params);
	virtual bool handle_focus(HELEMENT he, FOCUS_PARAMS& params)  override { return false; }
	virtual bool handle_event(HELEMENT he, BEHAVIOR_EVENT_PARAMS& params);
	virtual bool handle_method_call(HELEMENT he, METHOD_PARAMS& params)  override { return false; }
// �Ի�������
#ifdef AFX_DESIGN_TIME
	enum { IDD = IDD_DIALOG1 };
#endif
public:
	struct markets {
		mysqlpp::sql_int  ID;
		mysqlpp::sql_text Registermark;
		mysqlpp::sql_text WordName;
		mysqlpp::sql_text LegalPerson;
		mysqlpp::sql_text BusinessPlace;
		mysqlpp::sql_text Type;
		mysqlpp::sql_text Scopes;
		mysqlpp::sql_text Phone;
	};
	protected:
	void UpdateWindowSize();
	virtual void DoDataExchange(CDataExchange* pDX);	// DDX/DDV ֧
// ʵ��
protected:
	HICON m_hIcon;

	// ���ɵ���Ϣӳ�亯��
	virtual BOOL OnInitDialog();
	afx_msg void OnSysCommand(UINT nID, LPARAM lParam);
	afx_msg void OnPaint();
	afx_msg HCURSOR OnQueryDragIcon();
	DECLARE_MESSAGE_MAP()
	BEGIN_FUNCTION_MAP
		FUNCTION_0("DepartmentList", DepartmentList);
	    FUNCTION_0("GetKeyStates", GetKeyStates);
		FUNCTION_1("StaffList", StaffList);
		FUNCTION_1("AddJOToTable", AddJOToTable);
		FUNCTION_1("LoadExcel", LoadExcel);
		FUNCTION_1("SinglePoshDoubleSet", SinglePoshDoubleSet);
		FUNCTION_1("ManyPoshDoubleSet", ManyPoshDoubleSet);
		FUNCTION_1("SaveToExcel", SaveToExcel);
		CHAIN_FUNCTION_MAP(CStochasticServerDlg);
	END_FUNCTION_MAP
public:
	BOOL isLoading;
	int FileType;//0,1��ʾ�г��ͼ����
	int ImportData(CString path);
	afx_msg void OnBnClickedOk();
	bool MysqlTest();
	std::string mDocPath;
	int DocLength;
private:
	//�洢ÿ��ִ����Ա��Ϣ
	std::vector<std::vector<std::string>>n_infos;
	//�洢��ȡ�г���Ϣ
	std::vector<markets>m_infos;
	json::value StaffList(json::value temp);
	json::value GetKeyStates();
	//��������
	static UINT InsertSQLThreadProc(LPVOID pParam);
	HANDLE							m_hSoapThread;
	//SOAP�߳�ID
	UINT							m_dwSoapThreadID;
public:
	int ReadExcelFile(std::string& m_filename);
	BOOL InsertTable(mysqlpp::ScopedConnection& cp,json::value temp);
	virtual BOOL PreTranslateMessage(MSG* pMsg);
	afx_msg int OnCreate(LPCREATESTRUCT lpCreateStruct);
	virtual BOOL PreCreateWindow(CREATESTRUCT& cs);
	virtual LRESULT WindowProc(UINT message, WPARAM wParam, LPARAM lParam);
//	afx_msg void OnWindowPosChanged(WINDOWPOS* lpwndpos);
	json::value LoadExcel(json::value filename);
	json::value AddJOToTable(json::value temp);
	BOOL NewStaff(json::value temp, mysqlpp::ScopedConnection& cp);
	int ReadStaffExcel(std::string& filename);
	json::value SinglePoshDoubleSet(json::value info);
	BOOL GetAllPairs();
	json::value ManyPoshDoubleSet(json::value info);
	json::value SaveToExcel(json::value group);
	json::value DepartmentList();
};

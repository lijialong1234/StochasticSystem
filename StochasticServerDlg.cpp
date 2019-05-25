
// StochasticServerDlg.cpp : 实现文件
//

#include "stdafx.h"
#include "StochasticServerDlg.h"
#include "afxdialogex.h"
#include <iostream>
#include <windows.h>
#include "libxl.h"
#include "Jason/json.h"
#include"GlobalFunctions.h"
#define MAXPATH 1024
using namespace libxl;
#pragma comment(lib, "lib_json.lib")
extern CSQLThreadConPool	*g_pMysqlCP;
extern CSQLThreadConPool	*g_pLogCP;
extern TCHAR g_ExePath[MAX_PATH];
extern volatile HWND g_hCurrentDlg;
extern const unsigned char resources[];
#ifdef _DEBUG
#define new DEBUG_NEW
#endif
#pragma warning(disable:4996)
sql_create_19(applyform_stock,
	1, 19,
	mysqlpp::sql_int, ApplyID,
	mysqlpp::sql_text, ApplyCode,
	mysqlpp::sql_text, PatientName,
	mysqlpp::sql_int, PatientAge,
	mysqlpp::sql_int, PatientSex,
	mysqlpp::sql_int, PatientID,
	mysqlpp::sql_text, PatientCode,
	mysqlpp::sql_text, VisitID,
	mysqlpp::sql_int, IDType,
	mysqlpp::sql_text, IDNumber,
	mysqlpp::sql_text, Phone,
	mysqlpp::sql_text, PatientAddress,
	mysqlpp::sql_text, Memo,
	mysqlpp::sql_text, ApplyTime,
	mysqlpp::sql_text, ApplyItem,
	mysqlpp::sql_int, SectionID,
	mysqlpp::sql_int, DoctorID,
	mysqlpp::sql_int, Finished,
	mysqlpp::sql_int, ReportCount
);
sql_create_8(applyforms,
	1, 8,
	mysqlpp::sql_int, ApplyID,
	mysqlpp::sql_int, PatientID,
	mysqlpp::sql_text, ApplyCode,
	mysqlpp::sql_text, ApplyItem,
	mysqlpp::sql_int, SectionID,
	mysqlpp::sql_int, DoctorID,
	mysqlpp::sql_text, Memo,
	mysqlpp::sql_text, ApplyTime
);
sql_create_12(patientinfo_stock,
	1, 12,
	mysqlpp::sql_int, PatientID,
	mysqlpp::sql_text, PatientName,
	mysqlpp::sql_int, PatientAge,
	mysqlpp::sql_int, PatientSex,
	mysqlpp::sql_text, PatientCode,
	mysqlpp::sql_text, VisitID,
	mysqlpp::sql_text, CreateTime,
	mysqlpp::sql_int, IDType,
	mysqlpp::sql_text, IDNumber,
	mysqlpp::sql_text, Phone,
	mysqlpp::sql_text, PatientAddress,
	mysqlpp::sql_text, PatientMemo);
sql_create_11(patientinfo,
	1, 11,
	mysqlpp::sql_text, PatientName,
	mysqlpp::sql_int, PatientAge,
	mysqlpp::sql_int, PatientSex,
	mysqlpp::sql_text, PatientCode,
	mysqlpp::sql_text, VisitID,
	mysqlpp::sql_int, IDType,
	mysqlpp::sql_text, IDNumber,
	mysqlpp::sql_text, CreateTime,
	mysqlpp::sql_text, Phone,
	mysqlpp::sql_text, PatientAddress,
	mysqlpp::sql_text, PatientMemo);
/*sql_create_7(staff,
	1, 7,
	mysqlpp::sql_text, Registermark,
	mysqlpp::sql_text, WordName,
	mysqlpp::sql_text, LegalPerson,
	mysqlpp::sql_text, BusinessPlace,
	mysqlpp::sql_text, Type,
	mysqlpp::sql_text, Scopes,
	mysqlpp::sql_text, Phone);*/
sql_create_8(market,
	1, 8,
	mysqlpp::sql_int, ID,
	mysqlpp::sql_text, Registermark,
	mysqlpp::sql_text, WordName,
	mysqlpp::sql_text, LegalPerson,
	mysqlpp::sql_text, BusinessPlace,
	mysqlpp::sql_text, Type,
	mysqlpp::sql_text, Scopes,
	mysqlpp::sql_text, Phone);
sql_create_6(staffs,
	1, 6,
	mysqlpp::sql_int, ID,
	mysqlpp::sql_text, Pair,
	mysqlpp::sql_text, Name,
	mysqlpp::sql_text, Duty,
	mysqlpp::sql_text, Number,
	mysqlpp::sql_text, Remarks);
sql_create_2(department,
	1, 2,
	mysqlpp::sql_int, ID,
	mysqlpp::sql_text, Pair);
// 用于应用程序“关于”菜单项的 CAboutDlg 对话框

class CAboutDlg : public CDialogEx
{
public:
	CAboutDlg();

	// 对话框数据
#ifdef AFX_DESIGN_TIME
	enum { IDD = IDD_ABOUTBOX };
#endif

protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV 支持

// 实现
protected:
	DECLARE_MESSAGE_MAP()
};

CAboutDlg::CAboutDlg() : CDialogEx(IDD_ABOUTBOX)
{
}

void CAboutDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialogEx::DoDataExchange(pDX);
}

BEGIN_MESSAGE_MAP(CAboutDlg, CDialogEx)
END_MESSAGE_MAP()


// CStochasticServerDlg 对话框


//Soap线程
unsigned __stdcall SoapThread(LPVOID lParam)
{
	CStochasticServerDlg* dlg = (CStochasticServerDlg*)lParam;
	dlg->ReadExcelFile(dlg->mDocPath);
	return 0;
}
CStochasticServerDlg::CStochasticServerDlg(CWnd* pParent /*=NULL*/)
	: CDialogEx(IDD_STOCHASTICSERVER_DIALOG, pParent), m_dwSoapThreadID(0)
	, mDocPath(""), isLoading(FALSE), DocLength(0), FileType(0)
{
	m_hIcon = AfxGetApp()->LoadIcon(IDI_ICON3);
	m_hSoapThread = (HANDLE)_beginthreadex(NULL, 0, SoapThread, this, CREATE_SUSPENDED, &m_dwSoapThreadID);
}

HWINDOW CStochasticServerDlg::get_hwnd()
{
	return this->GetSafeHwnd();
}

HINSTANCE CStochasticServerDlg::get_resource_instance()
{
	return theApp.m_hInstance;
}

bool CStochasticServerDlg::handle_mouse(HELEMENT he, MOUSE_PARAMS & params)
{
	bool bRet = false;
	std::wstring Id;
	CString Cmd;
	sciter::dom::element root = get_root();
	switch (params.button_state)
	{
	case MAIN_MOUSE_BUTTON:
		if (MOUSE_UP == params.cmd - PHASE_MASK::SINKING)
		{
			Id = id_or_name_or_tag(params.target).c_str();
			sciter::dom::element esrc = params.target;
			const wchar_t *pAT = esrc.get_attribute("acttype").c_str();
			sciter::dom::element root;
			sciter::dom::element table;
			root = get_root();
			wchar_t PageNumber[128] = { 0 };
			if (wcsncmp(Id.c_str(), L"Cmd_Btn_Exit-", strlen("Cmd_Btn_Exit")) == 0)
			{
				//
				OnCancel();
			}
			else if (wcsncmp(Id.c_str(), L"Cmd_Btn_Min-", strlen("Cmd_Btn_Min")) == 0)
			{
				//
				ShowWindow(SW_MINIMIZE);
				bRet = false;
			}
		}


		break;
	}
	//return SciterTraverseUIEvent(params.cmd,(LPVOID)params, NULL);
	return bRet;
}

bool CStochasticServerDlg::handle_key(HELEMENT he, KEY_PARAMS & params)
{
	std::wstring Id;
	CString Cmd;
	std::string pret;
	bool bRet = false;
	if (params.key_code == VK_RETURN&&params.cmd == SINKING) {

		sciter::dom::element root;
		sciter::dom::element table;
		root = get_root();
		wchar_t PageNumber[128] = { 0 };
		return true;
	}
	else if (VK_BACK == params.key_code)
	{
		return false;
	}
	return false;
}

bool CStochasticServerDlg::handle_event(HELEMENT he, BEHAVIOR_EVENT_PARAMS & params)
{
	std::wstring Id;
	CString Cmd;
	//char* server = "http://localhost:4567";
	//char server[MAX_PATH];

	std::string pret;
	bool bRet = false;
	switch (params.cmd)
	{
	case DOCUMENT_COMPLETE:
		UpdateWindowSize();
		break;
	case BUTTON_CLICK:
		Id = id_or_name_or_tag(params.heTarget).c_str();
		sciter::dom::element esrc = params.heTarget;
		const wchar_t *pAT = esrc.get_attribute("acttype").c_str();
		sciter::dom::element root;
		sciter::dom::element table;
		root = get_root();
		wchar_t PageNumber[128] = { 0 };
		if (wcsncmp(Id.c_str(), L"Cmd_Btn_Exit", strlen("Cmd_Btn_Exit")) == 0)
		{
			//
			OnCancel();
		}
		else if (wcsncmp(Id.c_str(), L"Cmd_Btn_Min", strlen("Cmd_Btn_Min")) == 0)
		{
			ShowWindow(SW_MINIMIZE);
			bRet = false;
		}
		break;
	}
	return false;
}

void CStochasticServerDlg::UpdateWindowSize()
{
	sciter::dom::element root = get_root();
	sciter::dom::element DisplaySet = root.get_element_by_id(L"display");
	CRect rt;
	if (DisplaySet.is_valid())
	{
		rt.left = DisplaySet.get_attribute_int("left");
		rt.top = DisplaySet.get_attribute_int("top");
		rt.right = DisplaySet.get_attribute_int("right");
		rt.bottom = DisplaySet.get_attribute_int("bottom");
	}
	if (rt.right == -1 && rt.bottom == -1)
	{
		//
		rt.right = GetSystemMetrics(SM_CXSCREEN);
		rt.bottom = GetSystemMetrics(SM_CYSCREEN);
	}
	int nWidth = GetPrivateProfileInt(_T("System"), _T("Width"), 1004, g_ExePath);
	int nHeight = GetPrivateProfileInt(_T("System"), _T("Height"), 700, g_ExePath);
	int iWidth = rt.Width() == 0 ? nWidth : rt.Width();
	int iHeight = rt.Height() == 0 ? nHeight : rt.Height();
	int iScreenCX = GetSystemMetrics(SM_CXSCREEN);
	int iScreenCY = GetSystemMetrics(SM_CYSCREEN);
	int iLeft = (iScreenCX - iWidth) / 2;
	int iTop = (iScreenCY - iHeight) / 2;
	RECT rt1;
	::SystemParametersInfo(SPI_GETWORKAREA, 0, &rt1, 0);    // 获得工作区大小
	if (iHeight == iScreenCY)
	{
		iHeight = rt1.bottom;
	}
	::SetWindowPos(this->GetSafeHwnd(), HWND_NOTOPMOST, iLeft, iTop, iWidth, iHeight, SWP_SHOWWINDOW | SWP_NOZORDER);
}

void CStochasticServerDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialogEx::DoDataExchange(pDX);
}

BEGIN_MESSAGE_MAP(CStochasticServerDlg, CDialogEx)
	ON_WM_SYSCOMMAND()
	ON_WM_PAINT()
	ON_WM_QUERYDRAGICON()
	ON_BN_CLICKED(IDOK, &CStochasticServerDlg::OnBnClickedOk)
	ON_WM_CREATE()
	//	ON_WM_WINDOWPOSCHANGED()
END_MESSAGE_MAP()


// CStochasticServerDlg 消息处理程序

BOOL CStochasticServerDlg::OnInitDialog()
{
	CDialogEx::OnInitDialog();
	g_hCurrentDlg = this->GetSafeHwnd();
	// 将“关于...”菜单项添加到系统菜单中。

	// IDM_ABOUTBOX 必须在系统命令范围内。
	ASSERT((IDM_ABOUTBOX & 0xFFF0) == IDM_ABOUTBOX);
	ASSERT(IDM_ABOUTBOX < 0xF000);

	CMenu* pSysMenu = GetSystemMenu(FALSE);
	if (pSysMenu != NULL)
	{
		BOOL bNameValid;
		CString strAboutMenu;
		bNameValid = strAboutMenu.LoadString(IDS_ABOUTBOX);
		ASSERT(bNameValid);
		if (!strAboutMenu.IsEmpty())
		{
			pSysMenu->AppendMenu(MF_SEPARATOR);
			pSysMenu->AppendMenu(MF_STRING, IDM_ABOUTBOX, strAboutMenu);
		}
	}

	// 设置此对话框的图标。  当应用程序主窗口不是对话框时，框架将自动
	//  执行此操作
	SetIcon(m_hIcon, TRUE);			// 设置大图标
	SetIcon(m_hIcon, FALSE);		// 设置小图标

	ShowWindow(SW_NORMAL);

	// TODO: 在此添加额外的初始化代码

	g_hCurrentDlg = this->GetSafeHwnd();
#ifdef _ONTIME_DEBUG
	TCHAR loginpage[MAX_PATH] = { 0 };
	wsprintf(loginpage, _T("%sLayout\\login.htm"), g_ExePath);
	if (-1 == (_taccess(loginpage, 0))
		)
	{
		return FALSE;
	}
	TCHAR szURL[MAX_PATH] = { 0 };
	wsprintf(szURL, _T("file:///%s"), loginpage);
	load_file(szURL);
#else
	load_file(L"this://app/login.htm");
#endif
	return TRUE;  // 除非将焦点设置到控件，否则返回 TRUE
}

void CStochasticServerDlg::OnSysCommand(UINT nID, LPARAM lParam)
{
	if ((nID & 0xFFF0) == IDM_ABOUTBOX)
	{
		CAboutDlg dlgAbout;
		dlgAbout.DoModal();
	}
	else
	{
		CDialogEx::OnSysCommand(nID, lParam);
	}
}

// 如果向对话框添加最小化按钮，则需要下面的代码
//  来绘制该图标。  对于使用文档/视图模型的 MFC 应用程序，
//  这将由框架自动完成。

void CStochasticServerDlg::OnPaint()
{
	if (IsIconic())
	{
		CPaintDC dc(this); // 用于绘制的设备上下文

		SendMessage(WM_ICONERASEBKGND, reinterpret_cast<WPARAM>(dc.GetSafeHdc()), 0);

		// 使图标在工作区矩形中居中
		int cxIcon = GetSystemMetrics(SM_CXICON);
		int cyIcon = GetSystemMetrics(SM_CYICON);
		CRect rect;
		GetClientRect(&rect);
		int x = (rect.Width() - cxIcon + 1) / 2;
		int y = (rect.Height() - cyIcon + 1) / 2;

		// 绘制图标
		dc.DrawIcon(x, y, m_hIcon);
	}
	else
	{
		CDialogEx::OnPaint();
	}
}

//当用户拖动最小化窗口时系统调用此函数取得光标
//显示。
HCURSOR CStochasticServerDlg::OnQueryDragIcon()
{
	return static_cast<HCURSOR>(m_hIcon);
}



int CStochasticServerDlg::ImportData(CString path)
{
	Book* book = xlCreateBook();
	const wchar_t * x = L"Halil Kural";
	const wchar_t * y = L"windows-2723210a07c4e90162b26966a8jcdboe";
	book->setKey(x, y);

	if (book)
	{
		int f[6];

		f[0] = book->addCustomNumFormat(L"0.0");
		f[1] = book->addCustomNumFormat(L"0.00");
		f[2] = book->addCustomNumFormat(L"0.000");
		f[3] = book->addCustomNumFormat(L"0.0000");
		f[4] = book->addCustomNumFormat(L"#,###.00 $");
		f[5] = book->addCustomNumFormat(L"#,###.00 $[Black][<1000];#,###.00 $[Red][>=1000]");

		Format* format[6];
		for (int i = 0; i < 6; ++i) {
			format[i] = book->addFormat();
			format[i]->setNumFormat(f[i]);
		}

		Sheet* sheet = book->addSheet(L"Custom formats");
		if (sheet)
		{
			sheet->setCol(0, 0, 20);
			sheet->writeNum(2, 0, 25.718, format[0]);
			sheet->writeNum(3, 0, 25.718, format[1]);
			sheet->writeNum(4, 0, 25.718, format[2]);
			sheet->writeNum(5, 0, 25.718, format[3]);

			sheet->writeNum(7, 0, 1800.5, format[4]);

			sheet->writeNum(9, 0, 500, format[5]);
			sheet->writeNum(10, 0, 1600, format[5]);
		}

		if (book->save(L"custom.xls"))
		{
			//::ShellExecute(NULL, L"open", L"custom.xls", NULL, NULL, SW_SHOW);
		}

		book->release();
		return 0;
	}
	return -1;
}


void CStochasticServerDlg::OnBnClickedOk()
{
	// TODO: 在此添加控件通知处理程序代码
	//ImportData(L"sheets.xls");
	//MysqlTest();
	//AfxBeginThread(InsertSQLThreadProc, (LPVOID)this);
	ResumeThread(m_hSoapThread);
	//CDialogEx::OnOK();
}


bool CStochasticServerDlg::MysqlTest()
{
	mysqlpp::ScopedConnection cp(*g_pMysqlCP, true);

	if (cp) {
		SYSTEMTIME st;
		GetSystemTime(&st);
		char strSQL[MAXPATH] = { 0 };
		//检验者信息更新
		sprintf(strSQL, "select patientinfo.* from patientinfo ORDER BY PatientID DESC;");
		mysqlpp::Query query1 = cp->query(strSQL);
		std::vector<patientinfo_stock> rpatients;
		query1.storein(rpatients);
		query1.clear();
		size_t len = rpatients.size();
		//Json::Value Result;
		for (size_t i = 0; i < len; i++) {

			patientinfo_stock pio = rpatients[i];
			//Json::Value temp;
			//UTF82C(ins.HospitalName.c_str(), sout);
			char *pTem = new char[1024];
			sprintf(pTem, "%s", pio.PatientCode.c_str());
			CString tempRes(pTem);
			MessageBox(tempRes);
			delete[]pTem;
		}
		return true;
	}
	else
	{
		return false;
	}
	cp->disconnect();
}

json::value CStochasticServerDlg::StaffList(json::value temp)
{
	json::value res;
	mysqlpp::ScopedConnection cp(*g_pMysqlCP, true);
	if (cp) {
		char strSQL[MAXPATH] = { 0 };
		//清除之前数据
		m_infos.clear();
		//市场信息更新
		int ncount = temp.get_item("count").get<int>();
		//SELECT * FROM staff ORDER BY RAND() limit 30
		sprintf(strSQL, "SELECT * FROM market WHERE ID >= ((SELECT MAX(ID) FROM market)\
			   -(SELECT MIN(ID) FROM market)) * RAND()\
               + (SELECT MIN(ID) FROM market) LIMIT %d;", ncount);
		mysqlpp::Query query1 = cp->query(strSQL);
		std::vector<market> rstaffs;
		query1.storein(rstaffs);
		query1.clear();
		//保存抽取结果到抽取市场vector中
		size_t len = rstaffs.size();
		Json::Value Result;
		for (size_t i = 0; i < len; i++) {
			market pio = rstaffs[i];
			markets ps;
			Json::Value temp;
			ps.ID = pio.ID;
			temp["ID"] = pio.ID;
			ps.BusinessPlace = pio.BusinessPlace;
			temp["BusinessPlace"] = pio.BusinessPlace;
			ps.LegalPerson = pio.LegalPerson;
			temp["LegalPerson"] = pio.LegalPerson;
			ps.Phone = pio.Phone;
			temp["Phone"] = pio.Phone;
			ps.Registermark = pio.Registermark;
			temp["Registermark"] = pio.Registermark;
			ps.Type = pio.Type;
			temp["Type"] = pio.Type;
			ps.WordName = pio.WordName;
			temp["WordName"] = pio.WordName;
			ps.Scopes = pio.Scopes;
			temp["Scopes"] = pio.Scopes;
			Result.append(temp);
			m_infos.push_back(ps);
		}
		std::string sRet = Result.toStyledString();
		Json::Reader reader;
		Json::Value root;
		Json::Value jret;
		// reader将Json字符串解析到root，root将包含Json里所有子元素  
		if (reader.parse(sRet.c_str(), root))
		{
			size_t length = root.size();
			for (size_t i = 0; i < length; i++)
			{
				Json::Value jv = root[i];
				std::string str;
				UTF82C(jv["BusinessPlace"].asString().c_str(), str);
				jv["BusinessPlace"] = str;
				UTF82C(jv["LegalPerson"].asString().c_str(), str);
				jv["LegalPerson"] = str;
				UTF82C(jv["Registermark"].asString().c_str(), str);
				jv["Registermark"] = str;
				UTF82C(jv["Type"].asString().c_str(), str);
				jv["Type"] = str;
				UTF82C(jv["WordName"].asString().c_str(), str);
				jv["WordName"] = str;
				UTF82C(jv["Scopes"].asString().c_str(), str);
				jv["Scopes"] = str;
				jret.append(jv);
			}
			sRet = jret.toStyledString();
		}
		res["str"] = sRet;
		return res;
	}
	cp->disconnect();
	return res;
}

json::value CStochasticServerDlg::GetKeyStates()
{
	json::value temp;
	//GetKeyState(VK_CAPITAL)
	if (FALSE == this->isLoading)
	{
		temp["result"] = 0;//0表示未加载
		temp["DocLength"] = DocLength;//0表示未加载
	}
	else
	{
		temp["result"] = 1;
	}
	return temp;
}




int CStochasticServerDlg::ReadExcelFile(std::string& m_filename)
{
	mysqlpp::ScopedConnection cp(*g_pMysqlCP, true);
	std::string::size_type idx;
	idx = m_filename.find(".xlsx");
	Book* book = xlCreateXMLBook();
	if (idx == std::string::npos)//不存在
	{
		book = xlCreateBook();
	}
	const wchar_t * x = L"Halil Kural";
	const wchar_t * y = L"windows-2723210a07c4e90162b26966a8jcdboe";
	book->setKey(x, y);
	sciter::dom::element root;
	sciter::dom::element table;
	root = get_root();
	int ncount = 0;
	if (book&&cp)
	{
		if (book->load(StringToWstring(m_filename).c_str()))
		{
			Sheet* sheet = book->getSheet(0);
			if (sheet)
			{
				int nrow = 1;
				int ncolumn = 0;

				while (NULL != sheet->readStr(nrow, 0))
				{
					json::value temp;
					while (TRUE)
					{
						auto str = sheet->readStr(nrow, ncolumn);
						if (str)
						{
							switch (ncolumn)
							{
							case 0:temp.set_item("ID", _wtoi(str)); break;
							case 1:temp.set_item("Registermark", str); break;
							case 2:temp.set_item("WordName", str); break;
							case 3:temp.set_item("LegalPerson", str); break;
							case 4:temp.set_item("BusinessPlace", str); break;
							case 5:temp.set_item("Type", str); break;
							case 6:temp.set_item("Scopes", str); break;
							case 7:temp.set_item("Phone", str); break;
							default:break;
							}
							ncolumn++;
						}
						else
						{
							break;
						}
					}
					ncount++;
					InsertTable(cp, temp);
					nrow++;
					ncolumn = 0;
				}
			}
		}
		book->release();
		cp->disconnect();
	}
	return ncount;
}


BOOL CStochasticServerDlg::InsertTable(mysqlpp::ScopedConnection& cp, json::value temp)
{
	std::string UTF8Str;
	Json::Value JSInfo;
	USES_CONVERSION;
	int num = temp.get_item("ID").get<int>();
	auto str = temp.get_item("Registermark").get(L"");
	std::string Registermark = W2A(str.c_str());

	str = temp.get_item("WordName").get(L"");
	C2UTF8(W2A(str.c_str()), UTF8Str);
	std::string WordName = UTF8Str;

	str = temp.get_item("LegalPerson").get(L"");
	C2UTF8(W2A(str.c_str()), UTF8Str);
	std::string LegalPerson = UTF8Str;

	str = temp.get_item("BusinessPlace").get(L"");
	C2UTF8(W2A(str.c_str()), UTF8Str);
	std::string BusinessPlace = UTF8Str;

	str = temp.get_item("Type").get(L"");
	C2UTF8(W2A(str.c_str()), UTF8Str);
	std::string Type = UTF8Str;

	str = temp.get_item("Scopes").get(L"");
	C2UTF8(W2A(str.c_str()), UTF8Str);
	std::string Scopes = UTF8Str;

	str = temp.get_item("Phone").get(L"");
	std::string Phone = W2A(str.c_str());

	char strSQL[MAXPATH] = { 0 };
	mysqlpp::Query query = cp->query(strSQL);
	market row(num, Registermark, WordName, LegalPerson, BusinessPlace, Type, Scopes, Phone);
	query.insert(row);
	query.execute();
	query.clear();
	return 0;
}


BOOL CStochasticServerDlg::PreTranslateMessage(MSG* pMsg)
{
	// TODO: 在此添加专用代码和/或调用基类
	LRESULT lResult;
	BOOL    bHandled;

	if (pMsg->message == WM_CHAR) {
		lResult = SciterProcND(this->GetSafeHwnd(), pMsg->message, pMsg->wParam, pMsg->lParam, &bHandled);
		//if (bHandled)      // if it was handled by the Sciter
		//return lResult; // then no further processing is required.
	}
	if (pMsg->wParam == VK_BACK || pMsg->wParam == VK_RETURN || pMsg->wParam == VK_ESCAPE || pMsg->wParam == VK_SPACE)
	{
		lResult = SciterProcND(this->GetSafeHwnd(), pMsg->message, pMsg->wParam, pMsg->lParam, &bHandled);
		if (bHandled)      // if it was handled by the Sciter
			return true; // then no further processing is required.
						 //pMsg->wParam = 0 ;
		return true;
	}
	else
	{
		return __super::PreTranslateMessage(pMsg);
	}
}


int CStochasticServerDlg::OnCreate(LPCREATESTRUCT lpCreateStruct)
{
	if (__super::OnCreate(lpCreateStruct) == -1)
		return -1;

	// TODO:  在此添加您专用的创建代码
	//SetWindowLongPtr(GetSafeHwnd(), GWLP_USERDATA, LONG_PTR(this));
	SetWindowLong(m_hWnd, GWL_EXSTYLE, GetWindowLong(m_hWnd, GWL_EXSTYLE) | WS_EX_APPWINDOW | WS_CHILD);
	this->setup_callback(); // attach sciter::host callbacks
	sciter::attach_dom_event_handler(this->GetSafeHwnd(), this); // attach this as a DOM events 
	return 0;
}


BOOL CStochasticServerDlg::PreCreateWindow(CREATESTRUCT& cs)
{
	// TODO: 在此添加专用代码和/或调用基类

	if (!CWnd::PreCreateWindow(cs))
		return FALSE;

	//cs.style&=~WS_MAXIMIZEBOX;  //禁用最大化按钮
	//cs.style&=~WS_THICKFRAME;  //禁止调整窗口大小
	cs.dwExStyle |= WS_EX_CLIENTEDGE;
	cs.style &= ~WS_BORDER;
	cs.style &= ~WS_MAXIMIZEBOX;
	cs.lpszClass = AfxRegisterWndClass(CS_HREDRAW | CS_VREDRAW | CS_DBLCLKS,
		::LoadCursor(NULL, IDC_ARROW), reinterpret_cast<HBRUSH>(COLOR_WINDOW + 1), NULL);


	return __super::PreCreateWindow(cs);
}


LRESULT CStochasticServerDlg::WindowProc(UINT message, WPARAM wParam, LPARAM lParam)
{
	// TODO: 在此添加专用代码和/或调用基类

	LRESULT lResult;
	BOOL    bHandled;

	lResult = SciterProcND(this->GetSafeHwnd(), message, wParam, lParam, &bHandled);
	if (bHandled)      // if it was handled by the Sciter
		return lResult; // then no further processing is required.
	return __super::WindowProc(message, wParam, lParam);
}


//void CStochasticServerDlg::OnWindowPosChanged(WINDOWPOS* lpwndpos)
//{
//	__super::OnWindowPosChanged(lpwndpos);
//
//	// TODO: 在此处添加消息处理程序代码
//}

UINT CStochasticServerDlg::InsertSQLThreadProc(LPVOID pParam)
{
	CStochasticServerDlg* dlg = (CStochasticServerDlg*)pParam;
	dlg->isLoading = TRUE;
	if (0 == dlg->FileType)
		dlg->DocLength = dlg->ReadExcelFile(dlg->mDocPath);
	else if (1 == dlg->FileType)
		dlg->DocLength = dlg->ReadStaffExcel(dlg->mDocPath);
	dlg->GetAllPairs();//导入完成后执行部门表的新增
	dlg->isLoading = FALSE;
	return 0;
}

json::value CStochasticServerDlg::LoadExcel(json::value filename)
{
	json::value ret;
	/*auto tablename = filename.to_string();
	//转化为相对路径
	CString mfilename = tablename.c_str();
	int pos = mfilename.ReverseFind(L':');
	mfilename = mfilename.Mid(pos - 1);*/
	auto str = filename.get_item("path").get(L"");
	int type = filename.get_item("type").get<int>();
	USES_CONVERSION;
	//std::string path = UnicodeToANSI((const std::wstring)mfilename);
	this->mDocPath = UnicodeToANSI((const std::wstring)str);
	this->FileType = type;
	//让专门的线程来做插入数据库操作，防止主线程阻塞
	AfxBeginThread(InsertSQLThreadProc, (LPVOID)this);
	//ResumeThread(SoapThread);
	//ReadExcelFile(path);
	return ret;
}


json::value CStochasticServerDlg::AddJOToTable(json::value temp)
{
	return json::value();
}


BOOL CStochasticServerDlg::NewStaff(json::value temp, mysqlpp::ScopedConnection& cp)
{
	std::string UTF8Str;
	Json::Value JSInfo;
	USES_CONVERSION;
	int num = temp.get_item("ID").get<int>();
	auto str = temp.get_item("Pair").get(L"");
	C2UTF8(W2A(str.c_str()), UTF8Str);
	std::string Pair = UTF8Str;

	str = temp.get_item("Name").get(L"");
	C2UTF8(W2A(str.c_str()), UTF8Str);
	std::string Name = UTF8Str;


	str = temp.get_item("Duty").get(L"");
	C2UTF8(W2A(str.c_str()), UTF8Str);
	std::string Duty = UTF8Str;

	str = temp.get_item("Number").get(L"");
	C2UTF8(W2A(str.c_str()), UTF8Str);
	std::string Number = UTF8Str;

	str = temp.get_item("Remarks").get(L"");
	C2UTF8(W2A(str.c_str()), UTF8Str);
	std::string Remarks = UTF8Str;

	char strSQL[MAXPATH] = { 0 };
	mysqlpp::Query query = cp->query(strSQL);
	//sprintf(strSQL, "insert into staff values(%d,%s,%s,%s,%s,%s);",num, Firms.c_str(), Name.c_str(), Duty.c_str(), Number.c_str(), Remark.c_str());
	// = cp->query(strSQL);
	staffs row(num, Pair.c_str(), Name.c_str(), Duty.c_str(), Number.c_str(), Remarks.c_str());
	query.insert(row);
	query.execute();
	query.clear();
	return 0;
}


int CStochasticServerDlg::ReadStaffExcel(std::string& filename)
{
	mysqlpp::ScopedConnection cp(*g_pMysqlCP, true);
	std::string::size_type idx;
	idx = filename.find(".xlsx");
	Book* book = xlCreateXMLBook();
	if (idx == std::string::npos)//不存在
	{
		book = xlCreateBook();
	}
	const wchar_t * x = L"Halil Kural";
	const wchar_t * y = L"windows-2723210a07c4e90162b26966a8jcdboe";
	book->setKey(x, y);
	sciter::dom::element root;
	sciter::dom::element table;
	root = get_root();
	int ncount = 0;
	if (book&&cp)
	{
		if (book->load(StringToWstring(filename).c_str()))
		{
			Sheet* sheet = book->getSheet(0);
			if (sheet)
			{
				int nrow = 1;
				int ncolumn = 0;
				while (sheet->readStr(nrow, 0) || sheet->readNum(nrow, 0))
				{
					json::value temp;
					while (TRUE)
					{
						if (NULL == sheet->readStr(nrow, ncolumn) && NULL == sheet->readNum(nrow, ncolumn))
							break;
						auto str = sheet->readStr(nrow, ncolumn);
						if (!str)
						{
							std::string emp = std::to_string((int)sheet->readNum(nrow, ncolumn));
							USES_CONVERSION;
							str = A2W(emp.c_str());
						}
						switch (ncolumn)
						{
						case 0:temp.set_item("ID", _wtoi(str)); break;
						case 1:temp.set_item("Pair", str); break;
						case 2:temp.set_item("Name", str); break;
						case 3:temp.set_item("Duty", str); break;
						case 4:temp.set_item("Number", str); break;
						case 5:temp.set_item("Remarks", str); break;
						default:break;
						}
						ncolumn++;
					}
					ncount++;
					NewStaff(temp, cp);
					nrow++;
					ncolumn = 0;
				}
			}
		}
		book->release();
		cp->disconnect();
	}
	return ncount;
}

//单部门双随机算法，传递参数为组数+每组人数
json::value CStochasticServerDlg::SinglePoshDoubleSet(json::value info)
{
	json::value res;
	mysqlpp::ScopedConnection cp(*g_pMysqlCP, true);
	if (cp) {
		char strSQL[MAXPATH] = { 0 };
		//清除执法人员信息
		n_infos.clear();
		int GroupCount = info.get_item("GroupCount").get<int>();
		int ncount = info.get_item("count").get<int>();
		mysqlpp::Query query1 = cp->query(strSQL);
		query1.clear();
		sprintf(strSQL, "SELECT * FROM department ORDER BY RAND() LIMIT %d;", GroupCount);
		query1 = cp->query(strSQL);
		std::vector<department>rpartments;
		query1.storein(rpartments);
		query1.clear();
		size_t len = rpartments.size();
		Json::Value Result;
		for (size_t i = 0; i < len; i++) {
			department pio = rpartments[i];
			Json::Value temp;
			temp["ID"] = pio.ID;
			temp["Pair"] = pio.Pair;
			Result.append(temp);
		}
		std::string sRet = Result.toStyledString();
		Json::Reader reader;
		Json::Value root;
		Json::Value jret;
		// reader将Json字符串解析到root，root将包含Json里所有子元素  

		if (reader.parse(sRet.c_str(), root))
		{
			size_t length = root.size();
			for (size_t i = 0; i < length; i++)
			{
				Json::Value jv = root[i];
				std::string str;
				str=jv["Pair"].asString();
				//C2UTF8(str.c_str(), str);
				sprintf(strSQL, "SELECT *\
								 FROM staffs\
								 WHERE Pair='%s'\
								 ORDER BY rand() LIMIT %d;", str.c_str(),ncount);
				query1 = cp->query(strSQL);
				std::vector<staffs>rstaffs;
				query1.storein(rstaffs);
				size_t lens = rstaffs.size();
				query1.clear();
				//用于存储抽取执法人员组别信息
				std::vector<std::string>tempRes;
				for (size_t j = 0; j< lens; j++) {
					staffs pio = rstaffs[j];
					Json::Value temp;
					temp["ID"] = pio.ID;
					//存储数据所在组号
					temp["Group"] = i+1;

					std::string str;
					UTF82C(pio.Pair.c_str(), str);
					temp["Pair"] = str;

					UTF82C(pio.Duty.c_str(), str);
					temp["Duty"] = str;
					
					//存储执法人员姓名，用于导出时使用
					tempRes.push_back(pio.Name);

					UTF82C(pio.Name.c_str(), str);
					temp["Name"] = str;
					
					UTF82C(pio.Remarks.c_str(), str);
					temp["Remarks"] = str;

					UTF82C(pio.Number.c_str(), str);
					temp["Number"] = str;

					jret.append(temp);
				}
				//对不同组别加入存储vector
				n_infos.push_back(tempRes);
			}
			sRet = jret.toStyledString();
		}
		res["str"] = sRet;
		return res;
	}
	cp->disconnect();
	return res;
}

//生成部门表
BOOL CStochasticServerDlg::GetAllPairs()
{
	mysqlpp::ScopedConnection cp(*g_pMysqlCP, true);
	if (cp)
	{
		char strSQL[MAXPATH] = { 0 };
		sprintf(strSQL, "INSERT into department\
			select ID,Pair\
			from staffs\
			GROUP BY Pair;\
			");
		mysqlpp::Query query = cp->query(strSQL);
		query.execute();
		query.clear();
	}
	cp->disconnect();
	return TRUE;
}

//多部门双随机算法，传递参数为组数+每组人数+领头部门
json::value CStochasticServerDlg::ManyPoshDoubleSet(json::value info)
{
	json::value res;
	std::string sRet;
	USES_CONVERSION;
	mysqlpp::ScopedConnection cp(*g_pMysqlCP, true);
	
	if (cp) {
		char strSQL[MAXPATH] = { 0 };
		//清除所有
		n_infos.clear();
		int GroupCount = info.get_item("GroupCount").get<int>();
		int ncount = info.get_item("count").get<int>();
		std::string  department_lead = W2A(info.get_item("department_lead").get(L"").c_str());
		C2UTF8(department_lead.c_str(), department_lead);
		mysqlpp::Query query1 = cp->query(strSQL);
		query1.clear();
		//随机领头部门执法人员
		sprintf(strSQL, "SELECT * FROM staffs WHERE Pair='%s' ORDER BY RAND() LIMIT %d;", department_lead.c_str(), GroupCount);
		query1 = cp->query(strSQL);
		std::vector<staffs>rstaffs_leader;
		query1.storein(rstaffs_leader);
		query1.clear();
		size_t len = rstaffs_leader.size();
		Json::Value Result;
		std::vector<std::string>tempRes;
		Json::Value jret;
		for (size_t i = 0; i < len; i++) {
			staffs pio = rstaffs_leader[i];
			Json::Value temp;
			temp["ID"] = pio.ID;
			//存储数据所在组号
			temp["Group"] = i + 1;

			std::string str;
			UTF82C(pio.Pair.c_str(), str);
			temp["Pair"] = str;

			UTF82C(pio.Duty.c_str(), str);
			temp["Duty"] = str;

			//存储执法人员姓名，用于导出时使用
			tempRes.push_back(pio.Name);

			UTF82C(pio.Name.c_str(), str);
			temp["Name"] = str;

			UTF82C(pio.Remarks.c_str(), str);
			temp["Remarks"] = str;

			UTF82C(pio.Number.c_str(), str);
			temp["Number"] = str;

			jret.append(temp);
			n_infos.push_back(tempRes);
			tempRes.clear();
		}
		//每组人数减一个领头执法人员
		//随机每组非领头执法人员
		int ncurCount = ncount - 1;
		for (size_t i = 0; i < len; i++)
		{
			sprintf(strSQL, "SELECT *\
								 FROM staffs\
								 WHERE Pair!='%s'\
								 ORDER BY rand() LIMIT %d; ", department_lead.c_str(), ncurCount);
			query1 = cp->query(strSQL);
			std::vector<staffs>rstaffs;
			query1.storein(rstaffs);
			size_t lens = rstaffs.size();
			query1.clear();
			//用于存储抽取执法人员组别信息
			for (size_t j = 0; j < lens; j++) {
				staffs pio = rstaffs[j];
				Json::Value temp;
				temp["ID"] = pio.ID;
				//存储数据所在组号
				temp["Group"] = i + 1;

				std::string str;
				UTF82C(pio.Pair.c_str(), str);
				temp["Pair"] = str;

				UTF82C(pio.Duty.c_str(), str);
				temp["Duty"] = str;

				//存储执法人员姓名，用于导出时使用

				n_infos[i].push_back(pio.Name);

				UTF82C(pio.Name.c_str(), str);
				temp["Name"] = str;

				UTF82C(pio.Remarks.c_str(), str);
				temp["Remarks"] = str;

				UTF82C(pio.Number.c_str(), str);
				temp["Number"] = str;

				jret.append(temp);
			}
		}
		sRet = jret.toStyledString();

		res["str"] = sRet;
		return res;
	}
	cp->disconnect();
	return res;
}


json::value CStochasticServerDlg::SaveToExcel(json::value group)
{
	//获取保存路径
	auto tablename = group.get_item("filepath").get(L"");
	//获取组别号
	int groupNumber= group.get_item("group").get<int>();
	json::value ret;
	//获取市场抽取长度
	size_t length = m_infos.size();
	Json::Value JSInfo;
	USES_CONVERSION;
	libxl::Book* book = xlCreateBook();
	const wchar_t * x = L"Halil Kural";
	const wchar_t * y = L"windows-2723210a07c4e90162b26966a8jcdboe";
	book->setKey(x, y);
	if (book)
	{
		libxl::Sheet* sheet = book->addSheet(L"sheet1");
		if (sheet)
		{
			int rows = 0, columns = 0;
			//写入存储title
			libxl::Font* font = book->addFont();//创建一个字体对象
			font->setColor(COLOR_BLACK);  //设置对象颜色
			font->setBold(true);        //设置粗体
			Format * xmlmat = book->addFormat();
			xmlmat->setFont(font);
			xmlmat->setBorder(BORDERSTYLE_THIN);
			xmlmat->setAlignH(ALIGNH_CENTER);
			sheet->setCol(0,10,30.0);
			sheet->setRow(0, 25.0);
			sheet->setAutoFitArea(0, 0, length, 8);
			sheet->setMerge(0, 0, 0, 7);
			sheet->writeStr(rows++, columns, L"绵阳市西南科技大学抽查结果", xmlmat);
			//写入抽取的为第几组
			CString groupstr;
			//生产流通领域联合抽查（第2组)
			groupstr.Format(L"生产流通领域联合抽查（第%d组)", groupNumber);
			sheet->setMerge(1, 1, 0, 7);
			sheet->setRow(1, 25.0);
			sheet->writeStr(rows++, columns, groupstr, xmlmat);
			//写入执法人员名单
			CString Lgstr;
			size_t mlen=n_infos[groupNumber - 1].size();
			Lgstr.Format(L"执法人员名单(%d)人:", mlen);
			std::string str;
			for (size_t i = 0; i <mlen; i++)
			{
				UTF82C(n_infos[groupNumber - 1][i].c_str(), str);
				Lgstr += A2W(str.c_str());
				Lgstr += L"  ";
			}
			sheet->setRow(2, 25.0);
			sheet->setMerge(2, 2, 0, 7);
			sheet->writeStr(rows++, columns, Lgstr, xmlmat);
			for (size_t i = 0; i < length; i++)
			{
				sheet->setRow(rows, 15.0);
				sheet->writeNum(rows, columns++,m_infos[i].ID, xmlmat);
				UTF82C(m_infos[i].Registermark.c_str(), str);
				sheet->writeStr(rows, columns++,A2W(str.c_str()), xmlmat);
				UTF82C(m_infos[i].WordName.c_str(), str);
				sheet->writeStr(rows, columns++, A2W(str.c_str()), xmlmat);
				UTF82C(m_infos[i].LegalPerson.c_str(), str);
				sheet->writeStr(rows, columns++, A2W(str.c_str()), xmlmat);
				UTF82C(m_infos[i].BusinessPlace.c_str(), str);
				sheet->writeStr(rows, columns++, A2W(str.c_str()), xmlmat);
				UTF82C(m_infos[i].Type.c_str(), str);
				sheet->writeStr(rows, columns++, A2W(str.c_str()), xmlmat);
				UTF82C(m_infos[i].Scopes.c_str(), str);
				sheet->writeStr(rows, columns++, A2W(str.c_str()), xmlmat);
				UTF82C(m_infos[i].Phone.c_str(), str);
				sheet->writeStr(rows, columns++, A2W(str.c_str()), xmlmat);
				rows++;//行数增加
				columns = 0;
			}
			
			//保存文件路径提取
			CString filename = tablename.c_str();
			int pos = filename.ReverseFind(L':');
			filename = filename.Mid(pos - 1);
			book->save(filename);
			book->release();
		}
	}
	return ret;
}


json::value CStochasticServerDlg::DepartmentList()
{
	json::value res;
	std::string sRet;
	USES_CONVERSION;
	mysqlpp::ScopedConnection cp(*g_pMysqlCP, true);

	if (cp) {
		char strSQL[MAXPATH] = { 0 };
		mysqlpp::Query query1 = cp->query(strSQL);
		query1.clear();
		//获取部门表
		sprintf(strSQL, "SELECT * FROM department;");
		query1 = cp->query(strSQL);
		std::vector<department>rdepartments;
		query1.storein(rdepartments);
		query1.clear();
		size_t len = rdepartments.size();
		Json::Value jret;
		for (size_t i = 0; i < len; i++) {
			department pio = rdepartments[i];
			Json::Value temp;
			temp["ID"] = pio.ID;
			std::string str;
			UTF82C(pio.Pair.c_str(), str);
			temp["Pair"] = str;
			jret.append(temp);
		}
		sRet = jret.toStyledString();

		res["str"] = sRet;
		return res;
	}
	cp->disconnect();
	return res;
}

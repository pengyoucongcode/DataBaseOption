// DBOption.cpp: 实现文件
//

#include "stdafx.h"
#include "数据库操作.h"
#include "DBOption.h"
#include "afxdialogex.h"


// CDBOption 对话框

IMPLEMENT_DYNAMIC(CDBOption, CDialogEx)

CDBOption::CDBOption(CWnd* pParent /*=nullptr*/)
	: CDialogEx(IDD_DB, pParent)
	, m_ID(_T(""))
	, m_Name(_T(""))
	, m_Sex(_T(""))
	, m_Age(_T(""))
{

}

CDBOption::~CDBOption()
{
}
BOOL CDBOption::InitInstance()
{
	AfxOleInit();
	return TRUE;
}
void CDBOption::DoDataExchange(CDataExchange* pDX)
{
	CDialogEx::DoDataExchange(pDX);
	DDX_Control(pDX, IDC_EDIT_LOG, m_editLog);
	DDX_Text(pDX, IDC_EDIT_ID, m_ID);
	DDX_Text(pDX, IDC_EDIT_NAME, m_Name);
	DDX_Text(pDX, IDC_EDIT_SEX, m_Sex);
	DDX_Text(pDX, IDC_EDIT_AGE, m_Age);
}


BEGIN_MESSAGE_MAP(CDBOption, CDialogEx)
	ON_BN_CLICKED(IDOK, &CDBOption::OnBnClickedOk)
	ON_BN_CLICKED(IDC_BUTTON_CONNECT, &CDBOption::OnButtonConnect)
	ON_BN_CLICKED(IDC_BUTTON_READ, &CDBOption::OnButtonRead)
	ON_BN_CLICKED(IDC_BUTTON_FIRST, &CDBOption::OnButtonFirst)
	ON_BN_CLICKED(IDC_BUTTON_LAST, &CDBOption::OnButtonLast)
	ON_BN_CLICKED(IDC_BUTTON_NEXT, &CDBOption::OnButtonNext)
	ON_BN_CLICKED(IDC_BUTTON_PREV, &CDBOption::OnButtonPrev)
	ON_BN_CLICKED(IDC_BUTTON_WRITE, &CDBOption::OnButtonWrite)
	ON_BN_CLICKED(IDC_BUTTON_DELETE, &CDBOption::OnButtonDelete)
END_MESSAGE_MAP()


// CDBOption 消息处理程序


void CDBOption::OnBnClickedOk()
{
	// TODO: 在此添加控件通知处理程序代码
	CDialogEx::OnOK();
}


void CDBOption::OnButtonConnect()
{
	// TODO: 在此添加控件通知处理程序代码
	HRESULT hr;
	try
	{
		hr = m_pConnection.CreateInstance("ADODB.Connection");
		if (SUCCEEDED(hr))  //如果ADO实例创建成功
		{
			hr = m_pConnection->Open("driver={SQL Server};Server=御承扬;Database=TEST1;UID=pyc-ycy;PWD=789147;", "", "", adModeUnknown);
		}
		if (m_pConnection->State)
			m_editLog.SetWindowText(L"数据库连接成功");
		else
			m_editLog.SetWindowText(L"数据库连接失败");
	}
	catch (_com_error e)
	{
		CString log;
		log.Format(L"数据库连接失败！\r\n原因：%s", e.ErrorMessage());
		MessageBox(log, L"提示", MB_ICONASTERISK);
	}
}

void CDBOption::GetRecordContent()
{
	if (m_pRecordset == NULL)  //记录集为NULL
		return;
	if (m_pRecordset->BOF || m_pRecordset->adoEOF)  //记录是头或尾
		return;
	_variant_t vID, vName, vSex, vAge;
	vID = m_pRecordset->GetCollect("ID");
	vName = m_pRecordset->GetCollect("Agent_Name");
	vSex = m_pRecordset->GetCollect("Sex");
	vAge = m_pRecordset->GetCollect("Age");
	m_ID = (LPCTSTR)(_bstr_t)(vID);
	m_Name = (LPCTSTR)(_bstr_t)(vName);
	m_Sex = (LPCTSTR)(_bstr_t)(vSex);
	m_Age = (LPCTSTR)(_bstr_t)(vAge);
	UpdateData(false);
}

void CDBOption::OnButtonRead()
{
	// TODO: 在此添加控件通知处理程序代码
	if (m_pConnection == NULL)
	{
		MessageBox(L"数据库还未连接\r\n，请连接后再重复此操作", L"提示", MB_ICONASTERISK);
		return;    //如果还未连接，则返回
	}
	m_pConnection->CursorLocation = adUseServer;  //设置连接使用光标类型
	m_pRecordset.CreateInstance("ADODB.Recordset"); //创建记录对象
	m_pRecordset->Open("SELECT * FROM Agent",_variant_t((IDispatch*)m_pConnection,true),adOpenStatic,adLockBatchOptimistic,adCmdText);
	GetRecordContent();
}


void CDBOption::OnButtonFirst()
{
	// TODO: 在此添加控件通知处理程序代码
	if (m_pRecordset == NULL)  //如果记录集为空
		return;
	m_pRecordset->MoveFirst();  //移动到记录集里的第一条记录
	GetRecordContent();
}


void CDBOption::OnButtonLast()
{
	// TODO: 在此添加控件通知处理程序代码
	if (m_pRecordset == NULL)  //如果记录集为空
		return;
	m_pRecordset->MoveLast();  //移动到记录集里的最后一条记录
	GetRecordContent();
}


void CDBOption::OnButtonNext()
{
	// TODO: 在此添加控件通知处理程序代码
	if (m_pRecordset == NULL)  //如果记录集为空
		return;
	if (m_pRecordset->adoEOF)
	{
		MessageBox(L"已经是最后一条记录！", L"提示", MB_ICONASTERISK);
		return;  //如果是最后一条记录则返回
	}
	m_pRecordset->MoveNext(); //向后移动一条记录
	GetRecordContent();
}


void CDBOption::OnButtonPrev()
{
	// TODO: 在此添加控件通知处理程序代码
	if (m_pRecordset == NULL)  //如果记录集为空
		return;
	if (m_pRecordset->BOF)
	{
		MessageBox(L"已经是第一条记录！", L"提示", MB_ICONASTERISK);
		return;  //如果第一条记录则返回
	}
	m_pRecordset->MovePrevious();//向前移动一条记录
	GetRecordContent();
}


void CDBOption::OnButtonWrite()
{
	// TODO: 在此添加控件通知处理程序代码
	if (m_pConnection == NULL || m_pRecordset == NULL)
	{
		MessageBox(L"数据库未连接\r\n或数据集不存在", L"提示", MB_ICONASTERISK);
		return;
	}
	UpdateData(true);
	CString strSql;
	strSql.Format(L"INSERT INTO Agent(ID, Agent_Name, Sex, Age) VALUES(%d, '%s', '%s', %d)", _ttoi(m_ID),m_Name,m_Sex,_ttoi(m_Age));
	try
	{

		m_pConnection->Execute(_bstr_t(strSql), 0, adCmdText);
		m_editLog.SetWindowText(L"写入成功！");
	}
	catch (_com_error e)
	{
		MessageBox(e.Description());
		return;
	}
}


void CDBOption::OnButtonDelete()
{
	// TODO: 在此添加控件通知处理程序代码
	if (m_pRecordset == NULL)
	{
		MessageBox(L"记录集不存在\r\n请先执行读取操作", L"提示", MB_ICONASTERISK);
		return;
	}
	try
	{
		UpdateData(true);
		bool key = false;
		_variant_t vName;
		m_pRecordset->Requery(NULL);
		if (!m_pRecordset->BOF)
			m_pRecordset->MoveFirst();
		while (true)
		{
			vName = m_pRecordset->GetCollect("Agent_Name");
			if (m_Name == (LPCTSTR)(_bstr_t)(vName))
			{
				key = true;
				break;
			}
			m_pRecordset->MoveNext();
			if (m_pRecordset->adoEOF)
			{
				if (m_Name == (LPCTSTR)(_bstr_t)(vName))
				{
					key = true;
					break;
				}
				else
				{
					MessageBox(L"该客户不存在\r\n无法进行删除\r\n请确认是否输入错误", L"提示", MB_ICONASTERISK);
					break;
				}
			}

		}
		if (key)
		{
			m_pRecordset->Delete(adAffectCurrent);
			m_pRecordset->Update();
			CString sql;
			m_pRecordset->Requery(NULL);
			GetRecordContent();
			m_editLog.SetWindowText(L"删除成功");
		}
	}
	catch (_com_error e)
	{
		CString log;
		log.Format(L"删除失败！\r\n原因：%s\r\n%s", e.ErrorMessage(), (LPCTSTR)e.Description());
		MessageBox(log, L"提示", MB_ICONASTERISK);
	}
}

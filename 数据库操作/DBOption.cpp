// DBOption.cpp: 实现文件
//

#include "stdafx.h"
#include "数据库操作.h"
#include "DBOption.h"
#include "afxdialogex.h"
#include "Excel.h"

using namespace std;

// CDBOption 对话框

IMPLEMENT_DYNAMIC(CDBOption, CDialogEx)

CDBOption::CDBOption(CWnd* pParent /*=nullptr*/)
	: CDialogEx(IDD_DB, pParent)
	, m_ID(_T(""))
	, m_Name(_T(""))
	, m_Sex(_T(""))
	, m_Age(_T(""))
	, m_strImportFilePath(_T(""))
{

}

CDBOption::~CDBOption()
{
}

void CDBOption::DoDataExchange(CDataExchange* pDX)
{
	CDialogEx::DoDataExchange(pDX);
	DDX_Control(pDX, IDC_EDIT_LOG, m_editLog);
	DDX_Text(pDX, IDC_EDIT_ID, m_ID);
	DDX_Text(pDX, IDC_EDIT_NAME, m_Name);
	DDX_Text(pDX, IDC_EDIT_SEX, m_Sex);
	DDX_Text(pDX, IDC_EDIT_AGE, m_Age);
	DDX_Control(pDX, IDC_LIST1, m_list);
	DDX_Text(pDX, IDC_EDIT_IMPORTPATH, m_strImportFilePath);
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
	ON_BN_CLICKED(IDC_BUTTON_CHANGE, &CDBOption::OnButtonChange)
	ON_BN_CLICKED(IDC_BUTTON_LISTREAD, &CDBOption::OnButtonListread)
	ON_BN_CLICKED(IDC_BUTTON_INPATH, &CDBOption::OnButtonInpath)
	ON_BN_CLICKED(IDC_BUTTON_READFILE, &CDBOption::OnButtonReadfile)
END_MESSAGE_MAP()


// CDBOption 消息处理程序
BOOL CDBOption::InitInstance()
{
	AfxOleInit();
	return TRUE;
}

BOOL CDBOption::OnInitDialog()
{
	CDialogEx::OnInitDialog();
	// 初始化列表控件
	{
		LONG lStyle = GetWindowLong(m_list.m_hWnd, GWL_STYLE);
		lStyle &= ~LVS_TYPEMASK;
		lStyle |= LVS_REPORT;
		SetWindowLong(m_list.GetSafeHwnd(), GWL_STYLE, lStyle);
		DWORD dwStyle = m_list.GetExtendedStyle();
		dwStyle |= LVS_EX_FULLROWSELECT; // 选中行，整行高亮
		dwStyle |= LVS_EX_GRIDLINES; // 网格线
		m_list.SetExtendedStyle(dwStyle);
		// 设置列，并设置大小
		{
			CRect rc;
			m_list.GetClientRect(rc);
			m_list.InsertColumn(EListIndexCode, _T("客户编号"), LVCFMT_LEFT, rc.Width() / EListIndexMaxLimit);
			m_list.InsertColumn(EListIndexName, _T("姓名"), LVCFMT_LEFT, rc.Width() / EListIndexMaxLimit);
			m_list.InsertColumn(EListIndexSex, _T("性别"), LVCFMT_LEFT, rc.Width() / EListIndexMaxLimit);
			m_list.InsertColumn(EListIndexAge, _T("年龄"), LVCFMT_LEFT, rc.Width() / EListIndexMaxLimit);
		}
	}
	return TRUE;
}

void CDBOption::OnBnClickedOk()
{
	// TODO: 在此添加控件通知处理程序代码
	CDialogEx::OnOK();
}

// 连接数据库
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
		log.Format(L"数据库连接失败！\r\n原因：%s\r\n", e.ErrorMessage(),e.Description());
		MessageBox(log, L"提示", MB_ICONASTERISK);
	}
}
// 显示记录
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
//读取记录
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

// 读取第一条记录
void CDBOption::OnButtonFirst()
{
	// TODO: 在此添加控件通知处理程序代码
	if (m_pRecordset == NULL)  //如果记录集为空
		return;
	m_pRecordset->MoveFirst();  //移动到记录集里的第一条记录
	GetRecordContent();
}

// 读取最后一条记录
void CDBOption::OnButtonLast()
{
	// TODO: 在此添加控件通知处理程序代码
	if (m_pRecordset == NULL)  //如果记录集为空
		return;
	m_pRecordset->MoveLast();  //移动到记录集里的最后一条记录
	GetRecordContent();
}

// 读取下一条记录
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

// 读取上一条记录
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

//添加记录
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

// 删除记录
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
		if (key==true)
		{
			CString sql;
			sql.Format(L"DELETE FROM Agent WHERE ID=%d",_ttoi(m_ID ));
			m_pConnection->Execute(_bstr_t(sql), 0, adCmdText);
			/*m_pRecordset->Delete(adAffectCurrent);
			m_pRecordset->Update();*/ 
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

// 修改数据记录
void CDBOption::OnButtonChange()
{
	// TODO: 在此添加控件通知处理程序代码
	bool key = false;
	try
	{
		UpdateData(true);
		_variant_t vID, vName, vSex, vAge;
		m_pRecordset->Requery(NULL);
		if (!m_pRecordset->BOF)
			m_pRecordset->MoveFirst();
		while (true)
		{
			vID = m_pRecordset->GetCollect("ID");
			if (m_ID == (LPCTSTR)(_bstr_t)(vID))
			{
				key = true;
				break;
			}
			m_pRecordset->MoveNext();
			if (m_pRecordset->adoEOF)
			{
				if (m_ID == (LPCTSTR)(_bstr_t)(vID))
				{
					key = true;
					break;
				}
				else
				{
					MessageBox(L"該编号的客户并不存在\r\n无法进行修改\r\n请检查是否有误", L"提示", MB_ICONASTERISK);
					break;
				}
			}
		}
		if (key == true)
		{
			UpdateData(true);
			CString sql,sql1,sql2;
			sql.Format(L"UPDATE Agent SET Agent_Name='"+ m_Name+L"' WHERE ID=%d", _ttoi(m_ID));
			sql1.Format(L"UPDATE Agent SET Sex='%s' WHERE ID=%d", m_Sex, _ttoi(m_ID));
			sql2.Format(L"UPDATE Agent SET Age=%d WHERE ID=%d", _ttoi(m_Age), _ttoi(m_ID));
			m_pConnection->Execute(_bstr_t(sql), 0, adCmdText);
			m_pConnection->Execute(_bstr_t(sql1), 0, adCmdText);
			m_pConnection->Execute(_bstr_t(sql2), 0, adCmdText);
			m_pRecordset->Requery(NULL);
			m_pRecordset->MoveLast();
			GetRecordContent();
			m_editLog.SetWindowText(L"修改成功");
		}
	}
	catch (_com_error e)
	{
		CString log;
		log.Format(L"修改失败！\r\n原因：%s\r\n%s", e.ErrorMessage(), (LPCTSTR)e.Description());
		MessageBox(log, L"提示", MB_ICONASTERISK);
	}
}

void CDBOption::UpdateListData()
{
	_variant_t vID, vName, vSex, vAge;
	vID = m_pRecordset->GetCollect("ID");
	vName = m_pRecordset->GetCollect("Agent_Name");
	vSex = m_pRecordset->GetCollect("Sex");
	vAge = m_pRecordset->GetCollect("Age");
	CString ID, Name, Sex, Age;
	ID = (LPCTSTR)(_bstr_t)(vID);
	Name = (LPCTSTR)(_bstr_t)(vName);
	Sex = (LPCTSTR)(_bstr_t)(vSex);
	Age = (LPCTSTR)(_bstr_t)(vAge);
	int index = m_list.GetItemCount();
	m_list.InsertItem(index, ID);
	m_list.SetItemText(index, EListIndexName, Name);
	m_list.SetItemText(index, EListIndexSex, Sex);
	m_list.SetItemText(index, EListIndexAge, Age);
	
}

void CDBOption::OnButtonListread()
{
	// TODO: 在此添加控件通知处理程序代码
	if (m_pRecordset == NULL)
	{
		MessageBox(L"记录集不存在\r\n请先执行读取操作", L"提示", MB_ICONASTERISK);
		return;
	}
	try
	{
		if (!m_pRecordset->BOF)
			m_pRecordset->MoveFirst();
		m_list.DeleteAllItems();
		while (true)
		{
			UpdateListData();
			m_pRecordset->MoveNext();
			if (m_pRecordset->adoEOF)
			{
				break;
			}
		}
		m_editLog.SetWindowText(L"读取成功，所有记录如表中所示！");
	}
	catch (_com_error e)
	{
		CString log;
		log.Format(L"批量读取失败！\r\n原因：%s\r\n%s", e.ErrorMessage(), (LPCTSTR)e.Description());
		MessageBox(log, L"提示", MB_ICONASTERISK);
	}
}


void CDBOption::OnButtonInpath()
{
	// TODO: 在此添加控件通知处理程序代码
	UpdateData(TRUE);
	/* 导入文件为 UTF8 无BOM 编码
	   格式是:
	 */
	using namespace std;
	static CString strFilename = _T("");
	static TCHAR szFilter[] = _T("文本l文件(*.xls)|*.xls|所有文件(*.*)|*.*||");
	CFileDialog dlg(TRUE, // 创建文件打开对话框:FALSE保存对话框　
		_T(".txt"),
		strFilename.GetString(),
		OFN_HIDEREADONLY | OFN_OVERWRITEPROMPT,
		szFilter
	);
	if (IDOK != dlg.DoModal()) {
		return;
	}
	// 显示选择的文件内容
	m_strImportFilePath = dlg.GetPathName();
	UpdateData(FALSE);
}

void CDBOption::OnButtonReadfile()
{
	if (m_pConnection == NULL)
	{
		MessageBox(L"数据库未连接", L"提示", MB_ICONASTERISK);
		return;
	}
	UpdateData(TRUE);
	m_list.DeleteAllItems();
	Excel excel;
	bool bInit = excel.initExcel();
	// CString 转 const char*
	USES_CONVERSION;
	std::string s(W2A(m_strImportFilePath));
	const char* cstr = s.c_str();
	bool bRect = excel.open(cstr); // 打开 Excel 文件
	if (bRect == false)
	{
		m_editLog.SetWindowText(L"文件打开失败！");
		return;
	}
	if (bRect == true)
	{
		CString strsheetName = excel.getSheetName(1); // 获取sheet名
		bool bLoad = excel.loadSheet(strsheetName);
		if (bLoad)
		{
			int nRow = excel.getRowCount(); // 获取sheet中的行数
			int ID, Age;
			CString  Name, Sex;
			for (int i = 2; i <= nRow; i++)
			{
				ID = excel.getCellInt(i, 1);
				Name = excel.getCellString(i, 2);
				Sex = excel.getCellString(i, 3);
				Age = excel.getCellInt(i, 4);
				CString strSql;
				strSql.Format(L"INSERT INTO Agent(ID, Agent_Name, Sex, Age) VALUES(%d, '%s', '%s', %d)", ID, Name, Sex, Age);
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
			m_editLog.SetWindowText(L"导入成功！");
		}
		else
		{
			m_editLog.SetWindowText(L"导入失败！");
			return;
		}
	}
}

#pragma once
#import "C:\Program Files\Common Files\System\ado\msado15.dll" \
        no_namespace rename("EOF", "adoEOF")

// CDBOption 对话框

class CDBOption : public CDialogEx
{
	DECLARE_DYNAMIC(CDBOption)

public:
	CDBOption(CWnd* pParent = nullptr);   // 标准构造函数
	virtual ~CDBOption();
	BOOL InitInstance();
	void GetRecordContent();
	BOOL OnInitDialog();

	_ConnectionPtr m_pConnection;
	_RecordsetPtr m_pRecordset;

// 对话框数据
#ifdef AFX_DESIGN_TIME
	enum { IDD = IDD_DB };
#endif

protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV 支持

	DECLARE_MESSAGE_MAP()
public:
	afx_msg void OnBnClickedOk();
	afx_msg void OnButtonConnect();
	// 状态信息
	CEdit m_editLog;
	// 客户编号
	CString m_ID;
	// 客户姓名
	CString m_Name;
	// 客户性别
	CString m_Sex;
	// 客户年龄
	CString m_Age;
	afx_msg void OnButtonRead();
	afx_msg void OnButtonFirst();
	afx_msg void OnButtonLast();
	afx_msg void OnButtonNext();
	afx_msg void OnButtonPrev();
	afx_msg void OnButtonWrite();
	afx_msg void OnButtonDelete();
	afx_msg void OnButtonChange();
	CListCtrl m_list;
	enum {
		EListIndexCode = 0,
		EListIndexName,
		EListIndexSex,
		EListIndexAge,
		EListIndexMaxLimit
	};
	afx_msg void OnButtonListread();
	void UpdateListData();
	// 获取导入路径的编辑框变量
	CString m_strImportFilePath;
	afx_msg void OnButtonInpath();
	afx_msg void OnButtonReadfile();
};

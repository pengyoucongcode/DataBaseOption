
// 数据库操作.h: 数据库操作 应用程序的主头文件
//
#pragma once

#ifndef __AFXWIN_H__
	#error "在包含此文件之前包含“stdafx.h”以生成 PCH 文件"
#endif

#include "resource.h"       // 主符号
#include "DBOption.h"
#include<vector>

// C数据库操作App:
// 有关此类的实现，请参阅 数据库操作.cpp
//

class C数据库操作App : public CWinApp
{
public:
	C数据库操作App() noexcept;


// 重写
public:
	virtual BOOL InitInstance();
	virtual int ExitInstance();

// 实现
private:
	// 存讲所有对话框的指针
	

	afx_msg void OnAppAbout();
	afx_msg void OnDBOption();

	DECLARE_MESSAGE_MAP()
};

extern C数据库操作App theApp;

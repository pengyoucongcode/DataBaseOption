
// 数据库操作View.cpp: C数据库操作View 类的实现
//

#include "stdafx.h"
// SHARED_HANDLERS 可以在实现预览、缩略图和搜索筛选器句柄的
// ATL 项目中进行定义，并允许与该项目共享文档代码。
#ifndef SHARED_HANDLERS
#include "数据库操作.h"
#endif

#include "数据库操作Doc.h"
#include "数据库操作View.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#endif


// C数据库操作View

IMPLEMENT_DYNCREATE(C数据库操作View, CView)

BEGIN_MESSAGE_MAP(C数据库操作View, CView)
	// 标准打印命令
	ON_COMMAND(ID_FILE_PRINT, &CView::OnFilePrint)
	ON_COMMAND(ID_FILE_PRINT_DIRECT, &CView::OnFilePrint)
	ON_COMMAND(ID_FILE_PRINT_PREVIEW, &CView::OnFilePrintPreview)
END_MESSAGE_MAP()

// C数据库操作View 构造/析构

C数据库操作View::C数据库操作View() noexcept
{
	// TODO: 在此处添加构造代码

}

C数据库操作View::~C数据库操作View()
{
}

BOOL C数据库操作View::PreCreateWindow(CREATESTRUCT& cs)
{
	// TODO: 在此处通过修改
	//  CREATESTRUCT cs 来修改窗口类或样式

	return CView::PreCreateWindow(cs);
}

// C数据库操作View 绘图

void C数据库操作View::OnDraw(CDC* /*pDC*/)
{
	C数据库操作Doc* pDoc = GetDocument();
	ASSERT_VALID(pDoc);
	if (!pDoc)
		return;

	// TODO: 在此处为本机数据添加绘制代码
}


// C数据库操作View 打印

BOOL C数据库操作View::OnPreparePrinting(CPrintInfo* pInfo)
{
	// 默认准备
	return DoPreparePrinting(pInfo);
}

void C数据库操作View::OnBeginPrinting(CDC* /*pDC*/, CPrintInfo* /*pInfo*/)
{
	// TODO: 添加额外的打印前进行的初始化过程
}

void C数据库操作View::OnEndPrinting(CDC* /*pDC*/, CPrintInfo* /*pInfo*/)
{
	// TODO: 添加打印后进行的清理过程
}


// C数据库操作View 诊断

#ifdef _DEBUG
void C数据库操作View::AssertValid() const
{
	CView::AssertValid();
}

void C数据库操作View::Dump(CDumpContext& dc) const
{
	CView::Dump(dc);
}

C数据库操作Doc* C数据库操作View::GetDocument() const // 非调试版本是内联的
{
	ASSERT(m_pDocument->IsKindOf(RUNTIME_CLASS(C数据库操作Doc)));
	return (C数据库操作Doc*)m_pDocument;
}
#endif //_DEBUG


// C数据库操作View 消息处理程序

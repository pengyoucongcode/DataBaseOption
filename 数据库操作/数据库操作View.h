
// 数据库操作View.h: C数据库操作View 类的接口
//

#pragma once


class C数据库操作View : public CView
{
protected: // 仅从序列化创建
	C数据库操作View() noexcept;
	DECLARE_DYNCREATE(C数据库操作View)

// 特性
public:
	C数据库操作Doc* GetDocument() const;

// 操作
public:

// 重写
public:
	virtual void OnDraw(CDC* pDC);  // 重写以绘制该视图
	virtual BOOL PreCreateWindow(CREATESTRUCT& cs);
protected:
	virtual BOOL OnPreparePrinting(CPrintInfo* pInfo);
	virtual void OnBeginPrinting(CDC* pDC, CPrintInfo* pInfo);
	virtual void OnEndPrinting(CDC* pDC, CPrintInfo* pInfo);

// 实现
public:
	virtual ~C数据库操作View();
#ifdef _DEBUG
	virtual void AssertValid() const;
	virtual void Dump(CDumpContext& dc) const;
#endif

protected:

// 生成的消息映射函数
protected:
	DECLARE_MESSAGE_MAP()
};

#ifndef _DEBUG  // 数据库操作View.cpp 中的调试版本
inline C数据库操作Doc* C数据库操作View::GetDocument() const
   { return reinterpret_cast<C数据库操作Doc*>(m_pDocument); }
#endif


// PageCache.h: interface for the PageCache class.
//
//////////////////////////////////////////////////////////////////////

#if !defined(AFX_PAGECACHE_H__97201EB6_5393_40A8_933E_436A90D42D89__INCLUDED_)
#define AFX_PAGECACHE_H__97201EB6_5393_40A8_933E_436A90D42D89__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000

#include <stdio.h>

//这个类提供路面灰度位图的缓存功能，由XDrawImage使用，来表现路面影像
//每个要缓存的位图大小必须一致，由构造函数的nWidth和nHeight指定
struct PtCrack;
class PageCache  
{
	struct Page
	{
		long		m_nPage;
		long		m_nScale;
		BYTE*		m_pBits;
		int dis(int nPage){ nPage-=m_nPage; return nPage>0?nPage:-nPage; }
		BYTE* Init(int i,BYTE* pBit,int size);
	};

    char                        m_root[512];
    char*                       m_fnA[1024];
    char                        m_filename[512];
    
	BYTE*						m_pMem;
	Page						m_Pages[128];
	Page*						m_pPages[128];
    FILE*                       m_fb4;
	long                        m_bFast;

	Page* AllocPage(int nPage,int nScale);
    const char* FileAt(int iPage);

	PageCache(PageCache&);
	void operator=(PageCache&);

public:
	long			        m_nWidth;       //单幅图片的宽度和
    long			        m_nHeight;      //高度
    long                    m_nBmpCount;    //图片的张数
    
	PageCache(const char* Path);
	~PageCache();

    void OpenFb4(const char* path);

    //功能：根据页面号得到页面的颜色矩阵
	const BYTE* GetPage(int nPage,int nScale=1);

    //功能：得到一行的颜色
	const BYTE* GetRow(int iY,int iScale);

    //描述：清除所有装入内存中的数据
    void ClearMemcache();
    
    void SetFast(BOOL bFast,double scale);
    void DrawImage(HDC hDC,POINT pt,int w,int h,double fScale);
    void GetBits(int x,int y,int w,int h,BYTE* pBits);
    void GetBitsEx(double x,double y,double step,int w,int h,BYTE* pBits);
    void GetBitsEx4(double x,double y,double step,int w,int h,BYTE* pBits);
};

#endif // !defined(AFX_PAGECACHE_H__97201EB6_5393_40A8_933E_436A90D42D89__INCLUDED_)

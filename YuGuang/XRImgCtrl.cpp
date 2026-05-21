#include "stdafx.h"
#include "PageCache.h"

PageCache*  m_page = 0;

extern "C" __declspec(dllexport) void BH_Open(const char* path,int* w,int* h)
{
    if(m_page!=NULL)
        delete m_page;

    m_page = NULL; *w = *h = 0;
    if(path==NULL || *path==NULL)
        return;
    m_page = new PageCache(path);
    if(m_page->m_nBmpCount==0)
    {
        delete m_page;
        m_page = NULL;
        return;
    }
    *w = m_page->m_nWidth;
    *h = m_page->m_nHeight*m_page->m_nBmpCount;
}

extern "C" __declspec(dllexport) void BH_DrawMain(HWND hmain,int y,int fast)
{
    if(m_page==NULL || hmain==NULL) return;

    RECT rc; double scale; HDC hdc; 
    POINT pt = {0,y};

    ::GetClientRect(hmain,&rc);
    scale = (double)rc.right/m_page->m_nWidth;
    
    hdc = GetDC(hmain);
    m_page->SetFast(fast,scale);
    m_page->DrawImage(hdc,pt,rc.right,rc.bottom,scale);
    
    ReleaseDC(hmain,hdc);
}

//描述: 显示放大图片
//参数: (x,y)鼠标在图像中的位置
static void _DrawLage(HWND hlarge,int x,int y)
{
    POINT pt; RECT rc; HDC hdc; 
    if( !::GetClientRect(hlarge,&rc) )
		return;
    hdc = GetDC(hlarge);
    pt.x = x-rc.right/2;
    pt.y = y-rc.bottom/2;
    if(pt.x<0) pt.x = 0;
    if(pt.y<0) pt.y = 0;
    if(pt.x+rc.right>m_page->m_nWidth)
        pt.x = m_page->m_nWidth-rc.right;

    m_page->SetFast(false,1);
    m_page->DrawImage(hdc,pt,rc.right,rc.bottom,1);
    
    //显示鼠标
    int n,A[] = {0,0,0,14,3,11,6,16,8,15,6,10,10,10,0,0};    
    n = sizeof(A)/8;
    x -= pt.x; y -= pt.y;
    while(--n>=0)
    {
        A[2*n] += x;
        A[2*n+1] += y;
    }
    n = sizeof(A)/8;
    ::Polygon(hdc,(POINT*)A,n);
    
    ReleaseDC(hlarge,hdc);
}

extern "C" __declspec(dllexport) void BH_DrawLage(HWND hmain,HWND hlarge,double cy)
{
    if(m_page==NULL || hmain==NULL || hlarge==NULL)
        return;

    POINT pt; RECT rc; double s;

    ::GetClientRect(hmain,&rc);
    if( !::GetCursorPos(&pt) )
        return;
    if( !::ScreenToClient(hmain,&pt) )
        return;
    if(pt.x<0 || pt.x>=rc.right)
        return;
    if(pt.y<0 || pt.y>=rc.bottom)
        return;
    
    s = (double)rc.right/m_page->m_nWidth;
    cy += (pt.y-rc.bottom*0.5)/s;

    _DrawLage(hlarge,(int)(pt.x/s+0.5),(int)(cy+0.5));
}

extern "C" __declspec(dllexport) BYTE* BH_GetBmpPtr(int y)
{
    if(m_page==NULL)
        return NULL;
    return (BYTE*)m_page->GetPage(y/2048);
}

extern "C" __declspec(dllexport) void BH_ClearImgCache()
{
    if(m_page!=NULL)
        m_page->ClearMemcache();
}
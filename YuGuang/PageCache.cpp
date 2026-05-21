// PageCache.cpp: implementation of the PageCache class.
//
//////////////////////////////////////////////////////////////////////

#include "stdafx.h"
#include "PageCache.h"
#include <math.h>
#include <algorithm>
#include <stdio.h>

//////////////////////////////////////////////////////////////////////
// Construction/Destruction
//////////////////////////////////////////////////////////////////////
BOOL ReadJpgGray(const char* pFileName,BYTE* pBits);
BOOL GetJpgSize(const char* pFileName,long& W,long& H);

extern "C" void YG_SetView(int x,int w,double fScale);
extern "C" void YG_Adjust(int w,int h,BYTE* bits);
extern "C" void YG_Ruihua(int w,int h,BYTE* bits);

static char** SearchJpg(const char* path,char** fnA)
{
    WIN32_FIND_DATAA D;

    static char strpool[1024*256];
    char* it = strpool;

	memset(&D,0,sizeof(D));
	strcpy(D.cFileName,path);
	strcat(D.cFileName,"*.jpg");

	HANDLE h = FindFirstFileA(D.cFileName,&D);
	if( h==NULL || h==INVALID_HANDLE_VALUE )
	{
		strcpy(D.cFileName+strlen(D.cFileName)-4,"*.bmp");
		h = FindFirstFileA(D.cFileName,&D);
		if( h==NULL || h==INVALID_HANDLE_VALUE )
		{
			strcpy(D.cFileName+strlen(D.cFileName)-4,"*.jpeg");
			h = FindFirstFileA(D.cFileName,&D);
			if( h==NULL || h==INVALID_HANDLE_VALUE )
				return fnA;
		}
	}
    do{
        if( D.cFileName[0]=='.' )
            continue;
        *fnA++ = strcpy(it,D.cFileName);
        it += strlen(it)+1;
        
	}while( FindNextFileA(h,&D) );
    FindClose(h);
    return fnA;
}

PageCache::PageCache(const char* Path)
{
    memset(this,0,sizeof(*this));
    strcpy(m_root,Path);
    if( m_root[strlen(Path)-1]!='\\' )
        strcat(m_root,"\\");
    m_nBmpCount = SearchJpg(m_root,m_fnA)-m_fnA;
    if(m_nBmpCount==0)
    {
        OpenFb4(Path);
        if( m_fb4==0 )
            return;
        fseek(m_fb4,18,SEEK_SET);
        fread(&m_nWidth,4,2,m_fb4);
        m_nWidth *= 4;
        m_nBmpCount = m_nHeight/512;
        m_nHeight = 2048;
        m_bFast = true;
        return;
    }
	
	if( !GetJpgSize(FileAt(0),m_nWidth,m_nHeight) )
        return;
	
	m_pMem = new BYTE[m_nWidth*m_nHeight*4];
	int i = sizeof(m_Pages)/sizeof(Page)-1;
	for(BYTE* pB=m_pMem; --i>=0; )
	{
		pB = m_Pages[i].Init(i,pB,m_nWidth*m_nHeight);
	}
	
	Page* p = m_Pages;
	m_pPages[1] = p;
	while( (++p)->m_pBits!=0 )
	{
		if(p[-1].m_nScale!=p[0].m_nScale)
			m_pPages[p->m_nScale] = p;
	}

    OpenFb4(Path);

}

PageCache::~PageCache()
{
	if(m_pMem!=0)
		delete m_pMem;
    if(m_fb4!=NULL)
        fclose(m_fb4);
}

void PageCache::OpenFb4(const char* path)
{
    char buf[512],*it;
    if(m_fb4!=NULL)
        fclose(m_fb4);
    m_fb4 = NULL;
    if(path==NULL)
        return;
    it = (char*)strrchr(path,'\\')+1;
    sprintf(buf,"%s\\%s.b4",path,it);    
    m_fb4 = fopen(buf,"rb");
}

const char* PageCache::FileAt(int i)
{
    if(i<0 || i>=m_nBmpCount)
        return NULL;
    strcpy(m_filename,m_root);
    strcat(m_filename,m_fnA[i]);
    return m_filename;
}

BYTE* PageCache::Page::Init(int i,BYTE* pBit,int size)
{
	m_nPage = -1;
	m_pBits = pBit;
	m_nScale = 1;
	i = (i/2)+1;
	while(i!=1)
	{
		m_nScale *= 2;
		i /= 2;
	}
	return pBit+size/m_nScale/m_nScale;
}

//描述：清除所有装入内存中的数据
void PageCache::ClearMemcache()
{
    int i = sizeof(m_Pages)/sizeof(Page);
    while(--i>=0)
    {
        m_Pages[i].m_nPage = -1;
    }
}

PageCache::Page* PageCache::AllocPage(int nPage,int nScale)
{
	Page *A=0,*P;
	if( (P=m_pPages[nScale])==0 )
		return NULL;
	for(;P->m_nScale==nScale; ++P)
	{
		if(P->m_pBits==NULL)
			return NULL;
		if(P->m_nPage==-1)
			return P;
		if(A==0 || P->dis(nPage) > A->dis(nPage) )
			A = P;
	}
	return A->m_pBits?A:0;
}

static void sampling(int n,int s,const BYTE* src,BYTE* dec)
{
    if(s==2)
    {//平均法
        unsigned int c=0;
        for(;--n>=0; src+=2)
        {
            /*取平均值
            c = src[0];
            c += src[1];
		    *dec++ = c/2;
            */
            *dec++ = src[0]<src[1]?src[0]:src[1];
        }
    }
    else
    {
        //采样法
        for(;--n>=0; src+=s)
		    *dec++ = *src;
    }
}

const BYTE* PageCache::GetPage(int nPage,int nScale)
{
    if(nScale>=sizeof(m_pPages)/sizeof(*m_pPages))
        return NULL;
	Page* p = m_pPages[nScale];
	if(p==NULL)
		return NULL;
	for(;p->m_nScale==nScale; ++p)
	{
		if(p->m_nPage==nPage)
			return p->m_pBits;
	}
	p = AllocPage(nPage,nScale);
	if(p==NULL)
		return NULL;
    if( nScale==1 )
    {
        if(nPage>=m_nBmpCount)
            return NULL;
		const char* pFile = FileAt(nPage);
        ReadJpgGray(pFile,p->m_pBits);
        YG_Ruihua(m_nWidth,m_nHeight,p->m_pBits);
		
		p->m_nPage = nPage;
		return p->m_pBits;
	}

	const BYTE* B = GetPage(nPage,nScale/2);
	if(B==NULL)
		return NULL;
	p->m_nPage = nPage;
	BYTE* pB = p->m_pBits;

    int R = 2*(m_nWidth/(nScale/2));
	int W = m_nWidth/nScale;
	int H = m_nHeight/nScale;
	for(; --H>=0; B+=R,pB+=W)
	{
		sampling(W,2,B,pB);
	}

	return p->m_pBits;
}

const BYTE* PageCache::GetRow(int iY,int iScale)
{
	int iPage = iY/m_nHeight;
	iY = iY%m_nHeight;
	const BYTE* pPage;
	pPage = GetPage(iPage,iScale);
	if(pPage!=0)
	{
		iY /= iScale;
		iY *= m_nWidth/iScale;
		return pPage+iY;
	}

	int Scale = iScale;
	while( (Scale/=2)>0 )
	{
		pPage = GetPage(iPage,Scale);
		if(pPage!=NULL)
			break;
	}
	if(pPage==NULL)
		return NULL;
	static BYTE bits[4096];
	const BYTE* B = pPage+(iY/Scale)*m_nWidth/Scale;
	sampling(m_nWidth/iScale,iScale/Scale,B,bits);
	return bits;
}

void ReverseBitsY(int w,int h,BYTE* p)
{
    static BYTE buf[4096*4];
    BYTE* q = p+w*(h-1);
    for(h/=2; --h>=0; p+=w,q-=w)
    {
        memcpy(buf,p,w);
        memcpy(p,q,w);
        memcpy(q,buf,w);
    }
}

//功能：取得指定矩形区域各个象素的灰度值
//参数x,y：矩形的最小x,y
//参数w,h：要取的数据的列数和行数
//参数pBits：返回的灰度值的内存块，函数的调用者必须准备足够多(不少于w*h)的空间
//			 [pBits[0]--pBits[w])是第一行数据,[pBits[w]--pBits[2*w])是第二行数据,……
//			 第一行的Y值是y,第二行的Y值是y+1,第三行的Y值是y+2,……
void PageCache::GetBits(int x,int y,int w,int h,BYTE* pBits)
{
	long nPageHeight,nWidth;
	nPageHeight = m_nHeight;
	nWidth = m_nWidth;

	if( (y%nPageHeight)+h > nPageHeight )
	{
		int n = nPageHeight-(y%nPageHeight);
		GetBits(x,y+n,w,h-n,pBits+w*n);
		h = n;
	}

	memset(pBits,-1,w*h);
	const BYTE* pPage = GetPage(y/nPageHeight);
	if(pPage==NULL)
		return;
	y = y%nPageHeight;
	const BYTE* p = pPage+y*nWidth+x; 
	int w1 = w;
	if(x+w>nWidth) w1 = nWidth-x;
    for(;--h>=0; pBits+=w,p+=nWidth)
    {
        memcpy(pBits,p,w1);
    }
}

//功能：取得指定矩形区域各个象素的灰度值
//参数x,y：矩形的最小x,y
//参数w,h：要取的数据的列数和行数
//参数pBits：返回的灰度值的内存块，函数的调用者必须准备足够多(不少于w*h)的空间
//			 [pBits[0]--pBits[w])是第一行数据,[pBits[w]--pBits[2*w])是第二行数据,……
//			 第一行的Y值是y,第二行的Y值是y+step,第三行的Y值是y+2*step,……
//			 第一列的X值是x,第二列的X值是x+step,第三列的X值是x+2*step,……
void PageCache::GetBitsEx(double x,double y,double step,int w,int h,BYTE* pBits)
{
	static long I[4096]; long s=1,W,i,*pI;
	while( s*1.6<step ) s*= 2;
	
	//得到每行的下标
	W = m_nWidth;
	for(W/=s,i=w; --i>=0;)
	{
		I[i] = (int)((x+i*step)/s+0.5);
		if(I[i]<0 || I[i]>=W) I[i] = -1;
	}

	//一行一行的取数据
	memset(pBits,-1,w*h);
	for(;--h>=0; y+=step)
	{
		const BYTE* R = GetRow((int)y,s);
		if(R==NULL) { pBits+=w; continue; }
		for(pI=I,i=w; --i>=0; ++pBits,++pI)
		{
			if(*pI!=-1) *pBits = R[*pI];
		}
	}
}

//功能：取得指定矩形区域各个象素的灰度值
//参数x,y：矩形的最小x,y
//参数w,h：要取的数据的列数和行数
//参数pBits：返回的灰度值的内存块，函数的调用者必须准备足够多(不少于w*h)的空间
//			 [pBits[0]--pBits[w])是第一行数据,[pBits[w]--pBits[2*w])是第二行数据,……
//			 第一行的Y值是y,第二行的Y值是y+step,第三行的Y值是y+2*step,……
//			 第一列的X值是x,第二列的X值是x+step,第三列的X值是x+2*step,……
void PageCache::GetBitsEx4(double x,double y,double step,int w,int h,BYTE* pBits)
{
	static long I[4096]; long W,i,*pI;
	
    struct CB4
    {
        BYTE*   m_mem;
        int     m_w,m_h,m_y0;

        CB4(FILE* file,int w,double y,double s,int h)
            :m_mem(0),m_w(w)
        {
            m_y0 = (int)(y+0.5);
            if(m_y0<0) m_y0=0;
            m_h = (int)(h*s+0.5);

            if( 0!=fseek(file,54+1024+w*m_y0,SEEK_SET) )
                return;
            m_mem = new BYTE[m_w*m_h];
            m_h = fread(m_mem,m_w,m_h,file);
        }
        
        ~CB4()
        {
            delete m_mem;
        }

        BYTE* GetRow(int y)
        {
            if(m_mem==0)
                return 0;
            y -= m_y0;
            if(y<0||y>=m_h)
                return 0;
            return m_mem+y*m_w;
        }
    };

    x /= 4; y /= 4; step /= 4; 
    W = m_nWidth/4;
    CB4 b4(m_fb4,W,y,step,h);
    
    //得到每行的下标
	for(i=w; --i>=0;)
	{
		I[i] = (int)(x+i*step+0.5);
		if(I[i]<0 || I[i]>=W) I[i] = -1;
	}

	//一行一行的取数据
	memset(pBits,-1,w*h);
	for(;--h>=0; y+=step)
	{
		const BYTE* R = b4.GetRow( (int)(y+0.5) );
		if(R==NULL) { pBits+=w; continue; }
		for(pI=I,i=w; --i>=0; ++pBits,++pI)
		{
			if(*pI!=-1) *pBits = R[*pI];
		}
	}
}

static BITMAPINFO* BmpInfo()
{
    static BYTE buf[sizeof(BITMAPINFOHEADER)+256*sizeof(RGBQUAD)];
    static BITMAPINFO* info = 0;
    if(info!=0)
        return info;
    info = (BITMAPINFO*)buf;
    memset(buf,0,sizeof(buf));
    BITMAPINFOHEADER& H = info->bmiHeader;
    H.biSize = 40;
    H.biPlanes = 1;
    H.biBitCount = 8;
    RGBQUAD* Q = info->bmiColors;
    for(int i=256; --i>=0;)
    {
        Q[i].rgbBlue = Q[i].rgbGreen = Q[i].rgbRed = i;
    }
    Q[0].rgbRed = 255;
    return info;
}

void PageCache::DrawImage(HDC hDC,POINT pt,int w,int h,double fScale)
{
	YG_SetView(pt.x,w,fScale);

    static BYTE bits[1024*1024*8];
    BITMAPINFO& BI = *BmpInfo();
	
    if(fScale<=1)
    {
        BI.bmiHeader.biWidth = w = ((w+3)/4)*4;
        BI.bmiHeader.biHeight = h;
        if(m_bFast)
            GetBitsEx4(pt.x,pt.y,1.0/fScale,w,h,bits);
        else if(fScale==1)
			GetBits(pt.x,pt.y,w,h,bits);
		else
			GetBitsEx(pt.x,pt.y,1.0/fScale,w,h,bits);

        YG_Adjust(w,h,bits);
		ReverseBitsY(w,h,bits);
        SetDIBitsToDevice(hDC,0,0,w,h,0,0,0,h,bits,&BI,DIB_RGB_COLORS);
    }
    else //(fScale>1)
    {
        int w1,h1;
        
        w1 = (int)(w/fScale);
        h1 = (int)ceil(h/fScale);

        w1 = ((w1+3)/4)*4;
        BI.bmiHeader.biWidth = w1;
        BI.bmiHeader.biHeight = h1;
        if(m_bFast) GetBitsEx4(pt.x,pt.y,1,w1,h1,bits);
        else GetBits(pt.x,pt.y,w1,h1,bits);
		YG_Adjust(w1,h1,bits);
        ReverseBitsY(w1,h1,bits);
        w = w1*fScale;
        h = h1*fScale;
        StretchDIBits(hDC,0,0,w,h,0,0,w1,h1,bits,&BI,DIB_RGB_COLORS,SRCCOPY);
    }
}

void PageCache::SetFast(BOOL bFast,double scale)
{
    if(m_pMem==NULL)
        m_bFast = true;
    else if(m_fb4==NULL)
        m_bFast = false;
    else if(scale*4<1)
        m_bFast = true;
    else
        m_bFast = bFast;
}
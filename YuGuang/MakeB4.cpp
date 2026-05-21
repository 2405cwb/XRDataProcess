#include "stdafx.h"
#include "PageCache.h"
#include <process.h>
#include <stdio.h>

extern PageCache*  m_page;
BOOL ReadJpgGray(const char* pFileName,BYTE* pBits);
BOOL GetJpgSize(const char* pFileName,long& W,long& H);

static void Bmp8_WriteHead(FILE* file,int w,int h)
{
    int H[] = { 54+1024+w*h,0,54+1024,
        40,w,h,0x80001,0,0,0,0,0,0 };
    RGBQUAD rgb[256]; 

    rewind(file);
    fwrite("BM",2,1,file);
    fwrite(H,4,13,file);
    memset(rgb,0,sizeof(rgb));
    for(int i=256; --i>=0;)
    {
        rgb[i].rgbRed = 
        rgb[i].rgbBlue = 
        rgb[i].rgbGreen = i;
    }
    fwrite(rgb,4,256,file);
}

static inline int c44(BYTE* p,int w)
{
    int c = 0;
    c += p[0]; c += p[1]; c += p[2]; c += p[3];
    p += w;
    c += p[0]; c += p[1]; c += p[2]; c += p[3];
    p += w;
    c += p[0]; c += p[1]; c += p[2]; c += p[3];
    p += w;
    c += p[0]; c += p[1]; c += p[2]; c += p[3];
    return c/16;
}

static void _WriteB4(FILE* file,int w,int h,BYTE* p)
{
    int i,w4;
    w4 = w/4;
    for(h/=4; --h>=0; p+=(w*4))
    {
        for(i=0; i<w4; ++i)
        {
            p[i] = c44(p+i*4,w);
        }
        fwrite(p,1,w4,file);
    }
}

static void _CopyB4(BYTE* Q,int w,int h,BYTE* p)
{
    int i,w4;
    w4 = w/4;
    for(h/=4; --h>=0; p+=(w*4))
    {
        for(i=0; i<w4; ++i)
        {
            p[i] = c44(p+i*4,w);
        }
        memcpy(Q,p,w4); Q+=w4;
    }
}

char** FindFiles(const char* path,const char* ext,char** fn);

struct B4_Gen
{
    long    m_w1,m_h1;
    long    m_w,m_h,m_n,m_c;
    char**  m_fnA;
    char*   m_path;
    BYTE*   m_mem;

    B4_Gen(char* path)
    {
        char buf[512];
        memset(this,0,sizeof(*this));
        m_path = path;
        m_fnA = new char*[8000];
        m_fnA[0] = (char*)(m_fnA+1024);
        m_n = FindFiles(path,"jpg",m_fnA)-m_fnA;
        if(m_n==0) return;
        sprintf(buf,"%s\\%s",path,m_fnA[0]);
        GetJpgSize(buf,m_w1,m_h1);
        m_w = m_w1/4; m_h = m_h1/4;
        m_mem = new BYTE[m_w*m_h*m_n];
    }

    ~B4_Gen()
    {
        delete m_fnA;
        delete m_mem;
    }

    void MB4(int n,int i)
    {
        BYTE* mem; char buf[512];

        mem = new BYTE[m_w1*m_h1];
        for(; i<m_n; i+=n)
        {
            sprintf(buf,"%s\\%s",m_path,m_fnA[i]);
            ReadJpgGray(buf,mem);
            _CopyB4(m_mem+i*m_w*m_h,m_w1,m_h1,mem);
            InterlockedIncrement(&m_c);
        }
        delete mem;
    }

    static void _MB4(void* p)
    {
        void** pp = (void**)p;
        B4_Gen& gen = *((B4_Gen*)pp[0]);
        gen.MB4((int)pp[1],(int)pp[2]);
        delete pp;
    }

    void Run()
    {
        int i,n; void** pp;
        SYSTEM_INFO info; 
        GetSystemInfo(&info);
        n = info.dwNumberOfProcessors;
        for(i=0; i<n; ++i)
        {
            pp = new void*[3];
            pp[0] = this;
            pp[1] = (void*)n;
            pp[2] = (void*)i;
            ::_beginthread(_MB4,0,pp);
        }
    }
};

static void PumpMessage()
{
    MSG msg;
    while(PeekMessageA(&msg, 0, 0, 0, 1))
	{
		TranslateMessage(&msg);
		DispatchMessageA(&msg);
	}
}

static void _FWrite(HWND hWnd,FILE* file,BYTE* mem,int cs,int n)
{
    int i; char buf[32];
    for(i=0; i<n; ++i)
    {
        fwrite(mem+cs*i,cs,1,file);
        if(20*(i+1)/n==(20*i/n))
            continue;
        sprintf(buf,"正在保存… %d%%",100*(i+1)/n);
        ::SetWindowText(hWnd,buf);
        PumpMessage();
    }
}

extern "C" __declspec(dllexport) BOOL BH_MakeB4(HWND hWnd,char* path)
{
    char buf[512]; 
    B4_Gen gen(path);
    if(gen.m_n==0)
        return false;

    gen.Run();
    while(gen.m_c<gen.m_n)
    {
        Sleep(100);

        sprintf(buf,"正在解压… %d%%",(gen.m_c*100/gen.m_n));
        ::SetWindowText(hWnd,buf);

        PumpMessage();
	}

    if( m_page!=NULL )
        m_page->OpenFb4(0);

    sprintf(buf,"%s%s.b4",path,strrchr(path,'\\'));
    FILE* file = fopen(buf,"wb");    
    if( NULL==file )
        return false;

    Bmp8_WriteHead(file,gen.m_w,gen.m_h*gen.m_n);
    _FWrite(hWnd,file,gen.m_mem,gen.m_w*gen.m_h,gen.m_n);
    fclose(file);

    if( m_page!=NULL )
        m_page->OpenFb4(path);
    return true;
}


char** FindFiles(const char* path,const char* ext,char** fn)
{
	WIN32_FIND_DATAA D;

	memset(&D,0,sizeof(D));
	sprintf(D.cFileName,"%s\\*.%s",path,ext);
	HANDLE h = FindFirstFileA(D.cFileName,&D);
	if( h==NULL || h==INVALID_HANDLE_VALUE )
		return fn;
	do{
		strcpy(fn[0],D.cFileName);
		fn[1] = fn[0]+strlen(fn[0])+1;
		++fn;
	}while( FindNextFileA(h,&D) );
	FindClose(h);
	return fn;
}


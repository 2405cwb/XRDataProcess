#include "stdafx.h"
#include "ijl15.h"
#include <atlbase.h>

#pragma comment(lib, "ijl15.lib") 

static void* LoadProg(const char* dllname,const char* funcname)
{
	char buf[256];
	GetModuleFileName(NULL,buf,sizeof(buf));
	strcpy(strrchr(buf,'\\')+1,dllname);
	HMODULE  h = LoadLibraryA(buf);
	if(h==NULL)
		return NULL;
	return GetProcAddress(h,funcname);
}

static BOOL ReadTiffGray(const char* filename,BYTE* pBits)
{
	typedef BOOL (*fun_t)(const char* filename,BYTE* pBits);
	static fun_t fun = 0;
	static BOOL loaded = 0;
	if( !loaded )
	{
		loaded = true;
		fun = (fun_t)LoadProg("ZyTiff.dll","ReadTiffGray");
	}
	if(fun==0)
		return false;
	return fun(filename,pBits);
}

static BOOL GetTiffSize(const char* filename,long& w,long& h)
{
	typedef BOOL (*fun_t)(const char* filename,long* w,long* h);
	static fun_t fun = 0;
	static BOOL loaded = 0;
	if( !loaded )
	{
		loaded = true;
		fun = (fun_t)LoadProg("ZyTiff.dll","GetTiffSize");
	}
	if(fun==0)
		return false;
	return fun(filename,&w,&h);
}

#define DIB_HEADER_MARKER   ((WORD) ('M' << 8) | 'B')

static void _bit2gray(int c,BYTE* g)
{
	g[0] = c&1?255:0;
	g[1] = c&2?255:0;
	g[2] = c&4?255:0;
	g[3] = c&8?255:0;
	g[4] = c&16?255:0;
	g[5] = c&32?255:0;
	g[6] = c&64?255:0;
	g[7] = c&128?255:0;
}

static void Bit2Gray(int w,int h,BYTE* pBits,BYTE* pGray)
{
	for(w=(w*h)/8; --w>=0; pGray+=8)
	{
		_bit2gray(*pBits++,pGray);
	}
}

static BOOL ReadBmpGray(const char* filename,BYTE* pBits)
{
	if( ReadTiffGray(filename,pBits) )
		return true;
	if( strstr(filename,".bmp")==0 && strstr(filename,".BMP")==0 )
		return false;

    BITMAPFILEHEADER	file_head;
	BITMAPINFOHEADER	info_head;
	RGBQUAD pal[256];

    FILE* fin = fopen(filename,"rb");

    //读取文件头
	fread(&file_head, sizeof(file_head),1,fin);
	fread(&info_head,sizeof(info_head),1,fin);

	int nRow = (info_head.biWidth*info_head.biBitCount)/8;
	nRow = 4*((nRow+3)/4);

	//读取调色板信息
	if( !(info_head.biClrUsed==256 || info_head.biBitCount==8) )
	{
		if( info_head.biBitCount!=1 )
		{
			fclose(fin);
			return false;
		}
		fread( pal,sizeof(RGBQUAD),2,fin );
		static BYTE* tmp = 0;
		if( tmp==0 )
			tmp = new BYTE[info_head.biHeight*info_head.biWidth];
		fread(tmp,nRow,info_head.biHeight,fin);
		Bit2Gray(info_head.biWidth,info_head.biHeight,tmp,pBits);
		nRow = ((info_head.biWidth+3)/4)*4;
	}
	else
	{
		fread( pal,sizeof(RGBQUAD),256,fin );
		fread(pBits,nRow,info_head.biHeight,fin);
	}
	void ReverseBitsY(int w,int h,BYTE* p);
	ReverseBitsY(nRow,info_head.biHeight,pBits);
	fclose(fin);
	return true;
}

BOOL ReadJpgGray(const char* pFileName,BYTE* pBits)
{
	if( ReadBmpGray(pFileName,pBits) )
		return true;
	JPEG_CORE_PROPERTIES image;
	memset(&image,0,sizeof(image));
	IJLERR er = ijlInit( &image );

	image.JPGFile = pFileName;

	er = ijlRead( &image, IJL_JFILE_READPARAMS );
	if( er!=IJL_OK )
		return false;
	image.DIBBytes = pBits;
	image.DIBWidth    = image.JPGWidth;//( image.JPGWidth + 3 )/4*4;
	image.DIBHeight   = image.JPGHeight;
	image.DIBColor = IJL_G;
	image.DIBChannels = 1;

	er = ijlRead( &image, IJL_JFILE_READWHOLEIMAGE );
    ijlFree(&image);

    return er==IJL_OK;
}

static BOOL GetBmpSize(const char* filename,long& w,long& h)
{
	if( GetTiffSize(filename,w,h) )
		return true;
	if( strstr(filename,".bmp")==0 && strstr(filename,".BMP")==0 )
		return false;

    BITMAPFILEHEADER	file_head;
	BITMAPINFOHEADER	info_head;

    FILE* fin = fopen(filename,"rb");
	fread(&file_head, sizeof(file_head),1,fin);
	fread(&info_head,sizeof(info_head),1,fin);
	w = info_head.biWidth;
	h = info_head.biHeight;
	w = ((w+3)/4)*4;
	fclose(fin);
	return true;
}

BOOL GetJpgSize(const char* pFileName,long& W,long& H)
{
	if(pFileName==0 )
		return false;
	if( GetBmpSize(pFileName,W,H) )
		return true;
    JPEG_CORE_PROPERTIES image;
	memset(&image,0,sizeof(image));
	IJLERR er = ijlInit( &image );

	image.JPGFile = pFileName;

	er = ijlRead( &image, IJL_JFILE_READPARAMS );
    ijlFree(&image);
	if( er!=IJL_OK )
		return false;
    W = image.JPGWidth;
    H = image.JPGHeight;
    return true;
}

BOOL ReadJpgRGB(const char* pFileName,long& nWidth,long& nHeight,BYTE* pBits)
{
    JPEG_CORE_PROPERTIES image;
	memset(&image,0,sizeof(image));
	IJLERR er = ijlInit( &image );

	USES_CONVERSION;

	image.JPGFile = pFileName;

	er = ijlRead( &image, IJL_JFILE_READPARAMS );
	if( er!=IJL_OK )
		return false;

	image.DIBWidth = nWidth = image.JPGWidth;//( image.JPGWidth + 3 )/4*4;
	image.DIBHeight = nHeight = image.JPGHeight;
    image.DIBBytes = pBits;
    if( pBits==NULL )
        return false;
    
	image.DIBColor = IJL_BGR;
	image.DIBChannels = 3;

	er = ijlRead( &image, IJL_JFILE_READWHOLEIMAGE );
    ijlFree(&image);
    return er==IJL_OK;
}

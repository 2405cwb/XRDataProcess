// RectTracker.h: interface for the CRectTracker class.
//
//////////////////////////////////////////////////////////////////////

#if !defined(AFX_RECTTRACKER_H__C9701282_6FDA_47F0_B9E4_80E7AA357175__INCLUDED_)
#define AFX_RECTTRACKER_H__C9701282_6FDA_47F0_B9E4_80E7AA357175__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000

#include <vector>

using namespace std;

class ZRectTracker
{
	HWND	m_hWnd;
	HDC		m_hDC;
	HPEN	m_hPen;
public:
	ZRectTracker(HWND hWnd)
		:m_hWnd(hWnd),m_hDC(0)
	{
		::SetCapture(hWnd);
		::UpdateWindow(hWnd);
		m_hDC = GetDC(hWnd);
		m_hPen = CreatePen(PS_DASHDOT,0,0);
		::SelectObject(m_hDC,m_hPen);
		::SetROP2(m_hDC,R2_NOTXORPEN);
	}

	~ZRectTracker()
	{
		::ReleaseDC(m_hWnd,m_hDC);
		::DeleteObject(m_hPen);
		::ReleaseCapture();
	}

	BOOL Track(POINT pt,POINT& out)
	{
		out = pt;
		while(true)
		{	
			MSG msg;
			if( !::GetMessage(&msg, NULL, 0, 0) )
				return false;
			if (::GetCapture() != m_hWnd)
				return false;

			switch (msg.message)
			{
			case WM_RBUTTONDOWN:
			case WM_LBUTTONDOWN:
				::Rectangle(m_hDC,pt.x,pt.y,out.x,out.y);
				return false;
			case WM_LBUTTONUP:
            case WM_RBUTTONUP:
				::Rectangle(m_hDC,pt.x,pt.y,out.x,out.y);
				return true;
			case WM_MOUSEMOVE:
                ::SendMessage(m_hWnd,WM_USER+1,0,0);
				if( msg.lParam&0x80008000 )
					break;
				::Rectangle(m_hDC,pt.x,pt.y,out.x,out.y);
				out.x = LOWORD(msg.lParam);
				out.y = HIWORD(msg.lParam);
				::Rectangle(m_hDC,pt.x,pt.y,out.x,out.y);
				break;
			case WM_KEYDOWN:
                if( ((char)msg.wParam)==27 )
                {//按下Esc
                    ::Rectangle(m_hDC,pt.x,pt.y,out.x,out.y);
                    return false;
                }
			default:
				DispatchMessage(&msg);
				break;
			}
		}
	}

    
    enum
    {
        hit_nothing,
        hit_all,
        hit_top,
        hit_bottom,
        hit_left,
        hit_right
    };

    int HitTest(RECT rc,POINT pt)
    {
        const int size = 3;
        if(pt.x<rc.left-size || pt.x>rc.right+size)
            return hit_nothing;
        if(pt.y<rc.top-size || pt.y>rc.bottom+size)
            return hit_nothing;
        if(pt.x<rc.left+size)
            return hit_left;
        if(pt.x>rc.right-size)
            return hit_right;
        if(pt.y<rc.top+size)
            return hit_top;
        if(pt.y>rc.bottom-size)
            return hit_bottom;
        return hit_all;
    }

    BOOL TrackRect(RECT& rc)
	{
        HCURSOR hX,hY,hXY,hNothing;
        hNothing = LoadCursor(NULL,IDC_ARROW);
        hX = LoadCursor(NULL,IDC_SIZEWE);
        hY = LoadCursor(NULL,IDC_SIZENS);
        hXY = LoadCursor(NULL,IDC_SIZEALL);
        ::Rectangle(m_hDC,rc.left,rc.top,rc.right,rc.bottom);
        POINT pt = {0,0},ptold; RECT rcold;
        int hit = hit_nothing;
		while(true)
		{	
			MSG msg;
			if( !::GetMessage(&msg, NULL, 0, 0) )
				return false;
			if (::GetCapture() != m_hWnd)
				return false;

			switch (msg.message)
			{
			case WM_RBUTTONDOWN:
                return true;
			case WM_LBUTTONDOWN:
                pt.x = LOWORD(msg.lParam);
                pt.y = HIWORD(msg.lParam);
                ptold = pt; rcold = rc;
                hit = HitTest(rc,pt);
                break;
			case WM_LBUTTONUP:
                hit = hit_nothing;
				break;
			case WM_MOUSEMOVE:
                ::SendMessage(m_hWnd,WM_USER+1,0,0);
				if( msg.lParam&0x80008000 )
					break;
                pt.x = LOWORD(msg.lParam);
                pt.y = HIWORD(msg.lParam);
                if( (msg.wParam&MK_LBUTTON)==0 )
                {   
                    switch(HitTest(rc,pt))
                    {
                    case hit_nothing: 
                        SetCursor(hNothing); break;
                    case hit_left: case hit_right:
                        SetCursor(hX); break;
                    case hit_bottom: case hit_top:
                        SetCursor(hY); break;
                    case hit_all:
                        SetCursor(hXY); break;
                    }
                    break;
                }

                ::Rectangle(m_hDC,rc.left,rc.top,rc.right,rc.bottom);
                switch(hit)
                {
                case hit_left:
                    rc.left = pt.x;
                    break;
                case hit_right:
                    rc.right = pt.x;
                    break;
                case hit_bottom:
                    rc.bottom = pt.y;
                    break;
                case hit_top:
                    rc.top = pt.y;
                    break;
                case hit_all:
                    pt.x -= ptold.x; pt.y -= ptold.y;
                    rc.left = rcold.left+pt.x; rc.right = rcold.right+pt.x;
                    rc.top = rcold.top+pt.y; rc.bottom = rcold.bottom+pt.y;
                    break;
                }
                ::Rectangle(m_hDC,rc.left,rc.top,rc.right,rc.bottom);

				break;
            
			default:
				DispatchMessage(&msg);
				break;
			}
		}
	}

    BOOL TrackPolyline(std::vector<long>& ptA)
	{
        ::DeleteObject(m_hPen);
        m_hPen = CreatePen(PS_SOLID,0,0);
        ::SelectObject(m_hDC,m_hPen);
		while(true)
		{	
			MSG msg;
			if( !::GetMessage(&msg, NULL, 0, 0) )
				return false;
			if (::GetCapture() != m_hWnd)
				return false;

			switch (msg.message)
			{
			case WM_RBUTTONDOWN: case WM_LBUTTONDOWN:
				return false;
            case WM_KEYDOWN:
                if( ((char)msg.wParam)==27 )//按下Esc
                    return false;
                break;
			case WM_LBUTTONUP: case WM_RBUTTONUP:
				return true;
			case WM_MOUSEMOVE:
				if( msg.lParam&0x80008000 )
					break;
				ptA.push_back( LOWORD(msg.lParam) );
				ptA.push_back( HIWORD(msg.lParam) );
                ::Polyline(m_hDC,(POINT*)&ptA[0],ptA.size()/2);
				break;
			default:
				DispatchMessage(&msg);
				break;
			}
		}
	}
};

#endif // !defined(AFX_RECTTRACKER_H__C9701282_6FDA_47F0_B9E4_80E7AA357175__INCLUDED_)

#include "stdafx.h"
#include "RectTracker.h"

extern "C" __declspec(dllexport) int BH_RectTrack(HWND hWnd,POINT* pt)
{
    ZRectTracker T(hWnd);
    if(pt[0].x==pt[1].x && pt[0].y==pt[1].y)
        return T.Track(pt[0],pt[1]);
    else
        return T.TrackRect(*(RECT*)pt);
}
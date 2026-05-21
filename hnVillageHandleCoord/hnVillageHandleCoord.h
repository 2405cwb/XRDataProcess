#ifndef HNCALCU_IRI_METHOD_API_H
#define HNCALCU_IRI_METHOD_API_H
#include "hnCalcuMethod_global.h"
#include <fstream>
#include"opencv2/opencv.hpp"

//using namespace std;
using namespace cv;

  class  HNCALCUIRIMETHOD_API hnVillageHandleCoord
{
public:
	hnVillageHandleCoord(void);
	~hnVillageHandleCoord(void);
	void init();
	void getHandelCoord(float& x,float& y);
private:
	cv::Mat matu;
	cv::Mat matv;
    int CutImgX_value;
    int CutImgY_value;
 };



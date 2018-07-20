#pragma once
#include "opencv2/imgproc/imgproc.hpp"
#include "opencv2/highgui/highgui.hpp"


using namespace cv;

template <typename T>
struct Color
{
	T r;
	T g;
	T b;
};

struct RGBColor : public Color<float>
{
	unsigned int EncodeRGB()
	{
		uint8_t bR = (uint8_t)(r * (uint8_t)0xFF);
		uint8_t bG = (uint8_t)(g * (uint8_t)0xFF);
		uint8_t bB = (uint8_t)(b * (uint8_t)0xFF);

		return ((BYTE)bR | ((WORD)((BYTE)bG) << 8)) | (((DWORD)(BYTE)bB) << 16);
	}

	RGBColor(float rr, float gg, float bb)
	{
		r = rr;
		g = gg;
		b = bb;
	}
};

struct CellInfo
{
	uint32_t row, col;
	float value;
};

struct Settings
{
	float minValue;
	float maxValue;

	float width;
	float height;	

	uint32_t rowsCount;
	uint32_t columnsCount;

	uint32_t kernelX;
	uint32_t kernelY;
	float sigmaX;
	float sigmaY;

	LPCWSTR   path;
};

struct Image
{
	Mat* srcImage;
	Mat* dstImage;
	uint32_t firstRow, firstCol, irangeRows, irangeColumns;
	float height, width;
	float wStep, hStep;
	float valueMin, valueMax;
};
Image image;

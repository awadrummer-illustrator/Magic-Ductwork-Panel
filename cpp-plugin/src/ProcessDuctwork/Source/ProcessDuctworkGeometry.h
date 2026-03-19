#ifndef __ProcessDuctworkGeometry_H__
#define __ProcessDuctworkGeometry_H__

#include "IllustratorSDK.h"

#include <string>
#include <vector>

struct DuctworkPoint
{
	double x;
	double y;
};

struct DuctworkPath
{
	AIArtHandle art;
	std::vector<DuctworkPoint> points;
	bool closed;
	std::string layerName;
};

namespace DuctworkGeometry
{
	bool GetPathPoints(AIArtHandle path, std::vector<DuctworkPoint>& outPoints, bool& outClosed);
	std::string GetArtLayerName(AIArtHandle art);
}

#endif // __ProcessDuctworkGeometry_H__

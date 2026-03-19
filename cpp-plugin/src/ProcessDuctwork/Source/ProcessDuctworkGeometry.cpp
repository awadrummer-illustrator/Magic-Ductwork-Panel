#include "IllustratorSDK.h"
#include "ProcessDuctworkGeometry.h"
#include "ProcessDuctworkSuites.h"

bool DuctworkGeometry::GetPathPoints(AIArtHandle path, std::vector<DuctworkPoint>& outPoints, bool& outClosed)
{
	outPoints.clear();
	outClosed = false;
	if (!path || !sAIPath || !sAIArt) {
		return false;
	}

	short artType = kUnknownArt;
	if (sAIArt->GetArtType(path, &artType) || artType != kPathArt) {
		return false;
	}

	ai::int16 count = 0;
	if (sAIPath->GetPathSegmentCount(path, &count) || count <= 0) {
		return false;
	}

	std::vector<AIPathSegment> segments(static_cast<size_t>(count));
	if (sAIPath->GetPathSegments(path, 0, count, &segments[0])) {
		return false;
	}

	outPoints.reserve(static_cast<size_t>(count));
	for (ai::int16 i = 0; i < count; ++i) {
		DuctworkPoint pt;
		pt.x = segments[i].p.h;
		pt.y = segments[i].p.v;
		outPoints.push_back(pt);
	}

	AIBoolean closed = false;
	if (!sAIPath->GetPathClosed(path, &closed)) {
		outClosed = (closed != 0);
	}
	return true;
}

std::string DuctworkGeometry::GetArtLayerName(AIArtHandle art)
{
	if (!art || !sAIArt || !sAILayer) {
		return std::string();
	}
	AILayerHandle layer = nullptr;
	if (sAIArt->GetLayerOfArt(art, &layer) || !layer) {
		return std::string();
	}
	ai::UnicodeString title;
	if (sAILayer->GetLayerTitle(layer, title)) {
		return std::string();
	}
	return title.as_UTF8();
}

#include "HighAccDiseaseOut.h"
#include <unordered_set>
bool _stdcall HighAccOut(const char* path, int diseaseNum, HighAccDisease* disease)
{

	std::vector<HighAccDisease> vecDisease;
	vecDisease.resize(diseaseNum);
	for (int i = 0; i < diseaseNum; i++)
	{
		vecDisease[i] = disease[i];
	}
	 

	DL_Dxf* dxf = new DL_Dxf();
	DL_Codes::version exportVersion = DL_Codes::AC1015;
	DL_WriterA* dw = dxf->out(path, exportVersion);
	if (dw == NULL)
	{
		return false;
	}

	initCAD(dxf, dw);
	dw->sectionEntities();
	//对病害进行分类 统计有多少类别的
	std::unordered_set<std::string> uniqueNames;
	for (auto dis : vecDisease)
	{
		std::string fullName = dis.name;
		auto index = fullName.find('.');
		std::string name;
		if (index != std::string::npos)
		{
			name = fullName.substr(0, index);
		}
		else
		{
			name = fullName;
		}

		uniqueNames.insert(name);
	}
	//创建图层
	dw->tableLayers(uniqueNames.size());
	{
		for (auto name : uniqueNames)
		{
			dxf->writeLayer(*dw,
				DL_LayerData(name, 0),
				DL_Attributes(
					std::string(""),
					DL_Codes::red,
					50,
					"CONTINUOUS", 1.0));
		}
		dw->tableEnd();
	}
	dw->sectionEntities();
	for (int diseaseIndex = 0; diseaseIndex < vecDisease.size(); ++diseaseIndex)
	{
		auto disease = vecDisease.at(diseaseIndex);
		//判断 病害所在图层名称
		std::string disName = disease.name;

		unsigned vertexCount = 4;
		int flags = 1;
		// flags |= 8;  //3D Polyline

		//获得图层名称
		std::string layoutName;
		for (auto tempName : uniqueNames)
		{
			if (disName.find(tempName) != std::string::npos)
			{
				layoutName = tempName;
				break;
			}
			else
			{
				layoutName = disName; //不可能发生
			}
		}
		double xLength = disease.p1.DiseaseLon - disease.p0.DiseaseLon;
		double yLenght = disease.p1.DiseaseLat - disease.p0.DiseaseLat;
		double length = sqrt(xLength * xLength + yLenght * yLenght);
		double slops = yLenght / xLength;
		double angleRadians = atan2(yLenght, xLength);
		double textRatio = length * 0.1;

		//将弧度转为度数
		double textAngle = angleRadians * 180.0 / M_PI;
		if (textAngle <= 0)
		{
			textAngle += 360;
		}
		textAngle = textAngle / 180.0 * M_PI;
		//写入多段线
		dxf->writePolyline(*dw, DL_PolylineData(static_cast<int>(vertexCount), 0, 0, flags), DL_Attributes(layoutName, 256, -1.0, "BYLAYER", 8.0));
		dxf->writeVertex(*dw, DL_VertexData(disease.p0.DiseaseLon, disease.p0.DiseaseLat, disease.p0.DiseaseHeight));
		dxf->writeVertex(*dw, DL_VertexData(disease.p1.DiseaseLon, disease.p1.DiseaseLat, disease.p1.DiseaseHeight));
		dxf->writeVertex(*dw, DL_VertexData(disease.p2.DiseaseLon, disease.p2.DiseaseLat, disease.p2.DiseaseHeight));
		dxf->writeVertex(*dw, DL_VertexData(disease.p3.DiseaseLon, disease.p3.DiseaseLat, disease.p3.DiseaseHeight)); 
		dxf->writePolylineEnd(*dw);
		double centerX = (disease.p0.DiseaseLon + disease.p1.DiseaseLon + disease.p2.DiseaseLon + disease.p3.DiseaseLon) / 4;
		double centerY = (disease.p0.DiseaseLat + disease.p1.DiseaseLat + disease.p2.DiseaseLat + disease.p3.DiseaseLat) / 4;

		dxf->writeText(*dw, DL_TextData(centerX, centerY, 0, disease.p1.DiseaseLon, disease.p1.DiseaseLat, 0, textRatio, 1, 0, 0, 0, disName, "Standard", textAngle), DL_Attributes(layoutName, 256, -1, "BYLAYER", 1.0));

	}
	dw->sectionEnd();
	// 写入dxf
	dxf->writeObjects(*dw);
	dxf->writeObjectsEnd(*dw);

	dw->dxfEOF();
	dw->close();
	delete dw;
	dw = NULL;
	delete dxf;
	dxf = NULL; 
	return true;
}

bool initCAD(DL_Dxf* dxf, DL_WriterA* dw)
{
	if (dw == NULL)
		return false;

	dxf->writeHeader(*dw);
	dw->sectionEnd();
	dw->sectionTables();
	dxf->writeVPort(*dw);
	dw->tableLinetypes(8);
	dxf->writeLinetype(*dw, DL_LinetypeData("BYBLOCK", "BYBLOCK", 0, 0, 0.0));
	dxf->writeLinetype(*dw, DL_LinetypeData("BYLAYER", "BYLAYER", 0, 0, 0.0));
	dxf->writeLinetype(*dw, DL_LinetypeData("CONTINUOUS", "Continuous", 0, 0, 0.0));
	dxf->writeLinetype(*dw, DL_LinetypeData("DASHED", "DASHED", 0, 0, 0.0));
	dxf->writeLinetype(*dw, DL_LinetypeData("DASHED2", "DASHED2", 0, 0, 0.0));
	double patternCenter[4] = { 31.75 ,-6.349999999999997 ,6.349999999999997 ,-6.349999999999997 };
	dxf->writeLinetype(*dw, DL_LinetypeData("CENTER", "Center ____ _ ____ _ ____ _ ____ _ ____ _ ____", 0/*70*/, 4/*73*/, 50.8/*40*/, patternCenter));
	double patternDot[4] = { 24.0 ,-3.0 ,0.0 ,-3.0 };
	dxf->writeLinetype(*dw, DL_LinetypeData("ACAD_ISO04W100", "ISO long-dash dot ____ . ____ . ____ . ____ . _", 0/*70*/, 4/*73*/, 30.0/*40*/, patternDot));

	double patternCenter2[4] = { 0.4 ,-0.1 ,0.1 ,-0.1 };
	dxf->writeLinetype(*dw, DL_LinetypeData("CENTERBLOCK", "Center ____ _ ____ _ ____ _ ____ _ ____ _ ____", 0/*70*/, 4/*73*/, 0.7/*40*/, patternCenter2));
	dw->tableEnd();

	int numberOfLayers = 1;
	dw->tableLayers(numberOfLayers);

	//0必须存在
	dxf->writeLayer(*dw,
		DL_LayerData("0", 0),
		DL_Attributes(
			std::string(""),
			DL_Codes::black,
			256,
			"CONTINUOUS", 1.0));

	dw->tableEnd();

	dw->tableStyle(1);
	DL_StyleData style("Standard", 0, 0.0, 1, 0.0, 0, 2.0, "宋体", "");
	style.bold = false;
	style.italic = false;
	dxf->writeStyle(*dw, style);
	dw->tableEnd();

	dxf->writeView(*dw);
	dxf->writeUcs(*dw);

	dw->tableAppid(1);
	dxf->writeAppid(*dw, "ACAD");
	dw->tableEnd();

	dxf->writeDimStyle(*dw, 1, 1, 1, 1, 1);

	dxf->writeBlockRecord(*dw);


	dw->tableEnd();

	dw->sectionEnd();

	dw->sectionBlocks();
	dxf->writeBlock(*dw, DL_BlockData("*Model_Space", 0, 0.0, 0.0, 0.0));
	dxf->writeEndBlock(*dw, "*Model_Space");
	dxf->writeBlock(*dw, DL_BlockData("*Paper_Space", 0, 0.0, 0.0, 0.0));
	dxf->writeEndBlock(*dw, "*Paper_Space");
	dxf->writeBlock(*dw, DL_BlockData("*Paper_Space0", 0, 0.0, 0.0, 0.0));
	dxf->writeEndBlock(*dw, "*Paper_Space0");
	dw->sectionEnd();
	return true;
}

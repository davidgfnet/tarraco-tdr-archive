
/*

  Cal3DVb, an ATL Object implementation class for Cal3D

  Thanks to Cal3D team for the lib.

  ATL bridge for Vb by davidgf (davidgf@tinet.org)

*/

// Cal3DObject.cpp : Implementation of CCal3DObject
#include "stdafx.h"
#include "Cal3dvb.h"
#include "Cal3DObject.h"
#include "cal3d/cal3d.h"
#include <comdef.h>
#include "windows.h"
#include "oleauto.h"

/////////////////////////////////////////////////////////////////////////////
// CCal3DObject

using namespace std;

void BSTRtoASC (BSTR str, string * strRet) {
 if ( str != NULL ) {
	 char * cstr;
	 unsigned long length = WideCharToMultiByte (
	 CP_ACP,
	 0, str,
	 SysStringLen(str), NULL, 0,
	 NULL, NULL
	 );

	 cstr = new char[length];

	 length = WideCharToMultiByte (
	 CP_ACP,
	 0, str,
	 SysStringLen(str), reinterpret_cast <char *>(cstr), length,
	 NULL, NULL
	 );

	 cstr[length] = '\0';
	 strRet = new string(cstr);
	 strRet->resize (length);
 }
}

STDMETHODIMP CCal3DObject::loadmesh(BSTR *file, int *meshid)
{
	USES_CONVERSION;
	std::string sfile = W2A (*file);

	*meshid = coremodel->loadCoreMesh(sfile);

	meshlist[nummeshes]=*meshid;
	nummeshes++;

	return S_OK;
}

STDMETHODIMP CCal3DObject::loadskeleton(BSTR *file, int *result)
{
	USES_CONVERSION;
	std::string sfile = W2A (*file);

	bool succeed;

	succeed = coremodel->loadCoreSkeleton(sfile);
	if (succeed==true) {*result=1;}else{*result=-1;}

	return S_OK;
}

STDMETHODIMP CCal3DObject::loadanimation(BSTR *file, int *animid)
{
	USES_CONVERSION;
	std::string sfile = W2A (*file);

	*animid=coremodel->loadCoreAnimation(sfile);

	return S_OK;
}

STDMETHODIMP CCal3DObject::getanimationduration(int animid, float *animationduration)
{
	*animationduration=coremodel->getCoreAnimation(animid)->getDuration();

	return S_OK;
}

STDMETHODIMP CCal3DObject::loadmaterial(BSTR *file, int *matid)
{
	USES_CONVERSION;
	std::string sfile = W2A (*file);

	*matid=coremodel->loadCoreMaterial(sfile);

	return S_OK;
}


STDMETHODIMP CCal3DObject::ready(int *texturelist, int *numtextures)
{
	int matid;
	for(matid = 0; matid < coremodel->getCoreMaterialCount(); matid++)
	{
		std::string strFilename;
		strFilename = coremodel->getCoreMaterial(matid)->getMapFilename(0);  //mapid=0
		strFilename.resize (256,32);

		CopyMemory (&texturelist[matid*64],&strFilename[0],256);

		coremodel->getCoreMaterial(matid)->setMapUserData(0, (Cal::UserData)matid);

		coremodel->createCoreMaterialThread(matid);
		coremodel->setCoreMaterialId(matid,0,matid);
    }
	*numtextures=coremodel->getCoreMaterialCount();

	return S_OK;
}

//////////////////////////////////////////////////////////////////////
//////////////////////////////////////////////////////////////////////

STDMETHODIMP CCal3DObject::createmodel(int *modelid)
{
	models[nummodels]= new CalModel(coremodel);
	
	int i;
	for (i=0; i<nummeshes; i++)
	{
		models[nummodels]->attachMesh(meshlist[i]);
	}

	models[nummodels]->setMaterialSet(0);

	*modelid=nummodels;
	nummodels++;

	return S_OK;
}

STDMETHODIMP CCal3DObject::setlod(int modelid, float lod)
{
	models[modelid]->setLodLevel(lod);
	
	return S_OK;
}

STDMETHODIMP CCal3DObject::update(int modelid, float eseconds)
{
	models[modelid]->update(eseconds);
	
	return S_OK;
}

STDMETHODIMP CCal3DObject::blendcycle(int modelid, int animid, float weight, float delay)
{
    models[modelid]->getMixer()->blendCycle(animid, weight, delay);
  	
	return S_OK;
}


STDMETHODIMP CCal3DObject::clearcycle(int modelid, int animid, float delay)
{
    models[modelid]->getMixer()->clearCycle(animid, delay);
  	
	return S_OK;
}

STDMETHODIMP CCal3DObject::setanimationtime(int modelid, float time)
{
    models[modelid]->getMixer()->setAnimationTime (time);
  	
	return S_OK;
}

STDMETHODIMP CCal3DObject::executeaction(int modelid, int animid, float delayin, float delayout, float weight)
{
    models[modelid]->getMixer()->executeAction (animid, delayin, delayout, weight);
  	
	return S_OK;
}

STDMETHODIMP CCal3DObject::render(int modelid, int *numvertices, int *numfaces, int *nummaterials, float *vertices, float *uvcoords, int *indices , int *textures, int *atributes, int using_vs)
{
	int mid, meshcount, submeshcount, sid,i;

	int vertexCount, textureCoordinateCount, faceCount;
	int matcont=0,subset=0;

	meshcount = coremodel->getCoreMeshCount();
	*numvertices=0;*numfaces=0;

	for(mid = 0; mid < meshcount; mid++)
	{	submeshcount = coremodel->getCoreMesh(mid)->getCoreSubmeshCount();
		for(sid = 0; sid < submeshcount; sid++)
		{
			textures[matcont]= models[modelid]->getMesh(mid)->getSubmesh(sid)->getCoreMaterialId ();
			matcont++;
		}
	}

	CalRenderer *render;
    render = models[modelid]->getRenderer();
	render->beginRendering();

	for(mid = 0; mid < meshcount; mid++)
	{
		// get the number of submeshes
		submeshcount = render->getSubmeshCount(mid);

		// loop through all submeshes of the mesh
		for(sid = 0; sid < submeshcount; sid++)
		{
			// select mesh and submesh for further data access
			render->selectMeshSubmesh(mid, sid);

			*nummaterials = *nummaterials+1;

			//if we are using VS then we do not have to interleave the xyz with uv components
			if (using_vs==1) {
				//get face indices
				faceCount = render->getFaces(&indices[*numfaces*3]);
				// get the texture coordinates of the submesh
				textureCoordinateCount = render->getTextureCoordinates(0, &uvcoords[*numvertices*2]);
				// get the transformed vertices of the submesh
				vertexCount = render->getVertices(&vertices[*numvertices*3]);
			}else{
				faceCount = render->getFaces(&indices[0]);
				textureCoordinateCount = render->getTextureCoordinates(0, &uvcoords[0]);
				vertexCount = render->getVertices(&myvertices[0]);

				for (i=0; i<(faceCount*3); i++) {
					CopyMemory (&vertices[(i+(*numfaces*3))*5],&myvertices[indices[i]*3],12);
					CopyMemory (&vertices[(i+(*numfaces*3))*5+3],&uvcoords[indices[i]*2],8);
				}
			}

			atributes[subset++]=*numfaces;
			atributes[subset++]=faceCount;
			*numfaces = *numfaces + faceCount;

			atributes[subset++]=*numvertices;
			*numvertices = *numvertices + vertexCount;
		}
	}

	render->endRendering();

	return S_OK;
}



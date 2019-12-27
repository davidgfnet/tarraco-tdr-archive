// Cal3DObject.h : Declaration of the CCal3DObject

#ifndef __CAL3DOBJECT_H_
#define __CAL3DOBJECT_H_

#include "resource.h"       // main symbols
#include "cal3d/cal3d.h"

/////////////////////////////////////////////////////////////////////////////
// CCal3DObject
class ATL_NO_VTABLE CCal3DObject : 
	public CComObjectRootEx<CComSingleThreadModel>,
	public CComCoClass<CCal3DObject, &CLSID_Cal3DObject>,
	public IDispatchImpl<ICal3DObject, &IID_ICal3DObject, &LIBID_CAL3DVBLib>
{
public:

	CCal3DObject()
	{
		coremodel = new CalCoreModel("tarraco model");

		//init all vars
		nummeshes=0;
		nummodels=0;
	}

	~CCal3DObject()
	{
		for (int i = 0; i<nummodels; i++) 
		{
			delete models[i];
		}

		delete coremodel;
	}

DECLARE_REGISTRY_RESOURCEID(IDR_CAL3DOBJECT)

DECLARE_PROTECT_FINAL_CONSTRUCT()

BEGIN_COM_MAP(CCal3DObject)
	COM_INTERFACE_ENTRY(ICal3DObject)
	COM_INTERFACE_ENTRY(IDispatch)
END_COM_MAP()

// ICal3DObject
public:
	STDMETHOD(loadmesh)(BSTR *file, int *meshid);
	STDMETHOD(loadskeleton)(BSTR *file, int *result);
	STDMETHOD(loadanimation)(BSTR *file, int *animid);
	STDMETHOD(loadmaterial)(BSTR *file, int *matid);
	STDMETHOD(createmodel)(int *modelid);

	STDMETHOD(update)(int modelid, float eseconds);
	STDMETHOD(setlod)(int modelid, float lod);
	STDMETHOD(render)(int modelid, int *numvertices, int *numfaces, int *nummaterials, float *vertices, float *uvcoords, int *indices , int *textures, int *atributes, int using_vs);
	STDMETHOD(blendcycle)(int modelid, int animid, float weight, float delay);
	STDMETHOD(clearcycle)(int modelid, int animid, float delay);
	STDMETHOD(setanimationtime)(int modelid, float time);
	STDMETHOD(ready)(int *texturelist, int *numtextures);
	STDMETHOD(getanimationduration)(int animid, float *animationduration);
	STDMETHOD(executeaction)(int modelid, int animid, float delayin, float delayout, float weight);
	
private:
	CalCoreModel *coremodel;

	int nummodels;
	CalModel *models[64];

	int meshlist[64];
	int nummeshes;

	float myvertices[300000];
};

#endif //__CAL3DOBJECT_H_

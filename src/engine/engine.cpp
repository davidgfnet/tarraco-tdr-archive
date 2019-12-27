/* 

  engine core Maths Library

  Created by David GF using code from the

    ColDet - C++ 3D Collision Detection Library
    Copyright (C) 2000   Amir Geva
  
  Under GNU LESSER GENERAL PUBLIC LICENSE (Version 2.1, February 1999)

          http://photoneffect.com/coldet/

*/

#include "stdafx.h"
#include "coldet.h"
#include <math.h>


#pragma warning(disable:4244)


typedef struct  D3DVECTOR
    {
    float x;
    float y;
    float z;
}	D3DVECTOR;


typedef struct  salida
    {
    int respuesta;
	D3DVECTOR puntocolision;
}	salida;


D3DVECTOR CrossProduct(const D3DVECTOR& v1, const D3DVECTOR& v2)
{
	D3DVECTOR resultado;
	resultado.x=v1.y*v2.z-v2.y*v1.z;
    resultado.y=v1.z*v2.x-v2.z*v1.x;
	resultado.z=v1.x*v2.y-v2.x*v1.y;
	return resultado;
}

float DotProduct(const D3DVECTOR& v1, const D3DVECTOR& v2)
{
	return v1.x * v2.x + v1.y * v2.y + v1.z * v2.z;
}

D3DVECTOR resta(const D3DVECTOR &v1, const D3DVECTOR &v2)
{
	D3DVECTOR resultado;
	resultado.x=v1.x-v2.x;
    resultado.y=v1.y-v2.y;
	resultado.z=v1.z-v2.z;
	return resultado;
}

D3DVECTOR suma(const D3DVECTOR& v1, const D3DVECTOR& v2)
{
	D3DVECTOR resultado;
	resultado.x=v1.x+v2.x;
    resultado.y=v1.y+v2.y;
	resultado.z=v1.z+v2.z;
	return resultado;
}


D3DVECTOR multiplica(const D3DVECTOR& v1, long fact)
{
	D3DVECTOR resultado;
	resultado.x=v1.x*fact;
    resultado.y=v1.y*fact;
	resultado.z=v1.z*fact;
	return resultado;
}

D3DVECTOR Normalize(const D3DVECTOR& v1)
{
	D3DVECTOR resultado;
	float module;
	module=sqrt(v1.x*v1.x + v1.y*v1.y + v1.z*v1.z);
	resultado.x=v1.x/module;
    resultado.y=v1.y/module;
	resultado.z=v1.z/module;
	return resultado;
}

D3DVECTOR copyv (const D3DVECTOR& v1)
{
	D3DVECTOR resultado;
	resultado.x = v1.x;
	resultado.y = v1.y;
	resultado.z = v1.z;
	return resultado;
}

int vcompare (const D3DVECTOR& v1, const D3DVECTOR& v2)
{
	return ((v1.x == v2.x) && (v1.y == v2.y) && (v1.z == v2.z));
}

/*
	 Processa el moviment horitzontal, les col·lisions horitzontals
     i la càmera (amb la seva corresponent col·lisió)
*/


__declspec( dllexport ) _stdcall process(D3DVECTOR *cam, D3DVECTOR *tri, long numtri, D3DVECTOR *pos, D3DVECTOR *pos2, D3DVECTOR *upv, double distance, double angle, double angleh, float disfromcol, float interpolationspeed, float midh, float hih, float avrframe, float speed, float dimensions, long *tris, long cameratype)
{
	double CameraDistanceProj, CameraRelX, CameraRelZ, pi;
	float point1[3],point2[3];
	long i;
	double collx,colly,collz;
	pi = 3.14159265358979;
	
	pos->x = sin(angleh * pi / 180) * speed * avrframe / 1000 + pos->x ;
    pos->z = cos(angleh * pi / 180) * speed * avrframe / 1000 + pos->z;


	D3DVECTOR tripos[6];
	CollisionModel3D* model= newCollisionModel3D(TRUE);
	CollisionModel3D* model2= newCollisionModel3D(TRUE);

	model->setTriangleNumber (numtri);
	model2->setTriangleNumber (2);

	double var;
	var = cos(pi/6) * dimensions;

	tripos[0].x = pos->x + dimensions;
	tripos[0].y = pos->y + midh;
	tripos[0].z = pos->z;
	tripos[1].x = pos->x - dimensions;
	tripos[1].y = pos->y + midh;
	tripos[1].z = pos->z + var;
	tripos[2].x = pos->x - dimensions;
	tripos[2].y = pos->y + midh;
	tripos[2].z = pos->z - var;
	tripos[3].x = pos->x - dimensions;
	tripos[3].y = pos->y + midh;
	tripos[3].z = pos->z;
	tripos[4].x = pos->x + dimensions;
	tripos[4].y = pos->y + midh;
	tripos[4].z = pos->z + var;
	tripos[5].x = pos->x + dimensions;
	tripos[5].y = pos->y + midh;
	tripos[5].z = pos->z - var;

	D3DVECTOR vec, cam2,cam_original;
	cam_original.x  = cam->x;cam_original.y  = cam->y;cam_original.z  = cam->z;

	CameraDistanceProj = (distance * cos(angle * pi / 180));
	CameraRelX = CameraDistanceProj * sin(angleh * pi / 180);
	CameraRelZ = CameraDistanceProj * cos(angleh * pi / 180);
	cam->x = CameraRelX + pos->x;
	cam->z = CameraRelZ + pos->z;
	cam->y = pos->y + (distance * sin(angle * pi / 180));
	cam2.x  = cam->x;cam2.y  = cam->y;cam2.z  = cam->z;
	
	BOOL chocan=FALSE;
	double varx1,varx2,varz1,varz2;
	var = pos->y + midh;
	varx1=pos->x+distance;
	varx2=pos->x-distance;
	varz1=pos->z+distance;
	varz2=pos->z-distance;
	double var2=cam->y;

	long tricounter;
	tricounter=0;

	/* Optimitació! Es descarten triangles utilitzant una bounding box */

	for (i=0;i<(numtri*3);i+=3) {
		if ( (tri[i].y > var) || (tri[i+1].y > var) || (tri[i+2].y > var) ) {
		if ( (tri[i].y < var2) || (tri[i+1].y < var2) || (tri[i+2].y < var2) ) {
		if ( (tri[i].x > varx2) || (tri[i+1].x > varx2) || (tri[i+2].x > varx2) ) {
		if ( (tri[i].x < varx1) || (tri[i+1].x < varx1) || (tri[i+2].x < varx1) ) {
		if ( (tri[i].z > varz2) || (tri[i+1].z > varz2) || (tri[i+2].z > varz2) ) {
		if ( (tri[i].z < varz1) || (tri[i+1].z < varz1) || (tri[i+2].z < varz1) ) {
			model->addTriangle (tri[i].x,tri[i].y,tri[i].z,tri[i+1].x,tri[i+1].y,tri[i+1].z,tri[i+2].x,tri[i+2].y,tri[i+2].z);
			tris[tricounter++]=i;
		} } } } } }
	}
	for (i=0;i<6;i+=3) {
		model2->addTriangle (tripos[i].x,tripos[i].y,tripos[i].z,tripos[i+1].x,tripos[i+1].y,tripos[i+1].z,tripos[i+2].x,tripos[i+2].y,tripos[i+2].z);
	}
	model->finalize ();
	model2->finalize ();

	int firstt, secondt;

	if (model2->collision (model)==TRUE) {
		pos->x = pos2->x;
		pos->z = pos2->z;

		var = cos(pi/6) * dimensions;

		tripos[0].x = pos->x + dimensions;
		tripos[0].y = pos->y + midh;
		tripos[0].z = pos->z;
		tripos[1].x = pos->x - dimensions;
		tripos[1].y = pos->y + midh;
		tripos[1].z = pos->z + var;
		tripos[2].x = pos->x - dimensions;
		tripos[2].y = pos->y + midh;
		tripos[2].z = pos->z - var;
		tripos[3].x = pos->x - dimensions;
		tripos[3].y = pos->y + midh;
		tripos[3].z = pos->z;
		tripos[4].x = pos->x + dimensions;
		tripos[4].y = pos->y + midh;
		tripos[4].z = pos->z + var;
		tripos[5].x = pos->x + dimensions;
		tripos[5].y = pos->y + midh;
		tripos[5].z = pos->z - var;

		CollisionModel3D* model4= newCollisionModel3D(TRUE);

		model4->setTriangleNumber (2);

		for (i=0;i<6;i+=3) {
			model4->addTriangle (tripos[i].x,tripos[i].y,tripos[i].z,tripos[i+1].x,tripos[i+1].y,tripos[i+1].z,tripos[i+2].x,tripos[i+2].y,tripos[i+2].z);
		}
		model4->finalize ();

		if (model4->collision (model)==TRUE) {

			model4->getCollidingTriangles (firstt, secondt);

			D3DVECTOR a,b,c,h,i,n;
			a=tri[tris[secondt]];
			b=tri[tris[secondt]+1];
			c=tri[tris[secondt]+2];

			h.x = b.x - a.x;
			h.y = b.y - a.y;
			h.z = b.z - a.z;

			i.x = c.x - a.x;
			i.y = c.y - a.y;
			i.z = c.z - a.z;

			n=CrossProduct (h,i);
			n=Normalize(n);
			n.x = n.x * dimensions;
			n.y = n.y * dimensions;
			n.z = n.z * dimensions;

			float pn1[3],pn2[3];
			pn1[0]=pos->x+n.x;
			pn1[1]=pos->y+midh+n.y;
			pn1[2]=pos->z+n.z;
			pn2[0]=-n.x;
			pn2[1]=-n.y;
			pn2[2]=-n.z;

			float point5[3];
			model->getCollisionPoint (point5);
			pos->x = point5[0]+n.x;
			pos->z = point5[2]+n.z;
		}

		delete model4;
	}

	if ( (pos->x-pos2->x)*(pos->x-pos2->x) + (pos->y-pos2->y)*(pos->y-pos2->y) + (pos->z-pos2->z)*(pos->z-pos2->z) > 4) {
		pos->x=pos2->x;
		pos->z=pos2->z;
	}

	if (cameratype==1) {
		vec.x = pos->x - cam->x;
		vec.y = pos->y - cam->y;
		vec.z = pos->z - cam->z;

		vec.x = vec.x / distance;
		vec.y = vec.y / distance;
		vec.z = vec.z / distance;

		point1[0] = pos->x;
		point1[1] = pos->y+hih;
		point1[2] = pos->z;

		point2[0] = cam->x - point1[0];
		point2[1] = cam->y - point1[1];
		point2[2] = cam->z - point1[2];
		
		chocan=FALSE;

		chocan=model->rayCollision (point1,point2,true,0.0f);

		float point[3];
		BOOL camb = FALSE;
		double dis2;

		upv->x=0;upv->z=0;upv->y=1;

		if (chocan == TRUE) {
			model->getCollisionPoint (point);
			collx = point[0];colly = point[1];collz = point[2];

			dis2 = sqrt( (pos->x-collx) * (pos->x-collx) + (pos->y-colly) * (pos->y-colly) + (pos->z-collz) * (pos->z-collz));

			if ((disfromcol + distance) > dis2) {
				cam->x = collx + vec.x * disfromcol;
				cam->z = collz + vec.z * disfromcol;
				
				point1[0] = cam->x;
				point1[1] = pos->y+hih;
				point1[2] = cam->z;

				point2[0] = 0;
				point2[1] = 1;
				point2[2] = 0;

				camb=model->rayCollision(point1,point2,true);

				if (camb==TRUE) {
				model->getCollisionPoint (point);
				if (point[1]<cam->y) {
					cam->y = colly - vec.y * disfromcol;

					point1[0] = pos->x;
					point1[1] = pos->y+hih;
					point1[2] = pos->z;

					point2[0] = cam->x - point1[0];
					point2[1] = cam->y - point1[1];
					point2[2] = cam->z - point1[2];

					if (model->rayCollision (point1,point2,true,0.0f) == TRUE) {
						model->getCollisionPoint (point);
						cam->x = point[0] + vec.x * disfromcol;
						cam->z = point[2] + vec.z * disfromcol;
						cam->y = point[1] + vec.y * disfromcol;
					}
				}
				}
			}
		}

				

		double interpolation = interpolationspeed * avrframe / 1000;

		cam->x = (cam->x-cam_original.x) * interpolation + cam_original.x;
		cam->y = (cam->y-cam_original.y) * interpolation + cam_original.y;
		cam->z = (cam->z-cam_original.z) * interpolation + cam_original.z;

		double v1x,v1y,v2x,v2y;
		double angulo;

		v1x = cam->x - pos->x;
		v1y = cam->z - pos->z;

		v2x = sin(angleh * pi / 180);
		v2y = cos(angleh * pi / 180);

		angulo = acos( (v1x * v2x + v1y * v2y)/( sqrt(v1x*v1x+v1y*v1y) * sqrt(v2x*v2x+v2y*v2y) ));

		if ((angulo / pi * 180) > 135) { upv->y = -1; }

	}

	delete model;delete model2;

	return 0;
}


/*
	 Processa el moviment vertical
*/


__declspec( dllexport ) _stdcall hprocess(D3DVECTOR *tri, D3DVECTOR *point, long numtri, salida *collide)
{
	CollisionModel3D* model= newCollisionModel3D();

	double 	pi = 3.14159265358979;

	model->setTriangleNumber (numtri);
	
	long i;
	BOOL chocan=FALSE;
	double var=point->y;


	/* Optimitació! Es descarten triangles */
	for (i=0;i<(numtri*3);i=i+3) {
		if ( (tri[i].y < var) || (tri[i+1].y < var) || (tri[i+2].y < var) ) {
		if ( !( (tri[i].x < point->x) && (tri[i+1].x < point->x) && (tri[i+2].x < point->x) ) ) {
		if ( !( (tri[i].x > point->x) && (tri[i+1].x > point->x) && (tri[i+2].x > point->x) ) ) {
		if ( !( (tri[i].z < point->z) && (tri[i+1].z < point->z) && (tri[i+2].z < point->z) ) ) {
		if ( !( (tri[i].z > point->z) && (tri[i+1].z > point->z) && (tri[i+2].z > point->z) ) ) {
			model->addTriangle (tri[i].x,tri[i].y,tri[i].z,tri[i+1].x,tri[i+1].y,tri[i+1].z,tri[i+2].x,tri[i+2].y,tri[i+2].z);
		} } } } }
	}

	model->finalize ();

	float point1[3],point2[3];
	float outpoint[3];
	point2[0]=0;
	point2[1]=-1;
	point2[2]=0;

	point1[0]=point->x;
	point1[1]=point->y;
	point1[2]=point->z;

	chocan=model->rayCollision (point1,point2,true,0.0f);

	if (chocan == TRUE) {
		collide->respuesta=1;
		model->getCollisionPoint (outpoint);
		collide->puntocolision.x = point->x;
		collide->puntocolision.y = outpoint[1];
		collide->puntocolision.z = point->z;

	}else{ collide->respuesta =0; }

	delete model;

	return 0;
}

/*
	 Col·lisió segment - triangles
*/

__declspec( dllexport ) _stdcall segintersect(D3DVECTOR *origin, D3DVECTOR *direction, D3DVECTOR *tri, long numtri, long numsegs, salida *collide)
{
	long i;
	BOOL chocan=FALSE;
	CollisionModel3D* model= newCollisionModel3D();

	model->setTriangleNumber (numtri);

	for (i=0;i<(numtri*3);i=i+3) {
		model->addTriangle (tri[i].x,tri[i].y,tri[i].z,tri[i+1].x,tri[i+1].y,tri[i+1].z,tri[i+2].x,tri[i+2].y,tri[i+2].z);
	}

	model->finalize ();

	float point1[3],point2[3],outpoint[3];

	for (i=0;i<(numsegs);i=i+1) {
		point1[0]=origin[i].x;
		point1[1]=origin[i].y;
		point1[2]=origin[i].z;

		point2[0]=direction[i].x;
		point2[1]=direction[i].y;
		point2[2]=direction[i].z;

		chocan=model->rayCollision (point1,point2,true,0.0f);

		if (chocan == TRUE) {
			collide[i].respuesta=1;
			model->getCollisionPoint (outpoint);
			collide[i].puntocolision.x = outpoint[0];
			collide[i].puntocolision.y = outpoint[1];
			collide[i].puntocolision.z = outpoint[2];

		}else{ collide[i].respuesta =0; }
	}

	delete model;
}

/*
	 Col·lisió segments - triangles
*/

__declspec( dllexport ) _stdcall segintersectfast(D3DVECTOR *origin, D3DVECTOR *direction, D3DVECTOR *tri, long numtri, long numsegs, salida *collide)
{
	CollisionModel3D* model= newCollisionModel3D();

	model->setTriangleNumber (numtri);
	
	long i; long j;
	BOOL chocan=FALSE;

	for (i=0;i<(numtri*3);i=i+3) {
		for (j=0;j<numsegs;j=j+1) {
			/* Optimitació! Es descarten triangles */
			if ( (tri[i].y < origin[j].y) || (tri[i+1].y < origin[j].y) || (tri[i+2].y < origin[j].y) ) {
			if ( !( (tri[i].x < origin[j].x) && (tri[i+1].x < origin[j].x) && (tri[i+2].x < origin[j].x) ) ) {
			if ( !( (tri[i].x > origin[j].x) && (tri[i+1].x > origin[j].x) && (tri[i+2].x > origin[j].x) ) ) {
			if ( !( (tri[i].z < origin[j].z) && (tri[i+1].z < origin[j].z) && (tri[i+2].z < origin[j].z) ) ) {
			if ( !( (tri[i].z > origin[j].z) && (tri[i+1].z > origin[j].z) && (tri[i+2].z > origin[j].z) ) ) {
				model->addTriangle (tri[i].x,tri[i].y,tri[i].z,tri[i+1].x,tri[i+1].y,tri[i+1].z,tri[i+2].x,tri[i+2].y,tri[i+2].z);
			} } } } }
		}
	}

	model->finalize ();

	float point1[3],point2[3],outpoint[3];

	for (i=0;i<(numsegs);i=i+1) {
		point1[0]=origin[i].x;
		point1[1]=origin[i].y;
		point1[2]=origin[i].z;

		point2[0]=direction[i].x;
		point2[1]=direction[i].y;
		point2[2]=direction[i].z;

		chocan=model->rayCollision (point1,point2,true,0.0f);

		if (chocan == TRUE) {
			collide[i].respuesta=1;
			model->getCollisionPoint (outpoint);
			collide[i].puntocolision.x = outpoint[0];
			collide[i].puntocolision.y = outpoint[1];
			collide[i].puntocolision.z = outpoint[2];

		}else{ collide[i].respuesta =0; }
	}

	delete model;

}

/*
	 Test de visibilitat - Col·lisió de segments
*/

__declspec( dllexport ) _stdcall visible(D3DVECTOR *pointfrom, D3DVECTOR *pointto, D3DVECTOR *tri, long numtri, long numsegs, salida *collide)
{
	CollisionModel3D* model= newCollisionModel3D();

	model->setTriangleNumber (numtri);
	
	long i; long j;
	float xmax, xmin, zmax, zmin, ymax, ymin;
	BOOL chocan=FALSE;

	for (i=0;i<(numtri*3);i=i+3) {
		
		for (j=0;j<numsegs;j=j+1) {
			if (pointfrom[j].x > pointto[j].x) { xmax=pointfrom[j].x; xmin=pointto[j].x; }else{ xmax=pointto[j].x; xmin=pointfrom[j].x; }
			if (pointfrom[j].y > pointto[j].y) { ymax=pointfrom[j].y; ymin=pointto[j].y; }else{ ymax=pointto[j].y; ymin=pointfrom[j].y; }
			if (pointfrom[j].z > pointto[j].z) { zmax=pointfrom[j].z; zmin=pointto[j].z; }else{ zmax=pointto[j].z; zmin=pointfrom[j].z; }

			/* Optimitació! Es descarten triangles */
			if ( !( (tri[i].x > xmax) && (tri[i+1].x > xmax) && (tri[i+2].x > xmax) ) ) {
			if ( !( (tri[i].x < xmin) && (tri[i+1].x < xmin) && (tri[i+2].x < xmin) ) ) {

			if ( !( (tri[i].y > ymax) && (tri[i+1].y > ymax) && (tri[i+2].y > ymax) ) ) {
			if ( !( (tri[i].y < ymin) && (tri[i+1].y < ymin) && (tri[i+2].y < ymin) ) ) {

			if ( !( (tri[i].z > zmax) && (tri[i+1].z > zmax) && (tri[i+2].z > zmax) ) ) {
			if ( !( (tri[i].z < zmin) && (tri[i+1].z < zmin) && (tri[i+2].z < zmin) ) ) {

				model->addTriangle (tri[i].x,tri[i].y,tri[i].z,tri[i+1].x,tri[i+1].y,tri[i+1].z,tri[i+2].x,tri[i+2].y,tri[i+2].z);
			} } } } } }
		}
	}

	model->finalize ();

	float point1[3],point2[3],outpoint[3];

	for (i=0;i<(numsegs);i=i+1) {
		point1[0]=pointfrom[i].x;
		point1[1]=pointfrom[i].y;
		point1[2]=pointfrom[i].z;

		point2[0]=pointto[i].x - pointfrom[i].x;
		point2[1]=pointto[i].y - pointfrom[i].y;
		point2[2]=pointto[i].z - pointfrom[i].z;

		chocan=model->rayCollision (point1,point2,true,0.0f);

		if (chocan == TRUE) {
			collide[i].respuesta=1;
			model->getCollisionPoint (outpoint);
			collide[i].puntocolision.x = outpoint[0];
			collide[i].puntocolision.y = outpoint[1];
			collide[i].puntocolision.z = outpoint[2];

		}else{ collide[i].respuesta =0; }
	}

	delete model;

}

/*
	 Càlcul de normals d'una cara triangular
*/

__declspec( dllexport ) _stdcall computenormals(D3DVECTOR *verts, long numverts, D3DVECTOR *normals)
{
	long i, tric = 0;
	D3DVECTOR vc1, vc2;

	for (i=0 ; i<numverts; i=i+3 )
	{
		vc1 = resta(verts[i+2],verts[i+1]);
		vc2 = resta(verts[i+1],verts[i]);

		normals[tric++] = Normalize(CrossProduct(vc1,vc2));
	}
}


VOID AddEdge( long* pEdges, long& dwNumEdges, long v0, long v1 )
{
    // Remove interior edges (which appear in the list twice)
    for( long i=0; i < dwNumEdges; i++ )
    {
        if( ( pEdges[2*i+0] == v0 && pEdges[2*i+1] == v1 ) ||
            ( pEdges[2*i+0] == v1 && pEdges[2*i+1] == v0 ) )
        {
            if( dwNumEdges > 1 )
            {
                pEdges[2*i+0] = pEdges[2*(dwNumEdges-1)+0];
                pEdges[2*i+1] = pEdges[2*(dwNumEdges-1)+1];
            }
            dwNumEdges--;
            return;
        }
    }

    pEdges[2*dwNumEdges+0] = v0;
    pEdges[2*dwNumEdges+1] = v1;
    dwNumEdges++;
}

/*
	 Creació de la projecció d'una ombra a partir del càlcul de la seva silueta
*/

__declspec( dllexport ) _stdcall extrudeshadow(long proj, D3DVECTOR *tris, long numtri, D3DVECTOR *light, D3DVECTOR *trisout, long& numtrisout)
{
		long* pEdges = new long[numtri*6];
		long dwNumEdges = 0;

		D3DVECTOR lightbig;
		lightbig = Normalize(*light);
		lightbig.x = lightbig.x * proj;
		lightbig.y = lightbig.y * proj;
		lightbig.z = lightbig.z * proj;

		long i;
		numtrisout=0;

		D3DVECTOR vnormal, vc1, vc2;

		for (i=0; i<numtri; i=i+3) 
		{
			vc1 = resta(tris[i+2],tris[i+1]);
			vc2 = resta(tris[i+1],tris[i]);

			vnormal = Normalize(CrossProduct(vc1,vc2));

			if (DotProduct(vnormal,lightbig)>=0) {
				AddEdge (pEdges, dwNumEdges, i , i+1 );
				AddEdge (pEdges, dwNumEdges, i+1 , i+2 );
				AddEdge (pEdges, dwNumEdges, i+2 , i );

				vnormal = multiplica(vnormal, -0.002);

				tris[i]=resta(tris[i],vnormal);
				tris[i+1]=resta(tris[i+1],vnormal);
				tris[i+2]=resta(tris[i+2],vnormal);

				trisout[numtrisout++]=resta(tris[i],lightbig);
				trisout[numtrisout++]=resta(tris[i+1],lightbig);
				trisout[numtrisout++]=resta(tris[i+2],lightbig);

				trisout[numtrisout++]=copyv(tris[i+2]);
				trisout[numtrisout++]=copyv(tris[i+1]);
				trisout[numtrisout++]=copyv(tris[i]);
			}
		}

		for (i=0; i<dwNumEdges; i++) 
		{
			D3DVECTOR vec1,vec2,vec3,vec4;
			vec1 = copyv(tris[pEdges[2*i]]);
			vec2 = copyv(tris[pEdges[2*i+1]]);
			vec3 = resta(vec1,lightbig);
			vec4 = resta(vec2,lightbig);

			trisout[numtrisout++]=vec1;
			trisout[numtrisout++]=vec2;
			trisout[numtrisout++]=vec3;

			trisout[numtrisout++]=vec2;
			trisout[numtrisout++]=vec4;
			trisout[numtrisout++]=vec3;
		}

		delete pEdges;
}

/*
	 Canvi de format de pixel de textura de R8G8B8 a X8R8G8B8
*/

__declspec( dllexport ) _stdcall formalisetex(unsigned char *entrada, long width, long height, long pitch, unsigned char *salida)
{
	long x, y;
	unsigned char p1, p2, p3;
	for (x=0; x<height; x++) 
	{
		for (y=0; y<width; y++) 
		{
			p1 = entrada[(width * 3 * (height - 1 - x)) + y * 3];
			p2 = entrada[(width * 3 * (height - 1 - x)) + y * 3+1];
			p3 = entrada[(width * 3 * (height - 1 - x)) + y * 3+2];

			salida[x * pitch + y * 4] = p1;
			salida[x * pitch + y * 4+1] = p2;
			salida[x * pitch + y * 4+2] = p3;
			salida[x * pitch + y * 4+3] = 0;
		}
	}

}

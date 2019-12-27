/*   ColDet - C++ 3D Collision Detection Library
 *   Copyright (C) 2000   Amir Geva
 */
#include "sysdep.h"
#include "coldetimpl.h"

__CD__BEGIN

CollisionModel3D* newCollisionModel3D(bool Static)
{
  return new CollisionModel3DImpl(Static);
}

CollisionModel3DImpl::CollisionModel3DImpl(bool Static)
: m_Root(Vector3D::Zero, Vector3D::Zero,0),
  m_Transform(Matrix3D::Identity),
  m_InvTransform(Matrix3D::Identity),
  m_ColTri1(Vector3D::Zero,Vector3D::Zero,Vector3D::Zero),
  m_ColTri2(Vector3D::Zero,Vector3D::Zero,Vector3D::Zero),
  m_iColTri1(0),
  m_iColTri2(0),
  m_Final(false),
  m_Static(Static)
{}

void CollisionModel3DImpl::addTriangle(const Vector3D& v1, const Vector3D& v2, const Vector3D& v3)
{
  if (m_Final) throw Inconsistency();
  m_Triangles.push_back(BoxedTriangle(v1,v2,v3));
}

void CollisionModel3DImpl::setTransform(const Matrix3D& m)
{
  m_Transform=m;
  if (m_Static) m_InvTransform=m_Transform.Inverse();
}

void CollisionModel3DImpl::finalize()
{
  if (m_Final) throw Inconsistency();
  // Prepare initial triangle list
  m_Final=true;
  for(unsigned i=0;i<m_Triangles.size();i++)
  {
    BoxedTriangle& bt=m_Triangles[i];
    m_Root.m_Boxes.push_back(&bt);
  }
  int logdepth=0;
  for(int num=m_Triangles.size();num>0;num>>=1,logdepth++);
  m_Root.m_logdepth=int(logdepth*1.5f);
  m_Root.divide(0);
}

__CD__END

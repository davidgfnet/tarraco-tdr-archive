/*   ColDet - C++ 3D Collision Detection Library
 *   Copyright (C) 2000   Amir Geva
 */
#include "sysdep.h"

#ifdef GCC

#include <sys/time.h>

// Returns a time index in milliseconds
DWORD GetTickCount()
{
  static struct timezone tz={0,0};
  static const double t1=1000.0;
  static const double t2=0.001;
  timeval t;
  gettimeofday(&t,&tz);
  return long((t.tv_sec&0x000FFFFF)*t1 + t.tv_usec*t2);
}

#endif

/* this file contains the actual definitions of */
/* the IIDs and CLSIDs */

/* link this file in with the server and any clients */


/* File created by MIDL compiler version 5.01.0164 */
/* at Sat Jun 09 23:03:46 2007
 */
/* Compiler settings for C:\Documents and Settings\David\Mis documentos\gameproject\cal3dvb\cal3dvb.idl:
    Oicf (OptLev=i2), W1, Zp8, env=Win32, ms_ext, c_ext
    error checks: allocation ref bounds_check enum stub_data 
*/
//@@MIDL_FILE_HEADING(  )
#ifdef __cplusplus
extern "C"{
#endif 


#ifndef __IID_DEFINED__
#define __IID_DEFINED__

typedef struct _IID
{
    unsigned long x;
    unsigned short s1;
    unsigned short s2;
    unsigned char  c[8];
} IID;

#endif // __IID_DEFINED__

#ifndef CLSID_DEFINED
#define CLSID_DEFINED
typedef IID CLSID;
#endif // CLSID_DEFINED

const IID IID_ICal3DObject = {0x01E7F1E4,0xD966,0x46A2,{0x9D,0x48,0x78,0x27,0x3F,0xA6,0xB8,0x6F}};


const IID LIBID_CAL3DVBLib = {0x0FEE9E62,0x3C15,0x448D,{0xAD,0xA0,0x5D,0x79,0xC8,0x0E,0x39,0x96}};


const CLSID CLSID_Cal3DObject = {0x6F1BCB9E,0x077E,0x4F5C,{0x87,0x32,0x9B,0x36,0x1F,0x77,0x63,0xE9}};


#ifdef __cplusplus
}
#endif


Attribute VB_Name = "D3DShaders"

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
'  Copyright (C) 1999-2001 Microsoft Corporation.  All Rights Reserved.
'
'  File:       D3DShader.bas
'  Content:    Shader constants
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Option Explicit

Global Const D3DSI_COMMENTSIZE_SHIFT = 16
Global Const D3DSI_COMMENTSIZE_MASK = &H7FFF0000
Global Const D3DVS_INPUTREG_MAX_V1_1 = 16
Global Const D3DVS_TEMPREG_MAX_V1_1 = 12
Global Const D3DVS_CONSTREG_MAX_V1_1 = 96
Global Const D3DVS_TCRDOUTREG_MAX_V1_1 = 8
Global Const D3DVS_ADDRREG_MAX_V1_1 = 1
Global Const D3DVS_ATTROUTREG_MAX_V1_1 = 2
Global Const D3DVS_MAXINSTRUCTIONCOUNT_V1_1 = 128
Global Const D3DPS_INPUTREG_MAX_DX8 = 8
Global Const D3DPS_TEMPREG_MAX_DX8 = 8
Global Const D3DPS_CONSTREG_MAX_DX8 = 8
Global Const D3DPS_TEXTUREREG_MAX_DX8 = 8


Enum D3DVSD_TOKENTYPE
    D3DVSD_TOKEN_NOP = 0
    D3DVSD_TOKEN_STREAM = 1
    D3DVSD_TOKEN_STREAMDATA = 2
    D3DVSD_TOKEN_TESSELLATOR = 3
    D3DVSD_TOKEN_constMEM = 4
    D3DVSD_TOKEN_EXT = 5
    D3DVSD_TOKEN_END = 7
End Enum

Global Const D3DVSD_TOKENTYPESHIFT = 29
Global Const D3DVSD_TOKENTYPEMASK = &HE0000000
Global Const D3DVSD_STREAMNUMBERSHIFT = 0
Global Const D3DVSD_STREAMNUMBERMASK = &HF&
Global Const D3DVSD_DATALOADTYPESHIFT = 28
Global Const D3DVSD_DATALOADTYPEMASK = &H10000000
Global Const D3DVSD_DATATYPESHIFT = 16
Global Const D3DVSD_DATATYPEMASK = &HF& * 2 ^ D3DVSD_DATATYPESHIFT
Global Const D3DVSD_SKIPCOUNTSHIFT = 16
Global Const D3DVSD_SKIPCOUNTMASK = &HF& * 2 ^ D3DVSD_SKIPCOUNTSHIFT
Global Const D3DVSD_VERTEXREGSHIFT = 0
Global Const D3DVSD_VERTEXREGMASK = &HF& * 2 ^ D3DVSD_VERTEXREGSHIFT
Global Const D3DVSD_VERTEXREGINSHIFT = 20
Global Const D3DVSD_VERTEXREGINMASK = &HF& * 2 ^ D3DVSD_VERTEXREGINSHIFT
Global Const D3DVSD_CONSTCOUNTSHIFT = 25
Global Const D3DVSD_CONSTCOUNTMASK = &HF& * 2 ^ D3DVSD_CONSTCOUNTSHIFT
Global Const D3DVSD_CONSTADDRESSSHIFT = 0
Global Const D3DVSD_CONSTADDRESSMASK = &H7F&
Global Const D3DVSD_CONSTRSSHIFT = 16
Global Const D3DVSD_CONSTRSMASK = &H1FFF0000
Global Const D3DVSD_EXTCOUNTSHIFT = 24
Global Const D3DVSD_EXTCOUNTMASK = &H1F& * 2 ^ D3DVSD_EXTCOUNTSHIFT
Global Const D3DVSD_EXTINFOSHIFT = 0
Global Const D3DVSD_EXTINFOMASK = &HFFFFFF
Global Const D3DVSDT_FLOAT1 = 0&
Global Const D3DVSDT_FLOAT2 = 1&
Global Const D3DVSDT_FLOAT3 = 2&
Global Const D3DVSDT_FLOAT4 = 3&
Global Const D3DVSDT_D3DCOLOR = 4&
Global Const D3DVSDT_UBYTE4 = 5&
Global Const D3DVSDT_SHORT2 = 6&
Global Const D3DVSDT_SHORT4 = 7&
Global Const D3DVSDE_POSITION = 0&
Global Const D3DVSDE_BLENDWEIGHT = 1&
Global Const D3DVSDE_BLENDINDICES = 2&
Global Const D3DVSDE_NORMAL = 3&
Global Const D3DVSDE_PSIZE = 4&
Global Const D3DVSDE_DIFFUSE = 5&
Global Const D3DVSDE_SPECULAR = 6&
Global Const D3DVSDE_TEXCOORD0 = 7&
Global Const D3DVSDE_TEXCOORD1 = 8&
Global Const D3DVSDE_TEXCOORD2 = 9&
Global Const D3DVSDE_TEXCOORD3 = 10&
Global Const D3DVSDE_TEXCOORD4 = 11&
Global Const D3DVSDE_TEXCOORD5 = 12&
Global Const D3DVSDE_TEXCOORD6 = 13&
Global Const D3DVSDE_TEXCOORD7 = 14&
Global Const D3DVSDE_POSITION2 = 15&
Global Const D3DVSDE_NORMAL2 = 16&

' Maximum supported number of texture coordinate sets
Global Const D3DDP_MAXTEXCOORD = 8

Global Const D3DSI_OPCODE_MASK = &HFFFF&

Enum D3DSHADER_INSTRUCTION_OPCODE_TYPE
    D3DSIO_NOP = 0            ' PS/VS
    D3DSIO_MOV = 1                ' PS/VS
    D3DSIO_ADD = 2                ' PS/VS
    D3DSIO_SUB = 3                ' PS
    D3DSIO_MAD = 4                ' PS/VS
    D3DSIO_MUL = 5                ' PS/VS
    D3DSIO_RCP = 6                ' VS
    D3DSIO_RSQ = 7                ' VS
    D3DSIO_DP3 = 8                ' PS/VS
    D3DSIO_DP4 = 9                ' VS
    D3DSIO_MIN = 10                ' VS
    D3DSIO_MAX = 11                ' VS
    D3DSIO_SLT = 12                ' VS
    D3DSIO_SGE = 13                ' VS
    D3DSIO_EXP = 14                ' VS
    D3DSIO_LOG = 15                ' VS
    D3DSIO_LIT = 16                ' VS
    D3DSIO_DST = 17                ' VS
    D3DSIO_LRP = 18                ' PS
    D3DSIO_FRC = 19                ' VS
    D3DSIO_M4x4 = 20               ' VS
    D3DSIO_M4x3 = 21               ' VS
    D3DSIO_M3x4 = 22               ' VS
    D3DSIO_M3x3 = 23               ' VS
    D3DSIO_M3x2 = 24               ' VS

    D3DSIO_TEXCOORD = 64           ' PS
    D3DSIO_TEXKILL = 65            ' PS
    D3DSIO_TEX = 66                ' PS
    D3DSIO_TEXBEM = 67             ' PS
    D3DSIO_TEXBEML = 68            ' PS
    D3DSIO_TEXREG2AR = 69          ' PS
    D3DSIO_TEXREG2GB = 70          ' PS
    D3DSIO_TEXM3x2PAD = 71         ' PS
    D3DSIO_TEXM3x2TEX = 72         ' PS
    D3DSIO_TEXM3x3PAD = 73         ' PS
    D3DSIO_TEXM3x3TEX = 74         ' PS
    D3DSIO_TEXM3x3DIFF = 75        ' PS
    D3DSIO_TEXM3x3SPEC = 76        ' PS
    D3DSIO_TEXM3x3VSPEC = 77       ' PS
    D3DSIO_EXPP = 78               ' VS
    D3DSIO_LOGP = 79               ' VS
    D3DSIO_CND = 80                ' PS
    D3DSIO_DEF = 81                ' PS

    D3DSIO_COMMENT = &HFFFE&
    D3DSIO_END = &HFFFF&
End Enum

Global Const D3DSI_COISSUE = &H40000000
Global Const D3DSP_REGNUM_MASK = &HFFF&
Global Const D3DSP_WRITEMASK_0 = &H10000
Global Const D3DSP_WRITEMASK_1 = &H20000
Global Const D3DSP_WRITEMASK_2 = &H40000
Global Const D3DSP_WRITEMASK_3 = &H80000
Global Const D3DSP_WRITEMASK_ALL = &HF0000
Global Const D3DSP_DSTMOD_SHIFT = 20
Global Const D3DSP_DSTMOD_MASK = &HF00000

Enum D3DSHADER_PARAM_DSTMOD_TYPE

    D3DSPDM_NONE = 0 * 2 ^ D3DSP_DSTMOD_SHIFT
    D3DSPDM_SATURATE = 1 * 2 ^ D3DSP_DSTMOD_SHIFT

End Enum

Global Const D3DSP_DSTSHIFT_SHIFT = 24
Global Const D3DSP_DSTSHIFT_MASK = &HF000000
Global Const D3DSP_REGTYPE_SHIFT = 28
Global Const D3DSP_REGTYPE_MASK = &H70000000
Global Const D3DVSD_STREAMTESSSHIFT = 28
Global Const D3DVSD_STREAMTESSMASK = 2 ^ D3DVSD_STREAMTESSSHIFT

Enum D3DSHADER_PARAM_REGISTER_TYPE

    D3DSPR_TEMP = &H0&
    D3DSPR_INPUT = &H20000000
    D3DSPR_CONST = &H40000000
    D3DSPR_ADDR = &H60000000
    D3DSPR_TEXTURE = &H60000000
    D3DSPR_RASTOUT = &H80000000
    D3DSPR_ATTROUT = &HA0000000
    D3DSPR_TEXCRDOUT = &HC0000000
End Enum

Enum D3DVS_RASTOUT_OFFSETS
    D3DSRO_POSITION = 0
    D3DSRO_FOG = 1
    D3DSRO_POINT_SIZE = 2
End Enum

Global Const D3DVS_ADDRESSMODE_SHIFT = 13
Global Const D3DVS_ADDRESSMODE_MASK = (2 ^ D3DVS_ADDRESSMODE_SHIFT)

Enum D3DVS_ADRRESSMODE_TYPE
    D3DVS_ADDRMODE_ABSOLUTE = 0
    D3DVS_ADDRMODE_RELATIVE = 2 ^ D3DVS_ADDRESSMODE_SHIFT
End Enum

Global Const D3DVS_SWIZZLE_SHIFT = 16
Global Const D3DVS_SWIZZLE_MASK = &HFF0000
Global Const D3DVS_X_X = (0 * 2 ^ D3DVS_SWIZZLE_SHIFT)
Global Const D3DVS_X_Y = (1 * 2 ^ D3DVS_SWIZZLE_SHIFT)
Global Const D3DVS_X_Z = (2 * 2 ^ D3DVS_SWIZZLE_SHIFT)
Global Const D3DVS_X_W = (3 * 2 ^ D3DVS_SWIZZLE_SHIFT)
Global Const D3DVS_Y_X = (0 * 2 ^ (D3DVS_SWIZZLE_SHIFT + 2))
Global Const D3DVS_Y_Y = (1 * 2 ^ (D3DVS_SWIZZLE_SHIFT + 2))
Global Const D3DVS_Y_Z = (2 * 2 ^ (D3DVS_SWIZZLE_SHIFT + 2))
Global Const D3DVS_Y_W = (3 * 2 ^ (D3DVS_SWIZZLE_SHIFT + 2))
Global Const D3DVS_Z_X = (0 * 2 ^ (D3DVS_SWIZZLE_SHIFT + 4))
Global Const D3DVS_Z_Y = (1 * 2 ^ (D3DVS_SWIZZLE_SHIFT + 4))
Global Const D3DVS_Z_Z = (2 * 2 ^ (D3DVS_SWIZZLE_SHIFT + 4))
Global Const D3DVS_Z_W = (3 * 2 ^ (D3DVS_SWIZZLE_SHIFT + 4))
Global Const D3DVS_W_X = (0 * 2 ^ (D3DVS_SWIZZLE_SHIFT + 6))
Global Const D3DVS_W_Y = (1 * 2 ^ (D3DVS_SWIZZLE_SHIFT + 6))
Global Const D3DVS_W_Z = (2 * 2 ^ (D3DVS_SWIZZLE_SHIFT + 6))
Global Const D3DVS_W_W = (3 * 2 ^ (D3DVS_SWIZZLE_SHIFT + 6))
Global Const D3DVS_NOSWIZZLE = (D3DVS_X_X Or D3DVS_Y_Y Or D3DVS_Z_Z Or D3DVS_W_W)
Global Const D3DSP_SWIZZLE_SHIFT = 16
Global Const D3DSP_SWIZZLE_MASK = &HFF0000

Global Const D3DSP_NOSWIZZLE = _
    ((0 * 2 ^ (D3DSP_SWIZZLE_SHIFT + 0)) Or _
      (1 * 2 ^ (D3DSP_SWIZZLE_SHIFT + 2)) Or _
      (2 * 2 ^ (D3DSP_SWIZZLE_SHIFT + 4)) Or _
      (3 * 2 ^ (D3DSP_SWIZZLE_SHIFT + 6)))

Global Const D3DSP_REPLICATEALPHA = _
    ((3 * 2 ^ (D3DSP_SWIZZLE_SHIFT + 0)) Or _
      (3 * 2 ^ (D3DSP_SWIZZLE_SHIFT + 2)) Or _
      (3 * 2 ^ (D3DSP_SWIZZLE_SHIFT + 4)) Or _
      (3 * 2 ^ (D3DSP_SWIZZLE_SHIFT + 6)))
      
Global Const D3DSP_SRCMOD_SHIFT = 24
Global Const D3DSP_SRCMOD_MASK = &HF000000

Enum D3DSHADER_PARAM_SRCMOD_TYPE
    D3DSPSM_NONE = 0 * 2 ^ D3DSP_SRCMOD_SHIFT    '0<<D3DSP_SRCMOD_SHIFT, ' nop
    D3DSPSM_NEG = 1 * 2 ^ D3DSP_SRCMOD_SHIFT     ' negate
    D3DSPSM_BIAS = 2 * 2 ^ D3DSP_SRCMOD_SHIFT    ' bias
    D3DSPSM_BIASNEG = 3 * 2 ^ D3DSP_SRCMOD_SHIFT ' bias and negate
    D3DSPSM_SIGN = 4 * 2 ^ D3DSP_SRCMOD_SHIFT    ' sign
    D3DSPSM_SIGNNEG = 5 * 2 ^ D3DSP_SRCMOD_SHIFT ' sign and negate
    D3DSPSM_COMP = 6 * 2 ^ D3DSP_SRCMOD_SHIFT    ' complement
End Enum

Function D3DPS_VERSION(Major As Long, Minor As Long) As Long
    D3DPS_VERSION = (&HFFFF0000 Or ((Major) * 2 ^ 8) Or (Minor))
End Function

Function D3DVS_VERSION(Major As Long, Minor As Long) As Long
    D3DVS_VERSION = (&HFFFE0000 Or ((Major) * 2 ^ 8) Or (Minor))
End Function

Function D3DSHADER_VERSION_MAJOR(version As Long) As Long
    D3DSHADER_VERSION_MAJOR = (((version) \ 8) And &HFF&)
End Function
    
Function D3DSHADER_VERSION_MINOR(version As Long) As Long
    D3DSHADER_VERSION_MINOR = (((version)) And &HFF&)
End Function
    
Function D3DSHADER_COMMENT(DWordSize As Long) As Long
    D3DSHADER_COMMENT = ((((DWordSize) * 2 ^ D3DSI_COMMENTSIZE_SHIFT) And D3DSI_COMMENTSIZE_MASK) Or D3DSIO_COMMENT)
End Function
    
Function D3DPS_END() As Long
    D3DPS_END = &HFFFF&
End Function

Function D3DVS_END() As Long
   D3DVS_END = &HFFFF&
End Function

Function D3DVSD_MAKETOKENTYPE(tokenType As Long) As Long
    Dim Out As Long
    Select Case tokenType
        Case D3DVSD_TOKEN_NOP
            Out = 0
        Case D3DVSD_TOKEN_STREAM
            Out = &H20000000
        Case D3DVSD_TOKEN_STREAMDATA
            Out = &H40000000
        Case D3DVSD_TOKEN_TESSELLATOR
            Out = &H60000000
        Case D3DVSD_TOKEN_constMEM
            Out = &H80000000
        Case D3DVSD_TOKEN_EXT
            Out = &HA0000000
        Case D3DVSD_TOKEN_END
            Out = &HFFFFFFFF
    End Select
    D3DVSD_MAKETOKENTYPE = Out And D3DVSD_TOKENTYPEMASK
End Function

Function D3DVSD_STREAM(StreamNumber As Long) As Long
    D3DVSD_STREAM = (D3DVSD_MAKETOKENTYPE(D3DVSD_TOKEN_STREAM) Or (StreamNumber))
End Function

Function D3DVSD_STREAM_TESS() As Long
    D3DVSD_STREAM_TESS = (D3DVSD_MAKETOKENTYPE(D3DVSD_TOKEN_STREAM) Or (D3DVSD_STREAMTESSMASK))
End Function

Function D3DVSD_REG(VertexRegister As Long, dataType As Long) As Long
    D3DVSD_REG = (D3DVSD_MAKETOKENTYPE(D3DVSD_TOKEN_STREAMDATA) Or _
     ((dataType) * 2 ^ D3DVSD_DATATYPESHIFT) Or (VertexRegister))
End Function

Function D3DVSD_SKIP(DWORDCount As Long) As Long
    D3DVSD_SKIP = (D3DVSD_MAKETOKENTYPE(D3DVSD_TOKEN_STREAMDATA) Or &H10000000 Or _
     ((DWORDCount) * 2 ^ D3DVSD_SKIPCOUNTSHIFT))
End Function
    
Function D3DVSD_CONST(constantAddress As Long, count As Long) As Long
    D3DVSD_CONST = (D3DVSD_MAKETOKENTYPE(D3DVSD_TOKEN_constMEM) Or _
     ((count) * 2 ^ D3DVSD_CONSTCOUNTSHIFT) Or (constantAddress))
End Function

Function D3DVSD_TESSNORMAL(VertexRegisterIn As Long, VertexRegisterOut As Long) As Long
    D3DVSD_TESSNORMAL = (D3DVSD_MAKETOKENTYPE(D3DVSD_TOKEN_TESSELLATOR) Or _
     ((VertexRegisterIn) * 2 ^ D3DVSD_VERTEXREGINSHIFT) Or _
     ((&H2&) * 2 ^ D3DVSD_DATATYPESHIFT) Or (VertexRegisterOut))
End Function
   
Function D3DVSD_TESSUV(VertexRegister As Long) As Long
    D3DVSD_TESSUV = (D3DVSD_MAKETOKENTYPE(D3DVSD_TOKEN_TESSELLATOR) Or &H10000000 Or _
     ((&H1&) * 2 ^ D3DVSD_DATATYPESHIFT) Or (VertexRegister))
End Function

Function D3DVSD_END() As Long
        D3DVSD_END = &HFFFFFFFF
End Function

Function D3DVSD_NOP() As Long
    D3DVSD_NOP = 0
End Function

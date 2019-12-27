
cal3dvbps.dll: dlldata.obj cal3dvb_p.obj cal3dvb_i.obj
	link /dll /out:cal3dvbps.dll /def:cal3dvbps.def /entry:DllMain dlldata.obj cal3dvb_p.obj cal3dvb_i.obj \
		kernel32.lib rpcndr.lib rpcns4.lib rpcrt4.lib oleaut32.lib uuid.lib \

.c.obj:
	cl /c /Ox /DWIN32 /D_WIN32_WINNT=0x0400 /DREGISTER_PROXY_DLL \
		$<

clean:
	@del cal3dvbps.dll
	@del cal3dvbps.lib
	@del cal3dvbps.exp
	@del dlldata.obj
	@del cal3dvb_p.obj
	@del cal3dvb_i.obj

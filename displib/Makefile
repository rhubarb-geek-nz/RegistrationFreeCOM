# Copyright (c) 2024 Roger Brown.
# Licensed under the MIT License.

OBJDIR=obj\$(VSCMD_ARG_TGT_ARCH)
BINDIR=bin\$(VSCMD_ARG_TGT_ARCH)
DLLNAME=displib

all: $(BINDIR)\$(DLLNAME).dll

$(OBJDIR)\$(DLLNAME).h $(OBJDIR)\$(DLLNAME)_i.c: $(OBJDIR)\$(DLLNAME).tlb

$(OBJDIR)\$(DLLNAME).tlb: $(DLLNAME).idl $(OBJDIR)
	midl $(DLLNAME).idl /client none /server none /tlb $(DLLNAME).tlb /h $(DLLNAME).h /out $(OBJDIR)

clean:
	if exist $(OBJDIR) rmdir /q /s $(OBJDIR)
	if exist $(BINDIR) rmdir /q /s $(BINDIR)

$(BINDIR)\$(DLLNAME).dll: $(OBJDIR)\$(DLLNAME).obj $(OBJDIR)\$(DLLNAME)_i.obj $(OBJDIR)\$(DLLNAME).res $(DLLNAME).def $(BINDIR)
	cl /LD /Fe$@ /MT $(OBJDIR)\$(DLLNAME).obj $(OBJDIR)\$(DLLNAME)_i.obj $(OBJDIR)\$(DLLNAME).res ole32.lib oleaut32.lib advapi32.lib /link /INCREMENTAL:NO /PDB:NONE /DEF:$(DLLNAME).def /SUBSYSTEM:CONSOLE
	if not "$(CertificateThumbprint)"=="" signtool sign /a /sha1 $(CertificateThumbprint) /fd SHA256 /t http://timestamp.digicert.com $@

$(OBJDIR)\$(DLLNAME).obj: $(OBJDIR)\$(DLLNAME).h $(DLLNAME).c
	cl /c /Fo$@ $(DLLNAME).c /I$(OBJDIR) /DCOBJMACROS /DWIN32_LEAN_AND_MEAN

$(OBJDIR)\$(DLLNAME)_i.obj: $(OBJDIR)\$(DLLNAME).h $(OBJDIR)\$(DLLNAME)_i.c
	cl /c /Fo$@ $(OBJDIR)\$(DLLNAME)_i.c /I$(OBJDIR)

$(OBJDIR)\$(DLLNAME).res: $(DLLNAME).rc $(OBJDIR)\$(DLLNAME).tlb
	rc /r  $(RCFLAGS) /i $(OBJDIR) /fo$@ $(DLLNAME).rc

$(BINDIR) $(OBJDIR):
	mkdir $@

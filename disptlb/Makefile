# Copyright (c) 2024 Roger Brown.
# Licensed under the MIT License.

OBJDIR=obj\$(VSCMD_ARG_TGT_ARCH)
BINDIR=bin\$(VSCMD_ARG_TGT_ARCH)
DLLNAME=RhubarbGeekNzRegistrationFreeCOM
TLBFILE=..\displib\$(OBJDIR)\displib.tlb

all: $(BINDIR)\$(DLLNAME).dll

clean:
	if exist $(OBJDIR) rmdir /q /s $(OBJDIR)
	if exist $(BINDIR) rmdir /q /s $(BINDIR)

$(BINDIR)\$(DLLNAME).dll: $(BINDIR) $(TLBFILE)
	tlbimp $(TLBFILE) /asmversion:1.0.6.0 /namespace:RhubarbGeekNzRegistrationFreeCOM /machine:Agnostic /out:$@
	if not "$(CertificateThumbprint)"=="" signtool sign /a /sha1 $(CertificateThumbprint) /fd SHA256 /t http://timestamp.digicert.com $@

$(BINDIR) $(OBJDIR):
	mkdir $@

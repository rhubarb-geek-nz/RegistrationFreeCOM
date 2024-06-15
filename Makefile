# Copyright (c) 2024 Roger Brown.
# Licensed under the MIT License.

BINDIR=bin\$(VSCMD_ARG_TGT_ARCH)
APPNAME=dispapp
DLLNAME=displib

all: $(BINDIR) $(BINDIR)\$(APPNAME).exe $(BINDIR)\$(DLLNAME).dll

clean:
	if exist $(BINDIR) rmdir /q /s $(BINDIR)
	cd $(APPNAME)
	$(MAKE) clean
	cd ..
	cd $(DLLNAME)
	$(MAKE) clean
	cd ..

$(BINDIR):
	mkdir $@

$(BINDIR)\$(APPNAME).exe: $(APPNAME)\$(BINDIR)\$(APPNAME).exe
	copy $(APPNAME)\$(BINDIR)\$(APPNAME).exe $@

$(BINDIR)\$(DLLNAME).dll: $(DLLNAME)\$(BINDIR)\$(DLLNAME).dll
	copy $(DLLNAME)\$(BINDIR)\$(DLLNAME).dll $@

$(APPNAME)\$(BINDIR)\$(APPNAME).exe:
	cd $(APPNAME)
	$(MAKE) CertificateThumbprint=$(CertificateThumbprint)
	cd ..

$(DLLNAME)\$(BINDIR)\$(DLLNAME).dll:
	cd $(DLLNAME)
	$(MAKE) CertificateThumbprint=$(CertificateThumbprint)
	cd ..

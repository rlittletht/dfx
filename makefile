
ODIR=ship
SRCDIR=..

CS_FLAGS=/debug:full /incr+
SHARED=..\..\shared

!ifdef DEBUG
ODIR=debug
CS_FLAGS=$(CS_FLAGS) /define:DEBUG /debug:full /debug+
!else
!endif

target: chdir dfx.exe

clean: 
	-del /q $(ODIR)\*.*
	
chdir:
	@-mkdir $(ODIR) > NUL 2>&1
	@cd $(ODIR)  
	@echo Changed directory to $(ODIR)...

AssemblyInfo.netmodule: ..\AssemblyInfo.cs
	csc $(CS_FLAGS) /target:module /out:AssemblyInfo.netmodule ..\AssemblyInfo.cs

dfx.m_frm.resources: $(SRCDIR)\dfx.resx
	resgen $(SRCDIR)\dfx.resx dfx.m_frm.resources
	
dfx.exe: AssemblyInfo.netmodule $(SRCDIR)\dfx.cs $(SRCDIR)\App.ico $(SHARED)\ole32.cs $(SHARED)\ui.cs dfx.m_frm.resources
	csc $(CS_FLAGS) /res:dfx.m_frm.resources /target:winexe /out:dfx.exe /addmodule:AssemblyInfo.netmodule $(SRCDIR)\dfx.cs $(SHARED)\ole32.cs $(SHARED)\ui.cs
	

        

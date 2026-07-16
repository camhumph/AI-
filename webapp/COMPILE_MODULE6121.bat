@echo off
rem ============================================================
rem How to rebuild Module6121.swp (required when RunMacro ok=False)
rem ============================================================
echo.
echo Rebuild Module6121.swp for SolidWorks 2023
echo.
echo 1. Open SolidWorks 2023  ^(SOLIDWORKS (3)^)
echo 2. Tools  ^>  Macro  ^>  New
echo    Save temporarily as C:\CMS_Local_Workspace\Module6121_new.swp
echo 3. In the VBA editor: File  ^>  Import File...
echo    Select:  C:\CMS_AI\Module6121.bas
echo 4. In Project Explorer, confirm the module name is Module61211
echo    ^(SolidWorks often names it Module61211 — that is correct^)
echo 5. Debug  ^>  Compile VBAProject
echo 6. File  ^>  Save  Module6121_new.swp
echo 7. Close the VBA editor
echo 8. Copy / rename to:
echo      C:\CMS_Local_Workspace\Module6121.swp
echo.
echo RunMacro calls Module61211.main inside that .swp file.
echo.
echo Then copy launchers:
echo   copy /Y C:\CMS_AI\CMS_Launcher.vbs C:\CMS_Local_Workspace\
echo   copy /Y C:\CMS_AI\RunSolidWorksMacro.ps1 C:\CMS_Local_Workspace\
echo.
echo If RunMacro still fails, check:
echo   Tools ^> Options ^> System Options ^> Macro
echo   - enable macros / add C:\CMS_Local_Workspace as trusted
echo.
pause

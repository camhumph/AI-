CMS SWP-only source package

Use Module6121.bas as the source you import into the SolidWorks macro editor.
Save the macro package as:
  C:\CMS_Local_Workspace\Module6121.swp

The launcher tries C:\CMS_Local_Workspace\Module6121.swb first so source fixes run immediately.
Module6121.swp is still kept as the fallback saved SolidWorks macro package.
Module6121.swb is a text source copy of Module6121.bas.
Visual inspection is included and runs cms_visual_inspect.ps1 after ISO/CAD outputs.
DXF export uses the saved native model view: CMS_TOP is applied and persisted to standard *Top before the drawing/DXF view is created. No manual DXF-only rotation is used.
Purchased components are captured from BOM Purchase rows and vendor/part-number fallbacks.
Pullcore cam/key naming and H-13/A-2 fallback detection are included.

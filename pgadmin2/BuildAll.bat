@echo off
echo Building pgSchema...
"C:\Program Files\Microsoft Visual Studio\VB98\VB6.EXE" /make ..\pgschema\pgSchema.vbp /outdir ..\binaries

echo Building Abstract Exporter...
"C:\Program Files\Microsoft Visual Studio\VB98\VB6.EXE" /make Exporters\Abstract\pgAbsExp.vbp /outdir ..\binaries
echo Building ASCII Exporter...
"C:\Program Files\Microsoft Visual Studio\VB98\VB6.EXE" /make Exporters\ASCII\pgASCII.vbp /outdir ..\binaries
echo Building Excel Exporter...
"C:\Program Files\Microsoft Visual Studio\VB98\VB6.EXE" /make Exporters\Excel\pgExcel.vbp /outdir ..\binaries
echo Building HTML Exporter...
"C:\Program Files\Microsoft Visual Studio\VB98\VB6.EXE" /make Exporters\HTML\pgHTML.vbp /outdir ..\binaries

echo Building Abstract Plugin...
"C:\Program Files\Microsoft Visual Studio\VB98\VB6.EXE" /make Plugins\Abstract\pgAbsPlg.vbp /outdir ..\binaries
echo Building MSysConf Plugin...
"C:\Program Files\Microsoft Visual Studio\VB98\VB6.EXE" /make Plugins\MSysConf\pgMSysConf.vbp /outdir ..\binaries
echo Building Security Plugin...
"C:\Program Files\Microsoft Visual Studio\VB98\VB6.EXE" /make Plugins\Security\pgSecurity.vbp /outdir ..\binaries
echo Building Import Plugin...
"C:\Program Files\Microsoft Visual Studio\VB98\VB6.EXE" /make Plugins\Security\pgImport.vbp /outdir ..\binaries

echo Building Highlightbox...
"C:\Program Files\Microsoft Visual Studio\VB98\VB6.EXE" /make ..\ActiveX\Highlightbox\Highlightbox.vbp /outdir ..\binaries

echo Building pgAdmin...
"C:\Program Files\Microsoft Visual Studio\VB98\VB6.EXE" /make pgAdmin2.vbp /outdir ..\binaries

echo
echo Done!
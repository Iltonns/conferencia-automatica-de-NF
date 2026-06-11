@echo off

REM Define o diretório do projeto para garantir paths relativos consistentes.
cd /d "C:\Users\Dell\OneDrive\Documentos\GitHub\conferencia-automatica-de-NF"

REM Executa a conferência automática e registra stdout/stderr no log raiz.
"C:\Users\Dell\AppData\Local\Python\bin\python.exe" "C:\Users\Dell\OneDrive\Documentos\GitHub\conferencia-automatica-de-NF\script.py" >> "C:\Users\Dell\OneDrive\Documentos\GitHub\conferencia-automatica-de-NF\conferencia.log" 2>&1

exit

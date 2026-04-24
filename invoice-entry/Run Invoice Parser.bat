@echo off
title Invoice Parser
echo ================================
echo        CBS Invoice Parser
echo ================================
echo.
python "%~dp0invoice_parser.py" --outlook
echo.
if %ERRORLEVEL% NEQ 0 (
    echo Something went wrong. See error above.
) else (
    echo Done! Check your Documents folder for the Excel file.
)
echo.
pause

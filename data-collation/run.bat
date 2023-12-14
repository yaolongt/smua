@echo off
echo Running Python code...

python ./merge_files.py

if %errorlevel% neq 0 (
    echo Python script execution failed.
    pause
) else (
    echo Execution complete...
    exit /b 0
)
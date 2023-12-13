@echo off
cd src
echo Installing Python Dependencies...
pip install -r ./requirements.txt

if %errorlevel% neq 0 (
    echo Failed to install dependencies. Exiting...
    pause
    exit /b %errorlevel%
)

echo Dependencies installed successfully...
echo Running Python code...

python ./manage.py runserver

if %errorlevel% neq 0 (
    echo Python script execution failed.
    pause
) else (
    echo Execution complete...
    exit /b 0
)
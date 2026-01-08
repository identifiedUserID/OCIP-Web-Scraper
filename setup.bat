@echo off
echo Setting up Python environment...

py -3 -m venv venv

call venv\Scripts\activate

python -m pip install --upgrade pip
pip install -r requirements.txt

echo.
echo Setup complete.
pause

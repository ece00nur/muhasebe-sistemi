@echo off
title Treyler Cari ve Fatura Takip Sistemi
cd /d "%~dp0"

echo ======================================================================
echo    TREYLER VE DORSE SANAYI - CARI VE FATURA TAKIP SISTEMI
echo ======================================================================
echo.
echo  Sistem baslatiliyor...
echo  Tarayiciniz aciliyor: http://127.0.0.1:8000
echo.
echo  [BILGI] Bu siyah pencere acik kaldigi surece program calisir.
echo         Pencereyi simge durumuna kucultebilirsiniz (asagi alabilirsiniz).
echo ======================================================================
echo.

:: Tarayiciyi ac
start http://127.0.0.1:8000

:: Sunucuyu baslat
python app.py
pause

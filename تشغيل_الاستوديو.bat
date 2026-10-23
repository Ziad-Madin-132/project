@echo off
chcp 65001 >nul
title استوديو أثر للإنتاج الفني - Athar Studio
cd /d "%~dp0"
cls
echo ================================================================
echo          منظومة استوديو أَثَـــر للإنتاج الفني والتسجيل الصوتي
echo ================================================================
echo.

where node >nul 2>&1
if %errorlevel% neq 0 (
    echo  [تنبيه]: لم يتم العثور على Node.js على هذا الجهاز.
    echo  سيتم فتح المنظومة الآن مباشرة عبر المتصفح.
    echo.
    echo  • لمزامنة ملف الإكسل تلقائياً يمكنك تثبيت Node.js من:
    echo    https://nodejs.org
    echo ================================================================
    echo.
    start "" "%~dp0index.html"
    echo  تم تشغيل المنظومة في المتصفح بنجاح!
    echo.
    pause
    exit /b
)

echo  • جاري تشغيل الخادم المحلي...
echo  • فتح المنظومة في المتصفح: http://localhost:3000
echo  • ملف الإكسل المرتبط: حجوزات_استوديو_أثر.xlsx
echo.
echo  بيانات الدخول الافتراضية للوحة التحكم:
echo  - اسم المستخدم: admin
echo  - كلمة المرور : athar2026
echo ================================================================
echo.

start /b "" cmd /c "timeout /t 2 /nobreak >nul && start http://localhost:3000"
node server.js
pause

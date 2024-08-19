:: 目标目录
set "target_dir=C:\Program Files\combinexls"

:: 检查目标目录是否存在
if exist "!target_dir!" (
    echo Removing existing directory !target_dir!
    rmdir /S /Q "!target_dir!"
)

:: 复制 combinexls 文件夹到 C:\Program Files
echo Copying !source_dir! to !target_dir!...
xcopy /E /I "!source_dir!" "!target_dir!" 
echo complate copy task


:: 定义要添加的路径
set "newPath=C:\Program Files\combinexls"

:: 获取当前的 Path 环境变量
for /f "tokens=2*" %%A in ('reg query "HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Control\Session Manager\Environment" /v Path 2^>nul') do (
    set "currentPath=%%B"
)

:: 检查 Path 中是否包含 newPath
echo !currentPath! | findstr /i /c:"combinexls"
if %errorlevel% equ 0 (
    echo The environment variable value for combinexls already exists. ignore write...
) else (
    reg add "HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Control\Session Manager\Environment" /v Path /t REG_EXPAND_SZ /d "!Path!;!newPath!\\" /f
)

REG ADD HKLM\SOFTWARE\Classes\Directory\background\shell\combinexlsx /T reg_sz /d "RUN COMBINEXLSX" /f
REG ADD HKLM\SOFTWARE\Classes\Directory\background\shell\combinexlsx /V icon /T REG_SZ /d "C:\Program Files\combinexls\combinexls.exe" /f
REG ADD HKLM\SOFTWARE\Classes\Directory\background\shell\combinexlsx /V Position /T REG_SZ /d "Top" /f
REG ADD HKLM\SOFTWARE\Classes\Directory\background\shell\combinexlsx /V SeparatorAfter /T REG_SZ /f
REG ADD HKLM\SOFTWARE\Classes\Directory\background\shell\combinexlsx\Command /T REG_SZ /d "C:\Program Files\combinexls\combinexls.exe \"%%V\\" /f

echo All tasks completed.
pause
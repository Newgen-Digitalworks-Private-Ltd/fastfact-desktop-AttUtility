@echo OFF 
call "C:\Program Files\Microsoft Visual Studio\2022\Community\VC\Auxiliary\Build\vcvars64.bat" x86
echo "Starting Build for all Projects with proposed changes"
echo .  
devenv "D:\Jenkins\.jenkins\workspace\AttendanceUtility\AttendanceUtility.sln" /Rebuild Debug
echo . 
echo "All builds completed." 
pause
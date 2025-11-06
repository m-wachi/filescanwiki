cd C:\usr\src\python\filescanwiki
C:\usr\src\python\filescanwiki\py3env\Scripts\Activate.ps1



del filescanwiki05_all.log
del filescanwiki05_err_all.log

python filescanwiki05.py (\\some_computer\sample_folder1)
Get-Content filescanwiki05.log >> filescanwiki05_all.log
Get-Content filescanwiki05_err.log >> filescanwiki05_err_all.log

python filescanwiki05.py (\\some_computer\sample_folder2)
Get-Content filescanwiki05.log >> filescanwiki05_all.log
Get-Content filescanwiki05_err.log >> filescanwiki05_err_all.log

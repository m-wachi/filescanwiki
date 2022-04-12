set PYTHON_PATH=C:\Python27
set PYTHON_SCRIPTS=C:\Python27\Scripts
set PATH=%PYTHON_PATH%;%PYTHON_SCRIPTS%;%PATH%

cd C:\usr\src\python\filescanwiki
del filescanwiki04_all.log
del filescanwiki04_err_all.log

python filescanwiki04.py (\\some_computer\sample_folder1)
type filescanwiki04.log > filescanwiki04_all.log
type filescanwiki04_err.log > filescanwiki04_err_all.log

python filescanwiki04.py (\\some_computer\sample_folder2)
type filescanwiki04.log >> filescanwiki04_all.log
type filescanwiki04_err.log >> filescanwiki04_err_all.log


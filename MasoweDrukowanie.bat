mode con cp select=1250
setlocal EnableDelayedExpansion

set /A dzielnik=400
set /A licznik=0
set /A rozszerzenieNazwyPliku=1
set    rozszerzenieSzukanychPlikow=docx
set    sciezkaWord=C:\Program Files\Microsoft Office 15\root\office15\WINWORD.EXE


echo mode con cp select=1250 >drukuj%rozszerzenieNazwyPliku%.bat

for %%a in (*.%rozszerzenieSzukanychPlikow%) do (
	set /A licznik+=1
	SET /A mod = !licznik! %% %dzielnik%
	echo !mod!
	if !mod! == 0 (
	SET /A  rozszerzenieNazwyPliku=!rozszerzenieNazwyPliku! + 1
	echo mode con cp select=1250 >drukuj!rozszerzenieNazwyPliku!.bat
	echo "%sciezkaWord% " "%%a" /wait /mFilePrintDefault /mFileCloseOrExit >>drukuj!rozszerzenieNazwyPliku!.bat
	) else (
	echo "%sciezkaWord% " "%%a" /wait /mFilePrintDefault /mFileCloseOrExit >>drukuj!rozszerzenieNazwyPliku!.bat
	)
)


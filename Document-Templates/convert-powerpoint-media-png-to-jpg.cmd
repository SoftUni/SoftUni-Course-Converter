:: This script is designed to reduce the file size of PPTX presentations by
:: compressing the images in the slides. Typically PNG files 2-3 MB are reduced
:: to 150-200 KB. Only large images are compressed, the others stay untouched.
::
:: Unzip the PPTX archive, copy the script to `ppt/media`, run the script
:: Requires Image Magick to be installed (`magick.exe` on the system PATH)
::

mkdir newimg

del /s /q newimg\*.*

for %%i in (*.png) do (
	:: Process only files media bigger than 250 KB (256000 bytes)
	if %%~zi gtr 256000 (
		:: Convert the PNG image file to JPG using ImageMagick (85% quality)
		magick -quality 85 "%%i" "newimg\%%~ni.jpg"
	)
)

copy /y newimg\*.jpg *.png

del /s /q newimg\*.*

rmdir newimg

pause

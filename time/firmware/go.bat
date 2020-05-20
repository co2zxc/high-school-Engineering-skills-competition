@echo Copy go.bat to firmware directory
@echo off
del .\usbdrv\*.o
del main.hex
@echo on
make hex
pause
@echo off
del main.elf
del main.o
@echo on
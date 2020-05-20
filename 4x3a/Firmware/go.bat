@echo Copy go.bat to firmware directory
@echo off
make hex
make clean
make hex
del .\usbdrv\*.o
del main.hex
@echo on
make hex
pause
@echo off
del main.elf
del main.o
@echo on
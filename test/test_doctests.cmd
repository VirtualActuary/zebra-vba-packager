@echo off
pushd "%~dp0"
    python.exe test_doctests.py
popd

if /i "%comspec% /c ``%~0` `" equ "%cmdcmdline:"=`%" pause
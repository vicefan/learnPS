@echo off

if not (Test-Path "C:\tmp") then (
    mkdir "C:\tmp"
)

cd C:\tmp

if not exist "C:\tmp\MergePDF" (
    git clone https://github.com/yourusername/MergePDF.git
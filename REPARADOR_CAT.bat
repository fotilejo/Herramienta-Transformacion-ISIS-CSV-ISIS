@echo off
title Reparador CAT
chcp 65001 > nul
cd /d "%~dp0"
python reparador.py

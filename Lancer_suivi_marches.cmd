@echo off
setlocal EnableExtensions
chcp 65001 >NUL 2>&1
title Suivi des marches

REM ---------------------------------------------------------------------
REM Lanceur de l'application de suivi des marches.
REM Double-cliquer ce fichier : il verifie les dependances, les installe si
REM besoin, puis demarre l'application en journalisant dans run_logs\.
REM
REM Le travail est fait par lanceur.py. cmd.exe gere mal l'accent du nom du
REM script principal selon la page de code active ; Python, lui, ne s'en
REM soucie pas.
REM
REM   Lancer_suivi_marches.cmd            lance l'application
REM   Lancer_suivi_marches.cmd --maj      reinstalle les dependances puis lance
REM   Lancer_suivi_marches.cmd --verif    verifie seulement
REM ---------------------------------------------------------------------

cd /d "%~dp0"

REM Chercher un Python utilisable, par ordre de preference.
set "PYCMD="
for %%P in (py.exe py python.exe python python3.exe python3) do (
  if not defined PYCMD (
    where %%P >NUL 2>&1 && set "PYCMD=%%P"
  )
)

REM Dernier recours : emplacements d'installation par defaut, du plus recent.
if not defined PYCMD (
  for %%V in (313 312 311 310) do (
    if not defined PYCMD (
      if exist "%LocalAppData%\Programs\Python\Python%%V\python.exe" (
        set "PYCMD=%LocalAppData%\Programs\Python\Python%%V\python.exe"
      )
    )
  )
)

if not defined PYCMD (
  echo.
  echo [ERREUR] Python est introuvable sur ce poste.
  echo          Installer Python depuis https://www.python.org/downloads/
  echo          en cochant "Add python.exe to PATH" pendant l'installation.
  echo.
  pause
  exit /b 1
)

if not exist "lanceur.py" (
  echo.
  echo [ERREUR] lanceur.py est absent de ce dossier.
  echo          Placer ce fichier .cmd dans le dossier de l'application,
  echo          a cote de lanceur.py et du script principal.
  echo.
  pause
  exit /b 1
)

"%PYCMD%" -X utf8 lanceur.py %*
set "RC=%ERRORLEVEL%"

REM Garder la fenetre ouverte en cas d'echec, pour que le message reste lisible.
if not "%RC%"=="0" (
  echo.
  pause
)
exit /b %RC%

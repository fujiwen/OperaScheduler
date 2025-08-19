@echo off
chcp 65001
set NLS_LANG=AMERICAN_AMERICA.AL32UTF8



SQLPLUS "SYS/opera10g AS SYSDBA" @check_standby.sql
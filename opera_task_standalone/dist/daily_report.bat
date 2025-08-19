@echo off
chcp 65001
set NLS_LANG=AMERICAN_AMERICA.AL32UTF8

sqlplus sys/opera10g as sysdba @daily_report_dg.sql

sqlplus sys/opera10g@production as sysdba @daily_report_prod.sql

exit
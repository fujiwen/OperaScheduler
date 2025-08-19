select process, status, sequence# from v$managed_standby;
select database_role,controlfile_type,open_mode,protection_mode from v$database;
Select sequence#,applied from v$archived_log;
exit;


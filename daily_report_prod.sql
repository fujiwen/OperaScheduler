set pagesize 1000
set heading off
set serveroutput on
set trimspool on
set feedback off
set echo off
set time on
set termout on
set serverout on
set linesize 132
set pages 999
clear columns
break on report
column "INSTANCE NAME"		format	a8
column "DATABASE NAME"	 	format	a13
column "USER SESSIONS" 	 	format	999
column "DATABASE CONNECTIONS" 	format	999
column "EXT.MGMT"		format	a8
column "MAX SIZE (M)"		format	9999,990.000
column "FREE (M)"		format  999,990.000
column "FREE %"			format	990.00
column "SIZE (M)"		format	9999,990.000
column "TYPE"			format	a9
column "TABLESPACE"		format	a10
column "USED (M)"		format	9999,990.000
column "USED %"			format	990.00
column "USED TOTAL %"		format	990.00
column 	INPUT_TYPE		format	a10
column	OUTPUT_MBYTES		format	9,999,999		heading "OUTPUT (M)"
column	SESSION_RECID		format	999999			heading "SESSION RECID"
column	TIME_TAKEN_DISPLAY	format	a10			heading	"TIME TAKEN"
column	OUTPUT_INSTANCE 	format 	9999			heading "OUT INST"
COMPUTE SUM LABEL 'Total:' OF "SIZE (M)" "MAX SIZE (M)" on report

SET MARKUP HTML ON ENTMAP ON SPOOL ON PREFORMAT OFF -
HEAD "<style type='text/css'> -
body { -
font:10pt Arial,Helvetica,sans-serif; -
color:black; background:White;} -
p { -
font:10pt Arial,Helvetica,sans-serif; -
color:black; background:White;} -
table,tr,td { -
font:10pt Arial,Helvetica,sans-serif; -
color:Black; background:#f7f7e7; -
padding:0px 0px 0px 0px; margin:0px 0px 0px 0px;} -
th { -
font:bold 10pt Arial,Helvetica,sans-serif; -
color:#336699; background:#cccc99; -
padding:0px 0px 0px 0px;} -
h1 { -
font:bold 18pt Arial,Helvetica,Geneva,sans-serif; text-align:center; -
color:#336699; background-color:White; text-decoration: underline; -
border-bottom:0px solid #cccc99; -
margin-top:10pt; margin-bottom:5pt; padding:0px 0px 0px 0px;} -
h2 { -
font:bold 15pt Arial,Helvetica,Geneva,sans-serif; text-align:center; -
color:#336699; background-color:White; text-decoration: underline; -
border-bottom:0px solid #cccc99; -
margin-top:5pt; margin-bottom:5pt; padding:0px 0px 0px 0px;} -
h3 { -
font:bold 13pt Arial,Helvetica,Geneva,sans-serif; -
color:#336699; background-color:White; -
margin-top:4pt; margin-bottom:2pt;} -
a { -
font:9pt Arial,Helvetica,sans-serif; -
color:#663300; background:#ffffff; -
margin-top:0pt; margin-bottom:0pt; vertical-align:top;} -
</style> -
<title>Opera Data Guard Daily Report (Version 1.2)</title>"

spool d:\scripts\logs\daily_report.html append

SET MARKUP HTML ON ENTMAP OFF

prompt <br>
prompt <h2>Production Database</h2>

prompt <h3>General Database Information:</h3>
set heading on
select a.inst_id,a.name "DATABASE NAME", upper(c.instance_name) "INSTANCE NAME", '<b><font color="#8A0829">'||c.status||'</font></b>' "STATUS",
'<b><font color="#8A0829">'||c.host_name||'</font></b>' "HOST NAME", '<b><font color="#8A0829">'||a.database_role||'</font></b>' "DATABASE ROLE",
to_char(c.startup_time,'DD-MON-YYYY HH24:MI') "START TIME",to_char(sysdate,'DD-MON-YYYY HH24:MI') "SYSTEM DATE"
from gv$database a, gv$instance c
where a.inst_id = c.inst_id
order by a.inst_id
/
set heading off

prompt <h3>Opera Version Information:</h3>
set heading on
SELECT distinct version || ' ' || epatchlevel "OPERA VERSION"
FROM opera.installed_app
/
set heading off

prompt <h3>Oracle version Information:</h3>
set heading on
SELECT version || ' ' || comments version
FROM
(SELECT version, comments
FROM registry$history
where lower(comments) LIKE '%patch%'
ORDER BY action_time DESC)
WHERE ROWNUM=1
/
set heading off

set heading on
select platform_name from v$database;
set heading off

prompt <h3>Tablespace usage:</h3>
set heading on
set lines 200 pages 100
select
   c.tablespace_name                         "TABLESPACE",
   round(a.bytes/1048576,4)                  "SIZE (M)",
   round(maxbytes/1048576,4)                 "MAX SIZE (M)",
   round((a.bytes-b.bytes)/maxbytes,4) * 100 "USED %",
   c.contents                                "TYPE",
case when round((a.bytes-b.bytes)/maxbytes,4) * 100 > 90 then '<font size="+0" face="Arial,Helvetica,Geneva,sans-serif" color="#FF0000"><b>DANGER</b></font>'
     when round((a.bytes-b.bytes)/maxbytes,4) * 100 > 80 and round((a.bytes-b.bytes)/maxbytes,4) * 100 < 90 then '<font size="+0" face="Arial,Helvetica,Geneva,sans-serif" color="#FFBF00"><b>WARNING</b></font>'
     ELSE '<font size="+0" face="Arial,Helvetica,Geneva,sans-serif" color="#298A08"><b>OK</b></font>' END STATUS
   from
  ( select tablespace_name,
           sum(a.bytes) bytes,
           min(a.bytes) minbytes,
           sum(decode(a.autoextensible, 'YES', a.maxbytes,'NO', a.bytes)) maxbytes
      from DBA_DATA_FILES a
     group by tablespace_name
  union all
    select tablespace_name,
           sum(a.bytes) bytes,
           min(a.bytes) minbytes,
           sum(decode(a.autoextensible, 'YES', a.maxbytes,'NO', a.bytes)) maxbytes
      from DBA_TEMP_FILES a
     group by tablespace_name )                         a,
  ( select a.tablespace_name,
           nvl(sum(b.bytes),0) bytes
      from DBA_DATA_FILES          a
left outer join DBA_FREE_SPACE b
        on ( a.tablespace_name = b.tablespace_name
       and a.file_id           = b.file_id )
  group by a.tablespace_name )                          b,
       dba_tablespaces                                  c
 where a.tablespace_name = b.tablespace_name(+)
   and a.tablespace_name = c.tablespace_name
 order by 1,3
/
set heading off

prompt <h3>Number of database connections:</h3>
set heading on
select inst_id instance,(count(*)-1) "DATABASE CONNECTIONS"
from gv$session
where username is not null
and program not like '%ORACLE.EXE%'
group by inst_id
order by inst_id
/
set heading off

DECLARE
  v_owner        dba_objects.owner%TYPE;
  v_cnt         number;
 CURSOR Cursor_Invalids IS
                 SELECT distinct owner FROM dba_objects where status != 'VALID' and object_type != 'NEXT OBJECT' order by owner;
BEGIN
	select count(*) into v_cnt FROM dba_objects where status != 'VALID' and object_type != 'NEXT OBJECT';
	IF v_cnt > 0 THEN
	dbms_output.put_line('<h3>Invalid database objects:</h3>');
	dbms_output.put_line('<left><font size="+0" face="Arial,Helvetica,Geneva,sans-serif" color="#FF0000"><b>Please contact SHIJI support-4000211988.</b></font></left>');
        dbms_output.put_line('<br>');
	END IF;
	OPEN Cursor_Invalids;
  LOOP
    FETCH Cursor_Invalids INTO v_owner;
    IF Cursor_Invalids%FOUND THEN
		dbms_output.put_line('<left><font size="+0" face="Arial,Helvetica,Geneva,sans-serif" color="#FF0000"><b>'||TO_CHAR(Cursor_Invalids%ROWCOUNT)||' invalid object in '||v_owner||' schema.</b></font></left>');
    ELSE EXIT;
    END IF;
  END LOOP;
  dbms_output.put_line('<br>');
  CLOSE Cursor_Invalids;
END;
/

prompt <h3>List of last 3 days backups:</h3>
declare
v_cnt number;
BEGIN
select count(*) into v_cnt from V$RMAN_BACKUP_JOB_DETAILS where start_time > trunc(sysdate-3);
if v_cnt=0  then
dbms_output.put_line('<left><font size="+0" face="Arial,Helvetica,Geneva,sans-serif" color="#FF0000"><b>No backup information found. Please check backup logs.</b></font></left>');
dbms_output.put_line('<br>');
end if;
end;
/
set heading on
set lines 220 pages 1000
select
  j.session_recid,
  to_char(j.start_time, 'DD-MON-YYYY HH24:MI') start_time,
  to_char(j.end_time, 'DD-MON-YYYY HH24:MI') end_time,
  (j.output_bytes/1024/1024) output_mbytes, j.status, j.input_type,
  decode(to_char(j.start_time, 'd'), 1, 'Sunday', 2, 'Monday',
                                     3, 'Tuesday', 4, 'Wednesday',
                                     5, 'Thursday', 6, 'Friday',
                                     7, 'Saturday') DAY,
  j.time_taken_display, ro.inst_id output_instance
from V$RMAN_BACKUP_JOB_DETAILS j
  left outer join (select
                     d.session_recid, d.session_stamp
                   from
                     V$BACKUP_SET_DETAILS d
                     join V$BACKUP_SET s on s.set_stamp = d.set_stamp and s.set_count = d.set_count
                   where s.input_file_scan_only = 'NO'
                   group by d.session_recid, d.session_stamp) x
    on x.session_recid = j.session_recid and x.session_stamp = j.session_stamp
  left outer join (select o.session_recid, o.session_stamp, min(inst_id) inst_id
                   from GV$RMAN_OUTPUT o
                   group by o.session_recid, o.session_stamp)
    ro on ro.session_recid = j.session_recid and ro.session_stamp = j.session_stamp
where j.start_time > trunc(sysdate)- 3
order by j.start_time
/
set heading off

spool off
SET MARKUP HTML OFF ENTMAP OFF SPOOL OFF PREFORMAT ON
clear column

--host d:\scripts\send_report.bat
exit

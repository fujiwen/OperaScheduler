set echo off
set time on
set pagesize 9999
set heading off
set termout on
set serverout on
set linesize 132
set trimspool on
set pages 999
set feedback off
clear  column
column "INSTANCE NAME"	format	a13
column "DATABASE NAME"	format	a13

host echo RMAN target sys/opera10g cmdfile logs\del_app_arch.rcv log logs\del_app_arch.log > logs\del_arch_rman.bat

spool logs\del_app_arch.rcv

declare
v_thread number;
v_min number:=0;
v_max number:=0;

BEGIN
SELECT max(thread#) into v_thread from v$archived_log;
dbms_output.put_line('run {');
dbms_output.put_line('allocate channel dev0 type disk;');
dbms_output.put_line('crosscheck archivelog all;');
for d_seq in 1..v_thread
loop
select min(sequence#) into v_min from v$archived_log where thread#=d_seq and applied='YES' and deleted='NO' and trunc(completion_time) <= (trunc(sysdate)-2);
select max(sequence#) into v_max from v$archived_log where thread#=d_seq and applied='YES' and deleted='NO' and trunc(completion_time) <= (trunc(sysdate)-2);
if (v_min<>0 and v_max<>0) then
dbms_output.put_line('delete archivelog from sequence '||v_min||' until sequence '||v_max||' thread '||d_seq||';');
v_min:=0;
v_max:=0;
end if;
end loop;
dbms_output.put_line('}');
end;
/

spool off

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

spool logs\daily_report.html

SET MARKUP HTML ON ENTMAP OFF

select '<h1>Opera Data Guard Daily Report (Version 1.2)</h1>' || chr(10) || '<b><center><font size="+0" face="Arial,Helvetica,Geneva,sans-serif" color="#336699">'||to_char(sysdate,'DD-MON-YY')||'</font></b>' from dual;

prompt <br>

prompt <h2>Standby Database</h2>

prompt <h3>General Database Information:</h3>
set heading on
select a.inst_id,a.name "DATABASE NAME", upper(c.instance_name) "INSTANCE NAME", '<b><font color="#8A0829">'||c.status||'</font></b>' "STATUS",
'<b><font color="#8A0829">'||c.host_name||'</font></b>' "HOST NAME", '<b><font color="#8A0829">'||a.database_role||'</font></b>' "DATABASE ROLE",
(select protection_mode from v$database) "PROTECTION MODE", to_char(c.startup_time,'DD-MON-YYYY HH24:MI') "START TIME",to_char(sysdate,'DD-MON-YYYY HH24:MI') "SYSTEM DATE"
from gv$database a, gv$instance c
where a.inst_id = c.inst_id
order by a.inst_id
/
set heading off
declare
v_status v$instance.status%TYPE;
begin
	select status INTO v_status from v$instance;
	if v_status <> 'MOUNTED' then
		DBMS_OUTPUT.put_line('<center><font size="+0" face="Arial,Helvetica,Geneva,sans-serif" color="#FF0000"><b>(STANDBY database must be in MOUNTED state, please run start_standby.bat)</b></font></center>');
	else null;
	end if;
end;
/

prompt <h3>Applied logs:</h3>
set heading on
select 'Last Applied  : ' Logs, to_char(next_time,'DD-MON-YYYY HH24:MI') Time
from v$archived_log
where sequence# = (select max(sequence#) from v$archived_log where applied='YES')
union
select 'Last Received : ' Logs, to_char(next_time,'DD-MON-YYYY HH24:MI') Time
from v$archived_log
where sequence# = (select max(sequence#) from v$archived_log)
/
set heading off

prompt <h3>Archive gaps:</h3>
set heading on
select ''|| 
((select nvl(max(high_sequence#),0) - nvl(max(low_sequence#), 0) from v$archive_gap where thread#=1)
--(select  nvl(max(high_sequence#),0) - nvl(max(low_sequence#), 0) from v$archive_gap where thread#=2) +
--(select  nvl(max(high_sequence#),0) - nvl(max(low_sequence#), 0) from v$archive_gap where thread#=3) +
--(select  nvl(max(high_sequence#),0) - nvl(max(low_sequence#), 0) from v$archive_gap where thread#=4) +
--(select  nvl(max(high_sequence#),0) - nvl(max(low_sequence#), 0) from v$archive_gap where thread#=5) +
--(select  nvl(max(high_sequence#),0) - nvl(max(low_sequence#), 0) from v$archive_gap where thread#=6) 
-- please copy and insert the above line here to match with the number of threads you have in the cluster (eg: given below)
-- (select nvl(max(high_sequence#),0) - nvl(max(low_sequence#), 0) from v$archive_gap where thread#=3)
) "GAPS"
from dual
/
set heading off
declare
v_gap	number;
v_low_seq	number;
v_high_seq	number;
begin
	select 
	((select nvl(max(high_sequence#),0) - nvl(max(low_sequence#), 0) from v$archive_gap where thread#=1)
	-- (select  nvl(max(high_sequence#),0) - nvl(max(low_sequence#), 0) from v$archive_gap where thread#=2) +
	-- (select  nvl(max(high_sequence#),0) - nvl(max(low_sequence#), 0) from v$archive_gap where thread#=3) +
	-- (select  nvl(max(high_sequence#),0) - nvl(max(low_sequence#), 0) from v$archive_gap where thread#=4) +
	-- (select  nvl(max(high_sequence#),0) - nvl(max(low_sequence#), 0) from v$archive_gap where thread#=5) +
	-- (select  nvl(max(high_sequence#),0) - nvl(max(low_sequence#), 0) from v$archive_gap where thread#=6) 
	-- please copy and insert the above line here to match with the number of threads you have in the cluster (eg: given below)
	-- (select nvl(max(high_sequence#),0) - nvl(max(low_sequence#), 0) from v$archive_gap where thread#=3)
	) into v_gap
	from dual;
	
	-- 获取间隙的详细信息
	select nvl(max(low_sequence#), 0), nvl(max(high_sequence#), 0) 
	into v_low_seq, v_high_seq 
	from v$archive_gap where thread#=1;
	
	if v_gap > 0 then
		dbms_output.put_line('<center><font size="+0" face="Arial,Helvetica,Geneva,sans-serif" color="#FF0000"><b>归档日志间隙检查: 异常 (存在间隙)</b></font></center>');
		dbms_output.put_line('<center><font size="+0" face="Arial,Helvetica,Geneva,sans-serif" color="#FF0000"><b>间隙数量: ' || v_gap || ' 个日志文件</b></font></center>');
		dbms_output.put_line('<center><font size="+0" face="Arial,Helvetica,Geneva,sans-serif" color="#FF0000"><b>间隙范围: 序列号 ' || v_low_seq || ' 到 ' || v_high_seq || '</b></font></center>');
		dbms_output.put_line('<center><font size="+0" face="Arial,Helvetica,Geneva,sans-serif" color="#FF0000"><b>(Databases are not in sync, Please contact SHIJI support-4000211988.)</b></font></center>');
	else 
		dbms_output.put_line('<center><font size="+0" face="Arial,Helvetica,Geneva,sans-serif" color="#008000"><b>归档日志间隙检查: 正常 (无间隙)</b></font></center>');
	end if;
end;
/

prompt <h3>Logs not applied:</h3>
set heading on
select count(*) "NOT APPLIED"
from v$archived_log
where applied='NO' and creator='ARCH'
/
set heading off
declare
v_log	number;
begin
	select count(*) into v_log from v$archived_log where applied='NO' and creator='ARCH';
	if v_log > 0 then
		dbms_output.put_line('<center><font size="+0" face="Arial,Helvetica,Geneva,sans-serif" color="#FF0000"><b>未应用日志检查: 异常 (存在未应用日志)</b></font></center>');
		dbms_output.put_line('<center><font size="+0" face="Arial,Helvetica,Geneva,sans-serif" color="#FF0000"><b>未应用日志数量: ' || v_log || ' 个日志文件</b></font></center>');
		dbms_output.put_line('<center><font size="+0" face="Arial,Helvetica,Geneva,sans-serif" color="#FF0000"><b>(Some archive logs are not applied, Please contact SHIJI support-4000211988.)</b></font></center>');
	else 
		dbms_output.put_line('<center><font size="+0" face="Arial,Helvetica,Geneva,sans-serif" color="#008000"><b>未应用日志检查: 正常 (所有日志已应用)</b></font></center>');
	end if;
end;
/

prompt <h3>Deleted archive logs:</h3>
set heading on
select count(*) "DELETED ARCHIVE LOGS"
from v$archived_log where applied='YES' and deleted='NO'
and trunc(completion_time) <= (trunc(sysdate)-2)
/
set heading off

prompt <h3>Process on standby server:</h3>
set heading on
select process, status from v$managed_standby;
set heading off

spool off
SET MARKUP HTML OFF ENTMAP OFF SPOOL OFF PREFORMAT ON
clear column

host logs\del_arch_rman.bat

host del logs\del_arch_rman.bat

host del logs\del_app_arch.rcv

exit
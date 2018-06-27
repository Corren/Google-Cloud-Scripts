------------------snapshots-----------------
---creates snapshot---
CREATE DATABASE ADEPTO_SS
ON
(NAME=Empty,FILENAME='Z:\ss\ADEPTO_SS.ss')
AS SNAPSHOT OF ADEPTO

------------ check size of snapshot---------
USE ADEPTO_SS
EXEC sp_spaceused

--------------restores snapshot-------------
USE master
RESTORE DATABASE ADEPTO
FROM DATABASE_SNAPSHOT = 'ADEPTO_SS'

------drops snapshot----
------must do when done as backblaze will not back it up --------
drop DATABASE ADEPTO_SS
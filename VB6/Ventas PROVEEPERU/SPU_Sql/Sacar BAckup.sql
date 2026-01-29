BACKUP DATABASE [DbSistema_18072024]
TO DISK = 'D:\Sistema Vb6\Backup BD SQL\DbSistema_13082024.bak'
WITH FORMAT, 
MEDIANAME = 'SQLServerBackups', 
NAME = 'Full Backup of YourDatabaseName'


-- restaurar 
RESTORE DATABASE [DbSistema_18072024] -- nuevo nombre de la base de datos
FROM DISK = 'D:\Sistema Vb6\Backup BD SQL\DbSistema_18072024.bak'
WITH REPLACE,
MOVE 'DbSistema' TO 'C:\Program Files\Microsoft SQL Server\MSSQL16.SQLEXPRESS\MSSQL\DATA\DbSistema_18072024.mdf',
MOVE 'DbSistema_log' TO 'C:\Program Files\Microsoft SQL Server\MSSQL16.SQLEXPRESS\MSSQL\DATA\DbSistema_18072024.ldf'
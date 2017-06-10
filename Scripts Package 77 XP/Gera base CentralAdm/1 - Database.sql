-- =============================================
-- Basic Create Database Template
-- =============================================
IF EXISTS (SELECT * 
	   FROM   master..sysdatabases 
	   WHERE  name = 'CentralAdm')
	DROP DATABASE CentralAdm
GO

CREATE DATABASE CentralAdm
GO


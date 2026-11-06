from pathlib import Path
import subprocess,json
r=Path(__file__).resolve().parents[1];base=r/'../../Basededatos/Dbsishyt'
def save(p,t):p.write_bytes(t.replace('\r\n','\n').replace('\n','\r\n').encode('utf-8'))
sig='-- Firma: FRANCO LARA | Create date: 05/11/2026 | Ubigeo de FralseCont y direcciones de clientes/tiendas.\n'
specs=[('UbigeoDepartamentos','CodigoDepartamento','CHAR(2)','CodigoDepartamento,Nombre,Activo',[]),('UbigeoProvincias','CodigoProvincia','CHAR(4)','CodigoProvincia,CodigoDepartamento,Nombre,Activo',[('CodigoDepartamento','CHAR(2)','UbigeoDepartamentos','CodigoDepartamento')]),('UbigeoDistritos','CodigoUbigeo','CHAR(6)','CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo',[('CodigoDepartamento','CHAR(2)','UbigeoDepartamentos','CodigoDepartamento'),('CodigoProvincia','CHAR(4)','UbigeoProvincias','CodigoProvincia')])]
seeds=sig+'USE [SistemaWebHYT];\nGO\nSET NOCOUNT ON;\n'
for table,key,typ,cols,refs in specs:
 result=subprocess.run(['sqlcmd','-S',r'localhost\SQLEXPRESS','-E','-C','-I','-b','-f','65001','-d','FralseCont','-y','0','-Q',f'SET NOCOUNT ON; SELECT {cols} FROM dbo.{table} ORDER BY {key} FOR JSON PATH;'],capture_output=True,check=True)
 rows=json.loads(''.join(result.stdout.decode('utf-8-sig').splitlines()).strip());print(table,len(rows))
 for row in rows:
  values=[('1' if v else '0') if isinstance(v,bool) else "N'"+str(v).replace("'","''")+"'" for v in row.values()]
  seeds+=f"IF NOT EXISTS(SELECT 1 FROM dbo.{table} AS destino WHERE destino.{key}=N'{row[key]}') INSERT INTO dbo.{table} ({','.join(row)}) VALUES ({','.join(values)});\n"
seeds+='GO\n';save(base/'Script/024_UbigeosCatalogo.sql',seeds)
save(base/'Script/025_DireccionesYConsideraciones.sql',sig+'''USE [SistemaWebHYT];
GO
SET NOCOUNT ON;
IF COL_LENGTH(N'dbo.Cliente',N'CodigoUbigeo') IS NULL ALTER TABLE dbo.Cliente ADD CodigoUbigeo CHAR(6) NULL CONSTRAINT FK_Cliente_Ubigeo REFERENCES dbo.UbigeoDistritos(CodigoUbigeo);
IF COL_LENGTH(N'dbo.Tienda',N'CodigoUbigeo') IS NULL ALTER TABLE dbo.Tienda ADD CodigoUbigeo CHAR(6) NULL CONSTRAINT FK_Tienda_Ubigeo REFERENCES dbo.UbigeoDistritos(CodigoUbigeo);
IF COL_LENGTH(N'dbo.ConfiguracionHYT',N'ConsideracionesPresupuesto') IS NULL ALTER TABLE dbo.ConfiguracionHYT ADD ConsideracionesPresupuesto NVARCHAR(MAX) NULL;
ALTER TABLE dbo.PresupuestoRevision ALTER COLUMN ClienteDireccion NVARCHAR(600) NULL;
ALTER TABLE dbo.PresupuestoRevision ALTER COLUMN TiendaDireccion NVARCHAR(600) NULL;
GO
''')


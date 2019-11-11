
![CLEVER DATA GIT REPO](https://github.com/congmingshuju/git-resources/blob/master/images/0-clever-data-github.png "李聪明 数据")


# 仅使用SQL自动导入Excel
### Automate Excel Import With SQL
**发布-日期:  2018年2月20日 (评论)**

![Automate Import With SQL](images/automate-excel-importing-with-sql.png?raw=true "Automate Import With SQL")


## Contents

- [中文](#中文)
- [English](#English)
- [SQL-Logic](#Logic)
- [Build Quality](#Build-Quality)
- [Author](#Author)
- [License](#License) 


## 中文
以下示例可以在不使用任何其他数据服务的情况下将Excel数据直接导入SQL Server。


## English
Here’s an example of something you can write to import Excel Data directly into SQL Server without the use of any other data services.

---
## Logic
```SQL
use [master];
set nocount on
 
-- 	allow queries to directly access excel files
-- 	允许直接访问excel文件
exec [master]..sp_configure 'show advanced options', 1;
reconfigure with override;
exec [master]..sp_configure 'ad hoc distributed queries', 1;
reconfigure with override;
exec [master].dbo.sp_msset_oledb_prop N'microsoft.ace.oledb.12.0', N'allowinprocess', 1;
exec [master].dbo.sp_msset_oledb_prop N'microsoft.ace.oledb.12.0', N'dynamicparameters', 1;
 
use [compliance];
set nocount on
 
-- 	get simple file list
-- 	获取简单文件列表
declare @import_path    varchar(255) = 'C:\SQLIMPORTS\'
declare @files      table ([subdirectory] varchar(255), [depth] int, [file] int)
insert into @files exec master..xp_dirtree @import_path, 1, 1; delete from @files where [file] = 0 
  
if object_id('tempdb..#import_folder') is not null
drop table      #import_folder
create table    #import_folder  ([file_info] nvarchar(255))
declare         @file_list  table ([file_name] nvarchar(255))
  
-- 	get file meta data
-- 	获取文件元数据
Insert into     #import_folder exec xp_cmdshell 'dir C:\SQLIMPORTS\*.*'
delete from     #import_folder where [file_info] is null 
delete from     #import_folder where [file_info] like '%dir%' 
delete from     #import_folder where [file_info] like '%volume%' 
delete from     #import_folder where [file_info] like '%bytes%'
  
-- 	combine meta data with file name
-- 	合并元数据与文件名
if object_id('tempdb..#files_found') is not null
drop table      #files_found
create table    #files_found ([create_date] datetime, [file_name] nvarchar(255))
 
-- 	compile list of all files found.
-- 	编译找到的所有文件列表
declare @compile_list       varchar(max)
set     @compile_list       = ''
select  @compile_list       = @compile_list + 
'insert into #files_found  ([create_date], [file_name]) select cast(left([file_info], 20) as datetime), ''' + [subdirectory] + ''' from #import_folder where [file_info] like ''%' + [subdirectory] + '%'';' + char(10)
from    @files
exec    (@compile_list)
  
-- 	create table name based on latest file added to directory
-- 	根据添加到的最新文件创建表名
declare @new_file   varchar(255) = (select top 1 [file_name] from #files_found where [file_name] like '%.xlsx' and left([file_name], 8) like 'MSSQL FY%' order by [create_date] desc)
declare @new_table  varchar(255) = (select 'IMPORT_' + replace(substring(@new_file, 7, charindex('Q', @new_file) + 3), ' ', '') + '_00')
 
-- 	create next import table based on excel columns found.
-- 	根据找到的excel列创建下一个导入表
declare @target varchar(255) = (select top 1 [table_name] from information_schema.tables where [table_name] like 'IMPORT_FY%' order by [table_name] desc)
declare @next_table varchar(255) = (select isnull(@target, @new_table))
select  @next_table = upper(left(@new_table, 14)) + format(cast(right(@next_table, 2) as int) + 1, '00')
 
declare @create_table   varchar(max) = 'create table [' + @next_table + ']' + char(10) + 
'
(
   [F1]  nvarchar(max),  [F2]  nvarchar(max),  [F3] nvarchar(max),   [F4]  nvarchar(max)
,  [F5]  nvarchar(max),  [F6]  nvarchar(max),  [F7] nvarchar(max),   [F8]  nvarchar(max)
,  [F9]  nvarchar(max),  [F10] nvarchar(max),  [F11] nvarchar(max),  [F12] nvarchar(max)
,  [F13] nvarchar(max),  [F14] nvarchar(max),  [F15] nvarchar(max),  [F16] nvarchar(max)
,  [F17] nvarchar(max),  [F18] nvarchar(max),  [F19] nvarchar(max),  [F20] nvarchar(max)
,  [F21] nvarchar(max),  [F22] nvarchar(max),  [F23] nvarchar(max),  [F24] nvarchar(max)
,  [F25] nvarchar(max),  [F26] nvarchar(max),  [F27] nvarchar(max),  [F28] nvarchar(max)
,  [F29] nvarchar(max),  [F30] nvarchar(max),  [F31] nvarchar(max),  [F32] nvarchar(max)
,  [F33] nvarchar(max),  [F34] nvarchar(max),  [F35] nvarchar(max),  [F36] nvarchar(max)
,  [F37] nvarchar(max),  [F38] nvarchar(max),  [F39] nvarchar(max),  [F40] nvarchar(max)
,  [F41] nvarchar(max),  [F42] nvarchar(max),  [F43] nvarchar(max),  [F44] nvarchar(max)
,  [F45] nvarchar(max),  [F46] nvarchar(max),  [F47] nvarchar(max),  [F48] nvarchar(max)
,  [F49] nvarchar(max),  [F50] nvarchar(max),  [F51] nvarchar(max),  [F52] nvarchar(max)
,  [F53] nvarchar(max),  [F54] nvarchar(max),  [F55] nvarchar(max),  [F56] nvarchar(max)
,  [F57] nvarchar(max),  [F58] nvarchar(max),  [F59] nvarchar(max),  [F60] nvarchar(max)
,  [F61] nvarchar(max),  [F62] nvarchar(max),  [F63] nvarchar(max),  [F64] nvarchar(max)
,  [F65] nvarchar(max),  [F66] nvarchar(max),  [F67] nvarchar(max),  [F68] nvarchar(max)
,  [F69] nvarchar(max),  [F70] nvarchar(max),  [F71] nvarchar(max),  [F72] nvarchar(max)
,  [F73] nvarchar(max),  [F74] nvarchar(max),  [F75] nvarchar(max),  [F76] nvarchar(max)
,  [F77] nvarchar(max),  [F78] nvarchar(max),  [F79] nvarchar(max),  [F80] nvarchar(max)
,  [F81] nvarchar(max),  [F82] nvarchar(max),  [F83] nvarchar(max),  [F84] nvarchar(max)
,  [F85] nvarchar(max),  [F86] nvarchar(max),  [F87] nvarchar(max),  [F88] nvarchar(max)
,  [F89] nvarchar(max),  [F90] nvarchar(max),  [F91] nvarchar(max),  [F92] nvarchar(max)
,  [F93] nvarchar(max),  [F94] nvarchar(max),  [F95] nvarchar(max),  [F96] nvarchar(max)
,  [F97] nvarchar(max),  [F98] nvarchar(max),  [F99] nvarchar(max),  [F100] nvarchar(max)
,  [F101] nvarchar(max), [F102] nvarchar(max), [F103] nvarchar(max)
)'
exec (@create_table)
 
-- 	populate table with valuse from excel sheet.
-- 	使用excel表中的值生成表格
declare @populate_table varchar(max) = ('insert into [' + @next_table + '] select * from openrowset(''Microsoft.ACE.OLEDB.12.0'', ''Excel 12.0 Xml; HDR=YES ;Database=' + @import_path + @new_file + ''', ''SELECT * FROM [MSSQL$]'')')
exec (@populate_table)
 
-- 	create final table after import table is created.
-- 	导入表创建完成后创建最终表
declare @final_table    varchar(255) = (select right(@next_table,  9))
declare @final_build    varchar(max) = ('create table [' + @final_table + ']' + char(10) + 
'(
    [Issue]             nvarchar(255) --F1
,   [Severity]          nvarchar(255) --F19
,   [Control_Mapping]       nvarchar(555) --F12
,   [Condition]         nvarchar(max) --F3
,   [Server]            nvarchar(255) --F6
,   [Module]            nvarchar(255) --F5
,   [Version]           nvarchar(255) --F8
,   [Cause]             nvarchar(max) --F9
,   [Recommendation]        nvarchar(max) --F10, F11
,   [Comments]          nvarchar(max) --F11
)')
exec    (@final_build)

-- 	populate final table.   map F# columns from import table to final table and perfrom insert process to final table.
-- 	生成最终表格。将F＃列从导入表映射到最终表，并完成插入最终表的操作
declare @populate_final varchar(max) = ('insert into [' + @final_table + '] 
select [F1], [F19], [F12], [F3], upper([F6]), replace(upper([F5]), ''.MyDomain.com'', ''''), [F8], [F9], ([F10] + ''  '' +  [F11]), [F11]
from [' + @next_table + '] where [F1] not in (''Issue #'') and [F1] is not null order by [F1] asc; update [' + @final_table + '] set [comments] = NULL;')
exec    (@populate_final)

```


[![WorksEveryTime](https://forthebadge.com/images/badges/60-percent-of-the-time-works-every-time.svg)](https://shitday.de/)

## Build-Quality 
| [![Build status](https://ci.appveyor.com/api/projects/status/pjxh5g91jpbh7t84?svg=true)](https://ci.appveyor.com/project/tygerbytes/resourcefitness) | [![Coveralls](https://coveralls.io/repos/github/tygerbytes/ResourceFitness/badge.svg?branch=master)](https://coveralls.io/github/tygerbytes/ResourceFitness?branch=master) | [![nuget](https://img.shields.io/nuget/v/TW.Resfit.Core.svg?style=flat-square)](https://www.nuget.org/packages/TW.Resfit.Core/) |
|-|-|-|

>[![Build history](https://buildstats.info/appveyor/chart/tygerbytes/resourcefitness)](https://ci.appveyor.com/project/tygerbytes/resourcefitness/history)


## Author

- **李聪明 数据 Clever Data**
- **Mike的数据库宝典 Mikes Database Collection**
- **李聪明** "Lee Songming"

[![Gist](https://img.shields.io/badge/Gist-李聪明数据-<COLOR>.svg)](https://gist.github.com/congmingshuju)
[![Twitter](https://img.shields.io/badge/Twitter-mike的数据库宝典-<COLOR>.svg)](https://twitter.com/mikesdatawork?lang=en)
[![Wordpress](https://img.shields.io/badge/Wordpress-mike的数据库宝典-<COLOR>.svg)](https://mikesdatawork.wordpress.com/)

---
## License
[![LicenseCCSA](https://img.shields.io/badge/License-CreativeCommonsSA-<COLOR>.svg)](https://creativecommons.org/share-your-work/licensing-types-examples/)

![Lee Songming](https://github.com/congmingshuju/git-resources/blob/master/images/clever-data-gist-z5.png "李聪明 数据")





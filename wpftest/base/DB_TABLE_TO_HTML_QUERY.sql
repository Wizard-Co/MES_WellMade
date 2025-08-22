Declare @data varchar(4000)
Declare @tmp_data varchar(4000)
Declare @delimiter varchar(4000)
Declare @rVal varchar(4000)
Declare @total_len int

-- 추가할 테이블명 구분자 ',' 사용해서 입력
set @data = '';


set @delimiter = ',';
set @data = REPLACE(@data, ', ', ',');
set @data = @data + @delimiter;
set @tmp_data = @data
set @total_len = len(@tmp_data) - len(replace(@tmp_data, ',', ''))

Declare @i Int, @maxi Int
Declare @j Int, @maxj Int
Declare @sr int
Declare @Output varchar(4000)
Declare @last varchar(155), @current varchar(255), @typ varchar(255), @description varchar(4000)

create Table #Tables  (id int identity(1, 1), Object_id int, Name varchar(155), Type varchar(20), [description] varchar(4000))
create Table #Columns (id int identity(1,1), Name varchar(155), Type Varchar(155), Nullable varchar(2), [description] varchar(4000))
create Table #Fk(id int identity(1,1), Name varchar(155), col Varchar(155), refObj varchar(155), refCol varchar(155))
create Table #Constraint(id int identity(1,1), Name varchar(155), col Varchar(155), definition varchar(1000))
create Table #Indexes(id int identity(1,1), Name varchar(155), Type Varchar(25), cols varchar(1000))

Print '<head>'
Print '<title>::' + DB_name() + '::</title>'
Print '<style>'
    
Print '		body { font-size:9pt; }'
Print '		table { background:#aaa; }'
Print '		tr { background:#fff; }'		

Print '		th {'
Print '		font-size:9pt;'
Print '		background:#eee;'
Print '		}'

Print '		td {'
Print '		font-size:9pt;'
Print '		padding:0 5 0 5;'
Print '		}'

Print '	</style>'
Print '</head>'
Print '<body>'

set nocount on
	begin

	while(charindex(@delimiter, @data) > 0)
	begin
		-- 하나씩 자르기
		set @rVal = LEFT(@data, charindex(@delimiter, @data) - 1)

		-- 잘라낸 나머지로 다시 세팅
		set @data = substring(@data, charindex(@delimiter, @data) + 1, len(@data))

		if(@rVal <> '')
		begin
			insert into #Tables (Object_id, Name, Type, [description])
			Select o.object_id,  '[' + s.name + '].[' + o.name + ']', 
					case when type = 'V' then 'View' when type = 'U' then 'Table' end,  
					cast(p.value as varchar(4000))
					from sys.objects o 
						left outer join sys.schemas s on s.schema_id = o.schema_id 
						left outer join sys.extended_properties p on p.major_id = o.object_id and minor_id = 0 and p.name = 'MS_Description' 
					where type in ('U', 'V') 
					and o.name = @rVal
					order by type, s.name, o.name
		end
	end

	end
Set @maxi = @total_len
set @i = 1

print '<table border="0" cellspacing="0" cellpadding="0" width="550px" align="center"><tr><td colspan="3" style="height:50;font-size:14pt;text-align:center;"><a name="index"></a><b>Index</b></td></tr></table>'
print '<table border="0" cellspacing="1" cellpadding="0" width="550px" align="center"><tr><th>Sr</th><th>Object</th><th>Type</th></tr>' 
While(@i <= @maxi)
begin
	--Index
	select @Output =  '<tr><td align="center">' + Cast((@i) as varchar) + '</td><td><a href="#' + Type + ':' + name + '">' + name + '</a></td><td>' + Type + '</td></tr>' 
			from #Tables where id = @i
	
	print @Output
	set @i = @i + 1
end
print '</table><br />'

set @i = 1
While(@i <= @maxi)
begin
	--table header
	select @Output =  '<tr><th align="left"><a name="' + Type + ':' + name + '"></a><b>' + Type + ':' + name + '</b></th></tr>',  @description = [description]
			from #Tables where id = @i
	
	print '<br /><br /><br /><table border="0" cellspacing="0" cellpadding="0" width="750px"><tr><td align="right"><a href="#index">Index</a></td></tr>'
	print @Output
	print '</table><br />'
	print '<table border="0" cellspacing="0" cellpadding="0" width="750px"><tr><td><b>Description</b></td></tr><tr><td>' + isnull(@description, '') + '</td></tr></table><br />' 

	--table columns
	truncate table #Columns 
	begin
		insert into #Columns  (Name, Type, Nullable, [description])
		Select c.name, 
					type_name(user_type_id) + (
					case when (type_name(user_type_id) = 'varchar' or type_name(user_type_id) = 'nvarchar' or type_name(user_type_id) ='char' or type_name(user_type_id) ='nchar')
						then '(' + cast(max_length as varchar) + ')' 
						when type_name(user_type_id) = 'decimal'  
							then '(' + cast([precision] as varchar) + ',' + cast(scale as varchar)   + ')' 
					else ''
					end				
					), 
					case when is_nullable = 1 then 'Y' else 'N'  end,
					cast(p.value as varchar(4000))
		from sys.columns c
				inner join #Tables t on t.object_id = c.object_id
				left outer join sys.extended_properties p on p.major_id = c.object_id and p.minor_id  = c.column_id and p.name = 'MS_Description' 
		where t.id = @i
		order by c.column_id
	end

	Set @maxj =   @@rowcount
	set @j = 1

	print '<table border="0" cellspacing="0" cellpadding="0" width="750px"><tr><td><b>Table Columns</b></td></tr></table>' 
	print '<table border="0" cellspacing="1" cellpadding="0" width="750px"><tr><th>Sr.</th><th>Name</th><th>Datatype</th><th>Nullable</th><th>Description</th></tr>' 
	
	While(@j <= @maxj)
	begin
		select @Output = '<tr><td width="20px" align="center">' + Cast((@j) as varchar) + '</td><td width="150px">' + isnull(name,'')  + '</td><td width="150px">' +  upper(isnull(Type,'')) + '</td><td width="50px" align="center">' + isnull(Nullable,'N') + '</td><td>' + isnull([description],'') + '</td></tr>' 
			from #Columns  where id = @j
		
		print 	@Output 	
		Set @j = @j + 1;
	end

	print '</table><br />'

	--reference key
	truncate table #FK	
	begin
		insert into #FK  (Name, col, refObj, refCol)
		select f.name, COL_NAME (fc.parent_object_id, fc.parent_column_id) , object_name(fc.referenced_object_id) , COL_NAME (fc.referenced_object_id, fc.referenced_column_id)     
		from sys.foreign_keys f
			inner  join  sys.foreign_key_columns  fc  on f.object_id = fc.constraint_object_id	
			inner join #Tables t on t.object_id = f.parent_object_id
		where t.id = @i
		order by f.name
	end

	Set @maxj =   @@rowcount
	set @j = 1
	if (@maxj >0)
	begin

		print '<table border="0" cellspacing="0" cellpadding="0" width="750px"><tr><td><b>Refrence Keys</b></td></tr></table>' 
		print '<table border="0" cellspacing="1" cellpadding="0" width="750px"><tr><th>Sr.</th><th>Name</th><th>Column</th><th>Reference To</th></tr>' 

		While(@j <= @maxj)
		begin

			select @Output = '<tr><td width="20px" align="center">' + Cast((@j) as varchar) + '</td><td width="150px">' + isnull(name,'')  + '</td><td width="150px">' +  isnull(col,'') + '</td><td>[' + isnull(refObj,'N') + '].[' +  isnull(refCol,'N') + ']</td></tr>' 
				from #FK  where id = @j

			print @Output
			Set @j = @j + 1;
		end

		print '</table><br />'
	end

	--Default Constraints 
	truncate table #Constraint
	begin
		insert into #Constraint  (Name, col, definition)
		select c.name,  col_name(parent_object_id, parent_column_id), c.definition 
		from sys.default_constraints c
			inner join #Tables t on t.object_id = c.parent_object_id
		where t.id = @i
		order by c.name
	end

	Set @maxj =   @@rowcount
	set @j = 1
	if (@maxj >0)
	begin

		print '<table border="0" cellspacing="0" cellpadding="0" width="750px"><tr><td><b>Default Constraints</b></td></tr></table>' 
		print '<table border="0" cellspacing="1" cellpadding="0" width="750px"><tr><th>Sr.</th><th>Name</th><th>Column</th><th>Value</th></tr>' 

		While(@j <= @maxj)
		begin

			select @Output = '<tr><td width="20px" align="center">' + Cast((@j) as varchar) + '</td><td width="250px">' + isnull(name,'')  + '</td><td width="150px">' +  isnull(col,'') + '</td><td>' +  isnull(definition,'') + '</td></tr>' 
				from #Constraint  where id = @j

			print @Output
			Set @j = @j + 1;
		end

	print '</table><br />'
	end


	--Check  Constraints
	truncate table #Constraint
	begin
		insert into #Constraint  (Name, col, definition)
		select c.name,  col_name(parent_object_id, parent_column_id), definition 
		from sys.check_constraints c
			inner join #Tables t on t.object_id = c.parent_object_id
		where t.id = @i
		order by c.name
	end

	Set @maxj =   @@rowcount
	
	set @j = 1
	if (@maxj >0)
	begin

		print '<table border="0" cellspacing="0" cellpadding="0" width="750px"><tr><td><b>Check  Constraints</b></td></tr></table>' 
		print '<table border="0" cellspacing="1" cellpadding="0" width="750px"><tr><th>Sr.</th><th>Name</th><th>Column</th><th>Definition</th></tr>' 

		While(@j <= @maxj)
		begin

			select @Output = '<tr><td width="20px" align="center">' + Cast((@j) as varchar) + '</td><td width="250px">' + isnull(name,'')  + '</td><td width="150px">' +  isnull(col,'') + '</td><td>' +  isnull(definition,'') + '</td></tr>' 
				from #Constraint  where id = @j
			print @Output 
			Set @j = @j + 1;
		end

		print '</table><br />'
	end


	--Triggers 
	truncate table #Constraint	
	begin
		insert into #Constraint  (Name)
		select tr.name
		from sys.triggers tr
			inner join #Tables t on t.object_id = tr.parent_id
		where t.id = @i
		order by tr.name
	end

	Set @maxj =   @@rowcount
	
	set @j = 1
	if (@maxj >0)
	begin

		print '<table border="0" cellspacing="0" cellpadding="0" width="750px"><tr><td><b>Triggers</b></td></tr></table>' 
		print '<table border="0" cellspacing="1" cellpadding="0" width="750px"><tr><th>Sr.</th><th>Name</th><th>Description</th></tr>' 

		While(@j <= @maxj)
		begin
			select @Output = '<tr><td width="20px" align="center">' + Cast((@j) as varchar) + '</td><td width="150px">' + isnull(name,'')  + '</td><td></td></tr>' 
				from #Constraint  where id = @j
			print @Output 
			Set @j = @j + 1;
		end

		print '</table><br />'
	end

	--Indexes 
	truncate table #Indexes
	begin
		insert into #Indexes  (Name, type, cols)
		select i.name, case when i.type = 0 then 'Heap' when i.type = 1 then 'Clustered' else 'Nonclustered' end,  col_name(i.object_id, c.column_id)
			from sys.indexes i 
				inner join sys.index_columns c on i.index_id = c.index_id and c.object_id = i.object_id 
				inner join #Tables t on t.object_id = i.object_id
			where t.id = @i
			order by i.name, c.column_id
	end

	Set @maxj =   @@rowcount
	
	set @j = 1
	set @sr = 1
	if (@maxj >0)
	begin

		print '<table border="0" cellspacing="0" cellpadding="0" width="750px"><tr><td><b>Indexes</b></td></tr></table>' 
		print '<table border="0" cellspacing="1" cellpadding="0" width="750px"><tr><th>Sr.</th><th>Name</th><th>Type</th><th>Columns</th></tr>' 
		set @Output = ''
		set @last = ''
		set @current = ''
		While(@j <= @maxj)
		begin
			select @current = isnull(name,'') from #Indexes  where id = @j
					 
			if @last <> @current  and @last <> ''
				begin	
				print '<tr><td width="20px" align="center">' + Cast((@sr) as varchar) + '</td><td width="150px">' + @last + '</td><td width="150px">' + @typ + '</td><td>' + @Output  + '</td></tr>' 
				set @Output  = ''
				set @sr = @sr + 1
				end
			
				
			select @Output = @Output + cols + '<br />' , @typ = type
					from #Indexes  where id = @j
			
			set @last = @current 	
			Set @j = @j + 1;
		end
		if @Output <> ''
				begin	
				print '<tr><td width="20px" align="center">' + Cast((@sr) as varchar) + '</td><td width="150px">' + @last + '</td><td width="150px">' + @typ + '</td><td>' + @Output  + '</td></tr>' 
				end

		print '</table><br />'
	end

    Set @i = @i + 1;
	--Print @Output 
end


Print '</body>'
Print '</html>'

drop table #Tables
drop table #Columns
drop table #FK
drop table #Constraint
drop table #Indexes 
set nocount off




[System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") | Out-Null
#Below variables should be modified as per the Machine settings#
###########################################################################################################
$SqlConnectionString = "Server = <servername>; Database = master; User ID= <username>; Password= <password>"
###########################################################################################################


$SqlCommand = New-Object System.Data.SqlClient.SqlCommand
$SqlCommand.CommandTimeout = 0
$SqlConnection = New-Object System.Data.SqlClient.SqlConnection

$SqlConnection.ConnectionString =  $SqlConnectionString
$app = new-object -com Shell.Application
$fd = New-Object system.windows.forms.openfiledialog
$fd.InitialDirectory = 'c:'
$fd.MultiSelect = $false

$CurrentDate = Get-Date -format "dd-MMM-yyyy"





function SetSingleUserMode($DbName){

        write-host "Setting $DBName to Single User Mode..."
        $SqlCommand.CommandText = "ALTER DATABASE $DbName SET SINGLE_USER WITH ROLLBACK IMMEDIATE;"
        try{
            $SqlCommand.ExecuteNonQuery() | Out-Null
        }
        catch{
            write-host $_.Exception.Message -foregroundcolor "red"
            $SqlConnection.Close()
            break
        }


}

function SetMultipleUserMode($DbName){

        write-host "Setting $DBName to Multi User Mode..."
        $SqlCommand.CommandText = "ALTER DATABASE $DbName SET MULTI_USER;"
        try{
            $SqlCommand.ExecuteNonQuery() | Out-Null
        }
        catch{
            write-host $_.Exception.Message -foregroundcolor "red"
            $SqlConnection.Close()
            break
        }

}

function SetToOfflineMode($DbName){

        write-host "Setting $DBName to Offline mode..."
        $SqlCommand.CommandText = "ALTER DATABASE $DbName SET OFFLINE WITH ROLLBACK IMMEDIATE;"
        try{
            $SqlCommand.ExecuteNonQuery() | Out-Null
        }
        catch{
            write-host $_.Exception.Message -foregroundcolor "red"
            $SqlConnection.Close()
            break
        }

}
function SetToOnlineMode($DbName){
        
        write-host "Setting $DBName to Online mode..."
        $SqlCommand.CommandText = "ALTER DATABASE $DbName SET ONLINE;"
        try{
            $SqlCommand.ExecuteNonQuery() | Out-Null
        }
        catch{
            write-host $_.Exception.Message -foregroundcolor "red"
            $SqlConnection.Close()
            break
        }
}

function checkIfDbExists($DbName)
{

        $SqlCommand.CommandText = "select 1 from sys.databases where name = '$DbName';"
        $reader = $SqlCommand.ExecuteReader()
        try{
            if($reader.HasRows)
            {
                $reader.Close()
            }
            else{
                write-host "Database : $DbName doesn't exist, Script will exit now" -foregroundcolor "red"
                $reader.Close()
                $SqlConnection.Close()
                break
            }
        }
        catch{
            write-host $_.Exception.Message -foregroundcolor "red"
            $SqlConnection.Close()
            break
        }


}


function RestoreDBFromFile{
        $SourceDBName = Read-Host "Enter the Source DB name whose backup file you are using "
        $DbName = read-host "Enter the Destination DB name , you wish to restore"

        try{
            $SqlConnection.Open()
        }
        catch{
            write-host $_.Exception.Message -foregroundcolor "red"
            $SqlConnection.Close()
            break
        }

        $SqlCommand.Connection = $SqlConnection

        checkIfDbExists($DbName)

        checkIfDbExists($SourceDBName)

        $result = $fd.ShowDialog()

        if($result -ne "OK"){
            write-host "You didn't select a valid file"
            break
        }
        else{
           $bckFileName = $fd.FileName
        }


        SetSingleUserMode($DbName)
        SetToOfflineMode($DbName)

        write-host "Restoring $DBName from $bckFileName..."
        #$SqlCommand.CommandText = "RESTORE DATABASE $DbName  FROM DISK = '$bckFileName';"

        $SqlCommand.CommandText = "
        declare  @SrcDataName varchar(200)
	   ,@SrcLogName varchar(200)
	   ,@DestDataPhyLoc varchar(max)
	   ,@DestLogPhyLoc varchar(max)
	   ,@SourceDBName varchar(200)
	   ,@DestDBName varchar(200)
	   ,@DestDataName varchar(200)
	   ,@DestLogName varchar(200)
	   ,@BckUpPath varchar(max)
	   ,@v_error_msg    NVARCHAR(2048)  = NULL

            set @SourceDBName	= '$SourceDBName'
            set @DestDBName	    = '$DbName'
            set @BckUpPath		= '$bckFileName'




            select @SrcDataName =  mf.name from sys.master_files mf
            join sys.databases db on mf.database_id = db.database_id
            where db.name = @SourceDBName
            and type_desc = 'ROWS'

            select @SrcLogName = mf.name from sys.master_files mf
            join sys.databases db on mf.database_id = db.database_id
            where db.name = @SourceDBName
            and type_desc = 'Log'





            select @DestDataPhyLoc =  mf.physical_name from sys.master_files mf
            join sys.databases db on mf.database_id = db.database_id
            where db.name = @DestDBName
            and type_desc = 'ROWS'

            select @DestLogPhyLoc =  mf.physical_name from sys.master_files mf
            join sys.databases db on mf.database_id = db.database_id
            where db.name = @DestDBName
            and type_desc = 'LOG'

            print @SrcDataName
            Print @SrcLogName
            print @DestDataPhyLoc
            Print @DestLogPhyLoc

            print ('ALTER DATABASE ' + @DestDBName + ' SET SINGLE_USER WITH ROLLBACK IMMEDIATE;')
            EXEC ('ALTER DATABASE ' + @DestDBName + ' SET SINGLE_USER WITH ROLLBACK IMMEDIATE;')

            BEGIN TRY


            print ('
            RESTORE DATABASE ' + @DestDBName + ' FROM DISK = ''' + @BckUpPath + ''' 
            WITH REPLACE,
            MOVE ''' + @SrcDataName + ''' TO ''' + @DestDataPhyLoc + ''', MOVE ''' + @SrcLogName + ''' TO ''' + @DestLogPhyLoc + ''''
            )

            Exec ('
            RESTORE DATABASE ' + @DestDBName + ' FROM DISK = ''' + @BckUpPath + ''' 
            WITH REPLACE,
            MOVE ''' + @SrcDataName + ''' TO ''' + @DestDataPhyLoc + ''', MOVE ''' + @SrcLogName + ''' TO ''' + @DestLogPhyLoc + ''''
            )

            select @DestDataName =  mf.name from sys.master_files mf
            join sys.databases db on mf.database_id = db.database_id
            where db.name = @DestDBName
            and type_desc = 'ROWS'

            select @DestLogName = mf.name from sys.master_files mf
            join sys.databases db on mf.database_id = db.database_id
            where db.name = @DestDBName
            and type_desc = 'Log'

            IF(@DestDataName <> @DestDBName + '_Data')
            BEGIN
            print ('ALTER DATABASE ' + @DestDBName + ' MODIFY FILE(NAME = N''' + @DestDataName + ''', NEWNAME = N''' + @DestDBName + '_Data'')')
            EXEC ('ALTER DATABASE ' + @DestDBName + ' MODIFY FILE(NAME = N''' + @DestDataName + ''', NEWNAME = N''' + @DestDBName + '_Data'')')
            END
            
            IF(@DestLogName <> @DestDBName + '_Log')
            BEGIN
            print ('ALTER DATABASE ' + @DestDBName + ' MODIFY FILE(NAME = N''' + @DestLogName + ''', NEWNAME = N''' + @DestDBName + '_Log'')')
            EXEC ('ALTER DATABASE ' + @DestDBName + ' MODIFY FILE(NAME = N''' + @DestLogName + ''', NEWNAME = N''' + @DestDBName + '_Log'')')
            END

            print ('ALTER DATABASE ' + @DestDBName + ' SET MULTI_USER;')
            Exec ('ALTER DATABASE ' + @DestDBName + ' SET MULTI_USER;')

            END TRY
            BEGIN CATCH
                print ('ALTER DATABASE ' + @DestDBName + ' SET MULTI_USER;')
                Exec ('ALTER DATABASE ' + @DestDBName + ' SET MULTI_USER;')
                SET @v_error_msg = FORMATMESSAGE('%d|%d|%s', ERROR_NUMBER(), ERROR_LINE(), ERROR_MESSAGE());
                THROW 51000, @v_error_msg, 1
    
            END CATCH

        "




        try{
            $SqlCommand.ExecuteNonQuery() | Out-Null
        }
        catch{
            write-host $_.Exception.Message -foregroundcolor "red"

            SetToOnlineMode($DbName)
            SetMultipleUserMode($DbName)

            break
        }

        SetToOnlineMode($DbName)
        SetMultipleUserMode($DbName)

        try{
            $SqlConnection.Close()
        }
        catch{
            write-host $_.Exception.Message -foregroundcolor "red"
            break
        }
        Write-Host "Database $DbName successfully restored from $bckFileName" -ForegroundColor Green
       
}


function BackUpDB{
        $DbName = read-host "Enter the DB name whose backup file (.bak) you wish to save"
        
       

        try{
            $SqlConnection.Open()
        }
        catch{
            write-host $_.Exception.Message -foregroundcolor "red"
            $SqlConnection.Close()
            break
        }

        $SqlCommand.Connection = $SqlConnection
        checkIfDbExists($DbName)

         $folder = $app.BrowseForFolder(0, "Select Folder for saving the BackUp File", 0, "C:\")
        if ($folder.Self.Path -eq "") {
           # write-host "You selected " $folder.Self.Path
           write-host "You didn't select a valid folder, the script will exit"
           break
        }
        if ($folder.Self.Path -eq $null) {
           # write-host "You selected " $folder.Self.Path
           write-host "You didn't select a valid folder, the script will exit"
           break
        }
        $filelocation = $folder.Self.Path

        SetSingleUserMode($DbName)
       

        write-host "Creating the backup file for $DbName ..."
        $SqlCommand.CommandText = "BACKUP DATABASE $DbName TO DISK = '$filelocation\$DbName-$CurrentDate.bak' WITH FORMAT,NAME = '$DbName';"
        try{
            $SqlCommand.ExecuteNonQuery() | Out-Null
        }
        catch{
            write-host $_.Exception.Message -foregroundcolor "red"
            $SqlConnection.Close()
            break
        }
        
        SetMultipleUserMode($DbName)

        write-host "Database $DbName is backed up at $filelocation\$DbName-$CurrentDate.bak" -ForegroundColor Green

        try{
            $SqlConnection.Close()
        }
        catch{
            write-host $_.Exception.Message -foregroundcolor "red"
            break
        }

}


function RestoreFromDB{
        $SourceDbName = read-host "Enter the Source DB name whose backup you need"
        $DestDbName = read-host "Enter the Destination DB name , you wish to restore to"

       
        try{
            $SqlConnection.Open()
        }
        catch{
            write-host $_.Exception.Message -foregroundcolor "red"
            $SqlConnection.Close()
            break
        }

        $SqlCommand.Connection = $SqlConnection

        checkIfDbExists($SourceDbName)
        checkIfDbExists($DestDbName)

        SetSingleUserMode($SourceDbName)
       

        #write-host "Creating the backup file for $SourceDbName ..."
        $SqlCommand.CommandText = "BACKUP DATABASE $SourceDbName TO DISK = 'C:\$SourceDbName-$CurrentDate.bak' WITH FORMAT,NAME = '$SourceDbName';"

        
        try{
            $SqlCommand.ExecuteNonQuery() | Out-Null
        }
        catch{
            write-host $_.Exception.Message -foregroundcolor "red"
            $SqlConnection.Close()
            break
        }


 



        
        SetMultipleUserMode($SourceDbName)

        #write-host "Database $SourceDbName is backed up at $filelocation\$SourceDbName-$CurrentDate.bak"

        SetSingleUserMode($DestDbName)
        SetToOfflineMode($DestDbName)

        
        #$SqlCommand.CommandText = "RESTORE DATABASE $DestDbName  FROM DISK = 'C:\$SourceDbName-$CurrentDate.bak';"
        $SqlCommand.CommandText = "

        declare  @SrcDataName varchar(200)
	   ,@SrcLogName varchar(200)
	   ,@DestDataPhyLoc varchar(max)
	   ,@DestLogPhyLoc varchar(max)
	   ,@SourceDBName varchar(200)
	   ,@DestDBName varchar(200)
	   ,@DestDataName varchar(200)
	   ,@DestLogName varchar(200)
	   ,@BckUpPath varchar(max)
	   ,@v_error_msg    NVARCHAR(2048)  = NULL

            set @SourceDBName	= '$SourceDbName'
            set @DestDBName	    = '$DestDbName'
            set @BckUpPath		= 'C:\$SourceDbName-$CurrentDate.bak'




            select @SrcDataName =  mf.name from sys.master_files mf
            join sys.databases db on mf.database_id = db.database_id
            where db.name = @SourceDBName
            and type_desc = 'ROWS'

            select @SrcLogName = mf.name from sys.master_files mf
            join sys.databases db on mf.database_id = db.database_id
            where db.name = @SourceDBName
            and type_desc = 'Log'





            select @DestDataPhyLoc =  mf.physical_name from sys.master_files mf
            join sys.databases db on mf.database_id = db.database_id
            where db.name = @DestDBName
            and type_desc = 'ROWS'

            select @DestLogPhyLoc =  mf.physical_name from sys.master_files mf
            join sys.databases db on mf.database_id = db.database_id
            where db.name = @DestDBName
            and type_desc = 'LOG'

            print @SrcDataName
            Print @SrcLogName
            print @DestDataPhyLoc
            Print @DestLogPhyLoc

            print ('ALTER DATABASE ' + @DestDBName + ' SET SINGLE_USER WITH ROLLBACK IMMEDIATE;')
            EXEC ('ALTER DATABASE ' + @DestDBName + ' SET SINGLE_USER WITH ROLLBACK IMMEDIATE;')

            BEGIN TRY


            print ('
            RESTORE DATABASE ' + @DestDBName + ' FROM DISK = ''' + @BckUpPath + ''' 
            WITH REPLACE,
            MOVE ''' + @SrcDataName + ''' TO ''' + @DestDataPhyLoc + ''', MOVE ''' + @SrcLogName + ''' TO ''' + @DestLogPhyLoc + ''''
            )

            Exec ('
            RESTORE DATABASE ' + @DestDBName + ' FROM DISK = ''' + @BckUpPath + ''' 
            WITH REPLACE,
            MOVE ''' + @SrcDataName + ''' TO ''' + @DestDataPhyLoc + ''', MOVE ''' + @SrcLogName + ''' TO ''' + @DestLogPhyLoc + ''''
            )


            select @DestDataName =  mf.name from sys.master_files mf
            join sys.databases db on mf.database_id = db.database_id
            where db.name = @DestDBName
            and type_desc = 'ROWS'

            select @DestLogName = mf.name from sys.master_files mf
            join sys.databases db on mf.database_id = db.database_id
            where db.name = @DestDBName
            and type_desc = 'Log'

            IF(@DestDataName <> @DestDBName + '_Data')
            BEGIN
            print ('ALTER DATABASE ' + @DestDBName + ' MODIFY FILE(NAME = N''' + @DestDataName + ''', NEWNAME = N''' + @DestDBName + '_Data'')')
            EXEC ('ALTER DATABASE ' + @DestDBName + ' MODIFY FILE(NAME = N''' + @DestDataName + ''', NEWNAME = N''' + @DestDBName + '_Data'')')
            END
            
            IF(@DestLogName <> @DestDBName + '_Log')
            BEGIN
            print ('ALTER DATABASE ' + @DestDBName + ' MODIFY FILE(NAME = N''' + @DestLogName + ''', NEWNAME = N''' + @DestDBName + '_Log'')')
            EXEC ('ALTER DATABASE ' + @DestDBName + ' MODIFY FILE(NAME = N''' + @DestLogName + ''', NEWNAME = N''' + @DestDBName + '_Log'')')
            END


            print ('ALTER DATABASE ' + @DestDBName + ' SET MULTI_USER;')
            Exec ('ALTER DATABASE ' + @DestDBName + ' SET MULTI_USER;')

            END TRY
            BEGIN CATCH
                print ('ALTER DATABASE ' + @DestDBName + ' SET MULTI_USER;')
                Exec ('ALTER DATABASE ' + @DestDBName + ' SET MULTI_USER;')
                SET @v_error_msg = FORMATMESSAGE('%d|%d|%s', ERROR_NUMBER(), ERROR_LINE(), ERROR_MESSAGE());
                THROW 51000, @v_error_msg, 1
    
            END CATCH

        "
        try{
            $SqlCommand.ExecuteNonQuery() | Out-Null
        }
        catch{
            write-host $_.Exception.Message -foregroundcolor "red" 

            SetToOnlineMode($DestDbName)
            SetMultipleUserMode($DestDbName)

            break
        }

        SetToOnlineMode($DestDbName)
        SetMultipleUserMode($DestDbName)

        try{
            $SqlConnection.Close()
        }
        catch{
            write-host $_.Exception.Message -foregroundcolor "red"
            break
        }
        write-host "$DestDbName is restored from $SourceDbName" -ForegroundColor Green


}

function createDbFromExistingDB(){
        $SourceDbName = read-host "Enter the Source DB name whose copy you need"
        $DestDbName = read-host   "Enter the name of the new DB"
                
        if($DestDbName -eq ""){
            write-host "You didn't enter a valid name for new db, program will exit now" -ForegroundColor "red"
            break
        }
       
        try{
            $SqlConnection.Open()
        }
        catch{
            write-host $_.Exception.Message -foregroundcolor "red"
            $SqlConnection.Close()
            break
        }

        $SqlCommand.Connection = $SqlConnection
        checkIfDbExists($SourceDbName)
        $SqlScript = "
                        DECLARE	@DBToBeBackedUp VARCHAR(200)
	                           , @NewDBName VARCHAR(200)
	                           , @LogicalDBName VARCHAR(100)
	                           , @LogicalDBLogName VARCHAR(100)
	                           , @LogicalDBNameloc VARCHAR(200)
	                           , @LogicalDBLogNameloc VARCHAR(200);

                        SET @DBToBeBackedUp	    = '$SourceDbName';
                        SET @NewDBName		    = '$DestDbName';

                        IF NOT EXISTS
                        (
                            SELECT 1
                            FROM [sys].[databases]
                            WHERE [name] = @DBToBeBackedUp
                        )
                            BEGIN
                                PRINT('DB to be backed up i.e. '+@DBToBeBackedUp+' doesn''t exist');
                            END;
                        ELSE
                            BEGIN
                                BEGIN
                                    IF EXISTS
                                    (
                                        SELECT 1
                                        FROM [sys].[databases]
                                        WHERE [name] = @NewDBName
                                    )
                                        BEGIN
                                            PRINT('The New DB i.e. '+@NewDBName+' already exists');
                                        END;
                                    ELSE
                                        BEGIN


                                                SELECT @LogicalDBName = [mf].[name],
                                                       @LogicalDBNameloc = SUBSTRING([mf].[physical_name], 0, LEN([physical_name])-LEN(REVERSE(SUBSTRING(REVERSE([mf].[physical_name]), 0, CHARINDEX('\', REVERSE([mf].[physical_name])))))+1)
                                                FROM [sys].[databases] [db]
                                                     JOIN [sys].[master_files] [mf] ON [db].[database_id] = [mf].[database_id]
                                                WHERE [db].[name] = @DBToBeBackedUp
                                                      AND [type_desc] = 'ROWS';

                                                SELECT @LogicalDBLogName = [mf].[name],
                                                       @LogicalDBLogNameloc = SUBSTRING([mf].[physical_name], 0, LEN([physical_name])-LEN(REVERSE(SUBSTRING(REVERSE([mf].[physical_name]), 0, CHARINDEX('\', REVERSE([mf].[physical_name])))))+1)
                                                FROM [sys].[databases] [db]
                                                     JOIN [sys].[master_files] [mf] ON [db].[database_id] = [mf].[database_id]
                                                WHERE [db].[name] = @DBToBeBackedUp
                                                      AND [type_desc] = 'LOG';

                                                PRINT('Getting the backup of ['+@DBToBeBackedUp+'] and saving to '+@LogicalDBNameloc+@NewDBName+'.bak');
                                                EXEC ('BACKUP DATABASE ['+@DBToBeBackedUp+'] TO DISK = '''+@LogicalDBNameloc+@NewDBName+'.bak''');
                                                PRINT(@LogicalDBNameloc+@NewDBName+'.bak generated');

                                                PRINT('Restoring the backup to ['+@NewDBName+']');
                                                EXEC ('
						                          RESTORE DATABASE ['+@NewDBName+'] FROM DISK = '''+@LogicalDBNameloc+@NewDBName+'.bak''
						                          WITH 
						                            MOVE '''+@LogicalDBName+'''  TO '''+@LogicalDBNameloc+@NewDBName+'.mdf''
						                          , MOVE '''+@LogicalDBLogName+''' TO '''+@LogicalDBLogNameloc+@NewDBName+'_log.ldf''');

                                                PRINT('Restore finished');

                                                PRINT('Aligning the logical names');
                                                EXEC ('
						                          ALTER DATABASE ['+@NewDBName+'] MODIFY FILE(NAME = N'''+@LogicalDBName+''', NEWNAME = N'''+@NewDBName+''');
						                          ');
                                                EXEC ('
						                          ALTER DATABASE ['+@NewDBName+'] MODIFY FILE(NAME = N'''+@LogicalDBLogName+''', NEWNAME = N'''+@NewDBName+'_log'');
						                          ');

                                                PRINT('Logical names are updated');

                                        END;
                                END;
                            END;
                        "

        write-host "Creating the new DB $DestDbName from $SourceDbName"
        $SqlCommand.CommandText = $SqlScript
        try{
            $SqlCommand.ExecuteNonQuery() | Out-Null
        }
        catch{
            write-host $_.Exception.Message -foregroundcolor "red" 
            break
        }
        write-host "Successfully created $DestDbName from $SourceDbName" -ForegroundColor Green
}


write-host "Enter the task you wish to perform today

(1) Take Backup of DB to .bak file
(2) Restore DB from existing .bak file
(3) Restore Database to Database
(4) Create New Database from existing database
(5) Clicked by Mistake :)
" -foregroundcolor Yellow
$choice = Read-Host(
"
Enter your choice [1-5]
")  

if ($choice -eq '1'){
    cls
    write-host "**********************************************************************" -foregroundcolor Yellow
    write-host "**************BackUp Database to Disk Device(.bak files)**************" -foregroundcolor Yellow
    write-host "**********************************************************************" -foregroundcolor Yellow
    BackUpDB                                                                                            
}                                                                                                       
ElseIf($choice -eq '2'){                                                                                
    cls                                                                                                 
    write-host "**********************************************************************" -foregroundcolor Yellow
    write-host "*************Restore Database from backUp File(.bak file)*************" -foregroundcolor Yellow
    write-host "**********************************************************************" -foregroundcolor Yellow
    RestoreDBFromFile                                                                                 
}                                                                                                     
ElseIf($choice -eq '3'){                                                                              
    cls                                                                                               
    write-host "**********************************************************************" -foregroundcolor Yellow
    write-host "********Restore existing Database from Database (In same Server)******" -foregroundcolor Yellow
    write-host "**********************************************************************" -foregroundcolor Yellow
    RestoreFromDB                                                                                      
}                                                                                                      
ElseIf($choice -eq '4'){                                                                               
    cls                                                                                                
    write-host "**********************************************************************" -foregroundcolor Yellow
    write-host "**********Creating new Database from existing db (copy of db)*********" -foregroundcolor Yellow
    write-host "**********************************************************************" -foregroundcolor Yellow
    createDbFromExistingDB
}
ElseIf($choice -eq '5'){                                                                               
    cls                                                                                                
    write-host "**********************************************************************" -foregroundcolor Yellow
    write-host "*******************No issues, have a great day Bye!!******************" -foregroundcolor Yellow
    write-host "**********************************************************************" -foregroundcolor Yellow
    break
}
else{
    write-host "You didn't enter valid input, script will exit now"
    break
}


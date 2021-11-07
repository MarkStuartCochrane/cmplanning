<#  Summary of things planned to run - mark.cochrane@insight.com - jul-nov 2021
    update : change SQL access removing req for SQL module (thanks to Hugo Hucorne - 5/11/21)

    if SQL is on another box, you'll need to specify it in $sqlparms

    Backlog

        1) find a way to summarize client actions scheduled by the client settings. 
            Something like 
            Next WSUS evals : 15.000 at 10h, 12.000 at 16h...
            Next HINV : divide number of clients by frequency 
            ...
            So far, i only found Get-CMResultantSettings but not what's underlying. Way too slow.

        2) Sometimes, hangs the PSH ISE. PRobably due to a bad cleanup of WPF objects/runspace
    
    if you improve things, let me know so I can learn from your ideas

#>

#region Common CCM Functions

Function ConvertFrom-CCMSchedule {
    <#
    .SYNOPSIS
        Convert Configuration Manager Schedule Strings
    .DESCRIPTION
        This function will take a Configuration Manager Schedule String and convert it into a readable object, including
        the calculated description of the schedule
    .PARAMETER ScheduleString
        Accepts an array of strings. This should be a schedule string in the SCCM format
    .EXAMPLE
        PS C:\> ConvertFrom-CCMSchedule -ScheduleString 1033BC7B10100010
        SmsProviderObjectPath : SMS_ST_RecurInterval
        DayDuration : 0
        DaySpan : 2
        HourDuration : 2
        HourSpan : 0
        IsGMT : False
        MinuteDuration : 59
        MinuteSpan : 0
        StartTime : 11/19/2019 1:04:00 AM
        Description : Occurs every 2 days effective 11/19/2019 1:04:00 AM
    .NOTES
        This function was created to allow for converting SCCM schedule strings without relying on the SDK / Site Server
        It also happens to be a TON faster than the Convert-CMSchedule cmdlet and the CIM method on the site server
    #>
    Param(
        [Parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true)]
        [Alias('Schedules')]
        [string[]]$ScheduleString
    )
    begin {
        #region TypeMap for returning readable window type
        $TypeMap = @{
            1 = 'SMS_ST_NonRecurring'
            2 = 'SMS_ST_RecurInterval'
            3 = 'SMS_ST_RecurWeekly'
            4 = 'SMS_ST_RecurMonthlyByWeekday'
            5 = 'SMS_ST_RecurMonthlyByDate'
        }
        #endregion TypeMap for returning readable window type
    
        #region function to return a formatted day such as 1st, 2nd, or 3rd
        function Get-FancyDay {
            <#
                .SYNOPSIS
                Convert the input 'Day' integer to a 'fancy' value such as 1st, 2nd, 4d, 4th, etc.
            #>
            param(
                [int]$Day
            )
            $Suffix = switch -regex ($Day) {
                '1(1|2|3)$' {
                    'th'
                    break
                }
                '.?1$' {
                    'st'
                    break
                }
                '.?2$' {
                    'nd'
                    break
                }
                '.?3$' {
                    'rd'
                    break
                }
                default {
                    'th'
                    break
                }
            }
            [string]::Format('{0}{1}', $Day, $Suffix)
        }
        #endregion function to return a formatted day such as 1st, 2nd, or 3rd
    }
    process {
        # we will split the schedulestring input into 16 characters, as some are stored as multiple in one
        foreach ($Schedule in ($ScheduleString -split '(\w{16})' | Where-Object { $_ })) {
            $MW = [System.Collections.Specialized.OrderedDictionary]::new()
    
            # the first 8 characters are the Start of the MW, while the last 8 characters are the recurrence schedule
            $Start = $Schedule.Substring(0, 8)
            $Recurrence = $Schedule.Substring(8, 8)
    
            switch ($Start) {
                '00012000' {
                    # this is as 'simple' schedule, such as a CI that 'runs once a day' or 'every 8 hours'
                }
                default {
                    # Convert to binary string and pad left with 0 to ensure 32 character length for consistent parsing
                    $binaryStart = [Convert]::ToString([int64]"0x$Start".ToString(), 2).PadLeft(32, 48)
    
                    # Collect timedata and ensure we pad left with 0 to ensure 2 character length
                    [string]$StartMinute = ([Convert]::ToInt32($binaryStart.Substring(0, 6), 2).ToString()).PadLeft(2, 48)
                    [string]$MinuteDuration = [Convert]::ToInt32($binaryStart.Substring(26, 6), 2).ToString()
                    [string]$StartHour = ([Convert]::ToInt32($binaryStart.Substring(6, 5), 2).ToString()).PadLeft(2, 48)
                    [string]$StartDay = ([Convert]::ToInt32($binaryStart.Substring(11, 5), 2).ToString()).PadLeft(2, 48)
                    [string]$StartMonth = ([Convert]::ToInt32($binaryStart.Substring(16, 4), 2).ToString()).PadLeft(2, 48)
                    [String]$StartYear = [Convert]::ToInt32($binaryStart.Substring(20, 6), 2) + 1970
    
                    # set our StartDateTimeObject variable by formatting all our calculated datetime components and piping to Get-Date
                    $StartDateTimeString = [string]::Format('{0}-{1}-{2} {3}:{4}:00', $StartYear, $StartMonth, $StartDay, $StartHour, $StartMinute)
                    $StartDateTimeObject = Get-Date -Date $StartDateTimeString
                }
            }
            # Convert to binary string and pad left with 0 to ensure 32 character length for consistent parsing
            $binaryRecurrence = [Convert]::ToString([int64]"0x$Recurrence".ToString(), 2).PadLeft(32, 48)
    
            [bool]$IsGMT = [Convert]::ToInt32($binaryRecurrence.Substring(31, 1), 2)
    
            <#
                Day duration is found by calculating how many times 24 goes into our TotalHourDuration (number of times being DayDuration)
                and getting the remainder for HourDuration by using % for modulus
            #>
            $TotalHourDuration = [Convert]::ToInt32($binaryRecurrence.Substring(0, 5), 2)
    
            switch ($TotalHourDuration -gt 24) {
                $true {
                    $Hours = $TotalHourDuration % 24
                    $DayDuration = ($TotalHourDuration - $Hours) / 24
                    $HourDuration = $Hours
                }
                $false {
                    $HourDuration = $TotalHourDuration
                    $DayDuration = 0
                }
            }
    
            $RecurType = [Convert]::ToInt32($binaryRecurrence.Substring(10, 3), 2)
    
            $MW['SmsProviderObjectPath'] = $TypeMap[$RecurType]
            $MW['DayDuration'] = $DayDuration
            $MW['HourDuration'] = $HourDuration
            $MW['MinuteDuration'] = $MinuteDuration
            $MW['IsGMT'] = $IsGMT
            $MW['StartTime'] = $StartDateTimeObject
    
            Switch ($RecurType) {
                1 {
                    $MW['Description'] = [string]::Format('Occurs on {0}', $StartDateTimeObject)
                }
                2 {
                    $MinuteSpan = [Convert]::ToInt32($binaryRecurrence.Substring(13, 6), 2)
                    $Hourspan = [Convert]::ToInt32($binaryRecurrence.Substring(19, 5), 2)
                    $DaySpan = [Convert]::ToInt32($binaryRecurrence.Substring(24, 5), 2)
                    if ($MinuteSpan -ne 0) {
                        $Span = 'minutes'
                        $Interval = $MinuteSpan
                    }
                    elseif ($HourSpan -ne 0) {
                        $Span = 'hours'
                        $Interval = $HourSpan
                    }
                    elseif ($DaySpan -ne 0) {
                        $Span = 'days'
                        $Interval = $DaySpan
                    }
                                
                    $MW['Description'] = [string]::Format('Occurs every {0} {1} effective {2}', $Interval, $Span, $StartDateTimeObject)
                    $MW['MinuteSpan'] = $MinuteSpan
                    $MW['HourSpan'] = $Hourspan
                    $MW['DaySpan'] = $DaySpan
                }
                3 {
                    $Day = [Convert]::ToInt32($binaryRecurrence.Substring(13, 3), 2)
                    $WeekRecurrence = [Convert]::ToInt32($binaryRecurrence.Substring(16, 3), 2)
                    $MW['Description'] = [string]::Format('Occurs every {0} weeks on {1} effective {2}', $WeekRecurrence, $([DayOfWeek]($Day - 1)), $StartDateTimeObject)
                    $MW['Day'] = $Day
                    $MW['ForNumberOfWeeks'] = $WeekRecurrence
                }
                4 {
                    $Day = [Convert]::ToInt32($binaryRecurrence.Substring(13, 3), 2)
                    $ForNumberOfMonths = [Convert]::ToInt32($binaryRecurrence.Substring(16, 4), 2)
                    $WeekOrder = [Convert]::ToInt32($binaryRecurrence.Substring(20, 3), 2)
                    $WeekRecurrence = switch ($WeekOrder) {
                        0 {
                            'Last'
                        }
                        default {
                            $(Get-FancyDay -Day $WeekOrder)
                        }
                    }
                    $MW['Description'] = [string]::Format('Occurs the {0} {1} of every {2} months effective {3}', $WeekRecurrence, $([DayOfWeek]($Day - 1)), $ForNumberOfMonths, $StartDateTimeObject)
                    $MW['Day'] = $Day
                    $MW['ForNumberOfMonths'] = $ForNumberOfMonths
                    $MW['WeekOrder'] = $WeekOrder
                }
                5 {
                    $MonthDay = [Convert]::ToInt32($binaryRecurrence.Substring(13, 5), 2)
                    $MonthRecurrence = switch ($MonthDay) {
                        0 {
                            # $Today = [datetime]::Today
                            # [datetime]::DaysInMonth($Today.Year, $Today.Month)
                            'the last day'
                        }
                        default {
                            "day $PSItem"
                        }
                    }
                    $ForNumberOfMonths = [Convert]::ToInt32($binaryRecurrence.Substring(18, 4), 2)
                    $MW['Description'] = [string]::Format('Occurs {0} of every {1} months effective {2}', $MonthRecurrence, $ForNumberOfMonths, $StartDateTimeObject)
                    $MW['ForNumberOfMonths'] = $ForNumberOfMonths
                    $MW['MonthDay'] = $MonthDay
                }
                Default {
                    Write-Error "Parsing Schedule String resulted in invalid type of $RecurType"
                }
            }
    
            [pscustomobject]$MW
        }
    }
}

function Get-EnumFromBitmap($enum = @("Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"), $bitmap = 0) {
	$result = [array]@()
	foreach ($bit in ($($enum.Count - 1)..0) ) {
		$pow = [math]::pow(2, $bit)
		if ($bitmap -ge $pow) {
			$bitmap -= $pow
			$result += $enum[$bit]
		}
	}
	[array]::Reverse($result)
	return ($result -join ",")
}

function Get-DayLabel($dayofweek,$first=0){
    $dayindex = $dayofweek - $first
    switch($dayindex){
        0{"Sunday"}
        1{"Monday"}
        2{"Tuesday"}
        3{"Wednesday"}
        4{"Friday"}
        5{"Saturday"}
        default{"Sunday"}
    }
}

#endregion

#region GetData functions

Function Execute-SQLQuery ($ServerInstance = $env:COMPUTERNAME,[ValidateNotNullOrEmpty()][string]$Query,$querytimeout=120) {
    $Datatable = New-Object System.Data.DataTable
    $Connection = New-Object System.Data.SQLClient.SQLConnection
    $Connection.ConnectionString = "SERVER='$ServerInstance';TRUSTED_CONNECTION=TRUE;Connection Timeout=$querytimeout;"
    $Connection.Open()
    $Command = New-Object System.Data.SQLClient.SQLCommand
    $Command.Connection = $Connection
    $Command.CommandText = $Query
    $Reader = $Command.ExecuteReader()
    $Datatable.Load($Reader)
    $Connection.Close()
    return , $Datatable
}
	
function Get-SiteMaintenanceTasks($CMDBName){
	write-host "Site Maintenance tasks" -NoNewline
    $sqlq=@"
    with TS as (
    select TaskName, max(laststarttime) laststarttime ,max(LastCompletionTime) LastCompletionTime
    from SQLTaskSiteStatus 
    group by TaskName
    )
    select distinct T.Sitecode, T.TaskName, T.IsEnabled 'Enabled', T.DaysOfWeek, T.BeginTime,T.LatestBeginTime,t.DeleteOlderThan,t.BackupLocation DeviceName,
	    TS.Laststarttime,TS.LastCompletionTime,DATEDIFF(SECOND,TS.laststarttime,TS.LastCompletionTime) DurationSeconds
    from [dbo].[vSMS_SC_SQL_Task] T left join TS on T.TaskName = TS.TaskName
"@
	$o = (Measure-Command{
		Write-Host ": Querying SQL..." -NoNewline
        #$allmt = invoke-sqlcmd -query "use $CMDBName;$sqlq" @sqlparms
        $allmt = Execute-SQLQuery -query "use $CMDBName;$sqlq" @sqlparms
	}).TotalMilliseconds
	Write-Host "done ($o ms)"
    return $allmt | Select-Object * -ExcludeProperty Table,RowError,RowState,ItemArray,HasErrors
}

function Get-SummaryTasks($CMDBName){
	write-host "Summary tasks" -NoNewline
    $sqlq = @"
    select TaskName, Enabled, LastRunDuration, LastRunResult, LastSuccessfulCompletionTime, NextStartTime, RunInterval, SiteTypes 'Sitetype', TaskCommand, TaskParameter from [dbo].[v_SummaryTasks] S
"@
	$o = (Measure-Command{
		Write-Host ": Querying SQL" -NoNewline
        #$allst = invoke-sqlcmd -query "use $CMDBName;$sqlq" @sqlparms
        $allst = Execute-SQLQuery -query "use $CMDBName;$sqlq" @sqlparms
	}).TotalMilliseconds
	Write-Host ": done ($o ms)"
return $allst | Select-Object * -ExcludeProperty Table,RowError,RowState,ItemArray,HasErrors
}

function Get-CollevalRefresh($CMDBName){
	$compset = @"
select S.Sitecode, CP.Name, Value3 
from SC_Component_Property CP left join SC_Component C on CP.ComponentID = C.ID 
left join sites S on C.SiteNumber = S.SiteKey
where ComponentName = 'SMS_COLLECTION_EVALUATOR' and CP.Name = 'Incremental Interval'
"@
	#$ColleValSettings = invoke-sqlcmd -query "use $CMDBName;$compset" @sqlparms
	$ColleValSettings = Execute-SQLQuery -query "use $CMDBName;$compset" @sqlparms
	return $ColleValSettings | Select-Object * -ExcludeProperty Table,RowError,RowState,ItemArray,HasErrors
}

function Get-CollEvalCosts($CMDBName){
	write-host "Collection evaluations" -NoNewline
		
	$sqlq = @"
		-- set statistics io on
declare @mintime int = 0 -- select collections taking more than this number in seconds
declare @MinDeployments int = 0 -- seelct collections with at least this number of active pkg deployments 
declare @MinChildren int = 0 -- collections cith at least this number of children
declare @MinCIDeploys int = 0 -- at least this number of active CI deployments (Apps and Updates)
declare @RefreshTypes table (Id int);
insert into @RefreshTypes
values (0),
--     (1),
    (2),
    (4),
    (6);
-- 0 Manual, 1 Manual, 2 Schedule, 4 Incremental, 6 both
-- select @mintime 'Taking more than (seconds)', @MinDeployments 'Min deployments',@MinChildren 'Min children', @MinCIDeploys 'Min CI Deploys';
WITH 
QuerySchedules as (
    select SiteID ScheduleSiteId,
        case
            A.FirstWord
            when '00012000' then 'Simple'
            else 'Custom'
        End ScheduleType,
        CASE
            ((SW2 & 3670016) / 524288)
            when 1 then 'NonRecurring'
            when 2 then 'RecurInterval'
            when 3 then 'RecurWeekly'
            when 4 then 'RecurMonthlyByWeekday'
            when 5 then 'RecurMonthlyByDate'
        END strFlags,
        -- Always relevant
        (
            FB21 + (2 * FB22) + (4 * FB23) + (8 * FB24) + (16 * FB25)
        ) StartHour,
        (
            FB26 + (2 * FB27) + (4 * FB28) + (8 * FB29) + (16 * FB30) + (32 * FB31)
        ) StartMinute,
        (
            FB16 + (2 * FB17) + (4 * FB18) + (8 * FB19) + (16 * FB20)
        ) StartDay,
        (FB12 + (2 * FB13) + (4 * FB14) + (8 * FB15)) StartMonth,
        (
            FB6 + (2 * FB7) + (4 * FB8) + (8 * FB9) + (16 * FB10) + (32 * FB11)
        ) + 1970 StartYear,
        (
            sB22 + (2 * sB23) + (4 * sB24) + (8 * sB25) + (16 * sB26)
        ) DurationDays,
        (
            sB27 + (2 * SB28) + (4 * SB29) + (8 * SB30) + (16 * SB31)
        ) DurationHours,
        (
            FB0 + (2 * FB1) + (4 * FB2) + (8 * FB3) + (16 * FB4) + (32 * FB5)
        ) DurationMinutes,
        SB0 UTC -- depending on strFlags
        -- recur_none
        -- recur_interval
,
        iif (
            (SW2 & 3670016) / 524288 = 2,
            (
                sB13 + (2 * sB14) + (4 * sB15) + (8 * sB16) + (16 * sB17) + (32 * sB18)
            ),
            NULL
        ) IntervalMinutes,
        iif (
            (SW2 & 3670016) / 524288 = 2,
            (
                sB8 + (2 * sB9) + (4 * sB10) + (8 * sB11) + (16 * sB12)
            ),
            NULL
        ) IntervalHours,
        iif (
            (SW2 & 3670016) / 524288 = 2,
            (
                sB3 + (2 * sB4) + (4 * sB5) + (8 * sB6) + (16 * sB7)
            ),
            NULL
        ) IntervalDays -- recur_weekly
,
        iif (
            (SW2 & 3670016) / 524288 = 3,
            (sB13 + (2 * sB14) + (4 * sB15)),
            NULL
        ) WeeklyWeeks,
        iif (
            (SW2 & 3670016) / 524288 = 3,
            (sB16 + (2 * sB17) + (4 * sB18)),
            NULL
        ) WeeklyDay -- recur_monthlybyweekday
,
        iif (
            (SW2 & 3670016) / 524288 = 4,
            (sB16 + (2 * sB17) + (4 * sB18)),
            NULL
        ) MonthlyWeekDay,
        iif (
            (SW2 & 3670016) / 524288 = 4,
            (sB12 + (2 * sB13) + (4 * sB14) + (8 * sB15)),
            NULL
        ) MonthlyWDMonths,
        iif (
            (SW2 & 3670016) / 524288 = 4,
            (sB9 + (2 * sB10) + (4 * sB11)),
            NULL
        ) MonthlyWeekOrder -- recur_monthlybydate
,
        iif (
            (SW2 & 3670016) / 524288 = 5,
            (
                sB14 + (2 * sB15) + (4 * sB16) + (8 * sB17) + (16 * sB18)
            ),
            NULL
        ) MonthlyDate,
        iif (
            (SW2 & 3670016) / 524288 = 5,
            (sB10 + (2 * sB11) + (4 * sB12) + (8 * sB13)),
            NULL
        ) MonthlyDMonths
    from (
            select CollectionName,
                SiteID,
                schedule,
                RefreshType,
                '0x' + Substring(schedule, 1, 8) FirstWord,
                '0x' + substring(schedule, 9, 8) SecondWord,
                convert(
                    bigint,
                    convert(varbinary, '0x' + Substring(schedule, 1, 8), 1)
                ) FW2,
                convert(
                    bigint,
                    convert(varbinary, '0x' + substring(schedule, 9, 8), 1)
                ) SW2,
                CASE
                    (
                        POWER(2, 0) & convert(
                            bigint,
                            convert(varbinary, '0x' + substring(schedule, 1, 8), 1)
                        )
                    )
                    WHEN POWER(2, 0) THEN 1
                    WHEN 0 THEN 0
                END FB0,
                CASE
                    (
                        POWER(2, 1) & convert(
                            bigint,
                            convert(varbinary, '0x' + substring(schedule, 1, 8), 1)
                        )
                    )
                    WHEN POWER(2, 1) THEN 1
                    WHEN 0 THEN 0
                END FB1,
                CASE
                    (
                        POWER(2, 2) & convert(
                            bigint,
                            convert(varbinary, '0x' + substring(schedule, 1, 8), 1)
                        )
                    )
                    WHEN POWER(2, 2) THEN 1
                    WHEN 0 THEN 0
                END FB2,
                CASE
                    (
                        POWER(2, 3) & convert(
                            bigint,
                            convert(varbinary, '0x' + substring(schedule, 1, 8), 1)
                        )
                    )
                    WHEN POWER(2, 3) THEN 1
                    WHEN 0 THEN 0
                END FB3,
                CASE
                    (
                        POWER(2, 4) & convert(
                            bigint,
                            convert(varbinary, '0x' + substring(schedule, 1, 8), 1)
                        )
                    )
                    WHEN POWER(2, 4) THEN 1
                    WHEN 0 THEN 0
                END FB4,
                CASE
                    (
                        POWER(2, 5) & convert(
                            bigint,
                            convert(varbinary, '0x' + substring(schedule, 1, 8), 1)
                        )
                    )
                    WHEN POWER(2, 5) THEN 1
                    WHEN 0 THEN 0
                END FB5,
                CASE
                    (
                        POWER(2, 6) & convert(
                            bigint,
                            convert(varbinary, '0x' + substring(schedule, 1, 8), 1)
                        )
                    )
                    WHEN POWER(2, 6) THEN 1
                    WHEN 0 THEN 0
                END FB6,
                CASE
                    (
                        POWER(2, 7) & convert(
                            bigint,
                            convert(varbinary, '0x' + substring(schedule, 1, 8), 1)
                        )
                    )
                    WHEN POWER(2, 7) THEN 1
                    WHEN 0 THEN 0
                END FB7,
                CASE
                    (
                        POWER(2, 8) & convert(
                            bigint,
                            convert(varbinary, '0x' + substring(schedule, 1, 8), 1)
                        )
                    )
                    WHEN POWER(2, 8) THEN 1
                    WHEN 0 THEN 0
                END FB8,
                CASE
                    (
                        POWER(2, 9) & convert(
                            bigint,
                            convert(varbinary, '0x' + substring(schedule, 1, 8), 1)
                        )
                    )
                    WHEN POWER(2, 9) THEN 1
                    WHEN 0 THEN 0
                END FB9,
                CASE
                    (
                        POWER(2, 10) & convert(
                            bigint,
                            convert(varbinary, '0x' + substring(schedule, 1, 8), 1)
                        )
                    )
                    WHEN POWER(2, 10) THEN 1
                    WHEN 0 THEN 0
                END FB10,
                CASE
                    (
                        POWER(2, 11) & convert(
                            bigint,
                            convert(varbinary, '0x' + substring(schedule, 1, 8), 1)
                        )
                    )
                    WHEN POWER(2, 11) THEN 1
                    WHEN 0 THEN 0
                END FB11,
                CASE
                    (
                        POWER(2, 12) & convert(
                            bigint,
                            convert(varbinary, '0x' + substring(schedule, 1, 8), 1)
                        )
                    )
                    WHEN POWER(2, 12) THEN 1
                    WHEN 0 THEN 0
                END FB12,
                CASE
                    (
                        POWER(2, 13) & convert(
                            bigint,
                            convert(varbinary, '0x' + substring(schedule, 1, 8), 1)
                        )
                    )
                    WHEN POWER(2, 13) THEN 1
                    WHEN 0 THEN 0
                END FB13,
                CASE
                    (
                        POWER(2, 14) & convert(
                            bigint,
                            convert(varbinary, '0x' + substring(schedule, 1, 8), 1)
                        )
                    )
                    WHEN POWER(2, 14) THEN 1
                    WHEN 0 THEN 0
                END FB14,
                CASE
                    (
                        POWER(2, 15) & convert(
                            bigint,
                            convert(varbinary, '0x' + substring(schedule, 1, 8), 1)
                        )
                    )
                    WHEN POWER(2, 15) THEN 1
                    WHEN 0 THEN 0
                END FB15,
                CASE
                    (
                        POWER(2, 16) & convert(
                            bigint,
                            convert(varbinary, '0x' + substring(schedule, 1, 8), 1)
                        )
                    )
                    WHEN POWER(2, 16) THEN 1
                    WHEN 0 THEN 0
                END FB16,
                CASE
                    (
                        POWER(2, 17) & convert(
                            bigint,
                            convert(varbinary, '0x' + substring(schedule, 1, 8), 1)
                        )
                    )
                    WHEN POWER(2, 17) THEN 1
                    WHEN 0 THEN 0
                END FB17,
                CASE
                    (
                        POWER(2, 18) & convert(
                            bigint,
                            convert(varbinary, '0x' + substring(schedule, 1, 8), 1)
                        )
                    )
                    WHEN POWER(2, 18) THEN 1
                    WHEN 0 THEN 0
                END FB18,
                CASE
                    (
                        POWER(2, 19) & convert(
                            bigint,
                            convert(varbinary, '0x' + substring(schedule, 1, 8), 1)
                        )
                    )
                    WHEN POWER(2, 19) THEN 1
                    WHEN 0 THEN 0
                END FB19,
                CASE
                    (
                        POWER(2, 20) & convert(
                            bigint,
                            convert(varbinary, '0x' + substring(schedule, 1, 8), 1)
                        )
                    )
                    WHEN POWER(2, 20) THEN 1
                    WHEN 0 THEN 0
                END FB20,
                CASE
                    (
                        POWER(2, 21) & convert(
                            bigint,
                            convert(varbinary, '0x' + substring(schedule, 1, 8), 1)
                        )
                    )
                    WHEN POWER(2, 21) THEN 1
                    WHEN 0 THEN 0
                END FB21,
                CASE
                    (
                        POWER(2, 22) & convert(
                            bigint,
                            convert(varbinary, '0x' + substring(schedule, 1, 8), 1)
                        )
                    )
                    WHEN POWER(2, 22) THEN 1
                    WHEN 0 THEN 0
                END FB22,
                CASE
                    (
                        POWER(2, 23) & convert(
                            bigint,
                            convert(varbinary, '0x' + substring(schedule, 1, 8), 1)
                        )
                    )
                    WHEN POWER(2, 23) THEN 1
                    WHEN 0 THEN 0
                END FB23,
                CASE
                    (
                        POWER(2, 24) & convert(
                            bigint,
                            convert(varbinary, '0x' + substring(schedule, 1, 8), 1)
                        )
                    )
                    WHEN POWER(2, 24) THEN 1
                    WHEN 0 THEN 0
                END FB24,
                CASE
                    (
                        POWER(2, 25) & convert(
                            bigint,
                            convert(varbinary, '0x' + substring(schedule, 1, 8), 1)
                        )
                    )
                    WHEN POWER(2, 25) THEN 1
                    WHEN 0 THEN 0
                END FB25,
                CASE
                    (
                        POWER(2, 26) & convert(
                            bigint,
                            convert(varbinary, '0x' + substring(schedule, 1, 8), 1)
                        )
                    )
                    WHEN POWER(2, 26) THEN 1
                    WHEN 0 THEN 0
                END FB26,
                CASE
                    (
                        POWER(2, 27) & convert(
                            bigint,
                            convert(varbinary, '0x' + substring(schedule, 1, 8), 1)
                        )
                    )
                    WHEN POWER(2, 27) THEN 1
                    WHEN 0 THEN 0
                END FB27,
                CASE
                    (
                        POWER(2, 28) & convert(
                            bigint,
                            convert(varbinary, '0x' + substring(schedule, 1, 8), 1)
                        )
                    )
                    WHEN POWER(2, 28) THEN 1
                    WHEN 0 THEN 0
                END FB28,
                CASE
                    (
                        POWER(2, 29) & convert(
                            bigint,
                            convert(varbinary, '0x' + substring(schedule, 1, 8), 1)
                        )
                    )
                    WHEN POWER(2, 29) THEN 1
                    WHEN 0 THEN 0
                END FB29,
                CASE
                    (
                        POWER(2, 30) & convert(
                            bigint,
                            convert(varbinary, '0x' + substring(schedule, 1, 8), 1)
                        )
                    )
                    WHEN POWER(2, 30) THEN 1
                    WHEN 0 THEN 0
                END FB30,
                CASE
                    (
                        POWER(2, 15) & convert(
                            bigint,
                            convert(varbinary, '0x' + substring(schedule, 1, 4), 1)
                        )
                    )
                    WHEN POWER(2, 15) THEN 1
                    WHEN 0 THEN 0
                END FB31,
                CASE
                    (
                        POWER(2, 0) & convert(
                            bigint,
                            convert(varbinary, '0x' + substring(schedule, 9, 8), 1)
                        )
                    )
                    WHEN POWER(2, 0) THEN 1
                    WHEN 0 THEN 0
                END SB0,
                CASE
                    (
                        POWER(2, 1) & convert(
                            bigint,
                            convert(varbinary, '0x' + substring(schedule, 9, 8), 1)
                        )
                    )
                    WHEN POWER(2, 1) THEN 1
                    WHEN 0 THEN 0
                END SB1,
                CASE
                    (
                        POWER(2, 2) & convert(
                            bigint,
                            convert(varbinary, '0x' + substring(schedule, 9, 8), 1)
                        )
                    )
                    WHEN POWER(2, 2) THEN 1
                    WHEN 0 THEN 0
                END SB2,
                CASE
                    (
                        POWER(2, 3) & convert(
                            bigint,
                            convert(varbinary, '0x' + substring(schedule, 9, 8), 1)
                        )
                    )
                    WHEN POWER(2, 3) THEN 1
                    WHEN 0 THEN 0
                END SB3,
                CASE
                    (
                        POWER(2, 4) & convert(
                            bigint,
                            convert(varbinary, '0x' + substring(schedule, 9, 8), 1)
                        )
                    )
                    WHEN POWER(2, 4) THEN 1
                    WHEN 0 THEN 0
                END SB4,
                CASE
                    (
                        POWER(2, 5) & convert(
                            bigint,
                            convert(varbinary, '0x' + substring(schedule, 9, 8), 1)
                        )
                    )
                    WHEN POWER(2, 5) THEN 1
                    WHEN 0 THEN 0
                END SB5,
                CASE
                    (
                        POWER(2, 6) & convert(
                            bigint,
                            convert(varbinary, '0x' + substring(schedule, 9, 8), 1)
                        )
                    )
                    WHEN POWER(2, 6) THEN 1
                    WHEN 0 THEN 0
                END SB6,
                CASE
                    (
                        POWER(2, 7) & convert(
                            bigint,
                            convert(varbinary, '0x' + substring(schedule, 9, 8), 1)
                        )
                    )
                    WHEN POWER(2, 7) THEN 1
                    WHEN 0 THEN 0
                END SB7,
                CASE
                    (
                        POWER(2, 8) & convert(
                            bigint,
                            convert(varbinary, '0x' + substring(schedule, 9, 8), 1)
                        )
                    )
                    WHEN POWER(2, 8) THEN 1
                    WHEN 0 THEN 0
                END SB8,
                CASE
                    (
                        POWER(2, 9) & convert(
                            bigint,
                            convert(varbinary, '0x' + substring(schedule, 9, 8), 1)
                        )
                    )
                    WHEN POWER(2, 9) THEN 1
                    WHEN 0 THEN 0
                END SB9,
                CASE
                    (
                        POWER(2, 10) & convert(
                            bigint,
                            convert(varbinary, '0x' + substring(schedule, 9, 8), 1)
                        )
                    )
                    WHEN POWER(2, 10) THEN 1
                    WHEN 0 THEN 0
                END SB10,
                CASE
                    (
                        POWER(2, 11) & convert(
                            bigint,
                            convert(varbinary, '0x' + substring(schedule, 9, 8), 1)
                        )
                    )
                    WHEN POWER(2, 11) THEN 1
                    WHEN 0 THEN 0
                END SB11,
                CASE
                    (
                        POWER(2, 12) & convert(
                            bigint,
                            convert(varbinary, '0x' + substring(schedule, 9, 8), 1)
                        )
                    )
                    WHEN POWER(2, 12) THEN 1
                    WHEN 0 THEN 0
                END SB12,
                CASE
                    (
                        POWER(2, 13) & convert(
                            bigint,
                            convert(varbinary, '0x' + substring(schedule, 9, 8), 1)
                        )
                    )
                    WHEN POWER(2, 13) THEN 1
                    WHEN 0 THEN 0
                END SB13,
                CASE
                    (
                        POWER(2, 14) & convert(
                            bigint,
                            convert(varbinary, '0x' + substring(schedule, 9, 8), 1)
                        )
                    )
                    WHEN POWER(2, 14) THEN 1
                    WHEN 0 THEN 0
                END SB14,
                CASE
                    (
                        POWER(2, 15) & convert(
                            bigint,
                            convert(varbinary, '0x' + substring(schedule, 9, 8), 1)
                        )
                    )
                    WHEN POWER(2, 15) THEN 1
                    WHEN 0 THEN 0
                END SB15,
                CASE
                    (
                        POWER(2, 16) & convert(
                            bigint,
                            convert(varbinary, '0x' + substring(schedule, 9, 8), 1)
                        )
                    )
                    WHEN POWER(2, 16) THEN 1
                    WHEN 0 THEN 0
                END SB16,
                CASE
                    (
                        POWER(2, 17) & convert(
                            bigint,
                            convert(varbinary, '0x' + substring(schedule, 9, 8), 1)
                        )
                    )
                    WHEN POWER(2, 17) THEN 1
                    WHEN 0 THEN 0
                END SB17,
                CASE
                    (
                        POWER(2, 18) & convert(
                            bigint,
                            convert(varbinary, '0x' + substring(schedule, 9, 8), 1)
                        )
                    )
                    WHEN POWER(2, 18) THEN 1
                    WHEN 0 THEN 0
                END SB18,
                CASE
                    (
                        POWER(2, 19) & convert(
                            bigint,
                            convert(varbinary, '0x' + substring(schedule, 9, 8), 1)
                        )
                    )
                    WHEN POWER(2, 19) THEN 1
                    WHEN 0 THEN 0
                END SB19,
                CASE
                    (
                        POWER(2, 20) & convert(
                            bigint,
                            convert(varbinary, '0x' + substring(schedule, 9, 8), 1)
                        )
                    )
                    WHEN POWER(2, 20) THEN 1
                    WHEN 0 THEN 0
                END SB20,
                CASE
                    (
                        POWER(2, 21) & convert(
                            bigint,
                            convert(varbinary, '0x' + substring(schedule, 9, 8), 1)
                        )
                    )
                    WHEN POWER(2, 21) THEN 1
                    WHEN 0 THEN 0
                END SB21,
                CASE
                    (
                        POWER(2, 22) & convert(
                            bigint,
                            convert(varbinary, '0x' + substring(schedule, 9, 8), 1)
                        )
                    )
                    WHEN POWER(2, 22) THEN 1
                    WHEN 0 THEN 0
                END SB22,
                CASE
                    (
                        POWER(2, 23) & convert(
                            bigint,
                            convert(varbinary, '0x' + substring(schedule, 9, 8), 1)
                        )
                    )
                    WHEN POWER(2, 23) THEN 1
                    WHEN 0 THEN 0
                END SB23,
                CASE
                    (
                        POWER(2, 24) & convert(
                            bigint,
                            convert(varbinary, '0x' + substring(schedule, 9, 8), 1)
                        )
                    )
                    WHEN POWER(2, 24) THEN 1
                    WHEN 0 THEN 0
                END SB24,
                CASE
                    (
                        POWER(2, 25) & convert(
                            bigint,
                            convert(varbinary, '0x' + substring(schedule, 9, 8), 1)
                        )
                    )
                    WHEN POWER(2, 25) THEN 1
                    WHEN 0 THEN 0
                END SB25,
                CASE
                    (
                        POWER(2, 26) & convert(
                            bigint,
                            convert(varbinary, '0x' + substring(schedule, 9, 8), 1)
                        )
                    )
                    WHEN POWER(2, 26) THEN 1
                    WHEN 0 THEN 0
                END SB26,
                CASE
                    (
                        POWER(2, 27) & convert(
                            bigint,
                            convert(varbinary, '0x' + substring(schedule, 9, 8), 1)
                        )
                    )
                    WHEN POWER(2, 27) THEN 1
                    WHEN 0 THEN 0
                END SB27,
                CASE
                    (
                        POWER(2, 28) & convert(
                            bigint,
                            convert(varbinary, '0x' + substring(schedule, 9, 8), 1)
                        )
                    )
                    WHEN POWER(2, 28) THEN 1
                    WHEN 0 THEN 0
                END SB28,
                CASE
                    (
                        POWER(2, 29) & convert(
                            bigint,
                            convert(varbinary, '0x' + substring(schedule, 9, 8), 1)
                        )
                    )
                    WHEN POWER(2, 29) THEN 1
                    WHEN 0 THEN 0
                END SB29,
                CASE
                    (
                        POWER(2, 30) & convert(
                            bigint,
                            convert(varbinary, '0x' + substring(schedule, 9, 8), 1)
                        )
                    )
                    WHEN POWER(2, 30) THEN 1
                    WHEN 0 THEN 0
                END SB30,
                CASE
                    (
                        POWER(2, 15) & convert(
                            bigint,
                            convert(varbinary, '0x' + substring(schedule, 9, 4), 1)
                        )
                    )
                    WHEN POWER(2, 15) THEN 1
                    WHEN 0 THEN 0
                END SB31 -- cant use power(2,31)		
            from v_Collections
        ) A
)
select *
from (
        Select distinct -- FH.Name 'Console Folder',
            c.CollectionName,
            SiteID,
            CASE
                c.RefreshType
                WHEN 1 THEN 'Manual (1)'
                WHEN 2 THEN 'Schedule (2)'
                WHEN 4 THEN 'Incremental (4)'
                WHEN 6 THEN 'Both (6)'
                Else 'Unknown (' + cast(c.RefreshType as varchar(max)) + ')'
            End as 'CollRefresh',
            IIF (
                (
                    c.RefreshType = 2
                    OR c.RefreshType = 6
                ),
                case
                    (QS.strFlags)
                    when 'RecurInterval' then CASE
                        when (QS.IntervalDays > 0) then 'Every ' + convert(varchar(2), QS.IntervalDays) + ' day(s)'
                        when (QS.IntervalHours > 0) then 'Every ' + convert(varchar(2), QS.IntervalHours) + ' hour(s)'
                        when (QS.IntervalMinutes > 0) then 'Every ' + convert(varchar(2), QS.IntervalMinutes) + ' minute(s)'
                        else ''
                    END
                    when 'RecurWeekly' then 'Every ' + + convert(varchar(2), QS.WeeklyWeeks) + ' week(s) on day ' + + convert(varchar(2), QS.WeeklyDay)
                    when 'RecurMonthlyByWeekDay' then ''
                    when 'RecurMonthlyByDate' then ''
                    else 'One shot'
                end,
                'No Full'
            ) Interval,
            right(('00' + convert(varchar(2), QS.StartDay)), 2) + '/' + right('00' +(convert(varchar(2), QS.StartMonth)), 2) + '/' + convert(varchar(4), QS.StartYear) StartDate,
            right(('00' + convert(varchar(2), QS.StartHour)), 2) + ':' + right(('00' + convert(varchar(2), QS.StartMinute)), 2) StartTime,
            c.LimitToCollectionName 'FatherName' -- Collection weight Area
,
            c.MemberCount,
            EvaluationLength / 100 'Eval (1/10 sec)' -- ,c.Schedule
,
            (
                select count(*)
                from [v_CIAssignmentTargetedCollections] ATC
                where ATC.CollectionID = C.SiteID
            ) CIAssignments,
            (
                select count(*)
                from [dbo].[DeploymentSummary] DS
                where DS.CollectionId = C.SiteID
            ) PkgDeployments -- Collection usefulness Area End
			,Sitecode
        from v_Collections c
            --left join QueryScores s on c.CollectionId = s.CollectionID -- left join Collection_EvaluationAndCRCData ccrc on c.CollectionId = ccrc.CollectionId
            INNER JOIN dbo.Collections_L AS t1 ON c.Collectionid = t1.CollectionID
            left join FolderMembers FM on FM.InstanceKey = c.SiteID -- left join FolderHierarchy FH on FH.ContainerNodeID = FM.ContainerNodeID
            -- left join v_Collection F on F.CollectionID = C.LimitToCollectionID
            left join QuerySchedules QS on c.SiteID = QS.ScheduleSiteID -- left join QuerySchedules FS on C.LimitToCollectionID = FS.ScheduleSiteID
			left join Sites S on S.SiteKey = SiteNumber
        where EvaluationLength >= @mintime --		and LimitToCollectionID not like 'SMS%' 
            --  and F.RefreshType in (select id from @FatherRefreshTypes)
            and c.RefreshType in (
                select id
                from @RefreshTypes
            )

    ) A
where StartDate <> '00/00/1970'

"@
		
	$oldloc = Get-Location
		
	$o = (Measure-Command{
		Write-Host ": Querying SQL" -NoNewline
		#$CollEvalQueryResults = invoke-sqlcmd -query "use $CMDBName;$sqlq" @sqlparms
		$CollEvalQueryResults = Execute-SQLQuery -query "use $CMDBName;$sqlq" @sqlparms
	}).TotalMilliseconds
	Write-Host ": done ($o ms)"
		
	Set-Location $oldloc
    return $CollEvalQueryResults | Select-Object * -ExcludeProperty Table,RowError,RowState,ItemArray,HasErrors
}

function Get-PlannedDeployments($CMDBName){
	write-host "Recurring advertisements" -NoNewline
	$sqlq = "SELECT 
	PKG.Name 'PkgName',
	ADV.CollectionID , 
	-- ADV.OfferName,
	ADV.MandatorySched, 
	ADV.PkgID,
	COL.Name 'CollectionName',
	COL.MemberCount 'CollMembers'
	FROM [dbo].[vSMS_Advertisement] ADV
	left join v_Collection COL on ADV.CollectionID = COL.CollectionID
	left join SMSPackages PKG on ADV.PkgID = PKG.PkgID
	where MandatorySched <> ''
	"
		
	$o = (Measure-Command{
		Write-Host ": Querying SQL" -NoNewline
		#$PlannedAdvs = invoke-sqlcmd -query "use $CMDBName;$sqlq" @sqlparms
		$PlannedAdvs = Execute-SQLQuery  -query "use $CMDBName;$sqlq" @sqlparms
	}).TotalMilliseconds
	Write-Host ": done ($o ms)"
    return $PlannedAdvs | Select-Object * -ExcludeProperty Table,RowError,RowState,ItemArray,HasErrors
}

function Get-ClientSchedules($CMDBName){
}

#endregion

#region TransformData functions

function Get-SiteMaintenanceTasksSchedule($allmt){
	$SMTSchedule = @{}
	$dayindex = 6
	foreach($d in @("Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday")){
		$temp=[PSCustomObject]@{
			DayOfWeek = ($dayindex+7)%7		
			DayName = "$d"
			Type = "MaintenanceTask"
		}
		foreach($t in 0..23){
			$temp | Add-Member NoteProperty $t @() # was ""
		}
		$dayindex++
		$SMTSchedule[$d]=$temp
	}
		
	foreach ($smt in $allmt | Where-Object{$_.Enabled -eq 1}) {
		$daymap = (Get-EnumFromBitmap -bitmap $smt.DaysOfWeek).Split(",")
		if($null -ne $smt.LatestBeginTime){                
			$LatestBeginTime  = ("0000" + $smt.LatestBeginTime).Substring("$($smt.LatestBeginTime)".Length,4)
		}
		else{
			$LatestBeginTime = $null
		}
		if($null -ne $smt.BeginTime){
            $BeginTime  = ("0000" + $smt.BeginTime).Substring("$($smt.BeginTime)".Length,4)
        } 
		else{
			$BeginTime = $null
		}
		foreach($d in @("Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday")){
			if($daymap.Contains($d) -and ($null -ne $BeginTime)){
				if(($null -ne $BeginTime) -and ($null -ne $LatestBeginTime) ){
					$BeginIndex = [int]"$($BeginTime.Substring(0,2))"
                    #$LatestIndex = $BeginIndex + [math]::Round($smt.durationseconds / 3600, 0)
                    $hourcredit = $smt.DurationSeconds
                    for($TimeIndex = $BeginIndex ; $hourcredit -gt 0 ; $TimeIndex++){
                        if($hourcredit -gt 0){
                            if($hourcredit -gt 3600){$span=" (spanning later)"}
                            else{$span=""}
                            $SMTSchedule["$d"].$TimeIndex += [PSCustomObject]@{
                                Name = $smt.TaskName + $span
                                Duration = [math]::min($hourcredit, 3600)
                                Sitecode = $smt.Sitecode
                                LastStart = $smt.Laststarttime.toshorttimestring()
                            }
                        }
                        $hourcredit -= [math]::min($hourcredit, 3600)
					}
				}
			}
		}
	}
    return $SMTSchedule
}

function Get-SummaryTasksSchedule($allst){
	$STSchedule = @{}
	$dayindex = 6
	foreach($d in @("Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday")){
		$temp=[PSCustomObject]@{
			DayOfWeek = ($dayindex+7)%7		
			DayName = "$d"
			Type = "SummaryTask"
		}
		foreach($t in 0..23){
			$temp | Add-Member NoteProperty $t @() # ""
		}
		$dayindex++
		$STSchedule[$d]=$temp
	}

	foreach($st in $allst){
		$int = switch($st.RunInterval){
			"300"   { "Every 5 mins"}
			"600"   { "Every 10 mins"}
			"1200"   { "Every 20 mins"}
			"3600"  { "Every hour" }
			"7200"  { "Every two hours" }
			"14400"  { "Every four hours" }
			"82800" { "Every 23h"}
			"86400" { "Daily"}
			"604800"{ "Every seven days"}
			"1209600" { "Every 14 days"}
			default {"Every $($st.RunInterval) seconds" }
		}
        if(($null -ne $st.NextStartTime) -and ($st.NextStartTime.Gettype().Name -ne "DBNull") -and ($st.Enabled)){
            if($st.runInterval -le 3600){
                $hourly = ($st.LastRunDuration * 3600 / $st.runInterval) 
                $toadd = 3600 
            }
            else{
                $hourly = $st.LastRunDuration 
                $toadd = $st.RunInterval
            }
            $next = $st.nextstarttime + $shift
            do {
                $d = $next.DayofWeek
                $h = $next.Hour
                $STSchedule["$d"].$h += [PSCustomObject]@{
                    Name = $st.TaskName + " (" + $st.LastRunDuration + "s " + $int + ")"
                    Duration = $hourly
                }
                $next=$next.AddSeconds($toadd)
            } until ( ($next - $st.nextstarttime).totalHours -gt (24 * 7))
        }
	}		
    return $STSchedule
}	

function Get-CollEvalSchedule($refresh, $CollEvalQueryResults){
	
    #region Prepare Colleval arrays
		
	$sched = @{}
	$schedList = @{}
	$dayindex = 6
	foreach($d in @("Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday")){
		$temp=[PSCustomObject]@{
			DayOfWeek = ($dayindex+7)%7		
			DayName = "$d"
			Type = "Schedule"
		}
		$tempFull=[PSCustomObject]@{
			DayOfWeek = ($dayindex+7)%7		
			DayName = "$d"
			Type = "Schedule"
		}
		foreach($t in 0..23){
			$temp | Add-Member NoteProperty $t 0.0
            $tempFull | Add-Member NoteProperty $t @()
		}
		$dayindex++
		$sched[$d]=$temp
		$schedList[$d]=$tempFull
	}
		
	$inc = @{}
	$incList = @{}
	$dayindex = 6
	foreach($d in @("Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday")){
		$temp=[PSCustomObject]@{
			DayOfWeek = ($dayindex+7)%7		
			DayName = "$d"
			Type = "Incremental (/$refresh mins)"
		}
		$tempFull=[PSCustomObject]@{
			DayOfWeek = ($dayindex+7)%7		
			DayName = "$d"
			Type = "Incremental"
		}
		foreach($t in 0..23){
			$temp | Add-Member NoteProperty $t  0.0
            $tempFull | Add-Member NoteProperty $t @()
		}
		$dayindex++
		$inc[$d]=$temp
		$incList[$d]=$tempFull
	}
		
	#endregion
		
	foreach($coll in $CollEvalQueryResults){

        $cleancoll = $coll # | Select-Object * -ExcludeProperty Table,RowError,RowState,ItemArray,HasErrors
        
        if($coll.collrefresh.trim() -eq "Incremental (4)" -or $coll.collrefresh.trim() -eq "Both (6)"){
			foreach($d in @("Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday")){
				foreach($t in 0..23){
					$inc[$d].$t = $inc[$d].$t + ( $cleancoll.'Eval (1/10 sec)' )
                    $incList[$d].$t+=$cleancoll
				}
			}
        }

        if($coll.collrefresh.trim() -eq "Schedule (2)" -or $coll.collrefresh.trim() -eq "Both (6)"){
            $cleancoll.CollectionName = $cleancoll.CollectionName + " ($($coll.Interval))" 

            $startDay = (get-date $coll.StartDate).DayOfWeek 
			$startTime = (get-date $coll.Starttime).Hour
					
			switch($coll.Interval){
				"Every 1 day(s)" {
					foreach($d in @("Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday")){
						$sched[$d].$startTime = $sched[$d].$startTime + ( $cleancoll.'Eval (1/10 sec)' )
						$schedList[$d].$startTime+=$cleancoll
					}
				}
				"Every 7 day(s)" {
					$sched["$startDay"].$startTime = $sched["$startDay"]."$startTime" + ($cleancoll.'Eval (1/10 sec)'  )   
					$schedList["$startDay"].$startTime+=$cleancoll
				}
						
				"Every 1 week(s) on day 6" {
					$sched["Saturday"].$startTime = $sched["Saturday"].$startTime + ( $cleancoll.'Eval (1/10 sec)'  )
					$schedList["Saturday"].$startTime+=$cleancoll
				}
						
				"Every 12 hour(s)" {
					foreach($d in @("Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday")){
						$sched[$d].$startTime = $sched[$d].$startTime + ( $cleancoll.'Eval (1/10 sec)' )
						$sched[$d].$(($startTime+12)%24) = $sched[$d].$(($startTime+12)%24) + ( $cleancoll.'Eval (1/10 sec)' )
						
                        $schedList[$d].$startTime += $cleancoll 
						$schedList[$d].$(($startTime+12)%24) += $cleancoll
					}
				}
				"Every 1 hour(s)" {
					foreach($d in @("Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday")){
						foreach($t in 0..23){
							$sched[$d].$t = $sched[$d].$t + ( $cleancoll.'Eval (1/10 sec)'  )
							$schedList[$d].$t += $cleancoll
						}
					}
				}
				"Every 15 minute(s)" {
					foreach($d in @("Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday")){
						foreach($t in 0..23){
							$sched[$d].$t = $sched[$d].$t + ( $cleancoll.'Eval (1/10 sec)'   * 4)
							$schedList[$d].$t +=$cleancoll
						}
					}
				}
				"Every 5 minute(s)" {
					foreach($d in @("Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday")){
						foreach($t in 0..23){
							$sched[$d].$t = $sched[$d].$t + ( $cleancoll.'Eval (1/10 sec)'   * 12)
							$schedList[$d].$t += $cleancoll
						}
					}
				}
				default{
					# write-host "missing schedule"
				}
						
			}
        }
    }

    return [PSCustomObject]@{
        Incremental = $inc
        Scheduled = $sched
        IncFull = $incList
        SchedFull = $schedList
    }

}

function Get-PlannedDeploymentSchedule($PlannedAdvs){
    # write-host "Build Advertisements schedule"
	$oldloc = Get-Location
	$NextAdv = @{}
	$dayindex = 6
	foreach($d in @("Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday")){
		$temp=[PSCustomObject]@{
			DayOfWeek = ($dayindex+7)%7		
			DayName = "$d"
			Type = "Deployments"
		}
		foreach($t in 0..23){
			$temp | Add-Member NoteProperty $t @() # ""
		}
		$dayindex++
		$nextadv[$d]=$temp
	}
	foreach($GlobalPlannedAdv in $PlannedAdvs){
        $error.Clear()
        # write-host $GlobalPlannedAdv.MandatorySched -ForegroundColor Yellow
        $sc = @(ConvertFrom-CCMSchedule -ScheduleString $GlobalPlannedAdv.MandatorySched )

        try{
            for($subPlanningindex= 1; $subPlanningindex -le ($GlobalPlannedAdv.MandatorySched.Length)/16; $subPlanningindex++){
                $PlannedAdv = $GlobalPlannedAdv
                if($subPlanningindex -gt 1){
                    $PlannedAdv.MandatorySched = $GlobalPlannedAdv.MandatorySched.substring( ($subPlanningindex-1)*16 , 16 )
                }
			    switch($sc.SmsProviderObjectPath){
				    "SMS_ST_NonRecurring" {
					    if(((($sc.StartTime) -(get-date)).Days -lt 7)  -and ($sc.StartTime -gt (get-date) )   ){
						    # $NextAdv["$($sc.StartTime.DayOfWeek)"].$($sc.StartTime.Hour) += "$($PlannedAdv.PkgName) -> $($PlannedAdv.CollectionName) ($($PlannedAdv.CollMembers))"
						    $NextAdv["$($sc.StartTime.DayOfWeek)"].$($sc.StartTime.Hour) += [PSCustomObject]@{
                                    Name = $PlannedAdv.PkgName
                                    TargetCollection = $PlannedAdv.CollectionName
                                    TargetCount =  $PlannedAdv.CollMembers
                            }
					    }
				    }
				    "SMS_ST_RecurInterval" {
                        if($sc.DaySpan -ne 0){
						    foreach($d in @("Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday")){
							    $NextAdv["$d"].$($sc.StartTime.Hour) += [PSCustomObject]@{
                                    Name = "$($PlannedAdv.PkgName) (daily)" 
                                    TargetCollection = $PlannedAdv.CollectionName
                                    TargetCount =  $PlannedAdv.CollMembers
                                }
						    }                
					    }
					    elseif($sc.HourSpan -ne 0){
						    foreach($d in @("Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday")){
							    foreach($t in 0..23){
								    if((($sc.StartTime.Hour + $t) % $sc.HourSpan) -eq 0){
									    $NextAdv["$d"].$($t) += [PSCustomObject]@{
                                            Name = "$($PlannedAdv.PkgName) (every $($sc.HourSpan)h)"
                                            TargetCollection = $PlannedAdv.CollectionName
                                            TargetCount =  $PlannedAdv.CollMembers
                                        }
								    }
							    }
						    }
					    }
					    elseif($sc.MinuteSpan -ne 0){
					    }
					    else{  
                            write-host "Unhandled Recurinterval for $($PlannedAdv.PkgName) to $($PlannedAdv.CollectionName)" -ForegroundColor Yellow
                            write-host "==>" $sc.description
					    }										
				    }
				    "SMS_ST_RecurMonthlyByDate" {
                        $md = $sc.MonthDay
                        $nm = $sc.ForNumberOfMonths
                        $thisday = (get-date).Day
                        $delta = 1 - $thisday
                            
                        if($md -eq 0){
                            $lastdayofmonth = (get-date).AddDays($delta).AddMonths(1).AddDays(-1).DayOfWeek
                            # write-host "$thisday $delta Last $lastdayofmonth"
                            $t = $sc.StartTime.Hour
                            $NextAdv["$lastdayofmonth"].$($t) += [PSCustomObject]@{
                                Name = "$($PlannedAdv.PkgName) (every last day of month)"
                                TargetCollection = $PlannedAdv.CollectionName
                                TargetCount =  $PlannedAdv.CollMembers
                            }                                
                        }

                        if($md -eq 1){
                            $firstdayofmonth = (get-date).AddDays($delta).DayOfWeek
                            # write-host "$thisday $delta First $firstdayofmonth"
                            $t = $sc.StartTime.Hour
                            $NextAdv["$firstdayofmonth"].$($t) += [PSCustomObject]@{
                                Name = "$($PlannedAdv.PkgName) (every 1st day of month)"
                                TargetCollection = $PlannedAdv.CollectionName
                                TargetCount =  $PlannedAdv.CollMembers
                            }                                
                        }
                    }
				    "SMS_ST_RecurWeekly" {
                        $d = Get-DayLabel -dayofweek $sc.Day -first 1
                        $t = $sc.StartTime.Hour
                        $NextAdv["$d"].$($t) += [PSCustomObject]@{
                            Name = "$($PlannedAdv.PkgName) (every $($sc.ForNumberOfWeeks) $d)"
                            TargetCollection = $PlannedAdv.CollectionName
                            TargetCount =  $PlannedAdv.CollMembers
                        }                                
                    }
				    default {
					    write-host "Unmanaged token for $($plannedadv.pkgname) : $($plannedadv.MandatorySched)....." -ForegroundColor Yellow
				    }
                }

    		}        
        }
	    catch{
            Write-Host "Error building deployments planning" -ForegroundColor Red            
            write-host $error[-1].Exception 
            write-host $error[-1].ErrorDetails
            write-host $error[-1].ScriptStackTrace
        }

    }
	Set-Location $oldloc
    return $NextAdv
}	

#endregion

#region Start

Set-StrictMode -Version latest

Clear-Host

$ThisScriptName = $MyInvocation.InvocationName
$allmtout = $ThisScriptName.Replace(".ps1","-allmt.json")
$allstout = $ThisScriptName.Replace(".ps1","-allst.json")
$setsname = $ThisScriptName.Replace(".ps1","-sets.json")
$qresname = $ThisScriptName.Replace(".ps1","-eval.json")
$advsname = $ThisScriptName.Replace(".ps1","-advs.json")
$DBName = $ThisScriptName.Replace(".ps1","-cmdb.json")

$local = Get-Date
$shift = $local - $local.ToUniversalTime() # used to adjust UTC times to local

#endregion

#region GetData

if((-not (Test-Path $allmtout)) -or $True ){ #collect and save
    write-host "No previously saved capture: running live mode"
    $sqlparms = @{
        ServerInstance="localhost"
        QueryTimeout=120
    }
    $m=(Measure-Command {
        $CMDBName="CM_CNG"
        $CMDBName = (Invoke-Sqlcmd "select name from sys.databases where name like 'C%[_]___'" @sqlparms).Name

        write-host "Loading from live..."
        $allmt = Get-SiteMaintenanceTasks -CMDBName $CMDBName
        $allst = Get-SummaryTasks -CMDBName $CMDBName
        $CollEvalSettings = get-CollevalRefresh -CMDBName $CMDBName
        $CollEvalQueryResults = Get-CollEvalCosts -CMDBName $CMDBName
        $PlannedAdvs = Get-PlannedDeployments -CMDBName $CMDBName
        write-host "Data collected"

        write-host "Saving..."
        $allmt | ConvertTo-Json | Out-File $allmtout
        $allst | ConvertTo-Json | Out-File $allstout
        $CollEvalSettings | ConvertTo-Json | Out-File $setsname
        $CollEvalQueryResults | ConvertTo-Json | Out-File $qresname
        $PlannedAdvs | ConvertTo-Json | Out-File $advsname
        $CMDBName | ConvertTo-Json | Out-File $DBName
     }).TotalMilliseconds
    write-host "Done ($m ms)."
} 

if(((Test-Path $allmtout)) ){ 
    write-host "Using previously saved captures: Loading" -NoNewline
    $m=(Measure-Command{
    $allmt = (Get-content $allmtout | ConvertFrom-Json) 
    write-host "." -NoNewline
    $allst = (Get-content $allstout | ConvertFrom-Json)
    write-host "." -NoNewline
    $ColleValSettings = (Get-content $setsname | ConvertFrom-Json)
    write-host "." -NoNewline
    $CollEvalQueryResults = (Get-content $qresname | ConvertFrom-Json)
    write-host "." -NoNewline
    $PlannedAdvs = (Get-content $advsname | ConvertFrom-Json)
    write-host "." -NoNewline
    $CMDBName = (Get-content $dbname | ConvertFrom-Json)
    }).TotalMilliseconds
    write-host "Done ($m ms)."
}
else{
    write-host "Could neither collect nor use previously saved files. Quitting."
    break
}

#endregion

#region TransformData
# seems silly to reload collected data, but it ensures that offline &online are similar


write-host "Transforming data" -NoNewline
$STSchedule = Get-SummaryTasksSchedule -allst $allst
write-host "." -NoNewline
$SMTSchedule = Get-SiteMaintenanceTasksSchedule -allmt $allmt
write-host "." -NoNewline
$refresh = ($ColleValSettings | Measure-Object -Property Value3 -Average).Average
write-host "." -NoNewline
$AllScheduledEvals = Get-CollEvalSchedule -refresh $refresh -CollEvalQueryResults $CollEvalQueryResults
write-host "." -NoNewline
$sitecodes = @($ColleValSettings.Sitecode)
write-host "." -NoNewline
$NextAdv = Get-PlannedDeploymentSchedule -PlannedAdvs $PlannedAdvs		
write-host "." -NoNewline
$inc = $AllScheduledEvals.Incremental
write-host "." -NoNewline
$sched = $AllScheduledEvals.Scheduled
write-host "." -NoNewline
$masterview = $STSchedule.Values + $SMTSchedule.Values + $sched.Values + $inc.Values + $NextAdv.Values 
write-host "."
write-host "Done."	

#endregion

#region XAML

$startupVars = (Get-Variable | % {$_.Name}),"_","PSItem","Error","args","?","^"

#region Prerequisites


<#
    after paste from Blend:
        remove x:Class    
        ensure Window Title = "Summary of scheduled MEMCM activies ($CMDBName)" 
        ensure Label for incremental collevals contains (/$refresh mins on $($sitecodes.Count) in 

#>

[xml]$xaml=@"
<Window 
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    xmlns:d="http://schemas.microsoft.com/expression/blend/2008"
    xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
        Title="Summary of scheduled MEMCM activies ($CMDBName)" Width="1280" Height="772">
    <Window.Resources>
        <Style TargetType="TextBox">
            <Setter Property="HorizontalAlignment" Value="Stretch"/>
            <Setter Property="VerticalAlignment" Value="Top"/>
            <Setter Property="TextWrapping" Value="Wrap"/>
            <Setter Property="Width" Value="auto"/>
        </Style>
        <Style TargetType="ProgressBar">
            <Setter Property="Orientation" Value="Vertical"/>
            <Setter Property="Width" Value="15" />
            <Setter Property="Height" Value="25"/>
        </Style>
        <Style TargetType="Label">
            <Setter Property="HorizontalAlignment" Value="Left" />
            <Setter Property="VerticalAlignment" Value="Top"/>
        </Style>
        <Style x:Key="Zero" TargetType="ProgressBar" BasedOn="{StaticResource {x:Type ProgressBar}}">
            <Setter Property="Foreground" Value="#FF75C129" />
        </Style>
        <Style x:Key="Un"  TargetType="ProgressBar" BasedOn="{StaticResource {x:Type ProgressBar}}">
            <Setter Property="Foreground" Value="#FF219E9E"/>
        </Style>
        <Style x:Key="Deux"  TargetType="ProgressBar" BasedOn="{StaticResource {x:Type ProgressBar}}">
            <Setter Property="Foreground" Value="#FFD101FF"/>
        </Style>
        <Style x:Key="Trois"  TargetType="ProgressBar" BasedOn="{StaticResource {x:Type ProgressBar}}">
            <Setter Property="Foreground" Value="#FFFFDC01"/>
        </Style>
        <Style x:Key="Quatre" TargetType="ProgressBar" BasedOn="{StaticResource {x:Type ProgressBar}}">
            <Setter Property="Foreground" Value="#FF0152FF"/>
        </Style>
        <Style TargetType="DataGridCell">
            <Setter Property="BorderBrush" Value="Aquamarine"></Setter>
            <Style.Triggers>
                <Trigger Property="IsSelected" Value="True">
                    <Setter Property="BorderBrush" Value="Red"></Setter>
                    <Setter Property="BorderThickness" Value="2"></Setter>
                </Trigger>
            </Style.Triggers>
        </Style>

    </Window.Resources>
    <DockPanel VerticalAlignment="Stretch" HorizontalAlignment="Stretch" >
        <DataGrid Name="myDatagrid" AutoGenerateColumns="False" 
HorizontalAlignment="Stretch" VerticalAlignment="Stretch"
Width="Auto" Height="Auto" 
VerticalScrollBarVisibility="Auto"
Margin="10,10,10,10" 
Background="LightCyan" 
SelectionUnit="Cell">
            <DataGrid.Columns>
                <DataGridTemplateColumn Header="Start">
                    <DataGridTemplateColumn.CellTemplate>
                        <DataTemplate>
                            <Grid Name ="Start" >
                                <TextBlock Width="46" Height="19" Name="Time" Text="{Binding Path=Time}"></TextBlock>
                            </Grid>
                        </DataTemplate>
                    </DataGridTemplateColumn.CellTemplate>
                </DataGridTemplateColumn>
                <DataGridTemplateColumn Header="Monday&#x0a;M Su In Sc D">
                    <DataGridTemplateColumn.CellTemplate>
                        <DataTemplate>
                            <Grid >
                                <StackPanel Orientation="Horizontal" >
                                    <ProgressBar Style="{StaticResource Zero}"   Value="{Binding M[0]}"></ProgressBar>
                                    <ProgressBar Style="{StaticResource Un}"     Value="{Binding M[1]}"></ProgressBar>
                                    <ProgressBar Style="{StaticResource Deux}"   Value="{Binding M[2]}"></ProgressBar>
                                    <ProgressBar Style="{StaticResource Trois}"  Value="{Binding M[3]}"></ProgressBar>
                                    <ProgressBar Style="{StaticResource Quatre}" Value="{Binding M[4]}"></ProgressBar>
                                </StackPanel>
                            </Grid>
                        </DataTemplate>
                    </DataGridTemplateColumn.CellTemplate>
                </DataGridTemplateColumn>
                <DataGridTemplateColumn Header="Tuesday&#x0a;M Su In Sc D">
                    <DataGridTemplateColumn.CellTemplate>
                        <DataTemplate>
                            <Grid>
                                <StackPanel Orientation="Horizontal" >
                                    <ProgressBar Style="{StaticResource Zero}"  Value="{Binding Tu[0]}"></ProgressBar>
                                    <ProgressBar Style="{StaticResource Un}"    Value="{Binding Tu[1]}"></ProgressBar>
                                    <ProgressBar Style="{StaticResource Deux}"  Value="{Binding Tu[2]}"></ProgressBar>
                                    <ProgressBar Style="{StaticResource Trois}" Value="{Binding Tu[3]}"></ProgressBar>
                                    <ProgressBar Style="{StaticResource Quatre}" Value="{Binding Tu[4]}"></ProgressBar>
                                </StackPanel>
                            </Grid>
                        </DataTemplate>
                    </DataGridTemplateColumn.CellTemplate>
                </DataGridTemplateColumn>
                <DataGridTemplateColumn Header="Wednesday&#x0a;M Su In Sc D">
                    <DataGridTemplateColumn.CellTemplate>
                        <DataTemplate>
                            <Grid>
                                <StackPanel Orientation="Horizontal" >
                                    <ProgressBar Style="{StaticResource Zero}"  Value="{Binding W[0]}"></ProgressBar>
                                    <ProgressBar Style="{StaticResource Un}"    Value="{Binding W[1]}"></ProgressBar>
                                    <ProgressBar Style="{StaticResource Deux}"  Value="{Binding W[2]}"></ProgressBar>
                                    <ProgressBar Style="{StaticResource Trois}" Value="{Binding W[3]}"></ProgressBar>
                                    <ProgressBar Style="{StaticResource Quatre}" Value="{Binding W[4]}"></ProgressBar>
                                </StackPanel>
                            </Grid>
                        </DataTemplate>
                    </DataGridTemplateColumn.CellTemplate>
                </DataGridTemplateColumn>
                <DataGridTemplateColumn Header="Thursday&#x0a;M Su In Sc D">
                    <DataGridTemplateColumn.CellTemplate>
                        <DataTemplate>
                            <Grid>
                                <StackPanel Orientation="Horizontal" >
                                    <ProgressBar Style="{StaticResource Zero}"  Value="{Binding Th[0]}"></ProgressBar>
                                    <ProgressBar Style="{StaticResource Un}"    Value="{Binding Th[1]}"></ProgressBar>
                                    <ProgressBar Style="{StaticResource Deux}"  Value="{Binding Th[2]}"></ProgressBar>
                                    <ProgressBar Style="{StaticResource Trois}" Value="{Binding Th[3]}"></ProgressBar>
                                    <ProgressBar Style="{StaticResource Quatre}" Value="{Binding Th[4]}"></ProgressBar>
                                </StackPanel>
                            </Grid>
                        </DataTemplate>
                    </DataGridTemplateColumn.CellTemplate>
                </DataGridTemplateColumn>
                <DataGridTemplateColumn Header="Friday&#x0a;M Su In Sc D">
                    <DataGridTemplateColumn.CellTemplate>
                        <DataTemplate>
                            <Grid>
                                <StackPanel Orientation="Horizontal" >
                                    <ProgressBar Style="{StaticResource Zero}"  Value="{Binding F[0]}"></ProgressBar>
                                    <ProgressBar Style="{StaticResource Un}"    Value="{Binding F[1]}"></ProgressBar>
                                    <ProgressBar Style="{StaticResource Deux}"  Value="{Binding F[2]}"></ProgressBar>
                                    <ProgressBar Style="{StaticResource Trois}" Value="{Binding F[3]}"></ProgressBar>
                                    <ProgressBar Style="{StaticResource Quatre}" Value="{Binding F[4]}"></ProgressBar>
                                </StackPanel>
                            </Grid>
                        </DataTemplate>
                    </DataGridTemplateColumn.CellTemplate>
                </DataGridTemplateColumn>
                <DataGridTemplateColumn Header="Saturday&#x0a;M Su In Sc D">
                    <DataGridTemplateColumn.CellTemplate>
                        <DataTemplate>
                            <Grid>
                                <StackPanel Orientation="Horizontal" >
                                    <ProgressBar Style="{StaticResource Zero}"  Value="{Binding Sa[0]}"></ProgressBar>
                                    <ProgressBar Style="{StaticResource Un}"    Value="{Binding Sa[1]}"></ProgressBar>
                                    <ProgressBar Style="{StaticResource Deux}"  Value="{Binding Sa[2]}"></ProgressBar>
                                    <ProgressBar Style="{StaticResource Trois}" Value="{Binding Sa[3]}"></ProgressBar>
                                    <ProgressBar Style="{StaticResource Quatre}" Value="{Binding Sa[4]}"></ProgressBar>
                                </StackPanel>
                            </Grid>
                        </DataTemplate>
                    </DataGridTemplateColumn.CellTemplate>
                </DataGridTemplateColumn>
                <DataGridTemplateColumn Header="Sunday&#x0a;M Su In Sc D">
                    <DataGridTemplateColumn.CellTemplate>
                        <DataTemplate>
                            <Grid>
                                <StackPanel Orientation="Horizontal" >
                                    <ProgressBar Style="{StaticResource Zero}"  Value="{Binding Su[0]}"></ProgressBar>
                                    <ProgressBar Style="{StaticResource Un}"    Value="{Binding Su[1]}"></ProgressBar>
                                    <ProgressBar Style="{StaticResource Deux}"  Value="{Binding Su[2]}"></ProgressBar>
                                    <ProgressBar Style="{StaticResource Trois}" Value="{Binding Su[3]}"></ProgressBar>
                                    <ProgressBar Style="{StaticResource Quatre}" Value="{Binding Su[4]}"></ProgressBar>
                                </StackPanel>
                            </Grid>
                        </DataTemplate>
                    </DataGridTemplateColumn.CellTemplate>
                </DataGridTemplateColumn>
            </DataGrid.Columns>
        </DataGrid>
        <ListView HorizontalAlignment="Stretch" Margin="10,10,10,10">
            <Label Name="SlotDetail" Margin="0,0,0,0" FontSize="16" FontWeight="Bold" />
            <Expander Header="Overview" IsExpanded="True">
                <Grid>
                    <Grid.ColumnDefinitions>
                        <ColumnDefinition Width="auto"/>
                        <ColumnDefinition Width="auto"/>
                        <ColumnDefinition Width="auto"/>
                        <ColumnDefinition Width="auto"/>
                        <ColumnDefinition Width="auto"/>
                    </Grid.ColumnDefinitions>
                    <Grid.RowDefinitions>
                        <RowDefinition Height="auto"/>
                        <RowDefinition Height="auto"/>
                    </Grid.RowDefinitions>
                    <Label Grid.Column="2" Grid.Row="0" Content="ColleEvals (Incremental) &#x0a;Every $refresh mins on $($sitecodes.Count) site(s))"/>
                    <Label Grid.Column="2" Grid.Row="1" Name ="lblInc" Content="0" FontSize="32" HorizontalAlignment="Center"/>

                    <Label Grid.Column="3" Grid.Row="0" Content="CollEvals &#x0a;(Scheduled)" />
                    <Label Grid.Column="3" Grid.Row="1" Name ="lblSched" Content="0" FontSize="32" HorizontalAlignment="Center"/>

                    <Label Grid.Column="0" Grid.Row="0" Content="Maintenance Tasks &#x0a;(seconds)" />
                    <Label Grid.Column="0" Grid.Row="1" Name="lblMT" Content="0" FontSize="32" HorizontalAlignment="Center"/>

                    <Label Grid.Column="1" Grid.Row="0" Content="Summary Tasks&#x0a;(seconds)" />
                    <Label Grid.Column="1" Grid.Row="3" Name="lblST"  Content="0" FontSize="32" HorizontalAlignment="Center"/>

                    <Label Grid.Column="4" Grid.Row="0" Content="Deployments&#x0a;(sum of targets)" />
                    <Label Grid.Column="4" Grid.Row="1" Name="lblDepl"  Content="0" FontSize="32" HorizontalAlignment="Center"/>

                </Grid>
            </Expander>
            <Expander Header="Maintenance Tasks">
                
            <Grid Height ="250" Margin="10,0,10,0" HorizontalAlignment="Stretch">
                <DataGrid  Name="dgMaintenanceTasks" Margin="10,10,10,10" AutoGenerateColumns="False"  HorizontalAlignment="Stretch">
                    <DataGrid.Columns>
                        <DataGridTextColumn Header="MaintenanceTask Name" Binding="{Binding Path=Name}" />
                        <DataGridTextColumn Header="Duration" Binding="{Binding Path=Duration}"/>
                        <DataGridTextColumn Header="Sitecode" Binding="{Binding Path=Sitecode}"/>
                    </DataGrid.Columns>
                </DataGrid>

            </Grid>
            </Expander>
            <Expander Header="Summary Tasks">
            <Grid Height ="250" Margin="10,0,10,0" HorizontalAlignment="Stretch">
                <DataGrid Name="dgSummaryTasks" Margin="10,10,10,10" AutoGenerateColumns="False" VerticalAlignment="Stretch" HorizontalAlignment="Stretch">
                    <DataGrid.Columns>
                        <DataGridTextColumn Header="SummaryTask Name" Binding="{Binding Path=Name}" />
                        <DataGridTextColumn Header="Duration (Hourly sum)" Binding="{Binding Path=Duration}"/>
                    </DataGrid.Columns>
                </DataGrid>
            </Grid>
            </Expander>
            <Expander Header="Incremental Collection evaluations">
                <Grid Height ="250" Margin="10,0,10,0" HorizontalAlignment="Stretch">
                <DataGrid  Name="dgIncremental" Margin="10,10,10,10" AutoGenerateColumns="False"  HorizontalAlignment="Stretch">
                    <DataGrid.Columns>
                        <DataGridTextColumn Header="Incremental Eval CollName" Binding="{Binding Path=CollectionName}" />
                        <DataGridTextColumn Header="RefreshType" Binding="{Binding Path=CollRefresh}"/>
                        <DataGridTextColumn Header="MemberCount" Binding="{Binding Path=MemberCount}"/>
                        <DataGridTextColumn Header="SiteCode" Binding="{Binding Path=SiteCode}"/>
                        <DataGridTextColumn Header="Eval (1/10 s)" Binding="{Binding Path='Eval (1/10 sec)'}"/>
                        <DataGridTextColumn Header="Deployments" Binding="{Binding Path=PkgDeployments}"/>
                        <DataGridTextColumn Header="CI Assignments" Binding="{Binding Path=CIAssignments}"/>
                    </DataGrid.Columns>
                </DataGrid>

            </Grid>
            </Expander>
            <Expander Header="Scheduled Collection evaluations">
                <Grid Height ="250" Margin="10,0,10,0" HorizontalAlignment="Stretch">
                <DataGrid  Name="dgScheduled" Margin="10,10,10,10" AutoGenerateColumns="False"  HorizontalAlignment="Stretch">
                    <DataGrid.Columns>
                        <DataGridTextColumn Header="Scheduled Eval CollName" Binding="{Binding Path=CollectionName}" />
                        <DataGridTextColumn Header="RefreshType" Binding="{Binding Path=CollRefresh}"/>
                        <DataGridTextColumn Header="MemberCount" Binding="{Binding Path=MemberCount}"/>
                        <DataGridTextColumn Header="SiteCode" Binding="{Binding Path=SiteCode}"/>
                        <DataGridTextColumn Header="Eval (1/10 s)" Binding="{Binding Path='Eval (1/10 sec)'}"/>
                        <DataGridTextColumn Header="Deployments" Binding="{Binding Path=PkgDeployments}"/>
                        <DataGridTextColumn Header="CI Assignments" Binding="{Binding Path=CIAssignments}"/>
                    </DataGrid.Columns>
                </DataGrid>

            </Grid>
            </Expander>
            <Expander Header="Deployments">
                <Grid Height ="250" Margin="10,0,10,0" HorizontalAlignment="Stretch">
                <DataGrid Name="dgDeployments" Margin="10,10,10,10" AutoGenerateColumns="False" VerticalAlignment="Stretch" HorizontalAlignment="Stretch">
                    <DataGrid.Columns>
                        <DataGridTextColumn Header="PackageName" Binding="{Binding Path=Name}" />
                        <DataGridTextColumn Header="TargetCollection" Binding="{Binding Path=TargetCollection}"/>
                        <DataGridTextColumn Header="TargetCount" Binding="{Binding Path=TargetCount}"/>
                    </DataGrid.Columns>
                </DataGrid>
            </Grid>
            </Expander>


        </ListView>
    </DockPanel>
</Window>

"@

Add-Type -AssemblyName PresentationFramework, System.Drawing, System.Windows.Forms
$netpath = [System.Reflection.Assembly]::GetExecutingAssembly().ImageRuntimeVersion
$wpforms = "$env:SystemRoot\Microsoft.NET\Framework\$netpath\WPF\WindowsFormsIntegration.dll"
[System.Reflection.Assembly]::LoadFile($wpforms) | Out-Null
$xamlreader = (New-Object System.Xml.XmlNodeReader $xaml)
$Window=[Windows.Markup.XamlReader]::Load($xamlreader)
$xaml.SelectNodes("//*[@Name]") | ForEach-Object { Set-Variable -Name ($_.Name) -Value $Window.FindName($_.Name) }

#endregion

#region MasterScheduleView (left pane)

$PlanningForDatagrid = @()
Add-Type @"
public class Slot
{
    public string Time = "";
    public double[] M = new double[5];
    public double[] Tu = new double[5];
    public double[] W = new double[5];
    public double[] Th = new double[5];
    public double[] F = new double[5];
    public double[] Sa = new double[5];
    public double[] Su = new double[5];
    public Slot(int RequiredSlot){
        Time = RequiredSlot + "h";
    }
}
"@
foreach($t in 0..23){
	$item = New-Object Slot $t
	$PlanningForDatagrid+=$item 
}
$daymap=@{
	"Monday"="M"
	"Tuesday"="Tu"
	"Wednesday"="W"
	"Thursday"="Th"
	"Friday"="F"
	"Saturday"="Sa"
	"Sunday"="Su"
}

$AllSummaryDuration = ($allst | Where-Object {"$($_.lastrunduration)" -ne ""} | where-object {$_.LastRunDuration -ge 0 } | Measure-Object -Property LastRunDuration -sum).Sum
$AllMaintenanceDuration = ($allmt | Where-Object {$_.DurationSeconds -ge 0 } | Measure-Object -Property DurationSeconds -sum).Sum
$AllDeploymentTargets = 0 + (($CollEvalQueryResults | ?{$_.siteid -like "SMS00001"}) | Measure-Object -Property Membercount -Maximum).Maximum

foreach($activityday in $masterview){
	$zeday = $daymap[$activityday.Dayname]
	foreach ($zehour in 0..23) {
		switch ($activityday.Type) {
			"Schedule" {  
				$PlanningForDatagrid[$zehour]."$($zeday)"[3] = (100.0 * ($activityday.$zehour / 3600))
			}

			"Incremental (/$refresh mins)" {  
				$PlanningForDatagrid[$zehour]."$($zeday)"[2] = [math]::min((100.0 * ($activityday.$zehour / ($refresh * 60))) , 100.0)
			}

			"MaintenanceTask" {
    		    if($activityday.$zehour.Count -gt 0){
                    try{
                        $PlanningForDatagrid[$zehour]."$($zeday)"[0] = [math]::min((100.0 * (($activityday.$zehour | Where-Object {$_.Duration -gt 0 }| Measure-Object -Property Duration -sum).Sum) / 3600),100.0)
                    }
                    catch{
                        write-Host "MT $zehour $zeday"
                        read-host "Error building PlanningForGrid"
                    }
			    }
                else{
                    $PlanningForDatagrid[$zehour]."$($zeday)"[0]=0
                }
            }

			"SummaryTask" {
    		    if($activityday.$zehour.Count -gt 0){
    				$PlanningForDatagrid[$zehour]."$($zeday)"[1] = [math]::min((100.0 * (($activityday.$zehour | Measure-Object -Property Duration -sum).Sum) / $AllSummaryDuration),100.0)
			    }
                else{
                    $PlanningForDatagrid[$zehour]."$($zeday)"[1]=0
                }
            }

			"Deployments" {
    		    if($activityday.$zehour.Count -gt 0){
    				$PlanningForDatagrid[$zehour]."$($zeday)"[4] = [math]::min((100.0 * (($activityday.$zehour | Measure-Object -Property TargetCount -sum).Sum) / $AllDeploymentTargets),100.0)
			    }
                else{
                    $PlanningForDatagrid[$zehour]."$($zeday)"[4]=0
                }
            }

			Default {
                # should not happen
            }
		}
	}
}

$myDatagrid.ItemsSource = $PlanningForDatagrid

#endregion

#region DetailView (right pane)

$detailDeployments=@()
$detailMaintenanceTasks=@()
$detailSummaryTasks=@()
$detailDeployments=@()
$detailScheduled=@()
$detailIncrementals=@()

$myDatagrid.Add_SelectedCellsChanged({
	$col = ([System.Windows.Controls.DataGrid]$args[0]).CurrentCell.Column.DisplayIndex
	$row = ([System.Windows.Controls.DataGrid]$args[0]).CurrentItem.Time.Replace("h","")

    $detailDeployments=@()
    $detailMaintenanceTasks=@()
    $detailSummaryTasks=@()
    $detailDeployments=@()
    $detailScheduled=@()
    $detailIncrementals=@()

    $sumST=0
    $sumMt=0
    $sumDepl=0
    $sumInc = 0
    $sumSched =0

    if($col -gt 0){
        $Slotdetail.Content = "$(Get-DayLabel -dayofweek $col) at $($row)h" 
        $show = $masterview | Where-Object{$_.DayOfWeek -eq ($col - 1)} | Sort-Object -Property Type
        $sumInc = [math]::Round($show[1].$row, 1)
        $sumSched = [math]::Round($show[3].$row,1)

        $show[2].$row | %{
            $detailMaintenanceTasks+=[PSCustomObject]@{
                Name = $_.Name
                Duration = $_.Duration
                SiteCode  = $_.Sitecode
            }
            if($_.Duration -gt 0){
                $sumMt+=$_.Duration
            }
        }
        
        $show[0].$row | %{
            $detailDeployments+=[PSCustomObject]@{
                Name = $_.Name
                TargetCollection = $_.TargetCollection
                TargetCount = $_.TargetCount
            }
            $sumDepl+=$_.TargetCount
        }

        $show[4].$row | %{
            $detailSummaryTasks+=[PSCustomObject]@{
                Name = $_.Name
                Duration = $_.Duration
            }
            if($_.duration -gt 0){
                $sumSt += $_.Duration
            }
        }
    
        $detailIncrementals=$AllScheduledEvals.IncFull["$(Get-DayLabel -dayofweek $col)"].$row
        $detailScheduled=$AllScheduledEvals.SchedFull["$(Get-DayLabel -dayofweek $col)"].$row

        $sumInc = [math]::Round($show[1].$row * 0.1, 1) 
        $sumSched = [math]::Round($show[3].$row * 0.1,1) 

    }
    else{        
        $Slotdetail.Content = "Select day and time" 
    }
    $dgMaintenanceTasks.ItemsSource = $detailMaintenanceTasks
    $dgDeployments.ItemsSource = $detailDeployments
    $dgSummaryTasks.ItemsSource = $detailSummaryTasks
    $dgIncremental.ItemsSource = $detailIncrementals
    $dgScheduled.ItemsSource = $detailScheduled
    $lblInc.Content = "$sumInc"
    $lblSched.Content = "$sumSched"
    $lblST.Content = "$sumST"
    $lblMT.Content = "$sumMt"
    $lblDepl.Content = "$sumDepl"
})

#endregion

[void]$window.ShowDialog()

#Cleanup
#Get-Variable | Where-Object { $startupVars -notcontains $_.Name } | % { Remove-Variable -Name “$($_.Name)” -Force -Scope “global” }
#Get-Variable | Where-Object { $startupVars -notcontains $_.Name } 

#endregion


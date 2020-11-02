Class Task {
    [int]$id
    [int]$parentId = 0
    [String]$name
    [ValidateSet("STARTING","SUCCESS","ERROR","WARNING","IGNORED")][String]$status = "STARTING"
    [int]$level = 0
    [String]$type
    [String]$command
    [String]$details
    [datetime]$startTime = (Get-Date)
    [datetime]$endTime
    [TimeSpan]$elapsedTime
    [String]$errorMsg
    [String]$function
    $callstack = @()
    [Boolean]$hideTask = $false

    updateTask([Hashtable]$properties,[Boolean]$hidden){
        ForEach($item in $properties.keys){
            $this.$item = $properties[$item]
        }
        $this.endTime = (Get-Date)
        $this.elapsedTime = (New-TimeSpan -Start $this.startTime -End (Get-Date))
        if ( (!$this.hideTask) -and (!$hidden) ) {$this.logTask()}
    }
    updateTask([Hashtable]$properties){
        $this.updateTask($properties,$false)
    }

    logTask(){

        #Determine output color
        $color = $(Switch ($this.status){
            "STARTING" {"cyan"}
            "SUCCESS" {"green"}
            "WARNING" {"yellow"}
            "ERROR" {"red"}
            "IGNORED" {"yellow"}
            })
        
        $LogType = "INFO"
        if ( @("STARTING","SUCCESS").contains($this.status) ){
            $LogType = "INFO"
        } elseif ( $this.status -eq "WARNING" ){
            $LogType = "WARN"
        } elseif ( $this.status -eq "ERROR" ){
            $LogType = "ERROR"
        } elseif ( $this.status -eq "IGNORED" ){
            $LogType = "IGNORE"
        }

        $m = $this.getStatusMessage()
        if ( @("ERROR","IGNORE").contains($this.status) ) {
            $this.errorMsg = $this.getErrorMessage()
        }

        if ( ($this.status -ne "STARTING") -and ($this.level -eq 0) ){
            $m | EPM_Log-Item -WriteHost -LogType $LogType -HostColor $color -IncludeSeparator
        } else {
            $m | EPM_Log-Item -WriteHost -LogType $LogType -HostColor $color 
        }
        

        if ($this.status -eq "STARTING") {
            if ( ($this.command) -or ($this.details)){
                "[COMMAND] : $($this.command) $($this.details)" | EPM_Log-Item
            }
            "[FUNCTION] : $($this.function)" | EPM_Log-Item
            if ($this.callstack.Count -gt 1) {
                "[CALL STACK] : " | EPM_Log-Item
                ForEach ($item in $this.callstack) {"   $($item.trim())" | EPM_Log-Item}
            }
            if ( ($this.command) -or ($this.details)){
                "[COMMAND OUTPUT]" | EPM_Log-Item
            }
        }
    }

    [String] getStatusMessage(){
        [String] $prefix = ""

        # Determine Task Prefix for Logging
        if (@("STARTING","SUCCESS").contains($this.status)) {
            if ($this.level -eq 0){  
                $prefix = "==="
            }
            else {
                $prefix = "---$("--" * $this.level)"
            }
        } else {
            $prefix = "!!!$("!!" * $this.level)"
        }
        # Determine Message
        if ($this.status -eq "STARTING"){
            return "$prefix $($this.type) $(($this.status).PadRight(8,' ')) : $($this.name)"
        } else {
            return "$prefix $($this.type) $(($this.status).PadRight(8,' ')) : $($this.name) [Elapsed Time : $(EPM_Get-ElapsedTime -StartTime $($this.startTime))]"
        }    
    }

    [String] getErrorMessage(){


        #See if we can grab from error log
        if ($this.command) {
            $ErrorLog = (Get-ChildItem "$(Get-Variable EPM_PATH_SCRIPTS -ValueOnly)" `
                            -Filter "$($this.command)*.log" | `
                            Sort-Object LastWriteTime | Select-Object -Last 1)
            if ( $ErrorLog ) {
                $msg = (Get-Content -Path $ErrorLog.FullName | `
                            Select-String -Pattern "^EPM.*-(.*?)" -Context 0,1000 | Out-String)
                $msg = ($msg.Trim().Replace("> ",""))
            } else {
                # Grab Last line from $EPM_LOG_FULL
                $msg = (Get-Content -Path "$(Get-Variable EPM_LOG_FULL -ValueOnly)" -Tail 2)
                $msg = [regex]::Match($msg,"(.*)(EPM.*-[0-9].*:.*)").Groups[2].Value
                # TODO : maybe this is a parent, See if we can grab the error message from a child task? 
            }
        } else {$msg = ""}
        return $msg
    }

    [Hashtable] getProperties(){
        $res = @{}
        $this | Get-Member -MemberType Property | ForEach-Object {
            $res[$_.name] = $this.($_.name)
        }
        return $res
    }

    [String] display(){
        return ($this | Select-Object | Format-List | Out-String).trim() + "`n"
    }
}


Class TaskList {
    [Task[]] $Tasks

    [Task] addTask([Hashtable]$properties,[Boolean]$hidden){
        #Create a Task
        $task = [Task]$properties
        #Set Properties
        $id = $this.Tasks.Count + 1
        if ($task.level -eq 0){
            $type = "TASK"
        } else {
            $type = "SUB-TASK"
        }

        $FullStack = (Get-PSCallStack)
        If ($FullStack[2].FunctionName -eq "<ScriptBlock>") {
            $Function = "$($FullStack[$FullStack.Count-1].Location)"
        } else {
            $Function = "$($FullStack[2].FunctionName) | ARGS : $($FullStack[2].Arguments)"
        }

        $CallStack = @()
        For ($i = 2; $i -lt $FullStack.Count; $i++){
            if ($FullStack[$i].FunctioNName -eq "<ScriptBlock>") {
                $CallStack += "$($FullStack[$i].Location)"
            } else {
                $CallStack += "$($FullStack[$i].FunctionName) | $($FullStack[$i].Location)"
            }
            
        }
        
        #Write-Host -ForegroundColor Yellow "$Function"
       #Write-Host -ForegroundColor Yellow "$CallStack"


        $task.updateTask(@{
            id = $id;
            type = $type;
            function = $Function;
            callstack = $CallStack;
        },$hidden)
        $this.Tasks += $task
        return $task
    }

    [Task] addTask([Hashtable]$properties){
        return ($this.addTask($properties,$false))
    }

    [Task] getTask([String]$name){
        return ($this.Tasks | Where-Object name -eq $name)
    }

    [Task] getTask([Int]$id){
        return ($this.Tasks | Where-Object id -eq $id)
    }

    updateTask([Int]$id,[Hashtable]$properties,[Boolean]$hidden){
        $task = $this.getTask($id)
        $task.updateTask($properties,$hidden)
    }
    updateTask([Int]$id,[Hashtable]$properties){
        updateTask($id,$properties,$false)
    }

    [Int] countTasks([String]$property,[String]$value){
        return ($this.Tasks | Where-Object $property -eq $value).Count
    }

    [TaskList] getTasks([String]$property,[String]$op,[String]$value){
        [TaskList] $TL = [TaskList]::new()
        if ($op -eq "eq"){
            if ($value) {
                $Filtered = ($this.Tasks | Where-Object $property -eq $value)
            } else {
                $Filtered = ($this.Tasks | Where-Object $property -eq "")
            }
        } else {
            if ($value) {
                $Filtered = ($this.Tasks | Where-Object $property -ne $value)
            } else {
                $Filtered = ($this.Tasks | Where-Object $property -ne "")
            }
        }
        ForEach ($task in $Filtered){
            $props = $task.getProperties()
            $newTask = $TL.addTask($props,$true)
            $newTask.updateTask(@{
                id = $task.id;
                startTime = $task.startTime;
                endTime = $task.endTime;
                elapsedTime = $task.elapsedTime 
            },$true)
        }
        return $TL
    }
    [TaskList] getTasks([String]$property,[String]$value){
        return $this.getTasks($property,"eq",$value)
    }

    [TaskList] getTasks([String]$property,[String]$op,[Array]$values){
        [TaskList] $TL = [TaskList]::new()
        ForEach ($value in $values) {
            if ($op -eq "eq"){
                $Filtered = ($this.Tasks | Where-Object $property -eq $value)
            } elseif ($op -eq "ne") {
                $Filtered = ($this.Tasks | Where-Object $property -ne $value)
            } else {
                $Filtered = ($this.Tasks | Where-Object $property -eq $value)
            }
            ForEach ($task in $Filtered){
                $props = $task.getProperties()
                $newTask = $TL.addTask($props,$true)
                $newTask.updateTask(@{
                    id = $task.id;
                    startTime = $task.startTime;
                    endTime = $task.endTime;
                    elapsedTime = $task.elapsedTime 
                },$true)
            }       
        }
        return $TL
    }
    [TaskList] getTasks([String]$property,[Array]$values){
        return $this.getTasks($property,"eq",$values)
    }
}

function New-EPMTaskList(){
    return [TaskList]::new()
}

Export-ModuleMember -Function New-EPMTaskList

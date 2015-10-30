 Function Get-ProcessOwner {
        <#
        .SYNOPSIS
         obtain a list of running processes and add the process owner name to the results                       
        .DESCRIPTION
         obtain a list of running processes and add the process owner name to the results
        .EXAMPLE
         C:\\\\PS>Get-ProcessOwner
                  
        .NOTES
         NAME......:  Get-ProcessOwner
         AUTHOR....:  Craig Grady
         LAST EDIT.:  04/22/2014
         CREATED...:  04/22/2014
        #>
      
        
	# Build a hashtable that associates process ids with owners
    $processOwners = @{}
    Get-WmiObject -Class Win32_Process | ForEach-Object {
        $processOwner = $_.GetOwner()
        # Combine the domain and user information together to get the process owner
        $processOwners[[int]$_.ProcessId] = $processOwner.Domain + '\' + $processOwner.User 
    }
    # Now get all processes and add the owner information to them
    Get-Process | ForEach-Object {
        $processOwner = $null
        # If we have process owner information for the process, look up the owner in the table
        if ($processOwners.ContainsKey($_.Id)) {
            $processOwner = $processOwners[$_.Id]
        }
        # Add the owner information to the current process object
        Add-Member -InputObject $_ -MemberType NoteProperty -Name Owner -Value $processOwner
        # Return the current process object from the script block
        $_
    }
} 
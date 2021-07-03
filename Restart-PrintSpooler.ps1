Function Restart-PrintSpooler
{
<#
    .SYNOPSIS
        This CmdLet restart the print spooler service

    .DESCRIPTION
        This CmdLet enables an Administrator to restart the print spooler service and its dependent service LPD service remotely.
        
        To enable a none admin user to restart the Print Spooler service themselves:
        A GPO Computer Configuration Policy must be setup
        > Windows Settings > Security Settings > System Services
        Edit LPD Service and Print Spooler Properties > Startup = Automatic 
        Edit permissions and add the AD Users Group of your LSU with Security Permissions: Start, Stop and Pause - Allow.
        - With the above configuration a none admin user will be able to restart the Print Spooler service themselves.

    .PARAMETER ComputerName
        This parameter is not mandatory. By default the local host is set.

    .PARAMETER LogPath
        This parameter is not mandatory.

    .EXAMPLE
        Restart-PrintSpooler

        Simply restart the services on the local host

    .EXAMPLE
        Restart-PrintSpooler -verbose
        
        Restart the services on the local host with verbose mode turned on
    .EXAMPLE
        Restart-PrintSpooler -ComputerName ComputerName -LogPath 'C:\localData\ComputerName'
    
    .OUTPUTS
        A log file will be created here: '\\ServerName\logs$\SpoolerService'
        Message box will be showed for the user when the job is finished.
    
    .NOTES
        Version: 2.0
        Created by MisterDDM

#>

    [cmdletbinding()]
    Param
    (
        [Parameter(Mandatory=$false, 
                   ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true, 
                   ValueFromRemainingArguments=$false, 
                   Position=0,
                   ParameterSetName='ComputerName')]
        [String[]]$ComputerName = $env:COMPUTERNAME,
        
        [Parameter(Mandatory=$false, 
                   ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true, 
                   ValueFromRemainingArguments=$false, 
                   Position=1,
                   ParameterSetName='ComputerName')]
        [String]$LogPath
    )

    Begin 
    {
        if($PSBoundParameters['verbose']) 
        {
            $VerbosePreference = "Continue"
            Write-Verbose "$(Get-Date -Format HH:mm:ss.fff) `t Verbose mode is on"
        }
        else 
        {
            $VerbosePreference = "SilentlyContinue"
        }
        
        if($PSBoundParameters['LogPath']) 
        {
            if ( (Test-Path ${LogPath}) -eq $false )
            {
                New-Item ${LogPath} -ItemType Directory
                Write-Verbose "${LogPath} Created"
            }
        }
        
        if ( $Error[0] )
        {
            $Error.Clear()
            Write-Verbose "Previous error messages cleared"
        }
    }

    Process
    {
        if ($pscmdlet.ShouldProcess($ComputerName))
        {

            if ( Test-Connection $ComputerName -Quiet -Count 1 )
            {
                Write-Verbose 'Restarting Services Print Spooler and dependent service LPD Service'
                try 
                {
                    if ($PSBoundParameters['verbose'])
                    {
                        Get-Service -Name Spooler -ComputerName $ComputerName | Restart-Service -Force -Verbose
                    }
                    else 
                    {
                        Get-Service -Name Spooler -ComputerName $ComputerName | Restart-Service -Force
                    }
                
                }
                catch 
                {
                    if ( $Error[0] )
                    {
                        Write-Verbose 'There was an error'
                        Write-Verbose $Error[0].Exception.Message
                    }
                }

                if($PSBoundParameters['LogPath']) 
                {
                    $Object = [PSCustomObject]@{
                        ComputerName = $($ComputerName)
                        Date = $(Get-Date)
                        Errors = $($Error[0].Exception.Message)
                    }

                    Write-Verbose 'Exporting information to a csv file'
                    Write-Verbose "${LogPath}\$env:USERNAME.csv"
                    Write-Verbose '.'
                    Write-Verbose $Object

                    Export-Csv -Path "${LogPath}\$env:USERNAME.csv" -InputObject $Object -Encoding Default -Force -Delimiter ',' -Append -NoTypeInformation 
                }
            
            }
            else 
            {
                Write-Output $ComputerName is offline
            }
        }

        $Null = [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
        
        if ( $PSBoundParameters['ComputerName'] -eq $env:COMPUTERNAME )
        {
            $Message = "The Print Spooler Service has been restarted. If you are still unable to print, please contact your LSA."
        }

        else 
        {
            $Message = "The Print Spooler Service has been restarted on the remote Computer $ComputerName"
        }

        $Caption = "Print Spooler Service Restart"
        $Btn = [System.Windows.Forms.MessageBoxButtons]::OK
        $Icn = [System.Windows.Forms.MessageBoxIcon]::Information
        $Result = [System.Windows.Forms.MessageBox]::Show($Message,$Caption,$Btn,$Icn)
    }
    end
    {
        
    }
}

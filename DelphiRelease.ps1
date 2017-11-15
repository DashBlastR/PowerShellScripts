#20170920ab - script now works!
#20170323ab - script is not copying files from source to dest the first run.  It expects the directory
#             folders to already exist.  On the second run (same version folder), it doesn't copy
#             source subfolders to dest
#######################################################################################################
#######################################################################################################
#######################################################################################################

# add the required .NET assembly:
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName Microsoft.VisualBasic

# # Variables # #
    $localDirectory = "C:\OKApps\Development\"
    #$targetDirectory = "C:\tmp\MIS\Development\"
    $targetDirectory = "\\oks0009\MIS\Development\"

Function CloseMe([string]$inputString, [string]$errorCode)
{
    If(!($inputString))
    {
        $inputString = ("User cancelled process")
    }
    
    If($errorCode)
    {
        $inputString += (' - ' + $errorCode)
    }
    
    Write-Host ($inputString)
    Write-Host 'Terminating Application . . .'
    Exit 
}

Function Get-FileName([string]$initialDirectory)
{   
    $return = "" | Select-Object -Property GetMyBaseName, GetMyName, GetMyPath, GetMyVersion
    
    $OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $OpenFileDialog.initialDirectory = $initialDirectory
    $OpenFileDialog.filter = "Executable (*.exe)| *.exe|All files (*.*)| *.*"
    $OpenFileDialog.ShowDialog() | Out-Null

    $temp = $OpenFileDialog.FileName
    $return.GetMyPath = $temp
    $return.GetMyVersion = (Get-Item $temp).VersionInfo.FileVersion

    $return.GetMyBaseName = (Get-Item $temp).BaseName
    $return.GetMyName = (Get-Item $temp).Name
    
    If(!($OpenFileDialog.filename))
    {
        CloseMe 'No folder selected' 99 
    }
    
    return $return
}

Function VerifyContinue([string]$processName, [string]$processVal)
{
    $title = ($processVal + " already exists!")
    $msg = ($title + "`r`n`nDo you want to overwrite?")
    $msg = ($msg + "`r`n`nClick'Yes' to overwrite, 'No' to enter a new version or 'Cancel' to quit")
    $result = [System.Windows.Forms.MessageBox]::Show($msg, $title,
              [System.Windows.Forms.MessageBoxButtons]::YesNoCancel,
              [System.Windows.Forms.MessageBoxIcon]::Warning)

    switch($result)
    {
        "Yes" {
                $result = 0
        }
        "No" {
                $result = 1
        }
        "Cancel" {
                $result = 2
        }
    }

    return $result
}

Function CheckExistingFolder([string]$targetFolder, [string]$fileVersion)
{
    $checkVal = $null
    while ($checkVal -ne 0 -or !$versionNbr)
    {
        If(!$fileVersion)
        {
            $fileVersion = Get-ChildItem ($targetDirectory + ($targetFolder)) -erroraction stop| ?{ $_.PSIsContainer } | sort Name | select -last 1
        }

        $LastDir = "Release {0}" -f $fileVersion

        $versionNbr = [Microsoft.VisualBasic.Interaction]::InputBox("Please enter the folder name of this release", "Version Number", $LastDir)

        If(!$versionNbr)
        {
            $versionNbr = "CURRENT_VERSION"
        }

        $targetPath = "{0}{1}\{2}" -f $targetDirectory, $targetFolder, $versionNbr.trim()

        If(Test-Path -Path $targetPath)
        {
            $checkVal = VerifyContinue "Folder Exists!" $targetPath
            if($checkVal -eq 2)
            {
               CloseMe
            }
        }
        Else
        {
            $checkVal = 0
        }
    }
    return $targetPath
}

Function SEND-ZIP ($ZipFilename, $Filename) 
{ 
    # 
    # Check to see if the Zip file exists, if not create a blank one 
    # 
    If ( (TEST-PATH $ZipFilename) -eq $FALSE ) 
    {
        Compress-Archive $Filename -destinationPath $ZipFilename
    } 
} 

# *** Entry Point to Script ***

#Determine the source folder/file
    $getMyFile = Get-FileName -initialDirectory $localDirectory
    $sourcePath = split-path -Path $getMyFile.GetMyPath
    $sourceVer = $getMyFile.GetMyVersion

#Determine the version for target
    $targetdir = CheckExistingFolder (split-path $sourcePath -leaf) ($sourceVer)


#Copy source to destination
    write-output ('Copying ' + $sourcePath + ' to ' + $targetDir)
    New-Item -Path $targetDir -ItemType directory
    Get-ChildItem $sourcePath -recurse -exclude *.dcu,*.dsk | copy-item -destination $targetDir

#Clean out final directory
    write-output 'Removing unnecessary folders/files. . .'
    remove-item ($targetdir + '\' + '*.dcu') 
    remove-item ($targetdir + '\' + '*.dsk') 

#Zip the executable
    write-output 'Zipping the .EXE'
    [string]$ZipSource = "{0}\{1}" -f $targetDir, $getMyFile.GetMyName
    [string]$ZipDest = "{0}\{1}{2}" -f $targetDir, $getMyfile.GetMyBaseName, ".zip"
    SEND-ZIP $ZipDest $ZipSource
    
CloseMe('Finished!')

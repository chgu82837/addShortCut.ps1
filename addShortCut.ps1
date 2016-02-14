
Function Help
{
    echo "addShortCut.ps1"
    echo "create a shortcut for the command or executable"
    echo "================================================"
    echo "Usage:"
    echo "addShortCut [-d|-destination <path>] <command_or_path> [<more_command_or_path> ...]"
    exit 255
}

Function Set-ShortCut
{
    Param ([string] $src, [string] $des)
    $WshShell = New-Object -comObject WScript.Shell
    $Shortcut = $WshShell.CreateShortcut($des)
    $Shortcut.TargetPath = $src
    $Shortcut.Save()
}

If ($args.count -eq 0){
    Help
}

If ("$args[0]".StartsWith("-d")){
    If ($args.count -le 2){
        Help
    }
    $destination = $args[1]
    $args = [System.Collections.Generic.List[System.Object]]($args)
    $args.RemoveAt(0)
    $args.RemoveAt(0)
}

If ($destination.Length.Equals(0)){
    $destination = "$(pwd)"
}
Else {
    If (Test-Path $destination) {
        $destination = (Resolve-Path $destination).Path
    }
    Else {
        Help
    }
}
        
ForEach ($cp in $args){
    if (Test-Path $cp){
        $target = (Reslove-Path $cp).Path
    }
    Else {
        $target = (Get-Command $cp).Source 2> $null
    }
    If ($target) {
        $name = [System.IO.Path]::GetFileNameWithoutExtension($target)
        $shortcut = "$destination\$name.lnk"
        echo "Creating shortcut $name at $shortcut links to $target"
        Set-ShortCut $target $shortcut
    }
    Else {
        echo "$cp not found"
    }
}


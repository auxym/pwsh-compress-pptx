Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'

function Get-FfprobeData {
    [cmdletbinding()]
    param (
        [parameter(mandatory = $true, valuefrompipeline = $true)]
        [string]
        $Path
    )

    process {
        $ffprobe_output = & ffprobe -v quiet -print_format json -show_format -show_streams $Path
        $ffprobe_output | ConvertFrom-Json
    }
}

function isVideo {
    [cmdletbinding()]
    param (
        [parameter(mandatory = $true, valuefrompipeline = $true)]
        [string]
        $Path
    )

    process {
        $info = Get-FfprobeData $Path
        ($info.streams | Where-Object { [int]$_.nb_frames } | Measure-Object).Count -gt 0
    }

}
function Get-VideoFiles {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true, ValueFromPipeline = $true)]
        [string]
        $Path,

        [Parameter]
        [switch]
        $Recurse
    )
    begin {
        $video_extensions = 
        @('.mp4', '.avi', '.mts', '.webm', '.mkv', '.mov', '.wmv', '.flv', 'ogv', '.gifv', '.m4v', '.mpg', '.mpeg', '.3gp')
        # TODO Supported animated gif, png, webp
        # Need to change image relationship type to media type
        # And <a:blip> tag in slide xml to <p14:media> tag
    }

    process {
        Get-ChildItem -Path $Path -File -Recurse:$Recurse -Filter "media*.*"
        | Where-Object { $_.Extension -in $video_extensions -and (isVideo $_) }
    }
}

function transcode-video {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true, ValueFromPipeline = $true)]
        [string]
        $Path
    )

    begin {
        $ffmpeg_base_opts = @(
            '-threads', 2,
            '-c:a', 'aac',
            '-b:a', '320k', 
            '-c:v', 'libsvtav1',
            '-crf', 35,
            '-preset', 9
            #,'-t', '00:00:01'
        )

        $scale_opts = @(
            # Scale to 1080 vertical or horizontal resolution depending on
            # orientation of source video
            '-vf', "scale='if(gt(iw,ih),-1,1080)':'if(gt(iw,ih),1080,-1)'"
        )
    }

    process {
        $info = Get-FfprobeData $Path
        $ffmpeg_opts = $ffmpeg_base_opts

        # Min of horizontal and vertical resolution
        # Try / catch used to ignore streams that don't have width/height properties (eg audio streams)
        $res = ($info.streams | ForEach-Object { try {@($_.width, $_.height)} catch {}} | Measure-Object -Minimum).Minimum
        if ($res -gt 1080) {
            # Rescale only if source is larger than 1080 resolution
            $ffmpeg_opts += $scale_opts
        }

        $input_item = Get-Item $Path
        $tmpname = "$(New-Guid).mp4"
        $output = Join-Path $input_item.Directory "$($input_item.BaseName).mp4" 
        & ffmpeg -i $input_item $ffmpeg_opts $tmpname
        if ($LASTEXITCODE -eq 0) {
            Remove-Item $input_item
            Move-Item $tmpname $output
            Get-Item $output
        }
        else {
            Remove-Item $tmpname
            Write-Error "ffmpeg returned exit code $LASTEXITCODE"
            $input_item
        }
    }
}

function Update-Rels {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [string]
        $RootPath,

        [Parameter(Mandatory = $true)]
        [string]
        $OldName,

        [Parameter(Mandatory = $true)]
        [string]
        $NewName
    )

    $src_exp = "Target=""([^""]*)/$($OldName -replace '\.', '\.')"""
    $rep_exp = 'Target="$1/' + $NewName + '"' 

    Join-Path $RootPath "ppt/slides/_rels" | Get-ChildItem -File -Filter '*.xml.rels'
    | Where-Object { Select-String $_ -Pattern $src_exp }
    | ForEach-Object {
        Write-Verbose "Modify rels in $_"
        $content = Get-Content $_
        ($content -replace $src_exp, $rep_exp) | Out-File $_ -Encoding utf8NoBOM
    }
}

function Update-ContentTypes {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [string]
        $RootPath
    )
    # Powershell needs to espace the brackets here
    $cTypesXmlFile = Join-Path $RootPath '`[Content_Types`].xml' | Get-Item

    [xml]$doc = Get-Content -LiteralPath $cTypesXmlFile.FullName

    # Create the namespace manager
    $nsMgr = New-Object System.Xml.XmlNamespaceManager($doc.NameTable)
    $nsMgr.AddNamespace("t", "http://schemas.openxmlformats.org/package/2006/content-types")

    # Check if we already have a <Default> element for the .mp4 extension,
    # if not, we add it.
    if (-not $doc.SelectSingleNode("//t:Default[@Extension='mp4']", $nsMgr)) {
        # Find the <Types> root element
        $root = $doc.DocumentElement

        # Create a new <Default> node
        $newDefault = $doc.CreateElement("Default", $root.NamespaceURI)
        $newDefault.SetAttribute("Extension", "mp4")
        $newDefault.SetAttribute("ContentType", "video/mp4")

        # Insert the new node after the last <Default> element
        $lastDefault = $root.SelectNodes("t:Default", $nsMgr) | Select-Object -Last 1
        $root.InsertAfter($newDefault, $lastDefault) | Out-Null

        # Save the modified XML back to file
        # We need an absolute path here
        $doc.Save($cTypesXmlFile.FullName)
    }
}

function Compress-PptxMedia {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [string]
        $Path,

        [Parameter(Mandatory = $true)]
        [string]
        $DestinationPath,

        [Parameter()]
        [string]
        $SizeThreshold = 0,

        [Parameter()]
        [switch]
        $Overwrite
    )
    
    $bitrate_threshold = 4000 * 1000 # b/s
    
    #$tempdir = Get-Item (Join-Path $env:TEMP 'fc50475a-80df-47de-9e92-b57dbb8f45a9')
    #Get-ChildItem $tempdir | Remove-Item -Recurse -Force
    $tempdir = New-Item -Type Directory -Path (Join-Path $env:TEMP $(New-Guid))
    Expand-Archive -Path $Path -Destination $tempdir -Force
    
    $media_dir = Join-Path $tempdir "ppt/media"

    $transcodedVideoFiles = Get-VideoFiles $media_dir
    | Where-Object { $_.Length -ge $SizeThreshold }
    | Where-Object { [int]((Get-FfprobeData $_).format.bit_rate) -gt $bitrate_threshold }
    | ForEach-Object {
        $newfile = transcode-video -Path $_
        if ($newfile.Name -ne $_.Name) {
            Update-Rels -RootPath $tempdir -OldName $_.Name -NewName $newfile.Name
        }
        $newfile
    }

    # If we created new .mp4 files in the pptx, and there were none before,
    # we need to update the [ContentTypes].xml file to add a default content
    # type for the .mp4 extension.
    if (($transcodedVideoFiles | Measure-Object).Count -gt 0) {
        Update-ContentTypes -RootPath $tempdir
    }

    Compress-Archive -Path (Join-Path $tempdir "*") -DestinationPath $DestinationPath -CompressionLevel "Optimal" -Force:$Overwrite

    Remove-Item -Recurse -Force $tempdir
}

Export-ModuleMember -Function Compress-PptxMedia

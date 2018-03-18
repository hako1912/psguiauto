$Win32 = &{
    $cscode = @"
        [DllImport("USER32.dll")]
        public static extern void mouse_event(int dwFlags, int dx, int dy, int cButtons, int dwExtraInfo);

        [DllImport("USER32.dll")]
        public static extern void SetCursorPos(int X, int Y);
"@
    return (add-type -memberDefinition $cscode -name "Win32ApiFunctions" -passthru)
}

Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName mscorlib 

function click(){
    $win32::mouse_event(2, 0, 0, 0, 0) # left down
    $win32::mouse_event(4, 0, 0, 0, 0) # left up
}

function rightClick(){
    $win32::mouse_event(8, 0, 0, 0, 0) # right down
    $win32::mouse_event(16, 0, 0, 0, 0) # right up
}

function middleClick(){
    $win32::mouse_event(32, 0, 0, 0, 0) # right down
    $win32::mouse_event(64, 0, 0, 0, 0) # right up
}

function wheel($amount){
    $win32::mouse_event(2048, 0, 0, $amount, 0) # wheel
}

function mouseMove($x, $y, $ms){
    if($ms -eq $Null -or $ms -lt 1){
        $ms = 50
    }
    $nowX = [System.Windows.Forms.Cursor]::Position.X
    $nowY = [System.Windows.Forms.Cursor]::Position.Y

    $dx = $x - $nowX
    $dy = $y - $nowY

    for($i=1; $i -lt $ms; $i++){
        $percent =  $i / $ms

        $nextX = ($dx * $percent) + $nowX 
        $nextY = ($dy * $percent) + $nowY
        $Win32::SetCursorPos($nextX, $nextY)
        # Write-Host "${percent}: ($nextX, $nextY)"
        Start-Sleep -Milliseconds 1
    }
}

function typeWrite($msg){
    $msg.ToCharArray() | %{
        [Windows.Forms.Sendkeys]::SendWait($_)
        Start-Sleep -Milliseconds 10 # animation
    }
}

function locateOnScreen ($imagePath) {
    if($imagePath -eq $Null){
        Write-Host "imagePath = Null"
        return $Null
    }

    # screen capture
    [Windows.Forms.Sendkeys]::SendWait("{PrtSc}")
    Start-Sleep -Milliseconds 100
    $screenBitmap = New-Object System.Drawing.Bitmap ([Windows.Forms.Clipboard]::GetImage())
    $screen = toGrayScale($screenBitmap)

    # target image
    $targetBitmap = New-Object System.Drawing.Bitmap $imagePath
    $target = toGrayScale($targetBitmap)

    for($sx=0; $sx -lt $screen.Length; $sx++){
        if(($sx + $target.Length) -gt $screen.Length){
            break
        }
        for($sy=0; $sy -lt $screen[0].Length; $sy++){
            if(($sy + $target[0].Length) -gt $screen[0].Length){
                break
            }
            # Write-Host "($sx, $sy)"
            if($screen[$sx][$sy] -eq $target[0][0]){
                
                :targetLoop for($tx=0; $tx -lt $target.Length; $tx++){
                    for($ty=0; $ty -lt $target[0].Length; $ty++){
                        if($screen[$sx + $tx][$sy + $ty] -ne $target[$tx][$ty]){
                            break targetLoop
                        }
                    }
                }
                if($tx -eq $target.Length -and $ty -eq $target[0].Length){
                    Write-Host "find!!! ($sx, $sy)"
                    $ret = New-Object PSObject -Property @{x=$sx; y=$sy}
                    return $ret
                }
            }
        }
    }
    return $Null
}

function toGrayScale ($bitmap) {
    $resolution = 1
    $lengthX = [math]::round($bitmap.Width / $resolution)
    $lengthY = [math]::round($bitmap.Height / $resolution)
    $byte = New-Object System.Byte[][] (($lengthX), ($lengthY))

    $height = $bitmap.Height
    $width = $bitmap.Width
    
    $rect = New-Object System.Drawing.Rectangle (0, 0, $width, $height)
    $lockMode = [System.Drawing.Imaging.ImageLockMode]::ReadOnly
    $bitmapData = $bitmap.LockBits($rect, $lockMode, $bitmap.PixelFormat)
    $scan0 = $bitmapData.Scan0
    
    $iy = 0
    $idx = 0
    for($y = 0; $y -lt $height; $y += $resolution){
        $ix = 0
        for($x = 0; $x -lt $width; $x += $resolution){
            $r = [System.Runtime.InteropServices.Marshal]::ReadByte($scan0, $idx)
            $g = [System.Runtime.InteropServices.Marshal]::ReadByte($scan0, $idx + 1)
            $b = [System.Runtime.InteropServices.Marshal]::ReadByte($scan0, $idx + 2)
            # Write-Host "[$x][$y] -> [$ix][$iy]"
            $byte[$ix++][$iy] = [byte](($r + $g + $b) / 3)
            $idx += 4
        }
        $iy++
    }
    return $byte
}


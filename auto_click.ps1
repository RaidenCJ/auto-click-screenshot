# Auto Click + Screenshot + Send to Feishu API - Headless Compatible Version V2
# Uses P/Invoke via delegate (no Add-Type compilation) for headless environments
# Works when RDP session is disconnected

param(
    [int]$X1, [int]$Y1,
    [int]$X2, [int]$Y2,
    [int]$X3, [int]$Y3,
    [int]$X4, [int]$Y4,
    [int]$X5, [int]$Y5,
    [string]$AppId,
    [string]$AppSecret,
    [string]$UserId
)

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# ========== Load Config ==========
$configPath = "$PSScriptRoot\auto_click_config.json"
if (Test-Path $configPath) {
    $config = Get-Content $configPath | ConvertFrom-Json
} else {
    Write-Host "ERROR: Config file not found!"
    exit 1
}

if (-not $X1) {
    $coords = @($config.coordinates)
    $X1 = $coords[0].x; $Y1 = $coords[0].y
    $X2 = $coords[1].x; $Y2 = $coords[1].y
    $X3 = $coords[2].x; $Y3 = $coords[2].y
    $X4 = $coords[3].x; $Y4 = $coords[3].y
    $X5 = $coords[4].x; $Y5 = $coords[4].y
}
if (-not $AppId) {
    $AppId = $config.app_id
    $AppSecret = $config.app_secret
    $UserId = $config.user_id
}

# ========== Win32 API via P/Invoke (no compilation) ==========
$ApiCode = @'
using System;
using System.Runtime.InteropServices;

public static class Win32Api {
    [DllImport("user32.dll")]
    public static extern int GetSystemMetrics(int nIndex);
    
    [DllImport("user32.dll")]
    public static extern IntPtr GetDC(IntPtr hWnd);
    
    [DllImport("user32.dll")]
    public static extern int ReleaseDC(IntPtr hWnd, IntPtr hDC);
    
    [DllImport("gdi32.dll")]
    public static extern IntPtr CreateCompatibleDC(IntPtr hdc);
    
    [DllImport("gdi32.dll")]
    public static extern IntPtr CreateCompatibleBitmap(IntPtr hdc, int nWidth, int nHeight);
    
    [DllImport("gdi32.dll")]
    public static extern IntPtr SelectObject(IntPtr hdc, IntPtr hObject);
    
    [DllImport("gdi32.dll")]
    public static extern bool BitBlt(IntPtr hdcDest, int nXDest, int nYDest, int nWidth, int nHeight, IntPtr hdcSrc, int nXSrc, int nYSrc, int dwRop);
    
    [DllImport("gdi32.dll")]
    public static extern bool DeleteDC(IntPtr hdc);
    
    [DllImport("gdi32.dll")]
    public static extern bool DeleteObject(IntPtr hObject);
    
    [DllImport("user32.dll", SetLastError = true)]
    public static extern uint SendInput(uint nInputs, INPUT[] pInputs, int cbSize);
    
    [DllImport("user32.dll")]
    public static extern bool SetCursorPos(int x, int y);
    
    [DllImport("user32.dll")]
    public static extern void mouse_event(uint dwFlags, uint dx, uint dy, uint cButtons, uint dwExtraInfo);
    
    [DllImport("user32.dll")]
    public static extern bool SetProcessDPIAware();
    
    [DllImport("user32.dll")]
    public static extern IntPtr SendMessage(IntPtr hWnd, uint Msg, IntPtr wParam, IntPtr lParam);
    
    [DllImport("user32.dll")]
    public static extern IntPtr PostMessage(IntPtr hWnd, uint Msg, IntPtr wParam, IntPtr lParam);
    
    [DllImport("user32.dll")]
    public static extern IntPtr GetDesktopWindow();
    
    [DllImport("user32.dll")]
    public static extern IntPtr GetForegroundWindow();
    
    [DllImport("user32.dll")]
    public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);
    
    [StructLayout(LayoutKind.Sequential)]
    public struct INPUT {
        public uint type;
        public InputUnion U;
    }
    
    [StructLayout(LayoutKind.Explicit)]
    public struct InputUnion {
        [FieldOffset(0)] public MOUSEINPUT mi;
        [FieldOffset(0)] public KEYBOARDINPUT ki;
    }
    
    [StructLayout(LayoutKind.Sequential)]
    public struct MOUSEINPUT {
        public int dx;
        public int dy;
        public uint mouseData;
        public uint dwFlags;
        public uint time;
        public IntPtr dwExtraInfo;
    }
    
    [StructLayout(LayoutKind.Sequential)]
    public struct KEYBOARDINPUT {
        public ushort wVk;
        public ushort wScan;
        public uint dwFlags;
        public uint time;
        public IntPtr dwExtraInfo;
    }
    
    public const int SM_CXSCREEN = 0;
    public const int SM_CYSCREEN = 1;
    public const int SM_CXVIRTUALSCREEN = 78;
    public const int SM_CYVIRTUALSCREEN = 79;
    public const int SM_XVIRTUALSCREEN = 76;
    public const int SM_YVIRTUALSCREEN = 77;
    
    public const uint MOUSEEVENTF_MOVE = 0x0001;
    public const uint MOUSEEVENTF_LEFTDOWN = 0x0002;
    public const uint MOUSEEVENTF_LEFTUP = 0x0004;
    public const uint MOUSEEVENTF_ABSOLUTE = 0x8000;
    public const uint MOUSEEVENTF_VIRTUALDESK = 0x4000;
    public const uint MOUSEEVENTF_MOVE_NOCOALESCE = 0x2000;
    
    public const uint WM_LBUTTONDOWN = 0x0201;
    public const uint WM_LBUTTONUP = 0x0202;
    public const uint MK_LBUTTON = 0x0001;
    
    public const int SRCCOPY = 0x00CC0020;
    
    private static bool dpiSet = false;
    
    public static void EnsureDPI() {
        if (!dpiSet) {
            try { SetProcessDPIAware(); } catch { }
            dpiSet = true;
        }
    }
    
    public static int[] GetScreenSize() {
        EnsureDPI();
        int w = GetSystemMetrics(SM_CXSCREEN);
        int h = GetSystemMetrics(SM_CYSCREEN);
        if (w <= 0 || h <= 0) {
            w = GetSystemMetrics(SM_CXVIRTUALSCREEN);
            h = GetSystemMetrics(SM_CYVIRTUALSCREEN);
        }
        if (w <= 0) w = 1920;
        if (h <= 0) h = 1080;
        return new int[] { w, h };
    }
    
    public static bool SendClick(int x, int y) {
        EnsureDPI();
        int[] screen = GetScreenSize();
        int w = screen[0], h = screen[1];
        
        // Normalize to 0-65535
        int normX = (int)(((double)x / w) * 65535.0);
        int normY = (int)(((double)y / h) * 65535.0);
        normX = Math.Max(0, Math.Min(65535, normX));
        normY = Math.Max(0, Math.Min(65535, normY));
        
        var inputs = new INPUT[3];
        int size = Marshal.SizeOf(typeof(INPUT));
        
        // Move
        inputs[0].type = 0;
        inputs[0].U.mi.dx = normX;
        inputs[0].U.mi.dy = normY;
        inputs[0].U.mi.dwFlags = MOUSEEVENTF_MOVE | MOUSEEVENTF_ABSOLUTE | MOUSEEVENTF_VIRTUALDESK | MOUSEEVENTF_MOVE_NOCOALESCE;
        
        // Down
        inputs[1].type = 0;
        inputs[1].U.mi.dwFlags = MOUSEEVENTF_LEFTDOWN | MOUSEEVENTF_ABSOLUTE | MOUSEEVENTF_VIRTUALDESK;
        
        // Up
        inputs[2].type = 0;
        inputs[2].U.mi.dwFlags = MOUSEEVENTF_LEFTUP | MOUSEEVENTF_ABSOLUTE | MOUSEEVENTF_VIRTUALDESK;
        
        uint sent = SendInput(3, inputs, size);
        return sent == 3;
    }
    
    public static bool SendClickFallback(int x, int y) {
        EnsureDPI();
        try {
            SetCursorPos(x, y);
            System.Threading.Thread.Sleep(50);
            mouse_event(MOUSEEVENTF_LEFTDOWN | MOUSEEVENTF_LEFTUP, 0, 0, 0, 0);
            return true;
        } catch { return false; }
    }
    
    public static void SaveScreenshot(string path, int w, int h) {
        IntPtr hdcScreen = GetDC(IntPtr.Zero);
        if (hdcScreen == IntPtr.Zero) {
            throw new Exception("GetDC failed");
        }
        
        IntPtr hdcMem = CreateCompatibleDC(hdcScreen);
        IntPtr hBitmap = CreateCompatibleBitmap(hdcScreen, w, h);
        IntPtr hOld = SelectObject(hdcMem, hBitmap);
        
        BitBlt(hdcMem, 0, 0, w, h, hdcScreen, 0, 0, SRCCOPY);
        
        var bmp = System.Drawing.Image.FromHbitmap(hBitmap);
        bmp.Save(path, System.Drawing.Imaging.ImageFormat.Png);
        bmp.Dispose();
        
        SelectObject(hdcMem, hOld);
        DeleteObject(hBitmap);
        DeleteDC(hdcMem);
        ReleaseDC(IntPtr.Zero, hdcScreen);
    }
}
'@

try {
    Add-Type -TypeDefinition $ApiCode -Language CSharp -ReferencedAssemblies @("System.Drawing.dll") -ErrorAction Stop
    $useWin32 = $true
    Write-Host "Win32 API loaded successfully"
} catch {
    Write-Host "WARNING: Could not load Win32 API: $_"
    $useWin32 = $false
}

# ========== Feishu API Functions ==========
function Get-FeishuToken {
    param($AppId, $AppSecret)
    try {
        $body = @{app_id=$AppId; app_secret=$AppSecret} | ConvertTo-Json
        $r = Invoke-RestMethod -Uri "https://open.feishu.cn/open-apis/auth/v3/app_access_token/internal" -Method Post -Body $body -ContentType "application/json" -TimeoutSec 10
        if ($r.code -eq 0) { return $r.app_access_token }
        Write-Host "ERROR: $($r.msg)"
    } catch { Write-Host "ERROR: $_" }
    return $null
}

function Upload-Image {
    param($Path, $Token)
    try {
        $bytes = [IO.File]::ReadAllBytes($Path)
        $name = Split-Path $Path -Leaf
        $boundary = [Guid]::NewGuid().ToString()
        $LF = "`r`n"
        
        $ms = New-Object IO.MemoryStream
        $w = New-Object IO.StreamWriter($ms)
        $w.Write("--$boundary$LF")
        $w.Write("Content-Disposition: form-data; name=`"image_type`"$LF$LF")
        $w.Write("message$LF")
        $w.Write("--$boundary$LF")
        $w.Write("Content-Disposition: form-data; name=`"image`"; filename=`"$name`"$LF")
        $w.Write("Content-Type: image/png$LF$LF")
        $w.Flush()
        $ms.Write($bytes, 0, $bytes.Length)
        $w.Write("$LF--$boundary--$LF")
        $w.Flush()
        
        $r = Invoke-RestMethod -Uri "https://open.feishu.cn/open-apis/im/v1/images" -Method Post -Headers @{"Authorization"="Bearer $Token"} -ContentType "multipart/form-data; boundary=$boundary" -Body $ms.ToArray() -TimeoutSec 30
        $w.Dispose(); $ms.Dispose()
        
        if ($r.code -eq 0) { return $r.data.image_key }
        Write-Host "ERROR: $($r.msg)"
    } catch { Write-Host "ERROR: $_" }
    return $null
}

function Send-ToFeishu {
    param($ImageKey, $Token, $OpenId)
    try {
        $body = @{
            receive_id = $OpenId
            msg_type = "image"
            content = "{`"image_key`":`"$ImageKey`"}"
        } | ConvertTo-Json -Depth 10
        
        $r = Invoke-RestMethod -Uri "https://open.feishu.cn/open-apis/im/v1/messages?receive_id_type=open_id" -Method Post -Headers @{"Authorization"="Bearer $Token"; "Content-Type"="application/json"} -Body ([Text.Encoding]::UTF8.GetBytes($body)) -TimeoutSec 10
        if ($r.code -eq 0) { Write-Host "OK - Sent!"; return $true }
        Write-Host "ERROR: $($r.msg)"
    } catch { Write-Host "ERROR: $_" }
    return $false
}

function Invoke-Click {
    param([int]$X, [int]$Y)
    
    if ($useWin32) {
        # Method 1: SendInput
        if ([Win32Api]::SendClick($X, $Y)) {
            Write-Host "  Click sent via SendInput"
            return $true
        }
        # Method 2: Fallback
        if ([Win32Api]::SendClickFallback($X, $Y)) {
            Write-Host "  Click sent via mouse_event"
            return $true
        }
    }
    
    # Method 3: PowerShell Cursor (last resort)
    try {
        [System.Windows.Forms.Cursor]::Position = New-Object System.Drawing.Point($X, $Y)
        Start-Sleep -Milliseconds 50
        Add-Type -MemberDefinition '[DllImport("user32.dll")] public static extern void mouse_event(uint flags, uint dx, uint dy, uint cButtons, uint info);' -Name U32 -Namespace W -ErrorAction SilentlyContinue
        [W.U32]::mouse_event(0x02 -bor 0x04, 0, 0, 0, 0)
        Write-Host "  Click sent via Cursor.Position"
        return $true
    } catch {
        Write-Host "  ERROR: All click methods failed"
        return $false
    }
}

function Screenshot-Send {
    param($Token, $OpenId, $PosNum)
    Start-Sleep -Seconds 2
    
    $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
    $dir = "$env:USERPROFILE\Desktop\auto_click_screenshots"
    if (-not (Test-Path $dir)) { mkdir $dir | Out-Null }
    $path = "$dir\screenshot_$timestamp.png"
    
    # Get screen size
    $w = 1920; $h = 1080
    if ($useWin32) {
        $size = [Win32Api]::GetScreenSize()
        $w = $size[0]; $h = $size[1]
    }
    
    # Try BitBlt screenshot
    $screenshotOk = $false
    if ($useWin32) {
        try {
            [Win32Api]::SaveScreenshot($path, $w, $h)
            Write-Host "[$(Get-Date -Format 'HH:mm:ss')] Screenshot $PosNum saved (${w}x${h}, BitBlt)"
            $screenshotOk = $true
        } catch {
            Write-Host "WARNING: BitBlt screenshot failed: $_"
        }
    }
    
    # Fallback: CopyFromScreen
    if (-not $screenshotOk) {
        try {
            $screen = [System.Windows.Forms.Screen]::PrimaryScreen
            if ($screen -and $screen.Bounds.Width -gt 0) {
                $w = $screen.Bounds.Width
                $h = $screen.Bounds.Height
                $bmp = New-Object System.Drawing.Bitmap($w, $h)
                $g = [System.Drawing.Graphics]::FromImage($bmp)
                $g.CopyFromScreen($screen.Bounds.Location, [System.Drawing.Point]::Empty, $screen.Bounds.Size)
                $bmp.Save($path)
                $g.Dispose(); $bmp.Dispose()
                Write-Host "[$(Get-Date -Format 'HH:mm:ss')] Screenshot $PosNum saved (${w}x${h}, CopyFromScreen)"
                $screenshotOk = $true
            }
        } catch {
            Write-Host "WARNING: CopyFromScreen failed: $_"
        }
    }
    
    if (-not $screenshotOk) {
        Write-Host "ERROR: All screenshot methods failed"
        return
    }
    
    $key = Upload-Image -Path $path -Token $Token
    if ($key) { Send-ToFeishu -ImageKey $key -Token $Token -OpenId $OpenId }
}

# ========== Main Execution ==========
Write-Host "`n============================================================"
Write-Host "Auto Click + Screenshot + Send to Feishu (Headless V2)"
Write-Host "Time: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
Write-Host "============================================================`n"

# Display screen info
if ($useWin32) {
    $size = [Win32Api]::GetScreenSize()
    Write-Host "Screen Size: $($size[0])x$($size[1])"
}

$token = Get-FeishuToken -AppId $AppId -AppSecret $AppSecret
if (-not $token) { Write-Host "ERROR: No token"; exit 1 }
Write-Host "OK - Token obtained`n"

# Position 1
Write-Host "[$(Get-Date -Format 'HH:mm:ss')] Click position 1: ($X1, $Y1)"
Invoke-Click -X $X1 -Y $Y1
Write-Host "  Waiting 4 seconds..."
Start-Sleep -Seconds 4

# Positions 2-5
Write-Host "[$(Get-Date -Format 'HH:mm:ss')] Click position 2: ($X2, $Y2)"
Invoke-Click -X $X2 -Y $Y2
Screenshot-Send -Token $token -OpenId $UserId -PosNum 2

Write-Host "[$(Get-Date -Format 'HH:mm:ss')] Click position 3: ($X3, $Y3)"
Invoke-Click -X $X3 -Y $Y3
Screenshot-Send -Token $token -OpenId $UserId -PosNum 3

Write-Host "[$(Get-Date -Format 'HH:mm:ss')] Click position 4: ($X4, $Y4)"
Invoke-Click -X $X4 -Y $Y4
Screenshot-Send -Token $token -OpenId $UserId -PosNum 4

Write-Host "[$(Get-Date -Format 'HH:mm:ss')] Click position 5: ($X5, $Y5) - NO screenshot"
Invoke-Click -X $X5 -Y $Y5

Write-Host "`n============================================================"
Write-Host "Task completed! (3 screenshots sent)"
Write-Host "============================================================`n"

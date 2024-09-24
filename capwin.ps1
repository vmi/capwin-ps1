$ErrorActionPreference = "Stop"
Set-StrictMode -Version Latest

Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName System.Linq
Add-Type @"
using System;
using System.Text;
using System.Runtime.InteropServices;

public static class Win32 {
    [DllImport("user32.dll")]
    public extern static IntPtr GetForegroundWindow();

    [DllImport("dwmapi.dll")]
    public extern static int DwmGetWindowAttribute(IntPtr hWnd, int dwAttribute, ref RECT rc, int rcSize);

    [DllImport("user32.dll")]
    static extern int GetWindowText(IntPtr hWnd, StringBuilder lpString, int nMaxCount);

    [DllImport("user32.dll")]
    static extern int GetWindowTextLength(IntPtr hWnd);

    [StructLayout(LayoutKind.Sequential)]
    public struct RECT {
        public int Left;
        public int Top;
        public int Right;
        public int Bottom;
    }

    public static RECT GetWindowRect(IntPtr hWnd) {
        var rcWindow = new RECT();
        DwmGetWindowAttribute(hWnd, 9, ref rcWindow, Marshal.SizeOf(rcWindow));
        return rcWindow;
    }

    public static string GetWindowTitle(IntPtr hWnd) {
        var length = GetWindowTextLength(hWnd) + 1;
        var title = new StringBuilder(length);
        GetWindowText(hWnd, title, length);
        return title.ToString();
    }
}
"@

function Capture-ActiveWindow($writeTo) {
    $hwnd = [Win32]::GetForegroundWindow()
    $title = [Win32]::GetWindowTitle($hwnd)
    Write-Host "Target〚$title〛"
    $rect = [Win32]::GetWindowRect($hwnd)
    $bitmap = New-Object System.Drawing.Bitmap ($rect.Right - $rect.Left), ($rect.Bottom - $rect.Top)
    $graphics = [System.Drawing.Graphics]::FromImage($bitmap)
    $graphics.CopyFromScreen($rect.Left, $rect.Top, 0, 0, $bitmap.Size)
    $bitmap.Save($writeTo, [System.Drawing.Imaging.ImageFormat]::Png)
    $graphics.Dispose()
    $bitmap.Dispose()
    return $title
}

if ($args.Length -gt 0) {
    $wait = $args[0]
} else {
    $wait = 5
}

$prevImage = $null
$savePath = Join-Path (Get-Location) "images"
$null = New-Item $savePath -ItemType Directory -ErrorAction SilentlyContinue
Write-Host "* Save to ${savePath}"


while ($true) {
    $now = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    Write-Host -NoNewline "[${now}] Waiting $wait sec..."
    Start-Sleep -Seconds $wait
    $memoryStream = New-Object System.IO.MemoryStream
    $title = Capture-ActiveWindow($memoryStream)
    $curImage = $memoryStream.ToArray()
    $memoryStream.Dispose()
    if ($null -eq $prevImage -or -not [System.Linq.Enumerable]::SequenceEqual($prevImage, $curImage)) {
        $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
        $qTitle = $title -replace '[\s!"&''*/:;<>?\[\\\]^`|~]', "_"
        $suffix = $qTitle.SubString(0, [Math]::Min($qTitle.Length, 20))
        $filename = "${timestamp}-${suffix}.png"
        $path = Join-Path $savePath $filename
        [System.IO.File]::WriteAllBytes($path, $curImage)
        Write-Host "* Saved: ${filename}"
        $prevImage = $curImage
    }
}

[CmdletBinding()]
param(
    [Parameter()]
    [string]$ProjectRoot
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

if (-not ('SessionMinutesUI.DpiBootstrap' -as [type])) {
    Add-Type -TypeDefinition @'
using System;
using System.Runtime.InteropServices;

namespace SessionMinutesUI
{
    public static class DpiBootstrap
    {
        [DllImport("user32.dll", SetLastError = true)]
        private static extern bool SetProcessDpiAwarenessContext(IntPtr value);

        [DllImport("shcore.dll")]
        private static extern int SetProcessDpiAwareness(int value);

        [DllImport("user32.dll", SetLastError = true)]
        private static extern bool SetProcessDPIAware();

        public static string Enable()
        {
            try
            {
                if (SetProcessDpiAwarenessContext(new IntPtr(-4)))
                {
                    return "PerMonitorV2";
                }
                if (Marshal.GetLastWin32Error() == 5)
                {
                    return "AlreadyConfigured";
                }
            }
            catch (EntryPointNotFoundException) { }
            catch (DllNotFoundException) { }

            try
            {
                int result = SetProcessDpiAwareness(2);
                if (result == 0)
                {
                    return "PerMonitor";
                }
                if (result == unchecked((int)0x80070005))
                {
                    return "AlreadyConfigured";
                }
            }
            catch (EntryPointNotFoundException) { }
            catch (DllNotFoundException) { }

            try
            {
                return SetProcessDPIAware() ? "SystemAware" : "Unavailable";
            }
            catch
            {
                return "Unavailable";
            }
        }
    }
}
'@
}

$script:DpiAwarenessMode = [SessionMinutesUI.DpiBootstrap]::Enable()

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

if ($env:SESSION_MINUTES_HOSTED -ne '1') {
    [System.Windows.Forms.Application]::EnableVisualStyles()
    [System.Windows.Forms.Application]::SetCompatibleTextRenderingDefault($false)
}
if ($env:SESSION_MINUTES_HOSTED -ne '1') {
    [System.Windows.Forms.Application]::SetUnhandledExceptionMode(
        [System.Windows.Forms.UnhandledExceptionMode]::ThrowException
    )
}

if (-not ('SessionMinutesUI.RoundedPanel' -as [type])) {
    Add-Type -ReferencedAssemblies @('System.Windows.Forms', 'System.Drawing') -TypeDefinition @'
using System;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Windows.Forms;

namespace SessionMinutesUI
{
    internal static class RoundedGeometry
    {
        internal static GraphicsPath CreatePath(Rectangle bounds, int radius)
        {
            GraphicsPath path = new GraphicsPath();
            int diameter = Math.Min(radius * 2, Math.Min(bounds.Width, bounds.Height));
            if (diameter <= 2)
            {
                path.AddRectangle(bounds);
                return path;
            }

            Rectangle arc = new Rectangle(bounds.X, bounds.Y, diameter, diameter);
            path.AddArc(arc, 180, 90);
            arc.X = bounds.Right - diameter;
            path.AddArc(arc, 270, 90);
            arc.Y = bounds.Bottom - diameter;
            path.AddArc(arc, 0, 90);
            arc.X = bounds.X;
            path.AddArc(arc, 90, 90);
            path.CloseFigure();
            return path;
        }
    }

    public class RoundedPanel : Panel
    {
        public int CornerRadius { get; set; }
        public int BorderSize { get; set; }
        public Color BorderColor { get; set; }

        public RoundedPanel()
        {
            CornerRadius = 14;
            BorderSize = 1;
            BorderColor = Color.Transparent;
            SetStyle(ControlStyles.UserPaint | ControlStyles.AllPaintingInWmPaint |
                ControlStyles.OptimizedDoubleBuffer | ControlStyles.ResizeRedraw, true);
        }

        protected override void OnResize(EventArgs eventargs)
        {
            base.OnResize(eventargs);
            if (Width > 1 && Height > 1)
            {
                using (GraphicsPath path = RoundedGeometry.CreatePath(ClientRectangle, CornerRadius))
                {
                    Region = new Region(path);
                }
            }
        }

        protected override void OnPaintBackground(PaintEventArgs e)
        {
            e.Graphics.SmoothingMode = SmoothingMode.AntiAlias;
            using (GraphicsPath path = RoundedGeometry.CreatePath(ClientRectangle, CornerRadius))
            using (SolidBrush brush = new SolidBrush(BackColor))
            {
                e.Graphics.FillPath(brush, path);
            }
        }

        protected override void OnPaint(PaintEventArgs e)
        {
            base.OnPaint(e);
            if (BorderSize <= 0 || BorderColor == Color.Transparent)
            {
                return;
            }

            Rectangle borderBounds = ClientRectangle;
            borderBounds.Width -= 1;
            borderBounds.Height -= 1;
            e.Graphics.SmoothingMode = SmoothingMode.AntiAlias;
            using (GraphicsPath path = RoundedGeometry.CreatePath(borderBounds, CornerRadius))
            using (Pen pen = new Pen(BorderColor, BorderSize))
            {
                e.Graphics.DrawPath(pen, path);
            }
        }
    }

    public class RoundedButton : Button
    {
        private bool hovering;
        private bool pressing;

        public int CornerRadius { get; set; }
        public int BorderSize { get; set; }
        public Color BorderColor { get; set; }
        public Color HoverBackColor { get; set; }
        public Color PressedBackColor { get; set; }

        public RoundedButton()
        {
            CornerRadius = 8;
            BorderSize = 1;
            BorderColor = Color.Transparent;
            HoverBackColor = Color.Empty;
            PressedBackColor = Color.Empty;
            FlatStyle = FlatStyle.Flat;
            FlatAppearance.BorderSize = 0;
            UseVisualStyleBackColor = false;
            SetStyle(ControlStyles.UserPaint | ControlStyles.AllPaintingInWmPaint |
                ControlStyles.OptimizedDoubleBuffer | ControlStyles.ResizeRedraw, true);
        }

        protected override void OnResize(EventArgs e)
        {
            base.OnResize(e);
            if (Width > 1 && Height > 1)
            {
                using (GraphicsPath path = RoundedGeometry.CreatePath(ClientRectangle, CornerRadius))
                {
                    Region = new Region(path);
                }
            }
        }

        protected override void OnMouseEnter(EventArgs e)
        {
            hovering = true;
            Invalidate();
            base.OnMouseEnter(e);
        }

        protected override void OnMouseLeave(EventArgs e)
        {
            hovering = false;
            pressing = false;
            Invalidate();
            base.OnMouseLeave(e);
        }

        protected override void OnMouseDown(MouseEventArgs e)
        {
            pressing = true;
            Invalidate();
            base.OnMouseDown(e);
        }

        protected override void OnMouseUp(MouseEventArgs e)
        {
            pressing = false;
            Invalidate();
            base.OnMouseUp(e);
        }

        protected override void OnEnabledChanged(EventArgs e)
        {
            base.OnEnabledChanged(e);
            Invalidate();
        }

        protected override void OnPaint(PaintEventArgs e)
        {
            e.Graphics.SmoothingMode = SmoothingMode.AntiAlias;
            Color fill = BackColor;
            if (pressing && !PressedBackColor.IsEmpty)
            {
                fill = PressedBackColor;
            }
            else if (hovering && Enabled && !HoverBackColor.IsEmpty)
            {
                fill = HoverBackColor;
            }

            Rectangle bounds = ClientRectangle;
            bounds.Width -= 1;
            bounds.Height -= 1;
            using (GraphicsPath path = RoundedGeometry.CreatePath(bounds, CornerRadius))
            using (SolidBrush brush = new SolidBrush(fill))
            {
                e.Graphics.FillPath(brush, path);
                if (BorderSize > 0 && BorderColor != Color.Transparent)
                {
                    using (Pen pen = new Pen(BorderColor, BorderSize))
                    {
                        e.Graphics.DrawPath(pen, path);
                    }
                }
            }

            TextFormatFlags flags = TextFormatFlags.HorizontalCenter |
                TextFormatFlags.VerticalCenter | TextFormatFlags.EndEllipsis |
                TextFormatFlags.NoPrefix;
            if (RightToLeft == RightToLeft.Yes)
            {
                flags |= TextFormatFlags.RightToLeft;
            }
            TextRenderer.DrawText(e.Graphics, Text, Font, ClientRectangle, ForeColor, flags);

            if (Focused && ShowFocusCues)
            {
                Rectangle focus = ClientRectangle;
                focus.Inflate(-4, -4);
                ControlPaint.DrawFocusRectangle(e.Graphics, focus, ForeColor, fill);
            }
        }
    }

    public class SmoothProgressBar : Control
    {
        private readonly Timer timer;
        private int offset;
        private ProgressBarStyle style;

        public ProgressBarStyle Style
        {
            get { return style; }
            set { style = value; timer.Enabled = value == ProgressBarStyle.Marquee && Visible; Invalidate(); }
        }

        public int MarqueeAnimationSpeed
        {
            get { return timer.Interval; }
            set { timer.Interval = Math.Max(15, value); }
        }

        public Color TrackColor { get; set; }
        public Color FillColor { get; set; }

        public SmoothProgressBar()
        {
            TrackColor = Color.FromArgb(228, 225, 238);
            FillColor = Color.FromArgb(79, 70, 229);
            style = ProgressBarStyle.Marquee;
            timer = new Timer { Interval = 24 };
            timer.Tick += delegate { offset = (offset + 7) % Math.Max(1, Width + 80); Invalidate(); };
            SetStyle(ControlStyles.UserPaint | ControlStyles.AllPaintingInWmPaint |
                ControlStyles.OptimizedDoubleBuffer | ControlStyles.ResizeRedraw, true);
        }

        protected override void OnVisibleChanged(EventArgs e)
        {
            base.OnVisibleChanged(e);
            timer.Enabled = Visible && style == ProgressBarStyle.Marquee;
        }

        protected override void OnPaint(PaintEventArgs e)
        {
            e.Graphics.SmoothingMode = SmoothingMode.AntiAlias;
            Rectangle bounds = ClientRectangle;
            if (bounds.Width <= 1 || bounds.Height <= 1) return;
            int radius = Math.Max(2, bounds.Height / 2);
            using (GraphicsPath track = RoundedGeometry.CreatePath(bounds, radius))
            using (SolidBrush brush = new SolidBrush(TrackColor))
            {
                e.Graphics.FillPath(brush, track);
            }
            int chunkWidth = Math.Max(48, bounds.Width / 3);
            int x = offset - chunkWidth;
            Rectangle chunk = new Rectangle(x, 0, chunkWidth, bounds.Height);
            e.Graphics.SetClip(bounds);
            using (GraphicsPath fill = RoundedGeometry.CreatePath(chunk, radius))
            using (SolidBrush brush = new SolidBrush(FillColor))
            {
                e.Graphics.FillPath(brush, fill);
            }
            e.Graphics.ResetClip();
        }
    }
}
'@
}

$script:DisplayDpi = 96.0
try {
    $dpiGraphics = [System.Drawing.Graphics]::FromHwnd([IntPtr]::Zero)
    try {
        $script:DisplayDpi = [Math]::Max(96.0, [double]$dpiGraphics.DpiX)
    }
    finally {
        $dpiGraphics.Dispose()
    }
}
catch {
    $script:DisplayDpi = 96.0
}
$script:FontDpiScale = 1.0

$script:Colors = @{
    Canvas          = [System.Drawing.Color]::FromArgb(252, 248, 255)
    Surface         = [System.Drawing.Color]::FromArgb(255, 255, 255)
    SurfaceMuted    = [System.Drawing.Color]::FromArgb(245, 242, 255)
    SurfaceHigh     = [System.Drawing.Color]::FromArgb(248, 247, 255)
    SurfaceHighest  = [System.Drawing.Color]::FromArgb(240, 236, 249)
    Navy            = [System.Drawing.Color]::FromArgb(53, 37, 205)
    NavySoft        = [System.Drawing.Color]::FromArgb(238, 235, 255)
    Teal            = [System.Drawing.Color]::FromArgb(53, 37, 205)
    TealSoft        = [System.Drawing.Color]::FromArgb(238, 235, 255)
    Amber           = [System.Drawing.Color]::FromArgb(180, 83, 9)
    AmberSoft       = [System.Drawing.Color]::FromArgb(255, 247, 232)
    Red             = [System.Drawing.Color]::FromArgb(180, 35, 24)
    RedSoft         = [System.Drawing.Color]::FromArgb(255, 240, 238)
    Green           = [System.Drawing.Color]::FromArgb(4, 120, 87)
    GreenSoft       = [System.Drawing.Color]::FromArgb(236, 253, 245)
    Ink             = [System.Drawing.Color]::FromArgb(27, 27, 36)
    Muted           = [System.Drawing.Color]::FromArgb(100, 116, 139)
    Border          = [System.Drawing.Color]::FromArgb(226, 232, 240)
    Disabled        = [System.Drawing.Color]::FromArgb(240, 236, 249)
    DisabledText    = [System.Drawing.Color]::FromArgb(148, 163, 184)
    PrimaryText     = [System.Drawing.Color]::FromArgb(255, 255, 255)
    SurfaceLowest   = [System.Drawing.Color]::FromArgb(255, 255, 255)
}
$script:Colors['Background']   = $script:Colors.Canvas
$script:Colors['Primary']      = $script:Colors.Navy
$script:Colors['PrimaryLight'] = $script:Colors.Teal
$script:Colors['MutedText']    = $script:Colors.Muted
$script:Colors['Success']      = $script:Colors.Green
$script:Colors['Warning']      = $script:Colors.Amber

function New-AppFont {
    param(
        [float]$Size = 10,
        [System.Drawing.FontStyle]$Style = [System.Drawing.FontStyle]::Regular
    )
    $scaledSize = [Math]::Max(5.0, $Size * $script:FontDpiScale)
    return [System.Drawing.Font]::new('Segoe UI', $scaledSize, $Style, [System.Drawing.GraphicsUnit]::Point)
}

$script:ToolDirectory = $PSScriptRoot
$script:EnginePath = [System.IO.Path]::GetFullPath(
    (Join-Path $script:ToolDirectory '..\engine\SessionMinutesEngine.exe')
)
$script:AppIconPath = @(
    (Join-Path $script:ToolDirectory 'session_minutes.ico'),
    (Join-Path $script:ToolDirectory '..\assets\session_minutes.ico'),
    (Join-Path $script:ToolDirectory '..\..\..\assets\session_minutes.ico')
) | Where-Object { Test-Path -LiteralPath $_ -PathType Leaf } | Select-Object -First 1
$script:AppIcon = $null
$script:AppLogoPath = @(
    (Join-Path $script:ToolDirectory 'session_minutes_icon_source.png'),
    (Join-Path $script:ToolDirectory '..\assets\session_minutes_icon_source.png'),
    (Join-Path $script:ToolDirectory '..\..\..\assets\session_minutes_icon_source.png')
) | Where-Object { Test-Path -LiteralPath $_ -PathType Leaf } | Select-Object -First 1
$script:AppLogo = $null

if ([string]::IsNullOrWhiteSpace($ProjectRoot)) {
    $ProjectRoot = Join-Path $script:ToolDirectory '..\..\..'
}

try {
    $script:ProjectRoot = (Resolve-Path -LiteralPath $ProjectRoot -ErrorAction Stop).Path.TrimEnd('\')
}
catch {
    [System.Windows.Forms.MessageBox]::Show(
        "تعذر العثور على مجلد المشروع:`r`n$ProjectRoot",
        'تشغيل تعبئة المحاضر',
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Error
    ) | Out-Null
    exit 1
}

$script:OutputsRoot = Join-Path $script:ProjectRoot 'outputs'
$script:PreviewManifest = Join-Path $script:OutputsRoot 'preview\preview.json'
$script:PreviewReport = Join-Path $script:OutputsRoot 'preview\preview.md'
$script:PreviewId = $null
$script:BatchRoot = $null
$script:TemplatePath = $null
$script:CurrentProcess = $null
$script:StdoutTask = $null
$script:StderrTask = $null
$script:RunningMode = $null
$script:LastResultsManifest = $null
$script:LastOutputRoot = $null
$script:ReviewManifest = $null
$script:ReviewQueue = @()
$script:ReviewIndex = 0
$script:logList = $null

function ConvertTo-BidiSafeToken {
    param(
        [Parameter(Mandatory)]
        [AllowEmptyString()]
        [string]$Value
    )

    $leftToRightIsolate = [char]0x2066
    $popDirectionalIsolate = [char]0x2069
    return "$leftToRightIsolate$Value$popDirectionalIsolate"
}

function Add-LogLine {
    param(
        [Parameter(Mandatory)]
        [string]$Message,

        [ValidateSet('Normal', 'Success', 'Warning', 'Error')]
        [string]$Kind = 'Normal'
    )

    $stamp = Get-Date -Format 'HH:mm:ss'
    $wasFollowingLatest = (
        $script:txtLog.TextLength -eq 0 -or
        $script:txtLog.SelectionStart -ge ($script:txtLog.TextLength - 2)
    )
    $script:txtLog.SelectionStart = $script:txtLog.TextLength
    $script:txtLog.SelectionLength = 0
    $script:txtLog.SelectionColor = switch ($Kind) {
        'Success' { $script:Colors.Green }
        'Warning' { $script:Colors.Amber }
        'Error'   { $script:Colors.Red }
        default   { $script:Colors.Ink }
    }
    $script:txtLog.AppendText("[$stamp]  $Message`r`n")
    if ($wasFollowingLatest) {
        $script:txtLog.SelectionStart = $script:txtLog.TextLength
        $script:txtLog.ScrollToCaret()
    }
    if ($null -ne $script:logList) {
        $statusCaption = switch ($Kind) {
            'Success' { 'ناجح' }
            'Warning' { 'تنبيه' }
            'Error' { 'خطأ' }
            default { 'معلومة' }
        }
        $displayMessage = $Message
        $identifier = ''
        $pathMatch = [regex]::Match($Message, '(?i)(?:[A-Z]:\\|\\\\)[^\r\n]+')
        if ($pathMatch.Success) {
            $identifier = $pathMatch.Value.Trim([char]0x2066, [char]0x2069)
            $displayMessage = $Message.Remove($pathMatch.Index, $pathMatch.Length).Trim()
            $displayMessage = $displayMessage -replace '[:؛]\s*$', ''
        }
        elseif ($Message -match '(?i)([a-f0-9]{16,64})') {
            $identifier = $Matches[1]
            $displayMessage = ($Message -replace [regex]::Escape($identifier), '').Trim()
        }
        if ([string]::IsNullOrWhiteSpace($displayMessage)) {
            $displayMessage = 'تفاصيل تقنية'
        }
        $displayMessage = $displayMessage.Replace([string][char]0x2066, '').Replace([string][char]0x2069, '')
        $wasFollowingLatest = $script:logList.RowCount -eq 0 -or (
            $script:logList.FirstDisplayedScrollingRowIndex -ge 0 -and
            ($script:logList.FirstDisplayedScrollingRowIndex + $script:logList.DisplayedRowCount($false)) -ge $script:logList.RowCount
        )
        $technicalValue = if ([string]::IsNullOrWhiteSpace($identifier)) {
            ''
        }
        else {
            "$( [char]0x202A )$identifier$( [char]0x202C )"
        }
        $rowIndex = $script:logList.Rows.Add($stamp, $displayMessage, $statusCaption, $technicalValue)
        $row = $script:logList.Rows[$rowIndex]
        $row.Cells[0].Style.ForeColor = $script:Colors.Muted
        $row.Cells[0].Style.Alignment = [System.Windows.Forms.DataGridViewContentAlignment]::MiddleCenter
        $row.Cells[1].Style.ForeColor = $script:Colors.Ink
        $row.Cells[1].Style.Alignment = [System.Windows.Forms.DataGridViewContentAlignment]::MiddleRight
        $row.Cells[2].Style.ForeColor = $script:txtLog.SelectionColor
        $row.Cells[2].Style.Alignment = [System.Windows.Forms.DataGridViewContentAlignment]::MiddleCenter
        $row.Cells[3].Style.ForeColor = $script:Colors.Muted
        $row.Cells[3].Style.Alignment = [System.Windows.Forms.DataGridViewContentAlignment]::MiddleLeft
        if ($wasFollowingLatest -and $rowIndex -ge 0 -and $script:logList.IsHandleCreated -and $script:logList.DisplayedRowCount($false) -gt 0) {
            try { $script:logList.FirstDisplayedScrollingRowIndex = $rowIndex } catch { }
        }
    }
}

function Set-Status {
    param(
        [Parameter(Mandatory)]
        [string]$Message,

        [ValidateSet('Normal', 'Loading', 'Success', 'Warning', 'Error', 'Empty')]
        [string]$Kind = 'Normal'
    )

    $script:lblStatus.Text = $Message
    $style = switch ($Kind) {
        'Loading' { @{ Caption = 'قيد التنفيذ'; Symbol = '●'; Fore = $script:Colors.Navy; Back = $script:Colors.NavySoft } }
        'Success' { @{ Caption = 'مكتمل'; Symbol = '✓'; Fore = $script:Colors.Green; Back = $script:Colors.GreenSoft } }
        'Warning' { @{ Caption = 'تنبيه'; Symbol = '!'; Fore = $script:Colors.Amber; Back = $script:Colors.AmberSoft } }
        'Error'   { @{ Caption = 'خطأ'; Symbol = '×'; Fore = $script:Colors.Red; Back = $script:Colors.RedSoft } }
        'Empty'   { @{ Caption = 'بانتظار البدء'; Symbol = '○'; Fore = $script:Colors.Muted; Back = $script:Colors.SurfaceMuted } }
        default   { @{ Caption = 'جاهز'; Symbol = '●'; Fore = $script:Colors.Teal; Back = $script:Colors.TealSoft } }
    }
    $script:lblStatus.ForeColor = $style.Fore
    $script:lblStatusKind.ForeColor = $style.Fore
    $script:lblStatusKind.BackColor = $style.Back
    $script:lblStatusKind.Text = "$($style.Symbol)  $($style.Caption)"
    if ($null -ne $script:statusBadgePanel) {
        $script:statusBadgePanel.BackColor = $style.Back
        $script:statusBadgePanel.BorderColor = $style.Fore
        $script:statusBadgePanel.Invalidate()
    }
    $script:statusPanel.BackColor = $script:Colors.Surface
    if ($null -ne $script:statusLayout) {
        $script:statusLayout.BackColor = $script:Colors.Surface
    }
    if ($script:statusPanel.PSObject.Properties['BorderColor']) {
        $script:statusPanel.BorderColor = $script:Colors.Border
        $script:statusPanel.Invalidate()
    }
}

function Show-OperationError {
    param(
        [Parameter(Mandatory)]
        [string]$Message
    )

    Set-Status -Message $Message -Kind Error
    Add-LogLine -Message "خطأ: $Message" -Kind Error
    [System.Windows.Forms.MessageBox]::Show(
        $Message,
        'تشغيل تعبئة المحاضر',
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Error
    ) | Out-Null
}

function Test-Prerequisites {
    param(
        [switch]$Quiet
    )

    $missing = [System.Collections.Generic.List[string]]::new()
    $script:BatchRoot = $null
    $script:TemplatePath = $null
    if (-not (Test-Path -LiteralPath $script:EnginePath -PathType Leaf)) {
        $missing.Add("محرك النسخة المحمولة:`r`n$($script:EnginePath)")
    }
    $batchFolders = @(
        Get-ChildItem -LiteralPath $script:ProjectRoot -Directory -ErrorAction SilentlyContinue |
            Where-Object { $_.Name -match '^[1-9][0-9]*$' } |
            Sort-Object { [int64]$_.Name }
    )
    if ($batchFolders.Count -eq 0) {
        $missing.Add('مجلد دفعة رقمي رئيسي، مثل 8 أو 17')
    }
    elseif ($batchFolders.Count -gt 1) {
        $missing.Add(
            "يجب وجود مجلد دفعة رقمي واحد فقط؛ الموجود: " +
            (($batchFolders | ForEach-Object Name) -join ', ')
        )
    }
    else {
        $script:BatchRoot = $batchFolders[0].FullName
        $numberedSubfolders = @(
            Get-ChildItem -LiteralPath $script:BatchRoot -Directory -ErrorAction SilentlyContinue |
                Where-Object { $_.Name -match '^[1-9][0-9]*$' } |
                Sort-Object { [int64]$_.Name }
        )
        if ($numberedSubfolders.Count -eq 0) {
            $missing.Add("لا توجد مجلدات فرعية رقمية داخل $($batchFolders[0].Name)")
        }
        else {
            $actualNumbers = @($numberedSubfolders | ForEach-Object { [int]$_.Name })
            $expectedNumbers = @(1..($actualNumbers[-1]))
            $numberingDifferences = @(
                Compare-Object -ReferenceObject $expectedNumbers -DifferenceObject $actualNumbers
            )
            if ($numberingDifferences.Count -gt 0) {
                $missing.Add(
                    "المجلدات الفرعية يجب أن تبدأ من 1 وتكون متتابعة؛ الموجود: " +
                    ($actualNumbers -join ', ')
                )
            }
        }
    }
    $preferredTemplate = Join-Path $script:ProjectRoot 'نموذج التعبئة للمحاضر.docx'
    $templateCandidates = @(
        Get-ChildItem -LiteralPath $script:ProjectRoot -File -Filter 'نموذج التعبئة للمحاضر*.docx' -ErrorAction SilentlyContinue |
            Sort-Object Name
    )
    if (Test-Path -LiteralPath $preferredTemplate -PathType Leaf) {
        $script:TemplatePath = (Resolve-Path -LiteralPath $preferredTemplate).Path
    }
    elseif ($templateCandidates.Count -eq 1) {
        $script:TemplatePath = $templateCandidates[0].FullName
    }
    elseif ($templateCandidates.Count -eq 0) {
        $missing.Add("ملف DOCX يبدأ اسمه بـ نموذج التعبئة للمحاضر")
    }
    else {
        $missing.Add(
            "يوجد أكثر من نموذج تعبئة دون النموذج الافتراضي؛ احتفظ بنموذج واحد أو سمِّ النموذج المختار نموذج التعبئة للمحاضر.docx. الموجود: " +
            (($templateCandidates | ForEach-Object Name) -join '، ')
        )
    }

    if ($missing.Count -eq 0) {
        return $true
    }

    if (-not $Quiet) {
        Show-OperationError -Message ("المتطلبات التالية غير موجودة:`r`n`r`n- " + ($missing -join "`r`n- "))
    }
    return $false
}

function Get-TopLevelJsonValue {
    param(
        [Parameter(Mandatory)]
        [object]$Object,

        [Parameter(Mandatory)]
        [string[]]$Names
    )

    foreach ($property in $Object.PSObject.Properties) {
        foreach ($name in $Names) {
            if ($property.Name.Equals($name, [System.StringComparison]::OrdinalIgnoreCase)) {
                return $property.Value
            }
        }
    }
    return $null
}

function Read-JsonDocument {
    param(
        [Parameter(Mandatory)]
        [string]$Path
    )

    return (Get-Content -LiteralPath $Path -Raw -Encoding UTF8 | ConvertFrom-Json)
}

function Set-PreviewFromManifest {
    param(
        [switch]$WriteLog
    )

    $script:PreviewId = $null
    $script:txtPreviewId.Text = ''

    if (-not (Test-Path -LiteralPath $script:PreviewManifest -PathType Leaf)) {
        if ($WriteLog) {
            Add-LogLine -Message 'لم يُنشأ ملف preview.json؛ لن تُتاح التجربة أو المعالجة الكاملة.'
        }
        Update-Controls
        return $false
    }

    try {
        $manifest = Read-JsonDocument -Path $script:PreviewManifest
        $previewId = [string](Get-TopLevelJsonValue -Object $manifest -Names @('preview_id'))
        if ([string]::IsNullOrWhiteSpace($previewId)) {
            throw 'لا يحتوي ملف المعاينة على preview_id.'
        }

        $script:PreviewId = $previewId.Trim()
        $script:txtPreviewId.Text = $script:PreviewId
        if ($WriteLog) {
            $safePreviewId = ConvertTo-BidiSafeToken -Value $script:PreviewId
            Add-LogLine -Message "تم تحميل معرّف المعاينة: $safePreviewId" -Kind Success
            $summary = Get-TopLevelJsonValue -Object $manifest -Names @('summary')
            if ($null -ne $summary) {
                Add-LogLine -Message ("ملخص المعاينة: " + ($summary | ConvertTo-Json -Depth 8 -Compress))
            }
        }
        Update-Controls
        return $true
    }
    catch {
        if ($WriteLog) {
            Add-LogLine -Message "تعذر قراءة ملف المعاينة: $($_.Exception.Message)"
        }
        Update-Controls
        return $false
    }
}

function Get-PilotResultsManifestForPreview {
    param(
        [Parameter(Mandatory)]
        [string]$PreviewId
    )

    $prefixLength = [Math]::Min(12, $PreviewId.Length)
    $prefix = $PreviewId.Substring(0, $prefixLength)
    $manifestPath = Join-Path $script:OutputsRoot "pilot\$prefix\results.json"
    if (-not (Test-Path -LiteralPath $manifestPath -PathType Leaf)) {
        return $null
    }

    try {
        $pilot = Read-JsonDocument -Path $manifestPath
        $pilotPreviewId = [string](Get-TopLevelJsonValue -Object $pilot -Names @('preview_id'))
        $pilotMode = [string](Get-TopLevelJsonValue -Object $pilot -Names @('mode'))
        $sourcesUnchanged = Get-TopLevelJsonValue -Object $pilot -Names @('all_sources_unchanged')
        $validationsPassed = Get-TopLevelJsonValue -Object $pilot -Names @('all_validations_passed')
        if ([string]::IsNullOrWhiteSpace($pilotPreviewId)) {
            return $null
        }
        if (-not $pilotPreviewId.Equals($PreviewId, [System.StringComparison]::OrdinalIgnoreCase)) {
            return $null
        }
        if (
            -not $pilotMode.Equals('pilot', [System.StringComparison]::OrdinalIgnoreCase) -or
            $sourcesUnchanged -ne $true -or
            $validationsPassed -ne $true
        ) {
            return $null
        }
        return $manifestPath
    }
    catch {
        return $null
    }
}

function Get-LatestResultsManifest {
    $searchRoots = @(
        (Join-Path $script:OutputsRoot 'final'),
        (Join-Path $script:OutputsRoot 'pilot')
    )

    $manifests = foreach ($root in $searchRoots) {
        if (Test-Path -LiteralPath $root -PathType Container) {
            Get-ChildItem -LiteralPath $root -Recurse -File -Filter 'results.json' -ErrorAction SilentlyContinue
        }
    }

    return ($manifests | Sort-Object -Property LastWriteTimeUtc -Descending | Select-Object -First 1)
}

function Test-PathIsInsideOutputs {
    param(
        [Parameter(Mandatory)]
        [string]$Candidate
    )

    try {
        $outputFull = [System.IO.Path]::GetFullPath($script:OutputsRoot).TrimEnd('\') + '\'
        $candidateFull = [System.IO.Path]::GetFullPath($Candidate).TrimEnd('\') + '\'
        return $candidateFull.StartsWith($outputFull, [System.StringComparison]::OrdinalIgnoreCase)
    }
    catch {
        return $false
    }
}

function Load-LatestResultsContext {
    param(
        [switch]$WriteLog
    )

    $latest = Get-LatestResultsManifest
    if ($null -eq $latest) {
        $script:LastResultsManifest = $null
        $script:LastOutputRoot = $null
        return $false
    }

    try {
        $data = Read-JsonDocument -Path $latest.FullName
        $outputRootValue = [string](Get-TopLevelJsonValue -Object $data -Names @('output_root'))
        if ([string]::IsNullOrWhiteSpace($outputRootValue)) {
            $outputRootValue = $latest.Directory.FullName
        }
        elseif (-not [System.IO.Path]::IsPathRooted($outputRootValue)) {
            $outputRootValue = Join-Path $latest.Directory.FullName $outputRootValue
        }

        $outputRootValue = [System.IO.Path]::GetFullPath($outputRootValue)
        if (-not (Test-PathIsInsideOutputs -Candidate $outputRootValue)) {
            throw 'رفضت الواجهة فتح مسار لا يقع داخل مجلد outputs.'
        }
        if (-not (Test-Path -LiteralPath $outputRootValue -PathType Container)) {
            throw "مجلد الناتج المسجل غير موجود: $outputRootValue"
        }

        $manifestChanged = -not [string]::Equals(
            $script:LastResultsManifest,
            $latest.FullName,
            [System.StringComparison]::OrdinalIgnoreCase
        )
        $script:LastResultsManifest = $latest.FullName
        $script:LastOutputRoot = $outputRootValue

        if ($manifestChanged) {
            $script:ReviewManifest = $null
            $script:ReviewQueue = @()
            $script:ReviewIndex = 0
        }
        if ($WriteLog) {
            $safeResultsPath = ConvertTo-BidiSafeToken -Value $latest.FullName
            $safeOutputRoot = ConvertTo-BidiSafeToken -Value $outputRootValue
            Add-LogLine -Message "أحدث تقرير نتائج: $safeResultsPath"
            Add-LogLine -Message "مجلد النسخ الناتجة: $safeOutputRoot"
        }
        return $true
    }
    catch {
        if ($WriteLog) {
            Add-LogLine -Message "تعذر تحميل أحدث نتيجة: $($_.Exception.Message)"
        }
        $script:LastResultsManifest = $null
        $script:LastOutputRoot = $null
        return $false
    }
}

function Set-ActionButtonAppearance {
    param(
        [Parameter(Mandatory)]
        [System.Windows.Forms.Button]$Button
    )

    $backColor = $script:Colors.Surface
    $foreColor = $script:Colors.Ink
    $borderColor = $script:Colors.Border
    $hoverColor = $script:Colors.SurfaceHighest
    $pressedColor = $script:Colors.SurfaceMuted

    if (-not $Button.Enabled) {
        $backColor = $script:Colors.Disabled
        $foreColor = $script:Colors.DisabledText
        $borderColor = $script:Colors.Border
        $hoverColor = $script:Colors.Disabled
        $pressedColor = $script:Colors.Disabled
        $Button.Cursor = [System.Windows.Forms.Cursors]::Default
    }
    else {
        $Button.Cursor = [System.Windows.Forms.Cursors]::Hand
        switch ([string]$Button.Tag) {
        'Primary' {
            $backColor = $script:Colors.Navy
            $foreColor = $script:Colors.PrimaryText
            $borderColor = $script:Colors.Navy
            $hoverColor = [System.Drawing.Color]::FromArgb(116, 117, 255)
            $pressedColor = [System.Drawing.Color]::FromArgb(72, 73, 218)
        }
        'Pilot' {
            $backColor = $script:Colors.Navy
            $foreColor = $script:Colors.PrimaryText
            $borderColor = $script:Colors.Navy
            $hoverColor = [System.Drawing.Color]::FromArgb(116, 117, 255)
            $pressedColor = [System.Drawing.Color]::FromArgb(72, 73, 218)
        }
        'Final' {
            $backColor = $script:Colors.Navy
            $foreColor = $script:Colors.PrimaryText
            $borderColor = $script:Colors.Navy
            $hoverColor = [System.Drawing.Color]::FromArgb(79, 70, 229)
            $pressedColor = [System.Drawing.Color]::FromArgb(43, 30, 171)
        }
        'Review' {
            $backColor = $script:Colors.Surface
            $foreColor = $script:Colors.Navy
            $borderColor = $script:Colors.Navy
            $hoverColor = $script:Colors.NavySoft
            $pressedColor = $script:Colors.SurfaceMuted
        }
        default {
            $backColor = $script:Colors.Surface
            $foreColor = $script:Colors.Ink
            $borderColor = $script:Colors.Border
            $hoverColor = $script:Colors.SurfaceHigh
            $pressedColor = $script:Colors.SurfaceMuted
        }
        }
    }

    $Button.BackColor = $backColor
    $Button.ForeColor = $foreColor
    $Button.FlatAppearance.BorderColor = $borderColor
    $Button.FlatAppearance.MouseOverBackColor = $hoverColor
    $Button.FlatAppearance.MouseDownBackColor = $pressedColor
    if ($Button.PSObject.Properties.Name -contains 'BorderColor') {
        $Button.BorderColor = $borderColor
        $Button.BorderSize = 1
        $Button.HoverBackColor = $hoverColor
        $Button.PressedBackColor = $pressedColor
        $Button.Invalidate()
    }
}

function Set-StepCardAppearance {
    param(
        [Parameter(Mandatory)]
        [object]$Card,

        [Parameter(Mandatory)]
        [System.Windows.Forms.Label]$StateLabel,

        [Parameter(Mandatory)]
        [ValidateSet('Active', 'Completed', 'Ready', 'Locked')]
        [string]$State,

        [Parameter(Mandatory)]
        [string]$StepNumber
    )

    switch ($State) {
        'Active' {
            $Card.BackColor = $script:Colors.SurfaceHigh
            $Card.BorderColor = $script:Colors.Navy
            $Card.BorderSize = 2
            $StateLabel.ForeColor = $script:Colors.Navy
            $StateLabel.Text = "$StepNumber • جاري العمل"
        }
        'Completed' {
            $Card.BackColor = $script:Colors.Surface
            $Card.BorderColor = $script:Colors.Border
            $Card.BorderSize = 1
            $StateLabel.ForeColor = $script:Colors.Green
            $StateLabel.Text = "$StepNumber • مكتملة"
        }
        'Ready' {
            $Card.BackColor = $script:Colors.Surface
            $Card.BorderColor = $script:Colors.Border
            $Card.BorderSize = 1
            $StateLabel.ForeColor = $script:Colors.Navy
            $StateLabel.Text = "$StepNumber • جاهزة"
        }
        default {
            $Card.BackColor = $script:Colors.Surface
            $Card.BorderColor = $script:Colors.Border
            $Card.BorderSize = 1
            $StateLabel.ForeColor = $script:Colors.DisabledText
            $StateLabel.Text = "$StepNumber • بانتظار السابقة"
        }
    }
    foreach ($child in $Card.Controls) {
        if ($child -is [System.Windows.Forms.TableLayoutPanel]) {
            $child.BackColor = $Card.BackColor
            foreach ($grandChild in $child.Controls) {
                if ($grandChild -is [System.Windows.Forms.TableLayoutPanel]) {
                    $grandChild.BackColor = $Card.BackColor
                }
            }
        }
    }
    if ($StateLabel.Tag -is [SessionMinutesUI.RoundedPanel]) {
        $pill = [SessionMinutesUI.RoundedPanel]$StateLabel.Tag
        $pill.BackColor = switch ($State) {
            'Active' { $script:Colors.NavySoft }
            'Completed' { $script:Colors.GreenSoft }
            'Ready' { $script:Colors.SurfaceHighest }
            default { $script:Colors.SurfaceMuted }
        }
        $pill.BorderColor = $StateLabel.ForeColor
        $pill.Invalidate()
    }
    if ($null -ne $Card.Tag) {
        $visuals = $Card.Tag
        $visuals.IconTile.BackColor = if ($State -eq 'Active' -or $State -eq 'Ready') { $script:Colors.NavySoft } else { $script:Colors.SurfaceHighest }
        $visuals.Icon.ForeColor = switch ($State) {
            'Active' { $script:Colors.Navy }
            'Ready' { $script:Colors.Navy }
            'Completed' { $script:Colors.Muted }
            default { $script:Colors.DisabledText }
        }
        $visuals.Progress.Visible = $State -eq 'Active'
        $visuals.IconTile.Invalidate()
    }
    $Card.Invalidate()
}

function Update-ButtonVisuals {
    foreach ($button in @(
        $script:btnPreview,
        $script:btnPreviewReport,
        $script:btnPilot,
        $script:btnApply,
        $script:btnOpenOutput,
        $script:btnReviewNext
    )) {
        Set-ActionButtonAppearance -Button $button
    }
}

function Update-Controls {
    $busy = $null -ne $script:CurrentProcess
    $ready = Test-Prerequisites -Quiet
    $hasPreview = $ready -and
        -not [string]::IsNullOrWhiteSpace([string]$script:PreviewId) -and
        (Test-Path -LiteralPath $script:PreviewManifest -PathType Leaf)
    $pilotManifest = $null
    if ($hasPreview) {
        $pilotManifest = Get-PilotResultsManifestForPreview -PreviewId $script:PreviewId
    }

    $script:btnPreview.Enabled = $ready -and -not $busy
    $script:btnPreviewReport.Enabled = -not $busy -and (Test-Path -LiteralPath $script:PreviewReport -PathType Leaf)
    $script:btnPilot.Enabled = $hasPreview -and -not $busy
    $script:btnApply.Enabled = $hasPreview -and -not $busy

    $hasResults = Load-LatestResultsContext
    $script:btnOpenOutput.Enabled = -not $busy -and $hasResults
    $script:btnReviewNext.Enabled = -not $busy -and $hasResults

    $script:lblWorkflowState.Text = if ($busy) {
        'هناك عملية قيد التنفيذ؛ ستُتاح الإجراءات التالية عند اكتمالها.'
    }
    elseif ($script:btnApply.Enabled) {
        'اكتملت المعاينة. اختر نسخة تجريبية للمراجعة أو انتقل مباشرة إلى النسخ النهائية.'
    }
    elseif ($script:btnPilot.Enabled) {
        'المعاينة جاهزة. يمكنك إنشاء ملفات تجريبية للمراجعة.'
    }
    elseif ($hasPreview) {
        'راجع تقرير المعاينة وتأكد من صحة البيانات وسلامة الملفات قبل المتابعة.'
    }
    elseif ($ready) {
        'ابدأ بفحص الملفات وإنشاء معاينة آمنة دون تعديل المصادر.'
    }
    else {
        'أكمل المتطلبات الظاهرة في السجل قبل بدء سير العمل.'
    }

    $previewState = if ($busy -and $script:RunningMode -eq 'preview') {
        'Active'
    }
    elseif ($hasPreview) {
        'Completed'
    }
    elseif ($ready) {
        'Ready'
    }
    else {
        'Locked'
    }
    $pilotState = if ($busy -and $script:RunningMode -eq 'pilot') {
        'Active'
    }
    elseif (-not [string]::IsNullOrWhiteSpace([string]$pilotManifest)) {
        'Completed'
    }
    elseif ($hasPreview) {
        'Ready'
    }
    else {
        'Locked'
    }
    $applyState = if ($busy -and $script:RunningMode -eq 'apply') {
        'Active'
    }
    elseif ($script:btnApply.Enabled) {
        'Ready'
    }
    else {
        'Locked'
    }
    Set-StepCardAppearance -Card $script:stepPreviewCard -StateLabel $script:lblPreviewStepState -State $previewState -StepNumber '١'
    Set-StepCardAppearance -Card $script:stepPilotCard -StateLabel $script:lblPilotStepState -State $pilotState -StepNumber '٢'
    Set-StepCardAppearance -Card $script:stepApplyCard -StateLabel $script:lblApplyStepState -State $applyState -StepNumber '٣'
    Update-ButtonVisuals
}

function ConvertTo-NativeArgument {
    param(
        [Parameter(Mandatory)]
        [string]$Value
    )

    if ($Value.Contains('"')) {
        throw 'لا تدعم الواجهة علامة الاقتباس المزدوجة داخل مسار أو وسيطة.'
    }
    return '"' + $Value + '"'
}

function Start-EngineCommand {
    param(
        [Parameter(Mandatory)]
        [ValidateSet('preview', 'pilot', 'apply')]
        [string]$Mode
    )

    if ($null -ne $script:CurrentProcess) {
        return
    }
    if (-not (Test-Prerequisites)) {
        return
    }

    $engineArguments = @(
        $Mode,
        '--project-root',
        $script:ProjectRoot,
        '--template',
        $script:TemplatePath
    )

    if ($Mode -ne 'preview') {
        if ([string]::IsNullOrWhiteSpace([string]$script:PreviewId)) {
            Show-OperationError -Message 'شغّل المعاينة أولًا، ثم راجع تقريرها وتأكد من صحة البيانات وسلامة الملفات.'
            return
        }
        $engineArguments += @('--approve-preview-id', $script:PreviewId)
    }

    if ($Mode -eq 'apply') {
        $engineArguments += '--allow-direct-apply'
    }

    if ($Mode -eq 'preview') {
        $script:PreviewId = $null
        $script:txtPreviewId.Text = ''
    }

    try {
        $startInfo = [System.Diagnostics.ProcessStartInfo]::new()
        $startInfo.FileName = $script:EnginePath
        $startInfo.Arguments = (($engineArguments | ForEach-Object { ConvertTo-NativeArgument -Value $_ }) -join ' ')
        $startInfo.WorkingDirectory = Split-Path -Parent $script:EnginePath
        $startInfo.UseShellExecute = $false
        $startInfo.CreateNoWindow = $true
        $startInfo.RedirectStandardOutput = $true
        $startInfo.RedirectStandardError = $true
        $startInfo.EnvironmentVariables['PYTHONUTF8'] = '1'
        $startInfo.EnvironmentVariables['PYTHONIOENCODING'] = 'utf-8'
        try {
            $startInfo.StandardOutputEncoding = [System.Text.UTF8Encoding]::new($false)
            $startInfo.StandardErrorEncoding = [System.Text.UTF8Encoding]::new($false)
        }
        catch {
            # Windows PowerShell on older .NET versions may not expose these setters.
        }

        $process = [System.Diagnostics.Process]::new()
        $process.StartInfo = $startInfo
        if (-not $process.Start()) {
            throw 'تعذر بدء محرك التعبئة.'
        }

        $script:CurrentProcess = $process
        $script:StdoutTask = $process.StandardOutput.ReadToEndAsync()
        $script:StderrTask = $process.StandardError.ReadToEndAsync()
        $script:RunningMode = $Mode
        $script:progress.Style = [System.Windows.Forms.ProgressBarStyle]::Marquee
        $script:progress.MarqueeAnimationSpeed = 25
        $script:progress.Visible = $true
        $script:processTimer.Start()

        $modeLabel = switch ($Mode) {
            'preview' { 'الفحص والمعاينة' }
            'pilot'   { 'إنشاء الملفات التجريبية' }
            'apply'   { 'إنشاء النسخ النهائية' }
        }
        Add-LogLine -Message "بدأت عملية: $modeLabel"
        Set-Status -Message "جارٍ تنفيذ $modeLabel..." -Kind Loading
        Update-Controls
    }
    catch {
        $script:CurrentProcess = $null
        $script:StdoutTask = $null
        $script:StderrTask = $null
        $script:RunningMode = $null
        $script:progress.Visible = $false
        Update-Controls
        Show-OperationError -Message "تعذر تشغيل المحرك: $($_.Exception.Message)"
    }
}

function Complete-EngineCommand {
    if ($null -eq $script:CurrentProcess) {
        return
    }

    $process = $script:CurrentProcess
    $mode = $script:RunningMode
    $exitCode = $process.ExitCode
    $stdout = ''
    $stderr = ''

    try {
        $stdout = $script:StdoutTask.GetAwaiter().GetResult()
        $stderr = $script:StderrTask.GetAwaiter().GetResult()
    }
    catch {
        $stderr = "$stderr`r`nتعذر جمع سجل العملية: $($_.Exception.Message)"
    }

    $process.Dispose()
    $script:CurrentProcess = $null
    $script:StdoutTask = $null
    $script:StderrTask = $null
    $script:RunningMode = $null
    $script:progress.Visible = $false

    if (-not [string]::IsNullOrWhiteSpace($stdout)) {
        Add-LogLine -Message 'مخرجات المحرك:'
        foreach ($line in ($stdout.TrimEnd() -split "`r?`n")) {
            Add-LogLine -Message $line
        }
    }
    if (-not [string]::IsNullOrWhiteSpace($stderr)) {
        Add-LogLine -Message 'رسائل المحرك:'
        foreach ($line in ($stderr.TrimEnd() -split "`r?`n")) {
            Add-LogLine -Message $line
        }
    }

    if ($exitCode -eq 0) {
        if ($mode -eq 'preview') {
            $loaded = Set-PreviewFromManifest -WriteLog
            if ($loaded) {
                Set-Status -Message 'اكتملت المعاينة. راجع التقرير وتأكد من صحة البيانات وسلامة الملفات قبل المتابعة.' -Kind Success
            }
            else {
                Set-Status -Message 'انتهى المحرك، لكن ملف المعاينة غير صالح أو غير موجود.' -Kind Warning
            }
        }
        else {
            [void](Load-LatestResultsContext -WriteLog)
            $script:ReviewManifest = $null
            $script:ReviewQueue = @()
            $script:ReviewIndex = 0
            if ($mode -eq 'pilot') {
                Set-Status -Message 'اكتملت التجربة. افتح ملفاتها وراجعها في Word قبل التطبيق الكامل.' -Kind Success
            }
            else {
                Set-Status -Message 'اكتمل إنشاء النسخ النهائية. بقيت المراجعة البشرية في Word.' -Kind Success
            }
        }
    }
    else {
        Set-Status -Message "توقفت العملية بخطأ (رمز الخروج $exitCode). راجع السجل." -Kind Error
    }

    Add-LogLine -Message "انتهت العملية برمز خروج: $exitCode"
    Update-Controls
}

function Open-PreviewReport {
    if (-not (Test-Path -LiteralPath $script:PreviewReport -PathType Leaf)) {
        Show-OperationError -Message 'تقرير المعاينة غير موجود. شغّل الفحص والمعاينة أولًا.'
        return
    }

    try {
        Start-Process -FilePath $script:PreviewReport | Out-Null
        $safePreviewReport = ConvertTo-BidiSafeToken -Value $script:PreviewReport
        Add-LogLine -Message "فُتح تقرير المعاينة: $safePreviewReport" -Kind Success
        Set-Status -Message 'راجع كل التحذيرات والإجراءات في تقرير المعاينة.' -Kind Normal
    }
    catch {
        Show-OperationError -Message "تعذر فتح تقرير المعاينة: $($_.Exception.Message)"
    }
}

function Open-LatestOutput {
    if (-not (Load-LatestResultsContext -WriteLog)) {
        Show-OperationError -Message 'لا توجد نتائج تجريبية أو نهائية لفتحها.'
        return
    }

    try {
        Start-Process -FilePath 'explorer.exe' -ArgumentList @($script:LastOutputRoot) | Out-Null
        Set-Status -Message 'فُتح مجلد أحدث نتيجة.' -Kind Success
    }
    catch {
        Show-OperationError -Message "تعذر فتح مجلد النتائج: $($_.Exception.Message)"
    }
}

function Initialize-ReviewQueue {
    if (-not (Load-LatestResultsContext)) {
        return $false
    }

    if (
        $script:ReviewManifest -and
        $script:ReviewManifest.Equals($script:LastResultsManifest, [System.StringComparison]::OrdinalIgnoreCase) -and
        $script:ReviewQueue.Count -gt 0
    ) {
        return $true
    }

    $files = @(
        Get-ChildItem -LiteralPath $script:LastOutputRoot -Recurse -File -Filter '*.docx' -ErrorAction SilentlyContinue |
            Where-Object { -not $_.Name.StartsWith('~$') } |
            Sort-Object -Property FullName
    )

    $script:ReviewManifest = $script:LastResultsManifest
    $script:ReviewQueue = $files
    $script:ReviewIndex = 0
    Add-LogLine -Message "أُعدّت قائمة مراجعة تضم $($files.Count) ملف Word."
    return ($files.Count -gt 0)
}

function Open-NextReviewDocument {
    if (-not (Initialize-ReviewQueue)) {
        Show-OperationError -Message 'لا توجد ملفات Word ناتجة لفتحها للمراجعة.'
        return
    }

    if ($script:ReviewIndex -ge $script:ReviewQueue.Count) {
        $restart = [System.Windows.Forms.MessageBox]::Show(
            'وصلت إلى نهاية قائمة الملفات. هل تريد بدء القائمة من أول ملف؟',
            'اكتملت قائمة المراجعة',
            [System.Windows.Forms.MessageBoxButtons]::YesNo,
            [System.Windows.Forms.MessageBoxIcon]::Information,
            [System.Windows.Forms.MessageBoxDefaultButton]::Button2
        )
        if ($restart -ne [System.Windows.Forms.DialogResult]::Yes) {
            Set-Status -Message 'اكتملت قائمة المراجعة.' -Kind Success
            return
        }
        $script:ReviewIndex = 0
    }

    $document = $script:ReviewQueue[$script:ReviewIndex]
    if (-not (Test-PathIsInsideOutputs -Candidate $document.FullName)) {
        Show-OperationError -Message 'رفضت الواجهة فتح ملف مراجعة خارج مجلد outputs.'
        return
    }

    try {
        Start-Process -FilePath $document.FullName | Out-Null
        $script:ReviewIndex++
        $position = $script:ReviewIndex
        $total = $script:ReviewQueue.Count
        $safeDocumentPath = ConvertTo-BidiSafeToken -Value $document.FullName
        Add-LogLine -Message "فُتح للمراجعة ($position من $total): $safeDocumentPath" -Kind Success
        Set-Status -Message "فُتح الملف $position من $total في Word. راجع القوائم والقيم غير المكتملة." -Kind Success
    }
    catch {
        Show-OperationError -Message "تعذر فتح ملف Word: $($_.Exception.Message)"
    }
}

$script:form = [System.Windows.Forms.Form]::new()
$script:form.Text = 'محاضر جلسات اللجنة - النسخة المحمولة'
$script:form.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen
$script:form.AutoScaleMode = [System.Windows.Forms.AutoScaleMode]::Dpi
$workingArea = [System.Windows.Forms.Screen]::PrimaryScreen.WorkingArea
$usableWidth = [Math]::Max(560, $workingArea.Width - 24)
$usableHeight = [Math]::Max(500, $workingArea.Height - 48)
$targetWidth = [Math]::Min($usableWidth, [Math]::Max(960, [int][Math]::Floor($workingArea.Width * 0.90)))
$targetHeight = [Math]::Min($usableHeight, [Math]::Max(680, [int][Math]::Floor($workingArea.Height * 0.90)))
$minimumWidth = [Math]::Min($targetWidth, 980)
$minimumHeight = [Math]::Min($targetHeight, 700)
$script:form.MinimumSize = [System.Drawing.Size]::new($minimumWidth, $minimumHeight)
$script:form.ClientSize = [System.Drawing.Size]::new($targetWidth, $targetHeight)
$script:form.BackColor = $script:Colors.Background
$script:form.ForeColor = $script:Colors.Ink
$script:form.Font = New-AppFont -Size 10
$script:form.RightToLeft = [System.Windows.Forms.RightToLeft]::Yes
$script:form.RightToLeftLayout = $false
$script:form.KeyPreview = $true
if (-not [string]::IsNullOrWhiteSpace([string]$script:AppIconPath)) {
    try {
        $script:AppIcon = [System.Drawing.Icon]::new($script:AppIconPath)
        $script:form.Icon = $script:AppIcon
    }
    catch {
        $script:AppIcon = $null
    }
}
if (-not [string]::IsNullOrWhiteSpace([string]$script:AppLogoPath)) {
    try {
        $script:AppLogo = [System.Drawing.Bitmap]::new($script:AppLogoPath)
    }
    catch {
        $script:AppLogo = $null
    }
}

function New-UiLabel {
    param([string]$Text, [float]$Size = 9, [System.Drawing.FontStyle]$Style = [System.Drawing.FontStyle]::Regular,
          [System.Drawing.Color]$Color = $script:Colors.Ink,
          [System.Drawing.ContentAlignment]$Align = [System.Drawing.ContentAlignment]::MiddleRight)
    $label = [System.Windows.Forms.Label]::new()
    $label.Text = $Text
    $label.Dock = [System.Windows.Forms.DockStyle]::Fill
    $label.TextAlign = $Align
    $label.Font = New-AppFont -Size $Size -Style $Style
    $label.ForeColor = $Color
    $label.RightToLeft = [System.Windows.Forms.RightToLeft]::Yes
    $label.AutoEllipsis = $true
    return $label
}

function New-UiPanel {
    param([int]$Radius = 14, [System.Drawing.Color]$Back = $script:Colors.Surface)
    $panel = [SessionMinutesUI.RoundedPanel]::new()
    $panel.Dock = [System.Windows.Forms.DockStyle]::Fill
    $panel.BackColor = $Back
    $panel.BorderColor = $script:Colors.Border
    $panel.BorderSize = 1
    $panel.CornerRadius = $Radius
    return $panel
}

function New-WorkflowCard {
    param(
        [string]$Number,
        [string]$Title,
        [string]$Description,
        [string]$PrimaryText,
        [string]$SecondaryText,
        [int]$TabIndex,
        [int]$IconCode,
        [string]$PrimaryTag
    )

    $card = New-UiPanel -Radius 8
    $card.Margin = [System.Windows.Forms.Padding]::new(9, 7, 9, 7)
    $card.Padding = [System.Windows.Forms.Padding]::new(22, 17, 22, 17)

    $layout = [System.Windows.Forms.TableLayoutPanel]::new()
    $layout.Dock = [System.Windows.Forms.DockStyle]::Fill
    $layout.BackColor = $card.BackColor
    $layout.ColumnCount = 2
    $layout.RowCount = 5
    $layout.RightToLeft = [System.Windows.Forms.RightToLeft]::Yes
    [void]$layout.ColumnStyles.Add([System.Windows.Forms.ColumnStyle]::new([System.Windows.Forms.SizeType]::Absolute, 66))
    [void]$layout.ColumnStyles.Add([System.Windows.Forms.ColumnStyle]::new([System.Windows.Forms.SizeType]::Percent, 100))
    [void]$layout.RowStyles.Add([System.Windows.Forms.RowStyle]::new([System.Windows.Forms.SizeType]::Absolute, 44))
    [void]$layout.RowStyles.Add([System.Windows.Forms.RowStyle]::new([System.Windows.Forms.SizeType]::Absolute, 32))
    [void]$layout.RowStyles.Add([System.Windows.Forms.RowStyle]::new([System.Windows.Forms.SizeType]::Percent, 100))
    [void]$layout.RowStyles.Add([System.Windows.Forms.RowStyle]::new([System.Windows.Forms.SizeType]::Absolute, 10))
    [void]$layout.RowStyles.Add([System.Windows.Forms.RowStyle]::new([System.Windows.Forms.SizeType]::Absolute, 44))
    [void]$card.Controls.Add($layout)

    $iconTile = New-UiPanel -Radius 11 -Back $script:Colors.SurfaceHighest
    $iconTile.Dock = [System.Windows.Forms.DockStyle]::Fill
    $iconTile.Margin = [System.Windows.Forms.Padding]::new(13, 0, 0, 0)
    $iconTile.BorderSize = 0
    $icon = New-UiLabel -Text ([string][char]$IconCode) -Size 20 -Color $script:Colors.PrimaryLight -Align ([System.Drawing.ContentAlignment]::MiddleCenter)
    $icon.Font = [System.Drawing.Font]::new('Segoe MDL2 Assets', [Math]::Max(9.0, 20 * $script:FontDpiScale), [System.Drawing.FontStyle]::Regular)
    $icon.RightToLeft = [System.Windows.Forms.RightToLeft]::No
    [void]$iconTile.Controls.Add($icon)
    [void]$layout.Controls.Add($iconTile, 0, 0)

    $statePill = New-UiPanel -Radius 14 -Back $script:Colors.SurfaceHighest
    $statePill.Dock = [System.Windows.Forms.DockStyle]::None
    $statePill.Size = [System.Drawing.Size]::new(118, 30)
    # TableLayoutPanel mirrors anchor semantics when RightToLeft is enabled.
    $statePill.Anchor = [System.Windows.Forms.AnchorStyles]::Right
    $statePill.Margin = [System.Windows.Forms.Padding]::new(0, 7, 0, 7)
    $statePill.BorderSize = 1
    $statePill.BorderColor = $script:Colors.Border
    $stateLabel = New-UiLabel -Text "المرحلة $Number" -Size 8 -Style ([System.Drawing.FontStyle]::Bold) -Color $script:Colors.MutedText -Align ([System.Drawing.ContentAlignment]::MiddleCenter)
    $stateLabel.Tag = $statePill
    [void]$statePill.Controls.Add($stateLabel)
    [void]$layout.Controls.Add($statePill, 1, 0)
    $layout.Tag = [pscustomobject]@{ IconTile = $iconTile; StatePill = $statePill }

    $titleLabel = New-UiLabel -Text $Title -Size 13 -Style ([System.Drawing.FontStyle]::Bold) -Align ([System.Drawing.ContentAlignment]::BottomLeft)
    $layout.SetColumnSpan($titleLabel, 2)
    [void]$layout.Controls.Add($titleLabel, 0, 1)

    $descriptionLabel = New-UiLabel -Text $Description -Size 9 -Color $script:Colors.MutedText -Align ([System.Drawing.ContentAlignment]::TopLeft)
    $descriptionLabel.AutoEllipsis = $false
    $descriptionLabel.Padding = [System.Windows.Forms.Padding]::new(0, 8, 0, 4)
    $layout.SetColumnSpan($descriptionLabel, 2)
    [void]$layout.Controls.Add($descriptionLabel, 0, 2)

    $cardProgress = [SessionMinutesUI.SmoothProgressBar]::new()
    $cardProgress.Dock = [System.Windows.Forms.DockStyle]::Fill
    $cardProgress.Margin = [System.Windows.Forms.Padding]::new(0, 3, 0, 4)
    $cardProgress.Style = [System.Windows.Forms.ProgressBarStyle]::Marquee
    $cardProgress.MarqueeAnimationSpeed = 24
    $cardProgress.Visible = $false
    $layout.SetColumnSpan($cardProgress, 2)
    [void]$layout.Controls.Add($cardProgress, 0, 3)

    $actions = [System.Windows.Forms.TableLayoutPanel]::new()
    $actions.Dock = [System.Windows.Forms.DockStyle]::Fill
    $actions.BackColor = $card.BackColor
    $actions.ColumnCount = 2
    $actions.RightToLeft = [System.Windows.Forms.RightToLeft]::Yes
    [void]$actions.ColumnStyles.Add([System.Windows.Forms.ColumnStyle]::new([System.Windows.Forms.SizeType]::Percent, 58))
    [void]$actions.ColumnStyles.Add([System.Windows.Forms.ColumnStyle]::new([System.Windows.Forms.SizeType]::Percent, 42))
    $layout.SetColumnSpan($actions, 2)
    [void]$layout.Controls.Add($actions, 0, 4)

    $primary = [SessionMinutesUI.RoundedButton]::new()
    $primary.Text = $PrimaryText
    $primary.Dock = [System.Windows.Forms.DockStyle]::Fill
    $primary.Margin = [System.Windows.Forms.Padding]::new(0, 0, 6, 0)
    $primary.Font = New-AppFont -Size 9 -Style ([System.Drawing.FontStyle]::Bold)
    $primary.TabIndex = $TabIndex
    $primary.Tag = $PrimaryTag
    $primary.AccessibleName = $PrimaryText
    [void]$actions.Controls.Add($primary, 0, 0)

    $secondary = [SessionMinutesUI.RoundedButton]::new()
    $secondary.Text = $SecondaryText
    $secondary.Dock = [System.Windows.Forms.DockStyle]::Fill
    $secondary.Margin = [System.Windows.Forms.Padding]::new(6, 0, 0, 0)
    $secondary.Font = New-AppFont -Size 8
    $secondary.TabIndex = $TabIndex + 1
    $secondary.Tag = 'Review'
    $secondary.AccessibleName = $SecondaryText
    [void]$actions.Controls.Add($secondary, 1, 0)

    $card.Tag = [pscustomobject]@{
        IconTile = $iconTile
        Icon = $icon
        StatePill = $statePill
        Progress = $cardProgress
    }

    return [pscustomobject]@{
        Card = $card
        Layout = $layout
        State = $stateLabel
        Primary = $primary
        Secondary = $secondary
    }
}
$main = [System.Windows.Forms.TableLayoutPanel]::new()
$main.Dock = [System.Windows.Forms.DockStyle]::Fill
$main.Padding = [System.Windows.Forms.Padding]::new(40, 16, 40, 16)
$main.BackColor = $script:Colors.Background
$main.ColumnCount = 1
$main.RowCount = 5
$main.ColumnStyles.Add([System.Windows.Forms.ColumnStyle]::new([System.Windows.Forms.SizeType]::Percent, 100))
$main.RowStyles.Add([System.Windows.Forms.RowStyle]::new([System.Windows.Forms.SizeType]::Absolute, 58))
$main.RowStyles.Add([System.Windows.Forms.RowStyle]::new([System.Windows.Forms.SizeType]::Absolute, 136))
$main.RowStyles.Add([System.Windows.Forms.RowStyle]::new([System.Windows.Forms.SizeType]::Absolute, 365))
$main.RowStyles.Add([System.Windows.Forms.RowStyle]::new([System.Windows.Forms.SizeType]::Percent, 100))
$main.RowStyles.Add([System.Windows.Forms.RowStyle]::new([System.Windows.Forms.SizeType]::Absolute, 52))
$script:form.Controls.Add($main)

$topNav = New-UiPanel -Radius 12 -Back $script:Colors.Surface
$topNav.Margin = [System.Windows.Forms.Padding]::new(0, 0, 0, 8)
$topNav.Padding = [System.Windows.Forms.Padding]::new(14, 8, 14, 8)
$topNav.BorderColor = $script:Colors.NavySoft
$main.Controls.Add($topNav, 0, 0)

$nav = [System.Windows.Forms.TableLayoutPanel]::new()
$nav.Dock = [System.Windows.Forms.DockStyle]::Fill
$nav.BackColor = $topNav.BackColor
$nav.ColumnCount = 3
$nav.ColumnStyles.Add([System.Windows.Forms.ColumnStyle]::new([System.Windows.Forms.SizeType]::Absolute, 50))
$nav.ColumnStyles.Add([System.Windows.Forms.ColumnStyle]::new([System.Windows.Forms.SizeType]::Absolute, 250))
$nav.ColumnStyles.Add([System.Windows.Forms.ColumnStyle]::new([System.Windows.Forms.SizeType]::Percent, 100))
$topNav.Controls.Add($nav)
$accent = if ($null -ne $script:AppLogo) {
    $picture = [System.Windows.Forms.PictureBox]::new()
    $picture.Dock = [System.Windows.Forms.DockStyle]::Fill
    $picture.Padding = [System.Windows.Forms.Padding]::new(6)
    $picture.BackColor = $topNav.BackColor
    $picture.SizeMode = [System.Windows.Forms.PictureBoxSizeMode]::Zoom
    $picture.Image = $script:AppLogo
    $picture.AccessibleName = 'أيقونة محاضر جلسات اللجنة'
    $picture
}
else {
    $fallback = [System.Windows.Forms.Label]::new()
    $fallback.Dock = [System.Windows.Forms.DockStyle]::Fill
    $fallback.Text = [string][char]0xE713
    $fallback.Font = [System.Drawing.Font]::new('Segoe MDL2 Assets', [Math]::Max(9.0, 18 * $script:FontDpiScale))
    $fallback.ForeColor = $script:Colors.Primary
    $fallback.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
    $fallback.RightToLeft = [System.Windows.Forms.RightToLeft]::No
    $fallback
}
$nav.Controls.Add($accent, 0, 0)
$brand = New-UiLabel -Text 'محاضر جلسات اللجنة' -Size 15 -Style ([System.Drawing.FontStyle]::Bold) -Color $script:Colors.Ink -Align ([System.Drawing.ContentAlignment]::MiddleLeft)
$brand.Padding = [System.Windows.Forms.Padding]::new(0, 0, 12, 0)
$nav.Controls.Add($brand, 1, 0)
$projectLabel = New-UiLabel -Text $script:ProjectRoot -Size 9 -Color $script:Colors.MutedText -Align ([System.Drawing.ContentAlignment]::MiddleLeft)
$projectLabel.RightToLeft = [System.Windows.Forms.RightToLeft]::No
$projectLabel.Font = [System.Drawing.Font]::new('Segoe UI', [Math]::Max(5.0, 9 * $script:FontDpiScale))
$projectLabel.BackColor = $script:Colors.SurfaceHigh
$projectLabel.Padding = [System.Windows.Forms.Padding]::new(14, 0, 14, 0)
$projectLabel.AccessibleName = 'مسار مجلد المشروع'
$nav.Controls.Add($projectLabel, 2, 0)

$hero = New-UiPanel -Radius 14 -Back $script:Colors.SurfaceHigh
$hero.Margin = [System.Windows.Forms.Padding]::new(0, 4, 0, 10)
$hero.Padding = [System.Windows.Forms.Padding]::new(24, 12, 24, 8)
$hero.BorderColor = $script:Colors.Border
$main.Controls.Add($hero, 0, 1)
$heroGrid = [System.Windows.Forms.TableLayoutPanel]::new()
$heroGrid.Dock = [System.Windows.Forms.DockStyle]::Fill
$heroGrid.BackColor = $hero.BackColor
$heroGrid.RightToLeft = [System.Windows.Forms.RightToLeft]::No
$heroGrid.ColumnCount = 1
$heroGrid.RowCount = 3
$heroGrid.ColumnStyles.Add([System.Windows.Forms.ColumnStyle]::new([System.Windows.Forms.SizeType]::Percent, 100))
$heroGrid.RowStyles.Add([System.Windows.Forms.RowStyle]::new([System.Windows.Forms.SizeType]::Absolute, 42))
$heroGrid.RowStyles.Add([System.Windows.Forms.RowStyle]::new([System.Windows.Forms.SizeType]::Absolute, 30))
$heroGrid.RowStyles.Add([System.Windows.Forms.RowStyle]::new([System.Windows.Forms.SizeType]::Absolute, 30))
$hero.Controls.Add($heroGrid)
$heroTitle = New-UiLabel -Text 'إعداد محاضر الجلسات بثقة' -Size 20 -Style ([System.Drawing.FontStyle]::Bold) -Color $script:Colors.Ink -Align ([System.Drawing.ContentAlignment]::BottomLeft)
$heroTitle.RightToLeft = [System.Windows.Forms.RightToLeft]::Yes
$heroGrid.Controls.Add($heroTitle, 0, 0)
$heroSubtitle = New-UiLabel -Text 'افحص ملفات الجلسات، أنشئ نسخ المراجعة، ثم جهّز النسخ النهائية بخطوات واضحة وآمنة.' -Size 10 -Color $script:Colors.MutedText -Align ([System.Drawing.ContentAlignment]::TopLeft)
$heroSubtitle.RightToLeft = [System.Windows.Forms.RightToLeft]::Yes
$heroGrid.Controls.Add($heroSubtitle, 0, 1)
$heroSafety = New-UiLabel -Text 'معالجة محلية  •  حماية الملفات الأصلية  •  حفظ النتائج تلقائيًا داخل outputs' -Size 9 -Style ([System.Drawing.FontStyle]::Bold) -Color $script:Colors.Success -Align ([System.Drawing.ContentAlignment]::MiddleLeft)
$heroSafety.RightToLeft = [System.Windows.Forms.RightToLeft]::Yes
$heroGrid.Controls.Add($heroSafety, 0, 2)

$workflow = [System.Windows.Forms.TableLayoutPanel]::new()
$workflow.Dock = [System.Windows.Forms.DockStyle]::Fill
$workflow.Margin = [System.Windows.Forms.Padding]::new(0, 0, 0, 10)
$workflow.BackColor = $script:Colors.Background
$workflow.ColumnCount = 1
$workflow.RowCount = 4
$workflow.RowStyles.Add([System.Windows.Forms.RowStyle]::new([System.Windows.Forms.SizeType]::Absolute, 42))
$workflow.RowStyles.Add([System.Windows.Forms.RowStyle]::new([System.Windows.Forms.SizeType]::Absolute, 230))
$workflow.RowStyles.Add([System.Windows.Forms.RowStyle]::new([System.Windows.Forms.SizeType]::Absolute, 84))
$workflow.RowStyles.Add([System.Windows.Forms.RowStyle]::new([System.Windows.Forms.SizeType]::Percent, 100))
$main.Controls.Add($workflow, 0, 2)

$workflowHead = [System.Windows.Forms.TableLayoutPanel]::new()
$workflowHead.Dock = [System.Windows.Forms.DockStyle]::Fill
$workflowHead.BackColor = $workflow.BackColor
$workflowHead.RightToLeft = [System.Windows.Forms.RightToLeft]::No
$workflowHead.ColumnCount = 2
$workflowHead.ColumnStyles.Add([System.Windows.Forms.ColumnStyle]::new([System.Windows.Forms.SizeType]::Percent, 40))
$workflowHead.ColumnStyles.Add([System.Windows.Forms.ColumnStyle]::new([System.Windows.Forms.SizeType]::Percent, 60))
$workflow.Controls.Add($workflowHead, 0, 0)
$workflowHead.Controls.Add((New-UiLabel -Text 'الحالة الحالية  •  تتحدد حسب تقدمك' -Style ([System.Drawing.FontStyle]::Bold) -Color $script:Colors.Primary -Align ([System.Drawing.ContentAlignment]::MiddleLeft)), 0, 0)
$workflowHead.Controls.Add((New-UiLabel -Text 'سير العمل اليومي' -Size 14 -Style ([System.Drawing.FontStyle]::Bold) -Align ([System.Drawing.ContentAlignment]::MiddleLeft)), 1, 0)

$cards = [System.Windows.Forms.TableLayoutPanel]::new()
$cards.Dock = [System.Windows.Forms.DockStyle]::Fill
$cards.BackColor = $workflow.BackColor
$cards.ColumnCount = 3
$cards.ColumnStyles.Add([System.Windows.Forms.ColumnStyle]::new([System.Windows.Forms.SizeType]::Percent, 33.333))
$cards.ColumnStyles.Add([System.Windows.Forms.ColumnStyle]::new([System.Windows.Forms.SizeType]::Percent, 33.334))
$cards.ColumnStyles.Add([System.Windows.Forms.ColumnStyle]::new([System.Windows.Forms.SizeType]::Percent, 33.333))
$workflow.Controls.Add($cards, 0, 1)

$previewCard = New-WorkflowCard -Number '01' -Title 'الفحص والمعاينة' -Description 'تحليل أولي للنموذج وملفات المصدر وتجهيز تقرير الإجراءات والتحذيرات.' -PrimaryText 'بدء المعاينة' -SecondaryText 'فتح التقرير' -TabIndex 0 -IconCode 0xE721 -PrimaryTag 'Primary'
$script:stepPreviewCard=$previewCard.Card; $script:lblPreviewStepState=$previewCard.State
$script:btnPreview=$previewCard.Primary; $script:btnPreviewReport=$previewCard.Secondary
$cards.Controls.Add($script:stepPreviewCard, 0, 0)

$pilotCard = New-WorkflowCard -Number '02' -Title 'إنشاء ملفات تجريبية' -Description 'توليد مجموعة أولية آمنة ثم فتح الملفات بالتتابع للمراجعة في Word.' -PrimaryText 'إنشاء التجربة' -SecondaryText 'الملف التالي' -TabIndex 2 -IconCode 0xE70F -PrimaryTag 'Pilot'
$script:stepPilotCard=$pilotCard.Card; $script:lblPilotStepState=$pilotCard.State
$script:btnPilot=$pilotCard.Primary; $script:btnReviewNext=$pilotCard.Secondary
$cards.Controls.Add($script:stepPilotCard, 1, 0)

$applyCard = New-WorkflowCard -Number '03' -Title 'إنشاء النسخ النهائية' -Description 'اعتماد المحاضر بعد المراجعة وحفظها داخل مجلد نتائج مستقل.' -PrimaryText 'إنشاء النهائيات' -SecondaryText 'فتح النتائج' -TabIndex 4 -IconCode 0xE73E -PrimaryTag 'Final'
$script:stepApplyCard=$applyCard.Card; $script:lblApplyStepState=$applyCard.State
$script:btnApply=$applyCard.Primary; $script:btnOpenOutput=$applyCard.Secondary
$cards.Controls.Add($script:stepApplyCard, 2, 0)
$script:workflowCardLayouts = @($previewCard.Layout, $pilotCard.Layout, $applyCard.Layout)

$gate = New-UiPanel -Radius 10 -Back $script:Colors.NavySoft
$gate.Margin = [System.Windows.Forms.Padding]::new(7, 6, 7, 4)
$gate.Padding = [System.Windows.Forms.Padding]::new(20, 8, 20, 8)
$gate.BorderColor = $script:Colors.Navy
$workflow.Controls.Add($gate, 0, 2)
$gateGrid = [System.Windows.Forms.TableLayoutPanel]::new()
$gateGrid.Dock = [System.Windows.Forms.DockStyle]::Fill
$gateGrid.BackColor = $gate.BackColor
$gateGrid.ColumnCount = 1
$gateGrid.RowCount = 2
$gateGrid.ColumnStyles.Add([System.Windows.Forms.ColumnStyle]::new([System.Windows.Forms.SizeType]::Percent, 100))
$gateGrid.RowStyles.Add([System.Windows.Forms.RowStyle]::new([System.Windows.Forms.SizeType]::Absolute, 42))
$gateGrid.RowStyles.Add([System.Windows.Forms.RowStyle]::new([System.Windows.Forms.SizeType]::Percent, 100))
$gate.Controls.Add($gateGrid)
$script:lblPreviewId = New-UiLabel -Text 'لا توجد معاينة حالية' -Size 8 -Color $script:Colors.PrimaryLight
$script:lblPreviewId.Font = [System.Drawing.Font]::new('Consolas', [Math]::Max(5.0, 8.5 * $script:FontDpiScale))
$script:lblPreviewId.RightToLeft = [System.Windows.Forms.RightToLeft]::No
$script:txtPreviewId = $script:lblPreviewId
$script:lblPreviewId.Visible = $false
$script:lblWarningTitle = New-UiLabel -Text 'قبل المتابعة' -Size 10 -Style ([System.Drawing.FontStyle]::Bold) -Color $script:Colors.Navy -Align ([System.Drawing.ContentAlignment]::MiddleLeft)
$script:lblWarningTitle.RightToLeft = [System.Windows.Forms.RightToLeft]::Yes
$script:lblWarningTitle.Padding = [System.Windows.Forms.Padding]::new(0, 3, 0, 3)
$gateGrid.Controls.Add($script:lblWarningTitle, 0, 0)
$script:lblWarning = New-UiLabel -Text 'افتح تقرير المعاينة وراجع البيانات والملفات بعناية، ثم أنشئ الملفات التجريبية أو النسخ النهائية بعد التأكد من صحتها.' -Size 9 -Color $script:Colors.Ink -Align ([System.Drawing.ContentAlignment]::MiddleLeft)
$script:lblWarning.RightToLeft = [System.Windows.Forms.RightToLeft]::Yes
$script:lblWarning.AccessibleName = 'تنبيه مراجعة البيانات والملفات'
$gateGrid.Controls.Add($script:lblWarning, 0, 1)
$script:lblWorkflowState = New-UiLabel -Text 'بعد اكتمال المعاينة، يمكنك إنشاء نسخة تجريبية للمراجعة أو الانتقال مباشرة إلى النسخ النهائية.' -Color $script:Colors.MutedText -Align ([System.Drawing.ContentAlignment]::MiddleRight)
$script:lblWorkflowState.RightToLeft = [System.Windows.Forms.RightToLeft]::Yes
$workflow.Controls.Add($script:lblWorkflowState, 0, 3)

$logPanel = New-UiPanel -Radius 16 -Back $script:Colors.SurfaceLowest
$logPanel.Margin = [System.Windows.Forms.Padding]::new(0, 0, 0, 10)
$logPanel.Padding = [System.Windows.Forms.Padding]::new(16, 12, 16, 14)
$main.Controls.Add($logPanel, 0, 3)
$logGrid = [System.Windows.Forms.TableLayoutPanel]::new()
$logGrid.Dock = [System.Windows.Forms.DockStyle]::Fill
$logGrid.BackColor = $logPanel.BackColor
$logGrid.RowCount = 2
$logGrid.RowStyles.Add([System.Windows.Forms.RowStyle]::new([System.Windows.Forms.SizeType]::Absolute, 42))
$logGrid.RowStyles.Add([System.Windows.Forms.RowStyle]::new([System.Windows.Forms.SizeType]::Percent, 100))
$logPanel.Controls.Add($logGrid)
$logGrid.Controls.Add((New-UiLabel -Text 'سجل النشاط' -Size 11 -Style ([System.Drawing.FontStyle]::Bold) -Align ([System.Drawing.ContentAlignment]::MiddleLeft)), 0, 0)
$script:logBox = [System.Windows.Forms.RichTextBox]::new()
$script:logBox.Dock = [System.Windows.Forms.DockStyle]::Fill
$script:logBox.ReadOnly = $true
$script:logBox.BorderStyle = [System.Windows.Forms.BorderStyle]::None
$script:logBox.BackColor = $script:Colors.SurfaceLowest
$script:logBox.ForeColor = $script:Colors.Ink
$script:logBox.Font = New-AppFont -Size 9
$script:logBox.RightToLeft = [System.Windows.Forms.RightToLeft]::Yes
$script:logBox.WordWrap = $true
$script:logBox.ScrollBars = [System.Windows.Forms.RichTextBoxScrollBars]::Vertical
$script:logBox.TabStop = $false
$script:logBox.AccessibleName = 'سجل النشاط والحالة'
$script:txtLog = $script:logBox
$script:logList = [System.Windows.Forms.DataGridView]::new()
$script:logList.Dock = [System.Windows.Forms.DockStyle]::Fill
$script:logList.ReadOnly = $true
$script:logList.AllowUserToAddRows = $false
$script:logList.AllowUserToDeleteRows = $false
$script:logList.AllowUserToResizeRows = $false
$script:logList.AutoGenerateColumns = $false
$script:logList.AutoSizeRowsMode = [System.Windows.Forms.DataGridViewAutoSizeRowsMode]::AllCellsExceptHeaders
$script:logList.SelectionMode = [System.Windows.Forms.DataGridViewSelectionMode]::FullRowSelect
$script:logList.MultiSelect = $true
$script:logList.ClipboardCopyMode = [System.Windows.Forms.DataGridViewClipboardCopyMode]::EnableAlwaysIncludeHeaderText
$script:logList.RowHeadersVisible = $false
$script:logList.BorderStyle = [System.Windows.Forms.BorderStyle]::None
$script:logList.BackgroundColor = $script:Colors.SurfaceLowest
$script:logList.GridColor = $script:Colors.Border
$script:logList.CellBorderStyle = [System.Windows.Forms.DataGridViewCellBorderStyle]::SingleHorizontal
$script:logList.ColumnHeadersBorderStyle = [System.Windows.Forms.DataGridViewHeaderBorderStyle]::None
$script:logList.EnableHeadersVisualStyles = $false
$script:logList.Font = New-AppFont -Size 8.5
$script:logList.RightToLeft = [System.Windows.Forms.RightToLeft]::Yes
$script:logList.AccessibleName = 'جدول سجل النشاط والحالة'
$script:logList.DefaultCellStyle.BackColor = $script:Colors.SurfaceLowest
$script:logList.DefaultCellStyle.ForeColor = $script:Colors.Ink
$script:logList.DefaultCellStyle.SelectionBackColor = $script:Colors.NavySoft
$script:logList.DefaultCellStyle.SelectionForeColor = $script:Colors.Ink
$script:logList.DefaultCellStyle.Padding = [System.Windows.Forms.Padding]::new(8, 5, 8, 5)
$script:logList.DefaultCellStyle.WrapMode = [System.Windows.Forms.DataGridViewTriState]::True
$script:logList.AlternatingRowsDefaultCellStyle.BackColor = $script:Colors.SurfaceHigh
$script:logList.ColumnHeadersDefaultCellStyle.BackColor = $script:Colors.SurfaceMuted
$script:logList.ColumnHeadersDefaultCellStyle.ForeColor = $script:Colors.Muted
$script:logList.ColumnHeadersDefaultCellStyle.Alignment = [System.Windows.Forms.DataGridViewContentAlignment]::MiddleRight
$script:logList.ColumnHeadersDefaultCellStyle.Padding = [System.Windows.Forms.Padding]::new(8, 4, 8, 4)
$script:logList.ColumnHeadersHeightSizeMode = [System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode]::DisableResizing
$script:CurrentWindowScale = 1.0
$timeColumn = [System.Windows.Forms.DataGridViewTextBoxColumn]::new(); $timeColumn.Name = 'Time'; $timeColumn.HeaderText = 'الوقت'; $timeColumn.Width = 100; $timeColumn.SortMode = [System.Windows.Forms.DataGridViewColumnSortMode]::Automatic
$messageColumn = [System.Windows.Forms.DataGridViewTextBoxColumn]::new(); $messageColumn.Name = 'Message'; $messageColumn.HeaderText = 'العملية والتفاصيل'; $messageColumn.AutoSizeMode = [System.Windows.Forms.DataGridViewAutoSizeColumnMode]::Fill; $messageColumn.MinimumWidth = 260; $messageColumn.SortMode = [System.Windows.Forms.DataGridViewColumnSortMode]::Automatic
$statusColumn = [System.Windows.Forms.DataGridViewTextBoxColumn]::new(); $statusColumn.Name = 'Status'; $statusColumn.HeaderText = 'الحالة'; $statusColumn.Width = 110; $statusColumn.SortMode = [System.Windows.Forms.DataGridViewColumnSortMode]::Automatic
$idColumn = [System.Windows.Forms.DataGridViewTextBoxColumn]::new(); $idColumn.Name = 'Identifier'; $idColumn.HeaderText = 'المعرف أو المسار'; $idColumn.Width = 250; $idColumn.SortMode = [System.Windows.Forms.DataGridViewColumnSortMode]::Automatic
$idColumn.DefaultCellStyle.Alignment = [System.Windows.Forms.DataGridViewContentAlignment]::MiddleLeft
$idColumn.DefaultCellStyle.WrapMode = [System.Windows.Forms.DataGridViewTriState]::True
[void]$script:logList.Columns.AddRange([System.Windows.Forms.DataGridViewColumn[]]@($timeColumn, $messageColumn, $statusColumn, $idColumn))
$logGrid.Controls.Add($script:logList, 0, 1)

$script:statusPanel = New-UiPanel -Back $script:Colors.SurfaceHighest
$script:statusPanel.Margin = [System.Windows.Forms.Padding]::new(0)
$script:statusPanel.Padding = [System.Windows.Forms.Padding]::new(18, 8, 18, 8)
$main.Controls.Add($script:statusPanel, 0, 4)
$script:statusLayout = [System.Windows.Forms.TableLayoutPanel]::new()
$script:statusLayout.Dock = [System.Windows.Forms.DockStyle]::Fill
$script:statusLayout.BackColor = $script:statusPanel.BackColor
$script:statusLayout.RightToLeft = [System.Windows.Forms.RightToLeft]::No
$script:statusLayout.ColumnCount = 4
$script:statusLayout.ColumnStyles.Add([System.Windows.Forms.ColumnStyle]::new([System.Windows.Forms.SizeType]::Absolute, 150))
$script:statusLayout.ColumnStyles.Add([System.Windows.Forms.ColumnStyle]::new([System.Windows.Forms.SizeType]::Percent, 100))
$script:statusLayout.ColumnStyles.Add([System.Windows.Forms.ColumnStyle]::new([System.Windows.Forms.SizeType]::Absolute, 185))
$script:statusLayout.ColumnStyles.Add([System.Windows.Forms.ColumnStyle]::new([System.Windows.Forms.SizeType]::Absolute, 132))
$script:statusPanel.Controls.Add($script:statusLayout)
$script:authorLabel = New-UiLabel -Text 'Made By Salman' -Size 8 -Color $script:Colors.MutedText -Align ([System.Drawing.ContentAlignment]::MiddleLeft)
$script:authorLabel.RightToLeft = [System.Windows.Forms.RightToLeft]::No
$script:authorLabel.Font = [System.Drawing.Font]::new('Segoe UI Semibold', [Math]::Max(5.0, 8.5 * $script:FontDpiScale))
$script:statusLayout.Controls.Add($script:authorLabel, 0, 0)
$script:statusBadgePanel = New-UiPanel -Radius 16 -Back $script:Colors.TealSoft
$script:statusBadgePanel.Dock = [System.Windows.Forms.DockStyle]::Fill
$script:statusBadgePanel.Margin = [System.Windows.Forms.Padding]::new(8, 4, 0, 4)
$script:statusBadgePanel.BorderColor = $script:Colors.Teal
$script:statusBadge = New-UiLabel -Text '●  جاهز' -Size 9 -Style ([System.Drawing.FontStyle]::Bold) -Color $script:Colors.PrimaryLight -Align ([System.Drawing.ContentAlignment]::MiddleCenter)
$script:statusBadge.AutoEllipsis = $false
$script:statusBadgePanel.Controls.Add($script:statusBadge)
$script:lblStatusKind = $script:statusBadge
$script:statusLayout.Controls.Add($script:statusBadgePanel, 3, 0)
$script:statusLabel = New-UiLabel -Text 'النظام جاهز. ابدأ بمعاينة الملفات أو افتح المعاينة الحالية لمراجعتها.' -Align ([System.Drawing.ContentAlignment]::MiddleCenter)
$script:statusLabel.RightToLeft = [System.Windows.Forms.RightToLeft]::Yes
$script:lblStatus = $script:statusLabel
$script:statusLayout.Controls.Add($script:statusLabel, 1, 0)
$script:progress = [SessionMinutesUI.SmoothProgressBar]::new()
$script:progress.Dock = [System.Windows.Forms.DockStyle]::Fill
$script:progress.Margin = [System.Windows.Forms.Padding]::new(8, 12, 0, 12)
$script:progress.Style = [System.Windows.Forms.ProgressBarStyle]::Marquee
$script:progress.MarqueeAnimationSpeed = 28
$script:progress.Visible = $false
$script:statusLayout.Controls.Add($script:progress, 2, 0)
$script:ResponsiveFonts = [System.Collections.Generic.List[object]]::new()
$script:AppliedWindowScale = 0.0
function Register-ResponsiveFonts {
    param([System.Windows.Forms.Control]$Root)
    if (
        $Root -is [System.Windows.Forms.Label] -or
        $Root -is [System.Windows.Forms.Button] -or
        $Root -is [System.Windows.Forms.CheckBox] -or
        $Root -is [System.Windows.Forms.RichTextBox] -or
        $Root -is [System.Windows.Forms.DataGridView]
    ) {
        $script:ResponsiveFonts.Add([pscustomobject]@{
            Control = $Root
            Size = [float]$Root.Font.Size
            Family = $Root.Font.FontFamily.Name
            Style = $Root.Font.Style
        })
    }
    foreach ($child in $Root.Controls) {
        Register-ResponsiveFonts -Root $child
    }
}
Register-ResponsiveFonts -Root $script:form
function Update-ResponsiveLayout {
    $dpiScale = [Math]::Max(1.0, $script:DisplayDpi / 96.0)
    $resolutionScale = [Math]::Min(2.0, [Math]::Max(1.0, $script:form.ClientSize.Width / 1920.0))
    $desiredScale = [Math]::Max($dpiScale, $resolutionScale)
    $heightCapacity = [Math]::Max(0.70, ($script:form.ClientSize.Height - 180.0) / 652.0)
    $widthCapacity = [Math]::Max(0.70, $script:form.ClientSize.Width / 1000.0)
    $windowScale = [Math]::Min(2.5, [Math]::Min($desiredScale, [Math]::Min($heightCapacity, $widthCapacity)))
    $windowScale = [Math]::Round($windowScale * 20.0) / 20.0
    $compact = $windowScale -lt 1.15
    $script:CurrentWindowScale = $windowScale
    $fontScale = [Math]::Max(0.55, $windowScale / $dpiScale)
    if ([Math]::Abs($fontScale - $script:AppliedWindowScale) -gt 0.01) {
        foreach ($entry in $script:ResponsiveFonts) {
            $entry.Control.Font = [System.Drawing.Font]::new(
                $entry.Family,
                [Math]::Max(5.0, $entry.Size * $fontScale),
                $entry.Style,
                [System.Drawing.GraphicsUnit]::Point
            )
        }
        $script:AppliedWindowScale = $fontScale
    }
    # The warning title used to live in a fixed 32px row. At high DPI the
    # glyph height exceeded that row and Word users saw only the lower half of
    # "قبل المتابعة". Derive the row height from the actual rendered font.
    $warningTitleHeight = [int][Math]::Ceiling($script:lblWarningTitle.Font.GetHeight() + 16.0)
    $gateGrid.RowStyles[0].Height = [Math]::Max(38, $warningTitleHeight)
    if ($compact) {
        $main.Padding = [System.Windows.Forms.Padding]::new(18, 8, 18, 8)
        $main.RowStyles[0].Height = 52
        $main.RowStyles[1].Height = 132
        $main.RowStyles[2].Height = 310
        $main.RowStyles[4].Height = 48
        $script:statusLayout.ColumnStyles[0].Width = 132
        $script:statusLayout.ColumnStyles[2].Width = 135
        $hero.Padding = [System.Windows.Forms.Padding]::new(0, 10, 0, 4)
        $heroGrid.RowStyles[0].Height = 44
        $heroGrid.RowStyles[1].Height = 30
        $workflow.RowStyles[0].Height = 38
        $workflow.RowStyles[1].Height = 210
        $workflow.RowStyles[2].Height = 82
        foreach ($card in @($script:stepPreviewCard, $script:stepPilotCard, $script:stepApplyCard)) {
            $card.Margin = [System.Windows.Forms.Padding]::new(5, 4, 5, 4)
            $card.Padding = [System.Windows.Forms.Padding]::new(14, 10, 14, 10)
        }
        foreach ($layout in $script:workflowCardLayouts) {
            $layout.ColumnStyles[0].Width = 58
            $layout.RowStyles[0].Height = 38
            $layout.RowStyles[1].Height = 30
            $layout.RowStyles[3].Height = 8
            $layout.RowStyles[4].Height = 40
            $layout.Tag.IconTile.Margin = [System.Windows.Forms.Padding]::new(10, 0, 0, 0)
            $layout.Tag.StatePill.Size = [System.Drawing.Size]::new(104, 28)
        }
    }
    else {
        $widePadding = [Math]::Max(18, [int][Math]::Floor($script:form.ClientSize.Width * 0.035))
        $main.Padding = [System.Windows.Forms.Padding]::new($widePadding, 16, $widePadding, 16)
        $main.RowStyles[0].Height = [int][Math]::Round(58 * $windowScale)
        $main.RowStyles[1].Height = [int][Math]::Round(136 * $windowScale)
        $main.RowStyles[2].Height = [int][Math]::Round(365 * $windowScale)
        $main.RowStyles[4].Height = [int][Math]::Round(52 * $windowScale)
        $script:statusLayout.ColumnStyles[0].Width = [int][Math]::Round(150 * $windowScale)
        $script:statusLayout.ColumnStyles[2].Width = [int][Math]::Round(185 * $windowScale)
        $hero.Padding = [System.Windows.Forms.Padding]::new(0, [int][Math]::Round(18 * $windowScale), 0, [int][Math]::Round(8 * $windowScale))
        $heroGrid.RowStyles[0].Height = [int][Math]::Round(50 * $windowScale)
        $heroGrid.RowStyles[1].Height = [int][Math]::Round(34 * $windowScale)
        $workflow.RowStyles[0].Height = [int][Math]::Round(42 * $windowScale)
        $workflow.RowStyles[1].Height = [int][Math]::Round(230 * $windowScale)
        $workflow.RowStyles[2].Height = [int][Math]::Round(84 * $windowScale)
        $nav.ColumnStyles[0].Width = [int][Math]::Round(50 * $windowScale)
        $nav.ColumnStyles[1].Width = [int][Math]::Round(250 * $windowScale)
        foreach ($card in @($script:stepPreviewCard, $script:stepPilotCard, $script:stepApplyCard)) {
            $card.Margin = [System.Windows.Forms.Padding]::new([int][Math]::Round(9 * $windowScale), [int][Math]::Round(7 * $windowScale), [int][Math]::Round(9 * $windowScale), [int][Math]::Round(7 * $windowScale))
            $card.Padding = [System.Windows.Forms.Padding]::new([int][Math]::Round(22 * $windowScale), [int][Math]::Round(17 * $windowScale), [int][Math]::Round(22 * $windowScale), [int][Math]::Round(17 * $windowScale))
        }
        foreach ($layout in $script:workflowCardLayouts) {
            $layout.ColumnStyles[0].Width = [int][Math]::Round(66 * $windowScale)
            $layout.RowStyles[0].Height = [int][Math]::Round(44 * $windowScale)
            $layout.RowStyles[1].Height = [int][Math]::Round(32 * $windowScale)
            $layout.RowStyles[3].Height = [int][Math]::Round(10 * $windowScale)
            $layout.RowStyles[4].Height = [int][Math]::Round(44 * $windowScale)
            $layout.Tag.IconTile.Margin = [System.Windows.Forms.Padding]::new([int][Math]::Round(13 * $windowScale), 0, 0, 0)
            $layout.Tag.IconTile.CornerRadius = [int][Math]::Round(8 * $windowScale)
            $layout.Tag.StatePill.Size = [System.Drawing.Size]::new([int][Math]::Round(118 * $windowScale), [int][Math]::Round(30 * $windowScale))
            $layout.Tag.StatePill.CornerRadius = [int][Math]::Round(14 * $windowScale)
        }
    }
    $logPanel.Padding = [System.Windows.Forms.Padding]::new(
        [int][Math]::Round(16 * $windowScale),
        [int][Math]::Round(12 * $windowScale),
        [int][Math]::Round(16 * $windowScale),
        [int][Math]::Round(14 * $windowScale)
    )
    $logGrid.RowStyles[0].Height = [int][Math]::Round(42 * $windowScale)
    $timeWidth = [int][Math]::Round(100 * $windowScale)
    $stateWidth = [int][Math]::Round(110 * $windowScale)
    $idWidth = [int][Math]::Round(250 * $windowScale)
    $script:logList.ColumnHeadersHeight = [int][Math]::Round(34 * $windowScale)
    $script:logList.RowTemplate.MinimumHeight = [int][Math]::Round(28 * $windowScale)
    if ($script:logList.Columns.Count -ge 4) {
        $script:logList.Columns[0].Width = $timeWidth
        $script:logList.Columns[2].Width = $stateWidth
        $script:logList.Columns[3].Width = $idWidth
        $script:logList.Columns[3].DefaultCellStyle.Font = [System.Drawing.Font]::new('Consolas', [Math]::Max(5.0, 8.0 * $fontScale))
    }
}

function Fit-WindowToCurrentScreen {
    $area = [System.Windows.Forms.Screen]::FromControl($script:form).WorkingArea
    $newWidth = [Math]::Min($script:form.Width, $area.Width)
    $newHeight = [Math]::Min($script:form.Height, $area.Height)
    if ($newWidth -ne $script:form.Width -or $newHeight -ne $script:form.Height) {
        $script:form.Size = [System.Drawing.Size]::new($newWidth, $newHeight)
    }
    Update-ResponsiveLayout
}

$script:form.Add_Resize({ Update-ResponsiveLayout })
try {
    $script:form.Add_DpiChanged({ Fit-WindowToCurrentScreen })
}
catch {
    # Older .NET Framework versions still use the initial DPI-aware sizing path.
}

$toolTip = [System.Windows.Forms.ToolTip]::new()
$toolTip.AutoPopDelay = 9000
$toolTip.InitialDelay = 400
$toolTip.ReshowDelay = 100
$toolTip.SetToolTip($script:btnPreview, 'يفحص النموذج وملفات المصدر ويكتب تقريرًا فقط دون إنشاء ملفات Word.')
$toolTip.SetToolTip($script:btnPreviewReport, 'يفتح تقرير المعاينة لمراجعة الإجراءات والتحذيرات قبل إنشاء الملفات.')
$toolTip.SetToolTip($script:btnPilot, 'ينشئ مجموعة اختبار داخل outputs\pilot ولا يغيّر ملفات المصدر.')
$toolTip.SetToolTip($script:btnApply, 'ينشئ النسخ النهائية بعد معاينة الملفات، ويمكن استخدامه دون إنشاء تجربة مسبقة.')
$toolTip.SetToolTip($script:btnReviewNext, 'يفتح ملفات أحدث نتيجة واحدًا بعد الآخر في Microsoft Word.')
$toolTip.SetToolTip($script:btnOpenOutput, 'يفتح مجلد أحدث تجربة أو نتيجة نهائية تم التحقق منها.')

$script:processTimer = [System.Windows.Forms.Timer]::new()
$script:processTimer.Interval = 300
$script:processTimer.Add_Tick({
    if ($null -ne $script:CurrentProcess -and $script:CurrentProcess.HasExited) {
        $script:processTimer.Stop()
        Complete-EngineCommand
    }
})

$script:btnPreview.Add_Click({ Start-EngineCommand -Mode 'preview' })
$script:btnPreviewReport.Add_Click({ Open-PreviewReport })
$script:btnPilot.Add_Click({ Start-EngineCommand -Mode 'pilot' })
$script:btnApply.Add_Click({ Start-EngineCommand -Mode 'apply' })
$script:btnOpenOutput.Add_Click({ Open-LatestOutput })
$script:btnReviewNext.Add_Click({ Open-NextReviewDocument })

$script:form.Add_FormClosing({
    param($sender, $eventArgs)
    if ($null -ne $script:CurrentProcess) {
        [System.Windows.Forms.MessageBox]::Show(
            'انتظر حتى تنتهي العملية الحالية قبل إغلاق النافذة.',
            'عملية قيد التنفيذ',
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Information
        ) | Out-Null
        $eventArgs.Cancel = $true
    }
})

$script:form.Add_Shown({
    Fit-WindowToCurrentScreen
    $safeProjectPath = ConvertTo-BidiSafeToken -Value $script:ProjectRoot
    Add-LogLine -Message "مجلد المشروع: $safeProjectPath"
    Add-LogLine -Message 'الأداة لا تكتب داخل مجلد الدفعة الرقمي؛ كل النتائج داخل outputs.'
    if (Test-Prerequisites -Quiet) {
        $safeEnginePath = ConvertTo-BidiSafeToken -Value $script:EnginePath
        $safeBatchRoot = ConvertTo-BidiSafeToken -Value $script:BatchRoot
        Add-LogLine -Message "تم العثور على محرك النسخة المحمولة: $safeEnginePath" -Kind Success
        Add-LogLine -Message "مجلد الدفعة المكتشف: $safeBatchRoot" -Kind Success
        if (Set-PreviewFromManifest) {
            Add-LogLine -Message 'توجد معاينة سابقة. افتح تقريرها وتأكد من أنها تخص ملفات اليوم ومن صحة البيانات وسلامة الملفات.' -Kind Warning
            Set-Status -Message 'النظام جاهز. ابدأ بمعاينة جديدة أو افتح المعاينة الحالية لمراجعتها.' -Kind Normal
        }
        else {
            Set-Status -Message 'لا توجد معاينة حالية. ابدأ بالفحص والمعاينة.' -Kind Empty
        }
    }
    else {
        [void](Test-Prerequisites)
    }
    [void](Load-LatestResultsContext)
    Update-Controls
})

try {
    [void]$script:form.ShowDialog()
}
catch {
    Write-Error -ErrorRecord $_
    if ($env:SESSION_MINUTES_UI_DIAGNOSTIC -eq '1') {
        exit 1
    }
    [System.Windows.Forms.MessageBox]::Show(
        "تعذر تشغيل واجهة المحاضر:`r`n$($_.Exception.Message)",
        'محاضر جلسات اللجنة',
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Error
    ) | Out-Null
    exit 1
}

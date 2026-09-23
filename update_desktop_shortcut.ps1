$desktopPaths = @(
    [Environment]::GetFolderPath('Desktop'),
    "C:\Users\Administrador\Desktop",
    "C:\Users\Administrador\OneDrive\Desktop"
) | Select-Object -Unique

$projectDir = "C:\Users\Administrador\.gemini\antigravity-ide\scratch\ConsultorPro-II"
$icoPath = Join-Path $projectDir "public\icon.ico"
$chromeProxy = "C:\Program Files\Google\Chrome\Application\chrome_proxy.exe"
$arguments = "--profile-directory=Default --app-id=ibnejcobcbimbmpbffelbhaaabmbnbmd"

$wsh = New-Object -ComObject WScript.Shell

foreach ($dPath in $desktopPaths) {
    if (Test-Path $dPath) {
        $shortcutPath = Join-Path $dPath "Consultor Pro.lnk"
        $shortcut = $wsh.CreateShortcut($shortcutPath)
        $shortcut.TargetPath = $chromeProxy
        $shortcut.Arguments = $arguments
        $shortcut.WorkingDirectory = "C:\Program Files\Google\Chrome\Application"
        $shortcut.IconLocation = "$icoPath,0"
        $shortcut.Description = "Consultor Pró - Gestor Empresarial"
        $shortcut.Save()
        Write-Host "Atalho atualizado com sucesso em: $shortcutPath"
    }
}

# Notificar o Windows Explorer para atualizar os ícones na Área de Trabalho
try {
    Add-Type -TypeDefinition @"
    using System;
    using System.Runtime.InteropServices;
    public class ShellNotification {
        [DllImport("shell32.dll")]
        public static extern void SHChangeNotify(int wEventId, uint uFlags, IntPtr dwItem1, IntPtr dwItem2);
    }
"@ -ErrorAction SilentlyContinue
    [ShellNotification]::SHChangeNotify(0x08000000, 0, [IntPtr]::Zero, [IntPtr]::Zero)
    Write-Host "Área de Trabalho notificada para atualização de ícone com sucesso."
} catch {
    # Fallback se tipo já foi adicionado na sessão
    try {
        [ShellNotification]::SHChangeNotify(0x08000000, 0, [IntPtr]::Zero, [IntPtr]::Zero)
    } catch {}
}

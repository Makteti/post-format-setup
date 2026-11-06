<#
.SYNOPSIS
    Post-format app installer: a single-file PowerShell + WPF front end for winget.

.DESCRIPTION
    - Zero external files: the XAML, styles, C# helpers and icon logic all live in this script.
    - winget runs in a background runspace; every status change is pushed back to the UI
      thread with Dispatcher.BeginInvoke (see [PFI.Ui]), so the window never blocks.
    - Apps install one after another. winget/MSI do not handle parallel installs well.

.USAGE
    Right-click > "Run with PowerShell", or:
        powershell -NoProfile -ExecutionPolicy Bypass -File .\PostFormatInstaller.ps1
    Preview the UI without installing anything:
        powershell -NoProfile -ExecutionPolicy Bypass -File .\PostFormatInstaller.ps1 -DryRun

    Run it as your normal user, NOT "as administrator". Per-machine installers raise their own
    UAC prompt through winget. Some per-user installers (Spotify in particular) refuse to run elevated.

.PARAMETER DryRun
    Simulates downloads and installs so you can check the UI. Nothing is installed.
.PARAMETER Snapshot
    Development aid: starts a dry run, saves a PNG of the window to this path, then exits.
#>
[CmdletBinding()]
param(
    [switch]$DryRun,
    [string]$Snapshot,
    [string]$SnapshotSize   # development aid, e.g. "1280x860"
)

# =====================================================================================
#  APP CATALOG: one line per app, so adding software is a single line
# =====================================================================================
#  Columns (separated by | , spacing is ignored):
#    Category   the group heading in the list. A new name creates a new group.
#               Every app in "Browsers" also appears in the "Default browser" picker.
#    Name       what the user sees
#    winget ID  exact package ID. Find one with:  winget search <name>
#    Icon site  the publisher's website. Its official icon is fetched once and cached in
#               %LOCALAPPDATA%\PostFormatInstaller\icons. Leave empty to use the installed exe's icon.
#    On         "on" = ticked by default
#  Optional extras per app (by winget ID), below the table:
#    $IconOverrides   base64 PNG embedded in the script (fully offline icon)
#    $MatchOverrides  regex for the Add/Remove Programs name, used for the offline exe-icon fallback
#                     (default: the app name)
#  Every ID below was checked with "winget show --exact" on 2026-11-06.
$CatalogTable = @'
# Category          | Name                         | winget ID                                | Icon site                  | On
Browsers            | Brave                        | Brave.Brave                              | brave.com                  | on
Browsers            | Firefox                      | Mozilla.Firefox                          | firefox.com                |
Browsers            | Google Chrome                | Google.Chrome                            | chrome.com                 |
Browsers            | Opera GX                     | Opera.OperaGX                            | operagx.gg                 |
Browsers            | Opera                        | Opera.Opera                              | opera.com                  |
Browsers            | Vivaldi                      | Vivaldi.Vivaldi                          | vivaldi.com                |
Browsers            | LibreWolf                    | LibreWolf.LibreWolf                      | librewolf.net              |
Browsers            | Zen Browser                  | Zen-Team.Zen-Browser                     | zen-browser.app            |
Browsers            | Tor Browser                  | TorProject.TorBrowser                    | torproject.org             |

Communication       | Discord                      | Discord.Discord                          | discord.com                | on
Communication       | TeamSpeak 3                  | TeamSpeakSystems.TeamSpeakClient         | teamspeak.com              | on
# winget lists TeamSpeak 5 under ".Beta.5". Use "...Beta.6" for TeamSpeak 6.
Communication       | TeamSpeak 5                  | TeamSpeakSystems.TeamSpeakClient.Beta.5  | teamspeak.com              | on
Communication       | Telegram                     | Telegram.TelegramDesktop                 | telegram.org               |
Communication       | Signal                       | OpenWhisperSystems.Signal                | signal.org                 |
Communication       | Zoom                         | Zoom.Zoom                                | zoom.us                    |
Communication       | Microsoft Teams              | Microsoft.Teams                          | teams.microsoft.com        |
Communication       | Slack                        | SlackTechnologies.Slack                  | slack.com                  |
Communication       | Mumble                       | Mumble.Mumble.Client                     | mumble.info                |
Communication       | Element                      | Element.Element                          | element.io                 |
Communication       | Thunderbird                  | Mozilla.Thunderbird                      | thunderbird.net            |
Communication       | Viber                        | Rakuten.Viber                            | viber.com                  |

Gaming              | Steam                        | Valve.Steam                              | steampowered.com           |
Gaming              | GOG Galaxy                   | GOG.Galaxy                               | gog.com                    |
Gaming              | Epic Games Launcher          | EpicGames.EpicGamesLauncher              | epicgames.com              |
Gaming              | EA app                       | ElectronicArts.EADesktop                 | ea.com                     |
Gaming              | Ubisoft Connect              | Ubisoft.Connect                          | ubisoft.com                |
Gaming              | Battle.net                   | Blizzard.BattleNet                       | battle.net                 |
Gaming              | itch                         | ItchIo.Itch                              | itch.io                    |
Gaming              | Playnite                     | Playnite.Playnite                        | playnite.link              |
Gaming              | Heroic Games Launcher        | HeroicGamesLauncher.HeroicGamesLauncher  | heroicgameslauncher.com    |
Gaming              | Minecraft Launcher           | Mojang.MinecraftLauncher                 | minecraft.net              |
Gaming              | Prism Launcher               | PrismLauncher.PrismLauncher              | prismlauncher.org          |
Gaming              | MSI Afterburner              | Guru3D.Afterburner                       | msi.com                    |
Gaming              | NVCleanstall                 | TechPowerUp.NVCleanstall                 | techpowerup.com            |
Gaming              | Parsec                       | Parsec.Parsec                            | parsec.app                 |

Media & Creative    | Spotify                      | Spotify.Spotify                          | spotify.com                | on
Media & Creative    | VLC media player             | VideoLAN.VLC                             | videolan.org               |
Media & Creative    | AIMP                         | AIMP.AIMP                                | aimp.ru                    |
Media & Creative    | Plex                         | Plex.Plex                                | plex.tv                    |
Media & Creative    | Stremio                      | Stremio.Stremio                          | stremio.com                |
Media & Creative    | OBS Studio                   | OBSProject.OBSStudio                     | obsproject.com             |
Media & Creative    | Audacity                     | Audacity.Audacity                        | audacityteam.org           |
Media & Creative    | HandBrake                    | HandBrake.HandBrake                      | handbrake.fr               |
Media & Creative    | Kdenlive                     | KDE.Kdenlive                             | kdenlive.org               |
Media & Creative    | Blender                      | BlenderFoundation.Blender                | blender.org                |
Media & Creative    | GIMP                         | GIMP.GIMP.3                              | gimp.org                   |
Media & Creative    | Inkscape                     | Inkscape.Inkscape                        | inkscape.org               |
Media & Creative    | Krita                        | KDE.Krita                                | krita.org                  |
Media & Creative    | Paint.NET                    | dotPDN.PaintDotNet                       | getpaint.net               |
Media & Creative    | ShareX                       | ShareX.ShareX                            | getsharex.com              |
Media & Creative    | Greenshot                    | Greenshot.Greenshot                      | getgreenshot.org           |

Productivity        | Notion                       | Notion.Notion                            | notion.so                  |
Productivity        | Obsidian                     | Obsidian.Obsidian                        | obsidian.md                |
Productivity        | Todoist                      | Doist.Todoist                            | todoist.com                |
Productivity        | LibreOffice                  | TheDocumentFoundation.LibreOffice        | libreoffice.org            |
Productivity        | ONLYOFFICE                   | ONLYOFFICE.DesktopEditors                | onlyoffice.com             |
Productivity        | Adobe Acrobat Reader         | Adobe.Acrobat.Reader.64-bit              | adobe.com                  |
Productivity        | SumatraPDF                   | SumatraPDF.SumatraPDF                    | sumatrapdfreader.org       |
Productivity        | PowerToys                    | Microsoft.PowerToys                      |                            |
Productivity        | Flow Launcher                | Flow-Launcher.Flow-Launcher              | flowlauncher.com           |
Productivity        | Everything                   | voidtools.Everything                     | voidtools.com              |

Development         | Visual Studio Code           | Microsoft.VisualStudioCode               | code.visualstudio.com      |
Development         | Visual Studio 2022 Community | Microsoft.VisualStudio.2022.Community    | visualstudio.microsoft.com |
Development         | Cursor                       | Anysphere.Cursor                         | cursor.com                 |
Development         | Zed                          | ZedIndustries.Zed                        | zed.dev                    |
Development         | JetBrains Toolbox            | JetBrains.Toolbox                        | jetbrains.com              |
Development         | Notepad++                    | Notepad++.Notepad++                      | notepad-plus-plus.org      |
Development         | Neovim                       | Neovim.Neovim                            | neovim.io                  |
Development         | Git                          | Git.Git                                  | git-scm.com                |
Development         | GitHub Desktop               | GitHub.GitHubDesktop                     | desktop.github.com         |
Development         | Windows Terminal             | Microsoft.WindowsTerminal                |                            |
Development         | PowerShell 7                 | Microsoft.PowerShell                     |                            |
Development         | Node.js LTS                  | OpenJS.NodeJS.LTS                        | nodejs.org                 |
Development         | Python 3.13                  | Python.Python.3.13                       | python.org                 |
Development         | Go                           | GoLang.Go                                | go.dev                     |
Development         | Rust (rustup)                | Rustlang.Rustup                          | rust-lang.org              |
Development         | Java 21 (Temurin JDK)        | EclipseAdoptium.Temurin.21.JDK           | adoptium.net               |
Development         | Docker Desktop               | Docker.DockerDesktop                     | docker.com                 |
Development         | Postman                      | Postman.Postman                          | postman.com                |
Development         | DBeaver Community            | DBeaver.DBeaver.Community                | dbeaver.io                 |
Development         | WinSCP                       | WinSCP.WinSCP                            | winscp.net                 |
Development         | PuTTY                        | PuTTY.PuTTY                              | putty.org                  |

Utilities           | 7-Zip                        | 7zip.7zip                                | 7-zip.org                  |
Utilities           | NanaZip                      | M2Team.NanaZip                           |                            |
Utilities           | WinRAR                       | RARLab.WinRAR                            | win-rar.com                |
Utilities           | CPU-Z                        | CPUID.CPU-Z                              | cpuid.com                  |
Utilities           | GPU-Z                        | TechPowerUp.GPU-Z                        | techpowerup.com            |
Utilities           | HWiNFO                       | REALiX.HWiNFO                            | hwinfo.com                 |
Utilities           | CrystalDiskInfo              | CrystalDewWorld.CrystalDiskInfo          | crystalmark.info           |
Utilities           | WizTree                      | AntibodySoftware.WizTree                 | diskanalyzer.com           |
Utilities           | WinDirStat                   | WinDirStat.WinDirStat                    | windirstat.net             |
Utilities           | Revo Uninstaller             | RevoUninstaller.RevoUninstaller          | revouninstaller.com        |
Utilities           | BleachBit                    | BleachBit.BleachBit                      | bleachbit.org              |
Utilities           | Rufus                        | Rufus.Rufus                              | rufus.ie                   |
Utilities           | Ventoy                       | Ventoy.Ventoy                            | ventoy.net                 |
Utilities           | Logitech G HUB               | Logitech.GHUB                            | logitechg.com              |
Utilities           | Visual C++ Redist 2015+ x64  | Microsoft.VCRedist.2015+.x64             |                            |
Utilities           | .NET 8 Desktop Runtime       | Microsoft.DotNet.DesktopRuntime.8        | dotnet.microsoft.com       |

Security & Privacy  | Bitwarden                    | Bitwarden.Bitwarden                      | bitwarden.com              |
Security & Privacy  | 1Password                    | AgileBits.1Password                      | 1password.com              |
Security & Privacy  | KeePassXC                    | KeePassXCTeam.KeePassXC                  | keepassxc.org              |
Security & Privacy  | Malwarebytes                 | Malwarebytes.Malwarebytes                | malwarebytes.com           |
Security & Privacy  | Proton VPN                   | Proton.ProtonVPN                         | protonvpn.com              |
Security & Privacy  | WireGuard                    | WireGuard.WireGuard                      | wireguard.com              |
Security & Privacy  | Tailscale                    | Tailscale.Tailscale                      | tailscale.com              |

Cloud & Remote      | Google Drive                 | Google.GoogleDrive                       | drive.google.com           |
Cloud & Remote      | Dropbox                      | Dropbox.Dropbox                          | dropbox.com                |
Cloud & Remote      | MEGA                         | Mega.MEGASync                            | mega.io                    |
Cloud & Remote      | qBittorrent                  | qBittorrent.qBittorrent                  | qbittorrent.org            |
Cloud & Remote      | AnyDesk                      | AnyDesk.AnyDesk                          | anydesk.com                |
Cloud & Remote      | TeamViewer                   | TeamViewer.TeamViewer                    | teamviewer.com             |
'@

$IconOverrides = @{
    # 'Vendor.App' = 'iVBORw0KGgoAAAANSUhEUg...'
}
$MatchOverrides = @{
    'TeamSpeakSystems.TeamSpeakClient'        = '^TeamSpeak 3'
    'TeamSpeakSystems.TeamSpeakClient.Beta.5' = '^TeamSpeak( 5| Beta|$)'
    'Microsoft.VisualStudioCode'              = '^Microsoft Visual Studio Code'
    'Microsoft.VCRedist.2015+.x64'            = 'Visual C\+\+ 20(15|17|19|22).*x64'
    'Microsoft.DotNet.DesktopRuntime.8'       = 'Windows Desktop Runtime - 8\.'
}
# Monogram tile colour per category, shown until the real icon has loaded
$CategoryColors = @{
    'Browsers' = '#C2410C'; 'Communication' = '#4F46E5'; 'Gaming' = '#15803D'; 'Media & Creative' = '#BE185D'
    'Productivity' = '#0369A1'; 'Development' = '#7C3AED'; 'Utilities' = '#475569'; 'Security & Privacy' = '#B45309'
    'Cloud & Remote' = '#0E7490'
}

$Catalog = foreach ($line in ($CatalogTable -split "`r?`n")) {
    if ($line -match '^\s*(#|$)') { continue }
    $f = @($line -split '\|' | ForEach-Object { $_.Trim() })
    if ($f.Count -lt 3) { continue }
    $words = @($f[1] -split '[\s\-]+' | Where-Object { $_ })
    $mono  = if ($words.Count -ge 2 -and $words[1] -match '^\w') { "$($words[0][0])$($words[1][0])" } else { "$($f[1][0])" }
    [pscustomobject]@{
        Category   = $f[0]
        Name       = $f[1]
        Id         = $f[2]
        Domain     = if ($f.Count -gt 3) { $f[3] } else { '' }
        Selected   = ($f.Count -gt 4 -and $f[4] -eq 'on')
        Browser    = if ($f[0] -eq 'Browsers') { $f[1] } else { $null }
        Mono       = $mono.ToUpper()
        Brand      = if ($CategoryColors[$f[0]]) { $CategoryColors[$f[0]] } else { '#475569' }
        Match      = if ($MatchOverrides[$f[2]]) { $MatchOverrides[$f[2]] } else { '^' + [regex]::Escape($f[1]) }
        IconBase64 = $IconOverrides[$f[2]]
    }
}

# =====================================================================================
#  Bootstrap: WPF needs an STA thread (pwsh 7 starts MTA by default)
# =====================================================================================
if ([Threading.Thread]::CurrentThread.ApartmentState -ne 'STA') {
    $exe  = (Get-Process -Id $PID).Path
    $argv = @('-NoProfile', '-ExecutionPolicy', 'Bypass', '-STA', '-File', "`"$PSCommandPath`"")
    if ($DryRun)   { $argv += '-DryRun' }
    if ($Snapshot) { $argv += @('-Snapshot', "`"$Snapshot`"") }
    Start-Process -FilePath $exe -ArgumentList $argv
    return
}

Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase, System.Drawing

$AppDir  = Join-Path $env:LOCALAPPDATA 'PostFormatInstaller'
$IconDir = Join-Path $AppDir 'icons'
$null    = New-Item -ItemType Directory -Force -Path $IconDir
$LogPath = Join-Path $AppDir ('install-{0:yyyyMMdd-HHmmss}.log' -f (Get-Date))

# =====================================================================================
#  C# helpers: view models, Dispatcher marshalling, default-browser natives
# =====================================================================================
if (-not ('PFI.AppItem' -as [type])) {
Add-Type -ReferencedAssemblies ([System.Windows.Threading.Dispatcher].Assembly.Location) -TypeDefinition @'
using System;
using System.ComponentModel;
using System.Text;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using System.Security.Principal;
using System.Windows.Threading;
using Microsoft.Win32;

namespace PFI {
    public abstract class Bindable : INotifyPropertyChanged {
        public event PropertyChangedEventHandler PropertyChanged;
        protected void Set<T>(ref T field, T value, string name) {
            if (Equals(field, value)) return;
            field = value;
            var h = PropertyChanged; if (h != null) h(this, new PropertyChangedEventArgs(name));
        }
    }

    // One row in the list. State drives the pill colour in XAML:
    // Idle | Checking | Installed | Queued | Active | Success | Failed | Skipped
    public class AppItem : Bindable {
        string _name, _id, _status = "Ready", _state = "Idle", _brand = "#334155", _mono = "?", _detail;
        bool _sel = true, _editable = true, _indet;
        double _progress; byte[] _icon;
        string _category;
        public string Category  { get { return _category; } set { Set(ref _category, value, "Category"); } }
        public string Name      { get { return _name; }     set { Set(ref _name, value, "Name"); } }
        public string Id        { get { return _id; }       set { Set(ref _id, value, "Id"); } }
        public string Status    { get { return _status; }   set { Set(ref _status, value, "Status"); } }
        public string State     { get { return _state; }    set { Set(ref _state, value, "State"); } }
        public string BrandHex  { get { return _brand; }    set { Set(ref _brand, value, "BrandHex"); } }
        public string Monogram  { get { return _mono; }     set { Set(ref _mono, value, "Monogram"); } }
        public string Detail    { get { return _detail; }   set { Set(ref _detail, value, "Detail"); } }
        public bool IsSelected  { get { return _sel; }      set { Set(ref _sel, value, "IsSelected"); } }
        public bool IsEditable  { get { return _editable; } set { Set(ref _editable, value, "IsEditable"); } }
        public bool Indeterminate { get { return _indet; }  set { Set(ref _indet, value, "Indeterminate"); } }
        public double Progress  { get { return _progress; } set { Set(ref _progress, value, "Progress"); } }
        public byte[] IconBytes { get { return _icon; }     set { Set(ref _icon, value, "IconBytes"); Set(ref _hasIcon, value != null, "HasIcon"); } }
        bool _hasIcon; public bool HasIcon { get { return _hasIcon; } }
    }

    public class SessionState : Bindable {
        bool _busy; string _summary = ""; double _progress; int _columns = 3;
        public int Columns      { get { return _columns; }  set { Set(ref _columns, value, "Columns"); } }
        public bool Busy        { get { return _busy; }     set { Set(ref _busy, value, "Busy"); } }
        public string Summary   { get { return _summary; }  set { Set(ref _summary, value, "Summary"); } }
        public double Progress  { get { return _progress; } set { Set(ref _progress, value, "Progress"); } }
    }

    // Background runspaces call ONLY these methods. Each queues the property change onto the
    // WPF UI thread through Dispatcher.BeginInvoke, so worker threads never touch UI objects.
    public static class Ui {
        public static void Status(Dispatcher d, AppItem it, string status, string state, double progress, bool indeterminate, string detail) {
            d.BeginInvoke(new Action(delegate {
                it.Status = status; it.State = state; it.Indeterminate = indeterminate;
                if (progress >= 0) it.Progress = progress;
                if (detail != null) it.Detail = detail.Length == 0 ? null : detail;
            }));
        }
        public static void Icon(Dispatcher d, AppItem it, byte[] png) {
            d.BeginInvoke(new Action(delegate { it.IconBytes = png; }));
        }
        public static void ScanResult(Dispatcher d, AppItem it, bool installed) {
            d.BeginInvoke(new Action(delegate {
                if (it.State != "Checking") return;            // an install already took over this row
                if (installed) { it.State = "Installed"; it.Status = "Installed"; if (it.IsEditable) it.IsSelected = false; }
                else           { it.State = "Idle";      it.Status = "Not installed"; }
            }));
        }
        public static void Summary(Dispatcher d, SessionState s, string text, double progress) {
            d.BeginInvoke(new Action(delegate { if (text != null) s.Summary = text; if (progress >= 0) s.Progress = progress; }));
        }
        public static void Busy(Dispatcher d, SessionState s, bool busy) {
            d.BeginInvoke(new Action(delegate { s.Busy = busy; }));
        }
    }

    public static class Native {
        [DllImport("shlwapi.dll", CharSet = CharSet.Unicode)]
        static extern int AssocQueryString(int flags, int str, string assoc, string extra, StringBuilder outStr, ref uint cch);
        [DllImport("shell32.dll")]
        static extern void SHChangeNotify(int eventId, int flags, IntPtr item1, IntPtr item2);
        [DllImport("dwmapi.dll")]
        static extern int DwmSetWindowAttribute(IntPtr hwnd, int attr, ref int value, int size);

        // The ProgId Windows will actually launch for a protocol (e.g. "https"). This is the
        // effective default, so it is how the tool checks whether a change took effect.
        public static string ProgIdForProtocol(string protocol) {
            uint cch = 512; var sb = new StringBuilder(512);
            int hr = AssocQueryString(0x1000 /*ASSOCF_IS_PROTOCOL*/, 20 /*ASSOCSTR_PROGID*/, protocol, null, sb, ref cch);
            return hr == 0 ? sb.ToString() : null;
        }
        public static void NotifyAssocChanged() { SHChangeNotify(0x08000000 /*SHCNE_ASSOCCHANGED*/, 0, IntPtr.Zero, IntPtr.Zero); }
        public static void RoundCorners(IntPtr hwnd) { int pref = 2; try { DwmSetWindowAttribute(hwnd, 33, ref pref, 4); } catch { } }
    }

    // Legacy UserChoice hash (Windows 8 to Windows 11 builds before "UserChoiceLatest").
    // Windows rejects a UserChoice key whose Hash does not match
    //   lower(assoc + userSid + progId + keyLastWriteMinuteAsFileTimeHex + experienceString)
    // run through MD5 and a custom scramble. This port was checked against hashes Windows wrote itself.
    public static class UserChoice {
        const string Experience = "User Choice set via Windows User Experience {D18B6DD5-6124-4341-9318-804003BAFA0B}";
        static uint Swap(uint v) { return (v >> 16) | (v << 16); }

        public static string Hash(string assoc, string sid, string progId, long fileTimeUtc) {
            string s = (assoc + sid + progId + ((uint)(fileTimeUtc >> 32)).ToString("x8") + ((uint)fileTimeUtc).ToString("x8") + Experience).ToLowerInvariant();
            byte[] input = Encoding.Unicode.GetBytes(s + "\0");
            byte[] md5; using (var m = MD5.Create()) md5 = m.ComputeHash(input);
            uint m0 = BitConverter.ToUInt32(md5, 0) | 1, m1 = BitConverter.ToUInt32(md5, 4) | 1;
            uint[][] c0 = { new uint[] { m0, 0xCF98B111u, 0x87085B9Fu, 0x12CEB96Du, 0x257E1D83u }, new uint[] { m1, 0xA27416F5u, 0xD38396FFu, 0x7C932B89u, 0xBFA49F69u } };
            uint[][] c1 = { new uint[] { m0, 0xEF0569FBu, 0x689B6B9Fu, 0x79F8A395u, 0xC3EFEA97u }, new uint[] { m1, 0xC31713DBu, 0xDDCD1F0Fu, 0x59C3AF2Du, 0x35BD1EC9u } };
            int blocks = input.Length / 8;
            uint h0 = 0, h1 = 0, a0 = 0, a1 = 0;
            unchecked {
                for (int i = 0; i < blocks; i++) for (int j = 0; j < 2; j++) {
                    uint x = BitConverter.ToUInt32(input, (i * 2 + j) * 4);
                    uint[] a = c0[j], b = c1[j];
                    h0 += x; h0 *= a[0]; h0 = Swap(h0) * a[1]; h0 = Swap(h0) * a[2]; h0 = Swap(h0) * a[3]; h0 = Swap(h0) * a[4]; a0 += h0;
                    h1 += x; h1 = Swap(h1) * b[1] + h1 * b[0]; h1 = (h1 >> 16) * b[2] + h1 * b[3]; h1 = Swap(h1) * b[4] + h1; a1 += h1;
                }
            }
            byte[] o = new byte[8];
            BitConverter.GetBytes(h0 ^ h1).CopyTo(o, 0); BitConverter.GetBytes(a0 ^ a1).CopyTo(o, 4);
            return Convert.ToBase64String(o);
        }

        // Deletes and recreates the UserChoice key with ProgId + Hash. Throws if Windows blocks the
        // write, e.g. the UCPD driver denies powershell.exe access to http/https.
        public static void Write(string assoc, string progId) {
            bool isProtocol = !assoc.StartsWith(".");
            string path = isProtocol
                ? "Software\\Microsoft\\Windows\\Shell\\Associations\\UrlAssociations\\" + assoc + "\\UserChoice"
                : "Software\\Microsoft\\Windows\\CurrentVersion\\Explorer\\FileExts\\" + assoc + "\\UserChoice";
            string sid = WindowsIdentity.GetCurrent().User.Value.ToLowerInvariant();
            for (int attempt = 0; attempt < 3; attempt++) {
                DateTime now = DateTime.UtcNow;
                DateTime minute = new DateTime(now.Year, now.Month, now.Day, now.Hour, now.Minute, 0, DateTimeKind.Utc);
                string hash = Hash(assoc, sid, progId, minute.ToFileTimeUtc());
                try { Registry.CurrentUser.DeleteSubKey(path, false); } catch { }   // key carries a deny-SetValue ACE, so recreate it
                using (RegistryKey k = Registry.CurrentUser.CreateSubKey(path)) {
                    k.SetValue("ProgId", progId, RegistryValueKind.String);
                    k.SetValue("Hash", hash, RegistryValueKind.String);
                }
                if (DateTime.UtcNow.Minute == minute.Minute) return;           // key timestamp still matches the hash
            }
        }
    }
}
'@
}

# =====================================================================================
#  XAML
# =====================================================================================
[xml]$Xaml = @'
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Post-Format Setup" Width="1280" Height="860" MinWidth="1000" MinHeight="640"
        WindowStartupLocation="Manual" WindowStyle="None" ResizeMode="CanResize"
        Background="#101318" Foreground="#E7EAEE" UseLayoutRounding="True" SnapsToDevicePixels="True"
        FontFamily="Segoe UI Variable Text, Segoe UI" FontSize="14"
        TextOptions.TextFormattingMode="Display">
  <WindowChrome.WindowChrome>
    <WindowChrome CaptionHeight="44" ResizeBorderThickness="6" GlassFrameThickness="0" CornerRadius="0" UseAeroCaptionButtons="False"/>
  </WindowChrome.WindowChrome>

  <Window.Resources>
    <!-- Tokens -->
    <SolidColorBrush x:Key="Bg"          Color="#101318"/>
    <SolidColorBrush x:Key="Surface"     Color="#171B22"/>
    <SolidColorBrush x:Key="Surface2"    Color="#1E232C"/>
    <SolidColorBrush x:Key="Field"       Color="#12161C"/>
    <SolidColorBrush x:Key="Line"        Color="#232934"/>
    <SolidColorBrush x:Key="Line2"       Color="#2F3743"/>
    <SolidColorBrush x:Key="Line3"       Color="#4A5464"/>
    <SolidColorBrush x:Key="Text"        Color="#E7EAEE"/>
    <SolidColorBrush x:Key="Muted"       Color="#8D96A3"/>
    <SolidColorBrush x:Key="Accent"      Color="#22C55E"/>
    <SolidColorBrush x:Key="AccentHover" Color="#3BD573"/>
    <SolidColorBrush x:Key="AccentPress" Color="#1AA34D"/>
    <SolidColorBrush x:Key="OnAccent"    Color="#0B1410"/>
    <SolidColorBrush x:Key="ActiveFg"    Color="#5AA9FF"/>
    <SolidColorBrush x:Key="ActiveSoft"  Color="#1A5AA9FF"/>
    <SolidColorBrush x:Key="OkSoft"      Color="#1C22C55E"/>
    <SolidColorBrush x:Key="Danger"      Color="#F07178"/>
    <SolidColorBrush x:Key="DangerSoft"  Color="#1CF07178"/>
    <SolidColorBrush x:Key="Neutral"     Color="#12FFFFFF"/>
    <SolidColorBrush x:Key="Track"       Color="#262C36"/>
    <SolidColorBrush x:Key="Ring"        Color="#9CC9FF"/>
    <FontFamily x:Key="Glyphs">Segoe Fluent Icons, Segoe MDL2 Assets</FontFamily>
    <BooleanToVisibilityConverter x:Key="BoolVis"/>

    <Style x:Key="FocusRing">
      <Setter Property="Control.Template">
        <Setter.Value>
          <ControlTemplate><Rectangle Margin="-3" RadiusX="9" RadiusY="9" Stroke="{StaticResource Ring}" StrokeThickness="2"/></ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>

    <!-- Checkbox -->
    <Style TargetType="CheckBox">
      <Setter Property="Foreground" Value="{StaticResource Text}"/>
      <Setter Property="Cursor" Value="Hand"/>
      <Setter Property="MinHeight" Value="32"/>
      <Setter Property="FocusVisualStyle" Value="{StaticResource FocusRing}"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="CheckBox">
            <Grid Background="Transparent">
              <Grid.ColumnDefinitions><ColumnDefinition Width="Auto"/><ColumnDefinition Width="*"/></Grid.ColumnDefinitions>
              <Grid Width="18" Height="18" VerticalAlignment="Center">
                <Border x:Name="Box" CornerRadius="5" BorderThickness="1.5" BorderBrush="{StaticResource Line3}" Background="Transparent"/>
                <Path x:Name="Tick" Data="M 4.5,9.5 L 7.6,12.6 L 13.5,5.8" Stroke="{StaticResource OnAccent}" StrokeThickness="2"
                      StrokeStartLineCap="Round" StrokeEndLineCap="Round" StrokeLineJoin="Round" Visibility="Collapsed"/>
                <Rectangle x:Name="Dash" Width="8" Height="2" RadiusX="1" RadiusY="1" Fill="{StaticResource OnAccent}" Visibility="Collapsed"/>
              </Grid>
              <ContentPresenter x:Name="CP" Grid.Column="1" Margin="10,0,0,0" VerticalAlignment="Center" RecognizesAccessKey="True"/>
            </Grid>
            <ControlTemplate.Triggers>
              <Trigger Property="HasContent" Value="False"><Setter TargetName="CP" Property="Margin" Value="0"/></Trigger>
              <Trigger Property="IsMouseOver" Value="True"><Setter TargetName="Box" Property="BorderBrush" Value="{StaticResource Muted}"/></Trigger>
              <Trigger Property="IsChecked" Value="True">
                <Setter TargetName="Box" Property="Background" Value="{StaticResource Accent}"/>
                <Setter TargetName="Box" Property="BorderBrush" Value="{StaticResource Accent}"/>
                <Setter TargetName="Tick" Property="Visibility" Value="Visible"/>
              </Trigger>
              <Trigger Property="IsChecked" Value="{x:Null}">
                <Setter TargetName="Box" Property="Background" Value="{StaticResource Accent}"/>
                <Setter TargetName="Box" Property="BorderBrush" Value="{StaticResource Accent}"/>
                <Setter TargetName="Dash" Property="Visibility" Value="Visible"/>
              </Trigger>
              <Trigger Property="IsEnabled" Value="False"><Setter Property="Opacity" Value="0.45"/></Trigger>
            </ControlTemplate.Triggers>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>

    <!-- Buttons -->
    <Style x:Key="BtnBase" TargetType="Button">
      <Setter Property="Cursor" Value="Hand"/>
      <Setter Property="Height" Value="40"/>
      <Setter Property="Padding" Value="18,0"/>
      <Setter Property="FontSize" Value="14"/>
      <Setter Property="FocusVisualStyle" Value="{StaticResource FocusRing}"/>
    </Style>
    <Style x:Key="Primary" TargetType="Button" BasedOn="{StaticResource BtnBase}">
      <Setter Property="Foreground" Value="{StaticResource OnAccent}"/>
      <Setter Property="FontWeight" Value="SemiBold"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="Button">
            <Border x:Name="Bd" Background="{StaticResource Accent}" CornerRadius="8" Padding="{TemplateBinding Padding}">
              <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
            </Border>
            <ControlTemplate.Triggers>
              <Trigger Property="IsMouseOver" Value="True"><Setter TargetName="Bd" Property="Background" Value="{StaticResource AccentHover}"/></Trigger>
              <Trigger Property="IsPressed" Value="True"><Setter TargetName="Bd" Property="Background" Value="{StaticResource AccentPress}"/></Trigger>
              <Trigger Property="IsEnabled" Value="False"><Setter Property="Opacity" Value="0.4"/></Trigger>
            </ControlTemplate.Triggers>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>
    <Style x:Key="Ghost" TargetType="Button" BasedOn="{StaticResource BtnBase}">
      <Setter Property="Foreground" Value="{StaticResource Text}"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="Button">
            <Border x:Name="Bd" Background="Transparent" BorderBrush="{StaticResource Line2}" BorderThickness="1" CornerRadius="8" Padding="{TemplateBinding Padding}">
              <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
            </Border>
            <ControlTemplate.Triggers>
              <Trigger Property="IsMouseOver" Value="True"><Setter TargetName="Bd" Property="Background" Value="{StaticResource Surface2}"/><Setter TargetName="Bd" Property="BorderBrush" Value="{StaticResource Line3}"/></Trigger>
              <Trigger Property="IsPressed" Value="True"><Setter TargetName="Bd" Property="Background" Value="{StaticResource Line}"/></Trigger>
              <Trigger Property="IsEnabled" Value="False"><Setter Property="Opacity" Value="0.4"/></Trigger>
            </ControlTemplate.Triggers>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>
    <Style x:Key="Chrome" TargetType="Button">
      <Setter Property="Width" Value="46"/>
      <Setter Property="Height" Value="44"/>
      <Setter Property="Foreground" Value="{StaticResource Muted}"/>
      <Setter Property="FontFamily" Value="{StaticResource Glyphs}"/>
      <Setter Property="FontSize" Value="10"/>
      <Setter Property="IsTabStop" Value="False"/>
      <Setter Property="WindowChrome.IsHitTestVisibleInChrome" Value="True"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="Button">
            <Border x:Name="Bd" Background="Transparent"><ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/></Border>
            <ControlTemplate.Triggers>
              <Trigger Property="IsMouseOver" Value="True"><Setter TargetName="Bd" Property="Background" Value="{StaticResource Surface2}"/><Setter Property="Foreground" Value="{StaticResource Text}"/></Trigger>
            </ControlTemplate.Triggers>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>
    <Style x:Key="ChromeClose" TargetType="Button" BasedOn="{StaticResource Chrome}">
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="Button">
            <Border x:Name="Bd" Background="Transparent"><ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/></Border>
            <ControlTemplate.Triggers>
              <Trigger Property="IsMouseOver" Value="True"><Setter TargetName="Bd" Property="Background" Value="#C42B1C"/><Setter Property="Foreground" Value="White"/></Trigger>
            </ControlTemplate.Triggers>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>

    <!-- Thin progress bar with an indeterminate sweep -->
    <Style x:Key="Thin" TargetType="ProgressBar">
      <Setter Property="Foreground" Value="{StaticResource ActiveFg}"/>
      <Setter Property="Height" Value="3"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="ProgressBar">
            <Grid ClipToBounds="True">
              <Border x:Name="PART_Track" Background="{StaticResource Track}" CornerRadius="2"/>
              <Border x:Name="PART_Indicator" HorizontalAlignment="Left" Background="{TemplateBinding Foreground}" CornerRadius="2"/>
              <Border x:Name="Sweep" Width="90" HorizontalAlignment="Left" Background="{TemplateBinding Foreground}" CornerRadius="2" Visibility="Collapsed">
                <Border.RenderTransform><TranslateTransform X="-90"/></Border.RenderTransform>
              </Border>
            </Grid>
            <ControlTemplate.Triggers>
              <Trigger Property="IsIndeterminate" Value="True">
                <Setter TargetName="PART_Indicator" Property="Visibility" Value="Collapsed"/>
                <Setter TargetName="Sweep" Property="Visibility" Value="Visible"/>
                <Trigger.EnterActions>
                  <BeginStoryboard x:Name="SweepSb">
                    <Storyboard RepeatBehavior="Forever">
                      <DoubleAnimation Storyboard.TargetName="Sweep" Storyboard.TargetProperty="(UIElement.RenderTransform).(TranslateTransform.X)"
                                       From="-90" To="700" Duration="0:0:1.6">
                        <DoubleAnimation.EasingFunction><SineEase EasingMode="EaseInOut"/></DoubleAnimation.EasingFunction>
                      </DoubleAnimation>
                    </Storyboard>
                  </BeginStoryboard>
                </Trigger.EnterActions>
                <Trigger.ExitActions><StopStoryboard BeginStoryboardName="SweepSb"/></Trigger.ExitActions>
              </Trigger>
            </ControlTemplate.Triggers>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>

    <!-- Slim scrollbar -->
    <Style TargetType="ScrollBar">
      <Setter Property="Width" Value="10"/>
      <Setter Property="MinWidth" Value="10"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="ScrollBar">
            <Track x:Name="PART_Track" IsDirectionReversed="True">
              <Track.Thumb>
                <Thumb>
                  <Thumb.Template>
                    <ControlTemplate TargetType="Thumb"><Border Background="{StaticResource Line3}" CornerRadius="3" Margin="3,2"/></ControlTemplate>
                  </Thumb.Template>
                </Thumb>
              </Track.Thumb>
            </Track>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>

    <!-- Tooltip -->
    <Style TargetType="ToolTip">
      <Setter Property="Foreground" Value="{StaticResource Text}"/>
      <Setter Property="FontSize" Value="12"/>
      <Setter Property="MaxWidth" Value="420"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="ToolTip">
            <Border Background="{StaticResource Surface2}" BorderBrush="{StaticResource Line2}" BorderThickness="1" CornerRadius="6" Padding="10,7">
              <TextBlock Text="{Binding Content, RelativeSource={RelativeSource TemplatedParent}}" TextWrapping="Wrap"/>
            </Border>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>

    <!-- Small text button used in category headers -->
    <Style x:Key="Link" TargetType="Button">
      <Setter Property="Foreground" Value="{StaticResource Muted}"/>
      <Setter Property="FontSize" Value="12"/>
      <Setter Property="Cursor" Value="Hand"/>
      <Setter Property="FocusVisualStyle" Value="{StaticResource FocusRing}"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="Button">
            <Border x:Name="Bd" Background="Transparent" CornerRadius="6" Padding="8,4"><ContentPresenter/></Border>
            <ControlTemplate.Triggers>
              <Trigger Property="IsMouseOver" Value="True"><Setter TargetName="Bd" Property="Background" Value="{StaticResource Line}"/><Setter Property="Foreground" Value="{StaticResource Text}"/></Trigger>
              <Trigger Property="IsEnabled" Value="False"><Setter Property="Opacity" Value="0.4"/></Trigger>
            </ControlTemplate.Triggers>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>

    <!-- Sidebar category shortcut -->
    <Style x:Key="NavItem" TargetType="Button">
      <Setter Property="Foreground" Value="{StaticResource Text}"/>
      <Setter Property="FontSize" Value="13"/>
      <Setter Property="Cursor" Value="Hand"/>
      <Setter Property="HorizontalContentAlignment" Value="Stretch"/>
      <Setter Property="FocusVisualStyle" Value="{StaticResource FocusRing}"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="Button">
            <Border x:Name="Bd" Background="Transparent" CornerRadius="6" Padding="10,7"><ContentPresenter/></Border>
            <ControlTemplate.Triggers>
              <Trigger Property="IsMouseOver" Value="True"><Setter TargetName="Bd" Property="Background" Value="{StaticResource Line}"/></Trigger>
            </ControlTemplate.Triggers>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>

    <!-- Search box -->
    <Style x:Key="Search" TargetType="TextBox">
      <Setter Property="Foreground" Value="{StaticResource Text}"/>
      <Setter Property="CaretBrush" Value="{StaticResource Text}"/>
      <Setter Property="SelectionBrush" Value="{StaticResource ActiveFg}"/>
      <Setter Property="Height" Value="40"/>
      <Setter Property="FontSize" Value="14"/>
      <Setter Property="VerticalContentAlignment" Value="Center"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="TextBox">
            <Border x:Name="Bd" Background="{StaticResource Field}" BorderBrush="{StaticResource Line2}" BorderThickness="1" CornerRadius="8">
              <Grid Margin="12,0">
                <Grid.ColumnDefinitions><ColumnDefinition Width="Auto"/><ColumnDefinition Width="*"/></Grid.ColumnDefinitions>
                <TextBlock Text="&#xE721;" FontFamily="{StaticResource Glyphs}" FontSize="13" Foreground="{StaticResource Muted}" VerticalAlignment="Center" Margin="0,0,10,0"/>
                <TextBlock x:Name="Hint" Grid.Column="1" Text="{TemplateBinding Tag}" Foreground="{StaticResource Muted}" VerticalAlignment="Center"
                           IsHitTestVisible="False" Visibility="Collapsed" Margin="2,0,0,0"/>
                <ScrollViewer x:Name="PART_ContentHost" Grid.Column="1" VerticalAlignment="Center" Focusable="False"/>
              </Grid>
            </Border>
            <ControlTemplate.Triggers>
              <Trigger Property="Text" Value=""><Setter TargetName="Hint" Property="Visibility" Value="Visible"/></Trigger>
              <Trigger Property="IsMouseOver" Value="True"><Setter TargetName="Bd" Property="BorderBrush" Value="{StaticResource Line3}"/></Trigger>
              <Trigger Property="IsKeyboardFocused" Value="True"><Setter TargetName="Bd" Property="BorderBrush" Value="{StaticResource Ring}"/></Trigger>
            </ControlTemplate.Triggers>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>

    <!-- ComboBox -->
    <Style TargetType="ComboBoxItem">
      <Setter Property="Foreground" Value="{StaticResource Text}"/>
      <Setter Property="Cursor" Value="Hand"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="ComboBoxItem">
            <Border x:Name="Bd" Background="Transparent" CornerRadius="6" Padding="10,8"><ContentPresenter/></Border>
            <ControlTemplate.Triggers>
              <Trigger Property="IsHighlighted" Value="True"><Setter TargetName="Bd" Property="Background" Value="{StaticResource Line}"/></Trigger>
              <Trigger Property="IsSelected" Value="True"><Setter Property="Foreground" Value="{StaticResource Accent}"/></Trigger>
            </ControlTemplate.Triggers>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>
    <Style TargetType="ComboBox">
      <Setter Property="Foreground" Value="{StaticResource Text}"/>
      <Setter Property="Height" Value="40"/>
      <Setter Property="Cursor" Value="Hand"/>
      <Setter Property="FocusVisualStyle" Value="{StaticResource FocusRing}"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="ComboBox">
            <Grid>
              <ToggleButton Focusable="False" ClickMode="Press"
                            IsChecked="{Binding IsDropDownOpen, Mode=TwoWay, RelativeSource={RelativeSource TemplatedParent}}">
                <ToggleButton.Template>
                  <ControlTemplate TargetType="ToggleButton">
                    <Border x:Name="Bd" Background="{StaticResource Field}" BorderBrush="{StaticResource Line2}" BorderThickness="1" CornerRadius="8">
                      <TextBlock Text="&#xE70D;" FontFamily="{StaticResource Glyphs}" FontSize="10" Foreground="{StaticResource Muted}"
                                 HorizontalAlignment="Right" VerticalAlignment="Center" Margin="0,0,14,0"/>
                    </Border>
                    <ControlTemplate.Triggers>
                      <Trigger Property="IsMouseOver" Value="True"><Setter TargetName="Bd" Property="BorderBrush" Value="{StaticResource Line3}"/></Trigger>
                    </ControlTemplate.Triggers>
                  </ControlTemplate>
                </ToggleButton.Template>
              </ToggleButton>
              <ContentPresenter IsHitTestVisible="False" Margin="14,0,36,0" VerticalAlignment="Center"
                                Content="{TemplateBinding SelectionBoxItem}" ContentTemplate="{TemplateBinding SelectionBoxItemTemplate}"/>
              <Popup x:Name="PART_Popup" Placement="Bottom" AllowsTransparency="True" Focusable="False" PopupAnimation="Fade"
                     IsOpen="{Binding IsDropDownOpen, RelativeSource={RelativeSource TemplatedParent}}">
                <Border Background="{StaticResource Surface2}" BorderBrush="{StaticResource Line2}" BorderThickness="1" CornerRadius="8"
                        Margin="0,4,0,0" Padding="4" MaxHeight="260"
                        MinWidth="{Binding ActualWidth, RelativeSource={RelativeSource TemplatedParent}}">
                  <ScrollViewer><ItemsPresenter KeyboardNavigation.DirectionalNavigation="Contained"/></ScrollViewer>
                </Border>
              </Popup>
            </Grid>
            <ControlTemplate.Triggers>
              <Trigger Property="IsEnabled" Value="False"><Setter Property="Opacity" Value="0.45"/></Trigger>
            </ControlTemplate.Triggers>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>

    <!-- One app tile (whole tile toggles selection) -->
    <DataTemplate x:Key="AppRow">
      <Border x:Name="Row" Background="#07FFFFFF" BorderBrush="{StaticResource Line}" BorderThickness="1" CornerRadius="8"
              Margin="4" Padding="10,0,10,0" Height="54" Cursor="Hand">
        <Grid>
          <Grid.ColumnDefinitions>
            <ColumnDefinition Width="32"/>
            <ColumnDefinition Width="44"/>
            <ColumnDefinition Width="*"/>
            <ColumnDefinition Width="Auto"/>
          </Grid.ColumnDefinitions>

          <CheckBox Grid.Column="0" VerticalAlignment="Center" IsChecked="{Binding IsSelected, Mode=TwoWay}"
                    IsEnabled="{Binding IsEditable}" AutomationProperties.Name="{Binding Name}"/>

          <Grid Grid.Column="1" Width="32" Height="32" HorizontalAlignment="Left" VerticalAlignment="Center">
            <Border x:Name="Mono" CornerRadius="8" Background="{Binding BrandHex}">
              <TextBlock Text="{Binding Monogram}" Foreground="White" FontWeight="Bold" FontSize="13"
                         HorizontalAlignment="Center" VerticalAlignment="Center"/>
            </Border>
            <Image Source="{Binding IconBytes}" Width="32" Height="32" RenderOptions.BitmapScalingMode="HighQuality"/>
          </Grid>

          <TextBlock Grid.Column="2" VerticalAlignment="Center" Margin="0,0,12,0" Text="{Binding Name}" FontSize="14" FontWeight="SemiBold"
                     Foreground="{StaticResource Text}" TextTrimming="CharacterEllipsis"/>

          <Border Grid.Column="3" x:Name="Pill" CornerRadius="12" Padding="9,4" MinWidth="104" VerticalAlignment="Center"
                  Background="{StaticResource Neutral}" ToolTip="{Binding Detail}">
            <StackPanel Orientation="Horizontal">
              <Ellipse x:Name="Dot" Width="7" Height="7" Fill="{StaticResource Muted}" VerticalAlignment="Center" Margin="0,0,8,0"/>
              <TextBlock x:Name="StatusText" Text="{Binding Status}" FontSize="12" FontWeight="SemiBold" Foreground="{StaticResource Muted}"/>
            </StackPanel>
          </Border>

          <ProgressBar x:Name="Bar" Grid.Column="2" Grid.ColumnSpan="2" Style="{StaticResource Thin}" VerticalAlignment="Bottom"
                       Margin="0,0,0,6" Maximum="100" Value="{Binding Progress}" IsIndeterminate="{Binding Indeterminate}" Visibility="Collapsed"/>
        </Grid>
      </Border>
      <DataTemplate.Triggers>
        <DataTrigger Binding="{Binding IsSelected}" Value="True">
          <Setter TargetName="Row" Property="BorderBrush" Value="#6622C55E"/>
          <Setter TargetName="Row" Property="Background" Value="#0D22C55E"/>
        </DataTrigger>
        <Trigger SourceName="Row" Property="IsMouseOver" Value="True"><Setter TargetName="Row" Property="Background" Value="#12FFFFFF"/></Trigger>
        <DataTrigger Binding="{Binding IsEditable}" Value="False"><Setter TargetName="Row" Property="Cursor" Value="Arrow"/></DataTrigger>
        <DataTrigger Binding="{Binding HasIcon}" Value="True"><Setter TargetName="Mono" Property="Visibility" Value="Collapsed"/></DataTrigger>
        <DataTrigger Binding="{Binding State}" Value="Active">
          <Setter TargetName="Pill" Property="Background" Value="{StaticResource ActiveSoft}"/>
          <Setter TargetName="Dot" Property="Fill" Value="{StaticResource ActiveFg}"/>
          <Setter TargetName="StatusText" Property="Foreground" Value="{StaticResource ActiveFg}"/>
          <Setter TargetName="Bar" Property="Visibility" Value="Visible"/>
        </DataTrigger>
        <DataTrigger Binding="{Binding State}" Value="Queued">
          <Setter TargetName="Dot" Property="Fill" Value="Transparent"/>
          <Setter TargetName="Dot" Property="Stroke" Value="{StaticResource Muted}"/>
          <Setter TargetName="StatusText" Property="Foreground" Value="{StaticResource Text}"/>
        </DataTrigger>
        <DataTrigger Binding="{Binding State}" Value="Success">
          <Setter TargetName="Pill" Property="Background" Value="{StaticResource OkSoft}"/>
          <Setter TargetName="Dot" Property="Fill" Value="{StaticResource Accent}"/>
          <Setter TargetName="StatusText" Property="Foreground" Value="{StaticResource Accent}"/>
        </DataTrigger>
        <DataTrigger Binding="{Binding State}" Value="Installed">
          <Setter TargetName="Dot" Property="Fill" Value="{StaticResource Accent}"/>
          <Setter TargetName="StatusText" Property="Foreground" Value="{StaticResource Text}"/>
        </DataTrigger>
        <DataTrigger Binding="{Binding State}" Value="Failed">
          <Setter TargetName="Pill" Property="Background" Value="{StaticResource DangerSoft}"/>
          <Setter TargetName="Dot" Property="Fill" Value="{StaticResource Danger}"/>
          <Setter TargetName="StatusText" Property="Foreground" Value="{StaticResource Danger}"/>
        </DataTrigger>
      </DataTemplate.Triggers>
    </DataTemplate>
  </Window.Resources>

  <Border x:Name="Frame" BorderBrush="{StaticResource Line2}" BorderThickness="1" Background="{StaticResource Bg}">
    <Grid>
      <Grid.RowDefinitions>
        <RowDefinition Height="44"/>
        <RowDefinition Height="*"/>
        <RowDefinition Height="Auto"/>
      </Grid.RowDefinitions>

      <!-- Title bar -->
      <Grid Grid.Row="0">
        <StackPanel Orientation="Horizontal" Margin="16,0,0,0" VerticalAlignment="Center">
          <Border Width="20" Height="20" CornerRadius="5" Background="{StaticResource Accent}">
            <TextBlock Text="&#xE896;" FontFamily="{StaticResource Glyphs}" FontSize="11" Foreground="{StaticResource OnAccent}"
                       HorizontalAlignment="Center" VerticalAlignment="Center"/>
          </Border>
          <TextBlock Text="Post-Format Setup" Margin="10,0,0,0" FontSize="13" FontWeight="SemiBold" Foreground="{StaticResource Muted}" VerticalAlignment="Center"/>
          <Border x:Name="DryRunTag" Visibility="Collapsed" Margin="10,0,0,0" Padding="7,1" CornerRadius="4" Background="{StaticResource ActiveSoft}" VerticalAlignment="Center">
            <TextBlock Text="DRY RUN" FontSize="11" FontWeight="SemiBold" Foreground="{StaticResource ActiveFg}"/>
          </Border>
        </StackPanel>
        <StackPanel Orientation="Horizontal" HorizontalAlignment="Right">
          <Button x:Name="MinBtn" Style="{StaticResource Chrome}" Content="&#xE921;" AutomationProperties.Name="Minimize"/>
          <Button x:Name="FillBtn" Style="{StaticResource Chrome}" Content="&#xE923;" AutomationProperties.Name="Restore down"/>
          <Button x:Name="CloseBtn" Style="{StaticResource ChromeClose}" Content="&#xE8BB;" AutomationProperties.Name="Close"/>
        </StackPanel>
      </Grid>

      <!-- Body: header, then sidebar + app grid. Only the grid and the sidebar scroll. -->
      <Grid Grid.Row="1" Margin="28,6,28,18">
        <Grid.RowDefinitions><RowDefinition Height="Auto"/><RowDefinition Height="Auto"/><RowDefinition Height="*"/></Grid.RowDefinitions>

        <Grid Grid.Row="0" Margin="0,0,0,18">
          <Grid.ColumnDefinitions><ColumnDefinition Width="*"/><ColumnDefinition Width="Auto"/></Grid.ColumnDefinitions>
          <StackPanel VerticalAlignment="Bottom" Margin="0,0,24,0">
            <TextBlock Text="Install your apps" FontSize="26" FontWeight="SemiBold" Foreground="{StaticResource Text}"/>
            <TextBlock Margin="0,6,0,0" FontSize="14" Foreground="{StaticResource Muted}" TextWrapping="Wrap"
                       Text="Pick what this PC needs. Everything installs silently through winget, one app at a time, and you can keep working while it runs."/>
          </StackPanel>
          <TextBox x:Name="SearchBox" Grid.Column="1" Style="{StaticResource Search}" Width="420" VerticalAlignment="Bottom"
                   AutomationProperties.Name="Search apps"/>
        </Grid>

        <Border x:Name="Banner" Grid.Row="1" Visibility="Collapsed" Margin="0,0,0,16" Padding="14,12" CornerRadius="8"
                Background="{StaticResource DangerSoft}" BorderBrush="#40F07178" BorderThickness="1">
          <TextBlock x:Name="BannerText" Foreground="{StaticResource Text}" TextWrapping="Wrap" FontSize="13"/>
        </Border>

        <Grid Grid.Row="2">
          <Grid.ColumnDefinitions><ColumnDefinition Width="320"/><ColumnDefinition Width="20"/><ColumnDefinition Width="*"/></Grid.ColumnDefinitions>

          <!-- Sidebar -->
          <ScrollViewer Grid.Column="0" VerticalScrollBarVisibility="Auto" HorizontalScrollBarVisibility="Disabled" Focusable="False">
            <StackPanel>
              <!-- Default browser card -->
              <Border Padding="18,16" Background="{StaticResource Surface}" BorderBrush="{StaticResource Line}" BorderThickness="1" CornerRadius="10">
                <StackPanel>
                  <TextBlock Text="Default browser" FontSize="15" FontWeight="SemiBold"/>
                  <TextBlock Margin="0,4,0,12" FontSize="12" Foreground="{StaticResource Muted}" TextWrapping="Wrap"
                             Text="Windows protects this setting. A direct registry write is tried first. If Windows blocks it, Settings opens on the browser's page and you click 'Set default' once."/>
                  <ComboBox x:Name="BrowserCombo" AutomationProperties.Name="Browser to make default"/>
                  <Button x:Name="SetDefaultBtn" Style="{StaticResource Ghost}" Margin="0,10,0,0" Content="Set as default" HorizontalAlignment="Stretch"/>
                  <CheckBox x:Name="AutoDefault" Margin="0,8,0,0" FontSize="13" IsChecked="True">
                    <TextBlock Text="Apply automatically when installation finishes" TextWrapping="Wrap"/>
                  </CheckBox>
                  <CheckBox x:Name="EnforcePolicy" FontSize="13">
                    <TextBlock Text="Also enforce with policy (asks for admin, applies at next sign-in)" TextWrapping="Wrap"/>
                  </CheckBox>
                  <TextBlock x:Name="BrowserStatus" Margin="0,8,0,0" FontSize="13" Foreground="{StaticResource Muted}" TextWrapping="Wrap"/>
                </StackPanel>
              </Border>

              <!-- Category shortcuts -->
              <Border Margin="0,16,0,0" Padding="8,12" Background="{StaticResource Surface}" BorderBrush="{StaticResource Line}" BorderThickness="1" CornerRadius="10">
                <StackPanel>
                  <TextBlock Text="Categories" FontSize="15" FontWeight="SemiBold" Margin="10,0,0,8"/>
                  <ItemsControl x:Name="CatNav" Focusable="False"/>
                </StackPanel>
              </Border>
            </StackPanel>
          </ScrollViewer>

          <!-- App grid -->
          <Border Grid.Column="2" Background="{StaticResource Surface}" BorderBrush="{StaticResource Line}" BorderThickness="1" CornerRadius="10">
            <Grid>
              <Grid.RowDefinitions><RowDefinition Height="Auto"/><RowDefinition Height="*"/></Grid.RowDefinitions>
              <Grid Margin="18,10,18,8">
                <CheckBox x:Name="SelectAll" Content="Select all" FontSize="13" VerticalAlignment="Center"/>
                <TextBlock x:Name="CountText" HorizontalAlignment="Right" VerticalAlignment="Center" FontSize="12" Foreground="{StaticResource Muted}"/>
              </Grid>
              <ScrollViewer x:Name="ListScroll" Grid.Row="1" VerticalScrollBarVisibility="Auto" HorizontalScrollBarVisibility="Disabled" Focusable="False">
                <StackPanel Margin="10,0,10,12">
                  <ItemsControl x:Name="AppList" ItemTemplate="{StaticResource AppRow}" Focusable="False" KeyboardNavigation.TabNavigation="Continue">
                    <ItemsControl.ItemsPanel>
                      <ItemsPanelTemplate>
                        <UniformGrid Columns="{Binding DataContext.Columns, RelativeSource={RelativeSource AncestorType=Window}}"/>
                      </ItemsPanelTemplate>
                    </ItemsControl.ItemsPanel>
                    <ItemsControl.GroupStyle>
                      <GroupStyle>
                        <GroupStyle.HeaderTemplate>
                          <DataTemplate>
                            <Border BorderBrush="{StaticResource Line}" BorderThickness="0,1,0,0" Margin="4,10,4,2" Padding="4,12,0,4">
                              <Grid>
                                <StackPanel Orientation="Horizontal" VerticalAlignment="Center">
                                  <TextBlock Text="{Binding Name}" FontSize="14" FontWeight="SemiBold" Foreground="{StaticResource Text}"/>
                                  <TextBlock Text="{Binding ItemCount}" FontSize="12" Foreground="{StaticResource Muted}" Margin="8,0,0,0" VerticalAlignment="Center"/>
                                </StackPanel>
                                <StackPanel Orientation="Horizontal" HorizontalAlignment="Right">
                                  <Button Style="{StaticResource Link}" Content="All"  Tag="{Binding Name}" AutomationProperties.Name="Select all in category"/>
                                  <Button Style="{StaticResource Link}" Content="None" Tag="{Binding Name}" AutomationProperties.Name="Clear category" Margin="2,0,0,0"/>
                                </StackPanel>
                              </Grid>
                            </Border>
                          </DataTemplate>
                        </GroupStyle.HeaderTemplate>
                      </GroupStyle>
                    </ItemsControl.GroupStyle>
                  </ItemsControl>
                  <TextBlock x:Name="NoResults" Visibility="Collapsed" Margin="8,18,8,22" FontSize="13" Foreground="{StaticResource Muted}"
                             Text="No apps match that search."/>
                </StackPanel>
              </ScrollViewer>
            </Grid>
          </Border>
        </Grid>
      </Grid>

      <!-- Footer -->
      <Border Grid.Row="2" BorderBrush="{StaticResource Line}" BorderThickness="0,1,0,0" Background="{StaticResource Surface}" Padding="28,14">
        <Grid>
          <Grid.ColumnDefinitions><ColumnDefinition Width="*"/><ColumnDefinition Width="Auto"/></Grid.ColumnDefinitions>
          <StackPanel VerticalAlignment="Center" Margin="0,0,16,0">
            <TextBlock Text="{Binding Summary}" FontSize="13" Foreground="{StaticResource Text}" TextTrimming="CharacterEllipsis"/>
            <ProgressBar Style="{StaticResource Thin}" Foreground="{StaticResource Accent}" Margin="0,8,0,0" Width="320" HorizontalAlignment="Left"
                         Maximum="100" Value="{Binding Progress}" Visibility="{Binding Busy, Converter={StaticResource BoolVis}}"/>
          </StackPanel>
          <StackPanel Grid.Column="1" Orientation="Horizontal">
            <Button x:Name="LogBtn" Style="{StaticResource Ghost}" Content="Open log" Margin="0,0,10,0"/>
            <Button x:Name="InstallBtn" Style="{StaticResource Primary}" MinWidth="170" Content="Install selected"/>
          </StackPanel>
        </Grid>
      </Border>
    </Grid>
  </Border>
</Window>
'@

$Window = [Windows.Markup.XamlReader]::Load((New-Object System.Xml.XmlNodeReader $Xaml))
$ui = @{}
foreach ($n in 'Frame','DryRunTag','MinBtn','FillBtn','CloseBtn','CatNav','ListScroll','Banner','BannerText','SelectAll','CountText','AppList','SearchBox','NoResults',
               'BrowserCombo','SetDefaultBtn','AutoDefault','EnforcePolicy','BrowserStatus','LogBtn','InstallBtn') {
    $ui[$n] = $Window.FindName($n)
}

# =====================================================================================
#  State
# =====================================================================================
$Session = New-Object PFI.SessionState
$Window.DataContext = $Session
$Items = New-Object 'System.Collections.ObjectModel.ObservableCollection[PFI.AppItem]'
$Meta  = @{}
foreach ($c in $Catalog) {
    $it = New-Object PFI.AppItem
    $it.Category = $c.Category; $it.Name = $c.Name; $it.Id = $c.Id; $it.BrandHex = $c.Brand; $it.Monogram = $c.Mono
    $it.IsSelected = [bool]$c.Selected
    $Items.Add($it)
    $Meta[$c.Id] = [pscustomobject]@{
        Match      = $c.Match
        Browser    = $c.Browser
        IconBase64 = $c.IconBase64
        IconUrl    = if ($c.Domain) { "https://www.google.com/s2/favicons?domain=$($c.Domain)&sz=64" }
    }
}
$ui.AppList.ItemsSource = $Items

# Group by category (catalog order is kept) and filter from the search box
$View = [System.Windows.Data.CollectionViewSource]::GetDefaultView($Items)
$View.GroupDescriptions.Add((New-Object System.Windows.Data.PropertyGroupDescription 'Category'))
$Script:SearchText = ''
$View.Filter = [Predicate[object]]{
    param($o)
    if (-not $Script:SearchText) { return $true }
    return ($o.Name.IndexOf($Script:SearchText, [StringComparison]::OrdinalIgnoreCase) -ge 0) -or
           ($o.Category.IndexOf($Script:SearchText, [StringComparison]::OrdinalIgnoreCase) -ge 0) -or
           ($o.Id.IndexOf($Script:SearchText, [StringComparison]::OrdinalIgnoreCase) -ge 0)
}

# Everything the background runspaces may touch. They only read it, and every UI change goes through [PFI.Ui].
$Sync = [hashtable]::Synchronized(@{
    Dispatcher = $Window.Dispatcher
    Session    = $Session
    Items      = @($Items)
    Meta       = $Meta
    IconDir    = $IconDir
    LogPath    = $LogPath
    DryRun     = [bool]($DryRun -or $Snapshot)
    Winget     = $null
    Queue      = @()
    Cancel     = $false
})

function Write-Log([string]$Text) {
    try { [IO.File]::AppendAllText($LogPath, ('{0:HH:mm:ss}  {1}{2}' -f (Get-Date), $Text, [Environment]::NewLine)) } catch { }
}

# =====================================================================================
#  winget discovery (fresh installs sometimes need App Installer registered first)
# =====================================================================================
function Resolve-Winget {
    $cmd = Get-Command winget.exe -ErrorAction SilentlyContinue
    if ($cmd) { return $cmd.Source }
    $alias = Join-Path $env:LOCALAPPDATA 'Microsoft\WindowsApps\winget.exe'
    if (Test-Path $alias) { return $alias }
    try {
        # A fresh Windows image ships App Installer but sometimes has not registered it for the user yet.
        Add-AppxPackage -RegisterByFamilyName -MainPackage Microsoft.DesktopAppInstaller_8wekyb3d8bbwe -ErrorAction Stop
        if (Test-Path $alias) { return $alias }
    } catch { }
    return $null
}

# =====================================================================================
#  Background runspaces
# =====================================================================================
# Shared by both workers. They run in their own runspaces and cannot see this script's functions.
$WorkerCommon = @'
function Write-Log([string]$Text) {
    try { [IO.File]::AppendAllText($sync.LogPath, ('{0:HH:mm:ss}  {1}{2}' -f (Get-Date), $Text, [Environment]::NewLine)) } catch { }
}
function ConvertTo-IconPng {
    param([byte[]]$Bytes, [string]$Path)
    try {
        if ($Path) {
            if ($Path -match '\.ico$') { $bmp = (New-Object System.Drawing.Icon($Path, 64, 64)).ToBitmap() }
            else { $bmp = ([System.Drawing.Icon]::ExtractAssociatedIcon($Path)).ToBitmap() }
        } else {
            $ms  = New-Object IO.MemoryStream(,$Bytes)
            $bmp = [System.Drawing.Image]::FromStream($ms)
        }
        if ($bmp.Width -lt 16) { $bmp.Dispose(); return $null }   # 1px or broken placeholder
        $out = New-Object IO.MemoryStream
        $bmp.Save($out, [System.Drawing.Imaging.ImageFormat]::Png); $bmp.Dispose()
        return ,$out.ToArray()
    } catch { return $null }
}
# Dark logos on transparent backgrounds (Ubisoft, EA, Zed...) vanish on the dark UI.
# If the visible pixels are dark on average, draw the logo on a light rounded tile.
function Add-IconPlate([byte[]]$Png) {
    if (-not $Png) { return $null }
    try {
        $src = New-Object System.Drawing.Bitmap (New-Object IO.MemoryStream(,$Png))
        $sum = 0.0; $n = 0; $clear = 0
        $step = [math]::Max(1, [int]($src.Width / 24))
        for ($y = 0; $y -lt $src.Height; $y += $step) {
            for ($x = 0; $x -lt $src.Width; $x += $step) {
                $c = $src.GetPixel($x, $y)
                if ($c.A -lt 100) { $clear++; continue }
                $sum += 0.2126 * $c.R + 0.7152 * $c.G + 0.0722 * $c.B; $n++
            }
        }
        if ($n -eq 0 -or $clear -eq 0 -or ($sum / $n) -gt 70) { $src.Dispose(); return ,$Png }
        $size = 64; $r = 14
        $out = New-Object System.Drawing.Bitmap $size, $size
        $g = [System.Drawing.Graphics]::FromImage($out)
        $g.SmoothingMode = 'AntiAlias'; $g.InterpolationMode = 'HighQualityBicubic'
        $path = New-Object System.Drawing.Drawing2D.GraphicsPath
        $path.AddArc(0, 0, $r * 2, $r * 2, 180, 90); $path.AddArc($size - $r * 2 - 1, 0, $r * 2, $r * 2, 270, 90)
        $path.AddArc($size - $r * 2 - 1, $size - $r * 2 - 1, $r * 2, $r * 2, 0, 90); $path.AddArc(0, $size - $r * 2 - 1, $r * 2, $r * 2, 90, 90)
        $path.CloseFigure()
        $g.FillPath((New-Object System.Drawing.SolidBrush ([System.Drawing.Color]::FromArgb(255, 232, 235, 239))), $path)
        $g.DrawImage($src, 11, 11, 42, 42)
        $g.Dispose(); $src.Dispose()
        $ms = New-Object IO.MemoryStream; $out.Save($ms, [System.Drawing.Imaging.ImageFormat]::Png); $out.Dispose()
        return ,$ms.ToArray()
    } catch { return ,$Png }
}
function Get-InstalledIconPath([string]$Pattern) {
    if (-not $Pattern) { return $null }
    $roots = 'HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall',
             'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall',
             'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall'
    foreach ($r in $roots) {
        foreach ($k in (Get-ChildItem $r -ErrorAction SilentlyContinue)) {
            $p = Get-ItemProperty $k.PSPath -ErrorAction SilentlyContinue
            if ($p.DisplayName -match $Pattern -and $p.DisplayIcon) {
                $path = ($p.DisplayIcon -replace ',\s*-?\d+$', '').Trim('"', ' ')
                if (Test-Path -LiteralPath $path) { return $path }
            }
        }
    }
    return $null
}
'@

# Worker 1: "what's already installed" scan, then icons (embedded > cache > publisher favicon > installed exe).
$ScannerScript = $WorkerCommon + @'
$d = $sync.Dispatcher

# One "winget export" lists every installed package winget knows, with exact IDs. That takes a
# couple of seconds, against a few minutes for one "winget list" per app.
if ($sync.Winget -and -not $sync.DryRun) {
    $installed = New-Object 'System.Collections.Generic.HashSet[string]' ([StringComparer]::OrdinalIgnoreCase)
    $tmp = Join-Path ([IO.Path]::GetTempPath()) ("pfi-export-{0}.json" -f [guid]::NewGuid())
    try {
        $psi = New-Object Diagnostics.ProcessStartInfo $sync.Winget, "export -o `"$tmp`" --accept-source-agreements --disable-interactivity"
        $psi.UseShellExecute = $false; $psi.CreateNoWindow = $true; $psi.RedirectStandardOutput = $true
        $p = [Diagnostics.Process]::Start($psi); $null = $p.StandardOutput.ReadToEnd(); $p.WaitForExit()
        if (Test-Path $tmp) {
            foreach ($src in (Get-Content $tmp -Raw | ConvertFrom-Json).Sources) {
                foreach ($pkg in $src.Packages) { $null = $installed.Add($pkg.PackageIdentifier) }
            }
        }
    } catch { Write-Log "installed scan failed: $_" }
    finally { Remove-Item $tmp -ErrorAction SilentlyContinue }
    foreach ($it in $sync.Items) { [PFI.Ui]::ScanResult($d, $it, $installed.Contains($it.Id)) }
}

# Icons, pass 1: embedded or cached (instant)
$missing = New-Object System.Collections.ArrayList
foreach ($it in $sync.Items) {
    $m = $sync.Meta[$it.Id]; $png = $null
    if ($m.IconBase64) { $png = ConvertTo-IconPng -Bytes ([Convert]::FromBase64String($m.IconBase64)) }
    $cache = Join-Path $sync.IconDir ($it.Id + '.png')
    if (-not $png -and (Test-Path $cache)) { $png = [IO.File]::ReadAllBytes($cache) }
    if ($png) { [PFI.Ui]::Icon($d, $it, [byte[]](Add-IconPlate $png)) } else { $null = $missing.Add($it) }
}

# Pass 2: download the rest in parallel (first run only, then they come from the cache)
try { [Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12 } catch { }
[Net.ServicePointManager]::DefaultConnectionLimit = 16
$jobs = foreach ($it in $missing) {
    $url = $sync.Meta[$it.Id].IconUrl
    if (-not $url) { continue }
    $wc = New-Object Net.WebClient
    $wc.Headers['User-Agent'] = 'PostFormatInstaller'
    @{ Item = $it; Task = $wc.DownloadDataTaskAsync($url) }
}
foreach ($j in @($jobs)) {
    $png = $null
    try { if ($j.Task.Wait(15000) -and -not $j.Task.IsFaulted) { $png = ConvertTo-IconPng -Bytes $j.Task.Result } } catch { }
    if ($png) {
        try { [IO.File]::WriteAllBytes((Join-Path $sync.IconDir ($j.Item.Id + '.png')), $png) } catch { }
        [PFI.Ui]::Icon($d, $j.Item, [byte[]](Add-IconPlate $png))
        $null = $missing.Remove($j.Item)
    }
}

# Pass 3 (offline): the installed program's own icon
foreach ($it in @($missing)) {
    $exe = Get-InstalledIconPath $sync.Meta[$it.Id].Match
    if ($exe) { $png = ConvertTo-IconPng -Path $exe; if ($png) { [PFI.Ui]::Icon($d, $it, [byte[]](Add-IconPlate $png)) } }
}
'@

# Worker 2: sequential installs. Reads winget's stdout line by line and maps it to per-app status.
$InstallerScript = $WorkerCommon + @'
$d = $sync.Dispatcher; $s = $sync.Session

function ConvertTo-Bytes([string]$n, [string]$unit) {
    $v = [double]::Parse(($n -replace ',', '.'), [Globalization.CultureInfo]::InvariantCulture)
    switch ($unit) { 'KB' { $v * 1KB } 'MB' { $v * 1MB } 'GB' { $v * 1GB } default { $v } }
}

function Invoke-FakeInstall($it) {
    [PFI.Ui]::Status($d, $it, 'Preparing...', 'Active', 0, $true, $null); Start-Sleep -Milliseconds 500
    for ($i = 0; $i -le 100; $i += 4) { [PFI.Ui]::Status($d, $it, "Downloading... $i%", 'Active', $i, $false, $null); Start-Sleep -Milliseconds 70 }
    [PFI.Ui]::Status($d, $it, 'Installing...', 'Active', 100, $true, $null); Start-Sleep -Milliseconds 1400
    return @{ Ok = $true; Text = 'Completed'; Detail = 'Dry run: nothing was installed.' }
}

function Invoke-WingetInstall($it) {
    $wgArgs = "install --id $($it.Id) --exact --source winget --silent --accept-package-agreements --accept-source-agreements"
    if ($sync.WingetVersion -ge [version]'1.4') { $wgArgs += ' --disable-interactivity' }
    Write-Log "winget $wgArgs"

    $psi = New-Object Diagnostics.ProcessStartInfo $sync.Winget, $wgArgs
    $psi.UseShellExecute = $false; $psi.CreateNoWindow = $true
    $psi.RedirectStandardOutput = $true
    $psi.StandardOutputEncoding = [Text.Encoding]::UTF8

    [PFI.Ui]::Status($d, $it, 'Preparing...', 'Active', 0, $true, '')
    $p = [Diagnostics.Process]::Start($psi)

    # winget redraws its progress bar with '\r'. ReadLine() splits on '\r' too, so each redraw
    # arrives as its own "line".
    # winget's text is localized. Byte counts ("12.3 MB / 80 MB") are not, so the phase change is
    # inferred from them as well as from the English keywords.
    $phase = 'Preparing'; $lastPct = -1; $tail = New-Object Collections.Generic.List[string]
    while ($null -ne ($line = $p.StandardOutput.ReadLine())) {
        $t = ($line -replace '[▀-▟]', ' ').Trim()
        if ($t.Length -eq 0 -or $t -match '^[-\\|/]+$') { continue }

        $pct = $null
        if ($t -match '([\d.,]+)\s*(B|KB|MB|GB)\s*/\s*([\d.,]+)\s*(B|KB|MB|GB)') {
            $total = ConvertTo-Bytes $Matches[3] $Matches[4]
            if ($total -gt 0) { $pct = [int][math]::Min(100, [math]::Floor((ConvertTo-Bytes $Matches[1] $Matches[2]) / $total * 100)) }
        } elseif ($t -match '^(\d{1,3})\s*%$') { $pct = [int]$Matches[1] }

        if ($null -ne $pct) {
            if ($phase -eq 'Installing') { continue }   # some installers report their own percentage
            $phase = 'Downloading'
            if ($pct -ne $lastPct) { $lastPct = $pct; [PFI.Ui]::Status($d, $it, "Downloading... $pct%", 'Active', $pct, $false, $null) }
            continue
        }

        Write-Log "  $t"; $tail.Add($t); if ($tail.Count -gt 6) { $tail.RemoveAt(0) }

        if ($t -match '^Download') {
            if ($phase -eq 'Preparing') { $phase = 'Downloading'; [PFI.Ui]::Status($d, $it, 'Downloading...', 'Active', 0, $true, $null) }
        } elseif ($t -match 'Starting package install|Installing|verified installer hash' -or ($phase -eq 'Downloading' -and $lastPct -ge 0)) {
            if ($phase -ne 'Installing') { $phase = 'Installing'; [PFI.Ui]::Status($d, $it, 'Installing...', 'Active', 100, $true, $null) }
        }
    }
    $p.WaitForExit()

    $code = '0x{0:X8}' -f [BitConverter]::ToUInt32([BitConverter]::GetBytes([int]$p.ExitCode), 0)
    $detail = (($tail | Select-Object -Last 3) -join "`n")
    Write-Log "  exit $code"
    switch ($code) {
        '0x00000000' { return @{ Ok = $true;  Text = 'Completed';       Detail = $detail } }
        '0x8A15002B' { return @{ Ok = $true;  Text = 'Up to date';      Detail = 'Already installed, no newer version available.' } }
        '0x8A150061' { return @{ Ok = $true;  Text = 'Already installed'; Detail = $detail } }
        '0x8A150109' { return @{ Ok = $true;  Text = 'Restart needed';  Detail = 'Installed. Windows needs a restart to finish.' } }
        '0x8A150014' { return @{ Ok = $false; Text = 'Not found';       Detail = "No winget package with ID '$($it.Id)'." } }
        default      { return @{ Ok = $false; Text = 'Failed';          Detail = "winget exit code $code`n$detail`nFull log: $($sync.LogPath)" } }
    }
}

$queue = $sync.Queue; $total = $queue.Count; $n = 0; $ok = 0; $bad = 0
try {
    if (-not $sync.DryRun) {
        $v = (& $sync.Winget --version 2>$null) -replace '[^\d.]', ''
        $sync.WingetVersion = try { [version]$v } catch { [version]'1.0' }
        Write-Log "winget $($sync.WingetVersion)"
    }
    foreach ($it in $queue) {
        if ($sync.Cancel) { [PFI.Ui]::Status($d, $it, 'Skipped', 'Skipped', 0, $false, 'Stopped before this app started.'); continue }
        [PFI.Ui]::Summary($d, $s, ("Installing {0} ({1} of {2})" -f $it.Name, ($n + 1), $total), ($n / $total) * 100)
        try {
            $r = if ($sync.DryRun) { Invoke-FakeInstall $it } else { Invoke-WingetInstall $it }
        } catch {
            $r = @{ Ok = $false; Text = 'Failed'; Detail = $_.Exception.Message }
            Write-Log "  error: $($_.Exception.Message)"
        }
        if ($r.Ok) { $ok++ } else { $bad++ }
        [PFI.Ui]::Status($d, $it, $r.Text, $(if ($r.Ok) { 'Success' } else { 'Failed' }), 100, $false, $r.Detail)
        $n++
        [PFI.Ui]::Summary($d, $s, $null, ($n / $total) * 100)
    }
    $msg = "Finished: $ok of $total installed"
    if ($bad)         { $msg += ", $bad failed (hover a status for details)" }
    if ($sync.Cancel) { $msg += ', stopped early' }
    [PFI.Ui]::Summary($d, $s, $msg, 100)
    Write-Log $msg
} catch {
    Write-Log "worker crashed: $_"
    [PFI.Ui]::Summary($d, $s, "Something went wrong: $($_.Exception.Message)", -1)
} finally {
    [PFI.Ui]::Busy($d, $s, $false)    # always hand control back to the UI
}
'@

function Start-Worker([string]$Script) {
    $rs = [runspacefactory]::CreateRunspace()
    $rs.ApartmentState = 'MTA'
    $rs.Open()
    $rs.SessionStateProxy.SetVariable('sync', $Sync)
    $ps = [powershell]::Create(); $ps.Runspace = $rs
    $null = $ps.AddScript($Script)
    return @{ PS = $ps; Handle = $ps.BeginInvoke() }
}
function Stop-Worker($w) {
    if (-not $w) { return }
    try { if ($w.Handle.IsCompleted) { $null = $w.PS.EndInvoke($w.Handle) } else { $w.PS.Stop() } } catch { }
    try { $w.PS.Runspace.Dispose(); $w.PS.Dispose() } catch { }
}

# =====================================================================================
#  Selection
# =====================================================================================
$Script:Bulk = $false   # suppresses per-item refreshes while many checkboxes change at once

function Get-VisibleItems { @($Items | Where-Object { $View.PassesFilter($_) }) }

function Update-Selection {
    if ($Script:Bulk) { return }
    $sel = 0; foreach ($it in $Items) { if ($it.IsSelected) { $sel++ } }
    $visible = Get-VisibleItems
    $visSel = @($visible | Where-Object IsSelected).Count
    $ui.SelectAll.IsChecked = if ($visSel -eq 0) { $false } elseif ($visSel -eq $visible.Count) { $true } else { $null }
    $ui.SelectAll.Content = if ($Script:SearchText) { "Select all $($visible.Count) shown" } else { 'Select all' }
    $ui.CountText.Text = "$sel of $($Items.Count) selected"
    $ui.NoResults.Visibility = if ($visible.Count) { 'Collapsed' } else { 'Visible' }
    if (Get-Command Update-CategoryCounts -ErrorAction SilentlyContinue) { Update-CategoryCounts }
    if (-not $Session.Busy) {
        $ui.InstallBtn.IsEnabled = ($sel -gt 0) -and ([bool]$Sync.Winget -or $Sync.DryRun)
        $ui.InstallBtn.Content   = if ($sel -eq 1) { 'Install 1 app' } elseif ($sel -gt 1) { "Install $sel apps" } else { 'Install selected' }
        if (-not $Script:Finished) { $Session.Summary = if ($sel) { 'Ready when you are.' } else { 'Select at least one app.' } }
    }
}
foreach ($it in $Items) {
    $it.add_PropertyChanged({ param($s, $e) if ($e.PropertyName -eq 'IsSelected') { Update-Selection } })
}
function Set-Selected($Targets, [bool]$Value) {
    $Script:Bulk = $true
    try { foreach ($it in $Targets) { if ($it.IsEditable) { $it.IsSelected = $Value } } }
    finally { $Script:Bulk = $false }
    Update-Selection
}

# "Select all" acts on what the search currently shows
$ui.SelectAll.Add_Click({ Set-Selected (Get-VisibleItems) ($ui.SelectAll.IsChecked -eq $true) })

# Category header "All" / "None" buttons (one routed handler for every group header)
$ui.AppList.AddHandler([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent, [System.Windows.RoutedEventHandler]{
    param($sender, $e)
    $b = $e.OriginalSource
    if ($b -isnot [System.Windows.Controls.Button] -or -not ($b.Tag -is [string])) { return }
    if ($Session.Busy) { return }
    $inGroup = Get-VisibleItems | Where-Object { $_.Category -eq $b.Tag }
    Set-Selected $inGroup ([string]$b.Content -eq 'All')
})

# Clicking anywhere on a tile toggles it. The checkbox and buttons mark their own clicks
# as handled, so those never toggle twice.
$ui.AppList.AddHandler([System.Windows.UIElement]::MouseLeftButtonUpEvent, [System.Windows.Input.MouseButtonEventHandler]{
    param($sender, $e)
    $it = $e.OriginalSource.DataContext
    if ($it -is [PFI.AppItem] -and $it.IsEditable) { $it.IsSelected = -not $it.IsSelected; $e.Handled = $true }
})

# Sidebar category shortcuts: jump to a group and show "selected / total" for it
$Script:CatCounts = @{}
foreach ($cat in @($Catalog | ForEach-Object { $_.Category } | Select-Object -Unique)) {
    $btn = New-Object System.Windows.Controls.Button
    $btn.Style = $Window.FindResource('NavItem'); $btn.Tag = $cat
    [System.Windows.Automation.AutomationProperties]::SetName($btn, "Jump to $cat")
    $grid = New-Object System.Windows.Controls.Grid
    $name = New-Object System.Windows.Controls.TextBlock; $name.Text = $cat
    $count = New-Object System.Windows.Controls.TextBlock
    $count.HorizontalAlignment = 'Right'; $count.FontSize = 12; $count.VerticalAlignment = 'Center'
    $null = $grid.Children.Add($name); $null = $grid.Children.Add($count)
    $btn.Content = $grid
    $Script:CatCounts[$cat] = $count
    $btn.Add_Click({
        param($s, $e)
        $cat = [string]$s.Tag
        if ($ui.SearchBox.Text) { $ui.SearchBox.Clear() }
        $ui.AppList.UpdateLayout()
        $first = $Items | Where-Object { $_.Category -eq $cat } | Select-Object -First 1
        $c = $ui.AppList.ItemContainerGenerator.ContainerFromItem($first)
        if ($c) {
            $y = $c.TransformToAncestor($ui.ListScroll.Content).Transform((New-Object System.Windows.Point 0, 0)).Y
            $ui.ListScroll.ScrollToVerticalOffset([math]::Max(0, $y - 64))
        }
    })
    $null = $ui.CatNav.Items.Add($btn)
}
function Update-CategoryCounts {
    foreach ($cat in $Script:CatCounts.Keys) {
        $all = @($Items | Where-Object { $_.Category -eq $cat })
        $sel = @($all | Where-Object IsSelected).Count
        $Script:CatCounts[$cat].Text = if ($sel) { "$sel / $($all.Count)" } else { "$($all.Count)" }
        $Script:CatCounts[$cat].Foreground = $Window.FindResource($(if ($sel) { 'Accent' } else { 'Muted' }))
    }
}

$ui.SearchBox.Tag = "Search $($Items.Count) apps by name or category"
$ui.SearchBox.Add_TextChanged({
    $Script:SearchText = $ui.SearchBox.Text.Trim()
    $View.Refresh()
    Update-Selection
})
$ui.SearchBox.Add_KeyDown({ param($s, $e) if ($e.Key -eq 'Escape') { $ui.SearchBox.Clear() } })

# =====================================================================================
#  Install flow
# =====================================================================================
$Script:InstallWorker = $null
$Script:Finished = $false

$ui.InstallBtn.Add_Click({
    if ($Session.Busy) {
        $Sync.Cancel = $true
        $ui.InstallBtn.Content = 'Stopping...'; $ui.InstallBtn.IsEnabled = $false
        $Session.Summary = 'Stopping after the current app finishes...'
        return
    }
    $queue = @($Items | Where-Object IsSelected)
    if (-not $queue) { return }
    foreach ($it in $Items) { $it.IsEditable = $false }
    foreach ($it in $queue) { $it.State = 'Queued'; $it.Status = 'Queued'; $it.Progress = 0; $it.Indeterminate = $false; $it.Detail = $null }
    $Sync.Queue = $queue; $Sync.Cancel = $false
    $Session.Progress = 0
    $Session.Busy = $true
    Write-Log "Install started: $(($queue | ForEach-Object Id) -join ', ')"
    $Script:InstallWorker = Start-Worker $InstallerScript
})

$Session.add_PropertyChanged({
    param($s, $e)
    if ($e.PropertyName -ne 'Busy') { return }
    if ($Session.Busy) {
        $ui.InstallBtn.Content = 'Stop after current'
        $ui.SelectAll.IsEnabled = $false
        return
    }
    # Runs on the UI thread. The worker posted Busy=false as its last action.
    Stop-Worker $Script:InstallWorker; $Script:InstallWorker = $null
    $Script:Finished = $true
    foreach ($it in $Items) { $it.IsEditable = $true }
    $ui.SelectAll.IsEnabled = $true
    $ui.InstallBtn.IsEnabled = $true
    Update-Selection
    # Offline machines: once an app is installed, use its own exe icon if no icon was found earlier
    foreach ($it in $Items) {
        if (-not $it.HasIcon -and $it.State -eq 'Success') {
            $exe = Get-InstalledIconPathUi $Meta[$it.Id].Match
            if ($exe) {
                try {
                    $bmp = [System.Drawing.Icon]::ExtractAssociatedIcon($exe).ToBitmap()
                    $ms = New-Object IO.MemoryStream; $bmp.Save($ms, [System.Drawing.Imaging.ImageFormat]::Png); $bmp.Dispose()
                    $it.IconBytes = $ms.ToArray()
                } catch { }
            }
        }
    }
    Update-BrowserList
    if ($ui.AutoDefault.IsChecked -and $ui.BrowserCombo.SelectedItem) {
        $pick = $Script:BrowserMap[[string]$ui.BrowserCombo.SelectedItem]
        $installedNow = $Items | Where-Object { $_.State -eq 'Success' -and $Meta[$_.Id].Browser -eq $pick }
        if ($installedNow -and -not $Sync.DryRun) { Set-DefaultBrowser $pick }
    }
})

function Get-InstalledIconPathUi([string]$Pattern) {
    foreach ($r in 'HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall','HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall','HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall') {
        foreach ($k in (Get-ChildItem $r -ErrorAction SilentlyContinue)) {
            $p = Get-ItemProperty $k.PSPath -ErrorAction SilentlyContinue
            if ($Pattern -and $p.DisplayName -match $Pattern -and $p.DisplayIcon) {
                $path = ($p.DisplayIcon -replace ',\s*-?\d+$', '').Trim('"', ' ')
                if ($path -notmatch '\.ico$' -and (Test-Path -LiteralPath $path)) { return $path }
            }
        }
    }
}

# =====================================================================================
#  DEFAULT BROWSER
# =====================================================================================
#  Why this is hard, and what this tool does about it:
#
#  * Windows 8 and later store the http/https choice in
#      HKCU\...\Shell\Associations\UrlAssociations\<proto>\UserChoice  (ProgId + Hash)
#    The Hash is a salted, undocumented checksum. If it does not match, Windows ignores the key and
#    resets the default to Edge with an "app default was reset" notification. [PFI.UserChoice]
#    reproduces it, which was checked against hashes Windows wrote.
#
#  * Windows 10/11 updates since early 2024 add the "UCPD" kernel driver (User Choice Protection
#    Driver). It blocks writes to the http/https/.pdf UserChoice keys from processes that are not
#    allow-listed. powershell.exe, reg.exe and regedit are NOT allow-listed, so on a current build
#    step 1 below usually fails with "access denied". Disabling UCPD needs admin rights plus
#    turning off the "UCPD velocity" scheduled task and rebooting. This tool deliberately does
#    not do that.
#
#  * Newer Windows 11 builds also write a "UserChoiceLatest" key with a different, unpublished
#    hash. Once it exists, the legacy UserChoice key is no longer authoritative, so step 1 is skipped.
#
#  Strategy, from most to least automatic:
#    1. Legacy UserChoice + hash write for http, https, .htm, .html, then verify with
#       AssocQueryString (what Windows will actually launch).
#    2. If that is blocked or ignored: open Settings directly on the browser's page with
#       ms-settings:defaultapps?registeredAppUser=<name> (or registeredAppMachine). On Windows 11
#       (22H2 with the late-2023 update and later) that page has a one-click "Set default" button.
#       The tool polls and confirms the moment you click it.
#    3. Optional, needs admin: a Default Associations policy (DefaultAssociationsConfiguration XML),
#       the Microsoft-supported enterprise method. Applied at each sign-in. Because it re-applies
#       every sign-in, it also locks the choice. Officially supported on Pro/Enterprise/Education.
# =====================================================================================
$Script:BrowserMap = @{}
$Script:PollTimer  = $null

function Get-InstalledBrowsers {
    $found = [ordered]@{}
    foreach ($h in @(@{ Root = [Microsoft.Win32.Registry]::CurrentUser; Scope = 'User' }, @{ Root = [Microsoft.Win32.Registry]::LocalMachine; Scope = 'Machine' })) {
        $smi = $h.Root.OpenSubKey('SOFTWARE\Clients\StartMenuInternet')
        if (-not $smi) { continue }
        $ra = $h.Root.OpenSubKey('SOFTWARE\RegisteredApplications')
        foreach ($k in $smi.GetSubKeyNames()) {
            $cap = $smi.OpenSubKey("$k\Capabilities"); if (-not $cap) { continue }
            $urls = $cap.OpenSubKey('URLAssociations'); if (-not $urls) { continue }
            $progId = $urls.GetValue('https'); if (-not $progId) { continue }
            $files  = $cap.OpenSubKey('FileAssociations')
            $name = $smi.OpenSubKey($k).GetValue('')
            if (-not $name) { $name = $cap.GetValue('ApplicationName') }
            if (-not $name -or $name -like '@*') { $name = $k }
            if ($name -eq 'Internet Explorer' -or $k -eq 'IEXPLORE.EXE') { continue }
            $capPath = "Software\Clients\StartMenuInternet\$k\Capabilities"
            $regName = $null
            if ($ra) { foreach ($v in $ra.GetValueNames()) { if ([string]$ra.GetValue($v) -ieq $capPath) { $regName = $v } } }
            if (-not $found.Contains($name)) {
                $found[$name] = [pscustomobject]@{
                    Name = $name; ProgId = $progId; RegName = $regName; Scope = $h.Scope
                    HtmlProgId = if ($files) { $files.GetValue('.html') } else { $null }
                    HtmProgId  = if ($files) { $files.GetValue('.htm') }  else { $null }
                }
            }
        }
    }
    return @($found.Values)
}

function Find-Browser([string]$Name) {
    $all = Get-InstalledBrowsers
    $hit = $all | Where-Object Name -eq $Name | Select-Object -First 1
    # Registry names differ from catalog names, e.g. "Mozilla Firefox" or "Opera GX Stable"
    if (-not $hit) { $hit = $all | Where-Object { $_.Name -like "$Name*" } | Select-Object -First 1 }
    if (-not $hit) { $hit = $all | Where-Object { $_.Name -like "*$Name*" -or $Name -like "*$($_.Name)*" } | Select-Object -First 1 }
    return $hit
}

function Update-BrowserList {
    $prev = [string]$ui.BrowserCombo.SelectedItem
    $Script:BrowserMap = [ordered]@{}
    $current = [PFI.Native]::ProgIdForProtocol('https')
    foreach ($b in Get-InstalledBrowsers) {
        $label = if ($b.ProgId -eq $current) { "$($b.Name)  (current default)" } else { $b.Name }
        $Script:BrowserMap[$label] = $b.Name
    }
    foreach ($c in ($Catalog | Where-Object { $_.Browser })) {
        if (-not ($Script:BrowserMap.Values -contains $c.Browser) -and -not (Find-Browser $c.Browser)) {
            $Script:BrowserMap["$($c.Browser)  (installs first)"] = $c.Browser
        }
    }
    $ui.BrowserCombo.ItemsSource = @($Script:BrowserMap.Keys)
    $want = $Catalog | Where-Object { $_.Browser } | Select-Object -First 1
    $match = @($Script:BrowserMap.Keys) | Where-Object { $_ -eq $prev } | Select-Object -First 1
    if (-not $match -and $want) { $match = @($Script:BrowserMap.Keys) | Where-Object { $Script:BrowserMap[$_] -eq $want.Browser } | Select-Object -First 1 }
    if (-not $match) { $match = @($Script:BrowserMap.Keys) | Select-Object -First 1 }
    $ui.BrowserCombo.SelectedItem = $match
}

function Set-BrowserStatus([string]$Text, [string]$Tone = 'muted') {
    $ui.BrowserStatus.Text = $Text
    $ui.BrowserStatus.Foreground = $Window.FindResource($(switch ($Tone) { 'ok' { 'Accent' } 'bad' { 'Danger' } 'busy' { 'ActiveFg' } default { 'Muted' } }))
}

function Set-DefaultBrowser([string]$Name) {
    if ($Script:PollTimer) { $Script:PollTimer.Stop(); $Script:PollTimer = $null }
    $b = Find-Browser $Name
    if (-not $b) { Set-BrowserStatus "$Name isn't installed yet. Install it first, or leave 'Apply automatically' on." 'bad'; return }
    Write-Log "Default browser -> $($b.Name) ($($b.ProgId), regName=$($b.RegName), scope=$($b.Scope))"

    if ($ui.EnforcePolicy.IsChecked) { Invoke-DefaultAssocPolicy $b }

    if ([PFI.Native]::ProgIdForProtocol('https') -eq $b.ProgId) {
        Set-BrowserStatus "$($b.Name) is already your default browser." 'ok'; Update-BrowserList; return
    }

    # Step 1: legacy UserChoice + hash (works only where UCPD/UserChoiceLatest don't interfere)
    $latest = Test-Path 'HKCU:\Software\Microsoft\Windows\Shell\Associations\UrlAssociations\https\UserChoiceLatest'
    if (-not $latest) {
        $pairs = @(@('http', $b.ProgId), @('https', $b.ProgId), @('.htm', $b.HtmProgId), @('.html', $b.HtmlProgId))
        foreach ($pair in $pairs) {
            if (-not $pair[1]) { continue }
            try { [PFI.UserChoice]::Write($pair[0], $pair[1]); Write-Log "  UserChoice $($pair[0]) written" }
            catch { Write-Log "  UserChoice $($pair[0]) blocked: $($_.Exception.InnerException.Message)" }
        }
        [PFI.Native]::NotifyAssocChanged()
        if ([PFI.Native]::ProgIdForProtocol('https') -eq $b.ProgId) {
            Set-BrowserStatus "$($b.Name) is now your default browser." 'ok'; Update-BrowserList; return
        }
    } else {
        Write-Log '  UserChoiceLatest present: skipping legacy hash write'
    }

    # Step 2: the one-click Settings page, then confirm by polling.
    $uri = if ($b.RegName) {
        $param = if ($b.Scope -eq 'User') { 'registeredAppUser' } else { 'registeredAppMachine' }
        "ms-settings:defaultapps?$param=$([Uri]::EscapeDataString($b.RegName))"
    } else { 'ms-settings:defaultapps' }
    Write-Log "  opening $uri"
    Start-Process $uri
    Set-BrowserStatus "Windows needs one click: in the Settings window, press 'Set default' at the top of the $($b.Name) page. This will update when you do." 'busy'

    $Script:PollDeadline = (Get-Date).AddMinutes(3)
    $Script:PollTarget = $b
    $Script:PollTimer = New-Object Windows.Threading.DispatcherTimer
    $Script:PollTimer.Interval = [TimeSpan]::FromMilliseconds(800)
    $Script:PollTimer.Add_Tick({
        $target = $Script:PollTarget
        if ([PFI.Native]::ProgIdForProtocol('https') -eq $target.ProgId) {
            $Script:PollTimer.Stop(); Set-BrowserStatus "$($target.Name) is now your default browser." 'ok'; Update-BrowserList
            Write-Log '  confirmed via Settings'
        } elseif ((Get-Date) -gt $Script:PollDeadline) {
            $Script:PollTimer.Stop(); Set-BrowserStatus "Still not the default. Press 'Set as default' to try again." 'bad'
        }
    })
    $Script:PollTimer.Start()
}

function Invoke-DefaultAssocPolicy($b) {
    $assoc = @(
        @{ Id = 'http';  P = $b.ProgId }, @{ Id = 'https'; P = $b.ProgId },
        @{ Id = '.htm';  P = $b.HtmProgId }, @{ Id = '.html'; P = $b.HtmlProgId }
    ) | Where-Object { $_.P } | ForEach-Object {
        '  <Association Identifier="{0}" ProgId="{1}" ApplicationName="{2}" />' -f $_.Id, [Security.SecurityElement]::Escape($_.P), [Security.SecurityElement]::Escape($b.Name)
    }
    $xml = "<?xml version=`"1.0`" encoding=`"UTF-8`"?>`r`n<DefaultAssociations>`r`n$($assoc -join "`r`n")`r`n</DefaultAssociations>"
    $xmlB64 = [Convert]::ToBase64String([Text.Encoding]::UTF8.GetBytes($xml))
    $cmd = @"
`$dir = Join-Path `$env:ProgramData 'PostFormatInstaller'
New-Item -ItemType Directory -Force -Path `$dir | Out-Null
`$file = Join-Path `$dir 'DefaultAssociations.xml'
[IO.File]::WriteAllBytes(`$file, [Convert]::FromBase64String('$xmlB64'))
New-Item -Path 'HKLM:\SOFTWARE\Policies\Microsoft\Windows\System' -Force | Out-Null
Set-ItemProperty -Path 'HKLM:\SOFTWARE\Policies\Microsoft\Windows\System' -Name 'DefaultAssociationsConfiguration' -Value `$file
"@
    $enc = [Convert]::ToBase64String([Text.Encoding]::Unicode.GetBytes($cmd))
    try {
        # Not -Wait: the UAC prompt must not freeze the UI.
        Start-Process powershell.exe -Verb RunAs -WindowStyle Hidden -ArgumentList "-NoProfile -ExecutionPolicy Bypass -EncodedCommand $enc" | Out-Null
        Write-Log '  policy: elevated writer launched'
    } catch {
        Write-Log "  policy: cancelled ($($_.Exception.Message))"
        Set-BrowserStatus 'Policy step skipped: the admin prompt was cancelled.' 'bad'
    }
}

$ui.SetDefaultBtn.Add_Click({
    $label = [string]$ui.BrowserCombo.SelectedItem
    if ($label) { Set-DefaultBrowser $Script:BrowserMap[$label] }
})

# =====================================================================================
#  Window chrome + lifecycle
# =====================================================================================
$ui.MinBtn.Add_Click({ $Window.WindowState = 'Minimized' })

# Opens filling the screen's work area. That uses the whole screen but keeps the taskbar visible;
# WindowState=Maximized on a borderless window would cover the taskbar.
$Script:Filled = $false
function Set-Filled([bool]$Fill) {
    $Window.WindowState = 'Normal'
    $wa = [System.Windows.SystemParameters]::WorkArea
    if ($Fill) {
        $Window.Left = $wa.Left; $Window.Top = $wa.Top; $Window.Width = $wa.Width; $Window.Height = $wa.Height
        $ui.FillBtn.Content = [string][char]0xE923; [System.Windows.Automation.AutomationProperties]::SetName($ui.FillBtn, 'Restore down')
    } else {
        $Window.Width = [math]::Min(1280, $wa.Width); $Window.Height = [math]::Min(860, $wa.Height)
        $Window.Left = $wa.Left + ($wa.Width - $Window.Width) / 2; $Window.Top = $wa.Top + ($wa.Height - $Window.Height) / 2
        $ui.FillBtn.Content = [string][char]0xE922; [System.Windows.Automation.AutomationProperties]::SetName($ui.FillBtn, 'Fill screen')
    }
    $Script:Filled = $Fill
}
$ui.FillBtn.Add_Click({ Set-Filled (-not ($Script:Filled -or $Window.WindowState -eq 'Maximized')) })
if ($SnapshotSize -match '^(\d+)x(\d+)$') { $Window.Width = [int]$Matches[1]; $Window.Height = [int]$Matches[2]; $Window.Left = 0; $Window.Top = 0 }
else { Set-Filled $true }

# More columns of tiles as the window gets wider (one per ~330px)
$ui.ListScroll.Add_SizeChanged({
    $w = $ui.ListScroll.ActualWidth - 30
    $cols = [math]::Max(1, [math]::Floor($w / 330))
    if ($Session.Columns -ne $cols) { $Session.Columns = $cols }
})
$ui.CloseBtn.Add_Click({ $Window.Close() })
$ui.LogBtn.Add_Click({
    if (Test-Path $LogPath) { Start-Process notepad.exe -ArgumentList "`"$LogPath`"" }
    else { $Session.Summary = 'Nothing logged yet. The log starts with the first install.' }
})
$Window.Add_StateChanged({ $ui.Frame.Margin = New-Object Windows.Thickness ($(if ($Window.WindowState -eq 'Maximized') { 7 } else { 0 })) })
$Window.Add_SourceInitialized({
    [PFI.Native]::RoundCorners((New-Object System.Windows.Interop.WindowInteropHelper $Window).Handle)
})
$Window.Add_Closing({
    param($s, $e)
    if ($Session.Busy -and -not $Snapshot) {
        $r = [Windows.MessageBox]::Show($Window, "An installation is still running. If you close now, the current app keeps installing in the background and the rest of the queue is skipped.`n`nClose anyway?", 'Post-Format Setup', 'YesNo', 'Warning')
        if ($r -ne 'Yes') { $e.Cancel = $true; return }
        $Sync.Cancel = $true
    }
    if ($Script:PollTimer) { $Script:PollTimer.Stop() }
})

$Window.Add_ContentRendered({
    $Sync.Winget = Resolve-Winget
    if ($Sync.DryRun) {
        $ui.DryRunTag.Visibility = 'Visible'
    } elseif (-not $Sync.Winget) {
        $ui.Banner.Visibility = 'Visible'
        $ui.BannerText.Text = "winget wasn't found. Open Microsoft Store, update 'App Installer', then restart this tool. On Windows 10, run Windows Update first."
    }
    if ($Sync.Winget -and -not $Sync.DryRun) {
        foreach ($it in $Items) { $it.State = 'Checking'; $it.Status = 'Checking...' }
    }
    Update-Selection
    Update-BrowserList
    $Script:ScanWorker = Start-Worker $ScannerScript

    if ($Snapshot) {
        $ui.InstallBtn.RaiseEvent((New-Object Windows.RoutedEventArgs ([Windows.Controls.Primitives.ButtonBase]::ClickEvent)))
        $Script:SnapTimer = New-Object Windows.Threading.DispatcherTimer
        $Script:SnapTimer.Interval = [TimeSpan]::FromSeconds(4.2)
        $Script:SnapTimer.Add_Tick({
            $Script:SnapTimer.Stop()
            function Save-Visual($Visual, [string]$Path) {
                $rtb = New-Object Windows.Media.Imaging.RenderTargetBitmap ([int]$Visual.ActualWidth), ([int]$Visual.ActualHeight), 96, 96, ([Windows.Media.PixelFormats]::Pbgra32)
                $bg = New-Object Windows.Media.DrawingVisual; $dc = $bg.RenderOpen()
                $dc.DrawRectangle($Window.FindResource('Bg'), $null, (New-Object Windows.Rect 0, 0, $Visual.ActualWidth, $Visual.ActualHeight)); $dc.Close()
                $rtb.Render($bg); $rtb.Render($Visual)
                $enc = New-Object Windows.Media.Imaging.PngBitmapEncoder
                $enc.Frames.Add([Windows.Media.Imaging.BitmapFrame]::Create($rtb))
                $fs = [IO.File]::Create($Path); $enc.Save($fs); $fs.Close()
            }
            Save-Visual $Window.Content $Snapshot
            # Second image: the whole scrollable body, to review every row
            Save-Visual $ui.ListScroll.Content ([IO.Path]::ChangeExtension($Snapshot, '.full.png'))
            $Window.Close()
        })
        $Script:SnapTimer.Start()
    }
})

Write-Log "Post-Format Setup started (PowerShell $($PSVersionTable.PSVersion), dry run: $($Sync.DryRun))"
$null = $Window.ShowDialog()

Stop-Worker $Script:InstallWorker
Stop-Worker $Script:ScanWorker

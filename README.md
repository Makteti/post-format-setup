# Post-Format Setup

[![Free and open source](https://img.shields.io/badge/free-open%20source-22C55E)](LICENSE)
[![VirusTotal scan report](https://img.shields.io/badge/VirusTotal-scan%20report-394EFF?logo=virustotal&logoColor=white)](https://www.virustotal.com/gui/file/b70e0faa633ee4edf7234d842829588b67b5f9422a22a5d55f33c17c9b0ae57e/detection)

A single-file Windows tool that installs your usual apps after a fresh Windows install. Tick the apps you want, click **Install**, and it downloads and installs them silently, one after another, while you keep using the PC.

It's one PowerShell script with a built-in graphical interface. Nothing else needs downloading: no Python, no .NET SDK, no extra files. Under the hood it uses **winget**, the package manager built into Windows 10 and 11, so every app comes from the official winget catalogue.

**It's lightweight.** The whole tool is one script of about 100 KB. There's nothing to install and nothing left running afterwards. It only uses what already comes with Windows: PowerShell, WPF for the window, and winget. App icons are fetched once and saved on the PC, about 250 KB in total.

**It's free and open source.** No price, no ads, no sign-up, no tracking. The full source is this one plain-text script, so anyone can read exactly what it does before running it. It's released under the [MIT License](LICENSE): you're free to use, copy, change and share it.

**Scanned for malware.** See the [VirusTotal report](https://www.virustotal.com/gui/file/b70e0faa633ee4edf7234d842829588b67b5f9422a22a5d55f33c17c9b0ae57e/detection), which checks the script against more than 70 antivirus engines. The report is for this exact version of the file. Its SHA-256 fingerprint is:
`b70e0faa633ee4edf7234d842829588b67b5f9422a22a5d55f33c17c9b0ae57e`

## What it does

- **111 apps in 9 categories:** Browsers, Communication, Gaming, Media & Creative, Productivity, Development, Utilities, Security & Privacy, and Cloud & Remote. Examples include Steam, GOG Galaxy, Epic Games, Discord, TeamSpeak, Spotify, VS Code, Git, 7-Zip and Bitwarden.
- **Status for every app.** Each app shows where it is: *Queued → Downloading 42% → Installing → Completed / Failed*. Hover a failed status to see why.
- **The window never freezes.** Installs run in the background, so you can scroll, search or minimise while they work.
- **Skips what you already have.** On start-up it checks what's installed, marks those apps *Installed* and unticks them.
- **Official app icons.** Each icon is fetched once from the publisher's website and saved on the PC, so later runs load them instantly.
- **Search and quick selection.** Search by name or category, use *All* / *None* on any category, or *Select all*.
- **Default browser.** Pick a browser, for example Brave, and the tool makes it your default. See [Default browser](#default-browser) for what Windows allows.

## Step-by-step guide (no technical knowledge needed)

**1. Download it**

1. At the top of this GitHub page, click the green **Code** button.
2. Click **Download ZIP**. A file called `post-format-setup-main.zip` saves to your **Downloads** folder.

**2. Unzip it**

1. Open your **Downloads** folder.
2. Right-click the ZIP file and choose **Extract All…**, then click **Extract**.
3. A new folder opens. Inside it is a file called **PostFormatInstaller.ps1**. That's the tool.

> Can't see the `.ps1` at the end of the name? That's fine; Windows hides file extensions by default. Look for **PostFormatInstaller** with a type of *Windows PowerShell Script*.

**3. Start it**

1. **Right-click** `PostFormatInstaller.ps1`.
2. Choose **Run with PowerShell**. On Windows 11 it may be under **Show more options** first.
3. A blue or black PowerShell window may flash up briefly. Leave it open; it's what runs the tool. After a few seconds the **Post-Format Setup** window appears, filling the screen.

> **Don't** choose "Run as administrator". Just a normal right-click → **Run with PowerShell**.

**4. Pick your apps**

1. Tick the apps you want. Click anywhere on an app's tile to tick or untick it.
2. Use the **search box** at the top right to find an app quickly, or click a category on the left to jump to it.
3. Apps already on your PC say **Installed** and are skipped automatically.

**5. Install**

1. Click the green **Install** button at the bottom right.
2. Watch each app's status change: *Queued → Downloading → Installing → Completed*.
3. **Windows may ask for permission** for some apps (*"Do you want to allow this app to make changes to your device?"*). Check that the name matches an app you picked, then click **Yes**. More on this [below](#️-windows-will-ask-for-permission-during-installation).
4. When the bottom bar says **Finished**, you're done. Close the window with the **X** at the top right.

**If something goes wrong**

- **Nothing happens, or the window flashes and closes:** right-click the file → **Properties**. If there's an **Unblock** checkbox at the bottom, tick it, click **OK**, and try step 3 again.
- **"Windows protected your PC" (blue box):** click **More info → Run anyway**. Windows shows this for files downloaded from the internet that it hasn't seen many times before.
- **An app says *Failed*:** hover over the red status to see why, or click **Open log** at the bottom. Usually running the install again fixes it.

## How to run it (quick version)

1. Copy `PostFormatInstaller.ps1` to the new PC, for example on a USB stick.
2. Right-click it and choose **Run with PowerShell**.
   Or open PowerShell in the same folder and run:
   ```powershell
   powershell -NoProfile -ExecutionPolicy Bypass -File .\PostFormatInstaller.ps1
   ```
3. Tick the apps you want and click **Install**.

> **Run it as your normal user, not "as administrator".** Apps that need admin rights ask for them on their own (see below), and some installers, Spotify among them, refuse to run as administrator.

To look around without installing anything, add `-DryRun`. It shows the whole interface with simulated downloads and installs:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File .\PostFormatInstaller.ps1 -DryRun
```

## ⚠️ Windows will ask for permission during installation

Some apps install for every user on the PC, which needs administrator rights. For those apps Windows shows a **User Account Control (UAC)** prompt partway through: a dimmed screen asking *"Do you want to allow this app to make changes to your device?"*

**Always check the prompt before clicking Yes:**

- **Verify the program name** matches an app you selected in the list, for example "Steam Setup" while Steam shows *Installing…*.
- **Check the publisher** ("Verified publisher") is the company you expect, for example *Valve Corp.*, *Discord Inc.* or *Brave Software Inc.*
- If the name or publisher doesn't match what you selected, or the publisher shows as *Unknown*, click **No**. That app is marked *Failed* and the tool moves on to the next one.

The prompt may come up behind other windows, or flash in the taskbar. If an app sits on *Installing…* for a long time, look for a waiting permission prompt.

If you ticked **Also enforce with policy** in the Default browser box, you'll see one more prompt, from **Windows PowerShell**. That one is the tool itself writing the default-browser policy.

## Default browser

Windows 10 and 11 deliberately make it hard for programs to change your default browser. The tool tries, in this order:

1. **Direct registry change.** This works on older Windows builds, but recent updates block it.
2. **Settings page.** If Windows blocks step 1, Settings opens on your chosen browser's page. Click **Set default** once. The tool notices straight away and confirms.
3. **Policy (optional, needs admin).** Tick *Also enforce with policy* to apply Microsoft's own policy for default apps. It takes effect at your next sign-in and is reapplied at every sign-in.

With *Apply automatically when installation finishes* ticked, this runs as soon as the chosen browser has finished installing.

## Adding or removing apps

Every app is one line in a table near the top of the script:

```
# Category          | Name            | winget ID              | Icon site        | On
Gaming              | Steam           | Valve.Steam            | steampowered.com |
Communication       | Discord         | Discord.Discord        | discord.com      | on
```

- **winget ID:** find it with `winget search <name>` in a terminal.
- **Icon site:** the app's website, used to fetch its official icon. Leave it empty to use the app's own icon once installed.
- **On:** write `on` to have the app ticked by default.
- **Category:** a new name creates a new section and a new sidebar entry. Apps in **Browsers** also appear in the Default browser list.

Delete a line to remove an app.

## Requirements

- Windows 10 (1809 or newer) or Windows 11.
- **winget** (App Installer). It comes with Windows 11 and up-to-date Windows 10. If the tool can't find winget, it shows a message. Update **App Installer** from the Microsoft Store, then run the tool again.
- An internet connection.

## Files it creates

| Location | What |
|---|---|
| `%LOCALAPPDATA%\PostFormatInstaller\icons\` | Saved copies of the app icons |
| `%LOCALAPPDATA%\PostFormatInstaller\install-*.log` | A log of each run, also opened by the **Open log** button |
| `%ProgramData%\PostFormatInstaller\DefaultAssociations.xml` | Only if you used *Also enforce with policy* |

## Troubleshooting

- **An app says *Failed*.** Hover its status to see the error, or click **Open log** for winget's full output. Usual causes: you clicked *No* on the permission prompt, there was a network problem, or the app was already running.
- **An app says *Restart needed*.** It installed, but Windows needs a restart to finish setting it up.
- **The window doesn't open.** Your PC may block scripts. Run it with the `-ExecutionPolicy Bypass` command shown above.

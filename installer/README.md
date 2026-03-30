# Windows Installer Packaging

`dist/NT_DL` is the actual app `onedir` build.

`build_installer.ps1` now creates the installer as a folder-based package:

- `dist/NT_DL-Setup-<version>/`
- `dist/NT_DL-Setup-<version>.zip`

The setup EXE lives inside that installer folder. This avoids the old top-level self-extracting bootstrap EXE pattern, which could behave differently across machines because it was a single-file wrapper around the real `onedir` payload.

Release guidance:

- Preferred: distribute `NT_DL-Setup-<version>.zip`
- After extraction, run `NT_DL-Setup-<version>\\NT_DL-Setup-<version>.exe`
- Installed app target remains `%LOCALAPPDATA%\\Programs\\NT_DL`

# Windows Installer Packaging

TNT DL is released in two side-by-side editions:

- **NT DL Full** — includes Imprest Surrender and Imprest Old Date.
- **NT DL Standard** — excludes both Imprest engines and their UI controls.

Both editions include Receipt Reconciliation.

## Build application folders

```powershell
python -m PyInstaller --noconfirm --workpath build_release_full --distpath dist KDL.spec
python -m PyInstaller --noconfirm --workpath build_release_standard --distpath dist KDL_Standard.spec
```

The resulting application folders are:

- `dist/NT_DL_Full/NT_DL_Full.exe`
- `dist/NT_DL_Standard/NT_DL_Standard.exe`

## Build installers

```powershell
.\build_installer.ps1 -Edition Full
.\build_installer.ps1 -Edition Standard
```

The installers use separate application IDs and install directories, so both
editions can be installed on the same computer:

- `dist/NT_DL-Full-Setup-<version>.exe`
- `dist/NT_DL-Standard-Setup-<version>.exe`

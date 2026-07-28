# Windows Installer Packaging

TNT DL is released in two side-by-side editions:

- **NT DL Full** — includes all tools: Receipt/F.O. 30 reconciliation,
  Notes/IFMIS financial reports, Budget, Imprest Reconciliation, Filter Engine,
  Imprest Surrender, and Imprest Old Date.
- **NT DL Standard** — keeps Bank Statement Converter and Load History but
  excludes the Full-only reporting, reconciliation, filtering, budget, and
  Imprest tools and engines.

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

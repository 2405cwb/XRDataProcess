# Errors

Command failures and integration errors.

---

## [ERR-20260730-002] excel-temporary-lock-file

**Logged**: 2026-07-30T10:05:00+08:00
**Priority**: low
**Status**: resolved
**Area**: docs

### Summary
Workbook discovery included an Excel temporary lock file that cannot be opened as a ZIP workbook.

### Error
```
PermissionError: [Errno 13] Permission denied: '~$...xlsx'
```

### Suggested Fix
Exclude filenames beginning with `~$` when scanning user folders for Excel workbooks.

### Metadata
- Reproducible: yes
- Related Files: 原始平整度结果

### Resolution
- **Resolved**: 2026-07-30T10:05:00+08:00
- **Notes**: Analysis scan filters Excel temporary files.

---

## [ERR-20260730-001] powershell-empty-pipeline-report-stats

**Logged**: 2026-07-30T09:35:00+08:00
**Priority**: low
**Status**: resolved
**Area**: docs

### Summary
A compound PowerShell report-statistics command contained an empty pipeline element.

### Error
```
An empty pipe element is not allowed.
```

### Suggested Fix
Split multi-stage detail extraction into named intermediate variables before formatting output.

### Metadata
- Reproducible: yes
- Related Files: LP_KB_Test_Result.txt

### Resolution
- **Resolved**: 2026-07-30T09:35:00+08:00
- **Notes**: Continued with file-based aggregation in the document generator.

---

## [ERR-20260727-003] ref-parameter-lambda-capture

**Logged**: 2026-07-27T10:45:00+08:00
**Priority**: low
**Status**: resolved
**Area**: backend

### Summary
The local LP search attempted to capture a `ref` List parameter in LINQ lambdas.

### Error
```
CS1628: cannot use ref, out, or in parameter in an anonymous method or lambda expression
```

### Suggested Fix
Assign the current list reference to a non-ref local variable before using it in LINQ.

### Metadata
- Reproducible: yes
- Related Files: XRDataProcess/GlobalExcel.cs

### Resolution
- **Resolved**: 2026-07-27T10:45:00+08:00
- **Notes**: Lambda now uses a non-ref local list reference.

---

## [ERR-20260727-002] powershell-if-expression

**Logged**: 2026-07-27T10:35:00+08:00
**Priority**: low
**Status**: resolved
**Area**: infra

### Summary
A diagnostic projection attempted to use a PowerShell `if` statement as an expression.

### Error
```
if : The term 'if' is not recognized as the name of a cmdlet
```

### Suggested Fix
Assign the conditional result to a local variable before constructing the projected object.

### Metadata
- Reproducible: yes
- Related Files: IRIMTD/DAQ0/LP_KB_Test_Result.txt

### Resolution
- **Resolved**: 2026-07-27T10:35:00+08:00
- **Notes**: Continued with a statement-based projection.

---

## [ERR-20260727-001] display-path-star-separator

**Logged**: 2026-07-27T10:30:00+08:00
**Priority**: low
**Status**: resolved
**Area**: infra

### Summary
Project paths copied with asterisks as display separators did not match the actual Windows directory name.

### Error
```
Test-Path -LiteralPath returned false for the supplied path.
```

### Context
- The actual project uses underscores in its directory name.
- Located the directory safely by project identifier and timestamp before continuing.

### Suggested Fix
When project paths contain asterisks in user text, first resolve the physical directory by a stable project identifier and timestamp.

### Metadata
- Reproducible: yes
- Related Files: C:\\Users\\cwb\\Desktop\\job\\01二维公路软件\\temp

### Resolution
- **Resolved**: 2026-07-27T10:30:00+08:00
- **Notes**: Resolved to `C0A0360121_和丰驼洲_上行_1_江西省_南昌市_南昌县_20251024_105830`.

---

## [ERR-20260724-004] word-com-docx-render-timeout

**Logged**: 2026-07-24T11:40:30+08:00
**Priority**: low
**Status**: pending
**Area**: docs

### Summary
Hidden Word COM conversion of a generated DOCX to PDF timed out while another interactive Word instance was open.

### Error
```
command timed out after 60015 milliseconds
```

### Context
- Attempted `ExportAsFixedFormat` for visual page validation.
- LibreOffice was unavailable, so the normal `render_docx.py` path could not be used.
- A separate hidden WINWORD process remained after timeout; it was identified by start time and stopped without touching the user's older interactive Word process.
- Retrying from an ASCII-only temporary DOCX path reached the PDF export call but timed out again.
- Both bundled and system Python lacked `win32com`, so `DispatchEx` was not available as an isolation alternative.

### Suggested Fix
Prefer LibreOffice rendering when available. Otherwise copy the DOCX to an ASCII-only temporary path and launch an isolated Word instance, or fall back to OOXML and text/table structural validation while clearly reporting the rendering limitation.

### Metadata
- Reproducible: unknown
- Related Files: output/doc/IRI_KB等效反映到LP方案说明.docx

---

## [ERR-20260724-003] broad-output-config-search

**Logged**: 2026-07-24T00:00:00+08:00
**Priority**: low
**Status**: resolved
**Area**: infra

### Summary
A recursive PowerShell config search traversed unintended binary files and timed out.

### Error
```
command timed out after 20054 milliseconds
```

### Context
- Searched Output recursively with `-Include` and piped results to Select-String.
- PowerShell included binary build outputs despite the intended config-only scope.

### Suggested Fix
Resolve the exact settings file first and use `Select-String -LiteralPath`.

### Metadata
- Reproducible: yes
- Related Files: Output/XRSetting.ini

### Resolution
- **Resolved**: 2026-07-24T00:00:00+08:00
- **Notes**: Confirmed the active setting directly as sheetRoundingOffNum=5.

---

## [ERR-20260724-002] dotnet-msbuild-com-reference

**Logged**: 2026-07-24T00:00:00+08:00
**Priority**: low
**Status**: resolved
**Area**: infra

### Summary
The .NET SDK MSBuild cannot build the .NET Framework project because it contains Office COM references.

### Error
```
MSB4803: .NET Core MSBuild does not support ResolveComReference.
```

### Context
- Attempted to validate XRDataProcess.csproj with `dotnet msbuild`.
- The project targets .NET Framework 4.8 and contains Office Interop COM references.

### Suggested Fix
Use Visual Studio's full .NET Framework MSBuild executable for this project.

### Metadata
- Reproducible: yes
- Related Files: XRDataProcess/XRDataProcess.csproj

### Resolution
- **Resolved**: 2026-07-24T00:00:00+08:00
- **Notes**: Debug x64 build succeeded with Visual Studio 2022 MSBuild.

---

## [ERR-20260724-001] powershell-rg-probe

**Logged**: 2026-07-24T00:00:00+08:00
**Priority**: low
**Status**: resolved
**Area**: infra

### Summary
Repository inspection commands returned non-zero because of an invalid rg escape and because rg returns 1 when no matches are found.

### Error
```
rg: regex parse error: unrecognized escape sequence
```

### Context
- A PowerShell inspection command embedded a quoted MSBuild pattern in an rg regular expression.
- A later optional no-match search also propagated rg exit code 1.

### Suggested Fix
Use `Select-String -LiteralPath` for fixed MSBuild searches and append an explicit no-match handler when rg results are optional.

### Metadata
- Reproducible: yes
- Related Files: XRDataProcess/XRDataProcess.csproj

### Resolution
- **Resolved**: 2026-07-24T00:00:00+08:00
- **Notes**: Switched fixed metadata lookup to Select-String and will isolate optional rg exit codes.

---

# MS Project VB* Utils

Utilities for inspecting Microsoft Project MPP files with [VBScript](https://en.wikipedia.org/wiki/VBScript), [VBA](https://en.wikipedia.org/wiki/Visual_Basic_for_Applications), and Java.

## Report summary

Both reports output Microsoft Project file details without saving changes to the MPP file. The report includes:

- File size
- Project properties
- Built-in document properties
- Custom document properties
- Custom field names
- Project custom field values from the project summary task in the Java and VBScript reports

Custom fields are printed in this format:

```text
Text1: Custom Task Text
Number1: Custom Task Number
Flag1: Custom Task Flag
```

## Prerequisites

Requires:

- Windows
- Microsoft Project desktop installed
- Access to the target `.mpp` file
- Windows Script Host for running the VBScript report with `cscript`
- Microsoft Project VBA access for importing and running the VBA report
- Java 21 and Maven for running the Java console reports

Developed and tested on:

- Windows 11
- Microsoft Project Online Desktop Client MSO, Version 2606 Build 16.0.20131.20154, 64-bit

## Run the VBScript report

Use `vbscript/project-report.vbs` when you want to open an MPP file by path, print the report to the console, and close the file without saving changes.

From PowerShell, set the MPP path and run the script with `cscript`:

```powershell
$mppFile = "C:\path\to\project.mpp"
cscript //nologo .\vbscript\project-report.vbs $mppFile
```

Or pass the path directly:

```powershell
cscript //nologo .\vbscript\project-report.vbs "C:\path\to\project.mpp"
```

Optionally, to make `.vbs` files run with `cscript` by default, open an administrator terminal and run:

```powershell
cscript //H:CScript
```

After that, you can run the script without explicitly calling `cscript`:

```powershell
.\vbscript\project-report.vbs "C:\path\to\project.mpp"
```

The script starts or attaches to Microsoft Project, opens the MPP file, prints the report, then closes the MPP without saving changes.

## Run the MPXJ Java console report

Use `mpxj/` when you want a console report powered by [MPXJ](https://www.mpxj.org/) 16.7.0 instead of Microsoft Project automation.

Build and test the MPXJ module:

```powershell
cd mpxj
mvn verify
```

Run directly with Maven:

```powershell
mvn exec:java '-DmppPath=C:\path\to\project.mpp'
```

Run the shaded console jar:

```powershell
java -jar .\target\project-report-1.0.0-SNAPSHOT.jar "C:\path\to\project.mpp"
```

The MPXJ Java report prints file size, project properties, built-in document properties, custom document properties, custom field aliases with populated value counts where MPXJ exposes them, and project custom field values read from the project summary task. Unit tests run with JUnit 5, and `mvn verify` generates a JaCoCo coverage report under `mpxj/target/site/jacoco/`.

## Run the Aspose.Tasks Java console report

Use `aspose/` when you want the same console report shape powered by [Aspose.Tasks for Java](https://products.aspose.com/tasks/java/). The module uses the `jdk18` Aspose.Tasks artifact classifier from the Aspose Java repository, which works with Java 21.

Build and test the Aspose.Tasks module:

```powershell
cd aspose
mvn verify
```

Run directly with Maven:

```powershell
mvn exec:java '-DmppPath=C:\path\to\project.mpp'
```

Run the shaded console jar:

```powershell
java -jar .\target\project-report-aspose-1.0.0-SNAPSHOT.jar "C:\path\to\project.mpp"
```

The Aspose.Tasks Java report prints file size, project properties, built-in document properties, custom document properties, task/resource custom field aliases with populated value counts, project custom field values from the root project summary task, and assignments. Without an Aspose license, output may be subject to Aspose evaluation-mode behavior.

## Run the Immediate Window VBA report

Use `vba/project-report.bas` when the MPP file is already open in Microsoft Project and you want the report printed to the VBA Immediate Window.

1. Open the target `.mpp` file in Microsoft Project.
2. Press `Alt+F11` to open the VBA editor.
3. In the VBA editor, select the project for the open MPP file in the Project Explorer.
4. Select `File` > `Import File...`.
5. Choose `vba/project-report.bas` from this repo.
6. Press `Ctrl+G` to show the Immediate Window.
7. Run the macro `PrintActiveProjectReportToImmediate`.

You can run the macro from `Run` > `Run Sub/UserForm`, or from the Immediate Window with:

```vb
PrintActiveProjectReportToImmediate
```

The macro prints the report for the active project.

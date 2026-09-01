package com.github.jrg.msprojectutils.projectreport;

import java.io.IOException;
import java.io.PrintStream;
import java.nio.file.Files;
import java.nio.file.Path;

import org.mpxj.ProjectFile;
import org.mpxj.MPXJException;
import org.mpxj.reader.UniversalProjectReader;

public final class ProjectReportApp {
    private ProjectReportApp() {
    }

    public static void main(String[] args) {
        int exitCode = run(args, System.out, System.err);
        if (exitCode != 0) {
            System.exit(exitCode);
        }
    }

    static int run(String[] args, PrintStream out, PrintStream err) {
        if (args.length != 1) {
            err.println("Usage: java -jar project-report.jar <path-to-project-file>");
            return 1;
        }

        Path projectPath = Path.of(args[0]);
        if (!Files.isRegularFile(projectPath)) {
            err.println("Project file was not found: " + projectPath);
            return 1;
        }

        try {
            out.println("info: Running Java project-report");
            out.println("info: Opening project file: " + projectPath);
            out.println("info: Project file size: " + Files.size(projectPath) + " bytes");

            ProjectFile projectFile = new UniversalProjectReader().read(projectPath.toString());
            out.print(new ProjectReportFormatter().format(projectFile, projectPath));
            return 0;
        } catch (IOException | MPXJException e) {
            err.println("Could not read project file. " + e.getMessage());
            return 1;
        } catch (RuntimeException e) {
            err.println("Could not generate project report. " + e.getMessage());
            return 1;
        }
    }
}
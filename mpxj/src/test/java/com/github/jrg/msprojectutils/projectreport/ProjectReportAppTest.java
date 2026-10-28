package com.github.jrg.msprojectutils.projectreport;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.io.ByteArrayOutputStream;
import java.io.OutputStream;
import java.io.PrintStream;
import java.nio.charset.StandardCharsets;

import org.junit.jupiter.api.Test;

class ProjectReportAppTest {
    @Test
    void returnsUsageErrorWhenPathIsMissing() {
        ByteArrayOutputStream err = new ByteArrayOutputStream();

        int exitCode = ProjectReportApp.run(new String[0], nullPrintStream(), new PrintStream(err, true, StandardCharsets.UTF_8));

        assertEquals(1, exitCode);
        assertTrue(err.toString(StandardCharsets.UTF_8).contains("Usage: java -jar project-report.jar <path-to-project-file>"));
    }

    @Test
    void returnsNotFoundErrorWhenFileDoesNotExist() {
        ByteArrayOutputStream err = new ByteArrayOutputStream();

        int exitCode = ProjectReportApp.run(new String[] { "missing.mpp" }, nullPrintStream(), new PrintStream(err, true, StandardCharsets.UTF_8));

        assertEquals(1, exitCode);
        assertTrue(err.toString(StandardCharsets.UTF_8).contains("Project file was not found: missing.mpp"));
    }

    private PrintStream nullPrintStream() {
        return new PrintStream(OutputStream.nullOutputStream(), true, StandardCharsets.UTF_8);
    }
}
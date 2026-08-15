package com.github.jrg.msprojectutils.projectreport;

import static org.junit.jupiter.api.Assertions.assertTrue;

import java.nio.file.Path;

import org.junit.jupiter.api.Test;
import org.mpxj.ProjectFile;
import org.mpxj.Resource;
import org.mpxj.ResourceField;
import org.mpxj.Task;
import org.mpxj.TaskField;

class ProjectReportFormatterTest {
    @Test
    void formatsProjectPropertiesAndEmptyCustomFieldSections() {
        ProjectFile projectFile = new ProjectFile();
        projectFile.getProjectProperties().setProjectTitle("Network Voltage Control");
        projectFile.getProjectProperties().setAuthor("Engineering");
        projectFile.getProjectProperties().setCompany("Contoso");
        projectFile.getProjectProperties().setComments("Baseline schedule");

        String report = new ProjectReportFormatter().format(projectFile, Path.of("sample.mpp"));

        assertTrue(report.contains("Project properties"));
        assertTrue(report.contains("Title: Network Voltage Control"));
        assertTrue(report.contains("Author: Engineering"));
        assertTrue(report.contains("Company: Contoso"));
        assertTrue(report.contains("Comments: Baseline schedule"));
        assertTrue(report.contains("Task custom fields" + System.lineSeparator() + "------------------" + System.lineSeparator() + "<none>"));
        assertTrue(report.endsWith("------------------" + System.lineSeparator()));
    }

    @Test
    void formatsTaskAndResourceCustomFieldAliasesWithValueCounts() {
        ProjectFile projectFile = new ProjectFile();
        projectFile.getCustomFields().getOrCreate(TaskField.TEXT1).setAlias("PO Description");
        projectFile.getCustomFields().getOrCreate(ResourceField.FLAG2).setAlias("Work Order Executed");

        Task populatedTask = projectFile.addTask();
        populatedTask.set(TaskField.TEXT1, "Transformer install");
        projectFile.addTask();

        Resource populatedResource = projectFile.addResource();
        populatedResource.set(ResourceField.FLAG2, true);
        projectFile.addResource();

        String report = new ProjectReportFormatter().format(projectFile, Path.of("sample.mpp"));

        assertTrue(report.contains("Text1: PO Description (1 value)"));
        assertTrue(report.contains("Flag2: Work Order Executed (1 value)"));
        assertTrue(report.contains("Project custom fields" + System.lineSeparator() + "---------------------" + System.lineSeparator() + "<none>"));
    }
}
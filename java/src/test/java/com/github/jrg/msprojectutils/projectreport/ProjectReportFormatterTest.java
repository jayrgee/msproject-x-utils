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
        projectFile.getCustomFields().getOrCreate(TaskField.TEXT1).setAlias("Custom Task Text");
        projectFile.getCustomFields().getOrCreate(ResourceField.FLAG2).setAlias("Custom Resource Flag");

        Task populatedTask = projectFile.addTask();
        populatedTask.set(TaskField.TEXT1, "Task value");
        projectFile.addTask();

        Resource populatedResource = projectFile.addResource();
        populatedResource.set(ResourceField.FLAG2, true);
        projectFile.addResource();

        String report = new ProjectReportFormatter().format(projectFile, Path.of("sample.mpp"));

        assertTrue(report.contains("Text1: Custom Task Text (1 value)"));
        assertTrue(report.contains("Flag2: Custom Resource Flag (1 value)"));
        assertTrue(report.contains("Project custom fields" + System.lineSeparator() + "---------------------" + System.lineSeparator() + "<none>"));
    }

    @Test
    void usesTaskIdZeroAsProjectSummaryAndExcludesItFromTaskCounts() {
        ProjectFile projectFile = new ProjectFile();
        projectFile.getCustomFields().getOrCreate(TaskField.TEXT1).setAlias("Custom Project Text");

        Task projectSummaryTask = projectFile.addTask();
        projectSummaryTask.setID(0);
        projectSummaryTask.set(TaskField.TEXT1, "Project value");

        Task task = projectFile.addTask();
        task.setID(1);

        String report = new ProjectReportFormatter().format(projectFile, Path.of("sample.mpp"));

        assertTrue(report.contains("Task custom fields" + System.lineSeparator()
                + "------------------" + System.lineSeparator()
                + "Text1: Custom Project Text (0 values)"));
        assertTrue(report.contains("Project custom fields" + System.lineSeparator()
                + "---------------------" + System.lineSeparator()
                + "Text1: Custom Project Text (1 value)"));
    }

    @Test
    void excludesResourceIdZeroFromResourceCounts() {
        ProjectFile projectFile = new ProjectFile();
        projectFile.getCustomFields().getOrCreate(ResourceField.NUMBER1).setAlias("Custom Resource Number");

        Resource placeholderResource = projectFile.addResource();
        placeholderResource.setID(0);
        placeholderResource.set(ResourceField.NUMBER1, 0.0);

        Resource resource = projectFile.addResource();
        resource.setID(1);

        String report = new ProjectReportFormatter().format(projectFile, Path.of("sample.mpp"));

        assertTrue(report.contains("Resource custom fields" + System.lineSeparator()
                + "----------------------" + System.lineSeparator()
            + "Number1: Custom Resource Number (0 values)"));
    }
}
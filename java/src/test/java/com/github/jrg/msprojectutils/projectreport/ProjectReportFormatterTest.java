package com.github.jrg.msprojectutils.projectreport;

import static org.junit.jupiter.api.Assertions.assertTrue;

import java.nio.file.Path;

import org.junit.jupiter.api.Test;
import org.mpxj.ProjectFile;
import org.mpxj.Resource;
import org.mpxj.ResourceField;
import org.mpxj.Task;
import org.mpxj.TaskField;
import org.mpxj.mpp.CustomFieldValueItem;

class ProjectReportFormatterTest {
    @Test
    void formatsProjectPropertiesAndEmptyCustomFieldSections() {
        ProjectFile projectFile = new ProjectFile();
        projectFile.getProjectProperties().setProjectTitle("Sample Project");
        projectFile.getProjectProperties().setAuthor("Sample Author");
        projectFile.getProjectProperties().setCompany("Sample Company");
        projectFile.getProjectProperties().setComments("Sample comments");

        String report = new ProjectReportFormatter().format(projectFile, Path.of("sample.mpp"));

        assertTrue(report.contains("Project properties"));
        assertTrue(report.contains("Title: Sample Project"));
        assertTrue(report.contains("Author: Sample Author"));
        assertTrue(report.contains("Company: Sample Company"));
        assertTrue(report.contains("Comments: Sample comments"));
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
        projectFile.getCustomFields().getOrCreate(TaskField.TEXT30).setAlias("Custom Project Lookup");
        CustomFieldValueItem lookupItem = new CustomFieldValueItem(325);
        lookupItem.setValue("Lookup Option A");
        projectFile.getCustomFields().get(TaskField.TEXT30).getLookupTable().add(lookupItem);

        Task projectSummaryTask = projectFile.addTask();
        projectSummaryTask.setID(0);
        projectSummaryTask.set(TaskField.TEXT1, "Project value");
        projectSummaryTask.set(TaskField.TEXT30, "\u0001" + (char) 325);

        Task task = projectFile.addTask();
        task.setID(1);

        String report = new ProjectReportFormatter().format(projectFile, Path.of("sample.mpp"));

        assertTrue(report.contains("Task custom fields" + System.lineSeparator()
                + "------------------" + System.lineSeparator()
                + "Text1: Custom Project Text (0 values)"));
        assertTrue(report.contains("Project custom fields" + System.lineSeparator()
                + "---------------------" + System.lineSeparator()
            + "Text1: Custom Project Text = Project value (1 value)"));
        assertTrue(report.contains("Text30: Custom Project Lookup = Lookup Option A (1 value)"));
    }

    @Test
    void formatsProjectLookupValueWhenEncodedUniqueIdIsNotPrefixedByControlCharacter() {
        ProjectFile projectFile = new ProjectFile();
        projectFile.getCustomFields().getOrCreate(TaskField.TEXT30).setAlias("Custom Project Lookup");
        CustomFieldValueItem lookupItem = new CustomFieldValueItem(322);
        lookupItem.setValue("Lookup Option B");
        projectFile.getCustomFields().get(TaskField.TEXT30).getLookupTable().add(lookupItem);

        Task projectSummaryTask = projectFile.addTask();
        projectSummaryTask.setID(0);
        projectSummaryTask.set(TaskField.TEXT30, "\u0301\u0142");

        String report = new ProjectReportFormatter().format(projectFile, Path.of("sample.mpp"));

        assertTrue(report.contains("Text30: Custom Project Lookup = Lookup Option B (1 value)"));
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
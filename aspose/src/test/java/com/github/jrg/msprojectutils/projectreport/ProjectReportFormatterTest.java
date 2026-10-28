package com.github.jrg.msprojectutils.projectreport;

import static org.junit.jupiter.api.Assertions.assertTrue;

import java.math.BigDecimal;
import java.nio.file.Path;

import org.junit.jupiter.api.Test;

import com.aspose.tasks.ExtendedAttributeDefinition;
import com.aspose.tasks.Field;
import com.aspose.tasks.Prj;
import com.aspose.tasks.Project;
import com.aspose.tasks.Resource;
import com.aspose.tasks.ResourceType;
import com.aspose.tasks.Task;

class ProjectReportFormatterTest {
    @Test
    void formatsProjectPropertiesAndEmptyCustomFieldSections() {
        Project project = new Project();
        project.set(Prj.TITLE, "Sample Project");
        project.set(Prj.AUTHOR, "Sample Author");
        project.set(Prj.COMPANY, "Sample Company");
        project.set(Prj.COMMENTS, "Sample comments");

        String report = new ProjectReportFormatter().format(project, Path.of("sample.mpp"));

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
        Project project = new Project();
        ExtendedAttributeDefinition taskText1 = ExtendedAttributeDefinition.createTaskDefinition(Field.TaskText1, "Custom Task Text");
        ExtendedAttributeDefinition resourceFlag2 = ExtendedAttributeDefinition.createResourceDefinition(Field.ResourceFlag2, "Custom Resource Flag");
        project.getExtendedAttributes().add(taskText1);
        project.getExtendedAttributes().add(resourceFlag2);

        Task populatedTask = project.getRootTask().getChildren().add("Populated task");
        populatedTask.getExtendedAttributes().add(taskText1.createExtendedAttribute("Task value"));
        project.getRootTask().getChildren().add("Empty task");

        Resource populatedResource = project.getResources().add("Populated resource");
        populatedResource.getExtendedAttributes().add(resourceFlag2.createExtendedAttribute(true));
        project.getResources().add("Empty resource");

        String report = new ProjectReportFormatter().format(project, Path.of("sample.mpp"));

        assertTrue(report.contains("Text1: Custom Task Text (1 value)"));
        assertTrue(report.contains("Flag2: Custom Resource Flag (1 value)"));
    }

    @Test
    void formatsAssignmentListWithTaskResourceTypeAndTaskText10() {
        Project project = new Project();
        ExtendedAttributeDefinition taskText10 = ExtendedAttributeDefinition.createTaskDefinition(Field.TaskText10, "Task Review");
        project.getExtendedAttributes().add(taskText10);

        Task task = project.getRootTask().getChildren().add("Build interface");
        task.getExtendedAttributes().add(taskText10.createExtendedAttribute("Design review"));

        Resource resource = project.getResources().add("Alex Rivera");
        resource.setType(ResourceType.Work);

        project.getResourceAssignments().add(task, resource);

        String report = new ProjectReportFormatter().format(project, Path.of("sample.mpp"));

        assertTrue(report.contains("Assignments" + System.lineSeparator()
                + "-----------" + System.lineSeparator()
                + "taskname\tresourcename\tresourcetype\tText10" + System.lineSeparator()
                + "Build interface\tAlex Rivera\tWORK\tDesign review"));
    }

    @Test
    void usesRootTaskAsProjectSummaryAndExcludesItFromTaskCounts() {
        Project project = new Project();
        ExtendedAttributeDefinition taskText1 = ExtendedAttributeDefinition.createTaskDefinition(Field.TaskText1, "Custom Project Text");
        project.getExtendedAttributes().add(taskText1);

        project.getRootTask().getExtendedAttributes().add(taskText1.createExtendedAttribute("Project value"));
        project.getRootTask().getChildren().add("Child task");

        String report = new ProjectReportFormatter().format(project, Path.of("sample.mpp"));

        assertTrue(report.contains("Task custom fields" + System.lineSeparator()
                + "------------------" + System.lineSeparator()
                + "Text1: Custom Project Text (0 values)"));
        assertTrue(report.contains("Project custom fields" + System.lineSeparator()
                + "---------------------" + System.lineSeparator()
                + "Text1: Custom Project Text = Project value (1 value)"));
    }

    @Test
    void excludesResourceIdZeroFromResourceCounts() {
        Project project = new Project();
        ExtendedAttributeDefinition resourceNumber1 = ExtendedAttributeDefinition.createResourceDefinition(Field.ResourceNumber1, "Custom Resource Number");
        project.getExtendedAttributes().add(resourceNumber1);

        Resource placeholderResource = project.getResources().add();
    placeholderResource.getExtendedAttributes().add(resourceNumber1.createExtendedAttribute(BigDecimal.ZERO));

        project.getResources().add("Actual resource");

        String report = new ProjectReportFormatter().format(project, Path.of("sample.mpp"));

        assertTrue(report.contains("Resource custom fields" + System.lineSeparator()
                + "----------------------" + System.lineSeparator()
                + "Number1: Custom Resource Number (0 values)"));
    }
}
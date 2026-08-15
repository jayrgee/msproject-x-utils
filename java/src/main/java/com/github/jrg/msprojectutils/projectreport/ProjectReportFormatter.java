package com.github.jrg.msprojectutils.projectreport;

import java.nio.file.Path;
import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.Objects;
import java.util.function.Function;

import org.mpxj.CustomField;
import org.mpxj.FieldContainer;
import org.mpxj.FieldType;
import org.mpxj.ProjectFile;
import org.mpxj.ProjectProperties;
import org.mpxj.Resource;
import org.mpxj.ResourceField;
import org.mpxj.Task;
import org.mpxj.TaskField;

final class ProjectReportFormatter {
    private static final List<CustomFieldRange> CUSTOM_FIELD_RANGES = List.of(
            new CustomFieldRange("Text", "TEXT", 30),
            new CustomFieldRange("Number", "NUMBER", 20),
            new CustomFieldRange("Date", "DATE", 10),
            new CustomFieldRange("Duration", "DURATION", 10),
            new CustomFieldRange("Cost", "COST", 10),
            new CustomFieldRange("Flag", "FLAG", 20),
            new CustomFieldRange("Outline Code", "OUTLINE_CODE", 10));

    String format(ProjectFile projectFile, Path projectPath) {
        Objects.requireNonNull(projectFile, "projectFile");
        Objects.requireNonNull(projectPath, "projectPath");

        StringBuilder report = new StringBuilder();
        report.append("info: Project file opened:  ").append(projectPath).append(System.lineSeparator());
        appendProjectProperties(report, projectFile, projectPath);
        appendCustomFields(report, projectFile);
        report.append("------------------").append(System.lineSeparator());
        return report.toString();
    }

    private void appendProjectProperties(StringBuilder report, ProjectFile projectFile, Path projectPath) {
        ProjectProperties properties = projectFile.getProjectProperties();

        report.append(System.lineSeparator());
        report.append("Project properties").append(System.lineSeparator());
        report.append("------------------").append(System.lineSeparator());

        Map<String, Object> values = new LinkedHashMap<>();
        values.put("Name", projectPath.getFileName());
        values.put("FullName", projectPath.toAbsolutePath());
        values.put("Path", projectPath.toAbsolutePath().getParent());
        values.put("Title", properties.getProjectTitle());
        values.put("Subject", properties.getSubject());
        values.put("Author", properties.getAuthor());
        values.put("Manager", properties.getManager());
        values.put("Company", properties.getCompany());
        values.put("Comments", properties.getComments());
        values.put("CreationDate", properties.getCreationDate());
        values.put("LastSaveDate", properties.getLastSaved());
        values.put("Start", properties.getStartDate());
        values.put("Finish", properties.getFinishDate());
        values.put("StatusDate", properties.getStatusDate());
        values.put("CurrentDate", properties.getCurrentDate());
        values.put("Calendar", projectFile.getDefaultCalendar() == null ? null : projectFile.getDefaultCalendar().getName());
        values.put("CurrencySymbol", properties.getCurrencySymbol());
        values.put("CurrencyCode", properties.getCurrencyCode());

        values.forEach((name, value) -> appendNameValue(report, name, value));
        appendDocumentProperties(report, "Built-in document properties", builtInDocumentProperties(properties));
        appendDocumentProperties(report, "Custom document properties", properties.getCustomProperties());
    }

    private Map<String, Object> builtInDocumentProperties(ProjectProperties properties) {
        Map<String, Object> values = new LinkedHashMap<>();
        values.put("Title", properties.getProjectTitle());
        values.put("Subject", properties.getSubject());
        values.put("Author", properties.getAuthor());
        values.put("Manager", properties.getManager());
        values.put("Company", properties.getCompany());
        values.put("Category", properties.getCategory());
        values.put("Keywords", properties.getKeywords());
        values.put("Comments", properties.getComments());
        values.put("Creation Date", properties.getCreationDate());
        values.put("Last Saved", properties.getLastSaved());
        return values;
    }

    private void appendDocumentProperties(StringBuilder report, String heading, Map<String, ?> properties) {
        report.append(System.lineSeparator());
        report.append(heading).append(System.lineSeparator());
        report.append("-".repeat(heading.length())).append(System.lineSeparator());

        if (properties == null || properties.isEmpty()) {
            report.append("<none>").append(System.lineSeparator());
            return;
        }

        properties.forEach((name, value) -> appendNameValue(report, name, value));
    }

    private void appendCustomFields(StringBuilder report, ProjectFile projectFile) {
        report.append(System.lineSeparator());
        report.append("Custom fields").append(System.lineSeparator());
        report.append("-------------").append(System.lineSeparator());

        appendCustomFieldsForScope(report, "Task", projectFile, name -> TaskField.valueOf(name), tasksWithoutProjectSummary(projectFile));
        appendCustomFieldsForScope(report, "Resource", projectFile, name -> ResourceField.valueOf(name), resourcesWithoutPlaceholder(projectFile));
        appendProjectCustomFields(report, projectFile);
    }

    private List<Task> tasksWithoutProjectSummary(ProjectFile projectFile) {
        List<Task> tasks = new ArrayList<>();
        for (Task task : projectFile.getTasks()) {
            if (!isProjectSummaryTask(task)) {
                tasks.add(task);
            }
        }
        return tasks;
    }

    private List<Resource> resourcesWithoutPlaceholder(ProjectFile projectFile) {
        List<Resource> resources = new ArrayList<>();
        for (Resource resource : projectFile.getResources()) {
            if (!isPlaceholderResource(resource)) {
                resources.add(resource);
            }
        }
        return resources;
    }

    private void appendCustomFieldsForScope(
            StringBuilder report,
            String scopeName,
            ProjectFile projectFile,
            Function<String, FieldType> fieldTypeFactory,
            Iterable<? extends FieldContainer> items) {

        boolean foundAny = false;
        for (CustomFieldRange range : CUSTOM_FIELD_RANGES) {
            for (int index = 1; index <= range.count(); index++) {
                String enumName = range.enumPrefix() + index;
                FieldType fieldType = fieldTypeFactory.apply(enumName);
                CustomField customField = projectFile.getCustomFields().get(fieldType);
                String alias = customField == null ? "" : formatValue(customField.getAlias());
                if (alias.isEmpty()) {
                    continue;
                }

                if (!foundAny) {
                    appendScopeHeading(report, scopeName);
                    foundAny = true;
                }

                report.append(range.displayPrefix()).append(index).append(": ")
                        .append(alias)
                        .append(" (").append(valueCountLabel(countValues(items, fieldType))).append(")")
                        .append(System.lineSeparator());
            }
        }

        if (!foundAny) {
            appendNoCustomFieldsForScope(report, scopeName);
        }
    }

    private void appendProjectCustomFields(StringBuilder report, ProjectFile projectFile) {
        Task projectSummaryTask = projectSummaryTask(projectFile);
        if (projectSummaryTask == null) {
            appendNoCustomFieldsForScope(report, "Project");
            return;
        }

        boolean foundAny = false;
        for (CustomFieldRange range : CUSTOM_FIELD_RANGES) {
            for (int index = 1; index <= range.count(); index++) {
                String enumName = range.enumPrefix() + index;
                FieldType fieldType = TaskField.valueOf(enumName);
                CustomField customField = projectFile.getCustomFields().get(fieldType);
                String alias = customField == null ? "" : formatValue(customField.getAlias());
                if (alias.isEmpty()) {
                    continue;
                }

                if (!foundAny) {
                    appendScopeHeading(report, "Project");
                    foundAny = true;
                }

                report.append(range.displayPrefix()).append(index).append(": ")
                        .append(alias)
                        .append(" (").append(valueCountLabel(countValues(List.of(projectSummaryTask), fieldType))).append(")")
                        .append(System.lineSeparator());
            }
        }

        if (!foundAny) {
            appendNoCustomFieldsForScope(report, "Project");
        }
    }

    private Task projectSummaryTask(ProjectFile projectFile) {
        for (Task task : projectFile.getTasks()) {
            if (isProjectSummaryTask(task)) {
                return task;
            }
        }
        return null;
    }

    private boolean isProjectSummaryTask(Task task) {
        return Integer.valueOf(0).equals(task.getID());
    }

    private boolean isPlaceholderResource(Resource resource) {
        return Integer.valueOf(0).equals(resource.getID());
    }

    private int countValues(Iterable<? extends FieldContainer> items, FieldType fieldType) {
        int count = 0;
        for (FieldContainer item : items) {
            if (!formatValue(item.get(fieldType)).isEmpty()) {
                count++;
            }
        }
        return count;
    }

    private String valueCountLabel(int valueCount) {
        return valueCount == 1 ? "1 value" : valueCount + " values";
    }

    private void appendNoCustomFieldsForScope(StringBuilder report, String scopeName) {
        appendScopeHeading(report, scopeName);
        report.append("<none>").append(System.lineSeparator());
    }

    private void appendScopeHeading(StringBuilder report, String scopeName) {
        String heading = scopeName + " custom fields";
        report.append(System.lineSeparator());
        report.append(heading).append(System.lineSeparator());
        report.append("-".repeat(heading.length())).append(System.lineSeparator());
    }

    private void appendNameValue(StringBuilder report, String name, Object value) {
        report.append(name).append(": ").append(formatValue(value)).append(System.lineSeparator());
    }

    private String formatValue(Object value) {
        if (value == null) {
            return "";
        }

        return value.toString().trim();
    }

    private record CustomFieldRange(String displayPrefix, String enumPrefix, int count) {
    }
}
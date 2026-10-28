package com.github.jrg.msprojectutils.projectreport;

import java.lang.reflect.Field;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.Objects;
import java.util.function.Function;

import com.aspose.tasks.BuiltInProjectPropertyCollection;
import com.aspose.tasks.CustomFieldType;
import com.aspose.tasks.ExtendedAttribute;
import com.aspose.tasks.ExtendedAttributeCollection;
import com.aspose.tasks.ExtendedAttributeDefinition;
import com.aspose.tasks.Prj;
import com.aspose.tasks.Project;
import com.aspose.tasks.Property;
import com.aspose.tasks.Resource;
import com.aspose.tasks.ResourceAssignment;
import com.aspose.tasks.ResourceType;
import com.aspose.tasks.Task;
import com.aspose.tasks.Value;

final class ProjectReportFormatter {
    private static final List<CustomFieldRange> CUSTOM_FIELD_RANGES = List.of(
            new CustomFieldRange("Text", "Text", 30),
            new CustomFieldRange("Number", "Number", 20),
            new CustomFieldRange("Date", "Date", 10),
            new CustomFieldRange("Duration", "Duration", 10),
            new CustomFieldRange("Cost", "Cost", 10),
            new CustomFieldRange("Flag", "Flag", 20),
            new CustomFieldRange("Outline Code", "OutlineCode", 10));

    String format(Project project, Path projectPath) {
        Objects.requireNonNull(project, "project");
        Objects.requireNonNull(projectPath, "projectPath");

        StringBuilder report = new StringBuilder();
        report.append("info: Project file opened:  ").append(projectPath).append(System.lineSeparator());
        appendProjectProperties(report, project, projectPath);
        appendCustomFields(report, project);
        appendAssignments(report, project);
        report.append("------------------").append(System.lineSeparator());
        return report.toString();
    }

    private void appendAssignments(StringBuilder report, Project project) {
        report.append(System.lineSeparator());
        report.append("Assignments").append(System.lineSeparator());
        report.append("-----------").append(System.lineSeparator());
        report.append("taskname\tresourcename\tresourcetype\tText10").append(System.lineSeparator());

        boolean foundAny = false;
        for (ResourceAssignment assignment : project.getResourceAssignments()) {
            foundAny = true;
            Resource resource = assignment.getResource();
            Task task = assignment.getTask();

            report.append(formatValue(task == null ? null : task.getName())).append('\t')
                    .append(formatValue(resource == null ? null : resource.getName())).append('\t')
                    .append(formatValue(resource == null ? null : resourceTypeName(resource.getType()))).append('\t')
                    .append(formatAssignmentCustomText(assignment, task))
                    .append(System.lineSeparator());
        }

        if (!foundAny) {
            report.append("<none>").append(System.lineSeparator());
        }
    }

    private String formatAssignmentCustomText(ResourceAssignment assignment, Task task) {
        String value = formatExtendedAttributeValue(assignment.getExtendedAttributes(), "Text10");
        return value.isEmpty() && task != null ? formatExtendedAttributeValue(task.getExtendedAttributes(), "Text10") : value;
    }

    private String resourceTypeName(int resourceType) {
        if (resourceType == ResourceType.Work) {
            return "WORK";
        }
        if (resourceType == ResourceType.Material) {
            return "MATERIAL";
        }
        if (resourceType == ResourceType.Cost) {
            return "COST";
        }
        return String.valueOf(resourceType);
    }

    private void appendProjectProperties(StringBuilder report, Project project, Path projectPath) {
        report.append(System.lineSeparator());
        report.append("Project properties").append(System.lineSeparator());
        report.append("------------------").append(System.lineSeparator());

        Map<String, Object> values = new LinkedHashMap<>();
        values.put("Name", projectPath.getFileName());
        values.put("FullName", projectPath.toAbsolutePath());
        values.put("Path", projectPath.toAbsolutePath().getParent());
        values.put("Title", project.get(Prj.TITLE));
        values.put("Subject", project.get(Prj.SUBJECT));
        values.put("Author", project.get(Prj.AUTHOR));
        values.put("Manager", project.get(Prj.MANAGER));
        values.put("Company", project.get(Prj.COMPANY));
        values.put("Comments", project.get(Prj.COMMENTS));
        values.put("CreationDate", project.get(Prj.CREATION_DATE));
        values.put("LastSaveDate", project.get(Prj.LAST_SAVED));
        values.put("Start", project.get(Prj.START_DATE));
        values.put("Finish", project.get(Prj.FINISH_DATE));
        values.put("StatusDate", project.get(Prj.STATUS_DATE));
        values.put("CurrentDate", project.get(Prj.CURRENT_DATE));
        values.put("Calendar", project.get(Prj.CALENDAR) == null ? null : project.get(Prj.CALENDAR).getName());
        values.put("CurrencySymbol", project.get(Prj.CURRENCY_SYMBOL));
        values.put("CurrencyCode", project.get(Prj.CURRENCY_CODE));

        values.forEach((name, value) -> appendNameValue(report, name, value));
        appendDocumentProperties(report, "Built-in document properties", builtInDocumentProperties(project));
        appendDocumentProperties(report, "Custom document properties", customDocumentProperties(project));
    }

    private Map<String, Object> builtInDocumentProperties(Project project) {
        BuiltInProjectPropertyCollection properties = project.getBuiltInProps();
        Map<String, Object> values = new LinkedHashMap<>();
        values.put("Title", firstNonEmpty(properties.getTitle(), project.get(Prj.TITLE)));
        values.put("Subject", firstNonEmpty(properties.getSubject(), project.get(Prj.SUBJECT)));
        values.put("Author", firstNonEmpty(properties.getAuthor(), project.get(Prj.AUTHOR)));
        values.put("Manager", firstNonEmpty(properties.getManager(), project.get(Prj.MANAGER)));
        values.put("Company", firstNonEmpty(properties.getCompany(), project.get(Prj.COMPANY)));
        values.put("Category", properties.getCategory());
        values.put("Keywords", properties.getKeywords());
        values.put("Comments", firstNonEmpty(properties.getComments(), project.get(Prj.COMMENTS)));
        values.put("Creation Date", project.get(Prj.CREATION_DATE));
        values.put("Last Saved", project.get(Prj.LAST_SAVED));
        return values;
    }

    private Map<String, Object> customDocumentProperties(Project project) {
        Map<String, Object> values = new LinkedHashMap<>();
        for (Property property : project.getCustomProps()) {
            values.put(property.getName(), property.getValue());
        }
        return values;
    }

    private Object firstNonEmpty(Object first, Object second) {
        return formatValue(first).isEmpty() ? second : first;
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

    private void appendCustomFields(StringBuilder report, Project project) {
        report.append(System.lineSeparator());
        report.append("Custom fields").append(System.lineSeparator());
        report.append("-------------").append(System.lineSeparator());

        appendCustomFieldsForScope(report, "Task", project, index -> fieldId("Task", index), allTasks(project));
        appendCustomFieldsForScope(report, "Resource", project, index -> fieldId("Resource", index), resourcesWithoutPlaceholder(project));
        appendProjectCustomFields(report, project);
    }

    private List<Task> allTasks(Project project) {
        List<Task> tasks = new ArrayList<>();
        appendChildTasks(tasks, project.getRootTask());
        return tasks;
    }

    private void appendChildTasks(List<Task> tasks, Task parentTask) {
        for (Task task : parentTask.getChildren()) {
            tasks.add(task);
            appendChildTasks(tasks, task);
        }
    }

    private List<Resource> resourcesWithoutPlaceholder(Project project) {
        List<Resource> resources = new ArrayList<>();
        for (Resource resource : project.getResources()) {
            if (resource.getId() != 0 && !formatValue(resource.getName()).isEmpty()) {
                resources.add(resource);
            }
        }
        return resources;
    }

    private void appendCustomFieldsForScope(
            StringBuilder report,
            String scopeName,
            Project project,
            Function<CustomFieldIndex, Integer> fieldIdFactory,
            Iterable<? extends Object> items) {

        boolean foundAny = false;
        for (CustomFieldRange range : CUSTOM_FIELD_RANGES) {
            for (int index = 1; index <= range.count(); index++) {
                int fieldId = fieldIdFactory.apply(new CustomFieldIndex(range.fieldNamePart(), index));
                ExtendedAttributeDefinition definition = project.getExtendedAttributes().getById(fieldId);
                String alias = definition == null ? "" : formatValue(definition.getAlias());
                if (alias.isEmpty()) {
                    continue;
                }

                if (!foundAny) {
                    appendScopeHeading(report, scopeName);
                    foundAny = true;
                }

                report.append(range.displayPrefix()).append(index).append(": ")
                        .append(alias)
                        .append(" (").append(valueCountLabel(countValues(items, definition))).append(")")
                        .append(System.lineSeparator());
            }
        }

        if (!foundAny) {
            appendNoCustomFieldsForScope(report, scopeName);
        }
    }

    private void appendProjectCustomFields(StringBuilder report, Project project) {
        Task projectSummaryTask = project.getRootTask();
        boolean foundAny = false;
        for (CustomFieldRange range : CUSTOM_FIELD_RANGES) {
            for (int index = 1; index <= range.count(); index++) {
                int fieldId = fieldId("Task", new CustomFieldIndex(range.fieldNamePart(), index));
                ExtendedAttributeDefinition definition = project.getExtendedAttributes().getById(fieldId);
                String alias = definition == null ? "" : formatValue(definition.getAlias());
                if (alias.isEmpty()) {
                    continue;
                }

                if (!foundAny) {
                    appendScopeHeading(report, "Project");
                    foundAny = true;
                }

                String value = formatExtendedAttributeValue(projectSummaryTask.getExtendedAttributes(), definition);
                report.append(range.displayPrefix()).append(index).append(": ")
                        .append(alias)
                        .append(value.isEmpty() ? "" : " = " + value)
                        .append(" (").append(valueCountLabel(countValues(List.of(projectSummaryTask), definition))).append(")")
                        .append(System.lineSeparator());
            }
        }

        if (!foundAny) {
            appendNoCustomFieldsForScope(report, "Project");
        }
    }

    private int countValues(Iterable<? extends Object> items, ExtendedAttributeDefinition definition) {
        int count = 0;
        for (Object item : items) {
            ExtendedAttributeCollection attributes = extendedAttributes(item);
            if (!formatExtendedAttributeValue(attributes, definition).isEmpty()) {
                count++;
            }
        }
        return count;
    }

    private ExtendedAttributeCollection extendedAttributes(Object item) {
        if (item instanceof Task task) {
            return task.getExtendedAttributes();
        }
        if (item instanceof Resource resource) {
            return resource.getExtendedAttributes();
        }
        if (item instanceof ResourceAssignment assignment) {
            return assignment.getExtendedAttributes();
        }
        throw new IllegalArgumentException("Unsupported item type: " + item.getClass().getName());
    }

    private String valueCountLabel(int valueCount) {
        return valueCount == 1 ? "1 value" : valueCount + " values";
    }

    private String formatExtendedAttributeValue(ExtendedAttributeCollection attributes, String fieldName) {
        for (ExtendedAttribute attribute : attributes) {
            ExtendedAttributeDefinition definition = attribute.getAttributeDefinition();
            if (definition != null && fieldName.equalsIgnoreCase(definition.getFieldName())) {
                return formatExtendedAttributeValue(attribute, definition);
            }
        }
        return "";
    }

    private String formatExtendedAttributeValue(ExtendedAttributeCollection attributes, ExtendedAttributeDefinition definition) {
        for (ExtendedAttribute attribute : attributes) {
            ExtendedAttributeDefinition attributeDefinition = attribute.getAttributeDefinition();
            if (attributeDefinition != null && Objects.equals(attributeDefinition.getFieldId(), definition.getFieldId())) {
                return formatExtendedAttributeValue(attribute, definition);
            }
        }
        return "";
    }

    private String formatExtendedAttributeValue(ExtendedAttribute attribute, ExtendedAttributeDefinition definition) {
        String lookupValue = lookupValue(attribute, definition);
        if (!lookupValue.isEmpty()) {
            return lookupValue;
        }

        return switch (definition.getCfType()) {
            case CustomFieldType.Cost, CustomFieldType.Number -> formatValue(attribute.getNumericValue());
            case CustomFieldType.Date -> formatValue(attribute.getDateValue());
            case CustomFieldType.Duration -> formatValue(attribute.getDurationValue());
            case CustomFieldType.Flag -> formatValue(attribute.getFlagValue());
            default -> formatValue(attribute.getTextValue());
        };
    }

    private String lookupValue(ExtendedAttribute attribute, ExtendedAttributeDefinition definition) {
        String valueGuid = formatValue(attribute.getValueGuid());
        if (valueGuid.isEmpty()) {
            return "";
        }

        for (Value value : definition.getValueList()) {
            if (value.getValueGuid() != null && valueGuid.equalsIgnoreCase(value.getValueGuid().toString())) {
                return firstNonEmptyValue(value.getStringValue(), value.getVal(), value.getNumericValue(), value.getDateValue(), value.getDuration());
            }
        }
        return "";
    }

    private String firstNonEmptyValue(Object... values) {
        for (Object value : values) {
            String formatted = formatValue(value);
            if (!formatted.isEmpty()) {
                return formatted;
            }
        }
        return "";
    }

    private int fieldId(String scopeName, CustomFieldIndex index) {
        String fieldName = scopeName + index.fieldNamePart() + index.index();
        try {
            Field field = com.aspose.tasks.Field.class.getField(fieldName);
            return field.getInt(null);
        } catch (ReflectiveOperationException e) {
            throw new IllegalArgumentException("Aspose.Tasks field constant was not found: " + fieldName, e);
        }
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

    private record CustomFieldRange(String displayPrefix, String fieldNamePart, int count) {
    }

    private record CustomFieldIndex(String fieldNamePart, int index) {
    }
}
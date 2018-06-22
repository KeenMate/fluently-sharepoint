using System;
using System.Collections.Generic;
using System.Linq;
using KeenMate.FluentlySharePoint.Enums;
using KeenMate.FluentlySharePoint.Models;
using Microsoft.SharePoint.Client;
using ListTemplate = KeenMate.FluentlySharePoint.Enums.ListTemplate;

namespace KeenMate.FluentlySharePoint.Extensions
{
	public static class ListAlteration
	{
		public static CSOMOperation CreateList(this CSOMOperation operation, string name, string template = null)
		{
			operation.LogInfo($"Creating list {name}");

			ListCreationInformation listInformation = new ListCreationInformation
			{
				Title = name,
				ListTemplate = String.IsNullOrEmpty(template)
					? operation.LastWeb.ListTemplates.GetByName("Custom List")
					: operation.LastWeb.ListTemplates.GetByName(template)
			};

			var list = operation.LastWeb.Lists.Add(listInformation);

			operation.LastWeb.Context.Load(list);
			operation.SetLevel(OperationLevels.List, list);
			operation.ActionQueue.Enqueue(new DeferredAction { ClientObject = list, Action = DeferredActions.Load });

			return operation;
		}

		public static CSOMOperation CreateList(this CSOMOperation operation, string name, ListTemplate template)
		{
			return operation.CreateList(name, operation.LastWeb.ListTemplates.First(t => t.ListTemplateTypeKind == (int)template).Name);
		}

		public static CSOMOperation DeleteList(this CSOMOperation operation, string name)
		{
			operation.LogInfo($"Deleting list {name}");

			var list = operation.LoadedLists[name];

			operation.ActionQueue.Enqueue(new DeferredAction { ClientObject = list, Action = DeferredActions.Delete });

			return operation;
		}

		public static CSOMOperation ModifyList(this CSOMOperation operation, Action<Microsoft.SharePoint.Client.List> changes)
		{
			changes(operation.LastList);
			operation.LastList.Update();

			return operation;
		}

		public static CSOMOperation ModifyColumn(this CSOMOperation operation, string columnName, FieldType? type = null, string displayName = null, bool? required = null, bool? uniqueValues = null)
		{
			operation.LogInfo($"Modifying column {columnName}");

			var field = DecideFieldSource(operation).GetByInternalNameOrTitle(columnName);

			if (field == null)
				if (type.HasValue) field.TypeAsString = type.ToString();
			if (!String.IsNullOrEmpty(displayName)) field.Title = displayName;
			if (required.HasValue) field.Required = required.Value;
			if (uniqueValues.HasValue) field.EnforceUniqueValues = uniqueValues.Value;

			field.UpdateAndPushChanges(true);

			return operation;
		}

		public static CSOMOperation DeleteColumn(this CSOMOperation operation, string columnName)
		{
			operation.LogInfo($"Removing column {columnName}");

			var field = DecideFieldSource(operation).GetByInternalNameOrTitle(columnName);
			field.DeleteObject();

			return operation;
		}

		public static CSOMOperation AddChoiceField(this CSOMOperation operation, string name, List<string> choices, ChoiceTypes choiceType, string displayName = "", bool required = false, bool uniqueValues = false, string defaultValue = "", string group = "")
		{
			return operation.AddField(name, FieldType.Choice, displayName, required, uniqueValues, defaultValue, group, choices: choices, choiceType:choiceType);
		}

		public static CSOMOperation AddNumberField(this CSOMOperation operation, string name, string displayName = "",
			bool required = false, bool uniqueValues = false, string defaultValue = "", string group = "",
			bool percentage = false, int decimals = 2, int? min = null, int? max = null)
		{
			return operation.AddField(name, FieldType.Number, displayName, required, uniqueValues, defaultValue, group,
				null, percentage, decimals, max, min);
		}

		public static CSOMOperation AddTextField(this CSOMOperation operation, string name, string displayName = "",
			bool required = false, bool uniqueValues = false, string defaultValue = "", string group = "", int? maxLength = null)
		{
			return operation.AddField(name, FieldType.Text, displayName, required, uniqueValues,
				defaultValue, group, maxLength);
		}

		public static CSOMOperation AddLookupField(this CSOMOperation operation, string name, string list, string lookupField, string displayName = "", bool required = false, bool uniqueValues = false, string defaultValue = "", string group = "")
		{
			return operation.AddField(name, FieldType.Lookup, displayName, required, uniqueValues, defaultValue, group, lookupList:list, lookupField:lookupField);
		}

		/* Is uniqueValues required for boolean field? */
		public static CSOMOperation AddBooleanField(this CSOMOperation operation, string name, string displayName = "",
			bool required = false, bool uniqueValues = false, bool? defaultValue = null, string group = "")
		{
			return operation.AddField(name, FieldType.Boolean, displayName, required, uniqueValues, defaultValue.HasValue ? defaultValue.Value.ToString() : "", group);
		}

		//Generic method for all column types
		private static CSOMOperation AddField(this CSOMOperation operation, string name, FieldType type, string displayName = "", bool required = false, bool uniqueValues = false, string defaultValue = "", string group = "", int? maxLength = null, bool percentage = false, int decimals = 2, int? min = null, int? max = null, List<string> choices = null, ChoiceTypes choiceType = ChoiceTypes.Default, string lookupList = "", string lookupField = "")
		{
			operation.LogInfo($"Adding column {name}");

			FieldCreationInformation fieldInformation = new FieldCreationInformation
			{
				InternalName = name,
				DisplayName = String.IsNullOrEmpty(displayName) ? name : displayName,
				FieldType = type,
				Required = required,
				UniqueValues = uniqueValues,
				Group = group,
				Default = defaultValue,
				Percentage = percentage,
				Decimals = decimals,
				Min = min,
				Max = max,
				Choices = choices,
				Format = choiceType,
				List = lookupList,
				ShowField = lookupField,
				MaxLegth = maxLength
			};

			DecideFieldSource(operation).AddFieldAsXml(fieldInformation.ToXml(), true, AddFieldOptions.AddFieldInternalNameHint);

			return operation;
		}


		private static FieldCollection DecideFieldSource(CSOMOperation operation)
		{
			return operation.OperationLevel == OperationLevels.ContentType ? operation.LastContentType.Fields : operation.LastList.Fields;
		}
	}
}
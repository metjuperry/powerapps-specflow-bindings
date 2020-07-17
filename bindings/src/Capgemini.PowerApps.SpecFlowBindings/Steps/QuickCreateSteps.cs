﻿namespace Capgemini.PowerApps.SpecFlowBindings.Steps
{
    using System;
    using System.Globalization;
    using Capgemini.PowerApps.SpecFlowBindings.Extensions;
    using FluentAssertions;
    using Microsoft.Dynamics365.UIAutomation.Api.UCI;
    using TechTalk.SpecFlow;

    /// <summary>
    /// Step bindings related to quick creates.
    /// </summary>
    [Binding]
    public class QuickCreateSteps : PowerAppsStepDefiner
    {
        /// <summary>
        /// Sets the value for the field on the quick create.
        /// </summary>
        /// <param name="fieldValue">The field value.</param>
        /// <param name="fieldName">The field name.</param>
        /// <param name="fieldType">The field type.</param>
        [When(@"I enter '(.*)' into the '(.*)' (text|optionset|boolean|numeric|currency|datetime|lookup) field on the quick create")]
        public static void WhenIEnterInTheFieldOnTheQuickCreate(string fieldValue, string fieldName, string fieldType)
        {
            SetFieldValue(fieldName, fieldValue.ReplaceTemplatedText(), fieldType);
        }

        /// <summary>
        /// Sets the values for the fields in the table for the quick create.
        /// </summary>
        /// <param name="fields">Fields to be set.</param>
        [When(@"I enter the following into the quick create")]
        public static void WhenIEnterTheFollowingIntoTheQuickCreate(Table fields)
        {
            fields = fields ?? throw new ArgumentNullException(nameof(fields));

            foreach (TableRow row in fields.Rows)
            {
                SetFieldValue(row["Field"], row["Value"], row["Type"]);
            }
        }

        /// <summary>
        /// Cancels a quick create.
        /// </summary>
        [When(@"I cancel the quick create")]
        public static void WhenICancelTheQuickCreate()
        {
            XrmApp.QuickCreate.Cancel();
        }

        /// <summary>
        /// Clears the value for the field.
        /// </summary>
        /// <param name="field">The field name.</param>
        [When(@"I clear the '(.*)' (?:currency|numeric|text|datetime|boolean) field on the quick create")]
        public static void WhenIClearTheFieldOnTheQuickCreate(string field)
        {
            XrmApp.QuickCreate.ClearValue(field);
        }

        /// <summary>
        /// Clears the value for the optionset field.
        /// </summary>
        /// <param name="field">The field.</param>
        [When(@"I clear the '(.*)' optionset field on the quick create")]
        public static void WhenIClearTheOptionSetFieldOnTheQuickCreate(OptionSet field)
        {
            XrmApp.QuickCreate.ClearValue(field);
        }

        /// <summary>
        /// Clears the value for the lookup field.
        /// </summary>
        /// <param name="field">The field.</param>
        [When(@"I clear the '(.*)' lookup field on the quick create")]
        public static void WhenIClearTheLookupFieldOnTheQuickCreate(LookupItem field)
        {
            XrmApp.QuickCreate.ClearValue(field);
        }

        /// <summary>
        /// Saves a quick create.
        /// </summary>
        [When(@"I save the quick create")]
        public static void WhenISaveTheQuickCreate()
        {
            XrmApp.QuickCreate.Save();
        }

        /// <summary>
        /// Asserts that a value is shown in a text, currency or numeric field.
        /// </summary>
        /// <param name="expectedValue">The expected value.</param>
        /// <param name="field">The field name.</param>
        [Then("I can see a value of '(.*)' in the '(.*)' (?:currency|numeric|text) field on the quick create")]
        public static void ThenICanSeeAValueOfInTheFieldOnTheQuickCreate(string expectedValue, string field)
        {
            XrmApp.QuickCreate.GetValue(field).Should().Be(expectedValue);
        }

        /// <summary>
        /// Asserts that a value is shown in an option set field.
        /// </summary>
        /// <param name="expectedValue">The expected value.</param>
        /// <param name="field">The field name.</param>
        [Then("I can see a value of '(.*)' in the '(.*)' optionset field on the quick create")]
        public static void ThenICanSeeAValueOfInTheOptionSetFieldOnTheQuickCreate(string expectedValue, OptionSet field)
        {
            XrmApp.QuickCreate.GetValue(field).Should().Be(expectedValue);
        }

        /// <summary>
        /// Asserts that a value is shown in a lookup field.
        /// </summary>
        /// <param name="expectedValue">The expected value.</param>
        /// <param name="field">The field name.</param>
        [Then("I can see a value of '(.*)' in the '(.*)' lookup field on the quick create")]
        public static void ThenICanSeeAValueOfInTheOptionSetFieldOnTheQuickCreate(string expectedValue, LookupItem field)
        {
            XrmApp.QuickCreate.GetValue(field).Should().Be(expectedValue);
        }

        /// <summary>
        /// Asserts that a value is shown in a lookup field.
        /// </summary>
        /// <param name="expectedValue">The expected value.</param>
        /// <param name="field">The field name.</param>
        [Then("I can see a value of '(true|false)' in the '(.*)' boolean field on the quick create")]
        public static void ThenICanSeeAValueOfInTheBooleanFieldOnTheQuickCreate(bool expectedValue, BooleanItem field)
        {
            XrmApp.QuickCreate.GetValue(field).Should().Be(expectedValue);
        }

        private static void SetFieldValue(string fieldName, string fieldValue, string fieldType)
        {
            switch (fieldType)
            {
                case "optionset":
                    XrmApp.QuickCreate.SetValue(new OptionSet()
                    {
                        Name = fieldName,
                        Value = fieldValue,
                    });
                    break;
                case "boolean":
                    XrmApp.QuickCreate.SetValue(new BooleanItem()
                    {
                        Name = fieldName,
                        Value = bool.Parse(fieldValue),
                    });
                    break;
                case "datetime":
                    XrmApp.QuickCreate.SetValue(fieldName, DateTime.Parse(fieldValue, CultureInfo.CreateSpecificCulture("en-GB")));
                    break;
                case "lookup":
                    XrmApp.QuickCreate.SetValue(new LookupItem()
                    {
                        Name = fieldName,
                        Value = fieldValue,
                    });
                    break;
                case "currency":
                case "numeric":
                case "text":
                default:
                    XrmApp.QuickCreate.SetValue(fieldName, fieldValue);
                    break;
            }
        }
    }
}

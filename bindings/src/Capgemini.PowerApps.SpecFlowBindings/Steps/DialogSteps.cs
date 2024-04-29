namespace Capgemini.PowerApps.SpecFlowBindings.Steps
{
    using Capgemini.PowerApps.SpecFlowBindings.Extensions;
    using FluentAssertions;
    using Microsoft.Dynamics365.UIAutomation.Api.UCI;
    using Microsoft.Dynamics365.UIAutomation.Browser;
    using OpenQA.Selenium;
    using System;
    using System.Globalization;
    using System.Linq;
    using TechTalk.SpecFlow;

    /// <summary>
    /// Step bindings related to dialogs.
    /// </summary>
    [Binding]
    public class DialogSteps : PowerAppsStepDefiner
    {
        /// <summary>
        /// Clicks the confirmation button on a confirm dialog.
        /// </summary>
        /// <param name="option">The option to click.</param>
        [When(@"I (confirm|cancel) when presented with the confirmation dialog")]
        public static void WhenIConfirmWhenPresentedWithTheConfirmationDialog(string option)
        {
            XrmApp.Dialogs.ConfirmationDialog(option == "confirm");
            XrmApp.ThinkTime(2000);
        }

        /// <summary>
        /// Assigns to the current user.
        /// </summary>
        [When("I assign to me on the assign dialog")]
        public static void WhenIAssignToMeOnTheAssignDialog()
        {
            XrmApp.Dialogs.Assign(Dialogs.AssignTo.Me);
        }

        /// <summary>
        /// Assigns to a user or team with the given name.
        /// </summary>
        /// <param name="assignTo">User or team.</param>
        /// <param name="userName">The name of the user or team.</param>
        [When("I assign to a (user|team) named '(.*)' on the assign dialog")]
        public static void WhenIAssignToANamedOnTheAssignDialog(Dialogs.AssignTo assignTo, string userName)
        {
            XrmApp.Dialogs.Assign(assignTo, userName);
        }

        /// <summary>
        /// Closes an opportunity.
        /// </summary>
        /// <param name="status">Whether the opportunity was won.</param>
        [When("I close the opportunity as (won|lost)")]
        public static void WhenICloseTheOpportunityAs(string status)
        {
            XrmApp.Dialogs.CloseOpportunity(status == "won");
        }

        /// <summary>
        /// Closes a warning dialog.
        /// </summary>
        [When("I close the warning dialog")]
        public static void WhenICloseTheWarningDialog()
        {
            XrmApp.Dialogs.CloseWarningDialog();
        }

        /// <summary>
        /// Clicks an option on the publish dialog.
        /// </summary>
        /// <param name="option">The option to click.</param>
        [When("I click (confirm|cancel) on the publish dialog")]
        public static void WhenIClickOnThePublishDialog(string option)
        {
            XrmApp.Dialogs.PublishDialog(option == "confirm");
        }

        /// <summary>
        /// Click a button on a custom dialog
        /// </summary>
        /// <param name="buttonName">The option to click.</param>
        /// <param name="dialogName">Name of the dialog.</param>
        [When(@"I click the '(.*)' button on the '(.*)' custom dialog")]
        public static void WhenIClickButtonOnCustomDialog(string buttonName, string dialogName)
        {
            var dialog = Driver.FindElement(By.XPath($"//div[@role='dialog' and @data-id='{dialogName}']"));

            var childButton = dialog.FindElement(By.XPath($".//button[@data-id='{buttonName}']"));

            childButton.Click();

            Driver.WaitForPageToLoad();
            Driver.WaitForTransaction();
        }

        /// <summary>
        /// Sets the value for the field.
        /// </summary>
        /// <param name="fieldValue">The field value.</param>
        /// <param name="fieldName">The field name.</param>
        /// <param name="fieldType">The field type.</param>
        [When(@"I enter '(.*)' into the '(.*)' (text|optionset|multioptionset|boolean|numeric|currency|datetime|lookup) (field|header field) on the custom  dialog")]
        public static void WhenIEnterInTheFieldOnCustomDialog(string fieldValue, string fieldName, string fieldType)
        {

            SetCustomDialogFieldValue(fieldName, fieldValue.ReplaceTemplatedText(), fieldType);

            Driver.WaitForTransaction();
        }

        /// <summary>
        /// Sets the values of the fields in the table on the form.
        /// </summary>
        /// <param name="fields">The fields to set.</param>
        [When(@"I enter the following into the custom dialog")]
        public static void WhenIEnterTheFollowingIntoTheCustomDialog(Table fields)
        {
            fields = fields ?? throw new ArgumentNullException(nameof(fields));

            foreach (TableRow row in fields.Rows)
            {
                WhenIEnterInTheFieldOnCustomDialog(row["Value"], row["Field"], row["Type"]);
            }
        }

        /// <summary>
        /// Clicks an option on the set state dialog.
        /// </summary>
        /// <param name="option">The option to click.</param>
        [When("I click (ok|cancel) on the set state dialog")]
        public static void WhenIClickOnTheSetStateDialog(string option)
        {
            XrmApp.Dialogs.SetStateDialog(option == "ok");
        }

        /// <summary>
        /// Check if dialog is displayed.
        /// </summary>
        [Then(@"an alert dialog is displayed")]
        public static void ThenAlertDialogIsDisplayed()
        {
            var dialog = Driver.FindElement(By.XPath("//div[@data-id='alertdialog']"));

            dialog.Should().NotBeNull();
        }

        private static void SetCustomDialogFieldValue(string fieldName, string fieldValue, string fieldType)
        {
            switch (fieldType)
            {
                case "multioptionset":
                    XrmApp.Entity.SetValue(
                        new MultiValueOptionSet()
                        {
                            Name = fieldName,
                            Values = fieldValue
                                        .Split(',')
                                        .Select(v => v.Trim())
                                        .ToArray(),
                        },
                        true);
                    break;
                case "optionset":
                    XrmApp.Entity.SetValue(new OptionSet()
                    {
                        Name = fieldName,
                        Value = fieldValue,
                    });
                    break;
                case "boolean":
                    XrmApp.Entity.SetValue(new BooleanItem()
                    {
                        Name = fieldName,
                        Value = bool.Parse(fieldValue),
                    });
                    break;
                case "datetime":
                    XrmApp.Entity.SetValue(new DateTimeControl(fieldName)
                    {
                        Value = DateTime.Parse(fieldValue, CultureInfo.CurrentCulture),
                    });
                    break;
                case "lookup":
                    XrmApp.Entity.SetValue(new LookupItem()
                    {
                        Name = fieldName,
                        Value = fieldValue,
                    });
                    break;
                case "currency":
                case "numeric":
                case "text":
                default:
                    XrmApp.Entity.SetValue(fieldName, fieldValue);
                    break;
            }
        }
    }
}

namespace Capgemini.PowerApps.SpecFlowBindings.Steps
{
    using FluentAssertions;
    using Microsoft.Dynamics365.UIAutomation.Api.UCI;
    using Microsoft.Dynamics365.UIAutomation.Browser;
    using OpenQA.Selenium;
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
        /// Check if dialog is displayed.
        /// </summary>
        [Then(@"an alert dialog is displayed")]
        public static void ThenAlertDialogIsDisplayed()
        {
            var dialog = Driver.FindElement(By.XPath("//div[@data-id='alertdialog']"));

            dialog.Should().NotBeNull();
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
        /// Clicks an option on the set state dialog.
        /// </summary>
        /// <param name="option">The option to click.</param>
        [When("I click (ok|cancel) on the set state dialog")]
        public static void WhenIClickOnTheSetStateDialog(string option)
        {
            XrmApp.Dialogs.SetStateDialog(option == "ok");
        }
    }
}

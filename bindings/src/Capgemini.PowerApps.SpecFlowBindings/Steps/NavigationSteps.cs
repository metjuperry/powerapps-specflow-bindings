namespace Capgemini.PowerApps.SpecFlowBindings.Steps
{
    using System;
    using System.Linq;
    using FluentAssertions;
    using Microsoft.Dynamics365.UIAutomation.Api.UCI;
    using Microsoft.Dynamics365.UIAutomation.Browser;
    using OpenQA.Selenium;
    using TechTalk.SpecFlow;

    /// <summary>
    /// Step bindings relating to navigation.
    /// </summary>
    [Binding]
    public class NavigationSteps : PowerAppsStepDefiner
    {
        /// <summary>
        /// Opens a sub-area from the navigation bar.
        /// </summary>
        /// <param name="subAreaName">The name of the sub-area.</param>
        /// <param name="areaName">The name of the area.</param>
        [When("I open the sub area '(.*)' under the '(.*)' area")]
        public static void WhenIOpenTheSubAreaUnderTheArea(string subAreaName, string areaName)
        {
            XrmApp.Navigation.OpenSubArea(areaName, subAreaName);
        }

        /// <summary>
        /// Opens global search.
        /// </summary>
        [When("I open global search")]
        public static void WhenIOpenGlobalSearch()
        {
            XrmApp.Navigation.OpenGlobalSearch();
        }

        /// <summary>
        /// Opens an area.
        /// </summary>
        /// <param name="areaName">The name of the area.</param>
        [When("I open the '(.*)' area")]
        public static void WhenIOpenTheArea(string areaName)
        {
            XrmApp.Navigation.OpenArea(areaName);
        }

        /// <summary>
        /// Opens a sub area of a group.
        /// </summary>
        /// <param name="subAreaName">The name of the sub area.</param>
        /// <param name="groupName">The name of the group.</param>
        [When("I open the '(.*)' sub area of the '(.*)' group")]
        public static void WhenIOpenTheSubAreaOfTheGroup(string subAreaName, string groupName)
        {
            XrmApp.Navigation.OpenGroupSubArea(groupName, subAreaName);
        }

        /// <summary>
        /// Opens a sub area of a group.
        /// </summary>
        /// <param name="subAreaName">The name of the sub area.</param>
        /// <param name="groupName">The name of the group.</param>
        [When("I open the '(.*)' sub area under the '(.*)' group")]
        public static void WhenIOpenTheSubAreaOfTheGroupManually(string subAreaName, string groupName)
        {
            if (Driver.HasElement(By.XPath(AppElements.Xpath[AppReference.Navigation.SiteMapLauncherButton])))
            {
                var expanded = bool.Parse(Driver.FindElement(By.XPath(AppElements.Xpath[AppReference.Navigation.SiteMapLauncherButton])).GetAttribute("aria-expanded"));

                if (!expanded)
                {
                    Driver.ClickWhenAvailable(By.XPath(AppElements.Xpath[AppReference.Navigation.SiteMapLauncherButton]));
                }
            }

            var groups = Driver.FindElements(By.XPath(AppElements.Xpath[AppReference.Navigation.SitemapMenuGroup]));
            var groupList = groups.FirstOrDefault(g => g.GetAttribute("aria-label").ToLowerString() == groupName.ToLowerString());

            if (groupList == null)
            {
                throw new NotFoundException($"No group with the name '{groupName}' exists");
            }

            var subAreaItems = groupList.FindElements(By.XPath("." + AppElements.Xpath[AppReference.Navigation.SitemapMenuItems]));
            var subAreaItem = subAreaItems.FirstOrDefault(a => a.GetAttribute("data-text").ToLowerString() == subAreaName.ToLowerString());

            if (subAreaItem == null)
            {
                throw new NotFoundException($"No subarea with the name '{subAreaName}' exists inside of '{groupName}'");
            }

            subAreaItem.Click(true);

            Driver.WaitForPageToLoad();
            Driver.WaitForTransaction();
        }


        /// <summary>
        /// Opens a quick create for an entity.
        /// </summary>
        /// <param name="entity">The name of the entity.</param>
        [When("I open a quick create for the '(.*)' entity")]
        public static void WhenIOpenAQuickCreateForTheEntity(string entity)
        {
            XrmApp.Navigation.QuickCreate(entity);
        }

        /// <summary>
        /// Signs out.
        /// </summary>
        [When("I sign out")]
        public static void WhenISignOut()
        {
            XrmApp.Navigation.SignOut();
        }

        /// <summary>
        /// Gets the area.
        /// </summary>
        /// <param name="area">The area name.</param>
        [Then(@"I see the '(.*)' area")]
        public static void ThenICanSeeTheArea(string area)
        {
            Driver
                .WaitUntilAvailable(By.Id("areaSwitcherContainer"))
                .Text
                .Should().Contain(area);
        }

        /// <summary>
        /// Gets the sub-area.
        /// </summary>
        /// <param name="subArea">The area name.</param>
        [Then(@"I see the '(.*)' subarea")]
        public static void ThenICanSeeTheSubArea(string subArea)
        {
            Driver
                .WaitUntilAvailable(By.XPath($"//img[@title='{subArea}']"))
                .Text
                .Should().NotBeNull();
        }

        /// <summary>
        /// Gets the sub-area.
        /// </summary>
        /// <param name="groupName">The group.</param>
        [Then(@"I see the '(.*)' group")]
        public static void ThenICanSeeTheGroup(string groupName)
        {
            var groupNameWithoutWhitespace = groupName?.Replace(" ", string.Empty);

            Driver
                .WaitUntilAvailable(By.XPath($"//h3[@data-id='sitemap-sitemapAreaGroup-{groupNameWithoutWhitespace}']"))
                .Text
                .Should().Contain(groupName);
        }
    }
}

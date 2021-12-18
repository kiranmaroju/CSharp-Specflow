using OpenQA.Selenium;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using TechTalk.SpecFlow;
using Learn_Specflow.Hooks;
using NUnit.Framework;
using OpenQA.Selenium.Support.UI;
using Ledger_AutomationTesting.ExcelUtilities;
using Learn_Specflow.PageObjects;

namespace Learn_Specflow.Steps
{
    [Binding]
    public sealed class FujitsuTestsSteps
    {
        public int items=0;
        IWebDriver _driver;
        
        HomePage home = new HomePage();
        public FujitsuTestsSteps() => _driver = Hook.GetDriver();
        ExcelLib excel = new ExcelLib();
        


        //==============================================================================
        [Given(@"I Login to the Shopping Website")]
        public void GivenILoginToTheShoppingWebsite()
        {
                    
            home.NavigateToHome();
            home.Login(excel.getValueFromExcel("Login", 1, 2), excel.getValueFromExcel("Login", 1, 3));
        }


        [Given(@"Add that item '(.*)' with size '(.*)' quanity '(.*)' to basket")]
        public void GivenAddThatItemWithSizeQuanityToBasket(string p0, string p1, int p2)
        {
            
            //HomePage home = new HomePage();
            items = items + 1;
            home.AddItemWithQuickView(p0, p1, p2, items);
        }

        [Then(@"Verify the basket Items & Price")]
        public void ThenVerifyTheBasketItemsPrice()
        {
            home.VerifyCart();
        }



        [Then(@"Proceed to checkout with '(.*)'")]
        public void ThenProceedToCheckoutWith(string p0)
        {
            home.ProceedCheckout(p0);
        }

        [Then(@"Select previous order and comment")]
        public void ThenSelectPreviousOrderAndComment()
        {
            home.SelectLastOrder();
            home.SendMessgeOnOrder("Some message");
           
        }

        [Then(@"Verify the message")]
        public void ThenVerifyTheMessage()
        {
            home.VerifyMessgeOnOrder("Some message");
        }


        [Then(@"Select previous order")]
        public void ThenSelectPreviousOrder()
        {
            home.SelectLastOrder();
        }

        [Then(@"Verify item color is '(.*)'")]
        public void ThenVerifyItemColorIs(string p0)
        {
            home.VerifyItemColor(p0);
        }



        [Then(@"I Logout")]
        public void ThenILogout()
        {
            HomePage home = new HomePage();
            home.Logout();
        }

    }
}

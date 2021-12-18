using Learn_Specflow.Hooks;
using Learn_Specflow.SeleniumMethods ;
using OpenQA.Selenium;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Learn_Specflow.Config;
using NUnit.Framework;
using OpenQA.Selenium.Support.UI;
using Ledger_AutomationTesting.ExcelUtilities;
using System.Data;
using System.Data.OleDb;
using OpenQA.Selenium.Interactions;
using System.Threading;
using System.Collections.ObjectModel;
using Learn_Specflow.Steps;
namespace Learn_Specflow.PageObjects
{
    
    class HomePage
    {
        private IWebDriver _driver;

        
        DataTable dt = new DataTable();        
        public HomePage() => _driver = Hook.GetDriver();
        ExcelLib excel = new ExcelLib();
        
        IWebElement loginLink => _driver.FindElement(By.XPath("//*[@class='login']"));
        IWebElement emailField => _driver.FindElement(By.XPath("//input[@id='email']"));
        IWebElement passwdField => _driver.FindElement(By.XPath("//input[@id='passwd']"));
        IWebElement submitButton => _driver.FindElement(By.XPath("//button[@id='SubmitLogin']"));
        IWebElement accountInfo => _driver.FindElement(By.XPath("//p[@class='info-account']"));
        
        IWebElement logoutLink => _driver.FindElement(By.XPath("//*[@class='logout']"));

        IWebElement qtyField => _driver.FindElement(By.XPath("//input[@name='qty']"));
        IWebElement sizeField => _driver.FindElement(By.XPath("//*[contains(text(), 'Size')]/following-sibling::div//select"));
        IWebElement priceText => _driver.FindElement(By.XPath("//span[@id='our_price_display']"));

        IWebElement addToCartButton => _driver.FindElement(By.XPath("//button/span[contains(text(), 'Add to cart')]"));
        IWebElement okIcon => _driver.FindElement(By.XPath("//i[@class='icon-ok']"));
        IWebElement closeWindow => _driver.FindElement(By.XPath("//span[@title='Close window']"));
        IWebElement cartLink => _driver.FindElement(By.XPath("//div[@class='shopping_cart']/a"));

        IWebElement totalProductsPrice => _driver.FindElement(By.XPath("//table[@id='cart_summary']//td[@id='total_product']"));
        IWebElement totalShippingPrice => _driver.FindElement(By.XPath("//table[@id='cart_summary']//td[@id='total_shipping']"));

        IWebElement totalPriceWithoutTax => _driver.FindElement(By.XPath("//table[@id='cart_summary']//td[@id='total_price_without_tax']"));

        IWebElement totalTax => _driver.FindElement(By.XPath("//table[@id='cart_summary']//td[@id='total_tax']"));
        IWebElement totalPrice => _driver.FindElement(By.XPath("//table[@id='cart_summary']//td[@id='total_price_container']"));
        IWebElement checkoutLink => _driver.FindElement(By.XPath("//div[@class='columns-container']//span[contains(text(), 'Proceed to checkout')]"));
        IWebElement termsAcceptField => _driver.FindElement(By.XPath("//input[@id='cgv']"));
        IWebElement bankWireLink => _driver.FindElement(By.XPath("//a[@class='bankwire']"));
        IWebElement bankCheckLink => _driver.FindElement(By.XPath("//a[@class='cheque']"));
        IWebElement orderConfLink => _driver.FindElement(By.XPath("//button[@type='submit']/span[contains(text(), 'confirm my order')]"));
        IWebElement confMessage => _driver.FindElement(By.XPath("//*[@id='order-confirmation']//div[@class='box']"));

        IWebElement ordersLink => _driver.FindElement(By.XPath("//span[contains(text(), 'Order history and details')]"));
        IWebElement latestOrderLink => _driver.FindElement(By.XPath("(//tr[@class='first_item ']//a)[1]"));
        IWebElement latestOrderDate => _driver.FindElement(By.XPath("//tr[@class='first_item ']/td[@class='history_date bold']"));
        IWebElement msgTextArea => _driver.FindElement(By.XPath("//form[@id='sendOrderMessage']//textarea[@name='msgText']"));
        IWebElement submitMessageButton => _driver.FindElement(By.XPath("//form[@id='sendOrderMessage']//button[@name='submitMessage']"));

        IWebElement submittedMessage => _driver.FindElement(By.XPath("(//tr[@class='first_item item'])[2]//td[2]"));

        IWebElement itemDetails => _driver.FindElement(By.XPath("//div[@id='order-detail-content']//tbody/tr[@class='item']/td[2]"));


        

        public void NavigateToHome()
        {            
            string url= excel.getValueFromExcel("Login", 1, 1);
            _driver.Navigate().GoToUrl(url);
            Assert.IsTrue(_driver.Title.ToLower().Contains("my store"));
        }

        //Function Name: Login
        //Parameters: 
                   // 1. Email: String format and Mandatory
                   // 2. Passowrd: String format and Mandatory
        public void Login(string email, string password)
        {
            loginLink.Clicks();
            emailField.EnterText(email);
            passwdField.EnterText(password);
            submitButton.Clicks();
            String account_info = accountInfo.Text;
            Assert.IsTrue(account_info.ToLower().Contains("welcome"));
            Console.WriteLine("++++++++++++++++++++++++"+account_info);
        }


        public void AddItemWithQuickView(string itemName, string size, int qty, int items)
        {
            
            
            excel.writeToExcel("Items", items + 1, 2, itemName);
            excel.writeToExcel("Items", items + 1, 3, size);
            excel.writeToExcel("Items", items + 1, 4, qty.ToString());
            NavigateToHome();
            IWebElement productLink = _driver.FindElement(By.XPath("(//a[@title='" + itemName + "'][@class='product_img_link'])[1]"));
            Actions actions = new Actions(_driver);
            actions.MoveToElement(productLink).Perform();
            
            IWebElement quickViewLink = _driver.FindElement(By.XPath("//a[@title='"+ itemName + "']//following-sibling::a/span[contains(text(), 'Quick view')]"));
            try
            {
                bool productDisplayed = quickViewLink.Displayed;
                WebDriverWait wait = new WebDriverWait(_driver, TimeSpan.FromSeconds(30));
                
                ReadOnlyCollection<string> ss= _driver.WindowHandles;
                quickViewLink.Clicks();
                _driver.SwitchTo().ParentFrame();

                int fr=getFrameId(By.XPath("//body[@id='product']"), _driver);


                _driver.SwitchTo().Frame(fr);
                qtyField.Clear();
                qtyField.EnterText(qty.ToString());
                
                var selectElement = new SelectElement(sizeField);
                selectElement.SelectByText(size);

                string price = priceText.Text.Replace("$", "");
                Decimal priceVal = Math.Round(Convert.ToDecimal(price), 2);
                excel.writeToExcel("Items", items + 1, 5, priceVal.ToString());
                addToCartButton.Clicks();
                Thread.Sleep(3000);
                bool okIconDisplayed = okIcon.Displayed;

                if (okIconDisplayed)
                {
                    
                }
                else
                {
                    Assert.Fail(itemName + "-Item NOT added to cart, check logs");
                }

                
                closeWindow.Clicks();
                _driver.SwitchTo().DefaultContent();
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message+e.StackTrace);
                Assert.Fail(itemName + "-Item NOT added to cart, check logs");
                closeWindow.Clicks();
                _driver.SwitchTo().DefaultContent();
            }

            
            
        }


        public void VerifyCart()
        {
            
            NavigateToHome();
            Actions actions = new Actions(_driver);
            actions.MoveToElement(cartLink).Perform();
            cartLink.Clicks();

            DataTable dt = excel.ExcelToDataTable( "Items");            

            int products = dt.Rows.Count;
            Decimal expTotalProductsPriceVal = 0;
            for (int p = 0; p < products; p++)
            {
                try
                {
                    string expProductName = dt.Rows[p][1].ToString();
                    int expProductQty = int.Parse(dt.Rows[p][3].ToString());
                    Decimal expProductPrice = Math.Round(Convert.ToDecimal(dt.Rows[p][4].ToString()), 2);
                    expTotalProductsPriceVal = expTotalProductsPriceVal + (expProductPrice * expProductQty);
                    int id = p + 1;
                    IWebElement actItemNameElm = _driver.FindElement(By.XPath("//table[@id='cart_summary']//tr[contains(@id, 'product')][" + id.ToString() + "]//*[@class='product-name']"));
                    IWebElement actItemQtyElm = _driver.FindElement(By.XPath("//table[@id='cart_summary']//tr[contains(@id, 'product')][" + id.ToString() + "]//*[@class='cart_quantity_input form-control grey']"));
                    IWebElement actItemUnitPriceElm = _driver.FindElement(By.XPath("//table[@id='cart_summary']//tr[contains(@id, 'product')][" + id.ToString() + "]//*[@class='cart_unit']//span[@class='price']"));
                    Decimal actItemUnitPrice = Math.Round(Convert.ToDecimal(actItemUnitPriceElm.Text.Replace("$", "")), 2);
                    int actItemQty = int.Parse(actItemQtyElm.GetProperty("value"));


                    Assert.AreEqual(expProductName, actItemNameElm.Text);
                    Assert.AreEqual(expProductPrice, actItemUnitPrice);
                    Assert.AreEqual(expProductQty, actItemQty);
                }
                catch (Exception e) { Console.WriteLine(e.Message + e.StackTrace); }
                
            }

            Decimal actTotalProductsPriceVal = Math.Round(Convert.ToDecimal(totalProductsPrice.Text.Replace("$", "")), 2);
            Decimal totalShippingPriceVal = Math.Round(Convert.ToDecimal(totalShippingPrice.Text.Replace("$", "")), 2);
            Decimal actTotalPriceWithoutTaxVal = Math.Round(Convert.ToDecimal(totalPriceWithoutTax.Text.Replace("$", "")), 2);
            Decimal totalTaxVal = Math.Round(Convert.ToDecimal(totalTax.Text.Replace("$", "")), 2);
            Decimal acttotalPriceVal = Math.Round(Convert.ToDecimal(totalPrice.Text.Replace("$", "")), 2);

            Assert.AreEqual(expTotalProductsPriceVal, actTotalProductsPriceVal);
            Assert.AreEqual(expTotalProductsPriceVal+ totalShippingPriceVal, actTotalPriceWithoutTaxVal);
            Assert.AreEqual(expTotalProductsPriceVal + totalShippingPriceVal+ totalTaxVal, acttotalPriceVal);
        }
        
        public void ProceedCheckout(string modeOfPayment)
        {
            checkoutLink.Clicks();
            checkoutLink.Clicks();
            termsAcceptField.Clicks();
            checkoutLink.Clicks();
            switch (modeOfPayment)
            {
                case "Wire":
                    bankWireLink.Clicks();
                    break;
                case "Check":
                    bankCheckLink.Clicks();
                    break;
                default:
                    bankWireLink.Clicks();
                    break;
            }

            orderConfLink.Clicks();

            Thread.Sleep(3000);
            string orderConfMsg=confMessage.Text;

            Assert.IsTrue(orderConfMsg.ToLower().Contains("your order on my store is complete."));

        }


        public string SelectLastOrder()
        {

            ordersLink.Clicks();
            string lastOrderID=latestOrderLink.Text;
            string lastOrderDate = latestOrderDate.Text;
            latestOrderLink.Clicks();


            return lastOrderID;
        }


        public void SendMessgeOnOrder(string message)
        {
            latestOrderLink.Clicks();
            msgTextArea.EnterText(message);
            submitMessageButton.Clicks();
        }


        public void VerifyMessgeOnOrder(string message)
        {
            string actMesg=submittedMessage.Text;
            Assert.IsTrue(actMesg.ToLower().Contains(message.ToLower()));
        }


        public void VerifyItemColor(string color)
        {
            itemDetails.Clicks();
            string itemdetails = itemDetails.Text;
            
            if (itemdetails.ToLower().Contains(color.ToLower()))
            { }
            else
            {
                string errorImage = Hook.TakeScreenshot(_driver);
            }
            
            Assert.IsTrue(itemdetails.ToLower().Contains(color.ToLower()));
        }



        public void Logout()
        {
            logoutLink.Clicks();            
            Assert.IsTrue(_driver.Title.ToLower().Contains("login"));            
        }


        public static string ClickAndSwitchWindow(IWebElement elementToBeClicked,IWebDriver driver, int timer = 2000)
        {
            List<string> previousHandles = new List<string>();
            List<string> currentHandles = new List<string>();
            previousHandles.AddRange(driver.WindowHandles);
            elementToBeClicked.Click();

            Thread.Sleep(timer);
            for (int i = 0; i < 20; i++)
            {
                currentHandles.Clear();
                currentHandles.AddRange(driver.WindowHandles);
                foreach (string s in previousHandles)
                {
                    currentHandles.RemoveAll(p => p == s);
                }
                if (currentHandles.Count == 1)
                {
                    driver.SwitchTo().Window(currentHandles[0]);
                    Thread.Sleep(100);
                    return currentHandles[0];
                }
                else
                {
                    Thread.Sleep(500);
                }
            }
            return null;
        }

        public int getFrameId(By by,IWebDriver driver)
        {
            IList<IWebElement> frameCount = driver.FindElements(By.TagName("iframe"));
            bool IsElementPresent = false;
            IWebElement locator;
            int index=-1;
            for (int i = frameCount.Count - 1; i <= frameCount.Count-1; i--)
            {
                driver.SwitchTo().Frame(i);
                try
                {
                    locator = driver.FindElement(by);
                    if (locator.Displayed == true)
                    {
                        IsElementPresent = true;
                        Console.WriteLine("Index: " + i);
                        driver.SwitchTo().DefaultContent();
                        index = i;
                        break;
                    }
                    else
                    {

                    }
                }
                catch(Exception e){ Console.WriteLine(e.Message + e.StackTrace); }


            }
            return index;

        }

        
    }
}

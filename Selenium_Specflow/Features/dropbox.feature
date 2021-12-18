Feature: FujitsuTests
	

@test123
Scenario: 01)_HAPPY_PATH_Add 2 Items to Cart
	Given I Login to the Shopping Website
	And Add that item 'Blouse' with size 'M' quanity '2' to basket
	And Add that item 'Faded Short Sleeve T-shirts' with size 'M' quanity '1' to basket
	Then Verify the basket Items & Price
	Then Proceed to checkout with 'Wire'
	Then I Logout


@test123
Scenario: 02)_Review_Previous_Orders_And Add a Message
	Given I Login to the Shopping Website
	Then Select previous order and comment
	Then Verify the message
	Then I Logout

@test123
Scenario: 03)_CAPTURE_ERROR_IMAGES
	Given I Login to the Shopping Website
	Then Select previous order
	Then Verify item color is 'Red'
	Then I Logout

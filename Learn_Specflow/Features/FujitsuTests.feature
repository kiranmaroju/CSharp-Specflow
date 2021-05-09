Feature: FujitsuTests
	

@test123
Scenario: HAPPY_PATH  
	Given I Login to the Shopping Website
	And Add that item 'Blouse' with size 'M' quanity '2' to basket
	And Add that item 'Printed Chiffon Dress' with size 'M' quanity '1' to basket
	Then Verify the basket Items & Price
	Then Proceed to checkout with 'Wire'
	Then I Logout


@test123
Scenario: REVIEW PREVIOUS ORDERS AND ADD A MESSAGE
	Given I Login to the Shopping Website
	Then Select previous order and comment
	Then Verify the message
	Then I Logout
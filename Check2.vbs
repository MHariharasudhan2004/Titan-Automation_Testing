Function Mod1()
'Titan logo section'
	If Browser("Titan: The Official Website").Page("Titan: The Official Website").Image("TITAN").Exist(5)  Then
		Reporter.ReportEvent micPass,"Titan logo image","The titan logo image exists and working"
		Browser("Titan: The Official Website").Page("Titan: The Official Website").Image("TITAN").Highlight
		Reporter.ReportEvent micPass,"Titan logo image Highlighted","The titan logo image got Highlighted"
		Browser("Titan: The Official Website").Page("Titan: The Official Website").Image("TITAN").Click
		Reporter.ReportEvent micPass,"Titan logo image Clicked","The titan logo image got Clicked"
	Else
		Reporter.ReportEvent micFail,"Titan logo image didn't work","The titan logo image didn't work properly."
	End If
	'Serach image section'
	If Browser("Titan: The Official Website").Page("Titan: The Official Website").Image("Search_image").Exist(5)  Then
		 Reporter.ReportEvent micPass,"Search image exists","The search image exists and works"
		 Browser("Titan: The Official Website").Page("Titan: The Official Website").Image("Search_image").Highlight
		 Reporter.ReportEvent micPass,"Search image Highlighted","The search image got highlighted"
		 Browser("Titan: The Official Website").Page("Titan: The Official Website").Image("Search_image").Click
		 Reporter.ReportEvent micPass,"Search image Clicked","The search image got clicked"
		 Browser("Titan: The Official Website").Page("Titan: The Official Website").Image("TITAN").Click
		 Reporter.ReportEvent micPass,"Back to home page","Gone back to home page"
	Else
		Reporter.ReportEvent micFail,"Search image didn't work","The search image worked"
	End If
	'Account image section'
	If Browser("Titan: The Official Website").Page("Titan: The Official Website").Image("Acc_img").Exist(5) Then
		Reporter.ReportEvent micPass,"Account image exist","Account image exists and working"
		Browser("Titan: The Official Website").Page("Titan: The Official Website").Image("Acc_img").Highlight
		Reporter.ReportEvent micPass,"Account image Highlighted","Account image got highlighted"
		Browser("Titan: The Official Website").Page("Titan: The Official Website").Link("Account").Click
		Reporter.ReportEvent micPass,"Account image link clicked","The  account image link got clicked"
		If Browser("Titan: The Official Website").Page("Titan: The Official Website").WebElement("login_desc").Exist(5) Then
			Reporter.ReportEvent micPass,"Login description","The  Login desc exists"
			Dim t1
			t1=Browser("Titan: The Official Website").Page("Titan: The Official Website").WebElement("login_desc").GetROProperty("outertext")
			Reporter.ReportEvent micPass,"Login description","The  Login desc says:"&t1
		Else
			Reporter.ReportEvent micFail,"Login description didn't work","The Login description didn't work"
		End If
	Else
		Reporter.ReportEvent micFail,"Account image didn't work","The  account image didn't work"
	End If
	'Country code'
	If Browser("Titan: The Official Website").Page("Titan: The Official Website").WebElement("+91").Exist(5) Then
		Reporter.ReportEvent micPass,"Country code element exist","The country code element exists and working"
		Browser("Titan: The Official Website").Page("Titan: The Official Website").WebElement("+91").Highlight
		Reporter.ReportEvent micPass,"Country code element Highlighted","Country code element got Highlighted"
		Browser("Titan: The Official Website").Page("Titan: The Official Website").WebElement("+91").Click
		Reporter.ReportEvent micPass,"Country code element Clicked","Country code element got Clicked"
		Dim cl
		cl=Browser("Titan: The Official Website").Page("Titan: The Official Website").WebElement("CountryList").GetROProperty("innertext")
		Reporter.ReportEvent micPass,"All items of country list","Items in the country list:"&cl
		Browser("Titan: The Official Website").Page("Titan: The Official Website").WebElement("CountryList").Click(" +61 Australia")
		Reporter.ReportEvent micPass,"Clicked ","Australia selected got Clicked from the list"
	Else
		Reporter.ReportEvent micFail,"Country code element didn't work","Country code element didn't work"
	End If
	'Mobile number section'
	If Browser("Titan: The Official Website").Page("Titan: The Official Website").WebEdit("dwfrm_profile_phone").Exist(5) Then
		Reporter.ReportEvent micPass,"Mobile text box exist","The Mobile text box exists and working"
		Browser("Titan: The Official Website").Page("Titan: The Official Website").WebEdit("dwfrm_profile_phone").Highlight
		Reporter.ReportEvent micPass,"Mobile text box highlighted","The Mobile text box got highlighted"
		'Fail case
		     Browser("Titan: The Official Website").Page("Titan: The Official Website").WebEdit("dwfrm_profile_phone").Set("a") 
			Reporter.ReportEvent micPass,"Mobile number set to a","Alphabet set successfully"
			Dim e1,e2
			e1=Browser("Titan: The Official Website").Page("Titan: The Official Website").WebElement("alph_error").GetROProperty("innertext")
			Reporter.ReportEvent micFail,"Mobile number set to a","The error message is:"&e1
			Browser("Titan: The Official Website").Page("Titan: The Official Website").WebEdit("dwfrm_profile_phone").Set("98")
			e2=Browser("Titan: The Official Website").Page("Titan: The Official Website").WebElement("valid-error").GetROProperty("innertext")
			Reporter.ReportEvent micFail,"Mobile number set to a","The error message is:"&e2
		'Pass case'
			Browser("Titan: The Official Website").Page("Titan: The Official Website").WebEdit("dwfrm_profile_phone").Set("9894654###") 
			Reporter.ReportEvent micPass,"Mobile number is set ","Correct number is set"
		Else
			Reporter.ReportEvent micFail,"Mobile number error","The mobile number can't be set"
		End If
	'Terms of service'
	If Browser("Titan: The Official Website").Page("Titan: The Official Website").Link("Terms Of Service").Exist(2) Then
		Reporter.ReportEvent micPass,"Term of service link exist","The terms of service exists and working"
		Browser("Titan: The Official Website").Page("Titan: The Official Website").Link("Terms Of Service").Highlight
		Reporter.ReportEvent micPass,"Term of service link Highlighted","The terms of service link is highlighted"
		Browser("Titan: The Official Website").Page("Titan: The Official Website").Link("Terms Of Service").Click
		Reporter.ReportEvent micPass,"Term of service link Clicked","The terms of service link is Clicked"
	Else
		Reporter.ReportEvent micFail,"Terms of service error","The terms of service didn't open"
	End If
	'Privacy policies'
	If Browser("Titan: The Official Website").Page("Titan: The Official Website").Link("Privacy Policy").Exist(2) Then
		Reporter.ReportEvent micPass,"Privacy Policy link exist","The Privacy Policy exists and working"
		Browser("Titan: The Official Website").Page("Titan: The Official Website").Link("Privacy Policy").Highlight
		Reporter.ReportEvent micPass,"Term of service link Highlighted","The terms of service link is highlighted"
		Browser("Titan: The Official Website").Page("Titan: The Official Website").Link("Privacy Policy").Click
		Reporter.ReportEvent micPass,"Privacy Policy link Clicked","The Privacy Policy link is Clicked"
	Else
		Reporter.ReportEvent micFail,"Privacy Policy error","The Privacy Policy didn't open"
	End If
	'Continue button section'
	If Browser("Titan: The Official Website").Page("Titan: The Official Website").WebButton("Continue").Exist(5) Then
		Reporter.ReportEvent micPass,"Continue button exist","The Continue button exists and working"
		Browser("Titan: The Official Website").Page("Titan: The Official Website").WebButton("Continue").Highlight
		Reporter.ReportEvent micPass,"Continue button highlighted","The Continue button got highlighted"
		Browser("Titan: The Official Website").Page("Titan: The Official Website").WebButton("Continue").Click
		Reporter.ReportEvent micPass,"Continue button Clicked","The Continue button got Clicked"
	Else
		Reporter.ReportEvent micFail,"Continue button didn't work","The Continue button didn't work"
	End If
	'OTP Section'
	If Browser("Titan: The Official Website").Page("Titan: The Official Website").WebElement("OtpDesc").Exist(5) Then
		Reporter.ReportEvent micPass,"Otp page exist","The OTP Page exists and working"
		Browser("Titan: The Official Website").Page("Titan: The Official Website").WebElement("OtpDesc").Highlight
		Reporter.ReportEvent micPass,"Otp page Highlighted","The OTP Page got highlighted"
		Dim p1
		p1=Browser("Titan: The Official Website").Page("Titan: The Official Website").WebElement("OtpDesc").GetROProperty("innertext")
		Reporter.ReportEvent micPass,"Enter otp page description:","Enter OTP page description:"&p1
		'OTP Text box section'
		If Browser("Titan: The Official Website").Page("Titan: The Official Website").WebElement("textboxelement").Exist(5) Then
			Reporter.ReportEvent micPass,"OTP Text box exist","The OTP Text box exists and working"
			Browser("Titan: The Official Website").Page("Titan: The Official Website").WebElement("textboxelement").Highlight
			Reporter.ReportEvent micPass,"OTP Text box highlighted","The OTP Text box got highlighted"
		Else
			Reporter.ReportEvent micFail,"OTP Text box Error","OTP Text box didn't work"
		End If
		'Back button'
		If  Browser("Titan: The Official Website").Page("Titan: The Official Website").WebButton("Back").Exist(5) Then
			Reporter.ReportEvent micPass,"Back button exist","The Back button exists and working"
			Browser("Titan: The Official Website").Page("Titan: The Official Website").WebButton("Back").Highlight
			Reporter.ReportEvent micPass,"Back button highlighted","The Back button got highlighted"
			Browser("Titan: The Official Website").Page("Titan: The Official Website").WebButton("Back").Click
			Reporter.ReportEvent micPass,"Back button is clicked","The Back button got clicked successfully"
		Else
			Reporter.ReportEvent micFail,"Back button didn't work","Back button didn't work"
		End If
	Else
		Reporter.ReportEvent micFail,"Enter OTP page didn't work","Enter OTP page didn't work"	
	End If
	'Login close button'
	If Browser("Titan: The Official Website").Page("Titan: The Official Website").WebButton("LoginClose").Exist(5) Then
		Reporter.ReportEvent micPass,"Login close button exist","Login close button exists and working"
		Browser("Titan: The Official Website").Page("Titan: The Official Website").WebButton("LoginClose").Highlight
		Reporter.ReportEvent micPass,"Login close button highlighted","Login close button got highlighted"
		Browser("Titan: The Official Website").Page("Titan: The Official Website").WebButton("LoginClose").Click
		Reporter.ReportEvent micPass,"Login close button clicked","Login close button got clicked"
	Else
		Reporter.ReportEvent micFail,"Login close button work","The Login close button didn't work"
	End If
	'Search box section'
	If Browser("Titan: The Official Website").Page("Titan: The Official Website").WebEdit("Enter Keyword or Item").Exist(5)  Then
		Reporter.ReportEvent micPass,"Search box exist","The search box exists and working"
		Browser("Titan: The Official Website").Page("Titan: The Official Website").WebEdit("Enter Keyword or Item").Highlight
		Reporter.ReportEvent micPass,"Search box Highlighted","Search box got Highlighted"
		Browser("Titan: The Official Website").Page("Titan: The Official Website").WebEdit("Enter Keyword or Item").Set("Watches")
		Reporter.ReportEvent micPass,"Search box text set","Search box text is fixed"
		Set devRep = CreateObject("Mercury.DeviceReplay")
		devRep.PressKey 28
		Reporter.ReportEvent micPass,"Enter button pressed","Enter button pressed"
	Else
		Reporter.ReportEvent micFail,"Search box didn't work","The search box didn't work"
	End If
	
End Function
'Coupon check function'
Function Mod2()
	If Browser("Titan: The Official Website").Page("Titan: The Official Website").WebElement("Coupon").Exist(5) Then
		Reporter.ReportEvent micPass,"Coupon element exist","The coupon element exists and working"
		Browser("Titan: The Official Website").Page("Titan: The Official Website").WebElement("Coupon").Highlight
		Reporter.ReportEvent micPass,"Coupon element highlighted","The coupon element got highlighted"
		Dim c1
		c1=Browser("Titan: The Official Website").Page("Titan: The Official Website").WebElement("Coupon").GetROProperty("innertext")
		Reporter.ReportEvent micPass,"Coupon element fetched","The coupon element says:"&c1
	Else	
		Reporter.ReportEvent micFail,"Coupon element didn't work","The coupon element didn't work"
	End If
	'Popular search section'
	If Browser("Titan: The Official Website").Page("Titan: The Official Website").WebList("search-results").Exist(5) Then
		Reporter.ReportEvent micPass,"Popular Search element exist","The Popular search element exists and working"
		Browser("Titan: The Official Website").Page("Titan: The Official Website").WebList("search-results").Highlight
		Reporter.ReportEvent micPass,"Popular Search element highlighted","The Popular search element got highlighted"
		Dim c2
		c2=Browser("Titan: The Official Website").Page("Titan: The Official Website").WebList("search-results").GetROProperty("all items")
		Reporter.ReportEvent micPass,"Popular Search description","The Popular search element says:"&c2
	Else
		Reporter.ReportEvent micFail,"Popular search element error","The popular search element didn't work"
	End If
	'Trending search results'
	If Browser("Titan: The Official Website").Page("Titan: The Official Website").WebList("trending").Exist(5) Then
		Reporter.ReportEvent micPass,"Trending element exists","The trending element exists and working"
		Browser("Titan: The Official Website").Page("Titan: The Official Website").WebList("trending").Highlight
		Reporter.ReportEvent micPass,"Trending element highlighted","The Trending element got highlighted"
		Dim c3,c4
		c3=Browser("Titan: The Official Website").Page("Titan: The Official Website").WebList("trending").GetROProperty("all items")
		Reporter.ReportEvent micPass,"Trending element desc","The Trending element says: "&c3
		c4=Browser("Titan: The Official Website").Page("Titan: The Official Website").WebList("trending").GetROProperty("items count")
		Reporter.ReportEvent micPass,"Trending element count","The Trending count is : "&c4
	Else
		Reporter.ReportEvent micFail,"Trending element error","The Trending element didn't work"
	End If
	'Record button'
	If Browser("Titan: The Official Website").Page("Titan: The Official Website").Image("record").Exist(5) Then
		Reporter.ReportEvent micPass,"Record element exists","The Record element exists and working"
		Browser("Titan: The Official Website").Page("Titan: The Official Website").Image("record").Highlight
		Reporter.ReportEvent micPass,"Record element Highlighted","The Record element got highlighted"
		Browser("Titan: The Official Website").Page("Titan: The Official Website").Image("record").Click
		Reporter.ReportEvent micPass,"Record element Clicked","The Record element got clicked"
	Else
		Reporter.ReportEvent micFail,"Record element error","The Record element didn't work"	
	End If
	'Close search button'
	If Browser("Titan: The Official Website").Page("Titan: The Official Website").Image("clear search input").Exist(5) Then
		Reporter.ReportEvent micPass,"Search close element exists","The Search close element exists and working"
		 Browser("Titan: The Official Website").Page("Titan: The Official Website").Image("clear search input").Highlight
		 Reporter.ReportEvent micPass,"Search close element Highlighted","The Search close element got highlighted"
		Browser("Titan: The Official Website").Page("Titan: The Official Website").Image("clear search input").Click
		 Reporter.ReportEvent micPass,"Search close Clicked","The Search close element got clicked"	
	Else
		Reporter.ReportEvent micFail,"Search close element error","The Search element didn't work"	
	End If
End Function
'Wishlist function'
Function Mod3()
	If Browser("Titan: The Official Website").Page("Titan: The Official Website").Image("Wishlist").Exist(5) Then
		Reporter.ReportEvent micPass,"Wishlist element exists","Wishlist element exists and working"
		Browser("Titan: The Official Website").Page("Titan: The Official Website").Image("Wishlist").Highlight
		Reporter.ReportEvent micPass,"Wishlist element highlighted","Wishlist element got highlighted"
		Browser("Titan: The Official Website").Page("Titan: The Official Website").Image("Wishlist").Click
		Reporter.ReportEvent micPass,"Wishlist element clicked","Wishlist element got clicked"
	Else
		Reporter.ReportEvent micFail,"Wishlist element error","The Wishlist element didn't work"	
	End If
	'E-Gift card
	If Browser("Titan: The Official Website").Page("Titan: The Official Website").Link("eGift Card").Exist(5) Then
		Reporter.ReportEvent micPass,"E-Gift card element exists","E-Gift card element exists and working"
		Browser("Titan: The Official Website").Page("Titan: The Official Website").Link("eGift Card").Highlight
		Reporter.ReportEvent micPass,"E-Gift card element highlighted","E-Gift card element got highlighted"
		Browser("Titan: The Official Website").Page("Titan: The Official Website").Link("eGift Card").Click
		Reporter.ReportEvent micPass,"E-Gift card element clicked","E-Gift card element got clicked"
	Else
		Reporter.ReportEvent micFail,"E-Gift card element error","E-Gift card element didn't work"
	End If
	'Find a store button'
	If Browser("Titan: The Official Website").Page("Titan: The Official Website").Link("Find A Store").Exist(5)  Then
		Reporter.ReportEvent micPass,"Find a store element exists","Find a store element exists and working"
		 Browser("Titan: The Official Website").Page("Titan: The Official Website").Link("Find A Store").Highlight
		 Reporter.ReportEvent micPass,"Find a store element highlighted","Find a store element got highlighted"
		 Browser("Titan: The Official Website").Page("Titan: The Official Website").Link("Find A Store").Click
		 Reporter.ReportEvent micPass,"Find a store element Clicked","Find a store element got clicked"
	Else
		Reporter.ReportEvent micFail,"Find a store element error","Find a store element didn't work"
	End If
	'Help and contact button'
	If Browser("Titan: The Official Website").Page("Titan: The Official Website").Link("Help & Contact").Exist(5) Then
		Reporter.ReportEvent micPass,"Help and contact element exists","Help and contact element exists and working"
		Browser("Titan: The Official Website").Page("Titan: The Official Website").Link("Help & Contact").Highlight
		Reporter.ReportEvent micPass,"Help and contact element highlighted","Help and contact element got highlighted"
		Browser("Titan: The Official Website").Page("Titan: The Official Website").Link("Help & Contact").Click
		Reporter.ReportEvent micPass,"Help and contact element clicked","Help and contact element got clicked"
	Else
		Reporter.ReportEvent micFail,"Help and contact element error","Help and contact element didn't work"
	End If
	'FAQ'
	If Browser("Titan: The Official Website").Page("Titan: The Official Website").Link("FAQ").Exist(5) Then
		Reporter.ReportEvent micPass,"FAQ button exists","FAQ Button exists and working"
		 Browser("Titan: The Official Website").Page("Titan: The Official Website").Link("FAQ").Highlight
		 Reporter.ReportEvent micPass,"FAQ button highlighted","FAQ Button got highlighted"
		  Browser("Titan: The Official Website").Page("Titan: The Official Website").Link("FAQ").Click
		  Reporter.ReportEvent micPass,"FAQ button clicked","FAQ Button got clicked"
	Else
		Reporter.ReportEvent micFail,"FAQ element error","FAQ element didn't work"
	End If
	'Cart button'
	If Browser("Titan: The Official Website").Page("Titan: The Official Website").Image("cart").Exist(5) Then
		Reporter.ReportEvent micPass,"Cart button exists","Cart image exists and working"
		 Browser("Titan: The Official Website").Page("Titan: The Official Website").Image("cart").Highlight
		 Reporter.ReportEvent micPass,"Cart button highlighted","Cart image button got highlighted"
		 Browser("Titan: The Official Website").Page("Titan: The Official Website").Image("cart").Click
		  Reporter.ReportEvent micPass,"Cart button clicked","Cart image button got clicked"
	Else
		Reporter.ReportEvent micFail,"Cart button error","Cart button didn't work"
	End If
	'Track order button'
	If Browser("Titan: The Official Website").Page("Titan: The Official Website").Link("Track Order").Exist(5) Then
		Reporter.ReportEvent micPass,"Track order button exists","Track order button exists and working"
		Browser("Titan: The Official Website").Page("Titan: The Official Website").Link("Track Order").Highlight
		Reporter.ReportEvent micPass,"Track order button highlighted","Track order button highlighted"
		Browser("Titan: The Official Website").Page("Titan: The Official Website").Link("Track Order").Click
		Reporter.ReportEvent micPass,"Track order button clicked","Track order button got clicked"
		'Details inside track order'
		If Browser("Titan: The Official Website").Page("Titan: The Official Website").WebElement("TrackOrderDetails").Exist(5) Then
			Reporter.ReportEvent micPass,"Track order details exists","Track order details exists and working"
			Browser("Titan: The Official Website").Page("Titan: The Official Website").WebElement("TrackOrderDetails").Highlight
			Reporter.ReportEvent micPass,"Track order details highlighted","Track order details got highlighted"
			Dim t1
			t1=Browser("Titan: The Official Website").Page("Titan: The Official Website").WebElement("TrackOrderDetails").GetROProperty("innertext")
			Reporter.ReportEvent micPass,"Track order details","Track order details says:"&t1
		Else
			Reporter.ReportEvent micFail,"Track order details error","Track order details Cart button didn't work"
		End If
		'Email-id section'
		If Browser("Titan: The Official Website").Page("Titan: The Official Website").WebEdit("emailInput").Exist(5) Then
			Reporter.ReportEvent micPass,"Email text box exists","Email text box exists and working"
			Browser("Titan: The Official Website").Page("Titan: The Official Website").WebEdit("emailInput").Highlight
			Reporter.ReportEvent micPass,"Email text box highlighted","Email text box got highlighted"
			Browser("Titan: The Official Website").Page("Titan: The Official Website").WebEdit("emailInput").Set("hari")
			Set devRep1 = CreateObject("Mercury.DeviceReplay")
			devRep1.PressKey 28
			Reporter.ReportEvent micPass,"Enter button pressed","Enter button pressed"
			Dim em1
			em1=Browser("Titan: The Official Website").Page("Titan: The Official Website").WebElement("email-error").GetROProperty("innertext")
			Reporter.ReportEvent micFail,"Fail case of email ","Error on setting 'hari' says:"&em1
			Browser("Titan: The Official Website").Page("Titan: The Official Website").WebEdit("emailInput").Set("hariharasudhan.m.2026@gmail.com")
		Else
			Reporter.ReportEvent micFail,"Email text box error ","Email text box error occured"
		End If
		'Track order number field'
		If Browser("Titan: The Official Website").Page("Titan: The Official Website").WebEdit("orderno").Exist(5) Then
			Reporter.ReportEvent micPass,"Order number box exists","Order number box exists and working"
			Browser("Titan: The Official Website").Page("Titan: The Official Website").WebEdit("orderno").Highlight
			Reporter.ReportEvent micPass,"Order number box highlighted","Order number box got highlighted"
			Browser("Titan: The Official Website").Page("Titan: The Official Website").WebEdit("orderno").Set("20044002")
			Reporter.ReportEvent micPass,"Order number is set","Order number is set"
		Else
			Reporter.ReportEvent micFail,"Order box error","Order box didn't work"
		End If
	Else
		Reporter.ReportEvent micFail,"Track order button error","Track order button didn't work"
	End If
End Function
'HOME Page -Account
Function Mod4()
'Login account logo'
	If Browser("Titan: The Official Website").Page("Titan: The Official Website").WebElement("AccLogo").Exist(5) Then
		Reporter.ReportEvent micPass,"Login account logo exists","Login account logo exists and working"
		Browser("Titan: The Official Website").Page("Titan: The Official Website").WebElement("AccLogo").Highlight
		Reporter.ReportEvent micPass,"Login account logo highlighted","Login account logo got highlighted"
		Browser("Titan: The Official Website").Page("Titan: The Official Website").WebElement("AccLogo").HoverTap
		Reporter.ReportEvent micPass,"Login account logo hovered","Login account logo got hovered"
		Dim l1
		l1=Browser("Titan: The Official Website").Page("Titan: The Official Website").WebElement("AccLogo").GetROProperty("innertext")
		Reporter.ReportEvent micPass,"Login account logo contents","Login account logo says:"&l1
		Browser("Titan: The Official Website").Page("Titan: The Official Website").WebElement("AccLogo").Click
		Reporter.ReportEvent micPass,"Login account logo clicked","Login account logo got clicked"
	Else
		Reporter.ReportEvent micFail,"Login account logo error","Login account logo didn't work"
	End If
	'Account overview text'
	If Browser("Titan: The Official Website").Page("MyAccount").WebElement("Account Overview").Exist(5) Then
		Reporter.ReportEvent micPass,"Account overview text exist","Account overview exists and working"
		Browser("Titan: The Official Website").Page("MyAccount").WebElement("Account Overview").Highlight
		Reporter.ReportEvent micPass,"Account overview text highlighted","Account overview got highlighted"
		Dim l2
		l2=Browser("Titan: The Official Website").Page("MyAccount").WebElement("Account Overview").GetROProperty("innertext")
		Reporter.ReportEvent micPass,"Account overview text fetched","It says:"&l2
	Else
		Reporter.ReportEvent micFail,"Account overview text error","Account overview text didn't work"
	End If
	'Personal info text'
	If Browser("Titan: The Official Website").Page("MyAccount").WebElement("Personal Information").Exist(5) Then
		Reporter.ReportEvent micPass,"Personal Information text exist","Personal Information exists and working"
		Browser("Titan: The Official Website").Page("MyAccount").WebElement("Personal Information").Highlight
		Reporter.ReportEvent micPass,"Personal Information highlighted","Personal Information got highlighted"
		Dim l3
		l3=Browser("Titan: The Official Website").Page("MyAccount").WebElement("Personal Information").GetROProperty("innertext")
		Reporter.ReportEvent micPass,"Personal Information text fetched","It says:"&l3
	Else
		Reporter.ReportEvent micFail,"Personal Information text error","Personal Information text didn't work"
	End If
	'Personal information link'
	If Browser("Titan: The Official Website").Page("MyAccount").Link("Personal Information").Exist(5) Then
		Reporter.ReportEvent micPass,"Personal Information text exist","Personal Information exists and working"
		Browser("Titan: The Official Website").Page("MyAccount").Link("Personal Information").Highlight
		Reporter.ReportEvent micPass,"Personal Information highlighted","Personal Information got highlighted"
		Browser("Titan: The Official Website").Page("MyAccount").Link("Personal Information").Click
		Reporter.ReportEvent micPass,"Personal Information link clicked","Personal info got clicked"
	Else
		Reporter.ReportEvent micFail,"Personal Information text error","Personal Information text didn't work"
	End If
	'Personal information list'
	If Browser("Titan: The Official Website").Page("MyAccount").WebElement("Personal_info_element").Exist(5) Then
		Reporter.ReportEvent micPass,"Personal_info_element exist","Personal_info_element exists and working"	
		Browser("Titan: The Official Website").Page("MyAccount").WebElement("Title : Mr.").Highlight
		Reporter.ReportEvent micPass,"Title list highlighted","Title list exists got highlighted"
		Browser("Titan: The Official Website").Page("MyAccount").WebElement("First Name : Hariharasudhan").Highlight
		Reporter.ReportEvent micPass,"First name highlighted","First name got highlighted"
		Browser("Titan: The Official Website").Page("MyAccount").WebElement("Date of birth : -").Highlight
		Reporter.ReportEvent micPass,"Date of birth highlighted","Date of birth got highlighted"
		Browser("Titan: The Official Website").Page("MyAccount").WebElement("Anniversary : -").Highlight
		Reporter.ReportEvent micPass,"Anniversary highlighted","Anniversary got highlighted"
		Browser("Titan: The Official Website").Page("MyAccount").WebElement("Encircle ID : XXXXXXXX2591").Highlight
		Reporter.ReportEvent micPass,"Encircle ID highlighted","Encircle ID got highlighted"
		Browser("Titan: The Official Website").Page("MyAccount").WebElement("NeuCoins : 0 NeuCoins").Highlight
		Reporter.ReportEvent micPass,"NeuCoins highlighted","NeuCoins got highlighted"
	Else
		Reporter.ReportEvent micFail,"Personal Information list error","Personal Information list didn't work"	
	End If
	'Edit personal info
	If Browser("Titan: The Official Website").Page("MyAccount").WebButton("Edit").Exist(5) Then
		Reporter.ReportEvent micPass,"Edit button exists","Edit button exists and working"
		Browser("Titan: The Official Website").Page("MyAccount").WebButton("Edit").Highlight
		Reporter.ReportEvent micPass,"Edit button highlighted","Edit button got highlighted"
		Browser("Titan: The Official Website").Page("MyAccount").WebButton("Edit").Click
		Reporter.ReportEvent micPass,"Edit button clicked","Edit button got clicked"
		If Browser("Titan: The Official Website").Page("MyAccount").WebCheckBox("checkOffer").Exist(5) Then
			Reporter.ReportEvent micPass,"Check box exists","Check box exists and working"
			Browser("Titan: The Official Website").Page("MyAccount").WebCheckBox("checkOffer").Highlight
			Reporter.ReportEvent micPass,"Check box highlighted","Edit button got highlighted"
			Browser("Titan: The Official Website").Page("MyAccount").WebCheckBox("checkOffer").Set("ON")
			Reporter.ReportEvent micPass,"Check box selected","Check box is selected"
		Else
			Reporter.ReportEvent micFail,"Check box error","Check box didn't work"
	End If
	'Cancel button
	If Browser("Titan: The Official Website").Page("MyAccount").WebButton("Cancel").Exist(5) Then
		Reporter.ReportEvent micPass,"Cancel button exists","Cancel button exists and working"
		Browser("Titan: The Official Website").Page("MyAccount").WebButton("Cancel").Highlight
		Reporter.ReportEvent micPass,"Cancel button highlighted","Cancel button got highlighted"
		Browser("Titan: The Official Website").Page("MyAccount").WebButton("Cancel").Click
	Else
		Reporter.ReportEvent micFail,"Cancel button error","Cancel button didn't work"
	End If
	'Save button
	If Browser("Titan: The Official Website").Page("MyAccount").WebButton("Save").Exist(5) Then
		Reporter.ReportEvent micPass,"Save button exists","Save button exists and working"
		Browser("Titan: The Official Website").Page("MyAccount").WebButton("Save").Highlight
		Reporter.ReportEvent micPass,"Save button highlighted","Save button got highlighted"
		Browser("Titan: The Official Website").Page("MyAccount").WebButton("Save").Click
		Reporter.ReportEvent micPass,"Save button clicked","Save button got clicked"
	Else
		Reporter.ReportEvent micFail,"Save button error","Save button didn't work"
	End If
	Else
		Reporter.ReportEvent micFail,"Edit button error","Edit button didn't work"
	End If
	'Address book section
	If Browser("Titan: The Official Website").Page("MyAccount").Link("Address Book").Exist(5) Then
		Reporter.ReportEvent micPass,"Address book button exists","Address book button exists and working"
		Browser("Titan: The Official Website").Page("MyAccount").Link("Address Book").Highlight
		Reporter.ReportEvent micPass,"Address book button highlighted","Address book button got highlighted"
		Browser("Titan: The Official Website").Page("MyAccount").Link("Address Book").Click
		Reporter.ReportEvent micPass,"Address book button clicked","Address book button got clicked"
	Else
		Reporter.ReportEvent micFail,"Address book button error","Address book button didn't work"	
	End If
	'Address edit
	If Browser("Titan: The Official Website").Page("MyAccount").WebButton("AddEdit").Exist(5) Then
		Reporter.ReportEvent micPass,"Address book edit button lists","Address book edit button exists and working"
		Browser("Titan: The Official Website").Page("MyAccount").WebButton("AddEdit").Highlight
		Reporter.ReportEvent micPass,"Address edit button highlighted","Address book edit button got highlighted"
		Browser("Titan: The Official Website").Page("MyAccount").WebButton("AddEdit").Click
		Reporter.ReportEvent micPass,"Address edit button clicked","Address book edit button got clicked"
	Else
		Reporter.ReportEvent micFail,"Address book edit error","Address book edit button didn't work"
	End If
	'Edit address section
	If Browser("Titan: The Official Website").Page("Sites-Titan-Site").WebElement("Title Mr. Ms. Code* Mr.").Exist(5) Then
		Reporter.ReportEvent micPass,"Title button exists","Title button exists and working"
		Browser("Titan: The Official Website").Page("Sites-Titan-Site").WebElement("Title Mr. Ms. Code* Mr.").Highlight
		Reporter.ReportEvent micPass,"Title button highlighted","Title button got highlighted"
		Browser("Titan: The Official Website").Page("Sites-Titan-Site").WebElement("Mr.").Click
		Reporter.ReportEvent micPass,"Title Mr button is clicked","Title Mr button got clicked"
		'Contact info
		If Browser("Titan: The Official Website").Page("Sites-Titan-Site").WebEdit("contact").Exist(5) Then
			Reporter.ReportEvent micPass,"Contact info exists","Contact info exists and working"
			Browser("Titan: The Official Website").Page("Sites-Titan-Site").WebEdit("contact").Highlight
			Reporter.ReportEvent micPass,"Contact info highlighted","Contact info got highlighted"
			Dim l4
			l4=Browser("Titan: The Official Website").Page("Sites-Titan-Site").WebEdit("contact").GetROProperty("innertext")
			Reporter.ReportEvent micPass,"Contact info fetched","Contact info says:"&l4
		Else
			Reporter.ReportEvent micFail,"Contact info error","Contact info button didn't work"
		End If
		'Full name section
		If Browser("Titan: The Official Website").Page("Sites-Titan-Site").WebEdit("address_fullName").Exist(5) Then
			Reporter.ReportEvent micPass,"Full name text box exists","Full name text box exists and working"
			Browser("Titan: The Official Website").Page("Sites-Titan-Site").WebEdit("address_fullName").Highlight
			Reporter.ReportEvent micPass,"Full name text box highlighted","Full name text box got highlighted"
			Dim l5
			l5=Browser("Titan: The Official Website").Page("Sites-Titan-Site").WebEdit("address_fullName").GetROProperty("innertext")
			Reporter.ReportEvent micPass,"Full name text box entered","Full name text box says:"&l5
		Else
			Reporter.ReportEvent micFail,"Full name text box error","Full name text box didn't work"
		End If
		'Email address field
		If Browser("Titan: The Official Website").Page("Sites-Titan-Site").WebEdit("edit_address_email").Exist(5) Then
			Reporter.ReportEvent micPass,"Full name text box exists","Full name text box exists and working"
			Browser("Titan: The Official Website").Page("Sites-Titan-Site").WebEdit("edit_address_email").Highlight
			Reporter.ReportEvent micPass,"Full name text box highlighted","Full name text box got highlighted"
			Dim l6
			l6=Browser("Titan: The Official Website").Page("Sites-Titan-Site").WebElement("EmailAddressText").GetROProperty("innertext")
			Reporter.ReportEvent micPass,"Email address text is available:","Email address text says:"&l6
		Else
			Reporter.ReportEvent micFail,"Email address text box error","Email address text box didn't work"
		End If 
		'Landmark box
		If Browser("Titan: The Official Website").Page("Sites-Titan-Site").WebEdit("address_landmark").Exist(5) Then
			Reporter.ReportEvent micPass,"Landmark text box exists","Landmark text box exists and working"
			Browser("Titan: The Official Website").Page("Sites-Titan-Site").WebEdit("address_landmark").Highlight
			Reporter.ReportEvent micPass,"Landmark text box exists","Landmark text box exists and working"
		Else
			Reporter.ReportEvent micFail,"Land mark text box error","Land mark text box didn't work"
		End If
		'Postal code
		If Browser("Titan: The Official Website").Page("Sites-Titan-Site").WebEdit("dwfrm_address_postalCode").Exist(5) Then
			Reporter.ReportEvent micPass,"Postal code box exists","Postal code box exists and working"
			Browser("Titan: The Official Website").Page("Sites-Titan-Site").WebEdit("dwfrm_address_postalCode").Highlight
			Reporter.ReportEvent micPass,"Postal code box highlighted","Postal code got highlighted"
			Browser("Titan: The Official Website").Page("Sites-Titan-Site").WebEdit("dwfrm_address_postalCode").Set("635109")
			Reporter.ReportEvent micPass,"Postal code is set","Postal code is set"
		Else
			Reporter.ReportEvent micFail,"Postal code text box error","Postal code text box didn't work"
		End If
		'City text box
		If Browser("Titan: The Official Website").Page("Sites-Titan-Site").WebEdit("address_city").Exist(5) Then
			Reporter.ReportEvent micPass,"City text box exists","City text box exists and working"
			Browser("Titan: The Official Website").Page("Sites-Titan-Site").WebEdit("address_city").Highlight
			Reporter.ReportEvent micPass,"City text box highlighted","City text box highlighted"
		Else
			Reporter.ReportEvent micFail,"City text box error","City text box didn't work"
		End If
		'State text box
		If Browser("Titan: The Official Website").Page("Sites-Titan-Site").WebEdit("stateCode").Exist(5) Then
			Reporter.ReportEvent micPass,"State code box exists","State code boxexists and working"
			Browser("Titan: The Official Website").Page("Sites-Titan-Site").WebEdit("stateCode").Highlight
			Reporter.ReportEvent micPass,"State code box highlighted","State code box got highlighted"
		Else
			Reporter.ReportEvent micFail,"State code box error","State code box didn't work properly"
		End If
		'Address type element
		If Browser("Titan: The Official Website").Page("Sites-Titan-Site").WebElement("AddressType").Exist(5) Then
			Reporter.ReportEvent micPass,"Address type element exists","Address type element exists and working"
			Browser("Titan: The Official Website").Page("Sites-Titan-Site").WebElement("AddressType").Highlight
			Reporter.ReportEvent micPass,"Address type element highlighted","Address type element got highlighted"
			Dim a1
			a1=Browser("Titan: The Official Website").Page("Sites-Titan-Site").WebElement("AddressType").GetROProperty("innertext")
			Reporter.ReportEvent micPass,"Address type fetched","Address type element are:"&a1
		Else
			Reporter.ReportEvent micFail,"Address type failed","Address type didn't work properly"
		End If
		'Office button
		If Browser("Titan: The Official Website").Page("Sites-Titan-Site").WebElement("Office").Exist(5) Then
			Reporter.ReportEvent micPass,"Office element exists","Office element exists and working"
			Browser("Titan: The Official Website").Page("Sites-Titan-Site").WebElement("Office").Highlight
			Reporter.ReportEvent micPass,"Office element highlighted","Office element got highlighted"
			Browser("Titan: The Official Website").Page("Sites-Titan-Site").WebElement("Office").Click
			Reporter.ReportEvent micPass,"Office element clicked","Office element got clicked"
		Else
			Reporter.ReportEvent micFail,"Office element failed","Office element didn't work properly"
		End If
		'Other element
		If Browser("Titan: The Official Website").Page("Sites-Titan-Site").WebElement("Other").Exist(5) Then
			Reporter.ReportEvent micPass,"Other element exists","Other element exists and working"
			Browser("Titan: The Official Website").Page("Sites-Titan-Site").WebElement("Other").Highlight
			Reporter.ReportEvent micPass,"Other element highlighted","Other element got highlighted"
			Browser("Titan: The Official Website").Page("Sites-Titan-Site").WebElement("Office").Click
			Reporter.ReportEvent micPass,"Other element clicked","Other element got clicked"
		End If
		'Cancel button
		If Browser("Titan: The Official Website").Page("Sites-Titan-Site").Link("AddrCancel").Exist(5) Then
			Reporter.ReportEvent micPass,"Cancel button exists","Cancel button exists and working"
			Browser("Titan: The Official Website").Page("Sites-Titan-Site").Link("AddrCancel").Highlight
			Reporter.ReportEvent micPass,"Cancel button highlighted","Cancel button got highlighted" 
			Browser("Titan: The Official Website").Page("Sites-Titan-Site").Link("AddrCancel").Click
			Reporter.ReportEvent micPass,"Cancel button clicked","Cancel button got clicked" 
		Else
			Reporter.ReportEvent micFail,"Cancel button error","Cancel button didn't work"
		End If
		'Add new address button
		If Browser("Titan: The Official Website").Page("MyAccount").WebButton("Add New Address").Exist(5) Then
			Reporter.ReportEvent micPass,"Add new button exists","Add new button exists and working"
			Browser("Titan: The Official Website").Page("MyAccount").WebButton("Add New Address").Highlight
			Reporter.ReportEvent micPass,"Add new button exists","Add new button exists and working"
		Else
			Reporter.ReportEvent micFail,"Add new button error","Add new button gives error"
		End If
	Else
		Reporter.ReportEvent micFail,"Edit Address error","Edit Address button didn't work"
	End If
	'Wishlist button
	If Browser("Titan: The Official Website").Page("MyAccount").Link("Wishlist").Exist(5) Then
		Reporter.ReportEvent micPass,"Wishlist button exists","Wishlist button exists and working"	
		Browser("Titan: The Official Website").Page("MyAccount").Link("Wishlist").Highlight
		Reporter.ReportEvent micPass,"Wishlist button highlighted","Wishlist button got highlighted"
		Browser("Titan: The Official Website").Page("MyAccount").Link("Wishlist").Click
		Reporter.ReportEvent micPass,"Wishlist button clicked","Wishlist button got clicked"
	Else
		Reporter.ReportEvent micFail,"Wishlist error","Wishlist didn't work"
	End If
	'Order history button
	If Browser("Titan: The Official Website").Page("MyAccount").Link("Order History").Exist(5) Then
		Reporter.ReportEvent micPass,"Order history button exists","Order history button exists and working"
		Browser("Titan: The Official Website").Page("MyAccount").Link("Order History").Highlight
		Reporter.ReportEvent micPass,"Order history button highlighted","Order history button got highlighted"
		Browser("Titan: The Official Website").Page("MyAccount").Link("Order History").Click
		Reporter.ReportEvent micPass,"Order history button clicked","Order history button got clicked"
		If Browser("Titan: The Official Website").Page("MyAccount").WebElement("no order details yet !").Exist(5) Then
			Reporter.ReportEvent micPass,"No order text exists","No order text exists and working"
			Browser("Titan: The Official Website").Page("MyAccount").WebElement("no order details yet !").Highlight
			Reporter.ReportEvent micPass,"No order text highlighted","No order text got highlighted"
			Dim n1
			n1=Browser("Titan: The Official Website").Page("MyAccount").WebElement("no order details yet !").GetROProperty("innertext")
			Reporter.ReportEvent micPass,"No order text fetched","No order text says"&n1
			'Continue shopping button
			If Browser("Titan: The Official Website").Page("MyAccount").WebButton("Continue Shopping").Exist(5) Then
				Reporter.ReportEvent micPass,"Continue shopping button exists","Continue shopping exists and working"
				Browser("Titan: The Official Website").Page("MyAccount").WebButton("Continue Shopping").Highlight
				Reporter.ReportEvent micPass,"Continue shopping button highlighted","Continue shopping button got highlighted"
				Browser("Titan: The Official Website").Page("MyAccount").WebButton("Continue Shopping").Click
				Reporter.ReportEvent micPass,"Continue shopping button clicked","Continue shopping button got clicked"
			Else
				Reporter.ReportEvent micFail,"Continue shopping button error","Continue shopping button didn't work"
			End If
		Else
			Reporter.ReportEvent micFail,"No order error","No order didn't work"
		End If
	Else
		Reporter.ReportEvent micFail,"Order history error","Order history didn't work"
	End If
	'Gift card balance
	If Browser("Titan: The Official Website").Page("MyAccount").Link("Gift Card Balance").Exist(5) Then
		Reporter.ReportEvent micPass,"Gift card balance button exists","Gift card balance button exists and working"
		Browser("Titan: The Official Website").Page("MyAccount").Link("Gift Card Balance").Highlight
		Reporter.ReportEvent micPass,"Gift card balance button highlighted","Gift card balance button got highlighted"
		Browser("Titan: The Official Website").Page("MyAccount").Link("Gift Card Balance").Click
		Reporter.ReportEvent micPass,"Gift card balance button clicked","Gift card balance button got clicked"
		'Gift card text'
		If Browser("Titan: The Official Website").Page("MyAccount").WebElement("Gift Card").Exist(5) Then
			Reporter.ReportEvent micPass,"Gift card text exists","Gift card text exists and working"
			Browser("Titan: The Official Website").Page("MyAccount").WebElement("Gift Card").Highlight
			Reporter.ReportEvent micPass,"Gift card text highlighted","Gift card text got highlighted"
			Dim g1
			g1=Browser("Titan: The Official Website").Page("MyAccount").WebElement("Gift Card").GetROProperty("innertext")
			Reporter.ReportEvent micPass,"Gift card text fetched","Gift card text says:"&g1
		Else
			Reporter.ReportEvent micFail,"Gift card text error","Gift card text didn't work"		
		End If
		'Gift card balance text'
		If Browser("Titan: The Official Website").Page("MyAccount").WebElement("Gift Card Balance").Exist(5) Then
			Reporter.ReportEvent micPass,"Gift card balance text exists","Gift card balance text exists and working"
			Browser("Titan: The Official Website").Page("MyAccount").WebElement("Gift Card Balance").Highlight
			Reporter.ReportEvent micPass,"Gift card balance highlighted","Gift card balance got highlighted"
			Dim g2
			g2=Browser("Titan: The Official Website").Page("MyAccount").WebElement("Gift Card Balance").GetROProperty("innertext")
			Reporter.ReportEvent micPass,"Gift card balance text fetched","Gift card balance text says:"&g2
		Else
			Reporter.ReportEvent micFail,"Gift card balance text error","Gift card balance text didn't work"
		End If
		'16 digit text'
		If Browser("Titan: The Official Website").Page("MyAccount").WebElement("To view your card balance").Exist(5) Then
			Reporter.ReportEvent micPass,"To view your card balance text exists","The view your card balance text exists and working"
			Browser("Titan: The Official Website").Page("MyAccount").WebElement("To view your card balance").Highlight
			Reporter.ReportEvent micPass,"To view your card balance highlighted","The view your card balance text got highlighted"
			Dim g3
			g3=Browser("Titan: The Official Website").Page("MyAccount").WebElement("To view your card balance").GetROProperty("innertext")
			Reporter.ReportEvent micPass,"To view your card balance text fetched","The text says:"&g3
		Else
			Reporter.ReportEvent micFail,"To view your card balance text error","To view your card balance text didn't work"
		End If
		'Card number box
		If Browser("Titan: The Official Website").Page("MyAccount").WebEdit("GiftCardNum").Exist(5) Then
			Reporter.ReportEvent micPass,"Gift card number text box exists","The gift card number text box exists and working"
			Browser("Titan: The Official Website").Page("MyAccount").WebEdit("GiftCardNum").Highlight
			Reporter.ReportEvent micPass,"To gift number text box highlighted","The gift number text box got highlighted"
			Browser("Titan: The Official Website").Page("MyAccount").WebEdit("GiftCardNum").Set("4556659826562623")
			Reporter.ReportEvent micPass,"To gift number set","The gift number is set to 4556659826562623"
		Else
			Reporter.ReportEvent micFail,"Card number text error","Card number number text box didn't work"
		End If
		'Pin number box
		If Browser("Titan: The Official Website").Page("MyAccount").WebEdit("6DPin").Exist(5) Then
			Reporter.ReportEvent micPass,"6-DIGIT PIN box exists","The 6-DIGIT PIN box exists and working"
			Browser("Titan: The Official Website").Page("MyAccount").WebEdit("6DPin").Highlight
			Reporter.ReportEvent micPass,"6-DIGIT PIN box highlighted","The 6-DIGIT PIN box got highlighted"
			Browser("Titan: The Official Website").Page("MyAccount").WebEdit("6DPin").Set("635109")
			Reporter.ReportEvent micPass,"6-digit pin set","The 6-Digit pin is set to 4556659826562623"
		Else
			Reporter.ReportEvent micFail,"6-digit pin error","6-digit pin box didn't work"
		End If
		'Check balance button
		If Browser("Titan: The Official Website").Page("MyAccount").WebButton("Check Balance").Exist(5) Then
			Reporter.ReportEvent micPass,"Check balance button exists","The Check balance button exists and working"
			Browser("Titan: The Official Website").Page("MyAccount").WebButton("Check Balance").Highlight
			Reporter.ReportEvent micPass,"Check balance button highlighted","The Check balance button got highlighted"
			Browser("Titan: The Official Website").Page("MyAccount").WebButton("Check Balance").Click
			Reporter.ReportEvent micPass,"Check balance button clicked","The Check balance button got clicked"
		Else
			Reporter.ReportEvent micFail,"Check balance button error","Check balance button didn't work"
		End If
	Else
		Reporter.ReportEvent micFail,"Gift card error","Gift card didn't work"
	End If
	'Neu pass button 
	If Browser("Titan: The Official Website").Page("MyAccount").Link("NeuPass").Exist(5) Then
		Reporter.ReportEvent micPass,"Neu pass button exists","The Neu pass button exists and working"
		Browser("Titan: The Official Website").Page("MyAccount").Link("NeuPass").Highlight
		Reporter.ReportEvent micPass,"Neu pass button highlighted","The Neu pass button got highlighted"
		Browser("Titan: The Official Website").Page("MyAccount").Link("NeuPass").Click
		Reporter.ReportEvent micPass,"Neu pass button clicked","The Neu pass button got clicked"
		'Neu pass text
		If Browser("Titan: The Official Website").Page("MyAccount").WebElement("NeuPassText").Exist(5) Then
			Reporter.ReportEvent micPass,"Neu pass text exists","The Neu pass text exists and working"
			Browser("Titan: The Official Website").Page("MyAccount").WebElement("NeuPassText").Highlight
			Reporter.ReportEvent micPass,"Neu pass text highlighted","The Neu pass text got highlighted"
			Dim g4
			g4=Browser("Titan: The Official Website").Page("MyAccount").WebElement("NeuPassText").GetROProperty("innertext")
			Reporter.ReportEvent micPass,"Neu pass text fetched","The text says:"&g4
		Else
			Reporter.ReportEvent micFail,"Neu pass text error","Neu pass text didn't work"
		End If
		'Welcome text
		If Browser("Titan: The Official Website").Page("MyAccount").WebElement("Welcome to the rewarding").Exist(5) Then
			Reporter.ReportEvent micPass,"Welcome text exists","Welcome text exists and working"
			Browser("Titan: The Official Website").Page("MyAccount").WebElement("Welcome to the rewarding").Highlight
			Reporter.ReportEvent micPass,"Welcome text highlighted","Welcome text got highlighted"
			Dim g5
			g5=Browser("Titan: The Official Website").Page("MyAccount").WebElement("Welcome to the rewarding").GetROProperty("innertext")
			Reporter.ReportEvent micPass,"Welcome text fetched","The text says:"&g5
		Else
			Reporter.ReportEvent micFail,"Welcome text error","Welcome text didn't work"
		End If
		'NeuPass/Titan
		If Browser("Titan: The Official Website").Page("MyAccount").WebElement("NeuPass/Titan").Exist(5) Then
			Reporter.ReportEvent micPass,"NeuPass/TITAN text exists","NeuPass/TITAN text exists and working"
			Browser("Titan: The Official Website").Page("MyAccount").WebElement("NeuPass/Titan").Highlight
			Reporter.ReportEvent micPass,"NeuPass/TITAN text highlighted","NeuPass/TITAN text got highlighted"
			Dim g6,g7
			g6=Browser("Titan: The Official Website").Page("MyAccount").WebElement("NeuPass/Titan").GetROProperty("innertext")
			Reporter.ReportEvent micPass,"NeuPass/TITAN fetched","The text says:"&g6
			'You have ...section'
			Browser("Titan: The Official Website").Page("MyAccount").WebElement("You have 0 NeuCoins").Highlight
			Reporter.ReportEvent micPass,"You have 0 NeuCoins text highlighted","The text got highlighted"
			g7=Browser("Titan: The Official Website").Page("MyAccount").WebElement("You have 0 NeuCoins").GetROProperty("innertext")
			Reporter.ReportEvent micPass,"You have 0 NeuCoins text fetched","The text says:"&g7
		Else
			Reporter.ReportEvent micFail,"NeuPass/TITAN error","NeuPass/TITAN text didn't work"
		End If
		'Know more link
		If Browser("Titan: The Official Website").Page("MyAccount").Link("Know More About NeuPass").Exist(5) Then
			Reporter.ReportEvent micPass,"Know more about link exists","Know more about text exists and working"
			Browser("Titan: The Official Website").Page("MyAccount").Link("Know More About NeuPass").Highlight
			Reporter.ReportEvent micPass,"Know more about link highlighted","Know more about link highlighted"
			Browser("Titan: The Official Website").Page("MyAccount").Link("Know More About NeuPass").Click
			Reporter.ReportEvent micPass,"Know more about link clicked","Know more about link clicked"
		Else
			Reporter.ReportEvent micFail,"Know more about error","Know more about text didn't work"
		End If
	Else
		Reporter.ReportEvent micFail,"Neu pass button error","Neu pass button didn't work"
	End If
End Function

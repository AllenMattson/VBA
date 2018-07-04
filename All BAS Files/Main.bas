Attribute VB_Name = "Main"
Public Driver As New Selenium.FirefoxDriver
Public Assert As New Selenium.Assert
Public Verify As New Selenium.Verify
Public Waiter As New Selenium.Waiter
Public Keys As New Selenium.Keys
Public By As New Selenium.By

Public PageHome As New PageHome
Public PageLogin As New PageLogin
Public PageResult As New PageResult


Sub Main()
  ' Open login page
  PageHome.Go _
          .ClickLogin
  
  ' Type credentials
  Assert.Equals "Log in", PageLogin.Header
  PageLogin.Login "name", "password"
  Assert.Matches "^Login error", PageLogin.ErrorMessage
  
  ' Search content
  PageHome.Go _
          .Search "Eiffel tower"
  Assert.Equals "Eiffel Tower", PageResult.Header
  
  Set PageHome = Nothing
End Sub

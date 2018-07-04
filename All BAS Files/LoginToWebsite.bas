Attribute VB_Name = "LoginToWebsite"
'****************************************************************
'in VB editor goto tools references and select both
'microsoft internet controls and
'microsoft HTML object library
'****************************************************************

Sub Test()

    Const cURL = "http://mail.google.com" 'Enter the web address here
    Const cUsername = "XXXXXX" 'Enter your user name here
    Const cPassword = "XXXXXX" 'Enter your Password here
    
    Dim IE As InternetExplorer
    Dim doc As HTMLDocument
    Dim LoginForm As HTMLFormElement
    Dim UserNameInputBox As HTMLInputElement
    Dim PasswordInputBox As HTMLInputElement
    Dim SignInButton As HTMLInputButtonElement
    Dim HTMLelement As IHTMLElement
    Dim qt As QueryTable
        
    Set IE = New InternetExplorer
    
    IE.Visible = True
    IE.Navigate cURL
    
    'Wait for initial page to load
    
    Do While IE.ReadyState <> READYSTATE_COMPLETE Or IE.Busy: DoEvents: Loop
    
    Set doc = IE.Document
    
    'Get the only form on the page
    
    Set LoginForm = doc.forms(0)
    
    'Get the User Name textbox and populate it
    'input name="Email" id="Email" size="18" value="" class="gaia le val" type="text"
 
    Set UserNameInputBox = LoginForm.elements("Email")
    UserNameInputBox.Value = cUsername
    
    'Get the password textbox and populate it
    'input name="Passwd" id="Passwd" size="18" class="gaia le val" type="password"


    Set PasswordInputBox = LoginForm.elements("Passwd")
    PasswordInputBox.Value = cPassword
    
    'Get the form input button and click it
    'input class="gaia le button" name="signIn" id="signIn" value="Sign in" type="submit"
    
    Set SignInButton = LoginForm.elements("signIn")
    SignInButton.Click
            
    'Wait for the new page to load
    
    Do While IE.ReadyState <> READYSTATE_COMPLETE Or IE.Busy: DoEvents: Loop
    
    End Sub

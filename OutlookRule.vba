Sub CreateRule() 
    Dim colRules As Outlook.Rules 
    Dim oRule As Outlook.Rule 
    Dim colRuleActions As Outlook.RuleActions 
    Dim oMoveRuleAction As Outlook.MoveOrCopyRuleAction 
    Dim oFromCondition As Outlook.ToOrFromRuleCondition 
    Dim oExceptSubject As Outlook.TextRuleCondition 
    Dim oInbox As Outlook.Folder 
    Dim oMoveTarget As Outlook.Folder 



    'Specify target folder for rule move action 
    Set oInbox = Application.Session.GetDefaultFolder(olFolderInbox) 
    'Assume that target folder already exists 
    Set oMoveTarget = oInbox.Folders("Dan") 

    'Get Rules from Session.DefaultStore object 
    Set colRules = Application.Session.DefaultStore.GetRules() 

    'Create the rule by adding a Receive Rule to Rules collection 
    Set oRule = colRules.Create("Dan's rule", olRuleReceive) 

    'Specify the condition in a ToOrFromRuleCondition object 
    'Condition is if the message is from "Dan Wilson" 
    Set oFromCondition = oRule.Conditions.From 
    With oFromCondition 
        .Enabled = True 
        .Recipients.Add ("Dan Wilson") 
        .Recipients.ResolveAll 
    End With 

    'Specify the action in a MoveOrCopyRuleAction object 
    'Action is to move the message to the target folder 
    Set oMoveRuleAction = oRule.Actions.MoveToFolder 
    With oMoveRuleAction 
        .Enabled = True 
        .Folder = oMoveTarget 
    End With 

    'Specify the exception condition for the subject in a TextRuleCondition object 
    'Exception condition is if the subject contains "fun" or "chat" 
    Set oExceptSubject = _ 
        oRule.Exceptions.Subject 
    With oExceptSubject 
        .Enabled = True 
        .Text = Array("fun", "chat") 
    End With 

    'Update the server and display progress dialog 
    colRules.Save 
End Sub 
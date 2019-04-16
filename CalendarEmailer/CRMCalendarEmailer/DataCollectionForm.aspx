<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="DataCollectionForm.aspx.cs" Inherits="CRMCalendarEmailer.DataCollectionForm" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title>Calendar Creation Form</title>
</head>
<body>
    <form id="form1" runat="server" Visible="False">
    <div>
            <h1>Calendar Creation Form</h1>
        </div>
        
        <div id="FormFields">
            <fieldset>
                <legend>
                    Referring Physician
                </legend>
                <p>
                  <label>Send To</label>
                  <asp:TextBox runat="server" ID="SendTo"></asp:TextBox>
				 
			        <label>Subject</label>
			        <asp:TextBox runat="server" ID="Subject"></asp:TextBox>
                </p>
                <p>
                  <label>Location</label>
                  <asp:TextBox runat="server" ID="Location"></asp:TextBox>
				  
			        <label>StartTime</label>
			        <asp:TextBox runat="server" ID="StartTime"></asp:TextBox>
                </p>
                <p>
			        <label>Message</label>
			        <asp:TextBox runat="server" ID="Message"></asp:TextBox>
                </p>
		        <p>
			        <asp:Button 
                        runat="server"
                        ID="FormSubmissionButton"
                        Text="Submit"
                        onclick="SubmitForm"/>
		        </p>
            </fieldset>
        </div>
    </form>
</body>
</html>

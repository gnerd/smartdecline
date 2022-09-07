// Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license at the bottom of this file.

Office.initialize = function () {
}

var UID;
var extendedUID;


// Helper function to add a status message to
// the info bar.
function statusUpdate(icon, text) {
  Office.context.mailbox.item.notificationMessages.replaceAsync("status", {
    type: "informationalMessage",
    icon: icon,
    message: text,
    persistent: false
  });
}

// Adds text into the body of the item, then reports the results
// to the info bar.
function addTextToBody(text, icon, event) {
  Office.context.mailbox.item.body.setSelectedDataAsync(text, { coercionType: Office.CoercionType.Text }, 
    function (asyncResult){
      if (asyncResult.status == Office.AsyncResultStatus.Succeeded) {
        statusUpdate(icon, "\"" + text + "\" inserted successfully.");
      }
      else {
        Office.context.mailbox.item.notificationMessages.addAsync("addTextError", {
          type: "errorMessage",
          message: "Failed to insert \"" + text + "\": " + asyncResult.error.message
        });
      }
      event.completed();
    });
}

function getSubjectRequest(id) {
   // Return a GetItem operation request for the subject of the specified item. 
   var request = 
    '<?xml version="1.0" encoding="utf-8"?>' +
    '<soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"' +
    '               xmlns:xsd="http://www.w3.org/2001/XMLSchema"' +
    '               xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/"' +
    '               xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types">' +
    '  <soap:Header>' +
    '    <RequestServerVersion Version="Exchange2013" xmlns="http://schemas.microsoft.com/exchange/services/2006/types" soap:mustUnderstand="0" />' +
    '  </soap:Header>' +
    '  <soap:Body>' +
    '    <GetItem xmlns="http://schemas.microsoft.com/exchange/services/2006/messages">' +
    '      <ItemShape>' +
    '        <t:BaseShape>IdOnly</t:BaseShape>' +
    '        <t:AdditionalProperties>' +
    '            <t:ExtendedFieldURI DistinguishedPropertySetId="Meeting" PropertyId="3" PropertyType="Binary" />   ' +
    '            <t:FieldURI FieldURI="item:Subject"/>' +
    '            <t:FieldURI FieldURI="calendar:UID"/>' +
    '        </t:AdditionalProperties>' +
    '      </ItemShape>' +
    '      <ItemIds><t:ItemId Id="' + id + '"/></ItemIds>' +
    '    </GetItem>' +
    '  </soap:Body>' +
    '</soap:Envelope>';

   return request;
}

function getDeclineRequest(id, ChangeKey) {
   // Return a GetItem operation request for the subject of the specified item.
   //https://msdn.microsoft.com/en-us/library/office/dd633648(v=exchg.80).aspx

   var request = 
    '<?xml version="1.0" encoding="utf-8"?>' +
    '<soap:Envelope xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/"' +
    '               xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types">' +  
    '  <soap:Header>' +
    '    <RequestServerVersion Version="Exchange2013" xmlns="http://schemas.microsoft.com/exchange/services/2006/types" soap:mustUnderstand="0" />' +
    '  </soap:Header>' +             
    '  <soap:Body>' +
    '    <CreateItem xmlns="http://schemas.microsoft.com/exchange/services/2006/messages" MessageDisposition="SendOnly">' +
    '      <Items>' +
    '        <DeclineItem xmlns="http://schemas.microsoft.com/exchange/services/2006/types">' +
    '           <ReferenceItemId Id="' + id + '" ChangeKey="' + ChangeKey + '"/>' +
    '       </DeclineItem>' +
    '       </Items>' +
    '       </CreateItem>' +
    '  </soap:Body>' +
    '</soap:Envelope>';

   return request;
}

function getDeletedRequest() {
   // Return a GetItem operation request for the subject of the specified item.
   var request = 
'<?xml version="1.0" encoding="utf-8"?> ' +
'<soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" ' +
'      xmlns:m="http://schemas.microsoft.com/exchange/services/2006/messages" ' +
'      xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types" ' +
'      xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/"> ' +
'  <soap:Header> ' +
'      <t:RequestServerVersion Version="Exchange2013" /> ' +
'    </soap:Header> ' +
'    <soap:Body> ' +
'      <m:FindItem Traversal="Shallow"> ' +
'        <m:ItemShape> ' +
'          <t:BaseShape>IdOnly</t:BaseShape> ' +
'        </m:ItemShape> ' +
'       <m:Restriction> ' +
'            <t:IsEqualTo> ' +
'               <t:ExtendedFieldURI DistinguishedPropertySetId="Meeting" PropertyId="3" PropertyType="Binary" /> ' +
'           <t:FieldURIOrConstant> ' +
'           <t:Constant Value="' + extendedUID + '" /> ' +
'          </t:FieldURIOrConstant> ' +
'       </t:IsEqualTo> ' +
'      </m:Restriction> ' +
'      <m:ParentFolderIds> ' +
'       <t:DistinguishedFolderId Id="deleteditems" /> ' +
'      </m:ParentFolderIds> ' +
'    </m:FindItem> ' +
'  </soap:Body> ' +
' </soap:Envelope> ';

   //debugger;
   return request;
}


function getMeetingRequest() {
   // Return a GetItem operation request for the subject of the specified item.
   var request = 
'<?xml version="1.0" encoding="utf-8"?> ' +
'<soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" ' +
'      xmlns:m="http://schemas.microsoft.com/exchange/services/2006/messages" ' +
'      xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types" ' +
'      xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/"> ' +
'  <soap:Header> ' +
'      <t:RequestServerVersion Version="Exchange2013" /> ' +
'    </soap:Header> ' +
'    <soap:Body> ' +
'      <m:FindItem Traversal="Shallow"> ' +
'        <m:ItemShape> ' +
'          <t:BaseShape>IdOnly</t:BaseShape> ' +
'        </m:ItemShape> ' +
'       <m:Restriction> ' +
'            <t:IsEqualTo> ' +
'               <t:ExtendedFieldURI DistinguishedPropertySetId="Meeting" PropertyId="3" PropertyType="Binary" /> ' +
'           <t:FieldURIOrConstant> ' +
'           <t:Constant Value="' + extendedUID + '" /> ' +
'          </t:FieldURIOrConstant> ' +
'       </t:IsEqualTo> ' +
'      </m:Restriction> ' +
'      <m:ParentFolderIds> ' +
'       <t:DistinguishedFolderId Id="calendar" /> ' +
'      </m:ParentFolderIds> ' +
'    </m:FindItem> ' +
'  </soap:Body> ' +
' </soap:Envelope> ';

   //debugger;
   return request;
}


function getTentativeRequest(id, ChangeKey) {
   // Return a GetItem operation request for the subject of the specified item.
   //https://msdn.microsoft.com/en-us/library/office/dd633648(v=exchg.80).aspx

   var request = 
    '<?xml version="1.0" encoding="utf-8"?>' +
    '<soap:Envelope xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/"' +
    '               xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types">' +  
    '  <soap:Header>' +
    '    <RequestServerVersion Version="Exchange2013" xmlns="http://schemas.microsoft.com/exchange/services/2006/types" soap:mustUnderstand="0" />' +
    '  </soap:Header>' +             
    '  <soap:Body>' +
    '    <CreateItem xmlns="http://schemas.microsoft.com/exchange/services/2006/messages" MessageDisposition="SaveOnly">' +
    '      <Items>' +
    '        <TentativelyAcceptItem xmlns="http://schemas.microsoft.com/exchange/services/2006/types">' +
    '           <ReferenceItemId Id="' + id + '" ChangeKey="' + ChangeKey + '"/>' +
    '       </TentativelyAcceptItem>' +
    '       </Items>' +
    '       </CreateItem>' +
    '  </soap:Body>' +
    '</soap:Envelope>';

   return request;
}

function getUpdateRequest(id, ChangeKey) {
   // Return a GetItem operation request for the subject of the specified item.
   //https://msdn.microsoft.com/en-us/library/office/dd633648(v=exchg.80).aspx

   var request = 
    '<?xml version="1.0" encoding="utf-8"?>' +
    '<soap:Envelope xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/" xmlns:m="http://schemas.microsoft.com/exchange/services/2006/messages" ' +
    '    xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types">' +  
    '  <soap:Header>' +
    '    <t:RequestServerVersion Version="Exchange2013" soap:mustUnderstand="0" />' +
    '  </soap:Header>' +             
    '  <soap:Body>' +
    '    <m:UpdateItem MessageDisposition="SaveOnly" ConflictResolution="AlwaysOverwrite" SendMeetingInvitationsOrCancellations="SendToNone">' +
    '      <m:ItemChanges>' +
    '        <t:ItemChange>' +
    '           <t:ItemId Id="' + id + '" ChangeKey="' + ChangeKey + '"/>' +
    '           <t:Updates> ' +    
    '               <t:SetItemField> ' +
    '                   <t:FieldURI FieldURI="calendar:LegacyFreeBusyStatus" /> ' +
    '                       <t:CalendarItem>' +
    '                           <t:LegacyFreeBusyStatus>Free</t:LegacyFreeBusyStatus> ' +
    '                       </t:CalendarItem> ' +
    '               </t:SetItemField> ' +
    '           </t:Updates>' +
    '       </t:ItemChange>' +
    '      </m:ItemChanges> ' +
    '       </m:UpdateItem>' +
    '  </soap:Body>' +
    '</soap:Envelope>';
   //debugger;
   return request;
}

function sendRequest() {
   //Gets the UID,ExtendedUID
   //it callsback to callback
   Office.context.mailbox.makeEwsRequestAsync(
    getSubjectRequest(Office.context.mailbox.item.itemId), callback);
   
}


function sendDeletedRequest() {
   //finds the deleted meeting request
   //it calls back to deletedCallback
   Office.context.mailbox.makeEwsRequestAsync(
    getDeletedRequest(), deletedCallback);
   
}

function addDefaultMsgToBody(event) {
  addTextToBody("Inserted by the Add-in Command Demo add-in.", "blue-icon-16", event);
}

function addMsg1ToBody(event) {
    sendRequest();
  //addTextToBody("Hello World!", "red-icon-16", event);
}

function addMsg2ToBody(event) {
  addTextToBody("Add-in commands are cool!", "red-icon-16", event);
}

function addMsg3ToBody(event) {
  addTextToBody("Visit https://dev.outlook.com today for all of your add-in development needs.", "red-icon-16", event);
}

// Gets the subject of the item and displays it in the info bar.
function getSubjectOriginal(event) {
  var subject = Office.context.mailbox.item.subject;
  
  Office.context.mailbox.item.notificationMessages.addAsync("subject", {
    type: "informationalMessage",
    icon: "blue-icon-16",
    message: "Subject: " + subject,
    persistent: false
  });
  
  event.completed();
}

function callback(asyncResult)  {
   var result = asyncResult.value;
   var context = asyncResult.asyncContext;   
   
   // Process the returned response here.
   //get the ID change changekey, create the decline

   var response = $.parseXML(result);
   UID = response.getElementsByTagName("t:UID")[0].textContent;
   extendedUID = response.getElementsByTagName("t:Value")[0].textContent;
   var ChangeKey = response.getElementsByTagName("t:ItemId")[0].getAttribute("ChangeKey");

    Office.context.mailbox.makeEwsRequestAsync(
      getDeclineRequest(Office.context.mailbox.item.itemId,ChangeKey), declineCallback);
   
}

function declineCallback(asyncResult)  {
   var result = asyncResult.value;
   var context = asyncResult.asyncContext;

   // Process the returned response here.
   sendDeletedRequest();
}

function tentativeCallback(asyncResult)  {
   var result = asyncResult.value;
   var context = asyncResult.asyncContext;
 
   // Process the returned response here.
   //get the ID and changekey so we can mark it free and no reminder
      Office.context.mailbox.makeEwsRequestAsync(
    getMeetingRequest(), updateCallback);
 
}

function deletedCallback(asyncResult)  {
   var result = asyncResult.value;
   var context = asyncResult.asyncContext;
   
   // Process the returned response here.
   //get the ID and changekey

   var response = $.parseXML(result);
   var meetingRequest = response.getElementsByTagName("t:MeetingRequest")[0];
   var itemId = meetingRequest.getElementsByTagName("t:ItemId")[0].getAttribute("Id");
   var itemChangeKey = meetingRequest.getElementsByTagName("t:ItemId")[0].getAttribute("ChangeKey");

   // create the tentative and save it w/o sending it
      Office.context.mailbox.makeEwsRequestAsync(
        getTentativeRequest(itemId,itemChangeKey), tentativeCallback);
}

function updateCallback(asyncResult)  {
   var result = asyncResult.value;
   var context = asyncResult.asyncContext;


   // Process the returned response here.
   var response = $.parseXML(result);

   var itemId = response.getElementsByTagName("t:ItemId")[0].getAttribute("Id");
   var itemChangeKey = response.getElementsByTagName("t:ItemId")[0].getAttribute("ChangeKey");

   // create the tentative and save it w/o sending it
      Office.context.mailbox.makeEwsRequestAsync(
        getUpdateRequest(itemId,itemChangeKey), finalUpdateCallback);
}

function finalUpdateCallback(asyncResult)  {
   var result = asyncResult.value;
   var context = asyncResult.asyncContext;

   debugger;
   // Process the returned response here.
 
}

// Gets the subject of the item and displays it in the info bar.
function getSubject(event) {
  var subject = Office.context.mailbox.item.itemClass;
  
  Office.context.mailbox.item.notificationMessages.addAsync("subject", {
    type: "informationalMessage",
    icon: "blue-icon-16",
    message: "Subject: " + subject,
    persistent: false
  });
  
  event.completed();

}

// Gets the item class of the item and displays it in the info bar.
function getItemClass(event) {
  var itemClass = Office.context.mailbox.item.itemClass;
  
  Office.context.mailbox.item.notificationMessages.addAsync("itemClass", {
    type: "informationalMessage",
    icon: "red-icon-16",
    message: "Item Class: " + itemClass,
    persistent: false
  });
  
  event.completed();
}

// Gets the date and time when the item was created and displays it in the info bar.
function getDateTimeCreated(event) {
  var dateTimeCreated = Office.context.mailbox.item.dateTimeCreated;
  
  Office.context.mailbox.item.notificationMessages.addAsync("dateTimeCreated", {
    type: "informationalMessage",
    icon: "red-icon-16",
    message: "Created: " + dateTimeCreated.toLocaleString(),
    persistent: false
  });
  
  event.completed();
}
// Gets the ID of the item and displays it in the info bar.
function getItemID(event) {
  // Limited to 150 characters max in the info bar, so 
  // only grab the first 50 characters of the ID
  var itemID = Office.context.mailbox.item.itemId.substring(0, 50);
  
  Office.context.mailbox.item.notificationMessages.addAsync("itemID", {
    type: "informationalMessage",
    icon: "red-icon-16",
    message: "Item ID: " + itemID,
    persistent: false
  });
  
  event.completed();
}

// Gets the ID of the item and displays it in the info bar.
function getDecline(event) {
    //debugger;
    if (Office.context.mailbox.item.itemClass == "IPM.Schedule.Meeting.Request") {
        // Limited to 150 characters max in the info bar, so 
        // only grab the first 50 characters of the ID
        var itemID = Office.context.mailbox.item.itemId.substring(0, 50);
        
        Office.context.mailbox.item.notificationMessages.addAsync("itemID", {
            type: "informationalMessage",
            icon: "red-icon-16",
            message: "Sending decline and adding to your calendar",
            persistent: false
        });
        sendRequest();
    }
    event.completed();
}

// MIT License: 
 
// Permission is hereby granted, free of charge, to any person obtaining 
// a copy of this software and associated documentation files (the 
// ""Software""), to deal in the Software without restriction, including 
// without limitation the rights to use, copy, modify, merge, publish, 
// distribute, sublicense, and/or sell copies of the Software, and to 
// permit persons to whom the Software is furnished to do so, subject to 
// the following conditions: 
 
// The above copyright notice and this permission notice shall be 
// included in all copies or substantial portions of the Software. 
 
// THE SOFTWARE IS PROVIDED ""AS IS"", WITHOUT WARRANTY OF ANY KIND, 
// EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF 
// MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND 
// NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE 
// LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION 
// OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION 
// WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
// Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See LICENSE.txt in the project root for license information.

Office.initialize = function () {
}

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

function addXSignHeader(event) {
    Office.context.mailbox.item.internetHeaders.setAsync(
        { "x-sign": "x-sign"},
        function (asyncResult) {
            if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
                console.log("Successfully set headers");
                Office.context.mailbox.item.notificationMessages.addAsync("dateTimeCreated", {
                    type: "informationalMessage",
                    icon: "red-icon-16",
                    message: "Successfully set x-sign headers",
                    persistent: false
                });
            } else {
                console.log("Error setting x-sign header: " + JSON.stringify(asyncResult.error));
            }
        }

    );
}
function addXPDFHeader(event) {
    Office.context.mailbox.item.internetHeaders.setAsync(
        { "x-pdf": "x-pdf" },
        function (asyncResult) {
            if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
                
                Office.context.mailbox.item.notificationMessages.addAsync("dateTimeCreated", {
                    type: "informationalMessage",
                    icon: "red-icon-16",
                    message: "Successfully set x-pdf headers",
                    persistent: false
                });
            } else {
                console.log("Error setting x-pdf header: " + JSON.stringify(asyncResult.error));
            }
        }

    );
}
function addXEncryptHeader(event) {
    Office.context.mailbox.item.internetHeaders.setAsync(
        { "x-encrypt": "x-encrypt" },
        function (asyncResult) {
            if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
                
                Office.context.mailbox.item.notificationMessages.addAsync("dateTimeCreated", {
                    type: "informationalMessage",
                    icon: "red-icon-16",
                    message: "Successfully set x-encrypt headers",
                    persistent: false
                });
            } else {
                console.log("Error setting x-encrypt header: " + JSON.stringify(asyncResult.error));
            }
        }

    );
}
// Gets the subject of the item and displays it in the info bar.
function getSubject(event) {
  var subject = Office.context.mailbox.item.subject;
  
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
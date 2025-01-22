async function showAccountPrimaryContact(executionContext) {
    var formContext = executionContext.getFormContext();
    
    executionContext.getFormContext().ui.setFormNotification("v29", "INFO", "SAPC0001");
    
    try {
        var contact = await retrieveContact(formContext);
        if (contact == null) {
            toggleContactMandatory(formContext, true);
        } else {       
            populateForm(formContext, contact);
            hideEmptyFields(formContext, contact);
        }
    } catch (e) {
        //End script if no error message, otherwise displays message in an error dialogue
        if (e.message)
            Xrm.Navigation.openErrorDialog({ message: e.message });
    }
}

async function saveContactIfRequired(executionContext) {
    var formContext = executionContext.getFormContext();

    var contactAttribute = formContext.getAttribute("primarycontactid");
    if (contactAttribute.getRequiredLevel() != "required")
        return;
    
    var accountId = formContext.getAttribute("customerid").getValue()[0].id;
    var contactId = contactAttribute.getValue()[0].id;
    await saveAccountPrimaryContact(accountId, contactId.slice(1, contactId.length-1));
    
    toggleContactMandatory(formContext, false);
}

async function saveAccountPrimaryContact(accountId, contactId) {
    try {
        await Xrm.WebApi.updateRecord(
            "account",
            accountId,
            {
                "primarycontactid@odata.bind": `/contacts(${contactId})`
            }
        );
    } catch {
        Xrm.Navigation.openErrorDialog({ message: "Error saving primary contact to account" });
    }
}

async function retrieveContact(formContext) {
    var customerAttribute = getCustomerAttribute(formContext);  
    
    var customerRecord = getCustomerRecord(customerAttribute);
    if (customerRecord.entityType != "account")
        throw new Error();

    try {
        var account = await Xrm.WebApi.retrieveRecord(
            "account",
            customerRecord.id,
            "?$select=primarycontactid&$expand=primarycontactid($select=contactid,fullname,emailaddress1,mobilephone)"
        );
        return account.primarycontactid;
    } catch {
        throw new Error("There was an error retrieving account details.");
    }
}

function getCustomerAttribute(formContext) {
    var customerAttribute = formContext.getAttribute("customerid");
    if (customerAttribute == null)
        throw new Error("The Customer attribute could not be found in the Case table.");
    
    return customerAttribute;
}

function getCustomerRecord(customerAttribute) {
    var customerRecordArray = customerAttribute.getValue();
    if (customerRecordArray == null || customerRecordArray.length == 0)
        throw new Error();
    
    return customerRecordArray[0];
}

function toggleContactMandatory(formContext, mandatory) {
    var contactAttribute = formContext.getAttribute("primarycontactid");
    var contactControl = contactAttribute.controls.get((control) => {
        if (control.controlDescriptor.Label == "Account Primary Contact")
            return true;
    })[0];
    
    if (mandatory) {
        contactAttribute.setRequiredLevel("required");
        contactControl.setVisible(true);
    } else {
        contactAttribute.setRequiredLevel("none");
        contactControl.setVisible(false);
    }
}

function populateForm(formContext, contact) {   
    var contactRecord = {
        entityType: "contact",
        id: contact.contactid,
        name: contact.fullname
    };
    formContext.getAttribute("primarycontactid").setValue([contactRecord]);
}

function hideEmptyFields(formContext, contact) {
    var primaryContactForm = formContext.ui.quickForms.get("accountcontact_qfc");
    if (primaryContactForm == null)
        throw new Error("The quick form 'accountcontact_qfc' could not be found");
    
    if (contact.emailaddress1 == null)
        primaryContactForm.getControl("emailaddress1").setVisible(false);
    if (contact.mobilephone == null)
        primaryContactForm.getControl("mobilephone").setVisible(false);
}
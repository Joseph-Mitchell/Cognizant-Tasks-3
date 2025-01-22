async function showAccountPrimaryContact(executionContext) {
    var formContext = executionContext.getFormContext();
    
    executionContext.getFormContext().ui.setFormNotification("v22", "INFO", "SAPC0001");
    
    try {
        var contact = await retrieveContact(formContext);
        if (contact == null)
            return;
            
        populateForm(formContext, contact);
        hideEmptyFields(formContext, contact);     
    } catch (e) {
        //Uses errors without messages to exit script early, otherwise displays message in an error dialogue
        if (e.message)
            Xrm.Navigation.openErrorDialog({ message: e.message });
    }
}

function checkCustomerAttributeExists(formContext) {
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

async function retrieveContact(formContext) {
    var customerAttribute = checkCustomerAttributeExists(formContext);  
    
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

function hideEmptyFields(formContext, contact) {
    var primaryContactForm = formContext.ui.quickForms.get("accountcontact_qfc");
    if (primaryContactForm == null)
        throw new Error("The quick form 'accountcontact_qfc' could not be found");
    
    if (contact.emailaddress1 == null)
        primaryContactForm.getControl("emailaddress1").setVisible(false);
    if (contact.mobilephone == null)
        primaryContactForm.getControl("mobilephone").setVisible(false);
}

function populateForm(formContext, contact) {   
    var contactRecord = {
        entityType: "contact",
        id: contact.contactid,
        name: contact.fullname
    };
    formContext.getAttribute("primarycontactid").setValue([contactRecord]);
}